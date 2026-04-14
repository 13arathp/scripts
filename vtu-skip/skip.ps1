Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# TLS 1.2 fix for Windows PowerShell 5.1
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

# Increase connection limit for parallel async blasts
try { [Net.ServicePointManager]::DefaultConnectionLimit = 500 } catch { }

# Initialize Global HttpClient for fast asynchronous dispatch
try {
    Add-Type -AssemblyName System.Net.Http
    $Global:HttpClientHandler = [System.Net.Http.HttpClientHandler]::new()
    $Global:HttpClientHandler.UseCookies = $false
    $Global:HttpClientHandler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
    $Global:HttpClient = [System.Net.Http.HttpClient]::new($Global:HttpClientHandler)
    $Global:HttpClient.Timeout = [System.TimeSpan]::FromSeconds(15)
}
catch {
    Write-Warning "Could not load System.Net.Http for async requests."
}

# Enforce UTF8 output for modern UI characters
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

#region ------ CONFIGURATION ------------------------------------------------------------------------------------------------------------------------------------------------------------------

$Config = [pscustomobject]@{
    Origin         = 'https://online.vtu.ac.in'
    BaseUrl        = 'https://online.vtu.ac.in/api/v1'
    LoginEndpoint  = '/auth/login'
    EnrollEndpoint = '/student/my-enrollments'

    # Server strictly enforces chunk size between 0-120 seconds.
    WatchChunk     = 120

    # Absolute safety cap. Natural exit is ceil(totalSeconds / 120) chunks.
    # 200 x 120s = ~6.6 hours. Prevents infinite loops on bad API responses.
    MaxRetries     = 100

    RetryCount     = 3    # HTTP-level retries per failed API request
    RetryDelayMs   = 1500  # ms between HTTP retries
    DelayMs        = 470    # ms between consecutive API calls. Network RTT (~600ms) is the natural rate limiter.
}

#endregion ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#region ------ LOGGING ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

$Global:LogFile = $null

function Initialize-Logging {
    $logDir = Join-Path ([Environment]::GetFolderPath('Desktop')) 'VTULogs'
    $null = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction SilentlyContinue
    $Global:LogFile = Join-Path $logDir ('VTU_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Add-Content -Path $Global:LogFile -Encoding UTF8 -Value ('[{0}] Session started' -f (Get-Date))
}

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    if (-not $Global:LogFile) { return }
    $ts = (Get-Date).ToString('HH:mm:ss.fff')
    Add-Content -Path $Global:LogFile -Encoding UTF8 -Value "[$ts][$Level] $Message"
}

#endregion ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#region ------ UTILITIES ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function Get-UIPad {
    $w = $host.UI.RawUI.WindowSize.Width
    $boxW = 94
    if ($w -gt $boxW) { return ' ' * [math]::Floor(($w - $boxW) / 2) }
    return ''
}

function ConvertFrom-VTUDuration {
    # Parses VTU duration string "HH:MM:SS mins" -> total seconds (int).
    param([string]$Duration)
    if ([string]::IsNullOrWhiteSpace($Duration)) { return 0 }
    $parts = @()
    foreach ($p in ($Duration -split '[:\s]+')) { if ($p -match '^\d+$') { $parts += $p } }
    if ($parts.Count -lt 3) { return 0 }
    return ([int]$parts[0] * 3600) + ([int]$parts[1] * 60) + [int]$parts[2]
}

function ConvertTo-ProgressBar {
    param([int]$Current, [int]$Total, [int]$Width = 22)
    if ($Total -le 0) { return ('[' + ('=' * $Width) + '] 100%') }
    $pct = [math]::Min(100, [math]::Round(($Current / $Total) * 100))
    $fill = [math]::Round(($pct / 100) * $Width)
    $empty = $Width - $fill
    return ('[' + ('=' * $fill) + (' ' * $empty) + "] $pct%")
}

#endregion ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#region ------ HTTP HELPERS ------------------------------------------------------------------------------------------------------------------------------------------------------------------

function New-Headers {
    # Builds request headers to mimic a real browser session.
    param(
        [Parameter(Mandatory)][string]$Referer,
        [Parameter(Mandatory)][string]$CookieHeader,
        [switch]$JsonContent,
        [switch]$IncludeXHR
    )
    $h = @{
        'Accept'             = 'application/json'
        'Accept-Language'    = 'en-US,en;q=0.9'
        'DNT'                = '1'
        'Origin'             = $Config.Origin
        'Referer'            = $Referer
        'Sec-Fetch-Dest'     = 'empty'
        'Sec-Fetch-Mode'     = 'cors'
        'Sec-Fetch-Site'     = 'same-origin'
        'User-Agent'         = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36 Edg/143.0.0.0'
        'sec-ch-ua'          = '"Microsoft Edge";v="143", "Chromium";v="143", "Not A(Brand";v="24"'
        'sec-ch-ua-mobile'   = '?0'
        'sec-ch-ua-platform' = '"Windows"'
        'Cookie'             = $CookieHeader
    }
    if ($JsonContent) { $h['Content-Type'] = 'application/json' }
    if ($IncludeXHR) { $h['X-Requested-With'] = 'XMLHttpRequest' }
    return $h
}

function Invoke-Api {
    # Central HTTP wrapper: delays, logs, and captures response body on errors.
    param(
        [Parameter(Mandatory)][ValidateSet('GET', 'POST')][string]$Method,
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][hashtable]$Headers,
        [object]$Body,
        [Parameter(Mandatory)][Microsoft.PowerShell.Commands.WebRequestSession]$Session
    )
    Start-Sleep -Milliseconds $Config.DelayMs
    Write-Log "$Method $Url"
    try {
        if ($Method -eq 'GET') {
            return Invoke-RestMethod -Uri $Url -Method GET -Headers $Headers -WebSession $Session -ErrorAction Stop
        }
        else {
            $jsonBody = if ($null -ne $Body) { $Body | ConvertTo-Json -Depth 10 } else { '{}' }
            Write-Log "Body: $jsonBody" 'DEBUG'
            return Invoke-RestMethod -Uri $Url -Method POST -Headers $Headers -Body $jsonBody -WebSession $Session -ErrorAction Stop
        }
    }
    catch {
        Write-Log "FAILED $Method $Url : $($_.Exception.Message)" 'ERROR'
        try {
            if ($_.Exception.Response -and $_.Exception.Response.GetResponseStream()) {
                $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                Write-Log "ResponseBody: $($sr.ReadToEnd())" 'ERROR'
            }
        }
        catch { }
        throw
    }
}

function Get-CookieHeaderFromSession {
    # Extracts access_token and refresh_token from the session cookie jar.
    param([Parameter(Mandatory)][Microsoft.PowerShell.Commands.WebRequestSession]$Session)
    $cookies = $Session.Cookies.GetCookies([uri]$Config.Origin)
    $at = ($cookies | Where-Object Name -eq 'access_token'  | Select-Object -First 1).Value
    $rt = ($cookies | Where-Object Name -eq 'refresh_token' | Select-Object -First 1).Value
    if (-not $at -or -not $rt) { throw 'Login succeeded but tokens not found in cookies.' }
    return "access_token=$at; refresh_token=$rt"
}

function Format-EnrollmentArray {
    # Normalises enrollment API response into a flat array.
    # Handles both JSON array and object-with-numeric-keys formats.
    param([Parameter(Mandatory)]$Resp)
    $data = if ($Resp.data) { $Resp.data } else { $Resp }
    if ($data -is [System.Array]) { return $data }
    $items = @()
    foreach ($p in $data.PSObject.Properties) {
        if ($p.Name -match '^\d+$') { $items += $p.Value }
    }
    return $items
}

#endregion ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#region ------ API CALLS ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function Invoke-Login {
    param([string]$Email, [securestring]$Password)
    # ZeroFreeBSTR wipes the unmanaged BSTR memory immediately after use,
    # preventing the plain-text password from lingering in process memory.
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $url = $Config.BaseUrl + $Config.LoginEndpoint
    $headers = New-Headers -Referer "$($Config.Origin)/auth/login?returnTo=%2Fstudent%2Fenrollments" `
        -CookieHeader 'refresh_token=' -JsonContent
    $body = @{ email = $Email; password = $plain }

    # Send body as a plain JSON object (not array). Skip body logging to protect credentials.
    $loginJson = $body | ConvertTo-Json -Depth 10
    Write-Log "POST $url" ; Write-Log '[login body redacted]' 'DEBUG'
    Start-Sleep -Milliseconds $Config.DelayMs
    $null = Invoke-RestMethod -Uri $url -Method POST -Headers $headers -Body $loginJson -WebSession $session -ErrorAction Stop
    return [pscustomobject]@{
        Session      = $session
        CookieHeader = (Get-CookieHeaderFromSession -Session $session)
    }
}

function Get-Enrollments {
    param($Session, [string]$CookieHeader)
    $url = $Config.BaseUrl + $Config.EnrollEndpoint
    $headers = New-Headers -Referer "$($Config.Origin)/student/enrollments" -CookieHeader $CookieHeader -IncludeXHR
    $resp = Invoke-Api -Method GET -Url $url -Headers $headers -Session $Session
    return Format-EnrollmentArray -Resp $resp
}

function Get-CourseDetails {
    param($Session, [string]$CookieHeader, [string]$Slug)
    $url = "$($Config.BaseUrl)/student/my-courses/$Slug"
    $headers = New-Headers -Referer "$($Config.Origin)/student/course/$Slug" -CookieHeader $CookieHeader
    return Invoke-Api -Method GET -Url $url -Headers $headers -Session $Session
}

function Get-LectureDetails {
    # Fetches individual lecture metadata, including data.duration ("HH:MM:SS mins").
    param($Session, [string]$CookieHeader, [string]$Slug, [string]$LectureId)
    $url = "$($Config.BaseUrl)/student/my-courses/$Slug/lectures/$LectureId"
    $headers = New-Headers -Referer "$($Config.Origin)/student/learning/$Slug" -CookieHeader $CookieHeader -IncludeXHR
    return Invoke-Api -Method GET -Url $url -Headers $headers -Session $Session
}

function Start-ProgressRequestAsync {
    param(
        [string]$CookieHeader,
        [string]$Slug,
        [string]$LectureId,
        [int]$CurrentTime,
        [int]$TotalDuration,
        [int]$Watched
    )
    $url = "$($Config.BaseUrl)/student/my-courses/$Slug/lectures/$LectureId/progress"
    $body = @{
        current_time_seconds   = 99999
        total_duration_seconds = $TotalDuration
        seconds_just_watched   = $Watched
    }
    $jsonBody = $body | ConvertTo-Json -Depth 10

    $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $url)
    $req.Headers.Add('Accept', 'application/json')
    $req.Headers.Add('Accept-Language', 'en-US,en;q=0.9')
    $req.Headers.Add('DNT', '1')
    $req.Headers.Add('Origin', $Config.Origin)
    $req.Headers.Add('Referer', "$($Config.Origin)/student/learning/$Slug")
    $req.Headers.Add('Sec-Fetch-Dest', 'empty')
    $req.Headers.Add('Sec-Fetch-Mode', 'cors')
    $req.Headers.Add('Sec-Fetch-Site', 'same-origin')
    $req.Headers.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36 Edg/143.0.0.0')
    $req.Headers.Add('sec-ch-ua', '"Microsoft Edge";v="143", "Chromium";v="143", "Not A(Brand";v="24"')
    $req.Headers.Add('sec-ch-ua-mobile', '?0')
    $req.Headers.Add('sec-ch-ua-platform', '"Windows"')
    $req.Headers.Add('Cookie', $CookieHeader)
    $req.Headers.Add('X-Requested-With', 'XMLHttpRequest')
    
    $req.Content = [System.Net.Http.StringContent]::new($jsonBody, [System.Text.Encoding]::UTF8, 'application/json')
    return $Global:HttpClient.SendAsync($req)
}

#endregion ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#region ------ DISPLAY ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function Show-Banner {
    Write-Host ''
    $pad = Get-UIPad
    Write-Host ($pad + '+--------------------------------------------------------------------------------------------+') -ForegroundColor Cyan
    Write-Host ($pad + '|                                                                                            |') -ForegroundColor Cyan
    
    $art = @(
        ' ____   ____  _________  _____  _____             ______   ___  ____   _____  _______   ',
        '|_  _| |_  _||  _   _  ||_   _||_   _|          .'' ____ \ |_  ||_  _| |_   _||_   __ \  ',
        '  \ \   / /  |_/ | | \_|  | |    | |    ______  | (___ \_|  | |_/ /     | |    | |__) | ',
        '   \ \ / /       | |      | ''    '' |   |______|  _.____`.   |  __''.     | |    |  ___/  ',
        '    \ '' /       _| |_      \ \__/ /             | \____) | _| |  \ \_  _| |_  _| |_     ',
        '     \_/       |_____|      `.__.''               \______.''|____||____||_____||_____|    '
    )
    foreach ($line in $art) {
        Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
        Write-Host (' ' + $line).PadRight(92) -NoNewline -ForegroundColor White
        Write-Host '|' -ForegroundColor Cyan
    }

    Write-Host ($pad + '|                                                                                            |') -ForegroundColor Cyan

    $esc = [char]27
    $link = 'github.com/13arathp/scripts'
    $fullUrl = 'https://github.com/13arathp/scripts'
    $hyperlink = "$esc]8;;$fullUrl$esc\$link$esc]8;;$esc\"

    # Label line
    $label = '-- source --'
    $lbPad = [math]::Floor((92 - $label.Length) / 2)
    $lbPadR = 92 - $label.Length - $lbPad
    Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
    Write-Host (' ' * $lbPad + $label + ' ' * $lbPadR) -NoNewline -ForegroundColor DarkGray
    Write-Host '|' -ForegroundColor Cyan

    # URL line — bright yellow, centered
    $lPad = [math]::Floor((92 - $link.Length) / 2)
    $lPadR = 92 - $link.Length - $lPad
    Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
    Write-Host (' ' * $lPad) -NoNewline
    Write-Host $hyperlink -NoNewline -ForegroundColor Yellow
    Write-Host (' ' * $lPadR) -NoNewline
    Write-Host '|' -ForegroundColor Cyan

    Write-Host ($pad + '+--------------------------------------------------------------------------------------------+') -ForegroundColor Cyan
    Write-Host ''
}

function Show-InteractiveMenu {
    param([string]$Title, [string[]]$Options)
    $cursorHides = [Console]::CursorVisible
    if ($cursorHides -ne $null) { [Console]::CursorVisible = $false }
    
    $oldTreat = $false
    try { $oldTreat = [Console]::TreatControlCAsInput; [Console]::TreatControlCAsInput = $true } catch {}
    
    try {
        $selectedIndex = 0
        while ($true) {
            Clear-Host
            Show-Banner
            $pad = Get-UIPad
        
            Write-Host ($pad + '+--------------------------------------------------------------------------------------------+') -ForegroundColor Cyan
            Write-Host ($pad + '|                                                                                            |') -ForegroundColor Cyan
        
            $tPad = [math]::Floor((92 - $Title.Length) / 2)
            $tPadR = 92 - $Title.Length - $tPad
            Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
            Write-Host (' ' * $tPad + $Title + ' ' * $tPadR) -NoNewline -ForegroundColor White
            Write-Host '|' -ForegroundColor Cyan
        
            Write-Host ($pad + '|                                                                                            |') -ForegroundColor Cyan

            foreach ($i in 0..($Options.Count - 1)) {
                $opt = $Options[$i]
                $optStrLen = $opt.Length + 5
                $oPad = [math]::Floor((92 - $optStrLen) / 2)
                $oPadR = 92 - $optStrLen - $oPad
            
                Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
                Write-Host (' ' * $oPad) -NoNewline
                if ($i -eq $selectedIndex) {
                    Write-Host '-> ' -NoNewline -ForegroundColor Cyan
                    Write-Host "[$opt]" -NoNewline -ForegroundColor White
                }
                else {
                    Write-Host "   $opt  " -NoNewline -ForegroundColor DarkGray
                }
                Write-Host (' ' * $oPadR) -NoNewline
                Write-Host '|' -ForegroundColor Cyan
            }
            Write-Host ($pad + '|                                                                                            |') -ForegroundColor Cyan
            Write-Host ($pad + '+--------------------------------------------------------------------------------------------+') -ForegroundColor Cyan
            Write-Host ''
        
            $help = '[Use Up/Down Arrows to select, Enter to confirm]'
            $hPad = [math]::Floor((94 - $help.Length) / 2)
            Write-Host ($pad + ' ' * $hPad + $help) -ForegroundColor DarkGray
        
            $keyInfo = [Console]::ReadKey($true)
            if (([int]$keyInfo.Modifiers -band [int][ConsoleModifiers]::Control) -and $keyInfo.Key -eq 'C') {
                if ($cursorHides -ne $null) { try { [Console]::CursorVisible = $true } catch {} }
                return -1
            }
        
            $key = $keyInfo.Key
            if ($key -eq 'UpArrow') {
                $selectedIndex--
                if ($selectedIndex -lt 0) { $selectedIndex = $Options.Count - 1 }
            }
            elseif ($key -eq 'DownArrow') {
                $selectedIndex++
                if ($selectedIndex -ge $Options.Count) { $selectedIndex = 0 }
            }
            elseif ($key -eq 'Enter') {
                if ($cursorHides -ne $null) { try { [Console]::CursorVisible = $true } catch {} }
                return $selectedIndex
            }
        }
    }
    finally {
        try { [Console]::TreatControlCAsInput = $oldTreat } catch {}
    }
}

function Invoke-FetchDetails {
    param($Session, $CookieHeader, $Enrollments, $CourseCache)
    Clear-Host
    Show-Banner
    $pad = Get-UIPad
    Write-Host ($pad + '+-- COURSE OVERVIEW -------------------------------------------------------------------------+') -ForegroundColor Cyan
    $courseIndex = 0
    foreach ($e in $Enrollments) {
        $courseIndex++
        $slug = $e.details.slug
        $title = $e.details.title
        $titleShort = if ($title.Length -gt 85) { $title.Substring(0, 82) + '...' } else { $title }
        
        if ([string]::IsNullOrWhiteSpace($slug)) { continue }

        $progress = if ($null -ne $e.progress_percent) { $e.progress_percent } else { '?' }
        $progressNum = if ($progress -ne '?' -and $null -ne $progress) { [int][double]$progress } else { 0 }
        
        Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
        Write-Host " [$courseIndex/$($Enrollments.Count)] $titleShort".PadRight(92) -NoNewline -ForegroundColor White
        Write-Host '|' -ForegroundColor Cyan
        
        $course = if ($CourseCache.ContainsKey($slug)) { $CourseCache[$slug] } else { $null }
        if ($course) {
            $lessons = $course.data.lessons
            $done = 0; $total = 0
            if ($lessons) {
                foreach ($l in $lessons) {
                    if (-not $l.lectures) { continue }
                    foreach ($lec in $l.lectures) {
                        $total++
                        if ([bool]$lec.is_completed) { $done++ }
                    }
                }
            }
            $pending = $total - $done
            $bar = ConvertTo-ProgressBar -Current $progressNum -Total 100 -Width 40
            
            Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
            Write-Host "   => Progress : $bar".PadRight(92) -NoNewline -ForegroundColor DarkCyan
            Write-Host '|' -ForegroundColor Cyan
            
            Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
            Write-Host "   => Lectures : $done Done | $pending Pending".PadRight(92) -NoNewline -ForegroundColor DarkGray
            Write-Host '|' -ForegroundColor Cyan
            Write-Host ($pad + '|                                                                                            |') -ForegroundColor Cyan
        }
        else {
            Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
            Write-Host "   => [Course data unavailable]".PadRight(92) -NoNewline -ForegroundColor Red
            Write-Host '|' -ForegroundColor Cyan
        }
    }
    Write-Host ($pad + '+--------------------------------------------------------------------------------------------+') -ForegroundColor Cyan
    Write-Host ''
    
    $help = '[Press Enter to return to menu]'
    $hPad = [math]::Floor((94 - $help.Length) / 2)
    Write-Host ($pad + ' ' * $hPad + $help) -NoNewline -ForegroundColor DarkGray
    $null = Read-Host
}

function Show-Divider {
    Write-Host '  ----------------------------------------' -ForegroundColor DarkGray
}

function Show-Step {
    param([string]$Num, [string]$Text)
    Write-Host ''
    $pad = Get-UIPad
    Write-Host ($pad + " +- [$Num] ") -NoNewline -ForegroundColor Cyan
    Write-Host $Text -ForegroundColor White
}

function Show-CourseHeader {
    param([int]$Index, [int]$Total, [string]$Title, $Progress)
    $progressNum = if ($Progress -ne '?' -and $null -ne $Progress) { [int][double]$Progress } else { 0 }
    $bar = ConvertTo-ProgressBar -Current $progressNum -Total 100 -Width 40
    $titleShort = if ($Title.Length -gt 85) { $Title.Substring(0, 82) + '...' } else { $Title }
    $pad = Get-UIPad
    Write-Host ''
    Write-Host ($pad + '+--------------------------------------------------------------------------------------------+') -ForegroundColor Cyan
    
    Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
    Write-Host " Course $Index of $Total".PadRight(92) -NoNewline -ForegroundColor DarkGray
    Write-Host '|' -ForegroundColor Cyan

    Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
    Write-Host " $titleShort".PadRight(92) -NoNewline -ForegroundColor White
    Write-Host '|' -ForegroundColor Cyan

    Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
    Write-Host " $bar".PadRight(92) -NoNewline -ForegroundColor DarkCyan
    Write-Host '|' -ForegroundColor Cyan
    
    Write-Host ($pad + '+--------------------------------------------------------------------------------------------+') -ForegroundColor Cyan
}

function Show-Summary {
    param([int]$Skipped, [int]$Already, [int]$Failed, [string]$Elapsed)
    $pad = Get-UIPad
    Write-Host ''
    Write-Host ($pad + '+-- RUN COMPLETE ----------------------------------------------------------------------------+') -ForegroundColor Cyan
    
    Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
    Write-Host "  Completed : $Skipped lecture(s)".PadRight(92) -NoNewline -ForegroundColor Green
    Write-Host '|' -ForegroundColor Cyan

    Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
    Write-Host "  Already   : $Already lecture(s)".PadRight(92) -NoNewline -ForegroundColor Gray
    Write-Host '|' -ForegroundColor Cyan

    if ($Failed -gt 0) {
        Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
        Write-Host "  Failed    : $Failed lecture(s)".PadRight(92) -NoNewline -ForegroundColor Red
        Write-Host '|' -ForegroundColor Cyan
    }

    Write-Host ($pad + '|') -NoNewline -ForegroundColor Cyan
    Write-Host "  Time      : $Elapsed".PadRight(92) -NoNewline -ForegroundColor White
    Write-Host '|' -ForegroundColor Cyan

    Write-Host ($pad + '+--------------------------------------------------------------------------------------------+') -ForegroundColor Cyan
    Write-Host ''
}

function Get-AllCourseData {
    # Fetches course details for every enrollment and returns a hashtable keyed by slug.
    param($Session, $CookieHeader, $Enrollments)
    $map = @{}
    foreach ($e in $Enrollments) {
        $slug = $e.details.slug
        if ([string]::IsNullOrWhiteSpace($slug)) { continue }
        try { $map[$slug] = Get-CourseDetails -Session $Session -CookieHeader $CookieHeader -Slug $slug }
        catch { Write-Log "Could not prefetch course $slug : $($_.Exception.Message)" 'WARN' }
    }
    return $map
}

#endregion ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#region ------ MAIN ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function Start-VTUSkipper {
    Initialize-Logging
    Clear-Host
    Show-Banner

    # ------ Credentials Input ------------------------------------------------------------------------------------------------------------------------------------------------------------
    $pad = Get-UIPad
    Write-Host ($pad + '  :: AUTHENTICATION') -ForegroundColor Cyan
    Write-Host ($pad + '     Email    -> ') -NoNewline -ForegroundColor DarkGray
    $email = Read-Host
    Write-Host ($pad + '     Password -> ') -NoNewline -ForegroundColor DarkGray
    $secPass = Read-Host -AsSecureString
    Write-Host ''

    # ------ Login ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Show-Step '1/3' 'Logging in...'
    $auth = Invoke-Login -Email $email -Password $secPass
    $session = $auth.Session
    $cookie = $auth.CookieHeader
    Write-Host ($pad + '    OK') -ForegroundColor Green

    # ------ Check enrollments exist before entering menu ------------------------------------------------------------------
    Show-Step '2/3' 'Fetching enrollments...'
    $initialCheck = Get-Enrollments -Session $session -CookieHeader $cookie
    if (-not $initialCheck -or @($initialCheck).Count -eq 0) {
        Write-Host ($pad + '    No enrollments found.') -ForegroundColor Red
        return
    }
    Write-Host ($pad + "    $(@($initialCheck).Count) course(s) found") -ForegroundColor Green

    # ==== MENU LOOP ====
    while ($true) {
        # Prefetch everything before showing the menu so actions are instant
        $pad = Get-UIPad
        Write-Host ''
        Write-Host ($pad + '  Refreshing...') -ForegroundColor DarkGray
        $enrollments = Get-Enrollments    -Session $session -CookieHeader $cookie
        $courseCache = Get-AllCourseData  -Session $session -CookieHeader $cookie -Enrollments $enrollments

        $choices = @('Fetch Course Stats', 'Skip All Courses', 'Exit')
        $sel = Show-InteractiveMenu -Title 'MENU' -Options $choices

        if ($sel -eq -1 -or $sel -eq 2) {
            $pad = Get-UIPad
            Write-Host ''
            Write-Host ($pad + '  Thanks for using vtu-skip!') -ForegroundColor Cyan
            Write-Host ($pad + '  Star or contribute on GitHub:') -ForegroundColor DarkGray
            Write-Host ($pad + '  https://github.com/13arathp/scripts') -ForegroundColor Yellow
            Write-Host ''
            break
        }

        if ($sel -eq 0) {
            Invoke-FetchDetails    -Session $session -CookieHeader $cookie -Enrollments $enrollments -CourseCache $courseCache
        }
        elseif ($sel -eq 1) {
            Invoke-SkipAllCourses  -Session $session -CookieHeader $cookie -Enrollments $enrollments -CourseCache $courseCache
        }
    }
}

function Invoke-SkipAllCourses {
    param($Session, $CookieHeader, $Enrollments, $CourseCache)
    Clear-Host
    Show-Banner
    Show-Step '3/3' 'Processing courses...'
    Write-Host ''

    $totalSkipped = 0
    $totalAlready = 0
    $totalFailed = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $courseIndex = 0
    foreach ($e in $Enrollments) {
        $courseIndex++
        $slug = $e.details.slug
        $title = $e.details.title
        $progress = if ($null -ne $e.progress_percent) { $e.progress_percent } else { '?' }

        # Skip enrollments with no slug (malformed API data)
        if ([string]::IsNullOrWhiteSpace($slug)) { continue }

        # Fast-skip courses already at 100% --- no API call needed
        if ($progress -ne '?' -and [double]$progress -ge 100) {
            Show-CourseHeader $courseIndex $Enrollments.Count $title $progress
            Write-Host '   [--] Already at 100% --- skipped' -ForegroundColor DarkGray
            continue
        }

        Show-CourseHeader $courseIndex $Enrollments.Count $title $progress

        # Use pre-fetched course data from cache
        $course = if ($CourseCache.ContainsKey($slug)) { $CourseCache[$slug] } else { $null }
        if (-not $course) {
            Write-Host "   [!!] No cached data for course - skipping" -ForegroundColor Red
            continue
        }

        $lessons = $course.data.lessons
        if (-not $lessons) {
            Write-Host '   [??] No lessons found' -ForegroundColor Yellow
            continue
        }

        # Collect pending lectures using List to avoid O(n^2) array copying
        $pending = [System.Collections.Generic.List[object]]::new()
        $done = 0
        $total = 0

        for ($i = 0; $i -lt $lessons.Count; $i++) {
            $lects = $lessons[$i].lectures
            if (-not $lects) { continue }
            $lecNum = 1
            foreach ($lec in $lects) {
                $total++
                if ([bool]$lec.is_completed) { $done++ }
                else { $pending.Add([pscustomobject]@{ Week = ($i + 1); LecNum = $lecNum; Id = "$($lec.id)" }) }
                $lecNum++
            }
        }

        $totalAlready += $done

        Write-Host ''
        $pad = Get-UIPad
        Write-Host ($pad + "    => $done done  ") -NoNewline -ForegroundColor Green
        Write-Host "$($pending.Count) pending" -NoNewline -ForegroundColor Yellow
        Write-Host "  of $total" -ForegroundColor DarkGray

        if ($pending.Count -eq 0) {
            Write-Host ($pad + '    All lectures complete!') -ForegroundColor Green
            continue
        }

        Write-Host ''

        # ------ Send Progress Chunks for Each Pending Lecture ------------------------------------------------------------
        $lecIndex = 0
        foreach ($p in $pending) {
            $pad = Get-UIPad # Update dynamically for resizing
            $lecIndex++
            $prefix = "    [$lecIndex/$($pending.Count)]  W$($p.Week) L$($p.LecNum) : $($p.Id)   "

            # Fetch lecture duration for the progress bar and total_duration_seconds in the POST body.
            $totalSeconds = 0
            try {
                $lectDetail = Get-LectureDetails -Session $Session -CookieHeader $CookieHeader -Slug $slug -LectureId $p.Id
                $totalSeconds = ConvertFrom-VTUDuration -Duration $lectDetail.data.duration
            }
            catch {
                Write-Log "Could not fetch duration for lecture $($p.Id): $($_.Exception.Message)" 'WARN'
            }

            # Always use MaxRetries as the cap — let the server decide when it's done.
            # totalChunks is kept only as a denominator for the local progress bar estimate.
            $totalChunks = if ($totalSeconds -gt 0) { [math]::Ceiling($totalSeconds / $Config.WatchChunk) } else { $Config.MaxRetries }

            $currentTime = 0
            $chunk = $Config.WatchChunk
            $completed = $false
            $tries = 0
            $failed = $false
            $bar = ConvertTo-ProgressBar -Current 0 -Total 100 -Width 16

            # Print the static prefix for this lecture line
            Write-Host ($pad + $prefix) -NoNewline -ForegroundColor DarkGray

            # Send chunks asynchronously without waiting for responses in between.
            # We must use MaxRetries as the ceiling because fast async dispatches cause server-side database race conditions (dropped watched time).
            $maxToDispatch = $Config.MaxRetries
            $maxInFlight = if ($totalSeconds -gt 0) { $totalChunks } else { 3 }
            $tasks = [System.Collections.Generic.List[object]]::new()

            $checkTasks = {
                for ($i = $tasks.Count - 1; $i -ge 0; $i--) {
                    $tInfo = $tasks[$i]
                    if (-not $tInfo.Task.IsCompleted) { continue }
                    
                    $isError = $false
                    if ($tInfo.Task.IsFaulted -or $tInfo.Task.IsCanceled) {
                        $isError = $true
                        Write-Log "ASYNC PROGRESS failed lec=$($p.Id) current=$($tInfo.Current)" 'ERROR'
                    }
                    else {
                        try {
                            $r = $tInfo.Task.Result
                            $j = $r.Content.ReadAsStringAsync().Result
                            if ($r.IsSuccessStatusCode) {
                                $d = $j | ConvertFrom-Json
                                $pct = $d.data.percent
                                Write-Log "PROGRESS resp lec=$($p.Id) percent=$pct is_completed=$($d.data.is_completed)"
                                if ([bool]$d.data.is_completed -or ($null -ne $pct -and $pct -ge 100)) { 
                                    $completed = $true 
                                }
                                else {
                                    $bar = if ($null -ne $pct) { ConvertTo-ProgressBar ([int][double]$pct) 100 16 } else { ConvertTo-ProgressBar $tInfo.Tries $totalChunks 16 }
                                    $pad = Get-UIPad; Write-Host "`r$pad$prefix$bar" -NoNewline -ForegroundColor DarkCyan
                                }
                            }
                            else {
                                Write-Log "ASYNC PROGRESS http error status=$($r.StatusCode) payload=$j" 'WARN'
                                $isError = $true 
                            }
                        }
                        catch {
                            Write-Log "ASYNC PROGRESS parse error: $($_.Exception.Message)" 'ERROR'
                            $isError = $true
                        }
                    }

                    if ($isError) {
                        if ($tInfo.RetryCount -lt $Config.RetryCount) {
                            Write-Log "Retry $($tInfo.RetryCount+1)/$($Config.RetryCount) for chunk $($tInfo.Current)" 'WARN'
                            if ($Config.RetryDelayMs -gt 0) { Start-Sleep -Milliseconds $Config.RetryDelayMs }
                            try {
                                $newTask = Start-ProgressRequestAsync -CookieHeader $CookieHeader -Slug $slug -LectureId $p.Id -CurrentTime $tInfo.Current -TotalDuration $totalSeconds -Watched $tInfo.Watched
                                $tasks[$i] = @{ Task = $newTask; Tries = $tInfo.Tries; RetryCount = $tInfo.RetryCount + 1; Current = $tInfo.Current; Watched = $tInfo.Watched }
                                continue
                            }
                            catch { 
                                $failed = $true 
                            }
                        }
                        else {
                            $failed = $true
                        }
                    }
                    $tasks.RemoveAt($i)
                }
            }

            for ($tries = 1; $tries -le $maxToDispatch; $tries++) {
                if ($completed -or $failed) { break }
                while ($tasks.Count -ge $maxInFlight -and -not $completed -and -not $failed) {
                    Start-Sleep -Milliseconds 50
                    . $checkTasks
                }
                if ($completed -or $failed) { break }

                $currentTime += $chunk
                try {
                    $task = Start-ProgressRequestAsync -CookieHeader $CookieHeader -Slug $slug -LectureId $p.Id `
                        -CurrentTime $currentTime -TotalDuration $totalSeconds -Watched $chunk
                    $tasks.Add(@{ Task = $task; Tries = $tries; RetryCount = 0; Current = $currentTime; Watched = $chunk })
                    Write-Log "PROGRESS req lec=$($p.Id) current=$currentTime total=$totalSeconds watched=$chunk"
                }
                catch { 
                    $failed = $true 
                }

                if ($Config.DelayMs -gt 0) { Start-Sleep -Milliseconds $Config.DelayMs }
                . $checkTasks
            }

            # Drain any remaining in-flight tasks
            $drainTries = 0
            while (-not $completed -and -not $failed -and $tasks.Count -gt 0 -and $drainTries -lt 500) {
                Start-Sleep -Milliseconds 100
                $drainTries++
                . $checkTasks
            }

            # If MaxRetries hit without server confirmation, mark as TIMEOUT (not fake-done)

            # Clear the progress bar line, then print final coloured status
            $clearLine = ' ' * ($prefix.Length + 30)
            $pad = Get-UIPad
            Write-Host "`r$pad$clearLine`r" -NoNewline
            Write-Host ($pad + $prefix) -NoNewline -ForegroundColor DarkGray

            if ($completed) {
                $totalSkipped++
                $bar = ConvertTo-ProgressBar -Current 100 -Total 100 -Width 16
                Write-Host "$bar  " -NoNewline -ForegroundColor DarkCyan
                Write-Host 'DONE' -ForegroundColor Green
            }
            elseif ($failed) {
                $totalFailed++
                Write-Host "$bar  " -NoNewline -ForegroundColor DarkCyan
                Write-Host 'FAIL' -ForegroundColor Red
            }
            else {
                # Hit MaxRetries safety cap (only triggers on unknown-duration lectures)
                $totalFailed++
                Write-Host "$bar  " -NoNewline -ForegroundColor DarkCyan
                Write-Host 'TIME' -ForegroundColor Yellow
            }
        }
    }

    # ------ Final Summary ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    $sw.Stop()
    $elapsed = "$([math]::Floor($sw.Elapsed.TotalMinutes))m $($sw.Elapsed.Seconds)s"
    Show-Summary -Skipped $totalSkipped -Already $totalAlready -Failed $totalFailed -Elapsed $elapsed
    Write-Log "Done. Skipped=$totalSkipped Already=$totalAlready Failed=$totalFailed Time=$elapsed"
    
    $pad = Get-UIPad
    if ($Global:LogFile) {
        Write-Host ''
        Write-Host ($pad + "  [i] Logs saved to : ") -NoNewline -ForegroundColor DarkGray
        Write-Host $Global:LogFile -ForegroundColor Gray
    }

    Write-Host "`n$pad  Press [Enter] to return to menu... " -NoNewline -ForegroundColor DarkGray
    $null = Read-Host
}

#endregion ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Entry point -- runs automatically whether piped via iex or executed directly
try {
    Start-VTUSkipper
}
finally {
    if ([Console]::CursorVisible -ne $null) { try { [Console]::CursorVisible = $true } catch {} }
    try {
        [Console]::WriteLine('')
        if ($Global:LogFile) {
            [Console]::ForegroundColor = 'DarkGray'
            [Console]::WriteLine("  [i] Session logs saved to: $($Global:LogFile)")
        }
        [Console]::ForegroundColor = 'Cyan'
        [Console]::WriteLine('')
        [Console]::WriteLine('  Thanks for using vtu-skip!')
        [Console]::ForegroundColor = 'DarkGray'
        [Console]::WriteLine('  Star or contribute on GitHub:')
        [Console]::ForegroundColor = 'Yellow'
        [Console]::WriteLine('  https://github.com/13arathp/scripts')
        [Console]::WriteLine('')
        [Console]::ResetColor()
    }
    catch {}
}

