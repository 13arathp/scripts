Add-Type -AssemblyName System.Net.Http
$handler = [System.Net.Http.HttpClientHandler]::new()
$handler.UseCookies = $false
$client = [System.Net.Http.HttpClient]::new($handler)
$req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, 'https://jsonplaceholder.typicode.com/posts/1')
$task = $client.SendAsync($req)
$task.Wait()
$json = $task.Result.Content.ReadAsStringAsync().Result
Write-Output $json
