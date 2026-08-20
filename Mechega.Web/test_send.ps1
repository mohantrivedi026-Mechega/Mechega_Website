$body = @{ name='AgentTest'; email='mechega.safety@gmail.com'; message='Automated test send from agent' } | ConvertTo-Json
$response = Invoke-RestMethod -Uri 'http://localhost:5000/api/contact' -Method Post -ContentType 'application/json' -Body $body -Verbose
$response | ConvertTo-Json -Depth 4
