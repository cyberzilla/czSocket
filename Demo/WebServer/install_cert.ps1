$pfxPath = Join-Path $PSScriptRoot 'localhost.pfx'
$pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxPath, 'czSocket123', 'Exportable')
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'CurrentUser')
$store.Open('ReadWrite')
$store.Add($pfx)
$store.Close()
Write-Host "Certificate installed to Trusted Root (CurrentUser)."
Write-Host "Restart your browser, then https://localhost:8443/ will show Secure."
