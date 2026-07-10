$cert = New-SelfSignedCertificate -DnsName 'localhost' -CertStoreLocation 'Cert:\CurrentUser\My' -NotAfter (Get-Date).AddYears(5) -KeyExportPolicy Exportable -KeySpec KeyExchange
$pwd = ConvertTo-SecureString -String 'czSocket123' -Force -AsPlainText
$pfxPath = Join-Path $PSScriptRoot 'localhost.pfx'
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pwd | Out-Null
Remove-Item -Path "Cert:\CurrentUser\My\$($cert.Thumbprint)"
Write-Host "PFX created: $pfxPath"
