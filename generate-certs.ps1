$certDir = Join-Path $env:USERPROFILE '.office-addin-dev-certs'
if (-not (Test-Path $certDir)) {
    New-Item -ItemType Directory -Path $certDir -Force | Out-Null
}

Write-Host "Creating Root CA certificate..."
$rootCert = New-SelfSignedCertificate -Subject "CN=Developer CA for Microsoft Office Add-ins" -CertStoreLocation "Cert:\CurrentUser\My" -KeyUsage CertSign, CRLSign, DigitalSignature -KeyLength 2048 -NotAfter (Get-Date).AddYears(5)

$rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::Root, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
$rootStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
$rootStore.Add($rootCert)
$rootStore.Close()

Write-Host "Creating Localhost SSL certificate..."
$serverCert = New-SelfSignedCertificate -DnsName "localhost", "127.0.0.1" -CertStoreLocation "Cert:\CurrentUser\My" -Signer $rootCert -KeyLength 2048 -NotAfter (Get-Date).AddYears(2)

$certPath = Join-Path $certDir "localhost.crt"
$certBytes = $serverCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
[System.IO.File]::WriteAllBytes($certPath, $certBytes)

$pfxPath = Join-Path $certDir "localhost.pfx"
$pfxBytes = $serverCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, "Password123!")
[System.IO.File]::WriteAllBytes($pfxPath, $pfxBytes)

Write-Host "Trusted certificates successfully created in $certDir"
