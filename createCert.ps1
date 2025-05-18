# Path to cert and password file
$certPath = "$env:USERPROFILE\source\repos\voltura\QuickFolders\QuickFoldersDevCert.pfx"
$pwdFile = "$env:USERPROFILE\source\repos\voltura\QuickFolders\cert_pwd.bat"

# Parse CERT_PWD from cert_pwd.bat
$certPwdLine = Get-Content $pwdFile | Where-Object { $_ -match '^SET\s+CERT_PWD=' }
$certPwd = $certPwdLine -replace '^SET\s+CERT_PWD=', ''

# Create cert and export with parsed password
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=QuickFolders Dev" `
    -KeyExportPolicy Exportable -KeySpec Signature `
    -NotAfter (Get-Date).AddYears(3) -CertStoreLocation "Cert:\CurrentUser\My"

$securePwd = ConvertTo-SecureString -String $certPwd -Force -AsPlainText

Export-PfxCertificate -Cert $cert -FilePath $certPath -Password $securePwd
