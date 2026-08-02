$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$subject = 'CN=LaTeX Blocks Development'
$certificate = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $subject -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date).AddMonths(1) } |
    Sort-Object NotAfter -Descending | Select-Object -First 1
if (-not $certificate) {
    $certificate = New-SelfSignedCertificate -Subject $subject -Type CodeSigningCert -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(2)
}
$store = [Security.Cryptography.X509Certificates.X509Store]::new('TrustedPublisher', 'CurrentUser')
try {
    $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    if (-not ($store.Certificates | Where-Object Thumbprint -eq $certificate.Thumbprint)) { $store.Add($certificate) }
} finally { $store.Close() }
$props = @"
<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <ManifestCertificateThumbprint>$($certificate.Thumbprint)</ManifestCertificateThumbprint>
    <VisualStudioVersion>18.0</VisualStudioVersion>
  </PropertyGroup>
</Project>
"@
foreach ($projectDirectory in @('LaTeXBlocks.Word.AddIn', 'LaTeXBlocks.PowerPoint.AddIn')) {
    $propsPath = Join-Path $root "src\$projectDirectory\LaTeXBlocks.Development.props"
    [IO.File]::WriteAllText($propsPath, $props, [Text.UTF8Encoding]::new($false))
}
Write-Output "LaTeX Blocks signing initialized: $($certificate.Thumbprint)"
