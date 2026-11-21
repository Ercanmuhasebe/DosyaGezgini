param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [string]$Repository
)

$ErrorActionPreference = 'Stop'

if ($Version -notmatch '^\d+\.\d+\.\d+(\.\d+)?$') {
    throw "Version '$Version' geçerli bir sürüm değil. Örnek: 1.0.1"
}
if ($Repository -notmatch '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') {
    throw "Repository '$Repository' owner/repository biçiminde olmalı."
}

$root = $PSScriptRoot
$sourcePath = Join-Path $root 'src\FileExplorerStarter.cs'
$updaterSourcePath = Join-Path $root 'src\UpdateInstaller.cs'
$pdfSharpPath = Join-Path $root 'vendor\PdfSharp-gdi.dll'
$manifestPath = Join-Path $root 'src\longpath.manifest'
$iconPath = Join-Path $root 'src\DosyaGezgini.ico'
$argosPath = Join-Path $root 'src\argos_translate.py'
$templatesPath = Join-Path $root 'templates'
$distPath = Join-Path $root 'dist'
$objPath = Join-Path $root 'obj'
$sitePath = Join-Path $distPath 'site'
$siteUpdatesPath = Join-Path $sitePath 'updates'

$requiredFiles = @(
    $sourcePath,
    $updaterSourcePath,
    $pdfSharpPath,
    $manifestPath,
    $iconPath,
    $argosPath,
    (Join-Path $templatesPath 'berat.xslt'),
    (Join-Path $templatesPath 'kebir.xslt'),
    (Join-Path $templatesPath 'yevmiye.xslt')
)
foreach ($file in $requiredFiles) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
        throw "Derleme girdisi bulunamadı: $file"
    }
}

$frameworkDirectory = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$compiler = Join-Path $frameworkDirectory 'csc.exe'
if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) {
    throw "Windows .NET Framework C# derleyicisi bulunamadı: $compiler"
}

New-Item -ItemType Directory -Path $distPath, $objPath, $siteUpdatesPath -Force | Out-Null
$owner, $repoName = $Repository.Split('/', 2)
$manifestUrl = "https://$owner.github.io/$repoName/updates/latest.xml"

$sourceText = [System.IO.File]::ReadAllText($sourcePath, [System.Text.Encoding]::UTF8)
$versionPattern = 'private const string CurrentAppVersion = "[^"]*";'
$manifestPattern = 'private const string DefaultUpdateManifestUrl = "[^"]*";'
if ([regex]::Matches($sourceText, $versionPattern).Count -ne 1) {
    throw 'CurrentAppVersion tanımı kaynak dosyada tam bir kez bulunamadı.'
}
if ([regex]::Matches($sourceText, $manifestPattern).Count -ne 1) {
    throw 'DefaultUpdateManifestUrl tanımı kaynak dosyada tam bir kez bulunamadı.'
}
$sourceText = [regex]::Replace($sourceText, $versionPattern, "private const string CurrentAppVersion = `"$Version`";", 1)
$sourceText = [regex]::Replace($sourceText, $manifestPattern, "private const string DefaultUpdateManifestUrl = `"$manifestUrl`";", 1)
$buildSourcePath = Join-Path $objPath 'FileExplorerStarter.cs'
[System.IO.File]::WriteAllText($buildSourcePath, $sourceText, [System.Text.UTF8Encoding]::new($false))

$applicationPath = Join-Path $distPath 'DosyaGezgini.exe'
$updaterPath = Join-Path $distPath 'DosyaGezgini-Updater.exe'
$pdfSharpFullPath = (Resolve-Path -LiteralPath $pdfSharpPath).Path
$argosFullPath = (Resolve-Path -LiteralPath $argosPath).Path
$beratFullPath = (Resolve-Path -LiteralPath (Join-Path $templatesPath 'berat.xslt')).Path
$kebirFullPath = (Resolve-Path -LiteralPath (Join-Path $templatesPath 'kebir.xslt')).Path
$yevmiyeFullPath = (Resolve-Path -LiteralPath (Join-Path $templatesPath 'yevmiye.xslt')).Path

$compilerArguments = @(
    '/nologo',
    '/target:winexe',
    '/platform:x64',
    "/out:$applicationPath",
    "/win32manifest:$manifestPath",
    "/win32icon:$iconPath",
    "/reference:$pdfSharpFullPath",
    '/reference:System.IO.Compression.dll',
    '/reference:System.IO.Compression.FileSystem.dll',
    '/reference:System.Security.dll',
    '/reference:System.Web.Extensions.dll',
    "/resource:$beratFullPath,DosyaGezgini.berat.xslt",
    "/resource:$kebirFullPath,DosyaGezgini.kebir.xslt",
    "/resource:$yevmiyeFullPath,DosyaGezgini.yevmiye.xslt",
    "/resource:$argosFullPath,DosyaGezgini.argos_translate.py",
    "/resource:$pdfSharpFullPath,PdfSharp-gdi.dll",
    "/resource:$frameworkDirectory\System.Runtime.WindowsRuntime.dll,System.Runtime.WindowsRuntime.dll",
    "/resource:$frameworkDirectory\System.Runtime.InteropServices.WindowsRuntime.dll,System.Runtime.InteropServices.WindowsRuntime.dll",
    $buildSourcePath
)
& $compiler @compilerArguments
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $applicationPath -PathType Leaf)) {
    throw "Dosya Gezgini derlemesi başarısız oldu (csc kodu $LASTEXITCODE)."
}

$updaterArguments = @(
    '/nologo',
    '/target:winexe',
    '/platform:x64',
    "/out:$updaterPath",
    "/win32icon:$iconPath",
    '/reference:System.Windows.Forms.dll',
    $updaterSourcePath
)
& $compiler @updaterArguments
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $updaterPath -PathType Leaf)) {
    throw "Güncelleme yardımcısı derlemesi başarısız oldu (csc kodu $LASTEXITCODE)."
}

$setupNotesPath = Join-Path $distPath 'Kurulum.txt'
$setupNotes = @"
Dosya Gezgini $Version

1. Eski programı kapatın.
2. DosyaGezgini.exe ve DosyaGezgini-Updater.exe dosyalarını mevcut program klasörüne kopyalayıp eskilerin üzerine yazın.
3. Mevcut Dosya klasörünü silmeyin; bilgisayara özel ayarlar ve geçmiş burada tutulur.
4. DosyaGezgini.exe dosyasını çalıştırın.

Otomatik güncelleme denetimi: Ayarlar > İnternet üzerinden program güncellemesi.
"@
[System.IO.File]::WriteAllText($setupNotesPath, $setupNotes, [System.Text.UTF8Encoding]::new($true))
$setupArchive = Join-Path $distPath 'DosyaGezgini-IlkKurulum.zip'
Compress-Archive -LiteralPath $applicationPath, $updaterPath, $setupNotesPath -DestinationPath $setupArchive -Force

$sha256 = (Get-FileHash -LiteralPath $applicationPath -Algorithm SHA256).Hash
$updaterSha256 = (Get-FileHash -LiteralPath $updaterPath -Algorithm SHA256).Hash
$releaseDownloadUrl = "https://github.com/$Repository/releases/latest/download/DosyaGezgini.exe"
$updaterDownloadUrl = "https://github.com/$Repository/releases/latest/download/DosyaGezgini-Updater.exe"
$escapedDownloadUrl = [System.Security.SecurityElement]::Escape($releaseDownloadUrl)
$escapedUpdaterDownloadUrl = [System.Security.SecurityElement]::Escape($updaterDownloadUrl)
$manifestXml = @"
<?xml version="1.0" encoding="utf-8"?>
<update version="$Version">
  <downloadUrl>$escapedDownloadUrl</downloadUrl>
  <sha256>$sha256</sha256>
  <updaterUrl>$escapedUpdaterDownloadUrl</updaterUrl>
  <updaterSha256>$updaterSha256</updaterSha256>
  <notes>Dosya Gezgini $Version sürümü.</notes>
</update>
"@
[System.IO.File]::WriteAllText((Join-Path $siteUpdatesPath 'latest.xml'), $manifestXml, [System.Text.UTF8Encoding]::new($true))

$pageTemplatePath = Join-Path $root 'updates\index.html'
$pagePath = Join-Path $sitePath 'index.html'
$page = [System.IO.File]::ReadAllText($pageTemplatePath, [System.Text.Encoding]::UTF8)
$page = $page.Replace('OWNER', $owner).Replace('REPOSITORY', $repoName).Replace('CURRENT_VERSION', $Version)
[System.IO.File]::WriteAllText($pagePath, $page, [System.Text.UTF8Encoding]::new($false))

Write-Host "Derleme tamamlandı: $applicationPath"
Write-Host "Güncelleme yardımcısı: $updaterPath"
Write-Host "Sürüm manifesti: $(Join-Path $siteUpdatesPath 'latest.xml')"
Write-Host "SHA-256: $sha256"
