<#
.SYNOPSIS
    通用 MSIX/UWP 打包工具 v1.0
    将任意便携版 Windows 程序打包为 MSIX 格式，自动完成：
      1. 生成自签名证书（Python cryptography，有效期可配置）
      2. 从 exe 提取图标颜色，生成带背景的 UWP 图标
      3. 生成 AppxManifest.xml（含文件类型关联）
      4. 打包 → 签名 → 安装证书 → 安装 MSIX

.USAGE
    .\make_msix.ps1                          # 使用 config.json
    .\make_msix.ps1 -Config myapp.json       # 指定配置文件
    .\make_msix.ps1 -Config myapp.json -Verbose

.NOTES
    依赖：Python 3（用于生成证书）
    工具：tools\makeappx.exe, tools\signtool.exe（已内置）
#>

param(
    [string]$Config = "config.json",
    [switch]$Verbose
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─── 颜色输出 ────────────────────────────────────────────────────────────────
function Write-Step  { param($msg) Write-Host "`n[STEP] $msg" -ForegroundColor Cyan }
function Write-OK    { param($msg) Write-Host "  ✅ $msg" -ForegroundColor Green }
function Write-Warn  { param($msg) Write-Host "  ⚠️  $msg" -ForegroundColor Yellow }
function Write-Fail  { param($msg) Write-Host "  ❌ $msg" -ForegroundColor Red; throw $msg }

# ─── 加载配置 ────────────────────────────────────────────────────────────────
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$configPath = if ([System.IO.Path]::IsPathRooted($Config)) { $Config } else { Join-Path $scriptDir $Config }
if (-not (Test-Path $configPath)) { Write-Fail "配置文件不存在: $configPath" }

$cfg = Get-Content $configPath -Raw | ConvertFrom-Json
Write-Host "========================================" -ForegroundColor Magenta
Write-Host "  MSIX 打包工具 v1.0" -ForegroundColor Magenta
Write-Host "  应用: $($cfg.app.displayName) $($cfg.app.version)" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

# ─── 路径设置 ────────────────────────────────────────────────────────────────
$toolsDir   = Join-Path $scriptDir "tools"
$makeappx   = Join-Path $toolsDir "makeappx.exe"
$signtool   = Join-Path $toolsDir "signtool.exe"
$outDir     = if ($cfg.output.dir) { $cfg.output.dir } else { $scriptDir }
$workDir    = Join-Path $env:TEMP "msix_build_$($cfg.app.name)"
$pkgDir     = Join-Path $workDir "package"
$assetsDir  = Join-Path $pkgDir "Assets"
$msixPath   = Join-Path $outDir "$($cfg.app.name).msix"
$pfxPath    = Join-Path $workDir "cert.pfx"
$cerPath    = Join-Path $outDir "$($cfg.app.name).cer"

New-Item -ItemType Directory -Path $workDir  -Force | Out-Null
New-Item -ItemType Directory -Path $pkgDir   -Force | Out-Null
New-Item -ItemType Directory -Path $assetsDir -Force | Out-Null

# ─── STEP 1: 生成证书 ────────────────────────────────────────────────────────
Write-Step "生成自签名证书"

$certPy = Join-Path $workDir "make_cert.py"
$certSubject = $cfg.cert.subject
$certPassword = $cfg.cert.password
$certYears = $cfg.cert.validYears

@"
import sys, os, datetime
try:
    from cryptography import x509
    from cryptography.x509.oid import NameOID
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.hazmat.primitives.serialization import pkcs12, BestAvailableEncryption
except ImportError:
    print("NEED_INSTALL")
    sys.exit(1)

subject_str = r"$certSubject"
password    = b"$certPassword"
years       = $certYears
pfx_path    = r"$pfxPath"
cer_path    = r"$cerPath"

# 解析 CN=
cn = subject_str.split("CN=")[-1].split(",")[0].strip()

key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, cn)])
now  = datetime.datetime.now(datetime.timezone.utc)
cert = (
    x509.CertificateBuilder()
    .subject_name(name)
    .issuer_name(name)
    .public_key(key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(now)
    .not_valid_after(now + datetime.timedelta(days=365*years))
    .add_extension(x509.BasicConstraints(ca=True, path_length=None), critical=True)
    .sign(key, hashes.SHA256())
)

# 保存 PFX
pfx_data = pkcs12.serialize_key_and_certificates(
    name=cn.encode(), key=key, cert=cert, cas=None,
    encryption_algorithm=BestAvailableEncryption(password)
)
with open(pfx_path, "wb") as f: f.write(pfx_data)

# 保存 CER（公钥）
cer_data = cert.public_bytes(serialization.Encoding.DER)
with open(cer_path, "wb") as f: f.write(cer_data)

print(f"OK|{pfx_path}|{cer_path}")
"@ | Set-Content $certPy -Encoding UTF8

$pyOut = python $certPy 2>$null
if ($pyOut -match "NEED_INSTALL") {
    Write-Warn "安装 cryptography 库..."
    pip install cryptography -q 2>$null
    $pyOut = python $certPy 2>$null
}
if ($pyOut -notmatch "^OK\|") { Write-Fail "证书生成失败: $pyOut" }
Write-OK "证书生成: $pfxPath (有效期 $certYears 年)"

# 安装证书到 CurrentUser\Root（无需管理员）
$certObj = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cerPath)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","CurrentUser")
$store.Open("ReadWrite")
$existing = $store.Certificates | Where-Object { $_.Subject -eq $certObj.Subject }
if (-not $existing) {
    $store.Add($certObj)
    Write-OK "证书已安装到 CurrentUser\Root"
} else {
    Write-OK "证书已存在于 CurrentUser\Root（跳过）"
}
$store.Close()

# 同时保存 PFX 到输出目录
Copy-Item $pfxPath (Join-Path $outDir "$($cfg.app.name).pfx") -Force
Write-OK "PFX 已保存: $(Join-Path $outDir "$($cfg.app.name).pfx")"

# ─── STEP 2: 提取图标颜色 ────────────────────────────────────────────────────
Write-Step "提取图标颜色"

Add-Type -AssemblyName System.Drawing

function Get-ExeIconColors {
    param([string]$exePath)
    try {
        $ico = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
        $bmp = $ico.ToBitmap()
        
        # 采样边角（背景色）
        $corners = @(@(0,0),@(1,0),@(0,1),@(1,1))
        $bgSamples = $corners | ForEach-Object {
            $px = $bmp.GetPixel($_[0], $_[1])
            if ($px.A -gt 50) { $px } else { $null }
        } | Where-Object { $_ -ne $null }
        
        # 采样中心（主色）
        $cx = [int]($bmp.Width/2); $cy = [int]($bmp.Height/2)
        $centerPx = $bmp.GetPixel($cx, $cy)
        
        $bmp.Dispose()
        
        $bgColor = if ($bgSamples.Count -gt 0) {
            $r = [int](($bgSamples | Measure-Object -Property R -Average).Average)
            $g = [int](($bgSamples | Measure-Object -Property G -Average).Average)
            $b = [int](($bgSamples | Measure-Object -Property B -Average).Average)
            "#{0:X2}{1:X2}{2:X2}" -f $r,$g,$b
        } else { "#1E1E1E" }
        
        $logoColor = "#{0:X2}{1:X2}{2:X2}" -f $centerPx.R,$centerPx.G,$centerPx.B
        
        return @{ bg=$bgColor; logo=$logoColor }
    } catch {
        return @{ bg="#1E1E1E"; logo="#FFFFFF" }
    }
}

# 确定图标源
$iconSourcePng = $cfg.icon.sourcePng
$iconSourceExe = $cfg.icon.sourceExe

$bgHex   = $cfg.tile.backgroundColor
$logoHex = "#FFFFFF"

if ($cfg.icon.bgColor -ne "auto") { $bgHex = $cfg.icon.bgColor }

if ($iconSourceExe -and (Test-Path $iconSourceExe)) {
    $colors = Get-ExeIconColors $iconSourceExe
    if ($cfg.icon.bgColor   -eq "auto") { $bgHex   = $colors.bg }
    if ($cfg.icon.logoColor -eq "auto") { $logoHex = $colors.logo }
    Write-OK "从 exe 提取颜色: 背景=$bgHex  标志=$logoHex"
} else {
    Write-Warn "未找到 exe，使用默认颜色: 背景=$bgHex"
}

# ─── STEP 3: 生成 UWP 图标 ───────────────────────────────────────────────────
Write-Step "生成 UWP 图标"

function New-UwpIcon {
    param([string]$srcPng, [string]$exePath, [string]$bgColor, [string]$outDir, [int]$w, [int]$h, [string]$name)
    
    $bgR = [Convert]::ToInt32($bgColor.Substring(1,2),16)
    $bgG = [Convert]::ToInt32($bgColor.Substring(3,2),16)
    $bgB = [Convert]::ToInt32($bgColor.Substring(5,2),16)
    $bg  = [System.Drawing.Color]::FromArgb(255,$bgR,$bgG,$bgB)
    
    $dst = New-Object System.Drawing.Bitmap($w,$h)
    $g   = [System.Drawing.Graphics]::FromImage($dst)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.Clear($bg)
    
    if ($srcPng -and (Test-Path $srcPng)) {
        $src = [System.Drawing.Image]::FromFile($srcPng)
    } elseif ($exePath -and (Test-Path $exePath)) {
        $ico = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
        $src = $ico.ToBitmap()
    } else {
        $g.Dispose(); $dst.Dispose(); return
    }
    
    # 居中绘制，留 15% 边距
    $margin = [int]([Math]::Min($w,$h) * 0.15)
    $dw = $w - $margin*2; $dh = $h - $margin*2
    $scale = [Math]::Min($dw/$src.Width, $dh/$src.Height)
    $sw = [int]($src.Width*$scale); $sh = [int]($src.Height*$scale)
    $dx = [int](($w-$sw)/2); $dy = [int](($h-$sh)/2)
    $g.DrawImage($src, $dx, $dy, $sw, $sh)
    
    $src.Dispose(); $g.Dispose()
    $outPath = Join-Path $outDir $name
    $dst.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $dst.Dispose()
    return $outPath
}

$iconSizes = @(
    @{ name="Square44x44Logo.png";   w=44;  h=44  },
    @{ name="StoreLogo.png";         w=50;  h=50  },
    @{ name="Square150x150Logo.png"; w=150; h=150 },
    @{ name="Wide310x150Logo.png";   w=310; h=150 }
)

foreach ($s in $iconSizes) {
    $out = New-UwpIcon -srcPng $iconSourcePng -exePath $iconSourceExe `
                       -bgColor $bgHex -outDir $assetsDir `
                       -w $s.w -h $s.h -name $s.name
    Write-OK "$($s.name) ($($s.w)x$($s.h))"
}

# ─── STEP 4: 复制应用文件 ────────────────────────────────────────────────────
Write-Step "复制应用文件"

# 确定应用源目录（从 executable 路径推断）
$exeRelPath = $cfg.app.executable
$appSubDir  = Split-Path $exeRelPath -Parent   # e.g. "Sublime Text"
$appDstDir  = Join-Path $pkgDir $appSubDir

# 查找源目录
$exeFileName = Split-Path $exeRelPath -Leaf
$possibleSrc = @(
    "C:\Program Files\$appSubDir",
    "C:\Program Files (x86)\$appSubDir",
    (Join-Path $scriptDir "app"),
    (Join-Path $scriptDir $appSubDir)
)

$appSrcDir = $null
foreach ($p in $possibleSrc) {
    if (Test-Path (Join-Path $p $exeFileName)) {
        $appSrcDir = $p; break
    }
}

if ($appSrcDir) {
    Write-OK "源目录: $appSrcDir"
    if (-not (Test-Path $appDstDir)) {
        Copy-Item $appSrcDir $appDstDir -Recurse -Force
        $count = (Get-ChildItem $appDstDir -Recurse -File).Count
        Write-OK "已复制 $count 个文件"
    } else {
        Write-OK "目标目录已存在，跳过复制"
    }
} else {
    Write-Warn "未找到应用源目录，请手动将应用文件放到: $appDstDir"
    New-Item -ItemType Directory -Path $appDstDir -Force | Out-Null
}

# ─── STEP 5: 生成 AppxManifest.xml ──────────────────────────────────────────
Write-Step "生成 AppxManifest.xml"

# 构建文件类型关联 XML
$ftaXml = ""
if ($cfg.fileTypes.enabled) {
    # Windows 禁止注册的扩展名
    $blocked = @(".exe",".com",".bat",".cmd",".msi",".msix",".dll",".sys",".drv",".scr",".pif",".vbs",".vbe",".js",".jse",".wsf",".wsh",".msc",".cpl",".reg",".inf",".ins",".isp",".lnk",".url",".ps1",".psc1",".msh",".msh1",".msh2",".mshxml",".msh1xml",".msh2xml")
    # 注意：.ps1 在某些系统上也被保护，但通常可以注册
    $blocked = @(".exe",".com",".bat",".cmd",".msi",".msix",".dll",".sys",".drv",".scr",".pif",".vbs",".vbe",".jse",".wsf",".wsh",".msc",".cpl",".reg",".inf",".lnk",".url")
    
    foreach ($grp in $cfg.fileTypes.groups) {
        $validExts = $grp.extensions | Where-Object { $_ -notin $blocked }
        if ($validExts.Count -eq 0) { continue }
        
        $ftLines = ($validExts | ForEach-Object { "              <uap:FileType>$_</uap:FileType>" }) -join "`n"
        $ftaXml += @"

        <uap:Extension Category="windows.fileTypeAssociation">
          <uap:FileTypeAssociation Name="$($grp.name)">
            <uap:SupportedFileTypes>
$ftLines
            </uap:SupportedFileTypes>
            <uap:Logo>Assets\Square150x150Logo.png</uap:Logo>
          </uap:FileTypeAssociation>
        </uap:Extension>
"@
    }
}

$extensionsBlock = if ($ftaXml) {
    "      <Extensions>$ftaXml`n      </Extensions>"
} else { "" }

$manifest = @"
<?xml version="1.0" encoding="utf-8"?>
<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
         xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
         xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
         IgnorableNamespaces="uap rescap">

  <Identity Name="$($cfg.app.name)"
            Publisher="$($cfg.app.publisher)"
            Version="$($cfg.app.version)"
            ProcessorArchitecture="$($cfg.app.arch)"/>

  <Properties>
    <DisplayName>$($cfg.app.displayName)</DisplayName>
    <PublisherDisplayName>$($cfg.app.publisherDisplay)</PublisherDisplayName>
    <Description>$($cfg.app.description)</Description>
    <Logo>Assets\StoreLogo.png</Logo>
  </Properties>

  <Resources>
    <Resource Language="en-US"/>
  </Resources>

  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop"
                        MinVersion="10.0.17763.0"
                        MaxVersionTested="10.0.22621.0"/>
  </Dependencies>

  <Capabilities>
    <rescap:Capability Name="runFullTrust"/>
    <rescap:Capability Name="broadFileSystemAccess"/>
  </Capabilities>

  <Applications>
    <Application Id="$($cfg.app.name)"
                 Executable="$($cfg.app.executable)"
                 EntryPoint="Windows.FullTrustApplication">

      <uap:VisualElements DisplayName="$($cfg.app.displayName)"
                          Description="$($cfg.app.description)"
                          BackgroundColor="$($cfg.tile.backgroundColor)"
                          Square150x150Logo="Assets\Square150x150Logo.png"
                          Square44x44Logo="Assets\Square44x44Logo.png">
        <uap:DefaultTile Wide310x150Logo="Assets\Wide310x150Logo.png"/>
      </uap:VisualElements>

$extensionsBlock
    </Application>
  </Applications>
</Package>
"@

$manifest | Set-Content (Join-Path $pkgDir "AppxManifest.xml") -Encoding UTF8
Write-OK "AppxManifest.xml 生成完成"
if ($ftaXml) {
    $totalExts = ($cfg.fileTypes.groups | ForEach-Object { $_.extensions.Count } | Measure-Object -Sum).Sum
    Write-OK "文件类型关联: $totalExts 种扩展名"
}

# ─── STEP 6: 打包 ────────────────────────────────────────────────────────────
Write-Step "打包 MSIX"

$packResult = cmd /c "`"$makeappx`" pack /d `"$pkgDir`" /p `"$msixPath`" /o" 2>&1
if ($LASTEXITCODE -ne 0) {
    # 检查是否有受保护扩展名错误
    $badExt = $packResult | Select-String "can't register for the" | ForEach-Object {
        if ($_ -match '"(\.[^"]+)"') { $Matches[1] }
    }
    if ($badExt) {
        Write-Warn "以下扩展名受 Windows 保护，已自动移除: $($badExt -join ', ')"
        # 从 manifest 中移除这些扩展名
        $manifestContent = Get-Content (Join-Path $pkgDir "AppxManifest.xml") -Raw
        foreach ($ext in $badExt) {
            $manifestContent = $manifestContent -replace "(?m)^\s*<uap:FileType>\$([regex]::Escape($ext))</uap:FileType>\r?\n", ""
        }
        $manifestContent | Set-Content (Join-Path $pkgDir "AppxManifest.xml") -Encoding UTF8
        
        # 重新打包
        $packResult = cmd /c "`"$makeappx`" pack /d `"$pkgDir`" /p `"$msixPath`" /o" 2>&1
    }
}

if ($LASTEXITCODE -ne 0) { Write-Fail "打包失败:`n$packResult" }
$msixSize = [math]::Round((Get-Item $msixPath).Length/1MB, 1)
Write-OK "打包成功: $msixPath ($msixSize MB)"

# ─── STEP 7: 签名 ────────────────────────────────────────────────────────────
Write-Step "签名 MSIX"

$signResult = cmd /c "`"$signtool`" sign /fd SHA256 /a /f `"$pfxPath`" /p `"$certPassword`" `"$msixPath`"" 2>&1
if ($LASTEXITCODE -ne 0) { Write-Fail "签名失败:`n$signResult" }
Write-OK "签名成功"

# ─── STEP 8: 安装证书到系统根证书 ─────────────────────────────────────────
Write-Step "安装证书到系统根证书"

# 读取证书
$certBytes = [System.IO.File]::ReadAllBytes($cerPath)
$cert2 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certBytes)

# 尝试安装到 LocalMachine\Root（需要管理员权限）
$localMachineSuccess = $false
try {
    $storeLM = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
    $storeLM.Open("ReadWrite")
    $existingLM = $storeLM.Certificates | Where-Object { $_.Subject -eq $cert2.Subject -and $_.NotAfter -eq $cert2.NotAfter }
    if (-not $existingLM) {
        $storeLM.Add($cert2)
        Write-OK "证书已安装到 LocalMachine\Root（需要管理员权限）"
        $localMachineSuccess = $true
    } else {
        Write-OK "证书已存在于 LocalMachine\Root（跳过）"
        $localMachineSuccess = $true
    }
    $storeLM.Close()
} catch {
    Write-Warn "无法安装到 LocalMachine\Root（非管理员权限），尝试 CurrentUser..."
}

# 安装到 CurrentUser\Root（当前用户，无需管理员）
try {
    $storeCU = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
    $storeCU.Open("ReadWrite")
    $existingCU = $storeCU.Certificates | Where-Object { $_.Subject -eq $cert2.Subject }
    if (-not $existingCU) {
        $storeCU.Add($cert2)
        Write-OK "证书已安装到 CurrentUser\Root"
    } else {
        Write-OK "证书已存在于 CurrentUser\Root"
    }
    $storeCU.Close()
} catch {
    Write-Warn "证书安装失败: $_"
}

# ─── STEP 9: 安装 MSIX ──────────────────────────────────────────────────────
if ($cfg.output.autoInstall) {
    Write-Step "安装 MSIX"
    
    $existing = Get-AppxPackage "*$($cfg.app.name)*" -EA SilentlyContinue
    if ($existing) {
        Remove-AppxPackage -Package $existing.PackageFullName -EA SilentlyContinue
        Write-OK "已卸载旧版本"
    }
    
    # 检查是否以管理员身份运行
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin -and -not $localMachineSuccess) {
        # 非管理员且未安装到 LocalMachine → 尝试提权安装证书
        Write-Host "  正在请求管理员权限安装证书..."
        $certUtilScript = "certutil -addstore `"Root`" `"$cerPath`""
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "cmd.exe"
        $psi.Arguments = "/c $certUtilScript"
        $psi.Verb = "runas"
        $psi.UseShellExecute = $true
        $psi.WindowStyle = "Hidden"
        try {
            $proc = [System.Diagnostics.Process]::Start($psi)
            $proc.WaitForExit()
            if ($proc.ExitCode -eq 0) {
                Write-OK "证书已安装到 LocalMachine\Root（管理员权限）"
                $localMachineSuccess = $true
            }
        } catch {
            Write-Warn "管理员权限被拒绝"
        }
    }
    
    Add-AppxPackage -Path $msixPath -ForceApplicationShutdown 2>&1 | Out-Null
    
    $installed = Get-AppxPackage "*$($cfg.app.name)*"
    if ($installed) {
        Write-OK "安装成功: $($installed.PackageFullName)"
        Write-OK "安装路径: $($installed.InstallLocation)"
    } else {
        Write-Warn "自动安装失败（请确认证书已安装到系统根证书）"
        Write-Host "  手动安装步骤："
        Write-Host "  1. 右键 $($cfg.app.name).cer → 安装 → 选择「受信任的根证书颁发机构」→ 确定"
        Write-Host "  2. 双击 $($cfg.app.name).msix 安装应用"
        Write-Host "  或右键 $($cfg.app.name).msix → 分配 → 部署"
    }
}

# ─── 清理 ────────────────────────────────────────────────────────────────────
if (-not $cfg.output.keepPackageDir) {
    Remove-Item $workDir -Recurse -Force -EA SilentlyContinue
}

# ─── 完成 ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  🎉 打包完成！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  MSIX : $msixPath" -ForegroundColor Green
Write-Host "  证书 : $(Join-Path $outDir "$($cfg.app.name).pfx")  (密码: $certPassword)" -ForegroundColor Green
Write-Host "  CER  : $cerPath" -ForegroundColor Green
Write-Host ""
Write-Host "  📦 跨电脑部署：" -ForegroundColor Cyan
Write-Host "  1. 右键 $($cfg.app.name).cer → 安装 → 受信任的根证书颁发机构 → 确定" -ForegroundColor Yellow
Write-Host "  2. 双击 $($cfg.app.name).msix → 安装应用" -ForegroundColor Yellow
Write-Host ""
Write-Host "  💡 提示：以管理员身份运行脚本可自动安装证书到系统根证书" -ForegroundColor Cyan
Write-Host ""
