# SheetHierarchy Add-in Installer
# This script installs the SheetHierarchy Excel add-in

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SheetHierarchy Add-in Installer" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Step 0: Check if Excel is running and prompt user to close it
Write-Host "Step 0: Checking for Excel..." -ForegroundColor Yellow

$excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
if ($excelProcesses) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "EXCEL IS CURRENTLY RUNNING" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please:" -ForegroundColor Yellow
    Write-Host "1. Save all your Excel workbooks" -ForegroundColor White
    Write-Host "2. Close all Excel windows" -ForegroundColor White
    Write-Host "3. Run this script again" -ForegroundColor White
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 0
}
else {
    Write-Host "Excel is not running - proceeding with installation" -ForegroundColor Green
}

# Clear Office Add-in Cache
Write-Host ""
Write-Host "Clearing Office Add-in cache..." -ForegroundColor Yellow

# Clear Office cache folders
$cachePaths = @(
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\WEF\Cache",
    "$env:LOCALAPPDATA\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache",
    "$env:LOCALAPPDATA\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\INetCache",
    "$env:LOCALAPPDATA\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\INetCookies",
    "$env:LOCALAPPDATA\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\Temp"
)

foreach ($path in $cachePaths) {
    if (Test-Path $path) {
        try {
            Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Cleared: $path" -ForegroundColor Green
        }
        catch {
            Write-Host "Could not clear: $path (may be in use)" -ForegroundColor Yellow
        }
    }
}

# Clear browser cache (Edge WebView2 - used by Office)
$webView2Paths = @(
    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
    "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache",
    "$env:LOCALAPPDATA\Microsoft\EdgeWebView\User Data\Default\Cache",
    "$env:LOCALAPPDATA\Microsoft\EdgeWebView\User Data\Default\Code Cache"
)

foreach ($path in $webView2Paths) {
    if (Test-Path $path) {
        try {
            Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Cleared WebView2 cache: $path" -ForegroundColor Green
        }
        catch {
            Write-Host "Could not clear WebView2: $path (may be in use)" -ForegroundColor Yellow
        }
    }
}

Write-Host "Cache clearing completed" -ForegroundColor Green

# Step 1: Create directory
Write-Host ""
Write-Host "Step 1: Creating add-in directory..." -ForegroundColor Yellow
$addinPath = "C:\OfficeAddins"
if (!(Test-Path $addinPath)) {
    try {
        New-Item -ItemType Directory -Path $addinPath -Force | Out-Null
        Write-Host "Directory created: $addinPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create directory: $_" -ForegroundColor Red
        exit 1
    }
}
else {
    Write-Host "Directory already exists: $addinPath" -ForegroundColor Green
}

# Step 2: Download manifest
Write-Host ""
Write-Host "Step 2: Downloading manifest file..." -ForegroundColor Yellow
$manifestUrl = "https://martincurrin.github.io/SheetHierarchy/manifest.xml"
$manifestPath = "$addinPath\manifest.xml"

try {
    # Force fresh download (no cache)
    Invoke-WebRequest -Uri $manifestUrl -OutFile $manifestPath -ErrorAction Stop -Headers @{"Cache-Control"="no-cache"}
    Write-Host "Manifest downloaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to download manifest: $_" -ForegroundColor Red
    Write-Host "Please check your internet connection and try again." -ForegroundColor Red
    exit 1
}

# Step 3: Create registry key
Write-Host ""
Write-Host "Step 3: Registering add-in with Excel..." -ForegroundColor Yellow
$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

try {
    if (!(Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
        Write-Host "Created registry key" -ForegroundColor Green
    }
    
    # Remove existing entry first to ensure clean registration
    Remove-ItemProperty -Path $regPath -Name "SheetHierarchy" -ErrorAction SilentlyContinue
    
    New-ItemProperty -Path $regPath -Name "SheetHierarchy" -Value $manifestPath -PropertyType String -Force | Out-Null
    Write-Host "Add-in registered successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to register add-in: $_" -ForegroundColor Red
    exit 1
}

# Success message
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "Installation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Cache cleared and add-in installed!" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Wait 5 seconds before opening Excel" -ForegroundColor White
Write-Host "2. Open Excel (fresh instance)" -ForegroundColor White
Write-Host "3. Look for SheetHierarchy in Home > Add-ins" -ForegroundColor White
Write-Host "4. If issues persist, restart your computer" -ForegroundColor White
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")