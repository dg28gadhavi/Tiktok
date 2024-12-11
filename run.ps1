# Set Execution Policy to Bypass for Current User
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force

# Function to check if running as Administrator
function Check-Admin {
    $IsAdmin = [Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544'
    if (-not $IsAdmin) {
        Write-Host -ForegroundColor Red "You are not running as Administrator. Re-launching script with elevated privileges."

        # Re-launch the script as Administrator
        $myArgs = $MyInvocation.MyCommand.Definition
        Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -NoProfile -File $myArgs" -Verb RunAs
        exit
    }
}

# Ensure running as Administrator
Check-Admin

# Function to check if Python is installed
function Is-PythonInstalled {
    try {
        $pythonVersion = python --version 2>&1
        if ($pythonVersion -match "Python \d+\.\d+\.\d+") {
            Write-Host -ForegroundColor Green "Python is already installed: $pythonVersion"
            return $true
        }
    } catch {
        Write-Host -ForegroundColor Red "Python is not installed."
        return $false
    }
}

# Function to install Python
function Install-Python {
    if (Is-PythonInstalled) {
        Write-Host -ForegroundColor Yellow "Skipping Python installation."
        return
    }

    Write-Host -ForegroundColor Cyan "Downloading Python installer..."
    $url = "https://www.python.org/ftp/python/3.9.9/python-3.9.9-amd64.exe"
    $installerPath = "$env:TEMP\python_installer.exe"
    Invoke-WebRequest -Uri $url -OutFile $installerPath

    Write-Host -ForegroundColor Cyan "Installing Python..."
    Start-Process -FilePath $installerPath -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
    Write-Host -ForegroundColor Green "Python installation complete!"
}

# Function to install Python packages
function Install-PythonPackages {
    $packages = @(
        "email_decomposer==0.0.4",
        "openpyxl==3.1.5",
        "playwright==1.49.0",
        "PyQt6==6.7.1",
        "PyQt6_sip==13.8.0",
        "Requests==2.32.3",
        "tenacity==9.0.0"
    )

    foreach ($package in $packages) {
        $packageName = ($package -split '==')[0]
        $isInstalled = pip show $packageName 2>&1
        if ($isInstalled) {
            Write-Host -ForegroundColor Yellow "$packageName is already installed."
        } else {
            Write-Host -ForegroundColor Cyan "Installing $package..."
            pip install --no-cache-dir $package
        }
    }

    Write-Host -ForegroundColor Green "All packages are up-to-date."

    # Check and install Playwright Browsers (Firefox)
    try {
        $playwrightBrowsers = playwright show browsers
        if ($playwrightBrowsers -match "firefox") {
            Write-Host -ForegroundColor Yellow "Playwright Firefox browser is already installed."
        } else {
            Write-Host -ForegroundColor Cyan "Installing Playwright Firefox browser..."
            playwright install firefox
        }
    } catch {
        Write-Host -ForegroundColor Red "Error while checking or installing Playwright browsers: $_"
    }

    Write-Host -ForegroundColor Green "Playwright setup complete."
}

# Function to download and extract Tiktok.zip
function Setup-Tiktok {
    $zipUrl = "https://github.com/dg28gadhavi/Tiktok/raw/main/Tiktok.zip"
    $zipPath = Join-Path $env:TEMP "Tiktok.zip"
    $tiktokDir = Join-Path (Get-Location) "Tiktok"

    if (-not (Test-Path "$tiktokDir\main.py")) {
        Write-Host -ForegroundColor Cyan "Downloading Tiktok.zip..."
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath

        Write-Host -ForegroundColor Cyan "Extracting Tiktok.zip to Tiktok directory..."
        if (-not (Test-Path $tiktokDir)) { New-Item -ItemType Directory -Force -Path $tiktokDir }
        Expand-Archive -Path $zipPath -DestinationPath $tiktokDir -Force
        Write-Host -ForegroundColor Green "Tiktok setup complete."
    } else {
        Write-Host -ForegroundColor Yellow "Tiktok directory already contains main.py. Skipping extraction."
    }
}

# Function to execute main.py
function Load-Program {
    $scriptPath = Join-Path (Get-Location) "Tiktok\main.py"

    if (Test-Path $scriptPath) {
        Write-Host -ForegroundColor Cyan "Executing main.py..."
        python $scriptPath
    } else {
        Write-Host -ForegroundColor Red "main.py not found in the Tiktok directory."
    }
}

# Unified Setup and Execution
Write-Host -ForegroundColor Yellow "==================== Welcome to the Unified Setup and Execution Script ===================="

Write-Host -ForegroundColor Blue "Starting setup..."
Install-Python
Install-PythonPackages
Setup-Tiktok

Write-Host -ForegroundColor Green "Setup complete. Loading program..."
Load-Program

Write-Host -ForegroundColor Green "===================================== Thank you for using the script! ====================================="

