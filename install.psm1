function New-ScreepsTypeScriptSetup {
    param(
        $CodePath = $(Join-Path $(Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Personal').Personal '\screeps'),
        [switch]$ConfigureAccount,
        $Email,
        $Password,
        $Token,
        $ServerUrl = 'https://screeps.com',
        $ServerPassword,
        [switch]$GZip = $false,
        $NodeDownloadUrl
    )

    function Expand-ZIPFile
                                            {
    param(
        $file,
        $destination
    )
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item, 0x14)
    }
    }

    # Test for Node.js
    $NodejsRegPath = 'HKLM:\SOFTWARE\Node.js'
    if(Test-Path $NodejsRegPath)
                {
    $NodejsVersion = (Get-ItemProperty $NodejsRegPath -Name 'Version').Version
    $NodejsInstallPath = (Get-ItemProperty $NodejsRegPath -Name 'InstallPath').InstallPath
    Write-Output "Node.js version $NodejsVersion, install found."    
    }
                                                                                                                                        else{
    # Install Node.js

    # Check for admin rights
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
    {
        Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
        return
    }

    if(-not $NodeDownloadUrl)
    {
        if ([System.IntPtr]::Size -eq 4) 
        { 
            $NodeDownloadUrl = 'https://nodejs.org/dist/v6.11.3/node-v6.11.3-x86.msi'
        } 
        else 
        { 
            $NodeDownloadUrl = 'https://nodejs.org/dist/v6.11.3/node-v6.11.3-x64.msi'
        }
    
    }
    Write-Output "Downloading $(Split-Path $NodeDownloadUrl -Leaf)"
    $NodeDownloadPath = Join-Path $env:temp (Split-Path $NodeDownloadUrl -Leaf)
    (new-object Net.WebClient).DownloadFile($NodeDownloadUrl,$NodeDownloadPath)
    Write-Output "Installing Node.js"
    Start-Process msiexec -ArgumentList "/i $NodeDownloadPath /quiet /norestart /log $Env:temp\NodeInstall.log" -Wait
    Remove-Item $NodeDownloadPath -Force
    try
    {
        $NodejsInstallPath = (Get-ItemProperty $NodejsRegPath -Name 'InstallPath' -ErrorAction Stop).InstallPath
    }
    catch{
        Write-Output "There was an error installing Node.js.`nRead the log for more info. $Env:temp\NodeInstall.log"
        return
    }
    }

    # Download and extract screeps-typescript-starter.zip
    Write-Output "Downloading latest build of screeps-typescript-starter."
    (new-object Net.WebClient).DownloadFile('https://codeload.github.com/screepers/screeps-typescript-starter/zip/master',"$env:temp\screeps-typescript-starter.zip")

    # Create directory if needed
    if( -not (Test-Path $CodePath))
    {
        New-Item -ItemType Directory -Force -Path $CodePath | Out-Null
    }
    Write-Output "Extracting to $CodePath"
    Expand-ZIPFile "$env:temp\screeps-typescript-starter.zip" $CodePath

    # Remove downloaded zip
    Remove-Item "$env:temp\screeps-typescript-starter.zip" -Force

    Write-Output "Installing"
    Start-Process -FilePath "$($NodejsInstallPath)npm" -WorkingDirectory "$CodePath\screeps-typescript-starter-master" -ArgumentList 'install > %temp%\npm-install.log 2>&1' -Wait
    Write-Output "Install results logged to $env:temp\npm-install.log"

    # Create local dev path
    $LocalDevPath = "$env:LOCALAPPDATA\Screeps\scripts\127_0_0_1___21025\dev"
    if( -not (Test-Path $LocalDevPath))
    {
        New-Item -ItemType Directory -Force -Path $LocalDevPath | Out-Null
    }

    # Configure the local config
    Write-Output "Updating config.local.ts"
    $LocalConfigPath = "$CodePath\screeps-typescript-starter-master\config\config.local.ts"
    $LocalConfig = Get-Content $LocalConfigPath
    $EncodedPath = ($LocalDevPath) -replace '\\','\\'
    $LocalConfig = $LocalConfig -replace 'const localPath = "/home/USER_NAME/.config/Screeps/scripts/127_0_0_1___21025/default/";',"const localPath = `"$EncodedPath`";"
    $LocalConfig | Out-File $LocalConfigPath -Force

    Write-Output "Running local"
    Start-Process -FilePath "$($NodejsInstallPath)npm" -WorkingDirectory "$CodePath\screeps-typescript-starter-master" -ArgumentList 'run local > %temp%\npm-runLocal.log 2>&1' -Wait
    if(-not (Test-Path "$LocalDevPath\main.js"))
    {
        Write-Output "There was an error running local. main.js not created."
        Write-Output "npm run local logged to $env:temp\npm-runLocal.log"
    }
    else
    {
        $Main = Get-Item "$LocalDevPath\main.js"
        if($Main.LastWriteTime -lt (Get-Date).AddSeconds(-5))
        {
            Write-Output "There was an error running local. main.js not updated, last update $($Main.LastWriteTime)."
            Write-Output "npm run local logged to $env:temp\npm-runLocal.log"
        }
    }
    if($ConfigureAccount)
    {
        Write-Output "Creating credentials.json"
        $CredentialsPath = "$CodePath\screeps-typescript-starter-master\config\credentials.example.json"
        $Credentials = Get-Content $CredentialsPath | ConvertFrom-Json
        $Credentials.email = $Email
        $Credentials.password = $Password
        $Credentials.token = $Token
        $Credentials.serverUrl = $ServerUrl
        $Credentials.serverPassword = $ServerPassword
        $Credentials.gzip = $GZip
        $Credentials | ConvertTo-Json | Out-File "$CodePath\screeps-typescript-starter-master\config\credentials.json" -Force

        Write-Output "Running deploy"
        Start-Process -FilePath "$($NodejsInstallPath)npm" -WorkingDirectory "$CodePath\screeps-typescript-starter-master" -ArgumentList 'run deploy' -Wait
    }
}
