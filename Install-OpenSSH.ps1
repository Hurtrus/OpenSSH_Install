#region Script Variables
$ExpandOpenSSHPath = 'C:\Program Files\'
$OpenSSHPath = 'C:\Program Files\OpenSSH-Win64\'
$OpenSSHWin64ZipFile = "$env:USERPROFILE\downloads\OpenSSH-Win64.zip"
#endregion Script Variables

#region Download OpenSSH file
# To check for the most recent versions go to https://github.com/PowerShell/Win32-OpenSSH/releases/
$DownloadFile = @{
    URI     = 'https://github.com/PowerShell/Win32-OpenSSH/releases/download/v7.7.2.0p1-Beta/OpenSSH-Win64.zip'
    OutFile = $OpenSSHWin64ZipFile
}
Invoke-WebRequest  @DownloadFile

if (Test-Path $OpenSSHWin64ZipFile)
{
    Write-Host "Download of OpenSSH-Win64.zip was successful" -ForegroundColor Green
}
else
{
    Write-Host "** Error ** - Unable to locate the download of OpenSSH-Win64.zip" -ForegroundColor Red
    break
}
#endregion Download OpenSSH file

#region OpenSSH Install
# Unzip archive to Installation Path
Expand-Archive -Path $OpenSSHWin64ZipFile -DestinationPath $ExpandOpenSSHPath -Force
Test-Path $OpenSSHPath

# Run the SSH install script
pwsh.exe -ExecutionPolicy Bypass -File $OpenSSHPath"install-sshd.ps1"

#endregion OpenSSH Install

#region Setup Services
# Get SSH services and test that install-sshd.ps1 was successful
Get-Service *ssh*
# Set SSH services to Automatic
Set-Service sshd -StartupType Automatic
Set-Service ssh-agent -StartupType Automatic

# Start SSH services
Start-Service sshd
Start-Service ssh-agent

# Run netstat to verify listening on port 22
netstat -bano | Select-String -Pattern ':22'

#endregion Setup Services


#region Add OpenSSH to the Path environment
# Add OpenSSH to the Machine $Path
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";$OpenSSHPath", [EnvironmentVariableTarget]::Machine)

# Test that the settings were applied
$TestMachinePathAdded = ([Environment]::GetEnvironmentVariable('Path', [EnvironmentVariableTarget]::Machine) -split ';')
if ($TestMachinePathAdded -like $OpenSSHPath  )
{
    Write-Host "Added to Machine Path" -ForegroundColor Green
}
else
{
    Write-Host "Error - Not Added to Machine Path" -ForegroundColor Red
}


# Add OpenSSH to the User $Path
# Set the variable $Path to User
$Path = [System.Environment]::GetEnvironmentVariable("Path", "User")

# Add the new Path to the User Path environment
[Environment]::SetEnvironmentVariable("Path", $Path + ";$OpenSSHPath", [System.EnvironmentVariableTarget]::User)

# Test that the settings were applied
$TestUserPathAdded = ([Environment]::GetEnvironmentVariable('Path', [EnvironmentVariableTarget]::User) -split ';')
if ($TestUserPathAdded -like $OpenSSHPath  )
{
    Write-Host "Added to User Path" -ForegroundColor Green
}
else
{
    Write-Host "Error - Not Added to User Path" -ForegroundColor Red
}

#endregion Add OpenSSH to the Path environment

 
#region FireWall Rules

netsh advfirewall firewall add rule name="sshd" dir=in action=Allow protocol=TCP localport=22 

# Show Firewall rule has been created
netsh advfirewall firewall show rule name="sshd" 

#endregion FireWall Rules


#region Set default shell to PowerShell

$PwshDefaultShell = @{
    Path         = "HKLM:\SOFTWARE\OpenSSH" 
    Name         = "DefaultShell" 
    Value        = "C:\Program Files\PowerShell\6\pwsh.exe" 
    PropertyType = "String" 
    Force        = $true
}
New-ItemProperty @PwshDefaultShell

#endregion Set default shell to PowerShell


#region Edits to the sshd_config
# Setup Symbolic Link because OpenSSH doesn't handle a space in the path
New-Item -ItemType SymbolicLink -Path c:\pwsh -Value 'C:\Program Files\PowerShell\6'

# Test Symbolic Link works
Test-Path c:\pwsh

# Write necessary lines to the sshd_config file
$AddSSHConfigSubSystem = "Subsystem  powershell c:\pwsh\pwsh.exe -sshs -NoLogo -NoProfile"
Add-Content 'C:\ProgramData\ssh\sshd_config' "`r`n`r`n$AddSSHConfigSubSystem"

(Get-Content 'C:\ProgramData\ssh\sshd_config').replace('#PasswordAuthentication yes', 'PasswordAuthentication yes') |
    Set-Content 'C:\ProgramData\ssh\sshd_config'

# If we want to edit or see the config file in VSCode
# code 'C:\ProgramData\ssh\sshd_config'

#endregion Edits to the sshd_config

Restart-Service sshd
Restart-Service ssh-agent

# A restart of VSCode is required