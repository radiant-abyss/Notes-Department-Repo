$ErrorActionPreference = 'Stop'

$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $BaseDir 'config.json'
$CredentialPath = Join-Path $BaseDir 'napcat_password.dpapi'

if (-not (Test-Path $ConfigPath)) {
    Write-Host '[ERROR] config.json was not found. Run this script in the bot runtime folder.' -ForegroundColor Red
    exit 2
}

$config = Get-Content -Raw -Encoding UTF8 $ConfigPath | ConvertFrom-Json

$qq = ''
if ($null -ne $config.PSObject.Properties['napcat_login_qq']) {
    $qq = [string]$config.napcat_login_qq
}

if ([string]::IsNullOrWhiteSpace($qq)) {
    $qq = Read-Host 'Enter bot QQ number'
}

$qq = $qq.Trim()
if ($qq -notmatch '^\d+$') {
    Write-Host '[ERROR] QQ number must contain digits only.' -ForegroundColor Red
    exit 3
}

$secure = Read-Host 'Enter bot QQ password (input is hidden)' -AsSecureString
if ($secure.Length -le 0) {
    Write-Host '[ERROR] Password cannot be empty.' -ForegroundColor Red
    exit 4
}

# On Windows, ConvertFrom-SecureString without a key uses DPAPI for the current
# Windows user. The password is not written to config.json in plaintext.
$encrypted = $secure | ConvertFrom-SecureString
Set-Content -Path $CredentialPath -Value $encrypted -Encoding UTF8 -NoNewline

if ($null -ne $config.PSObject.Properties['napcat_login_qq']) {
    $config.napcat_login_qq = $qq
} else {
    $config | Add-Member -NotePropertyName 'napcat_login_qq' -NotePropertyValue $qq
}

if ($null -eq $config.PSObject.Properties['napcat_restart_wait_seconds']) {
    $config | Add-Member -NotePropertyName 'napcat_restart_wait_seconds' -NotePropertyValue 60
}

$config | ConvertTo-Json -Depth 20 | Set-Content -Path $ConfigPath -Encoding UTF8

Write-Host ''
Write-Host '[OK] NapCat login credential saved.' -ForegroundColor Green
Write-Host "QQ: $qq"
Write-Host "Credential file: $CredentialPath"
Write-Host 'The password was NOT written to config.json.'
Write-Host 'If QQ requires CAPTCHA or new-device verification, manual action may still be required.'
