$ErrorActionPreference = 'Stop'

$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $BaseDir 'config.json'
$CredentialPath = Join-Path $BaseDir 'napcat_password.dpapi'
$NapCatDir = 'C:\NapCat.Shell'
$Launcher = Join-Path $NapCatDir 'launcher-win10.bat'

# Safe state file: no account IDs, tokens, passwords, or verification URLs.
$StatusPath = Join-Path $BaseDir 'napcat_recovery_status.json'
function Write-RecoveryState([string]$State) {
    $value = @{state=$State; updated_at=[DateTime]::UtcNow.ToString('o')} | ConvertTo-Json -Compress
    $tempPath = $StatusPath + '.tmp'
    [IO.File]::WriteAllText($tempPath, $value, (New-Object Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $tempPath -Destination $StatusPath -Force
}
trap {
    try { Write-RecoveryState 'script_error' } catch {}
    Write-Host '[Watchdog] Recovery script failed. Check task permissions and local settings.'
    exit 20
}
function Test-LocalPort([int]$Port) {
    $client = New-Object Net.Sockets.TcpClient
    try {
        $pending = $client.ConnectAsync('127.0.0.1', $Port)
        if (-not $pending.Wait(700)) { return $false }
        return $client.Connected
    } catch { return $false } finally { $client.Dispose() }
}
function Get-LoginState {
    try {
        $web = Get-Content -Raw -Encoding UTF8 (Join-Path $NapCatDir 'config\webui.json') | ConvertFrom-Json
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $hash = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($web.token + '.napcat')))).Replace('-','').ToLower() } finally { $sha.Dispose() }
        $auth = Invoke-RestMethod -Uri 'http://127.0.0.1:6099/api/auth/login' -Method Post -ContentType 'application/json' -Body (@{hash=$hash} | ConvertTo-Json) -TimeoutSec 4
        if (-not $auth.data.Credential) { return 'existing_unverified' }
        $headers = @{Authorization='Bearer '+$auth.data.Credential}
        $status = Invoke-RestMethod -Uri 'http://127.0.0.1:6099/api/QQLogin/CheckLoginStatus' -Method Post -ContentType 'application/json' -Body '{}' -Headers $headers -TimeoutSec 4
        if ($null -eq $status.data) { return 'existing_unverified' }
        if ($status.data.isLogin) {
            $info = Invoke-RestMethod -Uri 'http://127.0.0.1:6099/api/QQLogin/GetQQLoginInfo' -Method Post -ContentType 'application/json' -Body '{}' -Headers $headers -TimeoutSec 4
            # Never treat a different QQ account as a successful recovery.
            if ([string]$info.data.uin -ne $qq) { return 'existing_unverified' }
            if (Test-LocalPort $OneBotPort) { return 'online_port_ready' }
            return 'online_port_missing'
        }
        if ([string]$status.data.loginError -match '验证码|验证|安全|captcha|verify') { return 'manual_verification' }
        return 'waiting_login'
    } catch { return 'existing_unverified' }
}
function Wait-Recovery {
    $deadline = [DateTime]::UtcNow.AddSeconds(90)
    do {
        if (Test-LocalPort 6099) {
            $state = Get-LoginState
            Write-RecoveryState $state
            if ($state -in @('online_port_ready','manual_verification','existing_unverified')) { return }
        }
        Start-Sleep -Seconds 5
    } while ([DateTime]::UtcNow -lt $deadline)
    if (Test-LocalPort 6099) { Write-RecoveryState (Get-LoginState) }
    else { Write-RecoveryState 'timeout' }
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Error 'This script must run elevated. Install the scheduled task first, then let restart_napcat.bat trigger it.'
    exit 10
}

if (-not (Test-Path $ConfigPath)) {
    Write-Error 'config.json was not found.'
    exit 2
}

if (-not (Test-Path $Launcher)) {
    Write-Error "NapCat launcher was not found: $Launcher"
    exit 3
}

if (-not (Test-Path $CredentialPath)) {
    Write-Error 'napcat_password.dpapi was not found. Run setup_napcat_credential.ps1 first.'
    exit 4
}

$config = Get-Content -Raw -Encoding UTF8 $ConfigPath | ConvertFrom-Json

$qq = ''
if ($null -ne $config.PSObject.Properties['napcat_login_qq']) {
    $qq = ([string]$config.napcat_login_qq).Trim()
}

if ($qq -notmatch '^\d+$') {
    Write-Error 'config.json does not contain a valid napcat_login_qq.'
    exit 5
}

$endpoint = [Uri]$config.ws_url
if ($endpoint.Host -notin @('127.0.0.1','localhost','::1')) { throw 'Local recovery requires a loopback OneBot address.' }
$OneBotPort = $endpoint.Port
# Never start another instance over an existing listener or known launcher.
$existingLauncher = @(Get-CimInstance Win32_Process -Filter "name='NapCatWinBootMain.exe'" -ErrorAction Stop | Where-Object { $_.ExecutablePath -eq (Join-Path $NapCatDir 'NapCatWinBootMain.exe') })
if ((Test-LocalPort 6099) -or (Test-LocalPort $OneBotPort) -or $existingLauncher.Count -gt 0) {
    Write-RecoveryState 'existing_unverified'
    Wait-Recovery
    exit 0
}
Write-RecoveryState 'starting'

$secure = (Get-Content -Raw -Encoding UTF8 $CredentialPath) | ConvertTo-SecureString
$bstr = [IntPtr]::Zero
$plain = $null

try {
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)

    if ([string]::IsNullOrEmpty($plain)) {
        throw 'The decrypted password is empty.'
    }

    # NapCat supports these environment variables for automatic fallback login.
    $env:NAPCAT_QUICK_ACCOUNT = $qq
    $env:NAPCAT_QUICK_PASSWORD = $plain

    # launcher-win10.bat forwards "-q <QQ>" to NapCat.
    # Because this script already runs elevated, launcher-win10.bat should not
    # need to open another UAC elevation window.
    $cmdLine = '"' + $Launcher + '" -q ' + $qq
    Start-Process -FilePath 'cmd.exe' `
        -ArgumentList @('/c', $cmdLine) `
        -WorkingDirectory $NapCatDir `
        -WindowStyle Hidden | Out-Null
}
finally {
    if ($bstr -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }

    $plain = $null
    Remove-Item Env:NAPCAT_QUICK_PASSWORD -ErrorAction SilentlyContinue
    Remove-Item Env:NAPCAT_QUICK_ACCOUNT -ErrorAction SilentlyContinue
}

Wait-Recovery
exit 0
