[CmdletBinding()]
param(
    [ValidateSet("prompt", "hermes", "openclaw", "custom")]
    [string]$Target = "prompt",

    [string]$ServerName = "word",

    [ValidateSet("stdio", "streamable-http", "sse")]
    [string]$Transport = "stdio",

    [string]$Author = "",
    [string]$AuthorInitials = "",
    [string]$HostAddress = "127.0.0.1",
    [int]$Port = 8000,
    [string]$McpPath = "/mcp",
    [string]$SsePath = "/sse",
    [string]$ConfigOut = "",
    [int]$Timeout = 180,
    [int]$ConnectTimeout = 30,

    [switch]$SkipUvInstall,
    [switch]$SkipTests,
    [switch]$SkipRegister,
    [switch]$SkipProbe
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Resolve-RepoRoot {
    if ($PSScriptRoot) {
        return (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
    }
    return (Get-Location).Path
}

function Resolve-InstallTarget {
    param([string]$Target)

    if ($Target -ne "prompt") {
        return $Target
    }

    Write-Step "Select MCP client target"
    Write-Host "1. hermes   - register with 'hermes mcp add' when available"
    Write-Host "2. openclaw - register with 'openclaw mcp set' when available"
    Write-Host "3. custom   - write and print the MCP config only"
    $selection = Read-Host "Target (1-3, default: custom)"

    $normalized = ""
    if ($selection) {
        $normalized = $selection.Trim().ToLowerInvariant()
    }

    switch ($normalized) {
        "" { return "custom" }
        "1" { return "hermes" }
        "hermes" { return "hermes" }
        "2" { return "openclaw" }
        "openclaw" { return "openclaw" }
        "3" { return "custom" }
        "custom" { return "custom" }
        default { throw "Unknown target '$selection'. Use hermes, openclaw, or custom." }
    }
}

function Get-RequiredPythonVersion {
    param([string]$RepoRoot)

    $pythonVersionPath = Join-Path $RepoRoot ".python-version"
    if (Test-Path -LiteralPath $pythonVersionPath) {
        $raw = (Get-Content -LiteralPath $pythonVersionPath -Raw).Trim()
        $match = [regex]::Match($raw, "\d+\.\d+(?:\.\d+)?")
        if ($match.Success) {
            return $match.Value
        }
    }

    $pyprojectPath = Join-Path $RepoRoot "pyproject.toml"
    if (Test-Path -LiteralPath $pyprojectPath) {
        $content = Get-Content -LiteralPath $pyprojectPath -Raw
        $match = [regex]::Match($content, 'requires-python\s*=\s*">=\s*(\d+\.\d+)')
        if ($match.Success) {
            return $match.Groups[1].Value
        }
    }

    return "3.12"
}

function Find-Uv {
    $cmd = Get-Command uv -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $candidates = @(
        (Join-Path $env:USERPROFILE ".local\bin\uv.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\uv\uv.exe"),
        (Join-Path $env:USERPROFILE ".cargo\bin\uv.exe")
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $null
}

function Install-Uv {
    Write-Step "Installing uv"
    Write-Host "Using the official uv standalone installer from https://astral.sh/uv/install.ps1"
    powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://astral.sh/uv/install.ps1 | iex"
}

function Invoke-Uv {
    param(
        [string]$UvPath,
        [string[]]$Arguments,
        [string]$WorkingDirectory
    )

    Push-Location -LiteralPath $WorkingDirectory
    try {
        $previousErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try {
            $output = & $UvPath @Arguments 2>&1
            $exitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $previousErrorActionPreference
        }
        if ($output) {
            $output | ForEach-Object { Write-Host $_ }
        }
        if ($exitCode -ne 0) {
            throw "uv $($Arguments -join ' ') failed with exit code $exitCode"
        }
    }
    finally {
        Pop-Location
    }
}

function Ensure-DefaultEnvFile {
    param(
        [string]$RepoRoot,
        [string]$Transport,
        [string]$Author,
        [string]$AuthorInitials,
        [string]$HostAddress,
        [int]$Port,
        [string]$McpPath,
        [string]$SsePath
    )

    $envPath = Join-Path $RepoRoot ".env"
    if (Test-Path -LiteralPath $envPath) {
        Write-Host ".env already exists; leaving it unchanged."
        return
    }

    $lines = @(
        "MCP_TRANSPORT=$Transport",
        "MCP_AUTHOR=$Author",
        "MCP_AUTHOR_INITIALS=$AuthorInitials"
    )

    if ($Transport -eq "streamable-http") {
        $lines += @(
            "MCP_HOST=$HostAddress",
            "MCP_PORT=$Port",
            "MCP_PATH=$McpPath"
        )
    }
    elseif ($Transport -eq "sse") {
        $lines += @(
            "MCP_HOST=$HostAddress",
            "MCP_PORT=$Port",
            "MCP_SSE_PATH=$SsePath"
        )
    }

    Set-Content -LiteralPath $envPath -Value $lines -Encoding UTF8
    Write-Host "Wrote default .env"
}

function Get-ProjectPythonPath {
    param([string]$RepoRoot)

    if ($IsWindows -or $env:OS -eq "Windows_NT") {
        return (Join-Path $RepoRoot ".venv\Scripts\python.exe")
    }
    return (Join-Path $RepoRoot ".venv/bin/python")
}

function Get-ProjectScriptPath {
    param(
        [string]$RepoRoot,
        [string]$ScriptName
    )

    if ($IsWindows -or $env:OS -eq "Windows_NT") {
        return (Join-Path $RepoRoot ".venv\Scripts\$ScriptName.exe")
    }
    return (Join-Path $RepoRoot ".venv/bin/$ScriptName")
}

function Get-McpUrl {
    param(
        [string]$Transport,
        [string]$HostAddress,
        [int]$Port,
        [string]$McpPath,
        [string]$SsePath
    )

    $path = if ($Transport -eq "sse") { $SsePath } else { $McpPath }
    return "http://${HostAddress}:$Port$path"
}

function New-HermesMcpConfig {
    param(
        [string]$RepoRoot,
        [string]$PythonPath,
        [string]$Transport,
        [string]$HostAddress,
        [int]$Port,
        [string]$McpPath,
        [string]$SsePath,
        [int]$Timeout,
        [int]$ConnectTimeout
    )

    if ($Transport -ne "stdio") {
        return [ordered]@{
            url = Get-McpUrl -Transport $Transport -HostAddress $HostAddress -Port $Port -McpPath $McpPath -SsePath $SsePath
            timeout = $Timeout
            connect_timeout = $ConnectTimeout
        }
    }

    $envConfig = [ordered]@{
        MCP_TRANSPORT = $Transport
        PYTHONPATH = $RepoRoot
    }

    return [ordered]@{
        command = $PythonPath
        args = @(
            "-m",
            "word_document_server.main"
        )
        env = $envConfig
        timeout = $Timeout
        connect_timeout = $ConnectTimeout
    }
}

function New-OpenClawMcpConfig {
    param(
        [string]$RepoRoot,
        [string]$PythonPath,
        [string]$Transport,
        [string]$HostAddress,
        [int]$Port,
        [string]$McpPath,
        [string]$SsePath,
        [int]$Timeout,
        [int]$ConnectTimeout
    )

    if ($Transport -ne "stdio") {
        return [ordered]@{
            url = Get-McpUrl -Transport $Transport -HostAddress $HostAddress -Port $Port -McpPath $McpPath -SsePath $SsePath
            transport = $Transport
            timeout = $Timeout
            connectTimeout = $ConnectTimeout
        }
    }

    return [ordered]@{
        command = $PythonPath
        args = @(
            "-m",
            "word_document_server.main"
        )
        cwd = $RepoRoot
        env = [ordered]@{
            MCP_TRANSPORT = "stdio"
        }
        timeout = $Timeout
        connectTimeout = $ConnectTimeout
    }
}

function Invoke-ClientCli {
    param(
        [string]$Executable,
        [string[]]$Arguments,
        [string]$InputText = ""
    )

    Write-Host "> $Executable $($Arguments -join ' ')" -ForegroundColor DarkGray
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        if ($InputText) {
            $output = $InputText | & $Executable @Arguments 2>&1
        }
        else {
            $output = & $Executable @Arguments 2>&1
        }
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    if ($output) {
        $output | ForEach-Object { Write-Host $_ }
    }
    if ($exitCode -ne 0) {
        throw "$Executable $($Arguments -join ' ') failed with exit code $exitCode"
    }
}

function Write-HermesManualCommand {
    param(
        [string]$ServerName,
        [string]$RepoRoot,
        [string]$CommandPath,
        [string]$Transport,
        [string]$Url
    )

    Write-Host "Manual Hermes command:" -ForegroundColor Yellow
    if ($Transport -eq "stdio") {
        Write-Host "hermes mcp add $ServerName --command `"$CommandPath`" --env MCP_TRANSPORT=stdio PYTHONPATH=`"$RepoRoot`""
    }
    else {
        Write-Host "hermes mcp add $ServerName --url `"$Url`""
    }
}

function Register-HermesMcp {
    param(
        [string]$ServerName,
        [string]$RepoRoot,
        [string]$CommandPath,
        [string]$Transport,
        [string]$Url
    )

    $hermes = Get-Command hermes -ErrorAction SilentlyContinue
    if (-not $hermes) {
        Write-Warning "Hermes CLI was not found on PATH. The config was still written."
        Write-HermesManualCommand -ServerName $ServerName -RepoRoot $RepoRoot -CommandPath $CommandPath -Transport $Transport -Url $Url
        return
    }

    Write-Step "Registering MCP server with Hermes"
    if ($Transport -eq "stdio") {
        $arguments = @(
            "mcp",
            "add",
            $ServerName,
            "--command",
            $CommandPath,
            "--env",
            "MCP_TRANSPORT=stdio",
            "PYTHONPATH=$RepoRoot"
        )
    }
    else {
        $arguments = @(
            "mcp",
            "add",
            $ServerName,
            "--url",
            $Url
        )
    }

    try {
        Invoke-ClientCli -Executable $hermes.Source -Arguments $arguments -InputText "Y"
    }
    catch {
        Write-Warning "Hermes registration failed: $_"
        Write-HermesManualCommand -ServerName $ServerName -RepoRoot $RepoRoot -CommandPath $CommandPath -Transport $Transport -Url $Url
    }
}

function Write-OpenClawManualCommand {
    param(
        [string]$ServerName,
        [string]$ConfigJson,
        [bool]$SkipProbe
    )

    Write-Host "Manual OpenClaw commands:" -ForegroundColor Yellow
    Write-Host "openclaw mcp set $ServerName '$ConfigJson'"
    if (-not $SkipProbe) {
        Write-Host "openclaw mcp doctor $ServerName --probe"
    }
}

function Register-OpenClawMcp {
    param(
        [string]$ServerName,
        [object]$Config,
        [bool]$SkipProbe
    )

    $configJson = $Config | ConvertTo-Json -Depth 20 -Compress
    $openclaw = Get-Command openclaw -ErrorAction SilentlyContinue
    if (-not $openclaw) {
        Write-Warning "OpenClaw CLI was not found on PATH. The config was still written."
        Write-OpenClawManualCommand -ServerName $ServerName -ConfigJson $configJson -SkipProbe $SkipProbe
        return
    }

    Write-Step "Registering MCP server with OpenClaw"
    try {
        Invoke-ClientCli -Executable $openclaw.Source -Arguments @("mcp", "set", $ServerName, $configJson)
        if (-not $SkipProbe) {
            Invoke-ClientCli -Executable $openclaw.Source -Arguments @("mcp", "doctor", $ServerName, "--probe")
        }
    }
    catch {
        Write-Warning "OpenClaw registration failed: $_"
        Write-OpenClawManualCommand -ServerName $ServerName -ConfigJson $configJson -SkipProbe $SkipProbe
    }
}

$repoRoot = Resolve-RepoRoot
$installTarget = Resolve-InstallTarget -Target $Target
$requiredPython = Get-RequiredPythonVersion -RepoRoot $repoRoot

if (-not $Author) {
    $Author = if ($env:MCP_AUTHOR) { $env:MCP_AUTHOR } else { "AI Agent" }
}
if (-not $AuthorInitials) {
    $AuthorInitials = if ($env:MCP_AUTHOR_INITIALS) { $env:MCP_AUTHOR_INITIALS } else { "AI" }
}

Write-Step "Preparing Word MCP Live"
Write-Host "Repository: $repoRoot"
Write-Host "Requested Python: $requiredPython"
Write-Host "MCP client target: $installTarget"
Write-Host "MCP server name: $ServerName"
Write-Host "Transport: $Transport"

$uvPath = Find-Uv
if (-not $uvPath) {
    if ($SkipUvInstall) {
        throw "uv was not found and -SkipUvInstall was supplied."
    }
    Install-Uv
    $uvPath = Find-Uv
}

if (-not $uvPath) {
    throw "uv installation finished, but uv.exe was not found. Restart PowerShell or add uv to PATH, then rerun this script."
}

$uvDir = Split-Path -Parent $uvPath
if (($env:Path -split [IO.Path]::PathSeparator) -notcontains $uvDir) {
    $env:Path = "$uvDir$([IO.Path]::PathSeparator)$env:Path"
}

Write-Host "uv: $uvPath"

Write-Step "Installing Python $requiredPython with uv if needed"
Invoke-Uv -UvPath $uvPath -Arguments @("python", "install", $requiredPython) -WorkingDirectory $repoRoot

Write-Step "Syncing project dependencies"
Invoke-Uv -UvPath $uvPath -Arguments @("sync") -WorkingDirectory $repoRoot

$pythonPath = Get-ProjectPythonPath -RepoRoot $repoRoot
if (-not (Test-Path -LiteralPath $pythonPath)) {
    throw "Expected project Python was not found at $pythonPath after uv sync."
}
$wordMcpCommandPath = Get-ProjectScriptPath -RepoRoot $repoRoot -ScriptName "word_mcp_server"
if (-not (Test-Path -LiteralPath $wordMcpCommandPath)) {
    throw "Expected console script was not found at $wordMcpCommandPath after uv sync."
}

Ensure-DefaultEnvFile `
    -RepoRoot $repoRoot `
    -Transport $Transport `
    -Author $Author `
    -AuthorInitials $AuthorInitials `
    -HostAddress $HostAddress `
    -Port $Port `
    -McpPath $McpPath `
    -SsePath $SsePath

if (-not $SkipTests) {
    Write-Step "Running install smoke tests"
    Invoke-Uv -UvPath $uvPath -Arguments @("run", "python", "-m", "py_compile", "word_document_server/main.py", "word_document_server/tools/live_v2_tools.py") -WorkingDirectory $repoRoot
    Invoke-Uv -UvPath $uvPath -Arguments @("run", "python", "-m", "pytest", "-q") -WorkingDirectory $repoRoot
}

if ($installTarget -eq "openclaw") {
    $config = New-OpenClawMcpConfig `
        -RepoRoot $repoRoot `
        -PythonPath $pythonPath `
        -Transport $Transport `
        -HostAddress $HostAddress `
        -Port $Port `
        -McpPath $McpPath `
        -SsePath $SsePath `
        -Timeout $Timeout `
        -ConnectTimeout $ConnectTimeout
}
else {
    $config = New-HermesMcpConfig `
        -RepoRoot $repoRoot `
        -PythonPath $pythonPath `
        -Transport $Transport `
        -HostAddress $HostAddress `
        -Port $Port `
        -McpPath $McpPath `
        -SsePath $SsePath `
        -Timeout $Timeout `
        -ConnectTimeout $ConnectTimeout
}

$configJson = $config | ConvertTo-Json -Depth 20

if (-not $ConfigOut) {
    $ConfigOut = "word-mcp-$installTarget.json"
}

$configPath = if ([IO.Path]::IsPathRooted($ConfigOut)) {
    $ConfigOut
}
else {
    Join-Path $repoRoot $ConfigOut
}
Set-Content -LiteralPath $configPath -Value $configJson -Encoding UTF8

if ($SkipRegister) {
    Write-Host "Skipping MCP client registration because -SkipRegister was supplied."
}
elseif ($installTarget -eq "hermes") {
    Register-HermesMcp `
        -ServerName $ServerName `
        -RepoRoot $repoRoot `
        -CommandPath $wordMcpCommandPath `
        -Transport $Transport `
        -Url (Get-McpUrl -Transport $Transport -HostAddress $HostAddress -Port $Port -McpPath $McpPath -SsePath $SsePath)
}
elseif ($installTarget -eq "openclaw") {
    Register-OpenClawMcp `
        -ServerName $ServerName `
        -Config $config `
        -SkipProbe ([bool]$SkipProbe)
}
else {
    Write-Host "Custom target selected; no MCP client registration was attempted."
}

Write-Step "$installTarget MCP config"
Write-Host "Wrote: $configPath"
if ($installTarget -eq "hermes") {
    Write-Host "This output is Hermes-specific. Add it as the server entry for your Hermes '$ServerName' MCP server if you do not use the CLI registration above." -ForegroundColor Green
}
elseif ($installTarget -eq "openclaw") {
    Write-Host "This output is OpenClaw-specific. For stdio, it uses cwd instead of PYTHONPATH for startup safety." -ForegroundColor Green
}
else {
    Write-Host "This output is a generic MCP client config object." -ForegroundColor Green
}
Write-Host ""
Write-Output $configJson
