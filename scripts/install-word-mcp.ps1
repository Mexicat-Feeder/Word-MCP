[CmdletBinding()]
param(
    [ValidateSet("stdio", "streamable-http", "sse")]
    [string]$Transport = "stdio",

    [string]$ServerName = "word",
    [string]$Author = "",
    [string]$AuthorInitials = "",
    [string]$HostAddress = "127.0.0.1",
    [int]$Port = 8000,
    [string]$McpPath = "/mcp",
    [string]$SsePath = "/sse",
    [string]$ConfigOut = "mcp-config.json",

    [switch]$SkipUvInstall,
    [switch]$SkipTests
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
        & $UvPath @Arguments
        if ($LASTEXITCODE -ne 0) {
            throw "uv $($Arguments -join ' ') failed with exit code $LASTEXITCODE"
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

function New-McpConfig {
    param(
        [string]$RepoRoot,
        [string]$UvPath,
        [string]$ServerName,
        [string]$Transport,
        [string]$Author,
        [string]$AuthorInitials,
        [string]$HostAddress,
        [int]$Port,
        [string]$McpPath,
        [string]$SsePath
    )

    $envConfig = [ordered]@{
        MCP_TRANSPORT = $Transport
        MCP_AUTHOR = $Author
        MCP_AUTHOR_INITIALS = $AuthorInitials
        PYTHONIOENCODING = "utf-8"
    }

    if ($Transport -eq "streamable-http") {
        $envConfig["MCP_HOST"] = $HostAddress
        $envConfig["MCP_PORT"] = "$Port"
        $envConfig["MCP_PATH"] = $McpPath
    }
    elseif ($Transport -eq "sse") {
        $envConfig["MCP_HOST"] = $HostAddress
        $envConfig["MCP_PORT"] = "$Port"
        $envConfig["MCP_SSE_PATH"] = $SsePath
    }

    return [ordered]@{
        mcpServers = [ordered]@{
            $ServerName = [ordered]@{
                command = $UvPath
                args = @(
                    "--directory",
                    $RepoRoot,
                    "run",
                    "python",
                    "-m",
                    "word_document_server.main"
                )
                env = $envConfig
            }
        }
    }
}

$repoRoot = Resolve-RepoRoot
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

$config = New-McpConfig `
    -RepoRoot $repoRoot `
    -UvPath $uvPath `
    -ServerName $ServerName `
    -Transport $Transport `
    -Author $Author `
    -AuthorInitials $AuthorInitials `
    -HostAddress $HostAddress `
    -Port $Port `
    -McpPath $McpPath `
    -SsePath $SsePath

$configJson = $config | ConvertTo-Json -Depth 20
$configPath = if ([IO.Path]::IsPathRooted($ConfigOut)) {
    $ConfigOut
}
else {
    Join-Path $repoRoot $ConfigOut
}
Set-Content -LiteralPath $configPath -Value $configJson -Encoding UTF8

Write-Step "MCP client config"
Write-Host "Wrote: $configPath"
Write-Host ""
Write-Output $configJson

Write-Host ""
Write-Host "MCP Inspector stdio fields:" -ForegroundColor Green
Write-Host "Command: $uvPath"
Write-Host "Arguments: --directory `"$repoRoot`" run python -m word_document_server.main"
