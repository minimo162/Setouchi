Param(
    [switch]$Release
)

$ErrorActionPreference = 'Stop'

function Invoke-Cargo {
    param(
        [string]$Command
    )
    if (-not (Get-Command cargo -ErrorAction SilentlyContinue)) {
        Write-Error "cargo が見つかりません。Rust ツールチェーンをインストールしてください。"
    }
    Push-Location "$PSScriptRoot/../apps/shell"
    try {
        if ($Release) {
            cargo tauri $Command --release
        } else {
            cargo tauri $Command
        }
    }
    finally {
        Pop-Location
    }
}

Invoke-Cargo -Command 'dev'
