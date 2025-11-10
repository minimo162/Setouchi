Param(
    [switch]$Release
)

$ErrorActionPreference = 'Stop'

function Invoke-Cargo {
    if (-not (Get-Command cargo -ErrorAction SilentlyContinue)) {
        Write-Error "cargo が見つかりません。Rust ツールチェーンをインストールしてください。"
    }
    Push-Location "$PSScriptRoot/../apps/shell"
    try {
        if ($Release) {
            cargo tauri build --release
        } else {
            cargo tauri build
        }
    }
    finally {
        Pop-Location
    }
}

Invoke-Cargo
