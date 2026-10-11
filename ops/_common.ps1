# Shared helpers for the setup scripts (dot-sourced: . "$PSScriptRoot\_common.ps1").
# Keep this file UTF-8 WITH BOM so Windows PowerShell 5.1 reads the Korean text correctly.

$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Write-Utf8NoBom([string]$Path, [string]$Text) {
    # A BOM at the start of .env would glue itself to the first key name and break it.
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllText($Path, $Text, $script:Utf8NoBom)
}

function Read-TextLines([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $text = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)  # drops a BOM if one exists
    if ($text.Length -eq 0) { return @() }
    return @($text -split "\r?\n" | Where-Object { $_ -ne '' })
}

function Clean-Value([string]$Value) {
    # Pasted values often carry spaces, quotes or invisible characters.
    if ($null -eq $Value) { return '' }
    $v = $Value -replace '[​-‍﻿]', ''
    $v = $v.Trim()
    $v = $v.Trim([char]34, [char]39, [char]0x201C, [char]0x201D, [char]0x2018, [char]0x2019)
    return $v.Trim()
}

function Read-EnvMap([string]$Path) {
    $map = [ordered]@{}
    foreach ($line in (Read-TextLines $Path)) {
        if ($line.TrimStart().StartsWith('#') -or -not $line.Contains('=')) { continue }
        $pair = $line.Split('=', 2)
        $map[$pair[0].Trim()] = $pair[1].Trim()
    }
    return $map
}

function Update-EnvFile([string]$Path, [System.Collections.IDictionary]$Updates) {
    # Replaces the lines of the given keys, appends missing ones, keeps every other line as it was.
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($l in (Read-TextLines $Path)) { $lines.Add($l) }
    foreach ($key in $Updates.Keys) {
        $newLine = $key + '=' + $Updates[$key]
        $done = $false
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i].Contains('=') -and -not $lines[$i].TrimStart().StartsWith('#') -and
                $lines[$i].Split('=', 2)[0].Trim() -eq $key) {
                $lines[$i] = $newLine
                $done = $true
                break
            }
        }
        if (-not $done) { $lines.Add($newLine) }
    }
    Write-Utf8NoBom $Path (($lines -join "`r`n") + "`r`n")
}

function Get-PythonInfo {
    # Finds a real Python >= 3.11. Rejects the Microsoft Store stub (it lives under WindowsApps).
    $ErrorActionPreference = 'Continue'
    $problem = 'python을 찾을 수 없습니다.'
    $candidates = @(
        @{ Exe = 'python'; Args = @() },
        @{ Exe = 'py'; Args = @('-3') }
    )
    foreach ($c in $candidates) {
        if (-not (Get-Command $c.Exe -ErrorAction SilentlyContinue)) { continue }
        $out = $null
        try {
            $out = @(& $c.Exe @($c.Args) -c "import sys; print(sys.executable); print(sys.version_info[0]); print(sys.version_info[1])" 2>$null)
        } catch { continue }
        if ($LASTEXITCODE -ne 0 -or $out.Count -lt 3) { continue }
        $path = ([string]$out[0]).Trim()
        if (-not (Test-Path -LiteralPath $path)) { continue }
        if ($path -match '\\WindowsApps\\') {
            $problem = 'Microsoft Store용 가짜 python이 잡혔습니다.'
            continue
        }
        $major = [int]$out[1]; $minor = [int]$out[2]
        if ($major -lt 3 -or ($major -eq 3 -and $minor -lt 11)) {
            $problem = "파이썬 버전이 낮습니다 ($major.$minor). 3.11 이상이 필요합니다."
            continue
        }
        return [pscustomobject]@{ Ok = $true; Path = $path; Version = "$major.$minor"; Problem = '' }
    }
    return [pscustomobject]@{ Ok = $false; Path = ''; Version = ''; Problem = $problem }
}
