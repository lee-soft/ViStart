<#
.SYNOPSIS
    Builds ViStart (prjVistaStart.vbp) with the VB6 command-line compiler — no IDE session.

.DESCRIPTION
    - Ensures the three custom typelibs (winU, GDIplus, SHELLLNK) are registered per-user,
      registering them from C:\TLBS via 32-bit PowerShell if missing (SHELLLNK.TLB is an
      old-format typelib that 64-bit oleaut32 cannot load).
    - Stops any running copy of the output exe so the compiler can overwrite it.
    - Compiles with VB6.EXE /make and reports the result.

.EXAMPLE
    .\build.ps1                 # build to %APPDATA%\ViStart
    .\build.ps1 -KeepRunning    # don't kill a running dev build first (build fails if the exe is locked)
#>
[CmdletBinding()]
param(
    [string]$OutDir = (Join-Path $env:APPDATA 'ViStart'),
    [switch]$KeepRunning
)

$ErrorActionPreference = 'Stop'

$Vb6     = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE'
$Ps32    = "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
$Project = Join-Path $PSScriptRoot 'prjVistaStart.vbp'
$TlbDir  = 'C:\TLBS'
$LogFile = Join-Path $env:TEMP 'vistart-vb6-build.log'

if (-not (Test-Path $Vb6))     { throw "VB6 compiler not found at $Vb6" }
if (-not (Test-Path $Project)) { throw "Project file not found at $Project" }

# --- 1. Typelib references (per-user registration is persistent; normally a no-op) ---
$tlbs = @(
    @{ File = 'winU.tlb';     Guid = '{64675040-3A82-101B-8181-00AA003743D3}'; Ver = '1.1' },
    @{ File = 'GDIplus.tlb';  Guid = '{3C4B4A34-9EF8-42A6-87A2-D2383DDC1EDB}'; Ver = '1.0' },
    @{ File = 'SHELLLNK.TLB'; Guid = '{11269241-F241-11CF-BD9A-00AA00575603}'; Ver = '1.0' }
)

$reg32Helper = @'
param([string]$TlbPath)
$src = @"
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class TlbReg {
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    public static extern ITypeLib LoadTypeLibEx(string file, int regKind);
    [DllImport("oleaut32.dll", PreserveSig = false)]
    public static extern void RegisterTypeLibForUser(ITypeLib tlb, string fullPath, string helpDir);
    public static void Register(string path) {
        ITypeLib tlb = LoadTypeLibEx(path, 2); // REGKIND_NONE
        RegisterTypeLibForUser(tlb, path, null);
    }
}
"@
Add-Type -TypeDefinition $src
[TlbReg]::Register($TlbPath)
'@

# oleaut32 stores these per-user typelib paths as ANSI bytes in a REG_SZ, so the value
# reads back as CJK mojibake — reinterpret the UTF-16 units as bytes to recover the path.
function Resolve-TlbRegPath([string]$raw) {
    if (-not $raw) { return $null }
    if (Test-Path -LiteralPath $raw -ErrorAction SilentlyContinue) { return $raw }
    $ansi = [Text.Encoding]::Default.GetString([Text.Encoding]::Unicode.GetBytes($raw)).TrimEnd([char]0)
    if (Test-Path -LiteralPath $ansi -ErrorAction SilentlyContinue) { return $ansi }
    return $null
}

foreach ($t in $tlbs) {
    $regKey  = "HKCU:\Software\Classes\TypeLib\$($t.Guid)\$($t.Ver)\0\win32"
    $tlbPath = Join-Path $TlbDir $t.File
    $registeredPath = $null
    if (Test-Path $regKey) { $registeredPath = Resolve-TlbRegPath ((Get-ItemProperty $regKey).'(default)') }

    if ($registeredPath) { continue }

    Write-Host "Registering typelib $($t.File)..."
    if (-not (Test-Path $tlbPath)) {
        $fallback = Join-Path "$env:USERPROFILE\Downloads\TLB" $t.File
        if (Test-Path $fallback) {
            New-Item -ItemType Directory -Force $TlbDir | Out-Null
            Copy-Item $fallback $tlbPath -Force
        } else {
            throw "Typelib $($t.File) not found in $TlbDir or Downloads\TLB. Get it from https://lee-soft.com/bin/TLB.zip"
        }
    }
    # 32-bit PowerShell handles old-format typelibs (SHELLLNK.TLB) that 64-bit oleaut32 rejects
    $helperFile = Join-Path $env:TEMP 'register-tlb-32.ps1'
    Set-Content -Path $helperFile -Value $reg32Helper -Encoding UTF8
    & $Ps32 -NoProfile -ExecutionPolicy Bypass -File $helperFile -TlbPath $tlbPath
    if ($LASTEXITCODE -ne 0) { throw "Failed to register $tlbPath" }
}

# --- 2. Free the output exe if a previous build is still running ---
$exeName = ([regex]::Match((Get-Content $Project -Raw), 'ExeName32="([^"]+)"')).Groups[1].Value
if (-not $exeName) { throw "Could not parse ExeName32 from $Project" }
$exePath = Join-Path $OutDir $exeName

if (-not $KeepRunning) {
    $running = Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -and ($_.Path -ieq $exePath) }
    if ($running) {
        Write-Host "Stopping running $exeName (PID $($running.Id -join ', '))..."
        $running | Stop-Process -Force
        Start-Sleep -Milliseconds 500
    }
}

# --- 3. Compile ---
New-Item -ItemType Directory -Force $OutDir | Out-Null
Remove-Item $LogFile -ErrorAction SilentlyContinue

Write-Host "Compiling $([IO.Path]::GetFileName($Project)) -> $exePath"
$p = Start-Process -FilePath $Vb6 -ArgumentList '/make', "`"$Project`"", '/outdir', "`"$OutDir`"", '/out', "`"$LogFile`"" -PassThru
if (-not $p.WaitForExit(300000)) {
    $p.Kill()
    throw "VB6 compile timed out after 5 minutes (probably stuck on a modal dialog)."
}

$log = if (Test-Path $LogFile) { (Get-Content $LogFile -Raw).Trim() } else { '<no build log produced>' }
if ($p.ExitCode -ne 0 -or $log -notmatch 'succeeded') {
    Write-Host $log
    throw "Build FAILED (VB6 exit code $($p.ExitCode))."
}

# --- 4. Embed the application manifest (RT_MANIFEST id 1) ---
$Manifest = Join-Path $PSScriptRoot 'ViStart.exe.manifest'
if (-not (Test-Path $Manifest)) { throw "Manifest not found at $Manifest" }
$mt = Get-ChildItem 'C:\Program Files (x86)\Windows Kits\10\bin\*\x64\mt.exe' -ErrorAction SilentlyContinue |
      Sort-Object { [version]$_.Directory.Parent.Name } | Select-Object -Last 1
if (-not $mt) { throw "mt.exe not found under Windows Kits 10 (needed to embed the manifest)" }
& $mt.FullName -nologo -manifest "$Manifest" -outputresource:"$exePath;#1"
if ($LASTEXITCODE -ne 0) { throw "Manifest injection failed (mt.exe exit code $LASTEXITCODE)" }
Write-Host "Embedded $([IO.Path]::GetFileName($Manifest)) into $exeName"

$exe = Get-Item $exePath
Write-Host $log
Write-Host ("Output: {0}  v{1}  {2:N0} bytes  {3}" -f $exe.FullName, $exe.VersionInfo.FileVersion, $exe.Length, $exe.LastWriteTime)
