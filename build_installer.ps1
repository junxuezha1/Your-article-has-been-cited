param(
    [switch]$SkipRuntime
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

$argsList = @("tools/installer_build.py")
if ($SkipRuntime) {
    $argsList += "--skip-runtime"
}

python @argsList
exit $LASTEXITCODE
