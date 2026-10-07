param(
    [string]$Environment = ""
)

$env:PYTHONDONTWRITEBYTECODE = "1"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot
$env:PYTHONPATH = Join-Path $projectRoot "src"

if ($Environment) {
    $env:EXCEL_COMPARATOR_ENV = $Environment.ToUpper()
} else {
    Remove-Item Env:EXCEL_COMPARATOR_ENV -ErrorAction SilentlyContinue
}

python -m excel_comparator
