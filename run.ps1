$env:PYTHONDONTWRITEBYTECODE = "1"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot
$env:PYTHONPATH = Join-Path $projectRoot "src"

python -m excel_comparator
