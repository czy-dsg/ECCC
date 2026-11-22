param(
    [string]$Repo = "https://github.com/czy-dsg/ECCC.git",
    [string]$Message = "Add ECC compression self-sensing curve generator"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "未找到 git。请先安装：winget install --id Git.Git -e --source winget" -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path ".git")) { git init | Out-Host }
git add .
git commit -m $Message
git branch -M main

$existing = git remote
if ($existing -contains "origin") { git remote remove origin }
git remote add origin $Repo

Write-Host "即将推送到 $Repo" -ForegroundColor Cyan
git push -u origin main
