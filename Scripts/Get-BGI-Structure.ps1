<#
.SYNOPSIS
    BetterGI 雙倉庫自動更新與架構提取工具
    
.DESCRIPTION
    1. 自動 Clone 或 Pull 指定的 GitHub 倉庫 (無需人工確認)。
    2. 生成目錄結構樹。
    3. 提取關鍵設定檔與說明檔內容。
    
.NOTES
    Author: Gemini
    Date: 2025-12-03
#>

# --- 設定區域 ---
$outputFilename = "BGI_Full_Architecture.txt"
$repos = @(
    @{
        Name = "bettergi-scripts-list"
        Url = "https://github.com/babalae/bettergi-scripts-list.git"
        # 讀取內容的副檔名
        ReadContentExt = @(".json", ".md", ".js", ".ts", ".d.ts")
        # 忽略的目錄
        IgnoreDirs = @(".git", ".github", "node_modules", "build", "archive", "dist")
    },
    @{
        Name = "better-genshin-impact"
        Url = "https://github.com/babalae/better-genshin-impact.git"
        # 主程式倉庫只讀取架構定義檔，避免檔案過大
        ReadContentExt = @(".md", ".sln", ".csproj", ".config", ".yml", ".xaml")
        IgnoreDirs = @(".git", ".github", "bin", "obj", "packages", "Build", "Assets")
    }
)
$maxFileSizeKB = 300 # 超過此大小只讀取前 50 行
# ----------------

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$outputFilePath = Join-Path $scriptPath $outputFilename

# 檢查 Git 是否安裝
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "錯誤：未檢測到 Git。請先安裝 Git (https://git-scm.com/) 才能使用自動下載功能。" -ForegroundColor Red
    Read-Host "按 Enter 退出"
    exit
}

# 初始化輸出檔案
$null = New-Item -Path $outputFilePath -ItemType File -Force

function Write-Log {
    param ([string]$Message, [switch]$Header)
    if ($Header) {
        $separator = "=" * 60
        $formattedMsg = "`r`n$separator`r`n$Message`r`n$separator`r`n"
    } else {
        $formattedMsg = "$Message`r`n"
    }
    Write-Host $Message -ForegroundColor Cyan
    Add-Content -Path $outputFilePath -Value $formattedMsg -Encoding UTF8
}

function Get-Tree {
    param ([string]$Path, [array]$IgnoreList, [string]$Prefix = "", [bool]$IsLast = $true)
    
    $items = Get-ChildItem -Path $Path -Force | Where-Object { 
        $_.Name -notin $IgnoreList -and $_.Name -ne $outputFilename 
    }
    
    $count = $items.Count
    $i = 0

    foreach ($item in $items) {
        $i++
        $isLastItem = ($i -eq $count)
        $currentPrefix = if ($isLastItem) { "└── " } else { "├── " }
        $childPrefix = if ($isLastItem) { "    " } else { "│   " }
        
        $line = "$Prefix$currentPrefix$($item.Name)"
        Add-Content -Path $outputFilePath -Value $line -Encoding UTF8
        
        if ($item.PSIsContainer) {
            Get-Tree -Path $item.FullName -IgnoreList $IgnoreList -Prefix "$Prefix$childPrefix" -IsLast $isLastItem
        }
    }
}

# --- 主程式開始 ---
Clear-Host
Write-Host "=== BetterGI 自動化架構提取工具 ===" -ForegroundColor Yellow
Write-Log "Report Generated on: $(Get-Date)" -Header

# 1. 自動下載或更新倉庫
foreach ($repo in $repos) {
    $repoPath = Join-Path $scriptPath $repo.Name
    
    Write-Host "`n正在處理: $($repo.Name)..." -ForegroundColor Green
    
    if (Test-Path $repoPath) {
        # 資料夾存在 -> 更新 (Git Pull)
        Write-Host "  -> 資料夾已存在，正在執行 Git Pull 更新..." -ForegroundColor Gray
        Push-Location $repoPath
        try {
            git pull
        } catch {
            Write-Host "  -> Git Pull 失敗 (可能是本地有修改衝突)，將使用當前版本分析。" -ForegroundColor Red
        }
        Pop-Location
    } else {
        # 資料夾不存在 -> 下載 (Git Clone)
        Write-Host "  -> 資料夾不存在，正在執行 Git Clone 下載..." -ForegroundColor Gray
        try {
            git clone $repo.Url $repoPath
        } catch {
            Write-Host "  -> Git Clone 失敗，跳過此倉庫。" -ForegroundColor Red
            continue
        }
    }
}

# 2. 執行分析與提取
Write-Host "`n開始分析檔案結構..." -ForegroundColor Yellow

foreach ($repo in $repos) {
    $repoPath = Join-Path $scriptPath $repo.Name
    
    if (Test-Path $repoPath) {
        Write-Host "正在提取: $($repo.Name)" -ForegroundColor Green
        Write-Log "REPOSITORY: $($repo.Name)" -Header
        
        # A. 生成目錄樹
        Write-Host "  - 掃描目錄結構"
        Add-Content -Path $outputFilePath -Value "`n[Directory Structure]`n" -Encoding UTF8
        Get-Tree -Path $repoPath -IgnoreList $repo.IgnoreDirs
        
        # B. 讀取關鍵檔案
        Write-Host "  - 讀取檔案內容"
        Add-Content -Path $outputFilePath -Value "`n[Key File Contents]`n" -Encoding UTF8
        
        $filesToRead = Get-ChildItem -Path $repoPath -Recurse | Where-Object {
            $ext = $_.Extension
            # 檢查副檔名是否在白名單中
            $isTargetExt = $repo.ReadContentExt -contains $ext
            # 特殊檔案強制讀取
            $isSpecialFile = ($_.Name -eq "repo.json") -or ($_.Name -eq "manifest.json")
            
            return ($isTargetExt -or $isSpecialFile) -and (-not $_.PSIsContainer)
        } | Where-Object {
            # 再次過濾忽略的路徑
            $path = $_.FullName
            $shouldIgnore = $false
            foreach ($ignore in $repo.IgnoreDirs) {
                if ($path -like "*\$ignore\*") { $shouldIgnore = $true; break }
            }
            -