# --- 強制開啟 TLS 1.2 (解決連線被拒問題) ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- 設定區 ---
$owner = "babalae"
$repo = "better-genshin-impact"
$branch = "main"
$targetFolder = "BetterGenshinImpact"
$outputFile = "BGI_Structure.txt"

# --- 建構網址 (使用 -f 格式化，防止變數解析錯誤) ---
# {0}=owner, {1}=repo, {2}=branch
$apiUrl = "https://api.github.com/repos/{0}/{1}/git/trees/{2}?recursive=1" -f $owner, $repo, $branch
$absoluteOutputPath = Join-Path $PWD $outputFile

# --- 顯示除錯資訊 ---
Clear-Host
Write-Host "正在執行 BGI 結構擷取腳本..." -ForegroundColor Cyan
Write-Host "------------------------------------------------" -ForegroundColor DarkGray
Write-Host "目標倉庫: $owner / $repo"
Write-Host "目標分支: $branch"
Write-Host "API 網址: $apiUrl" -ForegroundColor Yellow
Write-Host "------------------------------------------------" -ForegroundColor DarkGray

# --- 偽裝瀏覽器 Header ---
$headers = @{
    "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

try {
    Write-Host "正在連線 GitHub..." -ForegroundColor Cyan
    $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers -ErrorAction Stop

    Write-Host "連線成功，正在篩選 '$targetFolder' 目錄..." -ForegroundColor Cyan
    
    # 篩選並只取路徑
    $structure = $response.tree | Where-Object { $_.path -like "$targetFolder/*" } | Select-Object -ExpandProperty path

    if ($structure) {
        $structure | Out-File -FilePath $absoluteOutputPath -Encoding UTF8
        Write-Host "------------------------------------------------" -ForegroundColor Green
        Write-Host "★ 成功！" -ForegroundColor Green
        Write-Host "檔案已儲存至: $absoluteOutputPath" -ForegroundColor White
        Write-Host "共擷取到 $($structure.Count) 筆資料。" -ForegroundColor Gray
        Write-Host "------------------------------------------------" -ForegroundColor Green
    } else {
        Write-Warning "錯誤：連線成功但在該分支找不到 '$targetFolder' 資料夾。"
    }
}
catch {
    Write-Error "發生錯誤！"
    Write-Error $_.Exception.Message
}

Read-Host "執行完畢，請按 Enter 鍵離開..."