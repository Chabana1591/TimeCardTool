#エクセルファイル配置場所
$excelDir = (Join-Path $env:USERPROFILE "testdir")
$excelFileName = "test.xlsx"
$excelSheetName = "Sheet1"

#セル地点
$startRow = 3
$startCol = 3
$startDate = [DateTime]::ParseExact("2021/02/01", "yyyy/MM/dd", $null)
$currentDate = Get-Date

#時刻用の日付(Excelでは1900/01/00が0のため)
$baseDate = [DateTime]::ParseExact("1900/01/01", "yyyy/MM/dd", $null)
$baseOffsetDays = -2 

#テキスト出力用
$outDir = (Join-Path $env:USERPROFILE "testdir")
$outTxtFileName = ((Get-Date -Format "yyyyMMdd")+".txt")

#デフォルトの打刻
$defaultArriveTimeStr = "09:00"
$defaultLeaveTimeStr = "17:00"
$defaultRestTimeStr = "00:50"












