# Defaultの定数
$excelDir = (Join-Path $env:USERPROFILE "testdir")
$excelFileName = "test.xlsx"
$excelSheetName = "Sheet1"
$outDir = (Join-Path $env:USERPROFILE "testdir")
$outTxtFileName = ((Get-Date -Format "yyyyMMdd")+".txt")
$defaultArriveTimeStr = "09:00"
$defaultLeaveTimeStr = "17:00"
$defaultRestTimeStr = "00:50"
$baseDate = [DateTime]::ParseExact("1900/01/01", "yyyy/MM/dd", $null)
$baseOffsetDays = -2 
$startRow = 3
$startCol = 3
$startDate = [DateTime]::ParseExact("2021/02/01", "yyyy/MM/dd", $null)
$currentDate = Get-Date

# 定数をincludeする
. (Join-Path $PSScriptRoot "config.ps1")

do{
    $arriveTimeStr = Read-Host "出勤時刻を入力（デフォルト：`"$defaultArriveTimeStr`"）"
    # 空白時にデフォルト値を設定する
    if(!$arriveTimeStr){
        $arriveTimeStr = $defaultArriveTimeStr
    }
    # 適当な値で初期化
    [datetime]$arriveTime = [datetime]::MinValue

    # 変換して $parsedDate へ。失敗したら $parseSuccess が $false。
    [bool]$isArriveSuccess = [DateTime]::TryParseExact(($baseDate.toString("yyyy/MM/dd")+" "+$arriveTimeStr),
        "yyyy/MM/dd HH:mm", 
        [Globalization.DateTimeFormatInfo]::CurrentInfo,
        [Globalization.DateTimeStyles]::AllowWhiteSpaces, 
        [ref]$arriveTime)
}while(!$isArriveSuccess)

do{
    $leaveTimeStr = Read-Host "退勤時刻を入力（デフォルト：`"$defaultLeaveTimeStr`"）"
    # 空白時にデフォルト値を設定する
    if(!$leaveTimeStr){
        $leaveTimeStr = $defaultLeaveTimeStr
    }
    # 適当な値で初期化
    [datetime]$leaveTime = [datetime]::MinValue
    # 変換して $parsedDate へ。失敗したら $parseSuccess が $false。
    [bool]$isLeaveSuccess = [DateTime]::TryParseExact(($baseDate.toString("yyyy/MM/dd")+" "+$leaveTimeStr),
        "yyyy/MM/dd HH:mm", 
        [Globalization.DateTimeFormatInfo]::CurrentInfo,
        [Globalization.DateTimeStyles]::AllowWhiteSpaces, 
        [ref]$leaveTime)
}while(!$isLeaveSuccess)

do{
    $restTimeStr = Read-Host "退勤時刻を入力（デフォルト：`"$defaultRestTimeStr`"）"
    # 空白時にデフォルト値を設定する
    if(!$restTimeStr){
        $restTimeStr = $defaultRestTimeStr
    }
    # 適当な値で初期化
    [datetime]$restTime = [datetime]::MinValue
    # 変換して $parsedDate へ。失敗したら $parseSuccess が $false。
    [bool]$isRestSuccess = [DateTime]::TryParseExact(($baseDate.toString("yyyy/MM/dd")+" "+$restTimeStr),
        "yyyy/MM/dd HH:mm", 
        [Globalization.DateTimeFormatInfo]::CurrentInfo,
        [Globalization.DateTimeStyles]::AllowWhiteSpaces, 
        [ref]$restTime)
}while(!$isRestSuccess)

Write-Output ("出勤：" + $arriveTime.ToString("yyyy/MM/dd HH:mm"))
Write-Output ("退勤：" + $leaveTime.ToString("yyyy/MM/dd HH:mm"))
Write-Output ("休憩：" + $restTime.ToString("yyyy/MM/dd HH:mm"))

#このままExcelに代入するとずれるので補正する。
$arriveTime = $arriveTime.AddDays($baseOffsetDays)
$leaveTime = $leaveTime.AddDays($baseOffsetDays)
$restTime = $restTime.AddDays($baseOffsetDays)

try{
    # Excelオブジェクト作成
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $book = $excel.Workbooks.Open((Join-Path $excelDir $excelFileName))
    $sheet = $excel.Worksheets.Item($excelSheetName)

    # 現在日付までoffset
    $rowOffset = ($currentDate - $startDate).Days
    
    # 出社時刻・退勤時刻のセル取得
    $arriveTimeCell = $sheet.Cells.Item($startRow+$rowOffset, $startCol)
    $leaveTimeCell = $sheet.Cells.Item($startRow+$rowOffset, $startCol+1)
    $restTimeCell = $sheet.Cells.Item($startRow+$rowOffset, $startCol+2)
    
    # セルに値を代入
    $arriveTimeCell.Value = $arriveTime
    $leaveTimeCell.Value = $leaveTime
    $restTimeCell.Value = $restTime

    # 上書き保存 閉じる
    $book.Save()
    $excel.Quit()

} finally {

    # null破棄
    $excel,$book,$sheet,$arriveTimeCell,$leaveTimeCell,$restTimeCell | foreach{$_ = $null}

}