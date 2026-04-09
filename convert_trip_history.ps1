param(
  [Parameter(Mandatory = $true)]
  [string]$InputPath,

  [Parameter(Mandatory = $false)]
  [string]$OutputPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-FullPath {
  param([string]$Path)

  $resolved = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
  if ($resolved) {
    return $resolved.Path
  }

  return [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location) -ChildPath $Path))
}

function Convert-ToDate {
  param([object]$Value)

  if ($null -eq $Value) { return $null }

  if ($Value -is [datetime]) {
    return [datetime]$Value
  }

  if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) {
    try {
      return [datetime]::FromOADate([double]$Value)
    } catch {
      return $null
    }
  }

  $text = "$Value".Trim()
  if ([string]::IsNullOrWhiteSpace($text)) {
    return $null
  }

  try {
    return [datetime]::Parse($text)
  } catch {
    return $null
  }
}

function Get-HeaderIndex {
  param(
    [hashtable]$Map,
    [string]$Header
  )

  if (-not $Map.ContainsKey($Header)) {
    throw "입력 파일에 '$Header' 열이 없습니다."
  }

  return [int]$Map[$Header]
}

if (-not (Test-Path -LiteralPath $InputPath)) {
  throw "입력 파일을 찾을 수 없습니다: $InputPath"
}

$inputFullPath = Resolve-FullPath -Path $InputPath
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
  $baseDir = Split-Path -Parent $inputFullPath
  $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputFullPath)
  $OutputPath = Join-Path $baseDir ("{0}_정산변환.xlsx" -f $baseName)
}

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)

$excel = $null
$inWb = $null
$outWb = $null
$inWs = $null
$outWs = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $inWb = $excel.Workbooks.Open($inputFullPath)
  $inWs = $inWb.Worksheets.Item(1)

  $used = $inWs.UsedRange
  $rowCount = [int]$used.Rows.Count
  $colCount = [int]$used.Columns.Count

  if ($rowCount -lt 2) {
    throw "입력 파일에 데이터 행이 없습니다."
  }

  $headers = @{}
  for ($c = 1; $c -le $colCount; $c++) {
    $name = "$($inWs.Cells.Item(1, $c).Text)".Trim()
    if (-not [string]::IsNullOrWhiteSpace($name)) {
      $headers[$name] = $c
    }
  }

  $idxStart = Get-HeaderIndex -Map $headers -Header "시작일"
  $idxEnd = Get-HeaderIndex -Map $headers -Header "종료일"
  $idxDest = Get-HeaderIndex -Map $headers -Header "출장지"
  $idxTraveler = Get-HeaderIndex -Map $headers -Header "출장자"
  $idxVehicle = Get-HeaderIndex -Map $headers -Header "공무용차량여부"

  $records = New-Object System.Collections.Generic.List[object]

  for ($r = 2; $r -le $rowCount; $r++) {
    $traveler = "$($inWs.Cells.Item($r, $idxTraveler).Text)".Trim()
    if ([string]::IsNullOrWhiteSpace($traveler)) {
      continue
    }

    $startRaw = $inWs.Cells.Item($r, $idxStart).Value2
    $endRaw = $inWs.Cells.Item($r, $idxEnd).Value2

    $startDate = Convert-ToDate -Value $startRaw
    $endDate = Convert-ToDate -Value $endRaw

    if ($null -eq $startDate) {
      continue
    }

    if ($null -eq $endDate) {
      $endDate = $startDate
    }

    if ($endDate -lt $startDate) {
      $tmp = $startDate
      $startDate = $endDate
      $endDate = $tmp
    }

    $tripDays = [int](($endDate.Date - $startDate.Date).TotalDays + 1)
    if ($tripDays -lt 1) { $tripDays = 1 }

    $dest = "$($inWs.Cells.Item($r, $idxDest).Text)".Trim()
    $vehicleUse = "$($inWs.Cells.Item($r, $idxVehicle).Text)".Trim()
    $transport = if ($vehicleUse -eq "이용") { "관용차량" } else { "" }

    $mealUnit = 25000
    $dailyUnit = 25000
    $mealAmt = $mealUnit * $tripDays
    $dailyAmt = $dailyUnit * $tripDays

    $record = [PSCustomObject]@{
      Traveler = $traveler
      TripDate = $startDate
      Origin = "세종"
      Destination = $dest
      Transport = $transport
      TransportFare = 0
      TripDays = $tripDays
      MealUnit = $mealUnit
      MealAmt = $mealAmt
      DailyUnit = $dailyUnit
      DailyAmt = $dailyAmt
      LodgingNights = 0
      LodgingCost = 0
    }

    [void]$records.Add($record)
  }

  $sorted = $records |
    Sort-Object -Property @{Expression = { $_.Traveler }; Ascending = $true}, @{Expression = { $_.TripDate }; Ascending = $true}

  $outWb = $excel.Workbooks.Add()
  $outWs = $outWb.Worksheets.Item(1)
  $outWs.Name = "출장정산"

  $outHeaders = @(
    "출장자", "출장일", "출발지", "도착지", "교통수단", "교통요금",
    "출장일수", "식비단가", "식비금액", "일비단가", "일비금액",
    "숙박일수", "숙박비", "청구액", "지급액"
  )

  for ($i = 0; $i -lt $outHeaders.Count; $i++) {
    $outWs.Cells.Item(1, $i + 1).Value2 = $outHeaders[$i]
  }

  $row = 2
  foreach ($rec in $sorted) {
    $outWs.Cells.Item($row, 1).Value2 = $rec.Traveler
    $outWs.Cells.Item($row, 2).Value2 = $rec.TripDate
    $outWs.Cells.Item($row, 2).NumberFormat = "yyyy-mm-dd"
    $outWs.Cells.Item($row, 3).Value2 = $rec.Origin
    $outWs.Cells.Item($row, 4).Value2 = $rec.Destination
    $outWs.Cells.Item($row, 5).Value2 = $rec.Transport
    $outWs.Cells.Item($row, 6).Value2 = $rec.TransportFare
    $outWs.Cells.Item($row, 7).Value2 = $rec.TripDays
    $outWs.Cells.Item($row, 8).Value2 = $rec.MealUnit
    $outWs.Cells.Item($row, 9).Value2 = $rec.MealAmt
    $outWs.Cells.Item($row, 10).Value2 = $rec.DailyUnit
    $outWs.Cells.Item($row, 11).Value2 = $rec.DailyAmt
    $outWs.Cells.Item($row, 12).Value2 = $rec.LodgingNights
    $outWs.Cells.Item($row, 13).Value2 = $rec.LodgingCost
    $outWs.Cells.Item($row, 14).FormulaR1C1 = "=RC[-8]+RC[-5]+RC[-3]+RC[-1]"
    $outWs.Cells.Item($row, 15).FormulaR1C1 = "=RC[-1]"

    $row++
  }

  if ($row -gt 2) {
    $moneyRange = $outWs.Range("F2:O$($row - 1)")
    $moneyRange.NumberFormat = "#,##0"
  }

  $headerRange = $outWs.Range("A1:O1")
  $headerRange.Font.Bold = $true
  $headerRange.Interior.Color = 15773696
  $outWs.Columns.AutoFit() | Out-Null

  $outWb.SaveAs($outputFullPath)

  Write-Output "변환 완료: $outputFullPath"
  Write-Output "처리 건수: $($sorted.Count)"
}
finally {
  if ($inWb) { $inWb.Close($false) }
  if ($outWb) { $outWb.Close($true) }

  if ($inWs) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($inWs) }
  if ($outWs) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outWs) }
  if ($inWb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($inWb) }
  if ($outWb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outWb) }

  if ($excel) {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
