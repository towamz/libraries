Write-Output "ここに複数行を貼り付けてください。"

$lines = @()
$i=0
while ($true) {
    $line = Read-Host
    if ($line -eq 'EOF') { break }
    # if ([string]::IsNullOrWhiteSpace($line)) { break }
    if ([string]::IsNullOrWhiteSpace($line)) { 
        $i++ 
    }else{
        $i=0
    }
    if ($i -ge 3) { break }
    $lines += $line
}

$timeLines = @()
$i=0
foreach ($line in $lines) {
    if ($line -match '^\d{1,2}:\d{2}(:\d{2})?$') {
        $time = $line -replace ':', ''

        $nextLine = $lines[$i + 1]
        $station = ""
        if ($nextLine -match '発\s*(.+?)時刻表') {
            $station = $matches[1]
        }elseif ($nextLine -match "着\s*(.*?)(?=地図)") {
            $station = $matches[1]
        }
        $station = $station -replace "\(.*$", ""
        $station = $station.Trim()
        $timeLines += "$time$station"
    }elseif ($Line -match "(\d{2}):(\d{2})着(\d{2}):(\d{2})発\s+(.*?)(?=時刻表)") {
        $h1 = [int]$matches[1]
        $m1 = [int]$matches[2]
        $h2 = [int]$matches[3]
        $m2 = [int]$matches[4]
        $station = $matches[5]
        $station = $station -replace "\(.*$", ""
        $station = $station.Trim()

        # 時刻部分の処理
        $time = "{0:D2}{1:D2}" -f $h1, $m1

        if (
            ($h1 -eq $h2) -or
            ($h2 -eq ($h1 + 1) -and $m2 -lt $m1)
        ) {
            $time += ("{0:D2}" -f $m2)  # 発のhh省略
        } else {
            $time += ("{0:D2}{1:D2}" -f $h2, $m2)
        }
        $timeLines += "$time$station"
    }
    $i++
}

Write-Output "`n--- 整形済み運行時刻 ---"
$timeLines
