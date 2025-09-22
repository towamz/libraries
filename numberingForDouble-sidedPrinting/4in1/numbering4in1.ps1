#  powershell -ExecutionPolicy Bypass -File .\numbering4in1.ps1

# パラメータ読み込み
. "$PSScriptRoot\numbering4in1param.ps1"

# カレントディレクトリをスクリプトファイルのディレクトリに変更する
$targetPath = Join-Path $PSScriptRoot $subFoldderName
Set-Location -Path $targetPath
Write-Host $(Get-Location)

function Create-Blankfile {
    param (
        [int]$pageNumber
    )

    if (Test-Path $blankPdfFullname) {
        $blankPdf = Get-Item $blankPdfFullname
        # 新しいファイル名を作成
        $newName = "$pageNumber-$($blankPdf.BaseName)$(($blankPdf.Extension))"
        # 新しいフルパスを作成
        $newPath = Join-Path $targetPath $newName
        Copy-Item -Path $blankPdf -Destination $newPath
        Write-Host "ファイル生成:${newName}"
    } else {  
        # 新しいファイル名を作成
        $newName = "$pageNumber-blank.txt"
        # 新しいフルパスを作成
        $newPath = Join-Path $targetPath $newName
        Out-File -FilePath $newPath -Encoding UTF8
        Write-Host "ファイル生成:${newName}"
    }
}


# 最初のループ：すべてのファイル名が末尾に数値があるか確認
Get-ChildItem -Filter *.pdf -File | ForEach-Object {
    # 数値が末尾にない場合、即座にエラーメッセージを表示して処理を中止
    if ($_.BaseName -notmatch "-(\d+)$") {
        throw "無効なファイル名: $($_.Name)。数値が末尾にないため処理を中止します。"
    }else{
        Write-Host $_.Name
    }
}

$conf = Read-Host "ファイル名変更を実施しますか? (yesで実施)"

if ($conf -ne "yes") {
    throw "処理を中止します。"
}

## 2番目のループ：ファイル名をマッピングに従って変更
# 最大値を保持する変数
$maxNumber = 0
Get-ChildItem -Filter *.pdf -File | ForEach-Object {
    if ($_.BaseName -match "-(\d+)$") {
        $number = [int]$matches[1]

        # ページマッピングを取得する
        $moduloResult = $number % $pagesCountInPaper
        $newNumber = $moduloMapping[$moduloResult] + $number

        # 最大値の更新
        if ($newNumber -gt $maxNumber) {
            $maxNumber = $newNumber
        }

        # 新しいファイル名を作成
        $newName = "$newNumber-$($_.BaseName)$(($_.Extension))"

        # 新しいフルパスを作成
        $newPath = Join-Path $targetPath $newName

        # ファイル名を変更
        Rename-Item $_.FullName -NewName $newPath
        Write-Host $newName
    }
}


## 欠番に空ファイルを生成する
# 最大ページ数を取得
$pagesCount = [math]::Ceiling($maxNumber / $pagesCountInPaper) * $pagesCountInPaper
Write-Host "最大ページ数:${pagesCount},実ページ数:${maxNumber}"

$cnt = 0
Get-ChildItem -Filter *.pdf -File | 
# 数値順に並び替える(実施しないと文字列ソートになる<例>1,11,2,21,3)
Sort-Object { 
        if ($_.BaseName -match '^(\d+)-') { 
            [int]$matches[1] 
        } else { 
            [int]::MaxValue   
        }
} | 
ForEach-Object {
    if ($_.BaseName -match "^(\d+)-") {
        $newNumber = [int]$matches[1]

        $cnt++
        while ($cnt -lt $newNumber) {
            Create-Blankfile $cnt
            $cnt++
        }
        Write-Host $newNumber
    }
}

# ループを抜けた後、最大ページ数までblankfileを生成する
$cnt++
while ($cnt -le $pagesCount) {
    Create-Blankfile $cnt
    $cnt++
}
