<#
.SYNOPSIS
    設定ファイルに書いたフォルダ構成を、起点フォルダの下に一括作成するコンソール版ツール。

.DESCRIPTION
    - 設定ファイル: スクリプトと同じフォルダの newfolderlist.txt(パスは下の定数で変更可能)
        1行目     : 起点フォルダ(絶対パス。例: D:\Work\Project)
        2行目以降 : 追加するフォルダ名(起点フォルダからの相対パス、1行1フォルダ)
                    空行と、行頭が # の行は無視されます
    - 実行するとすぐプレビューを表示する([OK] / [重複] / [パス不正])
        すべてOK : [1] 実行する / [9] 中止する
        重複あり : [2] ないものだけ実行する / [9] 中止する
        パス不正 : 続行できません(フォルダ名を見直してください)
    - 隠しメニュー: 重複ありのとき Recreate と入力すると
      「既存フォルダの中身を削除して作り直す」(選択肢には表示されない)

.NOTES
    実行: powershell -ExecutionPolicy Bypass -File .\FolderMakerCui.ps1
    このファイルは UTF-8 (BOM付き) で保存してください(Windows PowerShell 5.1 で日本語が化けないため)。
#>

#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
# 設定(定数)
# ---------------------------------------------------------------
# 設定ファイルのパス: 既定はスクリプトと同じフォルダの newfolderlist.txt
$ScriptDir         = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$DEFAULT_LIST_PATH = Join-Path $ScriptDir 'newfolderlist.txt'

# ---------------------------------------------------------------
# ロジック
# ---------------------------------------------------------------

# 設定ファイルを読み、起点フォルダ(1行目)と追加フォルダ(2行目以降)に分けて検証する。
function Read-FolderList {
    param([string]$Path)

    $folders = New-Object System.Collections.Generic.List[string]
    $errors  = New-Object System.Collections.Generic.List[string]
    $entries = New-Object System.Collections.Generic.List[object]
    $seen    = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

    $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8)

    # 1行目: 起点フォルダ(絶対パス)
    $parent      = $null
    $parentError = $null
    $first = if ($lines.Count -gt 0) { ([string]$lines[0]).Trim().Trim('"') } else { '' }
    if ($first -eq '' -or $first -notmatch '^([A-Za-z]:[\\/]|\\\\)') {
        $parentError = '1行目に起点フォルダを絶対パスで書いてください(例: D:\Work)'
    }
    else {
        $parent = $first
    }

    # 2行目以降: 追加するフォルダ名
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $lineNo = $i + 1
        $raw    = [string]$lines[$i]
        $line   = $raw.Trim()
        if ($line -eq '' -or $line.StartsWith('#')) { continue }

        $text = $line                      # 表示用(書かれたまま)
        $line = ($line -replace '/', '\')

        # 絶対パス・ドライブ指定・UNCは不可(起点フォルダの外に出さない)
        if ([System.IO.Path]::IsPathRooted($line) -or $line -match '^[A-Za-z]:') {
            $errors.Add("行$($lineNo): 相対パスで書いてください ($raw)")
            $entries.Add([pscustomobject]@{ Text = $text; Valid = $false })
            continue
        }
        $line = $line.TrimEnd('\')

        $bad = $false
        foreach ($part in ($line -split '\\')) {
            if ($part -eq '' -or $part -eq '.' -or $part -eq '..' -or
                $part -match '[<>:"|?*\x00-\x1f]' -or
                $part -match '[. ]$' -or
                $part -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') {
                $bad = $true
                break
            }
        }
        if ($bad) {
            $errors.Add("行$($lineNo): 使えない名前が含まれています ($raw)")
            $entries.Add([pscustomobject]@{ Text = $text; Valid = $false })
            continue
        }

        if ($seen.Add($line)) {
            $folders.Add($line)
            $entries.Add([pscustomobject]@{ Text = $line; Valid = $true })
        }
    }

    [pscustomobject]@{
        Parent      = $parent
        ParentError = $parentError
        Folders     = $folders.ToArray()
        Errors      = $errors.ToArray()
        Entries     = $entries.ToArray()
    }
}

function Test-UnderAny {
    param([string]$Path, [string[]]$Roots)
    foreach ($r in $Roots) {
        if ($Path.StartsWith($r + '\', [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Invoke-FolderCreate {
    param(
        [string]$Parent,
        [string[]]$Folders,
        [ValidateSet('Abort', 'Skip', 'Recreate')][string]$Mode,
        [bool]$DryRun
    )

    # 浅い階層から順に処理する(親 -> 子)
    $sorted = @($Folders | Sort-Object { ($_ -split '\\').Count }, { $_ })

    # 作成前の存在状況をスナップショットとして取る
    $existing = @($sorted | Where-Object { Test-Path -LiteralPath (Join-Path $Parent $_) })

    if ($Mode -eq 'Abort' -and $existing.Count -gt 0) {
        Write-Log '既に存在するフォルダがあるため中断しました。何も作成していません。'
        foreach ($e in $existing) { Write-Log "  存在: $e" }
        return
    }

    $created = 0
    $skipped = 0
    $deleted = @()

    foreach ($f in $sorted) {
        $full = Join-Path $Parent $f

        if (($existing -contains $f) -and -not (Test-UnderAny $f $deleted)) {
            if ($Mode -eq 'Recreate') {
                Write-Log "[削除して再作成] $f"
                if (-not $DryRun) {
                    Remove-Item -LiteralPath $full -Recurse -Force
                    New-Item -ItemType Directory -Path $full | Out-Null
                }
                $deleted += $f
                $created++
            }
            else {
                Write-Log "[スキップ] $f"
                $skipped++
            }
        }
        else {
            Write-Log "[作成] $f"
            if (-not $DryRun) {
                New-Item -ItemType Directory -Path $full -Force | Out-Null
            }
            $created++
        }
    }

    $label = if ($DryRun) { '(プレビュー)' } else { '' }
    Write-Log ''
    Write-Log "完了$label : 作成 $created 件 / スキップ $skipped 件"
}

# ---------------------------------------------------------------
# コンソール用の入出力
# ---------------------------------------------------------------

# Invoke-FolderCreate が使う出力先(GUI版ではログ欄、ここではコンソール)
function Write-Log {
    param([string]$Message)
    Write-Host $Message
}

# 起動時チェック。続行してよければ $true、終了すべきなら $false を返す。
# (一覧ファイルがなければ空ファイルを作成する副作用あり)
function Invoke-StartupCheck {
    if (-not (Test-Path -LiteralPath $DEFAULT_LIST_PATH -PathType Leaf)) {
        try {
            New-Item -ItemType File -Path $DEFAULT_LIST_PATH | Out-Null
            Write-Host 'フォルダ一覧を作成して再実行してください。'
            Write-Host 'フォルダ一覧の空ファイルを作成しました。'
        }
        catch {
            Write-Host 'フォルダ一覧の空ファイルを作成できませんでした。' -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
        Write-Host $DEFAULT_LIST_PATH
        return $false
    }

    # 3バイト以下(0バイト・BOMのみ・空白のみ)のときだけ中身を確認する簡易チェック
    if ((Get-Item -LiteralPath $DEFAULT_LIST_PATH).Length -le 3) {
        $listText = [System.IO.File]::ReadAllText($DEFAULT_LIST_PATH, [System.Text.Encoding]::UTF8)
        if ([string]::IsNullOrWhiteSpace($listText.Trim([char]0xFEFF))) {
            Write-Host 'フォルダ一覧に入力がありません。'
            Write-Host $DEFAULT_LIST_PATH
            return $false
        }
    }
    return $true
}

# 一覧の各行の状態を表示し、不正の件数と既存の件数を返す。
function Show-FolderPreview {
    param([string]$Parent, [object[]]$Entries)

    $invalidCount = 0
    $existsCount  = 0
    foreach ($entry in $Entries) {
        if (-not $entry.Valid) {
            Write-Host "[パス不正] $($entry.Text)" -ForegroundColor Red
            $invalidCount++
        }
        elseif (Test-Path -LiteralPath (Join-Path $Parent $entry.Text)) {
            Write-Host "[重複] $($entry.Text)" -ForegroundColor Yellow
            $existsCount++
        }
        else {
            Write-Host "[OK] $($entry.Text)" -ForegroundColor Green
        }
    }
    [pscustomobject]@{ Invalid = $invalidCount; Exists = $existsCount }
}

# プレビュー結果に応じた選択肢を出す。中止なら $null、実行なら挙動(Abort/Skip/Recreate)を返す。
function Read-RunChoice {
    param([bool]$HasExisting)

    if ($HasExisting) { Write-Host '[2] ないものだけ実行する / [9] 中止する' }
    else              { Write-Host '[1] 実行する / [9] 中止する' }

    while ($true) {
        $answer = (Read-Host '番号を入力').Trim()
        if ($answer -eq '9')                          { return $null }
        if (-not $HasExisting -and $answer -eq '1')   { return 'Abort' }
        if ($HasExisting -and $answer -eq '2')        { return 'Skip' }
        if ($HasExisting -and $answer -eq 'Recreate') { return 'Recreate' }  # 隠しメニュー(表示しない)
        Write-Host '表示されている番号を入力してください。' -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------
# メイン
# ---------------------------------------------------------------
try {
    if (-not (Invoke-StartupCheck)) {
        [void](Read-Host 'Enterキーで終了します')
        exit 1
    }

    $parsed = Read-FolderList -Path $DEFAULT_LIST_PATH

    if ($parsed.ParentError) {
        Write-Host $parsed.ParentError -ForegroundColor Yellow
        [void](Read-Host 'Enterキーで終了します')
        exit 1
    }
    $parent = $parsed.Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        Write-Host "起点フォルダが見つかりません: $parent" -ForegroundColor Yellow
        [void](Read-Host 'Enterキーで終了します')
        exit 1
    }
    if ($parsed.Entries.Count -eq 0) {
        Write-Host '作成するフォルダが書かれていません。' -ForegroundColor Yellow
        [void](Read-Host 'Enterキーで終了します')
        exit 1
    }

    Write-Host "一覧ファイル: $DEFAULT_LIST_PATH"
    Write-Host "起点フォルダ: $parent"
    Write-Host "作成対象: $($parsed.Entries.Count) 件"
    Write-Host ''

    # プレビュー
    $report = Show-FolderPreview -Parent $parent -Entries $parsed.Entries
    Write-Host ''

    if ($report.Invalid -gt 0) {
        Write-Host '続行できません。フォルダ名を見直してください。' -ForegroundColor Red
        [void](Read-Host 'Enterキーで終了します')
        exit 1
    }

    $mode = Read-RunChoice -HasExisting ($report.Exists -gt 0)
    if (-not $mode) {
        Write-Host '中止しました。'
        [void](Read-Host 'Enterキーで終了します')
        exit 0
    }

    if ($mode -eq 'Recreate') {
        Write-Host '一覧にある既存フォルダの中身を完全に削除します(ごみ箱には入りません)。' -ForegroundColor Red
        if ((Read-Host '続行するには yes と入力') -ne 'yes') {
            Write-Host 'キャンセルしました。'
            [void](Read-Host 'Enterキーで終了します')
            exit 0
        }
    }

    Write-Host ''
    Invoke-FolderCreate -Parent $parent -Folders $parsed.Folders -Mode $mode -DryRun $false
    [void](Read-Host 'Enterキーで終了します')
    exit 0
}
catch {
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
    [void](Read-Host 'Enterキーで終了します')
    exit 1
}