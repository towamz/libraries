<#
.SYNOPSIS
    テキストファイルに書いたフォルダ構成を、選択した親フォルダの下に一括作成するGUIツール。

.DESCRIPTION
    - 親フォルダはGUIで選択
    - 作成するフォルダは、スクリプトと同じフォルダの newfolderlist.txt に
      「1行1フォルダ(親フォルダからの相対パス)」で書いておく(パスは下の定数で変更可能)
      空行と、行頭が # の行は無視されます
    - 既存フォルダがある場合の挙動をGUIで選択
        * 中断する(何も作成しない)
        * ないものだけ作成する
    - 隠しメニュー: Ctrl + Shift + A で「中身を削除して作り直す」が表示されます

.NOTES
    実行: powershell -ExecutionPolicy Bypass -File .\FolderMaker.ps1
    このファイルは UTF-8 (BOM付き) で保存してください(Windows PowerShell 5.1 で日本語が化けないため)。
#>

#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
# 設定(定数)
# ---------------------------------------------------------------
# フォルダ一覧ファイルのパス: 既定はスクリプトと同じフォルダの newfolderlist.txt
$ScriptDir         = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$DEFAULT_LIST_PATH = Join-Path $ScriptDir 'newfolderlist.txt'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------
# 起動時チェック: フォルダ一覧ファイル
#   - 3バイト以下 : 中身を確認し、未入力(0バイト・BOMのみ・空白のみ)なら終了
#   - 4バイト以上 : 入力ありとみなして通過(検証は実行時)
#   - それ以外  : 中身はここでは検証せず通過させる(検証は実行時)
# ---------------------------------------------------------------
if (-not (Test-Path -LiteralPath $DEFAULT_LIST_PATH -PathType Leaf)) {
    try {
        New-Item -ItemType File -Path $DEFAULT_LIST_PATH | Out-Null
        $startupMsg  = "フォルダ一覧を作成して再実行してください。`nフォルダ一覧の空ファイルを作成しました。"
        $startupIcon = 'Information'
    }
    catch {
        $startupMsg  = "フォルダ一覧の空ファイルを作成できませんでした。`n$($_.Exception.Message)"
        $startupIcon = 'Error'
    }
    [void][System.Windows.Forms.MessageBox]::Show($startupMsg, 'FolderMaker', 'OK', $startupIcon)
    exit
}
if ((Get-Item -LiteralPath $DEFAULT_LIST_PATH).Length -le 3) {
    $listText = [System.IO.File]::ReadAllText($DEFAULT_LIST_PATH, [System.Text.Encoding]::UTF8)
    if ([string]::IsNullOrWhiteSpace($listText.Trim([char]0xFEFF))) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "フォルダ一覧に入力がありません。`n`n$DEFAULT_LIST_PATH", 'FolderMaker', 'OK', 'Information')
        exit
    }
}

# ---------------------------------------------------------------
# ロジック
# ---------------------------------------------------------------

function Read-FolderList {
    param([string]$Path)

    $folders = New-Object System.Collections.Generic.List[string]
    $errors  = New-Object System.Collections.Generic.List[string]
    $seen    = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

    $lineNo = 0
    foreach ($raw in @(Get-Content -LiteralPath $Path -Encoding UTF8)) {
        $lineNo++
        $line = $raw.Trim()
        if ($line -eq '' -or $line.StartsWith('#')) { continue }

        $line = ($line -replace '/', '\')

        # 絶対パス・ドライブ指定・UNCは不可(親フォルダの外に出さない)
        if ([System.IO.Path]::IsPathRooted($line) -or $line -match '^[A-Za-z]:') {
            $errors.Add("行$($lineNo): 相対パスで書いてください ($raw)")
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
            continue
        }

        if ($seen.Add($line)) { $folders.Add($line) }
    }

    [pscustomobject]@{
        Folders = $folders.ToArray()
        Errors  = $errors.ToArray()
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
# GUI
# ---------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text            = 'フォルダ一括作成'
$form.ClientSize      = New-Object System.Drawing.Size(560, 394)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.KeyPreview      = $true
$form.Font            = New-Object System.Drawing.Font('Yu Gothic UI', 9)

# 親フォルダ
$lblParent = New-Object System.Windows.Forms.Label
$lblParent.Text = '親フォルダ'
$lblParent.Location = New-Object System.Drawing.Point(15, 20)
$lblParent.AutoSize = $true

$txtParent = New-Object System.Windows.Forms.TextBox
$txtParent.Location = New-Object System.Drawing.Point(110, 17)
$txtParent.Size = New-Object System.Drawing.Size(350, 24)

$btnParent = New-Object System.Windows.Forms.Button
$btnParent.Text = '参照...'
$btnParent.Location = New-Object System.Drawing.Point(470, 15)
$btnParent.Size = New-Object System.Drawing.Size(75, 27)

# 既存フォルダがある場合
$gb = New-Object System.Windows.Forms.GroupBox
$gb.Text = '既存のフォルダがある場合'
$gb.Location = New-Object System.Drawing.Point(15, 56)
$gb.Size = New-Object System.Drawing.Size(530, 108)

$rbAbort = New-Object System.Windows.Forms.RadioButton
$rbAbort.Text = '中断する(何も作成しない)'
$rbAbort.Location = New-Object System.Drawing.Point(15, 24)
$rbAbort.AutoSize = $true
$rbAbort.Checked = $true

$rbSkip = New-Object System.Windows.Forms.RadioButton
$rbSkip.Text = 'ないものだけ作成する'
$rbSkip.Location = New-Object System.Drawing.Point(15, 50)
$rbSkip.AutoSize = $true

# 隠しメニュー(Ctrl+Shift+A で表示)
$rbRecreate = New-Object System.Windows.Forms.RadioButton
$rbRecreate.Text = '[上級者] 既存フォルダの中身を削除して作り直す'
$rbRecreate.Location = New-Object System.Drawing.Point(15, 76)
$rbRecreate.AutoSize = $true
$rbRecreate.ForeColor = [System.Drawing.Color]::Firebrick
$rbRecreate.Visible = $false

$gb.Controls.AddRange(@($rbAbort, $rbSkip, $rbRecreate))

$chkDry = New-Object System.Windows.Forms.CheckBox
$chkDry.Text = 'プレビューのみ(実際には作成しない)'
$chkDry.Location = New-Object System.Drawing.Point(20, 177)
$chkDry.AutoSize = $true

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = '実行'
$btnRun.Location = New-Object System.Drawing.Point(440, 171)
$btnRun.Size = New-Object System.Drawing.Size(105, 32)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 214)
$txtLog.Size = New-Object System.Drawing.Size(530, 165)
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'

$form.Controls.AddRange(@(
    $lblParent, $txtParent, $btnParent,
    $gb, $chkDry, $btnRun, $txtLog
))

function Write-Log {
    param([string]$Message)
    $txtLog.AppendText($Message + "`r`n")
}

# --- イベント ---

$btnParent.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'フォルダを作成する親フォルダを選択してください'
    if ($dlg.ShowDialog() -eq 'OK') { $txtParent.Text = $dlg.SelectedPath }
})

# 隠しメニューの切り替え: Ctrl + Shift + A
$form.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.Shift -and $e.KeyCode -eq 'A') {
        $rbRecreate.Visible = -not $rbRecreate.Visible
        if (-not $rbRecreate.Visible -and $rbRecreate.Checked) { $rbAbort.Checked = $true }
        $e.SuppressKeyPress = $true
    }
})

$btnRun.Add_Click({
    $txtLog.Clear()
    $parent   = $txtParent.Text.Trim()
    $listPath = $DEFAULT_LIST_PATH

    if ([string]::IsNullOrWhiteSpace($parent) -or -not (Test-Path -LiteralPath $parent -PathType Container)) {
        Write-Log '親フォルダが見つかりません。'
        return
    }
    if (-not (Test-Path -LiteralPath $listPath -PathType Leaf)) {
        Write-Log "フォルダ一覧ファイルが見つかりません: $listPath"
        return
    }
    Write-Log "一覧ファイル: $listPath"

    try {
        $parsed = Read-FolderList -Path $listPath
        if ($parsed.Errors.Count -gt 0) {
            Write-Log '一覧ファイルに問題があります。修正してから再実行してください。'
            foreach ($err in $parsed.Errors) { Write-Log "  $err" }
            return
        }
        if ($parsed.Folders.Count -eq 0) {
            Write-Log '作成するフォルダが書かれていません。'
            return
        }

        $mode = 'Abort'
        if ($rbSkip.Checked) { $mode = 'Skip' }
        if ($rbRecreate.Visible -and $rbRecreate.Checked) { $mode = 'Recreate' }

        if ($mode -eq 'Recreate' -and -not $chkDry.Checked) {
            $answer = [System.Windows.Forms.MessageBox]::Show(
                "一覧にある既存フォルダの中身を完全に削除します(ごみ箱には入りません)。`n続行しますか?",
                '確認', 'YesNo', 'Warning', 'Button2')
            if ($answer -ne 'Yes') {
                Write-Log 'キャンセルしました。'
                return
            }
        }

        Invoke-FolderCreate -Parent $parent -Folders $parsed.Folders -Mode $mode -DryRun $chkDry.Checked
    }
    catch {
        Write-Log "エラー: $($_.Exception.Message)"
    }
})

[void]$form.ShowDialog()
