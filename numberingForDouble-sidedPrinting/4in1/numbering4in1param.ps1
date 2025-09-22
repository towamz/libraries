#  powershell -ExecutionPolicy Bypass -File .\numbering4in1.ps1

# ページ数指定(4in1の両面=8ページ)
$pagesCountInPaper = 8
$subFoldderName = 'numbering'
$delimiter = '-'
# $blankPdfFullname = 'C:\_雛形\blank_tate.pdf'
$blankPdfFullname = 'C:\_雛形\blank_yoko.pdf'

# 剰余残の結果からページマッピングを定義
$moduloMapping = @{
    1 = 0
    2 = 4
    3 = -1
    4 = 1
    5 = -2
    6 = 2
    7 = -3
    0 = -1
}
