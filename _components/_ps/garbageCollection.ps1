function garbageCollection(){
    # COMオブジェクトのみを取得
    $variables = Get-Variable | Where-Object { 
        $_.Value -is [System.__ComObject] -and $_.Name -notmatch '^env:|^global:|^function:' 
    }

    Write-Host "リリース予定: $($var.Name)"
    foreach ($var in $variables) {
        Write-Host "$($var.Name)"
    }

    # ユーザーに確認のプロンプトを表示
    $confirmation = Read-Host "上記変数をリリースしますか？ (yes でリリース)"

    # 2回目のループ：リリース処理
    if ($confirmation -eq "yes") {
        foreach ($var in $variables) {
            Write-Host "リリース中: $($var.Name)"
            # COMオブジェクトの解放
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($var.Value) | Out-Null
            # 変数の参照をnullに設定
            $var.Value = $null
        }

        # ガーベジコレクションを強制実行
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        
        Write-Host "COMオブジェクトの解放が完了しました"
    }
}

