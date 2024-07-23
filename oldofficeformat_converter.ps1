$logPath = [Environment]::GetFolderPath('MyDocuments') #ログファイルパス
$script:pptFlag = $FALSE
function firstMenu {
    # 変換元フォルダを選択するダイアログの設定
    $Dialog = New-Object System.Windows.Forms.OpenFileDialog
    $Dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop') # デスクトップを初期フォルダに
    $Dialog.Title = "フォルダを選択してください｜XlsDocPptTo2007" # ウィンドウタイトル
    $Dialog.ValidateNames = 1 # 有効な Win32 ファイル名のみを受け入れるかどうか
    $Dialog.CheckFileExists = 0 # 存在しないファイル名をユーザーが指定した場合に、ファイルダイアログで警告を表示するかどうか
    $Dialog.CheckPathExists = 1 # ユーザーが無効なパスとファイル名を入力した場合に警告を表示するかどうか
    $Dialog.FileName = "フォルダを選択"
    if($Dialog.ShowDialog() -eq "OK") { # ダイアログにてOKボタンが押されたら
        $inputDir = Split-Path -Parent $Dialog.FileName
        $inputDir = $inputDir + "\" #対象フォルダパスを作成
    }
    else { # ダイアログにてキャンセルボタンが押されたら
        echo "ダイアログ操作をキャンセルしました";
        [System.Windows.MessageBox]::Show("ダイアログ操作をキャンセルしました","メッセージ","OK","Information")
        exit
    }
    return $inputDir
}
function Convert {
    $ErrorActionPreference = "Continue" # 例外が出ても続行
    Get-ChildItem -Path $script:srcDir -Include "*.xls", "*.doc", "*.ppt" -Recurse | ForEach-Object { # 指定フォルダ内にxls, doc, pptファイルが見つかったら
        $originalFile = $_
        # docの場合
        if ($originalFile.Extension -eq ".doc") {
            try {
                if ($script:hasWord) { # Wordがインストールされていれば
                    $originalFile.FullName
                    # Wordオブジェクト作成
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $FALSE
                    $word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone
                    $doc = $word.Documents.Open($originalFile.FullName)
                    if ($doc.HasVBProject) { #マクロを含んでいるか
                        $destFile = $originalFile.FullName.Replace("doc", "docm")
                        $fileFormat = 17 # WdSaveFormatMacroEnabled
                    } else {
                        $destFile = $originalFile.FullName.Replace("doc", "docx")
                        echo $destFile
                        $fileFormat = 16 # WdSaveFormat
                    }
                    # 変換元、変換先のフルパスをログ出力
                    Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換元ファイル:$($originalFile.FullName)"
                    Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換先ファイル:$destFile"
                    if (-not (Test-Path -Path $destFile)) { # 変換先に同じファイル名が存在しなければ
                        $doc.SaveAs2([ref]$destFile, [ref]$fileFormat) # 名前をつけて保存
                        # ドキュメントを閉じる
                        $doc.Close() | Wait-Job
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
                        $doc = $FALSE
                        # 変換元のdocファイルを削除
                        Remove-Item -Path $originalFile.FullName -Force
                        Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換元ファイル$($originalFile.FullName)を削除しました。"
                    } else {
                        Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換先ファイル$($destFile)はすでに存在するため変換しませんでした。"
                    }
                }
                else {
                    Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換元ファイル:$($originalFile.FullName)"
                    Write-Host "$($originalFile.FullName) が見つかりましたが、v2007以降のWordがインストールしていない為スキップします。"
                }
            }
            catch {
                Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換エラー:$($originalFile.FullName): $_"
            }
            finally {
                if ($doc) { #ドキュメントが閉じられてなければ
                    $doc.Close() | Wait-Job
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
                }
                if ($word) { #Wordが起動していれば
                    # Wordを終了
                    $word.quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                    Remove-Variable word
                }
            }
        }
        # xlsの場合
        elseif ($originalFile.Extension -eq ".xls") {
            try {
                if ($script:hasExcel) { # Excelがインストールされていれば 
                    # Excelオブジェクト作成
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $FALSE
                    $excel.DisplayAlerts = $FALSE
                    $book = $excel.Workbooks.Open($originalFile.FullName)
                    if ($book.HasVBProject) { #マクロを含んでいるか
                        $destFile = $originalFile.FullName.Replace("xls", "xlsm")
                        $fileFormat = 52  # xlOpenXMLWorkbookMacroEnabled
                    } else {
                        $destFile = $originalFile.FullName.Replace("xls", "xlsx")
                        $fileFormat = 51  # xlOpenXMLWorkbook
                    }
                    # 変換元、変換先のフルパスをログ出力
                    Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換元ファイル:$($originalFile.FullName)"
                    Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換先ファイル:$destFile"
                    if(-not (Test-Path -Path $destFile)) { # 変換先に同じファイル名が存在しなければ
                        $book.SaveAs($destFile, $fileFormat)# 名前をつけて保存
                        # ドキュメントを閉じる
                        $book.Close() | Wait-Job
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book) | Out-Null
                        $book = $FALSE
                        # 変換元のxlsファイルを削除
                        Remove-Item -Path $originalFile.FullName -Force
                        Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換元ファイル$($originalFile.FullName)を削除しました。"
                    } else {
                        Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換先ファイル$($destFile)はすでに存在するため変換しませんでした。"
                    }                    
                }
                else {
                    Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換元ファイル:$($originalFile.FullName)"
                    Write-Host "$($originalFile.FullName) が見つかりましたが、v2007以降のExcelがインストールしていない為スキップします。"                   
                }
            }
            catch {
                Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換エラー:$($originalFile.FullName): $_"
            }
            finally {
                if ($book) { # ブックが閉じられてなければ
                    $book.Close() | Wait-Job
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book) | Out-Null
                }
                if ($excel) { # Excelが起動していれば
                    # Excelを終了
                    $excel.quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }
            }
        }
        # pptの場合
        else {
            try {
                if ($script:hasPowerPoint) { # PowerPointがインストールされていれば
                    # PowerPointオブジェクト作成
                    $powerpoint = New-Object -ComObject PowerPoint.Application
                    $powerpoint.Visible = $FALSE
                    $powerpoint.DisplayAlerts = $FALSE
                    $slide = $powerpoint.Documents.Open($originalFile.FullName)
                    if ($slide.HasVBProject) { #マクロを含んでいるか
                        $destFile = $originalFile.FullName.Replace("ppt", "pptm")
                        $fileFormat = 17  # WdSaveFormatMacroEnabled
                    } else {
                        $destFile = $originalFile.FullName.Replace("ppt", "pptx")
                        $fileFormat = 16  # WdSaveFormat
                    }
                    # 変換元、変換先のフルパスをログ出力
                    Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換元ファイル:$($originalFile.FullName)"
                    Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換先ファイル:$destFile"
                    if(-not (Test-Path -Path $destFile)) { # 変換先に同じファイル名が存在しなければ
                        $slide.SaveAs2([ref]$destFile, [ref]$fileFormat) # 名前をつけて保存
                        # ドキュメントを閉じる
                        $slide.Close() | Wait-Job
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($slide) | Out-Null
                        $slide = $FALSE
                        # 変換元のdocファイルを削除
                        Remove-Item -Path $originalFile.FullName -Force
                        Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換元ファイル$($originalFile.FullName)を削除しました。"
                    } else {
                        Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換先ファイル$($destFile)はすでに存在するため変換しませんでした。"
                    }
                }
                else {
                    Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換元ファイル:$($originalFile.FullName)"
                    Write-Host "$($originalFile.FullName) が見つかりましたが、v2007以降のPowerPointがインストールしていない為スキップします。"                  
                }
            }
            catch {
                Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 変換エラー:$($originalFile.FullName): $_"
            }
            finally {
                if ($slide) { #ドキュメントが閉じられてなければ
                    $slide.Close() | Wait-Job
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($slide) | Out-Null
                }
                if ($powerpoint) { # PowerPointが起動していれば
                    # PowerPointを終了
                    $powerpoint.quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
                    Remove-Variable word
                }
            }

        }
    }
}
function chcekOffice {
    $officeDir = "C:\Program Files\Microsoft Office\root" # Officeのインストールフォルダ
    if(Test-Path $officeDir) { # Officeのインストールフォルダが存在したら
        $folderNames = Get-ChildItem -Path $officeDir -Directory
        foreach ($folderName in $folderNames) { # フォルダ一覧のループ
            $match = [System.Text.RegularExpressions.Regex]::Match($folderName.Name, 'Office(\d+)') # "Office00"に該当するフォルダが存在するかどうかの真偽値を代入
            if ($match.Groups[1].Value) {
                $version = $match.Groups[1].Value # 00の部分の数字だけを代入
            }
        }
        if ($version -ge 12) { # バージョン番号12(2007)以上であれば
            if (Test-Path $officeDir'\Office'$version"\WINWORD.EXE") { # Wordがインストールされているか
                $script:hasWord = $true
            }
            else {
                $script:hasWord = $false                    
            }
            if (Test-Path $officeDir'\Office'$version"\EXCEL.EXE") { # Excelがインストールされているか
                $script:hasExcel = $true
            }
            else {
                $script:hasExcel = $false                    
            }
            if (Test-Path $officeDir'\Office'$version"\POWERPNT.EXE") { # PowerPointがインストールされているか
                $script:hasPowerPoint = $true
            }
            else {
                $script:hasPowerPoint = $false                    
            }
        }
        else { # バージョン番号12(2007)以下であれば
            [System.Windows.Forms.MessageBox]::Show("v2007以降のOfficeがインストールされていません。OfficeをインストールするかインストールされたPCで実行してください。","メッセージ","OK","Warning")
            echo "ログ取得終了";
            Stop-Transcript | Out-Null # ログ取得終了
            exit
        }
    }
    else { # Office00フォルダが存在しなければ
        [System.Windows.Forms.MessageBox]::Show("v2007以降のOfficeがインストールされていません。OfficeをインストールするかインストールされたPCで実行してください。","メッセージ","OK","Warning")
        echo "ログ取得終了";
        Stop-Transcript | Out-Null #ログ取得終了
        exit
    }
}
#=== メイン処理 =========================================================
# ログ取得開始
$logfilePath = $logPath + "\xlsdocto2007_" + $((Get-Date).toString('yyyyMMddHHmmss')) + "_" + $env:COMPUTERNAME + ".log"
Start-Transcript $logfilePath -Force | Out-Null
# System.Windows.Formsアセンブリを有効化
Add-Type -Assembly System.Windows.Forms
$result1 = [System.Windows.Forms.MessageBox]::Show("指定のフォルダ内にある.xls, .doc, .pptをv2007形式(xlsx, xlsm, docx, docm, pptx, pptm)に一括変換します。`n`n【※変換前のデータは削除されますので必ず事前にバックアップしてください】`n`n続行する場合は「OK」, 中断する場合は「キャンセル」を押してください","確認","OKCancel","Warning","Button2")
if($result1 -eq "Cancel") {
    echo "ログ取得終了";
    Stop-Transcript | Out-Null #ログ取得終了
    exit
}
chcekOffice
$script:srcDir = firstMenu
$result3 = [System.Windows.Forms.MessageBox]::Show("一括変換処理を開始します`n`nフォルダ内のファイル数が多い場合や`nネットワークフォルダ(NASなど)の場合は`n数時間～数日かかる可能性があります`n`nよろしいですか?`n`n","確認","YesNo","Question","Button2")
# System.Windows.Formsアセンブリを有効化
if($result3 -eq "No") {
    echo "ログ取得終了";
    Stop-Transcript | Out-Null #ログ取得終了
    exit
}
Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 一括変換処理を開始しました"
[System.Windows.Forms.MessageBox]::Show(" $((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 一括変換処理を開始しました","メッセージ","OK","Information") 
Convert -inputPath $script:srcDir
if(-! $script:hasWord) {
    [System.Windows.Forms.MessageBox]::Show(" $((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) .docファイルが見つかりましたが、v2007以降のWordがインストールしていない為手動で変換してください。`n`nファイルの場所は $logfilePath に記載されています","メッセージ","OK","Warning")    
}
if(-! $script:hasExcel) {
    [System.Windows.Forms.MessageBox]::Show(" $((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) .xlsファイルが見つかりましたが、v2007以降のExcelがインストールしていない為手動で変換してください。`n`nファイルの場所は $logfilePath に記載されています","メッセージ","OK","Warning")    
}
if(-! $script:hasPowerPoint) {
    [System.Windows.Forms.MessageBox]::Show(" $((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) .pptファイルが見つかりましたが、v2007以降のPowerPointがインストールしていない為手動で変換してください。`n`nファイルの場所は $logfilePath に記載されています","メッセージ","OK","Warning")    
}
Write-Host "$((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 一括変換処理が終了しました"
[System.Windows.Forms.MessageBox]::Show(" $((Get-Date).toString('yyyy-MM-dd HH:mm:ss')) 一括変換処理が終了しました","メッセージ","OK","Information")    
echo "ログ取得終了";
Stop-Transcript | Out-Null # ログ取得終了
# ログファイルを変換元フォルダ直下に移動
Move-Item -Path $logfilePath -Destination $script:srcDir
exit
#=======================================================================
