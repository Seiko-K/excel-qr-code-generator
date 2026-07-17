Attribute VB_Name = "QRGenerator"
'=============================================================
' Excel QR Code Generator
' Version : v0.1
' Author  : Seiko Atsuumi
' License : MIT
'
' Purpose:
' Generates multiple QR codes from Excel data, exports them as
' PNG files, inserts them into the worksheet, and records a log.
'
' 目的:
' Excelデータから複数のQRコードを一括生成し、PNGファイルとして
' 保存・ワークシートへ貼り付け・処理履歴を記録します。
'
' Important:
' This module currently uses MSXML2.XMLHTTP and ADODB.Stream.
' These components are primarily intended for Windows Excel.
'
' 重要:
' 現在のモジュールは MSXML2.XMLHTTP と ADODB.Stream を使用します。
' これらは主にWindows版Excel向けのコンポーネントです。
'=============================================================

Option Explicit

'=============================================================
' Public Entry Points
' ボタン・外部呼び出し用エントリーポイント
'=============================================================

' Starts the batch QR code generation workflow.
' QRコードの一括生成処理を開始します。
Public Sub Generate_QR_Codes()

    Export_QR_WithMargin_PNG

End Sub

' Deletes QR code images inserted in column D.
' D列に挿入されたQRコード画像を削除します。
Public Sub Clear_QR_Pictures()

    Delete_Pictures_In_ColumnD

End Sub

'=============================================================
' Configuration
' 設定
'=============================================================

' Name of the subfolder used to store generated PNG files.
' 生成したPNGファイルを保存するサブフォルダ名です。
Private Const FOLDER_PNG As String = "QRコード画像"

' QR code image size in pixels.
' QRコード画像のサイズ（ピクセル）です。
Private Const QR_PX As Long = 512

' Quiet-zone margin around the QR code in pixels.
' QRコード周囲の余白（quiet zone／ピクセル）です。
Private Const QR_MARGIN As Long = 40

' First worksheet row containing source data.
' 元データが開始するワークシート行です。
Private Const START_ROW As Long = 2

' Maximum worksheet row processed in one batch.
' 1回の一括処理で対象とする最大行です。
Private Const MAX_ROW As Long = 201

' Worksheet column containing the sequential number.
' 通し番号を格納する列です。
Private Const COL_NO As String = "A"

' Worksheet column containing the QR payload.
' QRコードへ格納するデータを含む列です。
Private Const COL_DATA As String = "B"

' Worksheet column used for OK / NG / Skip results.
' OK／NG／スキップの判定結果を表示する列です。
Private Const COL_JUDGE As String = "C"

' Worksheet column used to insert QR code images.
' QRコード画像を挿入する列です。
Private Const COL_PIC As String = "D"

' Controls whether generated QR images are inserted into Excel.
' 生成したQR画像をExcelシートへ挿入するかを制御します。
Private Const PUT_PICTURE_IN_SHEET As Boolean = True

' Row height applied when inserting QR code images.
' QRコード画像を挿入する際に設定する行の高さです。
Private Const ROW_HEIGHT As Double = 150

' Padding between the QR image and the target cell border.
' QR画像と対象セルの境界との余白です。
Private Const PAD As Double = 2

' Controls whether cleaned QR data overwrites column B.
' クリーン化したQRデータでB列を上書きするかを制御します。
Private Const CLEAN_B_COLUMN_BEFORE_EXPORT As Boolean = True

' Controls whether the user selects the output folder.
' ユーザーに保存先フォルダを選択させるかを制御します。
Private Const USE_FOLDER_PICKER As Boolean = True

'=============================================================
' Main Workflow
' メイン処理
'=============================================================

' Generates QR codes for all valid rows in the active worksheet.
' Validates inputs, exports PNG files, optionally inserts images,
' displays progress, and writes a processing log.
'
' アクティブシートの有効な行を対象にQRコードを生成します。
' 入力検証、PNG出力、画像貼り付け、進捗表示、処理ログ記録を行います。
Public Sub Export_QR_WithMargin_PNG()

    Dim ws As Worksheet
    Dim startTime As Double
    Dim baseFolder As String
    Dim pngDir As String
    Dim lastRow As Long
    Dim r As Long
    Dim okCount As Long
    Dim ngCount As Long
    Dim skipCount As Long
    Dim totalCount As Long
    Dim noVal As String
    Dim payload As String
    Dim fileBase As String
    Dim savePath As String
    Dim urlPng As String
    Dim ok As Boolean
    Dim errMsg As String
    Dim elapsedTime As Double

    Set ws = ActiveSheet
    startTime = Timer

    ' Ask the user for an output folder when folder selection is enabled.
    ' フォルダ選択が有効な場合は、ユーザーに保存先を選択してもらいます。
    If USE_FOLDER_PICKER Then
        baseFolder = SelectOutputFolder()

        ' Stop safely when the folder-selection dialog is cancelled.
        ' 保存先選択がキャンセルされた場合は、安全に処理を終了します。
        If baseFolder = "" Then
            MsgBox "保存先の選択がキャンセルされました。", vbInformation
            GoTo FINALLY
        End If
    Else
        baseFolder = GetBaseFolder()
    End If

    pngDir = baseFolder & "\" & FOLDER_PNG

    ' Ensure the required output folders exist before processing.
    ' 処理開始前に必要な出力フォルダが存在することを確認します。
    EnsureFolder baseFolder
    EnsureFolder pngDir

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo FINALLY

    ' Detect the last source-data row from columns A and B.
    ' A列とB列を確認し、元データの最終行を取得します。
    lastRow = GetLastDataRow(ws)

    ' Stop when no source data exists below the header row.
    ' ヘッダー行より下に元データがない場合は処理を終了します。
    If lastRow < START_ROW Then
        MsgBox "QRコードを生成するデータがありません。", vbInformation
        GoTo FINALLY
    End If

    ' Limit the batch to the configured maximum row.
    ' 設定された最大行を超えないよう、処理範囲を制限します。
    If lastRow > MAX_ROW Then
        lastRow = MAX_ROW
    End If

    ' Optionally clean and overwrite the source values in column B.
    ' 必要に応じてB列の元データをクリーン化し、上書きします。
    If CLEAN_B_COLUMN_BEFORE_EXPORT Then
        CleanColumnB ws, START_ROW, lastRow, COL_DATA
    End If

    totalCount = lastRow - START_ROW + 1

    For r = START_ROW To lastRow

        ' Display real-time batch progress in the Excel status bar.
        ' Excelのステータスバーへリアルタイムの処理進捗を表示します。
        Application.StatusBar = _
            "Generating QR Codes... " & _
            (r - START_ROW + 1) & _
            " / " & totalCount

        ' Keep Excel responsive during long batch operations.
        ' 大量処理中もExcelが応答できるようにします。
        DoEvents

        ws.Cells(r, COL_JUDGE).Value = ""

        ' Expand the row to fit the QR image when worksheet insertion is enabled.
        ' シートへの画像挿入が有効な場合、QR画像が収まるよう行を広げます。
        If PUT_PICTURE_IN_SHEET Then
            ws.Rows(r).RowHeight = ROW_HEIGHT
        End If

        noVal = Trim$(CStr(ws.Cells(r, COL_NO).Value))

        ' Remove hidden or unwanted characters before QR generation.
        ' QR生成前に不可視文字や不要文字を除去します。
        payload = CleanQrText(ws.Cells(r, COL_DATA).Value)

        ' Skip a completely empty row without treating it as an error.
        ' A列・B列の両方が空の行は、エラーにせずスキップします。
        If noVal = "" And payload = "" Then
            skipCount = skipCount + 1
            GoTo NextR
        End If

        ' Reject rows that do not contain a sequential number.
        ' 通し番号が入力されていない行はNGとします。
        If noVal = "" Then
            ws.Cells(r, COL_JUDGE).Value = "NG：A列(通し番号)空"
            ngCount = ngCount + 1
            GoTo NextR
        End If

        ' Reject rows that do not contain QR payload data.
        ' QRコードへ格納するデータがない行はNGとします。
        If payload = "" Then
            ws.Cells(r, COL_JUDGE).Value = "NG：B列(QRデータ)空"
            ngCount = ngCount + 1
            GoTo NextR
        End If

        ' Convert the sequential number to a zero-padded file name.
        ' 通し番号を3桁のゼロ埋めファイル名へ変換します。
        fileBase = ZeroPad3(noVal)
        savePath = pngDir & "\" & fileBase & ".png"

        ' Build the QR generation API URL with UTF-8 encoded data.
        ' UTF-8でエンコードしたデータを使い、QR生成APIのURLを組み立てます。
        urlPng = BuildQrServerUrl(payload, QR_PX, QR_MARGIN)

        ' Download the generated QR image as a PNG file.
        ' 生成されたQR画像をPNGファイルとしてダウンロードします。
        errMsg = ""
        ok = HttpDownloadBinary(urlPng, savePath, errMsg)

        If ok Then
            ws.Cells(r, COL_JUDGE).Value = "OK"
            okCount = okCount + 1

            ' Optionally insert the generated image into the worksheet.
            ' 必要に応じて生成した画像をワークシートへ挿入します。
            If PUT_PICTURE_IN_SHEET Then
                If Not TryInsertPictureFillCell(ws, r, COL_PIC, savePath, PAD) Then
                    ws.Cells(r, COL_JUDGE).Value = "OK（画像貼付不可）"
                End If
            End If
        Else
            ws.Cells(r, COL_JUDGE).Value = "NG：" & errMsg
            ngCount = ngCount + 1
        End If

NextR:
    Next r

    ' Calculate elapsed processing time, including midnight rollover.
    ' 深夜0時をまたぐ場合も考慮して処理時間を計算します。
    elapsedTime = Timer - startTime

    If elapsedTime < 0 Then
        elapsedTime = elapsedTime + 86400
    End If

    ' Record the batch result in the Log worksheet.
    ' 一括処理の結果をLogシートへ記録します。
    WriteProcessingLog _
        ThisWorkbook, _
        okCount, _
        ngCount, _
        skipCount, _
        elapsedTime, _
        pngDir

    ' Display the processing summary to the user.
    ' 処理結果のサマリーをユーザーへ表示します。
    MsgBox "QRコード生成が完了しました。" & vbCrLf & _
           "OK：" & okCount & "件" & vbCrLf & _
           "NG：" & ngCount & "件" & vbCrLf & _
           "スキップ：" & skipCount & "件" & vbCrLf & _
           "処理時間：" & Format$(elapsedTime, "0.00") & " 秒" & vbCrLf & _
           "保存先：" & pngDir, vbInformation

FINALLY:
    ' Always restore Excel application settings before exiting.
    ' 正常終了・エラー終了にかかわらず、Excelの設定を必ず元に戻します。
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'=============================================================
' Worksheet Range Utilities
' ワークシート範囲関連
'=============================================================

' Returns the lowest used row found in the number or data column.
' 通し番号列またはデータ列のうち、より下にある最終使用行を返します。
Private Function GetLastDataRow(ByVal ws As Worksheet) As Long

    Dim lastRowNo As Long
    Dim lastRowData As Long

    lastRowNo = ws.Cells(ws.Rows.Count, COL_NO).End(xlUp).Row
    lastRowData = ws.Cells(ws.Rows.Count, COL_DATA).End(xlUp).Row

    If lastRowNo > lastRowData Then
        GetLastDataRow = lastRowNo
    Else
        GetLastDataRow = lastRowData
    End If

End Function

' Cleans the configured QR source-data column in place.
' 指定されたQR元データ列をクリーン化し、セルへ上書きします。
Private Sub CleanColumnB( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal colLetter As String)

    Dim r As Long

    For r = startRow To endRow
        ws.Cells(r, colLetter).Value = CleanQrText(ws.Cells(r, colLetter).Value)
    Next r

End Sub

'=============================================================
' Output Folder Utilities
' 出力フォルダ関連
'=============================================================

' Opens a folder-picker dialog and returns the selected folder path.
' フォルダ選択ダイアログを開き、選択されたフォルダパスを返します。
Private Function SelectOutputFolder() As String

    Dim dialog As FileDialog
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)

    With dialog
        .Title = "QRコード画像の保存先を選択してください"
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectOutputFolder = .SelectedItems(1)
        Else
            SelectOutputFolder = ""
        End If
    End With

End Function

' Returns the default output folder when the folder picker is disabled.
' フォルダ選択を使用しない場合の既定保存先を返します。
Private Function GetBaseFolder() As String

    If ThisWorkbook.Path <> "" Then
        GetBaseFolder = ThisWorkbook.Path & "\output"
    Else
        GetBaseFolder = Environ$("USERPROFILE") & "\Pictures\QRコード生成"
    End If

End Function

' Creates a folder when it does not already exist.
' 対象フォルダが存在しない場合に新規作成します。
Private Sub EnsureFolder(ByVal path As String)

    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If

End Sub

'=============================================================
' QR Code Generation
' QRコード生成関連
'=============================================================

' Builds the QR Server API URL for a PNG image.
' PNG形式のQR画像を生成するQR Server APIのURLを組み立てます。
Private Function BuildQrServerUrl( _
    ByVal data As String, _
    ByVal px As Long, _
    ByVal marginPx As Long) As String

    Dim url As String

    url = "https://api.qrserver.com/v1/create-qr-code/?" & _
          "format=png" & _
          "&data=" & UrlEncodeUTF8(data)

    If px > 0 Then
        url = url & "&size=" & px & "x" & px
    End If

    If marginPx > 0 Then
        url = url & "&margin=" & marginPx
    End If

    BuildQrServerUrl = url

End Function

' Downloads binary HTTP content and saves it to the specified file.
' HTTPからバイナリデータを取得し、指定されたファイルへ保存します。
Private Function HttpDownloadBinary( _
    ByVal url As String, _
    ByVal savePath As String, _
    ByRef errMsg As String) As Boolean

    On Error GoTo EH

    Dim http As Object
    Dim stm As Object

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "ExcelVBA"
    http.send

    ' Stop when the QR generation service returns a non-success status.
    ' QR生成サービスが正常以外のステータスを返した場合は終了します。
    If http.Status <> 200 Then
        errMsg = "HTTP " & http.Status & "（PNG取得失敗）"
        HttpDownloadBinary = False
        Exit Function
    End If

    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 1
    stm.Open
    stm.Write http.responseBody
    stm.SaveToFile savePath, 2
    stm.Close

    HttpDownloadBinary = True
    Exit Function

EH:
    errMsg = "PNG保存エラー：" & Err.Description
    HttpDownloadBinary = False

End Function

' Encodes text as UTF-8 for use in a URL without including a BOM.
' URLへ安全に含めるため、BOMを除外して文字列をUTF-8エンコードします。
Private Function UrlEncodeUTF8(ByVal s As String) As String

    Dim stm As Object
    Dim bytes() As Byte
    Dim startIdx As Long
    Dim i As Long
    Dim b As Integer
    Dim out As String

    ' Remove any embedded Unicode BOM character before conversion.
    ' 変換前に文字列内のUnicode BOM文字を除去します。
    s = Replace(s, ChrW(&HFEFF), "")

    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s
    stm.Position = 0
    stm.Type = 1

    bytes = stm.Read
    stm.Close

    startIdx = 0

    ' Skip the UTF-8 BOM bytes EF BB BF when ADODB.Stream adds them.
    ' ADODB.Streamが付与したUTF-8 BOM（EF BB BF）を読み飛ばします。
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And _
           bytes(1) = &HBB And _
           bytes(2) = &HBF Then

            startIdx = 3
        End If
    End If

    ' Preserve RFC 3986 unreserved characters and percent-encode others.
    ' RFC 3986の非予約文字はそのまま残し、それ以外をパーセントエンコードします。
    For i = startIdx To UBound(bytes)

        b = bytes(i)

        Select Case b
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                out = out & Chr$(b)

            Case 32
                out = out & "%20"

            Case Else
                out = out & "%" & Right$("0" & Hex$(b), 2)
        End Select

    Next i

    UrlEncodeUTF8 = out

End Function

' Converts a numeric sequential number to a three-digit file name.
' 数値の通し番号を3桁ゼロ埋めのファイル名へ変換します。
Private Function ZeroPad3(ByVal s As String) As String

    Dim n As Double

    If IsNumeric(s) Then
        n = CDbl(s)

        If n >= 0 And n < 1000 Then
            ZeroPad3 = Format$(CLng(n), "000")
        Else
            ZeroPad3 = CStr(s)
        End If
    Else
        ZeroPad3 = CStr(s)
    End If

End Function

'=============================================================
' Worksheet Image Utilities
' ワークシート画像関連
'=============================================================

' Inserts a PNG image into the target cell while preserving aspect ratio.
' 縦横比を維持しながら、PNG画像を対象セル内へ挿入します。
Private Function TryInsertPictureFillCell( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal colLetter As String, _
    ByVal imgPath As String, _
    ByVal pad As Double) As Boolean

    On Error GoTo EH

    Dim tgt As Range
    Dim i As Long
    Dim pic As Shape
    Dim maxW As Double
    Dim maxH As Double

    Set tgt = ws.Cells(rowNum, colLetter)

    ' Delete an existing image anchored to the target cell.
    ' 対象セルに既存画像がある場合は削除します。
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).TopLeftCell.Address = tgt.Address Then
            ws.Shapes(i).Delete
        End If
    Next i

    Set pic = ws.Shapes.AddPicture( _
        Filename:=imgPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=tgt.Left + pad, _
        Top:=tgt.Top + pad, _
        Width:=-1, _
        Height:=-1)

    pic.Placement = xlMoveAndSize
    pic.LockAspectRatio = msoTrue

    maxW = tgt.Width - (pad * 2)
    maxH = tgt.Height - (pad * 2)

    pic.Height = maxH

    If pic.Width > maxW Then
        pic.Width = maxW
    End If

    ' Center the image inside the target cell.
    ' 対象セルの中央へ画像を配置します。
    pic.Left = tgt.Left + (tgt.Width - pic.Width) / 2
    pic.Top = tgt.Top + (tgt.Height - pic.Height) / 2

    TryInsertPictureFillCell = True
    Exit Function

EH:
    TryInsertPictureFillCell = False

End Function

' Removes all shapes whose top-left cell is located in column D.
' 左上の基準セルがD列にあるすべてのShapeを削除します。
Public Sub Delete_Pictures_In_ColumnD()

    Dim ws As Worksheet
    Dim i As Long

    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    On Error GoTo FINALLY

    ' Iterate backwards so deleting a shape does not disturb the loop.
    ' Shape削除によるコレクション番号のずれを避けるため、後ろから処理します。
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).TopLeftCell.Column = 4 Then
            ws.Shapes(i).Delete
        End If
    Next i

    MsgBox "D列の画像を削除しました。", vbInformation

FINALLY:
    Application.ScreenUpdating = True

End Sub

'=============================================================
' Text Cleanup
' 文字列クリーンアップ
'=============================================================

' Removes characters that can corrupt or unintentionally alter QR data.
' QRデータを壊したり意図せず変更したりする可能性のある文字を除去します。
Private Function CleanQrText(ByVal v As Variant) As String

    Dim s As String

    If IsError(v) Then
        CleanQrText = ""
        Exit Function
    End If

    s = CStr(v)

    ' Remove line breaks and tabs.
    ' 改行とタブを除去します。
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")

    ' Replace a non-breaking space with a normal space.
    ' ノーブレークスペースを通常のスペースへ置換します。
    s = Replace(s, ChrW(&HA0), " ")

    ' Remove zero-width and invisible formatting characters.
    ' ゼロ幅文字および不可視の書式制御文字を除去します。
    s = Replace(s, ChrW(&H200B), "")
    s = Replace(s, ChrW(&H200C), "")
    s = Replace(s, ChrW(&H200D), "")
    s = Replace(s, ChrW(&H2060), "")

    ' Remove Unicode BOM characters from anywhere in the text.
    ' 文字列内に混入したUnicode BOM文字をすべて除去します。
    s = Replace(s, ChrW(&HFEFF), "")

    CleanQrText = Trim$(s)

End Function

'=============================================================
' Processing Log
' 処理ログ
'=============================================================

' Appends one batch-processing result to the Log worksheet.
' 1回分の一括処理結果をLogシートの最終行へ追記します。
Private Sub WriteProcessingLog( _
    ByVal wb As Workbook, _
    ByVal okCount As Long, _
    ByVal ngCount As Long, _
    ByVal skipCount As Long, _
    ByVal elapsedTime As Double, _
    ByVal outputFolder As String)

    Dim logWs As Worksheet
    Dim nextRow As Long

    Set logWs = GetOrCreateLogSheet(wb)

    nextRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1

    If nextRow < 2 Then
        nextRow = 2
    End If

    logWs.Cells(nextRow, 1).Value = Now
    logWs.Cells(nextRow, 2).Value = okCount
    logWs.Cells(nextRow, 3).Value = ngCount
    logWs.Cells(nextRow, 4).Value = skipCount
    logWs.Cells(nextRow, 5).Value = elapsedTime
    logWs.Cells(nextRow, 6).Value = outputFolder

    logWs.Cells(nextRow, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    logWs.Cells(nextRow, 5).NumberFormat = "0.00"

End Sub

' Returns the existing Log worksheet or creates it when missing.
' Logシートが存在する場合は取得し、存在しない場合は新規作成します。
Private Function GetOrCreateLogSheet(ByVal wb As Workbook) As Worksheet

    Dim logWs As Worksheet

    On Error Resume Next
    Set logWs = wb.Worksheets("Log")
    On Error GoTo 0

    If logWs Is Nothing Then

        Set logWs = wb.Worksheets.Add( _
            After:=wb.Worksheets(wb.Worksheets.Count))

        logWs.Name = "Log"

        With logWs
            .Cells(1, 1).Value = "Executed At"
            .Cells(1, 2).Value = "OK Count"
            .Cells(1, 3).Value = "NG Count"
            .Cells(1, 4).Value = "Skip Count"
            .Cells(1, 5).Value = "Elapsed Time (sec)"
            .Cells(1, 6).Value = "Output Folder"

            .Rows(1).Font.Bold = True
            .Columns("A:F").AutoFit
        End With

    End If

    Set GetOrCreateLogSheet = logWs

End Function
