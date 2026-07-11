Option Explicit

Public Sub Generate_QR_Codes()

    Export_QR_WithMargin_PNG

End Sub

Public Sub Clear_QR_Pictures()

    Delete_Pictures_In_ColumnD

End Sub

' =========================
' Settings
' =========================
Const FOLDER_PNG As String = "QRコード画像"

Const QR_PX As Long = 512
Const QR_MARGIN As Long = 40

Const START_ROW As Long = 2
Const MAX_ROW As Long = 201

Const COL_NO As String = "A"
Const COL_DATA As String = "B"
Const COL_JUDGE As String = "C"
Const COL_PIC As String = "D"

Const PUT_PICTURE_IN_SHEET As Boolean = True
Const ROW_HEIGHT As Double = 150
Const PAD As Double = 2

Const CLEAN_B_COLUMN_BEFORE_EXPORT As Boolean = True

Public Sub Export_QR_WithMargin_PNG()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim baseFolder As String
    baseFolder = GetBaseFolder()

    Dim pngDir As String
    pngDir = baseFolder & "\" & FOLDER_PNG

    EnsureFolder baseFolder
    EnsureFolder pngDir

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo FINALLY
    
        Dim lastRow As Long
        lastRow = GetLastDataRow(ws)
        
        If lastRow < START_ROW Then
            MsgBox "QRコードを生成するデータがありません。", vbInformation
            GoTo FINALLY
        End If
        
        If lastRow > MAX_ROW Then
            lastRow = MAX_ROW
        End If
        
        If CLEAN_B_COLUMN_BEFORE_EXPORT Then
            CleanColumnB ws, START_ROW, lastRow, COL_DATA
        End If
        
        Dim r As Long
        Dim okCount As Long
        Dim ngCount As Long
        Dim skipCount As Long
        
        For r = START_ROW To lastRow

        ws.Cells(r, COL_JUDGE).Value = ""

        If PUT_PICTURE_IN_SHEET Then
            ws.Rows(r).RowHeight = ROW_HEIGHT
        End If

        Dim noVal As String
        noVal = Trim$(CStr(ws.Cells(r, COL_NO).Value))

        Dim payload As String
        payload = CleanQrText(ws.Cells(r, COL_DATA).Value)

        ' A列・B列の両方が空の場合は処理対象外
        If noVal = "" And payload = "" Then
            skipCount = skipCount + 1
            GoTo NextR
        End If

        If noVal = "" Then
            ws.Cells(r, COL_JUDGE).Value = "NG：A列(通し番号)空"
            ngCount = ngCount + 1
            GoTo NextR
        End If

        If payload = "" Then
            ws.Cells(r, COL_JUDGE).Value = "NG：B列(QRデータ)空"
            ngCount = ngCount + 1
            GoTo NextR
        End If

        Dim fileBase As String
        fileBase = ZeroPad3(noVal)

        Dim savePath As String
        savePath = pngDir & "\" & fileBase & ".png"

        Dim urlPng As String
        urlPng = BuildQrServerUrl(payload, QR_PX, QR_MARGIN)

        Dim ok As Boolean
        Dim errMsg As String

        ok = HttpDownloadBinary(urlPng, savePath, errMsg)

        If ok Then
            ws.Cells(r, COL_JUDGE).Value = "OK"
            okCount = okCount + 1

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

    MsgBox "QRコード生成が完了しました。" & vbCrLf & _
       "OK：" & okCount & "件" & vbCrLf & _
       "NG：" & ngCount & "件" & vbCrLf & _
       "スキップ：" & skipCount & "件" & vbCrLf & _
       "保存先：" & pngDir, vbInformation

FINALLY:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

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

Private Function GetBaseFolder() As String

    If ThisWorkbook.Path <> "" Then
        GetBaseFolder = ThisWorkbook.Path & "\output"
    Else
        GetBaseFolder = Environ$("USERPROFILE") & "\Pictures\QRコード生成"
    End If

End Function

Private Sub CleanColumnB(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal colLetter As String)

    Dim r As Long

    For r = startRow To endRow
        ws.Cells(r, colLetter).Value = CleanQrText(ws.Cells(r, colLetter).Value)
    Next r

End Sub

Private Function BuildQrServerUrl(ByVal data As String, ByVal px As Long, ByVal marginPx As Long) As String

    Dim url As String

    url = "https://api.qrserver.com/v1/create-qr-code/?" & _
          "format=png" & _
          "&data=" & UrlEncodeUTF8(data)

    If px > 0 Then url = url & "&size=" & px & "x" & px
    If marginPx > 0 Then url = url & "&margin=" & marginPx

    BuildQrServerUrl = url

End Function

Private Sub EnsureFolder(ByVal path As String)

    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If

End Sub

Private Function HttpDownloadBinary(ByVal url As String, ByVal savePath As String, ByRef errMsg As String) As Boolean

    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "ExcelVBA"
    http.send

    If http.Status <> 200 Then
        errMsg = "HTTP " & http.Status & "（PNG取得失敗）"
        HttpDownloadBinary = False
        Exit Function
    End If

    Dim stm As Object
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

Private Function UrlEncodeUTF8(ByVal s As String) As String

    s = Replace(s, ChrW(&HFEFF), "")

    Dim stm As Object
    Dim bytes() As Byte

    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s
    stm.Position = 0
    stm.Type = 1

    bytes = stm.Read
    stm.Close

    Dim startIdx As Long
    startIdx = 0

    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            startIdx = 3
        End If
    End If

    Dim i As Long
    Dim b As Integer
    Dim out As String

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

Private Function TryInsertPictureFillCell(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colLetter As String, ByVal imgPath As String, ByVal pad As Double) As Boolean

    On Error GoTo EH

    Dim tgt As Range
    Set tgt = ws.Cells(rowNum, colLetter)

    Dim shp As Shape

    For Each shp In ws.Shapes
        If shp.TopLeftCell.Address = tgt.Address Then
            shp.Delete
        End If
    Next shp

    Dim pic As Shape

    Set pic = ws.Shapes.AddPicture( _
        Filename:=imgPath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=tgt.Left + pad, _
        Top:=tgt.Top + pad, _
        Width:=-1, _
        Height:=-1 _
    )

    pic.Placement = xlMoveAndSize
    pic.LockAspectRatio = msoTrue

    Dim maxW As Double
    Dim maxH As Double

    maxW = tgt.Width - (pad * 2)
    maxH = tgt.Height - (pad * 2)

    pic.Height = maxH

    If pic.Width > maxW Then
        pic.Width = maxW
    End If

    pic.Left = tgt.Left + (tgt.Width - pic.Width) / 2
    pic.Top = tgt.Top + (tgt.Height - pic.Height) / 2

    TryInsertPictureFillCell = True
    Exit Function

EH:
    TryInsertPictureFillCell = False

End Function

Private Function CleanQrText(ByVal v As Variant) As String

    If IsError(v) Then
        CleanQrText = ""
        Exit Function
    End If

    Dim s As String
    s = CStr(v)

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")

    s = Replace(s, ChrW(&HA0), " ")

    s = Replace(s, ChrW(&H200B), "")
    s = Replace(s, ChrW(&H200C), "")
    s = Replace(s, ChrW(&H200D), "")
    s = Replace(s, ChrW(&H2060), "")

    s = Replace(s, ChrW(&HFEFF), "")

    CleanQrText = Trim$(s)

End Function

Public Sub Delete_Pictures_In_ColumnD()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim i As Long

    Application.ScreenUpdating = False
    On Error GoTo FINALLY

    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).TopLeftCell.Column = 4 Then
            ws.Shapes(i).Delete
        End If
    Next i

    MsgBox "D列の画像を削除しました。", vbInformation

FINALLY:
    Application.ScreenUpdating = True

End Sub
