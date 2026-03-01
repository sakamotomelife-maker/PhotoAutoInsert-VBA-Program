Attribute VB_Name = "selectFileandFolder"
Option Explicit


' STEP1：Excelファイルと画像フォルダの選択
Sub selectExcelFile()

    Dim filePath As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "商品一覧エクセル- Excelファイルを選択してください"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"

    If fd.Show <> -1 Then Exit Sub
    filePath = fd.SelectedItems(1)

    ThisWorkbook.Sheets("挿入ツール").Range("E10").Value = filePath
End Sub


Sub selectImagesInFolder()

    Dim folderPath As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "画像フォルダの選択"
    If fd.Show <> -1 Then Exit Sub
    folderPath = fd.SelectedItems(1)
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ThisWorkbook.Sheets("挿入ツール").Range("E12").Value = folderPath

End Sub



' STEP2：PNG＋JPG → JPEG 縮小（WIA）＋自動回転
Sub ResizeImagesToCompressedFolder()

    Dim folderPath As String
    Dim outFolder As String
    Dim newWidth As Long
    Dim newQuality As Long
    Dim f As String
    Dim srcPath As String
    Dim tmpPng As String
    Dim dstPath As String
    Dim files As Object
    Dim i As Long
    Dim ext As String

    folderPath = ThisWorkbook.Sheets("挿入ツール").Range("E12").Value   '画像フォルダパス
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    newWidth = CLng(ThisWorkbook.Sheets("挿入ツール").OLEObjects("TextBox1").Object.text)
    newQuality = CLng(ThisWorkbook.Sheets("挿入ツール").OLEObjects("TextBox2").Object.text)

    outFolder = folderPath & "圧縮\"
    If Dir(outFolder, vbDirectory) = "" Then MkDir outFolder

    Set files = CreateObject("System.Collections.ArrayList")

    ' PNG
    f = Dir(folderPath & "*.png")
    Do While f <> ""
        files.Add f
        f = Dir
    Loop

    ' JPG
    f = Dir(folderPath & "*.jpg")
    Do While f <> ""
        files.Add f
        f = Dir
    Loop

    ' JPEG
    f = Dir(folderPath & "*.jpeg")
    Do While f <> ""
        files.Add f
        f = Dir
    Loop

    If files.Count = 0 Then
        MsgBox "PNG / JPG ファイルが見つかりませんでした。"
        Exit Sub
    End If

    For i = 0 To files.Count - 1

        f = files(i)
        srcPath = folderPath & f
        ext = LCase$(Mid(f, InStrRev(f, ".") + 1))

        tmpPng = folderPath & "tmp_wia_" & CStr(i) & ".png"
        dstPath = outFolder & Left(f, InStrRev(f, ".")) & "jpg"

        If ext = "png" Then
            ' PNG → 再エンコード → JPEG
            WiaReencodePng srcPath, tmpPng
            WiaPngToJpegResize tmpPng, dstPath, newWidth, newQuality
            If Dir(tmpPng) <> "" Then Kill tmpPng

        ElseIf ext = "jpg" Or ext = "jpeg" Then
            ' JPG → 直接縮小
            WiaJpegResize srcPath, dstPath, newWidth, newQuality
        End If

    Next i

    MsgBox "圧縮フォルダへの画像生成が完了しました。"

End Sub


' PNG → PNG（再エンコード）
Sub WiaReencodePng(src As String, dst As String)

    Dim img As Object
    Dim ip As Object
    Dim result As Object

    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile src

    Set ip = CreateObject("WIA.ImageProcess")
    ip.Filters.Add ip.FilterInfos("Convert").FilterID
    ip.Filters(1).Properties("FormatID") = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"

    Set result = ip.Apply(img)

    If Dir(dst) <> "" Then Kill dst
    result.SaveFile dst

    Set img = Nothing
    Set ip = Nothing
    Set result = Nothing

End Sub



' PNG → JPEG 変換＋縮小＋自動回転
Sub WiaPngToJpegResize(src As String, dst As String, maxWidth As Long, quality As Long)

    Dim img As Object
    Dim ip As Object
    Dim ratio As Double
    Dim newW As Long, newH As Long
    Dim result As Object

    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile src

    Set ip = CreateObject("WIA.ImageProcess")

    '横長なら90°回転
    If img.Width > img.Height Then
        ip.Filters.Add ip.FilterInfos("RotateFlip").FilterID
        ip.Filters(1).Properties("RotationAngle") = 90
    End If

    '縮小
    ratio = img.Width / img.Height
    If img.Width > maxWidth Then
        newW = maxWidth
        newH = maxWidth / ratio
    Else
        newW = img.Width
        newH = img.Height
    End If

    ip.Filters.Add ip.FilterInfos("Scale").FilterID
    ip.Filters(ip.Filters.Count).Properties("MaximumWidth") = newW
    ip.Filters(ip.Filters.Count).Properties("MaximumHeight") = newH

    'JPEG変換
    ip.Filters.Add ip.FilterInfos("Convert").FilterID
    ip.Filters(ip.Filters.Count).Properties("FormatID") = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
    ip.Filters(ip.Filters.Count).Properties("Quality") = quality

    Set result = ip.Apply(img)

    If Dir(dst) <> "" Then Kill dst
    result.SaveFile dst

End Sub


' JPG → JPEG 縮小＋自動回転
Sub WiaJpegResize(src As String, dst As String, maxWidth As Long, quality As Long)

    Dim img As Object
    Dim ip As Object
    Dim ratio As Double
    Dim newW As Long, newH As Long
    Dim result As Object

    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile src

    Set ip = CreateObject("WIA.ImageProcess")

    ' 横長なら 90° 回転
    If img.Width > img.Height Then
        ip.Filters.Add ip.FilterInfos("RotateFlip").FilterID
        ip.Filters(1).Properties("RotationAngle") = 90
    End If

    ' 縮小
    ratio = img.Width / img.Height
    If img.Width > maxWidth Then
        newW = maxWidth
        newH = maxWidth / ratio
    Else
        newW = img.Width
        newH = img.Height
    End If

    ip.Filters.Add ip.FilterInfos("Scale").FilterID
    ip.Filters(ip.Filters.Count).Properties("MaximumWidth") = newW
    ip.Filters(ip.Filters.Count).Properties("MaximumHeight") = newH

    ' JPEG 変換
    ip.Filters.Add ip.FilterInfos("Convert").FilterID
    ip.Filters(ip.Filters.Count).Properties("FormatID") = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
    ip.Filters(ip.Filters.Count).Properties("Quality") = quality

    Set result = ip.Apply(img)

    If Dir(dst) <> "" Then Kill dst
    result.SaveFile dst

End Sub



' STEP3：圧縮フォルダ内の画像を Excel へ貼付
Sub InsertCompressedImagesToExcel()

    Dim folderPath As String
    Dim compressedFolder As String
    Dim filePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim startCell As Range
    Dim files As Object
    Dim f As String
    Dim i As Long
    Dim img As Shape
    Dim picPath As String

    folderPath = ThisWorkbook.Sheets("挿入ツール").Range("E12").Value
    If folderPath = "" Then
        MsgBox "画像フォルダが指定されていません。"
        Exit Sub
    End If
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    compressedFolder = folderPath & "圧縮\"
    If Dir(compressedFolder, vbDirectory) = "" Then
        MsgBox "圧縮フォルダが存在しません。先に圧縮処理を実行してください。"
        Exit Sub
    End If

    filePath = ThisWorkbook.Sheets("挿入ツール").Range("E10").Value
    If filePath = "" Then
        MsgBox "対象Excelが指定されていません。"
        Exit Sub
    End If

    Set wb = Workbooks.Open(filePath)
    Set ws = wb.ActiveSheet

    Set startCell = ws.Range("G4")
    Set files = CreateObject("System.Collections.ArrayList")

    f = Dir(compressedFolder & "*.jpg")
    Do While f <> ""
        files.Add compressedFolder & f
        f = Dir
    Loop

    files.Sort

    For i = 0 To files.Count - 1

        picPath = files(i)

        Set img = ws.Shapes.AddPicture( _
            fileName:=picPath, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=startCell.Left, _
            Top:=startCell.Offset(i, 0).Top, _
            Width:=-1, _
            Height:=-1)

        img.LockAspectRatio = msoTrue
        img.Width = 165

    Next i

    MsgBox "圧縮画像の貼付が完了しました。"

End Sub


' 写真が入るように幅の調整
Sub resizeExcel()

    Dim filePath As String
    Dim wb As Workbook
    Dim ws As Worksheet

    filePath = ThisWorkbook.Sheets("挿入ツール").Range("E10").Value
    If filePath = "" Then
        MsgBox "対象Excelが指定されていません。"
        Exit Sub
    End If

    Set wb = Workbooks.Open(filePath)
    Set ws = wb.ActiveSheet

    ws.Columns("G:G").ColumnWidth = 30
    ws.Range(Rows("4:4"), Rows("4:4").End(xlDown)).RowHeight = 226.8
    
    Dim last As Long
    last = ws.Cells(Rows.Count, 5).End(xlUp).Row
    ws.Range("F4:F" & last) = "〇"
    
    ws.Range("G1") = Date
   
   
   If ws.Range("G2") = "" Then
      MsgBox "処理が完了しました。" & vbCrLf & "※氏名を入力してください。"
   Else
      MsgBox "処理が完了しました。"
    End If
End Sub


