Sub FilterDataWithPlan()
    Dim dataWb As Workbook
    Dim planWb As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet, wsDelete As Worksheet, wsNew As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long, j As Long, deleteRowCount As Long, newRowCount As Long
    Dim id1 As String, id2 As String
    Dim fVal As String, rVal As String, lastVal As String
    Dim found As Boolean, shouldDelete As Boolean
    Dim filePath As String
    Dim wordlist As Variant
    Dim word As Variant


    ' A1. select testfile 
    filePath = GetFilePath("dataファイルを選択してください")
    If filePath = "" Then
        MsgBox "dataファイルが選択されませんでした。処理を中止します", vbInformation
        Exit Sub
    End If
    
    ' A2. open workbook which users has selected 
    Set dataWb = Workbooks.Open(filePath)

    ' A3. 让用户在要处理的sheet中点击任意单元格，以确定想用的sheet
    On Error Resume Next
    Set userSelection = Application.InputBox("処理したいシート上の任意のセルをクリックしてください", Type:=8)
    On Error GoTo 0

    If userSelection Is Nothing Then
        MsgBox "ワークシートが選択されませんでした。処理を中止します", vbInformation
        dataWb.Close False
        Exit Sub
    End If

    ' A4.
    Set ws1 = userSelection.Worksheet
    MsgBox "選択されたdataシート：" & ws1.Name, vbInformation
    lastRow1 = ws1.Cells(ws1.Rows.Count, "C").End(xlUp).Row

    ' B1. select planfile 
    filePath = GetFilePath("planファイルを選択してください")
    If filePath = "" Then
        MsgBox "planファイルが選択されませんでした。処理を中止します", vbInformation
        Exit Sub
    End If
    
    ' B2. open workbook which users has selected 
    Set planWb = Workbooks.Open(filePath)

    ' B3. 让用户在要处理的sheet中点击任意单元格，以确定想用的sheet
    On Error Resume Next
    Set userSelection = Application.InputBox("処理したいシート上の任意のセルをクリックしてください", Type:=8)
    On Error GoTo 0

    If userSelection Is Nothing Then
        MsgBox "ワークシートが選択されませんでした。処理を中止します", vbInformation
        dataWb.Close False
        Exit Sub
    End If

    ' B4.
    Set ws2 = userSelection.Worksheet
    MsgBox "選択されたplanシート：" & ws2.Name, vbInformation
    lastRow2 = ws2.Cells(ws2.Rows.Count, "E").End(xlUp).Row

    ' delete_sheet と new_data シートを再作成
    Application.DisplayAlerts = False
    On Error Resume Next
    If SheetExists(dataWb, "delete_sheet") Then dataWb.Sheets("delete_sheet").Delete
    If SheetExists(dataWb, "new_data") Then dataWb.Sheets("new_data").Delete
    On Error GoTo 0

    With dataWb
        .Sheets.Add(After:=.Sheets(1)).Name = "delete_sheet"
        Set wsDelete = .Sheets("delete_sheet")
        .Sheets.Add(After:=.Sheets("delete_sheet")).Name = "new_data"
        Set wsNew = .Sheets("new_data")
    End With
    Application.DisplayAlerts = True

    ' ヘッダー行をコピー
    ws1.Rows("1:2").Copy wsDelete.Range("A1")
    ws1.Rows("1:2").Copy wsNew.Range("A1")

    deleteRowCount = 3
    newRowCount = 3

    ' 条件分岐とデータ振り分け
    For i = 3 To lastRow1
        id1 = ws1.Cells(i, 3).Value
        fVal = ws1.Cells(i, 6).Value
        found = False
        shouldDelete = False

        ' 要识别的关键词，可以随时添加
        wordlist = Array("保管費", "試験費","立会対応費","総合試験")

            shouldDelete = False
            For Each word In wordlist
                If InStr(fVal, word) > 0 Then
                    Debug.Print "条件Fヒット：行 " & i & " → " & id1 & "（F列：" & fVal & "）→ 削除対象"
                    shouldDelete = True
                    Exit for
                End If
                Next word

                If Not shouldDelete Then
                    For j = 3 To lastRow2
                        id2 = ws2.Cells(j, 5).Value

                        If id1 = id2 Then
                            found = True
                            rVal = ws2.Cells(j, 18).Value
                            lastVal = ws2.Cells(j, 35).Value

                            If InStr(rVal, "T") > 0 Then
                                Debug.Print "条件Dヒット：R列にTを含む → " & id1 & " → 削除対象"
                                shouldDelete = True
                                Exit For
                            End If

                            If rVal = "" Or IsEmpty(rVal) Then
                                If lastVal = "社ｓ産" Then
                                    Debug.Print "条件Eヒット：R列空かつlast列「社ｓ産」 → " & id1 & " → 削除対象"
                                    shouldDelete = True
                                Else
                                    Debug.Print "条件Eヒット：R列空かつlast列その他 → " & id1 & " → 保存"
                                    shouldDelete = False
                                End If
                                Exit For
                            End If

                            If rVal <> "" Then
                                Debug.Print "条件A/Bヒット：R列に値あり → " & id1 & " → 保存"
                                shouldDelete = False
                                Exit For
                            End If

                            Exit For
                        End If
                    Next j

            If Not found Then
                Debug.Print "条件Cヒット：planに存在しないID → " & id1 & " → 保存"
                shouldDelete = False
            End If
        End If

        If shouldDelete = True Then
            Debug.Print ">>> 削除リストに追加：行 " & i & " → " & id1
            ws1.Rows(i).Copy wsDelete.Rows(deleteRowCount)
            deleteRowCount = deleteRowCount + 1
        Else
            ws1.Rows(i).Copy wsNew.Rows(newRowCount)
            newRowCount = newRowCount + 1
        End If
    Next i

    planWb.Close SaveChanges:=False

    MsgBox "処理が完了しました。" & vbCrLf & _
           "削除対象行は delete_sheet に、" & vbCrLf & _
           "残りのデータは new_data シートに保存されました。"
End Sub

' 汎用ファイル選択関数
Function GetFilePath(dialogTitle As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = dialogTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel ファイル", "*.xlsx; *.xls"
        .InitialFileName = "C:\Users\"
        If .Show = -1 Then
            GetFilePath = .SelectedItems(1)
        Else
            GetFilePath = ""
        End If
    End With
End Function

' シート存在確認関数
Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    SheetExists = False
    For Each ws In wb.Sheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function