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
    
    ' 現在開いているWorkbookを使用
    Set dataWb = ThisWorkbook
    Set ws1 = ActiveSheet ' 当前excel（thisworkbook）里正在活动的sheet
    lastRow1 = ws1.Cells(ws1.Rows.Count, "C").End(xlUp).Row  ' 从c列最后一行上跳，跳到最后一个有数据的单元格，把它的行号作为 lastRow1
    
    ' 调用user choose file功能 但是这个函数的定义在最下面写了 这和py不同 不用先定义再调用
    filePath = GetPlanFilePath()
    If filePath = "" Then '如果是空字符串就输出下面这行 空字符串对应user取消选择文件
        MsgBox "ファイルが選択されませんでした。処理を中止します。", vbInformation ' msgbox样示这是蓝色信息图标
        Exit Sub
    End If
    
    ' planファイルを開く
    Set planWb = Workbooks.Open(filePath)'Workbooks.Open这是函数 所以后面的filepath腰括号
    Set ws2 = planWb.Sheets("plan")
    lastRow2 = ws2.Cells(ws2.Rows.Count, "E").End(xlUp).Row ' 从 E 列最后一行上跳，跳到最后一个有数据的单元格，把它的行号作为 lastRow2
    'cells（行，列） 
    'ws2.Rows.Count这是总行数的意思
    'End(xlUp) 相当于按住 Ctrl + ↑，跳到数据区域的最上方

    
    ' delete_sheetとnew_dataシートを現在のワークブックに作成/再作成
    Application.DisplayAlerts = False
    On Error Resume Next
    ' 既存のシートがあれば削除
    If SheetExists(dataWb, "delete_sheet") Then
        dataWb.Sheets("delete_sheet").Delete
    End If
    If SheetExists(dataWb, "new_data") Then
        dataWb.Sheets("new_data").Delete
    End If
    On Error GoTo 0
    
    ' 現在のワークブックに新しいシートを追加
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
    
    ' カウンターの初期化
    deleteRowCount = 3 ' ヘッダー行の下から始める
    newRowCount = 3
    
    ' 削除対象行と保持行を振り分け
    For i = 3 To lastRow1  ' 3行目から開始（ヘッダー行をスキップ）
        id1 = ws1.Cells(i, 3).Value   ' test2_dataのC列からID取得
        fVal = ws1.Cells(i, 6).Value  ' F列の値
        found = False
        shouldDelete = False

        ' 条件F：F列が「保管費」または「試験費」
        If fVal = "保管費" Or fVal = "試験費" Then
            Debug.Print "条件Fヒット：行 " & i & " → " & id1 & "（F列：" & fVal & "）→ 削除対象"
            shouldDelete = True
        Else
            ' plan.xlsx 内のIDを検索
            For j = 3 To lastRow2  ' 3行目からスタート
                id2 = ws2.Cells(j, 5).Value  ' test_planのE列からID取得
                
                If id1 = id2 Then
                    found = True
                    rVal = ws2.Cells(j, 18).Value    ' R列 = 18列目
                    lastVal = ws2.Cells(j, 35).Value ' last列 = 35列目
                    
                    ' 条件D：R列に "T" を含む
                    If InStr(rVal, "T") > 0 Then
                        Debug.Print "条件Dヒット：R列にTを含む → " & id1 & " → 削除対象"
                        shouldDelete = True
                        Exit For
                    End If
                    
                    ' 条件E：R列が空、last列が「社ｓ産」→削除
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
                    
                    ' 条件A / B：R列に値がある → 保存
                    If rVal <> "" Then
                        Debug.Print "条件A/Bヒット：R列に値あり → " & id1 & " → 保存"
                        shouldDelete = False
                        Exit For
                    End If
                    
                    Exit For
                End If
            Next j
            
            ' 条件C：plan.xlsx に存在しないID → 保存
            If Not found Then
                Debug.Print "条件Cヒット：planに存在しないID → " & id1 & " → 保存"
                shouldDelete = False
            End If
        End If
        
        ' 行全体をコピーして適切なシートに貼り付け
        If shouldDelete = True Then
            Debug.Print ">>> 削除リストに追加：行 " & i & " → " & id1
            ws1.Rows(i).Copy wsDelete.Rows(deleteRowCount)
            deleteRowCount = deleteRowCount + 1
        Else
            ws1.Rows(i).Copy wsNew.Rows(newRowCount)
            newRowCount = newRowCount + 1
        End If
    Next i
    
    ' planファイルを閉じる（変更を保存しない）
    planWb.Close SaveChanges:=False
    
    MsgBox "処理が完了しました。" & vbCrLf & _
           "削除対象行は delete_sheet に、" & vbCrLf & _
           "残りのデータは new_data シートに保存されました。"
End Sub

' function for file choose
Function GetPlanFilePath() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "プランファイルを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel ファイル", "*.xlsx; *.xls"
        .InitialFileName = "C:\Users\" ' 初期フォルダを設定
        
        If .Show = -1 Then
            GetPlanFilePath = .SelectedItems(1)
        Else
            GetPlanFilePath = ""
        End If
    End With
End Function

' 指定したワークブックに特定の名前のシートが存在するかチェックする関数
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