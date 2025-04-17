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
    Dim headerRange As Range
    
    ' データファイルを参照
    Set dataWb = ThisWorkbook
    Set ws1 = dataWb.Sheets("test") ' ← シート名「test」
    lastRow1 = ws1.Cells(ws1.Rows.Count, "C").End(xlUp).Row
    
    ' ユーザーにplanファイルを選択させる
    filePath = GetPlanFilePath()
    If filePath = "" Then
        MsgBox "ファイルが選択されませんでした。処理を中止します。", vbExclamation
        Exit Sub
    End If
    
    ' planファイルを開く
    Set planWb = Workbooks.Open(filePath)
    Set ws2 = planWb.Sheets("plan")
    lastRow2 = ws2.Cells(ws2.Rows.Count, "E").End(xlUp).Row ' E列で最終行を取得
    
    ' delete_sheetとnew_dataシートを作成（既存の場合は削除して再作成）
    Call CreateOrReplaceSheet(dataWb, "delete_sheet", 2)
    Set wsDelete = dataWb.Sheets("delete_sheet")
    
    Call CreateOrReplaceSheet(dataWb, "new_data", 3)
    Set wsNew = dataWb.Sheets("new_data")
    
    ' ヘッダー行をコピー
    Set headerRange = ws1.Range("A1:Z2") ' 必要に応じて範囲を調整
    headerRange.Copy wsDelete.Range("A1")
    headerRange.Copy wsNew.Range("A1")
    
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
    
    planWb.Close SaveChanges:=False
    MsgBox "処理が完了しました。" & vbCrLf & _
           "削除対象行は delete_sheet に、" & vbCrLf & _
           "残りのデータは new_data シートに保存されました。"
End Sub

' ファイル選択ダイアログを表示し、選択されたファイルのパスを返す関数
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

' シートを作成または再作成する関数
Sub CreateOrReplaceSheet(wb As Workbook, sheetName As String, position As Integer)
    Dim ws As Worksheet
    
    ' 同名のシートが存在する場合は削除
    For Each ws In wb.Sheets
        If ws.Name = sheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws
    
    ' 新しいシートを追加
    wb.Sheets.Add After:=wb.Sheets(position - 1)
    ActiveSheet.Name = sheetName
End Sub