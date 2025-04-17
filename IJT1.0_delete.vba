Sub FilterDataWithPlan()
    Dim dataWb As Workbook
    Dim planWb As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long, j As Long
    Dim id1 As String, id2 As String
    Dim fVal As String, rVal As String, lastVal As String
    Dim found As Boolean, shouldDelete As Boolean
    ' データファイルを参照
    Set dataWb = ThisWorkbook
    Set ws1 = dataWb.Sheets("test") ' ← シート名「test」
    lastRow1 = ws1.Cells(ws1.Rows.Count, "C").End(xlUp).Row
    
    ' planファイルを開く
    Set planWb = Workbooks.Open("C:\Users\t6520274\Desktop\resign_freedom\chan\Development\test_plan.xlsx")
    Set ws2 = planWb.Sheets("plan")
    lastRow2 = ws2.Cells(ws2.Rows.Count, "E").End(xlUp).Row ' E列で最終行を取得
    
    ' 削除対象行を記録する配列
    Dim rowsToDelete As New Collection
    On Error Resume Next ' コレクションキー重複エラーを無視
    
    For i = lastRow1 To 3 Step -1  ' 3行目から開始（ヘッダー行をスキップ）
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
        
        ' 削除対象なら配列に追加
        If shouldDelete = True Then
            Debug.Print ">>> 削除リストに追加：行 " & i & " → " & id1
            rowsToDelete.Add i, CStr(i)
        End If
    Next i
    
    ' まとめて削除（降順で削除することでインデックスの問題を回避）
    On Error GoTo 0
    Dim deleteRow As Variant
    For Each deleteRow In rowsToDelete
        Debug.Print ">>> 削除実行：行 " & deleteRow
        ws1.Rows(deleteRow).Delete
    Next deleteRow
    
    planWb.Close SaveChanges:=False
    MsgBox "処理が完了しました。data.xlsx の内容は plan.xlsx に従ってフィルタされました。"
End Sub