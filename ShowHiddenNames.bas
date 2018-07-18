Attribute VB_Name = "ShowHiddenNames"
Sub ShowHiddenNames()
    
    '変数宣言
    Dim nameDef As Object
    Dim numOfInvisibleNames As Long: numOfInvisibleNames = 0
    
    '処理
    For Each nameDef In Names
        If Not (nameDef.Visible) Then
            numOfInvisibleNames = numOfInvisibleNames + 1
            nameDef.Visible = True
        End If
    Next
    
    MsgBox "Done!" & vbLf & vbLf & _
           "表示済みオブジェクトを含め、" & Format(Names.count, "#,###;-#,###;0") & " 件処理しました" & vbLf & _
           "内、" & Format(numOfInvisibleNames, "#,###;-#,###;0") & "件を非表示状態から表示状態に変更しました"
    
End Sub
