Attribute VB_Name = "ShowHiddenNames"
Sub ShowHiddenNames()
    
    '�ϐ��錾
    Dim nameDef As Object
    Dim numOfInvisibleNames As Long: numOfInvisibleNames = 0
    
    '����
    For Each nameDef In Names
        If Not (nameDef.Visible) Then
            numOfInvisibleNames = numOfInvisibleNames + 1
            nameDef.Visible = True
        End If
    Next
    
    MsgBox "Done!" & vbLf & vbLf & _
           "�\���ς݃I�u�W�F�N�g���܂߁A" & Format(Names.count, "#,###;-#,###;0") & " ���������܂���" & vbLf & _
           "���A" & Format(numOfInvisibleNames, "#,###;-#,###;0") & "�����\����Ԃ���\����ԂɕύX���܂���"
    
End Sub
