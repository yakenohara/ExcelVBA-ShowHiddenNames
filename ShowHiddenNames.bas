Attribute VB_Name = "ShowHiddenNames"
'<License>------------------------------------------------------------
'
' Copyright (c) 2018 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

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
