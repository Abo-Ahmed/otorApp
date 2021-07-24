VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FORM_ORDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim current_order As Integer
Dim current_state As String

Public Sub change_state(new_state As String)
    If current_state <> new_state Then
        If InStr(new_state, "CURRENT") > 0 Then
            Form_SUB_MANAGE.Requery
            Form_FORM_ORDER.RecordSource = "SELECT * FROM [ORDER] ORDER BY DEGREE, REST DESC"
            updateExpenses income:=Form_SUB_MANAGE.REST.Value, covers:=Form_SUB_MANAGE.EJAZA.Value * 8, panar:=Form_SUB_MANAGE.TREE.Value * 30, printer:=Form_SUB_MANAGE.EJAZA.Value * 20, STAMP:=Form_SUB_MANAGE.STAMP.Value * 100
        Else
            Form_FORM_ORDER.RecordSource = "SELECT * FROM [ORDER] WHERE STATE = '" & new_state & "' ORDER BY DEGREE, REST DESC"
        End If
        
        current_state = new_state
        change_order
    End If
End Sub
Public Sub change_order()
    If ID <> current_order Then
        current_order = ID
     '   Form_SUB_CONTENT.RecordSource = "SELECT * FROM CONTENT WHERE ORDER_ID = " & current_order
     '   Form_SUB_ORDER.RecordSource = "SELECT * FROM [ORDER] WHERE ID = " & current_order
     '   Form_SUB_SHEIKH.RecordSource = "SELECT * FROM SHEIKH WHERE ID = (SELECT SHEIKH_ID FROM [ORDER] WHERE [ORDER].ID = " & current_order & ")"
        
    End If
End Sub

Private Sub Command31_Click()
      change_order
      DoCmd.OpenForm "SUB_DETAILS", , , ""
      Form_SUB_DETAILS.RecordSource = "SELECT [ORDER].[STATE], [ORDER].[ID] AS ORDER_ID, [ORDER].[SHEIKH_ID], [ORDER].[SHEIKH_NAME], [ORDER].[COMMENT] AS ORDER_COMMENT, [ORDER].[CONTENTS], [ORDER].[COST], [ORDER].[PAID], [ORDER].[REST], [ORDER].[P_RECEIVER], [ORDER].[P_PHONE], [ORDER].[P_COUNTRY], [ORDER].[P_CITY], [ORDER].[P_ADDRESS], [ORDER].[INSERT_DATE], [ORDER].[UPDATE_DATE], [SHEIKH].[ID] AS SHEIKH_ID, [SHEIKH].[NAME], [SHEIKH].[INFO], [SHEIKH].[COMMENT] AS SHEIKH_COMMENT, [SHEIKH].[GENDER], [SHEIKH].[RECEIVER_NAME], [SHEIKH].[PHONE], [SHEIKH].[COUNTRY], [SHEIKH].[CITY], [SHEIKH].[ADDRESS] FROM SHEIKH INNER JOIN [ORDER] ON [SHEIKH].[ID] =[ORDER].[SHEIKH_ID] WHERE [ORDER].ID = " & current_order
      Form_SUB_DETAILS.Requery
End Sub

Private Sub REST_DblClick(Cancel As Integer)
End Sub

Private Sub SHEIKH_NAME_AfterUpdate()
 DoCmd.Save
 Refresh
 Form_FORM_ORDER.Requery
 DoCmd.OpenQuery "update_sheikh", acViewNormal, acEdit

End Sub


Private Sub STATE_Change()
    If STATE.Value = "DONE" Then
       UPDATE_DATE = Date
       DoCmd.Save
       DoCmd.RunSQL ("INSERT INTO ORDER_HISTORY SELECT * FROM [ORDER] WHERE [ORDER].STATE = 'DONE'")
       DoCmd.RunSQL ("DELETE * FROM [ORDER] WHERE STATE = 'DONE'")
    End If
    On Error GoTo finish:
    DoCmd.RefreshRecord
finish:
    DoCmd.Save
    DoCmd.Requery
End Sub

Private Sub Form_AfterInsert()
    DoCmd.RunSQL ("INSERT INTO CONTENT (TYPE , ORDER_ID) VALUES ('OTHER' , " & ID & ")")
    'Form_SUB_CONTENT.Requery
    'Form_SUB_SHEIKH.Requery
    'Form_SUB_ORDER.Requery
End Sub

Private Sub PAID_AfterUpdate()
    DoCmd.Save
    DoCmd.Requery
End Sub
Private Sub Command21_Click()
  change_order
  Shell "C:\WINDOWS\explorer.exe """ & "E:\sheikh" & "\" & SHEIKH_NAME.Value & "", vbNormalFocus
End Sub

Private Sub Text27_Click()
    PAID.SetFocus
    change_order
End Sub

Private Sub STATE_Click()
    change_order
End Sub

Private Sub STATE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    change_order
End Sub

Private Sub Text29_Click()
    change_order
End Sub

Private Sub SHEIKH_NAME_Click()
    change_order
End Sub

Private Sub SHEIKH_NAME_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    change_order
End Sub

Private Sub PAID_Click()
    change_order
End Sub

Private Sub REST_Click()
    change_order
End Sub

Private Sub COMMENT_Click()
    change_order
End Sub

Private Sub CONTENTS_Click()
    change_order
End Sub

Private Sub COST_Click()
    change_order
End Sub

Private Sub Detail_Click()
    change_order
End Sub

Private Sub Form_Click()
    change_order
End Sub

Private Sub Command28_Click()
    change_order
End Sub
