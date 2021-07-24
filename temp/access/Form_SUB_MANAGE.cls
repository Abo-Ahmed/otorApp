VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SUB_MANAGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub PROCESS_DblClick(Cancel As Integer)
 If PROCESS.Value = "DESIGN" Then
    Me.Recordset.MoveFirst
 Else
    Me.Recordset.MoveNext
 End If
 Refresh
 Form_FORM_ORDER.change_state (PROCESS.Value)
End Sub

Private Sub PROCESS_KeyUp(KeyCode As Integer, Shift As Integer)
 
 Form_FORM_ORDER.change_state (PROCESS.Value)

End Sub

Private Sub Text27_Click()
    PAID.SetFocus
End Sub
