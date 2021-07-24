VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_BOX_EXPENSES_SUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CATEGORY_DblClick(Cancel As Integer)

 If CATEGORY.Value = "DEPTS" Then
    Me.Recordset.MoveFirst
 Else
    Me.Recordset.MoveNext
 End If

 Refresh
 Form_BOX_EXPENSES_DETAILS.change_type (CATEGORY.Value)

End Sub

Private Sub CATEGORY_KeyUp(KeyCode As Integer, Shift As Integer)
 Form_BOX_EXPENSES_DETAILS.change_type (CATEGORY.Value)
End Sub
