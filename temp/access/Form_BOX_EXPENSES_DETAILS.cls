VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_BOX_EXPENSES_DETAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim current_type As String

Public Sub change_type(new_type As String)
    If current_type <> new_type Then
        Form_BOX_EXPENSES_DETAILS.RecordSource = "SELECT * FROM EXPENSES WHERE CATEGORY.VALUE = '" & new_type & "' ORDER BY AMOUNT DESC"
        current_type = new_type
    End If
End Sub


Private Sub Form_AfterUpdate()
     Form_BOX_EXPENSES_DETAILS.Requery
     Form_BOX_EXPENSES_SUM.Requery
End Sub

Private Sub Form_Load()
    Form_BOX_EXPENSES_DETAILS.change_type (Form_BOX_EXPENSES_SUM.CATEGORY.Value)
    Form_BOX_EXPENSES_DETAILS.Requery
End Sub
