VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SUB_CONTROL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command29_Click()

    If MsgBox("finish Printing?", vbYesNo) = vbYes Then
        DoCmd.RunSQL ("UPDATE [ORDER] SET [ORDER].STATE = 'POST' WHERE [ORDER].[STATE] = 'PRINT'")
    End If
End Sub

Private Sub Command30_Click()
 
    If MsgBox("Start Next Sessoin?", vbYesNo) = vbYes Then
         DoCmd.RunSQL ("UPDATE [ORDER] SET [ORDER].STATE = 'DESIGN' WHERE [ORDER].[STATE] = 'NEXT'")
    End If
    
End Sub

Private Sub Command32_Click()
        DoCmd.OpenForm "BOX_MANAGE", , , ""
End Sub

Private Sub Command54_Click()
      DoCmd.OpenForm "BOX_STATICS", , , ""
End Sub

Private Sub Command76_Click()
      DoCmd.OpenForm "BOX_CONTENT", , , ""
End Sub

Private Sub Command77_Click()
      DoCmd.OpenForm "BOX_SHEIKH", , , ""
End Sub

Private Sub Command78_Click()
      DoCmd.OpenForm "BOX_INSERT", , , ""
End Sub

Private Sub Command79_Click()
      DoCmd.OpenForm "BOX_POST", , , ""
End Sub

Private Sub Command80_Click()
    DoCmd.SelectObject acForm, "BOX_POST", True
    DoCmd.PrintOut acPrintAll
End Sub

Private Sub Command81_Click()
    
    If MsgBox("save Backup?", vbYesNo) = vbYes Then
        backupMacro
        MsgBox "Backup Done!"
    End If
     
End Sub

Private Sub Command91_Click()
    DoCmd.OpenForm "BOX_MANAGE", , , ""
End Sub

Private Sub Command92_Click()
    DoCmd.OpenForm "BOX_STATICS", , , ""
End Sub

Private Sub Command93_Click()
   Dim sh_name As String
   sh_name = InputBox("√œŒ· «”„ «·‘ÌŒ")
  
On Error GoTo checker:
   MkDir "E:\sheikh\" & sh_name
   MkDir "E:\sheikh\" & sh_name & "\template"
   FileCopy "E:\sanad\egaza.docx", "E:\sheikh\" & sh_name & "\template\egaza.docx"
   DoCmd.RunSQL ("INSERT INTO SHEIKH (NAME) VALUES ('" & sh_name & "')")
   Shell "C:\WINDOWS\explorer.exe """ & "E:\sheikh\" & sh_name & "\template", vbNormalFocus
    GoTo ender:
checker:
MsgBox "folder already exits"
ender:
End Sub
