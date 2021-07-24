VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form__MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    updateMacro
    Me.TimerInterval = 600000  ' every minutes
End Sub

Sub Form_Timer()
     updateMacro
     exportMacro
End Sub



