Attribute VB_Name = "dataModule"
Option Compare Database
Function updateMacro()
   
     DoCmd.OpenQuery "update_contents", acViewNormal, acEdit
     DoCmd.Save
     DoCmd.Requery
    
End Function

Function updateExpenses(income As Integer, covers As Integer, panar As Integer, printer As Integer, STAMP As Integer)
     
     DoCmd.RunSQL ("UPDATE EXPENSES SET AMOUNT = " & income & " WHERE EXPENSE = 'otor_income';")
     DoCmd.RunSQL ("UPDATE EXPENSES SET AMOUNT = " & covers & " WHERE EXPENSE = 'otor_covers'")
     DoCmd.RunSQL ("UPDATE EXPENSES SET AMOUNT = " & printer & " WHERE EXPENSE = 'otor_print'")
     DoCmd.RunSQL ("UPDATE EXPENSES SET AMOUNT = " & panar & " WHERE EXPENSE = 'otor_panar'")
     DoCmd.RunSQL ("UPDATE EXPENSES SET AMOUNT = " & STAMP & " WHERE EXPENSE = 'otor_stamp'")
   
End Function
Function backupMacro()

    DoCmd.Save
    Dim xlobj As Object
    Set xlobj = CreateObject("Scripting.FileSystemObject")
    xlobj.CopyFile "E:\sanad\otor.accdb", "E:\other\project\last\backup.accdb", True
    xlobj.CopyFile "E:\sanad\otor_be.accdb", "E:\other\project\last\backup_be.accdb", True
    Set xlobj = Nothing
    
    On Error GoTo missed:
    Set xlobj = CreateObject("Scripting.FileSystemObject")
    xlobj.CopyFile "E:\sanad\otor.accdb", "\\MAHMOUD-PC\Users\Public\otor.accdb", True
    xlobj.CopyFile "E:\sanad\otor_be.accdb", "\\MAHMOUD-PC\Users\Public\otor_be.accdb", True
    Set xlobj = Nothing
    
missed:
    
End Function


Function exportMacro()

    'DoCmd.OutputTo acOutputTable, "ORDER", acFormatXLSX, "E:\other\project\order.xlsx", False, "", , acExportQualityScreen
    'DoCmd.OutputTo acOutputTable, "ORDER", "HTML(*.html)", "\\MAHMOUD-PC\Users\Public\WebApp\data\order.html", False, "", , acExportQualityScreen
    'DoCmd.OutputTo acOutputTable, "SHEIKH", "HTML(*.html)", "\\MAHMOUD-PC\Users\Public\WebApp\data\sheikh.html", False, "", , acExportQualityScreen
    'DoCmd.OutputTo acOutputQuery, "MANAGE", "HTML(*.html)", "\\MAHMOUD-PC\Users\Public\WebApp\data\manage.html", False, "", , acExportQualityScreen
   
End Function
