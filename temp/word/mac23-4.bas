Attribute VB_Name = "NewMacros"
Sub Numbering()
Attribute Numbering.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.otor"
'
' Numbering Macro
'
'
    
    ' get the seprator
    Dim i As Integer
    Dim index As Integer
    
    index = 1
    
    For i = 1 To Selection.Words.Count
    
       If Trim(Selection.Words(i).Text) = ":" Then
           Selection.Words(i) = "(" + Trim(Str(index - 1)) + "): "
       End If
               
       If Trim(Selection.Words(i).Text) = "/" Then
           Selection.Words(i) = "(" + Trim(Str(index)) + ")/ "
           index = index + 1
       End If

    Next i
     
  ' make numbers subscript
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Superscript = False
        .Subscript = True
    End With
    With Selection.Find
        .Text = "\(([0-9]{1,2})\)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
End Sub
Sub Special_Numbering()
Attribute Special_Numbering.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Special_Numbering Macro
'
'
    
    ' get the seprator
    Dim i As Integer
    Dim index As Integer
    Dim Message, Title, Default, MyValue
    Message = "Enter a start"    ' Set prompt.
    Title = "Numbering"    ' Set title.
    Default = "1"    ' Set default.
    ' Display message, title, and default value.
    MyValue = InputBox(Message, Title, Default)
    index = CInt(MyValue)
    
    For i = 1 To Selection.Words.Count
    
       If Trim(Selection.Words(i).Text) = ":" Then
           Selection.Words(i) = "(" + Trim(Str(index - 1)) + "): "
       End If
               
       If Trim(Selection.Words(i).Text) = "/" Then
           Selection.Words(i) = "(" + Trim(Str(index)) + ")/ "
           index = index + 1
       End If

    Next i
     
  ' make numbers subscript
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Superscript = False
        .Subscript = True
    End With
    With Selection.Find
        .Text = "\(([0-9]{1,2})\)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
End Sub
Sub femalize()
'
' femalize Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "���������������� ��������"
        .Replacement.Text = "��������������� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Application.Keyboard (3073)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "���������������� ��������"
        .Replacement.Text = "��������������� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "��� ������ ���"
        .Replacement.Text = "��� ������� ���"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "��� ���� �� ������ ������ ������ ����������� ���������� ����������������"
        .Replacement.Text = _
            "��� ���� ��� ������ ������ ������ ����������� ����������� ����������������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "��������� �������� ��������� /"
        .Replacement.Text = "��������� �������� ��������� /"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "������ ������ ������� ��������� ���������"
        .Replacement.Text = "������ ������� ������� ��������� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Replacement.Text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Replacement.Text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "�� �� ������ ����� ����������� �� ��������� ��������� ������������ �������������"
        .Replacement.Text = _
            "�� �� ������ ����� ����������� �� ��������� ��������� ������������ �������������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "���������� ������� ������� ���������� ���� ���� �������� �������� ����������� ��������� �������� ����� ������ ������� ��� ����� ������"
        .Replacement.Text = _
            "���������� ������� ������� ���������� ���� ���� ������� ������� �� �������� �� ������ �������� ����� ������� ������� ��� ����� ������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            " ����� �� ���� ����� ����� ������ �������� ������������ �� ������������ ���������������� ������� ������� ����� ����� �������� ������ ������� ����������� ��� ������� ���������"
        .Replacement.Text = _
            " ������ �� ���� ����� ����� ������� �������� ������������ �� ������������ ���������������� ������� ������ ����� ����� ������� ������ �������� ���� �������� ��� �������� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
           "������ ��� ���������"
        .Replacement.Text = _
           "���� �� �����"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "������ ������ ������ ���������� ���������� ��� ��������� ��������� ������������ ��������������"
        .Replacement.Text = _
            "������ ������� ������ ���������� ���������� ��� ��������� ��������� ������������ ��������������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "����� ������ / "
        .Replacement.Text = "������ ������� / "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.LtrPara
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "����� ������ / "
        .Replacement.Text = "������ ������� / "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "����� �������� ������� �� ��������� "
        .Replacement.Text = "����� �������� ������� �� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "����� �������� ������� �� ���������"
        .Replacement.Text = "����� �������� ������� �� ��������� "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "���������� ������ ��� ������ �������� �� �������� �������� ���� ������ ����������� ����������� �� ��������� ����������� �� ����������"
        .Replacement.Text = _
            "��������� ������ ��� ������� �������� �� ������� �������� ���� ������ ����������� ����������� �� ��������� ����������� �� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "�������� ��������� ������� ���� ������� ����� ��������� ����� ��� �������� ��������� �������"
        .Replacement.Text = _
            "������� ��������� ������� ���� ������� ����� �������� ����� ��� �������� �������� �������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "�������� ��������� ������� ���� ������� ����� ��������� ����� ��� �������� ��������� �������"
        .Replacement.Text = _
            "������� ��������� ������� ���� ������� ����� �������� ����� ��� �������� �������� �������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "�������� �������� ���������"
        .Replacement.Text = "��������� �������� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "�� ������ ���� ���������� ����� ����������"
        .Replacement.Text = "�� ������ ���� ���������� ������ ����������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "��������� ���� ��������� ����� ��� ����������� ���� ��������"
        .Replacement.Text = _
            "��������� ������ ��������� ����� ��� ����������� ���� ��������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            " ���������� ����� ���������� ��������� ������������� ���� ������� ����������� ��� ����������� ������������ ������ ����������� ������ ������."
        .Replacement.Text = _
            " ���������� ����� ��������� ��������� ������������� ���� ������� ����������� ��� ����������� ������������ ������ ���������� ������ �������."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
         .Text = _
            "��� ��� ������ ���� ���"
        .Replacement.Text = _
            "��� ���� ������� ���� ���"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    MsgBox ("Femalization is done!")
End Sub
Sub PdfSaver()
Attribute PdfSaver.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Dim FileName As String
    FileName = Split(ActiveDocument.Name, ".", 2)(i)
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        ActiveDocument.Path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
    ChangeFileOpenDirectory ActiveDocument.Path
End Sub
Sub ardan()
Attribute ardan.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "����"
        .Replacement.Text = "����� �� ������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "� ��� ����� ������"
        .Replacement.Text = _
            "��� ��� ������ ���� ���: 1- .........................................."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "___"
        .Replacement.Text = "2- .........................................."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    FileName = Split(ActiveDocument.Name, ".", 2)(i)
    ChangeFileOpenDirectory ActiveDocument.Path
 
    ActiveDocument.SaveAs2 FileName:=FileName + " - ����" + ".docx", FileFormat:= _
     wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
     :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
     :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
     SaveAsAOCELetter:=False, CompatibilityMode:=14
    
   
    FileName = Split(ActiveDocument.Name, ".", 2)(i)
    ChangeFileOpenDirectory ActiveDocument.Path
    
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
    ActiveDocument.Path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
    OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
    wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
    IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
    wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
    True, UseISO19005_1:=False


End Sub
Sub Hyper()
Attribute Hyper.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Hyper Macro
'
'
    Dim FileName As String
    ' 1- the orignial one
    ' 2- save current file as pdf
    FileName = Split(ActiveDocument.Name, ".", 2)(i)
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        ActiveDocument.Path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
    ChangeFileOpenDirectory ActiveDocument.Path
    
    ' 3- convert file to female
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "���������������� ��������"
        .Replacement.Text = "��������������� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Application.Keyboard (3073)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "���������������� ��������"
        .Replacement.Text = "��������������� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "��� ������ ���"
        .Replacement.Text = "��� ������� ���"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "��� ���� �� ������ ������ ������ ����������� ���������� ����������������"
        .Replacement.Text = _
            "��� ���� ��� ������ ������ ������ ����������� ����������� ����������������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "��������� �������� ��������� /"
        .Replacement.Text = "��������� �������� ��������� /"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "������ ������ ������� ��������� ���������"
        .Replacement.Text = "������ ������� ������� ��������� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Replacement.Text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Replacement.Text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "�� �� ������ ����� ����������� �� ��������� ��������� ������������ �������������"
        .Replacement.Text = _
            "�� �� ������ ����� ����������� �� ��������� ��������� ������������ �������������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "���������� ������� ������� ���������� ���� ���� �������� �������� ����������� ��������� �������� ����� ������ ������� ��� ����� ������"
        .Replacement.Text = _
            "���������� ������� ������� ���������� ���� ���� ������� ������� �� �������� �� ������ �������� ����� ������� ������� ��� ����� ������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            " ����� �� ���� ����� ����� ������ �������� ������������ �� ������������ ���������������� ������� ������� ����� ����� �������� ������ ������� ����������� ��� ������� ���������"
        .Replacement.Text = _
            " ������ �� ���� ����� ����� ������� �������� ������������ �� ������������ ���������������� ������� ������ ����� ����� ������� ������ �������� ���� �������� ��� �������� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
           "������ ��� ���������"
        .Replacement.Text = _
           "���� �� �����"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "������ ������ ������ ���������� ���������� ��� ��������� ��������� ������������ ��������������"
        .Replacement.Text = _
            "������ ������� ������ ���������� ���������� ��� ��������� ��������� ������������ ��������������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "����� ������ / "
        .Replacement.Text = "������ ������� / "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.LtrPara
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "����� ������ / "
        .Replacement.Text = "������ ������� / "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "����� �������� ������� �� ��������� "
        .Replacement.Text = "����� �������� ������� �� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "����� �������� ������� �� ���������"
        .Replacement.Text = "����� �������� ������� �� ��������� "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "���������� ������ ��� ������ �������� �� �������� �������� ���� ������ ����������� ����������� �� ��������� ����������� �� ����������"
        .Replacement.Text = _
            "��������� ������ ��� ������� �������� �� ������� �������� ���� ������ ����������� ����������� �� ��������� ����������� �� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "�������� ��������� ������� ���� ������� ����� ��������� ����� ��� �������� ��������� �������"
        .Replacement.Text = _
            "������� ��������� ������� ���� ������� ����� �������� ����� ��� �������� �������� �������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "�������� ��������� ������� ���� ������� ����� ��������� ����� ��� �������� ��������� �������"
        .Replacement.Text = _
            "������� ��������� ������� ���� ������� ����� �������� ����� ��� �������� �������� �������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "�������� �������� ���������"
        .Replacement.Text = "��������� �������� ���������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "�� ������ ���� ���������� ����� ����������"
        .Replacement.Text = "�� ������ ���� ���������� ������ ����������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            "��������� ���� ��������� ����� ��� ����������� ���� ��������"
        .Replacement.Text = _
            "��������� ������ ��������� ����� ��� ����������� ���� ��������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = _
            " ���������� ����� ���������� ��������� ������������� ���� ������� ����������� ��� ����������� ������������ ������ ����������� ������ ������."
        .Replacement.Text = _
            " ���������� ����� ��������� ��������� ������������� ���� ������� ����������� ��� ����������� ������������ ������ ���������� ������ �������."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
         .Text = _
            "��� ��� ������ ���� ���"
        .Replacement.Text = _
            "��� ���� ������� ���� ���"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      
    ActiveDocument.SaveAs2 FileName:=FileName + " - ����" + ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=14
        
    ' 4- convert the current file to pdf
    
    FileName = Split(ActiveDocument.Name, ".", 2)(i)
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        ActiveDocument.Path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
    ChangeFileOpenDirectory ActiveDocument.Path
    
   ' 5- convert to ardan
   
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "����"
        .Replacement.Text = "����� �� ������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "� ��� ����� ������"
        .Replacement.Text = _
            "��� ��� ������ ���� ���: 1- .........................................."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "___"
        .Replacement.Text = "2- .........................................."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     ActiveDocument.SaveAs2 FileName:=FileName + " - ����" + ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=14
        
    '6- convert the current to pdf
    
     FileName = Split(ActiveDocument.Name, ".", 2)(i)
     ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        ActiveDocument.Path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
    ChangeFileOpenDirectory ActiveDocument.Path
    
End Sub
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Application.Keyboard (3073)
    Selection.MoveDown Unit:=wdScreen, Count:=1
    Selection.MoveDown Unit:=wdScreen, Count:=1
End Sub
Sub mode()
Attribute mode.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.mode"
'
' mode Macro
'
'
     If Options.ArabicNumeral = wdNumeralHindi Then
        Options.ArabicNumeral = wdNumeralContext
     Else
        Options.ArabicNumeral = wdNumeralHindi
     End If
    
End Sub

Sub smart_form()
Attribute smart_form.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.smart_form"
'
' smart_form Macro
'
'
Dim frm As New UserForm1
frm.Show

End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "mogez"
        .Replacement.Text = "����� ������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub MacroTemp()

Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject
clipboard.SetText "WWW"
clipboard.PutInClipboard
Dim strContents As String

Set clipboard = New MSForms.DataObject
clipboard.GetFromClipboard
strContents = clipboard.GetText

MsgBox (strContents)

End Sub

Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro3"
'
' inverse_numbering Macro
'
'
    ' get the seprator
    Dim i As Integer
    Dim index As Integer
    Dim Message, Title, Default, MyValue
    Message = "Enter a start"    ' Set prompt.
    Title = "Numbering"    ' Set title.
    Default = "1"    ' Set default.
    ' Display message, title, and default value.
    MyValue = InputBox(Message, Title, Default)
    index = CInt(MyValue)
    
    For i = 1 To Selection.Words.Count
    
       If Trim(Selection.Words(Selection.Words.Count - i).Text) = ":" Then
           Selection.Words(Selection.Words.Count - i) = "(" + Trim(Str(index - 1)) + "): "
       End If
               
       If Trim(Selection.Words(Selection.Words.Count - i).Text) = "/" Then
           Selection.Words(Selection.Words.Count - i) = "(" + Trim(Str(index)) + ")/ "
           index = index + 1
       End If

    Next i
     
  ' make numbers subscript
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Superscript = False
        .Subscript = True
    End With
    With Selection.Find
        .Text = "\(([0-9]{1,2})\)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub