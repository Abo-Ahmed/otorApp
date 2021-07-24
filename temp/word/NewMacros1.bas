Attribute VB_Name = "NewMacros1"
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
    
       If Trim(Selection.Words(i).text) = ":" Then
           Selection.Words(i) = "(" + Trim(Str(index - 1)) + "): "
       End If
               
       If Trim(Selection.Words(i).text) = "/" Then
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
        .text = "\(([0-9]{1,2})\)"
        .Replacement.text = ""
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
    
       If Trim(Selection.Words(i).text) = ":" Then
           Selection.Words(i) = "(" + Trim(Str(index - 1)) + "): "
       End If
               
       If Trim(Selection.Words(i).text) = "/" Then
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
        .text = "\(([0-9]{1,2})\)"
        .Replacement.text = ""
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
        .text = "���������������� ��������"
        .Replacement.text = "��������������� ���������"
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
        .text = "���������������� ��������"
        .Replacement.text = "��������������� ���������"
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
        .text = "��� ������ ���"
        .Replacement.text = "��� ������� ���"
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
        .text = _
            "��� ���� �� ������ ������ ������ ����������� ���������� ����������������"
        .Replacement.text = _
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
        .text = "��������� �������� ��������� /"
        .Replacement.text = "��������� �������� ��������� /"
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
        .text = "������ ������ ������� ��������� ���������"
        .Replacement.text = "������ ������� ������� ��������� ���������"
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
        .text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Replacement.text = _
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
        .text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Replacement.text = _
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
        .text = _
            "�� �� ������ ����� ����������� �� ��������� ��������� ������������ �������������"
        .Replacement.text = _
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
        .text = _
            "���������� ������� ������� ���������� ���� ���� �������� �������� ����������� ��������� �������� ����� ������ ������� ��� ����� ������"
        .Replacement.text = _
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
        .text = _
            " ����� �� ���� ����� ����� ������ �������� ������������ �� ������������ ���������������� ������� ������� ����� ����� �������� ������ ������� ����������� ��� ������� ���������"
        .Replacement.text = _
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
        .text = _
           "������ ��� ���������"
        .Replacement.text = _
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
        .text = _
            "������ ������ ������ ���������� ���������� ��� ��������� ��������� ������������ ��������������"
        .Replacement.text = _
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
        .text = "����� ������ / "
        .Replacement.text = "������ ������� / "
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
        .text = "����� ������ / "
        .Replacement.text = "������ ������� / "
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
        .text = "����� �������� ������� �� ��������� "
        .Replacement.text = "����� �������� ������� �� ���������"
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
        .text = "����� �������� ������� �� ���������"
        .Replacement.text = "����� �������� ������� �� ��������� "
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
        .text = _
            "���������� ������ ��� ������ �������� �� �������� �������� ���� ������ ����������� ����������� �� ��������� ����������� �� ����������"
        .Replacement.text = _
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
        .text = _
            "�������� ��������� ������� ���� ������� ����� ��������� ����� ��� �������� ��������� �������"
        .Replacement.text = _
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
        .text = _
            "�������� ��������� ������� ���� ������� ����� ��������� ����� ��� �������� ��������� �������"
        .Replacement.text = _
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
        .text = "�������� �������� ���������"
        .Replacement.text = "��������� �������� ���������"
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
        .text = "�� ������ ���� ���������� ����� ����������"
        .Replacement.text = "�� ������ ���� ���������� ������ ����������"
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
        .text = _
            "��������� ���� ��������� ����� ��� ����������� ���� ��������"
        .Replacement.text = _
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
        .text = _
            " ���������� ����� ���������� ��������� ������������� ���� ������� ����������� ��� ����������� ������������ ������ ����������� ������ ������."
        .Replacement.text = _
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
         .text = _
            "��� ��� ������ ���� ���"
        .Replacement.text = _
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
Sub malize()
'
' malize Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "��������������� ���������"
        .Replacement.text = "���������������� ��������"
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
        .text = "��������������� ���������"
        .Replacement.text = "���������������� ��������"
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
        .text = "��� ������� ���"
        .Replacement.text = "��� ������ ���"
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
        .text = "��� ���� ��� ������ ������ ������ ����������� ����������� ����������������"
        .Replacement.text = "��� ���� �� ������ ������ ������ ����������� ���������� ����������������"
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
        .text = "��������� �������� ��������� /"
        .Replacement.text = "��������� �������� ��������� /"
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
        .text = "������ ������� ������� ��������� ���������"
        .Replacement.text = "������ ������ ������� ��������� ���������"
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
        .text = "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Replacement.text = "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
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
        .text = "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Replacement.text = "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
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
        .text = "�� �� ������ ����� ����������� �� ��������� ��������� ������������ �������������"
        .Replacement.text = "�� �� ������ ����� ����������� �� ��������� ��������� ������������ �������������"
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
        .text = "���������� ������� ������� ���������� ���� ���� ������� ������� �� �������� �� ������ �������� ����� ������� ������� ��� ����� ������"
        .Replacement.text = "���������� ������� ������� ���������� ���� ���� �������� �������� ����������� ��������� �������� ����� ������ ������� ��� ����� ������"
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
        .text = " ������ �� ���� ����� ����� ������� �������� ������������ �� ������������ ���������������� ������� ������ ����� ����� ������� ������ �������� ���� �������� ��� �������� ���������"
         .Replacement.text = " ����� �� ���� ����� ����� ������ �������� ������������ �� ������������ ���������������� ������� ������� ����� ����� �������� ������ ������� ����������� ��� ������� ���������"
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
        .text = "���� �� �����"
        .Replacement.text = "������ ��� ���������"
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
        .text = "������ ������� ������ ���������� ���������� ��� ��������� ��������� ������������ ��������������"
        .Replacement.text = "������ ������ ������ ���������� ���������� ��� ��������� ��������� ������������ ��������������"
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
        .text = "������ ������� / "
        .Replacement.text = "����� ������ / "
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
        .text = "������ ������� / "
        .Replacement.text = "����� ������ / "
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
        .text = "����� �������� ������� �� ���������"
        .Replacement.text = "����� �������� ������� �� ��������� "
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
        .text = "����� �������� ������� �� ��������� "
        .Replacement.text = "����� �������� ������� �� ���������"
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
        .text = "��������� ������ ��� ������� �������� �� ������� �������� ���� ������ ����������� ����������� �� ��������� ����������� �� ���������"
        .Replacement.text = "���������� ������ ��� ������ �������� �� �������� �������� ���� ������ ����������� ����������� �� ��������� ����������� �� ����������"
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
        .text = "������� ��������� ������� ���� ������� ����� �������� ����� ��� �������� �������� �������"
        .Replacement.text = "�������� ��������� ������� ���� ������� ����� ��������� ����� ��� �������� ��������� �������"
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
        .text = "������� ��������� ������� ���� ������� ����� �������� ����� ��� �������� �������� �������"
        .Replacement.text = "�������� ��������� ������� ���� ������� ����� ��������� ����� ��� �������� ��������� �������"
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
        .text = "��������� �������� ���������"
        .Replacement.text = "�������� �������� ���������"
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
        .text = "�� ������ ���� ���������� ������ ����������"
        .Replacement.text = "�� ������ ���� ���������� ����� ����������"
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
        .text = "��������� ������ ��������� ����� ��� ����������� ���� ��������"
        .Replacement.text = "��������� ���� ��������� ����� ��� ����������� ���� ��������"
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
        .text = " ���������� ����� ��������� ��������� ������������� ���� ������� ����������� ��� ����������� ������������ ������ ���������� ������ �������."
        .Replacement.text = " ���������� ����� ���������� ��������� ������������� ���� ������� ����������� ��� ����������� ������������ ������ ����������� ������ ������."
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
         .text = "��� ���� ������� ���� ���"
         .Replacement.text = "��� ��� ������ ���� ���"
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
    MsgBox ("malization is done!")
End Sub
Sub PdfSaver()
Attribute PdfSaver.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Dim FileName As String
    FileName = Split(ActiveDocument.NAME, ".", 2)(i)
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        ActiveDocument.path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
    ChangeFileOpenDirectory ActiveDocument.path
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
        .text = "����"
        .Replacement.text = "����� �� ������"
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
        .text = "� ��� ����� ������"
        .Replacement.text = _
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
        .text = "___"
        .Replacement.text = "2- .........................................."
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
    
    FileName = Split(ActiveDocument.NAME, ".", 2)(i)
    ChangeFileOpenDirectory ActiveDocument.path
 
    ActiveDocument.SaveAs2 FileName:=FileName + " - ����" + ".docx", FileFormat:= _
     wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
     :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
     :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
     SaveAsAOCELetter:=False, CompatibilityMode:=14
    
   
    FileName = Split(ActiveDocument.NAME, ".", 2)(i)
    ChangeFileOpenDirectory ActiveDocument.path
    
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
    ActiveDocument.path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
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
    FileName = Split(ActiveDocument.NAME, ".", 2)(i)
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        ActiveDocument.path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
    ChangeFileOpenDirectory ActiveDocument.path
    
    ' 3- convert file to female
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "���������������� ��������"
        .Replacement.text = "��������������� ���������"
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
        .text = "���������������� ��������"
        .Replacement.text = "��������������� ���������"
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
        .text = "��� ������ ���"
        .Replacement.text = "��� ������� ���"
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
        .text = _
            "��� ���� �� ������ ������ ������ ����������� ���������� ����������������"
        .Replacement.text = _
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
        .text = "��������� �������� ��������� /"
        .Replacement.text = "��������� �������� ��������� /"
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
        .text = "������ ������ ������� ��������� ���������"
        .Replacement.text = "������ ������� ������� ��������� ���������"
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
        .text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Replacement.text = _
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
        .text = _
            "������� ���� �������� ������ ������������ ������������� �� ������������� ����� ����������� ����� ��������������"
        .Replacement.text = _
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
        .text = _
            "�� �� ������ ����� ����������� �� ��������� ��������� ������������ �������������"
        .Replacement.text = _
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
        .text = _
            "���������� ������� ������� ���������� ���� ���� �������� �������� ����������� ��������� �������� ����� ������ ������� ��� ����� ������"
        .Replacement.text = _
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
        .text = _
            " ����� �� ���� ����� ����� ������ �������� ������������ �� ������������ ���������������� ������� ������� ����� ����� �������� ������ ������� ����������� ��� ������� ���������"
        .Replacement.text = _
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
        .text = _
           "������ ��� ���������"
        .Replacement.text = _
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
        .text = _
            "������ ������ ������ ���������� ���������� ��� ��������� ��������� ������������ ��������������"
        .Replacement.text = _
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
        .text = "����� ������ / "
        .Replacement.text = "������ ������� / "
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
        .text = "����� ������ / "
        .Replacement.text = "������ ������� / "
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
        .text = "����� �������� ������� �� ��������� "
        .Replacement.text = "����� �������� ������� �� ���������"
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
        .text = "����� �������� ������� �� ���������"
        .Replacement.text = "����� �������� ������� �� ��������� "
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
        .text = _
            "���������� ������ ��� ������ �������� �� �������� �������� ���� ������ ����������� ����������� �� ��������� ����������� �� ����������"
        .Replacement.text = _
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
        .text = _
            "�������� ��������� ������� ���� ������� ����� ��������� ����� ��� �������� ��������� �������"
        .Replacement.text = _
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
        .text = _
            "�������� ��������� ������� ���� ������� ����� ��������� ����� ��� �������� ��������� �������"
        .Replacement.text = _
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
        .text = "�������� �������� ���������"
        .Replacement.text = "��������� �������� ���������"
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
        .text = "�� ������ ���� ���������� ����� ����������"
        .Replacement.text = "�� ������ ���� ���������� ������ ����������"
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
        .text = _
            "��������� ���� ��������� ����� ��� ����������� ���� ��������"
        .Replacement.text = _
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
        .text = _
            " ���������� ����� ���������� ��������� ������������� ���� ������� ����������� ��� ����������� ������������ ������ ����������� ������ ������."
        .Replacement.text = _
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
         .text = _
            "��� ��� ������ ���� ���"
        .Replacement.text = _
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
    
    FileName = Split(ActiveDocument.NAME, ".", 2)(i)
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        ActiveDocument.path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
    ChangeFileOpenDirectory ActiveDocument.path
    
   ' 5- convert to ardan
   
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "����"
        .Replacement.text = "����� �� ������"
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
        .text = "� ��� ����� ������"
        .Replacement.text = _
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
        .text = "___"
        .Replacement.text = "2- .........................................."
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
    
     FileName = Split(ActiveDocument.NAME, ".", 2)(i)
     ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        ActiveDocument.path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
    ChangeFileOpenDirectory ActiveDocument.path
    
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
        .text = "mogez"
        .Replacement.text = "����� ������"
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
Sub MacroBulkPdf()
 
 Dim fd As FileDialog
 Set fd = Application.FileDialog(msoFileDialogFilePicker)
 
 Dim vrtSelectedItem As Variant
 Dim wdApp As Word.Application
 Set wdApp = GetObject(, "Word.Application")
      
 With fd
 
 .AllowMultiSelect = True
 
 If .Show = -1 Then
 
     For Each vrtSelectedItem In .SelectedItems
     
        Documents.Open FileName:=vrtSelectedItem, ReadOnly:=False
  
        ActiveDocument.ExportAsFixedFormat OutputFileName:= _
          Split(vrtSelectedItem, ".", 2)(i) + ".pdf", ExportFormat:=wdExportFormatPDF, _
           OpenAfterExport:=True, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
           wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
           IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
           wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
           True, UseISO19005_1:=False
           
        ActiveDocument.Save
        wdApp.Documents(Split(vrtSelectedItem, ".", 2)(i) + ".docx").Close
       
     Next
 'If the user presses Cancel...
 Else
 End If
 End With
 
 'Set the object variable to Nothing.
 Set fd = Nothing

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
    
       If Trim(Selection.Words(Selection.Words.Count - i).text) = ":" Then
           Selection.Words(Selection.Words.Count - i) = "(" + Trim(Str(index - 1)) + "): "
       End If
               
       If Trim(Selection.Words(Selection.Words.Count - i).text) = "/" Then
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
        .text = "\(([0-9]{1,2})\)"
        .Replacement.text = ""
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
