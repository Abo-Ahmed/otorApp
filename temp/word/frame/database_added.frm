VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�����"
   ClientHeight    =   7344
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11568
   OleObjectBlob   =   "database_added.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub select_obydi(x)
        
        If x = 1 Then
            
            iPage = 11
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
    
            iPage = 10
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
            
            iPage = 9
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
                       
        ElseIf x = 2 Then
            
            iPage = 11
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
            
            iPage = 6
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
            
            iPage = 6
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
            
            iPage = 6
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
        
        ElseIf x = 3 Then
            
            iPage = 6
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
            
            iPage = 6
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
            
            iPage = 6
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
            
            iPage = 6
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
            
            iPage = 6
            With ActiveDocument
              Set Rng = .GoTo(What:=wdGoToPage, NAME:=iPage)
              Set Rng = Rng.GoTo(What:=wdGoToBookmark, NAME:="\page")
              Rng.Delete
            End With
        End If
End Sub
Private Sub set_sheikh_and_student(sheikh_name, sheikh_info, student_name, student_info)
    
    ' change sheikh name
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "sheikh_name"
        .Replacement.text = sheikh_name
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

    ' change student name
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "student_name"
        .Replacement.text = student_name
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

    ' set student info
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "student_info"
        .Replacement.text = student_info
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

    ' set sheikh info
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "sheikh_info"
        .Replacement.text = sheikh_info
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
Private Sub set_types(sheikh_type, student_type)
    ' set sheikh type
    If sheikh_type = 1 Then
    
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "mogez"
            .Replacement.text = "������"
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
            .text = "����� ������"
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
    Else
        
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .text = "mogez"
            .Replacement.text = "�����"
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

    End If
    

' set student type
    If student_type = False Then
            
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
    End If

End Sub
Private Sub set_qeraat(STATE, qeraat, rawy)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "egaza_content"
        .Replacement.text = qeraat
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
        .text = "rawy"
        .Replacement.text = rawy
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
        .text = "egaza_state"
        .Replacement.text = STATE
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
Private Sub set_snada(sanada)
 
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.SetText sanada
    clipboard.PutInClipboard
    Dim strContents As String

    Dim target As String
    Dim rngtarget As Range
    target = "sanada"
    Selection.HomeKey wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
    Do While .Execute(FindText:=target, Forward:=True, _
    MatchWildcards:=False, Wrap:=wdFindStop, MatchCase:=False) = True
    Selection.Paste
    Selection.Collapse wdCollapseEnd
    Selection.MoveRight wdCharacter, 1
    Loop
    
    End With
  
End Sub
Function sanadan(index As Integer) As String
     
     'adding sanad
     If index = -1 Then
         '��� ����
         sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
         sanadan = sanadan & "���� ����� ���� : ������ ��� ���� �� ���� ���: ����� ��� ����� � ��� : ����� ������ �� ����� ������ � ��� :����� ���� �� ���� �������� � ��� : ����� ���� �� ���� � ���: ����� ���� �� ���� ����� � ��� :���� ��� ���� �� ������ ������� � ����: ���� ��� ��� ���� �� ���� � ��� : ��� ���� : ����� ��� ������� ��� ��� ��� ����� ����� � ����: ���� ��� ��� ��� ���� �� ������ ������� � � ��� : ���� ��� ��� ���� �� ���� �� ����� � ���� : ���� ��� �������� � ���� : ���� ��� ���� " & vbNewLine
         sanadan = sanadan & "���� ����� ��� ����� : ������ ��� ���� �� ���� � ���:����� ���� �� ���� �� ����� � ��� : ����� ���� �� ���� ������� � ��� : ����� ��� ���� �� ����� � ��� : ����� ���� �� ���� ������� � ��� :����� ���� �� ������ ������� � ��� : ���� ��� ��� ���� � ���� ��� ���� : ����� ��� ������ ��� ��� ��� ������ �� ���� ������� ������� ���� �� : ���� ��� ��� ��� ��� ���� �� ����� ������ � ���� : ���� ��� ��� ��� ��� ���� ����� �� ���� �� ���� ������ ������ ������ �� ��� ���� �� ����� " & vbNewLine
         sanadan = sanadan & "������� ��� ���� ������� ������� : ��� ������� ����� �� ���� ���� ���� ���� � �������� �� ��� ���� �������� � ����� ��� ������� ���� ����� . ���� ������� �� ����� �� ���� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -" & vbNewLine
         snandan = sanadan & vbNewLine
         '����
         sanadan = sanadan & "��� ����� ������ ��������� ����� ��������:" & vbNewLine
         sanadan = sanadan & "���� ����� ��� ��� ����: ������ ��� ���� �� ���� �� ��� ������ ���: ����� �� ����� ���: ����� ������� �� ���� �� ��� ������� � ���:����� ��� ���:����� ���� �� ��� � ���: ����� ��� ��� �� ���� � ��� ��� ����: ����� ��� ������� ��� ��� ���� �� ���� ������� � � ��� ��: ���� ��� ��� ��� ����� ��� ������ �� ����� ������� � ����: ���� ��� ������� �� ��� ������ �� ���� ������� �������� ����: ���� ��� ���� �� ����� ������� � ����: ���� ��� ���� �� ���� ��������� � ����: ���� ��� ��� ���� �� ��� �� ��� ��� �� ����." & vbNewLine
         sanadan = sanadan & "���� ����� ��� : ������ ��� ��� ����� ����� �� ����� ������ � ��� : ����� ��� ����� ��� �� ���� �� ���� ������� ������ ������ ������� � ���: ����� ��� ������ ���� �� ��� �������� � ����: ���� ��� ��� ���� ���� �� ������ � ����: ���� ��� ��� � ����: ���� ��� ����� � ���� ��� ���� : ����� ��� ������ ��� ��� ����� ��� ����� ���� ��: ���� ��� ��� ������� ����: ���� ��� �������� �� ���� �� ��� �� ����� . " & vbNewLine
         sanadan = sanadan & "����� ���� ������� ����� ����� : ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ����  � ���� �� ���  � ���� �� ����  � ���� ���� �� ����� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� - � ��� �� �� ���� �� ����� �� �������  � ���� �����  � �� ���� ���� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
         snandan = sanadan & vbNewLine
        
         '�������
         sanadan = sanadan & "��� ��� ���� ������� �� �������:" & vbNewLine
         sanadan = sanadan & "���� ����� ��� ���� ������ : ������ ��� ��� ���� ��� ������ �� ��� �� ���� ������ � ��� : ����� ��� ��� ��� ���� �� ���� �� ������ ������� � ��� : ����� ���� �� ���� �� ��� ������� � ��� : ����� ��� ��� ������ � �� ������� � � ��� ��� ����� : ����� ��� ������� ��� ��� ��� ����� � ���� �� : ���� ��� ��� ��� ������ �� ����� � ���� : ���� ��� ��� ��� ���� �� ��� �� ������� ������� � � ��� :���� ��� ���� �� ���� � ���� : ���� ��� ��� ��� ������ � ���� : ���� ��� ������� ." & vbNewLine
         sanadan = sanadan & "���� ����� ��� ������ : ������ ��� ���� �� ���� � ��� : ����� ��� ��� ����� � ��� : ����� ���� �� ���� ( ������� ������) � �� ��� ������ � �� ������� � � ��� ����� ������ : ����� ��� ������� ��� ��� ���� �� ���� � � ��� �� : ���� ��� ��� ��� ����� ��� ������ �� ����� ������� � ���� : ���� ��� ��� ��� ������ ��� �� ��� � ���� : ���� ��� ���� �� ����� ������� ������ � ���� :���� ��� ���� �� ���� ( ������� ������) � ���� : ���� ��� ��� ������ � ���� : ���� ��� ������� ." & vbNewLine
         sanadan = sanadan & "����� ������� : ���� �� ���� ������ � ����� �� ��� �������� � ����� �� ��� ���� ������ � ������ �� ����� �������� ��� �� ���� ������ �������� �� ������� �� ���� � ��� ����� ����� ������ ." & vbNewLine
         sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
         sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
         snandan = sanadan & vbNewLine
        
         '���
         sanadan = sanadan & "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
         sanadan = sanadan & "��� ����� ����� ������ : ������ ��� ��� ��� ��� �� ����� ������� ���� ���� ���� � �� ���� ������ ������ ��� ������ ���� �� ������� �� ��� �������� ������� � ��� : ������ ����� � ��� : ������ ��� �������� ������ �� ����� ������� � ������ ��� ���� ���� �� ������ ������� � ������ ��� ������ ���� �� ��� ���� �� ����� ���������� � ������ ��� ����� ���� �� ��� ���� �� ���� �� ��� ������ ������� ���� ��� ��� ������ � ������ ��� ����� ����� �� ������� ������ ." & vbNewLine
         sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� �� �� ������� ��� ��� ���� ������ � ���� ���� ������� �������� � ���� �� ����� ��� ��� ��� ���� ���� �� ���� �� ��� ������ ������ � ���� ��� ��� ������ �� ���� � ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ������ ��� ���� �� ���� �� ����� �������� � ���� ��� ��� ��� ��� ���� �� ��� �� ���� ������ � ���� ��� ��� ��� ������ ���������� � ���� ��� ��� ��� ��� ��� ������ � ���� ��� ��� ����� ������ � ���� ��� ��� ��� ." & vbNewLine
         sanadan = sanadan & "���� ����� ����� : ������ ��� ���� �� ���� �� ������ ������� ������� ���� � ������ ��� �� ���� ���� ������ �� ��� ��� �� ����� �������� � ������ ��� ������ �� ���� ������� � ������ ��� ������� �� ��� �� ���� ������ � ������ ��� ����� ��� �� ���� �� ��� ���� ������ � ������ ��� ����� ������� �� ������ �� ��� ���� ������ ������� ������ � ������ ����� �� ��� ������ ������." & vbNewLine
         sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ����� ��� ���� ��� ������ �� ���� ������� � ������� ��� ��� ��� ��� ���� �� ���� �� ��� ������ ������ � ���� ��� ��� ������� �� ���� � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ���� ��� ������ � ��� : ����� ��� ������ �� ���� ��� ���� ��� �������� ������ ��� ����� ��� ������ �� ��� ������ ������� � ���� ������� ���� �� ����� �� ������� ������ � ���� ������ ������� ��� ��� ��� ��� ������ ��� ��� ���� ���� �� ������ ��������� � ������ ��� ��� ��� ��� ������ ��� ������ ����� �� ���� �� ���� ������� � ���� ��� ������� ������� ��� ��� ��� ��� ������ ������ ��� ������ ���� �� ��� �� ����� ������� � ���� ������� ��� �� ������ ��� ������ ��� ��� ���� �� ���� �� ����� �� ���� ������� � ���� ������� �������� ����� ��� ����� � ���� ����� ��� ��� � ����� ������ . " & vbNewLine
         sanadan = sanadan & "����� ��� : ����� ��� ���� ���� ���� � ������ �� ����� ������ ���� ��� ��� � ���� ��� ���� ���� �� ��� �������� ���� ������ ����� ����� ������ � ���� ��� ��� � ������� � ����� ��� ���� . ���� ������� ���� �� ������� ��� ���� �� ��� �� ��� ��� � ����� ������ . ��� : ���� ���� �� ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ���� � ���� �� ��� � ���� �� ���� � ���� ���� �� ����� � �� ����� � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
         sanadan = sanadan & "���� �� �� ���� �� ����� �� ������� � ���� ����� � �� ���� ���� � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -. ����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ . ���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
         snandan = sanadan & vbNewLine
         
        
        ElseIf index = -2 Then
        
        ' ��� ����
        sanadan = "��� ����� ������ / ��� ���� ������" & vbNewLine
        sanadan = sanadan & "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ��� ������ : ������ ��� ���� �� ���� �� ��� � ���: ������ ��� ���� ���� �� ���� �� ��� ��� ���� ���� ��������ɡ ���: ������ ��� ���� ������ �� ���� ���:����� ������� �� ��� ���� � ��� ��� ���� : ����� ��� ������� ��� �� ���� ��� ��� ������ ��� ����� ��� ���� �� �� ���� �� ���� �� ����� �������� ������� ������� � � ��� �� : ���� ��� ��� ��� ���� ��� ������ �� ��� �� ��� ���� ������� � �� �� ����� ���� � ���� �� : ���� ��� ��� ��� ��� �� ����� � ���� : ���� ��� ��� ��� ������� ��� ������ �� ����� ���� :���� ��� ��� ��� � ���� : ���� ��� ��� ������� � ���� ���� ��� ��� : ��� ����. " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� ������ : ������ ��� ��� �� ������� �� ���� ������ � ��� : ����� ��� ���� ����� �� ���� ������ � ��� : ����� ��� ��� ������ ���� �� ���� ������� � ��� : ������ ��� ���� � ��� : ������ ������� � �� ��� ���� � ���� ��� ���� : ����� ��� ������ ��� ������ ����� �� ������� ����������� �������� ��� ���� �� ���� ������� � ���� �� : ���� ��� ���� ��� ��� ���� �� ������ ������� � ���� �� : ���� ��� ������ ��� ���� ��� ��� ����� ���� �� ���� ������ � ���� : ���� ��� ��� ��� ���� � ���� : ���� ��� ������� � ���� : ���� ��� ��� ����" & vbNewLine
        sanadan = sanadan & "��� ��� ����: ������ ����� ������� ���� �� ���� �� ��� ����� �� ��� ������ �� ����� �� ������ �� ������� �� ��� ���� ������ ��� ���� ��� ����� ����� � ��� : ����� ��� ���� �� ������� �� ���� �� ������ �� ��� ���� �� ������� �� ��� ���� . " & vbNewLine
        sanadan = sanadan & "����� ��� ���� : ����� �� ��� ������ ��� ��� ������ � ��� ��� ��� : ����� � ����� �� ���� � ������ �� ���� � ����� �� ��� ���� � ���� ���� �� ���� � ����� �� ��� ������ �� ����� � ����� �� ��� ������ ������ � ��� ��� ������� : ���� �� ������� ������� ����� �� ����� � ����� �� ���� � ��� ��� ������ : ����� �� ��� ����� ������ � ���� �� ���� � ������� � ���� ����� ������� ��� ���� �� ������� ������ . " & vbNewLine
        sanadan = sanadan & "��� : ���� ���� �� ���� � �����ɡ ����� �� ���� � �� ��� ���� ���� ��� ���� �� ��� �� ��� ���� �� ���� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
         
        '�����
        sanadan = sanadan & "��� ����� ������ / ����� ������" & vbNewLine
        sanadan = sanadan & "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ����� ������ ��� ������ ���� �� ���� �� ����� ������ ������� ���� ���: ������ : ��� ������ ���� �� ��� ���� �� ��� ����� ������� ����� ���� � ������ ��� ���� ��� ������ �� ���� �� ������� � �� ����� ������ ��� ��� ��� ���� �� ������ ������ ����� ���� � ������ ��� ���� ���� �� ��� ������� ������� ������ ��� ����� ��� �� ���� �� ��� ������ � ������ ������� ������ ��� ����� ��� �� ���� �� ��� ������� � ������ ��� ������ ��� ���� �� ����� �� ������ ������ � ������ ��� ��� ���� �� ����� �� ���� ������ �������� � ������ ��� ��� ���� ���� �� ������� ������� ����� ." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ������ ��� ���� ��� ������ �� ���� �� ��� �������� � ������� ��� ��� ��� ������ ��� ��� ������ ����� ���� �� ���� ������ � ���� ��� ��� ������� �� ���� ��������� � ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ���� �� ��� �������� � ���� ��� ��� ������� ��� ���� �������� � ���� ��� ��� ��� ��� ����� �� ������ ������� � ���� ��� ��� : ������� � ���� ��� ��� ����� � � ���� ��� ��� ������ � ���� ��� ���� � ���� ��� ��� ����� . " & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ����� ��� ������ ���� �� ���� �� ������ �������� ������� ���� �� ������ ��� ����� ��� �� ���� ������� � ������ ��� ����� ������ ����� � ������ ��� ���� �������� � ������ ��� ����� ������ ����� � ������ ���� �� ������ ������� � ������ ��� ����� ��� �� ���� �� ������� �� ����� ������� ������ ������ ��� ������ ���� �� ����� �� ������ �� ������ ������ � ������ ��� ��� ���� �� ��� �� ���� �� ������ ������ ������ � ������ ��� �� ��� ������ ������ ." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ���� �� ���� �������� �������� � ������� ��� ��� ��� ������ ��� ��� ������ ��� ��� ���� ������ � ���� ��� ��� ��� ����� ������� ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ���� �� ��� � ���� ��� ��� ������� ��� ���� �� ���� � ���� ��� ��� ��� ������ ������� �� ����� �� ���� ������ � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ��� � ������ � ���� ��� ��� ��� ��� � ���� ��� ��� ��� � ���� ��� ��� ����� ." & vbNewLine
        sanadan = sanadan & "������� ����� ����� ����� ����� : ��� ������ ���� �� ������ ������ � ����� �� ����� � ����� �� ����� � ���� ������ ���� �� ���� �������� .���� �� ����� ��� ��� ��� ���� �� ������ ���� ���� ��� ���� ���� ���� � ������� ���� ������� ���� ���� ��� ����� ������ ���� ������ �� ��� ���� ������ ��� ���� ��� � ���� ���� ��� ���� �� ������� ���� ��� ��� ������� ������� ���� ��� ��� ���� ���� ��� ������ ��� ��� ���� ����� �� ����� �������� ���� ��� ��� ������� ������� ���� ��� ���� ���� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
  
         ElseIf index = -3 Then
        
         ' ����
        sanadan = "��� ����� ������ / ����" & vbNewLine
        sanadan = sanadan & "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "��� ����� ����� : ������ ��� ���� �� ��� �� ���� ������ � ���: ����� ���� �� ���� �� ���� � ���: ����� ��� ���� �� ���� ������ � ���:����� ����� �� ���ڡ � ��� ����� ������ : ����� ��� ������� ��� ��� ���� ��� ����� ���� �� ���� �� ���� �� ����� � ������� ������ � � ��� ��: ���� ��� ��� ��� ����� ��� ������ �� ����� ������ � ����: ���� ��� ������� �� ��� ������ � ����: ���� ��� ��� ��� ������ ���� �� ����� �� ���� �� ����� � ����:���� ��� ��� ��� ���� �� ���� �� ������ ����: ���� ��� ��� ���� ���� �� ����� � ����: ���� ��� ����� � ����: ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & "���� ����� ��� : ������ ��� ��� ��� ���� ���� �� ����� ������ ���� � ���: ����� ���� �� ������� �� ���� � ��� : ����� ��� ���� ��� �� ��� � ���: ����� ��� ���� ��� ����� �� ��� ������ � ��� : ����� ��� �� ���� � � ��� ����� ������ : ����� ��� ������� ��� ��� ���� ��� ������ ��� �� ������� �� ���� �� ����� ������� ���� � � ��� �� : ���� ��� ������ ��� ��� ���� ���� �� ����� ������� � ���� �� : ���� ��� ������ ��� ������� �� ��� ���� ������ � ���� : ���� ��� ��� ����� ���� �� ���� �� ���� ������ � ���� :���� ��� ��� ���� : ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & "����� ���� ����� ����� ���� : ��� ���� �� �� �� ������� ������ � ���� ���� ��� ������ �� ���� ������ � ����� �� ���� ������ � ���� ��� ���� ���� �� ���� ������ ����� � ���� ��� ���� �� ����� � ���� ����� ������� �� ��� ����� � ���� ���� � ���� ���� �� ���� �� ��� ����� � �� ��� �� ��� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
      
         ' ��� ����
        sanadan = sanadan & "��� ����� ������ / ��� ����" & vbNewLine
        sanadan = sanadan & "��� ���� ���� ������� �� �������  " & vbNewLine
        sanadan = sanadan & "���� ����� ����� : ������ ��� ���� �� ���� �� ������ � ���:����� ���� �� ���� � ���: ����� ��� �� ���� ����� � ���:����� ���� �� ��� ��� � ���: ���� ��� ����� �� ������ �� ���� � ����: ���� ��� ������� �� ��� ���� ����� � ���� : ���� ��� ��� ���� ���� � ��� ��� ���� : ����� ��� ������� ��� ��� ��� ������ ��� ������ �� ���� �� ���� ������� ������� � ���� ��: ���� ��� ������� ��� ��� ��� ��� ���� �� ����� ������ � ����: ���� ��� ��� ��� ����� ���� �� ����� ��� ��� � ����: ���� ��� ����� ." & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ��� ���� ���� �� ���� �������� � ���: ����� ��� ����� � ���: ���� ��� ���� � ����: ���� ��� ��� ����� ���� �� ��� ������ ����: ���� ��� ��� ����� �� ��� �� ���� � ����: ���� ��� ������� �� ��� ���� ����� � ����: ���� ��� ��� �� ���� � ����� �� ����� � ����� ����� ��� ��� ������ � � ��� ������� �������� : ����� ��� ������ ��� ��� ���� �� ���� ������ ������� ������ ����: ���� ��� ��� ��� ���� �� ������ �������� � ����: ���� ��� ��� ��� ����� ����: ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & " ������� ��� ���� ������� ����� ����� : ��� ���� �� ������ �������� ���� ���� ����  ������ �� ��� ��� ������ ���� ��� �� ������ � ������ ���� ��� ���� . ���� ��� ���� �� ��� �� ��� ����. ���� ����� �����ӡ �� ��� ���ӡ �� ��� � ���� �� ���� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ -  ��� �� ����� - ����� � ����� -." & vbNewLine
      
       ' ��� ����
        sanadan = sanadan & "��� ����� ������ / ��� ���� ������" & vbNewLine
        sanadan = sanadan & "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ��� ������ : ������ ��� ���� �� ���� �� ��� � ���: ������ ��� ���� ���� �� ���� �� ��� ��� ���� ���� ��������ɡ ���: ������ ��� ���� ������ �� ���� ���:����� ������� �� ��� ���� � ��� ��� ���� : ����� ��� ������� ��� �� ���� ��� ��� ������ ��� ����� ��� ���� �� �� ���� �� ���� �� ����� �������� ������� ������� � � ��� �� : ���� ��� ��� ��� ���� ��� ������ �� ��� �� ��� ���� ������� � �� �� ����� ���� � ���� �� : ���� ��� ��� ��� ��� �� ����� � ���� : ���� ��� ��� ��� ������� ��� ������ �� ����� ���� :���� ��� ��� ��� � ���� : ���� ��� ��� ������� � ���� ���� ��� ��� : ��� ����. " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� ������ : ������ ��� ��� �� ������� �� ���� ������ � ��� : ����� ��� ���� ����� �� ���� ������ � ��� : ����� ��� ��� ������ ���� �� ���� ������� � ��� : ������ ��� ���� � ��� : ������ ������� � �� ��� ���� � ���� ��� ���� : ����� ��� ������ ��� ������ ����� �� ������� ����������� �������� ��� ���� �� ���� ������� � ���� �� : ���� ��� ���� ��� ��� ���� �� ������ ������� � ���� �� : ���� ��� ������ ��� ���� ��� ��� ����� ���� �� ���� ������ � ���� : ���� ��� ��� ��� ���� � ���� : ���� ��� ������� � ���� : ���� ��� ��� ����" & vbNewLine
        sanadan = sanadan & "��� ��� ����: ������ ����� ������� ���� �� ���� �� ��� ����� �� ��� ������ �� ����� �� ������ �� ������� �� ��� ���� ������ ��� ���� ��� ����� ����� � ��� : ����� ��� ���� �� ������� �� ���� �� ������ �� ��� ���� �� ������� �� ��� ���� . " & vbNewLine
        sanadan = sanadan & "����� ��� ���� : ����� �� ��� ������ ��� ��� ������ � ��� ��� ��� : ����� � ����� �� ���� � ������ �� ���� � ����� �� ��� ���� � ���� ���� �� ���� � ����� �� ��� ������ �� ����� � ����� �� ��� ������ ������ � ��� ��� ������� : ���� �� ������� ������� ����� �� ����� � ����� �� ���� � ��� ��� ������ : ����� �� ��� ����� ������ � ���� �� ���� � ������� � ���� ����� ������� ��� ���� �� ������� ������ . " & vbNewLine
        sanadan = sanadan & "��� : ���� ���� �� ���� � �����ɡ ����� �� ���� � �� ��� ���� ���� ��� ���� �� ��� �� ��� ���� �� ���� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        '��� ����
         sanadan = sanadan & "��� ����� ������ / ��� ����" & vbNewLine
         sanadan = sanadan & "��� ��� ���� ������ �� �������:" & vbNewLine
         sanadan = sanadan & "���� ����� ���� : ������ ��� ���� �� ���� ���: ����� ��� ����� � ��� : ����� ������ �� ����� ������ � ��� :����� ���� �� ���� �������� � ��� : ����� ���� �� ���� � ���: ����� ���� �� ���� ����� � ��� :���� ��� ���� �� ������ ������� � ����: ���� ��� ��� ���� �� ���� � ��� : ��� ���� : ����� ��� ������� ��� ��� ��� ����� ����� � ����: ���� ��� ��� ��� ���� �� ������ ������� � � ��� : ���� ��� ��� ���� �� ���� �� ����� � ���� : ���� ��� �������� � ���� : ���� ��� ���� " & vbNewLine
         sanadan = sanadan & "���� ����� ��� ����� : ������ ��� ���� �� ���� � ���:����� ���� �� ���� �� ����� � ��� : ����� ���� �� ���� ������� � ��� : ����� ��� ���� �� ����� � ��� : ����� ���� �� ���� ������� � ��� :����� ���� �� ������ ������� � ��� : ���� ��� ��� ���� � ���� ��� ���� : ����� ��� ������ ��� ��� ��� ������ �� ���� ������� ������� ���� �� : ���� ��� ��� ��� ��� ���� �� ����� ������ � ���� : ���� ��� ��� ��� ��� ���� ����� �� ���� �� ���� ������ ������ ������ �� ��� ���� �� ����� " & vbNewLine
         sanadan = sanadan & "������� ��� ���� ������� ������� : ��� ������� ����� �� ���� ���� ���� ���� � �������� �� ��� ���� �������� � ����� ��� ������� ���� ����� . ���� ������� �� ����� �� ���� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -" & vbNewLine
         snandan = sanadan & vbNewLine
        
        
         '����
         sanadan = sanadan & "��� ����� ������ / ����" & vbNewLine
         sanadan = sanadan & "��� ����� ������ ��������� ����� ��������:" & vbNewLine
         sanadan = sanadan & "���� ����� ��� ��� ����: ������ ��� ���� �� ���� �� ��� ������ ���: ����� �� ����� ���: ����� ������� �� ���� �� ��� ������� � ���:����� ��� ���:����� ���� �� ��� � ���: ����� ��� ��� �� ���� � ��� ��� ����: ����� ��� ������� ��� ��� ���� �� ���� ������� � � ��� ��: ���� ��� ��� ��� ����� ��� ������ �� ����� ������� � ����: ���� ��� ������� �� ��� ������ �� ���� ������� �������� ����: ���� ��� ���� �� ����� ������� � ����: ���� ��� ���� �� ���� ��������� � ����: ���� ��� ��� ���� �� ��� �� ��� ��� �� ����." & vbNewLine
         sanadan = sanadan & "���� ����� ��� : ������ ��� ��� ����� ����� �� ����� ������ � ��� : ����� ��� ����� ��� �� ���� �� ���� ������� ������ ������ ������� � ���: ����� ��� ������ ���� �� ��� �������� � ����: ���� ��� ��� ���� ���� �� ������ � ����: ���� ��� ��� � ����: ���� ��� ����� � ���� ��� ���� : ����� ��� ������ ��� ��� ����� ��� ����� ���� ��: ���� ��� ��� ������� ����: ���� ��� �������� �� ���� �� ��� �� ����� . " & vbNewLine
         sanadan = sanadan & "����� ���� ������� ����� ����� : ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ����  � ���� �� ���  � ���� �� ����  � ���� ���� �� ����� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� - � ��� �� �� ���� �� ����� �� �������  � ���� �����  � �� ���� ���� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
         snandan = sanadan & vbNewLine
         
        '����
        sanadan = sanadan & "��� ����� ������ / ����" & vbNewLine
        sanadan = sanadan & "��� ���� ���� ������� �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� : ������ ��� ���� �� ���� � ��� : ����� ��� ����� � ����� ����� �� ��� ������ � ��� : ����� ��� � ���: �� ���� �� ���� � � ��� ����� ������ : ����� ��� ������� ��� ��� ��� ����� ����� � � ��� �� : ���� ��� ��� ��� ����� ���� �� ���� �� ���� ������� ������� � ���� �� : ���� ��� ��� ��� ������ ���� �� ����� �� ���� �� ����� � ���� �� :���� ��� ����� �� ��� ������ ��� �� ����� ������� ��� � ���� �� : ���� ��� ��� � ���� : ���� ��� ���� � � ��� : ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ���� �� ���� � ��� : ����� ���� �� ���� � ��� : ����� ���� �� ���� �� ����� ������ � �� ���� �� ���� �������� � �� ���� � �� ���� � �� ���� � � ��� ����� ������ : ����� ��� ������ ��� ��� ��� ����� ������ ����� � � ��� ��: ���� ��� ��� ��� ���� �� ������ ������� � ���� : ���� ��� ��� ���� �� ���� �� ����� � ���� : ���� ��� ��� ��� ���� �� ����� ������� ������ � ���� :���� ��� ���� ���� : ���� ��� ���� � ���� ���� ��� ����." & vbNewLine
        sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
        sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
       
        '�������
         sanadan = sanadan & "��� ����� ������ / �������" & vbNewLine
         sanadan = sanadan & "��� ��� ���� ������� �� �������:" & vbNewLine
         sanadan = sanadan & "���� ����� ��� ���� ������ : ������ ��� ��� ���� ��� ������ �� ��� �� ���� ������ � ��� : ����� ��� ��� ��� ���� �� ���� �� ������ ������� � ��� : ����� ���� �� ���� �� ��� ������� � ��� : ����� ��� ��� ������ � �� ������� � � ��� ��� ����� : ����� ��� ������� ��� ��� ��� ����� � ���� �� : ���� ��� ��� ��� ������ �� ����� � ���� : ���� ��� ��� ��� ���� �� ��� �� ������� ������� � � ��� :���� ��� ���� �� ���� � ���� : ���� ��� ��� ��� ������ � ���� : ���� ��� ������� ." & vbNewLine
         sanadan = sanadan & "���� ����� ��� ������ : ������ ��� ���� �� ���� � ��� : ����� ��� ��� ����� � ��� : ����� ���� �� ���� ( ������� ������) � �� ��� ������ � �� ������� � � ��� ����� ������ : ����� ��� ������� ��� ��� ���� �� ���� � � ��� �� : ���� ��� ��� ��� ����� ��� ������ �� ����� ������� � ���� : ���� ��� ��� ��� ������ ��� �� ��� � ���� : ���� ��� ���� �� ����� ������� ������ � ���� :���� ��� ���� �� ���� ( ������� ������) � ���� : ���� ��� ��� ������ � ���� : ���� ��� ������� ." & vbNewLine
         sanadan = sanadan & "����� ������� : ���� �� ���� ������ � ����� �� ��� �������� � ����� �� ��� ���� ������ � ������ �� ����� �������� ��� �� ���� ������ �������� �� ������� �� ���� � ��� ����� ����� ������ ." & vbNewLine
         sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
         sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
         snandan = sanadan & vbNewLine
          
        '��� ����
        sanadan = sanadan & "��� ����� ������ / ��� ����" & vbNewLine
        sanadan = sanadan & "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ����� : ������ ��� ����� ��� ��� ��� �� ����� �� ���� ������� ������� ���� ��� : ������ ��� ����� ��� �� ���� �� ��� ������ ������ ������ �� ������ ��� ����� ��� �� ����� ������ � ��� : ������ ��� ���� ��� ���� �� ��� �������� ������ ������ ��� ����� ��� ������ �� ��� ������ ������� � ������ ��� ��� ���� ���� �� ������ ��������� � ������ ��� ����� ���� �� ���� �� ������� ������ � ������ ��� ��� ���� �� ���� �� ����� ������ � ������ ��� ������ ����� �� ����� �� ���� ������ ������ ��� ����� ���� �� ���� �������� ������� ���� �� ���� ����� � ������ ���� �� �����." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ������ ��� ��� ���� ���� ��� ������ �� ��� ������ � ������� ��� ��� ��� ������ ��� ��� �������� ��� ��� ���� �� ���� �� ��� ������ ������ � ��� : ���� ��� ������ ��� ������ ������� �� ���� �� ���� ������� ��� : ���� ��� ��� ��� ����� ������ � ��� : ���� ��� ��� ������ ��� ����� ���� �� ��� ����� �� ����� �� ����� �������� � ��� : ���� ��� ��� ��� ������ ��� ����� �� ���� ������� � ��� : ���� ��� ��� ��� ���� ���� �� ����� ������ � ��� : ���� ��� ��� ��� ����� ������ ���: ���� ��� ��� ��� ��� �� ����� � ���: ���� ��� ��� ����� �� ����� � ��� : ���� ��� ��� �������� � ��� : ���� ��� ��� ����� � ��� : ���� ��� ��� ��� ����� . " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� : ������ ��� ��� ����� ������� �� ���� �� ������� �� ���� ������� ������� ���� �� ��� ��� ��� �� ����� �� ������ ������� � ������ ��� ����� �� ����� �������� � ������ ��� ���� ��� ������ � ������ ������� ��� ���� ���� �� ������ �� ����� ������� � ������ ������ ��� ������ ���� �� ����� ������ � ������ ��� ��� ����� �� ���� �������� � ������ ��� ����� ��� �� ���� ������� � ������ ��� ��� ���� �� ��� ������ �� ����� ������� � ������ ���� �� ���� �� ����� ������ ������� � ������ ���� �� ��� ���� �� ���� ������� � ������ ��� ������ ���� �� ��� ������ � ������ ��� ����� ���� �� ��� ������ ������ � ������ ���� �� ���� �� ������� �� ���� ��������� � ������ ������ �� ���� �� ��� �� ��� ���� �� ���� ������� � ������ ������� �� ���� �� ��� ���� ������ � ������ ������ �� ���� ��� ����." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ��� ���� �� ��� ������ ������ � ���� ��� ������ ��� ��� ���� �� ���� ������ � ���� ��� ��� ��� ����� �� ���� � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ������ � ���� ��� ��� ������� ��� ���� ���� �� ��� �� ���� ���� �� ���� � ���� ��� ��� ��� ��� ����� �� ����� ��������� � ���� ��� ��� ��� ��� ���� �� ��� ���� �� �������� ��������� � ���� ��� ��� ��� ��� ���� �� ���� �� ��� ������ � ���� ��� ��� ���� �� ���� �� ����� �������� � ���� ��� ��� ���� �� ���� ������ ������� � ���� ��� ��� ��� ���� � ���� ��� ��� ��� ��� ������ � ���� ��� ��� ��� ����� ������ � ���� ��� ��� ��� ���� � ���� ��� ��� ������� � ���� ��� ��� ��� ���� � ���� ��� ��� ��� ���� � ���� ��� ���� � ���� ����� � ��� ��� ���� ." & vbNewLine
        sanadan = sanadan & "������� ��� ���� ����� : ����� ��� ���� �� ���� �� ��� ����� � ���� ����� � ���� ���� . ���� ����� ������� ��� ��� �� ��� � ���� ��� ����� � ���� ���� � ���� ��� ��� �� ���� . ���� ��� �� ����� - ��� ���� ���� � ��� -� �� ����� - ���� ������ -  � �� �� ����� - ����� � ����� -." & vbNewLine
          
        '�����
        sanadan = sanadan & "��� ����� ������ / ����� ������" & vbNewLine
        sanadan = sanadan & "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ����� ������ ��� ������ ���� �� ���� �� ����� ������ ������� ���� ���: ������ : ��� ������ ���� �� ��� ���� �� ��� ����� ������� ����� ���� � ������ ��� ���� ��� ������ �� ���� �� ������� � �� ����� ������ ��� ��� ��� ���� �� ������ ������ ����� ���� � ������ ��� ���� ���� �� ��� ������� ������� ������ ��� ����� ��� �� ���� �� ��� ������ � ������ ������� ������ ��� ����� ��� �� ���� �� ��� ������� � ������ ��� ������ ��� ���� �� ����� �� ������ ������ � ������ ��� ��� ���� �� ����� �� ���� ������ �������� � ������ ��� ��� ���� ���� �� ������� ������� ����� ." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ������ ��� ���� ��� ������ �� ���� �� ��� �������� � ������� ��� ��� ��� ������ ��� ��� ������ ����� ���� �� ���� ������ � ���� ��� ��� ������� �� ���� ��������� � ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ���� �� ��� �������� � ���� ��� ��� ������� ��� ���� �������� � ���� ��� ��� ��� ��� ����� �� ������ ������� � ���� ��� ��� : ������� � ���� ��� ��� ����� � � ���� ��� ��� ������ � ���� ��� ���� � ���� ��� ��� ����� . " & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ����� ��� ������ ���� �� ���� �� ������ �������� ������� ���� �� ������ ��� ����� ��� �� ���� ������� � ������ ��� ����� ������ ����� � ������ ��� ���� �������� � ������ ��� ����� ������ ����� � ������ ���� �� ������ ������� � ������ ��� ����� ��� �� ���� �� ������� �� ����� ������� ������ ������ ��� ������ ���� �� ����� �� ������ �� ������ ������ � ������ ��� ��� ���� �� ��� �� ���� �� ������ ������ ������ � ������ ��� �� ��� ������ ������ ." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ���� �� ���� �������� �������� � ������� ��� ��� ��� ������ ��� ��� ������ ��� ��� ���� ������ � ���� ��� ��� ��� ����� ������� ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ���� �� ��� � ���� ��� ��� ������� ��� ���� �� ���� � ���� ��� ��� ��� ������ ������� �� ����� �� ���� ������ � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ��� � ������ � ���� ��� ��� ��� ��� � ���� ��� ��� ��� � ���� ��� ��� ����� ." & vbNewLine
        sanadan = sanadan & "������� ����� ����� ����� ����� : ��� ������ ���� �� ������ ������ � ����� �� ����� � ����� �� ����� � ���� ������ ���� �� ���� �������� .���� �� ����� ��� ��� ��� ���� �� ������ ���� ���� ��� ���� ���� ���� � ������� ���� ������� ���� ���� ��� ����� ������ ���� ������ �� ��� ���� ������ ��� ���� ��� � ���� ���� ��� ���� �� ������� ���� ��� ��� ������� ������� ���� ��� ��� ���� ���� ��� ������ ��� ��� ���� ����� �� ����� �������� ���� ��� ��� ������� ������� ���� ��� ���� ���� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
  
  
         '���
         sanadan = sanadan & "��� ����� ������ / ��� ������" & vbNewLine
         sanadan = sanadan & "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
         sanadan = sanadan & "��� ����� ����� ������ : ������ ��� ��� ��� ��� �� ����� ������� ���� ���� ���� � �� ���� ������ ������ ��� ������ ���� �� ������� �� ��� �������� ������� � ��� : ������ ����� � ��� : ������ ��� �������� ������ �� ����� ������� � ������ ��� ���� ���� �� ������ ������� � ������ ��� ������ ���� �� ��� ���� �� ����� ���������� � ������ ��� ����� ���� �� ��� ���� �� ���� �� ��� ������ ������� ���� ��� ��� ������ � ������ ��� ����� ����� �� ������� ������ ." & vbNewLine
         sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� �� �� ������� ��� ��� ���� ������ � ���� ���� ������� �������� � ���� �� ����� ��� ��� ��� ���� ���� �� ���� �� ��� ������ ������ � ���� ��� ��� ������ �� ���� � ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ������ ��� ���� �� ���� �� ����� �������� � ���� ��� ��� ��� ��� ���� �� ��� �� ���� ������ � ���� ��� ��� ��� ������ ���������� � ���� ��� ��� ��� ��� ��� ������ � ���� ��� ��� ����� ������ � ���� ��� ��� ��� ." & vbNewLine
         sanadan = sanadan & "���� ����� ����� : ������ ��� ���� �� ���� �� ������ ������� ������� ���� � ������ ��� �� ���� ���� ������ �� ��� ��� �� ����� �������� � ������ ��� ������ �� ���� ������� � ������ ��� ������� �� ��� �� ���� ������ � ������ ��� ����� ��� �� ���� �� ��� ���� ������ � ������ ��� ����� ������� �� ������ �� ��� ���� ������ ������� ������ � ������ ����� �� ��� ������ ������." & vbNewLine
         sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ����� ��� ���� ��� ������ �� ���� ������� � ������� ��� ��� ��� ��� ���� �� ���� �� ��� ������ ������ � ���� ��� ��� ������� �� ���� � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ���� ��� ������ � ��� : ����� ��� ������ �� ���� ��� ���� ��� �������� ������ ��� ����� ��� ������ �� ��� ������ ������� � ���� ������� ���� �� ����� �� ������� ������ � ���� ������ ������� ��� ��� ��� ��� ������ ��� ��� ���� ���� �� ������ ��������� � ������ ��� ��� ��� ��� ������ ��� ������ ����� �� ���� �� ���� ������� � ���� ��� ������� ������� ��� ��� ��� ��� ������ ������ ��� ������ ���� �� ��� �� ����� ������� � ���� ������� ��� �� ������ ��� ������ ��� ��� ���� �� ���� �� ����� �� ���� ������� � ���� ������� �������� ����� ��� ����� � ���� ����� ��� ��� � ����� ������ . " & vbNewLine
         sanadan = sanadan & "����� ��� : ����� ��� ���� ���� ���� � ������ �� ����� ������ ���� ��� ��� � ���� ��� ���� ���� �� ��� �������� ���� ������ ����� ����� ������ � ���� ��� ��� � ������� � ����� ��� ���� . ���� ������� ���� �� ������� ��� ���� �� ��� �� ��� ��� � ����� ������ . ��� : ���� ���� �� ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ���� � ���� �� ��� � ���� �� ���� � ���� ���� �� ����� � �� ����� � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
         sanadan = sanadan & "���� �� �� ���� �� ����� �� ������� � ���� ����� � �� ���� ���� � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -. ����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ . ���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
         snandan = sanadan & vbNewLine
 
        ElseIf index = -4 Then
        
        '�����
        sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan + "��� ����� ����� : ������ ��� ���� �� ��� �� ���� ������ � ���: ����� ���� �� ���� �� ���� � ���: ����� ��� ���� �� ���� ������ � ���:����� ����� �� ���ڡ � ��� ����� ������ : ����� ��� ������� ��� ��� ���� ��� ����� ���� �� ���� �� ���� �� ����� � ������� ������ � � ��� ��: ���� ��� ��� ��� ����� ��� ������ �� ����� ������ � ����: ���� ��� ������� �� ��� ������ � ����: ���� ��� ��� ��� ������ ���� �� ����� �� ���� �� ����� � ����:���� ��� ��� ��� ���� �� ���� �� ������ ����: ���� ��� ��� ���� ���� �� ����� � ����: ���� ��� ����� � ����: ���� ��� ���� ." & vbNewLine
        sanadan = sanadan + "����� ���� ����� ����� ���� : ��� ���� �� �� �� ������� ������ � ���� ���� ��� ������ �� ���� ������ � ����� �� ���� ������ � ���� ��� ���� ���� �� ���� ������ ����� � ���� ��� ���� �� ����� � ���� ����� ������� �� ��� ����� � ���� ���� � ���� ���� �� ���� �� ��� ����� � �� ��� �� ��� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
     
        ' ��� ����
        sanadan = sanadan & "��� ����� ������ / ��� ����" & vbNewLine
        sanadan = sanadan & "��� ���� ���� ������� �� �������  " & vbNewLine
        sanadan = sanadan & "���� ����� ����� : ������ ��� ���� �� ���� �� ������ � ���:����� ���� �� ���� � ���: ����� ��� �� ���� ����� � ���:����� ���� �� ��� ��� � ���: ���� ��� ����� �� ������ �� ���� � ����: ���� ��� ������� �� ��� ���� ����� � ���� : ���� ��� ��� ���� ���� � ��� ��� ���� : ����� ��� ������� ��� ��� ��� ������ ��� ������ �� ���� �� ���� ������� ������� � ���� ��: ���� ��� ������� ��� ��� ��� ��� ���� �� ����� ������ � ����: ���� ��� ��� ��� ����� ���� �� ����� ��� ��� � ����: ���� ��� ����� ." & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ��� ���� ���� �� ���� �������� � ���: ����� ��� ����� � ���: ���� ��� ���� � ����: ���� ��� ��� ����� ���� �� ��� ������ ����: ���� ��� ��� ����� �� ��� �� ���� � ����: ���� ��� ������� �� ��� ���� ����� � ����: ���� ��� ��� �� ���� � ����� �� ����� � ����� ����� ��� ��� ������ � � ��� ������� �������� : ����� ��� ������ ��� ��� ���� �� ���� ������ ������� ������ ����: ���� ��� ��� ��� ���� �� ������ �������� � ����: ���� ��� ��� ��� ����� ����: ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & " ������� ��� ���� ������� ����� ����� : ��� ���� �� ������ �������� ���� ���� ����  ������ �� ��� ��� ������ ���� ��� �� ������ � ������ ���� ��� ���� . ���� ��� ���� �� ��� �� ��� ����. ���� ����� �����ӡ �� ��� ���ӡ �� ��� � ���� �� ���� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ -  ��� �� ����� - ����� � ����� -."
      
        '��� ����
        sanadan = sanadan & "��� ����� ������ / ��� ����" & vbNewLine
        sanadan = sanadan & "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ����� : ������ ��� ����� ��� ��� ��� �� ����� �� ���� ������� ������� ���� ��� : ������ ��� ����� ��� �� ���� �� ��� ������ ������ ������ �� ������ ��� ����� ��� �� ����� ������ � ��� : ������ ��� ���� ��� ���� �� ��� �������� ������ ������ ��� ����� ��� ������ �� ��� ������ ������� � ������ ��� ��� ���� ���� �� ������ ��������� � ������ ��� ����� ���� �� ���� �� ������� ������ � ������ ��� ��� ���� �� ���� �� ����� ������ � ������ ��� ������ ����� �� ����� �� ���� ������ ������ ��� ����� ���� �� ���� �������� ������� ���� �� ���� ����� � ������ ���� �� �����." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ������ ��� ��� ���� ���� ��� ������ �� ��� ������ � ������� ��� ��� ��� ������ ��� ��� �������� ��� ��� ���� �� ���� �� ��� ������ ������ � ��� : ���� ��� ������ ��� ������ ������� �� ���� �� ���� ������� ��� : ���� ��� ��� ��� ����� ������ � ��� : ���� ��� ��� ������ ��� ����� ���� �� ��� ����� �� ����� �� ����� �������� � ��� : ���� ��� ��� ��� ������ ��� ����� �� ���� ������� � ��� : ���� ��� ��� ��� ���� ���� �� ����� ������ � ��� : ���� ��� ��� ��� ����� ������ ���: ���� ��� ��� ��� ��� �� ����� � ���: ���� ��� ��� ����� �� ����� � ��� : ���� ��� ��� �������� � ��� : ���� ��� ��� ����� � ��� : ���� ��� ��� ��� ����� . " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� : ������ ��� ��� ����� ������� �� ���� �� ������� �� ���� ������� ������� ���� �� ��� ��� ��� �� ����� �� ������ ������� � ������ ��� ����� �� ����� �������� � ������ ��� ���� ��� ������ � ������ ������� ��� ���� ���� �� ������ �� ����� ������� � ������ ������ ��� ������ ���� �� ����� ������ � ������ ��� ��� ����� �� ���� �������� � ������ ��� ����� ��� �� ���� ������� � ������ ��� ��� ���� �� ��� ������ �� ����� ������� � ������ ���� �� ���� �� ����� ������ ������� � ������ ���� �� ��� ���� �� ���� ������� � ������ ��� ������ ���� �� ��� ������ � ������ ��� ����� ���� �� ��� ������ ������ � ������ ���� �� ���� �� ������� �� ���� ��������� � ������ ������ �� ���� �� ��� �� ��� ���� �� ���� ������� � ������ ������� �� ���� �� ��� ���� ������ � ������ ������ �� ���� ��� ����." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ��� ���� �� ��� ������ ������ � ���� ��� ������ ��� ��� ���� �� ���� ������ � ���� ��� ��� ��� ����� �� ���� � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ������ � ���� ��� ��� ������� ��� ���� ���� �� ��� �� ���� ���� �� ���� � ���� ��� ��� ��� ��� ����� �� ����� ��������� � ���� ��� ��� ��� ��� ���� �� ��� ���� �� �������� ��������� � ���� ��� ��� ��� ��� ���� �� ���� �� ��� ������ � ���� ��� ��� ���� �� ���� �� ����� �������� � ���� ��� ��� ���� �� ���� ������ ������� � ���� ��� ��� ��� ���� � ���� ��� ��� ��� ��� ������ � ���� ��� ��� ��� ����� ������ � ���� ��� ��� ��� ���� � ���� ��� ��� ������� � ���� ��� ��� ��� ���� � ���� ��� ��� ��� ���� � ���� ��� ���� � ���� ����� � ��� ��� ���� ." & vbNewLine
        sanadan = sanadan & "������� ��� ���� ����� : ����� ��� ���� �� ���� �� ��� ����� � ���� ����� � ���� ���� . ���� ����� ������� ��� ��� �� ��� � ���� ��� ����� � ���� ���� � ���� ��� ��� �� ���� . ���� ��� �� ����� - ��� ���� ���� � ��� -� �� ����� - ���� ������ -  � �� �� ����� - ����� � ����� -." & vbNewLine
       
         ElseIf index = -5 Then
        
         ' ����
        sanadan = "��� ����� ������ / ����" & vbNewLine
        sanadan = sanadan & "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "��� ����� ����� : ������ ��� ���� �� ��� �� ���� ������ � ���: ����� ���� �� ���� �� ���� � ���: ����� ��� ���� �� ���� ������ � ���:����� ����� �� ���ڡ � ��� ����� ������ : ����� ��� ������� ��� ��� ���� ��� ����� ���� �� ���� �� ���� �� ����� � ������� ������ � � ��� ��: ���� ��� ��� ��� ����� ��� ������ �� ����� ������ � ����: ���� ��� ������� �� ��� ������ � ����: ���� ��� ��� ��� ������ ���� �� ����� �� ���� �� ����� � ����:���� ��� ��� ��� ���� �� ���� �� ������ ����: ���� ��� ��� ���� ���� �� ����� � ����: ���� ��� ����� � ����: ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & "���� ����� ��� : ������ ��� ��� ��� ���� ���� �� ����� ������ ���� � ���: ����� ���� �� ������� �� ���� � ��� : ����� ��� ���� ��� �� ��� � ���: ����� ��� ���� ��� ����� �� ��� ������ � ��� : ����� ��� �� ���� � � ��� ����� ������ : ����� ��� ������� ��� ��� ���� ��� ������ ��� �� ������� �� ���� �� ����� ������� ���� � � ��� �� : ���� ��� ������ ��� ��� ���� ���� �� ����� ������� � ���� �� : ���� ��� ������ ��� ������� �� ��� ���� ������ � ���� : ���� ��� ��� ����� ���� �� ���� �� ���� ������ � ���� :���� ��� ��� ���� : ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & "����� ���� ����� ����� ���� : ��� ���� �� �� �� ������� ������ � ���� ���� ��� ������ �� ���� ������ � ����� �� ���� ������ � ���� ��� ���� ���� �� ���� ������ ����� � ���� ��� ���� �� ����� � ���� ����� ������� �� ��� ����� � ���� ���� � ���� ���� �� ���� �� ��� ����� � �� ��� �� ��� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
      
         ' ��� ����
        sanadan = sanadan & "��� ����� ������ / ��� ����" & vbNewLine
        sanadan = sanadan & "��� ���� ���� ������� �� �������  " & vbNewLine
        sanadan = sanadan & "���� ����� ����� : ������ ��� ���� �� ���� �� ������ � ���:����� ���� �� ���� � ���: ����� ��� �� ���� ����� � ���:����� ���� �� ��� ��� � ���: ���� ��� ����� �� ������ �� ���� � ����: ���� ��� ������� �� ��� ���� ����� � ���� : ���� ��� ��� ���� ���� � ��� ��� ���� : ����� ��� ������� ��� ��� ��� ������ ��� ������ �� ���� �� ���� ������� ������� � ���� ��: ���� ��� ������� ��� ��� ��� ��� ���� �� ����� ������ � ����: ���� ��� ��� ��� ����� ���� �� ����� ��� ��� � ����: ���� ��� ����� ." & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ��� ���� ���� �� ���� �������� � ���: ����� ��� ����� � ���: ���� ��� ���� � ����: ���� ��� ��� ����� ���� �� ��� ������ ����: ���� ��� ��� ����� �� ��� �� ���� � ����: ���� ��� ������� �� ��� ���� ����� � ����: ���� ��� ��� �� ���� � ����� �� ����� � ����� ����� ��� ��� ������ � � ��� ������� �������� : ����� ��� ������ ��� ��� ���� �� ���� ������ ������� ������ ����: ���� ��� ��� ��� ���� �� ������ �������� � ����: ���� ��� ��� ��� ����� ����: ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & " ������� ��� ���� ������� ����� ����� : ��� ���� �� ������ �������� ���� ���� ����  ������ �� ��� ��� ������ ���� ��� �� ������ � ������ ���� ��� ���� . ���� ��� ���� �� ��� �� ��� ����. ���� ����� �����ӡ �� ��� ���ӡ �� ��� � ���� �� ���� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ -  ��� �� ����� - ����� � ����� -." & vbNewLine
      
       ' ��� ����
        sanadan = sanadan & "��� ����� ������ / ��� ���� ������" & vbNewLine
        sanadan = sanadan & "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ��� ������ : ������ ��� ���� �� ���� �� ��� � ���: ������ ��� ���� ���� �� ���� �� ��� ��� ���� ���� ��������ɡ ���: ������ ��� ���� ������ �� ���� ���:����� ������� �� ��� ���� � ��� ��� ���� : ����� ��� ������� ��� �� ���� ��� ��� ������ ��� ����� ��� ���� �� �� ���� �� ���� �� ����� �������� ������� ������� � � ��� �� : ���� ��� ��� ��� ���� ��� ������ �� ��� �� ��� ���� ������� � �� �� ����� ���� � ���� �� : ���� ��� ��� ��� ��� �� ����� � ���� : ���� ��� ��� ��� ������� ��� ������ �� ����� ���� :���� ��� ��� ��� � ���� : ���� ��� ��� ������� � ���� ���� ��� ��� : ��� ����. " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� ������ : ������ ��� ��� �� ������� �� ���� ������ � ��� : ����� ��� ���� ����� �� ���� ������ � ��� : ����� ��� ��� ������ ���� �� ���� ������� � ��� : ������ ��� ���� � ��� : ������ ������� � �� ��� ���� � ���� ��� ���� : ����� ��� ������ ��� ������ ����� �� ������� ����������� �������� ��� ���� �� ���� ������� � ���� �� : ���� ��� ���� ��� ��� ���� �� ������ ������� � ���� �� : ���� ��� ������ ��� ���� ��� ��� ����� ���� �� ���� ������ � ���� : ���� ��� ��� ��� ���� � ���� : ���� ��� ������� � ���� : ���� ��� ��� ����" & vbNewLine
        sanadan = sanadan & "��� ��� ����: ������ ����� ������� ���� �� ���� �� ��� ����� �� ��� ������ �� ����� �� ������ �� ������� �� ��� ���� ������ ��� ���� ��� ����� ����� � ��� : ����� ��� ���� �� ������� �� ���� �� ������ �� ��� ���� �� ������� �� ��� ���� . " & vbNewLine
        sanadan = sanadan & "����� ��� ���� : ����� �� ��� ������ ��� ��� ������ � ��� ��� ��� : ����� � ����� �� ���� � ������ �� ���� � ����� �� ��� ���� � ���� ���� �� ���� � ����� �� ��� ������ �� ����� � ����� �� ��� ������ ������ � ��� ��� ������� : ���� �� ������� ������� ����� �� ����� � ����� �� ���� � ��� ��� ������ : ����� �� ��� ����� ������ � ���� �� ���� � ������� � ���� ����� ������� ��� ���� �� ������� ������ . " & vbNewLine
        sanadan = sanadan & "��� : ���� ���� �� ���� � �����ɡ ����� �� ���� � �� ��� ���� ���� ��� ���� �� ��� �� ��� ���� �� ���� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        '��� ����
         sanadan = sanadan & "��� ����� ������ / ��� ����" & vbNewLine
         sanadan = sanadan & "��� ��� ���� ������ �� �������:" & vbNewLine
         sanadan = sanadan & "���� ����� ���� : ������ ��� ���� �� ���� ���: ����� ��� ����� � ��� : ����� ������ �� ����� ������ � ��� :����� ���� �� ���� �������� � ��� : ����� ���� �� ���� � ���: ����� ���� �� ���� ����� � ��� :���� ��� ���� �� ������ ������� � ����: ���� ��� ��� ���� �� ���� � ��� : ��� ���� : ����� ��� ������� ��� ��� ��� ����� ����� � ����: ���� ��� ��� ��� ���� �� ������ ������� � � ��� : ���� ��� ��� ���� �� ���� �� ����� � ���� : ���� ��� �������� � ���� : ���� ��� ���� " & vbNewLine
         sanadan = sanadan & "���� ����� ��� ����� : ������ ��� ���� �� ���� � ���:����� ���� �� ���� �� ����� � ��� : ����� ���� �� ���� ������� � ��� : ����� ��� ���� �� ����� � ��� : ����� ���� �� ���� ������� � ��� :����� ���� �� ������ ������� � ��� : ���� ��� ��� ���� � ���� ��� ���� : ����� ��� ������ ��� ��� ��� ������ �� ���� ������� ������� ���� �� : ���� ��� ��� ��� ��� ���� �� ����� ������ � ���� : ���� ��� ��� ��� ��� ���� ����� �� ���� �� ���� ������ ������ ������ �� ��� ���� �� ����� " & vbNewLine
         sanadan = sanadan & "������� ��� ���� ������� ������� : ��� ������� ����� �� ���� ���� ���� ���� � �������� �� ��� ���� �������� � ����� ��� ������� ���� ����� . ���� ������� �� ����� �� ���� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -" & vbNewLine
         snandan = sanadan & vbNewLine
        
        
         '����
         sanadan = sanadan & "��� ����� ������ / ����" & vbNewLine
         sanadan = sanadan & "��� ����� ������ ��������� ����� ��������:" & vbNewLine
         sanadan = sanadan & "���� ����� ��� ��� ����: ������ ��� ���� �� ���� �� ��� ������ ���: ����� �� ����� ���: ����� ������� �� ���� �� ��� ������� � ���:����� ��� ���:����� ���� �� ��� � ���: ����� ��� ��� �� ���� � ��� ��� ����: ����� ��� ������� ��� ��� ���� �� ���� ������� � � ��� ��: ���� ��� ��� ��� ����� ��� ������ �� ����� ������� � ����: ���� ��� ������� �� ��� ������ �� ���� ������� �������� ����: ���� ��� ���� �� ����� ������� � ����: ���� ��� ���� �� ���� ��������� � ����: ���� ��� ��� ���� �� ��� �� ��� ��� �� ����." & vbNewLine
         sanadan = sanadan & "���� ����� ��� : ������ ��� ��� ����� ����� �� ����� ������ � ��� : ����� ��� ����� ��� �� ���� �� ���� ������� ������ ������ ������� � ���: ����� ��� ������ ���� �� ��� �������� � ����: ���� ��� ��� ���� ���� �� ������ � ����: ���� ��� ��� � ����: ���� ��� ����� � ���� ��� ���� : ����� ��� ������ ��� ��� ����� ��� ����� ���� ��: ���� ��� ��� ������� ����: ���� ��� �������� �� ���� �� ��� �� ����� . " & vbNewLine
         sanadan = sanadan & "����� ���� ������� ����� ����� : ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ����  � ���� �� ���  � ���� �� ����  � ���� ���� �� ����� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� - � ��� �� �� ���� �� ����� �� �������  � ���� �����  � �� ���� ���� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
         snandan = sanadan & vbNewLine
         
        '����
        sanadan = sanadan & "��� ����� ������ / ����" & vbNewLine
        sanadan = sanadan & "��� ���� ���� ������� �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� : ������ ��� ���� �� ���� � ��� : ����� ��� ����� � ����� ����� �� ��� ������ � ��� : ����� ��� � ���: �� ���� �� ���� � � ��� ����� ������ : ����� ��� ������� ��� ��� ��� ����� ����� � � ��� �� : ���� ��� ��� ��� ����� ���� �� ���� �� ���� ������� ������� � ���� �� : ���� ��� ��� ��� ������ ���� �� ����� �� ���� �� ����� � ���� �� :���� ��� ����� �� ��� ������ ��� �� ����� ������� ��� � ���� �� : ���� ��� ��� � ���� : ���� ��� ���� � � ��� : ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ���� �� ���� � ��� : ����� ���� �� ���� � ��� : ����� ���� �� ���� �� ����� ������ � �� ���� �� ���� �������� � �� ���� � �� ���� � �� ���� � � ��� ����� ������ : ����� ��� ������ ��� ��� ��� ����� ������ ����� � � ��� ��: ���� ��� ��� ��� ���� �� ������ ������� � ���� : ���� ��� ��� ���� �� ���� �� ����� � ���� : ���� ��� ��� ��� ���� �� ����� ������� ������ � ���� :���� ��� ���� ���� : ���� ��� ���� � ���� ���� ��� ����." & vbNewLine
        sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
        sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
       
        '�������
         sanadan = sanadan & "��� ����� ������ / �������" & vbNewLine
         sanadan = sanadan & "��� ��� ���� ������� �� �������:" & vbNewLine
         sanadan = sanadan & "���� ����� ��� ���� ������ : ������ ��� ��� ���� ��� ������ �� ��� �� ���� ������ � ��� : ����� ��� ��� ��� ���� �� ���� �� ������ ������� � ��� : ����� ���� �� ���� �� ��� ������� � ��� : ����� ��� ��� ������ � �� ������� � � ��� ��� ����� : ����� ��� ������� ��� ��� ��� ����� � ���� �� : ���� ��� ��� ��� ������ �� ����� � ���� : ���� ��� ��� ��� ���� �� ��� �� ������� ������� � � ��� :���� ��� ���� �� ���� � ���� : ���� ��� ��� ��� ������ � ���� : ���� ��� ������� ." & vbNewLine
         sanadan = sanadan & "���� ����� ��� ������ : ������ ��� ���� �� ���� � ��� : ����� ��� ��� ����� � ��� : ����� ���� �� ���� ( ������� ������) � �� ��� ������ � �� ������� � � ��� ����� ������ : ����� ��� ������� ��� ��� ���� �� ���� � � ��� �� : ���� ��� ��� ��� ����� ��� ������ �� ����� ������� � ���� : ���� ��� ��� ��� ������ ��� �� ��� � ���� : ���� ��� ���� �� ����� ������� ������ � ���� :���� ��� ���� �� ���� ( ������� ������) � ���� : ���� ��� ��� ������ � ���� : ���� ��� ������� ." & vbNewLine
         sanadan = sanadan & "����� ������� : ���� �� ���� ������ � ����� �� ��� �������� � ����� �� ��� ���� ������ � ������ �� ����� �������� ��� �� ���� ������ �������� �� ������� �� ���� � ��� ����� ����� ������ ." & vbNewLine
         sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
         sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
         snandan = sanadan & vbNewLine
          
       
        ElseIf index = 1 Then
        ' ����
        sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "��� ����� ����� : ������ ��� ���� �� ��� �� ���� ������ � ���: ����� ���� �� ���� �� ���� � ���: ����� ��� ���� �� ���� ������ � ���:����� ����� �� ���ڡ � ��� ����� ������ : ����� ��� ������� ��� ��� ���� ��� ����� ���� �� ���� �� ���� �� ����� � ������� ������ � � ��� ��: ���� ��� ��� ��� ����� ��� ������ �� ����� ������ � ����: ���� ��� ������� �� ��� ������ � ����: ���� ��� ��� ��� ������ ���� �� ����� �� ���� �� ����� � ����:���� ��� ��� ��� ���� �� ���� �� ������ ����: ���� ��� ��� ���� ���� �� ����� � ����: ���� ��� ����� � ����: ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & "���� ����� ��� : ������ ��� ��� ��� ���� ���� �� ����� ������ ���� � ���: ����� ���� �� ������� �� ���� � ��� : ����� ��� ���� ��� �� ��� � ���: ����� ��� ���� ��� ����� �� ��� ������ � ��� : ����� ��� �� ���� � � ��� ����� ������ : ����� ��� ������� ��� ��� ���� ��� ������ ��� �� ������� �� ���� �� ����� ������� ���� � � ��� �� : ���� ��� ������ ��� ��� ���� ���� �� ����� ������� � ���� �� : ���� ��� ������ ��� ������� �� ��� ���� ������ � ���� : ���� ��� ��� ����� ���� �� ���� �� ���� ������ � ���� :���� ��� ��� ���� : ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & "����� ���� ����� ����� ���� : ��� ���� �� �� �� ������� ������ � ���� ���� ��� ������ �� ���� ������ � ����� �� ���� ������ � ���� ��� ���� ���� �� ���� ������ ����� � ���� ��� ���� �� ����� � ���� ����� ������� �� ��� ����� � ���� ���� � ���� ���� �� ���� �� ��� ����� � �� ��� �� ��� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 2 Then
        ' ��� ����
        sanadan = "��� ���� ���� ������� �� �������  " & vbNewLine
        sanadan = sanadan & "���� ����� ����� : ������ ��� ���� �� ���� �� ������ � ���:����� ���� �� ���� � ���: ����� ��� �� ���� ����� � ���:����� ���� �� ��� ��� � ���: ���� ��� ����� �� ������ �� ���� � ����: ���� ��� ������� �� ��� ���� ����� � ���� : ���� ��� ��� ���� ���� � ��� ��� ���� : ����� ��� ������� ��� ��� ��� ������ ��� ������ �� ���� �� ���� ������� ������� � ���� ��: ���� ��� ������� ��� ��� ��� ��� ���� �� ����� ������ � ����: ���� ��� ��� ��� ����� ���� �� ����� ��� ��� � ����: ���� ��� ����� ." & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ��� ���� ���� �� ���� �������� � ���: ����� ��� ����� � ���: ���� ��� ���� � ����: ���� ��� ��� ����� ���� �� ��� ������ ����: ���� ��� ��� ����� �� ��� �� ���� � ����: ���� ��� ������� �� ��� ���� ����� � ����: ���� ��� ��� �� ���� � ����� �� ����� � ����� ����� ��� ��� ������ � � ��� ������� �������� : ����� ��� ������ ��� ��� ���� �� ���� ������ ������� ������ ����: ���� ��� ��� ��� ���� �� ������ �������� � ����: ���� ��� ��� ��� ����� ����: ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & " ������� ��� ���� ������� ����� ����� : ��� ���� �� ������ �������� ���� ���� ����  ������ �� ��� ��� ������ ���� ��� �� ������ � ������ ���� ��� ���� . ���� ��� ���� �� ��� �� ��� ����. ���� ����� �����ӡ �� ��� ���ӡ �� ��� � ���� �� ���� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ -  ��� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 3 Then
        ' ��� ����
        sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ��� ������ : ������ ��� ���� �� ���� �� ��� � ���: ������ ��� ���� ���� �� ���� �� ��� ��� ���� ���� ��������ɡ ���: ������ ��� ���� ������ �� ���� ���:����� ������� �� ��� ���� � ��� ��� ���� : ����� ��� ������� ��� �� ���� ��� ��� ������ ��� ����� ��� ���� �� �� ���� �� ���� �� ����� �������� ������� ������� � � ��� �� : ���� ��� ��� ��� ���� ��� ������ �� ��� �� ��� ���� ������� � �� �� ����� ���� � ���� �� : ���� ��� ��� ��� ��� �� ����� � ���� : ���� ��� ��� ��� ������� ��� ������ �� ����� ���� :���� ��� ��� ��� � ���� : ���� ��� ��� ������� � ���� ���� ��� ��� : ��� ����. " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� ������ : ������ ��� ��� �� ������� �� ���� ������ � ��� : ����� ��� ���� ����� �� ���� ������ � ��� : ����� ��� ��� ������ ���� �� ���� ������� � ��� : ������ ��� ���� � ��� : ������ ������� � �� ��� ���� � ���� ��� ���� : ����� ��� ������ ��� ������ ����� �� ������� ����������� �������� ��� ���� �� ���� ������� � ���� �� : ���� ��� ���� ��� ��� ���� �� ������ ������� � ���� �� : ���� ��� ������ ��� ���� ��� ��� ����� ���� �� ���� ������ � ���� : ���� ��� ��� ��� ���� � ���� : ���� ��� ������� � ���� : ���� ��� ��� ����" & vbNewLine
        sanadan = sanadan & "��� ��� ����: ������ ����� ������� ���� �� ���� �� ��� ����� �� ��� ������ �� ����� �� ������ �� ������� �� ��� ���� ������ ��� ���� ��� ����� ����� � ��� : ����� ��� ���� �� ������� �� ���� �� ������ �� ��� ���� �� ������� �� ��� ���� . " & vbNewLine
        sanadan = sanadan & "����� ��� ���� : ����� �� ��� ������ ��� ��� ������ � ��� ��� ��� : ����� � ����� �� ���� � ������ �� ���� � ����� �� ��� ���� � ���� ���� �� ���� � ����� �� ��� ������ �� ����� � ����� �� ��� ������ ������ � ��� ��� ������� : ���� �� ������� ������� ����� �� ����� � ����� �� ���� � ��� ��� ������ : ����� �� ��� ����� ������ � ���� �� ���� � ������� � ���� ����� ������� ��� ���� �� ������� ������ . " & vbNewLine
        sanadan = sanadan & "��� : ���� ���� �� ���� � �����ɡ ����� �� ���� � �� ��� ���� ���� ��� ���� �� ��� �� ��� ���� �� ���� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -."
        
        ElseIf index = 4 Then
        '��� ����
        sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ���� �� ���� ���: ����� ��� ����� � ��� : ����� ������ �� ����� ������ � ��� :����� ���� �� ���� �������� � ��� : ����� ���� �� ���� � ���: ����� ���� �� ���� ����� � ��� :���� ��� ���� �� ������ ������� � ����: ���� ��� ��� ���� �� ���� � ��� : ��� ���� : ����� ��� ������� ��� ��� ��� ����� ����� � ����: ���� ��� ��� ��� ���� �� ������ ������� � � ��� : ���� ��� ��� ���� �� ���� �� ����� � ���� : ���� ��� �������� � ���� : ���� ��� ���� " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ����� : ������ ��� ���� �� ���� � ���:����� ���� �� ���� �� ����� � ��� : ����� ���� �� ���� ������� � ��� : ����� ��� ���� �� ����� � ��� : ����� ���� �� ���� ������� � ��� :����� ���� �� ������ ������� � ��� : ���� ��� ��� ���� � ���� ��� ���� : ����� ��� ������ ��� ��� ��� ������ �� ���� ������� ������� ���� �� : ���� ��� ��� ��� ��� ���� �� ����� ������ � ���� : ���� ��� ��� ��� ��� ���� ����� �� ���� �� ���� ������ ������ ������ �� ��� ���� �� ����� " & vbNewLine
        sanadan = sanadan & "������� ��� ���� ������� ������� : ��� ������� ����� �� ���� ���� ���� ���� � �������� �� ��� ���� �������� � ����� ��� ������� ���� ����� . ���� ������� �� ����� �� ���� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -" & vbNewLine
        
        ElseIf index = 5 Then
        '����
        sanadan = "��� ����� ������ ��������� ����� ��������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ��� ����: ������ ��� ���� �� ���� �� ��� ������ ���: ����� �� ����� ���: ����� ������� �� ���� �� ��� ������� � ���:����� ��� ���:����� ���� �� ��� � ���: ����� ��� ��� �� ���� � ��� ��� ����: ����� ��� ������� ��� ��� ���� �� ���� ������� � � ��� ��: ���� ��� ��� ��� ����� ��� ������ �� ����� ������� � ����: ���� ��� ������� �� ��� ������ �� ���� ������� �������� ����: ���� ��� ���� �� ����� ������� � ����: ���� ��� ���� �� ���� ��������� � ����: ���� ��� ��� ���� �� ��� �� ��� ��� �� ����." & vbNewLine
        sanadan = sanadan & "���� ����� ��� : ������ ��� ��� ����� ����� �� ����� ������ � ��� : ����� ��� ����� ��� �� ���� �� ���� ������� ������ ������ ������� � ���: ����� ��� ������ ���� �� ��� �������� � ����: ���� ��� ��� ���� ���� �� ������ � ����: ���� ��� ��� � ����: ���� ��� ����� � ���� ��� ���� : ����� ��� ������ ��� ��� ����� ��� ����� ���� ��: ���� ��� ��� ������� ����: ���� ��� �������� �� ���� �� ��� �� ����� . " & vbNewLine
        sanadan = sanadan & "����� ���� ������� ����� ����� : ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ����  � ���� �� ���  � ���� �� ����  � ���� ���� �� ����� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� - � ��� �� �� ���� �� ����� �� �������  � ���� �����  � �� ���� ���� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 6 Then
        '����
        sanadan = "��� ���� ���� ������� �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� : ������ ��� ���� �� ���� � ��� : ����� ��� ����� � ����� ����� �� ��� ������ � ��� : ����� ��� � ���: �� ���� �� ���� � � ��� ����� ������ : ����� ��� ������� ��� ��� ��� ����� ����� � � ��� �� : ���� ��� ��� ��� ����� ���� �� ���� �� ���� ������� ������� � ���� �� : ���� ��� ��� ��� ������ ���� �� ����� �� ���� �� ����� � ���� �� :���� ��� ����� �� ��� ������ ��� �� ����� ������� ��� � ���� �� : ���� ��� ��� � ���� : ���� ��� ���� � � ��� : ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ���� �� ���� � ��� : ����� ���� �� ���� � ��� : ����� ���� �� ���� �� ����� ������ � �� ���� �� ���� �������� � �� ���� � �� ���� � �� ���� � � ��� ����� ������ : ����� ��� ������ ��� ��� ��� ����� ������ ����� � � ��� ��: ���� ��� ��� ��� ���� �� ������ ������� � ���� : ���� ��� ��� ���� �� ���� �� ����� � ���� : ���� ��� ��� ��� ���� �� ����� ������� ������ � ���� :���� ��� ���� ���� : ���� ��� ���� � ���� ���� ��� ����." & vbNewLine
        sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
        sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 7 Then
        '�������
        sanadan = "��� ��� ���� ������� �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� ������ : ������ ��� ��� ���� ��� ������ �� ��� �� ���� ������ � ��� : ����� ��� ��� ��� ���� �� ���� �� ������ ������� � ��� : ����� ���� �� ���� �� ��� ������� � ��� : ����� ��� ��� ������ � �� ������� � � ��� ��� ����� : ����� ��� ������� ��� ��� ��� ����� � ���� �� : ���� ��� ��� ��� ������ �� ����� � ���� : ���� ��� ��� ��� ���� �� ��� �� ������� ������� � � ��� :���� ��� ���� �� ���� � ���� : ���� ��� ��� ��� ������ � ���� : ���� ��� ������� ." & vbNewLine
        sanadan = sanadan & "���� ����� ��� ������ : ������ ��� ���� �� ���� � ��� : ����� ��� ��� ����� � ��� : ����� ���� �� ���� ( ������� ������) � �� ��� ������ � �� ������� � � ��� ����� ������ : ����� ��� ������� ��� ��� ���� �� ���� � � ��� �� : ���� ��� ��� ��� ����� ��� ������ �� ����� ������� � ���� : ���� ��� ��� ��� ������ ��� �� ��� � ���� : ���� ��� ���� �� ����� ������� ������ � ���� :���� ��� ���� �� ���� ( ������� ������) � ���� : ���� ��� ��� ������ � ���� : ���� ��� ������� ." & vbNewLine
        sanadan = sanadan & "����� ������� : ���� �� ���� ������ � ����� �� ��� �������� � ����� �� ��� ���� ������ � ������ �� ����� �������� ��� �� ���� ������ �������� �� ������� �� ���� � ��� ����� ����� ������ ." & vbNewLine
        sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
        sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 8 Then
        '��� ����
        sanadan = "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ����� : ������ ��� ����� ��� ��� ��� �� ����� �� ���� ������� ������� ���� ��� : ������ ��� ����� ��� �� ���� �� ��� ������ ������ ������ �� ������ ��� ����� ��� �� ����� ������ � ��� : ������ ��� ���� ��� ���� �� ��� �������� ������ ������ ��� ����� ��� ������ �� ��� ������ ������� � ������ ��� ��� ���� ���� �� ������ ��������� � ������ ��� ����� ���� �� ���� �� ������� ������ � ������ ��� ��� ���� �� ���� �� ����� ������ � ������ ��� ������ ����� �� ����� �� ���� ������ ������ ��� ����� ���� �� ���� �������� ������� ���� �� ���� ����� � ������ ���� �� �����." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ������ ��� ��� ���� ���� ��� ������ �� ��� ������ � ������� ��� ��� ��� ������ ��� ��� �������� ��� ��� ���� �� ���� �� ��� ������ ������ � ��� : ���� ��� ������ ��� ������ ������� �� ���� �� ���� ������� ��� : ���� ��� ��� ��� ����� ������ � ��� : ���� ��� ��� ������ ��� ����� ���� �� ��� ����� �� ����� �� ����� �������� � ��� : ���� ��� ��� ��� ������ ��� ����� �� ���� ������� � ��� : ���� ��� ��� ��� ���� ���� �� ����� ������ � ��� : ���� ��� ��� ��� ����� ������ ���: ���� ��� ��� ��� ��� �� ����� � ���: ���� ��� ��� ����� �� ����� � ��� : ���� ��� ��� �������� � ��� : ���� ��� ��� ����� � ��� : ���� ��� ��� ��� ����� . " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� : ������ ��� ��� ����� ������� �� ���� �� ������� �� ���� ������� ������� ���� �� ��� ��� ��� �� ����� �� ������ ������� � ������ ��� ����� �� ����� �������� � ������ ��� ���� ��� ������ � ������ ������� ��� ���� ���� �� ������ �� ����� ������� � ������ ������ ��� ������ ���� �� ����� ������ � ������ ��� ��� ����� �� ���� �������� � ������ ��� ����� ��� �� ���� ������� � ������ ��� ��� ���� �� ��� ������ �� ����� ������� � ������ ���� �� ���� �� ����� ������ ������� � ������ ���� �� ��� ���� �� ���� ������� � ������ ��� ������ ���� �� ��� ������ � ������ ��� ����� ���� �� ��� ������ ������ � ������ ���� �� ���� �� ������� �� ���� ��������� � ������ ������ �� ���� �� ��� �� ��� ���� �� ���� ������� � ������ ������� �� ���� �� ��� ���� ������ � ������ ������ �� ���� ��� ����." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ��� ���� �� ��� ������ ������ � ���� ��� ������ ��� ��� ���� �� ���� ������ � ���� ��� ��� ��� ����� �� ���� � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ������ � ���� ��� ��� ������� ��� ���� ���� �� ��� �� ���� ���� �� ���� � ���� ��� ��� ��� ��� ����� �� ����� ��������� � ���� ��� ��� ��� ��� ���� �� ��� ���� �� �������� ��������� � ���� ��� ��� ��� ��� ���� �� ���� �� ��� ������ � ���� ��� ��� ���� �� ���� �� ����� �������� � ���� ��� ��� ���� �� ���� ������ ������� � ���� ��� ��� ��� ���� � ���� ��� ��� ��� ��� ������ � ���� ��� ��� ��� ����� ������ � ���� ��� ��� ��� ���� � ���� ��� ��� ������� � ���� ��� ��� ��� ���� � ���� ��� ��� ��� ���� � ���� ��� ���� � ���� ����� � ��� ��� ���� ." & vbNewLine
        sanadan = sanadan & "������� ��� ���� ����� : ����� ��� ���� �� ���� �� ��� ����� � ���� ����� � ���� ���� . ���� ����� ������� ��� ��� �� ��� � ���� ��� ����� � ���� ���� � ���� ��� ��� �� ���� . ���� ��� �� ����� - ��� ���� ���� � ��� -� �� ����� - ���� ������ -  � �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 9 Then
        '�����
        sanadan = "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ����� ������ ��� ������ ���� �� ���� �� ����� ������ ������� ���� ���: ������ : ��� ������ ���� �� ��� ���� �� ��� ����� ������� ����� ���� � ������ ��� ���� ��� ������ �� ���� �� ������� � �� ����� ������ ��� ��� ��� ���� �� ������ ������ ����� ���� � ������ ��� ���� ���� �� ��� ������� ������� ������ ��� ����� ��� �� ���� �� ��� ������ � ������ ������� ������ ��� ����� ��� �� ���� �� ��� ������� � ������ ��� ������ ��� ���� �� ����� �� ������ ������ � ������ ��� ��� ���� �� ����� �� ���� ������ �������� � ������ ��� ��� ���� ���� �� ������� ������� ����� ." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ������ ��� ���� ��� ������ �� ���� �� ��� �������� � ������� ��� ��� ��� ������ ��� ��� ������ ����� ���� �� ���� ������ � ���� ��� ��� ������� �� ���� ��������� � ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ���� �� ��� �������� � ���� ��� ��� ������� ��� ���� �������� � ���� ��� ��� ��� ��� ����� �� ������ ������� � ���� ��� ��� : ������� � ���� ��� ��� ����� � � ���� ��� ��� ������ � ���� ��� ���� � ���� ��� ��� ����� . " & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ����� ��� ������ ���� �� ���� �� ������ �������� ������� ���� �� ������ ��� ����� ��� �� ���� ������� � ������ ��� ����� ������ ����� � ������ ��� ���� �������� � ������ ��� ����� ������ ����� � ������ ���� �� ������ ������� � ������ ��� ����� ��� �� ���� �� ������� �� ����� ������� ������ ������ ��� ������ ���� �� ����� �� ������ �� ������ ������ � ������ ��� ��� ���� �� ��� �� ���� �� ������ ������ ������ � ������ ��� �� ��� ������ ������ ." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ���� �� ���� �������� �������� � ������� ��� ��� ��� ������ ��� ��� ������ ��� ��� ���� ������ � ���� ��� ��� ��� ����� ������� ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ���� �� ��� � ���� ��� ��� ������� ��� ���� �� ���� � ���� ��� ��� ��� ������ ������� �� ����� �� ���� ������ � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ��� � ������ � ���� ��� ��� ��� ��� � ���� ��� ��� ��� � ���� ��� ��� ����� ." & vbNewLine
        sanadan = sanadan & "������� ����� ����� ����� ����� : ��� ������ ���� �� ������ ������ � ����� �� ����� � ����� �� ����� � ���� ������ ���� �� ���� �������� .���� �� ����� ��� ��� ��� ���� �� ������ ���� ���� ��� ���� ���� ���� � ������� ���� ������� ���� ���� ��� ����� ������ ���� ������ �� ��� ���� ������ ��� ���� ��� � ���� ���� ��� ���� �� ������� ���� ��� ��� ������� ������� ���� ��� ��� ���� ���� ��� ������ ��� ��� ���� ����� �� ����� �������� ���� ��� ��� ������� ������� ���� ��� ���� ���� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 10 Then
        '���
        sanadan = "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "��� ����� ����� ������ : ������ ��� ��� ��� ��� �� ����� ������� ���� ���� ���� � �� ���� ������ ������ ��� ������ ���� �� ������� �� ��� �������� ������� � ��� : ������ ����� � ��� : ������ ��� �������� ������ �� ����� ������� � ������ ��� ���� ���� �� ������ ������� � ������ ��� ������ ���� �� ��� ���� �� ����� ���������� � ������ ��� ����� ���� �� ��� ���� �� ���� �� ��� ������ ������� ���� ��� ��� ������ � ������ ��� ����� ����� �� ������� ������ ." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� �� �� ������� ��� ��� ���� ������ � ���� ���� ������� �������� � ���� �� ����� ��� ��� ��� ���� ���� �� ���� �� ��� ������ ������ � ���� ��� ��� ������ �� ���� � ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ������ ��� ���� �� ���� �� ����� �������� � ���� ��� ��� ��� ��� ���� �� ��� �� ���� ������ � ���� ��� ��� ��� ������ ���������� � ���� ��� ��� ��� ��� ��� ������ � ���� ��� ��� ����� ������ � ���� ��� ��� ��� ." & vbNewLine
        sanadan = sanadan & "���� ����� ����� : ������ ��� ���� �� ���� �� ������ ������� ������� ���� � ������ ��� �� ���� ���� ������ �� ��� ��� �� ����� �������� � ������ ��� ������ �� ���� ������� � ������ ��� ������� �� ��� �� ���� ������ � ������ ��� ����� ��� �� ���� �� ��� ���� ������ � ������ ��� ����� ������� �� ������ �� ��� ���� ������ ������� ������ � ������ ����� �� ��� ������ ������." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ����� ��� ���� ��� ������ �� ���� ������� � ������� ��� ��� ��� ��� ���� �� ���� �� ��� ������ ������ � ���� ��� ��� ������� �� ���� � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ���� ��� ������ � ��� : ����� ��� ������ �� ���� ��� ���� ��� �������� ������ ��� ����� ��� ������ �� ��� ������ ������� � ���� ������� ���� �� ����� �� ������� ������ � ���� ������ ������� ��� ��� ��� ��� ������ ��� ��� ���� ���� �� ������ ��������� � ������ ��� ��� ��� ��� ������ ��� ������ ����� �� ���� �� ���� ������� � ���� ��� ������� ������� ��� ��� ��� ��� ������ ������ ��� ������ ���� �� ��� �� ����� ������� � ���� ������� ��� �� ������ ��� ������ ��� ��� ���� �� ���� �� ����� �� ���� ������� � ���� ������� �������� ����� ��� ����� � ���� ����� ��� ��� � ����� ������ . " & vbNewLine
        sanadan = sanadan & "����� ��� : ����� ��� ���� ���� ���� � ������ �� ����� ������ ���� ��� ��� � ���� ��� ���� ���� �� ��� �������� ���� ������ ����� ����� ������ � ���� ��� ��� � ������� � ����� ��� ���� . ���� ������� ���� �� ������� ��� ���� �� ��� �� ��� ��� � ����� ������ . ��� : ���� ���� �� ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ���� � ���� �� ��� � ���� �� ���� � ���� ���� �� ����� � �� ����� � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
        sanadan = sanadan & "���� �� �� ���� �� ����� �� ������� � ���� ����� � �� ���� ���� � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -. ����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ . ���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 11 Then
        '���
        sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan + "���� ����� ��� : ������ ��� ��� ��� ���� ���� �� ����� ������ ���� � ���: ����� ���� �� ������� �� ���� � ��� : ����� ��� ���� ��� �� ��� � ���: ����� ��� ���� ��� ����� �� ��� ������ � ��� : ����� ��� �� ���� � � ��� ����� ������ : ����� ��� ������� ��� ��� ���� ��� ������ ��� �� ������� �� ���� �� ����� ������� ���� � � ��� �� : ���� ��� ������ ��� ��� ���� ���� �� ����� ������� � ���� �� : ���� ��� ������ ��� ������� �� ��� ���� ������ � ���� : ���� ��� ��� ����� ���� �� ���� �� ���� ������ � ���� :���� ��� ��� ���� : ���� ��� ���� ." & vbNewLine
        sanadan = sanadan + "����� ���� ����� ����� ���� : ��� ���� �� �� �� ������� ������ � ���� ���� ��� ������ �� ���� ������ � ����� �� ���� ������ � ���� ��� ���� ���� �� ���� ������ ����� � ���� ��� ���� �� ����� � ���� ����� ������� �� ��� ����� � ���� ���� � ���� ���� �� ���� �� ��� ����� � �� ��� �� ��� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 21 Then
        '����
        sanadan = "��� ���� ���� ������� �� �������  " & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ��� ���� ���� �� ���� �������� � ���: ����� ��� ����� � ���: ���� ��� ���� � ����: ���� ��� ��� ����� ���� �� ��� ������ ����: ���� ��� ��� ����� �� ��� �� ���� � ����: ���� ��� ������� �� ��� ���� ����� � ����: ���� ��� ��� �� ���� � ����� �� ����� � ����� ����� ��� ��� ������ � � ��� ������� �������� : ����� ��� ������ ��� ��� ���� �� ���� ������ ������� ������ ����: ���� ��� ��� ��� ���� �� ������ �������� � ����: ���� ��� ��� ��� ����� ����: ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & " ������� ��� ���� ������� ����� ����� : ��� ���� �� ������ �������� ���� ���� ����  ������ �� ��� ��� ������ ���� ��� �� ������ � ������ ���� ��� ���� . ���� ��� ���� �� ��� �� ��� ����. ���� ����� �����ӡ �� ��� ���ӡ �� ��� � ���� �� ���� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ -  ��� �� ����� - ����� � ����� -."
        
        ElseIf index = 31 Then
        '������
        sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� ������ : ������ ��� ��� �� ������� �� ���� ������ � ��� : ����� ��� ���� ����� �� ���� ������ � ��� : ����� ��� ��� ������ ���� �� ���� ������� � ��� : ������ ��� ���� � ��� : ������ ������� � �� ��� ���� � ���� ��� ���� : ����� ��� ������ ��� ������ ����� �� ������� ����������� �������� ��� ���� �� ���� ������� � ���� �� : ���� ��� ���� ��� ��� ���� �� ������ ������� � ���� �� : ���� ��� ������ ��� ���� ��� ��� ����� ���� �� ���� ������ � ���� : ���� ��� ��� ��� ���� � ���� : ���� ��� ������� � ���� : ���� ��� ��� ����" & vbNewLine
        sanadan = sanadan & "��� ��� ����: ������ ����� ������� ���� �� ���� �� ��� ����� �� ��� ������ �� ����� �� ������ �� ������� �� ��� ���� ������ ��� ���� ��� ����� ����� � ��� : ����� ��� ���� �� ������� �� ���� �� ������ �� ��� ���� �� ������� �� ��� ���� . " & vbNewLine
        sanadan = sanadan & "����� ��� ���� : ����� �� ��� ������ ��� ��� ������ � ��� ��� ��� : ����� � ����� �� ���� � ������ �� ���� � ����� �� ��� ���� � ���� ���� �� ���� � ����� �� ��� ������ �� ����� � ����� �� ��� ������ ������ � ��� ��� ������� : ���� �� ������� ������� ����� �� ����� � ����� �� ���� � ��� ��� ������ : ����� �� ��� ����� ������ � ���� �� ���� � ������� � ���� ����� ������� ��� ���� �� ������� ������ . " & vbNewLine
        sanadan = sanadan & "��� : ���� ���� �� ���� � �����ɡ ����� �� ���� � �� ��� ���� ���� ��� ���� �� ��� �� ��� ���� �� ���� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -."
        
        ElseIf index = 41 Then
        '��� �����
        sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ����� : ������ ��� ���� �� ���� � ���:����� ���� �� ���� �� ����� � ��� : ����� ���� �� ���� ������� � ��� : ����� ��� ���� �� ����� � ��� : ����� ���� �� ���� ������� � ��� :����� ���� �� ������ ������� � ��� : ���� ��� ��� ���� � ���� ��� ���� : ����� ��� ������ ��� ��� ��� ������ �� ���� ������� ������� ���� �� : ���� ��� ��� ��� ��� ���� �� ����� ������ � ���� : ���� ��� ��� ��� ��� ���� ����� �� ���� �� ���� ������ ������ ������ �� ��� ���� �� ����� " & vbNewLine
        sanadan = sanadan & "������� ��� ���� ������� ������� : ��� ������� ����� �� ���� ���� ���� ���� � �������� �� ��� ���� �������� � ����� ��� ������� ���� ����� . ���� ������� �� ����� �� ���� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -" & vbNewLine
        
        ElseIf index = 51 Then
        '���
        sanadan = "��� ����� ������ ��������� ����� ��������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� : ������ ��� ��� ����� ����� �� ����� ������ � ��� : ����� ��� ����� ��� �� ���� �� ���� ������� ������ ������ ������� � ���: ����� ��� ������ ���� �� ��� �������� � ����: ���� ��� ��� ���� ���� �� ������ � ����: ���� ��� ��� � ����: ���� ��� ����� � ���� ��� ���� : ����� ��� ������ ��� ��� ����� ��� ����� ���� ��: ���� ��� ��� ������� ����: ���� ��� �������� �� ���� �� ��� �� ����� . " & vbNewLine
        sanadan = sanadan & "����� ���� ������� ����� ����� : ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ����  � ���� �� ���  � ���� �� ����  � ���� ���� �� ����� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� - � ��� �� �� ���� �� ����� �� �������  � ���� �����  � �� ���� ���� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 61 Then
        '����
        sanadan = "��� ���� ���� ������� �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ���� �� ���� � ��� : ����� ���� �� ���� � ��� : ����� ���� �� ���� �� ����� ������ � �� ���� �� ���� �������� � �� ���� � �� ���� � �� ���� � � ��� ����� ������ : ����� ��� ������ ��� ��� ��� ����� ������ ����� � � ��� ��: ���� ��� ��� ��� ���� �� ������ ������� � ���� : ���� ��� ��� ���� �� ���� �� ����� � ���� : ���� ��� ��� ��� ���� �� ����� ������� ������ � ���� :���� ��� ���� ���� : ���� ��� ���� � ���� ���� ��� ����." & vbNewLine
        sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
        sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 71 Then
        '��� ������
        sanadan = "��� ��� ���� ������� �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ������ : ������ ��� ���� �� ���� � ��� : ����� ��� ��� ����� � ��� : ����� ���� �� ���� ( ������� ������) � �� ��� ������ � �� ������� � � ��� ����� ������ : ����� ��� ������� ��� ��� ���� �� ���� � � ��� �� : ���� ��� ��� ��� ����� ��� ������ �� ����� ������� � ���� : ���� ��� ��� ��� ������ ��� �� ��� � ���� : ���� ��� ���� �� ����� ������� ������ � ���� :���� ��� ���� �� ���� ( ������� ������) � ���� : ���� ��� ��� ������ � ���� : ���� ��� ������� ." & vbNewLine
        sanadan = sanadan & "����� ������� : ���� �� ���� ������ � ����� �� ��� �������� � ����� �� ��� ���� ������ � ������ �� ����� �������� ��� �� ���� ������ �������� �� ������� �� ���� � ��� ����� ����� ������ ." & vbNewLine
        sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
        sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 81 Then
        '��� ����
        sanadan = "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� : ������ ��� ��� ����� ������� �� ���� �� ������� �� ���� ������� ������� ���� �� ��� ��� ��� �� ����� �� ������ ������� � ������ ��� ����� �� ����� �������� � ������ ��� ���� ��� ������ � ������ ������� ��� ���� ���� �� ������ �� ����� ������� � ������ ������ ��� ������ ���� �� ����� ������ � ������ ��� ��� ����� �� ���� �������� � ������ ��� ����� ��� �� ���� ������� � ������ ��� ��� ���� �� ��� ������ �� ����� ������� � ������ ���� �� ���� �� ����� ������ ������� � ������ ���� �� ��� ���� �� ���� ������� � ������ ��� ������ ���� �� ��� ������ � ������ ��� ����� ���� �� ��� ������ ������ � ������ ���� �� ���� �� ������� �� ���� ��������� � ������ ������ �� ���� �� ��� �� ��� ���� �� ���� ������� � ������ ������� �� ���� �� ��� ���� ������ � ������ ������ �� ���� ��� ����." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ��� ���� �� ��� ������ ������ � ���� ��� ������ ��� ��� ���� �� ���� ������ � ���� ��� ��� ��� ����� �� ���� � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ������ � ���� ��� ��� ������� ��� ���� ���� �� ��� �� ���� ���� �� ���� � ���� ��� ��� ��� ��� ����� �� ����� ��������� � ���� ��� ��� ��� ��� ���� �� ��� ���� �� �������� ��������� � ���� ��� ��� ��� ��� ���� �� ���� �� ��� ������ � ���� ��� ��� ���� �� ���� �� ����� �������� � ���� ��� ��� ���� �� ���� ������ ������� � ���� ��� ��� ��� ���� � ���� ��� ��� ��� ��� ������ � ���� ��� ��� ��� ����� ������ � ���� ��� ��� ��� ���� � ���� ��� ��� ������� � ���� ��� ��� ��� ���� � ���� ��� ��� ��� ���� � ���� ��� ���� � ���� ����� � ��� ��� ���� ." & vbNewLine
        sanadan = sanadan & "������� ��� ���� ����� : ����� ��� ���� �� ���� �� ��� ����� � ���� ����� � ���� ���� . ���� ����� ������� ��� ��� �� ��� � ���� ��� ����� � ���� ���� � ���� ��� ��� �� ���� . ���� ��� �� ����� - ��� ���� ���� � ��� -� �� ����� - ���� ������ -  � �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 91 Then
        '���
        sanadan = "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ����� ��� ������ ���� �� ���� �� ������ �������� ������� ���� �� ������ ��� ����� ��� �� ���� ������� � ������ ��� ����� ������ ����� � ������ ��� ���� �������� � ������ ��� ����� ������ ����� � ������ ���� �� ������ ������� � ������ ��� ����� ��� �� ���� �� ������� �� ����� ������� ������ ������ ��� ������ ���� �� ����� �� ������ �� ������ ������ � ������ ��� ��� ���� �� ��� �� ���� �� ������ ������ ������ � ������ ��� �� ��� ������ ������ ." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ���� �� ���� �������� �������� � ������� ��� ��� ��� ������ ��� ��� ������ ��� ��� ���� ������ � ���� ��� ��� ��� ����� ������� ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ���� �� ��� � ���� ��� ��� ������� ��� ���� �� ���� � ���� ��� ��� ��� ������ ������� �� ����� �� ���� ������ � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ��� � ������ � ���� ��� ��� ��� ��� � ���� ��� ��� ��� � ���� ��� ��� ����� ." & vbNewLine
        sanadan = sanadan & "������� ����� ����� ����� ����� : ��� ������ ���� �� ������ ������ � ����� �� ����� � ����� �� ����� � ���� ������ ���� �� ���� �������� .���� �� ����� ��� ��� ��� ���� �� ������ ���� ���� ��� ���� ���� ���� � ������� ���� ������� ���� ���� ��� ����� ������ ���� ������ �� ��� ���� ������ ��� ���� ��� � ���� ���� ��� ���� �� ������� ���� ��� ��� ������� ������� ���� ��� ��� ���� ���� ��� ������ ��� ��� ���� ����� �� ����� �������� ���� ��� ��� ������� ������� ���� ��� ���� ���� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 101 Then
        '�����
        sanadan = "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ����� : ������ ��� ���� �� ���� �� ������ ������� ������� ���� � ������ ��� �� ���� ���� ������ �� ��� ��� �� ����� �������� � ������ ��� ������ �� ���� ������� � ������ ��� ������� �� ��� �� ���� ������ � ������ ��� ����� ��� �� ���� �� ��� ���� ������ � ������ ��� ����� ������� �� ������ �� ��� ���� ������ ������� ������ � ������ ����� �� ��� ������ ������." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ����� ��� ���� ��� ������ �� ���� ������� � ������� ��� ��� ��� ��� ���� �� ���� �� ��� ������ ������ � ���� ��� ��� ������� �� ���� � ���� ��� ��� ��� ����� � ���� ��� ��� ��� ���� ��� ������ � ��� : ����� ��� ������ �� ���� ��� ���� ��� �������� ������ ��� ����� ��� ������ �� ��� ������ ������� � ���� ������� ���� �� ����� �� ������� ������ � ���� ������ ������� ��� ��� ��� ��� ������ ��� ��� ���� ���� �� ������ ��������� � ������ ��� ��� ��� ��� ������ ��� ������ ����� �� ���� �� ���� ������� � ���� ��� ������� ������� ��� ��� ��� ��� ������ ������ ��� ������ ���� �� ��� �� ����� ������� � ���� ������� ��� �� ������ ��� ������ ��� ��� ���� �� ���� �� ����� �� ���� ������� � ���� ������� �������� ����� ��� ����� � ���� ����� ��� ��� � ����� ������ . " & vbNewLine
        sanadan = sanadan & "����� ��� : ����� ��� ���� ���� ���� � ������ �� ����� ������ ���� ��� ��� � ���� ��� ���� ���� �� ��� �������� ���� ������ ����� ����� ������ � ���� ��� ��� � ������� � ����� ��� ���� . ���� ������� ���� �� ������� ��� ���� �� ��� �� ��� ��� � ����� ������ . ��� : ���� ���� �� ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ���� � ���� �� ��� � ���� �� ���� � ���� ���� �� ����� � �� ����� � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
        sanadan = sanadan & "���� �� �� ���� �� ����� �� ������� � ���� ����� � �� ���� ���� � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -. ����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ . ���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 12 Then
        '�����
        sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan + "��� ����� ����� : ������ ��� ���� �� ��� �� ���� ������ � ���: ����� ���� �� ���� �� ���� � ���: ����� ��� ���� �� ���� ������ � ���:����� ����� �� ���ڡ � ��� ����� ������ : ����� ��� ������� ��� ��� ���� ��� ����� ���� �� ���� �� ���� �� ����� � ������� ������ � � ��� ��: ���� ��� ��� ��� ����� ��� ������ �� ����� ������ � ����: ���� ��� ������� �� ��� ������ � ����: ���� ��� ��� ��� ������ ���� �� ����� �� ���� �� ����� � ����:���� ��� ��� ��� ���� �� ���� �� ������ ����: ���� ��� ��� ���� ���� �� ����� � ����: ���� ��� ����� � ����: ���� ��� ���� ." & vbNewLine
        sanadan = sanadan + "����� ���� ����� ����� ���� : ��� ���� �� �� �� ������� ������ � ���� ���� ��� ������ �� ���� ������ � ����� �� ���� ������ � ���� ��� ���� ���� �� ���� ������ ����� � ���� ��� ���� �� ����� � ���� ����� ������� �� ��� ����� � ���� ���� � ���� ���� �� ���� �� ��� ����� � �� ��� �� ��� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 22 Then
        '�����
        sanadan = "��� ���� ���� ������� �� �������  " & vbNewLine
        sanadan = sanadan & "���� ����� ����� : ������ ��� ���� �� ���� �� ������ � ���:����� ���� �� ���� � ���: ����� ��� �� ���� ����� � ���:����� ���� �� ��� ��� � ���: ���� ��� ����� �� ������ �� ���� � ����: ���� ��� ������� �� ��� ���� ����� � ���� : ���� ��� ��� ���� ���� � ��� ��� ���� : ����� ��� ������� ��� ��� ��� ������ ��� ������ �� ���� �� ���� ������� ������� � ���� ��: ���� ��� ������� ��� ��� ��� ��� ���� �� ����� ������ � ����: ���� ��� ��� ��� ����� ���� �� ����� ��� ��� � ����: ���� ��� ����� ." & vbNewLine
        sanadan = sanadan & " ������� ��� ���� ������� ����� ����� : ��� ���� �� ������ �������� ���� ���� ����  ������ �� ��� ��� ������ ���� ��� �� ������ � ������ ���� ��� ���� . ���� ��� ���� �� ��� �� ��� ����. ���� ����� �����ӡ �� ��� ���ӡ �� ��� � ���� �� ���� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ -  ��� �� ����� - ����� � ����� -."
        
        ElseIf index = 32 Then
        '������
        sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ��� ������ : ������ ��� ���� �� ���� �� ��� � ���: ������ ��� ���� ���� �� ���� �� ��� ��� ���� ���� ��������ɡ ���: ������ ��� ���� ������ �� ���� ���:����� ������� �� ��� ���� � ��� ��� ���� : ����� ��� ������� ��� �� ���� ��� ��� ������ ��� ����� ��� ���� �� �� ���� �� ���� �� ����� �������� ������� ������� � � ��� �� : ���� ��� ��� ��� ���� ��� ������ �� ��� �� ��� ���� ������� � �� �� ����� ���� � ���� �� : ���� ��� ��� ��� ��� �� ����� � ���� : ���� ��� ��� ��� ������� ��� ������ �� ����� ���� :���� ��� ��� ��� � ���� : ���� ��� ��� ������� � ���� ���� ��� ��� : ��� ����. " & vbNewLine
        sanadan = sanadan & "��� ��� ����: ������ ����� ������� ���� �� ���� �� ��� ����� �� ��� ������ �� ����� �� ������ �� ������� �� ��� ���� ������ ��� ���� ��� ����� ����� � ��� : ����� ��� ���� �� ������� �� ���� �� ������ �� ��� ���� �� ������� �� ��� ���� . " & vbNewLine
        sanadan = sanadan & "����� ��� ���� : ����� �� ��� ������ ��� ��� ������ � ��� ��� ��� : ����� � ����� �� ���� � ������ �� ���� � ����� �� ��� ���� � ���� ���� �� ���� � ����� �� ��� ������ �� ����� � ����� �� ��� ������ ������ � ��� ��� ������� : ���� �� ������� ������� ����� �� ����� � ����� �� ���� � ��� ��� ������ : ����� �� ��� ����� ������ � ���� �� ���� � ������� � ���� ����� ������� ��� ���� �� ������� ������ . " & vbNewLine
        sanadan = sanadan & "��� : ���� ���� �� ���� � �����ɡ ����� �� ���� � �� ��� ���� ���� ��� ���� �� ��� �� ��� ���� �� ���� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -."
        
        ElseIf index = 42 Then
        '����
        sanadan = "��� ��� ���� ������ �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ���� �� ���� ���: ����� ��� ����� � ��� : ����� ������ �� ����� ������ � ��� :����� ���� �� ���� �������� � ��� : ����� ���� �� ���� � ���: ����� ���� �� ���� ����� � ��� :���� ��� ���� �� ������ ������� � ����: ���� ��� ��� ���� �� ���� � ��� : ��� ���� : ����� ��� ������� ��� ��� ��� ����� ����� � ����: ���� ��� ��� ��� ���� �� ������ ������� � � ��� : ���� ��� ��� ���� �� ���� �� ����� � ���� : ���� ��� �������� � ���� : ���� ��� ���� " & vbNewLine
        sanadan = sanadan & "������� ��� ���� ������� ������� : ��� ������� ����� �� ���� ���� ���� ���� � �������� �� ��� ���� �������� � ����� ��� ������� ���� ����� . ���� ������� �� ����� �� ���� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -" & vbNewLine
        
        ElseIf index = 52 Then
        '����
        sanadan = "��� ����� ������ ��������� ����� ��������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ��� ����: ������ ��� ���� �� ���� �� ��� ������ ���: ����� �� ����� ���: ����� ������� �� ���� �� ��� ������� � ���:����� ��� ���:����� ���� �� ��� � ���: ����� ��� ��� �� ���� � ��� ��� ����: ����� ��� ������� ��� ��� ���� �� ���� ������� � � ��� ��: ���� ��� ��� ��� ����� ��� ������ �� ����� ������� � ����: ���� ��� ������� �� ��� ������ �� ���� ������� �������� ����: ���� ��� ���� �� ����� ������� � ����: ���� ��� ���� �� ���� ��������� � ����: ���� ��� ��� ���� �� ��� �� ��� ��� �� ����." & vbNewLine
        sanadan = sanadan & "����� ���� ������� ����� ����� : ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ����  � ���� �� ���  � ���� �� ����  � ���� ���� �� ����� � �� ����� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� - � ��� �� �� ���� �� ����� �� �������  � ���� �����  � �� ���� ���� - ��� ���� ���� � ��� - � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 62 Then
        '���
        sanadan = "��� ���� ���� ������� �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� : ������ ��� ���� �� ���� � ��� : ����� ��� ����� � ����� ����� �� ��� ������ � ��� : ����� ��� � ���: �� ���� �� ���� � � ��� ����� ������ : ����� ��� ������� ��� ��� ��� ����� ����� � � ��� �� : ���� ��� ��� ��� ����� ���� �� ���� �� ���� ������� ������� � ���� �� : ���� ��� ��� ��� ������ ���� �� ����� �� ���� �� ����� � ���� �� :���� ��� ����� �� ��� ������ ��� �� ����� ������� ��� � ���� �� : ���� ��� ��� � ���� : ���� ��� ���� � � ��� : ���� ��� ���� ." & vbNewLine
        sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
        sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 72 Then
        '��� ���� ������
        sanadan = "��� ��� ���� ������� �� �������:" & vbNewLine
        sanadan = sanadan & "���� ����� ��� ���� ������ : ������ ��� ��� ���� ��� ������ �� ��� �� ���� ������ � ��� : ����� ��� ��� ��� ���� �� ���� �� ������ ������� � ��� : ����� ���� �� ���� �� ��� ������� � ��� : ����� ��� ��� ������ � �� ������� � � ��� ��� ����� : ����� ��� ������� ��� ��� ��� ����� � ���� �� : ���� ��� ��� ��� ������ �� ����� � ���� : ���� ��� ��� ��� ���� �� ��� �� ������� ������� � � ��� :���� ��� ���� �� ���� � ���� : ���� ��� ��� ��� ������ � ���� : ���� ��� ������� ." & vbNewLine
        sanadan = sanadan & "����� ������� : ���� �� ���� ������ � ����� �� ��� �������� � ����� �� ��� ���� ������ � ������ �� ����� �������� ��� �� ���� ������ �������� �� ������� �� ���� � ��� ����� ����� ������ ." & vbNewLine
        sanadan = sanadan & "����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ ." & vbNewLine
        sanadan = sanadan & "���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� - ��� ���� ���� � ��� - �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 82 Then
        '��� �����
        sanadan = "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ��� ����� : ������ ��� ����� ��� ��� ��� �� ����� �� ���� ������� ������� ���� ��� : ������ ��� ����� ��� �� ���� �� ��� ������ ������ ������ �� ������ ��� ����� ��� �� ����� ������ � ��� : ������ ��� ���� ��� ���� �� ��� �������� ������ ������ ��� ����� ��� ������ �� ��� ������ ������� � ������ ��� ��� ���� ���� �� ������ ��������� � ������ ��� ����� ���� �� ���� �� ������� ������ � ������ ��� ��� ���� �� ���� �� ����� ������ � ������ ��� ������ ����� �� ����� �� ���� ������ ������ ��� ����� ���� �� ���� �������� ������� ���� �� ���� ����� � ������ ���� �� �����." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ������ ��� ��� ���� ���� ��� ������ �� ��� ������ � ������� ��� ��� ��� ������ ��� ��� �������� ��� ��� ���� �� ���� �� ��� ������ ������ � ��� : ���� ��� ������ ��� ������ ������� �� ���� �� ���� ������� ��� : ���� ��� ��� ��� ����� ������ � ��� : ���� ��� ��� ������ ��� ����� ���� �� ��� ����� �� ����� �� ����� �������� � ��� : ���� ��� ��� ��� ������ ��� ����� �� ���� ������� � ��� : ���� ��� ��� ��� ���� ���� �� ����� ������ � ��� : ���� ��� ��� ��� ����� ������ ���: ���� ��� ��� ��� ��� �� ����� � ���: ���� ��� ��� ����� �� ����� � ��� : ���� ��� ��� �������� � ��� : ���� ��� ��� ����� � ��� : ���� ��� ��� ��� ����� . " & vbNewLine
        sanadan = sanadan & "������� ��� ���� ����� : ����� ��� ���� �� ���� �� ��� ����� � ���� ����� � ���� ���� . ���� ����� ������� ��� ��� �� ��� � ���� ��� ����� � ���� ���� � ���� ��� ��� �� ���� . ���� ��� �� ����� - ��� ���� ���� � ��� -� �� ����� - ���� ������ -  � �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 92 Then
        '����
        sanadan = "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "���� ����� ���� : ������ ��� ����� ������ ��� ������ ���� �� ���� �� ����� ������ ������� ���� ���: ������ : ��� ������ ���� �� ��� ���� �� ��� ����� ������� ����� ���� � ������ ��� ���� ��� ������ �� ���� �� ������� � �� ����� ������ ��� ��� ��� ���� �� ������ ������ ����� ���� � ������ ��� ���� ���� �� ��� ������� ������� ������ ��� ����� ��� �� ���� �� ��� ������ � ������ ������� ������ ��� ����� ��� �� ���� �� ��� ������� � ������ ��� ������ ��� ���� �� ����� �� ������ ������ � ������ ��� ��� ���� �� ����� �� ���� ������ �������� � ������ ��� ��� ���� ���� �� ������� ������� ����� ." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� ������ ��� ���� ��� ������ �� ���� �� ��� �������� � ������� ��� ��� ��� ������ ��� ��� ������ ����� ���� �� ���� ������ � ���� ��� ��� ������� �� ���� ��������� � ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ���� �� ��� �������� � ���� ��� ��� ������� ��� ���� �������� � ���� ��� ��� ��� ��� ����� �� ������ ������� � ���� ��� ��� : ������� � ���� ��� ��� ����� � � ���� ��� ��� ������ � ���� ��� ���� � ���� ��� ��� ����� . " & vbNewLine
        sanadan = sanadan & "������� ����� ����� ����� ����� : ��� ������ ���� �� ������ ������ � ����� �� ����� � ����� �� ����� � ���� ������ ���� �� ���� �������� .���� �� ����� ��� ��� ��� ���� �� ������ ���� ���� ��� ���� ���� ���� � ������� ���� ������� ���� ���� ��� ����� ������ ���� ������ �� ��� ���� ������ ��� ���� ��� � ���� ���� ��� ���� �� ������� ���� ��� ��� ������� ������� ���� ��� ��� ���� ���� ��� ������ ��� ��� ���� ����� �� ����� �������� ���� ��� ��� ������� ������� ���� ��� ���� ���� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        
        ElseIf index = 102 Then
        '������
        sanadan = "��� ������ ��� ����� ���� �� ������ �� ����� ������� : " & vbNewLine
        sanadan = sanadan & "��� ����� ����� ������ : ������ ��� ��� ��� ��� �� ����� ������� ���� ���� ���� � �� ���� ������ ������ ��� ������ ���� �� ������� �� ��� �������� ������� � ��� : ������ ����� � ��� : ������ ��� �������� ������ �� ����� ������� � ������ ��� ���� ���� �� ������ ������� � ������ ��� ������ ���� �� ��� ���� �� ����� ���������� � ������ ��� ����� ���� �� ��� ���� �� ���� �� ��� ������ ������� ���� ��� ��� ������ � ������ ��� ����� ����� �� ������� ������ ." & vbNewLine
        sanadan = sanadan & "��� ��� ������ : ����� ��� ������� ��� ��� �� �� ������� ��� ��� ���� ������ � ���� ���� ������� �������� � ���� �� ����� ��� ��� ��� ���� ���� �� ���� �� ��� ������ ������ � ���� ��� ��� ������ �� ���� � ���� ��� ��� ��� �� ����� � ���� ��� ��� ��� ������ ��� ���� �� ���� �� ����� �������� � ���� ��� ��� ��� ��� ���� �� ��� �� ���� ������ � ���� ��� ��� ��� ������ ���������� � ���� ��� ��� ��� ��� ��� ������ � ���� ��� ��� ����� ������ � ���� ��� ��� ��� ." & vbNewLine
        sanadan = sanadan & "����� ��� : ����� ��� ���� ���� ���� � ������ �� ����� ������ ���� ��� ��� � ���� ��� ���� ���� �� ��� �������� ���� ������ ����� ����� ������ � ���� ��� ��� � ������� � ����� ��� ���� . ���� ������� ���� �� ������� ��� ���� �� ��� �� ��� ��� � ����� ������ . ��� : ���� ���� �� ��� ��� ������ ��� ���� �� ���� ������ � ���� ���� �� �� ���� � ����� ��� ��� ������ �� ����� �� ���� � ���� �� ��� ���� � ���� �� ��� � ���� �� ���� � ���� ���� �� ����� � �� ����� � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -." & vbNewLine
        sanadan = sanadan & "���� �� �� ���� �� ����� �� ������� � ���� ����� � �� ���� ���� � �� ����� - ���� ������ - � �� �� ����� - ����� � ����� -. ����� ���� : ����� ���� ����� ���� ��� ���� ������ �� ����� ������ � ����� �� ��� ������ �� ��� ���� ������ � ������ �� ���� � ���� ����� ������� � ������ ��� ������� � ������ �� ���� � ����� �� ���� ������ � ������ . ���� ������ �� ���� �� ���� � ���� ���� �� ����� �� ����� ��� ����� : ����� � ������� ����� �� ���� ������� � ��� �� ���� � ���� ��� ������ ������ � ������ � �� ��� ����� �� ����� �� ����� - ���� ������ -  �� �� ����� - ����� � ����� -." & vbNewLine
        Else
        sanadan = "sanada"
    End If

End Function
Function qeraatn(index As Integer) As String

        'adding sanad
        If index = -1 Then
        qeraatn = "������� ��� ������ ( ��� ���� � ���� � ������� � ��� )"
          
        ElseIf index = -2 Then
        qeraatn = "������ �������� ( ��� ���� � ����� ) "
      
        ElseIf index = -3 Then
        qeraatn = "��������� ����� ������"
      
        ElseIf index = -4 Then
        qeraatn = "������� ��� �����"
       
        ElseIf index = -5 Then
        qeraatn = "��������� �����"
       
        ElseIf index = 1 Then
        qeraatn = "������ ������ ���� �������"
        
        ElseIf index = 3 Then
        qeraatn = "������ ������ ��� ���� ������ �������"
        
        ElseIf index = 4 Then
        qeraatn = "������ ������ ��� ���� �������"
        
        ElseIf index = 5 Then
        qeraatn = "������ ������ ���� �������"
        
        ElseIf index = 6 Then
        qeraatn = "������ ������ ���� �������"
        
        ElseIf index = 7 Then
        qeraatn = "������ ������ ������� �������"
        
        ElseIf index = 8 Then
        qeraatn = "������ ������ ��� ���� �������"
        
        ElseIf index = 9 Then
        qeraatn = "������ ������ ����� �������"
        
        ElseIf index = 10 Then
        qeraatn = "������ ������ ��� ������ �������"
        
        ElseIf index = 11 Then
        qeraatn = "������ ��� �� ����"
        
        ElseIf index = 21 Then
        qeraatn = "������ ���� �� ��� ����"
        
        ElseIf index = 31 Then
        qeraatn = "������ ������ �� ��� ���� ������"
        
        ElseIf index = 41 Then
        qeraatn = "������ ��� ����� �� ��� ����"
        
        ElseIf index = 51 Then
        qeraatn = "������ ��� �� ����"
        
        ElseIf index = 61 Then
        qeraatn = "������ ���� �� ����"
        
        ElseIf index = 71 Then
        qeraatn = "������ ��� ������ �� �������"
        
        ElseIf index = 81 Then
        qeraatn = "������ ��� ���� �� ��� ����"
        
        ElseIf index = 91 Then
        qeraatn = "������ ��� �� �����"
        
        ElseIf index = 101 Then
        qeraatn = "������ ����� �� ��� ������"
        
        ElseIf index = 12 Then
        qeraatn = "������ ����� �� ����"
        
        ElseIf index = 22 Then
        qeraatn = "������ ����� �� ��� ����"
        
        ElseIf index = 32 Then
        qeraatn = "������ ������ �� ��� ���� ������"
        
        ElseIf index = 42 Then
        qeraatn = "������ ���� �� ��� ����"
        
        ElseIf index = 52 Then
        qeraatn = "������ ���� �� ����"
        
        ElseIf index = 62 Then
        qeraatn = "������ ��� �� ����"
        
        ElseIf index = 72 Then
        qeraatn = "������ ��� ���� ������ �� �������"
        
        ElseIf index = 82 Then
        qeraatn = "������ ��� ����� �� ��� ����"
        
        ElseIf index = 92 Then
        qeraatn = "������ ���� �� �����"
        
        ElseIf index = 102 Then
        qeraatn = "������ ������ �� ��� ������"
        Else
        qeraatn = "egaza_content"
    End If

End Function
Public Function rawye(index As Integer) As String

     'adding sanad
        If index = -1 Then
        rawye = "��� ������ / ��� ������"
        
        ElseIf index = -2 Then
        rawye = "��� ������ / ��������"
        
        ElseIf index = -3 Then
        rawye = "��� �������� �����"
         
        ElseIf index = -4 Then
        rawye = "��� ������ ��� �����"
          
        ElseIf index = -5 Then
        rawye = "��� �������� �����"
          
        ElseIf index = 1 Then
        rawye = "��� ����� ������ / ����"
        
        ElseIf index = 2 Then
        rawye = "��� ����� ������ / ��� ����"
        
        ElseIf index = 3 Then
        rawye = "��� ����� ������ / ��� ���� ������"
        
        ElseIf index = 4 Then
        rawye = "��� ����� ������ / ��� ����"
        
        ElseIf index = 5 Then
        rawye = "��� ����� ������ / ����"
        
        ElseIf index = 6 Then
        rawye = "��� ����� ������ / ����"
        
        ElseIf index = 7 Then
        rawye = "��� ����� ������ / �������"
        
        ElseIf index = 8 Then
        rawye = "��� ����� ������ / ��� ����"
        
        ElseIf index = 9 Then
        rawye = "��� ����� ������ / �����"
        
        ElseIf index = 10 Then
        rawye = "��� ����� ������ / ��� ������"
        
        ElseIf index = 11 Then
        rawye = "��� ����� / ���"
        
        ElseIf index = 21 Then
        rawye = "��� ����� / ����"
        
        ElseIf index = 31 Then
        rawye = "��� ����� / ������"
        
        ElseIf index = 41 Then
        rawye = "��� ����� / ��� �����"
        
        ElseIf index = 51 Then
        rawye = "��� ����� / ���"
        
        ElseIf index = 61 Then
        rawye = "��� ����� / ����"
        
        ElseIf index = 71 Then
        rawye = "��� ����� / ��� ������"
        
        ElseIf index = 81 Then
        rawye = "��� ����� / ��� ����"
        
        ElseIf index = 91 Then
        rawye = "��� ����� / ���"
        
        ElseIf index = 101 Then
        rawye = "��� ����� / �����"
        
        ElseIf index = 12 Then
        rawye = "��� ����� / �����"
        
        ElseIf index = 22 Then
        rawye = "��� ����� / �����"
        
        ElseIf index = 32 Then
        rawye = "��� ����� / ������"
        
        ElseIf index = 42 Then
        rawye = "��� ����� / ����"
        
        ElseIf index = 52 Then
        rawye = "��� ����� / ����"
        
        ElseIf index = 62 Then
        rawye = "��� ����� / ���"
        
        ElseIf index = 72 Then
        rawye = "��� ����� / ��� ����"
        
        ElseIf index = 82 Then
        rawye = "��� ����� / ��� �����"
        
        ElseIf index = 92 Then
        rawye = "��� ����� / ����"
        
        ElseIf index = 102 Then
        rawye = "��� ����� / ������"
        Else
        rawye = "rawy"
    End If

End Function
Public Function get_obydi() As Integer
    If OptionButton9.Value = True Then
        get_obydi = 1
    ElseIf OptionButton10.Value = True Then
        get_obydi = 2
    ElseIf OptionButton11.Value = True Then
        get_obydi = 3
    Else
        get_obydi = 4
    End If
End Function
Public Function get_sheikh_type() As Integer
    If OptionButton4.Value = True Then
        ' female
        get_sheikh_type = 1
    Else
        get_sheikh_type = -1
    End If
End Function
Public Function get_student_type() As Boolean
    If OptionButton6.Value = True Then
        'female
        get_student_type = False
    Else
        get_student_type = True
    End If
End Function
Public Function get_status() As String
 ' set egaza status
    If CheckBox39.Value = True Then
        get_status = "�������"
    End If
    
    If CheckBox40.Value = True Then
        get_status = "��� ������"
    Else
        get_status = "���� �����"
    End If
       
    If CheckBox41.Value = True Then
        get_status = get_status + " " + "���� �� ������"
    Else
        get_status = get_status + " " + "���� �� ��� ���"
    End If
   
End Function
Public Function get_index() As Integer
 ' set index

    If CheckBox38.Value = True Then
        ' �����
        get_index = -5
    End If
    
    If CheckBox6.Value = True Then
        ' ��� �����
        get_index = -4
    End If
    
    If CheckBox37.Value = True Then
         ' �����
         get_index = -3
    End If
    
    If CheckBox42.Value = True Then
         ' ��������
         get_index = -2
    End If
    
    If CheckBox5.Value = True Then
        ' ������
        get_index = -1
    End If
    
    If CheckBox7.Value = True Then
        '����
        get_index = 1
    End If
    
    If CheckBox8.Value = True Then
        '��� ����
        get_index = 2
    End If
   
    If CheckBox9.Value = True Then
        '��� ����
        get_index = 3
    End If
   
    If CheckBox10.Value = True Then
       '��� ����
        get_index = 4
    End If
     
    If CheckBox11.Value = True Then
       '����
        get_index = 5
    End If
     
    If CheckBox12.Value = True Then
       '����
        get_index = 6
    End If
     
    If CheckBox13.Value = True Then
       '�������
        get_index = 7
    End If
     
    If CheckBox14.Value = True Then
        '��� ����
         get_index = 8
    End If
   
    If CheckBox15.Value = True Then
       '�����
        get_index = 9
    End If
     
    If CheckBox16.Value = True Then
        '���
         get_index = 10
    End If
   
    ' set Rowayat
    If CheckBox17.Value = True Then
        '���
        get_index = 11
    End If
   
    If CheckBox18.Value = True Then
        '�����
        get_index = 12
    End If
    
    If CheckBox19.Value = True Then
        '����
         get_index = 21
    End If
     
    If CheckBox20.Value = True Then
        '�����
         get_index = 22
    End If
     
    If CheckBox21.Value = True Then
        '������
         get_index = 31
    End If
    
    If CheckBox22.Value = True Then
       '������
       get_index = 32
    End If
     
    If CheckBox23.Value = True Then
     '��� �����
     get_index = 41
    End If
    
    If CheckBox24.Value = True Then
      '���� �� ��� ����
      get_index = 42
    End If
     
    If CheckBox25.Value = True Then
     '���
     get_index = 51
    End If
     
    If CheckBox26.Value = True Then
    '����
    get_index = 52
    End If
   
    If CheckBox27.Value = True Then
     '����
     get_index = 61
    End If
     
    If CheckBox28.Value = True Then
      '���
      get_index = 62
    End If
     
    If CheckBox29.Value = True Then
       '��� ������
       get_index = 71
    End If
     
    If CheckBox30.Value = True Then
        '������ �� �������
        get_index = 72
    End If
     
    If CheckBox31.Value = True Then
    '��� ����
    get_index = 81
    End If
     
    If CheckBox32.Value = True Then
     '��� �����
     get_index = 82
    End If
     
    If CheckBox33.Value = True Then
      '���
      get_index = 91
    End If
     
    If CheckBox34.Value = True Then
       '����
       get_index = 92
    End If
     
    If CheckBox35.Value = True Then
       '�����
         get_index = 101
    End If
     
    If CheckBox36.Value = True Then
        '������
         get_index = 102
    End If

End Function

Public Function get_special_index(QERAA As String) As Integer
 ' set index

    If InStr(QERAA, "�����") > 0 Then
        '�����
        get_special_index = -5
    End If
    
    If InStr(QERAA, "��� �����") > 0 Then
        ' ��� �����
        get_special_index = -4
    End If
    
   If InStr(QERAA, "�����") > 0 Then
          ' �����
         get_special_index = -3
    End If
    
   If InStr(QERAA, "��������") > 0 Then
          ' ��������
         get_special_index = -2
    End If
    
   If InStr(QERAA, "������") > 0 Then
         ' ������
        get_special_index = -1
    End If
    
   If InStr(QERAA, "����") > 0 Then
         '����
        get_special_index = 1
    End If
    
   If InStr(QERAA, "��� ����") > 0 Then
         '��� ����
        get_special_index = 2
    End If
   
   If InStr(QERAA, "��� ����") > 0 Then
         '��� ����
        get_special_index = 3
    End If
   
   If InStr(QERAA, "��� ����") > 0 Then
        '��� ����
        get_special_index = 4
    End If
     
   If InStr(QERAA, "����") > 0 Then
        '����
        get_special_index = 5
    End If
     
   If InStr(QERAA, "����") > 0 Then
        '����
        get_special_index = 6
    End If
     
   If InStr(QERAA, "�������") > 0 Then
        '�������
        get_special_index = 7
    End If
     
   If InStr(QERAA, "��� ����") > 0 Then
         '��� ����
         get_special_index = 8
    End If
   
   If InStr(QERAA, "�����") > 0 Then
        '�����
        get_special_index = 9
    End If
     
   If InStr(QERAA, "��� ������") > 0 Then
         '��� ������
         get_special_index = 10
    End If
   
    ' set Rowayat
   If InStr(QERAA, "���") > 0 Then
         '���
        get_special_index = 11
    End If
   
   If InStr(QERAA, "�����") > 0 Then
         '�����
        get_special_index = 12
    End If
    
   If InStr(QERAA, "����") > 0 Then
         '����
         get_special_index = 21
    End If
     
   If InStr(QERAA, "�����") > 0 Then
         '�����
         get_special_index = 22
    End If
     
   If InStr(QERAA, "������") > 0 Then
         '������
         get_special_index = 31
    End If
    
   If InStr(QERAA, "������") > 0 Then
        '������
       get_special_index = 32
    End If
     
   If InStr(QERAA, "��� �����") > 0 Then
      '��� �����
     get_special_index = 41
    End If
    
   If InStr(QERAA, "����") > 0 Then
       '���� �� ��� ����
      get_special_index = 42
    End If
     
   If InStr(QERAA, "���") > 0 Then
      '���
     get_special_index = 51
    End If
     
   If InStr(QERAA, "����") > 0 Then
     '����
    get_special_index = 52
    End If
   
   If InStr(QERAA, "����") > 0 Then
      '����
     get_special_index = 61
    End If
     
   If InStr(QERAA, "���") > 0 Then
       '���
      get_special_index = 62
    End If
     
   If InStr(QERAA, "��� ������") > 0 Then
        '��� ������
       get_special_index = 71
    End If
     
   If InStr(QERAA, "������ �� �������") > 0 Then
         '������ �� �������
        get_special_index = 72
    End If
     
   If InStr(QERAA, "��� ����") > 0 Then
     '��� ����
    get_special_index = 81
    End If
     
   If InStr(QERAA, "��� �����") > 0 Then
      '��� �����
     get_special_index = 82
    End If
     
   If InStr(QERAA, "���") > 0 Then
       '���
      get_special_index = 91
    End If
     
   If InStr(QERAA, "����") > 0 Then
        '����
       get_special_index = 92
    End If
     
   If InStr(QERAA, "�����") > 0 Then
        '�����
         get_special_index = 101
    End If
     
   If InStr(QERAA, "������") > 0 Then
         '������
         get_special_index = 102
    End If

End Function

Public Function get_tareq() As String
    get_tareq = " �� ���� "
    If CheckBox3.Value = True Then
     
        If CheckBox14.Value = True Or CheckBox15.Value = True Or CheckBox16.Value = True Or CheckBox31.Value = True Or CheckBox32.Value = True Or CheckBox33.Value = True Or CheckBox34.Value = True Or CheckBox35.Value = True Or CheckBox36.Value = True Then
            get_tareq = get_tareq + "�����"
        Else
            get_tareq = get_tareq + "��������"
        End If
        
        If CheckBox37.Value = True Or CheckBox42.Value = True Or CheckBox6.Value = True Or CheckBox5.Value = True Then
            get_tareq = " �� ���� �������� � �����"
        End If
        
     End If
     
     If CheckBox4.Value = True And CheckBox3.Value = True Then
         get_tareq = get_tareq + " � ������"
     ElseIf CheckBox4.Value = True Then
         get_tareq = get_tareq + "������"
     End If

End Function
Private Sub removeBreakLines()

End Sub
Private Sub CommandButton1_Click()

    Dim index As Integer
    Dim obydi As Integer
    Dim sheikh_type As Integer
    Dim student_type As Boolean
    
    Dim sheikh_name As String
    Dim sheikh_info As String
    Dim student_name As String
    Dim student_info As String
      
    Dim Rng As Range, iPage As Long
    Dim status As String
    Dim qeraat As String
    Dim TAREQ As String
    Dim rawy As String
    Dim sanada As String
     
    sheikh_name = TextBox1.text
    student_name = TextBox2.text
    sheikh_info = TextBox3.text
    student_info = TextBox4.text
   
    obydi = 4
    sheikh_type = get_sheikh_type()
    student_type = get_student_type()
    status = get_status()
    index = get_index()
    
    ' make numbers arabic
    Options.ArabicNumeral = wdNumeralHindi
    set_sheikh_and_student sheikh_name:=sheikh_name, sheikh_info:=sheikh_info, student_name:=student_name, student_info:=student_info
    set_types sheikh_type:=sheikh_type, student_type:=student_type

    If index <> 0 Then
        
        TAREQ = get_tareq()
        sanada = sanadan(index)
        rawy = rawye(index)
        qeraat = qeraatn(index)
        qeraat = qeraat + TAREQ
        rawy = rawy + TAREQ
        
        set_qeraat STATE:=status, qeraat:=qeraat, rawy:=rawy
        set_snada (sanada)
        
        Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, NAME:="1"

    End If

    Dim tempForm As UserForm1
    For Each tempForm In UserForms
        Unload tempForm
    Next
    
    
End Sub
Private Sub moveToBack()
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    ShowVisualBasicEditor = True
End Sub
Private Sub add_image(imgPath)
 Dim pic As Shape
 Set pic = ActiveDocument.Shapes.AddPicture(FileName:=imgPath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=-62, _
        Top:=-38, _
        Width:=595, _
        Height:=842, _
        Anchor:=Selection.Range)
        pic.WrapFormat.Type = wdWrapNone
End Sub
Private Sub CommandButton2_Click()
        
    moveToBack
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "All Files", "*.*"

       If .Show = True Then
        add_image (.SelectedItems(1))
      End If
   End With
End Sub

Private Sub CommandButton3_Click()
    
    Dim temp As Integer
    Dim index As Integer
    Dim obydi As Integer
    Dim sheikh_type As Integer
    Dim student_type As Boolean
    
    Dim sheikh_name As String
    Dim sheikh_info As String
    Dim student_name As String
    Dim student_info As String
      
    Dim Rng As Range, iPage As Long
    Dim status As String
    Dim qeraat As String
    Dim TAREQ As String
    Dim rawy As String
    Dim sanada As String
     
    Dim originalFilePath As String
    Dim dlgOpen As FileDialog
    Dim IndexArray(30) As Integer
    Dim loop_counter As Integer
    
    IndexArray(1) = 1
    IndexArray(2) = 2
    IndexArray(3) = 3
    IndexArray(4) = 4
    IndexArray(5) = 5
    IndexArray(6) = 6
    IndexArray(7) = 7
    IndexArray(8) = 8
    IndexArray(9) = 9
    IndexArray(10) = 10
    IndexArray(11) = 11
    IndexArray(12) = 21
    IndexArray(13) = 31
    IndexArray(14) = 41
    IndexArray(15) = 51
    IndexArray(16) = 61
    IndexArray(17) = 71
    IndexArray(18) = 81
    IndexArray(19) = 91
    IndexArray(20) = 101
    IndexArray(21) = 12
    IndexArray(22) = 22
    IndexArray(23) = 32
    IndexArray(24) = 42
    IndexArray(25) = 52
    IndexArray(26) = 62
    IndexArray(27) = 72
    IndexArray(28) = 82
    IndexArray(29) = 92
    IndexArray(30) = 102
    
    loop_counter = 1
    
temp = MsgBox("Start group!", vbQuestion + vbYesNo, "Confirm")

If temp = 6 Then
   
    Set dlgOpen = Application.FileDialog(FileDialogType:=msoFileDialogOpen)
    With dlgOpen
    .AllowMultiSelect = False
    .Show
    End With
    originalFilePath = dlgOpen.SelectedItems(1)
          
    sheikh_name = TextBox1.text
    student_name = TextBox2.text
    sheikh_info = TextBox3.text
    student_info = TextBox4.text
     
    obydi = 4
    sheikh_type = get_sheikh_type()
    student_type = get_student_type()
    status = get_status()
    
    Dim wdApp As Word.Application
    Set wdApp = GetObject(, "Word.Application")
                   
  While loop_counter <= 30
    
     index = IndexArray(loop_counter)
     TAREQ = get_tareq()
     sanada = sanadan(index)
     rawy = rawye(index)
     qeraat = qeraatn(index)
     qeraat = qeraat + TAREQ
     rawy = rawy + TAREQ
 
     Documents.Open FileName:=originalFilePath, ReadOnly:=False
   
     ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path + Application.PathSeparator + Replace(rawy, "/", "") + ".docx", FileFormat:= _
     wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
     :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
     :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
     SaveAsAOCELetter:=False, CompatibilityMode:=14
        
    ' make numbers arabic
     Options.ArabicNumeral = wdNumeralHindi
     set_sheikh_and_student sheikh_name:=sheikh_name, sheikh_info:=sheikh_info, student_name:=student_name, student_info:=student_info
     set_types sheikh_type:=sheikh_type, student_type:=student_type
     set_qeraat STATE:=status, qeraat:=qeraat, rawy:=rawy
     set_snada (sanada)
        
     ActiveDocument.Save
     wdApp.Documents(ActiveDocument.Path + Application.PathSeparator + Replace(rawy, "/", "") + ".docx").Close
     
     loop_counter = loop_counter + 1
     
  Wend
       
    Dim tempForm As UserForm
    For Each tempForm In UserForms
        Unload tempForm
    Next

End If

End Sub

Private Sub CommandButton4_Click()

    Dim students As String
    Dim substrings() As String
    Dim counter As Integer
    
    Dim originalFilePath As String
    Dim dlgOpen As FileDialog
          
    Set dlgOpen = Application.FileDialog(FileDialogType:=msoFileDialogOpen)
    With dlgOpen
        .AllowMultiSelect = False
        .Show
    End With
    
    originalFilePath = dlgOpen.SelectedItems(1)
            
    Dim wdApp As Word.Application
    Set wdApp = GetObject(, "Word.Application")
         
    students = TextBox5.text
    substrings = Strings.Split(students, vbNewLine)
    counter = Val(substrings(0))
    
    For k = 0 To counter - 1
     
        Dim index As Integer
        Dim obydi As Integer
        Dim sheikh_type As Integer
        Dim student_type As Boolean
        
        Dim sheikh_name As String
        Dim sheikh_info As String
        Dim student_name As String
        Dim student_info As String
          
        Dim Rng As Range, iPage As Long
        Dim status As String
        Dim qeraat As String
        Dim TAREQ As String
        Dim rawy As String
        Dim sanada As String
          
        sheikh_name = TextBox1.text
        sheikh_info = TextBox3.text
        student_name = (substrings(1 + (k * 4)))
        student_info = (substrings(2 + (k * 4)))
        
        obydi = 4
        sheikh_type = get_sheikh_type()
        status = get_status()
      
        If (substrings(3 + (k * 4))) = "����" Then
        student_type = True
        Else
        student_type = False
        End If
        
        ' make numbers arabic
        Options.ArabicNumeral = wdNumeralHindi
          
        index = Val(substrings(4 + (k * 4)))
         
        If index <> 0 Then
            
            TAREQ = get_tareq()
            sanada = sanadan(index)
            rawy = rawye(index)
            qeraat = qeraatn(index)
            qeraat = qeraat + TAREQ
            rawy = rawy + TAREQ
                 
            Documents.Open FileName:=originalFilePath, ReadOnly:=False
            
            ActiveDocument.SaveAs2 FileName:=ActiveDocument.Path + Application.PathSeparator + student_name + ".docx", FileFormat:= _
            wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
            :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
            :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
            SaveAsAOCELetter:=False, CompatibilityMode:=14
      
            set_sheikh_and_student sheikh_name:=sheikh_name, sheikh_info:=sheikh_info, student_name:=student_name, student_info:=student_info
            set_types sheikh_type:=sheikh_type, student_type:=student_type
            set_qeraat STATE:=status, qeraat:=qeraat, rawy:=rawy
            set_snada (sanada)
            
            ActiveDocument.Save
            wdApp.Documents(ActiveDocument.Path + Application.PathSeparator + student_name + ".docx").Close

        End If

    Next k
End Sub

Private Sub CommandButton5_Click()

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim strDB As String
   Dim strSQL As String
   Dim ejaza_id As Integer
   
   Dim NAME As String
   Dim INFO As String
   Dim QERAA As String
   Dim TAREk As String
   Dim STATE As String
   Dim GENDER As Boolean
   
   strDB = "E:\\other\\otor.accdb"
   Set db = OpenDatabase(strDB)
   
   
   ejaza_id = InputBox("���� ���� �������")
   strSQL = "Select * from EJAZA where ID = " & ejaza_id
   Set rst = db.OpenRecordset(strSQL)
       
   If rst.RecordCount > 0 Then
      NAME = rst.Fields("STUDENT_NAME")
      INFO = rst.Fields("STUDENT_INFO")
      QERAA = rst.Fields("QERAA")
      TAREk = rst.Fields("TAREQ")
      STATE = rst.Fields("STATE")
      If (rst.Fields("STUDENT_GENDER") = "����") Then
        GENDER = True
      Else
        GENDER = False
      End If
      
      MsgBox (NAME & vbNewLine & INFO & vbNewLine & rst.Fields("STUDENT_GENDER") & vbNewLine & QERAA & vbNewLine & TAREk & vbNewLine & STATE)
   Else
      MsgBox ("���� ������� ��� �����")
   End If
   
   rst.Close
   db.Close
   Set db = Nothing
   Set rst = Nothing
   
    Dim obydi As Integer
    Dim sheikh_type As Integer
    Dim sheikh_name As String
    Dim sheikh_info As String
        
    Dim Rng As Range, iPage As Long
    Dim rawy As String
    Dim sanada As String
         
    Dim index As Integer
    Dim student_type As Boolean
    Dim student_name As String
    Dim student_info As String
    Dim status As String
    Dim qeraat As String
    Dim TAREQ As String
       
      
    sheikh_name = TextBox1.text
    sheikh_info = TextBox3.text
    obydi = 4
    sheikh_type = get_sheikh_type()
     
    student_name = NAME
    student_info = INFO
    student_type = GENDER
    
    ' make numbers arabic
    Options.ArabicNumeral = wdNumeralHindi
    set_sheikh_and_student sheikh_name:=sheikh_name, sheikh_info:=sheikh_info, student_name:=student_name, student_info:=student_info
    set_types sheikh_type:=sheikh_type, student_type:=student_type
     
     
    index = get_special_index(QERAA)
     
   ' set egaza status
    If InStr(STATE, "�������") > 0 Then
        status = "�������"
    ElseIf InStr(STATE, "���") > 0 Then
        status = "��� ������"
    ElseIf InStr(STATE, "����") > 0 Then
        status = "���� �����"
    Else
        status = STATE
    End If
    
    If InStr(STATE, "����") > 0 Then
        status = status + " " + "���� �� ��� ���"
    ElseIf InStr(STATE, "����") > 0 Then
        status = status + " " + "���� �� ������"
    Else
         status = STATE
    End If
    
    
    If InStr(TAREk, "����") > 0 And index < 0 And index > -5 Then
        TAREQ = " �� ���� �������� ������"
    ElseIf InStr(TAREk, "����") > 0 And index > 80 And index < 110 Then
        TAREQ = " �� ���� �����"
    ElseIf InStr(TAREk, "����") > 0 And index > 7 And index < 11 Then
        TAREQ = " �� ���� �����"
    ElseIf InStr(TAREk, "����") > 0 Then
        TAREQ = " �� ���� ��������"
    ElseIf InStr(TAREk, "����") > 0 Then
        TAREQ = " �� ���� ������"
    Else
        TAREQ = TAREk
    End If

    
    sanada = sanadan(index)
    rawy = rawye(index)
    qeraat = qeraatn(index)
    
    qeraat = qeraat + TAREQ
    rawy = rawy + TAREQ
    set_qeraat STATE:=status, qeraat:=qeraat, rawy:=rawy
    set_snada (sanada)
        
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, NAME:="1"
    Dim tempForm As UserForm1
    For Each tempForm In UserForms
        Unload tempForm
    Next
    

   MsgBox ("���")
End Sub

Private Sub OptionButton3_Click()
 TextBox3.text = "���� ����� ������ ������ ��������"
 
 
End Sub

Private Sub OptionButton4_Click()
 TextBox3.text = "����� ������ ������ ������ ��������"
End Sub

Private Sub UserForm_Click()

End Sub

