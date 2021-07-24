VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "«Ã«“…"
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
            .Replacement.text = "«·‘ÌŒ…"
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
            .text = "›ÌﬁÊ· «·‘ÌŒ…"
            .Replacement.text = "› ﬁÊ· «·‘ÌŒ…"
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
            .Replacement.text = "«·‘ÌŒ"
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
                .text = "·ˆ„ı”˙ ÛÕÛﬁ¯ˆÂÛ« «·„ıÃÛ«“"
                .Replacement.text = "·ˆ„ı”˙ ÛÕÛﬁ ÂÛ« «·„ıÃÛ«“…"
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
                .text = "·ˆ„ı”˙ ÛÕÛﬁ¯ˆÂÛ« «·„ıÃÛ«“"
                .Replacement.text = "·ˆ„ı”˙ ÛÕÛﬁ ÂÛ« «·„ıÃÛ«“…"
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
                .text = "«”„ «·ÿ«·» Â‰«"
                .Replacement.text = "«”„ «·ÿ«·»… Â‰«"
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
                    "‰›⁄ «··Â »Â Ê⁄Û›Û« ⁄Û‰˙Âı ÊÛ⁄Û‰˙ ÊÛ«·ˆœÛÌ˙Âˆ ÊÛ‘ıÌıÊŒˆÂ ÊÛ«·˙„ı”˙·ˆ„ˆÌ‰Û"
                .Replacement.text = _
                    "‰›⁄ «··Â »Â« Ê⁄Û›Û« ⁄Û‰˙Â« ÊÛ⁄Û‰˙ ÊÛ«·ˆœÛÌ˙Â« ÊÛ‘ıÌıÊŒˆÂ« ÊÛ«·˙„ı”˙·ˆ„ˆÌ‰Û"
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
                .text = "«·⁄Û„ˆÌﬁˆ «·ÿ«·ˆ»ı «·„ıÃÛ«“ı /"
                .Replacement.text = "«·⁄Û„ˆÌﬁˆ «·ÿ«·ˆ»… «·„ıÃÛ«“… /"
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
                .text = "·ÛﬁÛœ˙ ﬁÛ—Û√Û ⁄Û·ÛÌ¯Û «·ﬁı—˙¬‰Û «·ﬂÛ—ˆÌ„Û"
                .Replacement.text = "·ÛﬁÛœ˙ ﬁÛ—Û√Û  ⁄Û·ÛÌ¯Û «·ﬁı—˙¬‰Û «·ﬂÛ—ˆÌ„Û"
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
                    "ÊÛ»Û⁄˙œ √Û‰˙ ⁄Û·ˆ„Û ı „ˆ‰˙Âı «·œ¯ˆ—Û«ÌÛ…ˆ ÊÛ«·≈ˆ ˙ﬁÛ«‰ˆ ÊÛ √ÿ˙„Û√˙‰Û‰˙ ı ≈ˆ·ÛÏ ﬁˆ—Û«¡Û ˆÂˆ ﬂı·¯Û «·≈ÿ˙„ˆ∆˙‰Û«‰ˆ"
                .Replacement.text = _
                    "ÊÛ»Û⁄˙œ √Û‰˙ ⁄Û·ˆ„Û ı „ˆ‰˙Â« «·œ¯ˆ—Û«ÌÛ…ˆ ÊÛ«·≈ˆ ˙ﬁÛ«‰ˆ ÊÛ √ÿ˙„Û√˙‰Û‰˙ ı ≈ˆ·ÛÏ ﬁˆ—Û«¡Û ˆÂ« ﬂı·¯Û «·≈ÿ˙„ˆ∆˙‰Û«‰ˆ"
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
                    "ÊÛ»Û⁄˙œ √Û‰˙ ⁄Û·ˆ„Û ı „ˆ‰˙Âı «·œ¯ˆ—Û«ÌÛ…ˆ ÊÛ«·≈ˆ ˙ﬁÛ«‰ˆ ÊÛ √ÿ˙„Û√˙‰Û‰˙ ı ≈ˆ·ÛÏ ﬁˆ—Û«¡Û ˆÂˆ ﬂı·¯Û «·≈ÿ˙„ˆ∆˙‰Û«‰ˆ"
                .Replacement.text = _
                    "ÊÛ»Û⁄˙œ √Û‰˙ ⁄Û·ˆ„Û ı „ˆ‰˙Â« «·œ¯ˆ—Û«ÌÛ…ˆ ÊÛ«·≈ˆ ˙ﬁÛ«‰ˆ ÊÛ √ÿ˙„Û√˙‰Û‰˙ ı ≈ˆ·ÛÏ ﬁˆ—Û«¡Û ˆÂ« ﬂı·¯Û «·≈ÿ˙„ˆ∆˙‰Û«‰ˆ"
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
                    "ÊÛ ﬁœ ÿÛ·Û»Û „ˆ‰ÛÏ «·≈ˆÃÛ«“Û…Û ÊÛ ﬂˆ Û«»Û…Û «·”¯Û‰Ûœˆ ›Û√ÛÃÛ“˙ ıÂı »ˆ«·ﬁˆ—Û«¡Û…ˆ"
                .Replacement.text = _
                    "ÊÛ ﬁœ ÿÛ·Û»  „ˆ‰ÛÏ «·≈ˆÃÛ«“Û…Û ÊÛ ﬂˆ Û«»Û…Û «·”¯Û‰Ûœˆ ›Û√ÛÃÛ“˙ ıÂ« »ˆ«·ﬁˆ—Û«¡Û…ˆ"
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
                    "·ˆﬂÛÊ˙‰ˆÂˆ √ÛÂ˙·« ·–Û·ˆﬂÛ ÊÛ√Û–ˆ‰˙ ı ·ÛÂı √Û‰˙ ÌÛﬁ˙—Û√Û ÊÌıﬁ˙—ˆ∆ ÊÛÌı⁄Û·¯ˆ„ı ÊÛÌıÃˆÌ“ı €ÛÌ˙—ÛÂı »ˆ„Û« ﬁÛ—Û√Û ⁄Û·ÛÌ¯Û ›ˆÌ √ÛÌ¯ˆ „ÛﬂÛ«‰"
                .Replacement.text = _
                    "·ˆﬂÛÊ˙‰ˆÂ« √ÛÂ˙·« ·–Û·ˆﬂÛ ÊÛ√Û–ˆ‰˙ ı ·ÛÂ« √Û‰˙  ﬁ˙—Û√Û Ê ﬁ˙—ˆ∆ ÊÛ  ⁄Û·¯ˆ„ı ÊÛ  ÃˆÌ“ı €ÛÌ˙—ÛÂ« »ˆ„Û« ﬁÛ—Û√Û  ⁄Û·ÛÌ¯Û ›ˆÌ √ÛÌ¯ˆ „ÛﬂÛ«‰"
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
                    " ÕÛ·¯Ú ÊÛ ›¯ÛÏ √ÛÌ¯ˆ ﬁıÿ˙— ‰Û“Û·Û »ˆ‘Û—˙ÿˆ «·˙√Û„Û«‰Û…ˆ ÊÛ «·’¯ˆÌÛ«‰Û…ˆ ÊÛ«·˙„ıÿÛ«·Û⁄Û…ˆ ÊÛ√Û·Û« ÌÛﬁıÊ·Û ≈ˆ·Û« »ˆ„Û« ÌÛ⁄˙·Û„ı ›Û≈ˆ‰˙ »Ûœ¯Û·Û √ÛÊ˙€ÛÌ¯Û—Û √ÊÛ ÷ÛÌ¯Û⁄Û «·ﬁı—˙¬‰Û"
                .Replacement.text = _
                    " ÕÛ·¯Ú  ÊÛ ›¯ÛÏ √ÛÌ¯ˆ ﬁıÿ˙— ‰Û“Û·Û  »ˆ‘Û—˙ÿˆ «·˙√Û„Û«‰Û…ˆ ÊÛ «·’¯ˆÌÛ«‰Û…ˆ ÊÛ«·˙„ıÿÛ«·Û⁄Û…ˆ ÊÛ√Û·Û«  ﬁıÊ·Û ≈ˆ·Û« »ˆ„Û«  ⁄˙·Û„ı ›Û≈ˆ‰˙ »Ûœ¯Û·Û  √ÛÊ˙ €ÛÌ¯Û—Û  √ÊÛ ÷ÛÌ¯Û⁄Û  «·ﬁı—˙¬‰Û"
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
                   "ÊÛﬁÛ⁄Û ›ˆÌ «··¯ÛÕ˙‰ˆ"
                .Replacement.text = _
                   "Êﬁ⁄  ›Ï «··Õ‰"
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
                    "ÊÛﬁÛœ˙ ÿÛ·Û»Û „ˆ‰¯ˆÏ „Û⁄˙—ˆ›Û…Û ≈ˆ”˙‰Û«œˆÏ ›ˆÌ «·ﬁı—˙¬‰ˆ «·ﬂÛ—ˆÌ„ˆ ›Û√ÛÃÛ»˙ ıÂı ÊÛ√ÛŒ˙»Û—˙ ıÂı"
                .Replacement.text = _
                    "ÊÛﬁÛœ˙ ÿÛ·Û»Û  „ˆ‰¯ˆÏ „Û⁄˙—ˆ›Û…Û ≈ˆ”˙‰Û«œˆÏ ›ˆÌ «·ﬁı—˙¬‰ˆ «·ﬂÛ—ˆÌ„ˆ ›Û√ÛÃÛ»˙ ıÂ« ÊÛ√ÛŒ˙»Û—˙ ıÂ«"
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
                .text = "«·‘ÌŒ «·„Ã«“ / "
                .Replacement.text = "«·‘ÌŒ… «·„Ã«“… / "
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
                .text = "«·‘ÌŒ «·„Ã«“ / "
                .Replacement.text = "«·‘ÌŒ… «·„Ã«“… / "
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
                .text = "ÂÛ–Û« ÊÛ√ıÊ’ˆÌ ‰Û›˙”ˆÌ ÊÛ «·„ıÃÛ«“Û "
                .Replacement.text = "ÂÛ–Û« ÊÛ√ıÊ’ˆÌ ‰Û›˙”ˆÌ ÊÛ «·„ıÃÛ«“…"
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
                .text = "ÂÛ–Û« ÊÛ√ıÊ’ˆÌ ‰Û›˙”ˆÌ ÊÛ «·„ıÃÛ«“…"
                .Replacement.text = "ÂÛ–Û« ÊÛ√ıÊ’ˆÌ ‰Û›˙”ˆÌ ÊÛ «·„ıÃÛ«“… "
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
                    "·ˆÌÛ⁄˙—ˆ›Û ﬁÛœ˙—Û „Û« ÊÛ’Û·Û ≈ˆ·ÛÌ˙Âˆ ÊÛ √ı€˙œˆﬁÛ ⁄Û·ÛÌ˙Âˆ „Û‰˙ ÂÛ–ˆÂˆ «·‰¯ˆ⁄˙„Û…ˆ «·⁄ÛŸˆÌ„Û…ˆ ÊÛ «·„ˆ‰¯Û…ˆ «·ÃÛ”ˆÌ„Û…ˆ ÊÛ ·ˆÌı⁄Û·¯ˆ„"
                .Replacement.text = _
                    "·ˆ ⁄˙—ˆ›Û ﬁÛœ˙—Û „Û« ÊÛ’Û·Û  ≈ˆ·ÛÌ˙Âˆ ÊÛ √ı€˙œˆﬁ ⁄Û·ÛÌ˙Â« „Û‰˙ ÂÛ–ˆÂˆ «·‰¯ˆ⁄˙„Û…ˆ «·⁄ÛŸˆÌ„Û…ˆ ÊÛ «·„ˆ‰¯Û…ˆ «·ÃÛ”ˆÌ„Û…ˆ ÊÛ ·ˆ ⁄Û·¯ˆ„"
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
                    "ŒÛ«›ˆ÷« ÃÛ‰Û«ÕÛÂı ·ˆﬂı·¯ˆ „Û‰˙ √ı Û«Âı ÊÛ·Û« ÌÛﬁ˙ Û’Û— ⁄Û·ÛÏ „Û« ⁄ˆ‰˙œÛÂı ÊÛÌÛ ˙—ıﬂ «·Ãˆœ¯Û"
                .Replacement.text = _
                    "ŒÛ«›ˆ÷… ÃÛ‰Û«ÕÛÂ« ·ˆﬂı·¯ˆ „Û‰˙ √ı Û«Â« ÊÛ·Û«  ﬁ˙ Û’Û— ⁄Û·ÛÏ „Û« ⁄ˆ‰˙œÛÂ« ÊÛ  ˙—ıﬂ «·Ãˆœ¯Û"
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
                    "ŒÛ«›ˆ÷« ÃÛ‰Û«ÕÛÂı ·ˆﬂı·¯ˆ „Û‰˙ √ı Û«Âı ÊÛ·Û« ÌÛﬁ˙ Û’Û— ⁄Û·ÛÏ „Û« ⁄ˆ‰˙œÛÂı ÊÛÌÛ ˙—ıﬂ «·Ãˆœ¯Û"
                .Replacement.text = _
                    "ŒÛ«›ˆ÷… ÃÛ‰Û«ÕÛÂ« ·ˆﬂı·¯ˆ „Û‰˙ √ı Û«Â« ÊÛ·Û«  ﬁ˙ Û’Û— ⁄Û·ÛÏ „Û« ⁄ˆ‰˙œÛÂ« ÊÛ  ˙—ıﬂ «·Ãˆœ¯Û"
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
                .text = "Ê·ÌÛ“ˆœÂ «·⁄ˆ·˙„Û „ÛÕÛ«”ˆ‰Û"
                .Replacement.text = "Ê·ÌÛ“ˆœÂ« «·⁄ˆ·˙„Û „ÛÕÛ«”ˆ‰Û"
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
                .text = "ÊÛ ≈ˆ‰¯ˆÏ ﬁÛœ˙ √ÛÃÛ“˙ ıﬂÛ √ÛÌÂ« «·ÿ¯Û«·ˆ»ı"
                .Replacement.text = "ÊÛ ≈ˆ‰¯ˆÏ ﬁÛœ˙ √ÛÃÛ“˙ ıﬂˆ √ÛÌ Â« «·ÿ¯Û«·ˆ»…"
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
                    "›ÛÕÛ«›ˆŸı √ˆÌÂ «·„ıÃÛ«“ı ⁄Û·ÛÏ „Û« √Ûœ¯ÛÌ˙ ıÂı ·ÛﬂÛ ÃÛ⁄Û·ÛﬂÛ"
                .Replacement.text = _
                    "›ÛÕÛ«›ˆŸˆ √ˆÌ Â« «·„ıÃÛ«“… ⁄Û·ÛÏ „Û« √Ûœ¯ÛÌ˙ ıÂı ·ÛﬂÛ ÃÛ⁄Û·Ûﬂˆ"
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
                    " ÊÛ√ıÊ’ˆÌÂˆ √Û·Û« ÌÛ‰˙”Û«‰ˆÌ ÊÛ√ÛÂ˙·ˆÌ ÊÛ–Û—¯ˆÌ¯Û ˆÌ „ˆ‰˙ ’Û«·ˆÕˆ œÛ⁄ÛÊÛ« ˆÂˆ ›ˆÌ ŒÛ·ÛÊÛ« ˆÂˆ ÊÃÛ·ÛÊÛ« ˆÂˆ ÊÛ√Û‰˙ ÌÛ–˙ﬂı—Û‰ˆÌ ⁄ˆ‰˙œÛ —Û»¯ˆÂ."
                .Replacement.text = _
                    " ÊÛ√ıÊ’ˆÌÂ« √Û·Û«  ‰˙”Û«‰ˆÌ ÊÛ√ÛÂ˙·ˆÌ ÊÛ–Û—¯ˆÌ¯Û ˆÌ „ˆ‰˙ ’Û«·ˆÕˆ œÛ⁄ÛÊÛ« ˆÂ« ›ˆÌ ŒÛ·ÛÊÛ« ˆÂ« ÊÃÛ·ÛÊÛ« ˆÂ« ÊÛ√Û‰˙  –˙ﬂı—Û‰ˆÌ ⁄ˆ‰˙œÛ —Û»¯ˆÂ«."
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
                    "Êﬁœ ﬁ—√ «·ÿ«·» √Ì÷« ⁄·Ï"
                .Replacement.text = _
                    "Êﬁœ ﬁ—√  «·ÿ«·»… √Ì÷« ⁄·Ï"
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
         '«»‰ ⁄«„—
         sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
         sanadan = sanadan & "›√„« —Ê«Ì… Â‘«„ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ﬁ«·: ÕœÀ‰« «»‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« «·Õ”Ì‰ »‰ „Â—«‰ «·Ã„«· ° ﬁ«· :ÕœÀ‰« √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì ° ﬁ«· : ÕœÀ‰« Â‘«„ »‰ ⁄„«— ° ﬁ«·: ÕœÀ‰« ⁄—«ﬂ »‰ Œ«·œ «·„—Ì ° ﬁ«· :ﬁ—√  ⁄·Ï ÌÕÌÌ »‰ «·Õ«—À «·–„«—Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï ⁄»œ «··Â »‰ ⁄«„— ° ﬁ«· : √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ ‘ÌŒ‰« ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Ê ﬁ«· : ﬁ—√  »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ⁄»œ«‰ ° Êﬁ«· : ﬁ—√  ⁄·Ï «·Õ·Ê«‰Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï Â‘«„ " & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… «»‰ –ﬂÊ«‰ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ „Ê”Ï »‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« √Õ„œ »‰ ÌÊ”› «· €·»Ì ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ –ﬂÊ«‰ ° ﬁ«· : ÕœÀ‰« √ÌÊ» »‰  „Ì„ «· „Ì„Ì ° ﬁ«· :ÕœÀ‰« ÌÕÌÌ »‰ «·Õ«—À «·–„«—Ì ° ﬁ«· : ﬁ—√  ⁄·Ï «»‰ ⁄«„— ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ⁄»œ «·⁄“Ì“ »‰ Ã⁄›— «·›«—”Ì «·„ﬁ—Ì¡ Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ï »ﬂ— „Õ„œ »‰ «·Õ”‰ «·‰ﬁ«‘ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ⁄»œ «··Â Â«—Ê‰ »‰ „Ê”Ï »‰ ‘—Ìﬂ «·√Œ›‘ Ê—Ê«Â« «·√Œ›‘ ⁄‰ ⁄»œ «··Â »‰ –ﬂÊ«‰ " & vbNewLine
         sanadan = sanadan & "Ê—Ã‹‹«· «»‰ ⁄«„— «·‹–Ì‹‰ ”‹‹„«Â„ : √»Ê «·œ—œ«¡ ⁄ÊÌ„— »‰ ⁄«„— ’«Õ» —”Ê· «··Â ° Ê«·„€Ì—… »‰ √»Ì ‘Â«» «·„Œ“Ê„Ì ° Ê√Œ‹– √»Ê «·œ—œ«¡ ⁄‹‹‰ «·‰»Ì . Ê√Œ– «·„€Ì—… ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -" & vbNewLine
         snandan = sanadan & vbNewLine
         '⁄«’„
         sanadan = sanadan & "ﬁ«· √»‹‹Ê ⁄‹„‹—Ê «·‹œ«‰‹‹Ì ›‹‹‹Ì «·‹ Ì”Ì—:" & vbNewLine
         sanadan = sanadan & "›√„« —Ê«Ì… √»Ì »ﬂ— ‘⁄»…: ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ ⁄·Ì «·ﬂ« » ﬁ«·: ÕœÀ‰« »‰ „Ã«Âœ ﬁ«·: ÕœÀ‰« ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ⁄„— «·ÊﬂÌ⁄Ì ° ﬁ«·:ÕœÀ‰« √»Ì ﬁ«·:ÕœÀ‰« ÌÕÌÌ »‰ √œ„ ° ﬁ«·: ÕœÀ‰« √»Ê »ﬂ— ⁄‰ ⁄«’„ ° ﬁ«· √»Ê ⁄„—Ê: Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ «·„ﬁ—Ì¡ ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—Ì¡ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈»—«ÂÌ„ »‰ ⁄»œ «·—Õ„‰ »‰ √Õ„œ «·„ﬁ—Ì¡ «·»€œ«œÌ Êﬁ«·: ﬁ—√  ⁄·Ï ÌÊ”› »‰ Ì⁄ﬁÊ» «·Ê«”ÿÌ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‘⁄Ì» »‰ √ÌÊ» «·’—Ì›Ì‰Ì ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ÌÕÌÌ »‰ √œ„ ⁄‰ √»Ï »ﬂ— ⁄‰ ⁄«’„." & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… Õ›’ : ›ÕœÀ‰« »Â« √»Ê «·Õ”‰ ÿ«Â‹— »‰ €·»Ê‰ «·„ﬁ—∆ ° ﬁ«· : ÕœÀ‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ’«·Õ «·Â«‘„Ì «·÷—Ì— «·„ﬁ—∆ »«·»’—… ° ﬁ«·: ÕœÀ‰« √»Ê «·⁄»«” √Õ„œ »‰ ”Â· «·√‘‰«‰Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì „Õ„œ ⁄»Ìœ »‰ «·’»«Õ ° Êﬁ«·: ﬁ—√  ⁄·Ï Õ›’ ° Êﬁ«·: ﬁ—√  ⁄·Ï ⁄«’‹„ ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ‘ÌŒ‰« √»Ì «·Õ”‰ Êﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï «·Â«‘„Ì Êﬁ«·: ﬁ—√  ⁄·Ï «·√‘‰«‰Ì ⁄‰ ⁄»Ìœ ⁄‰ Õ›’ ⁄‰ ⁄«’‹„ . " & vbNewLine
         sanadan = sanadan & "Ê—Ã«· ⁄«’„ «·‹–Ì‹‰ ”„«Â„ «À‰«‰ : √»Ê ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ê „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·»  ° Ê√»Ì »‰ ﬂ⁄»  ° Ê“Ìœ »‰ À«»   ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï - ° √Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰  ° Ê«»‰ „”⁄Êœ  ° ⁄‰ —”Ê· «··Â - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         snandan = sanadan & vbNewLine
        
         '«·ﬂ”«∆Ï
         sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
         sanadan = sanadan & "›√„« —Ê«Ì… √»Ì ⁄„—Ê «·œÊ—Ì : ›ÕœÀ‰« »Â« √»Ê „Õ„œ ⁄»œ «·—Õ„‰ »‰ ⁄„— »‰ „Õ„œ «·„⁄œ· ° ﬁ«· : ÕœÀ‰« √»Ê ⁄„— ⁄»œ «··Â »‰ √Õ„œ »‰ œÌ“ÊÌÂ «·œ„‘ﬁÌ ° ﬁ«· : ÕœÀ‰« Ã⁄›— »‰ „Õ„œ »‰ √”œ «·‰’Ì»Ì ° ﬁ«· : ÕœÀ‰« √»Ê ⁄„— «·œÊ—Ì ° ⁄‰ «·ﬂ”«∆Ì ° Ê ﬁ«· √»Ê ⁄‹„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄·Ì »‰ «·Ã·‰œÌ «·„Ê’·Ì ° Ê ﬁ«· :ﬁ—√  ⁄·Ï Ã⁄›— »‰ „Õ„œ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì ⁄„— «·œÊ—Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï «·ﬂ”«∆Ì ." & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… √»Ì «·Õ«—À : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« »Â« «»‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« „Õ„œ »‰ ÌÕÌÏ ( «·ﬂ”«∆Ì «·’€Ì—) ° ⁄‰ √»Ì «·Õ«—À ° ⁄‰ «·ﬂ”«∆Ì ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—Ì¡ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·ﬁ«”„ “Ìœ »‰ ⁄·Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï √Õ„œ »‰ «·Õ”‰ «·„⁄—Ê› »«·»ÿÌ ° Êﬁ«· :ﬁ—√  ⁄·Ï „Õ„œ »‰ ÌÕÌÏ ( «·ﬂ”«∆Ì «·’€Ì—) ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì «·Õ«—À ° Êﬁ«· : ﬁ—√  ⁄·Ï «·ﬂ”«∆Ì ." & vbNewLine
         sanadan = sanadan & "Ê—Ã«· «·ﬂ”«∆Ì : Õ„“… »‰ Õ»Ì» «·“Ì«  ° Ê⁄Ì”Ï »‰ ⁄„— «·Â„–«‰Ì ° Ê„Õ„œ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° Ê€Ì—Â„ „‰ „‘ÌŒ… «·ﬂÊ›ÌÌ‰ €Ì— √‰ „«œ… ﬁ—«¡ Â Ê«⁄ „«œÂ ›Ì «Œ Ì«—Â ⁄‰ Õ„“… ° Êﬁœ –ﬂ—‰« « ’«· ﬁ—«¡ Â ." & vbNewLine
         sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
         sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         snandan = sanadan & vbNewLine
        
         'Œ·›
         sanadan = sanadan & "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
         sanadan = sanadan & "√„« —Ê«Ì… ≈œ—Ì” «·Ê—«ﬁ : ›ÕœÀ‰« »Â« √»Ê Õ›’ ⁄„— »‰ «·Õ”‰ »ﬁ—«¡ Ì ⁄·ÌÂ Ÿ«Â— œ„‘ﬁ ° ⁄‰ ‘ÌŒÂ «·≈„«„ «·ŒÿÌ» √»Ì «·⁄»«” √Õ„œ »‰ ≈»—«ÂÌ„ »‰ ⁄„— «·›«—Ê∆Ì «·‘«›⁄Ì ° ﬁ«· : √Œ»—‰« Ê«·œÌ ° ﬁ«· : √Œ»—‰« √»Ê «·”⁄«œ«  «·√”⁄œ »‰ ”·ÿ«‰ «·Ê«”ÿÌ ° √Œ»—‰« √»Ê «·⁄“ „Õ„œ »‰ «·Õ”Ì‰ «·Ê«”ÿÌ ° √Œ»—‰« √»Ê «·Õ”Ì‰ √Õ„œ »‰ ⁄»œ «··Â »‰ «·Œ÷— «·”Ê”‰Ã—œÌ ° √Œ»—‰« √»Ê «·Õ”‰ „Õ„œ »‰ ⁄»œ «··Â »‰ „Õ„œ »‰ „—… «·ÿÊ”Ì «·„⁄—Ê› »«»‰ √»Ì ⁄„— «·‰ﬁ«‘ ° √Œ»—‰« √»Ê Ì⁄ﬁÊ» ≈”Õ«ﬁ »‰ ≈»—«ÂÌ„ «·Ê—«ﬁ ." & vbNewLine
         sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ﬂ· „‰ «·‘ÌŒÌ‰ √»Ì ⁄»œ «··Â «·Õ‰›Ì ° Ê√»Ì „Õ„œ «·‘«›⁄Ì «·„’—ÌÌ‰ ° Êﬁ—√ ﬂ· „‰Â„« ⁄·Ï √»Ì ⁄»œ «··Â „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„’—Ì ° Êﬁ—√ »Â« ⁄·Ï «·ﬂ„«· »‰ ›«—” ° Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·ﬁ«”„ Â»… «··Â »‰ √Õ„œ »‰ «·ÿ»— «·»€œ«œÌ ° Êﬁ—√ »Â« ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄·Ì »‰ „Ê”Ï «·ŒÌ«ÿ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Õ”Ì‰ «·”Ê”‰Ã—œÌ ° Êﬁ—√ »Â« ⁄·Ï «»‰ √»Ì ⁄„— «·ÿÊ”Ì ° Êﬁ—√ »Â« ⁄·Ï ≈”Õ«ﬁ «·Ê—«ﬁ ° Êﬁ—√ »Â« ⁄·Ï Œ·› ." & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… ≈œ—Ì” : ›ÕœÀ‰« »Â« √Õ„œ »‰ „Õ„œ »‰ «·Õ”Ì‰ «·›«—”Ì »ﬁ—«¡ Ì ⁄·ÌÂ ° √Œ»—‰« ⁄·Ì »‰ √Õ„œ ›Ì„« ‘«›Â‰Ì »Â °⁄‰ “Ìœ »‰ «·Õ”‰ «·»€œ«œÌ ° √Œ»—‰« √»Ê «·ﬁ«”„ »‰ √Õ„œ «·Õ—Ì—Ì ° √Œ»—‰« √»Ê »ﬂ—„Õ„œ »‰ ⁄»Ì »‰ „Õ„œ «·ŒÌ«ÿ ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ⁄»œ «··Â «·Õ–«¡ ° √Œ»—‰« √»Ê ≈”Õ«ﬁ ≈»—«ÂÌ„ »‰ «·Õ”Ì‰ »‰ ⁄»œ «··Â «·‰”«Ã «·„⁄—Ê› »«·‘ÿÌ ° √Œ»—‰« ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ «·Õœ«œ." & vbNewLine
         sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·‘ÌŒ √»Ì „Õ„œ ⁄»œ «·—Õ„‰ »‰ √Õ„œ «·Ê«”ÿÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„⁄œ· ° Êﬁ—√ »Â« ⁄·Ï ≈»—«ÂÌ„ »‰ √Õ„œ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Ì„‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì „Õ„œ ”»ÿ «·ŒÌ«ÿ ° ﬁ«· : Êﬁ—√  »Â« «·ﬁ—¬‰ „‰ √Ê·Â ≈·Ï ¬Œ—Â ⁄·Ï «·≈„«„Ì‰ «·‘—Ì› √»Ì «·›÷· ⁄»œ «·ﬁ«Â— »‰ ⁄»œ «·”·«„ «·⁄»«”Ì ° Ê√»Ì «·„⁄«·Ì À«»  »‰ »‰œ«— »‰ ≈»—«ÂÌ„ «·»ﬁ«· ° ›√„« «·‘—Ì› ›√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â „Õ„œ »‰ «·Õ”Ì‰ «·ﬂ«—“Ì‰Ì ° Ê√Œ»—Â √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ √»Ì «·⁄»«” «·Õ”‰ »‰ ”⁄Ìœ »‰ Ã⁄›— «·„ÿÊ⁄Ì ° Ê√„« √»Ê «·„⁄«·Ì ›√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ «·ﬁ«÷Ì √»Ì «·⁄·«¡ „Õ„œ »‰ ⁄·Ì »‰ Ì⁄ﬁÊ» «·Ê«”ÿÌ ° Êﬁ—√ «·Ê«”ÿÌ »Â« „‰ «·ﬂ «» ⁄·Ï «·≈„«„ √»Ì »ﬂ— √Õ„œ »‰ Ã⁄›— »‰ Õ„œ«‰ »‰ „«·ﬂ «·ﬁÿÌ⁄Ì ° Êﬁ—√ «·ﬁÿÌ⁄Ì Ê«·„ÿÊ⁄Ì Ã„Ì⁄« ⁄·Ï ≈œ—Ì” ° Êﬁ—√ ≈œ—Ì” ⁄·Ï Œ·› ° Ê«··Â «·„Ê›ﬁ . " & vbNewLine
         sanadan = sanadan & "Ê—Ã«· Œ·› : Ê—Ã«· Œ·› ”·Ì„ ’«Õ» Õ„“… ° ÊÌ⁄ﬁÊ» »‰ Œ·Ì›… «·√⁄‘Ï ’«Õ» √»Ì »ﬂ— ° Ê√»Ê “Ìœ ”⁄Ìœ ”⁄Ìœ »‰ √Ê” «·√‰’«—Ì ’«Õ» «·„›÷· «·÷»Ì Ê√»«‰ «·⁄ÿ«— ° Êﬁ—√ √»Ê »ﬂ— ° Ê«·„›÷· ° Ê√»«‰ ⁄·Ï ⁄«’„ . Ê—ÊÏ «·ﬁ—«¡… √Ì÷« ⁄‰ «·ﬂ”«∆Ì Ê⁄‰ ÌÕÌÏ »‰ ¬œ„ ⁄‰ √»Ì »ﬂ— ° Ê«··Â «·„Ê›ﬁ . ﬁ·  : Ê√Œ– ⁄«’„ ⁄‰ √»Ì ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ì „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·» ° Ê√»Ì »‰ ﬂ⁄» ° Ê“Ìœ »‰ À«»  ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         sanadan = sanadan & "Ê√Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰ ° Ê«»‰ „”⁄Êœ ° ⁄‰ —”Ê· «··Â ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -. Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ . Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         snandan = sanadan & vbNewLine
         
        
        ElseIf index = -2 Then
        
        ' √»Ê ⁄„—Ê
        sanadan = "”‰œ ﬁ—«¡… «·≈„«„ / √»Ê ⁄„—Ê «·»’—Ï" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… √»Ì ⁄„— «·œÊ—Ì : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ ⁄·Ì ° ﬁ«·: √Œ»—‰« √»Ê ⁄Ì”Ï „Õ„œ »‰ √Õ„œ »‰ ﬁÿ‰ ”‰… À„«‰ ⁄‘—… ÊÀ·«À„«∆…° ﬁ«·: √Œ»—‰« √»Ê Œ·«œ ”·Ì„«‰ »‰ Œ·«œ ﬁ«·:ÕœÀ‰« «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê ° ﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â „‰ ÿ—Ìﬁ √»Ì ⁄„— «·œÊ—Ì ⁄·Ï ‘ÌŒ‰« ⁄»œ «·⁄“ Ì“ »‰ Ã⁄›— »‰ „Õ„œ »‰ ≈”Õ«ﬁ «·»€œ«œÌ «·›«—”Ì «·„ﬁ—Ì¡ ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì ÿ«Â— ⁄»œ «·Ê«Õœ »‰ ⁄„— »‰ √»Ì Â«‘„ «·„ﬁ—Ì¡ ° „« ·« √Õ’ÌÂ ﬂÀ—… ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì »ﬂ— »‰ „Ã«Âœ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·“⁄—«¡ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” Êﬁ«· :ﬁ—√  ⁄·Ï √»Ì ⁄„— ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï «·Ì“ÌœÌ ° Êﬁ«· ﬁ—√  »Â« ⁄·Ï : √»Ì ⁄„—Ê. " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… √»Ì ‘⁄Ì» «·”Ê”Ì : ›ÕœÀ‰« »Â« Œ·› »‰ ≈»—«ÂÌ„ »‰ „Õ„œ «·„ﬁ—∆ ° ﬁ«· : ÕœÀ‰« √»Ê „Õ„œ «·Õ”‰ »‰ —‘Ìﬁ «·„⁄œ· ° ﬁ«· : ÕœÀ‰« √»Ê ⁄»œ «·—Õ„‰ √Õ„œ »‰ ‘⁄Ì» «·‰”«∆Ì ° ﬁ«· : √Œ»—‰« √»Ê ‘⁄Ì» ° ﬁ«· : √Œ»—‰« «·Ì“ÌœÌ ° ⁄‰ √»Ì ⁄„—Ê ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â »≈ŸÂ«— «·√Ê· „‰ «·„À·Ì‰ Ê«·„ ﬁ«—»Ì‰ Ê»≈œ€«„Â ⁄·Ï ›«—” »‰ √Õ„œ «·„ﬁ—Ì¡ ° Êﬁ«· ·Ì : ﬁ—√  »Â« ﬂ–·ﬂ ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Êﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ﬂ·Â ﬂ–·ﬂ ⁄·Ï √»Ì ⁄„—«‰ „Ê”Ï »‰ Ã—Ì— «·‰ÕÊÌ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ‘⁄Ì» ° Êﬁ«· : ﬁ—√  ⁄·Ï «·Ì“ÌœÌ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì ⁄„—Ê" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê: ÊÕœÀ‰« »√’Ê· «·≈œ€«„ „Õ„œ »‰ √Õ„œ ⁄‰ «»‰ „Ã«Âœ ⁄‰ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” ⁄‰ «·œÊ—Ì ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ï ⁄„—Ê° ÊÕœÀ‰« »Â« √Ì÷« √»Ê «·Õ”‰ ‘ÌŒ‰« ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ «·„»«—ﬂ ⁄‰ Ã⁄›— »‰ ”·Ì„«‰ ⁄‰ √»Ì ‘⁄Ì» ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê . " & vbNewLine
        sanadan = sanadan & "Ê—Ã«· √»Ì ⁄„—Ê : Ã„«⁄… „‰ √Â· «·ÕÃ«“ Ê„‰ √Â· «·»’—… ° ›„‰ √Â· „ﬂ… : „Ã«Âœ ° Ê”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„… »‰ Œ«·œ ° Ê⁄ÿ«¡ »‰ √»Ì —»«Õ ° Ê⁄»œ «··Â »‰ ﬂÀÌ— ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ „ÕÌ’‰ ° ÊÕ„Ìœ »‰ ﬁÌ” «·√⁄—Ã «·ﬁ«—∆ ° Ê„‰ √Â· «·„œÌ‰… : Ì“Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—Ì¡ ÊÌ“Ìœ »‰ —Ê„«‰ ° Ê‘Ì»… »‰ ‰’«Õ ° Ê„‰ √Â· «·»’—… : «·Õ”‰ »‰ √»Ì «·Õ”‰ «·»’—Ì ° ÊÌÕÌ »‰ Ì⁄„— ° Ê€Ì—Â„« ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄„‰  ﬁœ„ „‰ «·’Õ«»… Ê€Ì—Â„ . " & vbNewLine
        sanadan = sanadan & "ﬁ·  : Ê√Œ– ”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„…° ÊÌÕÌÏ »‰ Ì⁄„— ° ⁄‰ «»‰ ⁄»«” Ê√Œ– «»‰ ⁄»«” ⁄‰ √»Ì »‰ ﬂ⁄» Ê“Ìœ »‰ À«»  ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         
        'Ì⁄ﬁÊ»
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / Ì⁄ﬁÊ» «·»’—Ï" & vbNewLine
        sanadan = sanadan & "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… —ÊÌ” : ›ÕœÀ‰« »Â« «·‘ÌŒ «·≈„«„ √»Ê «·⁄»«” √Õ„œ »‰ „Õ„œ »‰ «·Œ÷— «·Õ‰›Ì »ﬁ—«¡ Ì ⁄·ÌÂ ﬁ«·: √Œ»—‰« : √»Ê «·⁄»«” √Õ„œ »‰ √»Ì ÿ«·» »‰ √»Ì «·‰⁄„ «·’«·ÕÌ ﬁ—«¡… ⁄·ÌÂ ° √Œ»—‰« √»Ê ÿ«·» ⁄»œ «··ÿÌ› »‰ „Õ„œ »‰ «·ﬁ»ÌÿÌ ° ›Ì ﬂ «»Â √Œ»—‰« »Â« √»Ê »ﬂ— √Õ„œ »‰ «·„ﬁ—» «·ﬂ—ŒÌ ﬁ—«¡… ⁄·ÌÂ ° √Œ»—‰« √»Ê ÿ«Â— √Õ„œ »‰ ⁄·Ì «·„ﬁ—Ì¡ «·√” «– √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ⁄·Ì «·ŒÌ«ÿ ° √Œ»—‰« «·√” «– «·≈„«„ √»Ê «·Õ”‰ ⁄·Ì »‰ √Õ„œ »‰ ⁄„— «·Õ„«„Ì ° √Œ»—‰« √»Ê «·ﬁ«”„ ⁄»œ «··Â »‰ «·Õ”‰ »‰ ”·Ì„«‰ «·‰Œ«” ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ Â«—Ê‰ »‰ ‰«›⁄ «· „«— «·»€œ«œÌ ° √Œ»—‰« √»Ê ⁄»œ «··Â „Õ„œ »‰ «·„ Êﬂ· «·„⁄—Ê› »—ÊÌ” ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì „Õ„œ ⁄»œ «·—Õ„‰ »‰ √Õ„œ »‰ ⁄·Ì «·»€œ«œÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï «·≈„«„ «· ﬁÌ „Õ„œ »‰ √Õ„œ «·„’—Ì ° Êﬁ—√ »Â« ⁄·Ï ≈»—«ÂÌ„ »‰ √Õ„œ «·≈”ﬂ‰œ—Ì ° Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï ⁄»œ «··Â »‰ ⁄·Ì «·»€œ«œÌ ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì «·⁄“ «·ﬁ·«‰”Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄·Ì «·Õ”‰ »‰ «·ﬁ«”„ «·Ê«”ÿÌ ° Êﬁ—√ »Â« ⁄·Ï : «·Õ„«„Ì ° Êﬁ—√ »Â« ⁄·Ï «·‰Œ« ” ° Êﬁ—√ »Â« ⁄·Ï «· „«— ° Êﬁ—√ ⁄·Ï —ÊÌ” ° Êﬁ—√ »Â« ⁄·Ï Ì⁄ﬁÊ» . " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… —ÛÊÕ : ›ÕœÀ‰« »Â« «·‘ÌŒ √»Ê «·⁄»«” √Õ„œ »‰ „Õ„œ »‰ «·Õ”Ì‰ «·‘Ì—«“Ì »ﬁ—«¡ Ì ⁄·ÌÂ ⁄‰ «·≈„«„ √»Ì «·Õ”‰ ⁄·Ì »‰ √Õ„œ «·„ﬁœ”Ì ° √Œ»—‰« √»Ê «·Ì„‰ «·ﬂ‰œÌ ‘›«Â« ° √Œ»—‰« √»Ê „Õ„œ «·»€œ«œÌ ° √Œ»—‰« √»Ê «·›÷· «·‘—Ì› «·„ﬂÌ ° √Œ»—‰« „Õ„œ »‰ «·Õ”Ì‰ «·›«—”Ì ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ≈»—«ÂÌ„ »‰ Œ‘‰«„ «·„«·ﬂÌ «·»’—Ì √Œ»—‰« √»Ê «·⁄»«” „Õ„œ »‰ Ì⁄ﬁÊ» »‰ «·ÕÃ«Ã »‰ „⁄«ÊÌ… «· Ì„Ì ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ ÊÂ» »‰ ÌÕÌÏ »‰ «·⁄·«¡ «·Àﬁ›Ì «·ﬁ“«“ ° √Œ»—‰« —ÊÕ »‰ ⁄»œ «·„ƒ„‰ «·»’—Ì ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï „Õ„œ »‰ √Õ„œ »«·ﬁ«Â—… «·„Õ—Ê”… ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â «·’«∆€ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ≈”Õ«ﬁ «·œ„‘ﬁÌ Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï ⁄»œ «··Â »‰ ⁄·Ì ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì ÿ«Â— »‰ ”Ê«— ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·ﬁ«”„ «·„”«›— »‰ «·ÿÌ» »‰ ⁄»«œ «·»’—Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ Œ‘‰«„ ° Êﬁ—√ »Â« ⁄·Ï «»‰ ⁄»« ” «· Ì„Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ ÊÂ» ° Êﬁ—√ »Â« ⁄·Ï —ÊÕ ° Êﬁ—√ »Â« ⁄·Ï Ì⁄ﬁÊ» ." & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· Ì⁄ﬁÊ» «·–Ì‰ ”„«Â„ √—»⁄… : √»Ê «·„‰–— ”·«„ »‰ ”·Ì„«‰ «·ÿÊÌ· ° Ê‘Â«» »‰ ‘—‰›… ° Ê„ÂœÌ »‰ „Ì„Ê‰ ° Ê√»Ê «·√‘Â» Ã⁄›— »‰ ÕÌ«‰ «·⁄ÿ«—œÌ .ÊﬁÌ· ≈‰ Ì⁄ﬁÊ» ﬁ—√ ⁄·Ï √»Ì ⁄„—Ê »‰ «·⁄·«¡ Êﬁ—√ ”·«„ ⁄·Ï ⁄«’„ Ê√»Ì ⁄„—Ê ° Êﬁ‹‹‹—√ ‘Â«» «·ÃÕœ—Ì Êﬁ—√ ⁄«’„ ⁄·Ï «·Õ”‰ «·»’—Ì Ê⁄·Ï ”·Ì„«‰ »‰ ﬁ … Êﬁ—√ ”·Ì„«‰ ⁄·Ï «»‹‰ ⁄»« ” Êﬁ—√ „ÂœÌ ⁄·Ï ‘⁄Ì» »‰ «·Õ»Õ«» Êﬁ—√ ⁄·Ï √»Ì «·⁄«·Ì… «·—Ì«ÕÌ Êﬁ—√ ⁄·Ï √»Ì Ê“Ìœ Êﬁ—√ √»Ê «·√‘Â» ⁄·Ï √»Ì —Ã«¡ ⁄„—«‰ »‰ „·Õ«‰ «·⁄ÿ«—œÌ Êﬁ—√ ⁄·Ï √»Ì „Ê”‹‹‹Ï «·√‘⁄—Ì Êﬁ—√ ⁄·Ï —”Ê· «··Â ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
  
         ElseIf index = -3 Then
        
         ' ‰«›⁄
        sanadan = "”‰œ ﬁ—«¡… «·≈„«„ / ‰«›⁄" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "√„« —Ê«Ì… ﬁ«·Ê‰ : ›ÕœÀ‰« »Â« √Õ„œ »‰ ⁄„— »‰ „Õ„œ «·ÃÌ“Ì ° ﬁ«·: ÕœÀ‰« „Õ„œ »‰ √Õ„œ »‰ „‰Ì— ° ﬁ«·: ÕœÀ‰« ⁄»œ «··Â »‰ ⁄Ì”Ï «·„œ‰Ì ° ﬁ«·:ÕœÀ‰« ﬁ«·Ê‰ ⁄‰ ‰«›⁄° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ‘ÌŒÌ √»Ì «·› Õ ›«—” »‰ √Õ„œ »‰ „Ê”Ï »‰ ⁄„—«‰ ° «·„ﬁ—Ì¡ «·÷—Ì— ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—∆ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈»—«ÂÌ„ »‰ ⁄„— «·„ﬁ—∆ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”Ì‰ √Õ„œ »‰ ⁄À„«‰ »‰ Ã⁄›— »‰ »ÊÌ«‰ ° Êﬁ«·:ﬁ—√  ⁄·Ï √»Ì »ﬂ— √Õ„œ »‰ „Õ„œ »‰ «·√‘⁄À Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì ‰‘Ìÿ „Õ„œ »‰ Â«—Ê‰ ° Êﬁ«·: ﬁ—√  ⁄·Ï ﬁ«·Ê‰ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‰«›⁄ ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… Ê—‘ : ›ÕœÀ‰« »Â« √»Ê ⁄»œ «··Â √Õ„œ »‰ „Õ›ÊŸ «·ﬁ«÷Ì »„’— ° ﬁ«·: ÕœÀ‰« √Õ„œ »‰ ≈»—«ÂÌ„ »‰ Ã«„⁄ ° ﬁ«· : ÕœÀ‰« √»Ê „Õ„œ »ﬂ— »‰ ”Â· ° ﬁ«·: ÕœÀ‰« √»Ê „Õ„œ ⁄»œ «·’„œ »‰ ⁄»œ «·—Õ„‰ ° ﬁ«· : ÕœÀ‰« Ê—‘ ⁄‰ ‰«›⁄ ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ‘ÌŒÌ √»Ì «·ﬁ«”„ Œ·› »‰ ≈»—«ÂÌ„ »‰ „Õ„œ »‰ Œ«ﬁ«‰ «·„ﬁ—Ì¡ »„’— ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ⁄·Ï √»Ì Ã⁄›— √Õ„œ »‰ √”«„… «· ÃÌ»Ì ° Êﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·‰Õ«” ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì Ì⁄ﬁÊ» ÌÊ”› »‰ ⁄„—Ê »‰ Ì”«— «·√“—ﬁ ° Êﬁ«· :ﬁ—√  ⁄·Ï Ê—‘ Êﬁ«· : ﬁ—√  ⁄·Ï ‰«›⁄ ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· ‰«›⁄ «·–Ì‰ ”„«Â„ Œ„”… : √»Ê Ã⁄›— Ì“ Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—∆ ° Ê√»Ê œ«Êœ ⁄»œ «·—Õ„‰ »‰ Â—„“ «·√⁄—Ã ° Ê‘Ì»… »‰ ‰’«Õ «·ﬁ«÷Ì ° Ê√»Ê ⁄»œ «··Â „”·„ »‰ Ã‰œ» «·Â–·Ì «·ﬁ«’ ° Ê√»Ê —ÊÕ Ì“Ìœ »‰ —Ê„«‰ ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄‰ √»Ì Â—Ì—… ° Ê«»‰ ⁄»«” ° Ê⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° ⁄‰ √»Ì »‰ ﬂ⁄» ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
      
         ' «»‰ ﬂÀÌ—
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / «»‰ ﬂÀÌ—" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—  " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… «·»“Ì : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ «·ﬂ« » ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ „Ê”Ï ° ﬁ«·: ÕœÀ‰« „÷— »‰ „Õ„œ «·÷»Ì ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ √»Ì »“… ° ﬁ«·: ﬁ—√  ⁄·Ï ⁄ﬂ—„… »‰ ”·Ì„«‰ »‰ ⁄«„— ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·ﬁ”ÿ ° Êﬁ«· : ﬁ—√  ⁄·Ï «»‰ ﬂÀÌ— ‰›”Â ° ﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·ﬁ«”„ ⁄»œ «·⁄“Ì“ »‰ Ã⁄›— »‰ „Õ„œ «·„ﬁ—Ì¡ «·›«—”Ì ° Êﬁ«· ·Ì: ﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ «·Õ”‰ «·‰ﬁ«‘ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì —»Ì⁄… „Õ„œ »‰ ≈”Õ«ﬁ «·— »⁄Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï «·»“Ì ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… ﬁ‰»· : ›ÕœÀ‰« »Â« √»Ê „”·„ „Õ„œ »‰ √Õ„œ «·»€œ«œÌ ° ﬁ«·: ÕœÀ‰« «»‰ „Ã«Âœ ° ﬁ«·: ﬁ—√  ⁄·Ï ﬁ‰»· ° Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì «·Õ”‰ √Õ„œ »‰ ⁄Ê‰ «·ﬁÊ«” Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì «·«Œ— Ìÿ ÊÂ» »‰ Ê«÷Õ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·ﬁ”ÿ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‘»· »‰ ⁄»«œ Ê „⁄—Ê› »‰ „‘ﬂ«‰ ° Êﬁ«·« ﬁ—√‰« ⁄·Ï «»‰ ﬂÀ‹Ì‹— ° Ê ﬁ«· √»‹‹‹‹Ê ⁄‹‹„‹‹—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ «·Õ„’Ì «·„ﬁ—Ì¡ «·÷—Ì— Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·»€œ«œÌ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï «»‰ „Ã«Âœ Êﬁ«·: ﬁ—√  ⁄·Ï ﬁ‰»· ." & vbNewLine
        sanadan = sanadan & " Ê—Ã‹‹«· «»‰ ﬂÀÌ— «·‹–Ì‹‰ ”„«Â„ À·«À… : ⁄»œ «··Â »‰ «·”«∆» «·„Œ“Ê„Ì ’«Õ» —”Ê· «··Â  Ê„Ã«Âœ »‰ Ã»— √»Ê «·ÕÃ«Ã „Ê·Ï ﬁÌ” »‰ «·”«∆» ° Êœ—»«” „Ê·Ï «»‰ ⁄»«” . Ê√Œ– ⁄»œ «··Â ⁄‰ √»Ì »‰ ﬂ⁄» ‰›”Â. Ê√Œ– „Ã«Âœ Êœ—»«”° ⁄‰ «»‰ ⁄»«”° ⁄‰ √»Ì ° Ê“Ìœ »‰ À«»  ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  °⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
      
       ' √»Ê ⁄„—Ê
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / √»Ê ⁄„—Ê «·»’—Ï" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… √»Ì ⁄„— «·œÊ—Ì : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ ⁄·Ì ° ﬁ«·: √Œ»—‰« √»Ê ⁄Ì”Ï „Õ„œ »‰ √Õ„œ »‰ ﬁÿ‰ ”‰… À„«‰ ⁄‘—… ÊÀ·«À„«∆…° ﬁ«·: √Œ»—‰« √»Ê Œ·«œ ”·Ì„«‰ »‰ Œ·«œ ﬁ«·:ÕœÀ‰« «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê ° ﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â „‰ ÿ—Ìﬁ √»Ì ⁄„— «·œÊ—Ì ⁄·Ï ‘ÌŒ‰« ⁄»œ «·⁄“ Ì“ »‰ Ã⁄›— »‰ „Õ„œ »‰ ≈”Õ«ﬁ «·»€œ«œÌ «·›«—”Ì «·„ﬁ—Ì¡ ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì ÿ«Â— ⁄»œ «·Ê«Õœ »‰ ⁄„— »‰ √»Ì Â«‘„ «·„ﬁ—Ì¡ ° „« ·« √Õ’ÌÂ ﬂÀ—… ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì »ﬂ— »‰ „Ã«Âœ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·“⁄—«¡ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” Êﬁ«· :ﬁ—√  ⁄·Ï √»Ì ⁄„— ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï «·Ì“ÌœÌ ° Êﬁ«· ﬁ—√  »Â« ⁄·Ï : √»Ì ⁄„—Ê. " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… √»Ì ‘⁄Ì» «·”Ê”Ì : ›ÕœÀ‰« »Â« Œ·› »‰ ≈»—«ÂÌ„ »‰ „Õ„œ «·„ﬁ—∆ ° ﬁ«· : ÕœÀ‰« √»Ê „Õ„œ «·Õ”‰ »‰ —‘Ìﬁ «·„⁄œ· ° ﬁ«· : ÕœÀ‰« √»Ê ⁄»œ «·—Õ„‰ √Õ„œ »‰ ‘⁄Ì» «·‰”«∆Ì ° ﬁ«· : √Œ»—‰« √»Ê ‘⁄Ì» ° ﬁ«· : √Œ»—‰« «·Ì“ÌœÌ ° ⁄‰ √»Ì ⁄„—Ê ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â »≈ŸÂ«— «·√Ê· „‰ «·„À·Ì‰ Ê«·„ ﬁ«—»Ì‰ Ê»≈œ€«„Â ⁄·Ï ›«—” »‰ √Õ„œ «·„ﬁ—Ì¡ ° Êﬁ«· ·Ì : ﬁ—√  »Â« ﬂ–·ﬂ ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Êﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ﬂ·Â ﬂ–·ﬂ ⁄·Ï √»Ì ⁄„—«‰ „Ê”Ï »‰ Ã—Ì— «·‰ÕÊÌ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ‘⁄Ì» ° Êﬁ«· : ﬁ—√  ⁄·Ï «·Ì“ÌœÌ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì ⁄„—Ê" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê: ÊÕœÀ‰« »√’Ê· «·≈œ€«„ „Õ„œ »‰ √Õ„œ ⁄‰ «»‰ „Ã«Âœ ⁄‰ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” ⁄‰ «·œÊ—Ì ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ï ⁄„—Ê° ÊÕœÀ‰« »Â« √Ì÷« √»Ê «·Õ”‰ ‘ÌŒ‰« ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ «·„»«—ﬂ ⁄‰ Ã⁄›— »‰ ”·Ì„«‰ ⁄‰ √»Ì ‘⁄Ì» ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê . " & vbNewLine
        sanadan = sanadan & "Ê—Ã«· √»Ì ⁄„—Ê : Ã„«⁄… „‰ √Â· «·ÕÃ«“ Ê„‰ √Â· «·»’—… ° ›„‰ √Â· „ﬂ… : „Ã«Âœ ° Ê”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„… »‰ Œ«·œ ° Ê⁄ÿ«¡ »‰ √»Ì —»«Õ ° Ê⁄»œ «··Â »‰ ﬂÀÌ— ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ „ÕÌ’‰ ° ÊÕ„Ìœ »‰ ﬁÌ” «·√⁄—Ã «·ﬁ«—∆ ° Ê„‰ √Â· «·„œÌ‰… : Ì“Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—Ì¡ ÊÌ“Ìœ »‰ —Ê„«‰ ° Ê‘Ì»… »‰ ‰’«Õ ° Ê„‰ √Â· «·»’—… : «·Õ”‰ »‰ √»Ì «·Õ”‰ «·»’—Ì ° ÊÌÕÌ »‰ Ì⁄„— ° Ê€Ì—Â„« ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄„‰  ﬁœ„ „‰ «·’Õ«»… Ê€Ì—Â„ . " & vbNewLine
        sanadan = sanadan & "ﬁ·  : Ê√Œ– ”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„…° ÊÌÕÌÏ »‰ Ì⁄„— ° ⁄‰ «»‰ ⁄»«” Ê√Œ– «»‰ ⁄»«” ⁄‰ √»Ì »‰ ﬂ⁄» Ê“Ìœ »‰ À«»  ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        '«»‰ ⁄«„—
         sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / «»‰ ⁄«„—" & vbNewLine
         sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
         sanadan = sanadan & "›√„« —Ê«Ì… Â‘«„ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ﬁ«·: ÕœÀ‰« «»‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« «·Õ”Ì‰ »‰ „Â—«‰ «·Ã„«· ° ﬁ«· :ÕœÀ‰« √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì ° ﬁ«· : ÕœÀ‰« Â‘«„ »‰ ⁄„«— ° ﬁ«·: ÕœÀ‰« ⁄—«ﬂ »‰ Œ«·œ «·„—Ì ° ﬁ«· :ﬁ—√  ⁄·Ï ÌÕÌÌ »‰ «·Õ«—À «·–„«—Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï ⁄»œ «··Â »‰ ⁄«„— ° ﬁ«· : √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ ‘ÌŒ‰« ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Ê ﬁ«· : ﬁ—√  »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ⁄»œ«‰ ° Êﬁ«· : ﬁ—√  ⁄·Ï «·Õ·Ê«‰Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï Â‘«„ " & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… «»‰ –ﬂÊ«‰ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ „Ê”Ï »‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« √Õ„œ »‰ ÌÊ”› «· €·»Ì ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ –ﬂÊ«‰ ° ﬁ«· : ÕœÀ‰« √ÌÊ» »‰  „Ì„ «· „Ì„Ì ° ﬁ«· :ÕœÀ‰« ÌÕÌÌ »‰ «·Õ«—À «·–„«—Ì ° ﬁ«· : ﬁ—√  ⁄·Ï «»‰ ⁄«„— ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ⁄»œ «·⁄“Ì“ »‰ Ã⁄›— «·›«—”Ì «·„ﬁ—Ì¡ Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ï »ﬂ— „Õ„œ »‰ «·Õ”‰ «·‰ﬁ«‘ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ⁄»œ «··Â Â«—Ê‰ »‰ „Ê”Ï »‰ ‘—Ìﬂ «·√Œ›‘ Ê—Ê«Â« «·√Œ›‘ ⁄‰ ⁄»œ «··Â »‰ –ﬂÊ«‰ " & vbNewLine
         sanadan = sanadan & "Ê—Ã‹‹«· «»‰ ⁄«„— «·‹–Ì‹‰ ”‹‹„«Â„ : √»Ê «·œ—œ«¡ ⁄ÊÌ„— »‰ ⁄«„— ’«Õ» —”Ê· «··Â ° Ê«·„€Ì—… »‰ √»Ì ‘Â«» «·„Œ“Ê„Ì ° Ê√Œ‹– √»Ê «·œ—œ«¡ ⁄‹‹‰ «·‰»Ì . Ê√Œ– «·„€Ì—… ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -" & vbNewLine
         snandan = sanadan & vbNewLine
        
        
         '⁄«’„
         sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / ⁄«’„" & vbNewLine
         sanadan = sanadan & "ﬁ«· √»‹‹Ê ⁄‹„‹—Ê «·‹œ«‰‹‹Ì ›‹‹‹Ì «·‹ Ì”Ì—:" & vbNewLine
         sanadan = sanadan & "›√„« —Ê«Ì… √»Ì »ﬂ— ‘⁄»…: ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ ⁄·Ì «·ﬂ« » ﬁ«·: ÕœÀ‰« »‰ „Ã«Âœ ﬁ«·: ÕœÀ‰« ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ⁄„— «·ÊﬂÌ⁄Ì ° ﬁ«·:ÕœÀ‰« √»Ì ﬁ«·:ÕœÀ‰« ÌÕÌÌ »‰ √œ„ ° ﬁ«·: ÕœÀ‰« √»Ê »ﬂ— ⁄‰ ⁄«’„ ° ﬁ«· √»Ê ⁄„—Ê: Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ «·„ﬁ—Ì¡ ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—Ì¡ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈»—«ÂÌ„ »‰ ⁄»œ «·—Õ„‰ »‰ √Õ„œ «·„ﬁ—Ì¡ «·»€œ«œÌ Êﬁ«·: ﬁ—√  ⁄·Ï ÌÊ”› »‰ Ì⁄ﬁÊ» «·Ê«”ÿÌ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‘⁄Ì» »‰ √ÌÊ» «·’—Ì›Ì‰Ì ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ÌÕÌÌ »‰ √œ„ ⁄‰ √»Ï »ﬂ— ⁄‰ ⁄«’„." & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… Õ›’ : ›ÕœÀ‰« »Â« √»Ê «·Õ”‰ ÿ«Â‹— »‰ €·»Ê‰ «·„ﬁ—∆ ° ﬁ«· : ÕœÀ‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ’«·Õ «·Â«‘„Ì «·÷—Ì— «·„ﬁ—∆ »«·»’—… ° ﬁ«·: ÕœÀ‰« √»Ê «·⁄»«” √Õ„œ »‰ ”Â· «·√‘‰«‰Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì „Õ„œ ⁄»Ìœ »‰ «·’»«Õ ° Êﬁ«·: ﬁ—√  ⁄·Ï Õ›’ ° Êﬁ«·: ﬁ—√  ⁄·Ï ⁄«’‹„ ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ‘ÌŒ‰« √»Ì «·Õ”‰ Êﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï «·Â«‘„Ì Êﬁ«·: ﬁ—√  ⁄·Ï «·√‘‰«‰Ì ⁄‰ ⁄»Ìœ ⁄‰ Õ›’ ⁄‰ ⁄«’‹„ . " & vbNewLine
         sanadan = sanadan & "Ê—Ã«· ⁄«’„ «·‹–Ì‹‰ ”„«Â„ «À‰«‰ : √»Ê ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ê „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·»  ° Ê√»Ì »‰ ﬂ⁄»  ° Ê“Ìœ »‰ À«»   ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï - ° √Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰  ° Ê«»‰ „”⁄Êœ  ° ⁄‰ —”Ê· «··Â - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         snandan = sanadan & vbNewLine
         
        'Õ„“…
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / Õ„“…" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… Œ·› : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« «»‰ „Ã«Âœ ° ÕœÀ‰« ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ ° ﬁ«· : ÕœÀ‰« Œ·› ° ﬁ«·: ⁄‰ ”·Ì„ ⁄‰ Õ„“… ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·Õ”‰ ‘ÌŒ‰« ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ „Õ„œ »‰ ÌÊ”› »‰ ‰Â«— «·Õ— ﬂÌ »«·»’—… ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”Ì‰ √Õ„œ »‰ ⁄À„«‰ »‰ Ã⁄›— »‰ »ÊÌ«‰ ° Êﬁ«· ·Ì :ﬁ—√  ⁄·Ï ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ ﬁ»· √‰ Ìﬁ—Ì¡ »«Œ Ì«— Œ·› ° Êﬁ«· ·Ì : ﬁ—√  ⁄·Ï Œ·› ° Êﬁ«· : ﬁ—√  ⁄·Ï ”·Ì„ ° Ê ﬁ«· : ﬁ—√  ⁄·Ï Õ„“… ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… Œ·«œ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« √Õ„œ »‰ „Ê”Ï ° ﬁ«· : ÕœÀ‰« ÌÕÌÏ »‰ √Õ„œ »‰ Â«—Ê‰ «·„“Êﬁ ° ⁄‰ √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì ° ⁄‰ Œ·«œ ° ⁄‰ ”·Ì„ ° ⁄‰ Õ„“… ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ «·÷—Ì— ‘ÌŒ‰« ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ‘‰»Ê– ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ‘«–«‰ «·ÃÊÂ—Ì «·„ﬁ—Ì ° Êﬁ«· :ﬁ—√  ⁄·Ï Œ·«œ Êﬁ«· : ﬁ—√  ⁄·Ï ”·Ì„ ° Êﬁ—√ ”·Ì„ ⁄·Ï Õ„“…." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
        sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
       
        '«·ﬂ”«∆Ï
         sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / «·ﬂ”«∆Ï" & vbNewLine
         sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
         sanadan = sanadan & "›√„« —Ê«Ì… √»Ì ⁄„—Ê «·œÊ—Ì : ›ÕœÀ‰« »Â« √»Ê „Õ„œ ⁄»œ «·—Õ„‰ »‰ ⁄„— »‰ „Õ„œ «·„⁄œ· ° ﬁ«· : ÕœÀ‰« √»Ê ⁄„— ⁄»œ «··Â »‰ √Õ„œ »‰ œÌ“ÊÌÂ «·œ„‘ﬁÌ ° ﬁ«· : ÕœÀ‰« Ã⁄›— »‰ „Õ„œ »‰ √”œ «·‰’Ì»Ì ° ﬁ«· : ÕœÀ‰« √»Ê ⁄„— «·œÊ—Ì ° ⁄‰ «·ﬂ”«∆Ì ° Ê ﬁ«· √»Ê ⁄‹„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄·Ì »‰ «·Ã·‰œÌ «·„Ê’·Ì ° Ê ﬁ«· :ﬁ—√  ⁄·Ï Ã⁄›— »‰ „Õ„œ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì ⁄„— «·œÊ—Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï «·ﬂ”«∆Ì ." & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… √»Ì «·Õ«—À : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« »Â« «»‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« „Õ„œ »‰ ÌÕÌÏ ( «·ﬂ”«∆Ì «·’€Ì—) ° ⁄‰ √»Ì «·Õ«—À ° ⁄‰ «·ﬂ”«∆Ì ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—Ì¡ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·ﬁ«”„ “Ìœ »‰ ⁄·Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï √Õ„œ »‰ «·Õ”‰ «·„⁄—Ê› »«·»ÿÌ ° Êﬁ«· :ﬁ—√  ⁄·Ï „Õ„œ »‰ ÌÕÌÏ ( «·ﬂ”«∆Ì «·’€Ì—) ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì «·Õ«—À ° Êﬁ«· : ﬁ—√  ⁄·Ï «·ﬂ”«∆Ì ." & vbNewLine
         sanadan = sanadan & "Ê—Ã«· «·ﬂ”«∆Ì : Õ„“… »‰ Õ»Ì» «·“Ì«  ° Ê⁄Ì”Ï »‰ ⁄„— «·Â„–«‰Ì ° Ê„Õ„œ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° Ê€Ì—Â„ „‰ „‘ÌŒ… «·ﬂÊ›ÌÌ‰ €Ì— √‰ „«œ… ﬁ—«¡ Â Ê«⁄ „«œÂ ›Ì «Œ Ì«—Â ⁄‰ Õ„“… ° Êﬁœ –ﬂ—‰« « ’«· ﬁ—«¡ Â ." & vbNewLine
         sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
         sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         snandan = sanadan & vbNewLine
          
        '√»Ê Ã⁄›—
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / √»Ê Ã⁄›—" & vbNewLine
        sanadan = sanadan & "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… «»‰ Ê—œ«‰ : ›ÕœÀ‰« »Â« «·‘ÌŒ √»Ê Õ›’ ⁄„— »‰ «·Õ”‰ »‰ „“Ìœ «·„—«€Ì »ﬁ—«¡ Ì ⁄·ÌÂ ﬁ«· : √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ √Õ„œ »‰ ⁄»œ «·Ê«Õœ «·”⁄œÌ „‘«›Â… ⁄‰ «·≈„«„ √»Ì «·Ì„‰ “Ìœ »‰ «·Õ”‰ «··€ÊÌ ° ﬁ«· : √Œ»—‰« √»Ê „Õ„œ ⁄»œ «··Â »‰ ⁄·Ì «·»€œ«œÌ √Œ»—‰« «·‘—Ì› √»Ê «·›÷· ⁄»œ «·ﬁ«Â— »‰ ⁄»œ «·”·«„ «·⁄»«”Ì ° √Œ»—‰« √»Ê ⁄»œ «··Â „Õ„œ »‰ «·Õ”Ì‰ «·ﬂ«—“Ì‰Ì ° √Œ»—‰« √»Ê «·›—Ã „Õ„œ »‰ √Õ„œ »‰ ≈»—«ÂÌ„ «·‘ÿÊÌ ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ √Õ„œ »‰ Â«—Ê‰ «·—«“Ì ° √Œ»—‰« √»Ê «·⁄»«” «·›÷· »‰ ‘«–«‰ »‰ ⁄Ì”Ï «·—«“Ì √Œ»—‰« √»Ê «·Õ”‰ √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì °√Œ»—‰« ⁄Ì”Ï »‰ „Ì‰« ﬁ«·Ê‰ ° √Œ»—‰« ⁄Ì”Ï »‰ Ê—œ«‰." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â „Õ„œ ⁄»œ «·—Õ„‰ »‰ ⁄·Ì «·‰ÕÊÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï «·≈„‹‹«„ √»Ì ⁄»œ „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„’—Ì ° ﬁ«· : ﬁ—√  »Â« «·ﬁ—¬‰ ⁄·Ï «·ﬂ„«· ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ›«—” «· „Ì„Ì ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·Ì„‰ «·ﬂ‰œÌ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «·≈„«„ √»Ì „‰’Ê— „Õ„œ »‰ ⁄»œ «·„·ﬂ »‰ «·Õ”‰ »‰ ŒÌ—Ê‰ «·»€œ«œÌ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·ﬁ«”„ ⁄»œ «·”Ìœ »‰ ⁄ «» «·„ﬁ—Ì¡ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ÿ«Â— „Õ„œ »‰ Ì«”Ì‰ «·Õ·»Ì ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·›—Ã «·‘ÿÊÌ ﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì »ﬂ— »‰ Â«—Ê‰ ° ﬁ«·: ﬁ—√  »Â« ⁄·Ï «·›÷· »‰ ‘«–«‰ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «·Õ·Ê«‰Ì ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï ﬁ«·Ê‰ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «»‰ Ê—œ«‰ . " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… «»‰ Ã„«“ : ›ÕœÀ‰« »Â« √»Ê ≈”Õ«ﬁ ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ≈»—«ÂÌ„ »‰ Õ« „ «·Ã–«„Ì »ﬁ—«¡ Ì ⁄·ÌÂ ⁄‰ √»Ì Õ›’ ⁄„— »‰ €‹œÌ— »‰ «·ﬁÊ«” «·œ„‘ﬁÌ ° √‰»√‰« √»Ê «·Ì„‰ »‰ «·Õ”‰ «·»€œ«œÌ ° √Œ»—‰« √»Ê „Õ„œ ”»ÿ «·ŒÌ«ÿ ° √Œ»—‰« «·√” «– √»Ê «·⁄“ „Õ„œ »‰ «·Õ”Ì‰ »‰ »‰œ«— «·Ê«”ÿÌ ° √Œ»—‰« «·≈„«„ √»Ê «·ﬁ«”„ ÌÊ”› »‰ Ã»«—… «·Â–·Ì ° √Œ»—‰« √»Ê ‰’— „‰’Ê— »‰ „Õ„œ «·ﬁÂ‰œ“Ì ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ «·Œ»«“Ì ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ «·›÷· «·ÃÊÂ—Ì ° √Œ»—‰« „Õ„œ »‰ √Õ„œ »‰ «·Õ”‰ «·Àﬁ›Ì «·ﬂ”«∆Ì ° √Œ»—‰« „Õ„œ »‰ ⁄»œ «··Â »‰ ‘«ﬂ— «·’Ì—›Ì ° √Œ»—‰« √»Ê «·⁄»«” √Õ„œ »‰ ”Â· «·ÿÌ«‰ ° √Œ»—‰« √»Ê ⁄„—«‰ „Ê”Ï »‰ ⁄»œ «·—Õ„‰ «·»“«“ ° √Œ»—‰« „Õ„œ »‰ ⁄Ì”Ï »‰ ≈»—«ÂÌ„ »‰ —“Ì‰ «·√’»Â«‰Ì ° √Œ»—‰« ”·Ì„«‰ »‰ œ«Êœ »‰ ⁄·Ì »‰ ⁄»œ «··Â »‰ ⁄»«” «·Â«‘„Ì ° √Œ»—‰« ≈”„«⁄Ì· »‰ Ã⁄›— »‰ √»Ì ﬂÀÌ— «·„œ‰Ì ° √Œ»—‰« ”·Ì„«‰ »‰ „”·„ «»‰ Ã„«“." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì „Õ„œ »‰ ⁄»œ «·—Õ„‰ «·Õ‰›Ì ° Êﬁ—√ »Â« «·ﬁ—«‰ ﬂ·Â ⁄·Ï „Õ„œ »‰ √Õ„œ «·’«∆€ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ≈”Õ«ﬁ »‰ ›«—” ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Ì„‰ ° Êﬁ—√ »Â« ⁄·Ï ”»ÿ «·ŒÌ«ÿ ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì ÿ«Â— √Õ„œ »‰ ⁄·Ì »‰ ⁄»Ìœ «··Â »‰ ”Ê«— ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄·Ì «·Õ”‰ »‰ «·›÷· «·‘—„ﬁ«‰Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄»œ «··Â »‰ «·„“—»«‰ «·√’»Â«‰Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄„— „Õ„œ »‰ √Õ„œ »‰ ⁄„— «·Œ—ﬁÌ ° Êﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ Ã⁄›— »‰ „Õ„Êœ «·√‘‰«‰Ì ° Êﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ «·Àﬁ›Ì «·ﬂ”«∆Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ ‘«ﬂ— ° Êﬁ—√ »Â« ⁄·Ï «»‰ ”Â· «·ÿÌ«‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄„—«‰ «·»“«“ ° Êﬁ—√ »Â« ⁄·Ï «»‰ —“Ì‰ ° Êﬁ—√ »Â« ⁄·Ï «·Â«‘„Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ Ã⁄›— ° Êﬁ—√ »Â« ⁄·Ï «»‰ Ã„«“ ° Êﬁ—√ «»‰ Ã„«“ ° Ê«»‰ Ê—œ«‰ ° ⁄·Ï √»Ì Ã⁄›— ." & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· √»Ì Ã⁄›— À·«À… : „Ê·«Â ⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° Ê√»Ê Â—Ì—… ° Ê«»‰ ⁄»«” . Êﬁ—√ Âƒ·«¡ «·À·«À… ⁄·Ï √»Ì »‰ ﬂ⁄» ° Êﬁ—√ √»Ê Â—Ì—… ° Ê«»‰ ⁄»«” ° √Ì÷« ⁄·Ï “Ìœ »‰ À«»  . Ê√Œ– “Ìœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ -° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
          
        'Ì⁄ﬁÊ»
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / Ì⁄ﬁÊ» «·»’—Ï" & vbNewLine
        sanadan = sanadan & "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… —ÊÌ” : ›ÕœÀ‰« »Â« «·‘ÌŒ «·≈„«„ √»Ê «·⁄»«” √Õ„œ »‰ „Õ„œ »‰ «·Œ÷— «·Õ‰›Ì »ﬁ—«¡ Ì ⁄·ÌÂ ﬁ«·: √Œ»—‰« : √»Ê «·⁄»«” √Õ„œ »‰ √»Ì ÿ«·» »‰ √»Ì «·‰⁄„ «·’«·ÕÌ ﬁ—«¡… ⁄·ÌÂ ° √Œ»—‰« √»Ê ÿ«·» ⁄»œ «··ÿÌ› »‰ „Õ„œ »‰ «·ﬁ»ÌÿÌ ° ›Ì ﬂ «»Â √Œ»—‰« »Â« √»Ê »ﬂ— √Õ„œ »‰ «·„ﬁ—» «·ﬂ—ŒÌ ﬁ—«¡… ⁄·ÌÂ ° √Œ»—‰« √»Ê ÿ«Â— √Õ„œ »‰ ⁄·Ì «·„ﬁ—Ì¡ «·√” «– √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ⁄·Ì «·ŒÌ«ÿ ° √Œ»—‰« «·√” «– «·≈„«„ √»Ê «·Õ”‰ ⁄·Ì »‰ √Õ„œ »‰ ⁄„— «·Õ„«„Ì ° √Œ»—‰« √»Ê «·ﬁ«”„ ⁄»œ «··Â »‰ «·Õ”‰ »‰ ”·Ì„«‰ «·‰Œ«” ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ Â«—Ê‰ »‰ ‰«›⁄ «· „«— «·»€œ«œÌ ° √Œ»—‰« √»Ê ⁄»œ «··Â „Õ„œ »‰ «·„ Êﬂ· «·„⁄—Ê› »—ÊÌ” ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì „Õ„œ ⁄»œ «·—Õ„‰ »‰ √Õ„œ »‰ ⁄·Ì «·»€œ«œÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï «·≈„«„ «· ﬁÌ „Õ„œ »‰ √Õ„œ «·„’—Ì ° Êﬁ—√ »Â« ⁄·Ï ≈»—«ÂÌ„ »‰ √Õ„œ «·≈”ﬂ‰œ—Ì ° Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï ⁄»œ «··Â »‰ ⁄·Ì «·»€œ«œÌ ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì «·⁄“ «·ﬁ·«‰”Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄·Ì «·Õ”‰ »‰ «·ﬁ«”„ «·Ê«”ÿÌ ° Êﬁ—√ »Â« ⁄·Ï : «·Õ„«„Ì ° Êﬁ—√ »Â« ⁄·Ï «·‰Œ« ” ° Êﬁ—√ »Â« ⁄·Ï «· „«— ° Êﬁ—√ ⁄·Ï —ÊÌ” ° Êﬁ—√ »Â« ⁄·Ï Ì⁄ﬁÊ» . " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… —ÛÊÕ : ›ÕœÀ‰« »Â« «·‘ÌŒ √»Ê «·⁄»«” √Õ„œ »‰ „Õ„œ »‰ «·Õ”Ì‰ «·‘Ì—«“Ì »ﬁ—«¡ Ì ⁄·ÌÂ ⁄‰ «·≈„«„ √»Ì «·Õ”‰ ⁄·Ì »‰ √Õ„œ «·„ﬁœ”Ì ° √Œ»—‰« √»Ê «·Ì„‰ «·ﬂ‰œÌ ‘›«Â« ° √Œ»—‰« √»Ê „Õ„œ «·»€œ«œÌ ° √Œ»—‰« √»Ê «·›÷· «·‘—Ì› «·„ﬂÌ ° √Œ»—‰« „Õ„œ »‰ «·Õ”Ì‰ «·›«—”Ì ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ≈»—«ÂÌ„ »‰ Œ‘‰«„ «·„«·ﬂÌ «·»’—Ì √Œ»—‰« √»Ê «·⁄»«” „Õ„œ »‰ Ì⁄ﬁÊ» »‰ «·ÕÃ«Ã »‰ „⁄«ÊÌ… «· Ì„Ì ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ ÊÂ» »‰ ÌÕÌÏ »‰ «·⁄·«¡ «·Àﬁ›Ì «·ﬁ“«“ ° √Œ»—‰« —ÊÕ »‰ ⁄»œ «·„ƒ„‰ «·»’—Ì ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï „Õ„œ »‰ √Õ„œ »«·ﬁ«Â—… «·„Õ—Ê”… ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â «·’«∆€ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ≈”Õ«ﬁ «·œ„‘ﬁÌ Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï ⁄»œ «··Â »‰ ⁄·Ì ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì ÿ«Â— »‰ ”Ê«— ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·ﬁ«”„ «·„”«›— »‰ «·ÿÌ» »‰ ⁄»«œ «·»’—Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ Œ‘‰«„ ° Êﬁ—√ »Â« ⁄·Ï «»‰ ⁄»« ” «· Ì„Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ ÊÂ» ° Êﬁ—√ »Â« ⁄·Ï —ÊÕ ° Êﬁ—√ »Â« ⁄·Ï Ì⁄ﬁÊ» ." & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· Ì⁄ﬁÊ» «·–Ì‰ ”„«Â„ √—»⁄… : √»Ê «·„‰–— ”·«„ »‰ ”·Ì„«‰ «·ÿÊÌ· ° Ê‘Â«» »‰ ‘—‰›… ° Ê„ÂœÌ »‰ „Ì„Ê‰ ° Ê√»Ê «·√‘Â» Ã⁄›— »‰ ÕÌ«‰ «·⁄ÿ«—œÌ .ÊﬁÌ· ≈‰ Ì⁄ﬁÊ» ﬁ—√ ⁄·Ï √»Ì ⁄„—Ê »‰ «·⁄·«¡ Êﬁ—√ ”·«„ ⁄·Ï ⁄«’„ Ê√»Ì ⁄„—Ê ° Êﬁ‹‹‹—√ ‘Â«» «·ÃÕœ—Ì Êﬁ—√ ⁄«’„ ⁄·Ï «·Õ”‰ «·»’—Ì Ê⁄·Ï ”·Ì„«‰ »‰ ﬁ … Êﬁ—√ ”·Ì„«‰ ⁄·Ï «»‹‰ ⁄»« ” Êﬁ—√ „ÂœÌ ⁄·Ï ‘⁄Ì» »‰ «·Õ»Õ«» Êﬁ—√ ⁄·Ï √»Ì «·⁄«·Ì… «·—Ì«ÕÌ Êﬁ—√ ⁄·Ï √»Ì Ê“Ìœ Êﬁ—√ √»Ê «·√‘Â» ⁄·Ï √»Ì —Ã«¡ ⁄„—«‰ »‰ „·Õ«‰ «·⁄ÿ«—œÌ Êﬁ—√ ⁄·Ï √»Ì „Ê”‹‹‹Ï «·√‘⁄—Ì Êﬁ—√ ⁄·Ï —”Ê· «··Â ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
  
  
         'Œ·›
         sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / Œ·› «·»“«—" & vbNewLine
         sanadan = sanadan & "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
         sanadan = sanadan & "√„« —Ê«Ì… ≈œ—Ì” «·Ê—«ﬁ : ›ÕœÀ‰« »Â« √»Ê Õ›’ ⁄„— »‰ «·Õ”‰ »ﬁ—«¡ Ì ⁄·ÌÂ Ÿ«Â— œ„‘ﬁ ° ⁄‰ ‘ÌŒÂ «·≈„«„ «·ŒÿÌ» √»Ì «·⁄»«” √Õ„œ »‰ ≈»—«ÂÌ„ »‰ ⁄„— «·›«—Ê∆Ì «·‘«›⁄Ì ° ﬁ«· : √Œ»—‰« Ê«·œÌ ° ﬁ«· : √Œ»—‰« √»Ê «·”⁄«œ«  «·√”⁄œ »‰ ”·ÿ«‰ «·Ê«”ÿÌ ° √Œ»—‰« √»Ê «·⁄“ „Õ„œ »‰ «·Õ”Ì‰ «·Ê«”ÿÌ ° √Œ»—‰« √»Ê «·Õ”Ì‰ √Õ„œ »‰ ⁄»œ «··Â »‰ «·Œ÷— «·”Ê”‰Ã—œÌ ° √Œ»—‰« √»Ê «·Õ”‰ „Õ„œ »‰ ⁄»œ «··Â »‰ „Õ„œ »‰ „—… «·ÿÊ”Ì «·„⁄—Ê› »«»‰ √»Ì ⁄„— «·‰ﬁ«‘ ° √Œ»—‰« √»Ê Ì⁄ﬁÊ» ≈”Õ«ﬁ »‰ ≈»—«ÂÌ„ «·Ê—«ﬁ ." & vbNewLine
         sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ﬂ· „‰ «·‘ÌŒÌ‰ √»Ì ⁄»œ «··Â «·Õ‰›Ì ° Ê√»Ì „Õ„œ «·‘«›⁄Ì «·„’—ÌÌ‰ ° Êﬁ—√ ﬂ· „‰Â„« ⁄·Ï √»Ì ⁄»œ «··Â „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„’—Ì ° Êﬁ—√ »Â« ⁄·Ï «·ﬂ„«· »‰ ›«—” ° Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·ﬁ«”„ Â»… «··Â »‰ √Õ„œ »‰ «·ÿ»— «·»€œ«œÌ ° Êﬁ—√ »Â« ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄·Ì »‰ „Ê”Ï «·ŒÌ«ÿ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Õ”Ì‰ «·”Ê”‰Ã—œÌ ° Êﬁ—√ »Â« ⁄·Ï «»‰ √»Ì ⁄„— «·ÿÊ”Ì ° Êﬁ—√ »Â« ⁄·Ï ≈”Õ«ﬁ «·Ê—«ﬁ ° Êﬁ—√ »Â« ⁄·Ï Œ·› ." & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… ≈œ—Ì” : ›ÕœÀ‰« »Â« √Õ„œ »‰ „Õ„œ »‰ «·Õ”Ì‰ «·›«—”Ì »ﬁ—«¡ Ì ⁄·ÌÂ ° √Œ»—‰« ⁄·Ì »‰ √Õ„œ ›Ì„« ‘«›Â‰Ì »Â °⁄‰ “Ìœ »‰ «·Õ”‰ «·»€œ«œÌ ° √Œ»—‰« √»Ê «·ﬁ«”„ »‰ √Õ„œ «·Õ—Ì—Ì ° √Œ»—‰« √»Ê »ﬂ—„Õ„œ »‰ ⁄»Ì »‰ „Õ„œ «·ŒÌ«ÿ ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ⁄»œ «··Â «·Õ–«¡ ° √Œ»—‰« √»Ê ≈”Õ«ﬁ ≈»—«ÂÌ„ »‰ «·Õ”Ì‰ »‰ ⁄»œ «··Â «·‰”«Ã «·„⁄—Ê› »«·‘ÿÌ ° √Œ»—‰« ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ «·Õœ«œ." & vbNewLine
         sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·‘ÌŒ √»Ì „Õ„œ ⁄»œ «·—Õ„‰ »‰ √Õ„œ «·Ê«”ÿÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„⁄œ· ° Êﬁ—√ »Â« ⁄·Ï ≈»—«ÂÌ„ »‰ √Õ„œ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Ì„‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì „Õ„œ ”»ÿ «·ŒÌ«ÿ ° ﬁ«· : Êﬁ—√  »Â« «·ﬁ—¬‰ „‰ √Ê·Â ≈·Ï ¬Œ—Â ⁄·Ï «·≈„«„Ì‰ «·‘—Ì› √»Ì «·›÷· ⁄»œ «·ﬁ«Â— »‰ ⁄»œ «·”·«„ «·⁄»«”Ì ° Ê√»Ì «·„⁄«·Ì À«»  »‰ »‰œ«— »‰ ≈»—«ÂÌ„ «·»ﬁ«· ° ›√„« «·‘—Ì› ›√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â „Õ„œ »‰ «·Õ”Ì‰ «·ﬂ«—“Ì‰Ì ° Ê√Œ»—Â √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ √»Ì «·⁄»«” «·Õ”‰ »‰ ”⁄Ìœ »‰ Ã⁄›— «·„ÿÊ⁄Ì ° Ê√„« √»Ê «·„⁄«·Ì ›√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ «·ﬁ«÷Ì √»Ì «·⁄·«¡ „Õ„œ »‰ ⁄·Ì »‰ Ì⁄ﬁÊ» «·Ê«”ÿÌ ° Êﬁ—√ «·Ê«”ÿÌ »Â« „‰ «·ﬂ «» ⁄·Ï «·≈„«„ √»Ì »ﬂ— √Õ„œ »‰ Ã⁄›— »‰ Õ„œ«‰ »‰ „«·ﬂ «·ﬁÿÌ⁄Ì ° Êﬁ—√ «·ﬁÿÌ⁄Ì Ê«·„ÿÊ⁄Ì Ã„Ì⁄« ⁄·Ï ≈œ—Ì” ° Êﬁ—√ ≈œ—Ì” ⁄·Ï Œ·› ° Ê«··Â «·„Ê›ﬁ . " & vbNewLine
         sanadan = sanadan & "Ê—Ã«· Œ·› : Ê—Ã«· Œ·› ”·Ì„ ’«Õ» Õ„“… ° ÊÌ⁄ﬁÊ» »‰ Œ·Ì›… «·√⁄‘Ï ’«Õ» √»Ì »ﬂ— ° Ê√»Ê “Ìœ ”⁄Ìœ ”⁄Ìœ »‰ √Ê” «·√‰’«—Ì ’«Õ» «·„›÷· «·÷»Ì Ê√»«‰ «·⁄ÿ«— ° Êﬁ—√ √»Ê »ﬂ— ° Ê«·„›÷· ° Ê√»«‰ ⁄·Ï ⁄«’„ . Ê—ÊÏ «·ﬁ—«¡… √Ì÷« ⁄‰ «·ﬂ”«∆Ì Ê⁄‰ ÌÕÌÏ »‰ ¬œ„ ⁄‰ √»Ì »ﬂ— ° Ê«··Â «·„Ê›ﬁ . ﬁ·  : Ê√Œ– ⁄«’„ ⁄‰ √»Ì ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ì „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·» ° Ê√»Ì »‰ ﬂ⁄» ° Ê“Ìœ »‰ À«»  ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         sanadan = sanadan & "Ê√Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰ ° Ê«»‰ „”⁄Êœ ° ⁄‰ —”Ê· «··Â ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -. Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ . Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         snandan = sanadan & vbNewLine
 
        ElseIf index = -4 Then
        
        'ﬁ«·Ê‰
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan + "√„« —Ê«Ì… ﬁ«·Ê‰ : ›ÕœÀ‰« »Â« √Õ„œ »‰ ⁄„— »‰ „Õ„œ «·ÃÌ“Ì ° ﬁ«·: ÕœÀ‰« „Õ„œ »‰ √Õ„œ »‰ „‰Ì— ° ﬁ«·: ÕœÀ‰« ⁄»œ «··Â »‰ ⁄Ì”Ï «·„œ‰Ì ° ﬁ«·:ÕœÀ‰« ﬁ«·Ê‰ ⁄‰ ‰«›⁄° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ‘ÌŒÌ √»Ì «·› Õ ›«—” »‰ √Õ„œ »‰ „Ê”Ï »‰ ⁄„—«‰ ° «·„ﬁ—Ì¡ «·÷—Ì— ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—∆ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈»—«ÂÌ„ »‰ ⁄„— «·„ﬁ—∆ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”Ì‰ √Õ„œ »‰ ⁄À„«‰ »‰ Ã⁄›— »‰ »ÊÌ«‰ ° Êﬁ«·:ﬁ—√  ⁄·Ï √»Ì »ﬂ— √Õ„œ »‰ „Õ„œ »‰ «·√‘⁄À Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì ‰‘Ìÿ „Õ„œ »‰ Â«—Ê‰ ° Êﬁ«·: ﬁ—√  ⁄·Ï ﬁ«·Ê‰ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‰«›⁄ ." & vbNewLine
        sanadan = sanadan + "Ê—Ã«· ‰«›⁄ «·–Ì‰ ”„«Â„ Œ„”… : √»Ê Ã⁄›— Ì“ Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—∆ ° Ê√»Ê œ«Êœ ⁄»œ «·—Õ„‰ »‰ Â—„“ «·√⁄—Ã ° Ê‘Ì»… »‰ ‰’«Õ «·ﬁ«÷Ì ° Ê√»Ê ⁄»œ «··Â „”·„ »‰ Ã‰œ» «·Â–·Ì «·ﬁ«’ ° Ê√»Ê —ÊÕ Ì“Ìœ »‰ —Ê„«‰ ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄‰ √»Ì Â—Ì—… ° Ê«»‰ ⁄»«” ° Ê⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° ⁄‰ √»Ì »‰ ﬂ⁄» ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
     
        ' «»‰ ﬂÀÌ—
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / «»‰ ﬂÀÌ—" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—  " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… «·»“Ì : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ «·ﬂ« » ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ „Ê”Ï ° ﬁ«·: ÕœÀ‰« „÷— »‰ „Õ„œ «·÷»Ì ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ √»Ì »“… ° ﬁ«·: ﬁ—√  ⁄·Ï ⁄ﬂ—„… »‰ ”·Ì„«‰ »‰ ⁄«„— ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·ﬁ”ÿ ° Êﬁ«· : ﬁ—√  ⁄·Ï «»‰ ﬂÀÌ— ‰›”Â ° ﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·ﬁ«”„ ⁄»œ «·⁄“Ì“ »‰ Ã⁄›— »‰ „Õ„œ «·„ﬁ—Ì¡ «·›«—”Ì ° Êﬁ«· ·Ì: ﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ «·Õ”‰ «·‰ﬁ«‘ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì —»Ì⁄… „Õ„œ »‰ ≈”Õ«ﬁ «·— »⁄Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï «·»“Ì ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… ﬁ‰»· : ›ÕœÀ‰« »Â« √»Ê „”·„ „Õ„œ »‰ √Õ„œ «·»€œ«œÌ ° ﬁ«·: ÕœÀ‰« «»‰ „Ã«Âœ ° ﬁ«·: ﬁ—√  ⁄·Ï ﬁ‰»· ° Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì «·Õ”‰ √Õ„œ »‰ ⁄Ê‰ «·ﬁÊ«” Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì «·«Œ— Ìÿ ÊÂ» »‰ Ê«÷Õ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·ﬁ”ÿ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‘»· »‰ ⁄»«œ Ê „⁄—Ê› »‰ „‘ﬂ«‰ ° Êﬁ«·« ﬁ—√‰« ⁄·Ï «»‰ ﬂÀ‹Ì‹— ° Ê ﬁ«· √»‹‹‹‹Ê ⁄‹‹„‹‹—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ «·Õ„’Ì «·„ﬁ—Ì¡ «·÷—Ì— Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·»€œ«œÌ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï «»‰ „Ã«Âœ Êﬁ«·: ﬁ—√  ⁄·Ï ﬁ‰»· ." & vbNewLine
        sanadan = sanadan & " Ê—Ã‹‹«· «»‰ ﬂÀÌ— «·‹–Ì‹‰ ”„«Â„ À·«À… : ⁄»œ «··Â »‰ «·”«∆» «·„Œ“Ê„Ì ’«Õ» —”Ê· «··Â  Ê„Ã«Âœ »‰ Ã»— √»Ê «·ÕÃ«Ã „Ê·Ï ﬁÌ” »‰ «·”«∆» ° Êœ—»«” „Ê·Ï «»‰ ⁄»«” . Ê√Œ– ⁄»œ «··Â ⁄‰ √»Ì »‰ ﬂ⁄» ‰›”Â. Ê√Œ– „Ã«Âœ Êœ—»«”° ⁄‰ «»‰ ⁄»«”° ⁄‰ √»Ì ° Ê“Ìœ »‰ À«»  ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  °⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -."
      
        '√»Ê Ã⁄›—
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / √»Ê Ã⁄›—" & vbNewLine
        sanadan = sanadan & "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… «»‰ Ê—œ«‰ : ›ÕœÀ‰« »Â« «·‘ÌŒ √»Ê Õ›’ ⁄„— »‰ «·Õ”‰ »‰ „“Ìœ «·„—«€Ì »ﬁ—«¡ Ì ⁄·ÌÂ ﬁ«· : √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ √Õ„œ »‰ ⁄»œ «·Ê«Õœ «·”⁄œÌ „‘«›Â… ⁄‰ «·≈„«„ √»Ì «·Ì„‰ “Ìœ »‰ «·Õ”‰ «··€ÊÌ ° ﬁ«· : √Œ»—‰« √»Ê „Õ„œ ⁄»œ «··Â »‰ ⁄·Ì «·»€œ«œÌ √Œ»—‰« «·‘—Ì› √»Ê «·›÷· ⁄»œ «·ﬁ«Â— »‰ ⁄»œ «·”·«„ «·⁄»«”Ì ° √Œ»—‰« √»Ê ⁄»œ «··Â „Õ„œ »‰ «·Õ”Ì‰ «·ﬂ«—“Ì‰Ì ° √Œ»—‰« √»Ê «·›—Ã „Õ„œ »‰ √Õ„œ »‰ ≈»—«ÂÌ„ «·‘ÿÊÌ ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ √Õ„œ »‰ Â«—Ê‰ «·—«“Ì ° √Œ»—‰« √»Ê «·⁄»«” «·›÷· »‰ ‘«–«‰ »‰ ⁄Ì”Ï «·—«“Ì √Œ»—‰« √»Ê «·Õ”‰ √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì °√Œ»—‰« ⁄Ì”Ï »‰ „Ì‰« ﬁ«·Ê‰ ° √Œ»—‰« ⁄Ì”Ï »‰ Ê—œ«‰." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â „Õ„œ ⁄»œ «·—Õ„‰ »‰ ⁄·Ì «·‰ÕÊÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï «·≈„‹‹«„ √»Ì ⁄»œ „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„’—Ì ° ﬁ«· : ﬁ—√  »Â« «·ﬁ—¬‰ ⁄·Ï «·ﬂ„«· ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ›«—” «· „Ì„Ì ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·Ì„‰ «·ﬂ‰œÌ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «·≈„«„ √»Ì „‰’Ê— „Õ„œ »‰ ⁄»œ «·„·ﬂ »‰ «·Õ”‰ »‰ ŒÌ—Ê‰ «·»€œ«œÌ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·ﬁ«”„ ⁄»œ «·”Ìœ »‰ ⁄ «» «·„ﬁ—Ì¡ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ÿ«Â— „Õ„œ »‰ Ì«”Ì‰ «·Õ·»Ì ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·›—Ã «·‘ÿÊÌ ﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì »ﬂ— »‰ Â«—Ê‰ ° ﬁ«·: ﬁ—√  »Â« ⁄·Ï «·›÷· »‰ ‘«–«‰ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «·Õ·Ê«‰Ì ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï ﬁ«·Ê‰ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «»‰ Ê—œ«‰ . " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… «»‰ Ã„«“ : ›ÕœÀ‰« »Â« √»Ê ≈”Õ«ﬁ ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ≈»—«ÂÌ„ »‰ Õ« „ «·Ã–«„Ì »ﬁ—«¡ Ì ⁄·ÌÂ ⁄‰ √»Ì Õ›’ ⁄„— »‰ €‹œÌ— »‰ «·ﬁÊ«” «·œ„‘ﬁÌ ° √‰»√‰« √»Ê «·Ì„‰ »‰ «·Õ”‰ «·»€œ«œÌ ° √Œ»—‰« √»Ê „Õ„œ ”»ÿ «·ŒÌ«ÿ ° √Œ»—‰« «·√” «– √»Ê «·⁄“ „Õ„œ »‰ «·Õ”Ì‰ »‰ »‰œ«— «·Ê«”ÿÌ ° √Œ»—‰« «·≈„«„ √»Ê «·ﬁ«”„ ÌÊ”› »‰ Ã»«—… «·Â–·Ì ° √Œ»—‰« √»Ê ‰’— „‰’Ê— »‰ „Õ„œ «·ﬁÂ‰œ“Ì ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ «·Œ»«“Ì ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ «·›÷· «·ÃÊÂ—Ì ° √Œ»—‰« „Õ„œ »‰ √Õ„œ »‰ «·Õ”‰ «·Àﬁ›Ì «·ﬂ”«∆Ì ° √Œ»—‰« „Õ„œ »‰ ⁄»œ «··Â »‰ ‘«ﬂ— «·’Ì—›Ì ° √Œ»—‰« √»Ê «·⁄»«” √Õ„œ »‰ ”Â· «·ÿÌ«‰ ° √Œ»—‰« √»Ê ⁄„—«‰ „Ê”Ï »‰ ⁄»œ «·—Õ„‰ «·»“«“ ° √Œ»—‰« „Õ„œ »‰ ⁄Ì”Ï »‰ ≈»—«ÂÌ„ »‰ —“Ì‰ «·√’»Â«‰Ì ° √Œ»—‰« ”·Ì„«‰ »‰ œ«Êœ »‰ ⁄·Ì »‰ ⁄»œ «··Â »‰ ⁄»«” «·Â«‘„Ì ° √Œ»—‰« ≈”„«⁄Ì· »‰ Ã⁄›— »‰ √»Ì ﬂÀÌ— «·„œ‰Ì ° √Œ»—‰« ”·Ì„«‰ »‰ „”·„ «»‰ Ã„«“." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì „Õ„œ »‰ ⁄»œ «·—Õ„‰ «·Õ‰›Ì ° Êﬁ—√ »Â« «·ﬁ—«‰ ﬂ·Â ⁄·Ï „Õ„œ »‰ √Õ„œ «·’«∆€ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ≈”Õ«ﬁ »‰ ›«—” ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Ì„‰ ° Êﬁ—√ »Â« ⁄·Ï ”»ÿ «·ŒÌ«ÿ ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì ÿ«Â— √Õ„œ »‰ ⁄·Ì »‰ ⁄»Ìœ «··Â »‰ ”Ê«— ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄·Ì «·Õ”‰ »‰ «·›÷· «·‘—„ﬁ«‰Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄»œ «··Â »‰ «·„“—»«‰ «·√’»Â«‰Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄„— „Õ„œ »‰ √Õ„œ »‰ ⁄„— «·Œ—ﬁÌ ° Êﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ Ã⁄›— »‰ „Õ„Êœ «·√‘‰«‰Ì ° Êﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ «·Àﬁ›Ì «·ﬂ”«∆Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ ‘«ﬂ— ° Êﬁ—√ »Â« ⁄·Ï «»‰ ”Â· «·ÿÌ«‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄„—«‰ «·»“«“ ° Êﬁ—√ »Â« ⁄·Ï «»‰ —“Ì‰ ° Êﬁ—√ »Â« ⁄·Ï «·Â«‘„Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ Ã⁄›— ° Êﬁ—√ »Â« ⁄·Ï «»‰ Ã„«“ ° Êﬁ—√ «»‰ Ã„«“ ° Ê«»‰ Ê—œ«‰ ° ⁄·Ï √»Ì Ã⁄›— ." & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· √»Ì Ã⁄›— À·«À… : „Ê·«Â ⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° Ê√»Ê Â—Ì—… ° Ê«»‰ ⁄»«” . Êﬁ—√ Âƒ·«¡ «·À·«À… ⁄·Ï √»Ì »‰ ﬂ⁄» ° Êﬁ—√ √»Ê Â—Ì—… ° Ê«»‰ ⁄»«” ° √Ì÷« ⁄·Ï “Ìœ »‰ À«»  . Ê√Œ– “Ìœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ -° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
       
         ElseIf index = -5 Then
        
         ' ‰«›⁄
        sanadan = "”‰œ ﬁ—«¡… «·≈„«„ / ‰«›⁄" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "√„« —Ê«Ì… ﬁ«·Ê‰ : ›ÕœÀ‰« »Â« √Õ„œ »‰ ⁄„— »‰ „Õ„œ «·ÃÌ“Ì ° ﬁ«·: ÕœÀ‰« „Õ„œ »‰ √Õ„œ »‰ „‰Ì— ° ﬁ«·: ÕœÀ‰« ⁄»œ «··Â »‰ ⁄Ì”Ï «·„œ‰Ì ° ﬁ«·:ÕœÀ‰« ﬁ«·Ê‰ ⁄‰ ‰«›⁄° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ‘ÌŒÌ √»Ì «·› Õ ›«—” »‰ √Õ„œ »‰ „Ê”Ï »‰ ⁄„—«‰ ° «·„ﬁ—Ì¡ «·÷—Ì— ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—∆ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈»—«ÂÌ„ »‰ ⁄„— «·„ﬁ—∆ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”Ì‰ √Õ„œ »‰ ⁄À„«‰ »‰ Ã⁄›— »‰ »ÊÌ«‰ ° Êﬁ«·:ﬁ—√  ⁄·Ï √»Ì »ﬂ— √Õ„œ »‰ „Õ„œ »‰ «·√‘⁄À Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì ‰‘Ìÿ „Õ„œ »‰ Â«—Ê‰ ° Êﬁ«·: ﬁ—√  ⁄·Ï ﬁ«·Ê‰ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‰«›⁄ ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… Ê—‘ : ›ÕœÀ‰« »Â« √»Ê ⁄»œ «··Â √Õ„œ »‰ „Õ›ÊŸ «·ﬁ«÷Ì »„’— ° ﬁ«·: ÕœÀ‰« √Õ„œ »‰ ≈»—«ÂÌ„ »‰ Ã«„⁄ ° ﬁ«· : ÕœÀ‰« √»Ê „Õ„œ »ﬂ— »‰ ”Â· ° ﬁ«·: ÕœÀ‰« √»Ê „Õ„œ ⁄»œ «·’„œ »‰ ⁄»œ «·—Õ„‰ ° ﬁ«· : ÕœÀ‰« Ê—‘ ⁄‰ ‰«›⁄ ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ‘ÌŒÌ √»Ì «·ﬁ«”„ Œ·› »‰ ≈»—«ÂÌ„ »‰ „Õ„œ »‰ Œ«ﬁ«‰ «·„ﬁ—Ì¡ »„’— ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ⁄·Ï √»Ì Ã⁄›— √Õ„œ »‰ √”«„… «· ÃÌ»Ì ° Êﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·‰Õ«” ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì Ì⁄ﬁÊ» ÌÊ”› »‰ ⁄„—Ê »‰ Ì”«— «·√“—ﬁ ° Êﬁ«· :ﬁ—√  ⁄·Ï Ê—‘ Êﬁ«· : ﬁ—√  ⁄·Ï ‰«›⁄ ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· ‰«›⁄ «·–Ì‰ ”„«Â„ Œ„”… : √»Ê Ã⁄›— Ì“ Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—∆ ° Ê√»Ê œ«Êœ ⁄»œ «·—Õ„‰ »‰ Â—„“ «·√⁄—Ã ° Ê‘Ì»… »‰ ‰’«Õ «·ﬁ«÷Ì ° Ê√»Ê ⁄»œ «··Â „”·„ »‰ Ã‰œ» «·Â–·Ì «·ﬁ«’ ° Ê√»Ê —ÊÕ Ì“Ìœ »‰ —Ê„«‰ ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄‰ √»Ì Â—Ì—… ° Ê«»‰ ⁄»«” ° Ê⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° ⁄‰ √»Ì »‰ ﬂ⁄» ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
      
         ' «»‰ ﬂÀÌ—
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / «»‰ ﬂÀÌ—" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—  " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… «·»“Ì : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ «·ﬂ« » ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ „Ê”Ï ° ﬁ«·: ÕœÀ‰« „÷— »‰ „Õ„œ «·÷»Ì ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ √»Ì »“… ° ﬁ«·: ﬁ—√  ⁄·Ï ⁄ﬂ—„… »‰ ”·Ì„«‰ »‰ ⁄«„— ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·ﬁ”ÿ ° Êﬁ«· : ﬁ—√  ⁄·Ï «»‰ ﬂÀÌ— ‰›”Â ° ﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·ﬁ«”„ ⁄»œ «·⁄“Ì“ »‰ Ã⁄›— »‰ „Õ„œ «·„ﬁ—Ì¡ «·›«—”Ì ° Êﬁ«· ·Ì: ﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ «·Õ”‰ «·‰ﬁ«‘ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì —»Ì⁄… „Õ„œ »‰ ≈”Õ«ﬁ «·— »⁄Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï «·»“Ì ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… ﬁ‰»· : ›ÕœÀ‰« »Â« √»Ê „”·„ „Õ„œ »‰ √Õ„œ «·»€œ«œÌ ° ﬁ«·: ÕœÀ‰« «»‰ „Ã«Âœ ° ﬁ«·: ﬁ—√  ⁄·Ï ﬁ‰»· ° Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì «·Õ”‰ √Õ„œ »‰ ⁄Ê‰ «·ﬁÊ«” Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì «·«Œ— Ìÿ ÊÂ» »‰ Ê«÷Õ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·ﬁ”ÿ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‘»· »‰ ⁄»«œ Ê „⁄—Ê› »‰ „‘ﬂ«‰ ° Êﬁ«·« ﬁ—√‰« ⁄·Ï «»‰ ﬂÀ‹Ì‹— ° Ê ﬁ«· √»‹‹‹‹Ê ⁄‹‹„‹‹—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ «·Õ„’Ì «·„ﬁ—Ì¡ «·÷—Ì— Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·»€œ«œÌ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï «»‰ „Ã«Âœ Êﬁ«·: ﬁ—√  ⁄·Ï ﬁ‰»· ." & vbNewLine
        sanadan = sanadan & " Ê—Ã‹‹«· «»‰ ﬂÀÌ— «·‹–Ì‹‰ ”„«Â„ À·«À… : ⁄»œ «··Â »‰ «·”«∆» «·„Œ“Ê„Ì ’«Õ» —”Ê· «··Â  Ê„Ã«Âœ »‰ Ã»— √»Ê «·ÕÃ«Ã „Ê·Ï ﬁÌ” »‰ «·”«∆» ° Êœ—»«” „Ê·Ï «»‰ ⁄»«” . Ê√Œ– ⁄»œ «··Â ⁄‰ √»Ì »‰ ﬂ⁄» ‰›”Â. Ê√Œ– „Ã«Âœ Êœ—»«”° ⁄‰ «»‰ ⁄»«”° ⁄‰ √»Ì ° Ê“Ìœ »‰ À«»  ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  °⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
      
       ' √»Ê ⁄„—Ê
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / √»Ê ⁄„—Ê «·»’—Ï" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… √»Ì ⁄„— «·œÊ—Ì : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ ⁄·Ì ° ﬁ«·: √Œ»—‰« √»Ê ⁄Ì”Ï „Õ„œ »‰ √Õ„œ »‰ ﬁÿ‰ ”‰… À„«‰ ⁄‘—… ÊÀ·«À„«∆…° ﬁ«·: √Œ»—‰« √»Ê Œ·«œ ”·Ì„«‰ »‰ Œ·«œ ﬁ«·:ÕœÀ‰« «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê ° ﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â „‰ ÿ—Ìﬁ √»Ì ⁄„— «·œÊ—Ì ⁄·Ï ‘ÌŒ‰« ⁄»œ «·⁄“ Ì“ »‰ Ã⁄›— »‰ „Õ„œ »‰ ≈”Õ«ﬁ «·»€œ«œÌ «·›«—”Ì «·„ﬁ—Ì¡ ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì ÿ«Â— ⁄»œ «·Ê«Õœ »‰ ⁄„— »‰ √»Ì Â«‘„ «·„ﬁ—Ì¡ ° „« ·« √Õ’ÌÂ ﬂÀ—… ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì »ﬂ— »‰ „Ã«Âœ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·“⁄—«¡ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” Êﬁ«· :ﬁ—√  ⁄·Ï √»Ì ⁄„— ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï «·Ì“ÌœÌ ° Êﬁ«· ﬁ—√  »Â« ⁄·Ï : √»Ì ⁄„—Ê. " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… √»Ì ‘⁄Ì» «·”Ê”Ì : ›ÕœÀ‰« »Â« Œ·› »‰ ≈»—«ÂÌ„ »‰ „Õ„œ «·„ﬁ—∆ ° ﬁ«· : ÕœÀ‰« √»Ê „Õ„œ «·Õ”‰ »‰ —‘Ìﬁ «·„⁄œ· ° ﬁ«· : ÕœÀ‰« √»Ê ⁄»œ «·—Õ„‰ √Õ„œ »‰ ‘⁄Ì» «·‰”«∆Ì ° ﬁ«· : √Œ»—‰« √»Ê ‘⁄Ì» ° ﬁ«· : √Œ»—‰« «·Ì“ÌœÌ ° ⁄‰ √»Ì ⁄„—Ê ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â »≈ŸÂ«— «·√Ê· „‰ «·„À·Ì‰ Ê«·„ ﬁ«—»Ì‰ Ê»≈œ€«„Â ⁄·Ï ›«—” »‰ √Õ„œ «·„ﬁ—Ì¡ ° Êﬁ«· ·Ì : ﬁ—√  »Â« ﬂ–·ﬂ ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Êﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ﬂ·Â ﬂ–·ﬂ ⁄·Ï √»Ì ⁄„—«‰ „Ê”Ï »‰ Ã—Ì— «·‰ÕÊÌ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ‘⁄Ì» ° Êﬁ«· : ﬁ—√  ⁄·Ï «·Ì“ÌœÌ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì ⁄„—Ê" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê: ÊÕœÀ‰« »√’Ê· «·≈œ€«„ „Õ„œ »‰ √Õ„œ ⁄‰ «»‰ „Ã«Âœ ⁄‰ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” ⁄‰ «·œÊ—Ì ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ï ⁄„—Ê° ÊÕœÀ‰« »Â« √Ì÷« √»Ê «·Õ”‰ ‘ÌŒ‰« ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ «·„»«—ﬂ ⁄‰ Ã⁄›— »‰ ”·Ì„«‰ ⁄‰ √»Ì ‘⁄Ì» ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê . " & vbNewLine
        sanadan = sanadan & "Ê—Ã«· √»Ì ⁄„—Ê : Ã„«⁄… „‰ √Â· «·ÕÃ«“ Ê„‰ √Â· «·»’—… ° ›„‰ √Â· „ﬂ… : „Ã«Âœ ° Ê”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„… »‰ Œ«·œ ° Ê⁄ÿ«¡ »‰ √»Ì —»«Õ ° Ê⁄»œ «··Â »‰ ﬂÀÌ— ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ „ÕÌ’‰ ° ÊÕ„Ìœ »‰ ﬁÌ” «·√⁄—Ã «·ﬁ«—∆ ° Ê„‰ √Â· «·„œÌ‰… : Ì“Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—Ì¡ ÊÌ“Ìœ »‰ —Ê„«‰ ° Ê‘Ì»… »‰ ‰’«Õ ° Ê„‰ √Â· «·»’—… : «·Õ”‰ »‰ √»Ì «·Õ”‰ «·»’—Ì ° ÊÌÕÌ »‰ Ì⁄„— ° Ê€Ì—Â„« ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄„‰  ﬁœ„ „‰ «·’Õ«»… Ê€Ì—Â„ . " & vbNewLine
        sanadan = sanadan & "ﬁ·  : Ê√Œ– ”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„…° ÊÌÕÌÏ »‰ Ì⁄„— ° ⁄‰ «»‰ ⁄»«” Ê√Œ– «»‰ ⁄»«” ⁄‰ √»Ì »‰ ﬂ⁄» Ê“Ìœ »‰ À«»  ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        '«»‰ ⁄«„—
         sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / «»‰ ⁄«„—" & vbNewLine
         sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
         sanadan = sanadan & "›√„« —Ê«Ì… Â‘«„ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ﬁ«·: ÕœÀ‰« «»‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« «·Õ”Ì‰ »‰ „Â—«‰ «·Ã„«· ° ﬁ«· :ÕœÀ‰« √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì ° ﬁ«· : ÕœÀ‰« Â‘«„ »‰ ⁄„«— ° ﬁ«·: ÕœÀ‰« ⁄—«ﬂ »‰ Œ«·œ «·„—Ì ° ﬁ«· :ﬁ—√  ⁄·Ï ÌÕÌÌ »‰ «·Õ«—À «·–„«—Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï ⁄»œ «··Â »‰ ⁄«„— ° ﬁ«· : √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ ‘ÌŒ‰« ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Ê ﬁ«· : ﬁ—√  »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ⁄»œ«‰ ° Êﬁ«· : ﬁ—√  ⁄·Ï «·Õ·Ê«‰Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï Â‘«„ " & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… «»‰ –ﬂÊ«‰ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ „Ê”Ï »‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« √Õ„œ »‰ ÌÊ”› «· €·»Ì ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ –ﬂÊ«‰ ° ﬁ«· : ÕœÀ‰« √ÌÊ» »‰  „Ì„ «· „Ì„Ì ° ﬁ«· :ÕœÀ‰« ÌÕÌÌ »‰ «·Õ«—À «·–„«—Ì ° ﬁ«· : ﬁ—√  ⁄·Ï «»‰ ⁄«„— ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ⁄»œ «·⁄“Ì“ »‰ Ã⁄›— «·›«—”Ì «·„ﬁ—Ì¡ Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ï »ﬂ— „Õ„œ »‰ «·Õ”‰ «·‰ﬁ«‘ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ⁄»œ «··Â Â«—Ê‰ »‰ „Ê”Ï »‰ ‘—Ìﬂ «·√Œ›‘ Ê—Ê«Â« «·√Œ›‘ ⁄‰ ⁄»œ «··Â »‰ –ﬂÊ«‰ " & vbNewLine
         sanadan = sanadan & "Ê—Ã‹‹«· «»‰ ⁄«„— «·‹–Ì‹‰ ”‹‹„«Â„ : √»Ê «·œ—œ«¡ ⁄ÊÌ„— »‰ ⁄«„— ’«Õ» —”Ê· «··Â ° Ê«·„€Ì—… »‰ √»Ì ‘Â«» «·„Œ“Ê„Ì ° Ê√Œ‹– √»Ê «·œ—œ«¡ ⁄‹‹‰ «·‰»Ì . Ê√Œ– «·„€Ì—… ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -" & vbNewLine
         snandan = sanadan & vbNewLine
        
        
         '⁄«’„
         sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / ⁄«’„" & vbNewLine
         sanadan = sanadan & "ﬁ«· √»‹‹Ê ⁄‹„‹—Ê «·‹œ«‰‹‹Ì ›‹‹‹Ì «·‹ Ì”Ì—:" & vbNewLine
         sanadan = sanadan & "›√„« —Ê«Ì… √»Ì »ﬂ— ‘⁄»…: ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ ⁄·Ì «·ﬂ« » ﬁ«·: ÕœÀ‰« »‰ „Ã«Âœ ﬁ«·: ÕœÀ‰« ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ⁄„— «·ÊﬂÌ⁄Ì ° ﬁ«·:ÕœÀ‰« √»Ì ﬁ«·:ÕœÀ‰« ÌÕÌÌ »‰ √œ„ ° ﬁ«·: ÕœÀ‰« √»Ê »ﬂ— ⁄‰ ⁄«’„ ° ﬁ«· √»Ê ⁄„—Ê: Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ «·„ﬁ—Ì¡ ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—Ì¡ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈»—«ÂÌ„ »‰ ⁄»œ «·—Õ„‰ »‰ √Õ„œ «·„ﬁ—Ì¡ «·»€œ«œÌ Êﬁ«·: ﬁ—√  ⁄·Ï ÌÊ”› »‰ Ì⁄ﬁÊ» «·Ê«”ÿÌ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‘⁄Ì» »‰ √ÌÊ» «·’—Ì›Ì‰Ì ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ÌÕÌÌ »‰ √œ„ ⁄‰ √»Ï »ﬂ— ⁄‰ ⁄«’„." & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… Õ›’ : ›ÕœÀ‰« »Â« √»Ê «·Õ”‰ ÿ«Â‹— »‰ €·»Ê‰ «·„ﬁ—∆ ° ﬁ«· : ÕœÀ‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ’«·Õ «·Â«‘„Ì «·÷—Ì— «·„ﬁ—∆ »«·»’—… ° ﬁ«·: ÕœÀ‰« √»Ê «·⁄»«” √Õ„œ »‰ ”Â· «·√‘‰«‰Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì „Õ„œ ⁄»Ìœ »‰ «·’»«Õ ° Êﬁ«·: ﬁ—√  ⁄·Ï Õ›’ ° Êﬁ«·: ﬁ—√  ⁄·Ï ⁄«’‹„ ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ‘ÌŒ‰« √»Ì «·Õ”‰ Êﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï «·Â«‘„Ì Êﬁ«·: ﬁ—√  ⁄·Ï «·√‘‰«‰Ì ⁄‰ ⁄»Ìœ ⁄‰ Õ›’ ⁄‰ ⁄«’‹„ . " & vbNewLine
         sanadan = sanadan & "Ê—Ã«· ⁄«’„ «·‹–Ì‹‰ ”„«Â„ «À‰«‰ : √»Ê ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ê „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·»  ° Ê√»Ì »‰ ﬂ⁄»  ° Ê“Ìœ »‰ À«»   ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï - ° √Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰  ° Ê«»‰ „”⁄Êœ  ° ⁄‰ —”Ê· «··Â - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         snandan = sanadan & vbNewLine
         
        'Õ„“…
        sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / Õ„“…" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… Œ·› : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« «»‰ „Ã«Âœ ° ÕœÀ‰« ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ ° ﬁ«· : ÕœÀ‰« Œ·› ° ﬁ«·: ⁄‰ ”·Ì„ ⁄‰ Õ„“… ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·Õ”‰ ‘ÌŒ‰« ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ „Õ„œ »‰ ÌÊ”› »‰ ‰Â«— «·Õ— ﬂÌ »«·»’—… ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”Ì‰ √Õ„œ »‰ ⁄À„«‰ »‰ Ã⁄›— »‰ »ÊÌ«‰ ° Êﬁ«· ·Ì :ﬁ—√  ⁄·Ï ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ ﬁ»· √‰ Ìﬁ—Ì¡ »«Œ Ì«— Œ·› ° Êﬁ«· ·Ì : ﬁ—√  ⁄·Ï Œ·› ° Êﬁ«· : ﬁ—√  ⁄·Ï ”·Ì„ ° Ê ﬁ«· : ﬁ—√  ⁄·Ï Õ„“… ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… Œ·«œ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« √Õ„œ »‰ „Ê”Ï ° ﬁ«· : ÕœÀ‰« ÌÕÌÏ »‰ √Õ„œ »‰ Â«—Ê‰ «·„“Êﬁ ° ⁄‰ √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì ° ⁄‰ Œ·«œ ° ⁄‰ ”·Ì„ ° ⁄‰ Õ„“… ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ «·÷—Ì— ‘ÌŒ‰« ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ‘‰»Ê– ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ‘«–«‰ «·ÃÊÂ—Ì «·„ﬁ—Ì ° Êﬁ«· :ﬁ—√  ⁄·Ï Œ·«œ Êﬁ«· : ﬁ—√  ⁄·Ï ”·Ì„ ° Êﬁ—√ ”·Ì„ ⁄·Ï Õ„“…." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
        sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
       
        '«·ﬂ”«∆Ï
         sanadan = sanadan & "”‰œ ﬁ—«¡… «·≈„«„ / «·ﬂ”«∆Ï" & vbNewLine
         sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
         sanadan = sanadan & "›√„« —Ê«Ì… √»Ì ⁄„—Ê «·œÊ—Ì : ›ÕœÀ‰« »Â« √»Ê „Õ„œ ⁄»œ «·—Õ„‰ »‰ ⁄„— »‰ „Õ„œ «·„⁄œ· ° ﬁ«· : ÕœÀ‰« √»Ê ⁄„— ⁄»œ «··Â »‰ √Õ„œ »‰ œÌ“ÊÌÂ «·œ„‘ﬁÌ ° ﬁ«· : ÕœÀ‰« Ã⁄›— »‰ „Õ„œ »‰ √”œ «·‰’Ì»Ì ° ﬁ«· : ÕœÀ‰« √»Ê ⁄„— «·œÊ—Ì ° ⁄‰ «·ﬂ”«∆Ì ° Ê ﬁ«· √»Ê ⁄‹„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄·Ì »‰ «·Ã·‰œÌ «·„Ê’·Ì ° Ê ﬁ«· :ﬁ—√  ⁄·Ï Ã⁄›— »‰ „Õ„œ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì ⁄„— «·œÊ—Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï «·ﬂ”«∆Ì ." & vbNewLine
         sanadan = sanadan & "Ê√„« —Ê«Ì… √»Ì «·Õ«—À : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« »Â« «»‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« „Õ„œ »‰ ÌÕÌÏ ( «·ﬂ”«∆Ì «·’€Ì—) ° ⁄‰ √»Ì «·Õ«—À ° ⁄‰ «·ﬂ”«∆Ì ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—Ì¡ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·ﬁ«”„ “Ìœ »‰ ⁄·Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï √Õ„œ »‰ «·Õ”‰ «·„⁄—Ê› »«·»ÿÌ ° Êﬁ«· :ﬁ—√  ⁄·Ï „Õ„œ »‰ ÌÕÌÏ ( «·ﬂ”«∆Ì «·’€Ì—) ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì «·Õ«—À ° Êﬁ«· : ﬁ—√  ⁄·Ï «·ﬂ”«∆Ì ." & vbNewLine
         sanadan = sanadan & "Ê—Ã«· «·ﬂ”«∆Ì : Õ„“… »‰ Õ»Ì» «·“Ì«  ° Ê⁄Ì”Ï »‰ ⁄„— «·Â„–«‰Ì ° Ê„Õ„œ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° Ê€Ì—Â„ „‰ „‘ÌŒ… «·ﬂÊ›ÌÌ‰ €Ì— √‰ „«œ… ﬁ—«¡ Â Ê«⁄ „«œÂ ›Ì «Œ Ì«—Â ⁄‰ Õ„“… ° Êﬁœ –ﬂ—‰« « ’«· ﬁ—«¡ Â ." & vbNewLine
         sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
         sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
         snandan = sanadan & vbNewLine
          
       
        ElseIf index = 1 Then
        ' ‰«›⁄
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "√„« —Ê«Ì… ﬁ«·Ê‰ : ›ÕœÀ‰« »Â« √Õ„œ »‰ ⁄„— »‰ „Õ„œ «·ÃÌ“Ì ° ﬁ«·: ÕœÀ‰« „Õ„œ »‰ √Õ„œ »‰ „‰Ì— ° ﬁ«·: ÕœÀ‰« ⁄»œ «··Â »‰ ⁄Ì”Ï «·„œ‰Ì ° ﬁ«·:ÕœÀ‰« ﬁ«·Ê‰ ⁄‰ ‰«›⁄° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ‘ÌŒÌ √»Ì «·› Õ ›«—” »‰ √Õ„œ »‰ „Ê”Ï »‰ ⁄„—«‰ ° «·„ﬁ—Ì¡ «·÷—Ì— ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—∆ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈»—«ÂÌ„ »‰ ⁄„— «·„ﬁ—∆ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”Ì‰ √Õ„œ »‰ ⁄À„«‰ »‰ Ã⁄›— »‰ »ÊÌ«‰ ° Êﬁ«·:ﬁ—√  ⁄·Ï √»Ì »ﬂ— √Õ„œ »‰ „Õ„œ »‰ «·√‘⁄À Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì ‰‘Ìÿ „Õ„œ »‰ Â«—Ê‰ ° Êﬁ«·: ﬁ—√  ⁄·Ï ﬁ«·Ê‰ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‰«›⁄ ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… Ê—‘ : ›ÕœÀ‰« »Â« √»Ê ⁄»œ «··Â √Õ„œ »‰ „Õ›ÊŸ «·ﬁ«÷Ì »„’— ° ﬁ«·: ÕœÀ‰« √Õ„œ »‰ ≈»—«ÂÌ„ »‰ Ã«„⁄ ° ﬁ«· : ÕœÀ‰« √»Ê „Õ„œ »ﬂ— »‰ ”Â· ° ﬁ«·: ÕœÀ‰« √»Ê „Õ„œ ⁄»œ «·’„œ »‰ ⁄»œ «·—Õ„‰ ° ﬁ«· : ÕœÀ‰« Ê—‘ ⁄‰ ‰«›⁄ ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ‘ÌŒÌ √»Ì «·ﬁ«”„ Œ·› »‰ ≈»—«ÂÌ„ »‰ „Õ„œ »‰ Œ«ﬁ«‰ «·„ﬁ—Ì¡ »„’— ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ⁄·Ï √»Ì Ã⁄›— √Õ„œ »‰ √”«„… «· ÃÌ»Ì ° Êﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·‰Õ«” ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì Ì⁄ﬁÊ» ÌÊ”› »‰ ⁄„—Ê »‰ Ì”«— «·√“—ﬁ ° Êﬁ«· :ﬁ—√  ⁄·Ï Ê—‘ Êﬁ«· : ﬁ—√  ⁄·Ï ‰«›⁄ ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· ‰«›⁄ «·–Ì‰ ”„«Â„ Œ„”… : √»Ê Ã⁄›— Ì“ Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—∆ ° Ê√»Ê œ«Êœ ⁄»œ «·—Õ„‰ »‰ Â—„“ «·√⁄—Ã ° Ê‘Ì»… »‰ ‰’«Õ «·ﬁ«÷Ì ° Ê√»Ê ⁄»œ «··Â „”·„ »‰ Ã‰œ» «·Â–·Ì «·ﬁ«’ ° Ê√»Ê —ÊÕ Ì“Ìœ »‰ —Ê„«‰ ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄‰ √»Ì Â—Ì—… ° Ê«»‰ ⁄»«” ° Ê⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° ⁄‰ √»Ì »‰ ﬂ⁄» ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 2 Then
        ' «»‰ ﬂÀÌ—
        sanadan = "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—  " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… «·»“Ì : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ «·ﬂ« » ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ „Ê”Ï ° ﬁ«·: ÕœÀ‰« „÷— »‰ „Õ„œ «·÷»Ì ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ √»Ì »“… ° ﬁ«·: ﬁ—√  ⁄·Ï ⁄ﬂ—„… »‰ ”·Ì„«‰ »‰ ⁄«„— ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·ﬁ”ÿ ° Êﬁ«· : ﬁ—√  ⁄·Ï «»‰ ﬂÀÌ— ‰›”Â ° ﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·ﬁ«”„ ⁄»œ «·⁄“Ì“ »‰ Ã⁄›— »‰ „Õ„œ «·„ﬁ—Ì¡ «·›«—”Ì ° Êﬁ«· ·Ì: ﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ «·Õ”‰ «·‰ﬁ«‘ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì —»Ì⁄… „Õ„œ »‰ ≈”Õ«ﬁ «·— »⁄Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï «·»“Ì ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… ﬁ‰»· : ›ÕœÀ‰« »Â« √»Ê „”·„ „Õ„œ »‰ √Õ„œ «·»€œ«œÌ ° ﬁ«·: ÕœÀ‰« «»‰ „Ã«Âœ ° ﬁ«·: ﬁ—√  ⁄·Ï ﬁ‰»· ° Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì «·Õ”‰ √Õ„œ »‰ ⁄Ê‰ «·ﬁÊ«” Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì «·«Œ— Ìÿ ÊÂ» »‰ Ê«÷Õ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·ﬁ”ÿ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‘»· »‰ ⁄»«œ Ê „⁄—Ê› »‰ „‘ﬂ«‰ ° Êﬁ«·« ﬁ—√‰« ⁄·Ï «»‰ ﬂÀ‹Ì‹— ° Ê ﬁ«· √»‹‹‹‹Ê ⁄‹‹„‹‹—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ «·Õ„’Ì «·„ﬁ—Ì¡ «·÷—Ì— Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·»€œ«œÌ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï «»‰ „Ã«Âœ Êﬁ«·: ﬁ—√  ⁄·Ï ﬁ‰»· ." & vbNewLine
        sanadan = sanadan & " Ê—Ã‹‹«· «»‰ ﬂÀÌ— «·‹–Ì‹‰ ”„«Â„ À·«À… : ⁄»œ «··Â »‰ «·”«∆» «·„Œ“Ê„Ì ’«Õ» —”Ê· «··Â  Ê„Ã«Âœ »‰ Ã»— √»Ê «·ÕÃ«Ã „Ê·Ï ﬁÌ” »‰ «·”«∆» ° Êœ—»«” „Ê·Ï «»‰ ⁄»«” . Ê√Œ– ⁄»œ «··Â ⁄‰ √»Ì »‰ ﬂ⁄» ‰›”Â. Ê√Œ– „Ã«Âœ Êœ—»«”° ⁄‰ «»‰ ⁄»«”° ⁄‰ √»Ì ° Ê“Ìœ »‰ À«»  ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  °⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 3 Then
        ' √»Ê ⁄„—Ê
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… √»Ì ⁄„— «·œÊ—Ì : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ ⁄·Ì ° ﬁ«·: √Œ»—‰« √»Ê ⁄Ì”Ï „Õ„œ »‰ √Õ„œ »‰ ﬁÿ‰ ”‰… À„«‰ ⁄‘—… ÊÀ·«À„«∆…° ﬁ«·: √Œ»—‰« √»Ê Œ·«œ ”·Ì„«‰ »‰ Œ·«œ ﬁ«·:ÕœÀ‰« «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê ° ﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â „‰ ÿ—Ìﬁ √»Ì ⁄„— «·œÊ—Ì ⁄·Ï ‘ÌŒ‰« ⁄»œ «·⁄“ Ì“ »‰ Ã⁄›— »‰ „Õ„œ »‰ ≈”Õ«ﬁ «·»€œ«œÌ «·›«—”Ì «·„ﬁ—Ì¡ ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì ÿ«Â— ⁄»œ «·Ê«Õœ »‰ ⁄„— »‰ √»Ì Â«‘„ «·„ﬁ—Ì¡ ° „« ·« √Õ’ÌÂ ﬂÀ—… ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì »ﬂ— »‰ „Ã«Âœ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·“⁄—«¡ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” Êﬁ«· :ﬁ—√  ⁄·Ï √»Ì ⁄„— ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï «·Ì“ÌœÌ ° Êﬁ«· ﬁ—√  »Â« ⁄·Ï : √»Ì ⁄„—Ê. " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… √»Ì ‘⁄Ì» «·”Ê”Ì : ›ÕœÀ‰« »Â« Œ·› »‰ ≈»—«ÂÌ„ »‰ „Õ„œ «·„ﬁ—∆ ° ﬁ«· : ÕœÀ‰« √»Ê „Õ„œ «·Õ”‰ »‰ —‘Ìﬁ «·„⁄œ· ° ﬁ«· : ÕœÀ‰« √»Ê ⁄»œ «·—Õ„‰ √Õ„œ »‰ ‘⁄Ì» «·‰”«∆Ì ° ﬁ«· : √Œ»—‰« √»Ê ‘⁄Ì» ° ﬁ«· : √Œ»—‰« «·Ì“ÌœÌ ° ⁄‰ √»Ì ⁄„—Ê ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â »≈ŸÂ«— «·√Ê· „‰ «·„À·Ì‰ Ê«·„ ﬁ«—»Ì‰ Ê»≈œ€«„Â ⁄·Ï ›«—” »‰ √Õ„œ «·„ﬁ—Ì¡ ° Êﬁ«· ·Ì : ﬁ—√  »Â« ﬂ–·ﬂ ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Êﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ﬂ·Â ﬂ–·ﬂ ⁄·Ï √»Ì ⁄„—«‰ „Ê”Ï »‰ Ã—Ì— «·‰ÕÊÌ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ‘⁄Ì» ° Êﬁ«· : ﬁ—√  ⁄·Ï «·Ì“ÌœÌ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì ⁄„—Ê" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê: ÊÕœÀ‰« »√’Ê· «·≈œ€«„ „Õ„œ »‰ √Õ„œ ⁄‰ «»‰ „Ã«Âœ ⁄‰ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” ⁄‰ «·œÊ—Ì ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ï ⁄„—Ê° ÊÕœÀ‰« »Â« √Ì÷« √»Ê «·Õ”‰ ‘ÌŒ‰« ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ «·„»«—ﬂ ⁄‰ Ã⁄›— »‰ ”·Ì„«‰ ⁄‰ √»Ì ‘⁄Ì» ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê . " & vbNewLine
        sanadan = sanadan & "Ê—Ã«· √»Ì ⁄„—Ê : Ã„«⁄… „‰ √Â· «·ÕÃ«“ Ê„‰ √Â· «·»’—… ° ›„‰ √Â· „ﬂ… : „Ã«Âœ ° Ê”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„… »‰ Œ«·œ ° Ê⁄ÿ«¡ »‰ √»Ì —»«Õ ° Ê⁄»œ «··Â »‰ ﬂÀÌ— ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ „ÕÌ’‰ ° ÊÕ„Ìœ »‰ ﬁÌ” «·√⁄—Ã «·ﬁ«—∆ ° Ê„‰ √Â· «·„œÌ‰… : Ì“Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—Ì¡ ÊÌ“Ìœ »‰ —Ê„«‰ ° Ê‘Ì»… »‰ ‰’«Õ ° Ê„‰ √Â· «·»’—… : «·Õ”‰ »‰ √»Ì «·Õ”‰ «·»’—Ì ° ÊÌÕÌ »‰ Ì⁄„— ° Ê€Ì—Â„« ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄„‰  ﬁœ„ „‰ «·’Õ«»… Ê€Ì—Â„ . " & vbNewLine
        sanadan = sanadan & "ﬁ·  : Ê√Œ– ”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„…° ÊÌÕÌÏ »‰ Ì⁄„— ° ⁄‰ «»‰ ⁄»«” Ê√Œ– «»‰ ⁄»«” ⁄‰ √»Ì »‰ ﬂ⁄» Ê“Ìœ »‰ À«»  ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -."
        
        ElseIf index = 4 Then
        '«»‰ ⁄«„—
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… Â‘«„ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ﬁ«·: ÕœÀ‰« «»‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« «·Õ”Ì‰ »‰ „Â—«‰ «·Ã„«· ° ﬁ«· :ÕœÀ‰« √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì ° ﬁ«· : ÕœÀ‰« Â‘«„ »‰ ⁄„«— ° ﬁ«·: ÕœÀ‰« ⁄—«ﬂ »‰ Œ«·œ «·„—Ì ° ﬁ«· :ﬁ—√  ⁄·Ï ÌÕÌÌ »‰ «·Õ«—À «·–„«—Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï ⁄»œ «··Â »‰ ⁄«„— ° ﬁ«· : √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ ‘ÌŒ‰« ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Ê ﬁ«· : ﬁ—√  »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ⁄»œ«‰ ° Êﬁ«· : ﬁ—√  ⁄·Ï «·Õ·Ê«‰Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï Â‘«„ " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… «»‰ –ﬂÊ«‰ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ „Ê”Ï »‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« √Õ„œ »‰ ÌÊ”› «· €·»Ì ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ –ﬂÊ«‰ ° ﬁ«· : ÕœÀ‰« √ÌÊ» »‰  „Ì„ «· „Ì„Ì ° ﬁ«· :ÕœÀ‰« ÌÕÌÌ »‰ «·Õ«—À «·–„«—Ì ° ﬁ«· : ﬁ—√  ⁄·Ï «»‰ ⁄«„— ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ⁄»œ «·⁄“Ì“ »‰ Ã⁄›— «·›«—”Ì «·„ﬁ—Ì¡ Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ï »ﬂ— „Õ„œ »‰ «·Õ”‰ «·‰ﬁ«‘ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ⁄»œ «··Â Â«—Ê‰ »‰ „Ê”Ï »‰ ‘—Ìﬂ «·√Œ›‘ Ê—Ê«Â« «·√Œ›‘ ⁄‰ ⁄»œ «··Â »‰ –ﬂÊ«‰ " & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· «»‰ ⁄«„— «·‹–Ì‹‰ ”‹‹„«Â„ : √»Ê «·œ—œ«¡ ⁄ÊÌ„— »‰ ⁄«„— ’«Õ» —”Ê· «··Â ° Ê«·„€Ì—… »‰ √»Ì ‘Â«» «·„Œ“Ê„Ì ° Ê√Œ‹– √»Ê «·œ—œ«¡ ⁄‹‹‰ «·‰»Ì . Ê√Œ– «·„€Ì—… ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -" & vbNewLine
        
        ElseIf index = 5 Then
        '⁄«’„
        sanadan = "ﬁ«· √»‹‹Ê ⁄‹„‹—Ê «·‹œ«‰‹‹Ì ›‹‹‹Ì «·‹ Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… √»Ì »ﬂ— ‘⁄»…: ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ ⁄·Ì «·ﬂ« » ﬁ«·: ÕœÀ‰« »‰ „Ã«Âœ ﬁ«·: ÕœÀ‰« ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ⁄„— «·ÊﬂÌ⁄Ì ° ﬁ«·:ÕœÀ‰« √»Ì ﬁ«·:ÕœÀ‰« ÌÕÌÌ »‰ √œ„ ° ﬁ«·: ÕœÀ‰« √»Ê »ﬂ— ⁄‰ ⁄«’„ ° ﬁ«· √»Ê ⁄„—Ê: Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ «·„ﬁ—Ì¡ ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—Ì¡ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈»—«ÂÌ„ »‰ ⁄»œ «·—Õ„‰ »‰ √Õ„œ «·„ﬁ—Ì¡ «·»€œ«œÌ Êﬁ«·: ﬁ—√  ⁄·Ï ÌÊ”› »‰ Ì⁄ﬁÊ» «·Ê«”ÿÌ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‘⁄Ì» »‰ √ÌÊ» «·’—Ì›Ì‰Ì ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ÌÕÌÌ »‰ √œ„ ⁄‰ √»Ï »ﬂ— ⁄‰ ⁄«’„." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… Õ›’ : ›ÕœÀ‰« »Â« √»Ê «·Õ”‰ ÿ«Â‹— »‰ €·»Ê‰ «·„ﬁ—∆ ° ﬁ«· : ÕœÀ‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ’«·Õ «·Â«‘„Ì «·÷—Ì— «·„ﬁ—∆ »«·»’—… ° ﬁ«·: ÕœÀ‰« √»Ê «·⁄»«” √Õ„œ »‰ ”Â· «·√‘‰«‰Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì „Õ„œ ⁄»Ìœ »‰ «·’»«Õ ° Êﬁ«·: ﬁ—√  ⁄·Ï Õ›’ ° Êﬁ«·: ﬁ—√  ⁄·Ï ⁄«’‹„ ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ‘ÌŒ‰« √»Ì «·Õ”‰ Êﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï «·Â«‘„Ì Êﬁ«·: ﬁ—√  ⁄·Ï «·√‘‰«‰Ì ⁄‰ ⁄»Ìœ ⁄‰ Õ›’ ⁄‰ ⁄«’‹„ . " & vbNewLine
        sanadan = sanadan & "Ê—Ã«· ⁄«’„ «·‹–Ì‹‰ ”„«Â„ «À‰«‰ : √»Ê ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ê „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·»  ° Ê√»Ì »‰ ﬂ⁄»  ° Ê“Ìœ »‰ À«»   ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï - ° √Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰  ° Ê«»‰ „”⁄Êœ  ° ⁄‰ —”Ê· «··Â - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 6 Then
        'Õ„“…
        sanadan = "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… Œ·› : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« «»‰ „Ã«Âœ ° ÕœÀ‰« ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ ° ﬁ«· : ÕœÀ‰« Œ·› ° ﬁ«·: ⁄‰ ”·Ì„ ⁄‰ Õ„“… ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·Õ”‰ ‘ÌŒ‰« ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ „Õ„œ »‰ ÌÊ”› »‰ ‰Â«— «·Õ— ﬂÌ »«·»’—… ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”Ì‰ √Õ„œ »‰ ⁄À„«‰ »‰ Ã⁄›— »‰ »ÊÌ«‰ ° Êﬁ«· ·Ì :ﬁ—√  ⁄·Ï ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ ﬁ»· √‰ Ìﬁ—Ì¡ »«Œ Ì«— Œ·› ° Êﬁ«· ·Ì : ﬁ—√  ⁄·Ï Œ·› ° Êﬁ«· : ﬁ—√  ⁄·Ï ”·Ì„ ° Ê ﬁ«· : ﬁ—√  ⁄·Ï Õ„“… ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… Œ·«œ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« √Õ„œ »‰ „Ê”Ï ° ﬁ«· : ÕœÀ‰« ÌÕÌÏ »‰ √Õ„œ »‰ Â«—Ê‰ «·„“Êﬁ ° ⁄‰ √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì ° ⁄‰ Œ·«œ ° ⁄‰ ”·Ì„ ° ⁄‰ Õ„“… ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ «·÷—Ì— ‘ÌŒ‰« ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ‘‰»Ê– ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ‘«–«‰ «·ÃÊÂ—Ì «·„ﬁ—Ì ° Êﬁ«· :ﬁ—√  ⁄·Ï Œ·«œ Êﬁ«· : ﬁ—√  ⁄·Ï ”·Ì„ ° Êﬁ—√ ”·Ì„ ⁄·Ï Õ„“…." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
        sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 7 Then
        '«·ﬂ”«∆Ï
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… √»Ì ⁄„—Ê «·œÊ—Ì : ›ÕœÀ‰« »Â« √»Ê „Õ„œ ⁄»œ «·—Õ„‰ »‰ ⁄„— »‰ „Õ„œ «·„⁄œ· ° ﬁ«· : ÕœÀ‰« √»Ê ⁄„— ⁄»œ «··Â »‰ √Õ„œ »‰ œÌ“ÊÌÂ «·œ„‘ﬁÌ ° ﬁ«· : ÕœÀ‰« Ã⁄›— »‰ „Õ„œ »‰ √”œ «·‰’Ì»Ì ° ﬁ«· : ÕœÀ‰« √»Ê ⁄„— «·œÊ—Ì ° ⁄‰ «·ﬂ”«∆Ì ° Ê ﬁ«· √»Ê ⁄‹„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄·Ì »‰ «·Ã·‰œÌ «·„Ê’·Ì ° Ê ﬁ«· :ﬁ—√  ⁄·Ï Ã⁄›— »‰ „Õ„œ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì ⁄„— «·œÊ—Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï «·ﬂ”«∆Ì ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… √»Ì «·Õ«—À : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« »Â« «»‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« „Õ„œ »‰ ÌÕÌÏ ( «·ﬂ”«∆Ì «·’€Ì—) ° ⁄‰ √»Ì «·Õ«—À ° ⁄‰ «·ﬂ”«∆Ì ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—Ì¡ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·ﬁ«”„ “Ìœ »‰ ⁄·Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï √Õ„œ »‰ «·Õ”‰ «·„⁄—Ê› »«·»ÿÌ ° Êﬁ«· :ﬁ—√  ⁄·Ï „Õ„œ »‰ ÌÕÌÏ ( «·ﬂ”«∆Ì «·’€Ì—) ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì «·Õ«—À ° Êﬁ«· : ﬁ—√  ⁄·Ï «·ﬂ”«∆Ì ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· «·ﬂ”«∆Ì : Õ„“… »‰ Õ»Ì» «·“Ì«  ° Ê⁄Ì”Ï »‰ ⁄„— «·Â„–«‰Ì ° Ê„Õ„œ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° Ê€Ì—Â„ „‰ „‘ÌŒ… «·ﬂÊ›ÌÌ‰ €Ì— √‰ „«œ… ﬁ—«¡ Â Ê«⁄ „«œÂ ›Ì «Œ Ì«—Â ⁄‰ Õ„“… ° Êﬁœ –ﬂ—‰« « ’«· ﬁ—«¡ Â ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
        sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 8 Then
        '√»Ê Ã⁄›—
        sanadan = "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… «»‰ Ê—œ«‰ : ›ÕœÀ‰« »Â« «·‘ÌŒ √»Ê Õ›’ ⁄„— »‰ «·Õ”‰ »‰ „“Ìœ «·„—«€Ì »ﬁ—«¡ Ì ⁄·ÌÂ ﬁ«· : √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ √Õ„œ »‰ ⁄»œ «·Ê«Õœ «·”⁄œÌ „‘«›Â… ⁄‰ «·≈„«„ √»Ì «·Ì„‰ “Ìœ »‰ «·Õ”‰ «··€ÊÌ ° ﬁ«· : √Œ»—‰« √»Ê „Õ„œ ⁄»œ «··Â »‰ ⁄·Ì «·»€œ«œÌ √Œ»—‰« «·‘—Ì› √»Ê «·›÷· ⁄»œ «·ﬁ«Â— »‰ ⁄»œ «·”·«„ «·⁄»«”Ì ° √Œ»—‰« √»Ê ⁄»œ «··Â „Õ„œ »‰ «·Õ”Ì‰ «·ﬂ«—“Ì‰Ì ° √Œ»—‰« √»Ê «·›—Ã „Õ„œ »‰ √Õ„œ »‰ ≈»—«ÂÌ„ «·‘ÿÊÌ ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ √Õ„œ »‰ Â«—Ê‰ «·—«“Ì ° √Œ»—‰« √»Ê «·⁄»«” «·›÷· »‰ ‘«–«‰ »‰ ⁄Ì”Ï «·—«“Ì √Œ»—‰« √»Ê «·Õ”‰ √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì °√Œ»—‰« ⁄Ì”Ï »‰ „Ì‰« ﬁ«·Ê‰ ° √Œ»—‰« ⁄Ì”Ï »‰ Ê—œ«‰." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â „Õ„œ ⁄»œ «·—Õ„‰ »‰ ⁄·Ì «·‰ÕÊÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï «·≈„‹‹«„ √»Ì ⁄»œ „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„’—Ì ° ﬁ«· : ﬁ—√  »Â« «·ﬁ—¬‰ ⁄·Ï «·ﬂ„«· ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ›«—” «· „Ì„Ì ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·Ì„‰ «·ﬂ‰œÌ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «·≈„«„ √»Ì „‰’Ê— „Õ„œ »‰ ⁄»œ «·„·ﬂ »‰ «·Õ”‰ »‰ ŒÌ—Ê‰ «·»€œ«œÌ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·ﬁ«”„ ⁄»œ «·”Ìœ »‰ ⁄ «» «·„ﬁ—Ì¡ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ÿ«Â— „Õ„œ »‰ Ì«”Ì‰ «·Õ·»Ì ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·›—Ã «·‘ÿÊÌ ﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì »ﬂ— »‰ Â«—Ê‰ ° ﬁ«·: ﬁ—√  »Â« ⁄·Ï «·›÷· »‰ ‘«–«‰ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «·Õ·Ê«‰Ì ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï ﬁ«·Ê‰ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «»‰ Ê—œ«‰ . " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… «»‰ Ã„«“ : ›ÕœÀ‰« »Â« √»Ê ≈”Õ«ﬁ ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ≈»—«ÂÌ„ »‰ Õ« „ «·Ã–«„Ì »ﬁ—«¡ Ì ⁄·ÌÂ ⁄‰ √»Ì Õ›’ ⁄„— »‰ €‹œÌ— »‰ «·ﬁÊ«” «·œ„‘ﬁÌ ° √‰»√‰« √»Ê «·Ì„‰ »‰ «·Õ”‰ «·»€œ«œÌ ° √Œ»—‰« √»Ê „Õ„œ ”»ÿ «·ŒÌ«ÿ ° √Œ»—‰« «·√” «– √»Ê «·⁄“ „Õ„œ »‰ «·Õ”Ì‰ »‰ »‰œ«— «·Ê«”ÿÌ ° √Œ»—‰« «·≈„«„ √»Ê «·ﬁ«”„ ÌÊ”› »‰ Ã»«—… «·Â–·Ì ° √Œ»—‰« √»Ê ‰’— „‰’Ê— »‰ „Õ„œ «·ﬁÂ‰œ“Ì ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ «·Œ»«“Ì ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ «·›÷· «·ÃÊÂ—Ì ° √Œ»—‰« „Õ„œ »‰ √Õ„œ »‰ «·Õ”‰ «·Àﬁ›Ì «·ﬂ”«∆Ì ° √Œ»—‰« „Õ„œ »‰ ⁄»œ «··Â »‰ ‘«ﬂ— «·’Ì—›Ì ° √Œ»—‰« √»Ê «·⁄»«” √Õ„œ »‰ ”Â· «·ÿÌ«‰ ° √Œ»—‰« √»Ê ⁄„—«‰ „Ê”Ï »‰ ⁄»œ «·—Õ„‰ «·»“«“ ° √Œ»—‰« „Õ„œ »‰ ⁄Ì”Ï »‰ ≈»—«ÂÌ„ »‰ —“Ì‰ «·√’»Â«‰Ì ° √Œ»—‰« ”·Ì„«‰ »‰ œ«Êœ »‰ ⁄·Ì »‰ ⁄»œ «··Â »‰ ⁄»«” «·Â«‘„Ì ° √Œ»—‰« ≈”„«⁄Ì· »‰ Ã⁄›— »‰ √»Ì ﬂÀÌ— «·„œ‰Ì ° √Œ»—‰« ”·Ì„«‰ »‰ „”·„ «»‰ Ã„«“." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì „Õ„œ »‰ ⁄»œ «·—Õ„‰ «·Õ‰›Ì ° Êﬁ—√ »Â« «·ﬁ—«‰ ﬂ·Â ⁄·Ï „Õ„œ »‰ √Õ„œ «·’«∆€ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ≈”Õ«ﬁ »‰ ›«—” ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Ì„‰ ° Êﬁ—√ »Â« ⁄·Ï ”»ÿ «·ŒÌ«ÿ ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì ÿ«Â— √Õ„œ »‰ ⁄·Ì »‰ ⁄»Ìœ «··Â »‰ ”Ê«— ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄·Ì «·Õ”‰ »‰ «·›÷· «·‘—„ﬁ«‰Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄»œ «··Â »‰ «·„“—»«‰ «·√’»Â«‰Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄„— „Õ„œ »‰ √Õ„œ »‰ ⁄„— «·Œ—ﬁÌ ° Êﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ Ã⁄›— »‰ „Õ„Êœ «·√‘‰«‰Ì ° Êﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ «·Àﬁ›Ì «·ﬂ”«∆Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ ‘«ﬂ— ° Êﬁ—√ »Â« ⁄·Ï «»‰ ”Â· «·ÿÌ«‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄„—«‰ «·»“«“ ° Êﬁ—√ »Â« ⁄·Ï «»‰ —“Ì‰ ° Êﬁ—√ »Â« ⁄·Ï «·Â«‘„Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ Ã⁄›— ° Êﬁ—√ »Â« ⁄·Ï «»‰ Ã„«“ ° Êﬁ—√ «»‰ Ã„«“ ° Ê«»‰ Ê—œ«‰ ° ⁄·Ï √»Ì Ã⁄›— ." & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· √»Ì Ã⁄›— À·«À… : „Ê·«Â ⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° Ê√»Ê Â—Ì—… ° Ê«»‰ ⁄»«” . Êﬁ—√ Âƒ·«¡ «·À·«À… ⁄·Ï √»Ì »‰ ﬂ⁄» ° Êﬁ—√ √»Ê Â—Ì—… ° Ê«»‰ ⁄»«” ° √Ì÷« ⁄·Ï “Ìœ »‰ À«»  . Ê√Œ– “Ìœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ -° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 9 Then
        'Ì⁄ﬁÊ»
        sanadan = "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… —ÊÌ” : ›ÕœÀ‰« »Â« «·‘ÌŒ «·≈„«„ √»Ê «·⁄»«” √Õ„œ »‰ „Õ„œ »‰ «·Œ÷— «·Õ‰›Ì »ﬁ—«¡ Ì ⁄·ÌÂ ﬁ«·: √Œ»—‰« : √»Ê «·⁄»«” √Õ„œ »‰ √»Ì ÿ«·» »‰ √»Ì «·‰⁄„ «·’«·ÕÌ ﬁ—«¡… ⁄·ÌÂ ° √Œ»—‰« √»Ê ÿ«·» ⁄»œ «··ÿÌ› »‰ „Õ„œ »‰ «·ﬁ»ÌÿÌ ° ›Ì ﬂ «»Â √Œ»—‰« »Â« √»Ê »ﬂ— √Õ„œ »‰ «·„ﬁ—» «·ﬂ—ŒÌ ﬁ—«¡… ⁄·ÌÂ ° √Œ»—‰« √»Ê ÿ«Â— √Õ„œ »‰ ⁄·Ì «·„ﬁ—Ì¡ «·√” «– √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ⁄·Ì «·ŒÌ«ÿ ° √Œ»—‰« «·√” «– «·≈„«„ √»Ê «·Õ”‰ ⁄·Ì »‰ √Õ„œ »‰ ⁄„— «·Õ„«„Ì ° √Œ»—‰« √»Ê «·ﬁ«”„ ⁄»œ «··Â »‰ «·Õ”‰ »‰ ”·Ì„«‰ «·‰Œ«” ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ Â«—Ê‰ »‰ ‰«›⁄ «· „«— «·»€œ«œÌ ° √Œ»—‰« √»Ê ⁄»œ «··Â „Õ„œ »‰ «·„ Êﬂ· «·„⁄—Ê› »—ÊÌ” ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì „Õ„œ ⁄»œ «·—Õ„‰ »‰ √Õ„œ »‰ ⁄·Ì «·»€œ«œÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï «·≈„«„ «· ﬁÌ „Õ„œ »‰ √Õ„œ «·„’—Ì ° Êﬁ—√ »Â« ⁄·Ï ≈»—«ÂÌ„ »‰ √Õ„œ «·≈”ﬂ‰œ—Ì ° Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï ⁄»œ «··Â »‰ ⁄·Ì «·»€œ«œÌ ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì «·⁄“ «·ﬁ·«‰”Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄·Ì «·Õ”‰ »‰ «·ﬁ«”„ «·Ê«”ÿÌ ° Êﬁ—√ »Â« ⁄·Ï : «·Õ„«„Ì ° Êﬁ—√ »Â« ⁄·Ï «·‰Œ« ” ° Êﬁ—√ »Â« ⁄·Ï «· „«— ° Êﬁ—√ ⁄·Ï —ÊÌ” ° Êﬁ—√ »Â« ⁄·Ï Ì⁄ﬁÊ» . " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… —ÛÊÕ : ›ÕœÀ‰« »Â« «·‘ÌŒ √»Ê «·⁄»«” √Õ„œ »‰ „Õ„œ »‰ «·Õ”Ì‰ «·‘Ì—«“Ì »ﬁ—«¡ Ì ⁄·ÌÂ ⁄‰ «·≈„«„ √»Ì «·Õ”‰ ⁄·Ì »‰ √Õ„œ «·„ﬁœ”Ì ° √Œ»—‰« √»Ê «·Ì„‰ «·ﬂ‰œÌ ‘›«Â« ° √Œ»—‰« √»Ê „Õ„œ «·»€œ«œÌ ° √Œ»—‰« √»Ê «·›÷· «·‘—Ì› «·„ﬂÌ ° √Œ»—‰« „Õ„œ »‰ «·Õ”Ì‰ «·›«—”Ì ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ≈»—«ÂÌ„ »‰ Œ‘‰«„ «·„«·ﬂÌ «·»’—Ì √Œ»—‰« √»Ê «·⁄»«” „Õ„œ »‰ Ì⁄ﬁÊ» »‰ «·ÕÃ«Ã »‰ „⁄«ÊÌ… «· Ì„Ì ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ ÊÂ» »‰ ÌÕÌÏ »‰ «·⁄·«¡ «·Àﬁ›Ì «·ﬁ“«“ ° √Œ»—‰« —ÊÕ »‰ ⁄»œ «·„ƒ„‰ «·»’—Ì ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï „Õ„œ »‰ √Õ„œ »«·ﬁ«Â—… «·„Õ—Ê”… ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â «·’«∆€ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ≈”Õ«ﬁ «·œ„‘ﬁÌ Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï ⁄»œ «··Â »‰ ⁄·Ì ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì ÿ«Â— »‰ ”Ê«— ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·ﬁ«”„ «·„”«›— »‰ «·ÿÌ» »‰ ⁄»«œ «·»’—Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ Œ‘‰«„ ° Êﬁ—√ »Â« ⁄·Ï «»‰ ⁄»« ” «· Ì„Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ ÊÂ» ° Êﬁ—√ »Â« ⁄·Ï —ÊÕ ° Êﬁ—√ »Â« ⁄·Ï Ì⁄ﬁÊ» ." & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· Ì⁄ﬁÊ» «·–Ì‰ ”„«Â„ √—»⁄… : √»Ê «·„‰–— ”·«„ »‰ ”·Ì„«‰ «·ÿÊÌ· ° Ê‘Â«» »‰ ‘—‰›… ° Ê„ÂœÌ »‰ „Ì„Ê‰ ° Ê√»Ê «·√‘Â» Ã⁄›— »‰ ÕÌ«‰ «·⁄ÿ«—œÌ .ÊﬁÌ· ≈‰ Ì⁄ﬁÊ» ﬁ—√ ⁄·Ï √»Ì ⁄„—Ê »‰ «·⁄·«¡ Êﬁ—√ ”·«„ ⁄·Ï ⁄«’„ Ê√»Ì ⁄„—Ê ° Êﬁ‹‹‹—√ ‘Â«» «·ÃÕœ—Ì Êﬁ—√ ⁄«’„ ⁄·Ï «·Õ”‰ «·»’—Ì Ê⁄·Ï ”·Ì„«‰ »‰ ﬁ … Êﬁ—√ ”·Ì„«‰ ⁄·Ï «»‹‰ ⁄»« ” Êﬁ—√ „ÂœÌ ⁄·Ï ‘⁄Ì» »‰ «·Õ»Õ«» Êﬁ—√ ⁄·Ï √»Ì «·⁄«·Ì… «·—Ì«ÕÌ Êﬁ—√ ⁄·Ï √»Ì Ê“Ìœ Êﬁ—√ √»Ê «·√‘Â» ⁄·Ï √»Ì —Ã«¡ ⁄„—«‰ »‰ „·Õ«‰ «·⁄ÿ«—œÌ Êﬁ—√ ⁄·Ï √»Ì „Ê”‹‹‹Ï «·√‘⁄—Ì Êﬁ—√ ⁄·Ï —”Ê· «··Â ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 10 Then
        'Œ·›
        sanadan = "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "√„« —Ê«Ì… ≈œ—Ì” «·Ê—«ﬁ : ›ÕœÀ‰« »Â« √»Ê Õ›’ ⁄„— »‰ «·Õ”‰ »ﬁ—«¡ Ì ⁄·ÌÂ Ÿ«Â— œ„‘ﬁ ° ⁄‰ ‘ÌŒÂ «·≈„«„ «·ŒÿÌ» √»Ì «·⁄»«” √Õ„œ »‰ ≈»—«ÂÌ„ »‰ ⁄„— «·›«—Ê∆Ì «·‘«›⁄Ì ° ﬁ«· : √Œ»—‰« Ê«·œÌ ° ﬁ«· : √Œ»—‰« √»Ê «·”⁄«œ«  «·√”⁄œ »‰ ”·ÿ«‰ «·Ê«”ÿÌ ° √Œ»—‰« √»Ê «·⁄“ „Õ„œ »‰ «·Õ”Ì‰ «·Ê«”ÿÌ ° √Œ»—‰« √»Ê «·Õ”Ì‰ √Õ„œ »‰ ⁄»œ «··Â »‰ «·Œ÷— «·”Ê”‰Ã—œÌ ° √Œ»—‰« √»Ê «·Õ”‰ „Õ„œ »‰ ⁄»œ «··Â »‰ „Õ„œ »‰ „—… «·ÿÊ”Ì «·„⁄—Ê› »«»‰ √»Ì ⁄„— «·‰ﬁ«‘ ° √Œ»—‰« √»Ê Ì⁄ﬁÊ» ≈”Õ«ﬁ »‰ ≈»—«ÂÌ„ «·Ê—«ﬁ ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ﬂ· „‰ «·‘ÌŒÌ‰ √»Ì ⁄»œ «··Â «·Õ‰›Ì ° Ê√»Ì „Õ„œ «·‘«›⁄Ì «·„’—ÌÌ‰ ° Êﬁ—√ ﬂ· „‰Â„« ⁄·Ï √»Ì ⁄»œ «··Â „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„’—Ì ° Êﬁ—√ »Â« ⁄·Ï «·ﬂ„«· »‰ ›«—” ° Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·ﬁ«”„ Â»… «··Â »‰ √Õ„œ »‰ «·ÿ»— «·»€œ«œÌ ° Êﬁ—√ »Â« ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄·Ì »‰ „Ê”Ï «·ŒÌ«ÿ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Õ”Ì‰ «·”Ê”‰Ã—œÌ ° Êﬁ—√ »Â« ⁄·Ï «»‰ √»Ì ⁄„— «·ÿÊ”Ì ° Êﬁ—√ »Â« ⁄·Ï ≈”Õ«ﬁ «·Ê—«ﬁ ° Êﬁ—√ »Â« ⁄·Ï Œ·› ." & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… ≈œ—Ì” : ›ÕœÀ‰« »Â« √Õ„œ »‰ „Õ„œ »‰ «·Õ”Ì‰ «·›«—”Ì »ﬁ—«¡ Ì ⁄·ÌÂ ° √Œ»—‰« ⁄·Ì »‰ √Õ„œ ›Ì„« ‘«›Â‰Ì »Â °⁄‰ “Ìœ »‰ «·Õ”‰ «·»€œ«œÌ ° √Œ»—‰« √»Ê «·ﬁ«”„ »‰ √Õ„œ «·Õ—Ì—Ì ° √Œ»—‰« √»Ê »ﬂ—„Õ„œ »‰ ⁄»Ì »‰ „Õ„œ «·ŒÌ«ÿ ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ⁄»œ «··Â «·Õ–«¡ ° √Œ»—‰« √»Ê ≈”Õ«ﬁ ≈»—«ÂÌ„ »‰ «·Õ”Ì‰ »‰ ⁄»œ «··Â «·‰”«Ã «·„⁄—Ê› »«·‘ÿÌ ° √Œ»—‰« ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ «·Õœ«œ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·‘ÌŒ √»Ì „Õ„œ ⁄»œ «·—Õ„‰ »‰ √Õ„œ «·Ê«”ÿÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„⁄œ· ° Êﬁ—√ »Â« ⁄·Ï ≈»—«ÂÌ„ »‰ √Õ„œ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Ì„‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì „Õ„œ ”»ÿ «·ŒÌ«ÿ ° ﬁ«· : Êﬁ—√  »Â« «·ﬁ—¬‰ „‰ √Ê·Â ≈·Ï ¬Œ—Â ⁄·Ï «·≈„«„Ì‰ «·‘—Ì› √»Ì «·›÷· ⁄»œ «·ﬁ«Â— »‰ ⁄»œ «·”·«„ «·⁄»«”Ì ° Ê√»Ì «·„⁄«·Ì À«»  »‰ »‰œ«— »‰ ≈»—«ÂÌ„ «·»ﬁ«· ° ›√„« «·‘—Ì› ›√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â „Õ„œ »‰ «·Õ”Ì‰ «·ﬂ«—“Ì‰Ì ° Ê√Œ»—Â √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ √»Ì «·⁄»«” «·Õ”‰ »‰ ”⁄Ìœ »‰ Ã⁄›— «·„ÿÊ⁄Ì ° Ê√„« √»Ê «·„⁄«·Ì ›√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ «·ﬁ«÷Ì √»Ì «·⁄·«¡ „Õ„œ »‰ ⁄·Ì »‰ Ì⁄ﬁÊ» «·Ê«”ÿÌ ° Êﬁ—√ «·Ê«”ÿÌ »Â« „‰ «·ﬂ «» ⁄·Ï «·≈„«„ √»Ì »ﬂ— √Õ„œ »‰ Ã⁄›— »‰ Õ„œ«‰ »‰ „«·ﬂ «·ﬁÿÌ⁄Ì ° Êﬁ—√ «·ﬁÿÌ⁄Ì Ê«·„ÿÊ⁄Ì Ã„Ì⁄« ⁄·Ï ≈œ—Ì” ° Êﬁ—√ ≈œ—Ì” ⁄·Ï Œ·› ° Ê«··Â «·„Ê›ﬁ . " & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Œ·› : Ê—Ã«· Œ·› ”·Ì„ ’«Õ» Õ„“… ° ÊÌ⁄ﬁÊ» »‰ Œ·Ì›… «·√⁄‘Ï ’«Õ» √»Ì »ﬂ— ° Ê√»Ê “Ìœ ”⁄Ìœ ”⁄Ìœ »‰ √Ê” «·√‰’«—Ì ’«Õ» «·„›÷· «·÷»Ì Ê√»«‰ «·⁄ÿ«— ° Êﬁ—√ √»Ê »ﬂ— ° Ê«·„›÷· ° Ê√»«‰ ⁄·Ï ⁄«’„ . Ê—ÊÏ «·ﬁ—«¡… √Ì÷« ⁄‰ «·ﬂ”«∆Ì Ê⁄‰ ÌÕÌÏ »‰ ¬œ„ ⁄‰ √»Ì »ﬂ— ° Ê«··Â «·„Ê›ﬁ . ﬁ·  : Ê√Œ– ⁄«’„ ⁄‰ √»Ì ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ì „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·» ° Ê√»Ì »‰ ﬂ⁄» ° Ê“Ìœ »‰ À«»  ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        sanadan = sanadan & "Ê√Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰ ° Ê«»‰ „”⁄Êœ ° ⁄‰ —”Ê· «··Â ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -. Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ . Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 11 Then
        'Ê—‘
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan + "Ê√„« —Ê«Ì… Ê—‘ : ›ÕœÀ‰« »Â« √»Ê ⁄»œ «··Â √Õ„œ »‰ „Õ›ÊŸ «·ﬁ«÷Ì »„’— ° ﬁ«·: ÕœÀ‰« √Õ„œ »‰ ≈»—«ÂÌ„ »‰ Ã«„⁄ ° ﬁ«· : ÕœÀ‰« √»Ê „Õ„œ »ﬂ— »‰ ”Â· ° ﬁ«·: ÕœÀ‰« √»Ê „Õ„œ ⁄»œ «·’„œ »‰ ⁄»œ «·—Õ„‰ ° ﬁ«· : ÕœÀ‰« Ê—‘ ⁄‰ ‰«›⁄ ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ‘ÌŒÌ √»Ì «·ﬁ«”„ Œ·› »‰ ≈»—«ÂÌ„ »‰ „Õ„œ »‰ Œ«ﬁ«‰ «·„ﬁ—Ì¡ »„’— ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ⁄·Ï √»Ì Ã⁄›— √Õ„œ »‰ √”«„… «· ÃÌ»Ì ° Êﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·‰Õ«” ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì Ì⁄ﬁÊ» ÌÊ”› »‰ ⁄„—Ê »‰ Ì”«— «·√“—ﬁ ° Êﬁ«· :ﬁ—√  ⁄·Ï Ê—‘ Êﬁ«· : ﬁ—√  ⁄·Ï ‰«›⁄ ." & vbNewLine
        sanadan = sanadan + "Ê—Ã«· ‰«›⁄ «·–Ì‰ ”„«Â„ Œ„”… : √»Ê Ã⁄›— Ì“ Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—∆ ° Ê√»Ê œ«Êœ ⁄»œ «·—Õ„‰ »‰ Â—„“ «·√⁄—Ã ° Ê‘Ì»… »‰ ‰’«Õ «·ﬁ«÷Ì ° Ê√»Ê ⁄»œ «··Â „”·„ »‰ Ã‰œ» «·Â–·Ì «·ﬁ«’ ° Ê√»Ê —ÊÕ Ì“Ìœ »‰ —Ê„«‰ ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄‰ √»Ì Â—Ì—… ° Ê«»‰ ⁄»«” ° Ê⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° ⁄‰ √»Ì »‰ ﬂ⁄» ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 21 Then
        'ﬁ‰»·
        sanadan = "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—  " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… ﬁ‰»· : ›ÕœÀ‰« »Â« √»Ê „”·„ „Õ„œ »‰ √Õ„œ «·»€œ«œÌ ° ﬁ«·: ÕœÀ‰« «»‰ „Ã«Âœ ° ﬁ«·: ﬁ—√  ⁄·Ï ﬁ‰»· ° Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì «·Õ”‰ √Õ„œ »‰ ⁄Ê‰ «·ﬁÊ«” Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì «·«Œ— Ìÿ ÊÂ» »‰ Ê«÷Õ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·ﬁ”ÿ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‘»· »‰ ⁄»«œ Ê „⁄—Ê› »‰ „‘ﬂ«‰ ° Êﬁ«·« ﬁ—√‰« ⁄·Ï «»‰ ﬂÀ‹Ì‹— ° Ê ﬁ«· √»‹‹‹‹Ê ⁄‹‹„‹‹—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ «·Õ„’Ì «·„ﬁ—Ì¡ «·÷—Ì— Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·»€œ«œÌ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï «»‰ „Ã«Âœ Êﬁ«·: ﬁ—√  ⁄·Ï ﬁ‰»· ." & vbNewLine
        sanadan = sanadan & " Ê—Ã‹‹«· «»‰ ﬂÀÌ— «·‹–Ì‹‰ ”„«Â„ À·«À… : ⁄»œ «··Â »‰ «·”«∆» «·„Œ“Ê„Ì ’«Õ» —”Ê· «··Â  Ê„Ã«Âœ »‰ Ã»— √»Ê «·ÕÃ«Ã „Ê·Ï ﬁÌ” »‰ «·”«∆» ° Êœ—»«” „Ê·Ï «»‰ ⁄»«” . Ê√Œ– ⁄»œ «··Â ⁄‰ √»Ì »‰ ﬂ⁄» ‰›”Â. Ê√Œ– „Ã«Âœ Êœ—»«”° ⁄‰ «»‰ ⁄»«”° ⁄‰ √»Ì ° Ê“Ìœ »‰ À«»  ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  °⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -."
        
        ElseIf index = 31 Then
        '«·”Ê”Ï
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… √»Ì ‘⁄Ì» «·”Ê”Ì : ›ÕœÀ‰« »Â« Œ·› »‰ ≈»—«ÂÌ„ »‰ „Õ„œ «·„ﬁ—∆ ° ﬁ«· : ÕœÀ‰« √»Ê „Õ„œ «·Õ”‰ »‰ —‘Ìﬁ «·„⁄œ· ° ﬁ«· : ÕœÀ‰« √»Ê ⁄»œ «·—Õ„‰ √Õ„œ »‰ ‘⁄Ì» «·‰”«∆Ì ° ﬁ«· : √Œ»—‰« √»Ê ‘⁄Ì» ° ﬁ«· : √Œ»—‰« «·Ì“ÌœÌ ° ⁄‰ √»Ì ⁄„—Ê ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â »≈ŸÂ«— «·√Ê· „‰ «·„À·Ì‰ Ê«·„ ﬁ«—»Ì‰ Ê»≈œ€«„Â ⁄·Ï ›«—” »‰ √Õ„œ «·„ﬁ—Ì¡ ° Êﬁ«· ·Ì : ﬁ—√  »Â« ﬂ–·ﬂ ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Êﬁ«· ·Ì : ﬁ—√  »Â« «·ﬁ—«‰ ﬂ·Â ﬂ–·ﬂ ⁄·Ï √»Ì ⁄„—«‰ „Ê”Ï »‰ Ã—Ì— «·‰ÕÊÌ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ‘⁄Ì» ° Êﬁ«· : ﬁ—√  ⁄·Ï «·Ì“ÌœÌ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì ⁄„—Ê" & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê: ÊÕœÀ‰« »√’Ê· «·≈œ€«„ „Õ„œ »‰ √Õ„œ ⁄‰ «»‰ „Ã«Âœ ⁄‰ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” ⁄‰ «·œÊ—Ì ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ï ⁄„—Ê° ÊÕœÀ‰« »Â« √Ì÷« √»Ê «·Õ”‰ ‘ÌŒ‰« ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ «·„»«—ﬂ ⁄‰ Ã⁄›— »‰ ”·Ì„«‰ ⁄‰ √»Ì ‘⁄Ì» ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê . " & vbNewLine
        sanadan = sanadan & "Ê—Ã«· √»Ì ⁄„—Ê : Ã„«⁄… „‰ √Â· «·ÕÃ«“ Ê„‰ √Â· «·»’—… ° ›„‰ √Â· „ﬂ… : „Ã«Âœ ° Ê”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„… »‰ Œ«·œ ° Ê⁄ÿ«¡ »‰ √»Ì —»«Õ ° Ê⁄»œ «··Â »‰ ﬂÀÌ— ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ „ÕÌ’‰ ° ÊÕ„Ìœ »‰ ﬁÌ” «·√⁄—Ã «·ﬁ«—∆ ° Ê„‰ √Â· «·„œÌ‰… : Ì“Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—Ì¡ ÊÌ“Ìœ »‰ —Ê„«‰ ° Ê‘Ì»… »‰ ‰’«Õ ° Ê„‰ √Â· «·»’—… : «·Õ”‰ »‰ √»Ì «·Õ”‰ «·»’—Ì ° ÊÌÕÌ »‰ Ì⁄„— ° Ê€Ì—Â„« ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄„‰  ﬁœ„ „‰ «·’Õ«»… Ê€Ì—Â„ . " & vbNewLine
        sanadan = sanadan & "ﬁ·  : Ê√Œ– ”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„…° ÊÌÕÌÏ »‰ Ì⁄„— ° ⁄‰ «»‰ ⁄»«” Ê√Œ– «»‰ ⁄»«” ⁄‰ √»Ì »‰ ﬂ⁄» Ê“Ìœ »‰ À«»  ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -."
        
        ElseIf index = 41 Then
        '«»‰ –ﬂÊ«‰
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… «»‰ –ﬂÊ«‰ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ „Ê”Ï »‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« √Õ„œ »‰ ÌÊ”› «· €·»Ì ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ –ﬂÊ«‰ ° ﬁ«· : ÕœÀ‰« √ÌÊ» »‰  „Ì„ «· „Ì„Ì ° ﬁ«· :ÕœÀ‰« ÌÕÌÌ »‰ «·Õ«—À «·–„«—Ì ° ﬁ«· : ﬁ—√  ⁄·Ï «»‰ ⁄«„— ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ⁄»œ «·⁄“Ì“ »‰ Ã⁄›— «·›«—”Ì «·„ﬁ—Ì¡ Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ï »ﬂ— „Õ„œ »‰ «·Õ”‰ «·‰ﬁ«‘ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ⁄»œ «··Â Â«—Ê‰ »‰ „Ê”Ï »‰ ‘—Ìﬂ «·√Œ›‘ Ê—Ê«Â« «·√Œ›‘ ⁄‰ ⁄»œ «··Â »‰ –ﬂÊ«‰ " & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· «»‰ ⁄«„— «·‹–Ì‹‰ ”‹‹„«Â„ : √»Ê «·œ—œ«¡ ⁄ÊÌ„— »‰ ⁄«„— ’«Õ» —”Ê· «··Â ° Ê«·„€Ì—… »‰ √»Ì ‘Â«» «·„Œ“Ê„Ì ° Ê√Œ‹– √»Ê «·œ—œ«¡ ⁄‹‹‰ «·‰»Ì . Ê√Œ– «·„€Ì—… ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -" & vbNewLine
        
        ElseIf index = 51 Then
        'Õ›’
        sanadan = "ﬁ«· √»‹‹Ê ⁄‹„‹—Ê «·‹œ«‰‹‹Ì ›‹‹‹Ì «·‹ Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… Õ›’ : ›ÕœÀ‰« »Â« √»Ê «·Õ”‰ ÿ«Â‹— »‰ €·»Ê‰ «·„ﬁ—∆ ° ﬁ«· : ÕœÀ‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ’«·Õ «·Â«‘„Ì «·÷—Ì— «·„ﬁ—∆ »«·»’—… ° ﬁ«·: ÕœÀ‰« √»Ê «·⁄»«” √Õ„œ »‰ ”Â· «·√‘‰«‰Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì „Õ„œ ⁄»Ìœ »‰ «·’»«Õ ° Êﬁ«·: ﬁ—√  ⁄·Ï Õ›’ ° Êﬁ«·: ﬁ—√  ⁄·Ï ⁄«’‹„ ° Êﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï ‘ÌŒ‰« √»Ì «·Õ”‰ Êﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï «·Â«‘„Ì Êﬁ«·: ﬁ—√  ⁄·Ï «·√‘‰«‰Ì ⁄‰ ⁄»Ìœ ⁄‰ Õ›’ ⁄‰ ⁄«’‹„ . " & vbNewLine
        sanadan = sanadan & "Ê—Ã«· ⁄«’„ «·‹–Ì‹‰ ”„«Â„ «À‰«‰ : √»Ê ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ê „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·»  ° Ê√»Ì »‰ ﬂ⁄»  ° Ê“Ìœ »‰ À«»   ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï - ° √Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰  ° Ê«»‰ „”⁄Êœ  ° ⁄‰ —”Ê· «··Â - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 61 Then
        'Œ·«œ
        sanadan = "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… Œ·«œ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« √Õ„œ »‰ „Ê”Ï ° ﬁ«· : ÕœÀ‰« ÌÕÌÏ »‰ √Õ„œ »‰ Â«—Ê‰ «·„“Êﬁ ° ⁄‰ √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì ° ⁄‰ Œ·«œ ° ⁄‰ ”·Ì„ ° ⁄‰ Õ„“… ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ «·÷—Ì— ‘ÌŒ‰« ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ‘‰»Ê– ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ‘«–«‰ «·ÃÊÂ—Ì «·„ﬁ—Ì ° Êﬁ«· :ﬁ—√  ⁄·Ï Œ·«œ Êﬁ«· : ﬁ—√  ⁄·Ï ”·Ì„ ° Êﬁ—√ ”·Ì„ ⁄·Ï Õ„“…." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
        sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 71 Then
        '√»Ê «·Õ«—À
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… √»Ì «·Õ«—À : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« »Â« «»‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« „Õ„œ »‰ ÌÕÌÏ ( «·ﬂ”«∆Ì «·’€Ì—) ° ⁄‰ √»Ì «·Õ«—À ° ⁄‰ «·ﬂ”«∆Ì ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—Ì¡ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·ﬁ«”„ “Ìœ »‰ ⁄·Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï √Õ„œ »‰ «·Õ”‰ «·„⁄—Ê› »«·»ÿÌ ° Êﬁ«· :ﬁ—√  ⁄·Ï „Õ„œ »‰ ÌÕÌÏ ( «·ﬂ”«∆Ì «·’€Ì—) ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì «·Õ«—À ° Êﬁ«· : ﬁ—√  ⁄·Ï «·ﬂ”«∆Ì ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· «·ﬂ”«∆Ì : Õ„“… »‰ Õ»Ì» «·“Ì«  ° Ê⁄Ì”Ï »‰ ⁄„— «·Â„–«‰Ì ° Ê„Õ„œ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° Ê€Ì—Â„ „‰ „‘ÌŒ… «·ﬂÊ›ÌÌ‰ €Ì— √‰ „«œ… ﬁ—«¡ Â Ê«⁄ „«œÂ ›Ì «Œ Ì«—Â ⁄‰ Õ„“… ° Êﬁœ –ﬂ—‰« « ’«· ﬁ—«¡ Â ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
        sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 81 Then
        '«»‰ Ã„«“
        sanadan = "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… «»‰ Ã„«“ : ›ÕœÀ‰« »Â« √»Ê ≈”Õ«ﬁ ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ≈»—«ÂÌ„ »‰ Õ« „ «·Ã–«„Ì »ﬁ—«¡ Ì ⁄·ÌÂ ⁄‰ √»Ì Õ›’ ⁄„— »‰ €‹œÌ— »‰ «·ﬁÊ«” «·œ„‘ﬁÌ ° √‰»√‰« √»Ê «·Ì„‰ »‰ «·Õ”‰ «·»€œ«œÌ ° √Œ»—‰« √»Ê „Õ„œ ”»ÿ «·ŒÌ«ÿ ° √Œ»—‰« «·√” «– √»Ê «·⁄“ „Õ„œ »‰ «·Õ”Ì‰ »‰ »‰œ«— «·Ê«”ÿÌ ° √Œ»—‰« «·≈„«„ √»Ê «·ﬁ«”„ ÌÊ”› »‰ Ã»«—… «·Â–·Ì ° √Œ»—‰« √»Ê ‰’— „‰’Ê— »‰ „Õ„œ «·ﬁÂ‰œ“Ì ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ «·Œ»«“Ì ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ «·›÷· «·ÃÊÂ—Ì ° √Œ»—‰« „Õ„œ »‰ √Õ„œ »‰ «·Õ”‰ «·Àﬁ›Ì «·ﬂ”«∆Ì ° √Œ»—‰« „Õ„œ »‰ ⁄»œ «··Â »‰ ‘«ﬂ— «·’Ì—›Ì ° √Œ»—‰« √»Ê «·⁄»«” √Õ„œ »‰ ”Â· «·ÿÌ«‰ ° √Œ»—‰« √»Ê ⁄„—«‰ „Ê”Ï »‰ ⁄»œ «·—Õ„‰ «·»“«“ ° √Œ»—‰« „Õ„œ »‰ ⁄Ì”Ï »‰ ≈»—«ÂÌ„ »‰ —“Ì‰ «·√’»Â«‰Ì ° √Œ»—‰« ”·Ì„«‰ »‰ œ«Êœ »‰ ⁄·Ì »‰ ⁄»œ «··Â »‰ ⁄»«” «·Â«‘„Ì ° √Œ»—‰« ≈”„«⁄Ì· »‰ Ã⁄›— »‰ √»Ì ﬂÀÌ— «·„œ‰Ì ° √Œ»—‰« ”·Ì„«‰ »‰ „”·„ «»‰ Ã„«“." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì „Õ„œ »‰ ⁄»œ «·—Õ„‰ «·Õ‰›Ì ° Êﬁ—√ »Â« «·ﬁ—«‰ ﬂ·Â ⁄·Ï „Õ„œ »‰ √Õ„œ «·’«∆€ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ≈”Õ«ﬁ »‰ ›«—” ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Ì„‰ ° Êﬁ—√ »Â« ⁄·Ï ”»ÿ «·ŒÌ«ÿ ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì ÿ«Â— √Õ„œ »‰ ⁄·Ì »‰ ⁄»Ìœ «··Â »‰ ”Ê«— ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄·Ì «·Õ”‰ »‰ «·›÷· «·‘—„ﬁ«‰Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄»œ «··Â »‰ «·„“—»«‰ «·√’»Â«‰Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄„— „Õ„œ »‰ √Õ„œ »‰ ⁄„— «·Œ—ﬁÌ ° Êﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ Ã⁄›— »‰ „Õ„Êœ «·√‘‰«‰Ì ° Êﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ «·Àﬁ›Ì «·ﬂ”«∆Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ ‘«ﬂ— ° Êﬁ—√ »Â« ⁄·Ï «»‰ ”Â· «·ÿÌ«‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄„—«‰ «·»“«“ ° Êﬁ—√ »Â« ⁄·Ï «»‰ —“Ì‰ ° Êﬁ—√ »Â« ⁄·Ï «·Â«‘„Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ Ã⁄›— ° Êﬁ—√ »Â« ⁄·Ï «»‰ Ã„«“ ° Êﬁ—√ «»‰ Ã„«“ ° Ê«»‰ Ê—œ«‰ ° ⁄·Ï √»Ì Ã⁄›— ." & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· √»Ì Ã⁄›— À·«À… : „Ê·«Â ⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° Ê√»Ê Â—Ì—… ° Ê«»‰ ⁄»«” . Êﬁ—√ Âƒ·«¡ «·À·«À… ⁄·Ï √»Ì »‰ ﬂ⁄» ° Êﬁ—√ √»Ê Â—Ì—… ° Ê«»‰ ⁄»«” ° √Ì÷« ⁄·Ï “Ìœ »‰ À«»  . Ê√Œ– “Ìœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ -° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 91 Then
        '—ÊÕ
        sanadan = "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… —ÛÊÕ : ›ÕœÀ‰« »Â« «·‘ÌŒ √»Ê «·⁄»«” √Õ„œ »‰ „Õ„œ »‰ «·Õ”Ì‰ «·‘Ì—«“Ì »ﬁ—«¡ Ì ⁄·ÌÂ ⁄‰ «·≈„«„ √»Ì «·Õ”‰ ⁄·Ì »‰ √Õ„œ «·„ﬁœ”Ì ° √Œ»—‰« √»Ê «·Ì„‰ «·ﬂ‰œÌ ‘›«Â« ° √Œ»—‰« √»Ê „Õ„œ «·»€œ«œÌ ° √Œ»—‰« √»Ê «·›÷· «·‘—Ì› «·„ﬂÌ ° √Œ»—‰« „Õ„œ »‰ «·Õ”Ì‰ «·›«—”Ì ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ≈»—«ÂÌ„ »‰ Œ‘‰«„ «·„«·ﬂÌ «·»’—Ì √Œ»—‰« √»Ê «·⁄»«” „Õ„œ »‰ Ì⁄ﬁÊ» »‰ «·ÕÃ«Ã »‰ „⁄«ÊÌ… «· Ì„Ì ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ ÊÂ» »‰ ÌÕÌÏ »‰ «·⁄·«¡ «·Àﬁ›Ì «·ﬁ“«“ ° √Œ»—‰« —ÊÕ »‰ ⁄»œ «·„ƒ„‰ «·»’—Ì ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï „Õ„œ »‰ √Õ„œ »«·ﬁ«Â—… «·„Õ—Ê”… ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â «·’«∆€ ° Êﬁ—√ »Â« ⁄·Ï √»Ì ≈”Õ«ﬁ «·œ„‘ﬁÌ Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï ⁄»œ «··Â »‰ ⁄·Ì ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì ÿ«Â— »‰ ”Ê«— ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·ﬁ«”„ «·„”«›— »‰ «·ÿÌ» »‰ ⁄»«œ «·»’—Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ Œ‘‰«„ ° Êﬁ—√ »Â« ⁄·Ï «»‰ ⁄»« ” «· Ì„Ì ° Êﬁ—√ »Â« ⁄·Ï «»‰ ÊÂ» ° Êﬁ—√ »Â« ⁄·Ï —ÊÕ ° Êﬁ—√ »Â« ⁄·Ï Ì⁄ﬁÊ» ." & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· Ì⁄ﬁÊ» «·–Ì‰ ”„«Â„ √—»⁄… : √»Ê «·„‰–— ”·«„ »‰ ”·Ì„«‰ «·ÿÊÌ· ° Ê‘Â«» »‰ ‘—‰›… ° Ê„ÂœÌ »‰ „Ì„Ê‰ ° Ê√»Ê «·√‘Â» Ã⁄›— »‰ ÕÌ«‰ «·⁄ÿ«—œÌ .ÊﬁÌ· ≈‰ Ì⁄ﬁÊ» ﬁ—√ ⁄·Ï √»Ì ⁄„—Ê »‰ «·⁄·«¡ Êﬁ—√ ”·«„ ⁄·Ï ⁄«’„ Ê√»Ì ⁄„—Ê ° Êﬁ‹‹‹—√ ‘Â«» «·ÃÕœ—Ì Êﬁ—√ ⁄«’„ ⁄·Ï «·Õ”‰ «·»’—Ì Ê⁄·Ï ”·Ì„«‰ »‰ ﬁ … Êﬁ—√ ”·Ì„«‰ ⁄·Ï «»‹‰ ⁄»« ” Êﬁ—√ „ÂœÌ ⁄·Ï ‘⁄Ì» »‰ «·Õ»Õ«» Êﬁ—√ ⁄·Ï √»Ì «·⁄«·Ì… «·—Ì«ÕÌ Êﬁ—√ ⁄·Ï √»Ì Ê“Ìœ Êﬁ—√ √»Ê «·√‘Â» ⁄·Ï √»Ì —Ã«¡ ⁄„—«‰ »‰ „·Õ«‰ «·⁄ÿ«—œÌ Êﬁ—√ ⁄·Ï √»Ì „Ê”‹‹‹Ï «·√‘⁄—Ì Êﬁ—√ ⁄·Ï —”Ê· «··Â ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 101 Then
        '≈œ—Ì”
        sanadan = "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "Ê√„« —Ê«Ì… ≈œ—Ì” : ›ÕœÀ‰« »Â« √Õ„œ »‰ „Õ„œ »‰ «·Õ”Ì‰ «·›«—”Ì »ﬁ—«¡ Ì ⁄·ÌÂ ° √Œ»—‰« ⁄·Ì »‰ √Õ„œ ›Ì„« ‘«›Â‰Ì »Â °⁄‰ “Ìœ »‰ «·Õ”‰ «·»€œ«œÌ ° √Œ»—‰« √»Ê «·ﬁ«”„ »‰ √Õ„œ «·Õ—Ì—Ì ° √Œ»—‰« √»Ê »ﬂ—„Õ„œ »‰ ⁄»Ì »‰ „Õ„œ «·ŒÌ«ÿ ° √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ⁄»œ «··Â «·Õ–«¡ ° √Œ»—‰« √»Ê ≈”Õ«ﬁ ≈»—«ÂÌ„ »‰ «·Õ”Ì‰ »‰ ⁄»œ «··Â «·‰”«Ã «·„⁄—Ê› »«·‘ÿÌ ° √Œ»—‰« ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ «·Õœ«œ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·‘ÌŒ √»Ì „Õ„œ ⁄»œ «·—Õ„‰ »‰ √Õ„œ «·Ê«”ÿÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„⁄œ· ° Êﬁ—√ »Â« ⁄·Ï ≈»—«ÂÌ„ »‰ √Õ„œ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Ì„‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì „Õ„œ ”»ÿ «·ŒÌ«ÿ ° ﬁ«· : Êﬁ—√  »Â« «·ﬁ—¬‰ „‰ √Ê·Â ≈·Ï ¬Œ—Â ⁄·Ï «·≈„«„Ì‰ «·‘—Ì› √»Ì «·›÷· ⁄»œ «·ﬁ«Â— »‰ ⁄»œ «·”·«„ «·⁄»«”Ì ° Ê√»Ì «·„⁄«·Ì À«»  »‰ »‰œ«— »‰ ≈»—«ÂÌ„ «·»ﬁ«· ° ›√„« «·‘—Ì› ›√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â „Õ„œ »‰ «·Õ”Ì‰ «·ﬂ«—“Ì‰Ì ° Ê√Œ»—Â √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ √»Ì «·⁄»«” «·Õ”‰ »‰ ”⁄Ìœ »‰ Ã⁄›— «·„ÿÊ⁄Ì ° Ê√„« √»Ê «·„⁄«·Ì ›√Œ»—‰Ì √‰Â ﬁ—√ »Â« ⁄·Ï «·≈„«„ «·ﬁ«÷Ì √»Ì «·⁄·«¡ „Õ„œ »‰ ⁄·Ì »‰ Ì⁄ﬁÊ» «·Ê«”ÿÌ ° Êﬁ—√ «·Ê«”ÿÌ »Â« „‰ «·ﬂ «» ⁄·Ï «·≈„«„ √»Ì »ﬂ— √Õ„œ »‰ Ã⁄›— »‰ Õ„œ«‰ »‰ „«·ﬂ «·ﬁÿÌ⁄Ì ° Êﬁ—√ «·ﬁÿÌ⁄Ì Ê«·„ÿÊ⁄Ì Ã„Ì⁄« ⁄·Ï ≈œ—Ì” ° Êﬁ—√ ≈œ—Ì” ⁄·Ï Œ·› ° Ê«··Â «·„Ê›ﬁ . " & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Œ·› : Ê—Ã«· Œ·› ”·Ì„ ’«Õ» Õ„“… ° ÊÌ⁄ﬁÊ» »‰ Œ·Ì›… «·√⁄‘Ï ’«Õ» √»Ì »ﬂ— ° Ê√»Ê “Ìœ ”⁄Ìœ ”⁄Ìœ »‰ √Ê” «·√‰’«—Ì ’«Õ» «·„›÷· «·÷»Ì Ê√»«‰ «·⁄ÿ«— ° Êﬁ—√ √»Ê »ﬂ— ° Ê«·„›÷· ° Ê√»«‰ ⁄·Ï ⁄«’„ . Ê—ÊÏ «·ﬁ—«¡… √Ì÷« ⁄‰ «·ﬂ”«∆Ì Ê⁄‰ ÌÕÌÏ »‰ ¬œ„ ⁄‰ √»Ì »ﬂ— ° Ê«··Â «·„Ê›ﬁ . ﬁ·  : Ê√Œ– ⁄«’„ ⁄‰ √»Ì ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ì „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·» ° Ê√»Ì »‰ ﬂ⁄» ° Ê“Ìœ »‰ À«»  ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        sanadan = sanadan & "Ê√Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰ ° Ê«»‰ „”⁄Êœ ° ⁄‰ —”Ê· «··Â ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -. Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ . Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 12 Then
        'ﬁ«·Ê‰
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan + "√„« —Ê«Ì… ﬁ«·Ê‰ : ›ÕœÀ‰« »Â« √Õ„œ »‰ ⁄„— »‰ „Õ„œ «·ÃÌ“Ì ° ﬁ«·: ÕœÀ‰« „Õ„œ »‰ √Õ„œ »‰ „‰Ì— ° ﬁ«·: ÕœÀ‰« ⁄»œ «··Â »‰ ⁄Ì”Ï «·„œ‰Ì ° ﬁ«·:ÕœÀ‰« ﬁ«·Ê‰ ⁄‰ ‰«›⁄° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ‘ÌŒÌ √»Ì «·› Õ ›«—” »‰ √Õ„œ »‰ „Ê”Ï »‰ ⁄„—«‰ ° «·„ﬁ—Ì¡ «·÷—Ì— ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—∆ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈»—«ÂÌ„ »‰ ⁄„— «·„ﬁ—∆ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”Ì‰ √Õ„œ »‰ ⁄À„«‰ »‰ Ã⁄›— »‰ »ÊÌ«‰ ° Êﬁ«·:ﬁ—√  ⁄·Ï √»Ì »ﬂ— √Õ„œ »‰ „Õ„œ »‰ «·√‘⁄À Êﬁ«·: ﬁ—√  ⁄·Ï √»Ì ‰‘Ìÿ „Õ„œ »‰ Â«—Ê‰ ° Êﬁ«·: ﬁ—√  ⁄·Ï ﬁ«·Ê‰ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‰«›⁄ ." & vbNewLine
        sanadan = sanadan + "Ê—Ã«· ‰«›⁄ «·–Ì‰ ”„«Â„ Œ„”… : √»Ê Ã⁄›— Ì“ Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—∆ ° Ê√»Ê œ«Êœ ⁄»œ «·—Õ„‰ »‰ Â—„“ «·√⁄—Ã ° Ê‘Ì»… »‰ ‰’«Õ «·ﬁ«÷Ì ° Ê√»Ê ⁄»œ «··Â „”·„ »‰ Ã‰œ» «·Â–·Ì «·ﬁ«’ ° Ê√»Ê —ÊÕ Ì“Ìœ »‰ —Ê„«‰ ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄‰ √»Ì Â—Ì—… ° Ê«»‰ ⁄»«” ° Ê⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° ⁄‰ √»Ì »‰ ﬂ⁄» ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 22 Then
        '«·»“Ï
        sanadan = "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—  " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… «·»“Ì : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ «·ﬂ« » ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ „Ê”Ï ° ﬁ«·: ÕœÀ‰« „÷— »‰ „Õ„œ «·÷»Ì ° ﬁ«·:ÕœÀ‰« √Õ„œ »‰ √»Ì »“… ° ﬁ«·: ﬁ—√  ⁄·Ï ⁄ﬂ—„… »‰ ”·Ì„«‰ »‰ ⁄«„— ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈”„«⁄Ì· »‰ ⁄»œ «··Â «·ﬁ”ÿ ° Êﬁ«· : ﬁ—√  ⁄·Ï «»‰ ﬂÀÌ— ‰›”Â ° ﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·ﬁ«”„ ⁄»œ «·⁄“Ì“ »‰ Ã⁄›— »‰ „Õ„œ «·„ﬁ—Ì¡ «·›«—”Ì ° Êﬁ«· ·Ì: ﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ «·Õ”‰ «·‰ﬁ«‘ ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì —»Ì⁄… „Õ„œ »‰ ≈”Õ«ﬁ «·— »⁄Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï «·»“Ì ." & vbNewLine
        sanadan = sanadan & " Ê—Ã‹‹«· «»‰ ﬂÀÌ— «·‹–Ì‹‰ ”„«Â„ À·«À… : ⁄»œ «··Â »‰ «·”«∆» «·„Œ“Ê„Ì ’«Õ» —”Ê· «··Â  Ê„Ã«Âœ »‰ Ã»— √»Ê «·ÕÃ«Ã „Ê·Ï ﬁÌ” »‰ «·”«∆» ° Êœ—»«” „Ê·Ï «»‰ ⁄»«” . Ê√Œ– ⁄»œ «··Â ⁄‰ √»Ì »‰ ﬂ⁄» ‰›”Â. Ê√Œ– „Ã«Âœ Êœ—»«”° ⁄‰ «»‰ ⁄»«”° ⁄‰ √»Ì ° Ê“Ìœ »‰ À«»  ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  °⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -."
        
        ElseIf index = 32 Then
        '«·œÊ—Ï
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… √»Ì ⁄„— «·œÊ—Ì : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ ⁄·Ì ° ﬁ«·: √Œ»—‰« √»Ê ⁄Ì”Ï „Õ„œ »‰ √Õ„œ »‰ ﬁÿ‰ ”‰… À„«‰ ⁄‘—… ÊÀ·«À„«∆…° ﬁ«·: √Œ»—‰« √»Ê Œ·«œ ”·Ì„«‰ »‰ Œ·«œ ﬁ«·:ÕœÀ‰« «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê ° ﬁ«· √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â „‰ ÿ—Ìﬁ √»Ì ⁄„— «·œÊ—Ì ⁄·Ï ‘ÌŒ‰« ⁄»œ «·⁄“ Ì“ »‰ Ã⁄›— »‰ „Õ„œ »‰ ≈”Õ«ﬁ «·»€œ«œÌ «·›«—”Ì «·„ﬁ—Ì¡ ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì ÿ«Â— ⁄»œ «·Ê«Õœ »‰ ⁄„— »‰ √»Ì Â«‘„ «·„ﬁ—Ì¡ ° „« ·« √Õ’ÌÂ ﬂÀ—… ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì »ﬂ— »‰ „Ã«Âœ ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·“⁄—«¡ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” Êﬁ«· :ﬁ—√  ⁄·Ï √»Ì ⁄„— ° Êﬁ«· : ﬁ—√  »Â« ⁄·Ï «·Ì“ÌœÌ ° Êﬁ«· ﬁ—√  »Â« ⁄·Ï : √»Ì ⁄„—Ê. " & vbNewLine
        sanadan = sanadan & "ﬁ«· √»Ê ⁄„—Ê: ÊÕœÀ‰« »√’Ê· «·≈œ€«„ „Õ„œ »‰ √Õ„œ ⁄‰ «»‰ „Ã«Âœ ⁄‰ ⁄»œ «·—Õ„‰ »‰ ⁄»œÊ” ⁄‰ «·œÊ—Ì ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ï ⁄„—Ê° ÊÕœÀ‰« »Â« √Ì÷« √»Ê «·Õ”‰ ‘ÌŒ‰« ° ﬁ«· : ÕœÀ‰« ⁄»œ «··Â »‰ «·„»«—ﬂ ⁄‰ Ã⁄›— »‰ ”·Ì„«‰ ⁄‰ √»Ì ‘⁄Ì» ⁄‰ «·Ì“ÌœÌ ⁄‰ √»Ì ⁄„—Ê . " & vbNewLine
        sanadan = sanadan & "Ê—Ã«· √»Ì ⁄„—Ê : Ã„«⁄… „‰ √Â· «·ÕÃ«“ Ê„‰ √Â· «·»’—… ° ›„‰ √Â· „ﬂ… : „Ã«Âœ ° Ê”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„… »‰ Œ«·œ ° Ê⁄ÿ«¡ »‰ √»Ì —»«Õ ° Ê⁄»œ «··Â »‰ ﬂÀÌ— ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ „ÕÌ’‰ ° ÊÕ„Ìœ »‰ ﬁÌ” «·√⁄—Ã «·ﬁ«—∆ ° Ê„‰ √Â· «·„œÌ‰… : Ì“Ìœ »‰ «·ﬁ⁄ﬁ«⁄ «·ﬁ«—Ì¡ ÊÌ“Ìœ »‰ —Ê„«‰ ° Ê‘Ì»… »‰ ‰’«Õ ° Ê„‰ √Â· «·»’—… : «·Õ”‰ »‰ √»Ì «·Õ”‰ «·»’—Ì ° ÊÌÕÌ »‰ Ì⁄„— ° Ê€Ì—Â„« ° Ê√Œ– Âƒ·«¡ «·ﬁ—«¡… ⁄„‰  ﬁœ„ „‰ «·’Õ«»… Ê€Ì—Â„ . " & vbNewLine
        sanadan = sanadan & "ﬁ·  : Ê√Œ– ”⁄Ìœ »‰ Ã»Ì— ° Ê⁄ﬂ—„…° ÊÌÕÌÏ »‰ Ì⁄„— ° ⁄‰ «»‰ ⁄»«” Ê√Œ– «»‰ ⁄»«” ⁄‰ √»Ì »‰ ﬂ⁄» Ê“Ìœ »‰ À«»  ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -."
        
        ElseIf index = 42 Then
        'Â‘«„
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… Â‘«„ : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ﬁ«·: ÕœÀ‰« «»‰ „Ã«Âœ ° ﬁ«· : ÕœÀ‰« «·Õ”Ì‰ »‰ „Â—«‰ «·Ã„«· ° ﬁ«· :ÕœÀ‰« √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì ° ﬁ«· : ÕœÀ‰« Â‘«„ »‰ ⁄„«— ° ﬁ«·: ÕœÀ‰« ⁄—«ﬂ »‰ Œ«·œ «·„—Ì ° ﬁ«· :ﬁ—√  ⁄·Ï ÌÕÌÌ »‰ «·Õ«—À «·–„«—Ì ° Êﬁ«·: ﬁ—√  ⁄·Ï ⁄»œ «··Â »‰ ⁄«„— ° ﬁ«· : √»Ê ⁄„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ ‘ÌŒ‰« ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ⁄»œ «··Â »‰ «·Õ”Ì‰ «·„ﬁ—Ì¡ ° Ê ﬁ«· : ﬁ—√  »Â« ⁄·Ï „Õ„œ »‰ √Õ„œ »‰ ⁄»œ«‰ ° Êﬁ«· : ﬁ—√  ⁄·Ï «·Õ·Ê«‰Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï Â‘«„ " & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· «»‰ ⁄«„— «·‹–Ì‹‰ ”‹‹„«Â„ : √»Ê «·œ—œ«¡ ⁄ÊÌ„— »‰ ⁄«„— ’«Õ» —”Ê· «··Â ° Ê«·„€Ì—… »‰ √»Ì ‘Â«» «·„Œ“Ê„Ì ° Ê√Œ‹– √»Ê «·œ—œ«¡ ⁄‹‹‰ «·‰»Ì . Ê√Œ– «·„€Ì—… ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -" & vbNewLine
        
        ElseIf index = 52 Then
        '‘⁄»…
        sanadan = "ﬁ«· √»‹‹Ê ⁄‹„‹—Ê «·‹œ«‰‹‹Ì ›‹‹‹Ì «·‹ Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… √»Ì »ﬂ— ‘⁄»…: ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ »‰ ⁄·Ì «·ﬂ« » ﬁ«·: ÕœÀ‰« »‰ „Ã«Âœ ﬁ«·: ÕœÀ‰« ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ⁄„— «·ÊﬂÌ⁄Ì ° ﬁ«·:ÕœÀ‰« √»Ì ﬁ«·:ÕœÀ‰« ÌÕÌÌ »‰ √œ„ ° ﬁ«·: ÕœÀ‰« √»Ê »ﬂ— ⁄‰ ⁄«’„ ° ﬁ«· √»Ê ⁄„—Ê: Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ›«—” »‰ √Õ„œ «·„ﬁ—Ì¡ ° Ê ﬁ«· ·Ì: ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ «·„ﬁ—Ì¡ ° Êﬁ«·: ﬁ—√  ⁄·Ï ≈»—«ÂÌ„ »‰ ⁄»œ «·—Õ„‰ »‰ √Õ„œ «·„ﬁ—Ì¡ «·»€œ«œÌ Êﬁ«·: ﬁ—√  ⁄·Ï ÌÊ”› »‰ Ì⁄ﬁÊ» «·Ê«”ÿÌ ° Êﬁ«·: ﬁ—√  ⁄·Ï ‘⁄Ì» »‰ √ÌÊ» «·’—Ì›Ì‰Ì ° Êﬁ«·: ﬁ—√  »Â« ⁄·Ï ÌÕÌÌ »‰ √œ„ ⁄‰ √»Ï »ﬂ— ⁄‰ ⁄«’„." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· ⁄«’„ «·‹–Ì‹‰ ”„«Â„ «À‰«‰ : √»Ê ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ê „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·»  ° Ê√»Ì »‰ ﬂ⁄»  ° Ê“Ìœ »‰ À«»   ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï - ° √Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰  ° Ê«»‰ „”⁄Êœ  ° ⁄‰ —”Ê· «··Â - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 62 Then
        'Œ·›
        sanadan = "ﬁ«· √»‹Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… Œ·› : ›ÕœÀ‰« »Â« „Õ„œ »‰ √Õ„œ ° ﬁ«· : ÕœÀ‰« «»‰ „Ã«Âœ ° ÕœÀ‰« ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ ° ﬁ«· : ÕœÀ‰« Œ·› ° ﬁ«·: ⁄‰ ”·Ì„ ⁄‰ Õ„“… ° Ê ﬁ«· √»‹‹Ê ⁄‹„‹—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·Õ”‰ ‘ÌŒ‰« ° Ê ﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”‰ „Õ„œ »‰ ÌÊ”› »‰ ‰Â«— «·Õ— ﬂÌ »«·»’—… ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï √»Ì «·Õ”Ì‰ √Õ„œ »‰ ⁄À„«‰ »‰ Ã⁄›— »‰ »ÊÌ«‰ ° Êﬁ«· ·Ì :ﬁ—√  ⁄·Ï ≈œ—Ì” »‰ ⁄»œ «·ﬂ—Ì„ ﬁ»· √‰ Ìﬁ—Ì¡ »«Œ Ì«— Œ·› ° Êﬁ«· ·Ì : ﬁ—√  ⁄·Ï Œ·› ° Êﬁ«· : ﬁ—√  ⁄·Ï ”·Ì„ ° Ê ﬁ«· : ﬁ—√  ⁄·Ï Õ„“… ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
        sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 72 Then
        '√»Ê ⁄„—Ê «·œÊ—Ï
        sanadan = "ﬁ«· √»Ê ⁄„—Ê «·‹œ«‰Ì ›Ì «· Ì”Ì—:" & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… √»Ì ⁄„—Ê «·œÊ—Ì : ›ÕœÀ‰« »Â« √»Ê „Õ„œ ⁄»œ «·—Õ„‰ »‰ ⁄„— »‰ „Õ„œ «·„⁄œ· ° ﬁ«· : ÕœÀ‰« √»Ê ⁄„— ⁄»œ «··Â »‰ √Õ„œ »‰ œÌ“ÊÌÂ «·œ„‘ﬁÌ ° ﬁ«· : ÕœÀ‰« Ã⁄›— »‰ „Õ„œ »‰ √”œ «·‰’Ì»Ì ° ﬁ«· : ÕœÀ‰« √»Ê ⁄„— «·œÊ—Ì ° ⁄‰ «·ﬂ”«∆Ì ° Ê ﬁ«· √»Ê ⁄‹„—Ê : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï √»Ì «·› Õ ° Êﬁ«· ·Ì : ﬁ—√  »Â« ⁄·Ï ⁄»œ «·»«ﬁÌ »‰ «·Õ”‰ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄·Ì »‰ «·Ã·‰œÌ «·„Ê’·Ì ° Ê ﬁ«· :ﬁ—√  ⁄·Ï Ã⁄›— »‰ „Õ„œ ° Êﬁ«· : ﬁ—√  ⁄·Ï √»Ì ⁄„— «·œÊ—Ì ° Êﬁ«· : ﬁ—√  ⁄·Ï «·ﬂ”«∆Ì ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· «·ﬂ”«∆Ì : Õ„“… »‰ Õ»Ì» «·“Ì«  ° Ê⁄Ì”Ï »‰ ⁄„— «·Â„–«‰Ì ° Ê„Õ„œ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° Ê€Ì—Â„ „‰ „‘ÌŒ… «·ﬂÊ›ÌÌ‰ €Ì— √‰ „«œ… ﬁ—«¡ Â Ê«⁄ „«œÂ ›Ì «Œ Ì«—Â ⁄‰ Õ„“… ° Êﬁœ –ﬂ—‰« « ’«· ﬁ—«¡ Â ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ ." & vbNewLine
        sanadan = sanadan & "Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ - ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 82 Then
        '«»‰ Ê—œ«‰
        sanadan = "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… «»‰ Ê—œ«‰ : ›ÕœÀ‰« »Â« «·‘ÌŒ √»Ê Õ›’ ⁄„— »‰ «·Õ”‰ »‰ „“Ìœ «·„—«€Ì »ﬁ—«¡ Ì ⁄·ÌÂ ﬁ«· : √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ √Õ„œ »‰ ⁄»œ «·Ê«Õœ «·”⁄œÌ „‘«›Â… ⁄‰ «·≈„«„ √»Ì «·Ì„‰ “Ìœ »‰ «·Õ”‰ «··€ÊÌ ° ﬁ«· : √Œ»—‰« √»Ê „Õ„œ ⁄»œ «··Â »‰ ⁄·Ì «·»€œ«œÌ √Œ»—‰« «·‘—Ì› √»Ê «·›÷· ⁄»œ «·ﬁ«Â— »‰ ⁄»œ «·”·«„ «·⁄»«”Ì ° √Œ»—‰« √»Ê ⁄»œ «··Â „Õ„œ »‰ «·Õ”Ì‰ «·ﬂ«—“Ì‰Ì ° √Œ»—‰« √»Ê «·›—Ã „Õ„œ »‰ √Õ„œ »‰ ≈»—«ÂÌ„ «·‘ÿÊÌ ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ √Õ„œ »‰ Â«—Ê‰ «·—«“Ì ° √Œ»—‰« √»Ê «·⁄»«” «·›÷· »‰ ‘«–«‰ »‰ ⁄Ì”Ï «·—«“Ì √Œ»—‰« √»Ê «·Õ”‰ √Õ„œ »‰ Ì“Ìœ «·Õ·Ê«‰Ì °√Œ»—‰« ⁄Ì”Ï »‰ „Ì‰« ﬁ«·Ê‰ ° √Œ»—‰« ⁄Ì”Ï »‰ Ê—œ«‰." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì ⁄»œ «··Â „Õ„œ ⁄»œ «·—Õ„‰ »‰ ⁄·Ì «·‰ÕÊÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï «·≈„‹‹«„ √»Ì ⁄»œ „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„’—Ì ° ﬁ«· : ﬁ—√  »Â« «·ﬁ—¬‰ ⁄·Ï «·ﬂ„«· ≈»—«ÂÌ„ »‰ √Õ„œ »‰ ›«—” «· „Ì„Ì ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·Ì„‰ «·ﬂ‰œÌ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «·≈„«„ √»Ì „‰’Ê— „Õ„œ »‰ ⁄»œ «·„·ﬂ »‰ «·Õ”‰ »‰ ŒÌ—Ê‰ «·»€œ«œÌ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·ﬁ«”„ ⁄»œ «·”Ìœ »‰ ⁄ «» «·„ﬁ—Ì¡ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì ÿ«Â— „Õ„œ »‰ Ì«”Ì‰ «·Õ·»Ì ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï √»Ì «·›—Ã «·‘ÿÊÌ ﬁ«·: ﬁ—√  »Â« ⁄·Ï √»Ì »ﬂ— »‰ Â«—Ê‰ ° ﬁ«·: ﬁ—√  »Â« ⁄·Ï «·›÷· »‰ ‘«–«‰ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «·Õ·Ê«‰Ì ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï ﬁ«·Ê‰ ° ﬁ«· : ﬁ—√  »Â« ⁄·Ï «»‰ Ê—œ«‰ . " & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· √»Ì Ã⁄›— À·«À… : „Ê·«Â ⁄»œ «··Â »‰ ⁄Ì«‘ »‰ √»Ì —»Ì⁄… ° Ê√»Ê Â—Ì—… ° Ê«»‰ ⁄»«” . Êﬁ—√ Âƒ·«¡ «·À·«À… ⁄·Ï √»Ì »‰ ﬂ⁄» ° Êﬁ—√ √»Ê Â—Ì—… ° Ê«»‰ ⁄»«” ° √Ì÷« ⁄·Ï “Ìœ »‰ À«»  . Ê√Œ– “Ìœ ⁄‰ «·‰»Ì - ’·Ï «··Â ⁄·ÌÂ Ê ”·„ -° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 92 Then
        '—ÊÌ”
        sanadan = "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "›√„« —Ê«Ì… —ÊÌ” : ›ÕœÀ‰« »Â« «·‘ÌŒ «·≈„«„ √»Ê «·⁄»«” √Õ„œ »‰ „Õ„œ »‰ «·Œ÷— «·Õ‰›Ì »ﬁ—«¡ Ì ⁄·ÌÂ ﬁ«·: √Œ»—‰« : √»Ê «·⁄»«” √Õ„œ »‰ √»Ì ÿ«·» »‰ √»Ì «·‰⁄„ «·’«·ÕÌ ﬁ—«¡… ⁄·ÌÂ ° √Œ»—‰« √»Ê ÿ«·» ⁄»œ «··ÿÌ› »‰ „Õ„œ »‰ «·ﬁ»ÌÿÌ ° ›Ì ﬂ «»Â √Œ»—‰« »Â« √»Ê »ﬂ— √Õ„œ »‰ «·„ﬁ—» «·ﬂ—ŒÌ ﬁ—«¡… ⁄·ÌÂ ° √Œ»—‰« √»Ê ÿ«Â— √Õ„œ »‰ ⁄·Ì «·„ﬁ—Ì¡ «·√” «– √Œ»—‰« √»Ê «·Õ”‰ ⁄·Ì »‰ „Õ„œ »‰ ⁄·Ì «·ŒÌ«ÿ ° √Œ»—‰« «·√” «– «·≈„«„ √»Ê «·Õ”‰ ⁄·Ì »‰ √Õ„œ »‰ ⁄„— «·Õ„«„Ì ° √Œ»—‰« √»Ê «·ﬁ«”„ ⁄»œ «··Â »‰ «·Õ”‰ »‰ ”·Ì„«‰ «·‰Œ«” ° √Œ»—‰« √»Ê »ﬂ— „Õ„œ »‰ Â«—Ê‰ »‰ ‰«›⁄ «· „«— «·»€œ«œÌ ° √Œ»—‰« √»Ê ⁄»œ «··Â „Õ„œ »‰ «·„ Êﬂ· «·„⁄—Ê› »—ÊÌ” ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï «·≈„«„ √»Ì „Õ„œ ⁄»œ «·—Õ„‰ »‰ √Õ„œ »‰ ⁄·Ì «·»€œ«œÌ ° Ê√Œ»—‰Ì √‰Â ﬁ—√ »Â« «·ﬁ—¬‰ ﬂ·Â ⁄·Ï «·≈„«„ «· ﬁÌ „Õ„œ »‰ √Õ„œ «·„’—Ì ° Êﬁ—√ »Â« ⁄·Ï ≈»—«ÂÌ„ »‰ √Õ„œ «·≈”ﬂ‰œ—Ì ° Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï ⁄»œ «··Â »‰ ⁄·Ì «·»€œ«œÌ ° Êﬁ—√ »Â« ⁄·Ï «·√” «– √»Ì «·⁄“ «·ﬁ·«‰”Ì ° Êﬁ—√ »Â« ⁄·Ï √»Ì ⁄·Ì «·Õ”‰ »‰ «·ﬁ«”„ «·Ê«”ÿÌ ° Êﬁ—√ »Â« ⁄·Ï : «·Õ„«„Ì ° Êﬁ—√ »Â« ⁄·Ï «·‰Œ« ” ° Êﬁ—√ »Â« ⁄·Ï «· „«— ° Êﬁ—√ ⁄·Ï —ÊÌ” ° Êﬁ—√ »Â« ⁄·Ï Ì⁄ﬁÊ» . " & vbNewLine
        sanadan = sanadan & "Ê—Ã‹‹«· Ì⁄ﬁÊ» «·–Ì‰ ”„«Â„ √—»⁄… : √»Ê «·„‰–— ”·«„ »‰ ”·Ì„«‰ «·ÿÊÌ· ° Ê‘Â«» »‰ ‘—‰›… ° Ê„ÂœÌ »‰ „Ì„Ê‰ ° Ê√»Ê «·√‘Â» Ã⁄›— »‰ ÕÌ«‰ «·⁄ÿ«—œÌ .ÊﬁÌ· ≈‰ Ì⁄ﬁÊ» ﬁ—√ ⁄·Ï √»Ì ⁄„—Ê »‰ «·⁄·«¡ Êﬁ—√ ”·«„ ⁄·Ï ⁄«’„ Ê√»Ì ⁄„—Ê ° Êﬁ‹‹‹—√ ‘Â«» «·ÃÕœ—Ì Êﬁ—√ ⁄«’„ ⁄·Ï «·Õ”‰ «·»’—Ì Ê⁄·Ï ”·Ì„«‰ »‰ ﬁ … Êﬁ—√ ”·Ì„«‰ ⁄·Ï «»‹‰ ⁄»« ” Êﬁ—√ „ÂœÌ ⁄·Ï ‘⁄Ì» »‰ «·Õ»Õ«» Êﬁ—√ ⁄·Ï √»Ì «·⁄«·Ì… «·—Ì«ÕÌ Êﬁ—√ ⁄·Ï √»Ì Ê“Ìœ Êﬁ—√ √»Ê «·√‘Â» ⁄·Ï √»Ì —Ã«¡ ⁄„—«‰ »‰ „·Õ«‰ «·⁄ÿ«—œÌ Êﬁ—√ ⁄·Ï √»Ì „Ê”‹‹‹Ï «·√‘⁄—Ì Êﬁ—√ ⁄·Ï —”Ê· «··Â ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        
        ElseIf index = 102 Then
        '«·Ê—«ﬁ
        sanadan = "ﬁ«· «·≈„«„ √»Ê «·ŒÌ— „Õ„œ »‰ «·Ã“—Ì ›Ï  Õ»Ì— «· Ì”Ì— : " & vbNewLine
        sanadan = sanadan & "√„« —Ê«Ì… ≈œ—Ì” «·Ê—«ﬁ : ›ÕœÀ‰« »Â« √»Ê Õ›’ ⁄„— »‰ «·Õ”‰ »ﬁ—«¡ Ì ⁄·ÌÂ Ÿ«Â— œ„‘ﬁ ° ⁄‰ ‘ÌŒÂ «·≈„«„ «·ŒÿÌ» √»Ì «·⁄»«” √Õ„œ »‰ ≈»—«ÂÌ„ »‰ ⁄„— «·›«—Ê∆Ì «·‘«›⁄Ì ° ﬁ«· : √Œ»—‰« Ê«·œÌ ° ﬁ«· : √Œ»—‰« √»Ê «·”⁄«œ«  «·√”⁄œ »‰ ”·ÿ«‰ «·Ê«”ÿÌ ° √Œ»—‰« √»Ê «·⁄“ „Õ„œ »‰ «·Õ”Ì‰ «·Ê«”ÿÌ ° √Œ»—‰« √»Ê «·Õ”Ì‰ √Õ„œ »‰ ⁄»œ «··Â »‰ «·Œ÷— «·”Ê”‰Ã—œÌ ° √Œ»—‰« √»Ê «·Õ”‰ „Õ„œ »‰ ⁄»œ «··Â »‰ „Õ„œ »‰ „—… «·ÿÊ”Ì «·„⁄—Ê› »«»‰ √»Ì ⁄„— «·‰ﬁ«‘ ° √Œ»—‰« √»Ê Ì⁄ﬁÊ» ≈”Õ«ﬁ »‰ ≈»—«ÂÌ„ «·Ê—«ﬁ ." & vbNewLine
        sanadan = sanadan & "ﬁ«· «»‰ «·Ã“—Ì : Êﬁ—√  »Â« «·ﬁ—¡«‰ ﬂ·Â ⁄·Ï ﬂ· „‰ «·‘ÌŒÌ‰ √»Ì ⁄»œ «··Â «·Õ‰›Ì ° Ê√»Ì „Õ„œ «·‘«›⁄Ì «·„’—ÌÌ‰ ° Êﬁ—√ ﬂ· „‰Â„« ⁄·Ï √»Ì ⁄»œ «··Â „Õ„œ »‰ √Õ„œ »‰ ⁄»œ «·Œ«·ﬁ «·„’—Ì ° Êﬁ—√ »Â« ⁄·Ï «·ﬂ„«· »‰ ›«—” ° Êﬁ—√ »Â« ⁄·Ï “Ìœ »‰ «·Õ”‰ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·ﬁ«”„ Â»… «··Â »‰ √Õ„œ »‰ «·ÿ»— «·»€œ«œÌ ° Êﬁ—√ »Â« ⁄·Ï √»Ì »ﬂ— „Õ„œ »‰ ⁄·Ì »‰ „Ê”Ï «·ŒÌ«ÿ ° Êﬁ—√ »Â« ⁄·Ï √»Ì «·Õ”Ì‰ «·”Ê”‰Ã—œÌ ° Êﬁ—√ »Â« ⁄·Ï «»‰ √»Ì ⁄„— «·ÿÊ”Ì ° Êﬁ—√ »Â« ⁄·Ï ≈”Õ«ﬁ «·Ê—«ﬁ ° Êﬁ—√ »Â« ⁄·Ï Œ·› ." & vbNewLine
        sanadan = sanadan & "Ê—Ã«· Œ·› : Ê—Ã«· Œ·› ”·Ì„ ’«Õ» Õ„“… ° ÊÌ⁄ﬁÊ» »‰ Œ·Ì›… «·√⁄‘Ï ’«Õ» √»Ì »ﬂ— ° Ê√»Ê “Ìœ ”⁄Ìœ ”⁄Ìœ »‰ √Ê” «·√‰’«—Ì ’«Õ» «·„›÷· «·÷»Ì Ê√»«‰ «·⁄ÿ«— ° Êﬁ—√ √»Ê »ﬂ— ° Ê«·„›÷· ° Ê√»«‰ ⁄·Ï ⁄«’„ . Ê—ÊÏ «·ﬁ—«¡… √Ì÷« ⁄‰ «·ﬂ”«∆Ì Ê⁄‰ ÌÕÌÏ »‰ ¬œ„ ⁄‰ √»Ì »ﬂ— ° Ê«··Â «·„Ê›ﬁ . ﬁ·  : Ê√Œ– ⁄«’„ ⁄‰ √»Ì ⁄»œ «·—Õ„‰ ⁄»œ «··Â »‰ Õ»Ì» «·”·„Ì ° Ê√»Ì „—Ì„ “— »‰ Õ»Ì‘ ° Ê√Œ‹– √»Ê ⁄»œ «·—Õ„‰ ⁄‰ ⁄À„«‰ »‰ ⁄›«‰ ° Ê⁄·Ì »‰ √»Ì ÿ«·» ° Ê√»Ì »‰ ﬂ⁄» ° Ê“Ìœ »‰ À«»  ° Ê⁄»œ «··Â »‰ „”⁄Êœ ° ⁄‰ «·‰»Ì ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        sanadan = sanadan & "Ê√Œ– “— »‰ Õ»Ì‘ ⁄‰ ⁄À„«‰ »‰ ⁄‹›‹‹«‰ ° Ê«»‰ „”⁄Êœ ° ⁄‰ —”Ê· «··Â ° ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ - ° ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -. Ê—Ã«· Õ„“… : Ê—Ã«· Õ„“… Ã„«⁄… „‰Â„ √»Ê „Õ„œ ”·Ì„«‰ »‰ „Â—«‰ «·√⁄„‘ ° Ê„Õ„œ »‰ ⁄»œ «·—Õ„‰ »‰ √»Ì ·Ì·Ï «·ﬁ«÷Ì ° ÊÕ„—«‰ »‰ √⁄Ì‰ ° Ê√»Ê ≈”Õ«ﬁ «·”»Ì⁄Ì ° Ê„‰’Ê— «»‰ «·„⁄ „— ° Ê„€Ì—… »‰ „ﬁ”„ ° ÊÃ⁄›— »‰ „Õ„œ «·’«œﬁ ° Ê€Ì—Â„ . Ê√Œ– «·√⁄„‘ ⁄‰ ÌÕÌÏ »‰ ÊÀ«» ° Ê√Œ– ÌÕÌÏ ⁄‰ Ã„«⁄… „‰ √’Õ«» «»‰ „”⁄Êœ : ⁄·ﬁ„… ° Ê«·√”Êœ Ê⁄»Ìœ »‰ ‰÷·… «·Œ“«⁄Ì ° Ê“— »‰ Õ»Ì‘ ° Ê√»Ì ⁄»œ «·—Õ„‰ «·”·„Ì ° Ê€Ì—Â„ ° ⁄‰ «»‰ „”⁄Êœ ⁄‰ «·‰»Ì ⁄‰ Ã»—Ì· - ⁄·ÌÂ «·”·«„ -  ⁄‰ —» «·⁄“… -  »«—ﬂ Ê  ⁄«·Ï -." & vbNewLine
        Else
        sanadan = "sanada"
    End If

End Function
Function qeraatn(index As Integer) As String

        'adding sanad
        If index = -1 Then
        qeraatn = "»ﬁ—«¡«  √Â· «· Ê”ÿ ( «»‰ ⁄«„— Ê ⁄«’„ Ê «·ﬂ”«∆Ï Ê Œ·› )"
          
        ElseIf index = -2 Then
        qeraatn = "»ﬁ—¡«… «·»’—Ì«‰ ( √»Ê ⁄„—Ê Ê Ì⁄ﬁÊ» ) "
      
        ElseIf index = -3 Then
        qeraatn = "»«·ﬁ—«¡«  «·⁄‘— «·’€—Ï"
      
        ElseIf index = -4 Then
        qeraatn = "»ﬁ—«¡«  √Â· «·’·…"
       
        ElseIf index = -5 Then
        qeraatn = "»«·ﬁ—«¡«  «·”»⁄"
       
        ElseIf index = 1 Then
        qeraatn = "»ﬁ—«¡… «·≈„«„ ‰«›⁄ »—«ÊÌÌÂ"
        
        ElseIf index = 3 Then
        qeraatn = "»ﬁ—«¡… «·≈„«„ √»Ê ⁄„—Ê «·»’—Ï »—«ÊÌÌÂ"
        
        ElseIf index = 4 Then
        qeraatn = "»ﬁ—«¡… «·≈„«„ «»‰ ⁄«„— »—«ÊÌÌÂ"
        
        ElseIf index = 5 Then
        qeraatn = "»ﬁ—«¡… «·≈„«„ ⁄«’„ »—«ÊÌÌÂ"
        
        ElseIf index = 6 Then
        qeraatn = "»ﬁ—«¡… «·≈„«„ Õ„“… »—«ÊÌÌÂ"
        
        ElseIf index = 7 Then
        qeraatn = "»ﬁ—«¡… «·≈„«„ «·ﬂ”«∆Ï »—«ÊÌÌÂ"
        
        ElseIf index = 8 Then
        qeraatn = "»ﬁ—«¡… «·≈„«„ √»Ê Ã⁄›— »—«ÊÌÌÂ"
        
        ElseIf index = 9 Then
        qeraatn = "»ﬁ—«¡… «·≈„«„ Ì⁄ﬁÊ» »—«ÊÌÌÂ"
        
        ElseIf index = 10 Then
        qeraatn = "»ﬁ—«¡… «·≈„«„ Œ·› «·»“«— »—«ÊÌÌÂ"
        
        ElseIf index = 11 Then
        qeraatn = "»—Ê«Ì… Ê—‘ ⁄‰ ‰«›⁄"
        
        ElseIf index = 21 Then
        qeraatn = "»—Ê«Ì… ﬁ‰»· ⁄‰ «»‰ ﬂÀÌ—"
        
        ElseIf index = 31 Then
        qeraatn = "»—Ê«Ì… «·”Ê”Ï ⁄‰ √»Ê ⁄„—Ê «·»’—Ï"
        
        ElseIf index = 41 Then
        qeraatn = "»—«ÊÌ… «»‰ –ﬂÊ«‰ ⁄‰ «»‰ ⁄«„—"
        
        ElseIf index = 51 Then
        qeraatn = "»—Ê«Ì… Õ›’ ⁄‰ ⁄«’„"
        
        ElseIf index = 61 Then
        qeraatn = "»—Ê«Ì… Œ·«œ ⁄‰ Õ„“…"
        
        ElseIf index = 71 Then
        qeraatn = "»—Ê«Ì… √»Ï «·Õ«—À ⁄‰ «·ﬂ”«∆Ï"
        
        ElseIf index = 81 Then
        qeraatn = "»—Ê«Ì… «»‰ Ã„«“ ⁄‰ √»Ï Ã⁄›—"
        
        ElseIf index = 91 Then
        qeraatn = "»—Ê«Ì… —ÊÕ ⁄‰ Ì⁄ﬁÊ»"
        
        ElseIf index = 101 Then
        qeraatn = "»—Ê«Ì… ≈œ—Ì” ⁄‰ Œ·› «·»“«—"
        
        ElseIf index = 12 Then
        qeraatn = "»—Ê«Ì… ﬁ«·Ê‰ ⁄‰ ‰«›⁄"
        
        ElseIf index = 22 Then
        qeraatn = "»—Ê«Ì… «·»“Ï ⁄‰ «»‰ ﬂÀÌ—"
        
        ElseIf index = 32 Then
        qeraatn = "»—Ê«Ì… «·œÊ—Ï ⁄‰ √»Ê ⁄„—Ê «·»’—Ï"
        
        ElseIf index = 42 Then
        qeraatn = "»—Ê«Ì… Â‘«„ ⁄‰ «»‰ ⁄«„—"
        
        ElseIf index = 52 Then
        qeraatn = "»—Ê«Ì… ‘⁄»… ⁄‰ ⁄«’„"
        
        ElseIf index = 62 Then
        qeraatn = "»—Ê«Ì… Œ·› ⁄‰ Õ„“…"
        
        ElseIf index = 72 Then
        qeraatn = "»—Ê«Ì… √»Ê ⁄„—Ê «·œÊ—Ï ⁄‰ «·ﬂ”«∆Ï"
        
        ElseIf index = 82 Then
        qeraatn = "»—Ê«Ì… «»‰ Ê—œ«‰ ⁄‰ √»Ï Ã⁄›—"
        
        ElseIf index = 92 Then
        qeraatn = "»—Ê«Ì… —ÊÌ” ⁄‰ Ì⁄ﬁÊ»"
        
        ElseIf index = 102 Then
        qeraatn = "»—Ê«Ì… «·Ê—«ﬁ ⁄‰ Œ·› «·»“«—"
        Else
        qeraatn = "egaza_content"
    End If

End Function
Public Function rawye(index As Integer) As String

     'adding sanad
        If index = -1 Then
        rawye = "”‰œ ﬁ—«¡«  / √Â· «· Ê”ÿ"
        
        ElseIf index = -2 Then
        rawye = "”‰œ ﬁ—«¡«  / «·»’—Ì«‰"
        
        ElseIf index = -3 Then
        rawye = "”‰œ «·ﬁ—«¡«  «·⁄‘—"
         
        ElseIf index = -4 Then
        rawye = "”‰œ ﬁ—«¡«  √Â· «·’·…"
          
        ElseIf index = -5 Then
        rawye = "”‰œ «·ﬁ—«¡«  «·”»⁄"
          
        ElseIf index = 1 Then
        rawye = "”‰œ ﬁ—«¡… «·≈„«„ / ‰«›⁄"
        
        ElseIf index = 2 Then
        rawye = "”‰œ ﬁ—«¡… «·≈„«„ / «»‰ ﬂÀÌ—"
        
        ElseIf index = 3 Then
        rawye = "”‰œ ﬁ—«¡… «·≈„«„ / √»Ê ⁄„—Ê «·»’—Ï"
        
        ElseIf index = 4 Then
        rawye = "”‰œ ﬁ—«¡… «·≈„«„ / «»‰ ⁄«„—"
        
        ElseIf index = 5 Then
        rawye = "”‰œ ﬁ—«¡… «·≈„«„ / ⁄«’„"
        
        ElseIf index = 6 Then
        rawye = "”‰œ ﬁ—«¡… «·≈„«„ / Õ„“…"
        
        ElseIf index = 7 Then
        rawye = "”‰œ ﬁ—«¡… «·≈„«„ / «·ﬂ”«∆Ï"
        
        ElseIf index = 8 Then
        rawye = "”‰œ ﬁ—«¡… «·≈„«„ / √»Ê Ã⁄›—"
        
        ElseIf index = 9 Then
        rawye = "”‰œ ﬁ—«¡… «·≈„«„ / Ì⁄ﬁÊ»"
        
        ElseIf index = 10 Then
        rawye = "”‰œ ﬁ—«¡… «·≈„«„ / Œ·› «·»“«—"
        
        ElseIf index = 11 Then
        rawye = "”‰œ —Ê«Ì… / Ê—‘"
        
        ElseIf index = 21 Then
        rawye = "”‰œ —Ê«Ì… / ﬁ‰»·"
        
        ElseIf index = 31 Then
        rawye = "”‰œ —Ê«Ì… / «·”Ê”Ï"
        
        ElseIf index = 41 Then
        rawye = "”‰œ —Ê«Ì… / «»‰ –ﬂÊ«‰"
        
        ElseIf index = 51 Then
        rawye = "”‰œ —Ê«Ì… / Õ›’"
        
        ElseIf index = 61 Then
        rawye = "”‰œ —Ê«Ì… / Œ·«œ"
        
        ElseIf index = 71 Then
        rawye = "”‰œ —Ê«Ì… / √»Ï «·Õ«—À"
        
        ElseIf index = 81 Then
        rawye = "”‰œ —Ê«Ì… / «»‰ Ã„«“"
        
        ElseIf index = 91 Then
        rawye = "”‰œ —Ê«Ì… / —ÊÕ"
        
        ElseIf index = 101 Then
        rawye = "”‰œ —Ê«Ì… / ≈œ—Ì”"
        
        ElseIf index = 12 Then
        rawye = "”‰œ —Ê«Ì… / ﬁ«·Ê‰"
        
        ElseIf index = 22 Then
        rawye = "”‰œ —Ê«Ì… / «·»“Ï"
        
        ElseIf index = 32 Then
        rawye = "”‰œ —Ê«Ì… / «·œÊ—Ï"
        
        ElseIf index = 42 Then
        rawye = "”‰œ —Ê«Ì… / Â‘«„"
        
        ElseIf index = 52 Then
        rawye = "”‰œ —Ê«Ì… / ‘⁄»…"
        
        ElseIf index = 62 Then
        rawye = "”‰œ —Ê«Ì… / Œ·›"
        
        ElseIf index = 72 Then
        rawye = "”‰œ —Ê«Ì… / √»Ê ⁄„—Ê"
        
        ElseIf index = 82 Then
        rawye = "”‰œ —Ê«Ì… / «»‰ Ê—œ«‰"
        
        ElseIf index = 92 Then
        rawye = "”‰œ —Ê«Ì… / —ÊÌ”"
        
        ElseIf index = 102 Then
        rawye = "”‰œ —Ê«Ì… / «·Ê—«ﬁ"
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
        get_status = "«Œ »«—«"
    End If
    
    If CheckBox40.Value = True Then
        get_status = "»⁄÷ «·ﬁ—«‰"
    Else
        get_status = "Œ „… ﬂ«„·…"
    End If
       
    If CheckBox41.Value = True Then
        get_status = get_status + " " + "‰Ÿ—« „‰ «·„’Õ›"
    Else
        get_status = get_status + " " + "€Ì»« ⁄‰ ŸÂ— ﬁ·»"
    End If
   
End Function
Public Function get_index() As Integer
 ' set index

    If CheckBox38.Value = True Then
        ' «·”»⁄
        get_index = -5
    End If
    
    If CheckBox6.Value = True Then
        ' «Â· «·’·…
        get_index = -4
    End If
    
    If CheckBox37.Value = True Then
         ' «·⁄‘—
         get_index = -3
    End If
    
    If CheckBox42.Value = True Then
         ' «·»’—Ì«‰
         get_index = -2
    End If
    
    If CheckBox5.Value = True Then
        ' «· Ê”ÿ
        get_index = -1
    End If
    
    If CheckBox7.Value = True Then
        '‰«›⁄
        get_index = 1
    End If
    
    If CheckBox8.Value = True Then
        '«»‰ ﬂÀÌ—
        get_index = 2
    End If
   
    If CheckBox9.Value = True Then
        '«»Ê ⁄„—Ê
        get_index = 3
    End If
   
    If CheckBox10.Value = True Then
       '«»‰ ⁄«„—
        get_index = 4
    End If
     
    If CheckBox11.Value = True Then
       '⁄«’„
        get_index = 5
    End If
     
    If CheckBox12.Value = True Then
       'Õ„“…
        get_index = 6
    End If
     
    If CheckBox13.Value = True Then
       '«·ﬂ”«∆Ï
        get_index = 7
    End If
     
    If CheckBox14.Value = True Then
        '«»Ê Ã⁄›—
         get_index = 8
    End If
   
    If CheckBox15.Value = True Then
       'Ì⁄ﬁÊ»
        get_index = 9
    End If
     
    If CheckBox16.Value = True Then
        'Œ·›
         get_index = 10
    End If
   
    ' set Rowayat
    If CheckBox17.Value = True Then
        'Ê—‘
        get_index = 11
    End If
   
    If CheckBox18.Value = True Then
        'ﬁ«·Ê‰
        get_index = 12
    End If
    
    If CheckBox19.Value = True Then
        'ﬁ‰»·
         get_index = 21
    End If
     
    If CheckBox20.Value = True Then
        '«·»“Ï
         get_index = 22
    End If
     
    If CheckBox21.Value = True Then
        '«·”Ê”Ï
         get_index = 31
    End If
    
    If CheckBox22.Value = True Then
       '«·œÊ—Ï
       get_index = 32
    End If
     
    If CheckBox23.Value = True Then
     '«»‰ –ﬂÊ«‰
     get_index = 41
    End If
    
    If CheckBox24.Value = True Then
      'Â‘«„ ⁄‰ «»‰ ⁄«„—
      get_index = 42
    End If
     
    If CheckBox25.Value = True Then
     'Õ›’
     get_index = 51
    End If
     
    If CheckBox26.Value = True Then
    '‘⁄»…
    get_index = 52
    End If
   
    If CheckBox27.Value = True Then
     'Œ·«œ
     get_index = 61
    End If
     
    If CheckBox28.Value = True Then
      'Œ·›
      get_index = 62
    End If
     
    If CheckBox29.Value = True Then
       '«»Ï «·Õ«—À
       get_index = 71
    End If
     
    If CheckBox30.Value = True Then
        '«·œÊ—Ï ⁄‰ «‰ﬂ”«Ï∆
        get_index = 72
    End If
     
    If CheckBox31.Value = True Then
    '«»‰ Ã„«“
    get_index = 81
    End If
     
    If CheckBox32.Value = True Then
     '«»‰ Ê—œ«‰
     get_index = 82
    End If
     
    If CheckBox33.Value = True Then
      '—ÊÕ
      get_index = 91
    End If
     
    If CheckBox34.Value = True Then
       '—ÊÌ”
       get_index = 92
    End If
     
    If CheckBox35.Value = True Then
       '«œ—Ì”
         get_index = 101
    End If
     
    If CheckBox36.Value = True Then
        '«·Ê—«ﬁ
         get_index = 102
    End If

End Function

Public Function get_special_index(QERAA As String) As Integer
 ' set index

    If InStr(QERAA, "«·”»⁄") > 0 Then
        '«·”»⁄
        get_special_index = -5
    End If
    
    If InStr(QERAA, "«Â· «·’·…") > 0 Then
        ' «Â· «·’·…
        get_special_index = -4
    End If
    
   If InStr(QERAA, "«·⁄‘—") > 0 Then
          ' «·⁄‘—
         get_special_index = -3
    End If
    
   If InStr(QERAA, "«·»’—Ì«‰") > 0 Then
          ' «·»’—Ì«‰
         get_special_index = -2
    End If
    
   If InStr(QERAA, "«· Ê”ÿ") > 0 Then
         ' «· Ê”ÿ
        get_special_index = -1
    End If
    
   If InStr(QERAA, "‰«›⁄") > 0 Then
         '‰«›⁄
        get_special_index = 1
    End If
    
   If InStr(QERAA, "«»‰ ﬂÀÌ—") > 0 Then
         '«»‰ ﬂÀÌ—
        get_special_index = 2
    End If
   
   If InStr(QERAA, "«»Ê ⁄„—Ê") > 0 Then
         '«»Ê ⁄„—Ê
        get_special_index = 3
    End If
   
   If InStr(QERAA, "«»‰ ⁄«„—") > 0 Then
        '«»‰ ⁄«„—
        get_special_index = 4
    End If
     
   If InStr(QERAA, "⁄«’„") > 0 Then
        '⁄«’„
        get_special_index = 5
    End If
     
   If InStr(QERAA, "Õ„“…") > 0 Then
        'Õ„“…
        get_special_index = 6
    End If
     
   If InStr(QERAA, "«·ﬂ”«∆Ï") > 0 Then
        '«·ﬂ”«∆Ï
        get_special_index = 7
    End If
     
   If InStr(QERAA, "«»Ê Ã⁄›—") > 0 Then
         '«»Ê Ã⁄›—
         get_special_index = 8
    End If
   
   If InStr(QERAA, "Ì⁄ﬁÊ»") > 0 Then
        'Ì⁄ﬁÊ»
        get_special_index = 9
    End If
     
   If InStr(QERAA, "Œ·› «·⁄«‘—") > 0 Then
         'Œ·› «·⁄«‘—
         get_special_index = 10
    End If
   
    ' set Rowayat
   If InStr(QERAA, "Ê—‘") > 0 Then
         'Ê—‘
        get_special_index = 11
    End If
   
   If InStr(QERAA, "ﬁ«·Ê‰") > 0 Then
         'ﬁ«·Ê‰
        get_special_index = 12
    End If
    
   If InStr(QERAA, "ﬁ‰»·") > 0 Then
         'ﬁ‰»·
         get_special_index = 21
    End If
     
   If InStr(QERAA, "«·»“Ï") > 0 Then
         '«·»“Ï
         get_special_index = 22
    End If
     
   If InStr(QERAA, "«·”Ê”Ï") > 0 Then
         '«·”Ê”Ï
         get_special_index = 31
    End If
    
   If InStr(QERAA, "«·œÊ—Ì") > 0 Then
        '«·œÊ—Ï
       get_special_index = 32
    End If
     
   If InStr(QERAA, "«»‰ –ﬂÊ«‰") > 0 Then
      '«»‰ –ﬂÊ«‰
     get_special_index = 41
    End If
    
   If InStr(QERAA, "Â‘«„") > 0 Then
       'Â‘«„ ⁄‰ «»‰ ⁄«„—
      get_special_index = 42
    End If
     
   If InStr(QERAA, "Õ›’") > 0 Then
      'Õ›’
     get_special_index = 51
    End If
     
   If InStr(QERAA, "‘⁄»…") > 0 Then
     '‘⁄»…
    get_special_index = 52
    End If
   
   If InStr(QERAA, "Œ·«œ") > 0 Then
      'Œ·«œ
     get_special_index = 61
    End If
     
   If InStr(QERAA, "Œ·›") > 0 Then
       'Œ·›
      get_special_index = 62
    End If
     
   If InStr(QERAA, "«»Ï «·Õ«—À") > 0 Then
        '«»Ï «·Õ«—À
       get_special_index = 71
    End If
     
   If InStr(QERAA, "«·œÊ—Ï ⁄‰ «‰ﬂ”«Ï∆") > 0 Then
         '«·œÊ—Ï ⁄‰ «‰ﬂ”«Ï∆
        get_special_index = 72
    End If
     
   If InStr(QERAA, "«»‰ Ã„«“") > 0 Then
     '«»‰ Ã„«“
    get_special_index = 81
    End If
     
   If InStr(QERAA, "«»‰ Ê—œ«‰") > 0 Then
      '«»‰ Ê—œ«‰
     get_special_index = 82
    End If
     
   If InStr(QERAA, "—ÊÕ") > 0 Then
       '—ÊÕ
      get_special_index = 91
    End If
     
   If InStr(QERAA, "—ÊÌ”") > 0 Then
        '—ÊÌ”
       get_special_index = 92
    End If
     
   If InStr(QERAA, "«œ—Ì”") > 0 Then
        '«œ—Ì”
         get_special_index = 101
    End If
     
   If InStr(QERAA, "«·Ê—«ﬁ") > 0 Then
         '«·Ê—«ﬁ
         get_special_index = 102
    End If

End Function

Public Function get_tareq() As String
    get_tareq = " „‰ ÿ—Ìﬁ "
    If CheckBox3.Value = True Then
     
        If CheckBox14.Value = True Or CheckBox15.Value = True Or CheckBox16.Value = True Or CheckBox31.Value = True Or CheckBox32.Value = True Or CheckBox33.Value = True Or CheckBox34.Value = True Or CheckBox35.Value = True Or CheckBox36.Value = True Then
            get_tareq = get_tareq + "«·œ—…"
        Else
            get_tareq = get_tareq + "«·‘«ÿ»Ì…"
        End If
        
        If CheckBox37.Value = True Or CheckBox42.Value = True Or CheckBox6.Value = True Or CheckBox5.Value = True Then
            get_tareq = " „‰ ÿ—Ìﬁ «·‘«ÿ»Ì… Ê «·œ—…"
        End If
        
     End If
     
     If CheckBox4.Value = True And CheckBox3.Value = True Then
         get_tareq = get_tareq + " Ê «·ÿÌ»…"
     ElseIf CheckBox4.Value = True Then
         get_tareq = get_tareq + "«·ÿÌ»…"
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
      
        If (substrings(3 + (k * 4))) = "ÿ«·»" Then
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
   
   
   ejaza_id = InputBox("√œŒ· „⁄—› «·≈Ã«“…")
   strSQL = "Select * from EJAZA where ID = " & ejaza_id
   Set rst = db.OpenRecordset(strSQL)
       
   If rst.RecordCount > 0 Then
      NAME = rst.Fields("STUDENT_NAME")
      INFO = rst.Fields("STUDENT_INFO")
      QERAA = rst.Fields("QERAA")
      TAREk = rst.Fields("TAREQ")
      STATE = rst.Fields("STATE")
      If (rst.Fields("STUDENT_GENDER") = "ÿ«·»") Then
        GENDER = True
      Else
        GENDER = False
      End If
      
      MsgBox (NAME & vbNewLine & INFO & vbNewLine & rst.Fields("STUDENT_GENDER") & vbNewLine & QERAA & vbNewLine & TAREk & vbNewLine & STATE)
   Else
      MsgBox ("„⁄—› «·≈Ã«“… €Ì— „ÊÃÊœ")
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
    If InStr(STATE, "«Œ »«—«") > 0 Then
        status = "«Œ »«—«"
    ElseIf InStr(STATE, "»⁄÷") > 0 Then
        status = "»⁄÷ «·ﬁ—¬‰"
    ElseIf InStr(STATE, "Œ „…") > 0 Then
        status = "Œ „… ﬂ«„·…"
    Else
        status = STATE
    End If
    
    If InStr(STATE, "€Ì»«") > 0 Then
        status = status + " " + "€Ì»« ⁄‰ ŸÂ— ﬁ·»"
    ElseIf InStr(STATE, "‰Ÿ—«") > 0 Then
        status = status + " " + "‰Ÿ—« „‰ «·„’Õ›"
    Else
         status = STATE
    End If
    
    
    If InStr(TAREk, "’€—Ï") > 0 And index < 0 And index > -5 Then
        TAREQ = " „‰ ÿ—Ìﬁ «·‘«ÿ»Ì… Ê«·œ—…"
    ElseIf InStr(TAREk, "’€—Ï") > 0 And index > 80 And index < 110 Then
        TAREQ = " „‰ ÿ—Ìﬁ «·œ—…"
    ElseIf InStr(TAREk, "’€—Ï") > 0 And index > 7 And index < 11 Then
        TAREQ = " „‰ ÿ—Ìﬁ «·œ—…"
    ElseIf InStr(TAREk, "’€—Ï") > 0 Then
        TAREQ = " „‰ ÿ—Ìﬁ «·‘«ÿ»Ì…"
    ElseIf InStr(TAREk, "ﬂ»—Ï") > 0 Then
        TAREQ = " „‰ ÿ—Ìﬁ «·ÿÌ»…"
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
    

   MsgBox (" „ ")
End Sub

Private Sub OptionButton3_Click()
 TextBox3.text = "„ﬁ—∆ Ê„⁄·„ «·ﬁ—¬‰ «·ﬂ—Ì„ Ê«· ÃÊÌœ"
 
 
End Sub

Private Sub OptionButton4_Click()
 TextBox3.text = "„ﬁ—∆… Ê„⁄·„… «·ﬁ—¬‰ «·ﬂ—Ì„ Ê«· ÃÊÌœ"
End Sub

Private Sub UserForm_Click()

End Sub

