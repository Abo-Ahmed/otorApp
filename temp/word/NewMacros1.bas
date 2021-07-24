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
        .text = "áöãõÓúÊóÍóŞøöåóÇ ÇáãõÌóÇÒ"
        .Replacement.text = "áöãõÓúÊóÍóŞÊåóÇ ÇáãõÌóÇÒÉ"
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
        .text = "áöãõÓúÊóÍóŞøöåóÇ ÇáãõÌóÇÒ"
        .Replacement.text = "áöãõÓúÊóÍóŞÊåóÇ ÇáãõÌóÇÒÉ"
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
        .text = "ÇÓã ÇáØÇáÈ åäÇ"
        .Replacement.text = "ÇÓã ÇáØÇáÈÉ åäÇ"
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
            "äİÚ Çááå Èå æÚóİóÇ Úóäúåõ æóÚóäú æóÇáöÏóíúåö æóÔõíõæÎöå æóÇáúãõÓúáöãöíäó"
        .Replacement.text = _
            "äİÚ Çááå ÈåÇ æÚóİóÇ ÚóäúåÇ æóÚóäú æóÇáöÏóíúåÇ æóÔõíõæÎöåÇ æóÇáúãõÓúáöãöíäó"
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
        .text = "ÇáÚóãöíŞö ÇáØÇáöÈõ ÇáãõÌóÇÒõ /"
        .Replacement.text = "ÇáÚóãöíŞö ÇáØÇáöÈÉ ÇáãõÌóÇÒÉ /"
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
        .text = "áóŞóÏú ŞóÑóÃó Úóáóíøó ÇáŞõÑúÂäó ÇáßóÑöíãó"
        .Replacement.text = "áóŞóÏú ŞóÑóÃóÊ Úóáóíøó ÇáŞõÑúÂäó ÇáßóÑöíãó"
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
            "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåõ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåö ßõáøó ÇáÅØúãöÆúäóÇäö"
        .Replacement.text = _
            "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåÇ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåÇ ßõáøó ÇáÅØúãöÆúäóÇäö"
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
            "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåõ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåö ßõáøó ÇáÅØúãöÆúäóÇäö"
        .Replacement.text = _
            "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåÇ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåÇ ßõáøó ÇáÅØúãöÆúäóÇäö"
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
            "æó ŞÏ ØóáóÈó ãöäóì ÇáÅöÌóÇÒóÉó æó ßöÊóÇÈóÉó ÇáÓøóäóÏö İóÃóÌóÒúÊõåõ ÈöÇáŞöÑóÇÁóÉö"
        .Replacement.text = _
            "æó ŞÏ ØóáóÈÊ ãöäóì ÇáÅöÌóÇÒóÉó æó ßöÊóÇÈóÉó ÇáÓøóäóÏö İóÃóÌóÒúÊõåÇ ÈöÇáŞöÑóÇÁóÉö"
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
            "áößóæúäöåö ÃóåúáÇğ áĞóáößó æóÃóĞöäúÊõ áóåõ Ãóäú íóŞúÑóÃó æíõŞúÑöÆ æóíõÚóáøöãõ æóíõÌöíÒõ ÛóíúÑóåõ ÈöãóÇ ŞóÑóÃó Úóáóíøó İöí Ãóíøö ãóßóÇä"
        .Replacement.text = _
            "áößóæúäöåÇ ÃóåúáÇğ áĞóáößó æóÃóĞöäúÊõ áóåÇ Ãóäú ÊŞúÑóÃó æÊŞúÑöÆ æó ÊÚóáøöãõ æó ÊÌöíÒõ ÛóíúÑóåÇ ÈöãóÇ ŞóÑóÃóÊ Úóáóíøó İöí Ãóíøö ãóßóÇä"
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
            " Íóáøò æó İøóì Ãóíøö ŞõØúÑ äóÒóáó ÈöÔóÑúØö ÇáúÃóãóÇäóÉö æó ÇáÕøöíóÇäóÉö æóÇáúãõØóÇáóÚóÉö æóÃóáóÇ íóŞõæáó ÅöáóÇ ÈöãóÇ íóÚúáóãõ İóÅöäú ÈóÏøóáó ÃóæúÛóíøóÑó Ãæó ÖóíøóÚó ÇáŞõÑúÂäó"
        .Replacement.text = _
            " ÍóáøòÊ æó İøóì Ãóíøö ŞõØúÑ äóÒóáóÊ ÈöÔóÑúØö ÇáúÃóãóÇäóÉö æó ÇáÕøöíóÇäóÉö æóÇáúãõØóÇáóÚóÉö æóÃóáóÇ ÊŞõæáó ÅöáóÇ ÈöãóÇ ÊÚúáóãõ İóÅöäú ÈóÏøóáóÊ Ãóæú ÛóíøóÑóÊ Ãæó ÖóíøóÚóÊ ÇáŞõÑúÂäó"
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
           "æóŞóÚó İöí ÇááøóÍúäö"
        .Replacement.text = _
           "æŞÚÊ İì ÇááÍä"
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
            "æóŞóÏú ØóáóÈó ãöäøöì ãóÚúÑöİóÉó ÅöÓúäóÇÏöì İöí ÇáŞõÑúÂäö ÇáßóÑöíãö İóÃóÌóÈúÊõåõ æóÃóÎúÈóÑúÊõåõ"
        .Replacement.text = _
            "æóŞóÏú ØóáóÈóÊ ãöäøöì ãóÚúÑöİóÉó ÅöÓúäóÇÏöì İöí ÇáŞõÑúÂäö ÇáßóÑöíãö İóÃóÌóÈúÊõåÇ æóÃóÎúÈóÑúÊõåÇ"
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
        .text = "ÇáÔíÎ ÇáãÌÇÒ / "
        .Replacement.text = "ÇáÔíÎÉ ÇáãÌÇÒÉ / "
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
        .text = "ÇáÔíÎ ÇáãÌÇÒ / "
        .Replacement.text = "ÇáÔíÎÉ ÇáãÌÇÒÉ / "
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
        .text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒó "
        .Replacement.text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒÉ"
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
        .text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒÉ"
        .Replacement.text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒÉ "
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
            "áöíóÚúÑöİó ŞóÏúÑó ãóÇ æóÕóáó Åöáóíúåö æó ÃõÛúÏöŞó Úóáóíúåö ãóäú åóĞöåö ÇáäøöÚúãóÉö ÇáÚóÙöíãóÉö æó ÇáãöäøóÉö ÇáÌóÓöíãóÉö æó áöíõÚóáøöã"
        .Replacement.text = _
            "áöÊÚúÑöİó ŞóÏúÑó ãóÇ æóÕóáóÊ Åöáóíúåö æó ÃõÛúÏöŞ ÚóáóíúåÇ ãóäú åóĞöåö ÇáäøöÚúãóÉö ÇáÚóÙöíãóÉö æó ÇáãöäøóÉö ÇáÌóÓöíãóÉö æó áöÊÚóáøöã"
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
            "ÎóÇİöÖğÇ ÌóäóÇÍóåõ áößõáøö ãóäú ÃõÊóÇåõ æóáóÇ íóŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåõ æóíóÊúÑõß ÇáÌöÏøó"
        .Replacement.text = _
            "ÎóÇİöÖÉ ÌóäóÇÍóåÇ áößõáøö ãóäú ÃõÊóÇåÇ æóáóÇ ÊŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåÇ æóÊÊúÑõß ÇáÌöÏøó"
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
            "ÎóÇİöÖğÇ ÌóäóÇÍóåõ áößõáøö ãóäú ÃõÊóÇåõ æóáóÇ íóŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåõ æóíóÊúÑõß ÇáÌöÏøó"
        .Replacement.text = _
            "ÎóÇİöÖÉ ÌóäóÇÍóåÇ áößõáøö ãóäú ÃõÊóÇåÇ æóáóÇ ÊŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåÇ æóÊÊúÑõß ÇáÌöÏøó"
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
        .text = "æáíóÒöÏå ÇáÚöáúãó ãóÍóÇÓöäó"
        .Replacement.text = "æáíóÒöÏåÇ ÇáÚöáúãó ãóÍóÇÓöäó"
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
        .text = "æó Åöäøöì ŞóÏú ÃóÌóÒúÊõßó ÃóíåÇ ÇáØøóÇáöÈõ"
        .Replacement.text = "æó Åöäøöì ŞóÏú ÃóÌóÒúÊõßö ÃóíÊåÇ ÇáØøóÇáöÈÉ"
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
            "İóÍóÇİöÙõ Ãöíå ÇáãõÌóÇÒõ Úóáóì ãóÇ ÃóÏøóíúÊõåõ áóßó ÌóÚóáóßó"
        .Replacement.text = _
            "İóÍóÇİöÙö ÃöíÊåÇ ÇáãõÌóÇÒÉ Úóáóì ãóÇ ÃóÏøóíúÊõåõ áóßó ÌóÚóáóßö"
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
            " æóÃõæÕöíåö ÃóáóÇ íóäúÓóÇäöí æóÃóåúáöí æóĞóÑøöíøóÊöí ãöäú ÕóÇáöÍö ÏóÚóæóÇÊöåö İöí ÎóáóæóÇÊöåö æÌóáóæóÇÊöåö æóÃóäú íóĞúßõÑóäöí ÚöäúÏó ÑóÈøöå."
        .Replacement.text = _
            " æóÃõæÕöíåÇ ÃóáóÇ ÊäúÓóÇäöí æóÃóåúáöí æóĞóÑøöíøóÊöí ãöäú ÕóÇáöÍö ÏóÚóæóÇÊöåÇ İöí ÎóáóæóÇÊöåÇ æÌóáóæóÇÊöåÇ æóÃóäú ÊĞúßõÑóäöí ÚöäúÏó ÑóÈøöåÇ."
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
            "æŞÏ ŞÑÃ ÇáØÇáÈ ÃíÖÇ Úáì"
        .Replacement.text = _
            "æŞÏ ŞÑÃÊ ÇáØÇáÈÉ ÃíÖÇ Úáì"
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
        .text = "áöãõÓúÊóÍóŞÊåóÇ ÇáãõÌóÇÒÉ"
        .Replacement.text = "áöãõÓúÊóÍóŞøöåóÇ ÇáãõÌóÇÒ"
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
        .text = "áöãõÓúÊóÍóŞÊåóÇ ÇáãõÌóÇÒÉ"
        .Replacement.text = "áöãõÓúÊóÍóŞøöåóÇ ÇáãõÌóÇÒ"
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
        .text = "ÇÓã ÇáØÇáÈÉ åäÇ"
        .Replacement.text = "ÇÓã ÇáØÇáÈ åäÇ"
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
        .text = "äİÚ Çááå ÈåÇ æÚóİóÇ ÚóäúåÇ æóÚóäú æóÇáöÏóíúåÇ æóÔõíõæÎöåÇ æóÇáúãõÓúáöãöíäó"
        .Replacement.text = "äİÚ Çááå Èå æÚóİóÇ Úóäúåõ æóÚóäú æóÇáöÏóíúåö æóÔõíõæÎöå æóÇáúãõÓúáöãöíäó"
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
        .text = "ÇáÚóãöíŞö ÇáØÇáöÈÉ ÇáãõÌóÇÒÉ /"
        .Replacement.text = "ÇáÚóãöíŞö ÇáØÇáöÈõ ÇáãõÌóÇÒõ /"
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
        .text = "áóŞóÏú ŞóÑóÃóÊ Úóáóíøó ÇáŞõÑúÂäó ÇáßóÑöíãó"
        .Replacement.text = "áóŞóÏú ŞóÑóÃó Úóáóíøó ÇáŞõÑúÂäó ÇáßóÑöíãó"
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
        .text = "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåÇ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåÇ ßõáøó ÇáÅØúãöÆúäóÇäö"
        .Replacement.text = "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåõ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåö ßõáøó ÇáÅØúãöÆúäóÇäö"
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
        .text = "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåÇ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåÇ ßõáøó ÇáÅØúãöÆúäóÇäö"
        .Replacement.text = "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåõ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåö ßõáøó ÇáÅØúãöÆúäóÇäö"
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
        .text = "æó ŞÏ ØóáóÈÊ ãöäóì ÇáÅöÌóÇÒóÉó æó ßöÊóÇÈóÉó ÇáÓøóäóÏö İóÃóÌóÒúÊõåÇ ÈöÇáŞöÑóÇÁóÉö"
        .Replacement.text = "æó ŞÏ ØóáóÈó ãöäóì ÇáÅöÌóÇÒóÉó æó ßöÊóÇÈóÉó ÇáÓøóäóÏö İóÃóÌóÒúÊõåõ ÈöÇáŞöÑóÇÁóÉö"
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
        .text = "áößóæúäöåÇ ÃóåúáÇğ áĞóáößó æóÃóĞöäúÊõ áóåÇ Ãóäú ÊŞúÑóÃó æÊŞúÑöÆ æó ÊÚóáøöãõ æó ÊÌöíÒõ ÛóíúÑóåÇ ÈöãóÇ ŞóÑóÃóÊ Úóáóíøó İöí Ãóíøö ãóßóÇä"
        .Replacement.text = "áößóæúäöåö ÃóåúáÇğ áĞóáößó æóÃóĞöäúÊõ áóåõ Ãóäú íóŞúÑóÃó æíõŞúÑöÆ æóíõÚóáøöãõ æóíõÌöíÒõ ÛóíúÑóåõ ÈöãóÇ ŞóÑóÃó Úóáóíøó İöí Ãóíøö ãóßóÇä"
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
        .text = " ÍóáøòÊ æó İøóì Ãóíøö ŞõØúÑ äóÒóáóÊ ÈöÔóÑúØö ÇáúÃóãóÇäóÉö æó ÇáÕøöíóÇäóÉö æóÇáúãõØóÇáóÚóÉö æóÃóáóÇ ÊŞõæáó ÅöáóÇ ÈöãóÇ ÊÚúáóãõ İóÅöäú ÈóÏøóáóÊ Ãóæú ÛóíøóÑóÊ Ãæó ÖóíøóÚóÊ ÇáŞõÑúÂäó"
         .Replacement.text = " Íóáøò æó İøóì Ãóíøö ŞõØúÑ äóÒóáó ÈöÔóÑúØö ÇáúÃóãóÇäóÉö æó ÇáÕøöíóÇäóÉö æóÇáúãõØóÇáóÚóÉö æóÃóáóÇ íóŞõæáó ÅöáóÇ ÈöãóÇ íóÚúáóãõ İóÅöäú ÈóÏøóáó ÃóæúÛóíøóÑó Ãæó ÖóíøóÚó ÇáŞõÑúÂäó"
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
        .text = "æŞÚÊ İì ÇááÍä"
        .Replacement.text = "æóŞóÚó İöí ÇááøóÍúäö"
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
        .text = "æóŞóÏú ØóáóÈóÊ ãöäøöì ãóÚúÑöİóÉó ÅöÓúäóÇÏöì İöí ÇáŞõÑúÂäö ÇáßóÑöíãö İóÃóÌóÈúÊõåÇ æóÃóÎúÈóÑúÊõåÇ"
        .Replacement.text = "æóŞóÏú ØóáóÈó ãöäøöì ãóÚúÑöİóÉó ÅöÓúäóÇÏöì İöí ÇáŞõÑúÂäö ÇáßóÑöíãö İóÃóÌóÈúÊõåõ æóÃóÎúÈóÑúÊõåõ"
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
        .text = "ÇáÔíÎÉ ÇáãÌÇÒÉ / "
        .Replacement.text = "ÇáÔíÎ ÇáãÌÇÒ / "
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
        .text = "ÇáÔíÎÉ ÇáãÌÇÒÉ / "
        .Replacement.text = "ÇáÔíÎ ÇáãÌÇÒ / "
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
        .text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒÉ"
        .Replacement.text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒó "
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
        .text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒÉ "
        .Replacement.text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒÉ"
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
        .text = "áöÊÚúÑöİó ŞóÏúÑó ãóÇ æóÕóáóÊ Åöáóíúåö æó ÃõÛúÏöŞ ÚóáóíúåÇ ãóäú åóĞöåö ÇáäøöÚúãóÉö ÇáÚóÙöíãóÉö æó ÇáãöäøóÉö ÇáÌóÓöíãóÉö æó áöÊÚóáøöã"
        .Replacement.text = "áöíóÚúÑöİó ŞóÏúÑó ãóÇ æóÕóáó Åöáóíúåö æó ÃõÛúÏöŞó Úóáóíúåö ãóäú åóĞöåö ÇáäøöÚúãóÉö ÇáÚóÙöíãóÉö æó ÇáãöäøóÉö ÇáÌóÓöíãóÉö æó áöíõÚóáøöã"
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
        .text = "ÎóÇİöÖÉ ÌóäóÇÍóåÇ áößõáøö ãóäú ÃõÊóÇåÇ æóáóÇ ÊŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåÇ æóÊÊúÑõß ÇáÌöÏøó"
        .Replacement.text = "ÎóÇİöÖğÇ ÌóäóÇÍóåõ áößõáøö ãóäú ÃõÊóÇåõ æóáóÇ íóŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåõ æóíóÊúÑõß ÇáÌöÏøó"
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
        .text = "ÎóÇİöÖÉ ÌóäóÇÍóåÇ áößõáøö ãóäú ÃõÊóÇåÇ æóáóÇ ÊŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåÇ æóÊÊúÑõß ÇáÌöÏøó"
        .Replacement.text = "ÎóÇİöÖğÇ ÌóäóÇÍóåõ áößõáøö ãóäú ÃõÊóÇåõ æóáóÇ íóŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåõ æóíóÊúÑõß ÇáÌöÏøó"
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
        .text = "æáíóÒöÏåÇ ÇáÚöáúãó ãóÍóÇÓöäó"
        .Replacement.text = "æáíóÒöÏå ÇáÚöáúãó ãóÍóÇÓöäó"
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
        .text = "æó Åöäøöì ŞóÏú ÃóÌóÒúÊõßö ÃóíÊåÇ ÇáØøóÇáöÈÉ"
        .Replacement.text = "æó Åöäøöì ŞóÏú ÃóÌóÒúÊõßó ÃóíåÇ ÇáØøóÇáöÈõ"
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
        .text = "İóÍóÇİöÙö ÃöíÊåÇ ÇáãõÌóÇÒÉ Úóáóì ãóÇ ÃóÏøóíúÊõåõ áóßó ÌóÚóáóßö"
        .Replacement.text = "İóÍóÇİöÙõ Ãöíå ÇáãõÌóÇÒõ Úóáóì ãóÇ ÃóÏøóíúÊõåõ áóßó ÌóÚóáóßó"
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
        .text = " æóÃõæÕöíåÇ ÃóáóÇ ÊäúÓóÇäöí æóÃóåúáöí æóĞóÑøöíøóÊöí ãöäú ÕóÇáöÍö ÏóÚóæóÇÊöåÇ İöí ÎóáóæóÇÊöåÇ æÌóáóæóÇÊöåÇ æóÃóäú ÊĞúßõÑóäöí ÚöäúÏó ÑóÈøöåÇ."
        .Replacement.text = " æóÃõæÕöíåö ÃóáóÇ íóäúÓóÇäöí æóÃóåúáöí æóĞóÑøöíøóÊöí ãöäú ÕóÇáöÍö ÏóÚóæóÇÊöåö İöí ÎóáóæóÇÊöåö æÌóáóæóÇÊöåö æóÃóäú íóĞúßõÑóäöí ÚöäúÏó ÑóÈøöå."
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
         .text = "æŞÏ ŞÑÃÊ ÇáØÇáÈÉ ÃíÖÇ Úáì"
         .Replacement.text = "æŞÏ ŞÑÃ ÇáØÇáÈ ÃíÖÇ Úáì"
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
        .text = "ÛíÈÇ"
        .Replacement.text = "ÚÑÖÇğ ãä ÇáãÕÍİ"
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
        .text = "æ ßİì ÈÇááå ÔåíÏÇğ"
        .Replacement.text = _
            "æŞÏ ŞÑÃ ÇáØÇáÈ ÃíÖÇ Úáì: 1- .........................................."
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
 
    ActiveDocument.SaveAs2 FileName:=FileName + " - ÚÑÖÇ" + ".docx", FileFormat:= _
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
        .text = "áöãõÓúÊóÍóŞøöåóÇ ÇáãõÌóÇÒ"
        .Replacement.text = "áöãõÓúÊóÍóŞÊåóÇ ÇáãõÌóÇÒÉ"
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
        .text = "áöãõÓúÊóÍóŞøöåóÇ ÇáãõÌóÇÒ"
        .Replacement.text = "áöãõÓúÊóÍóŞÊåóÇ ÇáãõÌóÇÒÉ"
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
        .text = "ÇÓã ÇáØÇáÈ åäÇ"
        .Replacement.text = "ÇÓã ÇáØÇáÈÉ åäÇ"
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
            "äİÚ Çááå Èå æÚóİóÇ Úóäúåõ æóÚóäú æóÇáöÏóíúåö æóÔõíõæÎöå æóÇáúãõÓúáöãöíäó"
        .Replacement.text = _
            "äİÚ Çááå ÈåÇ æÚóİóÇ ÚóäúåÇ æóÚóäú æóÇáöÏóíúåÇ æóÔõíõæÎöåÇ æóÇáúãõÓúáöãöíäó"
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
        .text = "ÇáÚóãöíŞö ÇáØÇáöÈõ ÇáãõÌóÇÒõ /"
        .Replacement.text = "ÇáÚóãöíŞö ÇáØÇáöÈÉ ÇáãõÌóÇÒÉ /"
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
        .text = "áóŞóÏú ŞóÑóÃó Úóáóíøó ÇáŞõÑúÂäó ÇáßóÑöíãó"
        .Replacement.text = "áóŞóÏú ŞóÑóÃóÊ Úóáóíøó ÇáŞõÑúÂäó ÇáßóÑöíãó"
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
            "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåõ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåö ßõáøó ÇáÅØúãöÆúäóÇäö"
        .Replacement.text = _
            "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåÇ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåÇ ßõáøó ÇáÅØúãöÆúäóÇäö"
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
            "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåõ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåö ßõáøó ÇáÅØúãöÆúäóÇäö"
        .Replacement.text = _
            "æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåÇ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåÇ ßõáøó ÇáÅØúãöÆúäóÇäö"
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
            "æó ŞÏ ØóáóÈó ãöäóì ÇáÅöÌóÇÒóÉó æó ßöÊóÇÈóÉó ÇáÓøóäóÏö İóÃóÌóÒúÊõåõ ÈöÇáŞöÑóÇÁóÉö"
        .Replacement.text = _
            "æó ŞÏ ØóáóÈÊ ãöäóì ÇáÅöÌóÇÒóÉó æó ßöÊóÇÈóÉó ÇáÓøóäóÏö İóÃóÌóÒúÊõåÇ ÈöÇáŞöÑóÇÁóÉö"
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
            "áößóæúäöåö ÃóåúáÇğ áĞóáößó æóÃóĞöäúÊõ áóåõ Ãóäú íóŞúÑóÃó æíõŞúÑöÆ æóíõÚóáøöãõ æóíõÌöíÒõ ÛóíúÑóåõ ÈöãóÇ ŞóÑóÃó Úóáóíøó İöí Ãóíøö ãóßóÇä"
        .Replacement.text = _
            "áößóæúäöåÇ ÃóåúáÇğ áĞóáößó æóÃóĞöäúÊõ áóåÇ Ãóäú ÊŞúÑóÃó æÊŞúÑöÆ æó ÊÚóáøöãõ æó ÊÌöíÒõ ÛóíúÑóåÇ ÈöãóÇ ŞóÑóÃóÊ Úóáóíøó İöí Ãóíøö ãóßóÇä"
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
            " Íóáøò æó İøóì Ãóíøö ŞõØúÑ äóÒóáó ÈöÔóÑúØö ÇáúÃóãóÇäóÉö æó ÇáÕøöíóÇäóÉö æóÇáúãõØóÇáóÚóÉö æóÃóáóÇ íóŞõæáó ÅöáóÇ ÈöãóÇ íóÚúáóãõ İóÅöäú ÈóÏøóáó ÃóæúÛóíøóÑó Ãæó ÖóíøóÚó ÇáŞõÑúÂäó"
        .Replacement.text = _
            " ÍóáøòÊ æó İøóì Ãóíøö ŞõØúÑ äóÒóáóÊ ÈöÔóÑúØö ÇáúÃóãóÇäóÉö æó ÇáÕøöíóÇäóÉö æóÇáúãõØóÇáóÚóÉö æóÃóáóÇ ÊŞõæáó ÅöáóÇ ÈöãóÇ ÊÚúáóãõ İóÅöäú ÈóÏøóáóÊ Ãóæú ÛóíøóÑóÊ Ãæó ÖóíøóÚóÊ ÇáŞõÑúÂäó"
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
           "æóŞóÚó İöí ÇááøóÍúäö"
        .Replacement.text = _
           "æŞÚÊ İì ÇááÍä"
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
            "æóŞóÏú ØóáóÈó ãöäøöì ãóÚúÑöİóÉó ÅöÓúäóÇÏöì İöí ÇáŞõÑúÂäö ÇáßóÑöíãö İóÃóÌóÈúÊõåõ æóÃóÎúÈóÑúÊõåõ"
        .Replacement.text = _
            "æóŞóÏú ØóáóÈóÊ ãöäøöì ãóÚúÑöİóÉó ÅöÓúäóÇÏöì İöí ÇáŞõÑúÂäö ÇáßóÑöíãö İóÃóÌóÈúÊõåÇ æóÃóÎúÈóÑúÊõåÇ"
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
        .text = "ÇáÔíÎ ÇáãÌÇÒ / "
        .Replacement.text = "ÇáÔíÎÉ ÇáãÌÇÒÉ / "
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
        .text = "ÇáÔíÎ ÇáãÌÇÒ / "
        .Replacement.text = "ÇáÔíÎÉ ÇáãÌÇÒÉ / "
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
        .text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒó "
        .Replacement.text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒÉ"
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
        .text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒÉ"
        .Replacement.text = "åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒÉ "
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
            "áöíóÚúÑöİó ŞóÏúÑó ãóÇ æóÕóáó Åöáóíúåö æó ÃõÛúÏöŞó Úóáóíúåö ãóäú åóĞöåö ÇáäøöÚúãóÉö ÇáÚóÙöíãóÉö æó ÇáãöäøóÉö ÇáÌóÓöíãóÉö æó áöíõÚóáøöã"
        .Replacement.text = _
            "áöÊÚúÑöİó ŞóÏúÑó ãóÇ æóÕóáóÊ Åöáóíúåö æó ÃõÛúÏöŞ ÚóáóíúåÇ ãóäú åóĞöåö ÇáäøöÚúãóÉö ÇáÚóÙöíãóÉö æó ÇáãöäøóÉö ÇáÌóÓöíãóÉö æó áöÊÚóáøöã"
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
            "ÎóÇİöÖğÇ ÌóäóÇÍóåõ áößõáøö ãóäú ÃõÊóÇåõ æóáóÇ íóŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåõ æóíóÊúÑõß ÇáÌöÏøó"
        .Replacement.text = _
            "ÎóÇİöÖÉ ÌóäóÇÍóåÇ áößõáøö ãóäú ÃõÊóÇåÇ æóáóÇ ÊŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåÇ æóÊÊúÑõß ÇáÌöÏøó"
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
            "ÎóÇİöÖğÇ ÌóäóÇÍóåõ áößõáøö ãóäú ÃõÊóÇåõ æóáóÇ íóŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåõ æóíóÊúÑõß ÇáÌöÏøó"
        .Replacement.text = _
            "ÎóÇİöÖÉ ÌóäóÇÍóåÇ áößõáøö ãóäú ÃõÊóÇåÇ æóáóÇ ÊŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåÇ æóÊÊúÑõß ÇáÌöÏøó"
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
        .text = "æáíóÒöÏå ÇáÚöáúãó ãóÍóÇÓöäó"
        .Replacement.text = "æáíóÒöÏåÇ ÇáÚöáúãó ãóÍóÇÓöäó"
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
        .text = "æó Åöäøöì ŞóÏú ÃóÌóÒúÊõßó ÃóíåÇ ÇáØøóÇáöÈõ"
        .Replacement.text = "æó Åöäøöì ŞóÏú ÃóÌóÒúÊõßö ÃóíÊåÇ ÇáØøóÇáöÈÉ"
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
            "İóÍóÇİöÙõ Ãöíå ÇáãõÌóÇÒõ Úóáóì ãóÇ ÃóÏøóíúÊõåõ áóßó ÌóÚóáóßó"
        .Replacement.text = _
            "İóÍóÇİöÙö ÃöíÊåÇ ÇáãõÌóÇÒÉ Úóáóì ãóÇ ÃóÏøóíúÊõåõ áóßó ÌóÚóáóßö"
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
            " æóÃõæÕöíåö ÃóáóÇ íóäúÓóÇäöí æóÃóåúáöí æóĞóÑøöíøóÊöí ãöäú ÕóÇáöÍö ÏóÚóæóÇÊöåö İöí ÎóáóæóÇÊöåö æÌóáóæóÇÊöåö æóÃóäú íóĞúßõÑóäöí ÚöäúÏó ÑóÈøöå."
        .Replacement.text = _
            " æóÃõæÕöíåÇ ÃóáóÇ ÊäúÓóÇäöí æóÃóåúáöí æóĞóÑøöíøóÊöí ãöäú ÕóÇáöÍö ÏóÚóæóÇÊöåÇ İöí ÎóáóæóÇÊöåÇ æÌóáóæóÇÊöåÇ æóÃóäú ÊĞúßõÑóäöí ÚöäúÏó ÑóÈøöåÇ."
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
            "æŞÏ ŞÑÃ ÇáØÇáÈ ÃíÖÇ Úáì"
        .Replacement.text = _
            "æŞÏ ŞÑÃÊ ÇáØÇáÈÉ ÃíÖÇ Úáì"
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
    
      
    ActiveDocument.SaveAs2 FileName:=FileName + " - äÓÇÁ" + ".docx", FileFormat:= _
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
        .text = "ÛíÈÇ"
        .Replacement.text = "ÚÑÖÇğ ãä ÇáãÕÍİ"
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
        .text = "æ ßİì ÈÇááå ÔåíÏÇğ"
        .Replacement.text = _
            "æŞÏ ŞÑÃ ÇáØÇáÈ ÃíÖÇ Úáì: 1- .........................................."
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
    
     ActiveDocument.SaveAs2 FileName:=FileName + " - ÚÑÖÇ" + ".docx", FileFormat:= _
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
        .Replacement.text = "ÇáÔíÎ ÇáãÌíÒ"
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
