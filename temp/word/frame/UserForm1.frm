VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ÇÌÇÒÉ"
   ClientHeight    =   6828
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13872
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db_object As DAO.Database
Dim rst_object As DAO.Recordset
Dim wdApp As Word.Application
Dim database_path As String

Private Sub CommandButton9_Click()
    replace_text original_text:="ó", new_text:=""
    replace_text original_text:="ğ", new_text:=""
    replace_text original_text:="õ", new_text:=""
    replace_text original_text:="ñ", new_text:=""
    replace_text original_text:="ö", new_text:=""
    replace_text original_text:="ò", new_text:=""
    replace_text original_text:="ú", new_text:=""
    replace_text original_text:="ø", new_text:=""

End Sub

Private Sub UserForm_Initialize()
   database_path = "E:\sanad\otor_be.accdb"
   'On Error GoTo rapid:
   Set db_object = OpenDatabase(database_path)
   Set rst_object = db_object.OpenRecordset("Select ID , SHEIKH_NAME , DEGREE from  [ORDER] WHERE STATE NOT IN ( 'NEXT' ) ORDER BY DEGREE")
   Do Until rst_object.EOF
     ComboBox1.AddItem (rst_object.Fields("ID") & "-" & rst_object.Fields("SHEIKH_NAME"))
     rst_object.MoveNext
   Loop
    
rapid:
    ' make numbers arabic
    Options.ArabicNumeral = wdNumeralHindi
   
End Sub

Public Function replace_text(original_text As String, new_text As String)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = original_text
        .Replacement.text = new_text
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
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
End Function

Private Sub set_names()
    replace_text original_text:="sheikh_name", new_text:=TextBox1.text
    replace_text original_text:="sheikh_info", new_text:=TextBox3.text
    replace_text original_text:="student_name", new_text:=TextBox2.text
    replace_text original_text:="student_info", new_text:=TextBox4.text
End Sub
Private Sub set_sheikh_type(sheikh_type)
    If sheikh_type = False Then
        replace_text original_text:="mogez", new_text:="ÇáÔíÎÉ"
        replace_text original_text:="İíŞæá ÇáÔíÎÉ", new_text:="İÊŞæá ÇáÔíÎÉ"
    Else
        replace_text original_text:="mogez", new_text:="ÇáÔíÎ"
    End If
End Sub
    
Private Sub set_student_type(student_type As Boolean)
   
' set student type
    If student_type = False Then
            
            replace_text original_text:="áöãõÓúÊóÍóŞøöåóÇ ÇáãõÌóÇÒ", new_text:="áöãõÓúÊóÍóŞÊåóÇ ÇáãõÌóÇÒÉ"
            replace_text original_text:="ÇÓã ÇáØÇáÈ åäÇ", new_text:="ÇÓã ÇáØÇáÈÉ åäÇ"
            replace_text original_text:="äİÚ Çááå Èå æÚóİóÇ Úóäúåõ æóÚóäú æóÇáöÏóíúåö æóÔõíõæÎöå æóÇáúãõÓúáöãöíäó", new_text:="äİÚ Çááå ÈåÇ æÚóİóÇ ÚóäúåÇ æóÚóäú æóÇáöÏóíúåÇ æóÔõíõæÎöåÇ æóÇáúãõÓúáöãöíäó"
            replace_text original_text:="ÇáÚóãöíŞö ÇáØÇáöÈõ ÇáãõÌóÇÒõ /", new_text:="ÇáÚóãöíŞö ÇáØÇáöÈÉ ÇáãõÌóÇÒÉ /"
            replace_text original_text:="áóŞóÏú ŞóÑóÃó Úóáóíøó ÇáŞõÑúÂäó ÇáßóÑöíãó", new_text:="áóŞóÏú ŞóÑóÃóÊ Úóáóíøó ÇáŞõÑúÂäó ÇáßóÑöíãó"
            replace_text original_text:="æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåõ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåö ßõáøó ÇáÅØúãöÆúäóÇä", new_text:="æóÈóÚúÏ Ãóäú ÚóáöãóÊõ ãöäúåÇ ÇáÏøöÑóÇíóÉö æóÇáÅöÊúŞóÇäö æó ÃØúãóÃúäóäúÊõ Åöáóì ŞöÑóÇÁóÊöåÇ ßõáøó ÇáÇØúãöÆúäóÇäö"
            replace_text original_text:="æó ŞÏ ØóáóÈó ãöäóì ÇáÅöÌóÇÒóÉó æó ßöÊóÇÈóÉó ÇáÓøóäóÏö İóÃóÌóÒúÊõåõ ÈöÇáŞöÑóÇÁóÉ", new_text:="æó ŞÏ ØóáóÈÊ ãöäóì ÇáÅöÌóÇÒóÉó æó ßöÊóÇÈóÉó ÇáÓøóäóÏö İóÃóÌóÒúÊõåÇ ÈöÇáŞöÑóÇÁóÉ"
            replace_text original_text:="áößóæúäöåö ÃóåúáÇğ áĞóáößó æóÃóĞöäúÊõ áóåõ Ãóäú íóŞúÑóÃó æíõŞúÑöÆ æóíõÚóáøöãõ æóíõÌöíÒõ ÛóíúÑóåõ ÈöãóÇ ŞóÑóÃó Úóáóíøó İöí Ãóíøö ãóßóÇä", new_text:="áößóæúäöåÇ ÃóåúáÇğ áĞóáößó æóÃóĞöäúÊõ áóåÇ Ãóäú ÊŞúÑóÃó æÊŞúÑöÆ æó ÊÚóáøöãõ æó ÊÌöíÒõ ÛóíúÑóåÇ ÈöãóÇ ŞóÑóÃóÊ Úóáóíøó İöí Ãóíøö ãóßóÇä"
            replace_text original_text:="Íóáøò æó İøóì Ãóíøö ŞõØúÑ äóÒóáó ÈöÔóÑúØö ÇáúÃóãóÇäóÉö æó ÇáÕøöíóÇäóÉö æóÇáúãõØóÇáóÚóÉö æóÃóáóÇ íóŞõæáó ÅöáóÇ ÈöãóÇ íóÚúáóãõ İóÅöäú ÈóÏøóáó ÃóæúÛóíøóÑó Ãæó ÖóíøóÚó ÇáŞõÑúÂä", new_text:="ÍóáøòÊ æó İøóì Ãóíøö ŞõØúÑ äóÒóáóÊ ÈöÔóÑúØö ÇáúÃóãóÇäóÉö æó ÇáÕøöíóÇäóÉö æóÇáúãõØóÇáóÚóÉö æóÃóáóÇ ÊŞõæáó ÅöáóÇ ÈöãóÇ ÊÚúáóãõ İóÅöäú ÈóÏøóáóÊ Ãóæú ÛóíøóÑóÊ Ãæó ÖóíøóÚóÊ ÇáŞõÑúÂäó"
            replace_text original_text:="æóŞóÚó İöí ÇááøóÍúä", new_text:="æŞÚÊ İì ÇááÍä"
            replace_text original_text:="æóŞóÏú ØóáóÈó ãöäøöì ãóÚúÑöİóÉó ÅöÓúäóÇÏöì İöí ÇáŞõÑúÂäö ÇáßóÑöíãö İóÃóÌóÈúÊõåõ æóÃóÎúÈóÑúÊõå", new_text:="æóŞóÏú ØóáóÈóÊ ãöäøöì ãóÚúÑöİóÉó ÅöÓúäóÇÏöì İöí ÇáŞõÑúÂäö ÇáßóÑöíãö İóÃóÌóÈúÊõåÇ æóÃóÎúÈóÑúÊõåÇ"
            replace_text original_text:="ÇáÔíÎ ÇáãÌÇÒ / ", new_text:="ÇáÔíÎÉ ÇáãÌÇÒÉ / "
            replace_text original_text:="åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒó", new_text:="åóĞóÇ æóÃõæÕöí äóİúÓöí æó ÇáãõÌóÇÒÉ"
            replace_text original_text:="áöíóÚúÑöİó ŞóÏúÑó ãóÇ æóÕóáó Åöáóíúåö æó ÃõÛúÏöŞó Úóáóíúåö ãóäú åóĞöåö ÇáäøöÚúãóÉö ÇáÚóÙöíãóÉö æó ÇáãöäøóÉö ÇáÌóÓöíãóÉö æó áöíõÚóáøöã", new_text:="áöÊÚúÑöİó ŞóÏúÑó ãóÇ æóÕóáóÊ Åöáóíúåö æó ÃõÛúÏöŞ ÚóáóíúåÇ ãóäú åóĞöåö ÇáäøöÚúãóÉö ÇáÚóÙöíãóÉö æó ÇáãöäøóÉö ÇáÌóÓöíãóÉö æó áöÊÚóáøöã"
            replace_text original_text:="ÎóÇİöÖğÇ ÌóäóÇÍóåõ áößõáøö ãóäú ÃõÊóÇåõ æóáóÇ íóŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåõ æóíóÊúÑõß ÇáÌöÏøó", new_text:="ÎóÇİöÖÉ ÌóäóÇÍóåÇ áößõáøö ãóäú ÃõÊóÇåÇ æóáóÇ ÊŞúÊóÕóÑ Úóáóì ãóÇ ÚöäúÏóåÇ æóÊÊúÑõß ÇáÌöÏøó"
            replace_text original_text:="æáíóÒöÏå ÇáÚöáúãó ãóÍóÇÓöäó", new_text:="æáíóÒöÏåÇ ÇáÚöáúãó ãóÍóÇÓöäó"
            replace_text original_text:="æó Åöäøöì ŞóÏú ÃóÌóÒúÊõßó ÃóíåÇ ÇáØøóÇáöÈ", new_text:="æó Åöäøöì ŞóÏú ÃóÌóÒúÊõßö ÃóíÊåÇ ÇáØøóÇáöÈÉ"
            replace_text original_text:="İóÍóÇİöÙõ Ãöíå ÇáãõÌóÇÒõ Úóáóì ãóÇ ÃóÏøóíúÊõåõ áóßó ÌóÚóáóß", new_text:="İóÍóÇİöÙí ÃöíÊåÇ ÇáãõÌóÇÒÉ Úóáóì ãóÇ ÃóÏøóíúÊõåõ áóßó ÌóÚóáóß"
            replace_text original_text:=" æóÃõæÕöíåö ÃóáóÇ íóäúÓóÇäöí æóÃóåúáöí æóĞóÑøöíøóÊöí ãöäú ÕóÇáöÍö ÏóÚóæóÇÊöåö İöí ÎóáóæóÇÊöåö æÌóáóæóÇÊöåö æóÃóäú íóĞúßõÑóäöí ÚöäúÏó ÑóÈøöå.", new_text:=" æóÃõæÕöíåÇ ÃóáóÇ ÊäúÓóÇäöí æóÃóåúáöí æóĞóÑøöíøóÊöí ãöäú ÕóÇáöÍö ÏóÚóæóÇÊöåÇ İöí ÎóáóæóÇÊöåÇ æÌóáóæóÇÊöåÇ æóÃóäú ÊĞúßõÑóäöí ÚöäúÏó ÑóÈøöåÇ."
                 
    End If

End Sub

Private Sub update_sheikh_student(sheikh_type As Boolean, student_type As Boolean)
        set_names
        set_sheikh_type sheikh_type:=sheikh_type
        set_student_type student_type:=student_type
End Sub
Private Function set_qeraat(STATE As String, qeraat As String, rawy As String)
    replace_text original_text:="egaza_content", new_text:=qeraat
    replace_text original_text:="rawy", new_text:=rawy
    replace_text original_text:="egaza_state", new_text:=STATE
End Function
Private Sub set_snada(sanada)
 
   Dim index As Integer
   index = 0
   Do While index < 20
        On Error GoTo wow:
            replace_text original_text:="sanada", new_text:=Left(sanada, 200) & "sanada"
            index = index + 1
            sanada = Replace(sanada, Left(sanada, 200), "")
Loop
wow:
           replace_text original_text:="sanada", new_text:=""
End Sub
Function sanadan(index As Integer) As String
   On Error GoTo runner:
    sanadan = db_object.OpenRecordset("Select DETAILS from SANAD WHERE QERAAT = '" & isnad(index) & "'").Fields("DETAILS")
runner:
End Function
Function qeraatn(index As Integer) As String

        'adding sanad
        qeraatn = "egaza_content"
        If index = -1 Then qeraatn = "ÈŞÑÇÁÇÊ Ãåá ÇáÊæÓØ ( ÇÈä ÚÇãÑ æ ÚÇÕã æ ÇáßÓÇÆì æ Îáİ )"
        If index = -2 Then qeraatn = "ÈŞÑÁÇÉ ÇáÈÕÑíÇä ( ÃÈæ ÚãÑæ æ íÚŞæÈ ) "
        If index = -3 Then qeraatn = "ÈÇáŞÑÇÁÇÊ ÇáÚÔÑ ÇáÕÛÑì"
        If index = -4 Then qeraatn = "ÈŞÑÇÁÇÊ Ãåá ÇáÕáÉ"
        If index = -5 Then qeraatn = "ÈÇáŞÑÇÁÇÊ ÇáÓÈÚ"
        If index = 1 Then qeraatn = "ÈŞÑÇÁÉ ÇáÅãÇã äÇİÚ ÈÑÇæííå"
        If index = 3 Then qeraatn = "ÈŞÑÇÁÉ ÇáÅãÇã ÃÈæ ÚãÑæ ÇáÈÕÑì ÈÑÇæííå"
        If index = 4 Then qeraatn = "ÈŞÑÇÁÉ ÇáÅãÇã ÇÈä ÚÇãÑ ÈÑÇæííå"
        If index = 5 Then qeraatn = "ÈŞÑÇÁÉ ÇáÅãÇã ÚÇÕã ÈÑÇæííå"
        If index = 6 Then qeraatn = "ÈŞÑÇÁÉ ÇáÅãÇã ÍãÒÉ ÈÑÇæííå"
        If index = 7 Then qeraatn = "ÈŞÑÇÁÉ ÇáÅãÇã ÇáßÓÇÆì ÈÑÇæííå"
        If index = 8 Then qeraatn = "ÈŞÑÇÁÉ ÇáÅãÇã ÃÈæ ÌÚİÑ ÈÑÇæííå"
        If index = 9 Then qeraatn = "ÈŞÑÇÁÉ ÇáÅãÇã íÚŞæÈ ÈÑÇæííå"
        If index = 10 Then qeraatn = "ÈŞÑÇÁÉ ÇáÅãÇã Îáİ ÇáÈÒÇÑ ÈÑÇæííå"
        If index = 11 Then qeraatn = "ÈÑæÇíÉ æÑÔ Úä äÇİÚ"
        If index = 21 Then qeraatn = "ÈÑæÇíÉ ŞäÈá Úä ÇÈä ßËíÑ"
        If index = 31 Then qeraatn = "ÈÑæÇíÉ ÇáÓæÓì Úä ÃÈæ ÚãÑæ ÇáÈÕÑì"
        If index = 41 Then qeraatn = "ÈÑÇæíÉ ÇÈä ĞßæÇä Úä ÇÈä ÚÇãÑ"
        If index = 51 Then qeraatn = "ÈÑæÇíÉ ÍİÕ Úä ÚÇÕã"
        If index = 61 Then qeraatn = "ÈÑæÇíÉ ÎáÇÏ Úä ÍãÒÉ"
        If index = 71 Then qeraatn = "ÈÑæÇíÉ ÃÈì ÇáÍÇÑË Úä ÇáßÓÇÆì"
        If index = 81 Then qeraatn = "ÈÑæÇíÉ ÇÈä ÌãÇÒ Úä ÃÈì ÌÚİÑ"
        If index = 91 Then qeraatn = "ÈÑæÇíÉ ÑæÍ Úä íÚŞæÈ"
        If index = 101 Then qeraatn = "ÈÑæÇíÉ ÅÏÑíÓ Úä Îáİ ÇáÈÒÇÑ"
        If index = 12 Then qeraatn = "ÈÑæÇíÉ ŞÇáæä Úä äÇİÚ"
        If index = 22 Then qeraatn = "ÈÑæÇíÉ ÇáÈÒì Úä ÇÈä ßËíÑ"
        If index = 32 Then qeraatn = "ÈÑæÇíÉ ÇáÏæÑì Úä ÃÈæ ÚãÑæ ÇáÈÕÑì"
        If index = 42 Then qeraatn = "ÈÑæÇíÉ åÔÇã Úä ÇÈä ÚÇãÑ"
        If index = 52 Then qeraatn = "ÈÑæÇíÉ ÔÚÈÉ Úä ÚÇÕã"
        If index = 62 Then qeraatn = "ÈÑæÇíÉ Îáİ Úä ÍãÒÉ"
        If index = 72 Then qeraatn = "ÈÑæÇíÉ ÃÈæ ÚãÑæ ÇáÏæÑì Úä ÇáßÓÇÆì"
        If index = 82 Then qeraatn = "ÈÑæÇíÉ ÇÈä æÑÏÇä Úä ÃÈì ÌÚİÑ"
        If index = 92 Then qeraatn = "ÈÑæÇíÉ ÑæíÓ Úä íÚŞæÈ"
        If index = 102 Then qeraatn = "ÈÑæÇíÉ ÇáæÑÇŞ Úä Îáİ ÇáÈÒÇÑ"

End Function
Function isnad(index As Integer) As String

        isnad = "egaza_content"
        'adding sanad
        If index = -1 Then isnad = "Çåá ÇáÊæÓØ"
        If index = -2 Then isnad = "ÇáÈÕÑíÇä"
        If index = -3 Then isnad = "ÇáÚÔÑ"
        If index = -4 Then isnad = "Çåá ÇáÕáÉ"
        If index = -5 Then isnad = "ÇáÓÈÚ"
        If index = 1 Then isnad = "äÇİÚ"
        If index = 2 Then isnad = "ÇÈä ßËíÑ"
        If index = 3 Then isnad = "ÇÈæ ÚãÑæ ÇáÈÕÑì"
        If index = 4 Then isnad = "ÇÈä ÚÇãÑ"
        If index = 5 Then isnad = "ÚÇÕã"
        If index = 6 Then isnad = "ÍãÒÉ"
        If index = 7 Then isnad = "ÇáßÓÇÆì"
        If index = 8 Then isnad = "ÇÈæ ÌÚİÑ"
        If index = 9 Then isnad = "íÚŞæÈ"
        If index = 10 Then isnad = "Îáİ ÇáÈÒÇÑ"
        If index = 11 Then isnad = "æÑÔ"
        If index = 21 Then isnad = "ŞäÈá"
        If index = 31 Then isnad = "ÇáÓæÓì"
        If index = 41 Then isnad = "ÇÈä ĞßæÇä"
        If index = 51 Then isnad = "ÍİÕ"
        If index = 61 Then isnad = "ÎáÇÏ"
        If index = 71 Then isnad = "ÇÈì ÇáÍÇÑË"
        If index = 81 Then isnad = "ÇÈä ÌãÇÒ"
        If index = 91 Then isnad = "ÑæÍ"
        If index = 101 Then isnad = "ÅÏÑíÓ"
        If index = 12 Then isnad = "ŞÇáæä"
        If index = 22 Then isnad = "ÇáÈÒì"
        If index = 32 Then isnad = "ÇáÏæÑì"
        If index = 42 Then isnad = "åÔÇã"
        If index = 52 Then isnad = "ÔÚÈÉ"
        If index = 62 Then isnad = "Îáİ"
        If index = 72 Then isnad = "ÇÈæ ÚãÑæ ÇáÏæÑì"
        If index = 82 Then isnad = "ÇÈä æÑÏÇä"
        If index = 92 Then isnad = "ÑæíÓ"
        If index = 102 Then isnad = "ÇáæÑÇŞ"

End Function

Public Function rawye(index As Integer) As String

        rawye = "rawy"
     'adding sanad
        If index = -1 Then rawye = "ÓäÏ ŞÑÇÁÇÊ / Ãåá ÇáÊæÓØ"
        If index = -2 Then rawye = "ÓäÏ ŞÑÇÁÇÊ / ÇáÈÕÑíÇä"
        If index = -3 Then rawye = "ÓäÏ ÇáŞÑÇÁÇÊ ÇáÚÔÑ"
        If index = -4 Then rawye = "ÓäÏ ŞÑÇÁÇÊ Ãåá ÇáÕáÉ"
        If index = -5 Then rawye = "ÓäÏ ÇáŞÑÇÁÇÊ ÇáÓÈÚ"
        If index = 1 Then rawye = "ÓäÏ ŞÑÇÁÉ ÇáÅãÇã / äÇİÚ"
        If index = 2 Then rawye = "ÓäÏ ŞÑÇÁÉ ÇáÅãÇã / ÇÈä ßËíÑ"
        If index = 3 Then rawye = "ÓäÏ ŞÑÇÁÉ ÇáÅãÇã / ÃÈæ ÚãÑæ ÇáÈÕÑì"
        If index = 4 Then rawye = "ÓäÏ ŞÑÇÁÉ ÇáÅãÇã / ÇÈä ÚÇãÑ"
        If index = 5 Then rawye = "ÓäÏ ŞÑÇÁÉ ÇáÅãÇã / ÚÇÕã"
        If index = 6 Then rawye = "ÓäÏ ŞÑÇÁÉ ÇáÅãÇã / ÍãÒÉ"
        If index = 7 Then rawye = "ÓäÏ ŞÑÇÁÉ ÇáÅãÇã / ÇáßÓÇÆì"
        If index = 8 Then rawye = "ÓäÏ ŞÑÇÁÉ ÇáÅãÇã / ÃÈæ ÌÚİÑ"
        If index = 9 Then rawye = "ÓäÏ ŞÑÇÁÉ ÇáÅãÇã / íÚŞæÈ"
        If index = 10 Then rawye = "ÓäÏ ŞÑÇÁÉ ÇáÅãÇã / Îáİ ÇáÈÒÇÑ"
        If index = 11 Then rawye = "ÓäÏ ÑæÇíÉ / æÑÔ"
        If index = 21 Then rawye = "ÓäÏ ÑæÇíÉ / ŞäÈá"
        If index = 31 Then rawye = "ÓäÏ ÑæÇíÉ / ÇáÓæÓì"
        If index = 41 Then rawye = "ÓäÏ ÑæÇíÉ / ÇÈä ĞßæÇä"
        If index = 51 Then rawye = "ÓäÏ ÑæÇíÉ / ÍİÕ"
        If index = 61 Then rawye = "ÓäÏ ÑæÇíÉ / ÎáÇÏ"
        If index = 71 Then rawye = "ÓäÏ ÑæÇíÉ / ÃÈì ÇáÍÇÑË"
        If index = 81 Then rawye = "ÓäÏ ÑæÇíÉ / ÇÈä ÌãÇÒ"
        If index = 91 Then rawye = "ÓäÏ ÑæÇíÉ / ÑæÍ"
        If index = 101 Then rawye = "ÓäÏ ÑæÇíÉ / ÅÏÑíÓ"
        If index = 12 Then rawye = "ÓäÏ ÑæÇíÉ / ŞÇáæä"
        If index = 22 Then rawye = "ÓäÏ ÑæÇíÉ / ÇáÈÒì"
        If index = 32 Then rawye = "ÓäÏ ÑæÇíÉ / ÇáÏæÑì"
        If index = 42 Then rawye = "ÓäÏ ÑæÇíÉ / åÔÇã"
        If index = 52 Then rawye = "ÓäÏ ÑæÇíÉ / ÔÚÈÉ"
        If index = 62 Then rawye = "ÓäÏ ÑæÇíÉ / Îáİ"
        If index = 72 Then rawye = "ÓäÏ ÑæÇíÉ / ÃÈæ ÚãÑæ"
        If index = 82 Then rawye = "ÓäÏ ÑæÇíÉ / ÇÈä æÑÏÇä"
        If index = 92 Then rawye = "ÓäÏ ÑæÇíÉ / ÑæíÓ"
        If index = 102 Then rawye = "ÓäÏ ÑæÇíÉ / ÇáæÑÇŞ"

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
        get_status = "ÇÎÊÈÇÑÇ"
    End If
    
    If CheckBox40.Value = True Then
        get_status = "ÈÚÖ ÇáŞÑÇä"
    Else
        get_status = "ÎÊãÉ ßÇãáÉ"
    End If
       
    If CheckBox41.Value = True Then
        get_status = get_status + " " + "äÙÑÇ ãä ÇáãÕÍİ"
    Else
        get_status = get_status + " " + "ÛíÈÇ Úä ÙåÑ ŞáÈ"
    End If
   
End Function
Public Function get_index() As Integer
 ' set index

    If CheckBox38.Value = True Then
        ' ÇáÓÈÚ
        get_index = -5
    End If
    
    If CheckBox6.Value = True Then
        ' Çåá ÇáÕáÉ
        get_index = -4
    End If
    
    If CheckBox37.Value = True Then
         ' ÇáÚÔÑ
         get_index = -3
    End If
    
    If CheckBox42.Value = True Then
         ' ÇáÈÕÑíÇä
         get_index = -2
    End If
    
    If CheckBox5.Value = True Then
        ' ÇáÊæÓØ
        get_index = -1
    End If
    
    If CheckBox7.Value = True Then
        'äÇİÚ
        get_index = 1
    End If
    
    If CheckBox8.Value = True Then
        'ÇÈä ßËíÑ
        get_index = 2
    End If
   
    If CheckBox9.Value = True Then
        'ÇÈæ ÚãÑæ
        get_index = 3
    End If
   
    If CheckBox10.Value = True Then
       'ÇÈä ÚÇãÑ
        get_index = 4
    End If
     
    If CheckBox11.Value = True Then
       'ÚÇÕã
        get_index = 5
    End If
     
    If CheckBox12.Value = True Then
       'ÍãÒÉ
        get_index = 6
    End If
     
    If CheckBox13.Value = True Then
       'ÇáßÓÇÆì
        get_index = 7
    End If
     
    If CheckBox14.Value = True Then
        'ÇÈæ ÌÚİÑ
         get_index = 8
    End If
   
    If CheckBox15.Value = True Then
       'íÚŞæÈ
        get_index = 9
    End If
     
    If CheckBox16.Value = True Then
        'Îáİ
         get_index = 10
    End If
   
    ' set Rowayat
    If CheckBox17.Value = True Then
        'æÑÔ
        get_index = 11
    End If
   
    If CheckBox18.Value = True Then
        'ŞÇáæä
        get_index = 12
    End If
    
    If CheckBox19.Value = True Then
        'ŞäÈá
         get_index = 21
    End If
     
    If CheckBox20.Value = True Then
        'ÇáÈÒì
         get_index = 22
    End If
     
    If CheckBox21.Value = True Then
        'ÇáÓæÓì
         get_index = 31
    End If
    
    If CheckBox22.Value = True Then
       'ÇáÏæÑì
       get_index = 32
    End If
     
    If CheckBox23.Value = True Then
     'ÇÈä ĞßæÇä
     get_index = 41
    End If
    
    If CheckBox24.Value = True Then
      'åÔÇã Úä ÇÈä ÚÇãÑ
      get_index = 42
    End If
     
    If CheckBox25.Value = True Then
     'ÍİÕ
     get_index = 51
    End If
     
    If CheckBox26.Value = True Then
    'ÔÚÈÉ
    get_index = 52
    End If
   
    If CheckBox27.Value = True Then
     'ÎáÇÏ
     get_index = 61
    End If
     
    If CheckBox28.Value = True Then
      'Îáİ
      get_index = 62
    End If
     
    If CheckBox29.Value = True Then
       'ÇÈì ÇáÍÇÑË
       get_index = 71
    End If
     
    If CheckBox30.Value = True Then
        'ÇáÏæÑì Úä ÇäßÓÇìÆ
        get_index = 72
    End If
     
    If CheckBox31.Value = True Then
    'ÇÈä ÌãÇÒ
    get_index = 81
    End If
     
    If CheckBox32.Value = True Then
     'ÇÈä æÑÏÇä
     get_index = 82
    End If
     
    If CheckBox33.Value = True Then
      'ÑæÍ
      get_index = 91
    End If
     
    If CheckBox34.Value = True Then
       'ÑæíÓ
       get_index = 92
    End If
     
    If CheckBox35.Value = True Then
       'ÇÏÑíÓ
         get_index = 101
    End If
     
    If CheckBox36.Value = True Then
        'ÇáæÑÇŞ
         get_index = 102
    End If

End Function

Public Function get_special_index(QERAA As String) As Integer
 ' set index

    If InStr(QERAA, "ÇáÓÈÚ") > 0 Then
        'ÇáÓÈÚ
        get_special_index = -5
    End If
    
    If InStr(QERAA, "Çåá ÇáÕáÉ") > 0 Then
        ' Çåá ÇáÕáÉ
        get_special_index = -4
    End If
    
   If InStr(QERAA, "ÇáÚÔÑ") > 0 Then
          ' ÇáÚÔÑ
         get_special_index = -3
    End If
    
   If InStr(QERAA, "ÇáÈÕÑíÇä") > 0 Then
          ' ÇáÈÕÑíÇä
         get_special_index = -2
    End If
    
   If InStr(QERAA, "ÇáÊæÓØ") > 0 Then
         ' ÇáÊæÓØ
        get_special_index = -1
    End If
    
   If InStr(QERAA, "äÇİÚ") > 0 Then
         'äÇİÚ
        get_special_index = 1
    End If
    
   If InStr(QERAA, "ßËíÑ") > 0 Then
         'ÇÈä ßËíÑ
        get_special_index = 2
    End If
   
   If InStr(QERAA, "ÇÈæ ÚãÑæ") > 0 Then
         'ÇÈæ ÚãÑæ
        get_special_index = 3
    End If
   
   If InStr(QERAA, "ÇÈä ÚÇãÑ") > 0 Then
        'ÇÈä ÚÇãÑ
        get_special_index = 4
    End If
     
   If InStr(QERAA, "ÚÇÕã") > 0 Then
        'ÚÇÕã
        get_special_index = 5
    End If
     
   If InStr(QERAA, "ÍãÒÉ") > 0 Then
        'ÍãÒÉ
        get_special_index = 6
    End If
     
   If InStr(QERAA, "ÇáßÓÇÆì") > 0 Then
        'ÇáßÓÇÆì
        get_special_index = 7
    End If
     
   If InStr(QERAA, "ÇÈæ ÌÚİÑ") > 0 Then
         'ÇÈæ ÌÚİÑ
         get_special_index = 8
    End If
   
   If InStr(QERAA, "íÚŞæÈ") > 0 Then
        'íÚŞæÈ
        get_special_index = 9
    End If
     
   If InStr(QERAA, "Îáİ ÇáÚÇÔÑ") > 0 Then
         'Îáİ ÇáÚÇÔÑ
         get_special_index = 10
    End If
   
    ' set Rowayat
   If InStr(QERAA, "æÑÔ") > 0 Then
         'æÑÔ
        get_special_index = 11
    End If
   
   If InStr(QERAA, "ŞÇáæä") > 0 Then
         'ŞÇáæä
        get_special_index = 12
    End If
    
   If InStr(QERAA, "ŞäÈá") > 0 Then
         'ŞäÈá
         get_special_index = 21
    End If
     
   If InStr(QERAA, "ÇáÈÒì") > 0 Then
         'ÇáÈÒì
         get_special_index = 22
    End If
     
   If InStr(QERAA, "ÇáÓæÓì") > 0 Then
         'ÇáÓæÓì
         get_special_index = 31
    End If
    
   If InStr(QERAA, "ÇáÏæÑí") > 0 Then
        'ÇáÏæÑì
       get_special_index = 32
    End If
     
   If InStr(QERAA, "ÇÈä ĞßæÇä") > 0 Then
      'ÇÈä ĞßæÇä
     get_special_index = 41
    End If
    
   If InStr(QERAA, "åÔÇã") > 0 Then
       'åÔÇã Úä ÇÈä ÚÇãÑ
      get_special_index = 42
    End If
     
   If InStr(QERAA, "ÍİÕ") > 0 Then
      'ÍİÕ
     get_special_index = 51
    End If
     
   If InStr(QERAA, "ÔÚÈÉ") > 0 Then
     'ÔÚÈÉ
    get_special_index = 52
    End If
   
   If InStr(QERAA, "ÎáÇÏ") > 0 Then
      'ÎáÇÏ
     get_special_index = 61
    End If
     
   If InStr(QERAA, "Îáİ") > 0 Then
       'Îáİ
      get_special_index = 62
    End If
     
   If InStr(QERAA, "ÇÈì ÇáÍÇÑË") > 0 Then
        'ÇÈì ÇáÍÇÑË
       get_special_index = 71
    End If
     
   If InStr(QERAA, "ÇáÏæÑì Úä ÇäßÓÇìÆ") > 0 Then
         'ÇáÏæÑì Úä ÇäßÓÇìÆ
        get_special_index = 72
    End If
     
   If InStr(QERAA, "ÇÈä ÌãÇÒ") > 0 Then
     'ÇÈä ÌãÇÒ
    get_special_index = 81
    End If
     
   If InStr(QERAA, "ÇÈä æÑÏÇä") > 0 Then
      'ÇÈä æÑÏÇä
     get_special_index = 82
    End If
     
   If InStr(QERAA, "ÑæÍ") > 0 Then
       'ÑæÍ
      get_special_index = 91
    End If
     
   If InStr(QERAA, "ÑæíÓ") > 0 Then
        'ÑæíÓ
       get_special_index = 92
    End If
     
   If InStr(QERAA, "ÇÏÑíÓ") > 0 Then
        'ÇÏÑíÓ
         get_special_index = 101
    End If
     
   If InStr(QERAA, "ÇáæÑÇŞ") > 0 Then
         'ÇáæÑÇŞ
         get_special_index = 102
    End If

End Function

Public Function get_tareq() As String
    get_tareq = " ãä ØÑíŞ "
    If CheckBox3.Value = True Then
     
        If CheckBox14.Value = True Or CheckBox15.Value = True Or CheckBox16.Value = True Or CheckBox31.Value = True Or CheckBox32.Value = True Or CheckBox33.Value = True Or CheckBox34.Value = True Or CheckBox35.Value = True Or CheckBox36.Value = True Then
            get_tareq = get_tareq + "ÇáÏÑÉ"
        Else
            get_tareq = get_tareq + "ÇáÔÇØÈíÉ"
        End If
        
        If CheckBox37.Value = True Or CheckBox42.Value = True Or CheckBox6.Value = True Or CheckBox5.Value = True Then
            get_tareq = " ãä ØÑíŞ ÇáÔÇØÈíÉ æ ÇáÏÑÉ"
        End If
        
     End If
     
     If CheckBox4.Value = True And CheckBox3.Value = True Then
         get_tareq = get_tareq + " æ ÇáØíÈÉ"
     ElseIf CheckBox4.Value = True Then
         get_tareq = get_tareq + "ÇáØíÈÉ"
     End If

End Function

Private Sub CommandButton7_Click()
  
   Dim counter As Integer

  For counter = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(counter) Then
        DB_ejaza_creator (Split(ListBox1.Column(0, counter), "-")(0))
    End If
    Next counter
   
End Sub

Private Sub ComboBox1_Change()
        
   Set rst_object = db_object.OpenRecordset("Select CONTENT.ID , CONTENT.STUDENT_NAME , CONTENT.QERAA from  CONTENT , [ORDER] WHERE  CONTENT.ORDER_ID = [ORDER].ID AND TYPE IN ( 'EJAZA' , 'DSGN' )  AND [ORDER].SHEIKH_NAME = '" & Split(ComboBox1.Value, "-")(1) & "'")
   ListBox1.Clear
      Do Until rst_object.EOF
          ListBox1.AddItem (rst_object.Fields("ID") & "-" & rst_object.Fields("STUDENT_NAME") & "-" & rst_object.Fields("QERAA"))
        rst_object.MoveNext
       Loop

End Sub
Private Function close_db()
  rst_object.Close
  db_object.Close
  Set db_object = Nothing
  Set rst_object = Nothing
End Function
Private Sub CommandButton1_Click()

    Dim index As Integer
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
   
    sheikh_type = get_sheikh_type()
    student_type = get_student_type()
    status = get_status()
    index = get_index()
      
    If index <> 0 Then
        
        TAREQ = get_tareq()
        sanada = sanadan(index)
        rawy = rawye(index)
        qeraat = qeraatn(index)
        qeraat = qeraat + TAREQ
        rawy = rawy + TAREQ
        
        ActiveDocument.SaveAs2 FileName:=ActiveDocument.path + Application.PathSeparator + student_name + " - " + qeraat + ".docx", FileFormat:= _
         wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
         :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
         :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
         SaveAsAOCELetter:=False, CompatibilityMode:=14
             
        update_sheikh_student sheikh_type:=True, student_type:=True
        set_qeraat STATE:=status, qeraat:=qeraat, rawy:=rawy
        set_snada (sanada)
        Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, NAME:="1"
        ActiveDocument.Save
        
    End If
    
    Dim tempForm As UserForm1
    For Each tempForm In UserForms
        Unload tempForm
    Next
    
End Sub

Private Sub CommandButton3_Click()
    
    Dim temp As Integer
    Dim index As Integer
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
     
    sheikh_type = get_sheikh_type()
    student_type = get_student_type()
    status = get_status()
    
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
   
     ActiveDocument.SaveAs2 FileName:=ActiveDocument.path + Application.PathSeparator + Replace(rawy, "/", "") + ".docx", FileFormat:= _
     wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
     :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
     :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
     SaveAsAOCELetter:=False, CompatibilityMode:=14
        
     update_sheikh_student
     set_qeraat STATE:=status, qeraat:=qeraat, rawy:=rawy
     set_snada (sanada)
        
     ActiveDocument.Save
     wdApp.Documents(ActiveDocument.path + Application.PathSeparator + Replace(rawy, "/", "") + ".docx").Close
     
     loop_counter = loop_counter + 1
     
  Wend
       
    Dim tempForm As UserForm
    For Each tempForm In UserForms
        Unload tempForm
    Next

End If

End Sub
Private Sub DB_ejaza_finish()
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, NAME:="1"
        Dim tempForm As UserForm1
        For Each tempForm In UserForms
            Unload tempForm
            Dim x As Integer
        Next
       ActiveDocument.Save
End Sub
Private Sub DB_ejaza_creator(ejaza_id As String)

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
    
   Dim NAME As String
   Dim INFO As String
   Dim QERAA As String
   Dim TAREk As String
   Dim STATE As String
   Dim GENDER As Boolean
   
   Set db = OpenDatabase(database_path)
   Set rst = db.OpenRecordset("Select * from CONTENT where ID = " & ejaza_id)
       
   If rst.RecordCount > 0 Then
      
      NAME = rst.Fields("STUDENT_NAME")
      QERAA = rst.Fields("QERAA")
      TAREk = rst.Fields("TAREQ")
      STATE = rst.Fields("STATE")
      INFO = rst.Fields("STUDENT_INFO")
      
      If (rst.Fields("STUDENT_GENDER") = "ØÇáÈÉ") Then
        GENDER = False
      Else
        GENDER = True
      End If
      
      TextBox5.text = NAME & vbNewLine & INFO & vbNewLine & rst.Fields("STUDENT_GENDER") & vbNewLine & QERAA & vbNewLine & TAREk & vbNewLine & STATE
   Else
      TextBox5.text = "ãÚÑİ ÇáÅÌÇÒÉ ÛíÑ ãæÌæÏ"
      GoTo endF:
   End If
   
   
    Dim sheikh_type As Integer
    Dim sheikh_name As String
    Dim sheikh_info As String
        
    Dim student_type As Boolean
    Dim student_name As String
    Dim student_info As String
         
    Dim Rng As Range, iPage As Long
    Dim rawy As String
    Dim sanada As String
         
    Dim index As Integer
    Dim status As String
    Dim qeraat As String
    Dim TAREQ As String
    Dim short_name As String
      
    sheikh_name = TextBox1.text
    sheikh_info = TextBox3.text
    sheikh_type = get_sheikh_type()
     
    student_name = NAME
    TextBox2.text = NAME
    student_info = INFO
    student_type = GENDER
    TextBox4.text = INFO
    
    short_name = Split(student_name, " ")(0) + " " + Split(student_name, " ")(1) + " " + Split(student_name, " ")(2)
     
    index = get_special_index(QERAA)
     
   ' set egaza status
    If InStr(STATE, "ÇÎÊÈÇÑÇ") > 0 Then
        status = "ÇÎÊÈÇÑÇ"
    ElseIf InStr(STATE, "ÈÚÖ") > 0 Then
        status = "ÈÚÖ ÇáŞÑÂä"
    ElseIf InStr(STATE, "ÎÊãÉ") > 0 Then
        status = "ÎÊãÉ ßÇãáÉ"
    Else
        status = STATE
    End If
    
    If InStr(STATE, "ÛíÈÇ") > 0 Then
        status = status + " " + "ÛíÈÇ Úä ÙåÑ ŞáÈ"
    ElseIf InStr(STATE, "äÙÑÇ") > 0 Then
        status = status + " " + "äÙÑÇ ãä ÇáãÕÍİ"
    Else
         status = STATE
    End If
    
    
    If InStr(TAREk, "ÕÛÑì") > 0 And index < 0 And index > -5 Then
        TAREQ = " ãä ØÑíŞ ÇáÔÇØÈíÉ æÇáÏÑÉ"
    ElseIf InStr(TAREk, "ÕÛÑì") > 0 And index > 80 And index < 110 Then
        TAREQ = " ãä ØÑíŞ ÇáÏÑÉ"
    ElseIf InStr(TAREk, "ÕÛÑì") > 0 And index > 7 And index < 11 Then
        TAREQ = " ãä ØÑíŞ ÇáÏÑÉ"
    ElseIf InStr(TAREk, "ÕÛÑì") > 0 Then
        TAREQ = " ãä ØÑíŞ ÇáÔÇØÈíÉ"
    ElseIf InStr(TAREk, "ßÈÑì") > 0 Then
        TAREQ = " ãä ØÑíŞ ÇáØíÈÉ"
    Else
        TAREQ = TAREk
    End If

    
    sanada = sanadan(index)
    rawy = rawye(index)
    qeraat = qeraatn(index)
    
    qeraat = qeraat + TAREQ
    rawy = rawy + TAREQ
      
    createWord (ActiveDocument.path + Application.PathSeparator + short_name + " - " + QERAA + ".docx")
    update_sheikh_student sheikh_type:=True, student_type:=GENDER
    set_qeraat STATE:=status, qeraat:=qeraat, rawy:=rawy
    set_snada (sanada)
          
    If CheckBox43.Value = False Then
        PdfSaving
        sampleSaving
    End If
    
    closeWord (ActiveDocument.path + Application.PathSeparator + short_name + " - " + QERAA + ".docx")
    
endF:
End Sub
Private Sub closeWord(filePath As String)
 
    Set wdApp = GetObject(, "Word.Application")
    ActiveDocument.Save
    wdApp.Documents(filePath).Close
    Documents.Open FileName:=ActiveDocument.path + Application.PathSeparator + "temp.docx", ReadOnly:=False
  
End Sub

Private Sub createWord(filePath As String)
 
   ActiveDocument.SaveAs2 FileName:=filePath, FileFormat:= _
         wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
         :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
         :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
         SaveAsAOCELetter:=False, CompatibilityMode:=14
    
   ActiveDocument.SaveAs2 FileName:=ActiveDocument.path + Application.PathSeparator + "temp.docx", FileFormat:= _
         wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
         :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
         :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
         SaveAsAOCELetter:=False, CompatibilityMode:=14
       
    Documents.Open FileName:=filePath, ReadOnly:=False
   
End Sub
Private Sub CommandButton6_Click()
   
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
    sheikh_type = get_sheikh_type()
     
    student_name = "ÇÓã ÇáØÇáÈ åäÇ"
    student_info = "ÈíÇäÇÊ ÇáØÇáÈ åäÇ"
    student_type = GENDER
     
    ActiveDocument.SaveAs2 FileName:=ActiveDocument.path + Application.PathSeparator + "sample.docx", FileFormat:= _
    wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
    :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
    :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
    SaveAsAOCELetter:=False, CompatibilityMode:=14
    
    update_sheikh_student sheikh_type:=True, student_type:=True
    index = get_special_index("ÍİÕ")
    
    status = "ÎÊãÉ ßÇãáÉ"
    status = status + " " + "ÛíÈÇ Úä ÙåÑ ŞáÈ"
    TAREQ = " ãä ØÑíŞ ÇáÔÇØÈíÉ"
    
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
   ActiveDocument.Save

End Sub

Private Sub OptionButton3_Click()
 TextBox3.text = "ãŞÑÆ æãÚáã ÇáŞÑÂä ÇáßÑíã æÇáÊÌæíÏ"
End Sub

Private Sub OptionButton4_Click()
 TextBox3.text = "ãŞÑÆÉ æãÚáãÉ ÇáŞÑÂä ÇáßÑíã æÇáÊÌæíÏ"
End Sub

Sub PdfSaving()
    Dim FileName As String
    FileName = Split(ActiveDocument.NAME, ".", 2)(i)
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        ActiveDocument.path + Application.PathSeparator + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False

End Sub
Sub sampleSaving()
    Dim FileName As String
    FileName = Split(ActiveDocument.NAME, ".", 2)(i)
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        "E:\other\print\smpl" + Application.PathSeparator + "smpl-" + FileName + ".pdf", ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForOnScreen, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
End Sub
