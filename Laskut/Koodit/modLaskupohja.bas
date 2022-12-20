Attribute VB_Name = "modLaskupohja"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function SeuraavaLaskunNumero() As Integer
Attribute SeuraavaLaskunNumero.VB_ProcData.VB_Invoke_Func = " \n14"

Dim MyPath, MyName As String
Dim MyFirst, MyPresent As Integer
Dim lngLaskuri As Long

MyPath = "D:\Saku Yritys\Laskut\L‰htev‰t laskut\" & "*.xlsm"
MyName = Dir(MyPath)

lngLaskuri = 0

Do While MyName <> ""
       
       lngLaskuri = lngLaskuri + 1
       If lngLaskuri = 1 Then MyFirst = CInt(Trim$(Mid(MyName, 9, InStr(9, MyName, " ") - 18)))
       MyPresent = CInt(Trim$(Mid(MyName, 9, InStr(9, MyName, " ") - 18)))
       If MyPresent > MyFirst Then MyFirst = MyPresent
       MyName = Dir
Loop

SeuraavaLaskunNumero = MyFirst + 1
       
End Function

Public Function Viitetarkiste(strJono As String) As Byte

'Funktio palauttaa arvonaan tarkisteen, jos se ei ole oikein tai laskettiin uusi tarkiste,
'muuten palautetaan ' '.Laskentaalgoritmi on seuraava: kerrotaan numerosarjan numerot oikealta
'vasemmalle luvuilla 7, 3, 1, 7, 3, 1... ja lasketaansaaadut tulot yhteen.Tulojen summa v‰hennet‰‰n
'summaa seuraavasta nollaan p‰‰ttyv‰st‰ luvusta, jolloin erotuksesta tulee tarkiste.Esimerkiksi
'numerosarjan 484700 tarkisteen lasku:
'                                              4   8   4   7   0   0
'                                     painot   1   3   7   1   3   7
'                                     tulot    4  24   28  7   0   0

'tulojen summaksi tule 63 ja seuraava nollaan p‰‰ttyv‰ luku on 70, joten tarkisteeksi tulee 70 - 63 = 7
    
Dim intMerkkej‰ As Integer
Dim intSumma, i, intTemp As Integer
Dim strLoppu, strTarkistus, strJono2, strTuloste As String
    
strJono = Trim$(strJono)
intSumma = 0
intMerkkej‰ = Len(strJono)
For i = intMerkkej‰ To 1 Step -1
    intSumma = intSumma + CInt(Mid(strJono, i, 1)) * Choose((intMerkkej‰ - i) Mod 3 + 1, 7, 3, 1)
Next
Viitetarkiste = IIf(intSumma Mod 10 <> 0, 10 - intSumma Mod 10, intSumma Mod 10)
    
End Function

Public Function ReplaceStr(TextIn, SearchStr, Replacement, CompMode As Integer)

'******************************************************************************
' Replaces the SearchStr string with Replacement string in the TextIn string. *
' Uses CompMode to determine comparison mode                                  *
'******************************************************************************

Dim WorkText As String, Pointer As Integer
  If IsNull(TextIn) Then
    ReplaceStr = Null
  Else
    WorkText = TextIn
    Pointer = InStr(1, WorkText, SearchStr, CompMode)
    Do While Pointer > 0
      WorkText = Left(WorkText, Pointer - 1) & Replacement & Mid(WorkText, Pointer + Len(SearchStr))
      Pointer = InStr(Pointer + Len(Replacement), WorkText, SearchStr, CompMode)
    Loop
    ReplaceStr = WorkText
  End If
  
End Function

Sub TallennaJaAvaaUusiLasku()

Dim strLaskunNimi As String
Dim MyPath, MyName As String

MyPath = "D:\Saku Yritys\Laskut\L‰htev‰t laskut\PDF\" & "*.pdf"
MyName = Dir(MyPath)

strLaskunNimi = Trim$(ReplaceStr(Range("I8"), " ", "", 1)) & " " & Range("C9")

Range("C56:I60").Select
    
With Selection
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
End With
Selection.UnMerge

Workbooks.Open Filename:= _
"D:\Saku Yritys\Laskut\L‰htev‰t laskut\" 'ensimm‰inen lasku t‰h‰n

Range("C56:I60").Select
    
With Selection
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
End With
Selection.UnMerge

Windows("Laskupohja.xlsm").Activate

Range("B2:J64").Select
Selection.Copy
    
Windows("").Activate 'eka lasku t‰h‰n myˆs
    
Range("B2:J64").Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
    
Range("C2").Select
ActiveCell.FormulaR1C1 = "               Yritys (Y-tunnus: 1000000-1)" 'Firma t‰h‰n
With ActiveCell.Characters(Start:=16, Length:=18).Font
    .Name = "Calibri"
    .FontStyle = "Lihavoitu"
    .Size = 16
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontMinor
End With
With ActiveCell.Characters(Start:=34, Length:=21).Font
    .Name = "Calibri"
    .FontStyle = "Normaali"
    .Size = 10
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontMinor
End With
    
Range("C61").Select
ActiveCell.FormulaR1C1 = "Yritys (Y-tunnus: 1000000-1)" 'Firma t‰h‰nkin
With ActiveCell.Characters(Start:=1, Length:=18).Font
    .Name = "Calibri"
    .FontStyle = "Normaali"
    .Size = 12
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontMinor
End With
With ActiveCell.Characters(Start:=19, Length:=21).Font
    .Name = "Calibri"
    .FontStyle = "Normaali"
    .Size = 10
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontMinor
End With
    
Range("C56:I60").Select
    
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Selection.Merge

ActiveWorkbook.SaveAs Filename:= _
"D:\Saku Yritys\Laskut\L‰htev‰t laskut\Lasku - " & strLaskunNimi & ".xlsm", _
FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False

If MyName <> "" Then

    Do While MyName <> ""
       
        If Left$(MyName, Len(MyName) - 4) = Left$(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) Then
       
            ActiveSheet.Shapes.Range(Array("Button 10")).Select
            Selection.OnAction = "AvaaPDF"
            ActiveSheet.Buttons("Button 10").Font.ColorIndex = 1
        Else
       
            ActiveSheet.Shapes.Range(Array("Button 10")).Select
            Selection.OnAction = "Tyhj‰"
            ActiveSheet.Buttons("Button 10").Font.ColorIndex = 15
            
        End If
        MyName = Dir
    Loop

Else

    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    Selection.OnAction = "Tyhj‰"
    ActiveSheet.Buttons("Button 10").Font.ColorIndex = 15

End If

Range("A1").Select

Windows("Laskupohja.xlsm").Activate
    
Range("C56:I60").Select
    
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Selection.Merge
    
Application.CutCopyMode = False
Range("A1").Select

Windows("Lasku - " & Trim$(ReplaceStr(Range("I8"), " ", "", 1)) & " " & Range("C9") & ".xlsm").Activate

End Sub

Sub Sulje()

Application.DisplayAlerts = False
Application.Quit

End Sub

Sub Tyhj‰()

End Sub

Sub AvaaPDF()

Filename = "D:\Saku Yritys\Laskut\L‰htev‰t laskut\PDF\" & Left$(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & ".pdf"
ShellExecute 0, "Open", Filename, "", "", vbNormalNoFocus

End Sub

Sub Maksumuistutus()

Dim strTiedostonNimi, strMuistutus, strPaska As String
Dim intKorko, intSumma, intP‰iv‰t As Integer
Dim dblViiv‰styskorko As Double

strTiedostonNimi = ActiveSheet.ListBox1.Value & ".xlsm"

Workbooks.Open Filename:= _
"D:\Saku Yritys\Laskut\L‰htev‰t laskut\" & strTiedostonNimi

Windows(strTiedostonNimi).Activate

intKorko = CInt(Left$(CStr(Range("I7")), 1))
intSumma = CInt(Range("I47"))
intP‰iv‰t = CInt(Date - Range("I6"))

Range("G2") = "MAKSUMUISTUTUS"
Range("I4") = Date
Range("I9") = ""

If Range("I6") + 14 > Date Then

    strMuistutus = "1"
    
Else

    strMuistutus = "2"

End If

For i = 22 To 43

    Select Case True
    
        Case Range("C" & i) = ""
            
            Range("C" & i) = "Maksumuistutusmaksu"
            Range("D" & i) = 1
            Range("E" & i) = "kpl"
            
            If strMuistutus = "1" Then
            
                Range("F" & i) = 5
                Range("I" & i) = 5
                
            End If
            
            If strMuistutus = "2" Then
            
                Range("F" & i) = 10
                Range("I" & i) = 10
            
            End If
            
            Range("C" & i + 1) = "Viiv‰styskorko"
            Range("D" & i + 1) = 1
            Range("E" & i + 1) = "kpl"
            Range("F" & i + 1) = CCur(intKorko * intSumma * intP‰iv‰t / 36500)
            Range("I" & i + 1) = Range("F" & i + 1)
            dblViiv‰styskorko = CDbl(Range("F" & i + 1))
            Range("I47") = Range("I" & i) + Range("I" & i + 1) + intSumma
            Exit For
                
    End Select
    
Next

Range("C49:I52").Select

Muistutustekstialue

Range("C49") = "Kirjanpitomme mukaan emme ole saaneet " & Range("I6") & " menness‰ maksusuoritusta, jonka summa er‰p‰iv‰n‰ oli " & intSumma & " euroa. " & _
               "Kustakin maksumuistutuksesta veloitamme 5 euroa huomautuskulua. T‰m‰ on " & strMuistutus & ". muistutus, joten huomautuskulut ovat yhteens‰ " & CStr(CInt(strMuistutus) * 5) & " euroa. " & _
               "Lis‰ksi maksettava viiv‰styskoron m‰‰r‰ " & CStr(Format(dblViiv‰styskorko, "0.00")) & " euroa " & CStr(Date - Range("I6")) & " korkop‰iv‰lt‰. " & _
               "Maksumuistutus on aiheeton, mik‰li olette jo maksaneet laskun. Mik‰li Teill‰ on huomautettavaa laskun johdosta, ottanette yhteytt‰ mahdollisimman " & _
               "pian. Muussa tapauksessa katsomme, ett‰ kahden maksumuistutuksen j‰lkeen Teill‰ ei ole huomautettavaa saatavamme johdosta ja jos lasku on viel‰ maksamatta, lasku siirtyy perint‰‰n."
               
Range("C56").Select
ActiveCell.FormulaR1C1 = "=PankkiViivakoodi(R[-45]C[6],R[-9]C[6],R[-48]C[6],R[-50]C[6])"

Range("C56").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Application.CutCopyMode = False

Range("I6") = "HETI"

Range("A1").Select

ActiveWorkbook.SaveAs Filename:= _
        "D:\Saku Yritys\Laskut\L‰htev‰t laskut\Maksumuistutukset\Maksumuistutus " & strMuistutus & "." & " " & "-" & " " & Right$(strTiedostonNimi, Len(strTiedostonNimi) - 8), _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
               
End Sub

Sub Muistutustekstialue()

    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With

End Sub
