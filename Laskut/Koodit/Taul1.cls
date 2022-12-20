VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Taul1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim strTiedostonNimi As String
Dim MyPath, MyName As String

strTiedostonNimi = ActiveSheet.ListBox1.Value & ".xlsm"
    
Workbooks.Open Filename:= _
"D:\Saku Yritys\Laskut\Lähtevät laskut\" & strTiedostonNimi
    
MyPath = "D:\Saku Yritys\Laskut\Lähtevät laskut\PDF\" & "*.pdf"
MyName = Dir(MyPath)

Windows(strTiedostonNimi).Activate

If MyName <> "" Then

    Do While MyName <> ""
       
        If Left$(MyName, Len(MyName) - 4) = Left$(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) Then
       
            ActiveSheet.Shapes.Range(Array("Button 10")).Select
            Selection.OnAction = "AvaaPDF"
            ActiveSheet.Buttons("Button 10").Font.ColorIndex = 1
            Exit Do
        Else
       
            ActiveSheet.Shapes.Range(Array("Button 10")).Select
            Selection.OnAction = "Tyhjä"
            ActiveSheet.Buttons("Button 10").Font.ColorIndex = 15
            
        End If
        MyName = Dir
    Loop

Else

    ActiveSheet.Shapes.Range(Array("Button 10")).Select
    Selection.OnAction = "Tyhjä"
    ActiveSheet.Buttons("Button 10").Font.ColorIndex = 15

End If
    
ActiveSheet.Range("A1").Select
    
End Sub
