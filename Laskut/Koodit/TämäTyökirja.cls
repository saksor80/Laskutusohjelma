VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TämäTyökirja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

Dim MyPath, MyName As String

Application.DisplayAlerts = True

MyPath = "D:\Saku Yritys\Laskut\Lähtevät laskut\" & "*.xlsm"
MyName = Dir(MyPath)

If MyName <> "" Then

    Do While MyName <> ""
        
        MyName = Left$(MyName, Len(MyName) - 5)
        ActiveSheet.ListBox1.AddItem MyName
        MyName = Dir
        
    Loop

End If

Range("I5").Select
ActiveCell.FormulaR1C1 = "=SeuraavaLaskunNumero()"
ActiveSheet.Shapes.Range(Array("ListBox1")).Height = 550
Range("C8").Select

End Sub
