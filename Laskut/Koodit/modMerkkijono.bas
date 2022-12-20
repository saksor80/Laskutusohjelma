Attribute VB_Name = "modMerkkijono"
Public Function LPad(s As String, L As Integer, Optional Merkki As String) As String
    If Merkki = "" Then
        Merkki = " "
     End If
    If L - Len(s) > 0 Then
        Do Until Len(s) = L
            s = Merkki & s
        Loop
        LPad = Space(L - Len(s)) & s
    Else
        LPad = s
    End If
End Function
