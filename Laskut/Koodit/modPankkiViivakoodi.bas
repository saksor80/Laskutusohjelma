Attribute VB_Name = "modPankkiViivakoodi"
Public Function PankkiViivakoodi(strSaajantilinumero As String, strLaskunsumma As String, _
                                 strViitenumero As String, strEräpäivä As String) As String

'Tekijä:            Sakari Sorja
'Luotu:             4.2.2014
'Muokattu viimeksi: 6.2.2014
'Kuvaus:            PankkiViivakoodi Versio 4 käyttää viivakoodin Code 128 tulkinta C:tä.
'                   The check character is calculated from a weighted sum (modulo 103) of all the characters.
'                   128C (Code Set C) - 00-99 (encodes each two digits with one code) and FNC1.

'Esimerkki:
'                   Start C                                                                                      Stop
'           Value  [105] 45 81 01 71 00 00 00 12 20 00 48 29 90 00 00 00 05 59 58 22 43 29 46 71 12 01 31 [55] [stop]
'       Weighting    -   1  2  3  4  5  6  7  8  9  10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27

Dim strVersio, strVaralla, strTarkiste, strViivakoodi As String
Dim lngKerroin, lngSumma, lngPaikka As Long

strLaskunsumma = Format(strLaskunsumma, "0.00")
strVersio = "4"
strVaralla = "000"

lngSumma = 105
lngKerroin = 1
lngPaikka = 1

strViivakoodi = strVersio & ReplaceStr(Right$(strSaajantilinumero, 20), " ", "", 1) & LPad(Left$(strLaskunsumma, Len(strLaskunsumma) - 3), 6, 0) _
          & Right$(strLaskunsumma, 2) & strVaralla & LPad(ReplaceStr(strViitenumero, " ", "", 1), 20, 0) & Format(strEräpäivä, "yymmdd")

For i = 1 To 27
    
    lngSumma = lngSumma + CInt(Mid$(strViivakoodi, lngPaikka, 2)) * lngKerroin
    lngKerroin = lngKerroin + 1
    lngPaikka = lngPaikka + 2

Next

strTarkiste = CStr(lngSumma Mod 103)
strViivakoodi = strViivakoodi & strTarkiste
strViivakoodi = Chr("205") & MerkkijonoToAscii(strViivakoodi) & Chr("206")

PankkiViivakoodi = strViivakoodi

End Function

Public Function MerkkijonoToAscii(strViivakoodi As String) As String

' Tekijä:            Sakari Sorja
' Luotu:             1.7.2014
' Kuvaus:            Muuttaa numeroita sisältävän merkkijonon ASCII-merkeiksi Code 128 tulkinta C-viivakoodia varten.

Dim strTemp1, strTemp2, strPalautus As String
Dim intTemp1, intTemp2 As Integer

Do While Len(strViivakoodi) > 0
    
    strTemp1 = Left$(strViivakoodi, 2)
    
    Select Case True

        Case Len(strViivakoodi) = 3
             
             intTemp1 = CInt(strTemp1)
             intTemp1 = intTemp1 + 32
             strTemp1 = Chr(intTemp1)
             
             strTemp2 = Right$(strViivakoodi, 1)
             intTemp2 = CInt(strTemp2)
             intTemp2 = intTemp2 + 32
             strTemp2 = Chr(intTemp2)
             
             strPalautus = strPalautus & strTemp1 & strTemp2
             strViivakoodi = Right$(strViivakoodi, Len(strViivakoodi) - 3)
        
        Case strMerkkijono = "00"
            
             strTemp1 = Chr("207")
             
             strPalautus = strPalautus & strTemp1
             strViivakoodi = Right$(strViivakoodi, Len(strViivakoodi) - 2)
        
        Case strMerkkijono = "95" Or strMerkkijono = "96" Or strMerkkijono = "97" Or strMerkkijono = "98" Or strMerkkijono = "99"
             
             intTemp1 = CInt(strTemp1)
             intTemp1 = intTemp1 + 100
             strTemp1 = Chr(intTemp1)
             
             strPalautus = strPalautus & strTemp1
             strViivakoodi = Right$(strViivakoodi, Len(strViivakoodi) - 2)
             
        Case Else
                        
             intTemp1 = CInt(strTemp1)
             intTemp1 = intTemp1 + 32
             strTemp1 = Chr(intTemp1)
             
             strPalautus = strPalautus & strTemp1
             strViivakoodi = Right$(strViivakoodi, Len(strViivakoodi) - 2)

    End Select
    
Loop

MerkkijonoToAscii = strPalautus

End Function
