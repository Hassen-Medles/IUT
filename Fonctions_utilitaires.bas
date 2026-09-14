Attribute VB_Name = "Fonctions_utilitaires"
Option Explicit

Function CouleurPalette(index As Long) As Long
    Dim palette(0 To 9) As Long
    palette(0) = RGB(31, 119, 180)
    palette(1) = RGB(255, 127, 14)
    palette(2) = RGB(44, 160, 44)
    palette(3) = RGB(214, 39, 40)
    palette(4) = RGB(148, 103, 189)
    palette(5) = RGB(140, 86, 75)
    palette(6) = RGB(227, 119, 194)
    palette(7) = RGB(127, 127, 127)
    palette(8) = RGB(188, 189, 34)
    palette(9) = RGB(23, 190, 207)
    CouleurPalette = palette(index Mod 10)
End Function

Sub TrierArrayLong(ByRef arr As Variant)
    Dim i As Long, j As Long, temp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = LBound(arr) To UBound(arr) - 1 - i
            If CLng(arr(j)) > CLng(arr(j + 1)) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next j
    Next i
End Sub
