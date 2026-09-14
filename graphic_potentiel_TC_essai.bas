Attribute VB_Name = "graphic_potentiel_TC_essai"
Option Explicit
' MH GRAPHIQUE 3 : PE = f(Temps continu), TENDANCE (moyenne mobile) par ESSAI
Sub CreerGraphique3_PE_TempsContinu_ParEssai()
    Dim wsData As Worksheet, wsGraph As Worksheet
    Dim derLigne As Long, i As Long
    Dim dictEssais As Object
    Dim cht As ChartObject
    Dim serie As Series
    Dim compteurCouleur As Long
    Dim e As Variant

    Set wsData = ThisWorkbook.Worksheets(feuil_data)
    Set wsGraph = ThisWorkbook.Worksheets(feuil_graphic)
    derLigne = wsData.Cells(wsData.Rows.Count, index_data_essai).End(xlUp).Row
    If derLigne < 2 Then
        MsgBox "Aucune donnée en feuille " & feuil_data, vbExclamation
        Exit Sub
    End If
    '1. Lister les essais présents
    Set dictEssais = CreateObject("Scripting.Dictionary")
    Dim essaiVal As String
    For i = 2 To derLigne
        essaiVal = Trim(wsData.Cells(i, index_data_essai).Value)
        If essaiVal <> "" Then
            If Not dictEssais.Exists(essaiVal) Then dictEssais.Add essaiVal, True
        End If
    Next i
    If dictEssais.Count = 0 Then
        MsgBox "Aucune donnée trouvée en feuille " & feuil_data, vbExclamation
        Exit Sub
    End If
    Dim essaisTries As Variant
    essaisTries = dictEssais.Keys
    Call TrierArrayEssais(essaisTries)
    '2. Supprimer le graphique existant si présent
    On Error Resume Next
    wsGraph.ChartObjects("ChartPE_TempsContinu_ParEssai").Delete
    On Error GoTo 0
    Set cht = wsGraph.ChartObjects.Add(Left:=20, Top:=860, Width:=default_graphic_width, Height:=default_graphic_height)
    cht.Name = "ChartPE_TempsContinu_ParEssai"
    cht.Chart.ChartType = type_graphic_courbe
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = "PE = f(temps) - Tendance par essai"
    '3. Une série (lissée) par essai
    compteurCouleur = 0
    For Each e In essaisTries
        Dim xValsBrut() As Double, yValsBrut() As Double, nb As Long
        nb = 0
        ReDim xValsBrut(1 To derLigne)
        ReDim yValsBrut(1 To derLigne)
        For i = 2 To derLigne
            If Trim(wsData.Cells(i, index_data_essai).Value) = e Then
                Dim vTemps As Variant, vPE As Variant
                vTemps = wsData.Cells(i, index_data_temps_continue).Value
                vPE = wsData.Cells(i, index_data_Potentiel).Value

                If IsNumeric(vTemps) And IsNumeric(vPE) Then
                    nb = nb + 1
                    xValsBrut(nb) = CDbl(vTemps)
                    yValsBrut(nb) = CDbl(vPE)
                End If
            End If
        Next i

        If nb > 0 Then
            ReDim Preserve xValsBrut(1 To nb)
            ReDim Preserve yValsBrut(1 To nb)

            ' Tri par X croissant (évite tout artefact de diagonale)
            Call TrierPointsParX(xValsBrut, yValsBrut, nb)

            ' Calcul de la moyenne mobile (tendance lissée)
            Dim xLisse() As Double, yLisse() As Double, nbLisse As Long
            Call CalculerMoyenneMobile(xValsBrut, yValsBrut, nb, taille_fenetre_tendance, xLisse, yLisse, nbLisse)

            If nbLisse > 0 Then
                Set serie = cht.Chart.SeriesCollection.NewSeries
                With serie
                    .Name = CStr(e)
                    .XValues = xLisse
                    .Values = yLisse
                    .ChartType = type_graphic_courbe
                    .Format.Line.Weight = trait_graphic
                    .Format.Line.ForeColor.RGB = CouleurPalette(compteurCouleur)
                    .MarkerStyle = xlMarkerStyleNone
                End With
                compteurCouleur = compteurCouleur + 1
            End If
        End If
    Next e

    With cht.Chart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Temps"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "PE (tendance)"
        .HasLegend = True
    End With

    wsGraph.Activate
    wsGraph.Application.Goto wsGraph.Range(cht.TopLeftCell.Address), True
End Sub

' MH Trie deux tableaux parallèles (X et Y) par X croissant
Sub TrierPointsParX(ByRef xArr() As Double, ByRef yArr() As Double, ByVal n As Long)
    Dim i As Long, j As Long
    Dim tempX As Double, tempY As Double
    For i = 1 To n - 1
        For j = 1 To n - i
            If xArr(j) > xArr(j + 1) Then
                tempX = xArr(j): xArr(j) = xArr(j + 1): xArr(j + 1) = tempX
                tempY = yArr(j): yArr(j) = yArr(j + 1): yArr(j + 1) = tempY
            End If
        Next j
    Next i
End Sub

' MH Calcule une moyenne mobile centrée sur (xIn, yIn), fenêtre = taille
' Résultat dans xOut, yOut, nbOut (tableaux redimensionnés en sortie)
Sub CalculerMoyenneMobile(ByRef xIn() As Double, ByRef yIn() As Double, ByVal n As Long, _
                           ByVal taille As Long, ByRef xOut() As Double, ByRef yOut() As Double, ByRef nbOut As Long)
    Dim i As Long, k As Long, demiFenetre As Long
    Dim borneBas As Long, borneHaut As Long
    Dim sommeY As Double, sommeX As Double, compte As Long

    If taille < 1 Then taille = 1
    demiFenetre = taille \ 2

    ReDim xOut(1 To n)
    ReDim yOut(1 To n)
    nbOut = 0

    For i = 1 To n
        borneBas = i - demiFenetre
        borneHaut = i + demiFenetre
        If borneBas < 1 Then borneBas = 1
        If borneHaut > n Then borneHaut = n

        sommeX = 0: sommeY = 0: compte = 0
        For k = borneBas To borneHaut
            sommeX = sommeX + xIn(k)
            sommeY = sommeY + yIn(k)
            compte = compte + 1
        Next k

        nbOut = nbOut + 1
        xOut(nbOut) = sommeX / compte
        yOut(nbOut) = sommeY / compte
    Next i

    If nbOut > 0 Then
        ReDim Preserve xOut(1 To nbOut)
        ReDim Preserve yOut(1 To nbOut)
    End If
End Sub

' MH Trie un tableau de chaînes "Essai N" par N croissant
Sub TrierArrayEssais(ByRef arr As Variant)
    Dim i As Long, j As Long, temp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = LBound(arr) To UBound(arr) - 1 - i
            If NumeroEssai(CStr(arr(j))) > NumeroEssai(CStr(arr(j + 1))) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next j
    Next i
End Sub

Function NumeroEssai(ByVal essaiTxt As String) As Long
    Dim s As String
    s = Trim(essaiTxt)
    Dim p As Long
    p = InStrRev(s, " ")
    If p > 0 And IsNumeric(Mid(s, p + 1)) Then
        NumeroEssai = CLng(Mid(s, p + 1))
    Else
        NumeroEssai = 0
    End If
End Function

