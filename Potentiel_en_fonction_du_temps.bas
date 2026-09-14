Attribute VB_Name = "Potentiel_en_fonction_du_temps"
Option Explicit

' MH:PE = f(Temps continu), 1 couleur par tour
Sub CreerGraphique1_PE_TempsContinu()
    Dim wsData As Worksheet, wsGraph As Worksheet
    Dim derLigne As Long, i As Long
    Dim essaiCible As String
    Dim numTour As Long
    Dim dictTours As Object
    Dim cht As ChartObject
    Dim serie As Series
    Dim compteurCouleur As Long
    Dim t As Variant
    Set wsData = ThisWorkbook.Worksheets(feuil_data)
    Set wsGraph = ThisWorkbook.Worksheets(feuil_graphic)
    'récup de l'essai étudié
    essaiCible = Trim(wsGraph.Range(Cell_Essai_Choisie).Value)
    If essaiCible = "" Then
        MsgBox "Merci de renseigner l'essai à afficher en " & Cell_Essai_Choisie, vbExclamation
        Exit Sub
    End If
    If LCase(Left(essaiCible, 5)) <> "essai" Then essaiCible = "Essai " & essaiCible
    
    derLigne = wsData.Cells(wsData.Rows.Count, index_data_essai).End(xlUp).Row
    If derLigne < 2 Then
        MsgBox "Aucune donnée en feuille " & feuil_data, vbExclamation
        Exit Sub
    End If
    'Lister les tours présents pour cet essai
    Set dictTours = CreateObject("Scripting.Dictionary")
    For i = 2 To derLigne
        If wsData.Cells(i, index_data_essai).Value = essaiCible Then
            numTour = CLng(wsData.Cells(i, index_data_tour).Value)
            If Not dictTours.Exists(numTour) Then dictTours.Add numTour, True
        End If
    Next i

    If dictTours.Count = 0 Then
        MsgBox "Aucune donnée trouvée pour '" & essaiCible & "'.", vbExclamation
        Exit Sub
    End If
    Dim toursTries As Variant
    toursTries = dictTours.Keys
    Call TrierArrayLong(toursTries)
    'Supprimer le graphique existant si présent
    On Error Resume Next
    wsGraph.ChartObjects("ChartPE_TempsContinu").Delete
    On Error GoTo 0
    Set cht = wsGraph.ChartObjects.Add(Left:=20, Top:=20, Width:=700, Height:=400)
    cht.Name = "ChartPE_TempsContinu"
    cht.Chart.ChartType = type_graphic_courbe
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = "PE = f(temps) - " & essaiCible
    'Une série par tour
    compteurCouleur = 0
    For Each t In toursTries
        Dim xVals() As Double, yVals() As Double, nb As Long
        nb = 0
        ReDim xVals(1 To derLigne)
        ReDim yVals(1 To derLigne)
        'récupération des valeurs du potentiel en fonction du temps continue
        For i = 2 To derLigne
            If wsData.Cells(i, index_data_essai).Value = essaiCible And CLng(wsData.Cells(i, index_data_tour).Value) = CLng(t) Then
                nb = nb + 1
                xVals(nb) = wsData.Cells(i, index_data_temps_continue).Value
                yVals(nb) = wsData.Cells(i, index_data_Potentiel).Value
            End If
        Next i
        If nb > 0 Then
            ReDim Preserve xVals(1 To nb)
            ReDim Preserve yVals(1 To nb)

            Set serie = cht.Chart.SeriesCollection.NewSeries
            With serie
                .Name = "Tour " & t
                .XValues = xVals
                .Values = yVals
                .ChartType = type_graphic_courbe
                .Format.Line.Weight = trait_graphic
                .Format.Line.ForeColor.RGB = CouleurPalette(compteurCouleur)
                .MarkerStyle = xlMarkerStyleNone
            End With
            compteurCouleur = compteurCouleur + 1
        End If
    Next t
    
    With cht.Chart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Temps"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "PE"
        .HasLegend = True
    End With
End Sub
