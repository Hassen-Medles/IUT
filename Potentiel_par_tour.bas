Attribute VB_Name = "Potentiel_par_tour"
Option Explicit

' MH PE = f(Temps par tour) comparaison de tours sélectionnés
Sub CreerGraphique2_PE_TempsParTour()
    Dim wsData As Worksheet, wsGraph As Worksheet
    Dim derLigne As Long, i As Long, k As Long
    Dim essaiCible As String, toursTxt As String
    Dim listeTours() As String
    Dim cht As ChartObject
    Dim serie As Series
    Set wsData = ThisWorkbook.Worksheets(feuil_data)
    Set wsGraph = ThisWorkbook.Worksheets(feuil_graphic)
    ' essai étudié
    essaiCible = Trim(wsGraph.Range(Cell_Essai_Choisie).Value)
    If essaiCible = "" Then
        MsgBox "Merci de renseigner l'essai en " & Cell_Essai_Choisie, vbExclamation
        Exit Sub
    End If
    If LCase(Left(essaiCible, 5)) <> "essai" Then essaiCible = "Essai " & essaiCible
    'tours choisis
    toursTxt = Trim(wsGraph.Range(Cell_Tours_Choisis).Value)
    If toursTxt = "" Then
        MsgBox "Merci de renseigner la liste des tours en " & Cell_Tours_Choisis & " (ex: 1,10,20,30,40)", vbExclamation
        Exit Sub
    End If
    listeTours = Split(Replace(toursTxt, " ", ""), ",")
    derLigne = wsData.Cells(wsData.Rows.Count, index_data_essai).End(xlUp).Row
    If derLigne < 2 Then
        MsgBox "Aucune donnée en feuille " & feuil_data, vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    wsGraph.ChartObjects("ChartPE_TempsParTour").Delete
    On Error GoTo 0
    'Création du graphique
    Set cht = wsGraph.ChartObjects.Add(Left:=20, Top:=220, Width:=700, Height:=400)
    cht.Name = "ChartPE_TempsParTour"
    cht.Chart.ChartType = type_graphic_courbe
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = "Comparaison PE = f(temps par tour) - " & essaiCible
    'insertion des valeurs
    For k = 0 To UBound(listeTours)
        If Not IsNumeric(listeTours(k)) Then GoTo SuivantTour
        Dim numTour As Long
        numTour = CLng(listeTours(k))
        Dim xVals() As Double, yVals() As Double, nb As Long
        nb = 0
        ReDim xVals(1 To derLigne)
        ReDim yVals(1 To derLigne)
        For i = 2 To derLigne
            If wsData.Cells(i, index_data_essai).Value = essaiCible And CLng(wsData.Cells(i, index_data_tour).Value) = numTour Then
                nb = nb + 1
                xVals(nb) = wsData.Cells(i, index_data_temps_tour).Value
                yVals(nb) = wsData.Cells(i, index_data_Potentiel).Value
            End If
        Next i

        If nb > 0 Then
            ReDim Preserve xVals(1 To nb)
            ReDim Preserve yVals(1 To nb)
            #création des séries pour chaque tour
            Set serie = cht.Chart.SeriesCollection.NewSeries
            With serie
                .Name = "Tour " & numTour
                .XValues = xVals
                .Values = yVals
                .ChartType = type_graphic_courbe
                .Format.Line.Weight = trait_graphic
                .Format.Line.ForeColor.RGB = CouleurPalette(k)
                .MarkerStyle = xlMarkerStyleNone
            End With
        Else
            MsgBox "Aucune donnée pour le tour " & numTour & " dans " & essaiCible, vbInformation
        End If
SuivantTour:
    Next k

    With cht.Chart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Temps par tour"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "PE"
        .HasLegend = True
    End With
End Sub
