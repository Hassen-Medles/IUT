Attribute VB_Name = "Graphic_6"
Option Explicit

Sub CreerGraphique6_PEMin_ParTour()
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

    On Error Resume Next
    wsGraph.ChartObjects("ChartPEMin_ParTour").Delete
    On Error GoTo 0

    Set cht = wsGraph.ChartObjects.Add(Left:=20, Top:=1700, Width:=default_graphic_width, Height:=default_graphic_height)
    cht.Name = "ChartPEMin_ParTour"
    cht.Chart.ChartType = type_graphic_courbe
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = "PE min par tour - Comparaison des essais"

    compteurCouleur = 0
    For Each e In essaisTries
        Dim dictMin As Object
        Set dictMin = CreateObject("Scripting.Dictionary")

        Dim numTour As Long
        Dim vPE As Variant

        For i = 2 To derLigne
            If Trim(wsData.Cells(i, index_data_essai).Value) = e Then
                vPE = wsData.Cells(i, index_data_Potentiel).Value
                If IsNumeric(wsData.Cells(i, index_data_tour).Value) And IsNumeric(vPE) Then
                    numTour = CLng(wsData.Cells(i, index_data_tour).Value)
                    If Not dictMin.Exists(numTour) Then
                        dictMin.Add numTour, CDbl(vPE)
                    Else
                        If CDbl(vPE) < dictMin(numTour) Then dictMin(numTour) = CDbl(vPE)
                    End If
                End If
            End If
        Next i

        If dictMin.Count > 0 Then
            Dim toursTries As Variant
            toursTries = dictMin.Keys
            Call TrierArrayLong(toursTries)

            Dim n As Long
            n = UBound(toursTries) - LBound(toursTries) + 1
            Dim xVals() As Double, yVals() As Double
            ReDim xVals(1 To n)
            ReDim yVals(1 To n)

            Dim t As Variant, idx As Long
            idx = 0
            For Each t In toursTries
                idx = idx + 1
                xVals(idx) = CDbl(t)
                yVals(idx) = dictMin(CLng(t))
            Next t

            Set serie = cht.Chart.SeriesCollection.NewSeries
            With serie
                .Name = CStr(e)
                .XValues = xVals
                .Values = yVals
                .ChartType = type_graphic_courbe
                .Format.Line.Weight = trait_graphic
                .Format.Line.ForeColor.RGB = CouleurPalette(compteurCouleur)
                .MarkerStyle = xlMarkerStyleNone
            End With

            compteurCouleur = compteurCouleur + 1
        End If
    Next e

    With cht.Chart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Tour"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "PE min"
        .HasLegend = True
    End With
End Sub

Sub CreerGraphique7_PEMax_ParTour()
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

    On Error Resume Next
    wsGraph.ChartObjects("ChartPEMax_ParTour").Delete
    On Error GoTo 0

    Set cht = wsGraph.ChartObjects.Add(Left:=760, Top:=1700, Width:=default_graphic_width, Height:=default_graphic_height)
    cht.Name = "ChartPEMax_ParTour"
    cht.Chart.ChartType = type_graphic_courbe
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = "PE max par tour - Comparaison des essais"

    compteurCouleur = 0
    For Each e In essaisTries
        Dim dictMax As Object
        Set dictMax = CreateObject("Scripting.Dictionary")

        Dim numTour As Long
        Dim vPE As Variant

        For i = 2 To derLigne
            If Trim(wsData.Cells(i, index_data_essai).Value) = e Then
                vPE = wsData.Cells(i, index_data_Potentiel).Value
                If IsNumeric(wsData.Cells(i, index_data_tour).Value) And IsNumeric(vPE) Then
                    numTour = CLng(wsData.Cells(i, index_data_tour).Value)
                    If Not dictMax.Exists(numTour) Then
                        dictMax.Add numTour, CDbl(vPE)
                    Else
                        If CDbl(vPE) > dictMax(numTour) Then dictMax(numTour) = CDbl(vPE)
                    End If
                End If
            End If
        Next i

        If dictMax.Count > 0 Then
            Dim toursTries As Variant
            toursTries = dictMax.Keys
            Call TrierArrayLong(toursTries)

            Dim n As Long
            n = UBound(toursTries) - LBound(toursTries) + 1
            Dim xVals() As Double, yVals() As Double
            ReDim xVals(1 To n)
            ReDim yVals(1 To n)

            Dim t As Variant, idx As Long
            idx = 0
            For Each t In toursTries
                idx = idx + 1
                xVals(idx) = CDbl(t)
                yVals(idx) = dictMax(CLng(t))
            Next t

            Set serie = cht.Chart.SeriesCollection.NewSeries
            With serie
                .Name = CStr(e)
                .XValues = xVals
                .Values = yVals
                .ChartType = type_graphic_courbe
                .Format.Line.Weight = trait_graphic
                .Format.Line.ForeColor.RGB = CouleurPalette(compteurCouleur)
                .MarkerStyle = xlMarkerStyleNone
            End With

            compteurCouleur = compteurCouleur + 1
        End If
    Next e

    With cht.Chart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Tour"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "PE max"
        .HasLegend = True
    End With
End Sub

Sub CreerGraphique8_AmplitudePE_ParTour()
    Dim wsData As Worksheet, wsGraph As Worksheet
    Dim derLigne As Long, i As Long
    Dim dictEssais As Object
    Dim dictToursCommuns As Object
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

    Dim dictAmplitudeParEssai As Object
    Set dictAmplitudeParEssai = CreateObject("Scripting.Dictionary")
    Set dictToursCommuns = CreateObject("Scripting.Dictionary")

    For Each e In essaisTries
        Dim dictMin As Object, dictMax As Object
        Set dictMin = CreateObject("Scripting.Dictionary")
        Set dictMax = CreateObject("Scripting.Dictionary")

        Dim numTour As Long
        Dim vPE As Variant

        For i = 2 To derLigne
            If Trim(wsData.Cells(i, index_data_essai).Value) = e Then
                vPE = wsData.Cells(i, index_data_Potentiel).Value
                If IsNumeric(wsData.Cells(i, index_data_tour).Value) And IsNumeric(vPE) Then
                    numTour = CLng(wsData.Cells(i, index_data_tour).Value)
                    If Not dictMin.Exists(numTour) Then
                        dictMin.Add numTour, CDbl(vPE)
                        dictMax.Add numTour, CDbl(vPE)
                    Else
                        If CDbl(vPE) < dictMin(numTour) Then dictMin(numTour) = CDbl(vPE)
                        If CDbl(vPE) > dictMax(numTour) Then dictMax(numTour) = CDbl(vPE)
                    End If
                    If Not dictToursCommuns.Exists(numTour) Then dictToursCommuns.Add numTour, True
                End If
            End If
        Next i

        Dim dictAmplitude As Object
        Set dictAmplitude = CreateObject("Scripting.Dictionary")
        Dim t As Variant
        For Each t In dictMin.Keys
            dictAmplitude.Add t, dictMax(t) - dictMin(t)
        Next t
        dictAmplitudeParEssai.Add e, dictAmplitude
    Next e

    If dictToursCommuns.Count = 0 Then
        MsgBox "Aucune donnée trouvée en feuille " & feuil_data, vbExclamation
        Exit Sub
    End If

    Dim toursCommunsTries As Variant
    toursCommunsTries = dictToursCommuns.Keys
    Call TrierArrayLong(toursCommunsTries)

    Dim categories() As String
    Dim nbTours As Long
    nbTours = UBound(toursCommunsTries) - LBound(toursCommunsTries) + 1
    ReDim categories(1 To nbTours)
    Dim k As Long
    For k = 1 To nbTours
        categories(k) = CStr(toursCommunsTries(LBound(toursCommunsTries) + k - 1))
    Next k

    On Error Resume Next
    wsGraph.ChartObjects("ChartAmplitudePE_ParTour").Delete
    On Error GoTo 0

    Set cht = wsGraph.ChartObjects.Add(Left:=20, Top:=2120, Width:=default_graphic_width, Height:=default_graphic_height)
    cht.Name = "ChartAmplitudePE_ParTour"
    cht.Chart.ChartType = xlColumnClustered
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = "Amplitude PE (max - min) par tour - Comparaison des essais"

    compteurCouleur = 0
    For Each e In essaisTries
        Dim dictAmp As Object
        Set dictAmp = dictAmplitudeParEssai(e)

        Dim yVals() As Double
        ReDim yVals(1 To nbTours)
        For k = 1 To nbTours
            Dim tourCourant As Long
            tourCourant = CLng(categories(k))
            If dictAmp.Exists(tourCourant) Then
                yVals(k) = dictAmp(tourCourant)
            Else
                yVals(k) = 0
            End If
        Next k

        Set serie = cht.Chart.SeriesCollection.NewSeries
        With serie
            .Name = CStr(e)
            .XValues = categories
            .Values = yVals
            .ChartType = xlColumnClustered
            .Format.Fill.ForeColor.RGB = CouleurPalette(compteurCouleur)
        End With

        compteurCouleur = compteurCouleur + 1
    Next e

    With cht.Chart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Tour"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Amplitude PE"
        .HasLegend = True
    End With
End Sub
