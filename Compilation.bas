Attribute VB_Name = "Compilation"
Sub Compilation()

nb_sheets = Sheets.Count


''' RENOMMAGE des feuilles EXCEL (par soucis de lisibilité) '''

For i = 1 To nb_sheets
    nb_lines_activesheet = Sheets(i).Cells(Rows.Count, 1).End(xlUp).Row
    nb_columns_activesheet = Sheets(i).Cells(1, Columns.Count).End(xlToLeft).Column
    pos = InStr(1, Sheets(i).Name, "30")
    If pos = 0 Then
        pos = InStr(1, Sheets(i).Name, "50")
    End If
    Sheets(i).Name = Mid(Sheets(i).Name, pos, 14)
    
''' PARTIE FORMULES Excel Somme et Moyenne'''

    Sheets(i).Cells(1, nb_columns_activesheet + 1).Value = "Consommation_totale_par_date_sur_24_heures"
    Sheets(i).Cells(nb_lines_activesheet, 1).Value = "Moyenne_de_la_consommation_par_tranche_de_10_minutes_et_pour_toutes_les_dates"
    Adresse = Right(Cells(2, nb_columns_activesheet).Address, 4)
    Adresse = Left(Adresse, 2) + Right(Adresse, 1)
    Sheets(i).Select
    Sheets(i).Cells(2, nb_columns_activesheet + 1).FormulaLocal = "=SOMME(B2:" & Adresse & ")"
    Sheets(i).Cells(2, nb_columns_activesheet + 2).FormulaLocal = "=MOYENNE(B2:" & Adresse & ")"
    Range(Cells(2, nb_columns_activesheet + 1), Cells(2, nb_columns_activesheet + 2)).Select
    Selection.AutoFill Destination:=Range(Cells(2, nb_columns_activesheet + 1), Cells(nb_lines_activesheet, nb_columns_activesheet + 2))
    
''' PARTIE GRAPHIQUE '''
    
    Range(Cells(1, nb_columns_activesheet + 1), Cells(nb_lines_activesheet - 1, nb_columns_activesheet + 1)).Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.Parent.Cut
    Range("B2").Select
    ActiveSheet.Paste
    ActiveSheet.Shapes("Graphique 2").ScaleWidth 2.5, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("Graphique 2").ScaleHeight 1.5, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("Graphique 2").Name = "1_CPT_" & Sheets(i).Name
    ActiveSheet.ChartObjects("1_CPT_" & Sheets(i).Name).Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Consommation_totale_par_date_et_24_heures_CPT_" & Sheets(i).Name
    ActiveChart.ChartArea.Select

    Range(Cells(nb_lines_activesheet, 1), Cells(nb_lines_activesheet, nb_columns_activesheet)).Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.Parent.Cut
    Range("B30").Select
    ActiveSheet.Paste
    ActiveSheet.Shapes("Graphique 4").ScaleWidth 2.5, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("Graphique 4").ScaleHeight 1.5, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("Graphique 4").Name = "2_CPT_" & Sheets(i).Name
    ActiveSheet.ChartObjects("2_CPT_" & Sheets(i).Name).Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Moyenne_de_la_consommation_par_tranche_de_10_minutes_et_pour_toutes_les_dates_CPT_" & Sheets(i).Name
    ActiveChart.ChartArea.Select

Next

End Sub

