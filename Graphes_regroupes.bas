Attribute VB_Name = "Graphes_regroupes"
Sub Graphe_regroupe()

    Range("A1").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveSheet.Shapes("Graphique 5").Name = "Superposition"
    Application.CutCopyMode = False

nb_sheets = Sheets.Count

For i = 3 To nb_sheets
    ActiveSheet.ChartObjects("2_CPT_" & Sheets(i).Name).Activate
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.ChartArea.Copy
    ActiveSheet.ChartObjects("Graphique 5").Activate
    ActiveChart.Paste
Next

End Sub
