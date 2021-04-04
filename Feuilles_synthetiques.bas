Attribute VB_Name = "Feuilles_synthetiques"
Sub Feuilles_synthetiques()
Attribute Feuilles_synthetiques.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Sheets(1).Select
    Sheets.Add
    Sheets(1).Select
    Sheets(1).Name = "Feuil2"
    Sheets("Feuil2").Select
    Sheets("Feuil2").Name = "Conso_totale_date"
    Sheets(2).Select
    Sheets.Add
    Sheets(2).Select
    Sheets(2).Name = "Moyenne_conso_tranche"

nb_sheets = Sheets.Count

For i = 3 To nb_sheets
    Sheets(i).Select
    ActiveSheet.ChartObjects("1_CPT_" & Sheets(i).Name).Activate
    ActiveChart.ChartArea.Copy
    Sheets("Conso_totale_date").Select
    Cells(2 + (13 * (i - 3)), 2 + no_next_colonne).Select
    ActiveSheet.Paste
    ActiveSheet.Shapes("1_CPT_" & Sheets(i).Name).ScaleWidth 0.5, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("1_CPT_" & Sheets(i).Name).ScaleHeight 0.5, msoFalse, msoScaleFromTopLeft
    Sheets(i).Select
    ActiveSheet.ChartObjects("2_CPT_" & Sheets(i).Name).Activate
    ActiveChart.ChartArea.Copy
    Sheets("Moyenne_conso_tranche").Select
    Cells(2 + (13 * (i - 3)), 2 + no_next_colonne).Select
    ActiveSheet.Paste
    ActiveSheet.Shapes("2_CPT_" & Sheets(i).Name).ScaleWidth 0.5, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("2_CPT_" & Sheets(i).Name).ScaleHeight 0.5, msoFalse, msoScaleFromTopLeft
Next

End Sub
