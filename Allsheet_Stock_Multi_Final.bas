Attribute VB_Name = "Module1"
Sub Stock_Multi_Final()
    Dim xlsheet As Worksheet
    For Each xlsheet In ThisWorkbook.Worksheets
        xlsheet.Select
        Call Stock_Yr_Final
        xlsheet.Range("I:Q").Columns.AutoFit
    Next xlsheet

End Sub
