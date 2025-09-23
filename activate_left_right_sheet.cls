' 激活当前workbook下最左边的sheet
Sub ActivateLeftmostSheet()
    ThisWorkbook.Sheets(1).Activate
End Sub


' 激活当前workbook下最右边的sheet
Sub ActivateRightmostSheet()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    wb.Sheets(wb.Sheets.Count).Activate
End Sub