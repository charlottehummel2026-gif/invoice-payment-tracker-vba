Attribute VB_Name = "ModFormulas"
Option Explicit

Public Sub RepairFormulas()
    Dim lastRow As Long
    Dim table As ListObject
    With ThisWorkbook.Worksheets("Invoices")
        Set table = .ListObjects("InvoicesTable")
        If table.ListRows.Count = 0 Then table.ListRows.Add
        lastRow = table.Range.row + table.Range.Rows.Count - 1
        .Range("H9").Formula = "=IF(OR(A9="""",P9<>""OK"",E9>Overview!$K$4),"""",SUMIFS(PaymentsTable[Amount],PaymentsTable[Invoice ID],A9,PaymentsTable[Received on],""<=""&Overview!$K$4,PaymentsTable[Entry check],""OK""))"
        .Range("H9:H" & lastRow).FillDown
        .Range("I9").Formula = "=IF(ISNUMBER(H9),MAX(G9-H9,0),"""")"
        .Range("I9:I" & lastRow).FillDown
        .Range("J9").Formula = "=IF(ISNUMBER(H9),MAX(H9-G9,0),"""")"
        .Range("J9:J" & lastRow).FillDown
        .Range("K9").Formula = "=IF(ISNUMBER(I9),IF(I9>0,MAX(Overview!$K$4-F9,0),0),"""")"
        .Range("K9:K" & lastRow).FillDown
        .Range("L9").Formula = "=IF(A9="""","""",IF(P9<>""OK"",""Review entry"",IF(E9>Overview!$K$4,""Not issued"",IF(COUNTIFS(PaymentsTable[Invoice ID],A9,PaymentsTable[Entry check],""<>OK"",PaymentsTable[Entry check],""?*"")>0,""Review payments"",IF(J9>0,""Overpaid"",IF(I9=0,""Paid"",IF(K9>0,""Overdue"",IF(F9=Overview!$K$4,""Due today"",IF(H9>0,""Part paid"",""Open"")))))))))"
        .Range("L9:L" & lastRow).FillDown
        .Range("P9").Formula = "=IF(A9="""",IF(COUNTA(B9:G9)=0,"""",""Missing invoice ID""),IF(COUNTIFS(InvoicesTable[Invoice ID],A9)>1,""Duplicate invoice ID"",IF(OR(B9="""",D9="""",ISBLANK(E9),ISBLANK(F9),ISBLANK(G9)),""Missing required input"",IF(OR(NOT(ISNUMBER(E9)),NOT(ISNUMBER(F9)),NOT(ISNUMBER(G9))),""Check date or amount"",IF(OR(F9<E9,G9<=0),""Check due date / amount"",""OK"")))))"
        .Range("P9:P" & lastRow).FillDown
        .Range("E9:F" & lastRow).NumberFormat = "mmm d, yyyy"
        .Range("N9:N" & lastRow).NumberFormat = "mmm d, yyyy"
        .Range("G9:J" & lastRow).NumberFormat = "$#,##0.00;($#,##0.00);0.00"
    End With
    With ThisWorkbook.Worksheets("Payments")
        Set table = .ListObjects("PaymentsTable")
        If table.ListRows.Count = 0 Then table.ListRows.Add
        lastRow = table.Range.row + table.Range.Rows.Count - 1
        .Range("C9").Formula = "=IF(B9="""","""",IF(COUNTIFS(InvoicesTable[Invoice ID],B9)<>1,""Check invoice ID"",INDEX(InvoicesTable[Customer],MATCH(B9,InvoicesTable[Invoice ID],0))))"
        .Range("C9:C" & lastRow).FillDown
        .Range("J9").Formula = "=IF(A9="""",IF(COUNTA(B9,D9:I9)=0,"""",""Missing payment ID""),IF(COUNTIFS(PaymentsTable[Payment ID],A9)>1,""Duplicate payment ID"",IF(OR(B9="""",ISBLANK(D9),E9="""",ISBLANK(F9),G9="""",H9=""""),""Missing required input"",IF(COUNTIFS(InvoicesTable[Invoice ID],B9)<>1,""Check invoice ID"",IF(OR(NOT(ISNUMBER(D9)),NOT(ISNUMBER(F9))),""Check date or amount"",IF(OR(F9<=0,D9<INDEX(InvoicesTable[Issued],MATCH(B9,InvoicesTable[Invoice ID],0))),""Check date / amount"",""OK""))))))"
        .Range("J9:J" & lastRow).FillDown
        .Range("D9:D" & lastRow).NumberFormat = "mmm d, yyyy"
        .Range("F9:F" & lastRow).NumberFormat = "$#,##0.00;($#,##0.00);0.00"
    End With
End Sub
