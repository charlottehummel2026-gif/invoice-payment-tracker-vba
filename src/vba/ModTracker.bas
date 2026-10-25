Attribute VB_Name = "ModTracker"
Option Explicit

' Invoice and Payment Tracker - fictional portfolio demonstration.
' Worksheet forms and ordinary Excel objects work on Mac and Windows.
' No network calls, external libraries, or application security changes.

Public Sub NewInvoice()
    ThisWorkbook.Worksheets("Entry").Activate
    ThisWorkbook.Worksheets("Entry").Range("C10").Select
End Sub

Public Sub RecordPayment()
    ThisWorkbook.Worksheets("Entry").Activate
    ThisWorkbook.Worksheets("Entry").Range("I10").Select
End Sub

Public Sub GoOverview()
    ThisWorkbook.Worksheets("Overview").Activate
    ThisWorkbook.Worksheets("Overview").Range("A1").Select
End Sub

Public Sub ClearInvoice()
    Dim r As Variant
    For Each r In Array(10, 12, 14, 16, 18, 20, 22, 24, 26)
        ThisWorkbook.Worksheets("Entry").Range("C" & r & ":F" & r).ClearContents
    Next r
End Sub

Public Sub ClearPayment()
    Dim r As Variant
    For Each r In Array(10, 12, 14, 16, 18, 20, 22)
        ThisWorkbook.Worksheets("Entry").Range("I" & r & ":L" & r).ClearContents
    Next r
End Sub

Public Function NextRecordId(ByVal ids As Range, ByVal prefix As String) As String
    Dim c As Range, n As Long, largest As Long, suffix As String
    For Each c In ids.Cells
        If Left$(CStr(c.Value2), Len(prefix)) = prefix Then
            suffix = Mid$(CStr(c.Value2), Len(prefix) + 1)
            If Len(suffix) > 0 And Len(suffix) < 9 Then
                If IsNumeric(suffix) Then
                    n = CLng(Val(suffix))
                    If n > largest Then largest = n
                End If
            End If
        End If
    Next c
    NextRecordId = prefix & Format$(largest + 1, "000")
End Function

Public Function ValidAmount(ByVal value As Variant) As Boolean
    If Not IsNumeric(value) Then Exit Function
    ValidAmount = (CDbl(value) >= 0.01 And CDbl(value) <= 999999999.99)
End Function

Public Function ValidEntryDate(ByVal value As Variant) As Boolean
    If IsEmpty(value) Then Exit Function
    If Len(Trim$(CStr(value))) = 0 Then Exit Function
    If Not IsDate(value) Then Exit Function
    ValidEntryDate = (CDate(value) >= DateSerial(2000, 1, 1) And CDate(value) <= DateSerial(2100, 12, 31))
End Function

Private Function NewDataRow(ByVal table As ListObject) As ListRow
    ' Reuse only a completely empty input row; formula columns can be present.
    Dim row As ListRow, i As Long, used As Boolean, cols As Variant, c As Variant
    If table.Name = "InvoicesTable" Then
        cols = Array(1, 2, 3, 4, 5, 6, 7, 13, 14, 15)
    Else
        cols = Array(1, 2, 4, 5, 6, 7, 8, 9)
    End If
    For i = 1 To table.ListRows.Count
        Set row = table.ListRows(i)
        used = False
        For Each c In cols
            If Len(CStr(row.Range.Cells(1, c).Value2)) > 0 Then used = True
        Next c
        If Not used Then
            Set NewDataRow = row
            Exit Function
        End If
    Next i
    Set NewDataRow = table.ListRows.Add(AlwaysInsert:=False)
End Function

Private Sub PutText(ByVal target As Range, ByVal value As Variant)
    ' Text format prevents notes or IDs beginning with '=' becoming formulas.
    target.NumberFormat = "@"
    target.Value2 = Trim$(CStr(value))
End Sub

Public Sub SaveInvoice()
    Dim form As Worksheet, table As ListObject, row As ListRow
    Dim id As String, reason As String, c As Variant
    Dim previousEvents As Boolean, committed As Boolean
    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    Set form = ThisWorkbook.Worksheets("Entry")
    Set table = ThisWorkbook.Worksheets("Invoices").ListObjects("InvoicesTable")
    If Len(Trim$(CStr(form.Range("C10").Value2))) = 0 Then reason = "Choose a customer."
    If reason = "" And Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets("Overview").Range("A13:A18"), form.Range("C10").Value2) <> 1 Then reason = "Choose a customer from the demo list."
    If reason = "" And Len(Trim$(CStr(form.Range("C14").Value2))) = 0 Then reason = "Enter a description."
    If reason = "" And Not ValidEntryDate(form.Range("C16").value) Then reason = "Enter a valid issued date (2000-2100)."
    If reason = "" And Not ValidEntryDate(form.Range("C18").value) Then reason = "Enter a valid due date (2000-2100)."
    If reason = "" Then
        If CDate(form.Range("C18").value) < CDate(form.Range("C16").value) Then reason = "Due date must be on or after issued date."
    End If
    If reason = "" And Not ValidAmount(form.Range("C20").Value2) Then reason = "Enter a positive amount, up to $999,999,999.99."
    If reason = "" And Len(Trim$(CStr(form.Range("C22").Value2))) = 0 Then reason = "Enter an owner."
    If reason = "" And Len(CStr(form.Range("C24").Value2)) > 0 Then
        If Not ValidEntryDate(form.Range("C24").value) Then reason = "Enter a valid follow-up date or leave it blank."
    End If
    If reason <> "" Then
        form.Range("B30").Value2 = reason
        MsgBox reason, vbExclamation, "Check invoice"
        Exit Sub
    End If
    id = NextRecordId(table.ListColumns(1).DataBodyRange, "INV-" & Year(CDate(form.Range("C16").value)) & "-")
    Application.EnableEvents = False
    Set row = NewDataRow(table)
    PutText row.Range.Cells(1, 1), id
    PutText row.Range.Cells(1, 2), form.Range("C10").Value2
    PutText row.Range.Cells(1, 3), form.Range("C12").Value2
    PutText row.Range.Cells(1, 4), form.Range("C14").Value2
    row.Range.Cells(1, 5).value = CDate(form.Range("C16").value)
    row.Range.Cells(1, 6).value = CDate(form.Range("C18").value)
    row.Range.Cells(1, 7).Value2 = Application.WorksheetFunction.Round(CDbl(form.Range("C20").Value2), 2)
    PutText row.Range.Cells(1, 13), form.Range("C22").Value2
    If Len(CStr(form.Range("C24").Value2)) > 0 Then row.Range.Cells(1, 14).value = CDate(form.Range("C24").value)
    PutText row.Range.Cells(1, 15), form.Range("C26").Value2
    RepairFormulas
    RecalculateTracker
    committed = True
    ClearInvoice
    form.Range("B30").Value2 = "Saved " & id & ". Invoice register and dashboard updated."
    Application.EnableEvents = previousEvents
    Exit Sub
Failed:
    reason = Err.Description
    On Error Resume Next
    If Not row Is Nothing And Not committed Then
        For Each c In Array(1, 2, 3, 4, 5, 6, 7, 13, 14, 15)
            row.Range.Cells(1, c).ClearContents
        Next c
    End If
    Application.EnableEvents = previousEvents
    MsgBox "Invoice could not be saved: " & reason, vbExclamation, "Invoice tracker"
End Sub

Public Sub SavePayment()
    Dim form As Worksheet, table As ListObject, invoices As ListObject, row As ListRow
    Dim id As String, invoice As String, reason As String, match As Variant, c As Variant
    Dim previousEvents As Boolean, committed As Boolean, method As String
    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    Set form = ThisWorkbook.Worksheets("Entry")
    Set table = ThisWorkbook.Worksheets("Payments").ListObjects("PaymentsTable")
    Set invoices = ThisWorkbook.Worksheets("Invoices").ListObjects("InvoicesTable")
    invoice = Trim$(CStr(form.Range("I10").Value2))
    If Application.WorksheetFunction.CountIf(invoices.ListColumns(1).DataBodyRange, invoice) <> 1 Then reason = "Enter one existing invoice ID from the invoice register."
    If reason = "" Then
        match = Application.match(invoice, invoices.ListColumns(1).DataBodyRange, 0)
        If invoices.DataBodyRange.Cells(match, 16).Value2 <> "OK" Then reason = "Correct the invoice entry before recording a payment."
    End If
    If reason = "" And Not ValidEntryDate(form.Range("I12").value) Then reason = "Enter a valid received date (2000-2100)."
    If reason = "" Then
        If CDate(form.Range("I12").value) < CDate(invoices.DataBodyRange.Cells(match, 5).value) Then reason = "Payment date must be on or after the invoice issued date."
    End If
    method = Trim$(CStr(form.Range("I14").Value2))
    Select Case method
        Case "ACH", "Bank transfer", "Card", "Check", "Cash"
        Case Else
            If reason = "" Then reason = "Choose a payment method from the list."
    End Select
    If reason = "" And Not ValidAmount(form.Range("I16").Value2) Then reason = "Enter a positive payment amount, up to $999,999,999.99."
    If reason = "" And Len(Trim$(CStr(form.Range("I18").Value2))) = 0 Then reason = "Enter a receipt reference."
    If reason = "" And Len(Trim$(CStr(form.Range("I20").Value2))) = 0 Then reason = "Enter who recorded the payment."
    If reason <> "" Then
        form.Range("B30").Value2 = reason
        MsgBox reason, vbExclamation, "Check payment"
        Exit Sub
    End If
    id = NextRecordId(table.ListColumns(1).DataBodyRange, "PAY-")
    Application.EnableEvents = False
    Set row = NewDataRow(table)
    PutText row.Range.Cells(1, 1), id
    PutText row.Range.Cells(1, 2), invoice
    row.Range.Cells(1, 4).value = CDate(form.Range("I12").value)
    PutText row.Range.Cells(1, 5), method
    row.Range.Cells(1, 6).Value2 = Application.WorksheetFunction.Round(CDbl(form.Range("I16").Value2), 2)
    PutText row.Range.Cells(1, 7), form.Range("I18").Value2
    PutText row.Range.Cells(1, 8), form.Range("I20").Value2
    PutText row.Range.Cells(1, 9), form.Range("I22").Value2
    RepairFormulas
    RecalculateTracker
    committed = True
    ClearPayment
    form.Range("B30").Value2 = "Saved " & id & " against " & invoice & ". Dashboard updated."
    Application.EnableEvents = previousEvents
    Exit Sub
Failed:
    reason = Err.Description
    On Error Resume Next
    If Not row Is Nothing And Not committed Then
        For Each c In Array(1, 2, 4, 5, 6, 7, 8, 9)
            row.Range.Cells(1, c).ClearContents
        Next c
    End If
    Application.EnableEvents = previousEvents
    MsgBox "Payment could not be saved: " & reason, vbExclamation, "Invoice tracker"
End Sub

Public Sub RecalculateTracker()
    ThisWorkbook.Worksheets("Payments").Calculate
    ThisWorkbook.Worksheets("Invoices").Calculate
    ThisWorkbook.Worksheets("Overview").Calculate
End Sub

Public Sub RefreshTracker()
    On Error GoTo Failed
    RepairFormulas
    RecalculateTracker
    ThisWorkbook.Worksheets("Entry").Range("B30").Value2 = "Calculated columns restored and dashboard refreshed."
    MsgBox "Dashboard refreshed using the reporting date on Overview.", vbInformation, "Invoice tracker"
    Exit Sub
Failed:
    MsgBox "Refresh could not finish: " & Err.Description, vbExclamation, "Invoice tracker"
End Sub

Public Sub ShowOverdue()
    On Error GoTo Failed
    RecalculateTracker
    With ThisWorkbook.Worksheets("Invoices")
        .Activate
        If .FilterMode Then .ShowAllData
        .ListObjects("InvoicesTable").Range.AutoFilter Field:=11, Criteria1:=">0"
    End With
    Exit Sub
Failed:
    MsgBox "Filter could not be applied: " & Err.Description, vbExclamation, "Invoice tracker"
End Sub

Public Sub ShowAllInvoices()
    With ThisWorkbook.Worksheets("Invoices")
        .Activate
        If .FilterMode Then .ShowAllData
    End With
End Sub
