Attribute VB_Name = "modIFRSLayouts"
Option Explicit
'============================================================
' IFRS LAYOUT FORMATTERS
'  • FormatJournal        – general journal entry grid
'  • FormatTAccount       – 2?column T?account
'  • FormatLedger         – running?balance ledger
'  • FormatStatement      – financial statement body
'  • FormatCalcTable      – working schedule (e.g. amortisation)
'============================================================

'---------------------- JOURNAL ------------------------------
Public Sub FormatJournal()
    Dim rg As Range, hdr As Range
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rg = Selection
    Set hdr = rg.Rows(1)

    'Columns: Date | Account | Description | Debit | Credit
    hdr.Cells(1, 1).Resize(, 5).Value = _
        Array("Date", "Account", "Description", "Debit", "Credit")
    hdr.HorizontalAlignment = xlCenter
    hdr.Font.Bold = True
    
    rg.Columns(1).NumberFormat = "yyyy-mm-dd"
    
    'Single rule beneath header
    With hdr.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin
    End With
End Sub

'--------------------- T-ACCOUNT -----------------------------
Public Sub FormatTAccount()
    Dim rg As Range, hdr As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rg = Selection
    Set hdr = rg.Rows(1)
    
    hdr.Merge
    hdr.Value = "Account name"
    hdr.Font.Bold = True: hdr.HorizontalAlignment = xlCenter
    
    'Draw the “T”
    rg.Rows(2).Cells(1, 1).Resize(, 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
    rg.Columns(1).Borders(xlEdgeRight).LineStyle = xlContinuous

End Sub

'---------------------- LEDGER -------------------------------
Public Sub FormatLedger()
    Dim rg As Range, hdr As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rg = Selection
    Set hdr = rg.Rows(1)
    
    hdr.Cells(1, 1).Resize(, 6).Value = _
        Array("Date", "Account", "Description", "Debit", "Credit", "Balance")
    hdr.Font.Bold = True: hdr.HorizontalAlignment = xlCenter
    
    rg.Columns(1).NumberFormat = "yyyy-mm-dd"
    
    With hdr.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin
    End With
End Sub

'-------------------- STATEMENT BODY -------------------------
Public Sub FormatStatement()
    Dim rg As Range, hdr As Range, c As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rg = Selection
    Set hdr = rg.Rows(1)
    
    hdr.Value = "Statement heading": hdr.Merge
    hdr.Font.Bold = True: hdr.Font.Size = 12
    hdr.HorizontalAlignment = xlCenter
    
    'Totals / Net rows: bold + double?rule below
    For Each c In rg.Columns(1).Cells
        If LCase(Trim(c.Value)) Like "*total*" Or LCase(Trim(c.Value)) Like "*net*" Then
            c.EntireRow.Font.Bold = True
            With c.EntireRow.Borders(xlEdgeTop)
                .LineStyle = xlContinuous: .Weight = xlThin
            End With
            With c.EntireRow.Borders(xlEdgeBottom)
                .LineStyle = xlDouble: .Weight = xlMedium
            End With
        End If
    Next c
End Sub


