Attribute VB_Name = "modFormatLayouts"
Option Explicit

'?? Helper: write header row with bold, centered text and a bottom border ??
Private Sub WriteHeader(rowRange As Range, headers As Variant)
    With rowRange.Resize(1, UBound(headers) + 1)
        .Value = headers
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End With
End Sub

Public Sub FormatJournal()
    On Error GoTo ErrHandler
    Dim sel As Range: Set sel = Selection
    If Not TypeOf sel Is Range Or sel.Areas.Count > 1 Then Exit Sub
    If sel.Columns.Count < 5 Then
        MsgBox "Select at least 5 columns for Journal formatting", vbExclamation
        Exit Sub
    End If

    ' Header row
    WriteHeader sel.Rows(1), Array("Date", "Account", "Description", "Debit", "Credit")
    sel.Columns(1).NumberFormat = DefaultDateFmt
    Exit Sub

ErrHandler:
    MsgBox "FormatJournal error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Public Sub FormatLedger()
    On Error GoTo ErrHandler
    Dim sel As Range: Set sel = Selection
    If Not TypeOf sel Is Range Then Exit Sub
    If sel.Columns.Count < 6 Then
        MsgBox "Select at least 6 columns for Ledger formatting", vbExclamation
        Exit Sub
    End If

    ' Header row
    WriteHeader sel.Rows(1), Array("Date", "Account", "Description", "Debit", "Credit", "Balance")
    sel.Columns(1).NumberFormat = DefaultDateFmt
    Exit Sub

ErrHandler:
    MsgBox "FormatLedger error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Public Sub FormatTAccount()
    On Error GoTo ErrHandler
    Dim sel As Range: Set sel = Selection
    If Not TypeOf sel Is Range Then Exit Sub
    If sel.Rows.Count < 2 Or sel.Columns.Count < 2 Then
        MsgBox "Select at least a 2?2 range for a T?account", vbExclamation
        Exit Sub
    End If

    ' Header row
    With sel.Rows(1).Resize(1, sel.Columns.Count)
        .Merge
        .Value = "Account name"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End With

    ' Bottom rule on row 2, cols 1Ð2
    With sel.Rows(2).Cells(1, 1).Resize(1, 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' Vertical divider down colÊ1
    With sel.Columns(1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    Exit Sub

ErrHandler:
    MsgBox "FormatTAccount error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Public Sub FormatStatement()
    On Error GoTo ErrHandler
    Dim sel As Range: Set sel = Selection
    If Not TypeOf sel Is Range Then Exit Sub

    ' Statement heading
    With sel.Rows(1).Resize(1, sel.Columns.Count)
        .Merge
        .Value = "Statement heading"
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
    End With

    ' Totals/net rows
    Dim cell As Range
    For Each cell In sel.Columns(1).Cells
        If LCase(Trim(cell.Value & "")) Like "*total*" Or LCase(Trim(cell.Value & "")) Like "*net*" Then
            With cell.EntireRow
                .Font.Bold = True
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlDouble
                    .Weight = xlMedium
                End With
            End With
        End If
    Next
    Exit Sub

ErrHandler:
    MsgBox "FormatStatement error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub


