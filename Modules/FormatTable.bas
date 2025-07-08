Attribute VB_Name = "modFormatTable"
Option Explicit
'===============================================
'  IFRS Table Border & Total Formatting
'   ¥ Header row = any non?empty row that has an
'     empty row immediately above it (inside tbl)
'   ¥ Bold + thin underline on header text cells
'   ¥ Bold Total/Net row. Numeric cells in that
'     row get thin top line; if itÕs the last row
'     they also get a double bottom line.
'===============================================
Public Sub FormatTable()
    Dim tbl           As Range
    Dim lastRowIx     As Long          'relative index in tbl
    Dim r             As Long          'row counter
    Dim rowRange      As Range
    Dim cell          As Range
    Dim numRange      As Range         'numeric cells in Tot/Net row
    Dim cFirst        As Range         'first?column cell being scanned
    Dim relRowIx      As Long
    
    'ÐÐ validate selection
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set tbl = Selection
    lastRowIx = tbl.Rows.Count
    
    'ÐÐÐÐÐ 1. HEADER ROWS (may be several) ÐÐÐÐÐ
    For r = 1 To lastRowIx
        'Is this row non?empty?
        If Application.WorksheetFunction.CountA(tbl.Rows(r)) > 0 Then
            'Is the row above empty? (or r = 1 Ð top of selection)
            If r = 1 Or Application.WorksheetFunction.CountA(tbl.Rows(r - 1)) = 0 Then
                Set rowRange = tbl.Rows(r)
                
                'Bold + underline only non?numeric, non?empty cells
                For Each cell In rowRange.Cells
                    If Len(Trim(cell.Value)) > 0 And Not IsNumeric(cell.Value) Then
                        cell.Font.Bold = True
                        With cell.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                        End With
                    End If
                Next cell
            End If
        End If
    Next r
    
    'ÐÐÐÐÐ 2. TOTAL / NET ROW(S) ÐÐÐÐÐ
    For Each cFirst In tbl.Columns(1).Cells
        If LCase(Trim(cFirst.Value)) Like "*total*" Or _
           LCase(Trim(cFirst.Value)) Like "*net*" Then
           
            relRowIx = cFirst.Row - tbl.Row + 1
            Set rowRange = tbl.Rows(relRowIx)
            rowRange.Font.Bold = True        'bold the whole row
            
            'Collect all numeric, non?empty cells in this row
            Set numRange = Nothing
            For Each cell In rowRange.Cells
                If Len(Trim(cell.Value)) > 0 And IsNumeric(cell.Value) Then
                    If numRange Is Nothing Then
                        Set numRange = cell
                    Else
                        Set numRange = Union(numRange, cell)
                    End If
                End If
            Next cell
            
            'Apply borders if we found numeric cells
            If Not numRange Is Nothing Then
                'Thin line ABOVE those numeric cells
                With numRange.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                'Double line BELOW those numeric cells Ð only if last table row
                If relRowIx = lastRowIx Then
                    With numRange.Borders(xlEdgeBottom)
                        .LineStyle = xlDouble
                        .Weight = xlThick
                    End With
                End If
            End If
        End If
    Next cFirst
End Sub


