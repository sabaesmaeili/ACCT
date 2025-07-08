Attribute VB_Name = "modIFRSLibrary"

Option Explicit
'===========================================================
'  IFRS LIBRARY Ð common captions & line-items per IASÊ1/IFRS
'  ¥ BuildHiddenSheet : creates/refreshes the very-hidden list
'  ¥ IFRS_AccountsArr : returns the whole list as a 0-based array
'  ¥ ApplyIFRSDropDown rng : puts a validation list on any range
'===========================================================

'----------- 1A ¥ MASTER LIST (split into ²20-line blocks) ------------
'Keeping each block <20 continuations avoids the VBA compiler limit.
Private Const IFRS_RAW_1 As String = _
"ASSETS" & vbLf & "Current assets" & vbLf & "Cash and cash equivalents" & vbLf & _
"Trade and other receivables" & vbLf & "Other receivables" & vbLf & "Inventories" & vbLf & _
"Contract assets" & vbLf & "Current tax assets" & vbLf & "Prepayments" & vbLf & _
"Financial assets at FVTPL" & vbLf & "Financial assets at amortised cost" & vbLf & _
"Derivative assets" & vbLf & "Other current assets" & vbLf & _
"Non-current assets" & vbLf & "Property, plant and equipment" & vbLf & _
"Accumulated depreciation Ð PPE" & vbLf & "Right-of-use assets" & vbLf & _
"Investment property" & vbLf & "Goodwill" & vbLf & "Intangible assets"

Private Const IFRS_RAW_2 As String = _
"Investments in associates" & vbLf & "Investments in joint ventures" & vbLf & "Biological assets" & vbLf & _
"Deferred tax assets" & vbLf & "Other non-current assets" & vbLf & _
"LIABILITIES" & vbLf & "Current liabilities" & vbLf & "Trade and other payables" & vbLf & _
"Accruals" & vbLf & "Contract liabilities" & vbLf & "Borrowings Ð current" & vbLf & _
"Lease liabilities Ð current" & vbLf & "Provisions Ð current" & vbLf & "Current tax liabilities" & vbLf & _
"Derivative liabilities" & vbLf & "Bank overdraft" & vbLf & _
"Non-current liabilities" & vbLf & "Borrowings Ð non-current" & vbLf & _
"Lease liabilities Ð non-current" & vbLf & "Provisions Ð non-current"

Private Const IFRS_RAW_3 As String = _
"Deferred tax liabilities" & vbLf & "Employee benefit obligations" & vbLf & _
"Contingent consideration payable" & vbLf & "Other non-current liabilities" & vbLf & _
"EQUITY" & vbLf & "Share capital" & vbLf & "Share premium" & vbLf & "Treasury shares" & vbLf & _
"Other reserves" & vbLf & "Revaluation surplus" & vbLf & _
"Foreign currency translation reserve" & vbLf & "Retained earnings" & vbLf & _
"Non-controlling interests" & vbLf & _
"STATEMENTÊOFÊPROFITÊORÊLOSS" & vbLf & "Revenue" & vbLf & "Cost of sales" & vbLf & _
"Costs" & vbLf & "Net income" & vbLf & _
"Statement of Financial Position" & vbLf & _
"Statement of Profit or Loss" & vbLf & _
"Statement of Cash Flows" & vbLf & _
"Statement of Changes in Equity" & vbLf & _
"Gross profit" & vbLf & "Other income" & vbLf & "Selling and distribution expenses"


Private Const IFRS_RAW_4 As String = _
"Administrative expenses" & vbLf & "Research and development expenses" & vbLf & _
"Impairment losses" & vbLf & "Other expenses" & vbLf & "Operating profit" & vbLf & _
"Finance income" & vbLf & "Finance costs" & vbLf & "Share of profit of associates" & vbLf & _
"Profit before tax" & vbLf & "Income tax expense" & vbLf & "Profit for the year" & vbLf & _
"STATEMENTÊOFÊOCI" & vbLf & "Items that will not be reclassified" & vbLf & _
"Items that may be reclassified" & vbLf & "Other comprehensive income" & vbLf & _
"Total comprehensive income" & vbLf & _
"CASHÊFLOWÊSTATEMENT" & vbLf & "Net cash from operating activities" & vbLf & _
"Net cash used in investing activities" & vbLf & "Net cash from financing activities"

Private Const IFRS_RAW_5 As String = _
"Increase/(decrease) in cash and cash equivalents" & vbLf & _
"STATEMENTÊOFÊFINANCIALÊPOSITIONÊTOTALS" & vbLf & "Total current assets" & vbLf & _
"Total non-current assets" & vbLf & "Total assets" & vbLf & "Total current liabilities" & vbLf & _
"Total non-current liabilities" & vbLf & "Total liabilities" & vbLf & "Total equity" & vbLf & _
"Total liabilities and equity" & vbLf & _
"STATEMENTÊOFÊCHANGESÊINÊEQUITY" & vbLf & "Balance at 1ÊJanuary" & vbLf & _
"Total comprehensive income for the year" & vbLf & "Dividends" & vbLf & "Other movements" & vbLf & _
"Balance at 31ÊDecember"

'----------- 1A-(cont.)  Additional blocks for managerial & cost accounting -----
Private Const IFRS_RAW_6 As String = _
"MANAGERIAL ACCOUNTING" & vbLf & _
"Inventories Ð overview" & vbLf & _
"Raw materials inventory" & vbLf & "Work-in-progress (WIP) inventory" & vbLf & _
"Finished goods inventory" & vbLf & "Merchandise inventory" & vbLf & _
"Direct materials" & vbLf & "Indirect materials" & vbLf & _
"Direct labour" & vbLf & "Indirect labour" & vbLf & _
"Manufacturing overhead" & vbLf & "Prime cost" & vbLf & _
"Conversion cost" & vbLf & "Cost of goods manufactured" & vbLf & _
"Cost of goods sold" & vbLf & "Standard cost" & vbLf & _
"Actual cost" & vbLf & "Absorption costing" & vbLf & "Variable costing"

Private Const IFRS_RAW_7 As String = _
"Costing systems & inventory flow" & vbLf & _
"Activity-based costing (ABC)" & vbLf & "Job order costing" & vbLf & _
"Process costing" & vbLf & "Throughput costing" & vbLf & _
"Joint cost" & vbLf & "Split-off point" & vbLf & "By-product" & vbLf & _
"Equivalent units of production" & vbLf & "FIFO process costing" & vbLf & _
"Weighted-average process costing" & vbLf & "Target costing" & vbLf & _
"Kaizen costing" & vbLf & "Life-cycle costing" & vbLf & "Backflush costing" & vbLf & _
"Just-in-time (JIT) inventory" & vbLf & _
"Master budget" & vbLf & "Operating budget" & vbLf & "Flexible budget"

Private Const IFRS_RAW_8 As String = _
"COSTÐVOLUMEÐPROFIT & PERFORMANCE" & vbLf & _
"CVP analysis" & vbLf & "Break-even point" & vbLf & "Contribution margin" & vbLf & _
"Contribution margin ratio" & vbLf & "Degree of operating leverage" & vbLf & _
"Margin of safety" & vbLf & "Segment margin" & vbLf & _
"Responsibility accounting" & vbLf & "Cost centre" & vbLf & "Profit centre" & vbLf & _
"Investment centre" & vbLf & "Return on investment (ROI)" & vbLf & _
"Residual income" & vbLf & "Economic value added (EVA)" & vbLf & _
"Balanced scorecard" & vbLf & "Key performance indicator (KPI)" & vbLf & _
"Manufacturing cycle efficiency (MCE)" & vbLf & "Throughput time" & vbLf & _
"Delivery cycle time"

Private Const IFRS_RAW_9 As String = _
"STANDARD-COST VARIANCES" & vbLf & _
"Direct materials price variance" & vbLf & "Direct materials usage variance" & vbLf & _
"Direct labour rate variance" & vbLf & "Direct labour efficiency variance" & vbLf & _
"Variable overhead spending variance" & vbLf & "Variable overhead efficiency variance" & vbLf & _
"Fixed overhead budget variance" & vbLf & "Fixed overhead volume variance" & vbLf & _
"Sales price variance" & vbLf & "Sales volume variance" & vbLf & _
"Relevant cost" & vbLf & "Opportunity cost" & vbLf & "Sunk cost" & vbLf & _
"Product cost" & vbLf & "Period cost" & vbLf & "Differential cost" & vbLf & _
"Incremental cost" & vbLf & "Marginal cost"

'----------- 1B ¥ Convert the blocks to a single 0-based array -------------
Public Function IFRS_AccountsArr() As Variant
    Dim allTxt As String
    allTxt = IFRS_RAW_1 & vbLf & IFRS_RAW_2 & vbLf & IFRS_RAW_3 & vbLf & _
             IFRS_RAW_4 & vbLf & IFRS_RAW_5 & vbLf & IFRS_RAW_6 & vbLf & _
             IFRS_RAW_7 & vbLf & IFRS_RAW_8 & vbLf & IFRS_RAW_9
    IFRS_AccountsArr = Split(allTxt, vbLf)
End Function

'----------- 1C ¥ Build / refresh the hidden sheet & named range ----------
Public Sub BuildHiddenSheet()
    Dim wb As Workbook, ws As Worksheet, arr As Variant, i As Long, lastRow As Long
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets("IFRSLibrary")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "IFRSLibrary"
    End If
    
    ws.Visible = xlSheetVeryHidden          'cannot be un-hidden via UI
    ws.Cells.Clear
    
    arr = IFRS_AccountsArr
    For i = LBound(arr) To UBound(arr)
        ws.Cells(i + 1, 1).Value = arr(i)
    Next i
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    'Create/replace a global named range usable in DV lists
    On Error Resume Next
    wb.Names("IFRS_Accounts").Delete
    On Error GoTo 0
    wb.Names.Add Name:="IFRS_Accounts", _
                 RefersTo:=ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1))
End Sub

'----------- 1D ¥ Helper Ð apply the drop-down to any range  --------------
Public Sub ApplyIFRSDropDown(ByVal tgt As Range)
    BuildHiddenSheet                         'makes sure the list exists
    With tgt.Validation
        .Delete                              'clear previous DV
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=IFRS_Accounts"
        .IgnoreBlank = True
        .InCellDropdown = True               'auto-complete available in modern Excel
    End With
End Sub




