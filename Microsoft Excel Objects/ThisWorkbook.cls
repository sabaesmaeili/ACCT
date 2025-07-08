VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================
' Workbook: ThisWorkbook (Workbook Events)
' Purpose: Assign and clean up keyboard shortcuts for all formatting routines
'===============================================
Option Explicit

Private Sub Workbook_Open()
    ' Ensure the IFRS list is built on file open
    modIFRSLibrary.BuildHiddenSheet
    
    ' Assign keyboard shortcuts:
    ' CTRL+SHIFT+I: IFRS dropdown
    Application.OnKey "^+I", "IFRSDropDownSelection"
    ' CTRL+SHIFT+U: Remove data validation
    Application.OnKey "^+U", "RemoveDataValidationFromSelection"
    ' CTRL+SHIFT+B: IFRS table formatting
    Application.OnKey "^+B", "FormatTable"
    ' CTRL+SHIFT+J: Journal formatting
    Application.OnKey "^+J", "FormatJournal"
    ' CTRL+SHIFT+T: T?account formatting
    Application.OnKey "^+T", "FormatTAccount"
    ' CTRL+SHIFT+L: Ledger formatting
    Application.OnKey "^+L", "FormatLedger"
    ' CTRL+SHIFT+S: Statement formatting
    Application.OnKey "^+S", "FormatStatement"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clear all custom OnKey assignments
    Application.OnKey "^+I"
    Application.OnKey "^+U"
    Application.OnKey "^+B"
    Application.OnKey "^+J"
    Application.OnKey "^+T"
    Application.OnKey "^+L"
    Application.OnKey "^+S"
End Sub

