Attribute VB_Name = "modAccountingTemplate"
Option Explicit
'---------------------------------------------------------------
' Boot?strap + formatting helpers
'---------------------------------------------------------------
Public Sub InitializeTemplate()
Attribute InitializeTemplate.VB_ProcData.VB_Invoke_Func = "I\n14"
    BuildHiddenSheet                         'in modIFRSLibrary
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> IFRSLibSheetName Then ApplyDefaultFormatting ws
    Next ws
End Sub

Public Sub ApplyDefaultFormatting(ByVal ws As Worksheet)
    With ws.Cells
        .Font.Name = DefaultFontName
        .Font.Size = DefaultFontSize
        .NumberFormat = DefaultNumFmt
    End With
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

'---------------------------------------------------------------
'   IFRS dropdown for current selection
'---------------------------------------------------------------
Public Sub IFRSDropDownSelection()
Attribute IFRSDropDownSelection.VB_ProcData.VB_Invoke_Func = "D\n14"
    Dim sel As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    ' Make sure the hidden sheet + named range exist
    modIFRSLibrary.BuildHiddenSheet
    ' Apply the drop?down to whatever you’ve selected
    modIFRSLibrary.ApplyIFRSDropDown sel
End Sub
'---------------------------------------------------------------
'   Keyboard shortcuts
'---------------------------------------------------------------
Public Sub AssignShortcuts()
    With Application
        .OnKey "^+I", "IFRSDropDownSelection"           'Ctrl?Shift?I  ? dropdown
        .OnKey "^+U", "RemoveDataValidationFromSelection"
        .OnKey "^+B", "FormatTable"
        .OnKey "^+J", "FormatJournal"
        .OnKey "^+T", "FormatTAccount"
        .OnKey "^+L", "FormatLedger"
        .OnKey "^+S", "FormatStatement"
    End With
End Sub

Public Sub RemoveShortcuts()
    With Application
        .OnKey "^+I", ""
        .OnKey "^+U", ""
        .OnKey "^+B", ""
        .OnKey "^+J", ""
        .OnKey "^+T", ""
        .OnKey "^+L", ""
        .OnKey "^+S", ""
    End With
End Sub


