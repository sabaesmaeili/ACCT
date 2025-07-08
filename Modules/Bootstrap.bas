Attribute VB_Name = "modAccountingBootstrap"
Option Explicit
'===============================
'  Accounting template bootstrap
'===============================
Public Sub SetupAccountingTemplate()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim acctFmt As String

    Set wb = ThisWorkbook

    '? NEW format: no decimals, no red text
    acctFmt = "$#,##0_);($#,##0)"

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each ws In wb.Worksheets
        '1. Hide gridlines
        ws.Activate
        ActiveWindow.DisplayGridlines = False

        '2. Apply font + number format
        With ws.Cells
            .Font.Name = "Times New Roman"
            .Font.Size = 11
            .NumberFormat = acctFmt
        End With
    Next ws

    wb.Worksheets(1).Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Template initialised � all sheets ready.", vbInformation

End Sub

'===============================
'  Shortcut setup & handler
'===============================
Public Sub IFRSDropDownSelection()
    Dim sel As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    ' Make sure the hidden sheet + named range exist
    modIFRSLibrary.BuildHiddenSheet
    ' Apply the drop?down to whatever you�ve selected
    modIFRSLibrary.ApplyIFRSDropDown sel
End Sub

Public Sub RemoveDataValidationFromSelection()
    '�� Deletes any validation (including dropdowns) from the current Selection
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.Validation.Delete
End Sub

