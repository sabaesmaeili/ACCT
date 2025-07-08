Attribute VB_Name = "modIndent"

Option Explicit

'========================================
'  Indent selection by one level (~4 “spaces”)
'  Call repeatedly to keep pushing right
'========================================
Public Sub IndentPlus4()
Attribute IndentPlus4.VB_ProcData.VB_Invoke_Func = "P\n14"

    Dim c As Range
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    For Each c In Selection.Cells
        'Excel lets you go from 0 to 15 indent levels
        If c.IndentLevel < 15 Then c.IndentLevel = c.IndentLevel + 1
    Next c

End Sub

'Optional: quick “un?indent” companion
Public Sub IndentMinus4()
Attribute IndentMinus4.VB_ProcData.VB_Invoke_Func = "M\n14"

    Dim c As Range
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    For Each c In Selection.Cells
        If c.IndentLevel > 0 Then c.IndentLevel = c.IndentLevel - 1
    Next c

End Sub


