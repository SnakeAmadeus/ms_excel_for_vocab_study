VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KotobaOboeru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim colsToHide As Variant

Public Sub registerKeyStrokes()
    
    'Columns that will be hidden when macro is activated:
    colsToHide = Array(3, 4)
    
    Application.OnKey "+{F1}", "enableMacro"
    
    If Selection.Column = 1 Then
        Application.OnKey " ", "markCell"
    Else
        Application.OnKey " ", ""
    End If
    
End Sub

Public Function initColor(ByVal c As Long)
    
    Dim headerRow As Integer
    headerRow = 1
    
    Dim col As Variant
    For Each col In colsToHide:
        Dim contentRows As Range
        'Reference: https://stackoverflow.com/a/14390619
        Dim lastRow As Range
        Set lastRow = Cells.Find("*", [A1], , , xlByRows, xlPrevious)
        Set contentRows = Range(Cells(headerRow + 1, col), Cells(lastRow.Row, col))
        'Debug.Print "contentRows: " & contentRows.Address
        
        contentRows.Font.color = c
    Next

End Function

Private Sub Workbook_Activate()
    
    'Debug.Print "Current Selected: ("; prev.Cells.Column & ", " & prev.Cells.Row & ")."
    
    Call registerKeyStrokes
    
    macroEnabled = False
    
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    
    'Debug.Print "Current Selected: ("; prev.Cells.Column & ", " & prev.Cells.Row & ")."
    
    Call registerKeyStrokes

End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)
    
    If Not macroEnabled Then Exit Sub

    If Target.Cells.Count <> 1 Then Exit Sub
    
    Call registerKeyStrokes
    
    Application.ScreenUpdating = False
    initColor (RGB(250, 249, 246))
    Dim col As Variant
    For Each col In colsToHide
        Cells(Target.Row, col).Font.color = RGB(0, 0, 0)
    Next
    Application.ScreenUpdating = True
    
End Sub


