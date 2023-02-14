Attribute VB_Name = "Module2"
'Module2: Marking functionality by press spacebar on any cells in Column 1
Sub markCell()
    
    Dim sel As Range
    
    Set sel = Selection
    
    If sel.Cells.Count <> 1 Then Exit Sub
    
    If sel.Column <> 1 Then Exit Sub
    
'    If sel.Interior.color = RGB(255, 255, 255) Or sel.Cells.Interior.ColorIndex = 0 Then
'
'        sel.Interior.color = RGB(255, 0, 0)
'
'    Else
'        sel.Cells.Interior.ColorIndex = 0
'
'    End If

    If sel.Value = ChrW(&H274C) Then
        
        sel.Value = ""
    
    Else
        
        sel.Value = ChrW(&H274C)
        
        sel.HorizontalAlignment = xlCenter
        
    End If
        
    
End Sub
