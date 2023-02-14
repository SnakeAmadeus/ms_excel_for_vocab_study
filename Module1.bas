Attribute VB_Name = "Module1"
'Module1: Switch On/Off the Macro by pressing {F12}
Global macroEnabled As Boolean

Sub enableMacro()
    
    macroEnabled = Not macroEnabled
    
    If macroEnabled Then Call KotobaOboeru.registerKeyStrokes

    If macroEnabled Then
        Call KotobaOboeru.initColor(RGB(250, 249, 246))
    Else
        Call KotobaOboeru.initColor(RGB(0, 0, 0))
    End If
    
    Debug.Print "Macro Status: " & CStr(macroEnabled)
    
End Sub


