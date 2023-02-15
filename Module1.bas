Attribute VB_Name = "Module1"
'Module1: Switch On/Off the Macro by pressing {F12}
Global macroEnabled As Boolean

Sub enableMacro()
    
    macroEnabled = Not macroEnabled
    
    If macroEnabled Then Call KotobaOboeru.registerKeyStrokes

    Dim oldStatusBar
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    If macroEnabled Then
        Call KotobaOboeru.initColor(RGB(250, 249, 246))
        Application.StatusBar = "Macro is ON! Rendering worksheet..."
    Else
        Call KotobaOboeru.initColor(RGB(0, 0, 0))
        Application.StatusBar = "Macro is OFF, Restoring worksheet..."
    End If
    
    Debug.Print "Macro ON? : " & CStr(macroEnabled)
    
    Application.Wait (Now + TimeValue("0:00:02"))
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    
End Sub
