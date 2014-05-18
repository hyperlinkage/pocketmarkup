Attribute VB_Name = "AppStartEnd"
Option Explicit


Public Sub StartApp()
    
    InitMenuBar
    InitToolbars
    InitUserPreferences

End Sub


Public Sub EndApp()
      
    If CloseFile() = True Then
    
        SaveUserPreferences
        App.End
        
    End If
    
End Sub
