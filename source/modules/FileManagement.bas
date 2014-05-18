Attribute VB_Name = "FileManagement"
Option Explicit

Public FileName
Public FileModified

FileName = False
FileModified = False


Public Function CloseFile() As Boolean

    Dim OkToClose As Boolean
    OkToClose = True
    
    ' prompt for save
    If FileModified = True Then
        If MsgBox("Save file before closing?", 36, "PocketMarkup") = vbYes Then
            OkToClose = SaveFile
        End If
    End If

    If OkToClose = True Then
        ' clear text
        MainTextBox.Text = ""
        ' clear filename
        FileName = False
        FileModified = False
        
        CloseFile = True
    Else
        CloseFile = False
    End If
    
End Function


Public Sub OpenFile()
    
    If CloseFile Then
               
        FileName = GetFileName("open")
               
        If Not FileName = "" Then
           
            ' load file into text box
            MainTextBox.Text = LoadText(FileName)
            
            If GetOptionValue("Check XML Load") = "True" Then
            
                CheckWellformedDoc MainTextBox, False
            
            End If
            
        End If
    
    End If
    
End Sub


Public Sub OpenInitialFile(Path)

'   Description:
'       Opens a file without prompting for a name, or
'       prompting to save the current file.  Initially created
'       specifically to open config files for user debugging.

    FileName = Path
               
    If Not FileName = "" Then
           
        ' load file into text box
        MainTextBox.Text = LoadText(FileName)
        CheckWellformedDoc MainTextBox, False
        
    End If
    
End Sub


Public Function SaveFile() As Boolean
        
    Dim Form As Control
    Set Form = frmMain
           
    ' check for current filename
    If FileName = False Then
        SaveFile = SaveFileAs()
    Else
        SaveText MainTextBox, FileName
        FileModified = False
        SaveFile = True
    End If
    
End Function

Public Function SaveFileAs() As Boolean
    
    ' check for current filename
    FileName = GetFileName("save")
   
    If FileName = "" Then
        FileName = False
        SaveFileAs = False
    Else
        SaveText MainTextBox, FileName
        FileModified = False
        SaveFileAs = True
    End If
    
End Function



