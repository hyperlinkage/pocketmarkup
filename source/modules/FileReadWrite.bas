Attribute VB_Name = "FileReadWrite"
'   Module FileReadWrite.bas
'
'   Description:
'       A set of public functions to handle reading and writing to text file
'
'   Public Functions:
'       GetFileName(FileAction As String) As String
'       LoadText(FileName As String) As String
'       SaveText(FileData As String, FileName As String)
'       LoadXml(FileName As String) As DOMDocument
'       SaveXml(XmlDoc As DOMDocument, FileName As String)
'
'   Written by:
'       tim@yaffle.org
'
'   Version:
'       1.0.0
'       18th February 2002


Option Explicit


Public Function GetFileName(FileAction As String) As String

'   Description:
'       Accepts either "open" or "save" as FileAction.  Displays the
'       CommonDialog box and returns the full filename as string.
    
    ' get a reference to the CommonDialog control
    Dim CommonDialog As CommonDialog
    Set CommonDialog = frmMain.CommonDialog
    
    Dim FileTypes As String
    
    Select Case FileAction
        Case "open"
            ' set the CommonDialog to only display certain filetypes
            FileTypes = "HTML Documents (*.html, *.htm)|*.html;*.htm|XML Files (*.xml)|*.xml|XSL Stylesheets (*.xsl)|*.xsl|All Files (*.*)|*.*"
            CommonDialog.Filter = FileTypes
            CommonDialog.ShowOpen
        Case "save"
            FileTypes = "HTML Document (*.html)|*.html|XML File (*.xml)|*.xml|XSL Stylesheet (*.xsl)|*.xsl|All Files (*.*)|*.*"
            CommonDialog.Filter = FileTypes
            CommonDialog.ShowSave
    End Select
    
    ' return the filename
    GetFileName = CommonDialog.FileName
    
End Function


Public Function LoadText(FileName As String) As String

'   Description:
'       Reads text from the specified file, returns as string.

    Dim FileData As String

    If Not FileName = "" Then
        
        ' get a reference to the file input/output control
        Dim FileControl As FileControl
        Set FileControl = frmMain.FileControl
        
        ' display spinning timer thing on screen
        Screen.MousePointer = 11
        
        ' open the file
        FileControl.Open FileName, fsModeInput
        
        While Not FileControl.EOF
            ' read each line from the file, concat with CrLf char
            FileData = FileData & FileControl.LineInputString & vbCrLf
        Wend
        
        'close the file
        FileControl.Close
        
        ' hide spinning timer
        Screen.MousePointer = 0
        
    End If

    ' return the text
    LoadText = FileData

End Function


Public Sub SaveText(FileData As String, FileName As String)

'   Description:
'       Writes the FileData string back to text file specified by FileName.

    Dim FileControl As FileControl
    Set FileControl = frmMain.FileControl
        
    Dim intOfs1
    Dim intOfs2
    
    Screen.MousePointer = 11
    FileControl.Open FileName, fsModeOutput
    
    Do
        intOfs2 = InStr(intOfs1 + 1, FileData, Chr(13))
    
        If intOfs2 = 0 Then
            FileControl.LinePrint Mid(FileData, intOfs1 + 1)
            Exit Do
        Else
            FileControl.LinePrint Mid(FileData, intOfs1 + 1, intOfs2 - intOfs1)
        End If
    
        intOfs1 = intOfs2 + 1
    
    Loop
            
    FileControl.Close
    Screen.MousePointer = 0
    
End Sub


Public Function LoadXml(FileName As String) As DOMDocument
    
'   Description:
'       Loads XML file specified by FileName and returns as DOMDocument object
    
    Dim XmlString As String
    Dim XmlDoc As DOMDocument
    
    ' call function to get XML from the text file
    XmlString = LoadText(FileName)
    
    ' create an instance of the XML parser, and pass it the XML string
    Set XmlDoc = CreateObject("Microsoft.XMLDOM")
    XmlDoc.LoadXml XmlString
    
    ' return XML object
    Set LoadXml = XmlDoc
    
End Function


Public Sub SaveXml(XmlDoc As DOMDocument, FileName As String)
 
'   Description:
'       Writes XML stored in a DOMDocument object to a text file

    SaveText XmlDoc.xml, FileName
    
End Sub
