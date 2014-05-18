Attribute VB_Name = "TextFunctions"
Option Explicit

Public Function PrepTagName(TagName As String) As String
    ' converts user input to a valid tag name
    
    TagName = Trim(TagName) ' trim whitespace
    TagName = Replace(TagName, " ", "") ' remove spaces
    
    ' remove the lowercase for now
    ' set it somewhere in a prefs file
    'TagName = LCase(TagName)  ' convert to lowercase
    
    PrepTagName = TagName
    
End Function

