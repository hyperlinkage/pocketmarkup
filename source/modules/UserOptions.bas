Attribute VB_Name = "UserOptions"
Option Explicit

Dim OptionsXml As DOMDocument
Dim OptionsFile As String
OptionsFile = App.Path & "\options.xml"

Public MainTextBox As TextBox

Dim WordWrapActive As Boolean


Public Sub InitUserPreferences()

'   Description:
'       Loads user options from XML file, and initialises
'       the necessary values.

    ' Load the user options XML from file.
    Set OptionsXml = LoadXml(OptionsFile)
    
    If OptionsXml.parseError Then

        ' The options XML file must be well-formed.
        MsgBox "The options file is not well-formed XML.  PocketMarkup cannot start.", vbCritical, "PocketMarkup"
        App.End

    Else
        
        InitWordWrap
        InitFont
        InitWordList
        InitToolVisibility

    End If
    
End Sub


Public Sub SaveUserPreferences()

'   Description:
'       Writes the user options back to file.

   SaveWordList
   
   SaveXml OptionsXml, OptionsFile

End Sub


Public Function GetOptionValue(OptionName As String) As String

'   Description:
'       Gets the value of a named option by querying the
'       user options stored in the global OptionsXml object.
'
'   Returns:
'       The string value of the named option.

    Dim OptionNode As IXMLDOMNode
    Set OptionNode = OptionsXml.selectSingleNode("/options/option[ @name = '" & OptionName & "' ]")
    
    GetOptionValue = OptionNode.Text

End Function


Public Sub SetOptionValue(OptionName As String, OptionValue As String)

'   Description:
'       Sets the value of a named option by adding the node to the
'       OptionsXml object.  If the option doesn't exist, it is appended
'       to the list.

    Dim OptExists As Boolean
    OptExists = False
    
    ' Query XML for a list of the options
    Dim XmlNodeList As IXMLDOMNodeList1
    Set XmlNodeList = OptionsXml.selectNodes("/options/option/@name")

    Dim XmlNode As IXMLDOMNode

    For Each XmlNode In XmlNodeList
        ' Check each one is not same name as new option
        If XmlNode.nodeValue = OptionName Then
            OptExists = True
        End If
    Next
    
    Dim OptionNode As IXMLDOMNode
    
    If OptExists Then
    
        ' Of it exists, update the value
        Set OptionNode = OptionsXml.selectSingleNode("/options/option[ @name = '" & OptionName & "']")
        OptionNode.Text = OptionValue
    
    Else
    
        ' Otherwise, add the node with appropriate values
        Set OptionNode = OptionsXml.selectSingleNode("/options")
        
        AppendChildNode XmlWordlist, OptionNode, "option"
        AddAttribute XmlWordlist, OptionNode.lastChild, "name", OptionName
        OptionNode.lastChild.Text = OptionValue
        
    End If
    
End Sub


' Functions for handling of specific user options:

Private Sub InitWordWrap()
        
'   Description:
'       Sets the initial state of word wrap by getting the option value and
'       setting a global reference to the appropriate TextBox to use.
       
    ' Initialise the main text box, depending on the user word wrap preference.
    If GetOptionValue("Word Wrap") = "True" Then
        WordWrapActive = True
        Set MainTextBox = frmMain.txtMainOn
    Else
        WordWrapActive = False
        Set MainTextBox = frmMain.txtMainOff
    End If
    
    frmMain.MenuBar.Controls("mnuView").Items("mnuViewWordwrap").Checked = WordWrapActive
        
End Sub


Public Function ToggleWordWrap() As Boolean

'   Description:
'       Toggles word wrap on/off.
'       Because the scrollbar properties of TextBox are read only,
'       this has to be achieved by switching between two boxes and
'       storing a reference to the active one in a global variable.
'
'   Returns:
'       Whether word wrap has been enabled or disabled.

    ' Temporarily store a reference to the old text box
    Dim OldTextBox As TextBox
    Set OldTextBox = MainTextBox

    ' Switch the global reference to the main text box
    If WordWrapActive Then
    
        WordWrapActive = False
        Set MainTextBox = frmMain.txtMainOff
    Else
    
        WordWrapActive = True
        Set MainTextBox = frmMain.txtMainOn
    End If
    
    ' Copy the text over to the new box
    MainTextBox.Text = OldTextBox.Text
    
    ' Copy selection position over to new box
    Dim SelStrt, SelLnth As Integer
    SelLnth = OldTextBox.SelLength
    SelStrt = OldTextBox.SelStart
    MainTextBox.SelStart = SelStrt
    MainTextBox.SelLength = SelLnth
    
    ' Show the new box, hide the old one
    OldTextBox.Visible = False
    MainTextBox.Visible = True
    MainTextBox.SetFocus
    
    MainFormAlign
    InitFont
    
    ToggleWordWrap = WordWrapActive
        
End Function


Private Sub InitToolVisibility()

'   Description:
'       Initialises the visibility of the tag and attribute
'       toolbars.

    If GetOptionValue("Tag Toolbar") = "True" Then
        Toolbars(0).Visible = True
    Else
        Toolbars(0).Visible = False
    End If
    
    If GetOptionValue("Attribute Toolbar") = "True" Then
        Toolbars(1).Visible = True
    Else
        Toolbars(1).Visible = False
    End If
    
End Sub


Public Sub SetFont()

    frmMain.CommonDialog.FontName = GetOptionValue("Font Name")
    frmMain.CommonDialog.FontSize = GetOptionValue("Font Size")
    frmMain.CommonDialog.FontBold = GetOptionValue("Font Bold")
    frmMain.CommonDialog.FontItalic = GetOptionValue("Font Italic")

    frmMain.CommonDialog.ShowFont
    
    SetOptionValue "Font Name", frmMain.CommonDialog.FontName
    SetOptionValue "Font Size", frmMain.CommonDialog.FontSize
    SetOptionValue "Font Bold", frmMain.CommonDialog.FontBold
    SetOptionValue "Font Italic", frmMain.CommonDialog.FontItalic

    InitFont

End Sub

Public Sub InitFont()

    MainTextBox.FontName = GetOptionValue("Font Name")
    MainTextBox.FontSize = GetOptionValue("Font Size")
    MainTextBox.FontBold = GetOptionValue("Font Bold")
    MainTextBox.FontItalic = GetOptionValue("Font Italic")

End Sub
