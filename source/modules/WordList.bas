Attribute VB_Name = "WordList"
'   Module WordList.bas
'
'   Description:
'       Handles lists of tag and attribute names stored in an
'       external XML file
'
'   Public Functions:
'       InitWordList()
'       SwitchWordList(LangName As String)
'
'   Written by:
'       tim@yaffle.org
'
'   Version:
'       1.0.2
'       19th February 2002


Option Explicit

Dim XmlWordlist As DOMDocument
Dim WordListActive As Boolean
Dim WordFilePath As String
Dim CurrentLanguage As String

WordListActive = False


Public Sub InitWordList()

'   Description:
'       Initialises the word list from the XML file.  Calls the functions
'       to populate the view menu and combos.

    WordFilePath = Replace(GetOptionValue("Word File Path"), "[App]", App.Path)
    CurrentLanguage = GetOptionValue("Default Language")

    If GetOptionValue("Load Word List") = "True" Then

        ' load the XML from file
        Set XmlWordlist = LoadXml(WordFilePath)
    
        If XmlWordlist.parseError Then
    
            ' check that the word list is well-formed XML
            ' if not, offer to open it for editing
    
            If MsgBox("The word file is not well-formed XML.  Would you like to open it now for editing?", 20, "PocketMarkup") = vbYes Then
                OpenInitialFile WordFilePath
            End If
    
        Else
    
            ' otherwise, activate the word lists
            WordListActive = True
    
            GetMarkupLanguages
    
            GetList CurrentLanguage, "tags", frmMain.cboNewTag
            GetList CurrentLanguage, "attribs", frmMain.cboAttributeName
    
        End If

    End If

End Sub

Public Sub SaveWordList()

'   Description:
'

    If WordListActive Then
        
        If GetOptionValue("Save Word List") = "True" Then
    
            StoreList CurrentLanguage, "tags", frmMain.cboNewTag
            StoreList CurrentLanguage, "attribs", frmMain.cboAttributeName
    
            SaveXml XmlWordlist, WordFilePath
        
        End If

    End If

End Sub

Public Sub SwitchWordList(LangName As String)

'   Description:
'       Switches the current wordlist.

    If WordListActive Then

        If Not LangName = CurrentLanguage Then

            If LangName = "New..." Then
                ' create a new custom word list
                NewWordList
            Else
                ' store the lists, which may have been modified, in the XML object
                StoreList CurrentLanguage, "tags", frmMain.cboNewTag
                StoreList CurrentLanguage, "attribs", frmMain.cboAttributeName

                CurrentLanguage = LangName

                ' call the new word lists from the XML
                GetList LangName, "tags", frmMain.cboNewTag
                GetList LangName, "attribs", frmMain.cboAttributeName

                ' put a check mark next to the language name that was clicked
                Dim mnu
                Set mnu = frmMain.MenuBar.Controls("mnuView").Items("mnuViewWordlist")

                Dim i As Integer

                For i = 1 To mnu.SubItems.Count
                    If mnu.SubItems.Item(i).Caption = LangName Then
                        mnu.SubItems.Item(i).Checked = True
                    Else
                        mnu.SubItems.Item(i).Checked = False
                    End If
                Next

            End If
        End If
    End If

End Sub


Private Sub NewWordList()

'   Description:
'       Adds a new word list to the XML.  Prompts for a name, inserts
'       XML nodes, and adds it to the View > Word List menu.  Checks for
'       an existing word list of the same name.

    ' prompt for a name for the new language
    Dim NewLangName As String
    NewLangName = InputBox("Enter a name for the new word list:", "Add word list")

    ' check language does not exist
    Dim LangExists As Boolean
    LangExists = False
    
    ' query XML for a list of the languages
    Dim XmlNodeList As IXMLDOMNodeList1
    Set XmlNodeList = XmlWordlist.selectNodes("/wordlist/language/@name")

    Dim XmlNode As IXMLDOMNode

    For Each XmlNode In XmlNodeList
        ' check each one is not same name as new list
        If XmlNode.nodeValue = NewLangName Then
            LangExists = True
        End If
    Next
    
    If LangExists Then
        
        ' offer to switch to the existing list
        If MsgBox(NewLangName & " already exists in the wordlist.  Activate it now?", vbYesNo, "PocketMarkup") Then
            SwitchWordList NewLangName
        End If
        
    Else
        
        ' get reference to View > Word List menu
        Dim mnuViewWordlist As MenuBarLib.MenuBarMenu
        Set mnuViewWordlist = frmMain.MenuBar.Controls("mnuView").Items("mnuViewWordlist")
    
        ' add language name to menu
        mnuViewWordlist.SubItems.Add , "mnuViewWordlist" & mnuViewWordlist.SubItems.Count, NewLangName
        mnuViewWordlist.SubItems.Item(mnuViewWordlist.SubItems.Count - 1).Checked = True
    
        ' get a reference to the wordlist node in the XML
        Dim ListNode As IXMLDOMNode
        Set ListNode = XmlWordlist.selectSingleNode("/wordlist")
    
        ' create the new language node, with tags an attribs
        AppendChildNode XmlWordlist, ListNode, "language"
        AddAttribute XmlWordlist, ListNode.lastChild, "name", NewLangName
        AppendChildNode XmlWordlist, ListNode.lastChild, "tags"
        AppendChildNode XmlWordlist, ListNode.lastChild, "attribs"
    
        ' call the swtich to the new list
        SwitchWordList NewLangName
        
    End If

End Sub


Private Sub GetMarkupLanguages()

'   Description:
'       Gets a list of languages and populates the View > Word List menu

    ' query XML for a list of the languages
    Dim XmlNodeList As IXMLDOMNodeList1
    Set XmlNodeList = XmlWordlist.selectNodes("/wordlist/language/@name")

    ' get reference to View > Word List menu
    Dim mnuViewWordlist As MenuBarLib.MenuBarMenu
    Set mnuViewWordlist = frmMain.MenuBar.Controls("mnuView").Items("mnuViewWordlist")

    ' add language names to menu
    Dim XmlNode As IXMLDOMNode

    For Each XmlNode In XmlNodeList
        mnuViewWordlist.SubItems.Add , "mnuViewWordlist" & XmlNode.nodeValue, XmlNode.nodeValue
    Next

End Sub


Private Sub GetList(LangName As String, ListName As String, ListCombo As ComboBox)

'   Description:
'       Queries the XML for values according to LangName and ListName,
'       and populates ListCombo with values

    ' query XML for a list of the languages included in the tag list
    Dim XmlNodeList As IXMLDOMNodeList1
    Set XmlNodeList = XmlWordlist.selectSingleNode("/wordlist/language[ @name = '" & LangName & "' ]/" & ListName).childNodes

    ' clear the combo
    ListCombo.Clear

    ' loop through the values, adding each one to the combo
    Dim XmlNode As IXMLDOMNode

    For Each XmlNode In XmlNodeList
        AddToXmlCombo ListCombo, XmlNode.Text
    Next

End Sub


Private Sub StoreList(LangName As String, ListName As String, ListCombo As ComboBox)

'   Description:
'       Insert contents of word list combo back into the XML.
'       Allows the custom word list to be written back to the XML file.

    ' get a reference to the appropriate "tags" or "attribs" element
    Dim XmlWordSet As IXMLDOMNode
    Set XmlWordSet = XmlWordlist.selectSingleNode("/wordlist/language[ @name = '" & LangName & "' ]/" & ListName)

    ' remove all the child nodes
    While XmlWordSet.hasChildNodes()
        XmlWordSet.removeChild XmlWordSet.firstChild
    Wend

    ' set the name of the new child elements
    Dim ElName As String

    Select Case ListName
        Case "tags"
            ElName = "tag"
        Case "attribs"
            ElName = "attrib"
        Case "entities"
            ElName = "entity"
    End Select

    ' loop through the combo items and store each value in the XML
    Dim i As Integer

    For i = 0 To ListCombo.ListCount - 1

        ' create a new node
        Dim NewTag As IXMLDOMNode
        Set NewTag = XmlWordlist.createElement(ElName)

        ' give it the value from the combo
        NewTag.Text = ListCombo.List(i)

        ' append the new node to the list
        XmlWordSet.appendChild NewTag

    Next

End Sub


