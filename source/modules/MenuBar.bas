Attribute VB_Name = "MenuBar"
Option Explicit

Public Sub InitMenuBar()
    ' insert menu options into menu

    Dim MenuBar As MenuBar
    Set MenuBar = frmMain.MenuBar

    ' build File menu
    Dim mnuFile As MenuBarLib.MenuBarMenu
    Set mnuFile = MenuBar.Controls.AddMenu("File", "mnuFile")
    
    mnuFile.Items.Add , "mnuFileNew", "New"
    mnuFile.Items.Add , "mnuFileOpen", "Open..."
    mnuFile.Items.Add , , , mbrMenuSeparator
    mnuFile.Items.Add , "mnuFileSave", "Save"
    mnuFile.Items.Add , "mnuFileSaveAs", "Save As..."
    mnuFile.Items.Add , , , mbrMenuSeparator
    mnuFile.Items.Add , "mnuFileExit", "Exit"
    
    ' build Edit menu
    Dim mnuEdit As MenuBarLib.MenuBarMenu
    Set mnuEdit = MenuBar.Controls.AddMenu("Edit", "mnuEdit")
    
    mnuEdit.Items.Add , "mnuEditUndo", "Undo"
    mnuEdit.Items.Add , , , mbrMenuSeparator
    mnuEdit.Items.Add , "mnuEditCut", "Cut"
    mnuEdit.Items.Add , "mnuEditCopy", "Copy"
    mnuEdit.Items.Add , "mnuEditPaste", "Paste"
    mnuEdit.Items.Add , , , mbrMenuSeparator
    mnuEdit.Items.Add , "mnuEditDelete", "Delete"
    mnuEdit.Items.Add , "mnuEditSelectAll", "Select All"
    
    ' build View Menu
    Dim mnuView As MenuBarLib.MenuBarMenu
    Set mnuView = MenuBar.Controls.AddMenu("View", "mnuView")
    
    mnuView.Items.Add , "mnuViewWordlist", "Word List"
    mnuView.Items.Add , , , mbrMenuSeparator
    mnuView.Items.Add , "mnuViewTag", "Tag Toolbar"
    mnuView.Items.Add , "mnuViewAttrib", "Attribute Toolbar"
    mnuView.Items.Add , , , mbrMenuSeparator
    mnuView.Items.Add , "mnuViewWordwrap", "Word Wrap"
    mnuView.Items.Add , "mnuViewFont", "Font..."
    mnuView.Items.Add , , , mbrMenuSeparator
    mnuView.Items.Add , "mnuViewOptions", "Options..."
        
    ' add View > Word List menu
    Dim mnuViewWordlist As MenuBarLib.MenuBarMenu
    Set mnuViewWordlist = mnuView.Items("mnuViewWordlist")
   
    mnuViewWordlist.SubItems.Add , "mnuViewWordlistNew", "New..."
    mnuViewWordlist.SubItems.Add , , , mbrMenuSeparator
    
    ' build Help menu
    Dim mnuHelp As MenuBarLib.MenuBarMenu
    Set mnuHelp = MenuBar.Controls.AddMenu("Help", "mnuHelp")
    
    mnuHelp.Items.Add , "mnuHelpAbout", "About"
    
    ' load custom menu icons
    Dim img As ImageList
    Set img = frmMain.ImageList
    
    img.Add App.Path & "\btnCut.bmp"
    img.Add App.Path & "\btnCopy.bmp"
    img.Add App.Path & "\btnPaste.bmp"
    img.Add App.Path & "\btnWellformed.bmp"
    
    ' assign ImageList to the menu bar
    MenuBar.ImageList = img.hImageList

    ' add menu buttons
    Dim btn As MenuBarButton
    
    Set btn = MenuBar.Controls.AddButton("btnCut")
    btn.ToolTip = "Cut"
    btn.Image = 1
    
    Set btn = MenuBar.Controls.AddButton("btnCopy")
    btn.ToolTip = "Copy"
    btn.Image = 2
    
    Set btn = MenuBar.Controls.AddButton("btnPaste")
    btn.ToolTip = "Paste"
    btn.Image = 3
    
    Set btn = MenuBar.Controls.AddButton("btnWellformed")
    btn.ToolTip = "Check well-formed"
    btn.Image = 4
    
    ' initialise menu state
    mnuView.Items("mnuViewTag").Checked = True

End Sub


Public Sub MenuBarClick(ByVal Item As MenuBarLib.Item)

'   Description:
'       Handles clicking of main menu items


    Select Case Item.Key
        Case "mnuFileNew"
            Call CloseFile
        Case "mnuFileOpen"
            OpenFile
        Case "mnuFileSave"
            Call SaveFile
        Case "mnuFileSaveAs"
            Call SaveFileAs
        Case "mnuFileExit"
            EndApp
        Case "mnuEditUndo"
            EditUndo
            pInitEditMenu
        Case "mnuEditCut"
            EditCut
            pInitEditMenu
        Case "mnuEditCopy"
            EditCopy
            pInitEditMenu
        Case "mnuEditPaste"
            EditPaste
            pInitEditMenu
        Case "mnuEditDelete"
            EditDelete
            pInitEditMenu
        Case "mnuEditSelectAll"
            EditSelectAll
            pInitEditMenu
        Case "mnuViewTag"
            Item.Checked = ToggleToolbar(0)
        Case "mnuViewAttrib"
            Item.Checked = ToggleToolbar(1)
        Case "mnuViewWordwrap"
            Item.Checked = ToggleWordWrap()
        Case "mnuViewFont"
            SetFont
        Case "mnuViewOptions"
            frmOptions.Show
        Case "mnuHelpAbout"
            'show the about form
            frmAbout.Show
        Case Else
        
            ' switch word list
            If InStr(Item.Key, "mnuViewWordlist") Then
                SwitchWordList Item.Caption
            End If
        
    End Select

End Sub

Public Sub MenuButtonClick(ByVal button As MenuBarLib.Item)

'   Description:
'       Handles clicks of main menu bar buttons.

    Select Case button.Key
        Case "btnCut"
            EditCut
            pInitEditMenu
        Case "btnCopy"
            EditCopy
            pInitEditMenu
        Case "btnPaste"
            EditPaste
            pInitEditMenu
        Case "btnWellformed"
            CheckWellformedDoc MainTextBox, True
    End Select

End Sub



Public Sub pInitEditMenu()
  
    InitEditMenu frmMain.MenuBar.Controls(2).Items("mnuEditUndo"), _
        frmMain.MenuBar.Controls(2).Items("mnuEditCut"), _
        frmMain.MenuBar.Controls(2).Items("mnuEditCopy"), _
        frmMain.MenuBar.Controls(2).Items("mnuEditPaste"), _
        frmMain.MenuBar.Controls(2).Items("mnuEditDelete"), _
        frmMain.MenuBar.Controls(2).Items("mnuEditSelectAll")

End Sub


Public Sub pDisableEditMenu()
  
    DisableEditMenu frmMain.MenuBar.Controls(2).Items("mnuEditUndo"), _
        frmMain.MenuBar.Controls(2).Items("mnuEditCut"), _
        frmMain.MenuBar.Controls(2).Items("mnuEditCopy"), _
        frmMain.MenuBar.Controls(2).Items("mnuEditPaste"), _
        frmMain.MenuBar.Controls(2).Items("mnuEditDelete"), _
        frmMain.MenuBar.Controls(2).Items("mnuEditSelectAll")

End Sub

