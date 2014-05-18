Attribute VB_Name = "ToolbarHandling"
Option Explicit

Public Toolbars(1) As Frame


Public Sub InitToolbars()

    Set Toolbars(0) = frmMain.fraTagToolbar
    Set Toolbars(1) = frmMain.fraAttributeToolbar
   
End Sub


Public Function ToggleToolbar(ToolBar As Integer) As Boolean
    Set ToolBar = Toolbars(ToolBar)
    If ToolBar.Visible = False Then
        ToolBar.Visible = True
    Else
        ToolBar.Visible = False
    End If
    MainFormAlign
    ToggleToolbar = ToolBar.Visible
End Function


Public Sub MainFormAlign()

'   Description:
'       Calls the function to realign the main form elements
'       Accounts for changes in the SIP size and visibility
    
    Dim b As Frame
    Dim NextTop As Integer

    NextTop = frmMain.Height
     
    ' size main frame
    frmMain.fraEditView.Height = frmMain.Height
    frmMain.fraEditView.Width = frmMain.Width
    
    ' align toolbars
    For Each b In Toolbars
        If b.Visible Then
            NextTop = NextTop - b.Height
            b.Top = NextTop
        End If
    Next b
    
    ' size main text box
    MainTextBox.Height = NextTop
    MainTextBox.Width = frmMain.Width
    MainTextBox.Visible = True
    
End Sub


Public Sub AddToXmlCombo(ByRef cbo As Control, NewItem As String)

    Dim ItemExists, i
    
    ItemExists = False
    
    For i = 0 To cbo.ListCount
        If cbo.List(i) = NewItem Then
            ItemExists = True
        End If
    Next i

    If Not ItemExists = True Then
        cbo.AddItem NewItem
    End If
    
End Sub
