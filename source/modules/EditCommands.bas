Attribute VB_Name = "basEdit"
' Edit menu

Option Explicit

' API constants
Private Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const WM_UNDO = &H304

' API declarations
Public Declare Function SendMessage Lib "Coredll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Sub DisableEditMenu(EditUndo As MenuBarControl, EditCut As MenuBarControl, _
              EditCopy As MenuBarControl, EditPaste As MenuBarControl, _
              EditDelete As MenuBarControl, EditSelectAll As MenuBarControl)

' Disables all Edit menu items.
' IN:  EditUndo, Undo menu item (control)
'      EditCut, Cut menu item (control)
'      EditCopy, Copy menu item (control)
'      EditPaste, Paste menu item (control)
'      EditDelete, Delete menu item (control)
'      EditSelectAll, Select All menu item (control)
  
  EditUndo.Enabled = False
  EditCut.Enabled = False
  EditCopy.Enabled = False
  EditPaste.Enabled = False
  EditDelete.Enabled = False
  EditSelectAll.Enabled = False
  
End Sub
Public Sub InitEditMenu(EditUndo As MenuBarControl, EditCut As MenuBarControl, _
              EditCopy As MenuBarControl, EditPaste As MenuBarControl, _
              EditDelete As MenuBarControl, EditSelectAll As MenuBarControl)

' Handles activaion of Edit menu items.
' IN:  EditUndo, Undo menu item (control)
'      EditCut, Cut menu item (control)
'      EditCopy, Copy menu item (control)
'      EditPaste, Paste menu item (control)
'      EditDelete, Delete menu item (control)
'      EditSelectAll, Select All menu item (control)
  
  EditUndo.Enabled = pCanUndo()
  EditCut.Enabled = pCanCut()
  EditCopy.Enabled = pCanCopy()
  EditPaste.Enabled = pCanPaste()
  EditDelete.Enabled = pCanDelete()
  EditSelectAll.Enabled = pCanSelectAll()
  
End Sub
Private Function pCanUndo() As Boolean

' Check if Undo is possible.
' OUT: if Undo is possible
    
  pCanUndo = SendMessage(Screen.ActiveControl.hWnd, EM_CANUNDO, 0, 0) <> 0

End Function
Private Function pCanCut() As Boolean

' Check if Cut is possible.
' OUT: if Cut is possible
  
  pCanCut = (Screen.ActiveControl.SelLength > 0)
  
End Function
Private Function pCanCopy() As Boolean

' Check if Copy is possible.
' OUT: if Copy is possible
    
  pCanCopy = (Screen.ActiveControl.SelLength > 0)
  
End Function
Private Function pCanPaste() As Boolean

' Check if Paste is possible.
' OUT: if Paste is possible

  pCanPaste = Clipboard.GetFormat(vbCFText)
  
End Function
Private Function pCanDelete() As Boolean

' Check if Delete is possible.
' OUT: if Delete is possible
    
  pCanDelete = (Screen.ActiveControl.SelLength > 0)
  
End Function
Private Function pCanSelectAll() As Boolean

' Check if Select All is possible.
' OUT: if Select All is possible
    
  pCanSelectAll = (Screen.ActiveControl.SelLength <> Len(Screen.ActiveControl.Text))

End Function
Public Function EditUndo() As Long

' Perform Undo operation.
' OUT: return value from SendMessage API

  EditUndo = SendMessage(Screen.ActiveControl.hWnd, WM_UNDO, 0, 0)
  
End Function
Public Function EditCut() As Long

' Perform Cut operation.
' OUT: return value from SendMessage API

  EditCut = SendMessage(Screen.ActiveControl.hWnd, WM_CUT, 0, 0)
  
End Function
Public Function EditCopy() As Long

' Perform Copy operation.
' OUT: return value from SendMessage API
    
  EditCopy = SendMessage(Screen.ActiveControl.hWnd, WM_COPY, 0, 0)

End Function
Public Function EditPaste() As Long

' Perform Paste operation.
' OUT: return value from SendMessage API
  EditPaste = SendMessage(Screen.ActiveControl.hWnd, WM_PASTE, 0, 0)

End Function
Public Function EditDelete() As Long

' Perform Delete operation.
' OUT: return value from SendMessage API

  EditDelete = SendMessage(Screen.ActiveControl.hWnd, WM_CLEAR, 0, 0)
  
End Function
Public Sub EditSelectAll()

' Perform Select All operation.

  Screen.ActiveControl.SelStart = 0
  Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
  
End Sub

