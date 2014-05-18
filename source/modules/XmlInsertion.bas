Attribute VB_Name = "XmlInsertion"
Option Explicit

Private Sub InsertTag(e As Control, ByRef t As Control, typ As String)
    ' wraps selected text in new tag
   
    ' declare variables
    Dim SelStartPos, selEndPos, TagName, STag, ETag, EmptyElemTag
    
    ' prepare tag name
    TagName = PrepTagName(t.Text)
    
    If isValidTag(TagName) Then
    
        AddToXmlCombo t, TagName
        
        ' get start and end position of selection
        SelStartPos = e.SelStart
        selEndPos = SelStartPos + e.SelLength + 1
        
        e.SelLength = 0
        
        If typ = "Empty" Then
            EmptyElemTag = "<" & TagName & " />"
            ' insert empty tag
            e.SelText = EmptyElemTag
        Else
            STag = "<" & TagName & ">"
            ETag = "</" & TagName & ">"
            
            ' insert start tag
            e.SelText = STag
            
            ' move selection and insert end tag
            e.SelLength = 0
            e.SelStart = selEndPos + Len(STag) - 1
            e.SelText = ETag
        End If
        
        ' reselect text
        e.SetFocus
        e.SelStart = SelStartPos + Len(STag) + Len(EmptyElemTag)
        e.SelLength = selEndPos - SelStartPos - 1
    
    End If
    
End Sub


Private Sub InsertAttribute(ByRef Box As Control, NameBox As Control, ValBox As Control)

'   Description:
'       Inserts an attribute and value into the tag nearest the cursor.
'
'   Known bugs:
'       - Should not create duplicate attributes, update value instead
'       - Inserts into nearest tag, even if its an ending tag

    
    Dim CmdError As Boolean
    
    If Box.Text = "" Then
        CmdError = "Cannot insert attribute - no tags in document"
    End If

    If CmdError = "" Then

        Dim SelPos, SelLng, AttPos
        
        AddToXmlCombo NameBox, NameBox.Text
        AddToXmlCombo ValBox, ValBox.Text
        
        SelPos = Box.SelStart
        SelLng = Box.SelLength
        
        Box.SelLength = 0
        
        AttPos = InStr(SelPos, Box.Text, ">")
        
        If AttPos > 0 Then
        
            If InStr(AttPos - 1, Box.Text, "/") = AttPos - 1 Then
                
                Box.SelStart = AttPos - 2
                Box.SelText = " " & PrepTagName(NameBox.Text) & "=""" & ValBox.Text & """ "
    
            Else
            
                Box.SelStart = AttPos - 1
                Box.SelText = " " & PrepTagName(NameBox.Text) & "=""" & ValBox.Text & """"
            
            End If
                
        End If
        
    Else
        MsgBox CmdError, vbOKOnly, "Error"
    End If
    
End Sub

Private Function isValidTag(TagName As String) As Boolean

'   Description:
'       Validates a tag name
'       Currently returns true if TagName is not empty
'
'   Known bugs:
'       - Needs to be rewritten to check for well-formed qnames
'       - Needs to be duplicated for attribute names
'       - Seems to reset cursor to top of main text box
    
    If Not TagName = "" Then
        isValidTag = True
    Else
        isValidTag = False
    End If

End Function
