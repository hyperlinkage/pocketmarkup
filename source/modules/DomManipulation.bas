Attribute VB_Name = "DomManipulation"
'   Module DomManipulation.bas
'
'   Description:
'       A couple of functions to carry out some generic operations
'       on an XML object.
'
'   Public Functions:
'       AddAttribute(Parser As DOMDocument, Node As IXMLDOMNode, Name As String, Value As String)
'       AppendChildNode(Parser As DOMDocument, Node As IXMLDOMNode, Name As String)
'
'   Written by:
'       tim@yaffle.org
'
'   Version:
'       1.0.1
'       20th February 2002


Option Explicit


Public Sub AddAttribute(Parser As DOMDocument, Node As IXMLDOMNode, Name As String, Value As String)

'   Description:
'       Adds an attribute with Name and Value to an IXMLDOMNode

    ' create the attribute
    Dim Attrib As IXMLDOMAttribute
    Set Attrib = Parser.createAttribute(Name)
    Attrib.Value = Value

    ' add the attribute to the node
    Node.Attributes.setNamedItem Attrib

End Sub


Public Sub AppendChildNode(Parser As DOMDocument, Node As IXMLDOMNode, Name As String)

'   Description:
'       Appends a named child node to an IXMLDOMNode.  Does not give
'       the new child node a value.

    Dim ChildNode As IXMLDOMNode
    Set ChildNode = Parser.createElement(Name)

    Node.appendChild ChildNode

End Sub


Public Sub CheckWellformedDoc(txt As TextBox, NotifyOk As Boolean)

'   Description:
'       Checks that a TextBox contains well-formed XML.  If not, displays
'       an error message and moves cursor to the erroneous position.

    ' create instance of the msxml parser
    Dim Parser As DOMDocument
    Set Parser = CreateObject("Microsoft.XMLDOM")

    ' pass it the contents of the textbox
    Parser.LoadXml txt.Text

    'check for a parse error
    If Parser.parseError Then

        ' move cursor to error position
        txt.SelLength = 1
        txt.SelStart = Parser.parseError.filepos

        ' display the error message
        MsgBox Parser.parseError.reason, vbExclamation, "PocketMarkup"

    Else

        If NotifyOk Then
            ' display ok message
            MsgBox "This document is well-formed XML.", vbInformation, "PocketMarkup"

        End If

    End If

End Sub
