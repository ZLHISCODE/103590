Attribute VB_Name = "mdlXML"
Option Explicit
Public Sub XmlSetComment(Dom As MSXML2.DOMDocument, ByVal strComment As String)
'����ע��
    Dom.appendChild Dom.createComment(strComment)
End Sub
Public Function XmlSetRoot(Dom As MSXML2.DOMDocument, ByVal Name As String) As IXMLDOMElement
'�����ڵ�
Dim oRoot  As IXMLDOMElement        '���ڵ�
    Set oRoot = Dom.createElement(Name)
    Set Dom.documentElement = oRoot    '����Ϊ���ڵ�
    Set XmlSetRoot = oRoot
End Function
Public Function XmlCreateNode(ByVal TabNumber As Integer, ByVal Parent As IXMLDOMNode, _
    ByVal node_name As String, Optional ByVal Node_Value As String = "") As IXMLDOMNode
    Dim New_Node As IXMLDOMNode
    
    '�ַ�����ֵ���ã���Ӱ�����ݣ���ֻӰ���Ķ����۶�
    Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf & String(TabNumber, vbKeyTab))     '�����ı��ڵ�
    '�����½ڵ�
    Set New_Node = Parent.ownerDocument.createNode(NODE_ELEMENT, node_name, "")
    '�����ı�ֵ
    New_Node.Text = Node_Value
    '��ӵ����ڵ�
    Parent.appendChild New_Node
    Set XmlCreateNode = New_Node
End Function

Public Function GetElemnetValue(ByVal Dom As MSXML2.DOMDocument, ByVal Name As String, Optional ByVal ChildName As String) As String
'���ܣ��õ�ָ��Ԫ�ص�ֵ
    Dim ElementList As MSXML2.IXMLDOMNodeList, NodElement As MSXML2.IXMLDOMElement
    
    Set ElementList = Dom.getElementsByTagName(LCase(Name))
    If Not ElementList Is Nothing And ElementList.length >= 1 Then
        '�ҵ�ָ����Ԫ��
        If ChildName = "" Then
            GetElemnetValue = ElementList.Item(0).Text
        Else
            Set NodElement = ElementList.Item(0).selectSingleNode(LCase(ChildName))
            If Not NodElement Is Nothing Then
                GetElemnetValue = NodElement.Text
            End If
        End If
    End If
End Function
Public Sub XmlSetVersion(Dom As MSXML2.DOMDocument)
'��Ӱ汾��Ϣ
    Dim pi As IXMLDOMProcessingInstruction
    Set pi = Dom.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call Dom.insertBefore(pi, Dom.childNodes(0))
End Sub
