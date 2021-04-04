Attribute VB_Name = "mdlXML"
Option Explicit
Public Sub XmlSetComment(Dom As MSXML2.DOMDocument, ByVal strComment As String)
'创建注释
    Dom.appendChild Dom.createComment(strComment)
End Sub
Public Function XmlSetRoot(Dom As MSXML2.DOMDocument, ByVal Name As String) As IXMLDOMElement
'创根节点
Dim oRoot  As IXMLDOMElement        '根节点
    Set oRoot = Dom.createElement(Name)
    Set Dom.documentElement = oRoot    '设置为根节点
    Set XmlSetRoot = oRoot
End Function
Public Function XmlCreateNode(ByVal TabNumber As Integer, ByVal Parent As IXMLDOMNode, _
    ByVal node_name As String, Optional ByVal Node_Value As String = "") As IXMLDOMNode
    Dim New_Node As IXMLDOMNode
    
    '字符缩进值设置（不影响数据），只影响阅读美观度
    Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf & String(TabNumber, vbKeyTab))     '创建文本节点
    '创建新节点
    Set New_Node = Parent.ownerDocument.createNode(NODE_ELEMENT, node_name, "")
    '设置文本值
    New_Node.Text = Node_Value
    '添加到父节点
    Parent.appendChild New_Node
    Set XmlCreateNode = New_Node
End Function

Public Function GetElemnetValue(ByVal Dom As MSXML2.DOMDocument, ByVal Name As String, Optional ByVal ChildName As String) As String
'功能：得到指定元素的值
    Dim ElementList As MSXML2.IXMLDOMNodeList, NodElement As MSXML2.IXMLDOMElement
    
    Set ElementList = Dom.getElementsByTagName(LCase(Name))
    If Not ElementList Is Nothing And ElementList.length >= 1 Then
        '找到指定子元素
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
'添加版本信息
    Dim pi As IXMLDOMProcessingInstruction
    Set pi = Dom.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call Dom.insertBefore(pi, Dom.childNodes(0))
End Sub
