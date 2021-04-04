VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessageXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'����
Private mstrPath As String
'Private mdocXML As DOMDocument60
Private mdocXML As Object

'######################################################################################################################
'����


'######################################################################################################################
'����

Public Function LoadFile(ByVal strXMLFile As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Set mdocXML = InitXMLDoc
        
    mdocXML.Load strXMLFile
        
    '------------------------------------------------------------------------------------------------------------------
    LoadFile = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Public Function LoadString(ByVal strXMLString As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
        
    On Error GoTo errHand
    
    Set mdocXML = InitXMLDoc
    
    LoadString = mdocXML.loadXML(strXMLString)
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Public Function GetSingleNodeData(ByVal strNode As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsXML As ADODB.Recordset
'    Dim xn As IXMLDOMNode
'    Dim xlst As IXMLDOMNodeList
'    Dim xlst1 As IXMLDOMNodeList
'    Dim xnf As IXMLDOMNode
'    Dim xe As IXMLDOMElement

    Dim strTemp As String
    Dim i As Long
    Dim xn As Object
    Dim xlst As Object
    Dim xlst1 As Object
    Dim xnf As Object
    Dim xe As Object
    
    On Error GoTo errHand
        
    '------------------------------------------------------------------------------------------------------------------
    '
    Set rsXML = New ADODB.Recordset
    With rsXML
        .Fields.Append "id", adVarChar, 32, adFldKeyColumn
        .Fields.Append "parent_id", adVarChar, 32
        .Fields.Append "����", adVarChar, 50
        .Fields.Append "����", adVarChar, 100
        .Fields.Append "����", adTinyInt
        .Open
    End With
    
    
    Set xn = mdocXML.selectSingleNode(strNode)
    
    If Not (xn Is Nothing) Then
        
        Set xlst = xn.childNodes
        
        For Each xnf In xlst

            strTemp = strTemp & " " & xnf.nodeName & "=" & xnf.Text
            
'            If xnf.Attributes.length > 0 Then
'                strTemp = strTemp & vbCrLf & "���ԣ�"
'                For i = 0 To xnf.Attributes.length - 1
'                    strTemp = strTemp & vbCrLf & xnf.Attributes(i).nodeName & "=" & xnf.Attributes(i).nodeValue
'                Next
'            End If
'
'            If Not (xnf.childNodes Is Nothing) Then
'                strTemp = strTemp & vbCrLf & "�� " & xnf.childNodes.length & " ������"
'            End If
            
            

        Next
        
        MsgBox strTemp
    End If
'
            
'    '
'    Set xlst = mdocXML.selectNodes(strNode)
'    For Each xnf In xlst
'
'        If IsNull(xnf.nodeValue) Then
'            strTemp = xnf.nodeName & "=null"
'        Else
'            strTemp = xnf.nodeName & "=" & xnf.nodeValue
'        End If
'
'        If xnf.Attributes.length > 0 Then
'            strTemp = strTemp & vbCrLf & "���ԣ�"
'            For i = 0 To xnf.Attributes.length - 1
'
'                strTemp = strTemp & vbCrLf & xnf.Attributes(i).nodeName & "=" & xnf.Attributes(i).nodeValue
'
'            Next
'
'        End If
'
'        MsgBox strTemp
'
'    Next
        
    '------------------------------------------------------------------------------------------------------------------
    GetSingleNodeData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
    Resume
End Function

Public Function GetNodeData(ByVal strNode As String) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsXML As ADODB.Recordset
'    Dim objNodeList As IXMLDOMNodeList
    Dim objNodeList As Object
    
    On Error GoTo errHand
        
    '------------------------------------------------------------------------------------------------------------------
    Set rsXML = New ADODB.Recordset
    With rsXML
        .Fields.Append "id", adVarChar, 32, adFldKeyColumn
        .Fields.Append "parent_id", adVarChar, 32
        .Fields.Append "nodename", adVarChar, 50
        .Fields.Append "nodetext", adVarChar, 4000
        .Open
    End With
        
    Set objNodeList = mdocXML.selectNodes(strNode)
    Call GetXMLNode(rsXML, objNodeList)
                       
    If rsXML.RecordCount > 0 Then rsXML.MoveFirst
    Set GetNodeData = rsXML
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Private Sub GetXMLNode(ByRef rsXML As ADODB.Recordset, ByVal objNodeList As Object, Optional ByVal strUpKey As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
'    Dim objChildNodeList As IXMLDOMNodeList
    Dim objChildNodeList As Object
    Dim strKey As String
'    Dim objNode As IXMLDOMNode
    Dim objNode As Object
    Dim blnExistChildNode As Boolean
    
    On Error GoTo errHand
    
    For Each objNode In objNodeList
        strKey = zlCommFun.GetGUID
        
        rsXML.AddNew
        rsXML("id").Value = strKey
        rsXML("parent_id").Value = strUpKey
        rsXML("nodename").Value = objNode.nodeName
                
        blnExistChildNode = True
        If objNode.childNodes.length = 1 Then
            If objNode.childNodes.Item(0).hasChildNodes = False Then
                blnExistChildNode = False
            End If
        End If
        
        If blnExistChildNode = False Then
            rsXML("nodetext").Value = objNode.Text
        Else
            rsXML("nodetext").Value = ""
            Set objChildNodeList = objNode.childNodes
            Call GetXMLNode(rsXML, objChildNodeList, strKey)
        End If
    Next
    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Sub

Private Function ChangeSpecialChar(ByVal strXmlText As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strXml As String
    
'    &lt;    <   С��
'    &gt;    >   ����
'    &amp;   &   �ͺ�
'    &apos;  '   ������
'    &quot;  "   ����
    
    strXml = strXmlText
    If InStr(strXml, "<") > 0 Then strXml = Replace(strXml, "<", "&lt;")
    If InStr(strXml, ">") > 0 Then strXml = Replace(strXml, ">", "&gt;")
    If InStr(strXml, "&") > 0 Then strXml = Replace(strXml, "&", "&amp;")
    If InStr(strXml, "'") > 0 Then strXml = Replace(strXml, "'", "&apos;")
    If InStr(strXml, """") > 0 Then strXml = Replace(strXml, """", "&quot;")
    
    ChangeSpecialChar = strXml
End Function

Private Function RestoreSpecialChar(ByVal strXmlText As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strXml As String
    
    strXml = strXmlText
    If InStr(strXml, "&lt;") > 0 Then strXml = Replace(strXml, "&lt;", "<")
    If InStr(strXml, "&gt;") > 0 Then strXml = Replace(strXml, "&gt;", ">")
    If InStr(strXml, "&amp;") > 0 Then strXml = Replace(strXml, "&amp;", "&")
    If InStr(strXml, "&apos;") > 0 Then strXml = Replace(strXml, "&apos;", "'")
    If InStr(strXml, "&quot;") > 0 Then strXml = Replace(strXml, "&quot;", """")
    
    RestoreSpecialChar = strXml
End Function


Private Sub Class_Terminate()
    Set mdocXML = Nothing
End Sub