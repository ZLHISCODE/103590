VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'######################################################################################################################

Public Enum enumXmlType
    xsString = 1
    xsNumber = 2
    xsDate = 3
    xsTime = 4
    xsDateTime = 5
End Enum

Private mstrXmlText As String                                                       'XML�����¼
Private mintNodeLayer As Integer
Private mobjXML As Object
Private mstrXmlVersion As String

'######################################################################################################################
'������������

Public Sub ClearXmlText()
    '******************************************************************************************************************
    '���ܣ���ջ����XML�ַ���
    '******************************************************************************************************************
    mintNodeLayer = 1
    mstrXmlText = ""
End Sub

Public Sub AppendNode(ByVal strName As String, Optional ByVal blnCloseNode As Boolean)
    '******************************************************************************************************************
    '���ܣ����������ݵ�XML�ڵ㣬���ϼ��ڵ�
    '������strName - �ڵ�����blnCloseNode - �Ƿ�Ϊ�����ڵ�
    '���أ�
    '******************************************************************************************************************
    
    If blnCloseNode = False Then
        mstrXmlText = mstrXmlText & Space(mintNodeLayer) & "<" & strName & ">" & vbCrLf
        mintNodeLayer = mintNodeLayer + 1
    Else
        mintNodeLayer = mintNodeLayer - 1
        mstrXmlText = mstrXmlText & Space(mintNodeLayer) & "</" & strName & ">" & vbCrLf
    End If
    
End Sub

Public Sub AppendData(ByVal strName As String, ByVal strData As String, _
    Optional ByVal enumType As enumXmlType = xsString, _
    Optional ByVal blnIsCDATA As Boolean = False)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    If blnIsCDATA = False Then
        strData = ChangeSpecialChar(strData)
    End If
    
    Select Case enumType
    Case enumXmlType.xsString
    
    Case enumXmlType.xsNumber
    
    Case enumXmlType.xsDate
        If strData <> "" Then strData = Format(strData, "YYYY-MM-DD")
    Case enumXmlType.xsTime
        If strData <> "" Then strData = Format(strData, "hh:mm:ss")
    Case enumXmlType.xsDateTime
        If strData <> "" Then strData = Format(strData, "YYYY-MM-DDThh:mm:ss")
    End Select
    
    mstrXmlText = mstrXmlText & Space(mintNodeLayer) & "<" & strName & ">" & strData & "</" & strName & ">" & vbCrLf
End Sub

Public Property Get XmlText() As String
    '******************************************************************************************************************
    '���ܣ���ȡ�����XML�ַ������
    '���أ�XML�ַ���
    '******************************************************************************************************************
    XmlText = mstrXmlText
End Property

Public Function OpenXMLDocument(ByVal strXml As String) As Boolean
    '******************************************************************************************************************
    '���ܣ���XML�ĵ�
    '������strXML-XML�ַ���
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    On Error GoTo Errhand
    
    mstrXmlVersion = GetXMLVersion
    
    Set mobjXML = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
    
    OpenXMLDocument = mobjXML.loadXML(strXml)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    Set mobjXML = Nothing
    OpenXMLDocument = False
End Function

Public Function OpenXMLFile(ByVal strFile As String) As Boolean
'���ܣ���XML�ṹ���ļ�
'������
'  strFile���ļ�·����ȫ��
'���أ�True�ɹ���Falseʧ��
    
    On Error GoTo hErr
    
    mstrXmlVersion = GetXMLVersion
    
    Set mobjXML = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
    
    OpenXMLFile = mobjXML.Load(strFile)
    
    Exit Function

hErr:
    Set mobjXML = Nothing
End Function

Public Function GetSingleNodeValue(ByVal strNode As String, ByRef strValue As String, Optional ByVal enumType As enumXmlType = xsString) As Boolean
    '******************************************************************************************************************
    '���ܣ��Ӵ򿪵�XML�ĵ��л�ȡָ���ڵ�ֵ����һ�ڵ㣩
    '������strNode-�ڵ����ƣ���Meta/Release��strValue-���ؽڵ�ֵ
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    Dim strData As String
    Dim objNode As Object
    
    On Error GoTo Errhand
        
    If mobjXML Is Nothing Then Exit Function
    
    strNode = ".//" & Replace(strNode, "/", "//")
    
    Set objNode = mobjXML.selectSingleNode(strNode)
    strData = objNode.Text
    
    strValue = RestoreSpecialChar(strData)
    
    Select Case enumType
    Case enumXmlType.xsDateTime
        If strValue <> "" Then strValue = Replace(strValue, "T", " ")
    End Select
    
    GetSingleNodeValue = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    GetSingleNodeValue = False
End Function

Public Function GetNodePropertyValue(ByVal strNode As String, ByVal strProperty As String, ByRef strValue As String, _
    Optional ByVal enumType As enumXmlType = xsString) As Boolean
'���ܣ��Ӵ򿪵�XML�ĵ��л�ȡָ���ڡ����Ե�ֵ
'������
'  strNode���ڵ����ƣ���Meta/Release
'  strProperty����������
'  strValue�����ؽڵ�ֵ
'���أ��ɹ�����True�����򷵻�False
    
    Dim strData As String
    Dim objNode As Object
    
    On Error GoTo hErr
    
    If mobjXML Is Nothing Then Exit Function
    
    strNode = ".//" & Replace(strNode, "/", "//")
    
    Set objNode = mobjXML.selectSingleNode(strNode)
    strData = objNode.Attributes.getNamedItem(strProperty).Text
    
    strValue = RestoreSpecialChar(strData)
    
    Select Case enumType
    Case enumXmlType.xsDateTime
        If strValue <> "" Then strValue = Replace(strValue, "T", " ")
    End Select
    
    GetNodePropertyValue = True
    
    Exit Function
    
hErr:
End Function

Public Function GetMultiNodeRecord(ByVal strNode As String, ByRef rsValue As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ��Ӵ򿪵�XML�ĵ��л�ȡָ���ڵ㼰�ڵ����µ��������ݣ��Լ�¼����ʽ����
    '������strNode-�ڵ����ƣ���Meta/Release��rsValue-���ؽڵ��¼��
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    Dim objNodeList As Object
    Dim lngRow As Long
    
    On Error GoTo Errhand
        
    If mobjXML Is Nothing Then Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
    Set rsValue = New ADODB.Recordset
    With rsValue
        .Fields.Append "id", adBigInt, , adFldKeyColumn
        .Fields.Append "parent_id", adBigInt
        .Fields.Append "node_name", adVarChar, 50
        .Fields.Append "node_value", adVarChar, 4000
        .Open
    End With
    
    strNode = ".//" & Replace(strNode, "/", "//")
    
    Set objNodeList = mobjXML.SelectNodes(strNode)
    
    lngRow = 0
    If GetXMLNode(rsValue, objNodeList, lngRow, 0) = False Then Exit Function
                       
    If rsValue.RecordCount > 0 Then rsValue.MoveFirst
    
    GetMultiNodeRecord = True
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    GetMultiNodeRecord = False
End Function

Public Function GetAllNodeRecord(ByRef rsValue As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ��Ӵ򿪵�XML�ĵ��л�ȡָ���ڵ㼰�ڵ����µ��������ݣ��Լ�¼����ʽ����
    '������strNode-�ڵ����ƣ���Meta/Release��rsValue-���ؽڵ��¼��
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    Dim objNodeList As Object
    Dim lngRow As Long
    
    On Error GoTo Errhand
        
    If mobjXML Is Nothing Then Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
    Set rsValue = New ADODB.Recordset
    With rsValue
        .Fields.Append "id", adBigInt, , adFldKeyColumn
        .Fields.Append "parent_id", adBigInt
        .Fields.Append "node_name", adVarChar, 50
        .Fields.Append "node_value", adVarChar, 4000
        .Open
    End With
    
    Set objNodeList = mobjXML.SelectNodes(".//message")
    
    lngRow = 0
    If GetXMLNode(rsValue, objNodeList, lngRow, 0) = False Then Exit Function
                       
    If rsValue.RecordCount > 0 Then rsValue.MoveFirst
    
    GetAllNodeRecord = True
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    GetAllNodeRecord = False
End Function

Public Function CloseXMLDocument() As Boolean
    '******************************************************************************************************************
    '���ܣ��ر�XML�ĵ�
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    
    On Error GoTo Errhand
    
    Set mobjXML = Nothing
    CloseXMLDocument = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    CloseXMLDocument = False
End Function

'######################################################################################################################
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

Private Function GetXMLNode(ByRef rsXML As ADODB.Recordset, ByVal objNodeList As Object, ByRef lngRow As Long, ByVal lngParentRow As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objChildNodeList As Object
    Dim objNode As Object
    Dim blnExistChildNode As Boolean
    
    On Error GoTo Errhand
    
    For Each objNode In objNodeList
        
        lngRow = lngRow + 1
                                
        rsXML.AddNew
        rsXML("id").Value = lngRow
        If lngParentRow <> 0 Then
            rsXML("parent_id").Value = lngParentRow
        End If
        rsXML("node_name").Value = objNode.nodeName
        
        blnExistChildNode = True
        If objNode.childNodes.Length = 1 Then
            If objNode.childNodes.Item(0).hasChildNodes = False Then
                blnExistChildNode = False
            End If
        End If
        
        If blnExistChildNode = False Then
            rsXML("node_value").Value = objNode.Text
        Else
            rsXML("node_value").Value = ""
            Set objChildNodeList = objNode.childNodes
            Call GetXMLNode(rsXML, objChildNodeList, lngRow, lngRow)
        End If
    Next
    
    GetXMLNode = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    GetXMLNode = False
End Function


Private Function GetXMLVersion() As String
    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXML As Object
    
    On Error GoTo Errhand
        
    varXMLVersion = Split(".6.0,.4.0", ",")
    
    On Error Resume Next
    If IsDesinMode = True Then
        For intLoop = 0 To UBound(varXMLVersion)
            Err = 0
            Set objXML = CreateObject("MSXML2.DOMDocument" & varXMLVersion(intLoop))
            If Err = 0 Then
                strXMLVer = varXMLVersion(intLoop)
                Exit For
            End If
        Next
        On Error GoTo Errhand
        
        If strXMLVer = "" Then
'            MsgBox "����MSXML2.DOMDocument����ʧ��"
            Exit Function
        End If
    Else
        strXMLVer = ""
    End If
    GetXMLVersion = strXMLVer
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
'    MsgBox Err.Description
End Function

Public Sub SetSingleNodeValue(ByVal strNode As String, ByVal strValue As String, Optional ByVal strParentNode As String)
'���ܣ��޸�ָ������ֵ
'������
'  strNode���޸ĵĽ��
'  strValue���޸ĵĽ��ֵ

    Dim objNode As Object
    Dim strPath As String
    
    If mobjXML Is Nothing Then Exit Sub
    
    On Error GoTo hErr
    
    If strParentNode = "" Then
        strPath = "//" & strNode
    Else
        strPath = "//" & strParentNode & "/" & strNode
    End If
    
    Set objNode = mobjXML.selectSingleNode(strPath)
    If objNode Is Nothing Then Exit Sub
    
    objNode.nodeTypedValue = strValue
    
    Exit Sub
    
hErr:
'    MsgBox Err.Description, vbInformation, "clsXML��"
End Sub

Public Sub SaveXMLFile(ByVal strFile As String)
'���ܣ���mobjXML�����XML���浽ָ���ļ���
'������
'  strFile���ļ�·����ȫ��

    If mobjXML Is Nothing Then Exit Sub
    
    On Error GoTo hErr
    
    mobjXML.Save strFile
    
    Exit Sub
    
hErr:
'    MsgBox Err.Description, vbInformation, "clsXML��"
End Sub

Private Function IsDesinMode() As Boolean
'���ܣ� ȷ����ǰģʽΪ���ģʽ
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
