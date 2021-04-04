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
Attribute VB_Exposed = True
Option Explicit
'######################################################################################################################

Public Enum enumXmlType
    xsString = 1
    xsNumber = 2
    xsDate = 3
    xsTime = 4
    xsDateTime = 5
End Enum

Private mstrXmlText As String                                                       'XML缓存记录
Private mintNodeLayer As Integer
Private mobjFso As FileSystemObject
Private mobjTextStream As TextStream
Private mobjXML As Object
Private mstrXmlVersion As String

'######################################################################################################################
'公共方法属性

Public Sub ClearXmlText()
    '******************************************************************************************************************
    '功能：清空缓存的XML字符串
    '******************************************************************************************************************
    mintNodeLayer = 1
    mstrXmlText = ""
End Sub

Public Sub AppendNode(ByVal strName As String, Optional ByVal blnCloseNode As Boolean)
    '******************************************************************************************************************
    '功能：添加无内容的XML节点，即上级节点
    '参数：strName - 节点名；blnCloseNode - 是否为结束节点
    '返回：
    '******************************************************************************************************************
    
    If blnCloseNode = False Then
        mstrXmlText = mstrXmlText & Space(mintNodeLayer) & "<" & strName & ">" & vbCrLf
        mintNodeLayer = mintNodeLayer + 1
    Else
        mintNodeLayer = mintNodeLayer - 1
        mstrXmlText = mstrXmlText & Space(mintNodeLayer) & "</" & strName & ">" & vbCrLf
    End If
    
End Sub

Public Sub AppendData(ByVal strName As String, ByVal strData As String, Optional ByVal enumType As enumXmlType = xsString)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    strData = ChangeSpecialChar(strData)
    
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
    '功能：获取缓存的XML字符串结果
    '返回：XML字符串
    '******************************************************************************************************************
    XmlText = mstrXmlText
End Property

Public Function OpenXMLDocument(ByVal strXml As String) As Boolean
    '******************************************************************************************************************
    '功能：打开XML文档
    '参数：strXML-XML字符串
    '返回：成功返回True，否则返回False
    '******************************************************************************************************************
    On Error GoTo errHand
    
    mstrXmlVersion = GetXMLVersion
    
    Set mobjXML = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
    
    OpenXMLDocument = mobjXML.loadXML(strXml)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    Set mobjXML = Nothing
    OpenXMLDocument = False
End Function

Public Function GetSingleNodeValue(ByVal strNode As String, ByRef strValue As String, Optional ByVal enumType As enumXmlType = xsString) As Boolean
    '******************************************************************************************************************
    '功能：从打开的XML文档中获取指定节的值（单一节点）
    '参数：strNode-节点名称，如Meta/Release；strValue-返回节点值
    '返回：成功返回True，否则返回False
    '******************************************************************************************************************
    Dim strData As String
    Dim objNode As Object
    
    On Error GoTo errHand
        
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
errHand:
    GetSingleNodeValue = False
End Function

Public Function GetMultiNodeRecord(ByVal strNode As String, ByRef rsValue As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：从打开的XML文档中获取指定节点及节点以下的所有内容，以记录集形式返回
    '参数：strNode-节点名称，如Meta/Release；rsValue-返回节点记录集
    '返回：成功返回True，否则返回False
    '******************************************************************************************************************
    Dim objNodeList As Object
    Dim lngRow As Long
    
    On Error GoTo errHand
        
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
    
    Set objNodeList = mobjXML.selectNodes(strNode)
    
    lngRow = 0
    If GetXMLNode(rsValue, objNodeList, lngRow, 0) = False Then Exit Function
                       
    If rsValue.RecordCount > 0 Then rsValue.MoveFirst
    
    GetMultiNodeRecord = True
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    GetMultiNodeRecord = False
End Function

Public Function GetAllNodeRecord(ByRef rsValue As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：从打开的XML文档中获取指定节点及节点以下的所有内容，以记录集形式返回
    '参数：strNode-节点名称，如Meta/Release；rsValue-返回节点记录集
    '返回：成功返回True，否则返回False
    '******************************************************************************************************************
    Dim objNodeList As Object
    Dim lngRow As Long
    
    On Error GoTo errHand
        
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
    
    Set objNodeList = mobjXML.selectNodes(".//message")
    
    lngRow = 0
    If GetXMLNode(rsValue, objNodeList, lngRow, 0) = False Then Exit Function
                       
    If rsValue.RecordCount > 0 Then rsValue.MoveFirst
    
    GetAllNodeRecord = True
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    GetAllNodeRecord = False
End Function

Public Function CloseXMLDocument() As Boolean
    '******************************************************************************************************************
    '功能：关闭XML文档
    '返回：成功返回True，否则返回False
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Set mobjXML = Nothing
    CloseXMLDocument = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    CloseXMLDocument = False
End Function

'######################################################################################################################
Private Function ChangeSpecialChar(ByVal strXmlText As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strXml As String
    
'    &lt;    <   小于
'    &gt;    >   大于
'    &amp;   &   和号
'    &apos;  '   单引号
'    &quot;  "   引号
    
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
    '功能：
    '参数：
    '返回：
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
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objChildNodeList As Object
    Dim objNode As Object
    Dim blnExistChildNode As Boolean
    
    On Error GoTo errHand
    
    For Each objNode In objNodeList
        
        lngRow = lngRow + 1
                                
        rsXML.AddNew
        rsXML("id").value = lngRow
        If lngParentRow <> 0 Then
            rsXML("parent_id").value = lngParentRow
        End If
        rsXML("node_name").value = objNode.nodeName
        
        blnExistChildNode = True
        If objNode.childNodes.Length = 1 Then
            If objNode.childNodes.Item(0).hasChildNodes = False Then
                blnExistChildNode = False
            End If
        End If
        
        If blnExistChildNode = False Then
            rsXML("node_value").value = objNode.Text
        Else
            rsXML("node_value").value = ""
            Set objChildNodeList = objNode.childNodes
            Call GetXMLNode(rsXML, objChildNodeList, lngRow, lngRow)
        End If
    Next
    
    GetXMLNode = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    GetXMLNode = False
End Function


Private Function GetXMLVersion() As String
    
    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXML As Object
    
    On Error GoTo errHand
        
    varXMLVersion = Split(".6.0,.4.0", ",")
    
    On Error Resume Next
    If gobjComLib.OS.IsDesinMode = True Or gobjComLib.zlRegInfo("授权性质") <> "1" Then
        For intLoop = 0 To UBound(varXMLVersion)
            Err = 0
            Set objXML = CreateObject("MSXML2.DOMDocument" & varXMLVersion(intLoop))
            If Err = 0 Then
                strXMLVer = varXMLVersion(intLoop)
                Exit For
            End If
        Next
        On Error GoTo errHand
        
        If strXMLVer = "" Then
            MsgBox "创建MSXML2.DOMDocument对象失败"
            Exit Function
        End If
    Else
        strXMLVer = ""
    End If
    GetXMLVersion = strXMLVer
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function
