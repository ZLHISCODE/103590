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

Private mstrXmlText As String                                                       'XML缓存记录
Private mintNodeLayer As Integer
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

Public Sub AppendData(ByVal strName As String, ByVal strData As String, _
    Optional ByVal enumType As enumXmlType = xsString, _
    Optional ByVal blnIsCDATA As Boolean = False)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
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
'功能：打开XML结构的文件
'参数：
'  strFile：文件路径和全名
'返回：True成功；False失败
    
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
    '功能：从打开的XML文档中获取指定节的值（单一节点）
    '参数：strNode-节点名称，如Meta/Release；strValue-返回节点值
    '返回：成功返回True，否则返回False
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
'功能：从打开的XML文档中获取指定节、属性的值
'参数：
'  strNode：节点名称，如Meta/Release
'  strProperty：属性名称
'  strValue：返回节点值
'返回：成功返回True，否则返回False
    
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
    '功能：从打开的XML文档中获取指定节点及节点以下的所有内容，以记录集形式返回
    '参数：strNode-节点名称，如Meta/Release；rsValue-返回节点记录集
    '返回：成功返回True，否则返回False
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
    '功能：从打开的XML文档中获取指定节点及节点以下的所有内容，以记录集形式返回
    '参数：strNode-节点名称，如Meta/Release；rsValue-返回节点记录集
    '返回：成功返回True，否则返回False
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
    '功能：关闭XML文档
    '返回：成功返回True，否则返回False
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
'            MsgBox "创建MSXML2.DOMDocument对象失败"
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
'功能：修改指定结点的值
'参数：
'  strNode：修改的结点
'  strValue：修改的结点值

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
'    MsgBox Err.Description, vbInformation, "clsXML类"
End Sub

Public Sub SaveXMLFile(ByVal strFile As String)
'功能：将mobjXML对象的XML保存到指定文件中
'参数：
'  strFile：文件路径和全名

    If mobjXML Is Nothing Then Exit Sub
    
    On Error GoTo hErr
    
    mobjXML.Save strFile
    
    Exit Sub
    
hErr:
'    MsgBox Err.Description, vbInformation, "clsXML类"
End Sub

Private Function IsDesinMode() As Boolean
'功能： 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function

