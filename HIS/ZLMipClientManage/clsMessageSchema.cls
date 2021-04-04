VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessageSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量
Private mstrPath As String
'Private mdocXML As DOMDocument
'Private mdocSegmentXML As DOMDocument

Private mdocXML As Object
Private mdocSegmentXML As Object
'######################################################################################################################
'属性


'######################################################################################################################
'方法

Public Function GetMessageFormat(ByVal strFormatFile As String, ByRef rsFormat As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim xn As Object
    Dim xlst As Object
    Dim strFileName As String
    Dim lngLoop As Long
    Dim lngLength As Long
    Dim strSchema As String
    Dim strLayerKey(0 To 100) As String
    Dim objFso As New FileSystemObject
    Dim objFile As TextStream
    Dim lngStartPos As Long
    Dim lngEndPos As Long
    Dim strTemp As String
    Dim strTitle As String
    Dim strType As String
    Dim strMinOccurs As String
    Dim strMaxOccurs As String
    Dim intLayer As Integer
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    Set rsFormat = New ADODB.Recordset
    With rsFormat
        .Fields.Append "id", adVarChar, 32, adFldKeyColumn
        .Fields.Append "parent_id", adVarChar, 32
        .Fields.Append "节点类型", adVarChar, 50
        .Fields.Append "节点标题", adVarChar, 30
        .Fields.Append "数据类型", adVarChar, 30
        .Fields.Append "重复频率", adVarChar, 30
        .Fields.Append "节点说明", adVarChar, 200
        .Fields.Append "数据重复", adVarChar, 200
        .Fields.Append "数据赋值", adVarChar, 200
        .Open
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '判断并打开文件
    Set objFile = objFso.OpenTextFile(strFormatFile, ForReading)
    strSchema = objFile.ReadAll
    
    lngStartPos = InStr(strSchema, "<xs:schema")
    If lngStartPos > 0 Then
        lngEndPos = InStr(strSchema, "</xs:schema>")
        strSchema = Mid(strSchema, lngStartPos + 1, lngEndPos - lngStartPos - 1)
        lngStartPos = InStr(strSchema, ">")
        strSchema = Mid(strSchema, lngStartPos + 1)
    End If
        
    intLayer = 0
    
    Do While Trim(strSchema) <> ""
        
        lngStartPos = InStr(strSchema, "<")
        lngEndPos = InStr(strSchema, ">")
        
        strTemp = Trim(Mid(strSchema, lngStartPos + 1, lngEndPos - lngStartPos - 1))
        strSchema = Trim(Mid(strSchema, lngEndPos + 1))
        
        If InStr(strTemp, "xs:element ") = 1 Then
            
            strType = ""
            strTitle = ""
            strMaxOccurs = "1"
            strMinOccurs = "1"
            
            lngStartPos = InStr(strTemp, " name=""")
            lngEndPos = InStr(lngStartPos + 7, strTemp, """")
            If lngEndPos > lngStartPos Then strTitle = Mid(strTemp, lngStartPos + 7, lngEndPos - lngStartPos - 7)
                                    
            lngStartPos = InStr(strTemp, " type=""")
            lngEndPos = InStr(lngStartPos + 7, strTemp, """")
            If lngEndPos > lngStartPos And lngStartPos > 0 Then strType = Mid(strTemp, lngStartPos + 7, lngEndPos - lngStartPos - 7)
            
            lngStartPos = InStr(strTemp, " minOccurs=""")
            lngEndPos = InStr(lngStartPos + 12, strTemp, """")
            If lngEndPos > lngStartPos And lngStartPos > 0 Then strMinOccurs = Mid(strTemp, lngStartPos + 12, lngEndPos - lngStartPos - 12)
            
            lngStartPos = InStr(strTemp, " maxOccurs=""")
            lngEndPos = InStr(lngStartPos + 12, strTemp, """")
            If lngEndPos > lngStartPos And lngStartPos > 0 Then strMaxOccurs = Mid(strTemp, lngStartPos + 12, lngEndPos - lngStartPos - 12)
            If strMaxOccurs = "unbounded" Then strMaxOccurs = "u"
            
            '填写记录,name,type,minOccurs,maxOccurs,Layer
            
            rsFormat.AddNew
            rsFormat("id").Value = zlCommFun.GetGUID
            rsFormat("parent_id").Value = strLayerKey(intLayer)
            rsFormat("节点类型").Value = "Data"
            rsFormat("节点标题").Value = strTitle
            rsFormat("数据类型").Value = strType
            rsFormat("重复频率").Value = strMinOccurs & " ～ " & strMaxOccurs
            rsFormat("数据重复").Value = ""
            rsFormat("数据赋值").Value = ""
            
        ElseIf InStr(strTemp, "xs:complexType") = 1 Then
                                    
            rsFormat("节点类型").Value = "Complex"
            intLayer = intLayer + 1
            strLayerKey(intLayer) = rsFormat("id").Value
            
        ElseIf InStr(strTemp, "xs:sequence") = 1 Then
           
           '不处理
           
        ElseIf InStr(strTemp, "/xs:element") = 1 Then
            
            '不处理
            
        ElseIf InStr(strTemp, "/xs:complexType") = 1 Then
            
            intLayer = intLayer - 1
            
        ElseIf InStr(strTemp, "/xs:sequence") = 1 Then
            
            '不处理
            
        End If
        
ContinuePoint:
        
        strTemp = strSchema
        strTemp = Replace(strTemp, Chr(10), "")
        strTemp = Replace(strTemp, Chr(13), "")
        If strTemp = "" Then strSchema = ""
        
    Loop
        
    Dim strRootKey As String
    Dim strDataKey As String
    
    If rsFormat.RecordCount > 0 Then
        rsFormat.MoveFirst
        
        '删除根节下的head节点（head是系统自动产生的节点）
        strRootKey = rsFormat("id").Value
        rsFormat.Filter = "parent_id='" & strRootKey & "' And 节点标题='head'"
        If rsFormat.RecordCount > 0 Then
            rsFormat.MoveFirst
            
            strDataKey = rsFormat("id").Value
            
            rsFormat.Filter = "parent_id='" & strDataKey & "'"
            If rsFormat.RecordCount > 0 Then
                rsFormat.MoveFirst
                Do While Not rsFormat.EOF
                    rsFormat.Delete adAffectCurrent
                    rsFormat.MoveNext
                Loop
            End If
            
            rsFormat.Filter = ""
            rsFormat.Filter = "id='" & strDataKey & "'"
            If rsFormat.RecordCount > 0 Then rsFormat.Delete adAffectCurrent
        End If
        
        '更改根节点的直接下级节点的上级节点为空
        rsFormat.Filter = ""
        rsFormat.Filter = "parent_id='" & strRootKey & "'"
        If rsFormat.RecordCount > 0 Then
            rsFormat.MoveFirst
            Do While Not rsFormat.EOF
                rsFormat.Update "parent_id", ""
                rsFormat.MoveNext
            Loop
        End If
        
        
        rsFormat.Filter = ""
        rsFormat.Filter = "id='" & strRootKey & "'"
        If rsFormat.RecordCount > 0 Then rsFormat.Delete adAffectCurrent
        
        rsFormat.Filter = ""
    End If
    
    If rsFormat.RecordCount > 0 Then rsFormat.MoveFirst
    
    
    '------------------------------------------------------------------------------------------------------------------
    GetMessageFormat = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
'    Resume
End Function

Private Function GetMessageNode(ByVal xlst As Object, ByVal intLayer As Integer, ByVal strParentDataKey As String, ByRef rsFormat As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim xn As Object
    Dim xlst1 As Object
    Dim xlst2 As Object
    Dim xnf As Object
'    Dim xe As IXMLDOMElement
    Dim strTemp As String
    Dim i As Long
    Dim objPropertyCollect As Collection
    Dim strItem As String
    Dim strDataKey As String
    Dim strSegmentName As String
    Dim strNodeName As String
    Dim strDataType As String
    Dim strNodeType As String
    Dim strDataMin As String
    Dim strDataMax As String
    Dim strNodeInfo As String
                    
    On Error GoTo errHand
    
    intLayer = intLayer + 1
    
    For Each xnf In xlst
        
        '获取当前节点的属性
        Set objPropertyCollect = Nothing
        Set objPropertyCollect = New Collection
        
        If xnf.Attributes.length > 0 Then
            For i = 0 To xnf.Attributes.length - 1
                objPropertyCollect.Add xnf.Attributes(i).Text, "K_" & UCase(xnf.Attributes(i).nodeName)
            Next
        End If
        
        strNodeType = xnf.nodeName
        strNodeName = ""
        strDataType = ""
        strNodeInfo = ""
        strDataMin = ""
        strDataMax = ""
        
        On Error Resume Next
        
        'strNodeName
        '--------------------------------------------------------------------------------------------------------------
        Select Case UCase(strNodeType)
        Case UCase("Group")
            strNodeName = objPropertyCollect(UCase("K_name"))
        Case UCase("Segment")
            strNodeName = objPropertyCollect(UCase("K_nodeName"))
            strSegmentName = objPropertyCollect(UCase("K_ref"))
            If strNodeName = "" Then strNodeName = strSegmentName
            
        Case UCase("Data")
            strNodeName = objPropertyCollect(UCase("K_nodeName"))
            If strNodeName = "" Then strNodeName = objPropertyCollect(UCase("K_ref"))
        Case UCase("Composite")
            strNodeName = objPropertyCollect(UCase("K_nodeName"))
        End Select
        
        'strDataType
        '--------------------------------------------------------------------------------------------------------------
        Select Case UCase(strNodeType)
        Case UCase("Composite")
            strDataType = objPropertyCollect(UCase("K_ref"))
        End Select
        
        'strDataMin
        '--------------------------------------------------------------------------------------------------------------
        strItem = ""
        strItem = objPropertyCollect(UCase("K_minOccurs"))
        If strItem <> "" Then strDataMin = strItem
                                    
        'strDataMax
        '--------------------------------------------------------------------------------------------------------------
        strItem = ""
        strItem = objPropertyCollect(UCase("K_maxOccurs"))
        If strItem <> "" Then strDataMax = strItem
        
        
        'strNodeInfo
        '--------------------------------------------------------------------------------------------------------------
        Select Case UCase(strNodeType)
        Case UCase("Composite")
            strNodeInfo = objPropertyCollect(UCase("K_info"))
        End Select
        
        On Error GoTo errHand
        
        '填写记录
        '--------------------------------------------------------------------------------------------------------------
        rsFormat.AddNew
        strDataKey = zlCommFun.GetGUID
        rsFormat("id").Value = strDataKey
        rsFormat("parent_id").Value = strParentDataKey
        rsFormat("节点类型").Value = Space(2 * (intLayer - 1)) & strNodeType
        rsFormat("节点标题").Value = strNodeName
        rsFormat("数据类型").Value = strDataType
        
        Select Case UCase(strDataMax)
        Case UCase("unbounded")
            rsFormat("重复频率").Value = Val(strDataMin) & " ～ u"
        Case Else
            rsFormat("重复频率").Value = Val(strDataMin) & " ～ " & Val(strDataMax)
        End Select
                    
        rsFormat("节点说明").Value = strNodeInfo
        
        '段
        '--------------------------------------------------------------------------------------------------------------
        If UCase(strNodeType) = UCase("Segment") Then
            Set xn = mdocSegmentXML.selectSingleNode(".//Elements//Segment[@name='" & strSegmentName & "']")
            If Not (xn Is Nothing) Then
                Set xlst1 = xn.childNodes
                If GetMessageNode(xlst1, intLayer, strDataKey, rsFormat) = False Then Exit Function
            End If
        End If
        
        '搜索下级
        '--------------------------------------------------------------------------------------------------------------
        If Not (xnf.childNodes Is Nothing) Then
            
            Set xlst2 = xnf.childNodes
            If GetMessageNode(xlst2, intLayer, strDataKey, rsFormat) = False Then Exit Function
        End If
        
    Next
            
    GetMessageNode = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
    Resume
End Function


