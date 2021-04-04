Attribute VB_Name = "mdlXML"
Option Explicit
Public gobjXml As MSXML2.DOMDocument
Private mintDebug As Integer
Public Function zlXML_Init(Optional ByVal strNode As String = "DATA", _
    Optional blnNotMsg As Boolean = False, Optional strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化XML,增加声明和根节点
    '入参:strNode-接点
    '出参:strErrMsg-返回错误信息
    '返回: 初始成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-27 10:58:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim nodData As MSXML2.IXMLDOMElement
    mintDebug = -1
    On Error Resume Next
    Set gobjXml = New MSXML2.DOMDocument
    If Err <> 0 Then
        DebugTools "初始化XML对象失败(InitXML)《 " & strNode & "》"
        strErrMsg = "初始化XML对象失败(InitXML)《 " & strNode & "》"
        Err.Raise vbObjectError + 1, , strErrMsg
    End If
    '根节点
    Err = 0: On Error GoTo errHand:
    Set nodData = gobjXml.createElement(strNode)
    Set gobjXml.documentElement = nodData
    zlXML_Init = True: Exit Function
errHand:
    If blnNotMsg Then strErrMsg = Err.Description: Exit Function
    If ErrCenter = 1 Then Resume
End Function
Public Function zlXML_InsertNodes(nodParent As MSXML2.IXMLDOMElement, _
    ByVal cllData As Collection) As Boolean
    '------------------------------------------------------------------s---------------------------------------------------------------------------
    '功能:插入接点数
    '入参:nodParent-父接点
    '        cllData-数据Array(接点名,接点值)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-27 11:03:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    On Error GoTo errHandle
    For i = 1 To cllData.Count
        Call zlXML_InsertNode(nodParent, cllData(i)(0), cllData(i)(1))
    Next
    zlXML_InsertNodes = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlXML_InsertNode(nodParent As MSXML2.IXMLDOMElement, _
    ByVal Name As String, ByVal value As String, Optional ByRef OutNod As MSXML2.IXMLDOMElement) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:在指定XML元素下增加子元素
    '入参:nodParent-父接点
    '        Name-接点名
    '        Value-接点值
    '出参:OutNod-返回接点对象
    '返回:增加成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-27 11:26:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHandle
    Set OutNod = gobjXml.createElement(Name)
    OutNod.Text = value
    nodParent.appendChild OutNod
    zlXML_InsertNode = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlXML_GetXMLString(Optional blnHead As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取XML字符串
    '入参:blnHead-是否包含头数据
    '返回:完整的XML串
    '编制:刘兴洪
    '日期:2011-05-27 11:31:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnHead Then
        zlXML_GetXMLString = gobjXml.xml
    Else
        zlXML_GetXMLString = "<?xml version=""1.0"" encoding=""gb2312""?>" & vbCrLf & gobjXml.xml
    End If
End Function

Public Function zlXML_GetRows(ByVal strNodeName As String, ByRef lngOutRows As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取XML行数
    '入参:strNodeName-接点名
    '出参:lngOutRows-返回XML行数
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-27 10:51:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    lngOutRows = 0
    Err = 0: On Error GoTo errHand:
    If mintDebug = 5 Then
       Call zlTempLoadXML
    End If
    lngOutRows = gobjXml.getElementsByTagName(strNodeName).length
    DebugTools "获取XML的记录行数(GetOutXMLRows)《 " & strNodeName & "》:" & lngOutRows
    zlXML_GetRows = True
    Exit Function
errHand:
    DebugTools "获取XML的记录行数(GetOutXMLRows)《 " & strNodeName & "》" & vbCrLf & "错误序号:" & vbCrLf & "   " & Err.Description
    If ErrCenter = 1 Then Resume
End Function
Private Sub zlTempLoadXML()
    'J调试用:临时加载XML文件
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Set objText = objFile.OpenTextFile(App.Path & "\xml.txt", ForReading)
    Call zlXML_LoadXMLToDOMDocument(objText.ReadAll)
End Sub

Public Function zlXML_GetNodeValue( _
    ByVal strNodeName As String, Optional ByVal lngRow As Long = 0, _
    Optional ByRef strOutPut As String, Optional ByRef strErrMsg As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:得到指定元素的值
    '入参:strNodeName-接点名
    '       lngRow-指定行数
    '       strErrMsg-错误信息
    '出参:strOutPut-返回值
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-27 10:52:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim xmlElement As MSXML2.IXMLDOMElement
    Err = 0: On Error GoTo errHand:
    strOutPut = "": strErrMsg = ""
    If lngRow >= 0 Then
        Set xmlElement = gobjXml.getElementsByTagName(strNodeName).Item(lngRow)
    Else
        Set xmlElement = gobjXml.documentElement.selectSingleNode(strNodeName)
    End If
    If Not xmlElement Is Nothing Then
        '找到指定子元素
        strOutPut = Replace(xmlElement.Text, Chr(10), "")
    Else
        strErrMsg = strNodeName & "不存在，请检查!"
        DebugTools strErrMsg
        If Not gobjXml Is Nothing Then
            DebugTools gobjXml.xml
        Else
            DebugTools "gobjXml.xml=nothing"
        End If
    End If
    zlXML_GetNodeValue = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlXML_LoadXMLToDOMDocument(ByVal strXMLInPut As String, _
    Optional blnAddHead As Boolean = True, Optional blnNotMsg As Boolean = False, _
    Optional ByRef strErrMsg As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载XML给相关的对象
    '入参:strXMLInPut-相关的XML串
    '出参:strErrMsg-返回错误信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-03-13 11:07:00
    '-----------------------------------------------------------------------------------------------------------
    Dim strXml As String
    Err = 0: On Error GoTo errHand:
    If Not blnAddHead Then
        strXml = strXMLInPut
    Else
        strXml = Replace("'<?xml version=''1.0'' encoding=''gb2312''?>'", "'", Chr(34)) & vbCrLf & strXMLInPut
    End If
    If Not gobjXml.loadXML(strXml) Then
        strErrMsg = "XML解析错误(LoadXML):" & vbCrLf & strXml
        DebugTools strErrMsg
        Err.Raise vbObjectError + 1, , strErrMsg
    End If
    zlXML_LoadXMLToDOMDocument = True
    Exit Function
errHand:
    strErrMsg = Err.Description
    If blnNotMsg Then Exit Function
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub DebugTools(ByVal strInfo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:跟踪调试信息
    '入参:strInfo-调试信息
    '编制:刘兴洪
    '日期:2011-05-27 11:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call LogWrite("一卡通接口调试日志", glngModul, "读卡接口返回", strInfo)
End Sub
 
Public Function zlXML_GetChildNodeValue(strParentName As String, strChildName As String, Optional ByVal lngParentRow As Long = 0, Optional ByVal lngChildRow As Long = 0, Optional ByRef strOutPut As String, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取对应父节点下子节点的值
    '入参:strParentName-父节点名称 strChildName-子节点名称 lngParentRow - 指定父行数 lngChildRow - 指定子行数
    '出参:strOutPut-返回值 strErrMsg -错误信息
    '编制:王吉
    '日期:2012-12-12 11:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim xmlParentElement As MSXML2.IXMLDOMElement
    Dim xmlChildElement As MSXML2.IXMLDOMElement
    Err = 0: On Error GoTo errHand:
    If lngParentRow >= 0 Then
        Set xmlParentElement = gobjXml.documentElement.getElementsByTagName(strParentName).Item(lngParentRow)
    Else
        Set xmlParentElement = gobjXml.documentElement.selectSingleNode(strParentName)
    End If
    If lngChildRow >= 0 Then
        Set xmlChildElement = xmlParentElement.getElementsByTagName(strChildName).Item(lngChildRow)
    Else
        Set xmlChildElement = xmlParentElement.selectSingleNode(strChildName)
    End If
    If Not xmlChildElement Is Nothing Then
        '找到指定子元素
        strOutPut = Replace(xmlChildElement.Text, Chr(10), "")
    Else
        strErrMsg = strChildName & "不存在，请检查!"
        DebugTools strErrMsg
        If Not gobjXml Is Nothing Then
            DebugTools gobjXml.xml
        Else
            DebugTools "gobjXml.xml=nothing"
        End If
    End If
    zlXML_GetChildNodeValue = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlXML_GetChildRows(ByVal strParentName, ByVal strChildName As String, ByRef lngOutRows As Long, Optional ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取对应父节点下子节点的数量
    '入参:strParentName-父节点名 ,strChildName-子节点名
    '出参:lngOutRows-返回XML行数
    '返回:成功,返回true,否则返回False
    '编制:王吉
    '日期:2011-12-12 10:51:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    lngOutRows = 0
    Err = 0: On Error GoTo errHand:
    If mintDebug = 5 Then
       Call zlTempLoadXML
    End If
    lngOutRows = gobjXml.getElementsByTagName(strParentName).Item(lngRow).selectNodes(strChildName).length
    DebugTools "获取XML的记录行数(GetOutXMLRows)《 " & strChildName & "》:" & lngOutRows
    zlXML_GetChildRows = True
    Exit Function
errHand:
    DebugTools "获取XML的记录行数(GetOutXMLRows)《 " & strChildName & "》" & vbCrLf & "错误序号:" & vbCrLf & "   " & Err.Description
    If ErrCenter = 1 Then Resume
End Function

Public Function zlXML_GetChildNodes(ByVal strParentName As String) As IXMLDOMNodeList
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取对应父节点下子节点集合
    '入参:strParentName-父节点名
    '返回:lngOutRows-返回XML行数
    '返回:成功,返回true,否则返回False
    '编制:王吉
    '日期:2011-12-13 14:51:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim xmlParentElement As MSXML2.IXMLDOMElement

    Set xmlParentElement = gobjXml.documentElement.selectSingleNode(strParentName)
    If Not xmlParentElement Is Nothing Then
       Set zlXML_GetChildNodes = xmlParentElement.childNodes
    Else
       Set zlXML_GetChildNodes = Nothing
    End If
End Function

Public Function zlXML_ExistNode(ByVal strXMLIn As String, ByVal strNodeName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:XML串中是否存在某节点
    '入参:strXMLInPut-相关的XML串
    '   strNodeName-节点名
    '返回:节点存在,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    zlXML_ExistNode = InStr(strXMLIn, "<" & strNodeName & ">") > 0
    Exit Function
ErrHandler:
    zlXML_ExistNode = False
End Function

