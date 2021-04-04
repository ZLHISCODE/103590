VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public gobjXml As MSXML2.DOMDocument
Private mintDebug As Integer
Public Function zlXML_Init(Optional ByVal strNode As String = "DATA", _
    Optional ByVal blnNotErrMsg As Boolean, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化XML,增加声明和根节点
    '入参:strNode-接点
    '出参:strErrMsg-返回的错误信息
    '返回:初始成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-27 10:58:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim nodData As MSXML2.IXMLDOMElement
    '根节点
    Set nodData = gobjXml.createElement(strNode)
    Set gobjXml.documentElement = nodData
    zlXML_Init = True
End Function

Public Function zlXML_InsertNodes(nodParent As MSXML2.IXMLDOMElement, _
    ByVal cllData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlXML_InsertNode(nodParent As MSXML2.IXMLDOMElement, _
    ByVal Name As String, ByVal Value As String, _
    Optional ByRef OutNod As MSXML2.IXMLDOMElement) As Boolean
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
    OutNod.Text = Value
    nodParent.appendChild OutNod
    zlXML_InsertNode = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
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
         ' "<?xml version=""1.0"" encoding=""utf-8"" ?>"
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
    lngOutRows = gobjXml.getElementsByTagName(strNodeName).Length
    DebugTools "获取XML的记录行数(GetOutXMLRows)《 " & strNodeName & "》:" & lngOutRows
    zlXML_GetRows = True
    Exit Function
errHand:
    DebugTools "获取XML的记录行数(GetOutXMLRows)《 " & strNodeName & "》" & vbCrLf & "错误序号:" & vbCrLf & "   " & Err.Description
    If gobjComLib.ErrCenter() = 1 Then Resume
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
    Optional ByRef strOutput As String, Optional ByRef strErrMsg As String = "") As Boolean
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
    strOutput = "": strErrMsg = ""
    If lngRow >= 0 Then
        Set xmlElement = gobjXml.getElementsByTagName(strNodeName).Item(lngRow)
    Else
        Set xmlElement = gobjXml.documentElement.selectSingleNode(strNodeName)
        
    End If
    If Not xmlElement Is Nothing Then
        '找到指定子元素
        strOutput = Replace(xmlElement.Text, Chr(10), "")
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
    DebugTools strErrMsg
End Function
Public Function zlXML_LoadXMLToDOMDocument(ByVal strXMLInPut As String, _
     Optional blnAddHead As Boolean = True, Optional strErrMsg As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载XML给相关的对象
    '入参:strXMLInPut-相关的XML串
    '出参:
    '       strErrMsg-返回错误信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-03-13 11:07:00
    '-----------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Err = 0: On Error GoTo errHand:
    If Not blnAddHead Then
        strXML = strXMLInPut
    Else
        strXML = Replace("'<?xml version=''1.0'' encoding=''gb2312''?>'", "'", Chr(34)) & vbCrLf & strXMLInPut
    End If
    DebugTools "进入加载XML给对象:"
    If Not gobjXml.loadXML(strXML) Then
        strErrMsg = "XML解析错误"
        DebugTools "XML解析错误" & strXML
        Exit Function
    End If
    DebugTools "加载XML给对象成功:"
    zlXML_LoadXMLToDOMDocument = True
    Exit Function
errHand:
    strErrMsg = zlTranErrInfor(Err.Description)
    DebugTools "加载XML给对象失败:" & strErrMsg
End Function

Private Sub DebugTools(ByVal strInfo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:跟踪调试信息
    '入参:strInfo-调试信息
    '编制:刘兴洪
    '日期:2011-05-27 11:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call gobjComLib.LogWrite("一卡通接口调试日志", glngModul, "读卡接口返回", strInfo)
End Sub
Public Property Get zlGetXMLHeadString() As String
    zlGetXMLHeadString = "<?xml version=""1.0"" encoding=""utf-8"" ?>"
End Property
Private Function TranErrInfor(ByVal strErrMsg As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:格式化错误信息
    '编制:刘兴洪
    '日期:2012-06-21 10:49:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If InStr(UCase(strErrMsg), "[ZLSOFT]") > 0 Then
        TranErrInfor = Split(strErrMsg, "[ZLSOFT]")(1)
    Else
        TranErrInfor = strErrMsg
    End If
    Exit Function
errHandle:
      TranErrInfor = strErrMsg
End Function
Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    Set gobjXml = New MSXML2.DOMDocument
    mintDebug = -1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set gobjXml = Nothing
    mintDebug = -1
End Sub
