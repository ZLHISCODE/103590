VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

'常量
'----------------------------------------------------------------------------------------------------------------------
Private Const MAX_IP = 5                                                    'To make a buffer... i dont think you have more than 5 ip on your pc..

'类型
'----------------------------------------------------------------------------------------------------------------------
Private Type IPINFO
    dwAddr As Long                                                          ' IP address
    dwIndex As Long                                                         ' interface index
    dwMask As Long                                                          ' subnet mask
    dwBCastAddr As Long                                                     ' broadcast address
    dwReasmSize  As Long                                                    ' assembly size
    unused1 As Integer                                                      ' not currently used
    unused2 As Integer                                                      '; not currently used
End Type

Private Type MIB_IPADDRTABLE
    dEntrys As Long                                                         'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO                                               'array of IP address entries
End Type

Private Type MipClient
    OnLine As Boolean                                                       '是否在线
    System As Long                                                          '系统编号
    Module As Long                                                          '模块编号
    Privs As String                                                         '模块权限
    count As Long                                                           '
    MipModule As clsMipModule                                               '对应的平台客户端分支
End Type

'API
'----------------------------------------------------------------------------------------------------------------------
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'变量
'----------------------------------------------------------------------------------------------------------------------
Private mobjMipClient(1 To 100) As MipClient                                '记录ZLHIS消息应用场合（如导航台、模块等）
Private mstrDbUser As String                                                '登录ZLHIS的DB用户
Private mstrTitle As String                                                 '
Private mstrError As String                                                 '
Private mstrSQL As String                                                   '
Private mstrStationName As String                                           '
Private mstrStationIP As String                                             '

Private mlngPersonKey As Long                                               '
Private mlngPersonDeptKey As Long
Private mobjCollectProperty As Collection
Private mobjCollectRole As Collection
Private mstrUser As String
Private mstrUserName As String
Private mstrMspServerIP As String
Private mlngMspServerPort As Long
Private mstrMspUser As String
Private mstrMspUserPassword As String
Private mobjXML As Object
Private mobjXML2 As Object
Private mstrXmlVersion As String
Private mobjFso As New FileSystemObject
Private mrsModule As ADODB.Recordset
'Private WithEvents mclsMipClientShell As zlMipClientShell.clsMipClientShell

'######################################################################################################################
Public Function ConnectMip(ByVal lngParentForm As Long) As Boolean
'    '******************************************************************************************************************
'    '功能：连接消息服务平台
'    '参数：无
'    '说明：连接成功返回True，否则返回False
'    '******************************************************************************************************************
'    Dim rsTmp As ADODB.Recordset
'    Dim strPara As String
'    Dim varPara As Variant
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    '1.检查ZLHIS是否启用了消息
'    '------------------------------------------------------------------------------------------------------------------
'    strSQL = "Select 1 From zltools.zlOptions Where 参数号=16 And 参数值='1'"
'    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, mstrTitle)
'    If rsTmp.BOF = True Then Exit Function
'
'    '2.启动消息客户端标志程序
'    '------------------------------------------------------------------------------------------------------------------
'    On Error Resume Next
'    Set mclsMipClientShell = New zlMipClientShell.clsMipClientShell
'    Call mclsMipClientShell.Initialize(lngParentForm)
'    On Error GoTo errHand
'
'    '3.获取本机工作站名称及IP地址
'    '------------------------------------------------------------------------------------------------------------------
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "读取连接消息服务平台的相关参数")
'
'    mstrUser = gstrDBUser
'    mstrStationName = AnalyseComputer
'    mstrStationIP = GetIpAddress(AnalyseIP)
'
'    '4.获取ZLHIS产品登录用户及对应的部门
'    '------------------------------------------------------------------------------------------------------------------
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "获取ZLHIS产品登录用户及对应的部门")
'
'    Set rsTmp = gobjComLib.zlDatabase.GetUserInfo
'    If rsTmp.BOF = False Then
'        mstrUserName = gobjComLib.zlCommFun.NVL(rsTmp("姓名").value)
'        mlngPersonKey = gobjComLib.zlCommFun.NVL(rsTmp("ID").value, 0)
'        mlngPersonDeptKey = gobjComLib.zlCommFun.NVL(rsTmp("部门id").value, 0)
'    End If
'
'    '5.获取消息服务平台连接参数
'    '------------------------------------------------------------------------------------------------------------------
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "获取消息服务平台连接参数")
'
'    mstrSQL = "Select 参数值 From zltools.zlOptions Where 参数号=18 And 参数值 Is Not Null"
'    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle)
'    If rsTmp.BOF = False Then
'        strPara = rsTmp("参数值").value
'        varPara = Split(strPara, ";")
'        mstrMspUser = varPara(0)
'        mstrMspUserPassword = varPara(1)
'        mstrMspServerIP = varPara(2)
'        mlngMspServerPort = Val(varPara(3))
'    End If
'
'    If mstrMspServerIP = "" Or mlngMspServerPort = 0 Or mstrMspUser = "" Then
'        If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "未配置连接消息服务平台的用户或IP地址或端口号")
'        Exit Function
'    End If
'
'    mstrSQL = "Select mipuser,mipuserpassword From zltools.zlmip_clientuser Where station=[1] And mipuser Is Not Null"
'    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, UCase(mstrStationName))
'    If rsTmp.BOF = False Then
'        mstrMspUser = rsTmp("mipuser").value
'        mstrMspUserPassword = rsTmp("mipuserpassword").value
'    End If
'
'    '6.连接消息服务平台服务器
'    '------------------------------------------------------------------------------------------------------------------
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "启动消息服务平台客户端收发服务")
'
'    Call mclsMipClientShell.ConnectMip(mstrMspUser, mstrMspUserPassword, mstrMspServerIP, mlngMspServerPort)
'
'    If mclsMipClientShell.IsConnect = False Then Exit Function
'
'    '7.填写当前操作人员在线记录（此块内容应算属于ZLHIS自身管理工具范围）
'    '------------------------------------------------------------------------------------------------------------------
'    mstrSQL = "zltools.zlMip_OnLineUser_Insert('" & mstrStationName & "','" & gstrDBUser & "')"
'    Call gobjComLib.zlDatabase.ExecuteProcedure(mstrSQL, mstrTitle)
'
'    mstrSQL = "Select 系统, 序号, 标题 From Zltools.Zlprograms"
'    Set mrsModule = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle)
'
'    ConnectMip = mclsMipClientShell.IsConnect
'
'    mstrXmlVersion = GetXMLVersion
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mstrError = Err.Description
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "启动消息服务平台客户端收发服务程序出错：" & vbCrLf & mstrError)
End Function

Public Function AddModule(ByRef objMipModule As clsMipModule) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim intLoop As Integer
'
'    For intLoop = 1 To 100
'        If mobjMipClient(intLoop).OnLine = False Then
'
'            If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "模块登入（系统号=" & objMipModule.System & "，模块号=" & objMipModule.Module & "）")
'
'            mobjMipClient(intLoop).OnLine = True
'            mobjMipClient(intLoop).System = objMipModule.System
'            mobjMipClient(intLoop).Module = objMipModule.Module
'            mobjMipClient(intLoop).Privs = ";" & objMipModule.Privs & ";"
'            Set mobjMipClient(intLoop).MipModule = objMipModule
'            Set mobjMipClient(intLoop).MipModule.MipClient = Me
'
'            Exit For
'
'        End If
'    Next
'
'    AddModule = True
    
End Function

Public Function DeleteModule(ByRef objMipModule As clsMipModule) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim intLoop As Integer
'
'    For intLoop = 1 To 100
'        If mobjMipClient(intLoop).OnLine = True Then
'            If mobjMipClient(intLoop).System = objMipModule.System And mobjMipClient(intLoop).Module = objMipModule.Module And mobjMipClient(intLoop).MipModule.Instance = objMipModule.Instance Then
'
'                If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "模块登出（系统号=" & objMipModule.System & "，模块号=" & objMipModule.Module & "）")
'
'                Set mobjMipClient(intLoop).MipModule = Nothing
'                mobjMipClient(intLoop).System = 0
'                mobjMipClient(intLoop).Module = 0
'                mobjMipClient(intLoop).Privs = ""
'                mobjMipClient(intLoop).OnLine = False
'                Exit For
'            End If
'        End If
'    Next
'
'    DeleteModule = True
    
End Function

Public Property Get IsConnect() As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'
'    If Not (mclsMipClientShell Is Nothing) Then
'        IsConnect = mclsMipClientShell.IsConnect
'    End If
        
End Property

'Friend Function SendMessage(ByVal strMsgIdentity As String, _
'                                ByVal strMsgContent As String, _
'                                Optional ByRef strRequestResult As String, _
'                                Optional ByVal lngSystem As Long, _
'                                Optional ByVal lngModule As Long, _
'                                Optional ByVal strInstance As String, _
'                                Optional ByVal strProgram As String) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim strXml As String
'    Dim strItemFlow As String
'    Dim blnRequest As Boolean
'    Dim strMessageKey As String
'    Dim strSendLogKey As String
'    Dim rsTmp As ADODB.Recordset
'
'    On Error GoTo errHand
'
'    If mclsMipClientShell Is Nothing Then Exit Function
'    If mclsMipClientShell.IsConnect = False Then Exit Function
'
'    Call mclsMipClientShell.WriteRunLog("信息", "ZLHIS消息中心收到来自（系统号=" & lngSystem & "，模块号=" & lngModule & "）的消息" & vbCrLf & "消息编号：" & strMsgIdentity & "消息内容：" & strMsgContent)
'
'    If Not (mclsMipClientShell Is Nothing) Then
'
'        '先检查待发的消息是否合法
'        '--------------------------------------------------------------------------------------------------------------
'        If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "检查消息有效性")
'
'        mstrSQL = "Select item_title,item_flow From zltools.zlmip_item Where item_code=[1]"
'        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, strMsgIdentity)
'        If rsTmp.BOF = False Then
'            strMessageKey = GetGUID
'
'            strItemFlow = gobjComLib.zlCommFun.NVL(rsTmp("item_flow").value)
'
'            '添加消息头
'            '----------------------------------------------------------------------------------------------------------
'            If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "构造消息头内容")
'
'            strXml = "<?xml version=""1.0"" encoding=""gbk""?>" & vbCrLf
'            strXml = strXml & "<message>" & vbCrLf
'            strXml = strXml & " <head>" & vbCrLf
'            strXml = strXml & "  <" & LCase(strMsgIdentity) & ">" & rsTmp("item_title").value & "</" & LCase(strMsgIdentity) & ">" & vbCrLf
'            strXml = strXml & "  <msg_id>" & strMessageKey & "</msg_id>" & vbCrLf
'            strXml = strXml & "  <msg_item_code>" & strMsgIdentity & "</msg_item_code>" & vbCrLf
'            strXml = strXml & "  <send_station_name>" & mstrStationName & "</send_station_name>" & vbCrLf
'            strXml = strXml & "  <send_station_ip>" & mstrStationIP & "</send_station_ip>" & vbCrLf
'            strXml = strXml & "  <send_program>" & strProgram & "</send_program>" & vbCrLf
'            strXml = strXml & "  <send_instance>" & strInstance & "</send_instance>" & vbCrLf
'            strXml = strXml & "  <send_system_code>" & lngSystem & "</send_system_code>" & vbCrLf
'            strXml = strXml & "  <send_module_code>" & lngModule & "</send_module_code>" & vbCrLf
'            strXml = strXml & "  <send_mipuser>" & mstrMspUser & "</send_mipuser>" & vbCrLf
'
'            If mstrUserName = "" Then
'                strXml = strXml & "  <send_hisuser>" & mstrUser & "</send_hisuser>" & vbCrLf
'            Else
'                strXml = strXml & "  <send_hisuser>" & mstrUser & "/" & mstrUserName & "</send_hisuser>" & vbCrLf
'            End If
'            strXml = strXml & "  <send_time>" & Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "</send_time>" & vbCrLf
'            strXml = strXml & " </head>" & vbCrLf
'            strXml = strXml & strMsgContent
'            strXml = strXml & "</message>"
'
'            '填写消息发送记录
'            '----------------------------------------------------------------------------------------------------------
'            strSendLogKey = GetGUID
'
'            If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "登记发送记录" & vbCrLf & "ID=" & strSendLogKey)
'
'            Call WriteSendLog(strSendLogKey, strMessageKey, strMsgIdentity, Replace(strXml, "'", "''"))
'
'            '发送消息
'            '----------------------------------------------------------------------------------------------------------
'            If Not (mclsMipClientShell Is Nothing) Then SendMessage = mclsMipClientShell.SendMessage(strXml, strItemFlow, blnRequest, strRequestResult)
'
'            '更新消息发送记录状态
'            '----------------------------------------------------------------------------------------------------------
'            If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "更新消息发送记录状态")
'            mrsModule.Filter = ""
'            mrsModule.Filter = "系统=" & lngSystem & " And 序号=" & lngModule
'            If mrsModule.RecordCount > 0 Then
'                mstrSQL = "zltools.zlMip_SendLog_UpdateState('" & strSendLogKey & "'," & IIf(SendMessage = True, 1, 2) & ",'" & mstrMspUser & "','" & mstrUser & "','" & mstrUserName & "','" & mstrStationName & "(" & mstrStationIP & ")" & "','" & mrsModule("标题").value & "',Null,'')"
'                Call gobjComLib.zlDatabase.ExecuteProcedure(mstrSQL, mstrTitle)
'            End If
'        Else
'            If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "非法的消息标识" & vbCrLf & "消息标识=" & strMsgIdentity)
'        End If
'    Else
'        If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "消息服务平台连接已丢失")
'    End If
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mstrError = Err.Description
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "发送消息过程中出错：" & vbCrLf & mstrError)
'End Function
'
'Friend Function ShowMessage(ByVal strMsgKey As String, ByVal strMsgContent As String, ByVal strInstance As String) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim intCount As Integer
'
'    If Not (mclsMipClientShell Is Nothing) Then
'        For intCount = 1 To 100
'            If mobjMipClient(intCount).OnLine = True Then
'                If mobjMipClient(intCount).MipModule.Instance = strInstance Then
'                    Call mclsMipClientShell.ShowMessage(strMsgKey, strMsgContent, strInstance)
'                    Exit For
'                End If
'            End If
'        Next
'    End If
'
'End Function

Public Function DisConnectMip() As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim intLoop As Integer
'
'    On Error GoTo errHand
'
''    If mblnConnected = False Then Exit Function
'
'    For intLoop = 1 To 100
'        If mobjMipClient(intLoop).OnLine = True Then
'            mobjMipClient(intLoop).OnLine = False
'            mobjMipClient(intLoop).System = 0
'            mobjMipClient(intLoop).Module = 0
'            mobjMipClient(intLoop).Privs = ""
'        End If
'    Next
'
'    If Not (mclsMipClientShell Is Nothing) Then
'        If mclsMipClientShell.IsConnect Then
'            '删除当前操作人员在线记录
'            mstrSQL = "zltools.zlMip_OnLineUser_Delete('" & mstrStationName & "','" & gstrDBUser & "')"
'            Call gobjComLib.zlDatabase.ExecuteProcedure(mstrSQL, mstrTitle)
'
'            mclsMipClientShell.DisConnectMip
'            Call mclsMipClientShell.WriteRunLog("信息", "停止消息服务平台客户端收发服务")
'        End If
'        Set mclsMipClientShell = Nothing
'    End If
'
'    DisConnectMip = True
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mstrError = Err.Description
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "停止消息服务平台客户端收发服务过程中出错：" & vbCrLf & mstrError)
End Function

Private Sub Class_Initialize()
    mstrTitle = "消息平台客户端"
End Sub

Private Sub Class_Terminate()
        
'    If Not (mobjCollectRole Is Nothing) Then Set mobjCollectRole = Nothing
'    If Not (mobjCollectProperty Is Nothing) Then Set mobjCollectProperty = Nothing
'    If Not (mobjXML Is Nothing) Then Set mobjXML = Nothing
'    If Not (mobjXML2 Is Nothing) Then Set mobjXML2 = Nothing
'
'    If Not (mclsMipClientShell Is Nothing) Then Set mclsMipClientShell = Nothing
    If Not (mobjFso Is Nothing) Then Set mobjFso = Nothing
'    If Not (mrsModule Is Nothing) Then Set mrsModule = Nothing
    
End Sub

'Private Sub mclsMipClientShell_ConnectStateChanged(ByVal IsConnected As Boolean)
'    Dim intLoop As Integer
'
'
'    For intLoop = 1 To 100
'        If mobjMipClient(intLoop).OnLine = True Then
'            Call mobjMipClient(intLoop).MipModule.ConnectStateChanged(IsConnected)
'        End If
'    Next
'
'End Sub
'
'Private Sub mclsMipClientShell_OpenLink(ByVal strMsgKey As String, ByVal bytLinkType As Byte, ByVal strLinkPara As String, ByVal strModuleObjectKey As String)
'
'    Dim intCount As Integer
'
'    For intCount = 1 To 100
'        If mobjMipClient(intCount).OnLine = True Then
'
'            If strModuleObjectKey = "" Then
'                If mobjMipClient(intCount).MipModule.Module = 0 And mobjMipClient(intCount).MipModule.System = 0 Then
'                    Call mobjMipClient(intCount).MipModule.OpenLink(strMsgKey, bytLinkType, strLinkPara)
'                    Exit For
'                End If
'
'            Else
'                If mobjMipClient(intCount).MipModule.Instance = strModuleObjectKey Then
'                    Call mobjMipClient(intCount).MipModule.OpenLink(strMsgKey, bytLinkType, strLinkPara)
'                    Exit For
'                End If
'            End If
'        End If
'    Next
'End Sub
'
'Private Sub mclsMipClientShell_ReceiveMessage(ByVal strMsgContent As String)
'    Dim intCount As Integer
'    Dim intLoop As Integer
'    Dim strModule As String
'    Dim strNote As String
'    Dim strItemKey As String
'    Dim strReceiveLogKey As String
'    Dim strMsgid As String
'    Dim blnExisCommon As Boolean
'    Dim blnAllow As Boolean
'    Dim objNode As Object
'    Dim strTemp As String
'    Dim varTemp As Variant
'    Dim objCollectModule As New Collection
'
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "客户端服务收到来自消息服务平台的消息" & vbCrLf & strMsgContent)
'
'    blnExisCommon = False
'    blnAllow = False
'    '------------------------------------------------------------------------------------------------------------------
'    '分析消息内容
'
'    Set mobjXML = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
'    If mobjXML.loadXML(strMsgContent) = True Then
'        'msgid
'        Set objNode = mobjXML.selectSingleNode(".//head/msg_id")
'        strMsgid = objNode.Text
'
'        'itemidentity
'        Set objNode = mobjXML.selectSingleNode(".//head/msg_item_code")
'        strItemKey = objNode.Text
'
'    End If
'    If strMsgid = "" Or strItemKey = "" Then Exit Sub
'
'    '检查消息是否为当前工作站的消息
'    '------------------------------------------------------------------------------------------------------------------
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "检查消息是否为当前工作站的消息")
'
'    If CheckIsAllowReceive(strItemKey, strMsgContent, blnExisCommon, objCollectModule) = False Then
'        If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "当前消息并非是当前工作站的消息")
'        Exit Sub
'    End If
'
'    '分发消息
'    '------------------------------------------------------------------------------------------------------------------
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "分发消息到相应的模块")
'    strModule = ""
'    For intCount = 1 To 100
'        If mobjMipClient(intCount).OnLine = True Then
'            If mobjMipClient(intCount).Module = 0 Then
'
'                If blnExisCommon = True Then
'                    If Not (mobjMipClient(intCount).MipModule Is Nothing) Then
'
'                        If Not (mclsMipClientShell Is Nothing) Then
'                            If mobjMipClient(intCount).MipModule.ValidCommonMessageSchema(strMsgContent) Then
'                                Call mclsMipClientShell.ShowMessage(strItemKey, strMsgContent, mobjMipClient(intCount).MipModule.Instance)
'                                strModule = strModule & ",公共冒泡提醒"
'                            Else
'                                Call mobjMipClient(intCount).MipModule.ReceiveMessage(strItemKey, strMsgContent)
'                                strModule = strModule & ",导航台"
'                            End If
'                        End If
'                    End If
'
'                End If
'
'            Else
'                 If objCollectModule.count > 0 Then
'                    For intLoop = 1 To objCollectModule.count
'                        strTemp = objCollectModule.Item(intLoop)
'                        varTemp = Split(strTemp, "_")
'                        If varTemp(2) = "基本" Then varTemp(2) = ""
'
'                        If mobjMipClient(intCount).System = Val(varTemp(0)) And mobjMipClient(intCount).Module = Val(varTemp(1)) And (InStr(mobjMipClient(intCount).Privs, ";" & Trim(varTemp(2)) & ";") > 0 Or varTemp(2) = "") Then
'                            If Not (mobjMipClient(intCount).MipModule Is Nothing) Then
'                                Call mobjMipClient(intCount).MipModule.ReceiveMessage(strItemKey, strMsgContent)
'
'                                Select Case LenB(strModule & "," & Trim(varTemp(2)))
'                                Case Is > 88
'
'                                    mrsModule.Filter = ""
'                                    mrsModule.Filter = "系统=" & Val(varTemp(0)) & " And 序号=" & Val(varTemp(1))
'                                    If mrsModule.RecordCount > 0 Then
'                                        strModule = strModule & "," & mrsModule("标题").value & IIf(varTemp(2) = "", "", "(" & Trim(varTemp(2)) & ")") & "等"
'                                    End If
'
'                                Case Is > 90
'
'                                Case Else
'                                    mrsModule.Filter = ""
'                                    mrsModule.Filter = "系统=" & Val(varTemp(0)) & " And 序号=" & Val(varTemp(1))
'                                    If mrsModule.RecordCount > 0 Then
'                                        strModule = strModule & "," & mrsModule("标题").value & IIf(varTemp(2) = "", "", "(" & Trim(varTemp(2)) & ")")
'                                    End If
'
'                                End Select
'
'                            End If
'                        End If
'                    Next
'                 End If
'            End If
'        End If
'    Next
'    If strModule <> "" Then
'        strModule = Mid(strModule, 2)
'        If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", strModule)
'    Else
'        If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "当前消息没有可执行的模块")
'    End If
'
'    '填写接收日志
'    '------------------------------------------------------------------------------------------------------------------
'    strReceiveLogKey = GetGUID
'
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("信息", "登记接收记录" & vbCrLf & "ID=" & strReceiveLogKey)
'
'    Call WriteReceiveLog(strReceiveLogKey, strMsgid, strItemKey, mstrMspUser, mstrUser, mstrUserName, strMsgContent, strModule, strNote)
'
'    Exit Sub
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mstrError = Err.Description
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "接收消息过程中出错：" & vbCrLf & mstrError)
'End Sub
'
''######################################################################################################################
'Private Sub WriteSendLog(ByVal strSendLogKey As String, ByVal strMessageKey As String, ByVal strMsgIdentity As String, ByVal strMsgContent As String)
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim intLoop As Integer
'    Dim intCount As Integer
'    Dim intSQL As Integer
'    Dim strSQL() As String
'    Dim sgltimer As Single
'
'    If LenB(strMsgContent) > 4000 Then
'
'        mstrSQL = "zltools.zlMip_SendLog_Insert("
'        mstrSQL = mstrSQL & "'" & strSendLogKey & "'"
'        mstrSQL = mstrSQL & ",'" & strMessageKey & "'"
'        mstrSQL = mstrSQL & ",'" & strMsgIdentity & "'"
'        mstrSQL = mstrSQL & ",Null"
'        mstrSQL = mstrSQL & ")"
'
'        intSQL = intSQL + 1
'        ReDim Preserve strSQL(intSQL)
'        strSQL(intSQL) = mstrSQL
'
'        intCount = 0 - Int(0 - Len(strMsgContent) / 2000)
'        For intLoop = 1 To intCount
'            intSQL = intSQL + 1
'            ReDim Preserve strSQL(intSQL)
'            strSQL(intSQL) = "zltools.zlMip_Clob_Append(1,'" & strSendLogKey & "','" & Mid(strMsgContent, CLng(intLoop) * 2000 - 1999, 2000) & "'," & IIf(intLoop = 1, 1, 0) & "," & IIf(intLoop = intCount, 1, 0) & ")"
'        Next
'    Else
'        mstrSQL = "zltools.zlMip_SendLog_Insert("
'        mstrSQL = mstrSQL & "'" & strSendLogKey & "'"
'        mstrSQL = mstrSQL & ",'" & strMessageKey & "'"
'        mstrSQL = mstrSQL & ",'" & strMsgIdentity & "'"
'        mstrSQL = mstrSQL & ",'" & strMsgContent & "'"
'        mstrSQL = mstrSQL & ")"
'
'        intSQL = intSQL + 1
'        ReDim Preserve strSQL(intSQL)
'        strSQL(intSQL) = mstrSQL
'    End If
'
'    On Error GoTo errHand
'
''    sgltimer = timer
'    For intLoop = 1 To intSQL
'        Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL(intLoop), mstrTitle)
'    Next
'
''    MsgBox Format(timer - sgltimer, "0.000")
'
'    Exit Sub
'errHand:
''    Resume
'End Sub
'
'Private Sub WriteReceiveLog(ByVal strReceiveLogKey As String, _
'                            ByVal strMsgid As String, _
'                            ByVal strItemKey As String, _
'                            ByVal mstrMspUser As String, _
'                            ByVal mstrUser As String, _
'                            ByVal mstrUserName As String, _
'                            ByVal strMsgContent As String, _
'                            ByVal strModule As String, _
'                            ByVal strNote As String)
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim intLoop As Integer
'    Dim intCount As Integer
'    Dim intSQL As Integer
'    Dim strSQL() As String
'
'    If LenB(strMsgContent) > 4000 Then
'
'        mstrSQL = "zltools.zlMip_ReceiveLog_Insert("
'        mstrSQL = mstrSQL & "'" & strReceiveLogKey & "'"
'        mstrSQL = mstrSQL & ",'" & strMsgid & "'"
'        mstrSQL = mstrSQL & ",'" & strItemKey & "'"
'        mstrSQL = mstrSQL & ",'" & mstrMspUser & "'"
'        mstrSQL = mstrSQL & ",'" & mstrUser & "'"
'        mstrSQL = mstrSQL & ",'" & mstrUserName & "'"
'        mstrSQL = mstrSQL & ",'" & mstrStationName & "(" & mstrStationIP & ")" & "'"
'        mstrSQL = mstrSQL & ",'" & strModule & "'"
'        mstrSQL = mstrSQL & ",Null"
'        mstrSQL = mstrSQL & ",'" & strNote & "'"
'        mstrSQL = mstrSQL & ")"
'
'        intSQL = intSQL + 1
'        ReDim Preserve strSQL(intSQL)
'        strSQL(intSQL) = mstrSQL
'
'        intCount = 0 - Int(0 - Len(strMsgContent) / 2000)
'        For intLoop = 1 To intCount
'            intSQL = intSQL + 1
'            ReDim Preserve strSQL(intSQL)
'            strSQL(intSQL) = "zltools.zlMip_Clob_Append(2,'" & strReceiveLogKey & "','" & Mid(strMsgContent, CLng(intLoop) * 2000 - 1999, 2000) & "'," & IIf(intLoop = 1, 1, 0) & "," & IIf(intLoop = intCount, 1, 0) & ")"
'        Next
'    Else
'        mstrSQL = "zltools.zlMip_ReceiveLog_Insert("
'        mstrSQL = mstrSQL & "'" & strReceiveLogKey & "'"
'        mstrSQL = mstrSQL & ",'" & strMsgid & "'"
'        mstrSQL = mstrSQL & ",'" & strItemKey & "'"
'        mstrSQL = mstrSQL & ",'" & mstrMspUser & "'"
'        mstrSQL = mstrSQL & ",'" & mstrUser & "'"
'        mstrSQL = mstrSQL & ",'" & mstrUserName & "'"
'        mstrSQL = mstrSQL & ",'" & mstrStationName & "(" & mstrStationIP & ")" & "'"
'        mstrSQL = mstrSQL & ",'" & strModule & "'"
'        mstrSQL = mstrSQL & ",'" & strMsgContent & "'"
'        mstrSQL = mstrSQL & ",'" & strNote & "'"
'        mstrSQL = mstrSQL & ")"
'
'        intSQL = intSQL + 1
'        ReDim Preserve strSQL(intSQL)
'        strSQL(intSQL) = mstrSQL
'    End If
'
'    On Error GoTo errHand
'
'    For intLoop = 1 To intSQL
'        Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL(intLoop), mstrTitle)
'    Next
'
'    Exit Sub
'errHand:
'    '
'End Sub
'
'Private Function CheckIsAllowReceive(ByVal strItemKey As String, ByVal strMessageContent As String, ByRef blnAllowCommon As Boolean, ByRef objCollectModule As Collection) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim blnAllowReceive As Boolean
'    Dim rsTmp As New ADODB.Recordset
'    Dim strDeliverXML As String
'    Dim blnGroupPersonKey As Boolean
'    Dim blnGroupDeptKey As Boolean
'    Dim blnGroupPersonProperty As Boolean
'    Dim blnGroupPersonRole As Boolean
'    Dim blnGroupStation As Boolean
'    Dim blnGroupModule As Boolean
'    Dim blnGroupAllow As Boolean
'
'    On Error GoTo errHand
'
'    Set objCollectModule = New Collection
'
'    mstrSQL = "Select b.deliver_object.getstringval() as deliver_object From zltools.zlmip_item a,zltools.zlmip_item_deliver b Where a.item_code=[1] And a.id=b.item_id Order By b.deliver_order"
'    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, strItemKey)
'    If rsTmp.BOF = False Then
'
'        Do While Not rsTmp.EOF
'
'            strDeliverXML = gobjComLib.zlCommFun.NVL(rsTmp("deliver_object").value)
'
'            If strDeliverXML <> "" Then
'                blnGroupPersonKey = False
'                blnGroupDeptKey = False
'                blnGroupPersonProperty = False
'                blnGroupPersonRole = False
'                blnGroupStation = False
'                blnGroupModule = False
'                blnGroupAllow = False
'
'                Set mobjXML = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
'                Set mobjXML2 = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
'                Call mobjXML2.loadXML(strMessageContent)
'
'                If mobjXML.loadXML(strDeliverXML) = True Then
'
'                    blnGroupPersonKey = CheckGroupIsAllowReceive(1, ".//persons/person/key", ".//persons/message")
'                    blnGroupDeptKey = CheckGroupIsAllowReceive(2, ".//depts/dept/key", ".//messagedaa/message")
'                    blnGroupPersonProperty = CheckGroupIsAllowReceive(3, ".//personpropertys/personproperty", ".//personpropertys/message")
'                    blnGroupPersonRole = CheckGroupIsAllowReceive(4, ".//roles/role", ".//roles/message")
'                    blnGroupStation = CheckGroupIsAllowReceive(5, ".//stations/station", ".//stations/message")
'
'                    blnGroupModule = CheckGroupIsAllowReceive(6, ".//modules/module/key", ".//modules/message", objCollectModule)
'
'                End If
'
'                blnGroupAllow = (blnGroupPersonKey And blnGroupDeptKey And blnGroupPersonProperty And blnGroupPersonRole And blnGroupStation)
'
'                If blnAllowReceive = False Then blnAllowReceive = blnGroupAllow
'                If blnAllowCommon = False Then blnAllowCommon = (blnGroupModule = False And blnGroupAllow = True)
'            Else
'                blnAllowReceive = True
'                blnAllowCommon = True
'            End If
'
'            rsTmp.MoveNext
'        Loop
'    Else
'        blnAllowCommon = True
'        blnAllowReceive = True
'    End If
'
'    CheckIsAllowReceive = blnAllowReceive
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mstrError = Err.Description
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "检查消息是否为当前工作站过程中出错：" & vbCrLf & mstrError)
'End Function
'
'Private Function CheckGroupIsAllowReceive(ByVal bytMode As Byte, ByVal strProductConfig As String, ByVal strMessageConfig As String, Optional ByRef objCollectModule As Collection) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim blnAllowReceive As Boolean
'    Dim objNodeList As Object
'    Dim objNodeProductList As Object
'    Dim objNodeMessageList As Object
'    Dim objNode As Object
'    Dim objNode2 As Object
'    Dim objTmp As Object
'
'    On Error GoTo errHand
'
'    blnAllowReceive = False
'
'    Set objNodeProductList = mobjXML.selectNodes(strProductConfig)
'    Set objNodeMessageList = mobjXML.selectNodes(strMessageConfig)
'
'    Select Case bytMode
'    '------------------------------------------------------------------------------------------------------------------
'    Case 1, 2, 3, 4, 5
'        If objNodeProductList.Length = 0 And objNodeMessageList.Length = 0 Then
'            blnAllowReceive = True
'        Else
'            Select Case bytMode
'            '----------------------------------------------------------------------------------------------------------
'            Case 1          '人员
'                If objNodeProductList.Length > 0 Then
'                    For Each objNode In objNodeProductList
'                        If Val(objNode.Text) = mlngPersonKey Then
'                            blnAllowReceive = True
'                            Exit For
'                        End If
'                    Next
'                End If
'
'                If blnAllowReceive = False Then
'                    If objNodeMessageList.Length > 0 Then
'                        For Each objNode In objNodeMessageList
'                            Set objNodeList = mobjXML2.selectNodes("./" & Trim(objNode.Text))
'                            If objNodeList.Length = 0 Then
'                                blnAllowReceive = True
'                            Else
'                                For Each objNode2 In objNodeList
'                                    If Val(objNode2.Text) = mlngPersonKey Then
'                                        blnAllowReceive = True
'                                        Exit For
'                                    End If
'                                Next
'                            End If
'                        Next
'                    End If
'                End If
'            '----------------------------------------------------------------------------------------------------------
'            Case 2          '部门
'                If objNodeProductList.Length > 0 Then
'                    For Each objNode In objNodeProductList
'                        If Val(objNode.Text) = mlngPersonDeptKey Then
'                            blnAllowReceive = True
'                            Exit For
'                        End If
'                    Next
'                End If
'
'                If blnAllowReceive = False Then
'                    If objNodeMessageList.Length > 0 Then
'                        For Each objNode In objNodeMessageList
'                            Set objNodeList = mobjXML2.selectNodes("./" & Trim(objNode.Text))
'                            If objNodeList.Length = 0 Then
'                                blnAllowReceive = True
'                            Else
'                                For Each objNode2 In objNodeList
'                                    If Val(objNode2.Text) = mlngPersonDeptKey Then
'                                        blnAllowReceive = True
'                                        Exit For
'                                    End If
'                                Next
'                            End If
'                        Next
'                    End If
'                End If
'            '----------------------------------------------------------------------------------------------------------
'            Case 3          '人员性质
'                If objNodeProductList.Length > 0 Then
'                    For Each objNode In objNodeProductList
'                        Set objTmp = Nothing
'                        On Error Resume Next
'                        Set objTmp = mobjCollectProperty.Item("K_" & Trim(objNode.Text))
'                        On Error GoTo 0
'                        If Not (objTmp Is Nothing) Then
'                            blnAllowReceive = True
'                            Exit For
'                        End If
'                    Next
'                End If
'
'                If blnAllowReceive = False Then
'                    If objNodeMessageList.Length > 0 Then
'                        For Each objNode In objNodeMessageList
'                            Set objNodeList = mobjXML2.selectNodes("./" & Trim(objNode.Text))
'                            If objNodeList.Length = 0 Then
'                                blnAllowReceive = True
'                            Else
'                                For Each objNode2 In objNodeList
'                                    Set objTmp = Nothing
'                                    On Error Resume Next
'                                    Set objTmp = mobjCollectProperty.Item("K_" & Trim(objNode2.Text))
'                                    On Error GoTo 0
'                                    If Not (objTmp Is Nothing) Then
'                                        blnAllowReceive = True
'                                        Exit For
'                                    End If
'                                Next
'                            End If
'                        Next
'                    End If
'                End If
'            '----------------------------------------------------------------------------------------------------------
'            Case 4          '人员角色
'                If objNodeProductList.Length > 0 Then
'                    For Each objNode In objNodeProductList
'                        Set objTmp = Nothing
'                        On Error Resume Next
'                        Set objTmp = mobjCollectRole.Item("K_" & Trim(objNode.Text))
'                        On Error GoTo 0
'                        If Not (objTmp Is Nothing) Then
'                            blnAllowReceive = True
'                            Exit For
'                        End If
'                    Next
'                End If
'
'                If blnAllowReceive = False Then
'                    If objNodeMessageList.Length > 0 Then
'                        For Each objNode In objNodeMessageList
'                            Set objNodeList = mobjXML2.selectNodes("./" & Trim(objNode.Text))
'                            If objNodeList.Length = 0 Then
'                                blnAllowReceive = True
'                            Else
'                                For Each objNode2 In objNodeList
'                                    Set objTmp = Nothing
'                                    On Error Resume Next
'                                    Set objTmp = mobjCollectRole.Item("K_" & Trim(objNode2.Text))
'                                    On Error GoTo 0
'                                    If Not (objTmp Is Nothing) Then
'                                        blnAllowReceive = True
'                                        Exit For
'                                    End If
'                                Next
'                            End If
'                        Next
'                    End If
'                End If
'            '----------------------------------------------------------------------------------------------------------
'            Case 5          '站点
'                If objNodeProductList.Length > 0 Then
'                    For Each objNode In objNodeProductList
'                        If Trim(objNode.Text) = Trim(mstrStationIP) Then
'                            blnAllowReceive = True
'                            Exit For
'                        End If
'                    Next
'                End If
'
'                If blnAllowReceive = False Then
'                    If objNodeMessageList.Length > 0 Then
'                        For Each objNode In objNodeMessageList
'                            Set objNodeList = mobjXML2.selectNodes("./" & Trim(objNode.Text))
'                            If objNodeList.Length = 0 Then
'                                blnAllowReceive = True
'                            Else
'                                For Each objNode2 In objNodeList
'                                    If Trim(objNode2.Text) = Trim(mstrStationIP) Then
'                                        blnAllowReceive = True
'                                        Exit For
'                                    End If
'                                Next
'                            End If
'                        Next
'                    End If
'                End If
'            End Select
'        End If
'    '------------------------------------------------------------------------------------------------------------------
'    Case 6          '模块
'        If objNodeProductList.Length > 0 Then
'            For Each objNode In objNodeProductList
'                On Error Resume Next
'                Call objCollectModule.Add(Trim(objNode.Text), "K_" & Trim(objNode.Text))
'                On Error GoTo errHand
'                If Trim(objNode.Text) <> "" Then blnAllowReceive = True
'            Next
'        End If
'
'        If objNodeMessageList.Length > 0 Then
'            For Each objNode In objNodeMessageList
'                Set objNodeList = mobjXML2.selectNodes("./" & Trim(objNode.Text))
'                If objNodeList.Length > 0 Then
'                    For Each objNode2 In objNodeList
'                        On Error Resume Next
'                        Call objCollectModule.Add(Trim(objNode2.Text), "K_" & Trim(objNode2.Text))
'                        On Error GoTo errHand
'                        If Trim(objNode.Text) <> "" Then blnAllowReceive = True
'                    Next
'                End If
'            Next
'        End If
'    End Select
'
'    CheckGroupIsAllowReceive = blnAllowReceive
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mstrError = Err.Description
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "检查消息是否为当前工作站过程中出错：" & vbCrLf & mstrError)
'End Function
'
'Private Function AnalyseComputer() As String
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim strComputer As String * 256
'
'    Call GetComputerName(strComputer, 255)
'    AnalyseComputer = strComputer
'    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
'
'End Function
'
'Private Function AnalyseIP() As String
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim ret As Long, Tel As Long
'    Dim bBytes() As Byte
'    Dim TempList() As String
'    Dim TempIP As String
'    Dim Tempi As Long
'    Dim Listing As MIB_IPADDRTABLE
'    Dim L3 As String
'
'    On Error GoTo errHand
'
'    GetIpAddrTable ByVal 0&, ret, True
'
'
'    If ret <= 0 Then Exit Function
'    ReDim bBytes(0 To ret - 1) As Byte
'    ReDim TempList(0 To ret - 1) As String
'
'    'retrieve the data
'    GetIpAddrTable bBytes(0), ret, False
'
'    'Get the first 4 bytes to get the entry's.. ip installed
'    CopyMemory Listing.dEntrys, bBytes(0), 4
'
'    For Tel = 0 To Listing.dEntrys - 1
'        'Copy whole structure to Listing..
'        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
'        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
'    Next Tel
'    'Sort Out The IP For WAN
'        TempIP = TempList(0)
'        For Tempi = 0 To Listing.dEntrys - 1
'            L3 = Left(TempList(Tempi), 3)
'            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
'                TempIP = TempList(Tempi)
'            End If
'        Next Tempi
'        AnalyseIP = TempIP 'Return The TempIP
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mstrError = Err.Description
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "分析本工作站IP地址过程中出错：" & vbCrLf & mstrError)
'End Function
'
'Private Function ConvertAddressToString(longAddr As Long) As String
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '说明：
'    '******************************************************************************************************************
'    Dim myByte(3) As Byte
'    Dim Cnt As Long
'
'    On Error GoTo errHand
'
'    CopyMemory myByte(0), longAddr, 4
'    For Cnt = 0 To 3
'        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
'    Next Cnt
'    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mstrError = Err.Description
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "转换本工作站IP地址过程中出错：" & vbCrLf & mstrError)
'End Function
'
'Private Function GetIpAddress(Optional ByVal strDefaultIp_Address As String = "") As String
'    '******************************************************************************************************************
'    '功能:通过oracle获取的计算机的IP地址
'    '入参:strDefaultIp_Address-缺省IP地址
'    '出参:
'    '返回:返回IP地址
'    '******************************************************************************************************************
'    Dim rsTemp As ADODB.Recordset
'    Dim strIp_Address As String
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
'    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取IP地址")
'    If rsTemp.EOF = False Then
'        strIp_Address = gobjComLib.zlCommFun.NVL(rsTemp!Ip_Address)
'    End If
'    If strIp_Address = "" Then strIp_Address = strDefaultIp_Address
'    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
'
'    GetIpAddress = strIp_Address
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    mstrError = Err.Description
'    If Not (mclsMipClientShell Is Nothing) Then Call mclsMipClientShell.WriteRunLog("错误", "获取本工作站IP地址过程中出错：" & vbCrLf & mstrError)
'End Function
'
'Private Function GetXMLVersion() As String
'
'    Dim varXMLVersion As Variant
'    Dim strXMLVer As String
'    Dim intLoop As Integer
'    Dim objXML As Object
'
'    On Error GoTo errHand
'
'    varXMLVersion = Split(".6.0,.4.0", ",")
'
'    On Error Resume Next
'    If gobjComLib.OS.IsDesinMode = True Or gobjComLib.zlRegInfo("授权性质") <> "1" Then
'        For intLoop = 0 To UBound(varXMLVersion)
'            Err = 0
'            Set objXML = CreateObject("MSXML2.DOMDocument" & varXMLVersion(intLoop))
'            If Err = 0 Then
'                strXMLVer = varXMLVersion(intLoop)
'                Exit For
'            End If
'        Next
'
'        On Error GoTo errHand
'
'        If strXMLVer = "" Then
'            MsgBox "创建MSXML2.DOMDocument对象失败"
'            Exit Function
'        End If
'    Else
'        strXMLVer = ""
'    End If
'    GetXMLVersion = strXMLVer
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    MsgBox Err.Description
'End Function

