VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChatNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Public Function InitChat(ByRef objMain As Object, ByVal strChatURL As String, ByVal strMyChatUrl As String, _
    ByVal strUser As String, Optional ByRef strMsg As String) As Boolean
      '功能:初始化
      '参数:
      'objMain-导航台对象
      'strChatURL-发起讨论地址
      'strMyChatUrl-我参与的讨论地址
      'strUser-当前操作员
      'strMsg             -返回错误信息(避免弹出模态提示引起主进程挂起,将提示信息返回给主进程处理。)
      '                    返回格式:提示类型[,]提示语句
          Dim strFilePath As String
          
1         On Error GoTo ErrH
          '如果本地没有日志文件新建日志文件
2         strFilePath = App.Path & "\" & App.EXEName & ".ini"
3         If Dir(strFilePath) = "" Then
              '[DEBUG]
              'OutPutDebugInfo = 0 ;输出调试信息到zlCISRule.log文件，0-不输出,1-输出所有接口参数信息
4             WriteIni "DEBUG", "OutPutDebugInfo", "0 ;输出调试信息到" & App.EXEName & ".log文件,0-不输出,1-输出", strFilePath
5         End If
6         gblnLog = (Val(ReadIni("Debug", "OutPutDebugInfo", strFilePath)) = 1)
          
7         WriteLog "函数：InitChat 开始" & vbNewLine & _
                         "入参：objMain=" & (Not objMain Is Nothing) & ",发起讨论地址=" & strChatURL & ",我的讨论地址=" & strMyChatUrl & _
                         ",当前用户=" & strUser & vbNewLine
8         gstrSysName = "通用讨论"
9         gstrChatURL = strChatURL
10        gstrMyChatUrl = strMyChatUrl
11        gstrUser = strUser
          
12        Set gobjMain = objMain
13        Set gfrmMain = New frmMain
14        Set gcolChat = New Collection
          
15        gfrmMain.Visible = False
          '加入系统托盘
16        Call gfrmMain.SetNotifyIcon(0)
          

17        InitChat = True
18        WriteLog "函数：InitChat 结束" & vbNewLine
                  
19        Exit Function

ErrH:
20        strMsg = vbExclamation & "[,]" & "在zlChatNotify.clsChatNotify.InitChat的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description
21        WriteLog strMsg & vbNewLine
End Function

Public Function SendChat(ByVal strSubject As String, ByVal strSysCode As String, _
    ByVal strMainCode As String, ByVal dblMainId As Double, ByVal strSender As String, _
    ByVal strReceivers As String, Optional ByRef strMsg As String) As Boolean
'参数:
'strSubject         -讨论标题
'strSysCode         -系统编码
'strMainCode        -主体编码
'dblMainId          -主体ID
'strSender        -发起人
'strReceivers    -参与人(多个参与人用分隔符";"分开)
'strMsg             -返回错误信息(避免弹出模态提示引起主进程挂起,将提示信息返回给主进程处理。)
'                    返回格式:提示类型[,]提示语句
    SendChat = OpenChatRoom("", strSubject, strSysCode, strMainCode, dblMainId, strSender, strReceivers, strMsg)
End Function

Public Function AddWaitList(ByVal rsList As ADODB.Recordset, Optional ByRef strMsg As String) As Boolean
'功能:将未读消息加入等待清单
'strMsg             -返回错误信息(避免弹出模态提示引起主进程挂起,将提示信息返回给主进程处理。)
'                    返回格式:提示类型[,]提示语句
          Dim i As Long
          
1         On Error GoTo ErrH

2         If grsList Is Nothing Then Call InitRsList
          '1-将未读消息加入等待清单
3         For i = 1 To rsList.RecordCount
4             Call grsList.AddNew(Array("ID", "Url", "Sys_Code", "Main_Code", "Main_ID", "Subject"), _
                  Array(Val(rsList!Id), rsList!Url & "", rsList!Sys_Code & "", rsList!Main_Code & "", rsList!Main_ID & "", rsList!Subject & ""))
5             rsList.MoveNext
6         Next
7         grsList.Filter = ""
8         If gblnShow Then
9             If Not frmChatList.RefreshList Then Exit Function
10        Else
              '2-开启系统托盘闪烁
11            Call gfrmMain.SetIcon(1)
12        End If
13        AddWaitList = True

14        Exit Function

ErrH:
15        strMsg = vbExclamation & "[,]" & "在zlChatNotify.clsChatNotify.AddWaitList的第" & Erl() & "行出错：" & vbCrLf & _
                  "错误号: " & Err.Number & vbCrLf & _
                  "错误描述：" & Err.Description
16        WriteLog strMsg & vbNewLine
End Function

Public Function ShowMyChat(ByVal lngUserId As Long, Optional ByRef strMsg As String) As Boolean
'功能:我参与的讨论
'strMsg   -返回错误信息(避免弹出模态提示引起主进程挂起,将提示信息返回给主进程处理。)
'         返回格式:提示类型[,]提示语句
    Dim strKey As String
    Dim objChat As frmChat
    WriteLog "函数：ShowMyChat 开始" & vbNewLine & _
                         "入参：用户ID=" & lngUserId & vbNewLine
    strKey = "K_USER_" & lngUserId
    On Error Resume Next
    Set objChat = gcolChat(strKey)
    On Error GoTo 0
    If objChat Is Nothing Then
        Set objChat = New frmChat
        gcolChat.Add objChat, strKey
    End If
    ShowMyChat = objChat.ShowMyChat(lngUserId, strMsg)
    WriteLog "函数：ShowMyChat 结束" & vbNewLine
End Function

Private Sub Class_Terminate()
    Dim i As Long
    
    On Error Resume Next
    Unload gfrmMain
    
    Set gfrmMain = Nothing
    Set gobjMain = Nothing
    Set grsList = Nothing
    For i = 1 To gcolChat.Count
        Unload gcolChat(i)
        gcolChat.Remove i
    Next
    Set gcolChat = Nothing
End Sub

 
