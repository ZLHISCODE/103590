VERSION 5.00
Begin VB.Form frmMipPollService 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2175
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "frmMipPollService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义
Private mblnConnected As Boolean
Private mstrCurrentPackageKey As String
Private mblnResponeAck As Boolean
Private mstrResponeAck As String
Private mblnError As Boolean
Private mblnRunning As Boolean
Private mintCurrent As Integer
Private mstrSQL As String
Private mblnStartUp As Boolean
Private mblnDataChanged As Boolean
Private mintValue As Integer
Private mintMessageTotal As Integer
Private mblnStartService As Boolean
Private mcolWinsock As New Collection
Private mstrRemoteHost As String
Private mlngRemotePort As Long
Private mstrTitle As String
Private mintConnectTime As Integer
Private mintSendTime As Integer
Private mintSendService As Integer

Private WithEvents mclstimer As clsTimer
Attribute mclstimer.VB_VarHelpID = -1
Private WithEvents mclsMipPoll As clsMipPoll
Attribute mclsMipPoll.VB_VarHelpID = -1
Private mclsMipServiceLog As clsMipServiceLog
Private mclsMipServiceData As clsMipServiceData

Private Type UseTime
    Total As Single
    MakeData As Single
    SendMessage As Single
    SendData As Single
    ReadHeadData As Single
    ReadLoopData As Single
    DeleteData As Single
    WaitRespone As Single
    MakePackage As Single
    InitWinsock As Single
    ConnectWinsock As Single
    WriteLog As Single
End Type

Private usrUseTime As UseTime

Public Event AfterStateInfoChange(ByVal intState As Integer, ByVal strInfo As String)

'######################################################################################################################
'接口方法

Public Function InitService() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsCondition As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rs As zlDataSQLite.SQLiteRecordset
    
    Set mclsMipServiceLog = New clsMipServiceLog
    Set mclsMipServiceData = New clsMipServiceData
    
    mintConnectTime = 5
    mintSendTime = 5
    mintSendService = 5
    
    If mclsMipServiceData.OpenFile(App.Path & "\Data\zlMspPollService.db") = True Then
        '取参数
        
        Set rsCondition = zlCommFun.CreateCondition
'        Call zlCommFun.SetCondition(rsCondition, "参数编号", "1")
'        rs = mclsMipServiceData.GetPara("Filter", rsCondition)
'        If rs.DataSet.BOF = False Then
'            strPara = zlCommFun.NVL(rs.DataSet("Content").Value)
'        End If
        
        Call zlCommFun.SetCondition(rsCondition, "参数编号", "2")
        rs = mclsMipServiceData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            mintConnectTime = Val(zlCommFun.NVL(rs.DataSet("Content").Value))
        End If
        
        Call zlCommFun.SetCondition(rsCondition, "参数编号", "3")
        rs = mclsMipServiceData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            mintSendTime = Val(zlCommFun.NVL(rs.DataSet("Content").Value))
        End If
        
        Call zlCommFun.SetCondition(rsCondition, "参数编号", "4")
        rs = mclsMipServiceData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            mintSendService = Val(zlCommFun.NVL(rs.DataSet("Content").Value))
        End If
        
    End If
        
    mclsMipServiceData.CloseFile
        
    Set mclsMipPoll = New clsMipPoll
    Set rsTmp = gclsBusiness.GetUserInfo(gstrDbUser)
    If rsTmp.BOF = False Then
        Call mclsMipPoll.Initialize(gstrDbUser, zlCommFun.NVL(rsTmp("姓名").Value))
    Else
        Call mclsMipPoll.Initialize(gstrDbUser)
    End If
    
    InitService = True
    
End Function

Public Function StartService() As Boolean
    '******************************************************************************************************************
    '功能：启动轮询服务
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    Call mclsMipServiceLog.OpenFile(App.Path & "\Data\zlMspPollServiceLog.db")
    Call mclsMipServiceLog.WriteRunLog("信息", "启动消息轮询服务")
    
    If mclsMipPoll.ConnectMip = False Then Exit Function
    
    DoEvents
        
    '启用定时器
    Set mclstimer = New clsTimer
    mclstimer.Interval = 1000
    
    StartService = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

Public Function StopService() As Boolean
    '******************************************************************************************************************
    '功能：停止轮询服务
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand
        
    Do While mblnRunning = True
        DoEvents
    Loop
        
    Call mclsMipPoll.DisConnectMip
    
    mclstimer.Interval = 0
    Set mclstimer = Nothing
    Call mclsMipServiceLog.WriteRunLog("信息", "停止消息轮询服务")
        
    StopService = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
    
End Function

Public Function ServerRunState() As Boolean
    ServerRunState = mblnRunning
End Function

Private Sub mclsMipPoll_AfterInfo(ByVal strInfoType As String, ByVal strInfoContent As String)
    Call mclsMipServiceLog.WriteRunLog(strInfoType, strInfoContent)
End Sub

'######################################################################################################################

Private Sub mclstimer_ThatTime()
    
    '1.处理时先停止定时器(即单进程）
    mclstimer.Interval = 0
    mblnRunning = True
    DoEvents
    
    '2.处理
    Call mclsMipPoll.RunPoll
    
    mblnRunning = False
    DoEvents
    
    '3.处理完再启用定时器
    mclstimer.Interval = Val(1000) * Val(60) * Val(mintSendService)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcolWinsock = Nothing
    Set mclstimer = Nothing
            
    If Not (mclsMipServiceLog Is Nothing) Then
        mclsMipServiceLog.CloseFile
        Set mclsMipServiceLog = Nothing
    End If
    
    If Not (mclsMipServiceData Is Nothing) Then
        mclsMipServiceData.CloseFile
        Set mclsMipServiceData = Nothing
    End If
    
    If Not (mclsMipPoll Is Nothing) Then Set mclsMipPoll = Nothing
    
End Sub

