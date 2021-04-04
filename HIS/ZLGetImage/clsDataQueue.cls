VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjMsgQueue() As Object     '消息内容队列
Private mintQueueIndex As Integer    '记录队列中第一个消息的索引

Public Function MsgInQueue(objMsg As Object) As Boolean
'------------------------------------------------
'功能：消息入队，等候被处理
'参数： strMsg －－需要入队的消息内容，消息内容由存储地址，存储用户名，存储密码，存储目录，本机目录，存储类型组成
'返回：True－－入队成功，False－－入队失败
'-----------------------------------------------
    Dim intCount As Integer
    
    On Error GoTo err

    MsgInQueue = False
    
    '处理消息入队
    intCount = UBound(mobjMsgQueue) + 1
frmMain.BUGEX "MsgInQueue intCount=" & intCount & "    " & objMsg.SubDir
    ReDim Preserve mobjMsgQueue(intCount) As Object
    
    Set mobjMsgQueue(intCount) = objMsg
        
    MsgInQueue = True
    
    Exit Function
err:
    '出错就退出，暂时不做处理
End Function

Public Function MsgOutQueue() As Object
'------------------------------------------------
'功能：消息出队，调用过程负责处理出队的消息
'参数：
'返回：返回出队的消息内容
'-----------------------------------------------
    Dim intCount As Integer
    Dim objMsg As Object
    
    On Error GoTo err
    
    Set MsgOutQueue = Nothing        '初始化为空消息

    '消息出队处理
    intCount = UBound(mobjMsgQueue)

    If intCount = 0 Then Exit Function      '队列为空，不用出队
    
    '从队列中提取消息
    Set objMsg = mobjMsgQueue(mintQueueIndex)
frmMain.BUGEX "mintQueueIndex= " & mintQueueIndex & "    intCount= " & intCount & "    " & objMsg.SubDir
    '处理队列指针
    mintQueueIndex = mintQueueIndex + 1
    
    If mintQueueIndex > intCount Then
        '如果当前取出来的是队列中的最后一个消息，则清空消息队列
        ReDim Preserve mobjMsgQueue(0) As Object

        mintQueueIndex = 1
    End If
    
    Set MsgOutQueue = objMsg
    
    Exit Function
err:
    '出错就退出，暂时不做处理
frmMain.BUGEX "MsgOutQueue Err " & err.Description
End Function

Public Function QueueCount() As Integer
    QueueCount = UBound(mobjMsgQueue)
End Function

Private Sub Class_Initialize()
    ReDim mobjMsgQueue(0) As Object
    mintQueueIndex = 1
End Sub
