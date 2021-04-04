VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWinsock 
   ClientHeight    =   825
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   2250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   2250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   1320
      Top             =   120
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event DataArrival(ByVal lngNoticeCode As Long, ByVal intChangeType As Integer, ByVal strTableOwner As String, ByVal strTableName As String, ByVal strRowid As String)
Public blnDcnState As Boolean

Private mlngCheckDcnInterval As Long    '检查dcn服务器状态时间间隔
Private mlngCheck As Long

Private mcolInterval As New Collection  '保存通知Interval
Private mcolData As New Collection '保存变动信息
Private mcolTime As New Collection '保存变动信息缓存时间

Public Function StartUdp(ByVal lngPort As Long, Optional ByRef strError As String) As Boolean
    '开启端口,成功返回True
    On Error Resume Next

    winSock.LocalPort = lngPort
    winSock.Bind

    If Err.Number <> 0 Then
        strError = Err.Description
        Exit Function
    End If
    StartUdp = True
End Function


Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim intType As Integer, strOwner As String, strTable As String
    Dim strRowid As String, lngNoticeCode As Long
    Dim lngInterval As Long, arrTmp() As String
    
    On Error GoTo errH
    
    winSock.GetData strData
    If strData = "" Then Exit Sub
    
    '获取检查DCN的时间间隔
    If mlngCheckDcnInterval = 0 Then
        mlngCheckDcnInterval = GetCheckInterval
        blnDcnState = GetDcnState
    End If
    
    '消息格式:   NoticeCode-变动类型-所有者-表名-Rowid
    arrTmp = Split(strData, "-")
    lngNoticeCode = arrTmp(0): intType = arrTmp(1)
    strOwner = arrTmp(2): strTable = arrTmp(3): strRowid = arrTmp(4)
    
    lngInterval = GetNoticeInterval(lngNoticeCode)
    
    If lngInterval = 0 Then '如果Interval为0 ,说明不需要进行缓存,直接抛出事件
        RaiseEvent DataArrival(lngNoticeCode, intType, strOwner, strTable, strRowid)
    Else
        If GetValueFromList(mcolInterval, lngNoticeCode) = "" Then  '集合中没有找到noticecode,就添加到缓存中,达到设定时间后再抛出事件
            mcolData.Add strData, "_" & lngNoticeCode
            mcolInterval.Add lngInterval, "_" & lngNoticeCode
            mcolTime.Add 0, "_" & lngNoticeCode
        End If
    End If
    
    Exit Sub
errH:
    gobjComLib.ErrCenter
End Sub


Private Sub Timer_Timer()
    Dim i As Integer, arrTmp() As String
    Dim intType As Integer, strOwner As String, strTable As String
    Dim strRowid As String, lngNoticeCode As Long
    Dim lngInterval As Long, arrRemove() As Long
    
    '数据缓存
    ReDim arrRemove(0)
    
    For i = 1 To mcolTime.Count
        mcolTime.Item(i) = mcolTime.Item(i) + 1
        
        If mcolTime.Item(i) >= mcolInterval.Item(i) Then    '达到刷新时间
            arrTmp = Split(mcolData.Item(i), "-")
            lngNoticeCode = arrTmp(0): intType = arrTmp(1)
            strOwner = arrTmp(2): strTable = arrTmp(3): strRowid = arrTmp(4)
                    
            RaiseEvent DataArrival(lngNoticeCode, intType, strOwner, strTable, strRowid)    '抛出事件
            
            ReDim Preserve arrRemove(UBound(arrRemove) + 1)     '将已经抛出事件的通知记录在数组中
            arrRemove(UBound(arrRemove)) = lngNoticeCode
        End If
    Next
    
    For i = 1 To UBound(arrRemove) '循环数组,删除已抛出事件通知
        mcolData.Remove "_" & arrRemove(i)
        mcolInterval.Remove "_" & arrRemove(i)
        mcolTime.Remove "_" & arrRemove(i)
    Next
    
    'Dcn状态检查
    If mlngCheckDcnInterval <> 0 Then
        mlngCheck = mlngCheck + 1
        
        If mlngCheck > mlngCheckDcnInterval Then
            blnDcnState = GetDcnState
            mlngCheck = 0
        End If
    End If
End Sub

Private Function GetNoticeInterval(lngNoticeCode As Long) As Long
    '功能:根据NoticeCode获取Interval
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select Interval From zltools.ZlnoticeLists where NoticeCode = [1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取Interval", lngNoticeCode)
    
    If rsTmp.RecordCount > 0 Then
        GetNoticeInterval = Val(rsTmp!Interval & "")
    End If
    
    Exit Function
errH:
    gobjComLib.ErrCenter
End Function

Private Function GetValueFromList(colList As Collection, lngIndex As Long) As String
    '功能:根据Index在集合中获取值,如果没有找到则返回空
    Dim strResult As String
    
    On Error Resume Next
    
    strResult = colList.Item(lngIndex)
    
    If Err.Number = 0 Then
        GetValueFromList = strResult
    End If
End Function


Private Function GetCheckInterval() As Long
    '获取 "DCN存活时间更新间隔"
    Dim rsTmp As New ADODB.Recordset
       
    Set rsTmp = GetZLOptions(32)
    If rsTmp.RecordCount > 0 Then GetCheckInterval = Val(rsTmp!参数值)
End Function

Private Function GetDcnState() As Boolean
    '检查DCN服务器是否正常运行
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 1" & vbNewLine & _
                "From (Select To_Date(参数值, 'yyyy-mm-dd hh24:mi:ss') 存活时间 From zlOptions Where 参数号 = 31) A," & vbNewLine & _
                "     (Select 参数值 更新间隔 From zlOptions Where 参数号 = 32) B" & vbNewLine & _
                "Where Sysdate < a.存活时间 + b.更新间隔 / 24 / 60"

    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "DCN存活时间更新间隔")
    
    GetDcnState = rsTmp.RecordCount > 0
    Exit Function
errH:
    gobjComLib.ErrCenter
End Function


