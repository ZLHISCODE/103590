Attribute VB_Name = "mdlDcn"
Option Explicit
Private Const GWL_WNDPROC = -4
Public Const GWL_USERDATA = (-21)
Public Const WM_SIZE = &H5
Public Const WM_USER = &H400
Public Const WM_BROADCAST = &H218
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Declare Function OCI_ConnCreate Lib "zlNoticeLib.dll" (ByVal strServer As String, ByVal strUser As String, ByVal strPwd As String) As Boolean
Public Declare Sub OCI_Register Lib "zlNoticeLib.dll" (ByVal lngHandler As Long, ByVal strTable As String)
Public Declare Sub OCI_UnRigister Lib "zlNoticeLib.dll" ()

Public lpPrevWndProc As Long    '窗体句柄
Public gcolNotice As New Collection '通知缓存集合
Public gstrBuild           As New clsStringBulider

Public gstrIp As String
Public glngPort As Long
Public glngSid As Long
Public gintState As Integer
Public gintLog As Integer
Public gintInterval As Integer

Public Function Hook(ByVal hwnd As Long) As Long
    '指定自定义的窗口过程
    lpPrevWndProc = GetWindowLong(hwnd, GWL_WNDPROC)
    SetWindowLong hwnd, GWL_WNDPROC, AddressOf WindowProc
    
    Hook = lpPrevWndProc
End Function

Public Sub UnHook(ByVal hwnd As Long)
    Dim temp As Long
    'Cease subclassing.
    temp = SetWindowLong(hwnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim byteRowid(100) As Byte, strNotice As String
    '调用原来的窗口过程
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
    
    If uMsg = WM_USER + 1 Then
        On Error Resume Next
        
        CopyMemory byteRowid(0), ByVal wParam, 100
        strNotice = StrConv(byteRowid, vbUnicode)
        strNotice = zlCommFun.TruncZero(strNotice)
        gcolNotice.Add strNotice
    End If
    
End Function


'--------------------------------------------------------------------------------------------------------
'数据操作
Public Function GetNoticeList() As ADODB.Recordset
    '功能:获取NoticeList,并返回一个记录集
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select Noticecode, Noticename, Tableowner, Tablename, Receivercols,  Changetype," & vbNewLine & _
                    "SplitChar,Noticekind, Comments, Status ,Filter,ReceiverTab ,ReceiverRelas,ReceiverIP ," & vbNewLine & _
                    "ReceiverStaffKind ,ReceiverDeptKind,Interval " & vbNewLine & _
                    "From Zltools.Zlnoticelists Order by 1"
                        
    Set GetNoticeList = zlDatabase.OpenSQLRecord(strSql, "获取NoticeList")
    Exit Function
errH:
    ErrCenter
End Function

Public Function GetClientPort() As ADODB.Recordset
    '功能:获取客户端UDP端口设置
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select 参数值 From zloptions Where 参数号=9"
                        
    Set GetClientPort = zlDatabase.OpenSQLRecord(strSql, "获取客户端UDP端口")
    Exit Function
errH:
    ErrCenter
End Function

Public Function GetJobs() As ADODB.Recordset
    Dim strSql  As String
    
    On Error GoTo errH
    strSql = "Select 间隔时间, Nvl(作业号, 0) 作业号 From zlAutoJobs Where 类型 = 3 And 序号 = 3 And 系统 Is Null"
                        
    Set GetJobs = zlDatabase.OpenSQLRecord(strSql, "获取客户端UDP端口")
    Exit Function
errH:
    ErrCenter
End Function

Public Function GetUserNotices() As ADODB.Recordset
    '功能:获取自定义通知,并返回一个记录集
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select a.序号, a.提醒内容, a.检查周期, a.提醒周期, To_Char(a.开始时间, 'yyyy/mm/dd hh24:mi') 开始时间," & vbNewLine & _
                "       To_Char(a.终止时间, 'yyyy/mm/dd hh24:mi') 终止时间" & vbNewLine & _
                "From zlNotices A" & vbNewLine & _
                "Order By 1"

    Set GetUserNotices = zlDatabase.OpenSQLRecord(strSql, "获取自定义通知")
    Exit Function
errH:
    ErrCenter
End Function

Public Function GetSid() As Long
    '功能:获取当前会话的Sid
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select Userenv('SessionID') AudSid From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取SID")
    GetSid = Val(rsTmp!AudSid & "")
    
    Exit Function
errH:
    ErrCenter
End Function

Public Function CheckSidState(ByVal lngSid As Long) As Boolean
    '检查指定的Sid是否在线
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 1 From GV$session Where Audsid = [1] And STATUS <> 'KILLED'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取SID", lngSid)
    CheckSidState = rsTmp.RecordCount <> 0
    Exit Function
errH:
    ErrCenter
End Function

Public Function UpdateDcnState2DB(ByVal intType As Integer) As Boolean
    '在数据库中修改消息收发器状态
    'intType=1:在线     intType=0:离线
    Dim strSql As String, rsClient As ADODB.Recordset
    Dim strValue As String
    On Error GoTo errH
    
    strValue = gstrIp & ";" & glngPort & ";" & intType & ";" & glngSid
    
    '由于调用部件时,已经向zlclientSession表中插入了当前会话信息,因此先根据会话号删除
    
    If intType = 1 Then
        strSql = "Delete From zltools.zlclientsession Where 会话号 = " & glngSid
        gcnOracle.Execute strSql
    End If
    
    '检查收发器状态
    strSql = "SELECT 参数值 FROM zltools.zloptions WHERE 参数号=[1]"
    Set rsClient = zlDatabase.OpenSQLRecord(strSql, "获取参数", 27)
    
    If rsClient.RecordCount > 0 Then
        If rsClient!参数值 & "" <> "" Then
            If intType = 1 And Split(rsClient!参数值 & "", ";")(2) = 1 And CheckSidState(Split(rsClient!参数值 & "", ";")(3)) Then
                MsgBox "数据变动通知服务已经在IP" & Split(rsClient!参数值 & "", ";")(0) & "已开启，无法再次开启！"
                Exit Function
            End If
        End If
        strSql = "Update Zltools.zlOptions Set 参数值 = '" & strValue & "' Where 参数号 =27"
        gcnOracle.Execute strSql
    Else
        MsgBox "基础数据缺失，请检查zlOption中的27号参数是否存在。", vbExclamation, "注意"
        Exit Function
    End If
    
    UpdateDcnState2DB = True
    Exit Function
errH:
    ErrCenter
End Function

Public Function ChangeServerSet2DB(ByVal lngPort As Integer, ByVal intLog As Integer, ByVal intInterval As Integer) As Boolean
    '功能:修改DCN服务器的配置
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strValue As String
    
    On Error GoTo errH
    glngPort = lngPort: gintLog = intLog: gintInterval = intInterval
    strValue = gstrIp & ";" & glngPort & ";" & gintState & ";" & glngSid
        
    strSql = "Update Zltools.zlOptions Set 参数值 = '" & strValue & "' Where 参数号 = 27"
    gcnOracle.Execute strSql
    
    strSql = "Update Zltools.zlOptions Set 参数值 = " & intLog & " Where 参数号 = 28"
    gcnOracle.Execute strSql
    
    strSql = "Update Zltools.zlOptions Set 参数值 = " & intInterval & " Where 参数号 = 29"
    gcnOracle.Execute strSql
    
    ChangeServerSet2DB = True
    Exit Function
errH:
    ErrCenter
End Function

Public Function ChangeClientSet2DB(ByVal lngPortS As Long, ByVal lngPortE As Long, ByVal lngCheckInterval As Long) As Boolean
    '功能:修改客户端消息接收端口和检查存活状态频率
    Dim strSql As String
    
    On Error GoTo errH
    
    '修改端口
    If lngPortS = 0 Or lngPortE = 0 Then
        strSql = "Update zloptions set 参数值 = 0 Where 参数号 = 9"
    Else
        strSql = "Update zloptions set 参数值 = '" & lngPortS & "-" & lngPortE & "' Where 参数号 = 9"
    End If
    gcnOracle.Execute strSql
        
    '修改检查存活状态的频率
    strSql = "Update zloptions set 参数值 = " & lngCheckInterval & " Where 参数号 = 32"
    gcnOracle.Execute strSql
    
    ChangeClientSet2DB = True

    Exit Function
errH:
    ErrCenter
End Function

Public Function ChangeJobSet2DB(ByVal intType As Integer, Optional ByVal intInterval As Integer = 5) As Boolean
    '功能:修改自动作业信息
    'intType =1: 提交自动任务  intType =2:修改自动任务 intType =3:删除自动任务
    'intInterval-自动任务执行频率,默认每5分钟执行一次
    Dim strSql As String
    
    On Error GoTo errH
    Select Case intType
    Case 1
        strSql = "Begin" & vbNewLine & _
                    "  Execute Immediate 'Update  zlAutoJobs Set 间隔时间 = " & intInterval & "  Where 类型 = 3 And 序号 = 3 And 系统 Is Null';" & vbNewLine & _
                    "  zltools.Zl_Jobsubmit(Null,3,3);" & vbNewLine & _
                    "End;"
    Case 2
        strSql = "Begin" & vbNewLine & _
                    "  Execute Immediate 'Update  zlAutoJobs Set 间隔时间 = " & intInterval & "  Where 类型 = 3 And 序号 = 3 And 系统 Is Null';" & vbNewLine & _
                    "  zltools.Zl_Jobchange(Null,3,3);" & vbNewLine & _
                    "End;"
    Case 3
        strSql = "Begin" & vbNewLine & _
                    "  zltools.Zl_Jobremove(Null,3,3);" & vbNewLine & _
                    "End;"
    End Select
    
    gcnZltools.Execute strSql
    ChangeJobSet2DB = True
    Exit Function
errH:
    ErrCenter
End Function

Private Function GetZltoolsConnection(ByVal strPwd As String) As Boolean
    '功能: 获取zltools连接对象
    
    On Error Resume Next
    
    If gcnZltools Is Nothing Then
        Set gcnZltools = New ADODB.Connection
    End If
    
    With gcnZltools
        .Provider = "OraOLEDB.Oracle"
        .Open "PLSQLRSet=1;Data Source=" & gstrServer, "ZLTOOLS", strPwd
      
        If .State = adStateOpen Then
            GetZltoolsConnection = True
        End If
    End With

End Function

Public Function GetZltools() As Boolean
    Dim blnResult As Boolean, strPwd As String
    
    blnResult = GetZltoolsConnection("ZLTOOLS")
    
    If Not blnResult Then
        blnResult = GetZltoolsConnection("ZLSOFT")
    End If
    
    If Not blnResult Then
        blnResult = frmUserCheckLogin.GetZltoolsByLogin
    End If
    
    GetZltools = blnResult
End Function

Public Function GetCheckInterval() As Long
    '获取 "DCN存活时间更新间隔"
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 参数值 From zltools.zlOptions Where 参数号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "DCN存活时间更新间隔", 32)
    
    GetCheckInterval = Val(rsTmp!参数值)
    Exit Function
errH:
    ErrCenter
End Function

Public Function UpdateDcnTime() As Boolean
    '功能:维护服务器最新存活时间
    Dim strSql As String, rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 1" & vbNewLine & _
                "From Dba_Change_Notification_Regs" & vbNewLine & _
                "Where Table_Name In (Select Tableowner || '.' || Tablename From Zlnoticelists)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查DCN状态")
    
    If rsTmp.RecordCount > 0 Then
        strSql = "Update zloptions set 参数值 = to_char(Sysdate,'YYYY-MM-DD hh24:mi:ss') where 参数号=31"
        gcnOracle.Execute strSql
    End If
    
    UpdateDcnTime = True
    Exit Function
errH:
    ErrCenter
End Function

Public Function UpdateNoticeInterval(ByVal lngNoticeCode As Long, ByVal lngInterval As Long)
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Update zlNoticeLists set Interval= " & lngInterval & " where NoticeCode=" & lngNoticeCode
    gcnOracle.Execute strSql
        
    UpdateNoticeInterval = True
    Exit Function
errH:
    ErrCenter
End Function
