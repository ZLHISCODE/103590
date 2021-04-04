VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHisCrust 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自动升级"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   ControlBox      =   0   'False
   Icon            =   "frmHisCrust.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   7050
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdLog 
      Caption         =   "查看日志(&C)"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "完成(&O)"
      Height          =   375
      Left            =   5325
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ListView lvwMan 
      Height          =   2430
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   4286
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "部件"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "错误信息"
         Object.Width           =   9402
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHisCrust.frx":030A
            Key             =   "Ok"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHisCrust.frx":08A4
            Key             =   "Err"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHisCrust.frx":0E3E
            Key             =   "List"
         EndProperty
      EndProperty
   End
   Begin zlHisCrustCom.UsrProgressBar prgPross 
      Height          =   264
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   6792
      _ExtentX        =   11986
      _ExtentY        =   450
      Color           =   16750899
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "客户端正在升级,请稍候..."
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1740
      TabIndex        =   4
      Top             =   480
      Width           =   4020
   End
   Begin VB.Image imgUpdate 
      Height          =   720
      Left            =   240
      Picture         =   "frmHisCrust.frx":13D8
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "正在注册部件"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   1530
      Width           =   1080
   End
End
Attribute VB_Name = "frmHisCrust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOP                          As Long = 0
Private Const SWP_NOSIZE                        As Long = &H1
Private mblnOperateCompleted                    As Boolean      '操作是否完成
Private mblnOK                                  As Boolean      '是否没有错误
Private mintColumn                              As Integer      '当前列
Private mintVB6                                 As Integer      'VB6进程的处理策略,0-尚未处理，1-保留VB6 2-杀掉VB6
Private mlngTimes                               As Long         'Tmr执行次数
Private Enum ErrListCol
    ELC_部件 = 0 '不能索引
    ELC_错误信息 = 1
End Enum

Private Enum ControlType
    CT_KillProc = 0         '只杀掉进程
    CT_KillProcAndSvr = 1   '停止服务并杀掉进程
    CT_StartSvr = 2         '启用服务
End Enum

Private Enum UpdateCheck
    UC_IgnorUp = 7          '升级中发生占用，忽略升级
    UC_SvrMD5Null = 6       '服务器MD5为空
    UC_NotExists = 5        '默认本地不存在
    UC_Normal = 4           '无需升级
    UC_AddtionUp = 3        '更新附加路径文件
    UC_RegAgain = 2         '上次注册未成功，需要再次注册
    UC_Update = 1           '需要下载更新
    UC_NewDown = 0          '需下载，本地不存在
End Enum
Private mcllOldComs         As Collection '老的部件且没有在清单中存在

Private Sub cmdLog_Click()
    Dim lngRet As Long
    Dim strNotPad As String
    
    On Error Resume Next
    strNotPad = gstrSystemPath & "\notepad.exe"
    If gobjFSO.FileExists(strNotPad) Then
        lngRet = ShellExecute(0&, "open", strNotPad, gobjTrace.LogFile, gobjTrace.LogFile, 5)    'SW_SHOW
        If lngRet = 31 Then
           If Not gblnHelperMain Then MsgBox "没有找到适当的程序来打开它,请安装有效的程序!", vbInformation, "客户端自动升级"
        End If
    Else
        If Not gblnHelperMain Then MsgBox "本机没有安装记事本程序,不能打开日志文件!" & vbCrLf & "请手工用其它程序打开,记事本路径为:" & vbCrLf & gobjTrace.LogFile, vbInformation, "客户端自动升级"
    End If
End Sub

Private Sub cmdOK_Click()
    If Not gblnHelperMain Then Call CallHISEXE(mblnOK)
    Call gobjTrace.CloseLog
    Call gobjMe.ExitApp
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call gobjMe.ExitApp
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    imgUpdate.Width = Me.ScaleWidth
    
    cmdOK.Top = ScaleHeight - cmdOK.Height - 50
    cmdOK.Left = ScaleWidth - cmdOK.Width - 100

    cmdLog.Top = ScaleHeight - cmdLog.Height - 50
    cmdLog.Left = ScaleWidth - cmdLog.Width - cmdOK.Width - 200
End Sub

Private Sub lvwMan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mblnOperateCompleted = False Then Exit Sub
    
    On Error Resume Next
    If mintColumn = ColumnHeader.Index - 1 Then
        lvwMan.SortOrder = IIf(lvwMan.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMan.SortKey = mintColumn
        lvwMan.SortOrder = lvwAscending
    End If
End Sub

Public Sub tmrStart_Timer()
    Dim lstTmp      As ListItem
    Call SetWindowPos(Me.hwnd, HWND_TOP, ((Screen.Width - Me.Width) / 2) / 15, ((Screen.Height - Me.Height) / 2) / 15, 0, 0, SWP_NOSIZE)
    Me.cmdOK.Caption = "取消(&C)"
    If gotCurType = OT_PreUpgrade Or gotCurType = OT_CheckFile Then
        Me.Hide
    Else
        Me.Show
    End If
    mlngTimes = 0
    prgPross.Value = 0
    prgPross.Min = 0
    prgPross.Max = 100
    lblInfor.Caption = "加载安装清单数据..."
    
    mintVB6 = 0
    '杀掉HIS相关进程
    Call ControlProcAndSvr
    MousePointer = 11
    If FileUpgrade = False Then
        gobjTrace.WriteSection "升级结束", SL_LevelOne
        gobjTrace.WriteInfo "升级结束", "升级结果", "存在一个或多个部件升级或注册未成功"
        cmdOK.Caption = "取消(&C)"
        MousePointer = 0
        If gblnHelperMain Then Call cmdOK_Click
    Else
        cmdOK.Caption = "完成(&O)"
        cmdOK.Visible = False
        '升级完成需退出
        mblnOK = True
        gobjTrace.WriteSection "升级结束", SL_LevelOne
        gobjTrace.WriteInfo "升级结束", "升级结果", "升级成功"
        Call cmdOK_Click
        MousePointer = 0
    End If
End Sub

Private Function ControlProcAndSvr(Optional ByVal lngCurPro As Long, Optional ByVal lngCurIncPro As Long, Optional ByVal ctType As ControlType = CT_KillProc) As Boolean
'功能：控制进程与服务，常规是杀掉停止
'参数：lngCurPro=当前进度
'      lngCurIncPro=当前过程执行完毕后的增量进度
'      ctType=0-不进行服务处理,1-进行进程的停止并进行进程杀掉,2-只启用服务
'返回：若发生不可预知错误则为False,否则为True
    Dim lngHwnd     As Long, lngZlhisHwnd   As Long, lngVBHwnd      As Long
    Dim lngProcess  As Long, lngPid         As Long
    Dim i           As Long, lngTotal       As Long
    Dim strReturn   As String, strErr       As String
    Dim objShell    As New clsShell
    
    On Error GoTo ErrH
    '如果预升级,就退出。
    If gotCurType = OT_PreUpgrade Or gotCurType = OT_CheckFile Then
        ControlProcAndSvr = True
        prgPross.Value = lngCurPro + lngCurIncPro
        Exit Function
    End If
    If ctType = CT_KillProcAndSvr Then
        grsFileUpgrade.Filter = "自动注册=" & RegFileType.RFT_NETServer
        lngTotal = grsFileUpgrade.RecordCount
        For i = 1 To lngTotal
            prgPross.Value = lngCurPro + lngCurIncPro * 0.2 * (i / lngTotal)
            If gobjFSO.FileExists(grsFileUpgrade!实际路径) Then
                gobjTrace.WriteInfo "ControlProcAndSvr", "停止服务", gobjFSO.GetBaseName(grsFileUpgrade!实际路径)
                lblInfor.Caption = "正在停止服务：" & gobjFSO.GetBaseName(grsFileUpgrade!实际路径)
                If objShell.Run("NET STOP " & gobjFSO.GetBaseName(grsFileUpgrade!实际路径), strReturn, strErr, 30000) Then
                End If
                gobjTrace.WriteInfo "ControlProcAndSvr", "启动结果", strReturn, "错误信息", strErr
            End If
            grsFileUpgrade.MoveNext
        Next
    End If
    If ctType <> CT_StartSvr Then
        Do While True
            lblInfor.Caption = "正在检查残余进程：ZLHIS+与VB6.EXE"
            prgPross.Value = lngCurPro + lngCurIncPro * 0.3
            lngHwnd = FindWindow(vbNullString, "导航台")
            If lngHwnd = 0 Then
               lngHwnd = FindWindow(vbNullString, "医院信息系统")
               If lngHwnd = 0 Then
                   Exit Do
               End If
            End If
            If lngHwnd <> 0 Then
                '区分是否是VB在调用导航台还是程序直接执行导航台
                lngZlhisHwnd = FindExitsProcess("ZLHIS+.EXE")
                If lngZlhisHwnd <> 0 Then
                    gobjTrace.WriteInfo "KillHisProcess", "杀掉进程", "ZLHISEXE"
                    Call TerminateProcess(lngZlhisHwnd, 1&)
                Else
                    lngVBHwnd = FindExitsProcess("VB6.EXE")
                    If lngVBHwnd <> 0 Then
                        If mintVB6 = 0 Then
                            If Not gblnHelperMain Then
                                If MsgBox("升级程序检测到VB6加载了可能会升级的部件." & vbCrLf & "为了保证系统正常升级,是否关闭VB6进程!", vbQuestion + vbYesNo, "客户端自动升级") = vbYes Then
                                    Call GetWindowThreadProcessId(lngVBHwnd, lngPid)
                                    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
                                    Call TerminateProcess(lngProcess, 1&)
                                    gobjTrace.WriteInfo "KillHisProcess", "杀掉进程", "VB6EXE"
                                    mintVB6 = 2
                                Else
                                    mintVB6 = 1
                                    Exit Do
                                End If
                            Else
                                mintVB6 = 1
                                Exit Do
                            End If
                        ElseIf mintVB6 = 2 Then
                            Call GetWindowThreadProcessId(lngVBHwnd, lngPid)
                            lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
                            Call TerminateProcess(lngProcess, 1&)
                            gobjTrace.WriteInfo "KillHisProcess", "杀掉进程", "VB6EXE"
                        Else
                            Exit Do
                        End If
                    Else
                        Call GetWindowThreadProcessId(lngHwnd, lngPid)
                        lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPid)
                        Call TerminateProcess(lngProcess, 1&)
                        gobjTrace.WriteInfo "KillHisProcess", "杀掉进程", "VB6EXE"
                    End If
                End If
            End If
        Loop
        '正在检查ZLHISCrust.exe的其他进程
        lblInfor.Caption = "正在检查残余进程：ZLHISCRUST.EXE"
        prgPross.Value = lngCurPro + lngCurIncPro * 0.4
        lngHwnd = FindExitsProcess("ZLHISCRUST.EXE", GetCurrentProcessId)
        If lngHwnd <> 0 Then
            gobjTrace.WriteInfo "KillHisProcess", "杀掉进程", "ZLHISCRUST.EXE(其他)"
            Call TerminateProcess(lngHwnd, 1&)
        End If
        lngTotal = UBound(garrKillProcess) + 1
        For i = LBound(garrKillProcess) To UBound(garrKillProcess)
            If garrKillProcess(i) <> "VB6.EXE" And garrKillProcess(i) <> "ZLHISCRUST.EXE" Then
                lblInfor.Caption = "正在检查残余进程：" & garrKillProcess(i)
                prgPross.Value = lngCurPro + lngCurIncPro * 0.4 + lngCurIncPro * 0.6 * (i + 1) / lngTotal
                lngHwnd = FindExitsProcess(garrKillProcess(i))
                If lngHwnd <> 0 Then
                    gobjTrace.WriteInfo "KillHisProcess", "杀掉进程", garrKillProcess(i)
                    Call TerminateProcess(lngHwnd, 1&)
                End If
            End If
        Next
    End If
    If ctType = CT_StartSvr Then
        grsFileUpgrade.Filter = "自动注册=" & RegFileType.RFT_NETServer
        lngTotal = grsFileUpgrade.RecordCount
        For i = 1 To lngTotal
            prgPross.Value = lngCurPro + lngCurIncPro * (i / lngTotal)
            If gobjFSO.FileExists(grsFileUpgrade!实际路径) Then
                gobjTrace.WriteInfo "ControlProcAndSvr", "启动服务", gobjFSO.GetBaseName(grsFileUpgrade!实际路径)
                lblInfor.Caption = "正在启动服务：" & gobjFSO.GetBaseName(grsFileUpgrade!实际路径)
                If objShell.Run("NET STOP " & gobjFSO.GetBaseName(grsFileUpgrade!实际路径), strReturn, strErr, 30000) Then
                End If
                gobjTrace.WriteInfo "ControlProcAndSvr", "启动结果", strReturn, "错误信息", strErr
            End If
            grsFileUpgrade.MoveNext
        Next
    End If
    prgPross.Value = lngCurPro + lngCurIncPro
    ControlProcAndSvr = True
    Exit Function
ErrH:
    gobjTrace.WriteInfo "ControlProcAndSvr", "服务进程控制出错", Err.Description
    Err.Clear
End Function

Private Function FileUpgrade() As Boolean
''功能：升级的主体逻辑代码
    Dim strErr          As String, lngRet   As Long
    Dim lngCurPro       As Long, i          As Long, lngTotal       As Long
    Dim blnOperateOK    As Boolean
    
    
    blnOperateOK = False
    On Error GoTo ErrH
    gobjTrace.WriteSection "文件升级", SL_LevelTwo
    '检查文件更新
    If Not CheckUpdate(prgPross.Value, IIf(gotCurType = OT_PreUpgrade, 10, 10)) Then GoTo ErrEnd
    If gotCurType <> OT_CheckFile Then
        '文件下载解压
        If Not DownAndDecFiles(prgPross.Value, IIf(gotCurType = OT_PreUpgrade, 60, 60)) Then GoTo ErrEnd
    End If
    
    prgPross.Value = 70
    '预升级不需要这些
    If gotCurType <> OT_PreUpgrade And gotCurType <> OT_CheckFile Then
        '杀掉进程
        Call ControlProcAndSvr(prgPross.Value, 3)
        '删除弃用文件
        If Not DeleteExpiredFile(prgPross.Value, 3) Then GoTo ErrEnd
        If gotCurType = OT_Repair Then
            grsFileUpgrade.Filter = "(更新<" & UC_Normal & " And 错误信息=NULL) Or (更新=" & UC_Normal & " And 清理文件路径<>NULL) OR  (更新=" & UC_Normal & " And 清理文件路径=NULL And 文件类型<>" & FT_System & ")"
        Else
            grsFileUpgrade.Filter = "(更新<" & UC_Normal & " And 错误信息=NULL) Or (更新=" & UC_Normal & " And 清理文件路径<>NULL)"
        End If
        If grsFileUpgrade.RecordCount <> 0 Then
            '杀掉进程
            Call ControlProcAndSvr(prgPross.Value, 3, CT_KillProcAndSvr)
            '文件安装注册
            If Not SetupFiles(prgPross.Value, 10) Then GoTo ErrEnd
            '启用服务
            Call ControlProcAndSvr(prgPross.Value, 3, CT_StartSvr)
            '批处理文件执行
            If Not ExecBatFile(prgPross.Value, 3) Then GoTo ErrEnd
        Else
            gobjTrace.WriteInfo "FileUpgrade", "安装清理检查", "不存在需要清理或者安装的文件"
        End If
    ElseIf gotCurType = OT_CheckFile Then
        If Not ReportCheckInfo(blnOperateOK) Then GoTo ErrEnd
    End If
    
    If gotCurType <> OT_CheckFile Then
        prgPross.Value = 95
        '杀掉7Z.exe进程
        lblInfor.Caption = "正在清理7Z.EXE进程"
        lngRet = FindExitsProcess("7Z.EXE")
        If lngRet <> 0 Then Call TerminatePID(lngRet)
        lblInfor.Caption = "正在关闭文件服务器连接"
        Call gclsConnect.CloseConnect
    End If
    
    If gotCurType <> OT_CheckFile And gotCurType <> OT_PreUpgrade Then
        If Not LoadDetailList(prgPross.Value, 3) Then GoTo ErrEnd
        blnOperateOK = lvwMan.ListItems.Count = 0
        prgPross.Value = 99
        lblInfor.Caption = "正在清理临时目录"
        Call ClearFolder(gstrSetupPath & "\ZLUPTMP", blnOperateOK)

        '再次启动导航台时会重新检测安装部件
        If gblnReCheckComs Then
            gobjTrace.WriteInfo "FileUpgrade", "重新检测安装部件", True
            SaveSetting "ZLSOFT", "注册信息", "本机部件", ""
        End If
    End If
    
    If Not blnOperateOK Then
        Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:" & Decode(gotCurType, OT_OfficialUpgrade, "升级失败", OT_Repair, "修复失败", OT_PreUpgrade, "预升级失败", OT_CheckFile, "检查完成,需要升级") & " 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        lblInfor.Caption = "正在标记本次升级状态"
        Call SetOperateProcess(gotCurType, IIf(gotCurType = OT_CheckFile, OS_Failure, OS_NotInProcessing), SumErrMsg, glngFileBatch)  '标识升级完成
    Else
        Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:" & Decode(gotCurType, OT_OfficialUpgrade, "升级成功", OT_Repair, "修复成功", OT_PreUpgrade, "预升级失败", OT_CheckFile, "检查完成,无需升级") & " 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
        lblInfor.Caption = "正在标记本次升级状态"
        Call SetOperateProcess(gotCurType, OS_Completed, SumErrMsg, glngFileBatch)  '标识升级完成
        
        If gotCurType = OT_OfficialUpgrade Then
            lblInfor.Caption = "正在检测安装OO4O组件"
            gobjTrace.WriteSection "OO4O安装", SL_LevelThree
            If Not InstallOO4O(strErr) Then
                gobjTrace.WriteInfo "InstallOO4O", "安装安装OO4O组件", "失败：" & strErr
            Else
                gobjTrace.WriteInfo "InstallOO4O", "安装安装OO4O组件", "成功：" & strErr
            End If
        End If
    End If
    prgPross.Value = 100
    lblInfor.Caption = "升级结束"
    cmdOK.Visible = True
    mblnOperateCompleted = True
    If gotCurType <> OT_CheckFile And gotCurType <> OT_PreUpgrade Then
        FileUpgrade = blnOperateOK
    Else
        FileUpgrade = True
    End If
    Exit Function
ErrH:
    prgPross.Value = 100
    gobjTrace.WriteInfo "FileUpgrade", "升级过程发生致命错误", Err.Description
    If Not gblnHelperMain Then MsgBox "升级过程发生致命错误，请联系管理员！信息：" & Err.Description, vbInformation, App.Title
    Err.Clear
ErrEnd:
    Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:升级失败 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
    Call SetOperateProcess(gotCurType, IIf(gotCurType = OT_CheckFile, OS_Failure, OS_NotInProcessing), SumErrMsg, glngFileBatch)       '标识升级结束
    FileUpgrade = False
    cmdOK.Visible = True
    mblnOperateCompleted = True
End Function

Private Function LoadDetailList(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
    Dim strFile     As String, strErr       As String
    Dim lstTmp      As ListItem, intKey     As Integer
    Dim intIndex    As Integer
    Dim lngTotal    As Long, i              As Long
    
    On Error GoTo ErrH
    gobjTrace.WriteSection "升级概要", SL_LevelTwo
    grsFileUpgrade.Filter = ""
    grsFileUpgrade.Sort = "更新,文件名"
    lngTotal = grsFileUpgrade.RecordCount
    lblInfor.Caption = "正在加载错误清单"
    lvwMan.ListItems.Clear
    With grsFileUpgrade
        Do While Not .EOF
            i = i + 1
            prgPross.Value = lngCurPro + lngCurIncPro * 0.2 * (i / lngTotal)
            If !更新 <= UC_Normal Then
                If !错误信息 & "" <> "" Then
                    intKey = intKey + 1
                    Set lstTmp = lvwMan.ListItems.Add(, "K" & intKey, !文件名, "List", "List")
                    lstTmp.SubItems(ELC_错误信息) = !错误信息 & ""
                    lstTmp.SmallIcon = "Err"
                    gobjTrace.WriteInfo "LoadDetailList", !文件名, "失败", "信息", !错误信息
                ElseIf !更新 < UC_Normal Then
                    gobjTrace.WriteInfo "LoadDetailList", !文件名, "成功"
                Else
                    gobjTrace.WriteInfo "LoadDetailList", !文件名, "无需更新", "信息", !检查信息
                End If
            ElseIf !更新 > UC_SvrMD5Null Then
                gobjTrace.WriteInfo "LoadDetailList", !文件名, "成功但是存在警告", "警告", !检查信息
            Else
                gobjTrace.WriteInfo "LoadDetailList", !文件名, "无需更新", "信息", !检查信息
            End If
            .MoveNext
        Loop
    End With
    If lvwMan.ListItems.Count <> 0 Then
        Me.Height = 5445
        Me.Refresh
    End If
    prgPross.Value = lngCurPro + lngCurIncPro
    LoadDetailList = True
    Exit Function
ErrH:
    strErr = Err.Description
    gobjTrace.WriteInfo "CheckUpdate", "加载错误清单致命错误", strErr
    If Not gblnHelperMain Then MsgBox "加载错误清单致命错误，请联系管理员！信息：" & strErr, vbInformation, App.Title
    Call RecordErrMsg(MT_ChcekUpdate, "加载错误清单致命错误", strErr)
    Err.Clear
End Function

Private Function CheckUpdate(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
'功能：检查并获取更新文件
'参数：lngCurPro=当前进度
'      lngCurIncPro=当前过程执行完毕后的增量进度
'返回：若发生不可预知错误则为False,否则为True
    Dim rsFileList   As ADODB.Recordset, ucUpdate       As UpdateCheck, intPreDown  As Integer
    Dim lngRecCount     As Long, i                      As Long, lngBeach           As Integer
    Dim strFile         As String, strWrongFile         As String, strAddSetFile    As String
    Dim arrComs         As Variant
    Dim rsTmp           As ADODB.Recordset
    Dim strlocVersion   As String, strLocModiTime       As String, strLocMd5        As String
    Dim lngTotal        As Long, lngLoop                As Long
    Dim strTmpErr       As String, lngSort               As Long, strNoSubfix       As String
    Dim strOldFile      As String
    
    On Error GoTo ErrH
    Set mcllOldComs = New Collection
    gobjTrace.WriteSection "更新检查(0)", SL_LevelThree
    lblInfor.Caption = "正在进行文件更新检测..."
    '读取预升级下载的升级配置文件
    If gobjFSO.FileExists(gstrPreTempPath & "\ZLList.adtg") And (gotCurType = OT_OfficialUpgrade Or gotCurType = OT_Repair) Then
        Set rsFileList = New ADODB.Recordset
        rsFileList.Open gstrPreTempPath & "\ZLList.adtg", , adOpenStatic, adLockOptimistic, adCmdFile
        On Error Resume Next
        rsFileList.Sort = "文件名"
        If Err.Number = 0 Then
            lngRecCount = rsFileList.RecordCount
            gobjTrace.WriteInfo "CheckUpdate", "预升级清单记录", lngRecCount
        Else
            gobjTrace.WriteInfo "CheckUpdate", "预升级清单记录", Err.Description
            Err.Clear
        End If
        On Error GoTo ErrH
    End If
    
    With grsFileUpgrade
        '先进行路径转换，并判断本地文件是否存在,并进行简单的升级判断
        lngTotal = grsFileUpgrade.RecordCount
        For lngLoop = 1 To lngTotal
            gobjTrace.WriteSection "-", SL_LevelThree
            lblInfor.Caption = "正在检测文件：" & !文件名
            prgPross.Value = lngCurPro + lngCurIncPro * 0.75 * (lngLoop / lngTotal)
            strFile = GetActualPath(!安装路径 & "", Val(!文件类型 & ""), !文件名)
            lngSort = Decode(Val(!文件类型 & ""), FT_Apply, 5, FT_Public, 4, FT_System, 3, FT_AdditionFile, 0, FT_Help, 0, FT_Other, 2, 1)
            strNoSubfix = !标准文件名
            If InStr(strNoSubfix, ".") > 0 Then
                strNoSubfix = Mid(strNoSubfix, 1, InStrRev(strNoSubfix, ".") - 1)
            End If
            gobjTrace.WriteInfo "CheckUpdate", "部件", !文件名, "安装路径", !安装路径, _
                            "自动注册", Decode(!自动注册, RFT_NotReg, "不注册", RFT_NormalReg, "自动识别注册", RFT_NETGAC, "NET全局程序集缓存注册", RFT_NETServer, "NET系统服务器注册", RFT_NETComReg, "NETCOM注册"), _
                            "文件类型", Decode(!文件类型, FT_Apply, "业务部件", FT_Public, "公共部件", FT_System, "系统文件", FT_AdditionFile, "附件文件", FT_Help, "帮助文件", FT_Other, "三方文件", "未识别的文件"), _
                            "强制覆盖", !强制覆盖, "业务部件", !业务部件, "MD5", !MD5, "修改日期", !修改日期, "文件版本", !版本号, "附加安装路径", !附加安装路径
            intPreDown = 0: ucUpdate = UC_NotExists: lngBeach = 0: strTmpErr = "": strWrongFile = "": strAddSetFile = ""
            If InStr(",ZLMIPCLIENTSHELL.EXE,ZLIDKIND.OCX,ZLIDCARD.DLL,ZLICCARD.DLL,ZLREGISTER.DLL,ZL9COMLIB.DLL,ZLLOGIN.DLL,ZLHIS+.EXE,", "," & !标准文件名 & ",") > 0 Then
                lngBeach = -1
            ElseIf !自动注册 = RFT_NotReg Then
                lngBeach = -2
            End If
            If lngRecCount <> 0 Then
                rsFileList.Filter = "标准文件名='" & !标准文件名 & "'"
                If Not rsFileList.EOF Then
                    If !MD5 = rsFileList!MD5 Then
                        If gobjFSO.FileExists(gstrPreTempPath & "\" & rsFileList!文件名 & ".7z") Then
                            ucUpdate = rsFileList!更新 '从预升级获取文件
                            intPreDown = 1
                            lngBeach = rsFileList!判断批次
                            strTmpErr = rsFileList!检查信息 & ""
                        End If
                    End If
                    Call rsFileList.Delete
                    Call rsFileList.UpdateBatch
                    lngRecCount = lngRecCount - 1
                End If
            End If
            '判断能否升级
            If ucUpdate = UC_NotExists Then
                If Not IsNull(!MD5) Then
                    If gobjFSO.FileExists(strFile) Then
                        gobjTrace.WriteInfo "CheckUpdate", "存在本地文件", strFile
                        If !文件类型 = FT_System And !强制覆盖 = 0 Then
                            ucUpdate = UC_Normal
                            strTmpErr = "该文件是系统文件，本地存在且不需要强制覆盖,无需升级"
                        Else
                            strLocMd5 = FileMD5(strFile)
                            If !MD5 = strLocMd5 Then
                                ucUpdate = UC_Normal
                                strTmpErr = "本地和服务器MD5相同,无需升级"
                            Else
                                ucUpdate = UC_Update
                                strTmpErr = "本地和服务器MD5不相同,需要升级"
                            End If
                        End If
                    ElseIf !文件类型 = FT_Apply Then
                        If IsNull(!业务部件) Then '应用部件的业务部件为空不下载
                            ucUpdate = UC_NotExists
                            strTmpErr = "该文件是应用部件,本地不存在但业务部件为空,无需下载"
                        ElseIf UCase(!业务部件) = !标准文件名 Then '应用部件的业务部件是自身则强制下载
                            ucUpdate = UC_NewDown
                            strTmpErr = "该文件是应用部件,本地不存在但业务部件等于自身,需要下载"
                        End If
                    Else '普通文件不存在就下载
                        ucUpdate = UC_NewDown
                        strTmpErr = "该文件是非应用部件,本地不存,需要下载"
                    End If
                Else '服务器MD5为空不进行升级
                    ucUpdate = UC_SvrMD5Null '表示服务器MD5为空
                    strTmpErr = "ZLFilesUpgrade中没有该文件的MD5信息，无法进行部件升级检查"
                End If
            End If
            '获取错误路径文件
            '清理文件路径不能包含附加安装路径中的路径,此处根据是否注册，使两者分开。
            If !自动注册 > RFT_NotReg Then '需要注册的文件才清理错误路径
                strWrongFile = GetWrongFiles(!文件名, strFile)
                If strWrongFile <> "" Then '存在错误路径文件，但是文件不下载，则自动标记为下载更新
                    If ucUpdate = UC_NotExists Then
                        ucUpdate = UC_NewDown
                        strTmpErr = strTmpErr & IIf(strTmpErr <> "", ";", "") & "文件路径错误(错误路径：" & strWrongFile & "),因此需要重新下载"
                    ElseIf ucUpdate <> UC_SvrMD5Null Then
                        strTmpErr = strTmpErr & IIf(strTmpErr <> "", ";", "") & "文件路径错误(错误路径：" & strWrongFile & "),因此需要重新注册"
                    End If
                End If
            Else
                strAddSetFile = GetAdditionSetup(!文件名, !MD5 & "", !附加安装路径 & "")
                If strAddSetFile <> "" Then '存在附加安装路径，但是文件不下载，则自动标记为下载更新
                    If ucUpdate = UC_NotExists Then
                        ucUpdate = UC_NewDown
                        strTmpErr = strTmpErr & IIf(strTmpErr <> "", ";", "") & "需要附加安装文件(附加安装路径：" & strAddSetFile & "),因此需要重新下载"
                    ElseIf ucUpdate <> UC_SvrMD5Null Then
                        If ucUpdate = UC_Normal Then ucUpdate = UC_AddtionUp
                        strTmpErr = strTmpErr & IIf(strTmpErr <> "", ";", "") & "需要附加安装文件(附加安装路径：" & strAddSetFile & "),因此需要重新安装附加文件"
                    Else
                        strTmpErr = strTmpErr & IIf(strTmpErr <> "", ";", "") & "需要附加安装文件(附加安装路径：" & strAddSetFile & "),但是服务器文件MD5为空，无法下载"
                    End If
                End If
            End If
            grsFileUpgrade.Update Array("更新", "实际路径", "清理文件路径", "附加实际路径", "预升级下载", "判断批次", "检查信息", "无后缀文件名", "类型排序"), _
                                Array(ucUpdate, strFile, IIf(strWrongFile = "", Null, strWrongFile), IIf(strAddSetFile = "", Null, strAddSetFile), intPreDown, lngBeach, IIf(strTmpErr = "", Null, strTmpErr), strNoSubfix, lngSort)
            gobjTrace.WriteInfo "CheckUpdate", "更新", ucUpdate, "预升级下载", intPreDown, "判断批次", lngBeach, "实际路径", strFile, "清理文件路径", strWrongFile, "附加实际路径", strAddSetFile, "无后缀文件名", strNoSubfix, "类型排序", lngSort
            gobjTrace.WriteInfo "CheckUpdate", "检查说明(0)", strTmpErr
            grsFileUpgrade.MoveNext
        Next
        '轮询遍历，每次对没有判定更新且本地不存在的存在业务部件设置的应用部件进行判断
        '判断的参考依靠每次判断变动的文件。
        grsFileUpgrade.Filter = ""
        lngRecCount = 0: lngBeach = 0
        Set rsTmp = CopyNewRec(grsFileUpgrade)
        grsFileUpgrade.Filter = "更新=" & UC_NotExists & " And 业务部件<>NULL And 文件类型=" & FT_Apply
        Do While lngRecCount <> grsFileUpgrade.RecordCount
            lngRecCount = grsFileUpgrade.RecordCount
            lngBeach = lngBeach + 1
            If lngRecCount > 0 Then gobjTrace.WriteSection "更新检查(" & lngBeach & ")", SL_LevelThree
            For lngLoop = 1 To lngRecCount
                gobjTrace.WriteSection "-", SL_LevelThree
                lblInfor.Caption = "正在检测文件：" & !文件名
                prgPross.Value = lngCurPro + lngCurIncPro * 0.75 + lngCurIncPro * 0.25 * IIf(lngBeach > 3, 1, (lngBeach / 3) * (lngLoop / lngRecCount))
                '业务部件判断
                ucUpdate = UC_NotExists: strTmpErr = !检查信息 & ""
                arrComs = Split(UCase(grsFileUpgrade!业务部件), ",")
                For i = LBound(arrComs) To UBound(arrComs)
                    If arrComs(i) Like "*.*" Then '业务部件带后缀
                        rsTmp.Filter = "标准文件名='" & arrComs(i) & "'"
                    Else
                        rsTmp.Filter = "无后缀文件名 = '" & arrComs(i) & "'"
                    End If
                    If Not rsTmp.EOF Then
                        If rsTmp!更新 < UC_NotExists Then     '需要更新或这本地存在且不需要更新
                            ucUpdate = UC_NewDown
                            If rsTmp!更新 = UC_Update Then
                                strTmpErr = "业务部件""" & rsTmp!文件名 & """本地存在且需要更新，因此该部件需要下载"
                            ElseIf rsTmp!更新 = UC_NewDown Then
                                strTmpErr = "业务部件""" & rsTmp!文件名 & """已经标记为下载，因此该部件需要下载"
                            Else 'UC_normal,本地存在且不需要更新
                                strTmpErr = "业务部件""" & rsTmp!文件名 & """本地存在但是不需要更新，因此该部件需要下载"
                            End If
                        ElseIf rsTmp!更新 = UC_SvrMD5Null And gobjFSO.FileExists(rsTmp!实际路径) Then
                            ucUpdate = UC_NewDown
                            strTmpErr = "业务部件""" & rsTmp!文件名 & """本地存在尽管业务部件服务器MD5为空，因此该部件需要下载"
                        End If
                    ElseIf IsOldComponentExists(arrComs(i), strOldFile) Then
                        ucUpdate = UC_NewDown
                        strTmpErr = "业务部件""" & strOldFile & """本地存在（可能是已经不再使用的业务部件，部件清单中不存在该业务部件），因此该部件需要下载"
                    End If
                    If ucUpdate = UC_NewDown Then Exit For
                Next
                If ucUpdate = UC_NewDown Then
                    grsFileUpgrade.Update Array("更新", "判断批次", "检查信息"), Array(ucUpdate, lngBeach, IIf(strTmpErr = "", Null, strTmpErr))
                    gobjTrace.WriteInfo "CheckUpdate", "部件", !文件名, "安装路径", !安装路径, _
                                    "自动注册", Decode(!自动注册, RFT_NotReg, "不注册", RFT_NormalReg, "自动识别注册", RFT_NETGAC, "NET全局程序集缓存注册", RFT_NETServer, "NET系统服务器注册", RFT_NETComReg, "NETCOM注册"), _
                                    "文件类型", Decode(!文件类型, FT_Apply, "业务部件", FT_Public, "公共部件", FT_System, "系统文件", FT_AdditionFile, "附件文件", FT_Help, "帮助文件", FT_Other, "三方文件", "未识别的文件"), _
                                    "强制覆盖", !强制覆盖, "业务部件", !业务部件, "MD5", !MD5, "修改日期", !修改日期, "文件版本", !版本号
                    gobjTrace.WriteInfo "CheckUpdate", "更新", ucUpdate, "判断批次", lngBeach
                    gobjTrace.WriteInfo "CheckUpdate", "检查说明(" & lngBeach & ")", strTmpErr
                Else
                    gobjTrace.WriteInfo "CheckUpdate", "部件", !文件名, "检查说明(" & lngBeach & ")", "待定"
                End If
                grsFileUpgrade.MoveNext
            Next
            grsFileUpgrade.Filter = "判断批次=" & lngBeach '获取本轮判断需要下载的部件
            If grsFileUpgrade.EOF Then Exit Do
            Set rsTmp = CopyNewRec(grsFileUpgrade)
            grsFileUpgrade.Filter = "更新=" & UC_NotExists & " And 业务部件<>NULL And 文件类型=" & FT_Apply
        Loop
        '上次注册检查
        '读取预升级下载的升级配置文件
        If gobjFSO.FileExists(gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg") And (gotCurType = OT_OfficialUpgrade Or gotCurType = OT_Repair) Then
            Set rsFileList = New ADODB.Recordset
            rsFileList.Open gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg", , adOpenStatic, adLockOptimistic, adCmdFile
            On Error Resume Next
            rsFileList.Sort = "文件名"
            If Err.Number = 0 Then '因为流无效，行集不能被加载。
                lngRecCount = rsFileList.RecordCount
                gobjTrace.WriteInfo "CheckUpdate", "上次注册失败记录", lngRecCount
                lngBeach = lngBeach + 1
                gobjTrace.WriteSection "更新检查(" & lngBeach & ")", SL_LevelThree
                Do While Not rsFileList.EOF
                    strTmpErr = ""
                    .Filter = "标准文件名='" & rsFileList!标准文件名 & "'"
                    If Not .EOF Then
                        '新下载或者更新或者错误文件不为空的都是需要重新注册的。
                        If (!更新 = UC_Normal Or !更新 = UC_SvrMD5Null And gobjFSO.FileExists(!实际路径)) And !自动注册 <> RFT_NotReg Then
                            strTmpErr = !检查信息 & ""
                            strTmpErr = strTmpErr & IIf(strTmpErr = "", "", ";") & "上次升级注册不成功，需要重新注册"
                            .Update Array("更新", "检查信息"), Array(UC_RegAgain, strTmpErr)
                            gobjTrace.WriteInfo "CheckUpdate", "部件", !文件名, "安装路径", !安装路径, _
                                            "自动注册", Decode(!自动注册, RFT_NotReg, "不注册", RFT_NormalReg, "自动识别注册", RFT_NETGAC, "NET全局程序集缓存注册", RFT_NETServer, "NET系统服务器注册", RFT_NETComReg, "NETCOM注册"), _
                                            "文件类型", Decode(!文件类型, FT_Apply, "业务部件", FT_Public, "公共部件", FT_System, "系统文件", FT_AdditionFile, "附件文件", FT_Help, "帮助文件", FT_Other, "三方文件", "未识别的文件"), _
                                            "强制覆盖", !强制覆盖, "业务部件", !业务部件, "MD5", !MD5, "修改日期", !修改日期, "文件版本", !版本号
                            gobjTrace.WriteInfo "CheckUpdate", "更新", ucUpdate
                            gobjTrace.WriteInfo "CheckUpdate", "检查说明(" & lngBeach & ")", strTmpErr
                        End If
                    End If
                    rsFileList.MoveNext
                Loop
            Else
                gobjTrace.WriteInfo "CheckUpdate", "上次注册失败记录", Err.Description
                Err.Clear
            End If
            On Error GoTo ErrH
        End If
    End With
    '将未判定的部件调整为无需升级
    Call UpdateRec(grsFileUpgrade, "更新=" & UC_NotExists, "检查信息", "本地不存在且不需要更新")
    Call UpdateRec(grsFileUpgrade, "自动注册=" & RFT_NETServer, "判断批次", lngBeach + 1)
    prgPross.Value = lngCurPro + lngCurIncPro
    CheckUpdate = True
    Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    strTmpErr = Err.Description
    gobjTrace.WriteInfo "CheckUpdate", "检查文件更新发生致命错误", strTmpErr
    If Not gblnSilence And Not gblnHelperMain Then MsgBox "检查文件更新发生致命错误，请联系管理员！信息：" & strTmpErr, vbInformation, App.Title
    Call RecordErrMsg(MT_ChcekUpdate, "检查文件更新发生致命错误", strTmpErr)
    Err.Clear
End Function

Private Function DownAndDecFiles(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
'功能：下载并解压文件
'参数：lngCurPro=当前进度
'      lngCurIncPro=当前过程执行完毕后的增量进度
'返回：若发生不可预知错误则为False,否则为True
    Dim lngTotal        As Long, lngLoop                As Long
    Dim strTmpPath      As String, strErrTmp            As String
    Dim strlocVersion   As String, strLocModiTime       As String, strLocMd5    As String
    Dim strErrInfo      As String
    Dim rsTmp           As ADODB.Recordset
    Dim blnZip          As Boolean
    
    On Error GoTo ErrH
    lblInfor.Caption = "正在下载文件..."
    gobjTrace.WriteSection "下载解压", SL_LevelThree
    With grsFileUpgrade
        .Filter = "更新< " & UC_RegAgain
        lngTotal = .RecordCount
        gblnReCheckComs = lngTotal > 0
        For lngLoop = 1 To lngTotal
            gobjTrace.WriteSection "-", SL_LevelThree
            lblInfor.Caption = IIf(gotCurType <> OT_PreUpgrade, IIf(!预升级下载 = 0, "正在下载并解压文件:", "正在解压文件:"), "正在下载文件:") & !文件名
            prgPross.Value = lngCurPro + lngCurIncPro * 0.9 * (lngLoop / lngTotal)
            strErrInfo = ""
            If gotCurType = OT_PreUpgrade Or !预升级下载 = 1 Then
                strTmpPath = gstrPreTempPath
            Else
                strTmpPath = gstrTempPath
            End If
            '已经打包的文件，不再添加后缀以及解压（在安装后解压），在整体安装完成之后再进行单独的安装处理
            blnZip = !标准文件名 Like "*.7Z"
            If !预升级下载 = 0 Then
                If gclsConnect.IsServerFileExists(!标准文件名 & IIf(blnZip, "", ".7z")) Then
                    DoEvents
                    If gclsConnect.DownloadFile(!标准文件名 & IIf(blnZip, "", ".7z"), strTmpPath, strErrTmp) Then
                    Else
                        strErrInfo = "文件下载失败，" & strErrTmp & "(服务器：" & gclsConnect.ServerPath & ")"
                    End If
                Else
                    strErrInfo = "服务器文件不存在(服务器：" & gclsConnect.ServerPath & ")"
                End If
            End If
            
            If gotCurType <> OT_PreUpgrade And strErrInfo = "" Then
                DoEvents
                If Not blnZip Then
                    If Not gobj7zZip.UnZipFile(strTmpPath & "\" & !文件名 & ".7z", strTmpPath & "\" & !文件名, , strErrTmp) Then
                        If strErrTmp = "" Then
                            strErrInfo = "解压后文件" & strTmpPath & "\" & !文件名 & "不存在,可能被杀毒软件杀掉"
                        Else
                            strErrInfo = "文件解压失败，" & strErrTmp
                        End If
                    End If
                End If
                
                If strErrInfo = "" Then
                    If gobjFSO.FileExists(strTmpPath & "\" & !文件名) Then
                        strLocMd5 = FileMD5(strTmpPath & "\" & !文件名)
                        If !MD5 <> strLocMd5 Then
                            If gblnMD5Check Then
                                strErrInfo = "服务器文件受损(服务器文件和收集的MD5不匹配)(服务器：" & gclsConnect.ServerPath & ")"
                            Else
                                Call RecordErrMsg(MT_DownAndDec, !文件名, "服务器文件受损(服务器文件和收集的MD5不匹配)(服务器：" & gclsConnect.ServerPath & ")")
                                gobjTrace.WriteInfo "DownAndDecFiles", "服务器文件受损(服务器文件和收集的MD5不匹配)(服务器：" & gclsConnect.ServerPath & ")"
                            End If
                        End If
                    End If
                End If
            End If
            If strErrInfo <> "" Then
                grsFileUpgrade.Update "错误信息", strErrInfo
                Call RecordErrMsg(MT_DownAndDec, !文件名, strErrInfo)
                gobjTrace.WriteInfo "DownAndDecFiles", IIf(gotCurType <> OT_PreUpgrade, IIf(!预升级下载 = 0, "下载并解压文件", "解压文件"), "下载文件"), !文件名, "失败信息", strErrInfo
            Else
                gobjTrace.WriteInfo "DownAndDecFiles", IIf(gotCurType <> OT_PreUpgrade, IIf(!预升级下载 = 0, "下载并解压文件", "解压文件"), "下载文件"), !文件名
            End If
            grsFileUpgrade.MoveNext
        Next
        '保存预升级文件清单
        If gotCurType = OT_PreUpgrade Then
            lblInfor.Caption = "正在保存预升级文件清单"
            .Filter = "更新<" & UC_RegAgain
            If .RecordCount > 0 Then
                Set rsTmp = CopyNewRec(grsFileUpgrade)
                rsTmp.Sort = "文件名"
                If gobjFSO.FileExists(gstrPreTempPath & "\ZLList.adtg") Then
                    gobjFSO.DeleteFile gstrPreTempPath & "\ZLList.adtg", True
                End If
                rsTmp.Save gstrPreTempPath & "\ZLList.adtg", adPersistADTG
                rsTmp.Close
            End If
        End If
        prgPross.Value = lngCurPro + lngCurIncPro
    End With
    DownAndDecFiles = True
    Exit Function
ErrH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "DownAndDecFiles", "下载解压文件发生致命错误", strErrInfo
    Call RecordErrMsg(MT_DownAndDec, "下载解压文件发生致命错误", strErrInfo)
    If Not gblnHelperMain Then MsgBox "下载解压文件发生致命错误，请联系管理员！信息：" & strErrInfo, vbInformation, App.Title
    Err.Clear
End Function

Private Function DeleteExpiredFile(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
'功能：删除弃用文件
    Dim strSQL      As String, rsTmp        As ADODB.Recordset
    Dim rsSys       As ADODB.Recordset
    Dim strFile     As String
    Dim i           As Integer, lngCount    As Long, strErr     As String

    On Error Resume Next
    gobjTrace.WriteSection "清理弃用文件", SL_LevelThree
    strSQL = "Select 文件名,Upper(文件名) 标准文件名,安装路径,系统编号,系统版本 From zlFilesExpired"
    Set rsTmp = OpenSQLRecord(strSQL, "舍弃文件")
    lblInfor.Caption = "正在检测弃用文件..."
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrH
    If Not rsTmp Is Nothing Then '可能该表不存在
        If Not rsTmp.EOF Then
            strSQL = "Select 内容 版本号, 0 编号 From Zlreginfo Where 项目 = '版本号' Union All Select 版本号, 编号 From Zlsystems"
            Set rsSys = OpenSQLRecord(strSQL, "获取系统版本")
        End If
        lngCount = rsTmp.RecordCount
        For i = 1 To lngCount
            prgPross.Value = lngCurPro + lngCurIncPro * (i / lngCount)
            lblInfor.Caption = "正在检测弃用文件：" & rsTmp!文件名
            rsSys.Filter = "编号=" & Val(rsTmp!系统编号 & "")
            gobjTrace.WriteInfo "DeleteExpiredFile", "检查文件", rsTmp!文件名, "检查版本", rsTmp!系统版本, "检查路径", rsTmp!安装路径
            If Not rsSys.EOF Then
                '文件启用版本小于当前系统版本就可以弃用了
                If VerFull(rsTmp!系统版本) <= VerFull(rsSys!版本号) Then
                    On Error Resume Next
                    strFile = gcllSetPath("K_" & UCase(rsTmp!安装路径)) & "\" & rsTmp!文件名
                    If Err.Number <> 0 Then
                        gobjTrace.WriteInfo "DeleteExpiredFile", "无法清理", "文件路径无法转换：" & rsTmp!安装路径
                        Err.Clear
                        On Error GoTo ErrH
                    Else
                        On Error GoTo ErrH
                        '只处理不在升级文件清单中的文件，因为
                        If gobjFSO.FileExists(strFile) Then
                            grsFileUpgrade.Filter = "标准文件名='" & rsTmp!标准文件名 & "'"
                            If grsFileUpgrade.EOF Then
                                gobjTrace.WriteInfo "DeleteExpiredFile", "清理文件", strFile
                                On Error Resume Next
                                If FileSystem.GetAttr(strFile) <> vbNormal Then
                                    FileSystem.SetAttr strFile, vbNormal
                                End If
                                Call gclsRegCom.UnRegCom(strFile)
                                Call gobjFSO.DeleteFile(strFile, True)
                                If Err.Number <> 0 Then Err.Clear
                            Else
                                gobjTrace.WriteInfo "DeleteExpiredFile", "无需清理", "文件已经存在升迁文件列表中"
                            End If
                        Else
                            gobjTrace.WriteInfo "DeleteExpiredFile", "无需清理", "弃用文件不存在"
                        End If
                    End If
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    prgPross.Value = lngCurPro + lngCurIncPro
    DeleteExpiredFile = True
    Exit Function
ErrH:
    strErr = Err.Description
    gobjTrace.WriteInfo "DeleteExpiredFile", "清理弃用文件发生致命错误", strErr
    Call RecordErrMsg(MT_SetUp, "清理弃用文件发生致命错误", strErr)
    If Not gblnHelperMain Then MsgBox "清理弃用文件发生致命错误，请联系管理员！信息：" & strErr, vbInformation, App.Title
    Err.Clear
End Function

Private Function SetupFiles(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
    Dim arrComs     As Variant, i       As Integer
    Dim lngLoop     As Long, lngTotal   As Long
    Dim strErrInfo  As String, blnCanUp As Boolean, strErrTmp As String
    Dim blnRegErr   As Boolean
    Dim rsTmp           As ADODB.Recordset
    
    gobjTrace.WriteSection "安装注册文件", SL_LevelThree
    On Error GoTo ErrH
    With grsFileUpgrade
        If gotCurType = OT_Repair Then
            grsFileUpgrade.Filter = "(更新<" & UC_Normal & " And 错误信息=NULL) Or (更新=" & UC_Normal & " And 清理文件路径<>NULL) OR  (更新=" & UC_Normal & " And 清理文件路径=NULL And 文件类型<>" & FT_System & ")"
        Else
            grsFileUpgrade.Filter = "(更新<" & UC_Normal & " And 错误信息=NULL) Or (更新=" & UC_Normal & " And 清理文件路径<>NULL)"
        End If
        .Sort = "判断批次,自动注册,类型排序"
        lngTotal = .RecordCount
        For lngLoop = 1 To .RecordCount
            gobjTrace.WriteSection "-", SL_LevelThree
            prgPross.Value = lngCurPro + lngCurIncPro * (lngLoop / lngTotal)
            If !更新 > 0 Then
                gobjTrace.WriteInfo "SetupFiles", "安装注册文件", !文件名
            Else
                gobjTrace.WriteInfo "SetupFiles", "清理错误文件", !文件名
            End If
            strErrInfo = "": blnCanUp = True: blnRegErr = False
            If Not IsNull(!清理文件路径) Then
                On Error Resume Next
                lblInfor.Caption = "正在清理文件:" & !文件名
                arrComs = Split(!清理文件路径, "|")
                For i = LBound(arrComs) To UBound(arrComs)
                    '文件存在,则改属性
                    If FileSystem.GetAttr(arrComs(i)) <> vbNormal Then
                        FileSystem.SetAttr arrComs(i), vbNormal
                    End If
                    Call gclsRegCom.UnRegCom(arrComs(i), , !自动注册)
                    Call gobjFSO.DeleteFile(arrComs(i), True)
                    If Err.Number <> 0 Then Err.Clear
                Next
                On Error GoTo ErrH
            End If
            If !更新 < UC_RegAgain Then
                lblInfor.Caption = "正在安装文件:" & !文件名
                If SetupOneFile(!标准文件名, IIf(!预升级下载 = 1, gstrPreTempPath, gstrTempPath) & "\" & !文件名, !实际路径, !强制覆盖 = 1, strErrInfo) Then
                    '可能被占用，且忽略
                    If strErrInfo <> "" Then
                    ElseIf gobjFSO.FileExists(!实际路径) Then
                        gobjTrace.WriteInfo "SetupFiles", "成功安装文件", !文件名
                        '若是7z压缩文件，则自动解压，并不删除压缩文件，留作下一次判断。
                        If !标准文件名 Like "*.7Z" Then
                            If Not gobj7zZip.UnZipFile(!实际路径, Mid(!实际路径, 1, Len(!实际路径) - 3), False, strErrTmp, True) Then
                                gobjTrace.WriteInfo "SetupFiles", "压缩包解压失败", !文件名 & ":" & strErrTmp
                            Else
                                gobjTrace.WriteInfo "SetupFiles", "压缩包解压成功", !文件名
                            End If
                        End If
                    Else
                        blnCanUp = False
                        strErrInfo = "文件" & !实际路径 & "安装后不存在,可能被杀毒软件杀掉"
                        gobjTrace.WriteInfo "SetupFiles", "安装文件失败", "文件" & !实际路径 & "安装后不存在,可能被杀毒软件杀掉"
                    End If
                Else
                    blnCanUp = False
                    gobjTrace.WriteInfo "SetupFiles", "安装文件失败", strErrInfo
                End If
            End If
            If blnCanUp And strErrInfo = "" And Not IsNull(!附加实际路径) Then
                lblInfor.Caption = "正在进行附加安装:" & !文件名
                arrComs = Split(!附加实际路径, "|")
                For i = LBound(arrComs) To UBound(arrComs)
                    If SetupOneFile(!标准文件名, !实际路径, arrComs(i), !强制覆盖 = 1, strErrTmp) Then
                        If strErrTmp <> "" Then
                        ElseIf gobjFSO.FileExists(arrComs(i)) Then
                            gobjTrace.WriteInfo "SetupFiles", "成功安装附加安装文件", arrComs(i)
                        Else
                            blnCanUp = False
                            strErrTmp = "附加安装文件" & arrComs(i) & "安装后不存在,可能被杀毒软件杀掉"
                            gobjTrace.WriteInfo "SetupFiles", "安装附加安装文件失败", "文件" & arrComs(i) & "安装后不存在,可能被杀毒软件杀掉"
                        End If
                    Else
                        blnCanUp = False
                        gobjTrace.WriteInfo "SetupFiles", "安装附加安装文件失败", strErrTmp
                    End If
                    If strErrTmp <> "" Then
                        strErrInfo = strErrInfo & ";" & strErrInfo
                    End If
                Next
            End If
            If strErrInfo = "" And NVL(!自动注册, 0) <> 0 Then
                lblInfor.Caption = "正在注册文件:" & !文件名
                If Not gclsRegCom.RegCom(!实际路径, strErrInfo, !自动注册) Then
                    blnCanUp = False: blnRegErr = True
                    strErrInfo = "注册失败(" & strErrInfo & ")"
                    gobjTrace.WriteInfo "SetupFiles", "注册文件失败", strErrInfo
                Else
                    If strErrInfo <> "" Then '只是警告信息
                        gobjTrace.WriteInfo "SetupFiles", "注册文件失败", strErrInfo
                    Else
                        gobjTrace.WriteInfo "SetupFiles", "成功注册文件", !文件名
                    End If
                End If
            End If
            
            If strErrInfo <> "" Then
                If blnCanUp Then
                    grsFileUpgrade.Update Array("更新", "检查信息"), Array(UC_IgnorUp, strErrInfo)
                Else
                    grsFileUpgrade.Update Array("错误信息", "注册错误"), Array(strErrInfo, IIf(blnRegErr, 1, 0))
                End If
                Call RecordErrMsg(MT_SetUp, !文件名, strErrInfo)
            End If
            .MoveNext
        Next
        lblInfor.Caption = "正在保存注册失败文件清单"
        .Filter = "注册错误=1"
        If .RecordCount > 0 Then
            Set rsTmp = CopyNewRec(grsFileUpgrade)
            rsTmp.Sort = "文件名"
            If gobjFSO.FileExists(gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg") Then
                gobjFSO.DeleteFile gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg", True
            End If
            rsTmp.Save gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg", adPersistADTG
            rsTmp.Close
        Else
            If gobjFSO.FileExists(gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg") Then
                gobjFSO.DeleteFile gstrSetupPath & "\ZLUPTMP\ZLRegErr.adtg", True
            End If
        End If
    End With
    
    prgPross.Value = lngCurPro + lngCurIncPro
    SetupFiles = True
    Exit Function
ErrH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "SetupFiles", "安装注册文件发生致命错误", strErrInfo
    Call RecordErrMsg(MT_SetUp, "安装注册文件发生致命错误", strErrInfo)
    If Not gblnHelperMain Then MsgBox "安装注册文件发生致命错误，请联系管理员！信息：" & strErrInfo, vbInformation, App.Title
    Err.Clear
    lblInfor.Caption = "正在保存注册失败文件清单"
    grsFileUpgrade.Filter = "注册错误=1"
    If grsFileUpgrade.RecordCount > 0 Then
        Set rsTmp = CopyNewRec(grsFileUpgrade)
        rsTmp.Sort = "文件名"
        If gobjFSO.FileExists(gstrSetupPath & "ZLUPTMP\ZLRegErr.adtg") Then
            gobjFSO.DeleteFile gstrSetupPath & "ZLUPTMP\ZLRegErr.adtg", True
        End If
        rsTmp.Save gstrSetupPath & "ZLUPTMP\ZLRegErr.adtg", adPersistADTG
        rsTmp.Close
    End If
End Function

Public Function SetupOneFile(ByVal strSTFileName As String, ByVal strTmpFile As String, ByVal strSetupFile As String, Optional ByVal blnForceCover As Boolean, Optional ByRef strErrReturn As String) As Boolean
'功能：将升级文件放在安装路径
'说明：该功能独立出来是因为该过程存在较多Goto语句
    Dim blnGoto  As Boolean, sgResult           As VbMsgBoxResult
    Dim cllProcess  As New Collection   '进程集array(进程,Exe文件名,模块进程)
    Dim i           As Long, strMsgBox          As String
    Dim blnReturn   As Boolean, strErr          As String
    
    On Error Resume Next
    If gobjFSO.FileExists(strSetupFile) Then
        If FileSystem.GetAttr(strSetupFile) <> vbNormal Then
            FileSystem.SetAttr strSetupFile, vbNormal
        End If
    End If
    blnGoto = False
SartSetup:
     blnReturn = True: strErrReturn = ""
    If Err.Number <> 0 Then Err.Clear
    If Not gobjFSO.FileExists(strTmpFile) Then
        strErrReturn = "解压后的文件不存在，可能被杀毒软件报杀"
        Call RecordErrMsg(MT_SetUp, gobjFSO.GetFileName(strTmpFile), strErrReturn)
        gobjTrace.WriteInfo "SetupFile", "安装失败", strErrReturn
        Exit Function
    End If
    '2、开始拷贝，以及处理拷贝遇到的相关问题
    gobjFSO.CopyFile strTmpFile, strSetupFile, True
    If Err.Number <> 0 Then '拒绝权限先改名
        gobjTrace.WriteInfo "SetupFile", "拷贝安装文件失败", Err.Number & "-" & Err.Description
        If Not strSTFileName Like "ZL*" Then '系统文件
            strErrReturn = Err.Description
            Err.Clear '清除错误
            If blnForceCover Then  '强制覆盖
                If gobjFSO.FileExists(strSetupFile & "_old") Then Kill (strSetupFile & "_old")
                Call Kill(strSetupFile & "_old")
                Name strSetupFile As strSetupFile & "_old"
                Call Kill(strSetupFile & "_old")
                '重新拷贝文件
                If Err.Number <> 0 Then Err.Clear
                If Not blnGoto Then
                    blnGoto = True
                    GoTo SartSetup
                End If
            End If
        Else
            '发生错误,肯定存在文件是只读或被独占打开或已执行
            If Err.Number <> 70 And Err.Number <> 70 - 2146828288 Then
                If Not gblnHelperMain Then
                    sgResult = MsgBox("注意：" & vbCrLf & _
                                "     文件“" & strSetupFile & "”，不能升级 ,原因如下：" & vbCrLf & Err.Number & "-" & Err.Description & vbCrLf & _
                                "『重试』表示手工已经解除相关错误，重新执行升级！" & vbCrLf & _
                                "『取消』表示取消本次升级！", vbQuestion + vbRetryCancel + vbDefaultButton1, "自动升级")
                Else
                    sgResult = vbRetry
                End If
                If sgResult = vbRetry Then
                    '重新执行一次拷贝
                    If Not blnGoto Then
                        blnGoto = True
                        GoTo SartSetup
                    Else
                        blnReturn = False
                        strErrReturn = strSetupFile & "安装失败（" & Err.Number & "-" & Err.Description & ")"
                    End If
                Else
                    blnReturn = False
                    strErrReturn = strSetupFile & "安装失败（" & Err.Number & "-" & Err.Description & ")"
                End If
            Else
                Call zlGetFileProcess(strSetupFile, cllProcess)
                If strSTFileName Like "*.EXE" Then
                    If Not gblnHelperMain Then
                        sgResult = MsgBox("注意：" & vbCrLf & _
                               "     文件“" & strSetupFile & "”正在执行，不能升级！" & vbCrLf & _
                               "『终止』表示取消本部件升级！" & vbCrLf & _
                               "『重试』表示终止被运行的程序，重新执行升级！" & vbCrLf & _
                               "『忽略』表示本次不进行升级！", vbQuestion + vbAbortRetryIgnore, "自动升级")    'vbAbortRetryIgnore
                    Else
                        sgResult = vbRetry
                    End If
                ElseIf strSTFileName Like "*.OCX" Or strSTFileName Like "*.DLL" Then
                    strMsgBox = ""
                    For i = 1 To cllProcess.Count
                        If UCase(cllProcess(i)(1)) = "ZLHISCRUST.EXE" Then
                            Err.Clear
                            strErrReturn = strSetupFile & "安装失败（ZLHISCRUST.EXE自身占用，已经忽略）"
                            Exit Function
                        End If
                        If i > 2 Then
                            strMsgBox = strMsgBox & Space(5) & cllProcess(i)(0) & "：" & cllProcess(i)(1) & vbCrLf & Space(5) & "...."
                            Exit For
                        Else
                            strMsgBox = strMsgBox & Space(5) & cllProcess(i)(0) & "：" & cllProcess(i)(1) & vbCrLf
                        End If
                    Next
                    If Not gblnHelperMain Then
                        sgResult = MsgBox("注意：" & vbCrLf & _
                                "     文件“" & strSetupFile & "”正被如下程序引用，不能升级 ！" & vbCrLf & _
                                strMsgBox & vbCrLf & _
                                "『终止』表示取消本部件升级！" & vbCrLf & _
                                "『重试』表示终止被运行的程序，重新执行升级！" & vbCrLf & _
                                "『忽略』表示本次不进行升级！", vbQuestion + vbAbortRetryIgnore, "自动升级")    'vbAbortRetryIgnore
                    Else
                        sgResult = vbRetry
                    End If
                Else
                    If Not gblnHelperMain Then
                        sgResult = MsgBox("注意：" & vbCrLf & _
                                "     文件“" & strSetupFile & "”正被其他文件独站打开，不能升级 ！" & vbCrLf & _
                                "『终止』表示取消本部件升级！" & vbCrLf & _
                                "『重试』表示手工已经解除独站运行的程序，重新执行升级！" & vbCrLf & _
                                "『忽略』表示本次不进行升级！", vbQuestion + vbAbortRetryIgnore, "自动升级")
                    Else
                        sgResult = vbRetry
                    End If
                End If
                If sgResult = vbAbort Then
                    blnReturn = False
                    If strErrReturn = "" Then strErrReturn = strSetupFile & "安装失败（正被其他文件占用)"
                ElseIf sgResult = vbRetry Then
                    If strSTFileName Like "*.EXE" Or strSTFileName Like "*.OCX" Or strSTFileName Like "*.DLL" Then
                        For i = 1 To cllProcess.Count
                            Call TerminatePID(cllProcess(i)(0))
                        Next
                    End If
                    If Not blnGoto Then
                        blnGoto = True
                        GoTo SartSetup
                    End If
                ElseIf sgResult = vbIgnore Then
                    If strErrReturn = "" Then strErrReturn = strSetupFile & "安装失败（正被其他文件占用,已经忽略)"
                End If
            End If
        End If
    End If
    If Err.Number <> 0 Then Err.Clear
    SetupOneFile = blnReturn
End Function

Private Function ExecBatFile(ByVal lngCurPro As Long, ByVal lngCurIncPro As Long) As Boolean
'功能：执行批处理文件
    Dim strAutoRun      As String, strAutoRunBat As String
    Dim lngRet          As Long
    Dim lngTaskID       As Double, lngProcID As Long
    Dim objBat          As TextStream
    Dim strErrInfo      As String
    
    On Error GoTo ErrH
    gobjTrace.WriteSection "-", SL_LevelThree
    '提前升级不需要执行批处理
    If gotCurType <> OT_PreUpgrade Then
        '执行批处理文件
        lblInfor.Caption = "正在检测批处理：zlAutoRun.bat"
        strAutoRun = gstrSetupPath & "\zlAutoRun.ini"
        strAutoRunBat = gstrSetupPath & "\zlAutoRun.bat"
        '修复模式，自动生成批处理
        If gotCurType = OT_Repair And Not (gobjFSO.FileExists(strAutoRun) Or gobjFSO.FileExists(strAutoRunBat)) Then
            Set objBat = gobjFSO.CreateTextFile(strAutoRun, True)
            objBat.WriteLine gstrSetupPath & "\PUBLIC\zlMipClientShell.exe /regserver"
            objBat.WriteLine "IF ""%1""=="""" for %%c in (" & gstrSetupPath & "\PUBLIC\*.dll) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF ""%1""=="""" for %%c in (" & gstrSetupPath & "\PUBLIC\*.ocx) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF not ""%1""=="""" for %%c in (" & gstrSetupPath & "\PUBLIC\*.dll) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF not ""%1""=="""" for %%c in (" & gstrSetupPath & "\PUBLIC\*.ocx) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF ""%1""=="""" for %%c in (" & gstrSetupPath & "\apply\*.dll) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF ""%1""=="""" for %%c in (" & gstrSetupPath & "\apply\*.ocx) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF not ""%1""=="""" for %%c in (" & gstrSetupPath & "\apply\*.dll) do regsvr32.exe /s %%c"
            objBat.WriteLine "IF not ""%1""=="""" for %%c in (" & gstrSetupPath & "\apply\*.ocx) do regsvr32.exe /s %%c"
            objBat.Close
            Set objBat = Nothing
        End If
        If gobjFSO.FileExists(strAutoRun) Or gobjFSO.FileExists(strAutoRunBat) Then
            On Error Resume Next
            If gobjFSO.FileExists(strAutoRun) Then
                If gobjFSO.FileExists(strAutoRunBat) Then Call gobjFSO.DeleteFile(strAutoRunBat, True)
                Name strAutoRun As gstrSetupPath & "\zlAutoRun.bat"
            End If
            Call Kill(strAutoRun)
            lngTaskID = Shell(gstrSetupPath & "\zlAutoRun.bat", vbHide)  'SW_SHOW
            If lngTaskID <> 0 Then
                lngProcID = OpenProcess(SYNCHRONIZE, False, lngTaskID)
                If lngProcID <> 0 Then
                    DoEvents
                    lngRet = WaitForSingleObject(lngProcID, INFINITE)
                    lngRet = CloseHandle(lngProcID)
                End If
                gobjTrace.WriteInfo "ExecBatFile", "批处理文件执行", "成功"
            Else
                gobjTrace.WriteInfo "ExecBatFile", "批处理文件执行", "失败"
                Call RecordErrMsg(MT_ExeBat, "批处理文件执行", "批处理文件执行失败")
                prgPross.Value = lngCurPro + lngCurIncPro
                Me.Refresh
                Exit Function
            End If
        End If
    End If
    prgPross.Value = lngCurPro + lngCurIncPro
    Me.Refresh
    ExecBatFile = True
    Exit Function
ErrH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "SetupFiles", "批处理文件执行发生致命错误", strErrInfo
    Call RecordErrMsg(MT_SetUp, "批处理文件执行发生致命错误", strErrInfo)
    If Not gblnHelperMain Then MsgBox "批处理文件执行发生致命错误，请联系管理员！信息：" & strErrInfo, vbInformation, App.Title
    Err.Clear
End Function

Private Function IsOldComponentExists(ByVal strName As String, Optional ByRef strExistsFile As String) As Boolean
'功能：当部件清单不存在时，判断是否是老的部件，没有在部件清单中
'参数：strName=部件名，可能不带后缀
    Dim varItem         As Variant, strFileTmp              As String
    Dim strTmp As String
    
    On Error Resume Next
    If mcllOldComs Is Nothing Then
        Set mcllOldComs = New Collection
    End If
    strName = UCase(strName)
    strExistsFile = ""
    strExistsFile = mcllOldComs("K_" & strName)
    If Err.Number <> 0 Then
        Err.Clear
        If strName Like "*.DLL" Or strName Like "*.EXE" Then
            For Each varItem In gcllSetPath
                strFileTmp = varItem & "\" & strName
                If gobjFSO.FileExists(strFileTmp) Then
                    strExistsFile = strFileTmp
                    mcllOldComs.Add strFileTmp, "K_" & strName
                    mcllOldComs.Add strFileTmp, "K_" & Mid(strName, 1, Len(strName) - 4)
                    IsOldComponentExists = True
                    Exit For
                End If
            Next
        Else
            For Each varItem In gcllSetPath
                strFileTmp = varItem & "\" & strName & ".DLL"
                If gobjFSO.FileExists(strFileTmp) Then
                    strExistsFile = strFileTmp
                    mcllOldComs.Add strFileTmp, "K_" & strName & ".DLL"
                    mcllOldComs.Add strFileTmp, "K_" & strName
                    IsOldComponentExists = True
                    Exit For
                End If
                strFileTmp = varItem & "\" & strName & ".EXE"
                If gobjFSO.FileExists(strFileTmp) Then
                    strExistsFile = strFileTmp
                    mcllOldComs.Add strFileTmp, "K_" & strName & ".EXE"
                    mcllOldComs.Add strFileTmp, "K_" & strName
                    IsOldComponentExists = True
                    Exit For
                End If
            Next
        End If
    Else
        IsOldComponentExists = strExistsFile <> ""
    End If
End Function

Private Function ReportCheckInfo(ByRef blnOprateOK As Boolean) As Boolean
    Dim lngLoop     As Long
    
    gobjTrace.WriteSection "安装注册文件", SL_LevelThree
    On Error GoTo ErrH
    With grsFileUpgrade
        grsFileUpgrade.Filter = "(更新<" & UC_Normal & " And 错误信息=NULL) Or (更新=" & UC_Normal & " And 清理文件路径<>NULL)"
        .Sort = "判断批次,自动注册,类型排序"
        blnOprateOK = .RecordCount = 0
        For lngLoop = 1 To .RecordCount
            gobjTrace.WriteSection "-", SL_LevelThree
            If !更新 > 0 Then
                gobjTrace.WriteInfo "SetupFiles", "安装注册文件", !文件名
            Else
                gobjTrace.WriteInfo "SetupFiles", "清理错误文件", !文件名
            End If
            If !更新 > 0 Then
                Call ReportInfo("部件：" & !文件名 & "(" & !实际路径 & ")需要更新")
            ElseIf !更新 = UC_NewDown Then
                Call ReportInfo("部件：" & !文件名 & "(" & !实际路径 & ")缺失，需要下载")
            End If
            
            If Not IsNull(!清理文件路径) Then
                Call ReportInfo("部件：" & !文件名 & "(" & !实际路径 & ")存在错误路径文件：" & Replace(!清理文件路径, "|", ","))
            End If
            
            If Not IsNull(!附加实际路径) Then
                Call ReportInfo("部件：" & !文件名 & "(" & !实际路径 & ")的如下附加安装路径需要更新：" & Replace(!附加实际路径, "|", ","))
            End If
            .MoveNext
        Next
    End With
    ReportCheckInfo = True
    Exit Function
ErrH:
    Call RecordErrMsg(MT_ChcekUpdate, "上传信息检查出现错误", Err.Description)
End Function

