VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLabMBControl 
   BorderStyle     =   0  'None
   Caption         =   "仪器控制"
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraMain 
      Height          =   1605
      Left            =   30
      TabIndex        =   0
      Top             =   -75
      Width           =   6390
      Begin VB.CommandButton cmdCanCel 
         Cancel          =   -1  'True
         Caption         =   "停止(&S)"
         Height          =   350
         Left            =   2670
         TabIndex        =   3
         Top             =   1080
         Width           =   1100
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   405
         Left            =   150
         TabIndex        =   1
         Top             =   585
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   714
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSCommLib.MSComm MSComm 
         Left            =   5625
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.Label lbl信息 
         AutoSize        =   -1  'True
         Caption         =   "准备："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   225
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmLabMBControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conCmd_Begin = "Begin"            '从解析程序中取得仪器命令的命令常量
Const ConCmd_Out = "Out"                '出板命令
Const ConCmd_In = "In"                 '进板命令
Const conCmd_End = "End"
Const conCmd_Revert = "Revert"
Const conCmd_Play = "Play"
Const conCmd_ReadData = "ReadData" '读数
Const conCmd_SpecialConnType = "SpecialConnType"

Dim mstrCmdRevert As String              ' 仪器通用应答命令
Dim mobjDevice As Object                '接口
Public bln_Init As Boolean                 '仪器是否已初始化

Public Function MB_Start(objfrm As Object, ByVal strMachineID As Long) As Boolean
    Dim rsTmp As New adodb.Recordset
    Dim str通讯口 As String, str波特率 As String, str数据位 As String, str停止位 As String, str校验位 As String

    On Error GoTo ErrHandle
    
   
    str通讯口 = zlDatabase.GetPara("frmLabMB_通讯口", 100, 1208, "COM1")
    str波特率 = zlDatabase.GetPara("frmLabMB_波特率", 100, 1208, "9600")
    str数据位 = zlDatabase.GetPara("frmLabMB_数据位", 100, 1208, "8")
    str停止位 = zlDatabase.GetPara("frmLabMB_停止位", 100, 1208, "1")
    str校验位 = zlDatabase.GetPara("frmLabMB_校验位", 100, 1208, "N")
    
    
    str通讯口 = Replace(str通讯口, "COM", "")
    str校验位 = Replace(Replace(Replace(Replace(Replace(str校验位, "E-偶数", "E"), "M-标记", "M"), "N-缺省", "N"), "O-奇数", "O"), "S-空格", "S")
    str校验位 = Replace(str校验位, "None", "N")
    
    gstrSql = "select 通讯程序名 from 检验仪器 where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strMachineID)
    
    If rsTmp.EOF = True Then MsgBox "没有找到仪器!", vbInformation, Me.Caption: Exit Function
    
    If Nvl(rsTmp("通讯程序名")) = "" Then MsgBox "仪器通讯程序名为空,请到检验仪器管理中修改!", vbInformation: Exit Function
    
    If Not mobjDevice Is Nothing Then Set mobjDevice = Nothing
    Set mobjDevice = CreateObject(rsTmp("通讯程序名"))
    If mobjDevice Is Nothing Then
        MsgBox "创建解析部件失败！", vbInformation, Me.Caption
        Exit Function
    End If
    
    
    '发送控制指令
    If MSComm.PortOpen Then MSComm.PortOpen = False
    MSComm.CommPort = CInt(str通讯口)
    MSComm.Settings = str波特率 & "," & str校验位 & "," & str数据位 & "," & str停止位    '"9600,N,8,1"
    MSComm.InputLen = 0
    MSComm.PortOpen = True
    
    Me.Show , objfrm
    If bln_Init Then
        Call MB_SendCommand(conCmd_End, "释放仪器控制......", 3)
    End If
    
    '=========================================开始控制======================================
    If Not MB_SendCommand(conCmd_Begin, "连接仪器...", 2) Then Exit Function
    '========================================================================================
    
    '==================================OUT 可选步骤=========================================
    If Not MB_SendCommand(ConCmd_Out, "正在弹出微孔板...", 2) Then Exit Function
    
    '========================================================================================
    bln_Init = True
    MB_Start = bln_Init
    Me.Hide
    Exit Function
ErrHandle:
    MsgBox "连接仪器时，出现错误！" & vbNewLine & "[" & Err.Number & "] " & Err.Description, vbInformation, Me.Caption
End Function

Public Sub MB_Stop()
    '结束,断开连接
    Dim strCmd As String
    On Error GoTo errH
    Me.Show
    If bln_Init Then
        Call MB_SendCommand(conCmd_End, "释放仪器控制......", 3)
        If MSComm.PortOpen Then MSComm.PortOpen = False
        bln_Init = False
    End If
    Me.Hide
    Exit Sub
errH:
    MsgBox "释放仪器控制时出现错误！" & vbNewLine & "[" & Err.Number & "] " & Err.Description, vbInformation, Me.Caption
End Sub

Public Sub ShowMe(objfrm As Object, ByVal strControl As String, strResult As String)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '参数               Objfrm 父窗体
    '                   strMachineID    仪器ID
    '                   strControl (1:2:3:4:5:6) (波长;振板频率;振板时间;进板方式;空白形式:参考波长)
    '                   返回结果
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim aItem() As String
    Dim intLoop As Integer             '循环时使用
    Dim aRow() As String, aResult() As String
    Dim intRow As Integer, intCol As Integer
    Dim TestData(1, 1 To 8, 1 To 12) As String
    
    On Error GoTo errH
    If Not bln_Init Then
        MsgBox "请先连接仪器！", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.Show , objfrm
    
    '==================================开始准备测量==================================================
    Call MB_SendCommand(ConCmd_In, "正在关闭微孔板......", 2)
    
    aItem = Split(strControl, ";")
    For intLoop = 0 To UBound(aItem) - 1
        Select Case intLoop
            Case 0
                If aItem(intLoop) <> "" Then
                    If Not MB_SendCommand(aItem(intLoop), "正在设置波长...") Then Exit Sub
                End If
            Case 1
                If aItem(intLoop) <> "" Then
                    If Not MB_SendCommand(aItem(intLoop), "正在设置振板频率...") Then Exit Sub
                End If
            Case 2
                If aItem(intLoop) <> "" Then
                    If Not MB_SendCommand(aItem(intLoop), "正在设置振板时间...") Then Exit Sub
                End If
            Case 3
                If aItem(intLoop) <> "" Then
                    If Not MB_SendCommand(aItem(intLoop), "正在设置进板方式...") Then Exit Sub
                End If
            Case 4
                If aItem(intLoop) <> "" Then
                    If Not MB_SendCommand(aItem(intLoop), "正在设置空白形式...") Then Exit Sub
                End If
        End Select
    Next
    '========================================开始测量(主波长)=====================================================
    If Not MB_SendCommand(conCmd_Play, "开始测量(主波长)...") Then Exit Sub
    If Not MB_SendCommand(conCmd_ReadData, "开始读取数据(主波长)...", 15, 1, strResult) Then Exit Sub
    
    '========================================开始测量(参考波长)===================================================
    If aItem(5) <> "" Then
        '--保存主波长
        aRow = Split(strResult, "|")
        For intRow = 1 To 8
            aResult = Split(aRow(intRow - 1), ";")
            For intCol = 1 To 12
                TestData(0, intRow, intCol) = aResult(intCol - 1)
            Next
        Next
        
        '设置参考波长
        If Not MB_SendCommand(aItem(5), "正在设置参考波长...") Then Exit Sub
        
        '开始测量参考波长
        strResult = ""
        If Not MB_SendCommand(conCmd_Play, "开始测量(参考波长)...") Then Exit Sub
        
        If Not MB_SendCommand(conCmd_ReadData, "开始读取数据(参考波长)...", 15, 1, strResult) Then Exit Sub
        
        '保存参考波长
        aRow = Split(strResult, "|")
        For intRow = 1 To 8
            aResult = Split(aRow(intRow - 1), ";")
            For intCol = 1 To 12
                TestData(1, intRow, intCol) = aResult(intCol - 1)
            Next
        Next

        '计算
        strResult = ""
        For intRow = 1 To 8
            strResult = strResult & "|"
            For intCol = 1 To 12
                strResult = strResult & ";" & TestData(0, intRow, intCol) - TestData(1, intRow, intCol)
            Next
        Next
        strResult = Replace(strResult, "|;", "|")
        strResult = Mid(strResult, 2)
    End If
    
    '=======================================弹出微孔板，可选过程=======================================
    If Not MB_SendCommand(ConCmd_Out, "正在弹出微孔板...", 2) Then Exit Sub
    
    Me.Hide
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function MB_SendCommand(ByVal str_Command As String, ByVal str_Info As String, Optional ByVal intOutTime As Integer = 10, Optional ByVal intType As Integer = 0, Optional ByRef str_Data As String) As Boolean
    '发送命令到酶标仪
    
    'str_Command :要发到仪器的命令,是内部定义的，每个仪器都一样。
    'str_Info   ：提示信息
    'intOutTime ：执行此命令时，超时设置，默认10
    'intType   : 0-不返回数据 1-返回数据
    'str_Data  : 保存 返回的数据
    Dim str_R As String                         '每个命令发出后，仪器的应答指令。
    Dim strCmdRevert As String                  '通用的应答指令
    Dim strCmd As String                        '控制命令
    Dim lngBeginDate As Long, strGetCmd As String
    
    If mobjDevice Is Nothing Then Exit Function '解析程序未初始化，退出
    If Not MSComm.PortOpen Then Exit Function '串口未打开，退出
    
    Dim strReserved As String  '调用解析时用，占位置用
    Dim blnErr As Boolean, int_I As Integer, var_R As Variant
    Dim int超时 As Integer
    Dim strSpecialConnType As String
    Dim strgetval As String
    
    On Error GoTo ErrHandle
    
    If mstrCmdRevert = "" Then mstrCmdRevert = mobjDevice.CmdAnalyse(conCmd_Revert)
    
    strCmd = mobjDevice.CmdAnalyse(str_Command)
    int超时 = Val(mobjDevice.CmdAnalyse(str_Command & "_TimeOut"))
    If int超时 <= 0 Then
        int超时 = intOutTime
    End If
    
    '读取数据时的需要不停的发送指令(特殊仪器才会用到)=1时生效
    strSpecialConnType = mobjDevice.CmdAnalyse(conCmd_SpecialConnType)
    
    '--- 日志
    MbLog "frmLabMBControl", "MB_SendCommand", strCmd, int超时
    
    lngBeginDate = Timer
    If Trim(strCmd) <> "" Then
        
        lbl信息.Caption = str_Info
        
        If InStr(strCmd, "|") > 0 Then '有专门的应答指令
            str_R = Mid(strCmd, InStr(strCmd, "|") + 1)
            strCmd = Mid(strCmd, 1, InStr(strCmd, "|") - 1)
        Else
            str_R = mstrCmdRevert      '通用的应答指令
        End If
        MSComm.Output = strCmd
        strGetCmd = ""
        If intType = 0 Then '不返回数据
            Do
                DoEvents
                strGetCmd = strGetCmd & MSComm.Input
                
                Call ShowPbar((CLng(Timer) - lngBeginDate) / int超时 * 100)
            Loop Until InStr(strGetCmd, str_R) Or (CLng(Timer) - lngBeginDate > int超时)
                            '--- 日志
            MbLog "frmLabMBControl", "接收应答指令", strGetCmd, str_R
            
            If Trim(strGetCmd) = "" Then
                '超时处理
                Debug.Print Timer & " " & lngBeginDate
                MsgBox "执行" & str_Command & "命令超时!", vbInformation, Me.Caption
                Exit Function
            Else
                If InStr(str_R, "|") > 0 Then
                    var_R = Split(str_R, "|")
                    blnErr = True
                    For int_I = LBound(var_R) To UBound(var_R)
                        If InStr(strGetCmd, var_R(int_I)) >= 0 Then
                            blnErr = False
                            Exit For
                        End If
                    Next
                Else
                    blnErr = InStr(strGetCmd, str_R) <= 0
                End If
                If blnErr Then
                    MsgBox "执行" & str_Command & "命令，仪器返回的数据有误!" & vbNewLine & strGetCmd, vbInformation, Me.Caption
                    Exit Function
                End If
            End If
        Else                '要返回解析数据
            Do
               DoEvents
               '处理特殊的仪器不停的发连接指令
                If strSpecialConnType = "1" Then
                    MSComm.Output = strCmd
                    Call Sleep(1000)
                End If
               strGetCmd = strGetCmd & MSComm.Input
               Call ShowPbar((CLng(Timer) - lngBeginDate) / int超时 * 100)
               mobjDevice.Analyse strGetCmd, str_Data, strReserved, ""
               strgetval = strGetCmd
               strGetCmd = strReserved
            Loop Until str_Data <> "" Or (CLng(Timer) - lngBeginDate > int超时)
            
            '--- 日志
            MbLog "frmLabMBControl", "接收酶标数据", strgetval & "|" & strGetCmd, str_Data

            If Trim(str_Data) = "" And Trim(strGetCmd) = "" Then
                MsgBox "接收数据失败!", vbInformation, Me.Caption
                Exit Function
            ElseIf Trim(strGetCmd) <> "" And Trim(str_Data) = "" Then
                MsgBox "解析数据失败!", vbInformation, Me.Caption
                Exit Function
            End If
        End If
        

    End If
    MB_SendCommand = True
    Exit Function
ErrHandle:
    'If MSComm.PortOpen Then MSComm.PortOpen = False
    MsgBox "执行" & str_Command & "命令出现错误!" & vbNewLine & "[" & Err.Number & "] " & Err.Description
End Function

Private Sub CmdCancel_Click()
    If MsgBox("将关闭与仪器的连接，仪器未传输的数据将不能接收，请确认？" & vbNewLine & "点[确认]，将关闭与仪器的连接；点[取消]，继续原来的操作。", vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then
        Unload Me
    End If
End Sub

Private Sub ShowPbar(ByVal sinValue As Single)
    On Error Resume Next
    ProgressBar1.Value = sinValue
End Sub


Private Sub MbLog(ByVal strModule As String, ByVal strFunc As String, ByVal strInput As String, ByVal strOutput As String)
    '调用公共方法记录日志
    Call zl9Comlib.LogWrite("LIS老版通讯程序调试日志", strModule, strFunc, strInput & vbCrLf & strOutput)
End Sub

