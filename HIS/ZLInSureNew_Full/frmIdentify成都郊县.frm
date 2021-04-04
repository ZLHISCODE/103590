VERSION 5.00
Begin VB.Form frmIdentify成都郊县 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify成都郊县.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.CheckBox chk生育标志 
      Caption         =   "生育标志"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1590
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4350
      Width           =   2985
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   1950
      Locked          =   -1  'True
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2010
      Width           =   2595
   End
   Begin VB.CheckBox chk手工 
      Caption         =   "手工输入卡内数据(&M)"
      Height          =   240
      Left            =   1470
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   660
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.ComboBox cbo保险类别 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   360
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2490
      Width           =   2595
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "校验(&T)"
      Height          =   405
      Left            =   240
      TabIndex        =   20
      Top             =   4920
      Width           =   1305
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   1965
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   17
      Top             =   3870
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1965
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   15
      Top             =   3420
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1965
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   2970
      Width           =   2595
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1965
      Locked          =   -1  'True
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1530
      Width           =   2595
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H8000000F&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1965
      Locked          =   -1  'True
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1065
      Width           =   2595
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1965
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   615
      Width           =   2595
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   405
      Left            =   1965
      TabIndex        =   21
      Top             =   4920
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   3450
      TabIndex        =   22
      Top             =   4920
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   -210
      TabIndex        =   19
      Top             =   4665
      Width           =   6660
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "医疗照顾人员"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   330
      TabIndex        =   8
      Top             =   2070
      Width           =   1530
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "保险类别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   840
      TabIndex        =   10
      Top             =   2550
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "确认密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   840
      TabIndex        =   16
      Top             =   3930
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "新密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   1095
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1350
      TabIndex        =   12
      Top             =   3030
      Width           =   510
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "分中心编码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   585
      TabIndex        =   6
      Top             =   1590
      Width           =   1275
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "个人编码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   1125
      Width           =   1020
   End
   Begin VB.Label lblNote 
      Caption         =   "请在正确刷卡之后，输入个人密码。"
      Height          =   255
      Left            =   930
      TabIndex        =   0
      Top             =   225
      Width           =   3645
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "卡号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1350
      TabIndex        =   2
      Top             =   675
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frmIdentify成都郊县.frx":030A
      Top             =   345
      Width           =   480
   End
End
Attribute VB_Name = "frmIdentify成都郊县"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<0表示错误码;0连接正确且插卡;1-连接正确但未插卡
Private Declare Function IC_Status_1 Lib "JKIC32.DLL" Alias "IC_Status" (ByVal lngICDev As Long) As Integer
'<0错误;串口标识符
Private Declare Function IC_InitComm_1 Lib "JKIC32.DLL" Alias "IC_InitComm" (ByVal IntPort As Integer) As Long
'<0未能正常关闭串口;>=0正常关闭
Private Declare Function IC_ExitComm_1 Lib "JKIC32.DLL" Alias "IC_ExitComm" (ByVal lngICDev As Long) As Integer
'<0错误;=0正常
Private Declare Function IC_InitType_1 Lib "JKIC32.DLL" Alias "IC_InitType" (ByVal lngICDev As Long, ByVal intType As Integer) As Integer
'<0错误;=0正常
Private Declare Function IC_Read_1 Lib "JKIC32.DLL" Alias "IC_Read_Hex" (ByVal lngICDev As Long, ByVal intOffset As Integer, ByVal intLen As Integer, ByVal strData As String) As Integer

'------------------------------------------------------------
Private Declare Function IC_InitComm_2 Lib "ftic_32.dll" Alias "ic_init" (ByVal Port%, ByVal baud As Long) As Long
Private Declare Function IC_Status_2 Lib "ftic_32.dll" Alias "get_status" (ByVal icdev As Long, intCard As Integer) As Integer
Private Declare Function chk_card Lib "ftic_32.dll" (ByVal icdev As Long) As Integer
Private Declare Function IC_Read_2 Lib "ftic_32.dll" Alias "srd_4442" (ByVal icdev As Long, ByVal offset As Long, ByVal Length As Long, ByVal r_string As String) As Integer
Private Declare Function IC_Down_2 Lib "ftic_32.dll" Alias "auto_pull" (ByVal icdev As Long) As Integer
Private Declare Function ic_exit% Lib "ftic_32.dll" (ByVal icdev As Long)
Private Declare Function hex_asc% Lib "ftic_32.dll" (ByVal hex As String, ByVal asc$, ByVal le&)

Private Declare Function srd_4442 Lib "ftic_32.dll" (ByVal icdev As Long, ByVal offset As Long, ByVal Length As Long, ByRef r_string As Byte) As Integer
'------------------------------------------------------------

'说明：由于完成的功能极其相似，本窗体除了完成贵阳医保的身份验证外。还完成了成都郊县医保的身份验证
Private mstr卡号 As String
Private mstr医保号 As String
Private mstr分中心编号 As String
Private mstr密码 As String
Private mstr保险类别 As Integer
Private mintInsure As Integer
Private mblnPass As Boolean
Private mblnChangePassword As Boolean
Private mbln生育标志 As Boolean

Private mint卡类型 As Integer
Private mint端口号 As Integer
Private mint地区 As Integer

Private mblnOK As Boolean

Private Sub chk手工_Click()
    Dim lngColor As Long, blnEnable As Boolean
    
    If mintInsure <> TYPE_贵阳市 Then Exit Sub
    cmdOK.Enabled = False
    If chk手工.Value = 1 Then
        '打开手工输入项:分中心编号,保险类别
        blnEnable = True
        lngColor = &H80000005
        
        txtEdit(2).TabStop = True
    Else
        '关闭手工输入项
        blnEnable = False
        lngColor = &H8000000F
        
        txtEdit(2).Text = ""
        cbo保险类别.ListIndex = 0
    End If
    
    txtEdit(2).Locked = Not blnEnable
    txtEdit(2).BackColor = lngColor
    cbo保险类别.Enabled = blnEnable
    cbo保险类别.BackColor = lngColor
    txtEdit(0).SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    If cmdOK.Enabled = False Then Exit Sub
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If txtEdit(lngIndex).Visible = True Then
            If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), IIf(lngIndex = 0, 20, txtEdit(lngIndex).MaxLength)) = False Then
                If txtEdit(lngIndex).Enabled Then txtEdit(lngIndex).SetFocus
                Exit Sub
            End If
        End If
    Next
    If mintInsure = TYPE_贵阳市 Then
        If Trim(txtEdit(0).Text) = "" Then
            MsgBox "未正确地刷卡,不能通过验证！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If Trim(txtEdit(1).Text) = "" Then
        MsgBox "未正确地刷卡,不能通过验证！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstr卡号 = Trim(txtEdit(0).Text)
    mstr医保号 = Trim(txtEdit(1).Text)
    mstr分中心编号 = Trim(txtEdit(2).Text)
    mstr密码 = Trim(txtEdit(3).Text)
    mstr保险类别 = cbo保险类别.ListIndex + 1
    mbln生育标志 = (chk生育标志.Value = 1)
    
    If mblnChangePassword = True Then
        '当前功能是更改密码
        If mintInsure = TYPE_贵阳市 Then
            If txtEdit(4).Text <> "" Or txtEdit(5).Text <> "" Then
                If txtEdit(4).Text <> txtEdit(5).Text Then
                    MsgBox "两次输入的新密码不相同，请重新输入。", vbInformation, gstrSysName
                    If txtEdit(5).Enabled Then txtEdit(5).SetFocus
                    Exit Sub
                End If
            End If
        Else
            If txtEdit(4).Text = "" Then
                MsgBox "请输入新的密码。", vbInformation, gstrSysName
                If txtEdit(4).Enabled Then txtEdit(4).SetFocus
                Exit Sub
            End If
            If txtEdit(4).Text <> txtEdit(5).Text Then
                MsgBox "两次输入的新密码不相同，请重新输入。", vbInformation, gstrSysName
                If txtEdit(5).Enabled Then txtEdit(5).SetFocus
                Exit Sub
            End If
        End If
        
        If Trim(txtEdit(4).Text) <> "" Then
            If mintInsure = type_成都郊县 Then
                If 更改密码_成都郊县(mstr卡号, mstr医保号, mstr分中心编号, mstr密码, txtEdit(4).Text) = False Then Exit Sub
            ElseIf mintInsure = TYPE_新都 Then
                If 更改密码_新都(mstr卡号, mstr医保号, mstr分中心编号, mstr密码, txtEdit(4).Text) = False Then Exit Sub
            ElseIf mintInsure = TYPE_贵阳市 Then
                If 更改密码_贵阳市(txtEdit(0).Tag, mstr密码, txtEdit(4).Text) = False Then Exit Sub
            End If
            mstr密码 = Trim(txtEdit(4).Text)
        End If
    End If
    mblnOK = True
    Unload Me
End Sub

Public Function GetIdentify(ByVal intinsure As Integer, str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, _
                            Optional ByVal blnPass As Boolean = True, Optional ByVal blnChangePassword As Boolean = False, Optional ByRef bln生育标志 As Boolean = False) As Boolean
    Dim sinDec As Single
    Dim intAdjust As Integer
    Dim rsTemp As New ADODB.Recordset
    
    mblnOK = False
    mblnPass = blnPass
    mblnChangePassword = blnChangePassword
    mintInsure = intinsure
    
    If intinsure = type_成都郊县 Or intinsure = TYPE_新都 Then
        '以密码显示卡号等信息
        txtEdit(0).PasswordChar = "*"
        txtEdit(1).PasswordChar = "*"
        txtEdit(2).PasswordChar = "*"
        
        '屏蔽贵阳医保部分控件
        chk手工.Visible = False
        lblEdit(6).Visible = False
        cbo保险类别.Visible = False
        lblEdit(7).Visible = False
        txtEdit(6).Visible = False
        cmdTest.Visible = False
        cmdOK.Enabled = True
        '调整位置
        sinDec = txtEdit(0).Top - lblEdit(0).Top
        For intAdjust = 0 To 5
            If intAdjust = 0 Then
                txtEdit(intAdjust).Top = chk手工.Top
                lblEdit(intAdjust).Top = txtEdit(intAdjust).Top - sinDec
            Else
                txtEdit(intAdjust).Top = txtEdit(intAdjust - 1).Top + 510
                lblEdit(intAdjust).Top = txtEdit(intAdjust).Top - sinDec
            End If
        Next
        Frame1.Top = txtEdit(5).Top + txtEdit(5).Height + 180
        cmdOK.Top = Frame1.Top + 200
        cmdCancel.Top = cmdOK.Top
        cmdTest.Top = cmdOK.Top
        
        '曾明春(2005-12-28):提取适用地区
        gstrSQL = "Select 参数名,Nvl(参数值,0) Value From 保险参数 Where 参数名='适用地区' and 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取上传入院信息参数值", intinsure)
        mint地区 = rsTemp("value")
        
        '取保险参数
        mint卡类型 = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "卡类型", 0)
        mint端口号 = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "IC设备端口", 1)
    ElseIf intinsure = TYPE_贵阳市 Then
        cmdOK.Enabled = False
        
        With cbo保险类别
            .Clear
            .AddItem "企业职工基本医疗保险"
            .AddItem "企业离休医疗保险"
            .AddItem "机关事业单位医疗保险"
            .ListIndex = 0
        End With
        chk手工.Value = 0
    End If
    
    lblEdit(4).Visible = blnChangePassword
    lblEdit(5).Visible = blnChangePassword
    txtEdit(4).Visible = blnChangePassword
    txtEdit(5).Visible = blnChangePassword
    
    If blnChangePassword = False Then
        Frame1.Top = txtEdit(4).Top
        cmdOK.Top = Frame1.Top + 200
        cmdCancel.Top = cmdOK.Top
        cmdTest.Top = cmdOK.Top
    End If
    
    frmIdentify成都郊县.Height = cmdOK.Top + cmdOK.Height + 500
    frmIdentify成都郊县.Show vbModal
    
    GetIdentify = mblnOK
    If mblnOK = True Then
        If mintInsure = TYPE_贵阳市 Then
            str卡号 = mstr卡号 & "^" & mstr保险类别
        Else
            str卡号 = mstr卡号
        End If
        str医保号 = mstr医保号
        str分中心编号 = mstr分中心编号
        str密码 = mstr密码
    End If
    
    bln生育标志 = mbln生育标志
End Function

Private Sub cmdTest_Click()
    Dim str类别 As String
    Dim str性别 As String
    '将卡号传给接口分析,取其返回结果并更新界面
    
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "请刷卡！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If InitXML = False Then Exit Sub
    If chk手工.Value = 0 Then
        '必须先修改密码
        If mblnChangePassword = True Then
            '当前功能是更改密码
            If txtEdit(4).Text <> "" Or txtEdit(5).Text <> "" Then
                If txtEdit(4).Text <> txtEdit(5).Text Then
                    MsgBox "两次输入的新密码不相同，请重新输入。", vbInformation, gstrSysName
                    If txtEdit(5).Enabled Then txtEdit(5).SetFocus
                    Exit Sub
                End If
            End If
            
            If Trim(txtEdit(4).Text) <> "" Then
                If 更改密码_贵阳市(txtEdit(0).Text, txtEdit(3).Text, txtEdit(4).Text) = False Then Exit Sub
                txtEdit(3).Text = txtEdit(4).Text
                txtEdit(4).Text = ""
                txtEdit(5).Text = ""
                mstr密码 = Trim(txtEdit(3).Text)
            End If
        End If
    
        If InitXML = False Then Exit Sub
        Call InsertChild(mdomInput.documentElement, "CARDDATA", txtEdit(0).Text)            ' 磁卡数据
        Call InsertChild(mdomInput.documentElement, "PASSWORD", txtEdit(3).Text)            ' 密码
    Else
        Call InsertChild(mdomInput.documentElement, "CARDID", txtEdit(0).Text)              ' 磁卡数据
        Call InsertChild(mdomInput.documentElement, "CENTERCODE", txtEdit(2).Text)          ' 分中心编码
        Call InsertChild(mdomInput.documentElement, "INSURETYPE", cbo保险类别.ListIndex + 1) ' 保险类别
        Call InsertChild(mdomInput.documentElement, "PASSWORD", txtEdit(3).Text)            ' 密码
    End If
    
    '调用接口
    If CommServer(IIf(chk手工.Value = 0, "READCARD", "READCARD_M")) = False Then Exit Sub
    
    '取得返回值
    txtEdit(0).Tag = txtEdit(0).Text                    '保存卡内数据，以便更新密码时使用
    txtEdit(0).Text = GetElemnetValue("CARDID")
    txtEdit(1).Text = GetElemnetValue("PERSONCODE")
    txtEdit(2).Text = GetElemnetValue("CENTERCODE")
    txtEdit(6).Text = IIf(Val(GetElemnetValue("CAREPSNFLAG")) = 0, "否", "是")
    str性别 = GetElemnetValue("SEX")
    str性别 = Switch(str性别 = "1", "男", str性别 = "2", "女", str性别 = "9", "其它", True, str性别)
    If str性别 = "女" Then chk生育标志.Enabled = True
    cbo保险类别.ListIndex = GetElemnetValue("INSURETYPE") - 1
    cmdOK.Enabled = True
End Sub

Private Sub Form_Load()
  gblnLED = Val(GetSetting("ZLSOFT", "公共全局", "使用", 0)) <> 0
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mintInsure = TYPE_贵阳市 Then
        If chk手工.Value = 0 Then
            If Index = 0 Then cmdOK.Enabled = False
        Else
            If Index = 0 Or Index = 2 Then cmdOK.Enabled = False
        End If
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    '请刷卡
    If Index = 0 Then
        If gblnLED And txtEdit(0).Text = "" Then
            zl9LedVoice.Speak "#5"
        End If
    End If
    '请输入密码
    If Index = 3 Then
        If gblnLED And txtEdit(3).Text = "" Then
            zl9LedVoice.Speak "#0"
        End If
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim varSplit As Variant

    
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            If mintInsure = TYPE_贵阳市 Then
                '由自己对刷卡内容进行解析(此项功能禁止,由对方接口分析)
'                txtEdit(0).Text = Replace(txtEdit(0).Text, vbCr, "")
'                txtEdit(0).Text = Replace(txtEdit(0).Text, vbLf, "")
'                If Right(txtEdit(0).Text, 1) = "?" Then
'                    '可以对刷卡信息进行分解了
'                    '磁卡数据格式为：;:卡号=个人编码=分中心编码?
'                    varSplit = Split(txtEdit(0).Text, "=")
'                    txtEdit(0).Text = Mid(varSplit(0), 3)
'                    If UBound(varSplit) > 0 Then txtEdit(1).Text = varSplit(1)
'                    If UBound(varSplit) > 1 Then txtEdit(2).Text = Mid(varSplit(2), 1, Len(varSplit(2)) - 1)
'
'                    If mblnPass = True Then
'                        txtEdit(3).SetFocus
'                    Else
'                        cmdOK_Click
'                    End If
'                End If
                If Trim(txtEdit(0).Text) <> "" Then
                    If mblnPass = True Then
                        txtEdit(3).SetFocus
                    Else
                        Call cmdTest_Click
                    End If
                End If
            Else
                '成都郊县，调用系统提供的函数进行解析
                Dim lngDev As Long, lngBoud As Long, lngReturn As Long, intReturn As Integer, intCard As Integer
                Dim str中心编号  As String, str医保号 As String, str卡号 As String
                Dim str医保号_IC As String * 256
                Dim str卡号_IC As String * 256
                Dim strData As String * 256
                
                '曾明春(2006-01-18):增加变量
                Dim by卡号_IC(256) As Byte
                Dim intnum As Integer
                Dim StrHex As String, strTemp As String, str卡号_IC2 As String, str医保号_IC2 As String
                
                '如果是IC卡,需要先解析出卡内的数据
                If mint卡类型 <> 0 Then
                    '初始化端口号
                    lngBoud = 9600
                    Call DebugTool("准备打开端口")
                    If mint卡类型 = 1 Then
                        lngDev = IC_InitComm_1(mint端口号 - 1)
                        If lngDev < 0 Then
                            Call ShowErr(lngDev)
                            Exit Sub
                        End If
                    Else
                        lngDev = IC_InitComm_2(mint端口号 - 1, lngBoud)
                        If lngDev > 0 Then
                        ElseIf lngDev = -149 Then
                            MsgBox "当前端口已被其它程序占用！", vbInformation, gstrSysName
                            Exit Sub
                        Else
                            MsgBox "初始化端口失败！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                    
                    '判断是否已插卡
                    Call DebugTool("判断是否已插卡")
                    If mint卡类型 = 1 Then
                        intReturn = IC_Status_1(lngDev)
                        Select Case intReturn
                        Case Is < 0
                            Call ICErr(intReturn, lngDev)
                            Exit Sub
                        Case 1
                            MsgBox "请插卡！", vbInformation, gstrSysName
                            Call CloseCommon(lngDev)
                            Exit Sub
                        End Select
                    Else
                        '读取卡的状态
                        intReturn = IC_Status_2(lngDev, intCard)
                        If intReturn < 0 Then
                           MsgBox "读设备状态错误", vbInformation, gstrSysName
                           Exit Sub
                        Else
                           If intCard = 0 Then
                              MsgBox "请插卡", vbInformation, gstrSysName
                              Call CloseCommon(lngDev)
                              Exit Sub
                           Else
                              intReturn = chk_card(lngDev)
                              If intReturn < 0 Then
                                MsgBox "检测卡失败", vbInformation, gstrSysName
                                Call CloseCommon(lngDev)
                                Exit Sub
                              Else
                                Select Case intReturn
                                   Case 0
                                      MsgBox "未知卡型", vbInformation, gstrSysName
                                      Call CloseCommon(lngDev)
                                      Exit Sub
                                   Case 21
'                                      MsgBox "所插入卡为SLE4432", vbInformation, gstrSysName
'                                      Exit Sub
                                   Case Else
                                      MsgBox "所插入卡不为SLE4432", vbInformation, gstrSysName
                                      Call CloseCommon(lngDev)
                                      Exit Sub
                                End Select
                              End If
                           End If
                        End If
                    End If
                    
                    '设置卡类型
                    Call DebugTool("初始化卡类型为西门子4432/4442")
                    If mint卡类型 = 1 Then
                        intReturn = IC_InitType_1(lngDev, 16)
                        If intReturn < 0 Then
                            Call ICErr(intReturn, lngDev)
                            Exit Sub
                        End If
                    End If
                    
                    '读取数据
                    Call DebugTool("读取卡号：从59位开始，连续取6位")
                    If mint卡类型 = 1 Then
                        intReturn = IC_Read_1(lngDev, 29, 3, str卡号_IC)
                        If intReturn < 0 Then
                            Call ICErr(intReturn, lngDev)
                            Exit Sub
                        End If
                        Call DebugTool("读取医保号：从17位开始，连续取17位")
                        intReturn = IC_Read_1(lngDev, 8, 9, str医保号_IC)
                        If intReturn < 0 Then
                            Call ICErr(intReturn, lngDev)
                            Exit Sub
                        End If
                    
                        str卡号 = TruncZero(str卡号_IC)
                        str医保号 = TruncZero(str医保号_IC)
                        str医保号 = Mid(str医保号, 1, 17)
                    Else
'                        intReturn = IC_Read_2(lngDev, 0, 64, str卡号_IC)
'                        Call hex_asc%(str卡号_IC, strData, 32)
                     '曾明春(2006-01-18):由于hex_asc函数解析时会出现问题,使用手工解析
                     intReturn = srd_4442(lngDev, 8, 32, by卡号_IC(0))
                     For intnum = 0 To 9
                         If Len(CStr(hex(by卡号_IC(intnum)))) = 1 Then
                             StrHex = "0" & CStr(hex(by卡号_IC(intnum)))
                         Else
                             StrHex = CStr(hex(by卡号_IC(intnum)))
                         End If

                         strTemp = strTemp & Trim(StrHex)
                     Next
                     For intnum = 21 To 23
                         If Len(CStr(hex(by卡号_IC(intnum)))) = 1 Then
                             StrHex = "0" & CStr(hex(by卡号_IC(intnum)))
                         Else
                             StrHex = CStr(hex(by卡号_IC(intnum)))
                         End If
                         str卡号_IC2 = str卡号_IC2 & Trim(StrHex)
                     Next
                     str医保号_IC2 = Mid(strTemp, 1, 16) & Mid(strTemp, 17, 1) & Mid(strTemp, 20, 1)
                        
'                        str医保号_IC = Mid(strData, 13, 17)
'                        str卡号_IC = Mid(strData, 55, 6)
                        str卡号 = Replace(str卡号_IC2, " ", "")
                        str医保号 = Replace(str医保号_IC2, " ", "")
                    End If
                    
                    '关闭端口号
                    Call DebugTool("关闭端口号，以备下次使用")
                    Call CloseCommon(lngDev)
'                    str中心编号 = "22"
'                    '曾明春(2005-12-28):使用IC卡,需要自己判断中心编号
'                    If mintInsure = TYPE_新都 And mint地区 = 1 Then
'                       str中心编号 = "81"
'                    End If
                     str中心编号 = mintIC卡分中心
                Else
                    If mintInsure = TYPE_新都 Then
                        If 卡解析_新都(txtEdit(0).Text, str医保号, str卡号, str中心编号) = False Then Exit Sub
                    Else
                        If 卡解析_成都郊县(txtEdit(0).Text, str医保号, str卡号, str中心编号) = False Then Exit Sub
                    End If
                End If
                
                txtEdit(0).Text = str卡号
                txtEdit(1).Text = str医保号
                txtEdit(2).Text = str中心编号
                txtEdit(3).SetFocus
            End If
        ElseIf Index = 3 Then
            If mintInsure = TYPE_贵阳市 Then
                If cmdOK.Enabled = False Then
                    Call cmdTest_Click
                Else
                    Call cmdOK_Click
                End If
            Else
                If cmdOK.Enabled Then Call cmdOK_Click
            End If
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub ShowErr(ByVal lngError As Long)
    Dim strMsg As String
    
    lngError = Abs(lngError)
    Select Case lngError
    Case &H80
        strMsg = "读错误！"
    Case &H81
        strMsg = "写错误！"
    Case &H82
        strMsg = "通讯错误！"
    Case &H83
        strMsg = "密码错误！"
    Case &H84
        strMsg = "通讯超时！"
    Case &H85
        strMsg = "校验和错误！"
    Case &H86
        strMsg = "请插卡！"
    Case &H87
        strMsg = "函数参数格式错误！"
    Case Else
        strMsg = "未知错误！"
    End Select
    MsgBox strMsg & "错误号：" & lngError, vbInformation, gstrSysName
End Sub

Private Sub CloseCommon(ByVal lngDev As Long)
    Dim intReturn As Integer
    
    If mint卡类型 = 0 Then Exit Sub
    
    If mint卡类型 = 1 Then
        intReturn = IC_ExitComm_1(lngDev)
    Else
        intReturn = IC_Down_2(lngDev)
        intReturn = ic_exit%(lngDev)
    End If
    If intReturn < 0 Then
        MsgBox "    关闭端口时发生未知错误，你只能通过关闭操作系统来关闭端口！" & vbCrLf & _
                "由于端口无法关闭，本工作站将无法对卡完成读或写操作", vbInformation, gstrSysName
    End If
End Sub

Private Sub ICErr(ByVal lngErr As Long, ByVal lngDev As Long)
    If lngErr < 0 Then
        Call ShowErr(lngErr)
        Call CloseCommon(lngDev)
    End If
End Sub
