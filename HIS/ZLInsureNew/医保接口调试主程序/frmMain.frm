VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保接口测试程序"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10455
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd功能 
      Caption         =   "费用审核管理(&1)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   22
      Left            =   480
      TabIndex        =   32
      Tag             =   "1620;zl9I_Manager;800"
      Top             =   5130
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "结算数据管理(&2)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   23
      Left            =   3000
      TabIndex        =   31
      Tag             =   "1621;zl9I_Manager;800"
      Top             =   5130
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "出院数据审核(&3)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   24
      Left            =   5520
      TabIndex        =   30
      Tag             =   "1622;zl9I_Manager;800"
      Top             =   5130
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "保险病种(&4)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   29
      Left            =   7950
      TabIndex        =   29
      Tag             =   "1623;zl9I_Manager;800"
      Top             =   5130
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "票据使用监控(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   31
      Left            =   5520
      TabIndex        =   28
      Tag             =   "1501;zl9CashBill;100"
      Top             =   2865
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "指标超限管理(&7)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   30
      Left            =   5520
      TabIndex        =   27
      Tag             =   "1626;zl9I_Manager;800"
      Top             =   5835
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "特殊项目管理(&5)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   28
      Left            =   480
      TabIndex        =   26
      Tag             =   "1624;zl9I_Manager;800"
      Top             =   5835
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "参数设置(&9)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   27
      Left            =   480
      TabIndex        =   25
      Tag             =   "1620;zl9I_Manager;800"
      Top             =   6555
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "住院标准限额分科管理(&8)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   26
      Left            =   7950
      TabIndex        =   24
      Tag             =   "1620;zl9I_Manager;800"
      Top             =   5835
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "考核指标管理(&6)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   25
      Left            =   2985
      TabIndex        =   23
      Tag             =   "1625;zl9I_Manager;800"
      Top             =   5835
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "门诊分诊(&D)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   21
      Left            =   7950
      TabIndex        =   22
      Tag             =   "1113;zl9RegEvent;100"
      Top             =   810
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "医技记帐(&H)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   20
      Left            =   7950
      TabIndex        =   21
      Tag             =   "1135;zl9InExse;100"
      Top             =   1455
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "医保工具(&Z)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   18
      Left            =   7950
      TabIndex        =   20
      Tag             =   "1607;zl9insure;100"
      Top             =   4320
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "住院护士(&J)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   19
      Left            =   3000
      TabIndex        =   19
      Tag             =   "1262;zl9CisJob;100"
      Top             =   2145
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "医技站(&G)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   17
      Left            =   5520
      TabIndex        =   18
      Tag             =   "1263;zl9CisJob;100"
      Top             =   1455
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "保险病种(&Y)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   16
      Left            =   5520
      TabIndex        =   17
      Tag             =   "1606;zl9insure;100"
      Top             =   4320
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "医保结算(&X)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   15
      Left            =   3000
      TabIndex        =   15
      Tag             =   "1605;zl9insure;100"
      Top             =   4320
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "医保大类(&U)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   12
      Left            =   5520
      TabIndex        =   12
      Tag             =   "1602;zl9insure;100"
      Top             =   3630
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "结算规则(&T)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   11
      Left            =   3000
      TabIndex        =   11
      Tag             =   "1601;zl9insure;100"
      Top             =   3630
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "医保项目(&V)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   13
      Left            =   7950
      TabIndex        =   13
      Tag             =   "1603;zl9insure;100"
      Top             =   3630
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "保险帐户(&W)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   14
      Left            =   480
      TabIndex        =   14
      Tag             =   "1604;zl9insure;100"
      Top             =   4320
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "保险类别(&S)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   10
      Left            =   480
      TabIndex        =   10
      Tag             =   "1600;zl9insure;100"
      Top             =   3630
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "住院医生(&I)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   9
      Left            =   480
      TabIndex        =   9
      Tag             =   "1261;zl9CisJob;100"
      Top             =   2145
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "门诊医生(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   8
      Left            =   5520
      TabIndex        =   8
      Tag             =   "1260;zl9CisJob;100"
      Top             =   810
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "费用查询(&M)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Tag             =   "1139;zl9InExse;100"
      Top             =   2865
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "分散记帐(&L)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   6
      Left            =   7950
      TabIndex        =   6
      Tag             =   "1134;zl9InExse;100"
      Top             =   2145
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "入出管理(&F)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Tag             =   "1132;zl9Inpatient;100"
      Top             =   1455
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "入院登记(&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Tag             =   "1131;zl9Inpatient;100"
      Top             =   1455
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "住院结帐(&N)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   5
      Left            =   3000
      TabIndex        =   5
      Tag             =   "1137;zl9InExse;100"
      Top             =   2865
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "住院记帐(&K)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   4
      Left            =   5520
      TabIndex        =   4
      Tag             =   "1133;zl9InExse;100"
      Top             =   2145
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "门诊收费(&B)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Tag             =   "1121;zl9OutExse;100"
      Top             =   810
      Width           =   2265
   End
   Begin VB.CommandButton cmd功能 
      Caption         =   "门诊挂号(&A)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Tag             =   "1111;zl9RegEvent;100"
      Top             =   810
      Width           =   2265
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      X1              =   -180
      X2              =   43650
      Y1              =   4995
      Y2              =   4995
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   -180
      X2              =   43650
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      X1              =   -345
      X2              =   42000
      Y1              =   3555
      Y2              =   3555
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   -345
      X2              =   42000
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   360
      Picture         =   "frmMain.frx":1CFA
      Top             =   210
      Width           =   240
   End
   Begin VB.Label lblNote 
      Caption         =   "    警告：该程序仅允许中联公司人员使用，且只能应用于测试医保接口，任何人不得用于非法用途"
      Height          =   240
      Index           =   0
      Left            =   1140
      TabIndex        =   16
      Top             =   270
      Width           =   8985
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum 功能清单
    门诊挂号 = 0
    门诊收费
    入院登记
    入出院管理
    住院记帐
    住院结算
End Enum
Const mstrServer As String = "orcl"               '主机串
Const mstrUser As String = "ZLHIS"                  '用户名
Const mstrPass As String = "HIS"                  '密码
Private mobjCommon As Object
Private mcnOracle As New ADODB.Connection           '连接

Private mobjTest As Object

Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private Const 时间间隔 As Long = 500 '毫秒
Dim TW As SYSTEMTIME

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Static 上次时间 As Long
    Static 上次按键 As Integer
    Dim i           As Integer
    '读当前时间
    Call GetLocalTime(TW)
    '判断时差
    If TW.wSecond * 1000& + TW.wMilliseconds - 上次时间 < 时间间隔 And KeyCode = 上次按键 Then
        For i = 0 To cmd功能.Count - 1
            If InStr(1, cmd功能(i).Caption, UCase(Chr(KeyCode))) > 0 Then
                cmd功能_Click (i)
            End If
        Next
    End If
    上次时间 = TW.wSecond * 1000& + TW.wMilliseconds
    上次按键 = KeyCode
End Sub

Private Sub cmd功能_Click(Index As Integer)
    Dim strClass As String
    Dim lngModul As Long
    Dim lngSys As Long
    
    If mcnOracle Is Nothing Then Exit Sub
    If mcnOracle.State = 0 Then Exit Sub
    If mobjCommon Is Nothing Then
        MsgBox "未初始化公共部件！", vbInformation
        Exit Sub
    End If
'    mobjCommon.gstrNodeNo = "3"
    lngSys = Val(Split(cmd功能(Index).Tag, ";")(2))
    lngModul = Val(Split(cmd功能(Index).Tag, ";")(0))
    strClass = Split(cmd功能(Index).Tag, ";")(1)
    strClass = strClass & ".cls" & Mid(strClass, 4)
    
'    On Error Resume Next
    If Not mobjTest Is Nothing Then
        Call mobjTest.CloseWindows
        Set mobjTest = Nothing
    End If
    Err = 0
    Set mobjTest = CreateObject(strClass)
    If Err <> 0 Then
        MsgBox "无法创建该部件，请确认是否已安装！", vbInformation
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    Me.WindowState = 1
    
    Call mobjTest.CodeMan(lngSys, lngModul, mcnOracle, Nothing, mstrUser)
    Exit Sub
ErrHand:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub Form_Load()
    '打开数据库连接
    With mcnOracle
        If .State = 1 Then .Close
        .Provider = "MsDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & mstrServer, mstrUser, mstrPass
    End With
    
    On Error Resume Next
    Set mobjCommon = CreateObject("ZL9ComLib.clsComLib")
    Call mobjCommon.InitCommon(mcnOracle)
    Call mobjCommon.SetDbUser(mstrUser)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not mobjTest Is Nothing Then
        Call mobjTest.CloseWindows
        Set mobjTest = Nothing
    End If
End Sub

