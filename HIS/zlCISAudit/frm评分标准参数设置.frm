VERSION 5.00
Begin VB.Form frm评分标准参数设置 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "基本参数设置"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frm评分标准参数设置.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin zl9CISAudit.tipPopup tipPopup1 
      Height          =   420
      Left            =   972
      Top             =   1170
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   162
      TabIndex        =   5
      Top             =   1695
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2217
      TabIndex        =   2
      Top             =   1695
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3447
      TabIndex        =   3
      Top             =   1695
      Width           =   1100
   End
   Begin VB.Frame fra1 
      Caption         =   "系统参数"
      Height          =   1380
      Left            =   147
      TabIndex        =   4
      Top             =   135
      Width           =   4380
      Begin VB.CheckBox chk参数 
         Caption         =   "要求病案首页编目后再评分(&P)"
         Height          =   210
         Index           =   91
         Left            =   975
         TabIndex        =   1
         Top             =   810
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chk参数 
         Caption         =   "评分等级自动写入病案主页(&Y)"
         Height          =   210
         Index           =   90
         Left            =   975
         TabIndex        =   0
         Top             =   435
         Width           =   3015
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   315
         Picture         =   "frm评分标准参数设置.frx":000C
         Top             =   450
         Width           =   480
      End
   End
End
Attribute VB_Name = "frm评分标准参数设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'==============================================================================
'=功能： 参数控制添加描述'显示Tips
'==============================================================================
Private Sub chk参数_GotFocus(Index As Integer)
    On Error GoTo errH
    If Index = 90 Then
        ShowTips fra1, chk参数(90), "是否在病案首页的等级为空时，将评分结果等级自动写入病案首页。默认为否。", "评分等级自动写入病案首页"
    ElseIf Index = 91 Then
        ShowTips fra1, chk参数(91), "是否要求病案首页编目后才能评分。默认为是。", "要求病案首页编目后再评分"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 点击取消退出
'==============================================================================
Private Sub cmdCancel_Click()
    On Error GoTo errH
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 点击帮助
'==============================================================================
Private Sub cmdHelp_Click()
    On Error GoTo errH
    ShowHelp App.ProductName, Me.hWnd, Me.Name, 3
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 点击确定保存参数
'==============================================================================
Private Sub cmdOK_Click()
    On Error GoTo errH
    If Save参数() = False Then Exit Sub
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 窗口控件初始化
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo errH
    InitCommonControls
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 窗口初始化
'==============================================================================
Private Sub Form_Load()
    chk参数(90).Value = zlDatabase.GetPara(90, glngSys)
    chk参数(91).Value = zlDatabase.GetPara(91, glngSys)
End Sub

'==============================================================================
'=功能:保存编辑的内容到各个与系统参数相关的表中
'=返回值:成功返回True,否则为False
'==============================================================================
Private Function Save参数() As Boolean
    Dim i           As Integer
    
    On Error GoTo errH
    
    Save参数 = False
    gcnOracle.BeginTrans
    For i = 90 To 91
        Call zlDatabase.SetPara(i, IIf(chk参数(i).Value = 1, 1, 0), ParamInfo.系统号, 0)
    Next
    gcnOracle.CommitTrans
    Save参数 = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能:控件通用Tips显示
'==============================================================================
Private Sub ShowTips(ctl0 As Control, ctl As Control, str内容 As String, Optional str标题 As String = "提示信息", Optional lng时间 As Long = 3000)
    Dim X           As Single
    Dim Y           As Single
    
    On Error GoTo errH
    
    X = (ctl.Left + ctl.Width / 2 + ctl0.Left) / Screen.TwipsPerPixelX
    Y = (ctl.Top + ctl.Height + ctl0.Top) / Screen.TwipsPerPixelY
    If Len(str内容) > 0 Then
        tipPopup1.Hide
        tipPopup1.StandardIcon = IDI_INFORMATION
        tipPopup1.ShowCloseButton = True
        tipPopup1.TimeOut = lng时间
        tipPopup1.Title = str标题
        tipPopup1.Text = str内容
        tipPopup1.Show Me.hWnd, X, Y
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
