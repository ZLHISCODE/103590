VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotifySet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "frmNotifySet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4935
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdSetup 
      Caption         =   "设置(&S)"
      Height          =   350
      Left            =   165
      TabIndex        =   4
      Top             =   2295
      Width           =   1100
   End
   Begin VB.CommandButton cmdPriv 
      Caption         =   "预览(&O)"
      Height          =   350
      Left            =   1335
      TabIndex        =   5
      Top             =   2295
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   2490
      TabIndex        =   6
      Top             =   2295
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3645
      TabIndex        =   7
      Top             =   2295
      Width           =   1100
   End
   Begin VB.Frame fra条件 
      Caption         =   "条件设置"
      Height          =   2010
      Left            =   195
      TabIndex        =   8
      Top             =   120
      Width           =   4470
      Begin VB.TextBox txt催款金额 
         Height          =   300
         Left            =   1215
         TabIndex        =   3
         Top             =   915
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Left            =   1230
         TabIndex        =   1
         Top             =   390
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   126156803
         CurrentDate     =   36576
      End
      Begin VB.Label lblEdit 
         Caption         =   "催款金额"
         Height          =   180
         Left            =   375
         TabIndex        =   2
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "通知单将打印病人在指定截止日期所在期间内的费用欠款情况！"
         ForeColor       =   &H00800000&
         Height          =   450
         Left            =   690
         TabIndex        =   9
         Top             =   1365
         Width           =   3465
      End
      Begin VB.Label lbl截止日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "截止日期"
         Height          =   180
         Left            =   420
         TabIndex        =   0
         Top             =   465
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmNotifySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mbytType As Byte '1表示打印,2表示预览
Private mblnFirst As Boolean
Private mstrPrivs As String
Private mblncmdPriv As Boolean
Private mblnOk As Boolean, mstr载止日期 As String, mdbl催款金额 As Double
Public Function ShowSet(ByVal frmMain As Form, strPrivs As String, ByVal blncmdPriv As Boolean, ByRef bytType As Byte, ByRef str载止日期 As String, ByRef dbl催款金额 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示催款单条件窗体设置
    '入参:frmMain-调用的父窗口
    '     blncmdPriv-是否显示预览按钮
    '出参:bytType-0 表示取消 1表示打印,2表示预览
    '     str载止日期
    '     dbl催款金额
    '返回:
    '编制:刘兴洪
    '日期:2010-01-20 11:58:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytType = 0: mblnOk = False: mdbl催款金额 = 0: mstr载止日期 = str载止日期: mstrPrivs = strPrivs: mblncmdPriv = blncmdPriv
    Me.Show 1, frmMain
    str载止日期 = mstr载止日期: dbl催款金额 = mdbl催款金额
    ShowSet = mblnOk: bytType = mbytType
End Function
Private Sub cmdCancel_Click()
    mblnOk = False:
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    mstr载止日期 = Format(dtp.Value, "yyyy-mm-dd"): mdbl催款金额 = Val(txt催款金额.Text)
    mbytType = 1: mblnOk = True:  Unload Me
End Sub

Private Sub cmdPriv_Click()
    mstr载止日期 = Format(dtp.Value, "yyyy-mm-dd"): mdbl催款金额 = Val(txt催款金额.Text)
    mbytType = 2: mblnOk = True: Unload Me:
End Sub

Private Sub cmdSetup_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1139_3", Me
End Sub

Private Sub dtp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mstr载止日期 <> "" And IsDate(mstr载止日期) Then dtp.Value = CDate(mstr载止日期)
End Sub

Private Sub Form_Load()
    mblnFirst = True: mbytType = 0
    txt催款金额.Text = zlDatabase.GetPara("催款金额", glngSys, 1139, "", Array(txt催款金额, lblEdit), InStr(1, mstrPrivs, ";参数设置;") > 0)
    dtp.Value = DateAdd("d", -1, zlDatabase.Currentdate)
    cmdPriv.Visible = mblncmdPriv
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
  Call zlDatabase.SetPara("催款金额", Val(txt催款金额.Text), glngSys, 1139, InStr(1, mstrPrivs, ";参数设置;") > 0)
End Sub

Private Sub txt催款金额_GotFocus()
    zlControl.TxtSelAll txt催款金额
End Sub

Private Sub txt催款金额_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub txt催款金额_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt催款金额, KeyAscii, m负金额式
End Sub

Private Sub txt催款金额_LostFocus()
    txt催款金额.Text = Format(Val(txt催款金额.Text), "0.00")
End Sub
