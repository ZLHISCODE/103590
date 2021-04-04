VERSION 5.00
Begin VB.Form frmFeeVrerfyParaSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数设置"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   3090
      TabIndex        =   1
      Top             =   2730
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   4695
      TabIndex        =   5
      Top             =   2730
      Width           =   1500
   End
   Begin VB.CheckBox chk审核 
      Caption         =   "门诊转住院必须先审核(&V)"
      Height          =   270
      Left            =   1155
      TabIndex        =   0
      Top             =   1710
      Width           =   3405
   End
   Begin VB.Frame fraSplit 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   1
      Left            =   -45
      TabIndex        =   3
      Top             =   2400
      Width           =   8925
   End
   Begin VB.Frame fraSplit 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   1275
      Width           =   8925
   End
   Begin VB.Label lblTittle 
      Caption         =   $"frmFeeVrerfyParaSet.frx":0000
      Height          =   945
      Left            =   990
      TabIndex        =   4
      Top             =   315
      Width           =   5205
   End
   Begin VB.Image imgPit 
      Height          =   720
      Left            =   105
      Picture         =   "frmFeeVrerfyParaSet.frx":0093
      Top             =   435
      Width           =   720
   End
End
Attribute VB_Name = "frmFeeVrerfyParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************
'**该参数已废弃,已经放在了费用部分的公共部分进行设置

Private mlngModule As String, mstrPrivs As String, mblnOk As Boolean
Public Function ShowMe(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置相关参数的入口
    '返回:参数设置成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-09 11:35:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOk = False: mlngModule = lngModule: mstrPrivs = strPrivs
    Me.Show 1, frmMain
    ShowMe = mblnOk
End Function
Private Sub LoadPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载参数
    '编制:刘兴洪
    '日期:2011-02-09 11:36:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    chk审核.Value = IIf(Val(zlDatabase.GetPara("门诊转住院先审核", glngSys, mlngModule, 0, Array(chk审核), InStr(1, mstrPrivs, ";参数设置;") > 0)) = 1, 1, 0)
End Sub
Private Sub SavePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存参数
    '编制:刘兴洪
    '日期:2011-02-09 11:36:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call zlDatabase.SetPara("门诊转住院先审核", IIf(chk审核.Value = 1, 1, 0), glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SavePara
    Unload Me
    mblnOk = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Call LoadPara
End Sub
