VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#10.0#0"; "zlIDKind.ocx"
Begin VB.Form frmSendCardAndDepositErrEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "医疗卡及预交异常管理"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15060
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendCardAndDepositErrEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   15060
   StartUpPosition =   1  '所有者中心
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   3765
      Left            =   30
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1605
      Width           =   14970
      _Version        =   589884
      _ExtentX        =   26405
      _ExtentY        =   6641
      _StockProps     =   64
   End
   Begin VB.PictureBox picTotal 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   45
      ScaleHeight     =   495
      ScaleWidth      =   4230
      TabIndex        =   21
      Top             =   5520
      Width           =   4230
      Begin zlIDKind.ucQRCodePayButton btQRCode 
         Height          =   465
         Left            =   3480
         TabIndex        =   22
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   820
         Appearance      =   1
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "本次支付合计:0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   30
         TabIndex        =   23
         Top             =   150
         Width           =   3315
      End
   End
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   -135
      ScaleHeight     =   780
      ScaleWidth      =   16005
      TabIndex        =   20
      Top             =   6090
      Width           =   16005
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   13230
         TabIndex        =   16
         Top             =   165
         Width           =   1470
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   11625
         TabIndex        =   15
         Top             =   195
         Width           =   1470
      End
      Begin VB.Line lnW 
         BorderColor     =   &H00FFFFFF&
         X1              =   90
         X2              =   14910
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line lnDown 
         BorderColor     =   &H80000000&
         X1              =   -105
         X2              =   15700
         Y1              =   45
         Y2              =   45
      End
   End
   Begin VB.PictureBox picPati 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   -15
      ScaleHeight     =   1425
      ScaleWidth      =   15240
      TabIndex        =   17
      Top             =   120
      Width           =   15240
      Begin VB.Frame Frame 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   -45
         TabIndex        =   19
         Top             =   690
         Width           =   15495
      End
      Begin VB.Frame fraInfo 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   -105
         TabIndex        =   18
         Top             =   -90
         Width           =   15570
         Begin VB.TextBox txt年龄 
            BackColor       =   &H80000004&
            Height          =   360
            Left            =   5625
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   765
         End
         Begin VB.TextBox txt姓名 
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   240
            Width           =   1545
         End
         Begin VB.TextBox txt床号 
            BackColor       =   &H80000004&
            Height          =   360
            Left            =   7545
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   240
            Width           =   1110
         End
         Begin VB.TextBox txt性别 
            BackColor       =   &H80000004&
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   3615
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "热键：F11"
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txt付款方式 
            BackColor       =   &H80000004&
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "热键：F11"
            Top             =   870
            Width           =   2085
         End
         Begin VB.TextBox txt费别 
            BackColor       =   &H80000004&
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "热键：F11"
            Top             =   870
            Width           =   1545
         End
         Begin VB.TextBox txt住院号 
            BackColor       =   &H80000004&
            Height          =   360
            Left            =   10290
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   210
            Width           =   2160
         End
         Begin VB.Label lbl费别 
            AutoSize        =   -1  'True
            Caption         =   "费别"
            Height          =   240
            Left            =   300
            TabIndex        =   10
            Top             =   930
            Width           =   480
         End
         Begin VB.Label lbl年龄 
            AutoSize        =   -1  'True
            Caption         =   "年龄"
            Height          =   240
            Left            =   5160
            TabIndex        =   4
            Top             =   300
            Width           =   480
         End
         Begin VB.Label lbl性别 
            AutoSize        =   -1  'True
            Caption         =   "性别"
            Height          =   240
            Left            =   3075
            TabIndex        =   2
            Top             =   300
            Width           =   480
         End
         Begin VB.Label lbl姓名 
            AutoSize        =   -1  'True
            Caption         =   "病人"
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   315
            TabIndex        =   0
            Top             =   300
            Width           =   480
         End
         Begin VB.Label lbl床号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "床号"
            Height          =   240
            Left            =   7005
            TabIndex        =   6
            Top             =   315
            Width           =   480
         End
         Begin VB.Label lbl付款方式 
            AutoSize        =   -1  'True
            Caption         =   "付款方式"
            Height          =   240
            Left            =   2850
            TabIndex        =   12
            Top             =   930
            Width           =   960
         End
         Begin VB.Label lbl住院号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院号"
            Height          =   240
            Left            =   9570
            TabIndex        =   8
            Top             =   270
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "frmSendCardAndDepositErrEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------
'接口入参变量
Private mfrmain As Object
Private mint操作类型 As Integer  '0-增加;1-异常重收;2-异常作废
Private mlng异常ID As Long
Private mlngMoudle As Long
Private WithEvents mfrmSendCardAndDeposit As frmSendCardAndDeposit
Attribute mfrmSendCardAndDeposit.VB_VarHelpID = -1
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'----------------------------------------------------------------------------------------------
Private Const Pane_Pati = 10
Private Const Pane_SendCard = 11
Private Const Pane_Down = 12
Attribute Pane_Down.VB_VarHelpID = -1
Private mbln门诊预交 As Boolean
Private mbln住院预交 As Boolean
Private mbln医保缴预交 As Boolean
Private mbln发卡 As Boolean
Private mlngCardTypeID As Long
Private mblnUnload  As Boolean
Private mintSuccess As Integer
Private mblnFirst As Boolean
Private mbyt场景 As Byte    '
Private mobjPati As clsPatientInfo

Public Function zlShowWindow(ByVal frmMain As Object, ByVal int操作类型 As Integer, ByVal lng异常ID As Long, ByVal lngMoudle As Long, _
    Optional ByVal int场景 As Byte = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:异常重收或作废
    '入参:int操作类型-0-增加;1-异常重收;2-异常作废
    '     int场景:1-医疗卡发卡;2-病人信息登记;3-病人入院 登记;4-预约挂号接收
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-28 10:11:04
    '调用者:
    '    入院及信息登记
    '说明:场景为1,4的暂不支持
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mintSuccess = 0: mint操作类型 = int操作类型: mlng异常ID = lng异常ID: mlngMoudle = lngMoudle
    mbyt场景 = int场景
    
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlShowWindow = mintSuccess > 0
End Function


Private Function InitFace() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-28 10:21:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
     
    Call InitPage
    If mfrmSendCardAndDeposit.zlInit(Me, mlngMoudle, True, True, mlngCardTypeID, True, True, mbln医保缴预交, btQRCode, lblTotal, False, , True) = False Then mblnUnload = False: Exit Function
    InitFace = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2011-06-28 15:22:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:
     
 
    Set mfrmSendCardAndDeposit = New frmSendCardAndDeposit
    Load mfrmSendCardAndDeposit
    
    Set objItem = tbPage.InsertItem(1, "发卡及缴预交", mfrmSendCardAndDeposit.hWnd, 0)
    objItem.Tag = 1
    With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
 
Private Function ReadCardAndDepositData(ByVal lng异常ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取卡费及预交数据
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-28 10:20:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    If mfrmSendCardAndDeposit.zlReadCardAndDepositErrData(mint操作类型, lng异常ID, mobjPati) = False Then
        EnableWindow mfrmSendCardAndDeposit.hWnd, True
        Exit Function
    End If
    EnableWindow mfrmSendCardAndDeposit.hWnd, True
    
    If mint操作类型 = 0 Then ReadCardAndDepositData = True: Exit Function
    If mobjPati Is Nothing Then Exit Function
    
    '加载病人信息
    txt姓名.Text = mobjPati.姓名
    txt性别.Text = mobjPati.性别
    txt年龄.Text = mobjPati.年龄
    txt床号.Text = mobjPati.床号
    txt住院号.Text = mobjPati.住院号
    txt费别.Text = mobjPati.费别
    txt付款方式.Text = mobjPati.医疗付款方式
    ReadCardAndDepositData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '确定保存
    If mfrmSendCardAndDeposit.zlSaveDataBeforCheckIsValid(False, mobjPati, False) = False Then
        EnableWindow mfrmSendCardAndDeposit.hWnd, True
        Exit Sub
    End If
    If mfrmSendCardAndDeposit.zlSaveData(False, mobjPati) = False Then
        EnableWindow mfrmSendCardAndDeposit.hWnd, True
        Exit Sub
    End If
    EnableWindow mfrmSendCardAndDeposit.hWnd, True
    mfrmSendCardAndDeposit.zlSaveDataAfter '保存后处理
    mintSuccess = mintSuccess + 1
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
End Sub

Private Sub Form_Load()
    mblnFirst = True
    cmdOk.Caption = IIf(mint操作类型 = 2, "作废(&D)", "确定(&O)")
    Me.Caption = IIf(mint操作类型 = 2, "医疗卡及预交异常作废", "医疗卡及预交异常重收")
    mblnUnload = Not InitFace
    If mblnUnload = False Then mblnUnload = Not ReadCardAndDepositData(mlng异常ID)
End Sub

 
Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    Set mfrmSendCardAndDeposit = Nothing
End Sub

 

Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbTab
End Sub
