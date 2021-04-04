VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Begin VB.Form frmFeeGroupChargeAndBillTotal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "收款及票据汇总"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9930
   Icon            =   "frmFeeGroupChargeAndBillTotal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9930
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9855
      TabIndex        =   0
      Top             =   3360
      Width           =   9855
      Begin VB.TextBox txtDate 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   98
         Width           =   2760
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   98
         Width           =   1800
      End
      Begin VB.TextBox txtNO 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   98
         Width           =   1920
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "轧帐日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6120
         TabIndex        =   5
         Top             =   150
         Width           =   840
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "收费员"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3240
         TabIndex        =   3
         Top             =   150
         Width           =   630
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         Caption         =   "轧帐单号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   840
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFeeGroupChargeAndBillTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long, mintType As Integer, mstrNO As String, mstrDate As String, mstrName As String   '传入的收款记录信息
Private mlngModule As Long, mstrPrivs As String
Private mobjChargeBill As New clsChargeBill, mfrmChargeBill As Form '收款信息与票据对象

Public Enum BillType
    BT_收费员扎帐 = 0
    BT_小组收款 = 1
End Enum

Public Sub ShowMe(frmMain As Object, btIn As BillType, ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngID As Long, _
                  ByVal strNO As String, ByVal strDate As String, ByVal strName As String)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:调用接口
    '入参:frmMain-调用窗体
    '     btIn-单据类型
    '     lngModule-模块号
    '     strPrivs-权限串
    '     lngID-单据ID
    '     strNO-单据号码
    '     strDate-收费/轧帐日期
    '     strName-收费员
    '编制:刘尔旋
    '日期:2013-09-29
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mintType = btIn
    mlngID = lngID
    mstrDate = strDate
    mstrNO = strNO
    mstrName = strName
    mstrPrivs = strPrivs
    mlngModule = lngModule
    Me.Show vbModal, frmMain
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionAttaching Then Cancel = True
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Set mfrmChargeBill = mobjChargeBill.GetChargeAndBillTotalForm
    mobjChargeBill.SetFontSize lblNO.Font.Size
    Call SetDockingPanel
    If mintType = 1 Then
        lblNO.Caption = "收款单号"
        lblDate.Caption = "收款日期"
        mobjChargeBill.LoadChargeAndBillTotalData Me, mlngModule, mstrPrivs, EM_小组收款, mlngID
    End If
    If mintType = 0 Then
        lblNO.Caption = "轧帐单号"
        lblDate.Caption = "轧帐日期"
        mobjChargeBill.LoadChargeAndBillTotalData Me, mlngModule, mstrPrivs, EM_收费员轧帐, mlngID
    End If
    
    txtNO.Text = mstrNO
    txtDate.Text = mstrDate
    txtName.Text = mstrName
    
End Sub

Private Sub SetDockingPanel()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:创建DOCKINGPANEL控件
    '编制:刘尔旋
    '日期:2013-09-29
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim objPanel As Pane
    On Error GoTo errHandle
    With dkpMain
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(1, 2000, 60, DockTopOf)
        objPanel.Handle = picInfo.hwnd
        objPanel.Title = "基本信息"
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.MinTrackSize.Height = 25
        objPanel.MaxTrackSize.Height = 25
        Set objPanel = .CreatePane(2, 2000, 1000, DockBottomOf, objPanel)
        objPanel.Handle = mfrmChargeBill.hwnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.Title = "收费票据使用信息"
        Set .PaintManager.CaptionFont = lblNO.Font
        .Options.HideClient = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmChargeBill Is Nothing Then Unload mfrmChargeBill
    Set mfrmChargeBill = Nothing
    If Not mobjChargeBill Is Nothing Then Set mobjChargeBill = Nothing
End Sub
