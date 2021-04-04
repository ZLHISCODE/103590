VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholMaterials_Change 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "换缸"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4155
   Icon            =   "frmPatholMaterials_Change.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cbxTimeLen 
      Height          =   300
      ItemData        =   "frmPatholMaterials_Change.frx":179A
      Left            =   1680
      List            =   "frmPatholMaterials_Change.frx":17B0
      TabIndex        =   1
      Text            =   "12"
      Top             =   720
      Width           =   1785
   End
   Begin VB.CommandButton cmdReject_Sure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   960
      TabIndex        =   2
      Top             =   1275
      Width           =   1215
   End
   Begin VB.CommandButton cmdReject_Cancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   2280
      TabIndex        =   3
      Top             =   1275
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpStartTime 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   155058179
      CurrentDate     =   40646.4399652778
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   6
      Top             =   300
      Width           =   900
   End
   Begin VB.Label labSubmitDoctor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "所需时长(小时)："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   780
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   780
      Width           =   255
   End
End
Attribute VB_Name = "frmPatholMaterials_Change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mblnIsSure As Boolean




Property Get IsSure() As Boolean
    IsSure = mblnIsSure
End Property


Property Get StartTime() As Date
    StartTime = dtpStartTime.value
End Property


Property Get TimeLen() As Double
    TimeLen = Val(cbxTimeLen.Text)
End Property




Public Sub ShowChangeWindow(owner As Form)
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
    
    dtpStartTime.value = zlDatabase.Currentdate
    mblnIsSure = False
    
    Call Me.Show(1, owner)
End Sub





Private Sub cmdReject_Cancel_Click()
    mblnIsSure = False
    
    Me.Hide
End Sub

Private Sub cmdReject_Sure_Click()
    If Val(cbxTimeLen.Text) <= 0 Then
        Call MsgBoxD(Me, "请录入有效的时间长度。", vbOKOnly, Me.Caption)
        Call cbxTimeLen.SetFocus
        
        Exit Sub
    End If
    
    mblnIsSure = True
    
    Me.Hide
End Sub

Private Sub Form_Initialize()
    mblnIsSure = False
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub
