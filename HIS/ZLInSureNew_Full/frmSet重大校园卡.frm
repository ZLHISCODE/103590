VERSION 5.00
Begin VB.Form frmSet重大校园卡 
   AutoRedraw      =   -1  'True
   Caption         =   "设置"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   ControlBox      =   0   'False
   Icon            =   "frmSet重大校园卡.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   3735
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Txt限额 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1500
      TabIndex        =   1
      Text            =   "2000"
      Top             =   1380
      Width           =   930
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   7575
      TabIndex        =   7
      Top             =   855
      Width           =   7575
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1395
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "1"
         Top             =   75
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "当前串口(&D)"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   9
         Top             =   135
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "号串口"
         Height          =   180
         Index           =   4
         Left            =   1800
         TabIndex        =   8
         Top             =   135
         Width           =   540
      End
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -165
      TabIndex        =   6
      Top             =   1905
      Width           =   7755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2595
      TabIndex        =   3
      Top             =   2070
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1335
      TabIndex        =   2
      Top             =   2070
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   30
      TabIndex        =   5
      Top             =   690
      Width           =   7665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "元"
      Height          =   180
      Left            =   2505
      TabIndex        =   11
      Top             =   1425
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "每次交易限额"
      Height          =   180
      Left            =   390
      TabIndex        =   10
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmSet重大校园卡.frx":000C
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "设置设备的串口号和交易限额."
      Height          =   315
      Left            =   540
      TabIndex        =   4
      Top             =   390
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet重大校园卡"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlng医保中心 As Long
Private mlng险类 As Long

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    
    If Trim(txtEdit) = "" Then Exit Sub
    SaveRegInFor g公共模块, "操作", "端口号", Me.txtEdit
    gintComport_重大校园卡 = Val(txtEdit)
    
    '删除已经数据
    On Error GoTo errHand
    gstrSQL = "zl_保险参数_Update(" & mlng险类 & ",NULL,'交易限额' ,'" & Val(Txt限额.Text) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mblnReturn = True
    Unload Me
    Exit Sub
errHand:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    mblnReturn = False
    
    Call GetRegInFor(g公共模块, "操作", "端口号", strReg)
    
    If Val(strReg) = 0 Then
        txtEdit.Text = 0
    Else
        txtEdit.Text = Val(strReg)
    End If
    
     gstrSQL = "Select * From 保险参数 where 参数名 ='交易限额' and 险类=" & mlng险类
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If Not rsTemp.EOF Then
            Txt限额.Text = Format(Val(NVL(rsTemp!参数值)), "####0.00;-####0.00; ;")
    End If
End Sub

Public Function ShowME(ByVal lng险类 As Long, ByVal lng医保中心 As Long) As Boolean
    mlng医保中心 = lng医保中心
    mlng险类 = lng险类
    Me.Show 1
    ShowME = mblnReturn
End Function
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
        zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m数字式
End Sub



Private Sub Txt限额_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt限额, KeyAscii, m金额式
End Sub
