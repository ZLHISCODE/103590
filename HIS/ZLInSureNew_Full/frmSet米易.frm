VERSION 5.00
Begin VB.Form frmSet米易 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmSet米易.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3210
      TabIndex        =   5
      Top             =   1050
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1950
      TabIndex        =   4
      Top             =   1050
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   4635
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1545
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "号串口"
      Height          =   180
      Index           =   4
      Left            =   1950
      TabIndex        =   2
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "当前串口(&D)"
      Height          =   180
      Index           =   3
      Left            =   450
      TabIndex        =   1
      Top             =   300
      Width           =   990
   End
End
Attribute VB_Name = "frmSet米易"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlng险类 As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtEdit) = "" Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo errHand
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",NULL,'端口号','" & txtEdit.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    gintComPort = txtEdit.Text
    mblnReturn = True
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mblnReturn = False
    
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取参数", mlng险类)
    
    If Not rsTemp.EOF Then txtEdit.Text = rsTemp!参数值
End Sub

Public Function ShowME(ByVal lng险类 As Long) As Boolean
    mlng险类 = lng险类
    Me.Show 1
    ShowME = mblnReturn
End Function
