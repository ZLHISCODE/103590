VERSION 5.00
Begin VB.Form frmSet凯里 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1253
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "1"
      Top             =   163
      Width           =   360
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   2310
      TabIndex        =   2
      Top             =   870
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   3450
      TabIndex        =   1
      Top             =   870
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   1
      Left            =   -292
      TabIndex        =   0
      Top             =   673
      Width           =   5265
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "当前串口(&D)"
      Height          =   180
      Index           =   3
      Left            =   158
      TabIndex        =   5
      Top             =   223
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "号串口"
      Height          =   180
      Index           =   4
      Left            =   1658
      TabIndex        =   4
      Top             =   223
      Width           =   540
   End
End
Attribute VB_Name = "frmSet凯里"
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
    On Error GoTo ErrHand

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
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mblnReturn = False

    gstrSQL = "Select 参数值 From 保险参数 Where 险类=" & mlng险类
    Call OpenRecordset(rsTemp, "读取参数")

    If Not rsTemp.EOF Then txtEdit.Text = rsTemp!参数值
End Sub

Public Function ShowME(ByVal lng险类 As Long) As Boolean
    mlng险类 = lng险类
    Me.Show 1
    ShowME = mblnReturn
End Function

