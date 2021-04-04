VERSION 5.00
Begin VB.Form frmSet兴安 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数设置"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chk住院明细 
      Caption         =   "住院结算补传明细"
      Height          =   300
      Left            =   840
      TabIndex        =   8
      Top             =   1875
      Width           =   1755
   End
   Begin VB.TextBox txt等待 
      Height          =   300
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "400"
      Top             =   1140
      Width           =   465
   End
   Begin VB.CheckBox chk自动 
      Caption         =   "自动启动中联监听(&C)"
      Height          =   210
      Left            =   840
      TabIndex        =   4
      Top             =   1620
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   1
      Left            =   -150
      TabIndex        =   3
      Top             =   810
      Width           =   5265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   2
      Top             =   2430
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2430
      TabIndex        =   1
      Top             =   2430
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   0
      Left            =   -150
      TabIndex        =   0
      Top             =   2220
      Width           =   5265
   End
   Begin VB.Label lblInif 
      AutoSize        =   -1  'True
      Caption         =   "向医保中发出请求等待时间       秒"
      Height          =   180
      Left            =   270
      TabIndex        =   6
      Top             =   1230
      Width           =   2970
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   120
      Picture         =   "frmSet兴安.frx":0000
      Top             =   210
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "配置相关的参数."
      Height          =   195
      Index           =   0
      Left            =   690
      TabIndex        =   5
      Top             =   420
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet兴安"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    If Val(txt等待.Text) <= 10 Then
        ShowMsgbox "等待时间不能小于10秒"
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_兴安 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_保险参数_Insert(" & TYPE_兴安 & ",null,'自动启动监听','" & IIf(chk自动.Value = 1, 1, 0) & "',0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_兴安 & ",null,'请求等待时间','" & Val(txt等待.Text) & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_兴安 & ",null,'住院结算补传明细','" & IIf(chk住院明细.Value = 1, 1, 0) & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans
    mblnOK = True
    Unload Me
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_兴安
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!参数名)
            Case "自动启动监听"
                chk自动.Value = IIf(Nvl(!参数值, 1) = 1, 1, 0)
            Case "请求等待时间"
                txt等待.Text = Nvl(!参数值, 400)
            Case "住院结算补传明细"
                chk住院明细.Value = IIf(Nvl(!参数值, 1) = 1, 1, 0)
            End Select
            .MoveNext
        Loop
    End With

End Sub

Public Function 参数设置() As Boolean
    frmSet兴安.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function
End Function


Private Sub txt等待_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt等待, KeyAscii, m数字式
End Sub
