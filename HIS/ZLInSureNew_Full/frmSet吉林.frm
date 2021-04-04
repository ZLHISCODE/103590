VERSION 5.00
Begin VB.Form frmSet吉林 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSet吉林.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   5265
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2280
      TabIndex        =   3
      Top             =   2340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3420
      TabIndex        =   2
      Top             =   2340
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   5265
   End
   Begin VB.CheckBox chk明细 
      Caption         =   "明细时实上传(&D)"
      Height          =   210
      Left            =   870
      TabIndex        =   0
      Top             =   1500
      Width           =   3375
   End
   Begin VB.Label lbl 
      Caption         =   "配置相关的参数."
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   570
      Width           =   7125
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   270
      Picture         =   "frmSet吉林.frx":014A
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmSet吉林"
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
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_吉林 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    gstrSQL = "zl_保险参数_Insert(" & TYPE_吉林 & ",null,'明细时实上传','" & IIf(chk明细.Value = 1, 1, 0) & "',0)"
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
    
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_吉林
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!参数名)
            Case "明细时实上传"
                chk明细.Value = IIf(Nvl(!参数值, 1) = 1, 1, 0)
            End Select
            .MoveNext
        Loop
    End With

End Sub

Public Function 参数设置() As Boolean
    frmSet吉林.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function
End Function

