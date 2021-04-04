VERSION 5.00
Begin VB.Form frmDiluteSample 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "标本稀释"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   2820
      TabIndex        =   4
      Top             =   1110
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   1350
      TabIndex        =   3
      Top             =   1110
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -240
      TabIndex        =   2
      Top             =   960
      Width           =   4425
   End
   Begin VB.TextBox txtDilute 
      Height          =   300
      Left            =   1110
      MaxLength       =   6
      TabIndex        =   1
      Top             =   300
      Width           =   2775
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "稀释倍数:"
      Height          =   210
      Left            =   300
      TabIndex        =   0
      Top             =   345
      Width           =   1260
   End
End
Attribute VB_Name = "frmDiluteSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlngSampleID As Long
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub ShowME(frmObj As Object, lngSampleID As Long)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                   打开稀释窗口
    '参数                   frmObj 父窗口对象
    '                       lngSampleID 标本ID
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    mlngSampleID = lngSampleID
    
    Me.Show vbModal, frmObj
    
End Sub

Private Sub cmdOK_Click()
    '保存稀释倍数
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Val(Me.txtDilute) = 0 Then
        MsgBox "请输入稀释倍数!"
        Me.txtDilute.SetFocus
        Exit Sub
    End If
    
    strSQL = "Zl_检验标本稀释_update(" & mlngSampleID & "," & Val(Me.txtDilute) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "Zl_重新计算结果_Cale(" & mlngSampleID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    MsgBox "稀释调整完成!"
    
    Unload Me
    
    Exit Sub
errH:
    
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDilute_KeyPress(KeyAscii As Integer)
    KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789.*")
End Sub
