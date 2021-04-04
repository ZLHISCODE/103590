VERSION 5.00
Begin VB.Form frmSet乐山 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSet乐山.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra医保服务器 
      Caption         =   "医保服务器"
      Height          =   1605
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   3165
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   4
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1110
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&A)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   390
         Width           =   810
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "操作员(&U)"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   3
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "密码(&W)"
         Height          =   180
         Index           =   2
         Left            =   570
         TabIndex        =   5
         Top             =   1170
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3510
      TabIndex        =   7
      Top             =   390
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   8
      Top             =   870
      Width           =   1100
   End
End
Attribute VB_Name = "frmSet乐山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum 参数
    地址 = 0
    登录操作员
    登录口令
End Enum
Private mblnReturn As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    If Not Valid Then Exit Sub
    gcnOracle.BeginTrans
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_乐山 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_乐山 & ",NULL,'服务器地址','" & txtEdit(参数.地址).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_乐山 & ",NULL,'登录操作员','" & txtEdit(参数.登录操作员).Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_乐山 & ",NULL,'登录口令','" & txtEdit(参数.登录口令).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    mblnReturn = True
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Function Valid() As Boolean
    Dim strPara As String, arrPara
    Dim intDO As Integer, intUbound As Integer
    '检查是否必需输入的参数都输入了
    strPara = "服务器地址||登录操作员||口令"
    arrPara = Split(strPara, "||")
    
    intUbound = txtEdit.Count - 1
    For intDO = 0 To intUbound
        If Trim(txtEdit(intDO)) = "" Then
            MsgBox arrPara(intDO) & "不能为空！", vbInformation, gstrSysName
            txtEdit(intDO).SetFocus
            Exit Function
        End If
    Next
    
    Valid = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    '获取服务器地址
    gstrSQL = " Select 参数名,参数值 From 保险参数" & _
              " Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保相关参数", TYPE_乐山)
    
    With rsTemp
        Do While Not .EOF
            Select Case !参数名
            Case "服务器地址"
                txtEdit(参数.地址).Text = Nvl(!参数值)
            Case "登录操作员"
                txtEdit(参数.登录操作员).Text = Nvl(!参数值)
            Case "登录口令"
                txtEdit(参数.登录口令).Text = Nvl(!参数值)
            End Select
            .MoveNext
        Loop
    End With
End Sub

Public Function ShowME() As Boolean
    mblnReturn = False
    Me.Show 1
    ShowME = mblnReturn
End Function
