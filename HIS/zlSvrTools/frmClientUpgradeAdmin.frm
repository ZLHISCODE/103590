VERSION 5.00
Begin VB.Form frmClientUpgradeAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "客户端管理用户密码设置"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4185
   Icon            =   "frmClientUpgradeAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1155
      ScaleHeight     =   270
      ScaleWidth      =   2820
      TabIndex        =   7
      Top             =   720
      Width           =   2850
      Begin VB.TextBox txtPass 
         BorderStyle     =   0  'None
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   60
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   30
         Width           =   2715
      End
   End
   Begin VB.PictureBox picUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1155
      ScaleHeight     =   270
      ScaleWidth      =   2820
      TabIndex        =   6
      Top             =   255
      Width           =   2850
      Begin VB.TextBox txtUser 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   60
         TabIndex        =   8
         Text            =   "Administrator"
         Top             =   30
         Width           =   2715
      End
   End
   Begin VB.TextBox txtPass123 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1155
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.TextBox txtUser1111 
      Height          =   315
      Left            =   1155
      TabIndex        =   0
      Text            =   "Administrator"
      Top             =   255
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.CommandButton cmd保存 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   1785
      TabIndex        =   2
      Top             =   1200
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   2895
      TabIndex        =   4
      Top             =   1200
      Width           =   1100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "管理密码"
      Height          =   180
      Left            =   225
      TabIndex        =   5
      Top             =   795
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "管理用户"
      Height          =   180
      Left            =   225
      TabIndex        =   3
      Top             =   315
      Width           =   720
   End
End
Attribute VB_Name = "frmClientUpgradeAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOk As Boolean

'关闭
Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub Cmd保存_Click()
    Dim strUser As String
    Dim strPass As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strUser = txtUser.Text
    strPass = Cipher(txtPass.Text)
    
    '保存或新建账号
    gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 = '管理员账号'"
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    If rsTmp.EOF = False Then
        strSQL = "Update zlRegInfo Set 内容='" & strUser & "' Where 项目='管理员账号'"
        gcnOracle.Execute strSQL
    Else
        strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('管理员账号',Null,'" & strUser & "')"
        gcnOracle.Execute strSQL
    End If
    
    '保存或新建密码
    gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 = '管理员密码'"
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    If rsTmp.EOF = False Then
        strSQL = "Update zlRegInfo Set 内容='" & strPass & "' Where 项目='管理员密码'"
        gcnOracle.Execute strSQL
    Else
        strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('管理员密码',Null,'" & strPass & "')"
        gcnOracle.Execute strSQL
    End If
    
    mblnOk = True
    Unload Me
  Exit Sub
errHand:
    MsgBox err.Description, vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Sub Form_Load()
    
    '读取管理员用户名及密码
    Call LoadReadAdmin
End Sub

Private Sub Form_Resize()
    With Cmd保存
        .Top = Cmd保存.Top
        .Left = cmdCancel.Left - .Width - 30
    End With
End Sub

Private Sub LoadReadAdmin()
    Dim rsTmp As New ADODB.Recordset
    Dim rsTmpPass As New ADODB.Recordset
    gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 like '管理员账号'"
    Call OpenRecordset(rsTmp, gstrSQL, "管理")
    
    If rsTmp.RecordCount = 1 Then
        txtUser.Text = Trim(Nvl(rsTmp!内容))
        gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 like '管理员密码'"
        Call OpenRecordset(rsTmpPass, gstrSQL, Me.Caption)
        If rsTmpPass.RecordCount = 1 Then
            txtPass.Text = Decipher(Trim(Nvl(rsTmpPass!内容)))
        Else
            txtPass.Text = ""
        End If
    Else
        txtUser.Text = "Administrator"
        txtPass.Text = ""
    End If
    
End Sub


'密码加密程序
Private Function Cipher(stext As String)
    Const MIN_ASC = 32
    Const MAX_ASC = 126
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim offset As Long
    Dim strlen As Integer
    Dim i As Integer
    Dim ch As Integer
    Dim ptext As String
    offset = 123
    Rnd (-1)
    Randomize (offset)
    strlen = Len(stext)
    For i = 1 To strlen
       ch = Asc(Mid(stext, i, 1))
       If ch >= MIN_ASC And ch <= MAX_ASC Then
           ch = ch - MIN_ASC
           offset = Int((NUM_ASC + 1) * Rnd())
           ch = ((ch + offset) Mod NUM_ASC)
           ch = ch + MIN_ASC
           ptext = ptext & Chr(ch)
       End If
    Next i
    Cipher = ptext
End Function

'解密程序
Private Function Decipher(stext As String)      '密码解密程序
    Const MIN_ASC = 32 '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim offset As Long
    Dim strlen As Integer
    Dim i As Integer
    Dim ch As Integer
    Dim ptext As String
    offset = 123
    Rnd (-1)
    Randomize (offset)
    strlen = Len(stext)
    For i = 1 To strlen
       ch = Asc(Mid(stext, i, 1)) '取字母转变成ASCII码
       If ch >= MIN_ASC And ch <= MAX_ASC Then
           ch = ch - MIN_ASC
           offset = Int((NUM_ASC + 1) * Rnd())
           ch = ((ch - offset) Mod NUM_ASC)
           If ch < 0 Then
               ch = ch + NUM_ASC
           End If
           ch = ch + MIN_ASC
           ptext = ptext & Chr(ch)
       End If
    Next i
    Decipher = ptext
End Function
