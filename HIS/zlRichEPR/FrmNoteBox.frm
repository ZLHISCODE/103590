VERSION 5.00
Begin VB.Form FrmNoteBox 
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   Icon            =   "FrmNoteBox.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6720
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAll 
      Caption         =   "语法检测(&A)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3135
      TabIndex        =   3
      Top             =   4890
      Width           =   1320
   End
   Begin VB.TextBox TEdit 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4635
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      ToolTipText     =   "双击关闭"
      Top             =   90
      Width           =   6495
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4635
      TabIndex        =   1
      Top             =   4890
      Width           =   960
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5625
      TabIndex        =   0
      Top             =   4890
      Width           =   960
   End
End
Attribute VB_Name = "FrmNoteBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fStrText            As String
Private fStrTile            As String
Private fReadOnly           As Boolean
Private fSqlCheck           As Boolean

Public Property Get StrText() As String
    StrText = fStrText
End Property

Public Property Let StrText(ByVal vStrText As String)
    fStrText = vStrText
End Property

Public Property Let StrTile(ByVal vStrTile As String)
    fStrTile = vStrTile
End Property

Public Property Let ReadOnly(ByVal vReadOnly As Boolean)
    fReadOnly = vReadOnly
End Property

Public Property Let SqlCheck(ByVal vSqlCheck As Boolean)
    fSqlCheck = vSqlCheck
End Property

Private Sub cmdAll_Click()
    On Error GoTo ErrH
    If TEdit.Text <> "" Then Call CheckAuditSql_IN(TEdit.Text, True)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdCancel_Click()
On Error GoTo ErrH
    Unload Me
    Exit Sub
ErrH:
    Err.Clear
    Unload Me
    Exit Sub
End Sub

Private Sub CmdOK_Click()
On Error GoTo ErrH
    StrText = TEdit.Text
    Me.Hide
    Unload Me
    Exit Sub
ErrH:
    Err.Clear
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Caption = fStrTile
    TEdit.Text = fStrText
    cmdAll.Visible = fSqlCheck
    If fReadOnly Then
        TEdit.Locked = True
        TEdit.ForeColor = vbBlue
        CmdOk.Visible = False
        CmdCancel.Caption = "关闭(&C)"
    End If
End Sub

Private Sub Form_Resize()
On Error GoTo ErrH
    If Me.WindowState <> 1 Then
        With TEdit
            .Left = 10
            .Top = 10
            .Height = Me.ScaleHeight - .Top - CmdOk.Height - 300
            .Width = Me.ScaleWidth - 2 * .Left
        End With
        With CmdOk
            .Top = TEdit.Top + TEdit.Height + 200
            .Left = Me.ScaleWidth - .Width - 1400
        End With
        With CmdCancel
            .Top = CmdOk.Top
            .Left = Me.ScaleWidth - .Width - 320
        End With
        With cmdAll
            .Top = CmdOk.Top
            .Left = CmdOk.Left - .Width - 100
        End With
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub TEdit_DblClick()
On Error Resume Next
    If Me.Enabled = True Then
        fStrText = TEdit.Text
    End If
    Unload Me
End Sub
