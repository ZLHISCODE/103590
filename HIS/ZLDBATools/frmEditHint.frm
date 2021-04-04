VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEditHint 
   BackColor       =   &H80000005&
   Caption         =   "编辑自定义提示"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   13935
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pctOperation 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   13935
      TabIndex        =   2
      Top             =   7440
      Width           =   13935
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消（&C）"
         Height          =   350
         Left            =   12600
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   11280
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblTip 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "根据需求修改提示字，请不要改动提示字之外的语句，否则可能会导致SQL PROFILE添加失败或无效。"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   205
         Width           =   8010
      End
   End
   Begin VB.PictureBox pctText 
      BackColor       =   &H80000005&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5115
      ScaleWidth      =   13875
      TabIndex        =   0
      Top             =   0
      Width           =   13935
      Begin RichTextLib.RichTextBox rctSql 
         Height          =   4335
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   7646
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmEditHint.frx":0000
      End
   End
End
Attribute VB_Name = "frmEditHint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSqlID As String
Private mstrSqlText As String
Private mstrCaption   As String
Attribute mstrCaption.VB_VarHelpID = -1
Private mintInstID As String

Public Sub ShowEdit(ByVal strOldSql As String, ByVal strText As String, ByVal intInstID As Integer, ByVal strReturn As String)
    mstrSqlID = strOldSql
    mstrSqlText = strText
    rctSql.Text = strText
    mintInstID = intInstID
    Me.Show 1
    
    strReturn = mstrCaption
End Sub

Private Sub cmdCancel_Click()
    mstrCaption = "操作被取消。"
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strNewSQL As String
    Dim strChild As String
    
    '如果未修改，当作取消处理
    If mstrSqlText = rctSql.Text Then
        mstrCaption = "语句未发生变化，不做修改。"
        Unload Me
        Exit Sub
    End If
    
    strNewSQL = ChangeSQL(5, mstrSqlID, rctSql.Text, strChild, mintInstID)
    
    If strNewSQL = "5" Then
        lblTip.Caption = "编辑提示字失败，请检查输入后再执行。"
        Exit Sub
    End If
    
    If CreateSqlProfiles(mstrSqlID, strNewSQL, strChild) Then
        mstrCaption = "自定义提示字编辑成功。"
        Unload Me
    Else
        lblTip.Caption = "编辑提示字失败，请检查输入后再执行。"
    End If
End Sub


Private Sub Form_Load()
    Me.Icon = Nothing
End Sub

Private Sub Form_Resize()
    pctText.Height = Abs(Me.ScaleHeight - pctOperation.Height)
    pctText.Width = Me.ScaleWidth
    pctOperation.Top = pctText.Height
    pctOperation.Width = Me.ScaleWidth
End Sub

Private Sub pctText_Resize()
    With pctText
        rctSql.Move 0, 0, .Width, .Height
    End With
End Sub


Private Sub pctOperation_resize()
    lblTip.Left = 65
    cmdCancel.Left = pctOperation.ScaleWidth - cmdCancel.Width - 65
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 65
End Sub
