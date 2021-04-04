VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFgaMore 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SQL语句详情和绑定变量"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9405
   Icon            =   "frmFgaMore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9405
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pctBind 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1080
      ScaleHeight     =   1695
      ScaleWidth      =   7575
      TabIndex        =   3
      Top             =   3360
      Width           =   7575
      Begin RichTextLib.RichTextBox txtBind 
         Height          =   1335
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2355
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmFgaMore.frx":6852
      End
      Begin VB.Label lblBind 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "绑定变量"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.PictureBox pctSQL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2535
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin RichTextLib.RichTextBox txtSql 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3625
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmFgaMore.frx":68EF
      End
      Begin VB.Label lblSQL 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "SQL语句"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmFgaMore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowMe(ByVal strSql As String, ByVal strBind As String)
    txtSql.Text = strSql: txtBind.Text = strBind
    Me.Show 1
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    pctBind.Top = Me.ScaleHeight - pctBind.Height
    pctBind.Left = 0
    pctBind.Width = Me.ScaleWidth
    
    txtBind.Width = pctBind.ScaleWidth - txtBind.Left - 120
    
    pctSQL.Top = 0
    pctSQL.Left = 0
    pctSQL.Height = pctBind.Top - 60
    pctSQL.Width = Me.ScaleWidth
    
    txtSql.Width = txtBind.Width
    txtSql.Height = pctSQL.ScaleHeight - txtSql.Top
End Sub
