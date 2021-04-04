VERSION 5.00
Begin VB.Form frmHelpTopic 
   BorderStyle     =   0  'None
   Caption         =   "帮助主题搜索"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1740
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMainSkin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   0
      Picture         =   "frmHelpTopic.frx":0000
      ScaleHeight     =   1740
      ScaleWidth      =   2820
      TabIndex        =   0
      Top             =   0
      Width           =   2820
      Begin VB.CommandButton Command2 
         Caption         =   "搜索(&S)"
         Default         =   -1  'True
         Height          =   300
         Left            =   1725
         TabIndex        =   4
         Top             =   1035
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "选项(&O)"
         Height          =   300
         Left            =   135
         TabIndex        =   3
         Top             =   1035
         Width           =   960
      End
      Begin VB.TextBox txtThis 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmHelpTopic.frx":024B
         Top             =   405
         Width           =   2550
      End
      Begin VB.Label lblMSG 
         BackStyle       =   0  'Transparent
         Caption         =   "请问你要做什么？"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   135
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmHelpTopic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    If txtThis.Enabled And txtThis.Visible Then txtThis.SetFocus
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i
    i = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)    '窗体在最前
    
    Dim WindowRegion As Long
    
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)  '异型窗体
    SetWindowRgn Me.hWnd, WindowRegion, True
    
    SetOpacityForm Me, 230  '设置透明度，0～255。
End Sub

Private Sub txtThis_GotFocus()
    txtThis.SelStart = 0
    txtThis.SelLength = Len(txtThis)
End Sub
