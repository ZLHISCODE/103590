VERSION 5.00
Begin VB.Form frmPosSample 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "位置"
   ClientHeight    =   4800
   ClientLeft      =   6975
   ClientTop       =   4740
   ClientWidth     =   6675
   Icon            =   "frmPosSample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmr 
      Interval        =   500
      Left            =   945
      Top             =   4065
   End
   Begin VB.PictureBox pic 
      Height          =   3750
      Left            =   0
      Picture         =   "frmPosSample.frx":000C
      ScaleHeight     =   3690
      ScaleWidth      =   4905
      TabIndex        =   0
      Top             =   0
      Width           =   4965
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "主页图片"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   300
         Left            =   2115
         TabIndex        =   2
         Top             =   1395
         Width           =   1260
      End
      Begin VB.Shape shp 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         Height          =   510
         Left            =   1560
         Top             =   2535
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5805
      TabIndex        =   1
      Top             =   165
      Width           =   1100
   End
End
Attribute VB_Name = "frmPosSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarFirst As Boolean
Private mvarPos As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If mvarFirst = False Then Exit Sub
    mvarFirst = False
    
    lblNote.Caption = mvarPos
    Me.Caption = mvarPos & "的显示位置"
    
    Select Case mvarPos
    Case "主页图片"
        Call ResizeControl(shp, 1725, 810, 2220, 1425)
    Case "主页背景", "页面背景"
        Call ResizeControl(shp, 780, 420, 4140, 3120)
    Case "宣传标语"
        Call ResizeControl(shp, 780, 15, 4140, 375)
    Case "标志图片"
        Call ResizeControl(shp, 15, 15, 750, 375)
    Case "广告图片"
        Call ResizeControl(shp, 15, 3045, 750, 510)
    End Select
    Call ResizeControl(lblNote, shp.Left + (shp.Width - lblNote.Width) / 2, shp.Top + (shp.Height - lblNote.Height) / 2, lblNote.Width, lblNote.Height)
    If lblNote.Left < shp.Left Then lblNote.Left = shp.Left
End Sub

Private Sub Form_Load()
    With frmPosSample
        .Width = 5070 - 30
        .Height = 4095 - 30
        .Left = Screen.Width - .Width
        .Top = Screen.Height - .Height - 450
        
        .Left = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\界面\" & Me.Name, "Left", .Left)
        .Top = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\界面\" & Me.Name, "Top", .Top)
        
    End With

    
    mvarFirst = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    pic.Left = 0
    pic.Top = 0
    pic.Width = Me.ScaleWidth
    pic.Height = Me.ScaleHeight
    
End Sub

Public Sub ShowPageSample(ByVal strPos As String)
    mvarPos = strPos
    With frmPosSample
        .Show 1
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With frmPosSample
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\界面\" & Me.Name, "Left", .Left
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\界面\" & Me.Name, "Top", .Top
    End With
End Sub

Private Sub pic_Paint()
    Call RaisEffect(pic, -1)
End Sub

Private Sub tmr_Timer()
    shp.Visible = Not shp.Visible
    lblNote.Visible = Not lblNote.Visible
End Sub
