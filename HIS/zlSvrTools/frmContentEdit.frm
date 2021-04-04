VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmContentEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F4E4&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   4935
      TabIndex        =   1
      Top             =   2910
      Width           =   4965
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3720
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2520
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
   End
   Begin RichTextLib.RichTextBox rtfEditor 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmContentEdit.frx":0000
   End
End
Attribute VB_Name = "frmContentEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrContent As String
Private mblnOk As Boolean
Private mlngX As Long
Private mlngY As Long
Private mlngW As Long '父窗体宽度
Private mlngH As Long '父窗体高度
Private mlngL As Long '父窗体left
Private mlngT As Long '父窗体top
Private mlngBL As Long '父窗体左边框长度
Private mlngBT As Long '父窗体上边框长度
Private mlngWidth As Long '设置当前界面宽度
Private mlngHeight As Long '设置当期界面高度
Public Function ShowMe(ByRef frmParent As Object, ByRef strContent As String, ByVal lngX As Long, ByVal lngY As Long, Optional ByVal lngHeight As Long, Optional ByVal lngWidth As Long) As Boolean
'功能：显示表单项目编辑器
    mstrContent = strContent
    mblnOk = False
    mlngX = lngX
    mlngY = lngY
    mlngW = frmParent.Width
    mlngH = frmParent.Height
    mlngL = frmParent.Left
    mlngT = frmParent.Top
    mlngWidth = lngWidth
    mlngHeight = lngHeight
    mlngBL = (frmParent.Width - frmParent.ScaleWidth) / 2
    mlngBT = (frmParent.Height - frmParent.ScaleHeight) * 2 / 3
    
    Me.Show 1
    strContent = mstrContent
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mstrContent = rtfEditor.Text Then
        mblnOk = False
    Else
        mblnOk = True
    End If
    mstrContent = rtfEditor.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long
    If mlngWidth = 0 Then mlngWidth = 4965
    If mlngHeight = 0 Then mlngHeight = 3525
    
    Me.Width = mlngWidth
    Me.Height = mlngHeight
    
    rtfEditor.Text = mstrContent
    '设置窗体位置
    If mlngX + Me.Width > mlngW + mlngL Then  '超出父窗体边界
        Me.Left = mlngW + mlngL - Me.Width
    Else
        Me.Left = mlngX
    End If
    
    If mlngY + Me.Height > mlngH + mlngT Then '超出父窗体边界
        Me.Top = mlngH + mlngT - mlngBT - Me.Height
    Else '未超出父窗体边界
        Me.Top = mlngY
    End If
    Call Form_Resize
End Sub

Private Sub Form_Resize()

    rtfEditor.Width = Me.ScaleWidth
    rtfEditor.Height = Me.ScaleHeight - picBottom.Height - 5
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'功能：回车定位下一个控件
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PressKey(vbKeyTab)
    End If
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = picBottom.ScaleWidth - cmdCancel.Width - 60
    cmdOk.Left = cmdCancel.Left - cmdOk.Width - 60
End Sub
