VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStTableItemEdit 
   BorderStyle     =   0  'None
   Caption         =   "表单项目编辑"
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5550
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
      ScaleWidth      =   5520
      TabIndex        =   3
      Top             =   2940
      Width           =   5550
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2880
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4200
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
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5106
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmStTableItemEdit.frx":0000
   End
End
Attribute VB_Name = "frmStTableItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrContent As String
Private mblnOK As Boolean
Private mlngX As Long
Private mlngY As Long
Private mlngW As Long '父窗体宽度
Private mlngH As Long '父窗体高度
Private mlngL As Long '父窗体left
Private mlngT As Long '父窗体top
Private mlngBL As Long '父窗体左边框长度
Private mlngBT As Long '父窗体上边框长度
Public Function ShowMe(ByRef frmParent As Object, ByRef strContent As String, ByVal lngX As Long, ByVal lngY As Long) As Boolean
'功能：显示表单项目编辑器
    mstrContent = strContent
    mblnOK = False
    mlngX = lngX
    mlngY = lngY
    mlngW = frmParent.Width
    mlngH = frmParent.Height
    mlngL = frmParent.Left
    mlngT = frmParent.Top
    mlngBL = (frmParent.Width - frmParent.ScaleWidth) / 2
    mlngBT = (frmParent.Height - frmParent.ScaleHeight) * 2 / 3
    
    Me.Show 1
    strContent = mstrContent
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mstrContent = rtfEditor.Text Then
        mblnOK = False
    Else
        mblnOK = True
    End If
    mstrContent = rtfEditor.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long
    
    rtfEditor.Text = mstrContent
    '设置窗体位置
    If mlngX + Me.Width > mlngW + mlngL - mlngBL Then '超出父窗体边界
        Me.Left = mlngW + mlngL - mlngBL - Me.Width
    Else
        Me.Left = mlngX
    End If
    
    If mlngY + Me.Height > mlngH + mlngT - mlngBT Then '超出父窗体边界
        Me.Top = mlngH + mlngT - mlngBT - Me.Height
    Else '未超出父窗体边界
        Me.Top = mlngY
    End If
    
End Sub

Private Sub Form_Resize()

    rtfEditor.Width = Me.ScaleWidth
    rtfEditor.Height = Me.ScaleHeight - picBottom.Height - 5
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'功能：回车定位下一个控件
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub
