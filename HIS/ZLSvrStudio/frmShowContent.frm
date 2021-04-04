VERSION 5.00
Begin VB.Form frmShowContent 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "通用显示内容窗"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtContent 
      Appearance      =   0  'Flat
      Height          =   2805
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmShowContent.frx":0000
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "frmShowContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytMode As Byte
Private mstrCaption As String
Private mstrContent As String

Public Enum Enum_ShowContentMode
    scmCommonLogContent = 0
End Enum

Public Property Get Mode() As Enum_ShowContentMode
    Mode = mbytMode
End Property
Public Property Let Mode(ByVal value As Enum_ShowContentMode)
    mbytMode = value
End Property
        
Public Sub ShowMe(ByVal strCaption As String, ByVal strContent As String)
    mstrCaption = strCaption
    mstrContent = strContent
    Me.Show vbModal, frmMDIMain
End Sub

Private Sub Form_Initialize()
    '变量缺省值
    mstrCaption = scmCommonLogContent
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    '初始化
    Select Case mbytMode
    Case scmCommonLogContent
        Me.KeyPreview = True
        Me.Caption = mstrCaption
        Me.Width = 640 * 15
        Me.Height = 240 * 15
        txtContent.Text = mstrContent
        txtContent.Locked = True
        mstrContent = ""
        mstrCaption = ""
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Select Case mbytMode
    Case scmCommonLogContent
        txtContent.Move 30, 30, Me.ScaleWidth - 30 * 2, Me.ScaleHeight - 30 * 2
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '...
End Sub

Private Sub txtContent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        txtContent.SelStart = 0
        txtContent.SelLength = Len(txtContent.Text)
    End If
End Sub
