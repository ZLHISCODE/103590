VERSION 5.00
Begin VB.Form frmInputBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输入"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frmInputBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   5445
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1650
      Width           =   5445
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4215
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3015
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   6000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   -60
         X2              =   6000
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.TextBox txtMultiLine 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   420
      Width           =   5175
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入XXX:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   4
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mintLen As Integer
Private mblnNULL As Boolean
Private mblnEnter As Boolean
Private mstrText As String
Private mblnOk As Boolean
Private mlngTXTProc As Long

Public Function InputBox(frmParent As Object, ByVal 标题 As String, ByVal 提示 As String, _
    ByVal 长度 As Integer, ByVal 行数 As Byte, ByVal 允许空 As Boolean, _
    ByVal 允许回车 As Boolean, ByRef 内容 As String) As Boolean
   
    Load Me
    
    Caption = 标题
    lblInfo.Caption = 提示
    txtMultiLine.MaxLength = 长度
    txtMultiLine.Height = IIf(行数 = 0, 1, 行数) * 300
    Me.Height = txtMultiLine.Top + txtMultiLine.Height + picCmd.Height + 500
    
    txtMultiLine.Text = 内容
    
    mintLen = 长度
    mblnNULL = 允许空
    mblnEnter = 允许回车
    Me.Show 1, frmParent
    
    If Not mblnOk Then 内容 = "": Exit Function
    内容 = mstrText
    InputBox = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not mblnNULL And txtMultiLine.Text = "" Then
        MsgBox "必须输入" & Caption & "！", vbInformation, gstrSysName
        txtMultiLine.SetFocus: Exit Sub
    End If
    If mintLen > 0 Then
        If zlCommFun.ActualLen(txtMultiLine.Text) > mintLen Then
            MsgBox Caption & "最多允许输入 " & mintLen & " 个字符或 " & mintLen \ 2 & " 个汉字！", vbInformation, gstrSysName
            txtMultiLine.SetFocus: Exit Sub
        End If
    End If

    mstrText = txtMultiLine.Text
    
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("';|" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOk = False
End Sub

Private Sub txtMultiLine_GotFocus()
    zlControl.TxtSelAll txtMultiLine
End Sub

Private Sub txtMultiLine_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtMultiLine.Text = "" Then
            cmdOK_Click
        ElseIf Not mblnEnter Then
            cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub txtMultiLine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        mlngTXTProc = GetWindowLong(txtMultiLine.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtMultiLine.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtMultiLine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtMultiLine.hwnd, GWL_WNDPROC, mlngTXTProc)
    End If
End Sub
