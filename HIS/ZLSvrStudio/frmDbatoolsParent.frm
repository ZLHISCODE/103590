VERSION 5.00
Begin VB.Form frmDbatoolsParent 
   BackColor       =   &H00FFFFFF&
   Caption         =   "性能优化工具"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   ControlBox      =   0   'False
   DrawMode        =   3  'Not Merge Pen
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmDbatoolsParent.frx":0000
   ScaleHeight     =   6540
   ScaleWidth      =   8085
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctContent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5700
      Left            =   0
      ScaleHeight     =   5700
      ScaleWidth      =   7605
      TabIndex        =   1
      Top             =   480
      Width           =   7605
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性能优化"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   960
   End
End
Attribute VB_Name = "frmDbatoolsParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmTools As Object
Attribute mfrmTools.VB_VarHelpID = -1

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Private Sub Form_Resize()

    If mfrmTools Is Nothing Then Exit Sub
    
    On Error Resume Next
    pctContent.Height = Me.ScaleHeight - pctContent.Top
    pctContent.Width = Me.ScaleWidth
    
    mfrmTools.WindowState = 0
    mfrmTools.Move 0, 0, pctContent.ScaleWidth, pctContent.ScaleHeight
End Sub

Public Sub ShowToolsForm(ByVal strMoudle As String)
    Static objTools As Object
    
    On Error GoTo errH
    If objTools Is Nothing Then
        Set objTools = CreateObject("zlDbaTools.clsToolsMain")
    End If
    
    If objTools Is Nothing Then
        Me.Show
        frmMDIMain.stbThis.Panels(2).Text = "DBA工具加载失败，请检查zlDbaTools.dll是否成功注册。"
        Exit Sub
    End If
    
    Set mfrmTools = objTools.GetFrmByMdoudle(strMoudle, gblnDBA, gcnOracle, gstrUserName, gstrPassword)
    
    Select Case strMoudle
    Case "0601"
        Call ShowFlash("正在加载数据库性能分析工具...")
        lblTitle.Caption = "数据库性能分析"
    Case "0602"
        Call ShowFlash("正在加载SQL性能分析与优化工具...")
        lblTitle.Caption = "SQL性能分析与优化"
    Case "0604"
        Call ShowFlash("正在加载会话解锁工具...")
        lblTitle.Caption = "会话解锁"
    Case "0605"
        Call ShowFlash("正在加载外键索引工具...")
        lblTitle.Caption = "外键索引"
    Case "0606"
        Call ShowFlash("正在加载空间分析与整理工具...")
        lblTitle.Caption = "空间分析与整理"
    End Select

    '窗体应该有一个ShowMe方法。
    If mfrmTools Is Nothing Then
        Me.Show
        frmMDIMain.stbThis.Panels(2).Text = "当前用户不是DBA用户，权限不足，无法使用该功能。"
        Call ShowFlash("")
        Exit Sub
    Else
        LockWindowUpdate Me.hwnd
        SetParent mfrmTools.hwnd, pctContent.hwnd
        mfrmTools.ShowMe
    End If
    
    Form_Resize
    Call ShowFlash("")
    LockWindowUpdate 0
    Exit Sub
errH:
	Call ShowFlash("")
    MsgBox err.Description
End Sub


