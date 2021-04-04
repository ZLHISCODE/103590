VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendBodyMutilEditor 
   BackColor       =   &H00FFFFFF&
   Caption         =   "批量录入体温单"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11550
   Icon            =   "frmTendBodyMutilEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   11550
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin zl9TemperatureChartSX.UserMutilEditor BodyMutilEditor 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11245
   End
   Begin VB.PictureBox picPrompt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1440
      ScaleHeight     =   285
      ScaleWidth      =   9465
      TabIndex        =   1
      Top             =   6840
      Width           =   9465
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   15
         TabIndex        =   2
         Top             =   80
         Width           =   90
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6720
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendBodyMutilEditor.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17463
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTendBodyMutilEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BodyMutilEditor_UsrExit()
    Unload Me
End Sub

Private Sub BodyMutilEditor_UsrHelp()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With BodyMutilEditor
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - stbThis.Height
    End With
    With picPrompt
        .Top = Me.ScaleHeight - stbThis.Height + 50
        .Height = stbThis.Height - 100
        .Left = stbThis.Panels(2).Left + 50
        .Width = stbThis.Panels(2).Width - 100
    End With
    
    With lblPrompt
        .Width = picPrompt.Width
        .Height = TextHeight("刘")
        .Top = (picPrompt.Height - .Height) \ 2
    End With
End Sub

Public Sub ShowMe(ByVal frmParent As Form, ByVal lngDeptID As Long, ByVal strPrivs As String, Optional ByVal bytSize As Byte = 0)
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngDeptID           当前病区ID
    '       strPrivs            权限
    '返回： 无
    '******************************************************************************************************************

    Err = 0
    On Error GoTo Errhand
    If Not BodyMutilEditor.ShowMe(Me, lngDeptID, strPrivs, bytSize) Then Exit Sub
    Call ReSetFontSize(bytSize)
    '窗体显示
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If

    Unload Me
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub BodyMutilEditor_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean)
    lblPrompt.Caption = strInfo
    lblPrompt.ForeColor = IIf(blnImportant, &HFF&, &H80000008)
End Sub

Private Sub ReSetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    
    Me.FontSize = bytFontSize
    Me.FontName = "宋体"
    
    lblPrompt.FontSize = bytFontSize
    Call Form_Resize
End Sub
