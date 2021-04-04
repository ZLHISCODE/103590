VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendFileMutilEditor 
   BackColor       =   &H00FFFFFF&
   Caption         =   "批量录入护理记录单"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   Icon            =   "frmTendFileMutilEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPrompt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1410
      ScaleHeight     =   285
      ScaleWidth      =   9585
      TabIndex        =   2
      Top             =   7620
      Width           =   9585
      Begin VB.Label lblPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   30
         TabIndex        =   3
         Top             =   60
         Width           =   10500
      End
   End
   Begin zl9TendFile.usrTendFileMutilEditor usrTendFileMutilEditor 
      Height          =   5085
      Left            =   300
      TabIndex        =   1
      Top             =   630
      Width           =   7665
      _extentx        =   13520
      _extenty        =   8969
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendFileMutilEditor.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17224
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
Attribute VB_Name = "frmTendFileMutilEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmTipInfo As Object

Private Sub Form_Load()
    If mfrmTipInfo Is Nothing Then
        Set mfrmTipInfo = New frmTipInfo
    End If
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With usrTendFileMutilEditor
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
        .FontSize = Me.FontSize
        .Width = picPrompt.Width
        .Height = TextHeight("刘")
        .Top = (picPrompt.Height - .Height) \ 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmTipInfo Is Nothing Then
        Unload mfrmTipInfo
        Set mfrmTipInfo = Nothing
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Public Sub ShowMe(ByVal frmParent As Form, ByVal lngDeptID As Long, ByVal strPrivs As String, Optional ByVal bytSize As Byte = 0)
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngFileID           护理文件格式句柄
    '       lngPatiID           病人id
    '       lngPageID           主页id
    '       intBaby             婴儿标志
    '返回： 无
    '******************************************************************************************************************
    
    Err = 0
    On Error GoTo errHand
    Me.FontSize = IIf(bytSize = 1, 12, 9)
    If Not usrTendFileMutilEditor.ShowMe(Me, lngDeptID, strPrivs, bytSize) Then Exit Sub
    
    '窗体显示
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub usrTendFileMutilEditor_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean)
    lblPrompt.Caption = strInfo
    lblPrompt.ForeColor = IIf(blnImportant, &HFF&, &H80000008)
End Sub

Private Sub usrTendFileMutilEditor_ShowTipInfo(ByVal vsfObj As Object, ByVal strInfo As String, ByVal blnMultiRow As Boolean)
    Call mfrmTipInfo.ShowTipInfo(vsfObj, strInfo, blnMultiRow)
End Sub

Private Sub usrTendFileMutilEditor_UsrExit()
    Unload Me
End Sub

Private Sub usrTendFileMutilEditor_UsrHelp()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub
