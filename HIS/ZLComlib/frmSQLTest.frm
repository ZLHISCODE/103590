VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSQLTest 
   AutoRedraw      =   -1  'True
   Caption         =   "SQL速度测试"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5220
   Icon            =   "frmSQLTest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5220
   Visible         =   0   'False
   Begin VB.TextBox txt 
      Height          =   2445
      Left            =   315
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   4095
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   2865
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   5054
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmSQLTest.frx":014A
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   1290
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtObject 
      Height          =   1830
      Left            =   885
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileClear 
         Caption         =   "清除(&C)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFile_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "打开(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "另存为(&A)..."
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileText 
         Caption         =   "纯文本(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuFileAnalyse 
         Caption         =   "对象分析(&N)"
      End
      Begin VB.Menu mnuFile_Top 
         Caption         =   "总在最前(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Stop 
         Caption         =   "输出SQL脚本(&E)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmSQLTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFile As String

Private Sub SetTop(blnTop As Boolean)
    If blnTop Then
        SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub Form_Load()
    'gobjComLib.RestoreWinState Me, App.ProductName
    mnuFileText.Checked = GetSetting("ZLSOFT", "公共全局", "文本SQL", False)
    
    txt.Visible = mnuFileText.Checked
    rtf.Visible = Not mnuFileText.Checked
    
    mstrFile = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim varMsg As VbMsgBoxResult
    
    If Not gblnOK And UnloadMode = 0 Then
        varMsg = MsgBox("内容已经更改，要保存吗？", vbQuestion + vbYesNoCancel + vbDefaultButton2, Me.Caption)
        If varMsg = vbCancel Then Cancel = 1: Exit Sub
        If varMsg = vbYes Then SaveSQLTest
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    rtf.Left = Me.ScaleLeft
    rtf.Top = Me.ScaleTop
    rtf.Width = Me.ScaleWidth
    rtf.Height = Me.ScaleHeight

    txt.Left = Me.ScaleLeft
    txt.Top = Me.ScaleTop
    txt.Width = Me.ScaleWidth
    txt.Height = Me.ScaleHeight

    txtObject.Left = Me.ScaleLeft
    txtObject.Top = Me.ScaleTop
    txtObject.Width = Me.ScaleWidth
    txtObject.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Not gobjComLib Is Nothing Then gobjComLib.SaveWinState Me, App.ProductName
    gblnShow = False
End Sub

Private Sub mnuFile_Stop_Click()
    mnuFile_Stop.Checked = mnuFile_Stop.Checked Xor True
    SaveSetting "ZLSOFT", "公共全局", "SQLTest", IIf(mnuFile_Stop.Checked, 1, 0)
End Sub

Private Sub mnuFile_Top_Click()
    mnuFile_Top.Checked = Not mnuFile_Top.Checked
    Call SetTop(mnuFile_Top.Checked)
End Sub

Private Sub mnuFileAnalyse_Click()
    mnuFileAnalyse.Checked = Not mnuFileAnalyse.Checked
    txtObject.Visible = mnuFileAnalyse.Checked
    If mnuFileAnalyse.Checked Then
        txt.Visible = False
        rtf.Visible = False
    Else
        txt.Visible = mnuFileText.Checked
        rtf.Visible = Not mnuFileText.Checked
    End If
End Sub

Private Sub mnuFileClear_Click()
    txt.Text = ""
    rtf.Text = ""
    txtObject.Text = ""
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    cdg.DialogTitle = "打开SQL测试内容"
    If rtf.Visible Then
        cdg.Filter = "RTF文件|*.RTF|文本文件|*.TXT"
    Else
        cdg.Filter = "文本文件|*.TXT"
    End If
    cdg.Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdg.InitDir = GetSetting("ZLSOFT", "公共全局", "SQLTestPath", App.Path)
    cdg.FileName = ""
    cdg.CancelError = True
    On Error Resume Next
    cdg.ShowOpen
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        SaveSetting "ZLSOFT", "公共全局", "SQLTestPath", Left(cdg.FileName, Len(cdg.FileName) - Len(cdg.FileTitle))
        mstrFile = cdg.FileName
        Call rtf.LoadFile(mstrFile, IIf(UCase(mstrFile) Like "*.RTF", rtfRTF, rtfText))
        rtf.SelStart = 0
    End If
    gblnOK = True
End Sub

Private Sub mnuFileSave_Click()
    SaveSQLTest
End Sub

Private Sub SaveSQLTest(Optional blnSaveAs As Boolean)
    If Not blnSaveAs And mstrFile <> "" Then
        Call rtf.SaveFile(mstrFile, IIf(UCase(mstrFile) Like "*.RTF", rtfRTF, rtfText))
    Else
        cdg.DialogTitle = "保存SQL测试内容"
        If rtf.Visible Then
            cdg.Filter = "RTF文件|*.RTF|文本文件|*.TXT"
            cdg.FileName = UCase(gstrDBUser) & Format(Date, "yyMMdd") & ".RTF" '缺省文件名
        Else
            cdg.Filter = "文本文件|*.TXT"
            cdg.FileName = UCase(gstrDBUser) & Format(Date, "yyMMdd") & ".TXT" '缺省文件名
        End If
        cdg.Flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
        cdg.InitDir = GetSetting("ZLSOFT", "公共全局", "SQLTestPath", App.Path)
        
        cdg.CancelError = True
        On Error Resume Next
        cdg.ShowSave
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            SaveSetting "ZLSOFT", "公共全局", "SQLTestPath", Left(cdg.FileName, Len(cdg.FileName) - Len(cdg.FileTitle))
            mstrFile = cdg.FileName
            Call rtf.SaveFile(mstrFile, IIf(UCase(mstrFile) Like "*.RTF", rtfRTF, rtfText))
        End If
    End If
    gblnOK = True
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveSQLTest True
End Sub

Private Sub mnuFileText_Click()
    mnuFileText.Checked = Not mnuFileText.Checked
    txt.Visible = mnuFileText.Checked
    rtf.Visible = Not mnuFileText.Checked
    If txt.Visible Then
        txt.Text = rtf.Text
        txt.SelStart = Len(txt.Text)
        rtf.Text = ""
    Else
        rtf.Text = txt.Text
        rtf.SelStart = Len(rtf.Text)
        txt.Text = ""
    End If
    SaveSetting "ZLSOFT", "公共全局", "文本SQL", mnuFileText.Checked
End Sub
