VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmScheme_BaseQueryCfg 
   BorderStyle     =   0  'None
   Caption         =   "基础查询设置"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCheck 
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2835
      ScaleWidth      =   9795
      TabIndex        =   6
      Top             =   240
      Width           =   9855
      Begin VB.CommandButton cmdSQLVerify 
         Caption         =   "SQL验证"
         Height          =   350
         Left            =   8520
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdSQLInsert 
         Caption         =   "插入参数"
         Height          =   350
         Left            =   8520
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame fraCheckSQL 
         Height          =   30
         Left            =   1800
         TabIndex        =   8
         Top             =   120
         Width           =   6255
      End
      Begin VB.TextBox txtHinCfg 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   2040
         Width           =   5775
      End
      Begin RichTextLib.RichTextBox rtbCheckSQL 
         Height          =   1650
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   2910
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         TextRTF         =   $"frmScheme_BaseQueryCfg.frx":0000
      End
      Begin VB.Label lblCheckSQL 
         AutoSize        =   -1  'True
         Caption         =   "数据查询语句"
         Height          =   180
         Left            =   720
         TabIndex        =   13
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label lblHintCfg 
         AutoSize        =   -1  'True
         Caption         =   "历史库查询HINT配置"
         Height          =   180
         Left            =   480
         TabIndex        =   12
         Top             =   2160
         Width           =   1620
      End
   End
   Begin VB.PictureBox picDetail 
      Height          =   3015
      Left            =   240
      ScaleHeight     =   2955
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   3240
      Width           =   9855
      Begin VB.CommandButton cmdDetailVerify 
         Caption         =   "SQL验证"
         Height          =   350
         Left            =   8520
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdDetailInsert 
         Caption         =   "插入参数"
         Height          =   350
         Left            =   8520
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.Frame fraDetail 
         Height          =   30
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   7215
      End
      Begin RichTextLib.RichTextBox rtbDetail 
         Height          =   2295
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4048
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         TextRTF         =   $"frmScheme_BaseQueryCfg.frx":009D
      End
      Begin VB.Label lblDetail 
         AutoSize        =   -1  'True
         Caption         =   "数据明细语句"
         Height          =   180
         Left            =   480
         TabIndex        =   5
         Top             =   120
         Width           =   1080
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   120
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmScheme_BaseQueryCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnIsEdit As Boolean
Public mintVerify As Integer    '数据查询语句，0-未验证；1-验证失败；2-验证成功
Private mstrQuerySql As String

Public Event DoCheckVerify(ByVal strHint As String)    '数据查询语句验证

Private Sub cmdDetailInsert_Click()
'插入参数
On Error GoTo errHandle
    Dim strPar As String
    Dim frmPar As New frmSetPara
    
    strPar = frmPar.ShowParameterWindow(False, Me, , 2)
    If strPar <> "" Then
        rtbDetail.SelText = strPar
    End If
    
    Unload frmPar
    Set frmPar = Nothing
    
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
    
    If Not frmPar Is Nothing Then Unload frmPar
    Set frmPar = Nothing
End Sub

Private Sub cmdDetailVerify_Click()
    Dim strResult As String

    On Error GoTo errHandle
    
    RaiseEvent DoCheckVerify("正在进行查询验证...（如果时间过长，请检查查询语句）")
    
    strResult = SqlVerify(rtbDetail.Text, True)
    
    RaiseEvent DoCheckVerify("")
    
    If Len(strResult) = 0 Then
        MsgBox "验证成功！", vbInformation, Me.Caption
    Else
        MsgBox "验证失败，原因是：" & strResult, vbInformation, Me.Caption
        rtbDetail.SetFocus
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdSQLInsert_Click()
'插入参数
    Dim strPar As String
    Dim frmPar As New frmSetPara
    
    On Error GoTo errHandle
    
    strPar = frmPar.ShowParameterWindow(False, Me)
    If strPar <> "" Then
        rtbCheckSQL.SelText = strPar
    End If
    
    Unload frmPar
    Set frmPar = Nothing
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
    
    If Not frmPar Is Nothing Then Unload frmPar
    Set frmPar = Nothing
End Sub

Private Sub cmdSqlVerify_Click()
    Dim strResult As String
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As Recordset
    Dim strItem As String
    Dim strPar As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    RaiseEvent DoCheckVerify("正在进行查询验证...（如果时间过长，请检查查询语句）")
    
    strResult = SqlVerify(rtbCheckSQL.Text)
    
    If Len(strResult) = 0 Then
        strResult = IsHaveID(rtbCheckSQL.Text)
    End If
    
    RaiseEvent DoCheckVerify("")
    If Len(strResult) = 0 Then
        mintVerify = 2
        MsgBox "验证成功！", vbInformation, Me.Caption
    Else
        mintVerify = 1
        MsgBox "验证失败，原因是：" & strResult, vbInformation, Me.Caption
        rtbCheckSQL.SetFocus
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Load()
    Call InitDockPannel
    Call RefreshWindowState(False)
    Call SetFontSize(gbytFontSize)
End Sub

Public Sub UnloadMe()
    Unload Me
End Sub


Public Sub ShowQuerySet(objSqlScheme As clsSqlScheme)
    rtbCheckSQL.Text = objSqlScheme.Query
    rtbDetail.Text = objSqlScheme.Detail
    txtHinCfg.Text = objSqlScheme.HistoryDBHint
End Sub

Private Sub picCheck_Resize()
    On Error Resume Next
  
    '查询语句部分
    lblCheckSQL.Move picCheck.Left + 100, picCheck.Top + 100
    fraCheckSQL.Move lblCheckSQL.Left + lblCheckSQL.Width, lblCheckSQL.Top + lblCheckSQL.Height / 2, picCheck.ScaleWidth - fraCheckSQL.Left
    rtbCheckSQL.Move picCheck.ScaleLeft + 100, fraCheckSQL.Top + 200, picCheck.ScaleWidth - 300 - cmdSQLInsert.Width, picCheck.ScaleHeight - lblHintCfg.Height - 750
    cmdSQLInsert.Move rtbCheckSQL.Left + rtbCheckSQL.Width + 100, rtbCheckSQL.Top
    cmdSQLVerify.Move cmdSQLInsert.Left, cmdSQLInsert.Top + cmdSQLInsert.Height + 200

    lblHintCfg.Move rtbCheckSQL.Left, rtbCheckSQL.Top + rtbCheckSQL.Height + 200
    txtHinCfg.Move lblHintCfg.Left + lblHintCfg.Width + 60, rtbCheckSQL.Top + rtbCheckSQL.Height + 100, rtbCheckSQL.Width - lblHintCfg.Width - 60
End Sub

Private Sub picDetail_Resize()
    On Error Resume Next
    
    '检查明细部分
    lblDetail.Move picDetail.Left + 100, picDetail.Top + 100
    fraDetail.Move lblDetail.Left + lblDetail.Width, lblDetail.Top + lblDetail.Height / 2, picDetail.ScaleWidth - fraDetail.Left
    rtbDetail.Move rtbCheckSQL.Left, fraDetail.Top + 200, rtbCheckSQL.Width, picDetail.Height - 800
    cmdDetailInsert.Move cmdSQLInsert.Left, rtbDetail.Top
    cmdDetailVerify.Move cmdSQLInsert.Left, cmdDetailInsert.Top + cmdDetailInsert.Height + 200
End Sub

Private Sub rtbCheckSQL_Change()
    mstrQuerySql = rtbCheckSQL.Text
    mblnIsEdit = True
    mintVerify = 0
    
    If cmdSQLInsert.Enabled Then
        If Len(Trim(rtbCheckSQL.Text)) = 0 Then
            cmdSQLVerify.Enabled = False
        Else
            cmdSQLVerify.Enabled = True
        End If
    Else
        cmdSQLVerify.Enabled = False
    End If
End Sub

Public Function GetQuerySql() As String
    GetQuerySql = mstrQuerySql
End Function

Public Sub RefreshWindowState(blnState As Boolean)
    rtbCheckSQL.Locked = Not blnState
    rtbDetail.Locked = Not blnState
    txtHinCfg.Enabled = blnState
    
    If blnState Then
        rtbCheckSQL.BackColor = &H80000005
        rtbDetail.BackColor = &H80000005
    Else
        rtbCheckSQL.BackColor = &H8000000F
        rtbDetail.BackColor = &H8000000F
    End If
    
    cmdDetailInsert.Enabled = blnState
    cmdSQLInsert.Enabled = blnState
    
    If Len(Trim(rtbCheckSQL.Text)) = 0 Or Not blnState Then
        cmdSQLVerify.Enabled = False
    Else
        cmdSQLVerify.Enabled = True
    End If
    
    If Len(Trim(rtbDetail.Text)) = 0 Or Not blnState Then
        cmdDetailVerify.Enabled = False
    Else
        cmdDetailVerify.Enabled = True
    End If
End Sub

Public Sub SetQueryCfg(objSqlScheme As clsSqlScheme)
    objSqlScheme.Query = rtbCheckSQL.Text
    objSqlScheme.Detail = rtbDetail.Text
    objSqlScheme.HistoryDBHint = Trim(txtHinCfg.Text)
End Sub

Private Sub rtbDetail_Change()
    mblnIsEdit = True
    If cmdDetailInsert.Enabled Then
        If Len(Trim(rtbDetail.Text)) = 0 Then
            cmdDetailVerify.Enabled = False
        Else
            cmdDetailVerify.Enabled = True
        End If
    Else
        cmdDetailVerify.Enabled = False
    End If
End Sub

Public Function IsEnabledToSave() As Boolean
    IsEnabledToSave = False
    
    If Len(Trim(rtbCheckSQL.Text)) = 0 Then
        Exit Function
    End If
    
    IsEnabledToSave = True
End Function

Public Sub rtbCheckSQLSetFocus()
    rtbCheckSQL.SetFocus
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    Dim lngCmdHeight As Long
    Dim lngCmdWithd As Long
    
    If bytFontSize = 9 Then
        lngCmdHeight = 350
        lngCmdWithd = 1100
        rtbCheckSQL.Width = 7455
    ElseIf bytFontSize = 12 Then
        lngCmdHeight = 385
        lngCmdWithd = 1300
        rtbCheckSQL.Width = 7255
    ElseIf bytFontSize = 15 Then
        lngCmdHeight = 420
        lngCmdWithd = 1500
        rtbCheckSQL.Width = 7055
    End If
    
    rtbCheckSQL.Font.Size = bytFontSize
    rtbDetail.Font.Size = bytFontSize
    lblCheckSQL.FontSize = bytFontSize
    lblDetail.FontSize = bytFontSize
    txtHinCfg.FontSize = bytFontSize
    lblHintCfg.FontSize = bytFontSize
    
    cmdSQLInsert.FontSize = bytFontSize
    cmdSQLInsert.Height = lngCmdHeight
    cmdSQLInsert.Width = lngCmdWithd
    cmdSQLVerify.FontSize = bytFontSize
    cmdSQLVerify.Height = lngCmdHeight
    cmdSQLVerify.Width = lngCmdWithd
    cmdDetailInsert.FontSize = bytFontSize
    cmdDetailInsert.Height = lngCmdHeight
    cmdDetailInsert.Width = lngCmdWithd
    cmdDetailVerify.FontSize = bytFontSize
    cmdDetailVerify.Height = lngCmdHeight
    cmdDetailVerify.Width = lngCmdWithd
    
    Call picCheck_Resize
    Call picDetail_Resize
End Sub

Private Sub txtHinCfg_Change()
    mblnIsEdit = True
End Sub

'布局绑定
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo errHandle
    
    Select Case Item.Id
        Case 1
            Item.Handle = picCheck.hwnd
        Case 2
            Item.Handle = picDetail.hwnd
    End Select
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

'界面布局
Private Sub InitDockPannel()
    Dim objPane As Pane
    
    On Error GoTo errHandle
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockTopOf, Nothing)
    objPane.Title = "picCheck"
    objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, objPane)
    objPane.Title = "picDetail"
    objPane.Options = PaneNoCaption
    
    Set objPane = Nothing
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub
