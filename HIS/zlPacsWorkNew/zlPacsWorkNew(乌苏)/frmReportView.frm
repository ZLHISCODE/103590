VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmReportView 
   BorderStyle     =   0  'None
   Caption         =   "报告所见"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picSigns 
      Height          =   1455
      Left            =   1680
      ScaleHeight     =   1395
      ScaleWidth      =   3075
      TabIndex        =   3
      Top             =   4560
      Width           =   3135
      Begin VB.TextBox txtReview 
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
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtSigns 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   350
         Left            =   690
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblReview 
         Caption         =   "随访："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   650
      End
      Begin VB.Label lblSign 
         Caption         =   "签名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   9
         Top             =   30
         Width           =   650
      End
   End
   Begin VB.PictureBox picAdvice 
      Height          =   1215
      Left            =   3960
      ScaleHeight     =   1155
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   3000
      Width           =   3015
      Begin RichTextLib.RichTextBox rTxtAdvice 
         Height          =   855
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmReportView.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picResult 
      Height          =   1215
      Left            =   600
      ScaleHeight     =   1155
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   2400
      Width           =   3015
      Begin RichTextLib.RichTextBox rtxtResult 
         Height          =   975
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1720
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmReportView.frx":009D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picCheckView 
      Height          =   2175
      Left            =   2520
      ScaleHeight     =   2115
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      Begin RichTextLib.RichTextBox rtxtCheckView 
         Height          =   1935
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3413
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmReportView.frx":013A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblFormat 
         Caption         =   "新报告"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3855
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   600
      Top             =   480
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Private mlngAdviceID As Long    '医嘱ID
'Private mlngSendNo As Long      '发送号
Private mblnSingleWindow As Boolean     '是否使用独立窗口显示报告编辑器，True-独立窗口显示；False-嵌入式显示
Private mReportID As Long         '病历文件id
Private mFileID As Long           '病历模板ID
Private mlngCY1 As Long                 '检查所见的高度
Private mlngCY2 As Long                 '诊断意见的高度
Private mlngCY3 As Long                 '建议的高度
Private mlngCY4 As Long                 '签名的高度
Private mblnCheckModify As Boolean      '是否启动内容变化记录
Private mblnEdiatble As Boolean         '是否可以编辑内容
Private mstrModifyEdit As String        '当前报告是否在修订状态被其他人修订保存后没有签名？记录保存人的姓名，空表示不是这种情况
Private mblnShowWord As Boolean         '显示词句示范，True--显示词句示范；False--双击标题才显示词句示范
Private mblnMoved As Boolean            '是否转储

Public pModified As Boolean          '记录当前内容是否有改变
Private mingFlag As Integer          '为1时说明已经执行过检查所见的GetFocue方法

'本窗体的事件
Public Event CheckViewClick(ByVal strContext As String)
Public Event ResultClick(ByVal strContext As String)
Public Event AdviceClick(ByVal strContext As String)
Public Event ShowWord(intReportViewType As Integer, strContext As String)

Public Sub zlRefreshLblFormat(strFormatInfo As String)
    lblFormat.Caption = strFormatInfo
End Sub

Public Sub zlRefresh(ReportID As Long, blnSingleWindow As Boolean, FileID As Long, _
    blnDeptChanged As Boolean, blnEditable As Boolean, strModifyEdit As String, _
    strInfo As String, blnShowWord As Boolean, strFormatInfo As String, ByVal blnMoved As Boolean)
'参数说明：
'           blnEditable----当前是否可以编辑,True--可编辑；False--不可编辑
'           strModifyEdit----当前报告是否在修订状态被其他人修订保存后没有签名？记录保存人的姓名，空表示不是这种情况
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer

    mReportID = ReportID
    mFileID = FileID
    mblnEdiatble = blnEditable
    mstrModifyEdit = strModifyEdit
    mblnShowWord = blnShowWord
    mblnMoved = blnMoved
    mingFlag = 0
    lblFormat.Caption = strFormatInfo
    
    rtxtCheckView.Text = ""
    rtxtResult.Text = ""
    rTxtAdvice.Text = ""
    rtxtCheckView.Tag = ""
    rtxtResult.Tag = ""
    rTxtAdvice.Tag = ""
    
    txtInfo.Visible = blnSingleWindow
    If blnSingleWindow Then txtInfo.Text = strInfo
    
    mblnCheckModify = False         '关闭内容变化记录
    pModified = False

    If mblnSingleWindow <> blnSingleWindow Then
        mblnSingleWindow = blnSingleWindow
        Call InitLoaclParas     '读取本机参数
        Call InitFaceScheme     '初始界面布局
    End If
    
    If blnDeptChanged = True Then
        For i = 1 To dkpMain.PanesCount
            Select Case dkpMain.Panes(i).Tag
            Case 0
                dkpMain.Panes(i).Title = pReport_CheckViewName
            Case 1
                dkpMain.Panes(i).Title = pReport_ResultName
            Case 2
                dkpMain.Panes(i).Title = pReport_AdviceName
            End Select
        Next i
    End If
    
    '根据病历文件ID，初始化三个文本编辑器
    '查找报告单模板ID
    If mReportID = 0 Then       '报告为空，需要创建报告，从报告模板中提取信息
        strSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文 " & _
                 " From 病历文件结构 a, 病历文件结构 b" & _
                 " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父id And b.对象类型 = 2 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mFileID)
    Else
        strSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文 From 电子病历内容 a,电子病历内容 b " & _
                 " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 And b.终止版 = 0"
        If mblnMoved = True Then
            strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
    End If
    While rsTemp.EOF = False
        Select Case Nvl(rsTemp!标题)
            Case "检查所见"
                rtxtCheckView.Tag = Nvl(rsTemp!对象属性)
                zlWriteReport Nvl(rsTemp!正文), 0
            Case "诊断意见"
                rtxtResult.Tag = Nvl(rsTemp!对象属性)
                zlWriteReport Nvl(rsTemp!正文), 1
            Case "建议"
                rTxtAdvice.Tag = Nvl(rsTemp!对象属性)
                zlWriteReport Nvl(rsTemp!正文), 2
        End Select
        rsTemp.MoveNext
    Wend
    
    If mblnSingleWindow Then
        RaiseEvent CheckViewClick("")
    Else
        On Error GoTo errH
        If rtxtCheckView.Visible Then rtxtCheckView.SetFocus
errH:
    End If
    
    '设置界面控件是否可以编辑
    rtxtCheckView.Locked = Not mblnEdiatble
    rtxtResult.Locked = Not mblnEdiatble
    rTxtAdvice.Locked = Not mblnEdiatble
    
    rtxtCheckView.BackColor = IIf(rtxtCheckView.Locked, &H8000000F, &H80000005)
    rtxtResult.BackColor = IIf(rtxtResult.Locked, &H8000000F, &H80000005)
    rTxtAdvice.BackColor = IIf(rTxtAdvice.Locked, &H8000000F, &H80000005)
    
    rtxtCheckView.ToolTipText = IIf(mblnEdiatble = False And mstrModifyEdit <> "", "本报告已经由" & mstrModifyEdit & "正在修订，现在是否要进行修订？需要修订请双击。", "")
    rtxtResult.ToolTipText = rtxtCheckView.ToolTipText
    rTxtAdvice.ToolTipText = rtxtCheckView.ToolTipText
    
'    picCheckView.Enabled = mblnEdiatble
'    picResult.Enabled = mblnEdiatble
'    picAdvice.Enabled = mblnEdiatble
    
    mblnCheckModify = True      '内容装载完毕，启动内容变化记录
End Sub

Public Sub zlWriteReport(strText As String, intType As Integer)
    'intType---0 检查所见；1 诊断意见；2 建议
    Dim rText As RichTextBox
    Dim lngCount As Long
    Dim lngSelStart As Long
    Dim lngPosStart As Long
    Dim lngPosEnd As Long
    
    On Error GoTo err
    
    If intType = 0 Then
        Set rText = rtxtCheckView
    ElseIf intType = 1 Then
        Set rText = rtxtResult
    ElseIf intType = 2 Then
        Set rText = rTxtAdvice
    End If
    
    lngSelStart = rText.SelStart
    rText.SelLength = 0
    rText.SelText = strText
    '设置颜色
    rText.SelStart = lngSelStart
    rText.SelLength = Len(strText)
    rText.SelColor = vbBlack
    
    On Error Resume Next
    'rText.Tag 是电子病历格式的对象属性，用“|”分隔，总共26个元素
    rText.SelStart = 0
    rText.SelLength = Len(rText.Text)
    rText.SelFontName = Split(rText.Tag, "|")(15)     '  rText.SelFontName
    rText.SelFontSize = Split(rText.Tag, "|")(16)     ' rText.SelFontSize
    rText.SelBold = Split(rText.Tag, "|")(17)     'rText.SelBold
    rText.SelItalic = Split(rText.Tag, "|")(18)   'rText.SelItalic
    On Error GoTo 0
    
    '解析当前输入的文字，是否有要素，如果有则用蓝色表示出来
    '先查多选要素
    For lngCount = 1 To Len(strText)
        lngPosStart = InStr(lngCount, strText, "{{")
        lngPosEnd = InStr(lngCount, strText, "}}")
        If lngPosStart <> 0 And lngPosEnd <> 0 And lngPosEnd > lngPosStart Then
            '查找到要素，则对要素做蓝色显示
            rText.SelStart = lngSelStart + lngPosStart - 1
            rText.SelLength = lngPosEnd - lngPosStart + 2
            rText.SelColor = vbBlue
            lngCount = lngPosEnd
        Else
            Exit For
        End If
    Next lngCount
    
    '再查单选要素
    For lngCount = 1 To Len(strText)
        lngPosStart = InStr(lngCount, strText, "{<")
        lngPosEnd = InStr(lngCount, strText, ">}")
        If lngPosStart <> 0 And lngPosEnd <> 0 And lngPosEnd > lngPosStart Then
            '查找到要素，则对要素做蓝色显示
            rText.SelStart = lngSelStart + lngPosStart - 1
            rText.SelLength = lngPosEnd - lngPosStart + 2
            rText.SelColor = vbBlue
            lngCount = lngPosEnd
        Else
            Exit For
        End If
    Next lngCount
    
    rText.SelStart = lngSelStart + Len(strText)
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Dim strContext As String
    
    Select Case Pane.Tag
        Case 0
            strContext = rtxtCheckView.Text
        Case 1
            strContext = rtxtResult.Text
        Case 2
            strContext = rTxtAdvice.Text
    End Select
    
    If Action = PaneActionFloating And mblnShowWord = False Then
'        frmReportWord.Show 1, Me
    
        '触发事件，显示词句示范窗口
        RaiseEvent ShowWord(Pane.Tag, strContext)
    End If
    Cancel = True
End Sub

Private Sub Form_Load()
    mingFlag = 0
    pModified = False
    mblnSingleWindow = False    '默认设置为嵌入式窗体
    
    Call InitLoaclParas     '读取本机参数
    Call InitFaceScheme     '初始界面布局
End Sub

Private Sub InitLoaclParas()
    Dim strRegPath As String
    
    '读取检查所见区域，诊断意见区域，建议区域 和签名区域的高度
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReportView\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReportView"
    End If
    mlngCY1 = GetSetting("ZLSOFT", strRegPath, "CY1", 500)
    mlngCY2 = GetSetting("ZLSOFT", strRegPath, "CY2", 200)
    mlngCY3 = GetSetting("ZLSOFT", strRegPath, "CY3", 100)
    mlngCY4 = GetSetting("ZLSOFT", strRegPath, "CY4", 100)
End Sub

Private Sub InitFaceScheme()
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane
    With Me.dkpMain
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 0, mlngCY1, DockTopOf, Nothing)
    Pane1.Title = pReport_CheckViewName
    Pane1.Handle = picCheckView.hWnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane1.Tag = 0
    
    Set Pane2 = dkpMain.CreatePane(2, 0, mlngCY2, DockBottomOf, Pane1)
    Pane2.Title = pReport_ResultName
    Pane2.Handle = picResult.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane2.Tag = 1
    
    Set Pane3 = dkpMain.CreatePane(3, 0, mlngCY3, DockBottomOf, Pane2)
    Pane3.Title = pReport_AdviceName
    Pane3.Handle = picAdvice.hWnd
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane3.Tag = 2
    
    Set pane4 = dkpMain.CreatePane(4, 0, mlngCY4, DockBottomOf, Pane3)
    pane4.Title = "签名"
    pane4.Handle = picSigns.hWnd
    pane4.Options = PaneNoCaption Or PaneNoCloseable
    pane4.Tag = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReportView\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReportView"
    End If
    '保存检查所见区域，诊断意见区域，建议区域和签名区域的高度
    '285是Pane的标题高度，使用了标题，就需要加回这个高度
    SaveSetting "ZLSOFT", strRegPath, "CY1", picCheckView.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "CY2", picResult.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "CY3", picAdvice.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "CY4", picSigns.Height
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
    End If
    SaveSetting "ZLSOFT", strRegPath, "CX2", Me.Width
    SaveSetting "ZLSOFT", strRegPath, "CY21", Me.Height
End Sub

Private Sub picAdvice_Resize()
On Error Resume Next

    rTxtAdvice.Left = 20
    rTxtAdvice.Top = 20
    If picAdvice.Height > 50 And picAdvice.Width > 50 Then
        rTxtAdvice.Width = Abs(picAdvice.Width - 100)
        rTxtAdvice.Height = Abs(picAdvice.Height - 100)
    End If
End Sub

Private Sub picCheckView_Resize()
On Error Resume Next

    lblFormat.Left = 10
    lblFormat.Top = 10
    lblFormat.Width = picCheckView.Width
    lblFormat.Height = 400
    
    rtxtCheckView.Left = 20
    rtxtCheckView.Top = lblFormat.Height
    If picCheckView.Width > 50 And picCheckView.Height > 50 Then
        rtxtCheckView.Width = Abs(picCheckView.Width - 100)
        rtxtCheckView.Height = Abs(picCheckView.Height - 100 - lblFormat.Height)
    End If
End Sub

Private Sub picResult_Resize()
On Error Resume Next

    rtxtResult.Left = 20
    rtxtResult.Top = 20
    If picResult.Width > 50 And picResult.Height > 50 Then
        rtxtResult.Width = Abs(picResult.Width - 100)
        rtxtResult.Height = Abs(picResult.Height - 100)
    End If
End Sub

Private Sub picSigns_Resize()
On Error Resume Next

    lblSign.Left = 0
    txtSigns.Left = lblSign.Width
    txtSigns.Top = 30
    txtSigns.Width = Abs(picSigns.ScaleWidth - lblSign.Width - 50)
    
    lblReview.Left = 0
    txtReview.Left = lblSign.Width
    txtReview.Top = 360
    txtReview.Width = txtSigns.Width
    
    txtInfo.Left = 10
    txtInfo.Top = txtReview.Top + txtReview.Height + 10
    txtInfo.Width = Abs(picSigns.ScaleWidth - 50)
    txtInfo.Height = Abs(picSigns.ScaleHeight - txtSigns.Height - txtReview.Height - 50)
End Sub

Private Sub rTxtAdvice_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub rTxtAdvice_DblClick()
    If mblnEdiatble = False And mstrModifyEdit <> "" Then
        rTxtAdvice.Locked = False
        rTxtAdvice.ToolTipText = ""
    Else
        Call richTextBoxShowElements(rTxtAdvice)
    End If
End Sub

Private Sub rTxtAdvice_GotFocus()
On Error GoTo err
    If gblnIsStudyChage Then
        If rtxtCheckView.Visible Then rtxtCheckView.SetFocus '切换检查后，定位到检查所见文本框，具体见问题81704
        gblnIsStudyChage = False
        Exit Sub
    End If
    mingFlag = 0
    
    Call zlCommFun.OpenIme(True)
    RaiseEvent AdviceClick(rTxtAdvice.Text)
err:
End Sub
 

Private Sub rtxtCheckView_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub rtxtCheckView_DblClick()
    If mblnEdiatble = False And mstrModifyEdit <> "" Then
        rtxtCheckView.Locked = False
        rtxtCheckView.ToolTipText = ""
    Else
        Call richTextBoxShowElements(rtxtCheckView)
    End If
End Sub

Private Sub rtxtCheckView_GotFocus()
    If mingFlag = 1 Then Exit Sub
    
    mingFlag = 1
    Call zlCommFun.OpenIme(True)
    RaiseEvent CheckViewClick(rtxtCheckView.Text)
End Sub
 
Private Sub rtxtResult_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub rtxtResult_DblClick()
    If mblnEdiatble = False And mstrModifyEdit <> "" Then
        rtxtResult.Locked = False
        rtxtResult.ToolTipText = ""
    Else
        Call richTextBoxShowElements(rtxtResult)
    End If
End Sub

Private Sub rtxtResult_GotFocus()
On Error GoTo err
    If gblnIsStudyChage Then
        If rtxtCheckView.Visible Then rtxtCheckView.SetFocus '切换检查后，定位到检查所见文本框，具体见问题81704
        gblnIsStudyChage = False
        Exit Sub
    End If
    mingFlag = 0
    
    Call zlCommFun.OpenIme(True)
    RaiseEvent ResultClick(rtxtResult.Text)
err:
End Sub
 

Private Sub txtReview_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txtReview_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub
 
