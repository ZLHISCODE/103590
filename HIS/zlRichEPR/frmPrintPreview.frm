VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintPreview 
   Caption         =   "打印预览"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   9345
   Icon            =   "frmPrintPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9345
   StartUpPosition =   1  '所有者中心
   Begin zlSubclass.Subclass Subclass1 
      Left            =   4365
      Top             =   1215
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSComctlLib.ImageList imlPages 
      Left            =   4095
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox picZoom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   3375
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   11
      Top             =   45
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.VScrollBar VS 
      DragIcon        =   "frmPrintPreview.frx":038A
      Height          =   2145
      LargeChange     =   20
      Left            =   8955
      Max             =   100
      MouseIcon       =   "frmPrintPreview.frx":0694
      SmallChange     =   10
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2070
      Width           =   250
   End
   Begin VB.HScrollBar HS 
      DragIcon        =   "frmPrintPreview.frx":07E6
      Height          =   250
      LargeChange     =   20
      Left            =   2835
      Max             =   100
      MouseIcon       =   "frmPrintPreview.frx":0AF0
      SmallChange     =   10
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4410
      Width           =   6105
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   2835
      ScaleHeight     =   2145
      ScaleWidth      =   6060
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2250
      Width           =   6060
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   135
         MouseIcon       =   "frmPrintPreview.frx":0C42
         MousePointer    =   99  'Custom
         ScaleHeight     =   930
         ScaleWidth      =   5790
         TabIndex        =   7
         Top             =   180
         Width           =   5820
         Begin VB.PictureBox picBlank 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   40
            Left            =   0
            MouseIcon       =   "frmPrintPreview.frx":0D94
            MousePointer    =   99  'Custom
            ScaleHeight     =   45
            ScaleWidth      =   825
            TabIndex        =   13
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   225
         ScaleHeight     =   960
         ScaleWidth      =   5820
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   5820
      End
   End
   Begin VB.PictureBox pic页面 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   135
      ScaleHeight     =   2220
      ScaleWidth      =   2310
      TabIndex        =   4
      Top             =   1350
      Visible         =   0   'False
      Width           =   2310
      Begin VSFlex8Ctl.VSFlexGrid vfg页面 
         Height          =   1695
         Left            =   45
         TabIndex        =   5
         Top             =   135
         Width           =   2100
         _cx             =   3704
         _cy             =   2990
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   10197915
         ForeColor       =   -2147483640
         BackColorFixed  =   10197915
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483634
         BackColorBkg    =   10197915
         BackColorAlternate=   10197915
         GridColor       =   10197915
         GridColorFixed  =   10197915
         TreeColor       =   16777215
         FloodColor      =   192
         SheetBorder     =   10197915
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPrintPreview.frx":0FA6
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   7
         PicturesOver    =   -1  'True
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   1
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picBuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   2655
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   645
   End
   Begin zlRichEditor.Editor edtThis 
      Height          =   600
      Left            =   900
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   1058
      Title           =   ""
      WithViewButtonas=   0   'False
      ShowRuler       =   0   'False
   End
   Begin XtremeSuiteControls.TabControl tabThis 
      Height          =   1230
      Left            =   90
      TabIndex        =   2
      Top             =   810
      Width           =   2595
      _Version        =   589884
      _ExtentX        =   4577
      _ExtentY        =   2170
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7065
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPrintPreview.frx":0FE9
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13573
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
   Begin zlRichEditor.Editor edtBuff 
      Height          =   600
      Left            =   1620
      TabIndex        =   12
      Top             =   45
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   1058
      Title           =   ""
      WithViewButtonas=   0   'False
      ShowRuler       =   0   'False
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   4770
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   1
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event PrintEpr(ByVal lngRecordId As Long)

'文件 "File"
Private Const ID_File_SaveCopy = 302    '保存副本(A)...
Private Const ID_File_SaveTxt = 303     '保存为文本(V)...
Private Const ID_FILE_PRINT = 304       '打印(P)...
Private Const ID_FILE_EXIT = 305        '退出(X)
Private Const ID_File_SaveAsPic = 306   '另存为图片(I)
Private Const ID_FILE_PRINTINWORD = 307 '在Word中打印(W)

'视图 "View"
Private Const ID_View_ToolBar = 310     '工具栏(T)
Private Const ID_View_StatusBar = 311   '状态栏(S)
Private Const ID_View_ZoomFactor = 312  '缩放比例(C)
Private Const ID_View_First = 313       '第一页
Private Const ID_View_Prev = 314        '前一页
Private Const ID_View_Next = 315        '后一页
Private Const ID_View_Last = 316        '最后一页
Private Const ID_View_ActualSize = 317  '实际大小 Ctrl+1
Private Const ID_View_FitSize = 318     '适合页面 Ctrl+2
Private Const ID_View_FitWidth = 319    '适合宽度 Ctrl+3
Private Const ID_View_FitHeight = 320   '适合高度 Ctrl+4
Private Const ID_View_Size_250 = 330    '250%
Private Const ID_View_Size_200 = 331    '200%
Private Const ID_View_Size_150 = 332    '150%
Private Const ID_View_Size_100 = 333    '100%
Private Const ID_View_Size_75 = 334     '75%
Private Const ID_View_Size_50 = 335     '50%
Private Const ID_View_Size_25 = 336     '25%
Private Const ID_View_ZoomIn = 337      '放大
Private Const ID_View_ZoomOut = 338     '缩小

Private Const ID_View_StartPage = 340   '起始页面

'帮助 "Help"
Private Const ID_HELP_CONTENT = 500     '帮助主题
Private Const ID_HELP_CONTACT = 502     '发送反馈
Private Const ID_HELP_ONLINE = 503      '在线医业
Private Const ID_HELP_ABOUT = 504       '关于...

Private cboStartPage As CommandBarComboBox  '起始页面

'文件信息结构体
Private Type FileInfo
    ID As Long              'ID
    FileType As String      '病历种类
    PatiID   As Long        '病人ID
    PageID As Long          '主页ID
End Type

Private Files() As FileInfo             '多文档打印时的文件信息

Private mfrmTipInfo As New frmTipInfo
Private m_bSubClassing As Boolean
Private lngX As Long, lngY As Long, bMouseDown As Boolean
Private mlngPageCount As Long           '总页数
Private mlngCurPage As Long             '当前页
Private ZoomFactor As Double            '缩放比例
Private Const Shadow_W = 60             '阴影厚度

Private mlngCopies As Long              '控制打印份数＝0不指定，不控制，否则控制打印份数不能修改；并通过参数返回打印份数
Private mlngStartPage As Long           '起始页面
Private mlngBlankHeight As Long         '起始页面上部留白高度

'################################################################################################################
'## 功能：  获取系统默认临时路径
'################################################################################################################
Private Function GetSysTmpPath() As String
    GetSysTmpPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
End Function

Private Sub ReplaceKey(ByVal WordApp As Object)
Dim Fd As Object
    WordApp.Selection.Start = 0
    WordApp.Selection.End = 99999
    If WordApp.Selection.Find.Execute("{页码}") Then
        WordApp.Selection.Start = 99999
        Set Fd = WordApp.Selection.Fields.Add(Range:=WordApp.Selection.Range, Type:=33)   'wdFieldPage
        Fd.Copy
        
        WordApp.Selection.Start = 0
        WordApp.Selection.End = 99999
        WordApp.Selection.Find.Execute FindText:="{页码}", ReplaceWith:="^c"
        Fd.Cut
    End If
    
    WordApp.Selection.Start = 0
    WordApp.Selection.End = 99999
    If WordApp.Selection.Find.Execute("{总页数}") Then
        WordApp.Selection.Start = 99999
        Set Fd = WordApp.Selection.Fields.Add(Range:=WordApp.Selection.Range, Type:=26)   'wdFieldNumPages
        Fd.Copy
        
        WordApp.Selection.Start = 0
        WordApp.Selection.End = 99999
        WordApp.Selection.Find.Execute FindText:="{总页数}", ReplaceWith:="^c"
        Fd.Cut
    End If
    Set Fd = Nothing
End Sub
'################################################################################################################
'## 功能：  导出为RTF，然后通过Word打印当前病历
'################################################################################################################
Private Sub PrintInWord()
    On Error Resume Next
    Dim strF As String, strPicFile As String
    strF = GetSysTmpPath & "\PrintInWord_TMP" & App.ThreadID & ".rtf"
    If gobjFSO.FileExists(strF) Then gobjFSO.DeleteFile strF, True
    
    '更改所有左对齐为两端对齐
    Dim i As Long, j As Long
    edtThis.ForceEdit = True
    Do
        i = InStr(i + 1, edtThis.Text, vbCrLf)
        If i > 0 Then
            If edtThis.TOM.TextDocument.Range(i - 2, i - 2).Para.Alignment = tomAlignLeft Then
                edtThis.TOM.TextDocument.Range(i - 2, i - 2).Para.Alignment = tomAlignJustify
            End If
        End If
    Loop Until i <= 0
    edtThis.ForceEdit = False

    edtThis.SaveDoc strF
    If gobjFSO.FileExists(strF) Then
        Dim WordApp As Object   'Word.Application
        Dim WordDoc As Object   'Word.Document
        Set WordApp = CreateObject("Word.Application")
        Set WordDoc = WordApp.Documents.Open(strF)      '打开RTF文档
        
        If WordApp Is Nothing Then
            MsgBox "无法创建Word对象，请安装 Microsoft Office Word 产品！", vbOKOnly + vbInformation, gstrSysName
        Else
            zlCommFun.ShowFlash "请稍候..."
            Screen.MousePointer = vbHourglass
            
            WordApp.Visible = False
            WordApp.ScreenUpdating = False
            '页面大小设置
            WordDoc.PageSetup.LeftMargin = Me.ScaleX(edtThis.MarginLeft, vbTwips, vbPoints)
            WordDoc.PageSetup.RightMargin = Me.ScaleX(edtThis.MarginRight, vbTwips, vbPoints)
            WordDoc.PageSetup.TopMargin = Me.ScaleY(edtThis.MarginTop, vbTwips, vbPoints)
            WordDoc.PageSetup.BottomMargin = Me.ScaleY(edtThis.MarginBottom, vbTwips, vbPoints)
            WordDoc.PageSetup.PageWidth = Me.ScaleX(edtThis.PaperWidth, vbTwips, vbPoints)
            WordDoc.PageSetup.PageHeight = Me.ScaleY(edtThis.PaperHeight, vbTwips, vbPoints)
            
            If WordApp.ActiveWindow.ActivePane.View.Type = 1 Or WordApp.ActiveWindow.ActivePane.View.Type = 2 Then
                WordApp.ActiveWindow.ActivePane.View.Type = 3
                'wdNormalView=1     wdOutlineView=2     wdPrintView=3
            End If
            WordApp.ActiveWindow.View = 5   'wdMasterView
            
            '添加当前的页眉页脚到RTF文件中
            WordApp.ActiveWindow.View.SeekView = 9  'wdSeekCurrentPageHeader'页眉
            WordApp.Selection.ParagraphFormat.Alignment = 0     'wdAlignParagraphLeft
            If Not (edtThis.Picture Is Nothing) Then
                If edtThis.Picture.Handle <> 0 Then
                    strPicFile = GetSysTmpPath & "\zlDocHead" & App.ThreadID & ".BMP"
                    If gobjFSO.FileExists(strPicFile) Then gobjFSO.DeleteFile strPicFile, True
                    SavePicture edtThis.Picture, strPicFile
                    If gobjFSO.FileExists(strPicFile) Then
                        WordApp.Selection.InlineShapes.AddPicture Filename:=strPicFile, LinkToFile:=False, SaveWithDocument:=True
                        gobjFSO.DeleteFile strPicFile, True
                        WordApp.Selection.TypeParagraph
                    End If
                End If
            End If
            
            edtThis.DocHeadReplaceKey
            edtThis.DocHeadCopyWithFormat
            WordApp.Selection.Paste
            '去掉 其中的总页数,页码
            ReplaceKey WordApp
            
            WordApp.ActiveWindow.View.SeekView = 10 'wdSeekCurrentPageFooter'页脚
            edtThis.DocFootReplaceKey
            edtThis.DocFootCopyWithFormat
            WordApp.Selection.Paste
            '去掉 其中的总页数,页码
            ReplaceKey WordApp
            
            WordApp.ActiveWindow.View.SeekView = 3      'wdPrintView
            WordDoc.PrintPreview
            WordApp.ScreenUpdating = True
            WordApp.Visible = True
            WordApp.Activate
            
            Do
                DoEvents
                If Not WordDoc.Windows.Item(WordDoc.Windows.Count).View = 4 Then Exit Do    'wdPrintPreview=4
            Loop
            
            zlCommFun.StopFlash
            Screen.MousePointer = vbDefault
        End If
        
        WordDoc.Close False
        WordApp.Quit
        Set WordDoc = Nothing
        Set WordApp = Nothing
        gobjFSO.DeleteFile strF, True
    End If
End Sub

'################################################################################################################
'## 功能：  显示单文件打印预览窗体
'##
'## 参数：  frmParent       ：父窗体
'##         lng病人ID       ：病人ID
'##         lng主页ID       ：主页ID
'##         lng种类         ：文件种类：1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-诊断文书;6-知情文件;7-诊疗申请;8-诊疗报告
'##         str页面编号     ：页面编号
'##         lngRecID        ：记录ID
'##         blnPrintDirectly：直接全文打印，不显示窗体
'##         mblnOrigMode    ：是否是初始状态，默认打印最终状态
'##         blnNoAsk        ：静默打印
'##         blnMoved        ：打印的是已转储的数据
'##         lngCopies       ：指定打印份数，＝0不指定，不控制，否则控制打印份数不能修改；并通过参数返回打印份数
'## 说明：
'################################################################################################################
Public Sub DoOnlyDocPreview(ByRef frmParent As Object, _
    ByVal eDocType As EPRDocTypeEnum, _
    Optional ByVal lng病人ID As Long, _
    Optional ByVal lng主页ID As Long, _
    Optional ByVal lng种类 As Long, _
    Optional ByVal str页面编号 As String, _
    Optional ByVal lngRecId As Long = -1, _
    Optional ByVal blnPrintDirectly As Boolean = False, _
    Optional ByVal mblnOrigMode As Boolean = False, _
    Optional ByVal blnNoAsk As Boolean = False, _
    Optional ByVal blnMoved As Boolean, _
    Optional ByVal lngAdviceID As Long, _
    Optional ByRef strPrinterDeviceName As String, _
    Optional ByVal lngCopies As Long)

    ZoomFactor = 1#
    If Not blnNoAsk Then
        zlCommFun.ShowFlash "请稍候..."
        Screen.MousePointer = vbHourglass
    End If
    '=================================================================================================
    Dim strF As String, j As Integer
    Dim Doc As New cEPRDocument
    
    strF = CreateTmpFile("tmp", App.hInstance & "_1") '保存临时文件
    Doc.InitEPRDoc cprEM_修改, cprET_单病历审核, lngRecId, IIf(eDocType = cpr门诊病历, cprPF_门诊, cprPF_住院), lng病人ID, CStr(lng主页ID), 0, glngDeptId, lngAdviceID, blnMoved
    Doc.OpenEPRDoc Doc.frmEditor.Editor1, blnMoved   '打开该文件，转换为最终清洁模式
    
    ReDim Files(1 To 1) As FileInfo
    Files(1).ID = lngRecId
    Files(1).PatiID = Doc.EPRPatiRecInfo.病人ID
    Files(1).PageID = Doc.EPRPatiRecInfo.主页ID
    Files(1).FileType = Doc.EPRPatiRecInfo.病历种类
    
    '如果是报告图片则进行缩放
    If eDocType = cpr诊疗报告 Then
        Dim t_StdPic As StdPicture
        For j = 1 To Doc.Tables.Count
            If Doc.Tables(j).TableType = tte_报告图片组 Then
                Set t_StdPic = Doc.Tables(j).GetFinalPic(False)
                Doc.Tables(j).ZoomPicture = True
                Doc.Tables(j).Refresh Doc.frmEditor.Editor1, t_StdPic
                Doc.Tables(j).ZoomPicture = False
            End If
        Next
        Call RemoveSign(Doc.frmEditor.Editor1, Doc)
        Call Doc.GetReplacedHeadFootString(Doc.frmEditor.Editor1, True)
    End If
    
    If Doc.frmEditor.SaveDocToFile(strF, Not mblnOrigMode, False) Then      '存储清洁临时文件，但是保留原来的关键字（保证分页效果不变！）
        edtThis.Freeze
        edtThis.ReadOnly = False
        edtThis.ForceEdit = True
        edtThis.InProcessing = True
        edtThis.Tag = "LoadFile"
        Doc.frmEditor.Editor1.Freeze
        Doc.frmEditor.Editor1.ReadOnly = False
        Doc.frmEditor.Editor1.ForceEdit = True
        Doc.frmEditor.Editor1.InProcessing = True
        '打开文件
        edtThis.OpenDoc strF
        '删除临时文件
        If gobjFSO.FileExists(strF) Then gobjFSO.DeleteFile strF, True
        
        For j = 1 To Doc.Elements.Count '删除空的要素，展开型要素重新刷新以去掉下波浪线
            If Doc.Elements(j).内容文本 = "" Or Doc.Elements(j).终止版 <> 0 Then
                Doc.Elements(j).DeleteFromEditor edtThis
            ElseIf Doc.Elements(j).输入形态 = 1 Then
                Doc.Elements(j).Refresh edtThis
            End If
        Next
    End If
    
    With edtThis
        '设置页眉页脚
        Set .Picture = Doc.frmEditor.Editor1.Picture
        .HeadFileTextRTF = Doc.frmEditor.Editor1.HeadFileTextRTF
        .FootFileTextRTF = Doc.frmEditor.Editor1.FootFileTextRTF
        
        Call Doc.GetReplacedHeadFootString(edtThis)
        .ForceEdit = False
        .PaperKind = Doc.frmEditor.Editor1.PaperKind
        .HeadFontFormat = Doc.frmEditor.Editor1.HeadFontFormat
        .FootFontFormat = Doc.frmEditor.Editor1.FootFontFormat
        .PaperWidth = Doc.frmEditor.Editor1.PaperWidth
        .PaperHeight = Doc.frmEditor.Editor1.PaperHeight
        .MarginLeft = Doc.frmEditor.Editor1.MarginLeft
        .MarginRight = Doc.frmEditor.Editor1.MarginRight
        .MarginTop = Doc.frmEditor.Editor1.MarginTop
        .MarginBottom = Doc.frmEditor.Editor1.MarginBottom
        .PaperOrient = Doc.frmEditor.Editor1.PaperOrient
        '分页
        .DoVirtualPrint
    End With
    
    Doc.frmEditor.Editor1.Modified = False
    Doc.frmEditor.Editor1.UnFreeze
    Doc.frmEditor.Editor1.RefreshTargetDC
    Doc.frmEditor.Editor1.ReadOnly = True
    Doc.frmEditor.Editor1.ForceEdit = False
    Doc.frmEditor.Editor1.InProcessing = False
    Unload Doc.frmEditor
    Set Doc = Nothing
    edtThis.UnFreeze
    edtThis.RefreshTargetDC
    edtThis.ReadOnly = True
    edtThis.ForceEdit = False
    edtThis.InProcessing = False
    edtThis.Tag = ""
    '=================================================================================================
    If Not blnNoAsk Then
        zlCommFun.StopFlash
        Screen.MousePointer = vbDefault
    End If
    mlngCopies = lngCopies
    If blnPrintDirectly Then
        If edtThis.PrintDoc(blnNoAsk, 1, 0, strPrinterDeviceName, mlngCopies) Then
            Call PrintTag(Files(1).ID, Files(1).FileType, Files(1).PatiID, Files(1).PageID) '记录打印记录
            RaiseEvent PrintEpr(lngRecId)
        End If
    Else
        pAttachMessages     '滚动条消息
        Call InitCommandBars    '工具栏初始化
        With edtThis
            '刷新缩略图
            mlngPageCount = .PageCount
            mlngCurPage = 1
            vfg页面.Rows = mlngPageCount
            vfg页面.ColWidth(0) = 0
            vfg页面.ColWidth(1) = 2100
            vfg页面.RowHeightMin = 2900
            vfg页面.FixedRows = 0
            vfg页面.FixedCols = 0
            
            '记录每页起始、终止位置
            Dim dblZoom As Double
            If .PaperWidth / 2000 > .PaperHeight / 3000 Then
                dblZoom = 2000 / .PaperWidth
                vfg页面.RowHeightMin = .PaperHeight * dblZoom + 20
            Else
                dblZoom = 3000 / .PaperHeight
            End If
            
            picBuff.Width = .PaperWidth
            picBuff.Height = .PaperHeight
            picZoom.Width = .PaperWidth * dblZoom
            picZoom.Height = .PaperHeight * dblZoom
            If Not cboStartPage Is Nothing Then
                cboStartPage.Clear
            End If
            For j = 1 To mlngPageCount
                picBuff.Cls
                cboStartPage.AddItem "第 " & CStr(j) & " 页"
                edtThis.PrintPage j, picBuff, True
                vfg页面.Cell(flexcpText, j - 1, 0) = j
                '缩放图片
                picZoom.Cls
                '采用半色调缩放效果最好！
                SetStretchBltMode picZoom.hDC, HALFTONE
                StretchBlt picZoom.hDC, 0, 0, .PaperWidth * dblZoom, .PaperHeight * dblZoom, picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, SRCCOPY
                
                picZoom.Line (0, 0)-(picZoom.ScaleWidth - 15, picZoom.ScaleHeight - 15), RGB(99, 99, 99), B
                picZoom.Line (15, 15)-(picZoom.ScaleWidth - 30, picZoom.ScaleHeight - 30), vbBlack, B
                imlPages.ListImages.Add j, "K" & j, picZoom.Image
                vfg页面.Cell(flexcpPicture, j - 1, 1) = imlPages.ListImages("K" & j).Picture
                vfg页面.Cell(flexcpPictureAlignment, j - 1, 1) = 3
            Next
            imlPages.ListImages.Clear
            vfg页面.ROW = 0
            vfg页面_RowColChange
            .ForceEdit = False
        End With
        Me.Show vbModal, frmParent
    End If
End Sub
'################################################################################################################
'## 功能：  显示多文件打印预览窗体，只用于住院医生工作站
'##
'## 参数：  frmParent       ：父窗体
'##         lng病人ID       ：病人ID
'##         lng主页ID       ：主页ID
'##         lng种类         ：文件种类：1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-诊断文书;6-知情文件;7-诊疗申请;8-诊疗报告
'##         str页面编号     ：页面编号
'##         lngRecID        ：记录ID
'##         blnPrintDirectly：直接全文打印，不显示窗体
'##         blnOrigMode    ：是否是初始状态，默认打印最终状态
'##         blnNoAsk        ：静默打印
'##         blnMoved        ：打印的是已转储的数据
'##         lngCopies       ：指定打印份数，＝0不指定，不控制，否则控制打印份数不能修改；并通过参数返回打印份数
'## 说明：  多文件按照顺序组合为清洁模式供打印，需要记录下每个文件的起始、终止位置；每页的起始、终止位置。
'################################################################################################################
Public Sub DoMultiDocPreview(ByRef frmParent As Object, _
    ByVal eDocType As EPRDocTypeEnum, _
    Optional ByVal lng病人ID As Long, _
    Optional ByVal lng主页ID As Long, _
    Optional ByVal lng种类 As Long, _
    Optional ByVal str页面编号 As String, _
    Optional ByVal lngRecId As Long = -1, _
    Optional ByVal blnPrintDirectly As Boolean = False, _
    Optional ByVal blnOrigMode As Boolean = False, _
    Optional ByVal blnNoAsk As Boolean = False, _
    Optional ByVal blnMoved As Boolean, _
    Optional ByVal lngAdviceID As Long, _
    Optional ByRef strPrinterDeviceName As String, _
    Optional ByVal lngCopies As Long)
    
    Dim rs As New ADODB.Recordset, strIDs As String, varPar() As String
    Dim i As Long, lngStart As Long, lngLen1 As Long, lngLen2 As Long
    Dim strWhere As String
    
    On Error GoTo errHand
    
    If str页面编号 = "" Then
        strWhere = ""
    Else
        strWhere = " And f.编号 = [3] "
    End If
    
    gstrSQL = "Select Count(C.Id) As 数目, c.Id, c.病历名称, c.文件id, c.创建时间" & vbNewLine & _
                "From 病历文件列表 F, 病历文件列表 B, 电子病历记录 C" & vbNewLine & _
                "Where f.种类 = b.种类 And f.页面 = b.页面 And b.Id = c.文件id And c.Id = [1]" & vbNewLine & _
                "Group By c.Id, c.病历名称, c.文件id, c.创建时间"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", lngRecId)
    If rs!数目 = 1 Or eDocType = cpr门诊病历 Then '独立页面直接打印 或外部调用时明明是住院类型，却传门诊病历类型
        Call DoOnlyDocPreview(frmParent, eDocType, lng病人ID, lng主页ID, lng种类, str页面编号, lngRecId, blnPrintDirectly, blnOrigMode, blnNoAsk, blnMoved, lngAdviceID, strPrinterDeviceName, lngCopies)
        Exit Sub
    End If
    
    ZoomFactor = 1#
    If Not blnPrintDirectly Then
        zlCommFun.ShowFlash "请稍候，正在加载病历内容！"
        Screen.MousePointer = vbHourglass
    End If

    edtThis.Freeze
    edtThis.ReadOnly = False
    edtThis.ForceEdit = True
    edtThis.InProcessing = True
    edtThis.Tag = "LoadFile"
    
    '设置页眉页脚
    Call SetHeadFoot(edtThis, rs!文件ID)
    If rs!数目 > 1 Then '是共享性质的病历
        '读取共享页面文件ID
        strIDs = GetFileRange(rs!文件ID, lngRecId, Format(rs!创建时间, "yyyy-MM-dd HH:mm:ss"), eDocType, lng病人ID, lng主页ID, blnMoved)
        '读取共享页面的文件ID
        gstrSQL = "Select /*+ rule*/ a.Id, a.文件id, a.病历名称, a.病历种类, a.病人id, a.主页id, a.最后版本, a.保存人, a.完成时间, a.保存时间" & vbNewLine & _
                    "From 电子病历记录 A," & LongIDsTable(strIDs, varPar) & vbNewLine & _
                    "Where a.Id = b.Id" & vbNewLine & _
                    "Order By a.序号, a.创建时间"
        If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9))
    
        ReDim Files(1 To rs.RecordCount) As FileInfo
        i = 0
        Do Until rs.EOF
            i = i + 1
            Files(i).ID = rs!ID
            Files(i).PatiID = rs!病人ID
            Files(i).PageID = rs!主页ID
            Files(i).FileType = rs!病历种类
            rs.MoveNext
        Loop
    End If
    
    For i = 1 To UBound(Files)
        edtBuff.Freeze
        edtBuff.ReadOnly = False
        edtBuff.ForceEdit = True
        edtBuff.InProcessing = True
        edtBuff.Tag = "LoadFile"
        If Not blnPrintDirectly Then
            zlCommFun.ShowFlash "请稍待，正在加载第" & i & "份病历内容！"
        End If
        '读取RTF文件
        Call ReadRTF(edtBuff, Files(i).ID, Not blnOrigMode, blnMoved)
        '追加RTF文件
        lngLen1 = Len(edtBuff.Text) '记录临时文件开始、结束位置
        lngLen2 = Len(edtThis.Text) '将文件添加到主文档末尾
        edtThis.Range(lngLen2, lngLen2).Selected
        edtBuff.SelectAll
        edtBuff.CopyWithFormat
        edtThis.PasteWithFormat
        lngStart = Len(edtThis.Text)
        If i < UBound(Files) Then
            '只要不是最后一份文件，末尾保证有一个回车，以备追加下一个文件
            If edtThis.Range(lngStart - 2, lngStart) = vbCrLf Then
                edtThis.Range(lngStart - 2, lngStart).Font.Hidden = False
            Else
                edtThis.Range(lngStart, lngStart).Text = vbCrLf
                edtThis.Range(lngStart, lngStart + 2).Font.Hidden = False
            End If
        End If
        edtThis.TOM.TextDocument.Range(lngStart, lngStart).Para = edtBuff.TOM.TextDocument.Range(lngLen1, lngLen1).Para '.Duplicate
    Next
    
    '替换页眉页脚关键字
    Call ReplacedHeadFootString(edtThis, lngRecId, blnMoved)
    
    '分页
    edtBuff.UnFreeze
    edtBuff.RefreshTargetDC
    edtBuff.ReadOnly = True
    edtBuff.ForceEdit = False
    edtBuff.InProcessing = False
    edtBuff.Tag = ""
    edtThis.UnFreeze
    edtThis.RefreshTargetDC
    edtThis.ReadOnly = True
    edtThis.ForceEdit = False
    edtThis.InProcessing = False
    edtThis.Tag = ""
    edtThis.DoVirtualPrint
    
    If Not blnPrintDirectly Then
        zlCommFun.StopFlash
        Screen.MousePointer = vbDefault
    End If
    mlngCopies = lngCopies
    If blnPrintDirectly Then '直接打印
        If edtThis.PrintDoc(blnNoAsk, 1, 0, strPrinterDeviceName, mlngCopies) Then
            For i = 1 To UBound(Files) '作打印标记
                Call PrintTag(Files(i).ID, Files(i).FileType, Files(i).PatiID, Files(i).PageID)
                RaiseEvent PrintEpr(Files(i).ID)
            Next
        End If
    Else                    '显示预览窗口
        pAttachMessages     '滚动条消息
        Call InitCommandBars    '工具栏初始化
        '刷新缩略图
        mlngPageCount = edtThis.PageCount
        mlngCurPage = 1
        vfg页面.Rows = mlngPageCount
        vfg页面.ColWidth(0) = 0
        vfg页面.ColWidth(1) = 2100
        vfg页面.RowHeightMin = 2900
        vfg页面.FixedRows = 0
        vfg页面.FixedCols = 0
        
        '记录每页起始、终止位置
        Dim dblZoom As Double
        If edtThis.PaperWidth / 2000 > edtThis.PaperHeight / 3000 Then
            dblZoom = 2000 / edtThis.PaperWidth
            vfg页面.RowHeightMin = edtThis.PaperHeight * dblZoom + 20
        Else
            dblZoom = 3000 / edtThis.PaperHeight
        End If
        
        picBuff.Width = edtThis.PaperWidth
        picBuff.Height = edtThis.PaperHeight
        picZoom.Width = edtThis.PaperWidth * dblZoom
        picZoom.Height = edtThis.PaperHeight * dblZoom
        cboStartPage.Clear
        For i = 1 To mlngPageCount
            picBuff.Cls
            cboStartPage.AddItem "第 " & CStr(i) & " 页"
            edtThis.PrintPage i, picBuff, True
            vfg页面.Cell(flexcpText, i - 1, 0) = i
            '缩放图片
            picZoom.Cls
            '采用半色调缩放效果最好！
            SetStretchBltMode picZoom.hDC, HALFTONE
            StretchBlt picZoom.hDC, 0, 0, edtThis.PaperWidth * dblZoom, edtThis.PaperHeight * dblZoom, picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, SRCCOPY
            
            picZoom.Line (0, 0)-(picZoom.ScaleWidth - 15, picZoom.ScaleHeight - 15), RGB(99, 99, 99), B
            picZoom.Line (15, 15)-(picZoom.ScaleWidth - 30, picZoom.ScaleHeight - 30), vbBlack, B
            imlPages.ListImages.Add i, "K" & i, picZoom.Image
            vfg页面.Cell(flexcpPicture, i - 1, 1) = imlPages.ListImages("K" & i).Picture
            vfg页面.Cell(flexcpPictureAlignment, i - 1, 1) = 3
        Next
        imlPages.ListImages.Clear
        vfg页面.ROW = 0
        vfg页面_RowColChange
        Me.Show vbModal, frmParent
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Sub DoSingleDocPreview(ByRef Editor1 As Editor, ByRef frmParent As Object, ByRef Doc As cEPRDocument, _
        Optional ByVal blnClearMode As Boolean = False, Optional ByVal blnPreview As Boolean = True, Optional ByVal blnNoAsk As Boolean)
    '******************************************************************************************************************
    '## 功能：  显示单文件打印预览窗体
    '##
    '## 参数：  Editor1         ：编辑器控件
    '##         frmParent       ：父窗体
    '******************************************************************************************************************
    Dim strFootSign As String, strPrintName As String
    Dim int替换域 As Integer
    Dim intLoop As Integer
    On Error GoTo errHand
    
    strPrintName = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & frmParent.Name, "PrintName", "")
    Call InitCommandBars    '工具栏初始化
    ZoomFactor = 1#
    mlngCopies = 0
    If Not blnNoAsk Then
        zlCommFun.ShowFlash "请稍候..."
        Screen.MousePointer = vbHourglass
    End If
    '=================================================================================================
    
    Dim lngLen As Long, strF As String
'    strF = App.Path & "\TMP.rtf"
    strF = CreateTmpFile("tmp", "zlRichFile")
    With Me.edtThis
        .Freeze
        .ReadOnly = False
        .ForceEdit = True
        .InProcessing = True
        .Tag = "LoadFile"
        
        lngLen = Len(Editor1.Text)
        'RTF内容赋值
        Editor1.SaveDoc strF
        .OpenDoc strF
        .HeadFileTextRTF = Editor1.HeadFileTextRTF
        .FootFileTextRTF = Editor1.FootFileTextRTF
        Call frmParent.Document.GetReplacedHeadFootString(edtThis)
        If blnClearMode Then
            .AuditMode = True
            .AcceptAuditText
        End If
        
        '如果是报告图片则进行缩放
        Dim j As Long, t_StdPic As StdPicture
        For j = 1 To Doc.Tables.Count
            If Doc.Tables(j).TableType = tte_报告图片组 Then
                Set t_StdPic = Doc.Tables(j).GetFinalPic(False)
                Doc.Tables(j).ZoomPicture = True
                Doc.Tables(j).Refresh Me.edtThis, t_StdPic
                Doc.Tables(j).ZoomPicture = False
            End If
        Next
        
        
        For intLoop = 1 To frmParent.Document.Elements.Count '要预览窗体控件上删除空要素，展开型要素重新刷新以去掉下波浪线
            If frmParent.Document.Elements(intLoop).内容文本 = "" Or frmParent.Document.Elements(intLoop).终止版 <> 0 Then
                frmParent.Document.Elements(intLoop).DeleteFromEditor edtThis
            ElseIf Doc.Elements(intLoop).输入形态 = 1 Then
                frmParent.Document.Elements(intLoop).Refresh edtThis
            End If
        Next
        
        If frmParent.Document.EPRPatiRecInfo.病历种类 = cpr诊疗报告 Then
            Call RemoveSign(Me.edtThis, frmParent.Document)
            Call frmParent.Document.GetReplacedHeadFootString(edtThis, True)
        End If
        
        If gobjFSO.FileExists(strF) Then gobjFSO.DeleteFile strF
        .ForceEdit = False
        Dim i As Long
        i = 0
        .ForceEdit = True
        '去掉所有颜色（清洁模式）
        lngLen = Len(.Text)
        For i = 0 To lngLen - 1
            If .Range(i, i + 1).Font.BackColor = ELE_BACKCOLOR Then
                .Range(i, i + 1).Font.BackColor = tomAutoColor
            End If
            If .Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR Then
                .Range(i, i + 1).Font.ForeColor = tomAutoColor
            End If
        Next
        .SelectAll
        .Selection.Font.BackColor = tomAutoColor
        .Range(0, 0).Selected
        Set .Picture = Editor1.Picture
        .PaperKind = Editor1.PaperKind
        .HeadFontFormat = Editor1.HeadFontFormat
        .FootFontFormat = Editor1.FootFontFormat
        .PaperWidth = Editor1.PaperWidth
        .PaperHeight = Editor1.PaperHeight
        .MarginLeft = Editor1.MarginLeft
        .MarginRight = Editor1.MarginRight
        .MarginTop = Editor1.MarginTop
        .MarginBottom = Editor1.MarginBottom
        .PaperOrient = Editor1.PaperOrient
        '分页
        .DoVirtualPrint
        
        '刷新缩略图
        mlngPageCount = .PageCount
        mlngCurPage = 1
        vfg页面.Rows = mlngPageCount
        vfg页面.ColWidth(0) = 0
        vfg页面.ColWidth(1) = 2100
        vfg页面.RowHeightMin = 2900
        vfg页面.FixedRows = 0
        vfg页面.FixedCols = 0
        
        '记录每页起始、终止位置
        Dim dblZoom As Double
        If .PaperWidth / 2000 > .PaperHeight / 3000 Then
            dblZoom = 2000 / .PaperWidth
            vfg页面.RowHeightMin = .PaperHeight * dblZoom + 20
        Else
            dblZoom = 3000 / .PaperHeight
        End If
        
        picBuff.Width = .PaperWidth
        picBuff.Height = .PaperHeight
        picZoom.Width = .PaperWidth * dblZoom
        picZoom.Height = .PaperHeight * dblZoom
        cboStartPage.Clear
        For i = 1 To mlngPageCount
            picBuff.Cls
            cboStartPage.AddItem "第 " & CStr(i) & " 页"
            .PaperWidth = Editor1.PaperWidth
            .PaperHeight = Editor1.PaperHeight
            .PrintPage i, picBuff, True
            vfg页面.Cell(flexcpText, i - 1, 0) = i
            '缩放图片
            picZoom.Cls
            '采用半色调缩放效果最好！
            SetStretchBltMode picZoom.hDC, HALFTONE
            StretchBlt picZoom.hDC, 0, 0, .PaperWidth * dblZoom, .PaperHeight * dblZoom, picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, SRCCOPY
'            picZoom.PaintPicture picBuff.Image, 0, 0, .PaperWidth * dblZoom, .PaperHeight * dblZoom
            
            picZoom.Line (0, 0)-(picZoom.ScaleWidth - 15, picZoom.ScaleHeight - 15), RGB(99, 99, 99), B
            picZoom.Line (15, 15)-(picZoom.ScaleWidth - 30, picZoom.ScaleHeight - 30), vbBlack, B
            imlPages.ListImages.Add i, "K" & i, picZoom.Image
            
            vfg页面.Cell(flexcpPicture, i - 1, 1) = imlPages.ListImages("K" & i).Picture
            vfg页面.Cell(flexcpPictureAlignment, i - 1, 1) = 3
        Next
        .UnFreeze
        .RefreshTargetDC
        .ReadOnly = True
        .ForceEdit = False
        .InProcessing = False
        .Tag = ""
        vfg页面.ROW = 0
        vfg页面_RowColChange
    End With
    '=================================================================================================
    ReDim Files(1 To 1) As FileInfo
    Files(1).ID = Doc.EPRPatiRecInfo.ID
    Files(1).PatiID = Doc.EPRPatiRecInfo.病人ID
    Files(1).PageID = Doc.EPRPatiRecInfo.主页ID
    Files(1).FileType = Doc.EPRPatiRecInfo.病历种类
    zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    If blnPreview Then
        pAttachMessages     '滚动条消息
        Me.Show vbModal, frmParent
    Else
        If edtThis.PrintDoc(blnNoAsk, 1, 0, strPrintName, mlngCopies) Then
            Call PrintTag(Files(1).ID, Files(1).FileType, Files(1).PatiID, Files(1).PageID) '记录打印记录
            RaiseEvent PrintEpr(Files(1).ID)
        End If
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & frmParent.Name, "PrintName", strPrintName
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub picBlank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngY = Y
    bMouseDown = True
End Sub

Private Sub picBlank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bMouseDown Then
        Unload mfrmTipInfo
        Dim lngTop As Long
        lngTop = IIf((picBlank.Top + (Y - lngY)) < 0, 0, picBlank.Top + (Y - lngY))
        lngTop = IIf(lngTop > picPage.ScaleHeight, picPage.ScaleHeight - picBlank.Height, lngTop)
        picBlank.Top = lngTop
        mlngBlankHeight = IIf(picBlank.Top > 100, picBlank.Top, 0)
        mlngBlankHeight = mlngBlankHeight / ZoomFactor
    Else
        mfrmTipInfo.ShowTipInfo picBlank.hwnd, "请把需要遮罩的最后一行和其下方的行间距空白一并遮罩。", True
    End If
End Sub

Private Sub picBlank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '刷新半透明矩形框
    Call DrawAlphaRect(mlngBlankHeight * ZoomFactor)
    bMouseDown = False
End Sub

Private Sub DrawAlphaRect(lngHeight As Long)
    '绘制半透明矩形框
    Dim lBlend As Long
    Dim bf As BLENDFUNCTION
    
    ' Draw the first picture:
    bf.BlendOp = AC_SRC_OVER
    bf.BlendFlags = 0
    bf.SourceConstantAlpha = 255
    bf.AlphaFormat = 0
    CopyMemory lBlend, bf, 4
    picBuff.Cls
    edtThis.PrintPage mlngCurPage, picBuff, True
    
    SetStretchBltMode picPage.hDC, HALFTONE
    StretchBlt picPage.hDC, 0, 0, edtThis.PaperWidth * ZoomFactor, edtThis.PaperHeight * ZoomFactor, picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, SRCCOPY

'    AlphaBlend picPage.hdc, 0, 0, _
'        picPage.ScaleWidth \ Screen.TwipsPerPixelX, _
'        picPage.ScaleHeight \ Screen.TwipsPerPixelY, _
'        picBuff.hdc, 0, 0, _
'        picBuff.ScaleWidth \ Screen.TwipsPerPixelX, _
'        picBuff.ScaleHeight \ Screen.TwipsPerPixelY, _
'        lBlend
    
    bf.SourceConstantAlpha = 65
    CopyMemory lBlend, bf, 4
    AlphaBlend picPage.hDC, 0, 0, _
        picPage.ScaleWidth \ Screen.TwipsPerPixelX, _
        lngHeight \ Screen.TwipsPerPixelY, _
        picBlank.hDC, 0, 0, _
        picBlank.ScaleWidth \ Screen.TwipsPerPixelX, _
        picBlank.ScaleHeight \ Screen.TwipsPerPixelY, _
        lBlend
    picPage.Refresh
End Sub

Private Sub Subclass1_WndProc(msg As Long, wParam As Long, lParam As Long, Result As Long)
    '自定义的消息处理函数
    Dim tP As POINTAPI
    Dim sngX As Single, sngY As Single   '鼠标坐标
    Dim intShift As Integer              '鼠标按键
    Dim bWay As Boolean                  '鼠标方向
    Dim bMouseFlag As Boolean            '鼠标事件激活标志

    Select Case msg
    Case WM_MOUSEWHEEL   '滚动
        Dim wzDelta, wKeys As Integer
        'wzDelta传递滚轮滚动的快慢，该值小于零表示滚轮向后滚动（朝用户方向），
        '大于零表示滚轮向前滚动（朝显示器方向）
        wzDelta = HIWORD(wParam)
        'wKeys指出是否有CTRL=8、SHIFT=4、鼠标键(左=2、中=16、右=2、附加)按下，允许复合
        wKeys = LOWORD(wParam)
        tP.X = LOWORD(lParam)    'pt鼠标的坐标
        tP.Y = HIWORD(lParam)
        '--------------------------------------------------
        If wzDelta < 0 Then  '朝用户方向
           bWay = True
        Else                 '朝显示器方向
           bWay = False
        End If
        '--------------------------------------------------
        '将屏幕坐标转换为Form1.窗口坐标
        ScreenToClient hwnd, tP
        sngX = tP.X
        sngY = tP.Y
        intShift = wKeys
        bMouseFlag = True  '置滚动标志
        If bMouseFlag = True Then
            bMouseFlag = False
            DoMouseWheel bWay, intShift, sngX, sngY, CLng(wzDelta)
        End If
    End Select
End Sub

Private Sub DoMouseWheel(bBackDirection As Boolean, Shift As Integer, X As Single, Y As Single, Value As Single)
    '鼠标滚动的处理
    If Shift = 8 Then
        '缩放处理
        Dim r As Double
        If bBackDirection Then
            '缩小
            r = IIf(ZoomFactor - 0.25 < 0.25, 0.25, ZoomFactor - 0.25)
        Else
            r = IIf(ZoomFactor + 0.25 > 1#, 1#, ZoomFactor + 0.25)
        End If
        ZoomFactor = r
        PreviewPage mlngCurPage
    Else
        Dim lngR As Long
        lngR = VS.Value - IIf(Value < 0, -1, 1) * 50
        lngR = IIf(lngR > VS.Max, VS.Max, lngR)
        lngR = IIf(lngR < VS.Min, VS.Min, lngR)
        VS.Value = lngR
    End If
End Sub

Private Sub picPage_Resize()
    picBlank.Left = 0: picBlank.Width = picPage.ScaleWidth
    picBlank.Top = mlngBlankHeight * ZoomFactor
    picShadow.Move picPage.Left + Shadow_W, picPage.Top + Shadow_W, picPage.Width, picPage.Height
End Sub

Private Sub vfg页面_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vfg页面.ToolTipText = "第" & vfg页面.MouseRow + 1 & "页/共" & vfg页面.Rows & "页"
End Sub

Private Sub vfg页面_RowColChange()
    vfg页面.ShowCell vfg页面.ROW, 1
    mlngCurPage = vfg页面.ROW + 1
    edtThis.CurPage = mlngCurPage
    PreviewPage mlngCurPage
End Sub

Private Sub PreviewPage(ByVal PageNum As Long)
    '预览第PageNum页的页面
    picBlank.Visible = (PageNum = mlngStartPage)
    LockWindowUpdate picPage.hwnd
    picBuff.Cls
    picBuff.Width = edtThis.PaperWidth
    picBuff.Height = edtThis.PaperHeight
    Me.edtThis.PrintPage PageNum, picBuff, True
    '缩放图片
    picPage.Width = edtThis.PaperWidth * ZoomFactor
    picPage.Height = edtThis.PaperHeight * ZoomFactor
    picPage.Cls
    '采用半色调缩放效果最好！
    SetStretchBltMode picPage.hDC, HALFTONE
    StretchBlt picPage.hDC, 0, 0, edtThis.PaperWidth * ZoomFactor, edtThis.PaperHeight * ZoomFactor, picBuff.hDC, 0, 0, picBuff.Width, picBuff.Height, SRCCOPY
    Call Reposition
    If PageNum = mlngStartPage And mlngBlankHeight > 100 Then Call DrawAlphaRect(mlngBlankHeight * ZoomFactor)
    LockWindowUpdate 0
    UpdateWindow picPage.hwnd
    stbThis.Panels(2).Text = " 第 " & mlngCurPage & " 页/ 共 " & edtThis.PageCount & " 页"
End Sub

'################################################################################################################
'## 功能：  另存为RTF文件
'################################################################################################################
Private Function SaveAsRTFFile() As Boolean
    On Error GoTo LL
    Dim strF As String
    dlgThis.Filename = ""
    dlgThis.Filter = "*.rtf|*.rtf|*.*|*.*"
    dlgThis.ShowSave
    strF = dlgThis.Filename
    If strF <> "" Then
        '保存到文件
        Me.edtThis.SaveDoc strF
        SaveAsRTFFile = True
        MsgBox "保存成功！文件名:" & vbCrLf & strF, vbOKOnly + vbInformation, gstrSysName
    End If
    Exit Function
LL:
    MsgBox "保存失败！", vbOKOnly + vbInformation, gstrSysName
    SaveAsRTFFile = False
End Function

'################################################################################################################
'## 功能：  另存为TXT文件
'################################################################################################################
Private Function SaveAsTxtFile() As Boolean
    On Error GoTo LL
    Dim strF As String
    dlgThis.Filename = ""
    dlgThis.Filter = "*.txt|*.txt|*.*|*.*"
    dlgThis.ShowSave
    strF = dlgThis.Filename
    If strF <> "" Then
        '保存到文件
        Const ForReading = 1, ForWriting = 2, ForAppending = 3
        Dim fs As FileSystemObject, f As TextStream
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.OpenTextFile(strF, ForWriting, TristateUseDefault)
        f.Write Me.edtBuff.Text
        f.Close
        SaveAsTxtFile = True
        MsgBox "保存成功！文件名:" & vbCrLf & strF, vbOKOnly + vbInformation, gstrSysName
    End If
    Exit Function
LL:
    MsgBox "保存失败！", vbOKOnly + vbInformation, gstrSysName
    SaveAsTxtFile = False
End Function

'################################################################################################################
'## 功能：  另存为图片文件
'################################################################################################################
Private Function SaveAsPicture() As Boolean
    On Error GoTo LL
    Dim strF As String
    dlgThis.Filename = ""
    dlgThis.Filter = "*.bmp|*.bmp|*.*|*.*"
    dlgThis.ShowSave
    strF = dlgThis.Filename
    If strF <> "" Then
        '保存到文件
        SavePicture picPage.Image, strF
        SaveAsPicture = True
        MsgBox "保存成功！文件名:" & vbCrLf & strF, vbOKOnly + vbInformation, gstrSysName
    End If
    Exit Function
LL:
    MsgBox "保存失败！", vbOKOnly + vbInformation, gstrSysName
    SaveAsPicture = False
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim i As Long, strPrintName As String
    Select Case Control.ID
    Case ID_File_SaveCopy
        '保存副本(A)...
        Call SaveAsRTFFile
    Case ID_File_SaveTxt
        '保存为文本(V)...
        Call SaveAsTxtFile
    Case ID_File_SaveAsPic
        Call SaveAsPicture
    Case ID_FILE_PRINT
        '打印(P)...
        strPrintName = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmPrintPreview", "PrintName", "")
        If edtThis.PrintDoc(False, mlngStartPage, mlngBlankHeight, strPrintName, mlngCopies) Then
            For i = 1 To UBound(Files) '作打印标记
                If Files(i).ID > 0 Then
                    Call PrintTag(Files(i).ID, Files(i).FileType, Files(i).PatiID, Files(i).PageID) '记录打印记录
                End If
                RaiseEvent PrintEpr(Files(i).ID)
            Next
        End If
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmPrintPreview", "PrintName", strPrintName
    Case ID_FILE_PRINTINWORD
        Call PrintInWord
        For i = 1 To UBound(Files) '作打印标记
            If Files(i).ID > 0 Then
                Call PrintTag(Files(i).ID, Files(i).FileType, Files(i).PatiID, Files(i).PageID) '记录打印记录
            End If
            RaiseEvent PrintEpr(Files(i).ID)
        Next
    Case ID_FILE_EXIT
        '退出(X)
        Unload Me
    Case ID_View_ToolBar
        '工具栏(T)
    Case ID_View_StatusBar
        '状态栏(S)
        stbThis.Visible = Not stbThis.Visible
        cbsThis.RecalcLayout
    Case ID_View_ZoomFactor
        '缩放比例(C)
        Dim r As Double
        r = Val(Control.Text) / 100#
        ZoomFactor = r
        PreviewPage mlngCurPage
    Case ID_View_ZoomIn
        '放大
        ZoomFactor = IIf(ZoomFactor + 0.25 > 1#, 1#, ZoomFactor + 0.25)
        PreviewPage mlngCurPage
    Case ID_View_ZoomOut
        '缩小
        ZoomFactor = IIf(ZoomFactor - 0.25 < 0.25, 0.25, ZoomFactor - 0.25)
        PreviewPage mlngCurPage
    Case ID_View_First
        '第一页
        vfg页面.ROW = 0
    Case ID_View_Prev
        '前一页
        vfg页面.ROW = IIf(vfg页面.ROW - 1 > 0, vfg页面.ROW - 1, 0)
    Case ID_View_Next
        '后一页
        vfg页面.ROW = IIf(vfg页面.ROW + 1 > vfg页面.Rows, vfg页面.Rows, vfg页面.ROW + 1)
    Case ID_View_Last
        '最后一页
        vfg页面.ROW = vfg页面.Rows - 1
    Case ID_View_ActualSize
        '实际大小 Ctrl+1
        ZoomFactor = 1#
        PreviewPage mlngCurPage
    Case ID_View_FitSize
        '适合页面 Ctrl+2
        If picBack.ScaleWidth / edtThis.PaperWidth < picBack.ScaleHeight / edtThis.PaperHeight Then
            ZoomFactor = (picBack.ScaleWidth - Shadow_W * 4) / edtThis.PaperWidth
        Else
            ZoomFactor = (picBack.ScaleHeight - Shadow_W * 4) / edtThis.PaperHeight
        End If
        PreviewPage mlngCurPage
    Case ID_View_FitWidth
        '适合宽度 Ctrl+3
        ZoomFactor = (picBack.ScaleWidth - Shadow_W * 4) / edtThis.PaperWidth
        PreviewPage mlngCurPage
    Case ID_View_FitHeight
        '适合高度 Ctrl+4
        ZoomFactor = (picBack.ScaleHeight - Shadow_W * 4) / edtThis.PaperHeight
        PreviewPage mlngCurPage
    Case ID_View_Size_250
        '250%
        ZoomFactor = 2.5
        PreviewPage mlngCurPage
    Case ID_View_Size_200
        '200%
        ZoomFactor = 2#
        PreviewPage mlngCurPage
    Case ID_View_Size_150
        '150%
        ZoomFactor = 1.5
        PreviewPage mlngCurPage
    Case ID_View_Size_100
        '100%
        ZoomFactor = 1#
        PreviewPage mlngCurPage
    Case ID_View_Size_75
        '75%
        ZoomFactor = 0.75
        PreviewPage mlngCurPage
    Case ID_View_Size_50
        '50%
        ZoomFactor = 0.5
        PreviewPage mlngCurPage
    Case ID_View_Size_25
        '25%
        ZoomFactor = 0.25
        PreviewPage mlngCurPage
    Case ID_HELP_CONTENT
        '帮助主题
        ShowHelp App.ProductName, Me.hwnd, "frmPrintPreview", Int((glngSys) / 100)
    Case ID_HELP_CONTACT
        '发送反馈
        Call zlMailTo(Me.hwnd)
    Case ID_HELP_ONLINE
        '在线医业
        Call zlHomePage(Me.hwnd)
    Case ID_HELP_ABOUT
        '关于...
        ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
    Case ID_View_StartPage
        mlngStartPage = Val(Mid(Control.Text, 3))
        If mlngStartPage = 0 Or mlngStartPage > edtThis.PageCount Then Exit Sub
        
        mlngCurPage = mlngStartPage
        edtThis.CurPage = mlngStartPage
        vfg页面.RowHeightMin = 0
        For i = 0 To mlngStartPage - 2
            vfg页面.ROWHEIGHT(i) = 0
        Next
        For i = mlngStartPage - 1 To edtThis.PageCount - 1
            vfg页面.ROWHEIGHT(i) = 2900
        Next
        vfg页面.ROW = mlngStartPage - 1
        picBlank.Top = 0
        mlngBlankHeight = 0
        picBlank.Visible = True
        vfg页面_RowColChange
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height / Screen.TwipsPerPixelY
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    Me.cbsThis.GetClientRect Left, Top, Right, Bottom
    tabThis.Move (Left + 1) * Screen.TwipsPerPixelX, (Top + 1) * Screen.TwipsPerPixelY, 2500, (Bottom - Top - 2) * Screen.TwipsPerPixelY
    picBack.Move tabThis.Left + tabThis.Width + Screen.TwipsPerPixelX, _
        (Top + 1) * Screen.TwipsPerPixelY, _
        (Right - Left - 2) * Screen.TwipsPerPixelX - 2500 - VS.Width, _
        (Bottom - Top - 2) * Screen.TwipsPerPixelY - HS.Height
    Reposition
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next
If Not Me.Visible Then Exit Sub
    Select Case Control.ID
    Case ID_File_SaveCopy
        '保存副本(A)...
    Case ID_File_SaveTxt
        '保存为文本(V)...
    Case ID_FILE_PRINT
        '打印(P)...
    Case ID_FILE_EXIT
        '退出(X)
    Case ID_View_ToolBar
        '工具栏(T)
    Case ID_View_StatusBar
        '状态栏(S)
        Control.Checked = stbThis.Visible
    Case ID_View_ZoomFactor
        '缩放比例(C)
        Control.Text = Format(ZoomFactor, "0%")
    Case ID_View_ZoomIn
        '放大
        Control.Enabled = (ZoomFactor < 1#) And (Abs(ZoomFactor - 1#) > 0.00001)
    Case ID_View_ZoomOut
        '缩小
        Control.Enabled = (ZoomFactor > 0.25) And (Abs(ZoomFactor - 0.25) > 0.00001)
    Case ID_View_First
        '第一页
        Control.Enabled = (mlngPageCount > 1) And (mlngCurPage > mlngStartPage)
    Case ID_View_Prev
        '前一页
        Control.Enabled = (mlngPageCount > 1) And (mlngCurPage > mlngStartPage)
    Case ID_View_Next
        '后一页
        Control.Enabled = (mlngPageCount > 1) And (mlngCurPage < mlngPageCount)
    Case ID_View_Last
        '最后一页
        Control.Enabled = (mlngPageCount > 1) And (mlngCurPage < mlngPageCount)
    Case ID_View_ActualSize
        '实际大小 Ctrl+1
        Control.Checked = (Abs(ZoomFactor - 1#) < 0.00001)
    Case ID_View_FitSize
        '适合页面 Ctrl+2
    Case ID_View_FitWidth
        '适合宽度 Ctrl+3
    Case ID_View_FitHeight
        '适合高度 Ctrl+4
    Case ID_View_Size_250
        '250%
    Case ID_View_Size_200
        '200%
    Case ID_View_Size_150
        '150%
    Case ID_View_Size_100
        '100%
        Control.Checked = (Abs(ZoomFactor - 1#) < 0.00001)
    Case ID_View_Size_75
        '75%
        Control.Checked = (Abs(ZoomFactor - 0.75) < 0.00001)
    Case ID_View_Size_50
        '50%
        Control.Checked = (Abs(ZoomFactor - 0.5) < 0.00001)
    Case ID_View_Size_25
        '25%
        Control.Checked = (Abs(ZoomFactor - 0.25) < 0.00001)
    Case ID_HELP_CONTENT
        '帮助主题
    Case ID_HELP_CONTACT
        '发送反馈
    Case ID_HELP_ONLINE
        '在线医业
    Case ID_HELP_ABOUT
        '关于...
    End Select
End Sub

Private Sub InitCommandBars()
    '## 菜单初始化
Dim cbpPopup As CommandBarPopup                     '临时对象
Dim cbpPopupSub As CommandBarPopup                  '临时对象
Dim objControl As CommandBarControl                 '工具栏控件
Dim Combo As CommandBarComboBox                     '工具栏下拉框控件
Dim BarPreview As CommandBar
Dim cbp文件 As CommandBarPopup          '文件菜单
Dim cbp视图 As CommandBarPopup          '视图菜单
Dim cbp帮助 As CommandBarPopup          '帮助菜单
    
    '窗体位置恢复
    Call RestoreWinState(Me, App.ProductName)
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    Set cbp文件 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "文件(&F)")
    With cbp文件.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_File_SaveCopy, "保存副本(&A)..."): objControl.IconId = 104
        Set objControl = .Add(xtpControlButton, ID_File_SaveTxt, "另存为文本(&T)...")
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "打印(&P)..."): objControl.IconId = 103
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINTINWORD, "在Word中打印(&W)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "退出(&X)"): objControl.IconId = 191
        objControl.BeginGroup = True
    End With
    
    Set cbp视图 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "视图(&V)")
    With cbp视图.CommandBar.Controls
        Set cbpPopup = .Add(xtpControlPopup, 0, "工具栏(&T)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, "工具栏列表"
        Set objControl = .Add(xtpControlButton, ID_View_StatusBar, "状态栏(&S)"): objControl.IconId = conMenu_View_StatusBar
        
        Set cbpPopup = .Add(xtpControlPopup, 0, "缩放比例(&C)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ActualSize, "实际大小(&A)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_View_Size_75, "75%"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_View_Size_50, "50%"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_View_Size_25, "25%"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_View_ZoomIn, "放大"): objControl.IconId = 502
        objControl.BeginGroup = True: objControl.IconId = 502
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_View_ZoomOut, "缩小"): objControl.IconId = 513
        
        Set objControl = .Add(xtpControlButton, ID_View_First, "第一页(&F)   "): objControl.BeginGroup = True: objControl.IconId = 7401
        Set objControl = .Add(xtpControlButton, ID_View_Prev, "前一页(&P)   "): objControl.IconId = 7402
        Set objControl = .Add(xtpControlButton, ID_View_Next, "后一页(&N)   "): objControl.IconId = 7403
        Set objControl = .Add(xtpControlButton, ID_View_Last, "最后一页(&L) "): objControl.IconId = 7404
    End With
    
    Set cbp帮助 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "帮助(&H)")
    With cbp帮助.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "帮助主题(&H)")
        objControl.BeginGroup = True
        Set cbpPopupSub = .Add(xtpControlPopup, 0, "&Web上的" & gstrProductName)
        objControl.BeginGroup = True
        Set objControl = cbpPopupSub.CommandBar.Controls.Add(xtpControlButton, ID_HELP_ONLINE, gstrProductName & "在线(&H)"): objControl.IconId = conMenu_Help_Web_Forum
        Set objControl = cbpPopupSub.CommandBar.Controls.Add(xtpControlButton, ID_HELP_CONTACT, "发送反馈(&M)"): objControl.IconId = conMenu_Help_Web_Mail
        Set objControl = .Add(xtpControlButton, ID_HELP_ABOUT, "关于(&A)..."): objControl.IconId = conMenu_Help_About
        objControl.BeginGroup = True
    End With
    
    Set BarPreview = cbsThis.Add("打印预览", xtpBarTop)
    With BarPreview.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "打印"): objControl.IconId = 103
        objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINTINWORD, "在Word中打印")
        objControl.STYLE = xtpButtonIconAndCaption
    
        Set objControl = .Add(xtpControlButton, ID_View_ActualSize, "实际大小")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_View_ZoomIn, "放大"): objControl.IconId = 502
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_View_ZoomOut, "缩小"): objControl.IconId = 513
        Set Combo = .Add(xtpControlComboBox, ID_View_ZoomFactor, "缩放比例")
        Combo.AddItem "100%", 1
        Combo.AddItem "75%", 2
        Combo.AddItem "50%", 3
        Combo.AddItem "25%", 4
        Combo.ListIndex = 1
        Combo.Width = 80
        Combo.DropDownWidth = 80
        Combo.DropDownListStyle = True
        
        Set objControl = .Add(xtpControlButton, ID_View_First, "第一页"): objControl.IconId = 7401
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_View_Prev, "前一页"): objControl.IconId = 7402
        Set objControl = .Add(xtpControlButton, ID_View_Next, "后一页"): objControl.IconId = 7403
        Set objControl = .Add(xtpControlButton, ID_View_Last, "最后一页"): objControl.IconId = 7404
        
        Set objControl = .Add(xtpControlLabel, 0, "起始页面:")
        objControl.BeginGroup = True
        Set cboStartPage = .Add(xtpControlComboBox, ID_View_StartPage, "起始页面")
        cboStartPage.AddItem "第 1 页", 1
        cboStartPage.ListIndex = 1
        cboStartPage.Width = 80
        cboStartPage.DropDownWidth = 80
        cboStartPage.DropDownListStyle = True
        
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "关闭(&Q)"): objControl.IconId = 191
        objControl.BeginGroup = True
    End With
    
    
    '热键绑定
    cbsThis.KeyBindings.Add FCONTROL, Asc("S"), ID_File_SaveCopy
    cbsThis.KeyBindings.Add FCONTROL, Asc("P"), ID_FILE_PRINT
    cbsThis.KeyBindings.Add FCONTROL, Asc("Q"), ID_FILE_EXIT
    cbsThis.KeyBindings.Add FCONTROL, Asc("1"), ID_View_ActualSize
    
    cbsThis.KeyBindings.Add 0, VK_F1, ID_HELP_CONTENT
    cbsThis.KeyBindings.Add 0, vbKeyHome, ID_View_First
    cbsThis.KeyBindings.Add 0, vbKeyEnd, ID_View_Last
    cbsThis.KeyBindings.Add 0, vbKeyPageUp, ID_View_Prev
    cbsThis.KeyBindings.Add 0, vbKeyPageDown, ID_View_Next
    cbsThis.KeyBindings.Add 0, vbKeyAdd, ID_View_ZoomIn
    cbsThis.KeyBindings.Add 0, vbKeySubtract, ID_View_ZoomOut
    
    'TAB控件的初始化
    tabThis.Icons = zlCommFun.GetPubIcons
    tabThis.InsertItem 0, "页面缩略图 ", pic页面.hwnd, 513
    
    With tabThis.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .ShowIcons = True
        .DisableLunaColors = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    imlPages.ListImages.Clear
    ImageList_Destroy imlPages.hImageList
    Set picBack.Picture = Nothing
    Set picBlank.Picture = Nothing
    Set picBuff.Picture = Nothing
    Set picPage.Picture = Nothing
    Set picShadow.Picture = Nothing
    Set picZoom.Picture = Nothing
    Set pic页面.Picture = Nothing
    Set cboStartPage = Nothing
    Erase Files
    
    '保存位置信息
    Call SaveWinState(Me, App.ProductName)
    pDetachMessages
    '手动释放内存
'    SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
    EmptyWorkingSet GetCurrentProcess()
    Unload mfrmTipInfo
    Set mfrmTipInfo = Nothing
End Sub

Private Sub pic页面_Resize()
    vfg页面.Move 0, 0, pic页面.ScaleWidth, pic页面.ScaleHeight
End Sub

Private Sub picback_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngX = X: lngY = Y
    If Button = 2 Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        Set Popup = cbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            .Add xtpControlButton, ID_View_Size_100, "100%"
            .Add xtpControlButton, ID_View_Size_75, "75%"
            .Add xtpControlButton, ID_View_Size_50, "50%"
            .Add xtpControlButton, ID_View_Size_25, "25%"
            Set Control = .Add(xtpControlButton, ID_View_ZoomIn, "放大")
            Control.BeginGroup = True
            .Add xtpControlButton, ID_View_ZoomOut, "缩小"
            
            Set Control = .Add(xtpControlButton, ID_File_SaveAsPic, "另存为图片(&I)...")
            Control.BeginGroup = True
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub picback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If VS.Enabled Then
            If (Y - lngY) / 15 > 0 Then
                VS.Value = IIf(VS.Value - (Y - lngY) / 15 < VS.Min, VS.Min, VS.Value - (Y - lngY) / 15)
            Else
                VS.Value = IIf(VS.Value - (Y - lngY) / 15 > VS.Max, VS.Max, VS.Value - (Y - lngY) / 15)
            End If
        End If
        If HS.Enabled Then
            If (X - lngX) / 15 > 0 Then
                HS.Value = IIf(HS.Value - (X - lngX) / 15 < HS.Min, HS.Min, HS.Value - (X - lngX) / 15)
            Else
                HS.Value = IIf(HS.Value - (X - lngX) / 15 > HS.Max, HS.Max, HS.Value - (X - lngX) / 15)
            End If
        End If
    End If
End Sub

Private Sub picPage_DblClick()
    Dim r As Double
    r = ZoomFactor + 0.25
    If r > 1# Then r = 0.25
    ZoomFactor = r
    PreviewPage mlngCurPage
End Sub

Private Sub picPage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngX = X: lngY = Y
    If Button = 1 Then Set picPage.MouseIcon = HS.MouseIcon
    If Button = 2 Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        Set Popup = cbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            .Add xtpControlButton, ID_View_Size_100, "100%"
            .Add xtpControlButton, ID_View_Size_75, "75%"
            .Add xtpControlButton, ID_View_Size_50, "50%"
            .Add xtpControlButton, ID_View_Size_25, "25%"
            Set Control = .Add(xtpControlButton, ID_View_ZoomIn, "放大")
            Control.BeginGroup = True
            .Add xtpControlButton, ID_View_ZoomOut, "缩小"
            
            Set Control = .Add(xtpControlButton, ID_File_SaveAsPic, "另存为图片(&I)...")
            Control.BeginGroup = True
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub picPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If VS.Enabled Then
            If (Y - lngY) / 15 > 0 Then
                VS.Value = IIf(VS.Value - (Y - lngY) / 15 < VS.Min, VS.Min, VS.Value - (Y - lngY) / 15)
            Else
                VS.Value = IIf(VS.Value - (Y - lngY) / 15 > VS.Max, VS.Max, VS.Value - (Y - lngY) / 15)
            End If
        End If
        If HS.Enabled Then
            If (X - lngX) / 15 > 0 Then
                HS.Value = IIf(HS.Value - (X - lngX) / 15 < HS.Min, HS.Min, HS.Value - (X - lngX) / 15)
            Else
                HS.Value = IIf(HS.Value - (X - lngX) / 15 > HS.Max, HS.Max, HS.Value - (X - lngX) / 15)
            End If
        End If
    End If
End Sub

Private Sub Reposition()
    VS.Top = picBack.Top
    VS.Left = ScaleWidth - VS.Width
    VS.Height = picBack.Height
    
    HS.Left = picBack.Left
    HS.Top = picBack.Top + picBack.Height
    HS.Width = picBack.Width
    
    '调整预览页
    
    If picBack.ScaleWidth >= picPage.Width + Shadow_W * 4 Then
        picPage.Left = (picBack.ScaleWidth - (picPage.Width + Shadow_W * 4)) / 2 + Shadow_W * 2
        picShadow.Left = picPage.Left + Shadow_W
        HS.Enabled = False
    Else
        HS.Max = (picPage.Width + Shadow_W * 4 - picBack.ScaleWidth) / 15
        If HS.Max / 3 < HS.SmallChange Then
            HS.LargeChange = HS.SmallChange
        Else
            HS.LargeChange = HS.Max / 3
        End If
        HS.Value = 0
        HS.Enabled = True
        HS_Change
    End If
    If picBack.ScaleHeight >= picPage.Height + Shadow_W * 4 Then
        picPage.Top = (picBack.ScaleHeight - (picPage.Height + Shadow_W * 4)) / 2 + Shadow_W
        picShadow.Top = picPage.Top + Shadow_W
        VS.Enabled = False
    Else
        VS.Max = (picPage.Height + Shadow_W * 4 - picBack.ScaleHeight) / 15
        If VS.Max / 3 < VS.SmallChange Then
            VS.LargeChange = VS.SmallChange
        Else
            VS.LargeChange = VS.Max / 3
        End If
        VS.Value = 0
        VS.Enabled = True
        VS_Change
    End If
End Sub

Private Sub picPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Set picPage.MouseIcon = VS.MouseIcon
End Sub

Private Sub VS_Change()
    picPage.Top = -VS.Value * 15# + Shadow_W * 2
    picPage_Resize
    Me.Refresh
End Sub

Private Sub VS_Scroll()
    picPage.Top = -VS.Value * 15# + Shadow_W * 2
    picPage_Resize
    Me.Refresh
End Sub

Private Sub HS_Change()
    picPage.Left = -HS.Value * 15# + Shadow_W * 2
    picPage_Resize
    Me.Refresh
End Sub

Private Sub HS_Scroll()
    picPage.Left = -HS.Value * 15# + Shadow_W * 2
    picPage_Resize
    Me.Refresh
End Sub

Private Sub pAttachMessages()
'消息捕获绑定
On Error Resume Next
    Subclass1.hwnd = Me.hwnd
    Subclass1.Messages(WM_MOUSEWHEEL) = True
    m_bSubClassing = True
End Sub

Private Sub pDetachMessages()
'取消消息捕获
    On Error Resume Next
    
    If (m_bSubClassing) Then
        Subclass1.Messages(WM_MOUSEWHEEL) = False
        m_bSubClassing = False
    End If
End Sub

Private Sub PrintTag(ByVal lngID As Long, ByVal lngFileType As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long)
On Error GoTo errHand
    If lngPatiID = 0 Then Exit Sub '范文、定义
    gstrSQL = "Zl_电子病历打印_Insert(" & lngID & "," & lngFileType & "," & lngPatiID & "," & lngPageId & ",'" & gstrUserName & "')"
    zlDatabase.ExecuteProcedure gstrSQL, "frmPrintPreview"
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

