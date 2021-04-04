VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7D52C334-5021-43A4-8EB4-86CC21862ABF}#1.2#0"; "zlTable.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "zlRichEPR"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   7170
   ScaleWidth      =   9930
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtFeedBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   300
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   500
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtContent 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   4560
      MaxLength       =   500
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   5000
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2310
      Left            =   6450
      ScaleHeight     =   2310
      ScaleWidth      =   3690
      TabIndex        =   20
      Top             =   1110
      Width           =   3690
      Begin VB.PictureBox picHistoryInfo 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   105
         ScaleHeight     =   360
         ScaleWidth      =   1500
         TabIndex        =   22
         Top             =   60
         Width           =   1500
         Begin VB.Label lblHistoryInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "前 5 天的历史内容"
            Height          =   180
            Left            =   60
            TabIndex        =   23
            Top             =   75
            Width           =   1530
         End
      End
      Begin zlRichEditor.Editor edtThis 
         Height          =   2610
         Left            =   270
         TabIndex        =   21
         Top             =   645
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   4604
         ShowRuler       =   0   'False
      End
   End
   Begin zlRichEPR.ucPacsImgCanvas ucPacsImgCanvas1 
      Height          =   915
      Left            =   2160
      TabIndex        =   19
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
      _extentx        =   1296
      _extenty        =   1614
   End
   Begin zlRichEPR.ucPictureEditor ucPictureEditor1 
      Height          =   975
      Left            =   1080
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
      _extentx        =   1508
      _extenty        =   1720
   End
   Begin VB.Timer tmrAutoSaveEPR 
      Interval        =   1000
      Left            =   9000
      Top             =   765
   End
   Begin VB.PictureBox picPatiInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   45
      MouseIcon       =   "frmMain.frx":058A
      ScaleHeight     =   375
      ScaleWidth      =   9735
      TabIndex        =   7
      Top             =   3735
      Width           =   9735
      Begin VB.Label lblPatiInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "基本信息"
         Height          =   180
         Left            =   90
         TabIndex        =   12
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblPatiIns 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         ForeColor       =   &H00008000&
         Height          =   180
         Index           =   1
         Left            =   7410
         TabIndex        =   11
         Top             =   90
         Width           =   90
      End
      Begin VB.Label lblPatiIns 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号:"
         Height          =   180
         Index           =   0
         Left            =   6750
         TabIndex        =   10
         Top             =   90
         Width           =   630
      End
      Begin VB.Label lblPatiState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   9480
         TabIndex        =   9
         Top             =   90
         Width           =   90
      End
      Begin VB.Label lblPatiState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病况:"
         Height          =   180
         Index           =   0
         Left            =   8985
         TabIndex        =   8
         Top             =   90
         Width           =   450
      End
   End
   Begin VB.PictureBox picPenInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   5315
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   4185
      Visible         =   0   'False
      Width           =   1030
      Begin VB.TextBox txtPenInput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   20
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   20
         Width           =   1005
      End
   End
   Begin VB.Timer tmrThis 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8370
      Top             =   720
   End
   Begin VB.PictureBox picTMP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   7815
      ScaleHeight     =   420
      ScaleWidth      =   510
      TabIndex        =   3
      Top             =   5700
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picDropDown 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   7875
      ScaleHeight     =   315
      ScaleWidth      =   270
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5175
      Visible         =   0   'False
      Width           =   330
   End
   Begin zlRichEditor.Editor edtBuff 
      Height          =   600
      Left            =   360
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
   End
   Begin zlRichEPR.F1ColorPicker ColorFillColor 
      Height          =   2190
      Left            =   5940
      TabIndex        =   2
      Top             =   1170
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      AutoColor       =   16777215
   End
   Begin zlTable.Table tblThis 
      Height          =   1230
      Left            =   5400
      TabIndex        =   4
      Top             =   4905
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   2170
      SingleLine      =   0   'False
   End
   Begin zlRichEPR.ColorPicker ColorPaperBackColor 
      Height          =   2190
      Left            =   5760
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      AutoColor       =   16777215
   End
   Begin zlRichEPR.ColorPicker ColorHighlight 
      Height          =   2190
      Left            =   5580
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   810
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      AutoColor       =   16777215
   End
   Begin zlRichEPR.ColorPicker ColorForeColor 
      Height          =   2190
      Left            =   5400
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      Color           =   0
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   2790
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   8370
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C74
            Key             =   "HIGHLIGHT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DE9
            Key             =   "FORECOLOR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F36
            Key             =   "FILLCOLOR"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   17
      Top             =   6792
      Width           =   9936
      _ExtentX        =   17515
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4339
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2716
            MinWidth        =   2716
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1658
            MinWidth        =   1658
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "Ins"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin zlRichEditor.Editor Editor1 
      Height          =   2715
      Left            =   255
      TabIndex        =   13
      Top             =   810
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   4789
   End
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   5160
      MousePointer    =   7  'Size N S
      Top             =   2925
      Width           =   5115
   End
   Begin XtremeCommandBars.CommandBars cbrThis 
      Left            =   1215
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane DkpThis 
      Bindings        =   "frmMain.frx":10AA
      Left            =   300
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'## 全局变量
'######################################################################################################################
Public Document As cEPRDocument                         '文档类对象
Public glngCurEleKey As Long                            '当前元素ID

Public WithEvents mfrmCompends As frmCompends           '文档结构图
Attribute mfrmCompends.VB_VarHelpID = -1

Private WithEvents mfrmSentenceDetailed As frmSentenceDetailed       '示范词句窗体
Attribute mfrmSentenceDetailed.VB_VarHelpID = -1
Private WithEvents mfrmSegments As frmSegmentList        '示范片段窗体
Attribute mfrmSegments.VB_VarHelpID = -1
Private WithEvents mfrmModElement As frmElementEdit     '数据编辑窗体
Attribute mfrmModElement.VB_VarHelpID = -1
Private WithEvents mfrmInsElement As frmInsElement      '插入诊治要素窗体
Attribute mfrmInsElement.VB_VarHelpID = -1
Private WithEvents mfrmDicSelect As frmDicSelect        '插入字典项目
Attribute mfrmDicSelect.VB_VarHelpID = -1
Private WithEvents mfrmStyleMan As frmStyleMan          '段落样式维护
Attribute mfrmStyleMan.VB_VarHelpID = -1
Private WithEvents cPicEditor As cPictureEditor         '位图编辑器对象
Attribute cPicEditor.VB_VarHelpID = -1
Private WithEvents mfrmMultiDocView As frmMultiDocView  '多文档组合查阅
Attribute mfrmMultiDocView.VB_VarHelpID = -1
Private WithEvents mfrmPacsPic As frmPACSImg            'PACS图片列表窗体
Attribute mfrmPacsPic.VB_VarHelpID = -1
Private WithEvents mfrmMainError As frmMainMsg
Attribute mfrmMainError.VB_VarHelpID = -1
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mfrmPreview As frmPrintPreview
Attribute mfrmPreview.VB_VarHelpID = -1
Private WithEvents mfrmDocksymbol As frmDockSymbol      '特殊符号
Attribute mfrmDocksymbol.VB_VarHelpID = -1
Private WithEvents mfrmHistoryReport As frmDockReportHistory       '历史报告窗口
Attribute mfrmHistoryReport.VB_VarHelpID = -1

Private cDropDown As cDropDownToolWindow                '表格选择器

Private mlngHP As Long, blnSpaceEvent As Boolean        '记录自动增加空格的位置！
Private lngYOld As Long, lngXOld As Long, blnIsDown As Boolean
Private lngIndex As Long                                '文档Key（新建为“未命名文档1”、“未命名文档2”。。。等，依此类推）
Private lngPicPosition As Long                          '图片编辑位置
Private lngTablePosition As Long                        '表格编辑位置
Private mblnExistHistroy As Boolean                     '是否有历史内容列表
Private mlngSelHightlightColor As OLE_COLOR             '当前选中的文字高亮色
Private mlngSelForeColor As OLE_COLOR                   '当前选中的文字前景色
Private mlngCellFillColor As OLE_COLOR                  '单元格填充色

Private mParaFmt As cParaFormat                         '存储格式刷段落属性
Private mFontFmt As cFontFormat                         '存储格式刷字体属性
Private mblnFmtBrushDown As Boolean                     '格式刷按下与否

Private mblnIsMultiMode As Boolean                      '是否是多文档编辑模式
Private mintStyle As Integer                              '嵌入编辑 -1 非模态 vbModeless=0 模态 vbModal=1

Private Type PatiInfor
    姓名    As String
    身份证号 As String
End Type
Private mPatiInfor As PatiInfor
Private mblnPatiSign As Boolean
Private mblnEnPtSign As Boolean

Private Type UndoInfo
    Filename As String
    SelStart As Long
    SelEnd As Long
End Type

Private Type EleLimit
    变动原因 As Byte
    原因要件id As Long
    原因要素 As String
    原因内容 As String
    变动结果 As Byte
    结果提纲id As Long
    结果要件id As Long
    结果要素 As String
    结果值域 As String
    原始值域 As String
End Type
Private mEleLimit() As EleLimit
Private mlDiseaseID As Long
Private mlDiagnoseID As Long

Private UndoList() As UndoInfo                          '历史文件列表，1开始编号
Private p_Undo As Long                                  '当前的Undo指针
Private mblnAutosave As Boolean                         '是否开启自动缓存
Private mlngUndoLimit As Long                           'Undo最大步数，默认20步
Private mlngSaveInterval As Long                        '自动缓存时间间隔，秒
Private mblnAutoSaveEPR As Boolean                      '是否开启自动保存
Private mlngSaveIntervalEPR As Long                     '自动保存时间间隔，分钟
Private mblnAutoPageCount As Boolean                    '自动分页计数
Private mblnAutoPageNote As Boolean                     '自动增页提醒
Private mintSharePages As Integer                       '显示共享页面文件内容的数量
Private mblnNoAsk As Boolean                            '静默打印
Private mblnSignAutoAlter As Boolean                    '诊疗单据,签名自动移位
Private DT1_EPR As Date, DT2_EPR As Date
Private DT1 As Date, DT2 As Date, mblnChange As Boolean
Private mbEditInTable As Boolean                        '是否当前在表格编辑状态
Private mbln签名要素 As Boolean                         '是否有可用的签名要素
Private mblnChildMode As Boolean                        '是否是嵌入编辑的子窗体
Private mblnCanPrint As Boolean                         '是否可以打印预览
Private mblnReadOnly As Boolean                         '是否只读
Private mstrSex As String                               '病人的性别
Private mblnPrecess As Boolean                          '是否正在处理事务,处理中禁止关闭窗口
Private mbln返修处理 As Boolean                         '是否是返修处理传染病报告卡
Private mblnFBContentChanged As Boolean                 '返修处理说明是否修改了

'######################################################################################################################
Public Property Get ReadOnly() As Boolean
    ReadOnly = mblnReadOnly
End Property

Public Property Let ReadOnly(vData As Boolean)
    mblnReadOnly = vData
    Editor1.ReadOnly = vData
    If vData Then
        DkpThis.FindPane(ID_VIEW_PHRASEDEMO).Close
        DkpThis.FindPane(ID_VIEW_SEGMENT).Close
    Else
        DkpThis.ShowPane ID_VIEW_PHRASEDEMO
        DkpThis.ShowPane ID_VIEW_SEGMENT
    End If
End Property

Public Property Get CanPrint() As Boolean
    CanPrint = mblnCanPrint
End Property

Public Property Let CanPrint(vData As Boolean)
    mblnCanPrint = vData
End Property

Public Property Get ChildMode() As Boolean
    ChildMode = mblnChildMode
End Property

Public Property Let ChildMode(vData As Boolean)
    mblnChildMode = vData
    If mblnChildMode Then
        Me.BorderStyle = 0
        SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) Xor WS_BORDER Xor WS_THICKFRAME Xor WS_DLGFRAME
    Else
        Me.BorderStyle = 2
    End If
End Property

Private Function AutoMoveSignPos() As Boolean

    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long
    
    With Editor1
        lS = .Selection.StartPos
        If .SelLength > 0 Then Exit Function
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then
            Editor1.SelStart = lKEE
        End If
'        If sKeyType = "S" Then
'            If Editor1.Selection.StartPos >= lKSS And Editor1.Selection.StartPos <= lKSE And lKSS > 0 And lKSE > 0 Then
'                If lKSS > 0 Then Editor1.SelStart = lKSS
'            End If
'            If Editor1.Selection.StartPos >= lKES And Editor1.Selection.StartPos <= lKEE And lKES > 0 And lKEE > 0 Then
'                Editor1.SelStart = lKEE
'            End If
'        End If
    End With
                    
    AutoMoveSignPos = True
    
End Function
Private Function ShowSharePageHistory(ByVal Document As cEPRDocument, Optional ByVal intNumber As Integer = 5) As Boolean
    '******************************************************************************************************************
    '功能： 显示共享页面文件历史内容
    '参数： Document                    当前文档
    '       intNumber                   要显示历史的次数
    '返回： 有则返回真，否则假
    '******************************************************************************************************************
    Dim lngTMP As Long, rsTemp As New ADODB.Recordset, strTime As String, strIDs As String, varPar() As String
    Dim strFile As String, strZipFile As String, lngLen2 As Long, lngLen1 As Long, lngStart As Long
    Dim objEPRFileInfo As New cEPRFileDefineInfo
    Dim strTmpClipboard As String
    On Error GoTo errHand
    
    strTmpClipboard = Clipboard.GetText '临时记录粘贴板内容，本函数内部会用到粘贴板更改了粘贴内容
    If mintSharePages = 0 Then Exit Function
    
    edtThis.ReadOnly = False
    edtBuff.ReadOnly = False
    edtThis.Freeze
    edtThis.NewDoc
    edtBuff.NewDoc
    
    strTime = Format(Document.EPRPatiRecInfo.创建时间, "yyyy-MM-dd HH:mm:ss")
    lblHistoryInfo.Caption = "前 " & intNumber & " 次的历史内容"
    '查找当前病历前面的intNumber份病历
    strIDs = GetFileRange(Document.EPRFileInfo.ID, Document.EPRPatiRecInfo.ID, strTime, Document.EPRPatiRecInfo.病历种类, Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, False, Document.EPRPatiRecInfo.医嘱id)
    
    gstrSQL = "Select /*+ rule*/ a.Id, a.文件id, a.病历名称, a.病历种类, a.病人id, a.主页id,a.创建时间, a.最后版本, a.保存人, a.完成时间, a.保存时间" & vbNewLine & _
                "From 电子病历记录 A," & LongIDsTable(strIDs, varPar, 2) & vbNewLine & _
                "Where a.Id = b.Id" & vbNewLine & _
                "Order By a.序号, a.创建时间 Desc"
    gstrSQL = "Select ID,创建时间 From (" & gstrSQL & ") Where RowNum<=[1] Order by 创建时间" '限次数后依时间反向排序
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", intNumber, varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9))
    If rsTemp.BOF = False Then
        edtThis.ForceEdit = True
        edtBuff.ForceEdit = True

        Do While Not rsTemp.EOF
            strZipFile = zlBlobRead(5, Val(rsTemp!ID))

            If gobjFSO.FileExists(strZipFile) Then
                strFile = zlFileUnzip(strZipFile)
                If gobjFSO.FileExists(strFile) Then

                    edtBuff.OpenDoc strFile
                    
                    lngTMP = Val(rsTemp!ID)

                    lngLen1 = Len(edtBuff.Text)
                    lngLen2 = Len(edtThis.Text)

                    edtThis.Range(lngLen2, lngLen2).Selected
                    edtBuff.SelectAll

                    edtBuff.CopyWithFormat
                    edtThis.PasteWithFormat

                    lngStart = Len(edtThis.Text)

                    If rsTemp.AbsolutePosition < rsTemp.RecordCount Then
                        '末尾保证有一个回车
                        If edtThis.Range(lngStart - 2, lngStart) = vbCrLf Then
                            edtThis.Range(lngStart - 2, lngStart).Font.Hidden = False
                        Else
                            edtThis.Range(lngStart, lngStart).Text = vbCrLf
                            edtThis.Range(lngStart, lngStart + 2).Font.Hidden = False
                        End If
                    End If
                    edtThis.TOM.TextDocument.Range(lngStart, lngStart).Para = edtBuff.TOM.TextDocument.Range(lngLen1, lngLen1).Para
                End If
                
                If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile
                If gobjFSO.FileExists(strZipFile) Then gobjFSO.DeleteFile strZipFile
            End If

            rsTemp.MoveNext
        Loop

        gstrSQL = "Select c.ID, a.格式 From   病历页面格式 a, 病历文件列表 b, 电子病历记录 c " & _
                " Where  c.文件id = b.id And a.种类 = b.种类 And a.编号 = b.页面 And c.ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", lngTMP)
        If Not rsTemp.EOF Then
            objEPRFileInfo.格式 = zlCommFun.NVL(rsTemp("格式").Value)
            objEPRFileInfo.SetFormat edtThis, objEPRFileInfo.格式
            edtThis.ResetWYSIWYG
        End If
                    
        If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile
        If gobjFSO.FileExists(strZipFile) Then gobjFSO.DeleteFile strZipFile
                
        '将定位到初始位置
        edtThis.Range(1, 1).Selected
        
        edtThis.ForceEdit = False
        edtBuff.ForceEdit = False
        
        ShowSharePageHistory = True
    End If
    
    edtThis.UnFreeze
    edtThis.RefreshTargetDC
    edtThis.ReadOnly = True
    edtThis.ReadOnly = True
    
    Set objEPRFileInfo = Nothing
    
    If Trim(strTmpClipboard) <> "" Then '恢复粘贴板内容
        DoEvents
        Clipboard.SetText strTmpClipboard
    Else
        DoEvents
        Clipboard.Clear
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    Set objEPRFileInfo = Nothing
End Function

'################################################################################################################
'## 功能：  获取系统默认临时路径
'################################################################################################################
Private Function GetSysTmpPath() As String
    GetSysTmpPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
End Function

'################################################################################################################
'## 功能：  是否可以撤销编辑
'################################################################################################################
Public Function CanUndo() As Boolean
    CanUndo = (p_Undo > 1) Or (p_Undo = 1 And UndoList(1).Filename <> "" And UBound(UndoList) = 1)
End Function

'################################################################################################################
'## 功能：  是否可以重做编辑
'################################################################################################################
Public Function CanRedo() As Boolean
    CanRedo = (p_Undo > 0) And (p_Undo < UBound(UndoList))
End Function

'################################################################################################################
'## 功能：  撤销编辑一步
'################################################################################################################
Public Sub Undo()
    If CanUndo = False Then Exit Sub
    If mblnChange Then AddUndoAction
    p_Undo = p_Undo - 1
    Me.Editor1.Tag = "Undo"
    If p_Undo = 0 Then p_Undo = 1
    Me.Document.ImportFromXMLFile Me.Editor1, UndoList(p_Undo).Filename, False, True
    Me.Editor1.Range(UndoList(p_Undo).SelStart, UndoList(p_Undo).SelEnd).Selected
    Me.Editor1.Tag = ""
    mblnChange = False
    DT1 = Now

    On Error Resume Next
    '刷新提纲列表
    RefCompends
    Editor1_SelChange Me.Editor1.ViewMode, UndoList(p_Undo).SelStart, UndoList(p_Undo).SelEnd   '手工刷新示范词句
End Sub

'################################################################################################################
'## 功能：  重做编辑一步
'################################################################################################################
Public Sub Redo()
    If CanRedo = False Then Exit Sub
    p_Undo = p_Undo + 1
    Me.Editor1.Tag = "Redo"
    Me.Document.ImportFromXMLFile Me.Editor1, UndoList(p_Undo).Filename, False, True
    Me.Editor1.Range(UndoList(p_Undo).SelStart, UndoList(p_Undo).SelEnd).Selected
    mblnChange = False
    Me.Editor1.Tag = ""
    DT1 = Now

    On Error Resume Next
    '刷新提纲列表
    RefCompends
    Editor1_SelChange Me.Editor1.ViewMode, UndoList(p_Undo).SelStart, UndoList(p_Undo).SelEnd   '手工刷新示范词句
End Sub

'################################################################################################################
'## 功能：  清除所有Undo队列临时文件
'################################################################################################################
Public Sub ClearUndoList()
    Dim i As Long
    For i = 1 To UBound(UndoList)
        If gobjFSO.FileExists(UndoList(i).Filename) Then gobjFSO.DeleteFile UndoList(i).Filename, True
    Next
    ReDim UndoList(1 To 1) As UndoInfo
    p_Undo = 0
    DT1 = Now
End Sub

'################################################################################################################
'## 功能：  清除无用的Undo列表（用于Undo后又输入了其他文字时，清除Undo后面的列表）
'################################################################################################################
Private Sub ClearNoUseUndoList()
    If mblnAutosave Then
        If p_Undo < 1 Then Exit Sub
        Dim i As Long
        If p_Undo < UBound(UndoList) Then
            For i = p_Undo + 1 To UBound(UndoList)
               '清除文件
               If gobjFSO.FileExists(UndoList(i).Filename) Then gobjFSO.DeleteFile UndoList(i).Filename, True
            Next
            ReDim Preserve UndoList(1 To p_Undo) As UndoInfo
        End If
    End If
End Sub

'################################################################################################################
'## 功能：  增加一个历史文件
'################################################################################################################
Public Sub AddUndoAction()
    If Me.Document Is Nothing Then Exit Sub
    If mblnAutosave = False Then Exit Sub

    Dim i As Long, j As Long, k As Long, strF As String
    '获取一个随机文件名
    Do
        '其中，采用在文件名中包含当前光标位置的方式
        k = Val(gfrmPublic.Tag)
        strF = GetSysTmpPath & "\EPRUndo_" & App.ThreadID & "_" & k & "_" & CLng(Rnd(Timer) * 1000) & ".xml"

        j = j + 1
        If j = 100 Then
            MsgBox "储存临时文件时出错！无法保存临时文件", vbOKOnly + vbInformation, gstrSysName
            Exit Sub
        End If
        k = k + 1
        gfrmPublic.Tag = k
    Loop While gobjFSO.FileExists(strF)

    ClearNoUseUndoList
    If Me.Document.ExportToXMLFile(Me.Editor1, strF) Then
        If UBound(UndoList) = mlngUndoLimit + 1 Then
            '已经到达存储限制

            If gobjFSO.FileExists(UndoList(1).Filename) Then gobjFSO.DeleteFile UndoList(1).Filename    '清除第一个文件
            For i = 1 To UBound(UndoList) - 1
                UndoList(i).Filename = UndoList(i + 1).Filename
                UndoList(i).SelStart = UndoList(i + 1).SelStart
                UndoList(i).SelEnd = UndoList(i + 1).SelEnd
            Next

            p_Undo = mlngUndoLimit + 1
        Else
            p_Undo = p_Undo + 1
            ReDim Preserve UndoList(1 To p_Undo) As UndoInfo
        End If
        UndoList(p_Undo).Filename = strF
        UndoList(p_Undo).SelStart = Me.Editor1.Selection.StartPos
        UndoList(p_Undo).SelEnd = Me.Editor1.Selection.EndPos
        mblnChange = False
        DT1 = Now
    End If
End Sub

'################################################################################################################
'## 功能：  刷新当前编辑文档的报告图像，用于采集站在采集新的图像之后调用
'################################################################################################################
Public Sub RefPacsPic()
    mfrmPacsPic.zlRefresh Document.EPRPatiRecInfo.医嘱id, Document.EPRFileInfo.lngModule
    mfrmPacsPic.Tag = "Loaded"
End Sub

'################################################################################################################
'## 功能：  清除当前编辑文档的报告图像刷新标志，用于新文档调用时可以刷新
'################################################################################################################
Public Sub ClsPacsPic()
    mfrmPacsPic.Tag = ""
End Sub

'################################################################################################################
'## 功能：  重新计算文档页数
'## 参数：  blnEstopNote-禁止增页提醒，在批量调入文件时，都应禁止
'################################################################################################################
Public Sub RecountPage(Optional blnEstopNote As Boolean)
    Dim lngPageCount As Long
    
    If mblnAutoPageCount = False Then Exit Sub
    
    If Me.Visible = False Then Exit Sub                 '不可见时不处理
    If Me.Editor1.ReadOnly Then Exit Sub                '只读时不处理
    If Me.Editor1.ViewMode <> cprNormal Then Exit Sub   '非普通模式不处理
    If Me.Editor1.Tag <> "" Then Exit Sub               '在批处理过程中不处理
    If Me.Editor1.InProcessing Then Exit Sub            '同样表示在处理过程中
    
    lngPageCount = Me.Editor1.PageCount
    Call Me.Editor1.DoVirtualPrint
    stbThis.Panels(2).Text = Editor1.CurrentLine & " 行,  " & Editor1.CurrentColumn & " 列,  共" & Editor1.LineCount & " 行,  共 " & Me.Editor1.PageCount & " 页"
    
    If blnEstopNote = True Then Exit Sub
    If mblnAutoPageNote = False Then Exit Sub
    If Me.Editor1.PageCount - lngPageCount <= 0 Then Exit Sub
    MsgBox "增加页数提醒：文档内容从" & lngPageCount & "页增加为" & Me.Editor1.PageCount & "页！", vbInformation, gstrSysName
    
End Sub
Private Function CanSetFormat() As Boolean
'功能：用于在审核模式下判断选中文字是否是当前版本，只要选中文字中有一个字不是当前版本新增的即返回False
Dim lStar As Long, lEnd As Long, i As Long, COLOR As OLE_COLOR
    With Editor1
        If .SelLength = 0 Then CanSetFormat = True: Exit Function
        lEnd = .Selection.EndPos: lStar = .Selection.StartPos
        For i = lStar To lEnd
            COLOR = .Range(i, i + 1).Font.ForeColor
            If Not Me.Document.IsNewCharColor(COLOR) Then Exit Function
            i = i + 1
        Next
    End With
    CanSetFormat = True
End Function

Private Sub Editor1_BeforeKeyDown(ViewMode As zlRichEditor.ViewModeEnum, KeyCode As Integer, Shift As Integer)
    If Me.Editor1.ReadOnly Then Exit Sub
'    Debug.Print KeyCode, Shift
    Select Case KeyCode
    Case 0, 16, 17, 18, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
        vbKeyEscape, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
        vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital

    Case Else
        If UBound(UndoList) = 1 And p_Undo = 0 And Me.Editor1.Tag = "" Then
            '首次保存
            AddUndoAction
        End If
    End Select
End Sub

Private Sub Editor1_Change(ViewMode As zlRichEditor.ViewModeEnum)
    If Me.Editor1.ReadOnly Then Exit Sub
    mblnChange = True
    If Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos + 1) = Chr(32) Then
        Dim blnForce As Boolean
        If Me.Editor1.Tag <> "" Then Exit Sub       '已经在其他处理过程中，不应去掉空格；否则导致词句加入的空格被去掉，产生混乱
        blnForce = Me.Editor1.ForceEdit
        Me.Editor1.Tag = "Change"
        Me.Editor1.ForceEdit = True
        Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos + 1) = ""
        Me.Editor1.ForceEdit = blnForce
        Me.Editor1.Tag = ""
    End If
    If (DateDiff("s", DT1, DT2) > mlngSaveInterval And Me.Editor1.Tag = "" And Me.Editor1.ForceEdit = False) Or (UBound(UndoList) = 1 And p_Undo = 0 And Me.Editor1.Tag = "" And Me.Editor1.ForceEdit = False) Then
        '保存
        AddUndoAction
    ElseIf Me.Editor1.Tag = "" Then
        ClearNoUseUndoList
    End If
    
    Call RecountPage
End Sub

Private Sub Editor1_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub Editor1_Resize(ViewMode As zlRichEditor.ViewModeEnum)
    If Me.ucPictureEditor1.Inited = False Then Exit Sub
    If Editor1.UIVisibled Then Editor1.ShowUIInterface
End Sub

Private Sub Editor1_UIClick(ViewMode As zlRichEditor.ViewModeEnum)
    If tblThis.Visible And tblThis.Enabled Then tblThis.SetFocus
End Sub

Private Sub Editor1_UIClose(UIhWnd As Long)
    If ucPacsImgCanvas1.Visible Then
        Dim lKey As Long
        lKey = Val(ucPacsImgCanvas1.Tag)
        
        If ucPictureEditor1.Visible Then
            ucPictureEditor1.Visible = False
            ucPictureEditor1.CloseMe ucPacsImgCanvas1.mMarkedPicture
            ucPacsImgCanvas1.LayoutPictures False
        End If
        
        ucPacsImgCanvas1.SavePictures
        If lKey > 0 Then Document.Tables("K" & lKey).Refresh Editor1
        ucPacsImgCanvas1.CloseMe
        If DkpThis.FindPane(ID_VIEW_SEGMENT).Closed = False Then DkpThis.ShowPane ID_VIEW_SEGMENT
        If DkpThis.FindPane(ID_VIEW_PHRASEDEMO).Closed = False Then DkpThis.ShowPane ID_VIEW_PHRASEDEMO
    Else
        If ucPictureEditor1.Visible Then
            ucPictureEditor1.Visible = False
            ucPictureEditor1.CloseMe
        End If
        
        If tblThis.Visible Then
            Dim lStart As Long, lEnd As Long
            lStart = Editor1.Selection.StartPos
            lEnd = Editor1.Selection.EndPos
            
            If Val(tblThis.Tag) <= 0 Then Exit Sub
            If tblThis.Modified And Me.Document.Tables.Count > 0 Then
                SaveUIToTable Me.Document.Tables("K" & tblThis.Tag), False
            End If
    '        Editor1.Range(lStart, lEnd).Selected
            
            tblThis.Visible = False
            mbEditInTable = False
            tblThis.SelectedCellKey = 0
            tblThis.Tag = ""
        End If
    End If
End Sub
Private Sub Editor1_UIOpen(UIhWnd As Long, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long)
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    bBeteenKeys = IsBetweenAnyKeys(Me.Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys And sKeyType = "T" Then
        If Me.Document.Tables("K" & lKey).TableType = tte_报告图片组 Then
            ucPacsImgCanvas1.ShowMe Me, UIhWnd, cbrThis, Me.Document.Tables("K" & lKey), lngLeft, lngTop, lngWidth, lngHeight
            ucPacsImgCanvas1.Tag = lKey
        Else
            If Val(tblThis.Tag) <= 0 Then Exit Sub
            ReadTableToUI Me.Document.Tables("K" & tblThis.Tag)
            lngWidth = tblThis.Width + 2 * lngLeft
            SetParent tblThis.hwnd, UIhWnd
            tblThis.Move lngLeft, lngTop
            tblThis.hWndBound = Editor1.hWndRTB
            tblThis.OffsetX = tblThis.Left + Editor1.UILeft - 390
            tblThis.OffsetY = tblThis.Top + Editor1.UITop
            tblThis.Visible = True
            mbEditInTable = True
        End If
    ElseIf bBeteenKeys And sKeyType = "P" Then
        If Me.Document.Pictures("K" & lKey).PictureType = EPRSignPicture Or Me.Document.Pictures("K" & lKey).PictureType = EPRPatiSign Then
            Me.Editor1.CloseUIInterface
            Exit Sub
        Else
            ucPictureEditor1.ShowMe Me, UIhWnd, cbrThis, Me.Document.Pictures("K" & lKey), lngLeft, lngTop, lngWidth, lngHeight, False
        End If
    End If
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, y As Single)
    Dim Popup As CommandBar
    Dim Control As CommandBarControl

    Set Popup = cbrThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
        Popup.ShowPopup
    End With
End Sub

Private Sub imgX_S_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX_S.Top = imgX_S.Top + y
    
    If imgX_S.Top < 1500 Then imgX_S.Top = 1500
    If Me.Height - imgX_S.Top - imgX_S.Height < 1000 Then imgX_S.Top = Me.Height - imgX_S.Height - 1000

    cbrThis.RecalcLayout
End Sub


Private Sub mfrmDocksymbol_GetPosFontSize()
    Dim lngSize As Long
    On Error Resume Next
    lngSize = Editor1.Selection.Font.Size
    mfrmDocksymbol.PicFontSize = lngSize
End Sub
'导入范文
Private Sub mfrmDocksymbol_InsertEPRDemo(lngEPRDemoID As Long)
    If Editor1.ReadOnly Then Exit Sub
    If lngEPRDemoID > 0 Then
                Call AddUndoPoint  '手动缓存
                Me.Document.ImportEPRDemo Me.Editor1, lngEPRDemoID
                Call ClearNoUseUndoList
                Call RecountPage(True)
            End If
End Sub

Private Sub mfrmDocksymbol_InsertPicSymbol(strInfor As String, picSy As StdPicture, strReturn As String)
    Dim blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    
    On Error GoTo ErrHandle
    If Editor1.ReadOnly Then Exit Sub
    If tblThis.Visible Then
        If strReturn = "" Then
            MsgBox "当前位置不支持此类特殊符号", vbInformation, gstrSysName: Exit Sub
        Else
            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = strReturn
            tblThis.Modified = True
            tblThis.Refresh False, True, tblThis.SelectedCellKey
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Else
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys = False Then
            Call AddUndoPoint  '手动缓存
            Editor1.Tag = "InsertPicSymbol"
            InsertPicture EPRFormulaPicture, picSy, picSy.Width, picSy.Height, strInfor
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
        Editor1.SetFocus
    End If
    Call RecountPage
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mfrmDocksymbol_InsertSymbol(strSymbol As String, intStrLen As Integer)
'intStrLen 是 strSymbol 内容字符串的长度，用来计算最终光标位置，除“医学单位、过敏药物、检验结果”为实际长度，其他默认为1
    Dim blnForce As Boolean, strFont As String, lngSelStart As Long, lngSymbolPos As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    On Error GoTo ErrHandle
    If Editor1.ReadOnly Then Exit Sub
    If tblThis.Visible Then
        If tblThis.SelectedCellKey > 0 Then
            If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False Then
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = tblThis.Cells("K" & tblThis.SelectedCellKey).Text & strSymbol
                tblThis.Modified = True
                tblThis.Refresh True, True, tblThis.SelectedCellKey
                tblThis_Resize tblThis.Width, tblThis.Height
            End If
        End If
        tblThis.SetFocus
    Else
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys = False And Editor1.Selection.Font.Protected = False Then
            Call AddUndoPoint  '手动缓存
            blnForce = Editor1.ForceEdit
            Editor1.ForceEdit = True
            Call Editor1_KeyDown(cprNormal, 32, 0)
            Editor1.Tag = "InsertstrSymbol"
            If Me.Editor1.AuditMode Then
                Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos).Selected
                '保留性属性（便于新增文本）
                On Error Resume Next
                Me.Editor1.OriginRTB.SelColor = Me.Document.GetNewCharColor(Me.Editor1.OriginRTB.SelColor)
                Me.Editor1.OriginRTB.SelStrikeThru = False
            End If
            strFont = Editor1.Selection.Font.Name
            lngSelStart = Editor1.SelStart
            Editor1.SelText = strSymbol
            lngSymbolPos = lngSelStart + intStrLen
            Editor1.SelStart = lngSymbolPos
            Editor1.Range(lngSelStart, lngSymbolPos).Font.Name = strFont 'Toshma字体会在保存、签名时报错，因为是UTF字符占位3个字节
            If intStrLen = 1 And Editor1.Range(lngSelStart, lngSymbolPos).Font.Name = "Tahoma" Then
                Editor1.Range(lngSelStart, lngSymbolPos).Font.Name = "宋体"
            End If
            Editor1.Selection.Font.Name = strFont
            
            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
        Editor1.SetFocus
    End If
    Call RecountPage
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mfrmDocksymbol_SetFouse()
    Me.Editor1.Enabled = False
    Me.Editor1.Enabled = True
End Sub

Private Sub mfrmHistoryReport_CopyClick(ByVal strContent As String)
    Clipboard.SetText strContent
    On Error Resume Next
    If cbrThis.FindControl(, ID_EDIT_PASTE).Enabled Then
        cbrThis.FindControl(, ID_EDIT_PASTE).Execute
    End If
End Sub

Private Sub mfrmHistoryReport_ReportCountChange(ByVal lngReportCount As Long)
On Error Resume Next
    DkpThis.FindPane(ID_VIEW_HISTORYREPORT).Title = "历史检查(" & lngReportCount & ")"
End Sub

Private Sub mfrmMainError_Location(ByVal Key As Long)

    Call Me.Document.Elements("K" & Key).Selected(Me.Editor1)

End Sub

Private Sub mfrmPacsPic_InsertPicture(pic As stdole.StdPicture, ByVal strUid As String, ByVal lngAdviceID As Long)
    If Editor1.ReadOnly Then Exit Sub
    If ucPacsImgCanvas1.Visible = False Then Exit Sub
    ucPacsImgCanvas1.AddPacsPicture pic, strUid, lngAdviceID
    Call RecountPage
End Sub

Private Sub mfrmPreview_PrintEpr(ByVal lngRecordId As Long)
    Me.Document.AfterPrinted Me.Document.EPRPatiRecInfo.ID
End Sub

Private Sub mfrmSegments_ModifiedOrDeleted(Action As Integer)
    If Me.Document.EditType = cprET_全文示范编辑 Then
        Err = 0: On Error Resume Next
        Call Me.Document.mfrmParent.RefreshList
    End If
End Sub

Private Sub mfrmSegments_RowDblClick(ByVal Row As XtremeReportControl.IReportRow)
    Dim rsTemp As New ADODB.Recordset, lngDemoId As Long
    Dim rsText As New ADODB.Recordset, strVSql As String
    Dim oCompend As cEPRCompend, lngStart As Long, lngTail As Long
    Dim oCell As cEPRCell, oElement As cEPRElement, oPicture As cEPRPicture
    Dim StrText As String, lngLen As Long, lngKey As Long, aryProp() As String
    
    If Me.Editor1.ViewMode <> cprNormal Or Me.Editor1.ReadOnly Then Exit Sub
    If Row Is Nothing Then Exit Sub
    lngDemoId = Row.Record(1).Value

    gstrSQL = "Select Id, 定义提纲id From 病历范文内容 Where 文件id = [1] And 对象类型 = 1 Order By 对象序号"
    strVSql = "Select Id, 对象类型, 对象属性, 内容文本, 是否换行, 要素名称, 诊治要素id, 替换域, 要素类型, 要素长度, 要素小数, 要素单位," & vbNewLine & _
            "       要素表示, 要素值域, 输入形态" & vbNewLine & _
            "From 病历范文内容" & vbNewLine & _
            "Where 文件id = [1] And 父id = [2]" & vbNewLine & _
            "Order By 对象序号"
    
    Me.Editor1.ForceEdit = True
    Me.Editor1.Tag = "mfrmSegments_RowDblClick"
    Me.Editor1.Freeze
    
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", lngDemoId)
    Do While Not rsTemp.EOF
        lngStart = 0: lngTail = 0
        For Each oCompend In Me.Document.Compends
            If oCompend.定义提纲ID > 0 And oCompend.定义提纲ID = Val("" & rsTemp!定义提纲ID) Then
                oCompend.GetPosition Me.Editor1, lngStart, lngTail
                Exit For
            End If
        Next
        If lngTail > 0 Then
            With Me.Editor1
                If .Range(lngTail - 2, lngTail) <> vbCrLf Then
                    .Range(lngTail, lngTail) = vbCrLf
                    With .Range(lngTail, lngTail + 2).Font
                        .Protected = False: .Hidden = False: .Strikethrough = False: .BackColor = tomAutoColor
                        .ForeColor = IIf(Me.Editor1.AuditMode, GetCharColor(Me.Document.目标版本, 0), tomAutoColor)
                    End With
                End If
                .SelStart = lngTail
            End With
            Set rsText = zlDatabase.OpenSQLRecord(strVSql, "提取信息", lngDemoId, CLng(rsTemp!ID))
            Do While Not rsText.EOF
                lngTail = Me.Editor1.SelStart
                Select Case rsText!对象类型
                Case 2  '文本
                    StrText = rsText!内容文本 & IIf(Val("" & rsText!是否换行) = 1, vbCrLf, "")
                    lngLen = Len(StrText)
                    With Me.Editor1
                        .Range(lngTail, lngTail) = StrText
                        With .Range(lngTail, lngTail + lngLen).Font
                            .Protected = False: .Hidden = False: .Strikethrough = False: .BackColor = tomAutoColor
                            .ForeColor = IIf(Me.Editor1.AuditMode, GetCharColor(Me.Document.目标版本, 0), tomAutoColor)
                        End With
                        .Range(lngTail + lngLen, lngTail + lngLen).Selected
                    End With
                Case 3  '表格
                    If Me.Document.EditType = cprET_全文示范编辑 Or Me.Document.EditType = cprET_单病历编辑 Then
                        lngKey = Me.Document.Tables.Add
                        With Me.Document.Tables("K" & lngKey)
                            Call .GetTableFromDB(cprET_全文示范编辑, lngDemoId, rsText!ID, False)
                            .ID = 0: .文件ID = 0: .父ID = 0: .开始版 = Me.Document.目标版本
                            For Each oCell In .Cells
                                oCell.ID = 0: oCell.文件ID = 0: oCell.父ID = 0: oCell.开始版 = Me.Document.目标版本
                            Next
                            For Each oElement In .Elements
                                oElement.ID = 0: oElement.文件ID = 0: oElement.父ID = 0: oElement.开始版 = Me.Document.目标版本
                                If oElement.替换域 = 1 And Me.Document.EditType = cprET_单病历编辑 Then
                                    oElement.内容文本 = GetReplaceEleValue(oElement.要素名称, _
                                        Me.Document.EPRPatiRecInfo.病人ID, _
                                        Me.Document.EPRPatiRecInfo.主页ID, _
                                        Me.Document.EPRPatiRecInfo.病人来源, _
                                        Me.Document.EPRPatiRecInfo.医嘱id, _
                                        Me.Document.EPRPatiRecInfo.婴儿)
                                        For Each oCell In .Cells
                                            If oCell.ElementKey = oElement.Key Then oCell.内容文本 = oElement.内容文本: Exit For
                                        Next
                                End If
                            Next
                            For Each oPicture In .Pictures
                                oPicture.ID = 0: oPicture.文件ID = 0: oPicture.父ID = 0: oPicture.开始版 = Me.Document.目标版本
                            Next
                            .InsertIntoEditor Me.Editor1, lngTail
                        End With
                    End If
                Case 4  '元素
                    lngKey = Me.Document.Elements.Add
                    With Me.Document.Elements("K" & lngKey)
                        .ID = 0
                        .内容文本 = "" & rsText!内容文本
                        .要素名称 = "" & rsText!要素名称
                        .诊治要素ID = Val("" & rsText!诊治要素ID)
                        .替换域 = Val("" & rsText!替换域)
                        .要素类型 = Val("" & rsText!要素类型)
                        .要素长度 = Val("" & rsText!要素长度)
                        .要素小数 = Val("" & rsText!要素小数)
                        .要素单位 = "" & rsText!要素单位
                        .要素表示 = Val("" & rsText!要素表示)
                        .要素值域 = "" & rsText!要素值域
                        .输入形态 = Val("" & rsText!输入形态)
                        .是否换行 = Val("" & rsText!是否换行)
                        If .替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                            .内容文本 = GetReplaceEleValue(.要素名称, _
                                Me.Document.EPRPatiRecInfo.病人ID, _
                                Me.Document.EPRPatiRecInfo.主页ID, _
                                Me.Document.EPRPatiRecInfo.病人来源, _
                                Me.Document.EPRPatiRecInfo.医嘱id, _
                                Me.Document.EPRPatiRecInfo.婴儿)
                        End If
                        .开始版 = Me.Document.目标版本
                        .InsertIntoEditor Me.Editor1, lngTail, , True
                    End With
                Case 5  '图形
                    If Me.Document.EditType = cprET_全文示范编辑 Or Me.Document.EditType = cprET_单病历编辑 Then
                        lngKey = Me.Document.Pictures.Add
                        With Me.Document.Pictures("K" & lngKey)
                            Call .GetPictureFromDB(cprET_全文示范编辑, lngDemoId, rsText!ID, False)
                            .ID = 0: .文件ID = 0: .父ID = 0
                            .InsertIntoEditor Me.Editor1, lngTail, True
                        End With
                    End If
                Case 7  '诊断
                    aryProp = Split("" & rsText!对象属性, ";")
                    lngKey = Me.Document.Diagnosises.Add
                    With Me.Document.Diagnosises("K" & lngKey)
                        .ID = 0
                        .描述 = "" & rsText!内容文本
                        .类型 = Val(aryProp(0))
                        .中医 = Val(aryProp(1))
                        .疾病id = Val(aryProp(2))
                        .诊断id = Val(aryProp(3))
                        .证候id = Val(aryProp(4))
                        .疑诊 = Val(aryProp(5))
                        .日期 = Format(Now(), "yyyy-mm-dd hh:mm:ss")
                        .开始版 = Me.Document.目标版本
                        .InsertIntoEditor Me.Editor1, lngTail, True
                    End With
                End Select
                rsText.MoveNext
            Loop
        End If
        rsTemp.MoveNext
    Loop
    Me.Editor1.ForceEdit = False
    Me.Editor1.Tag = ""
    Me.Editor1.UnFreeze
    Call RecountPage
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.Editor1.ForceEdit = False
End Sub

Private Sub mfrmSentenceDetailed_ShiftFocus()
    Me.Editor1.Enabled = False
    Me.Editor1.Enabled = True
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    Me.Document.AfterPrinted Me.Document.EPRPatiRecInfo.ID
End Sub

Private Sub picHistoryInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If X > 0 And X < picHistoryInfo.ScaleWidth And y > 0 And y < picHistoryInfo.ScaleHeight Then
        If picHistoryInfo.Tag = "" Then
            SetCapture picHistoryInfo.hwnd
            picHistoryInfo.Cls
            picHistoryInfo.BackColor = &HD2BDB6    ' &HD8D5D4 ' &HD2BDB6
            picHistoryInfo.Line (0, 0)-(picHistoryInfo.ScaleWidth - Screen.TwipsPerPixelX, picHistoryInfo.ScaleHeight - Screen.TwipsPerPixelY), &H6A240A, B
            picHistoryInfo.Tag = "Captured"
        End If
    Else
        ReleaseCapture
        picHistoryInfo.Cls
        picHistoryInfo.BackColor = &H8000000F
        picHistoryInfo.Line (0, 0)-(picHistoryInfo.ScaleWidth - Screen.TwipsPerPixelX, picHistoryInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
        picHistoryInfo.Tag = ""
    End If
End Sub

Private Sub picHistoryInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    ReleaseCapture
    picHistoryInfo.Cls
    picHistoryInfo.BackColor = &H8000000F
    picHistoryInfo.Line (0, 0)-(picHistoryInfo.ScaleWidth - Screen.TwipsPerPixelX, picHistoryInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
    picHistoryInfo.Tag = ""
End Sub

Private Sub picHistoryInfo_Resize()
    picHistoryInfo.Cls
    picHistoryInfo.BackColor = &H8000000F
    picHistoryInfo.Line (0, 0)-(picHistoryInfo.ScaleWidth - Screen.TwipsPerPixelX, picHistoryInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
    picHistoryInfo.Tag = ""
End Sub

Private Sub picPane_Resize()
    On Error Resume Next
    
    picHistoryInfo.Move 15, 15, picPane.Width - 30
    edtThis.Move 15, picHistoryInfo.Top + picHistoryInfo.Height, picPane.Width - 30, picPane.Height - 15 - (picHistoryInfo.Top + picHistoryInfo.Height)
End Sub

Private Sub tblThis_CancelEdit()
    Editor1.Modified = True
End Sub

Private Sub tblThis_SelectionChange(ByVal lrow As Long, ByVal lCol As Long)
    Dim lngKey As Long
    If Me.Editor1.AuditMode = True Then Exit Sub
    lngKey = Val(tblThis.Cell(lrow, lCol).Tag)
    If ucPictureEditor1.Visible Then ucPictureEditor1.CloseMe: tblThis.Modified = True
    If lngKey > 0 Then
        If Not tblThis.Cell(lrow, lCol).Picture Is Nothing Then
            '图片
            If Val(tblThis.Tag) > 0 Then
                '编辑图片
                Dim LL As Long, lT As Long, lW As Long, lH As Long
                tblThis.Cell(lrow, lCol).GetCellPictureBorder LL, lT, lW, lH
                
                    ucPictureEditor1.ShowMe Me, tblThis.hwnd, cbrThis, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey), _
                        LL, lT, lW, lH, True, Me.Document.Tables("K" & tblThis.Tag)
                
                mblnChange = True
                tblThis.Modified = True
            Else
                If ucPictureEditor1.Visible Then ucPictureEditor1.CloseMe: tblThis.Modified = True
            End If
        Else
            If ucPictureEditor1.Visible Then ucPictureEditor1.CloseMe: tblThis.Modified = True
        End If
    Else
        If ucPictureEditor1.Visible Then ucPictureEditor1.CloseMe: tblThis.Modified = True
    End If
End Sub

Private Sub tblThis_ModifyProtected(ByVal lKey As Long)
    Dim lLeft As Long, lTOp As Long, lRight As Long, lBottom As Long, lngKey As Long

    tblThis.Cells("K" & lKey).GetCellTextBorder lLeft, lTOp, lRight, lBottom

    lngKey = Val(tblThis.Cells("K" & lKey).Tag)
    If lngKey > 0 Then
        If tblThis.Cells("K" & lKey).Picture Is Nothing Then
            '诊治要素
            If Val(tblThis.Tag) > 0 Then
                If Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).替换域 = 2 Then
                    '字典项目
                    mfrmDicSelect.Tag = lngKey
                    mfrmDicSelect.ShowMe Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).要素名称, Me.Left + Editor1.Left + Editor1.UILeft + tblThis.Left + lLeft * 15 + 30, _
                        Me.Top + Editor1.Top + Editor1.UITop + tblThis.Top + lBottom * 15 + 500, IIf(mintStyle = -1, 0, mintStyle), Me, Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).内容文本
                Else
                    '诊治要素
                    mfrmModElement.Tag = lngKey
                    mfrmModElement.ShowMe Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey), _
                        Me.Left + Editor1.Left + Editor1.UILeft + tblThis.Left + lLeft * 15 + 30, _
                        Me.Top + Editor1.Top + Editor1.UITop + tblThis.Top + lBottom * 15 + 500, IIf(mintStyle = -1, 0, mintStyle), Me, Me.Document.EditType
                End If
            End If
        Else
            '图片
            If Me.Editor1.AuditMode = True Then Exit Sub
            If Val(tblThis.Tag) > 0 Then
                '编辑指定的图片
                If Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey).PictureType = EPRMarkedPicture Then
                    '编辑图片
                    Dim LL As Long, lT As Long, lW As Long, lH As Long
                    tblThis.Cells("K" & lKey).GetCellPictureBorder LL, lT, lW, lH
                    ucPictureEditor1.ShowMe Me, tblThis.hwnd, cbrThis, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey), _
                        LL, lT, lW, lH, True, Me.Document.Tables("K" & tblThis.Tag)
                ElseIf Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey).PictureType = EPROutPicture Then
                    cPicEditor.ShowPicEditor glngSys, gcnOracle, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey).OrigPic, _
                        lngKey, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey).保留对象, Me, False
                    '这里外部图片的保存在cPicEditor对象的pOK事件中处理！
                End If
            End If
        End If
        mblnChange = True
    End If
End Sub

Private Sub tblThis_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Dim Popup As CommandBar
        Dim cbpPopup As CommandBarPopup
        Dim Control As CommandBarControl
        Dim lngKey As Long

        Set Popup = cbrThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set Control = .Add(xtpControlButton, ID_EDIT_CUT, "剪切(&X)")
            Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
            Set Control = .Add(xtpControlButton, ID_EDIT_PASTE, "粘贴(&V)")
            Set Control = .Add(xtpControlButton, ID_EDIT_DELETE, "删除(&D)")
            Set Control = .Add(xtpControlButton, ID_TABLE_MERGE, "合并单元格(&M)"): Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, ID_TABLE_DELETETABLE, "删除表格(&T)")

            If tblThis.SelStartRow = tblThis.SelEndRow And tblThis.SelStartCol = tblThis.SelEndCol Then
                If Val(tblThis.Tag) > 0 And tblThis.SelectedCellKey > 0 Then
                    '编辑指定的图片
                    If Not tblThis.Cells("K" & tblThis.SelectedCellKey).Picture Is Nothing Then
                        lngKey = Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag)
                        Set Control = .Add(xtpControlButton, ID_EDIT_MARKEDPIC, "标记修改(&M)"): Control.BeginGroup = True
                        If Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey).PictureType = EPROutPicture Then
                            .Add xtpControlButton, ID_EDIT_OUTERPIC, "底图处理(&D)"
                        End If
                    End If
                End If
            End If

            Set cbpPopup = .Add(xtpControlPopup, ID_TABLE_CELLALIGNMENT, "单元格对齐方式")
            cbpPopup.CommandBar.SetTearOffPopup "单元格对齐方式", ID_TABLE_CELLALIGNMENT, 100
            cbpPopup.CommandBar.SetPopupToolBar True
            cbpPopup.BeginGroup = True
            cbpPopup.CommandBar.Width = 70
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT1, "靠上左对齐"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT2, "靠上居中"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT3, "靠上右对齐"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT4, "中部左对齐"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT5, "中部居中"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT6, "中部右对齐"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT7, "靠下左对齐"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT8, "靠下居中"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT9, "靠下右对齐"

            Set Control = .Add(xtpControlButton, ID_TABLE_PROPERTY, "表格属性(&R)..."): Control.BeginGroup = True

            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub tblThis_Resize(ByVal lWidth As Long, ByVal lHeight As Long)
    '编辑过程中动态改变表格大小
    Editor1.ResizeUIInterface lWidth, lHeight
    If ucPictureEditor1.Visible Then ucPictureEditor1.Visible = False
    If Val(tblThis.Tag) > 0 Then
        Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, bFinded As Boolean, bNeeded As Boolean
        Dim lW As Long
        bFinded = FindKey(Editor1, "T", tblThis.Tag, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then
            Editor1.InProcessing = True
            lW = Me.Editor1.PaperWidth - Me.Editor1.MarginLeft - Me.Editor1.MarginRight - Me.ScaleX(Me.Editor1.Range(lSE, lES).Para.LeftIndent + Me.Editor1.Range(lSE, lES).Para.FirstLineIndent, vbPixels, vbTwips) - 130
            picTMP.Width = IIf(tblThis.Width > lW, lW, tblThis.Width)
            picTMP.Height = tblThis.Height
            tblThis.DrawToDC picTMP.hDC
            picTMP.Picture = picTMP.Image
            '刷新并选中改表格图片
            Me.Document.Tables("K" & tblThis.Tag).Refresh Editor1, picTMP.Picture, True
'            Editor1.InProcessing = True
'            Editor1.Range(lSE, lES).Selected
'            Editor1.RefreshUIInterface
            Editor1.InProcessing = False
            mblnChange = True
        End If
    End If
End Sub

Private Sub tmrAutoSaveEPR_Timer()
    DT2_EPR = Now
    If mblnAutoSaveEPR Then
        If DateDiff("n", DT1_EPR, DT2_EPR) > mlngSaveIntervalEPR Then
            '自动保存文件
            If (Editor1.Modified) And (Me.Document.目标版本 <= 16) Then
                Call SaveEMRDoc(True)
                DT1_EPR = Now
            End If
        End If
    End If
End Sub

'################################################################################################################
'## 功能：  定时自动缓存
'################################################################################################################
Private Sub tmrThis_Timer()
    If Me.Document Is Nothing Then Exit Sub
    If mblnAutosave Then DT2 = Now
End Sub

'################################################################################################################
'## 功能：  手工增加还原点
'################################################################################################################
Private Sub AddUndoPoint()
    If Me.Document Is Nothing Then Exit Sub
    If mblnChange Then AddUndoAction
End Sub
'################################################################################################################
'## 功能：  刷新提纲
'################################################################################################################
Public Sub RefCompends()
    Document.Compends.UpdateOrdersFromText Editor1
    Document.Compends.FillTree mfrmCompends.Tree
    Call RefSentenceList
End Sub

'################################################################################################################
'## 功能：  更新提纲相关的示范词句
'################################################################################################################
Public Sub RefSentenceList()
    Dim lngCompend As Long, lngPatient As Long, lngVisit As Long, lngAdvice As Long
    Dim blnForce As Boolean         '文件定义时，强制刷新
    Dim strLimit As String          '由病种,要素引起的词句限制
    If mfrmCompends.Tree.SelectedItem Is Nothing Then Exit Sub
    
    If Me.Document.EditType = cprET_病历文件定义 Then
        If mfrmCompends.Tree.SelectedItem Is Nothing Then
            lngCompend = 0
        Else
            lngCompend = Me.Document.Compends(mfrmCompends.Tree.SelectedItem.Key).ID
        End If
        lngPatient = 0: lngVisit = 0: lngAdvice = 0
        blnForce = True
    ElseIf Me.Document.EditType = cprET_全文示范编辑 Then
        If mfrmCompends.Tree.SelectedItem Is Nothing Then
            lngCompend = 0
        Else
            lngCompend = Me.Document.Compends(mfrmCompends.Tree.SelectedItem.Key).定义提纲ID
        End If
        lngPatient = 0: lngVisit = 0: lngAdvice = 0
        blnForce = False
    Else
        If mfrmCompends.Tree.SelectedItem Is Nothing Then
            lngCompend = 0
        Else
            lngCompend = Me.Document.Compends(mfrmCompends.Tree.SelectedItem.Key).定义提纲ID
        End If
        lngPatient = Me.Document.EPRPatiRecInfo.病人ID
        lngVisit = Me.Document.EPRPatiRecInfo.主页ID
        lngAdvice = Me.Document.EPRPatiRecInfo.医嘱id
        blnForce = False
    End If
    strLimit = MakeSentenceLimit(lngCompend)
    Call mfrmSentenceDetailed.zlRefFromCompend(Me, lngCompend, lngPatient, lngVisit, lngAdvice, blnForce, strLimit)
End Sub

'################################################################################################################
'## 功能：  显示本编辑主窗体
'##
'## 参数：  frmParent       :父窗体
'##         blnFirst        :是否是第一次打开窗体（多文件编辑时使用，切换当前文件时置为False，表示不刷新组合文档）
'##         blnCanPrint     :是否允许预览、打印
'################################################################################################################
Public Sub ShowMe(frmParent As Object, Optional blnFirst As Boolean = True, Optional blnCanPrint As Boolean = True, Optional ByVal byteStyle As Integer)
    '设置窗体显示状态
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    mblnIsMultiMode = Me.Document.IsMultiEPRDoc
    mblnCanPrint = blnCanPrint
    mblnPrecess = False
    mintStyle = byteStyle
    mblnPatiSign = HavedPatiSign
    Call SetStateInfo
    Select Case Document.EditType
        Case cprET_单病历编辑
            stbThis.Panels(5).Visible = False
            DkpThis.FindPane(ID_VIEW_STRUCTURE).Close
            CommBar(ID_BAR_SIGN).Visible = True
            Call mfrmSegments.zlRefresh(Me)
        Case cprET_单病历审核
            Editor1.AuditMode = True
            DkpThis.FindPane(ID_VIEW_STRUCTURE).Close
            CommBar(ID_BAR_SIGN).Visible = True
            Call mfrmSegments.zlRefresh(Me)
        Case cprET_病历文件定义
            stbThis.Panels(5).Visible = False
            DkpThis.FindPane(ID_VIEW_SEGMENT).Close
            CommBar(ID_BAR_SIGN).Visible = False
        Case cprET_全文示范编辑
            stbThis.Panels(5).Visible = False
            If Document.EPRDemoInfo.性质 <> 0 Then
                CommBar(ID_BAR_FORMAT).Visible = False
                CommBar(ID_BAR_FORMAT).Delete
            End If
            CommBar(ID_BAR_SIGN).Visible = False
            Call mfrmSegments.zlRefresh(Me)
    End Select

    If Document.EditType = cprET_单病历编辑 Or Document.EditType = cprET_单病历审核 Then
        If gobjESign Is Nothing Then
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            Call gobjESign.Initialize(gcnOracle, glngSys)
        End If
        If Not gobjESign Is Nothing Then
            mblnEnPtSign = gobjESign.EnabledPatiSign
        End If
    End If
        
    If Me.Document.EPRFileInfo.种类 = cpr诊疗报告 Then
        mfrmPacsPic.zlRefresh Document.EPRPatiRecInfo.医嘱id, Document.EPRFileInfo.lngModule
        DkpThis.FindPane(ID_VIEW_HISTORYREPORT).Title = "历史检查(" & mfrmHistoryReport.zlRefresh(Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.科室ID, Document.EPRPatiRecInfo.ID) & ")"
    Else
        DkpThis.FindPane(ID_VIEW_PACSPIC).Close
        DkpThis.FindPane(ID_VIEW_HISTORYREPORT).Close
    End If

    If mblnIsMultiMode And blnFirst Then
        If mfrmMultiDocView Is Nothing Then Set mfrmMultiDocView = New frmMultiDocView
        '多文件组合查阅窗体的初始化
        mfrmMultiDocView.InitData Me, Me.Document, Me.Document.EPRPatiRecInfo.ID
        DkpThis.ShowPane ID_VIEW_MULTIDOCVIEW
    End If
    
    If mblnIsMultiMode Then
        mblnExistHistroy = ShowSharePageHistory(Me.Document, mintSharePages)
    End If
    If mblnExistHistroy = False Then picPane.Visible = mblnExistHistroy
    
    If mblnChildMode Then
        stbThis.Visible = False
        picPatiInfo.Visible = False
    Else
        stbThis.Visible = True
        picPatiInfo.Visible = True
    End If
    Call Me.Editor1.DoVirtualPrint
    stbThis.Panels(2).Text = Editor1.CurrentLine & " 行,  " & Editor1.CurrentColumn & " 列,  共" & Editor1.LineCount & " 行"
    If mblnAutoPageCount Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & ",  共 " & Me.Editor1.PageCount & " 页"
    
    Dim intLoop As Integer
    
    For intLoop = 1 To Me.Document.Elements.Count
        If Me.Document.Elements(intLoop).签名要素 Then
            mbln签名要素 = True
            Exit For
        End If
    Next
    
    If Me.Document.EditType = cprET_单病历编辑 Then  '跟据病种、要素已选项检查未选要素可用选项
        Call ReadElementLimit
        Call CheckLastDiagnose '审核情况下不能对既有的要素发生变动
    Else '避免出错，定义数组维数
        ReDim mEleLimit(0) As EleLimit
    End If
    '为刷新范文列表传递条件
    Call mfrmDocksymbol.SetItems(Document.EPRPatiRecInfo.文件ID, Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.医嘱id)
    mbln返修处理 = False
    mblnFBContentChanged = True
    If Document.EPRPatiRecInfo.病历种类 = cpr诊断文书 Then
        strSQL = "Select b.反馈内容 From 疾病申报记录 A, 疾病报告反馈 B" & vbNewLine & _
             "Where a.文件id = b.文件id and A.处理状态=4 And a.文件id = [1] And B.登记时间 = (Select Max(登记时间) From 疾病报告反馈 Where 文件id = [1])"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Me.Document.EPRPatiRecInfo.ID)
        If rs.RecordCount > 0 Then
            mbln返修处理 = True
            txtFeedBack.Text = NVL(rs!反馈内容)
            cbrThis.Item(2).Controls.Find(, 99999901).Visible = True
            cbrThis.Item(2).Controls.Find(, 99999902).Visible = True
            cbrThis.Item(2).Controls.Find(, 99999903).Visible = True
            cbrThis.Item(2).Controls.Find(, 99999904).Visible = True
        End If
    End If
    If mintStyle = vbModeless Or mintStyle = vbModal Then
        Me.Show mintStyle, frmParent
    End If
End Sub
Private Sub ReadElementLimit()
Dim rsTemp As ADODB.Recordset, i As Integer
Dim strLastDiagnose As String
'结构数组第0维不赋值,从第一维开始认为有效
    On Error GoTo errHand
    If Not (Document.EPRFileInfo.种类 = cpr门诊病历 Or Document.EPRFileInfo.种类 = cpr住院病历) Then Exit Sub '只针对门诊病历和住院病历
    
    gstrSQL = "Select a.变动原因,a.原因要件id, a.原因要素, a.原因内容, b.变动结果, b.病历提纲id 结果提纲id, b.结果要件id, b.结果要素, b.结果值域,b.原始值域" & vbNewLine & _
                "From 病历变动原因 A, 病历变动结果 B" & vbNewLine & _
                "Where a.病历文件id = [1] And b.变动原因id = a.Id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取变动", Document.EPRFileInfo.ID)
    ReDim mEleLimit(rsTemp.RecordCount) As EleLimit
    For i = 1 To rsTemp.RecordCount
        mEleLimit(i).变动原因 = rsTemp!变动原因
        mEleLimit(i).原因要件id = NVL(rsTemp!原因要件id, 0)
        mEleLimit(i).原因要素 = NVL(rsTemp!原因要素, "")
        mEleLimit(i).原因内容 = NVL(rsTemp!原因内容, "")
        mEleLimit(i).变动结果 = rsTemp!变动结果
        mEleLimit(i).结果提纲id = NVL(rsTemp!结果提纲id, 0)
        mEleLimit(i).结果要件id = NVL(rsTemp!结果要件id, 0)
        mEleLimit(i).结果要素 = NVL(rsTemp!结果要素, "")
        mEleLimit(i).结果值域 = NVL(rsTemp!结果值域, "")
        mEleLimit(i).原始值域 = NVL(rsTemp!原始值域, "")
        rsTemp.MoveNext
    Next
    
    strLastDiagnose = GetReplaceEleValue("最后诊断ID", Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.病人来源, Document.EPRPatiRecInfo.医嘱id, Me.Document.EPRPatiRecInfo.婴儿)
    mlDiseaseID = Val(Split(strLastDiagnose, "|")(0))
    mlDiagnoseID = Val(Split(strLastDiagnose, "|")(1))
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub CheckLastDiagnose()
'功能：在病历打开时对病种限制要素选项进行检查和设置
Dim intLm As Integer, intEl, lKey As Long
    On Error GoTo errHand
    If Not (Document.EPRFileInfo.种类 = cpr门诊病历 Or Document.EPRFileInfo.种类 = cpr住院病历) Then Exit Sub '只针对门诊病历和住院病历
    
    For intLm = 1 To UBound(mEleLimit)
        Select Case mEleLimit(intLm).变动原因
            Case 1 '要素引起的变化不在此处理，在编辑时处理，修改情况下要素的值域已确定，审核情况下不可处理
            Case 2, 3 '病种引起的变化
                If (mlDiseaseID = mEleLimit(intLm).原因要件id Or mlDiagnoseID = mEleLimit(intLm).原因要件id) Then
                    For intEl = 1 To Document.Elements.Count
                        lKey = Document.Elements(intEl).Key
                        If Document.Elements("K" & lKey).要素名称 = mEleLimit(intLm).结果要素 And Document.Elements("K" & lKey).诊治要素ID = mEleLimit(intLm).结果要件id Then
                            Select Case mEleLimit(intLm).变动结果
                                Case 1 '引起要素选项变化
                                    If Document.Elements("K" & lKey).输入形态 = 1 Then
                                        If InStr(Document.Elements("K" & lKey).内容文本, "●") = 0 And InStr(Document.Elements("K" & lKey).内容文本, "■") = 0 Then
                                            '结果要素没有选项被选中才更新值域，内容，并刷新显示
                                            Document.Elements("K" & lKey).要素值域 = mEleLimit(intLm).结果值域
                                            If Document.Elements("K" & lKey).要素表示 = 2 Then
                                                Document.Elements("K" & lKey).内容文本 = "○" & Replace(mEleLimit(intLm).结果值域, ";", "○")
                                            Else
                                                Document.Elements("K" & lKey).内容文本 = "□" & Replace(mEleLimit(intLm).结果值域, ";", "□")
                                            End If
                                            Document.Elements("K" & lKey).Refresh Editor1
                                        End If
                                    Else                                                             '结果要素没有选项被选中，更新值域
                                        If Document.Elements("K" & lKey).内容文本 = "" Then Document.Elements("K" & lKey).要素值域 = mEleLimit(intLm).结果值域
                                    End If
                                Case 3 '删除要素
                                    If Document.Elements("K" & lKey).输入形态 = 1 Then
                                        If InStr(Document.Elements("K" & lKey).内容文本, "●") = 0 And InStr(Document.Elements("K" & lKey).内容文本, "■") = 0 Then
                                            Document.Elements("K" & lKey).DeleteFromEditor Editor1
                                            Exit For
                                        End If
                                    Else
                                        If Document.Elements("K" & lKey).内容文本 = "" Then Document.Elements("K" & lKey).DeleteFromEditor Editor1: Exit For
                                    End If
                            End Select
                        End If
                    Next
                End If
        End Select
    Next
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub UpdateSameELement(ByVal lKey As Long)
'功能：当前处于编辑状态，在要素选中时对联动要素进行选项检查和设置
Dim leKey As Long
Dim strElementName As String, strElementValue As String, lEleId As Long, intLm As Integer, intEl As Integer, intLn As Integer, blnLm As Boolean
    On Error GoTo errHand
    
    If Not (Document.EPRFileInfo.种类 = cpr门诊病历 Or Document.EPRFileInfo.种类 = cpr住院病历) Then Exit Sub '只针对门诊病历和住院病历
    If Document.EditType = cprET_单病历审核 Then Exit Sub
    
    strElementName = Document.Elements("K" & lKey).要素名称
    strElementValue = Document.Elements("K" & lKey).内容文本
    lEleId = Document.Elements("K" & lKey).诊治要素ID
    
    For intLm = 1 To UBound(mEleLimit)
        If mEleLimit(intLm).变动原因 = 4 Then
            If strElementName = mEleLimit(intLm).原因要素 And lEleId = mEleLimit(intLm).原因要件id Then
                For intEl = 1 To Document.Elements.Count
                    leKey = Document.Elements(intEl).Key
                    If Document.Elements("K" & leKey).要素名称 = mEleLimit(intLm).结果要素 And Document.Elements("K" & leKey).诊治要素ID = mEleLimit(intLm).结果要件id Then
                        If Document.Elements("K" & leKey).输入形态 = 1 Then
                            '结果要素没有选项被选中才更新值域，内容，并刷新显示
                            Document.Elements("K" & leKey).内容文本 = strElementValue
                            Document.Elements("K" & leKey).Refresh Editor1
                        Else
                            '结果要素没有选项被选中才更新值域，内容，并刷新显示
                            Document.Elements("K" & leKey).内容文本 = strElementValue
                            Document.Elements("K" & leKey).Refresh Editor1
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub CheckElementLimit(ByVal lKey As Long)
'功能：当前处于编辑状态，在要素选中时对联动要素进行选项检查和设置
Dim leKey As Long
Dim strElementName As String, strElementValue As String, lEleId As Long, intLm As Integer, intEl As Integer, intLn As Integer, blnLm As Boolean
Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean

    On Error GoTo errHand
    If Not (Document.EPRFileInfo.种类 = cpr门诊病历 Or Document.EPRFileInfo.种类 = cpr住院病历) Then Exit Sub '只针对门诊病历和住院病历
    If Document.EditType = cprET_单病历审核 Then Exit Sub
    strElementName = Document.Elements("K" & lKey).要素名称
    strElementValue = Document.Elements("K" & lKey).内容文本
    lEleId = Document.Elements("K" & lKey).诊治要素ID
    
    For intLm = 1 To UBound(mEleLimit)
        Select Case mEleLimit(intLm).变动原因
            Case 2, 3
            Case 1, 4
                If strElementName = mEleLimit(intLm).原因要素 And lEleId = mEleLimit(intLm).原因要件id Then
                    Select Case mEleLimit(intLm).变动结果
                        Case 1
                            '所选项满足原因条件:名称、ID、选项
                            If (strElementValue = mEleLimit(intLm).原因内容 Or InStr(strElementValue, "●" & mEleLimit(intLm).原因内容) > 0 Or InStr(strElementValue, "■" & mEleLimit(intLm).原因内容) > 0) Then
                                '原因要素被选中
                                For intEl = 1 To Document.Elements.Count
                                    leKey = Document.Elements(intEl).Key
                                    If Document.Elements("K" & leKey).要素名称 = mEleLimit(intLm).结果要素 And Document.Elements("K" & leKey).诊治要素ID = mEleLimit(intLm).结果要件id Then
                                        If Document.Elements("K" & leKey).输入形态 = 1 Then
                                            Document.Elements("K" & leKey).要素值域 = mEleLimit(intLm).结果值域
                                            If Document.Elements("K" & leKey).要素表示 = 2 Then
                                                Document.Elements("K" & leKey).内容文本 = "○" & Replace(mEleLimit(intLm).结果值域, ";", "  ○")
                                            Else
                                                Document.Elements("K" & leKey).内容文本 = "□" & Replace(mEleLimit(intLm).结果值域, ";", "  □")
                                            End If
                                            Document.Elements("K" & leKey).Refresh Editor1
                                        Else
                                            Document.Elements("K" & leKey).要素值域 = mEleLimit(intLm).结果值域
                                        End If
                                    End If
                                Next
                            ElseIf strElementValue = "" Or Not (InStr(strElementValue, "●" & mEleLimit(intLm).原因内容) > 0 Or InStr(strElementValue, "■" & mEleLimit(intLm).原因内容) > 0) Then
                                '原因要素没被选中 或没有选中原因要素的指定条件
                                For intEl = 1 To Document.Elements.Count
                                    leKey = Document.Elements(intEl).Key
                                    If Document.Elements("K" & leKey).要素名称 = mEleLimit(intLm).结果要素 And Document.Elements("K" & leKey).诊治要素ID = mEleLimit(intLm).结果要件id Then
                                        '检查是否被其它原因限制，如被限制，则不还原;可能是同一原因要素不同选项
                                        blnLm = False
                                        For intLn = 1 To UBound(mEleLimit)
                                            With Document.Elements("K" & leKey)
                                                If .要素名称 = mEleLimit(intLn).结果要素 And .诊治要素ID = mEleLimit(intLn).结果要件id _
                                                    And mEleLimit(intLn).变动结果 = 1 And .要素值域 = mEleLimit(intLn).结果值域 Then
                                                    If Document.Elements("K" & lKey).输入形态 = 1 Then '展开型
                                                        If InStr(strElementValue, "●") > 0 Then blnLm = True
                                                        If InStr(strElementValue, "■" & mEleLimit(intLn).原因内容) > 0 Then blnLm = True
                                                    Else                                               '下拉型
                                                        If InStr(strElementValue, mEleLimit(intLn).原因内容) > 0 Then blnLm = True
                                                    End If
                                                     If blnLm = True Then Exit For
                                                End If
                                            End With
                                        Next
                                    
                                        If Not blnLm Then
                                            If Document.Elements("K" & leKey).输入形态 = 1 Then
                                                Document.Elements("K" & leKey).要素值域 = mEleLimit(intLm).原始值域
                                                If Document.Elements("K" & leKey).要素表示 = 2 Then
                                                    Document.Elements("K" & leKey).内容文本 = "○" & Replace(mEleLimit(intLm).原始值域, ";", "  ○")
                                                Else
                                                    Document.Elements("K" & leKey).内容文本 = "□" & Replace(mEleLimit(intLm).原始值域, ";", "  □")
                                                End If
                                                Document.Elements("K" & leKey).Refresh Editor1
                                            Else
                                                Document.Elements("K" & leKey).要素值域 = mEleLimit(intLm).原始值域
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        Case 3 '删除要素
                            If (strElementValue = mEleLimit(intLm).原因内容 Or InStr(strElementValue, "●" & mEleLimit(intLm).原因内容) > 0 Or InStr(strElementValue, "■" & mEleLimit(intLm).原因内容) > 0) Then
                                For intEl = 1 To Document.Elements.Count
                                    leKey = Document.Elements(intEl).Key
                                    If Document.Elements("K" & leKey).要素名称 = mEleLimit(intLm).结果要素 And Document.Elements("K" & leKey).诊治要素ID = mEleLimit(intLm).结果要件id Then
                                        If Document.Elements("K" & leKey).输入形态 = 1 Then
                                            If InStr(Document.Elements("K" & leKey).内容文本, "●") = 0 And InStr(Document.Elements("K" & leKey).内容文本, "■") = 0 Then
                                                '结果要素没有选项被选中才删除，并刷新显示
                                                Document.Elements("K" & leKey).DeleteFromEditor Editor1
                                                Exit For
                                            End If
                                        Else
                                            If Document.Elements("K" & leKey).内容文本 = "" Then Document.Elements("K" & leKey).DeleteFromEditor Editor1: Exit For
                                        End If
                                    End If
                                Next
                            End If
                        Case 4 '同时变更
                            If strElementName = mEleLimit(intLm).原因要素 And lEleId = mEleLimit(intLm).原因要件id Then
                                For intEl = 1 To Document.Elements.Count
                                    leKey = Document.Elements(intEl).Key
                                    If Document.Elements("K" & leKey).要素名称 = mEleLimit(intLm).结果要素 And Document.Elements("K" & leKey).诊治要素ID = mEleLimit(intLm).结果要件id Then
                                        If Document.Elements("K" & leKey).输入形态 = 1 Then
                                            '结果要素没有选项被选中才更新值域，内容，并刷新显示
                                            Document.Elements("K" & leKey).内容文本 = strElementValue
                                            Document.Elements("K" & leKey).Refresh Editor1
                                        Else
                                            '结果要素没有选项被选中才更新值域，内容，并刷新显示
                                            Document.Elements("K" & leKey).内容文本 = strElementValue
                                            Document.Elements("K" & leKey).Refresh Editor1
                                        End If
                                    End If
                                Next
                            End If
                        Case 5
                            If (strElementValue = mEleLimit(intLm).原因内容 Or InStr(strElementValue, "●" & mEleLimit(intLm).原因内容) > 0 Or InStr(strElementValue, "■" & mEleLimit(intLm).原因内容) > 0) Then
                                If FindKey(Editor1, "E", lKey, lSS, lSE, lES, lEE, bNeeded) Then
                                    Editor1.ForceEdit = True
                                    Editor1.Range(lEE, lEE).Selected
                                    Editor1.Range(lEE, lEE).Font.Protected = False
                                    If Editor1.Range(lEE, lEE + 1).Text <> "，" Or Editor1.Range(lEE, lEE + 1).Text <> "。" _
                                       Or Editor1.Range(lEE, lEE + 1).Text <> "," Or Editor1.Range(lEE, lEE + 1).Text <> "." Then
                                        Editor1.Range(lEE + 1, lEE + 1).Selected
                                    Else
                                        Editor1.Range(lEE, lEE).Selected
                                    End If
                                    mfrmSentenceDetailed_RowDblClick mEleLimit(intLm).结果要件id
                                End If
                            End If
                    End Select
                End If
        End Select
    Next
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function MakeSentenceLimit(ByVal lngCompend As Long) As String
'功能：通过比较病种限制词句，要素限制词句，返回以豆号分隔的,以逗号开始，以逗号结束
Dim strReturn As String, intLm As Integer, intEl As Integer, leKey As Long
    On Error GoTo errHand
    If Not (Document.EPRFileInfo.种类 = cpr门诊病历 Or Document.EPRFileInfo.种类 = cpr住院病历) Then Exit Function '只针对门诊病历和住院病历
    If Document.EditType = cprET_单病历审核 Then Exit Function
    
    For intLm = 1 To UBound(mEleLimit)
        Select Case mEleLimit(intLm).变动原因
            Case 2, 3 '因病种引起词句变化
                If (mlDiseaseID = mEleLimit(intLm).原因要件id Or mlDiagnoseID = mEleLimit(intLm).原因要件id) And _
                    mEleLimit(intLm).变动结果 = 2 And mEleLimit(intLm).结果提纲id = lngCompend Then
                    strReturn = strReturn & "," & mEleLimit(intLm).结果要件id
                End If
            Case 1 '因要素引起词句变化,需要核对原因要素是否已有值、名称、ID、所处位置
                If mEleLimit(intLm).变动结果 = 2 And mEleLimit(intLm).结果提纲id = lngCompend Then
                    For intEl = 1 To Document.Elements.Count
                        leKey = Document.Elements(intEl).Key
                        If Document.Elements("K" & leKey).内容文本 <> "" Then
                            If Document.Elements("K" & leKey).要素名称 = mEleLimit(intLm).原因要素 And Document.Elements("K" & leKey).诊治要素ID = mEleLimit(intLm).原因要件id Then
                                If Document.Elements("K" & leKey).内容文本 = mEleLimit(intLm).原因内容 Or _
                                    InStr(Document.Elements("K" & leKey).内容文本, "●" & mEleLimit(intLm).原因内容) > 0 Or _
                                    InStr(Document.Elements("K" & leKey).内容文本, "■" & mEleLimit(intLm).原因内容) > 0 Then
                                        strReturn = strReturn & "," & mEleLimit(intLm).结果要件id
                                End If
                            End If
                        End If
                    Next
                End If
        End Select
    Next
    MakeSentenceLimit = Decode(strReturn, "", "", strReturn & ",")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## 功能：  设置表格填充色
'################################################################################################################
Private Sub ColorFillColor_pOK()
    SendKeys "{ESCAPE}"
    mlngCellFillColor = IIf(ColorFillColor.COLOR = tomAutoColor, -1, ColorFillColor.COLOR)
    If tblThis.Visible Then
        Dim lRow1 As Long, lRow2 As Long, lCol1 As Long, lCol2 As Long, i As Long, j As Long
        If tblThis.SelStartRow > tblThis.SelEndRow Then
            lRow1 = tblThis.SelEndRow
            lRow2 = tblThis.SelStartRow
        Else
            lRow1 = tblThis.SelStartRow
            lRow2 = tblThis.SelEndRow
        End If
        If tblThis.SelStartCol > tblThis.SelEndCol Then
            lCol1 = tblThis.SelEndCol
            lCol2 = tblThis.SelStartCol
        Else
            lCol1 = tblThis.SelStartCol
            lCol2 = tblThis.SelEndCol
        End If
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then Exit Sub
        For i = lRow1 To lRow2
            For j = lCol1 To lCol2
                tblThis.Cell(i, j).BackColor = mlngCellFillColor
            Next j
        Next i
        SetColorIcon "FILLCOLOR", ID_DRAW_FILLCOLOR, mlngCellFillColor
    End If
    tblThis.Modified = True
    tblThis.Refresh False, False
    SendKeys "{ESCAPE}"
End Sub

'################################################################################################################
'## 功能：  设置字体前景色
'################################################################################################################
Private Sub ColorForeColor_pOK()
    SendKeys "{ESCAPE}"
    mlngSelForeColor = ColorForeColor.COLOR
    If tblThis.Visible Then
        Dim lRow1 As Long, lRow2 As Long, lCol1 As Long, lCol2 As Long, i As Long, j As Long
        If tblThis.SelStartRow > tblThis.SelEndRow Then
            lRow1 = tblThis.SelEndRow
            lRow2 = tblThis.SelStartRow
        Else
            lRow1 = tblThis.SelStartRow
            lRow2 = tblThis.SelEndRow
        End If
        If tblThis.SelStartCol > tblThis.SelEndCol Then
            lCol1 = tblThis.SelEndCol
            lCol2 = tblThis.SelStartCol
        Else
            lCol1 = tblThis.SelStartCol
            lCol2 = tblThis.SelEndCol
        End If
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then Exit Sub
        For i = lRow1 To lRow2
            For j = lCol1 To lCol2
                tblThis.Cell(i, j).ForeColor = mlngSelForeColor
            Next j
        Next i
        SetColorIcon "FORECOLOR", ID_FORMAT_FORECOLOR, mlngSelForeColor
    End If
    tblThis.Modified = True
    tblThis.Refresh False, False
    SendKeys "{ESCAPE}"
End Sub

'################################################################################################################
'## 功能：  设置字体背景色
'################################################################################################################
Private Sub ColorHighlight_pOK()
    SendKeys "{ESCAPE}"
    mlngSelHightlightColor = ColorHighlight.COLOR
    If Editor1.Selection.Font.Protected = False And Editor1.Selection.Font.Hidden = False Then
        Editor1.Tag = "ColorHighlight_pOK"
        Editor1.ForceEdit = True
        Editor1.Selection.Font.BackColor = mlngSelHightlightColor
        Editor1.ForceEdit = False
        Editor1.Tag = ""
        SetColorIcon "HIGHLIGHT", ID_FORMAT_HIGHLIGHT, IIf(mlngSelHightlightColor = tomAutoColor, vbWhite, mlngSelHightlightColor)
    End If
    If tblThis.Visible Then
        tblThis.SetFocus
    Else
        If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
    End If
End Sub

'################################################################################################################
'## 功能：  设置页面背景色
'################################################################################################################
Private Sub ColorPaperBackColor_pOK()
    Editor1.PaperColor = IIf(ColorPaperBackColor.COLOR = tomAutoColor, vbWhite, ColorPaperBackColor.COLOR)
End Sub

'################################################################################################################
'## 功能：  回退操作
'################################################################################################################
Private Sub ExecBackSpace()
    If Me.Editor1.ReadOnly Then Exit Sub
    Dim i As Long, j As Long, lLen As Long
    Dim lS As Long, lE As Long, lSS As Long, lSS2 As Long
    Dim lf As Long, LL As Long, lR As Long, W As Long
    
    If Editor1.UIVisibled Then Editor1.CloseUIInterface

    Call AddUndoPoint  '手动缓存

    With Editor1
        .Tag = "ExecBackSpace"
        If .AuditMode Then
            i = .Selection.StartPos
            j = .Selection.StartPos + .SelLength

            '退格键处理
            If i = j Then
                If .Range(i - 1, i).Font.Protected Or .Range(i - 1, i).Font.Hidden Then Exit Sub
                If Me.Document.IsNewCharColor(.Range(i - 1, i).Font.ForeColor) And .Range(i - 1, i).Font.Strikethrough = False Then
                    '前面一个字符已经是新增文本，则直接删除之
                    .Range(i - 1, i).Text = ""
                Else
                    '否则，标记前面文本为删除文本
                    .Range(i - 1, i).Font.Strikethrough = True
                    .Range(i - 1, i).Font.ForeColor = Me.Document.GetDelCharColor(.Range(i - 1, i).Font.ForeColor)
                    .Range(i - 1, i - 1).Selected
                End If
            Else
                If Me.Document.IsNewCharColor(.Range(i, j).Font.ForeColor) And .Range(i, j).Font.Strikethrough = False Then
                    '选中文本为新增文本，直接删除之
                    .Range(i, j) = ""
                ElseIf Me.Document.IsNewCharColor(.Range(i, j).Font.ForeColor) = False And Me.Document.IsDelCharColor(.Range(i, j).Font.ForeColor) = False And .Range(i, j).Font.ForeColor <> tomUndefined Then
                    '否则，如果为普通文本，直接标记为删除
                    .Range(i, j).Font.Strikethrough = True
                    .Range(i, j).Font.ForeColor = Me.Document.GetDelCharColor(.Range(i, j).Font.ForeColor)
                ElseIf .Range(i, j).Font.ForeColor = tomUndefined Then
                    '否则，如果为混合文本，则不处理。
                    .Range(j, j).Selected
                End If
            End If
        Else
            '普通书写模式
            lS = .Selection.StartPos
            lE = .Selection.EndPos
            lSS = IIf(lS - 2 > 0, lS - 2, 0)
            lSS2 = IIf(lS - 16 > 0, lS - 16, 0)
            If .Range(lSS, lS) = vbCrLf Or lS = 0 Or (.Range(lSS2, lSS2 + 3) = "OE(" And .Range(lSS2, lSS2 + 3).Font.Hidden = True) Then
                '行首，增加首行缩进
                lf = .Range(lS, lE).Para.FirstLineIndent
                LL = .Range(lS, lE).Para.LeftIndent
                lR = .Range(lS, lE).Para.RightIndent
                If lf = tomUndefined Then lf = 0
                If LL = tomUndefined Then LL = 0
                If lR = tomUndefined Then lR = 0

                W = (.PaperWidth - .MarginLeft - .MarginRight - 3000) * .ZoomFactor / 20

                If lf > 0 Then
                    lf = 0
                Else
                    LL = LL - .DefaultTabStop
                End If
                If LL < 0 Then LL = 0
                .ForceEdit = True
                .Range(lS, lE).Para.SetIndents lf, LL, lR
                .ForceEdit = False
            ElseIf .Range(lE - 1, lS).Font.Protected = False Then
                .ForceEdit = True
                .Range(lE - 1, lS) = ""
                .ForceEdit = False
            End If
            If tblThis.Visible Then
                tblThis.SetFocus
            Else
                If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
            End If
        End If
        .Tag = ""
    End With
    Call ClearNoUseUndoList
End Sub

'################################################################################################################
'## 功能：  内容的粘贴操作（修正要素关键字，对于删除的要素要表现为新增，修订文本也统一改为新增文本）
'################################################################################################################
Private Sub ExecPaste(ByRef edtThis As Object)
    If edtThis.ReadOnly Then Exit Sub
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim i As Long, bForce As Boolean, bFinded As Boolean, strTmp As String, lS As Long, lE As Long, lngLen As Long
    Dim ParaFmt As New cParaFormat
    
    If tblThis.Visible Then
        If tblThis.SelectedCellKey > 0 Then
            If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected Then Exit Sub
            If tblThis.InEdit Then
                tblThis.InsertText Clipboard.GetText
            Else
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Clipboard.GetText
                tblThis.Refresh False, True, tblThis.SelectedCellKey
                tblThis_Resize tblThis.Width, tblThis.Height
            End If
            tblThis.Modified = True
        End If
        Exit Sub
    End If

    bBeteenKeys = IsBetweenAnyKeys(edtThis, edtThis.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then Exit Sub    '不允许粘贴到元素内部
'    Call AddUndoPoint  '手动缓存

    If edtThis.Selection.Font.ForeColor = tomUndefined Or edtThis.Selection.Font.Protected Then Exit Sub

    '如果剪贴板为空，那么就粘贴内部数据
    Dim strClipboard As String
    strClipboard = Clipboard.GetText
    If Len(Trim(strClipboard)) > 0 Then
        '粘贴剪贴板数据
        lS = edtThis.Selection.StartPos
        lE = lS + Len(strClipboard)
        edtThis.Tag = "ExecPaste"
        edtThis.ForceEdit = True
        edtThis.Range(lS, edtThis.Selection.EndPos).Text = strClipboard
        edtThis.Range(lS, lE).Font.Strikethrough = False
        edtThis.Range(lS, lE).Font.Protected = False
        edtThis.Range(lS, lE).Font.ForeColor = IIf(Me.Document.EditType = cprET_单病历审核, Me.Document.GetNewCharColor(vbBlack), tomAutoColor)
        edtThis.ForceEdit = False
        edtThis.Tag = ""
        edtThis.Range(lE, lE).Selected
        Exit Sub
    End If

    '先修正关键字
    gfrmPublic.edtPublic.ForceEdit = True
    '替换μ为u，防止崩溃
'    gfrmPublic.edtPublic.Text = Replace(gfrmPublic.edtPublic.Text, "μ", "u") '会导致编辑器内隐藏字段的属性丢失，暂时屏蔽，找到好的解决方法再行修改
    For i = 1 To gfrmPublic.Elements.Count
        '加入要素
        lKey = Me.Document.Elements.AddExistNode(gfrmPublic.Elements(i).Clone, False)
        Me.Document.Elements("K" & lKey).开始版 = Me.Document.目标版本
        Me.Document.Elements("K" & lKey).终止版 = 0     '去掉终止版
        Me.Document.Elements("K" & lKey).保留对象 = False
        Me.Document.Elements("K" & lKey).ID = 0
        '修正关键字
        bFinded = FindKey(gfrmPublic.edtPublic, "E", gfrmPublic.Elements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            strTmp = Format(lKey, "00000000") & "," & IIf(Me.Document.Elements("K" & lKey).保留对象, 1, 0) & ",0)"
            gfrmPublic.edtPublic.Range(lKSS, lKSE) = "ES(" & strTmp
            gfrmPublic.edtPublic.Range(lKES, lKEE) = "EE(" & strTmp
            gfrmPublic.Elements(i).Key = lKey '更新文本的同时，更新Key
        End If
    Next

    '拷贝RTF内容，清除前景色和删除线
    bForce = edtThis.ForceEdit
    edtThis.Tag = "ExecPaste"
    edtThis.Freeze
    edtThis.ForceEdit = True

    lS = 0: lE = Len(gfrmPublic.edtPublic.Text)
    If Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核 Then
        For i = lS To lE - 1
            '审阅模式下，全部处理为新增文本，去掉保护
            If gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And gfrmPublic.edtPublic.Range(i, i + 1).Font.Protected Then
                '保护文本视为新增文本
                gfrmPublic.edtPublic.Range(i, i + 1).Font.Protected = False
            End If
            gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = IIf(Me.Document.EditType = cprET_单病历审核, Me.Document.GetNewCharColor(vbBlack), tomAutoColor)
        Next
    Else
        For i = lS To lE - 1
            '审阅模式下，全部处理为新增文本，去掉保护
            If gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And gfrmPublic.edtPublic.Range(i, i + 1).Font.Protected Then
                '保护文本不变
            Else
                gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = tomAutoColor
            End If
        Next
    End If

    gfrmPublic.edtPublic.SelectAll
    gfrmPublic.edtPublic.Selection.Font.Strikethrough = False
    lS = edtThis.Selection.StartPos
    '1、保存段落格式
    Set ParaFmt = edtThis.Range(lS, lS).Para.GetParaFmt
    '2、同时保存Tab制表位
    Dim j As Long
    Dim iT As Single, lA As Long, lLd As Long, LL As Long
    Dim iTabPos() As Long, lAlign() As Byte, lLeader() As Long
    j = edtThis.Range(lS, lS).Para.TabCount

    If j = tomUndefined Then j = 0
    ReDim iTabPos(0 To j) As Long
    ReDim lAlign(0 To j) As Byte
    ReDim lLeader(0 To j) As Long
    For i = 0 To j - 1
        edtThis.TOM.TextDocument.Range(lS, lS).Para.GetTab i, iT, lA, LL
        iTabPos(i) = iT * 20
        lAlign(i) = lA * 20
        lLeader(i) = lLd * 20
    Next

    lngLen = Len(gfrmPublic.edtPublic.Text)
    If lngLen > 0 Then
'        gfrmPublic.edtPublic.CopyWithFormat
'        edtThis.PasteWithFormat
        edtThis.TOM.TextDocument.Selection.FormattedText = gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText
        '3、恢复段落格式
        edtThis.Range(lS, lS).Para.SetParaFmt ParaFmt
        '4、恢复Tab制表位
        For i = 0 To UBound(iTabPos)
            If iTabPos(i) > 0 Then edtThis.TOM.TextDocument.Range(lS, lS).Para.AddTab iTabPos(i) / 20, lAlign(i), tomSpaces
        Next i
        '去掉末尾的回车换行符
        If edtThis.Range(lS + lngLen, lS + lngLen + 2) = vbCrLf And edtThis.Range(lS + lngLen, lS + lngLen + 2).Font.Protected = False Then
            edtThis.Range(lS + lngLen, lS + lngLen + 2) = ""
        End If
        edtThis.Range(lS + lngLen, lS + lngLen).Selected
    End If
'    Clipboard.Clear

    edtThis.ForceEdit = bForce
    edtThis.UnFreeze
    edtThis.Tag = ""
    Call ClearNoUseUndoList
'    '更新提纲列表
'    Me.Document.Compends.UpdateOrdersFromText edtThis
'    Me.Document.Compends.FillTree mfrmCompends.Tree
End Sub

'################################################################################################################
'## 功能：  另存为示范词句
'################################################################################################################
Private Sub ExecSaveAsPhrase()
    Dim lngCompendID As Long, lngRetuId As Long, lngClassId As Long
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo errHand
    '获取到提纲id
    If mfrmCompends.Tree.SelectedItem Is Nothing Then Exit Sub
    If Me.Document.EditType = cprET_病历文件定义 Then
        lngCompendID = Me.Document.Compends(mfrmCompends.Tree.SelectedItem.Key).ID
    Else
        lngCompendID = Me.Document.Compends(mfrmCompends.Tree.SelectedItem.Key).定义提纲ID
    End If
    If lngCompendID = 0 Then MsgBox "临时提纲不能定义示范词句！", vbInformation, gstrSysName: Exit Sub
    
    '获取打词句分类id
    gstrSQL = "Select 词句分类id From 病历提纲词句 Where 提纲id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", lngCompendID)
    If rsTemp.RecordCount > 0 Then
        lngClassId = rsTemp.Fields(0).Value
    Else
        MsgBox "当前提纲没有设置词句示范分类对应，请联系管理员初始化基础数据！", vbInformation, gstrSysName
        Exit Sub
    End If
    '获取到当前用户是否有词句示范管理权限
    If InStr(1, gstrPrivsEpr, "全院病历词句") = 0 And InStr(1, gstrPrivsEpr, "科室病历词句") = 0 And InStr(1, gstrPrivsEpr, "个人病历词句") = 0 Then
        MsgBox "你不具备词句示范管理的权限！", vbInformation, gstrSysName
        Exit Sub
    End If
    '调用添加词句窗体
    lngRetuId = frmSentenceEdit.ShowMe(Me, True, 0, lngClassId, , True)
    If lngRetuId = 0 Then Exit Sub
    '刷新词句列表
    Call mfrmSentenceDetailed.zlSubRefList(lngRetuId)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'################################################################################################################
'## 功能：  内容的复制操作（包括文本和要素）
'################################################################################################################
Private Sub ExecCopy(ByRef edt As Object)

    If edt.ReadOnly Then Exit Sub
    
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, bFinded As Boolean, lngLen As Long, lngSum As Long

    If tblThis.Visible Then
        If tblThis.Row > 0 And tblThis.Col > 0 Then
            Clipboard.Clear
            Clipboard.SetText tblThis.Cell(tblThis.Row, tblThis.Col).Text
        End If
        Exit Sub
    End If

    '扩展起始位置和终止位置，使得其包含完整的要素定义
    lS = edt.Selection.StartPos
    lE = edt.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(edt, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(edt, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    
    '先拷贝RTF内容
    gfrmPublic.edtPublic.NewDoc
    gfrmPublic.edtPublic.ForceEdit = True
'    edt.Range(lS, lE).Selected
'    edt.CopyWithFormat
'    gfrmPublic.edtPublic.PasteWithFormat
    gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText = edt.TOM.TextDocument.Range(lS, lE).FormattedText
    '拷贝要素，过滤其他元素（图片、诊断、表格等），关键字也要拷贝过去，保证与内容的隐藏关键字Key值一致！
    Set gfrmPublic.Elements = New cEPRElements
    lngSum = 0
    For i = lS To lE
        bFinded = FindNextAnyKey(edt, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                '范围内存在关键字
                If sKeyType = "E" Then
                    '如果是要素，那么拷贝到缓冲区
                    gfrmPublic.Elements.AddExistNode Me.Document.Elements("K" & lKey).Clone(True), True
                Else
                    '如果是其他元素，则清除之（在gfrmPublic.edtPublic中清除，并记录当前位置）！
                    gfrmPublic.edtPublic.Range(lKSS - lS - lngSum, lKEE - lS - lngSum) = ""
                    lngSum = lngSum + lKEE - lKSS   '记录删除内容的总长度
                End If
            Else
                '否则，超出范围，退出循环
                Exit For
            End If
            i = lKEE - 1
        Else
            '不存在任何元素，那么退出循环
            Exit For
        End If
    Next
    Clipboard.Clear
End Sub

'################################################################################################################
'## 功能：  内容的剪切操作（包括文本和要素）
'################################################################################################################
Private Sub ExecCut()
    If Me.Editor1.ReadOnly Then Exit Sub
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, bFinded As Boolean, lngNum As Long, lngSum As Long

    If tblThis.Visible Then
        If tblThis.SelectedCellKey > 0 Then
            If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected Then Exit Sub
            Clipboard.Clear
            Clipboard.SetText tblThis.Cells("K" & tblThis.SelectedCellKey).Text
            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = ""
            tblThis.Modified = True
            tblThis.Refresh False, True, tblThis.SelectedCellKey
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
        Exit Sub
    End If

    Call AddUndoPoint  '手动缓存

    '扩展起始位置和终止位置，使得其包含完整的要素定义
    lS = Editor1.Selection.StartPos
    lE = Editor1.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(Editor1, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(Editor1, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    '末尾位置还不能跨越提纲
    bFinded = FindNextKey(Editor1, lS + 1, "O", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        If lKSS < lE Then
            If Editor1.Range(lKSS - 2, lKSS) = vbCrLf Then
                lE = lKSS - 2
            Else
                lE = lKSS
            End If
        End If
    End If
    If Editor1.Range(lE - 2, lE) = vbCrLf Then lE = lE - 2

    '先拷贝RTF内容
    gfrmPublic.edtPublic.NewDoc
    gfrmPublic.edtPublic.ForceEdit = True
'    Me.Editor1.Range(lS, lE).Selected
'    Me.Editor1.CopyWithFormat
'    gfrmPublic.edtPublic.PasteWithFormat
    gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText = Me.Editor1.TOM.TextDocument.Range(lS, lE).FormattedText
    '拷贝要素，过滤其他元素（图片、诊断、表格等），关键字也要拷贝过去，保证与内容的隐藏关键字Key值一致！
    Set gfrmPublic.Elements = New cEPRElements
    lngSum = 0
    For i = lS To lE
        bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                '范围内存在关键字
                If sKeyType = "E" Then
                    '如果是要素，那么拷贝到缓冲区
                    gfrmPublic.Elements.AddExistNode Me.Document.Elements("K" & lKey), True
                Else
                    '如果是其他元素，则清除之（在gfrmPublic.edtPublic中清除，并记录当前位置）！
                    gfrmPublic.edtPublic.Range(lKSS - lS - lngSum, lKEE - lS - lngSum) = ""
                    lngSum = lngSum + lKEE - lKSS   '记录删除内容的总长度
                End If
            Else
                '否则，超出范围，退出循环
                Exit For
            End If
            i = lKEE - 1
        Else
            '不存在任何元素，那么退出循环
            Exit For
        End If
    Next

    '删除选中内容
    Dim bForce As Boolean, COLOR As OLE_COLOR, bProtect1 As Boolean, bProtect2 As Boolean
    bForce = Me.Editor1.ForceEdit
    Me.Editor1.Freeze
    Me.Editor1.Tag = "ExecCut"
    Me.Editor1.ForceEdit = True
    If Me.Editor1.AuditMode Then
        '审核模式的话，需要进行颜色和版本特殊处理！
        '处理元素
        For i = lS To lE
            bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bFinded Then
                If lKSS < lE Then
                    '范围内存在关键字
                    Select Case sKeyType
                    Case "E"    '要素
                        If Me.Document.Elements("K" & lKey).保留对象 = False Then
                            If Me.Document.Elements("K" & lKey).开始版 = Me.Document.目标版本 Then
                                '本次版本新建的要素
                                Me.Editor1.Range(lKSS, lKEE) = ""
                                lE = lE - (lKEE - lKSS)
                                i = lKSS - 1
                            ElseIf Me.Document.Elements("K" & lKey).开始版 < Me.Document.目标版本 And _
                                Me.Document.Elements("K" & lKey).终止版 = 0 Then
                                Me.Document.Elements("K" & lKey).终止版 = Me.Document.目标版本 - 1
                                Me.Document.Elements("K" & lKey).Refresh Me.Editor1
                                i = lKEE - 1
                            Else
                                i = lKEE - 1
                            End If
                        Else
                            i = lKEE - 1
                        End If
                    Case "D"    '诊断
                        If Me.Document.Diagnosises("K" & lKey).开始版 = Me.Document.目标版本 Then
                            '本次版本新建的要素
                            Me.Editor1.Range(lKSS, lKEE) = ""
                            lE = lE - (lKEE - lKSS)
                            i = lKSS - 1
                        ElseIf Me.Document.Diagnosises("K" & lKey).开始版 < Me.Document.目标版本 And _
                            Me.Document.Diagnosises("K" & lKey).终止版 = 0 Then
                            Me.Document.Diagnosises("K" & lKey).终止版 = Me.Document.目标版本 - 1
                            Me.Document.Diagnosises("K" & lKey).Refresh Me.Editor1
                            i = lKEE - 1
                        Else
                            i = lKEE - 1
                        End If
                    Case Else
                       '如果是其他元素，则不处理
                       i = lKEE - 1
                    End Select
                Else
                    '否则，超出范围，退出循环
                    Exit For
                End If
            Else
                '不存在任何元素，那么退出循环
                Exit For
            End If
        Next

        '处理文字
        For i = lS To lE - 1
            If Editor1.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And Editor1.Range(i, i + 1).Font.Protected Then
                '保护文本，不允许删除（不处理）
            ElseIf Editor1.Range(i, i + 1).Font.Protected = False Then
                COLOR = IIf(Editor1.Range(i, i + 1).Font.ForeColor = tomAutoColor Or Editor1.Range(i, i + 1).Font.ForeColor = tomUndefined, vbBlack, Editor1.Range(i, i + 1).Font.ForeColor)
                If Me.Document.IsNewCharColor(COLOR) And Editor1.Range(i, i + 1).Font.Strikethrough = False Then
                    '后面一个字符是新增文本，则直接删除之
                    Editor1.Range(i, i + 1).Text = ""
                    lE = lE - 1
                    i = i - 1
                ElseIf rgbBlue(COLOR) <> 0 And Me.Document.IsDelCharColor(COLOR) = False Then
                    '否则，如果前面文本为以前版本的删除文本，则不作任何处理！
                Else
                    '否则标记为删除文本
                    Editor1.Range(i, i + 1).Font.Strikethrough = True
                    Editor1.Range(i, i + 1).Font.ForeColor = Me.Document.GetDelCharColor(Editor1.Range(i, i + 1).Font.ForeColor)
                End If
            End If
        Next
        Me.Editor1.UnFreeze
        Me.Editor1.Range(lS, lE).Selected
    Else
        '非修订模式，则清除所有要素、图片、表格、诊断，不能删除提纲
        lngSum = 0
        For i = lS To lE - 1
            bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bFinded Then
                If lKSS < lE Then   '范围内存在关键字
                    '1、先处理前面的文字
                    lngNum = DelTextRange(Me.Editor1, i, lKSS)
                    lE = lE - lngNum
                    lngSum = lngSum + lngNum
                    i = lKSS - lngNum - 1
                    '2、处理后面一个要素、图片、表格、诊断
                    Select Case sKeyType
                    Case "E"    '要素
                        If Me.Document.Elements("K" & lKey).保留对象 = False Then
                            Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                            Me.Document.Elements.Remove "K" & lKey
                            lngSum = lngSum + (lKEE - lKSS)
                            lE = lE - (lKEE - lKSS)
                        Else
                            i = lKEE - lngNum - 1
                        End If
                    Case "P"    '图片
                        Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                        Me.Document.Pictures.Remove "K" & lKey
                        lngSum = lngSum + (lKEE - lKSS)
                        lE = lE - (lKEE - lKSS)
                    Case "T"    '表格
                        Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                        Me.Document.Tables.Remove "K" & lKey
                        lngSum = lngSum + (lKEE - lKSS)
                        lE = lE - (lKEE - lKSS)
                    Case "D"    '诊断
                        Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                        Me.Document.Diagnosises.Remove "K" & lKey
                        lngSum = lngSum + (lKEE - lKSS)
                        lE = lE - (lKEE - lKSS)
                    Case Else
                       '如果是其他元素，则不处理
                       i = lKEE - lngNum - 1
                    End Select
                Else
                    '否则，超出范围，退出循环
                    Exit For
                End If
            Else
                '不存在任何元素，那么退出循环
                Exit For
            End If
        Next
        If i < lE Then
            lngNum = DelTextRange(Me.Editor1, i, lE)
        End If
        Me.Editor1.UnFreeze
        Me.Editor1.SelLength = 0
        Me.Editor1.Range(lS, lS).Selected
    End If
    Me.Editor1.Tag = ""
    Me.Editor1.ForceEdit = bForce
    Clipboard.Clear
    Call ClearNoUseUndoList
End Sub
'################################################################################################################
'## 功能：  修订内容的删除操作
'################################################################################################################
Private Sub ExecAuditDelete()
    If Me.Editor1.ReadOnly Then Exit Sub
    Dim i As Long, j As Long, blnForce As Boolean
    Dim lngStart As Long, lngEnd As Long, lngLen As Long, blnWithEles As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim strQuestion As String, COLOR As OLE_COLOR, lS As Long, lE As Long, bFinded As Boolean
    Dim bProtect1 As Boolean, bProtect2 As Boolean, lStart As Long
    Dim lngNum As Long, lngSum As Long, lIndex As Long
     '选中内容非空
        '扩展起始位置和终止位置，使得其包含完整的要素定义
        lS = Editor1.Selection.StartPos
        lE = Editor1.Selection.EndPos
        '如果选中位置在文章末尾，则不处理
        If lS = Len(Editor1.Text) Then Exit Sub
        bBeteenKeys = IsBetweenAnyKeys(Editor1, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lS = lKSS
        bBeteenKeys = IsBetweenAnyKeys(Editor1, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lE = lKEE
        '末尾位置还不能跨越提纲
        bFinded = FindNextKey(Editor1, lS + 1, "O", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                If Editor1.Range(lKSS - 2, lKSS) = vbCrLf Then
                    lE = lKSS - 2
                Else
                    lE = lKSS
                End If
            End If
        End If
        If Editor1.Range(lE - 2, lE) = vbCrLf Then lE = lE - 2
    If Me.Editor1.AuditMode Then
            '审核模式的话，需要进行颜色和版本特殊处理！
            '处理元素
            For i = lS To lE
                bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bFinded Then
                    If lKSS < lE Then
                        '范围内存在关键字
                        Select Case sKeyType
                        Case "E"    '要素
                            If Me.Document.Elements("K" & lKey).保留对象 = False Then
                                If Me.Document.Elements("K" & lKey).开始版 = Me.Document.目标版本 Then
                                    '本次版本新建的要素
                                    Me.Editor1.Range(lKSS, lKEE) = ""
                                    lE = lE - (lKEE - lKSS)
                                    i = lKSS - 1
                                ElseIf Me.Document.Elements("K" & lKey).开始版 < Me.Document.目标版本 And _
                                    Me.Document.Elements("K" & lKey).终止版 = 0 Then
                                    Me.Document.Elements("K" & lKey).终止版 = Me.Document.目标版本 - 1
                                    Me.Document.Elements("K" & lKey).Refresh Me.Editor1
                                    i = lKEE - 1
                                Else
                                    i = lKEE - 1
                                End If
                            Else
                                i = lKEE - 1
                            End If
                        Case "D"    '诊断
                            If Me.Document.Diagnosises("K" & lKey).开始版 = Me.Document.目标版本 Then
                                '本次版本新建的要素
                                Me.Editor1.Range(lKSS, lKEE) = ""
                                lE = lE - (lKEE - lKSS)
                                i = lKSS - 1
                            ElseIf Me.Document.Diagnosises("K" & lKey).开始版 < Me.Document.目标版本 And _
                                Me.Document.Diagnosises("K" & lKey).终止版 = 0 Then
                                Me.Document.Diagnosises("K" & lKey).终止版 = Me.Document.目标版本 - 1
                                Me.Document.Diagnosises("K" & lKey).Refresh Me.Editor1
                                i = lKEE - 1
                            Else
                                i = lKEE - 1
                            End If
                        Case Else
                           '如果是其他元素，则不处理
                           i = lKEE - 1
                        End Select
                    Else
                        '否则，超出范围，退出循环
                        Exit For
                    End If
                Else
                    '不存在任何元素，那么退出循环
                    Exit For
                End If
            Next

            '处理文字
            For i = lS To lE - 1
                If Editor1.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And Editor1.Range(i, i + 1).Font.Protected Then
                    '保护文本，不允许删除（不处理）
                ElseIf Editor1.Range(i, i + 1).Font.Protected = False Then
                    COLOR = IIf(Editor1.Range(i, i + 1).Font.ForeColor = tomAutoColor Or Editor1.Range(i, i + 1).Font.ForeColor = tomUndefined, vbBlack, Editor1.Range(i, i + 1).Font.ForeColor)
                    If Me.Document.IsNewCharColor(COLOR) And Editor1.Range(i, i + 1).Font.Strikethrough = False Then
                        '后面一个字符是新增文本，则直接删除之
                        Editor1.Range(i, i + 1).Text = ""
                        lE = lE - 1
                        i = i - 1
                    ElseIf rgbBlue(COLOR) <> 0 And Me.Document.IsDelCharColor(COLOR) = False Then
                        '否则，如果前面文本为以前版本的删除文本，则不作任何处理！
                    ElseIf Editor1.Range(i, i + 2).Text = vbCrLf Then
                        '是回车换行符且不处于保护状态，则直接删除
                        If Not (Editor1.Range(i, i + 2).Font.Protected Or Editor1.Range(i, i + 2).Font.Hidden) Then
                            Editor1.Range(i, i + 2).Text = ""
                            If lE > lS Then lE = lE - 2
                            i = i - 1
                        End If
                    Else
                        '否则标记为删除文本
                        Editor1.Range(i, i + 1).Font.Strikethrough = True
                        Editor1.Range(i, i + 1).Font.ForeColor = Me.Document.GetDelCharColor(Editor1.Range(i, i + 1).Font.ForeColor)
                    End If
                End If
            Next
            Me.Editor1.UnFreeze
            Me.Editor1.Range(lE, lE).Selected
      End If
End Sub
'################################################################################################################
'## 功能：  内容的删除操作
'################################################################################################################
Private Sub ExecDelete()
    If Me.Editor1.ReadOnly Then Exit Sub
    Dim i As Long, j As Long, blnForce As Boolean
    Dim lngStart As Long, lngEnd As Long, lngLen As Long, blnWithEles As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim strQuestion As String, COLOR As OLE_COLOR, lS As Long, lE As Long, bFinded As Boolean
    Dim bProtect1 As Boolean, bProtect2 As Boolean, lStart As Long
    Dim lngNum As Long, lngSum As Long, lIndex As Long
    
    If tblThis.Visible Then
        '表格内部的删除
        lKey = tblThis.SelectedCellKey
        If lKey <= 0 Then Exit Sub
        If tblThis.Cells("K" & lKey).Protected = True And Me.Document.EditType <> cprET_病历文件定义 Then Exit Sub
        If tblThis.Cells("K" & lKey).Text = "" And tblThis.Cells("K" & lKey).Protected Then
            '删除要素/图片
            If Val(tblThis.Tag) > 0 Then
                If Val(tblThis.Cells("K" & lKey).Tag) > 0 Then
                    If tblThis.Cells("K" & lKey).Picture Is Nothing Then
                        Me.Document.Tables("K" & tblThis.Tag).Elements.Remove "K" & tblThis.Cells("K" & lKey).Tag
                    Else
                        Me.Document.Tables("K" & tblThis.Tag).Pictures.Remove "K" & tblThis.Cells("K" & lKey).Tag
                        Set tblThis.Cells("K" & lKey).Picture = Nothing
                    End If
                End If
            End If
            tblThis.Cells("K" & lKey).Tag = ""
            tblThis.Cells("K" & lKey).ToolTipText = ""
            tblThis.Cells("K" & lKey).Protected = False
            tblThis.Refresh False, True, lKey
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            '删除文本
            If tblThis.InEdit Then
                '如果在编辑状态
                tblThis.PressDelKey
                tblThis.Refresh False, True, lKey
                tblThis_Resize tblThis.Width, tblThis.Height
            Else
                tblThis.Cells("K" & lKey).Text = ""
                If Val(tblThis.Tag) > 0 Then
                    If Val(tblThis.Cells("K" & lKey).Tag) > 0 Then
                        If tblThis.Cells("K" & lKey).Picture Is Nothing Then
                            Me.Document.Tables("K" & tblThis.Tag).Elements("K" & tblThis.Cells("K" & lKey).Tag).内容文本 = ""
                        End If
                    End If
                End If
                tblThis.Refresh False, True, lKey
                tblThis_Resize tblThis.Width, tblThis.Height
            End If
        End If
        tblThis.Modified = True
        Exit Sub
    End If
    
    If Editor1.UIVisibled Then Editor1.CloseUIInterface
    Call AddUndoPoint  '手动缓存

    blnForce = Editor1.ForceEdit
    Editor1.Tag = "ExecDelete"
    Editor1.ForceEdit = True

    If Me.Editor1.SelLength > 0 Then
        '选中内容非空
        '扩展起始位置和终止位置，使得其包含完整的要素定义
        lS = Editor1.Selection.StartPos
        lE = Editor1.Selection.EndPos
        '如果选中位置在文章末尾，则不处理
        If lS = Len(Editor1.Text) Then Exit Sub
        bBeteenKeys = IsBetweenAnyKeys(Editor1, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lS = lKSS
        bBeteenKeys = IsBetweenAnyKeys(Editor1, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lE = lKEE
        '末尾位置还不能跨越提纲
        bFinded = FindNextKey(Editor1, lS + 1, "O", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                If Editor1.Range(lKSS - 2, lKSS) = vbCrLf Then
                    lE = lKSS - 2
                Else
                    lE = lKSS
                End If
            End If
        End If
        If Editor1.Range(lE - 2, lE) = vbCrLf Then lE = lE - 2

        '删除选中内容
        Me.Editor1.Freeze
        If Me.Editor1.AuditMode Then
            '审核模式的话，需要进行颜色和版本特殊处理！
            '处理元素
            For i = lS To lE
                bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bFinded Then
                    If lKSS < lE Then
                        '范围内存在关键字
                        Select Case sKeyType
                        Case "E"    '要素
                            If Me.Document.Elements("K" & lKey).保留对象 = False Then
                                If Me.Document.Elements("K" & lKey).开始版 = Me.Document.目标版本 Then
                                    '本次版本新建的要素
                                    Me.Editor1.Range(lKSS, lKEE) = ""
                                    lE = lE - (lKEE - lKSS)
                                    i = lKSS - 1
                                ElseIf Me.Document.Elements("K" & lKey).开始版 < Me.Document.目标版本 And _
                                    Me.Document.Elements("K" & lKey).终止版 = 0 Then
                                    Me.Document.Elements("K" & lKey).终止版 = Me.Document.目标版本 - 1
                                    Me.Document.Elements("K" & lKey).Refresh Me.Editor1
                                    i = lKEE - 1
                                Else
                                    i = lKEE - 1
                                End If
                            Else
                                i = lKEE - 1
                            End If
                        Case "D"    '诊断
                            If Me.Document.Diagnosises("K" & lKey).开始版 = Me.Document.目标版本 Then
                                '本次版本新建的要素
                                Me.Editor1.Range(lKSS, lKEE) = ""
                                lE = lE - (lKEE - lKSS)
                                i = lKSS - 1
                            ElseIf Me.Document.Diagnosises("K" & lKey).开始版 < Me.Document.目标版本 And _
                                Me.Document.Diagnosises("K" & lKey).终止版 = 0 Then
                                Me.Document.Diagnosises("K" & lKey).终止版 = Me.Document.目标版本 - 1
                                Me.Document.Diagnosises("K" & lKey).Refresh Me.Editor1
                                i = lKEE - 1
                            Else
                                i = lKEE - 1
                            End If
                        Case Else
                           '如果是其他元素，则不处理
                           i = lKEE - 1
                        End Select
                    Else
                        '否则，超出范围，退出循环
                        Exit For
                    End If
                Else
                    '不存在任何元素，那么退出循环
                    Exit For
                End If
            Next

            '处理文字
            For i = lS To lE - 1
                If Editor1.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And Editor1.Range(i, i + 1).Font.Protected Then
                    '保护文本，不允许删除（不处理）
                ElseIf Editor1.Range(i, i + 1).Font.Protected = False Then
                    COLOR = IIf(Editor1.Range(i, i + 1).Font.ForeColor = tomAutoColor Or Editor1.Range(i, i + 1).Font.ForeColor = tomUndefined, vbBlack, Editor1.Range(i, i + 1).Font.ForeColor)
                    If Me.Document.IsNewCharColor(COLOR) And Editor1.Range(i, i + 1).Font.Strikethrough = False Then
                        '后面一个字符是新增文本，则直接删除之
                        Editor1.Range(i, i + 1).Text = ""
                        lE = lE - 1
                        i = i - 1
                    ElseIf rgbBlue(COLOR) <> 0 And Me.Document.IsDelCharColor(COLOR) = False Then
                        '否则，如果前面文本为以前版本的删除文本，则不作任何处理！
                    ElseIf Editor1.Range(i, i + 2).Text = vbCrLf Then
                        '是回车换行符且不处于保护状态，则直接删除
                        If Not (Editor1.Range(i, i + 2).Font.Protected Or Editor1.Range(i, i + 2).Font.Hidden) Then
                            Editor1.Range(i, i + 2).Text = ""
                            If lE > lS Then lE = lE - 2
                            i = i - 1
                        End If
                    Else
                        '否则标记为删除文本
                        Editor1.Range(i, i + 1).Font.Strikethrough = True
                        Editor1.Range(i, i + 1).Font.ForeColor = Me.Document.GetDelCharColor(Editor1.Range(i, i + 1).Font.ForeColor)
                    End If
                End If
            Next
            Me.Editor1.UnFreeze
            Me.Editor1.Range(lE, lE).Selected
        Else
            '非修订模式，则清除所有要素、图片、表格、诊断，不能删除提纲
            lngSum = 0
            For i = lS To lE - 1
                bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bFinded Then
                    If lKSS < lE Then   '范围内存在关键字
                        '1、先处理前面的文字
                        lngNum = DelTextRange(Me.Editor1, i, lKSS)
                        lE = lE - lngNum
                        lngSum = lngSum + lngNum
                        i = lKSS - lngNum - 1
                        '2、处理后面一个要素、图片、表格、诊断
                        Select Case sKeyType
                        Case "E"    '要素
                            If Me.Document.Elements("K" & lKey).保留对象 = False Then
                                Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                                Me.Document.Elements.Remove "K" & lKey
                                lngSum = lngSum + (lKEE - lKSS)
                                lE = lE - (lKEE - lKSS)
                            Else
                                i = lKEE - lngNum - 1
                            End If
                        Case "P"    '图片
                            Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                            Me.Document.Pictures.Remove "K" & lKey
                            lngSum = lngSum + (lKEE - lKSS)
                            lE = lE - (lKEE - lKSS)
                        Case "T"    '表格
                            Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                            Me.Document.Tables.Remove "K" & lKey
                            lngSum = lngSum + (lKEE - lKSS)
                            lE = lE - (lKEE - lKSS)
                        Case "D"    '诊断
                            Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                            Me.Document.Diagnosises.Remove "K" & lKey
                            lngSum = lngSum + (lKEE - lKSS)
                            lE = lE - (lKEE - lKSS)
                        Case Else
                           '如果是其他元素，则不处理
                           i = lKEE - lngNum - 1
                        End Select
                    Else
                        '否则，超出范围，退出循环
                        Exit For
                    End If
                Else
                    '不存在任何元素，那么退出循环
                    Exit For
                End If
            Next
            If i < lE Then
                lngNum = DelTextRange(Me.Editor1, i, lE)
            End If
            Me.Editor1.UnFreeze
            Me.Editor1.SelLength = 0
            Me.Editor1.Range(lE - lngNum, lE - lngNum).Selected
        End If
        Clipboard.Clear
    Else
        '没有选择文本
        lS = Editor1.Selection.StartPos
        lE = Editor1.Selection.EndPos
        '如果选中位置在文章末尾，则不处理
        If lS = Len(Editor1.Text) Then Exit Sub
        If Editor1.AuditMode Then
            '审核模式
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys Then
                '删除单个诊治要素、图片或者表格
                Select Case sKeyType
                Case "E"
                    If Document.Elements("K" & lKey).终止版 > 0 Then Exit Sub
                    Editor1.Range(lKSE, lKES).Selected
                    strQuestion = "是否删除该诊治要素？"
                Case "D"
                    If Document.Diagnosises("K" & lKey).终止版 > 0 Then Exit Sub
                    Editor1.Range(lKSE, lKES).Selected
                    strQuestion = "是否删除该诊断？"
                Case Else
                    GoTo LL
                End Select
    '            If MsgBox(strQuestion, vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
                Select Case sKeyType
                Case "E"
                    If Document.Elements("K" & lKey).保留对象 = False Then
                        If Document.Elements("K" & lKey).开始版 = Me.Document.目标版本 Then
                            Document.Elements("K" & lKey).DeleteFromEditor Me.Editor1
                            Document.Elements.Remove "K" & lKey
                            Me.Editor1.SelLength = 0
                        Else
                            Document.Elements("K" & lKey).终止版 = Me.Document.目标版本 - 1
                            Document.Elements("K" & lKey).Refresh Me.Editor1
                            Me.Editor1.Range(lKEE, lKEE).Selected
                        End If
                    End If
                Case "D"
                    If Document.Diagnosises("K" & lKey).开始版 = Me.Document.目标版本 Then
                        Document.Diagnosises("K" & lKey).DeleteFromEditor Me.Editor1
                        Document.Diagnosises.Remove "K" & lKey
                        Me.Editor1.SelLength = 0
                    Else
                        Document.Diagnosises("K" & lKey).终止版 = Me.Document.目标版本 - 1
                        Document.Diagnosises("K" & lKey).Refresh Me.Editor1
                        Me.Editor1.Range(lKEE, lKEE).Selected
                    End If
                Case Else
                    GoTo LL
                End Select
    '            End If
            Else
                '文本的编辑
                With Editor1
                    i = .Selection.StartPos
                    j = .Selection.StartPos + .SelLength
                    If .Range(i, j).Font.Protected Or .Range(i, j).Font.Hidden Then GoTo LL
                    If .Range(i, i + 2) = vbCrLf Then
'                        COLOR = IIf(.Range(i, i + 2).Font.ForeColor = tomAutoColor Or .Range(i, i + 2).Font.ForeColor = tomUndefined, vbBlack, .Range(i, i + 2).Font.ForeColor)
                        If .Range(i, i + 2).Font.Protected And .Range(i, i + 2).Font.Hidden Then GoTo LL
'                        If Me.Document.IsNewCharColor(COLOR) And .Range(i, i + 2).Font.Strikethrough = False Then
                            '后面一个字符是新增文本，则直接删除之
                            .Range(i, i + 2).Text = ""
'                        ElseIf rgbBlue(COLOR) <> 0 And Me.Document.IsDelCharColor(COLOR) = False Then
'                            '否则，如果前面文本为以前版本的删除文本，则不作任何处理！
'                            .Range(i + 2, i + 2).Selected
'                        Else
'                            '否则标记为删除文本
'                            .Range(i, i + 2).Font.Strikethrough = True
'                            .Range(i, i + 2).Font.ForeColor = Me.Document.GetDelCharColor(.Range(i, i + 2).Font.ForeColor)
'                            .Range(i + 2, i + 2).Selected   '后移一位
'                        End If
                    Else
                        COLOR = IIf(.Range(i, i + 1).Font.ForeColor = tomAutoColor, vbBlack, .Range(i, i + 1).Font.ForeColor)
                        If .Range(i, i + 1).Font.Protected And .Range(i, i + 1).Font.Hidden Then GoTo LL
                        If Me.Document.IsNewCharColor(COLOR) And .Range(i, i + 1).Font.Strikethrough = False Then
                            '后面一个字符是新增文本，则直接删除之
                            .Range(i, i + 1).Text = ""
                        ElseIf rgbBlue(COLOR) <> 0 And Me.Document.IsDelCharColor(COLOR) = False Then
                            '否则，如果前面文本为以前版本的删除文本，则不作任何处理！
                            .Range(i + 1, i + 1).Selected
                        Else
                            '否则标记为删除文本
                            .Range(i, i + 1).Font.Strikethrough = True
                            .Range(i, i + 1).Font.ForeColor = Me.Document.GetDelCharColor(.Range(i, i + 1).Font.ForeColor)
                            .Range(i + 1, i + 1).Selected   '后移一位
                        End If
                    End If
                End With
            End If
        Else
            '书写模式：内容为纯文本
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys Then
                '删除单个诊治要素、图片或者表格
                Select Case sKeyType
                Case "T"
                    If Document.Tables("K" & lKey).预制提纲ID <> 0 Then
                        MsgBox "不能删除预制提纲中固有元素！除非删除该预制提纲本身。", vbOKOnly + vbInformation, gstrSysName
                        GoTo LL
                    Else
                        strQuestion = "是否删除该表格？"
                    End If
                Case "P"
                    strQuestion = "是否删除该图片？"
                Case "E"
                    Editor1.Range(lKSE, lKES).Selected
                    strQuestion = "是否删除该诊治要素？"
                Case "D"
                    Editor1.Range(lKSE, lKES).Selected
                    strQuestion = "是否删除该诊断？"
                Case Else
                    GoTo LL
                End Select
    '            If MsgBox(strQuestion, vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
                Select Case sKeyType
                Case "T"
                    Document.Tables.Remove "K" & lKey
                Case "P"
                    Document.Pictures.Remove "K" & lKey
                Case "E"
                    If Document.Elements("K" & lKey).保留对象 = False Or Me.Document.EditType = cprET_病历文件定义 Then
                        Document.Elements.Remove "K" & lKey
                    Else
                        GoTo LL
                    End If
                Case "D"
                    Document.Diagnosises.Remove "K" & lKey
                Case Else
                    GoTo LL
                End Select
                Editor1.Range(lKSS, lKEE) = ""
                If Editor1.Range(lKSS - 2, lKSS) = vbCrLf And Editor1.Range(lKSS - 2, lKSS).Font.Protected Then
                    Editor1.Range(lKSS - 2, lKSS) = ""
                    Editor1.Range(lKSS - 2, lKSS - 2).Font.Protected = False
                Else
                    Editor1.Range(lKSS, lKSS).Font.Protected = False
                End If
    '            End If
            Else
                '删除文本
                i = Editor1.Selection.StartPos
                j = Len(Editor1.Text)

                If Editor1.Range(i, i + 2).Font.Protected = False And Editor1.Range(i, i + 2) = vbCrLf And _
                    Editor1.Range(i + 2, i + 5) = "OS(" And Editor1.Range(i + 2, i + 5).Font.Hidden Then
                    '首先不允许删除提纲前面的回车，保持提纲在每段落开始位置！
                    Editor1.Range(i + 2, i + 2).Selected
                ElseIf Editor1.Range(i, i + 1).Font.Protected = False And (Editor1.Range(i + 1, i + 2).Font.Protected = True Or i = j - 1) Then
                    Editor1.Range(i, i + 1) = ""
                ElseIf Editor1.Range(i - 1, i).Font.Protected = True And Editor1.Range(i, i + 1).Font.Protected = False Then
                    If Editor1.Range(i, i + 2) = vbCrLf And Editor1.Range(i, i + 2).Font.Protected = False Then
                        Editor1.Range(i, i + 2) = ""
                        Editor1.Range(i, i).Font.Protected = False
                    Else
                        Editor1.Delete
                    End If
                ElseIf Editor1.Range(i, i + 2) = vbCrLf And Editor1.Range(i, i + 2).Font.Protected = False Then
                    Editor1.Range(i, i + 2) = ""
                    Editor1.Range(i, i).Font.Protected = False
                ElseIf Editor1.Range(i, i).Font.Protected = False And Editor1.Range(i, i + 1).Font.Protected = False Then
                    Editor1.Delete
                ElseIf Editor1.Range(i, i + 2) = vbCrLf And Editor1.Range(i, i + 2).Font.Protected Then
                    Editor1.Range(i + 2, i + 2).Selected
                Else
                    Editor1.Range(i + 1, i + 1).Selected
                End If
            End If
        End If
    End If
    Call ClearNoUseUndoList
LL:
    Editor1.ForceEdit = blnForce
    Editor1.Tag = ""
End Sub

'################################################################################################################
'## 功能：  删除指定范围的文本，剔除受保护的文本，返回删除的字符数
'################################################################################################################
Private Function DelTextRange(ByRef edtThis As Editor, ByVal lS As Long, ByVal lE As Long) As Long
    Dim i As Long, j As Long, lStart As Long, lNum As Long, lSum As Long
    Dim bProtect1 As Boolean, bProtect2 As Boolean
    edtThis.Tag = "DelTextRange"
    If Me.Document.EditType <> cprET_病历文件定义 Then
        '剔除保护文本
        lStart = lS
        For i = lS To lE - 1
            bProtect1 = edtThis.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And edtThis.Range(i, i + 1).Font.Protected
            bProtect2 = edtThis.Range(i + 1, i + 2).Font.ForeColor = PROTECT_FORECOLOR And edtThis.Range(i + 1, i + 2).Font.Protected
            If bProtect1 = bProtect2 Then
                '文本状态相同，不处理
            Else
                '文本状态不同
                If bProtect1 Then
                    '前一位置为保护文本，不处理
                    lStart = i + 1  '记录非保护文本的起始位置
                Else
                    '前一位置为非保护文本，则清除之
                    edtThis.Range(lStart, i + 1) = ""
                    lNum = i + 1 - lStart   '本次删除的字符数
                    lSum = lSum + lNum
                    lE = lE - lNum
                    i = lStart - 1          'lStart不变
                End If
            End If
        Next
        '如果一直是状态相同
        If (bProtect1 = bProtect2) And (bProtect1 = False) And lStart < lE Then
            edtThis.Range(lStart, lE) = ""
            lNum = lE - lStart
            lSum = lSum + lNum
        End If
        DelTextRange = lSum
    Else
        edtThis.Range(lS, lE) = ""
        DelTextRange = lE - lS
    End If
    edtThis.Tag = ""
End Function

'################################################################################################################
'## 功能：  允许用户个性化设置菜单和工具栏
'################################################################################################################
Private Sub cbrThis_Customization(ByVal Options As XtremeCommandBars.ICustomizeOptions)
    Dim Controls As CommandBarControls
    Set Controls = cbrThis.DesignerControls

    If (Controls.Count = 0) Then
        AddButton Controls, xtpControlButton, ID_FILE_CLEAR, "清空", , "清空当前文件所有内容", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_IMPORT, "引入...", , "引入外部定义的文件", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_SAVE, "保存", , "保存当前编辑的文件", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_SAVE_QUIT, "保存退出", , "保存当前编辑的文件并退出编辑器", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_SAVEASEPRDEMO, "另存为范文...", , "另存为范文", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_SAVEASSEGMENT, "另存为片段...", , "另存为示范片段", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_SAVEAS, "导出为RTF文件...", , "导出为RTF文件", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_EXPORTTOXML, "导出为XML文件...", , "导出为XML文件", xtpButtonAutomatic, "文件"
'        AddButton Controls, xtpControlButton, ID_FILE_EXPORTTOHTML, "导出为HTML文件...", , "导出为HTML文件", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_IMPORTFROMXML, "从XML文件导入...", , "从XML文件导入", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_PAGESETUP, "页面设置...", , "页眉页脚、页面尺寸设置", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_PRINTPREVIEW, "打印预览", , "打印预览", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_PRINT, "打印...", , "打印当前文件", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_PRINTINWORD, "通过Word打印", , "通过Word打印当前文件", xtpButtonAutomatic, "文件"
        AddButton Controls, xtpControlButton, ID_FILE_EXIT, "退出", , "退出系统", xtpButtonAutomatic, "文件"

        AddButton Controls, xtpControlButton, ID_EDIT_UNDO, "撤销", , "撤销最近一次编辑", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_REDO, "重做", , "重复最近一次编辑", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_CUT, "剪切", , "剪切", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_COPY, "复制", , "复制", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_PASTE, "粘贴", , "粘贴", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_FORMATBRUSH, "格式刷", , "格式刷", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_REFCOMPEND, "刷新提纲", , "刷新提纲", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_ADDCOMPEND, "新增提纲", , "新增提纲", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_DELCOMPEND, "删除提纲", , "删除提纲", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_MODCOMPEND, "修改提纲", , "修改提纲", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_DELETE, "删除", , "删除所选内容", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_SELECTALL, "全选", , "全选", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_FIND, "查找...", , "查找...", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_FINDNEXT, "查找下一个", , "查找下一个", xtpButtonAutomatic, "编辑"
        AddButton Controls, xtpControlButton, ID_EDIT_REPLACE, "替换...", , "替换...", xtpButtonAutomatic, "编辑"

        AddButton Controls, xtpControlButton, ID_VIEW_STRUCTURE, "文档结构图", , "文档结构图", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_PHRASEDEMO, "示范词句列表", , "示范词句列表", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_SEGMENT, "示范片段列表", , "示范片段列表", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_SEGMENT, "报告图列表", , "报告图列表", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_HISTORYWINDOW, "历史内容列表", , "历史内容列表", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_HISTORYREPORT, "历史报告列表", , "历史报告列表", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, XTP_ID_TOOLBARLIST, "工具栏列表", , "工具栏列表", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_HEADFOOT, "页眉页脚", , "页眉页脚", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_CHARCOUNT, "字数统计", , "字数统计", xtpButtonAutomatic, "视图"
'        AddButton Controls, xtpControlButton, ID_VIEW_GRID, "网格线", , "网格线", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_RULER, "标尺", , "标尺", xtpButtonAutomatic, "视图"
        AddButton Controls, xtpControlButton, ID_VIEW_PENWINDOW, "手写输入窗口", , "手写输入窗口", xtpButtonAutomatic, "视图"

        AddButton Controls, xtpControlButton, ID_INSERT_DATETIME, "日期和时间", , "日期和时间", xtpButtonAutomatic, "插入"
        AddButton Controls, xtpControlButton, ID_INSERT_DATE, "插入日期", , "插入日期", xtpButtonAutomatic, "插入"
        AddButton Controls, xtpControlButton, ID_INSERT_TIME, "插入时间", , "插入时间", xtpButtonAutomatic, "插入"
        AddButton Controls, xtpControlButton, ID_INSERT_SPECIALCHAR, "特殊符号", , "特殊符号", xtpButtonAutomatic, "插入"
        AddButton Controls, xtpControlButton, ID_TABLE_INSERTTABLE, "插入表格", , "插入表格", xtpButtonAutomatic, "插入"
        AddButton Controls, xtpControlButton, ID_INSERT_ELEMENT, "诊治要素", , "诊治要素", xtpButtonAutomatic, "插入"
        AddButton Controls, xtpControlButton, ID_EDIT_ADDCOMPEND, "插入提纲", , "插入提纲", xtpButtonAutomatic, "插入"
        AddButton Controls, xtpControlButton, ID_INSERT_EPRDEMO, "导入范文", , "导入范文", xtpButtonAutomatic, "插入"

        AddButton Controls, xtpControlButton, ID_FORMAT_FONT, "字体...", , "字体...", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_PARA, "段落...", , "段落...", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_BOLD, "粗体", , "粗体", xtpButtonAutomatic, "格式"
'        AddButton Controls, xtpControlButton, ID_FORMAT_ITALIC, "斜体", , "斜体", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_SUPER, "上标", , "上标", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_SUB, "下标", , "下标", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_NONE, "无下划线", , "无下划线", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_THIN, "细下划线", , "细下划线", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_THICK, "粗下划线", , "粗下划线", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_WAVE, "波浪下划线", , "波浪下划线", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_DOT, "点下划线", , "点下划线", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_DASH, "虚下划线", , "虚下划线", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT, "点划下划线", , "点划下划线", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT2, "双点划下划线", , "双点划下划线", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_ALIGNLEFT, "左对齐", , "左对齐", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_ALIGNCENTER, "居中对齐", , "居中对齐", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_ALIGNRIGHT, "右对齐", , "右对齐", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTNONE, "无项目符号与编号", , "无项目符号与编号", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTBULLETS, "项目符号(·)", , "项目符号(·)", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTARABIC, "阿拉伯数字(1,2,3,...)", , "阿拉伯数字(1,2,3,...)", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTLCHAR, "小写字母(a,b,c,...)", , "小写字母(a,b,c,...)", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTUCHAR, "大写字母(A,B,C,...)", , "大写字母(A,B,C,...)", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTLROME, "小写罗马数字(i,ii,iii,...)", , "小写罗马数字(i,ii,iii,...)", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTUROME, "大写罗马数字(I,II,III,...)", , "大写罗马数字(I,II,III,...)", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTSETUP, "自定义格式...", , "自定义格式...", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE1, "1.0", , "1.0", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE2, "1.3", , "1.3", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE3, "1.5", , "1.5", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE4, "2.0", , "2.0", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE5, "2.5", , "2.5", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE6, "3.0", , "3.0", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE7, "其他...", , "其他...", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_SPACEBEFORE, "段前间距", , "段前间距", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_SPACEAFTER, "段后间距", , "段后间距", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_FIRSTINDENT, "首行缩进", , "首行缩进", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_FIRSTHUNGING, "首行悬挂", , "首行悬挂", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_INDENTDECREASE, "减少缩进量", , "减少缩进量", xtpButtonAutomatic, "格式"
        AddButton Controls, xtpControlButton, ID_FORMAT_INDENTINCREASE, "增加缩进量", , "增加缩进量", xtpButtonAutomatic, "格式"

        AddButton Controls, xtpControlButton, ID_HELP_CONTENT, "帮助主题", , "帮助主题", xtpButtonAutomatic, "帮助"
        AddButton Controls, xtpControlButton, ID_HELP_ONLINE, gstrProductName & "在线", , gstrProductName & "在线", xtpButtonAutomatic, "帮助"
        AddButton Controls, xtpControlButton, ID_HELP_WEBFORUM, gstrProductName & "论坛", , gstrProductName & "论坛", xtpButtonAutomatic, "帮助"
        AddButton Controls, xtpControlButton, ID_HELP_CONTACT, "发送反馈", , "发送反馈", xtpButtonAutomatic, "帮助"
        AddButton Controls, xtpControlButton, ID_HELP_ABOUT, "关于...", , "关于...", xtpButtonAutomatic, "帮助"

        AddButton Controls, xtpControlButton, ID_REVISION_PREV, "前一处修订", , "前一处修订", xtpButtonAutomatic, "修订"
        AddButton Controls, xtpControlButton, ID_REVISION_NEXT, "后一处修订", , "后一处修订", xtpButtonAutomatic, "修订"
        AddButton Controls, xtpControlButton, ID_REVISION_RESET, "清除修订", , "清除修订", xtpButtonAutomatic, "修订"

        AddButton Controls, xtpControlButton, ID_DIAGNOSIS, "诊断", , "诊断", xtpButtonAutomatic, "诊断"

        AddButton Controls, xtpControlButton, ID_PATISIGN, "患者签名", , "患者签名", xtpButtonAutomatic, "签名"
        AddButton Controls, xtpControlButton, ID_SIGN, "签名", , "签名", xtpButtonAutomatic, "签名"
        AddButton Controls, xtpControlButton, ID_UNTREAD, "回退", , "回退", xtpButtonAutomatic, "签名"
        AddButton Controls, xtpControlButton, ID_SIGN_QUIT, "签名退出", , "签名并退出编辑", xtpButtonAutomatic, "签名"
    End If
End Sub

'################################################################################################################
'## 功能：  菜单&工具栏执行事件
'################################################################################################################

Private Function AskOutputMode(ByRef blnOrigMode As Boolean, ByVal blnPreview As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：判断是否采用最终格式还是原始格式打印/预览
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strAsk As String
    
    If Document.目标版本 > 1 And Document.EPRFileInfo.种类 = cpr诊疗报告 Then
        If zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1" Then
            blnOrigMode = False
        Else
            strAsk = "“" & Document.EPRFileInfo.名称 & "”进行过修订。"
            strAsk = strAsk & "可以按清洁格式或原始格式" & IIf(blnPreview, "预览", "打印") & "："
            strAsk = strAsk & vbCrLf & "    最终格式：不包含修改痕迹的清洁格式"
            strAsk = strAsk & vbCrLf & "    原始格式：包含修改痕迹的草稿格式"
            strAsk = strAsk & vbCrLf & "按“最终格式”模式" & IIf(blnPreview, "预览", "打印") & "吗？"
            
            Select Case MsgBox(strAsk, vbYesNoCancel + vbQuestion, gstrSysName)
            Case vbYes
                blnOrigMode = False
            Case vbNo
                blnOrigMode = True
            Case Else
                Exit Function
            End Select
        End If
        
        AskOutputMode = True
    Else
        AskOutputMode = True
    End If
        
End Function

Private Sub cbrThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    On Error Resume Next
    Dim i As Long, j As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim bFinded As Boolean
    Dim blnForce As Boolean, blnModified As Boolean
    Dim lngLen As Long
    Dim lRow1 As Long, lRow2 As Long, lCol1 As Long, lCol2 As Long, blnTmp As Boolean, sText As String, strPic As String, objPic As StdPicture
    Dim blnOrigMode As Boolean
    Dim lngVkState As Long
    
    If mblnPrecess Then
        Exit Sub
    End If
    
    mblnPrecess = True
    If tblThis.Visible Then
        If tblThis.SelStartRow > tblThis.SelEndRow Then
            lRow1 = tblThis.SelEndRow
            lRow2 = tblThis.SelStartRow
        Else
            lRow1 = tblThis.SelStartRow
            lRow2 = tblThis.SelEndRow
        End If
        If tblThis.SelStartCol > tblThis.SelEndCol Then
            lCol1 = tblThis.SelEndCol
            lCol2 = tblThis.SelStartCol
        Else
            lCol1 = tblThis.SelStartCol
            lCol2 = tblThis.SelEndCol
        End If
    End If

    blnForce = Me.Editor1.ForceEdit
    Select Case Control.ID
    Case ID_FILE_CLEAR: Call ClearDoc: Call RecountPage(True)
    Case ID_FILE_IMPORT: Call ImportEPRDoc: Call RecountPage(True)
    Case ID_FILE_SAVE, ID_FILE_SAVE_QUIT
        If SaveEMRDoc And Control.ID = ID_FILE_SAVE_QUIT Then
            mblnPrecess = False: Unload Me: Exit Sub
        End If
        If Editor1.Enabled Then
            Editor1.SetFocus
        End If
    Case ID_FILE_SAVEAS
        If SaveDocToFile Then MsgBox "导出成功！", vbOKOnly + vbInformation, gstrSysName
    Case ID_FILE_SAVEASEPRDEMO: Call SaveDocAsEPRDemo
    Case ID_FILE_SAVEASSEGMENT: Call SaveDocAsSegment
    Case conMenu_File_Parameter
        '参数设置
        Dim frmSetup As New frmAutoSaveSetup
        
        If frmSetup.ShowMe(Me, gstrPrivsEpr) Then
            mblnAutosave = zlDatabase.GetPara("AutoSave", glngSys, 1070, 1) = 1
            mlngUndoLimit = zlDatabase.GetPara("UndoLimit", glngSys, 1070, 20)
            mlngSaveInterval = zlDatabase.GetPara("SaveInterval", glngSys, 1070, 60)
            mblnAutoSaveEPR = zlDatabase.GetPara("AutoSaveEPR", glngSys, 1070, 0) = 1
            mlngSaveIntervalEPR = zlDatabase.GetPara("SaveIntervalEPR", glngSys, 1070, 5)
            mblnAutoPageCount = zlDatabase.GetPara("AutoPageCount", glngSys, 1070, 0) = 1
            mblnAutoPageNote = zlDatabase.GetPara("AutoPageNote", glngSys, 1070, 0) = 1
            
            If mintSharePages <> Val(zlDatabase.GetPara("SharePageCount", glngSys, 1070, 5)) Then
                mintSharePages = Val(zlDatabase.GetPara("SharePageCount", glngSys, 1070, 5))
                mblnExistHistroy = ShowSharePageHistory(Me.Document, mintSharePages)
            End If
        End If
    Case ID_FILE_EXPORTTOXML:
        Call ExportXML
    Case ID_FILE_EXPORTTOHTML
        '导出到HTML中
        Dim strHTML As String
        Select Case Me.Document.EditType
        Case cprET_病历文件定义
            dlgThis.Filename = "定义_" & Me.Document.EPRFileInfo.名称 & ".htm"
        Case cprET_全文示范编辑
            dlgThis.Filename = "范文_" & Me.Document.EPRFileInfo.名称 & "_" & Me.Document.EPRDemoInfo.名称 & ".htm"
        Case cprET_单病历编辑, cprET_单病历审核
            dlgThis.Filename = "记录_" & Me.Document.EPRFileInfo.名称 & "(" & Me.Document.EPRPatiRecInfo.ID & "," & Me.Document.目标版本 & ").htm"
        End Select

        dlgThis.Filter = "*.htm|*.htm|*.html|*.html|*.*|*.*"
        dlgThis.CancelError = True
        On Error GoTo out
        dlgThis.ShowSave
        strHTML = dlgThis.Filename
        If gobjFSO.FileExists(strHTML) Then
            If MsgBox("该文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then GoTo out
        End If
        If Me.Document.ExportToHTML(Me.Editor1, strHTML) Then
            MsgBox "成功导出为HTML文件！" & vbCrLf & "文件名:" & strHTML, vbOKOnly + vbInformation, gstrSysName
        End If
    Case ID_FILE_IMPORTFROMXML
        '从XML文件导入
        Dim strXML As String
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error GoTo out
        dlgThis.ShowOpen
        strXML = dlgThis.Filename
        If gobjFSO.FileExists(strXML) Then
            If Me.Document.EPRPatiRecInfo.签名级别 > cprSL_空白 Or Me.Document.EPRPatiRecInfo.最后版本 > 1 Then
                MsgBox "只允许在书写或者文件定义时进行XML导入导出操作！", vbOKOnly + vbInformation, gstrSysName
                GoTo out
            End If
            If Me.Document.Signs.Count > 0 And _
                (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                MsgBox "不允许在已签名的文件中进行XML导入操作！", vbOKOnly + vbInformation, gstrSysName
                GoTo out
            Else
                Call AddUndoPoint  '手动缓存
                If Me.Document.ImportFromXMLFile(Me.Editor1, strXML) Then
                    Me.Document.Compends.UpdateOrdersFromText Me.Editor1
                    Me.Document.Compends.FillTree mfrmCompends.Tree
                End If
                Call ClearNoUseUndoList
            End If
            Call RecountPage(True)
        End If
    Case ID_FILE_PAGESETUP
        Editor1.ShowPageSetupDlg
        Call RecountPage(True)
    Case ID_FILE_PRINTPREVIEW
        blnModified = Me.Editor1.Modified
        If AskOutputMode(blnOrigMode, False) Then
            Call PrintEPRDoc(True, Not blnOrigMode)
        End If
        Me.Editor1.Modified = blnModified
    Case ID_FILE_PRINT
        blnModified = Me.Editor1.Modified
        If AskOutputMode(blnOrigMode, False) Then
            Call PrintEPRDoc(False, Not blnOrigMode)
        End If
        Me.Editor1.Modified = blnModified
    Case ID_FILE_PRINTINWORD
        Call PrintInWord
    Case ID_FILE_EXIT, ID_COMMON_CANCEL
        mblnPrecess = False: Unload Me: Exit Sub
    Case ID_EDIT_UNDO
'        Editor1.Undo
        Call Undo
        Call RecountPage
    Case ID_EDIT_REDO
'        Editor1.Redo
        Call Redo
        Call RecountPage
    Case ID_EDIT_CUT
        gstrCopyPID = CStr(Document.EPRPatiRecInfo.病人ID)
        Call ExecCut
        Call RecountPage
    Case ID_EDIT_COPY
        If Control.Enabled And Control.Visible Then '快捷键执行时需要判断
            gstrCopyPID = CStr(Document.EPRPatiRecInfo.病人ID)
            If Me.ActiveControl Is edtThis Then
                edtThis.Copy    '允许以文本方式拷贝到其他程序（放到剪贴板）
                Call ExecCopy(Me.edtThis)   '缓存内容（关键字未修正）
            Else
                Editor1.Copy    '允许以文本方式拷贝到其他程序（放到剪贴板）
                Call ExecCopy(Me.Editor1)   '缓存内容（关键字未修正）
            End If
            Call RecountPage
        End If
    Case ID_EDIT_COPYSELF  '专用复制
            Call SpicalCopy(Control.Enabled, Control.Visible)
    Case ID_EDIT_COPYOUT  '复制到粘贴板
        Editor1.Copy
        gstrCopyPID = CStr(Document.EPRPatiRecInfo.病人ID)
    Case ID_EDIT_SAVEASPHRASE
        '存为示范词句
        Call ExecSaveAsPhrase
    Case ID_EDIT_PASTE
        If Control.Enabled And Control.Visible Then '快捷键执行时需要判断
            If Control.Parent Is Nothing Then
                'Control.Parent为空表示由绑定热键触发
                lngVkState = GetAsyncKeyState(Asc("V"))
                
                'lngVkState如果为0表示Ctrl+v的v键没有被按下，如果产品中使用了HooK，当松开v键时，也会触发绑定热键
                If lngVkState = 0 Then
                    GoTo out
                End If
            End If
            
            If gstrCopyPID <> "" And gstrCopyPID <> CStr(Document.EPRPatiRecInfo.病人ID) And Document.EPRFileInfo.种类 <> cpr诊疗报告 And InStr(gstrPrivsEpr, "复制他人病历") <= 0 Then
                MsgBox "仅能引用病人自身病历内容，禁止复制他人病历！", vbExclamation, gstrSysName
                gstrCopyPID = ""
                On Error Resume Next
                Clipboard.Clear
                gfrmPublic.edtPublic.NewDoc: Set gfrmPublic.Elements = New cEPRElements
                GoTo out
            End If
            Call ExecAuditDelete
            Call ExecPaste(Me.Editor1)   '粘贴内容（修正关键字）
            Call RecountPage
        End If
    Case ID_EDIT_DELETE
        If Editor1.ViewMode = cprNormal Then
            Call ExecDelete
            Call RecountPage
        End If
    Case ID_EDIT_BACKSPACE
        If Editor1.ViewMode = cprNormal Then
            Call ExecBackSpace
            Call RecountPage
        End If
    Case ID_EDIT_SELECTALL
        Editor1.SelectAll
    Case ID_EDIT_FIND
        Editor1.ShowFindReplaceDlg 0
    Case ID_EDIT_FINDNEXT
        Editor1.FindNext
    Case ID_EDIT_REPLACE
        Editor1.ShowFindReplaceDlg IIf(Me.Editor1.AuditMode, -1, 1)
        Call RecountPage(True)
    Case ID_VIEW_STRUCTURE
        If mfrmCompends.Visible Then
            DkpThis.FindPane(ID_VIEW_STRUCTURE).Close
        Else
            DkpThis.ShowPane ID_VIEW_STRUCTURE
        End If
    Case ID_VIEW_PHRASEDEMO
        If mfrmSentenceDetailed.Visible Then
            DkpThis.FindPane(ID_VIEW_PHRASEDEMO).Close
        Else
            DkpThis.ShowPane ID_VIEW_PHRASEDEMO
        End If
    Case ID_VIEW_SEGMENT
        If mfrmSegments.Visible Then
            DkpThis.FindPane(ID_VIEW_SEGMENT).Close
        Else
            DkpThis.ShowPane ID_VIEW_SEGMENT
        End If
    Case ID_VIEW_PACSPIC
        If mfrmPacsPic.Visible Then
            DkpThis.FindPane(ID_VIEW_PACSPIC).Close
        Else
            DkpThis.ShowPane ID_VIEW_PACSPIC
        End If
    Case ID_VIEW_HISTORYREPORT
        If mfrmHistoryReport.Visible Then
            DkpThis.FindPane(ID_VIEW_HISTORYREPORT).Close
        Else
            DkpThis.ShowPane ID_VIEW_HISTORYREPORT
        End If
    Case ID_VIEW_HISTORYWINDOW
        
        picPane.Visible = Not picPane.Visible
        Call picHistoryInfo_Resize
        
        cbrThis.RecalcLayout
        
    Case ID_EDIT_REFCOMPEND
        Me.Document.Compends.UpdateOrdersFromText Me.Editor1
        Me.Document.Compends.FillTree mfrmCompends.Tree
    Case ID_EDIT_ADDCOMPEND
        If Editor1.ViewMode = cprNormal Then
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys = False And Editor1.Selection.Font.Protected = False Then
                Call AddUndoPoint  '手动缓存
                Dim f_InsCompend As New frmInsCompend
                f_InsCompend.ShowMe Me, Me.Editor1, Me.Document.Compends
                Call ClearNoUseUndoList
                Call RecountPage(True)
            End If
        End If
    Case ID_EDIT_MODCOMPEND
        If Not mfrmCompends.Tree.SelectedItem Is Nothing Then
            lKey = mfrmCompends.Tree.SelectedItem.Tag
            If lKey > 0 Then
                If Document.Compends("K" & lKey).预制提纲ID <> 0 And Document.EditType <> cprET_病历文件定义 Then
                    MsgBox "不允许编辑保留提纲！", vbOKOnly + vbInformation, gstrSysName
                    GoTo out
                End If
                Call AddUndoPoint  '手动缓存
                Dim f_ModCompend As New frmInsCompend
                f_ModCompend.ShowMe Me, Me.Editor1, Me.Document.Compends, Me.Document.Compends("K" & lKey)
                Call ClearNoUseUndoList
            End If
        End If
    Case ID_EDIT_DELCOMPEND
        If Not mfrmCompends.Tree.SelectedItem Is Nothing Then
            lKey = mfrmCompends.Tree.SelectedItem.Tag
            If lKey > 0 Then
                Call AddUndoPoint  '手动缓存
                Me.DeleteOutline lKey
                Call ClearNoUseUndoList
                Call RecountPage(True)
            End If
        End If
    Case ID_VIEW_HEADFOOT
        Editor1.Foot = Document.EPRFileInfo.页脚
        Editor1.Head = Document.EPRFileInfo.页眉
        If Editor1.ShowHeadFootDlg Then
            Document.EPRFileInfo.页脚 = Editor1.Foot
            Document.EPRFileInfo.页眉 = Editor1.Head
            If Me.Editor1.ViewMode = cprPaper Then
                '重新分页
                Me.Editor1.Freeze
                Me.Editor1.ViewMode = cprNormal
                Me.Editor1.ViewMode = cprPaper
                Me.Editor1.UnFreeze
            End If
            Me.Editor1.Modified = True
            Call RecountPage(True)
        End If
    Case ID_VIEW_CHARCOUNT
        Editor1.ShowCharCountDlg
    Case ID_VIEW_RULER
        Control.Checked = Not Control.Checked
        Editor1.ShowRuler = Control.Checked
    Case ID_VIEW_PENWINDOW
        If picPenInput.Visible Then
            picPenInput.Visible = False
        Else
            picPenInput.Visible = True
            If txtPenInput.Visible And txtPenInput.Enabled Then txtPenInput.SetFocus
        End If
    Case ID_INSERT_DATETIME
        If Editor1.ViewMode = cprNormal Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False Then
                        sText = Editor1.ShowInsertDateTimeDlg(, , , False)
                        If sText <> "" Then
                            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = sText
                            tblThis.Modified = True
                            tblThis.Refresh False, True, tblThis.SelectedCellKey
                            tblThis_Resize tblThis.Width, tblThis.Height
                        End If
                    End If
                End If
            Else
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then
                    Call AddUndoPoint  '手动缓存
                    Editor1.ShowInsertDateTimeDlg
                    Call ClearNoUseUndoList
                End If
            End If
            Call RecountPage
        End If
    Case ID_INSERT_DATE
        If Editor1.ViewMode = cprNormal Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False Then
                        tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Format(Now, "YYYY年MM月DD日")
                        tblThis.Modified = True
                        tblThis.Refresh False, True, tblThis.SelectedCellKey
                        tblThis_Resize tblThis.Width, tblThis.Height
                    End If
                End If
            Else
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False And Editor1.Selection.Font.Protected = False Then
                    Call AddUndoPoint  '手动缓存
                    Editor1.ForceEdit = True
                    Editor1.Tag = "cbrThis_ExeCute"
                    If Me.Editor1.AuditMode Then
                        Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos).Selected
                        '保留性属性（便于新增文本）
                        On Error Resume Next
                        Me.Editor1.OriginRTB.SelColor = Me.Document.GetNewCharColor(Me.Editor1.OriginRTB.SelColor)
                        Me.Editor1.OriginRTB.SelStrikeThru = False
                    End If
                    Editor1.Selection.Text = Format(Now, "YYYY年MM月DD日")
                    Editor1.Range(Editor1.Selection.StartPos + Len(Editor1.Selection.Text), Editor1.Selection.StartPos + Len(Editor1.Selection.Text)).Selected
                    Me.Editor1.ForceEdit = blnForce
                    Editor1.Tag = ""
                    Call ClearNoUseUndoList
                End If
            End If
            Call RecountPage
        End If
    Case ID_INSERT_TIME
        If Editor1.ViewMode = cprNormal Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False Then
                        tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Format(Now, "HH时mm分")
                        tblThis.Modified = True
                        tblThis.Refresh False, True, tblThis.SelectedCellKey
                        tblThis_Resize tblThis.Width, tblThis.Height
                    End If
                End If
            Else
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False And Editor1.Selection.Font.Protected = False Then
                    Call AddUndoPoint  '手动缓存
                    Editor1.ForceEdit = True
                    Editor1.Tag = "cbrThis_ExeCute"
                    If Me.Editor1.AuditMode Then
                        Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos).Selected
                        '保留性属性（便于新增文本）
                        On Error Resume Next
                        Me.Editor1.OriginRTB.SelColor = Me.Document.GetNewCharColor(Me.Editor1.OriginRTB.SelColor)
                        Me.Editor1.OriginRTB.SelStrikeThru = False
                    End If
                    Editor1.Selection.Text = Format(Now, "HH时mm分")
                    Editor1.Range(Editor1.Selection.StartPos + Len(Editor1.Selection.Text), Editor1.Selection.StartPos + Len(Editor1.Selection.Text)).Selected
                    Me.Editor1.ForceEdit = blnForce
                    Editor1.Tag = ""
                    Call ClearNoUseUndoList
                End If
            End If
            Call RecountPage
        End If
    Case ID_INSERT_SPECIALCHAR
        If Editor1.ViewMode = cprNormal Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False Then
                        sText = Editor1.ShowInsertSymbolDlg(False, IIf(InStr(mstrSex, "男") > 0, 1, IIf(InStr(mstrSex, "女") > 0, 2, 0)), True, strPic, objPic)
                        If sText = "" Then GoTo out
                        If sText <> "" Then
                            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = sText
                            tblThis.Modified = True
                            tblThis.Refresh False, True, tblThis.SelectedCellKey
                            tblThis_Resize tblThis.Width, tblThis.Height
                        End If
                    End If
                End If
            Else
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then
                    Call AddUndoPoint  '手动缓存
                    sText = Editor1.ShowInsertSymbolDlg(True, IIf(InStr(mstrSex, "男") > 0, 1, IIf(InStr(mstrSex, "女") > 0, 2, 0)), False, strPic, objPic)
                    If Not objPic Is Nothing Then '图片公式
                        Editor1.Tag = "cbrThis_ExeSPECIALCHAR"
                        InsertPicture EPRFormulaPicture, objPic, objPic.Width, objPic.Height, strPic
                        Editor1.Tag = ""
                    End If
                    Call ClearNoUseUndoList
                End If
            End If
            Call RecountPage
        End If
    Case ID_INSERT_TABLE
        If Editor1.ViewMode = cprNormal Then
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys = False Then
                Call AddUndoPoint  '手动缓存
                Dim ScreenPoint As POINTAPI
                GetCursorPos ScreenPoint
                ShowTablePicker ScreenPoint.X * 15, ScreenPoint.y * 15
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_INSERTTABLE
        If Editor1.ViewMode = cprNormal Then
            If tblThis.Visible = False Then
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then
                    Dim frmInsTb As New frmInsTable, lR As Long, lC As Long
                    If frmInsTb.ShowMe(Me, lR, lC) Then
                        Call AddUndoPoint  '手动缓存

                        lKey = Me.Document.Tables.Add
                        tblThis.AutoHeight = True
                        tblThis.Redraw = False
                        tblThis.SingleClickEdit = False
                        tblThis.HighlightMode = HMFilledRectAlpha
                        tblThis.Width = Me.Editor1.PaperWidth - Me.Editor1.Selection.Para.LeftIndent - Me.Editor1.MarginLeft - Me.Editor1.MarginRight - 800
                        tblThis.Init lR, lC
                        tblThis.CellMargin = 10
                        For i = 1 To lR
                            For j = 1 To lC
                                Me.Document.Tables("K" & lKey).Cells.Add , i, j
                            Next j
                        Next i
                        tblThis.Tag = lKey
                        tblThis.ShowToolTipText = True
                        tblThis.MinRowHeight = 300
                        tblThis.Redraw = True
                        tblThis.Refresh
                        SaveUIToTable Me.Document.Tables("K" & lKey), True

                        Call ClearNoUseUndoList
                        Call RecountPage
                    End If
                    Unload frmInsTb: Set frmInsTb = Nothing
                End If
            End If
        End If
    Case ID_INSERT_PICTURE
        If Editor1.ViewMode = cprNormal Then
            Dim frmInsertPic  As New frmInsertPicture
            If mbEditInTable Then
                frmInsertPic.ShowMe Me
            ElseIf ucPacsImgCanvas1.Visible Then
                frmInsertPic.ShowMe Me
            Else
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then
                    Call AddUndoPoint  '手动缓存
                    frmInsertPic.ShowMe Me, lngMaxWidth:=(Me.Editor1.PaperWidth - Me.Editor1.MarginLeft - Me.Editor1.MarginRight) / Screen.TwipsPerPixelX, lngMaxHeight:=(Me.Editor1.PaperHeight - Me.Editor1.MarginTop - Me.Editor1.MarginBottom) / Screen.TwipsPerPixelY
                    Call ClearNoUseUndoList
                End If
            End If
        End If
    Case ID_INSERT_ELEMENT
        If mbEditInTable Then
            If tblThis.Cells("K" & tblThis.SelectedCellKey).Picture Is Nothing Then
                lKey = Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag)
                If lKey > 0 Then
                    mfrmInsElement.Tag = lKey
                End If
                If mfrmInsElement.Tag = "" Then
                    '新增
                    mfrmInsElement.ShowMe Me
                Else
                    '修改
                    If Val(tblThis.Tag) > 0 Then mfrmInsElement.ShowMe Me, Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lKey), True
                End If
            End If
        Else
            bBeteenKeys = IsBetweenAnyKeys(Me.Editor1, Me.Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            Call AddUndoPoint  '手动缓存
            If bBeteenKeys Then
                '修改要素
                If sKeyType = "E" Then
                    If Me.Document.EditType <> cprET_病历文件定义 And Document.Elements("K" & lKey).保留对象 And InStr(1, gstrPrivsEpr, "保护文本处理") = 0 Then
                        '非定义模式下，不允许修改保留的诊治要素！
                        MsgBox "不能修改保护的诊治要素！", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                    If Document.Elements("K" & lKey).开始版 = Me.Document.目标版本 Then
                        mfrmInsElement.Tag = lKey
                        mfrmInsElement.ShowMe Me, Me.Document.Elements("K" & lKey), True, (Me.Document.EditType = cprET_病历文件定义)
                    End If
                End If
            Else
                '新增要素
                If mfrmInsElement.Visible Then
                    mfrmInsElement.Hide
                Else
                    mfrmInsElement.ShowMe Me, , , (Me.Document.EditType = cprET_病历文件定义)
                End If
            End If
            Call ClearNoUseUndoList
        End If
    Case ID_INSERT_EPRDEMO
        '范文引入
        Dim f_EPRDemo As New frmImportEPRDemo, lngEPRDemoID As Long
        lngEPRDemoID = f_EPRDemo.ShowMe(Me)
        If lngEPRDemoID > 0 Then
            Call AddUndoPoint  '手动缓存
            Me.Document.ImportEPRDemo Me.Editor1, lngEPRDemoID
            Call ClearNoUseUndoList
            Call RecountPage(True)
        End If
    Case ID_INSERT_DOCADVISE
        Call AddUndoPoint  '手动缓存
        Call ImportDocAdvice
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_INSERT_PACSPIC
        Call AddUndoPoint  '手动缓存
        Call InsertPacsPicTable
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_STYLEWINDOW
        '样式窗格
        If mfrmStyleMan.Visible Then
            DkpThis.FindPane(ID_FORMAT_STYLEWINDOW).Close
        Else
            DkpThis.ShowPane ID_FORMAT_STYLEWINDOW
        End If
    Case ID_FORMAT_STYLE
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        If Control.Type = xtpControlComboBox Then
            Call AddUndoPoint  '手动缓存
            If Control.Text = "其他..." Then
                DkpThis.ShowPane ID_FORMAT_STYLEWINDOW
            Else
                SetCommonStyle Editor1, Control.Text, Editor1.Selection.StartPos, Editor1.Selection.EndPos, True
            End If
            Call ClearNoUseUndoList
            Call RecountPage
        End If
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
    Case ID_FORMAT_FONTNAME
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            blnTmp = Not tblThis.Cell(lRow1, lCol1).FontBold
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FontName = Control.Text
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            If Control.Type = xtpControlComboBox Then
                Call AddUndoPoint  '手动缓存
                Editor1.Selection.Font.Name = Control.Text
                Call ClearNoUseUndoList
            End If
            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
        End If
        Call RecountPage
    Case ID_FORMAT_FONTSIZE
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            blnTmp = Not tblThis.Cell(lRow1, lCol1).FontBold
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FontSize = GetFontSizeNumber(Control.Text)
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            If Control.Type = xtpControlComboBox Then
                Call AddUndoPoint  '手动缓存
                Editor1.Selection.Font.Size = GetFontSizeNumber(Control.Text)
                Call ClearNoUseUndoList
            End If
            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
        End If
        Call RecountPage
    Case ID_FORMAT_FONT
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.ShowFontDlg 2 ^ 5 + 2 ^ 4 + 2 ^ 3 + 2 ^ 2 + 2 ^ 1 + 2 ^ 0
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_PARA
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.ShowParaDlg False
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_PROTECT
        Call AddUndoPoint  '手动缓存
        If Control.Checked Then
            '取消保护
            If Editor1.Selection.Font.ForeColor <> PROTECT_FORECOLOR Then
                GoTo out
            Else
                Editor1.ForceEdit = True
                Editor1.Tag = "cbrThis_ExeCute"
                Editor1.Selection.Font.Protected = False
                Editor1.Selection.Font.ForeColor = tomAutoColor
                Me.Editor1.ForceEdit = blnForce
                Editor1.Tag = ""
            End If
        Else
            '设置保护
            If Editor1.Selection.Font.Protected = True Or Editor1.Selection.Font.Hidden = True Or _
                Editor1.Selection.Font.BackColor <> tomAutoColor Then
                GoTo out
            Else
                Editor1.ForceEdit = True
                Editor1.Tag = "cbrThis_ExeCute"
                Editor1.Selection.Font.Protected = True
                Editor1.Selection.Font.ForeColor = PROTECT_FORECOLOR
                Me.Editor1.ForceEdit = blnForce
                Editor1.Tag = ""
            End If
        End If
        Call ClearNoUseUndoList
    Case ID_FORMAT_BOLD
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            blnTmp = Not tblThis.Cell(lRow1, lCol1).FontBold
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FontBold = blnTmp
                    tblThis.Cell(i, j).FontWeight = 0
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Call AddUndoPoint  '手动缓存
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            Editor1.Selection.Font.Bold = Not Editor1.Selection.Font.Bold
            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
        Call RecountPage
    Case ID_FORMAT_SUPER
        If tblThis.Visible Then GoTo out
        If Editor1.AuditMode Then '审核模式下，只有当前版本新增的可以更改，原有的文字不能更改
            If Not CanSetFormat Then
                MsgBox "当前为审核模式，上下标功能只能应用于本次审核新增的内容，请检查。", vbInformation, gstrSysName
                GoTo out  '选中文字中只要有一个字是原有文字即不可变更
            End If
        End If
            
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Font.Superscript = Not Editor1.Selection.Font.Superscript
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_SUB
        If tblThis.Visible Then GoTo out
        If Editor1.AuditMode Then '审核模式下，只有当前版本新增的可以更改，原有的文字不能更改
            If Not CanSetFormat Then
                MsgBox "当前为审核模式，上下标功能只能应用于本次审核新增的内容，请检查。", vbInformation, gstrSysName
                GoTo out  '选中文字中只要有一个字是原有文字即不可变更
            End If
        End If
        
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Font.Subscript = Not Editor1.Selection.Font.Subscript
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_UNDERLINE, ID_FORMAT_UNDERLINE_NONE, ID_FORMAT_UNDERLINE_THIN
        Call ExecuteUnderLine(Control, blnForce)
    Case ID_FORMAT_UNDERLINE_THICK, ID_FORMAT_UNDERLINE_WAVE, ID_FORMAT_UNDERLINE_DOT
        Call ExecuteUnderLine(Control, blnForce)
    Case ID_FORMAT_UNDERLINE_DASH, ID_FORMAT_UNDERLINE_DASHDOT, ID_FORMAT_UNDERLINE_DASHDOT2
        Call ExecuteUnderLine(Control, blnForce)
    Case ID_FORMAT_ALIGNLEFT
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignLeft
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Call AddUndoPoint  '手动缓存
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            Editor1.Selection.Para.Alignment = cprHALeft
            If tblThis.Visible Then
                tblThis.SetFocus
            Else
                If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
            End If

            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
    Case ID_FORMAT_ALIGNCENTER
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignCentre
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Call AddUndoPoint  '手动缓存
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            Editor1.Selection.Para.Alignment = cprHACenter
            If tblThis.Visible Then
                tblThis.SetFocus
            Else
                If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
            End If

            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
    Case ID_FORMAT_ALIGNRIGHT
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignRight
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Call AddUndoPoint  '手动缓存
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            Editor1.Selection.Para.Alignment = cprHARight
            If tblThis.Visible Then
                tblThis.SetFocus
            Else
                If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
            End If

            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
    Case ID_FORMAT_LISTNONE
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.ListType = cprLTNone
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTLCHAR
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.ListStart = 1
        Editor1.Selection.Para.ListTab = 25
        Editor1.Selection.Para.ListType = cprLTNumberAsLCLetter
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTUCHAR
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.ListStart = 1
        Editor1.Selection.Para.ListTab = 25
        Editor1.Selection.Para.ListType = cprLTNumberAsUCLetter
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTLROME
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.ListStart = 1
        Editor1.Selection.Para.ListTab = 30
        Editor1.Selection.Para.ListType = cprLTNumberAsLCRoman
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTUROME
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.ListStart = 1
        Editor1.Selection.Para.ListTab = 30
        Editor1.Selection.Para.ListType = cprLTNumberAsUCRoman
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTSETUP
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.ShowItemNumberDlg
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_SPACEBEFORE
        Dim l1 As Single
        l1 = Val(InputBox("输入段前间距的值，单位：磅。", gstrSysName, "0"))
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Call AddUndoPoint  '手动缓存
        Editor1.Selection.Para.SpaceBefore = l1
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_SPACEAFTER
        Dim l2 As Single
        l2 = Val(InputBox("输入段前间距的值，单位：磅。", gstrSysName, "0"))
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Call AddUndoPoint  '手动缓存
        Editor1.Selection.Para.SpaceAfter = l2
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_FIRSTINDENT
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.SetIndents 21, 0, 0
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_FIRSTHUNGING
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.SetIndents -21, 21, 0
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_INDENTDECREASE
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.LeftIndent = IIf(Editor1.Selection.Para.LeftIndent - 21 <= 0, 0, Editor1.Selection.Para.LeftIndent - 21)
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_INDENTINCREASE
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.LeftIndent = IIf(Editor1.Selection.Para.LeftIndent + 21 >= 300, 300, Editor1.Selection.Para.LeftIndent + 21)
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTARABIC
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        If Control.Checked Then
            Editor1.Selection.Para.ListType = cprLTNone
        Else
            Editor1.Selection.Para.ListStart = 1
            Editor1.Selection.Para.ListTab = 25
            Editor1.Selection.Para.ListType = cprLTNumberAsArabic
        End If
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTBULLETS
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        If Control.Checked Then
            Editor1.Selection.Para.ListType = cprLTNone
        Else
            Editor1.Selection.Para.ListTab = 12
            Editor1.Selection.Para.ListType = cprLTBullet
        End If
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LINESPACE, ID_FORMAT_LINESPACE1, ID_FORMAT_LINESPACE2, ID_FORMAT_LINESPACE3
        Call ExecuteLineSpace(Control, True)
    Case ID_FORMAT_LINESPACE4, ID_FORMAT_LINESPACE5, ID_FORMAT_LINESPACE6, ID_FORMAT_LINESPACE7
        Call ExecuteLineSpace(Control, True)
    Case ID_FORMAT_HIGHLIGHT
        Call AddUndoPoint  '手动缓存
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        ColorHighlight_pOK
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
    Case ID_TABLE_CELLALIGNMENT
        Debug.Print "ID_TABLE_CELLALIGNMENT"

    Case ID_DRAW_FILLCOLOR
        ColorFillColor_pOK
    Case ID_FORMAT_FORECOLOR
        ColorForeColor_pOK
    'Public Const ID_HELP_CONTENT = 500
    'Public Const ID_HELP_ASSISTANT = 501
    'Public Const ID_HELP_CONTACT = 502
    'Public Const ID_HELP_ONLINE = 503
    'Public Const ID_HELP_ABOUT = 504
    Case ID_HELP_CONTENT
        ShowHelp App.ProductName, Me.hwnd, "frmMain", Int((glngSys) / 100)
    Case ID_HELP_CONTACT
        Call zlMailTo(Me.hwnd)
    Case ID_HELP_ONLINE
        Call zlHomePage(Me.hwnd)
    Case ID_HELP_WEBFORUM
        Call zlWebForum(Me.hwnd)
    Case ID_HELP_ABOUT
        ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
    Case ID_EDIT_FORMATBRUSH
        If mblnFmtBrushDown = False Then
            Call AddUndoPoint  '手动缓存
            mblnFmtBrushDown = True
            Me.Editor1.OriginRTB.MousePointer = 99
            Me.Editor1.OriginRTB.MouseIcon = picPatiInfo.MouseIcon
            Dim lS As Long, lE As Long
            With Editor1
                lS = .Selection.StartPos
                lE = .Selection.EndPos
                If lE > lS + 1 Then
                    If .Range(lE - 2, lE) = vbCrLf Then
                        '保存段落属性
                        Set mParaFmt = New cParaFormat
                        Set mFontFmt = New cFontFormat
                        Set mParaFmt = Editor1.Range(lS, lS + 1).Para.GetParaFmt
                        Set mFontFmt = Editor1.Range(lS, lS + 1).Font.GetFontFmt
                    Else
                        '只保存字体属性
                        Set mParaFmt = Nothing
                        Set mFontFmt = New cFontFormat
                        Set mFontFmt = Editor1.Range(lS, lS + 1).Font.GetFontFmt
                    End If
                Else
                    '只保存字体属性
                    Set mParaFmt = Nothing
                    Set mFontFmt = New cFontFormat
                    Set mFontFmt = Editor1.Range(lS, lS + 1).Font.GetFontFmt
                End If
            End With
            Call ClearNoUseUndoList
            Call RecountPage
        Else
            Me.Editor1.OriginRTB.MousePointer = 0
            mblnFmtBrushDown = False
            Set mParaFmt = Nothing
            Set mFontFmt = Nothing
        End If
    Case ID_INSERT_AUTORECOGNISE
        '自动识别诊治要素或者字典项目
        Dim strAuto As String
        If tblThis.Visible Then
            If Val(tblThis.Tag) > 0 Then
                If tblThis.InEdit Then tblThis.EndEdit
                strAuto = Trim(tblThis.Cells("K" & tblThis.SelectedCellKey).Text)
                If strAuto = "" Then GoTo out
                If Len(strAuto) > 100 Then strAuto = Left(strAuto, 100)
                ShowAutoRecSelector strAuto
            End If
        Else
            strAuto = Trim(Me.Editor1.SelText)
            If strAuto = "" Then GoTo out
            If Len(strAuto) > 100 Then strAuto = Left(strAuto, 100)
            Call AddUndoPoint  '手动缓存
            ShowAutoRecSelector strAuto
            Call ClearNoUseUndoList
        End If
        Call RecountPage
    Case ID_EDIT_MARKEDPIC
        If tblThis.Visible Then
            If Val(tblThis.Tag) > 0 Then
                '标记图
                lKey = Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag)
                If lKey > 0 Then
'                    Dim frmPictureEditor1 As New frmPictureEditor
'                    If frmPictureEditor1.ShowMe(Me, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey), True) Then
'                        '保存结果图片到表格中
'                        Set tblThis.Cells("K" & tblThis.SelectedCellKey).Picture = Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey).DrawFinalPic
'                        tblThis.Modified = True
'                    End If
                    '编辑图片
                    Dim LL As Long, lT As Long, lW As Long, lH As Long
                    tblThis.Cells("K" & tblThis.SelectedCellKey).GetCellPictureBorder LL, lT, lW, lH
                    ucPictureEditor1.ShowMe Me, tblThis.hwnd, cbrThis, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey), _
                        LL, lT, lW, lH, True, Me.Document.Tables("K" & tblThis.Tag)
                End If
            End If
        Else
            If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal And (Me.Editor1.AuditMode = False) Then
                '查找关键字 ID ！
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then GoTo out
                If sKeyType = "P" Then
                    Call AddUndoPoint  '手动缓存
'                    Dim frmPictureEditor2 As New frmPictureEditor
'                    frmPictureEditor2.ShowMe Me, Me.Document.Pictures("K" & lKey)
                    '编辑图片
                    Editor1.ShowUIInterface
                    ucPictureEditor1.ShowMe Me, Editor1.hwnd, cbrThis, Me.Document.Pictures("K" & lKey), _
                        Editor1.UILeft, Editor1.UITop, Editor1.UIWidth, Editor1.UIHeight, False

                    Call ClearNoUseUndoList
                End If
            End If
        End If
    Case ID_EDIT_OUTERPIC
        If tblThis.Visible Then
            If Val(tblThis.Tag) > 0 Then
                '标记图
                lKey = Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag)
                If lKey > 0 Then
                    cPicEditor.ShowPicEditor glngSys, gcnOracle, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey).OrigPic, _
                        lKey, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey).保留对象, Me, False
                    '这里外部图片的保存在cPicEditor对象的pOK事件中处理！
                End If
            End If
        Else
            If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal And (Me.Editor1.AuditMode = False) Then
                '查找关键字 ID ！
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then GoTo out
                If sKeyType = "P" Then
                    Call AddUndoPoint  '手动缓存
                    cPicEditor.ShowPicEditor glngSys, gcnOracle, Me.Document.Pictures("K" & lKey).OrigPic, lKey, Me.Document.Pictures("K" & lKey).保留对象, Me, False
                    Call ClearNoUseUndoList
                End If
            End If
        End If
    Case ID_PATISIGN
        Call PatiSign(Control)
    Case ID_SIGN, ID_SIGN_QUIT
        If AddSign Then
            Call RelateFeedback(True)
            If Control.ID = ID_SIGN_QUIT Then '签名并退出
                mblnPrecess = False
                Unload Me
                Exit Sub
            End If
        End If
        
        Call RecountPage
    Case ID_UNTREAD
        Call DoUntread
        txtContent.Enabled = True
        Call RelateFeedback(False)
        Call RecountPage
    Case ID_ELEMENT_UPDATE                  '更新要素

        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys = False Then GoTo out
        If sKeyType = "E" Then
        
            Call AddUndoPoint  '手动缓存
            With Me.Document.Elements("K" & lKey)
                If .替换域 = 1 Then
                                        
                    .内容文本 = GetReplaceEleValue(.要素名称, Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.主页ID, Me.Document.EPRPatiRecInfo.病人来源, Me.Document.EPRPatiRecInfo.医嘱id, Me.Document.EPRPatiRecInfo.婴儿)
                    .Refresh Me.Editor1
                    
                    If .自动转文本 Then
                        Me.Document.EleToString Me.Editor1, Me.Document.Elements("K" & lKey), False      '自动转化为纯文本（暂时不删除该要素）
                    End If
                End If
            End With
            Call ClearNoUseUndoList
        End If
        
        Call RecountPage
        
    Case ID_ELEMENT_CLEAR
    
        '清空要素
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys = False Then GoTo out
        If sKeyType = "E" Then
            Call AddUndoPoint  '手动缓存
            Me.Document.Elements("K" & lKey).内容文本 = ""
            Me.Document.Elements("K" & lKey).Refresh Me.Editor1
            Call ClearNoUseUndoList
        End If
        Call RecountPage
        
    Case ID_ELEMENT_TOSTRING
        '要素转化为纯文本
        If MsgBox("是否将该诊治要素转化为不含结构化信息的纯文本？", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys = False Then GoTo out
            If sKeyType = "E" Then
                Call AddUndoPoint  '手动缓存
                Dim str内容 As String
                str内容 = IIf(Me.Document.Elements("K" & lKey).内容文本 = "", "  ", Me.Document.Elements("K" & lKey).内容文本)
                lngLen = Len(str内容)
                With Me.Editor1
                    .Freeze
                    .ForceEdit = True
                    .Tag = "cbrThis_ExeCute"
                    .Range(lKSS, lKEE) = str内容
                    .Range(lKSS, lKSS + lngLen).Font.Protected = False
                    .Range(lKSS, lKSS + lngLen).Font.Hidden = False
                    .Range(lKSS, lKSS + lngLen).Font.BackColor = tomAutoColor
                    .Range(lKSS, lKSS + lngLen).Font.Underline = cprNone
                    .ForceEdit = False
                    .Tag = ""
                    .UnFreeze
                End With
                Me.Document.Elements.Remove "K" & lKey
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_REVISION_PREV
        Call GotoPrevRevision
    Case ID_REVISION_NEXT
        Call GotoNextRevision
    Case ID_REVISION_RESET
        Call ResetRevision
'        Me.Editor1.ResetAuditText
    Case ID_DIAGNOSIS
        '诊断
        Call AddUndoPoint  '手动缓存
        Call AddDiagnosis
        Call ClearNoUseUndoList
        Call RecountPage
    Case conMenu_Tool_Reference
        '诊断参考
        bFinded = IsBetweenKeys(Editor1, Editor1.Selection.StartPos + 1, "D", lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            Call Me.Document.Event_ClickDiagRef(Me.Document.Diagnosises("K" & lKey).诊断id, vbModal)
        End If
    Case ID_INSERT_TABLE
        Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.Selection.Font.Protected = False And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
        Call RecountPage
    Case ID_TABLE_CELLALIGNMENT1
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignLeft
                    tblThis.Cell(i, j).VAlignment = VALignTop
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT2
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignCentre
                    tblThis.Cell(i, j).VAlignment = VALignTop
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT3
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignRight
                    tblThis.Cell(i, j).VAlignment = VALignTop
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT4
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignLeft
                    tblThis.Cell(i, j).VAlignment = VALignVCentre
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT5
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignCentre
                    tblThis.Cell(i, j).VAlignment = VALignVCentre
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT6
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignRight
                    tblThis.Cell(i, j).VAlignment = VALignVCentre
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT7
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignLeft
                    tblThis.Cell(i, j).VAlignment = VALignBottom
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT8
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignCentre
                    tblThis.Cell(i, j).VAlignment = VALignBottom
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT9
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignRight
                    tblThis.Cell(i, j).VAlignment = VALignBottom
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CURRENCY
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FormatString = "￥#,0.00"
                    tblThis.Cell(i, j).HAlignment = HALignRight
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_PERCENT
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FormatString = "0.0%"
                    tblThis.Cell(i, j).HAlignment = HALignRight
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_KILOBIT
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FormatString = "#,0.00"
                    tblThis.Cell(i, j).HAlignment = HALignRight
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_MERGE
        If tblThis.Visible Then
            If Control.Checked Then
                tblThis.DisMergeCells tblThis.Row, tblThis.Col
            Else
                tblThis.MergeSelectedCells
            End If
            tblThis.Modified = True
            tblThis.Refresh False
            '调整UI界面
            tblThis_Resize tblThis.Width, tblThis.Height
            '保存背景图片
            If Val(tblThis.Tag) <= 0 Then GoTo out
            If tblThis.Modified Then SaveUIToTable Me.Document.Tables("K" & tblThis.Tag)
            Call RecountPage
        End If
    Case ID_TABLE_CELLPROTECTED
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.SelectedCellKey > 0 Then
            tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = Not tblThis.Cells("K" & tblThis.SelectedCellKey).Protected
        End If
        tblThis.Refresh False, False
    Case ID_TABLE_PROPERTY
        If tblThis.Visible Then
            tblThis.ShowProperty Me, tblThis
        End If
    Case ID_TABLE_INSERTCOLLEFT
        If tblThis.Visible Then
            If tblThis.Col > 0 And Val(tblThis.Tag) > 0 Then
                '如果包含合并单元格，那么不允许插入
                For i = 1 To tblThis.RowCount
                    If Len(tblThis.Cell(i, tblThis.Col).MergeInfo) > 0 Then
                        MsgBox "因为本列包含合并单元格，所以不允许插入其他列，请先取消合并后再试！", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '手动缓存
                '在表格控件中插入空白行
                Me.Document.Tables("K" & tblThis.Tag).InsertCol tblThis.Col - 1
                tblThis.InsertCol tblThis.Col - 1
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_INSERTCOLRIGHT
        If tblThis.Visible Then
            If tblThis.Col > 0 And Val(tblThis.Tag) > 0 Then
                '如果包含合并单元格，那么不允许插入
                For i = 1 To tblThis.RowCount
                    If Len(tblThis.Cell(i, tblThis.Col).MergeInfo) > 0 Then
                        MsgBox "因为本列包含合并单元格，所以不允许插入其他列，请先取消合并后再试！", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '手动缓存
                '在表格控件中插入空白行
                Me.Document.Tables("K" & tblThis.Tag).InsertCol tblThis.Col
                tblThis.InsertCol tblThis.Col
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_INSERTROWUP
        If tblThis.Visible Then
            If tblThis.Row > 0 And Val(tblThis.Tag) > 0 Then
                '如果包含合并单元格，那么不允许插入
                For i = 1 To tblThis.ColCount
                    If Len(tblThis.Cell(tblThis.Row, i).MergeInfo) > 0 Then
                        MsgBox "因为本行包含合并单元格，所以不允许插入其他行，请先取消合并后再试！", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '手动缓存
                '在表格控件中插入空白行
                Me.Document.Tables("K" & tblThis.Tag).InsertRow tblThis.Row - 1
                tblThis.InsertRow tblThis.Row - 1
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_INSERTROWDOWN
        If tblThis.Visible Then
            If tblThis.Row > 0 And Val(tblThis.Tag) > 0 Then
                '如果包含合并单元格，那么不允许插入
                For i = 1 To tblThis.ColCount
                    If Len(tblThis.Cell(tblThis.Row, i).MergeInfo) > 0 Then
                        MsgBox "因为本行包含合并单元格，所以不允许插入其他行，请先取消合并后再试！", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '手动缓存
                '在表格控件中插入空白行
                Me.Document.Tables("K" & tblThis.Tag).InsertRow tblThis.Row
                tblThis.InsertRow tblThis.Row
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_INSERTINHERITROW
        '插入继承行
        Dim lTag As Long
        If tblThis.Visible Then
            If tblThis.Row > 0 And Val(tblThis.Tag) > 0 Then
                '如果包含合并单元格，那么不允许插入
                For i = 1 To tblThis.ColCount
                    If Len(tblThis.Cell(tblThis.Row, i).MergeInfo) > 0 Then
                        MsgBox "因为本行包含合并单元格，所以不允许插入其他行，请先取消合并后再试！", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '手动缓存
                '先在表格控件中插入空白行
                lRow1 = tblThis.Row
                Me.Document.Tables("K" & tblThis.Tag).InsertRow lRow1
                tblThis.InsertRow lRow1
                '然后复制上一行内容
                For i = 1 To tblThis.ColCount
                    With tblThis.Cell(lRow1 + 1, i)
                        .Margin = tblThis.Cell(lRow1, i).Margin
                        .SingleLine = tblThis.Cell(lRow1, i).SingleLine
                        .Visibled = tblThis.Cell(lRow1, i).Visibled
                        .Width = tblThis.Cell(lRow1, i).Width
                        .Height = tblThis.Cell(lRow1, i).Height
                        .FixedWidth = tblThis.Cell(lRow1, i).FixedWidth
                        .AutoHeight = tblThis.Cell(lRow1, i).AutoHeight
                        .Icon = tblThis.Cell(lRow1, i).Icon
                        .Text = tblThis.Cell(lRow1, i).Text

'                        .Tag = tblThis.Cell(lRow1, i).Tag
                        If Val(tblThis.Cell(lRow1, i).Tag) > 0 Then
                            If tblThis.Cell(lRow1, i).Picture Is Nothing Then
                                '复制要素
                                lKey = Me.Document.Tables("K" & tblThis.Tag).Elements.AddExistNode(Me.Document.Tables("K" & tblThis.Tag).Elements("K" & tblThis.Cell(lRow1, i).Tag), False)
                                .Tag = lKey
                                Me.Document.Tables("K" & tblThis.Tag).Cell(lRow1 + 1, i).ElementKey = lKey
                            Else
                                '复制图片
                                lKey = Me.Document.Tables("K" & tblThis.Tag).Pictures.AddExistNode(Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & tblThis.Cell(lRow1, i).Tag), False)
                                .Tag = lKey
                                Me.Document.Tables("K" & tblThis.Tag).Cell(lRow1 + 1, i).PictureKey = lKey
                            End If
                        End If
                        .ToolTipText = tblThis.Cell(lRow1, i).ToolTipText
                        .FormatString = tblThis.Cell(lRow1, i).FormatString
                        .Indent = tblThis.Cell(lRow1, i).Indent
                        .HAlignment = tblThis.Cell(lRow1, i).HAlignment
                        .VAlignment = tblThis.Cell(lRow1, i).VAlignment
                        .ForeColor = tblThis.Cell(lRow1, i).ForeColor
                        .BackColor = tblThis.Cell(lRow1, i).BackColor
                        .GridLineColor = tblThis.Cell(lRow1, i).GridLineColor
                        .GridLineWidth = tblThis.Cell(lRow1, i).GridLineWidth
                        .FontName = tblThis.Cell(lRow1, i).FontName
                        .FontSize = tblThis.Cell(lRow1, i).FontSize
                        .FontBold = tblThis.Cell(lRow1, i).FontBold
                        .FontItalic = tblThis.Cell(lRow1, i).FontItalic
                        .FontStrikeout = tblThis.Cell(lRow1, i).FontStrikeout
                        .FontUnderline = tblThis.Cell(lRow1, i).FontUnderline
                        .FontWeight = tblThis.Cell(lRow1, i).FontWeight
                        .Protected = tblThis.Cell(lRow1, i).Protected
                        Set .Picture = tblThis.Cell(lRow1, i).Picture
                    End With
                Next
                tblThis.Refresh
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_FORMATCELL
        If tblThis.Visible Then
            tblThis.ShowProperty Me, tblThis, 3
        End If
    Case ID_TABLE_SAMECOLWIDTH
        '相同列宽
        Dim lSum As Long, lEvery As Long
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 Then
                If lCol1 <> lCol2 Then
                    For i = lCol1 To lCol2
                        lSum = lSum + tblThis.ColWidth(i)
                    Next
                    lEvery = lSum / (lCol2 - lCol1 + 1)
                    For i = lCol1 To lCol2
                        tblThis.ColWidth(i) = lEvery
                    Next
                    tblThis.Modified = True
                    tblThis.Refresh
                    tblThis_Resize tblThis.Width, tblThis.Height
                End If
            End If
        End If
    Case ID_TABLE_DELETEROW
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 And tblThis.RowCount > 1 And Val(tblThis.Tag) > 0 Then
                '如果包含合并单元格，那么不允许插入
                For i = 1 To tblThis.ColCount
                    If Len(tblThis.Cell(tblThis.Row, i).MergeInfo) > 0 Then
                        MsgBox "因为本行包含合并单元格，所以不允许删除，请先取消合并后再试！", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '手动缓存
                lRow1 = tblThis.Row
                Me.Document.Tables("K" & tblThis.Tag).DeleteRow lRow1
                tblThis.DeleteRow lRow1
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_DELETECOL
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 And tblThis.ColCount > 1 And Val(tblThis.Tag) > 0 Then
                '如果包含合并单元格，那么不允许插入
                For i = 1 To tblThis.RowCount
                    If Len(tblThis.Cell(i, tblThis.Col).MergeInfo) > 0 Then
                        MsgBox "因为本列包含合并单元格，所以不允许删除，请先取消合并后再试！", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '手动缓存
                lCol1 = tblThis.Col
                Me.Document.Tables("K" & tblThis.Tag).DeleteCol lCol1
                tblThis.DeleteCol lCol1
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_DELETETABLE
        If tblThis.Visible And Val(tblThis.Tag) > 0 Then
            Call AddUndoPoint  '手动缓存

            lKey = Val(tblThis.Tag)
            Me.Document.Tables("K" & lKey).DeleteFromEditor Me.Editor1
            Me.Document.Tables.Remove "K" & lKey
            Editor1.CloseUIInterface
            Editor1.Modified = True
            tblThis.Visible = False
            mblnChange = True

            Call ClearNoUseUndoList
            Call RecountPage
        End If
    Case Else
        If Control.ID >= conMenu_Tool_PlugIn_Item + 1 And Control.ID <= conMenu_Tool_PlugIn_Item + 99 And Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.ExecuteFunc(glngSys, 1070, Control.Parameter, Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.ID, Document.EPRFileInfo.ID)
            Err.Clear: On Error GoTo 0
        End If
    End Select
    
out: mblnPrecess = False
End Sub

'################################################################################################################
'## 功能：  上一处修订
'################################################################################################################
Private Sub GotoPrevRevision()
    On Error Resume Next
    Dim i As Long, lS As Long
    Dim lState1 As Long, lState2 As Long
    Dim lngStart As Long, lngEnd As Long
    Dim lng开始版1 As Long, lng终止版1 As Long
    Dim lng开始版2 As Long, lng终止版2 As Long

    With Me.Editor1
        .Freeze
        lS = .Selection.StartPos
        lState1 = Me.Document.GetTextState(Me.Editor1, lS, lS + 1, lng开始版1, lng终止版1) '获取当前文本状态
        If lS = 0 Then Exit Sub
        For i = lS - 1 To 0 Step -1   '循环查找不同状态文本
            lState2 = Me.Document.GetTextState(Me.Editor1, i, i + 1, lng开始版2, lng终止版2) '获取文本状态
            If lState2 <> lState1 Or lng开始版1 <> lng开始版2 Or lng终止版1 <> lng终止版2 Then
                If (lng开始版2 = Me.Document.目标版本 Or lng终止版2 = Me.Document.目标版本 - 1) Then
                    '状态不同了
                    lState1 = lState2
                    lngEnd = i + 1
                    Exit For
                Else
                    lState1 = lState2
                    lng开始版1 = lng开始版2
                    lng终止版1 = lng终止版2
                End If
            End If
        Next
        If lngEnd > 0 Then
            For i = lngEnd - 1 To 0 Step -1
                lState2 = Me.Document.GetTextState(Me.Editor1, i, i + 1, lng开始版2, lng终止版2) '获取文本状态
                If lState2 <> lState1 Or lng开始版1 <> lng开始版2 Or lng终止版1 <> lng终止版2 Then
                    '状态不同了
                    lState1 = lState2
                    lng开始版1 = lng开始版2
                    lng终止版1 = lng终止版2
                    lngStart = i + 1
                    Exit For
                End If
            Next
        End If
        If lngStart <> lngEnd Then .Range(lngStart, lngEnd).Selected
        .UnFreeze
    End With
End Sub

'################################################################################################################
'## 功能：  下一处修订
'################################################################################################################
Private Sub GotoNextRevision()
    On Error Resume Next
    Dim i As Long, lS As Long, lngLen As Long
    Dim lState1 As Long, lState2 As Long
    Dim lngStart As Long, lngEnd As Long
    Dim lng开始版1 As Long, lng终止版1 As Long
    Dim lng开始版2 As Long, lng终止版2 As Long

    With Me.Editor1
        .Freeze
        lS = .Selection.StartPos
        lngLen = Len(.Text)
        lState1 = Me.Document.GetTextState(Me.Editor1, lS, lS + 1, lng开始版1, lng终止版1) '获取当前文本状态
        If lS = 0 Then Exit Sub
        For i = lS To lngLen - 1
            lState2 = Me.Document.GetTextState(Me.Editor1, i, i + 1, lng开始版2, lng终止版2) '获取文本状态
            If lState2 <> lState1 Or lng开始版1 <> lng开始版2 Or lng终止版1 <> lng终止版2 Then
                If (lng开始版2 = Me.Document.目标版本 Or lng终止版2 = Me.Document.目标版本 - 1) Then
                    '状态不同了
                    lState1 = lState2
                    lng开始版1 = lng开始版2
                    lng终止版1 = lng终止版2
                    lngStart = i
                    Exit For
                Else
                    lState1 = lState2
                    lng开始版1 = lng开始版2
                    lng终止版1 = lng终止版2
                End If
            End If
        Next
        If lngStart < lngLen Then
            For i = lngStart + 1 To lngLen - 1
                lState2 = Me.Document.GetTextState(Me.Editor1, i, i + 1, lng开始版2, lng终止版2) '获取文本状态
                If lState2 <> lState1 Or lng开始版1 <> lng开始版2 Or lng终止版1 <> lng终止版2 Then
                    '状态不同了
                    lState1 = lState2
                    lng开始版1 = lng开始版2
                    lng终止版1 = lng终止版2
                    lngEnd = i
                    Exit For
                End If
            Next
        End If
        If lngStart <> lngEnd Then .Range(lngStart, lngEnd).Selected
        .UnFreeze
    End With
End Sub

'################################################################################################################
'## 功能：  获取用于签名的源文本（除去“书写签名”预制提纲的其他所有文本内容）
'################################################################################################################
Public Function GetSignSourceString(ByRef edtThis As zlRichEditor.Editor) As String
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean, lKey As Long
    Dim i As Long, strR As String, lS As Long, lE As Long, strS As String, lngLen As Long, lPos As Long
    
    edtThis.SaveDoc App.Path & "\tmp.RTF"
    gfrmPublic.edtBuff.OpenDoc App.Path & "\tmp.RTF"
    gobjFSO.DeleteFile App.Path & "\tmp.RTF"
    gfrmPublic.edtBuff.Freeze
    gfrmPublic.edtBuff.ForceEdit = True
    '去掉所有S关键字的签名对象
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "S", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSS, lEE).Text = ""
    Loop Until bFinded = False
    gfrmPublic.edtBuff.ForceEdit = False
    gfrmPublic.edtBuff.UnFreeze
    strS = gfrmPublic.edtBuff.Text
    strS = Replace(strS, Chr(32), "")
    strS = Replace(strS, vbCr, "")
    strS = Replace(strS, vbLf, "")
    
'   '此种方法错误，如果多次签名位置混乱就会无法验证-----------此段以下文字需保留
'    lngLen = Len(edtThis.Text)
'    If Me.Document.Signs.Count = 0 Then
'        strS = edtThis.Text
'    Else
'        For i = 1 To Me.Document.Signs.Count
'            bFinded = FindKey(edtThis, "S", Me.Document.Signs(i).Key, lSS, lSE, lES, lEE, bNeeded)
'            If bFinded Then
'                '剔除签名
'                If i = 1 Then
'                    strS = edtThis.Range(0, lSS)
'                Else
'                    strS = strS & edtThis.Range(lS, lSS)
'                End If
'                lS = lEE
'            End If
'        Next
'        If lEE < lngLen Then
'            '加入末端文本
'            strS = strS & edtThis.Range(lEE, lngLen).Text
'        End If
'    End If
    GetSignSourceString = strS
End Function
Private Sub PatiSign(Control As XtremeCommandBars.ICommandBarControl)
    If Control.Caption = "患者签名" Then
        Call PatiDoSign(Control)
    Else
        Call PatiUnDoSign(Control)
    End If
End Sub
Private Sub PatiUnDoSign(Control As XtremeCommandBars.ICommandBarControl)
'功能：撤消手写签名图片
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, blnNeeded As Boolean, lCurPos As Long, intLoop As Integer
    On Error GoTo errHand
    
    If mblnPatiSign Then
        lCurPos = 1
        Do Until False
            If FindNextKey(Editor1, lCurPos, "P", lKey, lKSS, lKSE, lKES, lKEE, blnNeeded) Then
                If Document.Pictures("K" & lKey).PictureType = EPRPatiSign Then '找到图片，检查是否是手签图
                    Exit Do
                End If
            Else '没找到任何图
                For intLoop = 1 To Document.Pictures.Count
                    If Document.Pictures(intLoop).PictureType = EPRPatiSign Then
                        Document.Pictures.Remove (intLoop)
                        Exit Do
                    End If
                Next

                GoTo undosign
            End If
            lCurPos = lKEE
        Loop
        Call Document.Pictures("K" & lKey).DeleteFromEditor(Editor1)
    End If

undosign:
    mblnPatiSign = False
    Call RecountPage
    Control.Caption = "患者签名"
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function HavedPatiSign() As Boolean
'功能：检查是否已经手签过
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, blnNeeded As Boolean, lCurPos As Long
    On Error GoTo errHand
    
    lCurPos = 1
    Do Until False
        If FindNextKey(Editor1, lCurPos, "P", lKey, lKSS, lKSE, lKES, lKEE, blnNeeded) Then
            If Document.Pictures("K" & lKey).PictureType = EPRPatiSign Then '找到图片，检查是否是手签图
                HavedPatiSign = True
                Exit Function
            End If
        Else '没找到任何图
            Exit Function
        End If
        lCurPos = lKEE
    Loop

    HavedPatiSign = False
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub PatiDoSign(Control As XtremeCommandBars.ICommandBarControl)
'功能：获取手写签名图片、签名认证返回信息
'参数：strSource 签名源文
'        strName 病人姓名，缺省签名人
'        strIdentifyNo 病人身份证号，缺省签名人证件号
'        strOtherParms  为日后不变更参数数量，保留其它可能的参数
'        strSignInfo 签名返回认证信息
'        strPenSignBase64 返回手写签名图片BASE64编码
'        objPenSignPic 返回手写签名图片
Dim strSource As String, strName As String, strIdentifyNo As String, strOtherParms As String, strSignInfo As String, strPenSignBase64 As String, objPenSignPic As Object
Dim blnReturn As Boolean, lWidth As Long, lHeight As Long
    On Error GoTo errHand
    strSource = GetSignSourceString(Me.Editor1)
    strName = mPatiInfor.姓名
    strIdentifyNo = mPatiInfor.身份证号
    strOtherParms = "" '
    blnReturn = gobjESign.PenSignature(strSource, strName, strIdentifyNo, strOtherParms, strSignInfo, strPenSignBase64, objPenSignPic)
    If blnReturn And Not objPenSignPic Is Nothing Then
        '插入图片
        Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
        If Editor1.ReadOnly Then Exit Sub
        If tblThis.Visible Then
            MsgBox "当前位置不支持病人签名", vbInformation, gstrSysName: Exit Sub
        Else
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys = False Then
                Call AddUndoPoint  '手动缓存
                Editor1.Tag = "InsertPatiSign"
                '将图片对象保存到类对象中
                lKey = Document.Pictures.Add()
                Set Document.Pictures("K" & lKey).OrigPic = objPenSignPic
                Document.Pictures("K" & lKey).Width = lWidth
                Document.Pictures("K" & lKey).Height = lHeight
                Document.Pictures("K" & lKey).OrigWidth = lWidth
                Document.Pictures("K" & lKey).OrigHeight = lHeight
                Document.Pictures("K" & lKey).PictureType = EPRPatiSign
                Document.Pictures("K" & lKey).InsertIntoEditor Editor1
                Document.Pictures("K" & lKey).内容文本 = strName & "|" & strIdentifyNo & "|" & strSignInfo   '保存信息
                Editor1.Tag = ""
                Call ClearNoUseUndoList
            End If
            Editor1.SetFocus
        End If
        mblnPatiSign = True
        Call RecountPage
        Control.Caption = "患者撤签"
        '界面控制
        '无法控制,如果控制了，签名都无法正常进行
    End If
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
'################################################################################################################
'## 功能：  新增签名
'################################################################################################################
Private Function AddSign() As Boolean
    If Me.Editor1.ViewMode <> cprNormal Then Exit Function
    Dim strTmp As String, lLen As Long, lngKey As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim strSource As String, blnR As Boolean, picSign As StdPicture, lngPicKey As Long
    Dim frmSign As New frmEPRSign, oSign As cEPRSign
    Dim blnModified As Boolean '是否在签名前已经修改了内容
    Dim lS As Long, l As Long
    Dim strSQL As String, strTime As String
    
    If AutoMoveSignPos = False Then Exit Function
    If Me.Editor1.Selection.Font.Protected Then
        MsgBox "对不起，您不能在当前位置进行签名！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Me.Document.用户签名级别 = cprSL_空白 Then
        MsgBox "当前用户尚未设置签名级别，请在人员管理中调整聘任职务！", vbInformation, gstrSysName: Exit Function
    End If
    For l = 1 To Document.Signs.Count
        If Document.Signs(l).签名级别 > Me.Document.用户签名级别 Then
            MsgBox "当前病历已有更高级别的签名！当前签名级别无权审签本病历", vbInformation, gstrSysName
            Exit Function
        End If
    Next
    
    If Not CheckAllObjects(True) Then Exit Function '检查必填要素
    
            
    If Not gobjPlugIn Is Nothing Then '签名前插件处理
        On Error Resume Next
        If Not gobjPlugIn.SignEMRBefore(glngSys, 1070, Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.ID) Then Exit Function
        Err.Clear: On Error GoTo 0
    End If

    With Editor1
        '如果没有本次版本的签名位置，则在当前位置插入签名
        .Tag = "签名"
        blnModified = .Modified
        lS = .Selection.StartPos
        
        If .SelLength > 0 And mbln签名要素 = False Then .Tag = "": Exit Function
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        
        If Editor1.Selection.StartPos >= lKSS And Editor1.Selection.StartPos <= lKSE Then
            Editor1.SelStart = IIf(lKSS = 0, 0, lKSS - 1)
'            Editor1.Selection.StartPos = lKSS - 1
        End If
        If .Selection.Font.Protected And mbln签名要素 = False Then
            .Tag = ""
            Exit Function
        End If
        
        If bBeteenKeys And mbln签名要素 = False Then
            .Tag = ""
            Exit Function    '保证不能插入关键字内部
        Else
            strSource = GetSignSourceString(Me.Editor1)
            Set oSign = frmSign.ShowMe(Me.Editor1, Me, strSource, picSign)
            If Not oSign Is Nothing Then
                Me.Editor1.Modified = blnModified
                If (Me.Editor1.Modified Or (Me.Editor1.AuditMode And Me.Document.EPRPatiRecInfo.签名级别 = cprSL_空白)) Or Me.Editor1.AuditMode = False Then
                    oSign.开始版 = Me.Document.目标版本
                Else
                    oSign.开始版 = Me.Document.目标版本 - 1
                End If
                If oSign.开始版 > 16 Then
                    MsgBox "目前系统支持的最大版本号为16，请回退或者重新整理！", vbOKOnly + vbInformation, gstrSysName
                    .Tag = ""
                    Exit Function
                End If
                lngKey = Me.Document.Signs.AddExistNode(oSign)
                
                If Me.Document.Signs("K" & lngKey).InsertIntoEditor(Me.Editor1, , True, Me.Document) = True Then
                    If oSign.签名图片 And Not picSign Is Nothing Then
                        lngPicKey = Document.Pictures.Add()
                        Set Document.Pictures("K" & lngPicKey).OrigPic = picSign
                        Document.Pictures("K" & lngPicKey).Width = Me.ScaleX(picSign.Width, vbHimetric, vbTwips)
                        Document.Pictures("K" & lngPicKey).Height = Me.ScaleY(picSign.Height, vbHimetric, vbTwips)
                        Document.Pictures("K" & lngPicKey).OrigWidth = Me.ScaleX(picSign.Width, vbHimetric, vbTwips)
                        Document.Pictures("K" & lngPicKey).OrigHeight = Me.ScaleY(picSign.Height, vbHimetric, vbTwips)
                        Document.Pictures("K" & lngPicKey).PictureType = EPRSignPicture
                        Document.Pictures("K" & lngPicKey).开始版 = oSign.开始版
                        Document.Pictures("K" & lngPicKey).InsertIntoEditor Editor1
                        Call FindKey(Editor1, "S", lngKey, lKSS, lKSE, lKES, lKEE, bNeeded)
                        If Editor1.ForceEdit = False Then Editor1.ForceEdit = True
                        Editor1.Range(lKSS, lKEE).Font.Hidden = True
                    End If
                    If oSign.签名方式 <> 2 Then Call AutoAlterSignInPage '调整签名位置 数字签名不允许调整签名位置
                    Me.Editor1.Modified = blnModified
                    blnR = Me.Document.SaveEPRDoc(Editor1, InStr(1, gstrPrivsEpr, "修改页面设置") > 0)
                    
    '                If oSign.签名方式 = 2 And oSign.签名规则 = 3 And blnR Then
    '                    Dim strSign As String, lngCertID As Long, str时间戳 As String
    '                    '在签名窗口，签名对象已进行初始化，以避免初始化失败但本地已保存情况，所以直接使用即可
    '                    strSource = GetSignSourceFromDB(Me.Document.EPRPatiRecInfo.ID, lngKey)
    '                    On Error Resume Next
    '                    strSign = gobjESign.signature(strSource, UCase(oSign.签名信息), lngCertID, str时间戳) '返回签名信息,lngCertID返回签名使用的证书记录ID
    '                    If strSign <> "" Then
    '                        oSign.签名信息 = strSign    '要素值域
    '                        oSign.证书ID = lngCertID    '对象属性
    '                        oSign.时间戳 = str时间戳    '要素单位
    '                        gstrSQL = "zl_电子病历内容_数字签名(" & Me.Document.EPRPatiRecInfo.ID & "," & lngKey & ",'" & oSign.对象属性 & "','" & oSign.签名信息 & "','" & oSign.时间戳 & "')  '调用过程保存"
    '                        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存签名信息")
    '                    ElseIf strSign = "" Or Err.Number <> 0 Then '签名失败需要回退
    '                        Call DoUntread(True, lngKey)
    '                        MsgBox "数字签名失败，本次签名已自动回退,请重新签名！", vbCritical, gstrSysName
    '                    End If
    '                End If

                    Call ClearUndoList      '清空Undo队列！
                    DT1_EPR = Now
                    Me.Editor1.Modified = False
                    AddSign = blnR
                End If
                
                If Not gobjPlugIn Is Nothing Then '签名后插件处理
                    On Error Resume Next
                    Call gobjPlugIn.SignEMRAfter(glngSys, 1070, Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.ID, oSign.姓名)
                    Err.Clear: On Error GoTo 0
                End If
            End If
        End If
    End With
    If mbln返修处理 And mblnFBContentChanged Then
        strTime = "to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
        strSQL = "Zl_疾病申报记录_Update(" & Me.Document.EPRPatiRecInfo.ID & ",5,null,null,null,'" & gstrUserName & "'," & strTime & ",'" & Trim(txtContent.Text) & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        mblnFBContentChanged = False
    End If
    If mblnIsMultiMode And Not mfrmMultiDocView Is Nothing Then
        mfrmMultiDocView.InitData Me, Me.Document, Me.Document.EPRPatiRecInfo.ID
    End If
    Call SetStateInfo
    Editor1.Tag = ""
End Function
Private Sub AutoAlterSignInPage()
'针对诊疗报告,自动调整签名在页面中位置
'查到的第一个签名前追加回车换行符，将签名及其后内容位移到页底直至分页（只在一页内显示报告）
    If Not mblnSignAutoAlter Then Exit Sub
    If Document.EPRFileInfo.种类 <> cpr诊疗报告 Then Exit Sub
    
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim blnForce As Boolean, l As Long
    
    With Editor1
        .Freeze
        blnForce = .ForceEdit
        .ForceEdit = True
    
        Call Editor1.DoVirtualPrint
        If Editor1.PageCount > 1 Then Exit Sub '内容已经超过一页不处理
    
        If Not FindNextKey(Editor1, 1, "S", lKey, lKSS, lKSE, lKES, lKEE, bNeeded) Then Exit Sub '没找到签名对象不处理
        .Range(lKSS, lKSS).Font.Protected = False
        For l = 1 To 40
            .Range(lKSS, lKSS).Text = vbCrLf
            .Range(lKSS, lKSS + 2).Font.Protected = False
            Call Editor1.DoVirtualPrint
            If Editor1.PageCount > 1 Then '超过一页，取消刚才追加的三个回车换行
                If .Range(lKSS, lKSS + 2) = vbCrLf Then .Range(lKSS, lKSS + 2).Text = ""
                lKSS = lKSS - 2
                If .Range(lKSS, lKSS + 2) = vbCrLf Then .Range(lKSS, lKSS + 2).Text = ""
                lKSS = lKSS - 2
                If .Range(lKSS, lKSS + 2) = vbCrLf Then .Range(lKSS, lKSS + 2).Text = ""
                Exit For
            Else
                lKSS = lKSS + 2
            End If
        Next
        
        .ForceEdit = blnForce
        .UnFreeze
    End With
    
End Sub
'################################################################################################################
'## 功能：  新增诊断
'################################################################################################################
Private Function AddDiagnosis() As Boolean
    Dim strTmp As String, lLen As Long, lngKey As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim strSource As String, blnR As Boolean
    Dim frmDiagnosis As New frmInsDiagnosis, oDiagnosis As cEPRDiagnosis

    With Editor1
        '在当前位置插入诊断
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then
            AddDiagnosis = False: Exit Function    '保证不能插入关键字内部
        Else
            Set oDiagnosis = frmDiagnosis.ShowMe(Me.Editor1, Me)
            If Not oDiagnosis Is Nothing Then
                oDiagnosis.开始版 = Me.Document.目标版本
                lngKey = Me.Document.Diagnosises.AddExistNode(oDiagnosis)
                Me.Document.Diagnosises("K" & lngKey).InsertIntoEditor Me.Editor1
            End If
            AddDiagnosis = True
        End If
    End With
End Function

'################################################################################################################
'## 功能：  取消当前选中内容的修订
'################################################################################################################
Private Sub ResetRevision()
    '恢复所选文本修订内容
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, lG As Long, COLOR As OLE_COLOR
    With Me.Editor1
        .Tag = "ResetRevision"
        .InProcessing = True
        .Freeze
        .ForceEdit = True
        lS = .Selection.StartPos
        lE = .Selection.EndPos

        '先处理要素和诊断
        For i = lS To lE
            bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bFinded Then
                If lKSS < lE Then
                    '范围内存在关键字
                    If sKeyType = "E" Then
                        '要素
                        If Me.Document.Elements("K" & lKey).开始版 = Me.Document.目标版本 Then
                            '本版新增的要素，那么删除之！
                            .Range(lKSS, lKEE).Text = ""
                            Me.Document.Elements.Remove "K" & lKey
                            lE = lE - (lKEE - lKSS)
                            i = i - 1
                        ElseIf Me.Document.Elements("K" & lKey).终止版 = Me.Document.目标版本 - 1 Then
                            '本版删除的要素，那么恢复之！
                            Me.Document.Elements("K" & lKey).终止版 = 0
                            Me.Document.Elements("K" & lKey).Refresh Me.Editor1
                            i = lKEE - 1
                        Else
                            '不处理
                            i = lKEE - 1
                        End If
                    ElseIf sKeyType = "D" Then
                        '诊断
                        If Me.Document.Diagnosises("K" & lKey).开始版 = Me.Document.目标版本 Then
                            '本版新增的要素，那么删除之！
                            .Range(lKSS, lKEE).Text = ""
                            Me.Document.Diagnosises.Remove "K" & lKey
                            lE = lE - (lKEE - lKSS)
                            i = i - 1
                        ElseIf Me.Document.Diagnosises("K" & lKey).终止版 = Me.Document.目标版本 - 1 Then
                            '本版删除的要素，那么恢复之！
                            Me.Document.Diagnosises("K" & lKey).终止版 = 0
                            Me.Document.Diagnosises("K" & lKey).Refresh Me.Editor1
                            i = lKEE - 1
                        Else
                            '不处理
                            i = lKEE - 1
                        End If
                    Else
                        '如果是其他元素，不处理
                        i = lKEE - 1
                    End If
                Else
                    '否则，超出范围，退出循环
                    Exit For
                End If
            Else
                '不存在任何元素，那么退出循环
                Exit For
            End If
        Next

        i = lS
        Do While i < lE
            If .Range(i, i + 1).Font.Protected = False And .Range(i, i + 1).Font.Hidden = False Then
                COLOR = IIf(.Range(i, i + 1).Font.ForeColor = tomAutoColor, vbBlack, .Range(i, i + 1).Font.ForeColor)
                If Me.Document.IsNewCharColor(COLOR) And .Range(i, i + 1).Font.Strikethrough = False Then
                    '下一个字符为新增文本，则直接删除之！
                    .Range(i, i + 1) = ""
                    lE = lE - 1
                ElseIf Me.Document.IsDelCharColor(COLOR) And .Range(i, i + 1).Font.Strikethrough = True Then
                    '下一个字符为删除文本，则恢复文本为（无删除线＋删除前的颜色）。
                    lG = rgbGreen(COLOR)
                    If lG <> 0 Then
                        '表示该文本在删除前是新增文本，那么应该恢复为新增状态
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.ForeColor = RGB(255, lG, 0)
                    Else
                        '否则恢复为黑色
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.ForeColor = tomAutoColor
                    End If
                    i = i + 1
                Else
                    i = i + 1
                End If
            Else
                '若为保护/隐藏文本，则直接后移一位。
                i = i + 1
            End If
        Loop
        .InProcessing = False
        .Range(i, i).Selected
        .UnFreeze
        .Tag = ""
    End With
End Sub

'################################################################################################################
'## 功能：  取消签名（回退操作）
'################################################################################################################
Private Sub DoUntread(Optional blnImmediate As Boolean = False, Optional lSignKey As Long)
    If Me.Editor1.ViewMode <> cprNormal Then Exit Sub
    Dim lngVersion As Long, lngSignKey As Long
    Dim frmUntread As New frmEPRUntread
    Dim i As Long, lngKey As Long, lngLen As Long, COLOR As OLE_COLOR
    Dim blnForce As Boolean, lng开始版 As Long
    Dim lngKeys() As Long, lngCount As Long, blnReadOnly As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    If blnImmediate Then '由数字签名失败时调用
        lngSignKey = lSignKey
    Else
        If frmUntread.ShowMe(Me.Document.EPRPatiRecInfo.ID, Me.Document.EditType, lngVersion, lngSignKey, Me) = False Then Exit Sub
        If lngSignKey > 0 Or lngVersion > 0 Then
            If MsgBox("注意：回退操作将不可恢复！是否继续？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    '进行回退处理并重新保存
    On Error GoTo errHand
    If lngSignKey > 0 Then
        If Me.Document.Signs("K" & lngSignKey).签名方式 = 2 And Not blnImmediate Then
            '数字签名验证
            If gobjESign Is Nothing Then
                Set gobjESign = CreateObject("zl9ESign.clsESign")
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
            If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        End If
    
        '清除签名
        blnReadOnly = Me.Editor1.ReadOnly
        Me.Editor1.ReadOnly = False
        Editor1.Tag = "DoUntread"
        If Me.Document.Signs("K" & lngSignKey).签名图片 Then
            '清除签名图片,因签名图片与签名无直接数据联系,签名版本在没更改的情况下不会增长，所以只能用签名后第一张签名图清除
            If FindKey(Editor1, "S", Me.Document.Signs("K" & lngSignKey).Key, lKSS, lKSE, lKES, lKEE, bNeeded) Then
                If FindNextKey(Editor1, lKSS, "P", lKey, lKSS, lKSE, lKES, lKEE, bNeeded) Then
                    If Me.Document.Pictures("K" & lKey).PictureType = EPRSignPicture Then
                        Me.Document.Pictures("K" & lKey).DeleteFromEditor Me.Editor1
                        Me.Document.Pictures.Remove "K" & lKey
                    End If
                End If
            End If
        End If
        
        Me.Document.Signs("K" & lngSignKey).DeleteFromEditor Me.Editor1, Me.Document
        Me.Document.Signs.Remove "K" & lngSignKey
        Editor1.Tag = ""
        Me.Editor1.ReadOnly = blnReadOnly
    ElseIf lngVersion > 0 Then
        Editor1.Tag = "DoUntread"
        Editor1.InProcessing = True
        
        '清除表格中当前版本单元，重新获取
        Dim oTable As cEPRTable
        For Each oTable In Me.Document.Tables
            If oTable.TableType = tte_默认 Then Call oTable.ReGetCellsFromDB(lngVersion)
        Next
    
        '清除诊治要素
        ReDim Preserve lngKeys(0 To 0) As Long
        For i = 1 To Me.Document.Elements.Count
            If Me.Document.Elements(i).开始版 >= lngVersion And lngVersion > 1 Then
                lngCount = UBound(lngKeys) + 1
                ReDim Preserve lngKeys(0 To lngCount) As Long
                lngKeys(lngCount) = Me.Document.Elements(i).Key
            End If
        Next
        For i = 1 To UBound(lngKeys)
            lngKey = lngKeys(i)
            Me.Document.Elements("K" & lngKey).DeleteFromEditor Me.Editor1
            Me.Document.Elements.Remove "K" & lngKey
        Next
        '恢复删除的诊治要素
        ReDim Preserve lngKeys(0 To 0) As Long
        For i = 1 To Me.Document.Elements.Count
            If Me.Document.Elements(i).终止版 >= lngVersion - 1 And lngVersion > 1 Then
                lngCount = UBound(lngKeys) + 1
                ReDim Preserve lngKeys(0 To lngCount) As Long
                lngKeys(lngCount) = Me.Document.Elements(i).Key
            End If
        Next
        For i = 1 To UBound(lngKeys)
            lngKey = lngKeys(i)
            Me.Document.Elements("K" & lngKey).终止版 = 0
        Next
        '清除诊断
        ReDim Preserve lngKeys(0 To 0) As Long
        For i = 1 To Me.Document.Diagnosises.Count
            If Me.Document.Diagnosises(i).开始版 >= lngVersion And lngVersion > 1 Then
                lngCount = UBound(lngKeys) + 1
                ReDim Preserve lngKeys(0 To lngCount) As Long
                lngKeys(lngCount) = Me.Document.Diagnosises(i).Key
            End If
        Next
        For i = 1 To UBound(lngKeys)
            lngKey = lngKeys(i)
            Me.Document.Diagnosises("K" & lngKey).DeleteFromEditor Me.Editor1
            Me.Document.Diagnosises.Remove "K" & lngKey
        Next
        '恢复删除的诊断
        ReDim Preserve lngKeys(0 To 0) As Long
        For i = 1 To Me.Document.Diagnosises.Count
            If Me.Document.Diagnosises(i).终止版 >= lngVersion - 1 And lngVersion > 1 Then
                lngCount = UBound(lngKeys) + 1
                ReDim Preserve lngKeys(0 To lngCount) As Long
                lngKeys(lngCount) = Me.Document.Diagnosises(i).Key
            End If
        Next
        For i = 1 To UBound(lngKeys)
            lngKey = lngKeys(i)
            Me.Document.Diagnosises("K" & lngKey).终止版 = 0
        Next
        '清除文本
        With Me.Editor1
            lngLen = Len(.Text)
            blnForce = .ForceEdit
            .Tag = "DoUntread"
            .Freeze
            .ForceEdit = True
            For i = 0 To lngLen - 1
                '判断 .Range(i, i + 1).Font.ForeColor 颜色值，用于确定文本版本
                COLOR = .Range(i, i + 1).Font.ForeColor
                If Me.Document.IsNewCharColor(COLOR) Then
                    '属于新增文本，那么清除该文本
                    If .Range(i, i + 1).Font.Hidden And .Range(i, i + 3).Text = "TS(" Then
                        i = i + InStr(1, .Range(i, i + 100), ")")
                    Else
                        .Range(i, i + 1) = ""
                        lngLen = lngLen - 1
                        i = i - 1
                    End If
                ElseIf Me.Document.IsDelCharColor(COLOR) Then
                    '属于删除文本，那么还原该文本
                    lng开始版 = Get开始版(COLOR)
                    .Range(i, i + 1).Font.ForeColor = GetCharColor(lng开始版, 0)
                    .Range(i, i + 1).Font.Strikethrough = False
                    lngLen = lngLen - 1
                End If
            Next
            .ForceEdit = blnForce
            .UnFreeze
            .Tag = ""
        End With
        Editor1.Tag = ""
        Editor1.InProcessing = False
    End If
    '保存文件
    Me.Editor1.Modified = False
    If Me.Document.SaveEPRDoc(Me.Editor1, InStr(1, gstrPrivsEpr, "修改页面设置") > 0) Then
        Call ClearUndoList      '清空Undo队列！
        DT1_EPR = Now
'        MsgBox "回退操作成功！", vbOKOnly + vbInformation, gstrSysName
        If mblnIsMultiMode And Not mfrmMultiDocView Is Nothing Then
            mfrmMultiDocView.InitData Me, Me.Document, Me.Document.EPRPatiRecInfo.ID
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'################################################################################################################
'## 功能：  显示自动识别诊治要素或者字典项目的选择器
'##
'## 参数：  strAuto     :IN     传入查询关键字
'################################################################################################################
Private Sub ShowAutoRecSelector(ByVal strF As String)
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys And tblThis.Visible = False Then Exit Sub    '保证不能插入关键字内部
    If Me.Editor1.Selection.Font.Protected And tblThis.Visible = False Then Exit Sub

    Dim rs As New ADODB.Recordset
    Dim lLeft As Long, lTOp As Long, lRight As Long, lBottom As Long

    '如果中文名或者英文名建了索引会更快一些！
    gstrSQL = "select  ID,编码,中文名 As 名称,单位,decode(替换域,2,'字典项目',1,'替换项目','外部输入项') As 类型 " & _
        "From 诊治所见项目 " & _
        "Where 中文名 Like '%" & strF & "%' Or 英文名 Like '%" & UCase(strF) & "%' " & _
        "Order By 类型"

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息")
    If rs.EOF Then Exit Sub
    Dim pt As POINTAPI, arrPara As String, T As Variant, lngId As Long
    Dim f As New frmSelectChild

    pt.X = 0
    pt.y = 0
    ClientToScreen Editor1.OriginRTB.hwnd, pt
    '获取起始位置坐标
    If tblThis.Visible Then
        If tblThis.SelectedCellKey > 0 Then
            tblThis.Cells("K" & tblThis.SelectedCellKey).GetCellTextBorder lLeft, lTOp, lRight, lBottom
            lLeft = Me.Left + Editor1.Left + Editor1.UILeft + tblThis.Left + lLeft * 15 + 30
            lTOp = Me.Top + Editor1.Top + Editor1.UITop + tblThis.Top + lBottom * 15 + 500
            arrPara = "0;830;2500;700;1000"
            strF = f.ShowSelectChild(Me, lLeft, lTOp, 5550, 3000, rs, arrPara)
        Else
            Exit Sub
        End If
    Else
        Editor1.Range(Editor1.Selection.StartPos, Editor1.Selection.StartPos + 1).GetPoint cprGPStart + cprGPLeft + cprGPBottom, lLeft, lTOp
        Call AddUndoPoint  '手动缓存
        arrPara = "0;830;2500;700;1000"
        strF = f.ShowSelectChild(Me, pt.X * Screen.TwipsPerPixelX + lLeft, pt.y * Screen.TwipsPerPixelY + lTOp, 5550, 3000, rs, arrPara)
    End If


    If strF = "" Then
        Exit Sub
    Else
        T = Split(strF, ";")
        lngId = T(0)
        rs.Close
        gstrSQL = "Select ID, 中文名, 类型, 长度, 小数, 单位, 表示法, 替换域, 初始值, 数值域 From 诊治所见项目 Where ID =[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", lngId)
        If Not rs.EOF Then
            '插入元素
            Dim Ele As New cEPRElement, aryTemp() As String, lngKey As Long, lngCount As Long
            With Ele
                .要素名称 = NVL(rs("中文名"))
                .诊治要素ID = NVL(rs("ID"), 0)
                .要素类型 = NVL(rs("类型"), 1)
                .要素长度 = NVL(rs("长度"), 0)
                .要素小数 = NVL(rs("小数"), 0)
                .要素单位 = NVL(rs("单位"))
                .要素表示 = IIf(NVL(rs("表示法"), 0) = 4, 2, NVL(rs("表示法"), 0))
                .替换域 = NVL(rs("替换域"), 0)      '0-外部输入项目；1-替换项目；2-字典项目
                .内容文本 = Trim(NVL(rs("初始值")))
                If .要素类型 = 0 Then
                    Select Case .要素表示
                    Case 0, 1
                        If Trim(NVL(rs("数值域"))) = "" Then
                            .要素值域 = ""
                        Else
                            aryTemp = Split(NVL(rs("数值域")), ";")
                            .要素值域 = Val(aryTemp(0)) & ";" & Val(aryTemp(1))
                        End If
                    Case 2
                        aryTemp = Split(NVL(rs("数值域")), ";")
                        For lngCount = 0 To UBound(aryTemp)
                            aryTemp(lngCount) = Val(aryTemp(lngCount))
                        Next
                        .要素值域 = Join(aryTemp(0), ";")
                    Case Else
                        .要素值域 = ""
                    End Select
                Else
                    Select Case .要素表示
                    Case 2, 3
                        .要素值域 = NVL(rs("数值域"))
                    Case Else
                        .要素值域 = ""
                    End Select
                End If
                .输入形态 = IIf(.要素表示 = 2 Or .要素表示 = 3, 1, 0) '0-文本 1-上下 2-单选 3-复选   如果为单选、复选，则这里默认值为展开项目   0-弹出;1-展开
            End With
            If tblThis.Visible Then
                If Val(tblThis.Tag) > 0 Then
                    lngKey = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements.AddExistNode(Ele)
                    If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                    Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).内容文本 = GetReplaceEleValue(Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).要素名称, _
                        Me.Document.EPRPatiRecInfo.病人ID, _
                        Me.Document.EPRPatiRecInfo.主页ID, _
                        Me.Document.EPRPatiRecInfo.病人来源, _
                        Me.Document.EPRPatiRecInfo.医嘱id, _
                        Me.Document.EPRPatiRecInfo.婴儿)
                    End If
                    '保存到单元格中
                    tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).内容文本
                    tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lngKey
                    tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).要素名称
                    tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
                    tblThis.Modified = True
                    tblThis.Refresh False, True, tblThis.SelectedCellKey
                    tblThis_Resize tblThis.Width, tblThis.Height
                End If
            Else
                lngKey = Me.Document.Elements.AddExistNode(Ele)
                '替换项目
                If Me.Document.Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                    Me.Document.Elements("K" & lngKey).内容文本 = GetReplaceEleValue(Me.Document.Elements("K" & lngKey).要素名称, _
                        Me.Document.EPRPatiRecInfo.病人ID, _
                        Me.Document.EPRPatiRecInfo.主页ID, _
                        Me.Document.EPRPatiRecInfo.病人来源, _
                        Me.Document.EPRPatiRecInfo.医嘱id, _
                        Me.Document.EPRPatiRecInfo.婴儿)
                End If
                Me.Document.Elements("K" & lngKey).开始版 = Me.Document.目标版本

                '插入诊治要素到编辑器中
                Dim blnForce As Boolean
                blnForce = Me.Editor1.ForceEdit
                Me.Editor1.ForceEdit = True
                Me.Editor1.Tag = "ShowAutoRecSelector"
                Me.Editor1.SelText = ""
                Me.Document.Elements("K" & lngKey).InsertIntoEditor Me.Editor1, , True
                Me.Editor1.ForceEdit = blnForce
                Me.Editor1.Tag = ""
                '同时弹出编辑窗体供输入
                If (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) And Me.Document.Elements("K" & lngKey).替换域 <> 1 Then
                    bInKeys = FindKey(Editor1, "E", lngKey, lSS, lSE, lES, lEE, bNeeded)
                    If bInKeys Then
                        '定位
                        Me.Editor1.Range(lSE, lES).Selected
                        ShowEleEditor 0, 0
                    End If
                End If
            End If
        End If
    End If
    If tblThis.Visible = False Then Call ClearNoUseUndoList
End Sub

Private Sub cbrThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

'################################################################################################################
'## 功能：  编辑器位置调整
'################################################################################################################
Private Sub cbrThis_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    On Error Resume Next
    cbrThis.GetClientRect Left, Top, Right, Bottom

    If Right >= Left And Bottom >= Top Then
        If Not Me.Document Is Nothing Then
            
            If imgX_S.Top > Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000 Then
                imgX_S.Top = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000
            End If
            
            imgX_S.Move Left, imgX_S.Top, Right - Left
            
            If picPane.Visible Then
            
                If imgX_S.Top > Bottom - Top - 1000 Then imgX_S.Top = Bottom - Top - 1000
                            
                picPane.Move Left, Top, Right - Left, imgX_S.Top - Top
                If imgX_S.Top < 0 Then imgX_S.Top = picPane.Top + picPane.Height
            End If
            
            If ChildMode = False And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                
                If picPane.Visible Then
                    picPatiInfo.Move Left, imgX_S.Top + imgX_S.Height, Right - Left
                    Editor1.Move Left, picPatiInfo.Top + picPatiInfo.Height, Right - Left, Bottom - Top - (picPane.Height + imgX_S.Height + picPatiInfo.Height)
                Else
                    picPatiInfo.Move Left, Top, Right - Left
                    Editor1.Move Left, picPatiInfo.Top + picPatiInfo.Height, Right - Left, Bottom - Top - picPatiInfo.Height
                End If
                picPatiInfo.Visible = True
            Else
                If picPane.Visible Then
                    Editor1.Move Left, imgX_S.Top + imgX_S.Height, Right - Left, Bottom - Top - picPane.Height - imgX_S.Height
                Else
                    Editor1.Move Left, Top, Right - Left, Bottom - Top
                End If
                picPatiInfo.Visible = False
            End If
        End If
    Else
        Editor1.Move 0, 0, 0, 0
    End If
    If Editor1.ViewMode = cprNormal Then
        picPenInput.Move Editor1.Width - picPenInput.Width - 300, Editor1.Height - picPenInput.Height - 300
    Else
        picPenInput.Visible = False
    End If
End Sub

'################################################################################################################
'## 功能：  菜单&工具栏更新事件
'################################################################################################################
Private Sub cbrThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, eEPRType As EPRDocTypeEnum
    
    If Me.Visible = False Then Exit Sub
    If Me.Document Is Nothing Then Exit Sub
    If mblnReadOnly And Control.ID <> ID_FILE_EXIT And Control.ID <> ID_COMMON_CANCEL Then
        '只读查阅模式下所有菜单无效（除了退出菜单）
        Control.Enabled = False
        Exit Sub
    End If
    eEPRType = Me.Document.EPRFileInfo.种类
    
    Select Case Control.ID
    Case ID_FILE_PRINTPREVIEW, ID_FILE_PRINT, ID_FILE_PRINTINWORD
        Control.Enabled = mblnCanPrint
        If Control.Enabled And (eEPRType = cpr住院病历 Or eEPRType = cpr门诊病历) Then
            Control.Enabled = IIf(Document.Signs.Count = 0, InStr(1, gstrPrivsEpr, "未签名打印") > 0, InStr(1, gstrPrivsEpr, "病历打印") > 0)
        End If
    Case ID_FILE_EXIT, ID_COMMON_CANCEL
        Control.Enabled = (mblnChildMode = False)
        Control.Visible = Control.Enabled
        If mintStyle = -1 Then
            Control.Enabled = False: Control.Visible = False
        End If
    Case ID_EDIT_UNDO
        Control.Enabled = CanUndo And mblnAutosave And (Me.Editor1.ReadOnly = False)
        Control.Visible = mblnAutosave
    Case ID_EDIT_REDO
        Control.Enabled = CanRedo And mblnAutosave And (Me.Editor1.ReadOnly = False)
        Control.Visible = mblnAutosave
    Case ID_VIEW_STRUCTURE
        Control.Checked = mfrmCompends.Visible
    Case ID_VIEW_PHRASEDEMO
        Control.Checked = mfrmSentenceDetailed.Visible
    Case ID_VIEW_SEGMENT
        Control.Enabled = (Me.Document.EditType <> cprET_病历文件定义)
        Control.Checked = mfrmSegments.Visible
    Case ID_VIEW_PACSPIC
        Control.Visible = (eEPRType = cpr诊疗报告)
        Control.Enabled = (eEPRType = cpr诊疗报告)
        Control.Checked = mfrmSegments.Visible
    Case ID_VIEW_HISTORYREPORT
        Control.Visible = (eEPRType = cpr诊疗报告)
        Control.Enabled = (eEPRType = cpr诊疗报告)
        Control.Checked = mfrmHistoryReport.Visible
    Case ID_VIEW_HISTORYWINDOW
        Control.Enabled = mblnExistHistroy
        Control.Visible = Control.Enabled
        Control.Checked = picHistoryInfo.Visible
    Case ID_FILE_CLEAR
        If Me.Editor1.AuditMode Then
            Control.Enabled = False
        Else
            If Document Is Nothing Then
                Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False)
            Else
                If Me.Document.EditType = cprET_单病历审核 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False)
                End If
            End If
        End If
    Case ID_FILE_IMPORT
        If Me.Editor1.AuditMode Then
            Control.Enabled = False
        Else
            If Document Is Nothing Then
                Control.Enabled = Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False And Editor1.AuditMode = False
            Else
                If Me.Document.EditType = cprET_病历文件定义 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False And Editor1.AuditMode = False
                End If
            End If
        End If
        If Control.Enabled Then Control.Enabled = InStr(gstrPrivsEpr, "历史文件") > 0
    Case ID_FILE_IMPORTFROMXML
        If Me.Editor1.AuditMode Then
            Control.Enabled = False
        Else
            If Document Is Nothing Then
                Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False)
            Else
                If Me.Document.EditType = cprET_单病历审核 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False)
                End If
            End If
        End If
        If Control.Enabled Then Control.Enabled = InStr(gstrPrivsEpr, "导出/入XML文件") > 0
    Case ID_FILE_SAVE, ID_FILE_SAVE_QUIT
        Control.Enabled = (Editor1.Modified) And (Me.Document.目标版本 <= 16) And (Editor1.ViewMode = cprNormal)
        If Control.ID = ID_FILE_SAVE_QUIT And mintStyle = -1 Then
            Control.Enabled = False: Control.Visible = False
        End If
    Case ID_FILE_SAVEAS
        Control.Enabled = (Editor1.ViewMode = cprNormal And mblnCanPrint)
        If Control.Enabled Then Control.Enabled = InStr(gstrPrivsEpr, "导出RTF文件") > 0
    Case ID_FILE_EXPORTTOHTML
        Control.Enabled = (Editor1.ViewMode = cprNormal And mblnCanPrint)
    Case ID_FILE_EXPORTTOXML
        Control.Enabled = (Editor1.ViewMode = cprNormal And mblnCanPrint)
        Control.Enabled = InStr(gstrPrivsEpr, "导出/入XML文件") > 0
    Case ID_FILE_SAVEASEPRDEMO, ID_FILE_SAVEASSEGMENT
        Control.Enabled = (Editor1.ViewMode = cprNormal And Me.Document.EditType <> cprET_病历文件定义)
    Case ID_EDIT_CUT:
        Control.Enabled = (tblThis.Visible) Or (Editor1.CanCopy And Editor1.ViewMode = cprNormal And (Me.Editor1.ReadOnly = False))
    Case ID_EDIT_COPY
        If Me.ActiveControl Is edtThis Then
            Control.Enabled = (edtThis.CanCopy And edtThis.ViewMode = cprNormal)
        Else
            Control.Enabled = (tblThis.Visible) Or (Editor1.CanCopy And Editor1.ViewMode = cprNormal)
        End If
        Control.Visible = InStr(gstrPrivsEpr, "内容复制") > 0
    Case ID_EDIT_COPYSELF
         Control.Enabled = (tblThis.Visible) Or (Me.Editor1.Selection.Font.ForeColor <> tomUndefined And Me.Editor1.Selection.Font.Protected = False) And Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False
         Control.Visible = InStr(gstrPrivsEpr, "专用复制") > 0
    Case ID_EDIT_COPYOUT
        If Me.ActiveControl Is edtThis Then
            Control.Enabled = (edtThis.CanCopy And edtThis.ViewMode = cprNormal)
        Else
            Control.Enabled = (tblThis.Visible) Or (Editor1.CanCopy And Editor1.ViewMode = cprNormal)
        End If
        Control.Visible = InStr(gstrPrivsEpr, "内容复制") > 0
    Case ID_EDIT_PASTE:
        Control.Enabled = (tblThis.Visible) Or (Me.Editor1.Selection.Font.ForeColor <> tomUndefined And Me.Editor1.Selection.Font.Protected = False) And Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False
        Control.Visible = InStr(gstrPrivsEpr, "内容复制") > 0
    Case ID_EDIT_DELETE
        Control.Enabled = Editor1.ViewMode = cprNormal
    Case ID_EDIT_FORMATBRUSH
        Control.Enabled = (tblThis.Visible = False) And (Editor1.ViewMode = cprNormal) And Editor1.AuditMode = False And Editor1.ReadOnly = False
        Control.Checked = mblnFmtBrushDown
    Case ID_EDIT_FIND, ID_EDIT_FINDNEXT     ', ID_EDIT_UNDO, ID_EDIT_REDO
        Control.Enabled = (Editor1.ViewMode = cprNormal)
    Case ID_EDIT_REPLACE
        Control.Enabled = (Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_INSERT_DATETIME, ID_INSERT_DATE, ID_INSERT_TIME, ID_INSERT_SPECIALCHAR
        Control.Enabled = Editor1.ViewMode = cprNormal And Me.Editor1.ReadOnly = False
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 Then
                On Error Resume Next
                Control.Enabled = Control.Enabled And (tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False)
            End If
        Else
            Control.Enabled = Control.Enabled And Editor1.Selection.Font.Protected = False
        End If
    Case ID_INSERT_ELEMENT
        Control.Enabled = Editor1.ViewMode = cprNormal And Me.Editor1.ReadOnly = False
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 Then
                On Error Resume Next
                Control.Enabled = Control.Enabled And (tblThis.Cells("K" & tblThis.SelectedCellKey).Picture Is Nothing)
            End If
        End If
    Case ID_INSERT_PICTURE
        Control.Enabled = (Editor1.ViewMode = cprNormal And (Editor1.Selection.Font.Protected = False Or mbEditInTable Or ucPacsImgCanvas1.Visible) And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 Then
                On Error Resume Next
                Control.Enabled = Control.Enabled And ((Not tblThis.Cells("K" & tblThis.SelectedCellKey).Picture Is Nothing) Or (tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = ""))
            End If
        End If
    Case ID_INSERT_DOCADVISE
        Control.Enabled = Not Editor1.AuditMode
        If Control.Enabled Then Control.Enabled = Not Editor1.ReadOnly
        Control.Visible = (eEPRType = cpr门诊病历)
    Case ID_INSERT_EPRDEMO
        If Me.Editor1.AuditMode Then
            Control.Enabled = False
        Else
            If Me.Document Is Nothing Then
                Control.Enabled = (Editor1.ViewMode = cprNormal) And Me.Editor1.ReadOnly = False
            Else
                Control.Enabled = (Editor1.ViewMode = cprNormal) And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) And Me.Editor1.ReadOnly = False
            End If
        End If
    Case ID_INSERT_PACSPIC
        Control.Enabled = (Document.EPRFileInfo.种类 = cpr诊疗报告 And Me.Editor1.ReadOnly = False)
        Control.Visible = Control.Enabled
    Case ID_EDIT_ADDCOMPEND, ID_EDIT_MODCOMPEND, ID_EDIT_REFCOMPEND, ID_EDIT_DELCOMPEND
        If Me.Editor1.AuditMode Then
            If Control.ID <> ID_EDIT_REFCOMPEND Then Control.Enabled = False
        Else
            If Editor1.ViewMode = cprNormal Then
                If Control.ID = ID_EDIT_DELCOMPEND Then
                    If mfrmCompends.Tree.SelectedItem Is Nothing Then
                        Control.Enabled = False
                    Else
                        Control.Enabled = (Me.Editor1.ReadOnly = False)
                    End If
                ElseIf Control.ID = ID_EDIT_MODCOMPEND Then
                    If mfrmCompends.Tree.SelectedItem Is Nothing Then
                        Control.Enabled = False
                    Else
                        Control.Enabled = (Me.Editor1.ReadOnly = False)
                    End If
                ElseIf Control.ID = ID_EDIT_ADDCOMPEND Then
                    Control.Enabled = (Editor1.Selection.Font.Protected = False)
                Else
                    Control.Enabled = (Me.Editor1.ReadOnly = False)
                End If
            Else
                Control.Enabled = False
            End If
        End If
        If Not Me.Document Is Nothing Then
            If Me.Document.EditType <> cprET_病历文件定义 And Control.ID <> ID_EDIT_REFCOMPEND Then '只能在定义时修改提纲
                Control.Enabled = False
                Control.Visible = False
            End If
        End If
    Case ID_Main_FORMAT
        Control.Enabled = Control.Visible And (InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0)
        Control.Visible = Control.Enabled
        If Control.Visible Then
            If Not Document Is Nothing Then
                Control.Visible = Not (Document.EPRDemoInfo.性质 <> 0 And Document.EditType = cprET_全文示范编辑)
            End If
        End If
    Case ID_FORMAT_FONTNAME
        Control.Visible = (Editor1.AuditMode = False) And (InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0)
        
        If Control.Type = xtpControlComboBox Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    On Error Resume Next
                    Control.Text = (tblThis.Cells("K" & tblThis.SelectedCellKey).FontName)
                End If
            Else
                Control.Text = CStr(Editor1.Selection.Font.Name)
            End If
        End If
    Case ID_FORMAT_PARA
        If Not Me.Document Is Nothing Then
            Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
            Control.Visible = (Editor1.AuditMode = False)
        End If
    Case ID_FORMAT_FONT
        Control.Visible = (Editor1.AuditMode = False) And (InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0)
        If Not Me.Document Is Nothing Then Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_FONTSIZE
        Control.Visible = (Editor1.AuditMode = False) And (InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0)
        If Control.Type = xtpControlComboBox Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    On Error Resume Next
                    Control.Text = (tblThis.Cells("K" & tblThis.SelectedCellKey).FontSize)
                End If
            Else
                Control.Text = CStr(IIf(Editor1.Selection.Font.Size = tomUndefined, "", GetFontSizeChinese(Editor1.Selection.Font.Size)))
            End If
        End If
    Case ID_FORMAT_BOLD
        Control.Visible = Control.Visible And (InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0)
        Control.Checked = Editor1.Selection.Font.Bold
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE
        Control.Visible = Control.Visible And (InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0)
        Control.Checked = (Editor1.Selection.Font.Underline <> cprNone)
        Control.Enabled = (Editor1.AuditMode = False And Editor1.Selection.Font.Protected = False)
    Case ID_FORMAT_ALIGNLEFT
        Control.Visible = InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0
        Control.Checked = (Editor1.Selection.Para.Alignment = cprHALeft)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_ALIGNCENTER
        Control.Visible = InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0
        Control.Checked = (Editor1.Selection.Para.Alignment = cprHACenter)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_ALIGNRIGHT
        Control.Visible = InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0
        Control.Checked = (Editor1.Selection.Para.Alignment = cprHARight)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_LINESPACE
        Control.Visible = InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0
        Control.Checked = (Editor1.Selection.Para.LineSpacing > 1#)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_SPACEBEFORE
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_SPACEAFTER
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_FIRSTINDENT
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_FIRSTHUNGING
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTARABIC
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNumberAsArabic)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTBULLETS
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTBullet)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTLCHAR
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNumberAsLCLetter)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTLROME
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNumberAsLCRoman)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTUCHAR
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNumberAsUCLetter)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTUROME
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNumberAsUCRoman)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTNONE
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNone)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTSETUP
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_PROTECT
        Control.Checked = (Editor1.Selection.Font.Protected And Editor1.Selection.Font.ForeColor = PROTECT_FORECOLOR)
        Control.Enabled = (Me.Document.EditType = cprET_病历文件定义 Or Me.Document.EditType = cprET_全文示范编辑 Or InStr(1, gstrPrivsEpr, "保护文本处理") > 0)
    Case ID_FORMAT_SUPER
        Control.Checked = Editor1.Selection.Font.Superscript
        Control.Enabled = (tblThis.Visible = False)
    Case ID_FORMAT_SUB
        Control.Checked = Editor1.Selection.Font.Subscript
        Control.Enabled = (tblThis.Visible = False)
    Case ID_FORMAT_UNDERLINE_DASH
        Control.Checked = (Editor1.Selection.Font.Underline = cprDash)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_DASHDOT
        Control.Checked = (Editor1.Selection.Font.Underline = cprDashDot)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_DASHDOT2
        Control.Checked = (Editor1.Selection.Font.Underline = cprDashDotDot)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_DOT
        Control.Checked = (Editor1.Selection.Font.Underline = cprDotted)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_THIN
        Control.Checked = (Editor1.Selection.Font.Underline = cprHair)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_THICK
        Control.Checked = (Editor1.Selection.Font.Underline = cprThick)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_WAVE
        Control.Checked = (Editor1.Selection.Font.Underline = cprWave)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_NONE
        Control.Checked = (Editor1.Selection.Font.Underline = cprNone)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_BACKGROUND
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_HIGHLIGHT
        Control.Enabled = (Editor1.AuditMode = False And Editor1.Selection.Font.Protected = False)
    Case ID_VIEW_RULER
        Control.Checked = Editor1.ShowRuler
    Case ID_VIEW_PENWINDOW
        Control.Checked = picPenInput.Visible
    Case ID_FORMAT_STYLE
        Control.Visible = (InStr(";" & gstrPrivsEpr & ";", ";字体格式设置;") > 0) And (Editor1.AuditMode = False)
        If Control.Visible Then Control.Enabled = (tblThis.Visible = False)
    Case ID_FORMAT_INDENTDECREASE, ID_FORMAT_INDENTINCREASE
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_STYLEWINDOW
        '样式窗格
        Control.Checked = mfrmStyleMan.Visible
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
        Control.Visible = (Editor1.AuditMode = False)
    Case ID_VIEW_HEADFOOT
        Control.Enabled = (Me.Document.EditType = cprET_病历文件定义) And (Editor1.AuditMode = False)
    Case ID_FILE_PAGESETUP
        Control.Enabled = (InStr(1, gstrPrivsEpr, "修改页面设置") > 0) Or ((Me.Document.EditType = cprET_病历文件定义) And (Editor1.AuditMode = False))
    Case ID_SIGN, ID_SIGN_QUIT
        If Me.Document.EditType = cprET_单病历编辑 Then
            Control.Enabled = (Me.Document.Signs.Count = 0 And Me.Editor1.ReadOnly = False)
        ElseIf Me.Document.EditType = cprET_单病历审核 Then
            Control.Enabled = (Me.Document.目标版本 <= 16 And Me.Editor1.ReadOnly = False)
        Else
            Control.Enabled = False: Control.Visible = False
        End If
        If Control.ID = ID_SIGN_QUIT And mintStyle = -1 Then
            Control.Enabled = False: Control.Visible = False
        End If
    Case ID_PATISIGN
        Control.Visible = False
        If Me.Document.EditType = cprET_单病历编辑 Then
            If eEPRType = cpr知情文件 And mblnEnPtSign Then
                Control.Visible = True
            End If
            Control.Enabled = (Editor1.AuditMode = False)
            If Control.Enabled Then Control.Enabled = (Me.Document.Signs.Count = 0 And Me.Editor1.ReadOnly = False)
            Control.Caption = IIf(mblnPatiSign, "患者撤签", "患者签名")
        Else
            Control.Enabled = False: Control.Visible = False 'Control.Visible = Me.Document.EditType = cprET_单病历编辑 会引起DockPane不能还原
        End If
    Case ID_UNTREAD
        If Me.Document.Signs.Count > 0 Then
            If Me.Document.Signs("K" & Document.Signs.GetMaxKey).姓名 <> gstrUserName And Me.Document.Signs("K" & Document.Signs.GetMaxKey).姓名 <> gstrSignName _
                And InStr(gstrPrivsEpr, "回退他人签名") = 0 Then  '回退他人签名权限控制
                Control.Visible = False
            Else
                Control.Visible = True
            End If
        End If
        
        If Me.Document.EditType = cprET_单病历编辑 Then
            Control.Enabled = IIf(mblnIsMultiMode, Me.Document.EPRPatiRecInfo.最后版本 = 1, Me.Document.Signs.Count > 0) And Me.Editor1.Modified = False And Editor1.ReadOnly = True
        ElseIf Me.Document.EditType = cprET_单病历审核 Then
            Control.Enabled = (Me.Document.EPRPatiRecInfo.最后版本 > 1 Or Me.Document.Signs.Count > 1) And Me.Editor1.Modified = False And Editor1.ReadOnly = False
        Else
            Control.Enabled = False: Control.Visible = False
        End If
    Case ID_REVISION_PREV, ID_REVISION_NEXT, ID_REVISION_RESET
        Control.Enabled = Me.Editor1.AuditMode And (Me.Editor1.ReadOnly = False)
        Control.Visible = Control.Enabled
    Case ID_DIAGNOSIS '诊断
        Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.Selection.Font.Protected = False) And Me.Editor1.ReadOnly = False And Me.Document.EditType <> cprET_病历文件定义
    Case ID_EDIT_MARKEDPIC, ID_EDIT_OUTERPIC
        Control.Enabled = Me.Editor1.ReadOnly = False And (Editor1.ViewMode = cprNormal) And (Editor1.AuditMode = False)
    Case ID_INSERT_TABLE
        Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.Selection.Font.Protected = False And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_FORMAT_FORECOLOR, ID_TABLE_CELLALIGNMENT, ID_TABLE_CURRENCY, ID_TABLE_PERCENT, ID_TABLE_KILOBIT, ID_TABLE_CELLPROTECTED, ID_TABLE_INSERTPICTURE, ID_TABLE_BEELEMENTS
        Control.Enabled = tblThis.Visible And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_TABLE_MERGE
        Control.Enabled = tblThis.Visible And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
        If tblThis.SelectedCellKey > 0 Then
            On Error Resume Next
            Control.Checked = (Len(tblThis.Cells("K" & tblThis.SelectedCellKey).MergeInfo) = 16)
        End If
    Case ID_TABLE_DELETETABLE, ID_TABLE_DELETECOL, ID_TABLE_DELETEROW, ID_TABLE_FORMATCELL, ID_TABLE_PROPERTY
        Control.Enabled = tblThis.Visible And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_TABLE_INSERTROWDOWN, ID_TABLE_INSERTROWUP, ID_TABLE_INSERTCOLLEFT, ID_TABLE_INSERTCOLRIGHT, ID_TABLE_INSERTINHERITROW
        Control.Enabled = tblThis.SelectedCellKey > 0 And tblThis.Visible And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_TABLE_INSERTTABLE
        Control.Enabled = (tblThis.Visible = False) And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_TABLE_FORMATROWHEIGHT, ID_TABLE_FORMATCOLWIDTH
        Control.Enabled = (tblThis.Visible = True) And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_TABLE_CELLALIGNMENT1, ID_TABLE_CELLALIGNMENT2, ID_TABLE_CELLALIGNMENT3, ID_TABLE_CELLALIGNMENT4, ID_TABLE_CELLALIGNMENT5, ID_TABLE_CELLALIGNMENT6, ID_TABLE_CELLALIGNMENT7, ID_TABLE_CELLALIGNMENT8, ID_TABLE_CELLALIGNMENT9
        Control.Enabled = (tblThis.Visible = True) And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_Main_HELP
        Control.Visible = Not mblnChildMode
    End Select
    
    If Not Control.Visible Then Control.Enabled = False
End Sub

Private Sub cPicEditor_pOK(ByRef FinalPicture As StdPicture, ByVal lngWidth As Long, ByVal lngHeight As Long)
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim blnForce As Boolean

    lKey = cPicEditor.lngKeyOfPic
    If lKey > 0 And FinalPicture <> 0 Then
        If tblThis.Visible Then
            '表格中的图片
            If Val(tblThis.Tag) > 0 Then
                '将图片对象保存到类对象中
                Dim ctlPic As VB.PictureBox
                Set ctlPic = gfrmPublic.Controls.Add("VB.PictureBox", "ctlPic" & CLng(Timer * 1000))
                ctlPic.AutoRedraw = True
                ctlPic.BorderStyle = 0
                ctlPic.Height = lngHeight
                ctlPic.Width = lngWidth
                ShowPicMarks ctlPic, FinalPicture, Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).PicMarks

                Set Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).OrigPic = FinalPicture
                Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).Width = lngWidth
                Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).Height = lngHeight
                Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).OrigWidth = lngWidth
                Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).OrigHeight = lngHeight

                '保存到单元格中
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = ""
                tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lKey
                tblThis.Cells("K" & tblThis.SelectedCellKey).Picture = ctlPic.Picture
                tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = ""
                tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
                tblThis.Modified = True
                tblThis.Refresh
                tblThis_Resize tblThis.Width, tblThis.Height
                Editor1.Modified = True

                gfrmPublic.Controls.Remove ctlPic
                Set ctlPic = Nothing
            End If
        Else
            '替换图片
            Set Me.Document.Pictures("K" & lKey).OrigPic = FinalPicture

            Me.Document.Pictures("K" & lKey).OrigWidth = lngWidth
            Me.Document.Pictures("K" & lKey).OrigHeight = lngHeight
            Me.Document.Pictures("K" & lKey).Width = lngWidth
            Me.Document.Pictures("K" & lKey).Height = lngHeight
            Me.Document.Pictures("K" & lKey).Modified = True

            bInKeys = FindKey(Me.Editor1, "P", lKey, lSS, lSE, lES, lEE, bNeeded)
            If bInKeys = False Then Exit Sub
            Dim ParaFmt As New cParaFormat

            With Me.Editor1
                blnForce = .ForceEdit
                .Freeze
                .Tag = "cPicEditor_pOK"
                .ForceEdit = True
                Set ParaFmt = .Range(lSE, lES).Para.GetParaFmt

                If Me.Document.Pictures("K" & lKey).是否换行 Then
                    .Range(lSS, lEE + 2).Text = ""
                Else
                    .Range(lSS, lEE).Text = ""
                End If
                Me.Document.Pictures("K" & lKey).InsertIntoEditor Me.Editor1, lSS, True

                .Range(lSE, lES).Para.SetParaFmt ParaFmt
                .Range(lSS, lEE).Font.Protected = True
                .ForceEdit = blnForce
                .Tag = ""
                .UnFreeze
            End With
        End If
    End If
End Sub

'################################################################################################################
'## 功能：  添加可停靠窗体
'################################################################################################################
Private Sub DkpThis_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case ID_VIEW_PHRASEDEMO     '示范词句窗体
            Item.Handle = mfrmSentenceDetailed.hwnd
        Case ID_VIEW_SEGMENT        '示范段落窗体
            Item.Handle = mfrmSegments.hwnd
        Case ID_VIEW_STRUCTURE      '文档结构图窗体
            Item.Handle = mfrmCompends.hwnd
        Case ID_VIEW_HISTORYWINDOW  '共享页面文件内容
            Item.Handle = picPane.hwnd
        Case ID_FORMAT_STYLEWINDOW  '段落样式维护
            Item.Handle = mfrmStyleMan.hwnd
        Case ID_VIEW_PACSPIC
            Item.Handle = mfrmPacsPic.hwnd
        Case ID_VIEW_MULTIDOCVIEW
            If mfrmMultiDocView Is Nothing Then
                If mblnIsMultiMode Then
                    Set mfrmMultiDocView = New frmMultiDocView
                    Item.Handle = mfrmMultiDocView.hwnd
                ElseIf Not DkpThis.FindPane(ID_VIEW_MULTIDOCVIEW) Is Nothing Then
                    DkpThis.FindPane(ID_VIEW_MULTIDOCVIEW).Hide
                End If
            Else
                Item.Handle = mfrmMultiDocView.hwnd
            End If
        Case ID_VIEW_HISTORYREPORT
            Item.Handle = mfrmHistoryReport.hwnd
        Case ID_VIEW_Assistant '输入助手
            Item.Handle = mfrmDocksymbol.hwnd
    End Select
End Sub

'################################################################################################################
'## 功能：  根据当前选中内容，双击打开图片、表格编辑器
'################################################################################################################
Private Sub Editor1_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    If Editor1.ReadOnly Then Exit Sub
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal And Editor1.AuditMode = False Then
        bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '查找关键字 ID ！
        If bInKeys = False Then Exit Sub
        If sType = "P" Then '编辑图片
            If Me.Document.Pictures("K" & lKey).PictureType = EPRMarkedPicture Then
                Editor1.ShowUIInterface
                ucPictureEditor1.ShowMe Me, Editor1.hwnd, cbrThis, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey), _
                    Editor1.UILeft, Editor1.UITop, Editor1.UIWidth, Editor1.UIHeight, False
            ElseIf Me.Document.Pictures("K" & lKey).PictureType = EPROutPicture Then
                cPicEditor.ShowPicEditor glngSys, gcnOracle, Me.Document.Pictures("K" & lKey).OrigPic, lKey, Me.Document.Pictures("K" & lKey).保留对象, Me, False
            End If
        End If
    ElseIf Editor1.ViewMode = cprNormal Then
        bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
        If bInKeys Then
            Select Case sType
            Case "E"
                If Me.Document.Elements("K" & lKey).输入形态 = 1 Then Exit Sub
                Me.Editor1.Range(lSE, lES).Selected
                ShowEleEditor 0, 0
            Case "S" '书写签名窗体
                Me.Editor1.Range(lSE, lES).Selected
            End Select
        End If
    End If
End Sub

Private Sub Editor1_GetDelCharColor(COLOR As stdole.OLE_COLOR)
    If Editor1.AuditMode Then
        COLOR = Me.Document.GetDelCharColor(COLOR)
    End If
End Sub

Private Sub Editor1_GetNewCharColor(COLOR As stdole.OLE_COLOR)
    If Editor1.AuditMode Then
        COLOR = Me.Document.GetNewCharColor(COLOR)
    End If
End Sub

Private Sub Editor1_IsDelCharColor(ByVal COLOR As stdole.OLE_COLOR, blnIsDelCharColor As Boolean)
    If Editor1.AuditMode Then
        blnIsDelCharColor = Me.Document.IsDelCharColor(COLOR)
    End If
End Sub

Private Sub Editor1_IsNewCharColor(ByVal COLOR As stdole.OLE_COLOR, blnIsNewCharColor As Boolean)
    If Editor1.AuditMode Then
        blnIsNewCharColor = Me.Document.IsNewCharColor(COLOR)
    End If
End Sub

'################################################################################################################
'## 功能：  特殊情况下输入的处理；
'##         如果选中诊治要素，则弹出诊治要素编辑器。
'################################################################################################################
Private Sub Editor1_KeyDown(ViewMode As zlRichEditor.ViewModeEnum, KeyCode As Integer, Shift As Integer)
    If Editor1.ReadOnly Then Exit Sub
    Dim i As Long, blnForce As Boolean
    If ViewMode = cprPaper Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyEscape Then Exit Sub
    Editor1.Tag = "Editor1.KeyDown"

    If Me.Editor1.AuditMode Then
        If Shift <> 0 Then Editor1.Tag = "": Exit Sub
        Select Case KeyCode
        Case 0, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
            vbKeyEscape, vbKeyDelete, vbKeyBack, vbKeyTab, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
            vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital, _
            vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12

            Editor1.Tag = ""
            Exit Sub
        End Select

        If Editor1.SelLength > 0 Then
            If Editor1.Selection.Font.Protected = False Then
                Editor1.ForceEdit = True
                Editor1.Selection.Font.ForeColor = Me.Document.GetDelCharColor(Editor1.Selection.Font.ForeColor)
                Editor1.Selection.Font.Strikethrough = True
                Editor1.Selection.Font.Underline = False
                Editor1.ForceEdit = False
            End If
            Editor1.Range(Editor1.Selection.StartPos + Len(Editor1.Selection.Text), Editor1.Selection.StartPos + Len(Editor1.Selection.Text)).Selected
            Editor1.SelLength = 0
        End If
        '审核模式下的特殊情况输入的处理
        '发现在隐藏关键字后面，则自动加一个空格（非保护和隐藏属性）
        With Editor1
            blnForce = .ForceEdit
            If .SelLength = 0 And .Selection.Font.ForeColor = PROTECT_FORECOLOR Then
                .ForceEdit = True
                .Selection.Font.ForeColor = tomAutoColor
                .ForceEdit = blnForce
            End If
            i = .Selection.StartPos
LL1:
            If .Range(i - 1, i).Font.Hidden And _
                .Range(i, i + 1).Font.Hidden = False And _
                .Range(i, i + 1).Font.Protected = False Then
                'A问题：（隐藏文本）|普通文本
                .ForceEdit = True
                .Range(i, i).Text = " "
                .Range(i, i + 1).Font.Protected = False
                .Range(i, i + 1).Font.Hidden = False
                .Range(i, i + 1).Font.ForeColor = Me.Document.GetNewCharColor(tomAutoColor)
                .Range(i, i + 1).Font.BackColor = tomAutoColor
                .Range(i, i + 1).Font.Underline = cprNone
                .Range(i, i + 1).Font.Strikethrough = False
                .Range(i, i + 1).Selected
                .ForceEdit = blnForce
            Else
                If .Range(i - 1, i).Font.Hidden And _
                    .Range(i, i + 1).Font.Hidden = False And _
                    .Range(i, i + 1).Font.Protected And .Range(i, i + 1).Font.ForeColor <> PROTECT_FORECOLOR Then
                    'B问题1：普通文本（隐藏文本）|（保护文本）（隐藏文本）普通文本
                    i = i - 16
                    If .Range(i - 1, i + 3) Like ")?S(" And _
                        .Range(i - 1, i + 3).Font.Hidden = True Then
                        'C问题：（隐藏文本）（保护文本）（隐藏文本）|（隐藏文本）（保护文本）（隐藏文本）
                        mlngHP = -1
                        .ForceEdit = True
                        .Range(i, i).Font.Protected = False
                        .Range(i, i).Font.Hidden = False
                        .Range(i - 1, i).Font.ForeColor = vbBlack
                        blnSpaceEvent = True
                        .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")

                        '新增文本的格式设置
                        .Range(i, i + 1).Font.Protected = False
                        .Range(i, i + 1).Font.Hidden = False
                        .Range(i, i + 1).Font.ForeColor = Me.Document.GetNewCharColor(.Range(i, i + 1).Font.ForeColor)
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.Underline = cprNone

                        .Range(i + 1, i + 1).Selected
                        .ForceEdit = blnForce
                    ElseIf .Range(i + 1, i + 3) = "E(" And .Range(i, i + 3).Font.Protected And .Range(i, i + 3).Font.ForeColor <> PROTECT_FORECOLOR And _
                        .Range(i + 16, i + 18) = vbCrLf And .Range(i + 16, i + 18).Font.Protected And .Range(i + 16, i + 18).Font.ForeColor <> PROTECT_FORECOLOR Then
                        'D问题：提纲后面跟图片，在之间没有文字时，无法插入其他文字
                        i = i + 16
                        mlngHP = -1
                        .ForceEdit = True
                        .Range(i, i).Font.Protected = False
                        .Range(i, i).Font.Hidden = False
                        .Range(i - 1, i).Font.ForeColor = vbBlack
                        blnSpaceEvent = True
                        .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")

                        '新增文本的格式设置
                        .Range(i, i + 1).Font.Protected = False
                        .Range(i, i + 1).Font.Hidden = False
                        .Range(i, i + 1).Font.ForeColor = Me.Document.GetNewCharColor(.Range(i, i + 1).Font.ForeColor)
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.Underline = cprNone

                        If (.Range(i - 16, i - 14) <> "EE") Then
                            .Range(i, i + 1).Selected
                        Else
                            .Range(i + 1, i + 1).Selected
                        End If
                        .ForceEdit = blnForce
                    Else
                        .Range(i, i).Selected
                        On Error Resume Next
                        .OriginRTB.SelColor = Me.Document.GetNewCharColor(.OriginRTB.SelColor)
                        .OriginRTB.SelStrikeThru = False    '去掉删除线
                        .OriginRTB.SelUnderline = False     '去掉下划线
                    End If
                ElseIf .Range(i - 1, i).Font.Hidden = False And _
                    .Range(i - 1, i).Font.Protected And .Range(i - 1, i).Font.ForeColor <> PROTECT_FORECOLOR And _
                    .Range(i, i + 1).Font.Hidden Then
                    'B问题2：普通文本（隐藏文本）（保护文本）|（隐藏文本）普通文本
                    i = i + 16
                    If .Range(i - 1, i + 3) Like ")?S(" And _
                        .Range(i - 1, i + 3).Font.Hidden = True Then
                        'C问题：（隐藏文本）（保护文本）（隐藏文本）|（隐藏文本）（保护文本）（隐藏文本）
                        mlngHP = -1
                        .ForceEdit = True
                        .Range(i, i).Font.Protected = False
                        .Range(i, i).Font.Hidden = False

                        '新增文本的格式设置
                        .Range(i - 1, i).Font.ForeColor = Me.Document.GetNewCharColor(.Range(i - 1, i).Font.ForeColor)
                        .Range(i - 1, i).Font.Strikethrough = False
                        .Range(i - 1, i).Font.Underline = cprNone

                        blnSpaceEvent = True
                        .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")

                        '新增文本的格式设置
                        .Range(i, i + 1).Font.Protected = False
                        .Range(i, i + 1).Font.Hidden = False
                        .Range(i, i + 1).Font.ForeColor = Me.Document.GetNewCharColor(.Range(i, i + 1).Font.ForeColor)
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.Underline = cprNone

                        .Range(i + 1, i + 1).Selected
                        .ForceEdit = blnForce
                    Else
                        GoTo LL1
                    End If
                ElseIf .Range(i - 1, i).Font.Hidden = False And .Range(i, i + 2) = vbCrLf And .Range(i, i + 2).Font.Protected And .Range(i, i + 2).Font.ForeColor <> PROTECT_FORECOLOR Then
                    .ForceEdit = True
                    .Range(i, i) = " "
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = Me.Document.GetNewCharColor(.Range(i, i + 1).Font.ForeColor)
                    .Range(i, i + 1).Font.Strikethrough = False
                    .Range(i, i + 1).Font.Underline = cprNone
                    .Range(i, i).Selected
                    .Selection.Font.ForeColor = Me.Document.GetNewCharColor(.Selection.Font.ForeColor)
                    .ForceEdit = blnForce
                Else
                    On Error Resume Next
                    .OriginRTB.SelColor = Me.Document.GetNewCharColor(.OriginRTB.SelColor)
                    .OriginRTB.SelStrikeThru = False    '去掉删除线
                    .OriginRTB.SelUnderline = False     '去掉下划线
                    .OriginRTB.SelItalic = True
                    If KeyCode = 1 Then KeyCode = 0
                End If
            End If
        End With
        Editor1.Tag = ""
        Exit Sub
    End If

    If Editor1.SelLength > 0 Then GoTo LL4
    If Shift <> 0 Then GoTo LL4
    Select Case KeyCode
    Case 0, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
        vbKeyEscape, vbKeyDelete, vbKeyBack, vbKeyTab, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
        vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital, _
        vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12

        GoTo LL4
    End Select

    '发现在隐藏关键字后面，则自动加一个空格（非保护和隐藏属性）
    With Editor1
        blnForce = .ForceEdit
        i = .Selection.StartPos
        If .Range(i, i + 2) = vbCrLf Then
            .ForceEdit = True
            .Selection.Font.Protected = False
            .Selection.Font.Hidden = False
            .Selection.Font.ForeColor = tomAutoColor
            .Selection.Font.BackColor = tomAutoColor
            .Selection.Font.Underline = cprNone
            .Selection.Font.Strikethrough = False
            .ForceEdit = blnForce
        End If
LL3:
        If .Range(i - 1, i).Font.Hidden And _
            .Range(i, i + 1).Font.Hidden = False And _
            .Range(i, i + 1).Font.Protected = False Then
            'A问题：（隐藏文本）|普通文本
            .ForceEdit = True
            .Range(i, i).Text = " "
            .Range(i, i + 1).Font.Protected = False
            .Range(i, i + 1).Font.Hidden = False
            .Range(i, i + 1).Font.ForeColor = tomAutoColor
            .Range(i, i + 1).Font.BackColor = tomAutoColor
            .Range(i, i + 1).Font.Underline = cprNone
            .Range(i, i + 1).Font.Strikethrough = False
            .Range(i, i + 1).Selected
            .ForceEdit = blnForce
        Else
            If .Range(i - 1, i).Font.Hidden And _
                .Range(i, i + 1).Font.Hidden = False And _
                .Range(i, i + 1).Font.Protected Then
                'B问题1：普通文本（隐藏文本）|（保护文本）（隐藏文本）普通文本
                i = i - 16
                If .Range(i - 1, i + 3) Like ")?S(" And _
                    .Range(i - 1, i + 3).Font.Hidden = True Then
                    'C问题：（隐藏文本）（保护文本）（隐藏文本）|（隐藏文本）（保护文本）（隐藏文本）
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    .Range(i + 1, i + 1).Selected
                    .ForceEdit = blnForce
                ElseIf .Range(i + 1, i + 3) = "E(" And .Range(i, i + 3).Font.Protected And .Range(i, i + 3).Font.ForeColor <> PROTECT_FORECOLOR And _
                    .Range(i + 16, i + 18) = vbCrLf And .Range(i + 16, i + 18).Font.Protected And .Range(i + 16, i + 18).Font.ForeColor <> PROTECT_FORECOLOR Then
                    'D问题：提纲后面跟图片，在之间没有文字时，无法插入其他文字
                    i = i + 16
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    If (.Range(i - 16, i - 14) <> "EE") Then
                        .Range(i, i + 1).Selected
                    Else
                        .Range(i + 1, i + 1).Selected
                    End If
                    .ForceEdit = blnForce
                Else
                    .Range(i, i).Selected
                End If
            ElseIf .Range(i - 1, i).Font.Hidden = False And _
                .Range(i - 1, i).Font.Protected And .Range(i - 1, i).Font.ForeColor <> PROTECT_FORECOLOR And _
                .Range(i, i + 1).Font.Hidden Then
                'B问题2：普通文本（隐藏文本）（保护文本）|（隐藏文本）普通文本
                i = i + 16
                If .Range(i - 1, i + 3) Like ")?S(" And _
                    .Range(i - 1, i + 3).Font.Hidden = True Then
                    'C问题：（隐藏文本）（保护文本）（隐藏文本）|（隐藏文本）（保护文本）（隐藏文本）
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    .Range(i + 1, i + 1).Selected
                    .ForceEdit = blnForce
                Else
                    GoTo LL3
                End If
            ElseIf .Range(i - 1, i).Font.Hidden = False And .Range(i, i + 2) = vbCrLf And .Range(i, i + 2).Font.Protected And .Range(i, i + 2).Font.ForeColor <> PROTECT_FORECOLOR Then
                mlngHP = -1
                .ForceEdit = True
                .Range(i, i).Font.Protected = False
                .Range(i, i).Font.Hidden = False
                .Range(i - 1, i).Font.ForeColor = vbBlack
                blnSpaceEvent = True
                .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")
                .Range(i, i + 1).Font.Protected = False
                .Range(i, i + 1).Font.Hidden = False
                .Range(i, i + 1).Font.ForeColor = vbBlack
                If (.Range(i - 16, i - 14) <> "EE") Then
                    .Range(i, i + 1).Selected
                Else
                    .Range(i + 1, i + 1).Selected
                End If
                .ForceEdit = blnForce
            End If
        End If
    End With

LL4:
    If ViewMode = cprNormal Then
        If KeyCode = vbKeyF2 Then
            '显示诊治要素编辑器
            If ViewMode = cprNormal Then Call ShowEleEditor(0, Shift)
        ElseIf KeyCode = vbKeyReturn Then
            If Me.Editor1.Selection.GetType = cprSTPicture Then
                Editor1_DblClick ViewMode
            End If
        ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
            
        End If
    End If
    Editor1.Tag = ""
End Sub

Private Sub Editor1_KeyUp(ViewMode As zlRichEditor.ViewModeEnum, KeyCode As Integer, Shift As Integer)
    If Editor1.ReadOnly Then Exit Sub
    If txtPenInput.Visible And txtPenInput.Enabled Then txtPenInput.SetFocus: Exit Sub
End Sub

'################################################################################################################
'## 功能：  用户试图修改保护文本。
'##         如果当前是诊治要素，则弹出诊治要素编辑器。
'################################################################################################################
Private Sub Editor1_ModifyProtected(ViewMode As zlRichEditor.ViewModeEnum, bAllowDoIt As Boolean, ByVal lStart As Long, ByVal lEnd As Long, KeyAscii As Integer, Shift As Integer)
    If Editor1.ReadOnly Then bAllowDoIt = False: Exit Sub
    bAllowDoIt = False

    '如果在诊治要素中，则弹出编辑器
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    bBeteenKeys = IsBetweenAnyKeys(Editor1, lStart, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then
        Select Case sKeyType
        Case "E"
            ShowEleEditor KeyAscii, Shift
        End Select
    End If
End Sub

'################################################################################################################
'## 功能：  根据当前位置，同步高亮显示提纲树节点。
'################################################################################################################
Private Sub Editor1_MouseUp(ViewMode As zlRichEditor.ViewModeEnum, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Editor1.ReadOnly Then Exit Sub
    On Error Resume Next
    If ViewMode <> cprNormal Then Exit Sub
    If mblnFmtBrushDown Then
        '格式刷的应用
        With Me.Editor1
            .Tag = "Editor1_MouseUp"
            .ForceEdit = True
            Dim lS As Long, lE As Long
            lS = .Selection.StartPos
            lE = .Selection.EndPos
            If lE > lS + 1 Then
                If .Range(lE - 2, lE) = vbCrLf Then
                    If Not mParaFmt Is Nothing Then
                        '设置段落属性
                        .Selection.Para.Alignment = mParaFmt.Alignment
                        .Selection.Para.FirstLineIndent = mParaFmt.FirstLineIndent
                        .Selection.Para.LeftIndent = mParaFmt.LeftIndent
                        .Selection.Para.SetLineSpacing mParaFmt.LineSpacingRule, mParaFmt.LineSpacing
                        .Selection.Para.ListAlignment = mParaFmt.ListAlignment
                        .Selection.Para.ListStart = mParaFmt.ListStart
                        .Selection.Para.ListTab = mParaFmt.ListTab
                        .Selection.Para.ListType = mParaFmt.ListType
                        .Selection.Para.RightIndent = mParaFmt.RightIndent
                        .Selection.Para.SpaceAfter = mParaFmt.SpaceAfter
                        .Selection.Para.SpaceBefore = mParaFmt.SpaceBefore
                    End If
                End If
            End If
            If Not mFontFmt Is Nothing Then
                '设置字体属性
                .Selection.Font.Bold = mFontFmt.Bold
                .Selection.Font.Italic = mFontFmt.Italic
                .Selection.Font.Name = mFontFmt.Name
                .Selection.Font.Size = mFontFmt.Size
                .Selection.Font.Subscript = mFontFmt.Subscript
                .Selection.Font.Superscript = mFontFmt.Superscript
            End If
            .ForceEdit = False
            .Tag = ""
        End With
        mblnFmtBrushDown = False
        Me.Editor1.OriginRTB.MousePointer = 0
        Exit Sub
    End If

    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bFinded As Boolean, bNeeded As Boolean
    '同步定位提纲
    Document.HighlightCurCompend Editor1, mfrmCompends.Tree

    If txtPenInput.Visible And txtPenInput.Enabled Then txtPenInput.SetFocus: Exit Sub
    If Editor1.SelLength > 0 Then Exit Sub
    bFinded = IsBetweenKeys(Editor1, Editor1.Selection.StartPos + 1, "E", lSS, lSE, lES, lEE, lKey, bNeeded)
    If bFinded Then
        '点击的是元素内部，则表示选择某个选项
        If Me.Document.Elements("K" & lKey).开始版 < Me.Document.目标版本 Or Me.Document.Elements("K" & lKey).终止版 > 0 Then Exit Sub
        If Me.Document.Elements("K" & lKey).输入形态 = 1 Then
            '展开形式的要素录入     '○●□■
            Dim strTmp As String, p As Long, P1 As Long, P2 As Long, blnForce As Boolean, lSP As Long
            With Editor1
                blnForce = .ForceEdit
                .Tag = "Editor1_MouseUp"
                .Freeze
                .ForceEdit = True
                strTmp = .Range(lSE, lES)
                p = .Selection.StartPos: lSP = p - lSE '点中位置，点中位置和关键字起点距离
                If Me.Document.Elements("K" & lKey).要素表示 = 2 Then '单选
                    P1 = .Selection.StartPos - lSE + 1
                    P1 = InStrRev(strTmp, "○", P1)
                    P2 = .Selection.StartPos - lSE + 1
                    P2 = InStrRev(strTmp, "●", P2)
                    If P1 > P2 And P1 > 0 Then
                        strTmp = Replace(strTmp, "●", "○")
                        Mid(strTmp, P1, 1) = "●"
                        .Range(lSE, lES) = strTmp
                    ElseIf P2 > P1 And P2 > 0 Then
                        strTmp = Replace(strTmp, "●", "○")
                        Mid(strTmp, P2, 1) = "○"
                        .Range(lSE, lES) = strTmp
                    End If
                    
                    If Me.Document.Elements("K" & lKey).动态域 = 1 Then '对动态域的单独处理
                        If InStrRev(strTmp, "●") > InStrRev(strTmp, "○") Then '最后一个选项被选中
                            Dim strSin As String
                            strSin = Trim(InputBox("请录入自定义要素选项" & vbCrLf & "最大输入长度200个汉字", "中联软件"))
                            If strSin <> "" Then
                                Me.Document.Elements("K" & lKey).内容文本 = Mid(strTmp, 1, InStrRev(strTmp, "●")) & strSin
                            Else
                                Me.Document.Elements("K" & lKey).内容文本 = Mid(strTmp, 1, InStrRev(strTmp, "●") - 1) & "○自定义"
                            End If
                        Else '最后一项没有被选中,将其变成 ○自定义
                            Me.Document.Elements("K" & lKey).内容文本 = Mid(strTmp, 1, InStrRev(strTmp, "○")) & "自定义"
                        End If
                        Me.Document.Elements("K" & lKey).Refresh Editor1
                    End If
                ElseIf Me.Document.Elements("K" & lKey).要素表示 = 3 Then '多选
                    P1 = .Selection.StartPos - lSE + 1
                    P1 = InStrRev(strTmp, "□", P1)
                    P2 = .Selection.StartPos - lSE + 1
                    P2 = InStrRev(strTmp, "■", P2)
                    If P1 > P2 And P1 > 0 Then
                        Mid(strTmp, P1, 1) = "■"
                        .Range(lSE, lES) = strTmp
                    ElseIf P2 > P1 And P2 > 0 Then
                        Mid(strTmp, P2, 1) = "□"
                        .Range(lSE, lES) = strTmp
                    End If
                    
                    If Me.Document.Elements("K" & lKey).动态域 = 1 And p > InStrRev(strTmp, "■") + lSE - 2 Then '对动态域的单独处理，在最后项位置点击
                        If InStrRev(strTmp, "■") > InStrRev(strTmp, "□") Then '最后一个选项被选中
                            Dim strMul As String
                            strMul = Trim(InputBox("请录入自定义要素选项" & vbCrLf & "最大输入长度200个汉字", "中联软件"))
                            If strMul <> "" Then
                                Me.Document.Elements("K" & lKey).内容文本 = Mid(strTmp, 1, InStrRev(strTmp, "■")) & strMul
                            Else
                                Me.Document.Elements("K" & lKey).内容文本 = Mid(strTmp, 1, InStrRev(strTmp, "■") - 1) & "□自定义"
                            End If
                        Else '最后一项没有被选中,将其变成 ○自定义
                            Me.Document.Elements("K" & lKey).内容文本 = Mid(strTmp, 1, InStrRev(strTmp, "□")) & "自定义"
                        End If
                        Me.Document.Elements("K" & lKey).Refresh Editor1
                    End If
                End If
                
                Call FindKey(Editor1, "E", lKey, lSS, lSE, lES, lEE, bNeeded)
                strTmp = .Range(lSE, lES)
                If (InStr(strTmp, "●") = 0 And Me.Document.Elements("K" & lKey).要素表示 = 2) _
                    Or (InStr(strTmp, "■") = 0 And Me.Document.Elements("K" & lKey).要素表示 = 3) Then '对没选中任何选项情况加波浪线
                    .Range(lSE, lES).Font.Underline = cprWave
                Else
                    .Range(lSE, lES).Font.Underline = cprNone
                End If
                Me.Document.Elements("K" & lKey).内容文本 = strTmp
                
                Call CheckElementLimit(lKey)
                
                Call FindKey(Editor1, "E", lKey, lSS, lSE, lES, lEE, bNeeded)
                lSP = lSE + lSP
                
                .Range(lSP, lSP).Selected
                .ForceEdit = blnForce
                .UnFreeze
                .Tag = ""
            End With
        End If
    End If
End Sub
'################################################################################################################
'## 功能：  用户按Tab键的处理。
'##
'## 说明：  当前是诊治要素，则跳到下一个诊治要素位置处。
'################################################################################################################
Private Sub Editor1_PressTabKey()
    If Editor1.ReadOnly Then Exit Sub
    If Editor1.ViewMode = cprNormal Then
        Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBetweenKeys As Boolean, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
        bBetweenKeys = IsBetweenKeys(Editor1, Editor1.Selection.StartPos + 1, "E", lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBetweenKeys Then
            Call AddUndoPoint  '手动缓存
            bFinded = FindNextKey(Editor1, Editor1.Selection.StartPos + 1, "E", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
            If bFinded Then
                Editor1.Range(lKSE, lKES - Len(Me.Document.Elements("K" & lKey).要素单位)).Selected
            Else
                bFinded = FindNextKey(Editor1, 1, "E", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
                If bFinded Then
                    Editor1.Range(lKSE, lKES - Len(Me.Document.Elements("K" & lKey).要素单位)).Selected
                End If
            End If
            Call ClearNoUseUndoList
        End If
    End If
End Sub

'################################################################################################################
'## 功能：  右键菜单申请
'################################################################################################################
Private Sub Editor1_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, y As Single)
    If Editor1.ReadOnly Then Exit Sub
    Dim Popup As CommandBar
    Dim Control As CommandBarControl, bFinded As Boolean, bOK As Boolean
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean

    Set Popup = cbrThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal Then
            bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
            If bInKeys = False Then Exit Sub
                        If sType = "P" Then
                                Set Control = .Add(xtpControlButton, ID_EDIT_MARKEDPIC, "标记修改(&M)")
                                Control.BeginGroup = True
                                If Me.Document.Pictures("K" & lKey).PictureType = EPRMarkedPicture Then
                                        .Add xtpControlButton, ID_EDIT_OUTERPIC, "底图处理(&D)"
                                End If
                                Popup.ShowPopup
                        End If
        Else
            bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
            If bInKeys Then
                If sType = "D" Then
                    '诊断
                    Set Control = .Add(xtpControlButton, ID_EDIT_DELETE, "删除(&D)"): Control.BeginGroup = True
                    Set Control = .Add(xtpControlButton, conMenu_Tool_Reference, "查阅诊断参考(&R)..."): Control.BeginGroup = True
                    Popup.ShowPopup
                ElseIf sType = "E" Then
                    '诊治要素
                    bFinded = FindPrevKey(Editor1, Editor1.Selection.StartPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
                    If bFinded Then
                        If Me.Document.Compends("K" & lKey).定义提纲ID <> 0 And Me.Editor1.SelLength > 0 Then bOK = True
                    End If
                    If mfrmModElement.Visible Then Exit Sub     '如果在编辑中，那么不能弹出菜单
                    Set Control = .Add(xtpControlButton, ID_EDIT_CUT, "剪切(&X)")
                    Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
                    Set Control = .Add(xtpControlButton, ID_EDIT_PASTE, "粘贴(&V)    ")
                    Set Control = .Add(xtpControlButton, ID_EDIT_DELETE, "删除(&D)")
                    If bOK Then
                        Set Control = .Add(xtpControlButton, ID_EDIT_SAVEASPHRASE, "存为示范词句(&S)..."): Control.BeginGroup = True
                    End If
                    bFinded = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
                    If bFinded Then
                        If sType = "E" Then
                            If Me.Document.Elements("K" & lKey).输入形态 = 0 Then
                                
                                Set Control = .Add(xtpControlButton, ID_ELEMENT_UPDATE, "更新内容(&U)"): Control.BeginGroup = True
                                
                                If Me.Document.Elements("K" & lKey).内容文本 <> "" Then
                                '非展开型要素，可以清除要素内容（主要用于字典项目的文本清空）
                                    Set Control = .Add(xtpControlButton, ID_ELEMENT_CLEAR, "清空内容(&R)")
                                End If
                                
                            End If
                            If Me.Document.Elements("K" & lKey).输入形态 = 0 And InStr(1, gstrPrivsEpr, "保护文本处理") > 0 Then
                                Set Control = .Add(xtpControlButton, ID_ELEMENT_TOSTRING, "转为文本(&T)"): Control.BeginGroup = True
                            End If
                        End If
                    End If
                    Popup.ShowPopup
                End If
            Else
'                Set Control = .Add(xtpControlButton, ID_EDIT_UNDO, "撤销(&U)")
'                Control.BeginGroup = True
'                .Add xtpControlButton, ID_EDIT_REDO, "重做(&R)"
                bFinded = FindPrevKey(Editor1, Editor1.Selection.StartPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
                If bFinded Then
                    If Me.Document.Compends("K" & lKey).定义提纲ID <> 0 And Me.Editor1.SelLength > 0 Then bOK = True
                End If
                Set Control = .Add(xtpControlButton, ID_EDIT_CUT, "剪切(&X)")
                Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
                Set Control = .Add(xtpControlButton, ID_EDIT_PASTE, "粘贴(&V)    ")
                Set Control = .Add(xtpControlButton, ID_EDIT_DELETE, "删除(&D)")
                Set Control = .Add(xtpControlButton, ID_EDIT_COPYSELF, "专用复制(&I)"): Control.BeginGroup = True
                Set Control = .Add(xtpControlButton, ID_EDIT_COPYOUT, "复制到粘贴板(&U)")
                If bOK Then
                    Set Control = .Add(xtpControlButton, ID_EDIT_SAVEASPHRASE, "存为示范词句(&S)..."): Control.BeginGroup = True
                End If
                Set Control = .Add(xtpControlButton, ID_EDIT_SELECTALL, "全选(&A)"): Control.BeginGroup = True
                Popup.ShowPopup
            End If
        End If
    End With
End Sub

'################################################################################################################
'## 功能：  根据当前位置，显示当前行、列位置。
'################################################################################################################
Private Sub Editor1_SelChange(ViewMode As zlRichEditor.ViewModeEnum, ByVal lStart As Long, ByVal lEnd As Long)
    If Me.Document Is Nothing Then Exit Sub
    
    Dim COLOR As OLE_COLOR
    If Me.Document.EditType = cprET_单病历审核 Then
        COLOR = IIf(Editor1.Selection.Font.ForeColor = tomAutoColor, vbBlack, Editor1.Selection.Font.ForeColor)
        If COLOR = tomUndefined Then
            stbThis.Panels(5).Text = "混合文本"
        Else
            If Get终止版(COLOR) > 0 Then
                '删除文本
                stbThis.Panels(5).Text = "第" & Get终止版(COLOR) + 1 & "版删除"
            ElseIf Get开始版(COLOR) > 0 Then
                '新增文本
                stbThis.Panels(5).Text = "第" & Get开始版(COLOR) & "版新增"
            Else
                stbThis.Panels(5).Text = ""
            End If
        End If
    End If

    On Error Resume Next
    If Editor1.InProcessing Then Exit Sub
    stbThis.Panels(2).Text = Editor1.CurrentLine & " 行,  " & Editor1.CurrentColumn & " 列,  共" & Editor1.LineCount & " 行"
    If mblnAutoPageCount Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & ",  共 " & Me.Editor1.PageCount & " 页"
    If Editor1.Tag = "" Then
        Document.HighlightCurCompend Editor1, mfrmCompends.Tree
        '刷新示范词句
        If Not mfrmCompends.Tree.SelectedItem Is Nothing Then
            Call RefSentenceList
        End If
    End If

    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    If Editor1.Tag = "" And Editor1.AuditMode = False Then
        If tblThis.Visible Then
            Editor1.CloseUIInterface
        ElseIf ucPictureEditor1.Visible Then
            Editor1.CloseUIInterface
        ElseIf ucPacsImgCanvas1.Visible Then
            Editor1.CloseUIInterface
        End If
        If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False Then
            
            '查找关键字 ID ！
            bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
            If bInKeys = False Then Exit Sub

            If sType = "T" Then
                '编辑表格
                If Me.Document.Tables("K" & lKey).TableType = tte_报告图片组 Then
                    '报告图片组
                    DkpThis.ShowPane ID_VIEW_PACSPIC
                    tblThis.Tag = lKey
                    Editor1.ShowUIInterface
                    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
                    Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
                ElseIf Me.Document.Tables("K" & lKey).预制提纲ID = 0 Then
                    If Me.Document.Tables("K" & lKey).TableType = tte_医嘱项目组 Then
                    Else
                        '读取数据到表格控件中！
                        tblThis.Tag = lKey
                    End If
                    Editor1.ShowUIInterface
                    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
                    Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
                End If
            ElseIf sType = "P" Then
                '编辑图片
                Editor1.ShowUIInterface
            End If
        End If
    ElseIf Editor1.Tag = "" And Editor1.AuditMode = True And Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False Then
        If tblThis.Visible Then Editor1.CloseUIInterface
        If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False Then
            '查找关键字 ID ！
            bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
            If bInKeys = False Then Exit Sub
            If sType <> "T" Then Exit Sub
            If Me.Document.Tables("K" & lKey).TableType = tte_报告图片组 Then Exit Sub
            If Me.Document.Tables("K" & lKey).TableType = tte_医嘱项目组 Then Exit Sub
            If Me.Document.Tables("K" & lKey).预制提纲ID = 0 Then
                '读取数据到表格控件中！
                tblThis.Tag = lKey
                Editor1.ShowUIInterface
                Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
                Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
            End If
        End If
    End If
End Sub

Private Function ReadTableToUI(ByRef oTable As cEPRTable) As Boolean
    On Error Resume Next
    If oTable Is Nothing Then ReadTableToUI = False: Exit Function
'    Dim t1 As Date, T2 As Date
'    t1 = Timer

    Dim i As Long, j As Long, strMerge As String, R1 As Long, C1 As Long, R2 As Long, C2 As Long
    Dim T As Variant, strColWidth As String, lKey As Long

    tblThis.Redraw = False
    tblThis.SingleClickEdit = False
    tblThis.HighlightMode = HMFilledRectAlpha
    tblThis.BorderWidth = oTable.BorderWidth
    tblThis.AutoHeight = oTable.AutoHeight
    tblThis.Init oTable.Rows, oTable.Cols
    tblThis.ExtendTag = oTable.ExtendTag
    tblThis.UserTag = oTable.标记
    strColWidth = oTable.ColWidthString
    T = Split(strColWidth, "|")
    On Error Resume Next
    If UBound(T) = -1 Then
        If oTable.Rows > 0 Then
            For i = 1 To oTable.Cols
                tblThis.ColWidth(i) = oTable.Cell(1, i).Width
            Next
        End If
    Else
        For i = 0 To UBound(T)
            tblThis.ColWidth(i + 1) = Val(T(i))
        Next
    End If

    For i = 1 To oTable.Rows
        tblThis.ROWHEIGHT(i) = oTable.Cell(i, 1).Height
        For j = 1 To oTable.Cols
            lKey = tblThis.CellKey(i, j)
            With oTable.Cell(i, j)
                tblThis.Cells("K" & lKey).Text = oTable.Cell(i, j).内容文本
                tblThis.Cells("K" & lKey).Margin = .Margin
                tblThis.Cells("K" & lKey).Width = .Width
                tblThis.Cells("K" & lKey).Height = .Height
'                tblThis.Cells("K" & lKey).MergeInfo = .MergeNo
                tblThis.Cells("K" & lKey).SingleLine = .SingleLine
                tblThis.Cells("K" & lKey).ForeColor = .ForeColor
                tblThis.Cells("K" & lKey).BackColor = .BackColor
                tblThis.Cells("K" & lKey).GridLineColor = .GridLineColor
                tblThis.Cells("K" & lKey).GridLineWidth = .GridLineWidth
                tblThis.Cells("K" & lKey).FixedWidth = .FixedWidth
                tblThis.Cells("K" & lKey).AutoHeight = .AutoHeight
                tblThis.Cells("K" & lKey).FontName = .FontName
                tblThis.Cells("K" & lKey).FontSize = .FontSize
                tblThis.Cells("K" & lKey).FontBold = .FontBold
                tblThis.Cells("K" & lKey).FontItalic = .FontItalic
                tblThis.Cells("K" & lKey).FontStrikeout = .FontStrikeout
                tblThis.Cells("K" & lKey).FontUnderline = .FontUnderline
                tblThis.Cells("K" & lKey).FontWeight = .FontWeight
                tblThis.Cells("K" & lKey).FormatString = .FormatString
                tblThis.Cells("K" & lKey).Indent = .Indent
                tblThis.Cells("K" & lKey).HAlignment = .HAlignment
                tblThis.Cells("K" & lKey).VAlignment = .VAlignment
                tblThis.Cells("K" & lKey).Protected = .Protected
                If oTable.Cell(i, j).ElementKey > 0 Then
                    tblThis.Cells("K" & lKey).ToolTipText = oTable.Elements("K" & oTable.Cell(i, j).ElementKey).要素名称
                    tblThis.Cells("K" & lKey).Tag = oTable.Cell(i, j).ElementKey
                End If
                If .PictureKey > 0 Then
                    oTable.Pictures("K" & .PictureKey).Row = i
                    oTable.Pictures("K" & .PictureKey).Col = j
                    Set tblThis.Cells("K" & lKey).Picture = oTable.Pictures("K" & .PictureKey).DrawFinalPic(oTable)
                    tblThis.Cells("K" & lKey).Tag = oTable.Cell(i, j).PictureKey
                End If
            End With
        Next
    Next

    For i = 1 To oTable.Cells.Count
        strMerge = oTable.Cells(i).MergeNo              '恢复单元格的合并
        If strMerge <> "" Then
            R1 = Val(Left(strMerge, 4))
            C1 = Val(Mid(strMerge, 5, 4))
            R2 = Val(Mid(strMerge, 9, 4))
            C2 = Val(Mid(strMerge, 13))
            tblThis.MergeCells R1, C1, R2, C2, False
        End If
    Next

    tblThis.ShowToolTipText = True
    tblThis.MinRowHeight = 300
    tblThis.Redraw = True
    tblThis.Refresh
    tblThis.FixCellsWidth
    If (Not tblThis.AutoHeight) Then tblThis.Height = oTable.Height

    Editor1.ResizeUIInterface tblThis.Width, tblThis.Height
    tblThis.Refresh True, False

'    T2 = Timer
'    Debug.Print "读取耗时：" & Format(T2 - t1, "0.00000000") & ",单元格总数：" & tblThis.Cells.Count
End Function

Private Function SaveUIToTable(ByRef oTable As cEPRTable, Optional ByVal bFirst As Boolean) As Boolean
    If oTable Is Nothing Then SaveUIToTable = False: Exit Function
    Dim strColWidth As String

    Dim i As Long, j As Long, lKey As Long
    For i = 1 To tblThis.ColCount
        If i = 1 Then
            strColWidth = tblThis.ColWidth(i)
        Else
            strColWidth = strColWidth & "|" & tblThis.ColWidth(i)
        End If
    Next

    oTable.Rows = tblThis.RowCount
    oTable.Cols = tblThis.ColCount
    oTable.ColWidthString = strColWidth
'    oTable.Width = tblThis.Width
    oTable.Height = tblThis.Height
    oTable.SingleLine = tblThis.SingleLine
    oTable.AlternateRowBackColor = tblThis.AlternateRowBackColor
    oTable.BackColor = tblThis.BackColor
    oTable.GridLineColor = tblThis.GridLineColor
    oTable.GridLineWidth = tblThis.GridLineWidth
    oTable.BorderColor = tblThis.BorderColor
    oTable.BorderWidth = tblThis.BorderWidth
    oTable.ForeColor = tblThis.ForeColor
    oTable.FontQuality = tblThis.FontQuality
    oTable.AutoHeight = tblThis.AutoHeight
    oTable.WordEllipsis = tblThis.WordEllipsis
    oTable.CellMargin = tblThis.CellMargin
    oTable.CellIndent = tblThis.CellIndent
    oTable.ExtendTag = tblThis.ExtendTag
    oTable.标记 = tblThis.UserTag
    For i = 1 To oTable.Rows
        For j = 1 To oTable.Cols
            If oTable.Cell(i, j) Is Nothing Then
                lKey = 0
                Call oTable.Cells.Add(lKey, i, j)
                oTable.Cell(i, j).ID = 0
                oTable.Cell(i, j).开始版 = Me.Document.目标版本
            End If
            lKey = tblThis.CellKey(i, j)
            If lKey > 0 And Not oTable.Cell(i, j) Is Nothing Then
                With oTable.Cell(i, j)
                    If .内容文本 <> tblThis.Cells("K" & lKey).Text And oTable.TableType = tte_默认 And Me.Document.EditType = cprET_单病历审核 Then
                        If .开始版 <> Me.Document.目标版本 And .ID <> 0 Then .ID = 0
                        .开始版 = Me.Document.目标版本
                    End If
                    .内容文本 = tblThis.Cells("K" & lKey).Text
                    .Margin = tblThis.Cells("K" & lKey).Margin
'                    .Width = tblThis.Cells("K" & lKey).Width
'                    .Height = tblThis.Cells("K" & lKey).Height
                    .Width = tblThis.ColWidth(j)
                    .Height = tblThis.ROWHEIGHT(i)
                    .MergeNo = tblThis.Cells("K" & lKey).MergeInfo
                    .SingleLine = tblThis.Cells("K" & lKey).SingleLine
                    .ForeColor = tblThis.Cells("K" & lKey).ForeColor
                    .BackColor = tblThis.Cells("K" & lKey).BackColor
                    .GridLineColor = tblThis.Cells("K" & lKey).GridLineColor
                    .GridLineWidth = tblThis.Cells("K" & lKey).GridLineWidth
                    .FixedWidth = tblThis.Cells("K" & lKey).FixedWidth
                    .AutoHeight = tblThis.Cells("K" & lKey).AutoHeight
                    .FontName = tblThis.Cells("K" & lKey).FontName
                    .FontSize = tblThis.Cells("K" & lKey).FontSize
                    .FontBold = tblThis.Cells("K" & lKey).FontBold
                    .FontItalic = tblThis.Cells("K" & lKey).FontItalic
                    .FontStrikeout = tblThis.Cells("K" & lKey).FontStrikeout
                    .FontUnderline = tblThis.Cells("K" & lKey).FontUnderline
                    .FontWeight = tblThis.Cells("K" & lKey).FontWeight
                    .FormatString = tblThis.Cells("K" & lKey).FormatString
                    .Indent = tblThis.Cells("K" & lKey).Indent
                    .VAlignment = tblThis.Cells("K" & lKey).VAlignment
                    .HAlignment = tblThis.Cells("K" & lKey).HAlignment
                    .Protected = tblThis.Cells("K" & lKey).Protected
                    If tblThis.Cells("K" & lKey).Picture Is Nothing Then
                        .ElementKey = Val(tblThis.Cells("K" & lKey).Tag)
                        .PictureKey = 0
                    Else
                        .ElementKey = 0
                        .PictureKey = Val(tblThis.Cells("K" & lKey).Tag)
                    End If
                End With
            End If
        Next j
    Next i

    '这里要限制图片宽度不能超过页面宽度
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, bFinded As Boolean, bNeeded As Boolean
    Dim lW As Long
    bFinded = FindKey(Editor1, "T", oTable.Key, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        lW = Me.Editor1.PaperWidth - Me.Editor1.MarginLeft - Me.Editor1.MarginRight - Me.ScaleX(Me.Editor1.Range(lSE, lES).Para.LeftIndent + Me.Editor1.Range(lSE, lES).Para.FirstLineIndent, vbPixels, vbTwips) - 130
        picTMP.Width = IIf(tblThis.Width > lW, lW, tblThis.Width)
    Else
        picTMP.Width = tblThis.Width
    End If
    picTMP.Height = tblThis.Height

    oTable.Width = picTMP.Width '保存实际表格宽度！！！！！

    tblThis.DrawToDC picTMP.hDC
    picTMP.Picture = picTMP.Image
    If bFirst Then
        Dim frmT As New frmTablePicCreator
        Me.Document.Tables("K" & tblThis.Tag).InsertIntoEditor Me.Editor1, , frmT.GetFinalPic(Me.Document.Tables("K" & tblThis.Tag)), True
        Unload frmT
        Set frmT = Nothing
    Else
        oTable.Refresh Editor1, picTMP.Picture, True
    End If
End Function

Private Sub Form_Activate()
    On Error Resume Next
    If tblThis.Visible Then
        tblThis.SetFocus
    ElseIf ActiveControl Is Editor1 Then
        If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
    End If
     
    Err.Clear
End Sub

'################################################################################################################
'## 功能：  窗体初始化
'################################################################################################################
Private Sub Form_Load()
    Dim i As Long, j As Long

    mblnAutosave = (zlDatabase.GetPara("AutoSave", glngSys, 1070, 1) = "1")
    mlngUndoLimit = zlDatabase.GetPara("UndoLimit", glngSys, 1070, 20)
    mlngSaveInterval = zlDatabase.GetPara("SaveInterval", glngSys, 1070, 60)
    mblnAutoSaveEPR = (zlDatabase.GetPara("AutoSaveEPR", glngSys, 1070, 0) = "1")
    mlngSaveIntervalEPR = zlDatabase.GetPara("SaveIntervalEPR", glngSys, 1070, 5)
    mblnAutoPageCount = (zlDatabase.GetPara("AutoPageCount", glngSys, 1070, 0) = "1")
    mblnAutoPageNote = (zlDatabase.GetPara("AutoPageNote", glngSys, 1070, 0) = "1")
    mintSharePages = zlDatabase.GetPara("SharePageCount", glngSys, 1070, 5)
    mblnSignAutoAlter = (zlDatabase.GetPara("签名自动位移", glngSys, 1070, 0) = "1")
            
    ReDim UndoList(1 To 1) As UndoInfo
    p_Undo = 0
    mblnChange = False


    '## 菜单初始化
    Dim cbrMenu As CommandBarPopup                      '主菜单
    Dim cbpPopup As CommandBarPopup                     '下拉对象
    Dim cbpPopupSub As CommandBarPopup                  '临时对象
    Dim objControl As CommandBarControl                 '工具栏控件
    Dim objCustControl As CommandBarControlCustom       '自定义控件
    Dim Combo As CommandBarComboBox                     '工具栏下拉框控件
    Dim cbrBar As CommandBar                           '工具栏
    Dim cbrCustom As CommandBarControlCustom

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbrThis.Icons = gfrmPublic.ImageManager.Icons
    cbrThis.Options.ShowExpandButtonAlways = False
    cbrThis.EnableCustomization (False)
    cbrThis.Options.UseDisabledIcons = True
    cbrThis.Options.AlwaysShowFullMenus = True
    cbrThis.StatusBar.Visible = False
    cbrThis.ActiveMenuBar.Title = "菜单栏"

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "文件(&F)"): cbrMenu.ID = ID_Main_FILE
    With cbrMenu.CommandBar.Controls
        .Add xtpControlButton, ID_FILE_CLEAR, "清空(&C)"

        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE, "保存(&S)")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE_QUIT, "保存退出(&Q)")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVEASEPRDEMO, "另存为范文(&D)...")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVEASSEGMENT, "另存为片段(&G)...")
        
        Set cbpPopup = .Add(xtpControlPopup, 0, "导出(&A)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILE_SAVEAS, "导出为&RTF文件..."
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILE_EXPORTTOXML, "导出为X&ML文件..."

        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORTFROMXML, "从XM&L文件导入")
        Set objControl = .Add(xtpControlButton, ID_FILE_PAGESETUP, "页面设置(&U)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FILE_PRINTPREVIEW, "打印预览(&V)"
        .Add xtpControlButton, ID_FILE_PRINT, "打印(&P)..."
        .Add xtpControlButton, ID_FILE_PRINTINWORD, "通过Word打印(&W)"

        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&T)"): objControl.BeginGroup = True

        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "退出(&X)")
        objControl.BeginGroup = True
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "编辑(&E)"): cbrMenu.ID = ID_Main_EDIT
    With cbrMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_EDIT_UNDO, "撤销(&U)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_REDO, "重做(&R)"
'
        Set objControl = .Add(xtpControlButton, ID_EDIT_CUT, "剪切(&X)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_COPY, "复制(&C)"
        
        .Add xtpControlButton, ID_EDIT_PASTE, "粘贴(&V)"
        Set objControl = .Add(xtpControlButton, ID_EDIT_COPYSELF, "专用复制(&I)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_COPYOUT, "复制到粘贴板(&U)"
    
        Set cbpPopup = .Add(xtpControlPopup, 0, "提纲(&M)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_EDIT_REFCOMPEND, "刷新提纲(&R)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_EDIT_ADDCOMPEND, "新增提纲(&A)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_EDIT_DELCOMPEND, "删除提纲(&D)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_EDIT_MODCOMPEND, "修改提纲(&M)"

        Set cbpPopup = .Add(xtpControlPopup, 0, "签名与修订(&S)")
        cbpPopup.BeginGroup = True
'            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_PATISIGN, "患者签名(&S)"): objControl.STYLE = xtpButtonIconAndCaption
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_SIGN, "签名(&S)"): objControl.STYLE = xtpButtonIconAndCaption
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_UNTREAD, "回退(&C)"): objControl.STYLE = xtpButtonIconAndCaption
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_SIGN_QUIT, "签名退出(&Q)"): objControl.STYLE = xtpButtonIconAndCaption
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_REVISION_PREV, "前一处修订(&P)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_REVISION_NEXT, "后一处修订(&N)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_REVISION_RESET, "取消所选修订(&E)")

        Set objControl = .Add(xtpControlButton, ID_EDIT_DELETE, "删除(&D)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_SELECTALL, "全选(&A)"

        Set objControl = .Add(xtpControlButton, ID_EDIT_FIND, "查找(&F)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_FINDNEXT, "查找下一个(&N)"
        .Add xtpControlButton, ID_EDIT_REPLACE, "替换(&R)..."
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "视图(&V)"): cbrMenu.ID = ID_Main_VIEW
    With cbrMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_VIEW_STRUCTURE, "文档结构图(&D)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_VIEW_PHRASEDEMO, "示范词句列表(&S)"
        .Add xtpControlButton, ID_VIEW_SEGMENT, "示范片段列表(&G)"
        .Add xtpControlButton, ID_VIEW_PACSPIC, "报告图列表(&P)"
        .Add xtpControlButton, ID_VIEW_HISTORYWINDOW, "历史内容列表(&P)"
        .Add xtpControlButton, ID_VIEW_HISTORYREPORT, "历史报告列表(&R)"
        
        Set cbpPopup = .Add(xtpControlPopup, 0, "工具栏(&T)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, "工具栏列表"

        Set objControl = .Add(xtpControlButton, ID_VIEW_HEADFOOT, "页眉页脚(&H)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_VIEW_CHARCOUNT, "字数统计(&N)..."
'        .Add xtpControlButton, ID_VIEW_GRID, "网格线(&G)"
        Set objControl = .Add(xtpControlButton, ID_VIEW_RULER, "标尺(&R)")
        objControl.Checked = True
        .Add xtpControlButton, ID_VIEW_PENWINDOW, "手写输入窗口(&W)"
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "插入(&I)"): cbrMenu.ID = ID_Main_INSERT
    With cbrMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_INSERT_DATETIME, "日期和时间(&D)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_INSERT_SPECIALCHAR, "特殊符号(&S)..."

        Set objControl = .Add(xtpControlButton, ID_INSERT_PICTURE, "图片(&P)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_TABLE_INSERTTABLE, "表格(&T)"
        .Add xtpControlButton, ID_INSERT_ELEMENT, "要素(&E)"
        .Add xtpControlButton, ID_EDIT_ADDCOMPEND, "提纲(&C)"
        .Add xtpControlButton, ID_INSERT_PACSPIC, "报告图(&R)"
        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORT, "历史文件(&H)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_INSERT_EPRDEMO, "导入范文(&F)..."
        .Add xtpControlButton, ID_INSERT_DOCADVISE, "本次医嘱(&A)"
        .Add xtpControlButton, ID_DIAGNOSIS, "诊断(&D)"
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "格式(&O)"): cbrMenu.ID = ID_Main_FORMAT
    With cbrMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_FORMAT_FONT, "字体(&F)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FORMAT_PARA, "段落(&P)..."
        Set cbpPopup = .Add(xtpControlPopup, ID_FORMAT_BACKGROUND, "背景色(&K)")
        cbpPopup.BeginGroup = True
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, ID_FORMAT_BACKGROUND, "")
        objCustControl.Handle = ColorPaperBackColor.hwnd

        Set objControl = .Add(xtpControlButton, ID_FORMAT_BOLD, "粗体(&B)")
        objControl.BeginGroup = True
'        .Add xtpControlButton, ID_FORMAT_ITALIC, "斜体(&I)"
        .Add xtpControlButton, ID_FORMAT_SUPER, "上标(&R)"
        .Add xtpControlButton, ID_FORMAT_SUB, "下标(&S)"
        .Add xtpControlButton, ID_FORMAT_PROTECT, "保护(&P)"

        Set cbpPopup = .Add(xtpControlPopup, ID_FORMAT_UNDERLINE, "下划线"): cbpPopup.ID = ID_FORMAT_UNDERLINE
        cbpPopup.CommandBar.SetPopupToolBar True
        cbpPopup.CommandBar.SetIconSize 60, 8
        cbpPopup.CommandBar.Width = 60
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_NONE, "<无下划线>"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_THIN, "细线"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_THICK, "粗线"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_WAVE, "波浪线"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DOT, "点线"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASH, "虚线"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT, "点划线"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT2, "双点划线"

        Set cbpPopup = .Add(xtpControlPopup, 0, "对齐方式(&A)")
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_ALIGNLEFT, "左对齐(&L)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_ALIGNCENTER, "居中对齐(&C)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_ALIGNRIGHT, "右对齐(&R)"

        Set cbpPopup = .Add(xtpControlPopup, 0, "项目符号与编号(&E)")
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTNONE, "无"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTBULLETS, "项目符号(·)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTARABIC, "阿拉伯数字(1,2,3,...)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTLCHAR, "小写字母(a,b,c,...)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTUCHAR, "大写字母(A,B,C,...)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTLROME, "小写罗马数字(i,ii,iii,...)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTUROME, "大写罗马数字(I,II,III,...)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_LISTSETUP, "自定义格式...")
        objControl.BeginGroup = True

        Set cbpPopup = .Add(xtpControlPopup, ID_FORMAT_SPACE, "间距(&L)"): cbpPopup.ID = ID_FORMAT_SPACE
        Set cbpPopupSub = cbpPopup.CommandBar.Controls.Add(xtpControlSplitButtonPopup, ID_FORMAT_LINESPACE, "行间距(&L)")
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE1, "1.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE2, "1.3"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE3, "1.5"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE4, "2.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE5, "2.5"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE6, "3.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE7, "其他..."

        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_SPACEBEFORE, "段前间距(&B)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_SPACEAFTER, "段后间距(&A)"

        Set cbpPopup = .Add(xtpControlPopup, 0, "缩进(&D)")
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_FIRSTINDENT, "首行缩进(&F)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_FIRSTHUNGING, "首行悬挂(&H)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_INDENTDECREASE, "减少缩进量(&D)")
        objControl.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_INDENTINCREASE, "增加缩进量(&I)"
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "表格(&T)"): cbrMenu.ID = ID_Main_TABLE
    With cbrMenu.CommandBar.Controls
        Set cbpPopup = .Add(xtpControlPopup, 0, "插入(&I)")
'        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTTABLE, "表格(&T)"): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTCOLLEFT, "列(在左侧)(&L)"): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTCOLRIGHT, "列(在右侧)(&R")
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTROWUP, "行(在上方)(&A)"): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTROWDOWN, "行(在下方)(&B)")
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTINHERITROW, "插入继承行(&I)")

        Set cbpPopup = .Add(xtpControlPopup, 0, "删除(&D)")
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_DELETETABLE, "表格(&T)"): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_DELETECOL, "列(&C)"): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_DELETEROW, "行(&R)")

        Set cbpPopup = .Add(xtpControlPopup, 0, "格式(&F)")
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_FORMATCELL, "单元格(&E)..."): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_SAMECOLWIDTH, "相同列宽(&C)")
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_MERGE, "合并(&M)"): objControl.BeginGroup = True
        Set cbpPopup = cbpPopup.CommandBar.Controls.Add(xtpControlPopup, ID_TABLE_CELLALIGNMENT, "单元格对齐方式")
        cbpPopup.CommandBar.SetTearOffPopup "单元格对齐方式", ID_TABLE_CELLALIGNMENT, 100
        cbpPopup.CommandBar.SetPopupToolBar True
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Width = 70
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT1, "靠上左对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT2, "靠上居中"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT3, "靠上右对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT4, "中部左对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT5, "中部居中"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT6, "中部右对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT7, "靠下左对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT8, "靠下居中"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT9, "靠下右对齐"
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "帮助(&H)"): cbrMenu.ID = ID_Main_HELP
    With cbrMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "帮助主题(&H)")
        objControl.BeginGroup = True
        Set cbpPopupSub = .Add(xtpControlPopup, 0, "&Web上的" & gstrProductName)
        objControl.BeginGroup = True
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_HELP_ONLINE, gstrProductName & "主页(&H)"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_HELP_WEBFORUM, gstrProductName & "论坛(&F)"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_HELP_CONTACT, "发送反馈(&M)"
        Set objControl = .Add(xtpControlButton, ID_HELP_ABOUT, "关于(&A)...")
        objControl.BeginGroup = True
    End With

    '## 工具栏初始化

    Set cbrBar = cbrThis.Add("常用", xtpBarTop): cbrBar.BarId = ID_BAR_NORMAL
    cbrBar.EnableDocking xtpFlagStretched
    With cbrBar.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_CLEAR, "清空")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE, "保存")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE_QUIT, "保存退出")
        

        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "打印")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FILE_PRINTPREVIEW, "打印预览"

        Set objControl = .Add(xtpControlButton, ID_EDIT_CUT, "剪切")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_COPY, "复制"
        .Add xtpControlButton, ID_EDIT_COPYSELF, "专用复制"
        .Add xtpControlButton, ID_EDIT_PASTE, "粘贴"
        .Add xtpControlButton, ID_EDIT_FORMATBRUSH, "格式刷"

        Set objControl = .Add(xtpControlButton, ID_EDIT_UNDO, "撤销")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_REDO, "重做"

        Set objControl = .Add(xtpControlButton, ID_INSERT_DATETIME, "插入日期与时间")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_INSERT_DATE, "插入日期"
        .Add xtpControlButton, ID_INSERT_TIME, "插入时间"
        .Add xtpControlButton, ID_INSERT_SPECIALCHAR, "插入特殊符号"

        Set objControl = .Add(xtpControlButton, ID_INSERT_PICTURE, "插入图形"): objControl.BeginGroup = True
        .Add xtpControlButton, ID_INSERT_ELEMENT, "诊治要素"
        .Add xtpControlButton, ID_INSERT_PACSPIC, "插入报告图"

        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORT, "历史文件"): objControl.BeginGroup = True
        .Add xtpControlButton, ID_INSERT_EPRDEMO, "导入范文"
        .Add xtpControlButton, ID_INSERT_DOCADVISE, "本次医嘱"
        .Add xtpControlButton, ID_DIAGNOSIS, "插入诊断"

        Set cbpPopup = .Add(xtpControlPopup, 0, "提纲"): cbpPopup.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_EDIT_REFCOMPEND, "刷新提纲"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_EDIT_ADDCOMPEND, "新增提纲"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_EDIT_DELCOMPEND, "删除提纲"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_EDIT_MODCOMPEND, "修改提纲"): objControl.STYLE = xtpButtonIconAndCaption

        Set objControl = .Add(xtpControlButton, ID_VIEW_STRUCTURE, "文档结构图"): objControl.BeginGroup = True
        .Add xtpControlButton, ID_VIEW_PHRASEDEMO, "示范词句列表"
        .Add xtpControlButton, ID_VIEW_SEGMENT, "示范片段列表"
        .Add xtpControlButton, ID_VIEW_PACSPIC, "报告图列表"
        .Add xtpControlButton, ID_VIEW_HISTORYWINDOW, "历史内容列表"
        
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "zlRichEMR 帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_COMMON_CANCEL, "退出(&Q)"): objControl.BeginGroup = True: objControl.STYLE = xtpButtonIconAndCaption
     
        Set objControl = .Add(xtpControlLabel, 99999901, "反馈内容:")
        objControl.flags = xtpFlagRightAlign
        objControl.Visible = False
        Set cbrCustom = .Add(xtpControlCustom, 99999902, "反馈内容")
        cbrCustom.Handle = Me.txtFeedBack.hwnd
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Visible = False
        
        Set objControl = .Add(xtpControlLabel, 99999903, "处理说明:")
        objControl.flags = xtpFlagRightAlign
        objControl.Visible = False
        Set cbrCustom = .Add(xtpControlCustom, 99999904, "处理说明")
        cbrCustom.Handle = Me.txtContent.hwnd
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Visible = False
    End With
    
    If Not gobjPlugIn Is Nothing Then '插件菜单及工具条
        Dim strFunc As String, lngFuncID As Long, strFuncName As String
        strFunc = gobjPlugIn.GetFuncNames(glngSys, 1070)
        '增加菜单部分
        If strFunc <> "" Then
            With cbrThis.ActiveMenuBar.Controls
                Set cbrMenu = .Find(, ID_Main_HELP)
                If Not cbrMenu Is Nothing Then
                    i = cbrMenu.Index
                Else
                    i = -1
                End If
                Set cbrMenu = .Add(xtpControlPopup, conMenu_Tool_PlugIn, "扩展功能", i, False)
                With cbrMenu.CommandBar.Controls
                    For i = 0 To UBound(Split(strFunc, ","))
                        lngFuncID = conMenu_Tool_PlugIn_Item + i + 1
                        strFuncName = Split(strFunc, ",")(i)
                        
                        If UCase(strFuncName) Like UCase("Auto:*") Then
                            strFuncName = Mid(strFuncName, 6)
                        End If
                        
                        Set objControl = .Add(xtpControlButton, lngFuncID, strFuncName)
                        If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                        objControl.IconId = conMenu_Tool_PlugIn_Item
                        objControl.Parameter = strFuncName
                    Next
                End With
                
                If .Count > 1 Then .Item(2).BeginGroup = True
            End With
            '增加工具栏部分
            Set cbrBar = cbrThis(2)
            Set objControl = cbrBar.FindControl(, ID_HELP_CONTENT)
            If Not objControl Is Nothing Then
                objControl.BeginGroup = True
                i = objControl.Index
            Else
                i = -1
            End If
            
            Set cbpPopup = cbrBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "扩展功能", i, False)
            cbpPopup.ID = conMenu_Tool_PlugIn
            cbpPopup.IconId = conMenu_Tool_PlugIn
            cbpPopup.BeginGroup = True
            With cbpPopup.CommandBar.Controls
                For i = 0 To UBound(Split(strFunc, ","))
                    lngFuncID = conMenu_Tool_PlugIn_Item + i + 1
                    strFuncName = Split(strFunc, ",")(i)
                    
                    If UCase(strFuncName) Like UCase("Auto:*") Then
                        strFuncName = Mid(strFuncName, 6)
                    End If
                    
                    Set objControl = .Add(xtpControlButton, lngFuncID, strFuncName)
                    If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                    objControl.IconId = conMenu_Tool_PlugIn_Item
                    objControl.Parameter = strFuncName
                Next
            End With
        End If
    End If

    Set cbrBar = cbrThis.Add("格式", xtpBarTop): cbrBar.BarId = ID_BAR_FORMAT
    With cbrBar.Controls
        .Add xtpControlButton, ID_FORMAT_STYLEWINDOW, "样式窗格"

        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_STYLE, "常用样式")
        Dim rs As New ADODB.Recordset
        Set rs = zlDatabase.OpenSQLRecord("select 名称 from 病历常用样式 order by 编号", "提取信息", "")
        i = 0
        Do While Not rs.EOF
            i = i + 1
            Combo.AddItem rs("名称")
            If rs("名称") = "正文" Then Combo.ListIndex = i
            rs.MoveNext
        Loop
        Combo.AddItem "其他..."
        Combo.Width = 50
        Combo.DropDownWidth = 220
        Combo.DropDownListStyle = True

        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_FONTNAME, "字体名称")
        Combo.BeginGroup = True
        For i = 0 To gfrmPublic.cmbFont.ListCount - 1
            Combo.AddItem gfrmPublic.cmbFont.List(i), i + 1
            If gfrmPublic.cmbFont.List(i) = "宋体" Then Combo.ListIndex = i + 1
        Next
        Combo.Width = 90
        Combo.DropDownWidth = 250
        Combo.DropDownListStyle = True

        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_FONTSIZE, "字体尺寸")
        '字号列表
        Combo.AddItem "初号", 1
        Combo.AddItem "小初", 2
        Combo.AddItem "一号", 3
        Combo.AddItem "小一", 4
        Combo.AddItem "二号", 5
        Combo.AddItem "小二", 6
        Combo.AddItem "三号", 7
        Combo.AddItem "小三", 8
        Combo.AddItem "四号", 9
        Combo.AddItem "小四", 10
        Combo.AddItem "五号", 11
        Combo.AddItem "小五", 12
        Combo.AddItem "六号", 13
        Combo.AddItem "小六", 14
        Combo.AddItem "七号", 15
        Combo.AddItem "八号", 16
        Combo.AddItem 5, 17
        Combo.AddItem 5.5, 18
        Combo.AddItem 6.5, 19
        Combo.AddItem 7.5, 20
        Combo.AddItem 8, 21
        Combo.AddItem 9, 22
        Combo.AddItem 10, 23
        Combo.AddItem 10.5, 24
        Combo.AddItem 11, 25
        Combo.AddItem 12, 26
        Combo.AddItem 14, 27
        Combo.AddItem 16, 28
        Combo.AddItem 18, 29
        Combo.AddItem 20, 30
        Combo.AddItem 22, 31
        Combo.AddItem 24, 32
        Combo.AddItem 26, 33
        Combo.AddItem 28, 34
        Combo.AddItem 36, 35
        Combo.AddItem 48, 36
        Combo.AddItem 72, 37

        Combo.ListIndex = 12
        Combo.Width = 50
        Combo.DropDownWidth = 80
        Combo.DropDownListStyle = True

        Set objControl = .Add(xtpControlButton, ID_FORMAT_BOLD, "粗体")
        objControl.BeginGroup = True

        Set cbpPopupSub = .Add(xtpControlSplitButtonPopup, ID_FORMAT_UNDERLINE, "下划线"): cbpPopupSub.ID = ID_FORMAT_UNDERLINE
        cbpPopupSub.CommandBar.SetPopupToolBar True
        cbpPopupSub.CommandBar.SetIconSize 60, 8
        cbpPopupSub.CommandBar.Width = 60
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_THIN, "细线"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_THICK, "粗线"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_WAVE, "波浪线"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DOT, "点线"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASH, "虚线"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT, "点划线"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT2, "双点划线"

        .Add xtpControlButton, ID_FORMAT_SUPER, "上标"
        .Add xtpControlButton, ID_FORMAT_SUB, "下标"
        .Add xtpControlButton, ID_FORMAT_PROTECT, "保护"

        Set objControl = .Add(xtpControlButton, ID_FORMAT_ALIGNLEFT, "左对齐"): objControl.BeginGroup = True
        .Add xtpControlButton, ID_FORMAT_ALIGNCENTER, "居中"
        .Add xtpControlButton, ID_FORMAT_ALIGNRIGHT, "右对齐"

        Set cbpPopupSub = .Add(xtpControlSplitButtonPopup, ID_FORMAT_LINESPACE, "行距")
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE1, "1.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE2, "1.3"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE3, "1.5"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE4, "2.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE5, "2.5"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE6, "3.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE7, "其他..."

        Set objControl = .Add(xtpControlButton, ID_FORMAT_INDENTDECREASE, "减少缩进量"): objControl.BeginGroup = True
        objControl.Visible = False
        Set objControl = .Add(xtpControlButton, ID_FORMAT_INDENTINCREASE, "增加缩进量")
        objControl.Visible = False

        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_FORMAT_HIGHLIGHT, "突出显示"): cbpPopup.BeginGroup = True
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorHighlight.hwnd
    End With
    cbpPopupSub.CommandBar.FindControl(, ID_FORMAT_LINESPACE1).Checked = True

    Set cbrBar = cbrThis.Add("签名", xtpBarTop): cbrBar.BarId = ID_BAR_SIGN
    cbrBar.EnableDocking xtpFlagHideWrap
    cbrBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    With cbrBar.Controls
        Set objControl = .Add(xtpControlButton, ID_REVISION_PREV, "前一处修订")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_REVISION_NEXT, "后一处修订"
        Set objControl = .Add(xtpControlButton, ID_REVISION_RESET, "清除修订")
        objControl.STYLE = xtpButtonIconAndCaption

        
        Set objControl = .Add(xtpControlButton, ID_PATISIGN, "患者签名")
        objControl.STYLE = xtpButtonIconAndCaption
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_SIGN, "签名")
        objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_UNTREAD, "回退")
        objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_SIGN_QUIT, "签名退出")
        objControl.STYLE = xtpButtonIconAndCaption
    End With

    Set cbrBar = cbrThis.Add("表格", xtpBarTop): cbrBar.BarId = ID_BAR_TABLE
    cbrBar.EnableDocking xtpFlagHideWrap
    With cbrBar.Controls
        Set objControl = .Add(xtpControlButton, ID_INSERT_TABLE, "插入表格"): objControl.BeginGroup = True

        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_FILLCOLOR, "填充颜色")
        cbpPopup.BeginGroup = True
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorFillColor.hwnd

        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_FORMAT_FORECOLOR, "字体颜色")
        cbpPopup.BeginGroup = True
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorForeColor.hwnd

        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_TABLE_CELLALIGNMENT, "单元格对齐方式")
        cbpPopup.CommandBar.SetTearOffPopup "单元格对齐方式", ID_TABLE_CELLALIGNMENT, 100
        cbpPopup.CommandBar.SetPopupToolBar True
        cbpPopup.CommandBar.Width = 70
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT1, "靠上左对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT2, "靠上居中"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT3, "靠上右对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT4, "中部左对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT5, "中部居中"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT6, "中部右对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT7, "靠下左对齐"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT8, "靠下居中"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT9, "靠下右对齐"

        Set objControl = .Add(xtpControlButton, ID_TABLE_CURRENCY, "货币样式"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_TABLE_PERCENT, "百分比样式")
        Set objControl = .Add(xtpControlButton, ID_TABLE_KILOBIT, "千位分隔样式")

        Set objControl = .Add(xtpControlButton, ID_TABLE_MERGE, "合并单元格"): objControl.BeginGroup = True
        objControl.BeginGroup = True
        If Not Document Is Nothing Then
            If Document.EditType = cprET_病历文件定义 Then
                Set objControl = .Add(xtpControlButton, ID_TABLE_CELLPROTECTED, "锁定单元格")
            End If
        End If
    End With
    
    '工具栏位置调整
    DockingRightOf CommBar(ID_BAR_FORMAT), CommBar(ID_BAR_SIGN)

    '以下是不常用的隐藏菜单
    cbrThis.Options.AddHiddenCommand ID_FILE_IMPORT
    cbrThis.Options.AddHiddenCommand ID_FILE_SAVEAS
    cbrThis.Options.AddHiddenCommand ID_FILE_SAVE_QUIT
    cbrThis.Options.AddHiddenCommand ID_FILE_PRINTPREVIEW
    cbrThis.Options.AddHiddenCommand ID_EDIT_FINDNEXT
    cbrThis.Options.AddHiddenCommand ID_EDIT_REPLACE
'    cbrThis.Options.AddHiddenCommand ID_VIEW_GRID
    cbrThis.Options.AddHiddenCommand ID_INSERT_DATETIME
    cbrThis.Options.AddHiddenCommand ID_FORMAT_PROTECT
    cbrThis.Options.AddHiddenCommand ID_TABLE_SHOWGRID

    '热键绑定
    cbrThis.KeyBindings.Add FCONTROL, Asc("S"), ID_FILE_SAVE
    cbrThis.KeyBindings.Add FCONTROL, Asc("P"), ID_FILE_PRINT
    cbrThis.KeyBindings.Add FCONTROL, Asc("Z"), ID_EDIT_UNDO
    cbrThis.KeyBindings.Add FCONTROL, Asc("Y"), ID_EDIT_REDO
    cbrThis.KeyBindings.Add FCONTROL, Asc("X"), ID_EDIT_CUT
    cbrThis.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    cbrThis.KeyBindings.Add FCONTROL, Asc("V"), ID_EDIT_PASTE
    cbrThis.KeyBindings.Add FCONTROL, Asc("A"), ID_EDIT_SELECTALL
    cbrThis.KeyBindings.Add FCONTROL, Asc("F"), ID_EDIT_FIND
    cbrThis.KeyBindings.Add FCONTROL, Asc("H"), ID_EDIT_REPLACE
    cbrThis.KeyBindings.Add FCONTROL, Asc("D"), ID_VIEW_STRUCTURE
    cbrThis.KeyBindings.Add FCONTROL, Asc("M"), ID_VIEW_PHRASEDEMO
    cbrThis.KeyBindings.Add FCONTROL, Asc("G"), ID_VIEW_SEGMENT
    cbrThis.KeyBindings.Add FCONTROL, Asc("E"), ID_INSERT_ELEMENT
    cbrThis.KeyBindings.Add FCONTROL, Asc("B"), ID_FORMAT_BOLD
'    cbrThis.KeyBindings.Add FCONTROL, Asc("I"), ID_FORMAT_ITALIC
    cbrThis.KeyBindings.Add FCONTROL, Asc("Q"), ID_FILE_EXIT
    cbrThis.KeyBindings.Add FCONTROL, Asc("J"), ID_INSERT_AUTORECOGNISE     '智能识别
    cbrThis.KeyBindings.Add FCONTROL, Asc("T"), ID_VIEW_PENWINDOW           '手写输入窗口
    cbrThis.KeyBindings.Add FCONTROL, Asc("O"), ID_SIGN_QUIT                '签名并退出编辑器

    cbrThis.KeyBindings.Add 0, VK_F1, ID_HELP_CONTENT
    cbrThis.KeyBindings.Add 0, VK_F3, ID_EDIT_FINDNEXT
    cbrThis.KeyBindings.Add 0, VK_F5, ID_INSERT_PICTURE
    cbrThis.KeyBindings.Add 0, VK_F6, ID_TABLE_INSERTTABLE
    cbrThis.KeyBindings.Add 0, VK_F7, ID_INSERT_ELEMENT
    cbrThis.KeyBindings.Add 0, VK_F8, ID_EDIT_ADDCOMPEND
    cbrThis.KeyBindings.Add 0, VK_F9, ID_INSERT_EPRDEMO
    cbrThis.KeyBindings.Add 0, VK_F10, ID_DIAGNOSIS                         '诊断
    cbrThis.KeyBindings.Add 0, VK_F11, ID_SIGN                              '签名
    cbrThis.KeyBindings.Add 0, VK_F12, ID_INSERT_AUTORECOGNISE              '智能识别
    cbrThis.KeyBindings.Add 0, VK_DELETE, ID_EDIT_DELETE
'    cbrThis.KeyBindings.Add 0, VK_BACK, ID_EDIT_BACKSPACE

    '屏蔽编辑器的默认快捷键（不处理）
    cbrThis.KeyBindings.Add FCONTROL, Asc("R"), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("L"), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("L"), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("="), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("1"), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("2"), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("5"), -1
    cbrThis.KeyBindings.Add FCONTROL + FSHIFT, Asc("A"), -1
    cbrThis.KeyBindings.Add FCONTROL + FSHIFT, Asc("L"), -1

    '## 可停靠窗体设置

    Set mfrmSentenceDetailed = New frmSentenceDetailed
    Set mfrmSegments = New frmSegmentList
    Set mfrmCompends = New frmCompends: mfrmCompends.SetParent Me
    Set mfrmModElement = New frmElementEdit
    Set mfrmInsElement = New frmInsElement
    Set mfrmDicSelect = New frmDicSelect
    Set mfrmStyleMan = New frmStyleMan
    Set cPicEditor = New cPictureEditor
    Set mfrmPacsPic = New frmPACSImg
    Set mfrmHistoryReport = New frmDockReportHistory
    Set mfrmDocksymbol = New frmDockSymbol
    DkpThis.SetCommandBars Me.cbrThis
    DkpThis.Options.ThemedFloatingFrames = True
    DkpThis.TabPaintManager.Position = xtpTabPositionTop

    Dim PaneCompend As XtremeDockingPane.Pane           '文档结构图
    Dim PaneSentence As XtremeDockingPane.Pane          '词句示范
    Dim PaneSegment As XtremeDockingPane.Pane           '示范段落
    Dim PaneStyleMan As XtremeDockingPane.Pane          '段落样式维护
    Dim PaneMultiDocView As XtremeDockingPane.Pane      '多文档预览
    Dim PanePacsPic As XtremeDockingPane.Pane           'Pacs图片组
    Dim PaneSharePage As XtremeDockingPane.Pane           '
    Dim PaneHistoryReport As XtremeDockingPane.Pane     '本病人历史报告
    Dim PaneDockSymbol As XtremeDockingPane.Pane        '特殊符号
    
    '共享页面病历查看
    Set PaneSharePage = DkpThis.CreatePane(ID_VIEW_HISTORYWINDOW, 200, 140, DockTopOf, Nothing)
    PaneSharePage.Title = "共享页面"
    PaneSharePage.Options = PaneNoFloatable Or PaneNoHideable
    PaneSharePage.Close

    '多文档预览
    Set PaneMultiDocView = DkpThis.CreatePane(ID_VIEW_MULTIDOCVIEW, 200, 140, DockLeftOf, Nothing)
    PaneMultiDocView.Title = "共用列表"
    PaneMultiDocView.Options = PaneNoCloseable
    
    '文档结构图
    Set PaneCompend = DkpThis.CreatePane(ID_VIEW_STRUCTURE, 200, 140, DockLeftOf, Nothing)
    PaneCompend.Title = "文档结构"
    PaneCompend.Options = PaneNoCloseable
    If Screen.Width / Screen.TwipsPerPixelX <= 800 Then PaneCompend.Hide
    DkpThis.AttachPane PaneCompend, PaneMultiDocView
    PaneMultiDocView.Close

    '示范词句列表
    Set PaneSentence = DkpThis.CreatePane(ID_VIEW_PHRASEDEMO, 200, 140, DockBottomOf, PaneCompend)
    PaneSentence.Title = "词句示范"
    PaneSentence.Options = PaneNoCloseable
    If Screen.Width / Screen.TwipsPerPixelX <= 800 Then PaneSentence.Hide
    DkpThis.AttachPane PaneSentence, PaneCompend

    '示范片段列表
    Set PaneSegment = DkpThis.CreatePane(ID_VIEW_SEGMENT, 200, 140, DockBottomOf, PaneSentence)
    PaneSegment.Title = "示范片段"
    PaneSegment.Options = PaneNoCloseable
    If Screen.Width / Screen.TwipsPerPixelX <= 800 Then PaneSegment.Hide
    DkpThis.AttachPane PaneSegment, PaneSentence


    '报告图片列表
    Set PanePacsPic = DkpThis.CreatePane(ID_VIEW_PACSPIC, 200, 140, DockBottomOf, PaneSentence)
    PanePacsPic.Title = "报告图片"
    PanePacsPic.Options = PaneNoCloseable
    If Screen.Width / Screen.TwipsPerPixelX <= 800 Then PanePacsPic.Hide
    DkpThis.AttachPane PanePacsPic, PaneSentence
    
    '本病人历史报告
    Set PaneHistoryReport = DkpThis.CreatePane(ID_VIEW_HISTORYREPORT, 200, 140, DockTopOf, Nothing)
    PaneHistoryReport.Title = "历史检查"
    PaneHistoryReport.Options = PaneNoCloseable
    If Screen.Width / Screen.TwipsPerPixelX <= 800 Then PaneHistoryReport.Hide
    DkpThis.AttachPane PaneHistoryReport, PaneSentence

    '段落样式维护
    Set PaneStyleMan = DkpThis.CreatePane(ID_FORMAT_STYLEWINDOW, 230, 140, DockRightOf, Nothing)
    PaneStyleMan.Title = "段落样式"
    PaneHistoryReport.Options = PaneNoCloseable
    
    Set PaneDockSymbol = DkpThis.CreatePane(ID_VIEW_Assistant, 230, 140, DockRightOf, Nothing)
    PaneDockSymbol.Title = "输入助手"
    PaneDockSymbol.Options = PaneNoCloseable
    PaneDockSymbol.MaxTrackSize.Width = 280: PaneDockSymbol.MinTrackSize.Width = 230
    If Screen.Width / Screen.TwipsPerPixelX <= 1024 Then PaneDockSymbol.Hide
    DkpThis.AttachPane PaneStyleMan, PaneDockSymbol
    PaneStyleMan.Close
    
    '## 其他初始化设置
    Editor1.Modified = False

    SetParent picPenInput.hwnd, Editor1.hwnd

    ColorForeColor.COLOR = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "ForeColor", vbBlack)
    ColorHighlight.COLOR = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "HighlightColor", vbYellow)
    ColorFillColor.COLOR = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "CellFillColor", vbWhite)
    Editor1.ShowRuler = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "ShowRuler", False)
    SetColorIcon "FILLCOLOR", ID_DRAW_FILLCOLOR, ColorFillColor.COLOR
    SetColorIcon "FORECOLOR", ID_FORMAT_FORECOLOR, ColorForeColor.COLOR
    SetColorIcon "HIGHLIGHT", ID_FORMAT_HIGHLIGHT, IIf(ColorHighlight.COLOR = tomAutoColor, vbWhite, ColorHighlight.COLOR)

    Call RestoreWinState(Me, App.ProductName)   '在创建所有对象后再恢复窗体布局

    If zlDatabase.GetPara("使用个性化风格", glngSys, , 0) = 0 Then
        Me.WindowState = vbMaximized
    End If

    CommBar(ID_BAR_TABLE).Visible = False
    CommBar(ID_BAR_NORMAL).FindControl(, ID_COMMON_CANCEL).STYLE = xtpButtonIconAndCaption     '恢复图标＋文本类型
    CommBar(ID_BAR_SIGN).FindControl(, ID_REVISION_RESET).STYLE = xtpButtonIconAndCaption
    CommBar(ID_BAR_SIGN).FindControl(, ID_SIGN).STYLE = xtpButtonIconAndCaption
    CommBar(ID_BAR_SIGN).FindControl(, ID_UNTREAD).STYLE = xtpButtonIconAndCaption
    CommBar(ID_BAR_SIGN).FindControl(, ID_SIGN_QUIT).STYLE = xtpButtonIconAndCaption
    
    If imgX_S.Top < 0 Then
        imgX_S.Top = 5460
    End If
    
    '调整几个窗格的显示顺序，保证：如果有词句示范，则靠前显示
    DkpThis.ShowPane ID_VIEW_PHRASEDEMO
    If GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneCompendHided", False) Then PaneCompend.Close
    If GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneSentenceHided", False) Then PaneSentence.Close
    If GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneSegmentHided", False) Then PaneSegment.Close
    If GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PanePacsPicHided", False) Then PanePacsPic.Close
    If GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneHistoryReportHided", False) Then PaneHistoryReport.Close
    If GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneDockSymbol", False) Then PaneDockSymbol.Hide

    '## 表格选择器
    Set cDropDown = New cDropDownToolWindow
    cDropDown.Create picDropDown

    tmrThis.Enabled = True
    DT1 = Now
    DT1_EPR = Now
End Sub

'################################################################################################################
'## 功能：  确认是否保存修改后的文件
'################################################################################################################
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If mblnPrecess Then
        Cancel = 1: Exit Sub
    End If
    If Editor1.Modified Then
        Dim r As Long
        r = MsgBox("是否保存对 """ & Me.Document.EPRFileInfo.名称 & """ 的更改？", vbYesNoCancel + vbExclamation, gstrSysName)
        If r = vbYes Then
            '保存文件
            If Me.Document.目标版本 > 16 Then
                MsgBox "目前系统支持的最大版本号为16，保存失败！", vbOKOnly + vbInformation, gstrSysName
                Cancel = 1
                Exit Sub
            End If
            If SaveEMRDoc = False Then Cancel = 1
        ElseIf r = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

'################################################################################################################
'## 功能：  检查是否修改，并提示是否保存
'################################################################################################################
Public Function CheckModified(Optional blnCannotCancel As Boolean = False) As Boolean
    CheckModified = True
    If Editor1.Modified And Me.Document.EPRFileInfo.ID <> 0 Then
        Dim r As Long
        If blnCannotCancel Then
            r = MsgBox("是否保存对 """ & Me.Document.EPRFileInfo.名称 & """ 的更改？", vbYesNo + vbExclamation, gstrSysName)
        Else
            r = MsgBox("是否保存对 """ & Me.Document.EPRFileInfo.名称 & """ 的更改？", vbYesNoCancel + vbExclamation, gstrSysName)
        End If
        If r = vbYes Then
            '保存文件
            If Me.Document.目标版本 > 16 Then
                MsgBox "目前系统支持的最大版本号为16，保存失败！", vbOKOnly + vbInformation, gstrSysName
                CheckModified = False
                Exit Function
            End If
            If SaveEMRDoc = False Then CheckModified = False
        ElseIf r = vbCancel Then
            CheckModified = False
        End If
    End If
End Function

'################################################################################################################
'## 功能：  保存相关信息，关闭窗体
'################################################################################################################
Private Sub Form_Unload(Cancel As Integer)
    '保存位置信息
    If mblnPrecess Then
        Cancel = 1
    End If
    Dim i As Long
    If Me.WindowState <> vbMinimized Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "ForeColor", ColorForeColor.COLOR
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "HighlightColor", ColorHighlight.COLOR
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "CellFillColor", ColorFillColor.COLOR
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "ShowRuler", Me.Editor1.ShowRuler
    End If
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneCompendHided", DkpThis.FindPane(ID_VIEW_STRUCTURE).Closed
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneSentenceHided", DkpThis.FindPane(ID_VIEW_PHRASEDEMO).Closed
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneSegmentHided", DkpThis.FindPane(ID_VIEW_SEGMENT).Closed
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PanePacsPicHided", DkpThis.FindPane(ID_VIEW_PACSPIC).Closed
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneHistoryReportHided", DkpThis.FindPane(ID_VIEW_HISTORYREPORT).Closed
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneDockSymbol", DkpThis.FindPane(ID_VIEW_Assistant).Hidden
    

    Call SaveWinState(Me, App.ProductName)
    
    On Error Resume Next
    If Not Me.Document Is Nothing Then
        Call Me.Document.AfterClosed(Me.Document.EPRPatiRecInfo.医嘱id) '触发关闭事件
    End If
       
        Dim objTmp As Object
    For Each objTmp In Me.Controls
        Me.Controls.Remove objTmp
    Next

    '清除临时文件
    For i = 1 To UBound(UndoList)
        If gobjFSO.FileExists(UndoList(i).Filename) Then gobjFSO.DeleteFile UndoList(i).Filename
    Next

    Erase mEleLimit
    Erase UndoList
    Set mParaFmt = Nothing
    Set mFontFmt = Nothing
    
    If Not mfrmCompends Is Nothing Then Unload mfrmCompends
    If Not mfrmSentenceDetailed Is Nothing Then Unload mfrmSentenceDetailed
    If Not mfrmSegments Is Nothing Then Unload mfrmSegments
    If Not mfrmModElement Is Nothing Then Unload mfrmModElement
    If Not mfrmInsElement Is Nothing Then Unload mfrmInsElement
    If Not mfrmDicSelect Is Nothing Then Unload mfrmDicSelect
    If Not mfrmStyleMan Is Nothing Then Unload mfrmStyleMan
    If Not mfrmMultiDocView Is Nothing Then Unload mfrmMultiDocView
    If Not mfrmPacsPic Is Nothing Then Unload mfrmPacsPic
    If Not mfrmHistoryReport Is Nothing Then Unload mfrmHistoryReport
    If Not mfrmMainError Is Nothing Then Unload mfrmMainError
    If Not mfrmPreview Is Nothing Then Unload mfrmPreview
    If Not mfrmDocksymbol Is Nothing Then Unload mfrmDocksymbol
'    If Not gfrmPublic Is Nothing Then Unload gfrmPublic
     
    Set mfrmCompends = Nothing
    Set mfrmSentenceDetailed = Nothing
    Set mfrmSegments = Nothing
    Set mfrmModElement = Nothing
    Set mfrmInsElement = Nothing
    Set mfrmDicSelect = Nothing
    Set mfrmStyleMan = Nothing
    Set cPicEditor = Nothing
    Set mfrmMultiDocView = Nothing
    Set mfrmPacsPic = Nothing
    Set mfrmHistoryReport = Nothing
    Set mfrmMainError = Nothing
    Set mfrmPreview = Nothing
    Set mfrmDocksymbol = Nothing
    Set mobjReport = Nothing
    Set cDropDown = Nothing

    Set Me.Document = Nothing

'    ImageManager.Icons.RemoveAll
    imgColor.ListImages.Clear
    ImageList_Destroy imgColor.hImageList
    Set imgX_S.Picture = Nothing
    Set picDropDown.Picture = Nothing
    Set picHistoryInfo.Picture = Nothing
    Set picPane.Picture = Nothing
    Set picPatiInfo.Picture = Nothing
    Set picPenInput.Picture = Nothing
    Set picTMP.Picture = Nothing
'    '手动释放内存
'    SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
End Sub

Public Sub ShowTablePicker(ByVal X As Long, ByVal y As Long)
   '显示表格下拉选择器
   DefaultSize picDropDown
   cDropDown.Show X, y
   picDropDown.Visible = True
   DrawTableChooser cDropDown, 0, 0, 0
End Sub

Private Sub HideDropDown()
   cDropDown.Hide
End Sub

Private Sub picDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    ' Mouse down handling:
    If (cDropDown.IsShown) Then
        ' Drop down window is visible
        If Not (cDropDown.InRect(X, y)) Then
            ' Mouse down outside drop-down area:
            HideDropDown
        Else
            ' Mouse down inside the drop down:
            DrawTableChooser cDropDown, X \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY, Button
        End If
    End If
End Sub

Private Sub picDropDown_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    ' Mouse move.  Note that because all mouse messages are captured,
    ' x and y may be outside the limits of picDropdown.  This is
    ' handled in the draw routine.
    DrawTableChooser cDropDown, X \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY, Button
End Sub

Private Sub picDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim xCellHit As Long, yCellHit As Long, bIn As Boolean

    ' Mouse up.  Determine whether mouse up over a cell:
    DrawTableChooser cDropDown, X \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY, Button, bIn, xCellHit, yCellHit
    ' Hide the drop down:
    HideDropDown
    ' If an item selected, say what it was:
    If (bIn) Then
        Dim strTmp As String, lLen As Long, lngKey As Long, i As Long, j As Long, lKey2 As Long
        Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean

        With Editor1
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.SelStart + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys = False Then
                lKey = Me.Document.Tables.Add
                tblThis.Redraw = False
                tblThis.SingleClickEdit = False
                tblThis.HighlightMode = HMFilledRectAlpha
                tblThis.Width = Me.Editor1.PaperWidth - Me.Editor1.Selection.Para.LeftIndent - Me.Editor1.MarginLeft - Me.Editor1.MarginRight - 800
                tblThis.Init yCellHit, xCellHit
                tblThis.CellMargin = 10
                For i = 1 To yCellHit
                    For j = 1 To xCellHit
                        Me.Document.Tables("K" & lKey).Cells.Add , i, j
'                        tblThis.Cell(i, j).Width = 100
                    Next j
                Next i
                tblThis.Tag = lKey
                tblThis.ShowToolTipText = True
                tblThis.MinRowHeight = 300
                tblThis.Redraw = True
                tblThis.Refresh
                SaveUIToTable Me.Document.Tables("K" & lKey), True
            End If
        End With
    End If
End Sub

Private Sub mfrmCompends_NodeSelected(lngCompendID As Long)
    '同步示范词句的显示
    Call RefSentenceList
End Sub

Private Sub mfrmDicSelect_pOK(strR As String)
    '字典项目返回值
    Dim strTmp As String, T As Variant
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    If tblThis.Visible Then
        '表格中编辑要素
        If Val(tblThis.Tag) > 0 And Val(mfrmDicSelect.Tag) > 0 Then
            T = Split(strR, ";")
            If UBound(T) > 0 And Me.Document.Tables("K" & tblThis.Tag).Elements("K" & mfrmDicSelect.Tag).替换域 = 2 Then
                Me.Document.Tables("K" & tblThis.Tag).Elements("K" & mfrmDicSelect.Tag).内容文本 = T(1)
                '保存到单元格中
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = T(1)
                tblThis.Refresh False, True, tblThis.SelectedCellKey
                tblThis_Resize tblThis.Width, tblThis.Height
                tblThis.Modified = True
                If mintStyle <> 0 Then '模态状态不能使用
                    tblThis.SetFocus
                End If
            End If
        End If
        Exit Sub
    End If

    bFinded = FindKey(Editor1, "E", glngCurEleKey, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        T = Split(strR, ";")
        If UBound(T) > 0 And Me.Document.Elements("K" & glngCurEleKey).替换域 = 2 Then
            strTmp = T(1)
            Me.Document.Elements("K" & glngCurEleKey).内容文本 = strTmp

            Me.Document.Elements("K" & glngCurEleKey).Refresh Me.Editor1

            Call UpdateSameELement(glngCurEleKey)
            Call FindKey(Editor1, "E", glngCurEleKey, lKSS, lKSE, lKES, lKEE, bNeeded) '因为UpdateSameELement有可能改变文本长度，从而改变选中位置
            
            If Trim(Me.Document.Elements("K" & glngCurEleKey).内容文本) <> "" Then
                '自动定位到下一个要素位置
                bFinded = FindNextKey(Editor1, Editor1.Selection.StartPos + 1, "E", glngCurEleKey, lKSS, lKSE, lKES, lKEE, bNeeded)
                If bFinded Then
                    Editor1.Range(lKSE, lKES).Selected
                End If
            Else
                Editor1.Range(lKSE, lKSE + Me.Document.Elements("K" & glngCurEleKey).GetValidTextLength).Selected
            End If
            Editor1.ForceEdit = False
            Editor1.UnFreeze
        End If
    End If
End Sub

'################################################################################################################
'## 功能：  取消诊治要素的插入
'################################################################################################################
Private Sub mfrmInsElement_pCancel()
    mfrmInsElement.Hide
    mfrmInsElement.Tag = ""
End Sub

'################################################################################################################
'## 功能：  接受诊治要素的插入
'################################################################################################################
Private Sub mfrmInsElement_pOK(Ele As cEPRElement)
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lngKey As Long, bInKeys As Boolean, bNeeded As Boolean, lKey2 As Long
    If mfrmInsElement.Tag <> "" Then
        '修改模式
        If mbEditInTable Then
            '表格中的要素
            If Val(tblThis.Tag) > 0 Then
                Me.Document.Tables("K" & Val(tblThis.Tag)).Elements.Remove "K" & mfrmInsElement.Tag
                lngKey = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements.AddExistNode(Ele, True)
                If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                    Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).内容文本 = GetReplaceEleValue(Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).要素名称, _
                        Me.Document.EPRPatiRecInfo.病人ID, _
                        Me.Document.EPRPatiRecInfo.主页ID, _
                        Me.Document.EPRPatiRecInfo.病人来源, _
                        Me.Document.EPRPatiRecInfo.医嘱id, _
                        Me.Document.EPRPatiRecInfo.婴儿)
                End If
                If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                    If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).自动转文本 Then Me.Document.EleToString Me.Editor1, Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey)  '自动转化为纯文本
                End If
                '保存到单元格中
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).内容文本
                tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lngKey
                tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).要素名称
                tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
                tblThis.Modified = True
                tblThis.Refresh False, True, tblThis.SelectedCellKey
                tblThis_Resize tblThis.Width, tblThis.Height
            End If
        Else
            '文本中的要素
            Me.Document.Elements.Remove "K" & mfrmInsElement.Tag
            lngKey = Me.Document.Elements.AddExistNode(Ele, True)
            If Me.Document.Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                Me.Document.Elements("K" & lngKey).内容文本 = GetReplaceEleValue(Me.Document.Elements("K" & lngKey).要素名称, _
                    Me.Document.EPRPatiRecInfo.病人ID, _
                    Me.Document.EPRPatiRecInfo.主页ID, _
                    Me.Document.EPRPatiRecInfo.病人来源, _
                    Me.Document.EPRPatiRecInfo.医嘱id, _
                    Me.Document.EPRPatiRecInfo.婴儿)
            End If
            Me.Document.Elements("K" & lngKey).Refresh Me.Editor1
            If Me.Document.Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                If Me.Document.Elements("K" & lngKey).自动转文本 Then Me.Document.EleToString Me.Editor1, Me.Document.Elements("K" & lngKey)  '自动转化为纯文本
            End If
            bInKeys = FindKey(Me.Editor1, "E", lngKey, lSS, lSE, lES, lEE, bNeeded)
            If bInKeys Then
                If Me.Document.Elements("K" & lngKey).输入形态 = 0 Then
                    Editor1.Range(lSE, lES).Selected
                Else
                    Editor1.Range(lSE + 1, lSE + 1).Selected
                End If
            End If
        End If
    Else
        If mbEditInTable Then
            '表格中的要素
            If Val(tblThis.Tag) > 0 Then
                lngKey = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements.AddExistNode(Ele)
                If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                    Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).内容文本 = GetReplaceEleValue(Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).要素名称, _
                        Me.Document.EPRPatiRecInfo.病人ID, _
                        Me.Document.EPRPatiRecInfo.主页ID, _
                        Me.Document.EPRPatiRecInfo.病人来源, _
                        Me.Document.EPRPatiRecInfo.医嘱id, _
                        Me.Document.EPRPatiRecInfo.婴儿)
                End If
                Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).开始版 = Me.Document.目标版本
                If Val(tblThis.Tag) <= 0 Then Exit Sub
                If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                    If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).自动转文本 Then Me.Document.EleToString Me.Editor1, Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey)  '自动转化为纯文本
                End If
                '保存到单元格中
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).内容文本
                tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lngKey
                tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).要素名称
                tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
                tblThis.Modified = True
                tblThis.Refresh False, True, tblThis.SelectedCellKey
                tblThis_Resize tblThis.Width, tblThis.Height
            End If
        Else
            lngKey = Me.Document.Elements.AddExistNode(Ele)
            If Me.Document.Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                Me.Document.Elements("K" & lngKey).内容文本 = GetReplaceEleValue(Me.Document.Elements("K" & lngKey).要素名称, _
                    Me.Document.EPRPatiRecInfo.病人ID, _
                    Me.Document.EPRPatiRecInfo.主页ID, _
                    Me.Document.EPRPatiRecInfo.病人来源, _
                    Me.Document.EPRPatiRecInfo.医嘱id, _
                    Me.Document.EPRPatiRecInfo.婴儿)
            End If
            Me.Document.Elements("K" & lngKey).开始版 = Me.Document.目标版本
            Me.Document.Elements("K" & lngKey).InsertIntoEditor Me.Editor1, , True
            If Me.Document.Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                If Me.Document.Elements("K" & lngKey).自动转文本 Then Me.Document.EleToString Me.Editor1, Me.Document.Elements("K" & lngKey)  '自动转化为纯文本
            End If
        End If
    End If
    mfrmInsElement.Tag = ""
End Sub

'################################################################################################################
'## 功能：  取消诊治要素的编辑
'################################################################################################################
Private Sub mfrmModElement_pCancel()
    On Error Resume Next
    Unload mfrmModElement
'    mfrmModElement.Hide
    Err.Clear
End Sub

'################################################################################################################
'## 功能：  接受诊治要素的编辑
'################################################################################################################
Private Sub mfrmModElement_pOK()
    '保存诊治要素编辑结果
    Dim strTmp As String, lngKey As Long, Ele As cEPRElement, lS As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    If tblThis.Visible Then
        '表格中编辑要素
        If Val(tblThis.Tag) > 0 Then
            Me.Document.Tables("K" & tblThis.Tag).Elements.Remove "K" & tblThis.Cells("K" & tblThis.SelectedCellKey).Tag
            Set Ele = mfrmModElement.Element.Clone(True)
            lngKey = Me.Document.Tables("K" & tblThis.Tag).Elements.AddExistNode(Ele, True)
            If Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).替换域 = 1 And mfrmModElement.Element.内容文本 = "" Then
                Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).内容文本 = GetReplaceEleValue(Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).要素名称, Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.主页ID, Me.Document.EPRPatiRecInfo.病人来源, Me.Document.EPRPatiRecInfo.医嘱id, Me.Document.EPRPatiRecInfo.婴儿)
            End If
            '保存到单元格中
            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).内容文本
            tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lngKey
            tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).要素名称
            tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
            tblThis.Refresh False, True, tblThis.SelectedCellKey
            tblThis_Resize tblThis.Width, tblThis.Height
            tblThis.Modified = True
            tblThis.SetFocus
        End If
        Exit Sub
    End If

    bFinded = FindKey(Editor1, "E", glngCurEleKey, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        lS = lKEE
        If mfrmModElement.Element.替换域 = 1 And (Me.Document.EditType = cprET_病历文件定义) Then
            '自动定位到下一个要素位置
            bFinded = FindNextKey(Editor1, Editor1.Selection.StartPos + 1, "E", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
            If bFinded Then
                '距离在ELE_JUMP_LIMIT个字符之间的才定位过去
                If lKSS - lS < ELE_JUMP_LIMIT Then Editor1.Range(lKSE, lKES).Selected
            End If
        Else
            Me.Document.Elements.Remove "K" & glngCurEleKey
            Set Ele = mfrmModElement.Element.Clone(True)
            lngKey = Me.Document.Elements.AddExistNode(Ele, True)
            If Me.Document.Elements("K" & lngKey).替换域 = 1 And mfrmModElement.Element.内容文本 = "" Then
                Me.Document.Elements("K" & lngKey).内容文本 = GetReplaceEleValue(Me.Document.Elements("K" & lngKey).要素名称, _
                    Me.Document.EPRPatiRecInfo.病人ID, _
                    Me.Document.EPRPatiRecInfo.主页ID, _
                    Me.Document.EPRPatiRecInfo.病人来源, _
                    Me.Document.EPRPatiRecInfo.医嘱id, _
                    Me.Document.EPRPatiRecInfo.婴儿)
            End If

            Me.Document.Elements("K" & lngKey).Refresh Me.Editor1
            bFinded = FindKey(Editor1, "E", lngKey, lKSS, lKSE, lKES, lKEE, bNeeded)
            If bFinded Then lS = lKEE   '修正lS

            If Me.Document.Elements("K" & lngKey).输入形态 = 0 Then
                '自定义项的单独处理
                If InStr(Trim(Me.Document.Elements("K" & lngKey).内容文本), "自定义") > 0 And Me.Document.Elements("K" & lngKey).动态域 = 1 Then
                    strTmp = Trim(InputBox("请录入自定义要素选项" & vbCrLf & "最大输入长度200个汉字", "中联软件"))
                    If strTmp <> "" Then
                        Me.Document.Elements("K" & lngKey).内容文本 = Replace(Me.Document.Elements("K" & lngKey).内容文本, "自定义", strTmp)
                    Else
                        Me.Document.Elements("K" & lngKey).内容文本 = Replace(Me.Document.Elements("K" & lngKey).内容文本, "自定义", "")
                    End If
                    Me.Document.Elements("K" & lngKey).Refresh Me.Editor1
                End If
                
                Call CheckElementLimit(lngKey)
                bFinded = FindKey(Editor1, "E", lngKey, lKSS, lKSE, lKES, lKEE, bNeeded)
                If bFinded Then lS = lKEE   '修正lS,因为CheckElementLimit有可能改变文本长度，从而改变选中位置
                
                If Trim(Me.Document.Elements("K" & lngKey).内容文本) <> "" Then
                    '自动定位到下一个要素位置
                    bFinded = FindNextKey(Editor1, lKEE, "E", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
                    If bFinded Then
                        '距离在ELE_JUMP_LIMIT个字符之间的才定位过去
                        If lKSS - lS < ELE_JUMP_LIMIT Then Editor1.Range(lKSE, lKES).Selected
                    End If
                Else
                    Editor1.Range(lKSE, lKSE + Me.Document.Elements("K" & lngKey).GetValidTextLength).Selected
                End If
            End If
        End If
    End If
End Sub

'################################################################################################################
'## 功能：  将工具条A放置到工具条B的同一行
'##
'## 参数：  BarToDock   ：加入的工具栏
'##         BarOnLeft   ：位于左边的工具条
'################################################################################################################
Private Sub DockingRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    cbrThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    cbrThis.DockToolBar BarToDock, 0, (Bottom + Top) / 2, BarOnLeft.Position
End Sub

'################################################################################################################
'## 功能：  清空文档
'################################################################################################################
Private Sub ClearDoc()
    If Len(Trim(Editor1.Text)) > 0 Then
        Dim r As Long
        If Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_全文示范编辑 Then
            r = MsgBox("警告：清空当前文档后文档将恢复至文件定义初始状态，所有已经录入的信息将丢失！" & vbCrLf & "确认要清空所有内容吗？", vbYesNo + vbExclamation, gstrSysName)
        ElseIf Me.Document.EditType = cprET_病历文件定义 Then
            r = MsgBox("警告：清空当前文档将丢失所有已经录入的信息！" & vbCrLf & "确认要清空所有内容吗？", vbYesNo + vbExclamation, gstrSysName)
        Else
            Exit Sub
        End If
        If r = vbYes Then
            Call AddUndoPoint  '手动缓存
            Editor1.InProcessing = True
            Editor1.Freeze
            Editor1.Tag = "ClearDoc"
            Editor1.NewDoc
            Set Me.Document.Compends = New cEPRCompends
            Set Me.Document.Elements = New cEPRElements
            Set Me.Document.Tables = New cEPRTables
            Set Me.Document.Pictures = New cEPRPictures
            If Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_全文示范编辑 Then
                Me.Document.ReadInitFileStructure Me.Editor1
            End If
            Editor1.UnFreeze
            Editor1.Tag = ""
            Me.RefCompends
            Call ClearNoUseUndoList
        Else
            Exit Sub
        End If
    End If
    Call SetStateInfo
    Editor1.Filename = ""
    Editor1.Modified = True
    Editor1.InProcessing = False
    If tblThis.Visible = False Then Editor1.SetFocus
End Sub

'################################################################################################################
'## 功能：  设置当前状态信息（标题栏、状态栏的设置）
'################################################################################################################
Public Sub SetStateInfo()
    Select Case Document.EditType
    Case cprET_病历文件定义
        stbThis.Panels(3).Text = "文件定义"
        Me.Caption = Document.EPRFileInfo.名称
    Case cprET_全文示范编辑
        stbThis.Panels(3).Text = "范文编辑"
        Me.Caption = Document.EPRFileInfo.名称
    Case cprET_单病历编辑
        stbThis.Panels(3).Text = "文件编辑"
        Me.Caption = Document.EPRFileInfo.名称 & " (第" & Document.目标版本 & "版) "
    Case cprET_单病历审核
        stbThis.Panels(3).Text = "文件修订"
        Me.Caption = Document.EPRFileInfo.名称 & " (第" & Document.目标版本 & "版)"
    End Select
    Select Case Document.EditMode
    Case cprEM_新增
        stbThis.Panels(4).Text = "【新增】"
    Case cprEM_修改
        stbThis.Panels(4).Text = "【修改】"
    End Select
    
    Me.Caption = gstrUserName & "：" & Me.Caption
End Sub

'################################################################################################################
'## 功能：  引入历史文件。
'################################################################################################################
Private Sub ImportEPRDoc()
On Error GoTo errHand

    Dim f As New frmEPRSearchMan, lngR As Long
    lngR = f.ShowSearchFile(Me, Me.Document.EPRFileInfo.ID, Me.Document.EPRPatiRecInfo.科室ID)
    If lngR > 0 Then
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "Select Zl_Fun_ImportEnable([1]) CopyEnable From Dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngR)
        If rsTemp!CopyEnable <> 1 Then
            MsgBox "选定的病历文件不允许引入", vbInformation, gstrSysName
            Exit Sub
        End If

        '导入指定的历史文件
        If Me.Document.ImportOldEPRFile(Me.Editor1, lngR) Then
            MsgBox "成功导入历史文件！", vbOKOnly + vbInformation, gstrSysName
        End If
    End If
    Exit Sub
errHand:
    MsgBox Err.Description, vbInformation, gstrSysName
End Sub

'################################################################################################################
'## 功能：  打开一个RTF文档。
'################################################################################################################
Private Sub OpenRTFDoc()
    If Editor1.ViewMode <> cprNormal Then Exit Sub
    If Editor1.Modified Then
        Dim r As Long
        r = MsgBox("是否保存对 """ & Editor1.Title & """ 的更改?            ", vbYesNoCancel + vbExclamation, gstrSysName)
        If r = vbYes Then
            '保存文件
            If Editor1.Filename = "" Then
                dlgThis.Filename = ""
                dlgThis.Filter = "*.rtf|*.rtf|*.*|*.*"
                dlgThis.ShowSave
                If dlgThis.Filename <> "" Then
                    Editor1.SaveDoc dlgThis.Filename
                Else
                    Exit Sub
                End If
            Else
                Editor1.SaveDoc
            End If
        ElseIf r = vbCancel Then
            Exit Sub
        End If
    End If
    dlgThis.Filename = ""
    dlgThis.Filter = "*.rtf|*.rtf|*.txt|*.txt|*.html|*.html|*.htm|*.htm|*.*|*.*"
    dlgThis.ShowOpen
    If dlgThis.Filename <> "" Then
        Editor1.OpenDoc dlgThis.Filename
        Me.Caption = Editor1.Title & " - zlRichEMR"
    End If
End Sub

'################################################################################################################
'## 功能：  保存文档至数据库
'################################################################################################################
Public Function SaveEMRDoc(Optional ByVal blnNoAsk As Boolean = False) As Boolean
    Dim blnR As Boolean, eEditMode As EditModeEnum
    Dim strSQL As String, strTime As String
    eEditMode = Me.Document.EditMode
    If Me.Editor1.Modified Then
        If blnNoAsk Then
            blnR = Me.Document.SaveEPRDoc(Editor1, InStr(1, gstrPrivsEpr, "修改页面设置") > 0)

            Editor1.Modified = Not blnR
        Else
            If tblThis.Visible Then
                Editor1_UIClose 0
                Editor1.CloseUIInterface
            End If
            blnR = Me.Document.SaveEPRDoc(Editor1, InStr(1, gstrPrivsEpr, "修改页面设置") > 0)
            Editor1.Modified = Not blnR
        End If
    Else
        If (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
            If blnNoAsk Then
                blnR = Me.Document.SaveEPRDoc(Editor1, InStr(1, gstrPrivsEpr, "修改页面设置") > 0)

                Me.Editor1.Modified = False
            Else
                If tblThis.Visible Then
                    Editor1_UIClose 0
                    Editor1.CloseUIInterface
                End If
                blnR = Me.Document.SaveEPRDoc(Editor1, InStr(1, gstrPrivsEpr, "修改页面设置") > 0)
                Me.Editor1.Modified = False
            End If
        End If
    End If
    
    If mbln返修处理 And mblnFBContentChanged Then
        strTime = "to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
        strSQL = "Zl_疾病申报记录_Update(" & Me.Document.EPRPatiRecInfo.ID & ",5,null,null,null,'" & gstrUserName & "'," & strTime & ",'" & Trim(txtContent.Text) & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        mblnFBContentChanged = False
    End If
    SaveEMRDoc = blnR
    If blnR And mblnIsMultiMode And Not mfrmMultiDocView Is Nothing Then
        mfrmMultiDocView.InitData Me, Me.Document, Me.Document.EPRPatiRecInfo.ID
    End If
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## 功能：  检查是否所有必须填入的对象都不为空，如果有空的给予提示
'##
'## 返回：  如果用户取消保存继续输入则返回False，否则返回True（表示强制保存）
'################################################################################################################
Private Function CheckAllObjects(Optional CheckItemMust As Boolean) As Boolean
    If Me.Editor1.AuditMode Or Me.Document.EditType = cprET_病历文件定义 And Me.Document.EPRFileInfo.保留 <> 3 Or Me.Document.EditType = cprET_全文示范编辑 Then
        CheckAllObjects = True: Exit Function
    End If
    Dim i As Long, blnOK As Boolean, lngIndex As Long, strMsg As String
    
    If mfrmMainError Is Nothing Then Set mfrmMainError = New frmMainMsg
    
    CheckAllObjects = mfrmMainError.ShowNotice(Me, CheckItemMust)

End Function

'################################################################################################################
'## 功能：  保存文档为范文
'################################################################################################################
Private Sub SaveDocAsEPRDemo()
    If Editor1.Modified Then MsgBox "另存操作之前，请先保存本次编辑！", vbExclamation, gstrSysName: Exit Sub
    Select Case Me.Document.EditType
    Case cprET_全文示范编辑
        Call frmEPRModelSaveAs.ShowMe(1, Me.Document.EPRDemoInfo.ID)
        Err = 0: On Error Resume Next
        Call Me.Document.mfrmParent.RefreshList
    Case cprET_单病历编辑, cprET_单病历审核
        If Me.Document.EPRPatiRecInfo.ID = 0 Then MsgBox "另存操作之前，请先保存本次编辑！", vbExclamation, gstrSysName: Exit Sub
        Call frmEPRModelSaveAs.ShowMe(2, Me.Document.EPRPatiRecInfo.ID)
    End Select
End Sub

'################################################################################################################
'## 功能：  保存文档为片段
'################################################################################################################
Private Sub SaveDocAsSegment()
    Dim strCompends As String, objNode As MSComctlLib.Node
    
    If Editor1.Modified Then MsgBox "另存操作之前，请先保存本次编辑！", vbExclamation, gstrSysName: Exit Sub
    strCompends = ""
    For Each objNode In mfrmCompends.Tree.Nodes
        If objNode.Checked = True And Me.Document.Compends("K" & objNode.Tag).ID > 0 And Me.Document.Compends("K" & objNode.Tag).定义提纲ID > 0 Then
            strCompends = strCompends & "," & Me.Document.Compends("K" & objNode.Tag).ID
        End If
    Next
    If strCompends = "" Then
        MsgBox "需要首先选择提纲，以便确定将这些提纲的内容另存为片段！", vbExclamation, gstrSysName
        Exit Sub
    End If
    strCompends = Mid(strCompends, 2)
    
    Select Case Me.Document.EditType
    Case cprET_全文示范编辑
        Call frmEPRModelSaveAs.ShowMe(1, Me.Document.EPRDemoInfo.ID, strCompends)
        Err = 0: On Error Resume Next
        Call Me.Document.mfrmParent.RefreshList
    Case cprET_单病历编辑, cprET_单病历审核
        If Me.Document.EPRPatiRecInfo.ID = 0 Then MsgBox "另存操作之前，请先保存本次编辑！", vbExclamation, gstrSysName: Exit Sub
        Call frmEPRModelSaveAs.ShowMe(2, Me.Document.EPRPatiRecInfo.ID, strCompends)
        Call mfrmSegments.zlRefresh(Me)
    End Select
End Sub

'################################################################################################################
'## 功能：  获取页眉页脚替换结果文本，只用于导出到Word打印
'################################################################################################################
Private Function GetReplacedHeadFootStr(strIn As String) As String
    Dim strR As String, strUnitName As String
    strUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    strR = strIn
    strR = Replace(strR, "{单位名称}", strUnitName)
    strR = Replace(strR, "{标题}", Editor1.Title)
    strR = Replace(strR, "{路径}", Left(Editor1.Filename, InStrRev(Editor1.Filename, "\")))
    strR = Replace(strR, "{文件名}", Mid(Editor1.Filename, InStrRev(Editor1.Filename, "\") + 1))
    strR = Replace(strR, "{日期}", Format(Now(), "yyyy年mm月dd日"))
    strR = Replace(strR, "{时间}", Format(Now(), "hh:MM:ss"))
    strR = Replace(strR, "{打印日期}", Format(Now(), "yyyy年mm月dd日"))
    strR = Replace(strR, "{打印时间}", Format(Now(), "hh:MM:ss"))
    GetReplacedHeadFootStr = strR
End Function

'################################################################################################################
'## 功能：  导出为RTF，然后通过Word打印当前文件
'################################################################################################################
Private Sub PrintInWord()
    On Error Resume Next
    Dim strF As String, strPicFile As String, Fd As Object
    strF = GetSysTmpPath & "\PrintInWord_TMP" & App.ThreadID & ".rtf"
    If gobjFSO.FileExists(strF) Then gobjFSO.DeleteFile strF, True
    '更改所有左对齐为两端对齐
    Dim i As Long, j As Long
    Do
        i = InStr(i + 1, Editor1.Text, vbCrLf)
        If i > 0 Then
            If Editor1.TOM.TextDocument.Range(i - 2, i - 2).Para.Alignment = tomAlignLeft Then
                Editor1.TOM.TextDocument.Range(i - 2, i - 2).Para.Alignment = tomAlignJustify
            End If
        End If
    Loop Until i <= 0
    
    If SaveDocToFile(strF) Then
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
                WordDoc.PageSetup.LeftMargin = Me.ScaleX(Editor1.MarginLeft, vbTwips, vbPoints)
                WordDoc.PageSetup.RightMargin = Me.ScaleX(Editor1.MarginRight, vbTwips, vbPoints)
                WordDoc.PageSetup.TopMargin = Me.ScaleY(Editor1.MarginTop, vbTwips, vbPoints)
                WordDoc.PageSetup.BottomMargin = Me.ScaleY(Editor1.MarginBottom, vbTwips, vbPoints)
                WordDoc.PageSetup.PageWidth = Me.ScaleX(Editor1.PaperWidth, vbTwips, vbPoints)
                WordDoc.PageSetup.PageHeight = Me.ScaleY(Editor1.PaperHeight, vbTwips, vbPoints)
                
                If WordApp.ActiveWindow.ActivePane.View.Type = 1 Or WordApp.ActiveWindow.ActivePane.View.Type = 2 Then
                    WordApp.ActiveWindow.ActivePane.View.Type = 3
                    'wdNormalView=1     wdOutlineView=2     wdPrintView=3
                End If
                
                WordApp.ActiveWindow.View = 5   'wdMasterView
                '添加当前的页眉页脚到RTF文件中
                WordApp.ActiveWindow.View.SeekView = 9  'wdSeekCurrentPageHeader
                WordApp.Selection.ParagraphFormat.Alignment = 0     'wdAlignParagraphLeft
                If Not (Editor1.Picture Is Nothing) Then
                    If Editor1.Picture.Handle <> 0 Then
                        strPicFile = GetSysTmpPath & "\zlDocHead" & App.ThreadID & ".BMP"
                        If gobjFSO.FileExists(strPicFile) Then gobjFSO.DeleteFile strPicFile, True
                        SavePicture Editor1.Picture, strPicFile
                        If gobjFSO.FileExists(strPicFile) Then
                            WordApp.Selection.InlineShapes.AddPicture Filename:=strPicFile, LinkToFile:=False, SaveWithDocument:=True
                            gobjFSO.DeleteFile strPicFile, True
                            WordApp.Selection.TypeParagraph
                        End If
                    End If
                End If
                
                edtBuff.HeadFileTextRTF = Editor1.HeadFileTextRTF: edtBuff.DocHeadReplaceKey: edtBuff.DocHeadCopyWithFormat
                WordApp.Selection.Paste
                '去掉 其中的总页数,页码
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
                    Clipboard.Clear
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
                    Clipboard.Clear
                End If
                
                WordApp.ActiveWindow.View.SeekView = 10 'wdSeekCurrentPageFooter'页脚
                edtBuff.FootFileTextRTF = Editor1.FootFileTextRTF: edtBuff.DocFootReplaceKey: edtBuff.DocFootCopyWithFormat
                WordApp.Selection.Paste
                '去掉 其中的总页数,页码
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
                    Clipboard.Clear
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
                    Clipboard.Clear
                End If

                Set Fd = Nothing
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
    End If
End Sub

Private Sub InsertHeadFootInWord(WordApp As Object, strR As String)
    Dim blnFinded As Boolean, i As Long, j As Long, k As Long
    i = 1
    j = Len(strR)
    k = i
    Do While (k <= j)
        If Mid(strR, k, 1) = "{" Then
            If Mid(strR, k, 4) = "{页码}" Then
                WordApp.Selection.Fields.Add Range:=WordApp.Selection.Range, Type:=33   'wdFieldPage
                WordApp.Selection.Start = 999999
                k = k + 4
            ElseIf Mid(strR, k, 5) = "{总页数}" Then
                WordApp.Selection.Fields.Add Range:=WordApp.Selection.Range, Type:=26   'wdFieldNumPages
                WordApp.Selection.Start = 999999
                k = k + 5
            Else
                WordApp.Selection.TypeText Mid(strR, k, 1)
                k = k + 1
            End If
        ElseIf Mid(strR, k, 2) = vbCrLf Then
            WordApp.Selection.TypeText Mid(strR, k, 1)
            k = k + 2
        Else
            WordApp.Selection.TypeText Mid(strR, k, 1)
            k = k + 1
        End If
    Loop
End Sub


'################################################################################################################
'## 功能：  保存文档至RTF文件
'##
'## 参数：  strFileName     ：文件名
'##         blnClearMode    ：是否是清洁模式
'################################################################################################################
Public Function SaveDocToFile(Optional ByVal strFileName As String = "", _
    Optional blnClearMode As Boolean = True, _
    Optional blnClearKeywords As Boolean = True) As Boolean

    On Error GoTo LL
    Dim strF As String
    If strFileName = "" Then
        If Editor1.ViewMode = cprPaper Then Exit Function
        Select Case Me.Document.EditType
        Case cprET_病历文件定义
            dlgThis.Filename = "定义_" & Me.Document.EPRFileInfo.名称 & ".rtf"
        Case cprET_全文示范编辑
            dlgThis.Filename = "范文_" & Me.Document.EPRFileInfo.名称 & "_" & Me.Document.EPRDemoInfo.名称 & ".rtf"
        Case cprET_单病历编辑, cprET_单病历审核
            dlgThis.Filename = "记录_" & Me.Document.EPRFileInfo.名称 & "(" & Me.Document.EPRPatiRecInfo.ID & "," & Me.Document.目标版本 & ").rtf"
        End Select
        dlgThis.Filter = "*.rtf|*.rtf|*.*|*.*"
        dlgThis.ShowSave
        strF = dlgThis.Filename
    Else
        strF = strFileName
    End If
    If strF <> "" Then
        '=================================================================================================
        Dim lngLen As Long, blnReadOnly As Boolean
        With Me.edtBuff
            .NewDoc
            blnReadOnly = .ReadOnly
            .ReadOnly = False
            .ForceEdit = True
            lngLen = Len(Me.Editor1.Text)
            'RTF内容赋值
            Me.Editor1.SaveDoc strF
            .OpenDoc strF
'            .TOM.TextDocument.Selection.FormattedText = Me.Editor1.TOM.TextDocument.Range(0, lngLen).FormattedText
'            .TOM.TextDocument.Range(lngLen, lngLen).Para = Me.Editor1.TOM.TextDocument.Range(lngLen, lngLen).Para.Duplicate

            '清除所有关键字
            Dim i As Long
            Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
            i = 0
            If blnClearKeywords Then
                bFinded = FindNextAnyKey(Me.edtBuff, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                Do While bFinded
                    .Range(lKSS, lKSE) = ""
                    .Range(lKSS + lKES - lKSE, lKSS + lKES - lKSE + 16) = ""
                    i = lKSS + lKES - lKSE
                    bFinded = FindNextAnyKey(Me.edtBuff, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                Loop
            End If
            .SelectAll
            If blnClearMode Then
                .AuditMode = True
                .AcceptAuditText    '清洁模式
            End If
            lngLen = Len(.Text)
            For i = 0 To lngLen - 1
                '只将背景色为要素背景色颜色去掉
                If .Range(i, i + 1).Font.BackColor = ELE_BACKCOLOR Then
                    .Range(i, i + 1).Font.BackColor = tomAutoColor
                End If
                If .Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR Then
                    .Range(i, i + 1).Font.ForeColor = tomAutoColor
                End If
            Next
            .ReadOnly = blnReadOnly
            '保存到文件
            .SaveDoc strF
        End With
    End If
    SaveDocToFile = True
    Exit Function
LL:
    SaveDocToFile = False
End Function
'################################################################################################################
'## 功能：  打印清洁文档
'################################################################################################################
Private Function PrintEPRDoc(ByVal blnPreview As Boolean, Optional ByVal blnClearMode As Boolean = False) As Boolean
'参数：blnPreview－预览
'      blnClearMode-最终格式(无修改痕迹)
    Dim intLoop As Integer
    Dim lngLen As Long, strF As String
    Dim rsTemp As ADODB.Recordset
    Dim strBillNo As String
    Dim strExseNo As String, intExseKind As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim strPicPath As String, strPicFile As String
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim aryPara(19) As String, intPCount As Integer
    Dim aryFlagPara(1) As String
    Dim intRows As Integer, intCols As Integer
    Dim dcmImages As New DicomImages, dcmResultImage As DicomImage
    Dim blnNoAsk As Boolean
    
    zlCommFun.ShowFlash "请稍候..."
    Screen.MousePointer = vbHourglass
    Err.Clear
    On Error GoTo errHand
    
    blnNoAsk = (zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1")

    If Me.Document.EPRFileInfo.种类 = cpr诊疗报告 And Me.Document.EPRFileInfo.通用 = 2 Then
        strBillNo = "ZLCISBILL" & Format(Document.EPRFileInfo.编号, "00000") & "-2"
    
        gstrSQL = "Select 记录性质, No From 病人医嘱发送 Where 医嘱id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取NO", CLng(Document.EPRPatiRecInfo.医嘱id))
        If rsTemp.RecordCount = 0 Then zlCommFun.StopFlash: Screen.MousePointer = vbDefault: Exit Function
        strExseNo = "" & rsTemp!NO
        intExseKind = Val("" & rsTemp!记录性质)
        
        If mobjReport Is Nothing Then Set mobjReport = New clsReport
        If Not blnNoAsk Then
            If mobjReport.ReportPrintSet(gcnOracle, glngSys, strBillNo, Me) = False Then zlCommFun.StopFlash: Screen.MousePointer = vbDefault: Exit Function
        End If
            
        '获取图像
        strPicPath = App.Path & "\TmpImage\"
        If objFile.FolderExists(strPicPath) = False Then objFile.CreateFolder strPicPath
            
            '获取报告图象(包括标记图)生成本地文件
            '一个报告表格中可能排列多个报告图
        intPCount = 0
        gstrSQL = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取ID", CLng(Document.EPRPatiRecInfo.ID))
        Do While Not rsTemp.EOF
            Set cTable = New cEPRTable
            If cTable.GetTableFromDB(cprET_单病历审核, CLng(Document.EPRPatiRecInfo.ID), Val("" & rsTemp!表格Id)) Then
                For intLoop = 1 To cTable.Pictures.Count
                    strPicFile = strPicPath & "PACSPic" & intLoop & ".JPG"
                    If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True
                    If cTable.Pictures(intLoop).PictureType = EPRMarkedPicture Then
                        Set oPicture = cTable.Pictures(intLoop).DrawFinalPic
                    Else
                        Set oPicture = cTable.Pictures(intLoop).OrigPic
                    End If
                    SavePicture oPicture, strPicFile
                    If objFile.FileExists(strPicFile) Then
                        '保存标记图和图象的路径
                        If cTable.Pictures(intLoop).PictureType = EPRMarkedPicture Then
                            aryFlagPara(0) = strPicFile
                        Else
                            aryPara(intPCount) = strPicFile
                            dcmImages.AddNew
                            dcmImages(dcmImages.Count).FileImport strPicFile, "BMP"
                            intPCount = intPCount + 1
                            If intPCount > UBound(aryPara) Then Exit Do
                        End If
                    End If
                Next
            End If
            rsTemp.MoveNext
        Loop
        
        '判断是否需要自动组合图象，自定义报表中只定义了一个图象框，则自动组合图象
        '重新查一次数据库
        gstrSQL = "Select b.名称,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = 1 And b.名称 not like '标记%'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参数", strBillNo)
        If rsTemp.RecordCount = 1 And intPCount >= 1 Then
            '组合图象
            ResizeRegion intPCount, rsTemp("W"), rsTemp("H"), intRows, intCols
            Set dcmResultImage = AssembleImage(dcmImages, intRows, intCols, rsTemp("H"), rsTemp("W"))
            dcmResultImage.FileExport Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "=")), "JPEG"
        End If
        
        '获取自定义报表中的图象定义
        intPCount = 0
        gstrSQL = "Select b.名称 From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = 1" & vbNewLine & _
        "       Order By b.名称" 'Trunc(b.y/567),Trunc(b.x/567)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参数", strBillNo)
        Do While Not rsTemp.EOF
            If aryPara(intPCount) = "" Then Exit Do '报表中的图形比报告中多
            '分别装载标记图和报告图像
            If InStr(rsTemp!名称, "标记") <> 0 Then
                If aryFlagPara(0) <> "" Then aryFlagPara(0) = rsTemp!名称 & "=" & aryFlagPara(0)
            Else
                aryPara(intPCount) = rsTemp!名称 & "=" & aryPara(intPCount)
                intPCount = intPCount + 1
                If intPCount > UBound(aryPara) Then Exit Do
            End If
            rsTemp.MoveNext
        Loop
        For intLoop = intPCount To UBound(aryPara) '报表中的图形比报告中少
            If aryPara(intLoop) Like "*=*" Then aryPara(intLoop) = ""
        Next
            
        '调用报表
        Call mobjReport.ReportOpen(gcnOracle, glngSys, strBillNo, Me, _
            "NO=" & strExseNo, "性质=" & intExseKind, "医嘱ID=" & CLng(Document.EPRPatiRecInfo.医嘱id), aryFlagPara(0), _
            aryPara(0), aryPara(1), aryPara(2), aryPara(3), aryPara(4), aryPara(5), _
            aryPara(6), aryPara(7), aryPara(8), aryPara(9), aryPara(10), aryPara(11), _
            aryPara(12), aryPara(13), aryPara(14), aryPara(15), aryPara(16), _
            aryPara(17), aryPara(18), aryPara(19), IIf(blnPreview, 1, 2))
    Else
        Set mfrmPreview = New frmPrintPreview
        Call mfrmPreview.DoSingleDocPreview(Me.Editor1, Me, Me.Document, blnClearMode, blnPreview, blnNoAsk)
        Unload mfrmPreview
        Set mfrmPreview = Nothing
    End If
    '=================================================================================================
    zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    PrintEPRDoc = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'################################################################################################################
'## 功能：  添加提纲到指定位置
'################################################################################################################
Public Function InsertCompend(ByVal lStart As Long, ByVal lEnd As Long, ByRef objCompend As cEPRCompend, Optional blnFirstIns As Boolean = True) As Boolean
    Dim strTmp As String, lLen As Long, lngKey As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean

    With Editor1
        bBeteenKeys = IsBetweenAnyKeys(Editor1, lStart + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then
            '当前位置在某个关键字对之间，则不允许插入提纲！
            InsertCompend = False
            Exit Function
        End If
        '添加提纲
        lLen = Len(objCompend.名称)
        objCompend.InsertIntoEditor Editor1, , blnFirstIns, Me.Document
        mfrmCompends_NodeSelected objCompend.预制提纲ID
    End With
    InsertCompend = True
End Function

'################################################################################################################
'## 功能：  保存修改后的提纲
'################################################################################################################
Public Function ModifyCompend(objCompend As cEPRCompend) As Boolean
    Dim strTmp As String, lIndex As Long, lLen As Long
    Dim lKey As Long, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, sKeyType As String, bNeeded As Boolean, bFinded As Boolean

    With Editor1
        If .ViewMode <> cprNormal Then ModifyCompend = False: Exit Function
        lKey = objCompend.Key
        bFinded = FindKey(Editor1, "O", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded = False Then
            ModifyCompend = False
            Exit Function
        End If
        .Freeze
        .Tag = "禁止同步"
        .ForceEdit = True
        .Range(lKSS, lKEE) = ""
        objCompend.InsertIntoEditor Editor1, lKSS, False, Me.Document

        .Tag = ""
        .ForceEdit = False
        .UnFreeze
    End With
    ModifyCompend = True
End Function

'################################################################################################################
'## 功能：  删除一个提纲
'################################################################################################################
Public Function DeleteOutline(lKey As Long) As Boolean
    Dim strTmp As String, lIndex As Long, lLen As Long, lLevel As Long, lNextKey As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, sKeyType As String, bNeeded As Boolean, bFinded As Boolean

    With Editor1
        If .ViewMode <> cprNormal Then DeleteOutline = False: Exit Function
        '确认文档中存在该提纲
        bFinded = FindKey(Editor1, "O", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded = False Then
            DeleteOutline = False
            Exit Function
        End If
        Dim lngR As Long

        If Document.Compends("K" & lKey).保留对象 = True And Me.Document.EditType <> cprET_病历文件定义 Then
            MsgBox "不能删除保留提纲！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If

'        If Document.Compends("K" & lKey).预制提纲ID <> 0 Then
'            lngR = MsgBox("确认删除：是否同时删除预制提纲 [" & Document.Compends("K" & lKey).名称 & "] 及其所有下级内容？" _
'                , vbYesNo + vbQuestion + vbDefaultButton2, "确认删除")
'            If lngR = vbNo Then lngR = vbCancel
'        Else
        lngR = MsgBox("确认删除：是否同时删除提纲 [" & Document.Compends("K" & lKey).名称 & "] 的所有下级内容？" & vbCrLf & _
            "注：如果保留，则合并到一级提纲的内容中。", vbYesNoCancel + vbQuestion + vbDefaultButton3, "确认删除")
'        End If
        If lngR = vbNo Then
            .InProcessing = True
            .Tag = "DeleteOutline"
            .Freeze
            .ForceEdit = True
            lLevel = Document.Compends("K" & lKey).Level
            Document.Compends.Remove "K" & lKey
            .Range(lKSS, lKEE) = ""
            .ForceEdit = False
            .Tag = ""
            .UnFreeze
            .InProcessing = False
            Document.Compends.CheckValidParentKeys '检查父Key的有效性！
            Document.Compends.FillTree mfrmCompends.Tree
            DeleteOutline = True
            Editor1.SelLength = 0
        ElseIf lngR = vbYes Then
            .InProcessing = True
            .Tag = "DeleteOutline"
            .Freeze
            .ForceEdit = True
            lLevel = Document.Compends("K" & lKey).Level
            Document.Compends.Remove "K" & lKey
            Dim i As Long, sText As String
            sText = Editor1.Text
            lLen = Len(sText)
            i = lKEE
LL1:
            '处理中间包含的其他元素，并删除！

            i = InStr(i, sText, "OS", vbTextCompare)
            If i <> 0 Then
                If .Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                    i = i + 1
                    GoTo LL1
                End If
                lNextKey = Val(.Range(i + 2, i + 10))
                If Document.Compends("K" & lNextKey).父Key = 0 Then
                    Document.Compends("K" & lNextKey).Level = 1
                End If
                If Document.Compends("K" & lNextKey).Level > lLevel Then
                    '层次小于当前层次，表明是子提纲！！！
                    Document.Compends.Remove "K" & lNextKey
                    i = i + 1
                    GoTo LL1
                End If
                Call ClearObjectsInArea(sText, lKSS, i - 1)
                .Range(lKSS, i - 1) = ""
            Else
                Call ClearObjectsInArea(sText, lKSS, lLen)
                .Range(lKSS, lLen) = ""
            End If
            .ForceEdit = False
            .UnFreeze
            .Tag = ""
            .InProcessing = False
            Document.Compends.CheckValidParentKeys '检查父Key的有效性！
            Document.Compends.FillTree mfrmCompends.Tree
            DeleteOutline = True
            Editor1.SelLength = 0
        Else
            '不进行删除操作
        End If
        Call RefSentenceList
    End With
End Function

'################################################################################################################
'   用途：  清除区间(lngStart,lngEnd)内的所有图片、诊治要素和表格对象。（用于清空某个提纲内部所有对象）
'################################################################################################################
Private Sub ClearObjectsInArea(ByRef StrText As String, ByVal lngStart As Long, ByVal lngEnd As Long)
    Dim lLen As Long, i As Long, lngKey As Long, blnForce As Boolean
    With Editor1
        blnForce = .ForceEdit
        .Tag = "ClearObjectinArea"
        .Freeze
        .ForceEdit = True
        lLen = Len(StrText)
        '处理中间包含的其他元素，并删除！
        i = IIf(lngStart = 0, 1, lngStart)
        i = InStr(i, StrText, "ES", vbTextCompare)
        Do While i > lngStart And i < lngEnd
            If .Range(i - 1, i).Font.Hidden Then    '若为关键字，必须是隐藏且受保护的。
                lngKey = Val(.Range(i + 2, i + 10))
                Document.Elements.Remove "K" & lngKey
            End If
            i = i + 1
            i = InStr(i, StrText, "ES", vbTextCompare)
        Loop
        i = IIf(lngStart = 0, 1, lngStart)
        i = InStr(i, StrText, "PS", vbTextCompare)
        Do While i > lngStart And i < lngEnd
            If .Range(i - 1, i).Font.Hidden Then    '若为关键字，必须是隐藏且受保护的。
                lngKey = Val(.Range(i + 2, i + 10))
                Document.Pictures.Remove "K" & lngKey
            End If
            i = i + 1
            i = InStr(i, StrText, "PS", vbTextCompare)
        Loop
        i = IIf(lngStart = 0, 1, lngStart)
        i = InStr(i, StrText, "TS", vbTextCompare)
        Do While i > lngStart And i < lngEnd
            If .Range(i - 1, i).Font.Hidden Then    '若为关键字，必须是隐藏且受保护的。
                lngKey = Val(.Range(i + 2, i + 10))
                Document.Tables.Remove "K" & lngKey
            End If
            i = i + 1
            i = InStr(i, StrText, "TS", vbTextCompare)
        Loop
        .ForceEdit = blnForce
        .Tag = ""
        .UnFreeze
    End With
End Sub

'################################################################################################################
'   用途：  插入图片。
'################################################################################################################
Public Function InsertPicture(bytPicType As Integer, objPic As StdPicture, lWidth As Long, lHeight As Long, Optional strOther As String) As Boolean
    'IsMarked:是否是标记图
    'objPic  :图片数据
    'lRow    :用于表格中的图片绑定，表示行
    'lCol    :用于表格中的图片绑定，表示列
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim lngKey As Long
    If tblThis.Visible Then '        '表格中的插入图形
        If Val(tblThis.Tag) > 0 Then
            '将图片对象保存到类对象中
            If Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag) = 0 Then
                lngKey = Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures.Add
            Else
                lngKey = Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag)
            End If
            Set Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).OrigPic = objPic
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).Width = lWidth
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).Height = lHeight
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).OrigWidth = lWidth
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).OrigHeight = lHeight
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).PictureType = bytPicType
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).Row = tblThis.Cells("K" & tblThis.SelectedCellKey).Row
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).Col = tblThis.Cells("K" & tblThis.SelectedCellKey).Col
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).内容文本 = strOther
            
            '保存到单元格中
            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = ""
            tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lngKey
            tblThis.Cells("K" & tblThis.SelectedCellKey).Picture = objPic
            tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = ""
            tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
            tblThis.Modified = True
            tblThis.Refresh
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    ElseIf ucPacsImgCanvas1.Visible Then '在报告框中插入图形
        If ucPictureEditor1.Visible Then ucPictureEditor1.Modified = False: ucPictureEditor1.CloseMe
        If ucPacsImgCanvas1.MarkedPicPosition = 0 Then ucPacsImgCanvas1.MarkedPicPosition = 1
        ucPacsImgCanvas1.AddMarkedPicture objPic, ucPacsImgCanvas1.MarkedPicPosition
    Else '内容中的图
        bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
        If bInKeys Then InsertPicture = False: Exit Function    '保证不能插入关键字内部
        '将图片对象保存到类对象中
        lngKey = Document.Pictures.Add()
        Set Document.Pictures("K" & lngKey).OrigPic = objPic
        Document.Pictures("K" & lngKey).Width = lWidth
        Document.Pictures("K" & lngKey).Height = lHeight
        Document.Pictures("K" & lngKey).OrigWidth = lWidth
        Document.Pictures("K" & lngKey).OrigHeight = lHeight
        Document.Pictures("K" & lngKey).PictureType = bytPicType
        Document.Pictures("K" & lngKey).InsertIntoEditor Editor1
        Document.Pictures("K" & lngKey).内容文本 = strOther
    End If
    InsertPicture = True
End Function

'################################################################################################################
'## 功能：  显示诊治要素编辑器  '○●□■
'################################################################################################################
Private Sub ShowEleEditor(KeyAscii As Integer, Shift As Integer)
    On Error Resume Next
    If Editor1.ViewMode <> cprNormal Then Exit Sub
'    If Me.Editor1.AuditMode Then Exit Sub
    glngCurEleKey = 0
'    picSmartSignal.Visible = False
    '判断当前位置是否在 CS 与 CE 之间：
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    bBeteenKeys = IsBetweenKeys(Editor1, Editor1.Selection.StartPos + 1, "E", lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    
    
    '签名要素禁止编辑
    '------------------------------------------------------------------------------------------------------------------
    If Document.Elements("K" & lKey).替换域 = 1 Then
        Select Case Document.Elements("K" & lKey).要素名称
        Case "经治医师签名", "主治医师签名", "主任医师签名"
            Exit Sub
        End Select
    End If
    
    Dim pt As POINTAPI
    pt.X = 0
    pt.y = 0
    ClientToScreen Editor1.OriginRTB.hwnd, pt

    If bBeteenKeys Then
        '此时记录坐标位置
        If Me.Editor1.AuditMode Then
            If Document.Elements("K" & lKey).开始版 < Me.Document.目标版本 Then Exit Sub
        End If
        glngCurEleKey = lKey
        Dim lLeft As Long, lTOp As Long, lRight As Long
        '获取起始位置坐标
        Editor1.Range(Editor1.Selection.StartPos, Editor1.Selection.StartPos + 1).GetPoint cprGPStart + cprGPLeft + cprGPBottom, lLeft, lTOp

        '显示编辑控件
        If Editor1.Range(lKEE, lKEE + 2) = vbCrLf Then
            Document.Elements("K" & lKey).是否换行 = True
        Else
            Document.Elements("K" & lKey).是否换行 = False
        End If
        If Document.Elements("K" & lKey).替换域 = 2 Then
            '字典项目
            mfrmDicSelect.ShowMe Document.Elements("K" & lKey).要素名称, pt.X * Screen.TwipsPerPixelX + lLeft, _
                pt.y * Screen.TwipsPerPixelY + lTOp, vbModeless, Me, Document.Elements("K" & lKey).内容文本
        Else
            '诊治要素
            mfrmModElement.Tag = Editor1.Selection.StartPos
            mfrmModElement.ShowMe Document.Elements("K" & lKey), _
                pt.X * Screen.TwipsPerPixelX + lLeft, _
                pt.y * Screen.TwipsPerPixelY + lTOp, IIf(mintStyle = -1, 0, mintStyle), Me, Me.Document.EditType
        End If
        If Chr(KeyAscii) <> " " And Chr(KeyAscii) <> Chr(13) And KeyAscii <> 0 Then SendKeys Chr(KeyAscii)
    Else
        '否则，定位到第一个诊治要素的位置
        bBeteenKeys = FindNextKey(Editor1, 1, "E", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bBeteenKeys Then
            Editor1.Range(lKSE, lKES).Selected
        End If

        glngCurEleKey = 0
    End If
End Sub
Private Sub mfrmSentenceDetailed_RowDblClick(ByVal lngSentenceID As Long)
    '双击插入示范词句
    If Me.Editor1.ViewMode <> cprNormal Or Me.Editor1.ReadOnly Then Exit Sub
    If Me.Editor1.Selection.Font.Protected And tblThis.Visible = False Then Exit Sub
    If tblThis.Visible Then
        If tblThis.SelectedCellKey > 0 Then If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected Then Exit Sub
    End If

    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys And tblThis.Visible = False Then Exit Sub
    Dim blnForce As Boolean
    Dim lngKey As Long, lngStart As Long, lngLen As Long, strTmp As String, rsTemp As New ADODB.Recordset
    Dim lngStartPos As Long, lngEndPos As Long, sText As String

    If lngSentenceID <= 0 Then Exit Sub
    mfrmSentenceDetailed.Tag = lngSentenceID    '保留原来的记录ID

    '词句内容恢复
    gstrSQL = "Select 词句id, 排列次序, 内容性质, 内容文本, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 要素值域, 输入形态, 对象属性" & vbNewLine & _
                "From 病历词句组成" & vbNewLine & _
                "Where 词句id = [1]" & vbNewLine & _
                "Order By 排列次序"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", lngSentenceID)
    With Editor1
        .Tag = "mfrmSentenceDetailed_RowDblClick"
        .Freeze
        blnForce = .ForceEdit
        .ForceEdit = True
        lngStartPos = .Selection.StartPos
        Do While Not rsTemp.EOF
            Select Case rsTemp("内容性质")
            Case 0 '自由文字
                '恢复RTF内容
                lngStart = .Selection.StartPos
                strTmp = NVL(rsTemp("内容文本"))
                lngLen = Len(strTmp)

                If tblThis.Visible Then
                    sText = sText & strTmp
                Else
                    .Range(lngStart, lngStart) = strTmp
                    .Range(lngStart, lngStart + lngLen).Font.Protected = False
                    .Range(lngStart, lngStart + lngLen).Font.Hidden = False
                    .Range(lngStart, lngStart + lngLen).Font.ForeColor = IIf(Me.Editor1.AuditMode, GetCharColor(Me.Document.目标版本, 0), tomAutoColor)
                    .Range(lngStart, lngStart + lngLen).Font.Strikethrough = False
                    .Range(lngStart, lngStart + lngLen).Font.BackColor = tomAutoColor
                    .Range(lngStart + lngLen, lngStart + lngLen).Selected
                End If
            Case 1, 2 '1-临时诊治要素,2-固定诊治要素
                If tblThis.Visible Then
                    If NVL(rsTemp("替换域"), 0) = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                        strTmp = GetReplaceEleValue(NVL(rsTemp("要素名称")), Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.主页ID, Me.Document.EPRPatiRecInfo.病人来源, Me.Document.EPRPatiRecInfo.医嘱id, Me.Document.EPRPatiRecInfo.婴儿)
                    Else
                        strTmp = "{" & NVL(rsTemp("要素名称")) & "}"
                    End If
                    sText = sText & strTmp
                Else
                    lngStart = .Selection.StartPos
                    lngKey = Me.Document.Elements.Add
                    Me.Document.Elements("K" & lngKey).ID = 0       '而非： NVL(rsTemp("词句ID"), 0) ，这个ID值不同！！！
                    Me.Document.Elements("K" & lngKey).内容文本 = NVL(rsTemp("内容文本"))
                    Me.Document.Elements("K" & lngKey).要素名称 = NVL(rsTemp("要素名称"))
                    Me.Document.Elements("K" & lngKey).诊治要素ID = NVL(rsTemp("诊治要素ID"), 0)
                    Me.Document.Elements("K" & lngKey).替换域 = NVL(rsTemp("替换域"), 0)
                    Me.Document.Elements("K" & lngKey).要素类型 = NVL(rsTemp("要素类型"), 0)
                    Me.Document.Elements("K" & lngKey).要素长度 = NVL(rsTemp("要素长度"), 0)
                    Me.Document.Elements("K" & lngKey).要素小数 = NVL(rsTemp("要素小数"), 0)
                    Me.Document.Elements("K" & lngKey).要素单位 = NVL(rsTemp("要素单位"))
                    Me.Document.Elements("K" & lngKey).要素表示 = NVL(rsTemp("要素表示"), 0)
                    Me.Document.Elements("K" & lngKey).要素值域 = NVL(rsTemp("要素值域"))
                    Me.Document.Elements("K" & lngKey).输入形态 = NVL(rsTemp("输入形态"), 0)
                    Me.Document.Elements("K" & lngKey).是否换行 = False
                    Me.Document.Elements("K" & lngKey).对象属性 = NVL(rsTemp!对象属性)
                    If Me.Document.Elements("K" & lngKey).替换域 = 1 And (Me.Document.EditType = cprET_单病历编辑 Or Me.Document.EditType = cprET_单病历审核) Then
                        Me.Document.Elements("K" & lngKey).内容文本 = GetReplaceEleValue(Me.Document.Elements("K" & lngKey).要素名称, _
                            Me.Document.EPRPatiRecInfo.病人ID, _
                            Me.Document.EPRPatiRecInfo.主页ID, _
                            Me.Document.EPRPatiRecInfo.病人来源, _
                            Me.Document.EPRPatiRecInfo.医嘱id, _
                            Me.Document.EPRPatiRecInfo.婴儿)
    '                    If Me.Document.Elements("K" & lngKey).内容文本 = "" Then Me.Document.Elements("K" & lngKey).内容文本 = "    "
                    End If
                    Me.Document.Elements("K" & lngKey).开始版 = Me.Document.目标版本
                    Me.Document.Elements("K" & lngKey).InsertIntoEditor Editor1, lngStart
                End If
            End Select
            rsTemp.MoveNext
        Loop
        lngEndPos = .Selection.StartPos
        .ForceEdit = False
        If tblThis.Visible Then
            sText = tblThis.Cells("K" & tblThis.SelectedCellKey).Text & sText
            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = sText
            tblThis.Modified = True
            tblThis.Refresh False, True, tblThis.SelectedCellKey
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            .Range(lngEndPos, lngEndPos).Selected
        End If
        .UnFreeze
        .Tag = ""
        .SetFocus
    End With
    Call RecountPage
    If Me.Editor1.Visible And Me.Editor1.Enabled And tblThis.Visible = False Then Me.Editor1.SetFocus
End Sub

Private Sub mfrmStyleMan_DblClick(ByVal lngStyleCode As Long)
    '改变当前选中内容的段落样式
    SetCommonStyle Me.Editor1, lngStyleCode, Me.Editor1.Selection.StartPos, Me.Editor1.Selection.EndPos, True
    Call RecountPage
End Sub

Private Sub picPatiInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If X > 0 And X < picPatiInfo.ScaleWidth And y > 0 And y < picPatiInfo.ScaleHeight Then
        If picPatiInfo.Tag = "" Then
            SetCapture picPatiInfo.hwnd
            picPatiInfo.Cls
            picPatiInfo.BackColor = &HD2BDB6    ' &HD8D5D4 ' &HD2BDB6
            picPatiInfo.Line (0, 0)-(picPatiInfo.ScaleWidth - Screen.TwipsPerPixelX, picPatiInfo.ScaleHeight - Screen.TwipsPerPixelY), &H6A240A, B
            picPatiInfo.Tag = "Captured"
        End If
    Else
        ReleaseCapture
        picPatiInfo.Cls
        picPatiInfo.BackColor = &H8000000F
        picPatiInfo.Line (0, 0)-(picPatiInfo.ScaleWidth - Screen.TwipsPerPixelX, picPatiInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
        picPatiInfo.Tag = ""
    End If
End Sub

Private Sub picPatiInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    ReleaseCapture
    picPatiInfo.Cls
    picPatiInfo.BackColor = &H8000000F
    picPatiInfo.Line (0, 0)-(picPatiInfo.ScaleWidth - Screen.TwipsPerPixelX, picPatiInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
    picPatiInfo.Tag = ""
End Sub

Private Sub picPatiInfo_Resize()
'    ReleaseCapture
    picPatiInfo.Cls
    picPatiInfo.BackColor = &H8000000F
    picPatiInfo.Line (0, 0)-(picPatiInfo.ScaleWidth - Screen.TwipsPerPixelX, picPatiInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
    picPatiInfo.Tag = ""
End Sub

'################################################################################################################
'   用途：  动态更新工具栏“颜色”图标。
'################################################################################################################
Private Sub SetColorIcon(Key As String, ID As Long, COLOR As OLE_COLOR)
    Dim ctlPictureBox As VB.PictureBox
    Set ctlPictureBox = Controls.Add("VB.PictureBox", "ctlPictureBox1")
    Dim ListImage As ListImage
    Set ListImage = imgColor.ListImages(Key)

    ctlPictureBox.AutoRedraw = True
    ctlPictureBox.AutoSize = True
    ctlPictureBox.BackColor = imgColor.MaskColor

    ctlPictureBox.Picture = ListImage.ExtractIcon

    If COLOR = vbWhite Then COLOR = RGB(254, 254, 254)
    ctlPictureBox.Line (1, ctlPictureBox.Height * 0.6)-(ctlPictureBox.Width, ctlPictureBox.Height), COLOR, BF
    ctlPictureBox.Refresh

    'Replace icon
    imgColor.ListImages.Remove imgColor.ListImages(Key).Index
    imgColor.ListImages.Add 1, Key, ctlPictureBox.Image
'    Set imgColor.ListImages(Key).Picture = ctlPictureBox.Image

    'OK Now replace Tag property
    imgColor.ListImages(1).Tag = ID

    cbrThis.AddImageList imgColor
    cbrThis.RecalcLayout

        Set ctlPictureBox.Picture = Nothing
    Me.Controls.Remove ctlPictureBox
    Set ctlPictureBox = Nothing
End Sub

'################################################################################################################
'   用途：  刷新病人信息
'################################################################################################################
Public Sub RefreshPatiInfo()
    If Me.Document.EditType <> cprET_单病历编辑 And Me.Document.EditType <> cprET_单病历审核 Then Exit Sub

    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo errHand
    If Me.Document.EPRPatiRecInfo.病人来源 <> 2 Then
        gstrSQL = "Select 姓名,身份证号,Rpad('门诊号:' || 门诊号, 18) || Rpad('姓名:' || 姓名, 18) ||" & vbNewLine & _
         "        Rpad('性别:' || 性别, 10) || Rpad('年龄:' || 年龄,10) As 信息, 医保号,性别 " & vbNewLine & _
         "From 病人信息" & vbNewLine & _
         "Where 病人id = [1]"
    ElseIf (Document.EPRPatiRecInfo.医嘱id <> 0 And Document.EPRPatiRecInfo.婴儿 <> 0) Then
        gstrSQL = "Select Decode([2], 2, RPad('住院号:' || c.住院号, 18) || RPad('床号:' || c.出院病床, 15), RPad('门诊号:' || a.门诊号, 18)) ||" & vbNewLine & _
                    "        RPad('姓名:' || Nvl(b.婴儿姓名, a.姓名 || '之子'), 18) || RPad('性别:' || b.婴儿性别, 10) || '年龄:' ||" & vbNewLine & _
                    "        To_Char(b.出生时间, 'YYYY-MM-DD HH24:MI:SS') As 信息, A.医保号, b.婴儿性别 性别,a.姓名,a.身份证号" & vbNewLine & _
                    "From 病人信息 A, 病案主页 C, 病人新生儿记录 B" & vbNewLine & _
                    "Where c.病人id = [1] And c.主页id = [3] And a.病人id = c.病人id And c.病人id = b.病人id And c.主页id = b.主页id And b.序号 = [4]"
    ElseIf (Document.EPRPatiRecInfo.婴儿 = 0) Then
           gstrSQL = "Select Decode([2], 2, RPad('住院号:' || b.住院号, 18) || RPad('床号:' || b.出院病床, 15), RPad('门诊号:' || a.门诊号, 18)) ||" & vbNewLine & _
                        "        RPad('姓名:' || a.姓名, 18) || RPad('性别:' || a.性别, 10) || RPad('年龄:' || a.年龄, 10) As 信息, a.医保号, a.性别,a.姓名,a.身份证号" & vbNewLine & _
                        "From 病人信息 A, 病案主页 B" & vbNewLine & _
                        "Where b.病人id = [1] And b.主页id = [3] And a.病人id = b.病人id"
    Else
        gstrSQL = "Select Decode([2], 2, RPad('母亲住院号:' || c.住院号, 18) || RPad('母亲床号:' || c.出院病床, 15), RPad('母亲门诊号:' || a.门诊号, 18)) ||" & vbNewLine & _
                    "        RPad('姓名:' || Nvl(b.婴儿姓名, a.姓名 || '之婴' || b.序号), 30) || RPad('性别:' || Nvl(b.婴儿性别, '未知'), 10) || '年龄:' ||" & vbNewLine & _
                    "        To_Char(b.出生时间, 'YYYY-MM-DD HH24:MI:SS') As 信息, a.医保号, a.性别,a.姓名,a.身份证号" & vbNewLine & _
                    "From 病人信息 A, 病案主页 C, 病人新生儿记录 B" & vbNewLine & _
                    "Where c.病人id = [1] And c.主页id = [3] And a.病人id = c.病人id And c.病人id = b.病人id And c.主页id = b.主页id And b.序号 = [4]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.病人来源, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.婴儿)
    If rsTemp.RecordCount > 0 Then
        mPatiInfor.姓名 = "" & rsTemp!姓名
        mPatiInfor.身份证号 = "" & rsTemp!身份证号
        Me.lblPatiInfo.Caption = "" & rsTemp!信息
        Me.lblPatiIns(1).Caption = "" & rsTemp!医保号
        mstrSex = "" & rsTemp!性别
        mfrmDocksymbol.HideSomeThing IIf(InStr(mstrSex, "男") > 0, 1, IIf(InStr(mstrSex, "女") > 0, 2, 0))
    Else
        mstrSex = ""
        Me.lblPatiInfo.Caption = ""
        Me.lblPatiIns(1).Caption = ""
    End If
    Err = 0: On Error Resume Next
    lblPatiIns(0).Left = lblPatiInfo.Left + lblPatiInfo.Width + 50
    lblPatiIns(1).Left = lblPatiIns(0).Left + lblPatiIns(0).Width
    lblPatiState(0).Left = lblPatiIns(1).Left + lblPatiIns(1).Width
    lblPatiState(1).Left = lblPatiState(0).Left + lblPatiState(0).Width
    Err.Clear: On Error GoTo errHand
    
    If Me.Document.EPRPatiRecInfo.医嘱id = 0 Then
        Me.lblPatiState(0).Caption = "病况:"
        Select Case Me.Document.EPRPatiRecInfo.病人来源
        Case cprPF_门诊
            gstrSQL = "Select r.急诊 From 病人挂号记录 r Where r.Id = [1] and r.记录性质=1  and r.记录状态=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", Me.Document.EPRPatiRecInfo.主页ID)
            If rsTemp.RecordCount > 0 Then
                Me.lblPatiState(1).Caption = IIf(Val("" & rsTemp!急诊) = 1, "急", "")
            Else
                Me.lblPatiState(1).Caption = ""
            End If
        Case cprPF_住院
            gstrSQL = "Select 入院病况, 出院日期, 出院方式 From 病案主页 Where 病人id = [1] And Nvl(主页id, 0) = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.主页ID)
            If rsTemp.RecordCount > 0 Then
                If IsNull(rsTemp!出院日期) Then
                    If "" & rsTemp!入院病况 = "一般" Then
                        Me.lblPatiState(1).ForeColor = Me.lblPatiState(0).ForeColor
                    Else
                        Me.lblPatiState(1).ForeColor = RGB(255, 0, 0)
                    End If
                    Me.lblPatiState(1).Caption = "" & rsTemp!入院病况
                Else
                    If "" & rsTemp!出院方式 <> "死亡" Then
                        Me.lblPatiState(1).ForeColor = Me.lblPatiState(0).ForeColor
                    Else
                        Me.lblPatiState(1).ForeColor = RGB(255, 0, 0)
                    End If
                    Me.lblPatiState(1).Caption = rsTemp!出院方式 & "(出院)"
                End If
            Else
                Me.lblPatiState(1).Caption = ""
            End If
        End Select
    Else
        Me.lblPatiState(0).Caption = "要求:"
        gstrSQL = "Select 紧急标志 From 病人医嘱记录 Where Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取信息", Me.Document.EPRPatiRecInfo.医嘱id)
        If rsTemp.RecordCount > 0 Then
            Me.lblPatiState(1).Caption = IIf(Val("" & rsTemp!紧急标志) = 1, "急", "")
        Else
            Me.lblPatiState(1).Caption = ""
        End If
    End If

    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'################################################################################################################
'   用途：  导入本次就诊医嘱到编辑器中当前位置
'################################################################################################################
Public Function ImportDocAdvice() As Boolean
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
Dim rsTemp As New ADODB.Recordset

On Error GoTo errHand
    With Me.Editor1
        If .Selection.Font.Protected Then Exit Function
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then Exit Function
        gstrSQL = "Select ID,长度,小数,单位,替换域,必填,动态域,类型 From 诊治所见项目 where 中文名='本次医嘱'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取要素")
        
        lKey = Me.Document.Elements.Add
        With Me.Document.Elements("K" & lKey)
            .要素名称 = "本次医嘱"
            .诊治要素ID = rsTemp!ID
            .要素类型 = rsTemp!类型
            .要素长度 = NVL(rsTemp!长度, 2)
            .要素小数 = NVL(rsTemp!小数, 2)
            .要素单位 = NVL(rsTemp!单位, 2)
            .替换域 = NVL(rsTemp!替换域, 0)
            .必填 = NVL(rsTemp!必填, 0)
            .动态域 = NVL(rsTemp!动态域, 0)
            .内容文本 = GetReplaceEleValue(.要素名称, Document.EPRPatiRecInfo.病人ID, Document.EPRPatiRecInfo.主页ID, Document.EPRPatiRecInfo.病人来源, Document.EPRPatiRecInfo.医嘱id, Me.Document.EPRPatiRecInfo.婴儿)
            .开始版 = Me.Document.目标版本
            .InsertIntoEditor Me.Editor1, , True
        End With
    End With
    Me.Document.EleToString Me.Editor1, Me.Document.Elements("K" & lKey)
    ImportDocAdvice = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 '################################################################################################################
'   用途：  插入一个Pacse图片组（表格）
'################################################################################################################
Public Sub InsertPacsPicTable()
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then Exit Sub
    
    lKey = Me.Document.Tables.Add
    Me.Document.Tables("K" & lKey).TableType = tte_报告图片组
    ucPacsImgCanvas1.ReadPicturesFromTable Me.Document.Tables("K" & lKey)
    
    Dim frmT As New frmTablePicCreator
    Me.Document.Tables("K" & lKey).InsertIntoEditor Editor1, , , True
    Unload frmT
    Set frmT = Nothing
End Sub

Private Sub txtContent_Change()
    If txtContent.Text <> txtContent.Tag Then
        mblnFBContentChanged = True
    End If
    txtContent.Tag = txtContent.Text
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/<>", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtPenInput_Change()
    If Editor1.ReadOnly Then GoTo LL
    If Me.Editor1.Selection.Font.Protected = False And Me.Editor1.Selection.Font.Hidden = False And txtPenInput.Tag = "" Then
        txtPenInput.Tag = "InProcessing"
        Me.Editor1.ForceEdit = True
        If Me.Editor1.AuditMode Then
            On Error Resume Next
            Editor1.Range(Editor1.Selection.StartPos + Len(Editor1.Selection.Text), Editor1.Selection.StartPos + Len(Editor1.Selection.Text)).Selected
            Me.Editor1.SelLength = 0
            Me.Editor1.OriginRTB.SelColor = Me.Document.GetNewCharColor(Me.Editor1.OriginRTB.SelColor)
            Me.Editor1.OriginRTB.SelStrikeThru = False    '去掉删除线
            Me.Editor1.OriginRTB.SelUnderline = False     '去掉下划线
            Me.Editor1.SelText = txtPenInput.Text
        Else
            Me.Editor1.SelText = txtPenInput.Text
        End If
        Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos).Selected
        Me.Editor1.ForceEdit = False
    End If
LL:
    txtPenInput.Text = ""
    txtPenInput.Tag = ""
End Sub

Private Sub txtPenInput_GotFocus()
    txtPenInput.Tag = "InProcessing"
    txtPenInput.Text = ""
    txtPenInput.Tag = ""
End Sub

Private Sub txtPenInput_KeyPress(KeyAscii As Integer)
    If Editor1.ReadOnly Then Exit Sub
    Select Case KeyAscii
    Case vbKeyBack
        If Me.Editor1.AuditMode Then Exit Sub
        Dim lngStart As Long, lngEnd As Long
        lngStart = Editor1.Selection.StartPos
        lngEnd = Editor1.Selection.EndPos
        If lngStart <> lngEnd Then
            If Me.Editor1.Range(lngStart, lngEnd).Font.Protected = False And Me.Editor1.Range(lngStart, lngEnd).Font.Hidden = False Then
                Editor1.TOM.TextDocument.Range(lngStart, lngEnd) = ""
            End If
        Else
            If Me.Editor1.Range(lngStart - 1, lngStart).Font.Protected = False And Me.Editor1.Range(lngStart - 1, lngStart).Font.Hidden = False Then
                Editor1.TOM.TextDocument.Range(lngStart - 1, lngStart) = ""
            End If
        End If
    Case vbKeyEscape
        SendKeys "{F11}"
    End Select
End Sub

Private Sub mfrmMultiDocView_RequestModifyDoc(ByVal lngFileID As Long)
'请求编辑指定文件
    Dim strPrivs As String
    If Editor1.Modified Then
        Dim r As Long
        r = MsgBox("当前文件已经被修改，是否先保存？", vbYesNoCancel + vbQuestion, gstrSysName)
        If r = vbCancel Then
            Exit Sub
        ElseIf r = vbYes Then
            If SaveEMRDoc = False Then Exit Sub
        ElseIf r = vbNo Then
            '
        End If
    End If
    '重新初始化Doc对象
    Me.Editor1.ReadOnly = False
    Me.Document.ClearAllIDs
    Me.Document.InitEPRDoc cprEM_修改, Me.Document.EditType, lngFileID, _
        Me.Document.EPRPatiRecInfo.病历种类, Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.主页ID, _
        Me.Document.EPRPatiRecInfo.婴儿, Me.Document.EPRPatiRecInfo.科室ID
    Me.Document.OpenEPRDoc Me.Editor1
    Me.Editor1.Modified = False
    Me.RefreshPatiInfo
    If Me.Document.EditType = cprET_单病历编辑 Then
        '修改
        Me.Editor1.ReadOnly = Not (Me.Document.EPRPatiRecInfo.最后版本 = 1 And Me.Document.EPRPatiRecInfo.签名级别 = cprSL_空白)
        If Not Me.Editor1.ReadOnly Then
            Select Case Me.Document.EPRFileInfo.种类
            Case cpr住院病历
                strPrivs = GetPrivFunc(glngSys, 1251)
                Me.Editor1.ReadOnly = Not (Me.Document.EPRPatiRecInfo.保存人 = gstrUserName Or InStr(1, strPrivs, "他人病历") > 0)
            Case cpr护理病历
                strPrivs = GetPrivFunc(glngSys, 1255)
                Me.Editor1.ReadOnly = Not (Me.Document.EPRPatiRecInfo.保存人 = gstrUserName Or InStr(1, strPrivs, "他人护理病历") > 0)
            End Select
        End If
    Else
        '审阅
        Me.Editor1.ReadOnly = (Me.Document.EPRPatiRecInfo.最后版本 = 1 And Me.Document.EPRPatiRecInfo.签名级别 = cprSL_空白)
    End If
    Me.ShowMe Me.Document.mfrmParent, False
End Sub

Private Sub ucPacsImgCanvas1_Resize(lngWidth As Long, lngHeight As Long)
    Dim lKey As Long
    lKey = Val(ucPacsImgCanvas1.Tag)
    If lKey > 0 Then
        Document.Tables("K" & lKey).Refresh Editor1, ucPacsImgCanvas1.FinalPic, True
    End If
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, bFinded As Boolean, bNeeded As Boolean
    Dim lW As Long
    bFinded = FindKey(Editor1, "T", tblThis.Tag, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        Editor1.InProcessing = True
        Editor1.Range(lSE, lES).Selected
        Editor1.InProcessing = False
    End If
    Editor1.ResizeUIInterface lngWidth, lngHeight
End Sub

Private Sub ucPacsImgCanvas1_SelectedMarkedPic(lLeft As Long, lTOp As Long, lWidth As Long, lHeight As Long)
    '选中标记图，则显示图片编辑器
    If ucPacsImgCanvas1.Visible = False Then Exit Sub
    Dim lKey As Long
    lKey = Val(ucPacsImgCanvas1.Tag)
    If lKey > 0 Then
        ucPacsImgCanvas1.SavePictures
        If Me.Document.Tables("K" & lKey).Pictures.Count > 0 Then
            ucPictureEditor1.ShowMe Me, ucPacsImgCanvas1.hwnd, cbrThis, _
                Me.Document.Tables("K" & lKey).Pictures(1), _
                lLeft, lTOp, lWidth, lHeight, True, Me.Document.Tables("K" & lKey)
        End If
    End If
End Sub

Private Sub ucPacsImgCanvas1_SelectedPacsPic()
    '保存标记图
    If Not ucPacsImgCanvas1.mMarkedPicture Is Nothing Then
        If ucPictureEditor1.Visible Then
            ucPictureEditor1.CloseMe ucPacsImgCanvas1.mMarkedPicture
            ucPacsImgCanvas1.LayoutPictures False
        End If
    ElseIf ucPacsImgCanvas1.Visible And ucPictureEditor1.Visible Then
        ucPictureEditor1.Visible = False
    End If
End Sub

Private Sub ucPictureEditor1_DblClick()
Dim objPic As StdPicture, strPic As String, lKey As String
'编辑公式图并返回
    If ucPictureEditor1.mcPicture.PictureType <> EPRFormulaPicture Then Exit Sub
    
    strPic = ucPictureEditor1.mcPicture.内容文本
    lKey = ucPictureEditor1.mcPicture.Key
    ucPictureEditor1.Visible = False
    ucPictureEditor1.CloseMe
    Editor1.CloseUIInterface
    Call Editor1.ShowInsertSymbolDlg(False, IIf(InStr(mstrSex, "男") > 0, 1, IIf(InStr(mstrSex, "女") > 0, 2, 0)), False, strPic, objPic)
    If objPic Is Nothing Then Exit Sub
    
    Editor1.Tag = "特殊符号编辑"
    Call Document.Pictures("K" & lKey).DeleteFromEditor(Editor1)
    Call Document.Pictures.Remove("K" & lKey)
    InsertPicture EPRFormulaPicture, objPic, objPic.Width, objPic.Height, strPic
    Editor1.Tag = ""

End Sub
Private Sub ExportXML()
'导出到XML文件
Dim strF As String, i As Integer
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
Dim bFinded As Boolean
    Select Case Me.Document.EditType
    Case cprET_病历文件定义
        dlgThis.Filename = "定义_" & Me.Document.EPRFileInfo.名称 & ".xml"
    Case cprET_全文示范编辑
        dlgThis.Filename = "范文_" & Me.Document.EPRFileInfo.名称 & "_" & Me.Document.EPRDemoInfo.名称 & ".xml"
    Case cprET_单病历编辑, cprET_单病历审核
        dlgThis.Filename = "记录_" & Me.Document.EPRFileInfo.名称 & "(" & Me.Document.EPRPatiRecInfo.ID & "," & Me.Document.目标版本 & ").xml"
    End Select

    dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
    dlgThis.CancelError = True
    On Error GoTo LL
    dlgThis.ShowSave
    strF = dlgThis.Filename
    If gobjFSO.FileExists(strF) Then
        If MsgBox("该文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
    End If

    Editor1.Freeze
    Editor1.ForceEdit = True
    Editor1.Tag = "cbrThis_ExeCute"
    '在当前控件中处理图片和表格的替换
    Me.Document.PreSavingRTFText Me.Editor1

    If Me.Document.ExportToXMLFile(Me.Editor1, strF) Then
        MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
    End If
    '恢复图片和表格
    Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat
    For i = 1 To Me.Document.Pictures.Count
        bFinded = FindKey(Editor1, "P", Me.Document.Pictures(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            '还原图片
            Set ParaFmt = Editor1.Range(lKSE, lKES).Para.GetParaFmt
            Set FontFmt = Editor1.Range(lKSE, lKES).Font.GetFontFmt

            Me.Document.Pictures(i).是否换行 = False
            Editor1.Range(lKSS, lKEE).Text = ""
            Me.Document.Pictures(i).InsertIntoEditor Editor1, lKSS, True

            Editor1.Range(lKSE, lKES).Para.SetParaFmt ParaFmt
            Editor1.Range(lKSE, lKES).Font.SetFontFmt FontFmt
            Editor1.Range(lKSS, lKEE).Font.Protected = True
        End If
    Next
    For i = 1 To Me.Document.Tables.Count
        bFinded = FindKey(Editor1, "T", Me.Document.Tables(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            '还原表格
            Set ParaFmt = Editor1.Range(lKSE, lKES).Para.GetParaFmt
            Set FontFmt = Editor1.Range(lKSE, lKES).Font.GetFontFmt

            Me.Document.Tables(i).是否换行 = False
            Editor1.Range(lKSS, lKEE).Text = ""
            Me.Document.Tables(i).InsertIntoEditor Editor1, lKSS, , , True

            Editor1.Range(lKSE, lKES).Para.SetParaFmt ParaFmt
            Editor1.Range(lKSE, lKES).Font.SetFontFmt FontFmt
            Editor1.Range(lKSS, lKEE).Font.Protected = True
        End If
    Next

    Editor1.ForceEdit = False
    Editor1.UnFreeze
    Editor1.Tag = ""
LL:
End Sub
Public Function CommBar(ByVal BarId As Long) As XtremeCommandBars.CommandBar
    For Each CommBar In cbrThis
        If CommBar.BarId = BarId Then Exit Function
    Next
End Function
Private Sub ExecuteUnderLine(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal blnForce As Boolean)
    If Editor1.Selection.Font.Protected Or Editor1.Selection.Font.Hidden Then Exit Sub
    
    Dim BarUnderLine As CommandBarPopup, objControl As CommandBarControl
    Call AddUndoPoint  '手动缓存
    Editor1.ForceEdit = True
    Editor1.Tag = "cbrThis_ExeCute"
    Set BarUnderLine = CommBar(ID_BAR_FORMAT).FindControl(xtpControlSplitButtonPopup, ID_FORMAT_UNDERLINE)
    
    
    Select Case Control.ID
        Case ID_FORMAT_UNDERLINE
            If BarUnderLine.Checked Then
                Editor1.Selection.Font.Underline = cprNone
            Else
                Select Case True
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_THIN).Checked
                    Editor1.Selection.Font.Underline = cprHair
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_THICK).Checked
                    Editor1.Selection.Font.Underline = cprThick
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_WAVE).Checked
                    Editor1.Selection.Font.Underline = cprWave
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_DOT).Checked
                    Editor1.Selection.Font.Underline = cprDotted
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_DASH).Checked
                    Editor1.Selection.Font.Underline = cprDash
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_DASHDOT).Checked
                    Editor1.Selection.Font.Underline = cprDashDot
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_DASHDOT2).Checked
                    Editor1.Selection.Font.Underline = cprDashDotDot
                Case Else
                    Editor1.Selection.Font.Underline = cprHair
                End Select
            End If
        Case ID_FORMAT_UNDERLINE_NONE
                Editor1.Selection.Font.Underline = cprNone
        Case ID_FORMAT_UNDERLINE_THIN
                Editor1.Selection.Font.Underline = cprHair
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_THICK
                Editor1.Selection.Font.Underline = cprThick
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_WAVE
                Editor1.Selection.Font.Underline = cprWave
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_DOT
                Editor1.Selection.Font.Underline = cprDotted
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_DASH
                Editor1.Selection.Font.Underline = cprDash
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_DASHDOT
                Editor1.Selection.Font.Underline = cprDashDot
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_DASHDOT2
                Editor1.Selection.Font.Underline = cprDashDotDot
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
    End Select
    
    Me.Editor1.ForceEdit = blnForce
    Editor1.Tag = ""
    Call ClearNoUseUndoList
    Call RecountPage
End Sub
Private Sub ExecuteLineSpace(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal blnForce As Boolean)
    Dim BarSpace As CommandBarPopup, objControl As CommandBarControl, BarLineSpace As CommandBarPopup, BarLineSpace1 As CommandBarPopup
    On Error GoTo errHand
    Call AddUndoPoint  '手动缓存
    Editor1.ForceEdit = True
    Editor1.Tag = "cbrThis_Execute"
    Set BarSpace = Me.cbrThis.FindControl(, ID_Main_FORMAT)
    If Not BarSpace Is Nothing Then Set BarLineSpace1 = BarSpace.CommandBar.FindControl(, ID_FORMAT_SPACE).CommandBar.FindControl(, ID_FORMAT_LINESPACE)
    Set BarLineSpace = CommBar(ID_BAR_FORMAT).FindControl(xtpControlSplitButtonPopup, ID_FORMAT_LINESPACE)
    If BarLineSpace Is Nothing Then Exit Sub
    Select Case Control.ID
        Case ID_FORMAT_LINESPACE
            If BarLineSpace.Checked Then
                Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 0
            Else
                Select Case True
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE1).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1#
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE2).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1.3
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE3).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1.5
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE4).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 2#
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE5).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 2.5
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE6).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 3#
                End Select
            End If
        Case ID_FORMAT_LINESPACE1
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1
        Case ID_FORMAT_LINESPACE2
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1.3
        Case ID_FORMAT_LINESPACE3
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1.5
        Case ID_FORMAT_LINESPACE4
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 2
        Case ID_FORMAT_LINESPACE5
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 2.5
        Case ID_FORMAT_LINESPACE6
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 3
        Case ID_FORMAT_LINESPACE7
            Editor1.ShowParaDlg True
    End Select
    If Control.ID <> ID_FORMAT_LINESPACE Then
        Call CheckMenu(Control.ID, BarLineSpace)  '工具条上选中
        Call CheckMenu(Control.ID, BarLineSpace1) '菜单选中
    End If
    Me.Editor1.ForceEdit = blnForce
    Editor1.Tag = ""
    Call ClearNoUseUndoList
    Call RecountPage
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function CheckMenu(ByVal ID As Long, ByVal obj As CommandBarPopup)
    Dim objControl As CommandBarControl
    For Each objControl In obj.CommandBar.Controls
        If objControl.ID = ID Then
            objControl.Checked = True
        Else
            objControl.Checked = False
        End If
    Next
End Function

Private Sub SpicalCopy(ByVal blnEnabled As Boolean, ByVal blnVisible As Boolean)
    '-----------------------
    '专用复制
    On Error GoTo errHand
    Dim frm As New frmContentCopy
    Dim blnCan As Boolean
    blnCan = frm.ShowMe(Me, Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.主页ID, Me.Document.EPRPatiRecInfo.病人来源)
    If blnCan Then
        If blnEnabled And blnVisible Then '快捷键执行时需要判断
            Call ExecPaste(Me.Editor1)   '粘贴内容（修正关键字）
            Call RecountPage
        End If
    End If
    Clipboard.Clear
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function RelateFeedback(ByVal isRelated As Boolean) As Boolean
'功能：传染病报告卡，关联阳性结果反馈单，或者取消关联
'参数：isRelated  true-关联；false-取消关联
    Dim objDisease As Object
On Error GoTo errHand
    If Me.Document.EPRPatiRecInfo.病历种类 <> cpr诊断文书 Or mbln返修处理 Then Exit Function
    Set objDisease = CreateObject("zl9Disease.cDockDisease")
    If objDisease Is Nothing Then Exit Function
    Call objDisease.InitDockDisease(glngSys, gcnOracle)
    Call objDisease.RelateFeedback(Me, Me.Document.EPRPatiRecInfo.ID, Me.Document.EPRPatiRecInfo.病人ID, Me.Document.EPRPatiRecInfo.主页ID, Me.Document.EPRPatiRecInfo.病人来源, isRelated)
    Set objDisease = Nothing
    RelateFeedback = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
