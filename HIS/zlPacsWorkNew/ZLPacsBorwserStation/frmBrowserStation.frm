VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmBrowserStation 
   Caption         =   "影像浏览工作站"
   ClientHeight    =   7305
   ClientLeft      =   10185
   ClientTop       =   345
   ClientWidth     =   11325
   Icon            =   "frmBrowserStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11325
   Begin ZLPacsBrowserStation.ucReadCard ucLocate 
      Height          =   330
      Left            =   6720
      TabIndex        =   16
      Top             =   1320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      Picture         =   "frmBrowserStation.frx":0E42
   End
   Begin VB.PictureBox PicWindow 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   1725
      ScaleHeight     =   4215
      ScaleWidth      =   9510
      TabIndex        =   1
      Top             =   2670
      Width           =   9510
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   9465
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   9465
         Begin VB.Frame fraRegist 
            Height          =   810
            Left            =   0
            TabIndex        =   7
            Top             =   -75
            Width           =   8700
            Begin VB.CommandButton cmdReportView 
               Appearance      =   0  'Flat
               Height          =   615
               Left            =   8040
               Picture         =   "frmBrowserStation.frx":1194
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "报告查看"
               Top             =   120
               Width           =   615
            End
            Begin VB.ComboBox cboTimes 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   260
               Width           =   6315
            End
            Begin VB.Label lblRegist 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "检查记录(&G)："
               Height          =   180
               Left            =   105
               TabIndex        =   9
               Top             =   320
               Width           =   1170
            End
         End
         Begin VB.Frame fraInfo 
            Height          =   700
            Left            =   0
            TabIndex        =   4
            Top             =   600
            Width           =   7410
            Begin VB.Label lblCash 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "收"
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   21.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   540
               Left            =   6840
               TabIndex        =   10
               Top             =   120
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label lbl检查信息 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "检查信息"
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   90
               TabIndex        =   6
               Top             =   400
               Width           =   720
            End
            Begin VB.Label lbl个人信息 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "个人信息"
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   90
               TabIndex        =   5
               Top             =   150
               Width           =   720
            End
         End
      End
      Begin XtremeSuiteControls.TabControl TabWindow 
         Height          =   2415
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   4260
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   30
      ScaleHeight     =   4275
      ScaleWidth      =   4500
      TabIndex        =   11
      Top             =   495
      Width           =   4495
      Begin VB.TextBox txtAppend 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BorderStyle     =   0  'None
         Height          =   2100
         Left            =   630
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1605
         Width           =   2010
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2685
         Left            =   450
         TabIndex        =   13
         Top             =   435
         Width           =   3360
         _cx             =   5927
         _cy             =   4736
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Left            =   2730
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(*)"
            Top             =   270
            Visible         =   0   'False
            Width           =   270
         End
      End
      Begin ZLPacsBrowserStation.ucReadCard ucFilter 
         Height          =   330
         Left            =   360
         TabIndex        =   17
         Top             =   120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         Picture         =   "frmBrowserStation.frx":1DD6
      End
      Begin XtremeCommandBars.CommandBars cbrdock 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Left            =   7200
      Top             =   600
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Bindings        =   "frmBrowserStation.frx":2128
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBrowserStation.frx":213C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7805
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList Imglist 
      Left            =   6570
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":29D0
            Key             =   "紧急"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":2F6A
            Key             =   "住院"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":3844
            Key             =   "阳性"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":399E
            Key             =   "影像"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":4118
            Key             =   "已缴"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":44B2
            Key             =   "绿色通道"
            Object.Tag             =   "6"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5940
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":460C
            Key             =   "复选留空"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":4BA6
            Key             =   "复选选中"
            Object.Tag             =   "90001"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmBrowserStation.frx":5140
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBrowserStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mintCur业务类型 As Integer = 1 '当前系统操作的业务类型

Private Const ConstrCol = "紧急;300|来源;400|阳性;300|质量;300|姓名;1200|检查号;1400|检查过程;800|性别;450|年龄;450" & _
                        "|标识号;1400|医嘱内容;2400|部位方法;1400|执行间;600|报到时间;1800|申请时间;1800|开嘱医生;800" & _
                        "|身高;450|体重;450|婴儿;450|登记人;800|报到人;800|完成人;800|打印胶片;800|报告操作;800" & _
                        "|绿色通道;0|报告打印;800|报告人;800|复核人;800|检查技师;800|采图时间;1800|随访描述;2400" & _
                        "|影像类别;0|病人ID;0|主页ID;0|挂号单;0|病人科室ID;0|医嘱ID;1200|发送号;0|检查UID;0" & _
                        "|检查状态;0|NO;0|记录性质;0|转出;0|床号;0|当前病区ID;0|报告发放;800|诊断分类;800" & _
                        "|执行科室ID;0|关联ID;0|病人科室;800|就诊卡号;800|单据号;800|身份证号;800"
Private mstrCol As String   '列表顺序窗体加载时读取注册表，若无值用ConstrCol为默认值

'ID_查找方式+100之后保留7个是作为查找方式选择的
'ID_影像类别之后保留40个号码作为影像类别，从4021-4060
Private Enum FilterID
    ID_门诊 = 4001: ID_住院 = 4002: ID_体检 = 4003: ID_外诊 = 4004
    ID_费用 = 4005: ID_已缴 = 4006: ID_未缴 = 4007: ID_登记 = 4008
    ID_报到 = 4009: ID_检查 = 4010: ID_报告 = 4011: ID_审核 = 4012: ID_完成 = 4013
    ID_查找方式 = 4014: ID_查找值 = 4015: ID_开始查找 = 4016: ID_本次住院 = 4017
    ID_影像类别 = 4020
End Enum

Private mblncmd门诊 As Boolean, mblncmd住院 As Boolean, mblncmd体检 As Boolean, mblncmd外诊 As Boolean, mblncmd已缴 As Boolean, mblncmd未缴 As Boolean
Private mblncmd登记 As Boolean, mblncmd报到 As Boolean, mblncmd检查 As Boolean, mblncmd报告 As Boolean, mblncmd审核 As Boolean, mblncmd完成 As Boolean
Private mblncmd本次 As Boolean
Private mintcmd影像类别 As Integer      '0表示没有选择影像类别，其他数字表示选择的影像类别的数量
Private mblncmd影像类别() As Boolean    '保存当前选择的影像类别是否被选择



Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private Enum IDKinds
    C0姓名或就诊卡 = 0
    C1医保号 = 1
    C2身份证号 = 2
    C3IC卡号 = 3
End Enum

'子窗体对像
Private mfrmPACSImg As frmPACSImg       '影像子窗体
Public mobjRichEPR As New zlRichEPR.cRichEPR
Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer           '观片站对象
Attribute mobjPacsCore.VB_VarHelpID = -1

'窗口变量
Private mlngAdviceID As Long
'Private mlngCur科室ID As Long                               '当前科室ID
Private mstr医生所属科室 As String                            '当前可用科室  ID_编码-名称
Private mstr医生所属科室IDs As String
'Private mstr当前医生科室 As String                            '当前科室 编码-名称
Private mblnInitOk As Boolean, mblnvsRefresh As Boolean     '初始化完成,装载表格
Private mstrPrivs As String, mlngModul As Long              '模块号，本模块权限
Private mblnAllDepts As Boolean                             '是否选择全部科室
Private mlngSortCol As Long                                 '病人列表中，当前进行排序的列
Private mintSortOrder As Integer                            '病人列表中，当前进行排序的方式

'流程控制变量
Private mblnShowImgAtReport As Boolean                      '打开报告时打开观片站
Private mBeforeDays As Integer                              '默认查询的天数
Private mlngRefreshInterval As Long                         '病人列表自动刷新间隔
Private mblnRelatingPatient As Boolean                      '是否启用关联病人
Private mblnMoved As Boolean                                '当前时间段内是否被转移过

Private mblnUse3D As Boolean                                '是否启用三维重建功能
Private mstr3DExeDir As String                              '三维重建程序路径
Private mstr3DPara As String                                '三维重建参数
Private mstr3DFunctions As String                           '三维重建功能

'过滤条件变量
Private Type Type_SQLCondition
    开始时间 As Date
    结束时间 As Date
    时间类型 As Integer                                 '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
    单据号 As String
    门诊号 As Double
    住院号 As Double
    就诊卡 As String
    姓名 As String
    性别 As String
    开始年龄 As Long
    结束年龄 As Long
    年龄条件 As String
    检查号 As Double
    身份证  As String
    IC卡 As String
    病人科室 As Long
    标本部位 As String
    诊断医生 As String
    审核医生 As String
    疾病诊断 As String
    报告内容 As String
    结果阳性 As Integer
    影像质量 As String
    检查技师 As String
    检查过程 As String
    影像类别 As String
    检查所见 As String
    诊断意见 As String
    建议 As String
    随访 As String
    病人ID As Long
End Type
Private SQLCondition As Type_SQLCondition

Private mlngHSendNo As Long
Private mstrHStudyUID As String
Private mlngExecuteStep As Long '检查执行过程
Private mblnHMoved As Boolean


Private Sub OpenReportPreview(ByVal lngAdviceID As Long)
    If mobjRichEPR Is Nothing Then Exit Sub
    
    On Error GoTo errHandle
        
        Dim strSQL As String
        Dim lngExecuteStep As Long
        Dim rsReport As ADODB.Recordset
        Dim blnCanPrint As Boolean
        
        strSQL = "select 执行过程,病历ID from 病人医嘱发送 A,病人医嘱报告 R where R.医嘱ID=A.医嘱ID and  A.医嘱ID=[1]"
        Set rsReport = zlDatabase.OpenSQLRecord(strSQL, "查询病历ID", lngAdviceID)
        
        If rsReport.EOF Then
            MsgBoxD Me, "没有找到当前检查对应的病历信息，请检查！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        lngExecuteStep = rsReport!执行过程
        
        'If lngExecuteStep <> 5 And rsReport!执行过程 <> 6 Then
        If rsReport!执行过程 <> 6 Then
            MsgBoxD Me, "报告尚未完成，不能进行查看！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If InStr(mstrPrivs, "PACS报告打印") > 0 Then
            blnCanPrint = True
        Else
            blnCanPrint = False
        End If
        
        Call mobjRichEPR.ViewDocument(Me, rsReport!病历Id, blnCanPrint)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume Next
    End If
End Sub


Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Help_click()
    '功能：调用帮助主题
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.Hwnd)
End Sub


Private Sub Menu_Help_Web_Mail_click()
    zlMailTo Hwnd
End Sub


Private Sub Menu_Manage_观片()
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        Call mfrmPACSImg.zlRefresh(mlngAdviceID, mlngHSendNo, mstrPrivs, mblnHMoved)
    End If
    
    Call mfrmPACSImg.zlMenuClick("影像处理")
End Sub


Private Sub Menu_Manage_对比观片()
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        Call mfrmPACSImg.zlRefresh(mlngAdviceID, mlngHSendNo, mstrPrivs, mblnHMoved)
    End If
    
    Call mfrmPACSImg.zlMenuClick("影像对比")
End Sub


Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer, cbrControl As CommandBarControl
    For i = 2 To cbrMain.Count
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
    Next
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub cboTimes_Click()
    If cboTimes.ListCount <= 1 Then Exit Sub
    If cboTimes.Tag = "" Then Exit Sub '此时cbotime项目未增加完成，属listindex赋值触发
    
    On Error GoTo errHandle
    
    mlngAdviceID = cboTimes.ItemData(cboTimes.ListIndex)
    'If mlngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) Then Call vsList_RowColChange: Exit Sub '当次与当前选中医嘱ID相同时不由本函数控制
    
    If mlngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) Then
      Call FillTxtInfor  '填充右上方病人基本信息
      Call FillTxtAppend '填充左下角医嘱附件
    
      Call RefreshTabWindow '刷新子窗体
        
    Else
      '以下三个过程调用有先后顺序，勿调换
      Call FillTxtInfor(mlngAdviceID)  '填充右上方病人基本信息
      Call FillTxtAppend(mlngAdviceID) '填充左下角医嘱附件
    
      Call RefreshTabWindow(mlngAdviceID) '刷新子窗体
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboTimes_DropDown()
    Call SendMessage(cboTimes.Hwnd, &H160, 500, 0)
End Sub

Private Sub cbrdock_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim strTemp As String
    Dim strCardName As String
    Dim strCardText As String
    Dim lngPatientID As Long
    
    Select Case control.ID
        Case ID_门诊
            mblncmd门诊 = Not control.Checked
        Case ID_住院
            mblncmd住院 = Not control.Checked
        Case ID_外诊
            mblncmd外诊 = Not control.Checked
        Case ID_体检
            mblncmd体检 = Not control.Checked
        Case ID_已缴
            mblncmd已缴 = Not control.Checked
            If mblncmd已缴 Then mblncmd未缴 = False
        Case ID_未缴
            mblncmd未缴 = Not control.Checked
            If mblncmd未缴 Then mblncmd已缴 = False
        Case ID_影像类别 + 1 To ID_影像类别 + 40
            control.Checked = Not control.Checked
            mblncmd影像类别(control.ID - ID_影像类别 - 1) = control.Checked
            If control.Checked = True Then
                mintcmd影像类别 = mintcmd影像类别 + 1
            Else
                mintcmd影像类别 = mintcmd影像类别 - 1
            End If
            Set objControl = cbrdock.FindControl(, ID_影像类别)
            If mintcmd影像类别 = 0 Then
                strTemp = "影像类别"
            Else
                strTemp = ""
                For i = 1 To objControl.CommandBar.Controls.Count
                    If objControl.CommandBar.FindControl(, ID_影像类别 + i).Checked = True Then
                        strTemp = IIf(strTemp = "", objControl.CommandBar.FindControl(, ID_影像类别 + i).Caption, strTemp & "," & objControl.CommandBar.FindControl(, ID_影像类别 + i).Caption)
                    End If
                Next i
            End If
            objControl.Caption = strTemp
        Case ID_登记
            mblncmd登记 = Not control.Checked
        Case ID_报到
            mblncmd报到 = Not control.Checked
        Case ID_检查
            mblncmd检查 = Not control.Checked
        Case ID_报告
            mblncmd报告 = Not control.Checked
        Case ID_审核
            mblncmd审核 = Not control.Checked
        Case ID_完成
            mblncmd完成 = Not control.Checked
        Case ID_本次住院
            control.Checked = Not control.Checked
            mblncmd本次 = Not mblncmd本次
        Case ID_开始查找
            Call ucFilter.GetCardValue(strCardName, strCardText, lngPatientID)
            Call subRefreshFilterCondition(strCardName, strCardText, lngPatientID)
    End Select
    
    cbrdock.RecalcLayout
    Call RefreshList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub subRefreshFilterCondition(ByVal strCardName As String, ByVal strCardText As String, ByVal lngPatientID As Long)
'------------------------------------------------
'功能：用txtFilter控件的内容更新过滤条件
'参数： strFilter --- 过滤条件
'返回：无
'------------------------------------------------

On Error GoTo err
    Dim strFilter As String
    
    strFilter = strCardText
    
    With SQLCondition
        .姓名 = ""
        .就诊卡 = ""
        .门诊号 = 0
        .住院号 = 0
        .单据号 = ""
        .检查号 = 0
        .身份证 = ""
        .IC卡 = ""
        .病人ID = 0
        
        Select Case strCardName
            Case "姓名", "姓  名", "姓   名" '保持与以前方式兼容
                .姓名 = Trim(strFilter)
                
            Case "就诊卡"
                .就诊卡 = Trim(strFilter)
                
            Case "门诊号"   '快捷方式是“*+数字”,VAL提取前，“*”要特殊处理
                If Left(strCardText, 1) = "*" Then
                    strFilter = Mid(strFilter, 2)
                End If
                .门诊号 = Val(strFilter)
                
            Case "住院号"   '快捷方式是“++数字”
                .住院号 = Val(strFilter)
                
            Case "单据号"
                .单据号 = Trim(strFilter)
                
            Case "检查号"
                .检查号 = Val(strFilter)
                
            Case "身份证号", "身份证"
                .身份证 = Trim(strFilter)
                
            Case "IC卡号", "IC卡"
                .IC卡 = Trim(strFilter)
                
            Case Else
                .病人ID = lngPatientID
        End Select
    End With
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrdock_Resize()
Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    Call Me.cbrdock.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    vsList.Top = lngTop: vsList.Left = lngLeft
    vsList.Width = picList.Width
    vsList.Height = picList.Height - lngTop - txtAppend.Height - 100


    txtAppend.Top = vsList.Top + vsList.Height + 100: txtAppend.Left = lngLeft + 100
    txtAppend.Width = picList.Width - 200
End Sub

Private Sub cbrdock_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case ID_门诊
            control.Checked = mblncmd门诊
            control.IconId = IIf(mblncmd门诊, 90001, 90000)
        Case ID_住院
            control.Checked = mblncmd住院
            control.IconId = IIf(mblncmd住院, 90001, 90000)
        Case ID_外诊
            control.Checked = mblncmd外诊
            control.IconId = IIf(mblncmd外诊, 90001, 90000)
        Case ID_体检
            control.Checked = mblncmd体检
            control.IconId = IIf(mblncmd体检, 90001, 90000)
        Case ID_费用
            control.Checked = mblncmd已缴 Xor mblncmd未缴
            control.Caption = IIf(mblncmd已缴 Xor mblncmd未缴, IIf(mblncmd已缴, " 已缴费", " 未缴费"), " 费  用")
        Case ID_已缴
            control.Checked = mblncmd已缴
            control.IconId = IIf(mblncmd已缴, 90001, 90000)
        Case ID_未缴
            control.Checked = mblncmd未缴
            control.IconId = IIf(mblncmd未缴, 90001, 90000)
        Case ID_影像类别
            control.IconId = IIf(mintcmd影像类别 = 0, 90000, 90001)
        Case ID_影像类别 + 1 To ID_影像类别 + 40
            control.Checked = mblncmd影像类别(control.ID - ID_影像类别 - 1)
            control.IconId = IIf(control.Checked, 90001, 90000)
        Case ID_登记
            control.Checked = mblncmd登记
            control.IconId = IIf(mblncmd登记, 90001, 90000)
        Case ID_报到
            control.Checked = mblncmd报到
            control.IconId = IIf(mblncmd报到, 90001, 90000)
        Case ID_检查
            control.Checked = mblncmd检查
            control.IconId = IIf(mblncmd检查, 90001, 90000)
        Case ID_报告
            control.Checked = mblncmd报告
            control.IconId = IIf(mblncmd报告, 90001, 90000)
        Case ID_审核
            control.Checked = mblncmd审核
            control.IconId = IIf(mblncmd审核, 90001, 90000)
        Case ID_完成
            control.Checked = mblncmd完成
            control.IconId = IIf(mblncmd完成, 90001, 90000)
        Case ID_本次住院
            control.IconId = IIf(control.Checked, 90001, 90000)
    End Select
End Sub
Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = stbThis.Height
End Sub



Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    
    If control.ID <> 0 Then
        If cbrMain.FindControl(, control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    cbrMain.RecalcLayout
    Select Case control.ID
    
'--------------------------文件------------------
        
        Case conMenu_Manage_Change_In   '隐藏列表
            If dkpMain.Panes(1).Hidden = False Then
                dkpMain.Panes(1).Hide
            Else
                dkpMain.ShowPane (1)
            End If

        Case conMenu_File_Exit '退出
            Unload Me
            
'---------------------------检查-----------------
        Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '打印诊疗单据
            Call FuncBillPrint(control)
'
        Case conMenu_Img_Look                         '观片
            Call Menu_Manage_观片
        
        Case conMenu_Img_Contrast                     '对比观片
            Call Menu_Manage_对比观片
        
        Case conMenu_Img_3D_MMPR                    '三维重建，MMPR
            Call sub三维重建("MMPR")
        Case conMenu_Img_3D_MPR                     '三维重建，MPR
            Call sub三维重建("MPR")
        Case conMenu_Img_3D_PF                     '三维重建,灌注成像
            Call sub三维重建("PF")
        Case conMenu_Img_3D_SA                     '三维重建，表面重建
            Call sub三维重建("SA")
        Case conMenu_Img_3D_VA                     '三维重建，容积重建
            Call sub三维重建("VA")
        Case conMenu_Img_3D_VE                     '三维重建，虚拟内窥镜
            Call sub三维重建("VE")

        Case conMenu_File_Preview ', conMenu_File_Print       '报告预览和打印
            Dim i As Integer
            '没报告不能打印和预览
            If vsList.TextMatrix(vsList.Row, GetCN("报告人")) = "" Then
                MsgBoxD Me, "当前病人没有检查报告，不能操作，请检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            Call OpenReportPreview(mlngAdviceID)

'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button                        '工具栏
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text                          '按钮文字
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size                          '大图标
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar                             '状态栏
            Call Menu_View_StatusBar_click(control)
        Case conMenu_View_Filter                                '过滤
            Call Menu_View_Filter_click
        Case conMenu_View_Refresh                               '刷新
            Call RefreshList
        Case conMenu_Help_Help                                  '帮助
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum                             '网上中联
'            Case zlWebForum(Me.Hwnd)
        Case conMenu_Help_Web_Home                              '网上中联
            Call zlHomePage(Me.Hwnd)
        Case conMenu_Help_Web_Mail                              '电邮中联
            Call zlMailTo(Me.Hwnd)
        Case conMenu_Help_About                                 '关于
            Call Menu_Help_About_click
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstr医生所属科室, "|")) + 1 '更改当前科室
            Call Menu_Dept_Select(control)
    End Select
End Sub


Private Sub Menu_Dept_Select(ByVal control As XtremeCommandBars.ICommandBarControl)
    '更换科室，根据新的条件，重新过滤病人
    '如果选择的是全部科室，则 mlngCur科室ID 不改变
    '如果选择的是某个具体科室，则改变 mlngCur科室ID
    If glngDeptId <> control.DescriptionText Or (control.DescriptionText <> 0 And mblnAllDepts = True) Then
        '选择了具体科室，才改变当前科室的设置
        If control.DescriptionText = 0 Then
            mblnAllDepts = True
        Else
            mblnAllDepts = False
            glngDeptId = control.DescriptionText
            gstrDeptName = Split(control.Caption, "(")(0)
            
        End If
        
        Call cbrMain.RecalcLayout
        Call RefreshList
    End If
End Sub


Private Sub Menu_View_Filter_click()
    On Error GoTo errHandle
    
    With frmPACSFilter
        .mlngModul = mlngModul
        .mBeforeDays = mBeforeDays
'        .mDept = mlngCur科室ID '当前科室
        .Show 1, Me
        If Not .mblnOK Then Exit Sub '没有返回条件
        
        '当使用时间条件时，清空固定条件
        ucFilter.CardText = ""
        SQLCondition.姓名 = ""
        SQLCondition.就诊卡 = ""
        SQLCondition.门诊号 = 0
        SQLCondition.住院号 = 0
        SQLCondition.单据号 = ""
        SQLCondition.检查号 = 0
        SQLCondition.身份证 = ""
        SQLCondition.IC卡 = ""
        
        SQLCondition.开始时间 = Format(.dtpBegin.value, "yyyy-MM-dd HH:mm:00")
        SQLCondition.结束时间 = Format(.dtpEnd.value, "yyyy-MM-dd HH:mm:59")
'        If Format(.dtpEnd.value, "yyyy-MM-dd HH:mm") = Format(.dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
'            SQLCondition.结束时间 = CDate(0) '表示取当前时间
'        Else
'            SQLCondition.结束时间 = Format(.dtpEnd.value, "yyyy-MM-dd HH:mm:59")
'        End If
        
        mblnMoved = MovedByDate(SQLCondition.开始时间)
        
        If .optFindType(1).value = True Then '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
            SQLCondition.时间类型 = 1
        ElseIf .optFindType(2).value = True Then
            SQLCondition.时间类型 = 2
        Else
            SQLCondition.时间类型 = 3
        End If
        
        If .cboPart.ListIndex <> 0 Then '检查标本部位
            SQLCondition.标本部位 = NeedName(.cboPart.Text)
        Else
            SQLCondition.标本部位 = ""
        End If
        
        '病人性别
        If NeedName(.cboSex.Text) = "全部" Then
            SQLCondition.性别 = ""
        Else
            SQLCondition.性别 = NeedName(.cboSex.Text)
        End If
        
        '病人年龄
        Select Case NeedName(.cboAgeType.Text)
            Case "岁"
                SQLCondition.开始年龄 = Val(.txtBeginAge.Text) * 365
                SQLCondition.结束年龄 = Val(.txtEndAge.Text) * 365
            Case "月"
                SQLCondition.开始年龄 = Val(.txtBeginAge.Text) * 30
                SQLCondition.结束年龄 = Val(.txtEndAge.Text) * 30
            Case "周"
                SQLCondition.开始年龄 = Val(.txtBeginAge.Text) * 7
                SQLCondition.结束年龄 = Val(.txtEndAge.Text) * 7
            Case "天"
                SQLCondition.开始年龄 = Val(.txtBeginAge.Text) * 1
                SQLCondition.结束年龄 = Val(.txtEndAge.Text) * 1
        End Select
        
        If Trim(.txtBeginAge.Text) = "" Then SQLCondition.开始年龄 = -1
        If Trim(.txtEndAge.Text) = "" Then SQLCondition.结束年龄 = -1
        
        SQLCondition.年龄条件 = Trim(.cboAgeWhere.Text)
        
        If .cboDept.ListIndex <> 0 Then '病人科室
            SQLCondition.病人科室 = .cboDept.ItemData(.cboDept.ListIndex)
        Else
            SQLCondition.病人科室 = 0
        End If

        If .cbodiagdoc.ListIndex <> 0 Then '诊断医生
            SQLCondition.诊断医生 = NeedName(.cbodiagdoc.Text)
        Else
            SQLCondition.诊断医生 = ""
        End If
        
        If .cboAuditing.ListIndex <> 0 Then '审核医生
            SQLCondition.审核医生 = NeedName(.cboAuditing.Text)
        Else
            SQLCondition.审核医生 = ""
        End If
        
        
'        If .cboCheckStep.ListIndex <> 0 Then '检查过程
'            SQLCondition.检查过程 = .cboCheckStep.Text
'        Else
'            SQLCondition.检查过程 = ""
'        End If
        
        
        If .cboModality.ListIndex <> 0 Then '影像类别
            SQLCondition.影像类别 = Split(.cboModality.Text, "--")(1)
        Else
            SQLCondition.影像类别 = ""
        End If
        
        
        If Trim(.Txt影像诊断) <> "" Then '影像诊断
            SQLCondition.疾病诊断 = Trim(.Txt影像诊断)
        Else
            SQLCondition.疾病诊断 = ""
        End If
        
        If Trim(.txt报告内容) <> "" Then '报告内容
            SQLCondition.报告内容 = Trim(.txt报告内容)
        Else
            SQLCondition.报告内容 = ""
        End If
        
        If NeedName(.cboYinYangXing.Text) = "阳性" Then
            SQLCondition.结果阳性 = 1
        ElseIf NeedName(.cboYinYangXing.Text) = "阴性" Then
            SQLCondition.结果阳性 = 0
        Else
            SQLCondition.结果阳性 = -1
        End If
        
        If .cbo质量.ListIndex = 0 Then
            SQLCondition.影像质量 = ""
        Else
            SQLCondition.影像质量 = NeedName(.cbo质量.Text)
        End If
        
        If .cbo检查技师.ListIndex = 0 Then
            SQLCondition.检查技师 = ""
        Else
            SQLCondition.检查技师 = NeedName(.cbo检查技师.Text)
        End If
        
        
        If Trim(.txtPacsRpt(0)) <> "" Then 'PACS报告检索
            SQLCondition.检查所见 = Trim(.txtPacsRpt(0))
        Else
            SQLCondition.检查所见 = ""
        End If
        
        If Trim(.txtPacsRpt(1)) <> "" Then
            SQLCondition.诊断意见 = Trim(.txtPacsRpt(1))
        Else
            SQLCondition.诊断意见 = ""
        End If
        
        If Trim(.txtPacsRpt(2)) <> "" Then
            SQLCondition.建议 = Trim(.txtPacsRpt(2))
        Else
            SQLCondition.建议 = ""
        End If
        
        If Trim(.txt随访.Text) <> "" Then
            SQLCondition.随访 = Trim(.txt随访.Text)
        Else
            SQLCondition.随访 = ""
        End If
        
        Call RefreshList '调用刷新
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
        Case conMenu_View_Filter * 10#
            With CommandBar.Controls
                If .Count = 0 Then
                    '先添加全部科室
                    Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100#, "全部科室")
                    objControl.BeginGroup = True
                    objControl.Category = "Main"
                    objControl.DescriptionText = 0
                    If mblnAllDepts = True Then objControl.Checked = True
                    
                    '再添加每一个具体科室
                    For i = 0 To UBound(Split(mstr医生所属科室, "|"))  'mstr医生所属科室=id_编码-名称|id_编码-名称
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i + 1, Split(Split(mstr医生所属科室, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstr医生所属科室, "|")(i), "_")(0)
                        If mblnAllDepts = False And glngDeptId = objControl.DescriptionText Then objControl.Checked = True
                    Next
                End If
            End With
    End Select
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnNoRecord As Boolean, intState As Integer, blnCancel As Boolean
    If Not mblnInitOk Then Exit Sub

    blnNoRecord = Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) = 0
    control.Style = xtpButtonIconAndCaption
    If Not blnNoRecord Then
        intState = Val(vsList.TextMatrix(vsList.Row, GetCN("检查状态")))
        blnCancel = vsList.TextMatrix(vsList.Row, GetCN("检查过程")) = "已拒绝"
    End If

    Select Case control.ID
        Case conMenu_Manage_LocateValue
            control.Enabled = Not blnNoRecord
        Case conMenu_View_Filter * 10#
            control.Caption = "当前科室:" & IIf(mblnAllDepts = True, "全部科室", gstrDeptName)
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstr医生所属科室, "|")) + 1
            If mblnAllDepts = True Then
                control.Checked = (control.DescriptionText = 0)
            Else
                control.Checked = (control.DescriptionText = glngDeptId)
            End If
        Case conMenu_View_ToolBar_Button '工具栏
            If cbrMain.Count >= 2 Then
                control.Checked = Me.cbrMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbrMain.Count >= 2 Then
                control.Checked = Not (Me.cbrMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            control.Checked = Me.cbrMain.Options.LargeIcons
        Case conMenu_View_StatusBar '状态栏
            control.Checked = Me.stbThis.Visible
        Case conMenu_View_Filter   '过滤

        Case conMenu_View_Refresh  '刷新

        Case conMenu_Manage_RequestPrint
            control.Enabled = control.CommandBar.Controls.Count > 0 And Not blnNoRecord

        Case conMenu_Img_Contrast, conMenu_Img_Look     '影像对比,影像观片
            If blnNoRecord Then control.Enabled = False: Exit Sub

            control.Enabled = mstrHStudyUID <> ""
                        
            'If control.Parent.Type <> xtpControlPopup Then control.Visible = control.Enabled
        Case conMenu_Img_3D     '三维重建
            If InStr(mstrPrivs, "三维重建操作") <> 0 And mblnUse3D = True Then
                control.Visible = True
            Else
                control.Visible = False
            End If
            If control.Visible = True Then
                If blnNoRecord Then control.Enabled = False: Exit Sub
                If control.Parent.Type <> xtpControlPopup Then
                    control.Visible = vsList.TextMatrix(vsList.Row, GetCN("检查UID")) <> ""
                    control.Enabled = control.Visible
                Else
                    control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("检查UID")) <> ""
                End If
            End If

'        Case conMenu_File_PrintSet     '打印设置(&S)
        Case conMenu_File_Preview, conMenu_File_Print '报告预览(&V) 报告打印(&P)
            control.Enabled = Not blnNoRecord And (mlngExecuteStep = 5 Or mlngExecuteStep = 6)
'        Case conMenu_File_Excel         '清单打印(&L)
'            control.Enabled = Not blnNoRecord
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99 '报表
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_Help_Help, conMenu_Help_About  '帮助
        Case conMenu_Help_Web, conMenu_Help_Web_Forum, conMenu_Help_Web_Home, conMenu_Help_Web_Mail '帮助WEB
        Case conMenu_File_Exit      '退出
        Case conMenu_View_ToolBar   '工具栏
        Case conMenu_Cap_DevSet     '影像设备设置
        Case conMenu_Manage_Change_In   '隐藏列表
    End Select
End Sub

Private Sub InitMvar(intType As Integer)
'功能:初始化模块级变量,窗体加载、科室改变时调用
'参数：intType---0从菜单或者FormLoad触发科室改变，刷新病人过滤开始时间；intType---1从病人列表触发科室改变，不用再刷新过滤开始时间

    On Error GoTo err
    
    '读取跟科室相关的流程管理参数
    mBeforeDays = 1 'Val(GetDeptPara(mlngCur科室ID, "默认过滤天数", 2)) '                   '默认过滤天数
    If mBeforeDays > 15 Or mBeforeDays <= 0 Then
        mBeforeDays = 2
    End If


    If intType = 0 Then    '从菜单或者FormLoad触发科室改变，刷新病人过滤开始时间
        SQLCondition.开始时间 = CDate(Format(zlDatabase.Currentdate - mBeforeDays, "yyyy-mm-dd 00:00"))
        mblnMoved = MovedByDate(SQLCondition.开始时间)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Private Sub cmdInfo_Click()
    On Error GoTo errHandle
    frmDegreeCard.ShowMe Val(vsList.TextMatrix(vsList.Row, GetCN("病人ID"))), Val(vsList.TextMatrix(vsList.Row, GetCN("主页ID")))
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdReportView_Click()
            '没报告不能打印和预览
            If vsList.TextMatrix(vsList.Row, GetCN("报告人")) = "" Then
                MsgBoxD Me, "当前病人没有检查报告，不能操作，请检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            Call OpenReportPreview(mlngAdviceID)
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picList.Hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = PicWindow.Hwnd
    End If
End Sub
Private Sub Form_Load()
    mstrPrivs = gstrPrivs           '权限
    mlngModul = glngModul           '模块号
'    mlngCur科室ID = 0
    mstr医生所属科室 = ""
'    mblnAllDepts = False            '默认不选择全部科室
    mlngSortCol = 0
    mintSortOrder = 0
    
    mblnInitOk = False  '初始数据,初始化完成之前不进行数据的提取
    mblnvsRefresh = False
    
    ucFilter.CardNames = "姓名;就诊卡;门诊号;住院号;单据号;检查号;身份证号;健康号;IC卡号;"
    Call ucFilter.InitCardType(glngSys, mlngModul, UserInfo.姓名, gcnOracle)
    
    ucLocate.CardNames = "姓名;就诊卡;门诊号;住院号;单据号;检查号;身份证号;健康号;IC卡号;"
    Call ucLocate.InitCardType(glngSys, mlngModul, UserInfo.姓名, gcnOracle)
    
    Call InitLocalPars '本地注册表参数
    If Not InitDepts Then Unload Me: Exit Sub '初始化医技科室
    
    ReDim gConnectedShardDir(0) As String   '初始化共享目录连接串
    
    Call InitMvar(0) '初始化模块级变量
    '初始子窗体
    Set mfrmPACSImg = New frmPACSImg
    
    Call mobjRichEPR.InitRichEPR(gcnOracle, Me, glngSys, False)
    
    Set mobjPacsCore = New zl9PacsCore.clsViewer

    Call InitFilterCmd
    Call InitCommandBars
    Call InitSubForm
    Call InitFaceScheme
    Call InitList

    Set mfrmPACSImg.pobjPacsCore = mobjPacsCore
    
    mblnInitOk = True '初始化完成
    
    Call RestoreWinState(Me, App.ProductName)
    '不能被restorewinstate冲掉，所以写在其后
    Call RefreshList
    
    ClearCacheFolder App.Path & "\TmpImage\"    '若临时目录满了，则清空该目录
    Me.stbThis.Panels(3).Text = "报告医生：" & UserInfo.姓名
End Sub

Private Function InitDepts() As Boolean
'功能：初始化医生所属科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str科室IDs As String
    
    On Error GoTo errH
    
    strSQL = "select distinct A.ID, A.编码,A.名称 from 部门表 A, 部门人员 B Where a.ID = b.部门ID And b.人员id = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngUserId)
    
    If rsTmp.EOF Then
        MsgBoxD Me, "没有发现所属部门信息,请先到部门管理中设置。", vbInformation, gstrSysName
        Exit Function
    Else
        Do Until rsTmp.EOF
            mstr医生所属科室 = mstr医生所属科室 & "|" & rsTmp!ID & "_" & rsTmp!编码 & "-" & rsTmp!名称
            mstr医生所属科室IDs = mstr医生所属科室IDs & "," & rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        
        mstr医生所属科室 = Mid(mstr医生所属科室, 2)
        mstr医生所属科室IDs = Mid(mstr医生所属科室IDs, 2)
        

        InitDepts = True
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    Dim i As Integer
    
    On Error Resume Next
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "门诊病人", IIf(mblncmd门诊, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "住院病人", IIf(mblncmd住院, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "外诊病人", IIf(mblncmd外诊, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "体检病人", IIf(mblncmd体检, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用已缴", IIf(mblncmd已缴, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用未缴", IIf(mblncmd未缴, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "登记病人", IIf(mblncmd登记, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报到病人", IIf(mblncmd报到, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "检查病人", IIf(mblncmd检查, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报告病人", IIf(mblncmd报告, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "审核病人", IIf(mblncmd审核, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "完成病人", IIf(mblncmd完成, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "过滤方式", ucFilter.CurCardName
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "定位方式", ucLocate.CurCardName
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "本次住院", IIf(mblncmd本次, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序列", mlngSortCol
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序方向", mintSortOrder
    
    If UBound(mblncmd影像类别) >= 0 Then
        strTemp = mblncmd影像类别(0)
    End If
    For i = 1 To UBound(mblncmd影像类别)
        strTemp = strTemp & "," & mblncmd影像类别(i)
    Next i
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "影像类别过滤", strTemp
    
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveWinState(Me, App.ProductName)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, mstrCol)

    
    Unload mfrmPACSImg


    If Not mobjPacsCore Is Nothing Then mobjPacsCore.Closefrom
    
    Set mobjIDCard = Nothing
    Set mobjPacsCore = Nothing
    Set mobjRichEPR = Nothing
End Sub

Private Function GetCN(ByVal Col As String) As Integer
Dim arrCol As Variant, i As Integer
    If mstrCol = "" Then mstrCol = ConstrCol
    arrCol = Split(mstrCol, "|")
    For i = 0 To UBound(arrCol)
        'If InStr(arrCol(i), Col) > 0 Then GetCN = i: Exit Function
        If Split(arrCol(i), ";")(0) = Col Then GetCN = i: Exit Function
    Next
    GetCN = 0
End Function

Private Function GetCW(ByVal Col As String) As Long
    Dim arrCol As Variant, i As Integer
    arrCol = Split(mstrCol, "|")
    For i = 0 To UBound(arrCol)
        'If InStr(arrCol(i), Col) > 0 Then GetCW = Split(arrCol(i), ";")(1): Exit Function
        If Split(arrCol(i), ";")(0) = Col Then GetCW = Split(arrCol(i), ";")(1): Exit Function
    Next
    GetCW = 0
End Function

Private Sub InitLocalPars()
'初始化临时本地参数，以个人设置，注册表参数为主,窗体加载，本地设置等调用
    Dim strTemp As String
    Dim strTempArry() As String
    Dim i As Integer
    
    On Error GoTo err
    mblncmd门诊 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "门诊病人", 1))
    mblncmd住院 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "住院病人", 1))
    mblncmd外诊 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "外诊病人", 1))
    mblncmd体检 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "体检病人", 1))
    mblncmd已缴 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用已缴", 0))
    mblncmd未缴 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用未缴", 0))
    mblncmd登记 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "登记病人", 1))
    mblncmd报到 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报到病人", 1))
    mblncmd检查 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "检查病人", 1))
    mblncmd报告 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报告病人", 1))
    mblncmd审核 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "审核病人", 1))
    mblncmd完成 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "完成病人", 1))

    ucFilter.CurCardName = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "过滤方式", "检查号")
    ucLocate.CurCardName = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "定位方式", "检查号")
    mblncmd本次 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "本次住院", "0"))
    mlngSortCol = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序列", 0))
    mintSortOrder = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序方向", 0))
    
    strTemp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "影像类别过滤", "")
    ReDim strTempArry(0)
    ReDim mblncmd影像类别(0)
    On Error Resume Next
    strTempArry = Split(strTemp, ",")
    If UBound(strTempArry) >= 0 Then ReDim mblncmd影像类别(UBound(strTempArry))
    For i = 0 To UBound(strTempArry)
        mblncmd影像类别(i) = IIf(UCase(strTempArry(i)) = "TRUE", True, False)
    Next i
    
    On Error GoTo err
    
    '读取三维重建参数
    mblnUse3D = Val(zlDatabase.GetPara("启用三维重建", glngSys, mlngModul, 0))
    mstr3DExeDir = zlDatabase.GetPara("3D程序路径", glngSys, mlngModul, "")
    mstr3DPara = zlDatabase.GetPara("3D参数", glngSys, mlngModul, "")
    mstr3DFunctions = zlDatabase.GetPara("3D功能", glngSys, mlngModul, "")

    With SQLCondition '------------------------ '过滤条件初始
        '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
        .时间类型 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "过滤时间类型", 1))
        .单据号 = ""
        .门诊号 = 0
        .住院号 = 0
        .就诊卡 = ""
        .姓名 = ""
        .性别 = ""
        .开始年龄 = -1
        .结束年龄 = -1
        .年龄条件 = "="
        .检查号 = 0
        .身份证 = ""
        .IC卡 = ""
        .病人科室 = 0
        .标本部位 = ""
        .诊断医生 = ""
        .审核医生 = ""
        .疾病诊断 = ""
        .报告内容 = ""
        .结果阳性 = -1
        .影像质量 = ""
        .检查技师 = ""
        .检查过程 = ""
        .影像类别 = ""
        .检查所见 = ""
        .诊断意见 = ""
        .建议 = ""
        .随访 = ""
    End With
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitFaceScheme()
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 240, 250, DockLeftOf, Nothing)
    Pane1.Title = "检查列表"
    Pane1.Handle = picList.Hwnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set Pane2 = dkpMain.CreatePane(2, 700, 250, DockRightOf, Nothing)
    Pane2.Title = "子窗体"
    Pane2.Handle = PicWindow.Hwnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
End Sub

Private Sub InitFilterCmd()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    Dim objPopbar As CommandBarPopup, objCusControl As CommandBarControlCustom
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim i As Integer

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrdock.VisualTheme = xtpThemeOfficeXP
    With Me.cbrdock.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    cbrdock.AddImageList img16 '以VB.ImageList的Tag与ID进行关联
    cbrdock.EnableCustomization False
    cbrdock.ActiveMenuBar.Visible = False
    
    Set objBar = cbrdock.Add("来源", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_门诊, "门诊")
            objControl.ToolTipText = "显示门诊病人"
        Set objControl = .Add(xtpControlButton, ID_住院, "住院")
            objControl.ToolTipText = "显示住院病人"
        Set objControl = .Add(xtpControlButton, ID_外诊, "外诊")
            objControl.ToolTipText = "显示外诊病人"
        Set objControl = .Add(xtpControlButton, ID_体检, "体检")
            objControl.ToolTipText = "显示体检病人"
        Set objControl = .Add(xtpControlButtonPopup, ID_费用, " 费  用")
            objControl.ToolTipText = "显示费用已缴/未缴病人"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_未缴, "未缴")
            cbrPopControl.ToolTipText = "显示费用未缴病人"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_已缴, "已缴")
            cbrPopControl.ToolTipText = "显示费用已缴病人"
        
        '添加所有影像类别
        Set objControl = .Add(xtpControlButtonPopup, ID_影像类别, "影像类别")
        objControl.ToolTipText = "显示影像类别"
        strSQL = "select 编码,名称 from 影像检查类别"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "影像检查类别")
        i = 1
        mintcmd影像类别 = 0
        strTemp = ""
        ReDim Preserve mblncmd影像类别(rsTemp.RecordCount - 1)
        While rsTemp.EOF = False
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_影像类别 + i, rsTemp("名称"))
            cbrPopControl.DescriptionText = rsTemp("编码")
            cbrPopControl.Style = xtpButtonIconAndCaption
            cbrPopControl.Checked = mblncmd影像类别(i - 1)
            cbrPopControl.CloseSubMenuOnClick = False
            If mblncmd影像类别(i - 1) = True Then
                mintcmd影像类别 = mintcmd影像类别 + 1
                strTemp = IIf(strTemp = "", cbrPopControl.Caption, strTemp & "," & cbrPopControl.Caption)
            End If
            rsTemp.MoveNext
            i = i + 1
        Wend
        If strTemp <> "" Then objControl.Caption = strTemp
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("状态", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_登记, "登记")
            objControl.ToolTipText = "显示已登记病人"
        Set objControl = .Add(xtpControlButton, ID_报到, "报到")
            objControl.ToolTipText = "显示已报到病人"
        Set objControl = .Add(xtpControlButton, ID_检查, "检查")
            objControl.ToolTipText = "显示已检查病人"
        Set objControl = .Add(xtpControlButton, ID_报告, "报告")
            objControl.ToolTipText = "显示已报告病人"
        Set objControl = .Add(xtpControlButton, ID_审核, "审核")
            objControl.ToolTipText = "显示已审核病人"
        Set objControl = .Add(xtpControlButton, ID_完成, "完成")
            objControl.ToolTipText = "显示已完成病人"
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("过滤", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
        
    Set objCusControl = objBar.Controls.Add(xtpControlCustom, ID_查找值, "查找值")
        objCusControl.Handle = ucFilter.Handle
        objCusControl.Flags = xtpFlagRightAlign
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_开始查找, "开始查找")
        objControl.Style = xtpButtonIconAndCaption
        objControl.IconId = conMenu_View_Filter
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_本次住院, "本次")
    objControl.ToolTipText = "只显示本次住院检查记录"
    objControl.Style = xtpButtonIconAndCaption
    objControl.IconId = conMenu_View_Filter
    
    With cbrdock.KeyBindings
        .Add FCONTROL, vbKey0, ID_门诊
        .Add FCONTROL, vbKey1, ID_住院
        .Add FCONTROL, vbKey2, ID_外诊
        .Add FCONTROL, vbKey3, ID_体检
        
        .Add FCONTROL, vbKey4, ID_登记
        .Add FCONTROL, vbKey5, ID_报到
        .Add FCONTROL, vbKey6, ID_检查
        .Add FCONTROL, vbKey7, ID_报告
        .Add FCONTROL, vbKey8, ID_审核
        .Add FCONTROL, vbKey9, ID_完成
        .Add FCONTROL, Asc("G"), ID_开始查找
    End With
    cbrdock.RecalcLayout
End Sub

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim str3DFuncs() As String
    Dim i As Integer
    Dim i3DFunc As Integer

    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    

'菜单定义
'Begin------------------------文件菜单--------------------------------------默认可见
    Me.cbrMain.ActiveMenuBar.Title = "菜单"
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)"): cbrControl.IconId = 181
        
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)"): cbrControl.IconId = 103
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "清单打印(&L)"): cbrControl.BeginGroup = True: cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Change_In, "隐藏列表")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"):: cbrControl.IconId = 191: cbrControl.BeginGroup = True
    End With


'Begin----------------------检查菜单--------------------------------------默认可见
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "检查(&S)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告查看(&V)"): cbrControl.IconId = 102:  cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "影像观片(&S)"): cbrControl.IconId = 8111
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "观片对比(&E)"): cbrControl.IconId = 8112
        
        '如果启用三维重建功能，则创建对应菜单
        If mblnUse3D = True Then
            Set cbrControl = .Add(xtpControlPopup, conMenu_Img_3D, "三维重建"): cbrControl.ID = conMenu_Img_3D
                If mstr3DFunctions <> "" Then
                    str3DFuncs = Split(mstr3DFunctions, ",")
                    For i = 1 To UBound(str3DFuncs)
                        i3DFunc = Val(str3DFuncs(i))
                        If i3DFunc >= 1 And i3DFunc <= 6 Then
                            Select Case i3DFunc
                                Case 1
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VA, "容积重建")
                                Case 2
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                Case 3
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                Case 4
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VE, "虚拟内窥镜")
                                Case 5
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_SA, "表面重建")
                                Case 6
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_PF, "灌注成像")
                            End Select
                        End If
                    Next i
                End If
        End If
        
    End With
    
    
'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar.Controls '二级菜单
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): cbrControl.Checked = True: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Manage_LocateType, "定位方式(&G)"): cbrControl.ID = conMenu_Manage_LocateType
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_Filter * 10#, "当前科室"): cbrControl.ID = conMenu_View_Filter * 10#
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "快速过滤(&K)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&F)")
    End With


'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题", -1, False)
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB上的中联(&E)")
            With cbrControl.CommandBar.Controls
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(&F)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False)
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(cbrMain, glngSys, mlngModul, mstrPrivs)
    
'----------------------快键绑定------------------------------------------
    With Me.cbrMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print '打印------------------Ctrl+P
        
        .Add 0, VK_F1, conMenu_Help_Help              '帮助-------------F1
        .Add 0, VK_F5, conMenu_View_Refresh           '刷新-------------F5
        .Add FCONTROL, Asc("G"), conMenu_Manage_LocateType    '定位方式---------Ctrl+F
        .Add 0, VK_F3, conMenu_View_Filter            '过滤-------------F3
    End With
    
'---------------------设置右上角当前科室----------------------------------
        Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_Filter * 10#, "当前科室")
            cbrControl.ID = conMenu_View_Filter * 10#
            cbrControl.Flags = xtpFlagRightAlign
            cbrControl.Category = "Main"
            
        Set cbrCustom = cbrMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Manage_LocateValue, "定位条件")
            cbrCustom.Handle = ucLocate.Handle
            cbrCustom.Flags = xtpFlagRightAlign
            cbrCustom.Style = xtpButtonIconAndCaption
            cbrCustom.Category = "Main"
    
'---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
'    cbrToolBar.EnableDocking xtpFlagStretched '+ xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告"): cbrControl.IconId = 102: cbrControl.ToolTipText = "报告查看"
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): cbrControl.IconId = 103: cbrControl.ToolTipText = "报告打印"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "观片"): cbrControl.ToolTipText = "影像观片"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "对比"): cbrControl.IconId = 8112: cbrControl.ToolTipText = "观片对比"
        '如果启用三维重建功能，则创建对应菜单
        If mblnUse3D = True Then
            Set cbrControl = .Add(xtpControlPopup, conMenu_Img_3D, "三维"): cbrControl.ID = conMenu_Img_3D: cbrControl.ToolTipText = "三维重建"
                If mstr3DFunctions <> "" Then
                    str3DFuncs = Split(mstr3DFunctions, ",")
                    For i = 1 To UBound(str3DFuncs)
                        i3DFunc = Val(str3DFuncs(i))
                        If i3DFunc >= 1 And i3DFunc <= 6 Then
                            Select Case i3DFunc
                                Case 1
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VA, "容积重建")
                                Case 2
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                Case 3
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                Case 4
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VE, "虚拟内窥镜")
                                Case 5
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_SA, "表面重建")
                                Case 6
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_PF, "灌注成像")
                            End Select
                        End If
                    Next i
                End If
        End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
        
    End With

End Sub

Private Sub InitSubForm()
Dim i As Integer
    With TabWindow
        .RemoveAll
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTaskPanelHighlightNone
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll
        
        .InsertItem 0, "影像记录", mfrmPACSImg.Hwnd, conMenu_Img_Look
        .Item(TabWindow.ItemCount - 1).Tag = "影像图象"
        
        
    End With

End Sub


Private Sub InitList()
'初始化表格
Dim C紧急 As Long, C来源 As Long, C阳性 As Long, C质量 As Long, C姓名 As Long, C检查号 As Long, C检查过程 As Long, C性别 As Long, C年龄 As Long
Dim C标识号 As Long, C医嘱内容 As Long, C部位方法 As Long, C执行间 As Long, C报到时间 As Long, C申请时间 As Long, C开嘱医生 As Long
Dim C身高 As Long, C体重 As Long, C婴儿 As Long, C登记人 As Long, C报到人 As Long, C完成人 As Long, C打印胶片 As Long, C报告操作 As Long
Dim C绿色通道 As Long, C报告打印 As Long, C报告人 As Long, C复核人 As Long, C检查技师 As Long, C采图时间 As Long, C随访描述 As Long
Dim C影像类别 As Long, C病人ID As Long, C主页ID As Long, C挂号单 As Long, C病人科室ID As Long, C医嘱ID As Long, C发送号 As Long, C检查UID As Long
Dim C检查状态 As Long, CNO As Long, C记录性质 As Long, C转出 As Long, C床号 As Long, C当前病区ID As Long, C报告发放 As Long
Dim C诊断分类 As Long, C执行科室ID As Long, C关联ID As Long, C病人科室 As Long, C就诊卡号 As Long, C单据号 As Long, C身份证号 As Long
 
    If mstrCol = "" Then
        mstrCol = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, ConstrCol)
        '判断是否修改过显示的列数，如果修改过，则读取默认值，而不是读取注册表
        If UBound(Split(mstrCol, "|")) <> UBound(Split(ConstrCol, "|")) Then
            mstrCol = ConstrCol
        End If
    End If
    With vsList
        .Clear
        .FixedRows = 1
        .Rows = 2
        .Cols = 53
        '提取列序
        C紧急 = GetCN("紧急"):           C来源 = GetCN("来源"):          C阳性 = GetCN("阳性")
        C质量 = GetCN("质量"):          C姓名 = GetCN("姓名"):          C检查号 = GetCN("检查号")
        C检查过程 = GetCN("检查过程"):  C性别 = GetCN("性别"):          C年龄 = GetCN("年龄")
        C标识号 = GetCN("标识号"):      C医嘱内容 = GetCN("医嘱内容"):  C部位方法 = GetCN("部位方法")
        C执行间 = GetCN("执行间"):      C报到时间 = GetCN("报到时间"):  C申请时间 = GetCN("申请时间")
        C开嘱医生 = GetCN("开嘱医生"):   C身高 = GetCN("身高"):          C体重 = GetCN("体重")
        C婴儿 = GetCN("婴儿"):          C登记人 = GetCN("登记人"):      C报到人 = GetCN("报到人")
        C完成人 = GetCN("完成人"):      C打印胶片 = GetCN("打印胶片"):  C报告操作 = GetCN("报告操作")
        C绿色通道 = GetCN("绿色通道"):  C报告打印 = GetCN("报告打印"):  C报告人 = GetCN("报告人")
        C复核人 = GetCN("复核人"):      C检查技师 = GetCN("检查技师"):  C采图时间 = GetCN("采图时间")
        C随访描述 = GetCN("随访描述"):  C影像类别 = GetCN("影像类别"):  C病人ID = GetCN("病人ID")
        C主页ID = GetCN("主页ID"):      C挂号单 = GetCN("挂号单"):      C医嘱ID = GetCN("医嘱ID")
        C发送号 = GetCN("发送号"):      C病人科室ID = GetCN("病人科室ID"): C检查UID = GetCN("检查UID")
        C检查状态 = GetCN("检查状态"):  CNO = GetCN("NO"):              C记录性质 = GetCN("记录性质")
        C转出 = GetCN("转出"):          C床号 = GetCN("床号"):          C当前病区ID = GetCN("当前病区ID")
        C报告发放 = GetCN("报告发放"):  C诊断分类 = GetCN("诊断分类"):  C执行科室ID = GetCN("执行科室ID")
        C关联ID = GetCN("关联ID"):      C病人科室 = GetCN("病人科室"):  C就诊卡号 = GetCN("就诊卡号")
        C单据号 = GetCN("单据号"):      C身份证号 = GetCN("身份证号")
        '提取并指定列宽
        .ColWidth(C紧急) = GetCW("紧急"):           .ColWidth(C来源) = GetCW("来源"):           .ColWidth(C阳性) = GetCW("阳性")
        .ColWidth(C质量) = GetCW("质量"):           .ColWidth(C姓名) = GetCW("姓名"):           .ColWidth(C检查号) = GetCW("检查号")
        .ColWidth(C检查过程) = GetCW("检查过程"):   .ColWidth(C性别) = GetCW("性别"):           .ColWidth(C年龄) = GetCW("年龄")
        .ColWidth(C标识号) = GetCW("标识号"):       .ColWidth(C医嘱内容) = GetCW("医嘱内容"):   .ColWidth(C部位方法) = GetCW("部位方法")
        .ColWidth(C执行间) = GetCW("执行间"):       .ColWidth(C报到时间) = GetCW("报到时间"):   .ColWidth(C申请时间) = GetCW("申请时间")
        .ColWidth(C开嘱医生) = GetCW("开嘱医生"):   .ColWidth(C身高) = GetCW("身高"):           .ColWidth(C体重) = GetCW("体重")
        .ColWidth(C婴儿) = GetCW("婴儿"):           .ColWidth(C登记人) = GetCW("登记人"):       .ColWidth(C报到人) = GetCW("报到人")
        .ColWidth(C完成人) = GetCW("完成人"):       .ColWidth(C打印胶片) = GetCW("打印胶片"):   .ColWidth(C报告操作) = GetCW("报告操作")
        .ColWidth(C绿色通道) = GetCW("绿色通道"):   .ColWidth(C报告打印) = GetCW("报告打印"):   .ColWidth(C报告人) = GetCW("报告人")
        .ColWidth(C复核人) = GetCW("复核人"):       .ColWidth(C检查技师) = GetCW("检查技师"):   .ColWidth(C采图时间) = GetCW("采图时间")
        .ColWidth(C随访描述) = GetCW("随访描述"):   .ColWidth(C影像类别) = GetCW("影像类别"):   .ColWidth(C病人ID) = GetCW("病人ID")
        .ColWidth(C主页ID) = GetCW("主页ID"):       .ColWidth(C挂号单) = GetCW("挂号单"):       .ColWidth(C医嘱ID) = GetCW("医嘱ID")
        .ColWidth(C发送号) = GetCW("发送号"):       .ColWidth(C病人科室ID) = GetCW("病人科室ID"): .ColWidth(C检查UID) = GetCW("检查UID")
        .ColWidth(C检查状态) = GetCW("检查状态"):   .ColWidth(CNO) = GetCW("NO"):               .ColWidth(C记录性质) = GetCW("记录性质")
        .ColWidth(C转出) = GetCW("转出"):           .ColWidth(C床号) = GetCW("床号"):           .ColWidth(C当前病区ID) = GetCW("当前病区ID")
        .ColWidth(C报告发放) = GetCW("报告发放"):   .ColWidth(C诊断分类) = GetCW("诊断分类"):   .ColWidth(C执行科室ID) = GetCW("执行科室ID")
        .ColWidth(C关联ID) = GetCW("关联ID"):       .ColWidth(C病人科室) = GetCW("病人科室"):   .ColWidth(C就诊卡号) = GetCW("就诊卡号")
        .ColWidth(C单据号) = GetCW("单据号"):       .ColWidth(C身份证号) = GetCW("身份证号")
        
        '列名称
        .Cell(flexcpData, 0, C紧急) = "紧急":               .Cell(flexcpData, 0, C来源) = "来源":               .Cell(flexcpData, 0, C阳性) = "阳性"
        .Cell(flexcpData, 0, C质量) = "质量":               .Cell(flexcpData, 0, C姓名) = "姓名":               .Cell(flexcpData, 0, C检查号) = "检查号"
        .Cell(flexcpData, 0, C检查过程) = "检查过程":       .Cell(flexcpData, 0, C性别) = "性别":               .Cell(flexcpData, 0, C年龄) = "年龄"
        .Cell(flexcpData, 0, C标识号) = "标识号":           .Cell(flexcpData, 0, C医嘱内容) = "医嘱内容":       .Cell(flexcpData, 0, C部位方法) = "部位方法"
        .Cell(flexcpData, 0, C执行间) = "执行间":           .Cell(flexcpData, 0, C报到时间) = "报到时间":       .Cell(flexcpData, 0, C申请时间) = "申请时间"
        .Cell(flexcpData, 0, C开嘱医生) = "开嘱医生":       .Cell(flexcpData, 0, C身高) = "身高":               .Cell(flexcpData, 0, C体重) = "体重"
        .Cell(flexcpData, 0, C婴儿) = "婴儿":               .Cell(flexcpData, 0, C登记人) = "登记人":           .Cell(flexcpData, 0, C报到人) = "报到人"
        .Cell(flexcpData, 0, C完成人) = "完成人":           .Cell(flexcpData, 0, C打印胶片) = "打印胶片":       .Cell(flexcpData, 0, C报告操作) = "报告操作"
        .Cell(flexcpData, 0, C绿色通道) = "绿色通道":       .Cell(flexcpData, 0, C报告打印) = "报告打印":       .Cell(flexcpData, 0, C报告人) = "报告人"
        .Cell(flexcpData, 0, C复核人) = "复核人":           .Cell(flexcpData, 0, C检查技师) = "检查技师":       .Cell(flexcpData, 0, C采图时间) = "采图时间"
        .Cell(flexcpData, 0, C随访描述) = "随访描述":       .Cell(flexcpData, 0, C影像类别) = "影像类别":       .Cell(flexcpData, 0, C病人ID) = "病人ID"
        .Cell(flexcpData, 0, C主页ID) = "主页ID":           .Cell(flexcpData, 0, C挂号单) = "挂号单":           .Cell(flexcpData, 0, C病人科室ID) = "病人科室ID"
        .Cell(flexcpData, 0, C医嘱ID) = "医嘱ID":           .Cell(flexcpData, 0, C发送号) = "发送号":           .Cell(flexcpData, 0, C检查UID) = "检查UID"
        .Cell(flexcpData, 0, C检查状态) = "检查状态":       .Cell(flexcpData, 0, CNO) = "NO":                   .Cell(flexcpData, 0, C记录性质) = "记录性质"
        .Cell(flexcpData, 0, C转出) = "转出":               .Cell(flexcpData, 0, C床号) = "床号":               .Cell(flexcpData, 0, C当前病区ID) = "当前病区ID"
        .Cell(flexcpData, 0, C报告发放) = "报告发放":       .Cell(flexcpData, 0, C诊断分类) = "诊断分类":       .Cell(flexcpData, 0, C执行科室ID) = "执行科室ID"
        .Cell(flexcpData, 0, C关联ID) = "关联ID":           .Cell(flexcpData, 0, C病人科室) = "病人科室":       .Cell(flexcpData, 0, C就诊卡号) = "就诊卡号"
        .Cell(flexcpData, 0, C单据号) = "单据号":           .Cell(flexcpData, 0, C身份证号) = "身份证号"
        
        '显示列名称
        Set .Cell(flexcpPicture, 0, C紧急) = Imglist.ListImages("紧急").Picture
        Set .Cell(flexcpPicture, 0, C来源) = Imglist.ListImages("住院").Picture
        Set .Cell(flexcpPicture, 0, C阳性) = Imglist.ListImages("阳性").Picture
        .TextMatrix(0, C质量) = "质":               .TextMatrix(0, C姓名) = "姓名":              .TextMatrix(0, C检查号) = "检查号"
        .TextMatrix(0, C检查过程) = "检查过程":     .TextMatrix(0, C性别) = "性别":             .TextMatrix(0, C年龄) = "年龄"
        .TextMatrix(0, C标识号) = "标识号":         .TextMatrix(0, C医嘱内容) = "医嘱内容":     .TextMatrix(0, C部位方法) = "部位方法"
        .TextMatrix(0, C执行间) = "执行间":         .TextMatrix(0, C报到时间) = "报到时间":     .TextMatrix(0, C申请时间) = "申请时间"
        .TextMatrix(0, C开嘱医生) = "开嘱医生":     .TextMatrix(0, C身高) = "身高":             .TextMatrix(0, C体重) = "体重"
        .TextMatrix(0, C婴儿) = "婴儿":             .TextMatrix(0, C登记人) = "登记人":         .TextMatrix(0, C报到人) = "报到人"
        .TextMatrix(0, C完成人) = "完成人":         .TextMatrix(0, C打印胶片) = "打印胶片":     .TextMatrix(0, C报告操作) = "报告操作"
        .TextMatrix(0, C绿色通道) = "绿色通道":     .TextMatrix(0, C报告打印) = "报告打印":     .TextMatrix(0, C报告人) = "报告人"
        .TextMatrix(0, C复核人) = "复核人":         .TextMatrix(0, C检查技师) = "检查技师":     .TextMatrix(0, C采图时间) = "采图时间"
        .TextMatrix(0, C随访描述) = "随访描述":     .TextMatrix(0, C影像类别) = "影像类别":     .TextMatrix(0, C病人ID) = "病人ID"
        .TextMatrix(0, C主页ID) = "主页ID":         .TextMatrix(0, C挂号单) = "挂号单":         .TextMatrix(0, C病人科室ID) = "病人科室ID"
        .TextMatrix(0, C医嘱ID) = "医嘱ID":         .TextMatrix(0, C发送号) = "发送号":         .TextMatrix(0, C检查UID) = "检查UID"
        .TextMatrix(0, C检查状态) = "检查状态":     .TextMatrix(0, CNO) = "NO":                 .TextMatrix(0, C记录性质) = "记录性质"
        .TextMatrix(0, C转出) = "转出":             .TextMatrix(0, C床号) = "床号":             .TextMatrix(0, C当前病区ID) = "当前病区ID"
        .TextMatrix(0, C报告发放) = "报告发放":     .TextMatrix(0, C诊断分类) = "诊断分类":     .TextMatrix(0, C执行科室ID) = "执行科室ID"
        .TextMatrix(0, C关联ID) = "关联ID":         .TextMatrix(0, C病人科室) = "病人科室":     .TextMatrix(0, C就诊卡号) = "就诊卡号"
        .TextMatrix(0, C单据号) = "单据号":         .TextMatrix(0, C身份证号) = "身份证号"
        
        Dim i As Integer
        For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        
        '读取和设置病人列表的字体
        .FontName = zlDatabase.GetPara("病人列表内容字体", glngSys, mlngModul, "宋体")
        .FontSize = Val(zlDatabase.GetPara("病人列表内容字号", glngSys, mlngModul, 9))
        .FontBold = zlDatabase.GetPara("病人列表内容粗体", glngSys, mlngModul, 0) = 1
        .FontItalic = zlDatabase.GetPara("病人列表内容斜体", glngSys, mlngModul, 0) = 1
        .Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = zlDatabase.GetPara("病人列表表头字体", glngSys, mlngModul, "宋体")
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = Val(zlDatabase.GetPara("病人列表表头字号", glngSys, mlngModul, 9))
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = zlDatabase.GetPara("病人列表表头粗体", glngSys, mlngModul, 0) = 1
        .Cell(flexcpFontItalic, 0, 0, 0, .Cols - 1) = zlDatabase.GetPara("病人列表表头斜体", glngSys, mlngModul, 0) = 1
        .Editable = flexEDNone
    End With
End Sub




Private Sub OpenViewerWithReport()
'跟据参数打开报告后同时打开观片站，判断是否打开观片站
    Dim lngOrderID As Long
    
    On Error GoTo err
    
    lngOrderID = vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))
    
    If mblnShowImgAtReport And vsList.TextMatrix(vsList.Row, GetCN("检查UID")) <> "" Then
        Dim intImageInverval As Integer
        
        intImageInverval = Val(mfrmPACSImg.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)
        Call OpenViewer(mobjPacsCore, lngOrderID, False, Me, , , False, intImageInverval)
    End If
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'功能：显示当前执行医嘱可以打印的诊疗单据在菜单上
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
        
    On Error GoTo errH
    
    objPopup.CommandBar.Controls.DeleteAll
    With vsList
        gstrSQL = "Select Distinct C.编号,C.名称,C.说明" & _
            " From 病人医嘱记录 A,病历单据应用 B,病历文件列表 C" & _
            " Where A.ID=[1] And A.相关ID IS NULL" & _
            " And A.诊疗项目ID=B.诊疗项目ID" & _
            " And B.应用场合=[2] And B.病历文件ID=C.ID And C.种类=7" & _
            " Order by C.编号"
        If .TextMatrix(.Row, GetCN("转出")) = 1 Then
            gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
            gstrSQL = Replace(gstrSQL, "病人医嘱发送", "H病人医嘱发送")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(.TextMatrix(.Row, GetCN("医嘱ID"))), CLng(Decode(.TextMatrix(.Row, GetCN("来源")), "门", 1, "住", 2, "外", 3, 4)))
    End With
    
    If Not rsTmp.EOF Then
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + 1, rsTmp!名称 & "(&0)")
            objControl.Parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" '对应的自定义报表编号
        End With
        cbrMain.KeyBindings.Add 0, vbKeyF10, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncBillPrint(objControl As CommandBarControl)
'功能：打印诊疗单据
    On Error GoTo errH
    If objControl.Parameter = "" Then '奇怪，直接按F10时，是一个空的Control
        Set objControl = cbrMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    If objControl.Parameter = "" Then Exit Sub
    

    If ReportPrintSet(gcnOracle, glngSys, objControl.Parameter, Me) Then
        Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & vsList.TextMatrix(vsList.Row, GetCN("NO")), "性质=" & vsList.TextMatrix(vsList.Row, GetCN("记录性质")), 1)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub RefreshList(Optional ByVal lngAdviceID As Long = 0)
Dim i As Integer, lngcur医嘱ID As Long, lngRow As Long, lngTopRow As Long
    With vsList
        If lngAdviceID <> 0 Then
            lngcur医嘱ID = lngAdviceID
        Else
            lngcur医嘱ID = Val(.TextMatrix(.Row, GetCN("医嘱ID"))) '当前行医嘱ID
            lngRow = .Row: lngTopRow = .TopRow               '当前行和顶行之间的差距
        End If
        
        Call LoadPatiList
        If lngcur医嘱ID = 0 Then
            Call .Select(1, GetCN("姓名"))
            Exit Sub
        End If
        
        '有记录时要重新定位回之前记录
        On Error Resume Next
        lngcur医嘱ID = .FindRow(CStr(lngcur医嘱ID), , GetCN("医嘱ID"))
        If lngcur医嘱ID <> -1 Then
            lngRow = Abs(lngRow - lngTopRow)
            If .Row = lngcur医嘱ID Then '相同时不会触发CHANGE事件
                Call vsList_RowColChange '强制刷新右边子窗体
            Else
                .Row = lngcur医嘱ID
            End If
            .TopRow = .Row - lngRow
        Else
            If .Row <> 1 Then
                .Row = 1
            Else
                Call vsList_RowColChange '强制刷新右边子窗体
            End If
        End If
        err.Clear
    End With
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraRegist.Left = 0
    fraRegist.Top = -75
    fraRegist.Width = picInfo.ScaleWidth
    cboTimes.Left = lblRegist.Width + 30
    cboTimes.Width = fraRegist.Width - lblRegist.Width - cmdReportView.Width - 80
    
    cmdReportView.Left = fraRegist.Width - cmdReportView.Width - 20
    
    fraInfo.Top = fraRegist.Height - 20
    fraInfo.Left = 0 'fraRegist.Left + fraRegist.Width
    fraInfo.Width = picInfo.ScaleWidth '- fraInfo.Left
    
    
    lblCash.Top = 120 '(fraInfo.Height - lblCash.Height) / 2 ' (picInfo.ScaleHeight - lblCash.Height) / 2 - fraInfo.Top
    lblCash.Left = fraInfo.Width - lblCash.Width - 100

    lbl个人信息.Width = lblCash.Left
    lbl检查信息.Width = lblCash.Left
End Sub

Private Sub LoadPatiList()
'功能：读取当前医技科室的执行医嘱(病人)清单
Dim strSQL As String, strSQLBak As String, i As Long, rsList As ADODB.Recordset
Dim str来源 As String
Dim strFilter As String
Dim strModalitys As String
Dim blnUseTime As Boolean       '是否使用时间条件

    If Not mblnInitOk Then Exit Sub      '初始化未完成
    mblnvsRefresh = True
    On Error GoTo errHandle
    With SQLCondition
        blnUseTime = False  '默认不使用时间条件
        '界面查找条件不使用时间索引
        If .门诊号 <> 0 Then
            strFilter = " And C.门诊号=[1]"
        ElseIf .住院号 <> 0 Then
            strFilter = " And C.住院号=[2]"
        ElseIf .就诊卡 <> "" Then
            strFilter = " And C.就诊卡号=[3]"
        ElseIf .姓名 <> "" And InStr(.姓名, "*") = 0 Then   '姓名特殊处理，带*号表示模糊查询
            strFilter = " And C.姓名=[4]"
        ElseIf .身份证 <> "" Then
            strFilter = " And C.身份证号=[5]"
        ElseIf .IC卡 <> "" Then
            strFilter = " And C.IC卡=[6]"
        ElseIf .单据号 <> "" Then
            strFilter = " And A.NO=[7] "
        ElseIf .检查号 <> 0 Then
            strFilter = " And H.检查号=[8] "
        Else
        '其他条件查询，使用时间索引
            blnUseTime = True
            '填写过滤时间条件
            '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
            If .时间类型 = 1 Then       '按申请时间
                strFilter = " And A.发送时间 Between [9] and "
            ElseIf .时间类型 = 2 Then   '按报到时间
                strFilter = " And A.首次时间 Between [9] and "
            Else                        '采图时间
                strFilter = " And H.接收日期 Between [9] and "
            End If
            If .结束时间 <> CDate(0) Then
                strFilter = strFilter & " [10] "
            Else
                strFilter = strFilter & " Sysdate+1/(24*3600) "
            End If
            
            '先处理姓名中带*号的，进行带时间索引的模糊查询
            If .姓名 <> "" And InStr(.姓名, "*") <> 0 Then
                .姓名 = Replace(.姓名, "*", "%")
                strFilter = strFilter & " And C.姓名 like [4]"
            End If
            
            If .性别 <> "" Then
                strFilter = strFilter & " And Nvl(H.性别,C.性别)=[29]"
            End If
        
        
            '病人年龄-开始年龄(只有当条件使用“到”，即在多少年龄之间时，才使用开始年龄)
            If .开始年龄 <> -1 Then
                If .年龄条件 = "~" Then
                    strFilter = strFilter & " And ZL_AgeToDays(C.年龄)>=[30]"
                End If
            End If
            
            '病人年龄-结束年龄
            If .结束年龄 <> -1 Then
                If .年龄条件 = "~" Then
                    strFilter = strFilter & " And ZL_AgeToDays(C.年龄)<=[31]"
                Else
                    strFilter = strFilter & " And ZL_AgeToDays(C.年龄)" & .年龄条件 & "[31]"
                End If
            End If
            
            If .病人科室 <> 0 Then
                strFilter = strFilter & " And B.病人科室ID+0=[11] "
            End If
        
            If .标本部位 <> "" Then
                strFilter = strFilter & " And instr(B.医嘱内容,[12])>0"
            End If
            
            If .结果阳性 <> -1 Then
                strFilter = strFilter & " And Nvl(A.结果阳性, 0)=[32]"
            End If
            
            If .诊断医生 <> "" Then
                strFilter = strFilter & " And H.报告人=[13] "
            End If
            
            If .审核医生 <> "" Then
                strFilter = strFilter & " And H.复核人=[14] "
            End If
            
            If .影像质量 <> "" Then
                strFilter = strFilter & " And H.影像质量=[15]"
            End If
            
            If .检查技师 <> "" Then
                strFilter = strFilter & " And H.检查技师=[16]"
            End If
            
            '影像类别有两个地方做过滤条件的选择，过滤窗口和主程序上面，以主程序中的为主
            If mintcmd影像类别 > 0 Then
                Dim objControl As CommandBarControl
                
                Set objControl = cbrdock.FindControl(, ID_影像类别)
                For i = 1 To objControl.CommandBar.Controls.Count
                    If objControl.CommandBar.FindControl(, ID_影像类别 + i).Checked = True Then
                        strModalitys = strModalitys & "," & objControl.CommandBar.FindControl(, ID_影像类别 + i).DescriptionText
                    End If
                Next i
                If strModalitys <> "" Then
                    strFilter = strFilter & " And instr([26],H.影像类别)>0 "
                End If
            Else
                If .影像类别 <> "" Then
                    strFilter = strFilter & " And H.影像类别=[17] "
                End If
            End If
            
            
            
            If .随访 <> "" Then
                strFilter = strFilter & " And  Instr(H.随访描述, [18]) > 0 "
            End If
            
            If .疾病诊断 <> "" Then
                strFilter = strFilter & " And B.ID IN ( Select t.医嘱id From 病人医嘱报告 t Where t.病历id In " & _
                                                                    " (Select Distinct A.ID  " & _
                                                                        "From 电子病历记录 A,电子病历内容 B " & _
                                                                        "Where A.创建时间>[9] AND A.Id=B.文件ID  " & _
                                                                            "And B.对象类型=7 And instr(B.对象属性,'52;')>0 And instr(B.内容文本,[19])>0))"
            End If
            
            Dim strSubFilter As String '增加PACS报告检索条件
            If .检查所见 <> "" Then
                strSubFilter = " (b.内容文本 ='检查所见' And Instr(c.内容文本, [20]) > 0)"
            End If
            
            If .诊断意见 <> "" Then
                If strSubFilter = "" Then
                    strSubFilter = " (b.内容文本 ='诊断意见' And Instr(c.内容文本, [21]) > 0)"
                Else
                    strSubFilter = strSubFilter & " or (b.内容文本 ='诊断意见' And Instr(c.内容文本, [21]) > 0)"
                End If
            End If
            
            If .建议 <> "" Then
                If strSubFilter = "" Then
                    strSubFilter = " (b.内容文本 ='建议' And Instr(c.内容文本, [22]) > 0)"
                Else
                    strSubFilter = strSubFilter & " or (b.内容文本 ='建议' And Instr(c.内容文本, [22]) > 0)"
                End If
            End If
            
            If strSubFilter <> "" Then
                strSubFilter = " (" & strSubFilter & ")"
                
                strFilter = strFilter & " And B.ID IN ( Select t.医嘱id From 病人医嘱报告 t Where t.病历id In  " _
                    & " (Select Distinct a.ID From 电子病历记录 a, 电子病历内容 b,电子病历内容 c " _
                    & " Where a.创建时间 > [9] And a.Id = b.文件id And b.Id = C.父ID And b.对象类型 = 3 And c.对象类型 = 2 And c.终止版 = 0 and " _
                    & strSubFilter & "))"
            End If
           
'            If .检查过程 <> "" Then
'                If .检查过程 = "全部" Then
'
'                ElseIf .检查过程 = "已登记" Then
'                    strFilter = strFilter & " And (A.执行过程 =0 or A.执行过程=1 Or A.执行过程 Is Null) "
'                ElseIf .检查过程 = "已报到" Then
'                    strFilter = strFilter & " And (A.执行过程 = 2 and h.报告人 is null) "
'                ElseIf .检查过程 = "已检查" Then
'                    strFilter = strFilter & " And (A.执行过程 = 3 and h.报告人 is null) "
'                ElseIf .检查过程 = "处理中" Then
'                    strFilter = strFilter & " And (not h.报告操作 is null) "
'                ElseIf .检查过程 = "报告中" Then
'                    strFilter = strFilter & " And ((A.执行过程 =2 or A.执行过程=3) and not h.报告人 is null and h.报告操作 is null) "
'                ElseIf .检查过程 = "已报告" Then
'                    strFilter = strFilter & " And (A.执行过程=4 and h.复核人 is null) "
'                ElseIf .检查过程 = "审核中" Then
'                    strFilter = strFilter & " And (A.执行过程=4 and not h.复核人 is null) "
'                ElseIf .检查过程 = "已审核" Then
'                    strFilter = strFilter & " And A.执行过程=5 "
'                ElseIf .检查过程 = "已完成" Then
'                    strFilter = strFilter & " And A.执行过程=6 "
'                End If
'            End If
        End If
        
        '“过滤窗口”和“界面查找”条件独立，界面查找条件不使用时间索引，以下条件为共用条件
        
        '病人来源 (1-门诊,2-住院,3-外来,4-体检)
        '如果四种来源都选择了，表示查找所有病人，则不添加病人来源的查询条件
        If mblncmd门诊 And mblncmd住院 And mblncmd体检 And mblncmd外诊 Then
        
        Else
            If mblncmd门诊 Then str来源 = "1,"
            If mblncmd住院 Then str来源 = str来源 & "2,"
            If mblncmd外诊 Then str来源 = str来源 & "3,"
            If mblncmd体检 Then str来源 = str来源 & "4,"
            If str来源 <> "" Then   'str来源为空，表示没有选择任何来源，则不添加病人来源的查询条件
                str来源 = Mid(str来源, 1, Len(str来源) - 1)
                strFilter = strFilter & " And Instr([23],B.病人来源)> 0"
            End If
        End If
        
'        If mstrRoom <> "" Then  '只显示执行间范围内的
'            If Not mblncmd登记 Then
'                strFilter = strFilter & " And Instr([24],','|| A.执行间 || ',' )>0"
'            Else
'                strFilter = strFilter & " And (Instr([24],','|| A.执行间 || ',' )>0 And Nvl(A.执行过程,0)>1 OR Nvl(A.执行过程,0)<2)"
'            End If
'        End If
    
'        If mblnNoShowCancel Then '不显示取消登记的检查
'            strFilter = strFilter & " And A.执行状态<>2 "
'        End If
        
        If mblncmd本次 Then        '只显示本次住院记录
            strFilter = strFilter & vbNewLine & " And (B.病人来源=2 And B.主页ID=C.住院次数 Or Nvl(B.病人来源,0)<>2)"
        End If

        '是否选择了全部科室
        If mblnAllDepts = True Then
            strFilter = strFilter & " And (Instr( [27], B.执行科室ID ) >0  or Instr( [27], B.开嘱科室ID ) > 0) "
        Else
            strFilter = strFilter & " AND (B.执行科室ID + 0 =[25] or B.开嘱科室ID + 0 = [25])"
        End If


        
         
        '检索报告内容
        If .报告内容 <> "" Then
            strFilter = strFilter & " And B.id IN ( Select t.医嘱id From 病人医嘱报告 t Where t.病历id In " & _
                                                                    " (Select Distinct A.ID " & _
                                                                    " From 电子病历记录 A,电子病历内容 B " & _
                                                                    " Where A.创建时间>[9] AND A.Id=B.文件ID " & _
                                                                    " And B.对象类型=2 And instr(B.内容文本,[28])>0 And B.终止版 = 0)) "
        End If
        
        gstrSQL = "Select /*+ RULE */ Distinct" & vbNewLine & _
                    "       A.医嘱ID,A.发送号,A.首次时间 报到时间,A.发送时间 申请时间,A.执行状态,nvl(A.执行过程,0) 检查过程,A.执行间,A.结果阳性 阳性," & vbNewLine & _
                    "       B.病人ID,B.主页ID,B.挂号单,B.病人科室ID,Decode(B.病人来源, 1, '门', 2, '住', 3, '外', 4, '体') 来源,B.医嘱内容,B.标本部位," & vbNewLine & _
                    "       Nvl(B.紧急标志, 0) 紧急标志, Nvl(B.婴儿, 0) 婴儿,B.开嘱医生,A.NO,C.当前床号,C.当前病区ID,Decode(B.病人来源,2,C.住院号,C.门诊号) 标识号," & vbNewLine & _
                    "       Nvl(H.姓名,C.姓名) 姓名,H.影像类别,H.检查号,Nvl(H.性别,C.性别) 性别,Nvl(H.年龄,C.年龄) 年龄,H.身高,H.体重,H.影像质量," & vbNewLine & _
                    "       Decode(B.病人来源,3,B.开嘱医生,A.发送人) 登记人,H.报到人,H.报告发放,H.关联ID, " & vbNewLine & _
                    "       H.完成人,H.是否打印,H.报告操作,H.绿色通道,H.报告打印,H.报告人,H.复核人,H.检查技师,H.接收日期 采图时间," & vbNewLine & _
                    "       H.随访描述,H.诊断分类,H.检查UID,A.执行部门ID as 执行科室ID,0 as 转出,F.名称 AS 病人科室, " & vbNewLine & _
                    "       C.就诊卡号,A.NO as 单据号,C.身份证号 " & vbNewLine & _
                    " From 病人医嘱发送 A,病人医嘱记录 B,病人信息 C,影像检查记录 H,影像检查项目 G,部门表 F " & vbNewLine & _
                    " Where B.相关ID is NULL And A.医嘱ID=B.ID And A.医嘱ID=H.医嘱ID(+) And A.发送号=H.发送号(+) " & vbNewLine & _
                    " And B.诊疗项目ID=G.诊疗项目ID And B.病人ID=C.病人ID And B.病人科室id=F.ID "
        gstrSQL = gstrSQL & vbNewLine & strFilter
        
        If mblncmd已缴 Xor mblncmd未缴 Then '互斥选择
            '根据病人来源的过滤条件判断查询哪些费用表,
            '同时查询门诊和住院费用表的情况：四个来源都选择；四个来源都不选择；选择了住院同时还选择了其他任何一个来源
            '只查住院表的情况：只选择了住院
            '只查门诊表的情况：只选择了门诊、外诊、体检，这三种来源
            
            If mblncmd住院 = True And mblncmd门诊 = False And mblncmd外诊 = False And mblncmd体检 = False Then
                '只查询住院表
                strFilter = "Select Distinct NO From 住院费用记录 D Where A.NO = D.NO And A.记录性质 = D.记录性质 And D.记录状态 = 1"
            ElseIf mblncmd住院 = False And (mblncmd门诊 = True Or mblncmd外诊 = True Or mblncmd体检 = True) Then
                '只查门诊表
                strFilter = "Select Distinct NO From 门诊费用记录 E Where A.NO = E.NO And A.记录性质 = E.记录性质 And E.记录状态 = 1"
            Else    '其他情况同时查两个表
                strFilter = "Select Distinct NO From 住院费用记录 D Where A.NO = D.NO And A.记录性质 = D.记录性质 And D.记录状态 = 1" & vbNewLine & _
                            "Union" & vbNewLine & _
                            "Select Distinct NO From 门诊费用记录 E Where A.NO = E.NO And A.记录性质 = E.记录性质 And E.记录状态 = 1"
            End If
            
            gstrSQL = gstrSQL & vbNewLine & IIf(mblncmd已缴, " And Exists ", " And Not Exists") & "(" & strFilter & ")"
        End If
        
        '当使用检查号查找时一定是报过到的，影像检查记录中有记录，此时取消左连接避免全表扫描
        '使用采集时间过滤，影像检查记录中有记录
        If .检查号 <> 0 Or (blnUseTime = True And SQLCondition.时间类型 = 3) Then
            gstrSQL = Replace(Replace(gstrSQL, "H.医嘱ID(+)", "H.医嘱ID"), "H.发送号(+)", "H.发送号")
        End If
        
        '如果有数据转出则还要检索后备表
        If mblnMoved Then
            strSQLBak = gstrSQL
            strSQLBak = Replace(strSQLBak, "病人医嘱记录", "H病人医嘱记录")
            strSQLBak = Replace(strSQLBak, "病人医嘱发送", "H病人医嘱发送")
            strSQLBak = Replace(strSQLBak, "影像检查记录", "H影像检查记录")
            strSQLBak = Replace(strSQLBak, "电子病历记录", "H电子病历记录")
            strSQLBak = Replace(strSQLBak, "电子病历内容", "H电子病历内容")
            strSQLBak = Replace(strSQLBak, "0 as 转出", "1 as 转出")
            strSQL = strSQL & " Union ALL " & strSQLBak
        End If
        gstrSQL = "Select * From (" & vbNewLine & gstrSQL & vbNewLine & ") Order by 检查过程,报到时间,申请时间"
    
        Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人列表", .门诊号, .住院号, .就诊卡, .姓名, .身份证, .IC卡, .单据号, _
                                            .检查号, .开始时间, .结束时间, .病人科室, .标本部位, .诊断医生, .审核医生, .影像质量, _
                                            .检查技师, .影像类别, .随访, .疾病诊断, .检查所见, .诊断意见, .建议, str来源, "", _
                                           glngDeptId, strModalitys, mstr医生所属科室IDs, .报告内容, .性别, .开始年龄, .结束年龄, .结果阳性)
    End With

    strFilter = ""
    If mblncmd登记 Then strFilter = "检查过程=0 or 检查过程=1 or "
    If mblncmd报到 Then strFilter = IIf(strFilter <> "", strFilter & "检查过程=2 or ", "检查过程=2 or ")
    If mblncmd检查 Then strFilter = IIf(strFilter <> "", strFilter & "检查过程=3 or ", "检查过程=3 or ")
    If mblncmd报告 Then strFilter = IIf(strFilter <> "", strFilter & "检查过程=4 or ", "检查过程=4 or ")
    If mblncmd审核 Then strFilter = IIf(strFilter <> "", strFilter & "检查过程=5 or ", "检查过程=5 or ")
    If mblncmd完成 Then strFilter = IIf(strFilter <> "", strFilter & "检查过程=6 or ", "检查过程=6 or ")
    
    If mblncmd登记 And mblncmd报到 And mblncmd检查 And mblncmd报告 Then ' And mblncmd审核 And mblncmd完成 Then
        strFilter = ""
    End If

    If strFilter <> "" Then
        strFilter = Mid(strFilter, 1, Len(strFilter) - 4)
        rsList.Filter = strFilter
    End If
    
    Call FillList(rsList)
    mblnvsRefresh = False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillList(ByVal rsTemp As ADODB.Recordset)
Dim rsBaby As ADODB.Recordset
    On Error GoTo errHandle
    Call InitList
    If rsTemp.EOF Then stbThis.Panels(2).Text = "没有找到任何匹配的记录": Exit Sub
    
    With vsList
        .Rows = rsTemp.RecordCount + 1
        Do Until rsTemp.EOF
            .Row = rsTemp.AbsolutePosition
            .Cell(flexcpData, .Row, GetCN("紧急")) = Val(rsTemp!紧急标志)
            If rsTemp!紧急标志 <> 0 Then
                Set .Cell(flexcpPicture, .Row, GetCN("紧急")) = Imglist.ListImages("紧急").Picture
            End If
            If rsTemp!来源 = "住" Then
                Set .Cell(flexcpPicture, .Row, GetCN("来源")) = Imglist.ListImages("住院").Picture
            End If
            .TextMatrix(.Row, GetCN("来源")) = rsTemp!来源
            .Cell(flexcpData, .Row, GetCN("来源")) = Decode(rsTemp!来源, "门", 1, "住", 2, "外", 3, 4)
            
            If Nvl(rsTemp!阳性, 0) <> 0 Then
                Set .Cell(flexcpPicture, .Row, GetCN("阳性")) = Imglist.ListImages("阳性").Picture
            End If
            
            If Nvl(rsTemp!绿色通道, 0) <> 0 Then
                Set .Cell(flexcpPicture, .Row, GetCN("姓名")) = Imglist.ListImages("绿色通道").Picture
            End If
            
            If Nvl(rsTemp!检查uid) <> "" Then
                Set .Cell(flexcpPicture, .Row, GetCN("检查号")) = Imglist.ListImages("影像").Picture
            End If
            
            .TextMatrix(.Row, GetCN("质量")) = Nvl(rsTemp!影像质量)
            .TextMatrix(.Row, GetCN("姓名")) = Nvl(rsTemp!姓名)
            .TextMatrix(.Row, GetCN("检查号")) = Nvl(rsTemp!检查号)
            .TextMatrix(.Row, GetCN("检查过程")) = IIf(rsTemp!执行状态 = 2, "已拒绝", Decode(Nvl(rsTemp!检查过程, 0), 0, "已登记", 1, "已登记", _
                                                                                        2, IIf(Nvl(rsTemp!报告操作) <> "", "处理中", _
                                                                                                IIf(Nvl(rsTemp!报告人) = "", "已报到", "报告中")), _
                                                                                        3, IIf(Nvl(rsTemp!报告操作) <> "", "处理中", _
                                                                                                IIf(Nvl(rsTemp!报告人) = "", "已检查", "报告中")), _
                                                                                        4, IIf(Nvl(rsTemp!报告操作) <> "", "处理中", _
                                                                                                IIf(Nvl(rsTemp!复核人) <> "", "审核中", "已报告")), _
                                                                                        5, "已审核", "已完成"))
            .TextMatrix(.Row, GetCN("性别")) = Nvl(rsTemp!性别)
            .TextMatrix(.Row, GetCN("年龄")) = Nvl(rsTemp!年龄)
            If InStr(Nvl(rsTemp!医嘱内容), ":") > 0 Then '新的模式保存在医嘱内容中信息是 名称,执行标记:部位(方法,方法),部位---
                .TextMatrix(.Row, GetCN("医嘱内容")) = Split(rsTemp!医嘱内容, ":")(0)
                .TextMatrix(.Row, GetCN("部位方法")) = Split(rsTemp!医嘱内容, ":")(1)
            Else
                .TextMatrix(.Row, GetCN("医嘱内容")) = Nvl(rsTemp!医嘱内容)
            End If
            .TextMatrix(.Row, GetCN("执行间")) = Nvl(rsTemp!执行间)
            .TextMatrix(.Row, GetCN("报到时间")) = Nvl(rsTemp!报到时间)
            .TextMatrix(.Row, GetCN("申请时间")) = Nvl(rsTemp!申请时间)
            .TextMatrix(.Row, GetCN("开嘱医生")) = Nvl(rsTemp!开嘱医生)
            .TextMatrix(.Row, GetCN("身高")) = Nvl(rsTemp!身高)
            .TextMatrix(.Row, GetCN("体重")) = Nvl(rsTemp!体重)
            .TextMatrix(.Row, GetCN("婴儿")) = Nvl(rsTemp!婴儿)
            .TextMatrix(.Row, GetCN("登记人")) = Nvl(rsTemp!登记人)
            .TextMatrix(.Row, GetCN("报到人")) = Nvl(rsTemp!报到人)
            .TextMatrix(.Row, GetCN("完成人")) = Nvl(rsTemp!完成人)
            .TextMatrix(.Row, GetCN("打印胶片")) = IIf(Nvl(rsTemp!是否打印) = 1, "已打印", "未打印")
            .TextMatrix(.Row, GetCN("报告操作")) = Nvl(rsTemp!报告操作)
            .TextMatrix(.Row, GetCN("绿色通道")) = Nvl(rsTemp!绿色通道)
            .TextMatrix(.Row, GetCN("报告打印")) = IIf(Nvl(rsTemp!报告打印) = 1, "已打印", "未打印")
            .TextMatrix(.Row, GetCN("报告人")) = Nvl(rsTemp!报告人)
            .TextMatrix(.Row, GetCN("复核人")) = Nvl(rsTemp!复核人)
            .TextMatrix(.Row, GetCN("检查技师")) = Nvl(rsTemp!检查技师)
            .TextMatrix(.Row, GetCN("采图时间")) = Nvl(rsTemp!采图时间)
            .TextMatrix(.Row, GetCN("影像类别")) = Nvl(rsTemp!影像类别)
            .TextMatrix(.Row, GetCN("病人ID")) = Nvl(rsTemp!病人ID, 0)
            .TextMatrix(.Row, GetCN("主页ID")) = Nvl(rsTemp!主页ID, 0)
            .TextMatrix(.Row, GetCN("挂号单")) = Nvl(rsTemp!挂号单)
            .TextMatrix(.Row, GetCN("病人科室ID")) = Nvl(rsTemp!病人科室ID, 0)
            .TextMatrix(.Row, GetCN("医嘱ID")) = Nvl(rsTemp!医嘱id)
            .TextMatrix(.Row, GetCN("发送号")) = Nvl(rsTemp!发送号)
            .TextMatrix(.Row, GetCN("检查UID")) = Nvl(rsTemp!检查uid)
            .TextMatrix(.Row, GetCN("检查状态")) = Nvl(rsTemp!检查过程)
            .TextMatrix(.Row, GetCN("随访描述")) = Nvl(rsTemp!随访描述)
            .TextMatrix(.Row, GetCN("NO")) = Nvl(rsTemp!no)
            .TextMatrix(.Row, GetCN("转出")) = Nvl(rsTemp!转出)
            .TextMatrix(.Row, GetCN("床号")) = Nvl(rsTemp!当前床号)
            .TextMatrix(.Row, GetCN("当前病区ID")) = Nvl(rsTemp!当前病区ID, 0)
            .TextMatrix(.Row, GetCN("标识号")) = Nvl(rsTemp!标识号)
            .TextMatrix(.Row, GetCN("报告发放")) = IIf(Nvl(rsTemp!报告发放, 0) = 0, "未发放", "已发放")
            .TextMatrix(.Row, GetCN("诊断分类")) = Nvl(rsTemp!诊断分类)
            .TextMatrix(.Row, GetCN("执行科室ID")) = Nvl(rsTemp!执行科室ID)
            .TextMatrix(.Row, GetCN("关联ID")) = Nvl(rsTemp!关联ID, 0)
            .TextMatrix(.Row, GetCN("病人科室")) = Nvl(rsTemp!病人科室)
            .TextMatrix(.Row, GetCN("就诊卡号")) = Nvl(rsTemp!就诊卡号)
            .TextMatrix(.Row, GetCN("单据号")) = Nvl(rsTemp!单据号)
            .TextMatrix(.Row, GetCN("身份证号")) = Nvl(rsTemp!身份证号)
            
            If Nvl(rsTemp!婴儿) <> 0 Then
                gstrSQL = "Select Nvl(A.婴儿姓名, B.姓名 || '之子' || Trim(To_Char(A.序号, '9'))) As 婴儿姓名, 婴儿性别, 出生时间" & vbNewLine & _
                            "From 病人新生儿记录 A, 病人信息 B" & vbNewLine & _
                            "Where A.病人id = [1] And A.主页id = [2] And A.病人id = B.病人id And A.序号 = [3]"

                Set rsBaby = zlDatabase.OpenSQLRecord(gstrSQL, "提取婴儿信息", CLng(rsTemp!病人ID), CLng(Nvl(rsTemp!主页ID, 0)), CLng(rsTemp!婴儿))
                If Not rsBaby.EOF Then
                    .TextMatrix(.Row, GetCN("姓名")) = rsBaby!婴儿姓名
                    .TextMatrix(.Row, GetCN("性别")) = Nvl(rsBaby!婴儿性别)
                    .TextMatrix(.Row, GetCN("年龄")) = Nvl(rsBaby!出生时间)
                End If
            End If
            
            If .TextMatrix(.Row, GetCN("检查过程")) = "已拒绝" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已拒绝
            If .TextMatrix(.Row, GetCN("检查过程")) = "已完成" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已完成
            If .TextMatrix(.Row, GetCN("检查过程")) = "已报到" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已报到
            If .TextMatrix(.Row, GetCN("检查过程")) = "已登记" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已登记
            If .TextMatrix(.Row, GetCN("检查过程")) = "已检查" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已检查
            If .TextMatrix(.Row, GetCN("检查过程")) = "已审核" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已审核
            If .TextMatrix(.Row, GetCN("检查过程")) = "处理中" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor处理中
            If .TextMatrix(.Row, GetCN("检查过程")) = "报告中" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor报告中
            If .TextMatrix(.Row, GetCN("检查过程")) = "审核中" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor审核中
            If .TextMatrix(.Row, GetCN("检查过程")) = "已报告" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已报告
            
            rsTemp.MoveNext
        Loop
    End With
    
    '恢复排序
    If mlngSortCol <> 0 And mintSortOrder <> 0 Then
        If mlngSortCol < vsList.Cols Then
            vsList.Col = mlngSortCol
            vsList.Sort = mintSortOrder
        End If
    End If
    
    stbThis.Panels(2).Text = "共 " & vsList.Rows - 1 & " 条记录": stbThis.Panels(2).Alignment = sbrCenter
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub PicWindow_Resize()
    On Error Resume Next
    With picInfo
        .Top = 0
        .Left = 0
        .Width = PicWindow.ScaleWidth
    End With
        
    With TabWindow
        .Top = picInfo.ScaleHeight
        .Left = 0
        .Width = PicWindow.ScaleWidth
        .Height = PicWindow.ScaleHeight - picInfo.ScaleHeight
    End With
End Sub


Private Sub TimerRefresh_Timer()
    '刷新病人列表
    Call RefreshList
End Sub


Private Sub SeekNextPati(ByVal blnFirst As Boolean, ByVal strCardName As String, _
    ByVal strCardText As String, ByVal lngPatientID As Long)
'------------------------------------------------
'功能：在病人列表中定位指定的记录
'参数： blnFirst -- 是否第一次查找
'返回：无，直接在病人列表中定位
'------------------------------------------------
    Dim blnOk As Boolean, lngCount As Long, intB As Integer
    Dim lngRow As Long

    '如果没有记录，则退出
    If Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) = 0 Then Exit Sub

    intB = 0
    
    On Error GoTo err
    
    If Not blnFirst Then
        intB = vsList.Row + 1
        If intB >= vsList.Rows Then intB = 1
    End If

    blnOk = False
    For lngCount = intB To vsList.Rows - 1 '在当前状态中查找
        Select Case strCardName
            Case "标识号", "住院号", "门诊号"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("标识号")), 0)) Like UCase(strCardText) & "*" Then blnOk = True
            Case "就诊卡", "IC卡号", "IC卡"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("就诊卡号")), 0)) Like UCase(strCardText) & "*" Then blnOk = True
            Case "单据号"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("NO")), 0)) Like UCase(strCardText) & "*" Then blnOk = True
            Case "检查号"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("检查号")), 0)) Like UCase(strCardText) & "*" Then blnOk = True
            Case "姓名", "姓 名", "姓  名", "姓   名"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("姓名")), "")) Like UCase(strCardText) & "*" Then blnOk = True
                If zlCommFun.SpellCode(Nvl(vsList.TextMatrix(lngCount, GetCN("姓名")), "")) Like UCase(strCardText) & "*" Then blnOk = True
            Case "身份证号", "身份证"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("身份证号")), 0)) Like UCase(strCardText) & "*" Then blnOk = True
            Case Else
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("病人ID")), 0)) = UCase(lngPatientID) Then blnOk = True
        End Select

        If blnOk Then
            ucLocate.Tag = ucLocate.CardText
            On Error Resume Next
            '计算当前行和顶行之间的差距
            lngRow = Abs(vsList.Row - vsList.TopRow)

            vsList.Row = lngCount
            vsList.TopRow = vsList.Row - lngRow

            Exit Sub
        End If
    Next
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub sub3DProcess(strCommand As String, strImageDir As String)
    Dim str3DCommand As String
    
    '组织三维重建语句
    str3DCommand = mstr3DExeDir & " " & mstr3DPara & " " & strCommand & " " & strImageDir
    On Error Resume Next
    Shell str3DCommand
End Sub

Private Sub sub三维重建(strCommand As String)
    Dim strImageDir As String
    
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), vsList.TextMatrix(vsList.Row, GetCN("发送号")), mstrPrivs, vsList.TextMatrix(vsList.Row, GetCN("转出")) = 1)
    End If
    
    '组织三维重建需要的图像
    strImageDir = mfrmPACSImg.ZLfun3DImgProcess
    If strImageDir <> "" Then
        Call sub3DProcess(strCommand, strImageDir)
    End If
End Sub


Private Sub ucFilter_OnClick(ByVal strCardName As String, ByVal strCardText As String, ByVal lngKindId As Long, ByVal lngCardLen As Long, ByVal lngSwipingType As Long, ByVal blnIsPwdInput As Boolean)
'当单击该控件时处理读卡
On Error GoTo errHandle
    Dim lngPatientID As Long
    
    '如果为1则处理读卡
    If lngSwipingType = 1 Then ucFilter.CardText = ucFilter.ReadCard(lngPatientID)
    
    Call subRefreshFilterCondition(strCardName, ucFilter.CardText, lngPatientID)
    Call RefreshList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucFilter_OnRead(ByVal strCardName As String, ByVal strCardText As String, ByVal lngPatientID As Long)
'开始查找数据
On Error GoTo errHandle
    Call subRefreshFilterCondition(strCardName, strCardText, lngPatientID)
    Call RefreshList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucFilter_OnResize()
On Error Resume Next
    Call cbrdock.RecalcLayout
err.Clear
End Sub

Private Sub ucLocate_OnClick(ByVal strCardName As String, ByVal strCardText As String, ByVal lngKindId As Long, ByVal lngCardLen As Long, ByVal lngSwipingType As Long, ByVal blnIsPwdInput As Boolean)
'当单击该控件时处理读卡
On Error GoTo errHandle
    Dim lngPatientID As Long
    
    '如果为1则处理读卡
    If lngSwipingType = 1 Then ucLocate.CardText = ucLocate.ReadCard(lngPatientID)
    
    Call SeekNextPati(ucLocate.Tag <> ucLocate.CardText, strCardName, ucLocate.CardText, lngPatientID)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucLocate_OnRead(ByVal strCardName As String, ByVal strCardText As String, ByVal lngPatientID As Long)

    
On Error GoTo errHandle
     Call SeekNextPati(ucLocate.Tag <> ucLocate.CardText, strCardName, strCardText, lngPatientID)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
Dim i As Integer, strCol As String
    For i = 0 To vsList.Cols - 1
        strCol = strCol & "|" & vsList.Cell(flexcpData, 0, i) & ";" & vsList.ColWidth(i)
    Next
    mstrCol = Mid(strCol, 2)
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'功能: 显示病人卡片按钮
    If vsList.TextMatrix(NewRow, GetCN("医嘱ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If vsList.LeftCol > GetCN("姓名") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, NewRow, GetCN("姓名")) + vsList.Cell(flexcpWidth, NewRow, GetCN("姓名")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("姓名")) + 15
            cmdInfo.Visible = True
        End If
    End If
End Sub
Private Sub vsList_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'功能:显示病人卡片按钮
    If vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If NewLeftCol > GetCN("姓名") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, vsList.Row, GetCN("姓名")) + vsList.Cell(flexcpWidth, vsList.Row, GetCN("姓名")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("姓名")) + 15
            cmdInfo.Visible = True
        End If
    End If
End Sub

Private Sub vsList_AfterSort(ByVal Col As Long, Order As Integer)
    mlngSortCol = Col
    mintSortOrder = Order
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'功能:显示病人卡片按钮
    If vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If vsList.LeftCol > GetCN("姓名") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, vsList.Row, GetCN("姓名")) + vsList.Cell(flexcpWidth, vsList.Row, GetCN("姓名")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("姓名")) + 15
            cmdInfo.Visible = True
        End If
    End If
    
    Dim i As Integer, strCol As String
    For i = 0 To vsList.Cols - 1 '暂存列序列宽，窗体关闭时存于注册表
        strCol = strCol & "|" & vsList.Cell(flexcpData, 0, i) & ";" & vsList.ColWidth(i)
    Next
    mstrCol = Mid(strCol, 2)
End Sub

Private Sub vsList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < GetCN("姓名") Then Cancel = True
End Sub

Private Sub vsList_DblClick()

    '没报告不能打印和预览
    If vsList.TextMatrix(vsList.Row, GetCN("报告人")) = "" Then
        MsgBoxD Me, "当前病人没有检查报告，不能操作，请检查！", vbInformation, gstrSysName
        Exit Sub
    End If
            
            
    If vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) <> "" Then
        mlngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))
        Call OpenReportPreview(mlngAdviceID)
    End If
    
End Sub

Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Dim control As CommandBarControl, Menucontrol As CommandBarControl
        Dim Popup As CommandBar
        Set Popup = cbrMain.Add("右键菜单", xtpBarPopup)
        For Each Menucontrol In cbrMain.ActiveMenuBar.Controls
'            If Menucontrol.Parent.BarID = conMenu_ManagePopup Then
            If (Menucontrol.ID <> conMenu_FilePopup And Menucontrol.ID <> conMenu_ToolPopup _
                And Menucontrol.ID <> conMenu_ViewPopup And Menucontrol.ID <> conMenu_HelpPopup) And Menucontrol.Type = xtpControlPopup Then
                
                For Each control In Menucontrol.CommandBar.Controls
                    control.Copy Popup
                Next
            End If
        Next
        Popup.ShowPopup
    End If
End Sub

Private Sub vsList_RowColChange()
    On Error GoTo errHandle
'    mblnIsHistory = False

    If mlngAdviceID = Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) Then Exit Sub

    mlngAdviceID = Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")))
    mstrHStudyUID = vsList.TextMatrix(vsList.Row, GetCN("检查UID"))
    
    If Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) = 0 Then  '无记录时处理
        Call RefreshTabWindow(0, True)
        cboTimes.Clear
        txtAppend = ""
        lbl个人信息.Caption = "姓  名:" & Space(12) & "性  别:" & Space(13) & "年  龄:" & Space(10) & "标识号:" & Space(12) & "床  号:" & Space(10)
        lbl检查信息.Caption = "检查号:" & Space(12) & "病人科室:" & Space(11) & "开嘱医生:" & Space(8) & "检查项目:"
        lblCash.Visible = False
    Else
        
        Call FillHistory '填充历次检查记录
        Call FillTxtInfor '填充右上方病人基本信息
        Call FillTxtAppend '填充左下角医嘱附件

        Call RefreshTabWindow

    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillTxtInfor(Optional lngAdviceID As Long = 0)
'填充右上方病人基本信息
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    With vsList
        lbl个人信息.Caption = "姓  名:" & Rpad(.TextMatrix(.Row, GetCN("姓名")), 12, " ") & "性  别:" & Rpad(.TextMatrix(.Row, GetCN("性别")), 13, " ") & _
                          "年  龄:" & Rpad(.TextMatrix(.Row, GetCN("年龄")), 10, " ") & "标识号:" & Rpad(.TextMatrix(.Row, GetCN("标识号")), 12, " ") & _
                          "床  号:" & Rpad(.TextMatrix(.Row, GetCN("床号")) & "", 10, " ")
                          
        If lngAdviceID = 0 Then '---------------------------非历次检查直接用列表中记录填充
            gstrSQL = "Select 名称 From 部门表 Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人科室", CLng(.TextMatrix(.Row, GetCN("病人科室ID"))))
            lbl检查信息.Caption = "检查号:" & Rpad(.TextMatrix(.Row, GetCN("检查号")), 12, " ") & "病人科室:" & Rpad(rsTemp!名称, 11, " ") & _
                                  "开嘱医生:" & Rpad(.TextMatrix(.Row, GetCN("开嘱医生")), 8, " ") & "检查项目:" & .TextMatrix(.Row, GetCN("医嘱内容"))
            If .TextMatrix(.Row, GetCN("部位方法")) <> "" Then lbl检查信息.Caption = lbl检查信息.Caption & "(" & .TextMatrix(.Row, GetCN("部位方法")) & ")"
            
            mlngHSendNo = Nvl(.TextMatrix(.Row, GetCN("发送号")), 0)
            mstrHStudyUID = Nvl(.TextMatrix(.Row, GetCN("检查UID")))
            mlngExecuteStep = Decode(.TextMatrix(.Row, GetCN("检查过程")), "已审核", 6, "已完成", 5, 0)
            mblnHMoved = IIf(.TextMatrix(.Row, GetCN("转出")) = 1, True, False)
            
            lblCash.Caption = "收"
            lblCash.Visible = False
            lblCash.Visible = CheckChargeState(.TextMatrix(.Row, GetCN("医嘱ID")), CLng(Decode(.TextMatrix(.Row, GetCN("来源")), "门", 1, "住", 2, "外", 3, 4))) = 1
        Else
            Dim strSQLBak As String
            gstrSQL = "Select A.ID, A.病人科室id, A.开嘱医生,A.病人来源, A.医嘱内容, Nvl(A.婴儿, 0) 婴儿,A.病人id, " & vbNewLine & _
                        " A.主页id, A.挂号单, B.检查号, B.检查uid, C.名称, D.执行过程, D.发送号,D.执行状态,0 as 转出,A.执行科室ID " & vbNewLine & _
                        "From 病人医嘱记录 A, 影像检查记录 B, 部门表 C, 病人医嘱发送 D" & vbNewLine & _
                        "Where A.ID = [1] And A.ID = B.医嘱id And A.病人科室id = C.ID And A.ID = D.医嘱id"
            strSQLBak = gstrSQL
            strSQLBak = Replace(strSQLBak, "病人医嘱记录", "H病人医嘱记录")
            strSQLBak = Replace(strSQLBak, "病人医嘱发送", "H病人医嘱发送")
            strSQLBak = Replace(strSQLBak, "影像检查记录", "H影像检查记录")
            strSQLBak = Replace(strSQLBak, "0 as 转出", "1 as 转出")
            gstrSQL = gstrSQL & vbNewLine & " Union ALL " & strSQLBak
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查历次记录信息", lngAdviceID)
            If Not rsTemp.EOF Then

                mlngHSendNo = Nvl(rsTemp!发送号, 0)
                mstrHStudyUID = Nvl(rsTemp!检查uid)
                mlngExecuteStep = Nvl(rsTemp!执行过程, 0)
                mblnHMoved = IIf(rsTemp!转出 = 1, True, False)
                
                fraInfo.Tag = rsTemp!病人ID & "|" & rsTemp!主页ID & "|" & rsTemp!ID & "|" & rsTemp!发送号 _
                            & "|" & rsTemp!病人科室ID & "|" & rsTemp!挂号单 & "|" & Nvl(rsTemp!病人来源, 3) _
                            & "|" & rsTemp!检查uid & "|" & rsTemp!转出 & "|" & rsTemp!执行状态 & "|" & rsTemp!执行科室ID
                            
                lbl检查信息.Caption = "检查号:" & Rpad(Nvl(rsTemp!检查号), 12, " ") & "病人科室:" & Rpad(rsTemp!名称, 11, " ") & _
                                      "开嘱医生:" & Rpad(rsTemp!开嘱医生, 8, " ") & "检查项目:" & rsTemp!医嘱内容
                If rsTemp!婴儿 <> 0 Then
                    Dim lngBaby As Integer, lngPatID As Long, lngPageID As Long
                    lngBaby = rsTemp!婴儿: lngPatID = rsTemp!病人ID: lngPageID = Nvl(rsTemp!主页ID, 0)
                    gstrSQL = "Select Nvl(A.婴儿姓名, B.姓名 || '之子' || Trim(To_Char(A.序号, '9'))) As 婴儿姓名, 婴儿性别, 出生时间" & vbNewLine & _
                            "From 病人新生儿记录 A, 病人信息 B" & vbNewLine & _
                            "Where A.病人id = [1] And A.主页id = [2] And A.病人id = B.病人id And A.序号 = [3]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取婴儿信息", lngPatID, lngPageID, lngBaby)
                    If Not rsTemp.EOF Then
                        lbl个人信息.Caption = "姓  名:" & Rpad(rsTemp!婴儿姓名, 12, " ") & "性  别:" & Rpad(rsTemp!婴儿性别, 13, " ") & _
                                            "年  龄:" & Rpad(rsTemp!出生时间, 10, " ") & "标识号:" & Rpad(.TextMatrix(.Row, GetCN("标识号")), 12, " ") & _
                                            "床  号:" & Rpad(.TextMatrix(.Row, GetCN("床号")) & "", 10, " ")
                    End If
                End If
            Else
                lbl检查信息.Caption = "检查号:" & Space(12) & "病人科室:" & Space(11) & "开嘱医生:" & Space(8) & "检查项目:"
            End If
            lblCash.Caption = "历": lblCash.Visible = True
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillTxtAppend(Optional lngAdviceIDtmp As Long = 0)
'填充左下角医嘱附件
Dim lngAdviceID As Long, strAppend As String, rsTemp As ADODB.Recordset, i As Integer
    On Error GoTo errHandle
    With vsList
        If lngAdviceIDtmp = 0 Then
            lngAdviceID = Val(.TextMatrix(.Row, GetCN("医嘱ID")))
        Else
            lngAdviceID = lngAdviceIDtmp
        End If
        
        If lngAdviceIDtmp = 0 Then '-------------------------------------------列表选择调用
            If .TextMatrix(.Row, GetCN("部位方法")) <> "" Then
                For i = 0 To UBound(Split(.TextMatrix(.Row, GetCN("部位方法")), "),"))
                    If i = 0 Then
                        txtAppend = "检查部位:" & vbCrLf & Space(2) & "1:" & Split(.TextMatrix(.Row, GetCN("部位方法")), "),")(i) & ")"
                    Else
                        txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(.TextMatrix(.Row, GetCN("部位方法")), "),")(i) & ")"
                    End If
                Next
                If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) '取掉最后的括号
            Else
                txtAppend = "检查部位:" & .TextMatrix(.Row, GetCN("医嘱内容"))
            End If
            gstrSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列"
            If .TextMatrix(.Row, GetCN("转出")) = 1 Then gstrSQL = Replace(gstrSQL, "病人医嘱附件", "H病人医嘱附件")
        Else                    '-------------------------------------------历次记录选择调用
            Dim strTemp As String
            txtAppend = ""
            strTemp = Mid(lbl检查信息.Caption, InStr(lbl检查信息.Caption, "检查项目:") + 5)
            If strTemp <> "" Then
                If InStr(strTemp, ":") > 0 Then
                    strTemp = Split(strTemp, ":")(1)
                    For i = 0 To UBound(Split(strTemp, "),"))
                        If i = 0 Then
                            txtAppend = "检查部位:" & vbCrLf & Space(2) & "1:" & Split(strTemp, "),")(i) & ")"
                        Else
                            txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(strTemp, "),")(i) & ")"
                        End If
                    Next
                    If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) '取掉最后的括号
                Else
                    txtAppend = strTemp
                End If
            End If
            gstrSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列" '根据历次记录是否转移判断查历史表
            If Split(fraInfo.Tag, "|")(8) = 1 Then gstrSQL = Replace(gstrSQL, "病人医嘱附件", "H病人医嘱附件")
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人附件", lngAdviceID)
        Do Until rsTemp.EOF
            strAppend = strAppend & rsTemp!项目 & ":" & Nvl(rsTemp!内容) & vbCrLf
            rsTemp.MoveNext
        Loop
        
        txtAppend = txtAppend & vbCrLf & vbCrLf & strAppend
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillHistory()
'填充历次检查记录
Dim rsTemp As ADODB.Recordset, strTemp As String
    On Error GoTo errHandle
    With vsList
        cboTimes.Tag = "" 'cbotime下拉时用到，用于区别是"增加项目"时触发还是"点击cbotimes"触发
        gstrSQL = "Select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
                   " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 C" & _
                   " Where A.病人id = [1] And A.相关id Is Null And B.医嘱ID=A.ID " & _
                   "" & IIf(.TextMatrix(.Row, GetCN("检查过程")) = "已拒绝", "", " And B.执行状态<>2 ") & _
                   " AND A.ID=C.医嘱ID"

        gstrSQL = gstrSQL & " And (A.执行科室id+0 =[2] or A.开嘱科室ID+0=[2])"


        '启用关联病人，才查询关联ID
        If mblnRelatingPatient = True And .TextMatrix(.Row, GetCN("关联ID")) <> 0 Then
            gstrSQL = gstrSQL & " union select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
                    " From 病人医嘱记录 A " & _
                    " Where A.id in (Select 医嘱ID from 影像检查记录 Where 关联ID =[4]) "
        End If

        strTemp = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
        strTemp = Replace(strTemp, "病人医嘱发送", "H病人医嘱发送")
        strTemp = Replace(strTemp, "影像检查记录", "H影像检查记录")
        gstrSQL = gstrSQL & vbNewLine & " Union ALL " & vbNewLine & strTemp
        gstrSQL = "Select * From (" & vbNewLine & gstrSQL & vbNewLine & ") Order By 开嘱时间 Asc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", CLng(.TextMatrix(.Row, GetCN("病人ID"))), _
                glngDeptId, "", CLng(.TextMatrix(.Row, GetCN("关联ID"))))

        cboTimes.Clear
        Do Until rsTemp.EOF
           cboTimes.AddItem "第" & rsTemp.AbsolutePosition & "次(" & Format(rsTemp!开嘱时间, "yyyy-mm-dd") & ")  " & Trim(rsTemp!医嘱内容)
           cboTimes.ItemData(cboTimes.NewIndex) = rsTemp!医嘱id
           If rsTemp!医嘱id = .TextMatrix(.Row, GetCN("医嘱ID")) Then cboTimes.ListIndex = cboTimes.NewIndex
           rsTemp.MoveNext
        Loop
        cboTimes.Tag = "完成"
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub RefreshTabWindow(Optional lngAdviceIDtmp As Long = 0, Optional blnClear As Boolean = False, Optional blnRefresh As Boolean = False)
'lngAdviceIDtmp历次记录时传入 , 其它传0, blnclear清空当前列表, blnRefresh强制刷新
'刷新当前页面,调用：列表选择，历次记录选择，子窗体选择
'历次记录时fraInfo.Tag = 0病人ID|1主页ID|2医嘱ID|3发送号|4病人科室ID|5挂号单|6病人来源|7检查UID|8转出|9执行状态
Dim lngAdviceID As Long, lngSendNO As Long, lngPatID As Long, lngPageID As Long, blnCanPrint As Boolean, blnIsInsidePatient As Boolean
Dim lngUnit As Long, lngPatDept As Long, strRegNo As String, intMoved As Boolean, intState As Integer, i As Integer, intPatientForm As Integer

    On Error GoTo errHandle
    If lngAdviceIDtmp = 0 Then '-----------------------列表选择调用
        If blnClear Then       '无记录时清空所有子窗体
            lngAdviceID = 0: lngSendNO = 0: lngPatID = 0: lngPageID = 0
            lngPatDept = 0: strRegNo = "": intMoved = 0: intState = 0: lngUnit = 0: blnCanPrint = False
        Else
            With vsList
                lngAdviceID = .TextMatrix(.Row, GetCN("医嘱ID")): lngSendNO = .TextMatrix(.Row, GetCN("发送号"))
                lngPatID = .TextMatrix(.Row, GetCN("病人ID")): lngPageID = Val(.TextMatrix(.Row, GetCN("主页ID")))
                lngPatDept = .TextMatrix(.Row, GetCN("病人科室ID")): strRegNo = .TextMatrix(.Row, GetCN("挂号单"))
                intMoved = .TextMatrix(.Row, GetCN("转出"))
                intState = IIf(.TextMatrix(.Row, GetCN("检查过程")) = "已拒绝", 2, IIf(.TextMatrix(.Row, GetCN("检查过程")) = "已完成", 1, 3))
                lngUnit = Val(.TextMatrix(.Row, GetCN("当前病区ID")))
'                blnCanPrint = IIf(mblnCanPrint, IIf(.Cell(flexcpData, .Row, GetCN("紧急")) = 1, .TextMatrix(.Row, GetCN("报告人")) <> "", .TextMatrix(.Row, GetCN("复核人")) <> ""), True)
                intPatientForm = Decode(.TextMatrix(.Row, GetCN("来源")), "门", 1, "住", 2, "外", 3, 4)
            End With
        End If
    Else                       '----------------------历次记录选择调用
        lngAdviceID = lngAdviceIDtmp: lngSendNO = Split(fraInfo.Tag, "|")(3)
        lngPatID = Split(fraInfo.Tag, "|")(0): lngPageID = Val(Split(fraInfo.Tag, "|")(1))
        lngPatDept = Split(fraInfo.Tag, "|")(4): strRegNo = Split(fraInfo.Tag, "|")(5)
        intMoved = Split(fraInfo.Tag, "|")(8): intState = Split(fraInfo.Tag, "|")(9)
        lngUnit = lngPatDept
'        blnCanPrint = True
        intPatientForm = Split(fraInfo.Tag, "|")(6)
    End If
    
    blnIsInsidePatient = (intPatientForm = 1) Or (intPatientForm = 2)
    
    mfrmPACSImg.zlRefresh lngAdviceID, lngSendNO, mstrPrivs, intMoved = 1, blnRefresh
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subTriggleRefreshTimer(blnEnable As Boolean)
    '启动或者关闭自动刷新的Timer
    If blnEnable = False Then
        TimerRefresh.Enabled = False
    Else
        TimerRefresh.Enabled = mlngRefreshInterval > 0
    End If
End Sub

Private Function GetDeptName(lngDeptID As Long, strDeptStrings As String) As String
'通过可用的科室串，读取指定科室ID的科室名称
    Dim strDepts() As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    strDepts = Split(strDeptStrings, "|")
    For i = 0 To UBound(strDepts)
        If Split(strDepts(i), "_")(0) = lngDeptID Then
            GetDeptName = Split(strDepts(i), "_")(1)
            Exit For
        End If
    Next i
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
