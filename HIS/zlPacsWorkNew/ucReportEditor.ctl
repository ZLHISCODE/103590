VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.UserControl ucReportEditor 
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   ScaleHeight     =   10815
   ScaleWidth      =   10380
   Begin VB.Timer timerTmp 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   7800
      Top             =   120
   End
   Begin VB.PictureBox picState 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   535
      Left            =   120
      ScaleHeight     =   540
      ScaleWidth      =   9975
      TabIndex        =   12
      Top             =   9840
      Width           =   9975
      Begin VB.Label labEditState 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   9600
         TabIndex        =   15
         Top             =   160
         Width           =   240
      End
      Begin VB.Label labFmt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   375
         TabIndex        =   29
         Top             =   0
         Width           =   9120
      End
      Begin VB.Label lab阳性 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "＋"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   0
         TabIndex        =   28
         ToolTipText     =   "阴阳性"
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lab危急 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   0
         TabIndex        =   27
         ToolTipText     =   "危急状态"
         Top             =   0
         Width           =   270
      End
      Begin VB.Label labSignTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "签名:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   375
         TabIndex        =   14
         Top             =   240
         Width           =   840
      End
      Begin VB.Label labSign 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.PictureBox picChar 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   5295
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5295
      Begin XtremeCommandBars.CommandBars cbrChar 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   120
      ScaleHeight     =   9015
      ScaleWidth      =   9975
      TabIndex        =   4
      Top             =   720
      Width           =   9975
      Begin VB.PictureBox picImageBack 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   9735
         TabIndex        =   8
         Top             =   480
         Width           =   9735
         Begin zl9PacsControl.ucSplitter ucSplitter1 
            Bindings        =   "ucReportEditor.ctx":0000
            Height          =   2895
            Left            =   5625
            TabIndex        =   9
            Top             =   0
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   5106
            SplitLevel      =   3
            Con1MinSize     =   2000
            Con2MinSize     =   1000
            Control1Name    =   "dcmReportImg"
            Control2Name    =   "dcmMarkImage"
         End
         Begin VB.PictureBox picMarkImgOper 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   5880
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   21
            Top             =   120
            Visible         =   0   'False
            Width           =   1815
            Begin VB.CommandButton cmdOper 
               BackColor       =   &H0080FF80&
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   7
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "向后移动报告图像"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               BackColor       =   &H00FFFFFF&
               Caption         =   "AU"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "删除报告图像"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               BackColor       =   &H0080C0FF&
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "向前移动报告图像"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               BackColor       =   &H00FF80FF&
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   720
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "向后移动报告图像"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               BackColor       =   &H008080FF&
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   6
               Left            =   1080
               Style           =   1  'Graphical
               TabIndex        =   25
               ToolTipText     =   "向后移动报告图像"
               Top             =   0
               Width           =   375
            End
         End
         Begin VB.PictureBox picReportImgOper 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   1095
            TabIndex        =   17
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
            Begin VB.CommandButton cmdOper 
               Height          =   375
               Index           =   2
               Left            =   720
               Picture         =   "ucReportEditor.ctx":0014
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "向后移动报告图像"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               Height          =   375
               Index           =   1
               Left            =   360
               Picture         =   "ucReportEditor.ctx":0716
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "向前移动报告图像"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   0
               Picture         =   "ucReportEditor.ctx":0E18
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "删除报告图像"
               Top             =   0
               Width           =   375
            End
         End
         Begin DicomObjects.DicomViewer dcmMarkImage 
            Height          =   2895
            Left            =   5760
            TabIndex        =   10
            Top             =   0
            Width           =   3975
            _Version        =   262147
            _ExtentX        =   7011
            _ExtentY        =   5106
            _StockProps     =   35
            BackColor       =   4210752
            CellSpacing     =   2
         End
         Begin DicomObjects.DicomViewer dcmReportImg 
            Height          =   2895
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   5625
            _Version        =   262147
            _ExtentX        =   9922
            _ExtentY        =   5106
            _StockProps     =   35
            BackColor       =   4210752
            CellSpacing     =   2
         End
      End
      Begin VB.PictureBox picDesc 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   9735
         TabIndex        =   7
         Top             =   3720
         Visible         =   0   'False
         Width           =   9735
         Begin RichTextLib.RichTextBox rtb所见 
            Height          =   1695
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   2990
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"ucReportEditor.ctx":115A
         End
      End
      Begin VB.PictureBox picOpin 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   9735
         TabIndex        =   6
         Top             =   5760
         Visible         =   0   'False
         Width           =   9735
         Begin RichTextLib.RichTextBox rtb意见 
            Height          =   1575
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   2778
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"ucReportEditor.ctx":11F7
         End
      End
      Begin VB.PictureBox picAdvi 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   9735
         TabIndex        =   5
         Top             =   7680
         Visible         =   0   'False
         Width           =   9735
         Begin RichTextLib.RichTextBox rtb建议 
            Height          =   975
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   1720
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"ucReportEditor.ctx":1294
         End
      End
      Begin XtremeDockingPane.DockingPane dkpMain 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   450
         _ExtentY        =   423
         _StockProps     =   0
      End
   End
   Begin MSComctlLib.ImageList listCur 
      Left            =   240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucReportEditor.ctx":1331
            Key             =   "pen"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtxtSaveElement 
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"ucReportEditor.ctx":200B
   End
   Begin VB.Menu menuReport 
      Caption         =   "报告图"
      Begin VB.Menu menuReport_Del 
         Caption         =   "删除(&D)"
      End
      Begin VB.Menu menuReport_Split 
         Caption         =   "-"
      End
      Begin VB.Menu menuReport_Last 
         Caption         =   "前移(&L)"
      End
      Begin VB.Menu menuReport_Next 
         Caption         =   "后移(&N)"
      End
   End
   Begin VB.Menu menuLab 
      Caption         =   "标注"
      Visible         =   0   'False
      Begin VB.Menu menuLab_Del 
         Caption         =   "删除(&D)"
      End
   End
End
Attribute VB_Name = "ucReportEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const Report_Element_报告签名 = "报告签名"

'签名状态
Private Enum EPRSignLevelEnum
    cprSL_空白 = 0              '未签名
    cprSL_经治 = 1              '经治医师签名
    cprSL_主治 = 2              '主治医师签名
    cprSL_主任 = 3              '主任医师签名
    cprSL_正高 = 4              '正高：签名级别不包含，只表示人员居右正高职称，以便区别副主任医师
End Enum

Private Type TReportInfo
    创建日期 As Date      '创建日期
    创建用户 As String    '创建人
    审核用户 As String    '审核人
    签名级别 As EPRSignLevelEnum
    最后版本 As Long
    目标版本 As Long
    
End Type


Private Enum TReportFmtFrom
    rffTemplate = 0 '来自模板
    rffSample = 1   '来自范文
    rffReport = 2   '来自报告
End Enum


Private Type paneInfo
    title As String
    ID As Long
    hwnd As Long
    hidden As Boolean
    iconid As Long
    options As PaneOptions
    tag As Long
End Type


Private mlngModule As Long
Private mlngDeptID As Long
Private mObjNotify As IEventNotify


Private mlngAdviceId As Long        '医嘱ID
Private mlngFileID As Long          '格式文件ID
Private mstrEprFmtName As String    '报告格式名称
Private mstrPrintFmts As String     '报告打印格式
Private mlngSampleId As Long        '范文ID
Private mlngReportID As Long        '报告ID
Private mblnIsMoved As Boolean      '是否转储
Private mblnIsLoadData As Boolean   '是否已经载入了数据
Private mstrReportImgPath As String '报告图路径
Private mftpConTag As TFtpConTag    'ftp连接标记

Private mintEditFontSize As Integer '编辑框字体大小
Private mrtbActive As RichTextBox   '当前编辑框
Private mlngSelReportImgIndex As Long   '选择的报告图索引
Private mblnIsInit As Boolean       '是否初始化
Private mstrPrivs As String         '模块权限

Private mobjSpePlugin As Object     '专科报告插件
Private mblnIsSpeState As Boolean   '是否专科报告编辑模式状态


Private mblnIsLockingEdit As Boolean   '是否锁定编辑中
Private mlngSignCount As Long           '签名数量
Private mlngSignLevel As TReportSignLevel   '签名级别
Private mstrFirstSignUser As String     '首次签名用户
Private mstrFinalSignUser As String     '最终签名用户
Private mintTargetVer As Integer        '目标版本
Private mintSourceVer As Integer
Private mstrCreateUser As String        '创建人
Private mstrSaveUser As String          '最后保存人
Private mlngCreateDeptId As Long        '创建科室ID

'需要从参数配置中读取
Private mblnTechReptSame As Boolean '只能填写自己检查的报告
Private mlngSignPassType As Long        '签名类型 '密码验证规则（系统参数） 0-密码；1－数字；2－两者皆可

Private mblnUseImgSign As Boolean   '是否使用图像签名
Private mblnVisibleSpecialty As Boolean '是否显示专科报告
Private mblnCheckPrintPara As Boolean   '平诊需要审核才能打印
 
Private mblnReportWithResult As Boolean '无影像诊断为阴性
Private mblnReportDefaultPositive As Boolean '结果默认阳性
Private mblnIgnoreResult As Boolean     '忽略结果阴阳性
Private mblnIsEditWithReportImage As Boolean    '有图像才能写报告

Private mstrDescTitle As String         '所见标题
Private mstrOpinTitle As String         '意见标题
Private mstrAdviTitle As String         '建议标题


Private mblnIsEditable As Boolean   '控制内容是否能够编辑
Private mblnIsReadOnly As Boolean   '控制非内容相关的功能，如审核，回退，预览打印等，在没有对应权限时会处于true状态
Private mblnIsComplete As Boolean   '是否完成，在readonly状态时，检查不一定为完成状态，可以进行报告删除相关操作

Private mblnIsModifyText As Boolean
Private mblnIsModifyImage As Boolean
Private mblnIsModifyMarks As Boolean
 
Private mlngMarkType As TImgMarkType
Private mstrMarkText As String

Private WithEvents mobjMarkProcessV2 As frmImageProcessV2
Attribute mobjMarkProcessV2.VB_VarHelpID = -1

Public Event OnOutlineChange(ByVal lngSelOutline As TOutlineType)
Public Event OnStateChange()
Public Event OnDelRepImg(ByVal strImgKey As String) '删除报告图事件


'当前句柄
Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'医嘱ID
Property Get AdviceId() As Long
    AdviceId = mlngAdviceId
End Property

'报告ID
Property Get ReportID() As Long
    ReportID = mlngReportID
End Property

'是否转储
Property Get IsMoved() As Boolean
    IsMoved = mblnIsMoved
End Property

'是否有专科报告
Property Get HasSpeReport() As Boolean
    HasSpeReport = IIf(mobjSpePlugin Is Nothing, False, True)
End Property

'是否专科报告状态
Property Get IsSpeState() As Boolean
    IsSpeState = mblnIsSpeState
End Property

Property Let IsSpeState(ByVal value As Boolean)
    Call ChangeSepState(value, False)
End Property

'范文ID
Property Get SampleId() As Long
    SampleId = mlngSampleId
End Property

'病历单据格式名称
Property Get EPRFmtName() As String
    EPRFmtName = mstrEprFmtName
End Property

'已签名版本
Property Get SourceVer() As Long
    SourceVer = mintSourceVer
End Property


'目标版本
Property Get TargetVer() As Integer
    TargetVer = mintTargetVer
End Property
 

'创建人
Property Get CreateUser() As String
    CreateUser = mstrCreateUser
End Property
 

'保存人
Property Get SaveUser() As String
    SaveUser = mstrSaveUser
End Property
 
'创建科室ID
Property Get CreateDeptId() As Long
    CreateDeptId = mlngCreateDeptId
End Property
 
'权限串
Property Get Privs() As String
    Privs = mstrPrivs
End Property

Property Let Privs(ByVal value As String)
    mstrPrivs = value
End Property


'编辑字体大小
Property Get EditFontSize() As Integer
    EditFontSize = mintEditFontSize
End Property

Property Let EditFontSize(ByVal value As Integer)
    mintEditFontSize = value
    
    If mintEditFontSize <> 0 Then
        Call SetContextFont(mintEditFontSize)
    Else
        Call SetContextFont(gbytFontSize + 3)
    End If
End Property

Public Sub SetContextFont(ByVal intFontSize As Integer)
    rtb所见.Font.Size = intFontSize
    rtb意见.Font.Size = intFontSize
    rtb建议.Font.Size = intFontSize

    rtb所见.SelFontSize = intFontSize
    rtb意见.SelFontSize = intFontSize
    rtb建议.SelFontSize = intFontSize
    
End Sub


'签名数量
Property Get SignCount() As Long
    SignCount = Val(labSign.tag)
End Property

'签名类型
Property Get SignPassType() As Long
    SignPassType = mlngSignPassType
End Property

Property Let SignPassType(ByVal value As Long)
    mlngSignPassType = value
End Property


'是否完成
Property Get IsComplete() As Boolean
    IsComplete = mblnIsComplete
End Property

'只读属性
Property Get IsReadOnly() As Boolean
    IsReadOnly = mblnIsReadOnly
End Property

'Property Let IsReadOnly(ByVal value As Boolean)
'    mblnIsReadOnly = value
'End Property

'可编辑属性
Property Get IsEditable() As Boolean
    IsEditable = mblnIsEditable
End Property

Property Let IsEditable(ByVal value As Boolean)
    mblnIsEditable = value
End Property

Property Get IsModify() As Boolean
'判断报告是否有修改
    IsModify = mblnIsModifyText Or mblnIsModifyImage Or mblnIsModifyMarks
    
    If Not mobjSpePlugin Is Nothing Then
        IsModify = IsModify Or mobjSpePlugin.pModified
    End If
End Property


'是否对文本内容更新修改
Property Get IsModifyText() As Boolean
    IsModifyText = mblnIsModifyText
End Property

'是否对报告图像更新修改
Property Get IsModifyImage() As Boolean
    IsModifyImage = mblnIsModifyImage
End Property

'图像标记是否被修改
Property Get IsModifyMarks() As Boolean
    IsModifyMarks = mblnIsModifyMarks
End Property


'专科
Property Get VisibleSpecialty() As Boolean
    VisibleSpecialty = mblnVisibleSpecialty
End Property

Property Let VisibleSpecialty(ByVal value As Boolean)
    mblnVisibleSpecialty = value
End Property





'检查所见标题-----------------------
Property Get DescTitle() As String
    DescTitle = mstrDescTitle
End Property

Property Let DescTitle(ByVal value As String)
    mstrDescTitle = value
End Property

'诊断建议标题-----------------------
Property Get AdviTitle() As String
    AdviTitle = mstrAdviTitle
End Property

Property Let AdviTitle(ByVal value As String)
    mstrAdviTitle = value
End Property

'诊断意见标题-----------------------
Property Get OpinTitle() As String
    OpinTitle = mstrOpinTitle
End Property

Property Let OpinTitle(ByVal value As String)
    mstrOpinTitle = value
End Property


Property Get DescContext() As String
'检查所见内容
    DescContext = rtb所见.Text
End Property

Property Get OpinContext() As String
'诊断意见内容
    OpinContext = rtb意见.Text
End Property

Property Get AdviContext() As String
'建议内容
    AdviContext = rtb建议.Text
End Property


'报告图
Property Get RepImageCount() As Long
    RepImageCount = dcmReportImg.Images.Count
End Property


Property Get RepImage(ByVal lngIndex As Long) As Object
On Error GoTo errhandle
    Set RepImage = dcmReportImg.Images(lngIndex)
Exit Sub
errhandle:
    Set RepImage = Nothing
End Property

'标记图
Property Get MarkImageCount() As Long
    MarkImageCount = dcmMarkImage.Images.Count
End Property

Property Get MarkImage() As Object
On Error GoTo errhandle
    Set MarkImage = dcmMarkImage.Images(0)
Exit Property
errhandle:
    Set MarkImage = Nothing
End Property


Property Get CurOutlineType() As TOutlineType
    CurOutlineType = otNone
    
    If mrtbActive Is Nothing Then Exit Property
    
    If mrtbActive Is rtb所见 Then
        CurOutlineType = otDesc
    End If
    
    If mrtbActive Is rtb意见 Then
        CurOutlineType = otOpin
    End If
    
    If mrtbActive Is rtb建议 Then
        CurOutlineType = otAdvi
    End If
End Property

Public Sub SetFontSize(ByVal intFontSize As Integer)
    Dim objCapFont As New StdFont
    
    FontSize = intFontSize
    
    objCapFont.Name = FontName
    objCapFont.Size = intFontSize + 3
  
    Set dkpMain.PaintManager.CaptionFont = objCapFont
        
    Set cbrChar.options.Font = objCapFont
    
    picChar.FontSize = FontSize
    
    dkpMain.RecalcLayout
End Sub


Public Sub ChangeSepState(ByVal blnState As Boolean, ByVal blnIsForceRefresh As Boolean)
    Dim i As Long
    Dim objPane As Pane
    Dim Left As Long, Right As Long
    Dim Top As Long, Bottom As Long
    Dim strErr As String
    
    mblnIsSpeState = False
    
    If mobjSpePlugin Is Nothing Then Exit Sub
    
    Call dkpMain.GetClientRect(Left, Top, Right, Bottom)
    
    If blnState Then
        If dkpMain.PanesCount < 5 Then
            '专科报告录入
            If dkpMain.Panes(1).Closed = False Then
                Set objPane = dkpMain.CreatePane(5, 0, 1000 - (picImageBack.Height / (Height - 3000)) * 1000, DockBottomOf, dkpMain.Panes(1))
            Else
                Set objPane = dkpMain.CreatePane(5, 0, 1000 - (picImageBack.Height / (Height - 3000)) * 1000, DockBottomOf)
            End If
            
            objPane.title = "专科录入"
            objPane.Handle = mobjSpePlugin.hwnd
            objPane.tag = 4
            objPane.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            objPane.Closed = True
        End If
    End If
    
    For i = 1 To dkpMain.PanesCount
        If dkpMain.Panes(i).tag <> 0 And dkpMain.Panes(i).tag <> 4 Then
             '标题有效才进行显示
             dkpMain.Panes(i).Closed = IIf(blnState, True, dkpMain.Panes(i).iconid = 0)
            
        ElseIf dkpMain.Panes(i).tag = 4 Then
            dkpMain.Panes(i).Closed = IIf(blnState, False, True)
            
            mblnIsSpeState = blnState
            
        End If
    Next
    
    If mblnIsSpeState Then
        
        If dkpMain.Panes(5).ID = mlngAdviceId And blnIsForceRefresh = False Then
            picChar.Visible = False
            Exit Sub
        End If
        
        On Error GoTo errhandle
            mobjSpePlugin.Refresh mlngAdviceId, mlngReportID, mblnIsEditable And Not mblnIsReadOnly, mblnIsMoved
errhandle:
        strErr = err.Description
        If err.Number <> 0 Then MsgboxH GetRootHwnd, "专科报告插件刷新错误:" & strErr, vbOKOnly, "提示"
        
        dkpMain.Panes(5).ID = mlngAdviceId
         
    End If
    
    picChar.Visible = False
End Sub


Private Sub cbrChar_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strErr As String
On Error GoTo errhandle
    Dim objWordCharCfg As frmWordCharCfgV2
    
    Select Case Control.ID
        Case 1  '配置常用词句
            Set objWordCharCfg = New frmWordCharCfgV2
            If objWordCharCfg.zlShowWordCharCfg(mlngModule, mObjNotify.Owner) Then
                Call InitReportChar
                
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_REFWCHR, , Parent.hwnd, glngSys, mlngModule)
            End If
            
        Case Else   '写入选择的词句
            If mblnIsEditable = False Then Exit Sub
            
            If mrtbActive Is Nothing Then Exit Sub
            mrtbActive.SelText = Control.Caption
    End Select
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

'Private Sub chkCritical_Click()
'On Error GoTo errhandle
'
'    If chkCritical.value = 0 Then
'        chkCritical.ForeColor = &H404040
'    Else
'        chkCritical.ForeColor = ColorConstants.vbRed
'    End If
'
'    If mblnIsLoadData = False Or mblnIsEditable = False Then Exit Sub
'
'    mblnIsModifyText = True
'Exit Sub
'errhandle:
'    Debug.Print "chkPositive_Click:" & err.Description
'End Sub
'
'Private Sub chkPositive_Click()
'On Error GoTo errhandle
'    If chkPositive.value = 0 Then
'        chkPositive.ForeColor = &H404040
'    Else
'        chkPositive.ForeColor = ColorConstants.vbRed
'    End If
'
'    If mblnIsLoadData = False Or mblnIsEditable = False Then Exit Sub
'
'    mblnIsModifyText = True
'Exit Sub
'errhandle:
'    Debug.Print "chkPositive_Click:" & err.Description
'End Sub


Public Sub Init(objNotify As IEventNotify, ByVal lngModuleNo As Long, ByVal lngDeptId As Long, _
    ByVal strPrivs As String, lngSignPassType As Long, Optional ByVal blnIsForce As Boolean = False)
'模块初始化
    mlngModule = lngModuleNo
    mlngDeptID = lngDeptId
    mlngSignPassType = lngSignPassType
    
    Set mObjNotify = objNotify
    
    mstrPrivs = strPrivs
    
    If mblnIsInit And blnIsForce = False Then Exit Sub
    
    Call InitPar
    
    Call InitReportChar
    
    Call Relayout
    
    mblnIsInit = True
End Sub


Public Sub InitReportChar()
    Dim cbrToolBar As CommandBar
    Dim strWord As String
    Dim aryWord() As String
    Dim i As Long
    Dim blnIsSetGroup As Boolean
    Dim lngWordLen As Long
    
    cbrChar.DeleteAll
    
    
    With cbrChar.options
        .UpdatePeriod = 800
        
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseDisabledIcons = False
        .LargeIcons = False
    End With
    
    Set cbrToolBar = cbrChar.Add("特殊字符", xtpBarTop)
     
    With cbrToolBar
        .Position = xtpBarTop
        .Customizable = False
        .ShowTextBelowIcons = True
        .Closeable = False
        .EnableDocking xtpFlagHideWrap
    End With
    
    strWord = zlDatabase.GetPara("报告常用词句", glngSys, mlngModule)
    aryWord = Split(strWord & vbCrLf, vbCrLf)
    
    blnIsSetGroup = False
    With cbrToolBar.Controls
        .Add(xtpControlButton, 1, "…").ToolTipText = "常用字符配置"
        
        For i = 0 To UBound(aryWord)
            If Len(aryWord(i)) > 0 Then
                lngWordLen = TextWidth(aryWord(i))
                
                If lngWordLen <= picChar.Width - 2000 Then
                    If blnIsSetGroup = False Then
                        .Add(xtpControlButton, i + 2, aryWord(i)).BeginGroup = True
                        blnIsSetGroup = True
                    Else
                        .Add xtpControlButton, i + 2, aryWord(i)
                    End If
                    
                    aryWord(i) = ""
                End If
            End If
        Next
        
        For i = 0 To UBound(aryWord)
            If Len(aryWord(i)) > 0 Then
                If blnIsSetGroup = False Then
                    .Add(xtpControlButton, i + 2, aryWord(i)).BeginGroup = True
                    blnIsSetGroup = True
                Else
                    .Add xtpControlButton, i + 2, aryWord(i)
                End If
            End If
        Next
    End With
End Sub


Private Sub InitPar()
'初始化参数
    mblnIgnoreResult = Val(GetDeptPara(mlngDeptID, "忽略结果阴阳性", 0)) <> 0 '        '忽略结果阴阳性
    mblnReportDefaultPositive = Val(GetDeptPara(mlngDeptID, "诊断结果默认阳性", 0)) <> 0
    mblnTechReptSame = Val(GetDeptPara(mlngDeptID, "只能填写自己检查的报告", 0)) <> 0
    mblnReportWithResult = Val(GetDeptPara(mlngDeptID, "无影像诊断为阴性", 0)) <> 0 '  '无影像诊断为阴性
    mblnVisibleSpecialty = Val(GetDeptPara(mlngDeptID, "显示专科报告", 0)) <> 0
    mblnUseImgSign = Val(GetDeptPara(mlngDeptID, "图像签名验证")) <> 0
    mblnIsEditWithReportImage = Val(GetDeptPara(mlngDeptID, "有图像才能写报告", 0)) <> 0
    
    mstrDescTitle = GetDeptPara(mlngDeptID, "检查所见名称", "检查所见")
    mstrOpinTitle = GetDeptPara(mlngDeptID, "诊断意见名称", "诊断意见")
    mstrAdviTitle = GetDeptPara(mlngDeptID, "建议名称", "诊断建议")
End Sub


Public Sub Refresh(ByVal lngAdviceId As Long, ByVal lngFileId As Long, ByVal lngSampleId As Long, ByVal lngReportID As Long, _
    Optional ByVal blnIsMoved As Boolean = False, Optional ByVal blnIsForce As Boolean = False)
    Dim lngSelStart As Long
    
    '如果数据相同，且非强制刷新，则直接退出
    If lngFileId = mlngFileID _
        And lngSampleId = mlngSampleId _
        And lngReportID = mlngReportID _
        And Not blnIsForce Then Exit Sub
    
    mblnIsLoadData = False
    
    picReportImgOper.Visible = False
    picMarkImgOper.Visible = False
    
    '如果医嘱ID和现有不同，说明是不同检查的报告，则报告编辑焦点不需要进行保留
    If lngAdviceId <> mlngAdviceId Then Set mrtbActive = Nothing
    
    mlngAdviceId = lngAdviceId
    mblnIsMoved = blnIsMoved
    mlngFileID = lngFileId
    mlngSampleId = lngSampleId
    mlngReportID = lngReportID ' 0 '
    mlngSelReportImgIndex = 0
    
    mblnIsModifyMarks = False
    mblnIsModifyImage = False
    mblnIsModifyText = False
    
    mblnIsEditable = False
    
    mlngMarkType = imtAuto ' imtNormal
    
    mftpConTag.Ip = ""
    
    mblnIsLockingEdit = False
    
    
    '如果界面没有加载，则不进行显示
'    If Extender.Visible = False Then Exit Sub     'And Not blnIsForce 需要考虑预览与打印

    '保存之前的光标所在文本框位置
    If Not mrtbActive Is Nothing Then lngSelStart = mrtbActive.SelStart
    
    Call ResetContext
    
    If mintEditFontSize <> 0 Then
        Call SetContextFont(mintEditFontSize)
    Else
        Call SetContextFont(gbytFontSize)
    End If
    
    mstrReportImgPath = GetReportImgPath(lngAdviceId, blnIsMoved)
    
    '载入报告
    Call LoadReport
    
    '恢复文本框的光标位置
    If Not mrtbActive Is Nothing And lngSelStart > 0 Then mrtbActive.SelStart = lngSelStart
    
    '载入专科报告
    If Not mobjSpePlugin Is Nothing Then
        
        If mblnIsSpeState Then
            '恢复到专科显示界面
            Call ChangeSepState(True, blnIsForce)
        Else
            '如果是强制刷新，则重新对id进行设置，以便后续切换到专科报告时能够进行刷新操作
            If blnIsForce And Not dkpMain.Panes(5) Is Nothing Then dkpMain.Panes(5).ID = -5
            
        End If
    End If
    
    Call ShowPrintFormat(mstrPrintFmts)
    
    mblnIsLoadData = True
End Sub


Public Sub ResetContext()
    mlngCreateDeptId = mlngDeptID ' 0
    
    mstrCreateUser = UserInfo.姓名 ' ""
    mstrSaveUser = UserInfo.姓名 ' ""
     
    mlngSignCount = 0
    mlngSignLevel = cprSL_空白
    mstrFirstSignUser = ""
    mstrFinalSignUser = ""
    mintTargetVer = 1
    mintSourceVer = 0
    
    labEditState.Caption = ""
    
    dcmMarkImage.Images.Clear
    dcmReportImg.Images.Clear
    picChar.Visible = False
    
    rtb所见.Text = ""
    rtb意见.Text = ""
    rtb建议.Text = ""
    
    labSign.Caption = ""
    labSign.tag = ""
    
'    chkPositive.value = 0
'    chkCritical.value = 0
    
    mblnIsModifyImage = False
    mblnIsModifyMarks = False
    mblnIsModifyText = False
    
'    If mblnIgnoreResult = False Then
'        '如果不忽略阴阳性，则设置阴阳性的默认值
'        chkPositive.value = Abs(CLng(mblnReportDefaultPositive))
'    End If
    
'    mblnTechReptSame = False
End Sub

Private Sub LoadReport()
'载入报告
    Dim strSQL As String
    Dim strPicSql As String
    Dim strContextSql As String
    Dim rsData As ADODB.Recordset
    Dim lngFileId As Long
    Dim lngDataFrom As TReportFmtFrom
    Dim strTmp As String
    Dim blnHas描述 As Boolean
    Dim blnHas意见 As Boolean
    Dim blnHas建议 As Boolean
    Dim strTitle As String
    Dim blnForceRead As Boolean
    Dim blnReportVisible As Boolean
    Dim blnMarkVisible As Boolean
    Dim i As Long
    Dim blnReadyRepImg As Boolean
    Dim strFile As String
    
    '不是相同报告时，才设置为nothing
'    Set mrtbActive = Nothing
    
    mstrEprFmtName = ""
    If mlngFileID <> 0 Then
        strSQL = "Select 名称 From 病历文件列表 where ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询病历名称", mlngFileID)
        
        If rsData.RecordCount > 0 Then mstrEprFmtName = nvl(rsData!名称)
    End If
    
    '报告图查询...
    If mlngReportID <> 0 Then
        lngFileId = mlngReportID
        lngDataFrom = rffReport
        
        '从电子病历内容中查询数据
        strSQL = "Select  Id As 表格Id From 电子病历内容" & _
                    " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2' " & _
                    " Order By 对象序号"
                    
        strPicSql = "select ID,文件ID,父ID,开始版,对象标记,对象属性,内容行次 from 电子病历内容 where  文件ID=[1] and 父ID=[2] and 对象类型=5 order by 对象标记"
        
        strContextSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文" & vbNewLine & _
                 " From 电子病历内容 a,电子病历内容 b " & _
                 " Where a.文件id = [1] And (a.对象类型 = 3) And a.Id = b.父ID And b.对象类型 = 2 And b.终止版 = 0"
        '(a.对象类型 = 3 or a.对象类型=1 ) order by a.对象序号 '支持不使用1*1的表格
        
        If mblnIsMoved Then
            strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
            strPicSql = Replace(strPicSql, "电子病历内容", "H电子病历内容")
            strContextSql = Replace(strContextSql, "电子病历内容", "H电子病历内容")
        End If
        
    Else
        If mlngSampleId <> 0 Then
            lngDataFrom = rffSample
            lngFileId = mlngSampleId
            
            '从范文中查询格式数据
            strSQL = "Select  Id As 表格Id From 病历范文内容 a " & _
                        " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2' " & _
                        " Order By 对象序号"
                        
            strPicSql = "select ID,文件ID,父ID,1 as 开始版,对象标记,对象属性,内容行次 from 病历范文内容 where  文件ID=[1] and 父ID=[2] and 对象类型=5 order by 对象标记"
            
            strContextSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文" & vbNewLine & _
                    " From 病历范文内容 a, 病历范文内容 b" & vbNewLine & _
                    " Where a.文件id = [1] And (a.对象类型 = 3 ) And a.Id = b.父id And b.对象类型 = 2"
            '(a.对象类型 = 3 or a.对象类型=1 ) order by a.对象序号 '支持不使用1*1的表格
        Else
            lngDataFrom = rffTemplate
            lngFileId = mlngFileID
            
            '从病历单据中查询格式数据
            strSQL = "Select  Id As 表格Id From 病历文件结构" & _
                        " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2' " & _
                        " Order By 对象序号"
                        
            strPicSql = "select ID,文件ID,父ID,1 as 开始版,对象标记,对象属性,内容行次 from 病历文件结构 where  文件ID=[1] and 父ID=[2] and 对象类型=5 order by 对象标记"
            
            strContextSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文 " & _
                     " From 病历文件结构 a, 病历文件结构 b" & _
                     " Where a.文件id = [1] And (a.对象类型 = 3 ) And a.Id = b.父id And b.对象类型 = 2 "
                     
            '(a.对象类型 = 3 or a.对象类型=1 ) order by a.对象序号 '支持不使用1*1的表格
        End If
    End If
    
    '读取报告图信息****************************************
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告图框", lngFileId)
    
    dcmReportImg.MultiColumns = 1
    dcmReportImg.MultiRows = 1
    
    dcmReportImg.Visible = False
    dcmMarkImage.Visible = False
    
    blnReportVisible = False
    blnMarkVisible = False
    
    dcmReportImg.Images.Clear
    dcmMarkImage.Images.Clear

    If rsData.RecordCount > 0 Then
        '读取标记图，报告图
        blnReportVisible = True
        dcmReportImg.Visible = True
        '图像对象查询
        dcmReportImg.tag = Val(nvl(rsData!表格ID))
        
        Set rsData = zlDatabase.OpenSQLRecord(strPicSql, "查询报告图片", lngFileId, Val(nvl(rsData!表格ID)))
        If rsData.RecordCount > 0 Then
            
            Call ParshReportImgData(rsData, lngDataFrom)
            
            If dcmMarkImage.Images.Count > 0 Then blnMarkVisible = True
        End If
        
        '读取预先设置的报告图
        '只有成功下载到本地的图像，才能自动添加的报告图中
        strPicSql = "select 图像UID from 影像检查图象 a, 影像检查序列 b, 影像检查记录 c where  a.序列UID=b.序列UID and b.检查UID=c.检查UID and c.医嘱ID=[1] and a.报告图>=0 order by 图像时间"
        If mblnIsMoved Then
            strPicSql = Replace(strPicSql, "影像检查图象", "H影像检查图象")
            strPicSql = Replace(strPicSql, "影像检查序列", "H影像检查序列")
            strPicSql = Replace(strPicSql, "影像检查记录", "H影像检查记录")
        End If
        
        Set rsData = zlDatabase.OpenSQLRecord(strPicSql, "查询预设报告图", mlngAdviceId)
        If rsData.RecordCount > 0 Then
            While Not rsData.EOF
                blnReadyRepImg = True
                For i = 1 To dcmReportImg.Images.Count
                    If nvl(rsData!图像UID) = dcmReportImg.Images(i).InstanceUID Then
                        blnReadyRepImg = False
                        Exit For
                    End If
                Next
                
                If blnReadyRepImg Then
                    strFile = FormatFilePath(mstrReportImgPath & "\" & nvl(rsData!图像UID))
                    
                    If FileExists(strFile) Then
                        '存在对应的图像
                        Call AddRepImgFile(strFile, , , True)
                    End If
                End If
                
                Call rsData.MoveNext
            Wend
        End If
    End If
    
    
    If blnReportVisible = False And blnMarkVisible = False Then
        '关闭报告图
        dkpMain.Panes(1).Closed = True
    Else
        '打开报告图
        dkpMain.Panes(1).Closed = False
        
        If blnMarkVisible = False Then
            dcmReportImg.Width = picImageBack.Width
            ucSplitter1.Visible = False
        Else
            If dcmReportImg.Width = picImageBack.Width Then
                If picImageBack.Width - dcmMarkImage.Width < 0 Then
                    dcmMarkImage.Width = 0.34 * picImageBack.Width
                End If
                
                dcmReportImg.Width = picImageBack.Width - dcmMarkImage.Width
                ucSplitter1.Left = dcmReportImg.Width
                ucSplitter1.RePaint
            End If
            
            ucSplitter1.Visible = True
        End If
    End If

    '读取报告文本内容****************************************
    Set rsData = zlDatabase.OpenSQLRecord(strContextSql, "查询报告文本", lngFileId)
    
    rtb所见.Text = ""
    rtb意见.Text = ""
    rtb建议.Text = ""
    
    blnHas描述 = False
    blnHas意见 = False
    blnHas建议 = False
    
    While rsData.EOF = False
        strTmp = nvl(rsData!对象属性)
        strTitle = nvl(rsData!标题)
        
'        If strTitle = "检查所见" Or InStr(strTitle, "所见") >= 1 Then
'            ReadReport nvl(rsData!正文), strTmp, rtb所见
'            blnHas描述 = True
'
'        ElseIf strTitle = "诊断意见" Or InStr(strTitle, "意见") >= 1 Then
'            ReadReport nvl(rsData!正文), strTmp, rtb意见
'            blnHas意见 = True
'
'        ElseIf strTitle = "建议" Or InStr(strTitle, "建议") >= 1 Then
'            ReadReport nvl(rsData!正文), strTmp, rtb建议
'            blnHas建议 = True
'        End If
        

        Select Case nvl(rsData!标题)
            Case "检查所见"
                ReadReport nvl(rsData!正文), strTmp, rtb所见
                blnHas描述 = True

            Case "诊断意见"
                ReadReport nvl(rsData!正文), strTmp, rtb意见
                blnHas意见 = True

            Case "建议"
                ReadReport nvl(rsData!正文), strTmp, rtb建议
                blnHas建议 = True

        End Select
        
        rsData.MoveNext
    Wend

    blnForceRead = False
    If blnHas描述 = False And blnHas意见 = False And blnHas建议 = False Then
        blnForceRead = True
        
        If lngFileId <> 0 Then
            labEditState.Caption = "无对应提纲关联"
            picChar.Visible = False
            MsgboxH GetRootHwnd, "无效的报告格式设置。", vbOKOnly, "提示"
        End If
        
        dkpMain.Panes(2).Closed = False
        dkpMain.Panes(3).Closed = False
        dkpMain.Panes(4).Closed = False
        
        dkpMain.Panes(2).iconid = 0
        dkpMain.Panes(3).iconid = 0
        dkpMain.Panes(4).iconid = 0
    Else
        dkpMain.Panes(2).Closed = Not blnHas描述
        dkpMain.Panes(3).Closed = Not blnHas意见
        dkpMain.Panes(4).Closed = Not blnHas建议
        
        dkpMain.Panes(2).iconid = IIf(blnHas描述, 2, 0)
        dkpMain.Panes(3).iconid = IIf(blnHas意见, 3, 0)
        dkpMain.Panes(4).iconid = IIf(blnHas建议, 4, 0)
    End If
        
    rtb所见.Enabled = blnHas描述
    rtb意见.Enabled = blnHas意见
    rtb建议.Enabled = blnHas建议
    
    '读取报告签名等相关信息****************************************
    lab阳性.ForeColor = &H808080
    lab危急.ForeColor = &H808080
    
    If mlngReportID <> 0 Then
        Call ReadResultTag(mlngAdviceId, mblnIsMoved, 0) '支持多报告情况下，才需要传递报告ID
        
        Call ReadVersion(mlngReportID)
        Call ReadSigns(mlngReportID)
    End If

    Call ConfigFaceState(blnForceRead)
    
End Sub

Public Sub ReadResultTag(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, Optional ByVal lngReportID As Long = 0)
'读取结果标记
'结果标记包含阴阳性，危急状态等信息
'lngReportID多报告情况下，可通过该参数指定某份报告
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    lab阳性.ForeColor = &H808080
    lab危急.ForeColor = &H808080
    
    If lngReportID = 0 Then
        strSQL = "select 结果阳性, 危急状态 " & _
                " from 影像检查记录 A, 病人医嘱发送 B " & _
                " where A.医嘱ID=B.医嘱ID and A.发送号=B.发送号 And A.医嘱id=[1]"
    Else
        'TODO:多报告情况下，根据报告ID进行查询...
    End If
    
    If blnMoved Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告结果状态", lngAdviceId)
    If rsData.RecordCount <= 0 Then Exit Sub
    
    If IsNull(rsData!结果阳性) = False Then
        If Val(nvl(rsData!结果阳性)) <> 0 Then
            lab阳性.ForeColor = vbRed
        Else
            lab阳性.ForeColor = &H808080
        End If
    End If
    
    If Val(nvl(rsData!危急状态)) <> 0 Then
        lab危急.ForeColor = vbRed
    Else
        lab危急.ForeColor = &H808080
    End If
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub ReadVersion(ByVal lngReportID As Long)
'读取报告版本
'签名时目标版本才需要增加1
On Error GoTo errhandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "Select 最后版本,签名级别,创建人,保存人,科室ID From 电子病历记录  Where Id =[1]"
    If mblnIsMoved = True Then
        strSQL = Replace(strSQL, "电子病历记录", "H电子病历记录")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询报告签名版本", lngReportID)
    
    If rsTemp.RecordCount > 0 Then
        mlngCreateDeptId = Val(nvl(rsTemp!科室ID))
        mstrCreateUser = nvl(rsTemp!创建人)
        mstrSaveUser = nvl(rsTemp!保存人)
        mlngSignLevel = nvl(rsTemp!签名级别, cprSL_空白)
        mintTargetVer = nvl(rsTemp!最后版本, 1)
'    Else
'        '没有报告时的赋值处理
'        mlngCreateDeptId = 0
'        mstrCreateUser = ""
'        mstrSaveUser = ""
    End If
    
    If mlngSignLevel = cprSL_空白 Then
        mintSourceVer = 0
    Else
        mintSourceVer = mintTargetVer
    End If
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub ReadSigns(ByVal lngReportID As Long)
'------------------------------------------------
'功能：读取签名对象，删除本次签名的对象，重新从数据库读取，确保签名对象的内容跟数据库的一致，签名回退刷新之后调用本过程
'参数： 无
'返回： 无
'-----------------------------------------------
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim reportSignInfo As TReportSignInfo
    Dim strSigns As String
    Dim strSignName As String
    
    mlngSignCount = 0
    
    strSQL = "Select Id,对象标记 From 电子病历内容 Where 文件id= [1] And 对象类型=8 Order By 对象标记"
    
    If mblnIsMoved = True Then
        strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询报告签名记录", lngReportID)
    While rsTemp.EOF = False
        If GetReportSignInfo(Val(rsTemp!ID), reportSignInfo, mblnIsMoved) Then
            reportSignInfo.Key = nvl(rsTemp!对象标记, 0)
            
            If Len(strSigns) > 0 Then strSigns = strSigns & "  "
            
            If InStr(reportSignInfo.姓名, M_STR_TAG_SIGNWITHIMG) > 0 Then
                strSignName = Mid(reportSignInfo.姓名, 1, InStr(reportSignInfo.姓名, M_STR_TAG_SIGNWITHIMG) - 1)
            Else
                strSignName = reportSignInfo.姓名
            End If
            
            strSigns = strSigns & reportSignInfo.前置文字 & strSignName
            
            mstrFinalSignUser = strSignName
            
            If mstrFirstSignUser = "" Then
                mstrFirstSignUser = strSignName
            End If
        End If
        
        rsTemp.MoveNext
    Wend
     
    mlngSignCount = rsTemp.RecordCount
    
    '填写签名文本框
    labSign.Caption = strSigns
    labSign.tag = mlngSignCount    '保留签名次数
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub ReadReport(ByVal strtext As String, ByVal strPros As String, rText As RichTextBox)
    'intType---0 检查所见；1 诊断意见；2 建议
    Dim lngCount As Long
    Dim lngSelStart As Long
    Dim lngPosStart As Long
    Dim lngPosEnd As Long
    Dim aryTextPros() As String
    
    On Error GoTo err
    
    lngSelStart = rText.SelStart
    aryTextPros = Split(strPros, "|")
    
    rText.tag = strPros
    
    rText.SelLength = 0
    rText.SelText = strtext
    '设置颜色
    rText.SelStart = lngSelStart
    rText.SelLength = Len(strtext)
    rText.SelColor = vbBlack
    
    On Error Resume Next
    'rText.Tag 是电子病历格式的对象属性，用“|”分隔，总共26个元素
    rText.SelStart = 0
    rText.SelLength = Len(rText.Text)
    rText.SelFontName = aryTextPros(15)     '  rText.SelFontName
    
    If mintEditFontSize <> 0 Then
        rText.SelFontSize = mintEditFontSize
    Else
        rText.SelFontSize = gbytFontSize
    End If
        
    rText.SelBold = aryTextPros(17)     'rText.SelBold
    rText.SelItalic = aryTextPros(18)   'rText.SelItalic
    
    On Error GoTo 0
    
    '解析当前输入的文字，是否有要素，如果有则用蓝色表示出来
    '先查多选要素
    For lngCount = 1 To Len(strtext)
        lngPosStart = InStr(lngCount, strtext, "{{")
        lngPosEnd = InStr(lngCount, strtext, "}}")
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
    For lngCount = 1 To Len(strtext)
        lngPosStart = InStr(lngCount, strtext, "{<")
        lngPosEnd = InStr(lngCount, strtext, ">}")
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
    
    rText.SelStart = lngSelStart + Len(strtext)
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ParshReportImgData(rsData As ADODB.Recordset, ByVal lngDataFrom As TReportFmtFrom)
'解析报告图像数据
    Dim aryImgPro() As String
    Dim reportImgTag As TReportImgTag
    Dim result As ftpResult
    Dim blnIsAbort As Boolean
    Dim objDcmImg As DicomImage
    
 
    If rsData Is Nothing Then Exit Sub
    
    rsData.MoveFirst
    blnIsAbort = False
    
    While Not rsData.EOF
        '第一个属性说明：0普通图像，1标记图像，2报告图像
        aryImgPro = Split(nvl(rsData!对象属性) & ";;;;;;;;;;;;;;;;;;;;", ";")
        
        reportImgTag.lngFileId = Val(rsData!文件ID)
        reportImgTag.lngTableId = Val(rsData!父ID)
        reportImgTag.strObjectTag = Val(rsData!对象标记)
        reportImgTag.strPros = nvl(rsData!对象属性)
        reportImgTag.lngStartVer = Val(rsData!开始版)
        reportImgTag.strKey = Val(rsData!ID)
        reportImgTag.strImgMarks = ""
        
        If Val(aryImgPro(0)) = 1 Then '标记图
            reportImgTag.lngImgType = ritMark
            
            Call ReadMarkImage(dcmMarkImage.Images, lngDataFrom, reportImgTag)
            
            dcmMarkImage.Visible = True
        End If
        
        If Val(aryImgPro(0)) = 2 Then '报告图
            reportImgTag.lngImgType = ritReport
            reportImgTag.lngFromAdvice = Val(GetReportImagePro(reportImgTag.strPros, "ADVICEID"))
            
            If blnIsAbort = False Then
                result = ReadReportImage(dcmReportImg.Images, reportImgTag)
            Else
                '加载替换图像
                Set objDcmImg = dcmReportImg.Images.AddNew
                dcmReportImg.Images(dcmReportImg.Images.Count).tag = reportImgTag
                
                Call DrawBorder(objDcmImg, 0)
                Call DrawErrorText(objDcmImg, "已被终止")
                
            End If
            
            Call CalcImgView
            
            If result = frAbort Then
                '如果下载异常，且选择终止下载，则退出图像加载处理
                blnIsAbort = True
            End If
        End If
        
        Call rsData.MoveNext
    Wend
End Sub

Private Function GetRootHwnd() As Long
    Dim lngCurHwnd As Long
    
    lngCurHwnd = GetAncestor(hwnd, GA_ROOT)
    
On Error GoTo errhandle
    '在窗口的queryunload事件中调用该方法时，对parent的任何访问都会提示客户端不可用错误
    If Parent.hwnd = lngCurHwnd Then
        If Parent.Visible = False Then
            lngCurHwnd = MainForm.hwnd
        End If
    End If
errhandle:
    If err.Number <> 0 Then
        lngCurHwnd = MainForm.hwnd
    End If
    
    GetRootHwnd = lngCurHwnd
End Function

Private Function ReadMarkImage(objImages As DicomImages, _
    ByVal lngDataFrom As TReportFmtFrom, reportImgTag As TReportImgTag) As Boolean
'读取标记图像
    Dim strFile As String
    Dim strSQL As String
    Dim lngAction As Long
    Dim rsTemp As ADODB.Recordset
    Dim objPicMarks As clsPicMarks
    Dim dblMarkZoom As Double
    Dim strError As String
    Dim strTableName As String
    Dim objDcmImg As DicomImage
    Dim strFileName As String
    
    ReadMarkImage = False
    
    Select Case lngDataFrom
        Case rffReport
            lngAction = 6
            strTableName = "电子病历内容"
            
            If mblnIsMoved Then strTableName = "H电子病历内容"
            
        Case rffSample
            lngAction = 4
            strTableName = "病历范文内容"
            
        Case rffTemplate
            lngAction = 2
            strTableName = "病历文件结构"
            
    End Select
    
    strFileName = "MarkImage_" & reportImgTag.lngFileId & "_" & reportImgTag.strKey & ".JPG"
    strFile = mstrReportImgPath & strFileName
                
    If DirExists(mstrReportImgPath) = False Then Call MkLocalDir(mstrReportImgPath)
    
    If FileExists(strFile) = False Then
        Call Sys.ReadLob(glngSys, lngAction, reportImgTag.strKey, strFile)
    End If
    
    If FileExists(strFile) Then
        Set objDcmImg = ReadDicomFile(strFile, strError)
        
        If objDcmImg Is Nothing Then
            MsgboxH GetRootHwnd, "标记图读取失败:" & strError, vbOKOnly, "提示"
        Else
            objDcmImg.tag = reportImgTag
            
            '绘制边框
            Call DrawBorder(objDcmImg, 0)
            
            Call objImages.Add(objDcmImg)
         
        
            '读取标记picMarks...
            strSQL = "Select 内容文本 " & _
                " From " & strTableName & _
                " Where 文件ID = [1] And 父id=[2] And 对象类型=6 " & _
                " Order By 内容行次"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询报告标记图标记", reportImgTag.lngFileId, reportImgTag.strKey)
        
            While Not rsTemp.EOF
                reportImgTag.strImgMarks = reportImgTag.strImgMarks & nvl(rsTemp!内容文本)
                
                Call rsTemp.MoveNext
            Wend
            
            reportImgTag.strImgFile = strFileName
            objImages(1).tag = reportImgTag
            
            Set objPicMarks = New clsPicMarks
            
            objPicMarks.对象属性 = reportImgTag.strImgMarks
            
            dblMarkZoom = objImages(1).SizeX / Val(GetReportImagePro(reportImgTag.strPros, "width")) * Screen.TwipsPerPixelX
            
            '绘制标记
            Call DrawMarks(objImages(1), objPicMarks, dblMarkZoom)
            
            ReadMarkImage = True
        End If
    Else
        MsgboxH GetRootHwnd, "标记图读取失败", vbOKOnly, "提示"
    End If
End Function
 
 
Private Function DownLoadFtpFile(ByVal lngAdviceId As Long, ByVal strFtpFile As String, ByVal strLocalFile As String, ByVal blnMoved As Boolean) As ftpResult
'下载ftp文件
    DownLoadFtpFile = frNormal
    If Len(mftpConTag.Ip) <= 0 Or Val(mftpConTag.tag) <> lngAdviceId Then
        mftpConTag = GetReportDevice(lngAdviceId, blnMoved)
        mftpConTag.tag = lngAdviceId
        
        If Len(mftpConTag.Ip) <= 0 Then
            DownLoadFtpFile = frAbort
            Exit Function
        End If
    End If
    
    DownLoadFtpFile = FtpDownload(mftpConTag, strFtpFile, strLocalFile)
End Function


Private Function UpLoadFtpFile(ByVal lngAdviceId As Long, ByVal strFtpFile As String, ByVal strLocalFile As String, ByVal blnMoved As Boolean) As ftpResult
'上传ftp文件
    UpLoadFtpFile = frNormal
    If Len(mftpConTag.Ip) <= 0 Or Val(mftpConTag.tag) <> lngAdviceId Then
        mftpConTag = GetReportDevice(lngAdviceId, blnMoved)
        mftpConTag.tag = lngAdviceId
        
        If Len(mftpConTag.Ip) <= 0 Then
            UpLoadFtpFile = frAbort
            Exit Function
        End If
    End If
    
    UpLoadFtpFile = FtpUpload(mftpConTag, strFtpFile, strLocalFile)
End Function

Private Function GetReportDevice(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean) As TFtpConTag
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    strSQL = "select NVl(相关ID, ID) as ID from 病人医嘱记录 where ID=[1]"
    If blnMoved Then strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询主医嘱ID", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then
        Call MsgboxH(GetRootHwnd, "医嘱数据校验失败，未找到报告关联医嘱信息。", vbOKOnly, "提示")
        Exit Function
    End If
    
    strSQL = " Select Decode(A.接收日期,Null,'',to_Char(A.接收日期,'YYYYMMDD')||'/') ||A.检查UID||'/' As URL," & _
            " B.设备号 as 设备号1, B.设备名 As 设备名1, B.FTP用户名 As User1,B.FTP密码 As Pwd1, B.IP地址 As Host1, " & _
                    " decode(B.Ftp目录, null, '/', '/'||B.Ftp目录||'/') As Root1,B.共享目录 as 共享目录1,B.共享目录用户名 as 共享目录用户名1,B.共享目录密码 as 共享目录密码1 " & _
            " From  影像检查记录 A,影像设备目录 B " & _
            " Where A.医嘱ID=[1] And nvl(A.位置一, A.位置二)=B.设备号(+)  "
    If blnMoved Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告图像存储", Val(rsData!ID))
            
    If rsData.RecordCount <= 0 Then
        Call MsgboxH(GetRootHwnd, "未找到报告图对应的存储设备，请检查数据是否正确。", vbOKOnly, "提示")
        Exit Function
    End If
    
    If nvl(rsData!Host1) <> "" Then
        GetReportDevice = FtpTagInstance(rsData!Host1, rsData!User1, rsData!Pwd1, rsData!Root1 & rsData!Url)
    End If
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadReportImage(objImages As DicomImages, reportImgTag As TReportImgTag) As ftpResult
'读取报告图
    Dim strFile As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objPicMarks As clsPicMarks
    Dim dblMarkZoom As Double
    Dim strError As String
    Dim objDcmImg As DicomImage
    Dim strFileName As String
    Dim blnImgReadState As Boolean
    Dim lngAdviceId As Long
    
    
    ReadReportImage = frNormal
    blnImgReadState = True
    
    strFileName = GetReportImagePro(reportImgTag.strPros, "PicName")
    If Len(strFileName) > 0 Then
        
        strFile = FormatFilePath(mstrReportImgPath & "\" & strFileName)
        
        '从ftp下载图像
        If FileExists(strFile) = False Then
            lngAdviceId = GetReportImagePro(reportImgTag.strPros, "ADVICEID")
            If lngAdviceId = mlngAdviceId Then
                ReadReportImage = DownLoadFtpFile(lngAdviceId, strFileName, strFile, mblnIsMoved)
            Else
                '从其他医嘱中下载报告图像
                strSQL = "Select ID From 病人医嘱记录 where Id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱记录", lngAdviceId)
                
                If rsTemp.RecordCount > 0 Then
                    ReadReportImage = DownLoadFtpFile(lngAdviceId, strFileName, strFile, False)
                Else
                    ReadReportImage = DownLoadFtpFile(lngAdviceId, strFileName, strFile, True)
                End If
            End If
            
            
            If ReadReportImage <> frNormal Then
                blnImgReadState = False
            End If
        End If
    Else
        strFile = FormatFilePath(mstrReportImgPath & "\报告图_" & reportImgTag.strKey & ".JPG")
        
        '从数据库读取图像
        If FileExists(strFile) = False Then
            Call Sys.ReadLob(glngSys, 6, reportImgTag.strKey, strFile)
        End If
    End If
    
    If FileExists(strFile) = False Then
        If Len(strError) <= 0 Then strError = "未找到报告图像文件 [" & strFile & "]"
        blnImgReadState = False
    End If
    
    If blnImgReadState Then
        '图像读取成功的处理
        Set objDcmImg = ReadDicomFile(strFile, strError)
        
        If Not objDcmImg Is Nothing Then
            reportImgTag.strImgFile = strFileName
            
            objDcmImg.InstanceUID = Replace(strFileName, ".JPG", "")
            objDcmImg.tag = reportImgTag
            
            Call objImages.Add(objDcmImg)
            Call DrawBorder(objDcmImg, 0)
        Else
            blnImgReadState = False
        End If
    End If
    
    If blnImgReadState = False Then
        '加载失败的图像
        
        Set objDcmImg = objImages.AddNew
        
        objImages(objImages.Count).tag = reportImgTag
        
        Call DrawBorder(objDcmImg, 0)
        Call DrawErrorText(objDcmImg, strError)
        
        If ReadReportImage = frNormal Then Call MsgboxH(GetRootHwnd, "图像读取失败。" & vbCrLf & strError, vbOKOnly, "提示")
    End If
End Function

'Private Function GetReportImgFiles() As String
''获取报告图文件
'End Function
'
'Private Function GetMarkImgFile(objMarks As cPicMarks) As String
''获取标记图文件
'End Function

Public Sub AutoSave()
'TODO:自动保存

End Sub

Public Function PromptSave(ByVal lngNewAdviceId As Long, ByVal lngNewReportId As Long, Optional ByVal blnIsForceHint As Boolean = False) As Boolean
'保存提示
    Dim blnIsHint As Boolean
    
    PromptSave = False
    '如果没有修改，则直接退出
    If mblnIsEditable = False Or IsModify = False Then Exit Function
    
    blnIsHint = False
     
    If lngNewAdviceId <> mlngAdviceId Then
        blnIsHint = True
    End If
    
    If lngNewReportId <> mlngReportID Then
        blnIsHint = True
    End If
    
    If blnIsHint Or blnIsForceHint Then
        If MsgboxH(GetRootHwnd, "报告已被修改，是否保存", vbYesNo Or vbDefaultButton1, "提示") = vbNo Then
            If mlngReportID = 0 Then
                '清除报告状态
                Call UpdateReporter(mlngAdviceId, "")
            End If
            
            mblnIsModifyImage = False
            mblnIsModifyMarks = False
            mblnIsModifyText = False
         
            Exit Function
        End If
    End If
    
    PromptSave = SaveReport
    
End Function


Private Function CreateReport() As Long
'创建报告
    Dim iType As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    
On Error GoTo errhandle
    CreateReport = 0
    
    ' iType：0-从病历文件列表创建报告, 1-从病历范文目录创建报告
    iType = 0
    If mlngSampleId <> 0 Then iType = 1
    
    '创建电子病历内容
    strSQL = "ZL_影像报告内容_创建(" & mlngAdviceId & "," & mlngFileID & "," & mlngSampleId & "," & iType & ")"
    zlDatabase.ExecuteProcedure strSQL, "创建报告格式"
    
    '新创建的报告，从数据库中读取报告内容ID
    strSQL = "Select 病历ID From 病人医嘱报告 Where 医嘱ID= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询新建报告ID", mlngAdviceId)
    
    If rsTemp.EOF = True Then
        MsgboxH GetRootHwnd, "病历创建不正确，无法查找到病历内容ID。", vbOKOnly, "提示"
        Exit Function
    End If
    
    CreateReport = Val(nvl(rsTemp!病历Id))
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DelReportData(ByVal lngReportID As Long, Optional ByVal blnIsErrHint As Boolean = False) As Boolean
'删除报告
    Dim strSQL As String
On Error GoTo errhandle
    DelReportData = False
    
    If Not mobjSpePlugin Is Nothing Then
        If PluginAction(lngReportID, 1) = False Then Exit Function   '删除报告
    End If
    
    strSQL = "Zl_电子病历记录_Delete(" & lngReportID & ")"
    
    zlDatabase.ExecuteProcedure strSQL, "删除报告"
    
    DelReportData = True
Exit Function
errhandle:
    If blnIsErrHint Then
        If ErrCenter() = 1 Then Resume
    End If
    
    Call SaveErrLog
End Function


Private Sub ResetRtbTag(rText As RichTextBox)
    Dim strItem() As String
    Dim i As Integer
    Dim intCnt As Integer
    
    
    '修改该文本框的TAG,如果TAG为空，则暂时不记录
    If rText.tag <> "" Then
        strItem = Split(rText.tag, "|")
        
        strItem(15) = nvl(rText.SelFontName, "宋体")     'FontName
        strItem(17) = nvl(rText.SelBold, "False")    'FontBold
        strItem(18) = nvl(rText.SelItalic, "False")    'FontItalic
        
        rText.tag = ""
        For i = 0 To UBound(strItem()) - 1
            rText.tag = rText.tag & strItem(i) & "|"
        Next i
                
    End If
End Sub

Private Function GetSpecialtyContext() As String
'获取专科报告内容
    Dim strSpeModifyContext As String
    GetSpecialtyContext = ""
    
    If mobjSpePlugin Is Nothing Then Exit Function
    If mobjSpePlugin.pModified = False Then Exit Function
    
    strSpeModifyContext = mobjSpePlugin.getElementString
    
    '如果处于修改状态，且专科报告内容为空，说明是删除了所有专科报告的内容
    If Len(strSpeModifyContext) <= 0 Then
        strSpeModifyContext = "[[@]]专科报告[[;]]"
    End If
    
    'TODO:返回内容部分为测试验证数据
    GetSpecialtyContext = strSpeModifyContext & _
                "[[@]]建议内容[[;]]这是一段报告建议的内容" & vbCrLf & "可以直接在专科报告中书写诊断." & _
                "[[@]]未定义要素[[;]]这个要素是没有定义的" & _
                "" '"[[@]]检查所见[[;]]这是一段检查所见的提纲内容"
End Function

Public Sub WriteContext(ByVal lngReportID As Long, ByRef arrSQL() As String)
    Dim strReport As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strElements As String
    Dim blnInTrans As Boolean
    Dim i As Integer
    Dim intLevel As Integer '签名级别
    Dim strSQLLevel As String '签名查询
    Dim rsTempLevel As ADODB.Recordset '签名查询结果
    Dim strUnitName As String
    Dim strSpecialtyContext As String


    On Error GoTo errhandle
    
    If mobjSpePlugin Is Nothing Then
        If mblnIsModifyText = False Then Exit Sub
    Else
        If mblnIsModifyText = False And mobjSpePlugin.pModified = False Then Exit Sub
    End If


    ReDim Preserve arrSQL(UBound(arrSQL) + 1)


    '修改报告签名要素，将其内容替换为“ ”
    strElements = SPLITER_REPORT & Report_Element_报告签名 & SPLITER_ELEMENT & " "
    '组织专科报告内容
    strSpecialtyContext = GetSpecialtyContext
    strElements = strElements & strSpecialtyContext
    
    '判断专科报告中是否包含检查所见等内容
    If Len(strSpecialtyContext) > 0 Then
        If InStr(strSpecialtyContext, "[[@]]检查所见[[;]]") > 0 Then
            rtb所见.Text = ParseSpecialtyElement(strSpecialtyContext, "检查所见")
        End If
        
        If InStr(strSpecialtyContext, "[[@]]诊断意见[[;]]") > 0 Then
            rtb意见.Text = ParseSpecialtyElement(strSpecialtyContext, "诊断意见")
        End If
        
        If InStr(strSpecialtyContext, "[[@]]建议[[;]]") > 0 Then
            rtb建议.Text = ParseSpecialtyElement(strSpecialtyContext, "建议")
        End If
    End If

    '组织大文本段的对象属性,如果Tag为空，则从数据库读取默认值
    If rtb所见.tag = "" Or rtb意见.tag = "" Or rtb建议.tag = "" Then
        strSQL = "Select a.内容文本 As 标题, b.对象属性 " & _
                " From 电子病历内容 a,电子病历内容 b " & _
                " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 And b.终止版 = 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询报告文本属性", lngReportID)

        While rsTemp.EOF = False
            Select Case rsTemp!标题
                Case "检查所见"
                    If rtb所见.tag = "" Then
                        rtb所见.tag = rsTemp!对象属性
                    End If
                Case "诊断意见"
                    If rtb意见.tag = "" Then
                        rtb意见.tag = rsTemp!对象属性
                    End If
                Case "建议"
                    If rtb建议.tag = "" Then
                        rtb建议.tag = rsTemp!对象属性
                    End If
            End Select
            
            rsTemp.MoveNext
        Wend
    End If
    
    ResetRtbTag rtb所见
    ResetRtbTag rtb意见
    ResetRtbTag rtb建议
    

    '最后保存大文本段内容，此时会根据数据库内容，自动更新报告中的要素
    strReport = SPLITER_REPORT & _
        "1" & rtb所见.tag & SPLITER_ELEMENT & rtb所见.Text & SPLITER_REPORT & _
        "2" & rtb意见.tag & SPLITER_ELEMENT & rtb意见.Text & SPLITER_REPORT & _
        "3" & rtb建议.tag & SPLITER_ELEMENT & rtb建议.Text

    '问题号：80185
    '使用数据里的签名级别
    '更改内容的时候，保存的签名级别始终是0，最后具体的签名级别通过签名的过程来更改
    strSQLLevel = " Select id as 病历id,签名级别 " & _
                " From 电子病历记录 Where id = [1] "
    Set rsTempLevel = zlDatabase.OpenSQLRecord(strSQLLevel, "提取是否签名", lngReportID)
    
    If rsTempLevel.EOF = True Then
        intLevel = 0
    Else
        intLevel = nvl(rsTempLevel!签名级别)
    End If

    strUnitName = zlRegInfo("单位名称")

    strSQL = "ZL_影像报告内容_update(" & mlngAdviceId & "," & _
                                        lngReportID & _
                                        ",'" & Replace(strReport, "'", "’") & _
                                        " ','" & strElements & "'," & _
                                        mintTargetVer & "," & _
                                        intLevel & _
                                        ",'" & strUnitName & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    
'    If chkCritical.value <> 0 Then
'        '危急状态
'        strSQL = "zl_影像检查_危急更新(" & mlngAdviceId & ",1)"
'
'        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'        arrSQL(UBound(arrSQL)) = strSQL
'    End If
'
'    If mblnIgnoreResult = False And chkPositive.value <> 0 Then
'        '没有忽略阴阳性
'        strSQL = "ZL_影像检查_结果(" & mlngAdviceId & ",1)"
'
'        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'        arrSQL(UBound(arrSQL)) = strSQL
'    End If
    
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub WritePicMarks(ByVal lngReportID As Long, ByVal blnCreate As Boolean, _
    ByRef arrSQL() As String)
'写入图像标记
On Error GoTo errhandle
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim reportImgTag As TReportImgTag
    
    '没有标记图则退出
    If mblnIsModifyMarks = False Then Exit Sub
    If dcmMarkImage.Visible = False Then Exit Sub
    If dcmMarkImage.Images.Count <= 0 Then Exit Sub
    
    reportImgTag = dcmMarkImage.Images(1).tag
    
    '没有标记则直接退出
    If Len(reportImgTag.strImgMarks) <= 0 Then
        '直接重新保存标记图片
        strSQL = "ZL_影像报告标注_保存(" & reportImgTag.strKey & ",'')"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    
        Exit Sub
    End If

    If blnCreate = True Then
        '新创建的报告，从电子病历内容中读取标记图ID
        strSQL = "Select Id From 电子病历内容 Where 文件ID=[1] And  对象类型= 5 And substr(对象属性,1,1)='1' "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询新建报告ID", lngReportID)
        
        If rsTemp.EOF = False Then  '有标记图
            reportImgTag.strKey = Val(rsTemp!ID)
        Else    '没有标记图
            reportImgTag.strKey = 0
        End If
        
        '更新标记图属性
        dcmMarkImage.Images(1).tag = reportImgTag
    End If
    
    strSQL = "ZL_影像报告标注_保存(" & reportImgTag.strKey & ",'" & reportImgTag.strImgMarks & "')"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function WriteImages(ByVal lngReportID As Long, ByVal blnCreate As Boolean, _
    ByRef arrSQL() As String, Optional ByRef strReportImgs As String = "") As Boolean
'写入报告图
'只有非转储状态的检查才能执行到此过程
    Dim lngTableId  As Double
    Dim reportImgTag As TReportImgTag
    Dim ftpResult As ftpResult
    Dim strLocalFile As String
    Dim iImgCount As Integer
    Dim strSQL  As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim strPicAttrs As String
    Dim strBufferDir As String
    Dim strRepImgName As String
    Dim lngImageAdiceId As Long
    
    On Error GoTo errhandle
    
    WriteImages = True
     
    If mblnIsModifyImage = False Then Exit Function
    If dcmReportImg.Visible = False Then Exit Function
       
    
    lngTableId = Val(dcmReportImg.tag)
    
    
    If blnCreate = True Then
        strSQL = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
            " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询报告图框ID", lngReportID)
        If rsTemp.RecordCount > 0 Then
            lngTableId = Val(nvl(rsTemp!表格ID))
            dcmReportImg.tag = lngTableId
        End If
        
        '如果是新建报告，但没有报告图像，则退出后续处理
        If dcmReportImg.Images.Count <= 0 Then Exit Function
    Else
        '非新建情况下，如果报告图数量为0，则需要清除报告图
        If dcmReportImg.Images.Count <= 0 Then
            strSQL = "ZL_影像报告图像_保存(" & lngTableId & ",'')"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
            
            Exit Function
        End If
    End If

  
    strBufferDir = mstrReportImgPath
    
    '判断目录是否存在
    Call MkLocalDir(strBufferDir)
    
    '分析和保存每一个图像表格
    iImgCount = dcmReportImg.Images.Count
    strPicAttrs = ""
    
    lngImageAdiceId = mlngAdviceId
    
    For i = 1 To iImgCount
        reportImgTag = dcmReportImg.Images(i).tag
       
        strLocalFile = FormatFilePath(strBufferDir & "\" & reportImgTag.strImgFile)
        If FileExists(strLocalFile) = False Then
            '如果本地文件不存在，则从dicom图像中导出
            'dcmReportImg.Images(i).FileExport strLocalFile, "JPG"
            dcmReportImg.Images(i).FileExport strLocalFile, "BMP"   '避免文件格式头错误...
        End If
        
        '说明报告图像是从其他关联检查中提取，需要将报告图像存储到对应的检查设备中
        If reportImgTag.lngFromAdvice <> 0 Then
            lngImageAdiceId = reportImgTag.lngFromAdvice
        End If
        
        ftpResult = UpLoadFtpFile(lngImageAdiceId, reportImgTag.strImgFile, strLocalFile, False)
                        
        If ftpResult <> frNormal Then
            WriteImages = False
            Exit Function
        End If
         
        '只有医嘱相同的时候，才需要更新报告图状态显示
        If lngImageAdiceId = mlngAdviceId Then
            strReportImgs = strReportImgs & ";" & reportImgTag.strKey
        End If
        
        strRepImgName = GetReportImagePro(reportImgTag.strPros, "picname")
        If Len(strRepImgName) <= 0 Then strRepImgName = reportImgTag.strImgFile
        
        strPicAttrs = strPicAttrs & ";" & strRepImgName & "," & lngImageAdiceId
        
        If Len(reportImgTag.strPros) <= 0 Then
            '如果是新增的报告图像，则没有strPros属性
            reportImgTag.strPros = strPicAttrs
            dcmReportImg.Images(i).tag = reportImgTag
        End If
    Next
 
    strSQL = "ZL_影像报告图像_保存(" & lngTableId & ",'" & strPicAttrs & "')"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ParseSpecialtyElement(ByVal strSpecialtyContext As String, ByVal strElementName As String) As String
'解析专科报告中可能包含的要素，
'要素格式为：格式[[@]]要素名称[[;]]专科报告内容
    Dim lngStartIndex As Long
    Dim strElementContext As String
    
    ParseSpecialtyElement = " "
    
    If Len(strSpecialtyContext) <= 0 Then Exit Function
    
    lngStartIndex = InStr(strSpecialtyContext, strElementName & "[[;]]")
    
    If lngStartIndex <= 0 Then Exit Function
    
    strElementContext = Mid(strSpecialtyContext, lngStartIndex + Len(strElementName & "[[;]]"))
    
    lngStartIndex = InStr(strElementContext, "[[@]]")
    
    If lngStartIndex > 0 Then
        ParseSpecialtyElement = Mid(strElementContext, 1, lngStartIndex - 1)
    Else
        ParseSpecialtyElement = strElementContext
    End If
End Function

Private Function WriteRtfFormat(ByVal lngReportID As Long, ByRef arrSQL() As String) As Boolean
'------------------------------------------------
'功能：保存报告格式RTF文件，对报告进行签名或者回退
'参数：     OneSign -- 不为空，则表示进行签名或者回退；为空，表示只是保存格式，不处理签名
'           blnAddSign 增加或者回退签名，True--增加签名,OneSign为空表示保存报告格式；False--回退签名
'返回： 无，直接保存RTF报告格式文档，对报告签名或者回退
'-----------------------------------------------
On Error GoTo errhandle
    Dim strZipFile As String
    Dim strTemp As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String
    Dim lngSignPos As Long
    Dim strReportFormatFile As String
    Dim strErrCount As String
    Dim strElementContext As String
    Dim strSpecialtyContext As String
    
    strErrCount = ""
    WriteRtfFormat = False
    
reLoad:
    strReportFormatFile = FormatFilePath(SysRootPath & "\ReportFmt" & strErrCount)
    
    '先复制报告格式
    If Dir(strReportFormatFile) <> "" Then Call RemoveFile(strReportFormatFile)
    
    '从数据库读取RTF报告格式文档
    strZipFile = Sys.ReadLob(glngSys, 5, lngReportID, strReportFormatFile)
    
    '解压缩文件
    strTemp = zlFileUnzip(strZipFile)
    
    If strTemp <> "" Then
        If FileExists(strTemp) = False Then
            If MsgboxH(GetRootHwnd, "报告ID为[" & lngReportID & "]的RTF格式文件获取失败，是否重试？", vbYesNo) = vbYes Then
                strErrCount = CStr(Val(strErrCount) + 1)
                GoTo reLoad
            End If
            
            Exit Function
        End If
        '解析文件，根据报告内容，修改其中要素内容
        '读取RTF文件内容
        rtxtSaveElement.Filename = strTemp
        strReport = rtxtSaveElement.TextRTF
        
        strSpecialtyContext = ""
        If Not mobjSpePlugin Is Nothing Then
            strSpecialtyContext = GetSpecialtyContext
        End If
       
        '读取数据库中的要素，把各个要素内容填写到格式中
        strSQL = "Select 对象标记,内容文本,要素名称 From 电子病历内容 Where 文件ID= [1] And 对象类型 = 4 And 终止版=0 and 保留对象 =0 order by 对象标记 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询报告要素", lngReportID)
        
        While (rsTemp.EOF = False)
            strElementContext = nvl(rsTemp!内容文本, "")
            
            If Len(strElementContext) <= 0 Then
                '判断专科报告中是否包含内容
                '格式[[@]]要素名称[[;]]专科报告内容
                
                strElementContext = ParseSpecialtyElement(strSpecialtyContext, nvl(rsTemp!要素名称))
            End If
            
            
            UpdateReportElement strReport, "E", rsTemp!对象标记, strElementContext
            rsTemp.MoveNext
        Wend
        
        '保存RTF文件
        rtxtSaveElement.TextRTF = strReport
        rtxtSaveElement.SaveFile strTemp
            
        '压缩文件
        strZipFile = zlFileZip(strTemp)
        
        '保存格式
        zlSaveLob 5, lngReportID, strZipFile, arrSQL
        
        WriteRtfFormat = True
    
        '删除临时zip文件
        Call RemoveFile(strZipFile)
    Else
        If MsgboxH(GetRootHwnd, "无法读取或者解压报告格式" & strReportFormatFile & vbCrLf & "请使用“病历编辑”的方法来编辑此报告或重试读取，是否重试？", vbYesNo) = vbYes Then
            If Dir(strReportFormatFile) <> "" Then Kill strReportFormatFile
            
            strErrCount = CStr(Val(strErrCount) + 1)
            GoTo reLoad
        End If
    End If
    
    Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function SaveReport(Optional ByRef strReportImages As String = "") As Boolean
'保存报告
    Dim i As Long
    Dim arySql() As String
    Dim blnIsNew As Boolean
    Dim blnIsSaveImg As Boolean
    Dim blnInTrans As Boolean
    
On Error GoTo errhandle
    
    '最后版本应为签名次数+1，如果没有签名，则最后版本为1
'    If blnIsSignSave Then 'If mlngSignLevel <> cprSL_空白 Then
'        mintTargetVer = mintSourceVer + 1   '如果是签名保存，则需要对目标版本加一
'    End If
    
    If Not IsHaveContent() Then
        MsgBoxD Me, "没有有效的报告内容，不允许保存。", vbInformation, gstrSysName
        Exit Function
    End If
            
    mintTargetVer = mlngSignCount + 1   '目标版本直接和签名次数相关联（签名次数+1），没有签名次数时，目标版本为1，如果签名次数为1，则目标版本为2
   
    SaveReport = False
    
    If Not mblnIsEditable Then
    '非编辑状态不允许保存
        MsgboxH GetRootHwnd, "非编辑状态下不允许保存。", vbOKOnly, "提示"
        Exit Function
    End If
    
    '未对报告进行修改时，不进行保存
    If IsModify = False Then Exit Function
    
    '判断报告文本段长度是否超过2000个字符，如果超过，则提示，并退出
    If Len(rtb所见.Text) > 2000 Or Len(rtb意见.Text) > 2000 Or Len(rtb建议.Text) > 2000 Then
        MsgboxH GetRootHwnd, "检查所见、诊断意见或建议的字数超过2000，请删减部分文字后保存。", vbOKOnly, "提示"
        Exit Function
    End If
    
    blnIsNew = False
    If mlngReportID = 0 Then
        '新建报告
        blnIsNew = True
        mlngReportID = CreateReport()
    End If
    
    If mlngReportID = 0 Then
        MsgboxH GetRootHwnd, "未取得有效的报告ID数据，不能继续此操作。", vbOKOnly, "提示"
        Exit Function
    End If
    
    ReDim arySql(0)
    
    Call WriteContext(mlngReportID, arySql)
    
    Call WritePicMarks(mlngReportID, blnIsNew, arySql)
    
    
    blnIsSaveImg = WriteImages(mlngReportID, blnIsNew, arySql, strReportImages)
    
    
    '如果是新建报告，即便没有保存报告内容数据，也需要写入rtf报告格式
    If blnIsNew Or UBound(arySql) > 0 Then Call WriteRtfFormat(mlngReportID, arySql)
    
    If blnIsSaveImg = False Then
        Call MsgboxH(GetRootHwnd, "报告图上传失败,未能保存，请稍后重试。", vbOKOnly, "提示")
    End If
   
    gcnOracle.BeginTrans        '----------保存报告内容
    blnInTrans = True
    For i = 0 To UBound(arySql)
        If Trim(arySql(i)) <> "" Then
            Call zlDatabase.ExecuteProcedure(CStr(arySql(i)), "保存报告内容[" & i & "]")
        End If
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    mblnIsModifyImage = False
    mblnIsModifyMarks = False
    mblnIsModifyText = False
    
    If Not mobjSpePlugin Is Nothing Then
        Call PluginAction(mlngReportID, 0)   '保存报告
        mobjSpePlugin.pModified = False
    End If
    
'    If mlngCreateDeptId <= 0 Then mlngCreateDeptId = mlngDeptID
'    If Len(mstrCreateUser) <= 0 Then mstrCreateUser = UserInfo.姓名
'    If Len(mstrSaveUser) <= 0 Then mstrSaveUser = UserInfo.姓名
    
    SaveReport = True
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
    If blnIsNew Then
        '删除新建的报告
        Call DelReportData(mlngReportID)
        
        mlngReportID = 0
    End If
    
    If blnInTrans Then gcnOracle.RollbackTrans
End Function


Private Function PluginAction(ByVal lngReportID As Long, ByVal lngActionType As Long) As Boolean
    Dim strErr As String
On Error GoTo errhandle
    PluginAction = True
    
    If mobjSpePlugin Is Nothing Then Exit Function
    
    PluginAction = mobjSpePlugin.zlReportAction(lngReportID, lngActionType)
Exit Function
errhandle:
    strErr = err.Description
     
    If lngActionType = 1 Then
        If MsgboxH(GetRootHwnd, "专科报告插件执行异常：" & strErr & vbCrLf & "是否强制删除该份报告？", vbYesNo, "提示") = vbNo Then
            PluginAction = False
        Else
            PluginAction = True
        End If
    Else
        Call MsgboxH(GetRootHwnd, "专科报告插件执行异常：" & strErr, vbOKOnly, "提示")
        
        PluginAction = False
    End If
 
End Function

Public Sub ShowPrintFormat(ByVal strFmtName As String)
On Error GoTo errhandle
    mstrPrintFmts = strFmtName
    
    labFmt.Caption = IIf(mstrEprFmtName <> "", mstrEprFmtName & "：", "") & mstrPrintFmts
Exit Sub
errhandle:

End Sub

Public Function ChangeReportFormat(ByVal lngFmtId As Long) As Boolean
'更改报告格式
    Dim strSQL As String
    Dim strPicSql As String
    Dim strContextSql As String
    Dim rsData As ADODB.Recordset
    Dim strTmp As String
    Dim lngDataFrom As Long
    Dim lngFileId As Long
    Dim blnHas描述 As Boolean
    Dim blnHas意见 As Boolean
    Dim blnHas建议 As Boolean
    
    Dim strSource描述 As String
    Dim strSource意见 As String
    Dim strSource建议 As String
    Dim blnReportVisible As Boolean
    Dim blnMarkVisible As Boolean
     
 
    ChangeReportFormat = False
    
    If mlngReportID <> 0 Or IsModify Then
        If MsgboxH(GetRootHwnd, "更改格式将会覆盖当前报告内容，是否继续？", vbYesNo, "提示") = vbNo Then Exit Function
    End If
     
    If lngFmtId = 0 Then
        '标准格式，既从病历文件结构中读取内容
        
        lngDataFrom = rffTemplate
        lngFileId = mlngFileID
        
        '从病历单据中查询格式数据
        strSQL = "Select  Id As 表格Id From 病历文件结构" & _
                    " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2' " & _
                    " Order By 对象序号"
                    
        strPicSql = "select ID,文件ID,父ID,1 as 开始版,对象标记,对象属性,内容行次 from 病历文件结构 where  文件ID=[1] and 父ID=[2] and 对象类型=5 order by 对象标记"
        
        strContextSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文 " & _
                 " From 病历文件结构 a, 病历文件结构 b" & _
                 " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父id And b.对象类型 = 2 "
    Else
        
        lngDataFrom = rffSample
        lngFileId = lngFmtId
            
        '从范文中查询格式数据
        strSQL = "Select  Id As 表格Id From 病历范文内容 a " & _
                    " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2' " & _
                    " Order By 对象序号"
                    
        strPicSql = "select ID,文件ID,父ID,1 as 开始版,对象标记,对象属性,内容行次 from 病历范文内容 where  文件ID=[1] and 父ID=[2] and 对象类型=5 order by 对象标记"
        
        strContextSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文" & vbNewLine & _
                " From 病历范文内容 a, 病历范文内容 b" & vbNewLine & _
                " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父id And b.对象类型 = 2"
    End If
    
    
    '读取报告文本内容****************************************
    Set rsData = zlDatabase.OpenSQLRecord(strContextSql, "查询报告文本", lngFileId)
    
    strSource描述 = rtb所见.Text
    strSource意见 = rtb意见.Text
    strSource建议 = rtb建议.Text
    
    rtb所见.Text = ""
    rtb意见.Text = ""
    rtb建议.Text = ""
    
    blnHas描述 = False
    blnHas意见 = False
    blnHas建议 = False
    
    While rsData.EOF = False
        strTmp = nvl(rsData!对象属性)
                           
        Select Case nvl(rsData!标题)
            Case "检查所见"
                ReadReport nvl(rsData!正文), strTmp, rtb所见
                blnHas描述 = True
                
            Case "诊断意见"
                ReadReport nvl(rsData!正文), strTmp, rtb意见
                blnHas意见 = True
                
            Case "建议"
                ReadReport nvl(rsData!正文), strTmp, rtb建议
                blnHas建议 = True
                
        End Select
        
        rsData.MoveNext
    Wend
    
    If blnHas描述 = False And blnHas意见 = False And blnHas建议 = False Then
        picChar.Visible = False
        MsgboxH GetRootHwnd, "无效的报告格式设置,不能进行切换。", vbOKOnly, "提示"
        
        '恢复到切换前的文本内容
        rtb所见.Text = strSource描述
        rtb意见.Text = strSource意见
        rtb建议.Text = strSource建议
        
        Exit Function
    Else
        dkpMain.Panes(2).Closed = Not blnHas描述
        dkpMain.Panes(3).Closed = Not blnHas意见
        dkpMain.Panes(4).Closed = Not blnHas建议
    End If
    
    mlngSampleId = lngFmtId
    
    rtb所见.Enabled = blnHas描述
    rtb意见.Enabled = blnHas意见
    rtb建议.Enabled = blnHas建议
    
    
    
    '读取报告图信息****************************************
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告图框", lngFileId)
    
    dcmReportImg.Visible = False
    dcmMarkImage.Visible = False
    
    blnReportVisible = False
    blnMarkVisible = False
    
    dcmReportImg.Images.Clear
    dcmMarkImage.Images.Clear

    If rsData.RecordCount > 0 Then
        '读取标记图，报告图
        
        blnReportVisible = True
        dcmReportImg.Visible = True
        
        '图像对象查询
        dcmReportImg.tag = Val(nvl(rsData!表格ID))
        
        Set rsData = zlDatabase.OpenSQLRecord(strPicSql, "查询报告图片", lngFileId, Val(nvl(rsData!表格ID)))
        If rsData.RecordCount > 0 Then
            
            Call ParshReportImgData(rsData, lngDataFrom)
            
            If dcmMarkImage.Images.Count > 0 Then blnMarkVisible = True
            
            mblnIsModifyMarks = True
        End If
    End If
    
    
    If blnReportVisible = False And blnMarkVisible = False Then
        '关闭报告图
        dkpMain.Panes(1).Closed = True
    Else
        '打开报告图
        dkpMain.Panes(1).Closed = False
        
        If blnMarkVisible = False Then
            dcmReportImg.Width = picImageBack.Width
            ucSplitter1.Visible = False
        Else
            If dcmReportImg.Width = picImageBack.Width Then
                If picImageBack.Width - dcmMarkImage.Width < 0 Then
                    dcmMarkImage.Width = 0.34 * picImageBack.Width
                End If
                
                dcmReportImg.Width = picImageBack.Width - dcmMarkImage.Width
                
                ucSplitter1.Left = dcmReportImg.Width
                ucSplitter1.RePaint
            End If
            
            ucSplitter1.Visible = True
        End If
    End If
    
    
    mblnIsModifyText = True
    mlngReportID = 0
    
    ChangeReportFormat = True
End Function

Public Function SignVerifiy(ByVal lngSignVer As Long) As Boolean
'签名验证
'------------------------------------------------
'功能：校验检查报告的电子签名(可对已转移的数据),校验版本为int签名版本 的签名
'参数： int签名版本 -- 本次需要验证的签名的版本
'       blnMoved -- 数据是否被迁移
'返回：
'-----------------------------------------------
    Dim strSource As String
    Dim dbl签名ID  As Double                  '签名所在的行的ID
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errhandle
    
    SignVerifiy = False
    
    '根据报告ID和签名版本查找签名内容
    strSQL = "Select Id , 开始版 From 电子病历内容 Where 文件ID = [1] And 对象类型 = 8 and 开始版 =[2] "
    If mblnIsMoved Then
        strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取最后签名版本", mlngReportID, lngSignVer)
    If rsTemp.RecordCount = 0 Then
        MsgboxH GetRootHwnd, "本次报告没有版本为" & lngSignVer & "的签名，无法对数字签名做验证。", vbInformation, gstrSysName
        Exit Function
    End If
    
    dbl签名ID = Val(rsTemp!ID)
    
    '提取源文
    strSource = GetSignSource(mlngReportID, lngSignVer, mblnIsMoved)
    
    '如果返回的规则=0，表示提取源文失败
    If Len(strSource) = 0 Then
        MsgboxH GetRootHwnd, "本次报告版本为" & lngSignVer & "的签名源文提取失败，无法对数字签名做验证。", vbOK, "提示"
        Exit Function
    End If
    
    '创建签名对象，对源文进行签名验证
    On Error Resume Next
    If gobjESign Is Nothing Then
        Set gobjESign = Interaction.GetObject(, "zl9ESign.clsESign")
        If gobjESign Is Nothing Then Set gobjESign = DynamicCreate("zl9ESign.clsESign", "电子签名")
        If err <> 0 Then err = 0
        
        If Not gobjESign Is Nothing Then
            Call gobjESign.Initialize(gcnOracle, glngSys)
        End If
    End If
        
    On Error GoTo errhandle
        
    If Not gobjESign Is Nothing Then
        '签名验证
        Call gobjESign.VerifySignature(strSource, dbl签名ID, 2)
        
        SignVerifiy = True
    End If
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Private Function GetSignSource(ByVal lngReportID As Long, ByVal int签名版本 As Integer, ByVal blnMoved As Boolean) As String
'------------------------------------------------
'功能：获取用于电子签名，签名验证的报告源文内容
'参数： int提取类型 -- 1、签名时提取源文；2、签名验证时提取源文
'       lngReportID -- 报告ID，电子病历记录ID
'       int签名版本 -- 本次签名/验证签名提取源文的版本号
'       blnMoved --- 报告数据是否已经转储
'       thisSign --- 签名对象，签名的时候传入此对象，验证签名的时候传入nothing
'       strSourceOut -- 【返回】签名源文
'返回： 签名/验证签名的源文生成规则
'-----------------------------------------------
    Dim intRule As Integer
    Dim lng签名ID  As Long                  '签名所在的行的ID
    Dim strSQL As String
    Dim rs病历记录 As ADODB.Recordset
    Dim rs病历内容 As ADODB.Recordset
    Dim rs签名记录 As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim str签名时间 As String
    Dim arr对象属性() As String
    Dim strSignName As String
    Dim strSignImgBase64 As String
    Dim strImgFiles As String
    Dim strSourceOut As String
    Dim lngImgAdviceId As Long
    
    '源文提取规则：
    'intRule = 1时，提取 ID，病人ID，婴儿，创建人，创建时间，医生姓名，签名级别，签名时间,检查所见，诊断意见，建议
    '验证签名的时候，医生姓名，签名级别，签名时间从签名记录中获取，分别是医生姓名= “内容文本”，签名级别=“要素表示”，签名时间 =“对象属性（5）”
    '签名的时候，医生姓名，签名级别，签名时间 从签名对象中获取
    On Error GoTo err
    
    If lngReportID = 0 Or int签名版本 = 0 Then Exit Function
    
    strSourceOut = ""
     
    '从电子病历记录中提取报告源文的基本信息
    strSQL = "Select ID,病人ID,婴儿,创建人,创建时间 From 电子病历记录 Where Id = [1]"
    If blnMoved Then strSQL = Replace(strSQL, "电子病历记录", "H电子病历记录")
    
    Set rs病历记录 = zlDatabase.OpenSQLRecord(strSQL, "提取报告源文基本信息", lngReportID)
    
    If rs病历记录.RecordCount <= 0 Then Exit Function
    
    '从电子病历内容中提取报告源文的内容信息
    strSQL = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文,b.开始版 as 版本 From 电子病历内容 a,电子病历内容 b " & _
             " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 and b.开始版 = [2]  "
    If blnMoved Then strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
    
    Set rs病历内容 = zlDatabase.OpenSQLRecord(strSQL, "提取报告源文内容信息", lngReportID, int签名版本)
    
    If rs病历内容.RecordCount = 0 Then Exit Function
    
     
    '验证签名，从签名记录中提取医生姓名，签名级别，签名时间信息,签名规则
    strSQL = "Select 内容文本 as 医生姓名 ,要素表示  as 签名级别 ,对象属性 From 电子病历内容 Where 文件ID = [1] And 对象类型 = 8 and 开始版 =[2] "
    If blnMoved Then strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
    
    Set rs签名记录 = zlDatabase.OpenSQLRecord(strSQL, "提取最后报告源文签名信息", lngReportID, int签名版本)
    
    If rs签名记录.RecordCount = 0 Then Exit Function
    
    '提取格式化的签名时间，签名规则
    arr对象属性 = Split(rs签名记录!对象属性, ";")
    If UBound(arr对象属性) >= 5 Then
        intRule = Val(arr对象属性(1))
        str签名时间 = Format(arr对象属性(4), "yyyy-MM-dd HH:mm:ss")
    End If
    If intRule = 0 Then Exit Function
    
    '根据规则组织报告源文： ID，病人ID，婴儿，创建人，创建时间，医生姓名，签名级别，签名时间,检查所见，诊断意见，建议
    If intRule = 1 Then
        '源文基本信息
        strSourceOut = rs病历记录!ID
        strSourceOut = strSourceOut & vbTab & nvl(rs病历记录!病人ID)
        strSourceOut = strSourceOut & vbTab & nvl(rs病历记录!婴儿)
        strSourceOut = strSourceOut & vbTab & nvl(rs病历记录!创建人)
        strSourceOut = strSourceOut & vbTab & nvl(rs病历记录!创建时间)
 
        '验证签名，从数据库签名记录提取
        strSignName = nvl(rs签名记录!医生姓名)
        If InStr(strSignName, M_STR_TAG_SIGNWITHIMG) > 0 Then
            strImgFiles = Split(strSignName, M_STR_TAG_SIGNWITHIMG)(1)
            strSignName = Split(strSignName, M_STR_TAG_SIGNWITHIMG)(0)
        End If

        strSourceOut = strSourceOut & vbTab & strSignName   '姓名
        strSourceOut = strSourceOut & vbTab & nvl(rs签名记录!签名级别)
        strSourceOut = strSourceOut & vbTab & str签名时间
  
        
        '源文报告内容
        rs病历内容.Filter = "标题 ='" & ReportViewType_检查所见 & "'"
        If rs病历内容.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & nvl(rs病历内容!正文)
        End If
        
        rs病历内容.Filter = "标题 ='" & ReportViewType_诊断意见 & "'"
        If rs病历内容.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & nvl(rs病历内容!正文)
        End If
        
        rs病历内容.Filter = "标题 ='" & ReportViewType_建议 & "'"
        If rs病历内容.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & nvl(rs病历内容!正文)
        End If
        
        '源文签名图像信息
        If mblnUseImgSign Then
            '从数据库签名记录提取
            lngImgAdviceId = Val(Split(strImgFiles & "[ADV]", "[ADV]")(1))
            
            If lngImgAdviceId <= 0 Then
                strSignImgBase64 = GetSignedImgB64(strImgFiles, mlngAdviceId, blnMoved)
            Else
                '判断图像对应的医嘱是否已经被转储
                strSQL = "select ID From 病人医嘱记录 where Id=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱记录", lngImgAdviceId)
                
                If rsData.RecordCount > 0 Then
                    strSignImgBase64 = GetSignedImgB64(Split(strImgFiles, "[ADV]")(0), lngImgAdviceId, False)
                Else
                    strSignImgBase64 = GetSignedImgB64(Split(strImgFiles, "[ADV]")(0), lngImgAdviceId, True)
                End If
            End If
            If Len(strSignImgBase64) <= 0 Then Exit Function
            
            strSourceOut = strSourceOut & vbTab & strSignImgBase64
        End If
    End If
    
    GetSignSource = strSourceOut
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetSignedImgB64(ByVal strSignImgFiles As String, ByVal lngImgAdviceId As Long, ByVal blnMoved As Boolean) As String
'获取图像的base64
    Dim i As Long
    Dim aryFile() As String
    Dim strLocalFile As String
    Dim strCurB64 As String
    
    GetSignedImgB64 = ""
    aryFile = Split(strSignImgFiles & ";", ";")
    For i = 0 To UBound(aryFile)
        If Len(aryFile(i)) > 0 Then
            strLocalFile = mstrReportImgPath & aryFile(i)
            If FileExists(strLocalFile) = False Then
                If DownLoadFtpFile(lngImgAdviceId, aryFile(i), strLocalFile, blnMoved) <> frNormal Then
                    Exit Function
                End If
            End If
            
            If FileExists(strLocalFile) = False Then
                GetSignedImgB64 = ""
                MsgboxH GetRootHwnd, "未获取到该版本对应的签名图像，不能验证。", vbOKOnly, "提示"
                Exit Function
            End If
            
            strCurB64 = zlStr.EncodeBase64_File(strLocalFile)
            
            If GetSignedImgB64 <> "" Then GetSignedImgB64 = GetSignedImgB64 & ";"
            GetSignedImgB64 = GetSignedImgB64 & strCurB64
        End If
    Next

End Function


Public Function SignUntread() As Boolean
'签名回退
    Dim signInfo As TReportSignInfo
    Dim strSQL As String
    Dim arrSQL() As String
    Dim blIsUntread As Boolean
    Dim intRobackType As Integer '回退签名类型
    Dim i As Long
    
    SignUntread = False
 
    
    If Val(labSign.tag) = 1 Then  '只有一个签名，表示当前是书写模式下的回退
        signInfo = frmEPRUntread.ShowUntread(mlngReportID, cprET_单病历编辑, Me)
    Else
        signInfo = frmEPRUntread.ShowUntread(mlngReportID, cprET_单病历审核, Me)
    End If
    
    If signInfo.ID <= 0 Then Exit Function
 
    If MsgboxH(GetRootHwnd, "注意：回退操作将不可恢复！是否继续？", vbYesNo + vbDefaultButton2 + vbQuestion, "提示") = vbNo Then Exit Function

    mblnIsLoadData = False '在调用loadreport方法前，需要设置此变量为false，避免载入数据后，文本内容的修改状态为true
    
    '处理两种回退方式
    If signInfo.Key > 0 Then
        ReDim arrSQL(1)

        '清除签名,并保存格式
        If SaveSignFormat(mlngAdviceId, mlngReportID, signInfo, "", True) = False Then Exit Function
    ElseIf signInfo.签名版本 > 1 Then  '回退修订
        '直接修改数据库内容就可以了  '把回退修订保存到数据库
        strSQL = "ZL_影像报告回退(0," & mlngReportID & "," & signInfo.签名版本 & ")"
        zlDatabase.ExecuteProcedure strSQL, "回退报告签名"
    End If
    
    Call ResetContext
    
    Call LoadReport
     
    '载入专科报告
    If Not mobjSpePlugin Is Nothing Then
        
        If mblnIsSpeState Then
            '恢复到专科显示界面
            Call ChangeSepState(True, True)
            
        End If
    End If
    
    mblnIsLoadData = True 'loadreport数据载入完成后，设置为true
    
    SignUntread = True
End Function


Public Function Sign() As Long
'签名 0-未签名，1-诊断签名，2-审核签名
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objSignForm As frmEPRSign
    Dim strImgBase64Code As String
    Dim strImgFiles As String
    Dim signInfo As TReportSignInfo
    Dim strRtfFile As String
    Dim lngImgAdviceId As Long
    
On Error GoTo errhandle
    
    
    Sign = 0
    If Not IsHaveContent() Then
        MsgBoxD Me, "没有有效的报告内容，不允许签名。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '删除rtf格式文件
    strRtfFile = FormatFilePath(SysRootPath & "\TMP.RTF")
    If Dir(strRtfFile) <> "" Then Call RemoveFile(strRtfFile)
    
    '签名之前先保存报告
    If IsModify Then
        If SaveReport() = False Then Exit Function
    Else
        mintTargetVer = mlngSignCount + 1
    End If
    
    If mlngReportID = 0 Then
        MsgboxH GetRootHwnd, "未找到对应报告信息，不能进行签名。", vbOKOnly, "提示"
        Exit Function
    End If
    
    If mintTargetVer >= 16 Then
        MsgboxH GetRootHwnd, "目前系统支持的最大签名版本号为16，请回退或者重新整理！", vbOKOnly, "提示"
        Exit Function
    End If
    
    '获取签名图像
    If mblnUseImgSign Then
        '有检查图的情况下才允许进行对图像进行签名
        If dcmReportImg.Images.Count > 0 Then
            lngImgAdviceId = 0
            
            If GetSignImgEncode(mlngReportID, mintTargetVer, strImgFiles, strImgBase64Code, lngImgAdviceId) = False Then Exit Function
            If lngImgAdviceId = 0 Then lngImgAdviceId = mlngAdviceId
            
            If StorageSignImg(lngImgAdviceId, mlngReportID, mintTargetVer) = False Then Exit Function
        Else
            '不允许单独对标记图进行签名
            If dcmMarkImage.Images.Count > 0 Then
                If MsgboxH(GetRootHwnd, "当前报告没有报告图像，不能单独对标记图进行签名，是否继续？", vbYesNo, "提示") = vbNo Then Exit Function
            End If
        End If
    End If
    
    If mstrFinalSignUser = UserInfo.姓名 Then
        If MsgboxH(GetRootHwnd, "[" & mstrFinalSignUser & "] 用户已进行签名处理，是否继续？", vbYesNo, "提示") = vbNo Then Exit Function
    End If
    
    Set objSignForm = New frmEPRSign
    
    signInfo = objSignForm.ShowSign(UserControl.Parent, mlngSignPassType, mlngReportID, _
                                    mstrPrivs, mlngSignLevel, mstrFirstSignUser, mintTargetVer, _
                                    strImgBase64Code, mlngAdviceId)
    '如果签名信息获取失败，则退出签名
    If signInfo.签名方式 = 0 Or (signInfo.签名方式 = 2 And Len(signInfo.签名信息) <= 0) Then Exit Function
    
    
    '签名格式保存成功后，需要递增版本号
    If SaveSignFormat(mlngAdviceId, mlngReportID, signInfo, strImgFiles) Then
    
        Sign = IIf(signInfo.签名级别 > 1, 2, 1)
        
        mstrFinalSignUser = UserInfo.姓名
        If mstrFirstSignUser = "" Then mstrFirstSignUser = UserInfo.姓名
        
        mlngSignCount = mlngSignCount + 1
        mlngSignLevel = signInfo.签名级别
        mintSourceVer = mintSourceVer + 1
        
        labSign.Caption = labSign.Caption & "  " & UserInfo.姓名
        
        '签名次数增加1
        labSign.tag = mlngSignCount
        
        Call ConfigFaceState
    End If
    
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function CreateEditor() As Object
'创建Editor控件
    Dim objEditor As Object
    
    Set objEditor = Controls.Add("zlRichEditor.Editor", "Editor")
    objEditor.Visible = False '使控件可见
    
    Set CreateEditor = objEditor
End Function

Private Function RemoveEditor(objEditor As Object)
'移除Editor控件
    Controls.Remove objEditor
    Set objEditor = Nothing
End Function



Private Function SaveSignFormat(ByVal lngAdviceId As Long, ByVal lngReportID As Long, signInfo As TReportSignInfo, _
    ByVal strSignFiles As String, Optional ByVal blnIsUntread As Boolean = False) As Boolean
'------------------------------------------------
'功能：保存报告格式RTF文件，对报告进行签名或者回退
'参数：     OneSign -- 不为空，则表示进行签名或者回退；为空，表示只是保存格式，不处理签名
'           blnAddSign 增加或者回退签名，True--增加签名,OneSign为空表示保存报告格式；False--回退签名
'返回： 无，直接保存RTF报告格式文档，对报告签名或者回退
'-----------------------------------------------
On Error GoTo errH
    Dim strZipFile As String
    Dim i As Long
    Dim strTemp As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String
    Dim lngSignPos As Long
    Dim strReportFormatFile As String
    Dim strErrCount As String
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long
    Dim bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    Dim objEditor As Object
    Dim arySql() As String
    Dim blnInTrans As Boolean

    SaveSignFormat = False
    
    strErrCount = ""
    If signInfo.签名方式 = 0 Or (signInfo.签名方式 = 2 And Len(signInfo.签名信息) <= 0) Then Exit Function
    
reLoad:
    strReportFormatFile = FormatFilePath(SysRootPath & "\TMP.RTF")
    If Dir(strReportFormatFile) <> "" Then
        '本地存在报告保存时的rtf格式文件时，则不需要重新读取
        strTemp = strReportFormatFile
    Else
        strReportFormatFile = FormatFilePath(SysRootPath & "\SignFmt" & strErrCount)
        
        '先复制报告格式
        If Dir(strReportFormatFile) <> "" Then Kill strReportFormatFile
        
        '从数据库读取RTF报告格式文档
        strZipFile = Sys.ReadLob(glngSys, 5, lngReportID, strReportFormatFile)
        
        '解压缩文件
        strTemp = zlFileUnzip(strZipFile)
    End If
    

    ReDim arySql(0)
    
    If strTemp <> "" Then
        If FileExists(strTemp) = False Then
            If MsgboxH(GetRootHwnd, "报告ID为[" & lngReportID & "]的RTF格式文件获取失败，是否重试？", vbYesNo) = vbYes Then
                strErrCount = CStr(Val(strErrCount) + 1)
                GoTo reLoad
            End If
            
            Exit Function
        End If
         
        Set objEditor = CreateEditor()
        objEditor.OpenDoc strTemp

        If blnIsUntread Then
            '回退签名
            Call DeleteFromEditor(objEditor, signInfo)
            
            '把回退签名保存到数据库
            strSQL = "ZL_影像报告回退(" & signInfo.ID & "," & lngReportID & ",0)"
            
            ReDim Preserve arySql(UBound(arySql) + 1)
            arySql(UBound(arySql)) = strSQL
        Else
            '查找写入签名的位置
            strSQL = "Select 对象标记 From 电子病历内容 Where 文件ID= [1] And 对象类型 = 4 And 要素名称 ='报告签名' "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询报告签名位", lngReportID)
            lngSignPos = -1
            If rsTemp.EOF = False Then
                bFinded = FindKey(objEditor, "E", nvl(rsTemp!对象标记, 0), lKSS, lKSE, lKES, lKEE, bNeeded)
                If bFinded = True Then lngSignPos = lKEE
            End If
            
            '向指定位置写入签名
            InsertIntoEditor objEditor, signInfo, lngSignPos
            
            '把签名保存到数据库
            strSQL = "ZL_影像报告签名_保存(" & lngReportID & "," & _
                    signInfo.开始版 & "," & signInfo.终止版 & " ,'" & signInfo.对象属性 & "','" & _
                    signInfo.姓名 & strSignFiles & "','" & signInfo.前置文字 & "','" & signInfo.时间戳 & "'," & signInfo.签名级别 & ",'" & signInfo.签名信息 & "')"
            
            ReDim Preserve arySql(UBound(arySql) + 1)
            arySql(UBound(arySql)) = strSQL
            
            
            '更新电子病历记录中的最后版本
            strSQL = "ZL_影像报告内容_update(" & lngAdviceId & "," & _
                                                lngReportID & ",'',''," & _
                                                signInfo.开始版 & "," & _
                                                signInfo.签名级别 & ")"
                                                
            ReDim Preserve arySql(UBound(arySql) + 1)
            arySql(UBound(arySql)) = strSQL
        End If
        
        
        '保存成临时文件
        objEditor.SaveDoc strTemp
        
        '压缩文件
        strZipFile = zlFileZip(strTemp)
        
        '保存格式
        zlSaveLob 5, lngReportID, strZipFile, arySql
    
        '删除临时zip文件
        Kill strZipFile
        
        RemoveEditor objEditor
        
        '批量写入格式
        gcnOracle.BeginTrans        '----------保存报告内容
        blnInTrans = True
        For i = 0 To UBound(arySql)
            If Trim(arySql(i)) <> "" Then
                Call zlDatabase.ExecuteProcedure(CStr(arySql(i)), IIf(blnIsUntread, "回退报告签名", "保存报告签名[" & i & "]"))
            End If
        Next i
        gcnOracle.CommitTrans
        
        SaveSignFormat = True
    Else
        If MsgboxH(GetRootHwnd, "无法读取或者解压报告格式" & strReportFormatFile & vbCrLf & "请使用“病历编辑”的方法来编辑此报告或重试读取，是否重试？", vbYesNo) = vbYes Then
            If Dir(strReportFormatFile) <> "" Then Kill strReportFormatFile
            
            strErrCount = CStr(Val(strErrCount) + 1)
            GoTo reLoad
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
    If blnInTrans Then gcnOracle.RollbackTrans
End Function


Private Function GetSignMarkImgName(ByVal lngReportID As Long, ByVal intSignVer As Integer) As String
    GetSignMarkImgName = "MarkImage_" & lngReportID & "_" & intSignVer & ".JPG"
End Function

Private Function StorageSignImg(ByVal lngImgAdviceId As Long, ByVal lngReportID As Long, ByVal intSignVer As Integer) As Boolean
'保存签名图像
    Dim strFile As String
    Dim strFileName As String
    Dim reportImgTag As TReportImgTag
    
    StorageSignImg = True
    If dcmMarkImage.Images.Count <= 0 Then Exit Function
    
    strFileName = GetSignMarkImgName(lngReportID, intSignVer)
    strFile = mstrReportImgPath & strFileName
    
    If FileExists(strFile) = False Then
        Call dcmMarkImage.Images(1).FileExport(strFile, "JPG")
    End If
    
    reportImgTag = dcmMarkImage.Images(1).tag
    
    If UpLoadFtpFile(lngImgAdviceId, strFileName, strFile, False) <> frNormal Then
        StorageSignImg = False
    End If
    
End Function


Private Function GetSignImgEncode(ByVal lngReportID As Long, ByVal intSignVer As Integer, _
    ByRef strImgFiles As String, ByRef strBase64Code As String, ByRef lngImgAdviceId As Long) As Boolean
'产生图像签名信息
'返回格式为
    Dim strErr As String
On Error GoTo errhandle
    Dim i As Integer
    Dim strFile As String
    Dim strFileName As String
    Dim strResult As String
    Dim strCurB64 As String
    Dim reportImgTag As TReportImgTag
    
    GetSignImgEncode = True
    
    strImgFiles = ""
    strBase64Code = ""
    
    If dcmMarkImage.Images.Count <= 0 And dcmReportImg.Images.Count <= 0 Then Exit Function

    strResult = ""
    
    '处理标记图
    If dcmMarkImage.Images.Count > 0 Then
        '"标记图_" & reportImgTag.lngFileId & "_" & reportImgTag.strKey & ".JPG"
        strFileName = GetSignMarkImgName(lngReportID, intSignVer)
        strFile = mstrReportImgPath & strFileName
        Call dcmMarkImage.Images(1).FileExport(strFile, "JPG")
        
        strCurB64 = zlStr.EncodeBase64_File(strFile)
        If Len(strCurB64) > 0 Then
            strResult = M_STR_TAG_SIGNWITHIMG & strFileName
            
            If strBase64Code <> "" Then strBase64Code = strBase64Code & ";"
            strBase64Code = strBase64Code & strCurB64
        Else
            GetSignImgEncode = False
            MsgboxH GetRootHwnd, "标记图转Base64失败，不能进行图像签名。", vbOKOnly, "提示"
            Exit Function
        End If
    End If
    
    lngImgAdviceId = 0
    
    '处理报告图
    For i = 1 To dcmReportImg.Images.Count
        reportImgTag = dcmReportImg.Images(i).tag
        
        If lngImgAdviceId = 0 Then lngImgAdviceId = reportImgTag.lngFromAdvice
        
        strFileName = reportImgTag.strImgFile
        
        strFile = mstrReportImgPath & strFileName
        
        If FileExists(strFile) = False Then
            Call dcmReportImg.Images(i).FileExport(strFile, "JPG")
        End If
        
        strCurB64 = zlStr.EncodeBase64_File(strFile)
        If Len(strCurB64) > 0 Then
            If strResult = "" Then
                strResult = M_STR_TAG_SIGNWITHIMG
            Else
                strResult = strResult & ";"
            End If
        
            strResult = strResult & strFileName
            If strBase64Code <> "" Then strBase64Code = strBase64Code & ";"
            strBase64Code = strBase64Code & strCurB64
        Else
            GetSignImgEncode = False
            MsgboxH GetRootHwnd, "报告图转Base64失败，不能进行图像签名。", vbOKOnly, "提示"
            Exit Function
        End If
    Next
    
    If lngImgAdviceId <> 0 Then
        strResult = strResult & "[ADV]" & lngImgAdviceId
    End If
    
    strImgFiles = strResult
    
    Exit Function:
errhandle:
    GetSignImgEncode = False
    strErr = err.Description
    
    MsgboxH GetRootHwnd, "图像Base64转换错误，不能进行签名。" & vbCrLf & strErr, vbOKOnly, "提示"
End Function

Public Sub ReportPreview(ByVal strReportNo As String, ByVal strPrintFmts As String)
'报告预览
    Call PrintReport(False, strReportNo, strPrintFmts)
End Sub

Public Function ReportPrint(ByVal strReportNo As String, ByVal strPrintFmts As String, Optional ByVal blnIsBat As Boolean = False) As Boolean
'报告打印
    ReportPrint = PrintReport(True, strReportNo, strPrintFmts, blnIsBat)
End Function


Private Function GetReportId(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean, ByRef lngFileFmtId As Long, Optional ByVal lngSpecifyReportId As Long = 0) As Long
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetReportId = 0
    
    If lngSpecifyReportId = 0 Then
        strSQL = "select a.病历ID, b.文件ID from 病人医嘱报告 a, 电子病历记录 b where a.病历id=b.id and a.医嘱ID=[1]"
        If blnIsMoved Then strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查医嘱报告", lngAdviceId)
    Else
        strSQL = "select a.病历ID, b.文件ID from 病人医嘱报告 a, 电子病历记录 b where a.病历id=b.id and a.病历ID=[1]"
        If blnIsMoved Then strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查医嘱报告", lngSpecifyReportId)
    End If
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    lngFileFmtId = Val(nvl(rsData!文件ID))
    GetReportId = Val(nvl(rsData!病历Id))
End Function


Public Function ReportPreviewEx(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean, _
    Optional ByVal lngSpecifyReportId As Long = 0, Optional ByVal blnIsOneFmt As Boolean = False) As Boolean
    Dim lngReportID As Long
    Dim lngFmtId As Long
    Dim strReportNo As String
    Dim strPrintFmt As String
     
    ReportPreviewEx = False
    
    lngReportID = GetReportId(lngAdviceId, blnIsMoved, lngFmtId, lngSpecifyReportId)
    
    If lngReportID <= 0 Then
        MsgboxH GetRootHwnd, "未找到可打印的检查报告，请确认报告是否书写。", vbOKOnly, "提示"
        Exit Function
    End If
    
    Call Refresh(lngAdviceId, lngFmtId, 0, lngReportID, blnIsMoved)
    
    If GetPrintFormat(mObjNotify.Owner, lngFmtId, strReportNo, strPrintFmt, blnIsOneFmt) = False Then Exit Function
    
    Call ReportPreview(strReportNo, strPrintFmt)
End Function


Private Function GetPrintFormat(Owner As Object, ByVal lngFileId As Long, _
    ByRef strReportNo As String, ByRef strPrintFmt As String, Optional ByVal blnIsOneFmt As Boolean = False) As Boolean
'初始化报告打印格式
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strRegReportNo As String
    Dim blnCancel As Boolean
      
    strReportNo = ""
'    strPrintFmt = ""
    GetPrintFormat = True
    
    '先判断是否使用自定义报表
    strSQL = "Select 通用,编号 From 病历文件列表  Where Id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取报告打印方式", lngFileId)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    If nvl(rsTemp!通用) <> 2 Then Exit Function
    
    strReportNo = "ZLCISBILL" & Format(nvl(rsTemp!编号), "00000") & "-2"
    
    '如果传递了格式，说明只需要获取报表编号
    If Len(strPrintFmt) > 0 Then
        If Split(strPrintFmt, ":")(0) = strReportNo Then
            strPrintFmt = Split(strPrintFmt & ":", ":")(1)
            Exit Function
        End If
    End If
        
    strSQL = "Select b.序号 as ID, a.编号, b.说明 as 名称 From zlreports a,zlrptfmts b Where a.Id=b.报表ID And a.编号=[1] Order By 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取自定义报表格式", strReportNo)
    
    If rsTemp.RecordCount <= 1 Then Exit Function
    
    MainForm.SetFocus
    
'    Call SetActiveWindow(GetRootHwnd)
    
    '判断格式是否允许多选
    If blnIsOneFmt Then
        Set rsTemp = zlDatabase.ShowSQLSelect(Parent, strSQL, 0, "格式选择", True, "ID", "请选择需要打印的格式...", False, False, False, _
                                        Screen.Width / 2 - 3000, Screen.Height / 2 - 2000, 2000, blnCancel, True, False, strReportNo)
    Else
        '如果格式多于一个，则弹出格式选择器
        Set rsTemp = zlDatabase.ShowSQLMultiSelect(Parent, strSQL, 0, "格式选择", True, "ID", "请选择需要打印的格式...", False, False, False, _
                                        Screen.Width / 2 - 3000, Screen.Height / 2 - 2000, 2000, blnCancel, True, False, strReportNo)
    End If
    
    If blnCancel Or rsTemp Is Nothing Then
        GetPrintFormat = False
        Exit Function
    End If
    
    If rsTemp.RecordCount <= 0 Then
        GetPrintFormat = False
        Exit Function
    End If
    
    While Not rsTemp.EOF
        strPrintFmt = strPrintFmt & Val(nvl(rsTemp!ID)) & ","
        Call rsTemp.MoveNext
    Wend
    
End Function

Public Function ReportPrintEx(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean, _
    Optional ByVal lngSpecifyReportId As Long = 0, Optional ByVal blnIsOneFmt As Boolean = False, Optional ByVal strPrintFmts As String = "") As Boolean
    Dim lngReportID As Long
    Dim lngFmtId As Long
    Dim strReportNo As String
    Dim strPrintFmt As String
    
    ReportPrintEx = False
    
    lngReportID = GetReportId(lngAdviceId, blnIsMoved, lngFmtId, lngSpecifyReportId)
    
    If lngReportID <= 0 Then
        MsgboxH GetRootHwnd, "未找到可打印的检查报告，请确认报告是否书写。", vbOKOnly, "提示"
        Exit Function
    End If
    
    Call Refresh(lngAdviceId, lngFmtId, 0, lngReportID, blnIsMoved)
    
    If Len(strPrintFmts) > 0 Then strPrintFmt = strPrintFmts
    If GetPrintFormat(mObjNotify.Owner, lngFmtId, strReportNo, strPrintFmt, blnIsOneFmt) = False Then Exit Function
    
    
    ReportPrintEx = ReportPrint(strReportNo, strPrintFmt)
End Function


'Private Function IsAllowPrint() As Boolean
''判断是否允许报告打印
'On Error GoTo errH
'    Dim strSQL As String
'    Dim rsTemp As ADODB.Recordset
'
'    strSQL = "Select a.报告人,a.复核人,b.紧急标志 ,b.Id From 影像检查记录 a ,病人医嘱记录 b Where a.医嘱id = b.Id And b.Id = [1] "
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "验证是否可以打印", mlngAdviceID)
'
'    If rsTemp.EOF = False Then
'        IsAllowPrint = IIf(nvl(rsTemp!紧急标志, 0) = 1, nvl(rsTemp!报告人) <> "", nvl(rsTemp!复核人) <> "")
'    Else
'        IsAllowPrint = False
'    End If
'
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function


Private Function PrintReport(ByVal blnIsPrint As Boolean, _
    Optional ByVal strReportNo As String, Optional ByVal strPrintFmts As String, _
    Optional ByVal blnSilent As Boolean = False) As Boolean
'blnIsPrint:是否审核后自动打印报告
'blnSilent批量打印调用时，才需要传递此参数

On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnUseCustomReport As Boolean
    Dim objRichEpr As Object
    Dim blnNoAsk As Boolean
    Dim objCusReport As Object
    
    PrintReport = False
    
    '判断报告是否可以打印
    '如果是审核后打印报告，此时数据库还未更新数据，不用调用chkPrintState判断
'    If mblnCheckPrintPara = True And blnIsPrint Then
'        If IsAllowPrint = False Then
'            MsgboxH GetRootHwnd, "当前报告不允许打印。", vbOKOnly, "提示"
'            Exit Function
'        End If
'    End If
    
    If mlngReportID = 0 Then
        MsgboxH GetRootHwnd, "未找到报告相关信息，请先保存报告。", vbOKOnly, "提示"
        Exit Function
    End If
    
    '打印预览前，需要判断是否保存报告
    If IsModify Then Call SaveReport

 
    '打印报告或者预览报告
    If Len(strReportNo) > 0 Then
        '是否静默打印
        blnNoAsk = (zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1")
        If blnSilent = True Then blnNoAsk = True
    
        Set objCusReport = DynamicCreate("zl9Report.clsReport", "自定义报表")
        If objCusReport Is Nothing Then Exit Function
        
'        If Not blnNoAsk Then
            
            '当没有设置打印机时，会弹出打印机设置窗口，因此需要设置一个默认的报告格式
            objCusReport.SetReportPrintSet gcnOracle, glngSys, strReportNo, "Format", Split(strPrintFmts & "-", "-")(0)
            
'            If objCusReport.ReportPrintSet(gcnOracle, glngSys, strReportNo) = False Then
'                '此处刷新会造成界面混乱
'                Exit Function
'            End If
'
'            strPrintFmts = objCusReport.GetReportPrintSet(gcnOracle, glngSys, strReportNo, UserInfo.用户名, 1, , "Format")
'        End If

        
        objCusReport.InitOracle gcnOracle
        
        PrintReport = CustomReportPrint(objCusReport, mlngReportID, strPrintFmts, strReportNo, blnIsPrint)
        
        Call objCusReport.CloseWindows
        Set objCusReport = Nothing
        
    Else        '使用编辑模式打印，调用病历的打印过程
        Set objRichEpr = DynamicCreate("zlRichEPR.cRichEPR", "电子病历")
        If objRichEpr Is Nothing Then Exit Function
    
        objRichEpr.InitRichEPR gcnOracle, mObjNotify.Owner, glngSys, False
        Call objRichEpr.PrintOrPreviewDoc(mObjNotify.Owner, cpr诊疗报告, mlngReportID, blnIsPrint, True)
        
        PrintReport = True
        
        Call objRichEpr.CloseWindows
        Set objRichEpr = Nothing
    End If
    
   
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function




Private Function CustomReportPrint(objCusReport As Object, ByVal lngReportID As Long, _
    ByVal strSelFmts As String, ByVal strReportNo As String, _
    ByVal blnPrint As Boolean) As Boolean
'使用自定义报表打印和预览报告
'参数： blnPrint---True打印；False预览
'       blnSilent ---强制静默打印，批量打印时需要
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strExseNo As String, intExseKind As Integer
    Dim intPCount As Integer
    Dim i As Integer, j As Integer, intParaCount As Integer
    Dim strPicFile As String
    Dim aryRepPara(19) As String, aryMarkPara(1) As String     '报告图中的图像记录
    Dim aryPrintPara(19) As String, strFlagString As String '实际传给自定义报表的内容
    Dim dcmMarkImages As New DicomImages
    Dim dcmRepImages As New DicomImages
    Dim dcmResultImage As DicomImage
    Dim arr报表格式() As String
    Dim int格式号 As Integer
    Dim intRows As Integer, intCols As Integer
    Dim blnIsImageReport As Boolean
    Dim strPicSql As String
    Dim aryImgPro() As String
    Dim reportImgTag As TReportImgTag
    Dim lngReportBoxCount As Long
    Dim lngReportImgCount As Long
    Dim strFirstFile As String

On Error GoTo errhandle
 
    CustomReportPrint = False
    
    '提取报告的记录性质和No
    strSQL = "Select 记录性质, No From 病人医嘱发送 Where 医嘱id = [1]"
    If mblnIsMoved = True Then strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提前记录性质和No", mlngAdviceId)
    If rsTemp.RecordCount = 0 Then Exit Function

    strExseNo = "" & rsTemp!no
    intExseKind = Val("" & rsTemp!记录性质)


    '获取报告图像（包括标记图）生成本地文件
    '一个报告表格中可能排列多个报告图
    strSQL = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
    If mblnIsMoved = True Then strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取图像", lngReportID)

    If rsTemp.RecordCount > 0 Then
        strPicSql = "select ID,文件ID,父ID,开始版,对象标记,对象属性,内容行次 from 电子病历内容 where  文件ID=[1] and 父ID=[2] and 对象类型=5 order by 对象标记"
        Set rsTemp = zlDatabase.OpenSQLRecord(strPicSql, "查询预览打印图片", lngReportID, Val(nvl(rsTemp!表格ID)))
        
        If rsTemp.RecordCount > 0 Then
            intPCount = 0
            Do While Not rsTemp.EOF
                aryImgPro = Split(nvl(rsTemp!对象属性) & ";;;;;;;;;;;;;;;;;;;;", ";")
                
                reportImgTag.lngFileId = Val(rsTemp!文件ID)
                reportImgTag.lngTableId = Val(rsTemp!父ID)
                reportImgTag.strObjectTag = Val(rsTemp!对象标记)
                reportImgTag.strPros = nvl(rsTemp!对象属性)
                reportImgTag.lngStartVer = Val(rsTemp!开始版)
                reportImgTag.strKey = Val(rsTemp!ID)
                reportImgTag.strImgMarks = ""
            
                If Val(aryImgPro(0)) = 1 Then '标记图
                    reportImgTag.lngImgType = ritMark
                    
                    strPicFile = mstrReportImgPath & GetSignMarkImgName(mlngReportID, 0)
                    If ReadMarkImage(dcmMarkImages, rffReport, reportImgTag) = False Then
                        Exit Function
                    End If
                    
                    If dcmMarkImages.Count > 0 Then
                        dcmMarkImages(1).FileExport strPicFile, "BMP"
                    End If
                    
                    aryMarkPara(0) = strPicFile
                End If
                
                If Val(aryImgPro(0)) = 2 Then '报告图
                    reportImgTag.lngImgType = ritReport
                    
                    '获取报告图文件名称
                    strPicFile = GetReportImagePro(reportImgTag.strPros, "PicName")
                    If Len(strPicFile) > 0 Then
                        strPicFile = FormatFilePath(mstrReportImgPath & "\" & strPicFile)
                    Else
                        '如果图片存储在数据库中，则没有picname属性
                        strPicFile = FormatFilePath(mstrReportImgPath & "\报告图_" & reportImgTag.strKey & ".JPG")
                    End If
                     
                    If ReadReportImage(dcmRepImages, reportImgTag) <> frNormal Then Exit Function
                    
                    aryRepPara(intPCount) = strPicFile
                    
                    intPCount = intPCount + 1
                    If intPCount > UBound(aryRepPara) Then Exit Do
                End If
                
                Call rsTemp.MoveNext
            Loop
        End If
    End If


    '根据选择的自定义报表格式，组织图像
    '如果只选择了一种格式，则检查是否只有一个图象框,只有一个图像框的时候，自动组合图像。
    '如果选择了2种以上的格式，则对只有一个图像框的情况不作自动组合
    arr报表格式 = Split(strSelFmts, ",")

    '处理没有选择格式的情况
    If Trim(strSelFmts) = "" Then
        ReDim arr报表格式(0)
        arr报表格式(0) = "1-1"
    End If
 

    '获取图像，调用报表
    lngReportImgCount = intPCount
    
    blnIsImageReport = False
    intPCount = 0       '记录图像的数量
    
    If lngReportImgCount > 0 Then strFirstFile = aryRepPara(0)
    
    For i = 0 To UBound(arr报表格式)
        If arr报表格式(i) <> "" Then
            int格式号 = Split(arr报表格式(i), "-")(0)
    
            strSQL = "Select b.名称,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
            "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = [2]" & vbNewLine & _
            "       Order By b.名称" 'Trunc(b.y/567),Trunc(b.x/567)
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取图象框", strReportNo, int格式号)
            
            lngReportBoxCount = rsTemp.RecordCount
            
            rsTemp.Filter = "名称 like '%标记%'"
            lngReportBoxCount = lngReportBoxCount - rsTemp.RecordCount
            
            rsTemp.Filter = ""
            
            '报告图框只有一个，而报告图有多个时，需要组合图像
            If lngReportImgCount > 1 Then
                aryRepPara(0) = strFirstFile
                
                If lngReportBoxCount = 1 Then
                    '组合图象
                    ResizeRegion lngReportImgCount, rsTemp("W"), rsTemp("H"), intRows, intCols
                    Set dcmResultImage = AssembleImage(dcmRepImages, intRows, intCols, rsTemp("H"), rsTemp("W"))
                    
                    aryRepPara(0) = Replace(Right(aryRepPara(0), Len(aryRepPara(0)) - InStr(aryRepPara(0), "=")), ".JPG", "") & "_GRP.JPG"
                    
                    dcmResultImage.FileExport aryRepPara(0), "JPEG"
                End If
            End If
    
            '装载图像数据
            intParaCount = 0
            Do While Not rsTemp.EOF
                blnIsImageReport = True
    
                '分别装在标记图和报告图
                If InStr(rsTemp!名称, "标记") <> 0 Then '标记图
                    If aryMarkPara(0) <> "" Then strFlagString = rsTemp!名称 & "=" & aryMarkPara(0)
                Else    '报告图
                    If intPCount > UBound(aryRepPara) Then Exit Do      '当遍历的报表中的图像数量超过实际报告图像数量，退出
                    If aryRepPara(intPCount) <> "" Then          '报表中的图象框比报告中的多，退出
                        aryPrintPara(intParaCount) = rsTemp!名称 & "=" & aryRepPara(intPCount)
                        intParaCount = intParaCount + 1
                    End If
                    
                    If lngReportBoxCount <> 1 Then intPCount = intPCount + 1
                    
                End If
                rsTemp.MoveNext
            Loop
    
            '处理报表中图形比报告中少的情况
            For j = intParaCount To UBound(aryPrintPara)
                If aryPrintPara(j) Like "*=*" Then aryPrintPara(j) = ""
            Next j
    
            '如果是报告预览，无图时，则不进行提示
            If blnIsImageReport And blnPrint Then
                If Trim(aryPrintPara(0)) = "" _
                    And Trim(aryPrintPara(1)) = "" _
                    And Trim(aryPrintPara(2)) = "" _
                    And Trim(aryPrintPara(3)) = "" _
                    And Trim(aryPrintPara(4)) = "" _
                    And Trim(aryPrintPara(5)) = "" _
                    And Trim(aryPrintPara(6)) = "" _
                    And Trim(aryPrintPara(7)) = "" _
                    And Trim(aryPrintPara(8)) = "" _
                    And Trim(aryPrintPara(9)) = "" Then
                    
                    If MsgboxH(GetRootHwnd, "未发现待打印的报告图像，是否继续打印？", vbYesNo, "提示") = vbNo Then
                        Exit Function
                    End If
                End If
            End If
    
            '调用报表
            Call objCusReport.ReportOpen(gcnOracle, glngSys, strReportNo, Nothing, _
                "NO=" & strExseNo, "性质=" & intExseKind, "医嘱ID=" & mlngAdviceId, strFlagString, _
                aryPrintPara(0), aryPrintPara(1), aryPrintPara(2), aryPrintPara(3), aryPrintPara(4), aryPrintPara(5), _
                aryPrintPara(6), aryPrintPara(7), aryPrintPara(8), aryPrintPara(9), aryPrintPara(10), aryPrintPara(11), _
                aryPrintPara(12), aryPrintPara(13), aryPrintPara(14), aryPrintPara(15), aryPrintPara(16), aryPrintPara(17), _
                aryPrintPara(18), aryPrintPara(19), "ReportFormat=" & int格式号, IIf(blnPrint, 2, 1))
                
            CustomReportPrint = True
        End If
    Next i

Exit Function
errhandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ReportReject() As Boolean
'报告驳回
Dim objfrmRj As frmReject
Dim i As Long
Dim lngAdviceColIndex As Long
Dim lngProcedureColIndex As Long
Dim lngRowIndex As Long
    
On Error GoTo errFree
    If mlngReportID <= 0 Then
        MsgboxH GetRootHwnd, "当前检查没有报告，不能进行驳回操作。", vbOKOnly, "提示"
        Exit Function
    End If
    
    Set objfrmRj = New frmReject
    
    ReportReject = objfrmRj.ShowRejectWindow(mlngAdviceId, mlngReportID, mObjNotify.Owner)
    
errFree:
    Unload objfrmRj
    Set objfrmRj = Nothing
End Function


Public Sub RejectHistory()
'显示驳回历史
Dim frmRj As frmReject
    
On Error GoTo errFree
    If mlngReportID <= 0 Then
        MsgboxH GetRootHwnd, "当前检查没有报告，不存在驳回历史记录。", vbInformation, "提示"
        Exit Sub
    End If
    
    Set frmRj = New frmReject
    
    Call frmRj.ShowRejectHistory(mlngAdviceId, mlngReportID, mObjNotify.Owner)
errFree:
    Unload frmRj
    Set frmRj = Nothing
End Sub


Public Sub RevisionHistory()
    Dim objHistory As New frmReportHistory
    
    Call objHistory.ZlShowMe(mObjNotify.Owner, mlngAdviceId, mlngReportID, mblnIsMoved)
End Sub

Public Sub ClearMark(Optional ByVal blnIsTriggerModify As Boolean = False)
'清除标记
    Dim reportImgTag As TReportImgTag
    
    If dcmMarkImage.Images.Count > 0 Then
        dcmMarkImage.Images(1).Labels.Clear
        dcmMarkImage.Refresh
        
        reportImgTag = dcmMarkImage.Images(1).tag
        reportImgTag.strImgMarks = ""
        
        dcmMarkImage.Images(1).tag = reportImgTag
        
        If blnIsTriggerModify Then Call EnterModify(, , True)
    End If
End Sub

Public Sub ClearReportImg()
'清除报告图
    dcmReportImg.Images.Clear
End Sub




Public Sub ClearInfo()
'清除签名信息
    mlngCreateDeptId = mlngDeptID ' 0
    mstrCreateUser = UserInfo.姓名 ' ""
    mstrSaveUser = UserInfo.姓名 ' ""
    
    mblnIsLockingEdit = False
    mlngReportID = 0
    
    mlngSignLevel = cprSL_空白
    mstrFirstSignUser = ""
    mstrFinalSignUser = ""
    mintTargetVer = 1
    mintSourceVer = 0
    
    labEditState.Caption = ""
    labSign.Caption = ""
    
'    chkPositive.value = 0
'    chkCritical.value = 0
    
    mblnIsModifyImage = False
    mblnIsModifyMarks = False
    mblnIsModifyText = False
End Sub

Public Function LockEditor(Optional ByRef strErrMsg As String) As Boolean
'锁定编辑人
    '使用全局临时表进行并发处理
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
On Error GoTo errhandle
    
    LockEditor = False
    
    If mblnIsLockingEdit Then
        LockEditor = True
        Exit Function
    End If
    
    strSQL = "select 报告操作 from 影像检查记录 where 医嘱ID=[1]"
    
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告操作人", mlngAdviceId)
    If rsData.RecordCount > 0 Then
        If nvl(rsData!报告操作) <> "" Then
            '已被锁定，判断锁定人是否相同
            If nvl(rsData!报告操作) = UserInfo.姓名 Then
                LockEditor = True
                mblnIsLockingEdit = True
            Else
                strErrMsg = "报告已被 [" & nvl(rsData!报告操作) & "] 编辑锁定."
            End If
            
            Exit Function
        End If
    End If
        
        
    '没有锁定，则进行锁定操作
    Call UpdateReporter(mlngAdviceId, UserInfo.姓名)
     
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告操作人", mlngAdviceId)
    If rsData.RecordCount <= 0 Then
        '锁定失败
        strErrMsg = "报告锁定失败."
        Exit Function
    End If
    
    If nvl(rsData!报告操作) = UserInfo.姓名 Then
        LockEditor = True
        mblnIsLockingEdit = True
    Else
        strErrMsg = "报告锁定失败,已被 [" & nvl(rsData!报告操作) & "] 编辑锁定."
    End If
    
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub UnlockEditor()
'解除编辑人
    Call UpdateReporter(mlngAdviceId, "")
End Sub

Public Function IsLockEditor(Optional ByRef strEditor As String = "")
'报告是否锁定编辑
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo errhandle
    
    IsLockEditor = False
    
    strSQL = "Select 报告操作 From 影像检查记录 Where 医嘱ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "读取记录", mlngAdviceId)
     
    If rsTemp.RecordCount > 0 Then
        If nvl(rsTemp!报告操作) <> "" And nvl(rsTemp!报告操作) <> UserInfo.姓名 Then
            strEditor = nvl(rsTemp!报告操作)
            IsLockEditor = True
        End If
    End If
    
    Exit Function
    
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ResetEditState(Optional ByVal blnForceRead As Boolean = False)
'是否只读
    '只读状态只能进行查看，预览，打印等操作
    
    '已经转储的报告为只读
    '已完成的检查，且没有补录权限为只读
    '已出院且归档的报告为只读
    
    
'是否编辑
    '非编辑状态可进行回退，驳回，审核等操作
    
    '如果没有修改他人报告的权限，且报告创建人为他人则不能编辑
    '如果报告已经审核，则不能编辑，除非创建人和审核人相同
    '如果没有报告编辑权限，则不能编辑
    '如果报告已经被其他用户锁定编辑时，则不能继续进行编辑
    
'只读状态下的报告，肯定不能进行编辑
    
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngAdviceState As Long
    Dim lngCurSignLevel As Long
    Dim strEditor As String
    
    If blnForceRead Then
    '强制读状态
        mblnIsEditable = False
        mblnIsReadOnly = True
        Exit Sub
    End If
    
    mblnIsEditable = Not mblnIsMoved    '已经转储的报告不能编辑
    mblnIsReadOnly = mblnIsMoved
    mblnIsComplete = mblnIsMoved
   
    If mblnIsReadOnly Then Exit Sub '如果已经是只读，则不需要后续判断
    
    '*****************************
    '查询医嘱执行状态,和是否出院归档
    strSQL = "Select b.执行科室ID, a.执行过程,c.出院日期,c.病案状态,c.封存时间 " & _
        " From 病人医嘱发送 a,病人医嘱记录 b,病案主页 c  " & _
        " Where a.医嘱ID = b.Id And  b.病人ID = c.病人ID(+) And b.主页ID = c.主页ID(+) And a.医嘱ID= [1] "
    If mblnIsMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询医嘱报告状态", mlngAdviceId)
    If rsTemp.RecordCount > 0 Then
        lngAdviceState = nvl(rsTemp!执行过程, 0)
        
        mblnIsComplete = IIf(lngAdviceState = 6, True, False)
        
        If mlngReportID = 0 And lngAdviceState = 6 Then '补录报告只能补录医嘱对应执行科室下的报告，不能跨科室补录报告
            '如果没有报告且检查已完成时，只有具备“补录报告”权限时，才能编辑
            If CheckPopedom(mstrPrivs, "补录报告") And Val(nvl(rsTemp!执行科室ID)) = mlngDeptID Then
                labEditState.Caption = "检查已完成"
                labEditState.ForeColor = vbBlue
                
                Exit Sub
            End If
        End If
    
        
        '已完成的报告，为只读状态
        mblnIsReadOnly = IIf(lngAdviceState = 6 Or lngAdviceState = 0, True, False)
        If mblnIsReadOnly Then
            labEditState.Caption = IIf(lngAdviceState = 0, "未报到...", "报告已完成")
            labEditState.ForeColor = vbBlue
            
            mblnIsEditable = False
            Exit Sub
        End If
        
        '出院且归档后，报告不可操作,病案状态为5表示审查归档
        mblnIsReadOnly = IIf(nvl(rsTemp!出院日期) <> "" And (nvl(rsTemp!病案状态, 0) = 5 Or nvl(rsTemp!封存时间, "") <> ""), True, False)
        If mblnIsReadOnly Then
            mblnIsComplete = True '已归档报告表示已完成
            labEditState.Caption = "报告已归档"
            labEditState.ForeColor = vbBlue
            
            mblnIsEditable = False
            Exit Sub
        End If
    End If
    
    '*****************************
    '低级别的医生不能修订高级别医生的报告，打开报告后，报告为只读的。
    '这种情况只有在报告已经签名后再去考虑，所以签名级别<>0。修改后未签名的，在后续的chkEditState中处理。
    If mintSourceVer > 0 Then
        '自己书写的报告，应该是可以回退的
        '提取当前用户的签名级别
        lngCurSignLevel = GetUserSignLevel(UserInfo.ID)
        If lngCurSignLevel < mlngSignLevel Then
            If mstrFirstSignUser = mstrFinalSignUser And mstrSaveUser = mstrFinalSignUser And mstrFinalSignUser = UserInfo.姓名 Then
                '自己创建并签名的报告，有可能后面被调整了用户签名级别
                
            Else
                labEditState.Caption = "级别不足无权编辑"
                labEditState.ForeColor = vbRed
                
                mblnIsReadOnly = True
                mblnIsEditable = False
                
                Exit Sub
            End If
        End If
    End If
    
    '*****************************
    '报告操作人判断
    If IsLockEditor(strEditor) Then
        labEditState.Caption = "报告正被[" & strEditor & "]编辑"
        labEditState.ForeColor = vbRed
        
        mblnIsReadOnly = True
        mblnIsEditable = False
        Exit Sub
    End If
    
    
    '*****************************
    '判断创建人和当前用户是否相同，如果不同，则不允许编辑
    If mintSourceVer = 0 And CheckPopedom(mstrPrivs, "PACS报告书写") Then
        '有报告书写权限
        If mstrCreateUser = UserInfo.姓名 Then
            mblnIsEditable = True
        ElseIf CheckPopedom(mstrPrivs, "PACS他人报告") And (mlngCreateDeptId = mlngDeptID Or IsContainDept(UserInfo.ID, mlngCreateDeptId)) Then   '有他人报告权限的，可以书写本科室的报告
            mblnIsEditable = True
        Else
            labEditState.Caption = "无权编辑[" & mstrCreateUser & "]的报告"
            labEditState.ForeColor = vbRed
            
            mblnIsReadOnly = True '无他人报告权限时，在他人没有签名情况下，不允许进行任何操作
            mblnIsEditable = False
            Exit Sub
        End If
    ElseIf mintSourceVer > 0 Then   '已经签名的报告，不允许直接删除，必须先进行回退处理
        If (CheckPopedom(mstrPrivs, "PACS报告修订")) And (mlngCreateDeptId = mlngDeptID Or IsContainDept(UserInfo.ID, mlngCreateDeptId)) Then  ' Or CheckPopedom(mstrPrivs, "PACS他人报告")
            '有报告修订权限
            '在报告修订的状态下，有报告修订权限的人，可以书写本科室的报告。
            'mstrCreateUser = UserInfo.姓名 And mstrSaveUser <> UserInfo.姓名表示报告由自己创建并已被他人修订或审核
            If (mstrSaveUser = UserInfo.姓名) Or (Not (mstrCreateUser = UserInfo.姓名 And mstrSaveUser <> UserInfo.姓名)) Then     '报告最后是自己最后保存的，或者前面的修改者已经签名
                mblnIsEditable = True
            Else
                '已经有人在修订这个报告,修改已经保存，但是没有签名，报告不可编辑，记录修订人名称
                labEditState.Caption = "已被[" & mstrSaveUser & "]修订"
                mblnIsEditable = False
                Exit Sub
            End If
        ElseIf mstrFirstSignUser = UserInfo.姓名 And mstrFinalSignUser = UserInfo.姓名 Then '如果没有修订或他人报告权限，则判断是否诊断签名和当前用户相同
            '当具备他人报告权限，在对他人创建的报告进行签名后，首次签名人和最终签名人是相同的
            mblnIsEditable = True
        Else
            '只有具备书写，修订，审核或他人报告且级别大于等于当前签名级别权限才能进行回退,如果只具备报告书写权限，在没有审核情况下是允许进行回退的
            If Not (CheckPopedom(mstrPrivs, "PACS报告审核") Or CheckPopedom(mstrPrivs, "PACS报告终审")) Then
                '无修订，无审核，无终审权限时为只读，不能进行回退和删除驳回
                mblnIsReadOnly = True
            End If
            
            If mstrCreateUser = UserInfo.姓名 And mstrFinalSignUser <> UserInfo.姓名 Then
                labEditState.Caption = "报告已被[" & mstrFinalSignUser & "]修订"
            Else
                labEditState.Caption = "无权修订[" & mstrCreateUser & "]的报告"
            End If
            labEditState.ForeColor = vbRed
            
            mblnIsEditable = False
            Exit Sub
        End If
    Else
        If mintSourceVer <= 0 Then  '没有进行任何签名,未签名报告不允许进行审核
            mblnIsReadOnly = True
        End If
        
        If mstrCreateUser <> UserInfo.姓名 Then
            labEditState.Caption = "无权编辑[" & mstrCreateUser & "]的报告"
        Else
            labEditState.Caption = "无报告书写权限"
        End If
        
        labEditState.ForeColor = vbRed
        
        mblnIsEditable = False
        Exit Sub
    End If
     
    '*****************************
    '只能填写检查技师为自己的报告
    If mblnTechReptSame Then
        strSQL = " select 检查技师 from 影像检查记录 where 医嘱id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询检查技师", mlngAdviceId)
        
        If rsTemp.RecordCount < 1 Then
            labEditState.Caption = "无效检查数据禁止编辑"
            labEditState.ForeColor = vbRed
            
            mblnIsEditable = False
            Exit Sub
        End If
        
        If nvl(rsTemp!检查技师) <> UserInfo.姓名 Then
            
            labEditState.Caption = "只能书写自己检查的报告，当前检查没有确定检查技师。"
            If nvl(rsTemp!检查技师, "") <> "" Then
                labEditState.Caption = labEditState.Caption & "，无权对[" & nvl(rsTemp!检查技师) & "]的检查书写"
            End If
            labEditState.ForeColor = vbRed
            
            mblnIsEditable = False
            Exit Sub
        Else
            mblnIsEditable = True
        End If
    
    End If

    '有图像才能书写报告
    If mblnIsEditWithReportImage Then
        strSQL = " select 检查UID from 影像检查记录 where 医嘱id = [1]"
        If mblnIsMoved Then strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询检查UID", mlngAdviceId)
        
        If rsTemp.RecordCount < 1 Then
            labEditState.Caption = "无效检查数据禁止编辑"
            labEditState.ForeColor = vbRed

            mblnIsEditable = False
            
            Exit Sub
        End If
        
        If Len(nvl(rsTemp!检查UID)) <= 0 Then
            labEditState.Caption = "无检查图像禁止书写"
            labEditState.ForeColor = vbRed

            mblnIsEditable = False
            Exit Sub
        End If
    End If
    
    '显示报告当前状态
    If mlngReportID = 0 Then
        labEditState.Caption = "开始编辑..."
        labEditState.ForeColor = vbBlue
    Else
        If mintSourceVer >= 1 Then
            labEditState.Caption = "报告修订中..."
        Else
            labEditState.Caption = "报告书写中..."
        End If
        
        labEditState.ForeColor = vbBlue
    End If
    
End Sub

Public Sub ConfigFaceState(Optional ByVal blnForceRead As Boolean = False, Optional ByVal strEditState As String = "")
'配置界面状态
    Call ResetEditState(blnForceRead)
    
    rtb所见.Locked = Not mblnIsEditable Or mblnIsReadOnly
    rtb意见.Locked = Not mblnIsEditable Or mblnIsReadOnly
    rtb建议.Locked = Not mblnIsEditable Or mblnIsReadOnly
    
'    chkPositive.Enabled = mblnIsEditable Or Not mblnIsReadOnly
'    chkCritical.Enabled = mblnIsEditable Or Not mblnIsReadOnly
'    picState.Enabled = mblnIsEditable Or Not mblnIsReadOnly
    picImageBack.Enabled = mblnIsEditable Or Not mblnIsReadOnly
    
    If Not mblnIsEditable Or mblnIsReadOnly Then
        rtb所见.BackColor = UserControl.BackColor
        rtb意见.BackColor = UserControl.BackColor
        rtb建议.BackColor = UserControl.BackColor
    Else
        rtb所见.BackColor = ColorConstants.vbWhite
        rtb意见.BackColor = ColorConstants.vbWhite
        rtb建议.BackColor = ColorConstants.vbWhite
    End If
    
    If Len(strEditState) > 0 Then labEditState.Caption = strEditState
End Sub


Public Sub AddRepImage(objDcmImg As Object, _
    Optional ByVal lngReleationImageAdviceId As Long = 0, Optional ByVal strFileName As String = "", _
    Optional ByVal blnForceAdd As Boolean = False)
'添加报告图
'lngReleationImageAdviceId:关联图像的医嘱id，有可能该检查查看的图像是从关联的其他检查中打开的

    Dim objCurDicom As DicomImage
    Dim reportTag As TReportImgTag
    
    
    '非编辑状态下，不允许对报告图进行操作
    If mblnIsEditable = False And blnForceAdd = False Then Exit Sub
    
    Set objCurDicom = objDcmImg
    
    reportTag.strKey = objCurDicom.InstanceUID
    
    If strFileName = "" Then
        If lngReleationImageAdviceId = 0 Then
            reportTag.strImgFile = objCurDicom.InstanceUID & ".JPG"
        Else
            reportTag.strImgFile = objCurDicom.InstanceUID & "_" & lngReleationImageAdviceId & ".JPG"
        End If
    Else
        reportTag.strImgFile = strFileName
    End If
    
    reportTag.lngFromAdvice = lngReleationImageAdviceId
    
    objCurDicom.tag = reportTag
 
    Call DrawBorder(objCurDicom, 0)
    
    Call dcmReportImg.Images.Add(objCurDicom)
    Call CalcImgView
    
    If blnForceAdd = False Then Call EnterModify(, True)
End Sub

Public Sub AddRepImgFile(ByVal strFile As String, _
    Optional ByVal lngImageAdviceId As Long = 0, Optional ByVal strFileName As String = "", _
    Optional ByVal blnForceAdd As Boolean = False)
'添加报告图文件
    Dim objCurDicom As DicomImage
    Dim strError As String
    
    Set objCurDicom = ReadDicomFile(strFile, strError, False)
    
    Call AddRepImage(objCurDicom, lngImageAdviceId, strFileName, blnForceAdd)
End Sub

Private Sub DelRepImage()
'删除当前报告图
    Dim strImgKey As String
    Dim strSQL As String
    
    If mlngSelReportImgIndex <= 0 Or mlngSelReportImgIndex > dcmReportImg.Images.Count Then Exit Sub
    
    strImgKey = dcmReportImg.Images(mlngSelReportImgIndex).InstanceUID
    
    '更新数据库
    strSQL = "Zl_影像检查_设置报告图('" & strImgKey & "',2)"
    Call zlDatabase.ExecuteProcedure(strSQL, "删除报告图")
    
    Call dcmReportImg.Images.Remove(mlngSelReportImgIndex)
    
    Call CalcImgView
    
    mlngSelReportImgIndex = 0
     
    Call EnterModify(, True)
    
    RaiseEvent OnDelRepImg(strImgKey)
End Sub


Public Sub Mark(ByVal MarkType As TImgMarkType, Optional ByVal strMark As String = "")
'指定文本标记
    mlngMarkType = MarkType
    mstrMarkText = strMark
End Sub


Public Sub InputWord(ByVal strFreeText As String, _
    ByVal str所见 As String, ByVal str意见 As String, ByVal str建议 As String)
'写入词句
    Dim blnIsUseSpecialty As Boolean
    Dim lngStartSel As Long
    
    blnIsUseSpecialty = False
    If Not mobjSpePlugin Is Nothing Then
        blnIsUseSpecialty = mblnIsSpeState
    End If
    
    If blnIsUseSpecialty = False Then
        If Len(strFreeText) > 0 Then
            If Not mrtbActive Is Nothing Then
                If mrtbActive.Enabled And mrtbActive.Locked = False And mrtbActive.Visible Then
                    lngStartSel = mrtbActive.SelStart
                    
                    mrtbActive.SelText = strFreeText
                    Call SetWordStyle(mrtbActive)
                    
                    mrtbActive.SelStart = lngStartSel + Len(strFreeText)
                    mrtbActive.SetFocus
                End If
            End If
        End If
        
        If Len(str所见) > 0 Then
            If rtb所见.Enabled And rtb所见.Locked = False And rtb所见.Visible Then
                lngStartSel = rtb所见.SelStart
                rtb所见.SelText = str所见
                
                Call SetWordStyle(rtb所见)
                
                rtb所见.SelStart = lngStartSel + Len(str所见)
                
                rtb所见.SetFocus
            Else
                If rtb所见.Visible = False Then MsgboxH GetRootHwnd, "该词句内容仅适用于 [所见] 所在提纲。", vbOKOnly, "提示"
            End If
        End If
        
        If Len(str意见) > 0 Then
            If rtb意见.Enabled And rtb意见.Locked = False And rtb意见.Visible Then
                lngStartSel = rtb意见.SelStart
                rtb意见.SelText = str意见
                
                Call SetWordStyle(rtb意见)
                
                rtb所见.SelStart = lngStartSel + Len(str意见)
                rtb所见.SetFocus
            Else
                If rtb意见.Visible = False Then MsgboxH GetRootHwnd, "该词句内容仅适用于 [意见] 所在提纲。", vbOKOnly, "提示"
            End If
        End If
        
        If Len(str建议) > 0 Then
            If rtb建议.Enabled And rtb建议.Locked = False And rtb建议.Visible Then
                lngStartSel = rtb建议.SelStart
                rtb建议.SelText = str建议
                
                Call SetWordStyle(rtb建议)
                
                rtb所见.SelStart = lngStartSel + Len(str建议)
                rtb所见.SetFocus
            Else
                If rtb建议.Visible = False Then MsgboxH GetRootHwnd, "该词句内容仅适用于 [建议] 所在提纲。", vbOKOnly, "提示"
            End If
        End If
    Else
        Call InputWordToSpecialty(strFreeText, str所见, str意见, str建议)
    End If
End Sub

Private Sub InputWordToSpecialty(ByVal strFreeText As String, _
    ByVal str所见 As String, ByVal str意见 As String, ByVal str建议 As String)
On Error GoTo errhandle
    If mobjSpePlugin Is Nothing Then Exit Sub
    
    Call mobjSpePlugin.InputWord(strFreeText, str所见, str意见, str建议)
Exit Sub
errhandle:
    
End Sub


Public Sub GetReportContext(ByRef str所见 As String, ByRef str意见 As String, ByRef str建议 As String, _
    Optional ByRef strSelText As String = "")
'获取报告内容
    str所见 = rtb所见.Text
    str意见 = rtb意见.Text
    str建议 = rtb建议.Text
    
    If Not mrtbActive Is Nothing Then
        strSelText = mrtbActive.SelText
    End If
End Sub



Private Function GetDkpStateString(ByVal strSourceStateString As String, ByVal strCurStateString As String) As String
    Dim strSourceFmt As String
    Dim arySourcePaneInfo() As String
 
    Dim strCurFmt As String
    Dim aryCurPaneInfo() As String

    Dim i As Long
    Dim strNewPaneFmt As String
    Dim strTitle As String
    Dim strSourcePaneFmt As String
    Dim lngPaneInfoCount As Long
    
    GetDkpStateString = ""

    strSourceFmt = strSourceStateString
    strSourceFmt = Mid(strSourceFmt, InStr(strSourceFmt, "<Pane-1"), 4096)
    strSourceFmt = Mid(strSourceFmt, 1, InStr(strSourceFmt, "</Common>") - 1)
 
    strCurFmt = strCurStateString
    strCurFmt = Mid(strCurFmt, InStr(strCurFmt, "<Pane-1"), 4096)
    strCurFmt = Mid(strCurFmt, 1, InStr(strCurFmt, "</Common>") - 1)

    arySourcePaneInfo = Split(strSourceFmt, "<Pane-")
    aryCurPaneInfo = Split(strCurFmt, "<Pane-")

    lngPaneInfoCount = UBound(arySourcePaneInfo)

    For i = 1 To lngPaneInfoCount
        strSourcePaneFmt = arySourcePaneInfo(i)
        strSourcePaneFmt = Mid(strSourcePaneFmt, InStr(strSourcePaneFmt, "Type="), 255)
        
        strTitle = GetDkpTitleValue(strSourcePaneFmt)
        
        If InStr(strSourcePaneFmt, "Type=""2""") > 0 Then            '
            strNewPaneFmt = strNewPaneFmt & "<Pane-" & i & " " & strSourcePaneFmt
            
        ElseIf InStr(strSourcePaneFmt, "Type=""1""") > 0 Then
            strNewPaneFmt = strNewPaneFmt & "<Pane-" & i & " " & GetDkpReleationFmt(i, i + 1, strSourceFmt, strCurFmt, arySourcePaneInfo, aryCurPaneInfo)
            
        ElseIf InStr(strSourcePaneFmt, "Type=""0""") > 0 Then
            strNewPaneFmt = strNewPaneFmt & "<Pane-" & i & " " & GetDkpNewFmt(strTitle, strCurFmt, i - 1)
            
        Else
            strNewPaneFmt = strNewPaneFmt & "<Pane-" & i & " " & strSourcePaneFmt
            
        End If
    Next
    
    strNewPaneFmt = "<Layout><Common CompactMode=""1"">" & GetDkpSummaryInfo(strSourceStateString) & strNewPaneFmt & "</Common></Layout>"
    
    GetDkpStateString = strNewPaneFmt
End Function

Private Function GetDkpSummaryInfo(ByVal strCurFmt As String) As String
    Dim strTmp As String
    
    strTmp = Mid(strCurFmt, InStr(strCurFmt, "<Summary"), 4096)
    strTmp = Mid(strTmp, 1, InStr(strTmp, "/>") - 1) & "/>"
    
    GetDkpSummaryInfo = strTmp
End Function

Private Function GetDkpReleationFmt(ByVal lngPaneIndex As Long, ByVal lngBindIndex As Long, _
    ByVal strSourceFmt As String, ByVal strCurFmt As String, _
    arySource() As String, aryCur() As String) As String
'根据指定的源pane索引，获取新的格式配置
    
    Dim strReleationTitle As String
    Dim lngReleationIndex As Long
    Dim strFmt As String
    Dim strTmp As String
    Dim strSourcePaneFmt As String
    
    
    GetDkpReleationFmt = ""
    strSourcePaneFmt = arySource(lngPaneIndex)
    
    strFmt = strSourcePaneFmt
    strFmt = Mid(strFmt, InStr(strFmt, "Pane-1=""") + 8, 4096)
    strTmp = Mid(strFmt, 1, InStr(strFmt, """/>") - 1)
    
    lngReleationIndex = Val(strTmp)
    
    
    strSourcePaneFmt = arySource(lngReleationIndex)
    strFmt = Mid(strSourcePaneFmt, InStr(strSourcePaneFmt, "Title=""") + 7, 4096)
    strReleationTitle = Mid(strFmt, 1, InStr(strFmt, """") - 1)
     
    
    
    strFmt = Mid(strCurFmt, InStr(strCurFmt, strReleationTitle), 4096)
    strFmt = Mid(strFmt, 1, InStr(strFmt, "/>") - 1)
    strFmt = Mid(strFmt, InStr(strFmt, "LastHolder=""") + 12, 4096)
    lngReleationIndex = Val(Mid(strFmt, 1, InStr(strFmt, """") - 1))
    
    
    
    strFmt = Mid(aryCur(lngReleationIndex), 3, 4096)
    
    If InStr(strFmt, "Selected=") > 0 Then
        strFmt = Mid(strFmt, 1, InStr(strFmt, "Selected=") - 1)
        
        GetDkpReleationFmt = strFmt & "Selected=""" & lngBindIndex & """ Pane-1=""" & lngBindIndex & """/>"
    Else
        GetDkpReleationFmt = strFmt
    End If
End Function

Private Function GetDkpTitleValue(ByVal strPaneFmt As String) As String
'从格式中获取pane标题
    Dim strTmp As String

    GetDkpTitleValue = ""
    If InStr(strPaneFmt, "Title=") <= 0 Then Exit Function
    
    strTmp = Mid(strPaneFmt, InStr(strPaneFmt, "Title=") + 6, 4096)
    
    GetDkpTitleValue = Mid(strTmp, 1, InStr(strTmp, " ID=") - 1)
End Function


Private Function GetDkpNewFmt(ByVal strTitle As String, ByVal strCurFmt As String, ByVal lngHolderIndex As Long) As String
'根据指定标题获取新的pane格式配置
    Dim strTmp As String
    
    strTmp = Mid(strCurFmt, InStrRev(strCurFmt, "Pane-", InStr(strCurFmt, strTitle)), 4096)
    strTmp = Mid(strTmp, 1, InStr(strTmp, "/>") - 1)
    strTmp = Mid(strTmp, InStr(strTmp, "Type="), 4096)
    
    strTmp = Mid(strTmp, 1, InStr(strTmp, "DockingHolder=") - 1)
    
    strTmp = strTmp & "DockingHolder=""" & lngHolderIndex & """ LastHolder=""" & lngHolderIndex & """/>"
    
    GetDkpNewFmt = strTmp
End Function


Public Function GetLayoutStr() As String
'返回格式字符串[Key=TESTNAME@picturebox1.width:20;picturebox1.height:30;]
    If dkpMain.PanesCount >= 5 Then
        Call dkpMain.DestroyPane(dkpMain.Panes(5))
    End If
    
    GetLayoutStr = "[KEY=EDITOR@" & _
                                        GetProFmt("DKPMAINSTATESTR", GetDkpStateString(dkpMain.tag, dkpMain.SaveStateToString())) & _
                                        GetProFmt("REPORTIMG.WIDTH", dcmReportImg.Width) & _
                                        GetProFmt("MARKIMG.WIDTH", dcmMarkImage.Width) & _
                                 "]"
                                  
End Function

Public Function GetFaceKey() As String
    Dim strKeyTag As String

    strKeyTag = IIf(dkpMain.Panes(1).Closed, "0", "1")

    strKeyTag = strKeyTag & IIf(dkpMain.Panes(2).Closed, "0", "1")

    strKeyTag = strKeyTag & IIf(dkpMain.Panes(3).Closed, "0", "1")

    strKeyTag = strKeyTag & IIf(dkpMain.Panes(4).Closed, "0", "1")
    
    GetFaceKey = strKeyTag
End Function

Public Sub SetLayout(ByVal strLayout As String)
    Dim strPros As String
    Dim strPro As String
    Dim arySourcePane() As paneInfo
    Dim i As Long
    Dim objPane As Pane

    If Len(strLayout) <= 0 Then Exit Sub

    strPros = GetPros(strLayout, "EDITOR")

    strPro = GetProValue(strPros, "DKPMAINSTATESTR")
    If Len(strPro) > 0 Then
        ReDim arySourcePane(dkpMain.PanesCount)
        
        For i = 1 To dkpMain.PanesCount
            arySourcePane(i - 1).ID = dkpMain.Panes(i).ID
            arySourcePane(i - 1).hwnd = dkpMain.Panes(i).Handle
    
            arySourcePane(i - 1).hidden = dkpMain.Panes(i).hidden
            arySourcePane(i - 1).iconid = dkpMain.Panes(i).iconid
            arySourcePane(i - 1).options = dkpMain.Panes(i).options
            arySourcePane(i - 1).tag = dkpMain.Panes(i).tag
            arySourcePane(i - 1).title = dkpMain.Panes(i).title
        Next
  
        Call dkpMain.LoadStateFromString(strPro)
        
        For i = dkpMain.PanesCount To 1 Step -1
    
            dkpMain.Panes(i).ID = arySourcePane(i - 1).ID
            dkpMain.Panes(i).Handle = arySourcePane(i - 1).hwnd
            dkpMain.Panes(i).hidden = arySourcePane(i - 1).hidden
            dkpMain.Panes(i).iconid = arySourcePane(i - 1).iconid
            dkpMain.Panes(i).tag = arySourcePane(i - 1).tag
            dkpMain.Panes(i).title = arySourcePane(i - 1).title
            dkpMain.Panes(i).options = arySourcePane(i - 1).options
        Next
        
'        If Not mobjSpePlugin Is Nothing Then
'            '附加专科报告
'            If dkpMain.PanesCount < 5 Then
'                Set objPane = dkpMain.CreatePane(5, 0, 700, DockBottomOf, dkpMain.Panes(1))
'                objPane.title = "专科录入"
'                objPane.Handle = mobjSpePlugin.hwnd
'                objPane.tag = 4
'                objPane.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'                objPane.Closed = True
'            Else
'                If dkpMain.Panes(5).Handle = 0 Then
'                    dkpMain.Panes(5).title = "专科录入"
'                    dkpMain.Panes(5).Handle = mobjSpePlugin.hwnd
'                    dkpMain.Panes(5).tag = 4
'                    dkpMain.Panes(5).options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'                    dkpMain.Panes(5).Closed = True
'                End If
'            End If
'        End If
      
    End If
    
    If (dcmReportImg.Visible Or Val(dcmReportImg.tag) <> 0) And (dcmMarkImage.Visible Or dcmMarkImage.Images.Count > 0) Then
        strPro = GetProValue(strPros, "REPORTIMG.WIDTH")
        If Val(strPro) > 0 Then dcmReportImg.Width = Val(strPro)
         
        strPro = GetProValue(strPros, "MARKIMG.WIDTH")
        If Val(strPro) > 0 Then dcmMarkImage.Width = Val(strPro)
    End If
End Sub

Public Sub Relayout()
    Dim i As Long
    
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane, Pane5 As Pane
    
'    If dkpMain.PanesCount > 0 Then
'        For i = 1 To dkpMain.PanesCount
'            SetParent dkpMain.Panes(i).Handle, hwnd
'            dkpMain.Panes(i).Handle = 0
'        Next
'    End If

    If dkpMain.PanesCount <= 0 Then
    
        With dkpMain
            .CloseAll
            .DestroyAll
            .options.HideClient = True
            .options.UseSplitterTracker = False '实时拖动
            .options.ThemedFloatingFrames = True
            .options.AlphaDockingContext = True
        End With
        
        '图像
        Set Pane1 = dkpMain.CreatePane(1, 0, 200, DockTopOf, Nothing)
        Pane1.title = "报告图像"
        Pane1.Handle = picImageBack.hwnd
        Pane1.tag = 0 '"REPIMG"
        Pane1.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        
        '描述
        Set Pane2 = dkpMain.CreatePane(2, 0, 400, DockBottomOf, Nothing)   'Pane1
        Pane2.title = mstrDescTitle
        Pane2.Handle = picDesc.hwnd
        Pane2.tag = 1 ' "DESC"
        Pane2.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        '意见
        Set Pane3 = dkpMain.CreatePane(3, 0, 300, DockBottomOf, Pane2)
        Pane3.title = mstrOpinTitle
        Pane3.Handle = picOpin.hwnd
        Pane3.tag = 2 '"OPIN"
        Pane3.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        '建议
        Set pane4 = dkpMain.CreatePane(4, 0, 100, DockBottomOf, Pane3)
        pane4.title = mstrAdviTitle
        pane4.Handle = picAdvi.hwnd
        pane4.tag = 3 '"ADVI"
        pane4.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        
        dkpMain.tag = dkpMain.SaveStateToString
    End If
    
    '配置专科
    If mblnVisibleSpecialty Then
        Call LoadSpecialtyPlugin
        
        If mobjSpePlugin Is Nothing Then
            mblnVisibleSpecialty = False
'        Else
'            '专科报告录入
'            Set Pane5 = dkpMain.CreatePane(5, 0, 700, DockBottomOf, Pane1)
'            Pane5.title = "专科录入"
'            Pane5.Handle = mobjSpePlugin.hwnd
'            Pane5.tag = 4
'            Pane5.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'            Pane5.Closed = True
        End If
    Else
        '如果不配置专科，则隐藏之前可能显示的专科录入界面
        If mblnIsSpeState Then
            Call ChangeSepState(False, True)
        End If
        
        If Not mobjSpePlugin Is Nothing Then SetParent mobjSpePlugin.hwnd, 0
        
        Set mobjSpePlugin = Nothing
    End If
'
'    Call dkpMain.RecalcLayout
'
    
    
'    chkPositive.Visible = Not mblnIgnoreResult
End Sub


Public Sub LocateEditBox()
    Dim objActive As Object

    
    If dkpMain.PanesCount <= 0 Then Exit Sub
    
    '专科报告录入时，则不允许对编辑框进行定位
    If mblnIsSpeState Then
        '...
        Exit Sub
    End If
    
    Set objActive = UserControl.ActiveControl
    
    If objActive Is Nothing Then
        Set objActive = mrtbActive
    Else
        If Not (TypeOf objActive Is RichTextBox) Then Set objActive = mrtbActive
    End If


    If Not objActive Is Nothing Then
        If TypeOf objActive Is RichTextBox Then
            If objActive.Visible And objActive.Locked = False Then objActive.SetFocus
            Exit Sub
        End If
    End If

    If rtb所见.Visible Then
        If rtb所见.Enabled Then
            rtb所见.SetFocus
            rtb所见.SelStart = Len(rtb所见.Text)
        End If
        Exit Sub
    End If

    If rtb意见.Visible Then
        If rtb意见.Enabled Then
            rtb意见.SetFocus
            rtb意见.SelStart = Len(rtb意见.Text)
        End If
        
        Exit Sub
    End If

    If rtb建议.Visible Then
        If rtb建议.Enabled Then
            rtb建议.SetFocus
            rtb建议.SelStart = Len(rtb建议.Text)
        End If
        Exit Sub
    End If
End Sub

Public Sub GetReport(ByRef str所见 As String, ByRef str意见 As String, ByRef str建议 As String)
'插件支持方法，获取当前编辑中录入的报告内容
    str所见 = rtb所见.Text
    str意见 = rtb意见.Text
    str建议 = rtb建议.Text
End Sub

Public Sub ClearReport(ByVal blnClearDesc As Boolean, ByVal blnClearOpin As Boolean, ByVal blnClearAdvi As Boolean)
'插件支持方法，清除当前报告中的文本内容
    If blnClearDesc Then rtb所见.Text = ""
    If blnClearOpin Then rtb意见.Text = ""
    If blnClearAdvi Then rtb建议.Text = ""
End Sub

Public Sub SendReport(ByVal str所见 As String, ByVal str意见 As String, ByVal str建议 As String)
'插件支持方法，发送专科报告中录入的文本内容
    rtb所见.Text = str所见
    rtb意见.Text = str意见
    rtb建议.Text = str建议
End Sub

Private Function LoadSpecialtyPlugin() As Boolean
'载入专科报告插件
    Dim objParent As Object
    Dim strErr As String
    
On Error GoTo errhandle
    LoadSpecialtyPlugin = False
    If mblnVisibleSpecialty = False Then Exit Function
    
    Set mobjSpePlugin = DynamicCreate("ZLPacsProReport.clsZLPacsProReport", "专科录入")
     
     If mobjSpePlugin Is Nothing Then Exit Function
     
    Set objParent = UserControl.Extender
    Call mobjSpePlugin.InitPlugin(gcnOracle, objParent)
    
    LoadSpecialtyPlugin = True
Exit Function
errhandle:
    mblnVisibleSpecialty = False
    strErr = err.Description
    
    MsgboxH GetRootHwnd, "专科报告初始化失败：" & strErr, vbOKOnly, "提示"
End Function

Private Sub ResizeEdit(rtbEdit As RichTextBox, picParent As PictureBox)
    rtbEdit.Left = 0
    rtbEdit.Top = 0
    rtbEdit.Width = picParent.Width
    rtbEdit.Height = picParent.Height
End Sub


Private Sub dcmMarkImage_DblClick()
'TASK:暂时不支持标记图的其他高级处理
''打开标记图处理
'    If dcmMarkImage.Images.Count <> 1 Then Exit Sub
'
'
'    Call ShowMarkImgProcess
End Sub

'Public Sub ShowMarkImgProcess()
'    Dim i As Long
'    Dim objDcmImg As DicomImage
'    Dim aryNull() As Object
'
'    If mobjMarkProcessV2 Is Nothing Then Set mobjMarkProcessV2 = New frmImageProcessV2
'
'    Set objDcmImg = dcmMarkImage.Images(1)
'
'    '在没有任何处理下，2秒后自动关闭大图预览
'    Call mobjMarkProcessV2.ZlShowMe(mObjNotify.Owner, mlngAdviceID, objDcmImg, aryNull, ptMark, 2, False)
'End Sub


Private Sub dcmMarkImage_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim lngFrame As Long
    
    If dcmMarkImage.Images.Count <= 0 Then Exit Sub
    If mlngMarkType = imtNormal Then
        '不进行标记处理
        picImageBack.MousePointer = 0
        picImageBack.MouseIcon = Nothing
        
        Exit Sub
    End If
    
    lngFrame = 2
    
    '设置鼠标
    If dcmMarkImage.ImageXPosition(X, Y) > lngFrame And dcmMarkImage.ImageXPosition(X, Y) < dcmMarkImage.Images(1).SizeX - lngFrame _
       And dcmMarkImage.ImageYPosition(X, Y) > lngFrame And dcmMarkImage.ImageYPosition(X, Y) < dcmMarkImage.Images(1).SizeY - lngFrame Then
        picImageBack.MousePointer = 99
        picImageBack.MouseIcon = listCur.ListImages("pen").Picture
        
        SetCapture dcmMarkImage.hwnd
    Else
        ReleaseCapture
        
        picImageBack.MousePointer = 0
        picImageBack.MouseIcon = Nothing
    End If
End Sub

Public Sub AddNumber()
'给文本段添加前导的数字序号
'mintReportViewType 0-检查所见CheckView，1-诊断意见Result，2-建议Advice

    Dim rText As RichTextBox
    Dim strtext As String
    Dim iCount As Integer
    Dim iStart As Integer
    
On Error GoTo err
 
    Set rText = mrtbActive
    
    If rText Is Nothing Then
        MsgboxH GetRootHwnd, "请选择需要编号的报告编辑框。", vbOKOnly, "信息提示"
        Exit Sub
    End If
    
    strtext = rText.Text
    
    '先判断文本段是否被锁定
    If rText.Locked = True Then
        MsgboxH GetRootHwnd, "文本段被锁定不允许编辑。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    '先判断该文本段中第一个字符是否数字1，如果是，则提示已经有数字编号，是否还要添加
    If Left(strtext, 1) = "1" Then
        If MsgboxH(GetRootHwnd, "本段文本中已经包含数字编号，是否还要添加数字编号？", vbOKCancel, "提示") = vbCancel Then
            Exit Sub
        End If
    End If
    
    '开始添加数字编号,每一个回车之后，如果不是空格，就添加序号
    iStart = 1
    
    '第一行也需要判断是否存在缩进
    If Left(strtext, 1) <> " " Then
        iCount = 1
        strtext = iCount & ". " & strtext
    Else
        iCount = 0
    End If
    iStart = InStr(iStart, strtext, vbCrLf)
    
    While (iStart <> 0)
        If Mid(strtext, iStart + 2, 1) <> " " And Mid(strtext, iStart + 2, 2) <> vbCrLf And Mid(strtext, iStart + 2, 1) <> "" Then
            iCount = iCount + 1
            strtext = Left(strtext, iStart + 1) & iCount & ". " & Right(strtext, Len(strtext) - iStart - 1)
        End If
        iStart = InStr(iStart + 1, strtext, vbCrLf)
    Wend
    
    rText.Text = strtext
    
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub AddCurMarks(ByVal X As Long, ByVal Y As Long)
    Dim objPicMarks As New clsPicMarks
    Dim lTemp As DicomLabel
    Dim lngBound As Long
    Dim dblMarkZoom As Double
    Dim ImgTag As TReportImgTag
    Dim strMaxOrder As String
    
    objPicMarks.对象属性 = dcmMarkImage.Images(1).tag.strImgMarks
    
    strMaxOrder = 0
    If objPicMarks.Count > 0 Then
        strMaxOrder = objPicMarks.Item(objPicMarks.Count).内容
    End If
    
    '画标注
    '两种类型的标注，一种是直接自动编号，另一种是手工编号
    lngBound = objPicMarks.Count + 1
    
    objPicMarks.Add lngBound
    objPicMarks(lngBound).Selected = False
    
    If IsNumeric(mstrMarkText) Or Len(mstrMarkText) <= 0 Then
        objPicMarks(lngBound).类型 = 6     '圆形编号
    Else
        objPicMarks(lngBound).类型 = 0      '0-表示文本
    End If
        
    If mlngMarkType = imtAuto Then
        objPicMarks(lngBound).内容 = Val(strMaxOrder) + 1
    ElseIf mlngMarkType = imtSpecify Then
        objPicMarks(lngBound).内容 = mstrMarkText
    Else
        Exit Sub
    End If
    
    '点集没有留空
    Set lTemp = New DicomLabel
    lTemp.Left = X
    lTemp.Top = Y
    lTemp.Width = 20
    lTemp.Height = 20
    lTemp.ImageTied = True
    lTemp.Rescale dcmMarkImage.Images(1)
    
    dblMarkZoom = dcmMarkImage.Images(1).SizeX / Val(GetReportImagePro(dcmMarkImage.Images(1).tag.strPros, "width")) * Screen.TwipsPerPixelX
    
    objPicMarks(lngBound).X1 = lTemp.Left / dblMarkZoom
    objPicMarks(lngBound).Y1 = lTemp.Top / dblMarkZoom
    objPicMarks(lngBound).X2 = objPicMarks(lngBound).X1
    objPicMarks(lngBound).Y2 = objPicMarks(lngBound).Y1
    objPicMarks(lngBound).填充色 = glngColor(lngBound Mod 9 + 1)
    objPicMarks(lngBound).填充方式 = -2
    '线条色留空，字体色留空
    objPicMarks(lngBound).线型 = 1
    objPicMarks(lngBound).线宽 = 1
    
    Set objPicMarks(lngBound).字体 = New StdFont '  "宋体"
    
    Call DrawMarks(dcmMarkImage.Images(1), objPicMarks, dblMarkZoom)
    
    ImgTag = dcmMarkImage.Images(1).tag
    ImgTag.strImgMarks = objPicMarks.对象属性
    
    dcmMarkImage.Images(1).tag = ImgTag

    Call EnterModify(, , True)
End Sub

Private Sub dcmMarkImage_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim lTemp As DicomLabel
    Dim strNum As Integer
    Dim objDcmLabs As DicomLabels
    Dim strErr As String

On Error GoTo errhandle
    
    If dcmMarkImage.Images.Count <= 0 Then Exit Sub
    If mblnIsEditable = False Then Exit Sub
    
    If Button = 2 Then
        Set objDcmLabs = dcmMarkImage.LabelHits(X, Y, False, False, True)
        If objDcmLabs.Count > 0 Then
            menuLab.tag = dcmMarkImage.Images(1).Labels.IndexOf(objDcmLabs.Item(objDcmLabs.Count))
            PopupMenu menuLab, 2
            
        End If
        
        Exit Sub
    End If

    If mlngMarkType = imtNormal Then Exit Sub
    If Button = 1 And picImageBack.MousePointer = 99 Then
        '画标注
        Call AddCurMarks(X, Y)
    End If
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub


Private Sub dcmReportImg_Click()
    Dim strErr As String
On Error GoTo errhandle
    If dcmReportImg.Images.Count <= 0 Then Exit Sub
    If mlngSelReportImgIndex <= 0 Then Exit Sub

    picReportImgOper.Left = (dcmReportImg.Width - picReportImgOper.Width) / 2
    picReportImgOper.Top = dcmReportImg.Height - picReportImgOper.Height

    picReportImgOper.Visible = mblnIsEditable
      
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub


Private Sub cmdOper_Click(Index As Integer)
    Dim strErr As String
On Error GoTo errhandle
    Select Case Index
        Case 0  '删除报告
            Call DelRepImage
            
        Case 1  '前移
            If mlngSelReportImgIndex <= 1 Then
'                MsgboxH GetRootHwnd, "不能向前移动。", vbOKOnly, "提示"
                Exit Sub
            End If
            
            dcmReportImg.Images.Move mlngSelReportImgIndex, mlngSelReportImgIndex - 1
            
            DrawBorder dcmReportImg.Images(mlngSelReportImgIndex - 1), 0
            
            mblnIsModifyImage = True
        Case 2  '后移
            If mlngSelReportImgIndex >= dcmReportImg.Images.Count Then
'                MsgboxH GetRootHwnd, "不能向后移动。", vbOKOnly, "提示"
                Exit Sub
            End If
            
            dcmReportImg.Images.Move mlngSelReportImgIndex, mlngSelReportImgIndex + 1
            
            DrawBorder dcmReportImg.Images(mlngSelReportImgIndex + 1), 0
            
            mblnIsModifyImage = True
        Case 3  '自动
            Call Mark(imtAuto)
            
        Case 4, 5, 6, 7 '标记1,2,3,4
            Call Mark(imtSpecify, Index - 3)
            
    End Select
    
    mlngSelReportImgIndex = 0
    
    picReportImgOper.Visible = False
'    picMarkImgOper.Visible = False
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub
 
 

Private Sub dcmReportImg_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim blnIsReportImgArea As Boolean
    Dim lngBound As Long
    
    If dcmReportImg.Images.Count <= 0 Then Exit Sub
    If mlngSelReportImgIndex <= 0 Then Exit Sub

    blnIsReportImgArea = False
    lngBound = 135
    
    '判断是否需要显示图像
    If (lngBound <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= dcmReportImg.Width - lngBound) And _
       (lngBound <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= dcmReportImg.Height - lngBound) Then
        blnIsReportImgArea = True
    End If

    picReportImgOper.Visible = blnIsReportImgArea And mblnIsEditable
    
    '注意：该事件中，不能对鼠标进行锁定，或者会造成显示的报告图相关处理按钮不能操作
    
End Sub

Private Sub dcmReportImg_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    Dim strErr As String
On Error GoTo errhandle
    If Button = 2 Then
        '鼠标右键
        If mblnIsEditable Then
            If dcmReportImg.Images.Count <= 0 Then Exit Sub
            PopupMenu menuReport, 2
        End If
    Else
        mlngSelReportImgIndex = dcmReportImg.ImageIndex(X, Y)
        
        If mlngSelReportImgIndex <= 0 Or mlngSelReportImgIndex > dcmReportImg.Images.Count Then Exit Sub
        
        For i = 1 To dcmReportImg.Images.Count
            Call DrawBorder(dcmReportImg.Images(i), 0)
        Next
            
        Call DrawBorder(dcmReportImg.Images(mlngSelReportImgIndex), ColorConstants.vbRed, True)
    End If
    
'    RaiseEvent OnMouseUp(Button, Shift, x, y)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
On Error Resume Next
    Call HideCharInput
End Sub

Private Sub dkpMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
    Call HideCharInput
End Sub

Private Sub HideCharInput()
'隐藏字符录入
    Dim blnHide As Boolean
    
    If mrtbActive Is Nothing Then Exit Sub

    blnHide = False
    If Not ActiveControl Is Nothing Then
        Select Case ActiveControl.hwnd
            Case picDesc.hwnd
                If rtb所见.Visible And rtb所见.Locked = False Then
                    rtb所见.SetFocus
                Else
                    blnHide = True
                End If
            Case picOpin.hwnd
                If rtb意见.Visible And rtb意见.Locked = False Then
                    rtb意见.SetFocus
                Else
                    blnHide = True
                End If
            Case picAdvi.hwnd
                If rtb建议.Visible And rtb建议.Locked = False Then
                    rtb建议.SetFocus
                Else
                    blnHide = True
                End If
            Case picImageBack.hwnd
                blnHide = True
        End Select
    Else
        blnHide = mrtbActive.Locked
    End If
    
    If blnHide Then picChar.Visible = False
End Sub

Private Sub menuLab_Del_Click()
'删除标注\
    Dim strErr As String
On Error GoTo errhandle
    Dim objLab As DicomLabel
    Dim strDelTag As String
    Dim strMarks As String
    Dim objReportImgTag As TReportImgTag
    Dim i As Long
    Dim aryMark() As String
    Dim strRemoveMarks  As String
    Dim objLinkLab As DicomLabel
    
    If menuLab.tag = "" Then Exit Sub
    If dcmMarkImage.Images.Count <= 0 Then Exit Sub
    
    Set objLab = dcmMarkImage.Images(1).Labels(menuLab.tag)
    If objLab Is Nothing Then Exit Sub
    
    strDelTag = ""
    
    If Not objLab.TagObject Is Nothing Then
        strDelTag = objLab.TagObject.Text
        Set objLinkLab = objLab.TagObject
    End If
    
    If strDelTag = "" Then strDelTag = objLab.Text
    
    Call dcmMarkImage.Images(1).Labels.Remove(menuLab.tag)
    
    If Not objLinkLab Is Nothing Then
        Call dcmMarkImage.Images(1).Labels.Remove(dcmMarkImage.Images(1).Labels.IndexOf(objLinkLab))
    End If
    
    strMarks = dcmMarkImage.Images(1).tag.strImgMarks
    
    strRemoveMarks = ""
    If Len(strMarks) > 0 Then
        aryMark = Split(strMarks, "0|6|")
        For i = 0 To UBound(aryMark)
            If aryMark(i) <> "" Then
                If Val(Mid(aryMark(i), 1, 2)) = Val(strDelTag) Then
                    strRemoveMarks = "0|6|" & aryMark(i)
                    Exit For
                End If
            End If
        Next
    End If
    
    objReportImgTag = dcmMarkImage.Images(1).tag
    strMarks = Replace(strMarks, strRemoveMarks, "")
    
    If Len(strMarks) > 0 Then
        If Right(strMarks, 2) = "||" Then
            strMarks = Mid(strMarks, 1, Len(strMarks) - 2)
        End If
    End If
    objReportImgTag.strImgMarks = strMarks
    
  
    
    dcmMarkImage.Images(1).tag = objReportImgTag
    
    Call dcmMarkImage.Images(1).Refresh(False)
    
    menuLab.tag = ""
    
    Call EnterModify(, , True)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub menuReport_Del_Click()
    Call cmdOper_Click(0)
End Sub

Private Sub menuReport_Last_Click()
    Call cmdOper_Click(1)
End Sub

Private Sub menuReport_Next_Click()
    Call cmdOper_Click(2)
End Sub

Private Sub mobjMarkProcessV2_OnSaveImage(ByVal emImageType As TImageType, dcmImage As DicomObjects.DicomImage)
'    Dim reportImgTag As TReportImgTag
'
'    If emImageType <> mtTagImage Then Exit Sub
'    If dcmMarkImage.Images.Count <= 0 Then Exit Sub
'
'    reportImgTag = dcmMarkImage.Images(1).tag
'
'    dcmMarkImage.Images.Clear
'
'    reportImgTag.strImgMarks = ""
'    dcmImage.tag = reportImgTag
'
'    Call DrawBorder(dcmImage, 0)
'
'    dcmMarkImage.Images.Add dcmImage
'    dcmMarkImage.Images(1).tag = reportImgTag
'
'    mblnIsModifyMarks = True
End Sub

Private Sub picDesc_Resize()
On Error Resume Next
    Call ResizeEdit(rtb所见, picDesc)
    
    If rtb所见.Visible And rtb所见.Locked = False Then
        Call SyncWordChar
    End If
End Sub

Private Sub picOpin_Resize()
On Error Resume Next
    Call ResizeEdit(rtb意见, picOpin)
    
    If rtb意见.Visible And rtb意见.Locked = False Then
        Call SyncWordChar
    End If
End Sub


Private Sub picAdvi_Resize()
On Error Resume Next
    Call ResizeEdit(rtb建议, picAdvi)
    
    If rtb建议.Visible And rtb建议.Locked = False Then
        Call SyncWordChar
    End If
End Sub


Private Sub SyncWordChar()
'显示字符录入栏
    Dim p As POINTAPI
    Dim p2rect As RECT
    Dim objPic As PictureBox
    Dim strOutlineTitle As String
    
    If mrtbActive Is Nothing Then Exit Sub
    
    Select Case mrtbActive.Name
        Case "rtb所见"
            Set objPic = picDesc
            strOutlineTitle = mstrDescTitle
        Case "rtb意见"
            Set objPic = picOpin
            strOutlineTitle = mstrOpinTitle
        Case "rtb建议"
            Set objPic = picAdvi
            strOutlineTitle = mstrAdviTitle
    End Select
    
    GetWindowRect objPic.hwnd, p2rect
    
    p.X = 0
    p.Y = p2rect.Top
    
    ScreenToClient UserControl.hwnd, p
    p.Y = ScaleY(p.Y, vbPixels, vbTwips)
    
    picChar.Left = Len(strOutlineTitle) * (TextWidth("啊") + 90)
    picChar.Top = p.Y - objPic.Top + 10 - 15
    picChar.Width = rtb所见.Width - picChar.Left
    picChar.Height = TextHeight("啊") + 120
    
    picChar.Visible = True
End Sub

Private Sub picImageBack_Resize()

On Error Resume Next
    
    If ucSplitter1.Visible Then
        Call ucSplitter1.RePaint(False)
    Else
        If dcmReportImg.Visible Or Val(dcmReportImg.tag) <> 0 Then
            dcmReportImg.Width = picImageBack.Width
            dcmReportImg.Height = picImageBack.Height
        End If
    End If
  
    Call CalcImgView
Exit Sub
errhandle:
'    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
 
End Sub


Private Sub CalcImgView()
    Dim iCols As Integer, iRows As Integer
    
    If dcmReportImg.Images.Count = 1 Then
        dcmReportImg.MultiColumns = 1
        dcmReportImg.MultiRows = 1
    
        Exit Sub
    End If
    
On Error Resume Next
      
    '调整图像显示布局
    ResizeRegion dcmReportImg.Images.Count, dcmReportImg.Width, dcmReportImg.Height, iRows, iCols

    dcmReportImg.MultiColumns = iCols
    dcmReportImg.MultiRows = iRows
    
    If dcmReportImg.Images.Count > 0 Then
        dcmReportImg.CurrentIndex = 1
    Else
        dcmReportImg.CurrentIndex = 0
    End If
End Sub

Private Sub EnterModify(Optional ByVal blnIsTextModify As Boolean = False, _
    Optional ByVal blnIsImageModify As Boolean = False, Optional ByVal blnIsMarkModify As Boolean = False)
'设置修改状态
    Dim strMsg As String
    
    If mlngReportID = 0 And (blnIsTextModify Or blnIsImageModify Or blnIsMarkModify) Then
        If LockEditor(strMsg) = False Then
            '锁定失败的处理
            ResetContext
            'ResetEditState
            Call ConfigFaceState
            
            MsgboxH GetRootHwnd, strMsg, vbOKOnly, "提示"
            
            
            
            Exit Sub
        End If
    End If
    
    If blnIsTextModify Then mblnIsModifyText = blnIsTextModify
    If blnIsImageModify Then mblnIsModifyImage = blnIsImageModify
    If blnIsMarkModify Then mblnIsModifyMarks = blnIsMarkModify
    

End Sub


Private Sub rtb建议_Change()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsLoadData = False Then Exit Sub
    If mblnIsEditable = False Then Exit Sub
    
    Call EnterModify(True)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub
 

Private Sub rtb建议_DblClick()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsEditable = False Then Exit Sub
    If Trim(rtb建议.Text) = "" Then Exit Sub
    
    timerTmp.Enabled = True
'    Call richTextBoxShowElements(rtb建议, Parent)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub rtb建议_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = 2 Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub rtb建议_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub rtb所见_Change()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsLoadData = False Then Exit Sub
    If mblnIsEditable = False Then Exit Sub
    
    Call EnterModify(True)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub rtb所见_DblClick()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsEditable = False Then Exit Sub
    If Trim(rtb所见.Text) = "" Then Exit Sub
    
    timerTmp.Enabled = True
'    Call richTextBoxShowElements(rtb所见, Parent)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub rtb所见_GotFocus()
On Error Resume Next
    Set mrtbActive = rtb所见
    
    If mrtbActive.Visible And mrtbActive.Locked = False Then
        Call SyncWordChar
    Else
        picChar.Visible = False
    End If
    
    RaiseEvent OnOutlineChange(otDesc)
End Sub
 

Private Sub rtb所见_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = 2 Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub rtb所见_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub rtb意见_Change()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsLoadData = False Then Exit Sub
    If mblnIsEditable = False Then Exit Sub
    
    Call EnterModify(True)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub rtb意见_DblClick()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsEditable = False Then Exit Sub
    If Trim(rtb意见.Text) = "" Then Exit Sub
    
    timerTmp.Enabled = True
'    Call richTextBoxShowElements(rtb意见, Parent)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub rtb意见_GotFocus()
On Error Resume Next
    Set mrtbActive = rtb意见
    
    If mrtbActive.Visible And mrtbActive.Locked = False Then
        Call SyncWordChar
    Else
        picChar.Visible = False
    End If
    
    RaiseEvent OnOutlineChange(otOpin)
End Sub

Private Sub rtb建议_GotFocus()
On Error Resume Next
    Set mrtbActive = rtb建议
    
    If mrtbActive.Visible And mrtbActive.Locked = False Then
        Call SyncWordChar
    Else
        picChar.Visible = False
    End If
    
    RaiseEvent OnOutlineChange(otAdvi)
End Sub

Private Sub rtb意见_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = 2 Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub rtb意见_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub timerTmp_Timer()
On Error GoTo errhandle
    timerTmp.Enabled = False
    If mrtbActive Is Nothing Then Exit Sub
    
    Call richTextBoxShowElements(mrtbActive, Parent)
    
Exit Sub
errhandle:
    timerTmp.Enabled = False
End Sub

Private Sub UserControl_Initialize()
    mblnIsLoadData = False
     
    mblnVisibleSpecialty = True     '调试
    mblnUseImgSign = False
    
    mstrDescTitle = "检查所见"
    mstrOpinTitle = "诊断意见"
    mstrAdviTitle = "建    议"
    
    mlngSelReportImgIndex = 0
    mintEditFontSize = 0
    mblnTechReptSame = False
    
    Set mobjSpePlugin = Nothing
End Sub

 
Private Sub UserControl_InitProperties()
On Error GoTo errhandle
    Call Relayout
Exit Sub
errhandle:

End Sub

Private Sub UserControl_Paint()
On Error GoTo errhandle
    If mblnIsLoadData = False Then
        'TODO:载入数据...
    End If
Exit Sub
errhandle:

End Sub

Private Sub UserControl_Resize()
On Error GoTo errhandle
    picContainer.Left = 0
    picContainer.Top = 0
    picContainer.Width = ScaleWidth
    picContainer.Height = ScaleHeight - picState.Height
    
    picState.Left = 0
    picState.Top = picContainer.Height
    picState.Width = ScaleWidth
    
    labFmt.Width = picState.ScaleWidth - labFmt.Left
    
    labSign.Width = picState.ScaleWidth - labSign.Left
    labEditState.Left = picState.ScaleWidth - labEditState.Width - 50
Exit Sub
errhandle:
    Debug.Print "UserControl_Resize:" & err.Description
End Sub

Public Sub Destory()

    ucSplitter1.Destory
    
    If Not mobjMarkProcessV2 Is Nothing Then
        Unload mobjMarkProcessV2
    End If
    
    Set mobjMarkProcessV2 = Nothing
    Set mobjSpePlugin = Nothing
    Set mObjNotify = Nothing
    Set mrtbActive = Nothing
End Sub


Private Sub UserControl_Terminate()
On Error GoTo errhandle

    SetParent picImageBack.hwnd, 0
    SetParent picDesc.hwnd, 0
    SetParent picOpin.hwnd, 0
    SetParent picAdvi.hwnd, 0
    
    If Not mobjSpePlugin Is Nothing Then SetParent mobjSpePlugin.hwnd, 0
    
    dkpMain.CloseAll
    
    Call Destory
Exit Sub
errhandle:
    Debug.Print "ucReportEditor_Terminate Err:" & err.Description
End Sub

Private Function IsHaveContent() As Boolean
'返回:报告是否存在内容
On Error GoTo errH
    IsHaveContent = True
    If rtb所见.Text = "" And rtb建议.Text = "" And rtb意见.Text = "" Then IsHaveContent = False
    
    Exit Function
errH:
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Function
 
