VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.UserControl UcQueue 
   ClientHeight    =   6034
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10598
   ScaleHeight     =   6034
   ScaleWidth      =   10598
   ToolboxBitmap   =   "UcQueue.ctx":0000
   Begin VB.PictureBox picPlace 
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   315
      ScaleHeight     =   28
      ScaleWidth      =   42
      TabIndex        =   15
      Top             =   2085
      Width           =   45
   End
   Begin VB.Timer timerCard 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   900
      Top             =   330
   End
   Begin VB.PictureBox picCallFace 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   5535
      ScaleHeight     =   4452
      ScaleWidth      =   3738
      TabIndex        =   8
      Top             =   855
      Width           =   3735
      Begin XtremeReportControl.ReportControl rptCallList 
         Height          =   4095
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   7223
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnSort =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption scCallInf 
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "呼叫列表："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   8421504
      End
   End
   Begin VB.PictureBox picQueueFace 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   495
      ScaleHeight     =   4452
      ScaleWidth      =   4452
      TabIndex        =   1
      Top             =   915
      Width           =   4455
      Begin XtremeReportControl.ReportControl rptQueueList 
         Height          =   4455
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   4260
         _Version        =   589884
         _ExtentX        =   7514
         _ExtentY        =   7858
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnSort =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CheckBox optOutQueue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3450
         TabIndex        =   14
         Top             =   45
         Width           =   195
      End
      Begin VB.CheckBox optOutQueue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2700
         TabIndex        =   13
         Top             =   45
         Width           =   195
      End
      Begin VB.CheckBox optOutQueue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1905
         TabIndex        =   12
         Top             =   45
         Width           =   195
      End
      Begin VB.CheckBox optOutQueue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   11
         Top             =   45
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Label lblQueueFilter 
         BackStyle       =   0  'Transparent
         Caption         =   "排队中"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1170
         TabIndex        =   7
         Top             =   30
         Width           =   750
      End
      Begin VB.Label lblQueueFilter 
         BackStyle       =   0  'Transparent
         Caption         =   "已完成"
         Height          =   255
         Index           =   3
         Left            =   3645
         TabIndex        =   6
         Top             =   30
         Width           =   750
      End
      Begin VB.Label lblQueueFilter 
         BackStyle       =   0  'Transparent
         Caption         =   "已弃号"
         Height          =   255
         Index           =   2
         Left            =   2865
         TabIndex        =   5
         Top             =   30
         Width           =   750
      End
      Begin VB.Label lblQueueFilter 
         BackStyle       =   0  'Transparent
         Caption         =   "已暂停"
         Height          =   255
         Index           =   1
         Left            =   2130
         TabIndex        =   4
         Top             =   30
         Width           =   750
      End
      Begin XtremeSuiteControls.ShortcutCaption scQueueInf 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "排队列表："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16744576
         GradientColorDark=   16761024
      End
   End
   Begin VB.TextBox txtLocateValue 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12.24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7965
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Timer tmrBroadCast 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   120
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   600
      Top             =   45
      _Version        =   589884
      _ExtentX        =   610
      _ExtentY        =   610
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "UcQueue.ctx":0312
      Left            =   3840
      Top             =   240
      _Version        =   589884
      _ExtentX        =   432
      _ExtentY        =   406
      _StockProps     =   0
   End
End
Attribute VB_Name = "UcQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'默认排序号码格式长度
Private Const M_LNG_FORMAT_ORDER_LEN As Long = 20

Private Const M_LNG_ICON_QUEUEING As Long = 8264 '807         '排队中
Private Const M_LNG_ICON_DIAGNOSE As Long = 3009        '接诊中
Private Const M_LNG_ICON_CALLING As Long = 745          '呼叫中
Private Const M_LNG_ICON_CALLED As Long = 732          '已呼叫

Private Const M_STR_ALL_POPEDOM As String = "[[#ALLPOPEDOM#]]"


'排队列表的当前选择状态
Public Enum TQueueFromType
    qftWaitQueue = 0
    qftCalledQueue = 1
    qftFindQueue = 2
End Enum


Public Enum TQueueSelState
    qss排队中 = 0
    qss已暂停 = 1
    qss已弃号 = 2
    qss已完成 = 3
End Enum


'菜单关键子定义
Public Enum TMenuId
    mi打号 = conMenu_Queue_PrintNumber
    mi顺呼 = conMenu_Queue_CallNext
    mi直呼 = conMenu_Queue_CallThis
    mi广播 = conMenu_Queue_Broadcast
    
    mi插队 = conMenu_Queue_InsertQueue

    mi重排 = conMenu_Queue_RestartQueue
    
    mi暂停 = conMenu_Queue_Pause
    mi弃号 = conMenu_Queue_Abandon
    mi恢复 = conMenu_Queue_Restore
    
    mi接诊 = conMenu_Queue_RecDiagnose
    mi完成 = conMenu_Queue_Finaled
    
    mi定位 = conMenu_Queue_Locate
    mi查找 = conMenu_Queue_Find
    
    mi修改 = conMenu_Queue_Update
    mi刷新 = conMenu_Queue_Refresh
    
    mi设置 = conMenu_Queue_Setup
        
End Enum


Private WithEvents mobjMsg As clsQueueMsgCenter
Attribute mobjMsg.VB_VarHelpID = -1

Private mobjQueueManage As clsQueueOperation
Attribute mobjQueueManage.VB_VarHelpID = -1
Private mcnOracle As ADODB.Connection
Private mobjOwner As Object
Private mstrProTag As String


Private mblnIsSelectedCallingList As Boolean    '是否选择已叫号队列
Private mintWorkType As Integer                 '业务类型
Private mstrPrivs As String                     '权限字符串

Private mstrCustomOrderColName As String        '自定义排序字段   作用于ReportControl控件
Private mblnIsShowBars As Boolean               '是否显示工具栏
Private mblnIsShowCalledQueue As Boolean        '是否显示已呼叫队列

Private mblnAutoComplete As Boolean             '自动完成已接诊队列
Private mblnShowMySelfCalled As Boolean         '仅显示自己呼叫的队列数据
Private mblnIsReleationQueueTag As Boolean      '是否关联排队标记,为true时 界面上显示的排队号码为 排队标记+排队号码

Private mstrFindWay As String

Private mblnInitOk As Boolean                   '是否初始化完成
Private mstrLoginUserName As String             '登录用户名

Private mstrLocateType As String                '定位类型
Private mlngLocateRowIndex As Long

Private mstrDataFields As String                '该字段如果不在显示字段列中，则自动隐藏
Private mstrDisplayQueueFields As String        '排队列表列名串
Private mstrDisplayCallFields As String         '呼叫列表列名串
Private mstrReason As String                    '插队原因

Private mstrQueryQueueNames As String           '要查询显示的队列名称
Private mstrGroupField As String                '分组名 如为空则不进行分组
Private mstrLastFixedQueue As String        '最后的分组名称

Private mlngInterval As Long                    '轮训时间

Private mrsVoiceContext As ADODB.Recordset  '待播放的语音数据集
Private mstrComputerName As String              '本地计算机名称
Private mdtLastVoiceDate As Date

'Private mlngQueueW1 As Long     '队列显示宽度
'Private mlngQueueW2 As Long     '呼叫队列显示宽度


Private mblnIsFindQueue As Boolean   '是否为查找队列

Private mlngMenuCaptionStyle As Long


'与刷卡相关变量定义
Private mlngReadCount As Long
Private mlngStartTime As Long
Private mlngAvgTime As Long



'公共事件
Public Event OnCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As TCallWay, ByRef strCallContext As String, blnCancel As Boolean)
Public Event OnCallPreAfter(ByVal lngQueueId As Long, ByVal lngCallWay As TCallWay)

Public Event OnPlayVoiceBefore(ByVal lngCallId As Long, ByVal lngQueueId As Long, ByVal strCallContext As String, ByRef blnCancel As Boolean)
Public Event OnPlayVoiceAfter(ByVal lngCallId As Long, ByVal lngQueueId As Long, ByVal strCallContext As String)

Public Event OnWorkBefore(ByVal lngListType As TQueueFromType, ByVal lngListRow As Long, ByVal lngQueueId As Long, ByVal lngOperationType As TOperationType, blnCancel As Boolean)
Public Event OnWorkAfter(ByVal lngQueueId As Long, ByVal strCurQueueName As String, ByVal lngOperationType As TOperationType)

Public Event OnReadBefore(rsDataRow As ADODB.Recordset, ByVal lngListType As TQueueFromType, blnCancel As Boolean)
Public Event OnReadAfter(rsDataRow As ADODB.Recordset, ByVal lngListType As TQueueFromType, objReportRecord As Object)

Public Event OnCreateQueueNo(ByVal lngQueueId As Long, ByVal strQueueName As String, ByRef strQueueNo As String)

'查询排队队列数据事件
Public Event OnQueryQueueData(rsData As ADODB.Recordset, blnUseCustom As Boolean)

'查找队列数据时触发此事件
Public Event OnFindData(ByVal strFindWay As String, ByVal strFindValue As String, txtFind As Object, rsData As ADODB.Recordset, ByRef blnUseCustom As Boolean)
    
'定位队列数据时触发此事件
Public Event OnLocateData(ByVal strLocateWay As String, ByVal strLocateValue As String, txtFind As Object, ByRef lngQueueId As Long, ByRef blnUseCustom As Boolean)
    
Public Event OnSelectionChanged(ByVal lngListType As TQueueFromType, ByVal lngQueueId As Long, objQueueList As Object, objReportRow As Object)
Public Event OnItemDblClick(ByVal lngListType As TQueueFromType, ByVal lngQueueId As Long, objReoprtRow As Object, objReportItem As Object)
Public Event OnQueueListChange(ByVal lngListType As TQueueFromType, objQueueList As Object)


Public Event OnCmdBarInit(objCommandBar As Object)
Public Event OnCmdBarUpdate(objComandBarControl As Object)
Public Event OnCmdBarExecute(objComandBarControl As Object, ByRef blnUseCustom As Boolean)

Public Event OnColumnInit(objQueueList As Object, objReportColumn As Object)

Public Event OnQueueListMouseDown(ByVal lngListType As TQueueFromType, Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnQueueListMouseUp(ByVal lngListType As TQueueFromType, Button As Integer, Shift As Integer, X As Long, Y As Long)

Public Event OnGroupHint(ByVal strHintContext As String)
Public Event OnFilter(rsData As ADODB.Recordset, ByRef blnCancel As Boolean, ByRef blnUseCustom As Boolean)
Public Event OnConfigEvent(ByRef blnUseCustom As Boolean)

Public Event OnModifyBefore(ByVal lngListType As TQueueFromType, ByVal lngQueueId As Long, ByRef objInput As Dictionary, ByRef blnCancel As Boolean, ByRef blnUseCustom As Boolean)
Public Event OnModifyAfter(ByVal lngQueueId As Long, objUpdateValue As Dictionary)

Public Event OnMsgRecevie(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset, ByRef blnUseCustom As Boolean)

Private Declare Function GetTickCount Lib "kernel32" () As Long


'***************************************************************
'只读属性定义


'当前使用的队列类型
Property Get CurQueueType() As TQueueFromType
    If mblnIsFindQueue Then
        CurQueueType = qftFindQueue             '查找队列
    Else
        If mblnIsSelectedCallingList = True Then
            CurQueueType = qftCalledQueue       '呼叫队列
        Else
            CurQueueType = qftWaitQueue         '排队队列
        End If
    End If
End Property


'获取业务类型
Property Get WorkType() As Long
    WorkType = mintWorkType
End Property


'排队处理对象
Property Get QueueOper() As clsQueueOperation
    Set QueueOper = mobjQueueManage
End Property

'CommandBar对象属性
Property Get CmdBar() As Object
    Set CmdBar = cbrMain
End Property


'等候呼叫列表对象
Property Get WaitQueueList() As Object
    Set WaitQueueList = rptQueueList
End Property

'已呼叫列表对象
Property Get CallQueueList() As Object
    Set CallQueueList = rptCallList
End Property



'***************************************************************
'读写属性定义


'是否关联排队标记
Property Get IsReleationQueueTag() As Boolean
    IsReleationQueueTag = mblnIsReleationQueueTag
End Property

Property Let IsReleationQueueTag(value As Boolean)
    mblnIsReleationQueueTag = value
End Property

'自定义排序字段
Property Get CustomOrderField() As String
    CustomOrderField = UCase(mstrCustomOrderColName)
End Property

Property Let CustomOrderField(value As String)
    If UCase(value) <> mstrCustomOrderColName Then
        mstrCustomOrderColName = UCase(value)
    End If

End Property

'报表编号
Property Get ReportNum() As String
    ReportNum = mobjQueueManage.ReportNum
End Property

Property Let ReportNum(value As String)
    mobjQueueManage.ReportNum = value
End Property


'最后分组队列
Property Get LastFixedQueue() As String
    LastFixedQueue = mstrLastFixedQueue
End Property


Property Let LastFixedQueue(value As String)
    mstrLastFixedQueue = value
End Property


'查找窗口中所包含的查找方式
Property Get FindWayEx() As String
    FindWayEx = mstrFindWay
End Property

Property Let FindWayEx(value As String)
    mstrFindWay = value
End Property


'是否显示排队叫号工具栏按钮
Property Get IsShowBars() As Boolean
    IsShowBars = mblnIsShowBars 'cbrMain.ActiveMenuBar.Visible
End Property

Property Let IsShowBars(value As Boolean)
    mblnIsShowBars = value
    cbrMain.ActiveMenuBar.Visible = value
End Property


'是否显示已呼叫排队队列
Property Get IsShowCalledQueue() As Boolean
    IsShowCalledQueue = mblnIsShowCalledQueue
End Property

Property Let IsShowCalledQueue(value As Boolean)
    mblnIsShowCalledQueue = value
    
    If DkpMain.PanesCount <= 0 Then Exit Property
    
    If value Then
        DkpMain.Panes(2).Closed = False
    Else
        DkpMain.Panes(2).Closed = True
    End If
End Property

'参与查询的数据字段
Property Get DataFields() As String
    DataFields = UCase(mstrDataFields)
End Property

Property Let DataFields(value As String)
    mstrDataFields = UCase(value)
End Property


'设置呼叫时的目的地
Property Get CalledTarget() As String
    CalledTarget = mobjQueueManage.CallTarget
End Property

Property Let CalledTarget(value As String)
    mobjQueueManage.CallTarget = value
End Property

'排队列表显示字段属性
Property Get DisplayQueueFields() As String
    DisplayQueueFields = UCase(mstrDisplayQueueFields)
End Property

Property Let DisplayQueueFields(value As String)
    If UCase(value) <> mstrDisplayQueueFields Then
        mstrDisplayQueueFields = UCase(value)
    End If
End Property

'呼叫列表显示字段属性
Property Get DisplayCallFields() As String
    DisplayCallFields = mstrDisplayCallFields
End Property

Property Let DisplayCallFields(value As String)
    If UCase(value) <> mstrDisplayCallFields Then
        mstrDisplayCallFields = UCase(value)
    End If
End Property


'轮训间隔时长(单位：毫秒)
Property Get Interval() As Long
    Interval = mlngInterval
End Property

Property Let Interval(value As Long)
    mlngInterval = value
End Property



'允许显示的排队队列名称  注：如果为空，则显示当前业务类型下的所有队列中的排队和呼叫数据
Property Get QueryQueueNames() As String
    QueryQueueNames = mstrQueryQueueNames
End Property

Property Let QueryQueueNames(value As String)
    mstrQueryQueueNames = value
End Property




'组名(对配对数据分组显示)
Property Get GroupField() As String
    GroupField = UCase(mstrGroupField)
End Property

Property Let GroupField(value As String)
    If UCase(value) <> mstrGroupField Then
        mstrGroupField = UCase(value)
    End If
End Property



'数据有效天数
Property Get ValidDays() As Long
    ValidDays = mobjQueueManage.ValidDays
End Property

Property Let ValidDays(value As Long)
    mobjQueueManage.ValidDays = value
End Property


'字体属性
Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Property Set Font(value As StdFont)
    Call SetFont(value)
End Property


'是否大图标
Property Get IsIconLarge() As Boolean
    IsIconLarge = cbrMain.Options.LargeIcons
End Property


Property Let IsIconLarge(value As Boolean)
    cbrMain.Options.LargeIcons = value
    
    Call cbrMain.RecalcLayout
End Property


'是否显示按钮文本
Property Get IsShowToolText() As Boolean
    IsShowToolText = IIf(mlngMenuCaptionStyle = xtpButtonIcon, False, True)
End Property

Property Let IsShowToolText(value As Boolean)
    Dim i As Integer
    Dim cbrControl As CommandBarControl

    If value = False Then
        '不显示文本
        mlngMenuCaptionStyle = xtpButtonIcon
        cbrMain(1).ShowTextBelowIcons = False
    Else
        '显示文本
        mlngMenuCaptionStyle = xtpButtonIconAndCaption
        cbrMain(1).ShowTextBelowIcons = True
    End If

    For Each cbrControl In cbrMain(1).Controls
        cbrControl.Style = mlngMenuCaptionStyle
    Next

    cbrMain.RecalcLayout
End Property

'插队原因
Property Get Reasons() As String
    Reasons = mstrReason
End Property

Property Let Reasons(value As String)
    mstrReason = value
End Property


Private Sub SetFont(ft As StdFont)
'设置字体显示
    Dim ftNew As StdFont
    
    If Font Is Nothing Then Exit Sub
    
    Set ftNew = New StdFont
    Call CopyFont(ft, ftNew)
    

    Set UserControl.Font = ftNew
    Set cbrMain.Options.Font = ftNew
    
    Set rptQueueList.PaintManager.CaptionFont = ftNew
    Set rptQueueList.PaintManager.TextFont = ftNew
    
    Set rptCallList.PaintManager.CaptionFont = ftNew
    Set rptCallList.PaintManager.TextFont = ftNew
    
    Set scQueueInf.Font = ftNew
    Set scCallInf.Font = ftNew

End Sub



Private Sub CopyFont(ftSource As StdFont, ByRef ftTarget As StdFont)
'复制字体属性
    
    ftTarget.Bold = ftSource.Bold
    ftTarget.Charset = ftSource.Charset
    ftTarget.Italic = ftSource.Italic
    ftTarget.Name = ftSource.Name
    ftTarget.Size = ftSource.Size
    ftTarget.Strikethrough = ftSource.Strikethrough
    ftTarget.Underline = ftSource.Underline
    ftTarget.Weight = ftSource.Weight
End Sub


Public Sub ApplyVoiceConfig()
'应用语音配置
    Dim str呼叫站点名称 As String
    
    
   '读取叫号方式
    If Val(GetSetting("ZLSOFT", gstrRegPath, "播放方式", 1)) Then
         str呼叫站点名称 = GetSetting("ZLSOFT", gstrRegPath, "远端呼叫站点", "")
         
         If Trim(str呼叫站点名称) = "" Then str呼叫站点名称 = AnalyseComputer
    Else
        str呼叫站点名称 = AnalyseComputer
    End If
    
    mstrComputerName = AnalyseComputer
    
    mobjQueueManage.PlayStation = str呼叫站点名称
    mobjQueueManage.LocalStation = mstrComputerName

    mobjQueueManage.PlayTimeLength = Val(GetSetting("ZLSOFT", gstrRegPath, "语音播放时长", 15))
    mobjQueueManage.PlayCount = Val(GetSetting("ZLSOFT", gstrRegPath, "语音播放次数", 2))
    mobjQueueManage.VoiceType = GetSetting("ZLSOFT", gstrRegPath, "语音类型", "")
    mobjQueueManage.IsPlayHintSound = Val(GetSetting("ZLSOFT", gstrRegPath, "语音呼叫前播放提示音", False))
    mobjQueueManage.PlaySpeed = Val(GetSetting("ZLSOFT", gstrRegPath, "语音播放语速", 0))
    mobjQueueManage.UseVbsPlay = IIf(Val(GetSetting("ZLSOFT", gstrRegPath, "启用VBS自定义呼叫", 1)) = 0, False, True)
    mobjQueueManage.CusVoiceScript = GetSetting("ZLSOFT", gstrRegPath, "VBS脚本", "")
    
    Interval = Val(GetSetting("ZLSOFT", gstrRegPath, "轮询间隔时间", 30))
    
    mblnAutoComplete = Val(GetSetting("ZLSOFT", gstrRegPath, "自动完成已接诊队列", 1))
    mblnShowMySelfCalled = Val(GetSetting("ZLSOFT", gstrRegPath, "只显示自己呼叫的队列", 1))
    
    If Val(GetSetting("ZLSOFT", gstrRegPath, "启用语音呼叫", 1)) = 0 Then
        Call StopVoice
    Else
        Call StartVoice
    End If
 
End Sub

Public Function ShowVoiceConfig() As Boolean
'显示语音配置
    ShowVoiceConfig = frmSetup.ShowMe(Me)
End Function


Public Sub UseMsgCenter(ByVal lngSys As Long, ByVal lngModule As Long, Optional ByVal strPrivs As String = "")
'启用消息中心
    Call mobjQueueManage.UseMsgCenter(lngSys, lngModule, strPrivs)
    
    Set mobjMsg = gobjMsgCenter
End Sub


'初始化队列
Public Sub InitQueue( _
    cnOracle As ADODB.Connection, _
    ByVal intWorkType As Integer, _
    ByVal objOwnerForm As Object, _
    ByVal strProTag As String, _
    Optional ByVal strLoginUser As String = "system", _
    Optional ByVal strPrivs As String = "[[#ALLPOPEDOM#]]")
      
    mblnInitOk = False
    
    '调用初始化队列方法
    Call mobjQueueManage.InitQueue(cnOracle, intWorkType, strLoginUser)
    
    '设置全局变量
    Set mcnOracle = cnOracle
    Set mobjOwner = objOwnerForm
    
    mstrProTag = strProTag
    mstrPrivs = strPrivs
    mintWorkType = intWorkType
    
    gstrRegPath = "公共模块\" & mstrProTag & "\排队叫号"

    mblnIsSelectedCallingList = False
    
    If Trim(mstrDataFields) = "" Then
        mstrDataFields = mobjQueueManage.DefQueryCols
    End If
     
    '当前登陆的用户名
    mstrLoginUserName = strLoginUser
    
    If DkpMain.PanesCount > 0 Then
        '重新配置界面
        DkpMain.CloseAll
        DkpMain.DestroyAll

        Call InitFaceScheme
    End If
    
    Call InitLocalParas                                                 '初始化参数
    Call SetCommandBarStyle
    Call InitCommandBars                                                '初始化各个工具栏按钮
    
    Call InitQueueList(rptQueueList, mstrGroupField, mstrCustomOrderColName, mstrDisplayQueueFields, UCase(mstrDataFields))          '初始化等待呼叫队列列表
    Call InitQueueList(rptCallList, mstrGroupField, mstrCustomOrderColName, mstrDisplayCallFields, UCase(mstrDataFields))             '初始化已呼叫列表

    mblnInitOk = True
    
End Sub


Private Sub InitFaceScheme()
'初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane
    Dim dblQueueRate As Double
    
    On Error GoTo errHandle
    
    With DkpMain
        .SetCommandBars cbrMain
        
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    dblQueueRate = 0.666
    
    If gstrRegPath <> "" Then
        dblQueueRate = Val(GetSetting("ZLSOFT", gstrRegPath, "QueueListWidthRate", "0.6"))
    End If
    
    Set Pane1 = DkpMain.CreatePane(0, dblQueueRate * 100, 2000, DockLeftOf, Nothing)
                
    Pane1.Title = "排队列表"
    Pane1.Tag = 0
    Pane1.Handle = picQueueFace.hwnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
       
    
    Set Pane2 = DkpMain.CreatePane(1, (1 - dblQueueRate) * 100, 2000, DockRightOf, Nothing)
    
    Pane2.Title = "呼叫列表"
    Pane2.Tag = 1
    Pane2.Handle = picCallFace.hwnd
    Pane2.Options = PaneNoFloatable Or PaneNoCaption

    If mblnIsShowCalledQueue Then
        DkpMain.Panes(2).Closed = False
    Else
        DkpMain.Panes(2).Closed = True
    End If

    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub lblQueueFilter_Click(Index As Integer)
On Error GoTo errHandle
    If optOutQueue(Index).value = 0 Then
        optOutQueue(Index).value = 1
    Else
        optOutQueue(Index).value = 0
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mobjMsg_OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset)
'消息接收处理(这里只能接收到排队叫号相关的消息)
    Dim blnUseCustom As Boolean
    Dim strValue As String
    
    '不处理语音播放消息
    If strMsgItemIdentity = G_STR_MSG_QUEUE_004 Then Exit Sub
    
    '判断消息中的队列名称是否需要进行处理的队列名称
    rsData.Filter = "node_name='queue_name'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "消息无效，检测到未包含有效的队列名称，终止消息处理。"
        Exit Sub
    End If
    
    strValue = Nvl(rsData!node_value)
    
    If InStr(mstrQueryQueueNames, strValue) <= 0 Then
        Debug.Print "该消息所属队列不属于当前业务处理范围，忽略消息处理。"
        Exit Sub
    End If
    
    blnUseCustom = False
    RaiseEvent OnMsgRecevie(strMsgItemIdentity, strXmlContext, rsData, blnUseCustom)
    
    '如果调用者进行了消息处理，则这里将直接退出
    If blnUseCustom Then Exit Sub
    
    
    
    Select Case strMsgItemIdentity
        Case G_STR_MSG_QUEUE_001    '入队消息
            Call LineQueueMsgProcess(rsData)
            
        Case G_STR_MSG_QUEUE_002    '完成消息
            '从队列中删除数据显示
            Call CompleteMsgProcess(rsData)
            
        Case G_STR_MSG_QUEUE_003    '状态同步消息
            '根据状态更新列表中的数据
            Call StateSyncMsgProcess(rsData)
    End Select
    
End Sub

Private Sub LineQueueMsgProcess(ByVal rsData As ADODB.Recordset)
'排队消息处理
    Dim lngQueueId As Long
    Dim lngQueueRow As Long
    Dim lngRecordIndex As Long
    Dim objQueueList As ReportControl
    
    '判断已呼叫队列中是否存在该数据，如果存在，则需要进行删除
    rsData.Filter = "node_name='queue_id'"
    If rsData.RecordCount > 0 Then lngQueueId = Nvl(rsData!node_value)
    
    Call LocateQueueRow(lngQueueId, objQueueList, lngQueueRow)
    
    If lngQueueRow > 0 Then
        lngRecordIndex = objQueueList.Rows(lngQueueRow).Record.Index
        objQueueList.Rows(lngQueueRow).Selected = False
        
        Call objQueueList.Records.RemoveAt(lngRecordIndex)
        Call objQueueList.Populate
        
        If objQueueList.Rows.Count > lngQueueRow Then
            objQueueList.Rows(lngQueueRow).Selected = True
        End If
    End If
    
    '刷新排队队列数据
    Call LoadWaitQueueData
End Sub

Private Sub CompleteMsgProcess(ByVal rsData As ADODB.Recordset)
'完成消息的处理过程
    Dim lngQueueId As Long
    Dim lngQueueRow As Long
    Dim lngRecordIndex As Long
    Dim objQueueList As ReportControl
    
    '获取消息中的队列ID
    rsData.Filter = "node_name='queue_id'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "消息无效，检测到未包含有效的队列ID，终止消息处理。"
        Exit Sub
    End If
    
    lngQueueId = Val(Nvl(rsData!node_value))
    
    Call LocateQueueRow(lngQueueId, objQueueList, lngQueueRow)
    
    If lngQueueRow > 0 Then
        lngRecordIndex = objQueueList.Rows(lngQueueRow).Record.Index
        objQueueList.Rows(lngQueueRow).Selected = False
        
        Call objQueueList.Records.RemoveAt(lngRecordIndex)
        Call objQueueList.Populate
        
        If objQueueList.Rows.Count > lngQueueRow Then
            objQueueList.Rows(lngQueueRow).Selected = True
        End If
    End If
    
End Sub


Private Sub StateSyncMsgProcess(ByVal rsData As ADODB.Recordset)
'排队消息处理
    Dim lngQueueId As Long
    Dim lngCurState As Long
    
    '获取消息中的队列ID
    rsData.Filter = "node_name='queue_id'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "消息无效，检测到未包含有效的队列ID，终止消息处理。"
        Exit Sub
    End If
    
    lngQueueId = Val(Nvl(rsData!node_value))
    
    '获取消息中的队列状态
    rsData.Filter = "node_name='queue_state'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "消息无效，检测到未包含有效的队列ID，终止消息处理。"
        Exit Sub
    End If
    
    lngCurState = Val(Nvl(rsData!node_value))
    
    '刷新界面中的状态数据
    Call RefreshQueueRowState(lngQueueId, lngCurState)
End Sub

Private Sub optOutQueue_Click(Index As Integer)
On Error GoTo errHandle
'    Dim i As Long
'
'    If Not mblnInitOk Then Exit Sub
'
'    If optOutQueue(Index).value <> 0 Then
'        optOutQueue(Index).Enabled = False
'        lblQueueFilter(Index).FontBold = True
'    End If
'
'    '设置未被选择的显示样式
'    For i = 0 To optOutQueue.Count - 1
'        If i <> Index Then
'            If optOutQueue(Index).value <> 0 Then
'                optOutQueue(i).value = 0
'                optOutQueue(i).Enabled = True
'
'                lblQueueFilter(i).FontBold = False
'            End If
'        End If
'    Next i
'
'    If optOutQueue(Index).value = 0 Then Exit Sub

    Call LoadWaitQueueData
    
    '设置当前排队队列为焦点队列
    mblnIsSelectedCallingList = False
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadWaitQueueData()
'载入等待队列数据
    Dim rsData As ADODB.Recordset
    
    Set rsData = QueryQueueData()
    
    If rsData Is Nothing Then
        rptQueueList.Records.DeleteAll
        rptQueueList.Populate
        
        Exit Sub
    End If
    
    Call LoadDataToList(rptQueueList, rsData)
    
    rptQueueList.Populate
End Sub

Private Sub LoadCallQueueData()
'载入呼叫队列数据
    Dim rsData As ADODB.Recordset
    
    Set rsData = QueryQueueData()
    
    If rsData Is Nothing Then
        rptCallList.Records.DeleteAll
        rptCallList.Populate
        
        Exit Sub
    End If
    
    Call LoadDataToList(rptCallList, rsData)

End Sub

Private Sub picCallFace_Resize()
On Error GoTo errHandle
   
    scCallInf.Left = 0
    scCallInf.Top = 0
    scCallInf.Width = picCallFace.Width

    rptCallList.Left = 0
    rptCallList.Top = scCallInf.Height
    rptCallList.Width = picCallFace.ScaleWidth
    rptCallList.Height = picCallFace.ScaleHeight - scCallInf.Height
    
errHandle:
End Sub

Private Sub picQueueFace_Resize()
On Error GoTo errHandle
    
    scQueueInf.Left = 0
    scQueueInf.Top = 0
    scQueueInf.Width = picQueueFace.Width

    rptQueueList.Left = 0
    rptQueueList.Top = scQueueInf.Height
    rptQueueList.Width = picQueueFace.ScaleWidth
    rptQueueList.Height = picQueueFace.ScaleHeight - scQueueInf.Height
    
    
    optOutQueue(TQueueSelState.qss排队中).Left = scQueueInf.Width - 4000
    optOutQueue(TQueueSelState.qss排队中).Top = 65

    optOutQueue(TQueueSelState.qss已暂停).Left = optOutQueue(0).Left + 1000
    optOutQueue(TQueueSelState.qss已暂停).Top = 65


    optOutQueue(TQueueSelState.qss已弃号).Left = optOutQueue(1).Left + 1000
    optOutQueue(TQueueSelState.qss已弃号).Top = 65


    optOutQueue(TQueueSelState.qss已完成).Left = optOutQueue(2).Left + 1000
    optOutQueue(TQueueSelState.qss已完成).Top = 65
    
    
    lblQueueFilter(TQueueSelState.qss排队中).Left = optOutQueue(0).Left + optOutQueue(0).Width + 20
    lblQueueFilter(TQueueSelState.qss排队中).Top = 55
    
    lblQueueFilter(TQueueSelState.qss已暂停).Left = optOutQueue(1).Left + optOutQueue(1).Width + 20
    lblQueueFilter(TQueueSelState.qss已暂停).Top = 55
    
    lblQueueFilter(TQueueSelState.qss已弃号).Left = optOutQueue(2).Left + optOutQueue(2).Width + 20
    lblQueueFilter(TQueueSelState.qss已弃号).Top = 55
    
    lblQueueFilter(TQueueSelState.qss已完成).Left = optOutQueue(3).Left + optOutQueue(3).Width + 20
    lblQueueFilter(TQueueSelState.qss已完成).Top = 55
    
errHandle:
End Sub


Private Sub SetCommandBarStyle()
On Error GoTo errHandle
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrMain.VisualTheme = xtpThemeOffice2003
    Set cbrMain.Icons = zlCommFun.GetPubIcons
    
    
    With cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        '.LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    
    cbrMain.EnableCustomization False
    cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
Exit Sub
errHandle:
End Sub


Private Function CheckPopedom(ByVal strFuncName As String) As Boolean
'检查权限是否存在
    CheckPopedom = False
    
    If mstrPrivs = M_STR_ALL_POPEDOM Then
        CheckPopedom = True
        Exit Function
    End If
    
    If InStr("," & mstrPrivs & ",", "," & strFuncName & ",") > 0 Then CheckPopedom = True
    If InStr(";" & mstrPrivs & ";", ";" & strFuncName & ";") > 0 Then CheckPopedom = True
    
End Function

Public Sub zlCreateMenuBars(cbrMenuBar As CommandBarPopup, Optional ByVal blnIsAllMenu As Boolean = False)
    Dim cbrControl As CommandBarControl
    
    If cbrMenuBar Is Nothing Then Exit Sub
    
    '创建常用呼叫菜单
    With cbrMenuBar.CommandBar.Controls
        If CheckPopedom("打号") Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_PrintNumber, "打号"): cbrControl.IconId = 3571
        End If
        
        If CheckPopedom("顺呼") Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "顺呼")
            cbrControl.BeginGroup = True
            cbrControl.IconId = 744
            cbrControl.ToolTipText = "按顺序呼叫下一个"
        End If
        
        If CheckPopedom("直呼") Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallThis, "直呼"): cbrControl.IconId = 732
        End If
        

        If CheckPopedom("广播") Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "广播"): cbrControl.IconId = 2600
        End If
        
        If Not blnIsAllMenu Then
            If CheckPopedom("过滤") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Filter, "过滤")
                cbrControl.BeginGroup = True
                cbrControl.IconId = 731
            End If
            
            If CheckPopedom("刷新") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Refresh, "刷新"): cbrControl.IconId = 3003
            End If
        End If
    End With
    
    If blnIsAllMenu Then
    
        With cbrMenuBar.CommandBar.Controls
            If CheckPopedom("插队") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_InsertQueue, "插队")
                cbrControl.IconId = 2600
                cbrControl.BeginGroup = True
            End If
            
            If CheckPopedom("重排") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RestartQueue, "重排"): cbrControl.IconId = 2614
            End If
            
    
            If CheckPopedom("暂停") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Pause, "暂停"): cbrControl.IconId = 746: cbrControl.BeginGroup = True
            End If
            
            If CheckPopedom("弃号") Then    '权限中使用弃号的名称
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Abandon, "弃呼"): cbrControl.IconId = 8113
            End If
            
            If CheckPopedom("恢复") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Restore, "恢复"): cbrControl.IconId = 252
            End If
            
            If CheckPopedom("接诊") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RecDiagnose, "接诊"): cbrControl.IconId = 8264
            End If
            
            If CheckPopedom("完成") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Finaled, "完成"): cbrControl.IconId = 747
            End If
            
            If CheckPopedom("过滤") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Filter, "过滤"): cbrControl.IconId = 731
            End If
            
            If CheckPopedom("刷新") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Refresh, "刷新"): cbrControl.IconId = 3003
            End If
        End With
    End If
    
    For Each cbrControl In cbrMenuBar.Controls
        If (cbrControl.Type = xtpControlButton) Or (cbrControl.Type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
        If cbrControl.Category = "" Then cbrControl.Category = "QUEUE" '设置排队菜单标识
    Next
End Sub

'初始化功能按钮
Private Sub InitCommandBars()
    Dim cbrToolBar1 As CommandBar
    Dim cbrToolBar2 As CommandBar
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrCustom  As CommandBarControlCustom
    Dim i As Integer
    
    cbrMain(1).Visible = False
    
    For i = cbrMain.Count To 2 Step -1
        cbrMain(i).Controls.DeleteAll
    Next i
    
    Set cbrMain.Icons = zlCommFun.GetPubIcons
    
    
    '排队呼叫工具栏定义
    If cbrMain.Count > 1 Then
        Set cbrToolBar1 = cbrMain(2)
    Else
        Set cbrToolBar1 = cbrMain.Add("工具栏", XTPBarPosition.xtpBarTop)
    End If
    
    cbrToolBar1.Closeable = False
    '将CommandBar工具栏 设置成 上图标下文本的形式
    cbrToolBar1.ShowTextBelowIcons = True

    With cbrToolBar1.Controls

        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_PrintNumber, "打号"): cbrControl.IconId = 103
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "顺呼"): cbrControl.IconId = 744: cbrControl.ToolTipText = "按顺序呼叫下一个": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallThis, "直呼"): cbrControl.IconId = 732
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "重呼"): cbrControl.IconId = 745
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_InsertQueue, "插队"): cbrControl.IconId = 2600: cbrControl.BeginGroup = True: cbrControl.ToolTipText = "设置为优先呼叫"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RestartQueue, "重排"): cbrControl.IconId = 2614: cbrControl.ToolTipText = "重新进入队列排队"

        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Pause, "暂停"): cbrControl.IconId = 746: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Abandon, "弃呼"): cbrControl.IconId = 8113: cbrControl.ToolTipText = "放弃呼叫"
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Restore, "恢复"): cbrControl.IconId = 252: cbrControl.ToolTipText = "将数据恢复到排队状态"
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RecDiagnose, "接诊"): cbrControl.IconId = 3009: cbrControl.ToolTipText = "对报到病人进行接诊处理"
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Finaled, "完成"): cbrControl.IconId = 747
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Filter, "过滤"): cbrControl.IconId = 731: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Refresh, "刷新"): cbrControl.IconId = 791
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Update, "修改"): cbrControl.IconId = 3003: cbrControl.ToolTipText = "修改排队信息"
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Setup, "设置"): cbrControl.IconId = 181: cbrControl.ToolTipText = "参数设置": cbrControl.BeginGroup = True
        
    End With
    
    Call DoCmdBarInitEvent(cbrToolBar1)
    
    For Each cbrControl In cbrToolBar1.Controls
        If (cbrControl.Type = xtpControlButton) Or (cbrControl.Type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
        If cbrControl.Category = "" Then cbrControl.Category = "QUEUE" '设置排队菜单标识
    Next
        
    cbrToolBar1.Visible = mblnIsShowBars
    txtLocateValue.Visible = False
    
    If CheckPopedom("查找") And CheckPopedom("定位") Then
        
        If cbrMain.Count > 2 Then
            Set cbrToolBar2 = cbrMain(3)
        Else
            Set cbrToolBar2 = cbrMain.Add("查找栏", XTPBarPosition.xtpBarTop)
        End If
        
        cbrToolBar2.Closeable = False
        cbrToolBar2.ShowTextBelowIcons = False
        
        With cbrToolBar2.Controls
            '创建定位等操作
            Set cbrMenuBar = .Add(xtpControlPopup, conMenu_Queue_LocateType, "排队号")
                cbrMenuBar.Id = conMenu_Queue_LocateType
                cbrMenuBar.Flags = xtpFlagRightAlign
    
            Set cbrCustom = .Add(xtpControlCustom, conMenu_Queue_LocateValue, "定位条件")
                cbrCustom.Handle = txtLocateValue.hwnd
                cbrCustom.Flags = xtpFlagRightAlign
                cbrCustom.Style = xtpButtonIconAndCaption
    
                txtLocateValue.Visible = True
                
            Set cbrCustom = .Add(xtpControlCustom, conMenu_Queue_LocateValue, "")
                cbrCustom.Handle = picPlace.hwnd
    
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Locate, "")
                cbrControl.Id = conMenu_Queue_Locate
                cbrControl.IconId = 8267
                cbrControl.Flags = xtpFlagRightAlign
                cbrControl.ToolTipText = "定位排队数据"
                cbrControl.Checked = True
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Find, "")
                cbrControl.Id = conMenu_Queue_Find
                cbrControl.IconId = 721
                cbrControl.Flags = xtpFlagRightAlign
                cbrControl.ToolTipText = "查找排队数据"
        End With
        
        For Each cbrControl In cbrToolBar2.Controls
            If (cbrControl.Type = xtpControlButton) Or (cbrControl.Type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
            If cbrControl.Category = "" Then cbrControl.Category = "QUEUE" '设置排队菜单标识
        Next
        
        Call DockRightOfCommandBar(cbrToolBar2, cbrToolBar1)
        
        cbrToolBar2.Visible = mblnIsShowBars
        txtLocateValue.Visible = mblnIsShowBars
    End If
End Sub


Private Sub DockRightOfCommandBar(cbBarToDock As CommandBar, cbBarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbrMain.RecalcLayout
    
    cbBarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    '使cbBarToDock工具栏始终显示在最右边
    cbrMain.DockToolBar cbBarToDock, 300000, (Bottom + Top) / 2, cbBarOnLeft.Position

End Sub

Private Sub DoCmdBarInitEvent(cbrToolBar As Object)
    '触发工具栏初始化事件
    RaiseEvent OnCmdBarInit(cbrToolBar)
End Sub


Private Function GetWaitQueueSelState() As TQueueSelState
'获取排队队列的数据显示状态，排队，暂停，弃呼，完成
    Dim lngState As TQueueSelState
    Dim strState As String
    
     '得到当前选中的排队状态值
    lngState = -1
    
'    If optOutQueue(TQueueSelState.qss排队中).value Then lngState = TQueueSelState.qss排队中     '排队中状态
'    If optOutQueue(TQueueSelState.qss已暂停).value Then lngState = TQueueSelState.qss已暂停     '已暂停状态
'    If optOutQueue(TQueueSelState.qss已弃号).value Then lngState = TQueueSelState.qss已弃号     '已弃号状态
'    If optOutQueue(TQueueSelState.qss已完成).value Then lngState = TQueueSelState.qss已完成     '已完成状态
    
    If rptQueueList.SelectedRows.Count > 0 Then
        If rptQueueList.SelectedRows(0).GroupRow = False Then
            strState = rptQueueList.SelectedRows(0).Record(GetColIndex("排队状态", rptQueueList)).value
            If strState = "排队中" Then
                lngState = TQueueSelState.qss排队中
            ElseIf strState = "已暂停" Then
                lngState = TQueueSelState.qss已暂停
            ElseIf strState = "已弃呼" Then
                lngState = TQueueSelState.qss已弃号
            Else
                lngState = TQueueSelState.qss已完成
            End If
            
            If lngState = -1 And strState <> "已完成" Then lngState = TQueueSelState.qss排队中    '当未选择任何状态时，则只加载排队中的数据
        End If
    End If
    
    GetWaitQueueSelState = lngState
End Function


Private Function QueryQueueData() As ADODB.Recordset
'查询队列数据到数据集
    Dim strSql As String
    Dim strTemp As String
    Dim blnUseCustom As Boolean
    Dim lngTimePoint As Long
    Dim strStartTime As String
    Dim strEndTime As String
    Dim blnHasQueueCol As Boolean
    Dim strSelectColumns As String
    Dim strOrderCondition As String
    Dim strCurQueryQueueNames As String
    Dim rsData As ADODB.Recordset
    Dim i As Long
    Dim dtNow As Date
    
    Set QueryQueueData = Nothing
    
    blnUseCustom = False
    RaiseEvent OnQueryQueueData(rsData, blnUseCustom)
    
    '当未使用自定义查询时（blnUseCustom为false），则使用系统默认的查询数据源
    If Not blnUseCustom Then
        lngTimePoint = Val(Format(Time, "h"))
        dtNow = zlDatabase.Currentdate
        If lngTimePoint <= 4 Then
            strStartTime = To_Date(Format(dtNow - 1, "yy-mm-dd 20:00:00"))
            strEndTime = To_Date(Format(dtNow, "yy-mm-dd 08:00:00"))
        Else
            strStartTime = To_Date(Format(dtNow, "yy-mm-dd 00:00:00"))
            strEndTime = To_Date(Format(dtNow, "yy-mm-dd 23:59:59"))
        End If
    
       '调用得到所有字段列方法
        strSelectColumns = mobjQueueManage.DefQueryCols
        strOrderCondition = mobjQueueManage.CustomOrder
    
        '给队列名称添加单引号，以便在查询SQL语句
        strCurQueryQueueNames = Replace(mstrQueryQueueNames, ",", "','")
        strTemp = IIf(strCurQueryQueueNames = "", "", "and 队列名称 in ('" & strCurQueryQueueNames & "') ")
        strSql = "select " & strSelectColumns & " from 排队叫号队列 where 排队时间 between " & strStartTime & " and " & strEndTime & " and 业务类型=" & mintWorkType & " " & strTemp & _
                IIf(strOrderCondition <> "", " order by " & strOrderCondition, "")
       
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询队列数据")
    End If

    If rsData Is Nothing Then
        MsgBox "返回的数据集对象为空，不能执行数据加载操作。", vbInformation, "排队叫号系统"
        Exit Function
    End If
    
    '判断是否包含排队id列
    blnHasQueueCol = False
    'rsData返回时必须包含排队ID数据
    For i = 0 To rsData.Fields.Count - 1
        If UCase(rsData.Fields.Item(i).Name) = "ID" Then
            blnHasQueueCol = True
            Exit For
        End If
    Next i


    '数据集中不存在排队ID，不执行加载数据操作 直接退出过程
    If Not blnHasQueueCol Then
        MsgBox "查询数据集中不存在排队ID，不能执行数据加载操作。", vbInformation, "排队叫号系统"
        Exit Function
    End If
    
    Set QueryQueueData = rsData
    
End Function

Private Sub LoadDataToList(objCurQueueList As ReportControl, rsData As ADODB.Recordset, Optional ByVal blSetFocus As Boolean = True)
'载入队列数据
'lngQueueType:0排队队列，1呼叫队列
'blSetFocus 是否禁止设置列表焦点，默认 True
On Error GoTo errHandle
'加载数据到列表
    Dim rptRecord As ReportRecord
    Dim lngQueueLoadModle As TQueueSelState
    Dim i As Long
    Dim lngCurSelRow As Long
'    Dim lngQueryQueueState As Long
    Dim blnCancel As Boolean
    Dim lngOrdIndex As Long
    Dim strFilter As String
    Dim blnIsWaitQueue As Boolean
    Dim blnLoadData As Boolean
    Dim strQueueState As String
    
    objCurQueueList.Records.DeleteAll
    If rsData Is Nothing Then
        objCurQueueList.Populate
        Exit Sub
    End If
    
    blnIsWaitQueue = IIf(objCurQueueList.Name = rptQueueList.Name, True, False)
    
    If blnIsWaitQueue Then
        'lngQueueLoadModle = GetWaitQueueSelState()
'        lngQueryQueueState = Decode(lngQueueLoadModle, _
'                            TQueueSelState.qss排队中, TQueueState.qsQueueing, _
'                            TQueueSelState.qss已弃号, TQueueState.qsAbstain, _
'                            TQueueSelState.qss已暂停, TQueueState.qsPause, _
'                            TQueueState.qsComplete)
                  
        If optOutQueue(0).value = 1 Then
            strQueueState = "排队状态 = 0"
            
            If strFilter <> "" Then strFilter = strFilter & " Or "
            strFilter = strFilter & strQueueState
        End If
        
        If optOutQueue(1).value = 1 Then
            strQueueState = "排队状态 = 3"
            
            If strFilter <> "" Then strFilter = strFilter & " Or "
            strFilter = strFilter & strQueueState
        End If
        
        If optOutQueue(2).value = 1 Then
            strQueueState = "排队状态 = 2"
            
            If strFilter <> "" Then strFilter = strFilter & " Or "
            strFilter = strFilter & strQueueState
        End If
        
        If optOutQueue(3).value = 1 Then
            strQueueState = "排队状态 = 4"
            
            If strFilter <> "" Then strFilter = strFilter & " Or "
            strFilter = strFilter & strQueueState
        End If
        
        If strFilter = "" Then strFilter = "排队状态=0 or 排队状态=2 or 排队状态=3 or 排队状态=4"
    Else
        '过滤出呼叫中，已呼叫，接诊中，待呼叫的数据
        strFilter = "排队状态=1 or 排队状态=7 or 排队状态=8 or 排队状态=9"
    End If
    
    rsData.Filter = strFilter
    
    '如果没有数据，则直接退出
    If rsData.RecordCount <= 0 Then
        objCurQueueList.Populate
        Exit Sub
    End If
    
    '得到当前焦点列表的焦点行数
    If objCurQueueList.SelectedRows.Count > 0 Then lngCurSelRow = objCurQueueList.SelectedRows(0).Index


    lngOrdIndex = GetColIndex("ORD", objCurQueueList)
    While Not rsData.EOF
    
        blnCancel = False
        RaiseEvent OnReadBefore(rsData, IIf(blnIsWaitQueue, TQueueFromType.qftWaitQueue, TQueueFromType.qftCalledQueue), blnCancel)
        
        If Not blnCancel Then
            blnLoadData = False
            
            '加载排队状态为当前所选状态的数据
'            If (Nvl(rsData("排队状态"), -1) = lngQueryQueueState And blnIsWaitQueue) _
'                Or ((Nvl(rsData("排队状态"), -1) = TQueueState.qsCalling _
'                Or Nvl(rsData("排队状态"), -1) = TQueueState.qsCalled _
'                Or Nvl(rsData("排队状态"), -1) = TQueueState.qsDiagnose _
'                Or Nvl(rsData("排队状态"), -1) = TQueueState.qsWaitCall) And Not blnIsWaitQueue) Then
                
                If mblnShowMySelfCalled Then
                    blnLoadData = IIf(UCase(Nvl(rsData!呼叫医生)) = UCase(mstrLoginUserName) Or Nvl(rsData!呼叫医生) = "", True, False)
                Else
                    blnLoadData = True
                End If
'            End If
            

            If blnLoadData Then
                Set rptRecord = objCurQueueList.Records.Add
                
                For i = 0 To objCurQueueList.Columns.Count - 1
                    rptRecord.AddItem ""
                Next
    
                Call SetReportRecordItem(objCurQueueList, rptRecord, rsData)
                
                '用于当使用数据库默认的排序时，能够根据数据所在索引，在数据对齐后，由排队界面控件进行排序处理
                If lngOrdIndex >= 0 Then
                    rptRecord.Item(lngOrdIndex).value = Format(rsData.AbsolutePosition, "00000000")
                End If
   
                '设置背景颜色
                lngQueueLoadModle = Decode(Nvl(rsData("排队状态"), -1), _
                                    TQueueState.qsQueueing, TQueueSelState.qss排队中, _
                                    TQueueState.qsAbstain, TQueueSelState.qss已弃号, _
                                    TQueueState.qsPause, TQueueSelState.qss已暂停, _
                                    TQueueState.qsComplete, TQueueSelState.qss已完成)
                                    
                Select Case lngQueueLoadModle
                    Case TQueueSelState.qss已暂停
                        Call SetReportRecordColor(objCurQueueList, rptRecord, vbYellow)
                    Case TQueueSelState.qss已弃号
                        Call SetReportRecordColor(objCurQueueList, rptRecord, &H8080FF)
                    Case TQueueSelState.qss已完成
                        Call SetReportRecordColor(objCurQueueList, rptRecord, &HFF00&)
                    Case TQueueSelState.qss排队中
                        Call SetReportRecordColor(objCurQueueList, rptRecord, vbWhite)
                End Select

                RaiseEvent OnReadAfter(rsData, IIf(blnIsWaitQueue, TQueueFromType.qftWaitQueue, TQueueFromType.qftCalledQueue), rptRecord)
            End If
                        
        End If

        rsData.MoveNext
    Wend

    objCurQueueList.Populate
    
    '恢复选择的排队数据
    If lngCurSelRow >= objCurQueueList.Rows.Count Then
        lngCurSelRow = IIf(objCurQueueList.Rows.Count <= 0, -1, rptQueueList.Rows.Count - 1)
    End If

    '103315 增加这个处理，避免无意义的设置焦点行 导致不能取消报告
    If  blSetFocus Then
        If lngCurSelRow > -1 Then
            objCurQueueList.Rows(lngCurSelRow).Selected = True
            Set objCurQueueList.FocusedRow = objCurQueueList.Rows(lngCurSelRow)
        End If
    End If
    
    '恢复排序，如果这里不进行排序，在直呼或者顺乎以及恢复排队状态后，队列数据可能不会按顺序显示
    objCurQueueList.SortOrder(objCurQueueList.SortOrder.Count - 1).SortAscending = True

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function GetColumnWidth(ByVal strColPros As String, strColName As String) As Long
'获取列的宽度
On Error GoTo errHandle
    Dim strColPro As String
    Dim lngColIndex As Long
    
    GetColumnWidth = 100
    
    If gstrRegPath = "" Then Exit Function
    
    lngColIndex = InStr(strColPros, strColName & ":")
    If lngColIndex <= 0 Then Exit Function
    
    strColPro = Mid(strColPros, lngColIndex, 255)
    strColPro = Replace(strColPro, strColName & ":", "")
    
    GetColumnWidth = Val(strColPro)
    
Exit Function
errHandle:
    GetColumnWidth = 100
End Function


Public Function GetValidCols(ByVal strCols As String, Optional ByVal strQueueTabPrefix As String) As String
'必须具备的查询列字段：ID,队列名称,业务ID,患者姓名,排队状态,排队序号,排队号码
    Dim strResult As String
    Dim strTabPrefix As String
    
    strResult = UCase(strCols)
    strTabPrefix = UCase(strQueueTabPrefix)
    
    If Trim(strResult) = "" Then
        GetValidCols = mobjQueueManage.GetAllQueueTabCols(strQueueTabPrefix)
        Exit Function
    End If
    
    strResult = ",," & strResult & ",,"
    
    strResult = Replace(strResult, ", ", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "ID,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "队列名称,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "业务ID,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "患者姓名,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "排队状态,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "排队序号,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "排队号码,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "呼叫医生,", ",")
    strResult = Replace(strResult, ",,", "")
    
    strResult = "[[TAB]].ID,[[TAB]].队列名称,[[TAB]].业务ID,[[TAB]].患者姓名,[[TAB]].排队状态, [[TAB]].排队序号,[[TAB]].排队号码,[[TAB]].呼叫医生 " & IIf(strResult = "", "", "," & strResult)
    strResult = Replace(strResult, "[[TAB]].", IIf(strTabPrefix <> "", strTabPrefix & ".", ""))
    
    GetValidCols = strResult
End Function


Private Sub InitQueueList(objQueueList As Object, ByVal strGroupCol As String, ByVal strOrderCols As String, _
    ByVal strDisplayCols As String, ByVal strDataCols As String)
'初始化排队列表

    Dim Column As ReportColumn
    Dim strAllColNames As String
    Dim strQueueColNames() As String
    Dim strCallColNames() As String
    Dim strOrderCondition() As String
    Dim blnIsOrders As Boolean  '是否包含多个排序字段
    Dim i As Integer
    Dim j As Integer
    Dim aryDisplayCols() As String
    Dim strOrderCol As String
    Dim blnIsConfigOrder As Boolean
    Dim aryCurOrderInf() As String
    Dim objCurQueueList As ReportControl
    Dim lngColIndex As Long
    Dim strColPros As String
    
    Err = 0: On Error Resume Next
    
    If objQueueList Is Nothing Then Exit Sub
    Set objCurQueueList = objQueueList
    

    '初始化排队队列显示字段
    Call objCurQueueList.Records.DeleteAll
    Call objCurQueueList.Columns.DeleteAll
    
    Set objCurQueueList.Icons = zlCommFun.GetPubIcons

    '初始化列表相关属性
    objCurQueueList.AllowColumnRemove = False
    objCurQueueList.ShowItemsInGroups = False
    objCurQueueList.SkipGroupsFocus = True
    objCurQueueList.MultipleSelection = False

    With objCurQueueList.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "将列标题拖动到此,可按该列分组..."
        .NoItemsText = "没有可显示的项目..."
        .VerticalGridStyle = xtpGridSolid
    End With
    
    '获取排队叫号默认的查询数据列，默认为所有的数据字段，返回格式为以“，”逗号分隔的字段名称
    strAllColNames = strDataCols
    
    '如果没有查询到字段，则退出
    If Trim(strAllColNames) = "" Then Exit Sub
    If Trim(strDisplayCols) = "" Then strDisplayCols = strAllColNames
    
    '加载列表显示字段
    If Trim(strAllColNames) <> "" And strAllColNames <> "*" Then
    
        strQueueColNames() = Split(strAllColNames, ",")
        aryDisplayCols() = Split(strDisplayCols, ",")
        
        lngColIndex = 0
        
        strColPros = GetSetting("ZLSOFT", gstrRegPath, objQueueList.Name)
        
        For i = LBound(aryDisplayCols) To UBound(aryDisplayCols)
            If Trim(aryDisplayCols(i)) <> "" Then
                '载入需要显示的数据字段
                Set Column = objCurQueueList.Columns.Add(lngColIndex, aryDisplayCols(i), GetColumnWidth(strColPros, aryDisplayCols(i)), True)
                lngColIndex = lngColIndex + 1
                
                '判断该列是否参与分组
                If InStr("," & strGroupCol & ",", "," & aryDisplayCols(i) & ",") > 0 Then
                    Column.Groupable = True
                    
                    '参与分组的列不进行显示
                    Column.Visible = False
                End If
                
                RaiseEvent OnColumnInit(objCurQueueList, Column)
                
            End If
        Next i

        For i = LBound(strQueueColNames) To UBound(strQueueColNames)
            If Trim(strQueueColNames(i)) <> "" And InStr(strDisplayCols, strQueueColNames(i)) <= 0 Then
                '载入不需要显示的字段
                Set Column = objCurQueueList.Columns.Add(lngColIndex, strQueueColNames(i), 100, True)
                lngColIndex = lngColIndex + 1
                
                If InStr("," & strGroupCol & ",", "," & strQueueColNames(i) & ",") > 0 Then
                    Column.Groupable = True
                End If
                
                Column.Visible = False
            End If
        Next i
        
        '用于未传递自定义排序子段后的分组排序
        Set Column = objCurQueueList.Columns.Add(lngColIndex, "ORD", 0, False)
        Column.Visible = False
    End If
    
    '如果没有使用自定义排序，则控件不进行排序处理，加载数据时使用数据库的默认排序
    blnIsConfigOrder = False
    If Trim(strOrderCols) <> "" Then
        aryCurOrderInf = Split(strOrderCols, ",")
        blnIsConfigOrder = True
    End If
    
    

    '处理分组以及排序的规则
    With objCurQueueList
    
        If Trim(strGroupCol) <> "" Then
            .GroupsOrder.DeleteAll

            '只允许按照其中一个字段进行分组
            For i = 0 To .Columns.Count
                If .Columns(i).Caption = strGroupCol Then
                    .GroupsOrder.Add .Columns(i)
                    Exit For
                End If
            Next i

            '分组之后,如果分组列不显示,分组列的排序是不变的
            .GroupsOrder(0).SortAscending = True ' False '
        End If
        
         '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.DeleteAll
        .AllowColumnSort = False 'blnIsConfigOrder
        
        
        If blnIsConfigOrder Then
            '配置排序字段
            For i = LBound(aryCurOrderInf) To UBound(aryCurOrderInf)
                strOrderCol = Trim(aryCurOrderInf(i))
                
                If strOrderCol <> "" Then
                    If InStr(strOrderCol, "DESC") > 0 Then
                        '按降序排序
                        strOrderCol = Replace(strOrderCol, "DESC", "")
    
                        For j = 0 To .Columns.Count - 1
                            If .Columns(j).Caption = strOrderCol Then
                                .SortOrder.Add .Columns(j)
                                .SortOrder(.SortOrder.Count - 1).SortAscending = False
                                Exit For
                            End If
                        Next j
                    Else
                        '按升序排序
                        strOrderCol = Replace(strOrderCol, "ASC", "")
                        
                        For j = 0 To .Columns.Count - 1
                            If .Columns(j).Caption = strOrderCol Then
                                .SortOrder.Add .Columns(j)
                                .SortOrder(.SortOrder.Count - 1).SortAscending = True
                                Exit For
                            End If
                        Next j
                    End If
                End If
            Next i
        Else
            '如果没有设置排序列，则根据数据库的加载顺序进行排序
            .SortOrder.Add .Columns(GetColIndex("ORD", objCurQueueList))
            .SortOrder(.SortOrder.Count - 1).SortAscending = True
        End If
        
    End With
    
End Sub

Private Function FormatQueueOrder(ByVal strQueueOrder As String) As String
'格式化排队序号
    Dim strQueueNum As String
    Dim strQueueNumFront As String
    Dim strQueueNumBehind As String
    
    FormatQueueOrder = ""
    strQueueNum = strQueueOrder
    
    '如果包含小数，则需要截取处理
    If InStr(strQueueNum, ".") > 0 Then
        strQueueNumFront = Mid(strQueueNum, 1, InStr(strQueueNum, ".") - 1)
        strQueueNumBehind = Mid(strQueueNum, InStr(strQueueNum, "."), Len(strQueueNum))
        
        strQueueNumFront = Replace(Space(M_LNG_FORMAT_ORDER_LEN - Len(strQueueNumFront)), " ", "0") & strQueueNumFront
    Else
        strQueueNum = Replace(Space(M_LNG_FORMAT_ORDER_LEN - Len(strQueueNum)), " ", "0") & strQueueNum
    End If
    
    '加载处理后的数据
    FormatQueueOrder = IIf(InStr(strQueueNum, ".") > 0, strQueueNumFront & strQueueNumBehind, strQueueNum)
End Function


Private Sub SetReportRecordItem(rptControl As ReportControl, rptItems As ReportRecord, rsData As ADODB.Recordset)
'设置ReportControl控件的数据
    On Error GoTo errHandle
    Dim i As Long
    Dim j As Long
    Dim intMaxNumLen As Integer
    Dim strQueueNum As String
    Dim strQueueNumFront As String
    Dim strQueueNumBehind As String
    Dim strValue As String
    Dim lngFitColWidth As Long
    Dim lngQueueState As Long
    Dim lngColIndex As Long
    Dim lngFirstCol As Long
    
    lngQueueState = Val(Nvl(rsData!排队状态))
    lngFirstCol = GetFirstDisplayColIndex(rptControl)
    
    '循环加载各个单元格数据
    For i = 0 To rptControl.Columns.Count - 1
        lngColIndex = rptControl.Columns(i).ItemIndex
        
        '因ReportControl控件的分组后排序的特殊性，需要对“排队序号”做相应字符串的处理，才能正确的进行排序
        If Not HasField(rsData, rptControl.Columns(i).Caption) Then
            rptItems(lngColIndex).value = ""
        Else
            strValue = Nvl(rsData("" & rptControl.Columns(i).Caption & "").value)
            
            If (Trim(rptControl.Columns(i).Caption) = "排队号码" Or Trim(rptControl.Columns(i).Caption) = "排队号") And mblnIsReleationQueueTag Then
                strValue = Nvl(rsData!排队标记) & strValue
            End If
            
            '判断是否是“排队序号”列，是则进入处理，否则直接加载
            If rptControl.Columns(i).Caption = "排队序号" Then
            
                strQueueNum = FormatQueueOrder(strValue)
                
                rptItems(lngColIndex).value = strQueueNum
            ElseIf rptControl.Columns(i).Caption = "排队状态" Then
            
                Select Case Val(rsData!排队状态)
                    Case TQueueState.qsQueueing
                        rptItems(lngColIndex).value = "排队中"
                    Case TQueueState.qsPause
                        rptItems(lngColIndex).value = "已暂停"
                    Case TQueueState.qsAbstain
                        rptItems(lngColIndex).value = "已弃呼"
                    Case TQueueState.qsComplete
                        rptItems(lngColIndex).value = "已完成"
                    Case TQueueState.qsCalled
                        rptItems(lngColIndex).value = "已呼叫"
                    Case TQueueState.qsCalling
                        rptItems(lngColIndex).value = "呼叫中"
                    Case TQueueState.qsDiagnose
                        rptItems(lngColIndex).value = "接诊中"
                    Case TQueueState.qsWaitCall
                        rptItems(lngColIndex).value = "待呼叫"
                End Select
                
            ElseIf rptControl.Columns(i).Caption = "队列名称" Then
                If InStr(Nvl(rsData!队列名称), IIf(Trim(mstrLastFixedQueue) <> "", mstrLastFixedQueue, "@<A...B  C.#.D>")) > 0 Then
                    If rptControl.GroupsOrder.Count > 0 Then
                        If rptControl.GroupsOrder(0).SortAscending = True Then
                            '按分组进行升序排序
                            rptItems(lngColIndex).value = " " & mstrLastFixedQueue
                        Else
                            '按分组进行降序排序
                            rptItems(lngColIndex).value = Chr(255) & mstrLastFixedQueue
                        End If
                    End If
                Else
                    rptItems(lngColIndex).value = strValue
                End If
                
            Else
                If IsDate(strValue) Then strValue = Format(strValue, "yyyy-mm-dd hh:mm:ss")
                rptItems(lngColIndex).value = strValue
                
            End If
    
            '设置患者姓名图标
            If lngColIndex = lngFirstCol Then
                If lngQueueState = TQueueState.qsDiagnose Then
                    rptItems(lngColIndex).Icon = M_LNG_ICON_DIAGNOSE
                ElseIf lngQueueState = TQueueState.qsCalling Then
                    rptItems(lngColIndex).Icon = M_LNG_ICON_CALLING
                ElseIf lngQueueState = TQueueState.qsCalled Then
                    rptItems(lngColIndex).Icon = M_LNG_ICON_CALLED
                Else
                    rptItems(lngColIndex).Icon = M_LNG_ICON_QUEUEING
                End If
            End If
            
            rptItems(i).BackColor = vbWhite
        End If
        
    Next i

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Public Function GetFirstDisplayColIndex(objQueueList As Object) As Long
'取得第一个显示的排队列
    Dim i As Long
    
    GetFirstDisplayColIndex = -1
    
    If Nvl(objQueueList.Tag, "") <> "" Then
        GetFirstDisplayColIndex = Val(Nvl(objQueueList.Tag))
        Exit Function
    End If
    
    For i = 0 To objQueueList.Columns.Count - 1
        If objQueueList.Columns(i).Visible Then
            GetFirstDisplayColIndex = objQueueList.Columns(i).ItemIndex
            objQueueList.Tag = objQueueList.Columns(i).ItemIndex
            Exit Function
        End If
    Next i
End Function


Private Sub SetReportRecordColor(objQueueList As Object, rrRow As ReportRecord, ByVal lngColor As Long)
'设置行的背景颜色
    Dim i As Long
    
    For i = 0 To objQueueList.Columns.Count - 1
        rrRow.Item(objQueueList.Columns(i).ItemIndex).BackColor = lngColor
    Next i
End Sub

Private Sub SetReportRecordBold(objQueueList As Object, rrRow As ReportRecord, ByVal blnBold As Boolean)
'设置行的字体加粗
    Dim i As Long
    
    For i = 0 To objQueueList.Columns.Count - 1
        rrRow.Item(objQueueList.Columns(i).ItemIndex).Bold = blnBold
    Next i
End Sub


Public Sub RefreshQueueData(Optional ByVal blSetFocus As Boolean = True )
'blSetFocus 是否禁止设置列表焦点，默认 True
'刷新排队队列数据
On Error GoTo errHandle
    Dim rsData As ADODB.Recordset
    
    Set rsData = QueryQueueData()
    If rsData Is Nothing Then Exit Sub

    Call LoadDataToList(rptQueueList, rsData, blSetFocus)
    Call LoadDataToList(rptCallList, rsData, blSetFocus)
    
    '恢复焦点列表
    If mblnIsSelectedCallingList Then
        Call SwitchActiveWindow(mblnIsSelectedCallingList)
    Else
        Call SwitchActiveWindow(mblnIsSelectedCallingList)
    End If
    
    Call ConfigQueueStateSel(True)
    mblnIsFindQueue = False

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub RefreshQueueRowData(ByVal lngQueueId As Long, ByVal strColName As String, ByVal strValue As String)
'刷新行数据
    Dim lngRowIndex As Long
    Dim lngColIndex As Long
    Dim objList As ReportControl
    
    Call LocateQueueRow(lngQueueId, objList, lngRowIndex)
    If lngRowIndex < 0 Then
        Exit Sub
    End If
    
    '查找列
    lngColIndex = GetColIndex(strColName, objList)
    If lngColIndex < 0 Then Exit Sub
    
    objList.Rows(lngRowIndex).Record(lngColIndex).value = strValue
End Sub


Public Sub RefreshQueueRowState(ByVal lngQueueId As Long, ByVal lngCurState As TQueueState)
'刷新队列行数据
    Dim lngRowIndex As Long
    Dim objList As ReportControl
    
    lngRowIndex = -1
    
    Call LocateQueueRow(lngQueueId, objList, lngRowIndex)
    If lngRowIndex < 0 Then
        '如果数据不存在，则刷新显示
        Call RefreshQueueData(False)
        Exit Sub
    End If
 
    If mblnIsFindQueue Then
        '如果是查找队列，则需要更新对应行数据的显示状态
        Call SetQueueRowState(objList, lngRowIndex, lngCurState)
        Call objList.Populate
    Else
        If objList.Name = rptQueueList.Name Then
            '如果列表选择的状态与数据当前状态不同，则删除
            If GetWaitQueueSelState() <> lngCurState Then
                Call DelQueueRecord(qftWaitQueue, lngRowIndex)
                
                '如果当前状态为待呼叫，则需要刷新呼叫队列进行数据显示
                If lngCurState = qsWaitCall Then
                    Call LoadCallQueueData
                End If
            Else
                Call SetQueueRowState(objList, lngRowIndex, lngCurState)
                Call objList.Populate
            End If
        Else
            '如果是已呼叫队列，则直接更新行状态
            If lngCurState <> qsCalled And lngCurState <> qsCalling And lngCurState <> qsDiagnose Then
                '如果已经不处于呼叫状态，则进行删除
                Call DelQueueRecord(qftCalledQueue, lngRowIndex)
                
                If GetWaitQueueSelState() = lngCurState Then
                    Call LoadWaitQueueData
                End If
            Else
                Call SetQueueRowState(objList, lngRowIndex, lngCurState)
                Call objList.Populate
            End If
        End If
    End If
End Sub

Private Sub LocateQueueRow(ByVal lngQueueId As Long, objList As ReportControl, ByRef lngRow As Long)
'根据队列ID定位行
    Dim i As Long
    Dim lngColIndex As Long
    
    lngRow = -1
    Set objList = Nothing
    
    '从排队队列开始查找
    lngColIndex = GetColIndex("ID", rptQueueList)
    
    For i = 0 To rptQueueList.Rows.Count - 1
        If rptQueueList.Rows(i).GroupRow = False Then
            If rptQueueList.Rows(i).Record(lngColIndex).value = lngQueueId Then
                lngRow = rptQueueList.Rows(i).Index
                Exit For
            End If
        End If
    Next i
    
    If lngRow <> -1 Then
        Set objList = rptQueueList
        Exit Sub
    End If
    
    '从呼叫队列开始查找
    lngColIndex = GetColIndex("ID", rptCallList)
    
    For i = 0 To rptCallList.Rows.Count - 1
        If rptCallList.Rows(i).GroupRow = False Then
            If rptCallList.Rows(i).Record(lngColIndex).value = lngQueueId Then
                lngRow = rptCallList.Rows(i).Index
                Exit For
            End If
        End If
    Next i
    
    If lngRow <> -1 Then Set objList = rptCallList
End Sub

Private Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function


Private Sub DoCmdBarExecute(Control As XtremeCommandBars.ICommandBarControl, ByRef blnUseCustom As Boolean)
    RaiseEvent OnCmdBarExecute(Control, blnUseCustom)
End Sub


Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnUseCustom As Boolean
    
    blnUseCustom = False
    
    Call DoCmdBarExecute(Control, blnUseCustom)
    
    '如果使用了自定义事件处理，则不执行后面的操作
    If blnUseCustom Then Exit Sub
    
    Select Case Control.Id
        Case conMenu_Queue_LocateType * 10# + 1 To conMenu_Queue_LocateType * 10# + 50   '定位
            Call Menu_View_Locate_Type_click(Control)

        Case conMenu_Queue_PrintNumber   '打号
            Call comMenu_打号

        Case conMenu_Queue_CallNext      '顺呼
            Call comMenu_顺呼

        Case conMenu_Queue_CallThis      '直呼
            Call comMenu_直呼

        Case conMenu_Queue_Broadcast     '广播
            Call comMenu_广播

        Case conMenu_Queue_InsertQueue   '插队
            Call comMenu_插队
            
        Case conMenu_Queue_RestartQueue     '重排
            Call comMenu_重排

        Case conMenu_Queue_RecDiagnose   '接诊
            Call comMenu_接诊

        Case conMenu_Queue_Pause         '暂停
            Call comMenu_暂停

        Case conMenu_Queue_Abandon       '弃号
            Call comMenu_弃号

        Case conMenu_Queue_Restore       '恢复
            Call comMenu_恢复

        Case conMenu_Queue_Finaled       '完成
            Call comMenu_完成

        Case conMenu_Queue_Filter        '刷新
            Call comMenu_过滤
            
        Case conMenu_Queue_Refresh       '刷新
            Call comMenu_刷新

        Case conMenu_Queue_Update        '修改
            Call comMenu_修改

        Case conMenu_Queue_Setup         '设置
            Call comMenu_设置
            
        Case conMenu_Queue_Locate       '定位
            Call SetLocateState(Control, True)
            
            Call LocateQueueData(mstrLocateType, txtLocateValue.Text)
        Case conMenu_Queue_Find          '查找
            Call SetLocateState(Control, False)

            Call FindQueueData(mstrLocateType, txtLocateValue.Text)

    End Select
End Sub


'按钮执行事件
Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Call zlExecuteCommandBars(Control)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub SetLocateState(Control As XtremeCommandBars.ICommandBarControl, ByVal blnIsLocate As Boolean)
    Dim objFindControl As XtremeCommandBars.ICommandBarControl
    
    Set objFindControl = cbrMain.FindControl(, IIf(blnIsLocate, conMenu_Queue_Find, conMenu_Queue_Locate), True, True)
    If Not objFindControl Is Nothing Then
        objFindControl.Checked = False
    End If
    
    Control.Checked = True
End Sub


Private Function IsFindModel() As Boolean
'判断是否为查找模式
    Dim cbrFind As CommandBarControl
    
    IsFindModel = False
    
    Set cbrFind = cbrMain.FindControl(, conMenu_Queue_Find, True, True)
    
    If cbrFind Is Nothing Then Exit Function

    IsFindModel = cbrFind.Checked
End Function

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim mintQueueState As Integer
    
    '获取排队队列的显示状态
    mintQueueState = GetWaitQueueSelState()

    If Not mblnInitOk Then Exit Sub
    
    Select Case Control.Id
        Case conMenu_Queue_LocateType   '定位或查找的类型
          Control.Visible = True
          Control.Enabled = True
          Control.Caption = mstrLocateType
            
        Case conMenu_Queue_LocateType * 10# + 1 To conMenu_Queue_LocateType * 10# + 50
          Control.Visible = True
          Control.Checked = (InStr(Control.Caption, mstrLocateType) > 0)
            
        Case conMenu_Queue_PrintNumber  '打号
          Control.Visible = CheckPopedom("打号")
          Control.Enabled = Trim(mobjQueueManage.ReportNum) <> ""
          
          
          
        Case conMenu_Queue_CallNext     '顺呼           '只有处于排队中的数据，才能进行顺乎操作
          Control.Visible = CheckPopedom("顺呼")
          Control.Enabled = Not mblnIsSelectedCallingList And mintQueueState = TQueueSelState.qss排队中 And Not mblnIsFindQueue
          
        Case conMenu_Queue_CallThis     '直呼           '排队列表中的数据，都可以进行直呼操作
          Control.Visible = CheckPopedom("直呼")
          Control.Enabled = Not mblnIsSelectedCallingList And mintQueueState = TQueueSelState.qss排队中 And Not mblnIsFindQueue
          
        Case conMenu_Queue_Broadcast    '广播           '只有呼叫后的数据才能进行广播
          Control.Visible = CheckPopedom("广播")
          Control.Enabled = mblnIsSelectedCallingList And rptCallList.SelectedRows.Count > 0
          
          
          
          
          
        Case conMenu_Queue_InsertQueue  '插队
          Control.Visible = CheckPopedom("插队")
          Control.Enabled = Not mblnIsSelectedCallingList And mintQueueState = TQueueSelState.qss排队中 And Not mblnIsFindQueue
          
        Case conMenu_Queue_RestartQueue '重排
          Control.Visible = CheckPopedom("重排")
          Control.Enabled = mintQueueState <> -1
          
        Case conMenu_Queue_RecDiagnose  '接诊
          Control.Visible = CheckPopedom("接诊")
          Control.Enabled = mblnIsSelectedCallingList Or mblnIsFindQueue Or (Not mblnIsSelectedCallingList And mintQueueState <> TQueueSelState.qss已完成 And mintQueueState <> -1 And mintQueueState <> TQueueSelState.qss排队中)
          
        Case conMenu_Queue_Pause        '暂停
          Control.Visible = CheckPopedom("暂停")
          Control.Enabled = (Not mblnIsSelectedCallingList And mintQueueState <> TQueueSelState.qss已暂停 And mintQueueState <> -1) Or mblnIsFindQueue Or mblnIsSelectedCallingList
          
        Case conMenu_Queue_Abandon      '弃号           '已呼叫数据可以进行弃号
          Control.Visible = CheckPopedom("弃号")
          Control.Enabled = (Not mblnIsSelectedCallingList And mintQueueState <> TQueueSelState.qss已弃号 And mintQueueState <> -1) Or mblnIsFindQueue Or mblnIsSelectedCallingList
          
        Case conMenu_Queue_Restore      '恢复
          Control.Visible = CheckPopedom("恢复")
          Control.Enabled = mblnIsSelectedCallingList Or (mintQueueState <> TQueueSelState.qss排队中 And mintQueueState <> -1 And Not mblnIsSelectedCallingList) Or mblnIsFindQueue
          
        Case conMenu_Queue_Finaled      '完成
          Control.Visible = CheckPopedom("完成")
          Control.Enabled = mblnIsSelectedCallingList Or mblnIsFindQueue Or (Not mblnIsSelectedCallingList And mintQueueState <> TQueueSelState.qss已完成 And mintQueueState <> -1 And mintQueueState <> TQueueSelState.qss排队中) '


          
        Case conMenu_Queue_Filter       '过滤
          Control.Visible = CheckPopedom("过滤")
          Control.Enabled = True
          
'        Case conMenu_Queue_Refresh      '刷新
'          Control.Visible = CheckPopedom("刷新")
'          Control.Enabled = True
          
        Case conMenu_Queue_Locate       '定位
          Control.Visible = CheckPopedom("定位")
          Control.Enabled = True
          
        Case conMenu_Queue_Find         '查找
          Control.Visible = CheckPopedom("查找")
          Control.Enabled = True
          
        Case conMenu_Queue_Update       '修改
          Control.Visible = CheckPopedom("修改")
          Control.Enabled = True
        
        Case conMenu_Queue_Setup        '设置
          Control.Visible = CheckPopedom("设置") Or CheckPopedom("参数设置")
          Control.Enabled = True
    End Select
    
    Call DoCmdBarUpdate(Control)
End Sub

'按钮更新事件
Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next
    Call zlUpdateCommandBars(Control)
    
    If mblnIsFindQueue Then
        scQueueInf.Caption = "查询结果："
    Else
        scQueueInf.Caption = "排队列表："
    End If
Err.Clear
End Sub

Private Sub DoCmdBarUpdate(Control As XtremeCommandBars.ICommandBarControl)
    RaiseEvent OnCmdBarUpdate(Control)
End Sub


Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    Dim aryKindInfo() As String
    
On Error Resume Next
    If CommandBar.Parent Is Nothing Then Exit Sub
    

    Select Case CommandBar.Parent.Id
        Case conMenu_Queue_LocateType
            With CommandBar.Controls
                If .Count = 0 Then '动态子菜单,扩1位
                    mstrFindWay = Replace(mstrFindWay, "姓名", "")
                    mstrFindWay = Replace(mstrFindWay, "排队号", "")
                    
                    mstrFindWay = "排队号,姓名," & mstrFindWay
                    aryKindInfo = Split(mstrFindWay, ",")
                    
                    For i = 0 To UBound(aryKindInfo)
                        If Trim(aryKindInfo(i)) <> "" Then
                            Set objControl = .Add(xtpControlButton, conMenu_Queue_LocateType * 10# + i + 1, aryKindInfo(i) & "(&" & IIf(i >= 9, Chr(65 + i - 9), i + 1) & ")"): objControl.Category = "CallFind"
                            If i = 0 Then objControl.Checked = True
                        End If
                    Next i
                End If
            End With
    End Select
End Sub


Private Sub Menu_View_Locate_Type_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next

    mstrLocateType = Split(Control.Caption, "(")(0)
    Call SaveSetting("ZLSOFT", gstrRegPath, "定位方式", mstrLocateType)
    
    cbrMain.RecalcLayout

    txtLocateValue.Text = ""
    txtLocateValue.PasswordChar = ""
    txtLocateValue.SetFocus
End Sub


Private Sub SwitchActiveWindow(ByVal blnIsCalledList As Boolean)
On Error Resume Next

    If blnIsCalledList Then
        scCallInf.GradientColorDark = &HFFC0C0
        scCallInf.GradientColorLight = &HFF8080

        scQueueInf.GradientColorDark = &HC0C0C0
        scQueueInf.GradientColorLight = &H808080
    Else
        scQueueInf.GradientColorDark = &HFFC0C0
        scQueueInf.GradientColorLight = &HFF8080

        scCallInf.GradientColorDark = &HC0C0C0
        scCallInf.GradientColorLight = &H808080
    End If
    
Err.Clear
End Sub

Public Sub comMenu_打号()
On Error GoTo errHandle
    Dim lngQueueId As Long
    Dim lngRowIndex As Long
    Dim strQueueName As String
    
    '获取可呼叫数据的行索引
    If CurQueueType = qftFindQueue Or CurQueueType = qftWaitQueue Then
        lngRowIndex = GetWaitQueueIndex()
    Else
        lngRowIndex = GetCalledQueueIndex()
    End If
           
    If lngRowIndex < 0 Then
        MsgBox "没有可供打印的队列数据，请刷新后重试。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "队列名称")
    
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otPrintNo) = False Then Exit Sub

    '调用打号功能
    If mobjQueueManage.PrintQueueNo(lngQueueId) = False Then Exit Sub

    Call DoWorkAfter(lngQueueId, strQueueName, otPrintNo)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub



Private Function GetWindowCaption() As String
'获取窗口标题
    GetWindowCaption = mobjOwner.Caption
End Function

Private Sub CopyWaitRowToCallRow(ByVal lngWaitRow As Long)
'从排队队列复制数据到呼叫队列
    Dim rptRecord As ReportRecord
    Dim objReportRecordItem As ReportRecordItem
    Dim lngWaitColIndex As Long
    Dim lngFirstColIndex As Long
    Dim i As Long
    
    Set rptRecord = rptCallList.Records.Insert(0)
    lngFirstColIndex = GetFirstDisplayColIndex(rptCallList)
    
    '创建默认的空数据
    For i = 0 To rptCallList.Columns.Count - 1
        Call rptRecord.AddItem("")
    Next i
    
    For i = 0 To rptCallList.Columns.Count - 1
        '查找呼叫队列对应的排队列索引
        lngWaitColIndex = GetColIndex(rptCallList.Columns(i).Caption, rptQueueList)
        
        If lngWaitColIndex >= 0 Then
            rptRecord(rptCallList.Columns(i).ItemIndex).value = rptQueueList.Rows(lngWaitRow).Record(lngWaitColIndex).value
            
            If rptQueueList.Rows(lngWaitRow).Record(lngWaitColIndex).Icon > 0 Then
                rptRecord(lngFirstColIndex).Icon = rptQueueList.Rows(lngWaitRow).Record(lngWaitColIndex).Icon
            End If
        End If
    Next
    
    rptCallList.Populate
End Sub


Private Sub CopyCallRowToWaitRow(ByVal lngCallRow As Long)
'从呼叫队列复制数据到排队队列
    Dim rptRecord As ReportRecord
    Dim objReportRecordItem As ReportRecordItem
    Dim lngCallColIndex As Long
    Dim lngFirstColIndex As Long
    Dim i As Long
    
    
    Set rptRecord = rptQueueList.Records.Insert(0)
    lngFirstColIndex = GetFirstDisplayColIndex(rptQueueList)
    
    '创建默认的空数据
    For i = 0 To rptQueueList.Columns.Count - 1
        Call rptRecord.AddItem("")
    Next i
    
    For i = 0 To rptQueueList.Columns.Count - 1
        lngCallColIndex = GetColIndex(rptQueueList.Columns(i).Caption, rptCallList)
        
        If lngCallColIndex >= 0 Then
            rptRecord(rptQueueList.Columns(i).ItemIndex).value = rptCallList.Rows(lngCallRow).Record(lngCallColIndex).value
            
            If rptCallList.Rows(lngCallRow).Record(lngCallColIndex).Icon > 0 Then
                rptRecord(lngFirstColIndex).Icon = rptCallList.Rows(lngCallRow).Record(lngCallColIndex).Icon
            End If
        End If
    Next
    
    rptQueueList.Populate
End Sub


Private Function CheckIsQueueing(ByVal lngQueueId As Long) As Boolean
'判断数据是否处于排队中
    Dim lngCurQueueState As Long
    
    lngCurQueueState = mobjQueueManage.GetQueueState(lngQueueId)
    
    CheckIsQueueing = IIf(lngCurQueueState = 0, True, False)
End Function

Public Function GetColumnIndex(ByVal lngQueueFromType As TQueueFromType, ByVal strColName As String) As Long
'取得对应列表的指定列索引
    If lngQueueFromType = qftFindQueue Or lngQueueFromType = qftWaitQueue Then
        GetColumnIndex = GetColIndex(strColName, rptQueueList)
    Else
        GetColumnIndex = GetColIndex(strColName, rptCallList)
    End If
End Function


Public Function GetRowIndex(ByVal lngQueueFromType As TQueueFromType, _
                            ByVal strColName As String, ByVal strValue As String) As Long
'获取对应值所在的行
    Dim objList As ReportControl
    Dim lngColIndex As Long
    Dim i As Long
    
    GetRowIndex = -1
    
    If lngQueueFromType = qftFindQueue Or lngQueueFromType = qftWaitQueue Then
        Set objList = rptQueueList
        lngColIndex = GetColIndex(strColName, rptQueueList)
    Else
        Set objList = rptCallList
        lngColIndex = GetColIndex(strColName, rptCallList)
    End If
    
    For i = 0 To objList.Rows.Count - 1
        If objList.Rows(i).GroupRow = False Then
            If objList.Rows(i).Record(lngColIndex).value = strValue Then
                GetRowIndex = objList.Rows(i).Index
                Exit Function
            End If
        End If
    Next i
    
End Function

Public Function GetCalledQueueIndex() As Long
'获取已呼叫队列的行选择索引
    Dim lngCalledQueueRowIndex As Long
    
    lngCalledQueueRowIndex = -1
    GetCalledQueueIndex = -1
    
    If rptCallList.SelectedRows.Count <= 0 Then Exit Function
    
    If rptCallList.SelectedRows(0).GroupRow <> True Then
        lngCalledQueueRowIndex = rptCallList.SelectedRows(0).Index
    Else
        lngCalledQueueRowIndex = rptCallList.SelectedRows(0).Childs(0).Index
    End If
    
    GetCalledQueueIndex = lngCalledQueueRowIndex
End Function

Public Function GetWaitQueueIndex() As Long
'取得直呼行索引
    Dim lngCallRowIndex As Long
    
    lngCallRowIndex = -1
    GetWaitQueueIndex = -1
    
    If rptQueueList.SelectedRows.Count <= 0 Then Exit Function
    
    If rptQueueList.SelectedRows(0).GroupRow <> True Then
        lngCallRowIndex = rptQueueList.SelectedRows(0).Index
    Else
        lngCallRowIndex = rptQueueList.SelectedRows(0).Childs(0).Index
    End If
    
    GetWaitQueueIndex = lngCallRowIndex
End Function


Public Sub DelQueueRecord(ByVal lngQueueFromType As TQueueFromType, ByVal lngRowIndex As Long)
'删除队列记录数据
    Dim lngRecordIndex As Long
    Dim objQueueList As ReportControl
    
    Select Case lngQueueFromType
        Case qftFindQueue, qftWaitQueue
            Set objQueueList = rptQueueList
        Case qftCalledQueue
            Set objQueueList = rptCallList
    End Select
    
    lngRecordIndex = objQueueList.Rows(lngRowIndex).Record.Index
    objQueueList.Rows(lngRowIndex).Selected = False
    
    Call objQueueList.Records.RemoveAt(lngRecordIndex)
    Call objQueueList.Populate
    
    If objQueueList.Rows.Count > lngRowIndex Then
        objQueueList.Rows(lngRowIndex).Selected = True
    End If
End Sub


Public Function GetListValue(ByVal lngQueueFromType As TQueueFromType, ByVal lngRowIndex As Long, ByVal strColName As String) As String
'获取列表行中对应的值
    Dim objCurQueueList As ReportControl
    Dim lngColIndex As Long
    
    GetListValue = ""
    
    Select Case lngQueueFromType
        Case TQueueFromType.qftFindQueue, TQueueFromType.qftWaitQueue
            Set objCurQueueList = rptQueueList
            lngColIndex = GetColIndex(strColName, rptQueueList)
        Case Else
            Set objCurQueueList = rptCallList
            lngColIndex = GetColIndex(strColName, rptCallList)
    End Select
    
    If objCurQueueList.Rows(lngRowIndex).GroupRow = True Then Exit Function
    
    GetListValue = objCurQueueList.Rows(lngRowIndex).Record(lngColIndex).value
End Function


Public Sub SetListValue(ByVal lngQueueFromType As TQueueFromType, ByVal lngRowIndex As Long, _
                            ByVal strColName As String, ByVal strValue As String)
'设置列表行中对应的值
    Dim objCurQueueList As ReportControl
    Dim lngColIndex As Long
    
    lngColIndex = -1
    
    Select Case lngQueueFromType
        Case TQueueFromType.qftFindQueue, TQueueFromType.qftWaitQueue
            Set objCurQueueList = rptQueueList
            lngColIndex = GetColIndex(strColName, rptQueueList)
        Case Else
            Set objCurQueueList = rptCallList
            lngColIndex = GetColIndex(strColName, rptCallList)
    End Select
    
    If lngColIndex < 0 Then Exit Sub
    
    objCurQueueList.Rows(lngRowIndex).Record(lngColIndex).value = strValue
End Sub


Public Sub Populate(Optional ByVal lngQueueFromType As TQueueFromType = -1)
'更新列表显示
     If lngQueueFromType = -1 Or lngQueueFromType = qftCalledQueue Then rptCallList.Populate
     If lngQueueFromType = -1 Or lngQueueFromType <> qftCalledQueue Then rptQueueList.Populate
End Sub

Public Function GetOrderCallIndex() As Long
'取得当前所选队列下的第一个可供呼叫的记录行索引
'顺乎的时候使用此方法依次获取待呼叫的排队数据

    Dim i As Long
    Dim lngQueueId As Long
    Dim lngCallRowIndex As Long         '待呼叫行索引
    Dim strQueueName As String
    Dim lngQueueNameColIndex As Long
    Dim lngRowIndex As Long
    Dim lngRecordIndex As Long
    
    
    lngCallRowIndex = -1
    GetOrderCallIndex = -1
    
    strQueueName = ""
    lngQueueNameColIndex = GetColIndex("队列名称", rptQueueList)
    
    '判断是否有排队记录被选中，如果存在，则使用选中的队列
    If rptQueueList.SelectedRows.Count > 0 Then
        '选择的记录不为分组行
        If rptQueueList.SelectedRows(0).GroupRow <> True Then
            strQueueName = rptQueueList.SelectedRows(0).Record(lngQueueNameColIndex).value
        Else
            strQueueName = rptQueueList.SelectedRows(0).Childs(0).Record(lngQueueNameColIndex).value
        End If
    Else
        If rptQueueList.Rows.Count <= 0 Then
            Exit Function
        End If
    End If
    
    lngRowIndex = 0
    
    '如果没有被选中的记录，则读取第一个队列的第一条记录
    Do While rptQueueList.Rows.Count > 0 And lngRowIndex < rptQueueList.Rows.Count
        If rptQueueList.Rows(lngRowIndex).GroupRow = True Then
            If rptQueueList.Rows(lngRowIndex).Childs(0).Record(lngQueueNameColIndex).value = strQueueName Or strQueueName = "" Then
                lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Childs(0).Record(GetColIndex("ID", rptQueueList)).value)
                
                '判断该数据是否能够进行呼叫
                If CheckIsQueueing(lngQueueId) Then
'                    '这里不进行数据删除处理，只有成功呼叫后，才删除数据
'                    lngRecordIndex = rptQueueList.Rows(lngRowIndex).Childs(0).Record.Index
'                    Call DelQueueRecord(rptQueueList, lngRecordIndex)
                    
                    lngCallRowIndex = rptQueueList.Rows(lngRowIndex).Childs(0).Index
                    
                    If rptQueueList.Rows.Count - 1 >= lngRowIndex Then
                        If rptQueueList.Rows(lngRowIndex).Childs.Count > 0 Then
                            rptQueueList.Rows(lngRowIndex).Childs(0).Selected = True
                        End If
                    Else
                        If rptQueueList.Rows.Count > 0 Then
                            rptQueueList.Rows(rptQueueList.Rows.Count - 1).Selected = True
                        End If
                    End If
                    
                    Exit Do
                Else
                    lngRecordIndex = rptQueueList.Rows(lngRowIndex).Childs(0).Index
                    Call DelQueueRecord(qftWaitQueue, lngRecordIndex)
                End If
            Else
                lngRowIndex = lngRowIndex + 1
            End If
        Else
            If rptQueueList.Rows(lngRowIndex).Record(lngQueueNameColIndex).value = strQueueName Or strQueueName = "" Then
                lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Record(GetColIndex("ID", rptQueueList)).value)
                
                
                '判断数据是否能够进行呼叫
                If CheckIsQueueing(lngQueueId) Then
'                    '这里不进行数据删除处理，只有成功呼叫后，才删除数据
'                    lngRecordIndex = rptQueueList.Rows(lngRowIndex).Record.Index
'                    Call DelQueueRecord(rptQueueList, lngRecordIndex)
                    
                    lngCallRowIndex = rptQueueList.Rows(lngRowIndex).Index
                    
                    If rptQueueList.Rows.Count - 1 >= lngRowIndex Then
                        rptQueueList.Rows(lngRowIndex).Selected = True
                    Else
                        If rptQueueList.Rows.Count > 0 Then
                            rptQueueList.Rows(rptQueueList.Rows.Count - 1).Selected = True
                        End If
                    End If
                    
                    Exit Do
                Else
                    lngRecordIndex = rptQueueList.Rows(lngRowIndex).Record.Index
                    Call DelQueueRecord(qftWaitQueue, lngRecordIndex)
                End If
            Else
                lngRowIndex = lngRowIndex + 1
            End If
        End If
    Loop
    
    GetOrderCallIndex = lngCallRowIndex
End Function


Private Sub rptCallList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnIsFindQueue Then Exit Sub
    
    If mblnIsSelectedCallingList = True Then
        RaiseEvent OnQueueListMouseDown(CurQueueType, Button, Shift, X, Y)
        Exit Sub
    End If

    mblnIsSelectedCallingList = True
    
    '当列表焦点改变后，重新配置列表标题的显示样式
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    Call DoQueueListChangeEvent(TQueueFromType.qftCalledQueue, rptCallList)
    
    RaiseEvent OnQueueListMouseDown(CurQueueType, Button, Shift, X, Y)
End Sub

Private Sub rptCallList_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
'触发OnGroupHint事件
On Error GoTo errHandle
    Dim objReportRow As ReportRow
    Dim objReportHitTest As ReportHitTestInfo
    Dim lngRowCount As Long
    Dim strGroupName As String
    
    Set objReportRow = Nothing
    
    If Not mblnInitOk Then Exit Sub
    
    lngRowCount = 0
    
    Set objReportHitTest = rptCallList.HitTest(X, Y)
    If Not objReportHitTest Is Nothing Then
        Set objReportRow = objReportHitTest.Row
        
        If Not objReportRow Is Nothing Then
            If objReportRow.GroupRow <> True Then
                lngRowCount = objReportRow.ParentRow.Childs.Count
                strGroupName = objReportRow.Record(GetColIndex("队列名称", rptCallList)).value
            Else
                '如果是分组，则直接获取分组下的行数量
                lngRowCount = objReportRow.Childs.Count
                strGroupName = objReportRow.Childs(0).Record(GetColIndex("队列名称", rptCallList)).value
            End If
        End If
    End If
    
    rptCallList.ToolTipText = IIf(lngRowCount <= 0, "", "[" & strGroupName & "] 已呼叫数量为：" & lngRowCount)
    RaiseEvent OnGroupHint(rptCallList.ToolTipText)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub rptCallList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    RaiseEvent OnQueueListMouseUp(CurQueueType, Button, Shift, X, Y)
End Sub

Private Sub rptCallList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim lngQueueId As Long
    
    '如果是队列数据查找状态，则呼叫列表将不使用
    If mblnIsFindQueue Then Exit Sub
    
    lngQueueId = Row.Record(GetColIndex("ID", rptCallList)).value
    
    RaiseEvent OnItemDblClick(qftCalledQueue, lngQueueId, Row, Item)
End Sub

Private Sub rptQueueList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnIsSelectedCallingList = False Then
        RaiseEvent OnQueueListMouseDown(CurQueueType, Button, Shift, X, Y)
        Exit Sub
    End If
    
    mblnIsSelectedCallingList = False
    
    '当列表焦点改变后，重新配置列表标题的显示样式
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    Call DoQueueListChangeEvent(IIf(mblnIsFindQueue, TQueueFromType.qftFindQueue, TQueueFromType.qftWaitQueue), rptQueueList)
    
    RaiseEvent OnQueueListMouseDown(CurQueueType, Button, Shift, X, Y)
End Sub


Private Sub DoQueueListChangeEvent(ByVal lngListType As TQueueFromType, objQueueList As Object)
    '触发列表切换事件
    RaiseEvent OnQueueListChange(lngListType, objQueueList)
End Sub



Private Sub rptQueueList_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
'触发OnGroupHint事件
On Error GoTo errHandle
    Dim objReportRow As ReportRow
    Dim objReportHitTest As ReportHitTestInfo
    Dim lngRowCount As Long
    Dim strGroupName As String
    
    Set objReportRow = Nothing
    
    If Not mblnInitOk Then Exit Sub
    
    lngRowCount = 0
    
    Set objReportHitTest = rptQueueList.HitTest(X, Y)
    If Not objReportHitTest Is Nothing Then
        Set objReportRow = objReportHitTest.Row
        
        If Not objReportRow Is Nothing Then
            If objReportRow.GroupRow <> True Then
                lngRowCount = objReportRow.ParentRow.Childs.Count
                strGroupName = objReportRow.Record(GetColIndex("队列名称", rptQueueList)).value
            Else
                '如果是分组，则直接获取分组下的行数量
                lngRowCount = objReportRow.Childs.Count
                strGroupName = objReportRow.Childs(0).Record(GetColIndex("队列名称", rptQueueList)).value
            End If
        End If
    End If
    
    rptQueueList.ToolTipText = IIf(lngRowCount <= 0, "", "[" & strGroupName & "] 剩余排队数量为：" & lngRowCount)
    RaiseEvent OnGroupHint(rptQueueList.ToolTipText)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub rptQueueList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    RaiseEvent OnQueueListMouseUp(CurQueueType, Button, Shift, X, Y)
End Sub

Private Sub rptQueueList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim lngQueueId As Long
    
    lngQueueId = Row.Record(GetColIndex("ID", rptQueueList)).value
    
    RaiseEvent OnItemDblClick(IIf(mblnIsFindQueue, qftFindQueue, qftWaitQueue), lngQueueId, Row, Item)
End Sub

Private Sub rptQueueList_SelectionChanged()
On Error GoTo errHandle
    Dim objReportRow As ReportRow
    Dim lngQueueId As Long
    
    Set objReportRow = Nothing
    
    If Not mblnInitOk Then Exit Sub
    
    If mblnIsSelectedCallingList <> False Then
        Call DoQueueListChangeEvent(IIf(mblnIsFindQueue, TQueueFromType.qftFindQueue, TQueueFromType.qftWaitQueue), rptQueueList)
    End If
    
    lngQueueId = 0
    
    If rptQueueList.SelectedRows.Count > 0 Then
        Set objReportRow = rptQueueList.SelectedRows(0)
        
        If objReportRow.GroupRow <> True Then
            lngQueueId = objReportRow.Record(GetColIndex("ID", rptQueueList)).value
        End If
    End If
    
    RaiseEvent OnSelectionChanged(IIf(mblnIsFindQueue, TQueueFromType.qftFindQueue, TQueueFromType.qftWaitQueue), lngQueueId, rptQueueList, objReportRow)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub rptCallList_SelectionChanged()
On Error GoTo errHandle
    Dim objReportRow As ReportRow
    Dim lngQueueId As Long

    If Not mblnInitOk Then Exit Sub
    If mblnIsFindQueue Then Exit Sub
    
    Set objReportRow = Nothing
    
    If mblnIsSelectedCallingList <> True Then
        Call DoQueueListChangeEvent(TQueueFromType.qftCalledQueue, rptCallList)
    End If

    If rptCallList.SelectedRows.Count > 0 Then
        Set objReportRow = rptCallList.SelectedRows(0)
        
        If objReportRow.GroupRow <> True Then
            lngQueueId = objReportRow.Record(GetColIndex("ID", rptCallList)).value
        End If
        
    End If
    
    RaiseEvent OnSelectionChanged(TQueueFromType.qftCalledQueue, lngQueueId, rptCallList, objReportRow)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_顺呼()
On Error GoTo errHandle
    Dim lngQueueId As Long
    Dim lngRowIndex As Long
    Dim blnCancel As Boolean
    Dim strCallContext As String
    Dim lngSendHwnd As Long
    Dim lngSendResult As Long
    
    '获取可呼叫数据的行索引
    lngRowIndex = GetOrderCallIndex()
           
    If lngRowIndex < 0 Then
        MsgBox "没有可供呼叫的队列数据，请刷新后重试。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Record(GetColIndex("ID", rptQueueList)).value)

    strCallContext = ""
    
    RaiseEvent OnCallPreBefore(lngQueueId, TCallWay.cwOrder, strCallContext, blnCancel)
    If blnCancel = True Then Exit Sub
    
'    If mobjQueueManage.CallTarget <> "" Then
'        '设置呼叫后所在的诊室目的地
'        Call mobjQueueManage.WriteTarget(lngQueueId)
'    End If
    
    '执行呼叫处理
    If mobjQueueManage.SpecifiedCall(lngQueueId, strCallContext) <= 0 Then Exit Sub
    
    '刷新以呼叫队列数据
    Call SetQueueRowState(rptQueueList, lngRowIndex, qsWaitCall)
    Call SetListValue(qftWaitQueue, lngRowIndex, "呼叫医生", mstrLoginUserName)
    Call SetListValue(qftWaitQueue, lngRowIndex, "呼叫时间", Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss"))
    
    Call CopyWaitRowToCallRow(lngRowIndex)
    Call rptCallList.Populate
    
    '删除已经呼叫的记录
    Call DelQueueRecord(qftWaitQueue, lngRowIndex)
    
    RaiseEvent OnCallPreAfter(lngQueueId, TCallWay.cwOrder)
    
    '呼叫后通知zlQueueShow，以便在显示有多页数据时，定位到当前呼叫病人。问题号:85290
    lngSendHwnd = FindWindow(vbNullString, "排队显示控制")
    
    If lngSendHwnd > 0 Then
        lngSendResult = PostMessage(lngSendHwnd, 1025, lngQueueId, 0)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function DoCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As TCallWay, ByRef strCallContext As String) As Boolean
'执行OnCallPreBefore事件
    Dim blnCancel As Boolean
    
    DoCallPreBefore = True
    blnCancel = False
    RaiseEvent OnCallPreBefore(lngQueueId, lngCallWay, strCallContext, blnCancel)
    
    DoCallPreBefore = Not blnCancel
End Function


Private Sub DoCallPreAfter(ByVal lngQueueId As Long, ByVal lngCallWay As TCallWay)
'执行OnCallPreAfter事件
    RaiseEvent OnCallPreAfter(lngQueueId, lngCallWay)
End Sub


Public Sub comMenu_直呼()
'调用类直接呼叫方法
On Error GoTo errHandle
    Dim lngQueueId As Long
    Dim lngRowIndex As Long
    Dim strCallContext As String
    Dim lngSendHwnd As Long
    Dim lngSendResult As Long
   
    lngRowIndex = GetWaitQueueIndex()
    
    If lngRowIndex < 0 Then
        MsgBox "没有需要呼叫的队列数据被选择。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Record(GetColIndex("ID", rptQueueList)).value)
    If Not CheckIsQueueing(lngQueueId) Then
        MsgBox "当前数据已被呼叫，请刷新后重试。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    '触发事件
    strCallContext = ""
    If DoCallPreBefore(lngQueueId, TCallWay.cwSpecify, strCallContext) = False Then Exit Sub
       
'    If mobjQueueManage.CallTarget <> "" Then
'        '设置呼叫后所在的诊室目的地
'        Call mobjQueueManage.WriteTarget(lngQueueId)
'    End If
    
    '执行呼叫处理
    If mobjQueueManage.SpecifiedCall(lngQueueId, strCallContext) <= 0 Then Exit Sub
    
    '刷新已呼叫列表数据显示
    Call SetQueueRowState(rptQueueList, lngRowIndex, qsWaitCall)
    Call SetListValue(qftWaitQueue, lngRowIndex, "呼叫医生", mstrLoginUserName)
    Call SetListValue(qftWaitQueue, lngRowIndex, "呼叫时间", Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss"))
    
    Call CopyWaitRowToCallRow(lngRowIndex)
    
    Call rptCallList.Populate
    
    '删除已经被呼叫的数据行
    Call DelQueueRecord(qftWaitQueue, lngRowIndex)
    
    Call DoCallPreAfter(lngQueueId, TCallWay.cwSpecify)
    
    '呼叫后通知zlQueueShow，以便在显示有多页数据时，定位到当前呼叫病人。问题号:85290
    lngSendHwnd = FindWindow(vbNullString, "排队显示控制")
    
    If lngSendHwnd > 0 Then
        lngSendResult = PostMessage(lngSendHwnd, 1025, lngQueueId, 0)
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_广播()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim strCallContext As String
    
    If mblnIsFindQueue Then
        lngRowIndex = GetWaitQueueIndex()
        If lngRowIndex > 0 Then lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Record(GetColIndex("ID", rptQueueList)).value)
    Else
        lngRowIndex = GetCalledQueueIndex()
        If lngRowIndex > 0 Then lngQueueId = Val(rptCallList.Rows(lngRowIndex).Record(GetColIndex("ID", rptCallList)).value)
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "没有可供呼叫的队列数据被选择。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If

    '执行广播处理
    If CheckIsQueueing(lngQueueId) Then
        MsgBox "当前数据处于排队状态，不能执行此操作。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    '触发事件
    strCallContext = ""
    If DoCallPreBefore(lngQueueId, TCallWay.cwBroadcast, strCallContext) = False Then Exit Sub
    
    If mobjQueueManage.BroadcastCall(lngQueueId, strCallContext) <= 0 Then Exit Sub
    
    Call DoCallPreAfter(lngQueueId, TCallWay.cwBroadcast)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
  
Public Sub comMenu_插队()
'执行队列插队操作
On Error GoTo errHandle
    Dim lngQueueId As Long
    Dim lngRowIndex As Long
    Dim strQueueName As String
    
    lngRowIndex = GetWaitQueueIndex()
    
    If lngRowIndex < 0 Then
        MsgBox "没有需要插队的队列数据被选择。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(GetListValue(qftWaitQueue, lngRowIndex, "ID"))
    strQueueName = GetListValue(qftWaitQueue, lngRowIndex, "队列名称")
        
    '触发事件
    If DoWorkBefore(qftWaitQueue, lngRowIndex, lngQueueId, TOperationType.otInsertQueue) = False Then Exit Sub
        
    If frmPriorityCause.ShowPriorityCause(Me, rptQueueList, lngRowIndex, mintWorkType, mstrReason) = True Then
        If GetWaitQueueSelState = qss排队中 Then
            Call LoadWaitQueueData
        End If
        
        Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otInsertQueue)
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Public Sub comMenu_重排()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim lngQueueSelState As Long
    Dim objRestoreQueue As New frmRestoreQueue

    Dim strCurQueueName As String
    Dim dtCurQueueDate As Date
    Dim strNewQueueNo As String
    Dim strQueueOrder As String
    Dim strNewQueueName As String
    Dim dtQueueDate As Date
    
    Dim lngMsgResult As Long
    
    '执行重排操作时，需要判断当前所选列表属于排队列表还是已呼叫列表，根据不同的列表取得需要重排的数据
    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
    Else
        lngRowIndex = GetWaitQueueIndex()
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "没有需要重排的队列数据被选择。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueSelState = GetWaitQueueSelState
    
    '获取所在队列的排队ID
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strCurQueueName = GetListValue(CurQueueType, lngRowIndex, "队列名称")
    
    '触发事件处理
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otRestore) = False Then Exit Sub
    
    dtCurQueueDate = Nvl(mobjQueueManage.GetQueueInf(lngQueueId, "排队时间")!排队时间, Now)
    
    Call objRestoreQueue.ShowRestoreQueueWindow(mstrQueryQueueNames, strCurQueueName, strNewQueueName, dtCurQueueDate, dtQueueDate, Me)
    If Trim(strNewQueueName) = "" Then Exit Sub
    
    '队列名称改变或者排队时间改变，都需要重新生成排队号码
    If strNewQueueName <> strCurQueueName Or dtQueueDate <> dtCurQueueDate Then
        '更改了当前的排队队列，需要产生新的排队号码
        strNewQueueNo = mobjQueueManage.GetQueueMaxNo(strNewQueueName, dtQueueDate)
        
        lngMsgResult = MsgBox("队列发生变化，已产生新的排队号码 [" & strNewQueueNo & "], 是否继续？", vbYesNo, "提示")
        If lngMsgResult = vbNo Then Exit Sub
        
        Call mobjQueueManage.UpdateQueue(lngQueueId, "队列名称=''" & strNewQueueName & "''" & IIf(strNewQueueName <> strCurQueueName, ",诊室=''''", "") & ",排队号码=''" & strNewQueueNo & "'',排队状态=-1,排队时间=To_Date(''" & dtQueueDate & "'', ''yyyy-mm-dd hh24:mi:ss'')")
        
        Call DoCreateQueueNo(lngQueueId, strNewQueueName, strNewQueueNo)
    End If
    
        
    '排队队列没有进行改变，重新在该队列排队，不需要产生新的排队号码
    strQueueOrder = mobjQueueManage.RestoreQueue(lngQueueId)
    If Trim(strQueueOrder) = "" Then Exit Sub
        
    strQueueOrder = FormatQueueOrder(strQueueOrder)
        
    If CurQueueType = qftCalledQueue Then
        Call DelQueueRecord(qftCalledQueue, lngRowIndex)
        If lngQueueSelState = qss排队中 Then
            Call LoadWaitQueueData
        End If

    ElseIf CurQueueType = qftWaitQueue Then
        If lngQueueSelState <> qss排队中 Then
            Call DelQueueRecord(qftWaitQueue, lngRowIndex)
        Else
            '重新载入排队队列数据
            Call LoadWaitQueueData
        End If
        
    Else
        '更改查找队列中的数据状态,如果更改了所在队列，则诊室为空
        If strNewQueueName <> strCurQueueName Then
            Call SetListValue(qftFindQueue, lngRowIndex, "诊室", "")
        End If
        
        Call SetListValue(qftFindQueue, lngRowIndex, "排队号码", strNewQueueNo)
        Call SetListValue(qftFindQueue, lngRowIndex, "排队序号", strQueueOrder)
        Call SetListValue(qftFindQueue, lngRowIndex, "队列名称", strNewQueueName)
        
        Call SetQueueRowState(rptQueueList, lngRowIndex, TQueueState.qsQueueing)
        
        Call rptQueueList.Populate
        
    End If
                    
    Call DoWorkAfter(lngQueueId, strNewQueueName, TOperationType.otRestore)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function DoWorkBefore(ByVal lngListType As TQueueFromType, ByVal lngListRow As Long, ByVal lngQueueId As Long, ByVal lngOperationType As TOperationType) As Boolean
'执行OnWorkBefore事件
    Dim blnCancel As Boolean
    
    DoWorkBefore = True
    blnCancel = False
    RaiseEvent OnWorkBefore(lngListType, lngListRow, lngQueueId, lngOperationType, blnCancel)
    
    DoWorkBefore = Not blnCancel
End Function


Private Sub DoWorkAfter(ByVal lngQueueId As Long, ByVal strCurQueueName As String, ByVal lngOperationType As TOperationType)
'执行OnWorkAfter事件
    RaiseEvent OnWorkAfter(lngQueueId, strCurQueueName, lngOperationType)
End Sub


Private Sub DoCreateQueueNo(ByVal lngQueueId As Long, ByVal strQueueName As String, ByRef strQueueNo As String)
'执行OnCreateQueueNo事件
    RaiseEvent OnCreateQueueNo(lngQueueId, strQueueName, strQueueNo)
End Sub


Private Sub AutoComplete()
'自动完成已接诊处理
'只能接诊属于自己呼叫的数据
    Dim i As Long
    Dim lngColIndex As Long
    Dim lngQueueId As Long
    Dim lngIdColIndex As Long
    Dim blnAutoAll As Boolean
    Dim lngQueueSelState As Long
    Dim lngQueueNameIndex As Long
    Dim strQueueName As String
    Dim lngQueueNoIndex As Long
    
    lngIdColIndex = GetColIndex("ID", rptCallList)
    
    lngColIndex = GetColIndex("排队状态", rptCallList)
    lngQueueNameIndex = GetColIndex("队列名称", rptCallList)
                    
    lngQueueSelState = GetWaitQueueSelState
    blnAutoAll = True
    
    For i = rptCallList.Rows.Count - 1 To 0 Step -1
        If Not rptCallList.Rows(i).GroupRow Then
            If rptCallList.Rows(i).Record(lngColIndex).value = "接诊中" And i <> rptCallList.SelectedRows(0).Index Then
                lngQueueId = rptCallList.Rows(i).Record(lngIdColIndex).value
                strQueueName = rptCallList.Rows(i).Record(lngQueueNameIndex).value
    
                '判断改检查是否属于自己呼叫的数据
                If Nvl(mobjQueueManage.GetQueueInf(lngQueueId, "呼叫医生")!呼叫医生) = mstrLoginUserName Then
                
                    '触发事件
                    If DoWorkBefore(CurQueueType, i, lngQueueId, TOperationType.otComplete) = False Then Exit Sub
        
                    Call mobjQueueManage.CompleteQueue(lngQueueId)
                    
                    Call SetQueueRowState(rptCallList, i, qsComplete)
                    
                    If lngQueueSelState = qss已完成 Then
                        Call CopyCallRowToWaitRow(i)
                    End If
                    
                    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otComplete)
                Else
                    blnAutoAll = False
                End If
            End If
        End If
    Next i
    
    '刷新列表
    Call Populate
    
    If Not blnAutoAll Then
        MsgBox "部分已接诊队列因不由[" & mstrLoginUserName & "]呼叫，未执行自动完成操作。", vbOKOnly, "提示"
    End If
End Sub


Private Sub DelCompleteQueue()
'删除完成队列数据
On Error Resume Next
    Dim i As Long
    Dim lngColIndex As Long
    Dim lngQueueNameColIndex As Long
    
    lngColIndex = GetColIndex("排队状态", rptCallList)
        
    For i = rptCallList.Rows.Count - 1 To 0 Step -1
        If Not rptCallList.Rows(i).GroupRow Then
            If rptCallList.Rows(i).Record(lngColIndex).value = "已完成" Then
                Call DelQueueRecord(qftCalledQueue, i)
            End If
        End If
    Next i
    
    rptCallList.Populate
Err.Clear
End Sub

Public Sub comMenu_接诊()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim strQueueName As String
    Dim objList As ReportControl
      

    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
        Set objList = rptCallList
    Else
        lngRowIndex = GetWaitQueueIndex()
        Set objList = rptQueueList
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "没有可供接诊的队列数据被选择。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "队列名称")
    
    
    '判断是否允许接诊处理
    If CheckIsQueueing(lngQueueId) Then
        MsgBox "当前数据处于排队状态，不能执行此操作。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    '触发事件
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otDiagnose) = False Then Exit Sub

    If Not mobjQueueManage.DiagnoseQueue(lngQueueId) Then Exit Sub
        
    If CurQueueType = qftCalledQueue Then
        '更改查找队列中的数据状态
        
        If mblnAutoComplete Then
            '将已接诊的数据修改为已完成
            Call AutoComplete
        End If
        
        Call SetQueueRowState(objList, lngRowIndex, TQueueState.qsDiagnose)
        
        If mblnAutoComplete Then
            Call DelCompleteQueue
        End If
    Else
        '排队队列需要更改数据后，转移到呼叫队列显示
        Call SetQueueRowState(objList, lngRowIndex, TQueueState.qsDiagnose)
        Call CopyWaitRowToCallRow(lngRowIndex)
        
        Call rptCallList.Populate
        
        '删除已经接诊的记录
        Call DelQueueRecord(qftWaitQueue, lngRowIndex)
    End If
        
    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otDiagnose)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_暂停()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim strQueueName As String
    Dim objList As ReportControl
    
    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
        Set objList = rptCallList
    Else
        lngRowIndex = GetWaitQueueIndex()
        Set objList = rptQueueList
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "没有需要暂停的队列数据被选择。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "队列名称")
    
    '触发事件处理
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otPause) = False Then Exit Sub
    
    If Not mobjQueueManage.PauseQueue(lngQueueId) Then Exit Sub
    
    Select Case CurQueueType
        Case qftCalledQueue
            '从已呼叫队列数据，转移到排队队列
            If GetWaitQueueSelState = qss已暂停 Then
                Call SetQueueRowState(objList, lngRowIndex, qsPause)
                Call CopyCallRowToWaitRow(lngRowIndex)
                
                Call rptQueueList.Populate
            End If
            
            Call DelQueueRecord(qftCalledQueue, lngRowIndex)
            
        Case qftFindQueue
            '更改查找队列中的数据状态
            Call SetQueueRowState(rptQueueList, lngRowIndex, TQueueState.qsPause)
            Call rptQueueList.Populate
        Case qftWaitQueue
            If optOutQueue(TQueueSelState.qss已暂停).value Then
                '直接更新列表状态
                Call SetQueueRowState(objList, lngRowIndex, qsPause)
                Call rptQueueList.Populate
            Else
                If GetWaitQueueSelState() <> qss已暂停 Then
                    Call DelQueueRecord(qftWaitQueue, lngRowIndex)
                End If
            End If
    End Select
    
    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otPause)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_弃号()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim strQueueName As String
    Dim objList As ReportControl
    
    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
        Set objList = rptCallList
    Else
        lngRowIndex = GetWaitQueueIndex()
        Set objList = rptQueueList
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "没有需要弃呼的队列数据被选择。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
       
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "队列名称")
    
    '触发事件处理
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otAbstain) = False Then Exit Sub
    
    If Not mobjQueueManage.AbstainQueue(lngQueueId) Then Exit Sub
    
    
    Select Case CurQueueType
        Case qftCalledQueue
            If GetWaitQueueSelState() = qss已弃号 Then
                Call SetQueueRowState(objList, lngRowIndex, qsAbstain)
                Call CopyCallRowToWaitRow(lngRowIndex)
                
                Call rptQueueList.Populate
            End If
            
            Call DelQueueRecord(qftCalledQueue, lngRowIndex)
        Case qftWaitQueue
            If optOutQueue(TQueueSelState.qss已弃号).value Then
                '直接更新列表状态
                Call SetQueueRowState(objList, lngRowIndex, qsAbstain)
                Call rptQueueList.Populate
            Else
                Call DelQueueRecord(qftWaitQueue, lngRowIndex)
            End If
        Case qftFindQueue
            '更改查找队列中的数据状态
            Call SetQueueRowState(rptQueueList, lngRowIndex, TQueueState.qsAbstain)
            Call rptQueueList.Populate
    End Select
    
    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otAbstain)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_恢复()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim lngQueueSelState As Long
    Dim strNewQueueNo As String
    Dim strQueueName As String
    
    '执行恢复操作时，需要判断当前所选列表属于排队列表还是已呼叫列表，根据不同的列表取得需要恢复的数据
    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
    Else
        lngRowIndex = GetWaitQueueIndex()
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "没有需要恢复的队列数据被选择。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueSelState = GetWaitQueueSelState
    
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "队列名称")
    
    '触发事件处理
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otStart) = False Then Exit Sub
    
    If Not mobjQueueManage.LineQueue(lngQueueId, strNewQueueNo) Then Exit Sub
    
    If strNewQueueNo <> "" Then
        Call DoCreateQueueNo(lngQueueId, strQueueName, strNewQueueNo)
    Else
        strNewQueueNo = GetListValue(CurQueueType, lngRowIndex, "排队号码")
    End If
        
    If CurQueueType = qftCalledQueue Then
        '刷新已呼叫列表数据显示
        Call SetQueueRowState(rptCallList, lngRowIndex, qsQueueing)
        
        If strNewQueueNo <> "" Then Call SetListValue(qftCalledQueue, lngRowIndex, "排队号码", strNewQueueNo)
        Call CopyCallRowToWaitRow(lngRowIndex)
        
        Call rptQueueList.Populate
        
        '删除已经被呼叫的数据行
        Call DelQueueRecord(qftCalledQueue, lngRowIndex)
        
    ElseIf CurQueueType = qftWaitQueue Then
        If lngQueueSelState <> qss排队中 Then
            Call DelQueueRecord(qftWaitQueue, lngRowIndex)
        End If
        
    Else
        '更改查找队列中的数据状态
        Call SetQueueRowState(rptQueueList, lngRowIndex, TQueueState.qsQueueing)
        
        If strNewQueueNo <> "" Then Call SetListValue(qftFindQueue, lngRowIndex, "排队号码", strNewQueueNo)
        Call SetListValue(qftFindQueue, lngRowIndex, "排队时间", Now)
        
        Call rptQueueList.Populate
        
    End If
    
    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otStart)
    
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub SetQueueRowState(objList As ReportControl, ByVal lngRow As Long, ByVal lngState As TQueueState)
'恢复排队显示状态
    Dim lngStateColIndex As Long
        
    lngStateColIndex = GetColIndex("排队状态", objList)
    If lngStateColIndex >= 0 Then
        objList.Rows(lngRow).Record(lngStateColIndex).value = Decode(lngState, _
                                                                    TQueueState.qsAbstain, "已弃呼", _
                                                                    TQueueState.qsCalling, "呼叫中", _
                                                                    TQueueState.qsCalled, "已呼叫", _
                                                                    TQueueState.qsComplete, "已完成", _
                                                                    TQueueState.qsPause, "已暂停", _
                                                                    TQueueState.qsQueueing, "排队中", _
                                                                    TQueueState.qsDiagnose, "接诊中", _
                                                                    TQueueState.qsWaitCall, "待呼叫", _
                                                                    "")
    End If
    
    lngStateColIndex = GetFirstDisplayColIndex(objList)
    
    If lngStateColIndex >= 0 Then
        Select Case lngState
            Case TQueueState.qsDiagnose
                objList.Rows(lngRow).Record(lngStateColIndex).Icon = M_LNG_ICON_DIAGNOSE
            Case TQueueState.qsCalling
                objList.Rows(lngRow).Record(lngStateColIndex).Icon = M_LNG_ICON_CALLING
            Case TQueueState.qsCalled
                objList.Rows(lngRow).Record(lngStateColIndex).Icon = M_LNG_ICON_CALLED
            Case TQueueState.qsQueueing
                objList.Rows(lngRow).Record(lngStateColIndex).Icon = M_LNG_ICON_QUEUEING
            Case Else   '暂停，完成，弃号均使用相同的图标
                objList.Rows(lngRow).Record(lngStateColIndex).Icon = M_LNG_ICON_QUEUEING
        End Select
        
    End If
End Sub


Public Sub comMenu_完成()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim lngQueueSelState As Long
    Dim strQueueName As String
    
    '执行恢复操作时，需要判断当前所选列表属于排队列表还是已呼叫列表，根据不同的列表取得需要恢复的数据
    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
    Else
        lngRowIndex = GetWaitQueueIndex()
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "没有需要完成的队列数据被选择。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueSelState = GetWaitQueueSelState
    
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "队列名称")
    
    If CheckIsQueueing(lngQueueId) Then
        MsgBox "当前数据处于排队状态，不能执行此操作。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    
    '触发事件
    If DoWorkBefore(qftCalledQueue, lngRowIndex, lngQueueId, TOperationType.otComplete) = False Then Exit Sub
    
    If Not mobjQueueManage.CompleteQueue(lngQueueId) Then Exit Sub
        
        
    If CurQueueType = qftCalledQueue Then
        If lngQueueSelState = qss已完成 Then
            Call SetQueueRowState(rptCallList, lngRowIndex, qsComplete)
            Call CopyCallRowToWaitRow(lngRowIndex)
            
            Call rptQueueList.Populate
        End If
        
        '从呼叫队列中删除已经完成的数据
        Call DelQueueRecord(qftCalledQueue, lngRowIndex)
        
    ElseIf CurQueueType = qftWaitQueue Then
        If optOutQueue(TQueueSelState.qss已完成).value Then
            '直接更新列表状态
            Call SetQueueRowState(rptQueueList, lngRowIndex, qsComplete)
            Call rptQueueList.Populate
        Else
            If lngQueueSelState <> qss已完成 Then
                Call DelQueueRecord(qftWaitQueue, lngRowIndex)
            End If
        End If
    Else
        '更改查找队列中的数据状态
        Call SetQueueRowState(rptQueueList, lngRowIndex, TQueueState.qsComplete)
        Call rptQueueList.Populate
        
    End If
    
    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otComplete)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_过滤()
On Error GoTo errHandle
    Dim strResult As String
    Dim strFilterWhere As String
    Dim strFilterValue As String
    Dim blnCancel As Boolean
    Dim blnUseCustom As Boolean
    Dim rsData As ADODB.Recordset
    
    blnUseCustom = False
    blnCancel = False
    
    RaiseEvent OnFilter(rsData, blnCancel, blnUseCustom)
    
    If blnCancel Then Exit Sub
    
    If Not blnUseCustom Then
        Call frmFilter.ShowFilterWindow(mobjQueueManage.DefQueryCols, strFilterWhere, strFilterValue, Me)
        If Trim(strFilterWhere) = "" Then Exit Sub
        
        RaiseEvent OnFindData(strFilterWhere, strFilterValue, Nothing, rsData, blnUseCustom)
        
        If Not blnUseCustom Then
            Set rsData = DefaultFind(strFilterWhere, strFilterValue)
        End If
        
    End If
    
    '删除当前数据
    Call rptQueueList.Records.DeleteAll
    Call rptQueueList.Populate
    
    Call rptCallList.Records.DeleteAll
    Call rptCallList.Populate
    
    '配置界面状态
    Call ConfigQueueStateSel(False)
    mblnIsFindQueue = True

    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub

    '载入查询数据
    Call LoadFindData(rptQueueList, rsData)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub comMenu_刷新()
On Error GoTo errHandle

    rptQueueList.Tag = ""
    rptCallList.Tag = ""
    
    '重新加载数据
    Call RefreshQueueData

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub ConfigDefaultInputPro(objInputCfg As Dictionary)
'载入默认录入项目
    'String,Number,Date,DateTime
    
    objInputCfg.Add "患者姓名", "STRING"
    objInputCfg.Add "排队号码", "STRING"
    objInputCfg.Add "诊室", "STRING"
    objInputCfg.Add "备注", "STRING"
    objInputCfg.Add "排队时间", "DATETIME"
    objInputCfg.Add "排队标记", "STRING"
End Sub

Public Sub comMenu_修改()
On Error GoTo errHandle
    Dim objUpdateWind As New frmUpdateInfo
    Dim lngQueueId As Long
    Dim lngRowIndex As Long
    Dim blnCancel As Boolean
    Dim blnUseCustom As Boolean
    

    Dim objInputCfg As New Dictionary
    Dim objReturn As New Dictionary
    Dim strKey As Variant
  
    lngQueueId = 0
    
    If CurQueueType = qftFindQueue Or CurQueueType = qftWaitQueue Then
        lngRowIndex = GetWaitQueueIndex()
    Else
        lngRowIndex = GetCalledQueueIndex()
    End If
           
    If lngRowIndex < 0 Then
        MsgBox "请选择需要修改的队列记录。", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    If CurQueueType = qftCalledQueue Then
        lngQueueId = Val(rptCallList.Rows(lngRowIndex).Record(GetColIndex("ID", rptCallList)).value)
    Else
        lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Record(GetColIndex("ID", rptQueueList)).value)
    End If
    
    blnCancel = False
    blnUseCustom = False
    
    Call ConfigDefaultInputPro(objInputCfg)
    
    RaiseEvent OnModifyBefore(CurQueueType, lngQueueId, objInputCfg, blnCancel, blnUseCustom)
    
    If blnCancel = True Then Exit Sub
    If blnUseCustom = True Then Exit Sub
    
    If objUpdateWind.zlShowMe(lngQueueId, objInputCfg, objReturn, mobjQueueManage, Me) = True Then
                      
        RaiseEvent OnModifyAfter(lngQueueId, objReturn)
        
        '同步更新列表中的数据
        For Each strKey In objReturn.Keys
            Call SetListValue(CurQueueType, lngRowIndex, strKey, objReturn.Item(strKey))
        Next
                
        Call Populate(CurQueueType)
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Public Sub comMenu_设置()
On Error GoTo errHandle
    '调用打开参数配置界面
    Dim blnUseCustom As Boolean
    
    blnUseCustom = False
    RaiseEvent OnConfigEvent(blnUseCustom)
    
    If blnUseCustom Then Exit Sub
    
    If ShowVoiceConfig Then
        Call ApplyVoiceConfig
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


'开始数据轮询显示
Private Sub StartCall()
    tmrBroadCast.Enabled = True
End Sub


'终止数据轮询显示
Private Sub AbortCall()
    tmrBroadCast.Enabled = False
End Sub


'轮询数据显示和呼叫
Private Sub LoopPlayVoice()
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngVoiceId As Long
    Dim lngQueueId As Long
    Dim strVoiceContext As String
    Dim blnCancel As Boolean
    Dim i As Long
    Dim lngRowIndex As Long
    Dim blnAllowQuery As Boolean
    
    '判断是否需要查询数据库数据
    blnAllowQuery = IIf(mrsVoiceContext Is Nothing, True, False)
    
    If Not mrsVoiceContext Is Nothing Then
        blnAllowQuery = IIf(mrsVoiceContext.RecordCount <= 0 Or mrsVoiceContext.EOF, True, False)
    End If
    
    If blnAllowQuery Then
        If Timer < mdtLastVoiceDate + mlngInterval Then Exit Sub
        mdtLastVoiceDate = Timer
        
        '查询需要播放的语音数据
        strSql = "select id,队列ID,呼叫内容,生成时间 from 排队语音呼叫  where 站点=[1] order by 生成时间"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询语音呼叫内容", mstrComputerName)
        
        If rsData.RecordCount > 0 Then
            Set mrsVoiceContext = zlDatabase.CopyNewRec(rsData)
            mrsVoiceContext.Sort = "生成时间 asc"
        End If
    End If
    
    If mrsVoiceContext Is Nothing Then Exit Sub
    If mrsVoiceContext.RecordCount <= 0 Or mrsVoiceContext.EOF Then Exit Sub
    
'    mdtLastVoiceDate = Timer
    
    lngVoiceId = Val(Nvl(mrsVoiceContext!Id))
    lngQueueId = Val(Nvl(mrsVoiceContext!队列ID))
    strVoiceContext = Nvl(mrsVoiceContext!呼叫内容)
    
    Call mrsVoiceContext.MoveNext

    blnCancel = False
    RaiseEvent OnPlayVoiceBefore(lngVoiceId, lngQueueId, strVoiceContext, blnCancel)
    
    If blnCancel Then Exit Sub
    
    If lngQueueId <= 0 Then
        '播放自定义的呼叫内容
        Call mobjQueueManage.PlayCustomVoice(lngVoiceId, False, strVoiceContext)
    Else
        '设置呼叫行的当前颜色
        lngRowIndex = GetRowIndex(qftCalledQueue, "ID", lngQueueId)
        
        If lngRowIndex >= 0 Then
            Call SetQueueRowState(rptCallList, lngRowIndex, qsCalling)
            Call rptCallList.Populate
        End If
        
        '播放语音
        Call mobjQueueManage.PlayQueueVoice(lngVoiceId, lngQueueId, False, strVoiceContext)
    End If
    
    RaiseEvent OnPlayVoiceAfter(lngVoiceId, lngQueueId, strVoiceContext)

    '呼叫成功后删除呼叫过的内容
    Call mobjQueueManage.DelVoiceData(lngVoiceId)
    
    If lngQueueId > 0 Then
        '设置呼叫行的当前颜色
        lngRowIndex = GetRowIndex(qftCalledQueue, "ID", lngQueueId)
                
        If lngRowIndex >= 0 Then
            If GetListValue(qftCalledQueue, lngRowIndex, "排队状态") = "呼叫中" Then
                Call SetQueueRowState(rptCallList, lngRowIndex, qsCalled)
                Call rptCallList.Populate
            End If
        End If
    End If
    
    
End Sub


Public Sub StartVoice()
'开始语音播放
    mdtLastVoiceDate = Timer - mlngInterval
    
    tmrBroadCast.Interval = mlngInterval
    tmrBroadCast.Enabled = True
    tmrBroadCast.Tag = 0
End Sub

Public Sub StopVoice()
'结束语音播放
On Error GoTo errHandle
    tmrBroadCast.Tag = 1
    tmrBroadCast.Enabled = False
    
    If Not mobjQueueManage Is Nothing Then Call mobjQueueManage.StopVoice
Exit Sub
errHandle:
    Debug.Print "StopVoice Err:" & Err.Description
End Sub


Private Sub timerCard_Timer()
On Error GoTo errHandle
    If GetTickCount - mlngStartTime > 200 Then
        '大于200毫秒时，自动认为刷卡结束
        timerCard.Enabled = False
        
        mlngStartTime = 0
        mlngAvgTime = 0
        mlngReadCount = 0
        
        Call zlControl.TxtSelAll(txtLocateValue)
        
        If IsFindModel Then
            '进入数据查找
            Call FindQueueData(mstrLocateType, txtLocateValue.Text)
        Else
            '进入数据定位
            Call LocateQueueData(mstrLocateType, txtLocateValue.Text)
        End If
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub LocateQueueData(ByVal strLocateType As String, ByVal strFindValue As String)
    Dim blnUseCustom As Boolean
    Dim lngQueueId As Long
    Dim i As Long, j As Long
    Dim lngIdColIndex As Long
    Dim blnOldBold As Boolean

    
    If Trim(strFindValue) = "" Then Exit Sub
    
    blnUseCustom = False
    RaiseEvent OnLocateData(strLocateType, strFindValue, txtLocateValue, lngQueueId, blnUseCustom)
    
    If Not blnUseCustom Then
        lngQueueId = DefaultLocate(strLocateType, strFindValue)
    End If
    

    Call rptQueueList.SelectedRows.DeleteAll
    Call rptCallList.SelectedRows.DeleteAll
    
    
    If lngQueueId <= 0 Then Exit Sub
    
    
    lngIdColIndex = GetColIndex("ID", rptQueueList)
    
    For i = 0 To rptQueueList.Rows.Count - 1
        If rptQueueList.Rows(i).GroupRow = True Then
            For j = 0 To rptQueueList.Rows(i).Childs.Count - 1
                If rptQueueList.Rows(i).Childs(j).Record(lngIdColIndex).value = lngQueueId Then
                    Call zlControl.TxtSelAll(txtLocateValue)
                    
                    rptQueueList.Rows(i).Expanded = True
                    rptQueueList.Rows(i).Childs(j).Selected = True
                    
                    Set rptQueueList.FocusedRow = rptQueueList.Rows(i).Childs(j)
        
                    mblnIsSelectedCallingList = False
                    Call SwitchActiveWindow(mblnIsSelectedCallingList)
                    
                    Exit Sub
                End If
            Next j
        End If
    Next i
    
    
    lngIdColIndex = GetColIndex("ID", rptCallList)
    
    For i = 0 To rptCallList.Rows.Count - 1
        If rptCallList.Rows(i).GroupRow = True Then
            For j = 0 To rptCallList.Rows(i).Childs.Count - 1
                If rptCallList.Rows(i).Childs(j).Record(lngIdColIndex).value = lngQueueId Then
                    Call zlControl.TxtSelAll(txtLocateValue)
                    
                    rptCallList.Rows(i).Expanded = True
                    rptCallList.Rows(i).Childs(j).Selected = True
                    
                    Set rptCallList.FocusedRow = rptCallList.Rows(i).Childs(j)
        
                    mblnIsSelectedCallingList = True
                    Call SwitchActiveWindow(mblnIsSelectedCallingList)
                    
                    Exit Sub
                End If
            Next j
        End If
    Next i
    
    
End Sub


Private Sub ConfigQueueStateSel(ByVal blnEnable As Boolean)
''配置队列过滤状态是否允许设置
'    lblQueueFilter(0).Enabled = blnEnable
'    lblQueueFilter(1).Enabled = blnEnable
'    lblQueueFilter(2).Enabled = blnEnable
'    lblQueueFilter(3).Enabled = blnEnable
'
'    optOutQueue(0).Enabled = blnEnable And optOutQueue(0).value = 0
'    optOutQueue(1).Enabled = blnEnable And optOutQueue(1).value = 0
'    optOutQueue(2).Enabled = blnEnable And optOutQueue(2).value = 0
'    optOutQueue(3).Enabled = blnEnable And optOutQueue(3).value = 0
'
'    If mblnIsShowCalledQueue = True Then
'        DkpMain.Panes(2).Closed = Not blnEnable
'    End If
    
End Sub


Public Sub FindQueueData(ByVal strLocateType As String, ByVal strFindValue As String)
    Dim blnUseCustom As Boolean
    Dim rsData As ADODB.Recordset
    
    If Trim(strFindValue) = "" Then
        '如果没有录入查询数据，则表示刷新数据
        Call RefreshQueueData
        Exit Sub
    End If
    
    blnUseCustom = False
    RaiseEvent OnFindData(strLocateType, strFindValue, txtLocateValue, rsData, blnUseCustom)
    
    If Not blnUseCustom Then
        '使用默认的查询
        Set rsData = DefaultFind(strLocateType, strFindValue)
    End If

    Call rptQueueList.Records.DeleteAll
    Call rptQueueList.Populate
    
    Call rptCallList.Records.DeleteAll
    Call rptCallList.Populate
    
    Call ConfigQueueStateSel(False)
    mblnIsFindQueue = True

    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub

    Call LoadFindData(rptQueueList, rsData)
    
End Sub


Private Sub LoadFindData(objQueueList As ReportControl, rsData As ADODB.Recordset)
'载入队列数据

On Error GoTo errHandle
    Dim rptRecord As ReportRecord
    Dim blnCancel As Boolean
    Dim i As Long

'加载数据到列表

    Call objQueueList.Records.DeleteAll
    Call objQueueList.Populate
    
    If rsData.RecordCount <= 0 Then Exit Sub

    While Not rsData.EOF

        blnCancel = False
        RaiseEvent OnReadBefore(rsData, TQueueFromType.qftFindQueue, blnCancel)
        
        If Not blnCancel Then
            Set rptRecord = objQueueList.Records.Add
            
            For i = 0 To objQueueList.Columns.Count - 1
                rptRecord.AddItem ""
            Next
    
            Call SetReportRecordItem(objQueueList, rptRecord, rsData)
            
            RaiseEvent OnReadAfter(rsData, TQueueFromType.qftFindQueue, rptRecord)
        End If

        rsData.MoveNext
    Wend

    objQueueList.Populate

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function DefaultFind(ByVal findType As String, ByVal findData As String) As ADODB.Recordset
    Dim strSql As String, strFilter As String
    Dim strQueueNames As String
    Dim varValue1 As Variant
    Dim varValue2 As Variant
    
    On Error GoTo errHandle
    
    strFilter = ""
    varValue1 = findData
    varValue2 = ""
    
    Select Case findType  ' '0-排队号;1-姓名;
    Case "排队号", "排队号码"
        varValue1 = findData
        
        If mblnIsReleationQueueTag Then
            strFilter = " and upper(排队标记)||upper(排队号码) = upper([2])"
        Else
            strFilter = " and upper(排队号码) = upper([2])"
        End If
    Case "姓名", "患者姓名"
        varValue1 = findData & "%"
        strFilter = " and upper(患者姓名) Like upper([2])"

    Case Else
        strFilter = " and upper(" & findType & ")=upper([2])"
        
        Select Case True
            Case IsNumeric(findData)
                varValue1 = Val(findData)
            Case IsDate(findData)
                
                If Format(findData, "hh:mm:ss") = "00:00:00" Then
                    varValue1 = CDate(Format(findData, "yyyy-mm-dd 00:00:00"))
                    varValue2 = CDate(Format(findData, "yyyy-mm-dd 23:59:59"))
                    strFilter = " and " & findType & " between [2] and [3] "
                Else
                    varValue1 = CDate(findData)
                    strFilter = " and " & findType & " = [2]"
                End If
                
            Case Else
                varValue1 = findData
        End Select
        
        
    End Select
    
    strQueueNames = mstrQueryQueueNames
    
    If strQueueNames <> "" Then
        strQueueNames = Replace(strQueueNames, ",", "','")
        strFilter = strFilter & " and 队列名称 in ('" & strQueueNames & "') "
    End If
    
    strSql = "select * from 排队叫号队列　where  业务类型=[1] " & strFilter & " order by 排队号码 "

    Set DefaultFind = zlDatabase.OpenSQLRecord(strSql, "按默认方式查找队列", mintWorkType, varValue1, varValue2)
     
    Exit Function
errHandle:
    Set DefaultFind = Nothing
    If ErrCenter = 1 Then Resume

End Function


Private Sub tmrBroadCast_Timer()
On Error GoTo errHandle
    If Val(tmrBroadCast.Tag) = 1 Then Exit Sub
    
    '停止轮训
    Call AbortCall
    
    If Val(tmrBroadCast.Tag) <> 1 Then
        '调用轮训方法
        Call LoopPlayVoice
    End If
    
    If Val(tmrBroadCast.Tag) = 1 Then Exit Sub
    
    '开始轮训
    Call StartCall

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitLocalParas()
On Error GoTo errHandle
    Dim i As Integer
    Dim X As Long, Y As Long, r As Long, b As Long

    mstrLocateType = GetSetting("ZLSOFT", gstrRegPath, "定位方式", "姓名")

'    mlngQueueW1 = GetSetting("ZLSOFT", gstrRegPath, "排队队列显示宽度", Round(Width / 3 * 2))
'    mlngQueueW2 = GetSetting("ZLSOFT", gstrRegPath, "呼叫队列显示宽度", Round(Width / 3))

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function DefaultLocate(ByVal strFindType As String, ByVal strFindData As String) As Long
    Dim i As Integer
    Dim j As Integer
    Dim rsData As ADODB.Recordset
    Dim strSql As String
    Dim lngPatientId As Long
    Dim blnFind As Boolean
    Dim lngFindColIndex As Long
    Dim lngStartIndex As Long
    Dim blnExpandState As Boolean
    
    DefaultLocate = 0
    lngFindColIndex = -1
    
    If strFindType = "排队号" Or strFindData = "排队号码" Then
        lngFindColIndex = GetColIndex("排队号码", rptQueueList)
    ElseIf strFindType = "姓名" Or strFindData = "患者姓名" Then
        lngFindColIndex = GetColIndex("患者姓名", rptQueueList)
    Else
        lngFindColIndex = GetColIndex(strFindType, rptQueueList)
    End If
    
    If lngFindColIndex < 0 Then Exit Function
    
    '获取开始查找的数据行
    If rptQueueList.SelectedRows.Count > 0 Then
        mlngLocateRowIndex = rptQueueList.SelectedRows(0).Index + 1
    End If
    
    If mlngLocateRowIndex >= rptQueueList.Rows.Count + rptCallList.Rows.Count - 1 Then mlngLocateRowIndex = 0
    
   
    blnFind = False
    For i = 0 To rptQueueList.Rows.Count - 1
        If rptQueueList.Rows(i).GroupRow = True Then
            blnExpandState = rptQueueList.Rows(i).Expanded
            rptQueueList.Rows(i).Expanded = True
            
            For j = 0 To rptQueueList.Rows(i).Childs.Count - 1
                If strFindType = "姓名" Or strFindType = "患者姓名" Then
                    blnFind = IIf(rptQueueList.Rows(i).Childs(j).Index >= mlngLocateRowIndex And UCase(rptQueueList.Rows(i).Childs(j).Record(lngFindColIndex).value) Like UCase(strFindData) & "*", True, False)
                Else
                    blnFind = IIf(rptQueueList.Rows(i).Childs(j).Index >= mlngLocateRowIndex And UCase(rptQueueList.Rows(i).Childs(j).Record(lngFindColIndex).value) = UCase(strFindData), True, False)
                End If
        
                If blnFind Then
                    DefaultLocate = rptQueueList.Rows(i).Childs(j).Record(GetColIndex("ID", rptQueueList)).value
                    Exit Function
                End If
            Next j
            
            rptQueueList.Rows(i).Expanded = blnExpandState
        End If
    Next i


    '获取开始查找的数据行
    If rptCallList.SelectedRows.Count > 0 Then
        mlngLocateRowIndex = rptQueueList.Rows.Count + rptCallList.SelectedRows(0).Index + 1
    End If
    
    If mlngLocateRowIndex > rptQueueList.Rows.Count + rptCallList.Rows.Count - 1 Then
        mlngLocateRowIndex = 0
        Exit Function
    End If
    
    '如果没有找到数据，则从已呼叫队列中查找
    For i = 0 To rptCallList.Rows.Count - 1
        If rptCallList.Rows(i).GroupRow = True Then
            blnExpandState = rptCallList.Rows(i).Expanded
            rptCallList.Rows(i).Expanded = True
            
            For j = 0 To rptCallList.Rows(i).Childs.Count - 1
                If strFindType = "姓名" Or strFindType = "患者姓名" Then
                    blnFind = IIf(rptCallList.Rows(i).Childs(j).Index >= mlngLocateRowIndex - rptQueueList.Rows.Count And UCase(rptCallList.Rows(i).Childs(j).Record(lngFindColIndex).value) Like UCase(strFindData) & "*", True, False)
                Else
                    blnFind = IIf(rptCallList.Rows(i).Childs(j).Index >= mlngLocateRowIndex - rptQueueList.Rows.Count And UCase(rptCallList.Rows(i).Childs(j).Record(lngFindColIndex).value) = UCase(strFindData), True, False)
                End If
        
                If blnFind Then
                    DefaultLocate = rptCallList.Rows(i).Childs(j).Record(GetColIndex("ID", rptCallList)).value
                    Exit Function
                End If
            Next j
            
            rptCallList.Rows(i).Expanded = blnExpandState
        End If
    Next i
    
    mlngLocateRowIndex = 0
    
End Function


Private Sub txtLocateValue_GotFocus()
    On Error Resume Next
    
    Call zlControl.TxtSelAll(txtLocateValue)
End Sub


Private Sub txtLocateValue_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        If IsFindModel Then
            '进入查找
            Call FindQueueData(mstrLocateType, txtLocateValue.Text)
        Else
            '进入定位
            Call LocateQueueData(mstrLocateType, txtLocateValue.Text)
        End If
        
        Exit Sub
    End If
    
    If KeyAscii = 8 Then Exit Sub
    
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    mlngReadCount = mlngReadCount + 1
    If mlngStartTime <> 0 Then
        If GetTickCount - mlngStartTime > 200 Then
            mlngReadCount = 1
            mlngAvgTime = 0
        Else
            mlngAvgTime = mlngAvgTime + (GetTickCount() - mlngStartTime)
        End If
    End If
    
    mlngStartTime = GetTickCount
    
    '取三次平均录入时间
    If mlngReadCount = 3 Then
        mlngAvgTime = Fix(mlngAvgTime / 3)
        
        If mlngAvgTime <= 30 Then timerCard.Enabled = True
    End If

End Sub


Private Sub UserControl_Initialize()
    mblnIsShowBars = True
    mblnIsShowCalledQueue = True
    mlngInterval = 30000    '默认30秒轮询一次
    mblnIsFindQueue = False
    mstrReason = ""
    mblnAutoComplete = True
    mblnShowMySelfCalled = True
    mblnIsReleationQueueTag = False
    
    Set mrsVoiceContext = Nothing
    Set mobjQueueManage = New clsQueueOperation
    
    InitFaceScheme

End Sub

Private Sub UserControl_Resize()
'调整控件位置方法
On Error Resume Next
    Call picCallFace_Resize
    Call picQueueFace_Resize
Err.Clear
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '参与查询的数据字段
    mstrDataFields = PropBag.ReadProperty("DataFields", "")

    '排队列表显示字段
    mstrDisplayQueueFields = PropBag.ReadProperty("DisplayQueueColNames", "")

    '呼叫列表显示字段
    mstrDisplayCallFields = PropBag.ReadProperty("DisplayCallColNames", "")

    '允许显示的队列名称  注：如果为空，则显示当前业务类型下的所有队列中的排队和呼叫数据
    mstrQueryQueueNames = PropBag.ReadProperty("QueryQueueNames", "")

    '组名(对配对数据分组显示)
    mstrGroupField = PropBag.ReadProperty("GroupField", "")
    
    '自定义排序列
    mstrCustomOrderColName = PropBag.ReadProperty("CustomOrderName", "")
    
    '附件查找方式
    mstrFindWay = PropBag.ReadProperty("FindWay", "")
    
    '呼叫轮询间隔时间
    mlngInterval = PropBag.ReadProperty("Interval", 30)
    
    '数据有效保存天数
    mobjQueueManage.ValidDays = PropBag.ReadProperty("ValidDays", 1)
    
    '是否显示工具栏
    IsShowBars = PropBag.ReadProperty("IsShowBars", True)
    
    '是否显示已呼叫队列
    IsShowCalledQueue = PropBag.ReadProperty("IsShowCalledQueue", True)
    
    '设置呼叫后的目的地
    mobjQueueManage.CallTarget = PropBag.ReadProperty("CalledTarget", "")
    
    '是否显示按钮文本
    mlngMenuCaptionStyle = PropBag.ReadProperty("IsShowToolText", xtpButtonIconAndCaption)
    IsShowToolText = IIf(mlngMenuCaptionStyle = xtpButtonIcon, False, True)
    
    '是否显示大图标
    IsIconLarge = PropBag.ReadProperty("IsIconLarge", True)
    
    '是否显示插队原因
    mstrReason = PropBag.ReadProperty("Reason", "")
End Sub

Private Sub UserControl_Terminate()
    
    Call StopVoice
    
    SaveColWidth rptQueueList
    SaveColWidth rptCallList
    
    If gstrRegPath <> "" Then
        SaveSetting "ZLSOFT", gstrRegPath, "QueueListWidthRate", picQueueFace.Width / ScaleWidth
    End If
        
    Set mobjQueueManage = Nothing

    Unload frmPriorityCause
    Unload frmSetup
    Unload frmFilter
End Sub

Private Sub SaveColWidth(objQueueList As Object)
'保存列的宽度
    Dim strColPro As String
    Dim i As Long
    
    If gstrRegPath = "" Then Exit Sub
    
    For i = 0 To objQueueList.Columns.Count - 1
        If objQueueList.Columns(i).Visible = True Then
            If strColPro <> "" Then strColPro = strColPro & ";"
            
            strColPro = strColPro & objQueueList.Columns(i).Caption & ":" & objQueueList.Columns(i).Width
        End If
    Next i
    
    SaveSetting "ZLSOFT", gstrRegPath, objQueueList.Name, strColPro
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DataFields", mstrDataFields, "")
    Call PropBag.WriteProperty("DisplayQueueColNames", mstrDisplayQueueFields, "")
    Call PropBag.WriteProperty("DisplayCallColNames", mstrDisplayCallFields, "")
    Call PropBag.WriteProperty("QueryQueueNames", mstrQueryQueueNames, "")
    Call PropBag.WriteProperty("GroupField", mstrGroupField, "")
    Call PropBag.WriteProperty("CustomOrderName", mstrCustomOrderColName, "")
    Call PropBag.WriteProperty("FindWay", mstrFindWay, "")
    Call PropBag.WriteProperty("Interval", mlngInterval, 30)
    Call PropBag.WriteProperty("ValidDays", mobjQueueManage.ValidDays, 1)
    Call PropBag.WriteProperty("IsShowBars", mblnIsShowBars, True)
    Call PropBag.WriteProperty("IsShowCalledQueue", mblnIsShowCalledQueue, True)
    Call PropBag.WriteProperty("IsShowToolText", mlngMenuCaptionStyle, xtpButtonIconAndCaption)
    Call PropBag.WriteProperty("IsIconLarge", cbrMain.Options.LargeIcons, True)
    Call PropBag.WriteProperty("Reasons", mstrReason, "")
    Call PropBag.WriteProperty("CalledTarget", mobjQueueManage.CallTarget, "")
End Sub


