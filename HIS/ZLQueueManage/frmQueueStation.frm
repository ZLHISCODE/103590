VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmQueueStation 
   BorderStyle     =   0  'None
   Caption         =   "排队叫号"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10575
   ControlBox      =   0   'False
   Icon            =   "frmQueueStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin zlIDKind.PatiIdentify Pati 
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmQueueStation.frx":0CCA
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindAppearance=   0
      ShowSortName    =   -1  'True
      DefaultCardType =   "就诊卡"
      IDKindWidth     =   900
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowAutoCommCard=   -1  'True
      NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
   End
   Begin VB.PictureBox picCallFace 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   5160
      ScaleHeight     =   4455
      ScaleWidth      =   3735
      TabIndex        =   6
      Top             =   600
      Width           =   3735
      Begin XtremeReportControl.ReportControl rptCallList 
         Height          =   3855
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   6800
         _StockProps     =   0
         BorderStyle     =   3
         AllowColumnSort =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption scCallInf 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "呼叫列表："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
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
   Begin VB.PictureBox picLabel 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10575
      TabIndex        =   0
      Top             =   5670
      Width           =   10575
      Begin VB.CheckBox chkOutQueue 
         Caption         =   "已弃号"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   855
      End
      Begin VB.CheckBox chkOutQueue 
         Caption         =   "已暂停"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkOutQueue 
         Caption         =   "已完成"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
      Begin VB.Label labError 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   0
         Width           =   6975
      End
   End
   Begin VB.Timer tmrBroadCast 
      Interval        =   30000
      Left            =   4440
      Top             =   0
   End
   Begin VB.PictureBox picQueueFace 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   600
      ScaleHeight     =   4455
      ScaleWidth      =   3735
      TabIndex        =   4
      Top             =   600
      Width           =   3735
      Begin XtremeReportControl.ReportControl rptQueueList 
         Height          =   3975
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   7011
         _StockProps     =   0
         BorderStyle     =   3
         AllowColumnSort =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption scQueueInf 
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "排队列表："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16761024
         GradientColorDark=   16744576
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   360
      Top             =   720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "frmQueueStation.frx":0D7D
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQueueStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'功能说明
'
'广播：不更新数据库中的相关数据，直接组织数据到“排队语音呼叫”中，可直接对队列中的数据进行广播
'直呼：触发事件，根据需要更新数据库中的数据，可选择对队列中的任何一条数据进行直呼
'顺呼：按队列的排列顺序进行直呼的操作
'接诊：当执行顺呼或者直呼后，如果病人到达诊室，则执行接诊
'
'对分诊台可开放直呼、顺呼、广播的功能
'对医生站可开放直呼、顺呼、接诊的功能
'
'
'取消原有的“重呼”功能（重呼：呼叫已经进行顺呼或者直呼过的队列数据）
'
'
'完成：将当前队列中的数据设置为完成状态
'弃号：放弃当前队列中的数据呼叫
'暂停：暂停当前队列中的数据呼叫
'恢复：恢复已经弃号、暂停、完成的队列数据
'回诊：重新进行排队，并确定回诊序号
'
Private mlngModule As Long

Private mblnCustomCfg As Boolean
Private mTQueueCols As TColsInfo
Private mTCallCols As TColsInfo

Private mcnOracle As ADODB.Connection
Private mstr队列名称() As String            '叫号系统中需要进行显示的数据队列名称
Private mstrCurrent队列名称 As String       '当前选中的队列名称
Private mIsUnload As Boolean                '是否退出
Private mstrBusinessIds As String           '保存界面上已经加载的业务ID

Private objVoice As Object                  '语音呼叫对象

Private mlng呼叫方式 As Integer             '呼叫方式 0 表示本地呼叫，1 表示远端呼叫
Private mint语音广播时间长度 As Integer     '语音播放的时间长度，默认值为15秒
Private mlng语音广播语速 As Long            '语音广播的语速（0-100）,用int的时候，有些机器会无法设置语速
Private mlng语音播放次数 As Long            '语音广播次数
Private mstr呼叫站点名称 As String          '执行呼叫的站点名称
Private mbln启用语音呼叫 As Boolean         '是否启用本地的语音呼叫功能
Private mlngCurPlayCount As Long            '当前已播放次数
Private mbln显示排队队列 As Boolean         '是否使用显示设备显示排队队列
Private mstrShowColumnInf As String         '显示列的配置信息
Private mstrShowCalledColumnInf As String   '呼叫显示列配置信息
Private mlng回诊病人优先 As Long            '回诊病人优先排队
Private mstr语音类型 As String              '语音类型
Private mlng轮询时间 As Long                '轮询时间长度
Private mlngQueueGroupType As Long          '排队分组类型
Private mlngOrderStyle As Long              '使用数据原始顺序排序
Private mblnIsSelectedCallingList As Boolean '是否选择已叫号队列
Private mlngQueueFocusRow As Long            '排队队列焦点行
Private mlngCallingFocusRow As Long          '呼叫队列焦点行
Private mstrLocateType As String             '定位条件
Private mblnIsLoad As String             '是否处于加载状态
Private mobjSquareCard As Object    '一卡通，卡结算部件
Private mstrPrivs As String                 '权限字符串

Private mstrCurrentWorkID As String           '当前选中的业务ID
Private mlngCurrentWorkType As Long         '当前业务类型
Private mlngCurrentQueueId As Double           '当前队列ID

Private mstrLoginUserName As String
Private mblnFuncState(7) As Boolean         '功能状态 0-恢复，1-直呼/顺呼，2-弃号 ，3-暂停，4-完成就诊，5,-广播， 6,回诊，7,接诊
Private mstr诊室条件 As String
Private mstr医生条件 As String
Private mstrExcludeData As String
Private mintViewDataType As Integer

Private mintDetonatEvent As Integer     '触发OnSelectedChange事件  0--初始值无作用，1--触发rptQueueList排队列表的事件   2--触发rptCallList呼叫列表的事件
Private mblnNotRefresh As Boolean          '为true时，排队列表或呼叫列表执行选择行变换事件后无需刷新
'界面布局
Private mlngQueueW1 As Long
Private mlngQueueW2 As Long
Private mlngLEDW As Long
Private mlngLEDH As Long

Private mintIconSize As Integer
Private mblnIsDisplayText As Boolean
Private mblnFirst As Boolean

Public mblnIsShowFindTools As Boolean   '是否显示查找工具栏

Private mlngMaxLen As Long '获取所有排队号码值中的最长度
Private mblnIsGroup As Boolean '排队叫号列表是否显示分组

Private Type TColsInfo
    lngColIndex_ID As Long   '字段名称
    lngColIndex_病人ID As Long   '字段名称
    lngColIndex_队列名称 As Long   '字段名称
    lngColIndex_业务ID As Long   '字段名称
    lngColIndex_科室ID As Long   '字段名称
    lngColIndex_排队号码 As Long   '字段名称
    lngColIndex_患者姓名 As Long   '字段名称
    lngColIndex_诊室 As Long   '字段名称
    lngColIndex_医生姓名 As Long   '字段名称
    lngColIndex_业务类型 As Long   '字段名称
End Type

Private Enum mCol
    队列名称 = 0: Id: 病人ID: 排队标记: 排队号码:  排队序号: 患者姓名: 优先: 回诊序号: 回诊排序号: 科室ID: 诊室: 医生姓名: 排队状态: 排队时间: 呼叫医生: 业务类型: 业务ID: 呼叫时间: 排序名称: ORD
End Enum

Public Event OnRefresh(str队列名称() As String, ByVal strCur队列名称 As String, ByVal strCur业务ID As String, ByVal strMustCols As String, _
    ByVal str诊室 As String, ByVal str医生 As String, ByVal strExcludeData As String, ByVal intViewDataType As Integer, ByVal str执行状态 As String, ByRef blnIsCustom As Boolean)
    
Public Event OnInitQueueList(ByRef objQueueList As Object, ByRef objCallList As Object, ByRef blnIsCustom As Boolean)
Public Event OnQueueRoomLoad(ByVal str业务ID As String, rsRoomData As ADODB.Recordset, rsDoctorData As ADODB.Recordset)
Public Event OnQueueExecuteBefore(ByVal str业务ID As String, ByVal byt操作类型 As Byte, blnCancel As Boolean, strNewQueueName As String)
Public Event OnQueueExecuteAfter(ByVal str业务ID As String, ByVal byt操作类型 As Byte)
Public Event OnRecevieDiagnose(ByVal str业务ID As String, ByVal lng业务类型 As Long)
Public Event OnSelectionChanged(ByVal blnIsCallingList As Boolean, objDataRow As XtremeReportControl.ReportRow, cbrMain As XtremeCommandBars.CommandBars)

'Public Sub zlShowMe(cnOracle As ADODB.Connection, str队列名称() As String, strCurrent队列名称 As String, lngCurrentWorkID As Long)
'    '队列的下标从1开始
'    Call zlRefresh(cnOracle, str队列名称, strCurrent队列名称, lngCurrentWorkID)
'
'    Me.Show
'End Sub

''''''''''''公共函数'''''''''''''''''''''''

Public Sub zlSetToolIcon(ByVal intIconSize As Integer, ByVal blnIsDisplayText As Boolean)
  mintIconSize = intIconSize
  mblnIsDisplayText = blnIsDisplayText
  
  Call Me.cbrMain.Options.SetIconSize(True, mintIconSize, mintIconSize)
  Call Me.cbrMain.RecalcLayout

'  Call SetCommandBarStyle
'  Call InitCommandBars
End Sub



Public Sub zlInitVar(cnOracle As ADODB.Connection, Optional lngSys As Long = 100, _
    Optional int业务类型 As Integer = 0, Optional intValidDays As Integer = 1, _
    Optional strPrivs As String = "", Optional strOption As String = "", Optional blnIsGroup As Boolean = True)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：初始化系统参数
    '入参：strOption-暂留,以后扩展
    '编制：刘兴洪
    '日期：2010-06-11 11:01:09
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    glngSys = lngSys
    glngModul = 1160
    mstrPrivs = strPrivs
    Set mcnOracle = cnOracle
    mblnIsGroup = blnIsGroup
    
    If Trim(mstrPrivs) = "" Then
        mstrPrivs = GetPrivFunc(glngSys, glngModul)
    End If
    
    mlngModule = Val(strOption)
    
    Call ClearQueueData(int业务类型, intValidDays)
End Sub


Private Sub ClearQueueData(ByVal int业务类型 As Integer, ByVal intValidDays As Integer)
    Dim strSql As String
    
    On Error GoTo errHandle

    strSql = "ZL_排队清除(" & int业务类型 & "," & intValidDays & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "清楚排队数据")
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Function GetQueueBusinessDataIDs() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取业务IDs
    '入参:bytType-0-挂号;1...
    '出参:
    '返回:成功返回业务IDs,多个用逗号分离,如:22,33,44
    '编制:刘兴洪
    '日期:2014-03-11 16:48:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    GetQueueBusinessDataIDs = mstrBusinessIds
 

End Function


Private Sub SwitchActiveWindow(ByVal blnIsCalledList As Boolean)
    On Error Resume Next
    
    If blnIsCalledList Then
        scCallInf.GradientColorDark = &HFF8080
        scCallInf.GradientColorLight = &HFFC0C0
        
        scQueueInf.GradientColorDark = &H808080
        scQueueInf.GradientColorLight = &HC0C0C0
    Else
        scQueueInf.GradientColorDark = &HFF8080
        scQueueInf.GradientColorLight = &HFFC0C0
        
        scCallInf.GradientColorDark = &H808080
        scCallInf.GradientColorLight = &HC0C0C0
    End If
End Sub



Private Sub SetReportRecordItem(rriItem As ReportRecord, rsData As ADODB.Recordset)
    Dim i As Integer
    
    On Error GoTo errHandle
    rriItem(mCol.Id).value = rsData("id")
    rriItem(mCol.病人ID).value = Nvl(rsData("病人ID"))
    
    rriItem(mCol.队列名称).Caption = rsData("部门名称") & ":" & IIf(InStr(1, Nvl(rsData("队列名称")), ":") <= 0, "", Mid(Nvl(rsData("队列名称")), InStr(1, Nvl(rsData("队列名称")), ":") + 1))
    rriItem(mCol.队列名称).value = Nvl(rsData("队列名称"))

    rriItem(mCol.患者姓名).value = Nvl(rsData("患者姓名"))
    rriItem(mCol.科室ID).value = Nvl(rsData("科室ID"))
    rriItem(mCol.排队标记).value = Nvl(rsData("排队标记"))
    rriItem(mCol.排队序号).value = Lpad(Nvl(rsData("排队序号")), 20)
    rriItem(mCol.排队号码).value = Lpad(Nvl(rsData("排队号码")), mlngMaxLen)
    rriItem(mCol.排队时间).value = Nvl(rsData("排队时间"))
    rriItem(mCol.呼叫时间).value = Nvl(rsData("呼叫时间"))
    rriItem(mCol.回诊序号).value = Nvl(rsData("回诊序号"))
    rriItem(mCol.回诊排序号).value = Nvl(rsData("回诊排序号"))
    rriItem(mCol.呼叫医生).value = Nvl(rsData("呼叫医生"))
    rriItem(mCol.排序名称).value = DeptNametransform(Nvl(rsData("部门名称")))
    rriItem(mCol.排序名称).Caption = (Nvl(rsData("部门名称")))
    rriItem(mCol.ORD).value = Format(rsData.AbsolutePosition, "00000000")
    
    If Nvl(rsData("回诊序号")) = "" Then
        rriItem(mCol.患者姓名).Icon = 807
    Else
        rriItem(mCol.患者姓名).Icon = 3504
    End If
    
    
    If Nvl(rsData("排队状态")) = 1 Then
        rriItem(mCol.排队状态).value = "呼叫中"
        For i = 0 To rptQueueList.Columns.Count - 1
            rriItem(i).BackColor = &HC0C0FF
        Next
    ElseIf Nvl(rsData("排队状态")) = 0 Then
        rriItem(mCol.排队状态).value = "排队中"
        For i = 0 To rptQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbWhite
        Next
    ElseIf Nvl(rsData("排队状态")) = 3 Then
        rriItem(mCol.排队状态).value = "暂停"
        For i = 0 To rptQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbYellow
        Next
    ElseIf Nvl(rsData("排队状态")) = 4 Then
        rriItem(mCol.排队状态).value = "完成"
        For i = 0 To rptQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbGreen
        Next
    ElseIf Nvl(rsData("排队状态")) = 7 Then
        rriItem(mCol.排队状态).value = "已呼叫"
'        For i = 0 To rptQueueList.Columns.Count - 1
'            rriItem(i).BackColor = &HFFC0C0
'        Next
    Else
        rriItem(mCol.排队状态).value = "已弃号"
        For i = 0 To rptQueueList.Columns.Count - 1
            rriItem(i).BackColor = &HC0C0C0
        Next
    End If
    
    If mlngQueueGroupType = 1 Then
        rriItem(mCol.医生姓名).value = Nvl(rsData("部门名称")) & ":" & Nvl(rsData("医生姓名"))
    Else
        rriItem(mCol.医生姓名).value = Nvl(rsData("医生姓名"))
    End If

    rriItem(mCol.业务类型).value = Nvl(rsData("业务类型"))
    rriItem(mCol.业务ID).value = Nvl(rsData("业务ID"))

    rriItem(mCol.优先).value = IIf(Nvl(rsData("优先")) = 1, "优先", "")
    
    If mlngQueueGroupType = 2 Then
        rriItem(mCol.诊室).value = Nvl(rsData("部门名称")) & ":" & Nvl(rsData("诊室"))
    Else
        rriItem(mCol.诊室).value = Nvl(rsData("诊室"))
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
 
 
 Private Sub SetFocusToCalledList()
    On Error Resume Next
    
    'If rptCallList.Visible Then rptCallList.SetFocus
    
    On Error GoTo 0
 End Sub
 
 
 Private Sub SetFocusToQueueList()
    On Error Resume Next
    
    'If rptQueueList.Visible Then rptQueueList.SetFocus
    
    On Error GoTo 0
 End Sub

Public Function zlRefresh(str队列名称() As String, ByVal strCur队列名称 As String, ByVal strCur业务ID As String, _
    Optional str诊室 As String = "", Optional str医生 As String = "", Optional strExcludeData As String = "", Optional intViewDataType As Integer = 0) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '功能：调用刷新指定医嘱id的报告内容，并根据情况提供编辑功能
    '入参：str队列名称():传入的指定队列数组(从1开始)
    '         strCur队列名称-当前队列名称
    '         lngCur业务ID-业务ID
    '         str诊室-限制为指定的诊室,可以为多个诊室:如"一诊室,二诊室,..."
    '         str医生-限制为制定的医生,可以传多个医生,用逗号分隔,如"张三,李四,..."
    '         strExcludeData-排除的指定业务ID
    '         intViewDataType数据显示类型，0显示当前科室下的所有数据，
    '                                      1显示诊室为当前诊室且医生姓名为空，或者医生姓名等于当前医生，或者诊室为空和医生为空的数据
    '                                      2显示诊室为当前诊室和医生姓名为空或医生姓名等于当前医生的数据
    '                                      3显示当前医生的数据
    '编制：刘兴洪
    '日期：2010-06-11 20:54:55
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsLocal As ADODB.Recordset
    Dim rptRecord As ReportRecord
    Dim rptCalling As ReportRecord
    Dim strSql As String, j As Integer, i As Integer, str执行状态 As String
    Dim strQueueId As String
    Dim strValues(0 To 10) As String, strValue As String, strUninTable As String
    Dim strFilter As String
    Dim blnIsCustom As Boolean
    Dim strMustCols As String

    err = 0: On Error GoTo errHandle
    If mblnNotRefresh Then Exit Function  '执行事件OnSelectionChanged后无需刷新
    
    mstr队列名称 = str队列名称
    mstrExcludeData = strExcludeData
    mstr诊室条件 = str诊室
    mstr医生条件 = str医生
    mintViewDataType = intViewDataType
    
    strMustCols = "ID;病人ID;队列名称;业务ID;科室ID;排队号码;患者姓名;诊室;医生姓名;"
    
    str执行状态 = ""
    If chkOutQueue(0).value = vbUnchecked Then str执行状态 = str执行状态 & ",2"
    If chkOutQueue(1).value = vbUnchecked Then str执行状态 = str执行状态 & ",3"
    If chkOutQueue(2).value = vbUnchecked Then str执行状态 = str执行状态 & ",4"
    If rptQueueList.SelectedRows.Count > 0 Then mlngQueueFocusRow = rptQueueList.SelectedRows(0).Index
    If rptCallList.SelectedRows.Count > 0 Then mlngCallingFocusRow = rptCallList.SelectedRows(0).Index
        
    RaiseEvent OnRefresh(str队列名称(), strCur队列名称, strCur业务ID, strMustCols, str诊室, str医生, strExcludeData, intViewDataType, str执行状态, blnIsCustom)
    If blnIsCustom Then
'        自定义流程需要在之前的Onfresh事件中处理好查询需要的数据并且显示。
        
        '兼容以前的功能，获取 mstrBusinessIds
        mstrBusinessIds = ""
        For i = 0 To rptQueueList.Records.Count - 1
            If rptQueueList.Rows(i).GroupRow <> True Then
                If rptQueueList.Rows(i).Record.Item(2).value <> "" Then
                    If mstrBusinessIds <> "" Then mstrBusinessIds = mstrBusinessIds & ";"
                    mstrBusinessIds = mstrBusinessIds & rptQueueList.Rows(i).Record.Item(2).value
                End If
            End If
        Next
        
        For i = 0 To rptCallList.Records.Count - 1
            If rptCallList.Rows(i).GroupRow <> True Then
                If rptCallList.Rows(i).Record.Item(2).value <> "" Then
                    If mstrBusinessIds <> "" Then mstrBusinessIds = mstrBusinessIds & ";"
                    mstrBusinessIds = mstrBusinessIds & rptCallList.Rows(i).Record.Item(2).value
                End If
            End If
        Next
        
    Else
        '非自定义流程，保持113794前的处理方式
        
        strFilter = "": strValue = "": j = 0: strUninTable = ""
        If SafeArrayGetDim(mstr队列名称) > 0 Then
            For i = 1 To UBound(mstr队列名称)
                If Trim(mstr队列名称(i)) <> "" Then
                    If j > 10 Then
                        strFilter = strFilter & " Or A.队列名称 ='" & str队列名称(i) & "'"
                    Else
                        If zlCommFun.ActualLen(strValue) > 2000 Then
                             strValues(j) = Mid(strValue, 2)
                             strUninTable = strUninTable & " Union ALL  Select  Column_Value as 队列名称 From Table(Cast(f_Str2list([" & j + 4 & "]) As zlTools.t_Strlist))  " & vbCrLf
                             strValue = "": j = j + 1
                        End If
                        strValue = strValue & "," & str队列名称(i)
                    End If
                End If
            Next i
            If strValue <> "" Then
                strValues(j) = Mid(strValue, 2)
                strUninTable = strUninTable & " Union ALL  Select  Column_Value as 队列名称 From Table(Cast(f_Str2list([" & j + 4 & "]) As zlTools.t_Strlist))  " & vbCrLf
            End If
        End If
        
        If strUninTable <> "" Then
            strUninTable = Mid(strUninTable, 11)
        Else
            labError.Caption = "没有可显示的叫号队列信息，请检查相关排队科室设置"
            Exit Function
        End If
        
        If strFilter <> "" Then strFilter = "( " & Mid(strFilter, 4) & ")"
        
        str执行状态 = ""
        If chkOutQueue(0).value = vbUnchecked Then str执行状态 = str执行状态 & ",2"
        If chkOutQueue(1).value = vbUnchecked Then str执行状态 = str执行状态 & ",3"
        If chkOutQueue(2).value = vbUnchecked Then str执行状态 = str执行状态 & ",4"
         
        '为了支持复制，需要将number类型的字段进行转换，可以使用to_Number方式
        strSql = "" & _
        "   Select /*+ Rule*/  to_Number(A.ID) as ID, to_Number(a.病人id) as 病人id, A.队列名称, A.排队序号, to_Number(A.业务类型) as 业务类型, to_Number(A.业务ID) as 业务ID," & _
        "           to_Number(科室ID) as 科室ID, x.名称 as 部门名称, 排队号码 , 排队标记,患者姓名,诊室,医生姓名," & _
        "            (select 姓名 from 人员表 J, 上机人员表 K where J.ID=K.人员ID and K.用户名=A.呼叫医生) as 呼叫医生, " & _
        "           to_Number(优先) as 优先, to_Number(回诊序号) as 回诊序号, To_Char(排队时间, 'yyyy-mm-dd hh24:mi:ss') as 排队时间, To_Char(呼叫时间, 'yyyy-mm-dd hh24:mi:ss') as 呼叫时间,to_Number(排队状态) as 排队状态, " & _
                    IIf(mlng回诊病人优先 = 1, "to_number(nvl(回诊序号, 9999999999)) as 回诊排序号", "0 as 回诊排序号") & _
        "   From 排队叫号队列 a, 部门表 x " & IIf(strUninTable <> "", ", (" & strUninTable & ") b ", "") & _
                    IIf(mintViewDataType = 1, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                    IIf(mintViewDataType = 2, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                    IIf(mintViewDataType = 3, ", Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D", "") & _
        "   Where   (nvl(是否分时点, 0)=0 and A.排队时间 <= trunc(sysdate + 1) - 1/24/60/60 or nvl(是否分时点, 0)=1 and sysdate > a.排队时间) " & IIf(strUninTable <> "", " and a.队列名称=b.队列名称 ", "") & " and instr([3],A.排队状态)=0  and x.ID=a.科室ID  " & _
                    IIf(mintViewDataType = 1, " and  ((a.诊室=C.Column_Value and a.医生姓名 is null) or a.医生姓名=D.Column_Value or (a.诊室 is null and a.医生姓名 is null))", "") & _
                    IIf(mintViewDataType = 2, " and (a.诊室=C.Column_Value and (a.医生姓名 is Null or a.医生姓名=D.Column_Value)) ", "") & _
                    IIf(mintViewDataType = 3, " and a.医生姓名=D.Column_Value", "") & _
        "           " & strFilter & _
        "   Order by  排队状态 desc, 排队序号,优先 Desc, 回诊排序号, 排队时间, 排队号码 "
        
    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询队列", mstr诊室条件, mstr医生条件, str执行状态, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
        Set rsLocal = zlDatabase.CopyNewRec(rsTemp)
        
        '删除需要排除的数据,并获取实际排队号码值得最长度
        If rsLocal.RecordCount > 0 Then rsLocal.MoveFirst
        While Not rsLocal.EOF
            If InStr(1, strExcludeData, rsLocal!业务类型 & ":" & rsLocal!业务ID) > 0 Then
                rsLocal.Delete
            End If
            
            If LenB(StrConv(Trim(Nvl(rsLocal("排队号码"))), vbFromUnicode)) > mlngMaxLen Then
                mlngMaxLen = LenB(StrConv(Trim(Nvl(rsLocal("排队号码"))), vbFromUnicode))
            End If
            
            rsLocal.MoveNext
        Wend
    
        rsLocal.Sort = "队列名称, 排队状态 desc, 排队序号, 优先 Desc, 回诊排序号, 排队时间, 排队号码"
        If rsLocal.RecordCount > 0 Then rsLocal.MoveFirst
        
        Call rptQueueList.Records.DeleteAll
        Call rptCallList.Records.DeleteAll
            
        While Not rsLocal.EOF
    
            If rsLocal("排队状态") = 7 Or rsLocal("排队状态") = 1 Then
                Set rptCalling = rptCallList.Records.Add
                For j = 0 To Me.rptCallList.Columns.Count - 1
                    rptCalling.AddItem ""
                Next
                
                Call SetReportRecordItem(rptCalling, rsLocal)
            Else
                Set rptRecord = rptQueueList.Records.Add
                For j = 0 To Me.rptQueueList.Columns.Count - 1
                    rptRecord.AddItem ""
                Next
                
                Call SetReportRecordItem(rptRecord, rsLocal)
            End If
            
            If mstrBusinessIds <> "" Then mstrBusinessIds = mstrBusinessIds & ","
            mstrBusinessIds = mstrBusinessIds & Nvl(rsLocal!业务ID)
            
            rsLocal.MoveNext
        Wend
        
        rptQueueList.Populate
        rptCallList.Populate
    
    End If
    
    On Error GoTo errShow
    
    '恢复选择的排队数据
    If mlngQueueFocusRow >= rptQueueList.Rows.Count Then
        mlngQueueFocusRow = IIf(rptQueueList.Rows.Count <= 0, -1, rptQueueList.Rows.Count - 1)
    End If
    
    If mlngQueueFocusRow > -1 Then
        rptQueueList.Rows(mlngQueueFocusRow).Selected = True
    End If
    
    '恢复选择的呼叫数据
    If mlngCallingFocusRow >= rptCallList.Rows.Count Then
        mlngCallingFocusRow = IIf(rptCallList.Rows.Count <= 0, -1, rptCallList.Rows.Count - 1)
    End If
        
    If mlngCallingFocusRow > -1 Then
        rptCallList.Rows(mlngCallingFocusRow).Selected = True
    End If
        
    '恢复焦点列表
    Call SwitchActiveWindow(mblnIsSelectedCallingList)

errShow:
    
    '显示排队队列
    Call ShowQueue
    
    zlRefresh = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    
    Call SaveErrLog
    
    zlRefresh = err.Number
End Function


Public Sub zlCommandBarSet(ByVal intFuncType As Integer, ByVal blnUseState As Boolean)
'************************************************************************************
'
'设置功能状态
'
'intFuncType：功能类型 0-恢复，1-直呼/顺呼，2-弃号 ，3-暂停，4-完成就诊，5,-广播, 6,-回诊, 7-接诊
'blnUseState：是否启用
'
'************************************************************************************
    If (intFuncType >= 0) And (intFuncType <= 7) Then
        mblnFuncState(intFuncType) = blnUseState
    End If
End Sub


Public Function zlQueueExec(str当前队列名 As String, lng业务类型 As Long, str业务ID As String, byt操作类型 As Byte) As Boolean
'*************************************************************************************
'
'执行叫号相关操作
'
'str当前队列名：需要操作的队列名称,当不能确定诊室或者医生的时候，使用科室ID作为队列名称
'
'lng业务ID：表示当前业务的ID数据
'
'byt操作类型：叫号操作的类型 0-恢复，1-直呼/顺呼（Lng业务ID=0为顺呼），2-弃号 ，3-暂停，4-完成就诊，5,-广播 6,-回诊
'
'*************************************************************************************

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngQueueId As Long
    Dim blnFind As Boolean
    Dim i As Integer
            
    On Error GoTo errHandle
    
    zlQueueExec = False
    mstrCurrent队列名称 = str当前队列名
        
    Select Case byt操作类型
        Case 0, 1, 2, 3, 4, 6
            strSql = "ZL_排队叫号队列_呼叫('" & str当前队列名 & "'," & byt操作类型 & ",'" & mstrLoginUserName & "','" & mstr呼叫站点名称 & "'," & lng业务类型 & ",'" & str业务ID & "')"
            zlDatabase.ExecuteProcedure strSql, "排队叫号"
        Case 5

            
            strSql = "select ID from 排队叫号队列 where 队列名称=[1] and 业务ID=[2] and 业务类型=[3]"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "广播", str当前队列名, str业务ID, lng业务类型)
            
            While Not rsTemp.EOF
                lngQueueId = rsTemp!Id
        
                strSql = "ZL_排队语音呼叫_INSERT(" & lngQueueId & ",'" & mstr呼叫站点名称 & "', 1)"
                Call zlDatabase.ExecuteProcedure(strSql, "广播")
                
                rsTemp.MoveNext
            Wend
    End Select

        
    '如果列表中存在该数据，则定位该数据
    blnFind = False
    For i = 0 To rptQueueList.Rows.Count - 1
        If rptQueueList.Rows(i).GroupRow <> True Then
            If rptQueueList.Rows(i).Record.Item(mTQueueCols.lngColIndex_业务ID).value = str业务ID _
                And rptQueueList.Rows(i).Record.Item(mTQueueCols.lngColIndex_业务类型).value = lng业务类型 _
                And rptQueueList.Rows(i).Record.Item(mTQueueCols.lngColIndex_队列名称).value = mstrCurrent队列名称 Then
            
                rptQueueList.Rows(i).Selected = True
                blnFind = True
            
                Exit For
            End If
        End If
    Next i
    
    '从已呼叫列表中查找数据
    If Not blnFind Then
        Call SetFocusToCalledList
        
        For i = 0 To rptCallList.Rows.Count - 1
            If rptCallList.Rows(i).GroupRow <> True Then
                If rptCallList.Rows(i).Record.Item(mTCallCols.lngColIndex_业务ID).value = str业务ID _
                    And rptCallList.Rows(i).Record.Item(mTCallCols.lngColIndex_业务类型).value = lng业务类型 _
                    And rptCallList.Rows(i).Record.Item(mTCallCols.lngColIndex_队列名称).value = mstrCurrent队列名称 Then
                
                    rptCallList.Rows(i).Selected = True
                
                    Exit For
                End If
            End If
        Next i
    End If
    
    
    zlQueueExec = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function CheckQueueDataIsHas(ByVal lngQueueId As Long) As Boolean
'***********************************************************************
'检测队列数据是否存在
'
'参数：
'lngQueueId：需要进行检查的队列ID
'***********************************************************************

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '判断队列ID是否已经存在
    strSql = "select /*+ RULE*/ ID from 排队叫号队列 where Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询叫号数据是否存在", mlngCurrentQueueId)
    
    CheckQueueDataIsHas = Not rsTemp.EOF
    
    Exit Function
errHandle:
    CheckQueueDataIsHas = False
    If ErrCenter = 1 Then Resume
End Function


Private Function CheckIsSelectedData() As Boolean
    On Error Resume Next
    
    '取排队列表数据
    If mblnIsSelectedCallingList = False Then
        If rptQueueList.SelectedRows.Count = 0 Then
            If rptQueueList.Rows.Count > 0 Then
                rptQueueList.Rows(1).Selected = True
                
                 Call rptQueueList_SelectionChanged
            Else
                CheckIsSelectedData = False
                Exit Function
            End If
        Else
            '选中的行不能是分组行,如果是分组行，需要设置到该分组下的第一行
            If rptQueueList.SelectedRows(0).GroupRow = True Then
                If rptQueueList.SelectedRows(0).Childs.Count > 0 Then
                    rptQueueList.SelectedRows(0).Childs(0).Selected = True
                    
                    Call rptQueueList_SelectionChanged
                Else
                    CheckIsSelectedData = False
                    Exit Function
                End If
            Else
                Call rptQueueList_SelectionChanged
            End If
        End If
    Else
    '取已呼叫列表数据
        If rptCallList.SelectedRows.Count = 0 Then
            If rptCallList.Rows.Count > 0 Then
                rptCallList.Rows(1).Selected = True
                
                Call rptQueueList_SelectionChanged
            Else
                CheckIsSelectedData = False
                Exit Function
            End If
        Else
            '选中的行不能是分组行,如果是分组行，需要设置到该分组下的第一行
            If rptCallList.SelectedRows(0).GroupRow = True Then
                If rptCallList.SelectedRows(0).Childs.Count > 0 Then
                    rptCallList.SelectedRows(0).Childs(0).Selected = True
                    
                    Call rptCallList_SelectionChanged
                Else
                    CheckIsSelectedData = False
                    Exit Function
                End If
            Else
                Call rptCallList_SelectionChanged
            End If
        End If
    End If
    
    CheckIsSelectedData = True
End Function

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnIsHasData As Boolean
    Dim i As Integer
    
    '执行工具栏命令
    On Error GoTo errHand
    
    labError.Caption = ""
    
    Select Case Control.Id
        Case conMenu_Queue_CallThis, conMenu_Queue_CallNext, _
             conMenu_Queue_CallFirst, conMenu_Queue_Restore, _
             conMenu_Queue_ReCall, conMenu_Queue_Abandon, _
             conMenu_Queue_Refresh, conMenu_Queue_Setup, _
             conMenu_Queue_Update, conMenu_Queue_Broadcast, _
             conMenu_Queue_Pause, conMenu_Queue_Finaled, _
             conMenu_Queue_Find, conMenu_Queue_ComeBack, _
             conMenu_Queue_RecDiagnose
             
        Case Else
            Exit Sub
    End Select
    
    
    '如果为顺乎操作，则呼叫时，会直接访问rptQueueList列表(顺呼，刷新，和设置都不需要设置列表)
    If Control.Id <> conMenu_Queue_CallNext _
        And Control.Id <> conMenu_Queue_Refresh _
        And Control.Id <> conMenu_Queue_Setup _
        And Control.Id <> conMenu_Queue_Find Then
    
        If Not CheckIsSelectedData Then
            MsgBox "没有选择需要执行的数据，不能执行该操作。", vbInformation, "排队叫号系统"
            Exit Sub
        End If
        
        If Not CheckQueueDataIsHas(mlngCurrentQueueId) Then
            MsgBox "数据不存在或已被执行，请运行刷新操作。", vbInformation, "排队叫号系统"
            Exit Sub
        End If
    End If
        
    
    Select Case Control.Id
        Case conMenu_Queue_CallThis '直呼
            '---
            Call comMenu_直呼
            
        Case conMenu_Queue_RecDiagnose '接诊
            '---
            Call comMenu_接诊
            
        Case conMenu_Queue_Broadcast '广播  如果执行“广播”，则不需要对数据进行刷新操作
            '---
            Call comMenu_广播
            
            Exit Sub
        Case conMenu_Queue_CallFirst    '优先
            '---
            Call comMenu_优先
        
        Case conMenu_Queue_Restore    '恢复
            '---
            Call comMenu_恢复
            
        Case conMenu_Queue_Abandon      '弃号
            '---
            Call comMenu_弃号
            
        Case conMenu_Queue_Pause       '暂停
            '---
            Call comMenu_暂停
            
        Case conMenu_Queue_Finaled      '完成
            '---
            Call comMenu_完成
                        
'        Case conMenu_Queue_Refresh      '刷新 该处不需要进行刷新，在执行任何操作后，会在该过程的最后进行刷新
'            Call comMenu_刷新

        Case conMenu_Queue_Find     '查找
            Call comMenu_查找
            
        Case conMenu_Queue_CallNext '顺呼
            Call comMenu_顺呼
        
        Case conMenu_Queue_Update       '修改
            Call comMenu_修改
            
        Case conMenu_Queue_Setup        '设置  如果是“设置”操作，则不需要对数据进行刷新
            Call comMenu_设置
            
            Exit Sub
    End Select
    
    Call zlRefresh(mstr队列名称, mstrCurrent队列名称, mstrCurrentWorkID, mstr诊室条件, mstr医生条件, mstrExcludeData, mintViewDataType)
    
    
    '当执行顺呼或者直呼之后，需要将焦点设置到呼叫列表
    If Control.Id = conMenu_Queue_CallThis Or Control.Id = conMenu_Queue_CallNext Then
        For i = 0 To rptCallList.Rows.Count - 1
            If rptCallList.Rows(i).GroupRow <> True Then
                If rptCallList.Rows(i).Record.Item(mTCallCols.lngColIndex_业务ID).value = mstrCurrentWorkID And rptCallList.Rows(i).Record.Item(mTCallCols.lngColIndex_业务类型).value = mlngCurrentWorkType Then
                    rptCallList.Rows(i).Selected = True
                    
                    Call rptCallList_SelectionChanged
                    Call SetFocusToCalledList
                    
                    mblnIsSelectedCallingList = True
                    
                    Call SwitchActiveWindow(mblnIsSelectedCallingList)

                    Exit For
                End If
            End If
        Next i
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub zlDefCommandBars(ByVal cbsThis As Object)
   '创建工具栏按钮
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    On Error GoTo errHandle
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "排队(&C)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.Id = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        If InStr(mstrPrivs, "直呼") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallThis, "直呼"): cbrControl.IconId = 732
        End If
        
        If InStr(mstrPrivs, "顺呼") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "顺呼"): cbrControl.IconId = 744: cbrControl.ToolTipText = "按顺序呼叫下一个"
        End If
        
'        If InStr(mstrPrivs, "重呼") > 0 Then
'            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_ReCall, "重呼"): cbrControl.IconId = 3014: cbrControl.ToolTipText = "再次呼叫"
'        End If
        
        If InStr(mstrPrivs, "广播") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "重呼"): cbrControl.IconId = 745
        End If
    End With

    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = cbsThis(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.Id, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.Id, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    
    With cbrToolBar.Controls
        If InStr(mstrPrivs, "直呼") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallThis, "直呼", cbrControl.Index + 1): cbrControl.IconId = 732: cbrControl.ToolTipText = "直接呼叫当前患者"
        
            cbrControl.BeginGroup = True
        End If
        
        If InStr(mstrPrivs, "顺呼") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "顺呼", cbrControl.Index + 1): cbrControl.IconId = 744: cbrControl.ToolTipText = "按顺序呼叫下一个"
        End If
        
'        If InStr(mstrPrivs, "重呼") > 0 Then
'            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_ReCall, "重呼", cbrControl.Index + 1): cbrControl.IconId = 3014: cbrControl.ToolTipText = "再次呼叫"
'
'            cbrControl.BeginGroup = True
'        End If
        
        If InStr(mstrPrivs, "广播") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "重呼", cbrControl.Index + 1): cbrControl.IconId = 745
        End If
    End With

    '命令的快键绑定
    '-----------------------------------------------------
    With cbsThis.KeyBindings

    End With

    '设置不常用命令
    '-----------------------------------------------------
    With cbsThis.Options

    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
'''''''''''''''''''''''''''''''''''''''''''''

Private Sub SetCommandBarStyle()
    On Error Resume Next
    
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
        .SetIconSize True, mintIconSize, mintIconSize
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    'Me.cbrMain.ActiveMenuBar.Visible = False
End Sub

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrMenuBar As CommandBarPopup
    
    On Error GoTo errHandle
    '-----------------------------------------------------
    
    '排队呼叫工具栏定义
    Call cbrMain.DeleteAll
    Set cbrToolBar = Me.cbrMain.Add("呼叫工具栏", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    
        
    Call cbrToolBar.Controls.DeleteAll
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Refresh, "刷新"): cbrControl.IconId = 791
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "重呼"): cbrControl.IconId = 745
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallThis, "直呼"): cbrControl.IconId = 732
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "顺呼"): cbrControl.IconId = 744: cbrControl.ToolTipText = "按顺序呼叫下一个"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
       
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RecDiagnose, "接诊"): cbrControl.IconId = 8264: cbrControl.ToolTipText = "对报到病人进行接诊处理"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallFirst, "优先"): cbrControl.IconId = 216: cbrControl.ToolTipText = "设置为优先呼叫"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Restore, "恢复"): cbrControl.IconId = 252: cbrControl.ToolTipText = "将数据恢复到排队状态"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Pause, "暂停"): cbrControl.IconId = 746
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Abandon, "弃呼"): cbrControl.IconId = 8113: cbrControl.ToolTipText = "放弃呼叫"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Finaled, "完成"): cbrControl.IconId = 747
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Find, "查找"): cbrControl.IconId = 721
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
                       
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Update, "修改"): cbrControl.IconId = 3003: cbrControl.ToolTipText = "修改排队信息"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Setup, "设置"): cbrControl.IconId = 181: cbrControl.ToolTipText = "参数设置"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")

        Set cbrCustom = .Add(xtpControlCustom, conMenu_Queue_LocateNew, "定位")
            cbrCustom.Handle = Pati.hwnd
            cbrCustom.Flags = xtpFlagRightAlign
            cbrCustom.Style = xtpButtonIconAndCaption
            cbrCustom.Category = "CallFind"
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If (cbrControl.Type = xtpControlButton) Or (cbrControl.Type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
        If cbrControl.Category = "" Then cbrControl.Category = "Main" '设置成主界面菜单
    Next
    
    cbrToolBar.Position = xtpBarTop
    
    
    

    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitFaceScheme()
    '初始界面布局
    
    Dim Pane1 As Pane, Pane2 As Pane, pane3 As Pane
    
    On Error GoTo errHandle
    
    With Me.DkpMain
        .CloseAll
        .SetCommandBars cbrMain
        
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = DkpMain.CreatePane(0, IIf(mlngQueueW1 < 1000, 1000, mlngQueueW1), _
                Me.Height, _
                DockLeftOf, Nothing)
                
    Pane1.Title = "排队列表"
    Pane1.Tag = 0
    Pane1.Handle = picQueueFace.hwnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    
    Set Pane2 = DkpMain.CreatePane(1, IIf(mlngQueueW2 < 1000, 1000, mlngQueueW2), _
                Me.Height, _
                DockRightOf, Nothing)
         
    
    Pane2.Title = "呼叫列表"
    Pane2.Tag = 1
    Pane2.Handle = picCallFace.hwnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    
    
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlUpdateCommandBars(Control)
End Sub

Private Sub chkOutQueue_Click(Index As Integer)
    Call comMenu_刷新
End Sub

Private Sub Form_Activate()
    '显示排队队列
    If mblnFirst Then
    Call ShowQueue
        mblnFirst = False
    End If
End Sub


Private Sub Form_Load()
On Error GoTo errh
    
    mblnIsLoad = True
    '当前登陆的用户名
    mstrLoginUserName = GetUserName()
    mblnFirst = True
    mintIconSize = 24
    mblnIsDisplayText = True
    mIsUnload = False
    mblnIsSelectedCallingList = False
    mlngQueueFocusRow = -1
    mlngCallingFocusRow = -1
    
    mintDetonatEvent = 0
    mblnNotRefresh = False
    
    
    Set objVoice = Nothing
    
    mblnCustomCfg = False

    Call InitLocalParas(False)
    Call SetCommandBarStyle
    Call InitCommandBars
    Call InitFaceScheme
    Call InitQueueList
    Call InitPati
    
    mblnIsLoad = False
    '调试控件位置
    Call picLabel_Resize
    
    Exit Sub
errh:
    err.Raise -1, , "初始化排队叫号窗体失败" & err.Description, vbInformation, "排队叫号系统"
End Sub

Private Function GetUserName() As String
'************************************************
'
'取得当前登陆的用户名
'
'************************************************
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
        Set rsTmp = zlDatabase.GetUserInfo
        
        If Not rsTmp.EOF Then
            GetUserName = IIf(IsNull(rsTmp!用户名), "", rsTmp!用户名)
        End If
        
    Exit Function
errHandle:
    GetUserName = ""
    If ErrCenter = 1 Then Resume
End Function


Private Sub InitLocalParas(blnIsLISForm As Boolean)
    Dim strReg As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    strReg = "公共全局\排队叫号"
        
    
    mint语音广播时间长度 = Val(zlDatabase.GetPara("语音广播时间长度", glngSys, glngModul, "15"))
    mlng语音广播语速 = Val(zlDatabase.GetPara("语音广播语速", glngSys, glngModul, "60"))
    mlng语音播放次数 = Val(zlDatabase.GetPara("语音播放次数", glngSys, glngModul, "1"))
    mbln启用语音呼叫 = zlDatabase.GetPara("启用语音呼叫", glngSys, glngModul, "1")
    mstrShowColumnInf = zlDatabase.GetPara("数据显示列", glngSys, glngModul, "号码,患者姓名,排队状态")
    mstrShowColumnInf = Replace(mstrShowColumnInf, "，", ",")
    mstrShowColumnInf = "," & mstrShowColumnInf & ","
    
    mstrShowCalledColumnInf = zlDatabase.GetPara("呼叫数据显示列", glngSys, glngModul, "号码,患者姓名")
    mstrShowCalledColumnInf = Replace(mstrShowCalledColumnInf, "，", ",")
    mstrShowCalledColumnInf = "," & mstrShowCalledColumnInf & ","
               
    
    mlng呼叫方式 = Val(zlDatabase.GetPara("叫号方式", glngSys, glngModul, "0"))
    
    If mlng呼叫方式 Then
        mstr呼叫站点名称 = zlDatabase.GetPara("远端呼叫站点", glngSys, glngModul, "")
        
        '如果为空就表示为本地站点
        If Trim(mstr呼叫站点名称) = "" Then
          mstr呼叫站点名称 = AnalyseComputer
        End If
    Else
        mstr呼叫站点名称 = AnalyseComputer
    End If
    
    
    mbln显示排队队列 = zlDatabase.GetPara("显示排队队列", glngSys, glngModul, "1")
    plngLEDModal = zlDatabase.GetPara("显示设备类别", glngSys, glngModul, "101")
    
    mstrLocateType = GetSetting("ZLSOFT", strReg, "定位方式", "姓名")
    
    mlng回诊病人优先 = zlDatabase.GetPara("回诊病人是否优先", glngSys, glngModul, "1")
    mlngQueueGroupType = zlDatabase.GetPara("排队分组类型", glngSys, glngModul, "0")
    mlngOrderStyle = zlDatabase.GetPara("使用数据原始顺序排序", glngSys, glngModul, "0")
    
    mstr语音类型 = zlDatabase.GetPara("语音类型", glngSys, glngModul, "系统默认")
    mlng轮询时间 = Val(zlDatabase.GetPara("轮询时间", glngSys, glngModul, "30"))
    
    If Not blnIsLISForm Then
        mlngQueueW1 = GetSetting("ZLSOFT", strReg, "队列显示宽度", Round(Me.Width / 3 * 2))
        mlngQueueW2 = GetSetting("ZLSOFT", strReg, "呼叫队列显示宽度", Round(Me.Width / 3))
        
        tmrBroadCast.Enabled = False
        tmrBroadCast.Interval = mlng轮询时间 * 1000
        tmrBroadCast.Enabled = True
    End If
    
    For i = 0 To 7
        mblnFuncState(i) = True
    Next i
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Public Sub InitLED(lngLEDModal As Long)
    If Not CreateObject_LED(lngLEDModal) Then Exit Sub
End Sub


Private Function CreateObject_LED(lngLEDModal As Long) As Boolean
    '创建LED显示对象
    
    Dim strSql As String
    Dim strObject As String

    On Error GoTo errHand
    
    '读取该LED显示接口的注册信息
    If prsLEDComponent.State = 0 Then
        strSql = "Select 部件类型,部件名,Nvl(启用,0) AS 启用 From 排队LED显示部件  "
        Set prsLEDComponent = zlDatabase.OpenSQLRecord(strSql, "提取该LED显示接口的注册信息")
    End If
    prsLEDComponent.Filter = "部件类型=" & lngLEDModal
    If prsLEDComponent.RecordCount = 0 Then
        prsLEDComponent.Filter = 0
        MsgBox "该LED接口还未注册！ 序号=" & lngLEDModal, vbInformation, "排队叫号系统"
        Exit Function
    End If
    strObject = UCase(prsLEDComponent!部件名)
    prsLEDComponent.Filter = 0
    
    '检查该对象是否存在
    On Error Resume Next
    If Not pobjLEDShow Is Nothing Then
        CreateObject_LED = True
        Exit Function
    End If
    
    '去掉文件名后缀
    strObject = Mid(strObject, 1, Len(strObject) - 4)
    '创建对象
    Set pobjLEDShow = CreateObject(strObject & ".Cls" & Mid(strObject, 4))
    
    
    '调用初始化函数
    CreateObject_LED = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Resize()
    Call picLabel_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strReg As String
    Dim str排队列宽 As String
    Dim str呼叫列宽 As String
    Dim i As Integer
    
    On Error GoTo err
    strReg = "公共全局\排队叫号"
    str排队列宽 = ""
    str呼叫列宽 = ""
        
    If Not mblnCustomCfg Then
        With Me.rptQueueList
            For i = 0 To 18
                str排队列宽 = IIf(str排队列宽 = "", .Columns.Column(i).Width, str排队列宽 & "," & .Columns.Column(i).Width)
            Next
        End With
        With Me.rptCallList
            For i = 0 To 18
                str呼叫列宽 = IIf(str呼叫列宽 = "", .Columns.Column(i).Width, str呼叫列宽 & "," & .Columns.Column(i).Width)
            Next
        End With
        SaveSetting "ZLSOFT", strReg, "排队列宽度配置", str排队列宽
        SaveSetting "ZLSOFT", strReg, "呼叫列宽度配置", str呼叫列宽
    Else
        With Me.rptQueueList
            For i = 0 To 18
                str排队列宽 = IIf(str排队列宽 = "", .Columns.Column(i).Width, str排队列宽 & "," & .Columns.Column(i).Width)
            Next
        End With
        With Me.rptCallList
            For i = 0 To 18
                str呼叫列宽 = IIf(str呼叫列宽 = "", .Columns.Column(i).Width, str呼叫列宽 & "," & .Columns.Column(i).Width)
            Next
        End With
        If mlngModule > 0 Then
            SaveSetting "ZLSOFT", "公共全局\自定义排队叫号" & CStr(mlngModule), "排队列宽度配置", str排队列宽
            SaveSetting "ZLSOFT", "公共全局\自定义排队叫号" & CStr(mlngModule), "呼叫列宽度配置", str呼叫列宽
        Else
            SaveSetting "ZLSOFT", "公共全局\自定义排队叫号", "排队列宽度配置", str排队列宽
            SaveSetting "ZLSOFT", "公共全局\自定义排队叫号", "呼叫列宽度配置", str呼叫列宽
        End If
    End If
    
    SaveSetting "ZLSOFT", strReg, "队列显示宽度", rptQueueList.Width
    SaveSetting "ZLSOFT", strReg, "呼叫队列显示宽度", rptCallList.Width
    SaveSetting "ZLSOFT", strReg, "定位方式", mstrLocateType
    
    Set mobjSquareCard = Nothing
    
    Set objVoice = Nothing
    '卸载优先原因窗体
    Unload frmPriorityCause
    
    '关闭LCD窗口
    If Not pobjLEDShow Is Nothing Then
        Call pobjLEDShow.zlClose
        Set pobjLEDShow = Nothing
    End If
    
    mIsUnload = True
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
    mIsUnload = True
End Sub

Private Sub Pati_GotFocus()
On Error Resume Next
    Pati.SelStart = 0
    Pati.SelLength = Len(Pati.Text)
End Sub

Private Sub Pati_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnIsLoad Then Exit Sub
    mstrLocateType = objCard.名称
End Sub

Private Sub Pati_KeyPress(KeyAscii As Integer)
    On Error GoTo errh
    
    Dim blnCard As Boolean
    
    '后续可进行只允许输入数字的控制
'    If Trim(Pati.GetCurCard.名称) = "住院号" Then
'        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
'            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
'        End If
'    End If

    If KeyAscii = 13 Then
        Call LocateQueueData(Pati.GetCurCard.名称, Pati.Text)
        Exit Sub
    End If
    
    If Pati.GetCurCard.是否刷卡 Then
        blnCard = Pati.zlIsBrushCard(Pati.objTxtInput, KeyAscii)
            
        If blnCard And Len(Pati.Text) = Pati.GetCardNoLen - 1 And KeyAscii <> 8 Then  '刷卡完毕处理
            Pati.Text = Pati.Text & Chr(KeyAscii)
            KeyAscii = 0
            
            Call LocateQueueData(Pati.GetCurCard.名称, Pati.Text)

        End If
    End If
    
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, "排队叫号系统"
End Sub

Private Sub picCallFace_Resize()
    On Error Resume Next
    
    scCallInf.Left = 0
    scCallInf.Top = 0
    scCallInf.Width = picCallFace.Width
    
    rptCallList.Left = 0
    rptCallList.Top = scCallInf.Height
    rptCallList.Width = picCallFace.ScaleWidth
    If picCallFace.Height < 1800 Then
        rptCallList.Height = 1800
    Else
        rptCallList.Height = picCallFace.ScaleHeight - scCallInf.Height - 340
    End If
End Sub

Private Sub picLabel_Resize()
    On Error Resume Next
    chkOutQueue(0).Left = 30
    chkOutQueue(0).Top = Round((picLabel.ScaleHeight - chkOutQueue(0).Height) / 2)
    
    If chkOutQueue(1).Visible Then
        chkOutQueue(1).Left = chkOutQueue(0).Left + chkOutQueue(0).Width + 100
        chkOutQueue(1).Top = chkOutQueue(0).Top
        
        chkOutQueue(2).Left = chkOutQueue(1).Left + chkOutQueue(1).Width + 100
        chkOutQueue(2).Top = chkOutQueue(0).Top
    Else
        chkOutQueue(2).Left = chkOutQueue(0).Left + chkOutQueue(0).Width + 100
        chkOutQueue(2).Top = chkOutQueue(0).Top
    End If
    
    labError.Left = chkOutQueue(2).Left + chkOutQueue(2).Width + 100
    labError.Top = chkOutQueue(0).Top
End Sub


Private Sub InitQueueList()
On Error GoTo errh
    Dim Column As ReportColumn
    Dim str排队列宽 As String
    Dim str呼叫列宽 As String
    Dim strReg As String
    Dim blnIsCustom As Boolean
        
    strReg = "公共全局\排队叫号"
    
    str排队列宽 = GetSetting("ZLSOFT", strReg, "排队列宽度配置", C_STR_QUEUEQUEUE)
    str呼叫列宽 = GetSetting("ZLSOFT", strReg, "呼叫列宽度配置", C_STR_QUEUECALL)
    
    If UBound(Split(str排队列宽, ",")) <> 18 Then
        str排队列宽 = C_STR_QUEUEQUEUE
    End If
    If UBound(Split(str呼叫列宽, ",")) <> 18 Then
        str呼叫列宽 = C_STR_QUEUECALL
    End If
    
   '初始化呼叫队列显示字段
    Call Me.rptCallList.Columns.DeleteAll
    Call Me.rptQueueList.Columns.DeleteAll

    RaiseEvent OnInitQueueList(rptQueueList, rptCallList, blnIsCustom)
    mblnCustomCfg = blnIsCustom
    
    If Not blnIsCustom Then
        '原来的流程
        With Me.rptCallList.Columns
        
            rptCallList.AllowColumnRemove = False
            rptCallList.ShowItemsInGroups = False
            rptCallList.SkipGroupsFocus = True
            rptCallList.MultipleSelection = False
            rptCallList.AutoColumnSizing = False
            
            With rptCallList.PaintManager
                .ColumnStyle = xtpColumnShaded
                .GridLineColor = RGB(225, 225, 225)
                .NoGroupByText = "将列标题拖动到此,可按该列分组..."
                .NoItemsText = "没有可显示的项目..."
                .VerticalGridStyle = xtpGridSolid
            End With
            
            Set Column = .Add(mCol.队列名称, IIf(mlngQueueGroupType = 0, "", "队列"), Val(Split(str排队列宽, ",")(0)), False)
            If mlngQueueGroupType = 0 Then
                Column.Groupable = True
            Else
                Column.Visible = False
            End If
            
            Set Column = .Add(mCol.Id, "ID", Val(Split(str呼叫列宽, ",")(1)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.病人ID, "病人ID", Val(Split(str呼叫列宽, ",")(2)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.排队标记, "标记", Val(Split(str呼叫列宽, ",")(3)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.排队号码, "号码", Val(Split(str呼叫列宽, ",")(4)), True)
            Column.Visible = True
            
            Set Column = .Add(mCol.排队序号, "排队序号", Val(Split(str呼叫列宽, ",")(5)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.患者姓名, "患者姓名", Val(Split(str呼叫列宽, ",")(6)), True)
            Column.Visible = True
            
            Set Column = .Add(mCol.优先, "优先", Val(Split(str呼叫列宽, ",")(7)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.回诊序号, "回诊序号", Val(Split(str呼叫列宽, ",")(8)), True)
            Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",回诊序号,") > 0, True, False)
            
            Set Column = .Add(mCol.回诊排序号, "回诊排序号", Val(Split(str呼叫列宽, ",")(9)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.科室ID, "科室ID", Val(Split(str呼叫列宽, ",")(10)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.诊室, IIf(mlngQueueGroupType = 2, "", "诊室"), Val(Split(str呼叫列宽, ",")(11)), True)
            If mlngQueueGroupType = 2 Then
                Column.Groupable = True
                Column.Visible = False
            Else
                Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",诊室,") > 0, True, False)
            End If
            
            Set Column = .Add(mCol.医生姓名, IIf(mlngQueueGroupType = 1, "", "医生姓名"), Val(Split(str呼叫列宽, ",")(12)), True)
            If mlngQueueGroupType = 1 Then
                Column.Groupable = True
                Column.Visible = False
            Else
                Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",医生姓名,") > 0, True, False)
            End If
            
            Set Column = .Add(mCol.排队状态, "排队状态", Val(Split(str呼叫列宽, ",")(13)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.排队时间, "排队时间", Val(Split(str呼叫列宽, ",")(14)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.呼叫医生, "呼叫人", Val(Split(str呼叫列宽, ",")(15)), True)
            Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",呼叫医生,") > 0, True, False)
            
            Set Column = .Add(mCol.业务类型, "业务类型", Val(Split(str呼叫列宽, ",")(16)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.业务ID, "业务ID", Val(Split(str呼叫列宽, ",")(17)), False)
            Column.Visible = False
                    
            Set Column = .Add(mCol.呼叫时间, "呼叫时间", Val(Split(str呼叫列宽, ",")(18)), True)
            Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",呼叫时间,") > 0, True, False)
                    
            Set Column = .Add(mCol.排序名称, "排序名称", 0, False)
            Column.Visible = False
            
            Set Column = .Add(mCol.ORD, "ORD", 0, False)
            Column.Visible = False
                    
        End With
        
        With Me.rptCallList
            Set .Icons = zlCommFun.GetPubIcons
            
            .GroupsOrder.DeleteAll
    
            If mlngQueueGroupType = 0 Then
                .GroupsOrder.Add .Columns(mCol.排序名称)
            ElseIf mlngQueueGroupType = 1 Then
                .GroupsOrder.Add .Columns(mCol.医生姓名)
            Else
                .GroupsOrder.Add .Columns(mCol.诊室)
            End If
            
            .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
            
            .SortOrder.DeleteAll
            
            If mlngOrderStyle = 1 Then
                .SortOrder.Add .Columns(mCol.ORD)
                .SortOrder(0).SortAscending = True
            Else
            
                .SortOrder.Add .Columns(mCol.排队状态)
                .SortOrder(0).SortAscending = False
                
                .SortOrder.Add .Columns(mCol.排队序号)
                .SortOrder(1).SortAscending = True
                
                .SortOrder.Add .Columns(mCol.呼叫时间)
                .SortOrder(2).SortAscending = True
        
                .SortOrder.Add .Columns(mCol.排队号码)
                .SortOrder(3).SortAscending = True
            End If
        End With
        
        '初始化排队队列显示字段
        Call Me.rptQueueList.Columns.DeleteAll
        With Me.rptQueueList.Columns
            
            rptQueueList.AllowColumnRemove = False
            rptQueueList.ShowItemsInGroups = False
            rptQueueList.SkipGroupsFocus = True
            rptQueueList.MultipleSelection = False
            rptQueueList.AutoColumnSizing = False
            
            With rptQueueList.PaintManager
                .ColumnStyle = xtpColumnShaded
                .GridLineColor = RGB(225, 225, 225)
                .NoGroupByText = "将列标题拖动到此,可按该列分组..."
                .NoItemsText = "没有可显示的项目..."
                .VerticalGridStyle = xtpGridSolid
            End With
            
            Set Column = .Add(mCol.队列名称, IIf(mlngQueueGroupType = 0, "", "队列"), Val(Split(str排队列宽, ",")(0)), False)
              
            If mlngQueueGroupType = 0 Then
                Column.Groupable = True
            Else
                Column.Visible = False
            End If
                    
            Set Column = .Add(mCol.Id, "ID", Val(Split(str排队列宽, ",")(1)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.病人ID, "病人ID", Val(Split(str排队列宽, ",")(2)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.排队标记, "标记", Val(Split(str排队列宽, ",")(3)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.排队号码, "号码", Val(Split(str排队列宽, ",")(4)), True)
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",号码,") > 0, True, False)
            
            Set Column = .Add(mCol.排队序号, "排队序号", Val(Split(str排队列宽, ",")(5)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.患者姓名, "患者姓名", Val(Split(str排队列宽, ",")(6)), True)
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",患者姓名,") > 0, True, False)
            
            Set Column = .Add(mCol.优先, "优先", Val(Split(str排队列宽, ",")(7)), True)
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, "优先") > 0, True, False)
            
            Set Column = .Add(mCol.回诊序号, "回诊序号", Val(Split(str排队列宽, ",")(8)), True)
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",回诊序号,") > 0, True, False)
            
            Set Column = .Add(mCol.回诊排序号, "回诊排序号", Val(Split(str排队列宽, ",")(9)), True)
            Column.Visible = False
            
            Set Column = .Add(mCol.科室ID, "科室ID", Val(Split(str排队列宽, ",")(10)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.诊室, IIf(mlngQueueGroupType = 2, "", "诊室"), Val(Split(str排队列宽, ",")(11)), True)
            If mlngQueueGroupType = 2 Then
                Column.Groupable = True
                Column.Visible = False
            Else
                Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",诊室,") > 0, True, False)
            End If
            
            Set Column = .Add(mCol.医生姓名, IIf(mlngQueueGroupType = 1, "", "医生姓名"), Val(Split(str排队列宽, ",")(12)), True)
            If mlngQueueGroupType = 1 Then
                Column.Groupable = True
                Column.Visible = False
            Else
                Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",医生姓名,") > 0, True, False)
            End If
            
            Set Column = .Add(mCol.排队状态, "排队状态", Val(Split(str排队列宽, ",")(13)), True)
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",排队状态,") > 0, True, False)
            
            Set Column = .Add(mCol.排队时间, "排队时间", Val(Split(str排队列宽, ",")(14)), True)
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",排队时间,") > 0, True, False)
            
            Set Column = .Add(mCol.呼叫医生, "呼叫人", Val(Split(str排队列宽, ",")(15)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.业务类型, "业务类型", Val(Split(str排队列宽, ",")(16)), False)
            Column.Visible = False
            
            Set Column = .Add(mCol.业务ID, "业务ID", Val(Split(str排队列宽, ",")(17)), False)
            Column.Visible = False
                    
            Set Column = .Add(mCol.呼叫时间, "呼叫时间", Val(Split(str排队列宽, ",")(18)), False)
            Column.Visible = False
                    
            Set Column = .Add(mCol.排序名称, "排序名称", 0, False)
            Column.Visible = False
    
            Set Column = .Add(mCol.ORD, "ORD", 0, False)
            Column.Visible = False
        End With
        
        With Me.rptQueueList
            Set .Icons = zlCommFun.GetPubIcons
            
            .GroupsOrder.DeleteAll
    
            If mlngQueueGroupType = 0 Then
                .GroupsOrder.Add .Columns(mCol.排序名称)
            ElseIf mlngQueueGroupType = 1 Then
                .GroupsOrder.Add .Columns(mCol.医生姓名)
            Else
                .GroupsOrder.Add .Columns(mCol.诊室)
            End If
            
            .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
            
            '队列名称 = 0: Id:排队标记: 排队号码: 优先: 患者姓名: 科室ID:  诊室: 医生姓名:排队状态 : 排队时间: 业务ID
            '分组之后可能失去记录集中的顺序,因此强行加入排序列
            .SortOrder.DeleteAll
            If mlngOrderStyle = 1 Then
                .SortOrder.Add .Columns(mCol.ORD)
                .SortOrder(0).SortAscending = True
            Else
                .SortOrder.Add .Columns(mCol.排队状态)
                .SortOrder(0).SortAscending = True
                
                .SortOrder.Add .Columns(mCol.排队序号)
                .SortOrder(1).SortAscending = True
                
                .SortOrder.Add .Columns(mCol.优先)
                .SortOrder(2).SortAscending = False
        
                .SortOrder.Add .Columns(mCol.回诊排序号)
                .SortOrder(3).SortAscending = True
        
                .SortOrder.Add .Columns(mCol.排队时间)
                .SortOrder(4).SortAscending = True
        
                .SortOrder.Add .Columns(mCol.排队号码)
                .SortOrder(5).SortAscending = True
            End If
        End With
    End If

    Call DoReportCtlHeadInfo(rptQueueList, mTQueueCols)
    Call DoReportCtlHeadInfo(rptCallList, mTCallCols)
    
    If Not mblnIsGroup Then
        '删除分组
        Call rptQueueList.GroupsOrder.DeleteAll
        Call rptCallList.GroupsOrder.DeleteAll
    End If
    Exit Sub
errh:
    MsgBox "排队叫号InitQueueList执行错误" & err.Description, vbOKOnly, "排队叫号系统"
End Sub

Public Sub QueueParameterSetup(frmParent As Form, lngSys As Long)
'提供给接口的 打开排队配置界面方法

    '得到模块号和系统号
    glngSys = lngSys
    glngModul = 1160
    
    frmSetup.Show 1, frmParent
    
    On Error GoTo errHandle
        Call InitLocalParas(True)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Function IsAllowCall(ByVal lngBusinessType As Long, ByVal lngBusinessId As String) As Boolean
'检查是否允许呼叫
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    IsAllowCall = False
    
    strSql = "select 排队状态 from 排队叫号队列 where 业务类型=[1] and 业务ID=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBusinessType, lngBusinessId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    IsAllowCall = IIf(rsData!排队状态 = 0, True, False)
End Function



Private Sub comMenu_直呼()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    Dim lngSendHwnd As Long
    Dim lngSendResult As Long
    
    On Error GoTo errHandle
    
    If mstrCurrent队列名称 <> "" Then
        
        If Not mblnIsSelectedCallingList And Not IsAllowCall(mlngCurrentWorkType, mstrCurrentWorkID) Then
            MsgBox "当前数据可能已被呼叫或执行，请选择其他记录进行呼叫操作。", vbOKOnly Or vbInformation, "排队叫号系统"
            Exit Sub
        End If
        
        blnCancel = False
        Call DoQueueExecuteBefore(mstrCurrentWorkID, 1, blnCancel, strNewQueueName)
            
        If Not blnCancel Then
            strSql = "ZL_排队叫号队列_呼叫('" & mstrCurrent队列名称 & "',1,'" & mstrLoginUserName & "','" & mstr呼叫站点名称 & "'," & mlngCurrentWorkType & ",'" & mstrCurrentWorkID & "')"
            zlDatabase.ExecuteProcedure strSql, "直呼"
        
            Call DoQueueExecuteAfter(mstrCurrentWorkID, 1)
                    
            '呼叫后通知zlQueueShow，以便在显示有多页数据时，定位到当前呼叫病人。问题号:85290
            lngSendHwnd = FindWindow(vbNullString, "排队显示控制")
            
            If lngSendHwnd > 0 Then
                lngSendResult = PostMessage(lngSendHwnd, 1025, mlngCurrentQueueId, 0)
            End If
        End If
    End If
    
    '当勾选“医生主动呼叫后才允许在队列中接诊”后，在排队列表选择数据直呼或顺呼后mintDetonatEvent值没有改变仍为1
    '这时接诊按钮处于可用状态，再次点击排队列表时，不会执行MouseDown，所以接诊按钮还是处于可用状态，这时点击接诊时
    '会导致排队列表的信息直接进入就诊病人列表中，从而不符合业务逻辑，因此需将mintDetonatEvent的值改为非1，强制执行MouseDown
    mintDetonatEvent = 2
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub



Private Sub DoQueueExecuteBefore(ByVal str业务ID As String, ByVal byt操作类型 As Byte, blnCancel As Boolean, strNewQueueName As String)
On Error GoTo errHandle
    RaiseEvent OnQueueExecuteBefore(str业务ID, byt操作类型, blnCancel, strNewQueueName)
    Exit Sub
errHandle:
    err.Description = "OnQueueExecuteBefore事件错误>>" + err.Description
   If ErrCenter = 1 Then Resume
End Sub

Private Sub DoQueueExecuteAfter(ByVal str业务ID As String, ByVal byt操作类型 As Byte)
On Error GoTo errHandle
    RaiseEvent OnQueueExecuteAfter(str业务ID, byt操作类型)
    Exit Sub
errHandle:
    err.Description = "OnQueueExecuteAfter事件错误>>" + err.Description
   If ErrCenter = 1 Then Resume
End Sub

Private Sub DoRecevieDiagnose(ByVal str业务ID As String, ByVal lng业务类型 As Long)
On Error GoTo errHandle
    RaiseEvent OnRecevieDiagnose(str业务ID, lng业务类型)
    Exit Sub
errHandle:
    err.Description = "OnRecevieDiagnose事件错误>>" + err.Description
   If ErrCenter = 1 Then Resume
End Sub

'触发事件
Private Sub DoSelectionChanged(ByVal blnIsCallingList As Boolean, objDataRow As XtremeReportControl.ReportRow, cbrMain As XtremeCommandBars.CommandBars)
On Error GoTo errHandle
    RaiseEvent OnSelectionChanged(blnIsCallingList, objDataRow, cbrMain)
    Exit Sub
errHandle:
    err.Description = "OnSelectionChanged事件错误>>" + err.Description
   If ErrCenter = 1 Then Resume
End Sub




Private Sub comMenu_顺呼()
    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strCurWorkId As String
    Dim intCurWorkType As Integer
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    Dim strTempQueueName As String
    Dim lngSendHwnd As Long
    Dim lngSendResult As Long
    Dim lngQueueId As Long
    
    On Error GoTo errHandle
    
    
    '判断是否选择顺呼的数据
    If rptQueueList.SelectedRows.Count = 0 Then
        If rptQueueList.Rows.Count > 0 Then
            rptQueueList.Rows(0).Selected = True
            
             Call rptQueueList_SelectionChanged
             
             strTempQueueName = mstrCurrent队列名称
        Else
            MsgBox "没有数据被选择，不能执行该操作。", vbOKOnly Or vbInformation, "排队叫号系统"
            Exit Sub
        End If
    Else
        '获取队列名称
        If rptQueueList.SelectedRows(0).GroupRow <> True Then
            strTempQueueName = rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_队列名称).value
        Else
            strTempQueueName = rptQueueList.SelectedRows(0).Childs(0).Record.Item(mTQueueCols.lngColIndex_队列名称).value
        End If
    End If
    
    
    '如果是顺呼，则将焦点设置为顺呼列表
    Call SetFocusToQueueList
    
    If mstrCurrent队列名称 <> "" Then
    
        strCurWorkId = ""
        intCurWorkType = 0
        lngQueueId = 0
        
        For i = 0 To rptQueueList.Rows.Count - 1
            If rptQueueList.Rows(i).GroupRow <> True Then
                If rptQueueList.Rows(i).Record.Item(mTQueueCols.lngColIndex_队列名称).value = strTempQueueName Then
                    strSql = "select ID,业务类型,业务ID from 排队叫号队列 where 队列名称=[1] and 业务ID=[2] and 业务类型=[3] and 排队状态=0"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "顺呼", strTempQueueName, CStr(Nvl(rptQueueList.Rows(i).Record.Item(mTQueueCols.lngColIndex_业务ID).value)), Val(Nvl(rptQueueList.Rows(i).Record.Item(mTQueueCols.lngColIndex_业务类型).value)))
                    If Not rsTemp.EOF Then
                        intCurWorkType = Val(Nvl(rsTemp!业务类型))
                        strCurWorkId = Nvl(rsTemp!业务ID)
                        lngQueueId = Nvl(rsTemp!Id)
                        
                        Exit For
                    End If
                End If
            Else
                If rptQueueList.Rows(i).Childs(0).Record.Item(mTQueueCols.lngColIndex_队列名称).value = strTempQueueName Then
                    rptQueueList.Rows(i).Childs(0).Selected = True
                End If
            End If
        Next i
        
        If Trim(strCurWorkId) = "" Then Exit Sub
        
'        If Not IsAllowCall(intCurWorkType, strCurWorkId) Then
'            MsgBox "当前数据可能已被呼叫或执行，请选择其他记录进行呼叫操作。", vbOKOnly Or vbInformation, "排队叫号系统"
'            Exit Sub
'        End If
        
        blnCancel = False
        Call DoQueueExecuteBefore(strCurWorkId, 1, blnCancel, strNewQueueName)
            
        If Not blnCancel Then
            strSql = "ZL_排队叫号队列_呼叫('" & strTempQueueName & "',7,'" & mstrLoginUserName & "','" & mstr呼叫站点名称 & "'," & intCurWorkType & ",'" & strCurWorkId & "')"
            zlDatabase.ExecuteProcedure strSql, "顺呼"
            
            mstrCurrentWorkID = strCurWorkId
            mlngCurrentWorkType = intCurWorkType
            
            Call DoQueueExecuteAfter(strCurWorkId, 1)
            
            '呼叫后通知zlQueueShow，以便在显示有多页数据时，定位到当前呼叫病人。问题号:85290
            lngSendHwnd = FindWindow(vbNullString, "排队显示控制")
            
            If lngSendHwnd > 0 Then
                lngSendResult = PostMessage(lngSendHwnd, 1025, lngQueueId, 0)
            End If
        End If
    End If
    
    mintDetonatEvent = 2
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub comMenu_暂停()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
        If mstrCurrent队列名称 <> "" Then
        
            blnCancel = False
            Call DoQueueExecuteBefore(mstrCurrentWorkID, 3, blnCancel, strNewQueueName)
            
            If Not blnCancel Then
                strSql = "ZL_排队叫号队列_呼叫('" & mstrCurrent队列名称 & "',3,'" & mstrLoginUserName & "','" & mstr呼叫站点名称 & "'," & mlngCurrentWorkType & ",'" & mstrCurrentWorkID & "')"
                zlDatabase.ExecuteProcedure strSql, "暂停"
                
                Call DoQueueExecuteAfter(mstrCurrentWorkID, 3)
            End If
        End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub comMenu_完成()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
        If mstrCurrent队列名称 <> "" Then
            blnCancel = False
            Call DoQueueExecuteBefore(mstrCurrentWorkID, 4, blnCancel, strNewQueueName)
            
            If Not blnCancel Then
                strSql = "ZL_排队叫号队列_呼叫('" & mstrCurrent队列名称 & "',4,'" & mstrLoginUserName & "','" & mstr呼叫站点名称 & "'," & mlngCurrentWorkType & ",'" & mstrCurrentWorkID & "')"
                zlDatabase.ExecuteProcedure strSql, "完成"
                
                Call DoQueueExecuteAfter(mstrCurrentWorkID, 4)
            End If
        End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub comMenu_查找()
    On Error GoTo errHandle
    
    Call frmFind.ShowFind(mcnOracle, 0, Me)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub comMenu_广播()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
    If mstrCurrent队列名称 <> "" Then
    
        blnCancel = False
        Call DoQueueExecuteBefore(mstrCurrentWorkID, 5, blnCancel, strNewQueueName)
        
        If Not blnCancel Then
            'strSql = "ZL_排队语音呼叫_INSERT(" & mlngCurrentQueueId & ",'" & mstr呼叫站点名称 & "', 1)" '1表示广播
            strSql = "ZL_排队语音呼叫_INSERT(" & mlngCurrentQueueId & ",'" & mstr呼叫站点名称 & "', 0)" '1表示广播
            Call zlDatabase.ExecuteProcedure(strSql, "广播")
            
            Call DoQueueExecuteAfter(mstrCurrentWorkID, 5)
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub comMenu_优先()
    Dim strSql As String
    Dim strTempQueueName As String
    Dim strSelectedName As String
    
    On Error GoTo errHandle
    
    If mstrCurrent队列名称 <> "" Then
        With rptQueueList
            '判断是否已选择数据
            If .Rows.Count > 0 Then
                If .SelectedRows.Count = 0 Then .Rows(0).Selected = True
                
                '获取队列名称
                If .SelectedRows(0).GroupRow <> True Then
                    strTempQueueName = .SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_队列名称).value
                    strSelectedName = .SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_排队号码).value & .SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_患者姓名).value & "," & .SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_ID).value
                Else
                    strTempQueueName = .SelectedRows(0).Childs(0).Record.Item(mTQueueCols.lngColIndex_队列名称).value
                    strSelectedName = .SelectedRows(0).Childs(0).Record.Item(mTQueueCols.lngColIndex_排队号码).value & .SelectedRows(0).Childs(0).Record.Item(mTQueueCols.lngColIndex_患者姓名).value & "," & .SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_ID).value
                End If
                
            Else
                MsgBox "没有加载数据，不能执行该操作。", vbOKOnly Or vbInformation, "排队叫号系统"
                Exit Sub
            End If
        End With
        
        
        '调用优先原因窗体
        Call frmPriorityCause.ShowPriorityCause(Me, mstrCurrent队列名称, mstrCurrentWorkID, strTempQueueName, strSelectedName)
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    
End Sub

Private Sub comMenu_恢复()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
        If mstrCurrent队列名称 <> "" Then
        
            blnCancel = False
            Call DoQueueExecuteBefore(mstrCurrentWorkID, 0, blnCancel, strNewQueueName)
            
            If Not blnCancel Then
                strSql = "ZL_排队叫号队列_呼叫('" & mstrCurrent队列名称 & "',0,'" & mstrLoginUserName & "','" & mstr呼叫站点名称 & "'," & mlngCurrentWorkType & ",'" & mstrCurrentWorkID & "')"
                zlDatabase.ExecuteProcedure strSql, "恢复"
                
                Call DoQueueExecuteAfter(mstrCurrentWorkID, 0)
            End If
        End If
        
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub comMenu_接诊()
    On Error GoTo errHandle
    
        If mstrCurrent队列名称 <> "" Then
            Call DoRecevieDiagnose(mstrCurrentWorkID, mlngCurrentWorkType)
        End If
        
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub comMenu_弃号()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
        If mstrCurrent队列名称 <> "" Then
        
            blnCancel = False
            Call DoQueueExecuteBefore(mstrCurrentWorkID, 2, blnCancel, strNewQueueName)
            
            If Not blnCancel Then
                strSql = "ZL_排队叫号队列_呼叫('" & mstrCurrent队列名称 & "',2,'" & mstrLoginUserName & "','" & mstr呼叫站点名称 & "'," & mlngCurrentWorkType & ",'" & mstrCurrentWorkID & "')"
                zlDatabase.ExecuteProcedure strSql, "弃号"
        
                Call DoQueueExecuteAfter(mstrCurrentWorkID, 2)
            End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub comMenu_设置()
    frmSetup.Show 1, Me
    
On Error GoTo errHandle
    Call InitLocalParas(False)
    Call InitQueueList
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub comMenu_刷新()
    Call zlRefresh(mstr队列名称, mstrCurrent队列名称, mstrCurrentWorkID, mstr诊室条件, mstr医生条件, mstrExcludeData, mintViewDataType)
End Sub

Private Sub comMenu_修改()
    Dim str队列名称 As String
    Dim str患者姓名 As String
    Dim str诊室 As String
    Dim str医生姓名 As String
    Dim strSql As String
    Dim lng业务类型 As Long
    Dim str业务ID As String
    Dim lng科室ID As Long
    Dim lng病人id As Long
    Dim blnIsAllowChangePar As Boolean
    Dim blnIsAlreadyProcessPar As Boolean
    Dim rsRoom As ADODB.Recordset
    Dim rsDoctor As ADODB.Recordset

    On Error GoTo errHandle
    
    
    '已经呼叫的数据不能进行修改
    '记录当前的队列名称和工作ID
    lng病人id = Val(Nvl(rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_病人ID).value))
    str队列名称 = Nvl(rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_队列名称).value)
    str患者姓名 = Nvl(rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_患者姓名).value)
    str诊室 = Nvl(rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_诊室).value)
    str医生姓名 = Nvl(rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_医生姓名).value)
    lng业务类型 = Val(Nvl(rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_业务类型).value))
    
    str业务ID = Nvl(rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_业务ID).value)
    lng科室ID = Val(Nvl(rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_科室ID).value))
    
    RaiseEvent OnQueueRoomLoad(str业务ID, rsRoom, rsDoctor)
    
    Set frmUpdateInfo.mrsDoctorData = rsDoctor
    Set frmUpdateInfo.mrsRoomData = rsRoom
    frmUpdateInfo.mlngCurrentQueueId = mlngCurrentQueueId
    
    If frmUpdateInfo.zlShowMe(Me, mstr队列名称, str队列名称, str患者姓名, str诊室, str医生姓名) = True Then
        
        '修改队列信息
        
        If frmUpdateInfo.mblnIsAlreadyProcess = True Then
            Call comMenu_刷新
            Exit Sub
        End If
        
        If str队列名称 <> rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_队列名称).value Then
            
            On Error GoTo DBError
            Call mcnOracle.BeginTrans
            
            strSql = "ZL_排队叫号队列_DELETE('" & rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_队列名称).value & "','" & str业务ID & "')"
            zlDatabase.ExecuteProcedure strSql, "再删除原排队信息"
            
            
            '如果队列名称有所改变，则需要出队再入队
            strSql = "ZL_排队叫号队列_INSERT('" & str队列名称 & "'," & lng业务类型 & ",'" & str业务ID _
                & "'," & lng科室ID & ",0,null,'" & str患者姓名 & "'," & lng病人id & ",'" & str诊室 & "','" & str医生姓名 & "', sysdate)"
            zlDatabase.ExecuteProcedure strSql, "先加入队列"
            
            Call mcnOracle.CommitTrans
            Exit Sub
DBError:
            Call mcnOracle.RollbackTrans
            
        Else    '没有修改队列名称，则直接修改信息即可
            strSql = "ZL_排队叫号队列_UPDATE('" & str队列名称 & "'," & lng业务类型 & ",'" & rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_业务ID).value _
                    & "'," & rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_科室ID).value & ",'" & str患者姓名 & "','" _
                    & str诊室 & "','" & str医生姓名 & "')"
            zlDatabase.ExecuteProcedure strSql, "修改队列信息"
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub


Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
      Case conMenu_Queue_LocateNew
        Control.Visible = mblnIsShowFindTools
      Case conMenu_Queue_Abandon '弃号
        Control.Visible = IIf(InStr(mstrPrivs, "弃号") <= 0, False, True)
        Control.Enabled = mblnFuncState(2)
      Case conMenu_Queue_Broadcast '广播
        Control.Visible = IIf(InStr(mstrPrivs, "广播") <= 0, False, True)
        Control.Enabled = mblnFuncState(5) And mblnIsSelectedCallingList
      Case conMenu_Queue_CallFirst '优先
        Control.Visible = IIf(InStr(mstrPrivs, "优先") <= 0, False, True)
        Control.Enabled = Not mblnIsSelectedCallingList
      Case conMenu_Queue_CallNext  '顺呼
        Control.Visible = IIf(InStr(mstrPrivs, "顺呼") <= 0, False, True)
      Case conMenu_Queue_RecDiagnose  '接诊
        Control.Visible = IIf(InStr(mstrPrivs, "接诊") <= 0, False, True)
        Control.Enabled = mblnFuncState(7)
      Case conMenu_Queue_Find   '查找
        Control.Visible = IIf(InStr(mstrPrivs, "查找") <= 0, False, True)
'      Case conMenu_Queue_Pause   '暂停
'        Control.Visible = IIf(InStr(mstrPrivs, "暂停") <= 0, False, True) And Not mblnIsSelectedCallingList
      Case conMenu_Queue_CallThis '直呼
        Control.Visible = IIf(InStr(mstrPrivs, "直呼") <= 0, False, True)
        Control.Enabled = mblnFuncState(1)
      Case conMenu_Queue_Finaled  '完成
        Control.Visible = IIf(InStr(mstrPrivs, "完成") <= 0, False, True)
        Control.Enabled = mblnFuncState(4)
      Case conMenu_Queue_Pause  '暂停
        Control.Visible = IIf(InStr(mstrPrivs, "暂停") <= 0, False, True)
        Control.Enabled = mblnFuncState(3) And Not mblnIsSelectedCallingList
        chkOutQueue(1).Visible = Control.Visible
        picLabel_Resize
        
'      Case conMenu_Queue_ReCall '重呼
'        Control.Visible = IIf(InStr(mstrPrivs, "重呼") <= 0, False, True)
      Case conMenu_Queue_Setup  '设置
        Control.Visible = IIf(InStr(mstrPrivs, "设置") <= 0, False, True)
      Case conMenu_Queue_Update '修改
        Control.Visible = IIf(InStr(mstrPrivs, "修改") <= 0, False, True)
        Control.Enabled = Not mblnIsSelectedCallingList
      Case conMenu_Queue_Restore '恢复
        Control.Visible = IIf(InStr(mstrPrivs, "恢复") <= 0, False, True)
        Control.Enabled = mblnFuncState(0)
'      Case conMenu_Queue_ComeBack '回诊
'        Control.Visible = IIf(InStr(mstrPrivs, "回诊") <= 0, False, True)
'        Control.Enabled = mblnFuncState(6)
    End Select
End Sub


Private Sub picQueueFace_Resize()
    On Error Resume Next
    scQueueInf.Left = 0
    scQueueInf.Top = 0
    scQueueInf.Width = picQueueFace.Width
    
    rptQueueList.Left = 0
    rptQueueList.Top = scQueueInf.Height
    rptQueueList.Width = picQueueFace.ScaleWidth
    If picQueueFace.Height < 1800 Then
        rptQueueList.Height = 1800
    Else
        rptQueueList.Height = picQueueFace.ScaleHeight - scQueueInf.Height - 340
    End If
End Sub

Private Sub rptCallList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    Dim objReportRow As XtremeReportControl.ReportRow
    
    mblnIsSelectedCallingList = True
    
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    If mintDetonatEvent <> 2 Then
        mintDetonatEvent = 2
        mblnNotRefresh = True '执行事件OnSelectionChanged后无需刷新
        
        If rptCallList.Rows.Count < 1 Then
           Set objReportRow = Nothing
           
           Call DoSelectionChanged(False, objReportRow, cbrMain)
        Else
            
            If rptCallList.SelectedRows.Count < 1 Then
               Set objReportRow = Nothing
               Call DoSelectionChanged(False, objReportRow, cbrMain)
            Else
               Set objReportRow = rptCallList.SelectedRows(0)
               Call DoSelectionChanged(True, objReportRow, cbrMain)
            End If
        End If
        
        mblnNotRefresh = False
        '触发OnSelectionChanged事件
'        RaiseEvent OnSelectionChanged(False, objReportRow, cbrMain)
    End If
End Sub



Private Sub rptCallList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    On Error GoTo errHandle
    
        If Not CheckIsSelectedData Then
            MsgBox "没有选择需要执行的数据，不能执行该操作。", vbInformation, "排队叫号系统"
            Exit Sub
        End If
        
        If Not CheckQueueDataIsHas(mlngCurrentQueueId) Then
            MsgBox "数据不存在或已被执行，请运行刷新操作。", vbInformation, "排队叫号系统"
            Exit Sub
        End If
            
        Call comMenu_接诊
        
        Call zlRefresh(mstr队列名称, mstrCurrent队列名称, mstrCurrentWorkID, mstr诊室条件, mstr医生条件, mstrExcludeData, mintViewDataType)
        
        Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub rptCallList_SelectionChanged()
    On Error GoTo errHandle
    
    If rptCallList.SelectedRows.Count = 0 Then Exit Sub
    
    If rptCallList.SelectedRows(0).GroupRow = True Then
        If rptCallList.SelectedRows(0).Childs.Count > 0 Then
            mstrCurrent队列名称 = rptCallList.SelectedRows(0).Record.Childs(0).Item(mTCallCols.lngColIndex_队列名称).value
            
            mstrCurrentWorkID = rptCallList.SelectedRows(0).Childs(0).Record.Item(mTCallCols.lngColIndex_业务ID).value
            mlngCurrentWorkType = StrNvl(rptCallList.SelectedRows(0).Childs(0).Record.Item(mTCallCols.lngColIndex_业务类型).value, 0)
            mlngCurrentQueueId = Val(rptCallList.SelectedRows(0).Childs(0).Record.Item(mTCallCols.lngColIndex_ID).value)
        End If
        
        Exit Sub
    End If

    '记录当前的队列名称和工作ID
    mstrCurrent队列名称 = rptCallList.SelectedRows(0).Record.Item(mTCallCols.lngColIndex_队列名称).value
    mstrCurrentWorkID = rptCallList.SelectedRows(0).Record.Item(mTCallCols.lngColIndex_业务ID).value
    mlngCurrentWorkType = StrNvl(rptCallList.SelectedRows(0).Record.Item(mTCallCols.lngColIndex_业务类型).value, 0)
    mlngCurrentQueueId = Val(rptCallList.SelectedRows(0).Record.Item(mTCallCols.lngColIndex_ID).value)
    
    mblnNotRefresh = True '执行事件OnSelectionChanged后无需刷新
    Call DoSelectionChanged(True, rptCallList.SelectedRows(0), cbrMain)
    
    mblnNotRefresh = False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub rptQueueList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objReportRow As XtremeReportControl.ReportRow
    
    mblnIsSelectedCallingList = False
    
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    
    If mintDetonatEvent <> 1 Then
        mintDetonatEvent = 1
        mblnNotRefresh = True '执行事件OnSelectionChanged后无需刷新
        
        If rptQueueList.Rows.Count < 1 Then
           Set objReportRow = Nothing
        Else
            
            If rptQueueList.SelectedRows.Count < 1 Then
               Set objReportRow = Nothing
            Else
               Set objReportRow = rptQueueList.SelectedRows(0)
            End If
        End If
        
        '触发OnSelectionChanged事件
        Call DoSelectionChanged(False, objReportRow, cbrMain)
        
        mblnNotRefresh = False
    End If
End Sub

Private Sub rptQueueList_SelectionChanged()
    On Error GoTo errHandle
    
    If rptQueueList.SelectedRows.Count = 0 Then Exit Sub
    
    If rptQueueList.SelectedRows(0).GroupRow = True Then
        If rptQueueList.SelectedRows(0).Childs.Count > 0 Then
            mstrCurrent队列名称 = rptQueueList.SelectedRows(0).Childs(0).Record.Item(mTQueueCols.lngColIndex_队列名称).value
            
            mstrCurrentWorkID = rptQueueList.SelectedRows(0).Childs(0).Record.Item(mTQueueCols.lngColIndex_业务ID).value
            mlngCurrentWorkType = StrNvl(rptQueueList.SelectedRows(0).Childs(0).Record.Item(mTQueueCols.lngColIndex_业务类型).value, 0)
            mlngCurrentQueueId = Val(rptQueueList.SelectedRows(0).Childs(0).Record.Item(mTQueueCols.lngColIndex_ID).value)
        End If
        
        Exit Sub
    End If

    '记录当前的队列名称和工作ID
    mstrCurrent队列名称 = rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_队列名称).value
    mstrCurrentWorkID = rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_业务ID).value
    mlngCurrentWorkType = StrNvl(rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_业务类型).value, 0)
    mlngCurrentQueueId = Val(rptQueueList.SelectedRows(0).Record.Item(mTQueueCols.lngColIndex_ID).value)
    
    mblnNotRefresh = True '执行事件OnSelectionChanged后无需刷新
    Call DoSelectionChanged(False, rptQueueList.SelectedRows(0), cbrMain)
    
    mblnNotRefresh = False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub MSSoundPlay(ByVal strConnetxt As String, ByVal lngSoundSpeed As Long)
    On Error Resume Next
    
    If objVoice Is Nothing Then
        Set objVoice = CreateObject("SAPI.SpVoice")
    End If
    
    objVoice.Rate = lngSoundSpeed   '速度:-10,10  0
    objVoice.Volume = 100 '声音:0,100   100
    
    objVoice.Speak strConnetxt, 1

End Sub


Private Sub tmrBroadCast_Timer()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim start As Date
    Dim strCallingContext As String
    
    On Error GoTo err
    
    tmrBroadCast.Enabled = False
    '显示排队队列
    Call ShowQueue
    
    
    '如果没有启用呼叫功能,则直接退出
    If Not mbln启用语音呼叫 Then Exit Sub
    '播放语音广播 如果播放方式为1 ，说明使用的是远端语音
    If mlng呼叫方式 = 1 Then Exit Sub
    
    
    strSql = "Select 呼叫内容,ID from 排队语音呼叫 where 站点=[1]  order by id "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "语音呼叫", mstr呼叫站点名称)
        
    While rsTemp.EOF = False
        '显示排队队列（注：每次呼叫时，可能需要较长的时间，因此需要调用该方法即时刷新一些参数）
        Call ShowQueue
                            
        strCallingContext = Nvl(rsTemp!呼叫内容)
                            
        strSql = "ZL_排队语音呼叫_DELETE(" & Nvl(rsTemp!Id) & ")"
        zlDatabase.ExecuteProcedure strSql, "语音呼叫完成"
        
        mlngCurPlayCount = 0
        While (mlngCurPlayCount < mlng语音播放次数)
            If mstr语音类型 = MS_SOUND_TYPE Then
                Call MSSoundPlay(strCallingContext, mlng语音广播语速)
            Else
                Call StartTextPlay(strCallingContext, mlng语音广播语速 * 10)
            End If
            
            mlngCurPlayCount = mlngCurPlayCount + 1
                                            
            start = Timer
            
            Do While Timer < start + mint语音广播时间长度
                Call Sleep(5)
                
                DoEvents
                
                '如果程序关闭，则退出
                If mIsUnload Then
                    Call StopPlayStr
                    
                    tmrBroadCast.Enabled = False
                    Exit Sub
                End If
            Loop
        Wend
           
        '如果程序关闭，则退出
        If mIsUnload Then
            tmrBroadCast.Enabled = False
            
            Exit Sub
        End If
            
        rsTemp.MoveNext
    Wend
    
    tmrBroadCast.Interval = mlng轮询时间 * 1000
    tmrBroadCast.Enabled = True
    
    Exit Sub
err:
    Call SaveErrLog
    
    labError.Caption = err.Description
        
    tmrBroadCast.Interval = mlng轮询时间 * 1000
    tmrBroadCast.Enabled = True
End Sub


Public Function QueueBroadcastCall(ByVal str呼叫文本 As String) As Boolean


    Dim start As Date
    On Error GoTo err
    
    '初始化参数
    Call InitLocalParas(True)

    QueueBroadcastCall = False
    
    '如果没有启用呼叫功能,则直接退出
    If Not mbln启用语音呼叫 Then Exit Function
    '播放语音广播 如果播放方式为1 ，说明使用的是远端语音
    If mlng呼叫方式 = 1 Then Exit Function

        
        mlngCurPlayCount = 0
        While (mlngCurPlayCount < mlng语音播放次数)
            If mstr语音类型 = MS_SOUND_TYPE Then
                Call MSSoundPlay(str呼叫文本, mlng语音广播语速)
            Else
                Call StartTextPlay(str呼叫文本, mlng语音广播语速 * 10)
            End If
            
            mlngCurPlayCount = mlngCurPlayCount + 1
                                            
            start = Timer
            
            Do While Timer < start + mint语音广播时间长度
                Call Sleep(5)
                
                DoEvents
                
                '如果程序关闭，则退出
                If mIsUnload Then
                    Call StopPlayStr
                    Exit Function
                End If
            Loop
        Wend
        
        QueueBroadcastCall = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Sub ShowQueue()

    On Error GoTo errHandle
    
    '显示排队队列
    If mbln显示排队队列 = True Then
        If pobjLEDShow Is Nothing Then
            Call InitLED(plngLEDModal)
        End If
        
        Call pobjLEDShow.zlShow(mcnOracle, mstr队列名称, mstr诊室条件, mstr医生条件, mstrExcludeData, mintViewDataType, mlng回诊病人优先 = 1)
    Else
        If Not pobjLEDShow Is Nothing Then
            '关闭LCD窗口
            Call pobjLEDShow.zlClose
            Set pobjLEDShow = Nothing
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function FindQueue(ByVal strLocateType As String, ByVal strLocateValue As String) As Boolean
    On Error Resume Next
    
    FindQueue = False
    
    If Trim(strLocateValue) <> "" Then
        FindQueue = LocateQueueData(strLocateType, Trim(strLocateValue))
    End If
End Function

Private Function DeptNametransform(ByVal strOldName) As String
'部门名称转化，目前只支持 一到十的处理 将小写数字转化为 abc 这种形式便于排序
    Dim strWord As String '单个字符
    Dim intCount As Integer
    Dim i As Integer
    
    On Error GoTo errh
    DeptNametransform = strOldName
    
    intCount = 0
    For i = 1 To Len(strOldName)
        strWord = Mid(strOldName, i, 1)
        If strWord = "一" Or strWord = "二" Or strWord = "三" Or strWord = "四" Or strWord = "五" Or strWord = "六" Or _
           strWord = "七" Or strWord = "八" Or strWord = "九" Or strWord = "十" Then
            intCount = intCount + 1
        End If
    Next
    
    If intCount = 1 Then
        DeptNametransform = Replace(strOldName, "一", "a")
        DeptNametransform = Replace(DeptNametransform, "二", "b")
        DeptNametransform = Replace(DeptNametransform, "三", "c")
        DeptNametransform = Replace(DeptNametransform, "四", "d")
        DeptNametransform = Replace(DeptNametransform, "五", "e")
        DeptNametransform = Replace(DeptNametransform, "六", "f")
        DeptNametransform = Replace(DeptNametransform, "七", "g")
        DeptNametransform = Replace(DeptNametransform, "八", "h")
        DeptNametransform = Replace(DeptNametransform, "九", "i")
        DeptNametransform = Replace(DeptNametransform, "十", "j")
    End If

    Exit Function
errh:
    DeptNametransform = strOldName
End Function
Private Sub DoReportCtlHeadInfo(ByRef reportCrt As ReportControl, ByRef ObjTColsInfo As TColsInfo)
'根据初始化的列头信息
On Error GoTo errh
    Dim i As Integer
    Dim iCount As Integer
    Dim ColumnSub As ReportColumn
    
    With reportCrt
        For i = 0 To .Columns.Count - 1
            Set ColumnSub = .Columns(i)
            If Not ColumnSub Is Nothing Then
                Select Case ColumnSub.Caption
                    
                    Case C_STR_COL_ID
                        ObjTColsInfo.lngColIndex_ID = ColumnSub.Index
                    Case C_STR_COL_病人ID
                        ObjTColsInfo.lngColIndex_病人ID = ColumnSub.Index
                    Case C_STR_COL_队列名称
                        ObjTColsInfo.lngColIndex_队列名称 = ColumnSub.Index
                    Case C_STR_COL_业务ID
                        ObjTColsInfo.lngColIndex_业务ID = ColumnSub.Index
                    Case C_STR_COL_科室ID
                        ObjTColsInfo.lngColIndex_科室ID = ColumnSub.Index
                    Case C_STR_COL_排队号码
                        ObjTColsInfo.lngColIndex_排队号码 = ColumnSub.Index
                    Case C_STR_COL_患者姓名
                        ObjTColsInfo.lngColIndex_患者姓名 = ColumnSub.Index
                    Case C_STR_COL_诊室
                        ObjTColsInfo.lngColIndex_诊室 = ColumnSub.Index
                    Case C_STR_COL_医生姓名
                        ObjTColsInfo.lngColIndex_医生姓名 = ColumnSub.Index
                    Case C_STR_COL_业务类型
                        ObjTColsInfo.lngColIndex_业务类型 = ColumnSub.Index
                End Select
                    
            End If
        Next
    End With
    
    Exit Sub
errh:
    MsgBox "获取列信息错误,排队叫号功能不能正常使用，请联系软件技术人员解决" & err.Description, vbInformation, "排队叫号系统"
End Sub

Private Sub InitPati()
On Error GoTo errh
    Dim strKinds As String
    Dim bl存在就诊卡 As Boolean
    Dim i As Integer
    
    '创建卡结算部件
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")

    '初始化卡结算部件
    mobjSquareCard.zlInitComponents Me, G_LNG_QUEUEMANAGE_MODULENUM, glngSys, mstrLoginUserName, mcnOracle
    bl存在就诊卡 = False
    
    For i = 1 To mobjSquareCard.zlGetCards(1).Count
        If mobjSquareCard.zlGetCards(1).Item(i).名称 = "就诊卡" Then
            bl存在就诊卡 = True
            Exit For
        End If
    Next
    
    If bl存在就诊卡 Then
        strKinds = ""
    Else
        strKinds = "就诊卡|就诊卡号|-1;"
    End If
    
    strKinds = strKinds & "姓名|姓名|-1;"
    strKinds = strKinds & "门诊号|门诊号|-1;"
    strKinds = strKinds & "医保号|医保号|-1;"
    strKinds = strKinds & "排队号|排队号|-1;"
    
    Pati.zlInit Me, glngSys, G_LNG_QUEUEMANAGE_MODULENUM, mcnOracle, mstrLoginUserName, mobjSquareCard, strKinds
    Pati.IDKindIDX = Pati.GetKindIndex(mstrLocateType)
    Exit Sub
errh:
    err.Raise -1, , "初始化刷卡查询控件失败" & err.Description, vbInformation, "排队叫号系统"
End Sub

Private Function LocateQueueData(ByVal findType As String, ByVal findData As String) As Boolean
'调整LocateQueueData处理方式：
'Pati.GetCurCard.接口序号 > 0 使用lngPatientID 定位 ，否则，分为 姓名，排队号  门诊号/医保号 3种情况处理  跟pati控件初始化的卡类型有关
'
On Error GoTo errh
    Dim i As Integer
    Dim j As Integer
    Dim lngPatientID As Long
    Dim blnFind As Boolean
    
    Dim lngFindIndex As Long '用于定位的字段索引
    Dim strFindValue As String '实际用于定位的值
    
    lngFindIndex = -1
    
    If Pati.GetCurCard.接口序号 > 0 Then
        If mobjSquareCard.zlGetPatiID(Pati.GetCurCard.接口序号, Pati.Text, , lngPatientID) Then
            strFindValue = lngPatientID
            lngFindIndex = mTQueueCols.lngColIndex_病人ID
        Else
            Exit Function
        End If
    Else
        '这里的名称跟 InitPati 中初始化的类型有关
        If findType = "姓名" Then
            strFindValue = findData
        ElseIf findType = "排队号" Then
            lngFindIndex = mTQueueCols.lngColIndex_排队号码
            strFindValue = findData
        ElseIf findType = "医保号" Or findType = "门诊号" Or findType = "就诊卡号" Then
            If mobjSquareCard.zlGetPatiID(findType, Pati.Text, , lngPatientID) Then
                strFindValue = lngPatientID
                lngFindIndex = mTQueueCols.lngColIndex_病人ID
            Else
                Exit Function
            End If
        Else
            '不明确的卡类型，不能定位
            Exit Function
        End If
    End If
        
    LocateQueueData = False
    
    If findType <> "姓名" And lngFindIndex = -1 Then Exit Function
    
    If mblnIsGroup Then
        For i = 0 To rptQueueList.Rows.Count - 1
            If rptQueueList.Rows(i).GroupRow = True Then
                For j = 0 To rptQueueList.Rows(i).Childs.Count - 1
                    blnFind = False
                    If findType = "姓名" Then
                        blnFind = IIf(rptQueueList.Rows(i).Childs(j).Record.Item(mTQueueCols.lngColIndex_患者姓名).value Like findData & "*", True, False)
                    Else
                        blnFind = IIf(rptQueueList.Rows(i).Childs(j).Record.Item(lngFindIndex).value = strFindValue, True, False)
                    End If
                    
                    If blnFind Then
                    
                        rptQueueList.Rows(i).Expanded = True
                        rptQueueList.Rows(i).Childs(j).Selected = True
                        
                        mblnIsSelectedCallingList = False
                        Call SwitchActiveWindow(mblnIsSelectedCallingList)
                        
                        Call rptQueueList.SetFocus
                        
                        LocateQueueData = True
                        
                        Exit Function
                    End If
                Next j
            End If
        Next i
    End If
        
    '如果没有找到数据，则从已呼叫队列中查找
    If mblnIsGroup Then
        For i = 0 To rptCallList.Rows.Count - 1
            If rptCallList.Rows(i).GroupRow = True Then
                For j = 0 To rptCallList.Rows(i).Childs.Count - 1
                
                    blnFind = False
                    If findType = "姓名" Then
                        blnFind = IIf(rptCallList.Rows(i).Childs(j).Record.Item(mTCallCols.lngColIndex_患者姓名).value Like findData & "*", True, False)
                    Else
                        blnFind = IIf(rptCallList.Rows(i).Childs(j).Record.Item(lngFindIndex).value = strFindValue, True, False)
                    End If
                    
                    If blnFind Then
                    
                        rptCallList.Rows(i).Expanded = True
                        rptCallList.Rows(i).Childs(j).Selected = True
                        
                        mblnIsSelectedCallingList = True
                        Call SwitchActiveWindow(mblnIsSelectedCallingList)
                        
                        Call rptCallList.SetFocus
                        
                        LocateQueueData = True
                        
                        Exit Function
                    End If
                Next j
            End If
        Next i
    End If

    Exit Function
errh:
    err.Raise -1, , "定位到列表" & err.Description, vbInformation, "排队叫号系统"
End Function

