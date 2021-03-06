VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport 
   Caption         =   "PACS 报告编辑"
   ClientHeight    =   9330
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   10950
   ClipControls    =   0   'False
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   10950
   Begin VB.Timer tmrCheckingReportState 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3240
      Tag             =   "0"
      Top             =   840
   End
   Begin VB.PictureBox picReportHistoryList 
      Height          =   5895
      Left            =   5400
      ScaleHeight     =   5835
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CheckBox chkOtherDeptReport 
         Caption         =   "查看其他科历史报告"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2655
      End
      Begin VB.Frame frmReportDetail 
         Height          =   4575
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   3615
         Begin VB.CommandButton cmdViewImage 
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            Picture         =   "frmReport.frx":0CCA
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "查看患者本次检查的影像"
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton cmdSelectWord 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            Picture         =   "frmReport.frx":1F6C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "将当前选中的文本写入报告"
            Top             =   240
            Width           =   1200
         End
         Begin RichTextLib.RichTextBox rtxtReport 
            Height          =   3495
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   6165
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmReport.frx":3166
         End
      End
      Begin MSComctlLib.ListView lvHistoryList 
         Height          =   930
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1640
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "dfd"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "dsd"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin zlRichEditor.Editor edtEditor 
      Height          =   1095
      Left            =   2760
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1931
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   5040
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "宋体"
   End
   Begin RichTextLib.RichTextBox rtxtSaveElement 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmReport.frx":3203
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   1080
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmReport.frx":3292
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请重新选择。"
Private Const M_STR_MODULE_MENU_TAG As String = "报告"

Private mlngModule As Long
Private mstrPrivs As String         '权限字符串
Private mlngDeptID As Long          '当前科室ID
Private mobjOwner As Object


Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mlngAdviceID As Long        '医嘱ID
Private mlngSendNo As Long          '发送号
Private mblnMoved As Boolean        '是否被转储
Private mlngStudyState As Long

Private WithEvents mfrmReportView As frmReportView
Attribute mfrmReportView.VB_VarHelpID = -1
Private mfrmReportImage As frmReportImage
Private mfrmReportSpecial As Object
Private WithEvents pobjPacsCore As zl9PacsCore.clsViewer     '观片站对象
Attribute pobjPacsCore.VB_VarHelpID = -1
Private WithEvents mfrmReportWord As frmReportWord          '词句示范窗体
Attribute mfrmReportWord.VB_VarHelpID = -1
Private WithEvents mobjCustomReport As clsReport                  '自定义报表对象
Attribute mobjCustomReport.VB_VarHelpID = -1
Private WithEvents mobjReport As zlRichEPR.cDockReport      '报告对象
Attribute mobjReport.VB_VarHelpID = -1
Private mobjWork_ActiveVideo As Object ' zl9PacsCapture.clsPacsCapture  '视频采集模块
Attribute mobjWork_ActiveVideo.VB_VarHelpID = -1

Private mblnUseActiveCapture As Boolean
Private mblnSingleWindow As Boolean     '是否使用独立窗口显示报告编辑器，True-独立窗口显示；False-嵌入式显示
Private mlngEPRDeptID As Long   '当前报告中“电子病历记录”所记录的科室ID
Private mstrEPR创建人 As String '当前报告中的“电子病历记录”所记录的创建人
Private mstrEPR保存人 As String '当前报告中的“电子病历记录”所记录的保存人
Private mlngEPR签名级别 As Long '当前报告中的“电子病历记录”所记录的签名级别
Private mdtReportTime As Date   '报告保存时间
Private mlngPassType As Long                 '密码验证规则（系统参数） 0-密码；1－数字；2－两者皆可

Private mFileID As Long         '病历文件ID,病历格式文件
Private mReportID As Long       '病历内容文件ID
Private mFormatID As Long     '病历范文ID
Private mModelName As String     '病历名称
Private mintEditType As Integer '病历状态 0 创建，1书写，2 修订
Private mintReportViewType As Integer ' 0-检查所见CheckView，1-诊断意见Result，2-建议Advice
Private miES As Integer
Private miEE As Integer



Private mHasChangeFormat As Boolean     '记录是否更改了格式


Private mblnModified As Boolean              '报告内容是否改变
Private mblnReadOnly As Boolean         '是否只读状态，不可以再修改报告

Public mblnEditable As Boolean         '是否可以编辑报告

Private mstrModifyEdit As String        '当前报告是否在修订状态被其他人修订保存后没有签名？记录保存人的姓名，空表示不是这种情况
Private mblnCanUntread As Boolean       '是否允许回退。当报告已经被打印，而且被审核后，不允许回退

Private mSigns As New cEPRSigns         '当前文档中的签名
Private m最后版本 As Integer         '最后版本
Private m目标版本 As Integer        '目标版本
Private m签名级别 As EPRSignLevelEnum        '1-书写;2-主治医师审阅;3-主任医师审阅。住院病历以外的病历只有书写和审阅状态
Private mModified As Boolean
Public mblnShowImage As Boolean            '是否显示报告图像
Private mblnShowSpecial As Boolean         '是否显示专科报告
Public mblnShowVideoCapture As Boolean     '是否显示图像采集


Private mstrSpecialForm As String           '专科报告窗体名称
Private mlngShowBigImg As Long              '报告中显示大图
Private mdblBigImgZoom As Double            '报告大图放大倍数
Private mintMinImageCount As Integer        '报告缩略图显示数量
Private mblnExitAfterPrint As Boolean       '报告打印后关闭窗体
Private mintImageDblClick As Integer        '缩略图双击后的操作：0--直接写入报告；1--打开图像编辑窗口

Private mblnIgnoreResult As Boolean         '忽略结果阴阳性
Private mintCriticalValues As Integer                       '危急值
Private mintConformDetermine As Integer                     '符合情况
Private mstrImageLevel As String                            '影像质量等级串
Private mstrReportLevel As String                           '报告质量等级串
Private mlngHintType As Long

Private mblnReportWithResult As Boolean      '无影像诊断为阴性

Private mblnShowWord As Boolean             '是否常态显示词句示范，True-一直显示词句示范，False--双击标题才显示词句示范窗口
Private mintWordDblClick As Integer         '词句双击后的操作：0--直接写入报告；1--打开词句编辑窗口
Private mblnRptImg2CapImg As Boolean
Private mstrFormatInfo As String
Private mblnCheckPrintPara As Boolean         '平诊需要审核才能打印 =true，参数定义
Private mblnCanPrint As Boolean             '该病人的报告，是否允许打印
Private mblnCheckOtherDeptReport As Boolean     '是否通过历史报告功能查看其他科的历史报告
Private mblnUntreadPrinted As Boolean           '审核打印后是否允许回退，True--可以回退；False--不可以回退。
Private mblnPrintView As Boolean            '控制在未找到对应病历文件的情况下 “打印”“预览”按钮的禁用状态，true 为禁用  false 为不禁用
Private mblnIsReportDelete As Boolean      '是否已删除报告单据
Private mblnTechReptSame As Boolean        '只能填写自己检查的报告

Private mblnIsPetitionScan As Boolean      '是否启用申请单扫描


Private mobjFSO As New Scripting.FileSystemObject    'FSO对象
Private mclsUnzip As New cUnzip
Private mclsZip As New cZip

Private mlngCY21 As Long                 '文本报告的高度
Private mlngCY22 As Long                 '专科报告的高度
Private mlngCX1 As Long                  '模板的宽度
Private mlngCX2 As Long                  '文本报告的宽度
Private mlngCX3 As Long                  '图像区域的宽度
Private mlngCY3 As Long                  '图像区域的高度
Private mlngCX4 As Long                  '视频采集区域的宽度
Private mlngCY4 As Long                  '视频采集区域的高度
Private mlngPicHistoryY As Long          '报告历史区域的高度
Private mlngPicHistoryX As Long          '报告历史区域的宽度
Private mlngPrivateWordY As Long         '私人常用词句区域的高度
Private mblnExitAfterSign As Boolean     '签名后退出
Private mintPaneID As Integer             '当前选中的Pane ID

Private mblnPrintOK As Boolean           '打印完成

Private mblnMenuDownState As Boolean    '避免双击工具栏产生错误

'检查的信息
Private mstr医嘱内容 As String             '医嘱内容
Private mstr医嘱附件 As String        '医嘱附件

Private Type rptFormat
    ID As Long          '报告格式ID
    strName As String   '报告格式名称
End Type
Private rptFormats() As rptFormat

'设置自定义报表的打印格式
Private mbln使用自定义报表 As Boolean           '打印格式是否是自定义报表
Private mstr报表编号 As String                  '自定义报表的编号
Private mblnRefreshRptFormat As Boolean         '打印格式需要刷新
Private mstr选中报表格式 As String              '被选中的自定义报表格式

'词句示范的添加和修改
Private mintWordPower As Integer        '词句管理权范围
'    mintWordPower=-1，不具备词句管理权;
'    mintWordPower=0，全院，这时显示所有的示范，也可以更改;
'    mintWordPower=1，科室，这时显示全院通用示范(科室id is null)和所在科室公有或部门内人员私有的示范，但不能更改全院通用示范;
'    mintWordPower=2，个人，这时显示全院通用示范(科室id is null)和所在科室通用示范(人员id is null)和个人示范，仅个人示范可更改

Private Const Report_Element_报告签名 = "报告签名"

Private mObjActiveMenuBar As CommandBars
Private mblnRefreshState As Boolean

Public mblnClosed As Boolean        '判断该报告编辑器是否已经被关闭

'本窗体的事件
Public Event AfterOpen()
Public Event BeforeEdit(ByVal lngOrderID As Long)

'frmOwnerForm主要用于在该事件执行时,如果在该事件中有模态窗口显示，就需要制定owner拥有者，且必须使用该参数，否则会
'造成程序处于假死状态，如在电子病例签名后，选择阳性率，阳性率窗口在未设置该参数的情况下，将造成这种情况。
Public Event AfterSaved(ByVal lngOrderID As Long, frmOwnerForm As Form, ByVal lngSaveType As Long)
Public Event AfterClosed(ByVal lngOrderID As Long)
Public Event AfterPrinted(ByVal lngOrderID As Long)
Public Event AfterDeleted(ByVal lngOrderID As Long)

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


'获取菜单接口对象
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


'获取当前报告的医嘱ID
Property Get AdviceId()
    AdviceId = mlngAdviceID
End Property


'设置报告的图像处理对象
Property Get PacsCore() As zl9PacsCore.clsViewer
    Set PacsCore = pobjPacsCore
End Property

Property Set PacsCore(objPacsCore As zl9PacsCore.clsViewer)
    Set pobjPacsCore = objPacsCore
End Property


Public Sub NotificationRefresh()
'通知刷新
    mblnRefreshState = False
End Sub


Public Sub VideoCallBack(EventType As Long, lngAdviceID As Long, _
    Optional strStudyUID As String, Optional strPatientName As String, Optional blnIsLock As Boolean)
    '便于编译通过......
End Sub

Public Sub UpdateImageVideoState(ByVal lngEventType As TVideoEventType, ByVal lngAdviceID As Long, _
    ByVal other As Variant)
    
    Dim i As Integer
    
    '如果报告编辑器中显示了报告图且事件类型为更新图像，则执行更新报告图像操作
    If mblnShowImage And (lngEventType = TVideoEventType.vetUpdateImg Or lngEventType = TVideoEventType.vetCaptureFirstImg Or lngEventType = TVideoEventType.vetDelAllImg) And lngAdviceID = mlngAdviceID Then
        Call RefPacsPic
        Exit Sub
    End If
    
    If Not mblnShowVideoCapture Then Exit Sub
    
    Select Case lngEventType
        Case TVideoEventType.vetLockStudy
            For i = 1 To dkpMain.PanesCount
                If dkpMain.Panes(i).Title Like "*视频采集*" Then
                    dkpMain.Panes(i).Title = "【" & other & "】视频采集"
                    Exit For
                End If
            Next i
        Case TVideoEventType.vetUnLockStudy
            For i = 1 To dkpMain.PanesCount
                If dkpMain.Panes(i).Title Like "*视频采集*" Then
                    dkpMain.Panes(i).Title = "视频采集"
                    Exit For
                End If
            Next i
'        Case TVideoEventType.vetUpdateImg
'            '更新图像
'            If lngAdviceID = mlngAdviceID Then
'                Call RefPacsPic
'            End If
    End Select
    
End Sub


'接口实现部分*********************************************************************************

Public Function IWorkMenu_zlIsModuleMenu(ByVal objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单
    IWorkMenu_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenu_zlCreateMenu(objMenuBar As Object)
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    Set mObjActiveMenuBar = objMenuBar
    
'    If Not HasMenu(objMenuBar, conMenu_EditPopup) Then
        '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
        '-----------------------------------------------------

        Set cbrMenuBar = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "报告", 3, False)
        cbrMenuBar.ID = conMenu_EditPopup
        cbrMenuBar.Category = ""
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Preview, "预览", "", 102, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Print, "打印", "", 103, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_BatPrint, "批量打印", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_Open, "书写", "", 3002, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_ClearWritingState, "清除状态", "", 21903, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Edit_Delete, "删除", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Open, "查阅", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_ExportToXML, "导出XML…", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Tool_Search, "报告检索…", "", 0, False)
        End With
'    End If
End Sub


Public Sub IWorkMenu_zlCreateToolBar(objToolBar As Object)
'创建工具栏
    Dim cbrControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim lngIndex As Long
    
    Set cbrLogOut = objToolBar.FindControl(, conMenu_Manage_Logout, , True)
    
    lngIndex = 4
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index

    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Preview, "预览", "报告预览", 102, True, lngIndex + 1)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Print, "打印", "报告打印", 103, False, lngIndex + 2)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_PacsReport_Open, "书写", "", 2607, False, lngIndex + 3) 'IconId=3002
End Sub


Public Sub IWorkMenu_zlClearMenu()
'清除所创建的菜单
    Exit Sub
End Sub


Public Sub IWorkMenu_zlClearToolBar()
'清除创建的工具栏
    Exit Sub
End Sub



Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
    Call cbrMain_Update(control)
End Sub

Public Sub IWorkMenu_zlExecuteMenu(ByVal lngMenuId As Long)
    Dim objControl As XtremeCommandBars.ICommandBarControl
    
    Set objControl = mObjActiveMenuBar.FindControl(, lngMenuId, , True)
    If objControl Is Nothing Then Exit Sub
    
    Call cbrMain_Execute(objControl)
End Sub


Public Sub IWorkMenu_zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单
    Exit Sub
End Sub

Public Sub IWorkMenu_zlRefreshSubMenu(objMenuBar As Object)
'刷新弹出的子菜单
    Exit Sub
End Sub

'*************************************************************************************************


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'创建该模块内的菜单
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '如果这里不指定id，则不能将有些菜单添加到右键菜单中
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = "" 'M_STR_MODULE_MENU_TAG
End Function


Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing, Optional blnSingleWindow As Boolean = False)
'初始化报告模块
    Dim blnRestoreWindow As Boolean
    
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngDeptID = lngDepartId
    
    If Not owner Is Nothing Then Set mobjOwner = owner

    
    mblnSingleWindow = blnSingleWindow
    
    blnRestoreWindow = IIf(mlngDeptID = 0, True, False)
    
    '初始子窗体
    If mfrmReportView Is Nothing Then Set mfrmReportView = New frmReportView      '报告所见
    If mfrmReportWord Is Nothing Then Set mfrmReportWord = New frmReportWord      '词句示范
    If mobjReport Is Nothing Then Set mobjReport = New zlRichEPR.cDockReport      '电子病历报告
    
    Call InitLoaclParas(mlngDeptID, mlngModule, mstrPrivs, mlngModule = G_LNG_PACSSTATION_MODULE)

    Call InitFaceScheme  '初始界面布局,跟科室相关
    
    Call subShowHistoryList
    
    '使RestoreWinState方法在该对象的生命周期只执行一次，否则造成嵌套的报告编辑器位置错位。
    If blnRestoreWindow Then Call RestoreWinState(Me, App.ProductName)
    
    '提取用户的词句示范权限，跟科室无关
    mintWordPower = zlGetWordPower
    
    '如果包含视频采集窗口，则在这里对视频采集窗口初始化
    If mblnShowVideoCapture Then
        If mblnUseActiveCapture Then
            Call InitActiveVideoModuleObj
        End If
    End If
End Sub


Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'同步检查医嘱信息
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    mblnRefreshState = True
    
    If mblnUseActiveCapture Then
        If Not mobjWork_ActiveVideo Is Nothing Then Call mobjWork_ActiveVideo.zlUpdateStudyInf(lngAdviceID, lngSendNO, lngStudyState, blnMoved)
    End If
End Sub


Public Function zlRefreshFace(Optional blnForceRefresh As Boolean = False) As Boolean
On Error GoTo ErrHandle
    Dim lngNewAdviceId As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngOldFileID As Long            '记录当前的文件ID，用来比较诊疗单据是否发生改变
    Dim blnPrinted As Boolean           '报告是否已经被打印
    Dim lngStudyState As Long           '检查的状态，“病人医嘱发送.执行过程”
    Dim str审核人 As String             '如果报告已经审核，查找最后的审核签名人
    Dim str随访描述  As String          '记录随访描述的内容
    Dim thisUserSignLevel As EPRSignLevelEnum   '当前用户的签名级别
    Dim arrReportFormat() As String
       
    
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then Exit Function
        
    lngNewAdviceId = mlngAdviceID
    mlngAdviceID = mlngTmpAdviceId
    
    '判断上一次的改动是否保存
    Call PromptModify
    
    '恢复医嘱ID
    mlngAdviceID = lngNewAdviceId
    
    mlngTmpAdviceId = lngNewAdviceId
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True
    
    
    With Me.cbrMain.Options
        If mblnSingleWindow = True Then
            .SetIconSize True, 24, 24
        Else
            .SetIconSize True, 16, 16
        End If
    End With
    
    lngOldFileID = mFileID
    mReportID = 0
    mFileID = 0
    mintReportViewType = -1
    mblnIsReportDelete = False
    mblnModified = False
    mblnReadOnly = True
    mstr医嘱附件 = ""
    blnPrinted = False
    mblnCanUntread = True
    mblnPrintView = False
    
    If mlngAdviceID <> 0 Then
        If mblnMoved = True Then    '被转储，则报告为只读
            mblnReadOnly = True
        Else
            '查询医嘱执行状态,和是否出院归档
            strSql = "Select a.执行过程,c.出院日期,c.病案状态 From 病人医嘱发送 a,病人医嘱记录 b,病案主页 c Where " _
                & " a.医嘱ID = b.Id And  b.病人ID = c.病人ID(+) And b.主页ID = c.主页ID(+) " _
                & " And a.医嘱ID= [1] "
                
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
            
            If rsTemp.EOF = False Then
                lngStudyState = Nvl(rsTemp!执行过程, 0)
                '已完成的报告，为只读状态
                mblnReadOnly = IIf(lngStudyState = 6 Or lngStudyState = 0, True, False)
                '出院且归档后，报告不可操作
                If mblnReadOnly = False Then mblnReadOnly = IIf(Nvl(rsTemp!出院日期) <> "" And Nvl(rsTemp!病案状态, 0) >= 1, True, False)
            End If
            
            '如果不是只读状态，再查询报告并发状态
            If mblnReadOnly = False Then
                If CheckConcurrentReport(Me, mlngAdviceID, True) = False Then
                    mblnReadOnly = True
                End If
            End If
        End If
        
        '查询病历文件ID
        strSql = "Select 病历ID From 病人医嘱报告 Where 医嘱ID= [1]"
        If mblnMoved = True Then
            strSql = Replace(strSql, "病人医嘱报告", "H病人医嘱报告")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
        
        If rsTemp.EOF = True Then
            '如果没有查到记录，说明病人还没有报告，需要根据诊疗项目创建报告
            strSql = "Select l.病人来源, a.病历文件id" & vbNewLine & _
                "From 病人医嘱记录 l, 病历单据应用 a" & vbNewLine & _
                "Where l.诊疗项目id = a.诊疗项目id(+) And a.应用场合(+) = Decode(l.病人来源, 2, 2, 4 ,4, 1) And l.Id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
            If rsTemp.EOF = True Then
                mFileID = 0
            Else
                mFileID = Nvl(rsTemp!病历文件id, 0)
                mintEditType = 0    '创建报告
            End If
            
            mReportID = 0
            mlngEPRDeptID = 0
            mstrEPR创建人 = UserInfo.姓名
            mstrEPR保存人 = UserInfo.姓名
            mlngEPR签名级别 = 0
            mdtReportTime = Now
        Else
            mReportID = Nvl(rsTemp!病历Id, 0)
            mintEditType = 1    '书写报告，在哪里修订报告呢？
            '查找设计格式的文件ID
            strSql = "Select 文件ID,科室ID,创建人,保存人,签名级别,保存时间 From 电子病历记录  Where Id =[1]"
            If mblnMoved = True Then
                strSql = Replace(strSql, "电子病历记录", "H电子病历记录")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
            
            mFileID = rsTemp!文件ID
            mlngEPRDeptID = rsTemp!科室ID
            mstrEPR创建人 = Nvl(rsTemp!创建人)
            mstrEPR保存人 = Nvl(rsTemp!保存人)
            mlngEPR签名级别 = Nvl(rsTemp!签名级别, 0)
            mdtReportTime = Nvl(rsTemp!保存时间, Now)
        End If
        
        '如果病历文件ID找不到，提示设置诊疗项目对应的病历文件
        If mFileID = 0 Then
            mlngAdviceID = 0
            '如果没有找到病历文件，则屏蔽相关功能，如果没有病历文件mlngAdviceID值就被修改为0
            mblnReadOnly = True
            mblnPrintView = True
            MsgBoxD Me, "未找到该诊疗项目对应的病历文件，请到诊疗项目管理中设置"
        ElseIf mFileID <> lngOldFileID Then '诊疗单据发生改变，需要改变诊疗单据对应的打印设置菜单
            mblnRefreshRptFormat = False
            mbln使用自定义报表 = False
            
            '当诊疗单据改变时，才清空选中的报表格式，第一次打开，使用原来设置的默认格式
            If lngOldFileID <> 0 Then
                mstr选中报表格式 = ""
            End If
    
            '先判断是否使用自定义报表
            strSql = "Select 通用,编号 From 病历文件列表  Where Id =[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取报告打印方式", mFileID)
            If rsTemp.EOF = False Then
                If Nvl(rsTemp!通用) = 2 Then
                    mbln使用自定义报表 = True     '使用自定义报表格式打印
                    If mstr报表编号 <> "ZLCISBILL" & Format(Nvl(rsTemp!编号), "00000") & "-2" Then
                        mstr选中报表格式 = ""
                    End If
                    
                    mstr报表编号 = "ZLCISBILL" & Format(Nvl(rsTemp!编号), "00000") & "-2"
                    mblnRefreshRptFormat = True
                Else
                    mbln使用自定义报表 = False    '使用编辑格式打印
                    '使用编辑格式打印，清空自定义报表格式的设置
                    mstr选中报表格式 = ""
                    mstr报表编号 = ""
                End If
            End If
        End If
        
        cbrMain.Item(2).Visible = True
        
        Call InitReportFormat        '初始化报告格式
        Call RefreshVersion     '刷新报告版本
        Call RefreshSigns       '刷新报告签名
        Call subShowHistoryList '填写报告历史
    Else    '医嘱ID为0
        cbrMain.Item(2).Visible = False
        '只读状态
        mblnReadOnly = True
    End If
    
    
    '从数据库查询随访描述和打印状态信息
    strSql = "Select 随访描述,报告打印 From 影像检查记录 Where 医嘱ID=[1] "
    If mblnMoved = True Then
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取随访记录和打印状态", mlngAdviceID)
    If rsTemp.EOF = False Then
        str随访描述 = Nvl(rsTemp!随访描述)
        blnPrinted = (Nvl(rsTemp!报告打印, 0) = 1)
    End If
    
    '根据打印和审核状态，确定本次报告是否可以书写
    '1、如果“审核打印后允许回退”=True，则本人可以回退，其他人可以继续修订
    '2、如果“审核打印后允许回退”=False，则已审核并且已打印的报告，只有本人可以修订，并且只能修订而不能回退，其他人为只读
    If blnPrinted And lngStudyState = 5 Then
        If mblnUntreadPrinted = False Then
            '需要先查找本次报告的审核人，最后保存人，不一定是审核人。
            '因为以下情况：A审核后，B也审核报告，然后B回退，此时保存人是B，但是审核人是A。之后再打印。
            
            strSql = "Select 要素表示 As 签名级别,内容文本 as 签名,开始版  From 电子病历内容 Where 文件ID=[1] " _
                            & " And 对象类型= 8 order by 开始版 "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取最后签名人", mReportID)
                    
            If rsTemp.EOF = False Then
                str审核人 = Split(Nvl(rsTemp!签名), ";")(0)
            End If
            
            If str审核人 <> UserInfo.姓名 Then
                mblnReadOnly = True
            Else
                '不处理只读状态，允许读写，但是不能回退
                mblnCanUntread = False
            End If
        End If
    End If
    
    '低级别的医生不能修订高级别医生的报告，打开报告后，报告为只读的。
    '这种情况只有在报告已经签名后再去考虑，所以签名级别<>0。修改后未签名的，在后续的chkEditState中处理。
    If m目标版本 > 1 And mlngEPR签名级别 <> 0 Then
        '自己书写的报告，应该是可以回退的
        '提取当前用户的签名级别
        thisUserSignLevel = GetUserSignLevel(UserInfo.ID)
        If thisUserSignLevel < mlngEPR签名级别 Then
            mblnReadOnly = True
        End If
    End If
    
    '判断报告是否可以编辑
    Call chkEditState(False)
    
    '判断报告是否可以打印
    If mblnCheckPrintPara = True Then
        Call chkPrintState
    Else
        mblnCanPrint = True
    End If
    
    '---------------------初始化各种窗体-----------------------------------------------------------
    
    Call ShowTitle(False)
    
    mblnExitAfterSign = IIf(Val(zlDatabase.GetPara("PACS报告签名后退出", glngSys, mlngModule, True, "0")) = 0, False, True)
    
    '初始化其他窗体
    '初始化报告内容编辑窗口
    strSql = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列"
    If mblnMoved = True Then
        strSql = Replace(strSql, "病人医嘱附件", "H病人医嘱附件")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取医嘱附件", mlngAdviceID)
    Do Until rsTemp.EOF
        mstr医嘱附件 = mstr医嘱附件 & rsTemp!项目 & ":" & Nvl(rsTemp!内容) & vbCrLf
        rsTemp.MoveNext
    Loop
    
    '初始化报告内容窗口
    mfrmReportView.txtReview.Text = str随访描述
    If InStr(mstrPrivs, "随访") > 0 Then
        mfrmReportView.txtReview.Enabled = True
    Else
        mfrmReportView.txtReview.Enabled = False
    End If
    
    mfrmReportView.zlRefresh mReportID, mblnSingleWindow, mFileID, True, mblnEditable, mstrModifyEdit, mstr医嘱内容 & vbCrLf & vbCrLf & mstr医嘱附件, mblnShowWord, mstrFormatInfo, mblnMoved
    
    '初始化报告图像窗口
    If mblnShowImage = True And (Not mfrmReportImage Is Nothing) Then
        mfrmReportImage.zlRefresh mlngAdviceID, mFileID, mReportID, mblnSingleWindow, mlngShowBigImg, mdblBigImgZoom, mintImageDblClick, mblnEditable, mblnMoved, mintMinImageCount, GetReportImageSelected, mlngModule
    End If
    '初始化专科报告窗口
    If mblnShowSpecial = True And (Not mfrmReportSpecial Is Nothing) Then
        If mstrSpecialForm <> Report_Form_frmReportCustom Then
            mfrmReportSpecial.zlRefresh Me, mlngAdviceID, mReportID, mblnSingleWindow, mblnEditable, mblnMoved
        Else
            mfrmReportSpecial.Refresh mlngAdviceID, mReportID, mblnEditable, mblnMoved
        End If
    End If

    '为了解决保存后上下左右键盘没有响应的问题
    If mblnShowImage = True And (Not mfrmReportImage Is Nothing) Then
        mfrmReportView.rtxtCheckView.Visible = False
        mfrmReportView.rtxtCheckView.Visible = True
    End If
    
    If mblnUseActiveCapture Then
        If Not mobjWork_ActiveVideo Is Nothing Then Call mobjWork_ActiveVideo.zlRefreshData
    End If
    
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Function


'Public Function zlRefresh(ByVal lngAdviceID As Long, _
'    frmParent As Form, ByVal blnMoved As Boolean, _
'    Optional strInfo As String) As Long
'
'    Dim strSql As String
'    Dim rsTemp As ADODB.Recordset
''    Dim blnDeptChanged As Boolean       '记录科室是否发生变化
'    Dim lngOldFileID As Long            '记录当前的文件ID，用来比较诊疗单据是否发生改变
'    Dim blnPrinted As Boolean           '报告是否已经被打印
'    Dim lngStudyState As Long           '检查的状态，“病人医嘱发送.执行过程”
'    Dim str审核人 As String             '如果报告已经审核，查找最后的审核签名人
'    Dim str随访描述  As String          '记录随访描述的内容
'    Dim thisUserSignLevel As EPRSignLevelEnum   '当前用户的签名级别
'
'    mblnMoved = blnMoved
''    Set mfrmParent = frmParent
'
'    On Error GoTo errHandle
'
'    With Me.cbrMain.Options
'        If mblnSingleWindow = True Then
'            .SetIconSize True, 24, 24
'        Else
'            .SetIconSize True, 16, 16
'        End If
'    End With
'
'    '判断上一次的改动是否保存
'    Call PromptModify
'
'    mblnIsReportDelete = False
'
'    lngOldFileID = mFileID
'
'    mlngAdviceID = lngAdviceID
''    mlngPetiCaptureAdviceID = lngAdviceID
'    mReportID = 0
'    mFileID = 0
'    mintReportViewType = -1
'    mblnModified = False
'    mblnReadOnly = True
'    mstr医嘱附件 = ""
'    blnPrinted = False
'    mblnCanUntread = True
'    mblnPrintView = False
'
'
'
'    If mlngAdviceID <> 0 Then
'        If mblnMoved = True Then    '被转储，则报告为只读
'            mblnReadOnly = True
'        Else
'            '查询医嘱执行状态,和是否出院归档
'            strSql = "Select a.执行过程,c.出院日期,c.病案状态 From 病人医嘱发送 a,病人医嘱记录 b,病案主页 c Where " _
'                & " a.医嘱ID = b.Id And  b.病人ID = c.病人ID(+) And b.主页ID = c.主页ID(+) " _
'                & " And a.医嘱ID= [1] "
'            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
'            If rsTemp.EOF = False Then
'                lngStudyState = Nvl(rsTemp!执行过程, 0)
'                '已完成的报告，为只读状态
'                mblnReadOnly = IIf(lngStudyState = 6 Or lngStudyState = 0, True, False)
'                '出院且归档后，报告不可操作
'                If mblnReadOnly = False Then mblnReadOnly = IIf(Nvl(rsTemp!出院日期) <> "" And Nvl(rsTemp!病案状态, 0) >= 1, True, False)
'            End If
'
'            '如果不是只读状态，再查询报告并发状态
'            If mblnReadOnly = False Then
'                If CheckConcurrentReport(Me, mlngAdviceID, True) = False Then
'                    mblnReadOnly = True
'                End If
'            End If
'        End If
'
'        '查询病历文件ID
'        strSql = "Select 病历ID From 病人医嘱报告 Where 医嘱ID= [1]"
'        If mblnMoved = True Then
'            strSql = Replace(strSql, "病人医嘱报告", "H病人医嘱报告")
'        End If
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
'        If rsTemp.EOF = True Then
'            '如果没有查到记录，说明病人还没有报告，需要根据诊疗项目创建报告
'            strSql = "Select l.病人来源, a.病历文件id" & vbNewLine & _
'                "From 病人医嘱记录 l, 病历单据应用 a" & vbNewLine & _
'                "Where l.诊疗项目id = a.诊疗项目id(+) And a.应用场合(+) = Decode(l.病人来源, 2, 2, 4 ,4, 1) And l.Id = [1]"
'            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
'            If rsTemp.EOF = True Then
'                mFileID = 0
'            Else
'                mFileID = Nvl(rsTemp!病历文件id, 0)
'                mintEditType = 0    '创建报告
'            End If
'            mReportID = 0
'            mlngEPRDeptID = 0
'            mstrEPR创建人 = UserInfo.姓名
'            mstrEPR保存人 = UserInfo.姓名
'            mlngEPR签名级别 = 0
'        Else
'            mReportID = Nvl(rsTemp!病历Id, 0)
'            mintEditType = 1    '书写报告，在哪里修订报告呢？
'            '查找设计格式的文件ID
'            strSql = "Select 文件ID,科室ID,创建人,保存人,签名级别 From 电子病历记录  Where Id =[1]"
'            If mblnMoved = True Then
'                strSql = Replace(strSql, "电子病历记录", "H电子病历记录")
'            End If
'            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
'            mFileID = rsTemp!文件ID
'            mlngEPRDeptID = rsTemp!科室ID
'            mstrEPR创建人 = Nvl(rsTemp!创建人)
'            mstrEPR保存人 = Nvl(rsTemp!保存人)
'            mlngEPR签名级别 = Nvl(rsTemp!签名级别, 0)
'        End If
'        '如果病历文件ID找不到，提示设置诊疗项目对应的病历文件
'        If mFileID = 0 Then
'            mlngAdviceID = 0
'            '如果没有找到病历文件，则屏蔽相关功能，如果没有病历文件mlngAdviceID值就被修改为0
'            mblnReadOnly = True
'            mblnPrintView = True
'            MsgBoxD Me, "未找到该诊疗项目对应的病历文件，请到诊疗项目管理中设置"
'        ElseIf mFileID <> lngOldFileID Then '诊疗单据发生改变，需要改变诊疗单据对应的打印设置菜单
'            mblnRefreshRptFormat = False
'            mbln使用自定义报表 = False
'            mstr报表编号 = ""
'            mstr选中报表格式 = ""
'
'            '先判断是否使用自定义报表
'            strSql = "Select 通用,编号 From 病历文件列表  Where Id =[1]"
'            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取报告打印方式", mFileID)
'            If rsTemp.EOF = False Then
'                If Nvl(rsTemp!通用) = 2 Then
'                    mbln使用自定义报表 = True     '使用自定义报表格式打印
'                    mstr报表编号 = "ZLCISBILL" & Format(Nvl(rsTemp!编号), "00000") & "-2"
'                    mblnRefreshRptFormat = True
'                Else
'                    mbln使用自定义报表 = False    '使用编辑格式打印
'                End If
'            End If
'        End If
'        cbrMain.Item(2).Visible = True
'
'        Call InitReportFormat        '初始化报告格式
'        Call RefreshVersion     '刷新报告版本
'        Call RefreshSigns       '刷新报告签名
'        Call subShowHistoryList '填写报告历史
'    Else    '医嘱ID为0
'        cbrMain.Item(2).Visible = False
'        '只读状态
'        mblnReadOnly = True
'    End If
'
'
'    '从数据库查询随访描述和打印状态信息
'    strSql = "Select 随访描述,报告打印 From 影像检查记录 Where 医嘱ID=[1] "
'    If mblnMoved = True Then
'        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
'    End If
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取随访记录和打印状态", mlngAdviceID)
'    If rsTemp.EOF = False Then
'        str随访描述 = Nvl(rsTemp!随访描述)
'        blnPrinted = (Nvl(rsTemp!报告打印, 0) = 1)
'    End If
'
'    '根据打印和审核状态，确定本次报告是否可以书写
'    '1、如果“审核打印后允许回退”=True，则本人可以回退，其他人可以继续修订
'    '2、如果“审核打印后允许回退”=False，则已审核并且已打印的报告，只有本人可以修订，并且只能修订而不能回退，其他人为只读
'    If blnPrinted And lngStudyState = 5 Then
'        If mblnUntreadPrinted = False Then
'            '需要先查找本次报告的审核人，最后保存人，不一定是审核人。
'            '因为以下情况：A审核后，B也审核报告，然后B回退，此时保存人是B，但是审核人是A。之后再打印。
'
'            strSql = "Select 要素表示 As 签名级别,内容文本 as 签名,开始版  From 电子病历内容 Where 文件ID=[1] " _
'                            & " And 对象类型= 8 order by 开始版 "
'            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取最后签名人", mReportID)
'
'            If rsTemp.EOF = False Then
'                str审核人 = Split(Nvl(rsTemp!签名), ";")(0)
'            End If
'
'            If str审核人 <> UserInfo.姓名 Then
'                mblnReadOnly = True
'            Else
'                '不处理只读状态，允许读写，但是不能回退
'                mblnCanUntread = False
'            End If
'        End If
'    End If
'
'    '低级别的医生不能修订高级别医生的报告，打开报告后，报告为只读的。
'    '这种情况只有在报告已经签名后再去考虑，所以签名级别<>0。修改后未签名的，在后续的chkEditState中处理。
'    If m目标版本 > 1 And mlngEPR签名级别 <> 0 Then
'        '自己书写的报告，应该是可以回退的
'        '提取当前用户的签名级别
'        thisUserSignLevel = GetUserSignLevel(UserInfo.ID)
'        If thisUserSignLevel < mlngEPR签名级别 Then
'            mblnReadOnly = True
'        End If
'    End If
'
'    '判断报告是否可以编辑
'    Call chkEditState(False)
'
'    '判断报告是否可以打印
'    If mblnCheckPrintPara = True Then
'        Call chkPrintState
'    Else
'        mblnCanPrint = True
'    End If
'
'    '---------------------初始化各种窗体-----------------------------------------------------------
'
'    Call ShowTitle(False)
'    mblnExitAfterSign = IIf(Val(zlDatabase.GetPara("PACS报告签名后退出", glngSys, mlngModule, True, "0")) = 0, False, True)
'
'    '初始化其他窗体
'    '初始化报告内容编辑窗口
'    strSql = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列"
'    If mblnMoved = True Then
'        strSql = Replace(strSql, "病人医嘱附件", "H病人医嘱附件")
'    End If
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取医嘱附件", mlngAdviceID)
'    Do Until rsTemp.EOF
'        mstr医嘱附件 = mstr医嘱附件 & rsTemp!项目 & ":" & Nvl(rsTemp!内容) & vbCrLf
'        rsTemp.MoveNext
'    Loop
'
'    '初始化报告内容窗口
'    mfrmReportView.txtReview.Text = str随访描述
'    If InStr(mstrPrivs, "随访") > 0 Then
'        mfrmReportView.txtReview.Enabled = True
'    Else
'        mfrmReportView.txtReview.Enabled = False
'    End If
'
'    mfrmReportView.zlRefresh mReportID, mblnSingleWindow, mFileID, True, mblnEditable, mstrModifyEdit, mstr医嘱内容 & vbCrLf & vbCrLf & mstr医嘱附件, mblnShowWord, mstrFormatInfo, mblnMoved
'
'    '初始化报告图像窗口
'    If mblnShowImage = True And (Not mfrmReportImage Is Nothing) Then
'        mfrmReportImage.zlRefresh mlngAdviceID, mFileID, mReportID, mblnSingleWindow, mlngShowBigImg, mdblBigImgZoom, mintImageDblClick, mblnEditable, mblnMoved, mintMinImageCount, GetReportImageSelected, mlngModule
'    End If
'    '初始化专科报告窗口
'    If mblnShowSpecial = True And (Not mfrmReportSpecial Is Nothing) Then
'        If mstrSpecialForm <> Report_Form_frmReportCustom Then
'            mfrmReportSpecial.zlRefresh Me, mlngAdviceID, mReportID, mblnSingleWindow, mblnEditable, mblnMoved
'        Else
'            mfrmReportSpecial.Refresh mlngAdviceID, mReportID, mblnEditable, mblnMoved
'        End If
'    End If
'
'    '为了解决保存后上下左右键盘没有响应的问题
'    If mblnShowImage = True And (Not mfrmReportImage Is Nothing) Then
'        mfrmReportView.rtxtCheckView.Visible = False
'        mfrmReportView.rtxtCheckView.Visible = True
'    End If
'
'    Exit Function
'
'errHandle:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'End Function


'Public Sub zlEditReport(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, _
'        ByVal frmParent As Object, ByVal blnMoved As Boolean, Optional strPatientName As String)
'    '记录当前编辑方式为独立窗口方式
'    mblnSingleWindow = True
'
'    Set mfrmParent = frmParent
'    Call Me.Show(, frmParent)
'
'    Call zlRefresh(lngAdviceID, frmParent, blnMoved, strPatientName)
'    RaiseEvent AfterOpen
'End Sub

Public Sub zlEditReport()
    '记录当前编辑方式为独立窗口方式
    mblnSingleWindow = True
    
    '使用该方法时，说明需要使用独立窗口打开报告编辑器，因此需要执行RestoreWinState方法恢复窗口位置
    Call RestoreWinState(Me, App.ProductName)
    
    Call Me.Show(, mobjOwner)
    
    Call zlRefreshFace
    
    RaiseEvent AfterOpen
End Sub


Private Sub chkEditState(blnShowMessage As Boolean)
    'blnShowMessage---在嵌入式的模式下，是否显示提示信息
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    mstrModifyEdit = ""
    
    If mblnReadOnly = True Then
        mblnEditable = False
        Exit Sub
    End If
    
    If m目标版本 = 1 And InStr(mstrPrivs, "PACS报告书写") > 0 Then
        If mstrEPR创建人 = UserInfo.姓名 Then
            mblnEditable = True
        ElseIf (InStr(mstrPrivs, "PACS他人报告") > 0 And mlngEPRDeptID = mlngDeptID) Then '有他人报告权限的，可以书写本科室的报告
            mblnEditable = True
        Else
            mblnEditable = False
            If mblnSingleWindow = True Or blnShowMessage = True Then  '独立窗口模式，或者嵌入式下需要提示，直接提示
                MsgBoxD Me, "本报告已经由" & mstrEPR创建人 & "正在填写，现在无权限修改。", vbOKOnly
            End If
        End If
    ElseIf m目标版本 > 1 And InStr(mstrPrivs, "PACS报告修订") > 0 Then
        '在报告修订的状态下，有报告修订权限的人，可以书写本科室的报告。
        If mstrEPR保存人 = UserInfo.姓名 Or mlngEPR签名级别 <> 0 Then   '报告最后是自己最后保存的，或者前面的修改者已经签名
            mblnEditable = True
        Else
            '已经有人在修订这个报告,修改已经保存，但是没有签名，报告不可编辑，记录修订人名称
            mstrModifyEdit = mstrEPR保存人
            mblnEditable = False
        End If
    Else
        mblnEditable = False
    End If
    
    If mblnTechReptSame And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        strSql = " select 检查技师 from 影像检查记录 where 医嘱id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
        
        If rsTemp.RecordCount < 1 Then Exit Sub
        
        If Nvl(rsTemp!检查技师) <> "" And Nvl(rsTemp!检查技师) <> UserInfo.姓名 Then
            mblnEditable = False
        Else
            mblnEditable = True
        End If
    
    End If
End Sub

Public Function chkModified() As Boolean
    
    '改变格式
    If mHasChangeFormat = True Then
        chkModified = True
        Exit Function
    End If
    
    '修改报告内容
    If mfrmReportView.pModified = True Then
        chkModified = True
        Exit Function
    End If
    
    '修改报告图或标记图
    If mblnShowImage = True And Not mfrmReportImage Is Nothing Then
        If mfrmReportImage.pMarkModified = True Or mfrmReportImage.pImageModified = True Then
            chkModified = True
            Exit Function
        End If
    End If
    
    '修改专科报告信息
    If mblnShowSpecial = True And Not mfrmReportSpecial Is Nothing Then
        If mfrmReportSpecial.pModified = True Then
            chkModified = True
            Exit Function
        End If
    End If
End Function


Private Sub RefreshSigns()
'------------------------------------------------
'功能：刷新签名对象，删除本次签名的对象，重新从数据库读取，确保签名对象的内容跟数据库的一致，签名回退刷新之后调用本过程
'参数： 无
'返回： 无
'-----------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim OneSign As cEPRSign
    Dim i As Integer
    Dim strSigns As String
    
    '清空原有签名
    For i = 1 To mSigns.Count
        mSigns.Remove 1
    Next i
    mSigns.UpdateMaxKey
    
    strSql = "Select Id,对象标记 From 电子病历内容 Where 文件id= [1] And 对象类型=8 Order By 对象标记"
    If mblnMoved = True Then
        strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
    While rsTemp.EOF = False
        Set OneSign = New cEPRSign
        If OneSign.GetSignFromDB(rsTemp!ID) = True Then
            OneSign.Key = Nvl(rsTemp!对象标记, 0)
            mSigns.AddExistNode OneSign, IIf(OneSign.Key = 0, False, True)
            strSigns = strSigns & " " & OneSign.前置文字 & OneSign.姓名
        End If
        rsTemp.MoveNext
    Wend
    
    '填写签名文本框
    mfrmReportView.txtSigns.Text = strSigns
End Sub

Private Sub RefreshVersion()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If mReportID = 0 Then
        '创建报告的情况下，最后版本=1和签名级别=0
        m最后版本 = 1
        m签名级别 = cprSL_空白
        m目标版本 = 1
    Else
        strSql = "Select 最后版本,签名级别 From 电子病历记录  Where Id =[1]"
        If mblnMoved = True Then
            strSql = Replace(strSql, "电子病历记录", "H电子病历记录")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
        m最后版本 = Nvl(rsTemp!最后版本, 1)
        m签名级别 = Nvl(rsTemp!签名级别, cprSL_空白)
        m目标版本 = m最后版本 + IIf(m签名级别 = cprSL_空白, 0, 1)
    End If
End Sub

Private Sub ShowTitle(blnChangeFormat As Boolean)
'blnChangeFormat 是否修改格式

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strName As String
    Dim strSex As String
    Dim strAge As String
    Dim lngStyle  As Long
    Dim strDoctor As String
    Dim strCheckNo As String
    Dim strAdvice As String
        
    On Error GoTo ErrHandle
    
    If blnChangeFormat = True Then  '更改格式
        If mFormatID = 0 Then
            strSql = "Select a.姓名,a.性别,a.年龄,a.检查号,b.医嘱内容 From 影像检查记录 a ,病人医嘱记录 b  Where a.医嘱ID= b.id and b.id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
            mModelName = "标准报告"
        Else
            strSql = "Select a.姓名,a.性别,a.年龄,a.检查号,b.名称,c.医嘱内容 From 影像检查记录 a,病历范文目录 b,病人医嘱记录 c  Where a.医嘱ID=c.id and c.id= [1] And b.Id =[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID, mFormatID)
            If rsTemp.EOF = False Then mModelName = rsTemp!名称
        End If
    Else
        If mReportID = 0 Then
            strSql = "Select a.姓名,a.性别,a.年龄,a.检查号,b.医嘱内容 From 影像检查记录 a ,病人医嘱记录 b  Where a.医嘱ID= b.id and b.id = [1]"
            If mblnMoved = True Then
                strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
                strSql = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
            mModelName = "标准报告"
        Else
            strSql = "Select a.姓名,a.性别,a.年龄,a.检查号,b.病历名称,b.保存人,b.完成时间,c.医嘱内容 From 影像检查记录 a,电子病历记录 b,病人医嘱记录 c  Where a.医嘱ID=c.id and c.id = [1] And b.Id =[2]"
            If mblnMoved = True Then
                strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
                strSql = Replace(strSql, "电子病历记录", "H电子病历记录")
                strSql = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID, mReportID)
            If rsTemp.EOF = False Then
                mModelName = rsTemp!病历名称
                If Nvl(rsTemp!完成时间) = "" Then
                    strDoctor = Nvl(rsTemp!保存人)
                End If
            End If
        End If
    End If

    If rsTemp.EOF = False Then
        strName = Nvl(rsTemp!姓名)
        strSex = Nvl(rsTemp!性别)
        strAge = Nvl(rsTemp!年龄)
        strCheckNo = Nvl(rsTemp!检查号)
        strAdvice = Nvl(rsTemp!医嘱内容)
        mstr医嘱内容 = strAdvice
    End If
    
    '没有隐藏标题，才更新标题栏
    lngStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
    If (lngStyle And WS_CAPTION) <> 0 Then
        Me.Caption = IIf(m目标版本 > 1, "[报告修订]", "[报告书写]") & "   【姓名：" & strName & " 性别：" & strSex & " 年龄：" & strAge & "】   报告医生：" & UserInfo.姓名 _
                     & " 检查号：" & strCheckNo & "   医嘱：" & strAdvice
    End If
    mstrFormatInfo = IIf(m目标版本 > 1, "[报告修订]", "[报告书写]") & "   " & mModelName
    If blnChangeFormat = False Then
        If mReportID = 0 Then
            mstrFormatInfo = mstrFormatInfo & " 新报告，还未开始书写"
        ElseIf strDoctor <> "" Then
            mstrFormatInfo = mstrFormatInfo & " " & strDoctor & " 正在书写报告"
        End If
    End If
    
    If mstr选中报表格式 <> "" Then
        If InStr(mstrFormatInfo, vbCrLf) <> 0 Then
            mstrFormatInfo = Left(mstrFormatInfo, InStr(mstrFormatInfo, vbCrLf) - 1)
        End If
        mstrFormatInfo = mstrFormatInfo & vbCrLf & "打印格式：" & mstr选中报表格式
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitFaceScheme()
'初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane, Pane5 As Pane
    Dim Pane6 As Pane
    Dim i As Integer
    Dim intPaneID As Integer
    
    '定义Pane的ID顺序： 1-检查所见；2-历史报告；3-词句示范；4-报告图；5-视频采集；6-专科报告。
    
    If mlngDeptID = 0 Then Exit Sub
    
    On Error Resume Next
    
    '设置总体显示策略
    With Me.dkpMain
        .VisualTheme = ThemeOffice2003
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .PanelPaintManager.BoldSelected = True
        .TabPaintManager.Position = xtpTabPositionLeft  'TAB放到左边显示
'        .TabPaintManager.OneNoteColors = True           '一个TAB一种颜色显示
        .TabPaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .TabPaintManager.BoldSelected = True
        dkpMain.Options.DefaultPaneOptions = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    End With
    
    '先从注册表读取预先设置好的窗口布局，然后再逐个设置
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmReport" & IIf(mblnSingleWindow = True, "\SingleWindow\", "\") & mlngModule & "\" & TypeName(dkpMain), _
                dkpMain.Name & mlngDeptID, "")
    End If
    
    
    
    '历史报告
    intPaneID = PaneHasShow("历史报告")
    If intPaneID = 0 Then
        '加载历史报告页面
        Set Pane1 = dkpMain.CreatePane(1, 300, 150, DockLeftOf)
        Pane1.Title = "历史报告" '历史报告
        'Pane1.Options = PaneNoCaption
    Else
        Set Pane1 = dkpMain.Panes(intPaneID)
    End If
    
    '检查所见
    intPaneID = PaneHasShow("检查所见")
    If intPaneID = 0 Then
        '加载检查所见页面
        Set Pane2 = dkpMain.CreatePane(2, 600, 150, DockRightOf, Pane1)
        Pane2.Title = "检查所见"
        Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable
    End If
    
    '词句示范
    intPaneID = PaneHasShow("词句示范")
    If intPaneID = 0 And mblnShowWord = True Then
        '加载词句示范页面
        Set Pane3 = dkpMain.CreatePane(3, 300, 150, DockTopOf, Pane1)
        Pane3.Title = "词句示范"
        'Pane3.Options = PaneNoCaption
        Pane3.AttachTo Pane1
        
    ElseIf intPaneID <> 0 And mblnShowWord = False Then
        '不用显示词句示范页面，卸载该页面
        Call dkpMain.DestroyPane(dkpMain.Panes(intPaneID))
    End If
    
    '报告图
    intPaneID = PaneHasShow("报告图")
    If intPaneID = 0 And mblnShowImage = True Then
        '加载报告图页面
        Set pane4 = dkpMain.CreatePane(4, 300, 150, DockTopOf, Pane1)
        pane4.Title = "报告图"
        'pane4.Options = PaneNoCaption
        pane4.AttachTo Pane1
    ElseIf intPaneID <> 0 And mblnShowImage = False Then
        '不用显示报告图页面，卸载该页面
        Call dkpMain.DestroyPane(dkpMain.Panes(intPaneID))
    End If
    
    '视频采集
    intPaneID = PaneHasShow("视频采集")
    If intPaneID = 0 And mblnShowVideoCapture = True Then
        '加载视频采集页面
        Set Pane5 = dkpMain.CreatePane(5, 300, 150, DockTopOf, Pane1)
        Pane5.Title = "视频采集"
        'Pane5.Options = PaneNoCaption
        Pane5.AttachTo Pane1
    ElseIf intPaneID <> 0 And mblnShowVideoCapture = False Then
        '不用显示视频采集页面，卸载该页面
        Call dkpMain.DestroyPane(dkpMain.Panes(intPaneID))
    End If
    
    '专科报告
    intPaneID = PaneHasShow("专科报告")
    If intPaneID = 0 And mblnShowSpecial = True Then
        '加载专科报告页面
        Set Pane6 = dkpMain.CreatePane(6, 300, 150, DockTopOf, Pane1)
        Pane6.Title = "专科报告"
        'Pane6.Options = PaneNoCaption
        Pane6.AttachTo Pane1
    ElseIf intPaneID <> 0 And mblnShowSpecial = False Then
        '不用显示专科报告页面，卸载该页面
        Call dkpMain.DestroyPane(dkpMain.Panes(intPaneID))
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '全部加载和显示完之后，再设置被选中的TAB
    If mintPaneID <= dkpMain.PanesCount Then
        Call dkpMain.Panes(mintPaneID).Select
    Else
        For i = 1 To dkpMain.PanesCount
            If dkpMain.Panes(i).Title = "词句示范" Then
                Call dkpMain.Panes(i).Select
            End If
        Next i
    End If
    
End Sub

Private Function PaneHasShow(strTitle As String) As Integer
'------------------------------------------------
'功能：查询DockingPane中的Pane是否已经显示
'参数： strTitle --- Pane的Title
'返回：如果找到Pane，返回Pane的ID，如果找不到，返回0
'------------------------------------------------
    Dim i As Integer
    
    For i = 1 To dkpMain.PanesCount
        If dkpMain.Panes(i).Title Like "*" & strTitle & "*" Then
            PaneHasShow = i
            Exit Function
        End If
    Next i

    PaneHasShow = 0
End Function

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim NewControl As XtremeCommandBars.CommandBarControl
    Dim strInfo As String
    
    If mblnMenuDownState Then Exit Sub
    
    mblnMenuDownState = True
    
    Select Case control.ID
        Case conMenu_PacsReport_Save        '保存报告
            Call SaveReport(True)
        Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_BatPrint       '打印报告,预览报告,批量打印
            '判断报告是否可以打印
            If mblnCheckPrintPara = True Then
                Call chkPrintState
            End If
            
            If PromptModify = False And mReportID = 0 Then
                    MsgBoxD Me, "新创建的报告，没有保存无法打印和预览，请先保存报告。"
                    mblnModified = True
            ElseIf mblnCanPrint = False Then
                MsgBoxD Me, "当前报告未审核，不能打印，请检查！", vbInformation, gstrSysName
            Else    '可以打印
                '打印前判断是否需要提示阴阳性和影像质量
                If control.ID = conMenu_File_Print And mlngHintType = 2 Then 'mlngHintType = 2表示打印前提醒
                    Dim strResultInput As String
                    
                    strResultInput = ""
                    If mblnReportWithResult Then '无影像诊断为阴性  -无提示自动标记
                        gstrSQL = "ZL_影像检查_结果(" & mlngAdviceID & ",0)"
                        zlDatabase.ExecuteProcedure gstrSQL, "标记阴阳性"
                    End If
                        
                    strSql = "Select B.危急状态, A.结果阳性, B.影像质量, B.报告质量, B.符合情况 " & _
                             "From 病人医嘱发送 A, 影像检查记录 B " & _
                             "Where A.医嘱id = B.医嘱id and B.医嘱ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取结果阳性", mlngAdviceID)
    
                    If IsNull(rsTemp!危急状态) And mintCriticalValues <> 0 Then strResultInput = "危急状态|"
                    If IsNull(rsTemp!结果阳性) And Not mblnIgnoreResult Then strResultInput = strResultInput & "结果阳性|"
                    If IsNull(rsTemp!影像质量) And mstrImageLevel <> "" Then strResultInput = strResultInput & "影像质量|"
                    If IsNull(rsTemp!报告质量) And mstrReportLevel <> "" Then strResultInput = strResultInput & "报告质量|"
                    If IsNull(rsTemp!符合情况) And mintConformDetermine <> 0 Then strResultInput = strResultInput & "符合情况|"
                        
                    If strResultInput <> "" Then Call PromptResult(mlngAdviceID, mlngModule, Me, mlngDeptID, strResultInput)
                End If
                    
                '打印报告或者预览报告
                If mbln使用自定义报表 = True Then
                    mblnPrintOK = False
                    Call subPrintReport(IIf(control.ID = conMenu_File_Preview, False, True), control.ID = conMenu_File_BatPrint)
                Else        '使用编辑模式打印，调用病历的打印过程
                    mobjReport.zlRefresh 0, 0, , , , mlngModule
                    mobjReport.zlRefresh mlngAdviceID, UserInfo.部门ID, , , mblnCanPrint, mlngModule
                    mblnPrintOK = False     '标记打印是否完成，在AfterPrinted事件中被设置成True
                    mobjReport.zlExecuteCommandBars control
                End If
                
                '打印后退出
                If mblnExitAfterPrint = True And control.ID = conMenu_File_Print _
                    And mblnSingleWindow = True And mblnPrintOK = True Then
                    Call PromptModify
                    Unload Me
                End If
            End If
        Case conMenu_Edit_Modify       '用病历编辑器打开报告
                mobjReport.zlRefresh 0, 0, , , , mlngModule
                If m目标版本 > 1 Then       '修订模式
                    Set NewControl = cbrMain.FindControl(, conMenu_Edit_Audit, False)
                    mobjReport.zlRefresh mlngAdviceID, UserInfo.部门ID, , , , mlngModule
                    mobjReport.zlExecuteCommandBars NewControl
                Else
                    mobjReport.zlRefresh mlngAdviceID, UserInfo.部门ID, , , , mlngModule
                    mobjReport.zlExecuteCommandBars control
                End If
        Case conMenu_File_Open, conMenu_File_ExportToXML, conMenu_Tool_Search      '查阅报告,导出XML，报告检索
            mobjReport.zlRefresh 0, 0, , , , mlngModule
            mobjReport.zlRefresh mlngAdviceID, UserInfo.部门ID, , , , mlngModule
            mobjReport.zlExecuteCommandBars control
        Case conMenu_PacsReport_Sign                        '签名
            Call AddSign
        Case conMenu_PacsReport_DelSign                     '回退
            Call DoUntread
        Case conMenu_PacsReport_Reject                      '驳回
            Call RejectReport
        Case conMenu_PacsReport_RejectHistory               '驳回历史
            Call ShowRejectHistory
        Case conMenu_PacsReport_VerifySign_Item             '签名验证
            Call FuncAdviceSignVerify(Val(control.parameter), mblnMoved)
        Case conMenu_PacsReport_SelFormat_Item              '选择格式
            Call ChangeFormat(Val(control.parameter))
        Case conMenu_PacsReport_RepFormat_Item              '选择打印格式
            control.Checked = Not control.Checked
            Call subChangeRptFormat(control)
        Case conMenu_PacsReport_FontSet                     '设置文本段字体
            Call SetTextFont
        Case conMenu_File_PrintSet                          '打印设置
            Call zlPrintSet
        Case conMenu_Edit_Delete                            '删除报告
            If mReportID = 0 Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            strInfo = "真的删除这份“" & mModelName & "”吗？"
            If MsgBoxD(Me, strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            strSql = "Zl_电子病历记录_Delete(" & mReportID & ")"
            
            mblnIsReportDelete = True
            
            
            zlDatabase.ExecuteProcedure strSql, Me.Caption
            
            err = 0: On Error GoTo 0
            RaiseEvent AfterDeleted(mlngAdviceID)
            
            Call Me.zlRefreshFace(True)
        Case conMenu_PacsReport_ClearWritingState           '清除报告“处理中”标记
            If mReportID = 0 Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            strInfo = "本报告的状态是有人正在处理中，确定要清除这份报告的状态吗？"
            If MsgBoxD(Me, strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            Call UpdateReporter(mlngAdviceID, "")
            
        Case conMenu_PacsReport_History                     '显示报告修订历史
            Call frmReportHistory.zlShowMe(Me, mlngAdviceID, mReportID)
        Case conMenu_PacsReport_SaveWord                    '保存词句示范
            Call subSaveWord(0)
        Case conMenu_PacsReport_AddNumber                   '给文本段添加序号
            Call AddNumber
        Case conMenu_PacsReport_PrivOrder                   '上一个医嘱
            Call ChangeOrder(1)
        Case conMenu_PacsReport_NextOrder                   '下一个医嘱
            Call ChangeOrder(2)
        
        Case comMenu_Petition_Capture                       '查看扫描单
            Call comMenu_Petition_扫描申请单
            
        Case conMenu_PacsReport_Default                     '重置默认界面
            Call ReStoreFace
            
        Case conMenu_File_Exit                              '   退出
            Call PromptModify
            Unload Me
        Case Else
        
    End Select
    
    mblnMenuDownState = False
    Exit Sub
ErrHandle:
    mblnMenuDownState = False
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ReStoreFace()
'删除注册表，重置默认界面
 On Error GoTo ErrHandle
 Dim strReportRegPath As String
 Dim strImageRegPath As String
 Dim strViewRegPath As String
 Dim strWordRegPath As String
 
 Dim strIndividualPath As String

    If mblnSingleWindow = True Then
        strReportRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
        strImageRegPath = "公共模块\" & App.ProductName & "\frmReportImage\SingleWindow"
        strViewRegPath = "公共模块\" & App.ProductName & "\frmReportView\SingleWindow"
        strWordRegPath = "公共模块\" & App.ProductName & "\frmReportWord\SingleWindow"
    Else
        strReportRegPath = "公共模块\" & App.ProductName & "\frmReport"
        strImageRegPath = "公共模块\" & App.ProductName & "\frmReportImage"
        strViewRegPath = "公共模块\" & App.ProductName & "\frmReportView"
        strWordRegPath = "公共模块\" & App.ProductName & "\frmReportWord"
        
        
    End If
 
    '关闭工作站 用于重置界面布局
    If MsgBoxD(Me, "恢复界面默认布局需要关闭工作站，是否继续？", vbYesNo, gstrSysName) = vbYes Then
        Unload mobjOwner
    Else
        Exit Sub
    End If
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        strIndividualPath = "公共模块\" & App.ProductName & "\frmReport\" & mlngModule & "\Dockingpane"
        
        Call DeleteSetting("ZLSOFT", strIndividualPath, "dkpMain" & mlngDeptID)
    End If
   
    Call DeleteSetting("ZLSOFT", strReportRegPath, "CX1")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "CX2")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "CX3")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "CY21")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "CY3")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "PicHistoryX")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "PicHistoryY")

    Call DeleteSetting("ZLSOFT", strImageRegPath, "CY1")
    Call DeleteSetting("ZLSOFT", strImageRegPath, "CY2")
    Call DeleteSetting("ZLSOFT", strImageRegPath, "CY3")
    Call DeleteSetting("ZLSOFT", strImageRegPath, "MarkW")
    Call DeleteSetting("ZLSOFT", strImageRegPath, "RptImgW")
    
    Call DeleteSetting("ZLSOFT", strViewRegPath, "CY1")
    Call DeleteSetting("ZLSOFT", strViewRegPath, "CY2")
    Call DeleteSetting("ZLSOFT", strViewRegPath, "CY3")
    Call DeleteSetting("ZLSOFT", strViewRegPath, "CY4")
    
    
    Call DeleteSetting("ZLSOFT", strWordRegPath, "PrivateWordH")
    Call DeleteSetting("ZLSOFT", strWordRegPath, "WordShowH")
    Call DeleteSetting("ZLSOFT", strWordRegPath, "WordTreeH")

    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub comMenu_Petition_扫描申请单()
'扫描申请单
On Error GoTo errFree
    Dim frmPetitionCap As New frmPetitionCapture
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    
    strSql = "select a.姓名,a.年龄,a.性别,a.医嘱内容,b.门诊号,b.住院号,c.名称 from 病人医嘱记录 a,病人信息 b,部门表 c where a.病人id = b.病人id and a.病人科室id = c.id and a.id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "得到病人信息", mlngAdviceID)
    
    If rsTemp.RecordCount = 0 Then
         MsgBoxD Me, "没有找到该病人相关记录", vbInformation, gstrSysName
         Exit Sub
    End If
    
    '打开扫描申请单窗口
    Call frmPetitionCap.ShowPetitionCaptureWind(mstrPrivs, _
                                                mlngDeptID, _
                                                Nvl(rsTemp!名称), _
                                                Nvl(rsTemp!姓名), _
                                                Nvl(rsTemp!年龄), _
                                                Nvl(rsTemp!性别), _
                                                Nvl(Mid(rsTemp!医嘱内容, 1, InStr(rsTemp!医嘱内容, ":") - 1)), _
                                                Nvl(Mid(rsTemp!医嘱内容, InStr(rsTemp!医嘱内容, ":") + 1, Len(rsTemp!医嘱内容))), True, False, _
                                                mlngAdviceID)
    
errFree:
    Call Unload(frmPetitionCap)
    Set frmPetitionCap = Nothing
End Sub

Private Sub SetTextFont()
    'mintReportViewType 0-检查所见CheckView，1-诊断意见Result，2-建议Advice
    Dim rText As RichTextBox
    
    If mintReportViewType < 0 Or mintReportViewType > 2 Then Exit Sub
    
    '判断是哪个文本段被选中,读取文本段的对象属性
    If mintReportViewType = 0 Then
        Set rText = mfrmReportView.rtxtCheckView
    ElseIf mintReportViewType = 1 Then
        Set rText = mfrmReportView.rtxtResult
    ElseIf mintReportViewType = 2 Then
        Set rText = mfrmReportView.rTxtAdvice
    End If
    
    rText.SelStart = 0
    rText.SelLength = Len(rText.Text)
    
    On Error GoTo dlgErr
    '显示字体选择窗口

    dlgFont.flags = cdlCFBoth
    dlgFont.CancelError = False  '把点取消当作错误处理
    dlgFont.FontName = Nvl(rText.SelFontName, "宋体")
    dlgFont.FontBold = Nvl(rText.SelBold, "False")
    dlgFont.FontItalic = Nvl(rText.SelItalic, "False")
    dlgFont.FontSize = Nvl(rText.SelFontSize, "14")
    dlgFont.ShowFont
    On Error GoTo 0
    '设置字体
    rText.SelFontName = dlgFont.FontName
    rText.SelFontSize = dlgFont.FontSize
    rText.SelBold = dlgFont.FontBold
    rText.SelItalic = dlgFont.FontItalic
    '修改该文本框的TAG
    RefreshViewTag rText
    Exit Sub
dlgErr:
    
End Sub

Private Sub RefreshViewTag(rText As RichTextBox)
    Dim strItem() As String
    Dim i As Integer
    
    '修改该文本框的TAG,如果TAG为空，则暂时不记录
    If rText.Tag <> "" Then
        strItem = Split(rText.Tag, "|")
        rText.Tag = ""
        strItem(15) = Nvl(rText.SelFontName, "宋体")     'FontName
        strItem(16) = Nvl(rText.SelFontSize, 14)    'FontSize
        strItem(17) = Nvl(rText.SelBold, "False")    'FontBold
        strItem(18) = Nvl(rText.SelItalic, "False")    'FontItalic
        For i = 0 To 25
            rText.Tag = rText.Tag & strItem(i) & "|"
        Next i
    End If
End Sub

Private Sub DoUntread()
'回退，回退签名和修订
    Dim lngVersion As Long
    Dim lngSignKey As Long
    Dim objESign As Object  '电子签名接口部件
    Dim strSql As String
    
    If mSigns.Count = 1 Then  '只有一个签名，表示当前是书写模式下的回退
        If frmEPRUntread.ShowMe(mReportID, cprET_单病历编辑, lngVersion, lngSignKey, Me) = False Then Exit Sub
    Else
        If frmEPRUntread.ShowMe(mReportID, cprET_单病历审核, lngVersion, lngSignKey, Me) = False Then Exit Sub
    End If
    If lngSignKey > 0 Or lngVersion > 0 Then
        If MsgBoxD(Me, "注意：回退操作将不可恢复！是否继续？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    End If
    
    '处理两种回退方式
    If lngSignKey > 0 Then
        '先处理数字签名，然后再清除签名
        If mSigns("K" & lngSignKey).签名方式 = 2 Then
            '数字签名验证
            err.Clear: On Error Resume Next
            If objESign Is Nothing Then
                Set objESign = CreateObject("zl9ESign.clsESign")
                If err <> 0 Then err = 0
            End If
            If Not objESign Is Nothing Then
                If objESign.Initialize(gcnOracle, glngSys) Then
                    If Not objESign.CheckCertificate(gstrDBUser) Then
                        MsgBoxD Me, "取消已签名文件时需要再次验证数字证书，但数字证书和当前用户不一致，请使用正确的数字证书回退报告。", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    MsgBoxD Me, "数字证书初始化失败，请使用正确的数字证书。", vbInformation + vbOKOnly, "书写签名"
                    Exit Sub
                End If
            Else
                MsgBoxD Me, "数字签名部件初始化失败！", vbOKOnly + vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '清除签名,并保存格式
        SaveReportFormat mSigns("K" & lngSignKey), False
        mSigns.Remove "K" & lngSignKey
        
    ElseIf lngVersion > 1 Then  '回退修订
        '直接修改数据库内容就可以了  '把回退修订保存到数据库
        strSql = "ZL_影像报告回退(0," & mReportID & "," & lngVersion & ")"
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        
        Call chkEditState(False)
        '刷新各个文本窗体，图像窗体不需要刷新
        mfrmReportView.zlRefresh mReportID, mblnSingleWindow, mFileID, False, mblnEditable, mstrModifyEdit, mstr医嘱内容 & vbCrLf & vbCrLf & mstr医嘱附件, mblnShowWord, mstrFormatInfo, mblnMoved
        If mblnShowSpecial = True Then
            If mstrSpecialForm <> Report_Form_frmReportCustom Then
                mfrmReportSpecial.zlRefresh Me, mlngAdviceID, mReportID, mblnSingleWindow, mblnEditable, mblnMoved
            Else
                mfrmReportSpecial.Refresh mlngAdviceID, mReportID, mblnEditable, mblnMoved
            End If
        End If
    End If

    Call RefreshVersion
    Call RefreshSigns
    '更新主界面 标题栏
    Call ShowTitle(False)
    '恢复修改标记
    Call subSetModifyFlag(False)
    
    Call AfterReportSaved(mlngAdviceID, 3)
    
    If mblnSingleWindow = True Then
        '对于弹出窗口，刷新窗口内容
        Call zlRefreshFace(True)
    End If
End Sub

Private Sub SaveReport(blnRaiseEvent As Boolean)
'------------------------------------------------
'功能：保存报告，保存报告格式，内容，但是不处理签名
'参数：     blnRaiseEvent -- 是否触发报告保存完成的事件，True-触发事件；False-不触发事件，当签名之前调用时，应该不触发事件，在签名的过程中单独触发
'返回： 无
'-----------------------------------------------
    Dim lngSaveAdviceID As Long '记录当前的医嘱ID，在报告保存的过程中，可能医嘱ID会被从外部改变
    
    On Error GoTo err
    '判断报告文本段长度是否超过2000个字符，如果超过，则提示，并退出
    If Len(mfrmReportView.rtxtCheckView.Text) > 2000 Or Len(mfrmReportView.rtxtResult.Text) > 2000 _
        Or Len(mfrmReportView.rTxtAdvice.Text) > 2000 Then
        MsgBoxD Me, "报告中检查所见、诊断意见或者建议的字数超过2000，请删减部分文字后再保存。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lngSaveAdviceID = mlngAdviceID
    
    If mHasChangeFormat = True Then     '更改了格式，要根据格式ID，重新创建报告
        If mFormatID = 0 Then
            Call SaveReportItems(True, 0)
        Else
            Call SaveReportItems(True, 1)
        End If
        mHasChangeFormat = False
    Else
        If mReportID = 0 Then    '创建报告
            Call SaveReportItems(True, 0)
        Else
            Call SaveReportItems(False, 0)
        End If
    End If
    mModified = False
    Call RefreshVersion
'    '更新主界面 标题栏
    Call ShowTitle(False)
    '恢复修改标记
    Call subSetModifyFlag(False)
    
    '清空报告操作人
    Call UpdateReporter(lngSaveAdviceID, "")
    
    '触发报告保存完成的事件
    If blnRaiseEvent Then Call AfterReportSaved(lngSaveAdviceID, 0)
    
    If mblnSingleWindow = True And blnRaiseEvent Then
        '对于弹出窗口，如果触发保存事件，则刷新窗口内容
        If mblnExitAfterSign = False Then Call zlRefreshFace(True)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AddSign()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngPatientID As Long
    Dim lngPageID As Long
    Dim strZipFile As String
    Dim strTemp As String
    Dim OneSign As cEPRSign
    Dim lngKey As Long
    Dim lngMaxSignLevel As Long
    Dim int开始版  As Integer   '本次报告签名的开始版
    Dim lngSaveType As Long
    
    '先保存报告,但是不触发报告保存完成的事件,然后再处理签名，签名之后触发报告保存完成的事件
    Call SaveReport(False)
        
    '查询病人ID和主页id
    strSql = "Select 病人id,主页id From  病人医嘱记录 Where id= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    lngPatientID = Nvl(rsTemp!病人ID, 0)
    lngPageID = Nvl(rsTemp!主页ID, 0)
    
    '获取最大签名级别
    strSql = "Select 要素表示 As 签名级别,内容文本 as 签名,开始版  From 电子病历内容 Where 文件ID=[1] " _
                        & " And 对象类型= 8 order by 签名级别 desc "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取最大签名级别", mReportID)
    If rsTemp.EOF = False Then
        lngMaxSignLevel = Nvl(rsTemp!签名级别, 0)
    End If
    
    '计算本次签名的开始版，处理签名版本
    If (mModified Or (mintEditType = 2 And m签名级别 = cprSL_空白)) Or (mintEditType = 1 Or mintEditType = 0) Then
        int开始版 = m目标版本
    Else
        int开始版 = m目标版本 - 1
    End If
    If int开始版 > 16 Then
        MsgBoxD Me, "目前系统支持的最大版本号为16，请回退或者重新整理！", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    
    '调用签名窗口
    Set OneSign = frmEPRSign.ShowMe(Me, mReportID, lngPatientID, lngPageID, mstrPrivs, lngMaxSignLevel, int开始版)
    
    lngSaveType = 0
    If Not OneSign Is Nothing Then
        '签名了，先判断一下，是否第二次诊断签名，如果是则提示是否确定要签名
        If OneSign.签名级别 = cprSL_经治 Then
            If mSigns.Count >= 1 Then
                If MsgBoxD(Me, "本次报告已经有签名了，是否还要再次签名？", vbOKCancel, "诊断签名重复") = vbCancel Then
                    Exit Sub
                End If
            End If
        End If
        
        '签名了，保存报告内容和签名
        lngKey = mSigns.AddExistNode(OneSign)
        
        '签名直接调用SaveReportFormat就可以了
        Call SaveReportFormat(mSigns("K" & lngKey), True)
        
        '刷新签名对象，确保跟数据库的一致
        Call RefreshSigns
        
        Call UpdateReporter(mlngAdviceID, "")
        
        lngSaveType = IIf(OneSign.签名级别 < cprSL_主治, 1, 2)
    End If
    
    
    '不管是否确认进行签名，都要触发报告保存完成的事件
    '触发报告保存完成的事件
    Call AfterReportSaved(mlngAdviceID, lngSaveType)
    
    '如果签名成功，而且设置了签名后退出，则卸载报告窗体
    If Not OneSign Is Nothing And mblnExitAfterSign = True And mblnSingleWindow = True Then
        Unload Me
    ElseIf mblnExitAfterSign = False And mblnSingleWindow = True Then
        Call zlRefreshFace(True)
    End If
End Sub

Private Sub RejectReport()
'驳回报告
Dim frmRj As frmReject
Dim i As Long
Dim lngAdviceColIndex As Long
Dim lngProcedureColIndex As Long
Dim lngRowIndex As Long
    
On Error GoTo errFree
    If mReportID <= 0 Then
        MsgBoxD Me, "当前检查没有报告，不能进行驳回操作。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Set frmRj = New frmReject
    
    Call frmRj.ShowRejectWindow(mlngAdviceID, mReportID, Me)
    
    If frmRj.IsOk Then
        Call SendMsgToMainWindow(Me, wetRejectReport, mlngAdviceID)
    End If
errFree:
    Unload frmRj
    Set frmRj = Nothing
End Sub


Private Sub ShowRejectHistory()
'显示驳回历史
Dim frmRj As frmReject
    
On Error GoTo errFree
    If mReportID <= 0 Then
        MsgBoxD Me, "当前检查没有报告，不存在驳回历史记录。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Set frmRj = New frmReject
    
    Call frmRj.ShowRejectHistory(mlngAdviceID, mReportID, Me)
errFree:
    Unload frmRj
    Set frmRj = Nothing
End Sub


Private Sub ChangeFormat(lngFormatID As Long)
    mFormatID = lngFormatID
    mHasChangeFormat = True
    '更新界面显示
    If mblnShowImage = True Then
        mfrmReportImage.zlChangeFormat lngFormatID
    End If
    '更新主界面 标题栏
    Call ShowTitle(True)
End Sub

Private Sub SaveReportItems(blnCreate As Boolean, iAction As Integer)
 ' iAction = 0   '从病历文件列表创建报告, iType = 1    '从病历范文目录创建报告

    If blnCreate = True Then        '创建报告
        CreateReport iAction
    End If
    
    '保存报告内容
    Call SaveReportView
    '保存标记图标记
    Call SavePicMarks(blnCreate)
    '保存报告图
    If mblnShowImage = True Then
       If mfrmReportImage.ImageCount > 0 Then
            '处理新报告，未点击报告图，保存出错的问题
            If mfrmReportImage.pImageModified = False And blnCreate = True Then
                mfrmReportImage.zlRefresh mlngAdviceID, mFileID, mReportID, mblnSingleWindow, mlngShowBigImg, mdblBigImgZoom, mintImageDblClick, mblnEditable, mblnMoved, mintMinImageCount, True, mlngModule
            End If
            
            If mfrmReportImage.pImageModified = True Or blnCreate = True Then
                Call SaveReportImages(blnCreate)
            End If
        End If
    End If
    '保存报告格式,签名对象传空，表示只保存格式，不处理签名
    Call SaveReportFormat(Nothing, True)
End Sub

Private Function CreateReport(iType As Integer) As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    ' iType = 0   '从病历文件列表创建报告, iType = 1    '从病历范文目录创建报告

    '创建电子病历内容
    strSql = "ZL_影像报告内容_创建(" & mlngAdviceID & "," & mFileID & "," & mFormatID & "," & iType & ")"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    '新创建的报告，从数据库中读取报告内容ID
    strSql = "Select 病历ID From 病人医嘱报告 Where 医嘱ID= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    If rsTemp.EOF = True Then
        MsgBoxD Me, "病历创建不正确，无法查找到病历内容ID"
        CreateReport = False
        Exit Function
    Else
        mReportID = rsTemp!病历Id
    End If
                
End Function

Private Sub SaveReportImages(blnCreate As Boolean)
    Dim lngImgTableID  As Long
    Dim strTabIDs() As String
    Dim iImgCount As Integer
    Dim strSql  As String
    Dim rsTemp As ADODB.Recordset
    Dim strTempFile As String
    Dim i As Integer
    Dim j As Integer
    Dim strPicAttrs As String
    Dim cFTP As New clsFtp
    Dim strFTPUser As String
    Dim strFTPPwd As String
    Dim strFtpIp As String
    Dim strFTPDirUrl As String
    Dim strSaveDeviceID As String
    Dim strBufferDir As String
    Dim strLocalDir As String
    
    On Error GoTo ErrHandle
    
    If mfrmReportImage.dcmReportImage.Count <= 1 Then Exit Sub
    
    strTempFile = App.Path & "\Temp.jpg"
    '获取表格ID串
    If blnCreate = True Then
        strSql = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
            " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
        If rsTemp.RecordCount > 0 Then
            ReDim strTabIDs(rsTemp.RecordCount - 1) As String
            For i = 0 To rsTemp.RecordCount - 1
                strTabIDs(i) = rsTemp!表格ID
                If i = 0 Then
                    mfrmReportImage.pTableID = rsTemp!表格ID
                Else
                    mfrmReportImage.pTableID = mfrmReportImage.pTableID & ";" & rsTemp!表格ID
                End If
                rsTemp.MoveNext
            Next i
        End If
    Else
        strTabIDs = Split(mfrmReportImage.pTableID, ";")
    End If
    
    '先判断数组是否为空
    If SafeArrayGetDim(strTabIDs) <> 0 Then
        '读取保存报告图的FTP信息
        strBufferDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
        
        strSql = "Select 位置一,位置二,检查UID,接收日期 From 影像检查记录 Where 医嘱ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
        If rsTemp.RecordCount <> 0 Then
            strSaveDeviceID = Nvl(rsTemp!位置一)
            If strSaveDeviceID = "" Then
                strSaveDeviceID = Nvl(rsTemp!位置二)
            End If
            strLocalDir = Format(Nvl(rsTemp!接收日期), "yyyyMMdd") & "/" & Nvl(rsTemp!检查UID)
            
            Call funGetStorageDevice(Me, strSaveDeviceID, strFTPDirUrl, strFtpIp, strFTPUser, strFTPPwd)
            cFTP.FuncFtpConnect strFtpIp, strFTPUser, strFTPPwd
        End If
        
        
        '分析和保存每一个图像表格
        For i = 0 To UBound(strTabIDs)
            lngImgTableID = Val(strTabIDs(i))
            iImgCount = mfrmReportImage.dcmReportImage(i + 1).Images.Count
            strPicAttrs = ""
            
            For j = 1 To iImgCount
                strPicAttrs = strPicAttrs & ";" & mfrmReportImage.dcmReportImage(i + 1).Images(j).Tag & "," & mlngAdviceID
            Next j
            
            strSql = "ZL_影像报告图像_保存(" & lngImgTableID & ",'" & strPicAttrs & "')"
            zlDatabase.ExecuteProcedure strSql, Me.Caption
            
            '保存报告图文件到FTP目录中
            For j = 1 To mfrmReportImage.dcmReportImage(i + 1).Images.Count
                mfrmReportImage.dcmReportImage(i + 1).Images(j).FileExport strBufferDir & "\" & strLocalDir & "\" & mfrmReportImage.dcmReportImage(i + 1).Images(j).Tag, "JPG"
                Call cFTP.FuncUploadFile(strFTPDirUrl & strLocalDir & "/", strBufferDir & "\" & strLocalDir & "\" & mfrmReportImage.dcmReportImage(i + 1).Images(j).Tag, mfrmReportImage.dcmReportImage(i + 1).Images(j).Tag)
            Next j
            
        Next i
    End If
    
    cFTP.FuncFtpDisConnect
    
    mfrmReportImage.pImageModified = False
    Exit Sub
ErrHandle:
    cFTP.FuncFtpDisConnect
    
    If ErrCenter() = 1 Then Resume Next
End Sub

Private Sub SavePicMarks(blnCreate As Boolean)
    Dim strMarks As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngMarkImageID As Long
    Dim i As Integer
    
    If mfrmReportImage Is Nothing Then Exit Sub
    
    If mfrmReportImage.pobjMarks Is Nothing Then
        mfrmReportImage.pMarkModified = False
        Exit Sub
    End If
    '创建标记文本
    For i = 1 To mfrmReportImage.pobjMarks.Count
        If i = 1 Then
            strMarks = mfrmReportImage.pobjMarks(i).对象属性
        Else
            strMarks = strMarks & "||" & mfrmReportImage.pobjMarks(i).对象属性
        End If
    Next i

    If blnCreate = True Then
        '新创建的报告，从电子病历内容中读取标记图ID
        strSql = "Select Id From 电子病历内容 Where 文件ID=[1] And  对象类型= 5 And substr(对象属性,1,1)='1' "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
        If rsTemp.EOF = False Then  '有标记图
            lngMarkImageID = rsTemp!ID
        Else    '没有标记图
            lngMarkImageID = 0
        End If
        
        mfrmReportImage.pMarkImageID = lngMarkImageID
    Else
        lngMarkImageID = mfrmReportImage.pMarkImageID
    End If
    
    strSql = "ZL_影像报告标注_保存(" & lngMarkImageID & ",'" & strMarks & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    mfrmReportImage.pMarkModified = False
End Sub

Private Sub SaveReportFormat(OneSign As cEPRSign, blnAddSign As Boolean)
'------------------------------------------------
'功能：保存报告格式RTF文件，对报告进行签名或者回退
'参数：     OneSign -- 不为空，则表示进行签名或者回退；为空，表示只是保存格式，不处理签名
'           blnAddSign 增加或者回退签名，True--增加签名,OneSign为空表示保存报告格式；False--回退签名
'返回： 无，直接保存RTF报告格式文档，对报告签名或者回退
'-----------------------------------------------
    Dim strZipFile As String
    Dim strTemp As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String
    Dim lngSignPos As Long
    
    '先复制报告格式
    If Dir(App.Path & "\ReportTemp") <> "" Then
        Kill App.Path & "\ReportTemp"
    End If
    '从数据库读取RTF报告格式文档
    strZipFile = zlBlobRead(5, mReportID, App.Path & "\ReportTemp")
    '解压缩文件
    strTemp = zlFileUnzip(strZipFile)
    
    If strTemp <> "" Then
        If blnAddSign = True Then
            '解析文件，根据报告内容，修改其中要素内容
            '读取RTF文件内容
            rtxtSaveElement.Filename = strTemp
            strReport = rtxtSaveElement.TextRTF
            
            '读取数据库中的要素，把各个要素内容填写到格式中
            strSql = "Select 对象标记,内容文本,要素名称 From 电子病历内容 Where 文件ID= [1] And 对象类型 = 4 And 终止版=0 and 保留对象 =0 order by 对象标记 "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
            While (rsTemp.EOF = False)
                ReplaceElement strReport, "E", rsTemp!对象标记, Nvl(rsTemp!内容文本, " ")
                rsTemp.MoveNext
            Wend
            
            '保存RTF文件
            rtxtSaveElement.TextRTF = strReport
            rtxtSaveElement.SaveFile strTemp
        End If
        
        '如果有签名，则保存签名
        If Not OneSign Is Nothing Then
            edtEditor.OpenDoc strTemp
            If blnAddSign = True Then   '增加签名
                '查找写入签名的位置
                strSql = "Select 对象标记 From 电子病历内容 Where 文件ID= [1] And 对象类型 = 4 And 要素名称 ='报告签名' "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
                lngSignPos = -1
                If rsTemp.EOF = False Then
                    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    
                    bFinded = FindKey(edtEditor, "E", Nvl(rsTemp!对象标记, 0), lKSS, lKSE, lKES, lKEE, bNeeded)
                    If bFinded = True Then lngSignPos = lKEE
                End If
                
                '向指定位置写入签名
                OneSign.InsertIntoEditor edtEditor, lngSignPos
                
                '把签名保存到数据库
                strSql = "ZL_影像报告签名_保存(" & mReportID & "," & OneSign.开始版 & "," & OneSign.终止版 & " ,'" & OneSign.对象属性 & "','" & OneSign.姓名 & _
                        "','" & OneSign.前置文字 & "','" & OneSign.时间戳 & "'," & OneSign.签名级别 & ",'" & OneSign.签名信息 & "')"
                zlDatabase.ExecuteProcedure strSql, Me.Caption
            Else    '回退签名
                OneSign.DeleteFromEditor edtEditor
                
                '把回退签名保存到数据库
                strSql = "ZL_影像报告回退(" & OneSign.ID & "," & mReportID & ",0)"
                zlDatabase.ExecuteProcedure strSql, Me.Caption
            End If
            
            '保存成临时文件
            edtEditor.SaveDoc strTemp
        End If
        
        '压缩文件
        strZipFile = zlFileZip(strTemp)
        '保存格式
        zlBlobSave 5, mReportID, strZipFile
    
        '删除临时zip文件
        Kill strZipFile
    Else
        MsgBoxD Me, "无法读取或者解压报告格式，请使用“病历编辑”的方法来编辑此报告。"
    End If
End Sub

Private Function ReplaceElement(strReport As String, strKeyType As String, lngKey As Long, strElement As String) As Boolean
    Dim sTMP As String
    Dim i As Long
    Dim j As Long
    Dim lLength As Long
    Dim strNewReport As String
    Dim lngES As Long
    Dim lngEE As Long
    Dim strChar As String
    Dim lulWave As Long
    Dim lulNone As Long
    
    sTMP = strKeyType & "S(" & Format(lngKey, "00000000")
    i = 1
LL1:
    i = InStr(i, strReport, sTMP)
    If i <> 0 Then
        '看是否关键字，若为关键字，必须是隐藏且受保护的。
        If ProtectAndHide(strReport, i - 1, i) = False Then
            i = i + 1
            GoTo LL1
        End If
        '已经找到起始关键字，往后查找字符，并替换这些字符
        j = i + 16
        lngES = j
        '查找结束关键字
        sTMP = strKeyType & "E(" & Format(lngKey, "00000000")
LL2:
        j = InStr(j, strReport, sTMP)
        If j <> 0 Then
            '看是否关键字，若为关键字，必须是隐藏且受保护的。
            If ProtectAndHide(strReport, j - 1, j) = False Then
                j = j + 1
                GoTo LL2
            End If
            lngEE = j - 1
            '已经找到结束关键字，说明中间就是需要替换的要素
            
            '过滤掉控制符号，\cfN,\highlightN,\v0
            If getElementPos(strReport, lngES, lLength, lngEE, lulWave, lulNone) = True Then
                strNewReport = strReport
                '先处理下划波浪，删除下划波浪的两个标记
                If lulWave <> 0 And lulNone <> 0 Then
                    strNewReport = Left(strNewReport, lulNone) & Right(strNewReport, Len(strNewReport) - lulNone - 7)
                End If
                '再处理要素内容，替换要素内容
                strChar = Mid(strElement, 1, 1)
                If (strChar >= "A" And strChar <= "Z") Or (strChar >= "a" And strChar <= "z") Or IsNumeric(strChar) Or strChar = " " Then
                    strNewReport = Left(strNewReport, lngES) & " " & StrToASC(strElement) & Right(strNewReport, Len(strNewReport) - lngES - lLength)
                Else
                    strNewReport = Left(strNewReport, lngES) & StrToASC(strElement) & Right(strNewReport, Len(strNewReport) - lngES - lLength)
                End If
                If lulWave <> 0 And lulNone <> 0 Then
                    strNewReport = Left(strNewReport, lulWave) & Right(strNewReport, Len(strNewReport) - lulWave - 7)
                End If
                strReport = strNewReport
                ReplaceElement = True
            End If
        End If
    End If
End Function

Private Function StrToASC(ByVal strIn As String) As String
    '将中文字符串转换为ASC串（包括英文一起）
    '先将特殊字符进行转义：
    strIn = Replace(strIn, Chr(9), "\TAB ")
    strIn = Replace(strIn, Chr(13) + Chr(10), "\par ")
    Dim i As Long, s As String, lsChar As String, lsPart1 As String, lsPart2 As String
    Dim lsCharHex As String
    For i = 1 To Len(strIn)
        lsChar = Mid(strIn, i, 1)
        If lsChar = "?" Then
            lsCharHex = LCase(Hex(Asc(lsChar)))
            If Len(lsCharHex) = 4 Then
                lsCharHex = "\'" + Mid(lsCharHex, 1, 2) + "\'" + Mid(lsCharHex, 3, 2)
            Else
                lsCharHex = lsChar
            End If
            s = s + lsCharHex
        Else
            lsCharHex = LCase(Hex(Asc(lsChar)))
            If Len(lsCharHex) = 4 Then
                lsCharHex = "\'" + Mid(lsCharHex, 1, 2) + "\'" + Mid(lsCharHex, 3, 2)
            Else
                lsCharHex = lsChar
            End If
            s = s + lsCharHex
        End If
    Next
    StrToASC = s
End Function

Private Function getElementPos(ByVal strReport As String, ByRef lStart As Long, ByRef lLength As Long, _
    ByVal lEnd As Long, ByRef lulWave As Long, ByRef lulNone As Long) As Boolean
'    lulWave   '下划波浪标记\ulwave的开始位置
'    lulNone    '关闭所有下划线标记\ulnone的开始位置
    '查找从lStart开始的，元素内容文本的开始位置和长度
    '查找和定位元素中的下划波浪标记\ulwave 和 关闭所有下划线标记\ulnone
    Dim lIndex As Long
    Dim lWordEnd As Long
    Dim blnSearch As Boolean
    Dim strChar As String
    Dim strNextChar As String
    Dim blnInWord As Boolean
    Dim strTemp As String
    
    lIndex = lStart
    blnSearch = True
    blnInWord = True
    
    While (blnSearch And lIndex < lEnd)
        strChar = Mid(strReport, lIndex, 1)
        If strChar = "\" Then       '上一个控制字符结束，下一个控制字符，或者是文本的开始
            strNextChar = Mid(strReport, lIndex + 1, 1)
            If strNextChar = "'" Or strNextChar = "{" Or strNextChar = "}" Or strNextChar = "\" Then     '文本的开始
                '往后找第一个控制符
                blnInWord = True
                lStart = lIndex - 1
                While (blnInWord And lIndex <= lEnd)
                    lIndex = lIndex + 1
                    strChar = Mid(strReport, lIndex, 1)
                    If strChar = "\" Then
                        strNextChar = Mid(strReport, lIndex + 1, 1)
                        If strNextChar = "'" Or strNextChar = "{" Or strNextChar = "}" Or strNextChar = "\" Then
                            lIndex = lIndex + 1
                        Else
                            lWordEnd = lIndex - 1
                            blnInWord = False   '退出内容循环
                        End If
                    End If
                Wend
            Else    '控制字符的开始
                '往后读取一直到控制字符结束
                strTemp = Mid(strReport, lIndex, 1)
                lIndex = lIndex + 1
                While (Mid(strReport, lIndex, 1) <> "\" And Mid(strReport, lIndex, 1) <> " ")
                    strTemp = strTemp & Mid(strReport, lIndex, 1)
                    lIndex = lIndex + 1
                Wend
                If strTemp = "\ulwave" Then
                    lulWave = lIndex - 8
                ElseIf strTemp = "\ulnone" Then
                    lulNone = lIndex - 8
                    blnSearch = False   '退出查找元素的循环
                End If
            End If
        ElseIf strChar = " " Then   '正文开始，而且正文的字符是英文，不是中文
            '往后找第一个控制符
            blnInWord = True
            lStart = lIndex - 1
            While (blnInWord And lIndex <= lEnd)
                lIndex = lIndex + 1
                strChar = Mid(strReport, lIndex, 1)
                If strChar = "\" Then
                    strNextChar = Mid(strReport, lIndex + 1, 1)
                    If strNextChar = "'" Or strNextChar = "{" Or strNextChar = "}" Or strNextChar = "\" Then
                        lIndex = lIndex + 1
                    Else
                        lWordEnd = lIndex - 1
                        blnInWord = False   '退出内容循环
                    End If
                End If
            Wend
            
        Else        '在不是正确的RTF文件，返回查找错误
            getElementPos = False
            Exit Function
        End If
    Wend
    lLength = lWordEnd - lStart
    If lWordEnd = 0 Then  '说明是查到要素结束了，才退出的，没有查找到内容文本
        getElementPos = False
    Else
        getElementPos = True
    End If
End Function


Private Function ProtectAndHide(ByRef strReport As String, ByVal lStart As Long, ByVal lEnd As Long) As Boolean
    Dim lOnPos As Long
    Dim lOffPos As Long
    
    '往前找隐藏和保护开始标记，\v和\protect
    lOnPos = InStrRev(strReport, "\v", lStart, vbTextCompare)
    lOffPos = InStrRev(strReport, "\v0", lStart, vbTextCompare)
    If lOnPos > lOffPos And lOnPos <> 0 Then
        '查找后面的隐藏标记
        lOnPos = InStr(lEnd, strReport, "\v", vbTextCompare)
        lOffPos = InStr(lEnd, strReport, "\v0", vbTextCompare)
        If lOffPos <= lOnPos And lOffPos <> 0 Then
            '查找前面的保护标记
            lOnPos = InStrRev(strReport, "\protect", lStart, vbTextCompare)
            lOffPos = InStrRev(strReport, "\protect0", lStart, vbTextCompare)
            If lOnPos > lOffPos And lOnPos <> 0 Then
                '查找后面的保护标记
                lOnPos = InStr(lEnd, strReport, "\protect", vbTextCompare)
                lOffPos = InStr(lEnd, strReport, "\protect0", vbTextCompare)
                If lOffPos <= lOnPos And lOffPos <> 0 Then
                    ProtectAndHide = True
                End If
            End If
        End If
    End If
End Function


Public Sub SaveReportView()
    Dim strReport As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strElements As String
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    
    arrSQL = Array()
    
    On Error GoTo ErrHandle
    
    '修改报告签名要素，将其内容替换为“ ”
    strElements = SPLITER_REPORT & Report_Element_报告签名 & SPLITER_ELEMENT & " "
    '组织专科报告内容
    If mblnShowSpecial = True Then
        strElements = strElements & mfrmReportSpecial.getElementString
    End If
    '组织大文本段的对象属性,如果Tag为空，则从数据库读取默认值
    If mfrmReportView.rtxtCheckView.Tag = "" Or mfrmReportView.rtxtResult.Tag = "" Or mfrmReportView.rTxtAdvice.Tag = "" Then
        strSql = "Select a.内容文本 As 标题, b.对象属性 From 电子病历内容 a,电子病历内容 b " & _
             " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 And b.终止版 = 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
        While rsTemp.EOF = False
            Select Case rsTemp!标题
                Case "检查所见"
                    If mfrmReportView.rtxtCheckView.Tag = "" Then
                        mfrmReportView.rtxtCheckView.Tag = rsTemp!对象属性
                        RefreshViewTag mfrmReportView.rtxtCheckView
                    End If
                Case "诊断意见"
                    If mfrmReportView.rtxtResult.Tag = "" Then
                        mfrmReportView.rtxtResult.Tag = rsTemp!对象属性
                        RefreshViewTag mfrmReportView.rtxtResult
                    End If
                Case "建议"
                    If mfrmReportView.rTxtAdvice.Tag = "" Then
                        mfrmReportView.rTxtAdvice.Tag = rsTemp!对象属性
                        RefreshViewTag mfrmReportView.rTxtAdvice
                    End If
            End Select
            rsTemp.MoveNext
        Wend
    End If
    
    '最后保存大文本段内容，此时会根据数据库内容，自动更新报告中的要素
    strReport = SPLITER_REPORT & "1" & mfrmReportView.rtxtCheckView.Tag & SPLITER_ELEMENT & mfrmReportView.rtxtCheckView.Text & SPLITER_REPORT _
        & "2" & mfrmReportView.rtxtResult.Tag & SPLITER_ELEMENT & mfrmReportView.rtxtResult.Text & SPLITER_REPORT _
        & "3" & mfrmReportView.rTxtAdvice.Tag & SPLITER_ELEMENT & mfrmReportView.rTxtAdvice.Text
    
    '更改内容的时候，保存的签名级别始终是0，最后具体的签名级别通过签名的过程来更改
    strSql = "ZL_影像报告内容_update(" & mlngAdviceID & "," & mReportID & ",'" & Replace(strReport, "'", "’") & " ','" & strElements & "'," & m目标版本 & ",0)"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSql
    
    '有随访权限的，才能保存随访描述
    If InStr(mstrPrivs, "随访") > 0 Then
        strSql = "Zl_影像随访_Update(" & mlngAdviceID & ",'" & Replace(mfrmReportView.txtReview.Text, "'", "’") & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSql
    End If
    
    gcnOracle.BeginTrans        '----------保存报告内容
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存报告内容")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    mfrmReportView.pModified = False
    If mblnShowSpecial = True Then
        mfrmReportSpecial.pModified = False
    End If
    
    Exit Sub
ErrHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
        Call SaveErrLog
End Sub

Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim cbrControlItem As CommandBarControl
    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    '添加格式选择弹出菜单
    If CommandBar.Parent.ID = conMenu_PacsReport_SelFormat Then
        CommandBar.Controls.DeleteAll
        
        '添加新的菜单项
        For i = 1 To UBound(rptFormats)
            Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_SelFormat_Item, rptFormats(i).strName, i)
            cbrControlItem.parameter = rptFormats(i).ID
        Next i
    ElseIf CommandBar.Parent.ID = conMenu_PacsReport_RepFormat Then
        If mblnRefreshRptFormat = True And mbln使用自定义报表 = True Then
            CommandBar.Controls.DeleteAll
            
            '添加新的菜单项
            strSql = "Select a.编号,b.序号,b.说明 From zlreports a,zlrptfmts b Where a.Id=b.报表ID And a.编号=[1] Order By 序号"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取自定义报表格式", mstr报表编号)
            
            While rsTemp.EOF = False
                Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_RepFormat_Item, rsTemp!序号 & "-" & Nvl(rsTemp!说明))
                cbrControlItem.Style = xtpButtonIconAndCaption
                cbrControlItem.Checked = (InStr(mstr选中报表格式, cbrControlItem.Caption) <> 0)
                cbrControlItem.parameter = rsTemp!序号
                cbrControlItem.CloseSubMenuOnClick = False
            
                rsTemp.MoveNext
            Wend
            
            '关闭刷新
            mblnRefreshRptFormat = False
        End If
    ElseIf CommandBar.Parent.ID = conMenu_PacsReport_VerifySign Then
        '签名验证的弹出菜单，列出可以验证的签名版本
        CommandBar.Controls.DeleteAll
        
        '添加新的签名验证菜单
        strSql = "Select 开始版,内容文本 as 签名医生 From 电子病历内容 Where 文件ID = [1] And 对象类型 =8  Order By 开始版"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取各个签名版本", mReportID)
        
        While rsTemp.EOF = False
            Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_VerifySign_Item, rsTemp!开始版 & "-" & Nvl(rsTemp!签名医生))
            cbrControlItem.Style = xtpButtonIconAndCaption
            cbrControlItem.Checked = False
            cbrControlItem.parameter = rsTemp!开始版
            rsTemp.MoveNext
        Wend
    End If
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    
    Select Case control.ID
        Case conMenu_File_Print, conMenu_File_Preview      '打印报告,预览报告
            If InStr(mstrPrivs, "PACS报告打印") <= 0 Then
                control.Visible = False
            Else
                control.Visible = True
            End If
            
            '如果未找到对应的病历文件，那么打印预览按钮会被禁用
            If mblnPrintView = True Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
        Case comMenu_Petition_Capture
            '读取启用申请单扫描参数
            mblnIsPetitionScan = IIf(Val(GetDeptPara(mlngDeptID, "启用申请单扫描", 1)) = 1, True, False)
            If mblnIsPetitionScan Then
                control.Visible = True
            Else
                control.Visible = False
            End If
            
        Case conMenu_Edit_Modify        '报告编辑
            '可见性Visible跟保存的条件一样，只不过没有状态条件Enable，只要可见就可以操作
            '在报告书写状态下，有报告书写权限的人，可以书写自己的报告，有他人报告权限的人，可以书写本科室别人的报告
            If m目标版本 = 1 And InStr(mstrPrivs, "PACS报告书写") > 0 Then
                If mstrEPR创建人 = UserInfo.姓名 Then
                    control.Visible = True
                ElseIf (InStr(mstrPrivs, "PACS他人报告") > 0 And mlngEPRDeptID = mlngDeptID) Then '有他人报告权限的，可以书写本科室的报告
                    control.Visible = True
                Else
                    control.Visible = False
                End If
            ElseIf m目标版本 > 1 And InStr(mstrPrivs, "PACS报告修订") > 0 Then
                '在报告修订的状态下，有报告修订权限的人，可以书写本科室的报告。
                control.Visible = True
            Else
                control.Visible = False
            End If
            
            control.Enabled = mblnEditable
            
        Case conMenu_PacsReport_Reject
            '判断是否具备报告驳回权限
            '判断当前报告所在状态是否允许驳回
            If InStr(mstrPrivs, "报告驳回") <= 0 Then
                control.Visible = False
            Else
                control.Visible = True
                control.Enabled = mReportID <> 0 And Not mblnReadOnly
            End If
        Case conMenu_PacsReport_RejectHistory
            If InStr(mstrPrivs, "报告驳回") > 0 Then
                control.Visible = False
            Else
                control.Visible = True
            End If
            
        Case conMenu_PacsReport_Save    '保存
            '在报告书写状态下，有报告书写权限的人，可以书写自己的报告，有他人报告权限的人，可以书写本科室别人的报告
            If m目标版本 = 1 And InStr(mstrPrivs, "PACS报告书写") > 0 Then
                If mstrEPR创建人 = UserInfo.姓名 Then
                    control.Visible = True
                ElseIf (InStr(mstrPrivs, "PACS他人报告") > 0 And mlngEPRDeptID = mlngDeptID) Then '有他人报告权限的，可以书写本科室的报告
                    control.Visible = True
                Else
                    control.Visible = False
                End If
            ElseIf m目标版本 > 1 And InStr(mstrPrivs, "PACS报告修订") > 0 Then
                '在报告修订的状态下，有报告修订权限的人，可以书写本科室的报告。
                control.Visible = True
            Else
                control.Visible = False
            End If
            
            '根据报告的状态，确定是否开启书写的Enable
            If control.Visible = True Then
                If mblnReadOnly = True Then
                    control.Enabled = False
                ElseIf mblnModified = False Then
                    mblnModified = chkModified
                    
                    If mblnModified = True Then
                        If mdtReportTime <> GetReportLastSaveTime(mlngAdviceID) Then
                            mblnModified = False
                            
                            Call zlUpdateAdviceInf(mlngAdviceID, mlngSendNo, mlngStudyState, mblnMoved)
                            Call zlRefreshFace(True)
                        Else
                            control.Enabled = True
                                                
                            '从非编辑模式，进入编辑模式，触发报告编辑事件
                            RaiseEvent BeforeEdit(mlngAdviceID)
                    
                            tmrCheckingReportState.Enabled = True
                        End If
                    Else
                        control.Enabled = False
                    End If
                Else
                    control.Enabled = True
                End If
            End If
            
        Case conMenu_PacsReport_Sign    '签名
            
            '在书写模式下，还没有签名的，可以签名
            '在修订模式下，签名数量没有超过16次的，可以签名。
            '只读模式下，什么都不能操作。
            If m目标版本 = 1 And InStr(mstrPrivs, "PACS报告书写") > 0 Then    '还没有签名,而且有书写权限
                If mstrEPR创建人 = UserInfo.姓名 Then '自己写的报告，自己签名
                    control.Visible = True
                ElseIf (InStr(mstrPrivs, "PACS他人报告") > 0 And mlngEPRDeptID = mlngDeptID) Then     '有他人报告权限的，可以给本科室的报告签名
                    control.Visible = True
                Else
                    control.Visible = False
                End If
            ElseIf m目标版本 > 1 And InStr(mstrPrivs, "PACS报告修订") > 0 Then    '已经有签名了，再次签名，则需要修订的权限
                control.Visible = (m目标版本 <= 16)
            Else
                control.Visible = False
            End If
            If control.Visible = True Then control.Enabled = Not mblnReadOnly
            
        Case conMenu_PacsReport_VerifySign  '签名验证
            '只有启用了数字签名，才显示签名验证按钮
            '只有报告书写，报告修订权限的人，才能对签名进行验证
            control.Visible = IIf(mlngPassType = 0, False, True)
            
            If control.Visible = True Then
                If m目标版本 > 1 And (InStr(mstrPrivs, "PACS报告修订") > 0 Or InStr(mstrPrivs, "PACS报告书写") > 0) Then
                    control.Visible = True
                Else
                    control.Visible = False
                End If
            End If
        Case conMenu_PacsReport_DelSign '回退
            
            '没有签名之前，不可以回退,只能回退自己的签名，或者通过“回退他人签名”的权限，回退本科室其他人的签名
            If m目标版本 > 1 Then   '只有签名过后才可以回退
                If mSigns("K" & mSigns.GetMaxKey).姓名 = UserInfo.姓名 And m签名级别 <> cprSL_空白 Then  '回退自己的签名
                    control.Visible = True
                ElseIf mstrEPR保存人 = UserInfo.姓名 And m签名级别 = cprSL_空白 Then   '回退自己的修订
                    control.Visible = True
                ElseIf InStr(mstrPrivs, "PACS他人报告") <> 0 And mlngEPRDeptID = mlngDeptID Then      '有他人报告权限的,可以回退本科室的他人签名
                    control.Visible = True
                Else
                    control.Visible = False
                End If
            Else
                control.Visible = False
            End If
            If control.Visible = True Then
                control.Enabled = (Not mblnReadOnly) And mblnCanUntread
            End If
            
        Case conMenu_PacsReport_SelFormat  '选择格式 '修订模式下，不可以设置格式
            If InStr(mstrPrivs, "PACS报告书写") <= 0 Then
                control.Visible = False
            Else
                control.Visible = IIf(m目标版本 = 1, True, False)
            End If
        Case conMenu_PacsReport_RepFormat   '选择打印格式
            control.Visible = mbln使用自定义报表
        Case conMenu_PacsReport_RepFormat_Item  '选择具体打印格式
            control.IconId = IIf(control.Checked, 90002, 90001)
        Case conMenu_PacsReport_FontSet                     '设置文本段字体
            If InStr(mstrPrivs, "PACS报告书写") > 0 Or InStr(mstrPrivs, "PACS报告修订") > 0 Then
                control.Visible = True
            Else
                control.Visible = False
            End If
        Case conMenu_PacsReport_SaveWord                    '保存词句示范
            control.Visible = IIf(mintWordPower <> -1, True, False)
        Case conMenu_Edit_Delete                            '删除报告
            control.Visible = (mReportID <> 0 And (InStr(mstrPrivs, "PACS报告书写") > 0 Or InStr(mstrPrivs, "PACS报告删除") > 0))
            If control.Visible = True And InStr(mstrPrivs, "PACS报告删除") > 0 And mlngEPRDeptID = mlngDeptID Then Exit Sub     '可以强制删除本科室的报告
            If control.Visible = True Then control.Visible = mlngEPRDeptID = mlngDeptID
            If control.Visible = True Then control.Visible = (m目标版本 = 1)
            If control.Visible = True Then control.Visible = (InStr(mstrPrivs, "PACS他人报告") > 0 Or mstrEPR创建人 = UserInfo.姓名)

            '这里先对删除报告的Enable进行基础设置，在外面工作站里面，还会根据报告的状态，再设置一次是否可以删除
            If control.Visible = True Then control.Enabled = Not mblnReadOnly
            
        Case conMenu_PacsReport_ClearWritingState       '清除报告“处理中”的状态,可以清除本科室的报告标记
            control.Visible = (mReportID <> 0 And InStr(mstrPrivs, "PACS报告删除") > 0 And mlngEPRDeptID = mlngDeptID)
            If control.Visible = True And mlngAdviceID <> 0 Then
                '“清除状态”的菜单触发时机是右键弹出菜单，或者下拉显示菜单，不是一直在刷新的
                '在显示之前先查询数据库，如果当前有操作人，则显示此菜单
                Dim rsTemp As ADODB.Recordset
                Dim strSql As String
                strSql = "Select 医嘱ID From 影像检查记录 Where 医嘱ID = [1] And 报告操作 Is Not Null "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "判断报告是否在处理中", mlngAdviceID)
                control.Visible = (rsTemp.RecordCount <> 0)
            End If
        
        Case conMenu_File_Exit      '退出,独立窗口模式下，显示“退出”按钮
            control.Visible = IIf(mblnSingleWindow = True, True, False)
            
        Case comMenu_Petition_Capture
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Enabled = False
            End If
        Case conMenu_PacsReport_Default
    End Select
End Sub

Private Function GetReportLastSaveTime(ByVal lngAdviceID As Long) As Date
'获取报告最后保存的时间
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetReportLastSaveTime = mdtReportTime
    
    strSql = "select 保存时间 from 病人医嘱报告 a, 电子病历记录 b where a.病历ID=b.ID and a.医嘱ID=[1]"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsData.RecordCount <= 0 Then
        If mReportID > 0 Then GetReportLastSaveTime = Now
        Exit Function
    End If
    
    GetReportLastSaveTime = Nvl(rsData!保存时间, mdtReportTime)
End Function

Private Sub chkOtherDeptReport_Click()
    mblnCheckOtherDeptReport = chkOtherDeptReport.value
    Call subShowHistoryList
End Sub

Private Sub cmdSelectWord_Click()
    Dim strReportVieweType As String
    
    On Error GoTo err
    
    ' mintReportViewType 0-检查所见CheckView，1-诊断意见Result，2-建议Advice
    If mintReportViewType = 0 Then
        strReportVieweType = ReportViewType_检查所见
    ElseIf mintReportViewType = 1 Then
        strReportVieweType = ReportViewType_诊断意见
    Else
        strReportVieweType = ReportViewType_建议
    End If
    
    If rtxtReport.SelText <> "" Then
        Call mfrmReportWord_WordSelected(rtxtReport.SelText, strReportVieweType, False)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub cmdViewImage_Click()
    '打开观片站看图像
    Dim lngViewAdviceID As Long
    
    On Error GoTo err
    If lvHistoryList.SelectedItem Is Nothing Then Exit Sub

    lngViewAdviceID = Mid(lvHistoryList.SelectedItem.Key, 2)
    Call OpenViewer(1, pobjPacsCore, lngViewAdviceID, True, mobjOwner)
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    '定义Pane的ID顺序： 1-检查所见；2-历史报告；3-词句示范；4-报告图；5-视频采集；6-专科报告。
    Select Case Item.ID
        Case 1  '历史报告
            Item.Handle = picReportHistoryList.hWnd
            picReportHistoryList.Visible = True
        Case 2  '检查所见
            Item.Handle = mfrmReportView.hWnd
        Case 3  '词句示范
            If Not mfrmReportWord Is Nothing Then Item.Handle = mfrmReportWord.hWnd
        Case 4  '报告图
            If Not mfrmReportImage Is Nothing Then Item.Handle = mfrmReportImage.hWnd
        Case 5  '视频采集
            If mblnUseActiveCapture Then
                If Not mobjWork_ActiveVideo Is Nothing Then
                    Item.Handle = mobjWork_ActiveVideo.ContainerHwnd
                End If
            End If
        Case 6  '专科报告
            If Not mfrmReportSpecial Is Nothing Then Item.Handle = mfrmReportSpecial.hWnd
    End Select
End Sub


Private Sub Form_Activate()
    '显示嵌套视频采集
    If mblnUseActiveCapture Then
        If Not mobjWork_ActiveVideo Is Nothing Then Call mobjWork_ActiveVideo.zlRefreshVideoWindow
    End If
End Sub


Private Sub InitActiveVideoModuleObj()
'初始化ActivexExe视频采集模块对象
    If mobjWork_ActiveVideo Is Nothing Then
        Set mobjWork_ActiveVideo = CreateObject("zl9PacsCapture.clsPacsCapture") ' New zl9PacsCapture.clsPacsCapture
        mobjWork_ActiveVideo.ParentWindowKey = Me.Name & IIf(mblnSingleWindow = True, "Dock", "")
        
        Call mobjWork_ActiveVideo.zlInitModule(gcnOracle, glngSys, mlngModule, mstrPrivs, mlngDeptID, Me.hWnd, Me, True)
    End If
End Sub



Public Sub RefreshVideo()
    If mblnUseActiveCapture Then
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlRefreshVideoWindow
        End If
    End If
End Sub


'发送报告，供专科报告插件调用
Public Sub SendReport(ByVal strDescription As String, _
    ByVal strResult As String, ByVal strAdvice As String)
    
    mfrmReportView.rtxtCheckView.Text = strDescription
    mfrmReportView.rtxtResult.Text = strResult
    mfrmReportView.rTxtAdvice.Text = strAdvice
    
End Sub

'获取报告，供专科报告插件调用
Public Sub GetReport(ByRef strDescription As String, _
    ByRef strResult As String, ByRef strAdvice As String)
    
    strDescription = mfrmReportView.rtxtCheckView.Text
    strResult = mfrmReportView.rtxtResult.Text
    strAdvice = mfrmReportView.rTxtAdvice.Text
    
End Sub

'清除报告，供专科报告插件调用
Public Sub ClearReport(Optional ByVal blnClearDescription As Boolean = True, _
    Optional ByVal blnClearResult As Boolean = True, _
    Optional ByVal blnClearAdvice As Boolean = True)
    
    If blnClearDescription Then mfrmReportView.rtxtCheckView.Text = ""
    If blnClearResult Then mfrmReportView.rtxtResult.Text = ""
    If blnClearAdvice Then mfrmReportView.rTxtAdvice.Text = ""
    
End Sub

Private Sub Form_Load()
    mblnClosed = False
    
    InitCommandBars '初始化菜单，跟科室无关
    
    mblnMenuDownState = False
End Sub


Private Function GetSignVerifyType() As Long
'获取签名类型，默认为密码签名,1表示数字签名
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetSignVerifyType = 0
    
    strSql = "select Zl_Fun_Getsignpar([1], 7) as 签名类型 from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询签名类型", mlngDeptID)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetSignVerifyType = Nvl(rsData!签名类型, 0)
End Function

Private Sub InitLoaclParas(lngDeptID As Long, lngModuleId As Long, strPrivs As String, Optional blnIsPacsStation As Boolean = False)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strRegPath As String
    Dim blnInCreatProReport As Boolean
    
    If lngDeptID = 0 Then Exit Sub
    
    mlngDeptID = lngDeptID
    mstrPrivs = strPrivs
    mlngModule = lngModuleId
    
    '设置默认值
    mblnShowImage = False       '默认不显示图像区域
    mblnShowSpecial = False     '默认不显示专科报告
    mblnShowVideoCapture = False '默认不显示图像采集区域
    
    mstrSpecialForm = ""
    mblnExitAfterPrint = False  '默认打印报告后不关闭窗体
    mintWordDblClick = 0        '默认词句双击后直接写入报告
    mintImageDblClick = 0       '默认缩略图双击后直接写入报告
    pReport_CheckViewName = "检查所见"  '默认名称
    pReport_ResultName = "诊断意见"     '默认名称
    pReport_AdviceName = "建议"         '默认名称
'    mblnIgnoreResult = False            '忽略结果阴阳性
'    mintResultInput = 1                 '输入提示，默认是签名后提示
    mblnShowWord = True                 '默认一直显示词句示范
    mblnCheckPrintPara = False             '默认允许打印
    mblnCheckOtherDeptReport = False    '默认不查看其他科的历史报告
    mblnUntreadPrinted = False          '默认审核打印后不允许回退
    mintPaneID = 1                      '默认选中Pane为第一个Pane
    
    mblnTechReptSame = GetDeptPara(lngDeptID, "只能填写自己检查的报告", 0) = "1"  '只能填写自己检查的报告
    
    
    
    '读取检查所见区域，诊断意见区域，建议区域 和签名区域的高度
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
    End If
    
    mlngCY21 = GetSetting("ZLSOFT", strRegPath, "CY21", 500)
    mlngCY22 = GetSetting("ZLSOFT", strRegPath, "CY22", 250)
    mlngCX1 = GetSetting("ZLSOFT", strRegPath, "CX1", 250)
    mlngCX2 = GetSetting("ZLSOFT", strRegPath, "CX2", 500)
    mlngCX3 = GetSetting("ZLSOFT", strRegPath, "CX3", 250)
    mlngCY3 = GetSetting("ZLSOFT", strRegPath, "CY3", 250)
    mlngCX4 = GetSetting("ZLSOFT", strRegPath, "CX4", 250)
    mlngCY4 = GetSetting("ZLSOFT", strRegPath, "CY4", 250)
    mlngPicHistoryX = GetSetting("ZLSOFT", strRegPath, "PicHistoryX", 250)
    mlngPicHistoryY = GetSetting("ZLSOFT", strRegPath, "PicHistoryY", 250)
    mlngPrivateWordY = GetSetting("ZLSOFT", strRegPath, "PrivateWordY", 250)
    
    mintPaneID = Val(GetSetting("ZLSOFT", strRegPath, "选中PANE", 1))
    mstr选中报表格式 = GetSetting("ZLSOFT", strRegPath, "选中报表格式", "")
    mstr报表编号 = GetSetting("ZLSOFT", strRegPath, "报表编号", "")
     
    mblnCheckOtherDeptReport = (Val(zlDatabase.GetPara("查看他科历史报告", glngSys, mlngModule, 0)) = 1)
    
'    '读取当前签名方式（系统参数26）,诊疗报告是从 3开始
'    mlngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), 3, 1))  '门诊,住院,医技,护理 (1111),为空默认采用密码模式
    mlngPassType = GetSignVerifyType()
    
    On Error GoTo err
    strSql = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    
    While Not rsTemp.EOF
        Select Case rsTemp!参数名
            Case "显示报告图像"
                mblnShowImage = Nvl(rsTemp!参数值, 0)
            Case "显示视频采集"
                mblnShowVideoCapture = Nvl(rsTemp!参数值, 0)
                
                If blnIsPacsStation Then
                  mblnShowVideoCapture = False
                End If
                
            Case "打印后退出"
                mblnExitAfterPrint = Nvl(rsTemp!参数值, 0)
            Case "报告中显示大图"
                mlngShowBigImg = Nvl(rsTemp!参数值, 0)
            Case "报告大图放大倍数"
                mdblBigImgZoom = Val(Nvl(rsTemp!参数值, 1))
                If mdblBigImgZoom = 0 Then mdblBigImgZoom = 1
            Case "报告缩略图数量"
                mintMinImageCount = Val(Nvl(rsTemp!参数值, 8))
            Case "显示专科报告"
                mblnShowSpecial = Nvl(rsTemp!参数值, 0)
            Case "专科报告页"
                mstrSpecialForm = Nvl(rsTemp!参数值)
            Case "缩略图双击操作"
                mintImageDblClick = Val(Nvl(rsTemp!参数值, 0))
            Case "检查所见名称"
                pReport_CheckViewName = Nvl(rsTemp!参数值, "检查所见")
            Case "诊断意见名称"
                pReport_ResultName = Nvl(rsTemp!参数值, "诊断意见")
            Case "建议名称"
                pReport_AdviceName = Nvl(rsTemp!参数值, "建议")
            Case "显示词句示范"
                mblnShowWord = IIf(Nvl(rsTemp!参数值, 0) = 0, True, False)
            Case "报告词句双击操作"
                mintWordDblClick = Val(Nvl(rsTemp!参数值, 0))
            Case "平诊需审核才能打报告"
                mblnCheckPrintPara = Nvl(rsTemp!参数值, 0) = 1
            Case "审核打印后允许回退"
                mblnUntreadPrinted = Nvl(rsTemp!参数值, 0) = 1
        End Select
        rsTemp.MoveNext
    Wend
    
    mblnUseActiveCapture = False
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Or mlngModule = G_LNG_PATHSTATION_MODULE Then
        mblnUseActiveCapture = GetSetting("ZLSOFT", "公共模块", "UseActiveVideo", "1")
        Call SaveSetting("ZLSOFT", "公共模块", "UseActiveVideo", mblnUseActiveCapture)
    End If
    
    mstrImageLevel = Nvl(GetDeptPara(mlngDeptID, "影像质量等级", "甲,乙"))
    mstrReportLevel = Nvl(GetDeptPara(mlngDeptID, "报告质量等级", "甲,乙"))

    mintCriticalValues = Val(GetDeptPara(mlngDeptID, "危急情况判断", 0))           '危急情况判断
    mblnIgnoreResult = GetDeptPara(mlngDeptID, "忽略结果阴阳性", 0) = "1" '        '忽略结果阴阳性
    mintConformDetermine = Val(GetDeptPara(mlngDeptID, "符合情况判定", 0))         '符合情况判定
    
    mlngHintType = Val(GetDeptPara(mlngDeptID, "诊断结果提示类型", 0))
    
    
    mblnReportWithResult = GetDeptPara(mlngDeptID, "无影像诊断为阴性", 0) = "1" '  '无影像诊断为阴性
    
    '处理词句示范窗口
    If mblnShowWord = True And (Not mfrmReportWord Is Nothing) Then
        mfrmReportWord.mblnShowWord = mblnShowWord
        '如果直接显示词句示范，则去掉采集窗体的控制框
        FormSetCaption mfrmReportWord, False, False
    Else
        mfrmReportWord.mblnShowWord = mblnShowWord
        '如果直接显示词句示范，则去掉采集窗体的控制框
        FormSetCaption mfrmReportWord, True, True
    End If
                
'    '卸载原有窗体,卸载后，会导致报告中看不到这个窗体，专科报告以后要修改成一个统一的窗体，目前暂时不处理
    If Not mfrmReportSpecial Is Nothing Then
'        If mstrSpecialForm <> Report_Form_frmReportCustom Then Unload mfrmReportSpecial
        If TypeName(mfrmReportSpecial) <> "clsZLPacsProReport" Then Unload mfrmReportSpecial
        
        Set mfrmReportSpecial = Nothing
    End If
'
'    If Not mfrmReportImage Is Nothing Then
'        Unload mfrmReportImage
'        Set mfrmReportImage = Nothing
'    End If
    
    
    '装载图像窗体
    If mblnShowImage = True Then
        If mfrmReportImage Is Nothing Then Set mfrmReportImage = New frmReportImage
    End If
    
    '设置专科报告窗体
    If mblnShowSpecial = True Then
        
        Select Case mstrSpecialForm
            Case Report_Form_frmReportES
                Set mfrmReportSpecial = New frmReportES
            Case Report_Form_frmReportUS
                Set mfrmReportSpecial = New frmReportUS
            Case Report_Form_frmReportPathology
                Set mfrmReportSpecial = New frmReportPathology
            Case Report_Form_frmReportCustom
                blnInCreatProReport = True
                Set mfrmReportSpecial = CreateObject("ZLPacsProReport.clsZLPacsProReport")
                Call mfrmReportSpecial.InitPlugin(gcnOracle, Me)
                blnInCreatProReport = False
        End Select
    End If
    
    If mfrmReportSpecial Is Nothing Then    '如果没有找到对应的专科窗体，则设置为不使用专科报告
        mblnShowSpecial = False
    End If
    
    Exit Sub
err:
    If blnInCreatProReport = True And (err.Number = 429 Or err.Number = -2147024770) Then
        MsgBoxD Me, "没有找到自定义专科报告部件“ZLPacsProReport.dll”，请注册此部件后重试。"
        Set mfrmReportSpecial = Nothing
        mblnShowSpecial = False
    Else
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End If
End Sub

Private Sub InitReportFormat()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i  As Integer
    
    ReDim rptFormats(1) As rptFormat
    rptFormats(1).ID = 0
    rptFormats(1).strName = "标准格式"
    
    If mFileID = 0 Then Exit Sub
    
    strSql = "Select Id,名称 From 病历范文目录 Where 文件ID = [1] And 性质= 0 And (通用级=0 Or (通用级=1 And 科室ID=[2]) " & _
            " Or (通用级=2 And 人员ID= [3])) "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mFileID, UserInfo.部门ID, UserInfo.ID)
    If rsTemp.RecordCount <> 0 Then
        ReDim Preserve rptFormats(rsTemp.RecordCount + 1) As rptFormat
        For i = 1 To rsTemp.RecordCount
            rptFormats(i + 1).ID = rsTemp!ID
            rptFormats(i + 1).strName = rsTemp!名称
            rsTemp.MoveNext
        Next i
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim strRegPath As String
    
    '提示是否保存报告
    Call PromptModify
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
    End If
    
    '保存报告历史区域的宽度和高度
    SaveSetting "ZLSOFT", strRegPath, "PicHistoryY", picReportHistoryList.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "PicHistoryX", picReportHistoryList.Width
    
    
    '保存历史报告显示状态
    zlDatabase.SetPara "查看他科历史报告", chkOtherDeptReport.value, glngSys, mlngModule
    
    '保存报告中的DockingPane位置
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        Call SaveSetting("ZLSOFT", strRegPath & "\" & mlngModule & "\" & TypeName(dkpMain), dkpMain.Name & mlngDeptID, dkpMain.SaveStateToString)
    End If
    
    '保存第一个被选中的PANE编号
    mintPaneID = 1
    For i = 1 To dkpMain.PanesCount
        If dkpMain.Panes(i).Selected Then
            mintPaneID = i
            Exit For
        End If
    Next i
    SaveSetting "ZLSOFT", strRegPath, "选中PANE", mintPaneID
    
    '如果是自定义报表格式打印，则保存选中报表格式
    If mbln使用自定义报表 Then
        SaveSetting "ZLSOFT", strRegPath, "选中报表格式", mstr选中报表格式
        SaveSetting "ZLSOFT", strRegPath, "报表编号", mstr报表编号
    End If


    If mblnShowVideoCapture Then
        If mblnUseActiveCapture Then
            If Not mobjWork_ActiveVideo Is Nothing Then
'                Call mobjWork_ActiveVideo.FreeResource
                Set mobjWork_ActiveVideo = Nothing
            End If
        End If
    End If
    
    '卸载子窗体
    If Not mfrmReportView Is Nothing Then
        Unload mfrmReportView       '报告所见
        Set mfrmReportView = Nothing
    End If
    
    If Not mobjReport Is Nothing Then
        Unload mobjReport.zlGetForm        '电子病历报告
        Set mobjReport = Nothing
    End If
    
    If Not mfrmReportWord Is Nothing Then
        Unload mfrmReportWord       '词句示范
        Set mfrmReportWord = Nothing
    End If
    
    If Not mfrmReportImage Is Nothing Then
        Unload mfrmReportImage   '图像选择
        Set mfrmReportImage = Nothing
    End If
    
    If Not mfrmReportSpecial Is Nothing Then
        If mstrSpecialForm <> Report_Form_frmReportCustom Then Unload mfrmReportSpecial
        Set mfrmReportSpecial = Nothing
    End If
    
    
    '独立窗口模式,此模式下记录窗口位置,触发关闭事件
    If mblnSingleWindow = True Then
        Call SaveWinState(Me, App.ProductName)
        
        RaiseEvent AfterClosed(mlngAdviceID)
        
'        If Not mobjOwner Is Nothing Then
'            mobjOwner.EditorClosed (mlngAdviceID)
'        End If
    End If
    
    mblnClosed = True
End Sub


Private Sub lvHistoryList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    zlControl.LvwSortColumn lvHistoryList, ColumnHeader.Index
End Sub

Private Sub lvHistoryList_DblClick()
    Dim lngViewAdviceID As Long
    Dim lngViewReportID As Long
    
    If lvHistoryList.SelectedItem Is Nothing Then Exit Sub
    
    lngViewAdviceID = Mid(lvHistoryList.SelectedItem.Key, 2)
    lngViewReportID = lvHistoryList.SelectedItem.SubItems(5)
    Call frmReportHistory.zlShowMe(Me, lngViewAdviceID, lngViewReportID)
End Sub

Public Sub WordItemClick(strReportViewType As String, strReportViewTypeAlias As String, strContext As String)
    If mblnShowWord = True Then Exit Sub


    If mblnSingleWindow = True Then
        Call mfrmReportWord.zlShowMe(Me, mFileID, strReportViewType, strReportViewTypeAlias, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mintWordPower, mblnEditable)
    Else
        Call mfrmReportWord.zlShowMe(mobjOwner, mFileID, strReportViewType, strReportViewTypeAlias, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mintWordPower, mblnEditable)
    End If
End Sub

Private Sub lvHistoryList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lngViewAdviceID As Long
    Dim lngViewReportID As Long
    Dim strSql As String
    Dim strSeparator2 As String
    Dim rsTemp As ADODB.Recordset
    Dim blnShow As Boolean
    Dim mstrOffset As String
    Dim strTitle As String
    Dim strText As String
    Dim lngStart As Long
    
    rtxtReport.Text = ""
    mstrOffset = "  "
    
    On Error GoTo err
    
    lngViewAdviceID = Mid(Item.Key, 2)
    lngViewReportID = Item.SubItems(5)
    '显示报告内容
    
    '读取报告的内容
    strSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文,b.开始版 as 版本 From 电子病历内容 a,电子病历内容 b " & _
             " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 and b.终止版=0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngViewReportID)
    
    While Not rsTemp.EOF
        blnShow = False
        Select Case rsTemp!标题
            Case "检查所见"
                strTitle = pReport_CheckViewName & strSeparator2
                strText = vbCrLf & mstrOffset & Nvl(rsTemp!正文) & vbCrLf & vbCrLf
                blnShow = True
            Case "诊断意见"
                strTitle = pReport_ResultName & strSeparator2
                strText = vbCrLf & mstrOffset & Nvl(rsTemp!正文) & vbCrLf & vbCrLf
                blnShow = True
            Case "建议"
                strTitle = pReport_AdviceName & strSeparator2
                strText = vbCrLf & mstrOffset & Nvl(rsTemp!正文) & vbCrLf & vbCrLf
                blnShow = True
        End Select
        
        If blnShow = True Then
            lngStart = Len(rtxtReport.Text)
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = 0
            rtxtReport.SelText = strTitle
            
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = Len(strTitle)
            rtxtReport.SelFontName = "宋体"
            rtxtReport.SelFontSize = 10
            rtxtReport.SelColor = vbBlue
            rtxtReport.SelBold = True
            
            lngStart = Len(rtxtReport.Text)
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = 0
            rtxtReport.SelText = strText
            
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = Len(strText)
            rtxtReport.SelFontName = "宋体"
            rtxtReport.SelFontSize = 10
            rtxtReport.SelColor = vbBlack
            rtxtReport.SelBold = False
        End If
        rsTemp.MoveNext
    Wend
    rtxtReport.SelStart = 0
    
    '检查是否有报告图像
    cmdViewImage.Enabled = False
    strSql = "Select 检查UID from 影像检查记录 where 医嘱ID =[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngViewAdviceID)
    If rsTemp.EOF = False Then
        If Nvl(rsTemp!检查UID) <> "" Then
            cmdViewImage.Enabled = True
        End If
    End If
    
    If InStr(mstrPrivs, "PACS报告书写") <= 0 Then
        cmdSelectWord.Enabled = False
    Else
        cmdSelectWord.Enabled = True
    End If

    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetControlFocus(objControl As Object)
On Error Resume Next
    If objControl.Visible Then objControl.SetFocus
err.Clear
End Sub


Private Sub mfrmReportView_AdviceClick(ByVal strContext As String)
    mintReportViewType = 2
    If mblnShowWord = True Then
        Call mfrmReportWord.zlRefresh(mFileID, ReportViewType_建议, pReport_AdviceName, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mblnShowWord, mintWordDblClick, mintWordPower, mblnEditable)
        Call SetControlFocus(mfrmReportView)
    End If
End Sub

Private Sub mfrmReportView_CheckViewClick(ByVal strContext As String)
    mintReportViewType = 0
    If mblnShowWord = True Then
        Call mfrmReportWord.zlRefresh(mFileID, ReportViewType_检查所见, pReport_CheckViewName, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mblnShowWord, mintWordDblClick, mintWordPower, mblnEditable)
        Call SetControlFocus(mfrmReportView)
    End If
End Sub

Private Sub mfrmReportView_ResultClick(ByVal strContext As String)
    mintReportViewType = 1
    If mblnShowWord = True Then
        Call mfrmReportWord.zlRefresh(mFileID, ReportViewType_诊断意见, pReport_ResultName, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mblnShowWord, mintWordDblClick, mintWordPower, mblnEditable)
        Call SetControlFocus(mfrmReportView)
    End If
End Sub

Private Sub mfrmReportWord_AddSampleWord(ByVal blnIsAllWord As Boolean)
    Call subSaveWord(IIf(blnIsAllWord, 2, 0))
End Sub

Private Sub mfrmReportWord_ModifySampleWord()
    Call subSaveWord(1)
End Sub

Private Sub mobjCustomReport_AfterPrint(ByVal ReportNum As String)
    '激活记录打印的事件
    If Not mobjOwner Is Nothing Then
        mobjOwner.AfterPrinted (mlngAdviceID)
    Else
        RaiseEvent AfterPrinted(mlngAdviceID)
    End If
    mblnPrintOK = True
End Sub

Private Sub mobjReport_AfterPrinted(ByVal lngOrderID As Long)
    
    '激活记录打印的事件
    If Not mobjOwner Is Nothing Then
        mobjOwner.AfterPrinted (lngOrderID)
    Else
        RaiseEvent AfterPrinted(lngOrderID)
    End If
    mblnPrintOK = True
End Sub

Private Sub mobjReport_AfterSaved(ByVal lngOrderID As Long, ByVal lngSaveType As Long)

    Call AfterReportSaved(lngOrderID, lngSaveType)
    '更新编辑器中的内容
    Call zlRefreshFace(True)
End Sub


Private Sub mfrmReportView_ShowWord(intReportViewType As Integer, strContext As String)
    Dim strReportViewType As String
    Dim strReportViewTypeAlias As String
    
    Select Case intReportViewType
        Case 0
            strReportViewType = ReportViewType_检查所见
            strReportViewTypeAlias = pReport_CheckViewName
        Case 1
            strReportViewType = ReportViewType_诊断意见
            strReportViewTypeAlias = pReport_ResultName
        Case 2
            strReportViewType = ReportViewType_建议
            strReportViewTypeAlias = pReport_AdviceName
    End Select
    
    If mblnSingleWindow = True Then
        Call mfrmReportWord.zlShowMe(Me, mFileID, strReportViewType, strReportViewTypeAlias, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mintWordPower, mblnEditable)
    Else
        Call mfrmReportWord.zlShowMe(mobjOwner, mFileID, strReportViewType, strReportViewTypeAlias, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mintWordPower, mblnEditable)
    End If
End Sub

Private Sub mfrmReportWord_WordSelected(strWord As String, strReportViewType As String, blnIsPopupWindInsert As Boolean)
    '判断文字应该回写到哪里？
    
    '如果报告不允许编辑，则不允许修改报告词句
    If mblnReadOnly Then Exit Sub
    
    strWord = strWord & vbCrLf
    Select Case strReportViewType
        Case ReportViewType_检查所见
            If blnIsPopupWindInsert Then mfrmReportView.rtxtCheckView.Text = ""
            Call mfrmReportView.zlWriteReport(strWord, 0)
        Case ReportViewType_诊断意见
            If blnIsPopupWindInsert Then mfrmReportView.rtxtResult.Text = ""
            Call mfrmReportView.zlWriteReport(strWord, 1)
        Case ReportViewType_建议
            If blnIsPopupWindInsert Then mfrmReportView.rTxtAdvice.Text = ""
            Call mfrmReportView.zlWriteReport(strWord, 2)
        Case ReportViewType_病理诊断
            If mfrmReportSpecial.Name = "frmReportES" Then
                If blnIsPopupWindInsert Then mfrmReportSpecial.txtPathologyDiag.Text = ""
                Call mfrmReportSpecial.zlWriteWord(strWord, strReportViewType)
            End If
        Case ReportViewType_活检部位
            If mfrmReportSpecial.Name = "frmReportES" Then
                If blnIsPopupWindInsert Then mfrmReportSpecial.txt活检部位.Text = ""
                Call mfrmReportSpecial.zlWriteWord(strWord, strReportViewType)
            End If
    End Select
End Sub

Private Sub picReportHistoryList_Resize()
    On Error Resume Next
    
    chkOtherDeptReport.Left = 0
    chkOtherDeptReport.Top = 0
    
    lvHistoryList.Left = 0
    lvHistoryList.Top = chkOtherDeptReport.Height + 10
    lvHistoryList.Width = picReportHistoryList.ScaleWidth
    lvHistoryList.Height = picReportHistoryList.ScaleHeight / 5
    lvHistoryList.Refresh
    
    frmReportDetail.Left = 20
    frmReportDetail.Top = lvHistoryList.Top + lvHistoryList.Height + 10
    frmReportDetail.Width = Abs(picReportHistoryList.ScaleWidth - 20)
    frmReportDetail.Height = Abs(picReportHistoryList.ScaleHeight - frmReportDetail.Top - 50)
    
    cmdViewImage.Left = 200
    cmdViewImage.Top = 200
    
    cmdSelectWord.Left = cmdViewImage.Left + cmdViewImage.Width + 200
    cmdSelectWord.Top = cmdViewImage.Top
    
    rtxtReport.Left = 50
    rtxtReport.Top = cmdViewImage.Top + cmdViewImage.Height + 50
    rtxtReport.Width = Abs(frmReportDetail.Width - 100)
    rtxtReport.Height = Abs(frmReportDetail.Height - cmdViewImage.Height - 300)
End Sub


Private Sub pobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    Call RefPacsPic '刷新图片
End Sub

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
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
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    '采集工具栏定义
    Set cbrToolBar = Me.cbrMain.Add("报告栏", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Open, "书写"): cbrControl.IconId = 3002: cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): cbrControl.IconId = 102: cbrControl.ToolTipText = "报告预览"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): cbrControl.IconId = 103: cbrControl.ToolTipText = "报告打印"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Save, "保存"): cbrControl.IconId = 3091: cbrControl.ToolTipText = "保存报告"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Sign, "签名"): cbrControl.IconId = 3003: cbrControl.ToolTipText = "签名"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Reject, "报告驳回"): cbrControl.IconId = 229: cbrControl.ToolTipText = "报告驳回"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_RejectHistory, "驳回历史"): cbrControl.IconId = 8341: cbrControl.ToolTipText = "驳回历史"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelSign, "回退"): cbrControl.IconId = 3004: cbrControl.ToolTipText = "回退签名"
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_VerifySign, "签名验证"): cbrControl.IconId = 8044: cbrControl.ToolTipText = "签名验证"
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_RepFormat, "打印格式"): cbrControl.IconId = 3031: cbrControl.ToolTipText = "选择自定义报表格式"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_AddNumber, "序号"): cbrControl.IconId = 9023: cbrControl.ToolTipText = "给段落文字添加序号"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_FontSet, "字体"): cbrControl.IconId = 509: cbrControl.ToolTipText = "字体设置"
        cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置"): cbrControl.IconId = 181: cbrControl.ToolTipText = "打印设置"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "病历修订"): cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_History, "历史"): cbrControl.IconId = 3564: cbrControl.ToolTipText = "查看当前和历史报告的的修订情况"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_SaveWord, "词句"): cbrControl.IconId = 741: cbrControl.ToolTipText = "将报告内容保存成词句示范"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_PrivOrder, "上一个"): cbrControl.IconId = 21802: cbrControl.ToolTipText = "上一个检查"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_NextOrder, "下一个"): cbrControl.IconId = 21801: cbrControl.ToolTipText = "下一个检查"
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_PacsReport_SelFormat, "报告格式"): cbrControl.IconId = 227: cbrControl.ToolTipText = "选择和更换报告单格式"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "病历编辑"): cbrControl.IconId = 3002: cbrControl.ToolTipText = "用电子病历方式编辑报告"
        Set cbrControl = .Add(xtpControlButton, comMenu_Petition_Capture, "申请单"): cbrControl.IconId = 3935: cbrControl.ToolTipText = "查看已扫描的申请单图像": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Default, "恢复界面"): cbrControl.IconId = 3936: cbrControl.ToolTipText = "恢复默认界面": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"):: cbrControl.IconId = 191
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If (cbrControl.type = xtpControlButton) Or (cbrControl.type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
        If cbrControl.Category = "" Then cbrControl.Category = "Main" '设置成主界面菜单
    Next
    cbrToolBar.Position = xtpBarTop
End Sub



' 从电子病历中复制过来的一些过程
'################################################################################################################
'## 功能：  将指定的LOB字段复制为临时文件
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  存放内容的文件名，失败则返回零长度""
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).value) Then Exit Do
        strText = rsLob.Fields(0).value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
End Function

'################################################################################################################
'## 功能：  将指定的文件保存到指定记录的LOB字段中
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  成功返回True，失败返回False
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobSave(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    err = 0: On Error GoTo errHand
    
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        gstrSQL = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "zlBlobSave")
    Next
    Close lngFileNum
    zlBlobSave = True
    Exit Function

errHand:
    Close lngFileNum
    zlBlobSave = False
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If strZipFile = "" Then Exit Function
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If mobjFSO.FileExists(strZipPath & "TMP.RTF") Then mobjFSO.DeleteFile strZipPath & "TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
End Function

'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If strFile = "" Then Exit Function
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function




Public Sub RefPacsPic()
    '刷新可选的报告图像
    If mblnShowImage = True Then
        If Not mfrmReportImage Is Nothing Then
            mfrmReportImage.RefPacsPic
        End If
    End If
End Sub

Private Sub subSetModifyFlag(blnModifyFlag As Boolean)
    mblnModified = blnModifyFlag
    mfrmReportView.pModified = blnModifyFlag
    If mblnShowImage = True Then
        mfrmReportImage.pMarkModified = blnModifyFlag
        mfrmReportImage.pImageModified = blnModifyFlag
    End If
    If mblnShowSpecial = True Then
        mfrmReportSpecial.pModified = blnModifyFlag
    End If
End Sub

Public Function PromptModify(Optional blnCancelEdit As Boolean = False) As Boolean
    'blnCancelEdit =True 表示直接取消保存
    If blnCancelEdit = True Then
        Call subSetModifyFlag(False)
        PromptModify = False
        Exit Function
    End If
    
    If mlngAdviceID <> 0 And mblnModified = True And (Not cbrMain.FindControl(, conMenu_PacsReport_Save, True, True) Is Nothing) And Not mblnIsReportDelete Then
        '模拟按下ESC键盘，避免在主界面的快速过滤窗口进行过滤后，过滤菜单不消失，仍然可以点击的情况
        keybd_event VK_ESCAPE, 0, 0, 0
        keybd_event VK_ESCAPE, 0, 2, 0
        
        If MsgBoxD(Me, "病人的报告有所改变，是否保存？", vbYesNo, gstrSysName) = vbYes Then
            Call SaveReport(True)
            PromptModify = True
        Else
            '不保存报告时，清空报告操作数据
            Call UpdateReporter(mlngAdviceID, "")
            
            Call subSetModifyFlag(False)
            PromptModify = False
            
            '对于嵌入式的报告方式，此时相当于是关闭窗口
            If mblnSingleWindow = False Then
                RaiseEvent AfterClosed(mlngAdviceID)
            End If
        End If
    End If
End Function

Private Sub subShowHistoryList()
    Dim strSql As String
    Dim strSQLBack As String
    Dim rsTemp As ADODB.Recordset
    Dim objItem As ListItem
    Dim strTime As String
    Dim iCount As Integer
    Dim strFilter As String
    
    '先检查权限，确定是否显示他科历史报告
    
    If InStr(mstrPrivs, "PACS报告他科报告") Then
        chkOtherDeptReport.value = IIf(mblnCheckOtherDeptReport = True, 1, 0)
        chkOtherDeptReport.Enabled = True
    Else
        chkOtherDeptReport.value = 0
        mblnCheckOtherDeptReport = False
        chkOtherDeptReport.Enabled = False
    End If
    
    If chkOtherDeptReport.value = 1 Then
        strFilter = ""
    Else
        strFilter = " And c.执行科室id in(select  部门id  from 部门人员 where 人员id = [2] union all select to_Number([3]) from dual) "
    End If
                    
    strSql = "Select c.Id As 医嘱id, a.影像类别, c.开嘱时间, c.医嘱内容, b.病历id " & _
            " From 影像检查记录 A, 病人医嘱报告 B, 病人医嘱记录 C, 影像检查记录 D, 病人医嘱记录 E " & _
            " Where a.医嘱id = b.医嘱id And d.医嘱id = e.Id And e.Id =[1] And b.医嘱id = c.Id And " & _
            " (c.病人id = e.病人id Or a.关联id = d.关联id) And c.相关id Is Null "
            
    strSql = strSql & strFilter
    
    If mblnMoved = True Then
        strSQLBack = strSql
        strSQLBack = Replace(strSQLBack, "影像检查记录", "H影像检查记录")
        strSQLBack = Replace(strSQLBack, "病人医嘱报告", "H病人医嘱报告")
        strSQLBack = Replace(strSQLBack, "病人医嘱记录", "H病人医嘱记录")
        strSql = strSql & " UNION ALL  " & strSQLBack
    End If
    
    strSql = strSql & " Order By 开嘱时间 Asc "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "显示报告历史", mlngAdviceID, UserInfo.ID, mlngDeptID)
    
    lvHistoryList.ListItems.Clear
    
    iCount = 1
    zlControl.LvwSelectColumns lvHistoryList, "医嘱ID,0,0,1;序号,600,0,1;类别,600,0,1;开嘱时间,1100,0,1;医嘱内容,3000,0,1;病历ID,0,0,1", True
    With lvHistoryList
        Do While Not rsTemp.EOF
            Set objItem = .ListItems.Add(, "K" & rsTemp!医嘱ID, rsTemp!医嘱ID)
            '添加子项目
            objItem.SubItems(1) = iCount
            iCount = iCount + 1
            objItem.SubItems(2) = Nvl(rsTemp!影像类别)
            strTime = Format(rsTemp!开嘱时间, "yyyy-mm-dd")
            objItem.SubItems(3) = strTime
            objItem.SubItems(4) = Nvl(rsTemp!医嘱内容)
            objItem.SubItems(5) = Nvl(rsTemp!病历Id)
            rsTemp.MoveNext
        Loop
    End With
    
    If lvHistoryList.ListItems.Count > 0 Then
        lvHistoryList.ListItems(1).Selected = True
        Call lvHistoryList_ItemClick(lvHistoryList.ListItems(1))
    Else
        rtxtReport.Text = ""
    End If
    
    dkpMain.FindPane(1).Title = "历史报告（" & lvHistoryList.ListItems.Count & "）"
End Sub

Private Sub ChangeOrder(intType As Integer)
    'intType 切换类型 1 --上一个；2--下一个
    Dim lngRowIndex As Long
    Dim lngNewOrderID As Long
    Dim lngNewSendNo As Long
    Dim blnMoved As Boolean
    
    On Error GoTo err
    
    If mobjOwner.ufgStudyList.DataGrid.Rows <= 1 Then Exit Sub

    lngRowIndex = mobjOwner.ufgStudyList.FindRowIndex(mlngAdviceID, "医嘱ID", True)
    
    If lngRowIndex <= 0 Then Exit Sub
    
    '只能在非登记状态下进行切换
    Do While True
        '查找上一个或下一个医嘱
        If intType = 1 Then     '上一个医嘱
            lngRowIndex = lngRowIndex - 1
            If lngRowIndex <= 0 Then lngRowIndex = mobjOwner.ufgStudyList.DataGrid.Rows - 1
        ElseIf intType = 2 Then         '下一个医嘱
            lngRowIndex = lngRowIndex + 1
            If lngRowIndex >= mobjOwner.ufgStudyList.DataGrid.Rows Then lngRowIndex = 1
        End If
        
        If mobjOwner.ufgStudyList.Text(lngRowIndex, "检查过程") <> "已登记" Then Exit Do
    Loop
        

    Call zlUpdateAdviceInf(Val(mobjOwner.ufgStudyList.Text(lngRowIndex, "医嘱ID")), _
                        Val(mobjOwner.ufgStudyList.Text(lngRowIndex, "发送号")), _
                        Val(mobjOwner.ufgStudyList.Text(lngRowIndex, "检查状态")), _
                        IIf(mobjOwner.ufgStudyList.Text(lngRowIndex, "转出") = 1, True, False))
        
    If mblnSingleWindow = True Then      '独立窗口，直接刷新本窗体
        Call zlRefreshFace(True)
    Else            '嵌入式窗口，通过外部事件触发刷新,同时刷新病历等其他工作页面
        mobjOwner.ufgStudyList.DataGrid.Row = lngRowIndex
    End If
            
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub AfterReportSaved(lngOrderID As Long, ByVal lngSaveType As Long)
'lngSaveType:0-普通保存，1-诊断签名，2-审核签名，3-回退签名
    On Error GoTo err
    If mblnSingleWindow = True Then
        '对弹出窗口执行父窗体的AfterReportSaved过程
        If Not mobjOwner Is Nothing Then
            Call mobjOwner.AfterReportSaved(lngOrderID, Me, lngSaveType)
        End If
    Else
        '对嵌入式窗口，触发AfterSaved事件
        RaiseEvent AfterSaved(lngOrderID, Me, lngSaveType)
        '对于嵌入式的报告方式，此时相当于是关闭窗口,触发AfterClosed事件
        RaiseEvent AfterClosed(lngOrderID)
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chkPrintState()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
        
    strSql = "Select a.报告人,a.复核人,b.紧急标志 ,b.Id From 影像检查记录 a ,病人医嘱记录 b Where a.医嘱id = b.Id And b.Id = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "验证是否可以打印", mlngAdviceID)
    
    If rsTemp.EOF = False Then
        mblnCanPrint = IIf(Nvl(rsTemp!紧急标志, 0) = 1, Nvl(rsTemp!报告人) <> "", Nvl(rsTemp!复核人) <> "")
    Else
        mblnCanPrint = False
    End If
End Sub

Private Sub AddNumber()
'给文本段添加前导的数字序号
'mintReportViewType 0-检查所见CheckView，1-诊断意见Result，2-建议Advice

    Dim rText As RichTextBox
    Dim strText As String
    Dim iCount As Integer
    Dim iStart As Integer
    
    If mintReportViewType < 0 Or mintReportViewType > 2 Then Exit Sub
    
    '判断是哪个文本段被选中,读取文本段的对象属性
    If mintReportViewType = 0 Then
        Set rText = mfrmReportView.rtxtCheckView
    ElseIf mintReportViewType = 1 Then
        Set rText = mfrmReportView.rtxtResult
    ElseIf mintReportViewType = 2 Then
        Set rText = mfrmReportView.rTxtAdvice
    End If
    
    On Error GoTo err
    strText = rText.Text
    '先判断文本段是否被锁定
    If rText.Locked = True Then
        MsgBoxD Me, "文本段被锁定，请先双击解锁后再添加数字编号。", vbOKOnly, "信息提示"
        Exit Sub
    End If
    '先判断该文本段中第一个字符是否数字1，如果是，则提示已经有数字编号，是否还要添加
    If Left(strText, 1) = "1" Then
        If MsgBoxD(Me, "本段文本中已经包含数字编号，是否还要添加数字编号？", vbOKCancel, "信息提示") = vbCancel Then
            Exit Sub
        End If
    End If
    '开始添加数字编号,每一个回车之后，如果不是空格，就添加序号
    iStart = 1
    '第一行也需要判断是否存在缩进
    If Left(strText, 1) <> " " Then
        iCount = 1
        strText = iCount & ". " & strText
    Else
        iCount = 0
    End If
    iStart = InStr(iStart, strText, vbCrLf)
    While (iStart <> 0)
        If Mid(strText, iStart + 2, 1) <> " " And Mid(strText, iStart + 2, 2) <> vbCrLf And Mid(strText, iStart + 2, 1) <> "" Then
            iCount = iCount + 1
            strText = Left(strText, iStart + 1) & iCount & ". " & Right(strText, Len(strText) - iStart - 1)
        End If
        iStart = InStr(iStart + 1, strText, vbCrLf)
    Wend
    
    rText.Text = strText
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subChangeRptFormat(ByVal control As XtremeCommandBars.ICommandBarControl)
'更改被选中的自定义报表打印格式
    On Error GoTo err
    
    If control.Checked = True Then
        '添加报表格式
        mstr选中报表格式 = IIf(mstr选中报表格式 = "", control.Caption, mstr选中报表格式 & "," & control.Caption)
    Else
        '删除报表格式
        If InStr(mstr选中报表格式, control.Caption) = 1 Then
            mstr选中报表格式 = Replace(mstr选中报表格式, control.Caption, "")
            If Left(mstr选中报表格式, 1) = "," Then
                mstr选中报表格式 = Mid(mstr选中报表格式, 2)
            End If
        Else
            mstr选中报表格式 = Replace(mstr选中报表格式, "," & control.Caption, "")
        End If
    End If
    
    If InStr(mstrFormatInfo, vbCrLf) <> 0 Then
        mstrFormatInfo = Left(mstrFormatInfo, InStr(mstrFormatInfo, vbCrLf) - 1)
    End If
    mstrFormatInfo = mstrFormatInfo & vbCrLf & "打印格式：" & mstr选中报表格式
    Call mfrmReportView.zlRefreshLblFormat(mstrFormatInfo)
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subPrintReport(blnPrint As Boolean, blnSilent As Boolean)
'使用自定义报表打印和预览报告
'参数： blnPrint---True打印；False预览
'       blnSilent ---强制静默打印，批量打印时需要
        
    Dim blnNoAsk As Boolean     '是否静默打印
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strExseNo As String, intExseKind As Integer
    Dim strPicPath As String
    Dim objFile As New Scripting.FileSystemObject
    Dim intPCount As Integer
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim i As Integer, j As Integer, intParaCount As Integer
    Dim strPicFile As String
    Dim aryPara(19) As String, aryFlagPara(1) As String     '报告图中的图像记录
    Dim aryPrintPara(19) As String, strFlagString As String '实际传给自定义报表的内容
    Dim dcmImages As New DicomImages, dcmResultImage As DicomImage
    Dim arr报表格式() As String
    Dim int格式号 As Integer
    Dim intRows As Integer, intCols As Integer
    
    On Error GoTo err
    
    If mblnCanPrint = False Then
        MsgBoxD Me, "当前报告未审核，不能打印，请检查！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否静默打印
    blnNoAsk = (zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1")
    If blnSilent = True Then blnNoAsk = True
    
    '提取报告的记录性质和No
    strSql = "Select 记录性质, No From 病人医嘱发送 Where 医嘱id = [1]"
    If mblnMoved = True Then strSql = Replace(strSql, "病人医嘱发送", "H病人医嘱发送")
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提前记录性质和No", mlngAdviceID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    strExseNo = "" & rsTemp!NO
    intExseKind = Val("" & rsTemp!记录性质)
    
    If mobjCustomReport Is Nothing Then Set mobjCustomReport = New clsReport
    
    If Not blnNoAsk Then
        If mobjCustomReport.ReportPrintSet(gcnOracle, glngSys, mstr报表编号, Me) = False Then
            zlRefreshFace True  '调用该语句强制刷新报告界面，以避免取消打印机设置后，界面显示错误
            Exit Sub
        End If
    End If
    
    '获取图像
    strPicPath = App.Path & "\TmpImage\"
    If objFile.FolderExists(strPicPath) = False Then objFile.CreateFolder strPicPath
    
    '获取报告图像（包括标记图）生成本地文件
    '一个报告表格中可能排列多个报告图
    intPCount = 0
    strSql = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
    If mblnMoved = True Then strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取图像", mReportID)
    Do While Not rsTemp.EOF
        Set cTable = New cEPRTable
        If cTable.GetTableFromDB(cprET_单病历审核, mReportID, Val("" & rsTemp!表格ID)) Then
            For i = 1 To cTable.Pictures.Count
                strPicFile = "PACSPic" & i & ".JPG"
                If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True
                If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                    Set oPicture = cTable.Pictures(i).DrawFinalPic
                Else
                    Set oPicture = cTable.Pictures(i).OrigPic
                End If
                SavePicture oPicture, strPicFile
                If objFile.FileExists(strPicFile) Then
                    '保存标记图和图象的路径
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                        aryFlagPara(0) = strPicFile
                    Else
                        aryPara(intPCount) = strPicFile
                        dcmImages.AddNew
                        dcmImages(dcmImages.Count).FileImport strPicFile, "BMP"
                        intPCount = intPCount + 1
                        If intPCount > UBound(aryPara) Then Exit Do
                    End If
                End If
            Next i
        End If
        rsTemp.MoveNext
    Loop
    
    '根据选择的自定义报表格式，组织图像
    '如果只选择了一种格式，则检查是否只有一个图象框,只有一个图像框的时候，自动组合图像。
    '如果选择了2种以上的格式，则对只有一个图像框的情况不作自动组合
    arr报表格式 = Split(mstr选中报表格式, ",")
    '处理没有选择格式的情况
    If UBound(arr报表格式) = -1 Then
        ReDim arr报表格式(0) As String
        arr报表格式(0) = "1-1"
    End If
    If UBound(arr报表格式) = 0 Then     '只有一种格式
        int格式号 = Split(arr报表格式(0), "-")(0)
        strSql = "Select b.名称,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = [2] And b.名称 not like '标记%'"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询是否需要组合图像", mstr报表编号, int格式号)
        If rsTemp.RecordCount = 1 And intPCount >= 1 Then
            '组合图象
            ResizeRegion intPCount, rsTemp("W"), rsTemp("H"), intRows, intCols
            Set dcmResultImage = AssembleImage(dcmImages, intRows, intCols, rsTemp("H"), rsTemp("W"))
            dcmResultImage.FileExport Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "=")), "JPEG"
        End If
    End If
    
    '获取图像，调用报表
    intPCount = 0       '记录图像的数量
    For i = 0 To UBound(arr报表格式)
        int格式号 = Split(arr报表格式(i), "-")(0)
        
        strSql = "Select b.名称 From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = [2]" & vbNewLine & _
        "       Order By b.名称" 'Trunc(b.y/567),Trunc(b.x/567)
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取图象框", mstr报表编号, int格式号)
        '装载图像数据
        intParaCount = 0
        Do While Not rsTemp.EOF
            
            '分别装在标记图和报告图
            If InStr(rsTemp!名称, "标记") <> 0 Then '标记图
                If aryFlagPara(0) <> "" Then strFlagString = rsTemp!名称 & "=" & aryFlagPara(0)
            Else    '报告图
                If intPCount > UBound(aryPara) Then Exit Do     '图像数量超过报告中的图像，退出
                If aryPara(intPCount) = "" Then Exit Do         '报表中的图象框比报告中的多，退出
                
                aryPrintPara(intParaCount) = rsTemp!名称 & "=" & aryPara(intPCount)
                intPCount = intPCount + 1
                intParaCount = intParaCount + 1
            End If
            rsTemp.MoveNext
        Loop
        
        '处理报表中图形比报告中少的情况
        For j = intParaCount To UBound(aryPrintPara)
            If aryPrintPara(j) Like "*=*" Then aryPrintPara(j) = ""
        Next j
        
        '调用报表
        Call mobjCustomReport.ReportOpen(gcnOracle, glngSys, mstr报表编号, Nothing, _
            "NO=" & strExseNo, "性质=" & intExseKind, "医嘱ID=" & mlngAdviceID, strFlagString, _
            aryPrintPara(0), aryPrintPara(1), aryPrintPara(2), aryPrintPara(3), aryPrintPara(4), aryPrintPara(5), _
            aryPrintPara(6), aryPrintPara(7), aryPrintPara(8), aryPrintPara(9), aryPrintPara(10), aryPrintPara(11), _
            aryPrintPara(12), aryPrintPara(13), aryPrintPara(14), aryPrintPara(15), aryPrintPara(16), aryPrintPara(17), _
            aryPrintPara(18), aryPrintPara(19), "ReportFormat=" & int格式号, IIf(blnPrint, 2, 1))
    Next i
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subSaveWord(intType As Integer)
'保存病历词句示范
'从报告词句树分类结点提取词句分类ID，分类名称,KEY="T-分类ID",TEXT="分类名称"
'从报告词句树叶子结点提取词句ID，词句名称，KEY="L-示范ID"，TEXT="示范名称"
'参数： intType ---  0 新增；1 修改

    Dim strWordString As String
    Dim rText As RichTextBox
    Dim lngClassID As Long      '分类ID
    Dim strClassName As String  '分类名称
    Dim objNode As Node
    Dim lngWordID As Long       '词句示范ID
    Dim strWordName As String   '词句示范名称
    
    
    If mfrmReportWord.trvWordTree.SelectedItem Is Nothing Then
        MsgBoxD Me, "请从词句树中先选择需要保存词句的位置。", vbOKOnly, gstrSysName
        Exit Sub
    End If
    Set objNode = mfrmReportWord.trvWordTree.SelectedItem
    
    If intType = 1 Then         '修改，需要读取词句ID和分类ID，新增，只需要读取分类ID
        '读取词句ID
        '先判断当前选中的结点是分类结点还是叶子结点，标记是分类结点KEY=“T-...”,叶子结点KEY=“L-...”
        ''是叶子结点，需要查找上级结点,是分类结点，直接提取分类ID和名称
        If Left(objNode.Key, 1) = "L" Then
            lngWordID = Right(objNode.Key, Len(objNode.Key) - 2)
            strWordName = objNode.Text
        Else
            MsgBoxD Me, "您现在选择的是分类，请选择需要修改的词句。", vbOKOnly, gstrSysName
            Exit Sub
        End If
    ElseIf intType = 2 Then
        strWordString = ""
        
        If mfrmReportView.rtxtCheckView.Text <> "" Then
            strWordString = "<<所见>>" & mfrmReportView.rtxtCheckView.Text
        End If
        
        If mfrmReportView.rtxtResult.Text <> "" Then
            strWordString = strWordString & vbCrLf & "<<诊断>>" & mfrmReportView.rtxtResult.Text
        End If
        
        If mfrmReportView.rTxtAdvice.Text <> "" Then
            strWordString = strWordString & vbCrLf & "<<建议>>" & mfrmReportView.rTxtAdvice.Text
        End If
    Else
        '从报告中读取词句内容
        '提取当前需要保存成词句的内容
                'mintReportViewType= 0-检查所见CheckView，1-诊断意见Result，2-建议Advice
                
        If mintReportViewType = 0 Then
            Set rText = mfrmReportView.rtxtCheckView
        ElseIf mintReportViewType = 1 Then
            Set rText = mfrmReportView.rtxtResult
        Else
            Set rText = mfrmReportView.rTxtAdvice
        End If
        
        If rText.SelLength = 0 Then
            strWordString = rText.Text
        Else
            strWordString = rText.SelText
        End If
    End If
    '提取当前词句分类ID
    If Left(objNode.Key, 1) = "L" Then  '当前结点是叶子结点，则指向上级分类结点
            Set objNode = objNode.Parent
    End If
    
    lngClassID = Right(objNode.Key, Len(objNode.Key) - 2)
    strClassName = objNode.Text
    
    Call frmReportWordList.zlShowMe(Me, strWordString, mintWordPower, lngClassID, strClassName, _
                                    mlngDeptID, lngWordID)
End Sub

Private Function GetReportImageSelected() As Boolean
'------------------------------------------------
'功能：检查报告图页面是否当前活动页面
'参数：
'返回：True－－是活动页面，False－－不是活动页面
'-----------------------------------------------
Dim i As Integer

On Error Resume Next

GetReportImageSelected = False

For i = 1 To dkpMain.PanesCount
    If dkpMain.Panes(i).Title = "报告图" Then
        GetReportImageSelected = dkpMain.Panes(i).Selected
        Exit For
    End If
Next i
End Function


Private Sub FuncAdviceSignVerify(int签名版本 As Integer, blnMoved As Boolean)
'------------------------------------------------
'功能：校验检查报告的电子签名(可对已转移的数据),校验版本为int签名版本 的签名
'参数： int签名版本 -- 本次需要验证的签名的版本
'       blnMoved -- 数据是否被迁移
'返回：
'-----------------------------------------------
    Dim strSource As String
    Dim objESign As Object                  '电子签名接口部件
    Dim lng签名ID  As Long                  '签名所在的行的ID
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim intRule As Integer                  '记录签名规则
    
    
    On Error GoTo err
    
    '根据报告ID和签名版本查找签名内容
    strSql = "Select Id , 开始版 From 电子病历内容 Where 文件ID = [1] And 对象类型 = 8 and 开始版 =[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取最后签名版本", mReportID, int签名版本)
    If rsTemp.RecordCount = 0 Then
        MsgBoxD Me, "本次报告没有版本为" & int签名版本 & "的签名，无法对数字签名做验证。"
        Exit Sub
    End If
    
    lng签名ID = rsTemp!ID
    
    '提取源文
    intRule = GetSignSourceString(2, mReportID, int签名版本, blnMoved, Nothing, strSource)
    '如果返回的规则=0，表示提取源文失败
    If intRule = 0 Then
        MsgBoxD Me, "本次报告版本为" & int签名版本 & "的签名源文提取失败，无法对数字签名做验证。"
        Exit Sub
    End If
    
    '创建签名对象，对源文进行签名验证
    Set objESign = CreateObject("zl9ESign.clsESign")
    If err <> 0 Then err = 0
    If Not objESign Is Nothing Then
        Call objESign.Initialize(gcnOracle, glngSys)
        '签名验证
        Call objESign.VerifySignature(strSource, lng签名ID, 2)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub tmrCheckingReportState_Timer()
'循环检查报告编辑状态，如果已被人编辑，则进行提示
On Error Resume Next
    If mblnReadOnly Then
        tmrCheckingReportState.Enabled = False
        Exit Sub
    End If
    
    If CheckConcurrentReport(Me, mlngAdviceID) Then
        mblnReadOnly = False
        
        tmrCheckingReportState.Tag = Val(tmrCheckingReportState.Tag) + 1
        
        '大于10秒则退出
        If Val(tmrCheckingReportState.Tag) > 5 Then tmrCheckingReportState.Enabled = False
    Else
        mblnReadOnly = True
    End If
    err.Clear
End Sub
