VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.UserControl ucPacsHelper 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   LockControls    =   -1  'True
   ScaleHeight     =   10020
   ScaleWidth      =   6885
   Begin VB.CommandButton cmdAttach 
      Caption         =   "‖"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "H"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   375
      Left            =   6120
      Picture         =   "ucPacsHelper.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cmdDel 
      Height          =   375
      Left            =   6120
      Picture         =   "ucPacsHelper.ctx":047A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   375
      Left            =   6120
      Picture         =   "ucPacsHelper.ctx":08F4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   6120
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox cmdMenu 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   6120
      ScaleHeight     =   405
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   3960
      Width           =   495
      Begin VB.Image imgMenu 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         Picture         =   "ucPacsHelper.ctx":0D6E
         Top             =   30
         Width           =   270
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   360
      ScaleHeight     =   9495
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   135
         Left            =   0
         TabIndex        =   1
         Top             =   2535
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   238
         BackColor       =   12632256
         MousePointer    =   7
         SplitType       =   0
         SplitLevel      =   3
         Con1MinSize     =   1000
         Con2MinSize     =   2000
         Control1Name    =   "picVideoContainer"
         Control2Name    =   "picHelperContainer"
      End
      Begin VB.PictureBox picVideoContainer 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   0
         ScaleHeight     =   2535
         ScaleWidth      =   5415
         TabIndex        =   4
         Top             =   0
         Width           =   5415
      End
      Begin VB.PictureBox picHelperContainer 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   6825
         Left            =   0
         ScaleHeight     =   6825
         ScaleWidth      =   5415
         TabIndex        =   2
         Top             =   2670
         Width           =   5415
         Begin zl9PACSWork.ucCacheImages ucCache 
            Height          =   4935
            Left            =   1800
            TabIndex        =   7
            Top             =   1320
            Visible         =   0   'False
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   8705
         End
         Begin zl9PACSWork.ucReportHistory ucHistory 
            Height          =   4695
            Left            =   1320
            TabIndex        =   6
            Top             =   960
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   8281
         End
         Begin zl9PACSWork.ucBgImgViewer ucImages 
            Height          =   4575
            Left            =   840
            TabIndex        =   5
            Top             =   600
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   8070
         End
         Begin XtremeSuiteControls.TabControl tabSelect 
            Height          =   375
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1815
            _Version        =   589884
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   64
         End
         Begin zl9PACSWork.ucReportSegment ucWord 
            Height          =   4335
            Left            =   480
            TabIndex        =   9
            Top             =   360
            Visible         =   0   'False
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   7646
         End
         Begin VB.PictureBox picTemp 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   4215
            Left            =   120
            ScaleHeight     =   4215
            ScaleWidth      =   3015
            TabIndex        =   8
            Top             =   120
            Width           =   3015
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   6120
      Top             =   2160
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "ucPacsHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const C_MODULE_NAME = "ucPacsHelper"


Private Const CON_TAB_TAG_图像 As String = "图像"
Private Const CON_TAB_TAG_报告图 As String = "报告图"
Private Const CON_TAB_TAG_词句 As String = "词句"
Private Const CON_TAB_TAG_历史 As String = "历史"
Private Const CON_TAB_TAG_缓存 As String = "缓存"



Private Const conMenu_Helper_Refresh = 8140         '刷新


'图像相关
Private Const conMenu_Helper_AddToReport = 8141     '加入报告图
Private Const conMenu_Helper_ImageProcess = 8142    '图像处理
Private Const conMenu_Helper_BigImageShow_Move = 8143    '移动显示大图
Private Const conMenu_Helper_BigImageShow_Click = 8144    '单击显示大图
Private Const conMenu_Helper_BigImageShow_Delay = 8145    '延迟关闭大图

Private Const conMenu_Helper_SelAll = 8146          '全选
Private Const conMenu_Helper_DelOper = 8147         '删除操作

Private Const conMenu_Helper_Import = 8148          '导入
Private Const conMenu_Helper_Export = 8149          '导出

Private Const conMenu_Helper_SendStudy = 8150       '发送到检查
Private Const conMenu_Helper_SendCache = 8151       '发送到缓存

Private Const conMenu_Helper_SplitPage = 8152
Private Const conMenu_Helper_ReDo = 8153        '重新尝试   ’只针对处理失败的图像
Private Const conMenu_Helper_ReDown = 8154      '重新下载
Private Const conMenu_Helper_ReUp = 8155        '重新上传

Private Const conMenu_Helper_OpenImgPos = 8156     '打开图像位置


'检查历史相关
Private Const conMenu_Helper_ImgViewer = 8157           '影像观片(&S)
Private Const conMenu_Helper_ImgContrast = 8158         '观片对比(&E)
Private Const conMenu_Helper_ReportOpen = 8159          '报告打开
Private Const conMenu_Helper_Analysis = 8160      '综合分析

Private Const conMenu_Helper_ViewReportImage = 8161          '查看报告图
Private Const conMenu_Helper_ViewReportContext = 8162        '查看报告内容
Private Const conMenu_Helper_WriteReport = 8163         '写入报告
Private Const conMenu_Helper_LinkViewer = 8164        '联动查看
Private Const conMenu_Helper_CloseViewer = 8165

Private Const conMenu_Helper_RelateCfg = 8166          '相关设置
Private Const conMenu_Helper_ThisTime = 8167            '本次相关
Private Const conMenu_Helper_OtherDept = 8168           '他科检查
Private Const conMenu_Helper_AutoLine = 8169            '自动换行

Private Const conMenu_Helper_DateRange = 8170          '日期范围
Private Const conMenu_Helper_OneMonth = 8171            '一个月
Private Const conMenu_Helper_TwoMonth = 8172            '二个月
Private Const conMenu_Helper_ThreeMonth = 8173          '三个月
Private Const conMenu_Helper_HalfYear = 8174            '半年
Private Const conMenu_Helper_OneYear = 8175             '一年
Private Const conMenu_Helper_TwoYear = 8176             '两年
Private Const conMenu_Helper_ThreeYear = 8177           '三年
Private Const conMenu_Helper_DateUn = 8178              '不限日期
Private Const conMenu_Helper_DateCus = 8179             '自定日期


'报告词句相关
Private Const conMenu_Helper_DirectWrite = 8180         '直接写入
Private Const conMenu_Helper_EditWrite = 8181           '编辑写入

Private Const conMenu_Helper_FullSave = 8182            '全套写入
Private Const conMenu_Helper_NewWord = 8183           '新增词句
Private Const conMenu_Helper_ModWord = 8184             '修改词句
Private Const conMenu_Helper_DelWord = 8185             '删除词句

Private Const conMenu_Helper_AutoHide = 8186            '自动隐藏
Private Const conMenu_Helper_DblWrite = 8187
Private Const conMenu_Helper_ExpandLevel = 8188        '展开层级
Private Const conMenu_Helper_OneLevel = 8189         '一级
Private Const conMenu_Helper_TwoLevel = 8190           '二级
Private Const conMenu_Helper_ThreeLevel = 8191         '三级
Private Const conMenu_Helper_AllLevel = 8192           '所有


Private Const conMenu_Helper_Log = 8300    '日志
 

Private Type DlgFileInfo
    iCount As Long
    sPath As String
    sFIle() As String
End Type


Private mObjNotify As IEventNotify
Private mobjEmbedVideo As Object
Private mobjLinkEditor As ucReportEditor
Private mobjSel As Object


Private WithEvents mobjImageProcessV2 As frmImageProcessV2
Attribute mobjImageProcessV2.VB_VarHelpID = -1

Private mblnIsEmbedVideoArea As Boolean     '是否嵌入视频采集区域
Private mblnImgAscOrder As Boolean
Private mblnAllowEmbedVideo As Boolean  '是否允许嵌入视频采集

Private mlngModuleNo As Long
Private mlngDeptID As Long
Private mstrPrivs As String
Private mstrGrantDeptIds As String
Private mstrParentName As String
 
Private mobjStudyInfo As clsStudyInfo
Private mlngReleationImgAdvice As Long

Private mlngFileID As Long
Private mblnBgImgTrans As Boolean         '后台图像处理
Private mblnIsTabIniting As Boolean
Private mblnIsValid As Boolean
Private mobjMainVideo As Object

Private mblnMoveBigImageShow As Boolean
Private mblnClickBigImageShow As Boolean
Private mblnDelayCloseImage As Boolean
Private mlngBigImageIndex As Long
Private mlngStartMoveTime As Long

Private mblnIgnoreResult As Boolean
    
Private mstrReportImageUids As String       '保存作为报告图的图像UID
Private mblnAllowWrite As Boolean
Private mblnIsProcessing As Boolean
Private mlngReleationImgDays As Long
Private mlngImageDBClickOper As Long    '图像双击操作方式

'词句相关事件
Public Event OnWordRequestState(ByRef lngOutlineType As TOutlineType, _
                            ByRef str所见内容 As String, ByRef str意见内容 As String, ByRef str建议内容 As String)
    
Public Event OnWordSendContext(ByVal strFreeText As String, _
                            ByVal str所见内容 As String, ByVal str意见内容 As String, ByVal str建议内容 As String)
                                       
Public Event OnTabChanged(ByVal strTabName As String)

Public Event OnLinkHistoryView(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, ByVal blnIsDBClick As Boolean)
                            
Public Event OnDockHideClick()
Public Event OnDockAttachClick()



Property Get IsSyncWordFragment() As Boolean
    IsSyncWordFragment = ucWord.IsSyncWordFragment
End Property


Property Get ImgCount() As Long
    ImgCount = ucImages.ImgCount
End Property

Property Get EmbedVideo() As Object
    Set EmbedVideo = mobjEmbedVideo
End Property
                            
Property Get Processing() As Boolean
    Processing = mblnIsProcessing
End Property

Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property


Property Get HideButtonEnable() As Boolean
    HideButtonEnable = cmdHide.Enabled
End Property

Property Let HideButtonEnable(ByVal value As Boolean)
    cmdHide.Enabled = value
End Property



Property Get IsValid() As Boolean
    IsValid = mblnIsValid
End Property
 
 
Property Get SelTabName() As String
    SelTabName = tabSelect.Selected.tag
End Property


'连接的报告编辑器
Property Get LinkEditor() As Object
    Set LinkEditor = mobjLinkEditor
End Property

Property Set LinkEditor(value As Object)
    Set mobjLinkEditor = value
End Property


'主视频窗口
Property Get MainVideoWindow() As Object
    Set MainVideoWindow = mobjMainVideo
End Property

Property Set MainVideoWindow(value As Object)
    Set mobjMainVideo = value
End Property


Property Get AllowLinkerViewer() As Boolean
    AllowLinkerViewer = ucHistory.AllowLinkViewer
End Property


Property Let AllowLinkerViewer(ByVal value As Boolean)
    ucHistory.AllowLinkViewer = value
End Property


Property Get TabEnable(ByVal strTabName As String) As Boolean
    Dim i As Long
    
    For i = 1 To tabSelect.ItemCount
        If UCase(tabSelect(i).tag) = UCase(strTabName) Then
            TabEnable = tabSelect(i).Enabled
            Exit Property
        End If
    Next
End Property

Property Let TabEnable(ByVal strTabName As String, ByVal value As Boolean)
    Dim i As Long
    Dim j As Long
    
    For i = 0 To tabSelect.ItemCount - 1
        If UCase(tabSelect(i).tag) = UCase(strTabName) Then
            tabSelect(i).Enabled = value
            
            If value = False And tabSelect(i).Selected Then
                For j = 0 To tabSelect.ItemCount - 1
                    If tabSelect(j).Enabled Then
                        tabSelect(j).Selected = True
                        Exit Property
                    End If
                Next
            End If
            
            Exit Property
        End If
    Next
End Property


Property Get TabVisible(ByVal strTabName As String) As Boolean
    Dim i As Long
    
    For i = 1 To tabSelect.ItemCount
        If UCase(tabSelect(i).tag) = UCase(strTabName) Then
            TabVisible = tabSelect(i).Visible
            Exit Property
        End If
    Next
End Property

Property Let TabVisible(ByVal strTabName As String, ByVal value As Boolean)
    Dim i As Long
    Dim j As Long
    
    For i = 0 To tabSelect.ItemCount - 1
        If UCase(tabSelect(i).tag) = UCase(strTabName) Then
            tabSelect(i).Visible = value
            
            If value = False And tabSelect(i).Selected Then
                For j = 0 To tabSelect.ItemCount - 1
                    If tabSelect(j).Visible Then
                        tabSelect(j).Selected = True
                        Exit Property
                    End If
                Next
            End If
            
            Exit Property
        End If
    Next
End Property

'是否允许写入
Property Get AllowWrite() As Boolean
    AllowWrite = mblnAllowWrite
End Property

Property Let AllowWrite(ByVal value As Boolean)
    mblnAllowWrite = value
    
    If mblnAllowWrite = False Then
        cmdAdd.Enabled = mblnAllowWrite
    Else
        cmdAdd.Enabled = IIf(tabSelect.Selected.tag <> CON_TAB_TAG_缓存, True, False)
    End If
    
    ucHistory.AllowWrite = value
End Property



'医嘱ID
Property Get AdviceId() As Long
        AdviceId = mobjStudyInfo.lngAdviceId
End Property

'Property Let AdviceId(ByVal value As Long)
'    mlngAdviceID = value
'End Property



'是否嵌入视频采集
Property Get IsEmbedVideoArea() As Boolean
    IsEmbedVideoArea = mblnIsEmbedVideoArea
End Property


'是否允许嵌入视频采集
Property Get AllowEmbedVideo() As Boolean
    AllowEmbedVideo = mblnAllowEmbedVideo
End Property

Property Let AllowEmbedVideo(ByVal value As Boolean)
    mblnAllowEmbedVideo = value
End Property


Property Get IsStudying() As Boolean
    IsStudying = IIf(mobjStudyInfo.intStep < 6 And mobjStudyInfo.intStep > 1, True, False)
End Property


Private Function HintError(objErr As ErrObject, ByVal strMethodName As String, _
    Optional ByVal blnIsDataErr As Boolean = True) As Long
    If mObjNotify Is Nothing Then Exit Function
    
    If blnIsDataErr Then
        HintError = mObjNotify.PrintErr(objErr, infDataErr, GetRootHwnd, C_MODULE_NAME, strMethodName)
    Else
        HintError = mObjNotify.PrintErr(objErr, infNormalErr, GetRootHwnd, C_MODULE_NAME, strMethodName)
    End If
End Function

Private Function HintMsg(ByVal strMsg As String, ByVal strMethodName As String, _
    Optional ByVal lngMsgType As Long = infHint) As Long
        HintMsg = mObjNotify.PrintInfo(strMsg, lngMsgType, GetRootHwnd, C_MODULE_NAME, strMethodName)
End Function


Public Sub HideEmbedVideo()
'隐藏嵌入式视频采集窗口
    mblnIsEmbedVideoArea = False
    
    picVideoContainer.Visible = mblnIsEmbedVideoArea
    ucSplitter1.Visible = mblnIsEmbedVideoArea
    
    '如果已经存在视频采集,则可能需要将视频采集恢复到主窗口
    Call picBack_Resize
End Sub


Public Function ShowEmbedVideo(objCapLinker As Object, Optional ByVal blnIsForce As Boolean = False) As Boolean
'嵌入视频
'blnIsForce:是否强制嵌入视频采集，不判断视频所在的根窗口是否相同,主要用于影像采集和检查报告模块页之前的视频切换
    Dim objCapHelper As ICapHelper
    Dim blnAfterOrLock As Boolean
    
    ShowEmbedVideo = False
    mblnIsEmbedVideoArea = False
    blnAfterOrLock = False
    
     
    '如果不允许嵌入视频采集，则退出
    If mblnAllowEmbedVideo = False Then
        Call HideEmbedVideo
        Exit Function
    End If
    
    If Not mobjEmbedVideo Is Nothing Then
        blnAfterOrLock = IIf(mobjEmbedVideo.isLock Or mobjEmbedVideo.IsAfter, True, False)
    End If
    
    Set objCapHelper = objCapLinker
    
'    If objCapHelper.IsAllowCapture = False And blnAfterOrLock = False Then
'        '当前状态不允许采集，且没有开启后台和锁定采集情况下，隐藏嵌入式视频窗口
'        If picVideoContainer.Visible = False Then Exit Function
'        Call HideEmbedVideo
'        Exit Function
'    Else
        '如果已经显示了嵌入式视频窗口，则直接退出
        If Not mobjEmbedVideo Is Nothing Then
            If picVideoContainer.Visible And (GetAncestor(GetAncestor(mobjEmbedVideo.VideoHwnd, GA_PARENT), GA_PARENT) = picVideoContainer.hwnd) Then
                Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False, True)
                
                mblnIsEmbedVideoArea = True
                ShowEmbedVideo = True
                Exit Function
            End If
        Else
            Set mobjEmbedVideo = New clsPacsCaptureV2
            Call mobjEmbedVideo.zlInitModule(gcnOracle, objCapLinker, glngSys, mlngModuleNo, mstrPrivs, mlngDeptID, hwnd, True)
        End If

    '判断是否弹出了独立的视频采集窗口，如果弹出，则不允许嵌入视频
    If mobjEmbedVideo.VideoDockState Then
        Call HideEmbedVideo
        Exit Function
    End If
    
    mblnIsEmbedVideoArea = True
    
    picVideoContainer.Visible = True
    ucSplitter1.Visible = True
    
    Call picBack_Resize
    Call picVideoContainer_Resize

    
    '如果根窗口是相同窗口，则不需要重新设置嵌入
    If blnIsForce = False And GetAncestor(mobjEmbedVideo.VideoHwnd, GA_ROOT) = GetAncestor(picVideoContainer.hwnd, GA_ROOT) Then
        If GetAncestor(mobjEmbedVideo.ContainerHwnd, GA_ROOT) = GetAncestor(picVideoContainer.hwnd, GA_ROOT) Then
            '更新是否只读状态
            If mobjEmbedVideo.ContainerHwnd = GetAncestor(mobjEmbedVideo.VideoHwnd, GA_PARENT) Then
               
            
                Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False, True)
                ShowEmbedVideo = True
                Exit Function
            End If
        End If
    End If
    
    SetParent mobjEmbedVideo.ContainerHwnd, picVideoContainer.hwnd
    
    ShowWindow mobjEmbedVideo.ContainerHwnd, 1
    
    Call mobjEmbedVideo.zlRefreshVideoWindow
    
    If mobjStudyInfo Is Nothing Then
        '初始化方法中调用嵌入视频显示时，mobjStudyInfo为nothing
        Call mobjEmbedVideo.zlRestoreWindow(True, False)
    Else
        Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False)
    End If
    ShowEmbedVideo = True
End Function


Public Sub LocateTab(ByVal strTabName As String)
'定位指定的tab页
    Dim i As Long
    
    For i = 0 To tabSelect.ItemCount - 1
        If UCase(tabSelect(i).tag) = UCase(strTabName) Then
            If tabSelect(i).Visible Then tabSelect(i).Selected = True
            Exit Sub
        End If
    Next
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    FontSize = bytFontSize
    
    picBack.FontSize = bytFontSize
    picVideoContainer.FontSize = bytFontSize
    picHelperContainer.FontSize = bytFontSize
    picTemp.FontSize = bytFontSize
    
    Call ucWord.SetFontSize(bytFontSize)
    Call ucHistory.SetFontSize(bytFontSize)
    Call ucCache.SetFontSize(bytFontSize)
    
    Set tabSelect.PaintManager.Font = Font
    
    '字体改变后，需要使用该属性刷新界面显示
    tabSelect.PaintManager.Layout = xtpTabLayoutAutoSize
End Sub

Public Sub Init(objNotify As IEventNotify, ByVal lngMudleNo As Long, ByVal lngDeptId As Long, _
    ByVal strPrivs As String, Optional ByVal blnIsForce As Boolean = False)
'初始化
'lngMudleNo:模块号
'lngDeptId：当前科室ID
'strGrantDepts:授权科室ID
'
    Set mObjNotify = objNotify
    
    mlngModuleNo = lngMudleNo
    mlngDeptID = lngDeptId
    mstrPrivs = strPrivs
    mstrParentName = Parent.Name
    
    Call InitTab
     
    Call InitPar
    
    Call ucImages.Init
    Call ucWord.Init(lngMudleNo, lngDeptId, blnIsForce)
    Call ucHistory.Init(lngMudleNo, lngDeptId, mstrPrivs, blnIsForce)
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbrMain.VisualTheme = xtpThemeWhidbey

    mblnIsValid = True
End Sub



Public Function GetFileFormatId(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean) As Long
'获取检查对应的诊疗单据格式ID
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetFileFormatId = 0
    
    strSQL = "Select l.病人来源, a.病历文件id" & vbNewLine & _
            " From 病人医嘱记录 l, 病历单据应用 a" & vbNewLine & _
            " Where l.诊疗项目id = a.诊疗项目id(+) And a.应用场合(+) = Decode(l.病人来源, 2, 2, 4 ,4, 1) And l.Id = [1]"
            
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询单据格式", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetFileFormatId = Val(nvl(rsData!病历文件id))
    
End Function


Public Sub zlRefresh(objStudyInfo As clsStudyInfo, ByVal lngFileId As Long, Optional ByVal blnIsForceRefresh As Boolean = False)
'刷新
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errhandle
    
    If mobjSel Is Nothing Then Exit Sub
    
    If Not objStudyInfo Is Nothing And Not mobjStudyInfo Is Nothing Then
        If mobjStudyInfo.IsEquals(objStudyInfo) And blnIsForceRefresh = False Then Exit Sub
    End If
    
    mblnIsProcessing = True
    
    If Not mobjImageProcessV2 Is Nothing Then
        '判断正在处理的图像是否保存
        Unload mobjImageProcessV2
    End If
    
    Set mobjStudyInfo = objStudyInfo
     
    mlngFileID = lngFileId
    mlngReleationImgAdvice = 0
    
    If mlngFileID = 0 Then mlngFileID = GetFileFormatId(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
    
    Call RefreshHelperComponent(mobjStudyInfo.lngAdviceId, mobjStudyInfo.strStudyUID, mlngFileID, mobjStudyInfo.blnMoved)
    
    If mblnIsEmbedVideoArea Then
        If Not mobjEmbedVideo Is Nothing Then
            Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False)
        End If
    End If
    
    mblnIsProcessing = False
Exit Sub
errhandle:
    mblnIsProcessing = False
    If HintError(err, "zlRefresh", False) = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshHelperComponent(ByVal lngAdviceId As Long, ByVal strSutdyUid As String, ByVal lngFileId As Long, ByVal blnMoved As Boolean)
'刷新帮助组件
    If mobjSel Is Nothing Then Exit Sub
    Select Case mobjSel.hwnd
        Case ucImages.hwnd  '图像
             ucImages.ClearAll
             
            '判断检查是否已经有图像
            If Len(strSutdyUid) > 0 Then Call LoadExamImages(lngAdviceId, blnMoved)
            
            If ucImages.ImgCount <= 0 Then
                If mlngReleationImgDays > 0 Then
                    '判断是否进行关联图像读取
                    If mlngReleationImgAdvice = 0 Then
                        mlngReleationImgAdvice = GetReleationImageAdvice(lngAdviceId)
                    End If
                    
                    If mlngReleationImgAdvice > 0 Then
                        Call LoadExamImages(mlngReleationImgAdvice, False)
                    End If
                Else
                    mstrReportImageUids = ""
                    Call ucImages.ClearAll
                End If
            End If
            
        Case ucWord.hwnd    '词句
            Call ucWord.Refresh(lngAdviceId, lngFileId)

        Case ucHistory.hwnd '历史
            ucHistory.AllowWrite = mblnAllowWrite And (IsStudying Or (CheckPopedom(mstrPrivs, "补录报告") And mobjStudyInfo.intStep > 5 And mobjStudyInfo.strReportDoctor = ""))
            Call ucHistory.Refresh(lngAdviceId)
            
        Case ucCache.hwnd   '缓存
            Call ucCache.Refresh
            
    End Select
End Sub

Private Function GetReleationImageAdvice(ByVal lngAdviceId As Long) As Long
'打开关联在线检查图像
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rsReturn As ADODB.Recordset
    Dim dtStudy As Date
    
    GetReleationImageAdvice = 0
    
    strSQL = "select a.执行部门id, a.报到时间, b.病人id,c.关联id,c.影像类别 " & _
            " from 病人医嘱发送 a, 病人医嘱记录 b,影像检查记录 c" & _
            " where a.医嘱id=b.id and a.医嘱id=c.医嘱id and a.发送号=c.发送号 and a.医嘱id=[1]"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查关联信息", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    '打开相同科室下，最近影像类别相同的检查，主要针对同一患者不同医嘱相同类别的不同部位在同一设备上做检查的情况
    strSQL = "Select Distinct * from (" & vbCrLf & _
            " select b.医嘱ID, b.检查号,b.姓名,b.性别,b.年龄,b.影像类别, a.医嘱内容, b.检查uid,b.位置一,b.位置二 " & _
            " from 病人医嘱记录 a, 影像检查记录 b" & vbCrLf & _
            " Where a.ID = b.医嘱ID And a.病人ID = [1] And b.执行科室id = [3]" & vbCrLf & _
            "       and b.接收日期 between [4] and [5] " & vbCrLf & _
            "       and b.检查UID is not null and b.医嘱ID<>[6] and b.影像类别=[7] " & vbCrLf & _
            " Union All " & vbCrLf & _
            " select a.医嘱ID, a.检查号,a.姓名,a.性别,a.年龄,a.影像类别, b.医嘱内容,a.检查uid,a.位置一,a.位置二 " & _
            " from 影像检查记录 a , 病人医嘱记录 b " & vbCrLf & _
            " Where a.医嘱ID=b.Id and a.关联id = [2] And a.执行科室id = [3]" & vbCrLf & _
            "   and a.接收日期 between [4] and [5] " & vbCrLf & vbCrLf & _
            " and a.检查UID is not null and a.医嘱ID<>[6] and a.影像类别=[7] " & vbCrLf & _
            " )"
            
    dtStudy = CDate(Format(nvl(rsData!报到时间, 0), "yyyy-mm-dd 00:00"))
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询关联检查数据", _
                            Val(nvl(rsData!病人ID)), Val(nvl(rsData!关联ID)), Val(nvl(rsData!执行部门ID)), _
                            CDate(Format(dtStudy - mlngReleationImgDays, "yyyy-mm-dd 00:00:00")), _
                            CDate(Format(dtStudy + mlngReleationImgDays, "yyyy-mm-dd 23:59:59")), _
                            lngAdviceId, nvl(rsData!影像类别)) '常数7表示查询的天数范围
                            
    If rsData.RecordCount <= 0 Then Exit Function
    
    If rsData.RecordCount = 1 Then
        If HintMsg("当前检查中未发现相关图像，是否从如下的最近检查提取?" & vbCrLf & _
                                "    检查号：" & nvl(rsData!检查号) & vbCrLf & _
                                "    姓名：" & nvl(rsData!姓名) & vbCrLf & _
                                "    性别：" & nvl(rsData!性别) & vbCrLf & _
                                "    年龄：" & nvl(rsData!年龄) & vbCrLf & _
                                "    " & nvl(rsData!医嘱内容), "GetReleationImageAdvice", vbYesNo) = vbNo Then
            GetReleationImageAdvice = -1
            Exit Function
        End If
        
        GetReleationImageAdvice = Val(nvl(rsData!医嘱ID))
    Else
        If HintMsg("当前检查中未发现相关图像，是否从最近的 [" & rsData.RecordCount & "] 条相关检查中提取图像？", "GetReleationImageAdvice", vbYesNo) = vbNo Then
            GetReleationImageAdvice = -1
            Exit Function
        End If

        If FS.ShowRecSelect(mObjNotify.Owner, cmdMenu, rsData, rsReturn, True, "医嘱ID,位置一,位置二,检查UID") Then
            GetReleationImageAdvice = Val(nvl(rsReturn!医嘱ID))
        End If
    End If
    
End Function

Public Sub ClearReportImgState(Optional ByVal strImgKey As String = "")
'清除报告图状态
    If Len(strImgKey) > 0 Then
        Call ucImages.ImgDrawHint(strImgKey, "", "报")
    Else
        Call ucImages.ClearDrawHint("报")
    End If
End Sub
 

Public Sub SyncReportImgState(ByVal strImgs As String)
'同步报告图状态
    Call ucImages.SyncDrawHint(strImgs, "报", "报")
    mstrReportImageUids = strImgs
End Sub

Public Function GetLayoutStr() As String
'返回格式字符串[Key=picturebox1.width:20;picturebox1.height:30;]
    Dim strPros As String
    strPros = "[KEY=HELPER@" & _
                        GetProFmt("PICVIDEO.HEIGHT", picVideoContainer.Height) & _
                        GetProFmt("PICHELPER.HEIGHT", picHelperContainer.Height) & _
                        ";]"
                        
    GetLayoutStr = strPros & ucWord.GetLayoutStr & ucHistory.GetLayoutStr
End Function

Public Sub SetLayout(ByVal strLayout As String)
    Dim strPros As String
    Dim strPro As String

    If Len(strLayout) <= 0 Then Exit Sub
    
    Call ucWord.SetLayout(strLayout)
    Call ucHistory.SetLayout(strLayout)
    
    
    strPros = GetPros(strLayout, "HELPER")
    
    strPro = GetProValue(strPros, "PICVIDEO.HEIGHT")
    If Val(strPro) > 0 Then picVideoContainer.Height = Val(strPro)
    
    strPro = GetProValue(strPro, "PICHELPER.HEIGHT")
    If Val(strPro) > 0 Then picHelperContainer.Height = Val(strPro)
End Sub

Private Sub LoadExamImages(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean)
'载入检查图像
    Dim objImgInf As New clsBgImgInfo
    Dim strStudyUID As String
    Dim rsData As ADODB.Recordset
    Dim strLocalPath As String
    Dim i As Long
    Dim strResult As String
    
    mstrReportImageUids = ""
    
    ucImages.ClearAll
    
    If lngAdviceId <= 0 Then
        Exit Sub
    End If
    
    objImgInf.PatientName = mobjStudyInfo.strPatientName
    objImgInf.ImgCommand = icDownload
    objImgInf.AdviceId = lngAdviceId
    objImgInf.Format = ifDcm
    
    
    strResult = ResetStorageDevice(lngAdviceId, objImgInf, blnMoved)
    
    If Len(strResult) > 0 Then
        HintMsg strResult, "ResetStorageDevice", vbOKOnly
        Exit Sub
    End If
    
    strStudyUID = objImgInf.StudyUID
    If Len(strStudyUID) <= 0 Then Exit Sub
    
    Set rsData = GetExamImgData(strStudyUID, blnMoved)
    
    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub
    
    '构造载入图像
    Call ucImages.EraseImgData
    
    i = 1
    While Not rsData.EOF
        objImgInf.Key = nvl(rsData!图像UID)
        
        '保存报告图的UID
        If mobjStudyInfo.lngAdviceId = lngAdviceId Then
            If nvl(rsData!报告图) <> "" Then
                mstrReportImageUids = mstrReportImageUids & ";" & objImgInf.Key & ";"
                
                If Val(rsData!报告图) = 2 Then
                    objImgInf.DrawHint = "" '"◆"
                Else
                    '1和0的状态
                    objImgInf.DrawHint = "报"
                End If
            Else
                objImgInf.DrawHint = ""
            End If
        End If
        
        objImgInf.FtpFile = nvl(rsData!图像UID) & IIf(nvl(rsData!图像描述) = "REPIMG", ".jpg", "")
        objImgInf.Filename = nvl(rsData!图像UID) & IIf(nvl(rsData!图像描述) = "REPIMG", ".jpg", "")
        objImgInf.AdviceDes = IIf(nvl(rsData!图像描述) = "REPIMG", "REPIMG", "")
        
        objImgInf.IsBackGround = mblnBgImgTrans
        objImgInf.JpgConvert = False
        
        If Val(nvl(rsData!动态图)) = ImgTag Then objImgInf.Format = ifDcm
        If Val(nvl(rsData!动态图)) = VIDEOTAG Then objImgInf.Format = ifAvi
        If Val(nvl(rsData!动态图)) = AUDIOTAG Then objImgInf.Format = ifWav
        If Val(nvl(rsData!动态图)) = BMPTAG Then objImgInf.Format = ifBmp
        
        objImgInf.SeriesNoTag = nvl(rsData!序列号, "*")
        objImgInf.ImageOrder = nvl(rsData!图像号, 0)
        
        Call ucImages.ConstructionImgData(objImgInf.CopyNew())
        
        i = i + 1
        
        Call rsData.MoveNext
    Wend
    
    Call ucImages.Refresh
End Sub

Public Function GetRootHwnd() As Long
    GetRootHwnd = GetAncestor(hwnd, GA_ROOT)
End Function


Private Function GetExamImgData(ByVal strStudyUID As String, ByVal blnMoved As Boolean) As ADODB.Recordset
'获取检查图像数据
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim blnReadAllImg As Boolean
    
On Error GoTo errhandle
    Set GetExamImgData = Nothing
    
    blnReadAllImg = True
    If mlngModuleNo = G_LNG_PACSSTATION_MODULE Then
            '影像医技的报告图只能根据影像检查记录中的报告图字段进行加载
        strSQL = "  Select 1 as 序列号, Replace(Trim(B.Column_Value),'.jpg','') as 图像UID, rownum as 图像号, 'REPIMG' as 图像描述, 2 as 报告图, 5 as 动态图, " & _
                        " null as 编码名称, null as 采集时间, null as 录制长度 " & _
                        " From 影像检查记录 A, Table(Cast(f_Str2list(Replace(A.报告图象,';',',')) As zlTools.t_Strlist)) B " & _
                        " Where 检查UID = [1]"
                            
        If blnMoved Then
            strSQL = Replace(strSQL, "影像检查记录", "影像检查记录")
        End If
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告图", strStudyUID)
        If rsData.RecordCount > 0 Then
            blnReadAllImg = False
        Else
            '如果没有报告图，则不进行图像加载
            Exit Function
        End If
    End If
    
    If blnReadAllImg Then
        strSQL = "Select B.序列号,A.图像UID, A.图像号,A.图像描述, A.报告图, A.动态图,A.编码名称,A.采集时间, A.录制长度 " & _
            " From 影像检查图象 A,影像检查序列 B" & _
            " Where A.序列UID=B.序列UID And B.检查UID=[1]"
     
        If mobjStudyInfo.blnMoved Then
            strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
            strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        End If
    
        If mblnImgAscOrder Then
            strSQL = strSQL & " order by B.序列号, A.采集时间, 图像号"
        Else
            strSQL = strSQL & " order by B.序列号 Desc, A.采集时间 Desc, 图像号 Desc"
        End If
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查图像数据", strStudyUID)
    End If

    Set GetExamImgData = rsData
Exit Function
errhandle:
    If HintError(err, "GetExamImgData") Then Resume
End Function


Public Sub SyncCaptureImage(objImgInfo As clsBgImgInfo, Optional ByVal blnIsProxyTrans As Boolean = False)
'同步采集图像
    Dim blnSyncStudyUID As Boolean
    mblnIsProcessing = True
On Error GoTo errhandle
    blnSyncStudyUID = IIf(ucImages.ImgCount <= 0 Or mobjStudyInfo.strStudyUID = "", True, False)
    
    If blnIsProxyTrans = False Then
        Call ucImages.AddImg(objImgInfo)
    Else
        Call ucImages.ProxyTransfer(objImgInfo)
    End If
    
    If blnSyncStudyUID Then
        mobjStudyInfo.strStudyUID = objImgInfo.StudyUID
    End If
    
    mblnIsProcessing = False
    
Exit Sub
errhandle:
    mblnIsProcessing = False
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

Public Sub SyncAfterCapture(objImg As Object, strAfterTag As String)
'同步后台采集
    Call ucCache.SyncAfterShow(objImg, strAfterTag)
End Sub


Public Sub SyncAfterTag(strAfterTag As String)
'同步后台标记
    Call ucCache.Refresh
End Sub


Public Sub SyncOutline(ByVal strOutlineName As String)
'同步提纲
     Call ucWord.SyncOutline(strOutlineName)
End Sub


Private Sub InitTab()
On Error GoTo errH
    Dim i As Integer
    Dim iCount As Integer
    Dim strName() As String
     
    mblnIsTabIniting = True
    
    If tabSelect.ItemCount >= 1 Then
        mblnIsTabIniting = False
        Exit Sub
    End If
   
    With tabSelect
    
    
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ShowIcons = False ' True
        .PaintManager.Layout = xtpTabLayoutAutoSize ' xtpTabLayoutFixed ' ' xtpTabLayoutAutoSize
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionLeft
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.ColorSet.ButtonNormal = &HE0E0E0
        .PaintManager.HeaderMargin.Left = 110 '135
        .PaintManager.HeaderMargin.Top = 0
        .PaintManager.HeaderMargin.Right = 2
        .PaintManager.HeaderMargin.Bottom = 0
        .PaintManager.ButtonMargin.Left = 0
        .PaintManager.ButtonMargin.Top = 0
        .PaintManager.ButtonMargin.Right = 2
        .PaintManager.ButtonMargin.Bottom = 3
 
        
        .RemoveAll
        
        
        If mlngModuleNo = G_LNG_PACSSTATION_MODULE Then
            .InsertItem 1, CON_TAB_TAG_报告图, picTemp.hwnd, 0
        Else
            .InsertItem 1, CON_TAB_TAG_图像, picTemp.hwnd, 0
        End If
        
        .Item(0).tag = CON_TAB_TAG_图像
         
        .InsertItem 2, CON_TAB_TAG_词句, picTemp.hwnd, 0
        .Item(1).tag = CON_TAB_TAG_词句

        .InsertItem 3, CON_TAB_TAG_历史, picTemp.hwnd, 0
        .Item(2).tag = CON_TAB_TAG_历史

        If mlngModuleNo = G_LNG_VIDEOSTATION_MODULE Then
            .InsertItem 4, CON_TAB_TAG_缓存, picTemp.hwnd, 0
            .Item(3).tag = CON_TAB_TAG_缓存
        End If

        .Item(0).Selected = True
        Set mobjSel = ucImages
        
        '默认显示为图像采集界面
        SetParent mobjSel.hwnd, picTemp.hwnd
        mobjSel.Visible = True
    End With
    
 
    
    mblnIsTabIniting = False
    Exit Sub
errH:
    mblnIsTabIniting = False
End Sub


Private Sub cmdAdd_Click()
On Error GoTo errhandle
    Call WriteData
Exit Sub
errhandle:
    HintError err, "cmdAdd_Click"
End Sub

Private Sub cmdAttach_Click()
On Error GoTo errhandle
    RaiseEvent OnDockAttachClick
Exit Sub
errhandle:
    HintError err, "cmdAttach_Click"
End Sub

Private Sub cmdDel_Click()
On Error GoTo errhandle
    Call DelData
Exit Sub
errhandle:
    HintError err, "cmdDel_Click"
End Sub
 

Private Sub cmdHide_Click()
On Error GoTo errhandle
    RaiseEvent OnDockHideClick
Exit Sub
errhandle:
    HintError err, "cmdHide_Click"
End Sub

Private Sub cmdMenu_Click()
On Error GoTo errhandle
    Call ShowPopupMenu(tabSelect.Selected.tag)
Exit Sub
errhandle:
End Sub


Private Sub cmdMenu_Resize()
On Error Resume Next
    imgMenu.Left = (cmdMenu.Width - imgMenu.Width) / 2
End Sub

Private Sub CmdRefresh_Click()
On Error GoTo errhandle
    Call RefreshData
Exit Sub
errhandle:
    HintError err, "cmdRefresh_Click"
End Sub

Private Sub imgMenu_Click()
    Call cmdMenu_Click
End Sub

Private Sub mobjImageProcessV2_OnSaveImage(ByVal emImageType As TImageType, dcmImage As DicomObjects.DicomImage)
'保存处理后的图像
    Dim strLineDeviceNo As String
    Dim strBackDeviceNo As String
    Dim objResult As clsBgImgInfo
    Dim strReportImgFile As String
    
On Error GoTo errhandle
    Select Case emImageType
        Case mtStudyImage   '保存到检查图
            strLineDeviceNo = GetDeptPara(mlngDeptID, "存储设备号")
            strBackDeviceNo = GetDeptPara(mlngDeptID, "备份设备号")
            
            Set objResult = SaveDicomImageToStudy(dcmImage, strLineDeviceNo, strBackDeviceNo)
            If Not objResult Is Nothing Then
                If mobjStudyInfo.strStudyUID = "" Or mobjStudyInfo.strStudyUID <> dcmImage.StudyUID Then
                    mobjStudyInfo.strStudyUID = dcmImage.StudyUID
                End If
            End If
        Case mtReportImage  '保存到报告图
            strLineDeviceNo = GetDeptPara(mlngDeptID, "存储设备号")
            strBackDeviceNo = GetDeptPara(mlngDeptID, "备份设备号")
            
            Set objResult = SaveDicomImageToStudy(dcmImage, strLineDeviceNo, strBackDeviceNo)
            If Not objResult Is Nothing Then
                If mobjStudyInfo.strStudyUID = "" Or mobjStudyInfo.strStudyUID <> dcmImage.StudyUID Then
                    mobjStudyInfo.strStudyUID = dcmImage.StudyUID
                End If
                
                strReportImgFile = objResult.FilePath & objResult.Filename & ".jpg"
                
                If FileExists(strReportImgFile) = False Then
                    Call dcmImage.FileExport(strReportImgFile, "BMP")
                End If
                
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_ADDIMG, , mobjStudyInfo.lngAdviceId, strReportImgFile)
            End If
    End Select
    
Exit Sub
errhandle:
    HintError err, "mobjImageProcessV2_OnSaveImage"
End Sub

Private Sub mobjImageProcessV2_OnUnload()
'    mlngBigImageIndex = 0
    Set mobjImageProcessV2 = Nothing
End Sub

Private Sub picBack_Resize()
On Error Resume Next
    If mblnIsEmbedVideoArea Then
        Call ucSplitter1.RePaint
    Else
        picHelperContainer.Move 0, 0, picBack.ScaleWidth, picBack.ScaleHeight
    End If
End Sub
 
Private Sub picHelperContainer_Resize()
On Error Resume Next
    tabSelect.Left = 0
    tabSelect.Top = 0
    tabSelect.Width = picHelperContainer.ScaleWidth
    tabSelect.Height = picHelperContainer.ScaleHeight
    
    cmdMenu.Left = 0
    cmdMenu.Top = picHelperContainer.Top
    cmdMenu.Width = ScaleWidth - ucImages.Width
    
    cmdRefresh.Left = 0
    cmdRefresh.Top = cmdMenu.Top + cmdMenu.Height
    cmdRefresh.Width = cmdMenu.Width
    
    cmdDel.Left = 0
    cmdDel.Top = cmdRefresh.Top + cmdRefresh.Height
    cmdDel.Width = cmdMenu.Width
    
    cmdAdd.Left = 0
    cmdAdd.Top = cmdDel.Top + cmdDel.Height
    cmdAdd.Width = cmdMenu.Width
    
    cmdHide.Left = 0
    cmdHide.Top = cmdAdd.Top + cmdAdd.Height
    cmdHide.Width = cmdMenu.Width
    
    cmdAttach.Left = 0
    cmdAttach.Top = cmdHide.Top + cmdHide.Height
    cmdAttach.Width = cmdMenu.Width
    
    
    imgMenu.Left = (cmdMenu.Width - imgMenu.Width) / 2
End Sub
 
Private Sub picTemp_Resize()
On Error Resume Next
    
    ucWord.Move 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight
    ucImages.Move 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight
    ucHistory.Move 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight
    ucCache.Move 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight
End Sub

 

 

Private Sub picVideoContainer_Resize()
On Error GoTo errhandle
    If mblnIsEmbedVideoArea = False Then Exit Sub
    If mobjEmbedVideo Is Nothing Then Exit Sub
    
    Call MoveWindow(mobjEmbedVideo.ContainerHwnd, 0, 0, _
            picVideoContainer.ScaleX(picVideoContainer.Width, vbTwips, vbPixels), _
            picVideoContainer.ScaleY(picVideoContainer.Height, vbTwips, vbPixels), 0)
    
    '显示窗口
    ShowWindow mobjEmbedVideo.ContainerHwnd, 1
      
Exit Sub
errhandle:

End Sub

Private Sub tabSelect_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errhandle

    If mblnIsTabIniting Then Exit Sub
    
    Select Case Item.tag
        Case CON_TAB_TAG_图像
            Set mobjSel = ucImages
            
        Case CON_TAB_TAG_词句
            Set mobjSel = ucWord

        Case CON_TAB_TAG_历史
            Set mobjSel = ucHistory
            
        Case CON_TAB_TAG_缓存
            Set mobjSel = ucCache
            
    End Select
    
    If ucWord.hwnd <> mobjSel.hwnd Then ucWord.Visible = False
    If ucImages.hwnd <> mobjSel.hwnd Then ucImages.Visible = False
    If ucHistory.hwnd <> mobjSel.hwnd Then ucHistory.Visible = False
    If ucCache.hwnd <> mobjSel.hwnd Then ucCache.Visible = False
    
    cmdDel.Enabled = IIf(Item.tag <> CON_TAB_TAG_历史, True, False)
    cmdAdd.Enabled = mblnAllowWrite And IIf(Item.tag <> CON_TAB_TAG_缓存, True, False)
    
    SetParent mobjSel.hwnd, picTemp.hwnd
    mobjSel.Visible = True
    
    Call RefreshHelperComponent(mobjStudyInfo.lngAdviceId, mobjStudyInfo.strStudyUID, mlngFileID, mobjStudyInfo.blnMoved)
    
    RaiseEvent OnTabChanged(Item.tag)
Exit Sub
errhandle:
    HintError err, "tabSelect_SelectedChanged", False
End Sub

Public Sub FreeVideo()
On Error GoTo errhandle:
    Set mobjSel = Nothing
    
    If Not mobjEmbedVideo Is Nothing Then
        mobjEmbedVideo.zlNotifyQuit
        
        ShowWindow mobjEmbedVideo.ContainerHwnd, 0
        SetParent mobjEmbedVideo.ContainerHwnd, 0
    End If
    
    Set mobjEmbedVideo = Nothing
    
Exit Sub
errhandle:

End Sub
 
Private Sub ucCache_OnDblClick()
On Error GoTo errhandle
    If ucCache.ImgCount <= 0 Then Exit Sub
    
    Call OpenImageProcess(True, True)
Exit Sub
errhandle:
    HintError err, "ucCache_OnDblClick", False
End Sub

Private Sub ucCache_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
'显示缓存右键菜单
On Error GoTo errhandle
    If Button = 2 Then Call ShowPopupMenu(tabSelect.Selected.tag)
Exit Sub
errhandle:
    HintError err, "ucCache_OnMouseUp", False
End Sub

Private Sub ucHistory_OnLinkView(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, ByVal blnIsDBClick As Boolean)
On Error GoTo errhandle
    RaiseEvent OnLinkHistoryView(lngAdviceId, blnMoved, blnIsDBClick)
Exit Sub
errhandle:
    HintError err, "ucCache_OnMouseUp", False
End Sub

Private Sub ucHistory_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'显示历史右键菜单
On Error GoTo errhandle
    If Button = 2 Then Call ShowPopupMenu(tabSelect.Selected.tag)
Exit Sub
errhandle:
    HintError err, "ucHistory_OnMouseUp", False
End Sub

Private Sub ucHistory_OnSend()
    Dim i As Long
    Dim strtext As String
    Dim arySelIndex() As Long
    Dim objImg As DicomImage
    
On Error GoTo errhandle
    If mobjLinkEditor Is Nothing Then
        HintMsg "尚未进入编辑状态，不能写入。", "ucHistory_OnSend", vbOKOnly
        Exit Sub
    End If
    
    strtext = ucHistory.SelReportText
    If Len(strtext) > 0 Then
        Call mobjLinkEditor.InputWord(strtext, "", "", "")
        Exit Sub
    End If
    
    arySelIndex = ucHistory.GetSelects
    If UBound(arySelIndex) <= 0 Then Exit Sub
    
    For i = 1 To UBound(arySelIndex)
        Set objImg = ucHistory.GetImage(arySelIndex(i))
        If Not objImg Is Nothing Then
            Call mobjLinkEditor.AddRepImage(objImg, mlngReleationImgAdvice)
        End If
    Next
Exit Sub
errhandle:
    HintError err, "ucHistory_OnSend", False
End Sub

Private Sub ucImages_OnClick(ByVal lngImgIndex As Long)
    '将图片发送到视频采集中进行处理
    Dim objImg As DicomImage
    Dim objImgTag As clsImageTagInf
    Dim objBgImgInfo As clsBgImgInfo
    
    Set objImg = ucImages.GetImage(lngImgIndex, objBgImgInfo)
    If objImg Is Nothing Then Exit Sub
        
    If objBgImgInfo.LoadState = lsError Or objBgImgInfo.LoadState = lsRedo Or objBgImgInfo.LoadState = lsSent Then
        If Not mobjEmbedVideo Is Nothing Then
            If mobjEmbedVideo.VideoDockState Then Exit Sub
            
            If mblnIsEmbedVideoArea Then
                Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False)
            Else
                Call mobjMainVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), True)
            End If
            
            Exit Sub
        End If
        
        If Not mobjMainVideo Is Nothing Then
            If mobjMainVideo.VideoDockState Then Exit Sub
            Call mobjMainVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), True)
        End If
        
        Exit Sub
    End If
    
    '配置tag
    Set objImgTag = New clsImageTagInf
    objImgTag.videoFile = objBgImgInfo.FilePath & objBgImgInfo.Filename
    objImgTag.EncoderName = ""
    objImgTag.CaptureTime = 0
    objImgTag.RecordTimeLen = 0
    objImgTag.FilePath = objBgImgInfo.FilePath
    
    Select Case objBgImgInfo.Format
        Case ifAvi
            objImgTag.tag = VIDEOTAG
        Case ifWav
            objImgTag.tag = AUDIOTAG
        Case ifDcm
            objImgTag.tag = ImgTag
    End Select
    
    Set objImg.tag = objImgTag
    
    
    If mblnIsEmbedVideoArea And Not mobjEmbedVideo Is Nothing Then
        If mobjEmbedVideo.VideoDockState Then Exit Sub
        '如果视频窗口和当前控件窗口不在同一个窗口，则不显示选择图像
        If GetAncestor(mobjEmbedVideo.VideoHwnd, GA_ROOT) <> GetAncestor(hwnd, GA_ROOT) Then Exit Sub
        Call mobjEmbedVideo.zlPreviewThumbnail(objImg)
        Exit Sub
    End If
    
    
    If Not mobjMainVideo Is Nothing Then
        '需要判断视频采集和pacshelper是否在同一个主窗口
        If GetAncestor(mobjMainVideo.VideoHwnd, GA_ROOT) <> GetAncestor(hwnd, GA_ROOT) Then Exit Sub
        If mobjMainVideo.VideoDockState Then Exit Sub
        
        Call mobjMainVideo.zlPreviewThumbnail(objImg)
    End If
    

    If mblnClickBigImageShow Then
        If lngImgIndex <> mlngBigImageIndex Then
            If objBgImgInfo.Format <> ifAvi And objBgImgInfo.Format <> ifWav Then
                '加载图像并显示
                Call ShowImageProcess(lngImgIndex, ptPreview)
            End If
        End If

        mlngBigImageIndex = lngImgIndex
    End If
End Sub

 

Private Sub ucImages_OnCmdEvent(ByVal strCmd As String)
    If strCmd = "REFRESH" Then Call LoadExamImages(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
End Sub

Private Sub ucImages_OnDBClick(ByVal lngImgIndex As Long)
    '是否双击添加报告图,且需要判断是否为采集图像
    
    '将图片发送到视频采集中进行处理
    Dim objImg As DicomImage
    Dim objImgTag As clsImageTagInf
    Dim objBgImgInfo As clsBgImgInfo
    Dim strImgFailedFile As String
    
    Set objImg = ucImages.GetImage(lngImgIndex, objBgImgInfo)
    If objImg Is Nothing Then Exit Sub
        
    If objBgImgInfo.LoadState = lsError Or objBgImgInfo.LoadState = lsRedo Or objBgImgInfo.LoadState = lsSent Then Exit Sub
    
    '配置tag
    Set objImgTag = New clsImageTagInf
    objImgTag.videoFile = objBgImgInfo.FilePath & objBgImgInfo.Filename
    objImgTag.EncoderName = ""
    objImgTag.CaptureTime = 0
    objImgTag.RecordTimeLen = 0
    objImgTag.FilePath = objBgImgInfo.FilePath
    
    Select Case objBgImgInfo.Format
        Case ifAvi
            objImgTag.tag = VIDEOTAG
        Case ifWav
            objImgTag.tag = AUDIOTAG
        Case ifDcm
            objImgTag.tag = ImgTag
    End Select
    
    Set objImg.tag = objImgTag
    
    If objImgTag.tag <> ImgTag Then
        '播放媒体
        If Not mobjMainVideo Is Nothing Then
            Call mobjMainVideo.playVideo(objImgTag.videoFile)
        End If
    Else
        strImgFailedFile = GetImgCmdFailed(objBgImgInfo)
        If objBgImgInfo.LoadState = lsError Or objBgImgInfo.LoadState = lsSent Or objBgImgInfo.LoadState = lsRedo _
            Or FileExists(strImgFailedFile) Then
            HintMsg "当前图像状态不允许处理。", "WriteData", vbOKOnly
            Exit Sub
        End If
            
            
        '根据需要添加报告图
        If mlngImageDBClickOper = 0 And mblnAllowWrite Then
            Call SendImageToReport
        Else
            Call OpenImageProcess
        End If
    End If
End Sub


Private Sub InitPar()
'读取本地参数
    Dim strPrivatePath As String
    
    strPrivatePath = GetPrivateRegPath("ucPacsHelper")
    
    mblnAllowEmbedVideo = Val(GetDeptPara(mlngDeptID, "显示视频采集", "0")) = 1
    
    mblnMoveBigImageShow = Val(GetSetting("ZLSOFT", strPrivatePath, "移动显示大图", 0)) = 1  '应该调整为用户参数
    mblnClickBigImageShow = Val(GetSetting("ZLSOFT", strPrivatePath, "单击显示大图", 0)) = 1
    mblnDelayCloseImage = Val(GetSetting("ZLSOFT", strPrivatePath, "延迟关闭大图", 0)) = 1
    
    ucHistory.IsOtherDept = Val(GetSetting("ZLSOFT", strPrivatePath, "他科检查", "0")) = 1
    ucHistory.IsThisTime = Val(GetSetting("ZLSOFT", strPrivatePath, "本次相关", "0")) = 1
    ucHistory.IsAutoLine = Val(GetSetting("ZLSOFT", strPrivatePath, "自动换行", "0")) = 1
    
    ucWord.AutoHide = Val(GetSetting("ZLSOFT", strPrivatePath, "自动隐藏", "0")) = 1
    ucWord.ExpandLevel = Val(GetSetting("ZLSOFT", strPrivatePath, "展开层级", "1"))
    ucWord.DblWrite = Val(GetSetting("ZLSOFT", strPrivatePath, "报告词句双击操作", "0")) = 0
    
    ucImages.PageRecordCount = Val(GetSetting("ZLSOFT", strPrivatePath, "缩略图数量(" & mstrParentName & ")", "8"))
    
    mblnIgnoreResult = GetDeptPara(mlngDeptID, "忽略结果阴阳性", 0) = "1" '        '忽略结果阴阳性
    mlngImageDBClickOper = GetDeptPara(mlngDeptID, "缩略图双击操作", 0)
    
    mlngReleationImgDays = Val(GetDeptPara(mlngDeptID, "自动打开历史图像天数", 0))
End Sub


Private Sub SavePar()
'保存本地参数
    Dim strPrivatePath As String
    
    strPrivatePath = GetPrivateRegPath("ucPacsHelper")
    
    Call SaveSetting("ZLSOFT", strPrivatePath, "移动显示大图", IIf(mblnMoveBigImageShow, 1, 0))
    Call SaveSetting("ZLSOFT", strPrivatePath, "单击显示大图", IIf(mblnClickBigImageShow, 1, 0))
    Call SaveSetting("ZLSOFT", strPrivatePath, "延迟关闭大图", IIf(mblnDelayCloseImage, 1, 0))
    
    SaveSetting "ZLSOFT", strPrivatePath, "他科检查", IIf(ucHistory.IsOtherDept, 1, 0)
    SaveSetting "ZLSOFT", strPrivatePath, "本次相关", IIf(ucHistory.IsThisTime, 1, 0)
    SaveSetting "ZLSOFT", strPrivatePath, "自动换行", IIf(ucHistory.IsAutoLine, 1, 0)
    
    
    SaveSetting "ZLSOFT", strPrivatePath, "自动隐藏", IIf(ucWord.AutoHide, 1, 0)
    SaveSetting "ZLSOFT", strPrivatePath, "展开层级", ucWord.ExpandLevel
    SaveSetting "ZLSOFT", strPrivatePath, "报告词句双击操作", IIf(ucWord.DblWrite, 0, 1)
    
    SaveSetting "ZLSOFT", strPrivatePath, "缩略图数量(" & mstrParentName & ")", ucImages.PageRecordCount
End Sub


Public Sub RefreshData(Optional ByVal strHelperName As String = "")
    Dim strSelName As String
    
    strSelName = tabSelect.Selected.tag
    
    If Len(strHelperName) > 0 Then strSelName = strHelperName
    
    Select Case strSelName
        Case CON_TAB_TAG_图像
            If Len(mobjStudyInfo.strStudyUID) > 0 Then Call LoadExamImages(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
            
            If ucImages.ImgCount <= 0 Then
                If mlngReleationImgAdvice <> 0 Then
                    Call LoadExamImages(mlngReleationImgAdvice, False)
                Else
                    mstrReportImageUids = ""
                    ucImages.ClearAll
                End If
            End If
    
        Case CON_TAB_TAG_缓存
            Call ucCache.Refresh
    
        Case CON_TAB_TAG_历史
            '刷新检查历史
            Call ucHistory.Refresh(mobjStudyInfo.lngAdviceId, True)
    
        Case CON_TAB_TAG_词句
            '刷新词句
            Call ucWord.Refresh(mobjStudyInfo.lngAdviceId, mlngFileID, , True)
    End Select
End Sub

Private Sub WriteData()
    Dim objImgInfo As clsBgImgInfo
    Dim strImgFailedFile As String
    Dim blnAllowAdditionInput As Boolean
    
    
    blnAllowAdditionInput = (CheckPopedom(mstrPrivs, "补录报告") And mobjStudyInfo.intStep > 5)
    
    If Not (mblnAllowWrite And (IsStudying Or blnAllowAdditionInput)) Then
        HintMsg "当前报告不能写入。", "WriteData", vbOKOnly
        Exit Sub
    End If
    
    Select Case tabSelect.Selected.tag
        Case CON_TAB_TAG_图像
            If ucImages.SelImgIndex <= 0 Then
                HintMsg "请选择需要写入的图像。", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            '如果是媒体图像，则不能直接加入报告图，只能进行媒体播放
            Call ucImages.GetImage(ucImages.SelImgIndex, objImgInfo)
            
            If objImgInfo.Format = ifAvi And objImgInfo.Format = ifWav Then
                HintMsg "当前格式数据不允许写入。", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            strImgFailedFile = GetImgCmdFailed(objImgInfo)
            If objImgInfo.LoadState = lsError Or objImgInfo.LoadState = lsSent Or objImgInfo.LoadState = lsRedo _
                Or FileExists(strImgFailedFile) Then
                HintMsg "当前图像状态不允许写入。", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            Call SendImageToReport
        
        Case CON_TAB_TAG_历史
            If ucHistory.IsReportEnable(ucHistory.SelAdviceId) = False Then
                HintMsg "当前状态不允许写入,请确认历史报告是否为空。", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            Call ucHistory.WriteReport
            
        Case CON_TAB_TAG_词句
            If ucWord.SelNodeType <> 2 Then
                HintMsg "请选择需要写入的词句节点。", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            Call ucWord.DirectWrite
            
'            Case CON_TAB_TAG_缓存
    End Select
End Sub

Private Sub DelData()
    Dim blnIsCancel As Boolean
    Dim aryIndex() As Long
    Dim objImgInfo As clsBgImgInfo
    
    Select Case tabSelect.Selected.tag
        Case CON_TAB_TAG_图像
            If IsStudying = False Then
                HintMsg "当前状态不允许删除。", "DelData", vbOKOnly
                Exit Sub
            End If
            
            aryIndex = ucImages.GetSelects
            If UBound(aryIndex) <= 0 Then
                HintMsg "请选择需要删除的数据。", "DelData", vbOKOnly
                Exit Sub
            End If
            
            Call ucImages.GetImage(aryIndex(1), objImgInfo)
            
            If objImgInfo.AdviceDes = "REPIMG" Then
                HintMsg "当前图像类别不允许删除。", "DelData", vbOKOnly
                Exit Sub
            End If
        
            '删除检查图像
            Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 0, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, blnIsCancel, hwnd)
            If blnIsCancel Then Exit Sub
            
            If DeleteStudyImage Then
                If ucImages.ImgCount <= 0 Then
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, -1, hwnd)    '-1表示删除最后一张图像
                Else
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, , hwnd)
                End If
            End If
                
'        Case CON_TAB_TAG_历史
            
        Case CON_TAB_TAG_词句
            If ucWord.SelNodeType <> 2 Then
                HintMsg "当前节点不允许删除。", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            Call ucWord.WordDelete
            
        Case CON_TAB_TAG_缓存
            Call DeleteCacheImage
            
    End Select
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnIsCancel As Boolean
    Dim lngImgCount As Long
    
On Error GoTo errhandle
    
    Call mObjNotify.Broadcast(BM_SYS__EVENT_MENU, 0, mobjStudyInfo.lngAdviceId, Control.ID, Control.Category)
    
    Select Case Control.ID
        Case conMenu_Helper_DelOper   '删除操作
            If tabSelect.Selected.tag = CON_TAB_TAG_图像 Then
                '删除检查图像
                Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 0, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, blnIsCancel, hwnd)
                If blnIsCancel Then Exit Sub
                
                If DeleteStudyImage Then
                    If ucImages.ImgCount <= 0 Then
                        Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, -1, hwnd)    '-1表示删除最后一张图像
                    Else
                        Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, , hwnd)
                    End If
                End If
            End If
            
            If tabSelect.Selected.tag = CON_TAB_TAG_缓存 Then
                '删除缓存图像
                Call DeleteCacheImage
            End If
            
        Case conMenu_Helper_Refresh    '刷新操作
            Call RefreshData
            
        Case conMenu_Helper_SelAll '全选
            If tabSelect.Selected.tag = CON_TAB_TAG_图像 Then
                Call ucImages.SelectedAll
            End If
            
            If tabSelect.Selected.tag = CON_TAB_TAG_缓存 Then
                Call ucCache.SelectedAll
            End If
            
        Case conMenu_Helper_AddToReport '加入报告图
            If Len(Control.Parameter) <= 0 Then
                Call SendImageToReport
            Else
                '播放媒体
                If Not mobjMainVideo Is Nothing Then
                    Call mobjMainVideo.playVideo(Control.Parameter)
                End If
            End If
            
        Case conMenu_Helper_ImageProcess '图像-图像处理
            If tabSelect.Selected.tag = CON_TAB_TAG_图像 Then
                If Len(Control.Parameter) <= 0 Then
                    Call OpenImageProcess
                Else
                    Call OpenImageProcess(True)
                End If
            ElseIf tabSelect.Selected.tag = CON_TAB_TAG_缓存 Then
                Call OpenImageProcess(True, True)
            End If
            
        Case conMenu_Helper_Log '传输日志查看
            If FileExists(Control.Parameter) Then
                Call OpenFilePos(Control.Parameter, True, "notepad")
            Else
                HintMsg "日志尚未生成。", "cbrMain_Execute", infHint
            End If
            
        Case conMenu_Helper_Import '图像-导入图像
            lngImgCount = ucImages.ImgCount
            
            Call ImportImageFile
            
            If ucImages.ImgCount > 0 Then
                If lngImgCount <= 0 Then
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_FIRST, , mobjStudyInfo.lngAdviceId, -1, , hwnd)    '-1表示首次采集
                ElseIf ucImages.ImgCount > lngImgCount Then
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_FIRST, , mobjStudyInfo.lngAdviceId, , , hwnd)
                End If
            End If
            
        Case conMenu_Helper_Export '图像-导出图像
            Call ExportImageFile
            
        Case conMenu_Helper_OpenImgPos '打开图像位置
            If tabSelect.Selected.tag = CON_TAB_TAG_图像 Then
                Call OpenImagePos
            End If
            
            If tabSelect.Selected.tag = CON_TAB_TAG_缓存 Then
                Call OpenCachePos
            End If
            
        Case conMenu_Helper_BigImageShow_Move '图像-移动显示大图
            Control.Checked = Not Control.Checked
            mblnMoveBigImageShow = Control.Checked
            
            
        Case conMenu_Helper_BigImageShow_Click  '单击显示大图
            Control.Checked = Not Control.Checked
            mblnClickBigImageShow = Control.Checked
            
        Case conMenu_Helper_BigImageShow_Delay  '延迟关闭大图
            Control.Checked = Not Control.Checked
            mblnDelayCloseImage = Control.Checked
        
'        Case conMenu_Helper_ReDo    '图像-重新尝试
'            Call ucImages.Redo
             
        Case conMenu_Helper_ReDown  '图像-重新下载
            Call ucImages.ReDown
            
        Case conMenu_Helper_ReUp    '图像-重新上传
            If HintMsg("重新上传会使FTP存储数据被替换，是否继续？", "cbrMain_Execute", vbYesNo) = vbNo Then Exit Sub
            Call ucImages.ReUp
            
        Case conMenu_Helper_SendCache '图像-发送到缓存
            Call SendCache
            
            If ucImages.ImgCount <= 0 Then
                Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, -1, hwnd)    '-1表示删除最后一张图像
            Else
                Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, , hwnd)
            End If
        
        Case conMenu_Helper_SendStudy '缓存-发送到检查
            lngImgCount = ucImages.ImgCount
            Call SendStudy
            
            If ucImages.ImgCount > 0 Then
                If lngImgCount <= 0 Then
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_FIRST, , mobjStudyInfo.lngAdviceId, -1, , hwnd)    '-1表示首次采集
                ElseIf ucImages.ImgCount > lngImgCount Then
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_FIRST, , mobjStudyInfo.lngAdviceId, , , hwnd)
                End If
            End If
            
        Case conMenu_Helper_ImgViewer '历史-观片
            Call mObjNotify.SendRequest(WM_IMG_OPENVIEW, , ucHistory.SelAdviceId)
            
        Case conMenu_Helper_ImgContrast '历史-对比
            Call mObjNotify.SendRequest(WM_IMG_CONTRASTVIEW, , ucHistory.SelAdviceId)
            
        Case conMenu_Helper_ReportOpen '历史-报告查看
            Call mObjNotify.SendRequest(WM_REPORT_VIEW, , ucHistory.SelAdviceId, ucHistory.SelMoved)
        
        Case conMenu_Helper_ViewReportImage '查看报告图
            Call ucHistory.ViewReportImage
            
        Case conMenu_Helper_ViewReportContext '查看报告文本
            Call ucHistory.ViewReportContext
            
        Case conMenu_Helper_WriteReport '写入报告
            Call ucHistory.WriteReport
            
        Case conMenu_Helper_LinkViewer
            ucHistory.LinkViewed = Not ucHistory.LinkViewed
            
        Case conMenu_Helper_CloseViewer
            Call ucHistory.CloseLinkViewer
        
        Case conMenu_Helper_HalfYear, conMenu_Helper_TwoYear, conMenu_Helper_TwoMonth, _
            conMenu_Helper_ThreeYear, conMenu_Helper_ThreeMonth, conMenu_Helper_OneYear, _
            conMenu_Helper_OneMonth, conMenu_Helper_DateCus, conMenu_Helper_DateUn, _
            conMenu_Helper_DateCus '历史-自定日期
             
            
            If Control.ID = conMenu_Helper_DateCus Then
                Call ucHistory.SetDateRange("自定义")
                Call ucHistory.ShowDateConfig
            Else
                Call ucHistory.SetDateRange(Control.Caption)
            End If
            
            Call ucHistory.Refresh(mobjStudyInfo.lngAdviceId, True)
            
        Case conMenu_Helper_ThisTime    '历史-是否本次相关
            Control.Checked = Not Control.Checked
            ucHistory.IsThisTime = Control.Checked
            
            ucHistory.Refresh mobjStudyInfo.lngAdviceId, True
        
        Case conMenu_Helper_OtherDept   '历史-是否他科检查
            Control.Checked = Not Control.Checked
            ucHistory.IsOtherDept = Control.Checked
            
            ucHistory.Refresh mobjStudyInfo.lngAdviceId, True
        
        Case conMenu_Helper_AutoLine  '历史-是否自动换行
            Control.Checked = Not Control.Checked
            ucHistory.IsAutoLine = Control.Checked
            
            ucHistory.Refresh mobjStudyInfo.lngAdviceId, True
            
        Case conMenu_Helper_DirectWrite '词句-直接写入
            Call ucWord.DirectWrite
            
        Case conMenu_Helper_EditWrite   '词句-编辑写入
            Call ucWord.EditWrite
            
        Case conMenu_Helper_FullSave    '词句-全套存入
            Call ucWord.FullSave
            
        Case conMenu_Helper_NewWord     '词句-新增词句
            Call ucWord.WordNew
            
        Case conMenu_Helper_ModWord     '词句-修改词句
            Call ucWord.WordModify
            
        Case conMenu_Helper_DelWord     '词句-删除词句
            Call ucWord.WordDelete
            
        Case conMenu_Helper_AutoHide    '词句-自动隐藏
            Control.Checked = Not Control.Checked
            ucWord.AutoHide = Control.Checked
            
        Case conMenu_Helper_DblWrite    '双击写入
            Control.Checked = Not Control.Checked
            ucWord.DblWrite = Control.Checked
            
        Case conMenu_Helper_AllLevel    '词句-展开所有
            Control.Checked = True
            ucWord.ExpandLevel = 0
            
        Case conMenu_Helper_OneLevel    '词句-展开一级
            Control.Checked = True
            ucWord.ExpandLevel = 1
            
        Case conMenu_Helper_TwoLevel    '词句-展开二级
            Control.Checked = True
            ucWord.ExpandLevel = 2
            
        Case conMenu_Helper_ThreeLevel  '词句-展开三级
            Control.Checked = True
            ucWord.ExpandLevel = 3
    End Select
    
    Call mObjNotify.Broadcast(BM_SYS__EVENT_MENU, 1, mobjStudyInfo.lngAdviceId, Control.ID, Control.Category)
Exit Sub
errhandle:
    HintError err, "cbrMain_Execute", False
End Sub

Private Sub SendImageToReport()
'发送图像到报告
    Dim i As Long
    Dim arySelIndex() As Long
    Dim objImg As DicomImage
    Dim strSQL As String
    
    
     arySelIndex = ucImages.GetSelects()
     
     If UBound(arySelIndex) <= 0 Then
        HintMsg "请选择需要发送到报告的图像。", "SendImageToReport", vbOKOnly
        Exit Sub
     End If
     
     For i = 1 To UBound(arySelIndex)
        Set objImg = ucImages.GetImage(arySelIndex(i))
        If Not objImg Is Nothing Then
        
            '更新数据库
            strSQL = "Zl_影像检查_设置报告图('" & objImg.InstanceUID & "',1)"
            Call zlDatabase.ExecuteProcedure(strSQL, "预置报告图")
                
            If Not mobjLinkEditor Is Nothing Then
                mobjLinkEditor.AddRepImage objImg, mlngReleationImgAdvice
            End If
            
            '设置报告图标记
            Call ucImages.ImgDrawHint(objImg.InstanceUID, "报")
        End If
     Next
End Sub

Private Sub SendStudy()
'发送图像到检查
    Dim i As Long
    Dim arySelIndex() As Long
    Dim strLineDeviceNo As String
    Dim strBackDeviceNo As String
    Dim objDcmImg As DicomImage
    
    arySelIndex = ucCache.GetSelects
    
    If UBound(arySelIndex) <= 0 Then
        HintMsg "请选择需要发送的检查图像。", "SendStudy", vbOKOnly
    End If
    
    strLineDeviceNo = GetDeptPara(mlngDeptID, "存储设备号")
    strBackDeviceNo = GetDeptPara(mlngDeptID, "备份设备号")
     
    For i = UBound(arySelIndex) To 1 Step -1
        Set objDcmImg = ucCache.GetImage(arySelIndex(i))
        If Not objDcmImg Is Nothing Then
            If Len(objDcmImg.InstanceUID) > 0 Then
                If Not SaveDicomImageToStudy(objDcmImg, strLineDeviceNo, strBackDeviceNo) Is Nothing Then
                    If mobjStudyInfo.strStudyUID = "" Or mobjStudyInfo.strStudyUID <> objDcmImg.StudyUID Then mobjStudyInfo.strStudyUID = objDcmImg.StudyUID
                End If
            End If
        End If
    Next
    
    Call ucCache.DeleteCacheImg(-1)
End Sub

Private Sub SendCache()
'发送到缓存
    Dim i As Long
    Dim arySelIndex() As Long
    Dim strCacheTag As String
    Dim strCachePath As String
    Dim objImg As DicomImage
    Dim objImgInfo As clsBgImgInfo
    
    
    arySelIndex = ucImages.GetSelects()
    
    If UBound(arySelIndex) <= 0 Then
        HintMsg "请选择需要发送到本地缓存的检查图像。", "SendCache", vbOKOnly
        Exit Sub
    End If
    
    If HintMsg("图像发送到缓存后将删除检查对应的数据，是否继续？", "SendCache", vbYesNo) = vbNo Then Exit Sub
    
    strCacheTag = mobjStudyInfo.strPatientName & "(临时)"  '  Format(Now, "hhmmss")
    strCachePath = GetCachePath(Format(Now, "YYYYMMDD"), strCacheTag)
    
    If DirExists(strCachePath) = False Then Call MkLocalDir(strCachePath)
    
    '复制文件到缓存目录
    For i = 1 To UBound(arySelIndex)
        Set objImg = ucImages.GetImage(arySelIndex(i), objImgInfo)
        If Not objImg Is Nothing Then
            FileCopy objImgInfo.FilePath & objImgInfo.Filename, strCachePath & objImgInfo.Filename
        End If
    Next
    
    '删除对应的图像数据
    Call DeleteStudyImage(True)
    
    HintMsg "图像已发送到标记为 [" & strCacheTag & "] 的缓存目录中。", "SendCache", vbOKOnly
End Sub


Private Sub OpenImageProcess(Optional ByVal blnForceRead As Boolean = False, Optional ByVal blnIsCache As Boolean = False)
    Dim arySelIndex() As Long
    
    If blnIsCache Then
        If ucCache.ImgCount <= 0 Then
            HintMsg "请选择需要查看的图像。", "ImageProcess", vbOKOnly
            Exit Sub
        End If
        
        arySelIndex = ucCache.GetSelects

    Else
        If ucImages.ImgCount <= 0 Then
            HintMsg "请选择需要处理的图像。", "ImageProcess", vbOKOnly
            Exit Sub
        End If
        
        arySelIndex = ucImages.GetSelects
    End If
    
    If UBound(arySelIndex) <= 0 Then
        HintMsg "尚未选择需用于当前操作的图像。", "ImageProcess", vbOKOnly
        Exit Sub
    End If
     
    Call ShowImageProcess(arySelIndex(1), ptProcess, blnForceRead, blnIsCache)
End Sub

Private Function GetDlgSelectFileInfo(strFileName As String) As DlgFileInfo
'------------------------------------------------
'功能：将文件名转化为全路径数组
'参数：strFileName--文件名，通过打开文件控件来获得。
'返回：全路径数组
'------------------------------------------------
    Dim sPath, tmpStr As String
    Dim sFIle() As String
    Dim iCount, i As Integer
    On Error GoTo errhandle
    sPath = CurDir()  '获得当前的路径，因为在CommonDialog中改变路径时会改变当前的Path
    tmpStr = Right$(strFileName, Len(strFileName) - Len(sPath)) '将文件名分离出来
    
    If Left$(tmpStr, 1) = Chr$(0) Then
        '选择了多个文件(表现为第一个字符为空格)
        For i = 1 To Len(tmpStr)
            If Mid$(tmpStr, i, 1) = Chr$(0) Then
                iCount = iCount + 1
                ReDim Preserve sFIle(iCount)
            Else
                sFIle(iCount) = sFIle(iCount) & Mid$(tmpStr, i, 1)
            End If
        Next i
    Else
        '只选择了一个文件(注意：根目录下的文件名除去路径后没有"\"）
        iCount = 1
        ReDim Preserve sFIle(iCount)
        If Left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
        sFIle(iCount) = tmpStr
    End If
    
    GetDlgSelectFileInfo.iCount = iCount
    
    ReDim GetDlgSelectFileInfo.sFIle(iCount)
    
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    GetDlgSelectFileInfo.sPath = sPath
    
    For i = 1 To iCount
        GetDlgSelectFileInfo.sFIle(i) = sFIle(i)
    Next i
    Exit Function
errhandle:
    HintError err, "GetDlgSelectFileInfo", False
End Function


Private Sub ImportImageFile()
'------------------------------------------------
'功能：打开外部文件，放入缩略图中
'参数：无
'返回：无
'------------------------------------------------
'TASK:》》》》》》》》》》》》》》》》》》》暂时不支持AVI导入，后续可以增加》》》》》》》》》》》》
On Error GoTo errH
    Dim DlgInfo As DlgFileInfo
    Dim i As Integer
    Dim objDcmImg As New DicomImage
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    Dim dcmInfo As TDicomBaseInfo
    Dim strError As String
    
    Dim strLineDeviceNo As String
    Dim strBackDeviceNo As String
    
    Dim lineFtpInfo As TFtpDeviceInf
    Dim backFtpInfo As TFtpDeviceInf
    
    Dim objImgInfo As clsBgImgInfo
    
    '选择文件
    With dlgOpen
        .CancelError = False
        .MaxFileSize = 32767 '被打开的文件名尺寸设置为最大，即32K
        .flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .DialogTitle = "选择文件"
        .Filter = "DICOM文件（*.dcm）(*.img)|*.dcm;*.img|图像文件 (*.BMP)(*.JPG)|*.BMP;*.JPG|所有文件（*.*）|*.*"
        .ShowOpen
        If .Filename <> "" Then
            DlgInfo = GetDlgSelectFileInfo(.Filename)
        End If
        '在打开了*.pif文件后须将Filename属性置空，否则当选取多个*.pif文件后，当前路径会改变
        .Filename = ""
    End With
    
    strLineDeviceNo = GetDeptPara(mlngDeptID, "存储设备号")
    strBackDeviceNo = GetDeptPara(mlngDeptID, "备份设备号")
    
    For i = 1 To DlgInfo.iCount
        
        Set objDcmImg = ReadDicomFile(DlgInfo.sPath & DlgInfo.sFIle(i), strError)
        
        If objDcmImg Is Nothing Then
            HintMsg "文件 " & DlgInfo.sPath & DlgInfo.sFIle(i) & " 读取异常。" & IIf(strError = "", "", vbCrLf & strError), "ImportImageFile", infNormalErr
            Exit Sub
        End If
        
        If Len(objDcmImg.InstanceUID) <= 0 Then
            HintMsg "文件 " & DlgInfo.sPath & DlgInfo.sFIle(i) & " 读取异常，未产生实例UID。" & IIf(strError = "", "", vbCrLf & strError), "ImportImageFile", infNormalErr
            Exit Sub
        End If
        
        If Not SaveDicomImageToStudy(objDcmImg, strLineDeviceNo, strBackDeviceNo) Is Nothing Then
            If mobjStudyInfo.strStudyUID = "" Or mobjStudyInfo.strStudyUID <> objDcmImg.StudyUID Then mobjStudyInfo.strStudyUID = objDcmImg.StudyUID
        End If
    Next
    
    Exit Sub
errH:
    HintError err, "ImportImageFile", False
End Sub


Private Function SaveDicomImageToStudy(objDcmImg As DicomImage, _
    ByVal strLineDeviceNo As String, ByVal strBackDeviceNo As String) As clsBgImgInfo
 
    Dim dcmInfo As TDicomBaseInfo
    Dim strError As String
     
    Dim lineFtpInfo As TFtpDeviceInf
    Dim backFtpInfo As TFtpDeviceInf
    
    Dim objImgInfo As clsBgImgInfo
    Dim objResult As clsBgImgInfo
    
    Set SaveDicomImageToStudy = Nothing
    
    dcmInfo = GetDicomBaseInfo(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
    
    Call WriteDicomPara(objDcmImg, dcmInfo)
    
    lineFtpInfo = GetLineFtpInfo(strLineDeviceNo, False, dcmInfo, strError)
    If Len(strError) > 0 Then
        err.Raise 0, "", strError
        Exit Function
    End If
    
    backFtpInfo = GetBackFtpInfo(strBackDeviceNo, dcmInfo, strError)
    If Len(strError) > 0 Then
        err.Raise 0, "", strError
        Exit Function
    End If
    
    If FileExists(GetStudyImgPath(dcmInfo) & dcmInfo.strInstanceUID) = False Then
        objDcmImg.WriteFile GetStudyImgPath(dcmInfo) & dcmInfo.strInstanceUID, True, "1.2.840.10008.1.2.1"
    End If
    
    Call SaveImageInfo(dcmInfo, lineFtpInfo)
    
    
    Set objImgInfo = GetBgImgInfo(dcmInfo, lineFtpInfo, backFtpInfo, True)
    
    objImgInfo.JpgConvert = True
    
    Set objResult = objImgInfo.CopyNew
    
    Call ucImages.AddImg(objImgInfo)
    
    Set SaveDicomImageToStudy = objResult
End Function

Private Sub ExportSingleImg(objImg As DicomImage, objImgInfo As clsBgImgInfo, _
    Optional ByVal blnUseFix As Boolean = False, Optional ByVal strExportFile As String = "", Optional ByVal strSuffix As String = "")
'导出单幅图像
    Dim strFileName As String
    Dim blnIsCopy As Boolean
    Dim strFileType As String
     
    
    blnIsCopy = False
    
    If blnUseFix = False Then
        '不使用前缀字符
        strFileName = Replace(UCase(strExportFile), "." & strSuffix, "")
    Else
        strFileName = Replace(UCase(strExportFile), "." & strSuffix, "") & "_" & objImg.InstanceUID
    End If
        
    If Len(strFileName) <= 0 Then Exit Sub
    
    If Trim(strSuffix) = "" Then
        blnIsCopy = True
    Else
        If objImgInfo.Format = ifAvi Or objImgInfo.Format = ifWav Then
            blnIsCopy = True
        Else
            strFileName = strFileName & "." & strSuffix
        End If
    End If
    
    If blnIsCopy Then
        Call FileCopy(objImgInfo.FilePath & objImgInfo.Filename, strFileName)
        Exit Sub
    End If
    
    strFileType = UCase(Right(Trim(strFileName), 3))
    
    Select Case strFileType
        Case "AVI"
            objImg.WriteAVI strFileName, 1, objImg.FrameCount, 1, "", 100, False
        Case "DCM"
            objImg.WriteFile strFileName, True
        Case "BMP"
            objImg.FileExport strFileName, "BMP"
        Case "JPG"
            objImg.FileExport strFileName, "JPG"
    End Select
End Sub

Private Sub ExportImageFile()
'------------------------------------------------
'功能：另存dcmView中的图像,支持的格式为AVI,DCM,BMP,JPE
'参数：无
'返回：无
'------------------------------------------------
    Dim i As Long
    
    Dim arySelIndex() As Long
    Dim objDcmImg As DicomImage
    Dim objImgInfo As clsBgImgInfo
    
    Dim strExt As String
    
    arySelIndex = ucImages.GetSelects
    
    If UBound(arySelIndex) <= 0 Then
        HintMsg "请选择需要导出的图像。", "ExportImageFile", vbOKOnly
        Exit Sub
    End If
    
    dlgOpen.Filter = "原始格式| |(*.dcm)|*.dcm|(*.bmp)|*.bmp|(*.jpg)|*.jpg|(*.avi)|*.avi|(*.mpeg)|*.mpeg"
    dlgOpen.FilterIndex = 1
    dlgOpen.Filename = mobjStudyInfo.strPatientName
    
    dlgOpen.ShowSave
    
    If Len(dlgOpen.Filename) <= 0 Then Exit Sub
    
    strExt = ""
    
    Select Case dlgOpen.FilterIndex
        Case 1
            strExt = ""
        Case 2
            strExt = "DCM"
        Case 3
            strExt = "BMP"
        Case 4
            strExt = "JPG"
        Case 5
            strExt = "AVI"
        Case 6
            strExt = "MPEG"
    End Select
        
    If UBound(arySelIndex) = 1 Then
        '导出单幅图像
        Set objDcmImg = ucImages.GetImage(arySelIndex(1), objImgInfo)

        Call ExportSingleImg(objDcmImg, objImgInfo, False, dlgOpen.Filename, strExt)
    Else
        '导出多幅图像
        For i = 1 To UBound(arySelIndex)
            Set objDcmImg = ucImages.GetImage(arySelIndex(i), objImgInfo)
            
            Call ExportSingleImg(objDcmImg, objImgInfo, True, dlgOpen.Filename, strExt)
        Next
    End If
    
    If HintMsg("文件导出完成，是否打开导出目录?", "ExportImageFile", vbYesNo) = vbYes Then
        Call OpenFilePos(dlgOpen.Filename)
    End If
End Sub

Private Sub OpenFilePos(ByVal strFile As String, Optional ByVal blnIsOpenFile As Boolean = False, Optional ByVal strOpenWay As String = "")
'打开文件位置
    If blnIsOpenFile = False Then
        ShellExecute 0, "open", Mid(strFile, 1, InStrRev(strFile, "\")), "", "", 1
    Else
        If Len(strOpenWay) <= 0 Then
            ShellExecute 0, "open", strFile, "", "", 1
        Else
            ShellExecute 0, "open", strOpenWay, strFile, "", 1
        End If
    End If
    
End Sub

Private Sub OpenImagePos()
'打开图像位置
    Dim objImgInfo As clsBgImgInfo
    
    If ucImages.ImgCount <= 0 Then Exit Sub
    
    Call ucImages.GetImage(1, objImgInfo)
    
    If objImgInfo Is Nothing Then Exit Sub
    
    Call OpenFilePos(objImgInfo.FilePath & objImgInfo.Filename)
End Sub


Private Sub OpenCachePos()
'打开缓存位置
    If ucCache.ImgCount <= 0 Then Exit Sub
    
    ucCache.OpenCachePath
End Sub

Private Sub DeleteCacheImage()
'删除缓存的图像
    If ucCache.ImgCount <= 0 Then Exit Sub
    
    '判断是否有图像被选中
    If ucCache.IsSelected = False Then
        HintMsg "请选择需要删除的缓存图像。", "DeleteCacheImage", vbOKOnly
        Exit Sub
    End If
    
    If HintMsg("缓存图像删除后将不能恢复，是否继续？", "DeleteCacheImage", vbYesNo) = vbNo Then Exit Sub
    
    Call ucCache.DeleteCacheImg(-1)
End Sub

Private Sub SetRemoveTag(ByVal strFile As String)
On Error Resume Next
    Call SetFileHide(strFile)
    
    Name strFile As strFile & ".DEL"
End Sub

Private Function DeleteStudyImage(Optional ByVal blnIsSendCache As Boolean = False) As Boolean
'------------------------------------------------
'功能：删除缩略图中被选中的图像，先从数据库中删除，然后从FTP中删除。
'参数：无
'返回：无，直接删除缩略图中最后一个图像
'------------------------------------------------
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    Dim i As Long
    Dim lngIndex  As Long
    Dim arySelIndex() As Long
    
    Dim objImgInfo As clsBgImgInfo
    Dim blnIsContainReport As Boolean
    
    Dim strucFtpTag As TFtpConTag

    DeleteStudyImage = False
    If ucImages.ImgCount <= 0 Then Exit Function
     
    blnIsContainReport = False
    arySelIndex = ucImages.GetSelects()
    
    If UBound(arySelIndex) <= 0 Then Exit Function
    
    If blnIsSendCache = False Then
        If HintMsg("图像删除后将不能恢复，是否继续？", "DeleteStudyImage", vbYesNo) = vbNo Then Exit Function
    End If
    
    
    '得到需要删除的图像uid中间用';'隔开
    For i = UBound(arySelIndex) To 1 Step -1
        lngIndex = arySelIndex(i)
        Call ucImages.GetImage(lngIndex, objImgInfo)

        If InStr(mstrReportImageUids, ";" & objImgInfo.Key & ";") <= 0 And InStr(objImgInfo.DrawHint, "报") <= 0 Then '如果是报告图，或者有报告图标记，都不允许进行删除
            If strucFtpTag.Ip <> objImgInfo.FtpIp Then
                strucFtpTag = FtpTagInstance(objImgInfo.FtpIp, objImgInfo.FtpUser, objImgInfo.FtpPwd, objImgInfo.FtpVirtualPath)
            End If
            
            strSQL = "ZL_影像图象_DELETE(" & objImgInfo.AdviceId & ",0,'" & objImgInfo.Key & "',Null)"
            zlDatabase.ExecuteProcedure strSQL, "删除检查图像"
            
            '删除图像文件
            If FtpDelete(strucFtpTag, objImgInfo.Key, False, False) = frAbort Then Exit Function
            '本地文件设置为隐藏
            Call SetRemoveTag(objImgInfo.FilePath & objImgInfo.Filename)
            
            '如果是dcm文件，才可能存在对应的jpg文件
            If objImgInfo.Format = ifDcm Then
                If FtpDelete(strucFtpTag, objImgInfo.Key & ".jpg", True, False) = frAbort Then Exit Function
                Call SetRemoveTag(objImgInfo.FilePath & objImgInfo.Key & ".jpg")
            End If
            
            Call ucImages.DelImgView(lngIndex)
        Else
            blnIsContainReport = True
        End If
    Next
    
    Call ucImages.Selected(lngIndex)
     
    If blnIsContainReport Then
        HintMsg IIf(blnIsSendCache, "发送到缓存的", "删除") & "的图像中包含报告图，报告图不允许删除。", "DeleteStudyImage", vbOKOnly
        Exit Function
    End If
    
    DeleteStudyImage = True
    
    '判断是否还存在检查图像
    If ucImages.ImgCount <= 0 Then
 
'        strSQL = "Select * from 病人医嘱发送 where 医嘱ID=[1] and 发送号=[2]"
'        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查医嘱状态", mlngAdviceID, mlngSendNo)
'
'        '如果检查状态为已检查，则再删除所有图像后，需要对图像进行回退
'        If Val(nvl(rsData!执行过程)) = 3 Then
'            '设置影像检查状态，如果删除最后一个图，且原检查过程为3，则修改为2
'            strSQL = "Zl_影像检查_State(" & mlngAdviceID & "," & mlngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngDeptID & ")"
'            zlDatabase.ExecuteProcedure strSQL, "删除最后一个图像"
'
'            '发送状态改变消息
''            Call mObjNotify.SendRequest(WM_LIST_SYNCROW, , mlngAdviceID, mlngSendNo)
'        End If
        
        '恢复实时视频显示
        If Not mobjEmbedVideo Is Nothing Then
            If mobjEmbedVideo.VideoDockState Then Exit Function
            
            If mblnIsEmbedVideoArea Then
                Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False)
            Else
                Call mobjMainVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), True)
            End If
        Else
            If Not mobjMainVideo Is Nothing Then
                If mobjMainVideo.VideoDockState Then Exit Function
                Call mobjMainVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), True)
            End If
        End If
    End If
        
    Exit Function
errH:
    If HintError(err, "DeleteStudyImage", False) = 1 Then Resume
End Function


Private Sub ShowPopupMenu(ByVal strTabTag As String)
'------------------------------------------------
'功能：创建鼠标右键弹出菜单
'intType: 1--缩略图，2--缓存图
'------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    Dim blnVisible As Boolean
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    Dim strCaption As String
    Dim blnAllowAdditionInput As Boolean
    Dim objImgInfo As clsBgImgInfo
    Dim lngImgFmt As TImageFmt
    Dim blnResultOk As Boolean
    Dim strImgFailedFile As String
    Dim blnIsReportImg As Boolean
    
    '鼠标右键弹出菜单
    cbrMain.DeleteAll
    
    Set cbrToolBar = cbrMain.Add("鼠标右键", xtpBarPopup)
    
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    
    blnAllowAdditionInput = (CheckPopedom(mstrPrivs, "补录报告") And mobjStudyInfo.intStep > 5)
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Helper_Refresh, "刷新(&R)")
                
 
        Select Case strTabTag
            Case CON_TAB_TAG_图像 '**********************************************************************************
                blnIsReportImg = False
                If ucImages.ImgCount <= 0 Then
                    blnVisible = False
                    blnResultOk = False
                Else
                    If ucImages.SelImgIndex > 0 Then
                        blnVisible = True
                        '如果是媒体图像，则不能直接加入报告图，只能进行媒体播放
                        Call ucImages.GetImage(ucImages.SelImgIndex, objImgInfo)
                        
                        lngImgFmt = objImgInfo.Format
                        blnIsReportImg = IIf(objImgInfo.AdviceDes = "REPIMG", True, False)
                        blnResultOk = IIf(objImgInfo.LoadState <> lsError And objImgInfo.LoadState <> lsSent And objImgInfo.LoadState <> lsRedo, True, False)
                        strImgFailedFile = GetImgCmdFailed(objImgInfo)
                    Else
                        blnVisible = False
                    End If
                End If
            
                If lngImgFmt <> ifAvi And lngImgFmt <> ifWav Then
                    Set cbrControl = .Add(xtpControlButton, conMenu_Helper_AddToReport, "加入报告图(&i)")
                    cbrControl.Parameter = ""
                    cbrControl.Visible = blnVisible And blnResultOk
                    cbrControl.Enabled = (IsStudying Or blnAllowAdditionInput) 'mblnAllowWrite And
                    
'                    If Not mobjLinkEditor Is Nothing Then
'                        cbrControl.Enabled = cbrControl.Enabled And mobjLinkEditor.IsEditable
'                    End If
                Else
                    Set cbrControl = .Add(xtpControlButton, conMenu_Helper_AddToReport, "播放(&P)")
                    cbrControl.Parameter = objImgInfo.FilePath & objImgInfo.Filename
                    cbrControl.Visible = blnVisible And blnResultOk
                End If
                
                cbrControl.BeginGroup = True
                
                    
'                If blnResultOk Then
                    Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ImageProcess, "图像处理(&C)")
                        cbrControl.Visible = blnVisible
                        cbrControl.Parameter = strImgFailedFile
'                Else
                    If FileExists(strImgFailedFile) Then
                        Set cbrControl = .Add(xtpControlButton, conMenu_Helper_Log, "传输日志(&L)")
                            cbrControl.Parameter = strImgFailedFile
                            cbrControl.Visible = blnVisible
                    End If
'                End If
                
                
                    
    '            Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SplitPage, "分页设置")
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SelAll, "全选图像(&F)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = IIf(ucImages.ImgCount > 0, True, False)
                    
                If mlngModuleNo = G_LNG_VIDEOSTATION_MODULE Then
                    Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SendCache, "发送到缓存(&E)")
                        cbrControl.Visible = blnVisible And blnResultOk
                        
                    Set cbrControl = .Add(xtpControlButton, conMenu_Helper_Import, "导入图像(&M)")
                        cbrControl.Visible = mlngModuleNo <> G_LNG_PACSSTATION_MODULE
                        cbrControl.Enabled = IsStudying
                End If
                 
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_Export, "导出图像(&T)")
                    cbrControl.Visible = blnVisible
                      
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_DelOper, "删除图像(&D)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible And blnIsReportImg = False
                    cbrControl.Enabled = IsStudying
                    

                
'                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ReDo, "重试...")
'                    cbrControl.BeginGroup = True
'                    cbrControl.Visible = blnVisible
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ReDown, "重新下载(&N)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ReUp, "重新上传(&U)")
                    cbrControl.Visible = blnVisible
                     
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_OpenImgPos, "打开图像位置(&O)...")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_BigImageShow_Move, "移动显示大图(&B)")
                    cbrControl.BeginGroup = True
                    cbrControl.Checked = mblnMoveBigImageShow
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_BigImageShow_Click, "单击显示大图(&K)")
                    cbrControl.Checked = mblnClickBigImageShow
                    
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_BigImageShow_Delay, "延迟关闭大图(&A)")
                    cbrControl.Checked = mblnDelayCloseImage
                    

            Case CON_TAB_TAG_缓存 '**********************************************************************************
                blnVisible = IIf(ucCache.ImgCount > 0, True, False)
    '            Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SplitPage, "分页设置")
    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ImageProcess, "图像预览(&V)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SendStudy, "发送到检查(&S)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = IsStudying
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_DelOper, "删除(&D)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SelAll, "全选图像(&F)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = IIf(ucCache.ImgCount > 0, True, False)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_OpenImgPos, "打开图像位置(&O)...")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
            
            Case CON_TAB_TAG_历史 '**********************************************************************************
                blnVisible = ucHistory.SelAdviceId > 0
                
                strCaption = IIf(mlngModuleNo <> G_LNG_PACSSTATION_MODULE, "影像查阅(&F)", "影像观片(&F)")
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ImgViewer, strCaption)
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsImageEnable(ucHistory.SelAdviceId) Or mlngReleationImgDays > 0
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ImgContrast, "影像对比(&C)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsImageEnable(ucHistory.SelAdviceId)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ReportOpen, "报告预览(&V)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsReportEnable(ucHistory.SelAdviceId)
                    
'                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_Analysis, "综合分析")
'                    cbrControl.Visible = blnVisible
'                    cbrControl.Enabled = ucHistory.IsReportEnable(ucHistory.SelAdviceId)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ViewReportImage, "查看报告图(&i)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsReportEnable(ucHistory.SelAdviceId)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ViewReportContext, "查看报告文本(&T)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsReportEnable(ucHistory.SelAdviceId)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_WriteReport, "写入报告(&W)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsReportEnable(ucHistory.SelAdviceId) And mblnAllowWrite And (IsStudying Or blnAllowAdditionInput)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_LinkViewer, "联动查看(&K)")
                    cbrControl.Visible = blnVisible And ucHistory.AllowLinkViewer
                    cbrControl.Checked = ucHistory.LinkViewed
                    cbrControl.BeginGroup = True
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_CloseViewer, "退出查看(&Q)")
                    cbrControl.Visible = blnVisible And ucHistory.AllowLinkViewer
                    
            
                '时间.........................................................
                Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Helper_DateRange, "日期(&D)")
                    cbrControl.ToolTipText = "日期"
                    cbrControl.BeginGroup = True
        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_OneMonth, "一个月(&1)")
                        objControl.Checked = IIf(ucHistory.DataRange = "一个月", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_TwoMonth, "二个月(&2)")
                        objControl.Checked = IIf(ucHistory.DataRange = "二个月", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_ThreeMonth, "三个月(&3)")
                        objControl.Checked = IIf(ucHistory.DataRange = "三个月", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_HalfYear, "半年(&4)")
                        objControl.Checked = IIf(ucHistory.DataRange = "半年", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_OneYear, "一年(&5)")
                        objControl.Checked = IIf(ucHistory.DataRange = "一年", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_TwoYear, "两年(&6)")
                        objControl.Checked = IIf(ucHistory.DataRange = "两年", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_ThreeYear, "三年(&7)")
                        objControl.Checked = IIf(ucHistory.DataRange = "三年", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_DateUn, "不限(&F)")
                        objControl.Checked = IIf(ucHistory.DataRange = "不限", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_DateCus, "自定义(&C)")
                        objControl.Checked = IIf(ucHistory.DataRange = "自定义", True, False)
 
                
                Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Helper_RelateCfg, "选项(&O)")
                    cbrControl.ToolTipText = "选项"
                
                    If mobjStudyInfo.lngPatientFrom = 2 Then '只有住院患者，才需要显示本次相关菜单
                        Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_ThisTime, "本次相关(&1)")
                            cbrPopControl.Checked = ucHistory.IsThisTime
                    End If
                    
                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_OtherDept, "他科检查(&2)")
                        cbrPopControl.Checked = ucHistory.IsOtherDept
                        
                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_AutoLine, "自动换行(&3)")
                        cbrPopControl.Checked = ucHistory.IsAutoLine
                        
            Case CON_TAB_TAG_词句 '**********************************************************************************
                blnVisible = ucWord.NodeCount > 0
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_DirectWrite, "直接写入(&E)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = mblnAllowWrite And (IsStudying Or blnAllowAdditionInput)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_EditWrite, "编辑写入(&W)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = mblnAllowWrite And (IsStudying Or blnAllowAdditionInput)
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_FullSave, "全套存入(&S)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_NewWord, "新增词句(&N)")
                    cbrControl.Visible = blnVisible
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ModWord, "修改词句(&M)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucWord.SelNodeType = 2
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_DelWord, "删除词句(&D)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucWord.SelNodeType = 2
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_AutoHide, "自动隐藏(&H)")
                    cbrControl.BeginGroup = True
                    cbrControl.Checked = ucWord.AutoHide
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_DblWrite, "双击写入(&O)")
                    cbrControl.Checked = ucWord.DblWrite
                    
                Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Helper_ExpandLevel, "展开层级(&V)")
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_OneLevel, "一级(&1)")
                        objControl.Checked = IIf(ucWord.ExpandLevel = 1, True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_TwoLevel, "二级(&2)")
                        objControl.Checked = IIf(ucWord.ExpandLevel = 2, True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_ThreeLevel, "三级(&3)")
                        objControl.Checked = IIf(ucWord.ExpandLevel = 3, True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_AllLevel, "所有(&F)")
                        objControl.BeginGroup = True
                        objControl.Checked = IIf(ucWord.ExpandLevel = 0, True, False)
                
        End Select
    End With
    
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub


Private Sub ucImages_OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errhandle
    Dim blnShowImg As Boolean
    Dim intCurImg As Integer
    Dim objDcmViewer As DicomViewer
    Dim objImgInfo As clsBgImgInfo
    
    If mblnMoveBigImageShow = False Then Exit Sub
 

    Set objDcmViewer = ucImages.Viewer

    '判断是否需要显示图像
    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= objDcmViewer.Width) And _
       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= objDcmViewer.Height) Then
        blnShowImg = True
    End If

    If blnShowImg Then        '显示图像
        SetCapture objDcmViewer.hwnd    '锁定鼠标
        
        If mlngStartMoveTime = 0 Then mlngStartMoveTime = GetTickCount
        
        '鼠标移动到图像上的时间打印500毫秒后，才开始显示大图
        If GetTickCount - mlngStartMoveTime < 500 Then
            mlngBigImageIndex = 0
            Exit Sub
        End If
        
        
        intCurImg = objDcmViewer.ImageIndex(X, Y)
        
        Call ucImages.GetImage(intCurImg, objImgInfo)

        If objImgInfo Is Nothing Then Exit Sub


        If intCurImg <> mlngBigImageIndex Then
            If objImgInfo.Format <> ifAvi And objImgInfo.Format <> ifWav Then
            '加载图像并显示
                If objImgInfo.LoadState = lsError Or objImgInfo.LoadState = lsRedo Or objImgInfo.LoadState = lsSent Then
                    Call ShowImageProcess(intCurImg, ptPreview, True)
                Else
                    Call ShowImageProcess(intCurImg, ptPreview)
                End If
            End If
        End If

        mlngBigImageIndex = intCurImg
    Else
        ReleaseCapture
        
        If mblnDelayCloseImage = False Then
            CloseImageProcess
        End If
        
        mlngStartMoveTime = 0
    End If
     
Exit Sub
errhandle:
    Call HintError(err, "ucImages_OnMouseMove", False)
End Sub


Public Sub CloseImageProcess()
    If mobjImageProcessV2 Is Nothing Then Exit Sub
    
    If Not mobjImageProcessV2 Is Nothing Then
        If mobjImageProcessV2.WinType = 0 Then
            Unload mobjImageProcessV2
             
            Set mobjImageProcessV2 = Nothing
        End If
    End If
End Sub


Public Sub ShowImageProcess(ByVal lngImgIndex As Long, ByVal lngType As TImgProcessType, _
    Optional ByVal blnForaceRead As Boolean = False, Optional ByVal blnIsCache As Boolean = False)
    Dim i As Long
    Dim objDcmImg As DicomImage
    Dim objImgInfos() As Object
    Dim blnReportImgState As Boolean
    Dim blnAllowAdditionInput As Boolean
    
    
    If lngImgIndex <= 0 Then Exit Sub
    
    If mobjImageProcessV2 Is Nothing Then
        Set mobjImageProcessV2 = New frmImageProcessV2
    Else
        '不是预览图像时，移动鼠标切换图像不刷新
        If mobjImageProcessV2.WinType <> 0 And lngType = 0 Then Exit Sub
    End If
           
    If blnIsCache Then
        Set objDcmImg = ucCache.GetImage(lngImgIndex)
        
        Call mobjImageProcessV2.SetButtonState(False, False)
        
        ReDim objImgInfos(0)
        
        Call mobjImageProcessV2.ZlShowMe(mObjNotify.Owner, 0, objDcmImg, objImgInfos, ptPreview, 30, mblnAllowWrite)
    Else
        Set objDcmImg = ucImages.GetImage(lngImgIndex)
        
        ReDim objImgInfos(ucImages.ImgCount - 1)
        
        For i = 0 To ucImages.ImgCount - 1
            Set objImgInfos(i) = ucImages.ImageInfo(i)
        Next
         
        blnAllowAdditionInput = (CheckPopedom(mstrPrivs, "补录报告") And mobjStudyInfo.intStep > 5)
        blnReportImgState = mblnAllowWrite And (IsStudying Or blnAllowAdditionInput)
        
        If Not mobjLinkEditor Is Nothing Then
            blnReportImgState = blnReportImgState And mobjLinkEditor.IsEditable
        End If
        
        If blnForaceRead Then
            Call mobjImageProcessV2.SetButtonState(False, False)
        Else
            Call mobjImageProcessV2.SetButtonState(IsStudying, blnReportImgState)
        End If
        
        '在没有任何处理下，2秒后自动关闭大图预览
        If mblnDelayCloseImage Then
            Call mobjImageProcessV2.ZlShowMe(mObjNotify.Owner, mobjStudyInfo.lngAdviceId, objDcmImg, objImgInfos, lngType, 2, mblnAllowWrite)
        Else
            Call mobjImageProcessV2.ZlShowMe(mObjNotify.Owner, mobjStudyInfo.lngAdviceId, objDcmImg, objImgInfos, lngType, 30, mblnAllowWrite)
        End If
    End If
     
    

End Sub


Private Sub ucImages_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
'显示图像右键菜单
On Error GoTo errhandle
    If Button = 2 Then Call ShowPopupMenu(tabSelect.Selected.tag)
Exit Sub
errhandle:
    HintError err, "ucImages_OnMouseUp", False
End Sub

Private Sub ucWord_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'显示词句右键菜单
On Error GoTo errhandle
    If Button = 2 Then Call ShowPopupMenu(tabSelect.Selected.tag)
Exit Sub
errhandle:
    HintError err, "ucWord_OnMouseUp", False
End Sub
 

Private Sub ucWord_OnRequestState(lngOutlineType As TOutlineType, str所见内容 As String, str意见内容 As String, str建议内容 As String)
'获取报告内容信息
    If mobjLinkEditor Is Nothing Then Exit Sub
    
    mobjLinkEditor.GetReportContext str所见内容, str意见内容, str建议内容
    
    lngOutlineType = mobjLinkEditor.CurOutlineType
End Sub

Private Sub ucWord_OnSendContext(ByVal strFreeText As String, ByVal str所见内容 As String, ByVal str意见内容 As String, ByVal str建议内容 As String)
'发送词句内容到报告
    If mobjLinkEditor Is Nothing Then Exit Sub
    
    mobjLinkEditor.InputWord strFreeText, str所见内容, str意见内容, str建议内容
    
End Sub

Private Sub UserControl_Initialize()
    mblnIsValid = False
    mblnIsEmbedVideoArea = True
    mblnAllowEmbedVideo = True
    mblnBgImgTrans = True
    mblnMoveBigImageShow = True
    mblnClickBigImageShow = False
    mblnDelayCloseImage = True
    mblnAllowWrite = True
    
    mlngReleationImgAdvice = 0
    
    Set mobjSel = Nothing
    
End Sub
 


Private Sub UserControl_Resize()
On Error Resume Next
    picBack.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub UserControl_Show()
'On Error Resume Next
'    Call InitTab
End Sub

Public Sub Destory()
    mlngReleationImgAdvice = 0
    
    Call FreeVideo
     
    
    ucSplitter1.Destory
    
    If Not mobjImageProcessV2 Is Nothing Then
        Unload mobjImageProcessV2
    End If
    
    Set mobjImageProcessV2 = Nothing
    
    ucImages.Visible = False
    ucWord.Visible = False
    ucCache.Visible = False
    ucHistory.Visible = False
    
    SetParent ucImages.hwnd, 0
    SetParent ucWord.hwnd, 0
    SetParent ucCache.hwnd, 0
    SetParent ucHistory.hwnd, 0
    
    ucImages.Destory
    ucWord.Destory
    ucCache.Destory
    ucHistory.Destory
    
    tabSelect.RemoveAll
    
    Set mobjSel = Nothing
    Set mobjLinkEditor = Nothing
    Set mobjMainVideo = Nothing
    Set mobjLinkEditor = Nothing
    Set mObjNotify = Nothing
    Set mobjStudyInfo = Nothing
End Sub

Private Sub UserControl_Terminate()
On Error GoTo errhandle
    mblnIsValid = False
    
    Call SavePar
    
    Call Destory
Exit Sub
errhandle:
    Debug.Print "ucPacsHelper_Terminate Err:" & err.Description
End Sub
