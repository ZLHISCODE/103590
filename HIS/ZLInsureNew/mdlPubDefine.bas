Attribute VB_Name = "mdlPubDefine"
Option Explicit
Public Const gstrSplitCmb = "-"

'公共部份菜单ID定义:*表示有图标
'*********************************************************************
Public Const conMenu_FilePopup = 1 '文件
Public Const conMenu_ManagePopup = 2 '管理
Public Const conMenu_EditPopup = 3 '编辑
Public Const conMenu_ReportPopup = 4 '报表
Public Const conMenu_ViewPopup = 7 '查看
Public Const conMenu_ToolPopup = 8 '工具
Public Const conMenu_HelpPopup = 9 '帮助
Public Const gconLockColor = &H80000000
Public Const gconEditColor = &HC0FFC0
'文件菜单
Public Const conMenu_File_Open = 100            '*打开(&O)…
Public Const conMenu_File_PrintSet = 101        '*打印设置(&S)…
Public Const conMenu_File_Preview = 102         '*预览(&V)
Public Const conMenu_File_Print = 103           '*打印(&P)
Public Const conMenu_File_Excel = 104           '输出到&Excel…
Public Const conMenu_File_MedRec = 105          '首页打印(&R)
Public Const conMenu_File_MedRecSetup = 1051        '打印设置(&S)
Public Const conMenu_File_MedRecPreview = 1052      '打印预览(&P)
Public Const conMenu_File_MedRecPrint = 1053        '打印首页(&V)
Public Const conMenu_File_RowPrint = 121        '记录打印(&R)
Public Const conMenu_File_BatPrint = 122        '批量打印(&B)
Public Const conMenu_File_Parameter = 181       '*参数设置(&M)
Public Const conMenu_File_RoomSet = 182         '执行间设备
Public Const conMenu_File_SendImg = 184         '发送图像
Public Const conMenu_File_Exit = 191            '*退出(&X)
Public Const conMenu_File_ExportToXML = 192     '另存为XML文档
Public Const conMenu_File_ImportFromXML = 193   '从XML文档导入
Public Const conMenu_File_BillPrintView = 194   '单据打印预览
Public Const conMenu_File_BillPrint = 195       '单据打印

'管理菜单:工作站自身的功能菜单
Public Const conMenu_Manage_Regist = 211      '*病人挂号(&H)
Public Const conMenu_Manage_Bespeak = 212     '预约挂号(&B),时间安排(&B)
Public Const conMenu_Manage_Transfer = 213    '*转诊处理(&C)
Public Const conMenu_Manage_Transfer_Send = 2131      '病人转诊(&S)
Public Const conMenu_Manage_Transfer_Cancel = 2132    '取消转诊(&C)
Public Const conMenu_Manage_Transfer_Incept = 2133    '转诊接收(&I)
Public Const conMenu_Manage_Transfer_Refuse = 2134    '转诊拒绝(&R)
Public Const conMenu_Manage_Transfer_Force = 2135     '强制续诊(&F)
Public Const conMenu_Manage_Receive = 214     '*病人接诊(&Z)
Public Const conMenu_Manage_Cancel = 215      '取消接诊(&Q)
Public Const conMenu_Manage_Finish = 216      '*完成接诊(&W)
Public Const conMenu_Manage_Redo = 217        '恢复接诊(&R)

Public Const conMenu_Manage_Call = 218      '呼叫
Public Const conMenu_Manage_CallNext = 21801      '下一个(&N)
Public Const conMenu_Manage_CallPrevious = 21802     '上一个(&P)

Public Const conMenu_Manage_Reset = 219     '调整顺序
Public Const conMenu_Manage_Up = 21901    '上移(&U)
Public Const conMenu_Manage_Down = 21902     '下移(&D)
Public Const conMenu_Manage_Discard = 21903      '弃号(&D)
Public Const conMenu_Manage_Recall = 21904      '召回(&R)
Public Const conMenu_Manage_Untread = 21905        '退号(&R)

Public Const conMenu_Manage_Plan = 221        '*执行报到(&P)
Public Const conMenu_Manage_Logout = 222      '取消报到(&L)
Public Const conMenu_Manage_Refuse = 223      '拒绝执行(&R)
Public Const conMenu_Manage_ReGet = 224       '取消拒绝(&G)
Public Const conMenu_Manage_Complete = 225    '*执行完成(&C)
Public Const conMenu_Manage_Undone = 226      '取消完成(&U)
Public Const conMenu_Manage_ThingAdd = 227    '*记录执行情况(&A)
Public Const conMenu_Manage_ThingModi = 228   '*调整执行情况(&M)
Public Const conMenu_Manage_ThingDel = 229    '*删除执行情况(&D)
Public Const conMenu_Manage_ClearUp = 233     '检查报告驳回(&U)

Public Const conMenu_Manage_Request = 231        '*申请(&V)
Public Const conMenu_Manage_RequestView = 2311           '查阅申请(&V)
Public Const conMenu_Manage_RequestPrint = 2312           '打印诊疗单据(&J)
Public Const conMenu_Manage_RequestBatPrint = 2313           '批量打印条码(&B)
Public Const conMenu_Manage_Report = 232         '*报告(&O)
Public Const conMenu_Manage_ReportEdit = 2321        '填写报告(&E)
Public Const conMenu_Manage_ReportView = 2322        '查阅报告(&W)
Public Const conMenu_Manage_ReportPrint = 2323       '报告打印(&P)
Public Const conMenu_Manage_ReportPreview = 2324     '执行预览(&V)
Public Const conMenu_Manage_LeaveMedi = 251 '寄存药品

Public Const conMenu_Manage_Audit = 252         '*审核申请
Public Const conMenu_Manage_UnAudit = 253       '*取消审核
Public Const conMenu_Manage_Arrange = 254       '*执行安排
Public Const conMenu_Manage_UnArrange = 255     '*取消安排

'医嘱(编辑)菜单：因较多,共用时按4位编号,50位分段,001-050,051-100,101-150,...
Public Const conMenu_Edit_NewItem = 3001    '*新项目(&A)
Public Const conMenu_Edit_Append = 3002     '*补充/补录(&Y)
Public Const conMenu_Edit_Modify = 3003     '*修改(&M)
Public Const conMenu_Edit_Delete = 3004     '*删除(&D)
Public Const conMenu_Edit_Blankoff = 3005   '*作废(&B)
Public Const conMenu_Edit_Stop = 3006       '*医嘱停止(&S)
Public Const conMenu_Edit_ReStop = 3007     '*确认停止(&C)
Public Const conMenu_Edit_Pause = 3008      '*暂停(&P)
Public Const conMenu_Edit_Reuse = 3009      '*启用(&U)
Public Const conMenu_Edit_Audit = 3010      '*审核/校对(&V)
Public Const conMenu_Edit_Price = 3011      '*计价调整(&I)
Public Const conMenu_Edit_ClearUp = 3012    '*医嘱重整(&F)
Public Const conMenu_Edit_Send = 3013       '*发送(&G)
Public Const conMenu_Edit_SendDrug = 30131      '*药疗医嘱发送(&1)
Public Const conMenu_Edit_SendOther = 30132     '其它医嘱发送(&2)
Public Const conMenu_Edit_Untread = 3014    '*回退(&R)
Public Const conMenu_Edit_SendBack = 3015   '*超期发送收回(&N)
Public Const conMenu_Edit_Test = 3016       '*皮试结果(&T)
Public Const conMenu_Edit_ChargeOff = 3017       '*费用冲销(&E)
Public Const conMenu_Edit_NoPrint = 3018    '屏蔽打印(&I)
Public Const conMenu_Edit_ChargeDelApply = 3019 '*销帐申请(&L)
Public Const conMenu_Edit_ChargeDelAudit = 3020 '*销帐审核(&U)

'病历(编辑)菜单
Public Const conMenu_Edit_NewParent = 3051   '*新分类(&N)
Public Const conMenu_Edit_Insert = 3052      '*插入(&I)
Public Const conMenu_Edit_ModifyParent = 3053 '*修改分类(&M)
Public Const conMenu_Edit_DeleteParent = 3054 '*删除分类(&D)
Public Const conMenu_Edit_MarkMap = 3061     '*图片(&I)…
Public Const conMenu_Edit_ApplyTo = 3062     '*适用科室(&T)
Public Const conMenu_Edit_Request = 3063     '限制要求(&R)
Public Const conMenu_Edit_Compend = 3064     '*内容构造(&F)
Public Const conMenu_Edit_Import = 3071      '*成批导入(&B)…
Public Const conMenu_Edit_Adjust = 3082      '*调整(&J)
Public Const conMenu_Edit_Archive = 3083     '*归档(&R)
Public Const conMenu_Edit_UnArchive = 3084     '取消归档(&D)
Public Const conMenu_Edit_Save = 3091        '*保存
Public Const conMenu_Edit_Sort = 3092        '*多文档排序
Public Const conMenu_Edit_Privacy = 3093     '*病人隐私保护设置

Public Const conMenu_Edit_Select = 3094      '*选择
Public Const conMenu_Edit_DeSelect = 3095    '*取消选择

Public Const conMenu_Edit_Merge = 3096

'Public Const conMenu_Manage_ThingAdd = 227    '接单(&A)
'Public Const conMenu_Manage_ThingModi = 228   '*调整执行情况(&M)
Public Const conMenu_Edit_Transf_Delete = 229   '撤消接单

'体检系统补增 32开头的号
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_SaveExit = 3200           '保存并退出
Public Const conMenu_Edit_SizeFit = 3201            '格式调整
Public Const conMenu_Edit_SourceFit = 3202          '来源调整
Public Const conMenu_Edit_Camera = 3203             '照相设备
Public Const conMenu_Edit_TakePicture = 3204        '拍照
Public Const conMenu_Edit_SelAll = 3205             '全选
Public Const conMenu_Edit_ClsAll = 3206             '全清
Public Const conMenu_Edit_CallBack = 3207           '复查设置
Public Const conMenu_Edit_Money = 3208              '费用方式
Public Const conMenu_Edit_Pay = 3209                '支付方式
Public Const conMenu_Edit_CheckItem = 3210          '体检项目
Public Const conMenu_Edit_ChargeItem = 3211         '收费项目

'病人项目(编辑)菜单 3501-3530
Public Const conMenu_Edit_Transf_Modify = 3502   '修改单据
Public Const conMenu_Edit_Transf_Save = 3503     '保存
Public Const conMenu_Edit_Transf_Cancle = 3504   '取消

Public Const conMenu_Edit_Transf_UndoEnd = 3505  '撤消完成
Public Const conMenu_Edit_Transf_Negative = 3506 '阳性(+)
Public Const conMenu_Edit_Transf_Positive = 3507 '阴性(-)
Public Const conMenu_Edit_Transf_Reprint = 3508  '重打单据

'病人座位(编辑)菜单 3531-3559
Public Const conMenu_Edit_Seat = 3530        '座位
Public Const conMenu_Edit_Seat_Add = 3531    '座位增加
Public Const conMenu_Edit_Seat_Modify = 3532 '座位修改
Public Const conMenu_Edit_Seat_Delete = 3533 '座位删除
Public Const conMenu_Edit_Seat_Clear = 3534  '清除占用的座位
Public Const conMenu_Edit_Seat_Set = 3535    '安排座位
Public Const conMenu_Edit_Seat_Swap = 3536    '调换座位

Public Const conMenu_Edit_Seat_View = 3551 '查看
Public Const conMenu_Edit_Seat_Icon = 3552 '图标方式
Public Const conMenu_Edit_Seat_List = 3553 '列表方式
Public Const conMenu_Edit_Seat_Report = 3554 '报表方式

'暂存药品(编辑)菜单 3561 -3579
Public Const conMenu_Edit_Leave_Add = 3561 '增加
Public Const conMenu_Edit_Leave_Modify = 3562 '修改
Public Const conMenu_Edit_Leave_Delete = 3563 '删除
Public Const conMenu_Edit_Leave_Post = 3564 '使用登记
Public Const conMenu_Edit_Leave_SavePost = 3565 '保存登记数据
Public Const conMenu_Edit_Leave_UndoPost = 3565 '撤消登记

Public Const conMenu_Edit_Leave_Repertory = 3571 '库存查询
Public Const conMenu_Edit_Leave_AccountBook = 3572 '库存台帐

'手麻系统补增 3580 -  3599
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_CopyNewItem = 3580        '*复制并新项目
Public Const conMenu_Edit_Default = 3582            '缺省结果
Public Const conMenu_Edit_MakeCharge = 3586         '生成费用
Public Const conMenu_Edit_Preferences = 3587         '参考方案

'血库系统补增 31开头的号
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_NewKind = 311             '新增品种
Public Const conMenu_Edit_ModifyKind = 312          '修改品种
Public Const conMenu_Edit_DeleteKind = 313          '删除品种
Public Const conMenu_Edit_StorgeLimit = 314         '库存限量
Public Const conMenu_Edit_StorgeDept = 315          '库房
Public Const conMenu_Edit_StorgePostion = 316       '货位
Public Const conMenu_Edit_Check = 3101              '核对
Public Const conMenu_Edit_View = 3102               '查阅
Public Const conMenu_Edit_ModifyBill = 3103         '修改发票
Public Const conMenu_Edit_Verify = 3104             '常规检验
Public Const conMenu_Edit_AdjustPrice = 3105        '调价

'LIS使用的采单
Public Const conMenu_Edit_QCRes = 3650         '质控品

'报表菜单
Public Const conMenu_Report_DrugQuery = 401    '药疗收发查询(&H)
Public Const conMenu_Report_Reports = 402      '病区常用报表(&W)
Public Const conMenu_Report_MultiBill = 403    '打印多病人单据(&K)
Public Const conMenu_Report_ClinicBill = 404   '打印诊疗单据(&J)…
Public Const conMenu_Report_AdviceBill1 = 405  '长期医嘱单(&P)
Public Const conMenu_Report_AdviceBill2 = 406  '临时医嘱单(&T)
Public Const conMenu_Report_AdviceBill3 = 407  '医嘱记录本(&B)
Public Const conMenu_Report_WorkLog = 408      '工作日报(&O)


'查看菜单
Public Const conMenu_View_ToolBar = 701              '工具栏(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Public Const conMenu_View_StatusBar = 702            '状态栏(&S)
Public Const conMenu_View_Append = 703               '附加信息(&A)
Public Const conMenu_View_Expend = 711               '展开/折叠组(&X)
Public Const conMenu_View_Expend_CurCollapse = 7111     '折叠当前组(&C)
Public Const conMenu_View_Expend_CurExpend = 7112       '展开当前组(&E)
Public Const conMenu_View_Expend_AllCollapse = 7113     '折叠所有组(&L)
Public Const conMenu_View_Expend_AllExpend = 7114       '展开所有组(&X)
Public Const conMenu_View_Find = 721                 '*查找(&F)
Public Const conMenu_View_FindNext = 722             '继续查找(&N)
Public Const conMenu_View_FindType = 723             '查找方式(&Y)
Public Const conMenu_View_ReadIC = 724               '读IC卡(&I)
Public Const conMenu_View_PatInfor = 725             '查看病人信息
Public Const conMenu_View_PriceBill = 727
Public Const conMenu_View_PriceTable = 728
Public Const conMenu_View_PriceList = 729
Public Const conMenu_View_FilterView = 730           '以过滤方式显示
Public Const conMenu_View_Filter = 731               '*数据过滤(&I),子窗体的过滤功能
Public Const conMenu_View_Notify = 732               '*医嘱提醒(&B)
Public Const conMenu_View_Busy = 733                 '诊室忙(&M)
Public Const conMenu_View_ShowAll = 734
Public Const conMenu_View_ShowHistory = 735
Public Const conMenu_View_ShowStoped = 736
Public Const conMenu_View_Hide = 741                 '*隐藏(&H)
Public Const conMenu_View_Show = 742                 '*显示(&S)
Public Const conMenu_View_Forward = 743              '*前进(&F)
Public Const conMenu_View_Backward = 744             '*后退(&B)
Public Const conMenu_View_Dept = 745                '查看部门
Public Const conMenu_View_Location = 746            '定位
Public Const conMenu_View_LocationItem = 747        '定位项目
Public Const conMenu_View_Option = 781               '选项(&O)
Public Const conMenu_View_Refresh = 791              '*刷新(&R)
Public Const conMenu_View_Jump = 792                 '跳转(&J)

'体检系统补增70开头的号
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_View_Single = 7040             '个人
Public Const conMenu_View_Group = 7041              '团体
Public Const conMenu_View_LocationMethod = 7042     '定位处理
Public Const conMenu_View_Column = 7043             '选择列项

'工具菜单
Public Const conMenu_Tool_Reference = 801       '*参考(&R)
Public Const conMenu_Tool_Reference_1 = 8011    '疾病诊断参考(&D)
Public Const conMenu_Tool_Reference_2 = 8012    '诊疗措施参考(&C)
Public Const conMenu_Tool_MedRec = 802          '*首页整理(&M)
Public Const conMenu_Tool_Meet = 803            '*病人会诊(&E)
Public Const conMenu_Tool_MeetFinish = 8031         '完成会诊(&F)
Public Const conMenu_Tool_MeetCancel = 8032         '取消完成(&C)
Public Const conMenu_Tool_Sign = 804            '*电子签名(&I)
Public Const conMenu_Tool_SignNew = 8041            '电子签名(&I)
Public Const conMenu_Tool_SignVerify = 8042         '验证签名(&V)
Public Const conMenu_Tool_SignEarse = 8043          '取消签名(&E)
Public Const conMenu_Tool_Monitor = 811         '*监测(&M)
Public Const conMenu_Tool_Monitor_1 = 81101         '时限要求监测(&T)
Public Const conMenu_Tool_Monitor_2 = 81102         '内容要求监测(&C)
Public Const conMenu_Tool_Assistant = 812       '*助手(&A)
Public Const conMenu_Tool_Analyse = 813         '*分析(&Y)
Public Const conMenu_Tool_Search = 814          '*检索(&S)
Public Const conMenu_Tool_Define = 815          '*定义(&D)
Public Const conMenu_Tool_Report = 816          '*报告(&P)
Public Const conMenu_Tool_Apply = 817           '*应用(&A)
Public Const conMenu_Tool_Option = 819          '选项(&O),子窗体的设置功能


'采集菜单
Public Const conMenu_Cap_Dynamic = 8100         '动态显示(&V)
Public Const conMenu_Cap_MarkMap = 8101       '影像采集(&C)
Public Const conMenu_Cap_Import = 8102        '影像导入(&I)
Public Const conMenu_Cap_DevSet = 8103          '影像设备设置(&D)
Public Const comMenu_Cap_Process = 8104         '影像处理
Public Const conMenu_Cap_Record = 8105          '录像(&R)
Public Const conMenu_Cap_Play = 8106          '播放(&P)
Public Const conMenu_Cap_Stop = 8107            '停止(&T)
Public Const conMenu_Cap_Forward = 8108         '快进(&F)
Public Const conMenu_Cap_Back = 8109            '快退(&B)
Public Const conMenu_Cap_SaveAs = 8110          '保存录像(&S)


Public Const conMenu_Img_Look = 8111        '影像观片(&S)
Public Const conMenu_Img_Contrast = 8112    '观片对比(&E)
Public Const conMenu_Img_Delete = 8113        '图象删除(&K)
Public Const conMenu_Img_Query = 8114        'Q/R获取图象(&Q)



'图像处理
Public Const conMenu_Process_Window = 501           '亮度对比度
Public Const conMenu_Process_Zoom = 502             '缩放
Public Const conMenu_Process_RRotate = 503          '顺时针旋转
Public Const conMenu_Process_LRotate = 504          '逆时针旋转
Public Const conMenu_Process_Sharpness = 505        '锐化
Public Const conMenu_Process_Filter = 506           '平滑
Public Const conMenu_Process_Arrow = 507            '箭头标注
Public Const conMenu_Process_Ellipse = 508          '圆形标注
Public Const conMenu_Process_Text = 509             '文字标注


'帮助菜单
Public Const conMenu_Help_Help = 901        '*帮助主题(&H)
Public Const conMenu_Help_Web = 902         '&WEB上的中联
Public Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_Help_Web_Forum = 9023      '中联论坛(&F)
Public Const conMenu_Help_Web_Mail = 9022       '*发送反馈(&M)
Public Const conMenu_Help_About = 991       '关于(&A)…

Public Const conMenu_Edit_MediAudit = 3564 '*药嘱审查(&U)(合理用药审查)

'其它常量定义
'*********************************************************************
'CommandBar固有常量定义
Public Const XTP_ID_WINDOW_LIST = 35000 '窗体列表
Public Const XTP_ID_TOOLBARLIST = 59392 '工具栏列表
Public Const ID_INDICATOR_CAPS = 59137 '状态栏（大写）
Public Const ID_INDICATOR_NUM = 59138 '状态栏（数字）
Public Const ID_INDICATOR_SCRL = 59139 '状态栏（滚动）

'CommandBar辅助热键
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

'CommandBar虚拟键
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_PAGEUP = &H21
Public Const VK_PAGEDOWN = &H22
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B

Public Const VsModiBackColor = &HD6FFCA        'vs控件，可编辑单元的背景色
'*********************************************************************
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Enum REGISTER
    注册信息
    私有模块
    私有全局
    公共模块
    公共全局
End Enum
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
     x As Long
     y As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hwnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
End Sub
Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption)
        
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        
    End With
    
    Set NewCommandBar = objControl
    
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.ID = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function


Public Function DockPannelCreate(ByRef dkpMain As DockingPane, ByVal intIndex As Integer, _
                                    ByVal lngCX As Long, ByVal lngCY As Long, _
                                    ByVal bytDirection As DockingDirection, _
                                    Optional ByVal objNeighbour As Pane = Nothing, _
                                    Optional ByVal strTitle As String, _
                                    Optional ByVal bytOptions As PaneOptions) As Pane
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Set DockPannelCreate = dkpMain.CreatePane(intIndex, lngCX, lngCY, bytDirection, objNeighbour)
    DockPannelCreate.Title = strTitle
    DockPannelCreate.Options = PaneNoCaption
    
End Function

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Public Function TabControlInit(ByRef tbc As TabControl, _
                                Optional ByVal bytAppearance As XTPTabAppearanceStyle = xtpTabAppearancePropertyPage2003) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With tbc
        
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
'            .Position = bytPosition
        End With
        
        Set .Icons = frmPubResource.imgPublic.Icons
        

        
    End With

    TabControlInit = True
    
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto 'xtpSystemThemeBlue
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    cbsMain.Options.LargeIcons = False
    
    CommandBarInit = True
    
End Function

Public Function CopyMenu(ByVal cbsMain As Object, Optional ByVal intNo As Integer = 2) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '弹出菜单处理
    
    On Error GoTo errHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Set cbrPopupItem2 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.ID, cbrControl2.Caption)
                cbrPopupItem2.Parameter = cbrControl2.Parameter
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
errHand:
    Err.Clear
    Resume Next
End Function

Public Sub SetDockRight(cbsMain As Object, BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsMain.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsMain.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

Public Sub LocationObj(ByRef objTxt As Object, Optional ByVal blnDoevents As Boolean = False)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If blnDoevents Then DoEvents
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
    
errHand:
    
End Sub

Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button            '工具栏
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              '图标文字
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '大图标
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet              '打印设置
    
        Call zlPrintSet
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel               '打印数据,预览数据,输出到Excel
        
        If objPrnVsf Is Nothing Then Exit Function
        
        If Not SearchPrintData(objPrnVsf, frmPubResource.msfPrint) Then
            MsgBox "你打印的网络不存在数据，请重新检视！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '调用打印部件处理
        Set objPrint.Body = frmPubResource.msfPrint
        objPrint.Title.Text = strPrintTitle
        Set objAppRow = New zlTabAppRow
        Call objAppRow.Add("")
        Call objAppRow.Add("打印时间:" & Now())
        Call objPrint.BelowAppRows.Add(objAppRow)

        Select Case Control.ID
        Case conMenu_File_Print
            bytMode = zlPrintAsk(objPrint)
            If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
        Case conMenu_File_Preview
            zlPrintOrView1Grd objPrint, 2
        Case conMenu_File_Excel
            zlPrintOrView1Grd objPrint, 3
        End Select
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '按钮文字
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '大图标
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '状态栏
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
    
    Case conMenu_Help_Help              '帮助主题
    
        Call ShowHelp(App.ProductName, frmMain.hwnd, frmMain.Name, Int((glngSys) / 100))
        
    Case conMenu_Help_Web_Home          'Web上的中联
        
        Call zlHomePage(frmMain.hwnd)
        
    Case conMenu_Help_Web_Forum         'Web上的论坛
    
        Call zlWebForum(frmMain.hwnd)
        
    Case conMenu_Help_Web_Mail          '发送反馈
        
        Call zlMailTo(frmMain.hwnd)
            
    Case conMenu_Help_About             '关于
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_File_Exit              '退出
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function


Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strFormat As String
    Dim lngNotPrintCols As Long
    Dim lngPrintCol As Long
    
    If objPrintVsf.Cols = 0 Then Exit Function
    If strNotPrintCol <> "" Then
        lngNotPrintCols = UBound(Split(strNotPrintCol, ",")) + 1
        strNotPrintCol = "," & strNotPrintCol & ","
    End If
    
    objPrintVsf.Rows = objVsf.Rows
    objPrintVsf.FixedRows = objVsf.FixedRows
    
    objPrintVsf.Cols = 0
    lngPrintCol = -1
    For lngCol = 0 To objVsf.Cols - 1
        
        If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
            
            If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                lngPrintCol = lngPrintCol + 1
                
                objPrintVsf.Cols = lngPrintCol + 1
                
                objPrintVsf.ColWidth(lngPrintCol) = objVsf.ColWidth(lngCol)
                objPrintVsf.ColAlignmentFixed(lngPrintCol) = objVsf.ColAlignment(lngCol)
                If objVsf.ColDataType(lngCol) = flexDTBoolean Then
                    objPrintVsf.ColAlignment(lngPrintCol) = 4
                Else
                    objPrintVsf.ColAlignment(lngPrintCol) = objVsf.ColAlignment(lngCol)
                End If
            End If
        End If
    Next
    
    If objPrintVsf.Cols = 0 Then Exit Function
    
    For lngRow = 0 To objVsf.Rows - 1

        objPrintVsf.RowHeight(lngRow) = IIf(objVsf.RowHeight(lngRow) < objVsf.RowHeightMin, objVsf.RowHeightMin, objVsf.RowHeight(lngRow))
        lngPrintCol = -1
        For lngCol = 0 To objVsf.Cols - 1
            
            If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
                If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                    lngPrintCol = lngPrintCol + 1
                    
                    If objVsf.ColDataType(lngCol) = flexDTBoolean And lngRow >= objVsf.FixedRows Then
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = IIf(Abs(Val(objVsf.TextMatrix(lngRow, lngCol))) = 1, "√", "")
                    Else
                        strFormat = objVsf.ColFormat(lngCol)
                        If strFormat = "" Then
                            objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Trim(objVsf.TextMatrix(lngRow, lngCol))
                        Else
                            objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Format(objVsf.TextMatrix(lngRow, lngCol), strFormat)
                        End If
                    End If
                End If
            End If
        Next
        Call SetMsfForeColor(objPrintVsf, lngRow, Val(objVsf.Cell(flexcpForeColor, lngRow, 1)))
    Next
    SearchPrintData = True
End Function

Public Sub SetMsfForeColor(ByRef msf As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCol As Integer

    With msf

        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .COL = intCol
            .CellForeColor = lngColor
        Next

    End With
End Sub


Public Sub SendLMouseButton(ByVal lngHwnd As Long, ByVal x As Single, ByVal y As Single)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngX As Long
    Dim lngY As Long
    Dim lngLoop As Long
    Dim lngXY As Long
            
    lngX = x / 15
    lngY = y / 15
        
    lngXY = 2
    For lngLoop = 1 To 15
        lngXY = lngXY * 2
    Next
    
    lngXY = lngXY * lngY + lngX
    
    SendMessage lngHwnd, WM_LBUTTONDOWN, 0, ByVal lngXY
    SendMessage lngHwnd, WM_LBUTTONUP, 0, ByVal lngXY

End Sub

Public Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    GetPersonSet = False
    If Val(GetPara("使用个性化风格")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '功能： 将指定的信息保存在注册表中
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strKeyValue-键值
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        Call SaveSetting("ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue)
        
    Case 私有模块

        Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case 私有全局

        Call SaveSetting("ZLSOFT", "私有全局\" & UserInfo.用户名 & "\" & strSection, strKey, strKeyValue)
        
    Case 公共模块

        Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case 公共全局
        
        Call SaveSetting("ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue)
        
    End Select
    
    SetRegister = True
    
errHand:
    
End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '功能： 将指定的注册信息读取出来
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strDefKeyValue-缺省键值
    '返回： strKeyValue-键值
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        strValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, strDefKeyValue)
        
    Case 私有模块

        strValue = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 私有全局

        strValue = GetSetting("ZLSOFT", "私有全局\" & UserInfo.用户名 & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共模块

        strValue = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共全局
        
        strValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function


'========================================================================================
'=名稱:检测(ChkRsState)
'=入口参数:Rs               类型:ADODB.Recordset
'=出口参数:ChkRsState       类型:Boolean
'=功能:检测记录集的状态
'=日期:2004-07-08
'=編程:谢荣
'========================================================================================
Function ChkRsState(rs As ADODB.Recordset) As Boolean
On Error GoTo ErrH:
    With rs
        If rs Is Nothing Then
            ChkRsState = True
            Exit Function
        Else
            ChkRsState = False
        End If
        If rs.State = 0 Then
            ChkRsState = True
            Exit Function
        Else
            ChkRsState = False
        End If
        If .RecordCount < 1 Then
            ChkRsState = True
        Else
            ChkRsState = False
        End If
        If .EOF Or .BOF Then
            ChkRsState = True
        Else
            ChkRsState = False
        End If
    End With
    Exit Function
ErrH:
    Err.Clear
End Function


'==================================================================================================
'=名称:去掉字符串中的单引号("'")(ConvertString)
'=入口参数:
'=1).sStr          类型:String
'=出口参数:空
'=功能:去掉字符串(sStr)中的单引号
'=日期:2010-12-11
'=编程:谢荣
'=说明:在SQL语句中不能带单引号
'==================================================================================================
Public Function ConvertString(ByVal sStr As String) As String
    Dim i               As Integer
    Dim strReturn       As String
    Dim strSystemChar   As String
On Error GoTo ErrH
    strSystemChar = "'|[]"
    '检测系统不许录入字符
    For i = 1 To Len(strSystemChar)
        sStr = Replace(sStr, Mid(strSystemChar, i, 1), "")
    Next
    strReturn = sStr
    ConvertString = strReturn
    Exit Function
ErrH:
    Err.Clear
    ConvertString = ""
End Function

'==================================================================================================
'检测长度是否超过长度(字节数)
'==================================================================================================
Public Function ChkStrUniCode(mStr As String, mLen As Long) As String
    Dim strL        As String
On Error GoTo ErrH
    mStr = ConvertString(mStr)
    If mLen <= 0 Then
        ChkStrUniCode = mStr
        Exit Function
    Else
        strL = StrConv(mStr, vbFromUnicode)
        strL = LeftB(strL, mLen)
        ChkStrUniCode = StrConv(strL, vbUnicode)
    End If
    Exit Function
ErrH:
    Err.Clear
    ChkStrUniCode = ""
    Exit Function
End Function

'==================================================================================================
'=名称:得到下拉列表框的Text属性中取得ID(Cmb_ID)
'=入口参数:
'=1).下拉列表框控件         类型:Control
'=出口参数:空
'=功能:得到下拉列表框的Text属性中取得ID
'=日期:2004-12-11
'=编程:谢荣
'=说明:在原因类别ID中的数据不能带"-"
'==================================================================================================
Function Cmb_ID(Combo As Object, Optional Index As Byte = 1) As String
    Dim xx          As Variant
On Error GoTo ErrH
    If Combo.Text = "" Then
        Cmb_ID = ""
    Else
        xx = Split(Combo.Text, gstrSplitCmb)
        If Index - 1 <= UBound(xx) Then '最大下标值小于输入值[证明有截取值]
            Cmb_ID = xx(Index - 1)
        Else                        '最大下标值大于等于输入值[证明有无截取值]返回无
            Cmb_ID = "[无]"
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function


'==================================================================================================
'=名称:得到下拉列表框的Text属性中取得ID(Cmb_EditIndex)
'=入口参数:
'=1).下拉列表框控件         类型:Control
'=出口参数:空
'=功能:得到下拉列表框的Text属性中取得ID
'=日期:2004-12-11
'=编程:谢荣
'=说明:在原因类别ID中的数据不能带"-"
'==================================================================================================
Function Cmb_EditIndex(Combo As Object, sID As String) As Long
    Dim lngCount    As Long
    Dim lngStep     As Long
    Dim xx          As Variant
On Error GoTo ErrH
    lngCount = Combo.ListCount - 1
    For lngStep = 0 To lngCount
        xx = Split(Combo.List(lngStep) & gstrSplitCmb, gstrSplitCmb)
        If sID = xx(0) Then
            Cmb_EditIndex = lngStep
            Exit For
        End If
    Next
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'========================================================================================
'=功能：查询网络数据并定位
'=入参：1、objVsf VSFlexGrid网格
'=      2、strFind 查找字符串
'=      3、查询列集，多列用“,”分隔
'========================================================================================
Public Sub vsfSetRow(ByRef objVsf As VSFlexGrid, ByVal strFind As String, ByVal strCols As String)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    Dim varCols     As Variant
    Dim strCol      As String
    Dim blnExit     As Boolean
    
    varCols = Split(strCols, ",")
    blnExit = False
    
    '读取大于当前行的记录数据
    For lngLoop = objVsf.Row + 1 To objVsf.Rows - 1
        For intCol = 0 To UBound(varCols)
            strCol = varCols(intCol)
            If InStr(UCase(objVsf.TextMatrix(lngLoop, objVsf.ColIndex(strCol))), UCase(strFind)) > 0 Then
                lngRow = lngLoop
                blnExit = True
                Exit For
            End If
        Next
        If blnExit Then Exit For
    Next
    
    '读取小于当前行的记录数据
    If lngRow = 0 Then
        For lngLoop = 0 To objVsf.Row
            For intCol = 0 To UBound(varCols)
                strCol = varCols(intCol)
                If InStr(UCase(objVsf.TextMatrix(lngLoop, objVsf.ColIndex(strCol))), UCase(strFind)) > 0 Then
                    lngRow = lngLoop
                    blnExit = True
                    Exit For
                End If
            Next
            If blnExit Then Exit For
        Next
    End If
    If objVsf.Rows > 1 And lngRow >= 1 Then objVsf.Row = lngRow
    DoEvents
    objVsf.ShowCell lngRow, 1
End Sub





