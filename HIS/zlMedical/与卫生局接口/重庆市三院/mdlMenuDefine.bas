Attribute VB_Name = "mdlMenuDefine"
Option Explicit
'公共部份菜单ID定义
'*********************************************************************
Public Const conMenu_FilePopup = 1 '文件
Public Const conMenu_ManagePopup = 2 '管理
Public Const conMenu_EditPopup = 3 '编辑
Public Const conMenu_ReportPopup = 4 '报表
Public Const conMenu_ViewPopup = 7 '查看
Public Const conMenu_ToolPopup = 8 '工具
Public Const conMenu_HelpPopup = 9 '帮助

'文件菜单

Public Const conMenu_File_Login = 100            '*打开(&O)…
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
Public Const conMenu_File_Exit = 191            '*退出(&X)

'编辑菜单
Public Const conMenu_Manage_Regist = 211      '*病人挂号(&H)
Public Const conMenu_Manage_Bespeak = 212     '预约挂号(&B)
Public Const conMenu_Manage_Transfer = 213    '病人转诊(&C)
Public Const conMenu_Manage_Receive = 214     '*病人接诊(&Z)
Public Const conMenu_Manage_Cancel = 215      '取消接诊(&Q)
Public Const conMenu_Manage_Finish = 216      '*完成接诊(&W)
Public Const conMenu_Manage_Redo = 217        '恢复接诊(&R)

'医嘱菜单：因较多,共用时按4位编号,50位分段,001-050,051-100,101-150,...
Public Const conMenu_Edit_Dept = 3001    '*新项目(&A)
Public Const conMenu_Edit_Diagnose = 3002     '*补充/补录(&Y)
Public Const conMenu_Edit_Check = 3003     '*修改(&M)
Public Const conMenu_Edit_Combo = 3004     '*删除(&D)
Public Const conMenu_Edit_Verify = 3005   '*作废(&B)
Public Const conMenu_Edit_Stop = 3006       '*医嘱停止(&S)
Public Const conMenu_Edit_ReStop = 3007     '*确认停止(&C)
Public Const conMenu_Edit_Pause = 3008      '*暂停(&P)
Public Const conMenu_Edit_Reuse = 3009      '*启用(&U)
Public Const conMenu_Edit_Audit = 3010      '*审核/校对(&V)
Public Const conMenu_Edit_Price = 3011      '*计价调整(&I)
Public Const conMenu_Edit_ClearUp = 3012    '*医嘱重整(&F)
Public Const conMenu_Task_Send = 3013       '*发送(&G)
Public Const conMenu_Edit_SendDrug = 30131      '*药疗医嘱发送(&1)
Public Const conMenu_Edit_SendOther = 30132     '其它医嘱发送(&2)
Public Const conMenu_Edit_Untread = 3014    '*回退(&R)
Public Const conMenu_Edit_SendBack = 3015   '*超期发送收回(&N)
Public Const conMenu_Edit_Test = 3016       '*皮试结果(&T)

'病历菜单
Public Const conMenu_Edit_NewParent = 3051   '*新分类(&N)
Public Const conMenu_Edit_Insert = 3052      '*插入(&I)
Public Const conMenu_Edit_MarkMap = 3061     '*图片(&I)…
Public Const conMenu_Edit_ApplyTo = 3062     '*适用科室(&T)
Public Const conMenu_Edit_Request = 3063     '限制要求(&R)
Public Const conMenu_Edit_Compend = 3064     '*内容构造(&F)
Public Const conMenu_Edit_Import = 3071      '*成批导入(&B)…
Public Const conMenu_Edit_Adjust = 3082      '*调整(&J)
Public Const conMenu_Edit_Archive = 3083     '*归档(&R)
Public Const conMenu_Task_Accept = 3091        '*保存
Public Const conMenu_Edit_Sort = 3092        '*多文档排序

'报表菜单
Public Const conMenu_Report_DrugQuery = 401    '药疗收发查询(&H)
Public Const conMenu_Report_Reports = 402      '病区常用报表(&W)
Public Const conMenu_Report_MultiBill = 403    '打印多病人单据(&K)
Public Const conMenu_Report_ClinicBill = 404   '打印诊疗单据(&J)…
Public Const conMenu_Report_AdviceBill1 = 405  '长期医嘱单(&P)
Public Const conMenu_Report_AdviceBill2 = 406  '临时医嘱单(&T)
Public Const conMenu_Report_AdviceBill3 = 407  '医嘱记录本(&B)
Public Const conMenu_Report_WorkLog = 408      '工作日报(&O)
Public Const conMenu_Report_Item = 451         '预留为以后的动态发布报表开始项

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
Public Const conMenu_View_Filter = 731               '*数据过滤(&I)
Public Const conMenu_View_Notify = 732               '*医嘱提醒(&B)
Public Const conMenu_View_Busy = 733                 '诊室忙(&M)
Public Const conMenu_View_Hide = 741                 '*隐藏(&H)
Public Const conMenu_View_Show = 742                 '*显示(&S)
Public Const conMenu_View_Backward = 743             '*后退(&B)
Public Const conMenu_View_Forward = 744              '*前进(&F)
Public Const conMenu_View_Option = 781               '选项(&O)
Public Const conMenu_View_Refresh = 791              '*刷新(&R)
Public Const conMenu_View_Jump = 792                 '跳转(&J)

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
Public Const conMenu_Tool_Option = 819          '选项(&O)

'帮助菜单
Public Const conMenu_Help_Help = 901        '*帮助主题(&H)
Public Const conMenu_Help_Web = 902         '&WEB上的中联
Public Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_Help_Web_Mail = 9022       '*发送反馈(&M)
Public Const conMenu_Help_About = 991       '关于(&A)…

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

'*********************************************************************

