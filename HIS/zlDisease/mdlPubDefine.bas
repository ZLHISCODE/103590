Attribute VB_Name = "mdlPubDefine"
Option Explicit

Public Const conMenu_FilePopup = 1    '文件
Public Const conMenu_ManagePopup = 2    '管理
Public Const conMenu_EditPopup = 3    '编辑
Public Const conMenu_ReportPopup = 4    '报表
Public Const conMenu_PlugPopup = 6    '外接程序；检验技师工作站使用6100-6199
Public Const conMenu_ViewPopup = 7    '查看
Public Const conMenu_ToolPopup = 8    '工具
Public Const conMenu_HelpPopup = 9    '帮助

'文件菜单
Public Const conMenu_File_Open = 100            '*打开(&O)…
Public Const conMenu_File_PrintSet = 101        '*打印设置(&S)…
Public Const conMenu_File_Preview = 102         '*预览(&V)
Public Const conMenu_File_Print = 103           '*打印(&P)
Public Const conMenu_File_RowPrint = 121        '记录打印(&R)
Public Const conMenu_File_Parameter = 181       '*参数设置(&M)
Public Const conMenu_File_Modify = 3003         '*修改(&M)
Public Const conMenu_File_Delete = 3004         '*删除(&D)
Public Const conMenu_File_Excel = 104           '输出到&Excel…
Public Const conMenu_File_ExportToXML = 192     '另存为XML文档
Public Const conMenu_File_Exit = 191            '退出

'编辑菜单
Public Const conMenu_Edit_Reuse = 3009
Public Const conMenu_Edit_NewItem = 3001    '*新项目(&A)
Public Const conMenu_Edit_Audit = 3010      '*审核/校对(&V)
Public Const conMenu_Edit_Refuse = 3004        '拒绝
Public Const conMenu_Edit_NewTable = 3001    '新增表单
Public Const conMenu_Edit_Add = 3002         '新增备注
Public Const conMenu_Edit_Modify = 3003     '*修改(&M)
Public Const conMenu_Edit_EditInfo = 3564    '*调整病人信息(&E)
Public Const conMenu_Edit_Delete = 3004     '*删除(&D)
Public Const conMenu_Edit_Send = 3013       '*发送(&G)
Public Const conMenu_Edit_Untread = 3014    '*回退(&R)
Public Const conMenu_Edit_ApplyTo = 3062     '*适用科室(&T)
Public Const conMenu_Edit_Request = 3063     '限制要求(&R)
Public Const conMenu_Edit_Compend = 3064     '*内容构造(&F)
Public Const conMenu_Edit_ElementChange = 3065      '*要素联动设置
Public Const conMenu_Edit_Privacy = 3093     '*病人隐私保护设置

'视图菜单
Public Const conMenu_View_ToolBar = 701              '工具栏(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Public Const conMenu_View_StatusBar = 702            '状态栏(&S)
Public Const conMenu_View_DiseaseRegist = 3031
Public Const conMenu_View_Refresh = 791              '*刷新(&R)

'工具菜单
Public Const conMenu_Tool_Send = 3013            '发送(&S)
Public Const conMenu_Tool_Transfer = 213        '转诊
Public Const conMenu_Tool_Finish = 3010          '完成
Public Const conMenu_Tool_OK = 225              '确认为传染病
Public Const conMenu_Tool_NO = 3021                 '非传染病
Public Const conMenu_Tool_ViewReport = 7045          '查看检验检查报告
Public Const conMenu_Tool_Cancel = 3565
Public Const conMenu_Tool_Incept = 252    '接收(&I)
Public Const conMenu_Tool_Refuse = 3004    '拒绝(&R)
Public Const conMenu_Tool_Aduit = 3010

'帮助菜单
Public Const conMenu_Help_Help = 901        '*帮助主题(&H)
Public Const conMenu_Help_Web = 902         '&WEB上的中联
Public Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_Help_Web_Forum = 9023      '中联论坛(&F)
Public Const conMenu_Help_Web_Mail = 9022       '*发送反馈(&M)
Public Const conMenu_Help_About = 991       '关于(&A)…


Public Const HWND_TOPMOST = -1              '最前面
Public Const ID_EDIT_COPY = 323                      '复制

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


