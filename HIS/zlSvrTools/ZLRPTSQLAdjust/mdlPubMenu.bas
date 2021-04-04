Attribute VB_Name = "mdlPubMenu"
Option Explicit
'公共部份菜单ID定义
'********************************************************************
Public Const conMenu_FilePopup = 1              '文件
Public Const conMenu_ManagePopup = 2            '管理
Public Const conMenu_EditPopup = 3              '编辑
Public Const conMenu_ReportPopup = 4            '报表
Public Const conMenu_ViewPopup = 7              '查看
Public Const conMenu_ToolPopup = 8              '工具
Public Const conMenu_HelpPopup = 9              '帮助

'文件菜单
Public Const conMenu_File_PrintSet = 101        '打印设置(&S)…
Public Const conMenu_File_Preview = 102         '预览(&V)
Public Const conMenu_File_Print = 103           '打印(&P)
Public Const conMenu_File_Excel = 104           '输出到&Excel…
Public Const conMenu_File_Parameter = 181       '参数设置(&M)
Public Const conMenu_File_LogOut = 190            '注消(&L)
Public Const conMenu_File_Exit = 191            '退出(&X)


'编辑菜单
Public Const conMenu_Edit_NewParent = 301       '新分类(&N)
Public Const conMenu_Edit_NewItem = 302         '新项目(&A)
Public Const conMenu_Edit_Modify = 303          '修改(&M)
Public Const conMenu_Edit_Delete = 304          '删除(&D)
Public Const conMenu_Edit_Audit = 305           '审核(&U)
Public Const conMenu_Edit_Blankoff = 306        '作废(&B)
Public Const conMenu_Edit_Disuse = 307          '停用(&P)
Public Const conMenu_Edit_Reuse = 308           '启用(&R)

'查看菜单
Public Const conMenu_View_ToolBar = 701              '工具栏(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Public Const conMenu_View_StatusBar = 702            '状态栏(&S)
Public Const conMenu_View_Expend = 711               '展开/折叠组(&X)
Public Const conMenu_View_Expend_AllCollapse = 7111     '折叠所有组(&L)
Public Const conMenu_View_Expend_AllExpend = 7112       '展开所有组(&X)
Public Const conMenu_View_Expend_CurCollapse = 7113     '折叠当前组(&C)
Public Const conMenu_View_Expend_CurExpend = 7114       '展开当前组(&E)
Public Const conMenu_View_Filter = 721               '过滤(&G)
Public Const conMenu_View_Find = 722                 '查找(&F)
Public Const conMenu_View_FindNext = 723             '查找下一个(&N)
Public Const conMenu_View_Refresh = 791              '刷新(&R)

Public Const conMenu_View_Navigation = 792              '功能导航(&D)

Public Const conMenu_View_Property = 793              '属性(&P)


'帮助菜单
Public Const conMenu_Help_Help = 901        '帮助主题(&H)
Public Const conMenu_Help_Web = 902         '&WEB上的中联
Public Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_Help_Web_Mail = 9022       '发送反馈(&M)
Public Const conMenu_Help_About = 991       '关于(&A)…

'快捷键
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

