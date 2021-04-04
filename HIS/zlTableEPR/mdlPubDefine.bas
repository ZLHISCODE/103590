Attribute VB_Name = "mdlPubDefine"
Option Explicit
'公共部份菜单ID定义:*表示有图标
'*********************************************************************
Public Const conMenu_FilePopup = 1 '文件
Public Const conMenu_ManagePopup = 2 '管理
Public Const conMenu_EditPopup = 3 '编辑
Public Const conMenu_ReportPopup = 4 '报表
Public Const conMenu_ViewPopup = 7 '查看
Public Const conMenu_ToolPopup = 8 '工具
Public Const conMenu_HelpPopup = 9 '帮助

Public Const conMenu_File_Exit = 191            '*退出(&X)
Public Const conMenu_Edit_NewItem = 3001    '*新项目(&A)
Public Const conMenu_Edit_Modify = 3003     '*修改(&M)
Public Const conMenu_Edit_Delete = 3004     '*删除(&D)
Public Const conMenu_Edit_Insert = 3052      '*插入(&I)
Public Const conMenu_Edit_Save = 3091        '*保存
Public Const conMenu_View_Option = 781               '选项(&O)
Public Const conMenu_View_Refresh = 791              '*刷新(&R)

'帮助菜单
Public Const conMenu_Help_Help = 901        '*帮助主题(&H)
Public Const conMenu_Help_Web = 902         '&WEB上的中联
Public Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_Help_Web_Forum = 9023      '中联论坛(&F)
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


