Attribute VB_Name = "mdlPubDefine"
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
Public Const conMenu_File_ViewLog = 101        '日志查看(&L)…
Public Const conMenu_File_Exit = 191            '退出(&X)

'编辑菜单
Public Const conMenu_Edit_Reuse = 5006           '启用(&R)
Public Const conMenu_Edit_Disuse = 5007          '停用(&P)

'查看菜单
Public Const conMenu_View_ToolBar = 701              '工具栏(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Public Const conMenu_View_StatusBar = 702            '状态栏(&S)

'帮助菜单
Public Const conMenu_Help_Help = 901        '帮助主题(&H)
Public Const conMenu_Help_Web = 902         '&WEB上的中联
Public Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_Help_Web_Forum = 9023      '中联论坛(&F)
Public Const conMenu_Help_Web_Mail = 9022       '发送反馈(&M)
Public Const conMenu_Help_About = 991       '关于(&A)…
