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
Public Const conMenu_File_Preview = 102         '预览交接班记录(&V)
Public Const conMenu_File_Print = 103           '打印交接班记录(&P)
Public Const conMenu_File_Excel = 104           '输出到Excel...
Public Const conMenu_File_TypeManage = 105      '班次管理
Public Const conMenu_File_Exit = 191            '退出(&X)

'编辑菜单
Public Const conMenu_Edit_NewItem = 302         '新项目(&A)
Public Const conMenu_Edit_Modify = 303          '修改(&M)
Public Const conMenu_Edit_Delete = 304          '删除(&D)
Public Const conMenu_Edit_FinOut = 305             '完成交班(&U)
Public Const conMenu_Edit_FinIn = 306              '完成接班(&B)
Public Const conMenu_Edit_FinRead = 307          '完成审阅(&P)
'Public Const conMenu_Edit_Out = 308            '交班(&U)
'Public Const conMenu_Edit_In = 309              '接班(&B)
'Public Const conMenu_Edit_Read = 310          '审阅(&P)
Public Const conMenu_Edit_CancelOut = 311            '取消完成交班(&U)
Public Const conMenu_Edit_CancelIn = 312             '取消完成接班(&B)
Public Const conMenu_Edit_CancelRead = 313          '取消完成审阅(&P)
Public Const conMenu_Edit_CheckOutSign = 314    '验证交班电子签名(&C)
Public Const conMenu_Edit_CheckInSign = 315    '验证接班电子签名(&C)
'报表菜单
Public Const conMenu_Report_Record = 601   '交接班情况查询


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

'其它常量定义
'********************************************************************
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
'********************************************************************
Public gstrProductName As String            'OEM产品名称
Public gstrSQL As String
Public gcnOracle As New ADODB.Connection
Public gobjRegister As Object               '注册授权部件zlRegister
Public grsUserInfo As ADODB.Recordset
Public glngSys As Long
Public gstrDbaUser As String
'1,'新入',2,'抢救',3,'一级护理',4,'术后',5,'术前',6,'死亡',7,'输血',8,'危',9,'其他',10,'危/重',11,'特检',12,'留观'

Public gstrSysName As String
Public gstrPrivs As String
Public glngModul As Long

Public gobjEmr As Object

