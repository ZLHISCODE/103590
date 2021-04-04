Attribute VB_Name = "mdlPubDefine"
Option Explicit
'全局变量，全局参数
'----------------------------------------------------------------------------------
Public Const ConMenu_Appfro_AddBill = 101           '增加分类
Public Const ConMenu_Appfro_DelBill = 102           '删除分类
Public Const ConMenu_Appfro_ModifyItem = 103        '修改项目
Public Const ConMenu_Appfro_Exit = 104              '退出
Public Const ConMenu_Appfro_DeptSel = 105           '执行科室选择
Public Const ConMenu_Appfro_Refresh = 106           '刷新
Public Const ConMenu_Appfro_ModifyBill = 107        '修改分类
Public Const ConMenu_Appfro_ModifyDept = 108        '修改执行科室
Public Const ConMenu_Appfor_ItemSort = 401          '调整顺序
Public Const ConMenu_Appfor_ClincHelp = 402         '调整顺序

Public Const ConMenu_Browse_SelAll = 109            '全选
Public Const ConMenu_Browse_ClsAll = 110            '全清
Public Const ConMenu_Browse_Refresh = 111           '刷新
Public Const ConMenu_Browse_Print = 112             '打印
Public Const ConMenu_Browse_Exit = 113              '退出
Public Const ConMenu_Browse_Find = 114              '查找


Public Const ConMenu_Browse_Save = 115              '保存
Public Const ConMenu_Browse_Cancel = 116            '取消
Public Const ConMenu_Browse_PrintView = 117         '打印预览
Public Const ConMenu_Browse_PrintSet = 118          '打印设置
Public Const ConMenu_Appfro_Group = 119             '选择分组
Public Const conFun_Sample_Auditing = 120           '复核
Public Const conFun_Sample_unAuditing = 121         '取消复核
Public Const ConMenu_Browse_unPrint = 122           '重置打印
Public Const ConMenu_Browse_PrintAll = 123          '打印所有
Public Const ConMenu_Browse_PrintViewAll = 124      '预览所有
Public Const ConMenu_Browse_PrintSetAll = 125       '打印设置

'-------------------------------------------------------------------
Public Const ConTab_Sample_History = 201            '历次
Public Const ConTab_Sample_Image = 202              '图像
Public Const ConTab_Sample_Comment = 203            '备注
'--------------------------------------------------------------------

'---------------------------------------------------------------------
Public Const ConMenu_pop_In = 301             '住院号
Public Const ConMenu_pop_bed = 302            '床号
Public Const ConMenu_pop_Dept = 303           '申请科室
Public Const ConMenu_pop_DeptDistrict = 304   '申请病区
Public Const ConMenu_pop_Out = 305            '门诊号
Public Const ConMenu_pop_PatiCard = 306       '就诊卡
Public Const ConMenu_pop_SampleCode = 307     '条码号

'---------------------------------------------------------------------
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

Public Const VK_P = &H50

Public Const VsModiBackColor = &HD6FFCA        'vs控件，可编辑单元的背景色
'*********************************************************************

Public Const conMenu_Tool_PlugIn = 890          '插件
Public Const conMenu_Tool_PlugIn_Item = 89000   '插件项,实际依次为 conMenu_Tool_PlugIn_Item + n, 1<=n<=99

'菜单按钮
Public Const CONFUN_UP = 501                            '上一个
Public Const CONFUN_DOWN = 502                          '下一个
