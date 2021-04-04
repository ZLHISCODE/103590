Attribute VB_Name = "mdlPubDefine"
Option Explicit


'查看菜单
Public Const conMenu_View_ToolBar = 701              '工具栏(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Public Const conMenu_View_StatusBar = 702            '状态栏(&S)
Public Const conMenu_View_Append = 703               '附加信息(&A)
Public Const conMenu_View_Difference = 704              '显示差异(&D)
Public Const conMenu_View_Contrast = 705                 '对比查看


'采集菜单
Public Const conMenu_Cap_Group = 8090           '采集
Public Const conMenu_Cap_Dynamic = 8100         '动态显示(&V)
Public Const conMenu_Cap_MarkMap = 8101         '影像采集(&C)
Public Const conMenu_Cap_Import = 8102          '影像导入(&I)
Public Const conMenu_Cap_DevSet = 8103          '影像设备设置(&D)
Public Const comMenu_Cap_Process = 8104         '影像处理
Public Const conMenu_Cap_Record = 8105          '录像(&R)
Public Const conMenu_Cap_DelImg = 8097          '删除图像
Public Const conMenu_Cap_Full_Screen = 8098     '全屏(&U)
Public Const conMenu_Cap_Record_Stop = 8099     '停止录像(&O)
Public Const conMenu_Cap_Play = 8106            '播放(&P)
Public Const conMenu_Cap_Stop = 8107            '停止(&T)
Public Const conMenu_Cap_Forward = 8108         '快进(&F)
Public Const conMenu_Cap_Back = 8109            '快退(&B)
Public Const conMenu_Cap_SaveAs = 8126    '8110          '保存录像(&S)
Public Const conMenu_Cap_OpenStudyList = 8122   '打开检查列表
Public Const conMenu_Cap_StudySyncState = 8123  '影像检查同步状态
Public Const conMenu_Cap_RecordAudio = 8125     '录音
Public Const conMenu_Cap_After_Capture = 8140   '后台采集
Public Const conMenu_Cap_After_Record = 8141    '后台录像
Public Const conMenu_Cap_After_Tag = 8142       '更新标记




'图像处理
Public Const conMenu_Process_Window = 501           '亮度对比度
Public Const conMenu_Process_Zoom = 502             '缩放
Public Const conMenu_Process_Corp = 512             '拖动
Public Const conMenu_Process_RRotate = 503          '顺时针旋转
Public Const conMenu_Process_LRotate = 504          '逆时针旋转
Public Const conMenu_Process_Sharpness = 505        '锐化
Public Const conMenu_Process_Filter = 506           '平滑
Public Const conMenu_Process_Arrow = 507            '箭头标注
Public Const conMenu_Process_Ellipse = 508          '圆形标注
Public Const conMenu_Process_Text = 509             '文字标注
Public Const conMenu_Process_RectZoom = 510         '裁剪采集
Public Const conMenu_Process_RectCapture = 511         '裁剪后采集
Public Const conMenu_Process_Restore = 8124         '恢复



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
Public Const XTP_ID_WINDOW_LIST = 35000    '窗体列表
Public Const XTP_ID_TOOLBARLIST = 59392    '工具栏列表
Public Const ID_INDICATOR_CAPS = 59137    '状态栏（大写）
Public Const ID_INDICATOR_NUM = 59138    '状态栏（数字）
Public Const ID_INDICATOR_SCRL = 59139    '状态栏（滚动）

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


