Attribute VB_Name = "mDefinitions"
Option Explicit

'用于注册系统热键
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

'系统虚拟键定义
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

'文件
Public Const ID_FILE_OPEN = 300                 '打开
Public Const ID_FILE_SAVE = 301                 '保存
Public Const ID_FILE_SAVEAS = 302               '另存为
Public Const ID_FILE_PRINT = 303                '打印
Public Const ID_FILE_EXIT = 304                 '退出

'编辑
Public Const ID_EDIT_UNDO = 400                 '撤销
Public Const ID_EDIT_REDO = 401                 '重做
Public Const ID_EDIT_COPY = 402                 '复制
Public Const ID_EDIT_PASTE = 403                '粘贴
Public Const ID_EDIT_SIZE = 404                 '调整尺寸
Public Const ID_EDIT_ORIENT = 405               '调整方向
Public Const ID_EDIT_SCROLLMODE = 406           '卷动模式
Public Const ID_EDIT_CROPMODE = 407             '剪切模式

'缩放
Public Const ID_ZOOM_IN = 500                   '放大
Public Const ID_ZOOM_OUT = 501                  '缩小
Public Const ID_ZOOM_11 = 502                   '1:1
Public Const ID_ZOOM_FIT = 503                  '适合

'颜色
Public Const ID_COLOR_BLACKWHITE = 600          '灰度-黑白
Public Const ID_COLOR_GREYS16 = 601             '灰度-16色
Public Const ID_COLOR_GREYS256 = 602            '灰度-256色
Public Const ID_COLOR_COLOR2 = 603              '彩色-2色
Public Const ID_COLOR_COLOR16 = 604             '彩色-16色
Public Const ID_COLOR_COLOR256 = 605            '彩色-256色
Public Const ID_COLOR_TRUECOLOR = 606           '真彩色

'调节
Public Const ID_ADJUST_BRIGHT = 700             '亮度
Public Const ID_ADJUST_CONTRAST = 701           '对比度
Public Const ID_ADJUST_SITUATION = 702          '饱和度
Public Const ID_ADJUST_FILTERBROWSER = 703      '滤镜浏览器

'滤镜
Public Const ID_FILTER_COLOR1 = 800             '颜色－灰度
Public Const ID_FILTER_COLOR2 = 801             '颜色－负片效果
Public Const ID_FILTER_COLOR3 = 802             '颜色－老照片
Public Const ID_FILTER_COLOR4 = 803             '颜色－颜色填充
Public Const ID_FILTER_COLOR5 = 804             '颜色－替换 HS...
Public Const ID_FILTER_COLOR6 = 805             '颜色－替换 L...
Public Const ID_FILTER_COLOR7 = 806             '颜色－曝光过度

Public Const ID_FILTER_DEF1 = 810               '清晰度－模糊
Public Const ID_FILTER_DEF2 = 811               '清晰度－柔化
Public Const ID_FILTER_DEF3 = 812               '清晰度－锐化
Public Const ID_FILTER_DEF4 = 813               '清晰度－扩散
Public Const ID_FILTER_DEF5 = 814               '清晰度－象素化
Public Const ID_FILTER_DEF6 = 815               '清晰度－去斑
Public Const ID_FILTER_DEF7 = 816               '清晰度－进一步去斑

Public Const ID_FILTER_EDGES1 = 820             '边缘－轮廓
Public Const ID_FILTER_EDGES2 = 821             '边缘－浮雕
Public Const ID_FILTER_EDGES3 = 822             '边缘－版画
Public Const ID_FILTER_EDGES4 = 823             '边缘－醒目

Public Const ID_FILTER_SPECIAL1 = 830           '特殊－噪音
Public Const ID_FILTER_SPECIAL2 = 831           '特殊－扫描线
Public Const ID_FILTER_SPECIAL3 = 832           '特殊－扩张
Public Const ID_FILTER_SPECIAL4 = 833           '特殊－腐蚀
Public Const ID_FILTER_SPECIAL5 = 834           '特殊－纹理...

'视图
Public Const ID_VIEW_TOOLBARLIST = 59392        '工具栏列表
Public Const ID_VIEW_PANORAMIC = 900            '缩略图
Public Const ID_VIEW_PROPERTY = 901             '属性

'帮助 "Help"
Public Const ID_HELP_CONTENT = 902              '帮助主题
Public Const ID_HELP_CONTACT = 903              '发送反馈
Public Const ID_HELP_ONLINE = 904               '在线医业
Public Const ID_HELP_ABOUT = 905                '关于...

Public Const ID_PANE_PREVIEW = 10000
Public Const ID_PANE_INFO = 10001
Public Const ID_PANE_FILTER = 10002
Public Const ID_PANE_TEXTURE = 10003
