Attribute VB_Name = "mdlDefine"
Option Explicit

'################################################################################################################
'##     菜单
'################################################################################################################
'主菜单
Public Const ID_File_Menu = 101 '文件
Public Const ID_Edit_Menu = 102 '编辑
Public Const ID_Insert_Menu = 103 '插入
Public Const ID_Com_Bar = 104   '常用
Public Const ID_Sign_Bar = 105  '签名
Public Const ID_Format_Bar = 106 '格式
Public Const ID_Table_Bar = 107 '表格制作

'文件 "File"
Public Const ID_FILE_CLEAR = 300                '清空
Public Const ID_FILE_IMPORT = 301               '引入
Public Const ID_FILE_CLOSE = 302                '关闭 ×
Public Const ID_FILE_SAVE = 303                 '保存
Public Const ID_FILE_SAVEAS = 304               '另存为
Public Const ID_FILE_PAGESETUP = 305            '页面设置
Public Const ID_FILE_PRINTPREVIEW = 306         '打印预览
Public Const ID_FILE_PRINT = 307                '打印
Public Const ID_FILE_EXIT = 308                 '退出
Public Const ID_FILE_SAVEASEPRDEMO = 309        '另存为范文
Public Const ID_FILE_EXPORTTOXML = 310          '导出为XML文件
Public Const ID_FILE_IMPORTFROMXML = 311        '从XML文件导入
Public Const ID_FILE_EXPORTTOHTML = 312         '导出为HTML文件
Public Const ID_FILE_PRINTINWORD = 313          '在Word中打印
Public Const ID_FILE_SAVEASSEGMENT = 314        '另存为片段
Public Const ID_FILE_SAVE_QUIT = 315            '保存并退出

'编辑 "Edit"
Public Const ID_EDIT_UNDO = 320                 '撤销
Public Const ID_EDIT_REDO = 321                 '重做
Public Const ID_EDIT_CUT = 322                  '剪切
Public Const ID_EDIT_COPY = 323                 '复制
Public Const ID_EDIT_PASTE = 324                '粘贴
Public Const ID_EDIT_DELETE = 325               '删除
Public Const ID_EDIT_SELECTALL = 326            '全选
Public Const ID_EDIT_FIND = 327                 '查找
Public Const ID_EDIT_REPLACE = 328              '替换
Public Const ID_EDIT_FINDNEXT = 329             '查找下一个
Public Const ID_EDIT_FORMATBRUSH = 330          '格式刷
Public Const ID_EDIT_ADDCOMPEND = 331           '新增提纲
Public Const ID_EDIT_MODCOMPEND = 332           '修改提纲
Public Const ID_EDIT_DELCOMPEND = 333           '删除提纲
Public Const ID_EDIT_REFCOMPEND = 334           '刷新提纲
Public Const ID_EDIT_SAVEASPHRASE = 335         '存为词句示范
Public Const ID_EDIT_COMPENDWORD = 336          '提纲词句对照

Public Const ID_EDIT_MARKEDPIC = 337            '标记修改
Public Const ID_EDIT_OUTERPIC = 338             '底图处理
Public Const ID_EDIT_DELETEELEMENT = 339        '删除要素

'视图 "View"
Public Const ID_VIEW_STRUCTURE = 340            '文档结构图
Public Const ID_VIEW_PHRASEDEMO = 341           '词句示范列表
Public Const ID_VIEW_SEGMENT = 342              '示范片段列表
Public Const ID_VIEW_HEADFOOT = 343             '页眉页脚
Public Const ID_VIEW_GRID = 344                 '网格线
Public Const ID_VIEW_PACSPIC = 345              'PACS图片组列表窗口
Public Const ID_VIEW_MULTIDOCVIEW = 346         '多文档查阅
Public Const ID_VIEW_CHARCOUNT = 347            '字数统计
Public Const ID_VIEW_RULER = 348                '标尺
Public Const ID_VIEW_PENWINDOW = 349            '手写输入窗口
Public Const ID_VIEW_HISTORYWINDOW = 3400       '共享页面内容

'插入 "Insert"
Public Const ID_INSERT_DATETIME = 350           '日期时间
Public Const ID_INSERT_SPECIALCHAR = 351        '特殊符号
Public Const ID_INSERT_PICTURE = 352            '图片
Public Const ID_INSERT_TABLE = 353              '表格
Public Const ID_INSERT_ELEMENT = 354            '诊治要素
Public Const ID_INSERT_EPRDEMO = 355            '全文示范
Public Const ID_INSERT_DATE = 356               '插入日期
Public Const ID_INSERT_TIME = 357               '插入时间
Public Const ID_INSERT_DOCADVISE = 358          '插入本次就诊医嘱
Public Const ID_INSERT_AUTORECOGNISE = 359      '智能识别（诊治要素、字典项目）
Public Const ID_INSERT_PRECOMPEND = 360         '插入预制提纲
Public Const ID_INSERT_PACSPIC = 361            '插入PACS图片组

'格式 "Format"
Public Const ID_FORMAT_FONT = 390               '字体
Public Const ID_FORMAT_BACKGROUND = 391         '背景色
Public Const ID_FORMAT_PROTECT = 392            '保护
Public Const ID_FORMAT_BOLD = 393               '粗体
Public Const ID_FORMAT_ITALIC = 394             '斜体
Public Const ID_FORMAT_SUPER = 395              '上标
Public Const ID_FORMAT_SUB = 396                '下标
Public Const ID_FORMAT_UNDERLINE_THIN = 397     '下划线：细线
Public Const ID_FORMAT_UNDERLINE_THICK = 398    '下划线：粗线
Public Const ID_FORMAT_UNDERLINE_WAVE = 399     '下划线：波浪线
Public Const ID_FORMAT_UNDERLINE_DOT = 400      '下划线：点线
Public Const ID_FORMAT_UNDERLINE_DASH = 401     '下划线：虚线
Public Const ID_FORMAT_UNDERLINE_DASHDOT = 402  '下划线：点划线
Public Const ID_FORMAT_UNDERLINE_DASHDOT2 = 403 '下划线：双点划线
Public Const ID_FORMAT_ALIGNLEFT = 404          '对齐方式：左对齐
Public Const ID_FORMAT_ALIGNCENTER = 405        '对齐方式：左对齐
Public Const ID_FORMAT_ALIGNRIGHT = 406         '对齐方式：左对齐
Public Const ID_FORMAT_LISTNONE = 407           '项目符号：无
Public Const ID_FORMAT_LISTBULLETS = 408        '项目符号：项目符号
Public Const ID_FORMAT_LISTLCHAR = 409          '项目符号：小写字母
Public Const ID_FORMAT_LISTUCHAR = 410          '项目符号：大写字母
Public Const ID_FORMAT_LISTLROME = 411          '项目符号：小写罗马数字
Public Const ID_FORMAT_LISTUROME = 412          '项目符号：大写罗马数字
Public Const ID_FORMAT_LINESPACE = 413          '行间距
Public Const ID_FORMAT_SPACEBEFORE = 414        '段前距离
Public Const ID_FORMAT_SPACEAFTER = 415         '段后距离
Public Const ID_FORMAT_FIRSTINDENT = 416        '首行缩进
Public Const ID_FORMAT_FIRSTHUNGING = 417       '首行悬挂
Public Const ID_FORMAT_INDENTDECREASE = 418     '减少缩进量
Public Const ID_FORMAT_INDENTINCREASE = 419     '增加缩进量
Public Const ID_FORMAT_UNDERLINE = 420          '下划线
Public Const ID_FORMAT_LISTARABIC = 421         '项目符号：阿拉伯数字
Public Const ID_FORMAT_PARA = 422               '段落属性
Public Const ID_FORMAT_LINESPACE1 = 423         '行间距：1.0倍
Public Const ID_FORMAT_LINESPACE2 = 424         '行间距：1.3倍
Public Const ID_FORMAT_LINESPACE3 = 425         '行间距：1.5倍
Public Const ID_FORMAT_LINESPACE4 = 426         '行间距：2.0倍
Public Const ID_FORMAT_LINESPACE5 = 427         '行间距：2.5倍
Public Const ID_FORMAT_LINESPACE6 = 428         '行间距：3.0倍
Public Const ID_FORMAT_LINESPACE7 = 429         '行间距：其他...
Public Const ID_FORMAT_HIGHLIGHT = 530          '高亮显示 ×
Public Const ID_FORMAT_FORECOLOR = 531          '字体颜色
Public Const ID_FORMAT_STYLE = 532              '字体样式
Public Const ID_FORMAT_FONTNAME = 533           '字体名称
Public Const ID_FORMAT_FONTSIZE = 534           '字体尺寸
Public Const ID_FORMAT_UNDERLINE_NONE = 535     '下划线：无
Public Const ID_FORMAT_LISTSETUP = 536          '项目符号设置
Public Const ID_FORMAT_STYLEWINDOW = 537        '样式窗格

'表格 "Table"
Public Const ID_TABLE_INSERTTABLE = 430         '插入表格 ×
Public Const ID_TABLE_INSERTCOLLEFT = 431       '插入列（左边）
Public Const ID_TABLE_INSERTCOLRIGHT = 432      '插入列（右边）
Public Const ID_TABLE_INSERTROWUP = 433         '插入行（靠上）
Public Const ID_TABLE_INSERTROWDOWN = 434       '插入行（靠下）
Public Const ID_TABLE_INSERTCELL = 435          '插入单元格...
Public Const ID_TABLE_DELETETABLE = 436         '删除表格 ×
Public Const ID_TABLE_DELETECOL = 437           '删除列
Public Const ID_TABLE_DELETEROW = 438           '删除行
Public Const ID_TABLE_DELETECELL = 439          '删除单元格
Public Const ID_TABLE_FORMATCELL = 440          '单元格格式
Public Const ID_TABLE_FORMATROWHEIGHT = 441     '行高
Public Const ID_TABLE_FORMATCOLWIDTH = 442      '列宽
Public Const ID_TABLE_INSERTPICTURE = 443       '插入图片
Public Const ID_TABLE_BEELEMENTS = 444          '关联诊治要素
Public Const ID_TABLE_MERGE = 445               '合并单元格
Public Const ID_TABLE_CELLALIGNMENT1 = 446      '单元格对齐方式
Public Const ID_TABLE_CELLALIGNMENT2 = 447      '单元格对齐方式
Public Const ID_TABLE_CELLALIGNMENT3 = 448      '单元格对齐方式
Public Const ID_TABLE_CELLALIGNMENT4 = 449      '单元格对齐方式
Public Const ID_TABLE_CELLALIGNMENT5 = 450      '单元格对齐方式
Public Const ID_TABLE_CELLALIGNMENT6 = 451      '单元格对齐方式
Public Const ID_TABLE_CELLALIGNMENT7 = 452      '单元格对齐方式
Public Const ID_TABLE_CELLALIGNMENT8 = 453      '单元格对齐方式
Public Const ID_TABLE_CELLALIGNMENT9 = 454      '单元格对齐方式
Public Const ID_TABLE_BORDERSTYLE1 = 455        '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE2 = 456        '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE3 = 457        '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE4 = 458        '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE5 = 459        '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE6 = 460        '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE7 = 461        '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE8 = 462        '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE9 = 463        '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE10 = 464       '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE11 = 465       '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE12 = 466       '单元格边框样式
Public Const ID_TABLE_BORDERSTYLE13 = 467            '单元格边框样式
Public Const ID_TABLE_INSERTINHERITROW = 468    '插入继承行
Public Const ID_TABLE_INSERTINHERITCOL = 469    '插入继承列
Public Const ID_TABLE_CELLPROTECTED = 470       '保护单元格

'帮助 "Help"
Public Const ID_HELP_CONTENT = 500              '帮助主题
Public Const ID_HELP_ASSISTANT = 501            '系统助手 ×
Public Const ID_HELP_CONTACT = 502              '发送反馈
Public Const ID_HELP_ONLINE = 503               '在线医业
Public Const ID_HELP_ABOUT = 504                '关于...
Public Const ID_HELP_WEBFORUM = 505             '中联论坛(&F)

'#########################################################################
'##     工具栏新增的ID
'#########################################################################

'绘图工具栏 "Draw"
Public Const ID_DRAW_SELECT = 550               '选择
Public Const ID_DRAW_MOVE = 551                 '移动
Public Const ID_DRAW_LINE = 552                 '直线
Public Const ID_DRAW_MLINE = 553                '折线
Public Const ID_DRAW_RECT = 554                 '矩形
Public Const ID_DRAW_MRECT = 555                '多边形
Public Const ID_DRAW_CIRCLE = 556               '椭圆
Public Const ID_DRAW_TEXT = 557                 '文本
Public Const ID_DRAW_DELETE = 558               '删除
Public Const ID_DRAW_UNDO = 559                 '取消
Public Const ID_DRAW_REDO = 560                 '重做
Public Const ID_DRAW_RESET = 561                '清空
Public Const ID_DRAW_FILLCOLOR = 562            '填充色
Public Const ID_DRAW_LINECOLOR = 563            '线条色
Public Const ID_DRAW_FONTCOLOR = 564            '字体色
Public Const ID_DRAW_FILLSTYLE = 565            '填充样式
Public Const ID_DRAW_LINESTYLE = 566            '线条样式
Public Const ID_DRAW_LINEWIDTH = 567            '线条宽度
Public Const ID_DRAW_FILLNONE = 630             '填充方式
Public Const ID_DRAW_FILLALL = 631
Public Const ID_DRAW_FILLH = 632
Public Const ID_DRAW_FILLV = 633
Public Const ID_DRAW_FILLHV = 634
Public Const ID_DRAW_FILLR = 635
Public Const ID_DRAW_FILLL = 636
Public Const ID_DRAW_FILLLR = 637
Public Const ID_DRAW_LINECONTINUE = 639         '线条样式
Public Const ID_DRAW_LINEDOT = 640
Public Const ID_DRAW_LINEDASH = 641
Public Const ID_DRAW_LINEDASHDOT = 642
Public Const ID_DRAW_LINEDASHDOT2 = 643
Public Const ID_DRAW_LINEWIDTH1 = 644           '线条宽度
Public Const ID_DRAW_LINEWIDTH2 = 645
Public Const ID_DRAW_LINEWIDTH3 = 646
Public Const ID_DRAW_LINEWIDTH4 = 647
Public Const ID_DRAW_LINEWIDTH5 = 648
Public Const ID_DRAW_SEQUENCENUMBER = 650       '顺序编号
Public Const ID_DRAW_CLEARNUMBERS = 651         '清除顺序编号

'表格工具栏
Public Const ID_TABLE_MERGEANDCENTER = 580      '合并并居中
Public Const ID_TABLE_SAMECOLWIDTH = 581        '相同列宽
Public Const ID_TABLE_SAMEROWHEIGHT = 582       '相同行高
Public Const ID_TABLE_CURRENCY = 583            '货币
Public Const ID_TABLE_PERCENT = 584             '百分比
Public Const ID_TABLE_KILOBIT = 585             '千分位
Public Const ID_TABLE_DIGITSINCREASE = 586      '增加小数点
Public Const ID_TABLE_DIGITSDECREASE = 587      '减少小数点
Public Const ID_TABLE_BORDERSTYLE = 588         '边框样式
Public Const ID_TABLE_CELLALIGNMENT = 589       '单元格对齐方式
Public Const ID_TABLE_FORMULA = 590             '公式栏
Public Const ID_TABLE_INSERTTABLE_BAR = 591     '插入表格
Public Const ID_TABLE_PROPERTY = 592            '表格属性


'#########################################################################
'##     其他新增的菜单
'#########################################################################
Public Const ID_SIGN = 710                      '签名
Public Const ID_UNTREAD = 711                   '回退
Public Const ID_SIGN_QUIT = 712                 '签名并退出编辑窗体
Public Const ID_REVISION_PREV = 715             '前一处修订
Public Const ID_REVISION_NEXT = 716             '后一处修订
Public Const ID_REVISION_RESET = 717            '取消所选修订
Public Const ID_DIAGNOSIS = 720                 '诊断
Public Const ID_ELEMENT_TOSTRING = 722          '转换为纯文本
Public Const ID_EDIT_BACKSPACE = 723            '编辑器中按BackSpace键
Public Const ID_ELEMENT_CLEAR = 724             '清空文本
Public Const ID_ELEMENT_UPDATE = 725            '更新文本
Public Const ID_DesignTest = 9999               '设计环境下测试按扭

'PACS报告
Public Const ID_PACS_DeleteMarkedPic = 900      '删除标记图
Public Const ID_PACS_DeletePacsImg = 901        '删除PACS报告图片
Public Const ID_PACS_Layout = 903               '布局调整
Public Const ID_PACS_Left = 904                 '标记图在左边
Public Const ID_PACS_Right = 905                '标记图在右边
Public Const ID_PACS_None = 906                 '无标记图
