Attribute VB_Name = "mdlXtremeSuite"
Option Explicit


'''''''''''''''''''Xtreme控件相关定义
''主菜单
Public Const mconMenu_FilePopup = 1 '文件
Public Const mconMenu_ManagePopup = 2 '管理
Public Const mconMenu_EditPopup = 3 '编辑
Public Const mconMenu_ReportPopup = 4 '报表
Public Const mconMenu_ViewPopup = 7 '查看
Public Const mconMenu_ToolPopup = 8 '工具
Public Const mconMenu_HelpPopup = 9 '帮助

''文件菜单
Public Const mconMenu_File_Open = 100               '*打开(&O)…
Public Const mconMenu_File_PrintSet = 101           '*打印设置(&S)…
Public Const mconMenu_File_Preview = 102            '*预览(&V)
Public Const mconMenu_File_Print = 103              '*打印(&P)
Public Const mconMenu_File_Excel = 104              '输出到&Excel…
Public Const mconMenu_File_BillPrint = 105          '单据打印
Public Const mconMenu_File_BillPreview = 106        '单据预览
Public Const mconMenu_File_Parameter = 181       '*参数设置(&M)

Public Const mconMenu_File_Exit = 191            '*退出(&X)

''查看菜单
Public Const mconMenu_View_ToolBar = 701              '工具栏(&T)
Public Const mconMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const mconMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const mconMenu_View_ToolBar_Size = 7013           '大图标(&B)
Public Const mconMenu_View_StatusBar = 702            '状态栏(&S)
Public Const mconMenu_View_Append = 703               '附加信息(&A)
Public Const mconMenu_View_Expend = 711               '展开/折叠组(&X)
Public Const mconMenu_View_Expend_CurCollapse = 7111     '折叠当前组(&C)
Public Const mconMenu_View_Expend_CurExpend = 7112       '展开当前组(&E)
Public Const mconMenu_View_Expend_AllCollapse = 7113     '折叠所有组(&L)
Public Const mconMenu_View_Expend_AllExpend = 7114       '展开所有组(&X)
Public Const mconMenu_View_Find = 721                 '*查找(&F)
Public Const mconMenu_View_FindNext = 722             '继续查找(&N)
Public Const mconMenu_View_FindType = 723             '查找方式(&Y)
Public Const mconMenu_View_ReadIC = 724               '读IC卡(&I)
Public Const mconMenu_View_PatInfor = 725             '查看病人信息
Public Const mconMenu_View_PriceBill = 727
Public Const mconMenu_View_PriceTable = 728
Public Const mconMenu_View_PriceList = 729
Public Const mconMenu_View_FilterView = 730           '以过滤方式显示
Public Const mconMenu_View_Filter = 731               '*数据过滤(&I),子窗体的过滤功能
Public Const mconMenu_View_Notify = 732               '*医嘱提醒(&B)
Public Const mconMenu_View_Busy = 733                 '诊室忙(&M)
Public Const mconMenu_View_ShowAll = 734
Public Const mconMenu_View_ShowHistory = 735
Public Const mconMenu_View_ShowStoped = 736
Public Const mconMenu_View_Hide = 741                 '*隐藏(&H)
Public Const mconMenu_View_Show = 742                 '*显示(&S)
Public Const mconMenu_View_Forward = 743              '*前进(&F)
Public Const mconMenu_View_Backward = 744             '*后退(&B)
Public Const mconMenu_View_Dept = 745                '查看部门
Public Const mconMenu_View_Location = 746            '定位
Public Const mconMenu_View_LocationItem = 747        '定位项目
Public Const mconMenu_View_ColSet = 748              '列设置
Public Const mconMenu_View_Option = 781               '选项(&O)
Public Const mconMenu_View_Refresh = 791              '*刷新(&R)
Public Const mconMenu_View_Jump = 792                 '跳转(&J)

Public Const mconMenu_View_SelAll = 7301              '全选
Public Const mconMenu_View_ClsAll = 7302              '全清

Public Const mconMenu_View_Navigatebeginning = 7401           '*第一个(&F)
Public Const mconMenu_View_Navigateleft = 7402                '*上一个(&F)
Public Const mconMenu_View_Navigateright = 7403               '*下一个(&F)
Public Const mconMenu_View_Navigateend = 7404                 '*最后一个(&F)

Public Const mconMenu_View_FontSize = 4004         '字号设置
Public Const mconMenu_View_FontSize_1 = 4004         '9号字
Public Const mconMenu_View_FontSize_2 = 4004         '11号字
Public Const mconMenu_View_FontSize_3 = 4004         '15号字

''帮助菜单
Public Const mconMenu_Help_Help = 901           '*帮助主题(&H)
Public Const mconMenu_Help_Web = 902            '&WEB上的中联
Public Const mconMenu_Help_Web_Home = 9021      '中联主页(&H)
Public Const mconMenu_Help_Web_Forum = 9023     '中联论坛(&F)
Public Const mconMenu_Help_Web_Mail = 9022      '*发送反馈(&M)
Public Const mconMenu_Help_About = 991          '关于(&A)…

'盘点新增
Public Const mconMenu_Edit_AddBill = 3001        '增加记录单
Public Const mconMenu_Edit_AddTable = 3002       '增加盘点表
Public Const mconMenu_Edit_AddTableAuto = 30021  '自动产生盘点表
Public Const mconMenu_Edit_AddTableTotal = 30022 '汇总记录单产生盘点表
Public Const mconMenu_Edit_AddTableZero = 30023  '全部盘为零
Public Const mconMenu_Edit_AddTableHouseAll = 30024  '库房全部药品盘点
Public Const mconMenu_Edit_AddTableSpecial = 30025   '特殊药品盘点
Public Const mconMenu_Edit_AddModify = 3003      '修改
Public Const mconMenu_Edit_AddDel = 3004         '删除
Public Const mconMenu_Edit_AddVerify = 3005      '审核
Public Const mconMenu_Edit_AddStrike = 3006      '冲销
Public Const mconMenu_Edit_AddAffirmant = 3007   '阅读确认
Public Const mconMenu_Edit_AddDisplay = 3008     '查看单据
Public Const mconMenu_Edit_CheckTable = 5001     '盘点表智能检查

''CommandBar辅助热键
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

''CommandBar虚拟键
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
