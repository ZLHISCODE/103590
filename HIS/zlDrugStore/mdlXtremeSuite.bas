Attribute VB_Name = "mdlXtremeSuite"
Option Explicit


'''''''''''''''''''Xtreme控件相关定义
''方块区域
'部门发药
Public Const mconPane_Dept_Condition = 1                    '条件栏

'PIVA管理
Public Const mconPane_PIVA_Condition = 1                    '条件栏

'处方发药
Public Const mconPane_Recipe_Condition = 1    '条件栏
Public Const mconPane_Recipe_List = 2         '处方列表

''TabControl分页
'PIVA管理
Public Const mconTab_PIVA_Check = 0                 '待核查
Public Const mconTab_PIVA_Dosage = 1                '待配药
Public Const mconTab_PIVA_Send = 2                  '待发送
Public Const mconTab_PIVA_Return = 3                '已发送

'部门发药
Public Const mconTab_Dept_Send = 0                  '未发药品清单
Public Const mconTab_Dept_SumSend = 1               '汇总清单
Public Const mconTab_Dept_Shortage = 2              '缺药清单
Public Const mconTab_Dept_Reject = 3                '拒发药清单
Public Const mconTab_Dept_Return = 4                '已发药清单

'处方发药
Public Const mconTab_Recipe_DosageOk = 0          '配药确认
Public Const mconTab_Recipe_Dosage = 1            '配药
Public Const mconTab_Recipe_Abolish = 2           '取消配药
Public Const mconTab_Recipe_Send = 3              '待发药
Public Const mconTab_Recipe_OverTime = 4          '超期未发
Public Const mconTab_Recipe_Return = 5            '退药

''主菜单
Public Const mconMenu_FilePopup = 1 '文件
Public Const mconMenu_ManagePopup = 2 '管理
Public Const mconMenu_EditPopup = 3 '编辑
Public Const mconMenu_ReportPopup = 4 '报表
Public Const mconMenu_PlanPopup = 5 '排班设置
Public Const mconMenu_ViewPopup = 7 '查看
Public Const mconMenu_ToolPopup = 8 '工具
Public Const mconMenu_HelpPopup = 9 '帮助

''文件菜单
Public Const mconMenu_File_Open = 100               '*打开(&O)…
Public Const mconMenu_File_PrintSet = 101           '*打印设置(&S)…
Public Const mconMenu_File_Preview = 102            '*预览(&V)
Public Const mconMenu_File_Print = 103              '*打印(&P)
Public Const mconMenu_File_Excel = 104              '输出到&Excel…

Public Const mconMenu_File_Parameter = 181          '*参数设置(&M)

Public Const mconMenu_File_Exit = 191               '*退出(&X)
Public Const mconMenu_File_Message = 10000            '消息菜单

'PIVA
Public Const mconMenu_File_PIVA_BillPrint = 151              '单据打印
Public Const mconMenu_File_PIVA_BillPrintWait = 152          '打印药品摆药单
Public Const mconMenu_File_PIVA_BillPrintLable = 153         '打印标签
Public Const mconMenu_File_PIVA_BillPrintTotal = 154         '打印发药清单
Public Const mconMenu_File_PIVA_BillPrintReturn = 155        '打印退药（销帐）清单
Public Const mconMenu_File_PIVA_BillPrintNext = 156          '续打瓶签
Public Const mconMenu_File_PIVA_BillPrintSum = 157         '打印汇总报表

'部门发药
Public Const mconMenu_File_Dept_BillPrint = 151              '单据打印
Public Const mconMenu_File_Dept_BillPrintTotal = 152         '打印汇总清单
Public Const mconMenu_File_Dept_BillPrintRestore = 153       '打印退药通知单
Public Const mconMenu_File_Dept_BillPrintWait = 154          '打印药品摆药单

'处方发药
Public Const mconMenu_File_Recipe_BillPrintDosage = 151            '打印配药单(&B)-F6
Public Const mconMenu_File_Recipe_BillPrintRecipe = 152            '打印处方签(&D)-F4
Public Const mconMenu_File_Recipe_BillPrintReport = 153            '打印发药清单(&W)
Public Const mconMenu_File_Recipe_BillPrintReturn = 154            '打印退药通知单(&R)
Public Const mconMenu_File_Recipe_BillPrintLable = 155             '打印药品标签(&L)-F11
Public Const mconMenu_File_Recipe_BillPrintBack = 156              '打印退费单据(T)
Public Const mconMenu_File_Recipe_BillPrintChange = 157            '打印医嘱更改通知单


''编辑菜单
'PIVA
Public Const mconMenu_Edit_PIVA_Check = 3301                '核查（审方）
Public Const mconMenu_Edit_PIVA_Prepare = 3302              '摆药
Public Const mconMenu_Edit_PIVA_AutoSetBatch = 3303         '自动分配批次
Public Const mconMenu_Edit_PIVA_Dosage = 3304               '配药
Public Const mconMenu_Edit_PIVA_CancelDosage = 3305         '取消配药
Public Const mconMenu_Edit_PIVA_Send = 3306                 '发送
Public Const mconMenu_Edit_PIVA_CancelSend = 3307           '取消发送
Public Const mconMenu_Edit_PIVA_Cancel = 3308               '取消
Public Const mconMenu_Edit_PIVA_PASS = 3309                 'PASS
Public Const mconMenu_Edit_PIVA_Delete = 3310               '删除作废
Public Const mconMenu_Edit_PIVA_ReVerify = 3311             '销帐审核
Public Const mconMenu_Edit_PIVA_Approve = 3312                '审核医嘱
Public Const mconMenu_Edit_PIVA_CancelApprove = 3313         '取消审核
Public Const mconMenu_Edit_PIVA_Lock = 3314                 '全部锁定
Public Const mconMenu_Edit_PIVA_UnLock = 3315               '全部解锁
Public Const mconMenu_Edit_PIVA_Beach = 3316                '调整批次
Public Const MCONMENU_EDIT_PIVA_REFUSE = 3317               '确认拒绝
Public Const MCONMENU_EDIT_PIVA_SURE = 3318                 '确认调整
Public Const MCONMENU_EDIT_PIVA_SORTSET = 3319              '排序设置
Public Const MCONMENU_EDIT_PIVA_PLAN = 3320                 '排班安排
Public Const MCONMENU_EDIT_PIVA_MedicalRecord = 3331        '电子病案查阅



'排班菜单
'PIVA
Public Const MCONMENU_PLAN_PIVA_DESK = 3501
Public Const MCONMENU_PLAN_PIVA_DESKDRUG = 3502
Public Const MCONMENU_PLAN_PIVA_PERWORK = 3503

'外挂发药业务
Public Const mconMenu_Edit_PlugIn = 3400          '扩展

'部门发药
Public Const mconMenu_Edit_Dept_Verify = 3101             '发药
Public Const mconMenu_Edit_Dept_Desire = 3102             '缺药申领
Public Const mconMenu_Edit_Dept_Reject = 3103             '拒发确认
Public Const mconMenu_Edit_Dept_Return = 3104             '退药
Public Const mconMenu_Edit_Dept_ReturnOther = 3105        '退其它药房的处方
Public Const mconMenu_Edit_Dept_ReVerify = 3106           '药品退药销账
Public Const mconMenu_Edit_Dept_StopFlag = 3107           '停止发药标记
Public Const mconMenu_Edit_Dept_RejectRestore = 3108      '拒发恢复
Public Const mconMenu_Edit_Dept_EMR = 3109                '病案查询
Public Const mconMenu_Edit_Dept_Packer = 3110             '分包机
Public Const mconMenu_Edit_Dept_VerifySign = 3112       '验证签名
Public Const mconMenu_Edit_Dept_Hot_IC = 3150             'IC卡按钮热键
Public Const mconMenu_Edit_Dept_CustomCheck = 3151             '自定义审核功能
Public Const mconMenu_Edit_Dept_MedicalRecord = 3161           '电子病案查阅

'处方发药
Public Const mconMenu_Edit_Recipe_Dosage = 3201                 '配药模式(&D)-^D
Public Const mconMenu_Edit_Recipe_Abolish = 3202                '取消模式(&A)-^A
Public Const mconMenu_Edit_Recipe_Send = 3203                   '发药模式(&C)-^C
Public Const mconMenu_Edit_Recipe_Return = 3204                 '退药模式(&H)-^H
Public Const mconMenu_Edit_Recipe_Batch = 3205                  '批量发药(&B)
Public Const mconMenu_Edit_Recipe_SendOther = 3206              '发其它药房的处方(&F)
Public Const mconMenu_Edit_Recipe_ReturnBatch = 3207            '退其它药房的处方(&T)
Public Const mconMenu_Edit_Recipe_SendByBill = 3208             '按票据号发药(&I)
Public Const mconMenu_Edit_Recipe_ReturnByBill = 3209           '按票据号退药(&R)
Public Const mconMenu_Edit_Recipe_Flag = 3210                   '停止发药标记(&S)
Public Const mconMenu_Edit_Recipe_Cancel = 3211                 '取消发药(&Q)-^Q
Public Const mconMenu_Edit_Recipe_Charge = 3212                 '门诊划价(&M)-F8
Public Const mconMenu_Edit_Recipe_Stuff = 3213                  '卫材发料(@W)-F9
Public Const mconMenu_Edit_Recipe_Change = 3214                 '切换配药人(&E)
Public Const mconMenu_Edit_Recipe_EMR = 3215                     '病案查询
Public Const mconMenu_Edit_Recipe_SendHot = 3216                '发药操作，用于快捷键调用
Public Const mconMenu_Edit_Recipe_AddSign = 3217                '补充签名
Public Const mconMenu_Edit_Recipe_Windows = 3218               '调整发药窗口(&N)
Public Const mconMenu_Edit_Recipe_Call = 3219                  '呼叫(&G)
Public Const mconMenu_Edit_Recipe_VerifySign = 3220               '验证签名
Public Const mconMenu_Edit_Recipe_Cancle = 3221                  '取消确认(&G)
Public Const mconMenu_Edit_Recipe_TakeDrug = 3222                '取药确认
Public Const mconMenu_Edit_Recipe_Hot_IC = 3250                 'IC卡按钮热键
Public Const mconMenu_Edit_Recipe_AutoSend = 3222                  '门诊自动发药设置
Public Const mconMenu_Edit_Recipe_AutoSend_Open = 32221            '启用处方上传
Public Const mconMenu_Edit_Recipe_AutoSend_Set = 32222            '设置webService路径
Public Const mconMenu_Edit_Recipe_AutoSend_LoadDrug = 32223           '上传药品基础数据
Public Const mconMenu_Edit_Recipe_AutoSend_LoadStock = 32224          '上传药品库存数据
Public Const mconMenu_Edit_Recipe_MedicalRecord = 3232          '电子病案查阅

'药房自动发药设置
Public Const mconMenu_AutoSend = 100000

'调整发药窗口
Public Const mconMenu_Edit_Recipe_Guide = 4001
Public Const mconMenu_Edit_Recipe_OK = 4002
Public Const mconMenu_Edit_Recipe_Average = 4003
'刷新791
'退出191


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

''弹出菜单（录入类型）
Public Const mconMenu_InputPopup = 1000                  '录入类型

'部门发药
Public Const mconMenu_Input_Dept_HosNumber = 1101             '住院号
Public Const mconMenu_Input_Dept_Name = 1102                  '姓名
Public Const mconMenu_Input_Dept_BedNumber = 1103             '床号
Public Const mconMenu_Input_Dept_NO = 1104                    '单据号
Public Const mconMenu_Input_Dept_Ident = 1105                 '病人ID
Public Const mconMenu_Input_Dept_ReceiveNO = 1106             '领药号
Public Const mconMenu_Input_Dept_BatchSendNO = 1107           '汇总发药号
Public Const mconMenu_Input_Dept_Dept = 1108                  '领药部门
Public Const mconMenu_Input_Dept_ICCard = 1109                'IC卡

'处方发药
Public Const mconMenu_Input_Recipe_NO = 1201                    '单据号(&1)
Public Const mconMenu_Input_Recipe_OPNO = 1202                  '门诊号(&2)
Public Const mconMenu_Input_Recipe_Name = 1203                  '姓名(&3)
Public Const mconMenu_Input_Recipe_IDCard = 1204                '身份证(&4)
Public Const mconMenu_Input_Recipe_ICCard = 1205                'IC卡(&5)
Public Const mconMenu_Input_Recipe_MINo = 1206                  '医保号(&6)
Public Const mconMenu_Input_Recipe_HosNumber = 1207             '住院号(&7)

''弹出菜单（部门列表内容）
Public Const mconMenu_ListPopup = 2000
Public Const mconMenu_List_OnlyShowDept = 2001              '仅显示部门列表
Public Const mconMenu_List_ShowOther = 2002                 '显示详细内容
Public Const mconMenu_List_ShowAll = 2010                   '显示所有科室
Public Const mconMenu_List_ShowClin = 2011                  '显示临床科室
Public Const mconMenu_List_ShowTech = 2012                  '显示医技科室
Public Const mconMenu_List_ShowArea = 2013                  '显示病区
Public Const mconMenu_List_ShowReject = 2014                '提取拒发药品
Public Const mconMenu_List_Sort = 2015                      '部门按发送时间排序


''弹出菜单（PASS）
Public Const mconMenu_PASS = 5000
Public Const mconMenu_PASS_Item = 5100
Public Const mconMenu_PASS_Spec = 5200

'弹出菜单（PIVA）查找
Public Const mconMenu_Look = 2000
Public Const mconMenu_Filter = 2100


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
