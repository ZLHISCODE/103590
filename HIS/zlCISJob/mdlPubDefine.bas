Attribute VB_Name = "mdlPubDefine"
Option Explicit
'公共部份菜单ID定义:*表示有图标
'*********************************************************************
'如果程序由主管理程序和子业务程序组成，属主管理程序中的功能，应从"管理"菜单扩展，如果是子业务程序的功能，应从"编辑"菜单扩展。
Public Const conMenu_FilePopup = 1    '文件
Public Const conMenu_ManagePopup = 2    '管理
Public Const conMenu_EditPopup = 3    '编辑
Public Const conMenu_ReportPopup = 4    '报表
Public Const conMenu_PlugPopup = 6    '外接程序；检验技师工作站使用6100-6199
Public Const conMenu_ViewPopup = 7    '查看
Public Const conMenu_ToolPopup = 8    '工具
Public Const conMenu_HelpPopup = 9    '帮助
Public Const conMenu_PlugIn = 10        '插件扩展
Public Const conMenu_PlugIn_Menu = 1000001      '插件扩展


'文件菜单
Public Const conMenu_File_Open = 100            '*打开(&O)…
Public Const conMenu_File_PrintSet = 101        '*打印设置(&S)…
Public Const conMenu_File_Preview = 102         '*预览(&V)
Public Const conMenu_File_Preview_Pati = 10201         '*预览病人列表(&V)    '费用查询有效
Public Const conMenu_File_Print = 103           '*打印(&P)
Public Const conMenu_File_NoAskPrint = 106      '静默打印
Public Const conMenu_File_Print_Pati = 10301         '*打印病人列表(&P) '费用查询有效
Public Const conMenu_File_Print_PatiPath = 10302         '*打印病人版临床路径表
Public Const conMenu_File_Excel = 104           '输出到&Excel…
Public Const conMenu_File_Excel_Pati = 10401         '*病人列表输出到&Excel… '费用查询有效

Public Const conMenu_File_MedRec = 105          '首页打印(&R)
Public Const conMenu_File_MedRecSetup = 1051        '打印设置(&S)
Public Const conMenu_File_MedRecPreview = 1052      '打印预览(&P)
Public Const conMenu_File_MedRecPrint = 1053        '打印首页(&V)
Public Const conMenu_File_RowPrint = 121        '记录打印(&R)
Public Const conMenu_File_BatPrint = 122        '批量打印(&B)
Public Const conMenu_File_Print_Bespeak = 123        '预约挂号单打印
Public Const conMenu_File_Parameter = 181       '*参数设置(&M)
Public Const conMenu_File_RoomSet = 182         '执行间设备
Public Const ConMenu_File_ShortcutSet = 183     '快捷功能设置
Public Const conMenu_File_SendImg = 184         '发送图像
Public Const conMenu_File_SaveJpeg = 185       '保存为图片
Public Const conMenu_File_Exit = 191            '*退出(&X)
Public Const conMenu_File_ExportToXML = 192     '另存为XML文档
Public Const conMenu_File_ImportFromXML = 193   '从XML文档导入
Public Const conMenu_File_BillPrintView = 194   '单据打印预览
Public Const conMenu_File_BillPrint = 195       '单据打印
Public Const conMenu_File_ExportAll = 196       '批量导出
Public Const conMenu_File_ExportToXMLs = 197     '批量另存为XML文档
Public Const conMenu_File_ImportFromXMLs = 198   '从XML文档批量导入
Public Const conMenu_File_BarcodePrint = 199   '条码打印
Public Const conMenu_File_BillPrintSet = 200    '票据打印设置

'病人费用查询
Public Const conMenu_File_PrintMultiBill = 110    '打印多张催款单
Public Const conMenu_File_PrintSingleBill = 732    '打印单张催款单
Public Const conMenu_File_SchemeSet = 737    '报警方案设置
Public Const conMenu_File_PrintDayDetail = 3554    '打印一日清单
Public Const conMenu_File_PrintBedCard = 3555    '打印床头卡
Public Const conMenu_File_PrintPageSet = 113    '打印帐页设置

'管理菜单:工作站自身的功能菜单
Public Const conMenu_Manage_Monitor = 201     '*监护仪
Public Const conMenu_Manage_FeeItemSet = 3065      '诊疗项目费用设置

Public Const conMenu_Manage_Regist = 211      '*病人挂号(&H)
Public Const conMenu_Manage_Bespeak = 212     '预约挂号(&B),时间安排(&B)
Public Const conMenu_Manage_Transfer = 213    '*转诊处理(&C)
Public Const conMenu_Manage_Transfer_Send = 2131      '病人转诊(&S)
Public Const conMenu_Manage_Transfer_Cancel = 2132    '取消转诊(&C)
Public Const conMenu_Manage_Transfer_Incept = 2133    '转诊接收(&I)
Public Const conMenu_Manage_Transfer_Refuse = 2134    '转诊拒绝(&R)
Public Const conMenu_Manage_Transfer_Force = 2135     '强制续诊(&F)
Public Const conMenu_Manage_Receive = 214     '*病人接诊(&Z)
Public Const conMenu_Manage_Cancel = 215      '取消接诊(&Q)
Public Const conMenu_Manage_Finish = 216      '*完成接诊(&W)
Public Const conMenu_Manage_Redo = 217        '恢复接诊(&R)
Public Const conMenu_Manage_ReBack = 2171        '回诊
Public Const conMenu_Manage_ReBackCancel = 2172        '取消回诊

Public Const conMenu_Manage_Call = 218            '呼叫
Public Const conMenu_Manage_CallNext = 21801        '下一个(&N)
Public Const conMenu_Manage_CallPrevious = 21802    '上一个(&P)

Public Const conMenu_Manage_Reset = 219     '调整顺序
Public Const conMenu_Manage_Up = 21901        '上移(&U)
Public Const conMenu_Manage_Down = 21902      '下移(&D)
Public Const conMenu_Manage_Discard = 21903   '弃号(&D)
Public Const conMenu_Manage_Recall = 21904    '召回(&R)
Public Const conMenu_Manage_Untread = 21905   '退号(&R)
Public Const conMenu_Manage_TagEnd = 21906  '标记为结束(&M)
Public Const conMenu_Manage_ShowAller = 220  '显示所有过敏记录

Public Const conMenu_Manage_InQueue = 230     '排队叫号入队
Public Const conMenu_Manage_Plan = 221        '*执行报到(&P)
Public Const conMenu_Manage_Logout = 222      '取消报到(&L)
Public Const conMenu_Manage_Refuse = 223      '拒绝执行(&R)
Public Const conMenu_Manage_ReGet = 224       '取消拒绝(&G)
Public Const conMenu_Manage_Complete = 225    '*执行完成(&C)
Public Const conMenu_Manage_Undone = 226      '取消完成(&U)
Public Const conMenu_Manage_ThingAdd = 227    '*记录执行情况(&A)
Public Const conMenu_Manage_ThingModi = 228   '*调整执行情况(&M)
Public Const conMenu_Manage_ThingDel = 229    '*删除执行情况(&D)
'Public Const conMenu_Manage_ModifBaseInfo = 230 '基本信息调整
Public Const conMenu_Manage_ThingAudit = 234    '*核对执行情况(&E)
Public Const conMenu_Manage_ThingDelAudit = 235    '*取消核对执行(&F)
Public Const conMenu_Manage_ClearUp = 233     '检查报告驳回(&U)

Public Const conMenu_Manage_Request = 231        '*申请(&V)
Public Const conMenu_Manage_RequestView = 2311           '查阅申请(&V)
Public Const conMenu_Manage_RequestPrint = 2312           '打印诊疗单据(&J)
Public Const conMenu_Manage_RequestBatPrint = 2313           '批量打印条码(&B)
Public Const conMenu_Manage_Report = 232         '*报告(&O)
Public Const conMenu_Manage_ReportEdit = 2321        '填写报告(&E)
Public Const conMenu_Manage_ReportView = 2322        '查阅报告(&W)
Public Const conMenu_Manage_ReportPrint = 2323       '报告打印(&P)
Public Const conMenu_Manage_ReportPreview = 2324     '执行预览(&V)
Public Const conMenu_Manage_ReportLisView = 2325     '查阅LIS报告(&L)
Public Const conMenu_Manage_RecipeAuditView = 2326
Public Const conMenu_Manage_AppendBill = 2327  '老版医站补费
Public Const conMenu_Manage_EditCritical = 2328  '危急值登记
Public Const conMenu_Manage_QueryCriticl = 2329  '危急值查看
Public Const conMenu_Manage_LeaveMedi = 251    '寄存药品

Public Const conMenu_Manage_Audit = 252         '*审核申请
Public Const conMenu_Manage_UnAudit = 253       '*取消审核
Public Const conMenu_Manage_Arrange = 254       '*执行安排
Public Const conMenu_Manage_UnArrange = 255     '*取消安排

Public Const ConMenu_pop_Dept = 303           '申请科室
Public Const ConMenu_pop_DeptDistrict = 304   '申请病区

'护士站病人入出转
Public Const conMenu_Manage_Change_In = 2600          '病人入住
Public Const conMenu_Manage_Change_Turn = 2601      '转科
Public Const conMenu_Manage_Change_Bed = 2602         '换床
Public Const conMenu_Manage_Change_House = 2603       '包房
Public Const conMenu_Manage_Change_Out = 2604         '病人出院
Public Const conMenu_Manage_Change_InPati = 2605      '转为住院病人
Public Const conMenu_Manage_Change_BedGrid = 2606     '更改床位等级
Public Const conMenu_Manage_Change_PatiInfo = 2607    '调整住院信息
Public Const conMenu_Manage_Change_Baby = 2608        '新生儿登记
Public Const conMenu_Manage_Change_ReCalcFee = 2609   '按费别重算费用
Public Const conMenu_Manage_Change_InsureSel = 2610   '医保病种选择
Public Const conMenu_Manage_Change_Undo = 2611         '撤销功能
Public Const conMenu_Manage_Print_Label = 2612         '打印腕带
Public Const conMenu_Manage_Change_TurnUnit = 2613      '转病区
Public Const conMenu_Manage_Change_TurnTeam = 2614      '转医疗小组
Public Const conMenu_Manage_Change_TransposeBed = 2615         '床位对换
Public Const conMenu_Manage_Change_PaitNote = 2616         '病人备注信息
Public Const conMenu_Manage_Change_NurseGroup = 2617         '护理小组(整体护理)

'医嘱(编辑)菜单：因较多,共用时按4位编号,50位分段,001-050,051-100,101-150,...
Public Const conMenu_Edit_NewItem = 3001    '*新项目(&A)
Public Const conMenu_Edit_NewItemQAdvice = 300101    '*切换到医嘱
Public Const conMenu_Edit_NewItemQEpr = 300102    '*切换到病历
Public Const conMenu_Edit_Append = 3002     '*补充/补录(&Y)
Public Const conMenu_Edit_Modify = 3003     '*修改(&M)
Public Const conMenu_Edit_Delete = 3004     '*删除(&D)
Public Const conMenu_Edit_Blankoff = 3005   '*作废(&B)
Public Const conMenu_Edit_Stop = 3006       '*医嘱停止(&S)
Public Const conMenu_Edit_ReStop = 3007     '*确认停止(&C)
Public Const conMenu_Edit_Pause = 3008      '*暂停(&P)
Public Const conMenu_Edit_Reuse = 3009      '*启用(&U)
Public Const conMenu_Edit_Audit = 3010      '*审核/校对(&V)
Public Const conMenu_Edit_Price = 3011      '*计价调整(&I)
Public Const conMenu_Edit_ClearUp = 3012    '*医嘱重整(&F)
Public Const conMenu_Edit_Send = 3013       '*发送(&G)
Public Const conMenu_Edit_StopAudit = 3050      '停嘱审核(&W)
Public Const conMenu_Edit_BatExecute = 3098       '*医嘱批量执行
Public Const conMenu_Edit_SendBilling = 30133   '发送记帐单
Public Const conMenu_Edit_SendCharge = 30134   '发送收费单

Public Const conMenu_Edit_SendDrug = 30131      '*药疗医嘱发送(&1)
Public Const conMenu_Edit_SendOther = 30132     '其它医嘱发送(&2)
Public Const conMenu_Edit_SendInfusion = 30135 '*发送输液药品医嘱
Public Const conMenu_Edit_Untread = 3014    '*回退(&R)
Public Const conMenu_Edit_SendBack = 3015   '*超期发送收回(&Z)
Public Const conMenu_Edit_Test = 3016       '*皮试结果(&T)
Public Const conMenu_Edit_ChargeOff = 3017       '*费用冲销(&E)
Public Const conMenu_Edit_NoPrint = 3018    '屏蔽打印(&I)
Public Const conMenu_Edit_ChargeDelApply = 3019    '*销帐申请(&L)
Public Const conMenu_Edit_ChargeDelAudit = 3020    '*销帐审核(&U)
Public Const conMenu_Edit_UnUse = 3021      '*标记未用(&H)
Public Const conMenu_Edit_Surplus = 3022      '*留存登记
Public Const conMenu_Edit_MediAudit = 3023          '*药嘱审查(&U)(合理用药审查)
Public Const conMenu_Edit_ExtraFeeMove = 3024       '*附费转移
Public Const conMenu_Edit_ExtraFeeExe = 3025        '*附费执行
Public Const conMenu_Edit_ExtraFeeUnExe = 3026      '*附费取消执行
Public Const conMenu_Edit_AdviceUnAudit = 3027      '*抗菌药物取消审核
Public Const conMenu_Edit_LISApply = 3028      '*检验申请
Public Const conMenu_Edit_LISApplyModi = 3029      '*修改检验申请
Public Const conMenu_Edit_LISApplyDel = 3030      '*取消检验申请

'临床路径管理(编辑)
Public Const conMenu_Edit_Report = 816          '*出径登记表(&P)

'护理记录中使用到的复制,粘贴以及插入特殊符号
Public Const conMenu_Edit_Copy = 3031      '*复制
Public Const conMenu_Edit_PASTE = 3032      '*粘贴
Public Const conMenu_Edit_SPECIALCHAR = 3033      '*插入特殊符号
Public Const conMenu_Edit_Clear = 3034      '*清除内容
Public Const conMenu_Edit_PrevPage = 3035   '*上一页
Public Const conMenu_Edit_NextPage = 3036   '*下一页
Public Const conMenu_Edit_Word = 3037       '*词句选择
Public Const conMenu_Edit_Brief = 3038      '*小结
Public Const conMenu_Edit_Group_New = 3039          '*新分组
Public Const conMenu_Edit_Group_Append = 3040       '*追加分组
Public Const conMenu_Edit_Curve = 3041       '*体温曲线编辑
Public Const conMenu_Edit_CurveTable = 3042       '*体温表格编辑
Public Const conMenu_Edit_Curve_Show = 3043       '*体温曲线显示

'财务监控菜单
Public Const conMenu_Edit_CollectFees = 3588    '财务组收款
Public Const conMenu_Edit_CollectFees_Cancel = 3589  '财务组收款作废
Public Const conMenu_Edit_ReprintReceipt = 3591     '重打收据
Public Const conMenu_Edit_RollingCurtain = 3511    '*轧账(&Z)
Public Const conMenu_Edit_RollingCurtain_Cancel = 3512    '*作废轧账(&D)
Public Const conMenu_Edit_CheckCash = 3513    '*现金点钞(&E)
Public Const conMenu_Edit_ChargeBook_Reprint = 3514   '*重打缴款书(&E)
Public Const conMenu_Edit_RollingCurtain_Zero = 3515   '*轧帐归零(&C)
Public Const conMenu_Edit_StandbyMoeny_PutOut = 3516   '*发放备用金(&L)
Public Const conMenu_Edit_StandbyMoeny_OnWork = 3522   '*发放上岗备用金
Public Const conMenu_Edit_Personnel_Group = 3517   '*人员分组(&F)
Public Const conMenu_Edit_StandbyMoeny_PutIn = 3518   '*收回备用金(&H)
Public Const conMenu_Edit_Collect_Other = 3519  '其他人员收款
Public Const conMenu_Edit_Collect_Manual = 3520   '手工收款(&M)
Public Const conMenu_Edit_Collect_RollingCurtain = 3588   '*轧帐收款(&S)
Public Const conMenu_Edit_Collect_Cancel = 3589   '*收款作废(&S)
Public Const conMenu_Edit_DrawBook_Reprint = 3521   '*重打领用单


'护理部分管理功能列表
Public Const conMenu_Edit_WaveSynchro = 3044    '*体温项目同步设置
Public Const conMenu_Edit_CollectMan = 3045     '*汇总项目管理
Public Const conMenu_Edit_AnimalPart = 3046     '*体温部位管理
Public Const conMenu_Edit_FileMan = 3047        '*护理文件管理
Public Const conMenu_Edit_WavyMan = 3048        '波动项目管理
Public Const conMenu_Edit_MOBILE = 3049         '移动基础设置:1)输液配套项目设置(设置液体名称\液体量对应的项目,用于自动生成入量草稿数据);2)体温薄管理;3)护理项目分类管理


'病历(编辑)菜单
Public Const conMenu_Edit_NewParent = 3051   '*新分类(&N)
Public Const conMenu_Edit_Insert = 3052      '*插入(&I)
Public Const conMenu_Edit_ModifyParent = 3053    '*修改分类(&M)
Public Const conMenu_Edit_DeleteParent = 3054    '*删除分类(&D)
Public Const conMenu_Edit_MarkMap = 3061     '*图片(&I)…
Public Const conMenu_Edit_MarkKeyMap = 3070     '关键图像
Public Const conMenu_Edit_ApplyTo = 3062     '*适用科室(&T)
Public Const conMenu_Edit_Request = 3063     '限制要求(&R)
Public Const conMenu_Edit_Compend = 3064     '*内容构造(&F)
Public Const conMenu_Edit_Affix = 3069      '附项模板设置
Public Const conMenu_Edit_ElementChange = 3065      '*要素联动设置
Public Const conMenu_Edit_ImportMerge = 3066      '导入合并路径
Public Const conMenu_Edit_UnImportMerge = 3067      '取消导入合并路径
Public Const conMenu_Edit_ViewMergeImport = 3068      '查看合并路径导入评估
Public Const conMenu_Edit_Import = 3071      '*成批导入(&B)…

'医嘱(编辑)菜单
Public Const conMenu_Edit_BatUnPack = 3072      '批量打包
Public Const conMenu_Edit_PacsApply = 3073      '*检查申请
Public Const conMenu_Edit_PacsApplyModi = 3074      '*修改检查申请
Public Const conMenu_Edit_Apply = 3076      '*下达申请
Public Const conMenu_Edit_ApplyModi = 3077      '*修改申请
Public Const conMenu_Edit_ApplyDel = 3078      '*取消申请
Public Const conMenu_Edit_BloodApply = 3079    '输血申请
Public Const conMenu_Edit_TraReaction = 30791    '输血反应
Public Const conMenu_Edit_TraReactionRecord = 30792    '输血反应(多个)
Public Const conMenu_Edit_BloodApplyModi = 3080    '修改输血申请
Public Const conMenu_Edit_ApplyView = 3081    '查看申请单

'病历(编辑)菜单
Public Const conMenu_Edit_Adjust = 3082      '*调整(&J)
Public Const conMenu_Edit_Archive = 3083     '*归档(&R)
Public Const conMenu_Edit_UnArchive = 3084     '取消归档(&D)
Public Const conMenu_Edit_QCReport = 3085     '质控报告(&D)

'质控报告
Public Const conMenu_Edit_ItemEdit = 3086      '编辑
Public Const conMenu_Edit_ItemUndo = 3087       '取消
Public Const conMenu_Edit_ItemSave = 3088       '保存
Public Const conMenu_Edit_Exit = 3089           '退出
Public Const conMenu_Verify_AuditingLogin = 3090    '归档
Public Const conMenu_Verify_LogOut = 3099           '取消归档



Public Const conMenu_Edit_Save = 3091        '*保存
Public Const conMenu_Edit_Sort = 3092        '*多文档排序
Public Const conMenu_Edit_Privacy = 3093     '*病人隐私保护设置
Public Const conMenu_Edit_Select = 3094      '*选择
Public Const conMenu_Edit_DeSelect = 3095    '*取消选择
Public Const conMenu_Edit_Merge = 3096
Public Const conMenu_Edit_Dilute = 3098      '标本稀释

Public Const conMenu_Edit_OperationApply = 3099    '手术申请
Public Const conMenu_Edit_ConsultationApply = 3100    '会诊申请

'临床路径应用 径外项目移动排序
Public Const conMenu_Edit_Up = 3301         '上移
Public Const conMenu_Edit_Down = 3302       '下移
Public Const conMenu_Edit_SaveSorted = 3303     '保存排序
'Public Const conMenu_Manage_ThingAdd = 227    '接单(&A)
'Public Const conMenu_Manage_ThingModi = 228   '*调整执行情况(&M)
Public Const conMenu_Edit_Transf_Delete = 229   '撤消接单

'医疗卡管理
Public Const conMenu_Edit_CardBound = 3839    '绑定卡
Public Const conMenu_Edit_CancelCardBound = 3842    '取消绑定卡
Public Const conMenu_Edit_CardLoss = 3833    '挂失
Public Const conMenu_Edit_CardCancelLoss = 3834    '取消挂失
Public Const conMenu_Edit_Cardtrade = 3835    '换卡
Public Const conMenu_Edit_CardFill = 3836    '补卡
Public Const conMenu_Edit_CardBackMoney = 3837    '退款
Public Const conMenu_Edit_ChangePassWord = 3838    '调整密码
Public Const conMenu_Edit_ChangePassWord_Force = 3843    '强制调整密码
Public Const conMenu_Edit_MzToZy = 3840
Public Const conMenu_Edit_ZyToMz = 3841
Public Const conMenu_Edit_Family = 3844
Public Const conMenu_View_Family = 3845

'病人费用查询
'----------------------------------------------------------------------
Public Const conMenu_Edit_PreBalance = 817    '预结当前病人
Public Const conMenu_Edit_PreBalanceAll = 818    '预结所有病人
Public Const conMenu_Edit_Balance = 3011    '结帐
Public Const conMenu_Edit_Billing = 3003    '记帐
Public Const conMenu_Edit_Billing_Mulit = 3872    '批量记帐

Public Const conMenu_Edit_ReBilling = 3004    '直接销帐

Public Const conMenu_Edit_ReBillingButton = 3017       '*费用冲销(&E)
Public Const conMenu_Edit_ReBillingApply = 3019    '*销帐申请(&L)
Public Const conMenu_Edit_ReBillingAudit = 3020    '*销帐审核(&U)

Public Const conMenu_Edit_FeeAudit = 3564    '审核或开始审核
Public Const conMenu_Edit_FeeUnAudit = 3565    '取消审核
Public Const conMenu_Edit_OverFeeAudit = 3566    '完成审核
Public Const conMenu_Edit_PatiMemo = 3567   '备注信息编辑
Public Const conMenu_Edit_PrePayMoney = 3568    '预交款


'消费卡管理
'---------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_CardPay = 3811    '发卡
Public Const conMenu_Edit_CardBathPay = 3812    '批量发卡
Public Const conMenu_Edit_CardBack = 3813    '退卡
Public Const conMenu_Edit_CardCancelBack = 38131    '取消回收
Public Const conMenu_Edit_CardCallBack = 3814    '回收
Public Const conMenu_Edit_CardCancelCallBack = 38141    '取消回收

Public Const conMenu_Edit_CardInFull = 3816    '充值
Public Const conMenu_Edit_CardInFullBack = 3817    '充值回退
Public Const conMenu_Edit_CardModify = 3818    '修改卡信息
Public Const conMenu_Edit_CardResume = 3819    '卡启用
Public Const conMenu_Edit_CardStop = 38191    '卡停用
Public Const conMenu_Edit_MoveCard = 3821    '发卡时，移出卡片
Public Const conMenu_Apply_AllCard = 3822    '发卡时，根据当前单据，应用于所有需要发卡的单据
Public Const conMenu_Apply_AllColumn = 3823    '发卡时，根据当前单据指定的列，应用于所有需要发卡的此列信息
Public Const conMenu_COMBOX_INTERFACE = 3820    '消费卡接口
Public Const conMenu_Square_BrushCard = 3824    '消费卡目录+接口序号

'票据入库
Public Const conMenu_Edit_DamnifyAdd = 3831    '报损增加
Public Const conMenu_Edit_DamnifyDelete = 3832  '报损删除
Public Const conMenu_Edit_UserType = 3833          '票据使用类别

'分诊管理
'----------------------------------------------------------------------------------
Public Const conMenu_Edit_Triage = 2604  '分诊
Public Const conMenu_Edit_ModiyPati = 2607  '调整病人信息
Public Const conMenu_Edit_ModiyPatiBaseInfo = 2610 '73743:病人基本信息调整
Public Const conmenu_Edit_BackHospitalize = 3086    '回诊
Public Const conmenu_Edit_BackHospitalizeCancel = 3087    '回诊取消

Public Const conmenu_Edit_ChangeNum = 3088    '变号
Public Const conmenu_Edit_Leave = 3089    '病人不就诊
Public Const conmenu_Edit_Wait = 3090    '病人待诊

Public Const conmenu_View_TriagePati = 7101    '显示已分诊病人
Public Const conmenu_View_AdmissionsPati = 7102    '显示已接诊病人
Public Const conmenu_View_OverPati = 7103    '显示已完成病人
Public Const conmenu_View_Leave = 7104    '显示不就诊病人
Public Const conmenu_View_AutoRefresh = 7120    '自动刷新

'挂号安排
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_PlanAdd = 6820    '计划安排
Public Const conMenu_Edit_PlanModify = 6821    '修改计划安排
Public Const conMenu_Edit_PlanDelete = 6822    '删除计划安排
Public Const conMenu_Edit_PlanVerify = 6823    '审核计划安排
Public Const conMenu_Edit_PlanCancel = 6824    '取消审核计划
Public Const conMenu_Edit_AllStartNO = 6825  '全部启用挂号序号控制
Public Const conMenu_Edit_AllStopNO = 6826    '全部停用挂号序号控制
Public Const conMenu_Edit_StopPlanTimes = 6827  '停用安排计划
Public Const conMenu_Edit_ClearStopPlan = 6828  '清除所有停用安排计划
Public Const comMenu_Edit_SetDateSegment = 6829    '挂号安排时段 时间段设置
Public Const conMenu_Edit_SetPlanDateSeqment = 6830    '挂号计划时段 时间段设置
Public Const comMenu_Edit_UnitRegModify = 6831      '挂号安排合作单位序号分配
Public Const ComMenu_Edit_UnitRegArrangeModify = 6832    '挂号安排计划合作单位序号分配
Public Const ComMenu_Edit_AutoDefaultLimitAppointment = 6833    '挂号安排自动默认限约数

'体检系统补增 32开头的号
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_SaveExit = 3200           '保存并退出
Public Const conMenu_Edit_SizeFit = 3201            '格式调整
Public Const conMenu_Edit_SourceFit = 3202          '来源调整
Public Const conMenu_Edit_Camera = 3203             '照相设备
Public Const conMenu_Edit_TakePicture = 3204        '拍照
Public Const conMenu_Edit_SelAll = 3205             '全选
Public Const conMenu_Edit_ClsAll = 3206             '全清
Public Const conMenu_Edit_CallBack = 3207           '复查设置
Public Const conMenu_Edit_Money = 3208              '费用方式
Public Const conMenu_Edit_Pay = 3209                '支付方式
Public Const conMenu_Edit_CheckItem = 3210          '体检项目
Public Const conMenu_Edit_ChargeItem = 3211         '收费项目

'病人项目(编辑)菜单 3501-3530
Public Const conMenu_Edit_Transf_Modify = 3502   '修改单据
Public Const conMenu_Edit_Transf_UndoEnd = 3505  '撤消完成
Public Const conMenu_Edit_Transf_Negative = 3506    '阳性(+)
Public Const conMenu_Edit_Transf_Positive = 3507    '阴性(-)
Public Const conMenu_Edit_Transf_Reprint = 3508  '重打单据

Public Const conMenu_Edit_Transf_Liquid = 3509   '配液操作
Public Const conMenu_Edit_Transf_Puncture = 3510    '穿刺操作


'病人座位(编辑)菜单 3531-3559
Public Const conMenu_Edit_Seat = 3530        '座位
Public Const conMenu_Edit_Seat_Add = 3531    '座位增加
Public Const conMenu_Edit_Seat_Modify = 3532    '座位修改
Public Const conMenu_Edit_Seat_Delete = 3533    '座位删除
Public Const conMenu_Edit_Seat_Clear = 3534  '清除占用的座位
Public Const conMenu_Edit_Seat_Set = 3535    '安排座位
Public Const conMenu_Edit_Seat_Swap = 3536    '调换座位

Public Const conMenu_Edit_Seat_View = 3551    '查看
Public Const conMenu_Edit_Seat_Icon = 3552    '图标方式
Public Const conMenu_Edit_Seat_List = 3553    '列表方式
Public Const conMenu_Edit_Seat_Report = 3554    '报表方式

Public Const conMenu_Edit_View_Seat = 3550  '座位图例
Public Const conMenu_Edit_View_GBed = 3555    '普通床位
Public Const conMenu_Edit_View_RBed = 3556    '占用床位
Public Const conMenu_Edit_View_YBed = 3557    '维护床位

Public Const conMenu_Edit_View_Gseat = 3558    '普通座位
Public Const conMenu_Edit_View_Rseat = 3559    '占用座位
Public Const conMenu_Edit_View_Yseat = 3560    '维护座位


'暂存药品(编辑)菜单 3561 -3579
Public Const conMenu_Edit_Leave_Add = 3561    '增加
Public Const conMenu_Edit_Leave_Modify = 3562    '修改
Public Const conMenu_Edit_Leave_Delete = 3563    '删除
Public Const conMenu_Edit_Leave_Post = 3564    '使用登记
Public Const conMenu_Edit_Leave_SavePost = 3565    '保存登记数据
Public Const conMenu_Edit_Leave_UndoPost = 3565    '撤消登记

Public Const conMenu_Edit_Leave_Repertory = 3571    '库存查询
Public Const conMenu_Edit_Leave_AccountBook = 3572    '库存台帐

'手麻系统补增 3580 -  3599
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_CopyNewItem = 3580        '*复制并新项目
Public Const conMenu_Edit_Default = 3582            '缺省结果
Public Const conMenu_Edit_MakeCharge = 3586         '生成费用
Public Const conMenu_Edit_Preferences = 3587         '参考方案

'血库系统补增 31开头的号
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_NewKind = 311             '新增品种
Public Const conMenu_Edit_ModifyKind = 312          '修改品种
Public Const conMenu_Edit_DeleteKind = 313          '删除品种
Public Const conMenu_Edit_StorgeLimit = 314         '库存限量
Public Const conMenu_Edit_StorgeDept = 315          '库房
Public Const conMenu_Edit_StorgePostion = 316       '货位
Public Const conMenu_Edit_Check = 3101              '核对
Public Const conMenu_Edit_View = 3102               '查阅
Public Const conMenu_Edit_ModifyBill = 3103         '修改发票
Public Const conMenu_Edit_Verify = 3104             '常规检验
Public Const conMenu_Edit_AdjustPrice = 3105        '调价

'病区床位管理(编辑)菜单 3601-3610
Public Const conMenu_Edit_Bed_Add = 3601            '新增
Public Const conMenu_Edit_Bed_Modify = 3602         '调整
Public Const conMenu_Edit_Bed_Delete = 3603         '撤销
Public Const conMenu_Edit_Bed_ToRepair = 3604       '转修缮
Public Const conMenu_Edit_Bed_ToEmpty = 3605        '转空床
Public Const conMenu_Edit_SelUnit = 3606            '病区选择


'院感编辑\查看菜单
Public Const conMenu_Edit_DelDayItem = 3802        '删除日报当前行信息
Public Const conMenu_Edit_BuildConstant = 3803        '生成常用消毒项目
'
'卫生材料\查看菜单
Public Const conMenu_Edit_CfPay = 4000        '按处方发料
Public Const conMenu_Edit_BillPay = 4001        '按票据发料
Public Const conMenu_Edit_BillBackPay = 4002        '按单据退料
Public Const conMenu_Edit_StopPay = 4003        '按停止发料标记

Public Const conMenu_View_FontSize = 4004         '字号设置
Public Const conMenu_View_FontSize_1 = 4004         '9号字
Public Const conMenu_View_FontSize_2 = 4004         '11号字
Public Const conMenu_View_FontSize_3 = 4004         '15号字

Public Const conMenu_Edit_OtherPay = 4005        '发其他库房处方

'预约登记管理
Public Const conMenu_Edit_AppRequest = 4200
Public Const conMenu_Edit_CancelRequest = 4201
Public Const conMenu_Edit_ViewRequest = 4202
Public Const conMenu_Edit_AppRequestManage = 4203


'LIS使用的采单 3650-3690
Public Const conMenu_Edit_QCRes = 3650         '*质控品
Public Const conMenu_LIS_Cancel = 3651         '*取消
Public Const conMenu_LIS_PatientInfo = 3652    '病人信息
Public Const conMenu_LIS_HideList = 3653       '隐藏病人列表
Public Const conMenu_LIS_TOQC = 3654           '置为质控
Public Const conMenu_LIS_SendReport = 3655     '发送报告单
Public Const conMenu_LIS_SignVerify = 3656     '验证签名
Public Const conMenu_LIS_MB_Connect = 3701     '酶标仪连接
Public Const conMenu_LIS_MB_Disconnect = 3702  '酶标仪断开
Public Const comMenu_LIS_TodayQC = 3703        '今日质控
Public Const comMenu_LIS_History = 3704        '历史质控
Public Const comMenu_LIS_ShowListHead = 3705   '选择要显示的列
Public Const conMenu_LIS_LJAverage = 3706      '均值LJ质控
Public Const conMenu_LIS_RightMenu = 3707      '右键菜单
Public Const conMenu_LIS_SaveSample = 3708   '保存标本
Public Const conMenu_LIS_DropSample = 3709   '销毁标本

'报表菜单
Public Const conMenu_Report_DrugQuery = 401    '药疗收发查询(&H)
Public Const conMenu_Report_Reports = 402      '病区常用报表(&W)
Public Const conMenu_Report_MultiBill = 403    '打印多病人单据(&K)
Public Const conMenu_Report_ClinicBill = 404   '打印诊疗单据(&J)…
Public Const conMenu_Report_AdviceBill1 = 405  '长期医嘱单(&P)
Public Const conMenu_Report_AdviceBill2 = 406  '临时医嘱单(&T)
Public Const conMenu_Report_AdviceBill3 = 407  '医嘱记录本(&B)
Public Const conMenu_Report_WorkLog = 408      '工作日报(&O)

'查看菜单
Public Const conMenu_View_ToolBar = 701              '工具栏(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Public Const conMenu_View_StatusBar = 702            '状态栏(&S)
Public Const conMenu_View_Append = 703               '附加信息(&A)
Public Const conMenu_View_Difference = 704              '显示差异(&D)
Public Const conMenu_View_Contrast = 705                 '对比查看
Public Const conMenu_View_NoticBoard = 706           '病区公告栏

Public Const conMenu_View_FontSize_S = 4041            '医嘱字体：小字体
Public Const conMenu_View_FontSize_M = 4040            '医嘱字体：中字体
Public Const conMenu_View_FontSize_L = 4042            '医嘱字体：大字体

Public Const conMenu_View_StPath = 4043                '查看标准路径
Public Const conMenu_View_Expend = 711               '展开/折叠组(&X)
Public Const conMenu_View_Expend_CurCollapse = 7111     '折叠当前组(&C)
Public Const conMenu_View_Expend_CurExpend = 7112       '展开当前组(&E)
Public Const conMenu_View_Expend_AllCollapse = 7113     '折叠所有组(&L)
Public Const conMenu_View_Expend_AllExpend = 7114       '展开所有组(&X)
Public Const conMenu_View_Find = 721                 '*查找(&F)
Public Const conMenu_View_FindNext = 722             '继续查找(&N)
Public Const conMenu_View_FindType = 723             '查找方式(&Y)
Public Const conMenu_View_ReadIC = 724               '读IC卡(&I)
Public Const conMenu_View_PatInfor = 725             '查看病人信息
Public Const conMenu_View_PatiInput = 726             '显示病人信息快捷输入面板
Public Const conMenu_View_PriceBill = 727
Public Const conMenu_View_PriceTable = 728
Public Const conMenu_View_PriceList = 729
Public Const conMenu_View_FilterView = 730           '以过滤方式显示
Public Const conMenu_View_Filter = 731               '*数据过滤(&I),子窗体的过滤功能
Public Const conMenu_View_Notify = 732               '*医嘱提醒(&B)
Public Const conMenu_View_Busy = 733                 '诊室忙(&M)
Public Const conMenu_View_ShowAll = 734
Public Const conMenu_View_ShowHistory = 735
Public Const conMenu_View_ShowStoped = 736
Public Const conMenu_View_ShowDel = 738           '显示删除安排
Public Const conMenu_View_Hide = 741                 '*隐藏(&H)
Public Const conMenu_View_Show = 742                 '*显示(&S)
Public Const conMenu_View_Forward = 743              '*前进(&F)
Public Const conMenu_View_Backward = 744             '*后退(&B)
Public Const conMenu_View_Dept = 745                '查看部门
Public Const conMenu_View_Location = 746            '定位
Public Const conMenu_View_LocationItem = 747        '定位项目
Public Const conMenu_View_Option = 781               '选项(&O)
Public Const conMenu_View_Refresh = 791              '*刷新(&R)
Public Const conMenu_View_RefreshSpare = 7911        '读取备用医嘱
Public Const conMenu_View_Jump = 792                 '跳转(&J)
Public Const conMenu_View_Warrant = 794              '担保信息查阅
Public Const conMenu_View_Shell = 3818                '外接程序


Public Const conMenu_View_Navigatebeginning = 7401           '*第一个(&F)
Public Const conMenu_View_Navigateleft = 7402                '*上一个(&F)
Public Const conMenu_View_Navigateright = 7403               '*下一个(&F)
Public Const conMenu_View_Navigateend = 7404                 '*最后一个(&F)
Public Const conMenu_View_OneWeek = 7405 '*第一周
Public Const conMenu_View_TwotWeek = 7406 '*第二周
Public Const conMenu_View_ThreeWeek = 7407 '*第三周
Public Const conMenu_View_FourWeek = 7408 '*第四周
'财务监控菜单
Public Const conMenu_View_Detail = 7501             '查看明细数据
Public Const conMenu_View_ChargeAndBilllTotal = 7502             '查看收款及票据汇总

'病人结帐管理
'----------------------------------------------------------------------
Public Const conMenu_File_CashCount = 4801    '现金点钞
Public Const conMenu_File_SetInsure = 4802    '保险类别
Public Const conMenu_Edit_ClinicBalance = 4803
Public Const conMenu_Edit_InHosBalance = 4804
Public Const conMenu_Edit_BatchBalance = 4805
Public Const conMenu_Edit_UnitBalance = 4806
Public Const conMenu_Edit_FeeManage = 4807
Public Const conMenu_Edit_ClinicToHos = 4808
Public Const conMenu_Edit_ToHosCancel = 4809
Public Const conMenu_Edit_ErrReBalance = 4810
Public Const conMenu_Edit_ErrCancelBalance = 4811
Public Const conMenu_Edit_ErrDelBalance = 4812
Public Const conMenu_Edit_CancelBalance = 4813
Public Const conMenu_Edit_PrintAmend = 4814
Public Const conMenu_Edit_PrintDetail = 4815
Public Const conMenu_Edit_PrintAmendByPati = 4816
Public Const conMenu_Edit_WriteCard = 4817
Public Const conMenu_View_RefreshType = 4818
Public Const conMenu_View_RefreshType_No = 4819
Public Const conMenu_View_RefreshType_Ask = 4820
Public Const conMenu_View_RefreshType_Auto = 4821
Public Const conMenu_File_FeeCollect = 4822

'病人费用查询
Public Const conMenu_View_Billing = 3551             '查看记帐单
Public Const conMenu_View_DateType = 781           '查询时间
Public Const conMenu_View_DetailType = 793          '清单类型
Public Const conMenu_View_GroupCol = 733            '分组字段

Public Const conMenu_View_ReBalance = 7510  '显示结帐作废
Public Const conMenu_View_ZeroFee = 7511    '显示零费用
Public Const conMenu_View_CheckFee = 7512   '显示体检费用
Public Const conMenu_View_Owe = 7513        '显示未结清病人
Public Const conMenu_View_UnAudit = 7514     '显示未审核病人
Public Const conMenu_View_OnePati = 7515    '多次住院只显一次病人
Public Const conMenu_View_TurnToWardFeeQuery = 7516    '转病区费用变动查询

'人员分组(用于ListView控件的显示方式:大图标;小图标;列表;详细资料)
Public Const conMenu_View_LargeICO = 7610  '大图标
Public Const conMenu_View_MinICO = 7611  '小图标
Public Const conMenu_View_ListICO = 7612  '列表
Public Const conMenu_View_DetailsICO = 7613  '详细列表

'体检系统补增70开头的号
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_View_Single = 7040             '个人
Public Const conMenu_View_Group = 7041              '团体
Public Const conMenu_View_LocationMethod = 7042     '定位处理
Public Const conMenu_View_Column = 7043             '选择列项

Public Const conMenu_View_LocationRange = 7044     '定位范围

'工具菜单
Public Const conMenu_Tool_Reference = 801       '*参考(&R)
Public Const conMenu_Tool_Reference_1 = 8011    '疾病诊断参考(&D)
Public Const conMenu_Tool_Reference_2 = 8012    '诊疗措施参考(&C)
Public Const conMenu_Tool_MedRec = 802          '*首页整理(&M)
Public Const conMenu_Tool_Meet = 803            '*病人会诊(&E)
Public Const conMenu_Tool_MeetOpen = 8030         '接受会诊
Public Const conMenu_Tool_MeetFinish = 8031         '完成会诊(&F)
Public Const conMenu_Tool_MeetCancel = 8032         '取消完成(&C)
Public Const conMenu_Tool_MeetIdea = 8033         '填写会诊意见(&W)
Public Const conMenu_Tool_Sign = 804            '*电子签名(&I)
Public Const conMenu_Tool_SignNew = 8041            '电子签名(&I)
Public Const conMenu_Tool_SignVerify = 8042         '验证签名(&V)
Public Const conMenu_Tool_SignEarse = 8043          '取消签名(&E)
Public Const conMenu_Tool_SignAuditAffirm = 8044         '*上级审签(&V)
Public Const conMenu_Tool_SignAuditCancel = 8045          '取消审签(&E)
Public Const conMenu_Tool_Community = 805       '*社区档案(&U)
Public Const conMenu_Tool_MedRecAudit = 806        '病案审查(&M)
Public Const conMenu_Tool_MedRecAuditSubmit = 8061      '提交审查(&S)
Public Const conMenu_Tool_MedRecAuditCancel = 8062      '取消提交(&C)
Public Const conMenu_Tool_MedRecAuditResponse = 8063    '审查反馈(&M)
Public Const conMenu_Tool_MedRecAuditWriteResponse = 8064    '书写审查意见
Public Const conMenu_Tool_Archive = 807         '*人员档案(&I) / 门诊住院 电子病案查阅
Public Const conMenu_Tool_ExaReport = 808 '查阅体检总检报告
Public Const conMenu_Tool_Monitor = 811         '*监测(&M)
Public Const conMenu_Tool_Monitor_1 = 81101         '时限要求监测(&T)
Public Const conMenu_Tool_Monitor_2 = 81102         '内容要求监测(&C)
Public Const conMenu_Tool_Assistant = 812       '*助手(&A)
Public Const conMenu_Tool_Analyse = 813         '*分析(&Y)
Public Const conMenu_Tool_Search = 814          '*检索(&S)
Public Const conMenu_Tool_Define = 815          '*定义(&D)
Public Const conMenu_Tool_Report = 816          '*报告(&P)
Public Const conMenu_Tool_Apply = 817           '*应用(&A)
Public Const conMenu_Tool_BathSend = 818        '批量发送到仪器
Public Const conMenu_Tool_Option = 819          '选项(&O),子窗体的设置功能
Public Const conMenu_Tool_KssAudit = 820        '*抗菌用药审核
Public Const conMenu_Tool_OPSAudit = 821        '手术审核管理
Public Const conMenu_Tool_CISMed = 822        '临床自管药
Public Const conMenu_Tool_TransAudit = 823   '输血分级管理
Public Const conMenu_Tool_MedRatio = 824      '药占比
Public Const conMenu_Tool_HealthCard = 825      '居民健康卡
Public Const conMenu_Tool_OPSEmpower = 826   '手术授权管理
Public Const conMenu_Tool_UnitSubject = 827     '病区标记设置
Public Const conMenu_Tool_UnitNBoard = 828      '病区公告栏设置
Public Const conMenu_Tool_RisPrint = 829       '打印RIS预约单
Public Const conMenu_Tool_RisPrintBat = 830    '批量打印RIS预约单

Public Const conMenu_Tool_Positive = 889     '传染病阳性结果
Public Const conMenu_Tool_PlugIn = 890          '插件
Public Const conMenu_Tool_PlugInPop = 891          '插件
Public Const conMenu_Tool_PlugIn_Item = 89000   '插件项,实际依次为 conMenu_Tool_PlugIn_Item + n, 1<=n<=99
Public Const conMenu_Tool_Critical = 892    '危急值

'PACS工作站菜单
Public Const conMenu_Manage_CriticalValues = 8342           '危急值记录
Public Const conMenu_Manage_CriticalSituation = 8343        '危急情况
Public Const conMenu_Manage_Normal = 8344                   '正常
Public Const conMenu_Manage_Critical = 8345                 '危急

Public Const conMenu_Manage_Result = 8300       '检查结果
Public Const conMenu_Manage_Negative = 8301      '检查结果阳性
Public Const conMenu_Manage_Positive = 8302      '检查结果阴性

Public Const conMenu_Manage_ImageQuality = 8303       '影像质量
Public Const conMenu_Manage_ImageFirst = 8304         '第一级
Public Const conMenu_Manage_ImageSecond = 8305        '第二级
Public Const conMenu_Manage_ImageThird = 8396         '第三级
Public Const conMenu_Manage_ImageFourth = 8397        '第四级

Public Const conMenu_Manage_ReportQuality = 8346       '报告质量
Public Const conMenu_Manage_ReportFirst = 8347         '第一级
Public Const conMenu_Manage_ReportSecond = 8348        '第二级
Public Const conMenu_Manage_ReportThird = 8349         '第三级
Public Const conMenu_Manage_ReportFourth = 8350        '第四级

Public Const conMenu_Manage_FuHeLevel = 8220         '符合情况
Public Const conMenu_Manage_FuHe = 8221             '符合
Public Const conMenu_Manage_JiBenFuHe = 8222        '基本符合
Public Const conMenu_Manage_BuFuHe = 8223           '不符合

Public Const conMenu_Manage_SwitchUser = 8338       '切换用户
Public Const conMenu_Manage_ChangeUser = 8306      '交换用户
Public Const conMenu_Manage_ChangeDevice = 8307     '更换设备影像类别
Public Const conMenu_Manage_ImageInterval = 8308    '打开图像间隔
Public Const conMenu_Manage_CopyCheck = 8200        '同一病人登记相同项目不同部位
Public Const conMenu_Manage_GChannel = 8201         '绿色通道
Public Const conMenu_Manage_GChannelOk = 8202       '绿色通道标记
Public Const conMenu_Manage_GChannelCancel = 8203   '绿色通道取消
Public Const conMenu_Manage_Review = 8204           '随访
Public Const conMenu_Manage_SelectAllImages = 8205      '全选图像
Public Const conMenu_Manage_UnSelectAllImages = 8206    '全清图像
Public Const conMenu_Manage_ReverseSelectImages = 8207  '反选图像
Public Const conMenu_Manage_TechDoctorExecute = 8208    '技师执行
Public Const conMenu_Manage_ReportRelease = 8209        '报告发放
Public Const conMenu_Manage_ReportExecutor = 8228       '报告执行
Public Const conMenu_Manage_RelatingPatiet = 8210       '关联病人
Public Const conMenu_Manage_LocateType = 8211           '定位方式
Public Const conMenu_Manage_LocateValue = 8212          '定位值
Public Const conMenu_Manage_DeleteSeries = 8213         '删除序列图像
Public Const conMenu_Manage_DeleteImage = 8214          '删除图像
Public Const conMenu_Manage_FilmRelease = 8215          '胶片发放
Public Const conMenu_Manage_ReportFilmRelease = 8216    '报告胶片同时发放
Public Const conMenu_Manage_Release = 8217              '发放
Public Const conMenu_Manage_Burn = 8218                 '刻录
Public Const conMenu_Manage_RefreshImg = 8219           '刷新图像
Public Const conMenu_Manage_Query = 8224                '查询
Public Const conMenu_Manage_ConfigQuery = 8225          '配置查询方案
Public Const conMenu_Manage_CustomQuery = 8226          '自定义查询
Public Const conMenu_Manage_SetXWParam = 8227             'RISPACS参数
Public Const conMenu_Manage_CloseQuery = 8229           '关闭查询
Public Const conMenu_Manage_FilmPrevew = 8230           '胶片预览
Public Const conMenu_Manage_FilmPrint = 8231            '胶片打印
Public Const conMenu_Manage_FilmDelete = 8232           '胶片删除
Public Const conMenu_Manage_CheckList = 8233            '查看申请单
Public Const conMenu_Manage_SendArrange = 8234          '发送安排
Public Const conMenu_Manage_PacsPlugIn = 8235           'Pacs第三方功能挂接
Public Const conMenu_Manage_PacsPlugCfg = 8236          '插件配置

'PACS报告编辑器
Public Const conMenu_PacsReport_Group = 8336        '报告
Public Const conMenu_PacsReport_SelFormat = 8309    '选择报告格式
Public Const conMenu_PacsReport_SelFormat_Item = 8310    '选择报告格式
Public Const conMenu_PacsReport_Save = 8311         '保存报告
Public Const conMenu_PacsReport_Sign = 8312         '报告签名
Public Const conMenu_PacsReport_Reject = 8340         '报告驳回
Public Const conMenu_PacsReport_RejectHistory = 8341  '驳回历史
Public Const conMenu_PacsReport_DelSign = 8313      '回退签名
Public Const conMenu_PacsReport_MoveUp = 8314       '图像前移
Public Const conMenu_PacsReport_MoveDown = 8315     '图像后移
Public Const conMenu_PacsReport_DelImage = 8316     '删除图像
Public Const conMenu_PacsReport_DelMarks = 8317     '清除标注
Public Const conMenu_PacsReport_Open = 8318         '打开报告编辑窗体
Public Const conMenu_PacsReport_FontSet = 8319      '设置大文本段字体
Public Const conMenu_PacsReport_History = 8320      '修订历史
Public Const conMenu_PacsReport_Mode_Orig = 8321    '原始状态
Public Const conMenu_PacsReport_Mode_Clear = 8322   '最终状态
Public Const conMenu_PacsReport_History_Times = 8323    '历史报告
Public Const conMenu_PacsReport_DelMiniImage = 8324     '删除报告缩略图
Public Const conMenu_PacsReport_SelMiniImage = 8325     '提取报告缩略图
Public Const conMenu_PacsReport_RptImg2CapImg = 8326    '在报告图工具和采集工具间切换
Public Const conMenu_PacsReport_PrivOrder = 8327        '上一个医嘱
Public Const conMenu_PacsReport_NextOrder = 8328        '下一个医嘱
Public Const conMenu_PacsReport_AddNumber = 8329        '给段落文字添加序号
Public Const conMenu_PacsReport_RepFormat = 8330        '自定义报表选择格式
Public Const conMenu_PacsReport_RepFormat_Item = 8331   '自定义报表选择的具体格式项
Public Const conMenu_PacsReport_SaveWord = 8332         '保存词句示范
Public Const conMenu_PacsReport_ClearWritingState = 8333    '清除报告当前操作人
Public Const conMenu_PacsReport_VerifySign = 8334           '报告签名验证
Public Const conMenu_PacsReport_VerifySign_Item = 8335      '报告签名验证,对具体版本做验证
Public Const conMenu_PacsReport_Default = 8337              '恢复默认界面
Public Const conMenu_PacsReport_FinalShowMode = 8339        '最终状态显示

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


'病理菜单 3900-3970
Public Const conMenu_PatholManage = 3900                '病理管理
Public Const conMenu_Pathol_Antibody_Manage = 3901      '抗体管理
Public Const conMenu_Pathol_MealManage = 3902           '套餐维护
Public Const conMenu_Pathol_Request = 3903              '病理申请
Public Const conMenu_Pathol_ReportDelay = 3904          '报告延迟
Public Const conMenu_Pathol_ConRequest = 3905           '会诊申请
Public Const conMenu_Pathol_ConFeedback = 3906          '会诊反馈
Public Const conMenu_Pathol_Decalin_Task = 3907         '脱钙任务管理
Public Const conMenu_Pathol_BatSlicesAccept = 3908      '制片批量接受
Public Const conMenu_Pathol_BatSlicesSure = 3909        '制片批量确认
Public Const conMenu_Pathol_BatSpeExamAccept = 3910     '特检批量接受
Public Const conMenu_Pathol_BatSpeExamSure = 3911       '特检批量确认
Public Const conMenu_Pathol_BatProcess = 3912           '批量处理
Public Const conMenu_Pathol_Quality_Manage = 3913       '病理质量管理
Public Const conMenu_Pathol_NumConfig = 3914            '病理号码配置
Public Const conMenu_Pathol_WorkModule = 3915           '站点模式配置
Public Const conMenu_PatholSlices_Quality = 3916        '病理制片质量

'病理标本菜单
Public Const conMenu_PatholSpecimen = 3940                  '病理标本
Public Const conMenu_PatholSpecimen_LAB = 3941              '标签
Public Const conMenu_PatholSpecimen_PreviewLab = 3942       '预览标签
Public Const conMenu_PatholSpecimen_PrintLab = 3943         '打印标签
Public Const conMenu_PatholSpecimen_ACP = 3944              '核收单
Public Const conMenu_PatholSpecimen_PrintAccept = 3945      '打印核收单
Public Const conMenu_PatholSpecimen_PreviewAccept = 3946    '预览核收单
Public Const conMenu_PatholSpecimen_Get = 3947              '提取标本
Public Const conMenu_PatholSpecimen_Del = 3948              '删除标本
Public Const conMenu_PatholSpecimen_Save = 3949             '保存标本
Public Const conMenu_PatholSpecimen_Accept = 3950           '核收标本
Public Const conMenu_PatholSpecimen_Reject = 3951           '拒收标本


'病理制片菜单
Public Const conMenu_PatholSlices = 3960                    '病理制片
Public Const conMenu_PatholSlices_LAB = 3961                '标签
Public Const conMenu_PatholSlices_PreviewLAB = 3962         '预览标签
Public Const conMenu_PatholSlices_PrintLAB = 3963           '打印标签
Public Const conMenu_PatholSlices_List = 3964               '清单
Public Const conMenu_PatholSlices_PreviewList = 3965        '预览清单
Public Const conMenu_PatholSlices_PrintList = 3966          '打印清单
Public Const conMenu_PatholSlices_RequestView = 3967        '申请查看
Public Const conMenu_PatholSlices_Accept = 3968             '制片接受
Public Const conMenu_PatholSlices_Finish = 3969             '制片完成


'病理取材菜单
Public Const conMenu_PatholMaterial = 3970                  '病理取材
Public Const conMenu_PatholMaterial_PrintAll = 3971         '打印所有
Public Const conMenu_PatholMaterial_PreviewAll = 3972       '预览所有
Public Const conMenu_PatholMaterial_PrintSingle = 3973      '单个打印
Public Const conMenu_PatholMaterial_PreviewSingle = 3974    '单个预览
Public Const conMenu_PatholMaterial_RequestView = 3975      '申请查看
Public Const conMenu_PatholMaterial_Get = 3976              '材块提取
Public Const conMenu_PatholMaterial_Del = 3977              '删除材块
Public Const conMenu_PatholMaterial_Save = 3978             '保存材块
Public Const conMenu_PatholMaterial_Sure = 3979             '确认取材
Public Const conMenu_PatholMaterial_Decalcification = 3980  '脱钙
Public Const conMenu_PatholMaterial_ChangeVat = 3981        '换缸
Public Const conMenu_PatholMaterial_CancelVat = 3982        '撤销
Public Const conMenu_PahtolMaterial_Finish = 3983           '完成


'病理特检菜单
Public Const conMenu_PatholSpeExam = 3990                   '病理特检
Public Const conMenu_PatholSpeExam_LAB = 3991               '标签
Public Const conMenu_PatholSpeExam_PreviewLAB = 3992        '预览标签
Public Const conMenu_PatholSpeExam_PrintLab = 3993          '打印标签
Public Const conMenu_PatholSpeExam_List = 3994              '清单
Public Const conMenu_PatholSpeExam_PreviewList = 3995       '预览清单
Public Const conMenu_PatholSpeExam_PrintList = 3996         '打印清单
Public Const conMenu_PatholSpeExam_RequestView = 3997       '申请查看
Public Const conMenu_PatholSpeExam_Accept = 3998            '特检接受
Public Const conMenu_PatholSpeExam_Finish = 3999            '特检完成

'病理过程报告
Public Const conMenu_PatholProRep = 4100                    '病理过程报告
Public Const conMenu_PatholProRep_Print = 4101              '打印
Public Const conMenu_PatholProRep_Preview = 4102            '预览
Public Const conMenu_PatholProRep_Already = 4103            '已阅
Public Const conMenu_PatholProRep_Back = 4104               '撤回
Public Const conMenu_PatholProRep_Clear = 4105              '清空内容
Public Const conMenu_PatholProRep_Input = 4106              '特检项目录入
Public Const conMenu_PatholProRep_New = 4107                '新增报告
Public Const conMenu_PatholProRep_Del = 4108                '删除报告
Public Const conMenu_PatholProRep_Save = 4109               '数据保存

'病理套餐维护
Public Const conMenu_PatholMeal_Save = 4110
Public Const conMenu_PatholMeal_Cancel = 4111
Public Const conMenu_PatholMeal_AddRecord = 4112
Public Const conMenu_PatholMeal_ModRecord = 4113
Public Const conMenu_PatholMeal_DelRecord = 4114
Public Const conMenu_PatholMeal_UpRow = 4115
Public Const conMenu_PatholMeal_DownRow = 4116

'档案管理菜单
Public Const conMenu_Pathol_ArchivesManage = 3920    '档案管理
Public Const conMenu_Pathol_ArchivesClass = 3921    '档案类别配置
Public Const conMenu_Pathol_ArchivesPlace = 3922    '档案位置配置
Public Const conMenu_Pathol_ArchivesFile = 3923  '档案文件归档
Public Const conMenu_Pathol_ArchivesMaterial = 3924  '档案材料归档
Public Const conMenu_Pathol_ArchivesLend = 3925      '档案借阅管理


'收藏管理菜单
Public Const conMenu_Collection = 3930
Public Const conMenu_Collection_Manage = 3931    '收藏管理
Public Const conMenu_Collection_To = 3932    '收藏到
Public Const conMenu_Collection_ViewShare = 3933    '查看共享
Public Const comMenu_Collection_Type = 3934    '动态收藏菜单

'检查菜单
Public Const comMenu_Petition_Capture = 3935    '扫描申请单
Public Const comMenu_Petition_View = 3936       '查看申请单

'Pacs占用部分
Public Const conMenu_Reserve_23 = 4120
Public Const conMenu_Reserve_24 = 4125
Public Const conMenu_Reserve_25 = 4130
Public Const conMenu_Reserve_26 = 4135
Public Const conMenu_Reserve_27 = 4140
Public Const conMenu_Reserve_28 = 4145
Public Const conMenu_Reserve_29 = 4150
Public Const conMenu_Reserve_30 = 4155
Public Const conMenu_Reserve_31 = 4160
Public Const conMenu_Reserve_32 = 4165
Public Const conMenu_Reserve_33 = 4170
Public Const conMenu_Reserve_34 = 4175
Public Const conMenu_Reserve_35 = 4180
Public Const conMenu_Reserve_36 = 4185
Public Const conMenu_Reserve_37 = 4190
Public Const conMenu_Reserve_38 = 4195


Public Const conMenu_Img_Group = 8110           '影像
Public Const conMenu_Img_Look = 8127            '影像观片(&S)
Public Const conMenu_Img_Contrast = 8112        '观片对比(&E)
Public Const conMenu_Img_Look3D = 8111            '影像观片(&S)
Public Const conMenu_Img_Delete = 8113          '图象删除(&K)
Public Const conMenu_Img_Query = 8114           'Q/R获取图象(&Q)


'三维重建菜单
Public Const conMenu_Img_3D = 8115          '三维重建
Public Const conMenu_Img_3D_VA = 8116       '容积重建
Public Const conMenu_Img_3D_MPR = 8117      'MPR
Public Const conMenu_Img_3D_MMPR = 8118     'MMPR
Public Const conMenu_Img_3D_VE = 8119       '虚拟内窥镜
Public Const conMenu_Img_3D_SA = 8120       '表面重建
Public Const conMenu_Img_3D_PF = 8121       '灌注成像

'排队叫号系统
Public Const conMenu_Queue_CallThis = 8250      '直呼
Public Const conMenu_Queue_CallNext = 8251      '顺呼，呼叫下一个
Public Const conMenu_Queue_CallFirst = 8252     '优先
Public Const conMenu_Queue_Restore = 8253     '恢复
Public Const conMenu_Queue_ReCall = 8254        '重呼
Public Const conMenu_Queue_Abandon = 8255       '弃号
Public Const conMenu_Queue_Refresh = 8256       '刷新
Public Const conMenu_Queue_Setup = 8257         '参数设置
Public Const conMenu_Queue_Update = 8258        '修改
Public Const conMenu_Queue_Broadcast = 8259     '广播
Public Const conMenu_Queue_Pause = 8260         '暂停
Public Const conMenu_Queue_Finaled = 8261       '完成就诊
Public Const conMenu_Queue_Find = 8262          '查找
Public Const conMenu_Queue_ComeBack = 8263      '回诊
Public Const conMenu_Queue_RecDiagnose = 8264   '接诊

Public Const conMenu_Queue_Locate = 8265        '定位
Public Const conMenu_Queue_LocateValue = 8266    '定位值
Public Const conMenu_Queue_LocateType = 8267    '定位类型

Public Const conMenu_Queue_PrintNumber = 8272       '打号
Public Const conMenu_Queue_InsertQueue = 8273       '插队
Public Const conMenu_Queue_RestartQueue = 8274      '重排

'抗生素分级管理
Public Const conMenu_Kss_Jurisdiction = 9001    '权限
Public Const conMenu_Kss_Grant = 9002           '授权
Public Const conMenu_Kss_Cancellation = 9003    '取消授权
Public Const conMenu_Kss_Adjustment = 9004      '调整权限
Public Const conMenu_Kss_ShowCancel = 9005      '显示取消授权的人员

'药疗收发查询表格字体菜单
Public Const conMenu_FontSet = 509
Public Const conMenu_FontSet_FontSize_S = 4041            '药疗收发查询表格：小字体
Public Const conMenu_FontSet_FontSize_L = 4042            '药疗收发查询表格：大字体

'图像处理
Public Const conMenu_Process_Window = 501           '亮度对比度
Public Const conMenu_Process_Zoom = 502             '缩放
Public Const conMenu_Process_Small = 513             '缩小
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

'临床路径编辑菜单
Public Const conMenu_Edit_OutLogModi = 601    '修改出径登记
Public Const conMenu_Edit_OutLogView = 602    '查看出径登记
'标准路径维护
Public Const conMenu_Edit_NewCourseItem = 3001    '新增段落
Public Const conMenu_Edit_ModifyCourseItem = 3003    '修改段落
Public Const conMenu_Edit_DelCourseItem = 3004   '删除段落

Public Const conMenu_Edit_ModifyTableContent = 3822    '修改表单内容

Public Const conMenu_Edit_NewPath = 9002    '新增路径
Public Const conMenu_Edit_ModifyPath = 9004    '修改路径
Public Const conMenu_Edit_DelPath = 9003   '删除路径

Public Const conMenu_Edit_NewTable = 3051    '新增表单
Public Const conMenu_Edit_ModifyTable = 3053    '修改表单
Public Const conMenu_Edit_DelTable = 3054   '删除表单
'标准路径表单内容编辑
Public Const conMenu_NewRow = 100    '新增一行
Public Const conMenu_NewCol = 101    '新增一列
Public Const conMenu_DelCol = 102    '删除列
Public Const conMenu_DelRow = 103    '删除行
Public Const conMenu_ClearItem = 104    '清除内容
Public Const conMenu_Save = 107    '保存
Public Const conMenu_Edit = 108    '编辑
Public Const conMenu_Exit = 111    '退出
'帮助菜单
Public Const conMenu_Help_Help = 901        '*帮助主题(&H)
Public Const conMenu_Help_Web = 902         '&WEB上的中联
Public Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Public Const conMenu_Help_Web_Forum = 9023      '中联论坛(&F)
Public Const conMenu_Help_Web_Mail = 9022       '*发送反馈(&M)
Public Const conMenu_Help_About = 991       '关于(&A)…

'病案查询打印
Public Const conMenu_File_Word = 7047           '输出到&Word…
Public Const conMenu_File_PDF = 7048            '输出到&PDF…

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


