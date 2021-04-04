Attribute VB_Name = "mdlPathology"
Option Explicit


'病理系统编号
Public Const G_LNG_PATHOLSYS_NUM = 1294

'病理归档模块编号
Public Const G_LNG_PATHOLARCHIVES_NUM = 1295

'病理借阅模块编号
Public Const G_LNG_PATHOLBORROW_NUM = 1296

'病理材料遗失管理模块编号
Public Const G_LNG_PATHOLLOSE_NUM = 1297


'图标id常量
Public Const G_INT_ICONID_SPECIMEN = 10015
Public Const G_INT_ICONID_MATERIAL = 10016
Public Const G_INT_ICONID_SLICES = 10017
Public Const G_INT_ICONID_SPEEXAM = 10018
Public Const G_INT_ICONID_PROREPORT = 10019
Public Const G_INT_ICONID_SLICESSURE = 8133
Public Const G_INT_ICONID_BATPROCESS = 802


'病理检查类型定义
Public Enum StudyType
    stNormal = 0  '常规
    stIce = 1     '冰冻
    stCell = 2    '细胞
    stMeet = 3    '会诊
    stAutopsy = 4 '尸检
    stSpeed = 5   '快速石蜡
End Enum



'病理检查过程定义
Public Enum TStudyProcedure
    spReserve = 0
    spMaterial = 1
    spSlices = 2
    spDiagnose = 3
    spMianyi = 4
    spTeran = 5
    spFenzi = 6
    spAgainMaterial = 8
    spAgainSlices = 9
    spFinal = 10
End Enum


'病理执行步骤定义
Public Enum TExecuteStep
    None = 0        '未执行
    NeedDo = 1      '需执行
    AcceptDo = 2    '已接受
    AlreadDo = 3    '已执行
End Enum


'申请类型
Public Enum TRequestType
    rtMianyi = 0    '免疫组化申请
    rtTeran = 1     '特殊染色申请
    rtFenzi = 2     '分子病理申请
    rtSlices = 3    '再制片申请
    rtMaterial = 4  '再取材申请
End Enum


'特检类型定义
Public Enum TSpeexamType
    stMianyi = 0    '免疫组化
    stTeshu = 1     '特殊染色
    stFenzi = 2     '分子病理
End Enum

'病理检查状态信息
Public Type TStudyStateInf
    lngPatholAdviceId As Long   '病理医嘱ID
    lngStudyType As Long        '检查类型
    lngMaterialStep As Long     '取材过程
    lngSlicesStep As Long       '制片过程
    lngMianYiStep As Long       '免疫过程
    lngFenZiStep As Long        '分子过程
    lngTeRanStep As Long        '特染过程
    strPatholNumber As String   '病理号
End Type


'标准行数量
Public Const glngStandardRowCount As Long = 51
'标准行高度
Public Const glngStandardRowHeight As Long = 300
'制片行数（特检或制片）
Public Const glngSlicesRowCount As Long = 51
'最大行数量
Public Const glngMaxRowCount As Long = 101



'完整的日期时间格式字符串
Public Const gstrFullDateTimeFormat = "yyyy-mm-dd hh:mm:ss"

'日期格式字符串
Public Const gstrDateFormat = "yyyy-mm-dd"

'时间格式字符串
Public Const gstrTimeFormat = "hh:mm:ss"

'数据有误的单元格颜色
Public Const gCellErrColor As Long = &HC0C0FF



'列定义格式为：列名称,是否隐藏(默认不隐藏),可否编辑(默认可编辑),是否Button按钮(默认不是),宽度
'
'如果列名称为“≡”则表示该列为扩展列，主要用于控制行的高度
'
'列属性如下：
'显示名称>字段名称
'hide：表示隐藏
'btn：表示该列有button按钮
'read：表示该列为只读
'merge：表示该列为合并列（行合并）
'check：表示是否有checkbox控件
'w1600：表示宽度为1600
'key:表示为关键字段
'□：如果列名称为“□”则表示该列为单独的CheckBox列，不包含数据
'fulldatetime：yyyy-mm-dd hh:mm:ss
'onlydate：yyyy-mm-dd
'onlytime：hh:mm:ss
'shortdatetime：yyyy-mm-dd hh:mm
'cbx<0-否,1-是,2-未设置>：表示该列为可选列
'Align<8,0>：补足位数对齐
'colleft,colcenter,colright：表示列的对齐方式
'txtleft,txtcenter,txtright：表示文本的对齐方式
'chkleft,chkcenter,chkright：表示check的对齐方式
'tdate：表示时间类型
'tnum：表示数字类型
'tstr：表示字符串类型
'uncfg：表示不允许配置隐藏


'列转换属性如下：
'如特检类型:0-免疫组化,1-特殊染色,2-分子病理,els-其他|当前状态:0-未处理,1-已接受,2-已完成|清单状态:0-<nocheck>未打印,1-<check>已打印
'<nocheck>表示该数据显示时，单元格会添加未选中的勾选框，
'<check>表示该数据显示时，单元格会添加已选中的勾选框，
'els:表示当条件都不满足时，取该值

'============================================================================================================================

Public Const gstrPatholCol_ID           As String = "ID"
Public Const gstrPatholCol_病理号       As String = "病理号"
Public Const gstrPatholCol_姓名         As String = "姓名"
Public Const gstrPatholCol_性别         As String = "性别"
Public Const gstrPatholCol_材块号       As String = "材块号"
Public Const gstrPatholCol_标本名称     As String = "标本名称"
Public Const gstrPatholCol_取材位置     As String = "取材位置"
Public Const gstrPatholCol_材料类别     As String = "材料类别"
Public Const gstrPatholCol_材料明细     As String = "材料明细"
Public Const gstrPatholCol_遗失原因     As String = "遗失原因"
Public Const gstrPatholCol_档案状态     As String = "档案状态"
Public Const gstrPatholCol_借阅状态     As String = "借阅状态"
Public Const gstrPatholCol_借阅号       As String = "借阅号"
Public Const gstrPatholCol_借阅人       As String = "借阅人"
Public Const gstrPatholCol_借阅日期     As String = "借阅日期"
Public Const gstrPatholCol_归还日期     As String = "归还日期"
Public Const gstrPatholCol_证件类型     As String = "证件类型"
Public Const gstrPatholCol_证件号码     As String = "证件号码"
Public Const gstrPatholCol_联系电话     As String = "联系电话"
Public Const gstrPatholCol_联系地址     As String = "联系地址"
Public Const gstrPatholCol_押金         As String = "押金"
Public Const gstrPatholCol_借阅类型     As String = "借阅类型"
Public Const gstrPatholCol_借阅天数     As String = "借阅天数"
Public Const gstrPatholCol_借阅原因     As String = "借阅原因"
Public Const gstrPatholCol_归还状态     As String = "归还状态"
Public Const gstrPatholCol_备注         As String = "备注"
Public Const gstrPatholCol_确认状态     As String = "确认状态"
Public Const gstrPatholCol_检查类型     As String = "检查类型"
Public Const gstrPatholCol_归档ID       As String = "归档ID"
Public Const gstrPatholCol_所属档案     As String = "所属档案"
Public Const gstrPatholCol_存放位置     As String = "存放位置"
Public Const gstrPatholCol_详细地址     As String = "详细地址"
Public Const gstrPatholCol_归还人       As String = "归还人"
Public Const gstrPatholCol_退还押金     As String = "退还押金"
Public Const gstrPatholCol_外诊医院     As String = "外诊医院"
Public Const gstrPatholCol_外诊医师     As String = "外诊医师"
Public Const gstrPatholCol_外诊意见     As String = "外诊意见"
Public Const gstrPatholCol_登记人       As String = "登记人"
Public Const gstrPatholCol_在档数量     As String = "在档数量"
Public Const gstrPatholCol_遗失数量     As String = "遗失数量"
Public Const gstrPatholCol_可借数量     As String = "可借数量"
Public Const gstrPatholCol_已借数量     As String = "已借数量"
Public Const gstrPatholCol_需借数量     As String = "需借数量"
Public Const gstrPatholCol_待还数量     As String = "待还数量"
Public Const gstrPatholCol_实还数量     As String = "实还数量"
Public Const gstrPatholCol_借阅数量     As String = "借阅数量"
Public Const gstrPatholCol_归还数量     As String = "归还数量"

Public Const gstrPatholCol_档案名称    As String = "档案名称"
Public Const gstrPatholCol_档案编号    As String = "档案编号"
Public Const gstrPatholCol_档案分类    As String = "档案分类"
Public Const gstrPatholCol_材料类型    As String = "材料类型"
Public Const gstrPatholCol_报表名称    As String = "报表名称"
Public Const gstrPatholCol_检查范围    As String = "检查范围"
Public Const gstrPatholCol_开始日期    As String = "开始日期"
Public Const gstrPatholCol_结束日期    As String = "结束日期"
Public Const gstrPatholCol_所属房间    As String = "所属房间"
Public Const gstrPatholCol_所属柜号    As String = "所属柜号"
Public Const gstrPatholCol_所属抽屉    As String = "所属抽屉"
Public Const gstrPatholCol_档案说明    As String = "档案说明"
Public Const gstrPatholCol_归档时间    As String = "归档时间"
Public Const gstrPatholCol_创建人      As String = "创建人"
Public Const gstrPatholCol_创建日期    As String = "创建日期"

Public Const gstrPatholCol_来源ID        As String = "来源ID"
Public Const gstrPatholCol_病理医嘱ID    As String = "病理医嘱ID"
Public Const gstrPatholCol_档案来源      As String = "档案来源"
Public Const gstrPatholCol_年龄          As String = "年龄"
Public Const gstrPatholCol_检查项目      As String = "检查项目"
Public Const gstrPatholCol_数量          As String = "数量"
Public Const gstrPatholCol_归档状态      As String = "归档状态"
Public Const gstrPatholCol_存放状态      As String = "存放状态"
Public Const gstrPatholCol_报到时间      As String = "报到时间"
Public Const gstrPatholCol_执行过程      As String = "执行过程"

Public Const gstrArchivesClass_ID           As String = "ID"
Public Const gstrArchivesClass_分类名称     As String = "分类名称"
Public Const gstrArchivesClass_材料类型     As String = "材料类型"
Public Const gstrArchivesClass_报表名称     As String = "报表名称"
Public Const gstrArchivesClass_创建人       As String = "创建人"
Public Const gstrArchivesClass_创建时间     As String = "创建时间"
Public Const gstrArchivesClass_备注         As String = "备注"



'材料遗失显示列
Public Const gstrMaterialLoseCols As String = "|材料类别,read,merge|病理号,merge,read,uncfg|ID,key,hide,uncfg|材块号>序号,read,uncfg|标本名称,read|取材位置,read|材料明细,read,uncfg|在档数量,read|遗失数量,read|遗失原因,read|存放状态,read|"
Public Const gstrMaterialLoseConvertFormat As String = "存放状态:0-存档中,1-部分遗失,2-已遗失|借阅状态:0-未借出,1-部分借出,2-已借出"



'借阅管理显示列
Public Const gstrMaterialBorrowCols As String = "|ID,hide,uncfg,key|借阅号,read,uncfg|借阅人,read|借阅日期>借阅时间,read,w1600,onlydate|归还日期,read,w1600,onlydate|证件类型,read|证件号码,read|联系电话,read|联系地址,read|押金,read|借阅类型,read|借阅天数,read|借阅原因,read|归还状态,read|确认状态,read|备注,read|"
Public Const gstrMaterialBorrowConvertFormat As String = "证件类型:0-身份证,1-学生证,2-军官证,3-驾驶证,4-护照,5-社保卡,6-残疾证,7-其他|借阅类型:0-内部借阅,1-外部借阅|归还状态:0-未归还,1-已归还,2-部分归还,3-遗失处理|确认状态:0-未确认,1-已确认"



'借阅材料明细列
Public Const gstrMaterialBorrowDetailCols As String = "|检查类型,merge,read|病理号,read,merge|归档ID,key,hide,uncfg|材块号>序号,read,uncfg|标本名称,read|取材位置,read|材料类别,read|材料明细,read,uncfg|借阅数量,read|归还数量,read|所属档案>档案名称,read|存放位置,read|详细地址,read|归还状态,read|"
Public Const gstrMaterialBorrowDetailConvertFormat As String = "检查类型:0-常规,1-冰冻,2-细胞,3-会诊,4-尸检,5-快速石蜡|归还状态:0-未归还,1-已归还,2-部分归还,3-遗失,4-待借出"



'借阅材料归还显示列
Public Const gstrMaterialBorrowReturnCols As String = "|检查类型,merge,read|病理号,read,merge|归档ID,key,hide,uncfg|材块号>序号,rowcheck,uncfg|实还数量|待还数量,read|标本名称,read|取材位置,read|材料类别,read|材料明细,read,uncfg|所属档案>档案名称,read|存放位置,read|详细地址,read|归还状态,read|"
Public Const gstrMaterialBorrowReturnConvertFormat As String = "检查类型:0-常规,1-冰冻,2-细胞,3-会诊,4-尸检,5-快速石蜡|归还状态:0-未归还,1-已归还,2-部分归还,3-遗失"



'借阅归还历史列
Public Const gstrMaterialBorrowBackCols As String = "|id,key,hide,uncfg|归还人,read,uncfg|归还日期,read,w1600,onlydate|退还押金,read|外诊医院,read|外诊医师,read|外诊意见,read,uncfg|登记人,read|备注,read|"
Public Const gstrMaterialBorrowBackConvertFormat As String = ""



'借阅材料显示列（借阅登记窗口） 姓名,merge,read|性别,merge,read|
Public Const gstrMaterialBorrowEnregCols As String = "|检查类型,merge,read|病理号,read,merge|姓名,merge,read|id,key,read,uncfg,hide|材块号>序号,rowcheck,uncfg|需借数量|可借数量,read|标本名称,read|取材位置,read|材料类别,read|材料明细,read|所属档案>档案名称,read|存放位置,read,w2400|详细地址,read,w1600|遗失数量,read|已借数量,read|存放状态,read|借阅状态,read|"
Public Const gstrMaterialBorrowEnregConvertFormat As String = "检查类型:0-常规,1-冰冻,2-细胞,3-会诊,4-尸检,5-快速石蜡|存放状态:0-存档中,1-部分遗失,2-已遗失|借阅状态:0-未借出,1-部分借出,2-已借出"

Public Const gstrMaterialBorrowEnregedCols As String = "|检查类型,merge,read|病理号,read,merge|id,key,read,uncfg,hide|材块号>序号,rowcheck,uncfg|标本名称,read|取材位置,read|材料类别,read|材料明细,read|借阅数量,read|所属档案>档案名称,read|存放位置,read|详细地址,read,w1600|"



'档案管理显示列
Public Const gstrArchivesManageCols As String = "|ID,hide,uncfg,key|档案名称,read,uncfg|档案编号,read|档案分类,read|材料类型,hide,read,uncfg|报表名称,hide,read,uncfg|检查范围,read|开始日期,read,onlydate,w1600|结束日期,read,onlydate,w1600|所属房间,read|所属柜号,read|所属抽屉,read|详细地址,read,w2400|档案说明,read|档案状态,read|归档时间,read,onlydate,w1600|创建人,read|创建日期,read,onlydate,w1600|"
Public Const gstrArchivesManageConvertFormat As String = "档案状态:0-未归档,1-已归档"



'档案材料类查询细显示列定义
Public Const gstrArchivesMaterialCols As String = "|材料类别,merge,read|检查类型,merge,read|病理号,merge,uncfg,read|姓名,merge,read,uncfg|性别,merge,read|年龄,merge,read|" & _
                                                "检查项目,merge,read|病理医嘱id,hide,uncfg|材块号>序号,rowcheck,uncfg|标本名称,read|取材位置,read|材料明细,read|" & _
                                                "数量,read|所属档案>档案名称,read|存放位置,read,w2400|详细地址,read,w1600|存放状态,read|报到时间,read,onlydate,w1600|执行过程,hide,uncfg|来源ID,key,hide,uncfg|档案来源,hide,uncfg|"
Public Const gstrArchivesMaterialConvertFormat As String = "检查类型:0-常规,1-冰冻,2-细胞,3-会诊,4-尸检,5-快速石蜡"

'档案材料明细显示列定义
Public Const gstrArchivesMaterialDetailCols As String = "|材料类别,merge,read|检查类型,merge,read|病理号,merge,uncfg,read|姓名,merge,read,uncfg|性别,merge,read|年龄,merge,read|" & _
                                                "检查项目,merge,read|病理医嘱id,hide,uncfg|材块号>序号,rowcheck,uncfg|标本名称,read|取材位置,read|材料明细,read|" & _
                                                "数量,read|存放状态,read|借阅状态,read|报到时间,read,onlydate,w1600|执行过程,hide,uncfg|来源ID,key,hide,uncfg|档案来源,hide,uncfg|"


'文字类档案明细显示列定义
Public Const gstrArchivesWordCols As String = "|来源ID,key,hide,uncfg|档案来源,hide,uncfg|病理号,rowcheck,uncfg|姓名,read,uncfg|性别,read|年龄,read|" & _
                                                "检查项目,read|检查类型,read|报到时间,read,onlydate,w1600|执行过程,hide,uncfg|病理医嘱id,hide,uncfg|存放状态,read,uncfg|"
Public Const gstrArchivesWordConvertFormat As String = "检查类型:0-常规,1-冰冻,2-细胞,3-会诊,4-尸检,5-快速石蜡"



'档案分类设置
Public Const gstrArchivesClassCols As String = "|ID,hide,key,uncfg|分类名称,uncfg,w1800|材料类型,cbx<0-表格材料,1-检查材料,2-报告材料>|报表名称, w1600|备注,w2400|创建人,w1800,read|创建时间,onlydate,w1800,read|"
Public Const gstrArchivesClassConvertFormat As String = "材料类型:0-表格材料,1-检查材料,2-报告材料"



'检查标本配置列表定义
Public Const gstrSpecimenModuleCols As String = "|ID,hide,key|标本名称,w1800|标本部位|标本类型,cbx<0-手术标本,1-穿刺细胞,2-脱落细胞,3-液基细胞>,w1800|默认标本量,w1800|默认制片数,w1800|简码|备注,w2400|"
Public Const gstrSpecimenModuleConvertFormat As String = "标本类型:0-手术标本,1-穿刺细胞,2-脱落细胞,3-液基细胞"

Public Const gstrSpecimenModule_ID         As String = "ID"
Public Const gstrSpecimenModule_标本名称   As String = "标本名称"
Public Const gstrSpecimenModule_标本部位   As String = "标本部位"
Public Const gstrSpecimenModule_标本类型   As String = "标本类型"
Public Const gstrSpecimenModule_默认标本量 As String = "默认标本量"
Public Const gstrSpecimenModule_默认制片数 As String = "默认制片数"
Public Const gstrSpecimenModule_简码       As String = "简码"
Public Const gstrSpecimenModule_备注       As String = "备注"



'标本显示定义
'Public Const gstrSpecimenCols As String = "|标本ID,hide,key|送检ID,hide|标本名称|标本类型,cbx<0-手术标本,1-小标本,2-穿刺细胞,3-脱落细胞,4-液基细胞>|采集部位|标本份数>数量|材料类别,cbx<0-标本,1-蜡块,2-玻片,3-白片,4-其他>|" & _
'                                        "存放位置|原有编号|备注|接收日期,fulldatetime,read,w2400|核收状态,read|"
'Public Const gstrSpecimenConvertFormat As String = "标本类型:0-手术标本,1-小标本,2-穿刺细胞,3-脱落细胞,4-液基细胞|材料类别:0-标本,1-蜡块,2-玻片,3-白片,4-其他"

Public Const gstrSpecimenCols As String = "|标本ID,hide,key,uncfg|送检ID,hide,uncfg|标本名称,uncfg|标本类型,cbx<0-手术标本,1-穿刺细胞,2-脱落细胞,3-液基细胞>|采集部位|标本份数>数量|材料类别,cbx<0-标本,1-蜡块,2-玻片,3-白片,4-其他>|" & _
                                        "存放位置|原有编号|备注|接收日期,fulldatetime,read,w2400|核收状态,read|"
                                        
Public Const gstrSpecimenConvertFormat As String = "标本类型:0-手术标本,1-穿刺细胞,2-脱落细胞,3-液基细胞|材料类别:0-标本,1-蜡块,2-玻片,3-白片,4-其他"


Public Const gSpecimen_标本ID   As String = "标本ID"
Public Const gSpecimen_送检ID   As String = "送检ID"
Public Const gSpecimen_标本名称 As String = "标本名称"
Public Const gSpecimen_标本类型 As String = "标本类型"
Public Const gSpecimen_采集部位 As String = "采集部位"
Public Const gSpecimen_数量     As String = "标本份数"
Public Const gSpecimen_材料类别 As String = "材料类别"
Public Const gSpecimen_存放位置 As String = "存放位置"
Public Const gSpecimen_原有编号 As String = "原有编号"
Public Const gSpecimen_接收日期 As String = "接收日期"
Public Const gSpecimen_备注     As String = "备注"
Public Const gSpecimen_核收状态 As String = "核收状态"



'检查质量定义
Public Const gstrPatholQualityCols As String = "|ID,hide,key,uncfg|评价项目,cbx< ,标本质量,制片质量,诊断质量,免疫组化,特殊染色,分子病理>,uncfg|评价结果,cbx< ,甲,乙,丙,丁>,uncfg|评价意见,w2000|改进方法,w2000|备注,w2000|评价人,read|评价时间,w1900,fulldatetime,read|"
                                        
Public Const gstrPatholQualityConvertFormat As String = ""


Public Const gstrPatholQuality_ID       As String = "ID"
Public Const gstrPatholQuality_病理号   As String = "病理号"
Public Const gstrPatholQuality_评价项目 As String = "评价项目"
Public Const gstrPatholQuality_评价结果 As String = "评价结果"
Public Const gstrPatholQuality_评价意见 As String = "评价意见"
Public Const gstrPatholQuality_改进方法 As String = "改进方法"
Public Const gstrPatholQuality_备注     As String = "备注"
Public Const gstrPatholQuality_评价人   As String = "评价人"
Public Const gstrpatholQuality_评价时间 As String = "评价时间"


'============================================================================================================================



'常规取材显示列
Public Const gstrNormalMaterialCols As String = "|材块ID,key,read,w1000,hide,uncfg|材块号>序号,read,w800,align<2,0>|取材位置,w2400|标本名称,uncfg|形状|蜡块数|制片数|是否蜡块,cbx<0-否 ,1-是>,uncfg|是否脱钙,cbx<0-否 ,1-是>|" & _
                                                "主取医师,uncfg|副取医师|取材时间,fulldatetime,w2400,uncfg|记录医师,read|取材类型,read|确认状态,read|"

'细胞取材显示列
Public Const gstrCellMaterialCols As String = "|材块ID,key,,read,w1000,hide,uncfg|材块号>序号,read,w800,align<2,0>|标本名称,uncfg|性质|颜色|标本量,cbx<, ml, 1张, 2张, 4张, 8张>,uncfg|细胞块数>蜡块数|是否蜡块,cbx<0-否 ,1-是>,uncfg|制片数|" & _
                                                "主取医师,uncfg|副取医师|取材时间,fulldatetime,w2400,uncfg|记录医师,read|取材类型,read|确认状态,read|"

'冰冻取材显示列
Public Const gstrIceMaterialCols As String = "|材块ID,key,,read,w1000,hide,uncfg|材块号>序号,read,w800,align<2,0>|取材位置,w2400|标本名称,uncfg|形状|是否蜡块,cbx<0-否 ,1-是>|是否冰余,cbx<0-否,1-是>,uncfg|蜡块数|制片数|" & _
                                                "主取医师,uncfg|副取医师|取材时间,fulldatetime,w2400,uncfg|记录医师,read|取材类型,read|确认状态,read|"
                                                
Public Const gstrMaterialConvertFormat As String = "是否冰余:0-否,1-是|是否脱钙:0-否 ,1-是|确认状态:0-未确认,1-已确认|是否蜡块:0-否,1-是"
                                                
                                                
Public Const gstrMaterial_材块ID   As String = "材块ID"
Public Const gstrMaterial_材块号     As String = "材块号"
Public Const gstrMaterial_标本名称 As String = "标本名称"
Public Const gstrMaterial_取材位置 As String = "取材位置"
Public Const gstrMaterial_形状     As String = "形状"
Public Const gstrMaterial_蜡块数   As String = "蜡块数"
Public Const gstrMaterial_制片数   As String = "制片数"
Public Const gstrMaterial_取材类型 As String = "取材类型"
Public Const gstrMaterial_取材时间 As String = "取材时间"
Public Const gstrMaterial_主取医师 As String = "主取医师"
Public Const gstrMaterial_副取医师 As String = "副取医师"
Public Const gstrMaterial_记录医师 As String = "记录医师"
Public Const gstrMaterial_性质     As String = "性质"
Public Const gstrMaterial_颜色     As String = "颜色"
Public Const gstrMaterial_标本量   As String = "标本量"
Public Const gstrMaterial_细胞块数 As String = "细胞块数"
Public Const gstrMaterial_是否冰余 As String = "是否冰余"
Public Const gstrMaterial_是否脱钙 As String = "是否脱钙"
Public Const gstrMaterial_是否蜡块 As String = "是否蜡块"
Public Const gstrMaterial_确认状态 As String = "确认状态"

                                                
'============================================================================================================================


'脱钙显示列
Public Const gstrDecalinCols As String = "|ID,key,hide,uncfg|标本名称,uncfg|开始时间,fulldatetime,w2400,uncfg|所需时长(小时)>所需时长,w1600,uncfg|操作员,uncfg|结束时间,fulldatetime,read,w2400|当前缸次,read|当前状态>完成状态,read|"
Public Const gstrDecalinTaskCols As String = "|ID,key,hide,uncfg|病理号|标本名称,uncfg|开始时间,fulldatetime,w2400,uncfg|所需时长(小时)>所需时长,w1600|剩余时长(分)>剩余时长,w1200|结束时间,fulldatetime,read,w2400,uncfg|当前缸次,read|操作员,read,uncfg|当前状态>完成状态,read|"

Public Const gstrDecalinConvertFormat As String = "当前状态:0-进行中,1-已完成"

Public Const gstrDecalin_ID       As String = "ID"
Public Const gstrDecalin_标本ID   As String = "标本ID"
Public Const gstrDecalin_标本名称 As String = "标本名称"
Public Const gstrDecalin_开始时间 As String = "开始时间"
Public Const gstrDecalin_所需时长 As String = "所需时长(小时)"
Public Const gstrDecalin_剩余时长 As String = "剩余时长(分)"
Public Const gstrDecalin_结束时间 As String = "结束时间"
Public Const gstrDecalin_当前缸次 As String = "当前缸次"
Public Const gstrDecalin_操作员   As String = "操作员"
Public Const gstrDecalin_当前状态 As String = "当前状态"

 
'============================================================================================================================



'制片显示列
Public Const gstrSlicesCols As String = "|制片ID>ID,key,w1000,hide,uncfg|材块ID,hide,uncfg|材块号>序号,w1000,align<2,0>,uncfg|取材位置|标本名称,uncfg|制片数|制片类型|制片方式|制片时间,fulldatetime,w2400|制片技师|当前状态|清单状态|"
Public Const gstrSlicesConvertFormat As String = "制片类型:0-石蜡制片,1-冰冻切片,2-细胞制片|制片方式:0-常规,1-重切,2-深切,3-连切,4-白片,5-重染,6-薄片|当前状态:0-未处理,1-已接受,2-已完成|清单状态:0-未打印,1-已打印"


Public Const gstrSlices_材块ID     As String = "材块ID"
Public Const gstrSlices_材块号     As String = "材块号"
Public Const gstrSlices_标本名称   As String = "标本名称"
Public Const gstrSlices_制片数     As String = "制片数"
Public Const gstrSlices_制片类型   As String = "制片类型"
Public Const gstrSlices_制片时间   As String = "制片时间"
Public Const gstrSlices_制片人     As String = "制片技师"
Public Const gstrSlices_当前状态   As String = "当前状态"
Public Const gstrSlices_清单状态   As String = "清单状态"


'制片质量显示列
Public Const gstrSlicesQualityCols As String = "|条码号,w1200,read|玻片类型,w1000,read|标本名称,w1500,read|取材位置,w1500,read|材块号,w800,read|ID,key,w1000,hide|来源类型,hide|来源Id,hide|材块ID,hide|玻片质量,w1000,cbx<,甲,乙,丙,丁>|评审人,w1100,read|评审日期,read,onlydate|"
Public Const gstrSlicesQualityConvertFormat As String = ""


Public Const gstrSlicesQuality_制片ID     As String = "制片ID"
Public Const gstrSlicesQuality_材块ID     As String = "材块ID"
Public Const gstrSlicesQuality_标本名称   As String = "标本名称"
Public Const gstrSlicesQuality_制片方式   As String = "制片方式"
Public Const gstrSlicesQuality_玻片序号   As String = "玻片序号"
Public Const gstrSlicesQuality_制片质量   As String = "制片质量"
Public Const gstrSlicesQuality_备注       As String = "备注"
Public Const gstrSlicesQuality_评审人     As String = "评审人"
Public Const gstrSlicesQuality_评审时间   As String = "评审时间"



'制片工作清单显示列
Public Const gstrSlicesWorkCols As String = "|病理号,rowcheck,merge,w1600,uncfg|病理医嘱ID,hide,uncfg|检查类型,merge|姓名,merge|制片ID>ID,key,w1000,hide,uncfg|材块ID,hide,uncfg|材块号>序号,w1000,align<2,0>,uncfg|取材位置|标本名称,w1600,uncfg|标本类型|制片类型|制片方式|制片数|取材时间,fulldatetime,w2400,uncfg|当前状态|清单状态|"
Public Const gstrSlicesWorkConvertFormat As String = "病理号:els-<check><source>|检查类型:0-常规,1-冰冻,2-细胞,3-会诊,4-尸检,5-快速石蜡|标本类型:0-根治标本,1-小标本,2-穿刺细胞,3-脱落细胞,4-液基细胞|制片类型:0-石蜡制片,1-冰冻切片,2-细胞制片|制片方式:0-常规,1-重切,2-深切,3-连切,4-白片,5-重染,6-薄片|当前状态:0-未处理,1-已接受,2-已完成|清单状态:0-未打印,1-已打印"



Public Const gstrSlicesWork_病理号    As String = "病理号"
Public Const gstrSlicesWork_病理医嘱ID    As String = "病理医嘱ID"
Public Const gstrSlicesWork_姓名      As String = "姓名"
Public Const gstrSlicesWork_检查类型  As String = "检查类型"
Public Const gstrSlicesWork_材块ID    As String = "材块ID"
Public Const gstrSlicesWork_材块号  As String = "材块号"
Public Const gstrSlicesWork_标本名称  As String = "标本名称"
Public Const gstrSlicesWork_标本类型  As String = "标本类型"
Public Const gstrSlicesWork_制片类型  As String = "制片类型"
Public Const gstrSlicesWork_制片数    As String = "制片数"
Public Const gstrSlicesWork_当前状态  As String = "当前状态"
Public Const gstrSlicesWork_清单状态  As String = "清单状态"




'制片确认显示列
Public Const gstrSlicesSureColsWithMaterialNum As String = "|病理号,rowcheck,merge,w1600|姓名,merge,read|检查类型,merge,read|制片Id>ID,read,w1000,hide|材块ID,hide|材块号>序号,w1000,read,align<2,0>|标本名称,read|制片类型,read|制片方式,read|需制片数,read|已确认数|当前状态,read|病理医嘱ID,key,hide|"
Public Const gstrSlicesSureConvertFormat = "病理号:els-<check><source>|检查类型:0-常规,1-冰冻,2-细胞,3-会诊,4-尸检,5-快速石蜡|制片类型:0-石蜡制片,1-冰冻切片,2-细胞制片|制片方式:0-常规,1-重切,2-深切,3-连切,4-白片,5-重染,6-薄片|当前状态:0-未处理,1-已接受,2-已完成"


Public Const gstrSlicesSure_ID         As String = "制片ID"
Public Const gstrSlicesSure_病理号     As String = "病理号"
Public Const gstrSlicesSure_姓名       As String = "姓名"
Public Const gstrSlicesSure_检查类型   As String = "检查类型"
Public Const gstrSlicesSure_制片状态   As String = "制片状态"
Public Const gstrSlicesSure_材块ID     As String = "材块ID"
Public Const gstrSlicesSure_材块号   As String = "材块号"
Public Const gstrSlicesSure_标本名称   As String = "标本名称"
Public Const gstrSlicesSure_当前状态   As String = "当前状态"
Public Const gstrSlicesSure_需制片数   As String = "需制片数"
Public Const gstrSlicesSure_已确认数   As String = "已确认数"
Public Const gstrSlicesSure_确认状态   As String = "确认状态"



'============================================================================================================================

'抗体信息显示列
Public Const gstrAntibodyCols As String = "|抗体ID,key,hide,uncfg|抗体名称,uncfg|使用人份|已用人份|生产日期,onlydate,w1600|有效期|过期日期,onlydate,w1600|克隆性|作用对象|理化性质|应用情况|使用状态|登记人,uncfg|登记时间,fulldatetime,w2400,uncfg|备注|"
Public Const gstrAntibodyConvertFormat As String = "克隆性:0-单克隆（浓缩型）,1-单克隆（即用型）,2-多克隆（浓缩型）,3-多克隆（即用型）|使用状态:0-已停止,1-使用中"



Public Const gstrAntibody_抗体ID   As String = "抗体ID"
Public Const gstrAntibody_抗体名称 As String = "抗体名称"
Public Const gstrAntibody_使用人份 As String = "使用人份"
Public Const gstrAntibody_已用人份 As String = "已用人份"
Public Const gstrAntibody_生产日期 As String = "生产日期"
Public Const gstrAntibody_有效期   As String = "有效期"
Public Const gstrAntibody_过期日期 As String = "过期日期"
Public Const gstrAntibody_克隆性   As String = "克隆性"
Public Const gstrAntibody_作用对象 As String = "作用对象"
Public Const gstrAntibody_理化性质 As String = "理化性质"
Public Const gstrAntibody_应用情况 As String = "应用情况"
Public Const gstrAntibody_使用状态 As String = "使用状态"
Public Const gstrAntibody_登记人   As String = "登记人"
Public Const gstrAntibody_登记时间 As String = "登记时间"
Public Const gstrAntibody_备注     As String = "备注"
        
        
'抗体反馈信息显示列
Public Const gstrAntibodyFeedbackCols As String = "|ID,key,hide,uncfg|参考病理号,w2400|实验类型|抗体评价|反馈意见,w3200,uncfg|反馈医生,uncfg|反馈时间,fulldatetime,w2400,uncfg|"
Public Const gstrAntibodyFeedbackConvertFormat As String = "实验类型:0-免疫组化,1-特殊染色,2-分子病理,3-其他"


Public Const gstrAntibodyFeedback_ID         As String = "ID"
Public Const gstrAntibodyFeedback_参考病理号 As String = "参考病理号"
Public Const gstrAntibodyFeedback_实验类型   As String = "实验类型"
Public Const gstrAntibodyFeedback_抗体评价   As String = "抗体评价"
Public Const gstrAntibodyFeedback_反馈意见   As String = "反馈意见"
Public Const gstrAntibodyFeedback_反馈医生   As String = "反馈医生"
Public Const gstrAntibodyFeedback_反馈时间   As String = "反馈时间"
        
        
        
'============================================================================================================================


'套餐信息显示列
Public Const gstrAntibodyMealCols As String = "|套餐ID,key,hide,uncfg|套餐名称,uncfg|套餐类别|套餐说明,w3200|创建时间,fulldatetime,read,w2400|创建人,read|"
Public Const gstrAntibodyMealConvertFormat As String = ""


Public Const gstrAntibodyMeal_套餐ID   As String = "套餐ID"
Public Const gstrAntibodyMeal_套餐名称 As String = "套餐名称"
Public Const gstrAntibodyMeal_套餐类别 As String = "套餐类别"
Public Const gstrAntibodyMeal_套餐说明 As String = "套餐说明"
Public Const gstrAntibodyMeal_创建时间 As String = "创建时间"
Public Const gstrAntibodyMeal_创建人   As String = "创建人"


'套餐抗体明细显示列
Public Const gstrAntibodyMealLinkCols As String = "|关联ID,hide,uncfg|抗体ID,key,hide,uncfg|抗体名称,rowcheck,uncfg,w1200|克隆性,read,w1700|理化性质,read|作用对象,read|应用情况,read,w2400|备注,read|抗体顺序,hide,uncfg|"
Public Const gstrAntibodyMealLinkConvertFormat As String = "克隆性:0-单克隆（浓缩型）,1-单克隆（即用型）,2-多克隆（浓缩型）,3-多克隆（即用型）"


Public Const gstrAntibodyMealLink_关联ID   As String = "关联ID"
Public Const gstrAntibodyMealLink_抗体ID   As String = "抗体ID"
Public Const gstrAntibodyMealLink_抗体名称 As String = "抗体名称"
Public Const gstrAntibodyMealLink_克隆性   As String = "克隆性"
Public Const gstrAntibodyMealLink_理化性质 As String = "理化性质"
Public Const gstrAntibodyMealLink_作用对象 As String = "作用对象"
Public Const gstrAntibodyMealLink_应用情况 As String = "应用情况"
Public Const gstrAntibodyMealLink_备注     As String = "备注"
Public Const gstrAntibodyMealLink_抗体顺序 As String = "抗体顺序"

'============================================================================================================================


'报告延迟显示列
Public Const gstrReportDelayCols As String = "|ID,key,hide,uncfg|延迟原因,btn,w3200,uncfg|延迟天数,uncfg|临时诊断,w3200|转达人|登记时间,fulldatetime,read,w2400|登记人,read|当前状态,read|"
Public Const gstrReportDelayConvertFormat As String = "当前状态:0-未打印,1-已打印"

Public Const gstrReportDelay_ID       As String = "ID"
Public Const gstrReportDelay_病理号   As String = "病理号"
Public Const gstrReportDelay_延迟原因 As String = "延迟原因"
Public Const gstrReportDelay_延迟天数 As String = "延迟天数"
Public Const gstrReportDelay_临时诊断 As String = "临时诊断"
Public Const gstrReportDelay_转达人   As String = "转达人"
Public Const gstrReportDelay_登记人   As String = "登记人"
Public Const gstrReportDelay_登记时间 As String = "登记时间"
Public Const gstrReportDelay_当前状态 As String = "当前状态"


'============================================================================================================================


'过程报告显示列
Public Const gstrProcedureRepCols As String = "|ID,key,hide,uncfg|报告图像,hide,uncfg|标本名称,uncfg|报告类型,uncfg|报告子项|检查结果,hide,uncfg|检查意见,hide,uncfg|报告人>报告医师,uncfg|报告日期,fulldatetime, w2400|当前状态|备注,hide,uncfg|"
Public Const gstrProcedureRepConvertFormat = "报告类型:0-冰冻报告,1-免疫报告,2-分子报告,3-特染报告|报告子项:0-无,1-鉴别,2-多药耐药,3-荧光,4-普通|当前状态:0-未打印,1-已查阅,2-已撤销,3-已打印"

Public Const gstrProcedureRep_ID       As String = "ID"
Public Const gstrProcedureRep_报告图像 As String = "报告图像"
Public Const gstrProcedureRep_标本名称 As String = "标本名称"
Public Const gstrProcedureRep_报告类型 As String = "报告类型"
Public Const gstrProcedureRep_报告子项 As String = "报告子项"
Public Const gstrProcedureRep_检查结果 As String = "检查结果"
Public Const gstrProcedureRep_检查意见 As String = "检查意见"
Public Const gstrProcedureRep_报告人   As String = "报告人"
Public Const gstrProcedureRep_报告日期 As String = "报告日期"
Public Const gstrProcedureRep_当前状态 As String = "当前状态"
Public Const gstrProcedureRep_备注     As String = "备注"


'============================================================================================================================


'申请显示列
Public Const gstrRequisitionCols As String = "|申请ID,key,hide,uncfg|申请人,uncfg|申请类型,uncfg|补费状态|申请细目|申请时间,fulldatetime,w2400|当前状态>申请状态|申请描述,w3200|完成时间,fulldatetime,w2400|"
Public Const gstrRequisitionViewCols As String = "|申请ID,key,hide,uncfg|申请人,uncfg|申请时间,fulldatetime,w2400|申请细目,uncfg|当前状态>申请状态|申请描述,w3200|补费状态|完成时间,fulldatetime,w2400|"
Public Const gstrRequisitionConvertFormat As String = "申请类型:0-免疫组化,1-特殊染色,2-分子病理,3-再制片,4-补取材|补费状态:0-无,1-需补费,2-已补费|申请细目:0-无,1-鉴别,2-多药耐药,3-荧光,4-普通|当前状态:0-已申请,1-已接受,2-已完成"


Public Const gstrRequisition_申请ID   As String = "申请ID"
Public Const gstrRequisition_申请人   As String = "申请人"
Public Const gstrRequisition_申请类型 As String = "申请类型"
Public Const gstrRequisition_补费状态 As String = "补费状态"
Public Const gstrRequisition_申请细目 As String = "申请细目"
Public Const gstrRequisition_申请时间 As String = "申请时间"
Public Const gstrRequisition_当前状态 As String = "当前状态"
Public Const gstrRequisition_申请描述 As String = "申请描述"


'特检申请内容明细显示列
Public Const gstrRequest_SpeExam_Cols As String = "|ID,key,hide,read,uncfg|抗体ID,hide,read,uncfg|材块ID,hide,read,uncfg|材块号>序号,w1000,read,align<2,0>,uncfg|标本名称,read,w1600,uncfg|抗体名称,btn,read,uncfg|制作类型,read|当前状态,read|项目结果,read|完成时间,fulldatetime,read,w2400|操作人>特检医师,read|"
Public Const gstrRequest_SpeExamConvertFormat As String = "制作类型:-1-补做,0-常规,els-第<source>次重做|当前状态:0-已申请,1-已接受,2-已完成"


Public Const gstrRequest_SpeExam_ID       As String = "ID"
Public Const gstrRequest_SpeExam_材块号 As String = "材块号"
Public Const gstrRequest_SpeExam_标本名称 As String = "标本名称"
Public Const gstrRequest_SpeExam_抗体ID   As String = "抗体ID"
Public Const gstrRequest_SpeExam_抗体名称 As String = "抗体名称"
Public Const gstrRequest_SpeExam_制作类型 As String = "制作类型"
Public Const gstrRequest_SpeExam_当前状态 As String = "当前状态"
Public Const gstrRequest_SpeExam_项目结果 As String = "项目结果"
Public Const gstrRequest_SpeExam_完成时间 As String = "完成时间"
Public Const gstrRequest_SpeExam_操作人   As String = "操作人"


'制片申请内容明细显示列
Public Const gstrRequest_Slices_Cols As String = "|ID,key,hide,read,uncfg|材块ID,hide,read,uncfg|材块号>序号,w1000,read,align<2,0>,uncfg|标本名称,read,uncfg|制片类型,read|制片方式,read|制片数量>制片数,read|当前状态,read|制片时间,fulldatetime,read,w2400|制片人,read|"
Public Const gstrRequest_SlicesConvertFormat As String = "制片类型:0-石蜡制片,1-冰冻切片,2-细胞制片|制片方式:0-常规,1-重切,2-深切,3-连切,4-白片,5-重染,6-薄片|当前状态:0-已申请,1-已接受,2-已完成"


Public Const gstrRequest_Slices_ID       As String = "ID"
Public Const gstrRequest_Slices_材块号   As String = "材块号"
Public Const gstrRequest_Slices_标本名称 As String = "标本名称"
Public Const gstrRequest_Slices_制片类型 As String = "制片类型"
Public Const gstrRequest_Slices_制片方式 As String = "制片方式"
Public Const gstrRequest_Slices_制片数量 As String = "制片数量"
Public Const gstrRequest_Slices_当前状态 As String = "当前状态"
Public Const gstrRequest_Slices_制片时间 As String = "制片时间"
Public Const gstrRequest_Slices_制片人   As String = "制片人"



'补取材申请完成情况显示列
Public Const gstrRequest_Material_Cols As String = "|材块ID,hide,key,read,uncfg|材块号>序号,w1000,read,align<2,0>,uncfg|标本名称,read,uncfg|标本量,read|蜡块数,read|取材时间,fulldatetime,read,w2400|主取医师,read|副取医师,read|记录医师,read|"
Public Const gstrRequest_MaterialConvertFormat As String = ""



Public Const gstrRequest_Material_材块号   As String = "材块号"
Public Const gstrRequest_Material_标本名称 As String = "标本名称"
Public Const gstrRequest_Material_标本量   As String = "标本量"
Public Const gstrRequest_Material_蜡块数   As String = "蜡块数"
Public Const gstrRequest_Material_取材时间 As String = "取材时间"
Public Const gstrRequest_Material_主取医师 As String = "主取医师"
Public Const gstrRequest_Material_副取医师 As String = "副取医师"
Public Const gstrRequest_Material_记录医师 As String = "记录医师"



'特检申请的抗体信息显示列
Public Const gstrRequestAntibodyCols As String = "|抗体ID,key,hide,uncfg|抗体名称,rowcheck,btn,w1600,uncfg|使用人份,read|已用人份,read|生产日期,onlydate,,read,w1600|有效期,read|过期日期,onlydate,read,w1600|项目顺序,hide,uncfg|"
Public Const gstrRequestAntibodyConvertFormat As String = ""

Public Const gstrRequestAntibody_抗体ID   As String = "抗体ID"
Public Const gstrRequestAntibody_抗体名称 As String = "抗体名称"
Public Const gstrRequestAntibody_使用人份 As String = "使用人份"
Public Const gstrRequestAntibody_已用人份 As String = "已用人份"
Public Const gstrRequestAntibody_生产日期 As String = "生产日期"
Public Const gstrRequestAntibody_有效期   As String = "有效期"
Public Const gstrRequestAntibody_过期日期 As String = "过期日期"
Public Const gstrRequestAntibody_项目顺序 As String = "项目顺序"




'============================================================================================================================



'会诊申请内容明细显示列
Public Const gstrConsultationCols As String = "|ID,key,hide,uncfg|申请医师,uncfg|会诊单位|会诊医师,uncfg|会诊类型|会诊时间,shortdatetime,w2400|截止时间,shortdatetime,w2400|初步诊断>检查描述,w2400|诊断结果,w3200,uncfg|诊断意见,w3200,uncfg|当前状态|完成时间,fulldatetime,w2400|备注,w3200|"
Public Const gstrConsultationConvertFormat As String = "会诊类型:0-科内会诊,1-院外会诊|当前状态:0-已申请,1-已撤销,2-已反馈,3-已查阅"


Public Const gstrConsultation_ID       As String = "ID"
Public Const gstrConsultation_申请医师 As String = "申请医师"
Public Const gstrConsultation_会诊单位 As String = "会诊单位"
Public Const gstrConsultation_会诊医师 As String = "会诊医师"
Public Const gstrConsultation_会诊类型 As String = "会诊类型"
Public Const gstrConsultation_会诊时间 As String = "会诊时间"
Public Const gstrConsultation_截止时间 As String = "截止时间"
Public Const gstrConsultation_初步诊断 As String = "初步诊断"
Public Const gstrConsultation_诊断结果 As String = "诊断结果"
Public Const gstrConsultation_诊断意见 As String = "诊断意见"
Public Const gstrConsultation_当前状态 As String = "当前状态"
Public Const gstrConsultation_完成时间 As String = "完成时间"
Public Const gstrConsultation_备注     As String = "备注"





'============================================================================================================================



'特检信息显示列
Public Const gstrSpeExamCols As String = "|特检ID>ID,key,read,w1000,hide,uncfg|材块ID,hide,uncfg|材块号>序号,read,w1000,align<2,0>,uncfg|标本名称,read,uncfg|申请ID,hide,uncfg|抗体ID,hide,uncfg|抗体名称,btn,read,uncfg|特检细目,read|制作类型,read|特检技师,read|申请时间,fulldatetime,read,w2400|完成时间,fulldatetime,read,w2400|当前状态,read|清单状态,read|特检类型,hide,uncfg|"
Public Const gstrSpeExamConvertFormat = "制作类型:-1-补做,0-常规,els-第<source>次重做|当前状态:0-已申请,1-已接受,2-已完成|特检细目:0-无,1-鉴别,2-多药耐药,3-荧光,4-普通|清单状态:0-未打印,1-已打印"


Public Const gstrSpeExam_ID       As String = "ID"
Public Const gstrSpeExam_材块ID   As String = "材块ID"
Public Const gstrSpeExam_材块号 As String = "材块号"
Public Const gstrSpeExam_标本名称 As String = "标本名称"
Public Const gstrSpeExam_申请ID   As String = "申请ID"
Public Const gstrSpeExam_抗体ID   As String = "抗体ID"
Public Const gstrSpeExam_抗体名称 As String = "抗体名称"
Public Const gstrSpeExam_特检细目 As String = "特检细目"
Public Const gstrSpeExam_制作类型 As String = "制作类型"
Public Const gstrSpeExam_当前状态 As String = "当前状态"
Public Const gstrSpeExam_项目结果 As String = "项目结果"
Public Const gstrSpeExam_申请时间 As String = "申请时间"
Public Const gstrSpeExam_完成时间 As String = "完成时间"
Public Const gstrSpeExam_特检医师 As String = "特检技师"
Public Const gstrSpeExam_清单状态 As String = "清单状态"
Public Const gstrSpeExam_特检类型 As String = "特检类型"


'特检工作清单显示列
Public Const gstrSpeExamWorkCols As String = "|病理号,rowcheck,merge,w1600,uncfg|病理医嘱ID,hide,uncfg|检查类型,merge|姓名,merge|特检ID>ID,key,read,w1000,hide,uncfg|材块ID,hide,uncfg|材块号>序号,w1000,align<2,0>,uncfg|标本名称,w1600,uncfg|特检类型|特检细目|抗体ID,hide,uncfg|抗体名称,uncfg|制作类型|当前状态|清单状态|申请时间,fulldatetime,read,w2400|完成时间,fulldatetime,read,w2400|"
Public Const gstrSpeExamWorkConvertFormat = "病理号:els-<check><source>|检查类型:0-常规,1-冰冻,2-细胞,3-会诊,4-尸检,5-快速石蜡|特检类型:0-免疫组化,1-特殊染色,2-分子病理|制作类型:-1-补做,0-常规,els-第<source>次重做|特检细目:0-无,1-鉴别,2-多药耐药,3-荧光,4-普通|当前状态:0-已申请,1-已接受,2-已完成|清单状态:0-未打印,1-已打印"


Public Const gstrSpeExamWork_ID             As String = "ID"
Public Const gstrSpeExamWork_检查类型       As String = "检查类型"
Public Const gstrSpeExamWork_病理号         As String = "病理号"
Public Const gstrSpeExamWork_病理医嘱ID     As String = "病理医嘱ID"
Public Const gstrSpeExamWork_姓名           As String = "姓名"
Public Const gstrSpeExamWork_材块ID         As String = "材块ID"
Public Const gstrSpeExamWork_材块号         As String = "材块号"
Public Const gstrSpeExamWork_标本名称       As String = "标本名称"
Public Const gstrSpeExamWork_特检类型       As String = "特检类型"
Public Const gstrSpeExamWork_抗体ID         As String = "抗体ID"
Public Const gstrSpeExamWork_抗体名称       As String = "抗体名称"
Public Const gstrSpeExamWork_制作类型       As String = "制作类型"
Public Const gstrSpeExamWork_当前状态       As String = "当前状态"
Public Const gstrSpeExamWork_清单状态       As String = "清单状态"
Public Const gstrSpeExamWork_申请时间       As String = "申请时间"
Public Const gstrSpeExamWork_完成时间       As String = "完成时间"


'============================================================================================================================


'特检结果提取显示列
Public Const gstrSpeExamResultGetCols As String = "|ID,Key,hide,uncfg|材块号>序号,rowcheck,w1000,align<2,0>,uncfg|标本名称,read,uncfg|抗体名称,read,uncfg|项目结果,uncfg|特检细目,read|制作类型,read|项目顺序,hide,uncfg|"
Public Const gstrSpeExamResultGetConvertFormat As String = "制作类型:-1-补做,0-常规,els-第<source>次重做|特检细目:0-无,1-鉴别,2-多药耐药,3-荧光,4-普通"


Public Const gstrSpeExamResultGet_材块号 As String = "材块号"
Public Const gstrSpeExamResultGet_标本名称 As String = "标本名称"
Public Const gstrSpeExamResultGet_抗体名称 As String = "抗体名称"
Public Const gstrSpeExamResultGet_项目结果 As String = "项目结果"
Public Const gstrSpeExamResultGet_项目顺序 As String = "项目顺序"



'============================================================================================================================






Public Function GetNumber(ByVal str As String) As Long
'获取字符串中的数字
    Dim strNum As String
    Dim i As Long
        
    For i = 1 To Len(str)
        If IsNumeric(Mid(str, i, 1)) Then
            strNum = strNum & Mid(str, i, 1)
        End If
    Next i
    
    GetNumber = CLng(IIf(strNum = "", -1, strNum))
    
End Function



Public Function CheckPopedom(ByVal strPrivs As String, ByVal strPopedom As String) As Boolean
'检查权限
    Dim strCurPrivs As String
    
    strCurPrivs = ";" & strPrivs & ";"
    
    CheckPopedom = InStr(1, UCase(strCurPrivs), UCase(";" & strPopedom & ";")) > 0
End Function


Public Sub GetPatholStudyState(ByVal lngAdviceID As Long, ByRef recStudy As TStudyStateInf)
'获取病理号等相关状态信息
    Dim strSql As String
    Dim rsPatholNum As ADODB.Recordset
    
    
    strSql = "select 病理医嘱ID,病理号,检查类型,取材过程,制片过程,免疫过程,特染过程,分子过程 from 病理检查信息 where 医嘱id=[1]"
    
    Set rsPatholNum = zlDatabase.OpenSQLRecord(strSql, "获取病理状态信息", lngAdviceID)
    
    If rsPatholNum.RecordCount <= 0 Then
        recStudy.lngPatholAdviceId = -1
        recStudy.lngStudyType = -1
        recStudy.lngMaterialStep = -1
        recStudy.lngSlicesStep = -1
        recStudy.lngMianYiStep = -1
        recStudy.lngFenZiStep = -1
        recStudy.lngTeRanStep = -1
        recStudy.strPatholNumber = ""
        Exit Sub
    End If
    
    recStudy.lngPatholAdviceId = Val(Nvl(rsPatholNum!病理医嘱id))
    recStudy.lngStudyType = Val(Nvl(rsPatholNum!检查类型))
    recStudy.lngMaterialStep = Val(Nvl(rsPatholNum!取材过程))
    recStudy.lngSlicesStep = Val(Nvl(rsPatholNum!制片过程))
    recStudy.lngMianYiStep = Val(Nvl(rsPatholNum!免疫过程))
    recStudy.lngFenZiStep = Val(Nvl(rsPatholNum!分子过程))
    recStudy.lngTeRanStep = Val(Nvl(rsPatholNum!特染过程))
    recStudy.strPatholNumber = Nvl(rsPatholNum!病理号)
End Sub

Public Function GetPatholMenuIndex(objMenuBar As Object) As Long
'获取病理菜单索引
    Dim cbrPathol As CommandBarControl
    
    Set cbrPathol = objMenuBar.FindControl(, conMenu_PatholManage)
    
    If Not cbrPathol Is Nothing Then
        GetPatholMenuIndex = cbrPathol.Index
    Else
        GetPatholMenuIndex = 3
    End If
End Function


Public Function HasMenu(objMenuBar As Object, ByVal lngMenuId As Long) As Boolean
'是否存在指定菜单
    Dim cbrParentMenu As CommandBarControl
    
    Set cbrParentMenu = objMenuBar.FindControl(, lngMenuId)
    
    HasMenu = IIf(cbrParentMenu Is Nothing, False, True)
End Function


Public Function GetHistoryQuerySql(ByVal strSourceSql As String) As String
'取得转存后的数据查询语句
    Dim strNewSql As String
    
    strNewSql = strSourceSql
    
    strNewSql = Replace(strNewSql, "病人医嘱记录", "H病人医嘱记录")
    strNewSql = Replace(strNewSql, "病人医嘱发送", "H病人医嘱发送")
    strNewSql = Replace(strNewSql, "影像检查记录", "H影像检查记录")
    
    strNewSql = Replace(strNewSql, "电子病历记录", "H电子病历记录")
    strNewSql = Replace(strNewSql, "电子病历内容", "H电子病历内容")
    
    
'    病理数据在10.32.0之后取消数据转储
'    strNewSql = Replace(strNewSql, "病理检查信息", "H病理检查信息")
'    strNewSql = Replace(strNewSql, "病理质量信息", "H病理质量信息")
'    strNewSql = Replace(strNewSql, "病理标本信息", "H病理标本信息")
'    strNewSql = Replace(strNewSql, "病理送检信息", "H病理送检信息")
'    strNewSql = Replace(strNewSql, "病理取材信息", "H病理取材信息")
'    strNewSql = Replace(strNewSql, "病理脱钙信息", "H病理脱钙信息")
'    strNewSql = Replace(strNewSql, "病理制片信息", "H病理制片信息")
'    strNewSql = Replace(strNewSql, "病理过程报告", "H病理过程报告")
'    strNewSql = Replace(strNewSql, "病理申请信息", "H病理申请信息")
'    strNewSql = Replace(strNewSql, "病理特检信息", "H病理特检信息")
'    strNewSql = Replace(strNewSql, "病理报告延迟", "H病理报告延迟")
'    strNewSql = Replace(strNewSql, "病理会诊信息", "H病理会诊信息")
'    strNewSql = Replace(strNewSql, "病理归档信息", "H病理归档信息")
  
    
    GetHistoryQuerySql = strNewSql
    
End Function




Public Sub InitDebugObject(ByVal lngModuleNum As Long, ByVal frmMain As Object, ByVal strUser As String, ByVal strPwd As String)
'初始化调试状态下的所需对象
    Set gcnOracle = New ADODB.Connection
    
    Call OraDataOpen("", strUser, strPwd)
    
    glngSys = 100
    gstrPrivs = ";PACS报告打印;PACS报告删除;PACS报告书写;PACS报告他科报告;PACS报告修订;PACS他人报告;采集参数设置;参数设置;存储管理;关联病人;基本;检查报到;检查登记;检查完成;绿色通道;排队叫号;清除图像;取消报到;取消检查完成;删除临时影像;视频采集;随访;所有科室;图像关联;未缴费报到;文件发送;无报告完成;影像质控;档案分类设置;Excel输出;"
    glngModul = lngModuleNum
    
    UserInfo.ID = 281
    UserInfo.姓名 = "张永康"
    UserInfo.用户名 = "ZLHIS"
    UserInfo.编号 = "1123"
    UserInfo.简码 = "WGY"
    UserInfo.部门ID = "65"
    
    
    Call InitCommon(gcnOracle)
    
    Call RegCheck
        
    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '初始化医嘱，病历核心部件
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
End Sub


Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSql As String
    Dim strError As String
    
    On Error Resume Next
    err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '保存错误信息
            strError = err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    err = 0
    On Error GoTo errHand
    
    gstrDBUser = UCase(strUserName)
    SetDbUser gstrDBUser
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    err = 0
End Function
