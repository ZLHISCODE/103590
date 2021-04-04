--参数数据
Insert Into zlParameters(ID,系统,模块,私有,本机,授权,固定,参数号,参数名,参数值,缺省值,参数说明)
Select Rownum+B.ID,A.* From (
  Select 系统,模块,私有,本机,授权,固定,参数号,参数名,参数值,缺省值,参数说明 From zlParameters Where ID=0 Union All
  Select 100,1294,1,0,0,0,40,'常规过滤', '1','1','过滤病理检查类别为常规的检查.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,41,'冰冻过滤', '1','1','过滤病理检查类别为冰冻的检查.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,42,'细胞过滤', '1','1','过滤病理检查类别为细胞的检查.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,43,'会诊过滤', '1','1','过滤病理检查类别为会诊的检查.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,44,'尸检过滤', '1','1','过滤病理检查类别为尸检的检查.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,45,'根治过滤', '1','1','过滤病理标本类型为根治的检查.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,46,'小标本过滤', '1','1','过滤病理标本类型为小标本的检查.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,47,'穿刺过滤', '1','1','过滤病理标本类型为穿刺的检查.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,48,'脱落过滤', '1','1','过滤病理标本类型为脱落的检查.'   From Dual   Union ALL
  Select 100,1294,1,0,0,0,49,'液基过滤', '1','1','过滤病理标本类型为液基的检查.'   From Dual   Union ALL    
  Select 100,1294,1,0,0,0,50,'过滤页面', '0','0','设置当前病理检查数据的过滤页面索引.'   From Dual
  ) A,(Select Nvl(Max(ID),0) AS ID From zlParameters) B;




--添加病理部件程序
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1294,'影像病理工作站','用于病理标本核收和取材、图像采集、报告填写、附加医嘱和费用的登记',100,'zl9PacsWork');


--功能模块
--影像病理工作站 1294
Insert Into zlProgFuncs(系统,序号,功能) Values(100,1294,'基本');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'所有科室',1,'可以查看所有科室PACS检查的权限。');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'检查登记',2,'填写检查登记,取消登记和修改登记。');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'无报告完成',3,'直接完成和回退无报告的检查。');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'文件发送',4,'可发送指定图像文件到共享目录的权限。');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'取消报到',5,'取消影像检查报到状态。');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'清除图像',6,'删除图像、取消关联、重新关联以及Q/R提取图像的权限。');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'检查完成',7,'确认本次检查费用和报告都已经录入完成。');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'取消检查完成',8,'取消本次检查完成的状态。');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'视频采集',9,'可进行视频影像采集的权限。');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'存储管理',10,'可进行在线近线离线图像数据的管理。');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'参数设置',11,'进行参数设定');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'采集参数设置',12,'进行参数设定');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'影像质控',13,'进行影像质量等级控制');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'PACS报告书写',15,'使用PACS报告编辑器书写报告');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'PACS报告修订',16,'使用PACS报告编辑器修订报告');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'PACS报告打印',17,'使用PACS报告编辑器打印报告');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'PACS报告删除',18,'使用PACS报告编辑器允许强制删除报告');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'检查报到',19,'确认报到');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'绿色通道',20,'对某次检查标记/取消绿色通道');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'PACS他人报告',21,'使用PACS报告编辑器,允许可书写报告的人员删改他人书写的报告');
Insert Into zlProgFuncs(系统,序号,功能,排列,说明) Values(100,1294,'排队叫号',22,'对报到的患者进行排队显示和语音呼叫');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'随访',23,'记录病人的随访信息');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'未缴费报到',24,'拥有该权限可以报到未缴费的检查记录');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'关联病人',25,'拥有该权限可以设置关联病人');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'PACS报告他科报告',26,'拥有该权限可以在PACS报告编辑器中，通过历史报告功能查看其他科室的报告');


Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'标本核收',28,'对送检标本进行核收');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'病理取材',29,'获取待检查的材块');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'病理制片',30,'对蜡块进行切片，包括细胞，冰冻，石蜡');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'报告延迟',31,'对需要延迟的病理检查进行登记管理');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'冰冻报告',32,'编辑冰冻过程报告');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'免疫报告',33,'编辑免疫过程报告');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'分子报告',34,'编辑分子过程报告');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'特染报告',35,'编辑特染过程报告');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'特检申请',36,'申请特殊检查');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'制片申请',37,'申请制片');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'补取申请',38,'申请取材');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'会诊申请',39,'申请会诊');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'会诊反馈',40,'反馈会诊结果');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'免疫组化',41,'免疫组化操作');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'特殊染色',42,'特殊染色操作');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'分子病理',43,'分子病理操作');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'抗体管理',44,'管理抗体信息');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'抗体反馈',45,'抗体使用情况反馈');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'套餐维护',46,'维护抗体套餐信息');
Insert Into zlprogfuncs(系统,序号,功能,排列,说明) Values(100,1294,'冰冻特检报告查阅',47,'查阅和撤销冰冻特检报告');


--基本权限
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','人员性质说明',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','部门性质说明',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','性别',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','费别',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','民族',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','职业',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','婚姻状况',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','医疗付款方式',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病案主页',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人医嘱记录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人医嘱发送',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人医嘱报告',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','电子病历记录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','电子病历内容',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病历单据应用',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人医嘱附件',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人医嘱执行',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','收费项目目录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人医嘱附费',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','门诊费用记录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','住院费用记录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','药品规格',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','药品特性',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','药品库存',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','医技执行房间',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗项目别名',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗执行科室',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗项目部位',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗检查部位',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗项目组合',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗用法用量',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗分类目录',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗互斥项目',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗检验标本',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗个人项目',user,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗项目目录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病历文件列表',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病历单据附项',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','号码控制表',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人医嘱记录_ID',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人余额',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','床位状况记录',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人挂号记录',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','检验项目参考',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','检验报告项目',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像设备目录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像检查记录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像检查项目',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像检查类别',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像检查图象',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像检查序列',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像临时记录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像临时图象',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像临时序列',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像屏幕布局',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像预设窗宽窗位',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像图像消隐表',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像鼠标按钮分配',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像界面参数表',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像标注存储表',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像图像信息表',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像打印机设置',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像胶片规格',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像打印格式',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像胶片打印字体',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像颜色清单',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像检查UID序号_ID',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','造影剂',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','服用造影剂',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像流程参数',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病历文件结构',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像病理类别',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人新生儿记录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','记帐报警线',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','保险模拟结算',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','医保病人关联表',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','医保病人档案',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1294,'基本',USER,'病历词句示范','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1294,'基本',USER,'病历词句组成','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1294,'基本',USER,'病历词句条件','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像操作记录',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像图像备注',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病人临床路径',User,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像病理标本',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像标本核收取材',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','影像病理标本部位',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','诊疗收费关系',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','Zl_影像操作记录_Update',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1294,'基本',USER,'f_Sentence_Matched','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像报告内容_创建',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像报告内容_update',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像报告标注_保存',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像报告签名_保存',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像报告回退',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像报告图像_保存',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','Zl_影像检查_检查技师',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像检查_STATE',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像费用执行',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像预设窗宽窗位_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像预设窗宽窗位_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像预设窗宽窗位_DELETE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像窗宽窗位_类型_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像界面参数表_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像界面参数表_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像屏幕布局_类型_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像屏幕布局_类型_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像屏幕布局_DELETE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像屏幕布局_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像图像消隐表_类型_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像图像消隐表_类型_Update',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像图像消隐表_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','Zl_影像图像消隐表_Delete',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像鼠标按钮分配_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像鼠标按钮分配_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像检查_结果',USER,'EXECUTE');
Insert Into zlprogprivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','Zl_绿色通道_Update',USER,'EXECUTE');
Insert Into Zlprogprivs(系统,序号,功能,对象,所有者,权限) values(100,1294,'基本','ZL_影像报告打印_Update',USER,'EXECUTE');
Insert Into Zlprogprivs(系统,序号,功能,对象,所有者,权限) values(100,1294,'基本','ZL_影像报告保存_Update',USER,'EXECUTE');
Insert Into zlprogprivs(系统,序号,功能,对象,所有者,权限) values(100,1294,'基本','ZL_影像报告标记_Clear',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_服用造影剂_INSERT',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_影像报告操作_Update',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_病人医嘱执行_取消拒绝',User,'EXECUTE');
Insert Into zlprogprivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','Zl_电子病历记录_Print',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','Zl_影像图像备注_Insert',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_GetNumber',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','ZL_AgeToDays',USER,'EXECUTE');

--病理基本权限
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理检查信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理标本信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理送检信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理取材信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理脱钙信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理制片信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理特检信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理过程报告',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理申请信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理报告延迟',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理会诊信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理抗体信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理抗体反馈',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理套餐信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理套餐关联',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理归档信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病理借阅信息',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病历词句组成',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'基本','病历词句分类',USER,'SELECT');


Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查登记','zl_挂号病人病案_INSERT',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查登记','ZL_病人医嘱记录_Insert',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查登记','ZL_病人医嘱发送_Insert',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查登记','Zl_病人医嘱附件_Insert',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查登记','ZL_影像检查_SET',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查登记','ZL_影像检查_BEGIN',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查登记','ZL_病人医嘱执行_拒绝执行',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查登记','zl_病人费用记录_医嘱',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查登记','NextNO',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查登记','Zl_病人信息_Update',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'取消报到','ZL_影像检查_CANCEL',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'清除图像','ZL_影像检查_PhotoDelete',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'清除图像','ZL_影像检查_SET',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'清除图像','ZL_影像检查_PhotoCancel',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'视频采集','ZL_影像图象_DELETE',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'视频采集','ZL_影像检查记录_SET',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'视频采集','ZL_影像序列_INSERT',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'视频采集','ZL_影像图象_INSERT',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'视频采集','ZL_影像检查报告_ADD',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'影像质控','Zl_影像质量_Update',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'排队叫号','排队叫号队列',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'排队叫号','排队语音呼叫',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'排队叫号','排队LED显示部件',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'排队叫号','ZL_排队叫号队列_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'排队叫号','ZL_排队叫号队列_UPDATE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'排队叫号','ZL_排队叫号队列_呼叫',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'排队叫号','ZL_排队叫号队列_优先',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'排队叫号','ZL_排队叫号队列_DELETE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'排队叫号','ZL_排队语音呼叫_INSERT',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'排队叫号','ZL_排队语音呼叫_DELETE',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'随访','影像诊断分类',USER,'SELECT');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'随访','Zl_影像随访_Update',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'随访','Zl_影像诊断分类_Update',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'关联病人','ZL_影像关联病人',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'关联病人','ZL_影像取消关联病人',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查报到','zl_挂号病人病案_INSERT',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查报到','ZL_影像检查_SET',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查报到','ZL_影像检查_BEGIN',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'检查报到','NextNO',User,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'未缴费报到','收费项目类别',USER,'SELECT');




--标本核收
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'标本核收','Zl_病理标本_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'标本核收','Zl_病理标本_更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'标本核收','Zl_病理标本_核收',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'标本核收','Zl_病理标本_拒收',USER,'EXECUTE');

--标本取材
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理脱钙_开始',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理脱钙_换缸',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理脱钙_撤销',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理脱钙_完成',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理取材_常规',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理取材_常规更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理取材_细胞',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理取材_细胞更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理取材_冰冻',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理取材_冰冻更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理取材_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理取材_信息保存',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理取材','Zl_病理取材_确认',USER,'EXECUTE');

--病理制片
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理制片','Zl_病理制片_接受',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理制片','Zl_病理制片_清单打印',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'病理制片','Zl_病理制片_确认',USER,'EXECUTE');

--报告延迟
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'报告延迟','Zl_病理报告延迟_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'报告延迟','Zl_病理报告延迟_更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'报告延迟','Zl_病理报告延迟_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'报告延迟','Zl_病理报告延迟_打印',USER,'EXECUTE');

--冰冻，免疫，特染，分子过程报告，冰冻特检报告查阅
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'冰冻报告','Zl_病理过程报告_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'冰冻报告','Zl_病理过程报告_更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'冰冻报告','Zl_病理过程报告_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'冰冻报告','Zl_病理过程报告_状态',USER,'EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'免疫报告','Zl_病理过程报告_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'免疫报告','Zl_病理过程报告_更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'免疫报告','Zl_病理过程报告_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'免疫报告','Zl_病理过程报告_状态',USER,'EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'分子报告','Zl_病理过程报告_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'分子报告','Zl_病理过程报告_更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'分子报告','Zl_病理过程报告_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'分子报告','Zl_病理过程报告_状态',USER,'EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特染报告','Zl_病理过程报告_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特染报告','Zl_病理过程报告_更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特染报告','Zl_病理过程报告_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特染报告','Zl_病理过程报告_状态',USER,'EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'冰冻特检报告查阅','Zl_病理过程报告_状态',USER,'EXECUTE');

--特检申请
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特检申请','Zl_病理申请_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特检申请','Zl_病理申请_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特检申请','Zl_病理申请_特检项目_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特检申请','Zl_病理申请_特检项目_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特检申请','Zl_病理申请_特检项目_重做',USER,'EXECUTE');

--制片申请
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'制片申请','Zl_病理申请_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'制片申请','Zl_病理申请_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'制片申请','Zl_病理申请_制片项目_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'制片申请','Zl_病理申请_制片项目_删除',USER,'EXECUTE');

--补取申请
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'补取申请','Zl_病理申请_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'补取申请','Zl_病理申请_删除',USER,'EXECUTE');

--会诊申请
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'会诊申请','Zl_病理会诊_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'会诊申请','Zl_病理会诊_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'会诊申请','Zl_病理会诊_状态',USER,'EXECUTE');

--会诊反馈
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'会诊反馈','Zl_病理会诊_反馈',USER,'EXECUTE');

--免疫检查
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'免疫组化','Zl_病理特检_接受',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'免疫组化','Zl_病理特检_清单打印',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'免疫组化','Zl_病理特检_项目录入',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'免疫组化','Zl_病理特检_确认',USER,'EXECUTE');

--分子检查
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'分子病理','Zl_病理特检_接受',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'分子病理','Zl_病理特检_清单打印',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'分子病理','Zl_病理特检_项目录入',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'分子病理','Zl_病理特检_确认',USER,'EXECUTE');

--特染检查
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特殊染色','Zl_病理特检_接受',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特殊染色','Zl_病理特检_清单打印',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特殊染色','Zl_病理特检_项目录入',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'特殊染色','Zl_病理特检_确认',USER,'EXECUTE');

--抗体管理
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'抗体管理','Zl_病理抗体_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'抗体管理','Zl_病理抗体_更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'抗体管理','Zl_病理抗体_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'抗体管理','Zl_病理抗体_使用状态',USER,'EXECUTE');

--抗体反馈
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'抗体反馈','Zl_病理抗体反馈_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'抗体反馈','Zl_病理抗体反馈_更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'抗体反馈','Zl_病理抗体反馈_删除',USER,'EXECUTE');

--套餐维护
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'套餐维护','Zl_病理套餐_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'套餐维护','Zl_病理套餐_更新',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'套餐维护','Zl_病理套餐_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'套餐维护','Zl_病理套餐关联_新增',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'套餐维护','Zl_病理套餐关联_删除',USER,'EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(100,1294,'套餐维护','Zl_病理套餐关联_删除1',USER,'EXECUTE');


--导航台菜单
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval, '188','影像病理工作站','病理工作站','C',230,'用于病理标本核收和取材、图像采集、报告填写、附加医嘱和费用的登记',100,1294);
