--体检系统菜单,填写到ZLSOFT中的公共数据

--zlComponent数据
insert into zlComponent(部件,名称,主版本,次版本,附版本,系统) values ('zl9Medical','体检管理部件',10,0,0,100);

----------------------------------------------
--zlPrograms数据
----------------------------------------------
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1850,'体检类型设置','建立、调整、修改体检类型及相应的体检项目。',100,'zl9Medical');
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1860,'体检预约申请','完成体检预约的申请及确认。',100,'zl9Medical');
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1861,'体检工作管理','完成各体检项目的报告填写及体检总结。',100,'zl9Medical');
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1862,'体检团体结算','维护体检人员和体检团体的信息资料。',100,'zl9Medical');
----------------------------------------------
--zlProgFuncs数据
----------------------------------------------
--体检类型设置
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1850,'基本','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1850,'增删改','');
--体检预约申请
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1860,'基本','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1860,'所有科室','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1860,'体检预约','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1860,'确认预约','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1860,'取消预约','');
--体检工作管理
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'基本','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'所有科室','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'开始体检','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'取消开始','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'完成体检','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'取消完成','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'体检项目','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'附加项目','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'添加成员','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'移除成员','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'填写报告','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'填写总结','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'打印报告','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'综合查询','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'费用处理','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'未收费体检','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1861,'科室小结','');

--体检团体结算
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1862,'基本','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1862,'所有科室','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1862,'体检结算','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1862,'结算作废','');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1862,'结算重打','');

----------------------------------------------
--zlProgPrivs数据
----------------------------------------------
--体检类型设置
--基本
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'体检类型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'体检类型目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'诊疗项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'诊疗项目别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'诊疗分类目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'诊疗收费关系','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'收费价目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'收费项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'检验报告项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'检验项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'诊治所见项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'床位状况记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'基本',user,'诊疗用法用量','SELECT');

--增删改
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'增删改',user,'ZL_体检类型_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'增删改',user,'ZL_体检类型_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'增删改',user,'ZL_体检类型_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'增删改',user,'ZL_体检类型目录_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1850,'增删改',user,'ZL_体检类型目录_DELETE','EXECUTE');


--体检预约申请
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'诊疗用法用量','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'床位状况记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'体检类型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'体检类型目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'诊疗项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'诊疗项目别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'诊疗分类目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'诊疗执行科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'诊疗项目组合','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'诊疗收费关系','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'收费价目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'收费项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'部门性质说明','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'部门人员','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'体检登记记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'体检组别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'体检项目清单','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'体检项目清单_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'体检人员档案','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'体检人员档案_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'病人信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'体检登记记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'合约单位_id','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'合约单位','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'检验项目参考','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'检验报告项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'检验项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'诊治所见项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'号码控制表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'号码控制表','UPDATE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'系统参数表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'性别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'民族','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'国籍','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'婚姻状况','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'学历','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'职业','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'身份','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'费别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'ZL_体检登记记录_STATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'基本',user,'ZL_体检登记记录_体检类型','EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检组别_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检组别_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检项目清单_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检项目清单_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检人员档案_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检人员档案_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检人员档案_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检登记记录_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检登记记录_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检登记记录_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'ZL_体检人员档案_CLASS','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'zl_病人信息_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'zl_病人信息_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'zl_合约单位_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1860,'体检预约',user,'zl_合约单位_Update','EXECUTE');

--体检工作管理

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'体检登记记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'体检人员档案','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'体检人员档案_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'体检项目清单','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'体检项目清单_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'体检项目医嘱','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'体检组别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'体检类型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'体检类型目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'部门人员','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'影像检查项目','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'病历模板分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'病历模板应用','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'病历模板内容','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'材料特性','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'学历','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'身份','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'费别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人医嘱记录_ID','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'号码控制表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'号码控制表','UPDATE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'zlGetReference','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'zlGetResult','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人病历记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人余额','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人医嘱记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人医嘱发送','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病案主页','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病历文件目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'部门性质说明','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验标本记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验普通结果','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验报告项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验仪器','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验药敏结果','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验用抗生素','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验细菌','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验标本形态','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验标本记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验普通结果_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验项目取值','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验项目参考','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验仪器项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验抗生素用药','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验抗生素组','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验细菌抗生素','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'检验细菌类型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗分类目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗项目别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗检验标本','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'医疗付款方式','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'收费执行科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'部门安排','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人医嘱附费','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人医嘱执行','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗收费关系','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'收费价目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'收费项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'收入项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'收费项目别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'收费项目类别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'药品规格','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'记帐报警线','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'保险模拟结算','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'收费从属项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'药品收发记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'药品库存','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'未发药品记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'费用类型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'药品单据性质','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人手术记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人手术情况','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人手术人员','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人诊断记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人医嘱状态','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人过敏记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人费用记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人病历文本段','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人病历外部图','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人病历所见单','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人病历内容','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人病历标记图','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人变动记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病情','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病案主页从表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病历元素目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病历标记图','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病历所见单','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病历示范目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病历文件组成','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'医技执行房间','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'手术岗位','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'药品材质分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'药品用途分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'药品特性','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'药品出库检查','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'收费执行部门','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'收费分类目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗项目类别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊治所见性质','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊治所见分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊治所见项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗执行科室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗用法用量','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗项目组合','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗频率项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗频率时间','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗互斥项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗单据应用','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病诊断目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病诊断别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病诊断分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病诊断属类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病诊断对照','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病编码目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病编码分类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病诊断参考','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病参考项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病编码类别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病诊疗措施','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'疾病诊断规则','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'床位状况记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'人员性质说明','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'婚姻状况','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'合约单位','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'国籍','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'地区','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'中药煎服脚注','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'职业','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'血型','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'性别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'系统参数表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'特殊符号','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'社会关系','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'区域','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'民族','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'门诊诊室','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'诊疗手术规模','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'ZL_麻醉记录项目_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'zl_PatiDayCharge','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'保险类别','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'保险参数','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'保险帐户','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'保险支付项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'保险支付大类','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',User,'病人病历修订记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',User,'病人挂号记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'病人医嘱计价','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'费别明细','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人过敏药物','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'门诊诊室','UPDATE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'病人麻醉用药','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人费用记录_更新医保','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人费用记录_上传','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人记帐记录_上传','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检人员档案_报到','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'保险病种','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'保险特准项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'保险项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'帐户年度信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'病人病历修订记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'体检人员结论','SELECT');

insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'基本',USER,'ZL_体检人员结论_REFRESH','EXECUTE');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'基本',USER,'ZL_体检人员结论_UPDATE','EXECUTE');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'基本',USER,'ZL_病人医嘱发送_计费','EXECUTE');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'基本',USER,'zl_门诊划价记录_Insert','EXECUTE');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'基本',USER,'zl_门诊记帐记录_Insert','EXECUTE');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'基本',USER,'zl_住院记帐记录_Insert','EXECUTE');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'基本',USER,'zl_住院记帐记录_DELETE','EXECUTE');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'基本',USER,'zl_门诊记帐记录_DELETE','EXECUTE');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'基本',USER,'zl_门诊划价记录_DELETE','EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'病人病历记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'病人病历内容_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'病人手麻记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历_归档','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历_作废','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历标记图_SAVE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历附加表_SAVE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历附加表单元_SAVE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历内容_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历内容_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历所见单_SAVE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历外部图_检查','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历文本段_SAVE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人出院诊断记录单_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人门诊诊断记录单_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人手术概要手麻记录_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_检验结果记录_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'病历示范目录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病历示范目录_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病历示范目录_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病历示范目录_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'ZL_病人病历记录体麻_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'ZL_病人麻醉记录手麻_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'ZL_病人麻醉记录用药_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'ZL_病人麻醉记录标注_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',USER,'ZL_病人麻醉记录标注_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人病历修订_INSERT','EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_诊疗单据_报告','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_诊疗单据_申请','EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'zl_病人信息_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'zl_病人信息_Update','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检人员档案_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检人员档案_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检登记记录_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检登记记录_UPDATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检登记记录_DELETE','EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检登记记录_Cancel','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检项目医嘱_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_门诊医嘱发送_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人医嘱记录_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检登记记录_STATE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检人员档案_总结','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'zl_病人病历_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检人员档案_复查','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检登记记录_Finish','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检登记记录_CancelFinish','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_病人医嘱附费_Insert','EXECUTE');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检人员档案_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检登记记录_ItemCancel','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检项目清单_DELETE','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检项目清单_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1861,'基本',user,'ZL_体检登记记录_单项填写','EXECUTE');

--体检团体结算

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'体检结算记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'体检结算记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'体检结算清单','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'系统参数表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'体检登记记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'病人结帐记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'病人费用记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'合约单位','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'号码控制表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'号码控制表','UPDATE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'病人预交记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'部门性质说明','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'部门人员','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'票据使用明细','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'票据领用记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'结算方式应用','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'结算方式','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'病人结帐记录_ID','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'病人医嘱记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'收入项目','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'收费项目别名','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'收费项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'结算作废',user,'zl_病人结帐记录_Delete','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'体检结算',user,'zl_病人结帐记录_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'体检结算',user,'zl_病人结帐票据_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'体检结算',user,'zl_结帐缴款记录_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'体检结算',user,'zl_结帐费用记录_Insert','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'结算重打',user,'zl_病人结帐记录_RePrint','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'体检结算',user,'ZL_体检结算记录_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'体检结算',user,'ZL_体检结算清单_INSERT','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values (100,1862,'基本',user,'ZL_体检结算记录_Cancel','EXECUTE');


----------------------------------------------
--zlMenus数据 
----------------------------------------------
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,null,'体检管理系统','体检管理','A',99,'',100,NULL);
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,zlMenus_id.nextval-1,'体检类型设置','体检类型','A',99,'建立、调整、修改体检类型及相应的体检项目。',100,1850);
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,zlMenus_id.nextval-2,'体检预约申请','体检预约','B',213,'完成体检预约的申请及确认。',100,1860);
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,zlMenus_id.nextval-3,'体检工作管理','体检工作','D',225,'完成各体检项目的报告填写及体检总结。',100,1861);
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval,zlMenus_id.nextval-4,'体检团体结算','团体结算','E',99,'维护体检人员和体检团体的信息资料。',100,1862);

--非体检系统调整
Alter Table 合约单位 Add 电子邮件 varchar(50);
Alter Table 合约单位 Add 说明 varchar(2000);

CREATE INDEX 病人医嘱记录_IX_挂号单 ON 病人医嘱记录(挂号单) PCTFREE 10 TABLESPACE zl9CisRec;
Insert Into 部门性质分类(编码,名称,简码,服务病人,说明) Select 'Q','体检','TJ',3,'' From dual;
insert into 号码控制表(项目序号,项目名称,最大号码,自动补缺,编号规则) values (78,'体检单号','',1,null);

---------------------------------------------
-- 对"合约单位"作增加操作
----------------------------------------------
CREATE OR REPLACE PROCEDURE zl_合约单位_Insert (
    ID_IN IN 合约单位.ID%TYPE,
    上级ID_IN IN 合约单位.上级ID%TYPE,
    编码_IN IN 合约单位.编码%TYPE,
    名称_IN IN 合约单位.名称%TYPE,
    简码_IN IN 合约单位.简码%TYPE := NULL,
    地址_IN IN 合约单位.地址%TYPE := NULL,
    电话_IN IN 合约单位.电话%TYPE := NULL,
    开户银行_IN IN 合约单位.开户银行%TYPE := NULL,
    帐号_IN IN 合约单位.帐号%TYPE := NULL,
    联系人_IN IN 合约单位.联系人%TYPE := NULL,
    末级_IN IN 合约单位.末级%TYPE := 1,
    电子邮件_IN IN 合约单位.电子邮件%TYPE := NULL,
    说明_IN IN 合约单位.说明%TYPE := NULL
)
IS
BEGIN
    --首先插入记录
    Insert INTO 合约单位
                    (
                        ID,
                        编码,
                        名称,
                        简码,
                        地址,
                        电话,
                        开户银行,
                        帐号,
                        联系人,
                        上级ID,
                        建档时间,
                        撤档时间,
                        末级,
			电子邮件,
			说明
                    )
          VALUES (
              ID_IN,
              编码_IN,
              名称_IN,
              简码_IN,
              地址_IN,
              电话_IN,
              开户银行_IN,
              帐号_IN,
              联系人_IN,
              上级ID_IN,
              SYSDATE,
              TO_DATE ('3000-01-01', 'yyyy-mm-dd'),
              末级_IN,
	      电子邮件_IN,
	      说明_IN
          );
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_合约单位_Insert;
/
---------------------------------------------
-- 对"合约单位"作修改操作
----------------------------------------------
CREATE OR REPLACE PROCEDURE zl_合约单位_UPDATE (
    ID_IN IN 合约单位.ID%TYPE,
    上级ID_IN IN 合约单位.上级ID%TYPE,
    编码_IN IN 合约单位.编码%TYPE,
    名称_IN IN 合约单位.名称%TYPE,
    简码_IN IN 合约单位.简码%TYPE,
    地址_IN IN 合约单位.地址%TYPE := NULL,
    电话_IN IN 合约单位.电话%TYPE := NULL,
    开户银行_IN IN 合约单位.开户银行%TYPE := NULL,
    帐号_IN IN 合约单位.帐号%TYPE := NULL,
    联系人_IN IN 合约单位.联系人%TYPE := NULL,
    原长度_IN IN PLS_INTEGER,
    电子邮件_IN IN 合约单位.电子邮件%TYPE := NULL,
    说明_IN IN 合约单位.说明%TYPE := NULL
)
IS
BEGIN
    --首先插入修改记录
    UPDATE 合约单位
        SET 编码 = 编码_IN,
             名称 = 名称_IN,
             简码 = 简码_IN,
             地址 = 地址_IN,
             电话 = 电话_IN,
             开户银行 = 开户银行_IN,
             帐号 = 帐号_IN,
             联系人 = 联系人_IN,
             上级ID = 上级ID_IN,
	     电子邮件=电子邮件_IN,
	     说明=说明_IN
     WHERE ID = ID_IN;

    --对它的下级也要修改编码
    UPDATE 合约单位
        SET 编码 = 编码_IN || SUBSTR (编码, 原长度_IN)
     WHERE ID IN (SELECT ID
                         FROM 合约单位
                        START WITH 上级ID = ID_IN
                      CONNECT BY PRIOR ID = 上级ID);
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_合约单位_UPDATE;
/

--体检系统数据表
Create Table 体检类型(
	序号		NUMBER(18),
	上级序号	NUMBER(18),			--new
	编码		VARCHAR2(10),	
	名称		VARCHAR2(30),
	简码		VARCHAR2(30),
	末级		number(1) default 0,		--new
	说明		VARCHAR2(100))
    TABLESPACE zl9BaseItem
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

Create Table 体检类型目录(
	序号		NUMBER(18),
	诊疗项目id	NUMBER(18))	
    TABLESPACE zl9BaseItem
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

Create Table 体检登记记录(
	ID		NUMBER(18),
	体检号		varchar2(10),		--规则:年月	
	记录性质	NUMBER(3),		--
	体检状态	NUMBER(3),		--1:新开预约;2:确认预约;3:撤消预约确认;4:正在体检;5:体检完成
	联系人		VARCHAR2(20),
	联系电话	VARCHAR2(30),
	移动电话	VARCHAR(20),
	联系地址	VARCHAR2(50),
	合约单位id	NUMBER(18),
	体检人数	NUMBER(5),
	结算折扣	NUMBER(5,2) DEFAULT 1,		--new
	体检时间	DATE,
	体检类型	VARCHAR2(1000),		--此项目所属类型new
	体检部门id	NUMBER(18),
	附加说明	VARCHAR2(2000),
	是否团体	NUMBER(1) DEFAULT 0,
	登记时间	DATE,
	完成时间	DATE)
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

Create Table 体检组别(
	登记id		NUMBER(18),
	组别名称	VARCHAR2(30),
	说明		VARCHAR2(100))
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

CREATE TABLE 体检项目清单(
	ID		NUMBER(18),
	登记id		NUMBER(18),
	组别名称	VARCHAR2(30),
	病人id		NUMBER(18),		--此项有值时，表示此病人的私人体检项目
	诊疗项目id	NUMBER(18),
	执行科室id	NUMBER(18),
	采集方式id	NUMBER(18),
	体检类型	VARCHAR2(30),		--此项目所属类型new
	结算途径	number(1) default 1,	
	检验标本	VARCHAR2(50),		--检验项目的标本类型
	检查部位	VARCHAR2(4000),		--多部位时，以逗号分隔保存名称，如
	检查部位id	VARCHAR2(4000))		--多部位时，以逗号分隔保存id，如234,67,9821))		
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

CREATE TABLE 体检项目医嘱(
	清单id		NUMBER(18),	
	病人id		NUMBER(18),	
	医嘱id		NUMBER(18))		--记录生成医嘱的主医嘱id(相关id为NULL)		
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

Create Table 体检人员档案(
	ID		NUMBER(18),
	登记id		NUMBER(18),
	病人id		NUMBER(18),
	体检状态	NUMBER(3) DEFAULT 1,		--1:预约;4:正在体检;5:体检完成
	组别名称	VARCHAR2(30),
	姓名		VARCHAR2(20),			
	性别		VARCHAR2(10),			
	年龄		VARCHAR2(20),			
	婚姻状况	VARCHAR2(20),			
	联系电话	VARCHAR2(30),			
	移动电话	VARCHAR2(20),			
	联系地址	VARCHAR2(50),			
	复查时间	DATE,
	体检报到	NUMBER(2) DEFAULT 0,
	体检病历id	number(18),
	完成时间	DATE)
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

CREATE TABLE 体检人员结论(
	登记id		NUMBER(18),
	病人id		NUMBER(18),
	科室id		NUMBER(18),			--NULL,表示总结
	结论id		NUMBER(18))
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

create table 体检结算记录
(
  ID         NUMBER(18),
  记录状态   NUMBER(1),
  合约单位ID NUMBER(18),
  结算ID     NUMBER(18),
  结算部门ID NUMBER(18),
  结算金额   NUMBER(16,5)
)
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

create table 体检结算清单
(
  结算ID NUMBER(18),
  登记ID NUMBER(18)
)
    TABLESPACE zl9CisRec
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);

Create Sequence 体检登记记录_ID start with 1;
Create Sequence 体检人员档案_ID start with 1;
Create Sequence 体检项目清单_ID start with 1;
Create Sequence 体检结算记录_ID start with 1;

--体检系统索引

CREATE INDEX 体检类型目录_IX_序号 on 体检类型目录(序号) PCTFREE 10 TABLESPACE zl9BaseItem;

CREATE INDEX 体检登记记录_IX_合约单位id on 体检登记记录(合约单位id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检登记记录_IX_体检部门id on 体检登记记录(体检部门id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检登记记录_IX_体检时间 on 体检登记记录(体检时间) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检组别_IX_登记id on 体检组别(登记id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检项目清单_IX_登记id on 体检项目清单(登记id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检项目清单_IX_执行科室id on 体检项目清单(执行科室id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检项目清单_IX_诊疗项目id on 体检项目清单(诊疗项目id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检项目清单_IX_采集方式id on 体检项目清单(采集方式id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检项目清单_IX_病人id on 体检项目清单(病人id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检项目医嘱_IX_清单id on 体检项目医嘱(清单id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检项目医嘱_IX_病人id on 体检项目医嘱(病人id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检项目医嘱_IX_医嘱id on 体检项目医嘱(医嘱id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检人员档案_IX_登记id on 体检人员档案(登记id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检人员档案_IX_病人id on 体检人员档案(病人id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检人员档案_IX_体检病历id on 体检人员档案(体检病历id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检结算记录_IX_合约单位id on 体检结算记录(合约单位id) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检结算记录_IX_结算部门ID on 体检结算记录(结算部门ID) PCTFREE 10 TABLESPACE zl9CisRec;
CREATE INDEX 体检结算记录_IX_结算id on 体检结算记录(结算id) PCTFREE 10 TABLESPACE zl9CisRec;

--体检系统约束
ALTER TABLE 体检类型 ADD CONSTRAINT 体检类型_PK PRIMARY KEY (序号) USING INDEX PCTFREE 15 TABLESPACE zl9BaseItem;
ALTER TABLE 体检类型 ADD CONSTRAINT 体检类型_UQ_编码 UNIQUE (编码) USING INDEX PCTFREE 5 TABLESPACE zl9BaseItem;
ALTER TABLE 体检类型 ADD CONSTRAINT 体检类型_UQ_名称 UNIQUE (名称) USING INDEX PCTFREE 5 TABLESPACE zl9BaseItem;
ALTER TABLE 体检类型 ADD CONSTRAINT 体检类型_CK_缺省 CHECK (缺省 IN(0,1));
ALTER TABLE 体检类型 ADD CONSTRAINT 体检类型_FK_上级序号 FOREIGN KEY (上级序号) REFERENCES 体检类型(序号) ON DELETE CASCADE;

ALTER TABLE 体检类型目录 ADD CONSTRAINT 体检类型目录_FK_序号 FOREIGN KEY (序号) REFERENCES 体检类型(序号) ON DELETE CASCADE;
ALTER TABLE 体检类型目录 ADD CONSTRAINT 体检类型目录_FK_诊疗项目id FOREIGN KEY (诊疗项目id) REFERENCES 诊疗项目目录(ID);

ALTER TABLE 体检登记记录 ADD CONSTRAINT 体检登记记录_PK PRIMARY KEY (ID) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE 体检登记记录 ADD CONSTRAINT 体检登记记录_UQ_体检号 UNIQUE (体检号) USING INDEX PCTFREE 5 TABLESPACE zl9CisRec;
ALTER TABLE 体检登记记录 ADD CONSTRAINT 体检登记记录_CK_是否团体 CHECK (是否团体 IN(0,1));
ALTER TABLE 体检登记记录 ADD CONSTRAINT 体检登记记录_CK_体检状态 CHECK (体检状态 IN(1,2,3,4,5));
ALTER TABLE 体检登记记录 ADD CONSTRAINT 体检登记记录_FK_合约单位id FOREIGN KEY (合约单位id) REFERENCES 合约单位(ID);
ALTER TABLE 体检登记记录 ADD CONSTRAINT 体检登记记录_FK_体检部门id FOREIGN KEY (体检部门id) REFERENCES 部门表(ID);

ALTER TABLE 体检组别 ADD CONSTRAINT 体检组别_PK PRIMARY KEY (登记id,组别名称) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE 体检组别 ADD CONSTRAINT 体检组别_FK_登记id FOREIGN KEY (登记id) REFERENCES 体检登记记录(ID);

ALTER TABLE 体检项目清单 ADD CONSTRAINT 体检项目清单_PK PRIMARY KEY (ID) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE 体检项目清单 ADD CONSTRAINT 体检项目清单_FK_登记id FOREIGN KEY (登记id) REFERENCES 体检登记记录(ID);
ALTER TABLE 体检项目清单 ADD CONSTRAINT 体检项目清单_FK_诊疗项目id FOREIGN KEY (诊疗项目id) REFERENCES 诊疗项目目录(ID);
ALTER TABLE 体检项目清单 ADD CONSTRAINT 体检项目清单_FK_执行科室id FOREIGN KEY (执行科室id) REFERENCES 部门表(ID);
ALTER TABLE 体检项目清单 ADD CONSTRAINT 体检项目清单_FK_采集方式id FOREIGN KEY (采集方式id) REFERENCES 诊疗项目目录(ID);
ALTER TABLE 体检项目清单 ADD CONSTRAINT 体检项目清单_FK_病人id FOREIGN KEY (病人id) REFERENCES 病人信息(病人id);

ALTER TABLE 体检项目医嘱 ADD CONSTRAINT 体检项目医嘱_PK PRIMARY KEY (清单id,病人id,医嘱id) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE 体检项目医嘱 ADD CONSTRAINT 体检项目医嘱_FK_清单id FOREIGN KEY (清单id) REFERENCES 体检项目清单(ID) ON DELETE CASCADE;
ALTER TABLE 体检项目医嘱 ADD CONSTRAINT 体检项目医嘱_FK_病人id FOREIGN KEY (病人id) REFERENCES 病人信息(病人id);
ALTER TABLE 体检项目医嘱 ADD CONSTRAINT 体检项目医嘱_FK_医嘱id FOREIGN KEY (医嘱id) REFERENCES 病人医嘱记录(ID);

ALTER TABLE 体检人员档案 ADD CONSTRAINT 体检人员档案_PK PRIMARY KEY (ID) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE 体检人员档案 ADD CONSTRAINT 体检人员档案_FK_登记id FOREIGN KEY (登记id) REFERENCES 体检登记记录(ID);
ALTER TABLE 体检人员档案 ADD CONSTRAINT 体检人员档案_FK_病人id FOREIGN KEY (病人id) REFERENCES 病人信息(病人id);
ALTER TABLE 体检人员档案 ADD CONSTRAINT 体检人员档案_CK_体检状态 CHECK (体检状态 IN(1,4,5));
ALTER TABLE 体检人员档案 ADD CONSTRAINT 体检人员档案_FK_体检病历id FOREIGN KEY (体检病历id) REFERENCES 病人病历记录(ID);

ALTER TABLE 体检人员结论 ADD CONSTRAINT 体检人员结论_PK PRIMARY KEY (登记id,病人id,科室id) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE 体检人员结论 ADD CONSTRAINT 体检人员结论_FK_登记id FOREIGN KEY (登记id) REFERENCES 体检登记记录(ID);
ALTER TABLE 体检人员结论 ADD CONSTRAINT 体检人员结论_FK_病人id FOREIGN KEY (病人id) REFERENCES 病人信息(病人id);
ALTER TABLE 体检人员结论 ADD CONSTRAINT 体检人员结论_FK_科室id FOREIGN KEY (科室id) REFERENCES 部门表(ID);
ALTER TABLE 体检人员结论 ADD CONSTRAINT 体检人员结论_FK_结论id FOREIGN KEY (结论id) REFERENCES 病人病历记录(ID);

ALTER TABLE 体检结算记录 ADD CONSTRAINT 体检结算记录_PK PRIMARY KEY (ID) USING INDEX PCTFREE 15 TABLESPACE zl9CisRec;
ALTER TABLE 体检结算记录 ADD CONSTRAINT 体检结算记录_CK_记录状态 CHECK (记录状态 IN(1,2));
ALTER TABLE 体检结算记录 ADD CONSTRAINT 体检结算记录_FK_合约单位id FOREIGN KEY (合约单位id) REFERENCES 合约单位(ID);
ALTER TABLE 体检结算记录 ADD CONSTRAINT 体检结算记录_FK_结算部门ID FOREIGN KEY (结算部门ID) REFERENCES 部门表(ID);
ALTER TABLE 体检结算记录 ADD CONSTRAINT 体检结算记录_FK_结算id FOREIGN KEY (结算id) REFERENCES 病人结帐记录(ID);


--体检系统过程
----------------------------------------------------------------------------
---  INSERT   for   体检类型
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检类型_INSERT(
	序号_IN IN 体检类型.序号%TYPE,
	编码_IN IN 体检类型.编码%TYPE,
	名称_IN IN 体检类型.名称%TYPE,
	简码_IN IN 体检类型.简码%TYPE,
	说明_IN IN 体检类型.说明%TYPE,
	上级序号_IN IN 体检类型.上级序号%TYPE:=NULL,
	末级_IN IN 体检类型.末级%TYPE:=1,
	同级调整_IN  NUMBER:=0
)
IS
	v_Extend number(18);
	v_Parent varchar2(30);
BEGIN	
	IF 末级_IN=0 THEN
		IF 同级调整_IN=1 THEN
			    --调整同级编码的长度
			IF NVL(上级序号_IN,0)<>0 THEN
			    SELECT 编码 INTO v_Parent FROM 体检类型 WHERE 序号=上级序号_IN;
			ELSE
			    v_Parent:=NULL;
			END IF;

			BEGIN
			    SELECT length(rtrim(编码_IN))-length(rtrim(编码)) INTO v_Extend
			    FROM 体检类型
			    WHERE 末级=0 AND (上级序号=上级序号_IN OR 上级序号 IS NULL AND NVL(上级序号_IN,0)=0) AND Rownum=1;
			EXCEPTION
			    WHEN OTHERS THEN v_Extend:=0;
			END;

			IF v_Extend>0 THEN
			    --扩充处理
			    IF v_Parent IS null THEN
				UPDATE 体检类型 SET 编码=lpad('0',v_Extend,'0')||编码 WHERE 序号<>序号_IN AND 末级=0;
			    ELSE
				UPDATE 体检类型 SET 编码=v_Parent||lpad('0',v_Extend,'0')||substr(编码,length(v_Parent)+1) WHERE 编码 LIKE v_Parent||'_%' AND 末级=0;
			    END IF;
			END IF;

			IF v_Extend<0 THEN
			    --压缩处理
			    IF v_Parent IS null THEN
				UPDATE 体检类型 SET 编码=substr(编码,1+abs(v_Extend)) WHERE 序号<>序号_IN AND 末级=0;
			    ELSE
				UPDATE 体检类型 SET 编码=v_Parent||substr(编码,length(v_Parent)+1+abs(v_Extend)) WHERE 编码 LIKE v_Parent||'_%' AND 末级=0;
			    END IF;
			END IF;

		END IF;
	END IF;
	Insert Into 体检类型(序号,上级序号,末级,编码,名称,简码,说明) VALUES(序号_IN,DECODE(上级序号_IN,0,NULL,上级序号_IN),末级_IN,编码_IN,名称_IN,简码_IN,说明_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检类型_INSERT;
/
----------------------------------------------------------------------------
---  UPDATE   for   体检类型
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检类型_UPDATE(
	序号_IN IN 体检类型.序号%TYPE,
	编码_IN IN 体检类型.编码%TYPE,
	名称_IN IN 体检类型.名称%TYPE,
	简码_IN IN 体检类型.简码%TYPE,
	说明_IN IN 体检类型.说明%TYPE,
	上级序号_IN IN 体检类型.上级序号%TYPE:=NULL,
	同级调整_IN  NUMBER:=0
)
IS
	v_OldCode  VARCHAR2(30);  --原来的编码
	v_Parent  VARCHAR2(30);  --上级编码
	v_Extend  NUMBER(18);    --扩充长度(为负表示压缩)
	Err_NotFind  EXCEPTION;
BEGIN
	
	SELECT rtrim(编码) INTO v_OldCode FROM 体检类型 WHERE 序号=序号_IN;
	IF v_OldCode is null THEN
		RAISE Err_NotFind;
	END IF;

	--修改项目本身
	Update 体检类型
		Set 编码=编码_IN,
		    名称=名称_IN,
		    简码=简码_IN,
		    说明=说明_IN,
		    上级序号=DECODE(上级序号_IN,0,NULL,上级序号_IN)
	WHERE 序号=序号_IN;    

	--修改本系各级下属编码

	UPDATE 体检类型 SET 编码=编码_IN||substr(编码,length(v_OldCode)+1) WHERE 编码<>编码_IN And 编码 LIKE v_OldCode||'_%' And 末级=0;

	--调整同级编码的长度
	IF 同级调整_IN=1 THEN
		IF NVL(上级序号_IN,0)<>0 THEN
		    SELECT 编码 INTO v_Parent FROM 体检类型 WHERE 序号=上级序号_IN;
		ELSE
		    v_Parent:=NULL;
		END IF;

		BEGIN
		    SELECT length(rtrim(编码_IN))-length(rtrim(编码)) INTO v_Extend FROM 体检类型 WHERE 末级=0 AND (上级序号=上级序号_IN OR 上级序号 IS NULL AND nvl(上级序号_IN,0)=0) AND 序号<>序号_IN AND Rownum=1;
		EXCEPTION
		    WHEN OTHERS THEN v_Extend:=0;
		END;

		IF v_Extend>0 THEN
		    --扩充处理
		    IF v_Parent IS null THEN
			UPDATE 体检类型 SET 编码=lpad('0',v_Extend,'0')||编码  WHERE 末级=0 and 序号 not in (select 序号 from 体检类型 WHERE 末级=0 start with 序号=序号_IN connect by prior 序号=上级序号);
		    ELSE
			UPDATE 体检类型	SET 编码=v_Parent||lpad('0',v_Extend,'0')||substr(编码,length(v_Parent)+1) WHERE 末级=0 AND 编码 LIKE v_Parent||'_%' and 序号 not in (select 序号 from 体检类型 where 末级=0 start with 序号=序号_IN connect by prior 序号=上级序号);
		    END IF;
		END IF;

		IF v_Extend<0 THEN
		    --压缩处理
		    IF v_Parent IS null THEN
			UPDATE 体检类型 SET 编码=substr(编码,1+abs(v_Extend)) WHERE 序号<>序号_IN AND 末级=0;
		    ELSE
			UPDATE 体检类型 SET 编码=v_Parent||substr(编码,length(v_Parent)+1+abs(v_Extend)) WHERE 编码 LIKE v_Parent||'_%' AND 序号<>序号_IN AND 末级=0;
		    END IF;
		END IF;
	END IF;
EXCEPTION
	WHEN Err_NotFind THEN Raise_application_error (-20101, '[ZLSOFT]该项目不存在，可能已被其他用户删除！[ZLSOFT]');
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检类型_UPDATE;
/

----------------------------------------------------------------------------
---  DELETE   for   体检类型
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检类型_DELETE(
	序号_IN IN 体检类型.序号%TYPE
)
IS
BEGIN
	DELETE FROM 体检类型 WHERE 序号=序号_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检类型_DELETE;
/
----------------------------------------------------------------------------
---  INSERT   for   体检类型目录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检类型目录_INSERT(
	序号_IN IN 体检类型目录.序号%TYPE,
	诊疗项目id_IN IN 体检类型目录.诊疗项目id%TYPE
)
IS
BEGIN
	Insert Into 体检类型目录(序号,诊疗项目id)
		VALUES(序号_IN,诊疗项目id_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检类型目录_INSERT;
/
----------------------------------------------------------------------------
---  DELETE   for   体检类型目录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检类型目录_DELETE(
	序号_IN IN 体检类型目录.序号%TYPE
)
IS
BEGIN
	DELETE FROM 体检类型目录 WHERE 序号=序号_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检类型目录_DELETE;
/
----------------------------------------------------------------------------
---  INSERT   for   体检登记记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检登记记录_INSERT(
	ID_IN IN 体检登记记录.ID%TYPE,
	体检号_IN IN 体检登记记录.体检号%TYPE,
	记录性质_IN IN 体检登记记录.记录性质%TYPE,
	体检状态_IN IN 体检登记记录.体检状态%TYPE,
	联系人_IN IN 体检登记记录.联系人%TYPE,
	联系电话_IN IN 体检登记记录.联系电话%TYPE,
	移动电话_IN IN 体检登记记录.移动电话%TYPE,
	联系地址_IN IN 体检登记记录.联系地址%TYPE,
	合约单位ID_IN IN 体检登记记录.合约单位ID%TYPE,
	体检人数_IN IN 体检登记记录.体检人数%TYPE,
	体检时间_IN IN 体检登记记录.体检时间%TYPE,
	体检部门ID_IN IN 体检登记记录.体检部门ID%TYPE,
	附加说明_IN IN 体检登记记录.附加说明%TYPE,
	登记时间_IN IN 体检登记记录.登记时间%TYPE,
	完成时间_IN IN 体检登记记录.完成时间%TYPE,
	是否团体_IN IN 体检登记记录.是否团体%TYPE:=0,
	结算折扣_IN IN 体检登记记录.结算折扣%TYPE:=1
)
IS
BEGIN
	Insert Into 体检登记记录
		(ID,体检号,记录性质,体检状态,联系人,联系电话,移动电话,联系地址,合约单位ID,体检人数,体检时间,体检部门ID,附加说明,登记时间,完成时间,是否团体,结算折扣)
		VALUES
		(ID_IN,体检号_IN,记录性质_IN,体检状态_IN,联系人_IN,联系电话_IN,移动电话_IN,联系地址_IN,合约单位ID_IN,体检人数_IN,体检时间_IN,体检部门ID_IN,附加说明_IN,登记时间_IN,完成时间_IN,是否团体_IN,DECODE(结算折扣_IN,0,1,结算折扣_IN));
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检登记记录_INSERT;
/
----------------------------------------------------------------------------
---  UPDATE   for   体检登记记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检登记记录_UPDATE(
	ID_IN IN 体检登记记录.ID%TYPE,
	体检号_IN IN 体检登记记录.体检号%TYPE,
	记录性质_IN IN 体检登记记录.记录性质%TYPE,
	体检状态_IN IN 体检登记记录.体检状态%TYPE,
	联系人_IN IN 体检登记记录.联系人%TYPE,
	联系电话_IN IN 体检登记记录.联系电话%TYPE,
	移动电话_IN IN 体检登记记录.移动电话%TYPE,
	联系地址_IN IN 体检登记记录.联系地址%TYPE,
	合约单位ID_IN IN 体检登记记录.合约单位ID%TYPE,
	体检人数_IN IN 体检登记记录.体检人数%TYPE,
	体检时间_IN IN 体检登记记录.体检时间%TYPE,
	体检部门ID_IN IN 体检登记记录.体检部门ID%TYPE,
	附加说明_IN IN 体检登记记录.附加说明%TYPE,
	登记时间_IN IN 体检登记记录.登记时间%TYPE,
	完成时间_IN IN 体检登记记录.完成时间%TYPE,
	结算折扣_IN IN 体检登记记录.结算折扣%TYPE:=1
)
IS
BEGIN
	Update 体检登记记录
		Set 		    
		    体检号=体检号_IN,
		    记录性质=记录性质_IN,
		    体检状态=体检状态_IN,
		    联系人=联系人_IN,
		    联系电话=联系电话_IN,
		    移动电话=移动电话_IN,
		    联系地址=联系地址_IN,
		    合约单位ID=合约单位ID_IN,
		    体检人数=体检人数_IN,
		    体检时间=体检时间_IN,
		    体检部门ID=体检部门ID_IN,
		    附加说明=附加说明_IN,
		    登记时间=登记时间_IN,
		    完成时间=完成时间_IN,
		    结算折扣=DECODE(结算折扣_IN,0,1,结算折扣_IN)
	WHERE ID=ID_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检登记记录_UPDATE;
/

----------------------------------------------------------------------------
---  GROUP   for   体检登记记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检登记记录_GROUP(
	ID_IN IN 体检登记记录.ID%TYPE,
	合约单位ID_IN IN 体检登记记录.合约单位ID%TYPE
)
IS
BEGIN
	Update 体检登记记录
		Set 
		    合约单位ID=合约单位ID_IN
	WHERE ID=ID_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检登记记录_GROUP;
/

----------------------------------------------------------------------------
---  DELETE   for   体检登记记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检登记记录_DELETE(
	ID_IN IN 体检登记记录.ID%TYPE
)
IS
BEGIN
	Delete from 体检人员档案 WHERE 登记id=ID_IN;
	Delete from 体检项目清单 WHERE 登记id=ID_IN;
	Delete from 体检组别 WHERE 登记id=ID_IN;
	Delete From 体检登记记录 WHERE ID=ID_IN;
EXCEPTION
	WHEN OTHERS THEN
		zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检登记记录_DELETE;
/
----------------------------------------------------------------------------
---  STATE   for   体检登记记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检登记记录_STATE(
	ID_IN		IN 体检登记记录.ID%TYPE,
	体检状态_IN	IN 体检登记记录.体检状态%TYPE,
	病人id_IN	IN 体检人员档案.病人id%TYPE:=0
)
IS
BEGIN
	IF 病人id_IN=0 THEN
		UPDATE 体检登记记录 SET 体检状态=体检状态_IN WHERE ID=ID_IN;
		IF 体检状态_IN=4 THEN
			UPDATE 体检人员档案 SET 体检状态=体检状态_IN WHERE 登记id=ID_IN;
			UPDATE 体检登记记录 SET 体检时间=SYSDATE WHERE ID=ID_IN;
		END IF;
	ELSE
		UPDATE 体检人员档案 SET 体检状态=体检状态_IN WHERE 登记id=ID_IN AND 病人id=病人id_IN;		
	END IF;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检登记记录_STATE;
/

----------------------------------------------------------------------------
---  INSERT   for   体检组别
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检组别_INSERT(
	登记id_IN IN 体检组别.登记id%TYPE,
	组别名称_IN IN 体检组别.组别名称%TYPE,
	说明_IN IN 体检组别.说明%TYPE:=null
)
IS
BEGIN
	Insert Into 体检组别(登记id,组别名称,说明)
		VALUES
		(登记id_IN,组别名称_IN,说明_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检组别_INSERT;
/

----------------------------------------------------------------------------
---  DELETE   for   体检组别
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检组别_DELETE(
	登记id_IN IN 体检组别.登记id%TYPE
)
IS
BEGIN
	Delete from 体检组别 where 登记id=登记id_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检组别_DELETE;
/
----------------------------------------------------------------------------
---  INSERT   for   体检组别项目
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检项目清单_INSERT(
	登记id_IN IN 体检项目清单.登记id%TYPE,
	组别名称_IN IN 体检项目清单.组别名称%TYPE,
	诊疗项目id_IN IN 体检项目清单.诊疗项目id%TYPE,
	体检类型_IN IN 体检项目清单.体检类型%TYPE,
	执行科室id_IN IN 体检项目清单.执行科室id%TYPE:=NULL,
	采集方式id_IN IN 体检项目清单.采集方式id%TYPE:=NULL,
	检验标本_IN IN 体检项目清单.检验标本%TYPE:=NULL,
	检查部位_IN IN 体检项目清单.检查部位%TYPE:=NULL,
	检查部位id_IN IN 体检项目清单.检查部位id%TYPE:=NULL,
	病人id_IN IN 体检项目清单.病人id%TYPE:=0,
	结算途径_IN IN 体检项目清单.结算途径%TYPE:=1
)
IS
BEGIN
	Insert Into 体检项目清单(ID,登记id,组别名称,诊疗项目id,执行科室id,采集方式id,检验标本,检查部位,检查部位id,病人id,体检类型,结算途径)
	VALUES(体检项目清单_ID.NEXTVAL,登记id_IN,组别名称_IN,诊疗项目id_IN,DECODE(执行科室id_IN,0,NULL,执行科室id_IN),DECODE(采集方式id_IN,0,NULL,采集方式id_IN),检验标本_IN,检查部位_IN,检查部位id_IN,DECODE(病人id_IN,0,NULL,病人id_IN),体检类型_IN,结算途径_IN);

EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检项目清单_INSERT;
/

----------------------------------------------------------------------------
---  体检类型   for   体检登记记录
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检登记记录_体检类型(
	登记id_IN IN 体检登记记录.ID%TYPE
)
IS
	Cursor c_Type is
		SELECT DISTINCT 体检类型 FROM 体检项目清单 WHERE 体检类型 IS NOT NULL AND 登记id=登记id_IN;

	v_体检类型 VARCHAR2(1000);
BEGIN
	For r_Type IN c_Type Loop
		IF INSTR(';'||v_体检类型||';',';'||r_Type.体检类型||';')<=0 THEN
			v_体检类型:=v_体检类型||';'||r_Type.体检类型;		
		END IF;
	end loop;
	IF v_体检类型 IS NOT NULL THEN 
		v_体检类型:=substr(v_体检类型,2,length(v_体检类型)-1);
	END IF;
	UPDATE 体检登记记录 SET 体检类型=v_体检类型 WHERE ID=登记id_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检登记记录_体检类型;
/

----------------------------------------------------------------------------
---  DELETE   for   体检组别项目
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检项目清单_DELETE(
	登记id_IN IN 体检项目清单.登记id%TYPE,
	组别名称_IN		varchar2:=null,
	诊疗项目id_IN	number:=0,
	病人id_IN	number:=0
)
IS
	Cursor c_Items is
		SELECT A.* FROM 体检项目清单 A,体检登记记录 B  WHERE A.登记id=B.ID AND B.ID=登记id_IN AND A.组别名称=组别名称_IN AND A.诊疗项目id=诊疗项目id_IN;

BEGIN

	if 诊疗项目id_IN=0 then
		Delete from 体检项目清单 where 登记id=登记id_IN;
	else
		if 组别名称_IN IS NULL THEN
			Delete from 体检项目清单 where 登记id=登记id_IN AND 组别名称 IS NULL AND 诊疗项目id=诊疗项目id_IN and 病人id=病人id_IN;
		ELSE
			Delete from 体检项目清单 where 登记id=登记id_IN AND 组别名称=组别名称_IN AND 诊疗项目id=诊疗项目id_IN;
		END IF;
	end if;

EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检项目清单_DELETE;
/

----------------------------------------------------------------------------
---  INSERT   for   体检项目清单
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检项目医嘱_INSERT(
	ID_IN IN 体检项目清单.ID%TYPE,
	病人id_IN IN 体检人员档案.病人id%TYPE,
	医嘱id_IN IN 体检项目医嘱.医嘱id%TYPE
)
IS
BEGIN
	INSERT INTO 体检项目医嘱(清单id,病人id,医嘱id)
	SELECT ID,病人id_IN,医嘱id_IN FROM 体检项目清单 WHERE ID=ID_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检项目医嘱_INSERT;
/

----------------------------------------------------------------------------
---  INSERT   for   体检人员档案
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检人员档案_INSERT(
	ID_IN IN 体检人员档案.ID%TYPE,
	登记ID_IN IN 体检人员档案.登记ID%TYPE,
	病人ID_IN IN 体检人员档案.病人ID%TYPE,
	组别名称_IN IN 体检人员档案.组别名称%TYPE
)
IS
BEGIN
	Insert Into 体检人员档案
		(ID,登记ID,病人ID,组别名称)
		VALUES
		(ID_IN,登记ID_IN,病人ID_IN,组别名称_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检人员档案_INSERT;
/
----------------------------------------------------------------------------
---  UPDATE   for   体检人员档案
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检人员档案_UPDATE(
	登记ID_IN IN 体检人员档案.登记ID%TYPE,
	病人ID_IN IN 体检人员档案.病人ID%TYPE,
	组别名称_IN IN 体检人员档案.组别名称%TYPE,		
	原病人ID_IN IN NUMBER:=0
)
IS
BEGIN
	Update 体检人员档案
		Set 登记ID=登记ID_IN,
		    病人ID=病人ID_IN,
		    组别名称=组别名称_IN		    
	WHERE 登记id=登记id_IN AND 病人ID=原病人ID_IN;
EXCEPTION
	WHEN OTHERS THEN
		zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检人员档案_UPDATE;
/
----------------------------------------------------------------------------
---  CLASS   for   体检人员档案
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检人员档案_CLASS(
	登记ID_IN IN 体检人员档案.登记ID%TYPE,
	病人ID_IN IN 体检人员档案.病人ID%TYPE,
	组别名称_IN IN 体检人员档案.组别名称%TYPE
)
IS
BEGIN
	Update 体检人员档案
		Set 组别名称=组别名称_IN		    
	WHERE 登记id=登记id_IN AND 病人ID=病人ID_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检人员档案_CLASS;
/
----------------------------------------------------------------------------
---  DELETE   for   体检人员档案
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_体检人员档案_DELETE(
	登记id_IN IN 体检人员档案.登记ID%TYPE,
	病人id_IN IN 体检人员档案.病人ID%TYPE:=0,
	医嘱作废_IN NUMBER:=0
)
IS
	Cursor c_Advice is	
		SELECT A.ID,A.相关id FROM 病人医嘱记录 A,体检登记记录 B WHERE A.病人id=病人id_IN AND A.医嘱状态 <>4 AND A.挂号单=B.体检号 AND B.ID=登记id_IN;
	
	r_Advice c_Advice%RowType;
	v_Count number(18);

	v_Have		number(1);
	Err_Custom	Exception;
	v_Error		Varchar2(255);
BEGIN
	
	--要作医嘱作废处理
	IF 医嘱作废_IN=1 AND 病人id_IN>0 THEN
		
		For r_Advice IN c_Advice Loop

			Update 病人医嘱发送 Set 执行状态=0,报告id=NULL WHERE 医嘱ID=r_Advice.ID;

			Update 病人费用记录 
				Set 执行状态=0,执行时间=NULL,执行人=NULL
			Where 收费类别 Not IN('5','6','7') 
				AND 医嘱序号=r_Advice.ID
				And (记录性质,NO) IN(
						Select 记录性质,NO From 病人医嘱附费 Where 医嘱id=r_Advice.ID
						Union ALL
						Select 记录性质,NO From 病人医嘱发送 Where 医嘱id=r_Advice.ID);
						
		END LOOP;
		
		DELETE FROM 体检人员档案 WHERE 登记id=登记id_IN AND 病人id=病人id_IN;
		

		For r_Advice IN c_Advice Loop
			IF r_Advice.相关id IS NULL THEN

				--判断是否存在有效的附费,划价单\收费单\记帐单
				v_Have:=0;
				BEGIN
					SELECT 1 INTO v_Have FROM 病人费用记录 
					WHERE  记录状态 IN (0,1) 
						AND (医嘱序号,NO) IN 
							(
							SELECT 医嘱id,NO 
							FROM 病人医嘱附费 
							WHERE 医嘱id IN (
									SELECT ID FROM 病人医嘱记录 
									WHERE ID=r_Advice.ID OR 相关id=r_Advice.ID
									)
							);

				EXCEPTION
					WHEN OTHERS THEN v_Have:=0;
				END;
				
				IF v_Have=1 THEN
					v_Error:='有体检项目还存在附费,请先对附费进行删除或作废！';
				        Raise Err_Custom;
				END IF;
				
				DELETE FROM 病人医嘱附费 WHERE 医嘱id IN (
									SELECT ID FROM 病人医嘱记录 
									WHERE ID=r_Advice.ID OR 相关id=r_Advice.ID);	

				ZL_病人医嘱记录_作废(r_Advice.ID);
			END IF;
		END LOOP;

		
	END IF;

	IF 病人ID_IN=0 THEN
		Delete From 体检人员档案 WHERE 登记id=登记id_IN;
		Delete from 体检项目医嘱 WHERE 清单id in (SELECT ID FROM 体检项目清单 WHERE 登记id=登记id_IN);
	ELSE
		Delete From 体检人员档案 WHERE 登记id=登记id_IN AND 病人id=病人ID_IN;
		Delete from 体检项目医嘱 WHERE  病人id=病人ID_IN AND 清单id in (SELECT ID FROM 体检项目清单 WHERE 登记id=登记id_IN);
	END IF;

	IF 医嘱作废_IN=1 THEN
		--如果移除成员后,没有了人员在体检,自动退到未开始体检状态
		v_Count:=0;
		BEGIN 
			SELECT COUNT(1) INTO v_Count FROM 体检人员档案 WHERE 体检报到=1 AND 登记id=登记id_IN;
		EXCEPTION
			WHEN OTHERS THEN v_Count:=0;
		END;

		IF v_Count=0 THEN
			UPDATE 体检登记记录 SET 体检状态=2 WHERE ID=登记id_IN;
		END IF;		
	END IF;
EXCEPTION
	When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检人员档案_DELETE;
/

----------------------------------------------------------------------------
---  CANCEL   for   体检登记记录
---	取消体检开始
----------------------------------------------------------------------------
CREATE OR REPLACE Procedure ZL_体检登记记录_Cancel(
	体检号_IN		varchar2
) IS
	Cursor c_Advice is
		SELECT ID,相关id FROM 病人医嘱记录 WHERE 病人来源=4 AND 挂号单=体检号_IN AND 医嘱状态<>4;
	Cursor c_Advice2 is
		SELECT ID FROM 病人医嘱记录 WHERE 相关id IS NULL AND 病人来源=4 AND 挂号单=体检号_IN AND 医嘱状态<>4;

	Cursor c_Report is
		SELECT 医嘱id,报告id FROM 病人医嘱发送 WHERE 报告id>0 AND 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE 病人来源=4 AND 挂号单=体检号_IN AND 医嘱状态<>4);

	Cursor c_Person is
		SELECT ID,体检病历id FROM 体检人员档案 WHERE 登记id=(SELECT ID FROM 体检登记记录 WHERE 体检号=体检号_IN);

	r_Row c_Advice%RowType;

	v_Have number(1);
	Err_Custom	Exception;
	v_Error		Varchar2(255);
Begin
	
	DELETE FROM 体检人员结论 WHERE 登记id=(SELECT ID FROM 体检登记记录 WHERE 体检号=体检号_IN);

	For r_Row IN c_Report Loop				
		UPDATE 病人医嘱发送 SET 报告id=NULL WHERE 医嘱id=r_Row.医嘱id;
	END LOOP;

	For r_Row IN c_Report Loop				
		DELETE FROM 病人病历记录 WHERE ID=r_Row.报告id;
	END LOOP;

	For r_Row IN c_Advice Loop

		Update 病人医嘱发送 Set 执行状态=0,报告id=NULL WHERE 医嘱ID=r_Row.ID;

		Update 病人费用记录 
			Set 执行状态=0,执行时间=NULL,执行人=NULL
		Where 收费类别 Not IN('5','6','7') 
			AND 医嘱序号=r_Row.ID
			And (记录性质,NO) IN(
					Select 记录性质,NO From 病人医嘱附费 Where 医嘱id=r_Row.ID
					Union ALL
					Select 记录性质,NO From 病人医嘱发送 Where 医嘱id=r_Row.ID);

		--DELETE FROM 病人医嘱附费 WHERE 医嘱id=r_Row.ID;
	END LOOP;
	
	For r_Row IN c_Person Loop
		UPDATE 体检人员档案 SET 体检状态=1,体检病历id=NULL,复查时间=NULL,体检报到=0 WHERE ID=r_Row.ID;
		DELETE FROM 病人病历记录 WHERE ID=r_Row.体检病历id;
	end loop;

	UPDATE 体检登记记录 SET 体检状态=2 WHERE 体检号=体检号_IN;
	DELETE FROM 体检项目医嘱 WHERE 清单id IN (SELECT ID FROM 体检项目清单 WHERE 登记id=(SELECT ID FROM 体检登记记录 WHERE 体检号=体检号_IN));
	DELETE FROM 体检项目清单 WHERE 病人id>0 AND 登记id=(SELECT ID FROM 体检登记记录 WHERE 体检号=体检号_IN);


	For r_Row IN c_Advice2 Loop		
		--判断是否存在有效的附费,划价单\收费单\记帐单
		v_Have:=0;
		BEGIN
			SELECT 1 INTO v_Have FROM 病人费用记录 
			WHERE  记录状态 IN (0,1) 
				AND (医嘱序号,NO) IN 
					(
					SELECT 医嘱id,NO 
					FROM 病人医嘱附费 
					WHERE 医嘱id IN (
							SELECT ID FROM 病人医嘱记录 
							WHERE ID=r_Row.ID OR 相关id=r_Row.ID
							)
					);

		EXCEPTION
			WHEN OTHERS THEN v_Have:=0;
		END;
		
		IF v_Have=1 THEN
			v_Error:='有体检项目还存在附费,请先对附费进行删除或作废！';
			Raise Err_Custom;
		END IF;

		DELETE FROM 病人医嘱附费 WHERE 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE ID=r_Row.ID OR 相关id=r_Row.ID);

		ZL_病人医嘱记录_作废(r_Row.ID);
	END LOOP;

Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检登记记录_Cancel;
/


CREATE OR REPLACE Procedure ZL_体检登记记录_ItemCancel(
	体检号_IN		varchar2,
	诊疗项目id_IN		number:=0,
	组别名称_IN		varchar2:=NULL,
	病人id_IN		number:=0
) IS
	Cursor c_Items is
		SELECT A.* FROM 体检项目清单 A,体检登记记录 B  WHERE A.登记id=B.ID AND B.体检号=体检号_IN AND A.组别名称=组别名称_IN AND A.诊疗项目id=诊疗项目id_IN;

	Cursor c_ItemsMembers is
		SELECT A.* FROM 体检项目清单 A,体检登记记录 B  WHERE A.登记id=B.ID AND B.体检号=体检号_IN AND A.组别名称 IS NULL AND A.诊疗项目id=诊疗项目id_IN AND A.病人id=病人id_IN;

	Cursor c_ItemsPerson is
		SELECT A.* FROM 体检项目清单 A,体检登记记录 B  WHERE A.登记id=B.ID AND B.体检号=体检号_IN AND A.诊疗项目id=诊疗项目id_IN AND A.病人id=病人id_IN;

	Cursor c_Advice(v_诊疗id number,v_体检号 varchar2,v_检验 number) is
		SELECT DECODE(v_检验,1,相关id,ID) AS ID FROM 病人医嘱记录 
		WHERE 病人来源=4 
			AND 挂号单=v_体检号 
			AND 医嘱状态<>4 
			AND 诊疗项目id=v_诊疗id;

	Cursor c_AdvicePerson(v_诊疗id number,v_体检号 varchar2,v_检验 number) is
		SELECT DECODE(v_检验,1,相关id,ID) AS ID FROM 病人医嘱记录 
		WHERE 病人来源=4 
			AND 挂号单=v_体检号 
			AND 医嘱状态<>4 
			AND 诊疗项目id=v_诊疗id 
			AND 病人id=病人id_IN;

	Cursor c_Advice2(v_医嘱id number) is
		SELECT ID FROM 病人医嘱记录 WHERE 相关id=v_医嘱id or ID=v_医嘱id;

	r_Advice c_Advice%RowType;
	r_AdvicePerson c_AdvicePerson%RowType;
	r_Item c_Items%RowType;
	r_ItemPerson c_ItemsPerson%RowType;

	r_Advice2 c_Advice2%RowType;
	
	v_主诊疗id number(18);
	v_Have number(1);
	v_Flag number(1);
	Err_Custom	Exception;
	v_Error		Varchar2(255);
Begin
	IF 病人id_IN=0 THEN
		For r_Item IN c_Items Loop
			--if r_Item.采集方式id IS NOT NULL AND r_Item.检验标本 IS NOT NULL then
			v_Flag:=0;

			v_主诊疗id:=r_Item.诊疗项目id;
			if r_Item.采集方式id IS NOT NULL then
				v_Flag:=1;
			end if;
			
			For r_Advice IN c_Advice(v_主诊疗id,体检号_IN,v_Flag) Loop
				--找出主医嘱id
				For r_Advice2 IN c_Advice2(r_Advice.ID) Loop

					Update 病人医嘱发送 Set 执行状态=0,报告id=NULL WHERE 医嘱ID=r_Advice2.ID;
					
					Update 病人费用记录 
							Set 执行状态=0,执行时间=NULL,执行人=NULL
					Where 收费类别 Not IN('5','6','7') 
						AND 医嘱序号=r_Advice2.ID
						And (记录性质,NO) IN(
							Select 记录性质,NO From 病人医嘱附费 Where 医嘱id=r_Advice2.ID
							Union ALL
							Select 记录性质,NO From 病人医嘱发送 Where 医嘱id=r_Advice2.ID);
				END LOOP;

				--判断是否存在有效的附费,划价单\收费单\记帐单
				v_Have:=0;
				BEGIN
					SELECT 1 INTO v_Have FROM 病人费用记录 
					WHERE  记录状态 IN (0,1) 
						AND (医嘱序号,NO) IN 
							(
							SELECT 医嘱id,NO 
							FROM 病人医嘱附费 
							WHERE 医嘱id IN (
									SELECT ID FROM 病人医嘱记录 
									WHERE ID=r_Advice.ID OR 相关id=r_Advice.ID
									)
							);

				EXCEPTION
					WHEN OTHERS THEN v_Have:=0;
				END;
				
				IF v_Have=1 THEN
					v_Error:='当前体检项目还存在附费,请先对附费进行删除或作废！';
				        Raise Err_Custom;
				END IF;
				
				DELETE FROM 病人医嘱附费 WHERE 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE ID=r_Advice.ID OR 相关id=r_Advice.ID);

				ZL_病人医嘱记录_作废(r_Advice.ID);
			END LOOP;
		end loop;
	ELSE
		For r_ItemPerson IN c_ItemsMembers Loop
			--if r_Item.采集方式id IS NOT NULL AND r_Item.检验标本 IS NOT NULL then
			
			v_Flag:=0;
			v_主诊疗id:=r_ItemPerson.诊疗项目id;

			if r_ItemPerson.采集方式id IS NOT NULL then
				v_Flag:=1;
			end if;
			
			For r_AdvicePerson IN c_AdvicePerson(v_主诊疗id,体检号_IN,v_Flag) Loop
				--找出主医嘱id
				For r_Advice2 IN c_Advice2(r_AdvicePerson.ID) Loop

					Update 病人医嘱发送 Set 执行状态=0,报告id=NULL WHERE 医嘱ID=r_Advice2.ID;
					
					Update 病人费用记录 
							Set 执行状态=0,执行时间=NULL,执行人=NULL
					Where 收费类别 Not IN('5','6','7') 
						AND 医嘱序号=r_Advice2.ID
						And (记录性质,NO) IN(
							Select 记录性质,NO From 病人医嘱附费 Where 医嘱id=r_Advice2.ID
							Union ALL
							Select 记录性质,NO From 病人医嘱发送 Where 医嘱id=r_Advice2.ID);
				END LOOP;

				--判断是否存在有效的附费,划价单\收费单\记帐单
				v_Have:=0;
				BEGIN
					SELECT 1 INTO v_Have FROM 病人费用记录 
					WHERE  记录状态 IN (0,1) 
						AND (医嘱序号,NO) IN 
							(
							SELECT 医嘱id,NO 
							FROM 病人医嘱附费 
							WHERE 医嘱id IN (
									SELECT ID FROM 病人医嘱记录 
									WHERE ID=r_AdvicePerson.ID OR 相关id=r_AdvicePerson.ID
									)
							);

				EXCEPTION
					WHEN OTHERS THEN v_Have:=0;
				END;
				
				IF v_Have=1 THEN
					v_Error:='当前体检项目还存在附费,请先对附费进行删除或作废！';
				        Raise Err_Custom;
				END IF;
				
				DELETE FROM 病人医嘱附费 WHERE 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE ID=r_AdvicePerson.ID OR 相关id=r_AdvicePerson.ID);
				ZL_病人医嘱记录_作废(r_AdvicePerson.ID);
			END LOOP;
		end loop;		
	END IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检登记记录_ItemCancel;
/

CREATE OR REPLACE Procedure ZL_体检登记记录_Finish(
	体检号_IN		varchar2,
	病人id_IN		number:=0
	--病人id_IN为0时表示此体检号的所有病人
) IS
	v_Count NUMBER(18);
Begin

	IF 病人id_IN=0 THEN
		UPDATE 体检人员档案 A SET A.体检状态=5,
					A.完成时间=SYSDATE,
					A.姓名=(SELECT 姓名 FROM 病人信息 WHERE 病人id=A.病人id),
					A.性别=(SELECT 性别 FROM 病人信息 WHERE 病人id=A.病人id),
					A.年龄=(SELECT 年龄 FROM 病人信息 WHERE 病人id=A.病人id),
					A.婚姻状况=(SELECT 婚姻状况 FROM 病人信息 WHERE 病人id=A.病人id),
					A.联系电话=(SELECT 联系电话 FROM 病人信息 WHERE 病人id=A.病人id),
					A.联系地址=(SELECT 联系地址 FROM 病人信息 WHERE 病人id=A.病人id)					
		WHERE A.登记id=(SELECT ID FROM 体检登记记录 WHERE 体检号=体检号_IN);
		UPDATE 体检登记记录 SET 体检状态=5 WHERE 体检号=体检号_IN;
	ELSE
		UPDATE 体检人员档案 SET 体检状态=5,
					完成时间=SYSDATE,
					姓名=(SELECT 姓名 FROM 病人信息 WHERE 病人id=病人id_IN),
					性别=(SELECT 性别 FROM 病人信息 WHERE 病人id=病人id_IN),
					年龄=(SELECT 年龄 FROM 病人信息 WHERE 病人id=病人id_IN),
					婚姻状况=(SELECT 婚姻状况 FROM 病人信息 WHERE 病人id=病人id_IN),
					联系电话=(SELECT 联系电话 FROM 病人信息 WHERE 病人id=病人id_IN),
					联系地址=(SELECT 联系地址 FROM 病人信息 WHERE 病人id=病人id_IN)
		WHERE 病人id=病人id_IN 
			AND 登记id=(SELECT ID FROM 体检登记记录 WHERE 体检号=体检号_IN);
		
		v_Count:=0;
		BEGIN
			SELECT NVL(COUNT(1),0) INTO v_Count FROM 体检人员档案 WHERE 体检状态<5 AND 登记id=(SELECT ID FROM 体检登记记录 WHERE 体检号=体检号_IN);
		EXCEPTION
			WHEN OTHERS THEN v_Count:=0;
		END;

		IF v_Count<=0 THEN
			UPDATE 体检登记记录 SET 体检状态=5,完成时间=SYSDATE WHERE 体检号=体检号_IN;		
		END IF;
	END IF;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检登记记录_Finish;
/

CREATE OR REPLACE Procedure ZL_体检登记记录_CancelFinish(
	体检号_IN		varchar2,
	病人id_IN		number:=0
	--病人id_IN为0时表示此体检号的所有病人
) IS
	v_No varchar2(30);
	v_Temp			Varchar2(255);
	v_人员编号		人员表.编号%Type;
	v_人员姓名		人员表.姓名%Type;
Begin
	
	--当前操作人员
	v_Temp:=zl_Identity;
	v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
	v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
	v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
	v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

	IF 病人id_IN=0 THEN
		UPDATE 体检人员档案 SET 体检状态=4,完成时间=NULL WHERE 登记id=(SELECT ID FROM 体检登记记录 WHERE 体检号=体检号_IN);

	ELSE
		UPDATE 体检人员档案 SET 体检状态=4,完成时间=NULL WHERE 病人id=病人id_IN AND 登记id=(SELECT ID FROM 体检登记记录 WHERE 体检号=体检号_IN);
	END IF;

	UPDATE 体检登记记录 SET 体检状态=4,完成时间=NULL WHERE 体检号=体检号_IN;		

	--取消结算作废
	BEGIN
		zl_病人结帐记录_Delete(v_No,v_人员编号,v_人员姓名,0);
	EXCEPTION
		WHEN OTHERS THEN v_No:='';
	END;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检登记记录_CancelFinish;
/

CREATE OR REPLACE Procedure ZL_体检登记记录_单项填写(
	病历id_IN		病人病历所见单.病历id%TYPE,
	所见项id_IN		病人病历所见单.所见项id%TYPE,
	所见内容_IN		病人病历所见单.所见内容%TYPE
) IS
Begin
	UPDATE 病人病历所见单 SET 所见内容=所见内容_IN WHERE 病历id=病历id_IN AND 所见项id=所见项id_IN;	
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检登记记录_单项填写;
/

CREATE OR REPLACE Procedure ZL_体检人员档案_总结(
	登记id_IN		体检人员档案.登记id%TYPE,
	病人id_IN		体检人员档案.病人id%TYPE,
	体检病历id_IN		体检人员档案.体检病历id%TYPE
) IS
Begin
	UPDATE 体检人员档案 SET 体检病历id=体检病历id_IN WHERE 登记id=登记id_IN AND 病人id=病人id_IN;	
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检人员档案_总结;
/

CREATE OR REPLACE Procedure ZL_体检人员档案_复查(
	登记id_IN		体检人员档案.登记id%TYPE,
	病人id_IN		体检人员档案.病人id%TYPE,
	复查时间_IN		体检人员档案.复查时间%TYPE
) IS
Begin
	UPDATE 体检人员档案 SET 复查时间=复查时间_IN WHERE 登记id=登记id_IN AND 病人id=病人id_IN;	
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检人员档案_复查;
/

CREATE OR REPLACE PROCEDURE ZL_体检人员结论_REFRESH(
	登记id_IN IN 体检人员结论.登记id%TYPE,
	病人id_IN IN 体检人员结论.病人id%TYPE:=0
)
IS
	Cursor c_PersonAll is	
		SELECT 病人id FROM 体检人员档案 WHERE 登记id=登记id_IN;

	Cursor c_Person is	
		SELECT 病人id FROM 体检人员档案 WHERE 登记id=登记id_IN AND 病人id=病人id_IN;

	Cursor c_List(v_病人id number) is	
		SELECT DISTINCT 执行科室id 
		FROM (       
			SELECT 执行科室id 
			FROM 体检项目清单 A,体检人员档案 B
			WHERE B.登记id=登记id_IN 
			      AND B.病人ID=v_病人id
			      AND A.登记ID=B.登记ID
			      AND A.组别名称=B.组别名称
			      AND A.病人id IS NULL   
			UNION ALL 
			SELECT 执行科室id 
			FROM 体检项目清单 A
			WHERE A.登记id=登记id_IN
			      AND A.病人ID=v_病人id      
		     );

	v_Have NUMBER(1);

BEGIN
	IF 病人id_IN=0 THEN
		For r_Person IN c_PersonAll Loop

			For r_Row IN c_List(r_Person.病人id) Loop
				
				v_Have:=0;
				BEGIN
					SELECT 1 INTO v_Have FROM 体检人员结论 WHERE 登记id=登记id_IN AND 病人id=r_Person.病人id AND 科室id=r_Row.执行科室id;
				EXCEPTION
					WHEN OTHERS THEN v_Have:=0;
				END;

				IF v_Have=0 THEN
					--没有,则增加
					INSERT INTO 体检人员结论(登记id,病人id,科室id,结论id) VALUES (登记id_IN,r_Person.病人id,r_Row.执行科室id,NULL);
				END IF;

			END LOOP;
			
			DELETE FROM 体检人员结论 
			WHERE 登记id=登记id_IN 
				AND 病人id=r_Person.病人id 
				AND 科室id NOT IN (
						SELECT DISTINCT 执行科室id 
						FROM (       
							SELECT 执行科室id 
							FROM 体检项目清单 A,体检人员档案 B
							WHERE B.登记id=登记id_IN 
							      AND B.病人ID=r_Person.病人id
							      AND A.登记ID=B.登记ID
							      AND A.组别名称=B.组别名称
							      AND A.病人id IS NULL   
							UNION ALL 
							SELECT 执行科室id 
							FROM 体检项目清单 A
							WHERE A.登记id=登记id_IN
							      AND A.病人ID=r_Person.病人id
						));
		end loop;
	ELSE
		For r_Row IN c_List(病人id_IN) Loop
			
			v_Have:=0;
			BEGIN
				SELECT 1 INTO v_Have FROM 体检人员结论 WHERE 登记id=登记id_IN AND 病人id=病人id_IN AND 科室id=r_Row.执行科室id;
			EXCEPTION
				WHEN OTHERS THEN v_Have:=0;
			END;

			IF v_Have=0 THEN
				--没有,则增加
				INSERT INTO 体检人员结论(登记id,病人id,科室id,结论id) VALUES (登记id_IN,病人id_IN,r_Row.执行科室id,NULL);
			END IF;

		END LOOP;
		
		DELETE FROM 体检人员结论 
		WHERE 登记id=登记id_IN 
			AND 病人id=病人id_IN
			AND 科室id NOT IN (
					SELECT DISTINCT 执行科室id 
					FROM (       
						SELECT 执行科室id 
						FROM 体检项目清单 A,体检人员档案 B
						WHERE B.登记id=登记id_IN 
						      AND B.病人ID=病人id_IN
						      AND A.登记ID=B.登记ID
						      AND A.组别名称=B.组别名称
						      AND A.病人id IS NULL   
						UNION ALL 
						SELECT 执行科室id 
						FROM 体检项目清单 A
						WHERE A.登记id=登记id_IN
						      AND A.病人ID=病人id_IN
					));
	END IF;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检人员结论_REFRESH;
/

CREATE OR REPLACE PROCEDURE ZL_体检人员结论_UPDATE(
	登记id_IN IN 体检人员结论.登记id%TYPE,
	病人id_IN IN 体检人员结论.病人id%TYPE,
	科室id_IN IN 体检人员结论.科室id%TYPE,
	结论id_IN IN 体检人员结论.结论id%TYPE
)
IS
BEGIN
	UPDATE 体检人员结论 SET 结论id=结论id_IN WHERE 登记id=登记id_IN AND 病人id=病人id_IN AND 科室id=科室id_IN;
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_体检人员结论_UPDATE;
/

CREATE OR REPLACE Procedure ZL_体检人员档案_报到(
	登记id_IN		体检人员档案.登记id%TYPE,
	病人id_IN		体检人员档案.病人id%TYPE,
	体检报到_IN		体检人员档案.体检报到%TYPE
) IS
Begin
	UPDATE 体检人员档案 SET 体检报到=体检报到_IN WHERE 登记id=登记id_IN AND 病人id=病人id_IN;	
	IF 体检报到_IN=1 THEN
		ZL_体检人员结论_REFRESH(登记id_IN,病人id_IN);
	END IF;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检人员档案_报到;
/

CREATE OR REPLACE Procedure ZL_体检结算记录_INSERT(
	ID_IN		IN	体检结算记录.ID%TYPE,
	合约单位id_IN	IN	体检结算记录.合约单位id%TYPE,
	结算id_IN	IN	体检结算记录.结算id%TYPE,
	结算金额_IN	IN	体检结算记录.结算金额%TYPE,
	结算部门id_IN	IN 	体检结算记录.结算部门id%TYPE
) IS
Begin
	INSERT INTO 体检结算记录(ID,记录状态,合约单位id,结算id,结算金额,结算部门id)
	VALUES (ID_IN,1,合约单位id_IN,结算id_IN,结算金额_IN,结算部门id_IN);
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检结算记录_INSERT;
/

CREATE OR REPLACE Procedure ZL_体检结算清单_INSERT(
	结算id_IN	IN	体检结算清单.结算id%TYPE,
	登记id_IN	IN	体检结算清单.登记id%TYPE
) IS
Begin
	INSERT INTO 体检结算清单(结算id,登记id)	VALUES (结算id_IN,登记id_IN);
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_体检结算清单_INSERT;
/

CREATE OR REPLACE Procedure zl_体检结算记录_Cancel(
	结算id_IN	IN	病人结帐记录.ID%TYPE
) IS
	Cursor c_Items is
		SELECT A.* FROM 病人结帐记录 A WHERE A.ID=结算id_IN;

	v_Temp			Varchar2(255);
	v_人员编号		人员表.编号%Type;
	v_人员姓名		人员表.姓名%Type;
Begin

	--当前操作人员
	v_Temp:=zl_Identity;
	v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
	v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
	v_人员编号:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
	v_人员姓名:=Substr(v_Temp,Instr(v_Temp,',')+1);

	For r_Row IN c_Items Loop
		zl_病人结帐记录_Delete(r_Row.No,v_人员编号,v_人员姓名,0);
	end loop;

	UPDATE 体检结算记录 SET 记录状态=2 WHERE 结算id=结算id_IN;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_体检结算记录_Cancel;
/

--报表：ZL1_BILL_1861/体检员项目清单
Insert Into zlReports(ID,编号,名称,说明,密码,W,H,纸张,纸向,进纸,打印机,动态纸张,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_BILL_1861','体检员项目清单','项目清单','Zn!t_jgnq1<S~aimD0[_',11904,16832,9,1,15,NULL,0,1,100,1861,'项目清单',Sysdate,Sysdate);
Insert Into zlRPTFMTs(报表ID,序号,说明,图样) Values(zlReports_ID.CurrVal,1,'体检员项目清单1',0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,NULL,0,'任意表1',11,'团体:[体检项目清单_数据.团体]',NULL,450,1425,2610,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,NULL,0,'任意表1',12,'体检项目清单',NULL,4478,660,2700,435,0,0,1,'宋体',22,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,NULL,0,'任意表1',13,'体检单:[体检项目清单_数据.体检号]',NULL,8235,1665,2970,180,0,2,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,NULL,0,NULL,0,'体检项目清单_数据',NULL,450,1950,10755,8460,255,0,0,'宋体',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[体检项目清单_数据.类别]','4^255^类别',0,0,1140,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[体检项目清单_数据.名称]','4^255^名称',0,0,6345,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-3,2,NULL,NULL,'[体检项目清单_数据.体检科室]','4^255^体检科室',0,0,2625,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,NULL,0,'任意表1',11,'姓名:[体检项目清单_数据.姓名]',NULL,450,1670,2610,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'体检项目清单_数据','类别,200|名称,200|体检科室,200|体检号,200|姓名,200|团体,200',USER||'.体检项目清单,'||USER||'.诊疗项目目录,'||USER||'.部门表,'||USER||'.诊疗项目类别,'||USER||'.体检项目医嘱,'||USER||'.体检登记记录,'||USER||'.合约单位,'||USER||'.病人信息',0);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'Select D.名称 AS 类别,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'       B.名称,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'       C.名称 AS 体检科室,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'       F.体检号,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'       H.姓名,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'       G.名称 AS 团体');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'From 体检项目清单 A,诊疗项目目录 B,部门表 C,诊疗项目类别 D,体检项目医嘱 E,体检登记记录 F,合约单位 G,病人信息 H');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'WHERE A.诊疗项目ID=B.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'      AND C.ID=A.执行科室ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'      AND B.类别=D.编码');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'      AND E.清单ID=A.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'      AND A.登记ID=F.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,'      AND E.病人ID=H.病人ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,'      AND F.合约单位id=G.ID(+)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'      AND A.登记ID=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'      AND E.病人ID=[1]');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'登记ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,1,'病人ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);

--报表：ZL1_BILL_1861_2/体检报告书
Insert Into zlReports(ID,编号,名称,说明,密码,W,H,纸张,纸向,进纸,打印机,动态纸张,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_BILL_1861_2','体检报告书','报告书打印','Zn:kA}6x|;0Tm=|sW*Q]',11904,16832,9,1,15,NULL,0,1,100,1861,'报告书打印',Sysdate,Sysdate);
Insert Into zlRPTFMTs(报表ID,序号,说明,图样) Values(zlReports_ID.CurrVal,1,'体检报告书1',0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,NULL,0,NULL,0,'第[页号]页 共[页数]页',NULL,345,15960,1965,180,0,1,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,NULL,0,'任意表1',11,'团体:[体检人员档案_数据.团体]',NULL,360,1220,2610,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,NULL,0,'任意表1',11,'姓名:[体检人员档案_数据.姓名]',NULL,360,1505,2610,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,NULL,0,'任意表1',12,'体检报告单',NULL,4860,675,2250,435,0,0,1,'宋体',22,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,NULL,0,'任意表1',13,'体检时间:[体检人员档案_数据.体检时间]',NULL,8280,1485,3330,180,0,1,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,NULL,0,NULL,0,'病人病历所见单_数据',NULL,360,1785,11250,12570,345,0,0,'宋体',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,8421504,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[病人病历所见单_数据.项目]','4^345^项目',0,0,4455,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[病人病历所见单_数据.结果]','4^345^结果',0,0,4110,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-3,2,NULL,NULL,'[病人病历所见单_数据.参考]','4^345^参考',0,0,1860,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-4,3,NULL,NULL,'[病人病历所见单_数据.提示]','4^345^提示',0,0,765,0,0,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表2',4,NULL,0,'任意表1',1,'体检总检_数据',NULL,360,14355,11250,1545,570,0,1,'宋体',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,8421504,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[体检总检_数据.项目]','1^345^总检',0,0,1260,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[体检总检_数据.结果]','1^345^总检',0,0,9855,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'体检人员档案_数据','姓名,200|团体,200|体检时间,200',USER||'.体检人员档案,'||USER||'.病人信息,'||USER||'.体检登记记录,'||USER||'.合约单位',0);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'SELECT 	B.姓名, ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'	D.名称 AS 团体,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'	TO_CHAR(C.体检时间,'||CHR(39)||'yyyy-mm-dd'||CHR(39)||') AS 体检时间');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'FROM 	体检人员档案 A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'	病人信息 B,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'	体检登记记录 C,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'	合约单位 D');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'WHERE 	A.病人id=B.病人id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'	AND C.ID=A.登记id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'	AND C.合约单位id=D.ID(+)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'	AND C.ID=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'	AND A.病人id=[1]');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'登记ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,1,'病人ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'体检总检_数据','项目,200|结果,200',USER||'.病人病历所见单,'||USER||'.诊治所见项目,'||USER||'.病人病历文本段,'||USER||'.体检人员档案,'||USER||'.病人病历内容',0);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'SELECT 	'||CHR(39)||'    '||CHR(39)||'||项目 AS 项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'	结果');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'FROM (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'select ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'       X.排列序号,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'       X1.内序号,              ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'       X1.项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'       X1.结果  ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'from      ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'     体检人员档案 A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'     病人病历内容 X,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'     (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,'     select  ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,'             A.病历ID,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'             A.控件号 AS 内序号,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'             B.中文名 AS 项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,17,'             DECODE(A.所见内容,NULL,NULL,A.所见内容||'||CHR(39)||' '||CHR(39)||'||B.单位) AS 结果');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,18,'      from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,19,'        病人病历所见单 A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,20,'        诊治所见项目 B');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,21,'      where A.所见项id=B.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,22,'            and 所见项id>0');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,23,'      ) X1');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,24,'where X.病历记录id=A.体检病历ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,25,'      AND X.ID=X1.病历id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,26,'      AND A.病人ID=[1]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,27,'      AND A.登记ID=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,28,'      AND X.元素类型=2       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,29,NULL);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,30,'union all ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,31,'      ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,32,'select     ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,33,'       X.排列序号,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,34,'       X1.内序号,        ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,35,'       X.标题文本 AS 项目,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,36,'       X1.结果  ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,37,'from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,38,'     体检人员档案 A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,39,'     病人病历内容 X,     ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,40,'      (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,41,'      select 病历id,0 AS 内序号,'||CHR(39)||''||CHR(39)||' AS 项目,内容 AS 结果 from 病人病历文本段   ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,42,'      ) X1');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,43,'where X.病历记录id=A.体检病历ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,44,'      AND X.ID=X1.病历id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,45,'      AND A.病人ID=[1]      ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,46,'      AND A.登记ID=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,47,'      AND X.元素类型 in (4,-5)          ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,48,') ORDER BY  排列序号,内序号 ');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'登记ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,1,'病人ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'病人病历所见单_数据','项目,200|结果,200|参考,200|提示,200|排序1,200|排序2,200|排序3,139',USER||'.体检项目医嘱,'||USER||'.病人医嘱记录,'||USER||'.病人医嘱发送,'||USER||'.病人病历记录,'||USER||'.体检项目清单,'||USER||'.病人病历所见单,'||USER||'.诊治所见项目,'||USER||'.病人病历文本段,'||USER||'.病人病历内容,'||USER||'.部门表,'||USER||'.诊疗项目目录,'||USER||'.体检人员结论',0);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'SELECT * FROM (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'  SELECT ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'         '||CHR(39)||'        '||CHR(39)||'||项目 AS 项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'         结果||DECODE(标志,NULL,'||CHR(39)||''||CHR(39)||',DECODE(SUBSTR(标志,3,100),'||CHR(39)||'正常'||CHR(39)||','||CHR(39)||''||CHR(39)||','||CHR(39)||'异常'||CHR(39)||','||CHR(39)||'(+)'||CHR(39)||','||CHR(39)||'偏低'||CHR(39)||','||CHR(39)||'↓'||CHR(39)||','||CHR(39)||'偏高'||CHR(39)||','||CHR(39)||'↑'||CHR(39)||')) AS 结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'         参考,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'         DECODE(标志,NULL,'||CHR(39)||''||CHR(39)||',SUBSTR(标志,3,100)) AS 提示,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'         体检科室 AS 排序1,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'         体检项目 AS 排序2,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'         3 AS 排序3       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'  FROM (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'  SELECT ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'         U.名称 AS 体检科室,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,'         T.名称 AS 体检项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,'         R.项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'         R.结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'         DECODE(SIGN(INSTR(R.标志参考,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(R.标志参考,1,INSTR(R.标志参考,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')-1),'||CHR(39)||''||CHR(39)||') AS 标志,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,17,'         DECODE(SIGN(INSTR(R.标志参考,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(R.标志参考,INSTR(R.标志参考,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')+1,1000),'||CHR(39)||''||CHR(39)||') AS 参考');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,18,'  FROM (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,19,'  select ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,20,'         A.执行部门ID,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,21,'         A.体检项目id,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,22,'         A.ID,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,23,'         X.排列序号,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,24,'         X1.内序号,              ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,25,'         X1.项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,26,'         DECODE(SIGN(INSTR(X1.结果,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(X1.结果,1,INSTR(X1.结果,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')-1),X1.结果) AS 结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,27,'         DECODE(SIGN(INSTR(X1.结果,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(X1.结果,INSTR(X1.结果,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')+1,1000),'||CHR(39)||''||CHR(39)||') AS 标志参考       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,28,'  from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,29,'       (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,30,'       select DISTINCT A1.医嘱ID,A3.执行部门ID,A4.ID,A5.诊疗项目id AS 体检项目id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,31,'        from 体检项目医嘱 A1,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,32,'             病人医嘱记录 A2,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,33,'             病人医嘱发送 A3,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,34,'             病人病历记录 A4,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,35,'             体检项目清单 A5');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,36,'        where A1.病人id=[1]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,37,'      AND A5.登记id=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,38,'              AND (A1.医嘱ID=A2.ID OR A1.医嘱ID=A2.相关id)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,39,'              AND A3.医嘱ID=A2.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,40,'              AND A4.ID=A3.报告ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,41,'              AND A5.ID=A1.清单ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,42,'       ) A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,43,'       病人病历内容 X,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,44,'       (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,45,'       select  ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,46,'               A.病历ID,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,47,'               A.控件号 AS 内序号,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,48,'               B.中文名 AS 项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,49,'               DECODE(A.所见内容,NULL,NULL,A.所见内容||'||CHR(39)||' '||CHR(39)||'||B.单位) AS 结果');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,50,'        from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,51,'          病人病历所见单 A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,52,'          诊治所见项目 B');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,53,'        where A.所见项id=B.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,54,'              and 所见项id>0');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,55,'        ) X1');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,56,'  where X.病历记录id=A.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,57,'        AND X.ID=X1.病历id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,58,'  union all     ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,59,'  select ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,60,'         A.执行部门ID,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,61,'         A.体检项目id,      ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,62,'         A.ID,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,63,'         X.排列序号,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,64,'         X1.内序号,        ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,65,'         X.标题文本 AS 项目,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,66,'         X1.结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,67,'         '||CHR(39)||''||CHR(39)||' AS 标志参考   ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,68,'  from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,69,'       (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,70,'       select DISTINCT A1.医嘱ID,A3.执行部门ID,A4.ID,A5.诊疗项目id AS 体检项目id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,71,'        from 体检项目医嘱 A1,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,72,'             病人医嘱记录 A2,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,73,'             病人医嘱发送 A3,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,74,'             病人病历记录 A4,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,75,'             体检项目清单 A5');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,76,'        where A1.病人id=[1]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,77,'      AND A5.登记id=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,78,'              AND (A1.医嘱ID=A2.ID OR A1.医嘱ID=A2.相关id)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,79,'              AND A3.医嘱ID=A2.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,80,'              AND A4.ID=A3.报告ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,81,'              AND A5.ID=A1.清单ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,82,'       ) A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,83,'       病人病历内容 X,     ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,84,'        (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,85,'        select 病历id,0 AS 内序号,'||CHR(39)||''||CHR(39)||' AS 项目,内容 AS 结果 from 病人病历文本段   ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,86,'        ) X1');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,87,'  where X.病历记录id=A.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,88,'        AND X.ID=X1.病历id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,89,'        AND X.元素类型 IN (0,4,-5)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,90,'        AND X.元素编码<>'||CHR(39)||'000009'||CHR(39));
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,91,'  ) R,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,92,'  部门表 U,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,93,'  诊疗项目目录 T');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,94,'  WHERE R.执行部门id=U.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,95,'        AND R.体检项目id=T.ID)              ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,96,NULL);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,97,'UNION ALL ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,98,'	');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,99,'	SELECT ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,100,'	       '||CHR(39)||'    '||CHR(39)||'||T.名称 AS 项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,101,'	       '||CHR(39)||'检查时间:'||CHR(39)||'||TO_CHAR(R.书写日期,'||CHR(39)||'yyyy-mm-dd hh24:mi'||CHR(39)||') AS 结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,102,'	       '||CHR(39)||'检查医生:'||CHR(39)||'||R.书写人 AS 参考,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,103,'		'||CHR(39)||''||CHR(39)||' AS 提示,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,104,'	       U.名称 AS 排序1,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,105,'	       T.名称 AS 排序2,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,106,'	       2 AS 排序3       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,107,'	FROM (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,108,'	SELECT DISTINCT');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,109,'	       A.执行部门ID,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,110,'	       A.体检项目id,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,111,'	       A.ID,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,112,'	       A.书写人,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,113,'	       A.书写日期                 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,114,'	from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,115,'	     (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,116,'	     SELECT DISTINCT A1.医嘱ID,A3.执行部门ID,A4.ID,A5.诊疗项目id AS 体检项目id,A4.书写人,A4.书写日期');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,117,'	      from 体检项目医嘱 A1,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,118,'	           病人医嘱记录 A2,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,119,'	           病人医嘱发送 A3,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,120,'	           病人病历记录 A4,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,121,'	           体检项目清单 A5');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,122,'	      where A1.病人id=[1]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,123,'		      AND A5.登记id=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,124,'	            AND (A1.医嘱ID=A2.ID OR A1.医嘱ID=A2.相关id)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,125,'	            AND A3.医嘱ID=A2.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,126,'	            AND A4.ID=A3.报告ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,127,'	            AND A5.ID=A1.清单ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,128,'	     ) A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,129,'	     病人病历内容 X');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,130,'	WHERE X.病历记录id=A.ID     ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,131,') R,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,132,'部门表 U,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,133,'诊疗项目目录 T');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,134,'WHERE R.执行部门id=U.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,135,'      AND R.体检项目id=T.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,136,'union all ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,137,'SELECT ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,138,'       体检科室 AS 项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,139,'       '||CHR(39)||''||CHR(39)||' AS 结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,140,'       '||CHR(39)||''||CHR(39)||' AS 参考,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,141,'	'||CHR(39)||''||CHR(39)||' AS 提示,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,142,'       体检科室 AS 排序1,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,143,'       '||CHR(39)||' '||CHR(39)||' AS 排序2,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,144,'       1 AS 排序3       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,145,'FROM (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,146,'     select DISTINCT U.名称 AS 体检科室');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,147,'      from 体检项目医嘱 A1,           ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,148,'           体检项目清单 A5,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,149,'           部门表 U');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,150,'      where A1.病人id=[1]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,151,'	      AND A5.登记id=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,152,'            AND A5.执行科室ID=U.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,153,'            AND A5.ID=A1.清单ID     ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,154,'     ) R');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,155,NULL);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,156,'union all ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,157,'SELECT ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,158,'       '||CHR(39)||'    小结'||CHR(39)||' AS 项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,159,'       DECODE(R.书写日期,NULL,'||CHR(39)||''||CHR(39)||','||CHR(39)||'小结时间:'||CHR(39)||'||TO_CHAR(R.书写日期,'||CHR(39)||'yyyy-mm-dd hh24:mi'||CHR(39)||')) AS 结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,160,'       DECODE(R.书写人,NULL,'||CHR(39)||''||CHR(39)||','||CHR(39)||'小结医生:'||CHR(39)||'||R.书写人) AS 参考,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,161,'	'||CHR(39)||''||CHR(39)||' as 提示,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,162,'       体检科室 AS 排序1,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,163,'       '||CHR(39)||'座座座座座座座座座座座座座座座座'||CHR(39)||' AS 排序2,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,164,'       4 AS 排序3       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,165,'FROM (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,166,'     select DISTINCT U.名称 AS 体检科室,A4.书写人,A4.书写日期');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,167,'      from 体检人员结论 A1,           ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,168,'           病人病历记录 A4,		');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,169,'           部门表 U');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,170,'      where A1.病人id=[1]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,171,'	      AND A1.登记id=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,172,'            AND A1.科室ID=U.ID ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,173,'	    AND A1.结论id=A4.ID(+)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,174,'     ) R         ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,175,'UNION ALL ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,176,'  SELECT ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,177,'         '||CHR(39)||'        '||CHR(39)||'||项目 AS 项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,178,'         结果||DECODE(标志,NULL,'||CHR(39)||''||CHR(39)||',DECODE(SUBSTR(标志,3,100),'||CHR(39)||'正常'||CHR(39)||','||CHR(39)||''||CHR(39)||','||CHR(39)||'异常'||CHR(39)||','||CHR(39)||'(+)'||CHR(39)||','||CHR(39)||'偏低'||CHR(39)||','||CHR(39)||'↓'||CHR(39)||','||CHR(39)||'偏高'||CHR(39)||','||CHR(39)||'↑'||CHR(39)||')) AS 结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,179,'         参考,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,180,'	'||CHR(39)||''||CHR(39)||' as 提示,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,181,'         体检科室 AS 排序1,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,182,'         '||CHR(39)||'座座座座座座座座座座座座座座座座'||CHR(39)||' AS 排序2,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,183,'         5 AS 排序3       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,184,'  FROM (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,185,'  SELECT ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,186,'         U.名称 AS 体检科室,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,187,'         R.项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,188,'         R.结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,189,'         DECODE(SIGN(INSTR(R.标志参考,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(R.标志参考,1,INSTR(R.标志参考,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')-1),'||CHR(39)||''||CHR(39)||') AS 标志,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,190,'         DECODE(SIGN(INSTR(R.标志参考,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(R.标志参考,INSTR(R.标志参考,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')+1,1000),'||CHR(39)||''||CHR(39)||') AS 参考');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,191,'  FROM (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,192,'  select ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,193,'         A.执行部门ID,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,194,'         A.ID,              ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,195,'         X1.项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,196,'         DECODE(SIGN(INSTR(X1.结果,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(X1.结果,1,INSTR(X1.结果,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')-1),X1.结果) AS 结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,197,'         DECODE(SIGN(INSTR(X1.结果,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')),1,SUBSTR(X1.结果,INSTR(X1.结果,'||CHR(39)||''||CHR(39)||''||CHR(39)||''||CHR(39)||')+1,1000),'||CHR(39)||''||CHR(39)||') AS 标志参考       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,198,'  from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,199,'       (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,200,'       select 科室id AS 执行部门ID,结论id AS ID  from 体检人员结论 WHERE 病人id=[1] AND 登记id=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,201,'       ) A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,202,'       病人病历内容 X,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,203,'       (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,204,'       select  ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,205,'               A.病历ID,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,206,'               A.控件号 AS 内序号,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,207,'               B.中文名 AS 项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,208,'               DECODE(A.所见内容,NULL,NULL,A.所见内容||'||CHR(39)||' '||CHR(39)||'||B.单位) AS 结果');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,209,'        from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,210,'          病人病历所见单 A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,211,'          诊治所见项目 B');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,212,'        where A.所见项id=B.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,213,'              and 所见项id>0');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,214,'        ) X1');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,215,'  where X.病历记录id=A.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,216,'        AND X.ID=X1.病历id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,217,'  union all     ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,218,'  select ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,219,'         A.执行部门ID,            ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,220,'         A.ID,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,221,'         X.标题文本 AS 项目,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,222,'         X1.结果,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,223,'         '||CHR(39)||''||CHR(39)||' AS 标志参考   ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,224,'  from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,225,'       (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,226,'       select 科室id AS 执行部门ID,结论id AS ID  from 体检人员结论 WHERE 病人id=[1] AND 登记id=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,227,'       ) A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,228,'       病人病历内容 X,     ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,229,'        (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,230,'        select 病历id,0 AS 内序号,'||CHR(39)||''||CHR(39)||' AS 项目,内容 AS 结果 from 病人病历文本段   ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,231,'        ) X1');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,232,'  where X.病历记录id=A.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,233,'        AND X.ID=X1.病历id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,234,'        AND X.元素类型 IN (0,4,-5)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,235,'  ) R,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,236,'  部门表 U');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,237,'  WHERE R.执行部门id=U.ID)       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,238,')       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,239,'ORDER BY 排序1,排序2,排序3');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'登记ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,1,'病人ID',1,NULL,0,NULL,NULL,NULL,NULL,NULL,NULL);

--报表：ZL1_BILL_1862/团体体检结算收据
Insert Into zlReports(ID,编号,名称,说明,密码,W,H,纸张,纸向,进纸,打印机,动态纸张,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_BILL_1862','团体体检结算收据','收据打印','Zp,fXhpso<0TfvnmI<BD',12191,5443,256,1,7,'Star AR-3200+',0,1,100,1862,'收据打印',Sysdate,Sysdate);
Insert Into zlRPTFMTs(报表ID,序号,说明,图样) Values(zlReports_ID.CurrVal,1,'团体体检收据',0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,NULL,0,NULL,0,'[收费汇总.票据号][收费汇总.来源]',NULL,555,780,3645,225,0,0,1,'宋体',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签5',2,NULL,0,NULL,0,'[收费汇总.日期]',NULL,555,4320,1350,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,NULL,0,NULL,0,'[收费汇总.姓名]',NULL,915,1125,1710,225,0,0,1,'宋体',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,NULL,0,NULL,0,'[收费汇总.大写]',NULL,1395,3945,1710,225,0,0,1,'宋体',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,NULL,0,NULL,0,'[收费汇总.合计]',NULL,1650,3555,1710,225,0,2,1,'宋体',11,0,0,0,0,16777215,0,NULL,'0.00',NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签6',2,NULL,0,NULL,0,'[收费汇总.操作员姓名]',NULL,2310,4320,1890,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签7',2,NULL,0,NULL,0,'[收费汇总.票据号][收费汇总.来源]',NULL,4410,765,3645,225,0,0,1,'宋体',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签8',2,NULL,0,NULL,0,'[收费汇总.日期]',NULL,4415,4310,1350,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签9',2,NULL,0,NULL,0,'[收费汇总.姓名]',NULL,4770,1110,1710,225,0,0,1,'宋体',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签10',2,NULL,0,NULL,0,'[收费汇总.大写]',NULL,5255,3935,1710,225,0,0,1,'宋体',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签11',2,NULL,0,NULL,0,'[收费汇总.合计]',NULL,5505,3540,1710,225,0,2,1,'宋体',11,0,0,0,0,16777215,0,NULL,'0.00',NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签12',2,NULL,0,NULL,0,'[收费汇总.操作员姓名]',NULL,6165,4305,1890,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签13',2,NULL,0,NULL,0,'[收费汇总.票据号][收费汇总.来源]',NULL,8265,765,3645,225,0,0,1,'宋体',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签14',2,NULL,0,NULL,0,'[收费汇总.日期]',NULL,8275,4315,1350,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签15',2,NULL,0,NULL,0,'[收费汇总.姓名]',NULL,8625,1110,1710,225,0,0,1,'宋体',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签16',2,NULL,0,NULL,0,'[收费汇总.大写]',NULL,9115,3940,1710,225,0,0,1,'宋体',11,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签17',2,NULL,0,NULL,0,'[收费汇总.合计]',NULL,9375,3555,1710,225,0,2,1,'宋体',11,0,0,0,0,16777215,0,NULL,'0.00',NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签18',2,NULL,0,NULL,0,'[收费汇总.操作员姓名]',NULL,10020,4305,1890,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,NULL,0,NULL,0,'收费明细',NULL,556,1938,2858,1485,465,0,0,'宋体',11,0,0,0,0,16777215,1,NULL,NULL,NULL,1,16777215,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[收费明细.项目]','4^30^#',0,0,1380,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[收费明细.金额]','4^30^#',0,0,1425,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表2',4,NULL,0,NULL,0,'收费明细',NULL,4416,1928,2858,1485,465,0,0,'宋体',11,0,0,0,0,16777215,1,NULL,NULL,NULL,1,16777215,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[收费明细.项目]','4^30^#',300,300,1380,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[收费明细.金额]','4^30^#',300,300,1425,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表3',4,NULL,0,NULL,0,'收费明细',NULL,8276,1933,2858,1485,465,0,0,'宋体',11,0,0,0,0,16777215,1,NULL,NULL,NULL,1,16777215,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[收费明细.项目]','4^30^#',600,600,1380,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[收费明细.金额]','4^30^#',600,600,1425,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'收费明细','项目,200|金额,200',USER||'.病人费用记录',0);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'--因为有部分退费重打,因此不管记录状态;多单据收费时,参数传入了多个NO');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'Select ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'	收据费目 as 项目,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'	Ltrim(To_Char(Sum(Nvl(结帐金额,0)),'||CHR(39)||'999999990.00'||CHR(39)||')) as 金额');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'From 病人费用记录');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'Where Mod(记录性质,10)=2 And 结帐id=[0] and 记录状态<>0');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'Group by 收据费目');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'Having Sum(Nvl(结帐金额,0))<>0');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'Order by 收据费目');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'结帐id',1,'0',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'收费汇总','页号,139|票据号,200|NO,200|姓名,200|来源,200|操作员编号,200|操作员姓名,200|日期,200|合计,200|大写,200',USER||'.病人费用记录,'||USER||'.系统参数表,'||USER||'.体检结算记录,'||USER||'.合约单位,'||USER||'.票据打印内容,'||USER||'.病人结帐记录,'||USER||'.票据使用明细',0);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'--票据号没有固定分配到具体的收费行次上,因此根据收据费目排序汇总');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'--汇总收据费目时,因为有部分退费重打,因此不管记录状态,且先按收据费目排序,再按票据号排序');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'--支持多张单据收费统一打印票据的方式(参数传入了多个NO),多张单据的票据打印内容ID相同。');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'--子表A：根据收据行次设置及单据中的收据费目,返回每张票据的汇总金额');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'--子表B：返回单据中的病人信息,因为多单据收费时可以单独修改,因此取最后有效记录');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'--子表C：返回单据对应分配票据号');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'Select ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'	A.页号,C.号码 As 票据号,B.NO,B.姓名,B.来源,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'	B.操作员编号,B.操作员姓名,To_Char(B.登记时间,'||CHR(39)||'YYYY-MM-DD'||CHR(39)||') As 日期,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'	Ltrim(To_Char(A.金额,'||CHR(39)||'9999999.00'||CHR(39)||')) As 合计,zlUppMoney(A.金额) As 大写');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'From (Select Ceil(A.序号/B.收据行次) As 页号,Sum(A.金额) As 金额');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'		From (Select Rownum As 序号,项目,金额');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,'				From (');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,'				Select 收据费目 As 项目,Sum(结帐金额) As 金额');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'				From 病人费用记录');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'				Where 结帐id=[0] And Mod(记录性质,10)=2 and 记录状态<>0');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,17,'				Group By 收据费目');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,18,'					)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,19,'			) A,(Select Nvl(Nvl(参数值,缺省值),3) as 收据行次 From 系统参数表 Where 参数号=4) B');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,20,'		Group By Ceil(A.序号/B.收据行次)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,21,'	) A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,22,'	(Select ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,23,'		Min(NO)||DeCode(Max(NO),Min(NO),Null,'||CHR(39)||'-'||CHR(39)||' || Max(NO)) As NO,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,24,'		Max(C.名称) as 姓名,Max(A.操作员编号) as 操作员编号,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,25,'		Max(A.操作员姓名) as 操作员姓名,Max(A.登记时间) As 登记时间,Decode(Max(A.门诊标志),2,'||CHR(39)||'(住院收费)'||CHR(39)||',NULL) as 来源');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,26,'		From 病人费用记录 A,体检结算记录 B,合约单位 C');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,27,'		Where Mod(A.记录性质,10)=2 And A.记录状态 In (1,3) And A.序号=1 And A.结帐id=[0] AND A.结帐id=B.结算id AND C.ID=B.合约单位id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,28,'	) B,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,29,'	(Select Rownum As 页号,A.号码');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,30,'		From 票据使用明细 A,票据打印内容 B');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,31,'		Where B.数据性质=1 And B.ID=(Select Max(A.ID) From 票据打印内容 A,病人结帐记录 B Where A.数据性质=1 And A.NO=B.NO AND B.ID=[0])');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,32,'			And A.打印ID=B.ID And A.票种=1 And A.性质=1');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,33,'	) C');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,34,'Where A.页号=C.页号(+)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,35,'Order By C.号码');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'结帐id',1,'0',0,NULL,NULL,NULL,NULL,NULL,NULL);

--报表：ZL1_REPORT_1876/科室工作量统计
Insert Into zlReports(ID,编号,名称,说明,密码,W,H,纸张,纸向,进纸,打印机,动态纸张,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1876','科室工作量统计','统计一段时间内体检科室体检工作量的情况','Ew?vNub{b-<XqdldZ2ZZ',11904,16832,9,1,15,NULL,0,0,100,1876,'基本',Sysdate,Sysdate);
Insert Into zlRPTFMTs(报表ID,序号,说明,图样) Values(zlReports_ID.CurrVal,1,'科室工作量统计1',0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,NULL,0,'汇总表1',11,'统计范围:[=开始日期]至[=结束日期]',NULL,675,1650,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,NULL,0,'汇总表1',12,'科室工作量统计',NULL,4260,810,2625,375,0,1,1,'宋体',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'汇总表1',5,NULL,0,NULL,0,'科室工作量',NULL,675,1935,9795,7950,255,0,0,'宋体',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,7,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'体检科室',NULL,0,0,1605,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,7,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'项目名称',NULL,0,0,6750,0,0,0,0,'宋体',0,0,0,0,0,0,0,'项目名称',NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,9,zlRPTItems_ID.CurrVal-3,0,NULL,NULL,'人次',NULL,0,0,735,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'科室工作量','体检科室,200|项目名称,200|人次,139',USER||'.体检登记记录,'||USER||'.体检项目清单,'||USER||'.体检项目医嘱,'||USER||'.病人医嘱记录,'||USER||'.病人医嘱发送,'||USER||'.诊疗项目目录,'||USER||'.部门表',1);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'select G.名称 AS 体检科室,F.名称 AS 项目名称, COUNT(1) AS 人次');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'from 体检登记记录 A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'     体检项目清单 B,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'     体检项目医嘱 C,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'     病人医嘱记录 D,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'     病人医嘱发送 E,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'     诊疗项目目录 F,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'     部门表 G');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'WHERE A.ID=B.登记ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'      AND C.清单ID=B.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'      AND (D.ID=C.医嘱id OR D.相关ID=C.医嘱id)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'      AND D.诊疗类别 IN ('||CHR(39)||'C'||CHR(39)||','||CHR(39)||'D'||CHR(39)||')');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,'      AND E.医嘱ID=D.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,'      AND E.报告ID>0 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'	AND A.体检状态=5');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'      AND F.ID=D.诊疗项目id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,17,'      AND G.ID=E.执行部门ID	');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,18,'      AND A.体检时间 BETWEEN [0] and [1]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,19,'GROUP BY G.名称,F.名称');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'开始日期',2,CHR(38)||'前一月日期',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,1,'结束日期',2,CHR(38)||'当前日期',0,NULL,NULL,NULL,NULL,NULL,NULL);

--报表：ZL1_REPORT_1877/医生工作量统计
Insert Into zlReports(ID,编号,名称,说明,密码,W,H,纸张,纸向,进纸,打印机,动态纸张,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1877','医生工作量统计','统计一段时间内体检医生体检工作量的情况','Ww?vNtbib-<XqddZ2ZZ',11904,16832,9,1,15,NULL,0,0,100,1877,'基本',Sysdate,Sysdate);
Insert Into zlRPTFMTs(报表ID,序号,说明,图样) Values(zlReports_ID.CurrVal,1,'医生工作量统计1',0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,NULL,0,'汇总表1',11,'统计范围:[=开始日期]至[=结束日期]',NULL,675,1650,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,NULL,0,'汇总表1',12,'医生工作量统计',NULL,4245,810,2655,360,0,1,1,'宋体',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'汇总表1',5,NULL,0,NULL,0,'医生工作量',NULL,675,1935,9795,7950,255,0,0,'宋体',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,7,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'医生',NULL,0,0,1605,0,255,0,0,'宋体',0,0,0,0,0,0,0,'医生',NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,7,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'项目名称',NULL,0,0,6750,0,0,0,0,'宋体',0,0,0,0,0,0,0,'项目名称',NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,9,zlRPTItems_ID.CurrVal-3,0,NULL,NULL,'人次',NULL,0,0,735,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'医生工作量','医生,200|项目名称,200|人次,139',USER||'.体检登记记录,'||USER||'.体检项目清单,'||USER||'.体检项目医嘱,'||USER||'.病人医嘱记录,'||USER||'.病人医嘱发送,'||USER||'.诊疗项目目录,'||USER||'.病人病历记录',1);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'select G.书写人 AS 医生,F.名称 AS 项目名称, COUNT(1) AS 人次');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'from 体检登记记录 A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'     体检项目清单 B,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'     体检项目医嘱 C,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'     病人医嘱记录 D,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'     病人医嘱发送 E,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'     诊疗项目目录 F,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'     病人病历记录 G');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'WHERE A.ID=B.登记ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'      AND C.清单ID=B.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'      AND (D.ID=C.医嘱id OR D.相关ID=C.医嘱id)');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'      AND D.诊疗类别 IN ('||CHR(39)||'C'||CHR(39)||','||CHR(39)||'D'||CHR(39)||')');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,'      AND E.医嘱ID=D.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,'      AND E.报告ID>0 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'      AND F.ID=D.诊疗项目id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'      AND G.ID=E.报告ID	');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,17,'	AND A.体检状态=5');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,18,'      AND A.体检时间 BETWEEN [0] and [1]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,19,'GROUP BY G.书写人,F.名称');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'开始日期',2,CHR(38)||'前一月日期',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,1,'结束日期',2,CHR(38)||'当前日期',0,NULL,NULL,NULL,NULL,NULL,NULL);

--报表：ZL1_REPORT_1878/体检人数统计分析
Insert Into zlReports(ID,编号,名称,说明,密码,W,H,纸张,纸向,进纸,打印机,动态纸张,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1878','体检人数统计分析','统计一段时间内各个团体体检人数情况','Zn*Venhe 4GqdooI"D]',11904,16832,9,1,15,NULL,0,0,100,1878,'基本',Sysdate,Sysdate);
Insert Into zlRPTFMTs(报表ID,序号,说明,图样) Values(zlReports_ID.CurrVal,1,'体检人数统计1',0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,NULL,0,'任意表1',11,'统计范围:[=开始日期]至[=结束日期]',NULL,330,1965,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签2',2,NULL,0,'任意表1',12,'体检人数统计',NULL,4550,1080,2250,375,0,1,1,'宋体',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,NULL,0,NULL,0,'体检登记记录_数据',NULL,330,2265,10690,7200,255,0,0,'宋体',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[体检登记记录_数据.团体名称]','1^255^团体名称|1^255^团体名称',0,0,3735,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[体检登记记录_数据.男性人数]','4^255^人数|4^255^男性',0,0,630,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-3,2,NULL,NULL,'[体检登记记录_数据.女性人数]','4^255^人数|4^255^女性',0,0,585,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-4,3,NULL,NULL,'[体检登记记录_数据.人数]','4^255^人数|4^255^合计',0,0,810,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-5,4,NULL,NULL,'[体检登记记录_数据.已检男性人数]','4^255^已检人数|4^255^男性',0,0,630,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-6,5,NULL,NULL,'[体检登记记录_数据.已检女性人数]','4^255^已检人数|4^255^女性',0,0,675,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-7,6,NULL,NULL,'[体检登记记录_数据.已检人数]','4^255^已检人数|4^255^合计',0,0,795,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-8,7,NULL,NULL,'[体检登记记录_数据.未检男性人数]','4^255^未检人数|4^255^男性',0,0,705,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-9,8,NULL,NULL,'[体检登记记录_数据.未检女性人数]','4^255^未检人数|4^255^女性',0,0,690,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-10,9,NULL,NULL,'[体检登记记录_数据.未检人数]','4^255^未检人数|4^255^合计',0,0,1005,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,'SUM',1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'体检登记记录_数据','团体名称,200|男性人数,200|女性人数,200|人数,200|已检男性人数,200|已检女性人数,200|已检人数,200|未检男性人数,200|未检女性人数,200|未检人数,200',USER||'.体检登记记录,'||USER||'.体检人员档案,'||USER||'.合约单位',0);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'SELECT ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'	团体名称,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'	DECODE(男性人数,0,NULL,男性人数) AS 男性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'	DECODE(女性人数,0,NULL,女性人数) AS 女性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'	DECODE(人数,0,NULL,人数) AS 人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'	DECODE(已检男性人数,0,NULL,已检男性人数) AS 已检男性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'	DECODE(已检女性人数,0,NULL,已检女性人数) AS 已检女性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'	DECODE(已检人数,0,NULL,已检人数) AS 已检人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'	DECODE(未检男性人数,0,NULL,未检男性人数) AS 未检男性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'	DECODE(未检女性人数,0,NULL,未检女性人数) AS 未检女性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'	DECODE(未检人数,0,NULL,未检人数) AS 未检人数              ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'FROM ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,'(');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,'SELECT B.名称 AS 团体名称,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'       A.男性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'       A.女性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,17,'       nvl(A.男性人数,0)+nvl(A.女性人数,0) AS 人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,18,'       A.已检男性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,19,'       A.已检女性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,20,'       nvl(A.已检男性人数,0)+nvl(A.已检女性人数,0) AS 已检人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,21,'       nvl(A.男性人数,0)-nvl(A.已检男性人数,0) AS 未检男性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,22,'       nvl(A.女性人数,0)-nvl(A.已检女性人数,0) AS 未检女性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,23,'       (nvl(A.男性人数,0)-nvl(A.已检男性人数,0))+(nvl(A.女性人数,0)-nvl(A.已检女性人数,0)) AS 未检人数              ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,24,'FROM ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,25,'(');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,26,'select A.合约单位id,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,27,'       SUM(DECODE(sign(instr(B.性别,'||CHR(39)||'女'||CHR(39)||')-0),1,0,1)) AS 男性人数,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,28,'       SUM(DECODE(sign(instr(B.性别,'||CHR(39)||'女'||CHR(39)||')-0),1,1,0)) AS 女性人数,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,29,'       SUM(DECODE(sign(0 - NVL(B.体检病历ID,0)),-1, DECODE(SIGN(instr(B.性别,'||CHR(39)||'女'||CHR(39)||')-0),1,0,1),0)) AS 已检男性人数,       ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,30,'       SUM(DECODE(sign(0 - NVL(B.体检病历ID,0)),-1, DECODE(SIGN(instr(B.性别,'||CHR(39)||'女'||CHR(39)||')-0),1,1,0),0)) AS 已检女性人数');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,31,'from 体检登记记录 A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,32,'     体检人员档案 B     ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,33,'WHERE A.ID=B.登记ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,34,'      AND A.合约单位id>0');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,35,'	AND A.体检状态=5');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,36,'      AND A.体检时间 BETWEEN [0] AND [1]+1-1/24/60/60');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,37,'GROUP BY A.合约单位id');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,38,') A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,39,'合约单位 B');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,40,'WHERE A.合约单位id=B.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,41,')');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'开始日期',2,CHR(38)||'前一月日期',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,1,'结束日期',2,CHR(38)||'当前日期',0,NULL,NULL,NULL,NULL,NULL,NULL);

--报表：ZL1_REPORT_1879/复查人员清单
Insert Into zlReports(ID,编号,名称,说明,密码,W,H,纸张,纸向,进纸,打印机,动态纸张,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_REPORT_1879','复查人员清单','统计一段时间内需要复查的人员','Hg*uSjnsc37PcmznL,PM',11904,16832,9,1,15,NULL,0,0,100,1879,'基本',Sysdate,Sysdate);
Insert Into zlRPTFMTs(报表ID,序号,说明,图样) Values(zlReports_ID.CurrVal,1,'复查人员清单1',0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,NULL,0,'任意表1',11,'复查日期:[=开始日期]至[=结束日期]',NULL,600,1755,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,NULL,0,'任意表1',12,'复查人员清单',NULL,4855,1125,2250,360,0,1,1,'宋体',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,NULL,0,NULL,0,'数据源',NULL,600,2040,10760,5865,255,0,0,'宋体',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[数据源.姓名]','4^255^姓名',0,0,1005,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[数据源.体检号]','4^255^体检号',0,0,1005,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-3,2,NULL,NULL,'[数据源.体检团体]','4^255^体检团体',0,0,5130,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-4,3,NULL,NULL,'[数据源.体检时间]','4^255^体检时间',0,0,1785,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-5,4,NULL,NULL,'[数据源.复查时间]','4^255^复查时间',0,0,1290,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'数据源','姓名,200|体检号,131|体检团体,200|体检时间,200|复查时间,200',USER||'.体检登记记录,'||USER||'.体检人员档案,'||USER||'.合约单位',0);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'select B.姓名,A.体检号,F.名称 AS 体检团体,TO_CHAR(A.体检时间,'||CHR(39)||'yyyy-mm-dd'||CHR(39)||') AS 体检时间,to_char(B.复查时间,'||CHR(39)||'yyyy-mm-dd'||CHR(39)||') as 复查时间');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'from 体检登记记录 A,体检人员档案 B,合约单位 F');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'where A.合约单位id=F.ID	');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'      AND A.体检状态=5');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'      AND A.ID=B.登记ID	      ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'      AND B.复查时间 IS NOT NULL');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'      AND B.复查时间 BETWEEN [0] and [1]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'开始日期',2,CHR(38)||'当前日期',0,NULL,'select id,decode(上级id,null,-1,上级id) as 上级id,
	名称,编码 from 合约单位
where 末级<>1
Start With 上级id is null
Connect By prior id=上级id','select * from 合约单位
where (编码 like '||CHR(39)||'%[*]%'||CHR(39)||'
or 名称 like '||CHR(39)||'%[*]%'||CHR(39)||'
or 简码 like '||CHR(39)||'%[*]%'||CHR(39)||') AND 末级=1','ID,131,'||CHR(38)||'R|上级ID,139,|名称,200,'||CHR(38)||'S|编码,200,','ID,131,'||CHR(38)||'B|上级ID,131,'||CHR(38)||'R|编码,200,'||CHR(38)||'S|名称,200,'||CHR(38)||'S'||CHR(38)||'D|简码,200,'||CHR(38)||'S|末级,131,|地址,200,|电话,200,|开户银行,200,|帐号,200,|联系人,200,|建档时间,135,|撤档时间,135,|电子邮件,200,|说明,200,',USER||'.合约单位|'||USER||'.合约单位');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,1,'结束日期',2,CHR(38)||'下一月日期',0,NULL,NULL,NULL,NULL,NULL,NULL);

--报表：ZL1_SUB_1875_1/团体体检结果分析(疾病)
Insert Into zlReports(ID,编号,名称,说明,密码,W,H,纸张,纸向,进纸,打印机,动态纸张,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_SUB_1875_1','团体体检结果分析(疾病)',NULL,'Zp,fI`z?<,6'||CHR(39)||''||CHR(38)||'pkq[0\L',11904,16832,9,1,15,NULL,0,0,100,NULL,NULL,Sysdate,NULL);
Insert Into zlRPTFMTs(报表ID,序号,说明,图样) Values(zlReports_ID.CurrVal,1,'团体体检结果分析(疾病)1',0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,NULL,0,NULL,0,'体检时间:[=开始时间]至[=结束时间]',NULL,585,1755,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,NULL,0,NULL,0,'体检团体:[数据源.体检团体]',NULL,585,2040,4230,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,NULL,0,'任意表1',12,'团体体检结果分析(疾病)',NULL,3712,1125,4140,360,0,1,1,'宋体',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,NULL,0,NULL,0,'数据源',NULL,615,2325,10335,6465,255,0,0,'宋体',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[数据源.疾病名称]','1^255^疾病名称(人数)',0,0,5865,0,255,0,1,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[数据源.名单]','4^255^名单',0,0,1740,0,255,1,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'数据源','体检团体,200|疾病名称,200|名单,200',USER||'.体检登记记录,'||USER||'.体检人员档案,'||USER||'.病人病历内容,'||USER||'.病人诊断记录,'||USER||'.合约单位',0);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'select 体检团体,疾病名称||'||CHR(39)||'('||CHR(39)||'||TO_CHAR(人数)||'||CHR(39)||')'||CHR(39)||' AS 疾病名称,名单 from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'(');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'select D.诊断描述,D.诊断描述 AS 疾病名称,B.姓名 AS 名单,F.名称 AS 体检团体');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'from 体检登记记录 A,体检人员档案 B,病人病历内容 C,病人诊断记录 D,合约单位 F');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'where A.合约单位id=F.ID	');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'      AND F.ID=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'      AND A.体检状态=5');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'      AND A.ID=B.登记ID	');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'      AND B.体检病历ID>0');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'      AND C.病历记录ID=B.体检病历ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'      AND C.元素类型=4');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'      AND D.病历ID=C.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,'      AND A.体检时间 BETWEEN [1] and [2]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,') A,');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,15,'(');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,16,'select D.诊断描述,COUNT(1) AS 人数');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,17,'from 体检登记记录 A,体检人员档案 B,病人病历内容 C,病人诊断记录 D,合约单位 E');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,18,'where A.合约单位id=E.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,19,'      AND E.ID=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,20,'      AND A.体检状态=5');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,21,'      AND A.ID=B.登记ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,22,'      AND B.体检病历ID>0');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,23,'      AND C.病历记录ID=B.体检病历ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,24,'      AND C.元素类型=4');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,25,'      AND D.病历ID=C.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,26,'      AND A.体检时间 BETWEEN [1] and [2]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,27,'GROUP BY D.诊断描述      ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,28,') B      ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,29,'WHERE A.诊断描述=B.诊断描述 ');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'体检团体',1,'选择器定义…',0,NULL,NULL,'select * from 合约单位
where (编码 like '||CHR(39)||'%[*]%'||CHR(39)||'
or 名称 like '||CHR(39)||'%[*]%'||CHR(39)||'
or 简码 like '||CHR(39)||'%[*]%'||CHR(39)||') AND 末级=1',NULL,'ID,131,'||CHR(38)||'B|上级ID,131,|编码,200,'||CHR(38)||'S|名称,200,'||CHR(38)||'S'||CHR(38)||'D|简码,200,'||CHR(38)||'S|末级,131,|地址,200,|电话,200,|开户银行,200,|帐号,200,|联系人,200,|建档时间,135,|撤档时间,135,|电子邮件,200,|说明,200,',USER||'.合约单位|');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,1,'开始时间',2,CHR(38)||'前一周日期',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,2,'结束时间',2,CHR(38)||'当前日期',0,NULL,NULL,'select * from 合约单位
where (编码 like '||CHR(39)||'%[*]%'||CHR(39)||'
or 名称 like '||CHR(39)||'%[*]%'||CHR(39)||'
or 简码 like '||CHR(39)||'%[*]%'||CHR(39)||') AND 末级=1',NULL,'ID,131,'||CHR(38)||'B|上级ID,131,|编码,200,'||CHR(38)||'S|名称,200,'||CHR(38)||'S'||CHR(38)||'D|简码,200,'||CHR(38)||'S|末级,131,|地址,200,'||CHR(38)||'S|电话,200,|开户银行,200,|帐号,200,|联系人,200,|建档时间,135,|撤档时间,135,|电子邮件,200,|说明,200,',USER||'.合约单位|');

--报表：ZL1_SUB_1875_2/团体体检结果分析(人)
Insert Into zlReports(ID,编号,名称,说明,密码,W,H,纸张,纸向,进纸,打印机,动态纸张,票据,系统,程序ID,功能,修改时间,发布时间) Values(zlReports_ID.NextVal,'ZL1_SUB_1875_2','团体体检结果分析(人)',NULL,'Zp,fI`y?$:6'||CHR(39)||'8sfpE:BZ',11904,16832,9,1,15,NULL,0,0,100,NULL,NULL,Sysdate,NULL);
Insert Into zlRPTFMTs(报表ID,序号,说明,图样) Values(zlReports_ID.CurrVal,1,'团体体检结果分析(人)1',0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签3',2,NULL,0,NULL,0,'体检时间:[=开始时间]至[=结束时间]',NULL,585,1755,2970,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签4',2,NULL,0,NULL,0,'体检团体:[数据源.体检团体]',NULL,585,2040,4230,180,0,0,1,'宋体',9,0,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'标签1',2,NULL,0,'任意表1',12,'团体体检结果分析(人)',NULL,3942,1125,3780,360,0,1,1,'宋体',18,1,0,0,0,16777215,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'任意表1',4,NULL,0,NULL,0,'数据源',NULL,615,2355,10435,5550,255,0,0,'宋体',9,0,0,0,0,16777215,1,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-1,0,NULL,NULL,'[数据源.姓名]','4^255^姓名',0,0,1785,0,255,0,1,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTItems(ID,报表ID,格式号,名称,类型,上级ID,序号,参照,性质,内容,表头,X,Y,W,H,行高,对齐,自调,字体,字号,粗体,斜体,下线,前景,背景,边框,排序,格式,汇总,分栏,网格,系统) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,NULL,6,zlRPTItems_ID.CurrVal-2,1,NULL,NULL,'[数据源.疾病名称]','4^255^疾病名称',0,0,7320,0,255,0,0,'宋体',0,0,0,0,0,0,0,NULL,NULL,NULL,1,0,0);
Insert Into zlRPTDatas(ID,报表ID,名称,字段,对象,类型) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'数据源','体检团体,200|姓名,200|疾病名称,200',USER||'.体检登记记录,'||USER||'.体检人员档案,'||USER||'.病人病历内容,'||USER||'.病人诊断记录,'||USER||'.合约单位',0);
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,1,'select 体检团体,姓名||'||CHR(39)||'('||CHR(39)||'||性别||'||CHR(39)||')'||CHR(39)||' AS 姓名,疾病名称 from ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,2,'(');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,3,'select D.诊断描述 AS 疾病名称,B.姓名,B.性别,F.名称 AS 体检团体');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,4,'from 体检登记记录 A,体检人员档案 B,病人病历内容 C,病人诊断记录 D,合约单位 F');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,5,'where A.合约单位id=F.ID	');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,6,'      AND F.ID=[0]');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,7,'      AND A.体检状态=5');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,8,'      AND A.ID=B.登记ID	');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,9,'      AND B.体检病历ID>0');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,10,'      AND C.病历记录ID=B.体检病历ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,11,'      AND C.元素类型=4');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,12,'      AND D.病历ID=C.ID');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,13,'      AND A.体检时间 BETWEEN [1] and [2]+1-1/24/60/60 ');
Insert Into zlRPTSQLs(源ID,行号,内容) Values(zlRPTDatas_ID.CurrVal,14,') A');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,0,'体检团体',1,'选择器定义…',0,NULL,NULL,'select * from 合约单位
where (编码 like '||CHR(39)||'%[*]%'||CHR(39)||'
or 名称 like '||CHR(39)||'%[*]%'||CHR(39)||'
or 简码 like '||CHR(39)||'%[*]%'||CHR(39)||') AND 末级=1',NULL,'ID,131,'||CHR(38)||'B|上级ID,131,|编码,200,'||CHR(38)||'S|名称,200,'||CHR(38)||'S'||CHR(38)||'D|简码,200,'||CHR(38)||'S|末级,131,|地址,200,|电话,200,|开户银行,200,|帐号,200,|联系人,200,|建档时间,135,|撤档时间,135,|电子邮件,200,|说明,200,',USER||'.合约单位|');
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,1,'开始时间',2,CHR(38)||'前一周日期',0,NULL,NULL,NULL,NULL,NULL,NULL);
Insert Into zlRPTPars(源ID,组名,序号,名称,类型,缺省值,格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象) Values(zlRPTDatas_ID.CurrVal,NULL,2,'结束时间',2,CHR(38)||'当前日期',0,NULL,NULL,'select * from 合约单位
where (编码 like '||CHR(39)||'%[*]%'||CHR(39)||'
or 名称 like '||CHR(39)||'%[*]%'||CHR(39)||'
or 简码 like '||CHR(39)||'%[*]%'||CHR(39)||') AND 末级=1',NULL,'ID,131,'||CHR(38)||'B|上级ID,131,|编码,200,'||CHR(38)||'S|名称,200,'||CHR(38)||'S'||CHR(38)||'D|简码,200,'||CHR(38)||'S|末级,131,|地址,200,'||CHR(38)||'S|电话,200,|开户银行,200,|帐号,200,|联系人,200,|建档时间,135,|撤档时间,135,|电子邮件,200,|说明,200,',USER||'.合约单位|');


--报表组：ZL1_GROUP_1875/团体体检结果分析
Insert Into zlRPTGroups(ID,编号,名称,说明,系统,程序ID,发布时间) Values(zlRPTGroups_ID.NextVal,'ZL1_GROUP_1875','团体体检结果分析','统计一段体检时间范围内团体体检结果分析统计情况',100,1875,Sysdate);
Insert Into zlRPTSubs(组ID,报表ID,序号,功能) Select zlRPTGroups_ID.CurrVal,ID,1,'团体体检结果分析(疾病)' From zlReports Where Upper(编号)=Upper('ZL1_SUB_1875_1') And 系统=100;
Insert Into zlRPTSubs(组ID,报表ID,序号,功能) Select zlRPTGroups_ID.CurrVal,ID,2,'团体体检结果分析(人)' From zlReports Where Upper(编号)=Upper('ZL1_SUB_1875_2') And 系统=100;

--报表：ZL1_BILL_1861/体检员项目清单
insert into zlProgFuncs(系统,序号,功能) values (100,1861,'项目清单');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'项目清单',USER,'体检项目清单','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'项目清单',USER,'诊疗项目目录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'项目清单',USER,'部门表','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'项目清单',USER,'诊疗项目类别','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'项目清单',USER,'体检项目医嘱','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'项目清单',USER,'体检登记记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'项目清单',USER,'合约单位','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'项目清单',USER,'病人信息','SELECT');

--报表：ZL1_BILL_1861_2/体检报告书
insert into zlProgFuncs(系统,序号,功能) values (100,1861,'报告书打印');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'体检人员档案','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'病人信息','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'体检登记记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'合约单位','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'病人病历所见单','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'诊治所见项目','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'病人病历文本段','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'病人病历内容','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'体检项目医嘱','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'病人医嘱记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'病人医嘱发送','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'病人病历记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'体检项目清单','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'部门表','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'诊疗项目目录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1861,'报告书打印',USER,'体检人员结论','SELECT');

--报表：ZL1_BILL_1862/团体体检结算收据
insert into zlProgFuncs(系统,序号,功能) values (100,1862,'收据打印');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1862,'收据打印',USER,'系统参数表','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1862,'收据打印',USER,'体检结算记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1862,'收据打印',USER,'合约单位','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1862,'收据打印',USER,'票据打印内容','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1862,'收据打印',USER,'病人结帐记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1862,'收据打印',USER,'票据使用明细','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1862,'收据打印',USER,'病人费用记录','SELECT');

--报表：ZL1_REPORT_1876/科室工作量统计
insert into zlPrograms(序号,标题,说明,系统,部件) values(1876,'科室工作量统计','统计一段时间内体检科室体检工作量的情况',100,'zl9Report');
insert into zlProgFuncs(系统,序号,功能) values (100,1876,'基本');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1876,'基本',USER,'体检登记记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1876,'基本',USER,'体检项目清单','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1876,'基本',USER,'体检项目医嘱','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1876,'基本',USER,'病人医嘱记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1876,'基本',USER,'病人医嘱发送','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1876,'基本',USER,'诊疗项目目录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1876,'基本',USER,'部门表','SELECT');
insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'科室工作量统计','科室工作量统计',NULL,105,'统计一段时间内体检科室体检工作量的情况',100,1876 From zlMenus Where 系统=100 And 组别='缺省' And 标题='体检管理系统' And 模块 is NULL;

--报表：ZL1_REPORT_1877/医生工作量统计
insert into zlPrograms(序号,标题,说明,系统,部件) values(1877,'医生工作量统计','统计一段时间内体检医生体检工作量的情况',100,'zl9Report');
insert into zlProgFuncs(系统,序号,功能) values (100,1877,'基本');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1877,'基本',USER,'体检登记记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1877,'基本',USER,'体检项目清单','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1877,'基本',USER,'体检项目医嘱','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1877,'基本',USER,'病人医嘱记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1877,'基本',USER,'病人医嘱发送','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1877,'基本',USER,'诊疗项目目录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1877,'基本',USER,'病人病历记录','SELECT');
insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'医生工作量统计','医生工作量统计',NULL,105,'统计一段时间内体检医生体检工作量的情况',100,1877 From zlMenus Where 系统=100 And 组别='缺省' And 标题='体检管理系统' And 模块 is NULL;

--报表：ZL1_REPORT_1878/体检人数统计分析
insert into zlPrograms(序号,标题,说明,系统,部件) values(1878,'体检人数统计分析','统计一段时间内各个团体体检人数情况',100,'zl9Report');
insert into zlProgFuncs(系统,序号,功能) values (100,1878,'基本');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1878,'基本',USER,'体检登记记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1878,'基本',USER,'体检人员档案','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1878,'基本',USER,'合约单位','SELECT');
insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'体检人数统计分析','体检人数统计分析',NULL,105,'统计一段时间内各个团体体检人数情况',100,1878 From zlMenus Where 系统=100 And 组别='缺省' And 标题='体检管理系统' And 模块 is NULL;

--报表：ZL1_REPORT_1879/复查人员清单
insert into zlPrograms(序号,标题,说明,系统,部件) values(1879,'复查人员清单','统计一段时间内需要复查的人员',100,'zl9Report');
insert into zlProgFuncs(系统,序号,功能) values (100,1879,'基本');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1879,'基本',USER,'体检登记记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1879,'基本',USER,'体检人员档案','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1879,'基本',USER,'合约单位','SELECT');
insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'复查人员清单','复查人员清单',NULL,105,'统计一段时间内需要复查的人员',100,1879 From zlMenus Where 系统=100 And 组别='缺省' And 标题='体检管理系统' And 模块 is NULL;

--报表组：ZL1_GROUP_1875/团体体检结果分析
insert into zlPrograms(序号,标题,说明,系统,部件) values(1875,'团体体检结果分析','统计一段体检时间范围内团体体检结果分析统计情况',100,'zl9Report');
insert into zlProgFuncs(系统,序号,功能) values (100,1875,'团体体检结果分析(疾病)');
insert into zlProgFuncs(系统,序号,功能) values (100,1875,'团体体检结果分析(人)');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1875,'团体体检结果分析(疾病)',USER,'体检登记记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1875,'团体体检结果分析(疾病)',USER,'体检人员档案','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1875,'团体体检结果分析(疾病)',USER,'病人病历内容','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1875,'团体体检结果分析(疾病)',USER,'病人诊断记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1875,'团体体检结果分析(疾病)',USER,'合约单位','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1875,'团体体检结果分析(人)',USER,'体检登记记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1875,'团体体检结果分析(人)',USER,'体检人员档案','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1875,'团体体检结果分析(人)',USER,'病人病历内容','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1875,'团体体检结果分析(人)',USER,'病人诊断记录','SELECT');
insert into zlProgPrivs(系统,序号,功能,所有者,对象,权限) values (100,1875,'团体体检结果分析(人)',USER,'合约单位','SELECT');
insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'团体体检结果分析','团体体检结果分析(疾病)',NULL,105,'统计一段体检时间范围内团体体检结果分析统计情况',100,1875 From zlMenus Where 系统=100 And 组别='缺省' And 标题='体检管理系统' And 模块 is NULL;

commit;

