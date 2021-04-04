--zlComponent
--Insert Into zlComponent(部件,名称,主版本,次版本,附版本,系统) Values('zl9Transfusion','门诊输液注射部件',10,15,0,100);

--zlPrograms
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(1264,'门诊输液注射管理','辅助门诊护士对接受治疗的门诊病人排队管理及治疗过程登记',100,'zl9CISJob');

--1264:输液排队(基本)
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1264,'基本',Null);
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1264,'所有科室','所有科室病人执行输液排队功能的权限');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1264,'座位安排','允许给就诊病人安排座位');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1264,'座位管理','增加、修改、删除权限');

Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1264,'排队管理','允许对本科室的病人队列进行调整');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1264,'医嘱接单','允许对医嘱执行项目进行操作');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1264,'医嘱执行','允许对医嘱执行项目进行操作');
Insert Into zlProgFuncs(系统,序号,功能,说明) Values(100,1264,'药品寄存','可否进行药品寄存操作');

--  1264:输液排队(基本)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',User,'执行打印记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',User,'暂存药品记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',User,'诊疗项目目录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',User,'药品规格','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',User,'病人医嘱执行','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'部门表','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'病人挂号记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'病人信息','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'病人医嘱记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'病人医嘱发送','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'病人诊断记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'排队记录','SELECT');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'座位状况记录','SELECT');

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'ZL_排队记录_Addqueue','EXECUTE'); 
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'Zl_排队记录_Update','EXECUTE');

-- 1264:输液排队(座位安排)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'座位安排',USER,'ZL_座位状况记录_Setseating','EXECUTE');
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'座位安排',USER,'ZL_座位状况记录_Clear','EXECUTE'); 
-- 1264:输液排队(座位管理)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'座位管理',USER,'Zl_座位状况记录_Update','EXECUTE'); 
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'座位管理',USER,'Zl_座位状况记录_Insert','EXECUTE'); 
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'座位管理',USER,'Zl_座位状况记录_Delete','EXECUTE'); 
-- 1264:输液排队(医嘱执行)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'医嘱执行',USER,'Zl_病人医嘱执行_Transfusion','EXECUTE'); 
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'医嘱执行',USER,'Zl_病人医嘱执行_Modify','EXECUTE'); 
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'医嘱执行',USER,'Zl_病人医嘱执行_Insert','EXECUTE'); 
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'医嘱执行',USER,'Zl_病人医嘱执行_Delete','EXECUTE'); 
-- 1264:输液排队(药品寄存)
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'药品寄存',USER,'Zl_暂存药品记录_Insert','EXECUTE'); 
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'药品寄存',USER,'Zl_暂存药品记录_Delete','EXECUTE'); 
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'药品寄存',USER,'Zl_暂存药品记录_Adviceused','EXECUTE'); 

--- 新增的过程，权限待定
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(100,1264,'基本',USER,'病人医嘱执行_流水号','SELECT');

--zlMenus
Insert Into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Values('缺省',zlMenus_id.nextval, zlMenus_id.nextval-5,'门诊输液注射管理','门诊输液','F',200,'辅助门诊护士对接受治疗的门诊病人排队管理及治疗过程登记',100,1264);

--zlBaseCode

commit;
