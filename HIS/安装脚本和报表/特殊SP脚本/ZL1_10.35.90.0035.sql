----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------



------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--130409:焦博,2018-11-01,自助挂号和自助预约管理模块增加zl_Fun_病人挂号记录_Check的执行权限
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1802,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 
Union All Select 'zl_Fun_病人挂号记录_Check','EXECUTE' From Dual) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1803,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 
Union All Select 'zl_Fun_病人挂号记录_Check','EXECUTE' From Dual) A;




-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------




------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0035' Where 编号=&n_System;
Commit;
