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



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
---标准部件信息
EXECUTE Zlfiles_Autoupdate('ZL9LabWork.DLL','302EA4EFFF9E02187186D3660E73A722','10.35.90.0094',to_date('2020/5/22 9:15:37','YYYY-MM-DD HH24:MI:SS'),SYSDATE,'1','[APPSOFT]\APPLY','ZL9LabWork.dll','25','新版LIS业务部件','1','1','');
-------------------------------------------------------------------------------


------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0094' Where 编号=&n_System;
Commit;
