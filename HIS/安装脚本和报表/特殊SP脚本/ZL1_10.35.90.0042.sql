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
--127035:黄捷,2018-12-20,专业版PACS支持查看患者多个病人ID的图像
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1288, 0, 0, 0, 0, 0, 0, 20, 'XWWeb检查列表观片地址', null, 'http://localhost:8081/ClinicList.aspx?colid0=103&'||'colvalue0=~in~[@PAT_NOs]',
         'XWWEB观片时的URL，传入多个病人ID，显示检查列表后观片', '使用XWWEB观片时的URL，支持输入多个病人ID', NULL, '适用于专业版PACS观片', Null
  From Dual;

--135806:焦博,2018-12-17,费用虚拟模块增加私有模块参数显示所有号别
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, 100, 9000, 1, 0, 0, 0, 0, 2, 19, '显示不当班号别', '0', '0',
         '保存门诊医生站挂号时对于是否显示不当班号别的选择,以便下次自动恢复', '0-不显示不当班号别；1-显示不当班号别', '', '适用于医生站挂号时常挂其他医生号别或不当班号别的情况', Null
  From Dual;



-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------




------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0042' Where 编号=&n_System;
Commit;