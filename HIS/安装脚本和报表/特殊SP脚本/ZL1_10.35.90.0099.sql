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
EXECUTE Zlfiles_Autoupdate('zl9WizardPubFee.DLL','7DA4F0347D2F5D15952C2D0C59DD8927','10.35.90.0099',to_date('2020/7/3 16:54:44','YYYY-MM-DD HH24:MI:SS'),SYSDATE,'1','[APPSOFT]\APPLY','zl9WizardCards.dll;zl9WizardDeposit.dll;zl9WizardFeeQuery.dll;zl9WizardPayFee.dll;zl9WizardRegEvent.dll;zl9WizardInvoice.dll;zl9WizardProof.dll;zl9WizardInsure.dll','26','自助身份识别及支付部件','1','0','');
-------------------------------------------------------------------------------


------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0099' Where 编号=&n_System;
Commit;
