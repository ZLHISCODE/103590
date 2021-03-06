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
--130694:焦博,2018-10-08,调整公共全局参数一卡通消费刷卡控制的参数值含义和警告说明
Update zlParameters
Set 参数值含义 = '参数格式：消费刷卡控制|退费刷卡控制' || Chr(10) ||
             '   1.消费刷卡控制：0-不进行刷卡控制；1-门诊消费时需要刷卡验证；2-门诊消费时设置密码的(只要存在一张卡有密码的，就代表设置了密码的)，则必须刷卡验证；<0表示N元内免密支付，表示病人在消费N元内必须刷卡，不必输入密码即可支付；否则必须输入密码。' || Chr(10) ||
             '   2.退费刷卡控制：0-不进行刷卡控制；1-门诊退费时需要刷卡验证；2-门诊退费时设置密码的(只要存在一张卡有密码的，就代表设置了密码的)，则必须刷卡验证。',
    警告说明 = '此参数建议不调整为"不进行刷卡控制"或"N元内免密支付"，这样可能会存在病人资金安全隐患,为了避免隐患，请要求每个病人都设置刷卡消费密码，以保证资金安全'
Where 系统 = &n_System And 模块 Is Null And 参数号 = 28;

--132356:冉俊明,2018-10-08,增加参数设置按病人补打票据时默认的费用天数
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1121, 0, 1, 0, 0, 0, 0, 118, '缺省发票打印天数', Null, '0',
         '在按病人补打票据时，将根据该参数缺省时间范围来查询费用进行票据打印', '按病人补打票据时缺省提取的费用天数，参数值大于0表示缺省提取的费用天数，参数值等于0表示忽略发生时间提取所有费用，默认为0', Null, '适用于在按病人补打票据时，用户需要自定义费用的缺省查询时间范围的情况',
         Null
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
Update zlSystems Set 版本号='10.35.90.0032' Where 编号=&n_System;
Commit;