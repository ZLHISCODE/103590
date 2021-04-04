----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--126035:蒋敏,2018-05-23,报表Zlmenus遗漏数据添加
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '预交款操作日报', '预交款操作日报', Null, 105, '汇总统计操作员的工作量及票据使用和收款情况。', 100, 1104 From zlMenus Where 标题 = '住院入出转管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '合约单位费用', '合约单位费用', Null, 105, '按合约单位统计其病人的汇总费用情况或欠费情况。', 100, 1105 From zlMenus Where 标题 = '住院入出转管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '挂号员操作日报', '挂号员操作日报', Null, 105, '汇总统计挂号员的工作量及票据使用和分类挂号情况。', 100, 1112 From zlMenus Where 标题 = '门急诊挂号系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '检验综合统计分析', '检验综合统计分析', Null, 105, '按各种条件进行分组统计分析检验的标本人次、项目人次及费用等。', 100, 1230 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '检验信息系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '学术统计', '学术统计', Null, 105, '指定仪器、项目、时间范围等条件，查询检验结果，标准差、差异率等统计数据。', 100, 1231 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '检验信息系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '抗生素药敏查询', '抗生素药敏查询', Null, 105, '抗生素对药菌的耐药、中介、敏感查询。', 100, 1232 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '检验信息系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '细菌药敏查询', '细菌药敏查询', Null, 105, '细菌对抗生素的耐药、中介、敏感查询。', 100, 1233 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '检验信息系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '细菌药物分布统计', '细菌药物分布统计', Null, 105, '指定条件下，指定某细菌对指定抗生素药敏测试结果。', 100, 1234 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '检验信息系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '检验试剂消耗统计', '检验试剂消耗统计', Null, 105, '统计一段时间内的检验试剂消耗情况。', 100, 1235 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '检验信息系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '病历书写检查', '病历书写检查', Null, 105, '查询各住院科室病历书写情况', 100, 1279 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '病案质控与评分系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '外购药品汇总表', '外购药品汇总(单位)', Null, 105, '按供应商或药品种类汇总外购药品数据，以供查询。', 100, 1312 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '住院药房管理系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '自制药品汇总表', '自制药品汇总表', Null, 105, '按药品种类汇总自制药品数据，以供查询。', 100, 1313 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '住院药房管理系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '部门领用汇总表', '部门分类汇总', Null, 105, '汇总各个用药部门一段时间的领用药品的数据。', 100, 1314 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '住院药房管理系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '药品移库汇总表', '移出汇总统计', Null, 105, '反映一段时间各个药品库房间药品转移的汇总情况。', 100, 1315 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '住院药房管理系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '药品调价汇总表', '药品调价汇总表', Null, 105, '查询药品在一段时间内调价变动汇总情况表。', 100, 1316 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '住院药房管理系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '计划情况表', '计划情况表', Null, 105, '主要查询商品在计划内未付、已付等信息。', 100, 1325 From zlMenus Where 标题 = '药库管理与药品会计系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '药房工作量分析', '药房工作量分析', Null, 105, '反映药房人员的工作量情况。', 100, 1346 From zlMenus Where 标题 = '住院药房管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '医师病案质量统计表', '医师病案质量统计表', Null, 105, '医师病案质量统计表', 100, 1570 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '病案质控与评分系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '科室病案质量统计表', '科室病案质量统计表', Null, 105, '科室病案质量统计表', 100, 1571 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '病案质控与评分系统' And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '病案评分结果清单', '病案评分结果清单', Null, 105, '病案评分结果清单', 100, 1572 From zlMenus Where 系统 = 100 And 组别 = '缺省' And 标题 = '病案质控与评分系统' And 模块 Is Null;
Insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'医保统计报表','住院费用汇总表',Null,105,'医保统计报表',100,1610 From zlMenus Where 系统=100 And 组别='缺省' And 标题='医保支持系统' And 模块 is NULL;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '外购卫材汇总表', '外购卫材来源清单', Null, 105, '按供应商或卫材种类汇总外购卫材数据，以供查询。', 100, 1730 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '部门领用汇总表', '部门分类汇总', Null, 105, '汇总各个用料部门一段时间的领用卫生材料的数据。', 100, 1731 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '卫材移库汇总表', '移出汇总统计', Null, 105, '反映一段时间各个卫生材料库房间卫生材料转移的汇总情况。', 100, 1732 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '卫材直接收支分析（部门）', '卫材直接收支分析(部门)', Null, 105, '卫材直接收支分析(部门)', 100, 1733 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '卫材流向分析', '特定卫材去向(部门)', Null, 105, '按部门或医生分析指定类别或具体卫材的不同去向。', 100, 1734 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '发料部门工作量分析', '发料部门工作量分析', Null, 105, '反映发料人员的工作量情况。', 100, 1735 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '卫材超储短缺分析', '卫材超储短缺分析', Null, 105, '根据卫材的存储上限和下限，查询库存卫材的超储短缺信息。', 100, 1736 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '卫材调价汇总表', '卫材调价汇总表', Null, 105, '查询卫材在一段时间内调价变动汇总情况表。', 100, 1737 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '卫材入出汇总表', '卫材入出汇总表', Null, 105, '按卫材入出经济业务汇总各个卫材库房不同卫材的库存变化数据。', 100, 1738 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert into zlMenus(组别,ID,上级ID,标题,短标题,快键,图标,说明,系统,模块) Select '缺省',zlMenus_ID.Nextval,ID,'卫材收发存汇总表','卫材收发存汇总表',Null,105,'按卫材分类汇总卫材的收发调价以及库存金额和差价。',100,1739 From zlMenus Where 系统=100 And 组别='缺省' And 标题='卫生材料管理系统' And 模块 is NULL;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '卫材效期报警分析', '卫材效期报警分析', Null, 105, '查询在今后一段时间内将失效的卫材的当前库存，以便及时处理。', 100, 1740 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;
Insert Into zlMenus(组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 系统, 模块) Select '缺省', Zlmenus_Id.Nextval, ID, '卫材滞用报警分析', '卫材滞用报警分析', Null, 105, '了解从某段时间以来，一直没有使用的积压卫材。', 100, 1741 From zlMenus Where 标题 = '卫生材料管理系统' And 组别 = '缺省' And 系统 = 100 And 模块 Is Null;


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
--123845:焦博,2018-05-25,增加三方接口Zl_Third_Getdepositbalance来获取病人可用预交余额
Create Or Replace Procedure Zl_Third_Getdepositbalance
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --------------------------------------------------------------------------------------------------
  --功能:获取病人预交款余额
  --入参:Xml_In:
  --    <IN>
  --        <BRID>病人ID</BRID>
  --        <ZYID>主页ID</ZYID> //住院预交查询时有效:传入主页ID，查询第几次的预交余额
  --        <YJLX>预交类型</YJLX> //1-门诊预交;2-住院预交;0-所有预交
  --              说明:如果预交类型没有传入,则缺省为0,读取门诊和住院预交
  --    </IN>
  --出参:Xml_Out
  --  <OUTPUT>
  --     <YJYE>预交余额</YJYE>
  --     ――如无下列错误结点则说明正确执行
  --    <ERROR>
  --      <MSG>错误信息</MSG>
  --    </ERROR>
  --  </OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_病人id   病人信息.病人id%Type;
  n_主页id   病案主页.主页id%Type;
  n_预交类型 病人预交记录.预交类别%Type;
  n_预交余额 病人余额.预交余额%Type;
  n_费用余额 病人余额.费用余额%Type;
  v_Temp     Varchar2(32767); --临时XML
  x_Templet  Xmltype; --模板XML
  v_Err_Msg  Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/ZYID'), Extractvalue(Value(A), 'IN/YJLX')
  Into n_病人id, n_主页id, n_预交类型
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '未传入病人ID,无法完成查询!';
    Raise Err_Item;
  End If;

  If Nvl(n_预交类型, 0) = 1 Then
    Select Nvl(Sum(预交余额), 0) - Nvl(Sum(费用余额), 0)
    Into n_预交余额
    From 病人余额
    Where 病人id = n_病人id And 类型 = 1;
  Elsif Nvl(n_预交类型, 0) = 2 Then
    If Nvl(n_主页id, 0) = 0 Then
      Select Nvl(Sum(预交余额), 0) - Nvl(Sum(费用余额), 0)
      Into n_预交余额
      From 病人余额
      Where 病人id = n_病人id And 类型 = 2;
    Else
      Select Nvl(Sum(金额), 0) - Nvl(Sum(冲预交), 0)
      Into n_预交余额
      From 病人预交记录
      Where 病人id = n_病人id And 主页id = n_主页id And 记录性质 In (1, 11);
      Select Nvl(Sum(金额), 0) Into n_费用余额 From 病人未结费用 Where 病人id = n_病人id And 主页id = n_主页id;
      n_预交余额 := n_预交余额 - n_费用余额;
    End If;
  Else
    Select Nvl(Sum(预交余额), 0) - Nvl(Sum(费用余额), 0) Into n_预交余额 From 病人余额 Where 病人id = n_病人id;
  End If;
  If Nvl(n_预交余额, 0) < 0 Then
    n_预交余额 := 0;
  End If;
  v_Temp := '<YJYE>' || n_预交余额 || '</YJYE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdepositbalance;
/

--125867:黄捷,2018-05-23,RIS接口出院患者有未缴费用不允许执行费用

Create Or Replace Package b_Zlxwinterface Is
  Type t_Refcur Is Ref Cursor;

  --1、接收RIS状态改变
  Procedure Receiverisstate
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
    状态_In     Number,
    操作人员_In 病人医嘱发送.完成人%Type,
    执行时间_In 病人医嘱发送.完成时间%Type := Null,
    执行说明_In 病人医嘱发送.执行说明%Type := Null,
    单独执行_In Number := 0
  );

  --2、费用确认
  Procedure 影像费用执行
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  );

  --3、取消费用确认
  Procedure 影像费用执行_Cancel
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  );

  --4、接收RIS的报告
  Procedure Receivereport
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
    报告所见_In 电子病历内容.内容文本%Type,
    报告意见_In 电子病历内容.内容文本%Type,
    报告建议_In 电子病历内容.内容文本%Type,
    报告医生_In 电子病历记录.创建人%Type
  );

  --5、修改申请单信息
  Procedure 影像病人信息_修改
  (
    医嘱id_In       病人医嘱记录.Id%Type,
    姓名_In         病人信息.姓名%Type,
    性别_In         病人信息.性别%Type,
    年龄_In         病人信息.年龄%Type,
    费别_In         病人信息.费别%Type,
    医疗付款方式_In 病人信息.医疗付款方式%Type,
    民族_In         病人信息.民族%Type,
    婚姻_In         病人信息.婚姻状况%Type,
    职业_In         病人信息.职业%Type,
    身份证号_In     病人信息.身份证号%Type,
    家庭地址_In     病人信息.家庭地址%Type,
    家庭电话_In     病人信息.家庭电话%Type,
    家庭地址邮编_In 病人信息.家庭地址邮编%Type,
    出生日期_In     病人信息.出生日期%Type := Null
  );

  --6、取消申请单信息
  Procedure 取消检查申请单
  (
    医嘱id_In     病人医嘱执行.医嘱id%Type,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := 0,
    拒绝原因_In   病人医嘱发送.执行说明%Type := Null
  );

  --7、插入医嘱操作失败记录
  Procedure Ris医嘱失败记录_Insert
  (
    病人来源_In   In Ris医嘱失败记录.病人来源%Type,
    病人id_In     In Ris医嘱失败记录.病人id%Type,
    主页id_In     In Ris医嘱失败记录.主页id%Type,
    挂号单号_In   In Ris医嘱失败记录.挂号单号%Type,
    发送号_In     In Ris医嘱失败记录.发送号%Type,
    体检任务id_In In Ris医嘱失败记录.体检任务id%Type,
    体检报到号_In In Ris医嘱失败记录.体检报到号%Type,
    发送类型_In   In Ris医嘱失败记录.发送类型%Type
  );

  --8、更新医嘱操作失败记录
  Procedure Ris医嘱失败记录_重发
  (
    Id_In       In Ris医嘱失败记录.Id%Type,
    操作类型_In In Number
  );

  --9、销账后新建住院记账单据
  Procedure 病人医嘱_重建单据
  (
    医嘱id_In In 病人医嘱发送.医嘱id%Type,
    No_In     In 病人医嘱发送.No%Type,
    Action_In In Number
  );

  --10、打印RIS检查预约通知单
  Procedure Ris检查预约_打印(医嘱id_In In Ris检查预约.医嘱id%Type);

  --11、更新RIS分科室启用参数
  Procedure Ris启用控制_Update
  (
    检查类型_In Ris启用控制.检查类型%Type,
    场合_In     Ris启用控制.场合%Type,
    部门ids_In  Varchar2,
    启用类型_In Number
  );

  --12、删除RIS分科室启用参数
  Procedure Ris启用控制_Delete;

  --13、根据元素名提取信息
  Function Ris_Replace_Element_Value
  (
    元素名_In   In 诊治所见项目.中文名%Type,
    病人id_In   In 电子病历记录.病人id%Type,
    就诊id_In   In 电子病历记录.主页id%Type,
    病人来源_In In 电子病历记录.病人来源%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type
  ) Return Varchar2;

  --14、删除RIS分院设置参数
  Procedure Ris分院设置_Delete;

  --15、更新RISRis分院设置参数
  Procedure Ris分院设置_Update
  (
    Id_In           Ris分院设置.Id%Type,
    医院名称_In     Ris分院设置.医院名称%Type,
    医院代码_In     Ris分院设置.医院代码%Type,
    用户名_In       Ris分院设置.用户名%Type,
    密码_In         Ris分院设置.密码%Type,
    数据库服务名_In Ris分院设置.数据库服务名%Type
  );
End b_Zlxwinterface;
/

Create Or Replace Package Body b_Zlxwinterface Is

  --1、接收RIS状态改变
  Procedure Receiverisstate
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
    状态_In     Number,
    操作人员_In 病人医嘱发送.完成人%Type,
    执行时间_In 病人医嘱发送.完成时间%Type := Null,
    执行说明_In 病人医嘱发送.执行说明%Type := Null,
    单独执行_In Number := 0
  ) Is
  
    --参数：医嘱ID_IN - 单独执行的医嘱ID。
    --      状态_IN - -1-删除；0-预约；1-登记；3-检查完成；4-检查中止；9-初步报告；12-报告审核；15-发放
    --     单独执行_In -0-全部执行；1-单独执行；检查医嘱组合是否采用对每个项目分散单独执行的方式
  
    Cursor c_Adviceinfo Is
      Select a.Id, a.相关id, Nvl(a.相关id, a.Id) As 组id, a.诊疗类别, a.病人来源, a.执行科室id, b.执行过程
      From 病人医嘱记录 A, 病人医嘱发送 B
      Where a.Id = b.医嘱id And ID = 医嘱id_In;
    r_Adviceinfo c_Adviceinfo%RowType;
  
    v_执行状态 病人医嘱发送.执行状态%Type;
    v_执行过程 病人医嘱发送.执行过程%Type;
    n_执行     Number; --标记是否需要更新状态，1：需要更新，其他不需要更新
    v_Count    Number;
    v_完成人   病人医嘱发送.完成人%Type;
    v_完成时间 病人医嘱发送.完成时间%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    v_执行状态 := 0;
    v_执行过程 := 0;
  
    --提取医嘱的主医嘱ID，及组ID
    Open c_Adviceinfo;
    Fetch c_Adviceinfo
      Into r_Adviceinfo;
    Close c_Adviceinfo;
  
    --根据状态_IN执行医嘱
    ---1-删除；0-预约；1-登记；3-检查完成；4-检查中止；9-初步报告；12-报告审核；13-取消审核；14-报告删除；15-发放
  
    If 状态_In = -1 Or 状态_In = 0 Then
      v_执行状态 := 0; --未执行
      v_执行过程 := 0;
    Elsif 状态_In = 1 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 2; --已报到
    Elsif 状态_In = 3 Or 状态_In = 14 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 3; --已检查
    Elsif 状态_In = 4 Then
      --不改变
      v_执行状态 := v_执行状态;
    Elsif 状态_In = 9 Or 状态_In = 13 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 4; --已报告
    Elsif 状态_In = 12 Then
      v_执行状态 := 3; --正在执行
      v_执行过程 := 5; --已审核
    Elsif 状态_In = 15 Then
      v_执行状态 := 1; --完全执行
      v_执行过程 := 6; --已完成
      v_完成人   := 操作人员_In;
      v_完成时间 := 执行时间_In;
    End If;
  
    n_执行 := 1; --默认都要更新状态
  
    If 状态_In = 13 Or 状态_In = 14 Then
      --删除对应报告数据
      Delete From 电子病历记录
      Where ID = (Select 病历id From 病人医嘱报告 Where 医嘱id = 医嘱id_In And Risid = Risid_In);
      Delete From 病人医嘱报告 Where 医嘱id = 医嘱id_In And Risid = Risid_In;
    
      --删除后判断是否还存在报告，若存在则医嘱状态保持不变，若报告全部删除则更新医嘱状态
      Select Count(1) Into v_Count From 病人医嘱报告 Where 医嘱id = 医嘱id_In;
    
      If v_Count > 0 Then
        n_执行 := 0; --若存在则医嘱状态保持不变
      End If;
    End If;
  
    --如果是登记，先判断此检查是否未执行
    If 状态_In = 1 Then
      If r_Adviceinfo.执行过程 >= 3 Then
        v_Error := '患者已经做过检查了，不能重复登记。';
        Raise Err_Custom;
      End If;
    End If;
  
    --开始执行医嘱
    If n_执行 = 1 Then
      If Nvl(单独执行_In, 0) = 1 Then
        -- 单个部位医嘱单独执行
        Update 病人医嘱发送
        Set 执行状态 = v_执行状态, 执行过程 = v_执行过程, 执行说明 = 执行说明_In, 完成人 = v_完成人, 完成时间 = v_完成时间
        Where 医嘱id = 医嘱id_In;
      Else
        Update 病人医嘱发送
        Set 执行状态 = v_执行状态, 执行过程 = v_执行过程, 执行说明 = 执行说明_In, 完成人 = v_完成人, 完成时间 = v_完成时间
        Where 医嘱id In (Select ID From 病人医嘱记录 Where (ID = r_Adviceinfo.组id Or 相关id = r_Adviceinfo.组id));
      End If;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receiverisstate;

  --2、费用确认
  Procedure 影像费用执行
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  ) Is
    --参数：医嘱ID_IN=单独执行的医嘱ID。
    --      单独执行_In=检查医嘱组合是否采用对每个项目分散单独执行的方式,0-不单独执行
    Cursor c_Advice Is
      Select ID, 相关id, Nvl(相关id, ID) As 组id, 诊疗类别, 病人来源 From 病人医嘱记录 Where ID = 医嘱id_In;
    r_Advice c_Advice%RowType;
  
    v_Temp     Varchar2(255);
    v_人员编号 人员表.编号%Type;
    v_人员姓名 人员表.姓名%Type;
    v_部门id   部门表.Id%Type;
    v_费用性质 病人医嘱发送.记录性质%Type;
    v_发送号   病人医嘱发送.发送号%Type;
    v_执行过程 病人医嘱发送.执行过程%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  Begin
  
    --取主医嘱ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    Select 发送号, 执行过程 Into v_发送号, v_执行过程 From 病人医嘱发送 Where 医嘱id = r_Advice.组id;
  
    --登记和完成才执行费用  2-登记，3-检查，4-报告，5-审核，6-完成
    If v_执行过程 >= 2 Or v_执行过程 <= 6 Then
      --取当前操作人员
      If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null And 执行部门id_In Is Not Null Then
        v_人员编号 := 操作员编号_In;
        v_人员姓名 := 操作员姓名_In;
        v_部门id   := 执行部门id_In;
      Else
        v_Temp     := Zl_Identity;
        v_部门id   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      End If;
    
      If r_Advice.病人来源 = 2 Then
        Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
        Into v_费用性质
        From 病人医嘱发送
        Where 发送号 = v_发送号 And 医嘱id = 医嘱id_In;
      Else
        v_费用性质 := 1;
      End If;
    
      --执行费用和自动发料
      If v_费用性质 = 1 Then
        Zl_门诊医嘱执行_Finish(医嘱id_In, v_发送号, 单独执行_In, v_人员编号, v_人员姓名, r_Advice.组id, r_Advice.诊疗类别, v_部门id);
      Else
        Zl_住院医嘱执行_Finish(医嘱id_In, v_发送号, 单独执行_In, v_人员编号, v_人员姓名, r_Advice.组id, r_Advice.诊疗类别, v_部门id);
      End If;
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 影像费用执行;

  --3、取消费用确认
  Procedure 影像费用执行_Cancel
  (
    医嘱id_In     影像检查记录.医嘱id%Type,
    单独执行_In   Number := 0,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := Null
  ) Is
    --参数：
    --      医嘱ID_IN=单独执行的医嘱ID。
    --      单独执行_In=检查医嘱组合是否采用对每个项目分散单独执行的方式,0-不单独执行
  
    Cursor c_Advice Is
      Select ID, 相关id, Nvl(相关id, ID) As 组id From 病人医嘱记录 Where ID = 医嘱id_In;
    r_Advice c_Advice%RowType;
  
    v_发送号 病人医嘱发送.发送号%Type;
    v_Count  Number;
    v_Error  Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    --取主医嘱ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    --先检查是否已经出院的住院病人，已经预出院或者出院的检查申请，不允执行费用
    Select Count(*)
    Into v_Count
    From 病人医嘱记录 A, 病案主页 B
    Where a.病人id = b.病人id And a.主页id = b.主页id And (b.出院日期 Is Not Null Or b.状态 = 3) And a.Id = r_Advice.组id;
  
    If v_Count > 0 Then
      v_Error := '住院病人已经出院或者预出院，不能取消费用。';
      Raise Err_Custom;
    End If;
  
    Select 发送号 Into v_发送号 From 病人医嘱发送 Where 医嘱id = r_Advice.组id;
  
    --调用统一的医嘱执行Cancel过程
    Zl_病人医嘱执行_Cancel(医嘱id_In, v_发送号, Null, 单独执行_In, 执行部门id_In, 操作员编号_In, 操作员姓名_In);
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 影像费用执行_Cancel;

  --4、接收RIS的报告
  Procedure Receivereport
  (
    医嘱id_In   病人医嘱发送.医嘱id%Type,
    Risid_In    病人医嘱报告.Risid%Type,
    报告所见_In 电子病历内容.内容文本%Type,
    报告意见_In 电子病历内容.内容文本%Type,
    报告建议_In 电子病历内容.内容文本%Type,
    报告医生_In 电子病历记录.创建人%Type
  ) Is
    --提取病人医嘱及报告的相关信息
    Cursor c_Advice
    (
      v_组id  Number,
      v_Risid Number
    ) Is
      Select e.Id, e.病人来源, e.病人id, e.主页id, e.婴儿, e.病人科室id, e.文件id, e.病历种类, e.病历名称, f.病历id, e.执行科室id
      From (Select c.Id, c.病人来源, c.病人id, c.主页id, c.婴儿, c.病人科室id, c.文件id, d.种类 病历种类, d.名称 病历名称, c.执行科室id
             From (Select a.Id, a.病人来源, a.病人id, a.主页id, a.婴儿, a.病人科室id, b.病历文件id 文件id, a.执行科室id
                    From 病人医嘱记录 A, 病历单据应用 B
                    Where a.Id = v_组id And a.诊疗项目id = b.诊疗项目id(+) And b.应用场合(+) = Decode(a.病人来源, 2, 2, 4, 4, 1)) C,
                  病历文件列表 D
             Where c.文件id = d.Id(+)) E, 病人医嘱报告 F
      Where e.Id = f.医嘱id(+) And f.Risid(+) = v_Risid;
  
    --查找文件的组成元素
    Cursor c_File(v_File Number) Is
      Select a.Id, a.文件id, a.父id, a.对象序号, a.对象类型, a.对象标记, a.保留对象, a.对象属性, a.内容行次, a.内容文本, a.是否换行, a.预制提纲id, a.复用提纲,
             a.使用时机, a.诊治要素id, a.替换域, a.要素名称, a.要素类型, a.要素长度, a.要素小数, a.要素单位, a.要素表示, a.输入形态, a.要素值域
      From 病历文件结构 A
      Where a.文件id = v_File
      Order By a.对象序号;
  
    Cursor c_Report(v_电子病历记录id Number) Is
      Select b.Id, a.内容文本
      From 电子病历内容 A, 电子病历内容 B
      Where a.对象类型 = 3 And a.Id = b.父id And b.对象类型 = 2 And b.终止版 = 0 And a.文件id = v_电子病历记录id;
  
    Cursor c_Content
    (
      v_文件id Number,
      v_表格id Number
    ) Is
      Select a.Id, a.文件id, a.父id, a.对象序号, a.对象类型, a.对象标记, a.保留对象, a.对象属性, a.内容行次, a.内容文本, a.是否换行, a.预制提纲id, a.复用提纲,
             a.使用时机, a.诊治要素id, a.替换域, a.要素名称, a.要素类型, a.要素长度, a.要素小数, a.要素单位, a.要素表示, a.输入形态, a.要素值域
      From 病历文件结构 A
      Where 文件id = v_文件id And 父id = v_表格id;
  
    r_Advice        c_Advice%RowType;
    v_病历id        电子病历内容.文件id%Type;
    v_病历内容id    电子病历内容.Id%Type;
    v_病历内容idnew 电子病历内容.Id%Type;
    v_对象序号      电子病历内容.对象序号%Type;
    v_父id          电子病历内容.父id%Type;
    v_内容文本      电子病历内容.内容文本%Type;
    v_定义提纲id    电子病历内容.定义提纲id%Type;
    --v_格式内容    电子病历格式.内容%Type;
    v_Error Varchar2(255);
    Err_Custom Exception;
    v_主医嘱id 病人医嘱发送.医嘱id%Type;
    v_表格     Varchar2(300);
    n_数量     Number;
    n_Rptcount Number;
    v_病历名称 电子病历记录.病历名称%Type;
    v_挂号单id 病人挂号记录.Id%Type;
  
    Function Getrptno
    (
      v_医嘱idin   病人医嘱发送.医嘱id%Type,
      v_病历名称in 电子病历记录.病历名称%Type
    ) Return Varchar As
      v_Return Number;
      v_No     Number;
      v_Count  Number;
    Begin
      Select Count(医嘱id) + 1 Into v_No From 病人医嘱报告 Where 医嘱id = v_医嘱idin;
      v_Count := 1;
      While v_Count = 1 Loop
        Select Count(ID)
        Into v_Count
        From 病人医嘱报告 A, 电子病历记录 B
        Where a.医嘱id = v_医嘱idin And a.病历id = b.Id And b.病历名称 = v_病历名称in || v_No;
        If v_Count = 1 Then
          v_No := v_No + 1;
        End If;
      End Loop;
      v_Return := v_No;
      Return v_Return;
    End Getrptno;
  
  Begin
  
    -- 提取主医嘱ID ，防止因为传入部位医嘱，导致报告保存出错
    Select Nvl(相关id, ID) As 组id Into v_主医嘱id From 病人医嘱记录 Where ID = 医嘱id_In;
  
    Open c_Advice(v_主医嘱id, Nvl(Risid_In, 0));
    Fetch c_Advice
      Into r_Advice;
  
    If Nvl(r_Advice.文件id, 0) = 0 Then
      v_Error := '本次检查项目没有对应相关的检查报告，请与管理员联系！';
      Raise Err_Custom;
    Else
      If Nvl(r_Advice.病历id, 0) > 0 Then
        ----产生过报告
        --找出检查已填写的报告提纲中含有"%所见%","%描述%","%建议%","%意见%",并用传入的参数更新
        For r_Report In c_Report(r_Advice.病历id) Loop
          If r_Report.内容文本 Like '%所见%' Then
            Update 电子病历内容 Set 内容文本 = 报告所见_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.内容文本 Like '%意见%' Then
            Update 电子病历内容 Set 内容文本 = 报告意见_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.内容文本 Like '%建议%' Then
            Update 电子病历内容 Set 内容文本 = 报告建议_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          End If;
        End Loop;
        --更新保存时间
        Update 电子病历记录
        Set 完成时间 = Sysdate, 保存人 = 报告医生_In, 保存时间 = Sysdate
        Where ID = r_Advice.病历id;
      Else
        --先判断单据中是否有对应的提纲和表格
        If Nvl(报告所见_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_数量
          From 病历文件结构 A, 病历文件结构 B
          Where a.父id = b.Id And a.对象类型 = 3 And b.对象类型 = 1 And a.内容文本 Like '%所见%' And a.文件id = r_Advice.文件id;
        
          If n_数量 <= 0 Then
            v_Error := '在诊疗单据中没有找到【所见】对应的提纲或表格，请联系管理员设置！';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(报告意见_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_数量
          From 病历文件结构 A, 病历文件结构 B
          Where a.父id = b.Id And a.对象类型 = 3 And b.对象类型 = 1 And a.内容文本 Like '%意见%' And a.文件id = r_Advice.文件id;
        
          If n_数量 <= 0 Then
            v_Error := '在诊疗单据中没有找到【意见】对应的提纲或表格，请联系管理员设置！';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(报告建议_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_数量
          From 病历文件结构 A, 病历文件结构 B
          Where a.父id = b.Id And a.对象类型 = 3 And b.对象类型 = 1 And a.内容文本 Like '%建议%' And a.文件id = r_Advice.文件id;
        
          If n_数量 <= 0 Then
            v_Error := '在诊疗单据中没有找到【建议】对应的提纲或表格，请联系管理员设置！';
            Raise Err_Custom;
          End If;
        End If;
      
        If r_Advice.病人来源 = 1 Then
          --门诊，提取挂号单ID
          Select Nvl(c.Id, 0)
          Into v_挂号单id
          From 病人医嘱记录 B, 病人挂号记录 C
          Where b.挂号单 = c.No(+) And c.记录状态 In (1, 3) And b.Id = v_主医嘱id;
        Else
          --体检或者外诊，无挂号单ID，直接设置为0
          v_挂号单id := 0;
        End If;
      
        --产生电子病历记录
        Select 电子病历记录_Id.Nextval Into v_病历id From Dual;
        n_Rptcount := Getrptno(医嘱id_In, r_Advice.病历名称);
        If n_Rptcount > 1 Then
          v_病历名称 := r_Advice.病历名称 || n_Rptcount;
        Else
          v_病历名称 := r_Advice.病历名称;
        End If;
        Insert Into 电子病历记录
          (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 完成时间, 保存人, 保存时间, 最后版本, 签名级别)
        Values
          (v_病历id, r_Advice.病人来源, r_Advice.病人id, Decode(r_Advice.病人来源, 2, r_Advice.主页id, v_挂号单id), r_Advice.婴儿,
           r_Advice.病人科室id, r_Advice.病历种类, r_Advice.文件id, v_病历名称, 报告医生_In, Sysdate, Sysdate, 报告医生_In, Sysdate, 1, 2);
      
        --产生医嘱报告记录
        Insert Into 病人医嘱报告 (医嘱id, 病历id, Risid) Values (v_主医嘱id, v_病历id, Risid_In);
      
        v_对象序号 := 0;
      
        --新产生报告内容
        For r_File In c_File(r_Advice.文件id) Loop
          Select 电子病历内容_Id.Nextval Into v_病历内容id From Dual;
          v_内容文本   := r_File.内容文本;
          v_定义提纲id := 0;
        
          If Nvl(r_File.对象类型, 0) = 1 And Nvl(r_File.父id, 0) = 0 Then
            --提纲
            v_定义提纲id := r_File.Id;
            v_父id       := v_病历内容id;
          End If;
        
          If Nvl(r_File.对象类型, 0) = 4 And r_File.要素名称 Is Not Null Then
            --元素
            v_内容文本 := Zl_Replace_Element_Value(r_File.要素名称, r_Advice.病人id, r_Advice.主页id, r_Advice.病人来源, r_Advice.Id);
          End If;
        
          If Nvl(r_File.父id, 0) <> 0 Then
            v_定义提纲id := 0;
          End If;
        
          v_对象序号 := v_对象序号 + 1;
        
          If Instr(v_表格, '|' || r_File.父id || '|') > 0 Then
            Null;
          Else
            Insert Into 电子病历内容
              (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域,
               要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
            Values
              (v_病历内容id, v_病历id, 1, 0, Decode(v_定义提纲id, 0, v_父id, Null), v_对象序号, r_File.对象类型, r_File.对象标记, r_File.保留对象,
               r_File.对象属性, Null, v_内容文本, r_File.是否换行, r_File.预制提纲id, r_File.复用提纲, r_File.使用时机, r_File.诊治要素id,
               r_File.替换域, r_File.要素名称, r_File.要素类型, r_File.要素长度, r_File.要素小数, r_File.要素单位, r_File.要素表示, r_File.输入形态,
               r_File.要素值域, Decode(v_定义提纲id, 0, Null, v_定义提纲id));
          End If;
        
          --为表格时，插入文本内容
          If Nvl(r_File.对象类型, 0) = 3 And Nvl(r_File.父id, 0) <> 0 Then
            v_表格 := v_表格 || ',|' || r_File.Id || '|';
          
            If r_File.内容文本 Like '%所见%' Then
              v_内容文本 := 报告所见_In || Chr(13) || Chr(13);
            Elsif r_File.内容文本 Like '%意见%' Then
              v_内容文本 := 报告意见_In || Chr(13) || Chr(13);
            Else
              v_内容文本 := 报告建议_In || Chr(13) || Chr(13);
            End If;
          
            For r_Con In c_Content(r_Advice.文件id, r_File.Id) Loop
              Select 电子病历内容_Id.Nextval Into v_病历内容idnew From Dual;
              v_对象序号 := v_对象序号 + 1;
            
              Insert Into 电子病历内容
                (ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id,
                 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
              Values
                (v_病历内容idnew, v_病历id, 1, 0, v_病历内容id, v_对象序号, 2, r_Con.对象标记, r_Con.保留对象, r_Con.对象属性, Null, v_内容文本,
                 r_Con.是否换行, r_Con.预制提纲id, r_Con.复用提纲, r_Con.使用时机, r_Con.诊治要素id, r_Con.替换域, r_Con.要素名称, r_Con.要素类型,
                 r_Con.要素长度, r_Con.要素小数, r_Con.要素单位, r_Con.要素表示, r_Con.输入形态, r_Con.要素值域,
                 Decode(v_定义提纲id, 0, Null, v_定义提纲id));
            End Loop;
          End If;
        End Loop;
      
        --因电子病历格式中含了内容文字格式，此种方法导入之后内容文字将不可见
        --Select 内容 Into v_格式内容 From 病历文件格式 Where 文件ID=r_Advice.文件ID;
        --Insert Into 电子病历格式 (文件ID,内容) Values (v_病历id,v_格式内容);
      
      End If;
    End If;
    Close c_Advice;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receivereport;

  --5、修改申请单信息
  Procedure 影像病人信息_修改
  (
    医嘱id_In       病人医嘱记录.Id%Type,
    姓名_In         病人信息.姓名%Type,
    性别_In         病人信息.性别%Type,
    年龄_In         病人信息.年龄%Type,
    费别_In         病人信息.费别%Type,
    医疗付款方式_In 病人信息.医疗付款方式%Type,
    民族_In         病人信息.民族%Type,
    婚姻_In         病人信息.婚姻状况%Type,
    职业_In         病人信息.职业%Type,
    身份证号_In     病人信息.身份证号%Type,
    家庭地址_In     病人信息.家庭地址%Type,
    家庭电话_In     病人信息.家庭电话%Type,
    家庭地址邮编_In 病人信息.家庭地址邮编%Type,
    出生日期_In     病人信息.出生日期%Type := Null
  ) As
  
    v_年龄     Varchar2(20);
    v_年龄单位 Varchar2(20);
    v_出生日期 Date;
    v_病人来源 病人医嘱记录.病人来源%Type;
    v_病人id   病人医嘱记录.病人id%Type;
  Begin
    Begin
      Select 病人来源, 病人id Into v_病人来源, v_病人id From 病人医嘱记录 Where ID = 医嘱id_In;
    Exception
      When Others Then
        Return;
    End;
  
    If 出生日期_In Is Null And 年龄_In Is Not Null Then
      --根据年龄求出生日期
      v_年龄单位 := Substr(年龄_In, Length(年龄_In), 1);
      If Instr('岁,月,天', v_年龄单位) <= 0 Then
        v_年龄单位 := Null;
      Else
        v_年龄 := Replace(年龄_In, v_年龄单位, '');
      End If;
      Begin
        v_年龄 := To_Number(v_年龄);
      Exception
        When Others Then
          v_年龄 := Null;
      End;
      If v_年龄 Is Not Null And v_年龄单位 Is Not Null Then
        Select Decode(v_年龄单位, '岁', Add_Months(Sysdate, -12 * v_年龄), '月', Add_Months(Sysdate, -1 * v_年龄), '天',
                       Sysdate - v_年龄)
        Into v_出生日期
        From Dual;
      End If;
    Else
      v_出生日期 := 出生日期_In;
    End If;
  
    If v_病人来源 = 3 Then
      Update 病人信息
      Set 姓名 = 姓名_In, 性别 = Nvl(性别_In, 性别), 年龄 = 年龄_In, 出生日期 = v_出生日期, 费别 = Nvl(费别_In, 费别),
          医疗付款方式 = Nvl(医疗付款方式_In, 医疗付款方式), 民族 = Nvl(民族_In, 民族), 婚姻状况 = Nvl(婚姻_In, 婚姻状况), 职业 = Nvl(职业_In, 职业),
          身份证号 = 身份证号_In, 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In, 家庭地址邮编 = 家庭地址邮编_In
      Where 病人id = v_病人id;
    
      --修改对应的医嘱记录
      Update 病人医嘱记录
      Set 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In
      Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
    Else
      Update 病人信息
      Set 民族 = Nvl(民族_In, 民族), 婚姻状况 = Nvl(婚姻_In, 婚姻状况), 职业 = Nvl(职业_In, 职业), 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In,
          家庭地址邮编 = 家庭地址邮编_In
      Where 病人id = v_病人id;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 影像病人信息_修改;

  --6、取消申请单信息
  Procedure 取消检查申请单
  (
    医嘱id_In     病人医嘱执行.医嘱id%Type,
    操作员编号_In 人员表.编号%Type := Null,
    操作员姓名_In 人员表.姓名%Type := Null,
    执行部门id_In 门诊费用记录.执行部门id%Type := 0,
    拒绝原因_In   病人医嘱发送.执行说明%Type := Null
  ) As
    --参数：医嘱ID_IN=单独执行的医嘱ID
  
    v_发送号 病人医嘱执行.发送号%Type;
  
  Begin
  
    Begin
      Select 发送号 Into v_发送号 From 病人医嘱发送 Where 医嘱id = 医嘱id_In;
    Exception
      When Others Then
        Return;
    End;
  
    Zl_病人医嘱执行_拒绝执行(医嘱id_In, v_发送号, 操作员编号_In, 操作员姓名_In, 执行部门id_In, 拒绝原因_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 取消检查申请单;

  --7、插入医嘱操作失败记录
  Procedure Ris医嘱失败记录_Insert
  (
    病人来源_In   In Ris医嘱失败记录.病人来源%Type,
    病人id_In     In Ris医嘱失败记录.病人id%Type,
    主页id_In     In Ris医嘱失败记录.主页id%Type,
    挂号单号_In   In Ris医嘱失败记录.挂号单号%Type,
    发送号_In     In Ris医嘱失败记录.发送号%Type,
    体检任务id_In In Ris医嘱失败记录.体检任务id%Type,
    体检报到号_In In Ris医嘱失败记录.体检报到号%Type,
    发送类型_In   In Ris医嘱失败记录.发送类型%Type
  ) Is
  Begin
    Insert Into Ris医嘱失败记录
      (ID, 病人来源, 病人id, 主页id, 挂号单号, 发送号, 体检任务id, 体检报到号, 发送类型, 发送时间, 重发次数)
    Values
      (Ris医嘱失败记录_Id.Nextval, 病人来源_In, 病人id_In, 主页id_In, 挂号单号_In, 发送号_In, 体检任务id_In, 体检报到号_In, 发送类型_In, Sysdate, 0);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris医嘱失败记录_Insert;

  --8、更新医嘱操作失败记录
  Procedure Ris医嘱失败记录_重发
  (
    Id_In       In Ris医嘱失败记录.Id%Type,
    操作类型_In In Number
  ) Is
    v_重发次数 Ris医嘱失败记录.重发次数%Type;
  Begin
    --操作类型_In -- 1 重发成功，删除记录；2--重发失败
  
    If 操作类型_In = 1 Then
      Delete From Ris医嘱失败记录 Where ID = Id_In;
    Else
      Select 重发次数 Into v_重发次数 From Ris医嘱失败记录 Where ID = Id_In;
      If v_重发次数 >= 99 Then
        v_重发次数 := 99;
      Else
        v_重发次数 := v_重发次数 + 1;
      End If;
      Update Ris医嘱失败记录 Set 发送时间 = Sysdate, 重发次数 = v_重发次数 Where ID = Id_In;
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris医嘱失败记录_重发;

  --9、销账后新建住院记账单据
  Procedure 病人医嘱_重建单据
  (
    医嘱id_In In 病人医嘱发送.医嘱id%Type,
    No_In     In 病人医嘱发送.No%Type,
    Action_In In Number
  ) Is
    -- Action_In: 1 重建单据；2 取消重建单据
    v_No 病人医嘱发送.No%Type;
  Begin
    If Action_In = 1 Then
      Select Nextno(14) Into v_No From Dual;
    
      Update 病人医嘱发送
      Set NO = v_No, 计费状态 = 0
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In);
      Update 住院费用记录 Set 医嘱序号 = Null Where NO = No_In;
    Elsif Action_In = 2 Then
      Update 住院费用记录 Set 医嘱序号 = 医嘱id_In Where NO = No_In;
      Update 病人医嘱发送
      Set NO = No_In, 计费状态 = 4
      Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End 病人医嘱_重建单据;

  --10、打印RIS检查预约通知单
  Procedure Ris检查预约_打印(医嘱id_In In Ris检查预约.医嘱id%Type) Is
    v_Temp     Varchar2(255);
    v_人员姓名 人员表.姓名%Type;
  Begin
    --取当前操作人员
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  
    Update Ris检查预约 Set 是否打印 = 1, 打印人 = v_人员姓名, 打印时间 = Sysdate Where 医嘱id = 医嘱id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris检查预约_打印;

  --11、更新RIS分科室启用参数
  Procedure Ris启用控制_Update
  (
    检查类型_In Ris启用控制.检查类型%Type,
    场合_In     Ris启用控制.场合%Type,
    部门ids_In  Varchar2,
    启用类型_In Number
  ) Is
  
    l_部门id   t_Numlist := t_Numlist();
    v_启用ris  Ris启用控制.是否启用ris%Type;
    v_启用预约 Ris启用控制.是否启用预约%Type;
  
    Cursor c_Dept(Dept_In Varchar2) Is
      Select Column_Value From Table(f_Num2list(Dept_In));
  Begin
  
    If 启用类型_In = 1 Then
      v_启用ris  := 1;
      v_启用预约 := Null;
      Delete From Ris启用控制 Where 检查类型 = 检查类型_In And 场合 = 场合_In And 是否启用ris = 1;
    Else
      v_启用ris  := Null;
      v_启用预约 := 1;
      Delete From Ris启用控制 Where 检查类型 = 检查类型_In And 场合 = 场合_In And 是否启用预约 = 1;
    End If;
  
    If 部门ids_In Is Null Then
      Insert Into Ris启用控制
        (ID, 检查类型, 场合, 部门id, 是否启用ris, 是否启用预约)
      Values
        (Ris启用控制_Id.Nextval, 检查类型_In, 场合_In, Null, v_启用ris, v_启用预约);
    Else
      Open c_Dept(部门ids_In);
      Fetch c_Dept Bulk Collect
        Into l_部门id;
      Close c_Dept;
    
      Forall I In 1 .. l_部门id.Count
        Insert Into Ris启用控制
          (ID, 检查类型, 场合, 部门id, 是否启用ris, 是否启用预约)
        Values
          (Ris启用控制_Id.Nextval, 检查类型_In, 场合_In, l_部门id(I), v_启用ris, v_启用预约);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris启用控制_Update;

  --12、删除RIS分科室启用参数
  Procedure Ris启用控制_Delete Is
  
  Begin
    Delete From Ris启用控制;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris启用控制_Delete;

  --13、根据元素名提取信息
  Function Ris_Replace_Element_Value
  (
    元素名_In   In 诊治所见项目.中文名%Type,
    病人id_In   In 电子病历记录.病人id%Type,
    就诊id_In   In 电子病历记录.主页id%Type,
    病人来源_In In 电子病历记录.病人来源%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type
  ) Return Varchar2 Is
    v_Return Varchar2(4000) := Null;
    Cursor c_Patient Is
      Select 姓名, 性别, Decode(性别, '男', 'M', '女', 'F', 'O') As 性别编码, 出生日期, 病人id, 联系人地址, 家庭电话, 联系人电话, 婚姻状况, 身份证号, 当前科室id,
             当前病区id, 当前床号 As 床号, 就诊卡号, 入院时间, 出院时间
      From 病人信息
      Where 病人id = 病人id_In;
    r_Patient c_Patient%RowType;
  
    Cursor c_Order Is
      Select 主页id, 婴儿, Decode(病人来源, 1, 'OUTPAT', 2, 'INPAT', 'UNK') As 病人来源, 开嘱医生, 开嘱时间, 校对护士, 医嘱内容, 紧急标志, 执行科室id
      From 病人医嘱记录
      Where ID = 医嘱id_In;
    r_Order c_Order%RowType;
  
    Cursor c_Diagnose Is
      Select 诊断描述 || Decode(Nvl(是否疑诊, 0), 0, '', ' (？)') As 临床诊断
      From 病人诊断医嘱 A, 病人诊断记录 B
      Where a.医嘱id = 医嘱id_In And a.诊断id = b.Id;
    r_Diagnose c_Diagnose%RowType;
  
    --获取指定表的行类型
    Procedure p_Get_Rowtype(Table_In In Varchar2) Is
    Begin
      If Table_In = '病人信息' Then
        Open c_Patient;
        Fetch c_Patient
          Into r_Patient;
      Elsif Table_In = '病人医嘱记录' Then
        Open c_Order;
        Fetch c_Order
          Into r_Order;
      Elsif Table_In = '病人诊断记录' Then
        Open c_Diagnose;
        Fetch c_Diagnose
          Into r_Diagnose;
      End If;
    Exception
      When Others Then
        Null;
    End p_Get_Rowtype;
  
  Begin
    Case
    --直接返回的输入元素
      When 元素名_In = '医嘱ID' Then
        v_Return := 医嘱id_In;
      When 元素名_In = '病人ID' Then
        v_Return := 病人id_In;
      
    --姓名，性别单独处理，可能是婴儿
      When Instr(',姓名,性别,性别编码,出生日期,', ',' || 元素名_In || ',') > 0 Then
        p_Get_Rowtype('病人医嘱记录');
        p_Get_Rowtype('病人信息');
        If Nvl(r_Order.婴儿, 0) = 0 Then
          If 元素名_In = '姓名' Then
            v_Return := r_Patient.姓名;
          Elsif 元素名_In = '性别' Then
            v_Return := r_Patient.性别;
          Elsif 元素名_In = '性别编码' Then
            v_Return := r_Patient.性别编码;
          Elsif 元素名_In = '出生日期' Then
            v_Return := To_Char(r_Patient.出生日期, 'YYYYMMDDMISS');
          End If;
        Else
          If 元素名_In = '姓名' Then
            Select Decode(婴儿姓名, Null, r_Patient.姓名 || '之婴' || Trim(To_Char(序号, '9')), 婴儿姓名) As 婴儿姓名
            Into v_Return
            From 病人新生儿记录
            Where 病人id = 病人id_In And 主页id = r_Order.主页id And 序号 = Nvl(r_Order.婴儿, 0);
          Elsif Instr('性别', 元素名_In) > 0 Then
            Select 婴儿性别
            Into v_Return
            From 病人新生儿记录
            Where 病人id = 病人id_In And 主页id = r_Order.主页id And 序号 = Nvl(r_Order.婴儿, 0);
            If 元素名_In = '性别编码' Then
              Select Decode(v_Return, '男', 'M', '女', 'F', 'O') Into v_Return From Dual;
            End If;
          Elsif 元素名_In = '出生日期' Then
            Select 出生时间
            Into v_Return
            From 病人新生儿记录
            Where 病人id = 病人id_In And 主页id = r_Order.主页id And 序号 = Nvl(r_Order.婴儿, 0);
            v_Return := To_Char(v_Return, 'YYYYMMDDMISS');
          End If;
        End If;
      
    --查询病人信息表返回的元素
      When Instr(',联系人地址,家庭电话,联系人电话,婚姻状况,身份证号,床号,就诊卡号,入院时间,出院时间,', ',' || 元素名_In || ',') > 0 Then
        p_Get_Rowtype('病人信息');
        Case 元素名_In
          When '联系人地址' Then
            v_Return := r_Patient.联系人地址;
          When '家庭电话' Then
            v_Return := r_Patient.家庭电话;
          When '联系人电话' Then
            v_Return := r_Patient.联系人电话;
          When '婚姻状况' Then
            v_Return := r_Patient.婚姻状况;
          When '身份证号' Then
            v_Return := r_Patient.身份证号;
          When '床号' Then
            v_Return := r_Patient.床号;
          When '就诊卡号' Then
            v_Return := r_Patient.就诊卡号;
          When '入院时间' Then
            v_Return := To_Char(r_Patient.入院时间, 'YYYYMMDDMISS');
          When '出院时间' Then
            v_Return := To_Char(r_Patient.出院时间, 'YYYYMMDDMISS');
          Else
            v_Return := '';
        End Case;
        --查询医嘱表返回的元素
      When Instr(',病人来源,开嘱医生,开嘱时间,校对护士,医嘱内容,紧急标志,紧急标志对码,', ',' || 元素名_In || ',') > 0 Then
        p_Get_Rowtype('病人医嘱记录');
        Case 元素名_In
          When '病人来源' Then
            v_Return := r_Order.病人来源;
          When '开嘱医生' Then
            v_Return := r_Order.开嘱医生;
          When '开嘱时间' Then
            v_Return := To_Char(r_Order.开嘱时间, 'YYYYMMDDMISS');
          When '校对护士' Then
            v_Return := r_Order.校对护士;
          When '医嘱内容' Then
            v_Return := r_Order.医嘱内容;
          When '紧急标志' Then
            v_Return := r_Order.紧急标志;
        End Case;
        --查询诊断记录返回的元素
      When 元素名_In = '临床诊断' Then
        p_Get_Rowtype('病人诊断记录');
        v_Return := r_Diagnose.临床诊断;
      
      Else
        --自行查询SQL返回值的元素
        If 元素名_In = '执行站点' Then
          p_Get_Rowtype('病人医嘱记录');
          Select Decode(站点, 1, 'SITE0002', 2, 'SITE0001', 3, 'SITE0003', 'SITE0001')
          Into v_Return
          From 部门表
          Where ID = r_Order.执行科室id;
        End If;
        If 元素名_In = '当前科室名称' Then
          p_Get_Rowtype('病人信息');
          Select 名称 Into v_Return From 部门表 Where ID = r_Patient.当前科室id;
        End If;
        If 元素名_In = '病区名称' Then
          p_Get_Rowtype('病人信息');
          Select 名称 Into v_Return From 部门表 Where ID = r_Patient.当前病区id;
        End If;
        If 元素名_In = '标识号' Then
          Select Decode(a.病人来源, 1, c.门诊号, 2, Decode(c.住院号, Null, c.门诊号, c.住院号), 4, c.健康号, c.门诊号)
          Into v_Return
          From 病人医嘱记录 A, 病人信息 C
          Where a.病人id = c.病人id And a.Id = 医嘱id_In;
        End If;
    End Case;
  
    Return Trim(v_Return);
  Exception
    When Others Then
      Return Null;
  End Ris_Replace_Element_Value;

  --14、删除RIS分院设置参数
  Procedure Ris分院设置_Delete Is
  Begin
    Delete From Ris分院设置;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris分院设置_Delete;

  --15、更新RISRis分院设置参数
  Procedure Ris分院设置_Update
  (
    Id_In           Ris分院设置.Id%Type,
    医院名称_In     Ris分院设置.医院名称%Type,
    医院代码_In     Ris分院设置.医院代码%Type,
    用户名_In       Ris分院设置.用户名%Type,
    密码_In         Ris分院设置.密码%Type,
    数据库服务名_In Ris分院设置.数据库服务名%Type
  ) Is
  
  Begin
  
    Insert Into Ris分院设置
      (ID, 医院名称, 医院代码, 用户名, 密码, 数据库服务名)
    Values
      (Id_In, 医院名称_In, 医院代码_In, 用户名_In, 密码_In, 数据库服务名_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris分院设置_Update;

End b_Zlxwinterface;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0013' Where 编号=&n_System;
Commit;
