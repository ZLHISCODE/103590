----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--123451:黄捷,2018-03-28,RIS接口预约增加打印人和打印时间
alter table RIS检查预约 add 打印时间 date;
alter table RIS检查预约 add 打印人 VARCHAR2(100);

--122954:余伟节,2018-03-26,中联合理用药
Create Global Temporary Table 中联合理用药参数(参数内容 clob) On Commit Delete Rows;



------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--112136:胡俊勇,2018-03-26,参数性质变化
Insert Into zlParameters(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select zlParameters_ID.Nextval,&n_System,1254,A.* From (
Select 私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0 Union All 
Select 0,0,0,0,0,0,84,'住院本科执行自动完成方案',NULL,NULL,'住院 本科执行自动完成医嘱类别 参数现在按不同的科室设置不同本科执行自动完成医嘱类别对照方案；','科室1,科室2;科室3,科室4···每个分号为一个方案','科室本科执行自动完成医嘱类别对照参数的方案对照；',NULL,NULL From Dual Union All    
Select 私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0) A;

--112136:胡俊勇,2018-03-26,参数性质变化
Declare
  v_部门ids Varchar2(4000);
Begin
  For P In (Select ID, 参数值, 部门
            From zlParameters
            Where 参数名 = '本科执行自动完成医嘱类别' And 模块 = 1254 And 系统 = &n_System) Loop
    If Nvl(p.部门, 0) = 0 Then
      Update zlParameters Set 参数值=null,缺省值=null,部门 = 1 Where ID = p.Id;
      If p.参数值 Is Not Null Then
        For R In (Select Distinct a.Id
                  From 部门表 A, 部门性质说明 B
                  Where b.部门id = a.Id And
                        (b.工作性质 = '临床' And ((b.服务对象 In (2, 3)) Or
                        (b.服务对象 = 1 And Exists (Select 1 From 床位状况记录 C Where b.部门id = c.科室id))) Or
                        b.服务对象 In (1, 2, 3) And b.工作性质 = '护理') And
                        (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)) Loop
          v_部门ids := v_部门ids || ',' || r.Id;
          Insert Into Zldeptparas (参数id, 部门id, 参数值) Values (p.Id, r.Id, p.参数值);
        End Loop;
        v_部门ids := Substr(v_部门ids, 2);
        Update zlParameters
        Set 参数值 = v_部门ids
        Where 参数名 = '住院本科执行自动完成方案' And 模块 = 1254 And 系统 = &n_System;
      End If;
    End If;
  End Loop;
End;
/

--123386:刘硕,2018-03-23,收费价目与收费对照锚点
Update Zlmsg_Lists Set Key_Define='<root><收费项目ID></收费项目ID></root>' Where Code='ZLHIS_DICT_053';
Update Zlmsg_Lists Set Key_Define='<root><诊疗项目ID></诊疗项目ID></root>' Where Code='ZLHIS_DICT_054';

--122954:余伟节,2018-03-26,中联合理用药
Insert into zlTables(系统,表名,表空间,分类) Values(100,'中联合理用药参数','','B2');

Insert Into zlParameters(ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
Select Zlparameters_Id.Nextval, &n_System, Null, a.* From (Select 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明 From zlParameters Where 1 = 0 Union All
Select 1, 0, 0, 0, 0, 0, 299, '药品说明书要点提示', NULL, NULL, '用于控制下达药品医嘱时是否弹出药品提示以及允许展示的药品提示项目','参数值为0|1代码;0-代表关闭,1代表启用;第一位代表是否开启要点提示;从第二位开始对应每一个要点提示项目的开启状态,参数值位数等于要点提示项目个数加1','在启用“合理用药监测接口”为中联信息的前提下有效', '根据个人需要关闭提示,灵活设置提示项目', NULL From Dual Union All
Select 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明 From zlParameters Where 1 = 0) A;

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--122954:余伟节,2018-03-26,中联合理用药
--1252 门诊医嘱下达
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1252,'合理用药监测',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_中联合理用药参数_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_中联合理用药参数','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

--1253:住院医嘱下达
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1253,'合理用药监测',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_中联合理用药参数_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_中联合理用药参数','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;
--1254:住院医嘱发送
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1254,'合理用药监测',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_中联合理用药参数_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_中联合理用药参数','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;
--1341:药品处方发药
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1341,'合理用药监测',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_中联合理用药参数_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_中联合理用药参数','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;
--1342:药品部门发药
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1342,'合理用药监测',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_中联合理用药参数_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_中联合理用药参数','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;
--1345:输液配置中心管理
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1345,'合理用药监测',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_Lob_Append','EXECUTE' From Dual Union All
Select 'Zl_中联合理用药参数_Update','EXECUTE' From Dual Union All
Select 'Zl_Read_中联合理用药参数','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;





-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--123451:黄捷,2018-03-28,RIS接口预约增加打印人和打印时间
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
    v_Count    Number;
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
    
      --先检查是否已经出院的住院病人，已经预出院或者出院的检查申请，不允执行费用
      Select Count(*)
      Into v_Count
      From 病人医嘱记录 A, 病案主页 B
      Where a.病人id = b.病人id And a.主页id = b.主页id And (b.出院日期 Is Not Null Or b.状态 = 3) And a.Id = r_Advice.组id;
    
      If v_Count > 0 Then
        --已经出院、预出院或转院，需要判断先是否死亡
        Select Count(*)
        Into v_Count
        From 病人医嘱记录 A, 诊疗项目目录 B, 病人医嘱发送 C
        Where a.Id = c.医嘱id And a.诊疗项目id = b.Id And b.类别 = 'Z' And b.操作类型 = 11 And
              a.病人id = (Select d.病人id From 病人医嘱记录 D Where d.Id = r_Advice.组id);
        If v_Count > 0 Then
          v_Error := '已经对患者下达死亡医嘱，不能执行费用。';
          Raise Err_Custom;
        End If;
        --再判断是否已经预约，已经预约可执行
        Select Count(*) Into v_Count From Ris检查预约 Where 医嘱id = r_Advice.组id;
        If v_Count = 0 Then
          --已经出院或者预出院，未预约，如果在旧版PACS已经报到，也可以执行
          Select Count(*) Into v_Count From 影像检查记录 Where 医嘱id = r_Advice.组id;
          If v_Count = 0 Then
            v_Error := '住院病人已经出院或者预出院，不能执行费用。';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
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

--123112:董露露,2018-03-27,毒麻处方代办人信息录入用药理由
Create Or Replace Procedure Zl_代办人信息_Insert
(
  病人id_In         In 病人信息从表.病人id%Type,
  身份证号_In       In 病人信息.身份证号%Type,
  代办人姓名_In     In 病人信息从表.信息值%Type,
  代办人身份证号_In In 病人信息从表.信息值%Type,
  就诊id_In         In 病人信息从表.就诊id%Type,
  代办人性别_In     In 病人信息从表.信息值%Type := Null,
  代办人年龄_In     In 病人信息从表.信息值%Type := Null,
  代办人电话_In     In 病人信息从表.信息值%Type := Null,
  用药理由_In       In 病人信息从表.信息值%Type := Null
) As
Begin
  --修改病人身份证号 
  Update 病人信息
  Set 身份证号 = 身份证号_In
  Where 病人id = 病人id_In And (身份证号 Is Null Or 身份证号 <> 身份证号_In);

  Update 病人信息从表
  Set 信息值 = 身份证号_In
  Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '病人身份证号';
  If Sql%RowCount = 0 Then
    Insert Into 病人信息从表
      (病人id, 就诊id, 信息名, 信息值)
    Values
      (病人id_In, 就诊id_In, '病人身份证号', 身份证号_In);
  End If;
  --修改病人信息从表——代办人姓名、代办人身份证号 
  If 代办人姓名_In Is Null Then
    Delete From 病人信息从表 Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '代办人姓名';
    Delete From 病人信息从表
    Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '代办人身份证号';
    Delete From 病人信息从表 Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '代办人性别';
    Delete From 病人信息从表 Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '代办人年龄';
    Delete From 病人信息从表 Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '代办人电话';
    Delete From 病人信息从表 Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '用药理由';
  Else
    Update 病人信息从表
    Set 信息值 = 代办人姓名_In
    Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '代办人姓名';
    If Sql%RowCount = 0 Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
      Values
        (病人id_In, 就诊id_In, '代办人姓名', 代办人姓名_In);
    End If;
  
    Update 病人信息从表
    Set 信息值 = 代办人身份证号_In
    Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '代办人身份证号';
    If Sql%RowCount = 0 Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
      Values
        (病人id_In, 就诊id_In, '代办人身份证号', 代办人身份证号_In);
    End If;
  
    Update 病人信息从表
    Set 信息值 = 代办人性别_In
    Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '代办人性别';
    If Sql%RowCount = 0 Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
      Values
        (病人id_In, 就诊id_In, '代办人性别', 代办人性别_In);
    End If;
  
    Update 病人信息从表
    Set 信息值 = 代办人年龄_In
    Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '代办人年龄';
    If Sql%RowCount = 0 Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
      Values
        (病人id_In, 就诊id_In, '代办人年龄', 代办人年龄_In);
    End If;
  
    Update 病人信息从表
    Set 信息值 = 代办人电话_In
    Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '代办人电话';
    If Sql%RowCount = 0 Then
      Insert Into 病人信息从表
        (病人id, 就诊id, 信息名, 信息值)
      Values
        (病人id_In, 就诊id_In, '代办人电话', 代办人电话_In);
    End If;

    Update 病人信息从表
    Set 信息值 = 用药理由_In
    Where 病人id = 病人id_In And Nvl(就诊id, 0) = Nvl(就诊id_In, 0) And 信息名 = '用药理由';
    If Sql%RowCount = 0 Then
      Insert Into 病人信息从表 (病人id, 就诊id, 信息名, 信息值) Values (病人id_In, 就诊id_In, '用药理由', 用药理由_In);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_代办人信息_Insert;
/
--122954:余伟节,2018-03-26,中联合理用药
Create Or Replace Procedure Zl_Lob_Append
(
  Tab_In     In Number,
  Key_In     In Varchar2,
  Txt_In     In Varchar2, --16进制的文件片段或文字片段
  Cls_In     In Number := 0, --是否清除原来的内容，第一片段传递时为1，以后为0
  Lobtype_In In Number := 0 --0-BLOB;1-CLOB
  --参数说明：
  --Tab_In：包含LOB的数据表
  --        0-病历标记图形;1-病历文件格式;2-病历文件图形;3-病历范文格式;4-病历范文图形;
  --        5/21-电子病历格式;6-电子病历图形;7-病历页面格式；8-电子病历附件;9-体温重叠标记
  --        10-临床路径文件,11-临床路径图标;14-人员证书记录;15-人员表;16-人员照片;
  --        19-部门扩展信息;20-人员扩展信息;22-医嘱报告内容;
  --        23-供应商照片;24-自定义申请单文件;25-医嘱申请单文件
  --        26-门诊路径文件,27-病人照片,28-咨询图片元素,29-咨询段落目录,30-中联合理用药参数
  --Key_In：数据记录的关键字
  --Txt_In：16进制的文件片段或文字片段
  --Cls_In：是否清除原来的内容，第一片段传递时为1，以后为0
  --Lobtype_In:--0-BLOB;1-CLOB
) Is
  l_Blob Blob;
  l_Clob Clob;
  t_Key  t_Strlist;
Begin
  If Tab_In = 0 Then
    If Cls_In = 1 Then
      Update 病历标记图形 Set 图形 = Empty_Blob() Where 编码 = Key_In;
    End If;
    Select 图形 Into l_Blob From 病历标记图形 Where 编码 = Key_In For Update;
  Elsif Tab_In = 1 Then
    If Cls_In = 1 Then
      Update 病历文件格式 Set 内容 = Empty_Blob() Where 文件id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历文件格式 (文件id, 内容) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 内容 Into l_Blob From 病历文件格式 Where 文件id = To_Number(Key_In) For Update;
  Elsif Tab_In = 2 Then
    If Cls_In = 1 Then
      Update 病历文件图形 Set 图形 = Empty_Blob() Where 对象id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历文件图形 (对象id, 图形) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 图形 Into l_Blob From 病历文件图形 Where 对象id = To_Number(Key_In) For Update;
  Elsif Tab_In = 3 Then
    If Cls_In = 1 Then
      Update 病历范文格式 Set 内容 = Empty_Blob() Where 文件id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历范文格式 (文件id, 内容) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 内容 Into l_Blob From 病历范文格式 Where 文件id = To_Number(Key_In) For Update;
  Elsif Tab_In = 4 Then
    If Cls_In = 1 Then
      Update 病历范文图形 Set 图形 = Empty_Blob() Where 对象id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病历范文图形 (对象id, 图形) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 图形 Into l_Blob From 病历范文图形 Where 对象id = To_Number(Key_In) For Update;
  Elsif Tab_In = 5 Then
    If Cls_In = 1 Then
      Update 电子病历格式 Set 内容 = Empty_Blob() Where 文件id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 电子病历格式 (文件id, 内容) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 内容 Into l_Blob From 电子病历格式 Where 文件id = To_Number(Key_In) For Update;
  Elsif Tab_In = 6 Then
    If Cls_In = 1 Then
      Update 电子病历图形 Set 图形 = Empty_Blob() Where 对象id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 电子病历图形 (对象id, 图形) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 图形 Into l_Blob From 电子病历图形 Where 对象id = To_Number(Key_In) For Update;
  Elsif Tab_In = 7 Then
    If Cls_In = 1 Then
      Update 病历页面格式
      Set 图形 = Empty_Blob()
      Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
    End If;
    Select 图形
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 8 Then
    If Cls_In = 1 Then
      Update 电子病历附件
      Set 内容 = Empty_Blob()
      Where 病历id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 序号 = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select 内容
    Into l_Blob
    From 电子病历附件
    Where 病历id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 序号 = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 9 Then
    If Cls_In = 1 Then
      Update 体温重叠标记 Set 标记图形 = Empty_Blob() Where 序号 = To_Number(Key_In);
    End If;
    Select 标记图形 Into l_Blob From 体温重叠标记 Where 序号 = To_Number(Key_In) For Update;
  Elsif Tab_In = 10 Then
    If Cls_In = 1 Then
      Update 临床路径文件
      Set 内容 = Empty_Blob()
      Where 路径id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            文件名 = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select 内容
    Into l_Blob
    From 临床路径文件
    Where 路径id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 文件名 = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 11 Then
    If Cls_In = 1 Then
      Update 临床路径图标 Set 图标 = Empty_Blob() Where ID = To_Number(Key_In);
    End If;
    Select 图标 Into l_Blob From 临床路径图标 Where ID = To_Number(Key_In) For Update;
  Elsif Tab_In = 12 Then
    If Cls_In = 1 Then
      Update 病历页面格式
      Set 页眉文件 = Empty_Blob()
      Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
    End If;
    Select 页眉文件
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 13 Then
    If Cls_In = 1 Then
      Update 病历页面格式
      Set 页脚文件 = Empty_Blob()
      Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3);
    End If;
    Select 页脚文件
    Into l_Blob
    From 病历页面格式
    Where 种类 = To_Number(Substr(Key_In, 1, 1)) And 编号 = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 14 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update 人员证书记录 Set 签章信息 = Empty_Clob() Where 人员id = To_Number(t_Key(1)) And Certsn = t_Key(2);
    End If;
    Select 签章信息 Into l_Clob From 人员证书记录 Where 人员id = To_Number(t_Key(1)) And Certsn = t_Key(2) For Update;
  Elsif Tab_In = 15 Then
    If Cls_In = 1 Then
      Update 人员表 Set 签名图片 = Empty_Blob() Where ID = To_Number(Key_In);
    End If;
    Select 签名图片 Into l_Blob From 人员表 Where ID = To_Number(Key_In) For Update;
    Update 人员表 Set 最后修改时间 = Sysdate Where ID = To_Number(Key_In);
  Elsif Tab_In = 16 Then
    If Cls_In = 1 Then
      Update 人员照片 Set 照片 = Empty_Blob() Where 人员id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 人员照片 (人员id, 照片) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 照片 Into l_Blob From 人员照片 Where 人员id = To_Number(Key_In) For Update;
    Update 人员表 Set 最后修改时间 = Sysdate Where ID = To_Number(Key_In);
  Elsif Tab_In = 19 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update 部门扩展信息 Set 图片 = Empty_Blob() Where 部门id = To_Number(t_Key(1)) And 项目 = t_Key(2);
    End If;
    Select 图片 Into l_Blob From 部门扩展信息 Where 部门id = To_Number(t_Key(1)) And 项目 = t_Key(2) For Update;
    Update 部门表 Set 最后修改时间 = Sysdate Where ID = To_Number(t_Key(1));
  Elsif Tab_In = 20 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update 人员扩展信息 Set 图片 = Empty_Blob() Where 人员id = To_Number(t_Key(1)) And 项目 = t_Key(2);
    End If;
    Select 图片 Into l_Blob From 人员扩展信息 Where 人员id = To_Number(t_Key(1)) And 项目 = t_Key(2) For Update;
    Update 人员表 Set 最后修改时间 = Sysdate Where ID = To_Number(t_Key(1));
  Elsif Tab_In = 21 Then
    If Cls_In = 1 Then
      Update 电子病历格式 Set 文本内容 = Empty_Clob() Where 文件id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 电子病历格式 (文件id, 文本内容) Values (To_Number(Key_In), Empty_Clob());
      End If;
    End If;
    Select 文本内容 Into l_Clob From 电子病历格式 Where 文件id = To_Number(Key_In) For Update;
  Elsif Tab_In = 22 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update 医嘱报告内容 Set 内容 = Empty_Blob() Where ID = To_Number(t_Key(1));
    End If;
    Select 内容 Into l_Blob From 医嘱报告内容 Where ID = To_Number(t_Key(1)) For Update;
  Elsif Tab_In = 23 Then
    If To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=0 Then
      If Cls_In = 1 Then
        Update 供应商照片 Set 许可证号照片 = Empty_Blob() Where 供应商ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into 供应商照片 (供应商ID, 许可证号照片,执照号照片,授权号照片) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select 许可证号照片 Into l_Blob From 供应商照片 Where 供应商ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    Elsif  To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=1 Then
      If Cls_In = 1 Then
        Update 供应商照片 Set 执照号照片 = Empty_Blob() Where 供应商ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into 供应商照片 (供应商ID, 许可证号照片,执照号照片,授权号照片) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select 执照号照片 Into l_Blob From 供应商照片 Where 供应商ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    Elsif To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=2 Then
     If Cls_In = 1 Then
        Update 供应商照片 Set 授权号照片 = Empty_Blob() Where 供应商ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into 供应商照片 (供应商ID, 许可证号照片,执照号照片,授权号照片) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select 授权号照片 Into l_Blob From 供应商照片 Where 供应商ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    End If;
  Elsif Tab_In = 24 Then
    If Cls_In = 1 Then
      Update 自定义申请单文件
      Set 内容 = Empty_Clob()
      Where 文件id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            类别 = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) ;
    End If;
    Select 内容
    Into l_Clob
    From 自定义申请单文件
    Where 文件id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 类别 = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) 
    For Update;
  ElsIf Tab_In = 25 Then
    If Cls_In = 1 Then
      Update 医嘱申请单文件
      Set 内容 = Empty_Clob()
      Where 医嘱id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 
            类别 = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))  ;
    End If;
    Select 内容
    Into l_Clob
    From 医嘱申请单文件
    Where 医嘱id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 类别 = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) 
    For Update;
  Elsif Tab_In = 26 Then
    If Cls_In = 1 Then
      Update 门诊路径文件
      Set 内容 = Empty_Blob()
      Where 路径id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            文件名 = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select 内容
    Into l_Blob
    From 门诊路径文件
    Where 路径id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 文件名 = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 27 Then
    If Cls_In = 1 Then
      Update 病人照片 Set 照片 = Empty_Blob() Where 病人id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into 病人照片 (病人id, 照片) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select 照片 Into l_Blob From 病人照片 Where 病人id = To_Number(Key_In) For Update;
  Elsif Tab_In = 28 Then
    If Cls_In = 1 Then
      Update 咨询图片元素 Set 图形 = Empty_Blob() Where 序号 = To_Number(Key_In);
    End If;
    Select 图形 Into l_Blob From 咨询图片元素 Where 序号 = To_Number(Key_In) For Update;
  Elsif Tab_In = 29 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update 咨询段落目录
      Set 段落文本 = Empty_Clob()
      Where 页面序号 = To_Number(t_Key(1)) And 段落序号 = To_Number(t_Key(2));
    End If;
    Select 段落文本
    Into l_Clob
    From 咨询段落目录
    Where 页面序号 = To_Number(t_Key(1)) And 段落序号 = To_Number(t_Key(2))
    For Update;
  Elsif Tab_In = 30 Then
    If Cls_In = 1 Then
      Insert Into 中联合理用药参数 (参数内容) Values (Empty_Clob());
    End If;
    Select 参数内容 Into l_Clob From 中联合理用药参数 For Update;
  End If;

  If Not Txt_In Is Null Then
    If Lobtype_In = 1 Then
      Dbms_Lob.Writeappend(l_Clob, Length(Txt_In), Txt_In);
    Else
      Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Lob_Append;
/

--123386:刘硕,2018-03-23,收费价目与收费对照锚点
Create Or Replace Procedure Zl_诊疗收费_Update
(
  诊疗项目id_In In 诊疗收费关系.诊疗项目id%Type,
  计价性质_In   诊疗项目目录.计价性质%Type,
  收费内容_In   In Varchar2, --以"|"分隔的诊疗收费内容，每条记录按"诊疗项目ID^数量^固定^从项^性质^部位^检查方法^收费方式"组织
  是否删除_In   Number := 1,
  适用科室id_In 诊疗收费关系.适用科室id%Type := Null,
  病人来源_In   诊疗收费关系.病人来源%Type := 0
) Is
  v_Records    Varchar2(4000);
  v_Currrec    Varchar2(1000);
  v_Fields     Varchar2(1000);
  v_收费项目id 诊疗收费关系.收费项目id%Type;
  v_收费数量   诊疗收费关系.收费数量%Type;
  v_固有对照   诊疗收费关系.固有对照%Type;
  v_从属项目   诊疗收费关系.从属项目%Type;
  v_费用性质   诊疗收费关系.费用性质%Type;
  v_检查部位   诊疗收费关系.检查部位%Type;
  v_检查方法   诊疗收费关系.检查方法%Type;
  v_收费方式   诊疗收费关系.收费方式%Type;
Begin
  Update 诊疗项目目录 Set 计价性质 = 计价性质_In Where ID = 诊疗项目id_In;
  If 是否删除_In = 1 Then
    Delete 诊疗收费关系
    Where 诊疗项目id = 诊疗项目id_In And Nvl(适用科室id, 0) = Nvl(适用科室id_In, 0) And 病人来源 = 病人来源_In;
  End If;
  If 收费内容_In Is Null Then
    v_Records := Null;
  Else
    v_Records := 收费内容_In || '|';
  End If;
  While v_Records Is Not Null Loop
    v_Currrec    := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields     := v_Currrec;
    v_收费项目id := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_收费数量   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_固有对照   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_从属项目   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_费用性质   := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_检查部位   := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_检查方法   := Substr(v_Fields, 1, Instr(v_Fields, '^') - 1);
    v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
    v_收费方式   := To_Number(v_Fields);
    Insert Into 诊疗收费关系
      (诊疗项目id, 收费项目id, 收费数量, 固有对照, 从属项目, 费用性质, 检查部位, 检查方法, 收费方式, 适用科室id, 病人来源)
    Values
      (诊疗项目id_In, v_收费项目id, v_收费数量, v_固有对照, v_从属项目, v_费用性质, v_检查部位, v_检查方法, v_收费方式, 适用科室id_In, 病人来源_In);
    v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
  End Loop;

  If Nvl(适用科室id_In, 0) = 0 And Nvl(病人来源_In, 0) = 0 Then
    b_Message.Zlhis_Dict_054(诊疗项目id_In);
  End If;
  --delete 诊疗收费关系 where 诊疗项目ID=诊疗项目ID_IN and 收费数量=0;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗收费_Update;
/

--123386:刘硕,2018-03-23,收费价目与收费对照锚点
Create Or Replace Procedure Zl_收费价目_Update
(
  收费细目id_In In 收费价目.收费细目id%Type := Null,
  收入项目id_In In 收费价目.收入项目id%Type := Null,
  原价_In       In 收费价目.原价%Type := Null,
  现价_In       In 收费价目.现价%Type := Null,
  附术收费率_In In 收费价目.附术收费率%Type := Null,
  加班加价率_In In 收费价目.加班加价率%Type := Null,
  调价说明_In   In 收费价目.调价说明%Type := Null,
  调价id_In     In 收费价目.调价id%Type := Null,
  调价人_In     In 收费价目.调价人%Type := Null,
  缺省价格_In   In 收费价目.缺省价格%Type := Null,
  价格等级_In   In 收费价目.价格等级%Type := Null
) Is
Begin
  Update 收费价目
  Set 原价 = 原价_In, 现价 = 现价_In, 收入项目id = 收入项目id_In, 加班加价率 = 加班加价率_In, 附术收费率 = 附术收费率_In, 调价说明 = 调价说明_In, 调价id = 调价id_In,
      调价人 = 调价人_In, 缺省价格 = 缺省价格_In
  Where 收费细目id = 收费细目id_In And Nvl(价格等级, '-') = Nvl(价格等级_In, '-') And
        Decode(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'), Null, 终止日期) Is Null;

  If Sql%NotFound Then
    --只有时价才会出现这种情况
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 加班加价率, 附术收费率, 变动原因, 调价说明, 调价id, 调价人, 执行日期, 终止日期, NO, 序号, 缺省价格, 调价汇总号, 价格等级)
    Values
      (收费价目_Id.Nextval, Null, 收费细目id_In, 原价_In, 现价_In, 收入项目id_In, 加班加价率_In, 附术收费率_In, 1, 调价说明_In, 调价id_In, 调价人_In,
       Sysdate, To_Date('3000-01-01', 'yyyy-mm-dd'), Nextno(9), 1, 缺省价格_In, Null, 价格等级_In);
  End If;

  If 价格等级_In Is Null Then
    b_Message.Zlhis_Dict_053(收费细目id_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_收费价目_Update;
/

--123386:刘硕,2018-03-23,收费价目与收费对照锚点
Create Or Replace Procedure Zl_收费价目_Stop
(
  收费细目id_In In 收费价目.收费细目id%Type,
  终止日期_In   In 收费价目.终止日期%Type := Null,
  价格等级_In   In 收费价目.价格等级%Type := Null
) Is
Begin
  Update 收费价目
  Set 终止日期 = 终止日期_In
  Where Decode(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'), Null, 终止日期) Is Null And 收费细目id = 收费细目id_In And
        Nvl(价格等级, '-') = Nvl(价格等级_In, '-');

  If 价格等级_In Is Null Then
    b_Message.Zlhis_Dict_053(收费细目id_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_收费价目_Stop;
/

--123386:刘硕,2018-03-23,收费价目与收费对照锚点
Create Or Replace Procedure Zl_收费价目_Insert
(
  Id_In         In 收费价目.Id%Type,
  原价id_In     In 收费价目.原价id%Type := Null,
  收费细目id_In In 收费价目.收费细目id%Type := Null,
  收入项目id_In In 收费价目.收入项目id%Type := Null,
  原价_In       In 收费价目.原价%Type := Null,
  现价_In       In 收费价目.现价%Type := Null,
  附术收费率_In In 收费价目.附术收费率%Type := Null,
  加班加价率_In In 收费价目.加班加价率%Type := Null,
  调价说明_In   In 收费价目.调价说明%Type := Null,
  调价id_In     In 收费价目.调价id%Type := Null,
  调价人_In     In 收费价目.调价人%Type := Null,
  执行日期_In   In 收费价目.执行日期%Type := Null,
  变动原因_In   In 收费价目.变动原因%Type := 1,
  No_In         In 收费价目.No%Type := Null,
  序号_In       In 收费价目.序号%Type := 1,
  缺省价格_In   In 收费价目.缺省价格%Type := Null,
  调价汇总号_In In 收费价目.调价汇总号%Type := Null,
  价格等级_In   In 收费价目.价格等级%Type := Null
) Is
Begin
  Insert Into 收费价目
    (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 加班加价率, 附术收费率, 变动原因, 调价说明, 调价id, 调价人, 执行日期, 终止日期, NO, 序号, 缺省价格, 调价汇总号, 价格等级)
  Values
    (Id_In, 原价id_In, 收费细目id_In, 原价_In, 现价_In, 收入项目id_In, 加班加价率_In, 附术收费率_In, 变动原因_In, 调价说明_In, 调价id_In, 调价人_In, 执行日期_In,
     To_Date('3000-01-01', 'yyyy-mm-dd'), No_In, 序号_In, 缺省价格_In, 调价汇总号_In, 价格等级_In);
  If 价格等级_In Is Null Then
    b_Message.Zlhis_Dict_053(收费细目id_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_收费价目_Insert;
/

--123386:刘硕,2018-03-23,收费价目与收费对照锚点
CREATE OR REPLACE Package b_Message Is
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  );
  --设置平台调用类型
  Procedure Set_Platform_Call(Platform_Call Number);
  --新增部门
  Procedure Zlhis_Dict_001(Id_In 部门表.Id%Type);
  --修改部门
  Procedure Zlhis_Dict_002(部门id_In 部门表.Id%Type);
  --停用部门
  Procedure Zlhis_Dict_003(部门id_In 部门表.Id%Type);
  --启用部门
  Procedure Zlhis_Dict_004(部门id_In 部门表.Id%Type);
  --新增人员
  Procedure Zlhis_Dict_005(人员id_In 人员表.Id%Type);
  --修改人员
  Procedure Zlhis_Dict_006(人员id_In 人员表.Id%Type);
  --停用人员
  Procedure Zlhis_Dict_007(人员id_In 人员表.Id%Type);
  --启用人员
  Procedure Zlhis_Dict_008(人员id_In 人员表.Id%Type);
  --新增收费项目
  Procedure Zlhis_Dict_009(细目id_In 收费项目目录.Id%Type);
  --修改收费项目
  Procedure Zlhis_Dict_010(细目id_In 收费项目目录.Id%Type);
  --停用收费项目
  Procedure Zlhis_Dict_011(细目id_In 收费项目目录.Id%Type);
  --启用收费项目
  Procedure Zlhis_Dict_012(细目id_In 收费项目目录.Id%Type);
  --新增诊疗项目
  Procedure Zlhis_Dict_013(诊疗id_In 诊疗项目目录.Id%Type);
  --修改诊疗项目
  Procedure Zlhis_Dict_014(诊疗id_In 诊疗项目目录.Id%Type);
  --停用诊疗项目
  Procedure Zlhis_Dict_015(诊疗id_In 诊疗项目目录.Id%Type);
  --启用诊疗项目
  Procedure Zlhis_Dict_016(诊疗id_In 诊疗项目目录.Id%Type);
  --新增检验项目
  Procedure Zlhis_Dict_017(诊疗id_In 诊疗项目目录.Id%Type);
  --修改检验项目
  Procedure Zlhis_Dict_018(诊疗id_In 诊疗项目目录.Id%Type);
  --删除检验项目
  Procedure Zlhis_Dict_019
  (
    诊疗id_In 诊疗项目目录.Id%Type,
    编码_In   诊治所见项目.编码%Type,
    中文名_In 诊治所见项目.中文名%Type,
    英文名_In 诊治所见项目.英文名%Type
  );

  --新增疾病编码目录
  Procedure Zlhis_Dict_021(疾病id_In 疾病编码目录.Id%Type);
  --修改疾病编码目录
  Procedure Zlhis_Dict_022(疾病id_In 疾病编码目录.Id%Type);
  --停用疾病编码目录
  Procedure Zlhis_Dict_023(疾病id_In 疾病编码目录.Id%Type);
  --启用疾病编码目录
  Procedure Zlhis_Dict_024(疾病id_In 疾病编码目录.Id%Type);
  --新增药品分类
  Procedure Zlhis_Dict_025
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --修改药品分类
  Procedure Zlhis_Dict_026
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --删除药品分类
  Procedure Zlhis_Dict_027
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  );
  --停用药品分类
  Procedure Zlhis_Dict_028
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --启用药品分类
  Procedure Zlhis_Dict_029
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --新增药品品种
  Procedure Zlhis_Dict_030
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --修改药品品种
  Procedure Zlhis_Dict_031
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --删除药品品种
  Procedure Zlhis_Dict_032
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type,
    编码_In 诊疗项目目录.编码%Type,
    名称_In 诊疗项目目录.名称%Type
  );
  --停用药品品种
  Procedure Zlhis_Dict_033
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --启用药品品种
  Procedure Zlhis_Dict_034
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  );
  --新增药品规格
  Procedure Zlhis_Dict_035
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --修改药品规格
  Procedure Zlhis_Dict_036
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --删除药品规格
  Procedure Zlhis_Dict_037
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type,
    编码_In   收费项目目录.编码%Type,
    名称_In   收费项目目录.名称%Type,
    规格_In   收费项目目录.规格%Type,
    产地_In   收费项目目录.产地%Type
  );
  --停用药品规格
  Procedure Zlhis_Dict_038
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --启用药品规格
  Procedure Zlhis_Dict_039
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --设置药品存储库房
  Procedure Zlhis_Dict_040
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --设置药品储备限额
  Procedure Zlhis_Dict_041
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  );
  --新增卫材品种
  Procedure Zlhis_Dict_042(Id_In 诊疗项目目录.Id%Type);
  --新增卫材规格
  Procedure Zlhis_Dict_043(Id_In 收费项目目录.Id%Type);
  --修改卫材规格
  Procedure Zlhis_Dict_044(Id_In 收费项目目录.Id%Type);
  --删除卫材规格
  Procedure Zlhis_Dict_045
  (
    Id_In 收费项目目录.Id%Type,
    编码_In   收费项目目录.编码%Type,
    名称_In   收费项目目录.名称%Type,
    规格_In   收费项目目录.规格%Type,
    产地_In   收费项目目录.产地%Type
  );
  --停用卫材规格
  Procedure Zlhis_Dict_046(Id_In 收费项目目录.Id%Type);
  --启用卫材规格
  Procedure Zlhis_Dict_047(Id_In 收费项目目录.Id%Type);
  --医保对码
  Procedure Zlhis_Dict_048
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type
  );
  --删除医保对码
  Procedure Zlhis_Dict_049
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type,
    项目编码_In   In 收费项目目录.编码%Type,
    项目名称_In   In 收费项目目录.名称%Type,
    医保编码_In   In 保险支付项目.项目编码%Type,
    医保名称_In   In 保险支付项目.项目名称%Type
  );
  --新增卫材分类
  Procedure Zlhis_Dict_050
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --修改卫材分类
  Procedure Zlhis_Dict_051
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  );
  --删除卫材分类
  Procedure Zlhis_Dict_052
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  );
  --收费价目变动
  Procedure Zlhis_Dict_053
  (
    收费项目Id_In       收费项目目录.Id%Type
  );
  --诊疗收费对照变动
  Procedure Zlhis_Dict_054
  (
    诊疗项目Id_In     诊疗分类目录.Id%Type
  );
  --新增诊疗检查类型
  Procedure Zlhis_Dictpacs_001
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  );
  --修改诊疗检查类型
  Procedure Zlhis_Dictpacs_002
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  );
  --删除诊疗检查类型
  Procedure Zlhis_Dictpacs_003
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  );
  --新增诊疗检查部位
  Procedure Zlhis_Dictpacs_004
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  );
  --修改诊疗检查部位
  Procedure Zlhis_Dictpacs_005
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  );
  --删除诊疗检查部位
  Procedure Zlhis_Dictpacs_006
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  );
  --新增诊疗项目部位
  Procedure Zlhis_Dictpacs_007
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  );
  --修改诊疗项目部位
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  );
  --删除诊疗项目部位
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  );
    --新增诊疗检验标本
  Procedure Zlhis_DictLis_004
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In 诊疗检验标本.名称%Type,
    简码_In   诊疗检验标本.简码%Type,
    适用性别_In   诊疗检验标本.适用性别%Type
  );
    --修改诊疗检验标本
  Procedure Zlhis_DictLis_005
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In 诊疗检验标本.名称%Type,
    简码_In   诊疗检验标本.简码%Type,
    适用性别_In   诊疗检验标本.适用性别%Type
  );
    --删除诊疗项目部位
  Procedure Zlhis_DictLis_006
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In 诊疗检验标本.名称%Type,
    简码_In   诊疗检验标本.简码%Type,
    适用性别_In   诊疗检验标本.适用性别%Type
  );
  --新增采血管类型
    Procedure Zlhis_DictLis_007
  (
    编码_In     采血管类型.编码%Type,
    名称_In 采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In   采血管类型.添加剂%Type,
    采血量_In   采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料ID_In   采血管类型.材料ID%Type
  );
  --修改采血管类型
    Procedure Zlhis_DictLis_008
  (
    编码_In     采血管类型.编码%Type,
    名称_In 采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In   采血管类型.添加剂%Type,
    采血量_In   采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料ID_In   采血管类型.材料ID%Type
  );
  --删除采血管类型
    Procedure Zlhis_DictLis_009
  (
    编码_In     采血管类型.编码%Type,
    名称_In 采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In   采血管类型.添加剂%Type,
    采血量_In   采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料ID_In   采血管类型.材料ID%Type
  );
  
  --药品备药发送
  Procedure Zlhis_Drug_001(No_In 药品收发记录.No%Type);
  --取消药品备药发送
  Procedure Zlhis_Drug_002(No_In 药品收发记录.No%Type);
  --药品移库单接收
  Procedure Zlhis_Drug_003(No_In 药品收发记录.No%Type);
  --药品移库单冲销
  Procedure Zlhis_Drug_004(No_In 药品收发记录.No%Type);
  --部门发药
  Procedure Zlhis_Drug_005
  (
    库房id_In 药品收发记录.库房id%Type,
    收发id_In 药品收发记录.Id%Type
  );
  --部门退药
  Procedure Zlhis_Drug_006
  (
    冲销收发id_In 药品收发记录.Id%Type,
    待发收发id_In 药品收发记录.Id%Type,
    数量_In       药品收发记录.实际数量%Type,
    费用id_In     门诊费用记录.Id%Type
  );
  --药品调价
  Procedure Zlhis_Drug_007
  (
    价格id_In   药品价格记录.Id%Type
  );
  --静配发送
  Procedure ZLHIS_DRUG_008
  (
    记录Ids_In Varchar2
  );
  --药品调售价
  Procedure Zlhis_Drug_009
  (
    价格id_In   药品价格记录.Id%Type,
    时价_In Number
  );
  --卫材调成本价
  Procedure Zlhis_Drug_010
  (
    价格id_In   成本价调价信息.ID%Type
  );
  --卫材调售价
  Procedure Zlhis_Drug_011
  (
    价格id_In   收费价目.Id%Type,
    时价_In Number
  );
  --2.停止患者医嘱，住院
  Procedure Zlhis_Cis_002
  (
    病人id_In  In 病人医嘱记录.病人id%Type,
    主页id_In  In 病人医嘱记录.主页id%Type,
    医嘱id_In  In 病人医嘱记录.Id%Type,
    医嘱ids_In In Varchar2
  );
  --3.作废患者医嘱，门诊/住院
  Procedure Zlhis_Cis_003
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --4.患者术后医嘱，住院
  Procedure Zlhis_Cis_004
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --5.撤消患者术后医嘱，住院
  Procedure Zlhis_Cis_005
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --6.患者护理常规医嘱，住院
  Procedure Zlhis_Cis_006
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --7.撤消患者护理常规医嘱，住院
  Procedure Zlhis_Cis_007
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --10.下达患者诊断，门诊/住院
  Procedure Zlhis_Cis_010
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    诊断id_In In 病人诊断记录.Id%Type
  );
  --11.撤消患者诊断
  Procedure Zlhis_Cis_011
  (
    病人id_In   In 病人挂号记录.病人id%Type,
    就诊id_In   In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    Id_In       In 病人诊断记录.Id%Type,
    疾病id_In   In 病人诊断记录.疾病id%Type,
    诊断id_In   In 病人诊断记录.诊断id%Type,
    诊断描述_In In 病人诊断记录.诊断描述%Type
  );

  --病区执行医嘱校对
  Procedure Zlhis_Cis_012
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --13.检验危急值阅读通知
  Procedure Zlhis_Cis_014
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  );

  --15.患者检验申请，门诊/住院
  Procedure Zlhis_Cis_016
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );
  --16.患者检查申请，门诊/住院
  Procedure Zlhis_Cis_017
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );
  --17.患者手术申请，门诊/住院
  Procedure Zlhis_Cis_018
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --18.患者输血申请，住院
  Procedure Zlhis_Cis_019
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --19.患者会诊申请，住院
  Procedure Zlhis_Cis_020
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --20.患者抢救医嘱，住院
  Procedure Zlhis_Cis_021
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --21.患者死亡医嘱，住院
  Procedure Zlhis_Cis_022
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --22.患者特殊治疗医嘱，住院
  Procedure Zlhis_Cis_023
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --24.检查危急值阅读通知
  Procedure Zlhis_Cis_025
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  );

  --病区执行医嘱发送
  Procedure Zlhis_Cis_026
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );

  --撤消患者检验申请
  Procedure Zlhis_Cis_036
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );

  --撤消患者检查申请
  Procedure Zlhis_Cis_037
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );

  --撤消患者手术申请
  Procedure Zlhis_Cis_038
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者输血申请
  Procedure Zlhis_Cis_039
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者会诊申请
  Procedure Zlhis_Cis_040
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者抢救医嘱
  Procedure Zlhis_Cis_041
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消患者死亡医嘱
  Procedure Zlhis_Cis_042
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消特殊治疗医嘱
  Procedure Zlhis_Cis_043
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  );

  --撤消病区执行医嘱
  Procedure Zlhis_Cis_044
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    发送数次_In In 病人医嘱发送.发送数次%Type,
    首次时间_In In 病人医嘱发送.首次时间%Type,
    末次时间_In In 病人医嘱发送.末次时间%Type,
    样本条码_In In 病人医嘱发送.样本条码%Type
  );
  --患者医嘱执行登记
  Procedure Zlhis_Cis_050
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type
  );

  --患者医嘱取消执行登记
  Procedure Zlhis_Cis_051
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type,
    本次数次_In In 病人医嘱执行.本次数次%Type,
    执行结果_In In 病人医嘱执行.执行结果%Type,
    执行摘要_In In 病人医嘱执行.执行摘要%Type,
    执行科室_In In 病人医嘱执行.执行科室id%Type,
    执行人_In   In 病人医嘱执行.执行人%Type,
    核对人_In   In 病人医嘱执行.核对人%Type,
    记录来源_In In 病人医嘱执行.记录来源%Type
  );
  --患者医嘱执行完成
  Procedure Zlhis_Cis_052
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --患者医嘱撤消执行完成
  Procedure Zlhis_Cis_053
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  );
  --病理申请发送后修改
  Procedure Zlhis_Cis_056
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type
  );

  --26.检查报告完成，检查完成时
  Procedure Zlhis_Pacs_001
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  );
  --27.检查状态同步，检查状态改变后
  Procedure Zlhis_Pacs_002
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  );
  --28.检查状态回退，检查状态回退后
  Procedure Zlhis_Pacs_003
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  );
  --29.检查报告撤销，撤销检查完成时
  Procedure Zlhis_Pacs_004
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  );
  --30.检查危急值通知，检查发生危急值时
  Procedure Zlhis_Pacs_005(医嘱id_In In 影像检查记录.医嘱id%Type);
  -- 检查预约通知，检查预约时
  Procedure Zlhis_Pacs_006
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    预约id_In In Ris检查预约.预约id%Type
  );
  -- 取消检查预约，取消预约时
  Procedure Zlhis_Pacs_007
  (
    医嘱id_In       In 影像检查记录.医嘱id%Type,
    预约id_In       In Ris检查预约.预约id%Type,
    预约日期_In     In Ris检查预约.预约日期%Type,
    预约序号_In     In Ris检查预约.序号%Type,
    检查设备名称_In In Ris检查预约.检查设备名称%Type
  );


  --36.患者发卡
  Procedure Zlhis_Patient_018
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  );

  --37.患者退卡
  Procedure Zlhis_Patient_019
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  );

  --38.患者退卡
  Procedure Zlhis_Patient_020
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    原卡号_In   In 病人医疗卡信息.卡号%Type,
    新卡号_In   In 病人医疗卡信息.卡号%Type
  );

  --39.病人挂号登记（包含预约登记)
  Procedure Zlhis_Regist_001
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  );

  --40.病人分诊
  Procedure Zlhis_Regist_002
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type,
    诊室_In   In 病人挂号记录.诊室%Type
  );

  --41.病人退号
  Procedure Zlhis_Regist_003
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  );

  --42.临床出诊安排调整
  Procedure Zlhis_Regist_004
  (
    变动原因_In In Integer, --1-停诊;2-替诊;3-诊室变动
    记录id_In   In 临床出诊记录.Id%Type,
    变动id_In   In 临床出诊变动记录.Id%Type
  );

  --43.门诊患者挂号换号操作
  Procedure Zlhis_Regist_005
  (
    No_In         In 病人挂号记录.No%Type,
    变动原因_In   Integer, --1-替诊;2-换号;3-预约日期变动,
    就诊变动id_In 就诊变动记录.Id%Type
  );


  --费用门诊收费及补充结算
  --结算类型_In:1-收费结算，2-补充结算
  Procedure Zlhis_Charge_002
  (
    结算类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  );


  --46.门诊退费单据
  Procedure Zlhis_Charge_004
  (
    退费类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  );

  --47.收预交款
  Procedure Zlhis_Charge_005
  (
    预交id_In In 病人预交记录.Id%Type,
    单据号_In In 病人预交记录.No%Type
  );

  --48.退预交款(包含负数退预交款部分)
  Procedure Zlhis_Charge_006
  (
    退预交id_In In 病人预交记录.Id%Type,
    单据号_In   In 病人预交记录.No%Type
  );

  --住院记帐单据
  Procedure Zlhis_Charge_007
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type
  );

  --住院记帐单据销账
  Procedure Zlhis_Charge_008
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type,
    收发ids_In  In Varchar2 := Null --可能费用ID对应多个收发id，对应格式：收发id,数量|收发id,数量；非药品不传
  );

  --53.住院患者入院登记
  Procedure Zlhis_Patient_001
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --54.住院患者入院入科
  Procedure Zlhis_Patient_002
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --56.住院患者床位变更
  Procedure Zlhis_Patient_004
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --57.住院患者病情变更
  Procedure Zlhis_Patient_005
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --58.住院患者变更撤消
  Procedure Zlhis_Patient_006
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    撤销方式_In In Varchar2
  );
  --59.住院患者医护变更
  Procedure Zlhis_Patient_007
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --住院患者护理等级变更
  Procedure Zlhis_Patient_008
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --60.住院患者预出院
  Procedure Zlhis_Patient_009
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --61.住院患者出院
  Procedure Zlhis_Patient_010
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --62.住院患者新生儿登记
  Procedure Zlhis_Patient_011
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  );
  --63.住院患者转入科室
  Procedure Zlhis_Patient_012
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );
  --64.新生儿登记作废
  Procedure Zlhis_Patient_013
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  );
  --65.门诊患者登记
  Procedure Zlhis_Patient_015(病人id_In In 病案主页.病人id%Type);
  --66.患者信息修改
  Procedure Zlhis_Patient_016(病人id_In In 病案主页.病人id%Type);

  --67.患者合并
  Procedure Zlhis_Patient_017
  (
    病人id_In   In 病案主页.病人id%Type,
    原病人id_In In 病案主页.病人id%Type
  );

  --69.患者转病区转入
  Procedure Zlhis_Patient_026
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  );

  Procedure Zlhis_Patient_028(病人id_In In 病案主页.病人id%Type);

  --血库:科室配血完成
  Procedure Zlhis_Blood_001(医嘱id_In In 病人医嘱记录.Id%Type);
  --血库:科室配血拒绝
  Procedure Zlhis_Blood_002(医嘱id_In In 病人医嘱记录.Id%Type);

  --70.检验标本审核
  Procedure Zlhis_Lis_001(标本id_In In 检验标本记录.Id%Type);
  --71.检验标本审核撤消
  Procedure Zlhis_Lis_002(标本id_In In 检验标本记录.Id%Type);
  --73.检验标本条码打印
  Procedure Zlhis_Lis_004
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  );
  --74.检验标本条码打印撤销
  Procedure Zlhis_Lis_005
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  );
  --75.检验标本核收
  Procedure Zlhis_Lis_006(标本id_In In 检验标本记录.Id%Type);
  --76.检验标本核收撤销
  Procedure Zlhis_Lis_007(标本id_In In 检验标本记录.Id%Type);
  --77.检验标本拒收
  Procedure Zlhis_Lis_008(标本id_In In 检验标本记录.Id%Type);
End b_Message;
/
CREATE OR REPLACE Package Body b_Message Is
  --是否是平台调用
  Is_Platform_Call Number(1) := 0;
  --消息公共方法
  Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --缓存消息查询结果
  Type Tmap_Msg_Using Is Table Of Number(1) Index By Varchar2(30);
  Zlmsg_Map Tmap_Msg_Using;
  --消息是否启用
  Function p_Msg_Using(Msg_Code_In Zlmsg_Lists.Code%Type) Return Number As
    n_Using Zlmsg_Lists.Using%Type;
    v_Code  Zlmsg_Lists.Code%Type;
  Begin
    If Is_Platform_Call = 1 Then
      Return 0;
    End If;
    v_Code := Upper(Msg_Code_In);
    Begin
      n_Using := Zlmsg_Map(v_Code);
      Return n_Using;
    Exception
      When No_Data_Found Then
        --不采取Max容错处理，错误相当于外键,用户可能没有采取同步修改或自己增加了消息类型但是未注册到Zlmsg_Lists，这两种情况会出现错误。


        Select Nvl(Using, 0) Into n_Using From Zlmsg_Lists A Where Code = v_Code;
        Zlmsg_Map(v_Code) := n_Using;
        --查询生成消息的人员，放在这里减少执行次数
        If Message_Creator Is Null Then
          Message_Creator := Zl_Username;
        End If;
        Return n_Using;
    End;
  Exception
    When Others Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || '未在Zlmsg_Lists中找到消息"' || v_Code || '"！请联系管理员进行处理。' || '[ZLSOFT]');
      Return 0;
  End;
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  ) Is
  Begin
    If p_Msg_Using(Msg_Code_In) = 0 Then
      Return;
    End If;
    Insert Into Zlmsg_Todo
      (ID, Msg_Code, Key_Value, State, Create_Time, Creator)
    Values
      (Zlmsg_Todo_Id.Nextval, Upper(Msg_Code_In), Key_Value_In, 0, Sysdate, Message_Creator);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Msg_Todo_Insert;
  --设置当前会话为平台调用
  Procedure Set_Platform_Call(Platform_Call Number) Is
  Begin
    Is_Platform_Call := Platform_Call;
  End Set_Platform_Call;
  --消息Zlhis_Dict_001
  Procedure Zlhis_Dict_001(Id_In 部门表.Id%Type) Is
    v_Define Xmltype;
    v_Value  Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_DICT_001') = 0 Then
      Return;
    End If;
    Begin
      Select Xmltype(Key_Define) Into v_Define From Zlmsg_Lists Where Code = 'ZLHIS_DICT_001';
    Exception
      When Others Then
        v_Define := Xmltype('<root><ID>NULL</ID></root>');
    End;
    Select Updatexml(v_Define, '/root/ID/text()', Id_In).Getstringval() Into v_Value From Dual;
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_001', v_Value);
  End Zlhis_Dict_001;
  --修改部门
  Procedure Zlhis_Dict_002(部门id_In 部门表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_002', v_Value);
  End Zlhis_Dict_002;
  --停用部门
  Procedure Zlhis_Dict_003(部门id_In 部门表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_003', v_Value);
  End Zlhis_Dict_003;
  --启用部门
  Procedure Zlhis_Dict_004(部门id_In 部门表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><部门ID>' || 部门id_In || '</部门ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_004', v_Value);
  End Zlhis_Dict_004;
  --新增人员
  Procedure Zlhis_Dict_005(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_005', v_Value);
  End Zlhis_Dict_005;
  --修改人员
  Procedure Zlhis_Dict_006(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_006', v_Value);
  End Zlhis_Dict_006;
  --停用人员
  Procedure Zlhis_Dict_007(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_007', v_Value);
  End Zlhis_Dict_007;
  --启用人员
  Procedure Zlhis_Dict_008(人员id_In 人员表.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><人员ID>' || 人员id_In || '</人员ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_008', v_Value);
  End Zlhis_Dict_008;
  --新增收费项目
  Procedure Zlhis_Dict_009(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_009', v_Value);
  End Zlhis_Dict_009;
  --修改收费项目
  Procedure Zlhis_Dict_010(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_010', v_Value);
  End Zlhis_Dict_010;
  --停用收费项目
  Procedure Zlhis_Dict_011(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_011', v_Value);
  End Zlhis_Dict_011;
  --启用收费项目
  Procedure Zlhis_Dict_012(细目id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><细目ID>' || 细目id_In || '</细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_012', v_Value);
  End Zlhis_Dict_012;
  --新增诊疗项目
  Procedure Zlhis_Dict_013(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_013', v_Value);
  End Zlhis_Dict_013;
  --修改诊疗项目
  Procedure Zlhis_Dict_014(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_014', v_Value);
  End Zlhis_Dict_014;
  --停用诊疗项目
  Procedure Zlhis_Dict_015(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_015', v_Value);
  End Zlhis_Dict_015;
  --启用诊疗项目
  Procedure Zlhis_Dict_016(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_016', v_Value);
  End Zlhis_Dict_016;
  --新增检验项目
  Procedure Zlhis_Dict_017(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID><系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_017', v_Value);
  End Zlhis_Dict_017;
  --修改检验项目
  Procedure Zlhis_Dict_018(诊疗id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID><系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_018', v_Value);
  End Zlhis_Dict_018;
  --删除检验项目
  Procedure Zlhis_Dict_019
  (
    诊疗id_In 诊疗项目目录.Id%Type,
    编码_In   诊治所见项目.编码%Type,
    中文名_In 诊治所见项目.中文名%Type,
    英文名_In 诊治所见项目.英文名%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗ID>' || 诊疗id_In || '</诊疗ID>' || '<编码>' || 编码_In || '</编码>' || '<中文名>' || 中文名_In || '</中文名>' ||
               '<英文名>' || 英文名_In || '</英文名>' || '<系统>1</系统></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_019', v_Value);
  End Zlhis_Dict_019;
  --新增疾病编码目录
  Procedure Zlhis_Dict_021(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_021', v_Value);
  End Zlhis_Dict_021;
  --修改疾病编码目录
  Procedure Zlhis_Dict_022(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_022', v_Value);
  End Zlhis_Dict_022;
  --停用疾病编码目录
  Procedure Zlhis_Dict_023(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_023', v_Value);
  End Zlhis_Dict_023;
  --启用疾病编码目录
  Procedure Zlhis_Dict_024(疾病id_In 疾病编码目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><疾病ID>' || 疾病id_In || '</疾病ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_024', v_Value);
  End Zlhis_Dict_024;
  --新增药品分类
  Procedure Zlhis_Dict_025
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_025', v_Value);
  End Zlhis_Dict_025;
  --修改药品分类
  Procedure Zlhis_Dict_026
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_026', v_Value);
  End Zlhis_Dict_026;
  --删除药品分类
  Procedure Zlhis_Dict_027
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In  || '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_027', v_Value);
  End Zlhis_Dict_027;
  --停用药品分类
  Procedure Zlhis_Dict_028
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_028', v_Value);
  End Zlhis_Dict_028;
  --启用药品分类
  Procedure Zlhis_Dict_029
  (
    类型_In 诊疗分类目录.类型%Type,
    Id_In   诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_029', v_Value);
  End Zlhis_Dict_029;
  --新增药品品种
  Procedure Zlhis_Dict_030
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_030', v_Value);
  End Zlhis_Dict_030;
  --修改药品品种
  Procedure Zlhis_Dict_031
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_031', v_Value);
  End Zlhis_Dict_031;
  --删除药品品种
  Procedure Zlhis_Dict_032
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type,
    编码_In 诊疗项目目录.编码%Type,
    名称_In 诊疗项目目录.名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In  || '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_032', v_Value);
  End Zlhis_Dict_032;
  --停用药品品种
  Procedure Zlhis_Dict_033
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_033', v_Value);
  End Zlhis_Dict_033;
  --启用药品品种
  Procedure Zlhis_Dict_034
  (
    类别_In 诊疗项目目录.类别%Type,
    Id_In   诊疗项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_034', v_Value);
  End Zlhis_Dict_034;
  --新增药品规格
  Procedure Zlhis_Dict_035
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_035', v_Value);
  End Zlhis_Dict_035;
  --修改药品规格
  Procedure Zlhis_Dict_036
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_036', v_Value);
  End Zlhis_Dict_036;
  --删除药品规格
  Procedure Zlhis_Dict_037
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type,
    编码_In   收费项目目录.编码%Type,
    名称_In   收费项目目录.名称%Type,
    规格_In   收费项目目录.规格%Type,
    产地_In   收费项目目录.产地%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID><编码>' || 编码_In || '</编码><名称>' || 名称_In  || '</名称><规格>' ||
            规格_In || '</规格><产地>' || 产地_In || '</产地></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_037', v_Value);
  End Zlhis_Dict_037;
  --停用药品规格
  Procedure Zlhis_Dict_038
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_038', v_Value);
  End Zlhis_Dict_038;
  --启用药品规格
  Procedure Zlhis_Dict_039
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_039', v_Value);
  End Zlhis_Dict_039;
  --设置药品存储库房
  Procedure Zlhis_Dict_040
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_040', v_Value);
  End Zlhis_Dict_040;
  --设置药品储备限额
  Procedure Zlhis_Dict_041
  (
    类别_In   收费项目目录.类别%Type,
    药品id_In 收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类别>' || 类别_In || '</类别><药品ID>' || 药品id_In || '</药品ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_041', v_Value);
  End Zlhis_Dict_041;
  --新增卫材品种
  Procedure Zlhis_Dict_042(Id_In 诊疗项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_042', v_Value);
  End Zlhis_Dict_042;
  --新增卫材规格
  Procedure Zlhis_Dict_043(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_043', v_Value);
  End Zlhis_Dict_043;
  --修改卫材规格
  Procedure Zlhis_Dict_044(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_044', v_Value);
  End Zlhis_Dict_044;
  --删除卫材规格
  Procedure Zlhis_Dict_045
  (
   Id_In 收费项目目录.Id%Type,
   编码_In   收费项目目录.编码%Type,
   名称_In   收费项目目录.名称%Type,
   规格_In   收费项目目录.规格%Type,
   产地_In   收费项目目录.产地%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID><编码>' || 编码_In || '</编码><名称>' || 名称_In  || '</名称><规格>' || 规格_In || '</规格><产地>' || 产地_In || '</产地></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_045', v_Value);
  End Zlhis_Dict_045;
  --停用卫材规格
  Procedure Zlhis_Dict_046(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_046', v_Value);
  End Zlhis_Dict_046;
  --启用卫材规格
  Procedure Zlhis_Dict_047(Id_In 收费项目目录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><材料ID>' || Id_In || '</材料ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_047', v_Value);
  End Zlhis_Dict_047;
  --医保对码
  Procedure Zlhis_Dict_048
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><险类>' || 险类_In || '</险类><收费细目ID>' || 收费细目id_In || '</收费细目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_048', v_Value);
  End Zlhis_Dict_048;
  --删除医保对码
  Procedure Zlhis_Dict_049
  (
    险类_In       In 保险支付项目.险类%Type,
    收费细目id_In In 保险支付项目.收费细目id%Type,
    项目编码_In   In 收费项目目录.编码%Type,
    项目名称_In   In 收费项目目录.名称%Type,
    医保编码_In   In 保险支付项目.项目编码%Type,
    医保名称_In   In 保险支付项目.项目名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><险类>' || 险类_In || '</险类><收费细目ID>' || 收费细目id_In || '</收费细目ID><项目编码>' || 项目编码_In || '</项目编码><项目名称>' ||
               项目名称_In || '</项目名称><医保编码>' || 医保编码_In || '</医保编码><医保名称>' || 医保名称_In || '</医保名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_049', v_Value);
  End Zlhis_Dict_049;
  --新增卫材分类
  Procedure Zlhis_Dict_050
  (
    类型_In 诊疗分类目录.类型%Type,
    ID_In 诊疗分类目录.ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || ID_In ||  '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_050', v_Value);
  End Zlhis_Dict_050;
  --修改卫材分类
  Procedure ZLHIS_DICT_051
  (
    类型_In 诊疗分类目录.类型%Type,
    ID_In 诊疗分类目录.ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || ID_In ||  '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_051', v_Value);
  End ZLHIS_DICT_051;
  --删除卫材分类
  Procedure ZLHIS_DICT_052
  (
    类型_In 诊疗分类目录.类型%Type,
    ID_In 诊疗分类目录.ID%Type,
    编码_In 诊疗分类目录.编码%Type,
    名称_In 诊疗分类目录.名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><ID>' || Id_In || '</ID><编码>' || 编码_In || '</编码><名称>' || 名称_In  || '</名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_052', v_Value);
  End ZLHIS_DICT_052;
   --收费价目变动
  Procedure Zlhis_Dict_053
  (
    收费项目Id_In       收费项目目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><收费项目ID>' || 收费项目Id_In || '</收费项目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_053', v_Value);
  End Zlhis_Dict_053;

  --诊疗收费对照变动
  Procedure Zlhis_Dict_054
  (
    诊疗项目Id_In     诊疗分类目录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><诊疗项目ID>' || 诊疗项目Id_In || '</诊疗项目ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_054', v_Value);
  End Zlhis_Dict_054;
  --新增诊疗检查类型
  Procedure Zlhis_Dictpacs_001
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '<简码/><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_001', v_Value);
  End Zlhis_Dictpacs_001;

  --修改诊疗检查类型
  Procedure Zlhis_Dictpacs_002
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '<简码/><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_002', v_Value);
  End Zlhis_Dictpacs_002;
  --删除诊疗检查类型
  Procedure Zlhis_Dictpacs_003
  (
    编码_In   诊疗检查类型.编码%Type,
    名称_In   诊疗检查类型.名称%Type,
    简码_In   诊疗检查类型.简码%Type,
    建病案_In 诊疗检查类型.建病案%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '<简码/><建病案>' || 建病案_In ||
               '</建病案></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_003', v_Value);
  End Zlhis_Dictpacs_003;
  --新增诊疗检查部位
  Procedure Zlhis_Dictpacs_004
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_004', v_Value);
  End Zlhis_Dictpacs_004;
  --修改诊疗检查部位
  Procedure Zlhis_Dictpacs_005
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_005', v_Value);
  End Zlhis_Dictpacs_005;
  --删除诊疗检查部位
  Procedure Zlhis_Dictpacs_006
  (
    类型_In     诊疗检查部位.类型%Type,
    编码_In     诊疗检查部位.编码%Type,
    名称_In     诊疗检查部位.名称%Type,
    分组_In     诊疗检查部位.分组%Type,
    备注_In     诊疗检查部位.备注%Type,
    方法_In     诊疗检查部位.方法%Type,
    适用性别_In 诊疗检查部位.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><类型>' || 类型_In || '</类型><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><分组>' || 分组_In ||
               '</分组><备注>' || 备注_In || '</备注><方法>' || 方法_In || '</方法><适用性别>' || 适用性别_In || '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_006', v_Value);
  End Zlhis_Dictpacs_006;
  --新增诊疗项目部位
  Procedure Zlhis_Dictpacs_007
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_007', v_Value);
  End Zlhis_Dictpacs_007;
  --修改诊疗项目部位
  Procedure Zlhis_Dictpacs_008
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_008', v_Value);
  End Zlhis_Dictpacs_008;
  --删除诊疗项目部位
  Procedure Zlhis_Dictpacs_009
  (
    Id_In     诊疗项目部位.Id%Type,
    项目id_In 诊疗项目部位.项目id%Type,
    类型_In   诊疗项目部位.类型%Type,
    部位_In   诊疗项目部位.部位%Type,
    方法_In   诊疗项目部位.方法%Type,
    默认_In   诊疗项目部位.默认%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><项目ID>' || 项目id_In || '</项目ID><类型>' || 类型_In || '</类型><部位>' || 部位_In ||
               '</部位><方法>' || 方法_In || '</方法><默认>' || 默认_In || '</默认></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_009', v_Value);
  End Zlhis_Dictpacs_009;
    --新增诊疗项目部位
  Procedure Zlhis_DictLis_004
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In 诊疗检验标本.名称%Type,
    简码_In   诊疗检验标本.简码%Type,
    适用性别_In   诊疗检验标本.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_004', v_Value);
  End Zlhis_DictLis_004;
      --修改诊疗项目部位
  Procedure Zlhis_DictLis_005
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In 诊疗检验标本.名称%Type,
    简码_In   诊疗检验标本.简码%Type,
    适用性别_In   诊疗检验标本.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_005', v_Value);
  End Zlhis_DictLis_005;
      --删除诊疗项目部位
  Procedure Zlhis_DictLis_006
  (
    编码_In     诊疗检验标本.编码%Type,
    名称_In 诊疗检验标本.名称%Type,
    简码_In   诊疗检验标本.简码%Type,
    适用性别_In   诊疗检验标本.适用性别%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><适用性别>' || 适用性别_In ||
               '</适用性别></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_006', v_Value);
  End Zlhis_DictLis_006;
   --新增采血管类型
  Procedure Zlhis_DictLis_007
  (
    编码_In     采血管类型.编码%Type,
    名称_In 采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In   采血管类型.添加剂%Type,
    采血量_In   采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料ID_In   采血管类型.材料ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料ID_In || '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_007', v_Value);
  End Zlhis_DictLis_007;
    --新增采血管类型
  Procedure Zlhis_DictLis_008
  (
    编码_In     采血管类型.编码%Type,
    名称_In 采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In   采血管类型.添加剂%Type,
    采血量_In   采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料ID_In   采血管类型.材料ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料ID_In || '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_008', v_Value);
  End Zlhis_DictLis_008;
     --新增采血管类型
  Procedure Zlhis_DictLis_009
  (
    编码_In     采血管类型.编码%Type,
    名称_In 采血管类型.名称%Type,
    简码_In   采血管类型.简码%Type,
    添加剂_In   采血管类型.添加剂%Type,
    采血量_In   采血管类型.采血量%Type,
    规格_In   采血管类型.规格%Type,
    颜色_In   采血管类型.颜色%Type,
    材料ID_In   采血管类型.材料ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><编码>' || 编码_In || '</编码><名称>' || 名称_In || '</名称><简码>' || 简码_In || '</简码><添加剂>' || 添加剂_In ||
               '</添加剂><采血量>' || 采血量_In || '</采血量><颜色>' || 颜色_In || '</颜色><规格>' || 规格_In || '</规格><材料ID_In>' || 材料ID_In || '</材料ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_009', v_Value);
  End Zlhis_DictLis_009;  
  --药品备药发送
  Procedure Zlhis_Drug_001(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_001', v_Value);
  End Zlhis_Drug_001;
  --取消药品备药发送
  Procedure Zlhis_Drug_002(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_002', v_Value);
  End Zlhis_Drug_002;
  --药品移库单接收
  Procedure Zlhis_Drug_003(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_003', v_Value);
  End Zlhis_Drug_003;
  --药品移库单冲销
  Procedure Zlhis_Drug_004(No_In 药品收发记录.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><单据号>' || No_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_004', v_Value);
  End Zlhis_Drug_004;
  --部门发药
  Procedure Zlhis_Drug_005
  (
    库房id_In 药品收发记录.库房id%Type,
    收发id_In 药品收发记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><库房ID>' || 库房id_In || '</库房ID><收发ID>' || 收发id_In || '</收发ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_005', v_Value);
  End Zlhis_Drug_005;
  --部门退药
  Procedure Zlhis_Drug_006
  (
    冲销收发id_In 药品收发记录.Id%Type,
    待发收发id_In 药品收发记录.Id%Type,
    数量_In       药品收发记录.实际数量%Type,
    费用id_In     门诊费用记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><冲销记录ID>' || 冲销收发id_In || '</冲销记录ID><待发记录ID>' || 待发收发id_In || '</待发记录ID><数量>' || 数量_In ||
               '</数量><费用ID>' || 费用id_In || '</费用ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_006', v_Value);
  End Zlhis_Drug_006;
  --药品调价
  Procedure ZLHIS_DRUG_007
  (
    价格ID_In 药品价格记录.ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格ID_In ||  '</价格ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_007', v_Value);
  End ZLHIS_DRUG_007;
  --静配发送
  Procedure ZLHIS_DRUG_008
  (
    记录Ids_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    n_记录id 输液配药记录.ID%Type;
    v_Tmp    varchar2(4000);
  Begin
    If 记录Ids_In Is Null Then
      v_Tmp := Null;
    Else
      v_Tmp := 记录Ids_In || ',';
    End If;

    v_Value := '<root><记录IDS>';

    While v_Tmp Is Not Null Loop
      --分解单据ID串
      n_记录id :=to_number(Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1));
      v_Tmp    := Replace(',' || v_Tmp, ',' || n_记录id || ',');

      v_Value:=v_Value || '<记录ID>' || n_记录id || '</记录ID>';
    End Loop;

    v_Value:=v_Value || '</记录IDS></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
  End ZLHIS_DRUG_008;
  --药品调售价
  Procedure ZLHIS_DRUG_009
  (
    价格ID_In 药品价格记录.ID%Type,
    时价_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格ID_In ||  '</价格ID><时价>' || 时价_In || '</时价></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_009', v_Value);
  End ZLHIS_DRUG_009;
  --卫材调成本价
  Procedure ZLHIS_DRUG_010
  (
    价格ID_In 成本价调价信息.ID%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格ID_In ||  '</价格ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_010', v_Value);
  End ZLHIS_DRUG_010;
  --卫材调售价
  Procedure ZLHIS_DRUG_011
  (
    价格ID_In 收费价目.ID%Type,
    时价_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><价格ID>' || 价格ID_In ||  '</价格ID><时价>' || 时价_In || '</时价></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_011', v_Value);
  End ZLHIS_DRUG_011;

  --2.停止患者医嘱，住院
  Procedure Zlhis_Cis_002
  (
    病人id_In  In 病人医嘱记录.病人id%Type,
    主页id_In  In 病人医嘱记录.主页id%Type,
    医嘱id_In  In 病人医嘱记录.Id%Type,
    医嘱ids_In In Varchar2
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_CIS_002') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002',
                                  '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                   '</ID></root>');
    Else
      For R In (Select '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || ID || '</ID></root>' As Xml_Value
                From 病人医嘱记录
                Where ID In (Select Column_Value From Table(f_Num2list(医嘱ids_In))) And 相关id Is Null) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Cis_002;
  --3.作废患者医嘱，门诊/住院
  Procedure Zlhis_Cis_003
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_003', v_Value);
  End Zlhis_Cis_003;

  --4.患者术后医嘱，住院
  Procedure Zlhis_Cis_004
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_004',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_004;

  --5.撤消患者术后医嘱，住院
  Procedure Zlhis_Cis_005
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_005',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_005;

  --6.患者护理常规医嘱，住院
  Procedure Zlhis_Cis_006
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_006',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In ||
                                 '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_006;

  --7.撤消患者护理常规医嘱，住院
  Procedure Zlhis_Cis_007
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_007', v_Value);
  End Zlhis_Cis_007;

  --10.下达患者诊断，门诊/住院
  Procedure Zlhis_Cis_010
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    诊断id_In In 病人诊断记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_010',
                                '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><ID>' || 诊断id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_010;
  --11.撤消患者诊断
  Procedure Zlhis_Cis_011
  (
    病人id_In   In 病人挂号记录.病人id%Type,
    就诊id_In   In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    Id_In       In 病人诊断记录.Id%Type,
    疾病id_In   In 病人诊断记录.疾病id%Type,
    诊断id_In   In 病人诊断记录.诊断id%Type,
    诊断描述_In In 病人诊断记录.诊断描述%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><ID>' || Id_In || '</ID><疾病ID>' ||
               疾病id_In || '</疾病ID><诊断ID>' || 诊断id_In || '</诊断ID><诊断描述>' || 诊断描述_In || '</诊断描述></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_011', v_Value);
  End Zlhis_Cis_011;

  --病区执行医嘱校对
  Procedure Zlhis_Cis_012
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_012',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><ID>' || 医嘱id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_012;

  --13.检验危急值阅读通知
  Procedure Zlhis_Cis_014
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><医嘱ID>' || 医嘱id_In || '</医嘱ID><ID>' ||
               消息id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_014', v_Value);
  End Zlhis_Cis_014;
  --15.患者检验申请，门诊/住院
  Procedure Zlhis_Cis_016
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_016', v_Value);
  End Zlhis_Cis_016;
  --16.患者检查申请，门诊/住院
  Procedure Zlhis_Cis_017
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    v_操作类型 诊疗项目目录.操作类型%Type;
  Begin
    Select MAX(A.操作类型) Into v_操作类型 From 诊疗项目目录 A,病人医嘱记录 B Where B.诊疗项目ID = a.ID And B.ID = 医嘱id_In;
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
    If v_操作类型 = '病理' Then
       b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_054', v_Value);
    Else
       b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_017', v_Value);
    End If;
  End Zlhis_Cis_017;
  --17.患者手术申请，门诊/住院
  Procedure Zlhis_Cis_018
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_018', v_Value);
  End Zlhis_Cis_018;
  --18.患者输血申请，住院
  Procedure Zlhis_Cis_019
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_019', v_Value);
  End Zlhis_Cis_019;
  --19.患者会诊申请，住院
  Procedure Zlhis_Cis_020
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_020', v_Value);
  End Zlhis_Cis_020;
  --20.患者抢救医嘱，住院
  Procedure Zlhis_Cis_021
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_021', v_Value);
  End Zlhis_Cis_021;
  --21.患者死亡医嘱，住院
  Procedure Zlhis_Cis_022
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_022', v_Value);
  End Zlhis_Cis_022;
  --22.患者特殊治疗医嘱，住院
  Procedure Zlhis_Cis_023
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_023', v_Value);
  End Zlhis_Cis_023;

  --24.检查危急值阅读通知
  Procedure Zlhis_Cis_025
  (
    病人id_In In 病人挂号记录.病人id%Type,
    就诊id_In In 病人挂号记录.Id%Type, --门诊病人 挂号ID，住院病人 主页ID
    医嘱id_In In 病人医嘱记录.Id%Type,
    消息id_In In 业务消息清单.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><就诊ID>' || 就诊id_In || '</就诊ID><医嘱ID>' || 医嘱id_In || '</医嘱ID><ID>' ||
               消息id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_025', v_Value);
  End Zlhis_Cis_025;

  --病区执行医嘱发送
  Procedure Zlhis_Cis_026
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_026',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In ||
                                 '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_026;

  --撤消患者检验申请
  Procedure Zlhis_Cis_036
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO><病人来源>' || 病人来源_In ||
               '</病人来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_036', v_Value);
  End Zlhis_Cis_036;

  --撤消患者检查申请
  Procedure Zlhis_Cis_037
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    v_操作类型 诊疗项目目录.操作类型%Type;
  Begin
    Select MAX(A.操作类型) Into v_操作类型 From 诊疗项目目录 A,病人医嘱记录 B Where B.诊疗项目ID = a.ID And B.ID = 医嘱id_In;
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO><病人来源>' || 病人来源_In ||
               '</病人来源></root>';
    If v_操作类型 = '病理' Then
       b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_055', v_Value);
    Else
       b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_037', v_Value);
    End If;
  End Zlhis_Cis_037;

  --撤消患者手术申请
  Procedure Zlhis_Cis_038
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_038', v_Value);
  End Zlhis_Cis_038;

  --撤消患者输血申请
  Procedure Zlhis_Cis_039
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_039', v_Value);
  End Zlhis_Cis_039;

  --撤消患者会诊申请
  Procedure Zlhis_Cis_040
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_040', v_Value);
  End Zlhis_Cis_040;

  --撤消患者抢救医嘱
  Procedure Zlhis_Cis_041
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_041', v_Value);
  End Zlhis_Cis_041;

  --撤消患者死亡医嘱
  Procedure Zlhis_Cis_042
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_042', v_Value);
  End Zlhis_Cis_042;

  --撤消特殊治疗医嘱
  Procedure Zlhis_Cis_043
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type,
    No_In     In 病人医嘱发送.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_043', v_Value);
  End Zlhis_Cis_043;

  --撤消病区执行医嘱
  Procedure Zlhis_Cis_044
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    No_In       In 病人医嘱发送.No%Type,
    发送数次_In In 病人医嘱发送.发送数次%Type,
    首次时间_In In 病人医嘱发送.首次时间%Type,
    末次时间_In In 病人医嘱发送.末次时间%Type,
    样本条码_In In 病人医嘱发送.样本条码%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><发送号>' || 发送号_In || '</发送号><ID>' ||
               医嘱id_In || '</ID><NO>' || No_In || '</NO><发送数次>' || 发送数次_In || '</发送数次><首次时间>' ||
               To_Char(首次时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</首次时间><末次时间>' ||
               To_Char(末次时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</末次时间><样本条码>' || 样本条码_In || '</样本条码></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_044', v_Value);
  End Zlhis_Cis_044;

  --患者医嘱执行登记
  Procedure Zlhis_Cis_050
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><要求时间>' || To_Char(要求时间_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</要求时间><执行时间>' || To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</执行时间></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_050', v_Value);
  End Zlhis_Cis_050;

  --患者医嘱取消执行登记
  Procedure Zlhis_Cis_051
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    要求时间_In In 病人医嘱执行.要求时间%Type,
    执行时间_In In 病人医嘱执行.执行时间%Type,
    本次数次_In In 病人医嘱执行.本次数次%Type,
    执行结果_In In 病人医嘱执行.执行结果%Type,
    执行摘要_In In 病人医嘱执行.执行摘要%Type,
    执行科室_In In 病人医嘱执行.执行科室id%Type,
    执行人_In   In 病人医嘱执行.执行人%Type,
    核对人_In   In 病人医嘱执行.核对人%Type,
    记录来源_In In 病人医嘱执行.记录来源%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><要求时间>' || To_Char(要求时间_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</要求时间><执行时间>' || To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi:ss') || '</执行时间><本次数次>' || 本次数次_In ||
               '</本次数次><执行结果>' || 执行结果_In || '</执行结果><执行摘要>' || 执行摘要_In || '</执行摘要><执行科室ID>' || 执行科室_In ||
               '</执行科室ID><执行人>' || 执行人_In || '</执行人><核对人>' || 核对人_In || '</核对人><记录来源>' || 记录来源_In || '</记录来源></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_051', v_Value);
  End Zlhis_Cis_051;
  --患者医嘱执行完成
  Procedure Zlhis_Cis_052
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_052',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In ||
                                 '</挂号单><发送号>' || 发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_052;
  --患者医嘱撤消执行完成
  Procedure Zlhis_Cis_053
  (
    病人id_In In 病人医嘱记录.病人id%Type,
    主页id_In In 病人医嘱记录.主页id%Type,
    挂号单_In In 病人挂号记录.No%Type,
    发送号_In In 病人医嘱发送.发送号%Type,
    医嘱id_In In 病人医嘱记录.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_053',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In ||
                                 '</挂号单><发送号>' || 发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID></root>');
  End Zlhis_Cis_053;

  --病理申请发送后修改
  Procedure Zlhis_Cis_056
  (
    病人id_In   In 病人医嘱记录.病人id%Type,
    主页id_In   In 病人医嘱记录.主页id%Type,
    挂号单_In   In 病人挂号记录.No%Type,
    发送号_In   In 病人医嘱发送.发送号%Type,
    医嘱id_In   In 病人医嘱记录.Id%Type,
    病人来源_In In 病人医嘱记录.病人来源%Type --1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    v_操作类型 诊疗项目目录.操作类型%Type;
  Begin
    Select MAX(A.操作类型) Into v_操作类型 From 诊疗项目目录 A,病人医嘱记录 B Where B.诊疗项目ID = a.ID And B.ID = 医嘱id_In;
    v_Value := '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><挂号单>' || 挂号单_In || '</挂号单><发送号>' ||
               发送号_In || '</发送号><ID>' || 医嘱id_In || '</ID><病人来源>' || 病人来源_In || '</病人来源></root>';
     b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_056', v_Value);
  End Zlhis_Cis_056;

  --26.检查报告完成，检查完成时
  Procedure Zlhis_Pacs_001
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_001') = 0 Then
      Return;
    End If;
    For R In (Select '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><报告ID>' || Column_Value || '</报告ID><报告类型>' || 报告类型_In ||
                      '<报告类型></root>' As Xml_Value
              From Table(f_Str2list(报告id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_001', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_001;
  --27.检查状态同步，检查状态改变后
  Procedure Zlhis_Pacs_002
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><原状态>' || 原状态_In || '</原状态><新状态>' || 新状态_In || '</新状态></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_002', v_Value);
  End Zlhis_Pacs_002;
  --28.检查状态回退，检查状态回退后
  Procedure Zlhis_Pacs_003
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    原状态_In In 病人医嘱发送.执行过程%Type,
    新状态_In In 病人医嘱发送.执行过程%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><原状态>' || 原状态_In || '</原状态><新状态>' || 新状态_In || '</新状态></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_003', v_Value);
  End Zlhis_Pacs_003;
  --29.检查报告撤销，撤销检查完成时
  Procedure Zlhis_Pacs_004
  (
    医嘱id_In   In 影像检查记录.医嘱id%Type,
    报告id_Ins  In Varchar2,
    报告类型_In In Number --1-老版PACS报告，2-老版病历编辑器报告，3-新版编辑器报告
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_004') = 0 Then
      Return;
    End If;
    For R In (Select '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><报告ID>' || Column_Value || '</报告ID><报告类型>' || 报告类型_In ||
                      '<报告类型></root>' As Xml_Value
              From Table(f_Str2list(报告id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_004', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_004;
  --30.检查危急值通知，检查发生危急值时
  Procedure Zlhis_Pacs_005(医嘱id_In In 影像检查记录.医嘱id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_005', v_Value);
  End Zlhis_Pacs_005;
  -- 检查预约通知，检查预约时
  Procedure Zlhis_Pacs_006
  (
    医嘱id_In In 影像检查记录.医嘱id%Type,
    预约id_In In Ris检查预约.预约id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><预约ID>' || 预约id_In || '</预约ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_006', v_Value);
  End Zlhis_Pacs_006;
  -- 取消检查预约，取消预约时
  Procedure Zlhis_Pacs_007
  (
    医嘱id_In       In 影像检查记录.医嘱id%Type,
    预约id_In       In Ris检查预约.预约id%Type,
    预约日期_In     In Ris检查预约.预约日期%Type,
    预约序号_In     In Ris检查预约.序号%Type,
    检查设备名称_In In Ris检查预约.检查设备名称%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID><预约ID>' || 预约id_In || '</预约ID><预约日期>' || 预约日期_In || '</预约日期><预约序号>' ||
               预约序号_In || '</预约序号><检查设备名称>' || 检查设备名称_In || '</检查设备名称></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_007', v_Value);
  End Zlhis_Pacs_007;


  --36.患者发卡或绑定卡
  Procedure Zlhis_Patient_018
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><卡号>' || 卡号_In || '</卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_018', v_Value);
  End;

  --37.患者退卡
  Procedure Zlhis_Patient_019
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    卡号_In     In 病人医疗卡信息.卡号%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><卡号>' || 卡号_In || '</卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_019', v_Value);
  End;

  --38.患者补卡/换卡
  Procedure Zlhis_Patient_020
  (
    变动id_In   In 病人变动记录.Id%Type,
    病人id_In   In 病人信息.病人id%Type,
    卡类别id_In In 医疗卡类别.Id%Type,
    原卡号_In   In 病人医疗卡信息.卡号%Type,
    新卡号_In   In 病人医疗卡信息.卡号%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动ID>' || 变动id_In || '</变动ID><病人ID>' || 病人id_In || '</病人ID><卡类别ID>' || 卡类别id_In ||
               '</卡类别ID><原卡号>' || 原卡号_In || '</原卡号><新卡号>' || 新卡号_In || '</新卡号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_020', v_Value);
  End;

  --39.病人挂号登记（包含预约登记)
  Procedure Zlhis_Regist_001
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_001', v_Value);
  End;

  --40.病人分诊
  Procedure Zlhis_Regist_002
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type,
    诊室_In   In 病人挂号记录.诊室%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO><诊室>' || Nvl(诊室_In, '') || '</诊室></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_002', v_Value);
  End;

  --41.病人退号（含取消预约)
  Procedure Zlhis_Regist_003
  (
    挂号id_In In 病人挂号记录.Id%Type,
    No_In     In 病人挂号记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><挂号ID>' || 挂号id_In || '</挂号ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_003', v_Value);
  End;

  --42.临床出诊安排调整
  Procedure Zlhis_Regist_004
  (
    变动原因_In In Integer, --1-停诊;2-替诊;3-诊室变动
    记录id_In   In 临床出诊记录.Id%Type,
    变动id_In   In 临床出诊变动记录.Id%Type

  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><变动原因>' || 变动原因_In || '</变动原因><记录ID>' || 记录id_In || '</记录ID><变动ID>' || 变动id_In ||
               '</变动ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_004', v_Value);
  End;

  --43.门诊患者挂号换号操作
  Procedure Zlhis_Regist_005
  (
    No_In         In 病人挂号记录.No%Type,
    变动原因_In   Integer, --1-替诊;2-换号;3-预约日期变动,
    就诊变动id_In 就诊变动记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><NO>' || No_In || '</NO><变动原因>' || 变动原因_In || '</变动原因><就诊变动ID>' || 就诊变动id_In ||
               '</就诊变动ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_005', v_Value);
  End;


  --费用门诊收费及补充结算
  Procedure Zlhis_Charge_002
  (
    结算类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --结算类型_In:1-收费结算，2-补充结算
    v_Value := '<root><结算类型>' || 结算类型_In || '</结算类型><结帐ID>' || 结帐id_In || '</结帐ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_002', v_Value);
  End;


  --46.门诊退费单据
  Procedure Zlhis_Charge_004
  (
    退费类型_In In Number,
    结帐id_In   In 门诊费用记录.结帐id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --退费类型_In:1-收费结算，2-补充结算
    v_Value := '<root><退费类型>' || 退费类型_In || '</退费类型><结帐ID>' || 结帐id_In || '</结帐ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_004', v_Value);
  End;

  --47.收预交款
  Procedure Zlhis_Charge_005
  (
    预交id_In In 病人预交记录.Id%Type,
    单据号_In In 病人预交记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><预交ID>' || 预交id_In || '</预交ID><单据号>' || 单据号_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_005', v_Value);
  End;

  --48.退预交款(包含负数退预交款部分)
  Procedure Zlhis_Charge_006
  (
    退预交id_In In 病人预交记录.Id%Type,
    单据号_In   In 病人预交记录.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><退预交ID>' || 退预交id_In || '</退预交ID><单据号>' || 单据号_In || '</单据号></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_006', v_Value);
  End;

  --住院记帐单据
  Procedure Zlhis_Charge_007
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><收费类别>' || 收费类别_In || '</收费类别><费用ID>' || 费用id_In || '</费用ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_007', v_Value);
  End;

  --住院记帐单据销账
  Procedure Zlhis_Charge_008
  (
    收费类别_In In 住院费用记录.收费类别%Type,
    费用id_In   In 住院费用记录.Id%Type,
    收发ids_In  In Varchar2 := Null --可能费用ID对应多个收发id，对应格式：收发id,数量|收发id,数量；非药品不传
  ) Is
    v_Value   Zlmsg_Todo.Key_Value%Type;
    v_Tmp     Varchar2(4000);
    v_Infotmp Varchar2(4000);
    v_Fields  Varchar2(4000);
    v_收发id  Varchar2(50);
    v_数量    Varchar2(20);
  Begin
    If p_Msg_Using('ZLHIS_CHARGE_008') = 0 Then
      Return;
    End If;
    v_Value := '<root><收费类别>' || 收费类别_In || '</收费类别><费用ID>' || 费用id_In || '</费用ID>';

    If 收发ids_In Is Null Then
      v_Infotmp := Null;
      v_Tmp     := '<收发IDS>' || '<收发ID>' || '</收发ID>' || '<数量>' || '</数量>' || '</收发IDS>';
    Else
      v_Infotmp := 收发ids_In || '|';
      While v_Infotmp Is Not Null Loop
        --分解收发ID串
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        v_收发id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        v_数量    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');

        v_Tmp := v_Tmp || '<收发IDS>' || '<收发ID>' || v_收发id || '</收发ID>' || '<数量>' || v_数量 || '</数量>' || '</收发IDS>';
      End Loop;
    End If;

    v_Value := v_Value || v_Tmp || '</root>';

    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_008', v_Value);
  End;

  --53.住院患者入院登记
  Procedure Zlhis_Patient_001
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_001') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 开始原因 = 1 And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_001',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_001;
  --54.住院患者入院入科
  Procedure Zlhis_Patient_002
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_002') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_002',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_002;
  --56.住院患者床位变更
  Procedure Zlhis_Patient_004
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_原床号   Varchar2(255);
    v_新床号   Varchar2(255);
    n_变动id   Number(18);
    n_开始原因 Number(3);
    d_开始时间 Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_004') = 0 Then
      Return;
    End If;
    Select ID, 床号, 开始时间, 开始原因
    Into n_变动id, v_新床号, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;

    Select Max(床号)
    Into v_原床号
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_004',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原床号>' ||
                                 v_原床号 || '</原床号>' || '<新床号>' || v_新床号 || '</新床号>' || '<变动ID>' || n_变动id || '</变动ID>' ||
                                 '</root>');
  End Zlhis_Patient_004;
  --57.住院患者病情变更
  Procedure Zlhis_Patient_005
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_005') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_005',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_005;
  --58.住院患者变更撤消
  Procedure Zlhis_Patient_006
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    撤销方式_In In Varchar2
  ) Is
    n_科室id     病人变动记录.科室id%Type;
    n_病区id     病人变动记录.病区id%Type;
    n_护理等级id 病人变动记录.护理等级id%Type;
    n_医疗小组id 病人变动记录.医疗小组id%Type;
    v_床号       病人变动记录.床号%Type;
    v_责任护士   病人变动记录.责任护士%Type;
    v_主任医师   病人变动记录.主任医师%Type;
    v_主治医师   病人变动记录.主治医师%Type;
    v_经治医师   病人变动记录.经治医师%Type;
    v_病情       病人变动记录.病情%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_006') = 0 Then
      Return;
    End If;
    Select Max(科室id), Max(病区id), Max(护理等级id), Max(医疗小组id), Max(床号), Max(责任护士), Max(主任医师), Max(主治医师), Max(经治医师), Max(病情)
    Into n_科室id, n_病区id, n_护理等级id, n_医疗小组id, v_床号, v_责任护士, v_主任医师, v_主治医师, v_经治医师, v_病情
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And (终止时间 Is Null Or 终止原因 = 1) And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_006',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><撤销方式>' || 撤销方式_In ||
                                 '</撤销方式><科室ID>' || n_科室id || '</科室ID>' || '<病区ID>' || n_病区id || '</病区ID>' || '<护理等级ID>' ||
                                 n_护理等级id || '</护理等级ID>' || '<医疗小组ID>' || n_医疗小组id || '</医疗小组ID>' || '<床号>' || v_床号 ||
                                 '</床号>' || '<责任护士>' || v_责任护士 || '</责任护士>' || '<主任医师>' || v_主任医师 || '</主任医师>' ||
                                 '<主治医师>' || v_主治医师 || '</主治医师>' || '<经治医师>' || v_经治医师 || '</经治医师>' || '<病情>' || v_病情 ||
                                 '</病情>' || '</root>');
  End Zlhis_Patient_006;
  --59.住院患者医护变更
  Procedure Zlhis_Patient_007
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_原住院医生 Varchar2(100);
    v_新住院医生 Varchar2(100);
    v_原主治医生 Varchar2(100);
    v_新主治医生 Varchar2(100);
    v_原主任医生 Varchar2(100);
    v_新主任医生 Varchar2(100);
    v_原责任护士 Varchar2(100);
    v_新责任护士 Varchar2(100);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_007') = 0 Then
      Return;
    End If;
    Select ID, 经治医师, 主治医师, 主任医师, 责任护士, 开始时间, 开始原因
    Into n_变动id, v_新住院医生, v_新主治医生, v_新主任医生, v_新责任护士, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;

    Select Max(经治医师), Max(主治医师), Max(主任医师), Max(责任护士)
    Into v_原住院医生, v_原主治医生, v_原主任医生, v_原责任护士
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_007',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原住院医生>' ||
                                 v_原住院医生 || '</原住院医生>' || '<新住院医生>' || v_新住院医生 || '</新住院医生>' || '<原主治医生>' || v_原主治医生 ||
                                 '</原主治医生>' || '<新主治医生>' || v_新主治医生 || '</新主治医生>' || '<原主任医生>' || v_原主任医生 || '</原主任医生>' ||
                                 '<新主任医生>' || v_新主任医生 || '</新主任医生>' || '<原责任护士>' || v_原责任护士 || '</原责任护士>' || '<新责任护士>' ||
                                 v_新责任护士 || '</新责任护士>' || '<变动ID>' || n_变动id || '</变动ID>' || '</root>');
  End Zlhis_Patient_007;
  --住院患者护理等级变更
  Procedure Zlhis_Patient_008
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_原护理等级id Number(18);
    v_新护理等级id Number(18);
    n_变动id       Number(18);
    n_开始原因     Number(3);
    d_开始时间     Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_008') = 0 Then
      Return;
    End If;
    Select ID, 护理等级id, 开始时间, 开始原因
    Into n_变动id, v_新护理等级id, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;

    Select Max(护理等级id)
    Into v_原护理等级id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_008',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<原护理等级ID>' ||
                                 v_原护理等级id || '</原护理等级ID>' || '<新护理等级ID>' || v_新护理等级id || '</新护理等级ID>' || '<变动ID>' ||
                                 n_变动id || '</变动ID>' || '</root>');
  End Zlhis_Patient_008;
  --60.住院患者预出院
  Procedure Zlhis_Patient_009
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    n_变动id 病人变动记录.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_009') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_变动id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_009',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><变动ID>' || n_变动id ||
                                 '</变动ID></root>');
  End Zlhis_Patient_009;
  --61.住院患者出院
  Procedure Zlhis_Patient_010
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_010',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID></root>');
  End Zlhis_Patient_010;
  --62.住院患者新生儿登记
  Procedure Zlhis_Patient_011
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_011',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><婴儿序号>' || 婴儿序号_In ||
                                 '</婴儿序号></root>');
  End Zlhis_Patient_011;
  --63.住院患者转入科室
  Procedure Zlhis_Patient_012
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_转出科室id Number(18);
    v_转入科室id Number(18);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_012') = 0 Then
      Return;
    End If;
    Select ID, 科室id, 开始时间, 开始原因
    Into n_变动id, v_转入科室id, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;

    Select Max(科室id)
    Into v_转出科室id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_012',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<转出科室ID>' ||
                                 v_转出科室id || '</转出科室ID>' || '<转入科室ID>' || v_转入科室id || '</转入科室ID>' || '<变动ID>' || n_变动id ||
                                 '</变动ID>' || '</root>');
  End Zlhis_Patient_012;
  --64.新生儿登记作废
  Procedure Zlhis_Patient_013
  (
    病人id_In   In 病案主页.病人id%Type,
    主页id_In   In 病案主页.主页id%Type,
    婴儿序号_In 病人医嘱记录.婴儿%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_013',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID><婴儿序号>' || 婴儿序号_In ||
                                 '</婴儿序号></root>');
  End Zlhis_Patient_013;
  --65.门诊患者登记
  Procedure Zlhis_Patient_015(病人id_In In 病案主页.病人id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_015', '<root><病人ID>' || 病人id_In || '</病人ID></root>');
  End Zlhis_Patient_015;
  --66.患者信息修改
  Procedure Zlhis_Patient_016(病人id_In In 病案主页.病人id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_016', '<root><病人ID>' || 病人id_In || '</病人ID></root>');
  End Zlhis_Patient_016;

  --67.患者合并
  Procedure Zlhis_Patient_017
  (
    病人id_In   In 病案主页.病人id%Type,
    原病人id_In In 病案主页.病人id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017',
                                '<root><病人ID>' || 病人id_In || '</病人ID><原病人ID>' || 原病人id_In || '</原病人ID></root>');
  End Zlhis_Patient_017;

  --69.住院患者转入病区
  Procedure Zlhis_Patient_026
  (
    病人id_In In 病案主页.病人id%Type,
    主页id_In In 病案主页.主页id%Type
  ) Is
    v_转出病区id Number(18);
    v_转入病区id Number(18);
    n_变动id     Number(18);
    n_开始原因   Number(3);
    d_开始时间   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_026') = 0 Then
      Return;
    End If;
    Select ID, 病区id, 开始时间, 开始原因
    Into n_变动id, v_转入病区id, d_开始时间, n_开始原因
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 Is Null And Nvl(附加床位, 0) = 0;

    Select Max(病区id)
    Into v_转出病区id
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 终止时间 = d_开始时间 And 终止原因 = n_开始原因 And Nvl(附加床位, 0) = 0;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_026',
                                '<root><病人ID>' || 病人id_In || '</病人ID><主页ID>' || 主页id_In || '</主页ID>' || '<转出病区ID>' ||
                                 v_转出病区id || '</转出病区ID>' || '<转入病区ID>' || v_转入病区id || '</转入病区ID>' || '<变动ID>' || n_变动id ||
                                 '</变动ID>' || '</root>');
  End Zlhis_Patient_026;

  Procedure Zlhis_Patient_028(病人id_In In 病案主页.病人id%Type) Is
    v_姓名     病人信息.姓名%Type;
    v_性别     病人信息.性别%Type;
    v_年龄     病人信息.年龄%Type;
    v_出生日期 病人信息.出生日期%Type;
    v_门诊号   病人信息.门诊号%Type;
    v_身份证号 病人信息.身份证号%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_028') = 0 Then
      Return;
    End If;
    Select 姓名, 性别, 年龄, 出生日期, 门诊号, 身份证号
    Into v_姓名, v_性别, v_年龄, v_出生日期, v_门诊号, v_身份证号
    From 病人信息
    Where 病人id = 病人id_In;

    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_028',
                                '<root><病人ID>' || 病人id_In || '</病人ID><姓名>' || v_姓名 || '</姓名>' || '<性别>' || v_性别 ||
                                 '</性别>' || '<年龄>' || v_年龄 || '</年龄>' || '<出生日期>' || v_出生日期 || '</出生日期>' || '<门诊号>' ||
                                 v_门诊号 || '</门诊号>' || '<身份证号>' || v_身份证号 || '</身份证号>' || '</root>');
  End Zlhis_Patient_028;

  --血库:科室配血完成
  Procedure Zlhis_Blood_001(医嘱id_In In 病人医嘱记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_001', v_Value);
    End If;
  End Zlhis_Blood_001;

  --血库:科室拒绝配血
  Procedure Zlhis_Blood_002(医嘱id_In In 病人医嘱记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><医嘱ID>' || 医嘱id_In || '</医嘱ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_002', v_Value);
    End If;
  End Zlhis_Blood_002;

  --70.检验报告审核
  Procedure Zlhis_Lis_001(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_001', v_Value);
    End If;
  End Zlhis_Lis_001;
  --71.检验报告审核撤消
  Procedure Zlhis_Lis_002(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_002', v_Value);
    End If;
  End Zlhis_Lis_002;
  --73.检验标本条码打印
  Procedure Zlhis_Lis_004
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_004') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', v_Value);
    Else
      For R In (Select '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>' As Xml_Value
                From 病人医嘱发送
                Where 医嘱id In (Select Column_Value From Table(f_Num2list(医嘱ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_004;
  --74.检验标本条码打印撤销
  Procedure Zlhis_Lis_005
  (
    样本条码_In In 病人医嘱发送.样本条码%Type,
    医嘱id_In   In 病人医嘱发送.医嘱id%Type,
    医嘱ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_005') = 0 Then
      Return;
    End If;
    If 医嘱id_In Is Not Null Then
      v_Value := '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', v_Value);
    Else
      For R In (Select '<root><样本条码>' || 样本条码_In || '</样本条码><医嘱ID>' || 医嘱id_In || '</医嘱ID><系统>1</系统></root>' As Xml_Value
                From 病人医嘱发送
                Where 医嘱id In (Select Column_Value From Table(f_Num2list(医嘱ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_005;
  --75.检验标本核收
  Procedure Zlhis_Lis_006(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_006', v_Value);
    End If;
  End Zlhis_Lis_006;
  --76.检验标本核收撤销
  Procedure Zlhis_Lis_007(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_007', v_Value);
    End If;
  End Zlhis_Lis_007;
  --77.检验标本拒收
  Procedure Zlhis_Lis_008(标本id_In In 检验标本记录.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If 标本id_In Is Not Null Then
      v_Value := '<root><标本ID>' || 标本id_In || '</标本ID><系统>1</系统></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_008', v_Value);
    End If;
  End Zlhis_Lis_008;

End b_Message;
/


--122954:余伟节,2018-03-26,中联合理用药

Create Or Replace Procedure Zl_中联合理用药参数_Update
(
  病人id_In In 病人信息.病人id%Type,
  主页id_In In 病案主页.主页id%Type,
  挂号id_In In 病人挂号记录.Id%Type := Null
) As

  --------------------------------------------------------------------------------------------------
  --功能:合理用药监测传入值；返回合理用药传入值
  --出参:Xml_Return  返回补充的XML串
  -- <details_xml>
  --    <patient_info>
  --      <info name="年龄数字" value="28114.45"/>
  --     <info name="年龄周期" value="成人"/>
  --      <info name="性别" value="女"/>
  --      <info name="职业" value="运动员"/>
  --      <info name="妊娠" value="1"/>
  --      <info name="哺乳" value="1"/>
  --      <info name="肝功能不全" value="1">
  --      <info name="严重肝功能不全" value="1">
  --      <info name="肾功能不全" value="1">
  --      <info name="严重肾功能不全" value="1">
  --      <info name="诊断" value="J18.000"/> --诊断传编码，多个诊断以逗号分隔
  --    </patient_info>
  --    <medicine_info>
  --      <medicine>
  --        <info name="医嘱ID" value="1"/>
  --        <info name="本位码" value="86900967000160" main="46d64420-8319-4768-9a11-f4b0f5e4ce7a"/> --main值是固定的
  --        <info name="诊疗项目ID" value="67232" main="4e19df1c-c1b9-4a43-a83d-0741a19961ab"/>
  --        <info name="输液组号" value="1"/>
  --        <info name="计量单位" value="ml"/>
  --        <info name="单次量" value="250"/>
  --        <info name="单次量-按体重" value="5.21"/>--单次量-按体重= trunc(单次量/病人体重,2)
  --        <info name="单次量-按体表" value="170.3"/>--单次量-按体表= trunc(单次量/(0.0061*病人身高+0.0128*病人体重-0.1529),2)
  --        <info name="每日量" value="250"/>
  --        1.每日量=单次量*日频次
  --        2.日频次计算：
  --            a.适用范围=-1，日频次=1
  --            b.间隔单位=天 and 频率间隔=1，日频次=频率次数
  --            c.间隔单位=天 and 频率间隔>1 and 频率次数=1，日频次=1
  --            d.间隔单位=小时 and 频率间隔<=24,日频次=24/频率间隔*频率次数
  --            e.间隔单位=小时 and 频率间隔>24 and 频率次数=1，日频次=1
  --            f.间隔单位=周 and 频率次数=1，日频次=1
  --        <info name="每日量-按体重" value="5.21"/>  --trunc(每日量/病人体重,2)
  --        <info name="每日量-按体表" value="170.3"/>  --每日量-按体表= trunc(每日量/(0.0061*病人身高+0.0128*病人体重-0.1529),2)
  --        <info name="给药频次" value="每日一次"/>
  --        <info name="给药途径" value="001"/>
  --      </medicine>
  --    </medicine_info>
  --  </details_xml>
  --------------------------------------------------------------------------------------------------
  Xml_Ret             Xmltype;
  Xml_Document        Xmldom.Domdocument;
  Xml_Nodelist        Xmldom.Domnodelist;
  Xml_Domelement      Xmldom.Domelement;
  Xml_Domnamednodemap Xmldom.Domnamednodemap;
  Xml_Node_Med        Xmldom.Domnode;
  Xml_Node            Xmldom.Domnode;
  Xml_Node_New        Xmldom.Domnode;
  ----------------------------------
  n_身高 Number(10, 2); --单位:cm
  n_体重 Number(10, 2); --体重:KG

  l_Clob    Clob;
  v_Err_Msg Varchar2(2000);
  v_Temp    Varchar2(200);
  v_Value   Varchar2(200);
  n_Nodenum Number(5);
  Err_Item Exception;
Begin
  --：
  --将CLOB数据提取到v_XML中
  Select 参数内容 Into l_Clob From 中联合理用药参数;
  Xml_Ret        := Xmltype(l_Clob); --缓存函数返回值
  Xml_Document   := Xmldom.Newdomdocument(Xml_Ret);
  Xml_Domelement := Xmldom.Getdocumentelement(Xml_Document);
  Xml_Nodelist   := Xmldom.Getelementsbytagname(Xml_Domelement, 'patient_info');
  --获取patient_info/INfo节点
  Xml_Nodelist := Xmldom.Getchildnodes(Xmldom.Item(Xml_Nodelist, 0));
  n_Nodenum    := Xmldom.Getlength(Xml_Nodelist);
  For I In 0 .. n_Nodenum - 1 Loop
    Xml_Domnamednodemap := Xmldom.Getattributes(Xmldom.Item(Xml_Nodelist, I));
    v_Temp              := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'name'));
    If v_Temp = '身高' Then
      n_身高 := Nvl(To_Number(Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'))), 0);
    End If;
    If v_Temp = '体重' Then
      n_体重 := Nvl(To_Number(Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'))), 0);
    End If;
  End Loop;
  --获取medicine/INfo节点

  Xml_Nodelist := Xmldom.Getelementsbytagname(Xml_Domelement, 'medicine');
  n_Nodenum    := Xmldom.Getlength(Xml_Nodelist);
  For I In 0 .. n_Nodenum - 1 Loop
    Xml_Node_Med := Xmldom.Item(Xml_Nodelist, I); --取第一个孩子节点medicine
    Xml_Nodelist := Xmldom.Getchildnodes(Xml_Node_Med); --infos
    Xml_Node     := Xmldom.Getfirstchild(Xml_Node_Med); --取第一个孩子节点
    While Not Xmldom.Isnull(Xml_Node) Loop
      Xml_Domnamednodemap := Xmldom.Getattributes(Xml_Node);
      v_Temp              := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'name'));
      v_Value             := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'));
      If v_Temp = '单次量' Then
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '单次量-按体重');
        If n_体重 > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', Trunc(To_Number(v_Value) / n_体重, 2));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
        --单次量-按体表trunc(每日量/(0.0061*病人身高+0.0128*病人体重-0.1529),2)
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '单次量-按体表');
        If n_体重 > 0 And n_身高 > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              Trunc(To_Number(v_Value) / (0.0061 * n_身高 + 0.0128 * n_体重 - 0.1529), 2));
        
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      End If;
    
      If v_Temp = '每日量' Then
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '每日量-按体重');
        If n_体重 > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', Trunc(To_Number(v_Value) / n_体重, 2));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      
        --每日量-按体表
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '每日量-按体表');
        If n_体重 > 0 And n_身高 > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              Trunc(To_Number(v_Value) / (0.0061 * n_身高 + 0.0128 * n_体重 - 0.1529), 2));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      End If;
      --取下一个兄弟节点
      Xml_Node := Xmldom.Getnextsibling(Xml_Node);
    End Loop;
  End Loop;

  --将函数返回值存入临时表,ZLHIS在事物结束前提取（因为受驱动限制返回值不能超过4000限制）
  Update 中联合理用药参数 Set 参数内容 = Xml_Ret.Getclobval();

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_中联合理用药参数_Update;
/

--122954:余伟节,2018-03-26,中联合理用药

Create Or Replace Function Zl_Read_中联合理用药参数(Pos_In In Number
                                            --参数说明：
                                            --Pos_In：从0开始不断读取，直到返回为空
                                            ) Return Varchar2 Is
  l_Clob    Clob;
  v_Buffer  Varchar2(32767);
  n_Amount  Number := 2000;
  n_Offset  Number := 1;
  v_Err_Msg Varchar2(2000);
  Err_Item Exception;
Begin
  Select 参数内容 Into l_Clob From 中联合理用药参数;
  n_Offset := n_Offset + Pos_In * n_Amount;

  If l_Clob Is Null Then
    v_Buffer := Null;
  Else
    Begin
      Dbms_Lob.Read(l_Clob, n_Amount, n_Offset, v_Buffer);
    Exception
      When No_Data_Found Then
        v_Buffer := Null;
    End;
  End If;
  Return v_Buffer;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Read_中联合理用药参数;
/

------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0004' Where 编号=&n_System;
Commit;