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
--0:梁唐彬,2018-05-18,集成平台消息调整；
Create Or Replace Procedure Zl_入院病案主页_Insert
(
  登记模式_In       Number,
  病人性质_In       病案主页.病人性质%Type,
  病人id_In         病人信息.病人id%Type,
  住院号_In         病人信息.住院号%Type,
  医保号_In         保险帐户.医保号%Type,
  姓名_In           病人信息.姓名%Type,
  性别_In           病人信息.性别%Type,
  年龄_In           病人信息.年龄%Type,
  费别_In           病人信息.费别%Type,
  出生日期_In       病人信息.出生日期%Type,
  国籍_In           病人信息.国籍%Type,
  民族_In           病人信息.民族%Type,
  学历_In           病人信息.学历%Type,
  婚姻状况_In       病人信息.婚姻状况%Type,
  职业_In           病人信息.职业%Type,
  身份_In           病人信息.身份%Type,
  身份证号_In       病人信息.身份证号%Type,
  出生地点_In       病人信息.出生地点%Type,
  家庭地址_In       病人信息.家庭地址%Type,
  家庭地址邮编_In   病人信息.家庭地址邮编%Type,
  家庭电话_In       病人信息.家庭电话%Type,
  户口地址_In       病人信息.户口地址%Type,
  户口地址邮编_In   病人信息.户口地址邮编%Type,
  联系人姓名_In     病人信息.联系人姓名%Type,
  联系人关系_In     病人信息.联系人关系%Type,
  联系人地址_In     病人信息.联系人地址%Type,
  联系人电话_In     病人信息.联系人电话%Type,
  工作单位_In       病人信息.工作单位%Type,
  合同单位id_In     病人信息.合同单位id%Type,
  单位电话_In       病人信息.单位电话%Type,
  单位邮编_In       病人信息.单位邮编%Type,
  单位开户行_In     病人信息.单位开户行%Type,
  单位帐号_In       病人信息.单位帐号%Type,
  担保人_In         病人信息.担保人%Type,
  担保额_In         病人信息.担保额%Type,
  担保性质_In       病人信息.担保性质%Type,
  入院科室id_In     病案主页.入院科室id%Type,
  护理等级id_In     病案主页.护理等级id%Type,
  入院病况_In       病案主页.入院病况%Type,
  入院方式_In       病案主页.入院方式%Type,
  住院目的_In       病案主页.住院目的%Type,
  二级院转入_In     病案主页.二级院转入%Type,
  门诊医师_In       病案主页.门诊医师%Type,
  籍贯_In           病人信息.籍贯%Type,
  区域_In           病案主页.区域%Type,
  入院时间_In       病案主页.入院日期%Type,
  是否陪伴_In       病案主页.是否陪伴%Type,
  床号_In           病案主页.入院病床%Type,
  付款方式_In       病案主页.医疗付款方式%Type,
  疾病id_In         病人诊断记录.疾病id%Type,
  诊断id_In         病人诊断记录.诊断id%Type,
  门诊诊断_In       病人诊断记录.诊断描述%Type,
  中医疾病id_In     病人诊断记录.疾病id%Type,
  中医诊断id_In     病人诊断记录.诊断id%Type,
  中医诊断_In       病人诊断记录.诊断描述%Type,
  险类_In           病案主页.险类%Type,
  操作员编号_In     病案主页.编目员编号%Type,
  操作员姓名_In     病案主页.编目员姓名%Type,
  新病人_In         Number := 1,
  备注_In           病案主页.备注%Type,
  入院病区id_In     病案主页.入院病区id%Type,
  再入院_In         病案主页.再入院%Type,
  入院属性_In       病案主页.入院属性%Type := Null,
  主页id_In         病案主页.主页id%Type := Null,
  住院次数_In       病人信息.住院次数%Type := Null,
  其他证件_In       病人信息.其他证件%Type := Null,
  病人类型_In       病案主页.病人类型%Type := Null,
  联系人身份证号_In 病人信息.联系人身份证号%Type := Null,
  手机号_In         病人信息.手机号%Type := Null,
  挂号id_In         病案主页.挂号id%Type := Null
) As
  -----------------------------------------------------------
  --功能：对入院病人新增一张病案主页，同时可能处理入科。
  --参数：
  --      登记模式_IN=0-正常登记,1-预约登记,2-接收预约(新病人_IN=0)
  --      病人性质_IN=对应"病案主页.病人性质"
  --      床号_IN=Null:不同时入科;'家庭病床':分配家庭病床,填为空;其他:分配具体床位。
  --      新病人_IN=如果是已有档案的病人入院,则该参数为0；缺省为新病人
  --      入院病区ID_IN=只有当使用[病区管理病床]模式(参数号99)时,并且入院同时入科分床时,才有值
  --      住院号_In = 登记门诊留观病人时 住院号_In 为病人门诊号
  -----------------------------------------------------------
  v_主页id   病案主页.主页id%Type;
  v_等级id   床位状况记录.等级id%Type;
  n_住院次数 病人信息.住院次数%Type;

  v_费别      病案主页.费别%Type;
  v_Count     Number;
  n_Uniqueid  Number;
  v_Date      Date;
  d_Indeptime Date;
  v_Error     Varchar2(255);
  Err_Custom Exception;
Begin
  --判断病人是否锁定
  Select Count(病人id) Into v_Count From 病人信息 Where 病人id = 病人id_In;
  If v_Count <> 0 Then
    Zl_病人信息_锁定检查(病人id_In);
  End If;

  Select Sysdate Into v_Date From Dual;
  Zl_病区标记记录_Clear(病人id_In);

  --身份证号不等于空,根据系统参数判读是否唯一建档病人
  If 身份证号_In Is Not Null Then
    n_Uniqueid := Nvl(zl_GetSysParameter(279), 0);
    If n_Uniqueid = 1 Then
      Select Count(1) Into v_Count From 病人信息 Where 身份证号 = 身份证号_In And 病人id <> Nvl(病人id_In, 0);
      If v_Count <> 0 Then
        v_Error := '已经存在身份证号为' || 身份证号_In || '的病人,不能再录入相同的身份证号!';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  --病人基本信息
  If 病人性质_In = 1 Then
    If 新病人_In = 1 Then
      Insert Into 病人信息
        (病人id, 门诊号, 住院号, 姓名, 性别, 年龄, 费别, 医疗付款方式, 出生日期, 国籍, 民族, 籍贯, 区域, 学历, 婚姻状况, 职业, 身份, 身份证号, 出生地点, 家庭地址, 家庭地址邮编, 家庭电话,
         户口地址, 户口地址邮编, 联系人姓名, 联系人关系, 联系人地址, 联系人电话, 工作单位, 合同单位id, 单位电话, 单位邮编, 单位开户行, 单位帐号, 担保人, 担保额, 担保性质, 险类, 登记时间, 其他证件,
         病人类型, 联系人身份证号, 手机号)
      Values
        (病人id_In, 住院号_In, Null, 姓名_In, 性别_In, 年龄_In, 费别_In, 付款方式_In, 出生日期_In, 国籍_In, 民族_In, 籍贯_In, 区域_In, 学历_In,
         婚姻状况_In, 职业_In, 身份_In, 身份证号_In, 出生地点_In, 家庭地址_In, 家庭地址邮编_In, 家庭电话_In, 户口地址_In, 户口地址邮编_In, 联系人姓名_In, 联系人关系_In,
         联系人地址_In, 联系人电话_In, 工作单位_In, Decode(合同单位id_In, 0, Null, 合同单位id_In), 单位电话_In, 单位邮编_In, 单位开户行_In, 单位帐号_In, 担保人_In,
         Decode(担保额_In, 0, Null, 担保额_In), 担保性质_In, 险类_In, v_Date, 其他证件_In, 病人类型_In, 联系人身份证号_In, 手机号_In);
    Else
      --老病人的门诊费别不变,除非是门诊留观病人
      Update 病人信息
      Set 门诊号 = 住院号_In, 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 费别 = Decode(病人性质_In, 1, 费别_In, 费别), 医疗付款方式 = 付款方式_In,
          出生日期 = 出生日期_In, 国籍 = 国籍_In, 民族 = 民族_In, 籍贯 = 籍贯_In, 区域 = 区域_In, 学历 = 学历_In, 婚姻状况 = 婚姻状况_In, 职业 = 职业_In,
          身份 = 身份_In, 身份证号 = 身份证号_In, 出生地点 = 出生地点_In, 家庭地址 = 家庭地址_In, 家庭地址邮编 = 家庭地址邮编_In, 家庭电话 = 家庭电话_In, 户口地址 = 户口地址_In,
          户口地址邮编 = 户口地址邮编_In, 联系人姓名 = 联系人姓名_In, 联系人关系 = 联系人关系_In, 联系人地址 = 联系人地址_In, 联系人电话 = 联系人电话_In, 工作单位 = 工作单位_In,
          合同单位id = Decode(合同单位id_In, 0, Null, 合同单位id_In), 单位电话 = 单位电话_In, 单位邮编 = 单位邮编_In, 单位开户行 = 单位开户行_In,
          单位帐号 = 单位帐号_In, 担保人 = 担保人_In, 担保额 = Decode(担保额_In, 0, Null, 担保额_In), 担保性质 = 担保性质_In, 险类 = 险类_In,
          其他证件 = 其他证件_In, 病人类型 = 病人类型_In, 联系人身份证号 = 联系人身份证号_In, 手机号 = Nvl(手机号_In, 手机号)
      Where 病人id = 病人id_In;
    End If;
  Else
    If 新病人_In = 1 Then
      Insert Into 病人信息
        (病人id, 住院号, 姓名, 性别, 年龄, 费别, 医疗付款方式, 出生日期, 国籍, 民族, 籍贯, 区域, 学历, 婚姻状况, 职业, 身份, 身份证号, 出生地点, 家庭地址, 家庭地址邮编, 家庭电话,
         户口地址, 户口地址邮编, 联系人姓名, 联系人关系, 联系人地址, 联系人电话, 工作单位, 合同单位id, 单位电话, 单位邮编, 单位开户行, 单位帐号, 担保人, 担保额, 担保性质, 险类, 登记时间, 其他证件,
         病人类型, 联系人身份证号, 手机号)
      Values
        (病人id_In, Decode(病人性质_In, 2, Null, 住院号_In), 姓名_In, 性别_In, 年龄_In, 费别_In, 付款方式_In, 出生日期_In, 国籍_In, 民族_In, 籍贯_In,
         区域_In, 学历_In, 婚姻状况_In, 职业_In, 身份_In, 身份证号_In, 出生地点_In, 家庭地址_In, 家庭地址邮编_In, 家庭电话_In, 户口地址_In, 户口地址邮编_In,
         联系人姓名_In, 联系人关系_In, 联系人地址_In, 联系人电话_In, 工作单位_In, Decode(合同单位id_In, 0, Null, 合同单位id_In), 单位电话_In, 单位邮编_In,
         单位开户行_In, 单位帐号_In, 担保人_In, Decode(担保额_In, 0, Null, 担保额_In), 担保性质_In, 险类_In, v_Date, 其他证件_In, 病人类型_In,
         联系人身份证号_In, 手机号_In);
    Else
      --老病人的门诊费别不变,除非是门诊留观病人
      Update 病人信息
      Set 住院号 = Decode(病人性质_In, 2, 住院号, Decode(住院号_In, Null, 住院号, 住院号_In)), 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In,
          费别 = Decode(病人性质_In, 1, 费别_In, 费别), 医疗付款方式 = 付款方式_In, 出生日期 = 出生日期_In, 国籍 = 国籍_In, 民族 = 民族_In, 籍贯 = 籍贯_In,
          区域 = 区域_In, 学历 = 学历_In, 婚姻状况 = 婚姻状况_In, 职业 = 职业_In, 身份 = 身份_In, 身份证号 = 身份证号_In, 出生地点 = 出生地点_In, 家庭地址 = 家庭地址_In,
          家庭地址邮编 = 家庭地址邮编_In, 家庭电话 = 家庭电话_In, 户口地址 = 户口地址_In, 户口地址邮编 = 户口地址邮编_In, 联系人姓名 = 联系人姓名_In, 联系人关系 = 联系人关系_In,
          联系人地址 = 联系人地址_In, 联系人电话 = 联系人电话_In, 工作单位 = 工作单位_In, 合同单位id = Decode(合同单位id_In, 0, Null, 合同单位id_In),
          单位电话 = 单位电话_In, 单位邮编 = 单位邮编_In, 单位开户行 = 单位开户行_In, 单位帐号 = 单位帐号_In, 担保人 = 担保人_In,
          担保额 = Decode(担保额_In, 0, Null, 担保额_In), 担保性质 = 担保性质_In, 险类 = 险类_In, 其他证件 = 其他证件_In, 病人类型 = 病人类型_In,
          联系人身份证号 = 联系人身份证号_In, 手机号 = Nvl(手机号_In, 手机号)
      Where 病人id = 病人id_In;
    End If;
  End If;
  If 新病人_In <> 1 then
	b_Message.Zlhis_Patient_016(病人id_In);
  End if;

  --病案信息
  Begin
    If 登记模式_In = 1 Then
      v_主页id := 0; --预约登记记录的主页ID=0
    Else
      If 主页id_In Is Null Then
        Select Nvl(Max(主页id), 0) + 1 Into v_主页id From 病案主页 Where 病人id = 病人id_In And Nvl(主页id, 0) <> 0;
      Else
        v_主页id := 主页id_In;
      End If;
    End If;
  Exception
    When Others Then
      Null;
  End;

  If 登记模式_In <> 1 Then
    Update 病人信息
    Set 主页id = v_主页id, 当前病区id = 入院病区id_In, 当前科室id = 入院科室id_In, 当前床号 = Decode(床号_In, '家庭病床', Null, 床号_In), 入院时间 = 入院时间_In,
        出院时间 = Null, 在院 = 1
    Where 病人id = 病人id_In;
  End If;

  --更新住院次数
  If 登记模式_In <> 1 And 病人性质_In = 0 Then
    If Nvl(住院次数_In, 0) = 0 Then
      Select Nvl(住院次数, 0) + 1 Into n_住院次数 From 病人信息 Where 病人id = 病人id_In;
    Else
      n_住院次数 := 住院次数_In;
    End If;
    Update 病人信息 Set 住院次数 = n_住院次数 Where 病人id = 病人id_In;
  End If;

  --取入科时间
  If 床号_In Is Null Then
    d_Indeptime := Null;
  Else
    d_Indeptime := 入院时间_In;
  End If;

  --状态：0-正常在院,1-等待入科,2-等待转科
  If 登记模式_In = 2 Then
    --处理病案主页从表
    Delete From 病案主页从表 Where 病人id = 病人id_In And Nvl(主页id, 0) = 0;
    --接收预约
    Update 病案主页
    Set 主页id = v_主页id, 病人性质 = 病人性质_In, 住院号 = Decode(病人性质_In, 1, Null, 2, Null, 住院号_In),
        留观号 = Decode(病人性质_In, 2, 住院号_In, Null),
        --主页ID变更,病人性质可能变更
        费别 = 费别_In, 入院病区id = 入院病区id_In, 入院科室id = 入院科室id_In, 入院日期 = 入院时间_In, 入科时间 = d_Indeptime, 入院病况 = 入院病况_In,
        入院方式 = 入院方式_In, 入院属性 = 入院属性_In, 二级院转入 = 二级院转入_In, 住院目的 = 住院目的_In, 入院病床 = Decode(床号_In, '家庭病床', Null, 床号_In),
        是否陪伴 = 是否陪伴_In, 当前病况 = 入院病况_In, 当前病区id = 入院病区id_In, 护理等级id = Decode(护理等级id_In, 0, Null, 护理等级id_In),
        出院科室id = 入院科室id_In, 出院病床 = Decode(床号_In, '家庭病床', Null, 床号_In), 门诊医师 = 门诊医师_In, 编目员编号 = 操作员编号_In,
        编目员姓名 = 操作员姓名_In, 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 婚姻状况 = 婚姻状况_In, 职业 = 职业_In, 国籍 = 国籍_In, 学历 = 学历_In,
        单位电话 = 单位电话_In, 单位邮编 = 单位邮编_In, 单位地址 = 工作单位_In, 区域 = 区域_In, 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In, 家庭地址邮编 = 家庭地址邮编_In,
        户口地址 = 户口地址_In, 户口地址邮编 = 户口地址邮编_In, 联系人姓名 = 联系人姓名_In, 联系人关系 = 联系人关系_In, 联系人地址 = 联系人地址_In, 联系人身份证号 = 联系人身份证号_In,
        联系人电话 = 联系人电话_In, 医疗付款方式 = 付款方式_In, 备注 = 备注_In, 险类 = 险类_In, 状态 = Decode(床号_In, Null, 1, 0), 登记人 = 操作员姓名_In,
        登记时间 = v_Date, 再入院 = 再入院_In, 病人类型 = 病人类型_In
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0;
    Update 病人预交记录
    Set 主页id = 主页id_In
    Where 病人id = 病人id_In And 主页id Is Null And 科室id = 入院科室id_In And 预交类别 = 2 And 冲预交 Is Null And
          Trunc(收款时间) = Trunc(Sysdate);
  Else
    --入院登记或预约登记
    Insert Into 病案主页
      (病人性质, 病人id, 主页id, 住院号, 留观号, 费别, 入院病区id, 入院科室id, 入院日期, 入科时间, 入院病况, 入院方式, 入院属性, 二级院转入, 住院目的, 入院病床, 是否陪伴, 当前病况,
       当前病区id, 护理等级id, 出院科室id, 出院病床, 门诊医师, 编目员编号, 编目员姓名, 状态, 姓名, 性别, 年龄, 婚姻状况, 职业, 国籍, 学历, 单位电话, 单位邮编, 单位地址, 区域, 家庭地址,
       家庭电话, 家庭地址邮编, 户口地址, 户口地址邮编, 联系人姓名, 联系人关系, 联系人地址, 联系人电话, 联系人身份证号, 医疗付款方式, 险类, 备注, 登记人, 登记时间, 再入院, 病人类型, 挂号id)
    Values
      (病人性质_In, 病人id_In, v_主页id, Decode(病人性质_In, 1, Null, 2, Null, 住院号_In), Decode(病人性质_In, 2, 住院号_In, Null), 费别_In,
       入院病区id_In, 入院科室id_In, 入院时间_In, d_Indeptime, 入院病况_In, 入院方式_In, 入院属性_In, 二级院转入_In, 住院目的_In,
       Decode(床号_In, '家庭病床', Null, 床号_In), 是否陪伴_In, 入院病况_In, 入院病区id_In, Decode(护理等级id_In, 0, Null, 护理等级id_In), 入院科室id_In,
       Decode(床号_In, '家庭病床', Null, 床号_In), 门诊医师_In, 操作员编号_In, 操作员姓名_In, Decode(床号_In, Null, 1, 0), 姓名_In, 性别_In, 年龄_In,
       婚姻状况_In, 职业_In, 国籍_In, 学历_In, 单位电话_In, 单位邮编_In, 工作单位_In, 区域_In, 家庭地址_In, 家庭电话_In, 家庭地址邮编_In, 户口地址_In, 户口地址邮编_In,
       联系人姓名_In, 联系人关系_In, 联系人地址_In, 联系人电话_In, 联系人身份证号_In, 付款方式_In, 险类_In, 备注_In, 操作员姓名_In, v_Date, 再入院_In, 病人类型_In,
       挂号id_In);
  End If;

  Begin
    If 登记模式_In <> 1 Then
      Update 在院病人 Set 病区id = Nvl(入院病区id_In, 0), 科室id = 入院科室id_In Where 病人id = 病人id_In;
      If Sql%RowCount = 0 Then
        Insert Into 在院病人
          (病人id, 科室id, 病区id, 主页id)
        Values
          (病人id_In, 入院科室id_In, Nvl(入院病区id_In, 0), Nvl(v_主页id, 0));
      End If;
    End If;
  Exception
    When Others Then
      Null;
  End;

  Select 费别 Into v_费别 From 病人信息 Where 病人id = 病人id_In;
  If v_费别 Is Null Then
    Update 病人信息
    Set 费别 =
         (Select 费别 From 病案主页 Where 病人id = 病人id_In And 主页id = v_主页id)
    Where 病人id = 病人id_In;
  End If;

  --医保号
  If 登记模式_In <> 1 Then
    Select Zl_住院日报_Count(入院科室id_In, Trunc(入院时间_In)) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '已产生业务时间内的住院日报,不能办理该业务!';
      Raise Err_Custom;
    End If;
  
    If 医保号_In Is Not Null Then
      Insert Into 病案主页从表 (病人id, 主页id, 信息名, 信息值) Values (病人id_In, v_主页id, '医保号', 医保号_In);
    End If;
  
    --病人变动记录
    --同时入科且非家庭病床时有等级
    If 床号_In Is Not Null And 床号_In <> '家庭病床' Then
      Select 等级id Into v_等级id From 床位状况记录 Where 病区id = 入院病区id_In And 床号 = 床号_In;
    End If;
  
    --如果同时入科,则入院和入科填写到一条入院变动
    Insert Into 病人变动记录
      (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 护理等级id, 床位等级id, 床号, 病情, 操作员编号, 操作员姓名)
    Values
      (病人变动记录_Id.Nextval, 病人id_In, v_主页id, 入院时间_In, 1, 0, 入院病区id_In, 入院科室id_In, Decode(护理等级id_In, 0, Null, 护理等级id_In),
       v_等级id, Decode(床号_In, '家庭病床', Null, 床号_In), 入院病况_In, 操作员编号_In, 操作员姓名_In);
  
    Insert Into 病人自动计算
      (ID, 病人id, 主页id, 开始时间, 开始原因, 性质, 病区id, 科室id, 护理等级id, 操作员编号, 操作员姓名)
    Values
      (病人自动计算_Id.Nextval, 病人id_In, v_主页id, 入院时间_In, 1, 1, 入院病区id_In, 入院科室id_In, Decode(护理等级id_In, 0, Null, 护理等级id_In),
       操作员编号_In, 操作员姓名_In);
    Insert Into 病人自动计算
      (ID, 病人id, 主页id, 开始时间, 开始原因, 性质, 附加床位, 病区id, 科室id, 床位等级id, 床号, 操作员编号, 操作员姓名)
    Values
      (病人自动计算_Id.Nextval, 病人id_In, v_主页id, 入院时间_In, 1, 2, 0, 入院病区id_In, 入院科室id_In, v_等级id,
       Decode(床号_In, '家庭病床', Null, 床号_In), 操作员编号_In, 操作员姓名_In);
    Insert Into 病人自动计算
      (ID, 病人id, 主页id, 开始时间, 开始原因, 性质, 附加床位, 病区id, 科室id, 床位等级id, 床号, 操作员编号, 操作员姓名)
    Values
      (病人自动计算_Id.Nextval, 病人id_In, v_主页id, 入院时间_In, 1, 3, 0, 入院病区id_In, 入院科室id_In, v_等级id,
       Decode(床号_In, '家庭病床', Null, 床号_In), 操作员编号_In, 操作员姓名_In);
  
    --同时入科且非家庭病床时床位被占用
    If 床号_In Is Not Null And 床号_In <> '家庭病床' Then
      Select Count(*) Into v_Count From 床位状况记录 Where 病区id = 入院病区id_In And 床号 = 床号_In And 状态 = '空床';
    
      If v_Count = 0 Then
        v_Error := '操作失败,床位 ' || 床号_In || ' 不是空床！';
        Raise Err_Custom;
      End If;
    
      Update 床位状况记录
      Set 状态 = '占用', 病人id = 病人id_In, 科室id = Decode(共用, 1, 入院科室id_In, 科室id)
      Where 病区id = 入院病区id_In And 床号 = 床号_In;
    End If;
  
    --病人诊断记录
    If 门诊诊断_In Is Not Null Or 疾病id_In Is Not Null Then
      Insert Into 病人诊断记录
        (ID, 病人id, 主页id, 记录来源, 诊断类型, 诊断次序, 疾病id, 诊断id, 诊断描述, 记录日期, 记录人)
      Values
        (病人诊断记录_Id.Nextval, 病人id_In, v_主页id, 2, 1, 1, 疾病id_In, 诊断id_In, 门诊诊断_In, Sysdate, 操作员姓名_In);
    End If;
    If 中医诊断_In Is Not Null Or 中医疾病id_In Is Not Null Then
      Insert Into 病人诊断记录
        (ID, 病人id, 主页id, 记录来源, 诊断类型, 诊断次序, 疾病id, 诊断id, 诊断描述, 记录日期, 记录人)
      Values
        (病人诊断记录_Id.Nextval, 病人id_In, v_主页id, 2, 11, 1, 中医疾病id_In, 中医诊断id_In, 中医诊断_In, Sysdate, 操作员姓名_In);
    End If;
    --病人担保记录
    Update 病人担保记录
    Set 到期时间 = Sysdate
    Where 病人id = 病人id_In And 到期时间 Is Not Null And 到期时间 > Sysdate;
  
    --病人费用审批项目
    If 登记模式_In <> 1 Then
      Delete From 病人审批项目 Where 病人id = 病人id_In;
      b_Message.Zlhis_Patient_001(病人id_In, v_主页id);
    End If;
  
    If 登记模式_In = 0 And ((门诊诊断_In Is Not Null Or 疾病id_In Is Not Null) Or (中医诊断_In Is Not Null Or 中医疾病id_In Is Not Null)) Then
      --产生病历书写时机
      Zl_电子病历时机_Insert(病人id_In, 主页id_In, 2, '诊断', 入院科室id_In, Null, Sysdate, Sysdate);
    End If;
  
    If 登记模式_In = 0 And 床号_In Is Not Null Then
      If 再入院_In = 0 Then
        Zl_电子病历时机_Insert(病人id_In, 主页id_In, 2, '入院', 入院科室id_In, Null, 入院时间_In, 入院时间_In);
      Else
        Zl_电子病历时机_Insert(病人id_In, 主页id_In, 2, '再次入院', 入院科室id_In, Null, 入院时间_In, 入院时间_In);
      End If;
    End If;
  
    If 床号_In Is Not Null Then
      --添加首份体温单
      Zl_病人体温单_Newfirst(病人id_In, 主页id_In, 入院病区id_In);
    End If;
  
    --并发操作检查
    Select Count(*) Into v_Count From 病案主页 Where 病人id = 病人id_In And 出院日期 Is Null;
    If v_Count > 1 Then
      v_Error := '发现病人存在非法的病案记录,当前操作不能继续！' || Chr(13) || Chr(10) || '这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
      Raise Err_Custom;
    End If;
  
    Select Count(*)
    Into v_Count
    From 病人变动记录
    Where 病人id = 病人id_In And 主页id = v_主页id And Nvl(附加床位, 0) = 0 And 开始时间 Is Not Null And 终止时间 Is Null;
    If v_Count > 1 Then
      v_Error := '发现病人存在非法的变动记录,当前操作不能继续！' || Chr(13) || Chr(10) || '这可能是由于网络并发操作引起的,请刷新病人状态后再试！';
      Raise Err_Custom;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_入院病案主页_Insert;
/







------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0012' Where 编号=&n_System;
Commit;
