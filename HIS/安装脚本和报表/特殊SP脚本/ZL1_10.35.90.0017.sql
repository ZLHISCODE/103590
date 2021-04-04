----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--122504:胡俊勇,2018-06-21,新门诊系统ZLHIS相关修改
Alter Table 病案主页 Add 挂号ID number(18); 
Create Table 三方服务配置目录(
    系统标识    varchar2(100),  
    服务名称    varchar2(100),  
    服务地址  varchar2(300))    
    TABLESPACE zl9BaseItem;
Alter Table 三方服务配置目录 Add Constraint 三方服务配置目录_PK Primary Key (系统标识,服务名称) Using Index Tablespace zl9IndexHis;    

Create Index 病案主页_IX_挂号ID On 病案主页(挂号ID)  Tablespace zl9Indexcis;

------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--127144:余伟节,2018-06-21,反向问诊
Insert Into 三方服务配置目录 (系统标识, 服务名称) Values ('知识库', '反向问诊');

--122504:胡俊勇,2018-06-21,新门诊系统ZLHIS相关修改
Insert into zlTables ( 系统,表名,表空间,分类 ) Values( &n_System,'三方服务配置目录','ZL9BASEITEM','A1');

Insert Into 三方服务配置目录(系统标识,服务名称) 
Select '新门诊系统','判断医嘱是否收费' From Dual Union All
Select '新门诊系统','门诊费用转住院费用确认' From Dual Union All
Select '新门诊系统','门诊费用转住院费用' From Dual;


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--122504:胡俊勇,2018-06-21,新门诊系统ZLHIS相关修改
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1011,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
Select 'Zl_三方服务配置目录_Update','EXECUTE' From Dual Union All 
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;  




-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------

--122504:胡俊勇,2018-06-21,新门诊系统ZLHIS相关修改
CREATE OR REPLACE Procedure Zl_门诊医嘱发送_Insert
(
  医嘱id_In     In 病人医嘱发送.医嘱id%Type,
  发送号_In     In 病人医嘱发送.发送号%Type,
  记录性质_In   In 病人医嘱发送.记录性质%Type,
  No_In         In 病人医嘱发送.No%Type,
  记录序号_In   In 病人医嘱发送.记录序号%Type,
  发送数次_In   In 病人医嘱发送.发送数次%Type,
  首次时间_In   In 病人医嘱发送.首次时间%Type,
  末次时间_In   In 病人医嘱发送.末次时间%Type,
  发送时间_In   In 病人医嘱发送.发送时间%Type,
  执行状态_In   In 病人医嘱发送.执行状态%Type,
  执行部门id_In In 病人医嘱发送.执行部门id%Type,
  计费状态_In   In 病人医嘱发送.计费状态%Type,
  First_In      In Number := 0,
  样本条码_In   In 病人医嘱发送.样本条码%Type := Null,
  操作员编号_In In 人员表.编号%Type := Null,
  操作员姓名_In In 人员表.姓名%Type := Null,
  原液皮试_In   In Varchar2 := Null
  --功能：填写病人医嘱发送记录
  --参数：First_IN=表示是否一组医嘱的第一医嘱行,以便处理医嘱相关内容(如成药,配方的第一行,因为给药途径,配方煎法,用法可能为叮嘱不发送)
  --      源液皮试_In 原液皮试医嘱ID，需求号7107/bug115972用于关联药品医嘱行和皮试医嘱行。关联字段为 病人医嘱发送.标本发送批号 存入药品行的医嘱ID值
  --      格式：1医嘱ID,2医嘱ID 前面一个为皮试医嘱的医嘱ID，第二个为药品行医嘱的医嘱ID
) Is
  --包含病人及医嘱(一组医嘱中第一行)相关信息的游标
  Cursor c_Advice Is
    Select Nvl(a.相关id, a.Id) As 组id, a.相关id, a.序号, a.病人id, a.挂号单, a.婴儿, b.姓名, c.操作类型, a.诊疗类别, a.医嘱状态, a.医嘱内容, a.开嘱医生,
           a.开始执行时间, a.执行时间方案, a.频率次数, a.频率间隔, a.间隔单位, Nvl(a.紧急标志, 0) As 紧急标志, a.诊疗项目id, a.收费细目id
    From 病人医嘱记录 A, 病人信息 B, 诊疗项目目录 C
    Where a.病人id = b.病人id And a.诊疗项目id = c.Id And a.Id = 医嘱id_In
    Group By a.相关id, a.Id, a.序号, a.病人id, a.挂号单, a.婴儿, b.姓名, c.操作类型, a.诊疗类别, a.医嘱状态, a.医嘱内容, a.开嘱医生, a.开始执行时间, a.执行时间方案,
             a.频率次数, a.频率间隔, a.间隔单位, a.紧急标志, a.诊疗项目id, a.收费细目id;
  r_Advice c_Advice%RowType;

  Cursor c_Pati(v_病人id 病人信息.病人id%Type) Is
    Select * From 病人信息 Where 病人id = v_病人id;
  r_Pati c_Pati%RowType;

  --其它临时变量
  v_Temp       Varchar2(255);
  v_Count      Number;
  v_病人性质   病案主页.病人性质%Type;
  v_人员编号   人员表.编号%Type;
  v_人员姓名   人员表.姓名%Type;
  v_入院方式   入院方式.名称%Type;
  n_挂号id     病人挂号记录.Id%Type;
  d_开始时间   病人医嘱记录.开始执行时间%Type;
  n_医嘱状态   病人医嘱记录.医嘱状态%Type;
  n_皮试标号   病人医嘱发送.医嘱id%Type;
  n_皮试医嘱id 病人医嘱发送.医嘱id%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;
Begin
  --当前操作人员
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    v_人员编号 := 操作员编号_In;
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;
  --如果首次时间为空则填入开始执行时间
  Select 开始执行时间, 医嘱状态 Into d_开始时间, n_医嘱状态 From 病人医嘱记录 Where ID = 医嘱id_In;

  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;
  --是一组医嘱的第一行时处理医嘱内容
  If Nvl(First_In, 0) = 1 Or n_医嘱状态 = 1 Then
    --并发操作检查
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.医嘱状态, 0) <> 1 Then
      v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被其他人发送。' || Chr(13) || Chr(10) ||
                 '该病人的医嘱发送失败。请重新读取发送清单再试。';
      Raise Err_Custom;
    End If;
  
    --发送后的医嘱处理:临嘱发送后自动停止
    ---------------------------------------------------------------------------------------
    Update 病人医嘱记录
    Set 医嘱状态 = 8, 执行终止时间 = 末次时间_In,
        --可能没有
        停嘱时间 = 发送时间_In,
        --要作为发送时间显示
        停嘱医生 = v_人员姓名 --要作为发送人显示,不同于住院,门诊医嘱无护士操作
    Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
  
    Insert Into 病人医嘱状态
      (医嘱id, 操作类型, 操作人员, 操作时间)
      Select ID, 8, v_人员姓名, 发送时间_In From 病人医嘱记录 Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
  
    --特殊医嘱的处理
    ---------------------------------------------------------------------------------------
    If r_Advice.诊疗类别 = 'Z' And Nvl(r_Advice.操作类型, '0') <> '0' And Nvl(r_Advice.婴儿, 0) = 0 Then
      --1-留观;2-住院;
      If Instr(',1,2,', r_Advice.操作类型) > 0 And 执行部门id_In Is Not Null Then
        --满足产生新的预约登记的条件：1.当前无预约,2.当前不在院,3-无要求预约时间内的住院记录
      
        --删除超过挂号有效天数的预约登记
        Begin
          Select Count(*) Into v_Count From 病案主页 Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = 0;
        Exception
          When Others Then
            v_Count := 0;
        End;
        If Nvl(v_Count, 0) > 0 Then
          Zl_入院病案主页_Delete(r_Advice.病人id, 0, 0, 0);
          v_Count := 0;
        End If;
      
        If v_Count = 0 Then
          Select Count(*) Into v_Count From 病案主页 Where 病人id = r_Advice.病人id And 出院日期 Is Null;
        End If;
        If v_Count = 0 Then
          Select Count(*)
          Into v_Count
          From 病案主页
          Where 病人id = r_Advice.病人id And (入院日期 >= r_Advice.开始执行时间 Or 出院日期 >= r_Advice.开始执行时间);
        End If;
        If v_Count = 0 Then
          If r_Advice.操作类型 = '1' Then
            --留观医嘱,将病人在"开始时间"留观到临床执行科室
            Begin
              v_病人性质 := 2;
              Select Decode(服务对象, 1, 1, 2)
              Into v_病人性质
              From 部门性质说明
              Where 工作性质 = '临床' And 部门id = 执行部门id_In;
            Exception
              When Others Then
                Null;
            End;
          Elsif r_Advice.操作类型 = '2' Then
            --住院医嘱,将病人在"开始时间"登记到临床执行科室
            v_病人性质 := 0;
          End If;
        
          Open c_Pati(r_Advice.病人id);
          Fetch c_Pati
            Into r_Pati;
        
          v_入院方式 := Null;
          If r_Advice.紧急标志 = 1 Then
            v_入院方式 := '急诊';
            Select Max(ID)
            Into n_挂号id
            From 病人挂号记录
            Where NO = r_Advice.挂号单 And 记录性质 = 1 And 记录状态 = 1;
          Else
            Select Decode(急诊, 1, '急诊', Null), ID
            Into v_入院方式, n_挂号id
            From 病人挂号记录
            Where NO = r_Advice.挂号单 And 记录性质 = 1 And 记录状态 = 1;
          End If;
        
          If v_病人性质 = 1 Then
            Zl_入院病案主页_Insert(1, v_病人性质, r_Pati.病人id, r_Pati.门诊号, Null, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.费别,
                             r_Pati.出生日期, r_Pati.国籍, r_Pati.民族, r_Pati.学历, r_Pati.婚姻状况, r_Pati.职业, r_Pati.身份,
                             r_Pati.身份证号, r_Pati.出生地点, r_Pati.家庭地址, r_Pati.家庭地址邮编, r_Pati.家庭电话, r_Pati.户口地址,
                             r_Pati.户口地址邮编, r_Pati.联系人姓名, r_Pati.联系人关系, r_Pati.联系人地址, r_Pati.联系人电话, r_Pati.工作单位,
                             r_Pati.合同单位id, r_Pati.单位电话, r_Pati.单位邮编, r_Pati.单位开户行, r_Pati.单位帐号, r_Pati.担保人, r_Pati.担保额,
                             r_Pati.担保性质, 执行部门id_In, Null, Null, v_入院方式, Null, Null, r_Advice.开嘱医生, r_Pati.籍贯, r_Pati.区域,
                             r_Advice.开始执行时间, Null, Null, r_Pati.医疗付款方式, Null, Null, Null, Null, Null, Null, r_Pati.险类,
                             v_人员编号, v_人员姓名, 0, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, n_挂号id);
          Else
            Zl_入院病案主页_Insert(1, v_病人性质, r_Pati.病人id, r_Pati.住院号, Null, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.费别,
                             r_Pati.出生日期, r_Pati.国籍, r_Pati.民族, r_Pati.学历, r_Pati.婚姻状况, r_Pati.职业, r_Pati.身份,
                             r_Pati.身份证号, r_Pati.出生地点, r_Pati.家庭地址, r_Pati.家庭地址邮编, r_Pati.家庭电话, r_Pati.户口地址,
                             r_Pati.户口地址邮编, r_Pati.联系人姓名, r_Pati.联系人关系, r_Pati.联系人地址, r_Pati.联系人电话, r_Pati.工作单位,
                             r_Pati.合同单位id, r_Pati.单位电话, r_Pati.单位邮编, r_Pati.单位开户行, r_Pati.单位帐号, r_Pati.担保人, r_Pati.担保额,
                             r_Pati.担保性质, 执行部门id_In, Null, Null, v_入院方式, Null, Null, r_Advice.开嘱医生, r_Pati.籍贯, r_Pati.区域,
                             r_Advice.开始执行时间, Null, Null, r_Pati.医疗付款方式, Null, Null, Null, Null, Null, Null, r_Pati.险类,
                             v_人员编号, v_人员姓名, 0, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, n_挂号id);
          End If;
          Close c_Pati;
        End If;
      End If;
    End If;
  End If;
  Close c_Advice;

  If 原液皮试_In Is Not Null Then
    v_Count      := Instr(原液皮试_In, ',');
    n_皮试医嘱id := Substr(原液皮试_In, 1, v_Count - 1);
    n_皮试标号   := Substr(原液皮试_In, v_Count + 1);
    Update 病人医嘱发送 Set 标本发送批号 = n_皮试标号 Where 医嘱id = n_皮试医嘱id;
  End If;
  --填写发送记录
  ---------------------------------------------------------------------------------------
  Insert Into 病人医嘱发送
    (医嘱id, 发送号, 记录性质, NO, 记录序号, 发送数次, 发送人, 发送时间, 执行状态, 执行部门id, 计费状态, 首次时间, 末次时间, 样本条码, 门诊记帐, 标本发送批号)
  Values
    (医嘱id_In, 发送号_In, 记录性质_In, No_In, 记录序号_In, 发送数次_In, v_人员姓名, 发送时间_In, 执行状态_In, 执行部门id_In, 计费状态_In,
     Nvl(首次时间_In, d_开始时间), Nvl(末次时间_In, d_开始时间), 样本条码_In, Decode(记录性质_In, 2, 1, Null), n_皮试标号);

  --手术和检查医嘱同步更新主医嘱的计费状态
  If 计费状态_In = 1 And r_Advice.组id <> 医嘱id_In And (r_Advice.诊疗类别 = 'D' Or r_Advice.诊疗类别 = 'F') Then
    Update 病人医嘱发送 Set 计费状态 = 1 Where 医嘱id = r_Advice.组id And 发送号 = 发送号_In;
  End If;

  --自动填为已执行时，需要同步处理费用执行状态及审核划价状态
  If 执行状态_In = 1 Then
    Zl_病人医嘱执行_Finish(医嘱id_In, 发送号_In, Null, Null, v_人员编号, v_人员姓名, 执行部门id_In);
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin zl_服务窗消息_发送(:1,:2); End;'
      Using 3, 发送号_In;
  Exception
    When Others Then
      Null;
  End;

  If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '6' Then
    --检验项目
    b_Message.Zlhis_Cis_016(r_Advice.病人id, Null, r_Advice.挂号单, 发送号_In, r_Advice.组id, 1);
  Elsif r_Advice.诊疗类别 = 'D' And r_Advice.相关id Is Null Then
    b_Message.Zlhis_Cis_017(r_Advice.病人id, Null, r_Advice.挂号单, 发送号_In, r_Advice.组id, 1);
  Elsif r_Advice.诊疗类别 = 'F' And r_Advice.相关id Is Null Then
    b_Message.Zlhis_Cis_018(r_Advice.病人id, Null, r_Advice.挂号单, 发送号_In, r_Advice.组id);
  Elsif r_Advice.诊疗类别 = 'K' Then
    b_Message.Zlhis_Cis_019(r_Advice.病人id, Null, r_Advice.挂号单, 发送号_In, r_Advice.组id);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊医嘱发送_Insert;
/

--122504:胡俊勇,2018-06-21,新门诊系统ZLHIS相关修改
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
         户口地址, 户口地址邮编, 联系人姓名, 联系人关系, 联系人地址, 联系人电话, 工作单位, 合同单位id, 单位电话, 单位邮编, 单位开户行, 单位帐号, 担保人, 担保额, 担保性质, 险类, 登记时间, 其他证件, 病人类型,
         联系人身份证号, 手机号)
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
          其他证件 = 其他证件_In, 病人类型=病人类型_In, 联系人身份证号 = 联系人身份证号_In, 手机号 = Nvl(手机号_In, 手机号)
      Where 病人id = 病人id_In;
    End If;
  Else
    If 新病人_In = 1 Then
      Insert Into 病人信息
        (病人id, 住院号, 姓名, 性别, 年龄, 费别, 医疗付款方式, 出生日期, 国籍, 民族, 籍贯, 区域, 学历, 婚姻状况, 职业, 身份, 身份证号, 出生地点, 家庭地址, 家庭地址邮编, 家庭电话,
         户口地址, 户口地址邮编, 联系人姓名, 联系人关系, 联系人地址, 联系人电话, 工作单位, 合同单位id, 单位电话, 单位邮编, 单位开户行, 单位帐号, 担保人, 担保额, 担保性质, 险类, 登记时间, 其他证件, 病人类型,
         联系人身份证号, 手机号)
      Values
        (病人id_In, Decode(病人性质_In, 2, Null, 住院号_In), 姓名_In, 性别_In, 年龄_In, 费别_In, 付款方式_In, 出生日期_In, 国籍_In, 民族_In, 籍贯_In,
         区域_In, 学历_In, 婚姻状况_In, 职业_In, 身份_In, 身份证号_In, 出生地点_In, 家庭地址_In, 家庭地址邮编_In, 家庭电话_In, 户口地址_In, 户口地址邮编_In,
         联系人姓名_In, 联系人关系_In, 联系人地址_In, 联系人电话_In, 工作单位_In, Decode(合同单位id_In, 0, Null, 合同单位id_In), 单位电话_In, 单位邮编_In,
         单位开户行_In, 单位帐号_In, 担保人_In, Decode(担保额_In, 0, Null, 担保额_In), 担保性质_In, 险类_In, v_Date, 其他证件_In, 病人类型_In, 联系人身份证号_In, 手机号_In);
    Else
      --老病人的门诊费别不变,除非是门诊留观病人
      Update 病人信息
      Set 住院号 = Decode(病人性质_In, 2, 住院号, Decode(住院号_In, Null, 住院号, 住院号_In)), 姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In,
          费别 = Decode(病人性质_In, 1, 费别_In, 费别), 医疗付款方式 = 付款方式_In, 出生日期 = 出生日期_In, 国籍 = 国籍_In, 民族 = 民族_In, 籍贯 = 籍贯_In,
          区域 = 区域_In, 学历 = 学历_In, 婚姻状况 = 婚姻状况_In, 职业 = 职业_In, 身份 = 身份_In, 身份证号 = 身份证号_In, 出生地点 = 出生地点_In, 家庭地址 = 家庭地址_In,
          家庭地址邮编 = 家庭地址邮编_In, 家庭电话 = 家庭电话_In, 户口地址 = 户口地址_In, 户口地址邮编 = 户口地址邮编_In, 联系人姓名 = 联系人姓名_In, 联系人关系 = 联系人关系_In,
          联系人地址 = 联系人地址_In, 联系人电话 = 联系人电话_In, 工作单位 = 工作单位_In, 合同单位id = Decode(合同单位id_In, 0, Null, 合同单位id_In),
          单位电话 = 单位电话_In, 单位邮编 = 单位邮编_In, 单位开户行 = 单位开户行_In, 单位帐号 = 单位帐号_In, 担保人 = 担保人_In,
          担保额 = Decode(担保额_In, 0, Null, 担保额_In), 担保性质 = 担保性质_In, 险类 = 险类_In, 其他证件 = 其他证件_In, 病人类型=病人类型_In, 联系人身份证号 = 联系人身份证号_In,
          手机号 = Nvl(手机号_In, 手机号)
      Where 病人id = 病人id_In;
    End If;
  End If;

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
       家庭电话, 家庭地址邮编, 户口地址, 户口地址邮编, 联系人姓名, 联系人关系, 联系人地址, 联系人电话, 联系人身份证号, 医疗付款方式, 险类, 备注, 登记人, 登记时间, 再入院, 病人类型,挂号id)
    Values
      (病人性质_In, 病人id_In, v_主页id, Decode(病人性质_In, 1, Null, 2, Null, 住院号_In), Decode(病人性质_In, 2, 住院号_In, Null), 费别_In,
       入院病区id_In, 入院科室id_In, 入院时间_In, d_Indeptime, 入院病况_In, 入院方式_In, 入院属性_In, 二级院转入_In, 住院目的_In,
       Decode(床号_In, '家庭病床', Null, 床号_In), 是否陪伴_In, 入院病况_In, 入院病区id_In, Decode(护理等级id_In, 0, Null, 护理等级id_In), 入院科室id_In,
       Decode(床号_In, '家庭病床', Null, 床号_In), 门诊医师_In, 操作员编号_In, 操作员姓名_In, Decode(床号_In, Null, 1, 0), 姓名_In, 性别_In, 年龄_In,
       婚姻状况_In, 职业_In, 国籍_In, 学历_In, 单位电话_In, 单位邮编_In, 工作单位_In, 区域_In, 家庭地址_In, 家庭电话_In, 家庭地址邮编_In, 户口地址_In, 户口地址邮编_In,
       联系人姓名_In, 联系人关系_In, 联系人地址_In, 联系人电话_In, 联系人身份证号_In, 付款方式_In, 险类_In, 备注_In, 操作员姓名_In, v_Date, 再入院_In, 病人类型_In,挂号id_In);
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

--122504:胡俊勇,2018-06-21,新门诊系统ZLHIS相关修改
Create Or Replace Procedure Zl_三方服务配置目录_Update
(
  系统标识_In In 三方服务配置目录.系统标识%Type,
  服务名称_In In 三方服务配置目录.服务名称%Type,
  服务地址_In In 三方服务配置目录.服务地址%Type
) Is
Begin
  Update 三方服务配置目录 Set 服务地址 = 服务地址_In Where 系统标识 = 系统标识_In And 服务名称 = 服务名称_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方服务配置目录_Update;
/

--127450:李南春,2018-06-19,挂号按先进先出原则使用预交款
Create Or Replace Procedure Zl_病人挂号记录_出诊_Insert
(
  出诊记录id_In    临床出诊记录.Id%Type,
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  序号_In          门诊费用记录.序号%Type,
  价格父号_In      门诊费用记录.价格父号%Type,
  从属父号_In      门诊费用记录.从属父号%Type,
  收费类别_In      门诊费用记录.收费类别%Type,
  收费细目id_In    门诊费用记录.收费细目id%Type,
  数次_In          门诊费用记录.数次%Type,
  标准单价_In      门诊费用记录.标准单价%Type,
  收入项目id_In    门诊费用记录.收入项目id%Type,
  收据费目_In      门诊费用记录.收据费目%Type,
  结算方式_In      Varchar2,
  应收金额_In      门诊费用记录.应收金额%Type,
  实收金额_In      门诊费用记录.实收金额%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  执行部门id_In    门诊费用记录.执行部门id%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  医生姓名_In      挂号安排.医生姓名%Type,
  医生id_In        挂号安排.医生id%Type,
  病历费_In        Number, --该条记录是否病历工本费
  急诊_In          Number,
  号别_In          挂号安排.号码%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  领用id_In        票据使用明细.领用id%Type,
  预交支付_In      病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
  保险大类id_In    门诊费用记录.保险大类id%Type,
  保险项目否_In    门诊费用记录.保险项目否%Type,
  统筹金额_In      门诊费用记录.统筹金额%Type,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
  收费票据_In      Number := 0, --挂号是否使用收费票据
  保险编码_In      门诊费用记录.保险编码%Type,
  复诊_In          病人挂号记录.复诊%Type := 0,
  号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
  社区_In          病人挂号记录.社区%Type := Null,
  预约接收_In      Number := 0,
  预约方式_In      预约方式.名称%Type := Null,
  生成队列_In      Number := 0,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  合作单位_In      病人预交记录.合作单位%Type := Null,
  操作类型_In      Number := 0,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  退号重用_In      Number := 1,
  冲预交病人ids_In Varchar2 := Null,
  修正病人费别_In  Number := 0,
  预约顺序号_In    临床出诊序号控制.预约顺序号%Type := Null,
  修正病人年龄_In  Number := 0,
  收费单_In        病人挂号记录.收费单%Type := Null,
  更新交款余额_In  Number := 1 --是否更新人员交款余额，主要是处理统一操作员登录多台自助机的情况
) As
  ---------------------------------------------------------------------------
  --
  --参数:
  --     操作类型_in:0-正常挂号或者预约 1-操作员拥有加号权限加号
  --     修正病人费别_In:0-不修改病人费别 1-修改病人费别
  ----------------------------------------------------------------------------
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id, Min(Decode(记录性质, 1, 收款时间, NULL)) as 收款时间
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 收款时间;
  --功能：新增一行病人挂号费用，并将结算汇总到病人预交记录
  --       同时汇总相关的汇总表(病人挂号汇总、费用汇总)
  --       第一行费用处理票据使用情况(领用ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_排队号码 排队叫号队列.排队号码%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;

  n_分时段       Number;
  n_原始分时段   Number;
  n_时段限号     Number;
  n_时段限约     Number;
  d_时段时间     Date;
  d_最大序号时间 Date;
  n_追加号       Number := 0; --处理时段过期 追加挂号的情况
  n_已约数       病人挂号汇总.已约数%Type;
  n_预约有效时间 Number;
  n_失效数       Number;
  n_失约挂号     Number := 0;
  n_已用数量     Number;
  n_锁定         Number := 0;

  n_费用id        门诊费用记录.Id%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_当前金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  n_预交id        病人预交记录.Id%Type;
  n_挂号id        病人挂号记录.Id%Type;
  v_冲预交病人ids Varchar2(4000);

  n_组id           财务缴款分组.Id%Type;
  n_门诊号         病人信息.门诊号%Type;
  n_序号           挂号序号状态.序号%Type;
  n_已用序号       挂号序号状态.序号%Type;
  n_序号控制       挂号安排.序号控制%Type;
  n_分诊台签到排队 Number;
  n_Count          Number;
  n_限号数         Number(18);
  d_排队时间       Date;
  v_结算方式记录   Varchar2(1000);
  d_序号时间       Date;
  v_费别           费别.名称%Type;
  n_时段序号       Number := -1;
  n_预约生成队列   Number;
  v_结算方式       结算方式.名称%Type;
  v_结算内容       Varchar2(1000);
  v_当前结算       Varchar2(200);
  v_结算号码       病人预交记录.结算号码%Type;
  n_结算金额       病人预交记录.冲预交%Type;
  n_三方卡标志     Number(2);
  n_预约顺序号     临床出诊序号控制.预约顺序号%Type;
  n_限约数         Number(18);
  n_已挂数         Number(4) := 0;
  n_Exists         Number;
  n_挂出的最大序号 Number(4) := 0;
  n_分时点显示     Number(3);
  n_结算模式       病人信息.结算模式%Type;
  v_排队序号       排队叫号队列.排队序号%Type;
  v_机器名         挂号序号状态.机器名%Type;
  v_序号操作员     挂号序号状态.操作员姓名%Type;
  v_序号机器名     挂号序号状态.机器名%Type;
  v_付款方式       病人挂号记录.医疗付款方式%Type;
  n_状态           临床出诊序号控制.挂号状态%Type;
Begin
  --记录锁定判断
  If 出诊记录id_In Is Not Null Then
    Begin
      Select 1
      Into n_Exists
      From 临床出诊记录
      Where ID = 出诊记录id_In And Nvl(是否发布, 0) = 1 And Nvl(是否锁定, 0) = 0;
    Exception
      When Others Then
        v_Err_Msg := '无法确定出诊记录，请检查出诊记录是否存在或被锁定！';
        Raise Err_Item;
    End;
  End If;

  --获取当前机器名称
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);

  If 费别_In Is Null Then
    Begin
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
      Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
        Raise Err_Item;
    End;
  Else
    v_费别 := 费别_In;
    If Nvl(修正病人费别_In, 0) = 1 Then
      Begin
        Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '没有找到对应的病人！';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(修正病人年龄_In, 0) = 1 Then
    Begin
      Update 病人信息 Set 年龄 = 年龄_In Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '没有找到对应的病人！';
        Raise Err_Item;
    End;
  End If;

  If 门诊号_In Is Not Null Then
    Begin
      Select Nvl(门诊号, 0) Into n_门诊号 From 病人信息 Where 病人id = 病人id_In;
    Exception
      When Others Then
        n_门诊号 := 0;
    End;
    If n_门诊号 = 0 Then
      Update 病人信息 Set 门诊号 = 门诊号_In Where 病人id = 病人id_In;
    End If;
  End If;

  Begin
    Update 临床出诊序号控制
    Set 挂号状态 = 0
    Where 记录id = 出诊记录id_In And 序号 = 号序_In And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
  Exception
    When Others Then
      Null;
  End;

  If Nvl(预约挂号_In, 0) = 0 Then
    --挂号或者预约接收
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*)
    Into n_Count
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 In (1, 3) And 序号 = 序号_In And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    --获取结算方式名称
    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_现金 := '现金';
    End;
    Begin
      Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
    Exception
      When Others Then
        v_个人帐户 := '个人帐户';
    End;
  End If;

  n_序号 := 号序_In;

  --获取是否分时段
  Begin
    Select Nvl(是否分时段, 0), Nvl(是否序号控制, 0), 限号数, 限约数
    Into n_分时段, n_序号控制, n_限号数, n_限约数
    From 临床出诊记录
    Where ID = 出诊记录id_In;
    n_原始分时段 := n_分时段;
  Exception
    When Others Then
      n_分时段     := 0;
      n_原始分时段 := n_分时段;
      n_序号控制   := 0;
      n_限号数     := Null;
      n_限约数     := Null;
  End;

  If n_序号 Is Null And n_分时段 = 1 And n_序号控制 = 0 Then
    Begin
      Select 序号
      Into n_序号
      From 临床出诊序号控制
      Where 记录id = 出诊记录id_In And 开始时间 = 发生时间_In And Rownum < 2;
    Exception
      When Others Then
        n_序号 := Null;
    End;
  End If;

  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    Begin
      Select Max(To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 临床出诊序号控制
      Where 记录id = 出诊记录id_In And Nvl(数量, 0) <> 0;
    
      n_追加号 := Case Sign(发生时间_In - d_最大序号时间)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_追加号 := 0;
    End;
  End If;
  d_时段时间 := 发生时间_In;

  If 序号_In = 1 And n_分时段 > 0 Then
    If Nvl(n_序号控制, 0) = 1 Then
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := 发生时间_In;
          n_时段限号 := 0;
          n_时段限约 := 0;
      End;
    Else
      --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 Is Null;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := 发生时间_In;
          n_时段限号 := 0;
          n_时段限约 := 0;
      End;
    End If;
  End If;

  If 序号_In = 1 Then
    --获取当前未使用的序号
    If Nvl(预约挂号_In, 0) = 0 Then
      n_预约有效时间 := Zl_To_Number(zl_GetSysParameter('预约有效时间', 1111));
      n_失约挂号     := Zl_To_Number(zl_GetSysParameter('失约用于挂号', 1111));
    End If;
    If Nvl(n_序号控制, 0) = 1 And n_分时段 = 0 Then
      --<序号控制 未设置时段 获取可用的最大序号,以及已经使用的数量>
      Begin
        --最大序号
        Select Count(1) Into n_已用数量 From 病人挂号记录 Where 出诊记录id = 出诊记录id_In And 记录状态 = 1;
        Select Max(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
      End;
      Begin
        --最大序号
        Select Sum(Nvl(数量, 0))
        
        Into n_已约数
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And Nvl(挂号状态, 0) = 2;
      Exception
        When Others Then
          n_已约数 := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 预约时间), 1, 1, 0))
            Into n_失效数
            From 病人挂号记录
            Where 出诊记录id = 出诊记录id_In And 记录状态 = 1 And 记录性质 = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If n_原始分时段 = 0 Then
        Begin
          Select Min(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And Nvl(挂号状态, 0) = 0;
          If n_序号 Is Null Then
            n_序号 := Nvl(n_已用序号, 0);
          End If;
        Exception
          When Others Then
            Select Max(序号)
            Into n_已用序号
            From 临床出诊序号控制
            Where 记录id = 出诊记录id_In And Nvl(挂号状态, 0) <> 0;
            If n_序号 Is Null Then
              n_序号 := Nvl(n_已用序号, 0) + 1;
            End If;
        End;
      Else
        Select Max(序号) Into n_已用序号 From 临床出诊序号控制 Where 记录id = 出诊记录id_In;
        If n_序号 Is Null Then
          n_序号 := Nvl(n_已用序号, 0) + 1;
        End If;
      End If;
      --<序号控制 未设置时段 获取可用的最大序号 以及已经使用的数量 --end>
    
      --非加号的情况需要检查是否超过了限制
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号时检查
          --启用序号控制未分时段 达到了限制
          If n_限号数 <= n_已用数量 And n_限号数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd') || '已达到最大限约数！';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --启用序号控制,未分时段 加号情况   不处理,如果以后有限制条件以后补充
      End If;
    
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 0 Then
      --<--普通号分时段 这里只有预约一种情况-->
      If 操作类型_In = 0 Then
        --<正常预约挂号-->
        Begin
          Select Count(0) As 已挂数, Nvl(Sum(Decode(Nvl(Sign(a.开始时间 - d_时段时间), 0), 0, 1, 0)), 0) As 已约数
          Into n_已挂数, n_已约数
          From 临床出诊序号控制 A
          Where 记录id = 出诊记录id_In And 挂号状态 Not In (0, 4, 5);
        Exception
          When Others Then
            n_已挂数 := 0;
            n_已约数 := 0;
        End;
      
        n_时段限约 := n_时段限号; --普通号分时段的情况,29中n_时段限约始终是0 这里特殊处理
        --检查限制数量
        If n_限约数 = 0 Then
          n_限约数 := n_限号数;
        End If;
        If n_时段限约 <= n_已约数 Or n_限约数 <= n_已约数 Then
          v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限约数！';
          Raise Err_Item;
        End If;
        If n_限号数 <= n_已挂数 Then
          v_Err_Msg := '号别' || 号别_In || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
          Raise Err_Item;
        End If;
      End If;
    
      --没有达到时段的限号数 号码在当前时段往后追加
    
      --获取当天挂出的最大号序
      Select Nvl(Max(序号), 0)
      Into n_挂出的最大序号
      From 临床出诊序号控制 A
      Where 记录id = 出诊记录id_In And 预约顺序号 Is Null And 挂号状态 Not In (0, 5);
      If 预约顺序号_In Is Not Null Then
        n_预约顺序号 := 预约顺序号_In;
      Else
        Begin
          Select Nvl(Max(预约顺序号), 0) + 1
          Into n_预约顺序号
          From 临床出诊序号控制
          Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 Is Not Null;
        Exception
          When Others Then
            n_预约顺序号 := Null;
        End;
      End If;
      --设置序号
      n_序号 := RPad(Nvl(n_时段序号, 0), Length(n_限号数) + Length(Nvl(n_时段序号, 0)), 0) + n_预约顺序号;
      If n_预约顺序号 Is Null Then
        n_序号 := Nvl(n_挂出的最大序号, 0) + 1;
      End If;
    
      --<--普通号分时段--End>
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 1 Then
      --<启用序号控制 设置时段
      --专家号分时段
      Begin
        Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(开始时间 - d_时段时间), 0, 1, 0))
        Into n_已用序号, n_已挂数, n_已用数量
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 挂号状态 Not In (0, 4, 5);
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
          n_已挂数   := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 开始时间), 1, 1, 0))
            Into n_失效数
            From 临床出诊序号控制
            Where 记录id = 出诊记录id_In And 开始时间 Between Trunc(Sysdate) And Sysdate And Nvl(挂号状态, 0) = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
        
          --挂号 追加号码时不检查时段限号数
          If n_时段限号 <= n_已用数量 And Nvl(n_追加号, 0) = 0 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 - n_失效数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --挂号
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
          If n_限约数 <= n_已用数量 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_序号 Is Null Then
        --设置序号
        If Nvl(n_已用序号, 0) < Nvl(n_时段序号, 0) Then
          n_已用序号 := Nvl(n_时段序号, 0);
        End If;
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
    Elsif Nvl(n_分时段, 0) = 0 And Nvl(n_序号控制, 0) = 0 And Nvl(病历费_In, 0) = 0 And Nvl(号别_In, 0) > 0 Then
      ---<--普通号  -->
      Begin
        Select 已挂数, 已约数 Into n_已用数量, n_已约数 From 临床出诊记录 Where ID = 出诊记录id_In;
      Exception
        When Others Then
          n_已用数量 := 0;
          n_已约数   := 0;
      End;
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--普通号  -->
    End If;
  End If;

  --更新挂号序号状态
  If 序号_In = 1 And Not n_序号 Is Null Then
    If n_分时段 = 1 Then
      d_序号时间 := 发生时间_In;
    Else
      d_序号时间 := Trunc(发生时间_In);
    End If;
    --锁定序号的处理
    Begin
      If n_预约顺序号 Is Null Then
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where Nvl(挂号状态, 0) = 5 And 记录id = 出诊记录id_In And 序号 = n_序号;
      Else
        Select 操作员姓名, 工作站名称
        Into v_序号操作员, v_序号机器名
        From 临床出诊序号控制
        Where Nvl(挂号状态, 0) = 5 And 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 = n_预约顺序号;
      End If;
      n_锁定 := 1;
    Exception
      When Others Then
        v_序号操作员 := Null;
        v_序号机器名 := Null;
        n_锁定       := 0;
    End;
    If n_锁定 = 0 Then
      If n_预约顺序号 Is Null Then
        Update 临床出诊序号控制
        Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
      Else
        Update 临床出诊序号控制
        Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 = n_预约顺序号 And Nvl(挂号状态, 0) = 3 And 操作员姓名 = 操作员姓名_In;
      End If;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_分时段, 0) > 0 Then
            If Nvl(n_序号控制, 0) = 1 Then
              --分时段后专家号 失约的预约号允许挂号
              Update 临床出诊序号控制
              Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
              Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) In (0, 2);
              If Sql%NotFound Then
                Begin
                  Select 挂号状态 Into n_状态 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And 序号 = n_序号;
                Exception
                  When Others Then
                    n_状态 := -1;
                End;
                If n_状态 = -1 Then
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注)
                    Select 出诊记录id_In, n_序号, d_序号时间, d_序号时间, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1), Null,
                           Null, Null, 操作员姓名_In, '追加号'
                    From Dual;
                Else
                  v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
                  Raise Err_Item;
                End If;
              End If;
            Else
              If Nvl(预约接收_In, 0) = 1 Then
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注, 预约顺序号)
                  Select 记录id, 序号, 开始时间, 终止时间, 1, 1, Decode(预约挂号_In, 1, 2, 1), Null, Null, Null, 操作员姓名_In, n_序号, n_预约顺序号
                  From 临床出诊序号控制
                  Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 Is Null;
              End If;
            End If;
          Else
            If Nvl(n_序号控制, 0) = 1 Then
              Update 临床出诊序号控制
              Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
              Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 0;
            
              If Sql%RowCount = 0 Then
                Begin
                  Select 挂号状态 Into n_状态 From 临床出诊序号控制 Where 记录id = 出诊记录id_In And 序号 = n_序号;
                Exception
                  When Others Then
                    n_状态 := -1;
                End;
                If n_状态 = -1 Then
                  Insert Into 临床出诊序号控制
                    (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 挂号状态, 锁号时间, 类型, 名称, 操作员姓名, 备注)
                    Select 出诊记录id_In, n_序号, 发生时间_In, 发生时间_In, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1),
                           Null, Null, Null, 操作员姓名_In, '追加号'
                    From Dual;
                Else
                  v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
                  Raise Err_Item;
                End If;
              End If;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
            Raise Err_Item;
        End;
      End If;
    Else
      If 操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
        v_Err_Msg := '序号' || n_序号 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
        Raise Err_Item;
      Else
        If n_预约顺序号 Is Null Then
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_序号 And Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In And 工作站名称 = v_机器名;
        Else
          Update 临床出诊序号控制
          Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1)
          Where 记录id = 出诊记录id_In And 序号 = n_时段序号 And 预约顺序号 = n_预约顺序号 And Nvl(挂号状态, 0) = 5 And 操作员姓名 = 操作员姓名_In And
                工作站名称 = v_机器名;
        End If;
      End If;
    End If;
  End If;

  --产生病人挂号费用(可能单独是或包括病历费用)
  Select 病人费用记录_Id.Nextval Into n_费用id From Dual; --应该通过程序得到

  Insert Into 门诊费用记录
    (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id, 收费类别,
     计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名,
     发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
  Values
    (n_费用id, 4, Decode(预约挂号_In, 1, 0, 1), 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, 票据号_In, 1, 急诊_In,
     病历费_In, Decode(预约挂号_In, 1, To_Char(n_序号), 诊室_In), Decode(病人id_In, 0, Null, 病人id_In),
     Decode(门诊号_In, 0, Null, 门诊号_In), 付款方式_In, 姓名_In, Decode(姓名_In, Null, Null, 性别_In), Decode(姓名_In, Null, Null, 年龄_In),
     v_费别, 病人科室id_In, 收费类别_In, 号别_In, 收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In,
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额_In)),
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In)), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 开单部门id_In,
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, 发生时间_In, 登记时间_In, 保险大类id_In,
     保险项目否_In, 保险编码_In, 统筹金额_In, Decode(收费单_In, Null, 摘要_In, '划价:' || 收费单_In), 预约方式_In, Decode(预约挂号_In, 1, Null, n_组id));

  --汇总结算到病人预交记录
  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 0 Then
    If Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0 And 序号_In = 1 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_现金, 0, 登记时间_In, 操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费',
         n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4);
    End If;
    If Nvl(现金支付_In, 0) <> 0 And 序号_In = 1 Then
      v_结算内容     := 结算方式_In || '|'; --以空格分开以|结尾,没有结算号码的
      v_结算方式记录 := '';
      While v_结算内容 Is Not Null Loop
        v_当前结算 := Substr(v_结算内容, 1, Instr(v_结算内容, '|') - 1);
        v_结算方式 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_结算金额 := To_Number(Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1));
      
        v_当前结算 := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        v_结算号码 := Substr(v_当前结算, 1, Instr(v_当前结算, ',') - 1);
      
        v_当前结算   := Substr(v_当前结算, Instr(v_当前结算, ',') + 1);
        n_三方卡标志 := To_Number(v_当前结算);
      
        If Instr('|' || v_结算方式记录 || '|', '|' || Nvl(v_结算方式, v_现金) || '|') <> 0 Then
          v_Err_Msg := '使用了重复的结算方式,请检查!';
          Raise Err_Item;
        Else
          v_结算方式记录 := v_结算方式记录 || '|' || Nvl(v_结算方式, v_现金);
        End If;
      
        If n_三方卡标志 = 0 Then
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, Null, Null, Null, Null, Null, 合作单位_In, 4, v_结算号码);
        Else
          Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
          Insert Into 病人预交记录
            (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明,
             合作单位, 结算性质, 结算号码)
          Values
            (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(v_结算方式, v_现金), Nvl(n_结算金额, 0), 登记时间_In,
             操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4,
             v_结算号码);
        
          If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
            Zl_病人卡结算记录_支付(结算卡序号_In, 卡号_In, 0, Nvl(n_结算金额, 0), n_预交id, 操作员编号_In, 操作员姓名_In, 登记时间_In);
          End If;
        End If;
      
        If Nvl(更新交款余额_In, 1) = 1 Then
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + n_结算金额
          Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金)
          Returning 余额 Into n_返回值;
        
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (操作员姓名_In, Nvl(v_结算方式, v_现金), 1, n_结算金额);
            n_返回值 := n_结算金额;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(v_结算方式, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End If;
      
        v_结算内容 := Substr(v_结算内容, Instr(v_结算内容, '|') + 1);
      End Loop;
    End If;
  
    --对于医保挂号
    If Nvl(个帐支付_In, 0) <> 0 And 序号_In = 1 Then
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_个人帐户, 个帐支付_In, 登记时间_In, 操作员编号_In,
         操作员姓名_In, 结帐id_In, '医保挂号', n_组id, 4);
    End If;
  
    --对于就诊卡通过预交金挂号
    If Nvl(预交支付_In, 0) <> 0 And 序号_In = 1 Then
      n_预交金额 := 预交支付_In;
      For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
        n_当前金额 := Case
                    When r_Deposit.金额 - n_预交金额 < 0 Then
                     r_Deposit.金额
                    Else
                     n_预交金额
                  End;
      
        If r_Deposit.结帐id = 0 Then
          --第一次冲预交(填上结帐ID,金额为0)
          Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
        
        End If;
        --冲上次剩余额
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
           冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 登记时间_In, 操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2);
        --检查是否已经处理完
        If r_Deposit.金额 < n_预交金额 Then
          n_预交金额 := n_预交金额 - r_Deposit.金额;
        Else
          n_预交金额 := 0;
        End If;
      
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
      If n_预交金额 > 0 Then
        v_Err_Msg := '预交余不够支付本次支付金额,不能继续操作！';
        Raise Err_Item;
      
      End If;
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  
    --相关汇总表的处理
    --人员缴款余额
    If 序号_In = 1 And Nvl(更新交款余额_In, 1) = 1 Then
      If Nvl(个帐支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 个帐支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
          n_返回值 := 个帐支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_个人帐户 And Nvl(余额, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --病人挂号汇总(只处理一次,且单独收取病历费不处理)
  If Nvl(预约挂号_In, 0) = 0 Then
    --病人本次就诊(以发生时间为准)
    If Nvl(病人id_In, 0) <> 0 And 序号_In = 1 Then
      Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
  
    --病人余额
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 1, Nvl(实收金额_In, 0), 0);
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 1, Nvl(实收金额_In, 0));
    End If;
  End If;

  --病人挂号记录
  If 号别_In Is Not Null And 序号_In = 1 Then
    --And Nvl(预约挂号_In, 0) = 0
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    Begin
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
    Exception
      When Others Then
        v_付款方式 := Null;
    End;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式, 出诊记录id, 收费单)
    Values
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In,
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 出诊记录id_In, 收费单_In);
  
    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
      Update 病人挂号记录
      Set 预约 = 1, 预约时间 = 发生时间_In, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In
      Where ID = n_挂号id;
    End If;
  
    n_预约生成队列 := 0;
    If Nvl(预约挂号_In, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;
  
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then
        n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(预约挂号_In, 0) = 1 And n_分时点显示 = 1 And n_分时段 = 1 Then
          n_分时点显示 := 1;
        Else
          n_分时点显示 := Null;
        End If;
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_队列名称 := 执行部门id_In;
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号);
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In,
                         n_分时点显示, v_排队序号);
      
        --挂号立即排队
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          Update 病人挂号记录 Set 记录标志 = 1 Where ID = n_挂号id;
        End If;
      End If;
    End If;
  End If;
  --病人担保信息
  If 病人id_In Is Not Null And 序号_In = 1 Then
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(结算模式, 0) Into n_结算模式 From 病人信息 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的病人信息,不允许挂号';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Exists (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
  
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = Sysdate
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) > Sysdate;
    End If;
  End If;
  If 序号_In = 1 Then
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 1, n_挂号id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_挂号id, 单据号_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_出诊_Insert;
/

--127450:李南春,2018-06-20,余额退款时增加退款记录的冲预交信息，避免被退预交再次使用
CREATE OR REPLACE Procedure Zl_病人预交记录_Insert
(
  Id_In           病人预交记录.Id%Type,
  单据号_In       病人预交记录.No%Type,
  票据号_In       票据使用明细.号码%Type,
  病人id_In       病人预交记录.病人id%Type,
  主页id_In       病人预交记录.主页id%Type,
  科室id_In       病人预交记录.科室id%Type,
  金额_In         病人预交记录.金额%Type,
  结算方式_In     病人预交记录.结算方式%Type,
  结算号码_In     病人预交记录.结算号码%Type,
  缴款单位_In     病人预交记录.缴款单位%Type,
  单位开户行_In   病人预交记录.单位开户行%Type,
  单位帐号_In     病人预交记录.单位帐号%Type,
  摘要_In         病人预交记录.摘要%Type,
  操作员编号_In   病人预交记录.操作员编号%Type,
  操作员姓名_In   病人预交记录.操作员姓名%Type,
  领用id_In       票据使用明细.领用id%Type,
  预交类别_In     病人预交记录.预交类别%Type := Null,
  卡类别id_In     病人预交记录.卡类别id%Type := Null,
  结算卡序号_In   病人预交记录.结算卡序号%Type := Null,
  卡号_In         病人预交记录.卡号%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  合作单位_In     病人预交记录.合作单位%Type := Null,
  收款时间_In     病人预交记录.收款时间%Type := Null,
  操作类型_In     Integer := 0,
  结帐id_In       病人预交记录.结帐id%Type := Null,
  结算性质_In     病人预交记录.结算性质%Type := Null,
  退款检查_In     Number := 0,
  强制退现_In     Number := 0,
  更新交款余额_In Number := 1,
  是否转账_In     Number := 0
) As
  ----------------------------------------------
  --操作类型_In:0-正常缴预交;1-存为划价单;3-余额退款
  --结帐ID_IN:>0时,表示某次结帐时,同步产生的预交记录
  --退款检查_In;0-忽略退款金额是否大于了病人余额；1-检查退款金额
  --更新交款余额_In:0-在 zl_人员缴款余额_Update 中更新；1-在本过程中更新
  --强制退现_In:0-不强制，1-三方卡或消费卡不允许退现但强制退现金给病人
  --是否转账_In:0-原样退或退现，1-转账到支持的三方卡上

  v_Err_Msg         Varchar2(200);
  Err_Item          Exception;

  v_性质            结算方式.性质%Type;
  v_打印id          票据打印内容.Id%Type;
  v_担保            病人信息.担保性质%Type;
  v_Date            Date;
  n_返回值          病人余额.预交余额%Type;
  n_组id            财务缴款分组.Id%Type;
  n_病人余额        病人余额.预交余额%Type;
  n_三方预交        病人余额.预交余额%Type;
  n_退款金额        病人预交记录.金额%Type;
  n_剩余款          病人预交记录.金额%Type;
  n_结帐id          病人结帐记录.ID%Type;
  
  Cursor C_冲预交 is
    Select A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明, 0 as 序号, A.收款时间, A.金额 AS 预交金
    From 病人预交记录 A Where RowNum < 2;
  r_冲预交 C_冲预交%Rowtype;
  
  Type Ty_剩余款 Is Ref Cursor;
  C_剩余款 Ty_剩余款; --动态游标变量 
Begin
  v_Date := 收款时间_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  n_组id := Zl_Get组id(操作员姓名_In);

  --插入预交缴款记录
  Insert Into 病人预交记录
    (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id, 预交类别,
     卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质)
  Values
    (Id_In, 单据号_In, 票据号_In, 1, Decode(操作类型_In, 1, 0, 1), 病人id_In, Decode(主页id_In, 0, Null, 主页id_In),
     Decode(科室id_In, 0, Null, 科室id_In), 金额_In, 结算方式_In, 结算号码_In, v_Date, 缴款单位_In, 单位开户行_In, 单位帐号_In, 操作员编号_In, 操作员姓名_In,
     摘要_In, n_组id, 预交类别_In, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 结帐id_In,
     Decode(结帐id_In, Null, Null, 0), 结算性质_In);
     
  If 操作类型_In = 1 Then
    --暂不处理汇总表
    Return;
  Elsif 操作类型_In = 3 Then
    --生成一条原预交ID的冲销记录，同时也生成一条余额退款的冲销记录
    Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    IF Nvl(卡类别id_In, 0) = 0 And Nvl(结算卡序号_In, 0) =0 then
      --退现，包括普通结算方式退现、强制退现、三方卡允许退现
      Open C_剩余款 For
           Select A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明, 
                   Min(decode(sign(A.金额),-1,0,1)) AS 序号, Min(decode(A.记录性质,1,A.收款时间,null)) AS 收款时间,  
                   Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) as 预交金
              From 病人预交记录 A, 医疗卡类别 B, 消费卡类别目录 C
             Where A.病人ID = 病人id_In And A.记录性质 In (1,11) And A.预交类别 = Nvl(预交类别_In, 2)
               And A.卡类别ID = B.ID(+) And Decode(强制退现_In, 1, 1, Nvl(B.是否退现, 1)) = 1
               And A.卡类别ID = C.编号(+) And Decode(强制退现_In, 1, 1, Nvl(C.是否退现, 1)) = 1
             Group By A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明
            Having Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) <> 0
             Order By 序号,收款时间;
    ElsIF Nvl(是否转账_In, 0) = 1 Then
      --转账，三方卡允许退现或者强制退现，传入的卡号可能不是原卡号,金额由同种卡类别的预交缴款分摊
      --目前只支持同一种卡转账
      Open C_剩余款 For
           Select A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明, 
                   Min(decode(sign(A.金额),-1,0,1)) AS 序号, Min(decode(A.记录性质,1,A.收款时间,null)) AS 收款时间,  
                   Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) as 预交金
              From 病人预交记录 A, 医疗卡类别 B
             Where A.病人ID = 病人id_In And A.记录性质 In (1,11) And A.预交类别 = Nvl(预交类别_In, 2)
               And A.卡类别ID = B.ID(+)
               And Nvl(卡类别id, 0) = Nvl(卡类别id_In, 0) And Nvl(交易流水号, '-') = Nvl(交易流水号_In, '-')
             Group By A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明
            Having Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) <> 0
             Order By 序号,收款时间;
    Else
      --退三方卡或者是消费卡，根据卡类别ID、结算卡序号、卡号、交易流水号缺省原预交记录，如果不能确定唯一则进行分摊
      Open C_剩余款 For
           Select A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明, 
                   Min(decode(sign(A.金额),-1,0,1)) AS 序号, Min(decode(A.记录性质,1,A.收款时间,null)) AS 收款时间,  
                   Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) as 预交金
              From 病人预交记录 A
             Where A.病人ID = 病人id_In And A.记录性质 In (1,11) And A.预交类别 = Nvl(预交类别_In, 2)
               And Nvl(A.卡类别id, 0) = Nvl(卡类别id_In, 0) And Nvl(A.结算卡序号, 0) = Nvl(结算卡序号_In, 0) 
               And Nvl(A.卡号, '-') = Nvl(卡号_In, '-') And Nvl(交易流水号, '-') = Nvl(交易流水号_In, '-')
             Group By A.NO, A.病人id, A.预交类别, A.卡类别id, A.卡号, A.交易流水号, A.交易说明
            Having Nvl(Sum(A.金额), 0) - Nvl(Sum(A.冲预交), 0) <> 0
             Order By 序号,收款时间;
    End IF;
    
    n_剩余款 := -1 * 金额_In;
    n_退款金额 := 0;
    Loop
      Fetch C_剩余款
        Into r_冲预交;
      Exit When C_剩余款%NotFound;
      IF r_冲预交.NO <> 单据号_In Then
        IF n_剩余款 > r_冲预交.预交金 then
           n_退款金额 := r_冲预交.预交金;
           n_剩余款 := n_剩余款 - n_退款金额;
        Else
           n_退款金额 := n_剩余款;
           n_剩余款 := 0;
        End IF;
          	  
        IF nvl(n_退款金额, 0) <> 0 THEN 
          UPDATE 病人预交记录  SET 结帐ID = n_结帐id WHERE NO = r_冲预交.NO AND 记录性质 = 1 AND 结帐ID IS NULL;
          Insert Into 病人预交记录
             (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 操作员编号,
             收款时间, 操作员姓名, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 1, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 缴款单位, 单位开户行, 单位帐号, 操作员编号_In,
             v_Date, 操作员姓名_In, 摘要, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, n_结帐id, n_退款金额, NULL
          From 病人预交记录
          Where NO = r_冲预交.NO And 记录性质 In (1, 11) And RowNum < 2;
        END IF;

        IF n_剩余款 = 0 Then 
          Exit;
        End IF;
      End IF;
    END LOOP;

    IF n_剩余款 <> 0 And Nvl(退款检查_In, 0) = 1 THEN 
      v_Err_Msg := '退款金额大于病人剩余预交余额。';
      Raise Err_Item;
    END IF;
    
    n_退款金额 := -1 * (-1 * 金额_In - n_剩余款);
    IF n_退款金额 <> 0 Then
      Update 病人预交记录 Set 结帐id = n_结帐id Where ID = Id_In;
      Insert Into 病人预交记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行, 单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id, 预交类别,
         卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 单据号_In, 票据号_In, 11, 1, 病人id_In, Decode(主页id_In, 0, Null, 主页id_In),
         Decode(科室id_In, 0, Null, 科室id_In), NULL, 结算方式_In, 结算号码_In, v_Date, 缴款单位_In, 单位开户行_In, 单位帐号_In, 操作员编号_In, 操作员姓名_In,
         摘要_In, n_组id, 预交类别_In, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, n_结帐id, n_退款金额, NULL);
    End IF;
  End If;

  --处理票据
  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into v_打印id From Dual;

    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (v_打印id, 2, 单据号_In);

    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
    Values
      (票据使用明细_Id.Nextval, 2, 票据号_In, 1, 1, 领用id_In, v_打印id, v_Date, 操作员姓名_In, 金额_In);

    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  End If;

  --相关汇总表处理

  --病人余额(预交余额现收)
  Begin
    Select 性质 Into v_性质 From 结算方式 Where 名称 = 结算方式_In;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(v_性质, 1) <> 5 Then
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + 金额_In
    Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 0) = Nvl(预交类别_In, 0)
    Returning 预交余额 Into n_返回值;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额
        (病人id, 性质, 类型, 预交余额, 费用余额)
      Values
        (病人id_In, 1, Nvl(预交类别_In, 0), 金额_In, 0);
      n_返回值 := 金额_In;
    End If;
    If Nvl(金额_In, 0) = 0 Then
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
    End If;
  End If;

  If 金额_In < 0 Then
    Begin
      Select Nvl(预交余额, 0) - Nvl(费用余额, 0)
      Into n_病人余额
      From 病人余额
      Where 性质 = 1 And 病人id = 病人id_In And Nvl(类型, 0) = Nvl(预交类别_In, 0);
    Exception
      When Others Then
        Null;
    End;
    --余额退款要考虑三方预交是否支持退现
    If 操作类型_In = 3 And Nvl(强制退现_In, 0) = 0 Then
      For c_三方预交 In (Select a.预交id, a.预交类别, a.卡类别id, a.结算卡序号 As 消费接口id, Nvl(b.编码, c.编号) As 编码, Nvl(b.名称, c.名称) As 名称,
                            Decode(b.编码, Null, c.是否全退, b.是否全退) As 是否全退, Decode(b.编码, Null, c.是否退现, b.是否退现) As 是否退现, a.卡号,
                            a.交易流水号, a.交易说明, a.预交余额
                     From (Select a.预交类别, Nvl(a.卡类别id, 0) As 卡类别id, Nvl(a.结算卡序号, 0) As 结算卡序号, a.卡号, a.交易流水号, a.交易说明,
                                   Max(Decode(Sign(金额), -1, Decode(a.记录状态, 1, 0, 2, 0, ID), ID)) As 预交id,
                                   Nvl(Sum(金额), 0) - Nvl(Sum(Nvl(冲预交, 0)), 0) As 预交余额
                            From 病人预交记录 A
                            Where a.病人id = 病人id_In And (Nvl(a.结算卡序号, 0) <> 0 Or Nvl(卡类别id, 0) <> 0)
                            Group By a.预交类别, Nvl(a.卡类别id, 0), Nvl(a.结算卡序号, 0), a.卡号, a.交易流水号, a.交易说明
                            Having Nvl(Sum(金额), 0) - Nvl(Sum(Nvl(冲预交, 0)), 0) <> 0) A, 医疗卡类别 B, 消费卡类别目录 C
                     Where a.预交类别 = Nvl(预交类别_In, 0) And a.卡类别id = b.Id(+) And a.结算卡序号 = c.编号(+) And Nvl(a.预交余额, 0) <> 0
                     Order By 编码, a.卡号, a.交易流水号, a.交易说明) Loop

        If Instr(',7,8,', ',' || v_性质 || ',') = 0 And Nvl(c_三方预交.是否退现, 0) = 0 And Nvl(c_三方预交.预交余额, 0) > 0 Then
          n_三方预交 := Nvl(n_三方预交, 0) + Nvl(c_三方预交.预交余额, 0);
        Elsif Instr(',7,8,', ',' || v_性质 || ',') > 0 Then
          If Nvl(c_三方预交.卡号, '0') = Nvl(卡号_In, '0') And Nvl(c_三方预交.交易流水号, '0') = Nvl(交易流水号_In, '0') And
             Nvl(c_三方预交.交易说明, '0') = Nvl(交易说明_In, '0') Then
            n_三方预交 := Nvl(n_三方预交, 0) + Nvl(c_三方预交.预交余额, 0);
          End If;
        End If;
      End Loop;
    End If;

    If Instr(',7,8,', ',' || v_性质 || ',') > 0 And Nvl(n_三方预交, 0) < 0 And 操作类型_In = 3 Then
      v_Err_Msg := '退款金额大于病人三方预交金额。';
      Raise Err_Item;
    Elsif Nvl(n_病人余额, 0) < 0 And 退款检查_In = 1 Then
      v_Err_Msg := '退款金额大于病人剩余预交余额。';
      Raise Err_Item;
    Elsif Instr(',7,8,', ',' || v_性质 || ',') = 0 And Nvl(n_病人余额, 0) - Nvl(n_三方预交, 0) < 0 And 操作类型_In = 3 And
          退款检查_In = 1 Then
      v_Err_Msg := '退款金额大于病人剩余预交余额。';
      Raise Err_Item;
    End If;
  End If;

  --人员缴款余额(现收)
  If Nvl(更新交款余额_In, 1) = 1 Then
    Update 人员缴款余额
    Set 余额 = Nvl(余额, 0) + 金额_In
    Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = 结算方式_In
    Returning 余额 Into n_返回值;

    If Sql%RowCount = 0 Then
      Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, 结算方式_In, 1, 金额_In);
      n_返回值 := 金额_In;
    End If;
    If Nvl(n_返回值, 0) = 0 Then
      Delete From 人员缴款余额
      Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = 结算方式_In And Nvl(余额, 0) = 0;
    End If;
  End If;
  --对临时担保的处理
  Select Nvl(担保性质, 0) Into v_担保 From 病人信息 Where 病人id = 病人id_In;
  If v_担保 = 1 And Nvl(金额_In, 0) > 0 Then
    Update 病人信息
    Set 担保额 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, Nvl(担保额, 0) - Nvl(金额_In, 0), Null),
        担保人 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, 担保人, Null),
        担保性质 = Decode(Sign(Nvl(担保额, 0) - Nvl(金额_In, 0)), 1, 担保性质, Null)
    Where 病人id = 病人id_In;
  End If;
  If 操作类型_In <> 1 And 结帐id_In Is Null Then
    If 金额_In < 0 Then
      b_Message.Zlhis_Charge_006(Id_In, 单据号_In);
    Else
      b_Message.Zlhis_Charge_005(Id_In, 单据号_In);
    End If;
    --消息推送;
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 11, Id_In;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_Insert;
/

--127450:李南春,2018-06-20,挂号按先进先出原则使用预交款
Create Or Replace Procedure Zl_病人挂号记录_Insert
(
  病人id_In        门诊费用记录.病人id%Type,
  门诊号_In        门诊费用记录.标识号%Type,
  姓名_In          门诊费用记录.姓名%Type,
  性别_In          门诊费用记录.性别%Type,
  年龄_In          门诊费用记录.年龄%Type,
  付款方式_In      门诊费用记录.付款方式%Type, --用于存放病人的医疗付款方式编号
  费别_In          门诊费用记录.费别%Type,
  单据号_In        门诊费用记录.No%Type,
  票据号_In        门诊费用记录.实际票号%Type,
  序号_In          门诊费用记录.序号%Type,
  价格父号_In      门诊费用记录.价格父号%Type,
  从属父号_In      门诊费用记录.从属父号%Type,
  收费类别_In      门诊费用记录.收费类别%Type,
  收费细目id_In    门诊费用记录.收费细目id%Type,
  数次_In          门诊费用记录.数次%Type,
  标准单价_In      门诊费用记录.标准单价%Type,
  收入项目id_In    门诊费用记录.收入项目id%Type,
  收据费目_In      门诊费用记录.收据费目%Type,
  结算方式_In      病人预交记录.结算方式%Type, --现金的结算名称
  应收金额_In      门诊费用记录.应收金额%Type,
  实收金额_In      门诊费用记录.实收金额%Type,
  病人科室id_In    门诊费用记录.病人科室id%Type,
  开单部门id_In    门诊费用记录.开单部门id%Type,
  执行部门id_In    门诊费用记录.执行部门id%Type,
  操作员编号_In    门诊费用记录.操作员编号%Type,
  操作员姓名_In    门诊费用记录.操作员姓名%Type,
  发生时间_In      门诊费用记录.发生时间%Type,
  登记时间_In      门诊费用记录.登记时间%Type,
  医生姓名_In      挂号安排.医生姓名%Type,
  医生id_In        挂号安排.医生id%Type,
  病历费_In        Number, --该条记录是否病历工本费
  急诊_In          Number,
  号别_In          挂号安排.号码%Type,
  诊室_In          门诊费用记录.发药窗口%Type,
  结帐id_In        门诊费用记录.结帐id%Type,
  领用id_In        票据使用明细.领用id%Type,
  预交支付_In      病人预交记录.冲预交%Type, --刷卡挂号时使用的预交金额,序号为1传入.
  现金支付_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
  个帐支付_In      病人预交记录.冲预交%Type, --挂号时个人帐户支付金额,,序号为1传入.
  保险大类id_In    门诊费用记录.保险大类id%Type,
  保险项目否_In    门诊费用记录.保险项目否%Type,
  统筹金额_In      门诊费用记录.统筹金额%Type,
  摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
  预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
  收费票据_In      Number := 0, --挂号是否使用收费票据
  保险编码_In      门诊费用记录.保险编码%Type,
  复诊_In          病人挂号记录.复诊%Type := 0,
  号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
  社区_In          病人挂号记录.社区%Type := Null,
  预约接收_In      Number := 0,
  预约方式_In      预约方式.名称%Type := Null,
  生成队列_In      Number := 0,
  卡类别id_In      病人预交记录.卡类别id%Type := Null,
  结算卡序号_In    病人预交记录.结算卡序号%Type := Null,
  卡号_In          病人预交记录.卡号%Type := Null,
  交易流水号_In    病人预交记录.交易流水号%Type := Null,
  交易说明_In      病人预交记录.交易说明%Type := Null,
  合作单位_In      病人预交记录.合作单位%Type := Null,
  操作类型_In      Number := 0,
  险类_In          病人挂号记录.险类%Type := Null,
  结算模式_In      Number := 0,
  记帐费用_In      Number := 0,
  退号重用_In      Number := 1,
  冲预交病人ids_In Varchar2 := Null,
  修正病人费别_In  Number := 0,
  修正病人年龄_In  Number := 0,
  收费单_In        病人挂号记录.收费单%Type := Null,
  更新交款余额_In  Number := 1
) As
  ---------------------------------------------------------------------------
  --
  --参数:
  --     操作类型_in:0-正常挂号或者预约 1-操作员拥有加号权限加号
  --     修正病人费别_In:0-不修改病人费别 1-修改病人费别
  --     更新交款余额_In:0-在zl_人员缴款余额_Update 中更新 1-在本过程中更新
  ----------------------------------------------------------------------------
  --该游标用于收费冲预交的可用预交列表
  --以ID排序，优先冲上次未冲完的。
  Cursor c_Deposit
  (
    v_病人id        病人信息.病人id%Type,
    v_冲预交病人ids Varchar2
  ) Is
    Select 病人id, NO, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态, Nvl(Max(结帐id), 0) As 结帐id,
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id, Min(Decode(记录性质, 1, 收款时间, NULL)) as 收款时间
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 收款时间;
  --功能：新增一行病人挂号费用，并将结算汇总到病人预交记录
  --       同时汇总相关的汇总表(病人挂号汇总、费用汇总)
  --       第一行费用处理票据使用情况(领用ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_排队号码 排队叫号队列.排队号码%Type;
  v_现金     结算方式.名称%Type;
  v_个人帐户 结算方式.名称%Type;
  v_队列名称 排队叫号队列.队列名称%Type;

  n_分时段       Number;
  n_时段限号     Number;
  n_时段限约     Number;
  d_时段时间     Date;
  d_最大序号时间 Date;
  n_追加号       Number := 0; --处理时段过期 追加挂号的情况
  n_已约数       病人挂号汇总.已约数%Type;
  n_预约有效时间 Number;
  n_失效数       Number;
  n_失约挂号     Number := 0;
  n_已用数量     Number;
  n_锁定         Number := 0;

  n_费用id        门诊费用记录.Id%Type;
  n_预交金额      病人预交记录.金额%Type;
  n_当前金额      病人预交记录.金额%Type;
  n_返回值        病人预交记录.金额%Type;
  n_预交id        病人预交记录.Id%Type;
  n_挂号id        病人挂号记录.Id%Type;
  v_冲预交病人ids Varchar2(4000);

  n_组id           财务缴款分组.Id%Type;
  n_门诊号         病人信息.门诊号%Type;
  n_序号           挂号序号状态.序号%Type;
  n_已用序号       挂号序号状态.序号%Type;
  n_序号控制       挂号安排.序号控制%Type;
  n_分诊台签到排队 Number;
  n_Count          Number;
  n_限号数         Number(18);
  d_排队时间       Date;
  d_序号时间       Date;
  v_费别           费别.名称%Type;
  n_时段序号       Number := -1;
  n_预约生成队列   Number;
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type := 0;
  v_星期           挂号安排限制.限制项目%Type;
  n_限约数         Number(18);
  n_已挂数         Number(4) := 0;

  n_挂出的最大序号 Number(4) := 0;
  n_结算模式       病人信息.结算模式%Type;
  v_排队序号       排队叫号队列.排队序号%Type;
  v_机器名         挂号序号状态.机器名%Type;
  v_序号操作员     挂号序号状态.操作员姓名%Type;
  v_序号机器名     挂号序号状态.机器名%Type;
  v_付款方式       病人挂号记录.医疗付款方式%Type;
  v_Temp           Varchar2(3000);
  v_时间段         时间段.时间段%Type;
  d_检查开始时间   时间段.开始时间%Type;
  d_检查结束时间   时间段.终止时间%Type;
  n_出诊记录id     临床出诊记录.Id%Type;
  n_分时点显示     Number(3);
  d_启用时间       Date;
Begin
  --获取当前机器名称
  Select Terminal Into v_机器名 From V$session Where Audsid = Userenv('sessionid');
  n_组id          := Zl_Get组id(操作员姓名_In);
  v_冲预交病人ids := Nvl(冲预交病人ids_In, 病人id_In);

  If 费别_In Is Null Then
    Begin
      Select 名称 Into v_费别 From 费别 Where 缺省标志 = 1 And Rownum < 2;
      Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '无法确定病人费别，请检查缺省费别是否正确设置！';
        Raise Err_Item;
    End;
  Else
    v_费别 := 费别_In;
    If Nvl(修正病人费别_In, 0) = 1 Then
      Begin
        Update 病人信息 Set 费别 = v_费别 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '没有找到对应的病人！';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(修正病人年龄_In, 0) = 1 Then
    Begin
      Update 病人信息 Set 年龄 = 年龄_In Where 病人id = 病人id_In;
    Exception
      When Others Then
        v_Err_Msg := '没有找到对应的病人！';
        Raise Err_Item;
    End;
  End If;

  If 门诊号_In Is Not Null Then
    Begin
      Select Nvl(门诊号, 0) Into n_门诊号 From 病人信息 Where 病人id = 病人id_In;
    Exception
      When Others Then
        n_门诊号 := 0;
    End;
    If n_门诊号 = 0 Then
      Update 病人信息 Set 门诊号 = 门诊号_In Where 病人id = 病人id_In;
    End If;
  End If;

  Begin
    Delete From 挂号序号状态
    Where 号码 = 号别_In And 日期 = 发生时间_In And 序号 = 号序_In And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
  Exception
    When Others Then
      Null;
  End;
  v_Temp := zl_GetSysParameter(256);
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
    Null;
  Else
    Begin
      d_启用时间 := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
    If d_启用时间 Is Not Null Then
      If 发生时间_In > d_启用时间 Then
        v_Err_Msg := '当前挂号的发生时间' || To_Char(发生时间_In, 'yyyy-mm-dd hh24:mi:ss') || '已经启用了出诊表排班模式,不能再使用计划排班模式挂号!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 Then
    --挂号或者预约接收
    --因为现在有按日编单据号规则,日挂号量不能超过10000次,所以要检查唯一约束。
    Select Count(*)
    Into n_Count
    From 门诊费用记录
    Where 记录性质 = 4 And 记录状态 In (1, 3) And 序号 = 序号_In And NO = 单据号_In;
    If n_Count <> 0 Then
      v_Err_Msg := '挂号单据号重复,不能保存！' || Chr(13) || '如果使用了按日顺序编号,当日挂号量不能超过10000人次。';
      Raise Err_Item;
    End If;
  
    --获取结算方式名称
    Begin
      Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
    Exception
      When Others Then
        v_现金 := '现金';
    End;
    Begin
      Select 名称 Into v_个人帐户 From 结算方式 Where 性质 = 3;
    Exception
      When Others Then
        v_个人帐户 := '个人帐户';
    End;
  End If;

  n_序号 := 号序_In;
  Select Decode(To_Char(发生时间_In, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null)
  Into v_星期
  From Dual;

  --挂号获取安排
  Begin
    Select a.Id, a.序号控制, Nvl(b.限号数, 0), Nvl(b.限约数, 0)
    Into n_安排id, n_序号控制, n_限号数, n_限约数
    From 挂号安排 A, 挂号安排限制 B
    Where a.Id = b.安排id(+) And b.限制项目(+) = v_星期 And a.号码 = 号别_In;
  
  Exception
    When Others Then
      n_安排id := -1;
  End;

  --如果是病历费或者号别为空时不检查
  If Nvl(病历费_In, 0) = 0 Or 号别_In Is Not Null Then
    If n_安排id = -1 Then
      v_Err_Msg := '不存相应的挂号安排数据,请检查';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 1 Then
    --首先获取计划
    Begin
      Select ID
      Into n_计划id
      From 挂号安排计划
      Where 安排id = n_安排id And 审核时间 Is Not Null And
            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.生效时间) As 生效
             From 挂号安排计划 A
             Where a.审核时间 Is Not Null And 发生时间_In Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.安排id = n_安排id) And
            发生时间_In Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
            Nvl(失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'));
    
    Exception
      When Others Then
        n_计划id := 0;
    End;
    If Nvl(n_计划id, 0) <> 0 Then
      Begin
        --获取计划的限制
        Select a.Id, a.序号控制, Nvl(b.限号数, 0) As 限号数, Nvl(b.限约数, 0) As 限约数
        Into n_计划id, n_序号控制, n_限号数, n_限约数
        From 挂号安排计划 A, 挂号计划限制 B
        Where a.号码 = 号别_In And a.Id = n_计划id And a.审核时间 Is Not Null And a.Id = b.计划id(+) And b.限制项目(+) = v_星期;
      Exception
        When Others Then
          v_Err_Msg := '不存相应的挂号安排或计划数据,请检查';
          Raise Err_Item;
      End;
    End If;
  End If;

  --获取是否分时段
  Begin
    If Nvl(n_计划id, 0) = 0 Then
      Select Count(Rownum) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum <= 1;
      Select Decode(To_Char(发生时间_In, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排
      Where ID = n_安排id;
    Else
      Select Count(Rownum) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum <= 1;
      Select Decode(To_Char(发生时间_In, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排计划
      Where ID = n_计划id;
    End If;
  Exception
    When Others Then
      v_时间段 := Null;
  End;

  If v_时间段 Is Not Null And d_启用时间 Is Not Null And 序号_In = 1 Then
    --检查是否跨模式挂号安排
    Select To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_检查开始时间, d_检查结束时间
    From 时间段
    Where 时间段 = v_时间段 And 站点 Is Null And 号类 Is Null;
    If d_检查开始时间 > d_检查结束时间 Then
      d_检查结束时间 := d_检查结束时间 + 1;
    End If;
    If d_检查开始时间 < d_启用时间 And d_检查结束时间 > d_启用时间 Then
      --获取出诊记录id
      Begin
        Select a.Id
        Into n_出诊记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = 号别_In And 上班时段 = v_时间段 And 发生时间_In Between 开始时间 And 终止时间;
      Exception
        When Others Then
          n_出诊记录id := Null;
      End;
    End If;
  End If;

  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    --发生时间_in>Sysdate 发生时间>最大的时段时间--号序_in is null
    Begin
      Select Max(To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And Nvl(限制数量, 0) <> 0;
      n_追加号 := Case Sign(发生时间_In - d_最大序号时间)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_追加号 := 0;
    End;
  End If;
  d_时段时间 := 发生时间_In;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 Then
    --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
    Begin
      Select Nvl(序号, 0),
             To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
      Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And
            (序号, 安排id, 星期) In (Select Nvl(Max(序号), -1), 安排id, 星期
                               From 挂号安排时段
                               Where 安排id = n_安排id And 星期 = v_星期 And
                                     Decode(操作类型_In + n_追加号, 0, To_Char(发生时间_In, 'hh24:mi'),
                                            To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By 安排id, 星期);
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := 发生时间_In;
        n_时段限号 := 0;
        n_时段限约 := 0;
    End;
  End If;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 1 And n_分时段 > 0 Then
    --预约号,取计划
    Begin
      If Nvl(n_计划id, 0) = 0 Then
        --没计划生效,取安排的数据
        Select Nvl(序号, 0),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号安排时段 C
        Where 安排id = n_安排id And 星期 = v_星期 And
              (序号, 安排id, 星期) In
              (Select Nvl(Max(c.序号), -1), 安排id, 星期
               From 挂号安排时段 C
               Where 安排id = n_安排id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(发生时间_In, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 安排id, 星期);
      Else
        --有计划生效取计划
        --没生效，代表是从挂号计划时段查询
        Select Nvl(序号, -1),
               To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号计划时段 C
        Where 计划id = n_计划id And 星期 = v_星期 And
              (序号, 计划id, 星期) In
              (Select Nvl(Max(c.序号), -1), 计划id, 星期
               From 挂号计划时段 C
               Where 计划id = n_计划id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(发生时间_In, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 计划id, 星期);
      End If;
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := 发生时间_In;
        n_时段限号 := 0;
        n_时段限约 := 0;
    End;
  End If;

  If 序号_In = 1 Then
  
    --获取当前未使用的序号
    If Nvl(n_序号控制, 0) = 1 And n_分时段 = 0 Then
      --<序号控制 未设置时段 获取可用的最大序号,以及已经使用的数量>
      Begin
        --最大序号
        If 退号重用_In = 1 Then
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
        End If;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
      End;
      If n_序号 Is Null Then
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
      --<序号控制 未设置时段 获取可用的最大序号 以及已经使用的数量 --end>
    
      --非加号的情况需要检查是否超过了限制
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号时检查
          --启用序号控制未分时段 达到了限制
          If n_限号数 <= n_已用数量 And n_限号数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(发生时间_In), 'yyyy-mm-dd') || '已达到最大限约数！';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --启用序号控制,未分时段 加号情况   不处理,如果以后有限制条件以后补充
      End If;
    
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 0 Then
      --<--普通号分时段 这里只有预约一种情况-->
      If 操作类型_In = 0 Then
        --<正常预约挂号-->
        Begin
          Select Count(0) As 已挂数, Nvl(Sum(Decode(Nvl(Sign(a.日期 - d_时段时间), 0), 0, 1, 0)), 0) As 已约数
          Into n_已挂数, n_已约数
          From 挂号序号状态 A
          Where a.号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And
                状态 Not In (4, 5);
        Exception
          When Others Then
            n_已挂数 := 0;
            n_已约数 := 0;
        End;
      
        n_时段限约 := n_时段限号; --普通号分时段的情况,29中n_时段限约始终是0 这里特殊处理
        --检查限制数量
        If n_限约数 = 0 Then
          n_限约数 := n_限号数;
        End If;
        If n_时段限约 <= n_已约数 Or n_限约数 <= n_已约数 Then
          v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限约数！';
          Raise Err_Item;
        End If;
        If n_限号数 <= n_已挂数 Then
          v_Err_Msg := '号别' || 号别_In || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
          Raise Err_Item;
        End If;
      End If;
    
      --没有达到时段的限号数 号码在当前时段往后追加
    
      --获取当天挂出的最大号序
      If 退号重用_In = 1 Then
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 Not In (4, 5);
      Else
        Select Nvl(Max(序号), 0)
        Into n_挂出的最大序号
        From 挂号序号状态 A
        Where a.日期 = d_时段时间 And 号码 = 号别_In And 状态 <> 5;
      End If;
    
      --设置序号
      n_序号 := RPad(Nvl(n_时段序号, 0), Length(n_限号数) + Length(Nvl(n_时段序号, 0)), 0) + n_已约数 + 1;
      If n_序号 <= Nvl(n_挂出的最大序号, 0) Then
        n_序号 := Nvl(n_挂出的最大序号, 0) + 1;
      End If;
    
      --<--普通号分时段--End>
    Elsif Nvl(n_分时段, 0) > 0 And Nvl(n_序号控制, 0) = 1 Then
      --<启用序号控制 设置时段
      --专家号分时段
      Begin
        If 退号重用_In = 1 Then
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
        End If;
      Exception
        When Others Then
          n_已用序号 := 0;
          n_已用数量 := 0;
          n_已挂数   := 0;
      End;
    
      n_失效数 := 0;
      If Nvl(预约挂号_In, 0) = 0 Then
        n_预约有效时间 := Zl_To_Number(zl_GetSysParameter('预约有效时间', 1111));
        n_失约挂号     := Zl_To_Number(zl_GetSysParameter('失约用于挂号', 1111));
        If Nvl(n_预约有效时间, 0) <> 0 And Nvl(n_失约挂号, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_预约有效时间 / 24 / 60) - 日期), 1, 1, 0))
            Into n_失效数
            From 挂号序号状态
            Where 号码 = 号别_In And 日期 Between Trunc(Sysdate) And Sysdate And Nvl(预约, 0) = 1 And 状态 = 2;
          Exception
            When Others Then
              n_失效数 := 0;
          End;
        End If;
      End If;
    
      If 操作类型_In = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
        
          --挂号 追加号码时不检查时段限号数
          If n_时段限号 <= n_已用数量 And Nvl(n_追加号, 0) = 0 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 - n_失效数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --挂号
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
          If n_限约数 <= n_已用数量 Then
            v_Err_Msg := '号别' || 号别_In || '在时段' || To_Char(d_时段时间, 'hh24:mi') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
          If n_限号数 <= n_已挂数 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_序号 Is Null Then
        --设置序号
        If Nvl(n_已用序号, 0) < Nvl(n_时段序号, 0) Then
          n_已用序号 := Nvl(n_时段序号, 0);
        End If;
        n_序号 := Nvl(n_已用序号, 0) + 1;
      End If;
    Elsif Nvl(n_分时段, 0) = 0 And Nvl(n_序号控制, 0) = 0 And Nvl(病历费_In, 0) = 0 And Nvl(号别_In, 0) > 0 Then
      ---<--普通号  -->
      Begin
        Select 已挂数, 已约数
        Into n_已用数量, n_已约数
        From 病人挂号汇总
        Where 日期 = Trunc(发生时间_In) And 号码 = 号别_In;
      Exception
        When Others Then
          n_已用数量 := 0;
          n_已约数   := 0;
      End;
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(发生时间_In, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--普通号  -->
    End If;
  End If;

  --更新挂号序号状态
  If 序号_In = 1 And Not n_序号 Is Null Then
    If n_分时段 = 1 Then
      d_序号时间 := 发生时间_In;
    Else
      d_序号时间 := Trunc(发生时间_In);
    End If;
    --锁定序号的处理
    Begin
      Select 操作员姓名, 机器名
      Into v_序号操作员, v_序号机器名
      From 挂号序号状态
      Where 状态 = 5 And 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
      n_锁定 := 1;
    Exception
      When Others Then
        v_序号操作员 := Null;
        v_序号机器名 := Null;
        n_锁定       := 0;
    End;
    If n_锁定 = 0 Then
      Update 挂号序号状态
      Set 状态 = Decode(预约挂号_In, 1, 2, 1), 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
      Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 3 And 操作员姓名 = 操作员姓名_In;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_分时段, 0) = 0 Or Nvl(预约挂号_In, 0) = 1 Or (Nvl(n_序号控制, 0) = 0 And Nvl(号序_In, 0) = 0) Then
            Insert Into 挂号序号状态
              (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
            Values
              (号别_In, d_序号时间, n_序号, Decode(预约挂号_In, 1, 2, 1), 操作员姓名_In, Decode(预约接收_In, 1, 1, 0), Sysdate);
            If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
              Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
            End If;
          Elsif Nvl(n_分时段, 0) > 0 Then
            --分时段后专家号 失约的预约号允许挂号
            Update 挂号序号状态
            Set 状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In, 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
            Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 2;
            If Sql%NotFound Then
              Insert Into 挂号序号状态
                (号码, 日期, 序号, 状态, 操作员姓名, 预约, 登记时间)
              Values
                (号别_In, d_序号时间, n_序号, Decode(预约挂号_In, 1, 2, 1), 操作员姓名_In, Decode(预约接收_In, 1, 1, 0), Sysdate);
            End If;
            If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
              Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '序号' || n_序号 || '已被使用,请重新选择一个序号.';
            Raise Err_Item;
        End;
      End If;
    Else
      If 操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名 Then
        v_Err_Msg := '序号' || n_序号 || '已被自助机' || v_机器名 || '锁定,请重新选择一个序号.';
        Raise Err_Item;
      Else
        Update 挂号序号状态
        Set 状态 = Decode(预约挂号_In, 1, 2, 1), 预约 = Decode(预约接收_In, 1, 1, 0), 登记时间 = Sysdate
        Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号 And 状态 = 5 And 操作员姓名 = 操作员姓名_In And 机器名 = v_机器名;
        If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
          Update 挂号序号状态 Set 预约 = 1 Where 号码 = 号别_In And 日期 = d_序号时间 And 序号 = n_序号;
        End If;
      End If;
    End If;
  End If;

  If n_出诊记录id Is Not Null Then
    Update 临床出诊序号控制
    Set 挂号状态 = Decode(预约挂号_In, 1, 2, 1), 操作员姓名 = 操作员姓名_In
    Where 记录id = n_出诊记录id And 序号 = n_序号;
    If 预约挂号_In = 1 Then
      Update 临床出诊记录 Set 已约数 = 已约数 + 1 Where ID = n_出诊记录id;
    Else
      If 预约接收_In = 1 Then
        Update 临床出诊记录
        Set 已约数 = 已约数 + 1, 已挂数 = 已挂数 + 1, 其中已接收 = 其中已接收 + 1
        Where ID = n_出诊记录id;
      Else
        Update 临床出诊记录 Set 已挂数 = 已挂数 + 1 Where ID = n_出诊记录id;
      End If;
    End If;
  End If;

  --产生病人挂号费用(可能单独是或包括病历费用)
  Select 病人费用记录_Id.Nextval Into n_费用id From Dual; --应该通过程序得到

  Insert Into 门诊费用记录
    (ID, 记录性质, 记录状态, 序号, 价格父号, 从属父号, NO, 实际票号, 门诊标志, 加班标志, 附加标志, 发药窗口, 病人id, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 病人科室id, 收费类别,
     计算单位, 收费细目id, 收入项目id, 收据费目, 付数, 数次, 标准单价, 应收金额, 实收金额, 结帐金额, 结帐id, 记帐费用, 开单部门id, 开单人, 划价人, 执行部门id, 执行人, 操作员编号, 操作员姓名,
     发生时间, 登记时间, 保险大类id, 保险项目否, 保险编码, 统筹金额, 摘要, 结论, 缴款组id)
  Values
    (n_费用id, 4, Decode(预约挂号_In, 1, 0, 1), 序号_In, Decode(价格父号_In, 0, Null, 价格父号_In), 从属父号_In, 单据号_In, 票据号_In, 1, 急诊_In,
     病历费_In, Decode(预约挂号_In, 1, To_Char(n_序号), 诊室_In), Decode(病人id_In, 0, Null, 病人id_In),
     Decode(门诊号_In, 0, Null, 门诊号_In), 付款方式_In, 姓名_In, Decode(姓名_In, Null, Null, 性别_In), Decode(姓名_In, Null, Null, 年龄_In),
     v_费别, 病人科室id_In, 收费类别_In, 号别_In, 收费细目id_In, 收入项目id_In, 收据费目_In, 1, 数次_In, 标准单价_In, 应收金额_In, 实收金额_In,
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 实收金额_In)),
     Decode(预约挂号_In, 1, Null, Decode(Nvl(记帐费用_In, 0), 1, Null, 结帐id_In)), Decode(Nvl(记帐费用_In, 0), 1, 1, 0), 开单部门id_In,
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, 发生时间_In, 登记时间_In, 保险大类id_In,
     保险项目否_In, 保险编码_In, 统筹金额_In, Decode(收费单_In, Null, 摘要_In, '划价:' || 收费单_In), 预约方式_In, Decode(预约挂号_In, 1, Null, n_组id));

  --汇总结算到病人预交记录
  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 0 Then
  
    If (Nvl(现金支付_In, 0) <> 0 Or (Nvl(现金支付_In, 0) = 0 And Nvl(个帐支付_In, 0) = 0 And Nvl(预交支付_In, 0) = 0)) And 序号_In = 1 Then
      Select 病人预交记录_Id.Nextval Into n_预交id From Dual;
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位,
         结算性质)
      Values
        (n_预交id, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), Nvl(结算方式_In, v_现金), Nvl(现金支付_In, 0), 登记时间_In,
         操作员编号_In, 操作员姓名_In, 结帐id_In, '挂号收费', n_组id, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, 4);
    
      If Nvl(结算卡序号_In, 0) <> 0 And Nvl(现金支付_In, 0) <> 0 Then
        Zl_病人卡结算记录_支付(结算卡序号_In, 卡号_In, 0, 现金支付_In, n_预交id, 操作员编号_In, 操作员姓名_In, 登记时间_In);
      End If;
    End If;
  
    --对于医保挂号
    If Nvl(个帐支付_In, 0) <> 0 And 序号_In = 1 Then
      Insert Into 病人预交记录
        (ID, 记录性质, 记录状态, NO, 病人id, 结算方式, 冲预交, 收款时间, 操作员编号, 操作员姓名, 结帐id, 摘要, 缴款组id, 结算性质)
      Values
        (病人预交记录_Id.Nextval, 4, 1, 单据号_In, Decode(病人id_In, 0, Null, 病人id_In), v_个人帐户, 个帐支付_In, 登记时间_In, 操作员编号_In,
         操作员姓名_In, 结帐id_In, '医保挂号', n_组id, 4);
    End If;
  
    --对于就诊卡通过预交金挂号
    If Nvl(预交支付_In, 0) <> 0 And 序号_In = 1 Then
      n_预交金额 := 预交支付_In;
      For r_Deposit In c_Deposit(病人id_In, v_冲预交病人ids) Loop
        n_当前金额 := Case
                    When r_Deposit.金额 - n_预交金额 < 0 Then
                     r_Deposit.金额
                    Else
                     n_预交金额
                  End;
      
        If r_Deposit.结帐id = 0 Then
          --第一次冲预交(填上结帐ID,金额为0)
          Update 病人预交记录 Set 冲预交 = 0, 结帐id = 结帐id_In, 结算性质 = 4 Where ID = r_Deposit.原预交id;
        
        End If;
        --冲上次剩余额
        Insert Into 病人预交记录
          (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 预交类别, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
           冲预交, 结帐id, 缴款组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
          Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 预交类别, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                 登记时间_In, 操作员姓名_In, 操作员编号_In, n_当前金额, 结帐id_In, n_组id, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
          From 病人预交记录
          Where NO = r_Deposit.No And 记录状态 = r_Deposit.记录状态 And 记录性质 In (1, 11) And Rownum = 1;
      
        --更新病人预交余额
        Update 病人余额
        Set 预交余额 = Nvl(预交余额, 0) - n_当前金额
        Where 病人id = r_Deposit.病人id And 性质 = 1 And 类型 = Nvl(1, 2);
        --检查是否已经处理完
        If r_Deposit.金额 < n_预交金额 Then
          n_预交金额 := n_预交金额 - r_Deposit.金额;
        Else
          n_预交金额 := 0;
        End If;
      
        If n_预交金额 = 0 Then
          Exit;
        End If;
      End Loop;
      If n_预交金额 > 0 Then
        v_Err_Msg := '预交余不够支付本次支付金额,不能继续操作！';
        Raise Err_Item;
      
      End If;
      Delete From 病人余额 Where 病人id = 病人id_In And 性质 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
    End If;
  
    --相关汇总表的处理
    --人员缴款余额
    If 序号_In = 1 And Nvl(更新交款余额_In, 1) = 1 Then
      If Nvl(现金支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 现金支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金)
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额
            (收款员, 结算方式, 性质, 余额)
          Values
            (操作员姓名_In, Nvl(结算方式_In, v_现金), 1, 现金支付_In);
          n_返回值 := 现金支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 结算方式 = Nvl(结算方式_In, v_现金) And 性质 = 1 And Nvl(余额, 0) = 0;
        End If;
      End If;
    
      If Nvl(个帐支付_In, 0) <> 0 Then
        Update 人员缴款余额
        Set 余额 = Nvl(余额, 0) + 个帐支付_In
        Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = v_个人帐户
        Returning 余额 Into n_返回值;
      
        If Sql%RowCount = 0 Then
          Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, v_个人帐户, 1, 个帐支付_In);
          n_返回值 := 个帐支付_In;
        End If;
        If Nvl(n_返回值, 0) = 0 Then
          Delete From 人员缴款余额
          Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = v_个人帐户 And Nvl(余额, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --病人挂号汇总(只处理一次,且单独收取病历费不处理)
  If Nvl(预约挂号_In, 0) = 0 Then
    --病人本次就诊(以发生时间为准)
    If Nvl(病人id_In, 0) <> 0 And 序号_In = 1 Then
      Update 病人信息 Set 就诊时间 = 发生时间_In, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
    End If;
  End If;

  If Nvl(预约挂号_In, 0) = 0 And Nvl(记帐费用_In, 0) = 1 Then
    --记帐
    If Nvl(病人id_In, 0) = 0 Then
      v_Err_Msg := '要针对病人的挂号费进行记帐，必须是建档病人才能记帐挂号。';
      Raise Err_Item;
    End If;
  
    --病人余额
    Update 病人余额
    Set 费用余额 = Nvl(费用余额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = Nvl(病人id_In, 0) And 性质 = 1 And 类型 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人余额 (病人id, 性质, 类型, 费用余额, 预交余额) Values (病人id_In, 1, 1, Nvl(实收金额_In, 0), 0);
    End If;
  
    --病人未结费用
    Update 病人未结费用
    Set 金额 = Nvl(金额, 0) + Nvl(实收金额_In, 0)
    Where 病人id = 病人id_In And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(病人科室id_In, 0) And
          Nvl(开单部门id, 0) = Nvl(开单部门id_In, 0) And Nvl(执行部门id, 0) = Nvl(执行部门id_In, 0) And 收入项目id + 0 = 收入项目id_In And
          来源途径 + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into 病人未结费用
        (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
      Values
        (病人id_In, Null, Null, 病人科室id_In, 开单部门id_In, 执行部门id_In, 收入项目id_In, 1, Nvl(实收金额_In, 0));
    End If;
  End If;

  --病人挂号记录
  If 号别_In Is Not Null And 序号_In = 1 Then
    --And Nvl(预约挂号_In, 0) = 0
    Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
    Begin
      Select 名称 Into v_付款方式 From 医疗付款方式 Where 编码 = 付款方式_In And Rownum < 2;
    Exception
      When Others Then
        v_付款方式 := Null;
    End;
    Insert Into 病人挂号记录
      (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 预约时间, 操作员编号,
       操作员姓名, 复诊, 号序, 社区, 预约, 预约方式, 摘要, 交易流水号, 交易说明, 合作单位, 接收时间, 接收人, 预约操作员, 预约操作员编号, 险类, 医疗付款方式, 收费单)
    Values
      (n_挂号id, 单据号_In, Decode(Nvl(预约挂号_In, 0), 1, 2, 1), 1, 病人id_In, 门诊号_In, 姓名_In, 性别_In, 年龄_In, 号别_In, 急诊_In, 诊室_In,
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, 发生时间_In, Decode(Nvl(预约挂号_In, 0), 1, 发生时间_In, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 收费单_In);
    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
      Update 病人挂号记录
      Set 预约 = 1, 预约时间 = 发生时间_In, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In
      Where ID = n_挂号id;
    End If;
    n_预约生成队列 := 0;
    If Nvl(预约挂号_In, 0) = 1 Then
      n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
    End If;
  
    --0-不产生队列;1-按医生或分诊台排队;2-先分诊,后医生站
    If Nvl(生成队列_In, 0) <> 0 And Nvl(预约挂号_In, 0) = 0 Or n_预约生成队列 = 1 Then
      n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
      If Nvl(n_分诊台签到排队, 0) = 0 Or n_预约生成队列 = 1 Then
        n_分时点显示 := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(预约挂号_In, 0) = 1 And n_分时点显示 = 1 And n_分时段 = 1 Then
          n_分时点显示 := 1;
        Else
          n_分时点显示 := Null;
        End If;
      
        --产生队列
        --.按”执行部门” 的方式生成队列
        v_队列名称 := 执行部门id_In;
        v_排队号码 := Zlgetnextqueue(执行部门id_In, n_挂号id, 号别_In || '|' || n_序号);
        v_排队序号 := Zlgetsequencenum(0, n_挂号id, 0);
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, 发生时间_In);
        --  队列名称_In , 业务类型_In, 业务id_In,科室id_In,排队号码_In,排队标记_In,患者姓名_In,病人ID_IN, 诊室_In, 医生姓名_In, 排队时间_In
        Zl_排队叫号队列_Insert(v_队列名称, 0, n_挂号id, 执行部门id_In, v_排队号码, Null, 姓名_In, 病人id_In, 诊室_In, 医生姓名_In, d_排队时间, 预约方式_In,
                         n_分时点显示, v_排队序号);
      
        --挂号立即排队
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          Update 病人挂号记录 Set 记录标志 = 1 Where ID = n_挂号id;
        End If;
      End If;
    End If;
  End If;
  --病人担保信息
  If 病人id_In Is Not Null And 序号_In = 1 Then
    --取参数:
    If Nvl(n_结算模式, 0) <> Nvl(结算模式_In, 0) Then
      --结算模式的确定
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(结算模式, 0) Into n_结算模式 From 病人信息 Where 病人id = 病人id_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的病人信息,不允许挂号';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_结算模式 = 1 And Nvl(结算模式_In, 0) = 0 Then
        --病人已经是"先诊疗后结算的",本次是"先结算后诊疗的",则检查是否存在未结数据
        Select Count(1)
        Into n_Count
        From 病人未结费用
        Where 病人id = 病人id_In And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --存在未结算数据，必须先结算后才允许执行
          v_Err_Msg := '当前病人的就诊模式为先诊疗后结算且存在未结费用，不允许调整该病人的就诊模式,你可以先对未结费用结帐后再挂号或不调整病人的就诊模式!';
          Raise Err_Item;
        End If;
        --检查
        --未发生医嘱业务的（即当时就挂号的,需要保证同一次的就诊模式是一至的(程序已经检查，不用再处理)
      End If;
      Update 病人信息 Set 结算模式 = 结算模式_In Where 病人id = 病人id_In;
    End If;
  
    Update 病人信息
    Set 担保人 = Null, 担保额 = Null, 担保性质 = Null
    Where 病人id = 病人id_In And Nvl(在院, 0) = 0 And Exists
     (Select 1
           From 病人担保记录
           Where 病人id = 病人id_In And 主页id Is Not Null And
                 登记时间 = (Select Max(登记时间) From 病人担保记录 Where 病人id = 病人id_In));
  
    If Sql%RowCount > 0 Then
      Update 病人担保记录
      Set 到期时间 = Sysdate
      Where 病人id = 病人id_In And 主页id Is Not Null And Nvl(到期时间, Sysdate) >= Sysdate;
    End If;
  End If;
  If 序号_In = 1 Then
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 1, n_挂号id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_挂号id, 单据号_In);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人挂号记录_Insert;
/

--127271:焦博,2018-06-19,结帐作废时,清空病人结帐记录中的实际票好
CREATE OR REPLACE Procedure Zl_病人结帐异常_Update
(
  登记时间_In 门诊费用记录.登记时间%Type,
  结帐id_In   门诊费用记录.结帐id%Type := Null
) As
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  d_Date Date;
  v_No   门诊费用记录.No%Type;
Begin
  --功能：更新异常单据的登记时间及收款时间
  --结帐ID_IN: 传入时,以结帐ID进行更新操作;否则以NO_IN进行操作

  d_Date := 登记时间_In;
  If d_Date Is Null Then
    d_Date := Sysdate;
  End If;

  --更新指定结帐的门诊费用及预交费的登记时间
  Update 病人结帐记录 Set 收费时间 = d_Date Where ID = 结帐id_In Returning NO Into v_No;
  Update 病人结帐记录 Set 实际票号 = Null Where NO = v_No;
  Update 病人预交记录
  Set 收款时间 = d_Date
  Where 结帐id = 结帐id_In And ((记录性质 = 1 And 结算性质 = 12) Or (记录性质 <> 1));

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人结帐异常_Update;
/







------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0017' Where 编号=&n_System;
Commit;
