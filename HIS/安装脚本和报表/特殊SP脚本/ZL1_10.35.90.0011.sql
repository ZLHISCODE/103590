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
--125974:胡俊勇,2018-05-18,西医科录入中医诊断
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1252, 0, 0, 0, 0, 0, 0, 68, '门诊西医科允许录入中医诊断', '0', '0',
         '若启用参数，病人所属科室无“中医科”属性时填写诊断时可以下达中医诊断', '0-表示不启用,1-表示启用', '本参数仅用于门诊场合', '适用于西医科需要录入中医诊断的情况', Null
  From Dual;  

--119329:冉俊明,2018-05-16,三方接口获取可挂号科室过程调整
Declare
  --功能：修正临床出诊挂号控制
  Cursor c_限制 Is
    Select Rowid From 临床出诊挂号控制 Where 控制方式 = 3 And 序号 = 0 And 数量 = 0;

  c_Rowid      t_Strlist := t_Strlist();
  n_Array_Size Number := 10000; --每批一万,多了可能PGA不够
  I            Number(8) := 0; --每修正10万条记录提交一次,多了可能Undo不够,少了提交过于频繁
  J            Number(16) := 0;
Begin
  Open c_限制();
  Loop
    Fetch c_限制 Bulk Collect
      Into c_Rowid Limit n_Array_Size;
    Exit When c_Rowid.Count = 0;
  
    Forall K In 1 .. c_Rowid.Count
      Update 临床出诊挂号控制 Set 控制方式 = 4 Where Rowid = c_Rowid(K);
  
    J := J + c_Rowid.Count;
    If I = 10 Then
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Close c_限制;
  Commit;
End;
/

--119329:冉俊明,2018-05-16,三方接口获取可挂号科室过程调整
Declare
  --功能：修正临床出诊挂号控制记录
  Cursor c_记录 Is
    Select Rowid From 临床出诊挂号控制记录 Where 控制方式 = 3 And 序号 = 0 And 数量 = 0;

  c_Rowid      t_Strlist := t_Strlist();
  n_Array_Size Number := 10000; --每批一万,多了可能PGA不够
  I            Number(8) := 0; --每修正10万条记录提交一次,多了可能Undo不够,少了提交过于频繁
  J            Number(16) := 0;
Begin
  Open c_记录();
  Loop
    Fetch c_记录 Bulk Collect
      Into c_Rowid Limit n_Array_Size;
    Exit When c_Rowid.Count = 0;
  
    Forall K In 1 .. c_Rowid.Count
      Update 临床出诊挂号控制记录 Set 控制方式 = 4 Where Rowid = c_Rowid(K);
  
    J := J + c_Rowid.Count;
    If I = 10 Then
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Close c_记录;
  Commit;
End;
/




-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--125932:胡俊勇,2018-05-17,对象权限缺失
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1252,'医嘱下达',User,'Zl_Fun_Bloodapplyrate','EXECUTE' From Dual;

--125588:秦龙,2018-05-16,排除调价预减记录
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1333, '基本', User, 'Zl_Fun_Getbatchpro', 'EXECUTE'
  From Dual;






-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--122403:殷瑞,2018-06-01,修正医嘱发送时分解输液单批次的错误
Create Or Replace Procedure Zl_输液配药记录_核查
(
  部门id_In   In 输液配药记录.部门id%Type,
  医嘱id_In   In Varchar2, --输液医嘱给药途径对应的医嘱ID:医嘱ID1,医嘱ID2...
  发送号_In   In 病人医嘱发送.发送号%Type,
  核查人_In   In 输液配药状态.操作人员%Type,
  核查时间_In In 输液配药状态.操作时间%Type
) Is
  v_Count    Number;
  v_序号     Number;
  v_执行时间 Date;

  v_相关id      Number;
  v_New相关id   Number;
  v_Old相关id   Number;
  v_发送号      Number;
  v_Tmp         Varchar2(200);
  I             Number;
  v_配药id      Number;
  v_批次        Number;
  v_Maxno       Varchar2(4000);
  v_Lableno     Varchar2(200);
  v_Maxbatch    Number;
  v_Curdose     Number;
  v_Sumdose     Number;
  v_Drugcount   Number;
  v_Currdate    Date;
  n_Needcheck   Number;
  n_Lngid       药品收发记录.Id%Type;
  n_Count       Number(3);
  n_单据        药品收发记录.单据%Type;
  v_No          药品收发记录.No%Type;
  n_发送次数    Number(5);
  n_病人id      病人信息.病人id%Type := 0;
  b_Change      Boolean := True;
  n_Sum         Number;
  n_调整批次    Number(1);
  n_Cur         Number(5);
  v_上次发送号  病人医嘱发送.发送号%Type;
  v_医嘱ids     Varchar2(4000);
  v_Tansid      Varchar2(12);
  v_当前病人    Varchar2(20);
  n_Num         Number(8);
  d_Old执行时间 Date;
  n_是否打包    Number(1);
  n_打包        Number(1);
  n_摆药单      Number(2);
  --控制参数
  v_医嘱类型       Number;
  v_输液总量       Number;
  v_大输液剂型     Varchar2(2000);
  v_大输液给药途径 Varchar2(2000);
  v_来源科室       Varchar2(4000);
  v_Continue       Number := 1;
  v_Nodosage       Number := 0;
  v_保持上次批次   Number := 0;
  d_手工打包时间   Date;
  n_Tpn处置方式    Number := 0;
  v_药品类型       Varchar2(20);
  n_打包药品批次   Number(1);
  n_特殊药品批次   Number(1);
  n_优先级         Number := 999;
  n_自动排批       Number := 0;
  n_科室id         Number := 0;
  n_Row            Number(2);
  n_备用批次       Number := 0;
  n_剩余数量       Number := 0;
  n_单次数量       Number := 0;
  n_累计数量       Number := 0;
  n_医嘱id         Number := 0;
  n_填写数量       Number := 0;
  v_配药类型       Varchar2(20);
  v_时间串         Varchar2(100);
  v_时间值         Date;
  v_Fields         Varchar2(100);
  v_是否改变       Varchar2(20);
  v_时间串1        Varchar2(100);
  Err_Item Exception;
  n_流通金额小数 Number;

  Cursor c_医嘱记录 Is
    Select /*+rule */
    Distinct e.医嘱id As 相关id, e.发送号, b.频率间隔, b.间隔单位, b.执行时间方案, a.姓名, a.性别, a.年龄, a.住院号, a.当前床号 As 床号, a.当前病区id As 病人病区id,
             a.当前科室id As 病人科室id, e.首次时间, e.末次时间, b.开始执行时间, Nvl(e.发送数次, 0) As 次数, e.发送时间, Decode(b.医嘱期效, 0, 1, 2) As 医嘱类型,
             b.诊疗项目id As 给药途径, b.病人id, Nvl(c.执行标记, 0) As 是否tpn
    From 病人医嘱发送 E, 病人医嘱记录 B, 病人信息 A, 诊疗项目目录 C, Table(f_Num2list(医嘱id_In)) D
    Where e.医嘱id = b.Id And b.病人id = a.病人id And c.类别 = 'E' And c.操作类型 = '2' And c.执行分类 = 1 And b.诊疗项目id = c.Id And
          e.医嘱id = d.Column_Value And e.发送号 = 发送号_In
    Order By b.病人id, e.医嘱id, e.发送号;

  Cursor c_单个医嘱记录 Is
    Select /*+rule */
    Distinct e.医嘱id, e.发送号, b.频率间隔, b.间隔单位, b.执行时间方案, a.姓名, a.性别, a.年龄, a.住院号, a.当前床号 As 床号, a.当前病区id As 病人病区id,
             a.当前科室id As 病人科室id, e.首次时间, e.末次时间, b.开始执行时间, Nvl(e.发送数次, 0) As 次数, e.发送时间, Decode(b.医嘱期效, 0, 1, 2) As 医嘱类型,
             b.诊疗项目id As 给药途径, b.病人id
    From 病人医嘱发送 E, 病人医嘱记录 B, 病人信息 A, 诊疗项目目录 C
    Where e.医嘱id = b.Id And b.病人id = a.病人id And b.诊疗项目id = c.Id And b.相关id = v_相关id And e.发送号 = 发送号_In
    Order By e.医嘱id, e.发送号;

  Cursor c_收发记录 Is
    Select Distinct c.Id As 收发id, c.序号, c.实际数量 As 数量, Nvl(e.是否不予配置, 0) As 是否不予配置, c.单据, c.No
    From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, 输液药品属性 E
    Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id = e.药品id(+) And
          b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And a.Id = n_医嘱id And b.发送号 = 发送号_In And c.序号 < 1000
    Order By c.No, c.序号;

  Cursor c_原始收发记录 Is
    Select Distinct c.Id As 收发id, c.序号, c.实际数量 As 数量, Nvl(e.是否不予配置, 0) As 是否不予配置, c.单据, c.No
    From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, 输液药品属性 E
    Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And c.药品id = e.药品id(+) And
          b.执行部门id + 0 = 部门id_In And c.单据 = 9 And c.审核日期 Is Null And a.相关id = v_相关id And b.发送号 = 发送号_In And c.序号 < 1000
    Order By c.No, c.序号;

  Cursor c_输液单记录 Is
    Select a.Id, a.执行时间, a.配药批次, a.医嘱id, d.发送时间
    From 输液配药记录 A, 病人医嘱记录 B, 配药工作批次 C, 病人医嘱发送 D
    Where a.医嘱id = b.Id And a.配药批次 = c.批次 And d.医嘱id = a.医嘱id And a.发送号 = d.发送号 And c.批次 <> 0 And c.药品类型 Is Null And
          b.病人id = n_病人id And a.操作状态 < 2 And a.执行时间 Between Trunc(v_时间值) And Trunc(v_时间值 + 1) - 1 / 24 / 60 / 60;

  v_输液单记录   c_输液单记录%RowType;
  v_医嘱记录     c_医嘱记录%RowType;
  v_收发记录     c_收发记录%RowType;
  v_单个医嘱记录 c_单个医嘱记录%RowType;

  Function Zl_Getpivaworkbatch
  (
    执行时间_In In Date,
    发送时间_In In Date,
    药品类型_In In Varchar2 := Null
  ) Return Number As
  
    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_配药批次 Is
      Select 批次, 配药时间, 给药时间, 打包, 药品类型
      From 配药工作批次
      Where 启用 = 1 And 配置中心id = 部门id_In
      Order By 药品类型, 批次;
  
    v_配药批次 c_配药批次%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(执行时间_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');
  
    Select Max(Nvl(批次, 0)) + 1 Into v_Maxbatch From 配药工作批次 Where 启用 = 1 And 配置中心id = 部门id_In;
  
    For v_配药批次 In c_配药批次 Loop
      v_Batch := 0;
    
      --当天发送的医嘱发送到备用批次
      If (Trunc(执行时间_In) >= Trunc(v_Currdate) And Trunc(发送时间_In) < Trunc(执行时间_In)) Or n_备用批次 = 0 Then
        If v_配药批次.批次 <> '0' And
           ((Nvl(v_配药批次.药品类型, '0') <> '0' And v_配药批次.药品类型 = 药品类型_In) Or Nvl(v_配药批次.药品类型, '0') = '0') Then
          v_Starttime := To_Date(Substr(v_配药批次.给药时间, 1, Instr(v_配药批次.给药时间, '-') - 1), 'hh24:mi');
          v_Endtime   := To_Date(Substr(v_配药批次.给药时间, Instr(v_配药批次.给药时间, '-') + 1), 'hh24:mi');
        
          If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
            v_Batch := v_配药批次.批次;
            n_打包  := v_配药批次.打包;
            Exit When v_Batch > 0;
          End If;
        End If;
      End If;
    End Loop;
  
    If v_Batch = 0 And (n_打包药品批次 <> 1 Or n_备用批次 = 1) Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;

  Function Zl_Getfirst
  (
    配药id_In In Number,
    科室id_In In Number
  ) Return Number As
    n_First  Number;
    n_科室id Number;
    Cursor c_优先级 Is
      Select 科室id, 配药类型, 优先级, 频次
      From 输液药品优先级
      Where (科室id = 科室id_In Or 科室id = 0)
      Order By 科室id, 优先级 Desc;
  
    r_优先级 c_优先级%RowType;
  Begin
    n_First := 0;
    For r_优先级 In c_优先级 Loop
      If n_科室id <> 0 And r_优先级.科室id = 0 Then
        Exit;
      End If;
      n_科室id := r_优先级.科室id;
    
      For r_配药记录 In (Select Distinct d.配药类型, e.执行频次
                     From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C, 输液药品属性 D, 病人医嘱记录 E
                     Where a.医嘱id = e.Id And a.Id = b.记录id And b.收发id = c.Id And c.药品id = d.药品id And a.Id = 配药id_In) Loop
        If Instr(r_配药记录.配药类型, r_优先级.配药类型, 1) > 0 And (Instr(r_优先级.频次, r_配药记录.执行频次, 1) > 0 Or r_优先级.频次 = '所有频次') Then
          n_First := r_优先级.优先级;
          Exit;
        End If;
      End Loop;
    End Loop;
  
    If n_First = 0 Then
      n_First := 999;
    End If;
    Return(n_First);
  End;
Begin
  n_Count          := 0;
  v_医嘱类型       := Zl_To_Number(Nvl(zl_GetSysParameter('医嘱类型', 1345), 1));
  v_输液总量       := Zl_To_Number(Nvl(zl_GetSysParameter('同批次输液总量', 1345), 0));
  v_大输液剂型     := Nvl(zl_GetSysParameter('大输液药品剂型', 1345), '');
  v_大输液给药途径 := Nvl(zl_GetSysParameter('输液给药途径', 1345), '');
  v_来源科室       := Nvl(zl_GetSysParameter('来源科室', 1345), '');
  v_保持上次批次   := Zl_To_Number(Nvl(zl_GetSysParameter('保持上次批次', 1345), 0));
  n_Tpn处置方式    := Zl_To_Number(Nvl(zl_GetSysParameter('静脉营养药物处置方式', 1345), 0));
  n_打包药品批次   := Zl_To_Number(Nvl(zl_GetSysParameter('单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包', 1345), 0));
  n_特殊药品批次   := Zl_To_Number(Nvl(zl_GetSysParameter('特殊药品按药品类型指定批次', 1345), 0));
  n_自动排批       := Zl_To_Number(Nvl(zl_GetSysParameter('启动自动排批', 1345), 0));
  n_备用批次       := Zl_To_Number(Nvl(zl_GetSysParameter('当天发送的医嘱产生的输液单全部到备用批次', 1345), 0));
  v_医嘱ids        := 医嘱id_In;
  v_当前病人       := '';
  n_发送次数       := 0;

  --取流通业务精度位数
  --类别:1-药品 2-卫材
  --内容：2-零售价 4-金额
  --单位：药品:1-售价 5-金额单位
  Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;

  Select Trunc(Sysdate) Into v_Currdate From Dual;

  Select Max(Nvl(批次, 0)) + 1 Into v_Maxbatch From 配药工作批次;

  --检查当前病人的医嘱是否有今天需要执行的输液单是锁定状态的
  If Instr(v_医嘱ids, ',') = 0 Then
    v_Tansid := v_医嘱ids;
  Else
    v_Tansid := Substr(v_Tmp, 1, Instr(v_医嘱ids, ',') - 1);
  End If;

  Select Count(ID)
  Into n_Num
  From 输液配药记录
  Where 是否锁定 = 1 And 执行时间 Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
        医嘱id In
        (Select 相关id
         From 病人医嘱记录
         Where 病人id = (Select 病人id From 病人医嘱记录 Where 相关id = v_Tansid And Rownum < 2) And (诊疗类别 = '5' Or 诊疗类别 = '6')) And
        Rownum < 2;

  If n_Num > 0 Then
    Select 姓名
    Into v_当前病人
    From 输液配药记录
    Where 是否锁定 = 1 And 执行时间 Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
          医嘱id In
          (Select 相关id
           From 病人医嘱记录
           Where 病人id = (Select 病人id From 病人医嘱记录 Where 相关id = v_Tansid And Rownum < 2) And (诊疗类别 = '5' Or 诊疗类别 = '6')) And
          Rownum < 2;
    Raise Err_Item;
  End If;

  --先将原收发记录的序号增大，新的收发记录产生后再删除
  --Update 药品收发记录
  --Set 序号 = 序号 + 10000
  --Where ID In (Select \*+rule *\
  --             Distinct c.Id
  --             From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D, Table(f_Num2list(医嘱id_In)) F
  --             Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And b.执行部门id + 0 = 部门id_In And
  --                   c.单据 = 9 And c.审核日期 Is Null And a.相关id = f.Column_Value And b.发送号 = 发送号_In And c.序号 < 10000);

  For v_医嘱记录 In c_医嘱记录 Loop
    v_Continue := 1;
    n_病人id   := v_医嘱记录.病人id;
    n_科室id   := v_医嘱记录.病人科室id;
  
    Select Count(1)
    Into v_Continue
    From 病人医嘱记录 A, 输液不配置药品 B, 住院费用记录 C
    Where c.收费细目id = b.药品id And c.医嘱序号 = a.Id And a.相关id = v_医嘱记录.相关id;
    If v_Continue = 0 Then
      v_Continue := 1;
    Else
      v_Continue := 0;
    End If;
  
    --参数控制产生输液单
    If (v_医嘱类型 = 1 And v_医嘱记录.医嘱类型 <> 1) Or (v_医嘱类型 = 2 And v_医嘱记录.医嘱类型 <> 2) Then
      v_Continue := 0;
    End If;
  
    If Not v_大输液给药途径 Is Null Then
      If Instr(',' || v_大输液给药途径 || ',', ',' || v_医嘱记录.给药途径 || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;
  
    If Not v_来源科室 Is Null Then
      If Instr(',' || v_来源科室 || ',', ',' || v_医嘱记录.病人科室id || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;
  
    v_药品类型 := Null;
    For r_药品类型 In (Select Decode(Nvl(d.抗生素, 0), 0, Decode(Nvl(d.是否肿瘤药, 0), 0, '', '肿瘤药'), '抗生素') 药品类型
                   From 病人医嘱记录 A, 药品规格 B, 住院费用记录 C, 药品特性 D
                   Where c.收费细目id = b.药品id And b.药名id = d.药名id And c.医嘱序号 = a.Id And a.相关id = v_医嘱记录.相关id) Loop
      If r_药品类型.药品类型 Is Not Null Then
        v_药品类型 := r_药品类型.药品类型;
      End If;
    End Loop;
  
    If v_药品类型 Is Null Then
      If v_医嘱记录.是否tpn = 2 Then
        v_药品类型 := '营养药';
      End If;
    End If;
  
    If v_Continue = 1 Then
      v_Old相关id := v_New相关id;
      v_相关id    := v_医嘱记录.相关id;
      v_New相关id := v_相关id;
      v_发送号    := v_医嘱记录.发送号;
      v_序号      := 0;
    
      If v_Continue = 1 Then
        --v_Count := Zl_Gettransexenumber(v_医嘱记录.开始执行时间, v_医嘱记录.首次时间, v_医嘱记录.末次时间, v_医嘱记录.频率间隔, v_医嘱记录.间隔单位, v_医嘱记录.执行时间方案);
        Select Count(医嘱id)
        Into v_Count
        From 医嘱执行时间
        Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号;
      
        v_Nodosage := 0;
      
        For I In 1 .. v_Count Loop
          Select 输液配药记录_Id.Nextval Into v_配药id From Dual;
          v_序号 := v_序号 + 1;
        
          If I > 1 Then
            --从医嘱执行时间表中取医嘱的执行时间
            Select 要求时间
            Into v_执行时间
            From 医嘱执行时间
            Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And 要求时间 > v_执行时间 And Rownum = 1
            Order By 要求时间;
          Else
            Select Min(要求时间)
            Into v_执行时间
            From 医嘱执行时间
            Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And Rownum = 1
            Order By 要求时间;
          End If;
        
          v_批次 := 0;
        
          If d_Old执行时间 <> Trunc(v_执行时间) Or d_Old执行时间 Is Null Then
            b_Change := True;
          End If;
        
          If b_Change = True Then
            If d_Old执行时间 <> Trunc(v_执行时间) Or d_Old执行时间 Is Null Then
              d_Old执行时间 := v_执行时间;
            
              Select Count(Distinct a.摆药单号)
              Into n_摆药单
              From 输液配药记录 A
              Where a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = v_医嘱记录.病人id And 相关id Is Null) And
                    a.执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60 And 操作状态 >= 2 And 操作状态 < 9;
            
              If n_摆药单 > 1 Then
                Update 输液配药记录
                Set 是否调整批次 = 1
                Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And
                     
                      执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                b_Change := False;
              
              End If;
            End If;
          End If;
        
          If b_Change = True Then
            n_病人id := v_医嘱记录.病人id;
            Select Count(ID)
            
            Into n_Sum
            From 输液配药记录
            Where 医嘱id = v_医嘱记录.相关id And 执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60;
            If n_Sum = 0 Then
              Update 输液配药记录
              Set 是否调整批次 = 1
              Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And
                   
                    执行时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
              b_Change := False;
            
            End If;
          
            If b_Change = True Then
              --检查输液单是否调整到打包状态
              Select Count(a.Id)
              Into n_Sum
              From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C
              Where a.Id = b.记录id And b.收发id = c.Id And
                    a.医嘱id In (Select ID
                               From 病人医嘱记录
                               Where 病人id = (Select 病人id From 病人医嘱记录 Where ID = v_医嘱记录.相关id And Rownum < 2)) And
                    a.执行时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60 And a.打包时间 Is Not Null;
              If n_Sum <> 0 Then
                Update 输液配药记录
                Set 是否调整批次 = 1
                Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_执行时间) And
                      Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                b_Change := False;
              End If;
            
              Select Count(医嘱id)
              Into n_Cur
              From 医嘱执行时间
              Where 医嘱id = v_医嘱记录.相关id And 要求时间 Between Trunc(v_执行时间) And Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60;
            
              Select Count(医嘱id)
              Into n_Sum
              From 医嘱执行时间
              Where 医嘱id = v_医嘱记录.相关id And 要求时间 Between Trunc(v_执行时间 - 1) And Trunc(v_执行时间) - 1 / 24 / 60 / 60;
            
              If n_Sum <> n_Cur Then
                Update 输液配药记录
                Set 是否调整批次 = 1
                Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_执行时间) And
                      Trunc(v_执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
                b_Change := False;
              End If;
            End If;
          End If;
        
          If v_时间串 <> Trunc(Sysdate) || ';false\' Or v_时间串 Is Null Then
            If Trunc(v_执行时间) = Trunc(Sysdate) Then
              If b_Change = False Then
                v_时间串 := Trunc(v_执行时间) || ';false\';
              Else
                v_时间串 := Trunc(v_执行时间) || ';true\';
              End If;
            End If;
          End If;
        
          If v_时间串1 <> Trunc(Sysdate + 1) || ';false\' Or v_时间串1 Is Null Then
            If Trunc(v_执行时间) = Trunc(Sysdate + 1) Then
              If b_Change = False Then
                v_时间串1 := Trunc(v_执行时间) || ';false\';
              Else
                v_时间串1 := Trunc(v_执行时间) || ';true\';
              End If;
            End If;
          End If;
        
          If v_药品类型 Is Null Or n_特殊药品批次 = 0 Then
            v_批次 := Zl_Getpivaworkbatch(v_执行时间, Sysdate);
          Else
            --药品类型不为空，直接根据药品类型匹配批次
            v_批次 := Zl_Getpivaworkbatch(v_执行时间, Sysdate, v_药品类型);
          End If;
        
          Select Count(医嘱id)
          Into n_发送次数
          From 医嘱执行时间
          Where 医嘱id = v_医嘱记录.相关id And 要求时间 <= v_执行时间
          Order By 要求时间;
        
          If n_发送次数 > 99 Then
            n_发送次数 := Mod(n_发送次数, 99);
          End If;
        
          If Length(v_医嘱记录.相关id) > 9 Then
            If n_发送次数 < 10 Then
              Select '91' || Substr(To_Char(v_医嘱记录.相关id), Length(v_医嘱记录.相关id) - 8) || To_Char(v_医嘱记录.相关id) || '0' ||
                      To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr(To_Char(v_医嘱记录.相关id), Length(v_医嘱记录.相关id) - 8) || To_Char(v_医嘱记录.相关id) ||
                      To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            End If;
          Else
            If n_发送次数 < 10 Then
              Select '91' || Substr('000000000', Length(v_医嘱记录.相关id) + 1) || To_Char(v_医嘱记录.相关id) || '0' ||
                      To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr('000000000', Length(v_医嘱记录.相关id) + 1) || To_Char(v_医嘱记录.相关id) || To_Char(n_发送次数)
              Into v_Maxno
              From Dual;
            End If;
          End If;
          n_调整批次 := 0;
          If b_Change = False Then
            n_调整批次 := 1;
          End If;
        
          If v_批次 <> 0 Then
            Select Nvl(Max(打包), 0), Max(药品类型)
            Into n_打包, v_配药类型
            From 配药工作批次
            Where 批次 = v_批次 And 配置中心id = 部门id_In;
          End If;
        
          If (Trunc(v_执行时间) <= v_Currdate Or n_打包 <> 0) And v_配药类型 Is Null Then
            n_是否打包     := 1;
            d_手工打包时间 := Null;
          Else
            n_是否打包     := 0;
            d_手工打包时间 := Null;
          End If;
        
          --如果是TPN不管其他条件如何都设置为配置
          If v_医嘱记录.是否tpn = 2 Then
            n_是否打包     := 0;
            d_手工打包时间 := Null;
          End If;
        
          If v_批次 = 0 Then
            n_是否打包 := 1;
          End If;
          --产生配药记录
          Insert Into 输液配药记录
            (ID, 部门id, 序号, 姓名, 性别, 年龄, 住院号, 床号, 病人病区id, 病人科室id, 执行时间, 医嘱id, 发送号, 配药批次, 瓶签号, 是否调整批次, 是否打包, 打包时间, 操作状态,
             操作人员, 操作时间)
          Values
            (v_配药id, 部门id_In, v_序号, v_医嘱记录.姓名, v_医嘱记录.性别, v_医嘱记录.年龄, v_医嘱记录.住院号, v_医嘱记录.床号, v_医嘱记录.病人病区id,
             v_医嘱记录.病人科室id, v_执行时间, v_医嘱记录.相关id, v_医嘱记录.发送号, v_批次, v_Maxno, n_调整批次, n_是否打包, d_手工打包时间, 1, 核查人_In, 核查时间_In);
        
          Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (v_配药id, 1, 核查人_In, 核查时间_In);
        
          For v_单个医嘱记录 In c_单个医嘱记录 Loop
            n_医嘱id   := v_单个医嘱记录.医嘱id;
            n_累计数量 := 0;
            n_剩余数量 := 0;
          
            Select Sum(c.实际数量)
            Into n_Sum
            From 病人医嘱记录 A, 病人医嘱发送 B, 药品收发记录 C, 住院费用记录 D
            Where a.Id = b.医嘱id And c.费用id = d.Id And a.Id = d.医嘱序号 And b.No = c.No And b.执行部门id + 0 = 部门id_In And
                  c.单据 = 9 And c.审核日期 Is Null And a.Id = n_医嘱id And b.发送号 = v_医嘱记录.发送号 And c.序号 < 1000;
          
            --产生配药记录对应的药品记录
            For v_收发记录 In c_收发记录 Loop
              If v_收发记录.是否不予配置 = 1 Then
                v_Nodosage := 1;
              End If;
            
              Select 药品收发记录_Id.Nextval Into n_Lngid From Dual;
              n_累计数量 := n_累计数量 + v_收发记录.数量;
            
              If n_剩余数量 = 0 Then
                n_剩余数量 := n_Sum / v_Count;
              End If;
              n_单次数量 := n_Sum / v_Count;
            
              If n_累计数量 >= n_Sum / v_Count * I Then
                n_Count := n_Count + 1;
                Insert Into 药品收发记录
                  (ID, 记录状态, 单据, NO, 序号, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期, 效期, 付数, 填写数量, 实际数量,
                   成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期,
                   灭菌效期, 产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号, 注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间)
                  Select n_Lngid, 记录状态, 单据, NO, n_Count + 1000, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号,
                         生产日期, 效期, 付数, n_剩余数量, n_剩余数量, 成本价, Round(成本价 * n_剩余数量, n_流通金额小数), 扣率, 零售价,
                         Round(零售价 * n_剩余数量, n_流通金额小数), Round(差价 * (实际数量 / n_剩余数量), n_流通金额小数), '复制', 填制人, 填制日期, 配药人,
                         配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期, 灭菌效期, 产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号,
                         注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间
                  From 药品收发记录
                  Where ID = v_收发记录.收发id;
              
                Zl_未审药品记录_Insert(n_Lngid);
              
                Insert Into 输液配药内容 (记录id, 收发id, 数量) Values (v_配药id, n_Lngid, n_剩余数量);
              
                n_剩余数量 := 0;
                Exit;
              Elsif n_累计数量 > (n_Sum / v_Count * (I - 1)) Then
                n_Count    := n_Count + 1;
                n_填写数量 := n_累计数量 - (n_Sum / v_Count * (I - 1)) - (n_单次数量 - n_剩余数量);
                Insert Into 药品收发记录
                  (ID, 记录状态, 单据, NO, 序号, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号, 生产日期, 效期, 付数, 填写数量, 实际数量,
                   成本价, 成本金额, 扣率, 零售价, 零售金额, 差价, 摘要, 填制人, 填制日期, 配药人, 配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期,
                   灭菌效期, 产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号, 注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间)
                  Select n_Lngid, 记录状态, 单据, NO, n_Count + 1000, 库房id, 供药单位id, 入出类别id, 对方部门id, 入出系数, 药品id, 批次, 产地, 批号,
                         生产日期, 效期, 付数, n_填写数量, n_填写数量, 成本价, Round(成本价 * n_填写数量, n_流通金额小数), 扣率, 零售价,
                         Round(零售价 * n_填写数量, n_流通金额小数), Round(差价 * (实际数量 / n_填写数量), n_流通金额小数), '复制', 填制人, 填制日期, 配药人,
                         配药日期, 审核人, 审核日期, 价格id, 费用id, 单量, 频次, 用法, 外观, 灭菌日期, 灭菌效期, 产品合格证, 发药方式, 发药窗口, 领用人, 批准文号, 汇总发药号,
                         注册证号, 库房货位, 商品条码, 内部条码, 核查人, 核查日期, 签到确认人, 签到时间
                  From 药品收发记录
                  Where ID = v_收发记录.收发id;
              
                Zl_未审药品记录_Insert(n_Lngid);
              
                Insert Into 输液配药内容 (记录id, 收发id, 数量) Values (v_配药id, n_Lngid, n_填写数量);
              
                n_剩余数量 := n_剩余数量 - n_填写数量;
              End If;
            End Loop;
          End Loop;
          n_优先级 := Zl_Getfirst(v_配药id, v_医嘱记录.病人科室id);
          Update 输液配药记录 Set 优先级 = n_优先级 Where ID = v_配药id;
        
        End Loop;
      
        For v_收发记录 In c_原始收发记录 Loop
          n_单据 := v_收发记录.单据;
        
          v_No := v_收发记录.No;
          Delete From 药品收发记录 Where ID = v_收发记录.收发id;
        End Loop;
      
        --单个药品或者不予配置的药品默认为0批次
        Select Count(收发id) Into n_Row From 输液配药内容 Where 记录id = v_配药id;
        If (v_Nodosage = 1 Or n_Row = 1) And n_打包药品批次 = 1 Then
          Update 输液配药记录
          Set 配药批次 = 0, 是否打包 = 1
          Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And 操作状态 < 2;
        End If;
        --如果存在“不予配置”属性的药品，也设置为打包
        If v_Nodosage = 1 Then
          Update 输液配药记录
          Set 是否打包 = 1
          Where 医嘱id = v_医嘱记录.相关id And 发送号 = v_医嘱记录.发送号 And 操作状态 < 2;
        End If;
      End If;
    End If;
  End Loop;

  For v_收发记录 In (Select ID From 药品收发记录 Where 序号 < 1000 And 单据 = n_单据 And NO = v_No) Loop
    n_Count := n_Count + 1;
    Update 药品收发记录 Set 序号 = n_Count + 1000, 摘要 = '复制' Where ID = v_收发记录.Id;
  End Loop;

  Update 药品收发记录
  Set 序号 = 序号 - 1000, 摘要 = '医嘱发送'
  Where 摘要 = '复制' And 序号 > 1000 And 单据 = n_单据 And NO = v_No;

  If n_备用批次 = 1 Then
  
    Select Count(a.Id)
    Into n_Sum
    From 输液配药记录 A, 病人医嘱发送 B
    Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And
          a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id And 相关id Is Null) And b.发送时间 Between Trunc(Sysdate) And
          Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And a.执行时间 Between Trunc(Sysdate) And
          Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And 操作状态 < 9;
    If n_Sum <> 0 Then
      b_Change  := False;
      v_时间串1 := Trunc(Sysdate + 1) || ';false\';
    
      Update 输液配药记录
      Set 是否调整批次 = 1
      Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(Sysdate + 1) And
            Trunc(Sysdate + 2) - 1 / 24 / 60 / 60 And 操作状态 < 2;
    End If;
  End If;
  If v_时间串 Is Null Then
    v_时间串 := v_时间串1;
  Else
    v_时间串 := v_时间串 || v_时间串1;
  End If;

  While v_时间串 Is Not Null Loop
    --分解单据ID串
    v_Fields   := Substr(v_时间串, 1, Instr(v_时间串, '\') - 1);
    v_时间值   := Substr(v_Fields, 1, Instr(v_Fields, ';') - 1);
    v_是否改变 := Substr(v_Fields, Instr(v_Fields, ';') + 1);
  
    v_时间串 := Replace('\' || v_时间串, '\' || v_Fields || '\');
  
    If v_是否改变 = 'true' Then
      b_Change := True;
    Else
      b_Change := False;
    End If;
  
    If b_Change = True Then
      Select Count(医嘱id)
      Into n_Cur
      From (Select Distinct a.要求时间, a.医嘱id
             From 医嘱执行时间 A, 输液配药记录 B
             Where a.要求时间 = b.执行时间 And a.医嘱id = b.医嘱id And a.要求时间 Between Trunc(v_时间值) And
                   Trunc(v_时间值 + 1) - 1 / 24 / 60 / 60 And
                   a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id And 相关id Is Null));
    
      Select Count(医嘱id)
      Into n_Sum
      From (Select Distinct a.要求时间, a.医嘱id
             From 医嘱执行时间 A, 输液配药记录 B
             Where a.要求时间 = b.执行时间 And a.医嘱id = b.医嘱id And a.要求时间 Between Trunc(v_时间值 - 1) And
                   Trunc(v_时间值) - 1 / 24 / 60 / 60 And
                   a.医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id And 相关id Is Null));
    
      If n_Cur <> n_Sum Then
        Update 输液配药记录
        Set 是否调整批次 = 1
        Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_时间值) And
              Trunc(v_时间值 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
        b_Change := False;
      End If;
    End If;
  
    If v_保持上次批次 = 1 And b_Change = True Then
      For v_输液单记录 In c_输液单记录 Loop
        Begin
          Select Distinct 配药批次
          Into v_批次
          From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C
          Where a.Id = b.记录id And b.收发id = c.Id And a.医嘱id = v_输液单记录.医嘱id And
                To_Char(a.执行时间, 'hh24:mi:ss') = To_Char(v_输液单记录.执行时间, 'hh24:mi:ss') And
                a.执行时间 Between Trunc(v_输液单记录.执行时间 - 1) And Trunc(v_输液单记录.执行时间) - 1 / 24 / 60 / 60 And Rownum = 1;
        Exception
          When Others Then
            Begin
              Select Distinct 配药批次
              Into v_批次
              From 输液配药记录 A
              Where a.医嘱id = v_输液单记录.医嘱id And To_Char(a.执行时间, 'hh24:mi:ss') = To_Char(v_输液单记录.执行时间, 'hh24:mi:ss') And
                    a.操作状态 <> 12 And a.执行时间 Between Trunc(v_输液单记录.执行时间 - 1) And Trunc(v_输液单记录.执行时间) - 1 / 24 / 60 / 60 And
                    Rownum = 1;
            Exception
              When Others Then
                v_批次 := v_输液单记录.配药批次;
            End;
        End;
      
        Update 输液配药记录
        Set 是否确认调整 = 0, 是否调整批次 = 0
        Where 医嘱id In (Select ID From 病人医嘱记录 Where 病人id = n_病人id) And 执行时间 Between Trunc(v_输液单记录.执行时间) And
              Trunc(v_输液单记录.执行时间 + 1) - 1 / 24 / 60 / 60 And 操作状态 < 2;
      
        If v_输液单记录.配药批次 <> v_批次 Then
          Update 输液配药记录 Set 配药批次 = v_批次 Where ID = v_输液单记录.Id;
          Select Nvl(Max(打包), 0) Into n_打包 From 配药工作批次 Where 批次 = v_批次 And 配置中心id = 部门id_In;
          If n_打包 <> 0 Then
            Update 输液配药记录 Set 是否打包 = n_打包 Where ID = v_输液单记录.Id;
          Else
            Select Nvl(Max(打包), 0)
            Into n_打包
            From 配药工作批次
            Where 批次 = v_输液单记录.配药批次 And 配置中心id = 部门id_In;
          
            If n_打包 <> 0 Then
              Update 输液配药记录 Set 是否打包 = 0 Where ID = v_输液单记录.Id;
            End If;
          End If;
        End If;
      End Loop;
    End If;
  
    If n_自动排批 = 1 And (b_Change = False Or v_保持上次批次 = 0) Then
      For v_输液单记录 In c_输液单记录 Loop
        v_批次 := Zl_Getpivaworkbatch(v_输液单记录.执行时间, v_输液单记录.发送时间);
        Update 输液配药记录 Set 配药批次 = v_批次 Where ID = v_输液单记录.Id;
      End Loop;
      Zl_输液配药记录_自动排批(n_病人id, n_科室id, 部门id_In, v_时间值);
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]病人' || v_当前病人 || '在输液配置中心有被锁定的输液单，发送失败！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_核查;
/

--125588:秦龙,2018-05-17,处理库存实际数量为0的数据
Create Or Replace Procedure Zl_材料收发记录_Adjust
(
  调价id_In   In Number, --调价记录的ID
  定价_In     In Number := 0, --是否转为定价销售（更新材料特性、收费细目中的变价）
  材料id_In   In Number := 0, --当不为0时表示是成本价调价，不处理售价相关内容
  Billinfo_In In Varchar2 := Null --用于时价卫材按批次调价。格式:"批次1,现价1|批次2,现价2|....."
) As
  n_入出类别id 药品收发记录.入出类别id%Type; --入出类别
  v_调价单据号 药品收发记录.No%Type; --调价单号
  d_生效日期   Date; --调价生效时间
  n_执行调价   Number(1); --调价时刻到了
  n_实价材料   Number(1); --时价药品
  n_收费细目id Number(18); --收费细目ID
  d_审核日期   药品收发记录.审核日期%Type;
  n_零售金额   药品库存.实际金额%Type;
  n_零售价     药品库存.零售价%Type;
  n_序号       Integer(8);
  v_Infotmp    Varchar2(4000);
  v_Fields     Varchar2(4000);
  n_批次       Number(18);
  n_现价       收费价目.现价%Type;
  n_原价       收费价目.原价%Type;
  n_收发id     药品收发记录.Id%Type;
  n_时价分批   Number(1);
  v_Lngid      药品收发记录.Id%Type; --收发ID
  n_价格id     收费价目.Id%Type;

  Cursor c_Price --普通调价
  Is
    Select 1 记录状态, 13 单据, v_调价单据号 NO, Rownum 序号, n_入出类别id 入出类别id, m.材料id 药品id, s.批次 批次, Null 批号, s.效期,
           Decode(s.上次产地, Null, q.产地, s.上次产地) 产地, 1 付数, s.实际数量 填写数量, 0 实际数量, a.原价 成本价, 0 成本金额, a.现价 零售价, 0 扣率,
           Nvl(s.零售价, 0) As 库存零售价, s.实际金额 As 库存金额, s.实际差价 As 库存差价, '卫材调价' 摘要, User 填制人, Sysdate 填制日期, s.库房id 库房id,
           1 入出系数, a.Id 价格id, s.上次生产日期, s.灭菌效期, s.批准文号, s.上次供应商id,
           Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, a.原价, Nvl(s.实际金额, 0) / s.实际数量)) As 原售价
    From 药品库存 S, 材料特性 M, 收费价目 A, 收费项目目录 Q
    Where s.药品id = m.材料id And m.材料id = q.Id And m.材料id = a.收费细目id And s.性质 = 1 And a.变动原因 = 0 And a.Id = 调价id_In And
          a.执行日期 <= Sysdate;

  Cursor c_时价按批次调价 --时价卫材按批次调价
  Is
    Select 1 记录状态, 13 单据, v_调价单据号 NO, n_序号 + Rownum 序号, n_入出类别id 入出类别id, s.药品id 药品id, s.批次 批次, Null 批号, s.效期,
           Decode(s.上次产地, Null, b.产地, s.上次产地) 产地, 1 付数, Nvl(s.实际数量, 0) 填写数量, 0 实际数量, a.原价 成本价, 0 成本金额, n_现价 零售价, 0 扣率,
           '卫材调价' 摘要, User 填制人, Sysdate 填制日期, s.库房id 库房id, 1 入出系数, a.Id 价格id, Nvl(b.是否变价, 0) As 时价, s.实际金额 As 库存金额,
           s.实际差价 As 库存差价, Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, a.原价, Nvl(s.实际金额, 0) / s.实际数量)) As 原售价
    From 药品库存 S, 材料特性 M, 收费价目 A, 收费项目目录 B
    Where s.药品id = m.材料id And m.材料id = a.收费细目id And a.收费细目id = b.Id And s.性质 = 1 And a.变动原因 = 0 And a.Id = 调价id_In And
          a.执行日期 <= Sysdate And Nvl(s.批次, 0) = n_批次;
Begin

  If 材料id_In <> 0 Then
    --成本价调价
    Zl_材料收发记录_成本价调价(材料id_In);
    Return;
  End If;

  --取入出类别ID
  Select 类别id Into n_入出类别id From 药品单据性质 Where 单据 = 13;

  --取序列
  Select Nextno(147) Into v_调价单据号 From Dual;
  --取调价记录生效日期
  Select 收费细目id, 执行日期 Into n_收费细目id, d_生效日期 From 收费价目 Where ID = 调价id_In;
  --取该材料是否是时价药品
  Select Nvl(是否变价, 0) Into n_实价材料 From 收费项目目录 Where ID = n_收费细目id;

  If Sysdate >= d_生效日期 Then
    n_执行调价 := 1;
  Else
    n_执行调价 := 0;
  End If;

  If n_执行调价 = 1 Then
    d_审核日期 := Sysdate;
    --普通调价处理
    If Billinfo_In = '' Or Billinfo_In Is Null Then
      --非时价药品调价
      For c_调价 In c_Price Loop
        n_价格id := c_调价.价格id;
        /*If Nvl(c_调价.填写数量, 0) = 0 And Nvl(c_调价.库存金额, 0) = 0 And Nvl(c_调价.库存差价, 0) = 0 Then
          Null;
        Elsif Nvl(c_调价.填写数量, 0) = 0 And (Nvl(c_调价.库存金额, 0) <> 0 Or Nvl(c_调价.库存差价, 0) <> 0) Then
          --数量=0 金额或差价<>0时只更新库存表中对应的零售价,并产生售价修正数据但是金额差=0，只记录最新售价，金额差和差价差不填数据

        
        
          --产生调价影响记录
          Select 药品收发记录_Id.Nextval Into v_Lngid From Dual;
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 摘要, 填制人, 填制日期,
             库房id, 入出系数, 价格id, 审核人, 审核日期, 生产日期, 灭菌效期, 批准文号, 供药单位id, 单量, 频次)
          Values
            (v_Lngid, c_调价.记录状态, c_调价.单据, c_调价.No, c_调价.序号, c_调价.入出类别id, c_调价.药品id, c_调价.批次, c_调价.批号, c_调价.效期, c_调价.产地,
             c_调价.付数, c_调价.填写数量, c_调价.实际数量, Decode(n_实价材料, 1, c_调价.原售价, c_调价.成本价), c_调价.成本金额, c_调价.零售价, c_调价.扣率, c_调价.摘要,
             c_调价.填制人, c_调价.填制日期, c_调价.库房id, c_调价.入出系数, c_调价.价格id, User, d_审核日期, c_调价.上次生产日期, c_调价.灭菌效期, c_调价.批准文号,
             c_调价.上次供应商id, c_调价.库存金额, c_调价.库存差价);
        
          Zl_未审药品记录_Insert(v_Lngid);
          --更新材料库存 ，只有时价卫材才更新零售价
          Update 药品库存
          Set 零售价 = Decode(n_实价材料, 1, Decode(Nvl(c_调价.批次, 0), 0, Null, c_调价.零售价), Null)
          Where 库房id = c_调价.库房id And 药品id = c_调价.药品id And 性质 = 1 And Nvl(批次, 0) = Nvl(c_调价.批次, 0);
        
          Zl_未审药品记录_Delete(v_Lngid);
        Else*/
        If n_实价材料 = 1 Then
          If c_调价.库存零售价 = 0 Then
            n_零售价 := c_调价.原售价;
          Else
            n_零售价 := c_调价.库存零售价;
          End If;
        Else
          n_零售价 := c_调价.成本价;
        End If;
        n_零售金额 := Round((c_调价.零售价 - n_零售价) * c_调价.填写数量, 2);
      
        --产生调价影响记录
        Select 药品收发记录_Id.Nextval Into v_Lngid From Dual;
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
           填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 生产日期, 灭菌效期, 批准文号, 供药单位id, 单量, 频次)
        Values
          (v_Lngid, c_调价.记录状态, c_调价.单据, c_调价.No, c_调价.序号, c_调价.入出类别id, c_调价.药品id, c_调价.批次, c_调价.批号, c_调价.效期, c_调价.产地,
           c_调价.付数, c_调价.填写数量, c_调价.实际数量, Decode(n_实价材料, 1, c_调价.原售价, c_调价.成本价), c_调价.成本金额, c_调价.零售价, c_调价.扣率, n_零售金额,
           n_零售金额, c_调价.摘要, c_调价.填制人, c_调价.填制日期, c_调价.库房id, c_调价.入出系数, c_调价.价格id, User, d_审核日期, c_调价.上次生产日期, c_调价.灭菌效期,
           c_调价.批准文号, c_调价.上次供应商id, c_调价.库存金额, c_调价.库存差价);
      
        Zl_未审药品记录_Insert(v_Lngid);
        --更新材料库存
        Update 药品库存
        Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_零售金额,
            零售价 = Decode(n_实价材料, 1, Decode(Nvl(c_调价.批次, 0), 0, Null, c_调价.零售价), Null)
        Where 库房id = c_调价.库房id And 药品id = c_调价.药品id And 性质 = 1 And Nvl(批次, 0) = Nvl(c_调价.批次, 0);
      
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 效期, 灭菌效期, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 批准文号, 零售价)
          Values
            (c_调价.库房id, c_调价.药品id, c_调价.批次, 1, 0, 0, n_零售金额, n_零售金额, c_调价.效期, c_调价. 灭菌效期, c_调价.上次供应商id, c_调价.成本价,
             c_调价.批号, c_调价.上次生产日期, c_调价.产地, c_调价.批准文号,
             Decode(n_实价材料, 1, Decode(Nvl(c_调价.批次, 0), 0, Null, c_调价.零售价), Null));
        End If;
      
        Zl_未审药品记录_Delete(v_Lngid);
        --End If;
      End Loop;
    
      --消息处理
      b_Message.Zlhis_Drug_011(n_价格id, 0);
    Else
      --时价分批调价处理
      n_序号 := 0;
      --时价药品按批次调价
      v_Infotmp := Billinfo_In || '|';
      While v_Infotmp Is Not Null Loop
        --分解单据ID串
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        n_批次    := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        n_现价    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        For v_时价按批次调价 In c_时价按批次调价 Loop
          If v_时价按批次调价.填写数量 <> 0 Then
            n_原价 := Nvl(v_时价按批次调价.库存金额, 0) / v_时价按批次调价.填写数量;
          Else
            n_原价 := v_时价按批次调价.成本价;
          End If;
        
          Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
          /*If Nvl(v_时价按批次调价.填写数量, 0) = 0 And Nvl(v_时价按批次调价.库存金额, 0) = 0 And Nvl(v_时价按批次调价.库存差价, 0) = 0 Then
            Null;
            n_价格id := Null;
          Elsif Nvl(v_时价按批次调价.填写数量, 0) = 0 And (Nvl(v_时价按批次调价.库存金额, 0) <> 0 Or Nvl(v_时价按批次调价.库存差价, 0) <> 0) Then
            --数量=0 金额或差价<>0时只更新库存表中对应的零售价,并产生售价修正数据但是金额差=0，只记录最新售价，金额差和差价差不填数据

          
          
            --产生调价影响记录
            Insert Into 药品收发记录
              (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 摘要, 填制人, 填制日期,
               库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次)
            Values
              (n_收发id, v_时价按批次调价.记录状态, v_时价按批次调价.单据, v_时价按批次调价.No, v_时价按批次调价.序号, v_时价按批次调价.入出类别id, v_时价按批次调价.药品id,
               v_时价按批次调价.批次, v_时价按批次调价.批号, v_时价按批次调价.效期, v_时价按批次调价.产地, v_时价按批次调价.付数, v_时价按批次调价.填写数量, v_时价按批次调价.实际数量,
               Decode(n_实价材料, 1, v_时价按批次调价.原售价, v_时价按批次调价.成本价), v_时价按批次调价.成本金额, v_时价按批次调价.零售价, v_时价按批次调价.扣率,
               v_时价按批次调价.摘要, v_时价按批次调价.填制人, v_时价按批次调价.填制日期, v_时价按批次调价.库房id, v_时价按批次调价.入出系数, v_时价按批次调价.价格id, User, d_审核日期,
               v_时价按批次调价.库存金额, v_时价按批次调价.库存差价);
            n_序号 := n_序号 + 1;
          
            Zl_未审药品记录_Insert(n_收发id);
            --处理库存
            --更新库存零售价,只有时价分批药品才能更新零售价字段
            Update 药品库存
            Set 零售价 = Decode(v_时价按批次调价.时价, 1, Decode(Nvl(v_时价按批次调价.批次, 0), 0, Null, v_时价按批次调价.零售价), Null)
            Where 库房id = v_时价按批次调价.库房id And 药品id = v_时价按批次调价.药品id And 性质 = 1 And Nvl(批次, 0) = Nvl(v_时价按批次调价.批次, 0);
          
            Zl_未审药品记录_Delete(n_收发id);
          
            n_价格id := n_收发id;
          Else*/
          n_零售价   := v_时价按批次调价.原售价;
          n_零售金额 := Round((n_现价 - n_零售价) * v_时价按批次调价.填写数量, 2);
          --产生调价影响记录
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
             填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次)
          Values
            (n_收发id, v_时价按批次调价.记录状态, v_时价按批次调价.单据, v_时价按批次调价.No, v_时价按批次调价.序号, v_时价按批次调价.入出类别id, v_时价按批次调价.药品id,
             v_时价按批次调价.批次, v_时价按批次调价.批号, v_时价按批次调价.效期, v_时价按批次调价.产地, v_时价按批次调价.付数, v_时价按批次调价.填写数量, v_时价按批次调价.实际数量,
             Decode(n_实价材料, 1, v_时价按批次调价.原售价, v_时价按批次调价.成本价), v_时价按批次调价.成本金额, v_时价按批次调价.零售价, v_时价按批次调价.扣率, n_零售金额,
             n_零售金额, v_时价按批次调价.摘要, v_时价按批次调价.填制人, v_时价按批次调价.填制日期, v_时价按批次调价.库房id, v_时价按批次调价.入出系数, v_时价按批次调价.价格id, User,
             d_审核日期, v_时价按批次调价.库存金额, v_时价按批次调价.库存差价);
          n_序号 := n_序号 + 1;
        
          Zl_未审药品记录_Insert(n_收发id);
          --处理库存
          If v_时价按批次调价.时价 = 1 And Nvl(v_时价按批次调价.批次, 0) > 0 Then
            n_时价分批 := 1;
          Else
            n_时价分批 := 0;
          End If;
        
          If Nvl(v_时价按批次调价.批次, 0) = 0 Then
            Update 药品库存
            Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_零售金额
            Where 库房id = v_时价按批次调价.库房id And 药品id = v_时价按批次调价.药品id And 性质 = 1 And (批次 Is Null Or 批次 = 0);
          Else
            Update 药品库存
            Set 实际金额 = Nvl(实际金额, 0) + n_零售金额, 实际差价 = Nvl(实际差价, 0) + n_零售金额, 零售价 = Decode(n_时价分批, 1, v_时价按批次调价.零售价, 零售价)
            Where 库房id = v_时价按批次调价.库房id And 药品id = v_时价按批次调价.药品id And 性质 = 1 And 批次 = v_时价按批次调价.批次;
          End If;
        
          If Sql%RowCount = 0 Then
            Insert Into 药品库存
              (库房id, 药品id, 批次, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 零售价)
            Values
              (v_时价按批次调价.库房id, v_时价按批次调价.药品id, v_时价按批次调价.批次, 1, 0, 0, n_零售金额, n_零售金额,
               Decode(n_时价分批, 1, v_时价按批次调价.零售价, Null));
          End If;
        
          Zl_未审药品记录_Delete(n_收发id);
        
          n_价格id := n_收发id;
          --End If;
        
          --消息处理
          If n_价格id Is Not Null Then
            b_Message.Zlhis_Drug_011(n_价格id, 1);
          End If;
        End Loop;
      End Loop;
    End If;
  
    Update 药品收发记录 Set 审核人 = User, 审核日期 = Sysdate Where 价格id = 调价id_In;
    Update 收费价目 Set 变动原因 = 1 Where ID = 调价id_In;
  
    --更新药品目录、收费细目中的变价
    If 定价_In = 1 Then
      Update 收费项目目录 Set 是否变价 = 0 Where ID = n_收费细目id;
    End If;
    --成本价调价
    Zl_材料收发记录_成本价调价(n_收费细目id);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料收发记录_Adjust;
/

--125925:胡俊勇,2018-05-17,移动护理接口封装
Create Or Replace Procedure Zl_Third_Advicecheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：病人医嘱核对/取消核对，数据写入
  --入参：xml_in
  --<IN>
  -- <TYPE>1</TYPE>   --操作类型：1、核对；0、取消核对
  -- <YZID>1162695</YZID>   --医嘱id
  -- <FSH>202704</FSH>   --发送号    
  -- <ZXSJ>2017-12-05 16:26:54</ZXSJ>   --执行时间

  --以下节点取消核对时传空
  -- <HDSJ>2017-12-05 10:00:00</HDSJ>   --核对时间  
  -- <HDR></HDR>   --核对人   
  --</IN>

  --出参：Xml_Out
  --<OUTPUT>
  --   <RESULT>true</RESULT>
  --</OUTPUT>

  --失败：
  --<OUTPUT>
  --   <RESULT>false</RESULT>
  --   <ERROR>
  --     <MSG>详细错误提示</MSG>
  --   </ERROR>
  --</OUTPUT>

  n_Type     Number;
  n_医嘱id   病人医嘱记录.Id%Type;
  n_发送号   病人医嘱发送.发送号%Type;
  d_执行时间 病人医嘱执行.要求时间%Type;
  d_核对时间 病人医嘱执行.要求时间%Type;
  v_核对人   病人医嘱执行.核对人%Type;
Begin
  Select Extractvalue(Value(A), 'IN/TYPE'), Extractvalue(Value(A), 'IN/YZID') As 医嘱id,
         Extractvalue(Value(A), 'IN/FSH') As 发送号,
         To_Date(Extractvalue(Value(A), 'IN/ZXSJ'), 'yyyy-mm-dd hh24:mi:ss') As 执行时间,
         To_Date(Extractvalue(Value(A), 'IN/HDSJ'), 'yyyy-mm-dd hh24:mi:ss') As d_核对时间,
         Extractvalue(Value(A), 'IN/HDR') As 核对人
  Into n_Type, n_医嘱id, n_发送号, d_执行时间, d_核对时间, v_核对人
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If n_Type = 1 Then
    Zl_病人医嘱核对_Insert(n_医嘱id, n_发送号, v_核对人, d_执行时间, d_核对时间);  
  Else
    Zl_病人医嘱核对_Delete(n_医嘱id, n_发送号, d_执行时间);  
  End If;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Advicecheck;
/

--124431:李南春,2018-05-17,挂号序号占用后重新取序号
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
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, NO;
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
  n_占用         Number;
  d_发生时间     门诊费用记录.发生时间%Type;
  
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

  IF Nvl(号序_In, 0) <> 0 then
    Begin
      Select 1 Into n_占用 From 临床出诊序号控制
      Where 记录ID = 出诊记录id_In And 序号 = 号序_In And (挂号状态 In (1,2,4) Or 挂号状态 in (3, 5) And (操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名));
    Exception
      When Others Then
        n_占用 := 0;
    End;
  End IF;
  IF Nvl(n_占用, 0) = 1 And 序号_In = 1 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 Then
    Begin
      Select 序号, 开始时间
      Into n_序号, d_发生时间
      From (Select 序号, 开始时间 From 临床出诊序号控制 Where 序号 > 号序_In And Nvl(挂号状态, 0) = 0 Order By 序号)
      Where Rownum < 2;
    Exception
      When Others Then
        n_序号 := Null;
        d_发生时间 := 发生时间_In;
    End;
  Else
    n_序号 := 号序_In;
    d_发生时间 := 发生时间_In;
  End IF;

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
      Where 记录id = 出诊记录id_In And 开始时间 = d_发生时间 And Rownum < 2;
    Exception
      When Others Then
        n_序号 := Null;
    End;
  End If;

  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    Begin
      Select Max(To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 临床出诊序号控制
      Where 记录id = 出诊记录id_In And Nvl(数量, 0) <> 0;
    
      n_追加号 := Case Sign(d_发生时间 - d_最大序号时间)
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
  d_时段时间 := d_发生时间;

  If 序号_In = 1 And n_分时段 > 0 Then
    If Nvl(n_序号控制, 0) = 1 Then
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := d_发生时间;
          n_时段限号 := 0;
          n_时段限约 := 0;
      End;
    Else
      --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
      Begin
        Select Nvl(序号, 0),
               To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               数量, Decode(Nvl(是否预约, 0), 0, 0, 数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 临床出诊序号控制
        Where 记录id = 出诊记录id_In And 序号 = n_序号 And 预约顺序号 Is Null;
      Exception
        When Others Then
          n_时段序号 := -1;
          n_分时段   := 0;
          d_时段时间 := d_发生时间;
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
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(d_发生时间), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(d_发生时间), 'yyyy-mm-dd') || '已达到最大限约数！';
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
          v_Err_Msg := '号别' || 号别_In || To_Char(d_发生时间, 'yyyy-mm-dd') || '已达到最大限制数！';
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
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(d_发生时间, 'yyyy-mm-dd') || '已达到最大限制数！';
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
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(d_发生时间, 'yyyy-mm-dd') || '已达到最大限制数！';
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
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(d_发生时间, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(d_发生时间, 'yyyy-mm-dd') || '已达到最大限制数！';
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
      d_序号时间 := d_发生时间;
    Else
      d_序号时间 := Trunc(d_发生时间);
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
                    Select 出诊记录id_In, n_序号, d_发生时间, d_发生时间, 1, Decode(预约挂号_In, 1, 1, 0), Decode(预约挂号_In, 1, 2, 1),
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
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, d_发生时间, 登记时间_In, 保险大类id_In,
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
      Update 病人信息 Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
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
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, d_发生时间, Decode(Nvl(预约挂号_In, 0), 1, d_发生时间, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 出诊记录id_In, 收费单_In);
  
    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
      Update 病人挂号记录
      Set 预约 = 1, 预约时间 = d_发生时间, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In
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
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, d_发生时间);
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

--124431:李南春,2018-05-17,挂号序号占用后重新取序号
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
           Max(Decode(记录性质, 1, ID, 0) * Decode(记录状态, 2, 0, 1)) As 原预交id
    From 病人预交记录
    Where 记录性质 In (1, 11) And 病人id In (Select Column_Value From Table(f_Num2list(v_冲预交病人ids))) And Nvl(预交类别, 2) = 1
     Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0
    Group By NO, 病人id
    Order By Decode(病人id, Nvl(v_病人id, 0), 0, 1), 结帐id, NO;
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
  n_占用           Number;
  d_发生时间       门诊费用记录.发生时间%Type;
  
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
  
  --如果当前序号不为0，则要检查需要是否被占用，如果被占用，按顺序去下一个序号及发生时间
  IF Nvl(号序_In, 0) <> 0 then
    Begin
      IF 退号重用_In = 1 Then
        Select 1
        Into n_占用
        From 挂号序号状态
        Where 号码 = 号别_In And Trunc(日期) = Trunc(发生时间_In) And 序号 = 号序_In And
              (状态 In (1, 2) Or 状态 In (3, 5) And (操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名));
      Else
        Select 1
        Into n_占用
        From 挂号序号状态
        Where 号码 = 号别_In And Trunc(日期) = Trunc(发生时间_In) And 序号 = 号序_In And
              (状态 In (1, 2, 4) Or 状态 In (3, 5) And (操作员姓名_In <> v_序号操作员 Or v_机器名 <> v_序号机器名));
      End IF;
    Exception
      When Others Then
        n_占用 := 0;
    End;
  End IF;
  IF Nvl(n_占用, 0) = 1 And 序号_In = 1 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 Then
    Begin
      If Nvl(n_计划id, 0) = 0 Then
        Select 序号, 时段时间
        Into n_序号, d_发生时间
        From (Select Nvl(序号, 0) As 序号,
                      To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'),
                               'yyyy-mm-dd hh24:mi:ss') As 时段时间
               From 挂号安排时段
               Where 安排id = n_安排id And 星期 = v_星期 And 序号 > Nvl(号序_In, 0) And
                     序号 Not In (Select 序号
                                From 挂号序号状态
                                Where 号码 = 号别_In And 状态 <> 0 And 状态 <> Decode(退号重用_In, 1, 4, 0) And
                                      日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60)
               Order By 序号)
        Where Rownum < 2;
      Else
        Select 序号, 时段时间
        Into n_序号, d_发生时间
        From (Select Nvl(序号, 0) As 序号,
                      To_Date(To_Char(发生时间_In, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'),
                               'yyyy-mm-dd hh24:mi:ss') As 时段时间
               From 挂号计划时段
               Where 计划id = n_计划id And 星期 = v_星期 And 序号 > Nvl(号序_In, 0) And
                     序号 Not In (Select 序号
                                From 挂号序号状态
                                Where 号码 = 号别_In And 状态 <> 0 And 状态 <> Decode(退号重用_In, 1, 4, 0) And
                                      日期 Between Trunc(发生时间_In) And Trunc(发生时间_In) + 1 - 1 / 24 / 60 / 60)
               Order By 序号)
        Where Rownum < 2;
      End IF;
    Exception
      When Others Then
        n_序号 := Null;
        d_发生时间 := 发生时间_In;
    End;
  Else
    n_序号 := 号序_In;
    d_发生时间 := 发生时间_In;
  End IF;
  
  --分时段挂号时判断是否是过期挂号 也就是追加号的情况 目前只针对专家号分时段进行处理
  If Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 And Nvl(n_序号控制, 0) = 1 And 号序_In Is Null And Nvl(操作类型_In, 0) = 0 Then
    --发生时间_in>Sysdate 发生时间>最大的时段时间--号序_in is null
    Begin
      Select Max(To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_最大序号时间
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And Nvl(限制数量, 0) <> 0;
      n_追加号 := Case Sign(d_发生时间 - d_最大序号时间)
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
  d_时段时间 := d_发生时间;

  If 序号_In = 1 And Nvl(预约挂号_In, 0) = 0 And n_分时段 > 0 Then
    --挂号时检查 是否分了时段,分了时段,把时段的限制条件给取出来
    Begin
      Select Nvl(序号, 0),
             To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
      Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
      From 挂号安排时段
      Where 安排id = n_安排id And 星期 = v_星期 And
            (序号, 安排id, 星期) In (Select Nvl(Max(序号), -1), 安排id, 星期
                               From 挂号安排时段
                               Where 安排id = n_安排id And 星期 = v_星期 And
                                     Decode(操作类型_In + n_追加号, 0, To_Char(d_发生时间, 'hh24:mi'),
                                            To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By 安排id, 星期);
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := d_发生时间;
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
               To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号安排时段 C
        Where 安排id = n_安排id And 星期 = v_星期 And
              (序号, 安排id, 星期) In
              (Select Nvl(Max(c.序号), -1), 安排id, 星期
               From 挂号安排时段 C
               Where 安排id = n_安排id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(d_发生时间, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 安排id, 星期);
      Else
        --有计划生效取计划
        --没生效，代表是从挂号计划时段查询
        Select Nvl(序号, -1),
               To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || ' ' || To_Char(c.开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As 时段时间,
               限制数量, Decode(Nvl(是否预约, 0), 0, 0, 限制数量)
        Into n_时段序号, d_时段时间, n_时段限号, n_时段限约
        From 挂号计划时段 C
        Where 计划id = n_计划id And 星期 = v_星期 And
              (序号, 计划id, 星期) In
              (Select Nvl(Max(c.序号), -1), 计划id, 星期
               From 挂号计划时段 C
               Where 计划id = n_计划id And c.星期 = v_星期 And
                     Decode(操作类型_In, 1, To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(d_发生时间, 'hh24:mi')) =
                     To_Char(Nvl(c.开始时间, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By 计划id, 星期);
      End If;
    Exception
      When Others Then
        n_时段序号 := -1;
        n_分时段   := 0;
        d_时段时间 := d_发生时间;
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
          Where 号码 = 号别_In And 日期 Between Trunc(d_发生时间) And Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(Nvl(序号, 0)), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Nvl(预约, 0))
          Into n_已用序号, n_已用数量, n_已约数
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(d_发生时间) And Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
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
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(d_发生时间), 'yyyy-mm-dd ') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约时检查
          If n_限约数 = 0 Then
            n_限约数 := n_限号数;
          End If;
        
          If n_限约数 <= n_已约数 And n_限约数 > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(Trunc(d_发生时间), 'yyyy-mm-dd') || '已达到最大限约数！';
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
          Where a.号码 = 号别_In And 日期 Between Trunc(d_发生时间) And Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And
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
          v_Err_Msg := '号别' || 号别_In || To_Char(d_发生时间, 'yyyy-mm-dd') || '已达到最大限制数！';
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
          Where 号码 = 号别_In And 日期 Between Trunc(d_发生时间) And Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 状态 Not In (4, 5);
        Else
          Select Max(序号), Sum(Decode(序号, Nvl(号序_In, 0), 0, 1)), Sum(Decode(Sign(日期 - d_时段时间), 0, 1, 0))
          Into n_已用序号, n_已挂数, n_已用数量
          From 挂号序号状态
          Where 号码 = 号别_In And 日期 Between Trunc(d_发生时间) And Trunc(d_发生时间) + 1 - 1 / 24 / 60 / 60 And 状态 <> 5;
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
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(d_发生时间, 'yyyy-mm-dd') || '已达到最大限制数！';
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
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(d_发生时间, 'yyyy-mm-dd') || '已达到最大限制数！';
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
        Where 日期 = Trunc(d_发生时间) And 号码 = 号别_In;
      Exception
        When Others Then
          n_已用数量 := 0;
          n_已约数   := 0;
      End;
      If Nvl(操作类型_In, 0) = 0 Then
        If Nvl(预约挂号_In, 0) = 0 Then
          --挂号
          If Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0 Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(d_发生时间, 'yyyy-mm-dd') || '已达到最大限制数！';
            Raise Err_Item;
          End If;
        Else
          --预约
          If (Nvl(n_限约数, 0) > 0 And Nvl(n_已约数, 0) >= Nvl(n_限约数, 0)) Or
             (Nvl(n_已用数量, 0) >= Nvl(n_限号数, 0) And Nvl(n_限号数, 0) > 0) Then
            v_Err_Msg := '号别' || 号别_In || '在' || To_Char(d_发生时间, 'yyyy-mm-dd') || '已达到最大限制数！';
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
      d_序号时间 := d_发生时间;
    Else
      d_序号时间 := Trunc(d_发生时间);
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
     操作员姓名_In, Decode(预约挂号_In, 1, 操作员姓名_In, Null), 执行部门id_In, 医生姓名_In, 操作员编号_In, 操作员姓名_In, d_发生时间, 登记时间_In, 保险大类id_In,
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
      Update 病人信息 Set 就诊时间 = d_发生时间, 就诊状态 = 1, 就诊诊室 = 诊室_In Where 病人id = 病人id_In;
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
       Null, 执行部门id_In, 医生姓名_In, 0, Null, 登记时间_In, d_发生时间, Decode(Nvl(预约挂号_In, 0), 1, d_发生时间, Null), 操作员编号_In,
       操作员姓名_In, 复诊_In, n_序号, 社区_In, Decode(预约接收_In, 1, 1, 0), 预约方式_In, 摘要_In, 交易流水号_In, 交易说明_In, 合作单位_In,
       Decode(Nvl(预约挂号_In, 0), 0, 登记时间_In, Null), Decode(Nvl(预约挂号_In, 0), 0, 操作员姓名_In, Null),
       Decode(Nvl(预约挂号_In, 0), 1, 操作员姓名_In, Null), Decode(Nvl(预约挂号_In, 0), 1, 操作员编号_In, Null),
       Decode(Nvl(险类_In, 0), 0, Null, 险类_In), v_付款方式, 收费单_In);
    If Nvl(预约挂号_In, 0) = 0 And 预约方式_In Is Not Null Then
      Update 病人挂号记录
      Set 预约 = 1, 预约时间 = d_发生时间, 预约操作员 = 操作员姓名_In, 预约操作员编号 = 操作员编号_In
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
        d_排队时间 := Zl_Get_Queuedate(n_挂号id, 号别_In, n_序号, d_发生时间);
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

--125588:秦龙,2018-05-17,处理卫材库存实际数量为0的数据
Create Or Replace Procedure Zl_材料收发记录_成本价调价(材料id_In In 药品收发记录.药品id%Type) As
  v_No         药品收发记录.No%Type;
  v_应付id     应付记录.Id%Type; --应付记录的ID 
  v_应付单据号 应付记录.No%Type;
  d_调价时间   Date;
  n_序号       Number(8);
  n_库房id     药品收发记录.库房id%Type;
  n_入出类别id 药品收发记录.入出类别id%Type;
  n_入出系数   药品收发记录.入出系数%Type;
  n_收发id     药品收发记录.Id%Type;
  n_调整额     药品收发记录.零售金额%Type;
  n_原成本价   药品收发记录.成本价%Type;
  n_新成本价   药品收发记录.成本价%Type;
  n_平均成本价 药品库存.平均成本价%Type;
  v_调价id     成本价调价信息.Id%Type;
  v_调价汇总号 成本价调价信息.调价汇总号%Type;
  n_Count      Number(1) := 0;

  Cursor c_Stock Is --当前库存 
    Select 上次供应商id, a.库房id, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, a.上次批号, a.效期, a.上次产地, a.灭菌效期,
           Decode(Sign(Nvl(a.批次, 0)), 1, a.上次采购价, a.平均成本价) As 原成本价
    From 药品库存 A
    Where a.性质 = 1 And Nvl(a.实际数量, 0) <> 0 And a.药品id = 材料id_In
    Order By a.库房id;

  v_Stock c_Stock%RowType;
Begin
  d_调价时间 := Sysdate;
  n_库房id   := 0;

  --判断是否存在无库存调价 
  Begin
    Select ID, 新成本价, 调价汇总号
    Into v_调价id, n_新成本价, v_调价汇总号
    From 成本价调价信息
    Where 执行日期 Is Null And Nvl(库房id, 0) = 0 And 药品id = 材料id_In;
  Exception
    When Others Then
      v_调价id   := 0;
      n_新成本价 := Null;
  End;

  --无库存调价 
  If v_调价id > 0 Then
    --根据当前库存重新产生调价信息 
    For v_Stock In c_Stock Loop
      Zl_材料成本调价_Insert(v_Stock.上次供应商id, v_Stock.库房id, v_Stock.材料id, v_Stock.批次, v_Stock.上次批号, v_Stock.原成本价, n_新成本价,
                       Null, Null, 0, 0, v_调价汇总号);
      n_Count := n_Count + 1;
    End Loop;
  
    If n_Count > 0 Then
      --如果当前有库存记录，则删除无库存调价记录 
      Delete 成本价调价信息 Where ID = v_调价id;
    Else
      Update 成本价调价信息 Set 执行日期 = d_调价时间 Where ID = v_调价id;
    
      Update 材料特性 Set 成本价 = n_新成本价 Where 材料id = 材料id_In And 成本价 <> n_新成本价;
    End If;
  End If;

  --取库存差价调整的入出类别ID 
  Select b.Id, b.系数
  Into n_入出类别id, n_入出系数
  From 药品单据性质 A, 药品入出类别 B
  Where a.类别id = b.Id And a.单据 = 33 And Rownum < 2;

  For c_成本调整 In (Select a.库房id, a.药品id As 材料id, Nvl(a.批次, 0) 批次, a.上次供应商id, a.实际数量, a.实际金额, a.实际差价, a.上次产地 As 产地,
                        a.上次批号 As 批号, a.灭菌效期, a.效期, a.上次生产日期 As 生产日期, a.批准文号, Nvl(a.平均成本价, 0) As 原成本价, b.新成本价, b.发票号,
                        b.发票日期, b.发票金额, Nvl(a.上次采购价, 0) As 上次采购价, b.Id As 调价id
                 From 药品库存 A, 成本价调价信息 B
                 Where a.药品id = b.药品id And Nvl(a.上次供应商id, 0) = Nvl(b.供药单位id, 0) And a.库房id = b.库房id And
                       Nvl(a.批次, 0) = Nvl(b.批次, 0) And a.性质 = 1 And b.执行日期 Is Null And a.药品id = 材料id_In
                 Order By a.库房id) Loop
    If n_库房id <> c_成本调整.库房id Then
      n_序号   := 1;
      n_库房id := c_成本调整.库房id;
      v_No     := Nextno(71, n_库房id);
    Else
      n_序号 := n_序号 + 1;
    End If;
  
    Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
    /*    If Nvl(c_成本调整.实际数量, 0) = 0 And Nvl(c_成本调整.实际金额, 0) = 0 And Nvl(c_成本调整.实际差价, 0) = 0 Then
      --数量,金额、差价都为0，则表示数据是填单下可用数量出库产生的单据，此单据还没有审核，因此只需要更新调价信息，其他不更新
      Update 材料特性 Set 成本价 = c_成本调整.新成本价 Where 材料id = c_成本调整.材料id;
    
      Update 成本价调价信息
      Set 收发id = n_收发id, 执行日期 = d_调价时间, 效期 = c_成本调整.效期, 灭菌效期 = c_成本调整.灭菌效期, 产地 = c_成本调整.产地, 批号 = c_成本调整.批号
      Where ID = c_成本调整.调价id;
    Elsif Nvl(c_成本调整.实际数量, 0) = 0 And (Nvl(c_成本调整.实际金额, 0) <> 0 Or Nvl(c_成本调整.实际差价, 0) <> 0) Then
      --数量=0 金额或差价<>0时只更新库存表中对应的平均成本价和特性表中成本价，并产生成本价修正数据但是差价差=0，只记录最新成本价 
      --产生调价记录，只记录最新成本价 
      Insert Into 药品收发记录
        (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期, 审核人,
         审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率)
      Values
        (n_收发id, 1, 18, v_No, n_序号, c_成本调整.库房id, n_入出类别id, c_成本调整.上次供应商id, n_入出系数, c_成本调整.材料id, c_成本调整.批次, c_成本调整.产地,
         c_成本调整.批号, c_成本调整.效期, 0, c_成本调整.实际金额, c_成本调整.实际差价, 0, '卫生材料成本价调价', Zl_Username, d_调价时间, Zl_Username, d_调价时间,
         c_成本调整.生产日期, c_成本调整.批准文号, c_成本调整.新成本价, 1, c_成本调整.原成本价);
    
      Zl_未审药品记录_Insert(n_收发id);
      --更新库存 
      Update 药品库存
      Set 平均成本价 = c_成本调整.新成本价, 上次采购价 = c_成本调整.新成本价
      Where 库房id = c_成本调整.库房id And 药品id = c_成本调整.材料id And Nvl(批次, 0) = c_成本调整.批次 And 性质 = 1;
      Update 材料特性 Set 成本价 = c_成本调整.新成本价 Where 材料id = c_成本调整.材料id;
    
      Update 成本价调价信息
      Set 收发id = n_收发id, 执行日期 = d_调价时间, 效期 = c_成本调整.效期, 灭菌效期 = c_成本调整.灭菌效期, 产地 = c_成本调整.产地, 批号 = c_成本调整.批号
      Where ID = c_成本调整.调价id;*/
    --Else
    --调整相应的库存:原成本金额-实新成本金额 
    n_调整额   := Round(c_成本调整.原成本价 * c_成本调整.实际数量, 2) - Round(c_成本调整.新成本价 * c_成本调整.实际数量, 2);
    n_原成本价 := c_成本调整.原成本价;
  
    If n_原成本价 <= 0 Then
      n_原成本价 := c_成本调整.上次采购价;
    End If;
  
    --目前：收发记录对应: 
    -- 扣率--> 原成本价 
    -- 单量-->新成本价 
    -- 填写数量-->库存实际数量 
    -- 零售价-->库存实际金额 
    -- 成本价-->库存实际差价 
    -- 差价-->本次调整额 
  
    Insert Into 药品收发记录
      (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 零售价, 成本价, 差价, 摘要, 填制人, 填制日期, 审核人,
       审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率)
    Values
      (n_收发id, 1, 18, v_No, n_序号, c_成本调整.库房id, n_入出类别id, c_成本调整.上次供应商id, n_入出系数, c_成本调整.材料id, c_成本调整.批次, c_成本调整.产地,
       c_成本调整.批号, c_成本调整.效期, c_成本调整.实际数量, c_成本调整.实际金额, c_成本调整.实际差价, n_调整额, '卫生材料成本价调价', Zl_Username, d_调价时间, Zl_Username,
       d_调价时间, c_成本调整.生产日期, c_成本调整.批准文号, c_成本调整.新成本价, 1, n_原成本价);
  
    Zl_未审药品记录_Insert(n_收发id);
    --更新库存 
    Update 药品库存
    Set 实际差价 = Nvl(实际差价, 0) + n_调整额
    Where 库房id = c_成本调整.库房id And 药品id = c_成本调整.材料id And Nvl(批次, 0) = Nvl(c_成本调整.批次, 0) And 性质 = 1;
  
    If Sql%NotFound Then
      Insert Into 药品库存
        (库房id, 药品id, 批次, 性质, 实际差价, 上次批号, 效期, 上次产地, 上次供应商id, 上次生产日期, 批准文号, 灭菌效期)
      Values
        (c_成本调整.库房id, c_成本调整.材料id, c_成本调整.批次, 1, n_调整额, c_成本调整.批号, c_成本调整.效期, c_成本调整.产地, c_成本调整.上次供应商id, c_成本调整.生产日期,
         c_成本调整.批准文号, c_成本调整.灭菌效期);
    End If;
  
    Update 药品库存
    Set 上次采购价 = c_成本调整.新成本价
    Where 药品id = c_成本调整.材料id And 上次采购价 <> c_成本调整.新成本价;
  
    Update 材料特性
    Set 成本价 = c_成本调整.新成本价
    Where 材料id = c_成本调整.材料id And 成本价 <> c_成本调整.新成本价;
  
    --重新计算库存表中的平均成本价 
    Update 药品库存
    Set 平均成本价 = Decode(Nvl(批次, 0), 0, Decode((实际金额 - 实际差价) / 实际数量, 0, 上次采购价, (实际金额 - 实际差价) / 实际数量), 上次采购价)
    Where 药品id = c_成本调整.材料id And Nvl(批次, 0) = Nvl(c_成本调整.批次, 0) And 库房id = c_成本调整.库房id And 性质 = 1 And Nvl(实际数量, 0) <> 0;
    If Sql%NotFound Then
      Select 成本价 Into n_平均成本价 From 材料特性 Where 材料id = c_成本调整.材料id;
      Update 药品库存
      Set 平均成本价 = n_平均成本价
      Where 药品id = c_成本调整.材料id And 库房id = c_成本调整.库房id And Nvl(批次, 0) = Nvl(c_成本调整.批次, 0) And 性质 = 1;
    End If;
  
    --更新成本价调价信息 
    Update 成本价调价信息
    Set 收发id = n_收发id, 执行日期 = d_调价时间, 原成本价 = n_原成本价, 效期 = c_成本调整.效期, 灭菌效期 = c_成本调整.灭菌效期, 产地 = c_成本调整.产地, 批号 = c_成本调整.批号
    Where ID = c_成本调整.调价id;
    --End If;
  
    --消息处理
    b_Message.Zlhis_Drug_010(c_成本调整.调价id);
  End Loop;

  --产生应付记录 
  For c_应付 In (Select Distinct a.供药单位id, a.药品id, a.发票号, a.发票日期, a.发票金额, b.名称, b.计算单位, b.规格
               From 成本价调价信息 A, 收费项目目录 B
               Where a.药品id = b.Id And Nvl(a.应付款变动, 0) = 1 And Nvl(a.供药单位id, 0) <> 0 And a.药品id = 材料id_In
               Order By a.供药单位id) Loop
  
    v_应付单据号 := Nextno(67);
  
    Select 应付记录_Id.Nextval Into v_应付id From Dual;
  
    Insert Into 应付记录
      (ID, 记录性质, 记录状态, 单位id, NO, 系统标识, 发票号, 发票日期, 发票金额, 品名, 规格, 填制人, 填制日期, 审核人, 审核日期, 摘要)
    Values
      (v_应付id, 1, 1, c_应付.供药单位id, v_应付单据号, 5, c_应付.发票号, c_应付.发票日期, c_应付.发票金额, c_应付.名称, c_应付.规格, Zl_Username, d_调价时间,
       Zl_Username, d_调价时间, '成本价调价自动产生应付款变动记录');
  
    If Nvl(c_应付.供药单位id, 0) <> 0 Then
      Update 应付余额 Set 金额 = Nvl(金额, 0) + Nvl(c_应付.发票金额, 0) Where 单位id = c_应付.供药单位id And 性质 = 1;
      If Sql%NotFound Then
        Insert Into 应付余额 (单位id, 性质, 金额) Values (c_应付.供药单位id, 1, Nvl(c_应付.发票金额, 0));
      End If;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_材料收发记录_成本价调价;
/

--125781:焦博,2018-05-17,调整oracle过程Zl_Third_Getvisitinfo,删除无用的外连接
CREATE OR REPLACE Procedure Zl_Third_Getvisitinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:根据挂号单号获取该次就诊详情(医嘱为主要显示) 
  --入参:Xml_In: 
  --<IN> 
  --    <GHDH>挂号单号</GHDH> 
  --    <JSKLB>结算卡类别</JSKLB> 
  --    <MXGL>明细过滤</MXGL> 0-不过滤,明细包含治疗 1-过滤,明细不包含治疗,默认为1 
  --</IN> 
  --出参:Xml_Out 
  --<OUTPUT> 
  --  <GH> 
  --     <GHDH>挂号单号</GHDH> //本次查询的挂号单号 
  --     <YYSJ>预约时间</YYSJ> //yyyy-mm-dd hh24:mi:ss 
  --     <JZSJ></JZSJ>      //实际就诊时间 
  --     <DJH></DJH>        //单据号 
  --     <JE></JE>          //金额 
  --     <DJLX></DJLX>      //单据类型,1-收费单，4-挂号单 
  --     <KDSJ></KDSJ>      //开单时间 
  --     <JKFS></JKFS>      //缴款方式,0-挂号或预约缴款;1-预约不缴款 
  --     <ZFZT></ZFZT>  //支付状态,0-待支付，1-已支付，2-已退费 
  --     <SFJSK></SFJSK>    //是否结算卡支付，0-否，1-是 
  --  </GH> 
  --  <YZLIST> 
  --     <YZ>                   //医嘱返回与HIS中显示的内容相同 
  --        <YZID><YZID>        //医嘱ID，返回组医嘱ID 
  --        <YZLX><YZLX>        //医嘱类型,如处方、检查、检验 
  --         <YZMC></YZMC>        //医嘱名称 
  --        <ZXKS></ZXKS>       //执行科室 
  --        <ZXKSID></ZXKSID>   //执行科室ID 
  --        <FYCK></FYCK>       //发药窗口 
  --        <YZMX> 
  --           <MX> 
  --              <YZNR></YZNR>        //医嘱内容 
  --              <ZXZT></ZXZT>        //医嘱执行状态 
  --              <SFFY>是否发药</SFFY> // 0-否 ，1-是 
  --              <GG>规格</GG> 
  --              <SL>数量</SL> 
  --              <DW>计算单位</DW> 
  --              <BZDJ>标准单价</BZDJ> 
  --              <YSJE>应收金额</YSJE> 
  --              <SSJE>实收金额</SSJE> 
  --           </MX> 
  --           <MX/> 
  --        </YZMX> 
  --        <BG></BG>                   //是否已出报告，是否签名 
  --        <BGLY></BGLY>               //是否外检项目,1-院内项目，2-外检项目 
  --        <BGLYSM></BGLYSM>           //外检项目说明 
  --        <JZBG></JZBG>                //禁止显示报告。0-允许，1-禁止 
  --        <JZTS></JZTS>                 //提示文字。对于禁止查看的报告，可返回用于提示病人的信息 
  --        <BLID></BLID>              //病历ID，如果<BG>字段为1，该值不为空 
  --        <DJLIST> 
  --           <DJ>                //费用单据信息 
  --              <DJH></DJH>      //费用单据号 
  --              <DJLX></DJLX>    //单据类型 
  --              <JE></JE>        //单据总金额 
  --              <KDSJ></KDSJ>    //开单时间 
  --              <ZFZT></ZFZT>    //支付状态,0-待支付，1-已支付，2-已退费,3-退费申请中,4-审核通过,5-审核未通过 
  --              <SHSM></SHSM>    //审核说明,审核未通过原因 
  --              <SFJSK></SFJSK>  //是否结算卡支付，0-否，1-是 
  --           </DJ> 
  --           <DJ/> 
  --        </DJLIST> 
  --     </YZ> 
  --  </YZLIST> 
  --    <ERROR><MSG></MSG></ERROR>                      //如果错误返回 
  --</OUTPUT> 

  -------------------------------------------------------------------------------------------------- 
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  x_Templet Xmltype; --模板XML 

  v_卡类别   Varchar2(100);
  n_卡类别id Number(18);
  v_挂号单   Varchar2(10);
  v_排队号码 Varchar2(10);
  n_Temp     Number(18);
  v_队列名称 排队叫号队列.队列名称%Type;

  n_Count Number(18);

  v_Temp       Varchar2(32767); --临时XML 
  v_队列       Varchar2(32767);
  v_No         Varchar2(50);
  n_Add_Djlist Number(1); --是否增加了DJLIST的 
  n_性质       Number(2);
  n_组医嘱id   Number(18);
  n_独立医嘱   Number(8);
  n_执行科室id Number(18);
  v_执行科室   Varchar2(50);
  n_退款金额   病人预交记录.冲预交%Type;
  n_明细过滤   Number(3);
  n_退费状态   病人退费申请.状态%Type;
  v_申请原因   病人退费申请.申请原因%Type;
  v_审核原因   病人退费申请.审核原因%Type;
  v_发药窗口   门诊费用记录.发药窗口%Type;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/MXGL')
  Into v_挂号单, v_卡类别, n_明细过滤
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_挂号单 Is Null Then
    v_Err_Msg := '不能找到指定的挂号单号(当前挂号单号为空)';
    Raise Err_Item;
  End If;
  If n_明细过滤 Is Null Then
    n_明细过滤 := 1;
  End If;
  n_Add_Djlist := 0;

  v_Err_Msg := Null;
  If v_卡类别 Is Not Null Then
    Begin
      n_卡类别id := To_Number(v_卡类别);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
  
    If n_卡类别id = 0 Then
      Begin
        Select ID, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_Err_Msg
        From 医疗卡类别
        Where 名称 = v_卡类别;
      Exception
        When Others Then
          v_Err_Msg := '卡类别:' || v_卡类别 || '不存在!';
      End;
    
    Else
    
      Begin
        Select ID, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_Err_Msg
        From 医疗卡类别
        Where ID = n_卡类别id;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  End If;
  n_性质 := 4;
  --1.获取挂号数据 

  Select Max(收费单) Into v_No From 病人挂号记录 Where NO = v_挂号单;

  If v_No Is Not Null Then
    Select Count(*) Into n_Count From 门诊费用记录 Where NO = v_No And 记录性质 = 1;
    If n_Count <> 0 Then
      n_性质 := 1;
    End If;
  End If;
  If n_性质 = 4 Then
    v_No := v_挂号单;
  End If;

  n_Count := 0;
  For c_挂号 In (Select a.Id, v_No As NO, n_性质 As 记录性质, a.执行部门id, c.名称 As 执行部门,
                      To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, To_Char(a.预约时间, 'yyyy-mm-dd hh24:mi:ss') As 预约时间,
                      a.接收时间, To_Char(a.发生时间, 'yyyy-mm-dd HH24:mi:ss') As 就诊时间, a.号别, a.号序, b.金额, a.记录状态,
                      Decode(Nvl(a.执行状态, 0), 0, '等待接诊', 1, '完成就诊', 2, '正在就诊', -1, '取消就诊') As 执行状态,
                      Decode(Nvl(b.结帐id, 0), 0, 0, 1) As 支付标志, Decode(Nvl(a.记录性质, 0), 2, 1, 0) As 缴款方式, b.结帐id As 结帐id
               From 病人挂号记录 A,
                    (Select Max(Decode(记录状态, 0, 0, 2, 0, Nvl(结帐id, 0))) As 结帐id, Sum(实收金额) As 金额
                      From 门诊费用记录 B
                      Where 记录性质 = n_性质 And NO = v_No) B, 部门表 C
               Where a.No = v_挂号单 And a.执行部门id = c.Id) Loop
  
    If Nvl(c_挂号.记录状态, 0) <> 1 Then
      v_Err_Msg := '单据号:' || v_挂号单 || '已经被退号!';
      Raise Err_Item;
    End If;
  
    Select Max(排队号码), Max(队列名称)
    Into v_排队号码, v_队列名称
    From 排队叫号队列
    Where 业务id = c_挂号.Id And Nvl(业务类型, 0) = 0;
  
    If v_排队号码 Is Not Null Then
      --业务id_In ,业务类型_In 排队号码_In Number := Null 
      n_Temp := Zl_Getsequencebeforperons(c_挂号.Id, 0, v_排队号码, v_队列名称);
      v_队列 := v_队列 || '<DL><XH>' || v_排队号码 || '</XH><QMRS>' || n_Temp || '</QMRS></DL>';
    End If;
    n_Temp := 0;
    If Nvl(n_卡类别id, 0) <> 0 Then
      Begin
        Select 1
        Into n_Temp
        From 病人预交记录
        Where 结帐id = c_挂号.结帐id And 记录性质 = 4 And 记录状态 In (1, 3) And 卡类别id = n_卡类别id And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    End If;
  
    v_Temp := '<GHDH>' || v_挂号单 || '</GHDH>';
    v_Temp := v_Temp || '<DJH>' || c_挂号.No || '</DJH>';
    v_Temp := v_Temp || '<YYSJ>' || c_挂号.预约时间 || '</YYSJ>';
    v_Temp := v_Temp || '<JZSJ>' || c_挂号.就诊时间 || '</JZSJ>';
    v_Temp := v_Temp || '<KDSJ>' || c_挂号.登记时间 || '</KDSJ>';
    v_Temp := v_Temp || '<JKFS>' || c_挂号.缴款方式 || '</JKFS>';
    v_Temp := v_Temp || '<JE>' || c_挂号.金额 || '</JE>';
    v_Temp := v_Temp || '<DJLX>' || n_性质 || '</DJLX>';
    v_Temp := v_Temp || '<ZFZT>' || c_挂号.支付标志 || '</ZFZT>';
    v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
    If v_队列 Is Not Null Then
      v_Temp := v_Temp || v_队列;
    End If;
    v_Temp := '<GH>' || v_Temp || '</GH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    n_Count := n_Count + 1;
  End Loop;

  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '未找到指定的挂号单据:' || v_挂号单 || '!';
    Raise Err_Item;
  End If;

  --2.组建医嘱及费用相关数据 
  n_组医嘱id := 0;

  For c_医嘱 In (With 医嘱费用 As
                  (Select 医嘱id, 发送号, 记录性质, NO, Max(Nvl(执行状态, 0)) As 执行状态
                  From (Select b.医嘱id, b.发送号, b.记录性质, b.No, Nvl(b.执行状态, 0) As 执行状态
                         From 病人医嘱记录 A, 病人医嘱发送 B
                         Where a.挂号单 = v_挂号单 And a.Id = b.医嘱id
                         Union All
                         Select b.医嘱id, b.发送号, b.记录性质, b.No, Nvl(c.执行状态, 0) As 执行状态
                         From 病人医嘱记录 A, 病人医嘱发送 C, 病人医嘱附费 B
                         Where a.挂号单 = v_挂号单 And a.Id = b.医嘱id And b.医嘱id = c.医嘱id And b.发送号 = c.发送号)
                  Group By 医嘱id, 发送号, 记录性质, NO)
                 
                 Select Nvl(a.相关id, a.Id) As 组id, Decode(a.相关id, Null, 0, 1) As 附医嘱, a.Id, a.相关id, e.发药窗口,
                        Max(Decode(a.诊疗类别, 'E', Decode(q.操作类型, '2', '处方', '4', '处方', '6', '检验', m.名称), m.名称)) As 医嘱类型,
                        a.执行科室id, d.名称 As 执行科室, Decode(a.相关id, Null, a.医嘱内容, Null) As 组医嘱内容,
                        Max(Decode(a.诊疗类别, '5', 1, '6', 1, '7', 1, 0) * Decode(Nvl(e.执行状态, 0), 1, 1, 3, 1, 0)) As 发药状态,
                        Decode(a.相关id, Null, Null, q.名称) As 明细医嘱内容, s.规格, (e.数次 * e.付数) As 数量, e.计算单位 As 单位,
                        Decode(Nvl(b.执行状态, 0), 0, '未执行', 1, '完全执行', 2, '拒绝执行', 3, '正在执行', '正在执行') As 执行状态,
                        Max(Decode(p.审核时间, Null, Decode(C1.完成时间, Null, 0, 1), 1)) As 是否已出报告, c.病历id, e.No, e.记录性质 As 单据类型,
                        Max(e.标准单价) As 标准单价, Sum(e.应收金额) As 应收金额, Sum(e.实收金额) As 实收金额,
                        To_Char(e.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 开单时间, Decode(Nvl(e.记录状态, 0), 0, 0, 3, 2, 1) As 支付状态,
                        a.病人id
                 
                 From 病人医嘱记录 A, 医嘱费用 B, 病人医嘱报告 C, 电子病历记录 C1, 部门表 D, 门诊费用记录 E, 诊疗项目类别 M, 诊疗项目目录 Q, 收费项目目录 S, 检验标本记录 P
                 Where a.Id = b.医嘱id And a.执行科室id = d.Id And c.病历id = C1.Id(+) And a.Id = c.医嘱id(+) And a.Id = p.医嘱id(+) And
                       b.医嘱id = e.医嘱序号 And e.收费细目id = s.Id And b.No = e.No And b.记录性质 = e.记录性质 And e.记录状态 <> 2 And
                       a.挂号单 = v_挂号单 And a.诊疗类别 = m.编码 And a.诊疗项目id = q.Id And a.医嘱状态 In (3, 8)
                 Group By a.Id, a.婴儿, a.序号, a.相关id, e.发药窗口, a.诊疗类别, a.执行科室id, d.名称, a.医嘱内容, q.名称, s.规格, e.数次 * e.付数,
                          e.计算单位, Decode(Nvl(b.执行状态, 0), 0, '未执行', 1, '完全执行', 2, '拒绝执行', 3, '正在执行', '正在执行'), C1.完成时间,
                          Decode(c.病历id, Null, 0, 1), c.病历id, e.No, e.记录性质, e.登记时间, Decode(Nvl(e.记录状态, 0), 0, 0, 3, 2, 1),
                          p.审核时间, a.病人id
                 Order By 组id, 附医嘱, Nvl(a.婴儿, 0), a.序号) Loop
    If Nvl(n_Add_Djlist, 0) = 0 Then
      --增加DJList节点 
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<YZLIST></YZLIST>')) Into x_Templet From Dual;
      n_Add_Djlist := 1;
    End If;
  
    If n_组医嘱id <> Nvl(c_医嘱.组id, 0) Then
      n_组医嘱id := Nvl(c_医嘱.组id, 0);
    
      Zl_Third_Custom_Getdeptinfo(n_组医嘱id, n_执行科室id, v_执行科室);
    
      If Nvl(n_执行科室id, 0) = 0 Then
        If c_医嘱.医嘱类型 = '检验' Then
          --检验医嘱以显示采集科室 
          n_执行科室id := c_医嘱.执行科室id;
          v_执行科室   := c_医嘱.执行科室;
        Else
          Begin
            Select b.Id, b.名称, c.发药窗口
            Into n_执行科室id, v_执行科室, v_发药窗口
            From 病人医嘱记录 A, 部门表 B, 门诊费用记录 C
            Where a.Id = c.医嘱序号 And a.相关id = n_组医嘱id And a.执行科室id = b.Id And Rownum <= 1;
          Exception
            When Others Then
              n_执行科室id := c_医嘱.执行科室id;
              v_执行科室   := c_医嘱.执行科室;
              v_发药窗口   := c_医嘱.发药窗口;
          End;
        End If;
      End If;
    
      v_Temp := '<YZID>' || n_组医嘱id || '</YZID>';
      v_Temp := v_Temp || '<YZLX>' || c_医嘱.医嘱类型 || '</YZLX>';
      v_Temp := v_Temp || '<YZMC>' || c_医嘱.组医嘱内容 || '</YZMC>';
      v_Temp := v_Temp || '<ZXKS>' || v_执行科室 || '</ZXKS>';
      v_Temp := v_Temp || '<ZXKSID>' || n_执行科室id || '</ZXKSID>';
      v_Temp := v_Temp || '<FYCK>' || v_发药窗口 || '</FYCK>';
      v_Temp := v_Temp || '<BG>' || c_医嘱.是否已出报告 || '</BG>';
      v_Temp := v_Temp || Zl_Third_Custom_Getrptfrom(n_组医嘱id);
      v_Temp := v_Temp || Zl_Third_Custom_Rptlimit(c_医嘱.病人id, n_组医嘱id);
      If Nvl(c_医嘱.是否已出报告, 0) = 1 And c_医嘱.病历id Is Not Null Then
        v_Temp := v_Temp || '<BLID>' || c_医嘱.病历id || '</BLID>';
      End If;
      v_Temp := '<YZ 医嘱ID="' || n_组医嘱id || '">' || v_Temp || '<YZMX></YZMX><DJLIST></DJLIST></YZ>';
      Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      For v_费用 In (
                   
                   Select a.No, Mod(a.记录性质, 10) As 单据类型, To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 开单时间,
                           Max(Decode(Nvl(a.记录状态, 0), 0, 0, 3, 2, 1)) As 支付状态, Sum(a.实收金额) As 单据金额, Max(a.结帐id) As 结算卡支付
                   From 门诊费用记录 A
                   Where (a.No, a.记录性质) In
                         (Select Distinct q.No, q.记录性质
                          From 病人医嘱记录 M, 病人医嘱发送 Q
                          Where m.Id = q.医嘱id And (m.Id = n_组医嘱id Or m.相关id = n_组医嘱id)
                          Union All
                          Select Distinct q.No, q.记录性质
                          From 病人医嘱记录 M, 病人医嘱附费 Q
                          Where m.Id = q.医嘱id And (m.Id = n_组医嘱id Or m.相关id = n_组医嘱id)) And Nvl(a.记录状态, 0) In (0, 1, 3)
                   Group By a.No, Mod(a.记录性质, 10), To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss')) Loop
        Begin
          Select 1
          Into n_Temp
          From 病人预交记录 A, 门诊费用记录 B
          Where a.结帐id = b.结帐id And b.No = v_费用.No And Mod(b.记录性质, 10) = 1 And b.记录状态 In (1, 3) And a.卡类别id = n_卡类别id And
                Rownum < 2;
        Exception
          When Others Then
            n_Temp := 0;
        End;
        Begin
          Select -1 * Sum(结帐金额)
          Into n_退款金额
          From 门诊费用记录 B
          Where b.No = v_费用.No And Mod(b.记录性质, 10) = 1 And b.记录状态 = 2;
        Exception
          When Others Then
            n_退款金额 := 0;
        End;
        Begin
          Select 状态, 申请原因, 审核原因
          Into n_退费状态, v_申请原因, v_审核原因
          From 病人退费申请
          Where NO = v_费用.No And Mod(记录性质, 10) = Mod(v_费用.单据类型, 10);
        Exception
          When Others Then
            n_退费状态 := -1;
            v_申请原因 := '';
            v_审核原因 := '';
        End;
      
        v_Temp := '<DJH>' || v_费用.No || '</DJH>';
        v_Temp := v_Temp || '<DJLX>' || v_费用.单据类型 || '</DJLX>';
        v_Temp := v_Temp || '<JE>' || v_费用.单据金额 || '</JE>';
        v_Temp := v_Temp || '<KDSJ>' || v_费用.开单时间 || '</KDSJ>';
        If n_退费状态 = -1 Then
          v_Temp := v_Temp || '<ZFZT>' || v_费用.支付状态 || '</ZFZT>';
        Else
          If n_退费状态 = 0 Then
            v_Temp := v_Temp || '<ZFZT>3</ZFZT>';
          End If;
          If n_退费状态 = 1 Then
            If v_费用.支付状态 = 2 Then
              v_Temp := v_Temp || '<ZFZT>2</ZFZT>';
            Else
              v_Temp := v_Temp || '<ZFZT>4</ZFZT>';
            End If;
          End If;
          If n_退费状态 = 2 Then
            v_Temp := v_Temp || '<ZFZT>5</ZFZT>';
          End If;
        End If;
      
        If n_退费状态 = -1 Then
          v_Temp := v_Temp || '<SHSM>' || '' || '</SHSM>';
        Else
          If n_退费状态 = 0 Then
            v_Temp := v_Temp || '<SHSM>' || v_申请原因 || '</SHSM>';
          End If;
          If n_退费状态 = 1 Then
            v_Temp := v_Temp || '<SHSM>' || v_审核原因 || '</SHSM>';
          End If;
          If n_退费状态 = 2 Then
            v_Temp := v_Temp || '<SHSM>' || v_审核原因 || '</SHSM>';
          End If;
        End If;
      
        v_Temp := v_Temp || '<YTJE>' || Nvl(n_退款金额, 0) || '</YTJE>';
        v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
        v_Temp := '<DJ>' || v_Temp || '</DJ>';
        Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@医嘱ID="' || n_组医嘱id || '"]/DJLIST', Xmltype(v_Temp))
        Into x_Templet
        From Dual;
      End Loop;
    End If;
  
    --只有一条记录的医嘱，在明细中增加该条医嘱，以获取执行状态 
    Select Decode(Count(*), 0, 1, 0) Into n_独立医嘱 From 病人医嘱记录 Where 相关id = n_组医嘱id;
    If n_独立医嘱 = 1 Then
      v_Temp := '<YZNR>' || c_医嘱.组医嘱内容 || '</YZNR>';
      v_Temp := v_Temp || '<GG>' || c_医嘱.规格 || '</GG>';
      v_Temp := v_Temp || '<SFFY>' || c_医嘱.发药状态 || '</SFFY>';
      v_Temp := v_Temp || '<SL>' || c_医嘱.数量 || '</SL>';
      v_Temp := v_Temp || '<DW>' || c_医嘱.单位 || '</DW>';
      v_Temp := v_Temp || '<BZDJ>' || Nvl(c_医嘱.标准单价, 0) || '</BZDJ>';
      v_Temp := v_Temp || '<YSJE>' || Nvl(c_医嘱.应收金额, 0) || '</YSJE>';
      v_Temp := v_Temp || '<SSJE>' || Nvl(c_医嘱.实收金额, 0) || '</SSJE>';
      v_Temp := v_Temp || '<ZXZT>' || c_医嘱.执行状态 || '</ZXZT>';
      v_Temp := '<MX>' || v_Temp || '</MX>';
      Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@医嘱ID="' || n_组医嘱id || '"]/YZMX', Xmltype(v_Temp))
      Into x_Templet
      From Dual;
    End If;
  
    If Nvl(c_医嘱.附医嘱, 0) = 1 Then
      If n_明细过滤 = 0 Or (n_明细过滤 = 1 And c_医嘱.医嘱类型 <> '治疗') Then
        v_Temp := '<YZNR>' || c_医嘱.明细医嘱内容 || '</YZNR>';
        v_Temp := v_Temp || '<GG>' || c_医嘱.规格 || '</GG>';
        v_Temp := v_Temp || '<SL>' || c_医嘱.数量 || '</SL>';
        v_Temp := v_Temp || '<DW>' || c_医嘱.单位 || '</DW>';
        v_Temp := v_Temp || '<SFFY>' || c_医嘱.发药状态 || '</SFFY>';
        v_Temp := v_Temp || '<ZXZT>' || c_医嘱.执行状态 || '</ZXZT>';
        v_Temp := v_Temp || '<BZDJ>' || Nvl(c_医嘱.标准单价, 0) || '</BZDJ>';
        v_Temp := v_Temp || '<YSJE>' || Nvl(c_医嘱.应收金额, 0) || '</YSJE>';
        v_Temp := v_Temp || '<SSJE>' || Nvl(c_医嘱.实收金额, 0) || '</SSJE>';
        v_Temp := '<MX>' || v_Temp || '</MX>';
        Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@医嘱ID="' || n_组医嘱id || '"]/YZMX', Xmltype(v_Temp))
        Into x_Templet
        From Dual;
      End If;
    End If;
  
  End Loop;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getvisitinfo;
/

--119329:冉俊明,2018-05-16,三方接口获取可挂号科室过程调整
Create Or Replace Procedure Zl_Third_Getdeptlist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --功能:获取可挂号科室
  --入参:Xml_In:
  --<IN>
  --  <CXTS>14</CXTS>        //查询天数
  --  <HZDW>支付宝</HZDW>    //合作单位
  --  <ZD></ZD>              //站点
  --</IN>
  --出参:Xml_Out
  --<OUTPUT>
  -- <KSLIST>
  --  <KS>
  --    <ID>科室ID</ID>       //科室ID
  --    <MC>科室名称</MC>     //科室名称
  --    <ZDYYTS>最大可预约天数</ZDYYTS>     //最大可预约天数
  --  </KS>
  --  <KS>
  --    ...
  --  </KS>
  -- </KSLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  x_Templet Xmltype; --模板XML
  v_Temp    Varchar2(5000); --临时XML

  n_查询天数 Number(5);
  v_合作单位 合作单位安排控制.合作单位%Type;
  v_站点     部门表.站点%Type;

  v_Para     Varchar2(4000);
  n_预约天数 Number(5);
  n_补充天数 Number(5);

  n_挂号模式 Number(3);
  d_启用时间 Date;
Begin
  x_Templet := Xmltype('<OUTPUT><KSLIST></KSLIST></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/CXTS'), Extractvalue(Value(A), 'IN/HZDW'), Nvl(Extractvalue(Value(A), 'IN/ZD'), '-')
  Into n_查询天数, v_合作单位, v_站点
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := To_Number(Substr(v_Para, 1, 1));
  If n_挂号模式 = 1 Then
    Begin
      d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_启用时间 := Null;
    End;
  End If;
  n_预约天数 := zl_GetSysParameter(66);

  If n_挂号模式 = 0 Then
    If v_合作单位 Is Null Then
      For r_Dept In (Select a.科室id, d.名称, Max(Nvl(a.预约天数, n_预约天数)) As 预约天数
                     From 挂号安排 A, 部门表 D
                     Where a.科室id = d.Id And a.停用日期 Is Null And Nvl(d.站点, v_站点) = v_站点 And
                           (d.撤档时间 Is Null Or d.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd'))
                     Group By a.科室id, d.名称) Loop
        v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>';
        v_Temp := v_Temp || '<ZDYYTS>' || r_Dept.预约天数 || '</ZDYYTS>' || '</KS>';
        Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
      End Loop;
    Else
      For r_Dept In (Select 科室id, 名称, Max(预约天数) As 预约天数
                     From (
                            --1.计划
                            Select a.科室id, d.名称, Nvl(a.预约天数, n_预约天数) As 预约天数
                            From 挂号安排 A, 挂号安排计划 C, 部门表 D
                            Where a.Id = c.安排id And a.科室id = d.Id And a.停用日期 Is Null And c.审核时间 Is Not Null And
                                  Not (c.失效时间 < Sysdate Or c.生效时间 > Sysdate + Nvl(n_查询天数, Nvl(a.预约天数, n_预约天数)))
                                 
                                  And
                                  (Not Exists (Select 1 From 合作单位计划控制 Where 计划id = c.Id And 合作单位 = v_合作单位) Or Exists
                                   (Select 1
                                    From 合作单位计划控制
                                    Where 计划id = c.Id And 合作单位 = v_合作单位 And 数量 <> 0))
                                 
                                  And Nvl(d.站点, v_站点) = v_站点 And
                                  (d.撤档时间 Is Null Or d.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd'))
                            --2.安排
                            Union All
                            Select a.科室id, d.名称, Nvl(a.预约天数, n_预约天数) As 预约天数
                            From 挂号安排 A, 部门表 D
                            Where a.科室id = d.Id And a.停用日期 Is Null And Not Exists
                             (Select 1 From 挂号安排计划 Where 安排id = a.Id)
                                 
                                  And
                                  (Not Exists (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位) Or Exists
                                   (Select 1
                                    From 合作单位安排控制
                                    Where 安排id = a.Id And 合作单位 = v_合作单位 And 数量 <> 0))
                                 
                                  And Nvl(d.站点, v_站点) = v_站点 And
                                  (d.撤档时间 Is Null Or d.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')))
                     Group By 科室id, 名称) Loop
        v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>';
        v_Temp := v_Temp || '<ZDYYTS>' || r_Dept.预约天数 || '</ZDYYTS>' || '</KS>';
        Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
      End Loop;
    End If;
  Else
    --出诊表排班模式
    n_补充天数 := Zl_Fun_Getappointmentdays;
    If v_合作单位 Is Null Then
      For r_Dept In (Select 科室id, 名称, Max(预约天数) As 预约天数
                     From (
                            --启用前
                            Select a.科室id, d.名称, Nvl(a.预约天数, n_预约天数) As 预约天数
                            From 挂号安排 A, 部门表 D
                            Where a.科室id = d.Id And a.停用日期 Is Null And Nvl(d.站点, v_站点) = v_站点 And
                                  (d.撤档时间 Is Null Or d.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')) And Sysdate < d_启用时间
                            --启用后
                            Union All
                            Select a.科室id, d.名称, Nvl(c.预约天数, n_预约天数) + n_补充天数 As 预约天数
                            From 临床出诊记录 A, 临床出诊号源 C, 部门表 D
                            Where a.号源id = c.Id And a.科室id = d.Id And a.出诊日期 Between Trunc(Sysdate) And
                                  Trunc(Sysdate) + Nvl(n_查询天数, Nvl(c.预约天数, n_预约天数) + n_补充天数) And a.开始时间 > d_启用时间 And
                                  Nvl(a.是否发布, 0) = 1
                                 --排除全时段停诊了的
                                  And (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) And Nvl(a.停诊开始时间, a.终止时间) > Sysdate Or
                                  a.终止时间 > Nvl(a.停诊终止时间, a.开始时间) And a.终止时间 > Sysdate Or Exists
                                   (Select 1
                                        From 临床出诊序号控制
                                        Where 记录id = a.Id And Nvl(是否停诊, 0) = 0 And Nvl(a.是否序号控制, 0) = 1 And
                                              Nvl(a.是否分时段, 0) = 1 And 开始时间 <> 终止时间 And 开始时间 >= Sysdate))
                                 --
                                  And (c.撤档时间 Is Null Or c.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')) And
                                  Nvl(d.站点, v_站点) = v_站点 And
                                  (d.撤档时间 Is Null Or d.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')))
                     Group By 科室id, 名称) Loop
        v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>';
        v_Temp := v_Temp || '<ZDYYTS>' || r_Dept.预约天数 || '</ZDYYTS>' || '</KS>';
        Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
      End Loop;
    Else
      For r_Dept In (Select 科室id, 名称, Max(预约天数) As 预约天数
                     From (
                            --1.启用前
                            --1.1.计划
                            Select a.科室id, d.名称, Nvl(a.预约天数, n_预约天数) As 预约天数
                            From 挂号安排 A, 挂号安排计划 C, 部门表 D
                            Where a.Id = c.安排id And a.科室id = d.Id And a.停用日期 Is Null And c.审核时间 Is Not Null And
                                  Not (c.失效时间 < Sysdate Or c.生效时间 > Sysdate + Nvl(n_查询天数, Nvl(a.预约天数, n_预约天数)))
                                 
                                  And
                                  (Not Exists (Select 1 From 合作单位计划控制 Where 计划id = c.Id And 合作单位 = v_合作单位) Or Exists
                                   (Select 1
                                    From 合作单位计划控制
                                    Where 计划id = c.Id And 合作单位 = v_合作单位 And 数量 <> 0))
                                 
                                  And Nvl(d.站点, v_站点) = v_站点 And
                                  (d.撤档时间 Is Null Or d.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')) And Sysdate < d_启用时间
                            --1.2.安排
                            Union All
                            Select a.科室id, d.名称, Nvl(a.预约天数, n_预约天数) As 预约天数
                            From 挂号安排 A, 部门表 D
                            Where a.科室id = d.Id And a.停用日期 Is Null And Not Exists
                             (Select 1 From 挂号安排计划 Where 安排id = a.Id)
                                 
                                  And
                                  (Not Exists (Select 1 From 合作单位安排控制 Where 安排id = a.Id And 合作单位 = v_合作单位) Or Exists
                                   (Select 1
                                    From 合作单位安排控制
                                    Where 安排id = a.Id And 合作单位 = v_合作单位 And 数量 <> 0))
                                 
                                  And Nvl(d.站点, v_站点) = v_站点 And
                                  (d.撤档时间 Is Null Or d.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')) And Sysdate < d_启用时间
                            --2.启用后
                            Union All
                            Select a.科室id, d.名称, Nvl(c.预约天数, n_预约天数) + n_补充天数 As 预约天数
                            From 临床出诊记录 A, 临床出诊号源 C, 部门表 D
                            Where a.号源id = c.Id And a.科室id = d.Id And a.出诊日期 Between Trunc(Sysdate) And
                                  Trunc(Sysdate) + Nvl(n_查询天数, Nvl(c.预约天数, n_预约天数) + n_补充天数) And a.开始时间 >= d_启用时间 And
                                  Nvl(a.是否发布, 0) = 1
                                 --排除全时段停诊了的
                                  And (a.开始时间 < Nvl(a.停诊开始时间, a.终止时间) And Nvl(a.停诊开始时间, a.终止时间) > Sysdate Or
                                  a.终止时间 > Nvl(a.停诊终止时间, a.开始时间) And a.终止时间 > Sysdate Or Exists
                                   (Select 1
                                        From 临床出诊序号控制
                                        Where 记录id = a.Id And Nvl(是否停诊, 0) = 0 And Nvl(a.是否序号控制, 0) = 1 And
                                              Nvl(a.是否分时段, 0) = 1 And 开始时间 <> 终止时间 And 开始时间 >= Sysdate))
                                 --临床出诊记录.预约控制：0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
                                  And
                                  (Not Exists (Select 1
                                               From 临床出诊挂号控制记录
                                               Where 记录id = a.Id And 性质 = 1 And 类型 = 1 And 名称 = v_合作单位) Or Exists
                                   (Select 1
                                    From 临床出诊挂号控制记录
                                    Where 记录id = a.Id And 性质 = 1 And 类型 = 1 And 名称 = v_合作单位
                                         --临床出诊挂号控制记录.控制方式：0-禁止预约;1-按比例控制预约;2-按总量控制预约;3-按序号控制预约;4-不作限制
                                          And (控制方式 In (1, 2, 3) And 数量 <> 0 Or 控制方式 = 4)))
                                 
                                  And (c.撤档时间 Is Null Or c.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')) And
                                  Nvl(d.站点, v_站点) = v_站点 And
                                  (d.撤档时间 Is Null Or d.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')))
                     Group By 科室id, 名称) Loop
        v_Temp := '<KS>' || '<ID>' || r_Dept.科室id || '</ID>' || '<MC>' || r_Dept.名称 || '</MC>';
        v_Temp := v_Temp || '<ZDYYTS>' || r_Dept.预约天数 || '</ZDYYTS>' || '</KS>';
        Select Appendchildxml(x_Templet, '/OUTPUT/KSLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
      End Loop;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdeptlist;
/

--125688:刘涛,2018-05-16,药品申领入库房批次处理
Create Or Replace Procedure Zl_药品申领_Insert
(
  No_In           In 药品收发记录.No%Type,
  序号_In         In 药品收发记录.序号%Type,
  库房id_In       In 药品收发记录.库房id%Type,
  对方部门id_In   In 药品收发记录.对方部门id%Type,
  药品id_In       In 药品收发记录.药品id%Type,
  批次_In         In 药品收发记录.批次%Type,
  填写数量_In     In 药品收发记录.填写数量%Type,
  实际数量_In     In 药品收发记录.实际数量%Type,
  成本价_In       In 药品收发记录.成本价%Type,
  成本金额_In     In 药品收发记录.成本金额%Type,
  零售价_In       In 药品收发记录.零售价%Type,
  零售金额_In     In 药品收发记录.零售金额%Type,
  差价_In         In 药品收发记录.差价%Type,
  填制人_In       In 药品收发记录.填制人%Type,
  产地_In         In 药品收发记录.产地%Type := Null,
  批号_In         In 药品收发记录.批号%Type := Null,
  效期_In         In 药品收发记录.效期%Type := Null,
  摘要_In         In 药品收发记录.摘要%Type := Null,
  填制日期_In     In 药品收发记录.填制日期%Type := Null,
  上次供应商id_In In 药品收发记录.供药单位id%Type := Null,
  批准文号_In     In 药品收发记录.批准文号%Type := Null,
  申领方式_In     In 药品收发记录.单量%Type := 0,
  结束时间_In     In 药品收发记录.频次%Type := Null,
  原产地_In       In 药品收发记录.原产地%Type := Null,
  修改人_In       In 药品收发记录.修改人%Type,
  修改日期_In     In 药品收发记录.修改日期%Type := Null
) Is
  v_Lngid        药品收发记录.Id%Type; --收发ID 
  n_出库收发id   药品收发记录.Id%Type; --出库库房收发id 
  v_入的类别id   药品收发记录.入出类别id%Type; --入出类别ID 
  v_出的类别id   药品收发记录.入出类别id%Type; --入出类别ID 
  d_上次生产日期 药品库存.上次生产日期%Type;
  v_批次         药品收发记录.批次%Type := Null; --主要针对入库中实行药库分批的药品
  v_是否分批     Integer; --判断入库是否药库分批   1:分批；0：不分批
  v_药库分批     Integer; --判断入库是否药库分批   1:分批；0：不分批
  v_药房分批     Integer; --判断入库是否药库分批   1:分批；0：不分批
Begin
  --首先找出入和出的类别ID 
  Select b.Id 
  Into v_入的类别id 
  From 药品单据性质 A, 药品入出类别 B 
  Where a.类别id = b.Id And a.单据 = 6 And b.系数 = 1 And Rownum < 2; 
  
  Select b.Id 
  Into v_出的类别id 
  From 药品单据性质 A, 药品入出类别 B 
  Where a.类别id = b.Id And a.单据 = 6 And b.系数 = -1 And Rownum < 2; 

  Select 药品收发记录_Id.Nextval Into v_Lngid From Dual;

  Begin
    Select 上次生产日期
    Into d_上次生产日期
    From 药品库存
    Where 性质 = 1 And 库房id = 库房id_In And 药品id = 药品id_In And Nvl(批次, 0) = Nvl(批次_In, 0);
  Exception
    When Others Then
      d_上次生产日期 := Null;
  End;

  Select 药品收发记录_Id.Nextval Into n_出库收发id From Dual;
  --插入类别为出的那一笔 
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 原产地,批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 发药方式, 供药单位id, 批准文号, 生产日期, 单量, 频次, 修改人, 修改日期)
  Values
    (n_出库收发id, 1, 6, No_In, 序号_In, 库房id_In, 对方部门id_In, v_出的类别id, -1, 药品id_In, 批次_In, 产地_In, 原产地_In,批号_In, 效期_In, 填写数量_In,
     实际数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, 1, 上次供应商id_In, 批准文号_In, d_上次生产日期, 申领方式_In,
     结束时间_In, 修改人_In, 修改日期_In);
  
  Zl_未审药品记录_Insert(n_出库收发id);

  --处理库存
  Zl_药品库存_Update(n_出库收发id, 0);

  Select Nvl(药库分批, 0), Nvl(药房分批, 0) Into v_药库分批, v_药房分批 From 药品规格 Where 药品id = 药品id_In;

  v_是否分批 := 0;
  If v_药房分批 = 0 Then
    If v_药库分批 = 1 Then
      Begin
        Select Distinct 0
        Into v_是否分批
        From 部门性质说明
        Where ((工作性质 Like '%药房') Or (工作性质 Like '制剂室')) And 部门id = 对方部门id_In;
      Exception
        When Others Then
          v_是否分批 := 1;
      End;
    End If;
  Else
    v_是否分批 := 1;
  End If;

  If v_是否分批 = 1 And Nvl(批次_In, 0) = 0 Then
    --入库分批且出库不分批
    v_批次 := Zl_Fun_Getbatchnum(药品id_In, 产地_In, 批号_In, 成本价_In, 零售价_In, v_Lngid, 上次供应商id_In);
  Elsif v_是否分批 = 0 Then
    --入库不分批
    v_批次 := 0;
  Elsif Nvl(批次_In, 0) <> 0 Then
    --入库分批且出库也分批
    v_批次 := 批次_In;
  End If;

  --插入类别为入的那一笔 
  Insert Into 药品收发记录
    (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地,原产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
     填制人, 填制日期, 发药方式, 供药单位id, 批准文号, 生产日期, 单量, 频次, 修改人, 修改日期)
  Values
    (v_Lngid, 1, 6, No_In, 序号_In + 1, 对方部门id_In, 库房id_In, v_入的类别id, 1, 药品id_In, v_批次, 产地_In,原产地_In, 批号_In, 效期_In, 填写数量_In,
     实际数量_In, 成本价_In, 成本金额_In, 零售价_In, 零售金额_In, 差价_In, 摘要_In, 填制人_In, 填制日期_In, 1, 上次供应商id_In, 批准文号_In, d_上次生产日期, 申领方式_In,
     结束时间_In, 修改人_In, 修改日期_In);
  
  Zl_未审药品记录_Insert(v_Lngid);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品申领_Insert;
/

--113688:余伟节,2018-05-14,解决取消再入院的门诊留观病人的登记后病人信息中的入出院信息没有更新的问题
--97423:余伟节,2018-05-14,处理再入院病人取消登记时找不到数据的问题
Create Or Replace Procedure Zl_入院病案主页_Delete
(
  病人id_In     病案主页.病人id%Type,
  主页id_In     病案主页.主页id%Type,
  转留观_In     Number := 0,
  清除住院号_In Number := 0
  --功能：取消病人入院/预约登记
  --     主页ID_IN:为0时表示取消预约登记
  --     转留观_IN:将正常入院登记病人转为住院留观病人
  --     清除住院号_In:第一次住院的病人转留观时是否清除住院号
) As
  v_入院时间   病案主页.入院日期%Type;
  v_入院科室   病案主页.入院科室id%Type;
  v_出院时间   病案主页.出院日期%Type;
  v_住院号     病案主页.住院号%Type;
  v_再入院     病案主页.再入院%Type;
  v_出院科室id 病案主页.出院科室id%Type;
  n_病人性质   病案主页.病人性质%Type;
  n_主页id     病案主页.主页id%Type;

  v_Count Number;
  v_Error Varchar2(255);
  Err_Custom Exception;

  Function Checkpatiadvice
  (
    病人id_In 病案主页.病人id%Type,
    主页id_In 病案主页.主页id%Type
  ) Return Varchar2 Is
    --本次住院所有医嘱记录都已作废
    v_Err Varchar2(255);
  Begin
    v_Err := Null;
  
    For r_Row In (Select 开嘱医生, Decode(医嘱状态, -1, '暂存', 1, '新开', 2, '校对疑问', '未作废') As 状态, 医嘱内容
                  From 病人医嘱记录
                  Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(医嘱状态, 0) <> 4 And Rownum < 2) Loop
      v_Err := '【' || r_Row.开嘱医生 || '】医生有' || r_Row.状态 || '的医嘱没有处理,不允许取消登记！';
    End Loop;
    Return v_Err;
  End Checkpatiadvice;
Begin
  Select Nvl(状态, 0), Nvl(病人性质, 0)
  Into v_Count, n_病人性质
  From 病案主页
  Where 病人id = 病人id_In And 主页id = 主页id_In;
  If v_Count <> 1 Then
    v_Error := '该病人已经入科,请先将病人撤消至入院状态。';
    Raise Err_Custom;
  End If;

  --删除电子病历时机
  Select 出院科室id, 再入院 Into v_出院科室id, v_再入院 From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
  If v_再入院 = 0 Then
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '入院', v_出院科室id);
  Else
    Zl_电子病历时机_Delete(病人id_In, 主页id_In, '再次入院', v_出院科室id);
  End If;

  --提取最近一次不为空的住院号
  Begin
    If 主页id_In = 0 Then
      Select 住院号
      Into v_住院号
      From 病案主页
      Where 病人id = 病人id_In And
            主页id =
            (Select Max(主页id) From 病案主页 Where 病人id = 病人id_In And Nvl(主页id, 0) <> 0 And Nvl(住院号, 0) <> 0);
    Else
      Select 住院号
      Into v_住院号
      From 病案主页
      Where 病人id = 病人id_In And
            主页id =
            (Select Max(主页id) From 病案主页 Where 病人id = 病人id_In And 主页id < 主页id_In And Nvl(住院号, 0) <> 0);
    End If;
  Exception
    When Others Then
      Null;
  End;

  b_Message.ZLHIS_PATIENT_006(病人id_In,主页id_In,'入院登记');

  If 转留观_In = 1 And Nvl(主页id_In, 0) <> 0 Then
    Update 病案主页
    Set 病人性质 = 2, 住院号 = Decode(清除住院号_In, 1, Null, 住院号)
    Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(病人性质, 0) = 0;
  
    --调整住院次数
    Update 病人信息 Set 住院次数 = Decode(Sign(住院次数 - 1), 1, 住院次数 - 1, Null) Where 病人id = 病人id_In;
    If 清除住院号_In = 1 Then
      Update 病人信息 Set 住院号 = v_住院号 Where 病人id = 病人id_In;
    End If;
  Else
    Begin
      Select b.入院日期, b.出院日期, b.入院科室id
      Into v_入院时间, v_出院时间, v_入院科室
      From 病人信息 A, 病案主页 B
      Where a.病人id = 病人id_In And a.病人id = b.病人id And a.主页id = b.主页id And Nvl(b.主页id, 0) <> 0;
    Exception
      When Others Then
        Null;
    End;
    --撤消预约登记病人不检查住院日报
    If Nvl(主页id_In, 0) <> 0 Then
      Select Zl_住院日报_Count(v_入院科室, v_入院时间) Into v_Count From Dual;
      If v_Count > 0 Then
        v_Error := '已产生业务时间内的住院日报,不能办理该业务!';
        Raise Err_Custom;
      End If;
    End If;
    --门诊留观病人下达入院通知后存在两条有效的病案主页记录（36549）
    Select Count(*) Into v_Count From 病案主页 Where 病人id = 病人id_In And 入院日期 Is Not Null And 出院日期 Is Null;
    If Not v_Count > 1 Then
      v_Count := 0;
      If Nvl(主页id_In, 0) <> 0 And Nvl(n_病人性质, 0) = 0 Then
        v_Count := 1;
      End If;
      --再入院病人,取消入院登记时,病人信息的入院时间和出院时间应该回退到上一次入院日期和出院日期
      If v_再入院 = 1 Then
        Begin
          Select 入院日期, 出院日期
          Into v_入院时间, v_出院时间
          From 病案主页
          Where 病人id = 病人id_In And
                主页id = (Select Max(主页id)
                        From 病案主页
                        Where 病人id = 病人id_In And 主页id < 主页id_In);
        Exception
          When Others Then
            --异常处理是为了屏蔽取不到数据的异常情况
            Null;
        End;
      End If;
    
      Update 病人信息
      Set 住院号 = v_住院号, 住院次数 = Decode(v_Count, 0, 住院次数, Decode(Sign(住院次数 - 1), 1, 住院次数 - 1, Null)), 当前科室id = Null,
          当前病区id = Null, 当前床号 = Null, 入院时间 = v_入院时间, 出院时间 = v_出院时间, 担保人 = Null, 担保额 = Null, 担保性质 = Null, 在院 = Null
      Where 病人id = 病人id_In;
      Delete From 在院病人 Where 病人id = 病人id_In;
    End If;
    Delete From 病人变动记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
    Delete From 病人自动计算 Where 病人id = 病人id_In And 主页id = 主页id_In;
    Delete From 病人诊断记录 Where 病人id = 病人id_In And 主页id = 主页id_In And 记录来源 = 2;
  
    --本次住院如果交了预交款,改为当作门诊交的
    Update 病人预交记录 Set 主页id = Null Where 病人id = 病人id_In And 主页id = 主页id_In;
  
    --本次发卡的,改变门诊发卡
    Update 住院费用记录 Set 主页id = Null Where 病人id = 病人id_In And 主页id = 主页id_In And 记录性质 = 5;
  
    --本次住院的所有费用记录无结算且已全部冲销，则将对应费用记录中的"主页ID"清除。
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From 住院费用记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And 记帐费用 = 1 And 结帐id Is Not Null;
  
    If v_Count = 0 Then
      Begin
        Select Nvl(Count(*), 0)
        Into v_Count
        From 住院费用记录
        Where 病人id = 病人id_In And 主页id = 主页id_In And 记帐费用 = 1
        Group By NO, 记录性质, 序号
        Having Nvl(Sum(实收金额), 0) <> 0;
      Exception
        When Others Then
          v_Count := 0;
      End;
    
      If v_Count = 0 Then
        Delete 病人未结费用 Where 病人id = 病人id_In And 主页id = 主页id_In And 金额 = 0;
        Update 住院费用记录 Set 主页id = Null Where 病人id = 病人id_In And 主页id = 主页id_In And 记帐费用 = 1;
      End If;
    End If;
  
    --本次住院所有医嘱记录都已作废
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From 病人医嘱记录
    Where 病人id = 病人id_In And 主页id = 主页id_In And Nvl(医嘱状态, 0) <> 4;
    If v_Count = 0 Then
      Delete From 病人医嘱记录 Where 病人id = 病人id_In And 主页id = 主页id_In;
    Else
      v_Error := Checkpatiadvice(病人id_In, 主页id_In);
      If v_Error Is Not Null Then
        Raise Err_Custom;
      End If;
    End If;
  
    --以下表,没有建病案主页(病人ID,主页ID)的外键,因为其主页ID可能是挂号ID
    Delete From 病人过敏记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 病人诊断记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 病人新生儿记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 电子病历记录 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    Delete From 电子病历打印 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    --如果入院发放了就诊卡,则删除会失败(病人费用记录主页ID有外键约束)
    Delete From 病案主页 Where 病人id = 病人id_In And Nvl(主页id, 0) = Nvl(主页id_In, 0);
    --修改病人信息的主页ID和住院次数
    Select Max(主页id) Into n_主页id From 病案主页 Where 病人id = 病人id_In And Nvl(主页id, 0) <> 0;
    Update 病人信息 Set 主页id = n_主页id Where 病人id = 病人id_In;
    If n_主页id Is Null Then
      Update 病人信息 Set 住院次数 = Null Where 病人id = 病人id_In;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_入院病案主页_Delete;
/

--125233:焦博,2018-05-14,往合作单位挂号汇总表里插入数据时没有添加序号导致添加失败
Create Or Replace Procedure Zl_三方机构挂号_Delete
(
  单据号_In     门诊费用记录.No%Type,
  交易流水号_In 病人预交记录.交易流水号%Type,
  交易说明_In   病人预交记录.交易说明%Type,
  退号时间_In   门诊费用记录.登记时间%Type := Null,
  预交id_In     病人预交记录.Id%Type := Null
) As
  v_Error Varchar2(255);
  Err_Custom Exception;

  --该游标用于判断是否单独收病历费,及挂号汇总表处理
  Cursor c_Registinfo
  (
    v_状态     病人挂号记录.记录状态%Type,
    v_性质     病人挂号记录.记录性质%Type,
    v_无效单据 Number := 0
  ) Is
    Select a.发生时间, a.登记时间, b.项目id, b.科室id, b.医生姓名, b.医生id, b.号码
    From 病人挂号记录 A, 挂号安排 B
    Where a.记录性质 = Decode(v_无效单据, 0, v_性质, a.记录性质) And a.记录状态 = v_状态 And a.No = 单据号_In And a.号别 = b.号码 And Rownum = 1;

  r_Registrow c_Registinfo%RowType;

  --该光标用于处理人员缴款余额中退的不同结算方式的金额
  Cursor c_Opermoney Is
    Select Distinct b.结算方式, b.冲预交
    From 门诊费用记录 A, 病人预交记录 B
    Where a.结帐id = b.结帐id And a.No = 单据号_In And a.记录性质 = 4 And a.记录状态 = 3 And b.记录性质 = 4 And b.记录状态 = 3 And
          Nvl(b.冲预交, 0) <> 0;

  n_执行状态       病人挂号记录.执行状态%Type;
  n_打印id         票据打印内容.Id%Type;
  n_结帐id         门诊费用记录.结帐id%Type;
  n_原结帐id       病人预交记录.结帐id%Type;
  n_病人id         病人信息.病人id%Type;
  n_返回值         病人余额.预交余额%Type;
  n_分诊台签到排队 Number;
  n_预约挂号       Number;
  n_无效单据       Number; --无效单据没有产生费用单据
  n_挂号生成队列   Number;
  n_Count          Number;
  n_组id           财务缴款分组.Id%Type;
  d_退号时间       Date;
  v_操作员编号     人员表.编号%Type;
  v_操作员姓名     人员表.姓名%Type;
  v_合作单位       合作单位挂号汇总.合作单位%Type;
  n_预约状态       病人挂号记录.预约%Type;
  v_Temp           Varchar2(100);
  d_登记时间       病人挂号记录.登记时间%Type;
  v_号别           病人挂号记录.号别%Type;
  n_号序           病人挂号记录.号序%Type;
  d_预约时间       病人挂号记录.预约时间%Type;
  n_合作单位限制   Number(18);
  n_预约生成队列   Number;
  n_记录性质       Number;
  n_状态           Number;
  n_退号重用       Number(3);
  n_挂号排班模式   Number;
  n_记帐           门诊费用记录.记帐费用%Type;
  n_挂号id         病人挂号记录.Id%Type;
  n_已结帐         Number;
  n_返回额         病人余额.费用余额%Type;
  n_预交支付       Number(3);
  n_正常支付       Number(3);
  n_安排id         挂号安排.Id%Type;
  n_计划id         挂号安排计划.Id%Type;
  v_时间段         时间段.时间段%Type;
  d_检查开始时间   Date;
  d_检查结束时间   Date;
  d_启用时间       Date;
  n_出诊记录id     Number(18);
  Function Zl_操作员
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
    -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
  Begin
    If Type_In = 0 Then
      --缺省部门
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --操作员编码
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --操作员姓名
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_三方机构挂号_出诊_Delete
  (
    单据号_In     门诊费用记录.No%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    退号时间_In   门诊费用记录.登记时间%Type := Null,
    预交id_In     病人预交记录.Id%Type := Null
  ) As
    v_Error Varchar2(255);
    Err_Custom Exception;
  
    --该游标用于判断是否单独收病历费,及挂号汇总表处理
    Cursor c_Registinfo
    (
      v_状态     病人挂号记录.记录状态%Type,
      v_性质     病人挂号记录.记录性质%Type,
      v_无效单据 Number := 0
    ) Is
      Select a.发生时间, a.登记时间, b.项目id, b.科室id, b.医生姓名, b.医生id, b.Id As 记录id, a.号别 As 号码
      From 病人挂号记录 A, 临床出诊记录 B
      Where a.记录性质 = Decode(v_无效单据, 0, v_性质, a.记录性质) And a.记录状态 = v_状态 And a.No = 单据号_In And a.出诊记录id = b.Id And
            Rownum < 2;
  
    r_Registrow c_Registinfo%RowType;
  
    --该光标用于处理人员缴款余额中退的不同结算方式的金额
    Cursor c_Opermoney Is
      Select Distinct b.结算方式, b.冲预交
      From 门诊费用记录 A, 病人预交记录 B
      Where a.结帐id = b.结帐id And a.No = 单据号_In And a.记录性质 = 4 And a.记录状态 = 3 And b.记录性质 = 4 And b.记录状态 = 3 And
            Nvl(b.冲预交, 0) <> 0;
  
    n_执行状态       病人挂号记录.执行状态%Type;
    n_打印id         票据打印内容.Id%Type;
    n_结帐id         门诊费用记录.结帐id%Type;
    n_原结帐id       病人预交记录.结帐id%Type;
    n_病人id         病人信息.病人id%Type;
    n_返回值         病人余额.预交余额%Type;
    n_分诊台签到排队 Number;
    n_预约挂号       Number;
    n_无效单据       Number; --无效单据没有产生费用单据
    n_挂号生成队列   Number;
    n_Count          Number;
    n_组id           财务缴款分组.Id%Type;
    d_退号时间       Date;
    v_操作员编号     人员表.编号%Type;
    v_操作员姓名     人员表.姓名%Type;
    v_合作单位       合作单位挂号汇总.合作单位%Type;
    n_预约状态       病人挂号记录.预约%Type;
    v_Temp           Varchar2(100);
    d_登记时间       病人挂号记录.登记时间%Type;
    v_号别           病人挂号记录.号别%Type;
    n_号序           病人挂号记录.号序%Type;
    d_预约时间       病人挂号记录.预约时间%Type;
    n_合作单位限制   Number(18);
    n_预约生成队列   Number;
    n_记录性质       Number;
    n_状态           Number;
    n_退号重用       Number(3);
    n_挂号id         病人挂号记录.Id%Type;
    n_记帐           门诊费用记录.记帐费用%Type;
    n_记录id         临床出诊记录.Id%Type;
  
    n_已结帐   Number;
    n_返回额   病人余额.费用余额%Type;
    n_预交支付 Number(3);
    n_正常支付 Number(3);
    Function Zl_操作员
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-获取缺省部门ID;1-获取操作员编号;2-获取操作员姓名
      -- SplitStr:格式为:部门ID,部门名称;人员ID,人员编号,人员姓名(用Zl_Identity获取的)
    Begin
      If Type_In = 0 Then
        --缺省部门
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --操作员编码
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --操作员姓名
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  Begin
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Custom;
    End If;
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
  
    n_组id := Zl_Get组id(v_操作员姓名);
  
    d_退号时间 := 退号时间_In;
    If d_退号时间 Is Null Then
      d_退号时间 := Sysdate;
    End If;
  
    --首先判断要退号/取消预约的记录是否存在
    Begin
      Select Decode(记录性质, 2, 1, 0), 记录性质, 登记时间, 号别, 号序, Nvl(预约时间, 发生时间), 合作单位, Nvl(预约, 0), Decode(记录状态, 0, 1, 0), 出诊记录id
      Into n_预约挂号, n_记录性质, d_登记时间, v_号别, n_号序, d_预约时间, v_合作单位, n_预约状态, n_无效单据, n_记录id
      From 病人挂号记录
      Where NO = 单据号_In And 记录状态 In (0, 1) And Rownum < 2;
    Exception
      When Others Then
        n_预约挂号 := -1;
    End;
  
    If n_预约挂号 = -1 Then
      v_Error := '单据可能已经被退号或单据输入错误!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(记帐费用, 0), Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
      Into n_记帐, n_已结帐
      From 门诊费用记录
      Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        n_记帐   := 0;
        n_已结帐 := 0;
    End;
  
    --预约检查是否添加合作单位控制
    --如果设置了合作单位控制 则
    Select Count(0) Into n_合作单位限制 From 临床出诊挂号控制记录 Where 类型 = 1 And 性质 = 1 And Rownum < 2;
    --更新挂号序号状态
    n_退号重用 := Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111));
    If n_退号重用 = 0 Then
      Update 临床出诊序号控制 Set 挂号状态 = 4 Where 记录id = n_记录id And (序号 = n_号序 Or 备注 = To_Char(n_号序));
    Else
      Update 临床出诊序号控制
      Set 挂号状态 = 0, 类型 = Null, 名称 = Null, 操作员姓名 = Null, 工作站名称 = Null
      Where 记录id = n_记录id And (序号 = n_号序 Or 备注 = To_Char(n_号序));
    End If;
    If Nvl(n_预约挂号, 0) = 1 Or Nvl(n_无效单据, 0) = 1 Then
      If Nvl(n_无效单据, 0) = 0 Then
        --N天内不能取消预约号
        n_Count := Zl_To_Number(zl_GetSysParameter('N天内不能取消预约号', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_登记时间 Then
            v_Error := '不能退掉预约在' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '以前的预约单!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_状态 := Case n_无效单据
                When 1 Then
                 0
                Else
                 1
              End;
      --减少已约数
      Open c_Registinfo(n_状态, 2, n_无效单据);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
            Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
            (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数, 已挂数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -n_预约状态, Decode(n_预约状态, 0, 1, 0));
      End If;
    
      Update 临床出诊记录
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where ID = n_记录id;
      Close c_Registinfo;
    
      If Nvl(n_无效单据, 0) = 0 Then
        --删除门诊费用记录
        Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
        --如果预约生成队列时需要清除队列
        n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
        If Nvl(n_挂号生成队列, 0) = 1 Then
          n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
          If Nvl(n_预约生成队列, 0) = 1 Then
            --要删除队列
            For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
              Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    
      --更新挂号序号状态
    
      --病人就诊状态
      Select 病人id
      Into n_病人id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 序号 = 1;
    
      If n_病人id Is Not Null Then
        Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
        --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      End If;
    
      --门诊费用记录
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位,
         付数, 数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间,
         登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
        Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
               收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
               v_操作员编号, v_操作员姓名, 发生时间, d_退号时间, n_结帐id,
               Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额, 摘要,
               Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
    
      --原始记录
      If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
        Update 门诊费用记录
        Set 记录状态 = 3, 结帐id = n_结帐id, 结帐金额 = 实收金额
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      Else
        Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      End If;
    
      n_原结帐id := 0;
      If n_记帐 = 0 Then
        --获取结帐ID
        Select Nvl(结帐id, 0)
        Into n_原结帐id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum < 2;
      End If;
    
      If n_记帐 = 1 Then
        --记帐
        For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                     From 门诊费用记录
                     Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Nvl(记帐费用, 0) = 1) Loop
          --病人余额
          Update 病人余额
          Set 费用余额 = Nvl(费用余额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1
          Returning 费用余额 Into n_返回额;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人余额
              (病人id, 性质, 类型, 费用余额, 预交余额)
            Values
              (n_病人id, 1, 1, -1 * Nvl(c_费用.实收金额, 0), 0);
            n_返回额 := Nvl(c_费用.实收金额, 0);
          End If;
          If Nvl(n_返回额, 0) = 0 Then
            Delete 病人余额
            Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
          --病人未结费用
          Update 病人未结费用
          Set 金额 = Nvl(金额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
                Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And
                收入项目id + 0 = c_费用.收入项目id And 来源途径 + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人未结费用
              (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
            Values
              (n_病人id, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, -1 * Nvl(c_费用.实收金额, 0));
          End If;
        End Loop;
        Delete 病人未结费用
        Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(金额, 0) = 0 And 来源途径 + 0 = 1;
      End If;
    
      If n_记帐 = 0 Then
        Begin
          Select 1
          Into n_预交支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_预交支付 := 0;
        End;
        Begin
          Select 1
          Into n_正常支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 4 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_正常支付 := 0;
        End;
        If n_预交支付 = 1 And n_正常支付 = 1 Then
          v_Error := '不能处理多种结算方式,请检查传入的退号单据是否正确!';
          Raise Err_Custom;
        End If;
        If n_预交支付 = 1 Then
          --原样退回预交
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                     d_退号时间, v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 预交id_In, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_退号时间,
                     v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          End If;
          --处理病人预交余额
          For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                       From 病人预交记录
                       Where 记录性质 In (1, 11) And 结帐id = n_结帐id
                       Group By 病人id, 预交类别) Loop
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.冲预交, 0)
            Where 病人id = c_预交.病人id And 类型 = Nvl(c_预交.预交类别, 2) And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 预交余额, 性质, 类型)
              Values
                (c_预交.病人id, Nvl(c_预交.冲预交, 0), 1, Nvl(c_预交.预交类别, 2));
              n_返回值 := Nvl(c_预交.冲预交, 0);
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End Loop;
        Else
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交,
                     n_结帐id, n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 预交id_In, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交, n_结帐id,
                     n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          End If;
          Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
        End If;
        --退卡收回票据(可能上次挂号使用票据,不能收回)
        --从最后一次的打印内容中取
        Select Max(ID)
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
        If n_打印id Is Not Null Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_退号时间, v_操作员姓名, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1;
        End If;
      End If;
    
      --相关汇总表的处理
    
      --病人挂号汇总
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --只收病历费时无号别,不处理
        Close c_Registinfo;
      Else
      
        --需要确定是否预约挂号
        --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
        --2.如果是正常挂号,则只减已挂数
        Begin
          Select Decode(预约, Null, 0, 0, 0, 1), 执行状态
          Into n_预约挂号, n_执行状态
          From 病人挂号记录
          Where NO = 单据号_In And 记录状态 = 1 And Rownum = 1;
        Exception
          When Others Then
            n_预约挂号 := 0;
        End;
        --0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
        If n_执行状态 > 0 Then
          If n_执行状态 = 1 Then
            v_Error := '该病人已经完成就诊,不能再退号!';
          Else
            v_Error := '该病人正在就诊, 不能退号!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
              Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
              (号码 = r_Registrow.号码 Or 号码 Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
             Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号);
        End If;
      
        Update 临床出诊记录
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where ID = n_记录id;
        Close c_Registinfo;
      End If;
    
      --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
      If n_记帐 = 0 Then
        For r_Opermoney In c_Opermoney Loop
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
          Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
          Returning 余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (v_操作员姓名, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
            n_返回值 := r_Opermoney.冲预交;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End Loop;
      End If;
    
      n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
      If n_挂号生成队列 <> 0 Then
        n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          --要删除队列
          For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
            Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
          End Loop;
        End If;
      End If;
    
      --医保产生的就诊登记记录
      Delete From 就诊登记记录
      Where (病人id, 主页id, 就诊时间) In (Select 病人id, 主页id, 发生时间 From 病人挂号记录 Where NO = 单据号_In);
    End If;
  
    If Nvl(n_无效单据, 0) = 0 Then
      Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1;
      If Sql%NotFound Then
        v_Error := '未找到挂号单据,请检查!';
        Raise Err_Custom;
      End If;
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位, 出诊记录id)
        Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_退号时间, 发生时间,
               v_操作员编号, v_操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号_In, 交易说明_In, 合作单位, 出诊记录id
        From 病人挂号记录
        Where NO = 单据号_In And 记录状态 = 3;
    End If;
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 2, 单据号_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_挂号排班模式 := To_Number(Substr(Nvl(zl_GetSysParameter('挂号排班模式'), 0), 1, 1));
  If n_挂号排班模式 = 1 Then
    --出诊表排班模式
    Zl_三方机构挂号_出诊_Delete(单据号_In, 交易流水号_In, 交易说明_In, 退号时间_In, 预交id_In);
  Else
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
    End If;
    --部门ID,部门名称;人员ID,人员编号,人员姓名
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '当前操作人员未设置对应的人员关系,不能继续。';
      Raise Err_Custom;
    End If;
    v_操作员编号 := Zl_操作员(1, v_Temp);
    v_操作员姓名 := Zl_操作员(2, v_Temp);
  
    n_组id := Zl_Get组id(v_操作员姓名);
  
    d_退号时间 := 退号时间_In;
    If d_退号时间 Is Null Then
      d_退号时间 := Sysdate;
    End If;
  
    --首先判断要退号/取消预约的记录是否存在
    Begin
      Select Decode(记录性质, 2, 1, 0), 记录性质, 登记时间, 号别, 号序, Nvl(预约时间, 发生时间), 合作单位, Nvl(预约, 0), Decode(记录状态, 0, 1, 0)
      Into n_预约挂号, n_记录性质, d_登记时间, v_号别, n_号序, d_预约时间, v_合作单位, n_预约状态, n_无效单据
      From 病人挂号记录
      Where NO = 单据号_In And 记录状态 In (0, 1) And Rownum <= 1;
    Exception
      When Others Then
        n_预约挂号 := -1;
    End;
  
    If n_预约挂号 = -1 Then
      v_Error := '单据可能已经被退号或单据输入错误!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(记帐费用, 0), Decode(Sign(Nvl(结帐id, 0)), 0, 0, 1)
      Into n_记帐, n_已结帐
      From 门诊费用记录
      Where 记录性质 = 4 And NO = 单据号_In And 记录状态 In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        n_记帐   := 0;
        n_已结帐 := 0;
    End;
  
    Begin
      Select a.Id Into n_安排id From 挂号安排 A Where a.号码 = v_号别;
    Exception
      When Others Then
        n_安排id := -1;
    End;
  
    Begin
      Select ID
      Into n_计划id
      From 挂号安排计划
      Where 安排id = n_安排id And 审核时间 Is Not Null And
            Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.生效时间) As 生效
             From 挂号安排计划 A
             Where a.审核时间 Is Not Null And d_预约时间 Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   Nvl(a.失效时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.安排id = n_安排id) And
            d_预约时间 Between Nvl(生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And
            Nvl(失效时间, To_Date('3000-01-01', 'yyyy-mm-dd'));
    Exception
      When Others Then
        n_计划id := 0;
    End;
    If Nvl(n_计划id, 0) = 0 Then
      Select Decode(To_Char(d_预约时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排
      Where ID = n_安排id;
    Else
      Select Decode(To_Char(d_预约时间, 'D'), '1', 周日, '2', 周一, '3', 周二, '4', 周三, '5', 周四, '6', 周五, '7', 周六, Null)
      Into v_时间段
      From 挂号安排计划
      Where ID = n_计划id;
    End If;
  
    If v_时间段 Is Not Null And d_启用时间 Is Not Null Then
      --检查是否跨模式挂号安排
      Select To_Date(To_Char(d_预约时间, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(d_预约时间, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_检查开始时间, d_检查结束时间
      From 时间段
      Where 时间段 = v_时间段 And 站点 Is Null And 号类 Is Null;
      If d_检查开始时间 > d_检查结束时间 Then
        d_检查结束时间 := d_检查结束时间 + 1;
      End If;
      If d_检查结束时间 > d_启用时间 Then
        --获取出诊记录id
        Select Max(a.Id)
        Into n_出诊记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = v_号别 And 上班时段 = v_时间段 And d_预约时间 Between 开始时间 And 终止时间;
      End If;
    End If;
  
    --预约检查是否添加合作单位控制
    --如果设置了合作单位控制 则
    Select Count(0) Into n_合作单位限制 From 合作单位安排控制 Where Rownum = 1;
    --更新挂号序号状态
    n_退号重用 := Zl_To_Number(zl_GetSysParameter('已退序号允许挂号', 1111));
    If n_退号重用 = 0 Then
      Update 挂号序号状态
      Set 状态 = 4
      Where 号码 = v_号别 And 序号 = n_号序 And 日期 Between Trunc(d_预约时间) And Trunc(d_预约时间 + 1) - 1 / 24 / 60 / 60;
    Else
      Delete 挂号序号状态
      Where 号码 = v_号别 And 序号 = n_号序 And 日期 Between Trunc(d_预约时间) And Trunc(d_预约时间 + 1) - 1 / 24 / 60 / 60;
    End If;
    If Nvl(n_预约挂号, 0) = 1 Or Nvl(n_无效单据, 0) = 1 Then
      If Nvl(n_无效单据, 0) = 0 Then
        --N天内不能取消预约号
        n_Count := Zl_To_Number(zl_GetSysParameter('N天内不能取消预约号', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_登记时间 Then
            v_Error := '不能退掉预约在' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '以前的预约单!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_状态 := Case n_无效单据
                When 1 Then
                 0
                Else
                 1
              End;
      --减少已约数
      Open c_Registinfo(n_状态, 2, n_无效单据);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update 病人挂号汇总
      Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
      Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
            Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
            (号码 = r_Registrow.号码 Or 号码 Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into 病人挂号汇总
          (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已约数, 已挂数)
        Values
          (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
           Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -n_预约状态, Decode(n_预约状态, 0, 1, 0));
      End If;
    
      If n_出诊记录id Is Not Null Then
        Update 临床出诊记录
        Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已挂数 = Nvl(已挂数, 0) - Decode(n_预约状态, 0, 1, 0)
        Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
        Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_号序;
      End If;
      If Nvl(n_合作单位限制, 0) <> 0 And v_合作单位 Is Not Null And Nvl(n_预约状态, 0) <> 0 Then
        Update 合作单位挂号汇总
        Set 已约数 = Nvl(已约数, 0) - n_预约状态, 已接数 = Nvl(已接数, 0) - Decode(n_预约状态, 0, 1, 0)
        Where 日期 = Trunc(r_Registrow.发生时间) And (号码 = r_Registrow.号码 Or 号码 Is Null) And 合作单位 = v_合作单位 And
              序号 = Nvl(n_号序, 0);
        If Sql%RowCount = 0 Then
          Insert Into 合作单位挂号汇总
            (日期, 号码, 已约数, 合作单位, 序号, 已接数)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.号码, -n_预约状态, v_合作单位, Nvl(n_号序, 0), -decode(n_预约状态, 0, 1, 0));
        End If;
      End If;
      Close c_Registinfo;
    
      If Nvl(n_无效单据, 0) = 0 Then
        --删除门诊费用记录
        Delete From 门诊费用记录 Where NO = 单据号_In And 记录性质 = 4 And 记录状态 = 0;
        --如果预约生成队列时需要清除队列
        n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
        If Nvl(n_挂号生成队列, 0) = 1 Then
          n_预约生成队列 := Zl_To_Number(zl_GetSysParameter('预约生成队列', 1113));
          If Nvl(n_预约生成队列, 0) = 1 Then
            --要删除队列
            For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
              Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    
      --更新挂号序号状态
    
      --病人就诊状态
      Select 病人id
      Into n_病人id
      From 门诊费用记录
      Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In And 序号 = 1;
    
      If n_病人id Is Not Null Then
        Update 病人信息 Set 就诊状态 = 0, 就诊诊室 = Null Where 病人id = n_病人id;
        --删除门诊号相关处理,只有当只有一条挂号记录并且病人建档日期与挂号日期近似时才会处理
      End If;
    
      --门诊费用记录
      Insert Into 门诊费用记录
        (ID, NO, 实际票号, 记录性质, 记录状态, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别, 收费细目id, 计算单位,
         付数, 数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, 应收金额, 实收金额, 开单部门id, 开单人, 执行部门id, 执行人, 操作员编号, 操作员姓名, 发生时间,
         登记时间, 结帐id, 结帐金额, 保险项目否, 保险大类id, 统筹金额, 摘要, 是否上传, 保险编码, 费用类型, 缴款组id)
        Select 病人费用记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 序号, 价格父号, 从属父号, 病人id, 病人科室id, 门诊标志, 标识号, 付款方式, 姓名, 性别, 年龄, 费别, 收费类别,
               收费细目id, 计算单位, 付数, -数次, 加班标志, 附加标志, 发药窗口, 收入项目id, 收据费目, 记帐费用, 标准单价, -应收金额, -实收金额, 开单部门id, 开单人, 执行部门id, 执行人,
               v_操作员编号, v_操作员姓名, 发生时间, d_退号时间, n_结帐id,
               Decode(n_记帐, 1, Decode(Nvl(n_已结帐, 0), 0, -1 * 实收金额, Null), -1 * 结帐金额), 保险项目否, 保险大类id, -1 * 统筹金额, 摘要,
               Decode(Nvl(附加标志, 0), 9, 1, 0), 保险编码, 费用类型, n_组id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
    
      --原始记录
      If n_记帐 = 1 And Nvl(n_已结帐, 0) = 0 Then
        Update 门诊费用记录
        Set 记录状态 = 3, 结帐id = n_结帐id, 结帐金额 = 实收金额
        Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      Else
        Update 门诊费用记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And NO = 单据号_In;
      End If;
    
      n_原结帐id := 0;
      If n_记帐 = 0 Then
        --获取结帐ID
        Select Nvl(结帐id, 0)
        Into n_原结帐id
        From 门诊费用记录
        Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Rownum < 2;
      End If;
    
      If n_记帐 = 1 Then
        --记帐
        For c_费用 In (Select 实收金额, 病人科室id, 开单部门id, 执行部门id, 收入项目id
                     From 门诊费用记录
                     Where 记录性质 = 4 And 记录状态 = 3 And NO = 单据号_In And Nvl(记帐费用, 0) = 1) Loop
          --病人余额
          Update 病人余额
          Set 费用余额 = Nvl(费用余额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1
          Returning 费用余额 Into n_返回额;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人余额
              (病人id, 性质, 类型, 费用余额, 预交余额)
            Values
              (n_病人id, 1, 1, -1 * Nvl(c_费用.实收金额, 0), 0);
            n_返回额 := Nvl(c_费用.实收金额, 0);
          End If;
          If Nvl(n_返回额, 0) = 0 Then
            Delete 病人余额
            Where 病人id = Nvl(n_病人id, 0) And 性质 = 1 And 类型 = 1 And Nvl(费用余额, 0) = 0 And Nvl(预交余额, 0) = 0;
          End If;
          --病人未结费用
          Update 病人未结费用
          Set 金额 = Nvl(金额, 0) - Nvl(c_费用.实收金额, 0)
          Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(病人科室id, 0) = Nvl(c_费用.病人科室id, 0) And
                Nvl(开单部门id, 0) = Nvl(c_费用.开单部门id, 0) And Nvl(执行部门id, 0) = Nvl(c_费用.执行部门id, 0) And
                收入项目id + 0 = c_费用.收入项目id And 来源途径 + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into 病人未结费用
              (病人id, 主页id, 病人病区id, 病人科室id, 开单部门id, 执行部门id, 收入项目id, 来源途径, 金额)
            Values
              (n_病人id, Null, Null, c_费用.病人科室id, c_费用.开单部门id, c_费用.执行部门id, c_费用.收入项目id, 1, -1 * Nvl(c_费用.实收金额, 0));
          End If;
        End Loop;
        Delete 病人未结费用
        Where 病人id = n_病人id And Nvl(主页id, 0) = 0 And Nvl(病人病区id, 0) = 0 And Nvl(金额, 0) = 0 And 来源途径 + 0 = 1;
      End If;
    
      If n_记帐 = 0 Then
        Begin
          Select 1
          Into n_预交支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 1 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_预交支付 := 0;
        End;
      
        Begin
          Select 1
          Into n_正常支付
          From 病人预交记录
          Where Mod(记录性质, 10) = 4 And 记录状态 = 1 And 结帐id = n_原结帐id And Rownum < 2;
        Exception
          When Others Then
            n_正常支付 := 0;
        End;
      
        If n_预交支付 = 1 And n_正常支付 = 1 Then
          v_Error := '不能处理多种结算方式,请检查传入的退号单据是否正确!';
          Raise Err_Custom;
        End If;
      
        If n_预交支付 = 1 Then
          --原样退回预交
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号,
                     d_退号时间, v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 金额, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, 收款时间, 操作员姓名, 操作员编号,
               冲预交, 结帐id, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算性质)
              Select 预交id_In, NO, 实际票号, 11, 记录状态, 病人id, 主页id, 科室id, Null, 结算方式, 结算号码, 摘要, 缴款单位, 单位开户行, 单位帐号, d_退号时间,
                     v_操作员姓名, v_操作员编号, -1 * 冲预交, n_结帐id, n_组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 4
              From 病人预交记录
              Where 记录性质 In (1, 11) And 结帐id = n_原结帐id And Nvl(冲预交, 0) <> 0;
          End If;
          --处理病人预交余额
          For c_预交 In (Select 病人id, 预交类别, -1 * Sum(Nvl(冲预交, 0)) As 冲预交
                       From 病人预交记录
                       Where 记录性质 In (1, 11) And 结帐id = n_结帐id
                       Group By 病人id, 预交类别) Loop
            Update 病人余额
            Set 预交余额 = Nvl(预交余额, 0) + Nvl(c_预交.冲预交, 0)
            Where 病人id = c_预交.病人id And 类型 = Nvl(c_预交.预交类别, 2) And 性质 = 1
            Returning 预交余额 Into n_返回值;
            If Sql%RowCount = 0 Then
              Insert Into 病人余额
                (病人id, 预交余额, 性质, 类型)
              Values
                (c_预交.病人id, Nvl(c_预交.冲预交, 0), 1, Nvl(c_预交.预交类别, 2));
              n_返回值 := Nvl(c_预交.冲预交, 0);
            End If;
            If Nvl(n_返回值, 0) = 0 Then
              Delete From 病人余额
              Where 病人id = c_预交.病人id And 性质 = 1 And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0;
            End If;
          End Loop;
        Else
          If Nvl(预交id_In, 0) = 0 Then
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 病人预交记录_Id.Nextval, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交,
                     n_结帐id, n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          Else
            Insert Into 病人预交记录
              (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 科室id, 摘要, 结算方式, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款组id, 交易流水号, 交易说明,
               合作单位, 结算序号, 卡类别id, 结算性质)
              Select 预交id_In, NO, 实际票号, 记录性质, 2, 病人id, 主页id, 科室id, 摘要, 结算方式, d_退号时间, v_操作员编号, v_操作员姓名, -冲预交, n_结帐id,
                     n_组id, 交易流水号_In, 交易说明_In, 合作单位, n_结帐id, 卡类别id, 结算性质
              From 病人预交记录
              Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
          End If;
        
          Update 病人预交记录 Set 记录状态 = 3 Where 记录性质 = 4 And 记录状态 = 1 And 结帐id = n_原结帐id;
        End If;
        --退卡收回票据(可能上次挂号使用票据,不能收回)
        --从最后一次的打印内容中取
        Select Max(ID)
        Into n_打印id
        From (Select b.Id
               From 票据使用明细 A, 票据打印内容 B
               Where a.打印id = b.Id And a.性质 = 1 And b.数据性质 = 4 And b.No = 单据号_In
               Order By a.使用时间 Desc)
        Where Rownum < 2;
        If n_打印id Is Not Null Then
          Insert Into 票据使用明细
            (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
            Select 票据使用明细_Id.Nextval, 票种, 号码, 2, 2, 领用id, 打印id, d_退号时间, v_操作员姓名, 票据金额
            From 票据使用明细
            Where 打印id = n_打印id And 性质 = 1;
        End If;
      End If;
    
      --相关汇总表的处理
    
      --病人挂号汇总
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --只收病历费时无号别,不处理
        Close c_Registinfo;
      Else
      
        --需要确定是否预约挂号
        --1.如果是退预约挂号产生的挂号记录,则需要减已挂数和其中已接数
        --2.如果是正常挂号,则只减已挂数
        Begin
          Select Decode(预约, Null, 0, 0, 0, 1), 执行状态
          Into n_预约挂号, n_执行状态
          From 病人挂号记录
          Where NO = 单据号_In And 记录状态 = 1 And Rownum = 1;
        Exception
          When Others Then
            n_预约挂号 := 0;
        End;
        --0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
        If n_执行状态 > 0 Then
          If n_执行状态 = 1 Then
            v_Error := '该病人已经完成就诊,不能再退号!';
          Else
            v_Error := '该病人正在就诊, 不能退号!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update 病人挂号汇总
        Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
        Where 日期 = Trunc(r_Registrow.发生时间) And 科室id = r_Registrow.科室id And 项目id = r_Registrow.项目id And
              Nvl(医生姓名, '医生') = Nvl(r_Registrow.医生姓名, '医生') And Nvl(医生id, 0) = Nvl(r_Registrow.医生id, 0) And
              (号码 = r_Registrow.号码 Or 号码 Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into 病人挂号汇总
            (日期, 科室id, 项目id, 医生姓名, 医生id, 号码, 已挂数, 其中已接收)
          Values
            (Trunc(r_Registrow.发生时间), r_Registrow.科室id, r_Registrow.项目id, r_Registrow.医生姓名,
             Decode(r_Registrow.医生id, 0, Null, r_Registrow.医生id), r_Registrow.号码, -1, -1 * n_预约挂号);
        End If;
        If n_出诊记录id Is Not Null Then
          Update 临床出诊记录
          Set 已挂数 = Nvl(已挂数, 0) - 1, 其中已接收 = Nvl(其中已接收, 0) - n_预约状态, 已约数 = Nvl(已约数, 0) - n_预约状态
          Where ID = n_出诊记录id And Nvl(已约数, 0) > 0;
          Update 临床出诊序号控制 Set 挂号状态 = Null, 操作员姓名 = Null Where 记录id = n_出诊记录id And 序号 = n_号序;
        End If;
        If Nvl(n_合作单位限制, 0) <> 0 And v_合作单位 Is Not Null And Nvl(n_预约状态, 0) <> 0 Then
          Update 合作单位挂号汇总
          Set 已接数 = Nvl(已接数, 0) - 1, 已约数 = Nvl(已约数, 0) - n_预约挂号
          Where 日期 = Trunc(r_Registrow.发生时间) And (号码 = r_Registrow.号码 Or 号码 Is Null) And 合作单位 = v_合作单位 And
                序号 = Nvl(n_号序, 0);
          If Sql%RowCount = 0 Then
            Insert Into 合作单位挂号汇总
              (日期, 号码, 已约数, 合作单位, 已接数, 序号)
            Values
              (Trunc(r_Registrow.发生时间), r_Registrow.号码, -1, v_合作单位, -1 * n_预约挂号, Nvl(n_号序, 0));
          End If;
        End If;
        Close c_Registinfo;
      End If;
    
      --人员缴款余额(包括个人帐户等的结算金额,不含退冲预交款)
      If n_记帐 = 0 Then
        For r_Opermoney In c_Opermoney Loop
          Update 人员缴款余额
          Set 余额 = Nvl(余额, 0) + (-1 * r_Opermoney.冲预交)
          Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1
          Returning 余额 Into n_返回值;
          If Sql%RowCount = 0 Then
            Insert Into 人员缴款余额
              (收款员, 结算方式, 性质, 余额)
            Values
              (v_操作员姓名, r_Opermoney.结算方式, 1, -1 * r_Opermoney.冲预交);
            n_返回值 := r_Opermoney.冲预交;
          End If;
          If Nvl(n_返回值, 0) = 0 Then
            Delete From 人员缴款余额
            Where 收款员 = v_操作员姓名 And 结算方式 = r_Opermoney.结算方式 And 性质 = 1 And Nvl(余额, 0) = 0;
          End If;
        End Loop;
      End If;
    
      n_挂号生成队列 := Zl_To_Number(zl_GetSysParameter('排队叫号模式', 1113));
      If n_挂号生成队列 <> 0 Then
        n_分诊台签到排队 := Zl_To_Number(zl_GetSysParameter('分诊台签到排队', 1113));
        If Nvl(n_分诊台签到排队, 0) = 0 Then
          --要删除队列
          For v_挂号 In (Select ID, 诊室, 执行部门id, 执行人 From 病人挂号记录 Where NO = 单据号_In) Loop
            Zl_排队叫号队列_Delete(v_挂号.执行部门id, v_挂号.Id);
          End Loop;
        End If;
      End If;
    
      --医保产生的就诊登记记录
      Delete From 就诊登记记录
      Where (病人id, 主页id, 就诊时间) In (Select 病人id, 主页id, 发生时间 From 病人挂号记录 Where NO = 单据号_In);
    End If;
  
    If Nvl(n_无效单据, 0) = 0 Then
      Update 病人挂号记录 Set 记录状态 = 3 Where NO = 单据号_In And 记录状态 = 1;
      If Sql%NotFound Then
        v_Error := '未找到挂号单据,请检查!';
        Raise Err_Custom;
      End If;
      Select 病人挂号记录_Id.Nextval Into n_挂号id From Dual;
      Insert Into 病人挂号记录
        (ID, NO, 记录性质, 记录状态, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, 登记时间, 发生时间, 操作员编号, 操作员姓名,
         复诊, 号序, 预约, 接收人, 接收时间, 交易流水号, 交易说明, 合作单位)
        Select n_挂号id, NO, 记录性质, 2, 病人id, 门诊号, 姓名, 性别, 年龄, 号别, 急诊, 诊室, 附加标志, 执行部门id, 执行人, 执行状态, 执行时间, d_退号时间, 发生时间,
               v_操作员编号, v_操作员姓名, 复诊, 号序, 预约, 接收人, 接收时间, 交易流水号_In, 交易说明_In, 合作单位
        From 病人挂号记录
        Where NO = 单据号_In And 记录状态 = 3;
    End If;
    --消息推送
    Begin
      Execute Immediate 'Begin ZL_服务窗消息_发送(:1,:2); End;'
        Using 2, 单据号_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_挂号id, 单据号_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方机构挂号_Delete;
/

--125588:秦龙,2018-05-14,处理库存实际数量为0的数据
Create Or Replace Procedure Zl_草药规格_Update
(
  药品id_In         In 药品规格.药品id%Type,
  编码_In           In 收费项目目录.编码%Type,
  规格_In           In 收费项目目录.规格%Type,
  产地_In           In 收费项目目录.产地%Type := Null,
  品名_In           In 收费项目别名.名称%Type := Null,
  拼音_In           In 收费项目别名.简码%Type := Null,
  五笔_In           In 收费项目别名.简码%Type := Null,
  数字码_In         In 收费项目别名.简码%Type := Null,
  标识码_In         In 药品规格.标识码%Type := Null,
  药品来源_In       In 药品规格.药品来源%Type := Null,
  批准文号_In       In 药品规格.批准文号%Type := Null,
  注册商标_In       In 药品规格.注册商标%Type := Null,
  售价单位_In       In 收费项目目录.计算单位%Type := Null,
  剂量系数_In       In 药品规格.剂量系数%Type := Null,
  门诊单位_In       In 药品规格.门诊单位%Type := Null,
  门诊包装_In       In 药品规格.门诊包装%Type := Null,
  药库单位_In       In 药品规格.药库单位%Type := Null,
  药库包装_In       In 药品规格.药库包装%Type := Null,
  申领单位_In       In 药品规格.申领单位%Type := 1,
  申领阀值_In       In 药品规格.申领阀值%Type := Null,
  是否变价_In       In 收费项目目录.是否变价%Type := Null,
  指导批发价_In     In 药品规格.指导批发价%Type := Null,
  扣率_In           In 药品规格.扣率%Type := 95,
  指导零售价_In     In 药品规格.指导零售价%Type := Null,
  加成率_In         In 药品规格.加成率%Type := Null,
  管理费比例_In     In 药品规格.管理费比例%Type := Null,
  药价级别_In       In 药品规格.药价级别%Type := Null,
  费用类型_In       In 收费项目目录.费用类型%Type := Null,
  服务对象_In       In 收费项目目录.服务对象%Type := Null,
  Gmp认证_In        In 药品规格.Gmp认证%Type := 0,
  招标药品_In       In 药品规格.招标药品%Type := 0,
  屏蔽费别_In       In 收费项目目录.屏蔽费别%Type := 0,
  住院可否分零_In   In 药品规格.住院可否分零%Type := 0,
  药库分批_In       In 药品规格.药库分批%Type := Null,
  药房分批_In       In 药品规格.药房分批%Type := Null,
  最大效期_In       In 药品规格.最大效期%Type := Null,
  差价让利比_In     In 药品规格.差价让利比%Type := 0,
  成本价_In         In 药品规格.成本价%Type := 0,
  当前售价_In       In 收费价目.现价%Type := 0,
  收入id_In         In 收费价目.收入项目id%Type := Null,
  合同单位id_In     In 药品规格.合同单位id%Type := Null,
  说明_In           In 收费项目目录.说明%Type := Null,
  动态分零_In       In 药品规格.动态分零%Type := 0,
  发药类型_In       In 药品规格.发药类型%Type := Null,
  备选码_In         In 收费项目目录.备选码%Type := Null,
  增值税率_In       In 药品规格.增值税率%Type := Null,
  基本药物_In       In 药品规格.基本药物%Type := Null,
  中药形态_In       In 药品规格.中药形态%Type := Null,
  站点_In           In 收费项目目录.站点%Type := Null,
  是否常备_In       In 药品规格.是否常备%Type := Null,
  病案费目_In       In 收费项目目录.病案费目%Type := Null,
  门诊可否分零_In   In 药品规格.门诊可否分零%Type := 0,
  送货单位_In       药品规格.送货单位%Type := Null,
  送货包装_In       药品规格.送货包装%Type := Null,
  是否摆药_In       药品规格.是否摆药%Type := Null,
  是否零差价管理_In In 药品规格.是否零差价管理%Type := Null,
  本位码_In         In 药品规格.本位码%Type := Null,
  原产地_In         In 药品规格.原产地%Type := Null
) Is
  v_药名id   诊疗项目目录.Id%Type;
  v_名称     诊疗项目目录.名称%Type;
  v_是否变价 收费项目目录.是否变价%Type; --允许定价药品随时改为时价，时价药品只能在未发生的情况下修改为定价，其它情况不允许修改定价属性 
  v_发生     Number(2);
  Err_Notfind Exception;
  v_No           收费价目.No%Type;
  v_Temp         收费项目目录.病案费目%Type;
  v_病案费目     收费项目目录.病案费目%Type;
  n_指导差价率   药品规格.指导差价率%Type;
  n_药品上次售价 药品规格.上次售价%Type;
  n_零售金额     药品库存.实际金额%Type;
  n_收发id       药品收发记录.Id%Type;
  n_流通金额小数 Number;
  n_序号         Number(8);
  Classid        Number(18); --入出类别
  v_Billno       药品收发记录.No%Type; --调价单号
  n_价格id       收费价目.Id%Type;
  n_收费价目现价 收费价目.现价%Type;
  n_原价         药品价格记录.原价%Type;
  n_药品价格记录 Number(1);
  v_类别         收费项目目录.类别%Type;
  --定价->时价后更新药品价格记录的值

  Cursor c_Priceadjust Is
    Select s.药品id, s.库房id, Nvl(s.批次, 0) As 批次, s.上次供应商id As 供应商id, s.上次批号 As 批号, s.效期, s.上次产地 As 产地,
           Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额, Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价,
           s.平均成本价, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期
    From 药品库存 S
    Where s.药品id = 药品id_In And s.性质 = 1 
    Order By s.药品id, s.批次, s.库房id;

  r_Priceadjust c_Priceadjust%RowType;

Begin
  v_病案费目 := 病案费目_In;
  --判断病案费目 
  If v_病案费目 Is Null Then
    If 收入id_In Is Not Null Then
      Begin
        Select 病案费目 Into v_Temp From 收入项目 Where ID = 收入id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_病案费目 := v_Temp;
      End If;
    End If;
  End If;
  --通用名称 
  Select ID, 名称
  Into v_药名id, v_名称
  From 诊疗项目目录
  Where ID = (Select 药名id From 药品规格 Where 药品id = 药品id_In);
  --取原始的定价属性 
  Select 是否变价 Into v_是否变价 From 收费项目目录 Where ID = 药品id_In;
  --规格信息 
  Update 收费项目目录
  Set 编码 = 编码_In, 名称 = v_名称, 规格 = 规格_In, 产地 = 产地_In, 计算单位 = 售价单位_In, 费用类型 = 费用类型_In, 服务对象 = 服务对象_In, 屏蔽费别 = 屏蔽费别_In,
      病案费目 = v_病案费目, 说明 = 说明_In, 备选码 = 备选码_In, 站点 = 站点_In
  Where ID = 药品id_In
  Returning 类别 Into v_类别;
  If Sql%RowCount = 0 Then
    Raise Err_Notfind;
  End If;
  n_指导差价率 := (1 - 1 / (1 + 加成率_In / 100)) * 100;
  Update 药品规格
  Set 标识码 = 标识码_In, 药品来源 = 药品来源_In, 批准文号 = 批准文号_In, 注册商标 = 注册商标_In, 剂量系数 = 剂量系数_In, 门诊单位 = 门诊单位_In, 门诊包装 = 门诊包装_In,
      住院单位 = 门诊单位_In, 住院包装 = 门诊包装_In, 药库单位 = 药库单位_In, 药库包装 = 药库包装_In, 申领单位 = 申领单位_In, 申领阀值 = 申领阀值_In, 指导批发价 = 指导批发价_In,
      扣率 = 扣率_In, 指导零售价 = 指导零售价_In, 指导差价率 = n_指导差价率, 管理费比例 = 管理费比例_In, 药价级别 = 药价级别_In, 住院可否分零 = 住院可否分零_In,
      药库分批 = 药库分批_In, 药房分批 = 药房分批_In, 最大效期 = 最大效期_In, 招标药品 = 招标药品_In, Gmp认证 = Gmp认证_In, 差价让利比 = 差价让利比_In,
      合同单位id = 合同单位id_In, 动态分零 = 动态分零_In, 发药类型 = 发药类型_In, 增值税率 = 增值税率_In, 基本药物 = 基本药物_In, 中药形态 = 中药形态_In, 是否常备 = 是否常备_In,
      门诊可否分零 = 门诊可否分零_In, 送货单位 = 送货单位_In, 送货包装 = 送货包装_In, 加成率 = 加成率_In, 是否摆药 = 是否摆药_In, 是否零差价管理 = 是否零差价管理_In,
      本位码 = 本位码_In, 原产地 = 原产地_In
  Where 药品id = 药品id_In;

  --朱玉宝修改：建立药品（西成药、中成药）时，缺省服务对象为门诊和住院，因此修改规格药品时，不再根据规格药品的服务对象更新药品的服务对象 
  --诊疗项目服务对象的更改 
  --select nvl(sum(distinct I.服务对象),0) into v_对象 
  --from 收费项目目录 I,药品规格 S 
  --where I.ID=S.药品ID and S.药名ID=v_药名ID; 
  --update 诊疗项目目录 
  --set 服务对象=decode(v_对象,0,0,1,1,2,2,3) 
  --where ID=v_药名ID; 

  --别名的处理 
  If 数字码_In Is Null Then
    Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 1 And 码类 = 3;
  Else
    Update 收费项目别名 Set 名称 = v_名称, 简码 = 数字码_In Where 收费细目id = 药品id_In And 性质 = 1 And 码类 = 3;
    If Sql%RowCount = 0 Then
      Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, v_名称, 1, 数字码_In, 3);
    End If;
  End If;
  If 品名_In Is Null Then
    Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3;
  Else
    If 拼音_In Is Null Then
      Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 1;
    Else
      Update 收费项目别名 Set 名称 = 品名_In, 简码 = 拼音_In Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 拼音_In, 1);
      End If;
    End If;
    If 五笔_In Is Null Then
      Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 2;
    Else
      Update 收费项目别名 Set 名称 = 品名_In, 简码 = 五笔_In Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 五笔_In, 2);
      End If;
    End If;
  End If;

  --定价信息：如果已经有发生，则不允许直接更改这些信息 
  Select Nvl(Count(*), 0) Into v_发生 From 药品收发记录 Where 药品id = 药品id_In And Rownum < 2;
  If v_发生 = 0 Then
    Update 药品规格 Set 成本价 = 成本价_In Where 药品id = 药品id_In;
    If 收入id_In Is Not Null Then
      Update 收费价目
      Set 现价 = 当前售价_In, 收入项目id = 收入id_In, 变动原因 = 1, 调价说明 = '修改定价', 调价人 = User
      Where 收费细目id = 药品id_In And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因 = 1;
      If Sql%RowCount = 0 Then
        v_No := Nextno(9);
        Insert Into 收费价目
          (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
        Values
          (收费价目_Id.Nextval, Null, 药品id_In, 0, 当前售价_In, 收入id_In, 1, '新增定价', User, Sysdate,
           To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
      End If;
    End If;
  Else
    --发生业务数据时，不能修改价格但是可以修改收入项目 
    Update 收费价目
    Set 收入项目id = 收入id_In
    Where 收费细目id = 药品id_In And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因 = 1;
  End If;

  --时价->定价
  If v_是否变价 = 1 And 是否变价_In = 0 Then
    Update 收费项目目录 Set 是否变价 = 是否变价_In Where ID = 药品id_In;
  
    Begin
      Select 现价, ID As 价格id
      Into n_收费价目现价, n_价格id
      From 收费价目
      Where 收费细目id = 药品id_In And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 变动原因 = 1;
    Exception
      When Others Then
        n_收费价目现价 := Null;
        n_价格id       := Null;
    End;
  
    Begin
      Select 上次售价 Into n_药品上次售价 From 药品规格 Where 药品id = 药品id_In;
    Exception
      When Others Then
        n_药品上次售价 := Null;
    End;
  
    If n_药品上次售价 Is Null Then
      n_药品上次售价 := n_收费价目现价;
    End If;
  
    Zl_收费价目_Stop(药品id_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, n_价格id, 药品id_In, n_收费价目现价, n_药品上次售价, 收入id_In, 1, '时价转定价', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;
  
    --取入出类别ID
    Select 类别id Into Classid From 药品单据性质 Where 单据 = 13;
  
    n_序号   := 0;
    v_Billno := Null;
  
    For r_Priceadjust In c_Priceadjust Loop
      If n_药品上次售价 <> r_Priceadjust.零售价 Then
        If v_Billno Is Null Then
          Select Nextno(147) Into v_Billno From Dual;
        End If;
        n_序号 := n_序号 + 1;
        Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
        n_零售金额 := Round(n_药品上次售价 * r_Priceadjust.实际数量, n_流通金额小数) -
                  Round(r_Priceadjust.零售价 * r_Priceadjust.实际数量, n_流通金额小数);
        --产生调价影响记录
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
           填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
        Values
          (n_收发id, 1, 13, v_Billno, n_序号, Classid, r_Priceadjust.药品id, r_Priceadjust.批次, r_Priceadjust.批号,
           r_Priceadjust.效期, r_Priceadjust.产地, 1, r_Priceadjust.实际数量, 0, r_Priceadjust.零售价, 0, n_药品上次售价,
           r_Priceadjust.扣率, n_零售金额, n_零售金额, '时价转定价', Zl_Username, Sysdate, r_Priceadjust.库房id, 1, n_价格id, Zl_Username,
           Sysdate, r_Priceadjust.实际金额, r_Priceadjust.实际差价, r_Priceadjust.供应商id);
      
        Zl_药品库存_Update(n_收发id, 2, 0);
      End If;
    End Loop;
  
    --定价->时价
  Elsif v_是否变价 = 0 And 是否变价_In = 1 Then
    Update 收费项目目录 Set 是否变价 = 是否变价_In Where ID = 药品id_In;
    Begin
      Select 现价, ID As 价格id
      Into n_收费价目现价, n_价格id
      From 收费价目
      Where 收费细目id = 药品id_In And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 变动原因 = 1;
    Exception
      When Others Then
        n_收费价目现价 := Null;
        n_价格id       := Null;
    End;
  
    Zl_收费价目_Stop(药品id_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, n_价格id, 药品id_In, n_收费价目现价, n_收费价目现价, 收入id_In, 1, '定价转时价', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    For r_Priceadjust In c_Priceadjust Loop
      n_药品价格记录 := 0;
      Begin
        Select 1, 现价
        Into n_药品价格记录, n_原价
        From 药品价格记录
        Where 药品id = r_Priceadjust.药品id And 库房id = r_Priceadjust.库房id And Nvl(批次, 0) = r_Priceadjust.批次 And 记录状态 = 1 And
              价格类型 = 1;
      Exception
        When Others Then
          n_药品价格记录 := 0;
          n_原价         := n_收费价目现价;
      End;
    
      If n_药品价格记录 = 1 Then
        Zl_药品价格记录_Stop(1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, Sysdate - 1 / 24 / 60 / 60, 2);
      End If;
      Zl_药品价格记录_Insert(0, 1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, n_原价, n_收费价目现价, Sysdate, '定价转时价',
                       Zl_Username, Null, r_Priceadjust.供应商id, r_Priceadjust.批号, r_Priceadjust.效期, r_Priceadjust.产地,
                       r_Priceadjust.灭菌效期, Null, Null, Null, Null, 1);
    
      Update 药品库存
      Set 零售价 = n_收费价目现价
      Where 性质 = 1 And 库房id = r_Priceadjust.库房id And 药品id = r_Priceadjust.药品id And Nvl(批次, 0) = r_Priceadjust.批次;
    
    End Loop;
  End If;

  --药品生产商比较增加 
  If 产地_In Is Not Null Then
    Update 药品生产商 Set 名称 = 产地_In Where 名称 = 产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 药品生产商
        (编码, 名称, 简码)
        Select Nvl(Max(To_Number(编码)), 0) + 1, 产地_In, zlSpellCode(产地_In, 10) From 药品生产商;
    End If;
  End If;

  --原产地较增加 
  If 原产地_In Is Not Null Then
    Update 药品生产商 Set 名称 = 原产地_In Where 名称 = 原产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 药品生产商
        (编码, 名称, 简码)
        Select Nvl(Max(To_Number(编码)), 0) + 1, 原产地_In, zlSpellCode(原产地_In, 10) From 药品生产商;
    End If;
  End If;

  --药品精度调整(零差价模式时)
  Zl_药品卫材精度_零差价调整;

  b_Message.Zlhis_Dict_036(v_类别, 药品id_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该规格不存在，可能已被其他用户删除！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_草药规格_Update;
/

--125588:秦龙,2018-05-14,处理库存实际数量为0的数据
Create Or Replace Procedure Zl_成药规格_Update
(
  药品id_In         In 药品规格.药品id%Type,
  编码_In           In 收费项目目录.编码%Type,
  规格_In           In 收费项目目录.规格%Type,
  产地_In           In 收费项目目录.产地%Type := Null,
  品名_In           In 收费项目别名.名称%Type := Null,
  拼音_In           In 收费项目别名.简码%Type := Null,
  五笔_In           In 收费项目别名.简码%Type := Null,
  数字码_In         In 收费项目别名.简码%Type := Null,
  标识码_In         In 药品规格.标识码%Type := Null,
  药品来源_In       In 药品规格.药品来源%Type := Null,
  批准文号_In       In 药品规格.批准文号%Type := Null,
  注册商标_In       In 药品规格.注册商标%Type := Null,
  售价单位_In       In 收费项目目录.计算单位%Type := Null,
  剂量系数_In       In 药品规格.剂量系数%Type := Null,
  门诊单位_In       In 药品规格.门诊单位%Type := Null,
  门诊包装_In       In 药品规格.门诊包装%Type := Null,
  住院单位_In       In 药品规格.住院单位%Type := Null,
  住院包装_In       In 药品规格.住院包装%Type := Null,
  药库单位_In       In 药品规格.药库单位%Type := Null,
  药库包装_In       In 药品规格.药库包装%Type := Null,
  申领单位_In       In 药品规格.申领单位%Type := 1,
  申领阀值_In       In 药品规格.申领阀值%Type := Null,
  是否变价_In       In 收费项目目录.是否变价%Type := Null,
  指导批发价_In     In 药品规格.指导批发价%Type := Null,
  扣率_In           In 药品规格.扣率%Type := 95,
  指导零售价_In     In 药品规格.指导零售价%Type := Null,
  加成率_In         In 药品规格.加成率%Type := Null,
  管理费比例_In     In 药品规格.管理费比例%Type := Null,
  药价级别_In       In 药品规格.药价级别%Type := Null,
  费用类型_In       In 收费项目目录.费用类型%Type := Null,
  服务对象_In       In 收费项目目录.服务对象%Type := Null,
  Gmp认证_In        In 药品规格.Gmp认证%Type := 0,
  招标药品_In       In 药品规格.招标药品%Type := 0,
  屏蔽费别_In       In 收费项目目录.屏蔽费别%Type := 0,
  住院可否分零_In   In 药品规格.住院可否分零%Type := 0,
  药库分批_In       In 药品规格.药库分批%Type := Null,
  药房分批_In       In 药品规格.药房分批%Type := Null,
  最大效期_In       In 药品规格.最大效期%Type := Null,
  差价让利比_In     In 药品规格.差价让利比%Type := 0,
  成本价_In         In 药品规格.成本价%Type := 0,
  当前售价_In       In 收费价目.现价%Type := 0,
  收入id_In         In 收费价目.收入项目id%Type := Null,
  合同单位id_In     In 药品规格.合同单位id%Type := Null,
  说明_In           In 收费项目目录.说明%Type := Null,
  动态分零_In       In 药品规格.动态分零%Type := 0,
  发药类型_In       In 药品规格.发药类型%Type := Null,
  备选码_In         In 收费项目目录.备选码%Type := Null,
  增值税率_In       In 药品规格.增值税率%Type := Null,
  基本药物_In       In 药品规格.基本药物%Type := Null,
  站点_In           In 收费项目目录.站点%Type := Null,
  是否常备_In       In 药品规格.是否常备%Type := Null,
  存储温度_In       In 输液药品属性.存储温度%Type := Null,
  存储条件_In       In 输液药品属性.存储条件%Type := Null,
  配药类型_In       In 输液药品属性.配药类型%Type := Null,
  是否不予配置_In   In 输液药品属性.是否不予配置%Type := Null,
  容量_In           In 药品规格.容量%Type := Null,
  病案费目_In       In 收费项目目录.病案费目%Type := Null,
  门诊可否分零_In   In 药品规格.门诊可否分零%Type := 0,
  Ddd值_In          In 药品规格.Ddd值%Type := 0,
  高危药品_In       药品规格.高危药品%Type := Null,
  送货单位_In       In 药品规格.送货单位%Type := Null,
  送货包装_In       In 药品规格.送货包装%Type := Null,
  输液注意事项_In   In 输液药品属性.输液注意事项%Type := Null,
  是否摆药_In       In 药品规格.是否摆药%Type := Null,
  是否零差价管理_In In 药品规格.是否零差价管理%Type := Null,
  本位码_In         In 药品规格.本位码%Type := Null
) Is
  v_药名id   诊疗项目目录.Id%Type;
  v_名称     诊疗项目目录.名称%Type;
  v_是否变价 收费项目目录.是否变价%Type; --允许定价药品随时改为时价，时价药品只能在未发生的情况下修改为定价，其它情况不允许修改定价属性 
  v_发生     Number(2);
  Err_Notfind Exception;
  v_No           收费价目.No%Type;
  v_Temp         收费项目目录.病案费目%Type;
  v_病案费目     收费项目目录.病案费目%Type;
  n_指导差价率   药品规格.指导差价率%Type;
  n_药品上次售价 药品规格.上次售价%Type;
  n_零售金额     药品库存.实际金额%Type;
  n_收发id       药品收发记录.Id%Type;
  n_流通金额小数 Number;
  n_序号         Number(8);
  Classid        Number(18); --入出类别
  v_Billno       药品收发记录.No%Type; --调价单号
  n_价格id       收费价目.Id%Type;
  n_收费价目现价 收费价目.现价%Type;
  n_原价         药品价格记录.原价%Type;
  n_药品价格记录 Number(1);
  v_类别         收费项目目录.类别%Type;
  --定价->时价后更新药品价格记录的值

  Cursor c_Priceadjust Is
    Select s.药品id, s.库房id, Nvl(s.批次, 0) As 批次, s.上次供应商id As 供应商id, s.上次批号 As 批号, s.效期, s.上次产地 As 产地,
           Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额, Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价,
           s.平均成本价, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期
    From 药品库存 S
    Where s.药品id = 药品id_In And s.性质 = 1 
    Order By s.药品id, s.批次, s.库房id;

  r_Priceadjust c_Priceadjust%RowType;

Begin
  v_病案费目 := 病案费目_In;
  --判断病案费目 
  If v_病案费目 Is Null Then
    If 收入id_In Is Not Null Then
      Begin
        Select 病案费目 Into v_Temp From 收入项目 Where ID = 收入id_In;
      Exception
        When Others Then
          v_Temp := Null;
      End;
      If v_Temp Is Not Null Then
        v_病案费目 := v_Temp;
      End If;
    End If;
  End If;
  --通用名称 
  Select ID, 名称
  Into v_药名id, v_名称
  From 诊疗项目目录
  Where ID = (Select 药名id From 药品规格 Where 药品id = 药品id_In);
  --取原始的定价属性 
  Select 是否变价 Into v_是否变价 From 收费项目目录 Where ID = 药品id_In;
  --规格信息 
  Update 收费项目目录
  Set 编码 = 编码_In, 名称 = v_名称, 规格 = 规格_In, 产地 = 产地_In, 计算单位 = 售价单位_In, 费用类型 = 费用类型_In, 服务对象 = 服务对象_In, 屏蔽费别 = 屏蔽费别_In,
      病案费目 = v_病案费目, 说明 = 说明_In, 备选码 = 备选码_In, 站点 = 站点_In
  Where ID = 药品id_In
  Returning 类别 Into v_类别;
  If Sql%RowCount = 0 Then
    Raise Err_Notfind;
  End If;
  n_指导差价率 := (1 - 1 / (1 + 加成率_In / 100)) * 100;
  Update 药品规格
  Set 标识码 = 标识码_In, 药品来源 = 药品来源_In, 批准文号 = 批准文号_In, 注册商标 = 注册商标_In, 剂量系数 = 剂量系数_In, 门诊单位 = 门诊单位_In, 门诊包装 = 门诊包装_In,
      住院单位 = 住院单位_In, 住院包装 = 住院包装_In, 药库单位 = 药库单位_In, 药库包装 = 药库包装_In, 申领单位 = 申领单位_In, 申领阀值 = 申领阀值_In, 指导批发价 = 指导批发价_In,
      扣率 = 扣率_In, 指导零售价 = 指导零售价_In, 指导差价率 = n_指导差价率, 管理费比例 = 管理费比例_In, 药价级别 = 药价级别_In, 住院可否分零 = 住院可否分零_In,
      药库分批 = 药库分批_In, 药房分批 = 药房分批_In, 最大效期 = 最大效期_In, 招标药品 = 招标药品_In, Gmp认证 = Gmp认证_In, 差价让利比 = 差价让利比_In,
      合同单位id = 合同单位id_In, 动态分零 = 动态分零_In, 发药类型 = 发药类型_In, 增值税率 = 增值税率_In, 基本药物 = 基本药物_In, 是否常备 = 是否常备_In, 容量 = 容量_In,
      门诊可否分零 = 门诊可否分零_In, Ddd值 = Ddd值_In, 高危药品 = 高危药品_In, 送货单位 = 送货单位_In, 送货包装 = 送货包装_In, 加成率 = 加成率_In, 是否摆药 = 是否摆药_In,
      是否零差价管理 = 是否零差价管理_In, 本位码 = 本位码_In
  Where 药品id = 药品id_In;

  --朱玉宝修改：建立药品（西成药、中成药）时，缺省服务对象为门诊和住院，因此修改规格药品时，不再根据规格药品的服务对象更新药品的服务对象 
  --诊疗项目服务对象的更改 
  --select nvl(sum(distinct I.服务对象),0) into v_对象 
  --from 收费项目目录 I,药品规格 S 
  --where I.ID=S.药品ID and S.药名ID=v_药名ID; 
  --update 诊疗项目目录 
  --set 服务对象=decode(v_对象,0,0,1,1,2,2,3) 
  --where ID=v_药名ID; 

  --别名的处理 
  If 数字码_In Is Null Then
    Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 1 And 码类 = 3;
  Else
    Update 收费项目别名 Set 名称 = v_名称, 简码 = 数字码_In Where 收费细目id = 药品id_In And 性质 = 1 And 码类 = 3;
    If Sql%RowCount = 0 Then
      Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, v_名称, 1, 数字码_In, 3);
    End If;
  End If;
  If 品名_In Is Null Then
    Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3;
  Else
    If 拼音_In Is Null Then
      Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 1;
    Else
      Update 收费项目别名 Set 名称 = 品名_In, 简码 = 拼音_In Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 拼音_In, 1);
      End If;
    End If;
    If 五笔_In Is Null Then
      Delete 收费项目别名 Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 2;
    Else
      Update 收费项目别名 Set 名称 = 品名_In, 简码 = 五笔_In Where 收费细目id = 药品id_In And 性质 = 3 And 码类 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (药品id_In, 品名_In, 3, 五笔_In, 2);
      End If;
    End If;
  End If;

  --定价信息：如果已经有发生，则不允许直接更改这些信息 
  Select Nvl(Count(*), 0) Into v_发生 From 药品收发记录 Where 药品id = 药品id_In And Rownum < 2;
  If v_发生 = 0 Then
    Update 药品规格 Set 成本价 = 成本价_In Where 药品id = 药品id_In;
    If 收入id_In Is Not Null Then
      Update 收费价目
      Set 现价 = 当前售价_In, 收入项目id = 收入id_In, 变动原因 = 1, 调价说明 = '修改定价', 调价人 = User
      Where 收费细目id = 药品id_In And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因 = 1;
      If Sql%RowCount = 0 Then
        v_No := Nextno(9);
        Insert Into 收费价目
          (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
        Values
          (收费价目_Id.Nextval, Null, 药品id_In, 0, 当前售价_In, 收入id_In, 1, '新增定价', User, Sysdate,
           To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
      End If;
    End If;
  Else
    --发生业务数据时，不能修改价格但是可以修改收入项目 
    Update 收费价目
    Set 收入项目id = 收入id_In
    Where 收费细目id = 药品id_In And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因 = 1;
  End If;

  --时价->定价
  If v_是否变价 = 1 And 是否变价_In = 0 Then
    Update 收费项目目录 Set 是否变价 = 是否变价_In Where ID = 药品id_In;
  
    Begin
      Select 现价, ID As 价格id
      Into n_收费价目现价, n_价格id
      From 收费价目
      Where 收费细目id = 药品id_In And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 变动原因 = 1;
    Exception
      When Others Then
        n_收费价目现价 := Null;
        n_价格id       := Null;
    End;
  
    Begin
      Select 上次售价 Into n_药品上次售价 From 药品规格 Where 药品id = 药品id_In;
    Exception
      When Others Then
        n_药品上次售价 := Null;
    End;
  
    If n_药品上次售价 Is Null Then
      n_药品上次售价 := n_收费价目现价;
    End If;
  
    Zl_收费价目_Stop(药品id_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, n_价格id, 药品id_In, n_收费价目现价, n_药品上次售价, 收入id_In, 1, '时价转定价', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;
  
    --取入出类别ID
    Select 类别id Into Classid From 药品单据性质 Where 单据 = 13;
  
    n_序号   := 0;
    v_Billno := Null;
  
    For r_Priceadjust In c_Priceadjust Loop
      If n_药品上次售价 <> r_Priceadjust.零售价 Then
        If v_Billno Is Null Then
          Select Nextno(147) Into v_Billno From Dual;
        End If;
        n_序号 := n_序号 + 1;
        Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
        n_零售金额 := Round(n_药品上次售价 * r_Priceadjust.实际数量, n_流通金额小数) -
                  Round(r_Priceadjust.零售价 * r_Priceadjust.实际数量, n_流通金额小数);
        --产生调价影响记录
        Insert Into 药品收发记录
          (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
           填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
        Values
          (n_收发id, 1, 13, v_Billno, n_序号, Classid, r_Priceadjust.药品id, r_Priceadjust.批次, r_Priceadjust.批号,
           r_Priceadjust.效期, r_Priceadjust.产地, 1, r_Priceadjust.实际数量, 0, r_Priceadjust.零售价, 0, n_药品上次售价,
           r_Priceadjust.扣率, n_零售金额, n_零售金额, '时价转定价', Zl_Username, Sysdate, r_Priceadjust.库房id, 1, n_价格id, Zl_Username,
           Sysdate, r_Priceadjust.实际金额, r_Priceadjust.实际差价, r_Priceadjust.供应商id);
      
        Zl_药品库存_Update(n_收发id, 2, 0);
      End If;
    End Loop;
  
    --定价->时价
  Elsif v_是否变价 = 0 And 是否变价_In = 1 Then
    Update 收费项目目录 Set 是否变价 = 是否变价_In Where ID = 药品id_In;
    Begin
      Select 现价, ID As 价格id
      Into n_收费价目现价, n_价格id
      From 收费价目
      Where 收费细目id = 药品id_In And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 变动原因 = 1;
    Exception
      When Others Then
        n_收费价目现价 := Null;
        n_价格id       := Null;
    End;
  
    Zl_收费价目_Stop(药品id_In, Sysdate - 1 / 24 / 60 / 60);
    v_No := Nextno(9);
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, n_价格id, 药品id_In, n_收费价目现价, n_收费价目现价, 收入id_In, 1, '定价转时价', Zl_Username, Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  
    For r_Priceadjust In c_Priceadjust Loop
      n_药品价格记录 := 0;
      Begin
        Select 1, 现价
        Into n_药品价格记录, n_原价
        From 药品价格记录
        Where 药品id = r_Priceadjust.药品id And 库房id = r_Priceadjust.库房id And Nvl(批次, 0) = r_Priceadjust.批次 And 记录状态 = 1 And
              价格类型 = 1;
      Exception
        When Others Then
          n_药品价格记录 := 0;
          n_原价         := n_收费价目现价;
      End;
    
      If n_药品价格记录 = 1 Then
        Zl_药品价格记录_Stop(1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, Sysdate - 1 / 24 / 60 / 60, 2);
      End If;
      Zl_药品价格记录_Insert(0, 1, r_Priceadjust.库房id, r_Priceadjust.药品id, r_Priceadjust.批次, n_原价, n_收费价目现价, Sysdate, '定价转时价',
                       Zl_Username, Null, r_Priceadjust.供应商id, r_Priceadjust.批号, r_Priceadjust.效期, r_Priceadjust.产地,
                       r_Priceadjust.灭菌效期, Null, Null, Null, Null, 1);
    
      Update 药品库存
      Set 零售价 = n_收费价目现价
      Where 性质 = 1 And 库房id = r_Priceadjust.库房id And 药品id = r_Priceadjust.药品id And Nvl(批次, 0) = r_Priceadjust.批次;
    
    End Loop;
  End If;

  --药品生产商比较增加 
  If 产地_In Is Not Null Then
    Update 药品生产商 Set 名称 = 产地_In Where 名称 = 产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 药品生产商
        (编码, 名称, 简码)
        Select Nvl(Max(To_Number(编码)), 0) + 1, 产地_In, zlSpellCode(产地_In, 10) From 药品生产商;
    End If;
  End If;

  --修改输液药品属性 
  Update 输液药品属性
  Set 存储温度 = 存储温度_In, 存储条件 = 存储条件_In, 配药类型 = 配药类型_In, 是否不予配置 = 是否不予配置_In, 输液注意事项 = 输液注意事项_In
  Where 药品id = 药品id_In;

  If Sql%NotFound Then
    Insert Into 输液药品属性
      (药品id, 存储温度, 存储条件, 配药类型, 是否不予配置, 输液注意事项)
    Values
      (药品id_In, 存储温度_In, 存储条件_In, 配药类型_In, 是否不予配置_In, 输液注意事项_In);
  End If;

  --药品精度调整(零差价模式时)
  Zl_药品卫材精度_零差价调整;

  b_Message.Zlhis_Dict_036(v_类别, 药品id_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该规格不存在，可能已被其他用户删除！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_成药规格_Update;
/

--125588:秦龙,2018-05-14,处理库存实际数量为0的数据
Create Or Replace Procedure Zl_药品收发记录_Adjust
(
  药品id_In   In Number, --药品ID,为0时检查所有预调价
  调价方式_In In Number := 0 --0-检查售价和成本价预调价,1-只检查售价预调价,2-只检查成本价预调价
) As
  Classid          Number(18); --入出类别
  v_Billno         药品收发记录.No%Type; --调价单号
  Adjustdate       Date; --调价时间
  n_批次           Number(18);
  n_现价           收费价目.现价%Type;
  n_原价           收费价目.原价%Type;
  n_序号           Number(8);
  n_原价id         收费价目.原价id%Type;
  n_零售金额       药品库存.实际金额%Type;
  n_收发id         药品收发记录.Id%Type;
  n_流通金额小数   Number;
  n_Stockid        药品收发记录.库房id%Type;
  n_入出类别id     药品收发记录.入出类别id%Type;
  n_入出系数       药品收发记录.入出系数%Type;
  n_价格id         收费价目.Id%Type;
  n_无库存调价模式 Number(1) := 0;
  n_分批属性       Number(1) := 0;
  n_消息调用       Number(1) := 0;
  --定价售价，时价售价预调价记录
  --价格类型：0-定价售价,1-时价售价
  Cursor c_Priceadjust Is
    Select 0 As 价格类型, p.Id As 价格id, p.原价id, p.执行日期, p.原价, p.现价, i.药品id, s.库房id As 库房id, Nvl(s.批次, 0) As 批次, s.上次批号 As 批号,
           s.效期, s.上次产地 As 产地, s.上次供应商id As 供应商id, Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额,
           Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价, s.平均成本价, s.Rowid As 库存记录, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期,
           p.调价汇总号
    From 收费价目 P, 药品规格 I, 收费项目目录 A, 药品库存 S
    Where i.药品id = a.Id And p.收费细目id = i.药品id And i.药品id = s.药品id(+) And s.性质(+) = 1 And Nvl(a.是否变价, 0) = 0 And
          Sysdate Between p.执行日期 And Nvl(p.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.变动原因, 0) = 0 And
          p.收费细目id = Decode(药品id_In, 0, p.收费细目id, 药品id_In) And 调价方式_In In (0, 1)
    Union All
    Select 价格类型, p.Id As 价格id, p.原价id, p.执行日期, p.原价, p.现价, i.药品id, p.库房id As 库房id, Nvl(p.批次, 0) As 批次, p.批号 As 批号, p.效期,
           p.产地 As 产地, s.上次供应商id As 供应商id, Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额,
           Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价, s.平均成本价, s.Rowid As 库存记录, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期,
           p.调价汇总号
    From 药品价格记录 P, 药品规格 I, 收费项目目录 A, 药品库存 S
    Where i.药品id = a.Id And p.药品id = i.药品id And p.库房id = s.库房id(+) And p.药品id = s.药品id(+) And
          Nvl(p.批次, 0) = Nvl(s.批次(+), 0) And s.性质(+) = 1 And Sysdate Between p.执行日期 And
          Nvl(p.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.记录状态, 0) = 0 And
          p.药品id = Decode(药品id_In, 0, p.药品id, 药品id_In) And 价格类型 = 1 And 调价方式_In In (0, 1)
    Order By 价格类型, 药品id, 批次, 库房id;

  r_Priceadjust c_Priceadjust%RowType;

  --成本价预调价记录
  Cursor c_Costadjust Is
    Select 价格类型, p.Id As 价格id, p.原价id, p.执行日期, p.原价, p.现价, i.药品id, p.库房id As 库房id, Nvl(p.批次, 0) As 批次, p.批号 As 批号, p.效期,
           p.产地 As 产地, s.上次供应商id As 供应商id, Nvl(s.实际数量, 0) As 实际数量, s.上次扣率 As 扣率, Nvl(s.实际金额, 0) As 实际金额,
           Nvl(s.实际差价, 0) As 实际差价, Nvl(s.零售价, 0) As 零售价, s.平均成本价, s.Rowid As 库存记录, s.灭菌效期, s.批准文号, s.上次生产日期 As 生产日期,
           p.调价汇总号
    From 药品价格记录 P, 药品规格 I, 收费项目目录 A, 药品库存 S
    Where i.药品id = a.Id And p.药品id = i.药品id And p.库房id = s.库房id(+) And p.药品id = s.药品id(+) And
          Nvl(p.批次, 0) = Nvl(s.批次(+), 0) And s.性质(+) = 1 And Sysdate Between p.执行日期 And
          Nvl(p.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.记录状态, 0) = 0 And
          p.药品id = Decode(药品id_In, 0, p.药品id, 药品id_In) And 价格类型 = 2 And 调价方式_In In (0, 2)
    Order By 药品id, 批次, 库房id;

  r_Costadjust c_Costadjust%RowType;

  --当前生效的价格，用于无库存调价
  Cursor c_Nostockadjust
  (
    Drugid_In 药品价格记录.药品id%Type,
    Type_In   药品价格记录.价格类型%Type
  ) Is
    Select a.价格类型, a.Id As 价格id, a.原价, a.现价, a.药品id, a.库房id, a.批次, a.供药单位id, a.批号, a.效期, a.产地
    From 药品价格记录 A
    Where Sysdate Between a.执行日期 And Nvl(a.终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And a.记录状态 = 1 And
          a.药品id = Drugid_In And a.价格类型 = Type_In And a.库房id Is Not Null
    Order By a.库房id, a.药品id, a.批次;

  r_Nostockadjust c_Nostockadjust%RowType;
Begin
  --取流通业务精度位数
  --类别:1-药品 2-卫材
  --内容：2-零售价 4-金额
  --单位：药品:1-售价 5-金额单位
  Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;

  --取入出类别ID
  Select 类别id Into Classid From 药品单据性质 Where 单据 = 13;

  Adjustdate := Sysdate;

  --售价调价处理
  If 调价方式_In = 0 Or 调价方式_In = 1 Then
  
    n_序号 := 0;
  
    --取调价NO取
    Select Nextno(147) Into v_Billno From Dual;
  
    For r_Priceadjust In c_Priceadjust Loop
      If r_Priceadjust.库房id Is Not Null Then
        --有库房id正常调价
      
        --取分批属性
        n_分批属性 := Zl_Fun_Getbatchpro(r_Priceadjust.库房id, r_Priceadjust.药品id);
      
        --产生调价盈亏记录的条件：1.要有库存记录，2.分批属性和库存批次一致
        If r_Priceadjust.库存记录 Is Not Null And ((n_分批属性 = 1 And r_Priceadjust.批次 > 0) Or
           (n_分批属性 = 0 And r_Priceadjust.批次 = 0)) Then
          n_序号 := n_序号 + 1;
        
          Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
        
          n_原价 := r_Priceadjust.原价;
        
          --时价调价，如果原价和当前库存不一致，则以当前库存为准
          If r_Priceadjust.价格类型 = 1 And r_Priceadjust.原价 <> r_Priceadjust.零售价 And r_Priceadjust.库存记录 Is Not Null Then
            n_原价 := r_Priceadjust.零售价;
          End If;
        
          n_零售金额 := Round(r_Priceadjust.现价 * r_Priceadjust.实际数量, n_流通金额小数) - Round(n_原价 * r_Priceadjust.实际数量, n_流通金额小数);
        
          n_价格id := r_Priceadjust.价格id;
          If r_Priceadjust.价格类型 = 1 Then
            Select ID
            Into n_价格id
            From 收费价目
            Where Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And 收费细目id = r_Priceadjust.药品id;
          End If;
        
          --产生调价影响记录
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 入出类别id, 药品id, 批次, 批号, 效期, 产地, 付数, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 扣率, 零售金额, 差价, 摘要, 填制人,
             填制日期, 库房id, 入出系数, 价格id, 审核人, 审核日期, 单量, 频次, 供药单位id)
          Values
            (n_收发id, 1, 13, v_Billno, n_序号, Classid, r_Priceadjust.药品id, r_Priceadjust.批次, r_Priceadjust.批号,
             r_Priceadjust.效期, r_Priceadjust.产地, 1, r_Priceadjust.实际数量, 0, n_原价, 0, r_Priceadjust.现价, r_Priceadjust.扣率,
             n_零售金额, n_零售金额, '药品调价', Zl_Username, Adjustdate, r_Priceadjust.库房id, 1, n_价格id, Zl_Username, Adjustdate,
             r_Priceadjust.实际金额, r_Priceadjust.实际差价, r_Priceadjust.供应商id);
        
          Zl_未审药品记录_Insert(n_收发id);
        
          --更新药品库存，无库存不执行
          If r_Priceadjust.库存记录 Is Not Null Then
            Zl_药品库存_Update(n_收发id, 2, 0);
          End If;
        End If;
      
        --更新原价格信息
        If r_Priceadjust.价格类型 = 1 Then
          Update 药品价格记录 Set 记录状态 = 2 Where ID = r_Priceadjust.原价id;
        End If;
      
        --时价调价更新价格表中的信息
        If r_Priceadjust.价格类型 = 1 Then
          --更新当前价格信息
          If r_Priceadjust.库存记录 Is Not Null Then
            Update 药品价格记录
            Set 批号 = r_Priceadjust.批号, 效期 = r_Priceadjust.效期, 产地 = r_Priceadjust.产地, 灭菌效期 = r_Priceadjust.灭菌效期,
                供药单位id = r_Priceadjust.供应商id, 原价 = n_原价, 收发id = n_收发id, 记录状态 = 1
            Where ID = r_Priceadjust.价格id;
          Else
            --无库存时只更新记录状态，收发id
            Update 药品价格记录 Set 收发id = n_收发id, 记录状态 = 1 Where ID = r_Priceadjust.价格id;
          End If;
        End If;
      
        --更新批号对照表售价
        If r_Priceadjust.价格类型 = 1 Then
          --如果是时价，则更新该药品批次对应的价格
          Update 药品批号对照
          Set 售价 = r_Priceadjust.现价
          Where 药品id = r_Priceadjust.药品id And Nvl(批次, 0) = r_Priceadjust.批次 And 售价 <> r_Priceadjust.现价;
        End If;
      
        --消息处理
        --定价只调用一次消息，时价可多次调用
        If (r_Priceadjust.价格类型 = 0 And n_消息调用 = 0) Or r_Priceadjust.价格类型 = 1 Then
          n_消息调用 := 1;
          b_Message.Zlhis_Drug_009(r_Priceadjust.价格id, r_Priceadjust.价格类型);
        End If;
      Else
        --无库存调价模式，价格表中该药品所有生效的价格都要按无库存调价时的价格调整
      
        If r_Priceadjust.价格类型 = 1 Then
          --更新原价格信息
          Update 药品价格记录 Set 记录状态 = 2 Where ID = r_Priceadjust.原价id;
        
          --更新现价格状态
          Update 药品价格记录 Set 记录状态 = 1 Where ID = r_Priceadjust.价格id;
        End If;
      
        n_无库存调价模式 := 0;
        For r_Nostockadjust In c_Nostockadjust(r_Priceadjust.药品id, 1) Loop
          If r_Priceadjust.现价 <> r_Nostockadjust.现价 Then
            Zl_药品价格记录_Stop(1, r_Nostockadjust.库房id, r_Nostockadjust.药品id, r_Nostockadjust.批次,
                           Adjustdate - 2 / 24 / 60 / 60);
            Zl_药品价格记录_Insert(1, 1, r_Nostockadjust.库房id, r_Nostockadjust.药品id, r_Nostockadjust.批次, Null,
                             r_Priceadjust.现价, Adjustdate - 1 / 24 / 60 / 60, '药品调价', Zl_Username, r_Priceadjust.调价汇总号,
                             r_Nostockadjust.供药单位id, r_Nostockadjust.批号, r_Nostockadjust.效期, r_Nostockadjust.产地);
            n_无库存调价模式 := 1;
          End If;
        End Loop;
        If n_无库存调价模式 = 1 Then
          Zl_药品收发记录_Adjust(r_Priceadjust.药品id, 1);
        End If;
      End If;
    
      --更新规格价格
      If r_Priceadjust.现价 <> r_Priceadjust.原价 Then
        Update 药品规格
        Set 上次售价 = r_Priceadjust.现价
        Where 药品id = r_Priceadjust.药品id And 上次售价 <> r_Priceadjust.现价;
      End If;
    
      If r_Priceadjust.价格类型 = 0 Then
        n_价格id := r_Priceadjust.价格id;
      Else
        Begin
          Select ID
          Into n_价格id
          From 收费价目
          Where Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(变动原因, 0) = 0 And
                收费细目id = r_Priceadjust.药品id;
        Exception
          When Others Then
            n_价格id := 0;
        End;
      End If;
    
      If n_价格id > 0 Then
        Update 收费价目 Set 变动原因 = 1 Where Nvl(变动原因, 0) = 0 And ID = n_价格id;
      End If;
    
      --更新批号对照表售价
      If r_Priceadjust.价格类型 = 0 Then
        --如果是定价，则更新该药品对应的所有批次的售价
        Update 药品批号对照
        Set 售价 = r_Priceadjust.现价
        Where 药品id = r_Priceadjust.药品id And 售价 <> r_Priceadjust.现价;
      End If;
    End Loop;
  End If;

  --成本价调价处理
  If 调价方式_In = 0 Or 调价方式_In = 2 Then
  
    n_序号    := 0;
    n_Stockid := 0;
  
    Select b.Id, b.系数
    Into n_入出类别id, n_入出系数
    From 药品单据性质 A, 药品入出类别 B
    Where a.类别id = b.Id And a.单据 = 5 And Rownum < 2;
  
    v_Billno := Nextno(25, n_Stockid);
  
    For r_Costadjust In c_Costadjust Loop
      If r_Costadjust.库房id Is Not Null Then
        --有库房id正常调价
      
        --取分批属性
        n_分批属性 := Zl_Fun_Getbatchpro(r_Costadjust.库房id, r_Costadjust.药品id);
      
        --产生调价盈亏记录的条件：1.要有库存记录，2.分批属性和库存批次一致
        If r_Costadjust.库存记录 Is Not Null And ((n_分批属性 = 1 And r_Costadjust.批次 > 0) Or
           (n_分批属性 = 0 And r_Costadjust.批次 = 0)) Then
          n_序号 := n_序号 + 1;
        
          --产生库存差价调整单
          Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
        
          --如果原价和当前库存不一致，则以当前库存为准
          n_原价 := r_Costadjust.原价;
          If r_Costadjust.原价 <> r_Costadjust.平均成本价 And r_Costadjust.库存记录 Is Not Null Then
            n_原价 := r_Costadjust.平均成本价;
          End If;
        
          n_零售金额 := Round(n_原价 * r_Costadjust.实际数量, n_流通金额小数) - Round(r_Costadjust.现价 * r_Costadjust.实际数量, n_流通金额小数);
        
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 库房id, 入出类别id, 供药单位id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 填写数量, 实际数量, 零售价, 零售金额, 成本价, 成本金额,
             差价, 摘要, 填制人, 填制日期, 审核人, 审核日期, 生产日期, 批准文号, 单量, 发药方式, 扣率, 灭菌效期)
          Values
            (n_收发id, 1, 5, v_Billno, n_序号, r_Costadjust.库房id, n_入出类别id, r_Costadjust.供应商id, n_入出系数, r_Costadjust.药品id,
             r_Costadjust.批次, r_Costadjust.产地, r_Costadjust.批号, r_Costadjust.效期, r_Costadjust.实际数量, 0, r_Costadjust.实际金额,
             0, r_Costadjust.实际差价, 0, n_零售金额, '成本价调价', Zl_Username, Adjustdate, Zl_Username, Adjustdate,
             r_Costadjust.生产日期, r_Costadjust.批准文号, r_Costadjust.现价, 1, n_原价, r_Costadjust.灭菌效期);
        
          Zl_未审药品记录_Insert(n_收发id);
        
          Zl_药品库存_Update(n_收发id, 2, 0);
        
          --更新当前价格信息
          Update 药品价格记录
          Set 批号 = r_Costadjust.批号, 效期 = r_Costadjust.效期, 产地 = r_Costadjust.产地, 灭菌效期 = r_Costadjust.灭菌效期,
              供药单位id = r_Costadjust.供应商id, 原价 = n_原价, 收发id = n_收发id, 记录状态 = 1
          Where ID = r_Costadjust.价格id;
        Else
          --无库存时只更新记录状态
          Update 药品价格记录 Set 记录状态 = 1 Where ID = r_Costadjust.价格id;
        End If;
      
        --更新原价格信息
        Update 药品价格记录 Set 记录状态 = 2 Where ID = r_Costadjust.原价id;
      
        --更新批号对照表成本价
        Update 药品批号对照
        Set 成本价 = r_Costadjust.现价
        Where 药品id = r_Costadjust.药品id And Nvl(批次, 0) = r_Costadjust.批次 And 成本价 <> r_Costadjust.现价;
      Else
        --无库存调价模式，价格表中该药品所有生效的价格都要按无库存调价时的价格调整
      
        --更新原价格信息
        Update 药品价格记录 Set 记录状态 = 2 Where ID = r_Costadjust.原价id;
      
        --更新现价格状态
        Update 药品价格记录 Set 记录状态 = 1 Where ID = r_Costadjust.价格id;
      
        n_无库存调价模式 := 0;
        For r_Nostockadjust In c_Nostockadjust(r_Costadjust.药品id, 2) Loop
          If r_Costadjust.现价 <> r_Nostockadjust.现价 Then
            Zl_药品价格记录_Stop(2, r_Nostockadjust.库房id, r_Nostockadjust.药品id, r_Nostockadjust.批次,
                           Adjustdate - 2 / 24 / 60 / 60);
            Zl_药品价格记录_Insert(1, 2, r_Nostockadjust.库房id, r_Nostockadjust.药品id, r_Nostockadjust.批次, Null,
                             r_Costadjust.现价, Adjustdate - 1 / 24 / 60 / 60, '成本价调价', Zl_Username, r_Costadjust.调价汇总号,
                             r_Nostockadjust.供药单位id, r_Nostockadjust.批号, r_Nostockadjust.效期, r_Nostockadjust.产地);
            n_无库存调价模式 := 1;
          End If;
        End Loop;
        If n_无库存调价模式 = 1 Then
          Zl_药品收发记录_Adjust(r_Costadjust.药品id, 2);
        End If;
      End If;
    
      --更新规格价格
      If r_Costadjust.原价 <> r_Costadjust.现价 Then
        Update 药品规格
        Set 成本价 = r_Costadjust.现价
        Where 药品id = r_Costadjust.药品id And 成本价 <> r_Costadjust.现价;
      End If;
    
      --消息处理
      b_Message.Zlhis_Drug_007(r_Costadjust.价格id);
    End Loop;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_Adjust;
/


--125656:李业庆,2018-05-14,销售退费未删除未发药品记录
Create Or Replace Procedure Zl_药品收发记录_销售退费
(
  费用id_In   In 门诊费用记录.Id%Type,
  销帐数量_In In 药品收发记录.实际数量%Type := 0,
  配药id_In   Varchar2 := Null,
  消息_In     Number := 0 --是否发送销帐消息
) Is
  ----------------------------------
  --功能：删除门诊收费单、门诊划价单、门诊收费销帐时用来处理药品库存、药品收发记录、未发药记录的过程
  --参数：
  --      费用id_In：门诊费用记录或者住院费用记录做删除或者销帐时被删除单据的id
  --      销帐数量_In：销帐审核时需要销帐的数量
  --      配药id_In：销帐审核时输液配置中心需要传递的记录id，以字符串传递，用逗号分割，如：1001,1002,1003
  --      为空表示冲销所有可冲销行
  -----------------------------------
  --该游标用于处理药品库存可用数量
  l_药品收发   t_Numlist := t_Numlist();
  n_单据       药品收发记录.单据%Type;
  v_No         药品收发记录.No%Type;
  n_原始数量   Number;
  n_销帐数量   Number;
  v_收费类别   收费项目目录.类别%Type;
  n_虚拟库房id 药品收发记录.库房id%Type;
  n_其他出库id 药品收发记录.Id%Type;
  n_库房id     药品收发记录.库房id%Type;
  n_备货卫材   Number;
  v_收发ids    Varchar2(4000); --用户消息锚点发送，格式:收发id,数量|收发id,数量...

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  --找出需要处理的药品单据
  Cursor c_药品收发记录 Is
    Select ID, 费用id, NO, 单据, 药品id, 库房id, Nvl(批次, 0) 批次, 批号, 产地,
           Decode(发药方式, Null, 1, -1, 0, 1) * Nvl(付数, 1) * Nvl(实际数量, 0) As 数量, 实际数量, 付数, 发药方式, 灭菌效期, 效期, 商品条码, 内部条码
    From 药品收发记录
    Where 单据 In (8, 9, 10, 21, 24, 25, 26) And Mod(记录状态, 3) = 1 And 审核人 Is Null And 费用id = 费用id_In;

  Cursor c_销帐审核 Is
    Select /*+ rule*/
     a.Id, a.费用id, a.No, a.单据, a.药品id, a.库房id, Nvl(a.批次, 0) 批次, a.批号, a.产地,
     Decode(a.发药方式, Null, 1, -1, 0, 1) * Nvl(a.付数, 1) * Nvl(a.实际数量, 0) As 数量, a.实际数量, a.付数, a.发药方式, a.灭菌效期, a.效期, a.商品条码,
     a.内部条码
    From 药品收发记录 A, Table(f_Str2list(配药id_In)) B, 输液配药内容 C
    Where a.单据 In (9, 10, 25, 26) And Mod(a.记录状态, 3) = 1 And a.审核人 Is Null And a.费用id = 费用id_In And a.Id = c.收发id And
          c.记录id = b.Column_Value
    Order By 填制日期;

  r_Row c_药品收发记录%RowType;
Begin
  n_单据 := 0;
  v_No   := '';

  If 销帐数量_In = 0 Then
  
    --数量为空表示是全部删除
    --打开游标
    Open c_药品收发记录;
  
    --遍历游标
    Loop
      Fetch c_药品收发记录
        Into r_Row;
      Exit When c_药品收发记录%NotFound;
    
      Select 类别 Into v_收费类别 From 收费项目目录 Where ID = r_Row.药品id;
    
      --处理药品库存
      If r_Row.库房id Is Not Null Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + r_Row.数量
        Where 库房id = r_Row.库房id And 药品id = r_Row.药品id And Nvl(批次, 0) = Nvl(r_Row.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期, 商品条码, 内部条码)
          Values
            (r_Row.库房id, r_Row.药品id, 1, Nvl(r_Row.批次, 0), r_Row.效期, r_Row.数量, r_Row.批号, r_Row.产地, r_Row.灭菌效期,
             r_Row.商品条码, r_Row.内部条码);
        End If;
      
        --删除多余的库存数据
        Delete From 药品库存
        Where 库房id = r_Row.库房id And 药品id = r_Row.药品id And Nvl(批次, 0) = Nvl(r_Row.批次, 0) And 性质 = 1 And Nvl(可用数量, 0) = 0 And
              Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
      
        Zl_药品库存_可用数量异常处理(r_Row.库房id, r_Row.药品id, r_Row.批次);
      End If;
    
      n_单据 := r_Row.单据;
      v_No   := r_Row.No;
      l_药品收发.Extend;
      l_药品收发(l_药品收发.Count) := r_Row.Id;
    
      v_收发ids := v_收发ids || '|' || r_Row.Id || ',' || 0;
    End Loop;
  
    --关闭游标
    Close c_药品收发记录;
  Else
    --数量不为空表示是销帐审核操作
    If 配药id_In Is Not Null Then
      Open c_销帐审核;
    Else
      Open c_药品收发记录;
    End If;
    n_销帐数量 := 销帐数量_In;
  
    --只有住院记账处理才会走这一步
    Loop
      If 配药id_In Is Not Null Then
        Fetch c_销帐审核
          Into r_Row;
        Exit When c_销帐审核%NotFound;
      Else
        Fetch c_药品收发记录
          Into r_Row;
        Exit When c_药品收发记录%NotFound;
      End If;
    
      n_虚拟库房id := Null;
      n_其他出库id := Null;
      Select 类别 Into v_收费类别 From 收费项目目录 Where ID = r_Row.药品id;
      If v_收费类别 = '4' Then
        Begin
          Select 1, 库房id, ID
          Into n_备货卫材, n_虚拟库房id, n_其他出库id
          From 药品收发记录
          Where 费用id = 费用id_In And 审核日期 Is Null And 单据 = 21 And Rownum = 1;
        Exception
          When Others Then
            n_备货卫材 := 0;
        End;
      Else
        n_备货卫材 := 0;
      End If;
    
      n_单据     := r_Row.单据;
      v_No       := r_Row.No;
      n_原始数量 := r_Row.数量;
    
      If n_销帐数量 >= n_原始数量 Then
        l_药品收发.Extend;
        l_药品收发(l_药品收发.Count) := r_Row.Id;
        v_收发ids := v_收发ids || '|' || r_Row.Id || ',' || 0;
        If Nvl(n_其他出库id, 0) > 0 Then
          l_药品收发.Extend;
          l_药品收发(l_药品收发.Count) := n_其他出库id;
          v_收发ids := v_收发ids || '|' || n_其他出库id || ',' || 0;
        End If;
        n_销帐数量 := n_销帐数量 - n_原始数量;
      Else
        If v_收费类别 = '7' Then
          --当前行的数量要大
          Update 药品收发记录
          Set 付数 = 1, 实际数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量,
              填写数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(填写数量, 0) - n_销帐数量,
              成本金额 =
               (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量) * 成本价,
              零售金额 =
               (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量) * 零售价,
              差价 = Round((Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量) * 零售价 -
                          (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量) * 成本价, 5)
          Where ID = r_Row.Id;
        Else
          Update 药品收发记录
          Set 实际数量 = Nvl(实际数量, 0) - n_销帐数量, 填写数量 = Nvl(填写数量, 0) - n_销帐数量,
              成本金额 =
               (Nvl(实际数量, 0) - n_销帐数量) * 成本价,
              零售金额 =
               (Nvl(实际数量, 0) - n_销帐数量) * 零售价,
              差价 = Round((Nvl(实际数量, 0) - n_销帐数量) * 零售价 - (Nvl(实际数量, 0) - n_销帐数量) * 成本价, 5)
          Where ID = r_Row.Id;
        End If;
      
        v_收发ids := v_收发ids || '|' || r_Row.Id || ',' || n_销帐数量;
      
        --更新其他出库单
        If Nvl(n_其他出库id, 0) <> 0 Then
          If v_收费类别 = '7' Then
            Update 药品收发记录
            Set 付数 = 1, 实际数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量,
                填写数量 = Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量,
                成本金额 =
                 (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量) * 成本价,
                零售金额 =
                 (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量) * 零售价,
                差价 = Round((Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量) * 零售价 -
                            (Decode(付数, Null, 1, 0, 1, 付数) * Nvl(实际数量, 0) - n_销帐数量) * 成本价, 5)
            Where ID = Nvl(n_其他出库id, 0);
          Else
            Update 药品收发记录
            Set 实际数量 = Nvl(实际数量, 0) - n_销帐数量, 填写数量 = Nvl(实际数量, 0) - n_销帐数量,
                成本金额 =
                 (Nvl(实际数量, 0) - n_销帐数量) * 成本价,
                零售金额 =
                 (Nvl(实际数量, 0) - n_销帐数量) * 零售价,
                差价 = Round((Nvl(实际数量, 0) - n_销帐数量) * 零售价 - (Nvl(实际数量, 0) - n_销帐数量) * 成本价, 5)
            Where ID = Nvl(n_其他出库id, 0);
          End If;
        
          v_收发ids := v_收发ids || '|' || n_其他出库id || ',' || n_销帐数量;
        End If;
        n_原始数量 := n_销帐数量;
        n_销帐数量 := 0;
      End If;
      If Nvl(n_备货卫材, 0) = 1 Then
        n_库房id := n_虚拟库房id;
      Else
        n_库房id := r_Row.库房id;
      End If;
    
      If n_库房id Is Not Null Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + n_原始数量
        Where 库房id = n_库房id And 药品id = r_Row.药品id And Nvl(批次, 0) = Nvl(r_Row.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期)
          Values
            (n_库房id, r_Row.药品id, 1, Nvl(r_Row.批次, 0), r_Row.效期, n_原始数量, r_Row.批号, r_Row.产地, r_Row.灭菌效期);
        End If;
        Delete 药品库存
        Where 库房id = n_库房id And 药品id = r_Row.药品id And Nvl(批次, 0) = Nvl(r_Row.批次, 0) And 性质 = 1 And Nvl(可用数量, 0) = 0 And
              Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
      
        Zl_药品库存_可用数量异常处理(r_Row.库房id, r_Row.药品id, r_Row.批次);
      End If;
    
      If Nvl(n_备货卫材, 0) = 1 Then
        Update 药品库存
        Set 可用数量 = Nvl(可用数量, 0) + n_原始数量
        Where 库房id = r_Row.库房id And 药品id = r_Row.药品id And Nvl(批次, 0) = Nvl(r_Row.批次, 0) And 性质 = 1;
        If Sql%RowCount = 0 Then
          Insert Into 药品库存
            (库房id, 药品id, 性质, 批次, 效期, 可用数量, 上次批号, 上次产地, 灭菌效期)
          Values
            (r_Row.库房id, r_Row.药品id, 1, Nvl(r_Row.批次, 0), r_Row.效期, n_原始数量, r_Row.批号, r_Row.产地, r_Row.灭菌效期);
        End If;
      
        Delete 药品库存
        Where 库房id = r_Row.库房id And 药品id = r_Row.药品id And Nvl(批次, 0) = Nvl(r_Row.批次, 0) And 性质 = 1 And Nvl(可用数量, 0) = 0 And
              Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0 And Nvl(实际差价, 0) = 0;
      
        Zl_药品库存_可用数量异常处理(r_Row.库房id, r_Row.药品id, r_Row.批次);
      End If;
    
      If n_销帐数量 = 0 Then
        Exit;
      End If;
    End Loop;
  
    --不跟踪卫材的,不检查:因为不跟噻的话,不会在药品收发记录中存在
    If Nvl(n_销帐数量, 0) <> 0 And Not (v_收费类别 = '4' And n_原始数量 = 0) Then
      --未分配完成,表示此药品可能已经执行.
      v_Err_Msg := '要销帐的费用中存在已发的药品或卫材，或已被其他人销帐；这可能是并发操作引起的。';
      Raise Err_Item;
    End If;
  End If;

  --删除药品收发记录
  Forall I In 1 .. l_药品收发.Count
    Delete From 药品收发记录 Where ID = l_药品收发(I) And 审核人 Is Null;

  --删除未发药品记录
  Delete From 未发药品记录 A
  Where NO = v_No And 单据 = n_单据 And Not Exists
   (Select 1
         From 药品收发记录
         Where 单据 = a.单据 And Nvl(库房id, 0) = Nvl(a.库房id, 0) And NO = v_No And Mod(记录状态, 3) = 1 And 审核人 Is Null);

  --发送销帐消息
  If 消息_In = 1 Then
    b_Message.Zlhis_Charge_008(v_收费类别, 费用id_In, v_收发ids);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_销售退费;
/

--125779:李业庆,2018-05-15,退药按药品id排序处理
Create Or Replace Procedure Zl_输液配药记录_销帐审核
(
  配药id_In   In Varchar2, --ID串:ID1,审核标志1,ID2,审核标志2....
  操作人员_In In 输液配药记录.操作人员%Type,
  操作时间_In In 输液配药记录.操作时间%Type
) Is
  v_Tansid     Varchar2(20);
  v_Tmp        Varchar2(4000);
  v_Usercode   Varchar2(100);
  v_发药id     药品收发记录.Id%Type;
  n_Count      Number(1);
  d_审核时间   药品收发记录.审核日期%Type;
  v_No         药品收发记录.No%Type;
  v_上次no     药品收发记录.No%Type;
  n_审核标志   Number(1);
  n_操作状态   Number(2);
  v_收发ids    Varchar2(4000);
  v_退药待发id 药品收发记录.Id%Type;
  v_原始id     药品收发记录.Id%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;

  Cursor c_销帐记录 Is
    Select Distinct a.费用id, b.操作时间
    From 药品收发记录 A, 输液配药记录 B, 输液配药内容 C
    Where a.Id = c.收发id And b.Id = c.记录id And b.Id = v_Tansid And b.操作状态 = 9;

  v_销帐记录 c_销帐记录%RowType;

  Cursor c_退药记录 Is
    Select Distinct a.Id As 退药id, c.收发id, c.数量, a.药品id, a.批次
    From 药品收发记录 A, 药品收发记录 B, 输液配药内容 C
    Where c.记录id = v_Tansid And b.Id = c.收发id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And
          a.药品id + 0 = b.药品id And a.序号 = b.序号 And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
    Order By a.药品id, a.批次;

  v_退药记录 c_退药记录%RowType;

  Cursor c_费用销帐 Is
    Select a.No, a.序号 || ':' || c.数量 || ':' || c.记录id As 费用序号
    From 住院费用记录 A, 药品收发记录 B, 输液配药内容 C
    Where a.Id = b.费用id And b.Id = c.收发id And Mod(b.记录状态, 3) = 1 And c.记录id = v_Tansid;

  v_费用销帐 c_费用销帐%RowType;

Begin
  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;

  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Tansid   := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
    n_审核标志 := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
  
    v_收发ids := Null;
  
    --检查当前输液单的状态是否为待摆药状态
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;
    
      If n_操作状态 <> 9 Then
        v_Error := '该数据已被操作，不能进行销帐审核！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    If n_审核标志 = 1 Then
      n_操作状态 := 10;
    Elsif n_审核标志 = 2 Then
      n_操作状态 := 11;
    End If;
  
    --查找输液单对应的收发NO
    Begin
      Select NO
      Into v_No
      From 药品收发记录
      Where ID In (Select 收发id From 输液配药内容 Where 记录id In (Select ID From 输液配药记录 Where ID = v_Tansid)) And Rownum < 2;
    Exception
      When Others Then
        v_No := '';
    End;
  
    --收发NO相同的配药ID，审核时间以此设置为延长1秒
    If v_No = v_上次no Then
      d_审核时间 := d_审核时间 + 1 / 24 / 60 / 60;
    Else
      d_审核时间 := 操作时间_In;
      v_上次no   := v_No;
    End If;
  
    --销帐记录处理
    For v_销帐记录 In c_销帐记录 Loop
      Zl_病人费用销帐_Audit(v_销帐记录.费用id, v_销帐记录.操作时间, 操作人员_In, d_审核时间, n_审核标志);
    End Loop;
  
    Select Count(*) Into n_Count From 输液配药状态 Where 配药id = v_Tansid And 操作时间 = 操作时间_In;
  
    If n_Count <> 1 Then
      Insert Into 输液配药状态
        (配药id, 操作类型, 操作人员, 操作时间)
      Values
        (v_Tansid, n_操作状态, 操作人员_In, 操作时间_In);
    End If;
    Update 输液配药记录 Set 操作人员 = 操作人员_In, 操作时间 = 操作时间_In, 操作状态 = n_操作状态 Where ID = v_Tansid;
  
    --退药处理
    If n_审核标志 = 1 Then
      For v_退药记录 In c_退药记录 Loop
        Zl_药品收发记录_部门退药(v_退药记录.退药id, 操作人员_In, 操作时间_In, Null, Null, Null, v_退药记录.数量, Null, 操作人员_In);
      
        --取退药待发id
        Select a.Id
        Into v_发药id
        From 药品收发记录 A, 药品收发记录 B
        Where b.Id = v_退药记录.退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
              a.序号 = b.序号 And Mod(a.记录状态, 3) = 1 And a.审核日期 Is Null;
      
        --输液配药内容中的收发ID更新为退药待发的收发ID
        Update 输液配药内容 Set 收发id = v_发药id Where 记录id = v_Tansid And 收发id = v_退药记录.收发id;
      
        If v_收发ids Is Null Then
          v_收发ids := v_发药id;
        Else
          v_收发ids := v_收发ids || ',' || v_发药id;
        End If;
      
        --取原始id
        Select a.Id
        Into v_原始id
        From 药品收发记录 A, 药品收发记录 B
        Where b.Id = v_退药记录.退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
              a.序号 = b.序号 And Mod(a.记录状态, 3) = 0 And a.审核日期 Is Not Null;
      
        Insert Into 输液配药内容
          (记录id, 收发id, 数量)
          Select 记录id, v_原始id, 数量 From 输液配药内容 Where 记录id = v_Tansid And 收发id = v_发药id;
      
        v_收发ids := v_收发ids || ',' || v_原始id;
      End Loop;
    
      --费用销帐
      For v_费用销帐 In c_费用销帐 Loop
        Zl_住院记帐记录_Delete(v_费用销帐.No, v_费用销帐.费用序号, v_Usercode, Zl_Username, 2, 1, 1, d_审核时间);
      End Loop;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_销帐审核;
/




------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0011' Where 编号=&n_System;
Commit;
