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
--128798:蒋廷中,2018-08-03,处理执行时间显示小时分钟
Create Or Replace Procedure Zl_门诊医嘱执行_Cancel
(
  医嘱id_In     病人医嘱执行.医嘱id%Type,
  发送号_In     病人医嘱执行.发送号%Type,
  单独执行_In   Number,
  操作员编号_In 人员表.编号%Type,
  操作员姓名_In 人员表.姓名%Type,
  组id_In       病人医嘱执行.医嘱id%Type,
  诊疗类别_In   病人医嘱记录.诊疗类别%Type,
  执行部门id_In 门诊费用记录.执行部门id%Type
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。 
  --      单独执行_In=检验医嘱组合是否采用对每个项目分散单独执行的方式 
) Is
  --医嘱相关的费用单据 
  Cursor c_No Is
    Select a.No || ':' || a.记录性质
    From 病人医嘱附费 A, 病人医嘱记录 B
    Where a.发送号 + 0 = 发送号_In And a.医嘱id = b.Id And b.诊疗类别 = 诊疗类别_In And (b.Id = 组id_In Or b.相关id = 组id_In)
    Union
    Select NO || ':' || 记录性质 From 病人医嘱发送 Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;

  Cursor c_Noone Is
    Select NO || ':' || 记录性质
    From 病人医嘱附费
    Where 发送号 + 0 = 发送号_In And 医嘱id = 医嘱id_In
    Union
    Select NO || ':' || 记录性质 From 病人医嘱发送 Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;

  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;

  n_执行次数 Number;
  n_剩余次数 Number;
  d_执行时间   Date;
  n_执行状态 Number;

  --要取消执行的费用行(不包含药品和跟踪在用的卫材) 
  Cursor c_Finish(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select Distinct a.Id, a.收费类别, a.收费细目id
           From 门诊费用记录 A, 病人医嘱记录 B,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
                  From Table(r_No)) N
           Where a.医嘱序号 = b.Id And (b.Id = 组id_In Or b.相关id = 组id_In) And a.No = n.No And
                 (a.记录性质 = n.记录性质 Or a.记录性质 = 11 And n.记录性质 = 1) And a.记录状态 In (0, 1, 3) And
                 (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)) A, 材料特性 B
    Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);

  Cursor c_Finishone(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select a.Id, a.收费类别, a.收费细目id
           From 门诊费用记录 A,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
                  From Table(r_No)) N
           Where a.医嘱序号 = 医嘱id_In And a.No = n.No And (a.记录性质 = n.记录性质 Or a.记录性质 = 11 And n.记录性质 = 1) And
                 a.记录状态 In (0, 1, 3) And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)) A, 材料特性 B
    Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);

  --取消执行中包含跟踪在用的发料卫料时，根据参数设置是否自动退料 
  --卫生材料医嘱目前不存在单独和组合执行的情况 
  Cursor c_Stuff(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id
    From 门诊费用记录 A, 药品收发记录 B, 病人医嘱记录 C,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
           From Table(r_No)) N
    Where a.Id = b.费用id And (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And b.审核人 Is Not Null And a.收费类别 = '4' And a.记录状态 = 1 And
          a.医嘱序号 = c.Id And c.诊疗类别 = 诊疗类别_In And (c.Id = 组id_In Or c.相关id = 组id_In) And a.No = n.No And
          (a.记录性质 = n.记录性质 Or a.记录性质 = 11 And n.记录性质 = 1) And b.单据 IN(24,25,26) And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)
    Order By b.药品id;
Begin
  Open c_Noone;
  Fetch c_Noone Bulk Collect
    Into r_No_Stuff;
  Close c_Noone;

  If Nvl(单独执行_In, 0) = 0 Then
    Open c_No;
    Fetch c_No Bulk Collect
      Into r_No;
    Close c_No;
  Else
    r_No := r_No_Stuff;
  End If;

  --主费用可能需要限制医嘱序号 
  --不包含药品和跟踪在用的卫材，因为这些都要发放才表示执行 
  If Nvl(单独执行_In, 0) = 0 Then
    Open c_Finish(r_No);
    Fetch c_Finish Bulk Collect
      Into r_Finish;
    Close c_Finish;
  Else
    Open c_Finishone(r_No);
    Fetch c_Finishone Bulk Collect
      Into r_Finish;
    Close c_Finishone;
  End If;

  Select Decode(a.执行状态, 1, a.发送数次, c.登记次数), Decode(a.执行状态, 1, 0, a.发送数次 - c.登记次数)
  Into n_执行次数, n_剩余次数
  From 病人医嘱发送 A,
       (Select 医嘱id_In 医嘱id, 发送号_In 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数
         From 病人医嘱执行 B
         Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In And Nvl(b.执行结果, 1) <> 0) C
  Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In;

  --如果全部执行则状态为1，未执行状态为0，部分执行状态为2 
  Select Decode(n_剩余次数, 0, 1, Decode(n_执行次数, 0, 0, 2)) Into n_执行状态 From Dual;

  --对于门诊单据（包含记账与收费）部分执行（2）与完全执行（1）,执行时间为执行完成的执行时间，执行人为执行完成的执行人 
  Forall I In 1 .. r_Finish.Count
    Update 门诊费用记录
    Set 执行状态 = n_执行状态, 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间), 执行人 = Decode(n_执行状态, 0, Null, 执行人)
    Where ID = r_Finish(I);

  --处理跟踪在用卫材自动发料 
  For r_Stuff In c_Stuff(r_No_Stuff) Loop
    Zl_材料收发记录_部门退料(r_Stuff.Id, 操作员姓名_In, Sysdate, Null, Null, Null, Null, 0, 操作员姓名_In);
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊医嘱执行_Cancel;
/

--128798:蒋廷中,2018-08-03,处理执行时间显示小时分钟
Create Or Replace Procedure Zl_住院医嘱执行_Cancel
(
  医嘱id_In     病人医嘱执行.医嘱id%Type,
  发送号_In     病人医嘱执行.发送号%Type,
  单独执行_In   Number,
  操作员编号_In 人员表.编号%Type,
  操作员姓名_In 人员表.姓名%Type,
  组id_In       病人医嘱执行.医嘱id%Type,
  诊疗类别_In   病人医嘱记录.诊疗类别%Type,
  执行部门id_In 门诊费用记录.执行部门id%Type
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --      单独执行_In=检验医嘱组合是否采用对每个项目分散单独执行的方式
) Is
  --医嘱相关的费用单据
  Cursor c_No Is
    Select a.No || ':' || a.记录性质
    From 病人医嘱附费 A, 病人医嘱记录 B
    Where a.发送号 + 0 = 发送号_In And a.医嘱id = b.Id And b.诊疗类别 = 诊疗类别_In And (b.Id = 组id_In Or b.相关id = 组id_In)
    Union
    Select NO || ':' || 记录性质 From 病人医嘱发送 Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;

  Cursor c_Noone Is
    Select NO || ':' || 记录性质
    From 病人医嘱附费
    Where 发送号 + 0 = 发送号_In And 医嘱id = 医嘱id_In
    Union
    Select NO || ':' || 记录性质 From 病人医嘱发送 Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;
  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;

  n_执行次数 Number;
  n_剩余次数 Number;
  d_执行时间   Date;
  n_执行状态 Number;
  n_Count    Number;
  --要取消执行的费用行(不包含药品和跟踪在用的卫材)
  Cursor c_Finish(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select Distinct a.Id, a.收费类别, a.收费细目id
           From 住院费用记录 A, 病人医嘱记录 B,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
                  From Table(r_No)) N
           Where a.医嘱序号 = b.Id And (b.Id = 组id_In Or b.相关id = 组id_In) And a.No = n.No And a.记录性质 = n.记录性质 And
                 a.记录状态 In (0, 1, 3) And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)) A, 材料特性 B
    Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);

  Cursor c_Finishone(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select a.Id, a.收费类别, a.收费细目id
           From 住院费用记录 A,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
                  From Table(r_No)) N
           Where a.医嘱序号 = 医嘱id_In And a.No = n.No And a.记录性质 = n.记录性质 And a.记录状态 In (0, 1, 3) And
                 (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)) A, 材料特性 B
    Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);

  --取消执行中包含跟踪在用的发料卫料时，根据参数设置是否自动退料
  --卫生材料医嘱目前不存在单独和组合执行的情况
  Cursor c_Stuff(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id
    From 住院费用记录 A, 药品收发记录 B, 病人医嘱记录 C,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
           From Table(r_No)) N
    Where a.Id = b.费用id And (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And b.审核人 Is Not Null And a.收费类别 = '4' And a.记录状态 = 1 And
          a.医嘱序号 = c.Id And c.诊疗类别 = 诊疗类别_In And (c.Id = 组id_In Or c.相关id = 组id_In) And a.No = n.No And a.记录性质 = n.记录性质 And
          b.单据 In (25, 26) And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)
    Order By b.药品id;

  --取消执行中包含药品时，本科执行的自动退药
  Cursor c_Drug(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id
    From 住院费用记录 A, 药品收发记录 B, 病人医嘱记录 C, 病案主页 D,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
           From Table(r_No)) N
    Where a.Id = b.费用id And (b.记录状态 = 1 Or Mod(b.记录状态, 3) = 0) And b.审核人 Is Not Null And b.库房id = 执行部门id_In And
          a.收费类别 In ('5', '6', '7') And a.记录状态 = 1 And a.医嘱序号 = c.Id And c.诊疗类别 = 诊疗类别_In And c.病人id = d.病人id And
          c.主页id = d.主页id And (c.Id = 组id_In Or c.相关id = 组id_In) And a.No = n.No And a.记录性质 = n.记录性质
    Order By b.药品id;

  v_医嘱期效 病人医嘱记录.医嘱期效%Type;
Begin
  Open c_Noone;
  Fetch c_Noone Bulk Collect
    Into r_No_Stuff;
  Close c_Noone;

  If Nvl(单独执行_In, 0) = 0 Then
    Open c_No;
    Fetch c_No Bulk Collect
      Into r_No;
    Close c_No;
  Else
    r_No := r_No_Stuff;
  End If;

  --主费用可能需要限制医嘱序号
  --不包含药品和跟踪在用的卫材，因为这些都要发放才表示执行
  If Nvl(单独执行_In, 0) = 0 Then
    Open c_Finish(r_No);
    Fetch c_Finish Bulk Collect
      Into r_Finish;
    Close c_Finish;
  Else
    Open c_Finishone(r_No);
    Fetch c_Finishone Bulk Collect
      Into r_Finish;
    Close c_Finishone;
  End If;

  Select Decode(a.执行状态, 1, a.发送数次, c.登记次数), Decode(a.执行状态, 1, 0, a.发送数次 - c.登记次数)
  Into n_执行次数, n_剩余次数
  From 病人医嘱发送 A,
       (Select 医嘱id_In 医嘱id, 发送号_In 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数
         From 病人医嘱执行 B
         Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In And Nvl(b.执行结果, 1) <> 0) C
  Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In;

  --如果全部执行则状态为1，未执行状态为0，部分执行状态为2
  Select Decode(n_剩余次数, 0, 1, Decode(n_执行次数, 0, 0, 2)) Into n_执行状态 From Dual;

  Forall I In 1 .. r_Finish.Count
    Update 住院费用记录
    Set 执行状态 = n_执行状态, 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间), 执行人 = Decode(n_执行状态, 0, Null, 执行人)
    Where ID = r_Finish(I);

  --处理跟踪在用卫材自动发料
  For r_Stuff In c_Stuff(r_No_Stuff) Loop
    Zl_材料收发记录_部门退料(r_Stuff.Id, 操作员姓名_In, Sysdate, Null, Null, Null, Null, 0, 操作员姓名_In);
  End Loop;

  --处理药品自动发药(只在护士站，本科药品才处理,本科由参数和游标判断)
  Select Max(a.医嘱期效), Max(Decode(b.病区id, 执行部门id_In, 1, 0))
  Into v_医嘱期效, n_Count
  From 病人医嘱记录 A, 病人变动记录 B
  Where a.病人id = b.病人id And a.主页id = b.主页id And a.Id = 医嘱id_In;

  If Substr(Zl_Getsysparameter('本科执行自动完成', 1254), v_医嘱期效 + 1, 1) = '1' And n_Count = 1 Then
    For r_Drug In c_Drug(r_No_Stuff) Loop
      Zl_药品收发记录_部门退药(r_Drug.Id, 操作员姓名_In, Sysdate);
    End Loop;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_住院医嘱执行_Cancel;
/

--128798:蒋廷中,2018-08-03,处理执行时间显示小时分钟
Create Or Replace Procedure Zl_病人医嘱执行_Update
(
  原执行时间_In 病人医嘱执行.执行时间%Type,
  医嘱id_In     病人医嘱执行.医嘱id%Type,
  发送号_In     病人医嘱执行.发送号%Type,
  要求时间_In   病人医嘱执行.要求时间%Type,
  本次数次_In   病人医嘱执行.本次数次%Type,
  执行摘要_In   病人医嘱执行.执行摘要%Type,
  执行人_In     病人医嘱执行.执行人%Type,
  执行时间_In   病人医嘱执行.执行时间%Type,
  执行结果_In   病人医嘱执行.执行结果%Type := 1,
  未执行原因_In 病人医嘱执行.说明%Type := Null,
  单独执行_In   Number := 0,
  操作员编号_In 人员表.编号%Type := Null,
  操作员姓名_In 人员表.姓名%Type := Null,
  执行部门id_In 门诊费用记录.执行部门id%Type := 0
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门
) Is
  --除了要执行的主记录,还包含了附加手术,检查部位的记录
  --手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同
  v_Temp     Varchar2(255); 
  v_人员姓名 人员表.姓名%Type;

  v_组id        病人医嘱记录.Id%Type;
  v_诊疗类别    病人医嘱记录.诊疗类别%Type;
  v_执行结果old 病人医嘱执行.执行结果%Type;
  n_本次数次old 病人医嘱执行.本次数次%Type;

  v_病人来源 病人医嘱记录.病人来源%Type;
  v_费用性质 病人医嘱发送.记录性质%Type;

  n_执行次数 Number;
  n_剩余次数 Number;
  n_执行状态 Number;
  n_发送数次 Number;
  n_单次数次 Number;
  v_Count    Number;
  n_登记数次 Number;
  d_要求时间 Date;
  d_执行时间 Date;

  d_登记时间   病人医嘱执行.登记时间%Type;
  n_取消执行   Number;
  n_Diffday    Number(18, 3);
  n_执行科室id Number;

  v_Date  Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --当前操作人员
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  Select Sysdate Into v_Date From Dual;
  Select Nvl(执行结果, 1), Nvl(本次数次, 0), 登记时间
  Into v_执行结果old, n_本次数次old, d_登记时间
  From 病人医嘱执行
  Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 执行时间 = 原执行时间_In;
  -----取消执行有效天数限制
  Select Zl_To_Number(Nvl(zl_GetSysParameter(220), '999')) Into n_取消执行 From Dual;
  Select v_Date - d_登记时间 Into n_Diffday From Dual;
  --登记时间超过取消执行天数的记录，不允许修改医嘱执行情况
  If n_Diffday > n_取消执行 Then
    v_Error := '医嘱执行登记时间超过了取消执行有效天数，不能修改医嘱执行情况！';
    Raise Err_Custom;
  End If;
  Select 执行部门id Into n_执行科室id From 病人医嘱发送 Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In;
  --病人医嘱执行
  Update 病人医嘱执行
  Set 要求时间 = 要求时间_In, 本次数次 = 本次数次_In, 执行摘要 = 执行摘要_In, 执行人 = 执行人_In, 执行时间 = 执行时间_In, 登记时间 = v_Date, 登记人 = v_人员姓名,
      执行结果 = 执行结果_In, 说明 = 未执行原因_In, 执行科室id = n_执行科室id
  Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 执行时间 = 原执行时间_In;
  --本次执行次数或这执行结果修改后需要更新单据的执行状态
  If v_执行结果old <> 执行结果_In Or n_本次数次old <> 本次数次_In Then
    Select 病人来源, Nvl(相关id, ID), 诊疗类别
    Into v_病人来源, v_组id, v_诊疗类别
    From 病人医嘱记录
    Where ID = 医嘱id_In;
  
    If v_病人来源 = 2 Then
      Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
      Into v_费用性质
      From 病人医嘱发送
      Where 发送号 = 发送号_In And 医嘱id = 医嘱id_In;
    Else
      v_费用性质 := 1;
    End If;
  
    Select Decode(a.执行状态, 1, a.发送数次, c.登记次数), Decode(a.执行状态, 1, 0, a.发送数次 - c.登记次数), a.发送数次, c.登记次数


    Into n_执行次数, n_剩余次数, n_发送数次, n_登记数次
    From 病人医嘱发送 A,
         (Select 医嘱id_In 医嘱id, 发送号_In 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数
           From 病人医嘱执行 B
           Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In And Nvl(b.执行结果, 1) = 1) C
    Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In;
  
    --如果全部执行则状态为1，未执行状态为0，部分执行状态为2
    Select Decode(n_剩余次数, 0, 1, Decode(n_执行次数, 0, 0, 2)) Into n_执行状态 From Dual;
  
    --更新医嘱执行计价.执行状态
    If n_发送数次 > 0 Then
      Select Count(Distinct 要求时间) Into v_Count From 医嘱执行计价 Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In;
      If v_Count > 0 Then
        n_单次数次 := n_发送数次 / v_Count;
        --已执行数量+本次数次 总共能够执行多少个时间点,取最大整数
        v_Count := Ceil((n_登记数次) / n_单次数次);
        If n_登记数次 = 0 Then
          Update 医嘱执行计价
          Set 执行状态 = 0
          Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And Nvl(执行状态, 0) <> 2;
        Else
          --获取执行截至要求时间
          Select 要求时间
          Into d_要求时间
          From (Select 要求时间, Rownum As 次数
                 From (Select Distinct 要求时间
                        From 医嘱执行计价
                        Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In
                        Order By 要求时间))
          Where 次数 = v_Count;
        
          If Not d_要求时间 Is Null Then
            --先检查是否已经退费
            Select Max(Nvl(执行状态, 0))
            Into v_Count
            From 医嘱执行计价
            Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 <= d_要求时间;
            If v_Count = 2 Then
              v_Error := '您指定的执行时间段的医嘱费用已经被退费，不允许再执行。';
              Raise Err_Custom;
            End If;
            --更新截至要求时间之前(含)的记录执行状态；
            Update 医嘱执行计价
            Set 执行状态 = 1
            Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 <= d_要求时间 And Nvl(执行状态, 0) <> 2;
            Update 医嘱执行计价
            Set 执行状态 = 0
            Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 > d_要求时间 And Nvl(执行状态, 0) <> 2;
          End If;
        End If;
      End If;
    End If;
  
    --执行次数不为0就标记为正在执行
    If Nvl(单独执行_In, 0) = 1 Then
      Update 病人医嘱发送
      Set 执行状态 = Decode(n_执行次数, 0, 0, 3), 完成人 = Null, 完成时间 = Null
      Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;
    Else
      Update 病人医嘱发送
      Set 执行状态 = Decode(n_执行次数, 0, 0, 3), 完成人 = Null, 完成时间 = Null
      Where 执行状态 In (0, 3) And 发送号 + 0 = 发送号_In And
            医嘱id In (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id) And 诊疗类别 = v_诊疗类别);
    End If;
  
    If v_费用性质 = 2 Then
      If Nvl(单独执行_In, 0) = 1 Then
        Update 住院费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 住院费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                     医嘱id In
                     (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id) And 诊疗类别 = v_诊疗类别));
      End If;
    Else
      If Nvl(单独执行_In, 0) = 1 Then
        Update 门诊费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 门诊费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间_In)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                     医嘱id In
                     (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id) And 诊疗类别 = v_诊疗类别));
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱执行_Update;
/

--128798:蒋廷中,2018-08-03,处理执行时间显示小时分钟
CREATE OR REPLACE Procedure Zl_病人医嘱执行_Delete
(
  医嘱id_In     In 病人医嘱执行.医嘱id%Type,
  发送号_In     In 病人医嘱执行.发送号%Type,
  执行时间_In   In 病人医嘱执行.执行时间%Type,
  单独执行_In   In Number := 0,
  自动取消_In   In Number := 0,
  执行部门id_In In 门诊费用记录.执行部门id%Type := 0
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门
) Is
  --除了要执行的主记录,还包含了附加手术,检查部位的记录
  --手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同
  v_组id     病人医嘱记录.Id%Type;
  v_诊疗类别 病人医嘱记录.诊疗类别%Type;
  v_病人来源 病人医嘱记录.病人来源%Type;
  v_费用性质 病人医嘱发送.记录性质%Type;
  v_操作类型 诊疗项目目录.操作类型%Type;

  n_病人id   病人医嘱记录.病人id%Type;
  n_主页id   病人医嘱记录.主页id%Type;
  v_挂号单   病人医嘱记录.挂号单%Type;
  n_本次数次 病人医嘱执行.本次数次%Type;
  v_执行摘要 病人医嘱执行.执行摘要%Type;
  n_执行科室 病人医嘱执行.执行科室id%Type;
  v_执行人   病人医嘱执行.执行人%Type;
  v_核对人   病人医嘱执行.核对人%Type;
  n_记录来源 病人医嘱执行.记录来源%Type;

  v_自动取消 Number;
  v_执行状态 Number;

  n_执行次数 Number;
  n_剩余次数 Number;
  n_执行状态 Number;
  n_执行结果 Number;

  n_发送数次 Number;
  n_单次数次 Number;
  v_Count    Number;
  n_登记数次 Number;
  d_要求时间 Date;
  d_执行时间 Date;

  d_登记时间 病人医嘱执行.登记时间%Type;
  n_取消执行 Number;
  n_Diffday  Number(18, 3);
  Err_Custom Exception;
  v_Error Varchar2(2000);
Begin
  Select a.病人来源, Nvl(a.相关id, a.Id), Nvl(a.诊疗类别, '*'), Nvl(b.操作类型, '0') 操作类型, a.病人id, a.主页id, a.挂号单
  Into v_病人来源, v_组id, v_诊疗类别, v_操作类型, n_病人id, n_主页id, v_挂号单
  From 病人医嘱记录 A, 诊疗项目目录 B
  Where a.Id = 医嘱id_In And a.诊疗项目id = b.Id(+);

  Select Nvl(a.执行结果, 1), a.登记时间, a.要求时间, a.本次数次, a.执行摘要, a.执行科室id, a.执行人, a.核对人, a.记录来源
  Into n_执行结果, d_登记时间, d_要求时间, n_本次数次,
       
       v_执行摘要, n_执行科室, v_执行人, v_核对人, n_记录来源
  
  From 病人医嘱执行 A
  Where a.医嘱id = 医嘱id_In And a.发送号 + 0 = 发送号_In And a.执行时间 = 执行时间_In;

  -----取消执行有效天数限制
  Select Zl_To_Number(Nvl(zl_GetSysParameter(220), '999')) Into n_取消执行 From Dual;
  Select Sysdate - d_登记时间 Into n_Diffday From Dual;
  --登记时间超过取消执行天数的记录，不允许删除医嘱执行记录
  If n_Diffday > n_取消执行 Then
    v_Error := '医嘱执行登记时间超过了取消执行有效天数，不能删除医嘱执行记录！';
    Raise Err_Custom;
  End If;

  If v_病人来源 = 2 Then
    Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
    Into v_费用性质
    From 病人医嘱发送
    Where 发送号 = 发送号_In And 医嘱id = 医嘱id_In;
  Else
    v_费用性质 := 1;
  End If;

  --病人医嘱执行
  Delete From 病人医嘱执行 Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In And 执行时间 = 执行时间_In;
  b_Message.Zlhis_Cis_051(n_病人id, n_主页id, v_挂号单, 发送号_In, 医嘱id_In, d_要求时间, 执行时间_In, n_本次数次, n_执行结果, v_执行摘要, n_执行科室,
                          v_执行人, v_核对人, n_记录来源);
  d_要求时间 := Null;

  --对于未执行的医嘱执行记录的删除，不更新医嘱发送以及费用信息的执行状态
  If n_执行结果 <> 0 Then
    Select Decode(a.执行状态, 1, a.发送数次, c.登记次数), Decode(a.执行状态, 1, 0, a.发送数次 - c.登记次数), a.发送数次, c.登记次数


    
    Into n_执行次数, n_剩余次数, n_发送数次, n_登记数次
    From 病人医嘱发送 A,
         (Select 医嘱id_In 医嘱id, 发送号_In 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数
           From 病人医嘱执行 B
           Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In And Nvl(b.执行结果, 1) <> 0) C
    Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In;
  
    --如果全部执行则状态为1，未执行状态为0，部分执行状态为2
    Select Decode(n_剩余次数, 0, 1, Decode(n_执行次数, 0, 0, 2)) Into n_执行状态 From Dual;
  
    --更新医嘱执行计价.执行状态
    If n_发送数次 > 0 Then
      If n_登记数次 > 0 Then
        Select Count(Distinct 要求时间) Into v_Count From 医嘱执行计价 Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In;
        If v_Count > 0 Then
          n_单次数次 := n_发送数次 / v_Count;
          --已执行数量+本次数次 总共能够执行多少个时间点,取最大整数
          v_Count := Ceil((n_登记数次) / n_单次数次);
          --获取执行截至要求时间
          Select 要求时间
          Into d_要求时间
          From (Select 要求时间, Rownum As 次数
                 From (Select Distinct 要求时间
                        From 医嘱执行计价
                        Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In
                        Order By 要求时间))
          Where 次数 = v_Count;
        
          If Not d_要求时间 Is Null Then
            --先检查是否已经退费
            Select Max(Nvl(执行状态, 0))
            Into v_Count
            From 医嘱执行计价
            Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 <= d_要求时间;
            If v_Count = 2 Then
              v_Error := '您指定的执行时间段的医嘱费用已经被退费，不允许再执行。';
              Raise Err_Custom;
            End If;
            --更新截至要求时间之前(含)的记录执行状态；
            Update 医嘱执行计价
            Set 执行状态 = 1
            Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 <= d_要求时间 And Nvl(执行状态, 0) <> 2;
            Update 医嘱执行计价
            Set 执行状态 = 0
            Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 > d_要求时间 And Nvl(执行状态, 0) <> 2;
          End If;
        End If;
      Else
        Update 医嘱执行计价 Set 执行状态 = 0 Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And Nvl(执行状态, 0) <> 2;
      End If;
    End If;
  
    --如果执行情况删除了就更新执行状态
    If Nvl(单独执行_In, 0) = 1 Then
      Update 病人医嘱发送
      Set 执行状态 = Decode(n_执行次数, 0, 0, 3), 完成人 = Null, 完成时间 = Null
      Where 执行状态 = 3 And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;
    Else
      Update 病人医嘱发送 A
      Set 执行状态 = Decode(n_执行次数, 0, 0, 3), 完成人 = Null, 完成时间 = Null
      Where 执行状态 = 3 And 发送号 + 0 = 发送号_In And
            医嘱id In
            (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id) And Nvl(诊疗类别, '*') = v_诊疗类别) And Not Exists
       (Select 1 From 病人医嘱执行 Where 发送号 + 0 = 发送号_In And 医嘱id = a.医嘱id);
    End If;
    --更新对应的费用执行状态为未执行
    --不应该处理药品和跟踪在用的卫材
    If v_费用性质 = 2 Then
      If Nvl(单独执行_In, 0) = 1 Then
        Update 住院费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 住院费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                     医嘱id In
                     (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id) And 诊疗类别 = v_诊疗类别));
      End If;
    Else
      If Nvl(单独执行_In, 0) = 1 Then
        Update 门诊费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 门诊费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人), 执行时间 = Decode(n_执行状态, 0, d_执行时间, 执行时间)
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                     医嘱id In
                     (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id) And 诊疗类别 = v_诊疗类别));
      End If;
    End If;
    --检验采集自动取消采集人采集时间
    If v_诊疗类别 = 'E' And v_操作类型 = '6' Then
      Update 病人医嘱发送 A
      Set a.采样人 = Null, a.采样时间 = Null
      Where 医嘱id In (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id)) And 发送号 = 发送号_In;
    End If;
  
    --已完成执行的，执行数次减少之后自动取消执行完成为正在执行或未执行(主要用于PDA自动执行)
    If Nvl(自动取消_In, 0) = 1 Then
      Begin
        Select Decode(Sign(Nvl(Sum(b.本次数次), 0) - a.发送数次), -1, 1, 0), Decode(Sign(Nvl(Sum(b.本次数次), 0)), 0, 0, 3)
        Into v_自动取消, v_执行状态
        From 病人医嘱发送 A, 病人医嘱执行 B
        Where a.医嘱id = b.医嘱id(+) And a.发送号 = b.发送号(+) And a.执行状态 = 1 And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In
        Group By a.发送数次;
      Exception
        When Others Then
          Null;
      End;
    
      If Nvl(v_自动取消, 0) = 1 Then
        Zl_病人医嘱执行_Cancel(医嘱id_In, 发送号_In, Null, 单独执行_In, 执行部门id_In);
      
        If v_执行状态 = 3 Then
          Select Nvl(相关id, ID), 诊疗类别 Into v_组id, v_诊疗类别 From 病人医嘱记录 Where ID = 医嘱id_In;
          Update 病人医嘱发送
          Set 执行状态 = 3, 完成人 = Null, 完成时间 = Null
          Where 发送号 + 0 = 发送号_In And
                医嘱id In (Select ID From 病人医嘱记录 Where (ID = v_组id Or 相关id = v_组id) And 诊疗类别 = v_诊疗类别);
        End If;
      
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱执行_Delete;
/

--128798:蒋廷中,2018-08-03,处理执行时间显示小时分钟
CREATE OR REPLACE Procedure Zl_病人医嘱执行_Insert
(
  医嘱id_In       In 病人医嘱执行.医嘱id%Type,
  发送号_In       In 病人医嘱执行.发送号%Type,
  要求时间_In     In 病人医嘱执行.要求时间%Type,
  本次数次_In     In 病人医嘱执行.本次数次%Type,
  执行摘要_In     In 病人医嘱执行.执行摘要%Type,
  执行人_In       In 病人医嘱执行.执行人%Type,
  执行时间_In     In 病人医嘱执行.执行时间%Type,
  单独执行_In     In Number := 0,
  自动完成_In     In Number := 0,
  执行结果_In     In 病人医嘱执行.执行结果%Type := 1,
  未执行原因_In   In 病人医嘱执行.说明%Type := Null,
  操作员编号_In   In 人员表.编号%Type := Null,
  操作员姓名_In   In 人员表.姓名%Type := Null,
  执行部门id_In   In 门诊费用记录.执行部门id%Type := 0,
  配液检查_In     In Number := 0,
  检验项目记帐_In In Number := 0,
  输液通道_In     In 病人医嘱执行.输液通道%Type := Null,
  记录来源_In     In 病人医嘱执行.记录来源%Type := Null
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --      执行结果_In=1- 完成   =0  -未执行
  --      如果是台式机调用 操作员编号_In 操作员姓名_In 这两个参数必须传入
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门
  --配液检查_In=移动工作站调用时，是否检查配液信息。
  --检验项目记帐_In=如果是检验项目时，需要记帐但不完成医嘱发送状态
) Is
  --除了要执行的主记录,还包含了附加手术,检查部位的记录
  --手术麻醉,中药煎法,采集方法单独控制,检验组合只填写在第一个项目上，但执行状态相同
  v_组id       病人医嘱记录.Id%Type;
  v_诊疗类别   病人医嘱记录.诊疗类别%Type;
  v_自动完成   Number;
  v_病人来源   病人医嘱记录.病人来源%Type;
  v_费用性质   病人医嘱发送.记录性质%Type;
  v_操作类型   诊疗项目目录.操作类型%Type;
  v_病区id     病案主页.当前病区id%Type;
  v_配液病区   Varchar2(200);
  v_Count      Number;
  v_Temp       Varchar2(255);
  v_人员编号   人员表.编号%Type;
  v_人员姓名   人员表.姓名%Type;
  n_期效       病人医嘱记录.医嘱期效%Type;
  n_诊疗项目id 病人医嘱记录.诊疗项目id%Type;
  v_叮嘱执行   Varchar2(5);
  n_本次数次   病人医嘱执行.本次数次%Type;
  n_病人id     病人医嘱记录.病人id%Type;
  n_主页id     病人医嘱记录.主页id%Type;
  v_挂号单     病人医嘱记录.挂号单%Type;

  n_执行次数   Number;
  n_剩余次数   Number;
  n_执行状态   Number;
  d_终止时间   Date;
  d_开始时间   Date;
  n_发送数次   Number;
  n_登记数次   Number;
  n_单次数次   Number;
  d_要求时间   Date;
  n_执行科室id Number;
  d_执行时间   Date;

  v_Date  Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --并发查检，防止产生多条执行记录
  Begin
    Select (a.发送数次 - c.登记次数) As 剩余数次, a.发送数次, a.执行部门id, Nvl(d.诊疗项目id, 0), c.登记次数
    Into v_Count, n_发送数次, n_执行科室id, n_诊疗项目id, n_登记数次
    From 病人医嘱发送 A,
         (Select 医嘱id_In As 医嘱id, 发送号_In As 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数
           From 病人医嘱执行 B
           Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In) C, 病人医嘱记录 D
    Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.医嘱id = d.Id And a.发送号 = 发送号_In;
  Exception
    When Others Then
      v_Count := 本次数次_In;
  End;
  v_叮嘱执行 := zl_GetSysParameter(288);
  n_本次数次 := 本次数次_In;
  If 本次数次_In > v_Count And (Not (n_诊疗项目id = 0 And v_叮嘱执行 = 1)) Then
    If Round(n_登记数次 + 本次数次_In) = 1 Then
      --表明是输血执行
      n_本次数次 := 1 - n_登记数次;
    Else
      v_Error := '由于并发操作可能已经被他人登记，请刷新后再试。';
      Raise Err_Custom;
    End If;
  End If;
  --当前操作人员
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    v_人员编号 := 操作员编号_In;
    v_人员姓名 := 操作员姓名_In;
  Else
    Begin
      Select 姓名, 编号 Into v_人员姓名, v_人员编号 From 人员表 Where 姓名 = 执行人_In;
    Exception
      When Others Then
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    End;
  End If;
  --对医嘱终止时间进行检查
  Select a.执行终止时间, a.开始执行时间, a.医嘱期效, a.病人id, a.主页id, a.挂号单
  Into d_终止时间, d_开始时间, n_期效, n_病人id, n_主页id, v_挂号单
  From 病人医嘱记录 A
  Where a.Id = 医嘱id_In;
  If Not d_终止时间 Is Null And n_期效 = 0 Then
    If 要求时间_In > d_终止时间 Then
      v_Error := '要求时间超过了医嘱终止时间，请确认医嘱是否提前停止！';
      Raise Err_Custom;
    End If;
  End If;
  If Not d_开始时间 Is Null Then
    If 执行时间_In < d_开始时间 Then
      v_Error := '执行时间必须大于医嘱的开始执行时间''' || To_Char(d_开始时间, 'yyyy-mm-dd HH24:mi:ss') || '''！';
      Raise Err_Custom;
    End If;
  End If;
  Select Sysdate Into v_Date From Dual;
  Select a.病人来源, 执行科室id, Nvl(a.相关id, a.Id), Nvl(a.诊疗类别, '*'), Nvl(b.操作类型, '0') 操作类型
  Into v_病人来源, v_病区id, v_组id, v_诊疗类别, v_操作类型
  From 病人医嘱记录 A, 诊疗项目目录 B
  Where a.Id = 医嘱id_In And a.诊疗项目id = b.Id(+);

  If v_病人来源 = 2 Then
    Select Decode(记录性质, 1, 1, Decode(门诊记帐, 1, 1, 2))
    Into v_费用性质
    From 病人医嘱发送
    Where 发送号 = 发送号_In And 医嘱id = 医嘱id_In;
  Else
    v_费用性质 := 1;
  End If;

  --移动系统配液检查
  If 配液检查_In = 1 Then
    --检查当前病人所属病区是否进行配液登记管理
    Select Nvl(zl_GetSysParameter(184), '') Into v_配液病区 From Dual;
  
    If v_配液病区 Is Not Null And 执行结果_In <> 0 Then
      If Instr(',' || v_配液病区 || ',', ',' || v_病区id || ',') > 0 Then
        v_病区id   := 0;
        v_配液病区 := 'Select 1 From 病区配液记录 where 医嘱ID=:YZID AND 发送号=:FSH AND 要求时间=:YQSJ';
        Begin
          Execute Immediate v_配液病区
            Into v_病区id
            Using 医嘱id_In, 发送号_In, 要求时间_In;
        Exception
          When Others Then
            Null;
        End;
        If v_病区id = 0 Then
          v_Error := '当前医嘱还未进行配液，不允许进行执行登记！';
          Raise Err_Custom;
        End If;
      End If;
    End If;
    --检查当前医嘱是否已配液
  End If;

  --病人医嘱执行
  Select Count(1)
  Into v_Count
  From 病人医嘱执行
  Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 执行时间 = 执行时间_In;
  If v_Count > 0 Then
    v_Error := '您指定的执行时间，已经执行过本条医嘱，请更改一个执行时间。';
    Raise Err_Custom;
  End If;
  Insert Into 病人医嘱执行
    (医嘱id, 发送号, 要求时间, 本次数次, 执行摘要, 执行人, 执行时间, 登记时间, 登记人, 执行结果, 说明, 输液通道, 执行科室id, 记录来源)
  Values
    (医嘱id_In, 发送号_In, 要求时间_In, n_本次数次, 执行摘要_In, 执行人_In, 执行时间_In, v_Date, v_人员姓名, 执行结果_In, 未执行原因_In, 输液通道_In, n_执行科室id,
     记录来源_In);

  b_Message.Zlhis_Cis_050(n_病人id, n_主页id, v_挂号单, 发送号_In, 医嘱id_In, 要求时间_In, 执行时间_In);
  
  --费用记录的执行状态进行更新
  Select Decode(a.执行状态, 1, a.发送数次, c.登记次数), Decode(a.执行状态, 1, 0, a.发送数次 - c.登记次数), c.登记次数
  Into n_执行次数, n_剩余次数, n_登记数次
  From 病人医嘱发送 A,
       (Select 医嘱id_In 医嘱id, 发送号_In 发送号, Nvl(Sum(b.本次数次), 0) As 登记次数
         From 病人医嘱执行 B
         Where b.医嘱id = 医嘱id_In And b.发送号 = 发送号_In And Nvl(b.执行结果, 1) = 1) C
  Where a.医嘱id = c.医嘱id And a.发送号 = c.发送号 And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In;
  --如果全部执行则状态为1，未执行状态为0，部分执行状态为2
  Select Decode(n_剩余次数, 0, 1, Decode(n_执行次数, 0, 0, 2)) Into n_执行状态 From Dual;

  --填写了执行状态后就标记为正在执行
  If Nvl(单独执行_In, 0) = 1 Then
    Update 病人医嘱发送
    Set 执行状态 = Decode(n_执行次数, 0, 0, 3)
    Where 执行状态 In (0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;
  Else
    Update 病人医嘱发送
    Set 执行状态 = Decode(n_执行次数, 0, 0, 3)
    Where 执行状态 In (0, 3) And 发送号 + 0 = 发送号_In And
          医嘱id In (Select ID
                   From 病人医嘱记录
                   Where ID = v_组id And Nvl(诊疗类别, '*') = v_诊疗类别
                   Union All
                   Select ID
                   From 病人医嘱记录
                   Where 相关id = v_组id And Nvl(诊疗类别, '*') = v_诊疗类别);
  End If;

  --更新对应的费用执行状态为已执行(无正在执行)
  --不应该处理药品和跟踪在用的卫材
  If 执行结果_In = 1 Then
    If n_执行状态 != 0 Then
      d_执行时间 := 执行时间_In;
    End If;
    If v_费用性质 = 2 Then
      If Nvl(单独执行_In, 0) = 1 Then
        Update 住院费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = d_执行时间
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 住院费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = d_执行时间
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In (Select 医嘱id, NO, 记录性质
                                   From 病人医嘱发送
                                   Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                                         医嘱id In (Select ID
                                                  From 病人医嘱记录
                                                  Where ID = v_组id And 诊疗类别 = v_诊疗类别
                                                  Union All
                                                  Select ID
                                                  From 病人医嘱记录
                                                  Where 相关id = v_组id And 诊疗类别 = v_诊疗类别));
      End If;
    Else
      If Nvl(单独执行_In, 0) = 1 Then
        --对于门诊单据n_执行状态可能为0（登记执行情况，选择执行结果为未执行），因此需判断
        Update 门诊费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = d_执行时间
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In
              (Select 医嘱id, NO, 记录性质
               From 病人医嘱发送
               Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In);
      Else
        Update 门诊费用记录 A
        Set 执行状态 = n_执行状态, 执行人 = Decode(n_执行状态, 0, Null, 执行人_In), 执行时间 = d_执行时间
        Where 收费类别 Not In ('5', '6', '7') And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And Not Exists
         (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) And a.记录状态 In (0, 1, 3) And
              (医嘱序号, NO, 记录性质) In (Select 医嘱id, NO, 记录性质
                                   From 病人医嘱发送
                                   Where 执行状态 = Decode(n_执行次数, 0, 0, 3) And 发送号 + 0 = 发送号_In And
                                         医嘱id In (Select ID
                                                  From 病人医嘱记录
                                                  Where ID = v_组id And 诊疗类别 = v_诊疗类别
                                                  Union All
                                                  Select ID
                                                  From 病人医嘱记录
                                                  Where 相关id = v_组id And 诊疗类别 = v_诊疗类别));
      End If;
    End If;
    --检验自动完成采集
    If v_诊疗类别 = 'E' And v_操作类型 = '6' Then
      Update 病人医嘱发送 A
      Set a.采样人 = 执行人_In, a.采样时间 = 执行时间_In
      Where 医嘱id In
            (Select ID From 病人医嘱记录 Where ID = v_组id Union All Select ID From 病人医嘱记录 Where 相关id = v_组id) And
            发送号 = 发送号_In;
    End If;
  
    --执行数次达到之后自动完成执行(主要用于PDA自动执行)，如果启用了移动临床，则护士站和PDA一致。
    v_自动完成 := 自动完成_In;
    If 自动完成_In = 1 Then
      --医嘱已经是完成状态则不用再调用执行完成过程此处先设为不自动完成
      Select Max(a.执行状态) Into v_Count From 病人医嘱发送 A Where a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In;
      If v_Count = 1 Then
        v_自动完成 := 0;
      End If;
      v_Count := Null;
    End If;
  
    If Nvl(v_自动完成, 0) = 0 And (v_病人来源 = 2 Or v_病人来源 = 1) And Instr('C,D', v_诊疗类别) = 0 Then
      Begin
        Execute Immediate 'Select Count(1) From ZLMBSYSTEMS'
          Into v_Count;
      Exception
        When Others Then
          Null;
      End;
      If v_Count > 0 Then
        v_自动完成 := 1;
      End If;
    End If;
  
    If Nvl(v_自动完成, 0) = 1 Or 检验项目记帐_In = 1 Then
      Begin
        Select Decode(Sign(Nvl(Sum(b.本次数次), 0) - a.发送数次), 1, 1, 0, 1, 0)
        Into v_自动完成
        From 病人医嘱发送 A, 病人医嘱执行 B
        Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And a.执行状态 In (0, 3) And a.医嘱id = 医嘱id_In And a.发送号 = 发送号_In
        Group By a.发送数次;
      Exception
        When Others Then
          Null;
      End;
    
      If Nvl(v_自动完成, 0) = 1 Or 检验项目记帐_In = 1 Then
        Zl_病人医嘱执行_Finish(医嘱id_In, 发送号_In, Null, 单独执行_In, v_人员编号, v_人员姓名, 执行部门id_In, 检验项目记帐_In);
      End If;
    End If;
    --更新医嘱执行计价.执行状态
    If n_发送数次 > 0 Then
      Select Count(Distinct 要求时间) Into v_Count From 医嘱执行计价 Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In;
      If v_Count > 0 Then
        n_单次数次 := n_发送数次 / v_Count;
        --已执行数量+本次数次 总共能够执行多少个时间点,取最大整数
        v_Count := Ceil((n_登记数次) / n_单次数次);
        --获取执行截至要求时间
        Select 要求时间
        Into d_要求时间
        From (Select 要求时间, Rownum As 次数
               From (Select Distinct 要求时间
                      From 医嘱执行计价
                      Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In
                      Order By 要求时间))
        Where 次数 = v_Count;
      
        If Not d_要求时间 Is Null Then
          --先检查是否已经退费
          Select Max(Nvl(执行状态, 0))
          Into v_Count
          From 医嘱执行计价
          Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 <= d_要求时间;
          If v_Count = 2 Then
            v_Error := '您指定的执行时间段的医嘱费用已经被退费，不允许再执行。';
            Raise Err_Custom;
          End If;
          --更新截至要求时间之前(含)的记录执行状态；
          Update 医嘱执行计价
          Set 执行状态 = 1
          Where 医嘱id = 医嘱id_In And 发送号 = 发送号_In And 要求时间 <= d_要求时间 And Nvl(执行状态, 0) <> 2;
        End If;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱执行_Insert;
/

--129387:王煜,2018-08-03,返回的病人基本信息不能满足需求
Create Or Replace Procedure Zl_Third_Getpatiinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:获取病人基本信息
  --入参:Xml_In: 
  --  <IN> 
  --      <BRID></BRID>     --病人ID
  --      <SFZH></SFZH>     --身份证号
  --      <CXKLB></CXKLB>   --查询卡类别
  --      <MZH></MZH>       --门诊号
  --      <GHDH></GHDH>     --挂号单号
  --      <YLKLB></YLKLB>   --医疗卡类别，ID或者名称
  --      <YLKH></YLKH>     --医疗卡号
  --      <BRXM></BRXM>     --病人姓名
  --  </IN> 
  --  病人识别顺序:
  --  1.传入病人ID ,以病人ID为准
  --  2.传入挂号单，则以挂号单为准
  --  3.传入卡类别ID，则以卡类别ID和卡号为准
  --  4.传入门诊号，则以门诊号为准
  --  5.传入身份证，则以身份证号和姓名为准  
  --出参:Xml_Out 
  -- <OUTPUT>
  --   <BR>
  --     <BRID></BRID>       --病人ID
  --     <XM></XM>           --姓名
  --     <XB></XB>           --性别
  --     <Nl></NL>           --年龄
  --     <CSRQ></CSRQ>       --出生日期
  --     <MZH></MZH>         --门诊号
  --     <HY></HY>           --婚姻
  --     <GJ></GJ>           --国籍
  --     <MZ></MZ>           --民族
  --     <XL></XL>           --学历
  --     <SF></SF>           --身份
  --     <ZY></ZY>           --职业
  --     <SFZH></SFZH>       --身份证号
  --     <FKFS></FKFS>       --付款方式
  --     <LXFS></LXFS>       --联系方式
  --     <LXRXM></LXRXM>     --联系人姓名
  --     <LXRDH></LXRDH>     --联系人电话
  --     <LXRDZ></LXRDZ>     --联系人地址
  --     <LXDH></LXDH>       --联系电话
  --     <XJZDZ></XJZDZ>     --现居住地址 
  --     <HJDZ></HJDZ>       --户籍地址
  --     <CSDD></CSDD>       --出生地点
  --     <KSID></KSID>       --科室ID
  --     <CXKH></CXKH>       --查询卡号
  --     <GMS></GMS>         --过敏史         
  --     <GHD></GHD>         --挂号单号
  --     <GHSJ></GHSJ>       --挂号时间
  --     <JZSJ></JZSJ>       --就诊时间
  --     <JZKS></JZKS>       --就诊科室
  --     <JZYS></JZYS>       --就诊医生
  --   </BR>
  -- </OUTPUT>
  -------------------------------------------------------------------------------------------------- 

  v_病人id       Varchar2(30000);
  v_医疗卡       Varchar2(500);
  v_门诊号       Varchar2(500);
  v_挂号单       病人挂号记录.No%Type;
  v_卡号         病人医疗卡信息.卡号%Type;
  v_姓名         病人信息.姓名%Type;
  v_身份证号     病人信息.身份证号%Type;
  v_查询卡类别   Varchar2(20);
  n_查询卡类别id 病人医疗卡信息.卡类别id%Type;
  n_卡类别id     医疗卡类别.Id%Type;
  v_No           病人挂号记录.No%Type;
  v_付款方式     病人挂号记录.医疗付款方式%Type;
  d_挂号时间     病人挂号记录.登记时间%Type;
  d_就诊时间     病人挂号记录.执行时间%Type;
  v_就诊科室     部门表.名称%Type;
  v_就诊医生     病人挂号记录.执行人%Type;
  v_过敏史       病人过敏记录.药物名%Type;
  v_Temp         Varchar2(32767); --临时XML 
  x_Templet      Xmltype; --模板XML 
  v_Err_Msg      Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(a), 'IN/SFZH'), Extractvalue(Value(a), 'IN/YLKLB'), Extractvalue(Value(a), 'IN/YLKH'),
         Extractvalue(Value(a), 'IN/BRXM'), Extractvalue(Value(a), 'IN/MZH'), Extractvalue(Value(a), 'IN/GHDH'),
         Extractvalue(Value(a), 'IN/BRID'), Extractvalue(Value(a), 'IN/CXKLB')
  Into v_身份证号, v_医疗卡, v_卡号, v_姓名, v_门诊号, v_挂号单, v_病人id, v_查询卡类别
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) a;

  If v_查询卡类别 Is Not Null Then
    Select Max(Id) Into n_查询卡类别id From 医疗卡类别 Where 名称 = v_查询卡类别;
    If n_查询卡类别id Is Null Then
      n_查询卡类别id := To_Number(v_查询卡类别);
    End If;
  End If;

  If v_病人id Is Null Then
  
    If v_身份证号 Is Null And v_医疗卡 Is Null And v_卡号 Is Null And v_姓名 Is Null And v_门诊号 Is Null And v_挂号单 Is Null Then
      v_Err_Msg := '未传入任何条件,无法完成查询!';
      Raise Err_Item;
    End If;
  
    If v_医疗卡 Is Not Null Then
      Select Max(Id) Into n_卡类别id From 医疗卡类别 Where 名称 = v_医疗卡;
      If n_卡类别id Is Null Then
        n_卡类别id := To_Number(v_医疗卡);
      End If;
    End If;
  
    If v_挂号单 Is Null Then
      If Nvl(n_卡类别id, 0) = 0 Then
        If Nvl(v_门诊号, 0) <> 0 Then
          Select 病人id Into v_病人id From 病人信息 Where 门诊号 = v_门诊号;
        Else
          Select Max(病人id)
          Into v_病人id
          From 病人信息
          Where Nvl(身份证号, '-') = Nvl(v_身份证号, Nvl(身份证号, '-')) And 姓名 = Nvl(v_姓名, 姓名);
        End If;
      Else
        Select Max(病人id) Into v_病人id From 病人医疗卡信息 Where 卡类别id = n_卡类别id And 卡号 = v_卡号;
      End If;
    Else
      Select Max(病人id) Into v_病人id From 病人挂号记录 Where No = v_挂号单 And 记录性质 = 1 And 记录状态 In (1, 3);
    End If;
  End If;
  If Nvl(v_病人id, 0) = 0 Then
    v_Err_Msg := '根据传入条件,无法完成查询!';
    Raise Err_Item;
  End If;
  For r_挂号 In (Select c.病人id, c.当前科室id, c.门诊号, c.姓名, c.性别, c.年龄, c.婚姻状况, c.国籍, c.出生日期, c.身份证号, c.职业, c.学历, c.民族, c.家庭电话,
                      c.家庭地址, c.户口地址, c.身份, c.手机号, c.联系人姓名, c.联系人电话, c.联系人地址, c.出生地点, Max(f.卡号) As 卡号
               From 病人信息 c, 病人医疗卡信息 f
               Where c.病人id = v_病人id And c.病人id = f.病人id(+) And f.卡类别id(+) = n_查询卡类别id And Nvl(f.状态, 0) = 0
               Group By c.病人id, c.当前科室id, c.门诊号, c.姓名, c.性别, c.出生日期, c.身份证号, c.职业, c.学历, c.民族, c.家庭电话, c.家庭地址, c.户口地址,
                        c.身份, c.手机号, c.联系人姓名, c.联系人电话, c.联系人地址, c.出生地点, c.年龄, c.婚姻状况, c.国籍) Loop
    v_Temp := '<BR>';
  
    If v_挂号单 Is Null Then
      Select Max(No), Max(医疗付款方式), Max(登记时间), Max(执行时间), Max(执行人), Max(就诊科室)
      Into v_No, v_付款方式, d_挂号时间, d_就诊时间, v_就诊医生, v_就诊科室
      From (Select a.No, a.医疗付款方式, a.登记时间, a.执行时间, a.执行人, b.名称 As 就诊科室
             From 病人挂号记录 a, 部门表 b
             Where a.执行部门id = b.Id(+) And a.病人id = r_挂号.病人id And a.记录性质 = 1 And a.记录状态 = 1
             Order By a.登记时间 Desc)
      Where Rownum < 2;
    Else
      Select Max(a.No), Max(a.医疗付款方式), Max(a.登记时间), Max(a.执行时间), Max(a.执行人), Max(b.名称) As 就诊科室
      Into v_No, v_付款方式, d_挂号时间, d_就诊时间, v_就诊医生, v_就诊科室
      From 病人挂号记录 a, 部门表 b
      Where a.执行部门id = b.Id(+) And a.No = v_挂号单 And a.记录性质 = 1 And a.记录状态 In (1, 3);
    End If;
  
    For r In (Select 药物名 From 病人过敏记录 Where 病人id = r_挂号.病人id) Loop
      v_过敏史 := v_过敏史 || ',' || r.药物名;
    End Loop;
    v_过敏史 := Substr(v_过敏史, 2);
  
    v_Temp := v_Temp || '<BRID>' || r_挂号.病人id || '</BRID>';
    v_Temp := v_Temp || '<XM>' || r_挂号.姓名 || '</XM>';
    v_Temp := v_Temp || '<XB>' || r_挂号.性别 || '</XB>';
    v_Temp := v_Temp || '<NL>' || r_挂号.年龄 || '</NL>';
    v_Temp := v_Temp || '<CSRQ>' || To_Char(r_挂号.出生日期, 'yyyy-mm-dd hh24:mi:ss') || '</CSRQ>';
    v_Temp := v_Temp || '<MZH>' || r_挂号.门诊号 || '</MZH>';
    v_Temp := v_Temp || '<HY>' || r_挂号.婚姻状况 || '</HY>';
    v_Temp := v_Temp || '<GJ>' || r_挂号.国籍 || '</GJ>';
    v_Temp := v_Temp || '<MZ>' || r_挂号.民族 || '</MZ>';
    v_Temp := v_Temp || '<XL>' || r_挂号.学历 || '</XL>';
    v_Temp := v_Temp || '<SF>' || r_挂号.身份 || '</SF>';
    v_Temp := v_Temp || '<ZY>' || r_挂号.职业 || '</ZY>';
    v_Temp := v_Temp || '<SFZH>' || r_挂号.身份证号 || '</SFZH>';
    v_Temp := v_Temp || '<FKFS>' || v_付款方式 || '</FKFS>';
    v_Temp := v_Temp || '<LXFS>' || r_挂号.手机号 || '</LXFS>';
    v_Temp := v_Temp || '<LXRXM>' || r_挂号.联系人姓名 || '</LXRXM>';
    v_Temp := v_Temp || '<LXRDH>' || r_挂号.联系人电话 || '</LXRDH>';
    v_Temp := v_Temp || '<LXRDZ>' || r_挂号.联系人地址 || '</LXRDZ>';
    v_Temp := v_Temp || '<LXDH>' || r_挂号.家庭电话 || '</LXDH>';
    v_Temp := v_Temp || '<XJZDZ>' || r_挂号.家庭地址 || '</XJZDZ>';
    v_Temp := v_Temp || '<HJDZ>' || r_挂号.户口地址 || '</HJDZ>';
    v_Temp := v_Temp || '<CSDD>' || r_挂号.出生地点 || '</CSDD>';
    v_Temp := v_Temp || '<KSID>' || r_挂号.当前科室id || '</KSID>';
    v_Temp := v_Temp || '<CXKH>' || r_挂号.卡号 || '</CXKH>';
    v_Temp := v_Temp || '<GMS>' || v_过敏史 || '</GMS>';
    v_Temp := v_Temp || '<GHD>' || v_No || '</GHD>';
    v_Temp := v_Temp || '<GHSJ>' || To_Char(d_挂号时间, 'yyyy-mm-dd hh24:mi:ss') || '</GHSJ>';
    v_Temp := v_Temp || '<JZSJ>' || To_Char(d_就诊时间, 'yyyy-mm-dd hh24:mi:ss') || '</JZSJ>';
    v_Temp := v_Temp || '<JZKS>' || v_就诊科室 || '</JZKS>';
    v_Temp := v_Temp || '<JZYS>' || v_就诊医生 || '</JZYS>';
    v_Temp := v_Temp || '</BR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;

  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Third_Getpatiinfo;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0023' Where 编号=&n_System;
Commit;
