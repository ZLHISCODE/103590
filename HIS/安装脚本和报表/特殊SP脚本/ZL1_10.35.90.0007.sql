----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--124650:李业庆,2018-04-20,材料特性增加是否分零属性
Alter Table 材料特性 Add 是否分零 Number(1) Default 0;

------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------

--123971:刘鹏飞,2018-04-16,血液接收后才允许执行登记
Insert Into Zlparameters
  (Id, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
Values
  (Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 0, 0, 301, '血液接收后才允许执行登记', 1, 1,
   '启用血库系统时医护人员取血回室后，是否需要进行血液接收核对环节才允许进行输血执行情况登记', '0-无需进行接收环节即可进行执行情况登记,1-必须进行血液接收核对环节才允许进行执行情况登记',
   '只有启用236号参数[启用血库管理系统]，才允许设置此参数，以及根据此参数控制是否需要进行血液接收环节', '医院可根据具体业务管理模式进行设置', Null);


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--122998:胡俊勇,2018-04-20,集成平台消息锚点
Create Or Replace Procedure Zl_字典管理_Execute(Sql_In In Varchar2) Is
  --一条完整的SQL语句，注意对象前一定要把所有者加上。
  --如UPDATE ZLHIS.结算方式 SET 缺省标志=0
  v_Rulesql Varchar2(8000);
  n_Pos     Number;
  v_Tmp     Varchar2(4000);
  v_Tab     Varchar2(100);
  v_Sql     Varchar2(8000);
  n_Count   Number;
  v_Owner   Varchar2(100);
  v_Code    Varchar2(100);
  v_Tmp1    Varchar2(8000);

  v_Err Varchar2(500);
  Err_Custom Exception;
Begin
  -------------------------
  --SQL校验
  ----------------------
  --1.格式化SQL语句
  v_Rulesql := Upper(Sql_In);
  v_Rulesql := Trim(Replace(v_Rulesql, Chr(10), ' '));
  v_Rulesql := Trim(Replace(v_Rulesql, Chr(13), ' '));
  --将双空格替换为单空格
  While Instr(v_Rulesql, '  ', 1) > 0 Loop
    v_Rulesql := Trim(Replace(v_Rulesql, '  ', ''));
  End Loop;
  v_Rulesql := Trim(v_Rulesql);
  --2、必须是标准的Insert,uPdate,Delete语句
  n_Pos := Instr(v_Rulesql, ' ');
  --三种标准的DML语句一定包含空格，并且空格的位置是第七位
  If n_Pos = 0 Or n_Pos <> 7 Then
    v_Err := '语法检查失败！语法错误或语句不是DML语句！';
    Raise Err_Custom;
  End If;
  v_Tmp := Trim(Substr(v_Rulesql, 1, n_Pos));
  v_Sql := Trim(Substr(v_Rulesql, n_Pos));

  If v_Tmp = 'INSERT' Or v_Tmp = 'DELETE' Or v_Tmp = 'UPDATE' Then
    --Insert 语句必须是Insert into tableName(col1,col2,...) values(val1,val2,...)
    If v_Tmp = 'INSERT' Then
      --Insert 语句是Insert into tableName(col1,col2,...) values(val1,val2,...)
      If v_Rulesql Like 'INSERT INTO %(%)%VALUES%(%)' Or v_Rulesql Like 'INSERT INTO %(%)%SELECT % FROM DUAL' Then
        --截取INTO TableName 字段
        n_Pos := Instr(v_Sql, '(');
        v_Tab := Trim(Substr(v_Sql, 1, n_Pos - 1));
        --截取OWNER.Table字段
        n_Pos := Instr(v_Tab, ' ');
        v_Tab := Trim(Substr(v_Tab, n_Pos));
      Else
        v_Err := '语法检查失败！Insert语句语法错误。';
        Raise Err_Custom;
      End If;
    Elsif v_Tmp = 'UPDATE' Then
      --Update 语句必须是Update tableName Set COl1=val1,.....
      If v_Rulesql Like 'UPDATE % SET %' Then
        --截取OWNER.Table字段
        n_Pos := Instr(v_Sql, 'SET');
        v_Tab := Trim(Substr(v_Sql, 1, n_Pos - 1));
      Else
        v_Err := '语法检查失败！UPDATE语句语法错误。';
        Raise Err_Custom;
      End If;
    Elsif v_Tmp = 'DELETE' Then
      --DELETE 语句必须是DELETE [From] tableName ,DELETE [From] tableName Where ..........
      If v_Rulesql Like 'DELETE % WHERE %' Then
        --delete语句含FROM
        If v_Rulesql Like 'DELETE FROM % WHERE %' Then
          --截取FROM TableName 字段
          n_Pos := Instr(v_Sql, 'WHERE');
          v_Tab := Trim(Substr(v_Sql, 1, n_Pos - 1));
          --截取OWNER.Table字段
          n_Pos := Instr(v_Tab, ' ');
          v_Tab := Trim(Substr(v_Tab, n_Pos));
          --delete语句不含FROM
        Else
          --截取OWNER.Table字段
          n_Pos := Instr(v_Sql, 'WHERE');
          v_Tab := Trim(Substr(v_Sql, 1, n_Pos - 1));
        End If;
      Elsif v_Rulesql Like 'DELETE % ' Then
        --delete语句含FROM
        If v_Rulesql Like 'DELETE FROM %' Then
          --截取OWNER.Table字段
          n_Pos := Instr(v_Tab, ' ');
          v_Tab := Trim(Substr(v_Sql, n_Pos));
          --delete语句不含FROM
        Else
          --截取OWNER.Table字段
          v_Tab := v_Sql;
        End If;
      Else
        v_Err := '语法检查失败！DELETE语句语法错误。';
        Raise Err_Custom;
      End If;
    End If;
  Else
    v_Err := '语法检查失败！语句必须是DML语句。';
    Raise Err_Custom;
  End If;
  --获取所有者以及系统号
  --没有带所有者时默认为标准版
  v_Tab := Trim(v_Tab);
  If v_Tab || ' ' <> ' ' Then
    n_Pos := Instr(v_Tab, '.');
    If n_Pos <> 0 Then
      v_Owner := Substr(v_Tab, 1, n_Pos - 1);
      v_Tab   := Substr(v_Tab, n_Pos + 1);
    Else
      Select Max(a.所有者) Into v_Owner From zlSystems A Where a.编号 = 100;
    End If;
  End If;

  --DML语句操作的表必须是ZLBASECODE中的非固定表
  Select Count(1)
  Into n_Count
  From zlBaseCode
  Where 固定 = 0 And 表名 = v_Tab And 系统 In (Select a.编号 From zlSystems A Where a.所有者 = v_Owner);

  If n_Count = 0 Then
    v_Err := '表' || v_Tab || '不是当前系统所有的非固定表。';
    Raise Err_Custom;
  End If;

  If v_Tab = '诊疗检查类型' Then
    --解析编码值
    If v_Tmp = 'INSERT' Then
      n_Pos  := Instr(v_Sql, 'VALUES');
      v_Tmp1 := Substr(v_Sql, n_Pos);
      n_Pos  := Instr(v_Tmp1, ',');
      v_Tmp1 := Substr(v_Tmp1, 1, n_Pos - 1);
      n_Pos  := Instr(v_Tmp1, '(');
      v_Tmp1 := Substr(v_Tmp1, n_Pos + 1);
      v_Code := Trim(Replace(v_Tmp1, '''', ''));
    Else
      n_Pos  := Instr(v_Sql, 'WHERE');
      v_Tmp1 := Substr(v_Sql, n_Pos);
      n_Pos  := Instr(v_Tmp1, '=');
      v_Tmp1 := Substr(v_Tmp1, n_Pos + 1);
      v_Code := Trim(Replace(v_Tmp1, '''', ''));
    End If;
  End If;

  --因为可能独立安装，所以只能动态执行  
  If v_Tmp = 'DELETE' Then
    If v_Tab = '诊疗检查类型' Then
      --删除记录
      For R In (Select a.编码, a.名称, a.简码, a.建病案 From 诊疗检查类型 A Where a.编码 = v_Code) Loop
        --b_Message.Zlhis_Dictpacs_003(r.编码, r.名称, r.简码, r.建病案);      
        Begin
          Execute Immediate 'call b_Message.Zlhis_Dictpacs_003(:1,:2,:3,:4)'
            Using r.编码, r.名称, r.简码, r.建病案;
        Exception
          When Others Then
            Null;
        End;
      End Loop;
    Elsif v_Tab = '诊疗检验标本' Then
      --删除记录
      For R In (Select a.编码, a.名称, a.简码, a.适用性别 From 诊疗检验标本 A Where a.编码 = v_Code) Loop
        --b_Message.Zlhis_Dictlis_006(r.编码, r.名称, r.简码, r.适用性别);          
        Begin
          Execute Immediate 'call b_Message.Zlhis_Dictlis_006(:1,:2,:3,:4)'
            Using r.编码, r.名称, r.简码, r.适用性别;
        Exception
          When Others Then
            Null;
        End;
      End Loop;
    End If;
  End If;

  Execute Immediate v_Rulesql;

  If v_Tab = '诊疗检查类型' Then
    If v_Tmp = 'INSERT' Or v_Tmp = 'UPDATE' Then
      For R In (Select a.编码, a.名称, a.简码, a.建病案 From 诊疗检查类型 A Where a.编码 = v_Code) Loop
        If v_Tmp = 'INSERT' Then
          --b_Message.Zlhis_Dictpacs_001(r.编码, r.名称, r.简码, r.建病案);         
          Begin
            Execute Immediate 'call b_Message.Zlhis_Dictpacs_001(:1,:2,:3,:4)'
              Using r.编码, r.名称, r.简码, r.建病案;
          Exception
            When Others Then
              Null;
          End;
        Else
          --b_Message.Zlhis_Dictpacs_002(r.编码, r.名称, r.简码, r.建病案);          
          Begin
            Execute Immediate 'call b_Message.Zlhis_Dictpacs_002(:1,:2,:3,:4)'
              Using r.编码, r.名称, r.简码, r.建病案;
          Exception
            When Others Then
              Null;
          End;
        End If;
      End Loop;
    End If;
  Elsif v_Tab = '诊疗检验标本' Then
    If v_Tmp = 'INSERT' Or v_Tmp = 'UPDATE' Then
      For R In (Select a.编码, a.名称, a.简码, a.适用性别 From 诊疗检验标本 A Where a.编码 = v_Code) Loop
        If v_Tmp = 'INSERT' Then
          --b_Message.Zlhis_Dictlis_004(r.编码, r.名称, r.简码, r.适用性别);        
          Begin
            Execute Immediate 'call b_Message.Zlhis_Dictlis_004(:1,:2,:3,:4)'
              Using r.编码, r.名称, r.简码, r.适用性别;
          Exception
            When Others Then
              Null;
          End;
        Else
          --b_Message.Zlhis_Dictlis_005(r.编码, r.名称, r.简码, r.适用性别);        
          Begin
            Execute Immediate 'call b_Message.Zlhis_Dictlis_005(:1,:2,:3,:4)'
              Using r.编码, r.名称, r.简码, r.适用性别;
          Exception
            When Others Then
              Null;
          End;
        End If;
      End Loop;
    End If;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_字典管理_Execute;
/

--124359:胡俊勇,2018-04-17,检查项目医嘱按部位执行或取消执行
Create Or Replace Procedure Zl_住院医嘱执行_Finish
(
  医嘱id_In     In 病人医嘱执行.医嘱id%Type,
  发送号_In     In 病人医嘱执行.发送号%Type,
  单独执行_In   In Number,
  操作员编号_In In 人员表.编号%Type,
  操作员姓名_In In 人员表.姓名%Type,
  组id_In       In 病人医嘱执行.医嘱id%Type,
  诊疗类别_In   In 病人医嘱记录.诊疗类别%Type,
  执行部门id_In In 住院费用记录.执行部门id%Type
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --      单独执行_In=检验医嘱组合是否采用对每个项目分散单独执行的方式
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门
) Is
  --医嘱相关的费用单据
  Cursor c_No Is
    Select a.No || ':' || a.记录性质
    From 病人医嘱附费 A, 病人医嘱记录 B
    Where a.发送号 + 0 = 发送号_In And a.医嘱id = b.Id And b.诊疗类别 = 诊疗类别_In And (b.Id = 组id_In Or b.相关id = 组id_In)
    Union
    Select NO || ':' || 记录性质
    From 病人医嘱发送
    Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;

  Cursor c_Noone Is
    Select NO || ':' || 记录性质
    From 病人医嘱附费
    Where 发送号 + 0 = 发送号_In And 医嘱id = 医嘱id_In
    Union
    Select NO || ':' || 记录性质
    From 病人医嘱发送
    Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;

  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;

  Cursor c_Finish(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select Distinct a.Id, a.收费类别, a.收费细目id
           From 住院费用记录 A, 病人医嘱记录 B,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
                  From Table(r_No)) N
           Where a.医嘱序号 = b.Id And (b.Id = 组id_In Or b.相关id = 组id_In) And a.No = n.No And a.记录性质 = n.记录性质 And
                 a.记录状态 In (0, 1, 3) And a.执行状态 <> 1 And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)) A, 材料特性 B
    Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);

  Cursor c_Finishone(r_No t_Strlist) Is
    Select /*+ RULE */
     a.Id
    From (Select a.Id, a.收费类别, a.收费细目id
           From 住院费用记录 A,
                (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                         To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
                  From Table(r_No)) N
           Where a.医嘱序号 = 医嘱id_In And a.No = n.No And a.记录性质 = n.记录性质 And a.记录状态 In (0, 1, 3) And a.执行状态 <> 1 And
                 (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)) A, 材料特性 B
    Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);

  --执行中包含跟踪在用的未发卫料时，根据参数设置是否自动发料
  --卫生材料医嘱目前不存在单独和组合执行的情况
  Cursor c_Stuff(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id, Decode(d.高值材料, 1, a.执行部门id, b.库房id) As 库房id
    From 住院费用记录 A, 药品收发记录 B, 病人医嘱记录 C, 材料特性 D,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
           From Table(r_No)) N
    Where d.材料id = a.收费细目id And a.Id = b.费用id And Mod(b.记录状态, 3) = 1 And b.审核人 Is Null And b.库房id Is Not Null And
          a.收费类别 = '4' And a.记录状态 = 1 And a.医嘱序号 = c.Id And c.诊疗类别 = 诊疗类别_In And (c.Id = 组id_In Or c.相关id = 组id_In) And
          a.No = n.No And a.记录性质 = n.记录性质 And b.单据 In (25, 26) And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And
          (c.Id = 医嘱id_In And 单独执行_In = 1 Or Nvl(单独执行_In, 0) <> 1)
    Order By b.库房id, b.药品id;

  --执行中包含的未发药品，本科执行的自动发药
  Cursor c_Drug(r_No t_Strlist) Is
    Select /*+ RULE */
     b.Id, b.库房id
    From 住院费用记录 A, 药品收发记录 B, 病人医嘱记录 C, 病案主页 D,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
           From Table(r_No)) N
    Where a.Id = b.费用id And Mod(b.记录状态, 3) = 1 And b.审核人 Is Null And b.库房id = d.当前病区id And a.收费类别 In ('5', '6', '7') And
          a.记录状态 = 1 And a.医嘱序号 = c.Id And c.诊疗类别 = 诊疗类别_In And c.病人id = d.病人id And c.主页id = d.主页id And
          (c.Id = 组id_In Or c.相关id = 组id_In) And a.No = n.No And a.记录性质 = n.记录性质 And
          (c.Id = 医嘱id_In And 单独执行_In = 1 Or Nvl(单独执行_In, 0) <> 1)
    Order By b.库房id, b.药品id;

  --未审核的费用行(包含药品和卫材)
  Cursor c_Verify(r_No t_Strlist) Is
    Select /*+ RULE */
    Distinct a.No, a.序号
    From 住院费用记录 A, 病人医嘱记录 C,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
           From Table(r_No)) N
    Where a.记帐费用 = 1 And a.记录状态 = 0 And a.价格父号 Is Null And a.医嘱序号 = c.Id And (c.Id = 组id_In Or c.相关id = 组id_In) And
          a.No = n.No And a.记录性质 = n.记录性质 And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)
    Order By NO, 序号;

  Cursor c_Verifyone(r_No t_Strlist) Is
    Select /*+ RULE */
     a.No, a.序号
    From 住院费用记录 A,
         (Select Substr(Column_Value, 1, Instr(Column_Value, ':') - 1) As NO,
                  To_Number(Substr(Column_Value, Instr(Column_Value, ':') + 1)) As 记录性质
           From Table(r_No)) N
    Where a.记帐费用 = 1 And a.记录状态 = 0 And a.价格父号 Is Null And a.医嘱序号 + 0 = 医嘱id_In And a.No = n.No And a.记录性质 = n.记录性质 And
          (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)
    Order By NO, 序号;

  v_No   病人医嘱发送.No%Type;
  v_序号 Varchar2(1000);

  v_发料号  药品收发记录.汇总发药号%Type;
  v_库房id  药品收发记录.库房id%Type;
  v_收发ids Varchar2(4000);

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
  Forall I In 1 .. r_Finish.Count
    Update 住院费用记录 Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 操作员姓名_In Where ID = r_Finish(I);

  --执行时自动审核对应的记帐划价单费用
  --包含医嘱对应的药品及卫材费用，因为医嘱已执行，费用应该生效。
  If Nvl(单独执行_In, 0) = 0 Then
    For r_Verify In c_Verify(r_No) Loop
      If r_Verify.No <> v_No And v_序号 Is Not Null Then
        Zl_住院记帐记录_Verify(v_No, 操作员编号_In, 操作员姓名_In, Substr(v_序号, 2));
        v_序号 := Null;
      End If;
      v_No   := r_Verify.No;
      v_序号 := v_序号 || ',' || r_Verify.序号;
    End Loop;
  Else
    For r_Verify In c_Verifyone(r_No) Loop
      If r_Verify.No <> v_No And v_序号 Is Not Null Then
        Zl_住院记帐记录_Verify(v_No, 操作员编号_In, 操作员姓名_In, Substr(v_序号, 2));
        v_序号 := Null;
      End If;
      v_No   := r_Verify.No;
      v_序号 := v_序号 || ',' || r_Verify.序号;
    End Loop;
  End If;
  If v_序号 Is Not Null Then
    Zl_住院记帐记录_Verify(v_No, 操作员编号_In, 操作员姓名_In, Substr(v_序号, 2));
  End If;

  --处理跟踪在用卫材自动发料
  For r_Stuff In c_Stuff(r_No_Stuff) Loop
    If v_发料号 Is Null Then
      v_发料号 := Nextno(20);
    End If;
  
    If r_Stuff.库房id <> Nvl(v_库房id, 0) Then
      If Nvl(v_库房id, 0) <> 0 And v_收发ids Is Not Null Then
        v_收发ids := Substr(v_收发ids, 2);
        Zl_药品收发记录_批量发料(v_收发ids, v_库房id, 操作员姓名_In, Sysdate, 1, 操作员姓名_In, v_发料号, 操作员姓名_In);
      End If;
      v_库房id  := r_Stuff.库房id;
      v_收发ids := Null;
    End If;
  
    v_收发ids := v_收发ids || '|' || r_Stuff.Id || ',0';
  End Loop;
  If Nvl(v_库房id, 0) <> 0 And v_收发ids Is Not Null Then
    v_收发ids := Substr(v_收发ids, 2);
    Zl_药品收发记录_批量发料(v_收发ids, v_库房id, 操作员姓名_In, Sysdate, 1, 操作员姓名_In, v_发料号, 操作员姓名_In);
  End If;

  --处理药品自动发药(只在护士站，本科药品才处理,本科由参数和游标判断)
  Select 医嘱期效 Into v_医嘱期效 From 病人医嘱记录 Where ID = 医嘱id_In;
  If Substr(zl_GetSysParameter('本科执行自动完成', 1254), v_医嘱期效 + 1, 1) = '1' Then
    v_发料号  := Null;
    v_收发ids := Null;
    For r_Drug In c_Drug(r_No_Stuff) Loop
      If v_发料号 Is Null Then
        v_发料号 := Nextno(20);
      End If;
    
      If r_Drug.库房id <> Nvl(v_库房id, 0) Then
        If Nvl(v_库房id, 0) <> 0 And v_收发ids Is Not Null Then
          v_收发ids := Substr(v_收发ids, 2);
          Zl_药品收发记录_批量发药(v_收发ids, v_库房id, 操作员姓名_In, Sysdate, 1, 操作员姓名_In, v_发料号);
        End If;
        v_库房id  := r_Drug.库房id;
        v_收发ids := Null;
      End If;
    
      v_收发ids := v_收发ids || '|' || r_Drug.Id || ',0';
    End Loop;
    If Nvl(v_库房id, 0) <> 0 And v_收发ids Is Not Null Then
      v_收发ids := Substr(v_收发ids, 2);
      Zl_药品收发记录_批量发药(v_收发ids, v_库房id, 操作员姓名_In, Sysdate, 1, 操作员姓名_In, v_发料号);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_住院医嘱执行_Finish;
/

--124359:胡俊勇,2018-04-17,检查项目医嘱按部位执行或取消执行
Create Or Replace Procedure Zl_住院医嘱执行_Cancel
(
  医嘱id_In     In 病人医嘱执行.医嘱id%Type,
  发送号_In     In 病人医嘱执行.发送号%Type,
  单独执行_In   In Number,
  操作员编号_In In 人员表.编号%Type,
  操作员姓名_In In 人员表.姓名%Type,
  组id_In       In 病人医嘱执行.医嘱id%Type,
  诊疗类别_In   In 病人医嘱记录.诊疗类别%Type,
  执行部门id_In In 门诊费用记录.执行部门id%Type
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --      单独执行_In=检验医嘱组合是否采用对每个项目分散单独执行的方式
) Is
  --医嘱相关的费用单据
  Cursor c_No Is
    Select a.No || ':' || a.记录性质
    From 病人医嘱附费 A, 病人医嘱记录 B
    Where a.发送号 + 0 = 发送号_In And a.医嘱id = b.Id And b.诊疗类别 = 诊疗类别_In And (b.Id = 组id_In Or b.相关id = 组id_In)
    Union
    Select NO || ':' || 记录性质
    From 病人医嘱发送
    Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;

  Cursor c_Noone Is
    Select NO || ':' || 记录性质
    From 病人医嘱附费
    Where 发送号 + 0 = 发送号_In And 医嘱id = 医嘱id_In
    Union
    Select NO || ':' || 记录性质
    From 病人医嘱发送
    Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;
  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;

  n_执行次数 Number;
  n_剩余次数 Number;
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
          b.单据 In (25, 26) And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And
          (c.Id = 医嘱id_In And 单独执行_In = 1 Or Nvl(单独执行_In, 0) <> 1)
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
          c.主页id = d.主页id And (c.Id = 组id_In Or c.相关id = 组id_In) And a.No = n.No And a.记录性质 = n.记录性质 And
          (c.Id = 医嘱id_In And 单独执行_In = 1 Or Nvl(单独执行_In, 0) <> 1)
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
    Set 执行状态 = n_执行状态, 执行时间 = Decode(n_执行状态, 0, Null, 执行时间), 执行人 = Decode(n_执行状态, 0, Null, 执行人)
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

  If Substr(zl_GetSysParameter('本科执行自动完成', 1254), v_医嘱期效 + 1, 1) = '1' And n_Count = 1 Then
    For r_Drug In c_Drug(r_No_Stuff) Loop
      Zl_药品收发记录_部门退药(r_Drug.Id, 操作员姓名_In, Sysdate);
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_住院医嘱执行_Cancel;
/

--124359:胡俊勇,2018-04-17,检查项目医嘱按部位执行或取消执行
Create Or Replace Procedure Zl_门诊医嘱执行_Finish
(
  医嘱id_In     In 病人医嘱执行.医嘱id%Type,
  发送号_In     In 病人医嘱执行.发送号%Type,
  单独执行_In   In Number,
  操作员编号_In In 人员表.编号%Type,
  操作员姓名_In In 人员表.姓名%Type,
  组id_In       In 病人医嘱执行.医嘱id%Type,
  诊疗类别_In   In 病人医嘱记录.诊疗类别%Type,
  执行部门id_In In 门诊费用记录.执行部门id%Type
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --      单独执行_In=检验医嘱组合是否采用对每个项目分散单独执行的方式
  --执行部门id_In=仅处理指定执行部门的费用，不传或传入0时不限制执行部门
) Is
  --医嘱相关的费用单据
  Cursor c_No Is
    Select a.No || ':' || a.记录性质
    From 病人医嘱附费 A, 病人医嘱记录 B
    Where a.发送号 + 0 = 发送号_In And a.医嘱id = b.Id And b.诊疗类别 = 诊疗类别_In And (b.Id = 组id_In Or b.相关id = 组id_In)
    Union
    Select NO || ':' || 记录性质
    From 病人医嘱发送
    Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;

  Cursor c_Noone Is
    Select NO || ':' || 记录性质
    From 病人医嘱附费
    Where 发送号 + 0 = 发送号_In And 医嘱id = 医嘱id_In
    Union
    Select NO || ':' || 记录性质
    From 病人医嘱发送
    Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;

  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;
  n_Cnt      Number;
  v_Error    Varchar2(2000);
  Err_Custom Exception;
  v_执行前先结算 Varchar2(500);

  Cursor c_Finish(r_No t_Strlist) Is
    Select a.Id
    From (Select Distinct a.Id, a.收费类别, a.收费细目id
           From 门诊费用记录 A, 病人医嘱记录 B,
                (Select /*+cardinality(f,10)*/
                   Substr(f.Column_Value, 1, Instr(f.Column_Value, ':') - 1) As NO,
                   To_Number(Substr(f.Column_Value, Instr(f.Column_Value, ':') + 1)) As 记录性质
                  From Table(r_No) F) N
           Where a.医嘱序号 = b.Id And (b.Id = 组id_In Or b.相关id = 组id_In) And a.No = n.No And Mod(a.记录性质, 10) = n.记录性质 And
                 a.记录状态 In (0, 1, 3) And a.执行状态 <> 1 And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)) A, 材料特性 B
    Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);

  Cursor c_Finishone(r_No t_Strlist) Is
    Select a.Id
    From (Select a.Id, a.收费类别, a.收费细目id
           From 门诊费用记录 A,
                (Select /*+cardinality(f,10)*/
                   Substr(f.Column_Value, 1, Instr(f.Column_Value, ':') - 1) As NO,
                   To_Number(Substr(f.Column_Value, Instr(f.Column_Value, ':') + 1)) As 记录性质
                  From Table(r_No) F) N
           Where a.医嘱序号 = 医嘱id_In And a.No = n.No And Mod(a.记录性质, 10) = n.记录性质 And a.记录状态 In (0, 1, 3) And a.执行状态 <> 1 And
                 (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)) A, 材料特性 B
    Where a.收费细目id = b.材料id(+) And Not a.收费类别 In ('5', '6', '7') And Not (a.收费类别 = '4' And Nvl(b.跟踪在用, 0) = 1);

  --执行中包含跟踪在用的未发卫料时，根据参数设置是否自动发料
  --卫生材料医嘱目前不存在单独和组合执行的情况
  Cursor c_Stuff(r_No t_Strlist) Is
    Select b.Id, Decode(d.高值材料, 1, a.执行部门id, b.库房id) As 库房id
    From 门诊费用记录 A, 药品收发记录 B, 病人医嘱记录 C, 材料特性 D,
         (Select /*+cardinality(f,10)*/
            Substr(f.Column_Value, 1, Instr(f.Column_Value, ':') - 1) As NO,
            To_Number(Substr(f.Column_Value, Instr(f.Column_Value, ':') + 1)) As 记录性质
           From Table(r_No) F) N
    Where d.材料id = a.收费细目id And a.Id = b.费用id And Mod(b.记录状态, 3) = 1 And b.审核人 Is Null And b.库房id Is Not Null And
          a.收费类别 = '4' And a.记录状态 = 1 And a.医嘱序号 = c.Id And c.诊疗类别 = 诊疗类别_In And (c.Id = 组id_In Or c.相关id = 组id_In) And
          a.No = n.No And Mod(a.记录性质, 10) = n.记录性质 And b.单据 In (24, 25, 26) And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And
          (c.Id = 医嘱id_In And 单独执行_In = 1 Or Nvl(单独执行_In, 0) <> 1)
    Order By b.库房id, b.药品id;

  --未审核的费用行(包含药品和卫材)
  Cursor c_Verify
  (
    r_No        t_Strlist,
    记帐费用_In Number := 1
  ) Is
    Select Distinct a.No, a.序号
    From 门诊费用记录 A, 病人医嘱记录 C,
         (Select /*+cardinality(f,10)*/
            Substr(f.Column_Value, 1, Instr(f.Column_Value, ':') - 1) As NO,
            To_Number(Substr(f.Column_Value, Instr(f.Column_Value, ':') + 1)) As 记录性质
           From Table(r_No) F) N
    Where Nvl(a.记帐费用, 0) = 记帐费用_In And a.记录状态 = 0 And a.价格父号 Is Null And a.医嘱序号 = c.Id And
          (c.Id = 组id_In Or c.相关id = 组id_In) And a.No = n.No And Mod(a.记录性质, 10) = n.记录性质 And
          (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)
    Order By NO, 序号;

  Cursor c_Verifyone
  (
    r_No        t_Strlist,
    记帐费用_In Number := 1
  ) Is
    Select a.No, a.序号
    From 门诊费用记录 A,
         (Select /*+cardinality(f,10)*/
            Substr(f.Column_Value, 1, Instr(f.Column_Value, ':') - 1) As NO,
            To_Number(Substr(f.Column_Value, Instr(f.Column_Value, ':') + 1)) As 记录性质
           From Table(r_No) F) N
    Where Nvl(a.记帐费用, 0) = 记帐费用_In And a.记录状态 = 0 And a.价格父号 Is Null And a.医嘱序号 + 0 = 医嘱id_In And a.No = n.No And
          Mod(a.记录性质, 10) = n.记录性质 And (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In)
    Order By NO, 序号;

  v_No   病人医嘱发送.No%Type;
  v_序号 Varchar2(1000);

  v_发料号  药品收发记录.汇总发药号%Type;
  v_库房id  药品收发记录.库房id%Type;
  v_收发ids Varchar2(4000);
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
  Select Count(1)
  Into n_Cnt
  From 门诊费用记录 A,
       (Select /*+cardinality(f,10)*/
          To_Number(f.Column_Value) As 费用id
         From Table(r_Finish) F) B
  Where a.Id = b.费用id And a.费用状态 = 1;

  If n_Cnt > 0 Then
    v_Error := '当前执行的医嘱对应的费用单据中存在异常单据。';
    Raise Err_Custom;
  End If;

  Select zl_GetSysParameter(163) Into v_执行前先结算 From Dual;
  Forall I In 1 .. r_Finish.Count
    Update 门诊费用记录 Set 执行状态 = 1, 执行时间 = Sysdate, 执行人 = 操作员姓名_In Where ID = r_Finish(I);

  --执行时自动审核对应的记帐划价单费用
  --包含医嘱对应的药品及卫材费用，因为医嘱已执行，费用应该生效。
  If Nvl(单独执行_In, 0) = 0 Then
    If Nvl(v_执行前先结算, '0') <> '0' Then
      For r_Verify In c_Verify(r_No, 0) Loop
        v_Error := '当前执行的医嘱还存在未收取的费用。';
        Raise Err_Custom;
      End Loop;
    End If;
    For r_Verify In c_Verify(r_No) Loop
      If r_Verify.No <> v_No And v_序号 Is Not Null Then
        Zl_门诊记帐记录_Verify(v_No, 操作员编号_In, 操作员姓名_In, Substr(v_序号, 2));
        v_序号 := Null;
      End If;
      v_No   := r_Verify.No;
      v_序号 := v_序号 || ',' || r_Verify.序号;
    End Loop;
  Else
    If Nvl(v_执行前先结算, '0') <> '0' Then
      For r_Verify In c_Verifyone(r_No, 0) Loop
        v_Error := '当前执行的医嘱还存在未收取的费用。';
        Raise Err_Custom;
      End Loop;
    End If;
    For r_Verify In c_Verifyone(r_No) Loop
      If r_Verify.No <> v_No And v_序号 Is Not Null Then
        Zl_门诊记帐记录_Verify(v_No, 操作员编号_In, 操作员姓名_In, Substr(v_序号, 2));
        v_序号 := Null;
      End If;
      v_No   := r_Verify.No;
      v_序号 := v_序号 || ',' || r_Verify.序号;
    End Loop;
  End If;
  If v_序号 Is Not Null Then
    Zl_门诊记帐记录_Verify(v_No, 操作员编号_In, 操作员姓名_In, Substr(v_序号, 2));
  End If;

  --处理跟踪在用卫材自动发料
  For r_Stuff In c_Stuff(r_No_Stuff) Loop
    If v_发料号 Is Null Then
      v_发料号 := Nextno(20);
    End If;
  
    If r_Stuff.库房id <> Nvl(v_库房id, 0) Then
      If Nvl(v_库房id, 0) <> 0 And v_收发ids Is Not Null Then
        v_收发ids := Substr(v_收发ids, 2);
        Zl_药品收发记录_批量发料(v_收发ids, v_库房id, 操作员姓名_In, Sysdate, 1, 操作员姓名_In, v_发料号, 操作员姓名_In);
      End If;
      v_库房id  := r_Stuff.库房id;
      v_收发ids := Null;
    End If;
  
    v_收发ids := v_收发ids || '|' || r_Stuff.Id || ',0';
  End Loop;
  If Nvl(v_库房id, 0) <> 0 And v_收发ids Is Not Null Then
    v_收发ids := Substr(v_收发ids, 2);
    Zl_药品收发记录_批量发料(v_收发ids, v_库房id, 操作员姓名_In, Sysdate, 1, 操作员姓名_In, v_发料号, 操作员姓名_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊医嘱执行_Finish;
/

--124359:胡俊勇,2018-04-17,检查项目医嘱按部位执行或取消执行
Create Or Replace Procedure Zl_门诊医嘱执行_Cancel
(
  医嘱id_In     In 病人医嘱执行.医嘱id%Type,
  发送号_In     In 病人医嘱执行.发送号%Type,
  单独执行_In   In Number,
  操作员编号_In In 人员表.编号%Type,
  操作员姓名_In In 人员表.姓名%Type,
  组id_In       In 病人医嘱执行.医嘱id%Type,
  诊疗类别_In   In 病人医嘱记录.诊疗类别%Type,
  执行部门id_In In 门诊费用记录.执行部门id%Type
  --参数：医嘱ID_IN=单独执行的医嘱ID，检验组合为显示的检验项目的ID。
  --      单独执行_In=检验医嘱组合是否采用对每个项目分散单独执行的方式
) Is
  --医嘱相关的费用单据
  Cursor c_No Is
    Select a.No || ':' || a.记录性质
    From 病人医嘱附费 A, 病人医嘱记录 B
    Where a.发送号 + 0 = 发送号_In And a.医嘱id = b.Id And b.诊疗类别 = 诊疗类别_In And (b.Id = 组id_In Or b.相关id = 组id_In)
    Union
    Select NO || ':' || 记录性质
    From 病人医嘱发送
    Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;

  Cursor c_Noone Is
    Select NO || ':' || 记录性质
    From 病人医嘱附费
    Where 发送号 + 0 = 发送号_In And 医嘱id = 医嘱id_In
    Union
    Select NO || ':' || 记录性质
    From 病人医嘱发送
    Where 医嘱id = 医嘱id_In And 发送号 + 0 = 发送号_In;

  r_No       t_Strlist;
  r_No_Stuff t_Strlist;
  r_Finish   t_Numlist;

  n_执行次数 Number;
  n_剩余次数 Number;
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
          (a.记录性质 = n.记录性质 Or a.记录性质 = 11 And n.记录性质 = 1) And b.单据 In (24, 25, 26) And
          (执行部门id_In = 0 Or a.执行部门id = 执行部门id_In) And (c.Id = 医嘱id_In And 单独执行_In = 1 Or Nvl(单独执行_In, 0) <> 1)
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
    Set 执行状态 = n_执行状态, 执行时间 = Decode(n_执行状态, 0, Null, 执行时间), 执行人 = Decode(n_执行状态, 0, Null, 执行人)
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


--124306:焦博,2018-04-17,调整Oracle过程Zl_Third_Charge_Del和Zl_Third_Registdel
Create Or Replace Procedure Zl_Third_Charge_Del
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  -------------------------------------------------------------------------------------------------- 
  --功能:三方退费交易 
  --入参:Xml_In: 
  --<IN> 
  --    <BRID>病人ID</BRID> 
  --    <XM>姓名</XM> 
  --    <SFZH>身份证号</SFZH> 
  --    <JE></JE> //退款总金额 
  --    <JSKLB></JSKLB>     //结算卡类别 
  --    <TFZY>退费摘要</TFZY> 
  --    <JCFP>1</JCFP>      //检查发票 
  --    <FYLIST> 
  --        <FY> 
  --           <DJH>退款单据号</DJH> 
  --           <XH>退款序号(格式:1,2,3..为空代表退剩余数量)</DJH> 
  --        <FY> 
  --    </FYLIST> 
  --    <TKLIST> 
  --        <TK> 
  --            <TKKLB>退款卡类别</TKKLB> 
  --            <TKKH>退款卡号</TKKH> 
  --            <TKFS>退款方式</TKFS> //退款方式:现金;支票,如果是三方卡,可以传空 
  --            <TKJE>支付金额</TKJE> 
  --            <JYLSH>交易流水号</JYLSH> 
  --            <TKZY>摘要</TKZY> 
  --            <TYJK>退回预交款</TYJK> //允冲预交时,只填JSJE节点:1-冲预交 
  --            <SFXFK>是否消费卡</SFXFK>   //(1-是消费卡),消费卡时,传入结算卡类别,结算卡号,结算金额等接点 
  --            <EXPENDLIST>  //扩展交易信息 
  --                <EXPEND> 
  --                    <JYMC>交易名称</JYMC> 
  --                    <JYLR>交易内容</JYLR> 
  --                </EXPEND> 
  --            </EXPENDLIST> 
  --        </TK> 
  --    </TKLIST> 
  --</IN> 

  --出参:Xml_Out 
  --  <OUT> 
  --    <CZSJ>操作时间</CZSJ>          //HIS的登记时间 
  --    <YJZID>原结帐ID</YJZID>       //原结帐ID 
  --    <CXID>冲销ID</CXID>          //冲销ID 
  --    ――如无下列错误结点则说明正确执行 
  --    <ERROR> 
  --      <MSG>错误信息</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  n_退款总额 门诊费用记录.实收金额%Type;
  n_卡类别id 医疗卡类别.Id%Type;
  v_结算方式 Varchar2(2000);

  n_病人id     门诊费用记录.病人id%Type;
  v_姓名       病人信息.姓名%Type;
  v_身份证号   病人信息.身份证号%Type;
  n_单据病人id 门诊费用记录.病人id%Type;
  v_操作员编码 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  n_冲销id     门诊费用记录.结帐id%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_结帐金额   门诊费用记录.结帐金额%Type;
  n_误差额     病人预交记录.冲预交%Type;
  n_原结算序号 病人预交记录.结算序号%Type;
  l_挂号单     t_Strlist := t_Strlist();
  n_Column     Number(18);
  n_检查发票   Number(3);
  n_是否打印   Number(3);
  v_结算卡类别 Varchar2(100);
  v_结帐ids    Varchar2(1000);

  n_消费卡id 消费卡信息.Id%Type;
  v_摘要     门诊费用记录.摘要%Type;
  n_Count    Number(18);

  d_退费时间 病人预交记录.收款时间%Type;

  v_退费结算 Varchar2(2000);
  v_普通结算 Varchar2(4000);
  n_Temp     Number(18);

  v_Temp    Varchar2(32767); --临时XML 
  x_Templet Xmltype; --模板XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  Procedure Third_Cardbalance_Modfiy
  (
    冲销id_In     病人预交记录.结帐id%Type,
    卡类别_In     Varchar2,
    卡号_In       病人预交记录.卡号%Type,
    退款金额_In   病人预交记录.冲预交%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    摘要_In       病人预交记录.摘要%Type,
    Xmlexpned_In  Xmltype
  ) Is
    n_卡类别id 医疗卡类别.Id%Type;
    v_结算方式 病人预交记录.结算方式%Type;
  
  Begin
    v_Err_Msg := Null;
    Begin
      n_卡类别id := To_Number(卡类别_In);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
    If n_卡类别id = 0 Then
      Begin
        Select ID, 结算方式, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_结算方式, v_Err_Msg
        From 医疗卡类别
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          n_卡类别id := -1;
          v_Err_Msg  := 卡类别_In || '不存在!';
      End;
    Else
      Begin
        Select ID, 结算方式, Decode(Nvl(是否启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_结算方式, v_Err_Msg
        From 医疗卡类别
        Where ID = n_卡类别id;
      Exception
        When Others Then
          n_卡类别id := -1;
          v_Err_Msg  := '未找到指定的结算支付信息!';
      End;
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  
    v_退费结算 := v_结算方式 || '|' || 退款金额_In || '|' || ' |' || Nvl(摘要_In, ' ');
    --   2.三方卡退费结算: 
    --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:结算方式|结算金额|结算号码|结算摘要 
    --     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入 
    --结算方式|结算金额|结算号码|结算摘要 
    Zl_门诊退费结算_Modify(2, n_病人id, 冲销id_In, v_退费结算, 0, n_卡类别id, 卡号_In, 交易流水号_In, 交易说明_In, 0, 0, 0, 0);
  
    --保存扩展结算信息 
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 0, 卡号_In, 冲销id_In, c_扩展.Jymc || '|' || c_扩展.Jylr, 0);
    End Loop;
  End Third_Cardbalance_Modfiy;

  Procedure Square_Cardbalance_Modfiy
  (
    冲销id_In     病人预交记录.结帐id%Type,
    卡类别_In     Varchar2,
    卡号_In       病人预交记录.卡号%Type,
    退款金额_In   病人预交记录.冲预交%Type,
    交易流水号_In 病人预交记录.交易流水号%Type,
    交易说明_In   病人预交记录.交易说明%Type,
    摘要_In       病人预交记录.摘要%Type,
    
    Xmlexpned_In Xmltype
  ) Is
    n_卡类别id 医疗卡类别.Id%Type;
    v_结算方式 病人预交记录.结算方式%Type;
  
  Begin
    v_Err_Msg := Null;
    Begin
      n_卡类别id := To_Number(卡类别_In);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
  
    If n_卡类别id = 0 Then
      Begin
        Select 编号, 结算方式, Decode(Nvl(启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_结算方式, v_Err_Msg
        From 消费卡类别目录
        Where 名称 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := '消费:' || 卡类别_In || '不存在!';
      End;
    
    Else
    
      Begin
        Select 编号, 结算方式, Decode(Nvl(启用, 0), 1, Null, 名称 || '未启用,不允许进行缴费!')
        Into n_卡类别id, v_结算方式, v_Err_Msg
        From 消费卡类别目录
        Where 编号 = 卡类别_In;
      Exception
        When Others Then
          v_Err_Msg := '未找到指定的结算支付信息!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  
    v_退费结算 := v_结算方式 || '|' || 退款金额_In || '|' || ' |' || Nvl(摘要_In, ' ');
    --   4-消费卡结算: 
    --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额|| 
    --     ②退支票额_In:传入零 
    Select ID
    Into n_消费卡id
    From 消费卡信息
    Where 接口编号 = n_卡类别id And 卡号 = 卡号_In And
          序号 = (Select Max(序号) From 消费卡信息 Where 接口编号 = n_卡类别id And 卡号 = 卡号_In);
  
    --卡类别ID|卡号|消费卡ID|消费金额||. 
    v_退费结算 := n_卡类别id || '|' || 卡号_In || '|' || n_消费卡id || '|' || 退款金额_In;
    Zl_门诊退费结算_Modify(4, n_病人id, 冲销id_In, v_退费结算, 0, Null, 卡号_In, 交易流水号_In, 交易说明_In, 0, 0, 0, 0);
  
    --保存扩展结算信息 
    For c_扩展 In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 1, 卡号_In, 冲销id_In, c_扩展.Jymc || '|' || c_扩展.Jylr, 0);
    End Loop;
  End Square_Cardbalance_Modfiy;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  --0.获取入参中的病人ID等信息 
  Select To_Number(Extractvalue(Value(A), 'IN/BRID')), To_Number(Extractvalue(Value(A), 'IN/JE')),
         Extractvalue(Value(A), 'IN/TFZY'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_病人id, n_退款总额, v_摘要, n_检查发票, v_结算卡类别, v_身份证号, v_姓名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;

  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '无法确定病人信息,请检查!';
    Raise Err_Item;
  End If;

  --人员id,人员编号,人员姓名 
  v_Temp       := Zl_Identity(1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员姓名 := v_Temp;
  v_Err_Msg    := Null;
  v_结帐ids    := Null;

  If v_结算卡类别 Is Not Null Then
    Begin
      n_卡类别id := To_Number(v_结算卡类别);
    Exception
      When Others Then
        n_卡类别id := 0;
    End;
    If n_卡类别id = 0 Then
      Begin
        Select ID Into n_卡类别id From 医疗卡类别 Where 名称 = v_结算卡类别;
      Exception
        When Others Then
          v_Err_Msg := '无法确认传入的结算卡！';
          Raise Err_Item;
      End;
    End If;
  Else
    n_卡类别id := 0;
  End If;

  If Nvl(n_卡类别id, 0) <> 0 Then
    Select 结算方式 Into v_结算方式 From 医疗卡类别 Where ID = n_卡类别id;
  End If;

  --1.先进行退费 

  Select 病人结帐记录_Id.Nextval, Sysdate Into n_冲销id, d_退费时间 From Dual;

  n_Count      := 0;
  n_原结算序号 := 0;
  For c_费用 In (Select Extractvalue(b.Column_Value, '/FY/DJH') As 单据号, Extractvalue(b.Column_Value, '/FY/XH') As 退款序号
               From Table(Xmlsequence(Extract(Xml_In, '/IN/FYLIST/FY'))) B) Loop
    Begin
      Select 结算序号, 结帐id, 病人id
      Into n_Temp, n_结帐id, n_单据病人id
      From 病人预交记录
      Where 结帐id In (Select 结帐id
                     From 门诊费用记录
                     Where NO = c_费用.单据号 And 记录性质 = 1 And Nvl(费用状态, 0) = 0 And 记录状态 In (1, 3) And Rownum < 2) And
            Rownum < 2;
    Exception
      When Others Then
        n_Temp := Null;
    End;
    If Instr(',' || v_结帐ids || ',', ',' || n_结帐id || ',') = 0 Then
      v_结帐ids := v_结帐ids || ',' || n_结帐id;
    End If;
  
    If n_Temp Is Null Then
      v_Err_Msg := '指定的单据号:' || c_费用.单据号 || '未找到,不能退费!';
      Raise Err_Item;
    End If;
  
    For c_挂号 In (Select b.No As 挂号单, b.收费单
                 From 门诊费用记录 A, 病人挂号记录 B
                 Where a.No = c_费用.单据号 And a.记录性质 = 1 And Nvl(费用状态, 0) = 0 And a.记录状态 In (1, 3) And a.挂号id = b.Id And
                       Instr(',' || b.收费单 || ',', ',' || c_费用.单据号 || ',') > 0 And Rownum < 2) Loop
      Select /*+ cardinality(b, 10) */
       Count(1)
      Into n_Column
      From 门诊费用记录 A
      Where a.记录性质 = 1 And a.No In (Select Column_Value From Table(f_Str2list(c_挂号.收费单))) And a.记录状态 = 1 And
            a.No <> c_费用.单据号 And 序号 = 1;
      If n_Column = 0 Then
        If Not c_挂号.挂号单 Is Null Then
          l_挂号单.Extend;
          l_挂号单(l_挂号单.Count) := c_挂号.挂号单;
        End If;
      End If;
    End Loop;
  
    If Nvl(n_单据病人id, 0) = 0 Then
      Begin
        Select 病人id
        Into n_单据病人id
        From 门诊费用记录
        Where NO = c_费用.单据号 And 记录性质 = 1 And Nvl(费用状态, 0) = 0 And 记录状态 In (1, 3) And Rownum < 2;
      Exception
        When Others Then
          n_单据病人id := 0;
      End;
    End If;
  
    If Nvl(n_病人id, 0) <> Nvl(n_单据病人id, 0) Then
      v_Err_Msg := '本次退费的收费单:' || c_费用.单据号 || '不是当前病人的收费单,不能退费!';
      Raise Err_Item;
    End If;
  
    If n_原结算序号 <> 0 And n_原结算序号 <> n_Temp Then
      v_Err_Msg := '本次退费的单据号不是一次收费结算,不能退费!';
      Raise Err_Item;
    End If;
    n_原结算序号 := n_Temp;
  
    Select Count(*) Into n_Temp From 费用补充记录 Where 收费结帐id = n_结帐id;
    If Nvl(n_Temp, 0) <> 0 Then
      v_Err_Msg := '本次退费的单据号已经进行了保险补充结算,不能退费!';
      Raise Err_Item;
    End If;
  
    If v_结算卡类别 Is Not Null Then
      Select Count(*) Into n_Temp From 病人预交记录 Where 结帐id = n_结帐id And 结算方式 = v_结算方式;
      If Nvl(n_Temp, 0) = 0 Then
        v_Err_Msg := '本次退费的单据不是' || v_结算方式 || '结算的,不能退费!';
        Raise Err_Item;
      End If;
    End If;
  
    If Nvl(n_检查发票, 0) = 1 Then
      Select Max(Decode(a.实际票号, Null, 0, 1))
      Into n_是否打印
      From 门诊费用记录 A
      Where NO = c_费用.单据号 And 记录性质 = 1;
      If Nvl(n_是否打印, 0) = 1 Then
        v_Err_Msg := '本次退费的单据号已开发票,不能退费!';
        Raise Err_Item;
      End If;
    End If;
  
    Zl_门诊收费记录_销帐(c_费用.单据号, v_操作员编码, v_操作员姓名, c_费用.退款序号, d_退费时间, v_摘要, n_冲销id);
    n_Count := n_Count + 1;
  End Loop;
  If n_Count = 0 Then
    v_Err_Msg := '未确定本次需要退费的单据,不能退费!';
    Raise Err_Item;
  End If;

  --2.处理退费的结算信息 

  n_结帐金额 := 0;

  --检查总金额是否正确 
  Select Sum(结帐金额) Into n_结帐金额 From 门诊费用记录 Where 结帐id = n_冲销id;

  n_误差额 := -1 * Nvl(n_结帐金额, 0) - Nvl(n_退款总额, 0);
  If Abs(n_误差额) > 1.00 Then
    v_Err_Msg := '单据缴款金额与实际结算的误差太大!';
    Raise Err_Item;
  End If;

  --2.确定支付方式 
  n_Count := 0;
  For c_结算方式 In (Select Extractvalue(b.Column_Value, '/TK/TKKLB') As 卡类别, Extractvalue(b.Column_Value, '/TK/TKKH') As 卡号,
                        Extractvalue(b.Column_Value, '/TK/TKFS') As 结算方式,
                        -1 * To_Number(Extractvalue(b.Column_Value, '/TK/TKJE')) As 退款金额,
                        Extractvalue(b.Column_Value, '/TK/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/TK/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/TK/TKZY') As 摘要,
                        Extractvalue(b.Column_Value, '/TK/TYJK') As 是否退预交,
                        Extractvalue(b.Column_Value, '/TK/SFXFK') As 是否消费卡,
                        Extract(b.Column_Value, '/TK/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/TKLIST/TK'))) B) Loop
  
    --1.退回三方卡 
    If c_结算方式.卡类别 Is Not Null And Nvl(c_结算方式.是否消费卡, 0) = 0 Then
      --1.三方卡结算 
      Third_Cardbalance_Modfiy(n_冲销id, c_结算方式.卡类别, c_结算方式.卡号, c_结算方式.退款金额, c_结算方式.交易流水号, c_结算方式.交易说明, c_结算方式.摘要,
                               c_结算方式.Expend);
    Elsif c_结算方式.卡类别 Is Not Null And Nvl(c_结算方式.是否消费卡, 0) = 1 Then
      --2.消费卡结算 
      Square_Cardbalance_Modfiy(n_冲销id, c_结算方式.卡类别, c_结算方式.卡号, c_结算方式.退款金额, c_结算方式.交易流水号, c_结算方式.交易说明, c_结算方式.摘要,
                                c_结算方式.Expend);
    Elsif Nvl(c_结算方式.是否退预交, 0) = 1 Then
      --3.退预交款 
      Zl_门诊退费结算_Modify(4, n_病人id, n_冲销id, Null, c_结算方式.退款金额, Null, Null, Null, Null, 0, 0, 0, 0);
    Else
      --4.普通结算 
      If c_结算方式.结算方式 Is Null Then
        v_Err_Msg := '未指定指付方式，不允缴款!';
        Raise Err_Item;
      End If;
      --结算方式|结算金额|结算号码|结算摘要||.. 
      v_退费结算 := c_结算方式.结算方式 || '|' || c_结算方式.退款金额 || '| |' || Nvl(c_结算方式.摘要, '  ');
      v_普通结算 := Nvl(v_普通结算, '') || '||' || v_退费结算;
    End If;
    n_Count := n_Count + 1;
  End Loop;

  --   0-原样退 
  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0 
  --   1-普通退费方式: 
  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空. 
  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新) 
  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.. 
  --     ②退支票额_In:传入零 
  If n_Count = 0 Then
    v_Err_Msg := '不能有效确认当前的支付方式!';
    Raise Err_Item;
  End If;

  --5.普通结算及完成结 
  If v_普通结算 Is Not Null Then
    v_普通结算 := Substr(v_普通结算, 3);
  End If;
  Zl_门诊退费结算_Modify(1, n_病人id, n_冲销id, v_普通结算, 0, Null, Null, Null, Null, 0, 0, n_误差额, 2);

  If v_结帐ids Is Not Null Then
    v_结帐ids := Substr(v_结帐ids, 2);
  End If;

  If l_挂号单.Count <> 0 Then
    For I In 0 .. l_挂号单.Count Loop
      x_Templet := Xmltype('<IN></IN>');
      v_Temp    := '<GHDH>' || l_挂号单(I) || '</GHDH>';
      Select Appendchildxml(x_Templet, '/IN', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<JSKLB>' || v_结算卡类别 || '</JSKLB>';
      Select Appendchildxml(x_Templet, '/IN', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Temp := '<GHJE>' || 0 || '</GHJE>';
      Select Appendchildxml(x_Templet, '/IN', Xmltype(v_Temp)) Into x_Templet From Dual;
      Zl_Third_Registdel(x_Templet, Xml_Out);
    End Loop;
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_退费时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<YJZID>' || v_结帐ids || '</YJZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CXID>' || n_冲销id || '</CXID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Charge_Del;
/

--124306:焦博,2018-04-17,调整Oracle过程Zl_Third_Charge_Del和Zl_Third_Registdel
CREATE OR REPLACE Procedure Zl_Third_Registdel
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:HIS退号 
  --入参:Xml_In: 
  --<IN> 
  --  <GHDH>A000001</GHDH>    //挂号单号 
  --  <JSKLB>支付宝</JSKLB>      //结算卡类别 
  --  <JCFP>1</JCFP>            //检查发票 
  --  <GHJE>20</GHJE>            //挂号金额 
  --  <LSH>34563</LSH>           //交易流水号 
  --  <JKFS>0</JKFS>             //缴款方式,0-挂号或预约缴款;1-预约不缴款 
  --  <YYFS></YYFS>              //缴款方式=1时传入，预约的预约方式 
  --</IN> 

  --出参:Xml_Out 
  --<OUTPUT> 
  -- <CZSJ>操作时间</CZSJ>          //HIS的登记时间 
  -- <YJZID>原结帐ID</YJZID> 
  -- <CXID>冲销ID</CXID> 
  -- <ERROR><MSG></MSG></ERROR> //为空表示取消挂号成功 
  --</OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  v_卡类别     Varchar2(100);
  v_No         病人挂号记录.No%Type;
  n_挂号金额   门诊费用记录.实收金额%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_结算方式   医疗卡类别.结算方式%Type;
  n_实收金额   门诊费用记录.实收金额%Type;
  v_交易流水号 病人预交记录.交易流水号%Type;
  n_存在       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --临时XML 
  x_Templet    Xmltype; --模板XML 
  v_Err_Msg    Varchar2(200);
  n_已开医嘱   Number(2);
  n_检查发票   Number(3);
  n_是否打印   Number(3);
  n_缴款方式   Number(3);
  n_结帐id     门诊费用记录.结帐id%Type;
  n_冲销id     门诊费用记录.结帐id%Type;
  d_登记时间   Date;
  v_预约方式   病人挂号记录.预约方式%Type;
  v_收费单     门诊费用记录.No%Type;
  n_病人id     门诊费用记录.病人id%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_退费结算   Varchar2(1000);
  n_Column     Number(18);
  Err_Item Exception;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS')
  Into v_No, v_卡类别, n_挂号金额, v_交易流水号, n_检查发票, n_缴款方式, v_预约方式
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Max(收费单) Into v_收费单 From 病人挂号记录 Where NO = v_No;

  n_缴款方式 := Nvl(n_缴款方式, 0);

  If n_缴款方式 = 1 Then
    Begin
      Select 1 Into n_存在 From 门诊费用记录 Where NO = v_No And 记录性质 = 4 And 结帐id Is Not Null And Rownum < 2;
      Select 1
      Into n_存在
      From 门诊费用记录
      Where NO In (Select Column_Value From Table(f_Str2list(v_收费单)) B) And 记录性质 = 1 And 结帐id Is Not Null And
            Rownum < 2;
    Exception
      When Others Then
        n_存在 := 0;
    End;
    If n_存在 = 1 Then
      v_Err_Msg := '传入的挂号单据不是预约挂号单,无法取消预约!';
      Raise Err_Item;
    End If;
    Begin
      Select 1 Into n_存在 From 病人挂号记录 A Where a.No = v_No And a.预约方式 = v_预约方式 And Rownum < 2;
    Exception
      When Others Then
        n_存在 := 0;
    End;
    If n_存在 = 0 Then
      v_Err_Msg := '传入的挂号单据不是' || v_预约方式 || '预约的,无法取消预约!';
      Raise Err_Item;
    End If;
  End If;

  If v_卡类别 Is Not Null And n_缴款方式 = 0 Then
    Select Nvl2(Translate(v_卡类别, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
    If v_Type = 'Num' Then
      --传入的是卡类别ID 
      Select 结算方式, ID Into v_结算方式, n_卡类别id From 医疗卡类别 Where ID = To_Number(v_卡类别);
    Else
      --传入的是卡类别名称 
      Select 结算方式, ID Into v_结算方式, n_卡类别id From 医疗卡类别 Where 名称 = v_卡类别;
    End If;
  
    Select Sum(实收金额) Into n_实收金额 From 门诊费用记录 Where NO = v_No And 记录性质 = 4;
  
    If Nvl(n_缴款方式, 0) = 0 Then
      --要退的单据不是以该结算卡结算的，则禁止退号 
      Begin
        Select 1
        Into n_存在
        From 病人预交记录 A,
             (Select Distinct 结帐id
               From 门诊费用记录
               Where NO = v_No And 记录性质 = 4
               Union
               Select Distinct 结帐id
               From 住院费用记录
               Where NO = v_No And 记录性质 = 5
               Union
               Select Distinct 结帐id
               From 门诊费用记录
               Where NO In (Select Column_Value From Table(f_Str2list(v_收费单)) B) And 记录性质 = 1) B
        Where a.结帐id = b.结帐id And 结算方式 = v_结算方式 And Rownum < 2;
      Exception
        When Others Then
          n_存在 := 0;
      End;
      If n_存在 = 0 Then
        v_Err_Msg := '传入的挂号单据不是' || v_结算方式 || '结算的,无法退号!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --补充结算检查，已存在补结算数据的，不能退号 
  Begin
    Select 1
    Into n_存在
    From 费用补充记录 A,
         (Select Distinct 结帐id
           From 门诊费用记录
           Where NO = v_No And 记录性质 = 4
           Union
           Select Distinct 结帐id
           From 住院费用记录
           Where NO = v_No And 记录性质 = 5
           Union
           Select Distinct 结帐id
           From 门诊费用记录
           Where NO In (Select Column_Value From Table(f_Str2list(v_收费单)) B) And 记录性质 = 1) B
    Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 1 And Nvl(a.费用状态, 0) <> 2 And Rownum < 2;
  Exception
    When Others Then
      n_存在 := 0;
  End;
  If n_存在 = 1 Then
    v_Err_Msg := '传入的挂号单据已经进行了二次结算,无法退号!';
    Raise Err_Item;
  End If;
  --医嘱检查，已经开过医嘱的，不能退号 
  Begin
    Select Distinct 1 Into n_已开医嘱 From 病人医嘱记录 Where 挂号单 = v_No;
  Exception
    When Others Then
      n_已开医嘱 := 0;
  End;
  If n_已开医嘱 = 1 Then
    v_Err_Msg := '传入的挂号单据已经开过医嘱,无法退号!';
    Raise Err_Item;
  End If;
  If Nvl(n_检查发票, 0) = 1 Then
    Select Max(Decode(a.实际票号, Null, 0, 1)) Into n_是否打印 From 门诊费用记录 A Where NO = v_No And 记录性质 = 4;
    If Nvl(n_是否打印, 0) = 1 Then
      v_Err_Msg := '本次退号的单据已开发票,不能退费!';
      Raise Err_Item;
    End If;
    Select Max(Decode(a.实际票号, Null, 0, 1))
    Into n_是否打印
    From 门诊费用记录 A
    Where NO In (Select Column_Value From Table(f_Str2list(v_收费单)) B) And 记录性质 = 1;
    If Nvl(n_是否打印, 0) = 1 Then
      v_Err_Msg := '本次退号的单据已开发票,不能退费!';
      Raise Err_Item;
    End If;
  End If;
  --获取操作员信息 
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
  d_登记时间 := Sysdate;

  Zl_三方机构挂号_Delete(v_No, v_交易流水号, '移动平台退号', d_登记时间);

  --同步处理划价单 
  If v_收费单 Is Not Null Then
  
    n_Column := 0;
    For c_挂号 In (Select NO, Max(记录状态) As 记录状态, Max(病人id) As 病人id, Max(Decode(记录状态, 2, 0, 结帐id)) As 原结帐id,
                        Max(Decode(记录状态, 2, 结帐id, 0)) As 冲销id
                 From 门诊费用记录
                 Where NO In (Select * From Table(f_Str2list(v_收费单)) B) And 记录性质 = 1) Loop
      If Nvl(c_挂号.记录状态, 0) = 0 Then
        Zl_门诊划价记录_Delete(c_挂号.No);
        n_结帐id := c_挂号.原结帐id;
        n_冲销id := c_挂号.冲销id;
      Elsif Nvl(c_挂号.记录状态, 0) = 1 Then
        If v_结算方式 Is Null Then
          v_Err_Msg := '本次挂号单据退款失败,请检查!';
          Raise Err_Item;
        End If;
        Select 病人结帐记录_Id.Nextval Into n_冲销id From Dual;
        Zl_门诊收费记录_销帐(c_挂号.No, v_操作员编号, v_操作员姓名, Null, d_登记时间, Null, n_冲销id);
        v_退费结算 := v_结算方式 || '|' || -1 * n_挂号金额 || '|' || ' |' || ' ';
        Zl_门诊退费结算_Modify(2, n_病人id, n_冲销id, v_退费结算, 0, n_卡类别id, Null, v_交易流水号, Null, 0, 0, 0, 2);
        n_结帐id := c_挂号.原结帐id;
        n_冲销id := c_挂号.冲销id;
        n_Column := n_Column + 1;
      Else
        n_结帐id := c_挂号.原结帐id;
        n_冲销id := c_挂号.冲销id;
      End If;
    
    End Loop;
    If n_Column > 1 Then
      v_Err_Msg := '本次挂号存在多次收费，请先退费后再退号!';
      Raise Err_Item;
    End If;
  
  Else
  
    Select Max(结帐id) Into n_结帐id From 门诊费用记录 Where NO = v_No And 记录性质 = 4 And 记录状态 = 3;
    Select Max(结帐id) Into n_冲销id From 门诊费用记录 Where NO = v_No And 记录性质 = 4 And 记录状态 = 2;
  
  End If;

  v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<YJZID>' || n_结帐id || '</YJZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CXID>' || n_冲销id || '</CXID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registdel;
/

--124650:李业庆,2018-04-20,材料特性增加是否分零属性
Create Or Replace Procedure Zl_卫生材料_Insert
(
  诊疗id_In         In 材料特性.诊疗id%Type,
  材料id_In         In 材料特性.材料id%Type,
  编码_In           In 收费项目目录.编码%Type,
  规格_In           In 收费项目目录.规格%Type,
  产地_In           In 收费项目目录.产地%Type := Null,
  标识主码_In       In 收费项目目录.标识主码%Type := Null,
  标识子码_In       In 收费项目目录.标识子码%Type := Null,
  备选码_In         In 收费项目目录.备选码%Type := Null,
  材料来源_In       In 材料特性.材料来源%Type := Null,
  货源情况_In       In 材料特性.货源情况%Type := Null,
  散装单位_In       In 收费项目目录.计算单位%Type := Null,
  包装单位_In       In 材料特性.包装单位%Type := Null,
  换算系数_In       In 材料特性.换算系数%Type := Null,
  是否变价_In       In 收费项目目录.是否变价%Type := Null,
  指导批发价_In     In 材料特性.指导批发价%Type := Null,
  扣率_In           In 材料特性.扣率%Type := 95,
  指导零售价_In     In 材料特性.指导零售价%Type := Null,
  指导差价率_In     In 材料特性.指导差价率%Type := Null,
  费用类型_In       In 收费项目目录.费用类型%Type := Null,
  服务对象_In       In 收费项目目录.服务对象%Type := Null,
  屏蔽费别_In       In 收费项目目录.屏蔽费别%Type := 0,
  库房分批_In       In 材料特性.库房分批%Type := Null,
  在用分批_In       In 材料特性.在用分批%Type := Null,
  最大效期_In       In 材料特性.最大效期%Type := Null,
  灭菌效期_In       In 材料特性.灭菌效期%Type := Null,
  无菌性材料_In     In 材料特性.无菌性材料%Type := Null,
  一次性材料_In     In 材料特性.一次性材料%Type := Null,
  原材料_In         In 材料特性.原材料%Type := Null,
  差价让利比_In     In 材料特性.差价让利比%Type := 0,
  成本价_In         In 材料特性.成本价%Type := 0,
  跟踪在用_In       In 材料特性.跟踪在用%Type := Null,
  核算材料_In       In 材料特性.核算材料%Type := 0,
  当前售价_In       In 收费价目.现价%Type := 0,
  收入id_In         In 收费价目.收入项目id%Type := Null,
  批准文号_In       In 材料特性.批准文号%Type := Null,
  注册商标_In       In 材料特性.注册商标%Type := Null,
  注册证号_In       In 材料特性.注册证号%Type := Null,
  许可证号_In       In 材料特性.许可证号%Type := Null,
  许可证有效期_In   In 材料特性.许可证有效期%Type := Null,
  材质分类_In       In 材料特性.材质分类%Type := Null,
  存储条件_In       In 材料特性.存储条件%Type := Null,
  跟踪病人_In       In 材料特性.跟踪病人%Type := 0,
  站点_In           In 收费项目目录.站点%Type := Null,
  品名_In           In 收费项目别名.名称%Type := Null,
  拼音_In           In 收费项目别名.简码%Type := Null,
  五笔_In           In 收费项目别名.简码%Type := Null,
  增值税率_In       In 材料特性.增值税率%Type := Null,
  说明_In           In 收费项目目录.说明%Type := Null,
  高值材料_In       In 材料特性.高值材料%Type := Null,
  条码管理_In       In 材料特性.是否条码管理%Type := Null,
  病案费目_In       In 收费项目目录.病案费目%Type := Null,
  器械包卫材单件_In In 材料特性.器械包卫材单件%Type := 0,
  注册证有效期_In   In 材料特性.注册证有效期%Type := Null,
  是否植入耗材_In   In 材料特性.是否植入耗材%Type := 0,
  加成率_In         In 材料特性.加成率%Type := Null,
  分零使用_In       In 材料特性.是否分零%Type := 0
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  v_No       收费价目.No%Type;
  v_名称     诊疗项目目录.名称%Type;
  v_Temp     收费项目目录.病案费目%Type;
  v_病案费目 收费项目目录.病案费目%Type;

  Cursor c_Item Is
    Select ID
    From 部门表 D
    Where ID In (Select Distinct 部门id
                 From 部门性质说明 A
                 Where 工作性质 In ('发料部门', '物资库房', '卫材库', '制剂室', '虚拟库房'));
Begin
  v_Err_Msg := 'NO';

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
  Begin
    Select 名称
    Into v_名称
    From 诊疗项目目录
    Where ID = 诊疗id_In And (撤档时间 Is Null Or To_Char(撤档时间, 'yyyy-mm-dd') = '3000-01-01');
  Exception
    When Others Then
      v_Err_Msg := 'Err';
  End;
  If v_Err_Msg = 'Err' Then
    v_Err_Msg := '[ZLSOFT]未找到指定的材料品种，可能该品种已被其他用户删除或停用！[ZLSOFT]';
    Raise Err_Item;
  End If;
  --规格信息 
  Insert Into 收费项目目录
    (类别, ID, 编码, 名称, 规格, 产地, 标识主码, 标识子码, 备选码, 计算单位, 费用类型, 服务对象, 屏蔽费别, 是否变价, 站点, 建档时间, 撤档时间, 说明, 病案费目)
  Values
    (4, 材料id_In, 编码_In, v_名称, 规格_In, 产地_In, 标识主码_In, 标识子码_In, 备选码_In, 散装单位_In, 费用类型_In, 服务对象_In, 屏蔽费别_In, 是否变价_In,
     站点_In, Sysdate, To_Date('3000-01-01', 'YYYY-MM-DD'), 说明_In, v_病案费目);

  --材料特性 
  Insert Into 材料特性
    (材料id, 诊疗id, 最大效期, 灭菌效期, 无菌性材料, 一次性材料, 原材料, 货源情况, 包装单位, 换算系数, 指导批发价, 指导零售价, 指导差价率, 扣率, 库房分批, 在用分批, 材料来源, 差价让利比, 成本价,
     跟踪在用, 核算材料, 批准文号, 注册商标, 注册证号, 注册证有效期, 许可证号, 许可证有效期, 材质分类, 存储条件, 跟踪病人, 增值税率, 高值材料, 是否条码管理, 器械包卫材单件, 是否植入耗材, 加成率,
     是否分零)
  Values
    (材料id_In, 诊疗id_In, 最大效期_In, 灭菌效期_In, 无菌性材料_In, 一次性材料_In, 原材料_In, 货源情况_In, 包装单位_In, 换算系数_In, 指导批发价_In, 指导零售价_In,
     指导差价率_In, 扣率_In, 库房分批_In, 在用分批_In, 材料来源_In, 差价让利比_In, 成本价_In, 跟踪在用_In, 核算材料_In, 批准文号_In, 注册商标_In, 注册证号_In,
     注册证有效期_In, 许可证号_In, 许可证有效期_In, 材质分类_In, 存储条件_In, 跟踪病人_In, 增值税率_In, 高值材料_In, 条码管理_In, 器械包卫材单件_In, 是否植入耗材_In, 加成率_In,
     分零使用_In);

  --别名的处理 
  Insert Into 收费项目别名
    (收费细目id, 名称, 性质, 简码, 码类)
    Select 材料id_In, 名称, 性质, 简码, 码类 From 诊疗项目别名 Where 诊疗项目id = 诊疗id_In;
  If (品名_In Is Not Null) And (拼音_In Is Not Null) Then
    Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (材料id_In, 品名_In, 3, 拼音_In, 1);
  End If;
  If (品名_In Is Not Null) And (五笔_In Is Not Null) Then
    Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (材料id_In, 品名_In, 3, 五笔_In, 2);
  End If;

  For r_Item In c_Item Loop
    Insert Into 材料储备限额 (库房id, 材料id, 上限, 下限, 盘点属性) Values (r_Item.Id, 材料id_In, 0, 0, '1111');
  End Loop;
  --定价信息 
  If 收入id_In Is Not Null Then
    v_No := Nextno(9);
    --非跟踪在用的时价卫材，在销售时相当于一般的收费项目，在调价时应按”最低限价、最高限价、缺省价格”进行设置。 
    Insert Into 收费价目
      (ID, 原价id, 收费细目id, 原价, 现价, 缺省价格, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
    Values
      (收费价目_Id.Nextval, Null, 材料id_In, 0, 当前售价_In, Decode(跟踪在用_In, 0, Decode(是否变价_In, 1, 当前售价_In, Null), Null), 收入id_In,
       1, '新增定价', User, Sysdate, To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
  End If;

  --材料生产商比较增加 
  If 产地_In Is Not Null Then
    Update 材料生产商 Set 名称 = 产地_In Where 名称 = 产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 材料生产商
        (编码, 名称, 简码)
        Select Nvl(Max(To_Number(编码)), 0) + 1, 产地_In, zlSpellCode(产地_In) From 材料生产商;
    End If;
  End If;

  --插入材料的服务科室 
  Insert Into 收费执行科室
    (收费细目id, 病人来源, 开单科室id, 执行科室id)
    Select 材料id_In, 病人来源, 开单科室id, 执行科室id From 诊疗执行科室 Where 诊疗项目id = 诊疗id_In;

  b_Message.Zlhis_Dict_043(材料id_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_卫生材料_Insert;
/

--124650:李业庆,2018-04-20,材料特性增加是否分零属性
Create Or Replace Procedure Zl_卫生材料_Update
(
  诊疗id_In         In 材料特性.诊疗id%Type,
  材料id_In         In 材料特性.材料id%Type,
  编码_In           In 收费项目目录.编码%Type,
  规格_In           In 收费项目目录.规格%Type,
  产地_In           In 收费项目目录.产地%Type := Null,
  标识主码_In       In 收费项目目录.标识主码%Type := Null,
  标识子码_In       In 收费项目目录.标识子码%Type := Null,
  备选码_In         In 收费项目目录.备选码%Type := Null,
  材料来源_In       In 材料特性.材料来源%Type := Null,
  货源情况_In       In 材料特性.货源情况%Type := Null,
  散装单位_In       In 收费项目目录.计算单位%Type := Null,
  包装单位_In       In 材料特性.包装单位%Type := Null,
  换算系数_In       In 材料特性.换算系数%Type := Null,
  是否变价_In       In 收费项目目录.是否变价%Type := Null,
  指导批发价_In     In 材料特性.指导批发价%Type := Null,
  扣率_In           In 材料特性.扣率%Type := 95,
  指导零售价_In     In 材料特性.指导零售价%Type := Null,
  指导差价率_In     In 材料特性.指导差价率%Type := Null,
  费用类型_In       In 收费项目目录.费用类型%Type := Null,
  服务对象_In       In 收费项目目录.服务对象%Type := Null,
  屏蔽费别_In       In 收费项目目录.屏蔽费别%Type := 0,
  库房分批_In       In 材料特性.库房分批%Type := Null,
  在用分批_In       In 材料特性.在用分批%Type := Null,
  最大效期_In       In 材料特性.最大效期%Type := Null,
  灭菌效期_In       In 材料特性.灭菌效期%Type := Null,
  无菌性材料_In     In 材料特性.无菌性材料%Type := Null,
  一次性材料_In     In 材料特性.一次性材料%Type := Null,
  原材料_In         In 材料特性.原材料%Type := Null,
  差价让利比_In     In 材料特性.差价让利比%Type := 0,
  成本价_In         In 材料特性.成本价%Type := 0,
  跟踪在用_In       In 材料特性.跟踪在用%Type := Null,
  核算材料_In       In 材料特性.核算材料%Type := 0,
  当前售价_In       In 收费价目.现价%Type := 0,
  收入id_In         In 收费价目.收入项目id%Type := Null,
  批准文号_In       In 材料特性.批准文号%Type := Null,
  注册商标_In       In 材料特性.注册商标%Type := Null,
  注册证号_In       In 材料特性.注册证号%Type := Null,
  许可证号_In       In 材料特性.许可证号%Type := Null,
  许可证有效期_In   In 材料特性.许可证有效期%Type := Null,
  材质分类_In       In 材料特性.材质分类%Type := Null,
  存储条件_In       In 材料特性.存储条件%Type := Null,
  跟踪病人_In       In 材料特性.跟踪病人%Type := 0,
  站点_In           In 收费项目目录.站点%Type := Null,
  品名_In           In 收费项目别名.名称%Type := Null,
  拼音_In           In 收费项目别名.简码%Type := Null,
  五笔_In           In 收费项目别名.简码%Type := Null,
  增值税率_In       In 材料特性.增值税率%Type := Null,
  说明_In           In 收费项目目录.说明%Type := Null,
  高值材料_In       In 材料特性.高值材料%Type := Null,
  条码管理_In       In 材料特性.是否条码管理%Type := Null,
  病案费目_In       In 收费项目目录.病案费目%Type := Null,
  器械包卫材单件_In In 材料特性.器械包卫材单件%Type := 0,
  注册证有效期_In   In 材料特性.注册证有效期%Type := Null,
  是否植入耗材_In   In 材料特性.是否植入耗材%Type := 0,
  修改类型_In       In Number := 0, --1-同步修改品种下所有注册证号和注册证有效期
  加成率_In         In 材料特性.加成率%Type := Null,
  分零使用_In       In 材料特性.是否分零%Type := 0
) Is
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  v_发生     Integer;
  v_跟踪在用 Integer;
  v_Count    Integer;
  v_No       收费价目.No%Type;
  v_名称     诊疗项目目录.名称%Type;
  v_Temp     收费项目目录.病案费目%Type;
  v_病案费目 收费项目目录.病案费目%Type;

Begin
  v_Err_Msg := '无';

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
  --修改诊疗项目 
  Begin
    Select 跟踪在用 Into v_跟踪在用 From 材料特性 Where 材料id = 材料id_In;
  Exception
    When Others Then
      v_Err_Msg := '[ZLSOFT]不存在规格材料,可能被其他用户删除了,请检查![ZLSOFT]';
  End;
  If v_Err_Msg <> '无' Then
    Raise Err_Item;
  End If;

  Begin
    Select 名称 Into v_名称 From 诊疗项目目录 Where ID = 诊疗id_In;
  Exception
    When Others Then
      v_Err_Msg := 'Err';
  End;

  If v_Err_Msg = 'Err' Then
    v_Err_Msg := '[ZLSOFT]未找到指定的材料品种，可能已被其他用户删除！[ZLSOFT]';
    Raise Err_Item;
  End If;

  --如果更新前的材料为跟踪在用,如果改为了不跟踪则需判断库存 
  If v_跟踪在用 = 1 And 跟踪在用_In <> 1 Then
    Begin
      Select Count(*)
      Into v_Count
      From 药品库存
      Where 药品id = 材料id_In And (Nvl(可用数量, 0) <> 0 Or Nvl(实际数量, 0) <> 0 Or Nvl(实际金额, 0) <> 0 Or Nvl(实际差价, 0) <> 0);
      If v_Count <> 0 Then
        v_Err_Msg := '[ZLSOFT]该卫生材料存在库存,不能取消跟踪在用属性,请检查![ZLSOFT]';
      End If;
    Exception
      When Others Then
        Null;
    End;
  End If;

  If v_Err_Msg <> '无' Then
    Raise Err_Item;
  End If;

  --规格信息 
  Update 收费项目目录
  Set 编码 = 编码_In, 名称 = v_名称, 规格 = 规格_In, 标识主码 = 标识主码_In, 标识子码 = 标识子码_In, 备选码 = 备选码_In, 产地 = 产地_In, 是否变价 = 是否变价_In,
      计算单位 = 散装单位_In, 费用类型 = 费用类型_In, 服务对象 = 服务对象_In, 屏蔽费别 = 屏蔽费别_In, 站点 = 站点_In, 说明 = 说明_In, 病案费目 = v_病案费目
  Where ID = 材料id_In;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]该卫生材料可能被其他用户删除了,请检查![ZLSOFT]';
    Raise Err_Item;
  End If;

  --材料特性 
  Update 材料特性
  Set 最大效期 = 最大效期_In, 灭菌效期 = 灭菌效期_In, 无菌性材料 = 无菌性材料_In, 一次性材料 = 一次性材料_In, 原材料 = 原材料_In, 货源情况 = 货源情况_In, 包装单位 = 包装单位_In,
      换算系数 = 换算系数_In, 指导批发价 = 指导批发价_In, 指导零售价 = 指导零售价_In, 指导差价率 = 指导差价率_In, 扣率 = 扣率_In, 库房分批 = 库房分批_In, 在用分批 = 在用分批_In,
      材料来源 = 材料来源_In, 差价让利比 = 差价让利比_In, 成本价 = 成本价_In, 跟踪在用 = 跟踪在用_In, 核算材料 = 核算材料_In, 批准文号 = 批准文号_In, 注册商标 = 注册商标_In,
      注册证号 = 注册证号_In, 注册证有效期 = 注册证有效期_In, 材质分类 = 材质分类_In, 存储条件 = 存储条件_In, 许可证号 = 许可证号_In, 许可证有效期 = 许可证有效期_In,
      诊疗id = 诊疗id_In, 跟踪病人 = 跟踪病人_In, 增值税率 = 增值税率_In, 高值材料 = 高值材料_In, 是否条码管理 = 条码管理_In, 器械包卫材单件 = 器械包卫材单件_In,
      是否植入耗材 = 是否植入耗材_In, 加成率 = 加成率_In, 是否分零 = 分零使用_In
  Where 材料id = 材料id_In;

  --同步修改改品种下所有规格
  If 修改类型_In = 1 Then
    Update 材料特性 Set 注册证号 = 注册证号_In, 注册证有效期 = 注册证有效期_In Where 诊疗id = 诊疗id_In;
  End If;

  --别名的处理 
  Delete 收费项目别名 Where 收费细目id = 材料id_In And 性质 = 1;

  Insert Into 收费项目别名
    (收费细目id, 名称, 性质, 简码, 码类)
    Select 材料id_In, 名称, 性质, 简码, 码类 From 诊疗项目别名 Where 诊疗项目id = 诊疗id_In And 性质 = 1;

  If 品名_In Is Null Then
    Delete 收费项目别名 Where 收费细目id = 材料id_In And 性质 = 3;
  Else
    If 拼音_In Is Null Then
      Delete 收费项目别名 Where 收费细目id = 材料id_In And 性质 = 3 And 码类 = 1;
    Else
      Update 收费项目别名 Set 名称 = 品名_In, 简码 = 拼音_In Where 收费细目id = 材料id_In And 性质 = 3 And 码类 = 1;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (材料id_In, 品名_In, 3, 拼音_In, 1);
      End If;
    End If;
    If 五笔_In Is Null Then
      Delete 收费项目别名 Where 收费细目id = 材料id_In And 性质 = 3 And 码类 = 2;
    Else
      Update 收费项目别名 Set 名称 = 品名_In, 简码 = 五笔_In Where 收费细目id = 材料id_In And 性质 = 3 And 码类 = 2;
      If Sql%RowCount = 0 Then
        Insert Into 收费项目别名 (收费细目id, 名称, 性质, 简码, 码类) Values (材料id_In, 品名_In, 3, 五笔_In, 2);
      End If;
    End If;
  End If;

  --定价信息：如果已经有发生，则不允许直接更改这些信息 
  Select Nvl(Count(*), 0) Into v_发生 From 药品收发记录 Where 药品id = 材料id_In And Rownum < 2;

  If v_发生 = 0 Then
    Update 收费项目目录 Set 是否变价 = 是否变价_In Where ID = 材料id_In;
    Update 材料特性 Set 成本价 = 成本价_In Where 材料id = 材料id_In;
  
    If 收入id_In Is Not Null Then
      Update 收费价目
      Set 现价 = 当前售价_In, 缺省价格 = Decode(跟踪在用_In, 0, Decode(是否变价_In, 1, 当前售价_In, 缺省价格), 缺省价格), 收入项目id = 收入id_In, 变动原因 = 1,
          调价说明 = '修改定价', 调价人 = User
      Where 收费细目id = 材料id_In
           --And (终止日期 Is Null Or 终止日期=to_date('3000-01-01','YYYY-MM-DD')); 
            And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因 = 1;
    
      If Sql%RowCount = 0 Then
        v_No := Nextno(9);
        Insert Into 收费价目
          (ID, 原价id, 收费细目id, 原价, 现价, 缺省价格, 收入项目id, 变动原因, 调价说明, 调价人, 执行日期, 终止日期, NO, 序号)
        Values
          (收费价目_Id.Nextval, Null, 材料id_In, 0, 当前售价_In, Decode(跟踪在用_In, 0, Decode(是否变价_In, 1, 当前售价_In, Null), Null),
           收入id_In, 1, '新增定价', User, Sysdate, To_Date('3000-01-01', 'YYYY-MM-DD'), v_No, 1);
      End If;
    End If;
  Else
    --有业务单据后不能直接修改价格，但是可以修改收入项目 
    Update 收费价目
    Set 收入项目id = 收入id_In
    Where 收费细目id = 材料id_In And (Sysdate Between 执行日期 And 终止日期 Or Sysdate >= 执行日期 And 终止日期 Is Null) And 变动原因 = 1;
  End If;

  --材料生产商比较增加 
  If 产地_In Is Not Null Then
    Update 材料生产商 Set 名称 = 产地_In Where 名称 = 产地_In;
    If Sql%RowCount = 0 Then
      Insert Into 材料生产商
        (编码, 名称, 简码)
        Select Nvl(Max(To_Number(编码)), 0) + 1, 产地_In, zlSpellCode(产地_In) From 材料生产商;
    End If;
  End If;

  b_Message.Zlhis_Dict_044(材料id_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_卫生材料_Update;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0007' Where 编号=&n_System;
Commit;