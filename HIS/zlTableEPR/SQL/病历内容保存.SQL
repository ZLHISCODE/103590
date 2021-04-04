alter table 电子病历记录 Add 编辑方式 Number(1) Default 0;
Create Or Replace Procedure Zl_病历文件结构_Apply
(
  Id_In         In 病历文件结构.ID%Type,
  文件id_In     In 病历文件结构.文件id%Type,
  父id_In       In 病历文件结构.父id%Type,
  对象序号_IN	In 病历文件结构.对象序号%Type,
  对象类型_In   In 病历文件结构.对象类型%Type,
  保留对象_In   In 病历文件结构.保留对象%Type,
  对象属性_In   In 病历文件结构.对象属性%Type,
  内容行次_In   In 病历文件结构.内容行次%Type,
  内容文本_In   In 病历文件结构.内容文本%Type,
  诊治要素id_In In 病历文件结构.诊治要素id%Type := Null,
  替换域_In     In 病历文件结构.替换域%Type := 0,
  要素名称_In   In 病历文件结构.要素名称%Type := Null,
  要素类型_In   In 病历文件结构.要素类型%Type := Null,
  要素长度_In   In 病历文件结构.要素长度%Type := Null,
  要素小数_In   In 病历文件结构.要素小数%Type := Null,
  要素单位_In   In 病历文件结构.要素单位%Type := Null,
  要素表示_In   In 病历文件结构.要素表示%Type := 0,
  输入形态_In   In 病历文件结构.输入形态%Type := 0,
  要素值域_In   In 病历文件结构.要素值域%Type := Null
) Is
Begin

  Update 病历文件结构
  Set 文件id = 文件id_In, 父id = 父id_In, 对象序号 = 对象序号_IN, 对象类型 = 对象类型_In, 保留对象 = 保留对象_In, 对象属性 = 对象属性_In, 内容行次 = 内容行次_In, 内容文本 = 内容文本_In,
      诊治要素id = 诊治要素id_In, 替换域 = 替换域_In, 要素名称 = 要素名称_In, 要素类型 = 要素类型_In, 要素长度 = 要素长度_In, 要素小数 = 要素小数_In, 要素单位 = 要素单位_In,
      要素表示 = 要素表示_In, 输入形态 = 输入形态_In, 要素值域 = 要素值域_In
  Where ID = Id_In;
  If Sql%RowCount = 0 Then
    Insert Into 病历文件结构
      (ID, 文件id, 父id, 对象序号, 对象类型, 保留对象, 对象属性, 内容行次, 内容文本, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域)
    Values
      (Id_In, 文件id_In, 父id_In, 对象序号_IN, 对象类型_In, 保留对象_In, 对象属性_In, 内容行次_In, 内容文本_In, 诊治要素id_In, 替换域_In, 要素名称_In, 要素类型_In,
       要素长度_In, 要素小数_In, 要素单位_In, 要素表示_In, 输入形态_In, 要素值域_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病历文件结构_Apply;
/

Create Or Replace Procedure Zl_病历文件结构_Delete(文件id_In In 病历文件结构.文件id%Type) Is
Begin

  Delete 病历文件结构 Where 文件id = 文件id_In;
  Update 病历页面格式 Set 图形 = Empty_Blob() Where (种类, 编号) = (Select 种类, 编号 From 病历文件列表 Where ID = 文件id_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病历文件结构_Delete;
/

Create Or Replace Procedure Zl_病历范文内容_Apply
(
  Id_In         In 病历范文内容.ID%Type,
  文件id_In     In 病历范文内容.文件id%Type,
  父id_In       In 病历范文内容.父id%Type,
  对象序号_In   In 病历范文内容.对象序号%Type,
  对象类型_In   In 病历范文内容.对象类型%Type,
  保留对象_In   In 病历范文内容.保留对象%Type,
  对象属性_In   In 病历范文内容.对象属性%Type,
  内容行次_In   In 病历范文内容.内容行次%Type,
  内容文本_In   In 病历范文内容.内容文本%Type,
  诊治要素id_In In 病历范文内容.诊治要素id%Type := Null,
  替换域_In     In 病历范文内容.替换域%Type := 0,
  要素名称_In   In 病历范文内容.要素名称%Type := Null,
  要素类型_In   In 病历范文内容.要素类型%Type := Null,
  要素长度_In   In 病历范文内容.要素长度%Type := Null,
  要素小数_In   In 病历范文内容.要素小数%Type := Null,
  要素单位_In   In 病历范文内容.要素单位%Type := Null,
  要素表示_In   In 病历范文内容.要素表示%Type := 0,
  输入形态_In   In 病历范文内容.输入形态%Type := 0,
  要素值域_In   In 病历范文内容.要素值域%Type := Null
) Is
Begin

  Update 病历范文内容
  Set 文件id = 文件id_In, 父id = 父id_In, 对象序号 = 对象序号_In, 对象类型 = 对象类型_In, 保留对象 = 保留对象_In, 对象属性 = 对象属性_In, 内容行次 = 内容行次_In,
      内容文本 = 内容文本_In, 诊治要素id = 诊治要素id_In, 替换域 = 替换域_In, 要素名称 = 要素名称_In, 要素类型 = 要素类型_In, 要素长度 = 要素长度_In, 要素小数 = 要素小数_In,
      要素单位 = 要素单位_In, 要素表示 = 要素表示_In, 输入形态 = 输入形态_In, 要素值域 = 要素值域_In
  Where ID = Id_In;
  If Sql%RowCount = 0 Then
    Insert Into 病历范文内容
      (ID, 文件id, 父id, 对象序号, 对象类型, 保留对象, 对象属性, 内容行次, 内容文本, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域)
    Values
      (Id_In, 文件id_In, 父id_In, 对象序号_In, 对象类型_In, 保留对象_In, 对象属性_In, 内容行次_In, 内容文本_In, 诊治要素id_In, 替换域_In, 要素名称_In,
       要素类型_In, 要素长度_In, 要素小数_In, 要素单位_In, 要素表示_In, 输入形态_In, 要素值域_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病历范文内容_Apply;
/

Create Or Replace Procedure Zl_病历范文内容_Delete(文件id_In In 病历范文内容.文件id%Type) Is
Begin

  Delete 病历范文内容 Where 文件id = 文件id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病历范文内容_Delete;
/
Create Or Replace Procedure Zl_电子病历记录_Apply
(
  Id_In       In 电子病历记录.ID%Type,
  病人来源_In In 电子病历记录.病人来源%Type,
  病人id_In   In 电子病历记录.病人id%Type,
  主页id_In   In 电子病历记录.主页id%Type,
  婴儿_In     In 电子病历记录.婴儿%Type,
  科室id_In   In 电子病历记录.科室id%Type,
  病历种类_In In 电子病历记录.病历种类%Type,
  文件id_In   In 电子病历记录.文件id%Type,
  病历名称_In In 电子病历记录.病历名称%Type,
  创建人_In   In 电子病历记录.创建人%Type,
  保存人_In   In 电子病历记录.保存人%Type,
  医嘱id_In   In 病人医嘱报告.医嘱id%Type:=0
) Is
  e_Submit Exception;
  e_Repeat Exception;
  d_保存时间 电子病历记录.保存时间%Type;
  n_处理状态 电子病历记录.处理状态%Type;
  n_Num      Number;
Begin
  d_保存时间 := Sysdate;

  Update 电子病历记录 Set 保存人 = 保存人_In, 保存时间 = d_保存时间 Where ID = Id_In;
  If Sql%RowCount = 0 Then
    Insert Into 电子病历记录
      (ID, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 最后版本, 创建人, 创建时间, 保存人, 保存时间,编辑方式)
    Values
      (Id_In, 病人来源_In, 病人id_In, 主页id_In, 婴儿_In, 科室id_In, 病历种类_In, 文件id_In, 病历名称_In, 1, 创建人_In, d_保存时间, 保存人_In, d_保存时间,1);
    If 病历种类_In = 7 And Nvl(医嘱id_In, 0) <> 0 Then
      Select Count(*)
      Into n_Num
      From 电子病历记录 L, 病人医嘱报告 R
      Where L.ID = R.病历id And R.医嘱id = 医嘱id_In And L.文件id = 文件id_In;
      If n_Num > 0 Then
        Raise e_Repeat;
      End If;
      Insert Into 病人医嘱报告 (医嘱id, 病历id) Values (医嘱id_In, Id_In);
    End If;
  Else
    Select Nvl(处理状态, 0) Into n_处理状态 From 电子病历记录 Where ID = Id_In;
    If n_处理状态 > 0 Then
      Raise e_Submit;
    End If;
  End If;
Exception
  When e_Submit Then
    Raise_Application_Error(-20101, '[ZLSOFT]不能更改被后续接收的病历！[ZLSOFT]');
  When e_Repeat Then
    Raise_Application_Error(-20101, '[ZLSOFT]其他人已经书写并保存了报告，不能再保存！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子病历记录_Apply;
/
Create Or Replace Procedure Zl_电子病历内容_Apply
(
  Id_In         In 电子病历内容.ID%Type,
  文件id_In     In 电子病历内容.文件id%Type,
  父id_In       In 电子病历内容.父id%Type,
  对象序号_In   In 电子病历内容.对象序号%Type,
  对象类型_In   In 电子病历内容.对象类型%Type,
  保留对象_In   In 电子病历内容.保留对象%Type,
  对象属性_In   In 电子病历内容.对象属性%Type,
  内容行次_In   In 电子病历内容.内容行次%Type,
  内容文本_In   In 电子病历内容.内容文本%Type,
  开始版_In     In 电子病历内容.开始版%Type,
  终止版_In     In 电子病历内容.终止版%Type := 0,
  诊治要素id_In In 电子病历内容.诊治要素id%Type := Null,
  替换域_In     In 电子病历内容.替换域%Type := 0,
  要素名称_In   In 电子病历内容.要素名称%Type := Null,
  要素类型_In   In 电子病历内容.要素类型%Type := Null,
  要素长度_In   In 电子病历内容.要素长度%Type := Null,
  要素小数_In   In 电子病历内容.要素小数%Type := Null,
  要素单位_In   In 电子病历内容.要素单位%Type := Null,
  要素表示_In   In 电子病历内容.要素表示%Type := 0,
  输入形态_In   In 电子病历内容.输入形态%Type := 0,
  要素值域_In   In 电子病历内容.要素值域%Type := Null
) Is
  v_保存人   电子病历记录.保存人%Type;
  d_保存时间 电子病历记录.保存时间%Type;
Begin
  Begin
    Select P.姓名 Into v_保存人 From 上机人员表 O, 人员表 P Where O.人员id = P.ID And 用户名 = User;
  Exception
    When Others Then
      v_保存人 := User;
  End;
  d_保存时间 := Sysdate;

  Update 电子病历内容
  Set 文件id = 文件id_In, 父id = 父id_In,对象序号 =对象序号_In, 对象类型 = 对象类型_In, 保留对象 = 保留对象_In, 对象属性 = 对象属性_In, 内容行次 = 内容行次_In, 内容文本 = 内容文本_In,
      诊治要素id = 诊治要素id_In, 替换域 = 替换域_In, 要素名称 = 要素名称_In, 要素类型 = 要素类型_In, 要素长度 = 要素长度_In, 要素小数 = 要素小数_In, 要素单位 = 要素单位_In,
      要素表示 = 要素表示_In, 输入形态 = 输入形态_In, 要素值域 = 要素值域_In, 开始版 = 开始版_In, 终止版 = 终止版_In
  Where ID = Id_In;
  If Sql%RowCount = 0 Then
    Insert Into 电子病历内容
      (ID, 文件id, 父id,对象序号, 对象类型, 保留对象, 对象属性, 内容行次, 内容文本, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 开始版,
       终止版)
    Values
      (Id_In, 文件id_In, 父id_In,对象序号_In, 对象类型_In, 保留对象_In, 对象属性_In, 内容行次_In, 内容文本_In, 诊治要素id_In, 替换域_In, 要素名称_In, 要素类型_In,
       要素长度_In, 要素小数_In, 要素单位_In, 要素表示_In, 输入形态_In, 要素值域_In, 开始版_In, 终止版_In);
  End If;

  If 终止版_In <> 0 And 对象类型_In In (6, 7, 8) Then
    --更新签名级别及完成时间
    Update 电子病历记录
    Set 保存人 = v_保存人, 完成时间 = d_保存时间, 保存时间 = d_保存时间, 最后版本 = 终止版_In,
        签名级别 = (Select Nvl(Sum(Power(2, 要素表示 - 1)), 0)
                 From (Select Distinct 要素表示 From 电子病历内容 Where ID = 文件id_In And 对象类型 In (6, 7, 8)))
    Where ID = 文件id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子病历内容_Apply;
/

CREATE OR REPLACE Procedure Zl_电子病历对象_Endvar
(
  Id_In     In 电子病历内容.ID%Type,
  终止版_In In 电子病历内容.终止版%Type
) Is
--用于终止某个版本
Begin
  Update 电子病历内容 Set 终止版 = 终止版_In Where ID = Id_In Or 父id = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子病历对象_Endvar;
/

Create Or Replace Procedure Zl_电子病历对象_Delete
(
  Id_In       In 电子病历内容.ID%Type,
  父id_In     In 电子病历内容.父id%Type,
  对象类型_In In 电子病历内容.对象类型%Type
) Is
Begin
  If 父id_In = 0 Then
    Delete 电子病历内容 Where ID = Id_In And 对象类型 = 对象类型_In;
  Else
    Delete 电子病历内容 Where 父id = 父id_In And 对象类型 = 对象类型_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子病历对象_Delete;
/
Create Or Replace Procedure Zl_电子病历内容_Untread
(
  文件id_In In 电子病历内容.文件id%Type,
  终止版_In In 电子病历内容.终止版%Type,
  Rollsign  In Number
) Is
  v_Signtime Varchar2(100);
  v_Savename Varchar2(64);
Begin
  If Rollsign = 1 Then
    --撤消签名
    If 终止版_In > 2 Then
      --撤消第二次以其之后的签名
      Delete 电子病历内容 Where 文件id = 文件id_In And 对象类型 In (6, 7, 8) And 终止版 = 终止版_In;
      --最后一次签名的签名时间和签名人用于更新电子病历记录
      Select 对象属性, 内容文本
      Into v_Signtime, v_Savename
      From 电子病历内容
      Where 文件id = 文件id_In And 对象类型 In (6, 7, 8) And 终止版 = 终止版_In - 1;
      v_Signtime := Substr(v_Signtime, Instr(v_Signtime, '|') + 1);
      v_Signtime := Substr(v_Signtime, Instr(v_Signtime, ';', 1, 4) + 1);
      v_Signtime := Substr(v_Signtime, 1, Instr(v_Signtime, ';', 1, 1) - 1);
      Update 电子病历记录
      Set 最后版本 = 最后版本 - 1, 完成时间 = To_Date(v_Signtime, 'yyyy-mm-dd HH24:mi:ss'), 保存人 = v_Savename, 保存时间 = Sysdate;
    Else
      --撤消第一次签名
      Update 电子病历内容
      Set 终止版 = 0, 对象属性 = Substr(对象属性, 1, Instr(对象属性, '|')), 内容文本 = Null, 要素名称 = Null, 要素单位 = Null, 要素表示 = 0,
          要素值域 = Null
      Where 文件id = 文件id_In And 对象类型 In (6, 7, 8) And 终止版 = 终止版_In;
      Update 电子病历记录 Set 最后版本 = 1, 完成时间 = Null, 保存人 = Null, 保存时间 = Sysdate Where ID = 文件id_In;
    End If;
  Else
    --撤消修订
    Delete 电子病历内容 Where 文件id = 文件id_In And 开始版 = 终止版_In;
    Update 电子病历内容 Set 终止版 = 0 Where 文件id = 文件id_In And 终止版 = 终止版_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子病历内容_Untread;
/
Create Or Replace Function f_Sentence_Usable
(
  提纲id_In   In 病历提纲词句.提纲id%Type := 0,
  病人id_In   In 电子病历记录.病人id%Type,
  主页id_In   In 电子病历记录.主页id%Type,
  医嘱id_In   In 病人医嘱记录.ID%Type := Null,
  Filetype_In In Number := 0
) Return t_Dic_Rowset Is
  a_Id_Retu t_Dic_Rowset := t_Dic_Rowset();

  n_符合 Number(18);
  Type t_Word_Table Is Table Of 病历词句示范.ID%Type;
  a_Words t_Word_Table := t_Word_Table();

  Type t_Term_Record Is Record(
    条件项 Varchar2(20),
    条件值 Varchar2(2000));
  Type t_Term_Table Is Table Of t_Term_Record;
  a_Terms t_Term_Table := t_Term_Table();

  r_Patient 病人信息%RowType := Null;
  r_In_Page 病案主页%RowType := Null;
  r_Advice  病人医嘱记录%RowType := Null;
  v_Value   Varchar2(2000);

Begin
  Begin
    Select * Into r_Patient From 病人信息 Where 病人id = 病人id_In;
    Select * Into r_In_Page From 病案主页 Where 病人id = 病人id_In And 主页id = 主页id_In;
  Exception
    When Others Then
      Null;
  End;
  Begin
    Select * Into r_Advice From 病人医嘱记录 Where ID = 医嘱id_In And 病人id + 0 = 病人id_In;
  Exception
    When Others Then
      Null;
  End;

  If 提纲id_In = 0 Then
    Select A.ID Bulk Collect
    Into a_Words
    From 病历词句示范 A, 病历词句分类 B
    Where Substr(范围, Filetype_In, 1) = 1 And A.分类id = B.ID;
  Else
    Select L.ID Bulk Collect
    Into a_Words
    From 病历词句示范 L, 病历提纲词句 A
    Where L.分类id = A.词句分类id And A.提纲id = 提纲id_In;
  End If;

  For n_Wcount In 1 .. a_Words.Count Loop
    n_符合  := 1;
    a_Terms := t_Term_Table();
    For r_Temp In (Select 条件项, 条件值 From 病历词句条件 Where 词句id = a_Words(n_Wcount) And 条件值 Is Not Null) Loop
      a_Terms.Extend;
      a_Terms(a_Terms.Count).条件项 := r_Temp.条件项;
      a_Terms(a_Terms.Count).条件值 := Chr(9) || r_Temp.条件值 || Chr(9);
    End Loop;
    For n_Tcount In 1 .. a_Terms.Count Loop
      If r_Patient.病人id Is Not Null Then
        If a_Terms(n_Tcount).条件项 = '病人性别' Then
          n_符合 := Instr(a_Terms(n_Tcount).条件值, Chr(9) || r_Patient.性别 || Chr(9));
        Elsif a_Terms(n_Tcount).条件项 = '婚姻状况' Then
          n_符合 := Instr(a_Terms(n_Tcount).条件值, Chr(9) || r_Patient.婚姻状况 || Chr(9));
        End If;
      End If;
      If r_In_Page.主页id Is Not Null Then
        If a_Terms(n_Tcount).条件项 = '住院目的' Then
          n_符合 := Instr(a_Terms(n_Tcount).条件值, Chr(9) || r_In_Page.住院目的 || Chr(9));
        Elsif a_Terms(n_Tcount).条件项 = '病人病情' Then
          n_符合 := Instr(a_Terms(n_Tcount).条件值, Chr(9) || r_In_Page.当前病况 || Chr(9));
        Elsif a_Terms(n_Tcount).条件项 = '入院方式' Then
          n_符合 := Instr(a_Terms(n_Tcount).条件值, Chr(9) || r_In_Page.入院方式 || Chr(9));
        End If;
      End If;
      If r_Advice.ID Is Not Null Then
        If a_Terms(n_Tcount).条件项 = '诊疗类别' Then
          Select 名称 Into v_Value From 诊疗项目类别 Where 编码 = r_Advice.诊疗类别;
          n_符合 := Instr(a_Terms(n_Tcount).条件值, Chr(9) || v_Value || Chr(9));
        Elsif a_Terms(n_Tcount).条件项 = '检查类型' Then
          Select 操作类型 Into v_Value From 诊疗项目目录 I Where ID = r_Advice.诊疗项目id;
          n_符合 := Instr(a_Terms(n_Tcount).条件值, Chr(9) || v_Value || Chr(9));
        Elsif a_Terms(n_Tcount).条件项 = '检查部位' Then
          n_符合 := 0;
          For r_Temp In (Select Distinct 标本部位 From 病人医嘱记录 Where ID = r_Advice.ID Or 相关id = r_Advice.ID) Loop
            If Instr(a_Terms(n_Tcount).条件值, Chr(9) || r_Temp.标本部位 || Chr(9)) <> 0 Then
              n_符合 := 1;
              Exit;
            End If;
          End Loop;
        Elsif a_Terms(n_Tcount).条件项 = '检查方法' Then
          n_符合 := 0;
          For r_Temp In (Select Distinct 检查方法 From 病人医嘱记录 Where ID = r_Advice.ID Or 相关id = r_Advice.ID) Loop
            If Instr(a_Terms(n_Tcount).条件值, Chr(9) || r_Temp.检查方法 || Chr(9)) <> 0 Then
              n_符合 := 1;
              Exit;
            End If;
          End Loop;
        End If;
      End If;
      If n_符合 = 0 Then
        Exit;
      End If;
    End Loop;
    If n_符合 <> 0 Then
      a_Id_Retu.Extend;
      a_Id_Retu(a_Id_Retu.Count) := t_Dic_Record(Null, Null, Null);
      a_Id_Retu(a_Id_Retu.Count).编码 := a_Words(n_Wcount);
    End If;
  End Loop;
  Return a_Id_Retu;
Exception
  When Others Then
    Return a_Id_Retu;
End f_Sentence_Usable;
/