----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--123137:焦博,2018-03-21,在保险支付项目表中增加字段保险费用等级
alter table 保险支付项目 add 保险费用等级  varchar2(50);




------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--123386:刘硕,2018-03-23,收费价目与收费对照锚点
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '字典', 'ZLHIS_DICT_053', '收费价目变动', '<root><ID></ID><变动类型></变动类型></root>', '收费细目管理:调价时'  From Dual Union All 
Select '字典', 'ZLHIS_DICT_054', '诊疗收费对照变动', '<root><ID></ID><原对照></原对照><现对照></现对照></root>', '诊疗项目管理:设置诊疗收费对照时'  From Dual;

--123263:王振涛,2018-03-22,集成平台消息锚点
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '字典', 'ZLHIS_DICTLIS_004', '新增诊疗标本类型', '<root><编码></编码><名称></名称><简码><简码/><适用性别></适用性别></root>', '字典管理工具:新增诊疗标本类型'  From Dual Union All 
Select '字典', 'ZLHIS_DICTLIS_005', '修改诊疗标本类型', '<root><编码></编码><名称></名称><简码><简码/><适用性别></适用性别></root>', '字典管理工具:修改诊疗标本类型'  From Dual Union All 
Select '字典', 'ZLHIS_DICTLIS_006', '删除诊疗标本类型', '<root><编码></编码><名称></名称><简码><简码/><适用性别></适用性别></root>', '字典管理工具:删除诊疗标本类型'  From Dual Union All 
Select '字典', 'ZLHIS_DICTLIS_007', '新增采血管', '<root><编码></编码><名称></名称><简码></简码><添加剂></添加剂><采血量></采血量><规格></规格><颜色></颜色><材料ID></材料ID><root>', '检验采血管设置:新增采血管'  From Dual Union All 
Select '字典', 'ZLHIS_DICTLIS_008', '修改采血管', '<root><编码></编码><名称></名称><简码></简码><添加剂></添加剂><采血量></采血量><规格></规格><颜色></颜色><材料ID></材料ID><root>', '检验采血管设置:修改采血管'  From Dual Union All 
Select '字典', 'ZLHIS_DICTLIS_009', '删除采血管', '<root><编码></编码><名称></名称><简码></简码><添加剂></添加剂><采血量></采血量><规格></规格><颜色></颜色><材料ID></材料ID></root>', '检验采血管设置:删除采血管'  From Dual ;


--122998:胡俊勇,2018-03-21,集成平台消息锚点
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '字典', 'ZLHIS_DICTPACS_001', '新增诊疗检查类型', '<root><编码></编码><名称></名称><简码><简码/><建病案></建病案></root>', '字典管理工具:新增诊疗检查类型'  From Dual Union All 
Select '字典', 'ZLHIS_DICTPACS_002', '修改诊疗检查类型', '<root><编码></编码><名称></名称><简码><简码/><建病案></建病案></root>', '字典管理工具:修改诊疗检查类型'  From Dual Union All 
Select '字典', 'ZLHIS_DICTPACS_003', '删除诊疗检查类型', '<root><编码></编码><名称></名称><简码><简码/><建病案></建病案></root>', '字典管理工具:删除诊疗检查类型'  From Dual Union All 
Select '字典', 'ZLHIS_DICTPACS_004', '新增诊疗检查部位', '<root><类型></类型><编码></编码><名称></名称><分组></分组><备注></备注><方法></方法><适用性别></适用性别><root>', '检查部位设置:新增诊疗检查部位'  From Dual Union All 
Select '字典', 'ZLHIS_DICTPACS_005', '修改诊疗检查部位', '<root><类型></类型><编码></编码><名称></名称><分组></分组><备注></备注><方法></方法><适用性别></适用性别><root>', '检查部位设置:修改诊疗检查部位'  From Dual Union All 
Select '字典', 'ZLHIS_DICTPACS_006', '删除诊疗检查部位', '<root><类型></类型><编码></编码><名称></名称><分组></分组><备注></备注><方法></方法><适用性别></适用性别></root>', '检查部位设置:删除诊疗检查部位'  From Dual Union All 
Select '字典', 'ZLHIS_DICTPACS_007', '新增诊疗项目部位', '<root><ID></ID><项目ID></项目ID><类型></类型><部位></部位><方法></方法><默认></默认></root>', '诊疗项目设置:新增诊疗项目部位'  From Dual Union All 
Select '字典', 'ZLHIS_DICTPACS_008', '修改诊疗项目部位', '<root><ID></ID><项目ID></项目ID><类型></类型><部位></部位><方法></方法><默认></默认></root>', '诊疗项目设置:修改诊疗项目部位'  From Dual Union All 
Select '字典', 'ZLHIS_DICTPACS_009', '删除诊疗项目部位', '<root><ID></ID><项目ID></项目ID><类型></类型><部位></部位><方法></方法><默认></默认></root>', '诊疗项目设置:删除诊疗项目部位'  From Dual;


--123312:梁唐彬,2018-03-22,新病理系统集成平台消息
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '临床', 'ZLHIS_CIS_056', '病理申请发送后修改', '<root><病人ID></病人ID><主页ID></主页ID><挂号单></挂号单><发送号></发送号><ID></ID><病人来源></病人来源></root>', '医技站修改病理申请单的标本信息时'  From Dual;


--123098:李南春,2018-03-20,发卡或绑定卡是否自动生成门诊号
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1107, 0, 0, 0, 0, 0, 0, 26, '自动门诊号', NULL, '1',
         '启用了此参数，在发卡或绑定卡时将会为没有门诊号的病人自动生成门诊号', '0-不自动生成门诊号,1-自动生成门诊号', NULL, '为病人自动生成门诊号，便于门诊号有序管理', Null
  From Dual;

Update zlParameters Set 参数值 = 1 where 系统 = &n_System And 模块 = 1107 And 参数名 = '自动门诊号' And Exists(Select 1　From zlParameters where 系统 = &n_System And 模块= 1111 And 参数名 = '自动门诊号' And 参数值 = 1);





-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--123120:蒋廷中,2018-03-21,拆分修改打包状态和配液批次权限
Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明, 缺省值)
  Select &n_System, 1254, '修改配液打包状态', 35, '临床护士有此权限时可以修改配液记录的打包状态，否则不能修改。', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where 系统 = &n_System And 序号 = 1254 And 功能 = '修改配液打包状态');

Insert Into zlProgFuncs
  (系统, 序号, 功能, 排列, 说明, 缺省值)
  Select &n_System, 1254, '修改配液批次', 36, '临床护士有此权限时可以修改配液记录的批次，否则不能修改。', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where 系统 = &n_System And 序号 = 1254 And 功能 = '修改配液批次');

Insert Into zlRoleGrant
  Select &n_System, 1254, b.角色, '修改配液打包状态'
  From zlRoleGrant B
  Where b.系统 = &n_System And b.序号 = 1254 And b.功能 = '修改打包状态和配液批次' And Not Exists
   (Select 1
         From zlRoleGrant C
         Where c.系统 = &n_System And c.序号 = 1254 And c.角色 = b.角色 And c.功能 = '修改配液打包状态');

Insert Into zlRoleGrant
  Select &n_System, 1254, b.角色, '修改配液批次'
  From zlRoleGrant B
  Where b.系统 = &n_System And b.序号 = 1254 And b.功能 = '修改打包状态和配液批次' And Not Exists
   (Select 1
         From zlRoleGrant C
         Where c.系统 = &n_System And c.序号 = 1254 And c.角色 = b.角色 And c.功能 = '修改配液批次');
         
Delete From zlProgFuncs Where 系统 = &n_System And 序号 = 1254 And 功能 = '修改打包状态和配液批次';






-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
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
  v_Old        Varchar2(4000);
Begin
  Update 诊疗项目目录 Set 计价性质 = 计价性质_In Where ID = 诊疗项目id_In;
  If 是否删除_In = 1 Then
    If Nvl(适用科室id_In, 0) = 0 And Nvl(病人来源_In, 0) = 0 Then
      Select f_List2str(Cast(Collect(收费项目id || '^' || 收费数量 || '^' || 固有对照 || '^' || 从属项目 || '^' || 费用性质 || '^' || 检查部位 || '^' || 检查方法 || '^' || 收费方式) As
                              t_Strlist), '|')
      Into v_Old
      From 诊疗收费关系
      Where 诊疗项目id = 诊疗项目id_In;
    End If;
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
    b_Message.Zlhis_Dict_054(诊疗项目id_In, v_Old, 收费内容_In);
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
  n_State Number(1);
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
    n_State := 1;
  Else
    n_State := 2;
  End If;

  If 价格等级_In Is Null Then
    b_Message.Zlhis_Dict_053(收费细目id_In, n_State);
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
    b_Message.Zlhis_Dict_053(收费细目id_In, 0);
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
    b_Message.Zlhis_Dict_053(收费细目id_In, 1);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_收费价目_Insert;
/

--123263:王振涛,2018-03-22,集成平台消息锚点
Create Or Replace Procedure Zl_采血管类型_Clear
(
  Type_In  In Number := 0,
  Oldno_In In 采血管类型.编码%Type := Null,
  Newno_In In 采血管类型.编码%Type := Null
) Is
Begin
  If Type_In = 0 Then
    --- 为保持兼容性 
    Delete 采血管类型;
  End If;

  If Type_In = 1 Then
    -- 改编码 
    If Nvl(Oldno_In, 0) <> 0 Then
      If Nvl(Newno_In, 0) <> 0 Then
        Update 采血管类型 Set 编码 = Newno_In Where 编码 = Oldno_In;
      Else
        For R In (Select 编码, 名称, 简码, 规格, 添加剂, 采血量, 颜色, 材料id From 采血管类型 A Where a.编码 = Oldno_In) Loop
          b_Message.Zlhis_Dictlis_009(r.编码, r.名称, r.简码, r.规格, r.添加剂, r.采血量, r.颜色, r.材料id);
        End Loop;
        Delete 采血管类型 Where 编码 = Oldno_In;
      End If;    
      Update 诊疗项目目录 Set 试管编码 = Newno_In Where 试管编码 = Oldno_In;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_采血管类型_Clear;
/

--123263:王振涛,2018-03-22,集成平台消息锚点
Create Or Replace Procedure Zl_采血管类型_Update
(
  编码_In   In 采血管类型.编码%Type,
  名称_In   In 采血管类型.名称%Type,
  规格_In   In 采血管类型.规格%Type,
  添加剂_In In 采血管类型.添加剂%Type,
  采血量_In In 采血管类型.采血量%Type,
  颜色_In   In 采血管类型.颜色%Type,
  材料id_In In 采血管类型.材料id%Type := Null
) Is
  v_材料id Number;
Begin
  If Nvl(材料id_In, 0) <> 0 Then
    v_材料id := 材料id_In;
  Else
    v_材料id := Null;
  End If;
  Update 采血管类型
  Set 名称 = 名称_In, 规格 = 规格_In, 添加剂 = 添加剂_In, 采血量 = 采血量_In, 颜色 = 颜色_In, 材料id = v_材料id
  Where 编码 = 编码_In;
  
  If Sql%NotFound Then
    Insert Into 采血管类型
      (编码, 名称, 规格, 添加剂, 采血量, 颜色, 材料id)
    Values
      (编码_In, 名称_In, 规格_In, 添加剂_In, 采血量_In, 颜色_In, v_材料id);
    b_Message.Zlhis_Dictlis_007(编码_In, 名称_In, Null, 规格_In, 添加剂_In, 采血量_In, 颜色_In, v_材料id);
  else
    b_Message.Zlhis_Dictlis_008(编码_In, 名称_In, Null, 规格_In, 添加剂_In, 采血量_In, 颜色_In, v_材料id);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_采血管类型_Update;
/

--122937:胡俊勇,2018-03-21,护理接口修改
Create Or Replace Procedure Zl_Third_Getadviceinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：获取医嘱基本信息/查询
  --参数：
  --入参数 Xml_In
  --<IN>
  --     <YZID>1156789</YZID>--主医嘱ID
  --</IN>

  --出参 Xml_Out
  --<OUTPUT>
  --    <YZ>
  --       <PATIID></PATIID>     --病人医嘱记录.病人ID
  --       <PAGEID></PAGEID>     --病人医嘱记录.主页ID
  --       <BABY></BABY>   --病人医嘱记录.婴儿
  --       <YZID>1145878</YZID>   --病人医嘱记录.医嘱ID， 主医嘱ID
  --       <RELATEDID></RELATEDID>   --病人医嘱记录.相关ID
  --       <ZXKSID></ZXKSID>   --病人医嘱记录.执行科室id
  --       <YZQX>0</YZQX>      --病人医嘱记录.医嘱期效
  --       <STATE>8</STATE>    --病人医嘱记录.医嘱状态
  --       <JJBZ>0</JJBZ>      --病人医嘱记录.紧急标志
  --       <KZYS>代翔</KZYS>   --病人医嘱记录.开嘱医生
  --       <KZSJ>2015-03-25 16:37:00</KZSJ>   --病人医嘱记录.开嘱时间
  --       <ZLXMID></ZLXMID>   --诊疗项目目录.ID
  --       <ZLLB>E</ZLLB>      --诊疗项目目录.类别
  --       <ZLXMMC></ZLXMMC>   --诊疗项目目录.名称 ，检查，检验(检验行 C)，手术(主手术行 F)，输血(K)，中药配方(服法行 E)，其它(本身)
  --       <ZLXMCZLX></<ZLXMCZLX>   --诊疗项目目录.操作类型
  --       <ZLXMZXFL></ZLXMZXFL>   --诊疗项目目录.执行分类
  --       <BZ>21</BZ> 诊疗项目目录.操作类型||诊疗项目目录.执行分类
  --       <YF>静脉滴注</YF>   --病人医嘱记录.医嘱内容 ，主医嘱行中的  医嘱内容
  --       <PC>BID</PC>   --诊疗频率项目.英文名称
  --       <ZXSJFY>18-20</ZXSJFY>   --病人医嘱记录.执行时间方案
  --       <PLCS>2</PLCS>   --病人医嘱记录.频率次数
  --       <PLJG>1</PLJG>   --病人医嘱记录.频率间隔
  --       <PSJG></PSJG>   --病人医嘱记录.皮试结果
  --       <YSZT></YSZT>   --病人医嘱记录.医生嘱托
  --       <KSZXSJ>2015-03-25 16:35:00</KSZXSJ>  --病人医嘱记录.开始执行时间
  --       <ZXZZSJ></ZXZZSJ>   --病人医嘱记录.执行终止时间
  --       <TZYS></TZYS>   --病人医嘱记录.停嘱医生
  --       <TZSJ></TZSJ>   --病人医嘱记录.停嘱时间
  --       <DW>次</DW>   --诊疗项目目录.计算单位
  --       <DL></DL>   --病人医嘱记录.单次用量
  --       <ZL></ZL>   --病人医嘱记录.总给予量

  --       <ITEMLIST> 仅输血项目和西/成药医嘱项目明细相关信息；输血的血袋信息，药品行明细信息
  --        <ITEM>
  --         <YSZT></YSZT>   --病人医嘱记录.医生嘱托
  --         <YZID>1145878</YZID>   --病人医嘱记录.医嘱ID
  --         <RELATEDID></RELATEDID>   --病人医嘱记录.相关ID
  --         <ZLXMID></ZLXMID>   --诊疗项目目录.ID
  --         <SFXMID></SFXMID>   --收费项目目录.id
  --         <SFXMMC></SFXMMC>   --收费项目目录.名称
  --         <SFXMGG></SFXMGG>   --收费项目目录.规格
  --         <BM></BM>           --收费项目别名.名称（商品名）
  --         <ZL></ZL>           --病人医嘱记录.总给予量
  --         <DL>10</DL>         --病人医嘱记录.单次用量
  --         <DW>ml</DW>         --收费项目目录.计算单位
  --         <ZLDW>ml</ZLDW>   --诊疗项目目录.计算单位
  --         <ZXXZ></ZXXZ>   --病人医嘱记录.执行性质
  --         <ZXKS></ZXKS>   --诊疗项目目录.执行科室
  --         <XDBH></XDBH>   --血液收发记录.血袋编号
  --         <SXXH></SXXH>   --血液收发记录.序号
  --        </ITEM>
  --        <ITEM/>...
  --       </ITEMLIST>
  --      </YZ>
  --</OUTPUT>

  n_医嘱id  病人医嘱记录.Id%Type;
  x_医嘱    Xmltype;
  x_Item    Xmltype;
  v_Xtmp    Clob; --临时XML
  n_Cnt     Number;
  x_Templet Xmltype;

  v_英文名     诊疗频率项目.英文名称%Type;
  v_试管名称   采血管类型.名称%Type;
  v_添加剂     采血管类型.添加剂%Type;
  v_试管规格   采血管类型.规格%Type;
  n_试管颜色   采血管类型.颜色%Type;
  v_收费商品名 收费项目别名.名称%Type;
  n_启用血库   Number;
  v_Sql血库    Varchar2(4000);
  n_血库申请id Number(18);
  v_Tmp输血    Varchar2(4000);

  Type Bloodlist_Type Is Ref Cursor;
  Cbloodlist Bloodlist_Type;

  Type t_Code Is Record(
    ID       收费项目目录.Id%Type,
    名称     收费项目目录.名称%Type,
    规格     收费项目目录.规格%Type,
    单位     收费项目目录.计算单位%Type,
    血袋编号 Varchar2(50),
    序号     Number(5));
  r_b t_Code;

Begin

  Select Extractvalue(Value(A), 'IN/YZID') Into n_医嘱id From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_Cnt := 0;
  For R In (Select a.病人id, a.主页id, a.婴儿, a.Id As 医嘱id, a.相关id, a.执行科室id, a.医嘱期效, a.医嘱状态, a.紧急标志, a.开嘱医生, a.开嘱时间, a.诊疗项目id,
                   a.诊疗类别, a.医嘱内容, a.执行时间方案, a.执行频次, a.频率次数, a.频率间隔, a.皮试结果, a.医生嘱托, a.开始执行时间, a.执行终止时间, a.停嘱医生, a.停嘱时间,
                   b.名称 As 项目名称, b.操作类型, b.执行分类, b.计算单位 As 诊疗单位, a.单次用量, a.总给予量, a.标本部位, a.检查方法, a.收费细目id, c.名称 As 收费名称,
                   c.规格, Null As 收费商品名, c.计算单位 As 收费单位, a.执行性质, b.执行科室, b.试管编码, c.产地, d.高危药品
            From 病人医嘱记录 A, 诊疗项目目录 B, 收费项目目录 C, 药品规格 D
            Where a.诊疗项目id = b.Id And a.收费细目id = c.Id(+) And a.收费细目id = d.药品id(+) And (a.Id = n_医嘱id Or a.相关id = n_医嘱id)
            Order By a.序号) Loop
    n_Cnt := n_Cnt + 1;
    If n_Cnt = 1 Then
      Select Max(a.英文名称) Into v_英文名 From 诊疗频率项目 A Where a.名称 = r.执行频次;
    End If;
    v_试管名称 := Null;
    v_添加剂   := Null;
    v_试管规格 := Null;
    n_试管颜色 := Null;
    If r.试管编码 Is Not Null Then
      Select Max(a.名称), Max(a.添加剂), Max(a.规格), Max(a.颜色)
      Into v_试管名称, v_添加剂, v_试管规格, n_试管颜色
      From 采血管类型 A
      Where a.编码 = r.试管编码;
    End If;
    --主医行
    If r.相关id Is Null Then
      v_Xtmp := '<YZ>';
      v_Xtmp := v_Xtmp || '<PATIID>' || r.病人id || '</PATIID>'; --病人医嘱记录.病人ID
      v_Xtmp := v_Xtmp || '<PAGEID>' || r.主页id || '</PAGEID>'; --病人医嘱记录.主页ID
      v_Xtmp := v_Xtmp || '<BABY>' || r.婴儿 || '</BABY>'; --病人医嘱记录.婴儿
      v_Xtmp := v_Xtmp || '<YZID>' || r.医嘱id || '</YZID>'; --病人医嘱记录.医嘱ID
      v_Xtmp := v_Xtmp || '<RELATEDID>' || r.相关id || '</RELATEDID>'; --病人医嘱记录.相关ID
      v_Xtmp := v_Xtmp || '<ZXKSID>' || r.执行科室id || '</ZXKSID>'; --病人医嘱记录.执行科室id
      v_Xtmp := v_Xtmp || '<YZQX>' || r.医嘱期效 || '</YZQX>'; --病人医嘱记录.医嘱期效
      v_Xtmp := v_Xtmp || '<STATE>' || r.医嘱状态 || '</STATE>'; --病人医嘱记录.医嘱状态
      v_Xtmp := v_Xtmp || '<JJBZ>' || r.紧急标志 || '</JJBZ>'; --病人医嘱记录.紧急标志
      v_Xtmp := v_Xtmp || '<KZYS>' || r.开嘱医生 || '</KZYS>'; --病人医嘱记录.开嘱医生
      v_Xtmp := v_Xtmp || '<KZSJ>' || To_Char(r.开嘱时间, 'yyyy-mm-dd hh24:mi:ss') || '</KZSJ>'; --病人医嘱记录.开嘱时间
      v_Xtmp := v_Xtmp || '<BZ>' || r.操作类型 || r.执行分类 || '</BZ>'; -- 诊疗项目目录.操作类型||诊疗项目目录.执行分类
      v_Xtmp := v_Xtmp || '<ZLXMID>' || r.诊疗项目id || '</ZLXMID>'; --诊疗项目目录.ID
      v_Xtmp := v_Xtmp || '<ZLLB>' || r.诊疗类别 || '</ZLLB>'; --诊疗项目目录.类别
      v_Xtmp := v_Xtmp || '<YZNR>' || r.医嘱内容 || '</YZNR>'; --医嘱内容
      v_Xtmp := v_Xtmp || '<YF>' || r.项目名称 || '</YF>'; --病人医嘱记录.医嘱内容
      v_Xtmp := v_Xtmp || '<PC>' || v_英文名 || '</PC>'; --诊疗频率项目.英文名称
      v_Xtmp := v_Xtmp || '<ZXSJFY>' || r.执行时间方案 || '</ZXSJFY>'; --病人医嘱记录.执行时间方案
      v_Xtmp := v_Xtmp || '<PLCS>' || r.频率次数 || '</PLCS>'; --病人医嘱记录.频率次数
      v_Xtmp := v_Xtmp || '<PLJG>' || r.频率间隔 || '</PLJG>'; --病人医嘱记录.频率间隔
      v_Xtmp := v_Xtmp || '<PSJG>' || r.皮试结果 || '</PSJG>'; --病人医嘱记录.皮试结果
      v_Xtmp := v_Xtmp || '<YSZT>' || r.医生嘱托 || '</YSZT>'; --病人医嘱记录.医生嘱托
      v_Xtmp := v_Xtmp || '<KSZXSJ>' || To_Char(r.开始执行时间, 'yyyy-mm-dd hh24:mi:ss') || '</KSZXSJ>'; --病人医嘱记录.开始执行时间
      v_Xtmp := v_Xtmp || '<ZXZZSJ>' || To_Char(r.执行终止时间, 'yyyy-mm-dd hh24:mi:ss') || '</ZXZZSJ>'; --病人医嘱记录.执行终止时间
      v_Xtmp := v_Xtmp || '<TZYS>' || r.停嘱医生 || '</TZYS>'; --病人医嘱记录.停嘱医生
      v_Xtmp := v_Xtmp || '<TZSJ>' || To_Char(r.停嘱时间, 'yyyy-mm-dd hh24:mi:ss') || '</TZSJ>'; --病人医嘱记录.停嘱时间
      v_Xtmp := v_Xtmp || '<ZLXMMC>' || r.项目名称 || '</ZLXMMC>'; --诊疗项目目录.名称
      v_Xtmp := v_Xtmp || '<ZLXMCZLX>' || r.操作类型 || '</ZLXMCZLX>'; --诊疗项目目录.操作类型
      v_Xtmp := v_Xtmp || '<ZLXMZXFL>' || r.执行分类 || '</ZLXMZXFL>'; --诊疗项目目录.执行分类
      --       (仅采血管返回)
      v_Xtmp := v_Xtmp || '<CXGMC>' || v_试管名称 || '</CXGMC>'; --采血管名称
      v_Xtmp := v_Xtmp || '<CXGTJJ>' || v_添加剂 || '</CXGTJJ>'; --采血管添加剂
      v_Xtmp := v_Xtmp || '<CXGGG>' || v_试管规格 || '</CXGGG>'; --采血管规格
      v_Xtmp := v_Xtmp || '<CXGYS>' || n_试管颜色 || '</CXGYS>'; --采血管颜色
      v_Xtmp := v_Xtmp || '<DW>' || r.诊疗单位 || '</DW>'; --诊疗项目目录.计算单位
      v_Xtmp := v_Xtmp || '<DL>' || r.单次用量 || '</DL>'; --病人医嘱记录.单次用量
      v_Xtmp := v_Xtmp || '<ZL>' || r.总给予量 || '</ZL>'; --病人医嘱记录.总给予量
      v_Xtmp := v_Xtmp || '</YZ>';
      x_医嘱 := Xmltype(v_Xtmp);
    End If;
  
    --输血
    If r.诊疗类别 = 'K' Then
      --判断是否安装血库
      Select Zl_Checkobject(1, '血液收发记录') Into n_启用血库 From Dual;
      If n_启用血库 > 0 Then
        n_血库申请id := r.医嘱id;
        --医嘱部分
        v_Xtmp    := '<YSZT>' || r.医生嘱托 || '</YSZT>'; --病人医嘱记录.医生嘱托
        v_Xtmp    := v_Xtmp || '<YZID>' || r.医嘱id || '</YZID>'; --病人医嘱记录.医嘱ID
        v_Xtmp    := v_Xtmp || '<RELATEDID>' || r.相关id || '</RELATEDID>'; --病人医嘱记录.相关ID
        v_Xtmp    := v_Xtmp || '<ZLXMID>' || r.诊疗项目id || '</ZLXMID>'; --诊疗项目目录.ID
        v_Xtmp    := v_Xtmp || '<ZL>' || r.总给予量 || '</ZL>'; --病人医嘱记录.总给予量
        v_Xtmp    := v_Xtmp || '<DL>' || r.单次用量 || '</DL>'; --病人医嘱记录.单次用量
        v_Xtmp    := v_Xtmp || '<ZLDW>' || r.诊疗单位 || '</ZLDW>'; --诊疗项目目录.计算单位
        v_Xtmp    := v_Xtmp || '<ZXXZ>' || r.执行性质 || '</ZXXZ>'; --病人医嘱记录.执行性质
        v_Xtmp    := v_Xtmp || '<ZXKS>' || r.执行科室 || '</ZXKS>'; --诊疗项目目录.执行科室
        v_Tmp输血 := v_Xtmp;
        If r.检查方法 = '1' Then
          v_Sql血库 := 'Select d.Id,d.名称,d.规格,d.计算单位 as 单位, a.血袋编号,a.序号
                       From 血液收发记录 a,血液发送记录 b,血液配血记录 c,收费项目目录 d
                       Where a.Id = b.收发id And b.配发id = c.Id and a.血液id =d.id  And c.申请id =:1';
        End If;
      End If;
    Elsif r.相关id Is Not Null And r.诊疗类别 = 'E' And r.操作类型 = '8' And Nvl(r.执行分类, 0) = 0 And n_启用血库 = 1 And
          v_Sql血库 Is Null Then
      v_Sql血库 := 'Select b.Id,b.名称,  b.规格,b.计算单位 as 单位, a.血袋编号,a.序号
                  From 血液收发记录 a,收费项目目录 b
                  Where a.血液id =b.id and a.配发id = (Select Id From 血液配血记录 Where 申请id=:1)';
    Else
      v_Sql血库 := Null;
    End If;
  
    If v_Sql血库 Is Not Null And n_血库申请id Is Not Null Then
      --输血医嘱，只有发医嘱后才可能有血袋信息
      x_Item := Xmltype('<ITEMLIST></ITEMLIST>');
      Open Cbloodlist For v_Sql血库
        Using n_血库申请id;
      Loop
        Fetch Cbloodlist
          Into r_b.Id, r_b.名称, r_b.规格, r_b.单位, r_b.血袋编号, r_b.序号;
        Exit When Cbloodlist%NotFound;
        v_收费商品名 := Null;
        If r_b.Id Is Not Null Then
          For Z In (Select a.名称, a.性质
                    From 收费项目别名 A
                    Where a.收费细目id = r_b.Id
                    Group By a.名称, a.性质
                    Order By a.性质) Loop
            v_收费商品名 := z.名称;
            If z.性质 = 3 Then
              v_收费商品名 := z.名称;
              Exit;
            End If;
          End Loop;
        End If;
      
        v_Xtmp := '<ITEM jsonArray="True" >';
      
        v_Xtmp := v_Xtmp || v_Tmp输血;
      
        --血库部分
        v_Xtmp := v_Xtmp || '<SFXMID>' || r_b.Id || '</SFXMID>'; --收费项目目录.id
        v_Xtmp := v_Xtmp || '<SFXMMC>' || r_b.名称 || '</SFXMMC>'; --收费项目目录.名称
        v_Xtmp := v_Xtmp || '<SFXMGG>' || r_b.规格 || '</SFXMGG>'; --收费项目目录.规格
        v_Xtmp := v_Xtmp || '<BM>' || v_收费商品名 || '</BM>'; --收费项目别名.名称（商品名）
        v_Xtmp := v_Xtmp || '<DW>' || r_b.单位 || '</DW>'; --收费项目目录.计算单位
        v_Xtmp := v_Xtmp || '<XDBH>' || r_b.血袋编号 || '</XDBH>'; --血液收发记录.血袋编号
        v_Xtmp := v_Xtmp || '<SXXH>' || r_b.序号 || '</SXXH>'; --血液收发记录.序号
      
        v_Xtmp := v_Xtmp || '</ITEM>';
        Select Appendchildxml(x_Item, '/ITEMLIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
      End Loop;
      Close Cbloodlist;
    End If;
  
    --西药成药医嘱
    If r.诊疗类别 = '5' Or r.诊疗类别 = '6' Then
      --西/成 药
      If x_Item Is Null Then
        --只初始化一次
        x_Item := Xmltype('<ITEMLIST></ITEMLIST>');
      End If;
      v_收费商品名 := Null;
      If r.收费细目id Is Not Null Then
        For Z In (Select a.名称, a.性质
                  From 收费项目别名 A
                  Where a.收费细目id = r.收费细目id
                  Group By a.名称, a.性质
                  Order By a.性质) Loop
          v_收费商品名 := z.名称;
          If z.性质 = 3 Then
            v_收费商品名 := z.名称;
            Exit;
          End If;
        End Loop;
      End If;
    
      v_Xtmp := '<ITEM jsonArray="True" >';
      v_Xtmp := v_Xtmp || '<YSZT>' || r.医生嘱托 || '</YSZT>'; --病人医嘱记录.医生嘱托
      v_Xtmp := v_Xtmp || '<YZID>' || r.医嘱id || '</YZID>'; --病人医嘱记录.医嘱ID
      v_Xtmp := v_Xtmp || '<RELATEDID>' || r.相关id || '</RELATEDID>'; --病人医嘱记录.相关ID
      v_Xtmp := v_Xtmp || '<GW>' || nvl(r.高危药品,0) || '</GW>'; --高危药标识，1表示高危药，0表示普通
      v_Xtmp := v_Xtmp || '<CDM>' || r.产地 || '</CDM>'; --产地名，收费项目目录.产地    
      v_Xtmp := v_Xtmp || '<ZLXMID>' || r.诊疗项目id || '</ZLXMID>'; --诊疗项目目录.ID
      v_Xtmp := v_Xtmp || '<SFXMID>' || r.收费细目id || '</SFXMID>'; --收费项目目录.id
      v_Xtmp := v_Xtmp || '<SFXMMC>' || r.收费名称 || '</SFXMMC>'; --收费项目目录.名称
      v_Xtmp := v_Xtmp || '<SFXMGG>' || r.规格 || '</SFXMGG>'; --收费项目目录.规格
      v_Xtmp := v_Xtmp || '<BM>' || v_收费商品名 || '</BM>'; --收费项目别名.名称（商品名）
      v_Xtmp := v_Xtmp || '<ZL>' || r.总给予量 || '</ZL>'; --病人医嘱记录.总给予量
      v_Xtmp := v_Xtmp || '<DL>' || r.单次用量 || '</DL>'; --病人医嘱记录.单次用量
      v_Xtmp := v_Xtmp || '<DW>' || r.收费单位 || '</DW>'; --收费项目目录.计算单位
      v_Xtmp := v_Xtmp || '<ZLDW>' || r.诊疗单位 || '</ZLDW>'; --诊疗项目目录.计算单位
      v_Xtmp := v_Xtmp || '<ZXXZ>' || r.执行性质 || '</ZXXZ>'; --病人医嘱记录.执行性质
      v_Xtmp := v_Xtmp || '<ZXKS>' || r.执行科室 || '</ZXKS>'; --诊疗项目目录.执行科室
      v_Xtmp := v_Xtmp || '<XDBH></XDBH>'; --血液收发记录.血袋编号
      v_Xtmp := v_Xtmp || '<SXXH></SXXH>'; --血液收发记录.序号
      v_Xtmp := v_Xtmp || '</ITEM>';
      Select Appendchildxml(x_Item, '/ITEMLIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
    End If;
  End Loop;
  If x_Item Is Not Null Then
    Select Appendchildxml(x_医嘱, '/YZ', x_Item) Into x_医嘱 From Dual;
  End If;
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  Select Appendchildxml(x_Templet, '/OUTPUT', x_医嘱) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getadviceinfo;
/
--123263:王振涛,2018-03-22,集成平台消息锚点
--122998:胡俊勇,2018-03-21,集成平台消息锚点
--122609:胡俊勇,2018-03-08,集成平台消息添加
--123312:梁唐彬,2018-03-22,新病理系统集成平台消息
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
    Id_In       收费项目目录.Id%Type,
    变动类型_In Number
  );
  --诊疗收费对照变动
  Procedure Zlhis_Dict_054
  (
    Id_In     诊疗分类目录.Id%Type,
    原对照_In Varchar2,
    现对照_In Varchar2
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
    Id_In       收费项目目录.Id%Type,
    变动类型_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><变动类型>' || 变动类型_In || '</变动类型></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_053', v_Value);
  End Zlhis_Dict_053;

  --诊疗收费对照变动
  Procedure Zlhis_Dict_054
  (
    Id_In     诊疗分类目录.Id%Type,
    原对照_In Varchar2,
    现对照_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><原对照>' || 原对照_In || '</原对照><现对照>' || 现对照_In || '</现对照></root>';
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

--123263:王振涛,2018-03-22,集成平台消息锚点
--122998:胡俊勇,2018-03-22,集成平台消息锚点
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

  If v_Tmp = 'DELETE' Then
    If v_Tab = '诊疗检查类型' Then
      --删除记录
      For R In (Select a.编码, a.名称, a.简码, a.建病案 From 诊疗检查类型 A Where a.编码 = v_Code) Loop
        b_Message.Zlhis_Dictpacs_003(r.编码, r.名称, r.简码, r.建病案);
      End Loop;
    Elsif v_Tab = '诊疗检验标本' Then
      --删除记录
      For R In (Select a.编码, a.名称, a.简码, a.适用性别 From 诊疗检验标本 A Where a.编码 = v_Code) Loop
        b_Message.Zlhis_Dictlis_006(r.编码, r.名称, r.简码, r.适用性别);
      End Loop;
    End If;
  End If;

  Execute Immediate v_Rulesql;

  If v_Tab = '诊疗检查类型' Then
    If v_Tmp = 'INSERT' Or v_Tmp = 'UPDATE' Then
      For R In (Select a.编码, a.名称, a.简码, a.建病案 From 诊疗检查类型 A Where a.编码 = v_Code) Loop
        If v_Tmp = 'INSERT' Then
          b_Message.Zlhis_Dictpacs_001(r.编码, r.名称, r.简码, r.建病案);
        Else
          b_Message.Zlhis_Dictpacs_002(r.编码, r.名称, r.简码, r.建病案);
        End If;
      End Loop;
    End If;
  Elsif v_Tab = '诊疗检验标本' Then
    If v_Tmp = 'INSERT' Or v_Tmp = 'UPDATE' Then
      For R In (Select a.编码, a.名称, a.简码, a.适用性别 From 诊疗检验标本 A Where a.编码 = v_Code) Loop
        If v_Tmp = 'INSERT' Then
          b_Message.Zlhis_Dictlis_004(r.编码, r.名称, r.简码, r.适用性别);
        Else
          b_Message.Zlhis_Dictlis_005(r.编码, r.名称, r.简码, r.适用性别);
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

--122998:胡俊勇,2018-03-22,集成平台消息锚点
Create Or Replace Procedure Zl_诊疗检查部位_Edit
(
  操作_In     In Number, --1:增加;2:修改;3:删除
  类型_In     In 诊疗检查部位.类型%Type,
  原编码_In   In 诊疗检查部位.编码%Type,
  新编码_In   In 诊疗检查部位.编码%Type := Null,
  名称_In     In 诊疗检查部位.名称%Type := Null,
  分组_In     In 诊疗检查部位.分组%Type := Null,
  备注_In     In 诊疗检查部位.备注%Type := Null,
  方法_In     In 诊疗检查部位.方法%Type := Null,
  适用性别_In In 诊疗检查部位.适用性别%Type := Null
) Is
  v_原名称 诊疗检查部位.名称%Type := Null;
  e_Notfind Exception;
  v_方法   Varchar2(1000);
  v_Fields Varchar2(1000);
  v_Tmp    Varchar2(1000);
  n_Count  Number;
  n_记录id 诊疗项目部位.Id%Type;
Begin
  If 操作_In = 1 Then
    Insert Into 诊疗检查部位
      (类型, 编码, 名称, 分组, 备注, 方法, 适用性别)
    Values
      (类型_In, 新编码_In, 名称_In, 分组_In, 备注_In, 方法_In, 适用性别_In);
    b_Message.Zlhis_Dictpacs_004(类型_In, 新编码_In, 名称_In, 分组_In, 备注_In, 方法_In, 适用性别_In);
  Elsif 操作_In = 2 Then
    Begin
      Select 名称 Into v_原名称 From 诊疗检查部位 Where 编码 = 原编码_In And 类型 = 类型_In;
    Exception
      When Others Then
        Null;
    End;
    If v_原名称 Is Null Then
      Raise e_Notfind;
    End If;
    Update 诊疗检查部位
    Set 编码 = 新编码_In, 名称 = 名称_In, 分组 = 分组_In, 备注 = 备注_In, 方法 = 方法_In, 适用性别 = 适用性别_In
    Where 编码 = 原编码_In And 类型 = 类型_In;
    b_Message.Zlhis_Dictpacs_005(类型_In, 新编码_In, 名称_In, 分组_In, 备注_In, 方法_In, 适用性别_In);
  
    --级联修改
    v_方法 := ';' || 方法_In;
    v_方法 := Replace(v_方法, ',', Chr(10));
    v_方法 := Replace(v_方法, Chr(9), ';');
    v_方法 := Replace(v_方法, ';0', Chr(10));
    v_方法 := Replace(v_方法, ';1', Chr(10));
    v_方法 := Replace(v_方法, Chr(10), ';');
    v_方法 := Replace(v_方法, ';;', ';');
    v_方法 := v_方法 || ';';
  
    v_方法 := Substr(v_方法, 2);
  
    --原有的方法，现在已经删除了或原有的部位的名称已经改变了
    For r_Used In (Select ID, 项目id, 部位, 方法, 类型, 默认 From 诊疗项目部位 Where 部位 = v_原名称 And 类型 = 类型_In) Loop
      If Instr(';' || v_方法, ';' || r_Used.方法 || ';') = 0 Then
        b_Message.Zlhis_Dictpacs_009(r_Used.Id, r_Used.项目id, r_Used.类型, r_Used.部位, r_Used.方法, r_Used.默认);
        Delete 诊疗项目部位
        Where 项目id = r_Used.项目id And 部位 = r_Used.部位 And 方法 = r_Used.方法 And 类型 = r_Used.类型;
      Else
        Update 诊疗项目部位
        Set 部位 = 名称_In
        Where 项目id = r_Used.项目id And 部位 = r_Used.部位 And 方法 = r_Used.方法 And 类型 = r_Used.类型;
        b_Message.Zlhis_Dictpacs_008(r_Used.Id, r_Used.项目id, r_Used.类型, r_Used.部位, r_Used.方法, r_Used.默认);
      End If;
    End Loop;
  
    --原来没有的方法现在新增
    v_Tmp := v_方法;
    While v_Tmp Is Not Null Loop
      --依次取每个项目
      v_Fields := Substr(v_Tmp, 1, Instr(v_Tmp, ';') - 1);
      v_Tmp    := Substr(v_Tmp, Instr(v_Tmp, ';') + 1);
    
      If v_Fields Is Not Null Then
        For r_Used In (Select Distinct 项目id From 诊疗项目部位 Where 部位 = 名称_In And 类型 = 类型_In) Loop
          Select Count(ID)
          Into n_Count
          From 诊疗项目部位
          Where 项目id = r_Used.项目id And 部位 = 名称_In And 类型 = 类型_In And 方法 = v_Fields;
        
          If n_Count = 0 Then
            Select 诊疗项目部位_Id.Nextval Into n_记录id From Dual;
            Insert Into 诊疗项目部位
              (ID, 项目id, 类型, 部位, 方法)
            Values
              (n_记录id, r_Used.项目id, 类型_In, 名称_In, v_Fields);
            b_Message.Zlhis_Dictpacs_007(n_记录id, r_Used.项目id, 类型_In, 名称_In, v_Fields, Null);
          End If;
        End Loop;
      End If;
    End Loop;
  Elsif 操作_In = 3 Then
    For R In (Select a.类型, a.编码, a.名称, a.分组, a.备注, a.方法, a.适用性别
              From 诊疗检查部位 A
              Where a.编码 = 原编码_In And a.类型 = 类型_In) Loop
      b_Message.Zlhis_Dictpacs_006(r.类型, r.编码, r.名称, r.分组, r.备注, r.方法, r.适用性别);
    End Loop;
    Delete 诊疗检查部位 Where 编码 = 原编码_In And 类型 = 类型_In;
  End If;

Exception
  When e_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该部位不存在，可能已被其他用户删除修改！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗检查部位_Edit;
/

--122998:胡俊勇,2018-03-22,集成平台消息锚点
Create Or Replace Procedure Zl_诊疗项目部位_Insert
(
  项目id_In In 诊疗项目部位.项目id%Type,
  类型_In   In 诊疗项目部位.类型%Type,
  部位_In   In 诊疗项目部位.部位%Type,
  方法_In   In 诊疗项目部位.方法%Type,
  默认_In   In 诊疗项目部位.默认%Type := Null
) As
  v_Code Varchar2(20); --编码
  Err_Notfind Exception;
  n_记录id 诊疗项目部位.Id%Type;
Begin
  Select RTrim(编码) Into v_Code From 诊疗项目目录 Where 类别 = 'D' And ID = 项目id_In;
  If v_Code Is Null Then
    Raise Err_Notfind;
  End If;
  Select 诊疗项目部位_Id.Nextval Into n_记录id From Dual;
  Insert Into 诊疗项目部位
    (ID, 项目id, 类型, 部位, 方法, 默认)
  Values
    (n_记录id, 项目id_In, 类型_In, 部位_In, 方法_In, 默认_In);
  b_Message.Zlhis_Dictpacs_007(n_记录id, 项目id_In, 类型_In, 部位_In, 方法_In, 默认_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]该项目不存在，可能已被其他用户删除！[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗项目部位_Insert;
/

--122998:胡俊勇,2018-03-21,集成平台消息锚点
Create Or Replace Procedure Zl_诊疗项目部位_Delete(项目id_In In 诊疗项目部位.项目id%Type) As
Begin
  For R In (Select a.Id, a.项目id, a.类型, a.部位, a.方法, a.默认 From 诊疗项目部位 A Where a.项目id = 项目id_In) Loop
    b_Message.Zlhis_Dictpacs_009(r.Id, r.项目id, r.类型, r.部位, r.方法, r.默认);
  End Loop;
  Delete 诊疗项目部位 Where 项目id = 项目id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_诊疗项目部位_Delete;
/

--123312:梁唐彬,2018-03-22,新病理系统集成平台消息
--123225:梁唐彬,2018-03-20,病理系统修改申请问题
CREATE OR REPLACE Procedure Zl_医嘱申请单文件_Edit
( 
  文件id_In 医嘱申请单文件.文件id%Type, 
  文件名_IN 医嘱申请单文件.文件名%Type, 
  类别_In   医嘱申请单文件.类别%Type, 
  医嘱ID_In   医嘱申请单文件.医嘱ID%Type
) As 
   n_病人ID 病人医嘱记录.病人ID%Type;
   n_主页id 病人医嘱记录.主页id%Type;
   v_挂号单 病人医嘱记录.挂号单%Type;
   n_发送号 病人医嘱发送.发送号%Type;
   n_组ID  病人医嘱记录.ID%TYPE;
   n_病人来源 病人医嘱记录.病人来源%Type;
Begin 
  Delete From 医嘱申请单文件 Where 文件id = 文件id_In And 医嘱ID = 医嘱ID_In And 类别 = 类别_In;
  If Sql%Rowcount <> 0 And 类别_In = 2 Then
    Select Max(a.病人id), Max(a.主页id), Max(a.挂号单), Max(b.发送号), Max(Nvl(a.相关id, a.Id)), Max(a.病人来源)
    Into n_病人id, n_主页id, v_挂号单, n_发送号, n_组id, n_病人来源
    From 病人医嘱记录 A, 病人医嘱发送 B
    Where a.Id = b.医嘱id And a.Id = 医嘱id_In;
    b_Message.Zlhis_Cis_056(n_病人ID, n_主页id, v_挂号单,n_发送号, n_组ID, n_病人来源);
  End If;
  Insert Into 医嘱申请单文件 
      (文件id,文件名, 医嘱ID, 类别) 
  Values 
      (文件id_In,文件名_IN, 医嘱ID_In, 类别_In); 
Exception 
  When Others Then 
    Zl_Errorcenter(Sqlcode, Sqlerrm); 
End Zl_医嘱申请单文件_Edit;
/







------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0003' Where 编号=&n_System;
Commit;
