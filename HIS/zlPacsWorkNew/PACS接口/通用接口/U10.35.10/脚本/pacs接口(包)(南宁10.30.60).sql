create or replace package b_PacsInterface is
  Type t_Refcur Is Ref Cursor;


  -----------------------------------------------------------------------------
  --获取检查科室信息
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  Procedure GetDeptItems
  (
  Cursor_Out  Out	t_Refcur,
  过滤条件_In  In  Varchar2:=Null
  );


  -----------------------------------------------------------------------------
  -- 功    能：获取费别
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  Procedure GetChargeTypes
  (
    Cursor_Out  Out  t_Refcur,
    过滤条件_In  In  Varchar2:=Null
  );


  -----------------------------------------------------------------------------
  -- 功    能：获取pacs检查项目
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  procedure GetPacsItems
  (
    Cursor_Out  Out  t_Refcur,
    过滤条件_In  In  Varchar2:=Null
  );


  -----------------------------------------------------------------------------
  -- 功    能：获取检查项目明细
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  procedure GetAdviceItems
  (
    Cursor_Out  Out  t_Refcur,
    医嘱id_In  In  病人医嘱记录.ID%Type
  );

  -----------------------------------------------------------------------------
  -- 功    能：获取检查费用明细
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  procedure GetAdviceFees
  (
    Cursor_Out  Out  t_Refcur,
    医嘱id_In  In  病人医嘱记录.ID%Type
  );


  -----------------------------------------------------------------------------
  -- 功    能：返回科室医生信息
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  procedure GetPacsDeptDoctor
  (
    Cursor_Out  Out  t_Refcur,
    过滤条件_In  In  Varchar2:=Null
  );

  -----------------------------------------------------------------------------
  -- 功    能：获取病人信息
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  Procedure GetPatient
  (
    Cursor_Out  Out  t_Refcur,
    查找方式_In  In  Number,
    查找内容_In  In  Varchar2
  );


  ---------------------------------------------------------------------------------------------------------------
  -- 功    能：获取医嘱申请状态
  -- 调用列表：
  --
  --
  ---------------------------------------------------------------------------------------------------------------
  Procedure GetRequestStatus
  (
    Cursor_Out  Out  t_Refcur,
    医嘱id_In  In  病人医嘱记录.ID%Type
  );


  ---------------------------------------------------------------------------------------------------------------
  -- 功    能：获取检查申请信息
  -- 调用列表：
  --
  --
  ---------------------------------------------------------------------------------------------------------------
  Procedure GetRequestInfo
  (
    Cursor_Out  Out  t_Refcur,
    查找方式_In  In  Number,
    查找内容_In  In  Varchar2,
    过滤条件_In  In  Varchar2:=null
  );



  ---------------------------------------------------------------------------------------------------------------
	-- 功    能：获取心电检查的申请信息
  -- 调用列表：
  --
  --
	---------------------------------------------------------------------------------------------------------------
	Procedure GetRequestInfo1
	(
		Cursor_Out	Out	t_Refcur,
		开始日期_In	In	Varchar2,
		结束日期_In	In	Varchar2,
    检查类别_In In  Varchar2
	);


  -----------------------------------------------------------------------------
  -- 功    能：取消检验/检查申请
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  Procedure CancelRequest
  (
    医嘱id_In  In  病人医嘱发送.医嘱ID%Type,
    单独执行_In   Number := 0,
    执行部门ID_IN 部门表.id%Type := 0
  );


    -----------------------------------------------------------------------------
  -- 功    能：删除报告信息
  -----------------------------------------------------------------------------
  PROCEDURE DeleteReport
  (
    医嘱id_IN 病人医嘱发送.医嘱ID%TYPE
  );


  -----------------------------------------------------------------------------
  -- 功    能：删除心电报告信息
  -----------------------------------------------------------------------------
  PROCEDURE DeleteElectrocardioReport
  (
    医嘱id_IN 病人医嘱发送.医嘱ID%TYPE
  );



  -----------------------------------------------------------------------------
  -- 功    能：清除报告数据
  -----------------------------------------------------------------------------
  PROCEDURE ClearPacsReport
  (
    医嘱id_IN 病人医嘱发送.医嘱ID%TYPE
  );

  -----------------------------------------------------------------------------
  -- 功    能：接收检验/检查申请
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  Procedure RecevieRequest
  (
    医嘱id_IN     病人医嘱发送.医嘱ID%TYPE,

    执行间_IN     病人医嘱发送.执行间%TYPE:=Null,
    检查号_IN     影像检查记录.检查号%TYPE:=NULL,
    检查设备_IN   影像检查记录.检查设备%TYPE:=Null,
    身高_IN       影像检查记录.身高%TYPE:=Null,
    体重_IN       影像检查记录.体重%TYPE:=Null,
    检查技师_IN   影像检查记录.检查技师%TYPE:=Null,
    执行时间_IN   病人医嘱发送.安排时间%TYPE:=Null,
    执行说明_IN   病人医嘱发送.执行说明%TYPE:=NULL,
    单独执行_In   Number := 0,
    执行部门ID_IN 部门表.id%Type := 0
  );

  -----------------------------------------------------------------------------
  -- 功    能：发送报告文本信息
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  procedure SendReport
  (
    医嘱id_IN       病人医嘱发送.医嘱ID%TYPE,
    报告所见_IN     电子病历内容.内容文本%TYPE,
    报告建议_IN     电子病历内容.内容文本%TYPE,
    报告医生_IN     电子病历记录.创建人%TYPE,
    审核医生_IN     影像检查记录.复核人%TYPE := Null,
    执行部门ID_IN 部门表.id%Type := 0
  );


  -----------------------------------------------------------------------------
  -- 功    能：发送心电报告信息
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  procedure SendElectrocardioReport
  (
    医嘱id_IN       病人医嘱发送.医嘱ID%TYPE,
    报告标题_IN     电子病历内容.内容文本%TYPE,
    诊断结果_IN     电子病历内容.内容文本%TYPE,
    诊断建议_IN     电子病历内容.内容文本%TYPE,
    报告医生_IN     电子病历记录.创建人%TYPE,
    审核医生_IN     影像检查记录.复核人%TYPE := Null
  );

  -----------------------------------------------------------------------------
  -- 功    能：清除报告附件
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  procedure ClearReportAffix
  (
    病历id_In       电子病历附件.病历ID%TYPE,
    附件标记_IN     电子病历附件.创建人%TYPE
  );


  -----------------------------------------------------------------------------
  -- 功    能：添加报告附件
  -- 调用列表：
  --
  --
  -----------------------------------------------------------------------------
  Procedure AddReportAffix
  (
    病历id_In In 电子病历附件.病历id%Type,
    文件名_In In 电子病历附件.文件名%Type,
    大小_In   In 电子病历附件.大小%Type,
    附件标记_IN in  电子病历附件.创建人%TYPE
  );


  -----------------------------------------------------------------------------
  -- 功    能：添加心电报告图像
  -- 调用列表：
  --返回图形行记录ID
  --
  -----------------------------------------------------------------------------
  function AddElectrocardioReportImage
  (
    医嘱id_IN       病人医嘱发送.医嘱ID%TYPE
  )return number;

end b_PacsInterface;
/
create or replace package body b_PacsInterface is

  -----------------------------------------------------------------------------
  --获取检查科室信息
  -----------------------------------------------------------------------------
  Procedure GetDeptItems
  (
  Cursor_Out  Out  t_Refcur,
  过滤条件_In  In  Varchar2:=Null
  ) is
  begin

    If 过滤条件_In Is Null Then
      Open Cursor_Out For
        Select p.ID, P.名称, p.编码, p.简码, p.位置 from 部门表 P, 部门性质说明 C where P.id = c.部门id and c.工作性质 = '检查';
    Else
      Open Cursor_Out For
        Select p.ID, P.名称, p.编码, p.简码, p.位置 from 部门表 P, 部门性质说明 C where P.id = c.部门id and c.工作性质 = '检查'
        And (p.编码 = 过滤条件_In Or p.名称 Like '%'||过滤条件_In||'%');
    End If;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetDeptItems;


  -----------------------------------------------------------------------------
  -- 功    能：获取费别
  -----------------------------------------------------------------------------
  Procedure GetChargeTypes
  (
    Cursor_Out  Out  t_Refcur,
    过滤条件_In  In  Varchar2:=Null
  ) is
  begin

    If 过滤条件_In Is Null Then
      Open Cursor_Out For
        Select 编码,名称,缺省标志 From 费别 a;

    Else
      Open Cursor_Out For
        Select 编码,名称,缺省标志 From 费别 a
        Where (a.编码 = 过滤条件_In Or a.名称 Like '%'||过滤条件_In||'%');
    End If;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetChargeTypes;


  -----------------------------------------------------------------------------
  -- 功    能：获取pacs检查项目
  -----------------------------------------------------------------------------
  procedure GetPacsItems
  (
    Cursor_Out  Out  t_Refcur,
    过滤条件_In  In  Varchar2:=Null
  ) is
  begin

    If 过滤条件_In Is Null Then
      Open Cursor_Out For
        Select /*+ RULE */
          Distinct a.Id as 诊疗项目ID, a.编码, a.名称, Decode(a.适用性别, 1, '男', 2, '女', '通用') 适用性别, a.计算单位 As 单位,
                   Decode(a.服务对象, 1, '门诊', 2, '住院', '通用') 适用场合, a.操作类型 检查类型, b.部位 检查部位,
                   b.方法 检查方法, (a.名称 || '_' || b.部位 || '_' || b.方法 ) as 部位方法组合
          From 诊疗项目目录 a, 诊疗项目部位 b
          Where Nvl(To_Char(a.撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And a.类别 = 'D' And a.Id = b.项目id(+);

    Else
      Open Cursor_Out For
        Select /*+ RULE */
          Distinct a.Id as 诊疗项目ID, a.编码, a.名称, Decode(a.适用性别, 1, '男', 2, '女', '通用') 适用性别, a.计算单位 As 单位,
                   Decode(a.服务对象, 1, '门诊', 2, '住院', '通用') 适用场合, a.操作类型 检查类型, b.部位 检查部位,
                   b.方法 检查方法, (a.名称 || '_' || b.部位 || '_' || b.方法 ) as 部位方法组合
          From 诊疗项目目录 a, 诊疗项目部位 b
          Where Nvl(To_Char(a.撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And a.类别 = 'D' And a.Id = b.项目id(+)
          And (a.编码 = 过滤条件_In Or a.名称 Like '%'||过滤条件_In||'%');
    End If;

   Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetPacsItems;


    -----------------------------------------------------------------------------
  -- 功    能：获取检查项目明细
  -----------------------------------------------------------------------------
  procedure GetAdviceItems
  (
    Cursor_Out  Out  t_Refcur,
    医嘱id_In  In  病人医嘱记录.ID%Type
  ) Is
  Begin
    Open Cursor_Out For
          select /*+ RULE */ a.ID As 部位医嘱ID, a.诊疗项目id, c.名称 As 诊疗项目名称, a.标本部位, a.检查方法, Decode(c.操作类型,'X线','DR','MRI','MR',c.操作类型) as 检查类型,
                (/*c.名称 || '_' || */a.标本部位 /*|| '_' */|| replace(replace(a.检查方法,'(','') ,')','')) as 部位方法组合
                from 病人医嘱记录 a, 诊疗项目目录 c ,病人医嘱发送 b
                Where a.诊疗项目id=c.id and  a.Id = b.医嘱ID And b.执行状态=0 And 相关id=医嘱id_In ;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetAdviceItems;

      -----------------------------------------------------------------------------
  -- 功    能：获取检查项目费用
  -----------------------------------------------------------------------------
  procedure GetAdviceFees
  (
    Cursor_Out  Out  t_Refcur,
    医嘱id_In  In  病人医嘱记录.ID%Type
  ) Is
  v_病人来源 病人医嘱记录.病人来源%Type;
  v_记录性质 病人医嘱发送.记录性质%Type;
  v_门诊记账 病人医嘱发送.门诊记帐%Type;
  v_单据号   病人医嘱发送.NO%Type;
  v_发送号   病人医嘱发送.发送号%Type;
  strSQL varchar2(2000);
  strFeeTable Varchar2(20);

  Begin

   Select a.病人来源,b.记录性质,nvl(b.门诊记帐,0) As 门诊记帐,b.NO,b.发送号 Into v_病人来源,v_记录性质,v_门诊记账,v_单据号,v_发送号
           From 病人医嘱记录 a,病人医嘱发送 b
           Where a.id=b.医嘱ID And a.id = 医嘱id_In;


    If v_病人来源 = 2 And v_记录性质 = 2 And v_门诊记账 = 0 Then
        --查 "住院费用记录"
        strFeeTable :='住院费用记录';
    Else
        --查 "门诊费用记录"
        strFeeTable :='门诊费用记录';
    End If;

     strSQL := 'Select  ''主费用'' As 费用类型,decode(A.记录性质,1,''收费单据'',''记帐单据'') As 单据类型,
                 A.NO As 单据号,A.应收金额,A.实收金额,A.数次 || '' '' || A.计算单位 as 数量,
                      Decode(A.记录性质,1,Decode(A.记录状态,0,''收费划价'',1,''已收费'',3,''已退费''),2,
                          Decode(A.记录状态,0,''记帐划价'',1,''已记帐'',3,''已销帐''),''未计费'') as 计费状态,e.名称|| '' ''|| e.规格 as 项目
                 From ' || strFeeTable || ' A,病人医嘱记录 B ,病人医嘱发送 C,收费项目目录 E
                 Where A.NO= ''' || v_单据号 || ''' And A.记录状态 IN(0,1,3) And A.医嘱序号+0=B.ID And A.记录性质= ' || v_记录性质
                       || ' And  c.医嘱ID=b.Id  And A.收费细目ID=E.Id
          Union ALL
          Select  ''附加费用'' As 费用类型,decode(B.记录性质,1,''收费单据'',''记帐单据'') As 单据类型,
                 B.NO As 单据号,B.应收金额,B.实收金额,B.数次 || '' '' || B.计算单位 as 数量,
                      Decode(B.记录性质,1,Decode(B.记录状态,0,''收费划价'',1,''已收费'',3,''已退费''),2,
                          Decode(B.记录状态,0,''记帐划价'',1,''已记帐'',3,''已销帐''),''未计费'') as 计费状态,e.名称|| '' ''|| e.规格 as 项目
                 From 病人医嘱记录 C,' || strFeeTable || ' B,病人医嘱附费 A ,病人医嘱发送 D ,收费项目目录 E
                 Where A.NO=B.NO And A.记录性质=B.记录性质 And A.医嘱ID=B.医嘱序号+0
                       And A.医嘱ID IN (Select ID From 病人医嘱记录 Where (ID= ' || 医嘱id_In || ' Or 相关ID= ' ||
                       医嘱id_In || ') )
                       And A.发送号= ' || v_发送号 || ' And B.记录状态 IN(0,1,3) And A.医嘱ID=C.ID And A.记录性质= ' ||
                       v_记录性质 || ' And d.医嘱ID =c.Id  And B.收费细目ID=E.Id ';

     Open Cursor_Out For strSQL;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetAdviceFees;

  -----------------------------------------------------------------------------
  -- 功    能：返回科室医生信息
  -----------------------------------------------------------------------------
  procedure GetPacsDeptDoctor
  (
    Cursor_Out  Out  t_Refcur,
    过滤条件_In  In  Varchar2:=Null
  ) is
  Begin
    --...
    null;
  end GetPacsDeptDoctor;


  -----------------------------------------------------------------------------
  -- 功    能：获取病人信息
  -----------------------------------------------------------------------------
  Procedure GetPatient
  (
    Cursor_Out  Out  t_Refcur,
    查找方式_In  In  Number,
    查找内容_In  In  Varchar2
  ) is
  begin
    If 查找方式_In=1 Then
      Open Cursor_Out For
        Select 病人id,姓名,性别,年龄,出生日期,replace(身份证号,'未带','') As  身份证号,婚姻状况,民族,国籍,/*职业,籍贯,*/学历,联系人姓名,nvl(联系人电话,家庭电话) As 联系人电话,nvl(家庭地址,联系人地址) As 联系人地址,工作单位,就诊卡号,健康号,门诊号,住院号,费别,当前床号 From 病人信息 Where 病人id = zl_to_number(查找内容_In);
    ElsIf 查找方式_In=2 Then
      Open Cursor_Out For
        Select 病人id,姓名,性别,年龄,出生日期,replace(身份证号,'未带','') As  身份证号,婚姻状况,民族,国籍,/*职业,籍贯,*/学历,联系人姓名,nvl(联系人电话,家庭电话) As 联系人电话,nvl(家庭地址,联系人地址) As 联系人地址,工作单位,就诊卡号,健康号,门诊号,住院号,费别,当前床号 From 病人信息 Where 住院号 = zl_to_number(查找内容_In);
    ElsIf 查找方式_In=3 Then
      Open Cursor_Out For
        Select 病人id,姓名,性别,年龄,出生日期,replace(身份证号,'未带','') As  身份证号,婚姻状况,民族,国籍,/*职业,籍贯,*/学历,联系人姓名,nvl(联系人电话,家庭电话) As 联系人电话,nvl(家庭地址,联系人地址) As 联系人地址,工作单位,就诊卡号,健康号,门诊号,住院号,费别,当前床号 From 病人信息 Where 门诊号 = zl_to_number(查找内容_In);
    ElsIf 查找方式_In=4 Then
      Open Cursor_Out For
        Select 病人id,姓名,性别,年龄,出生日期,replace(身份证号,'未带','') As  身份证号,婚姻状况,民族,国籍,/*职业,籍贯,*/学历,联系人姓名,nvl(联系人电话,家庭电话) As 联系人电话,nvl(家庭地址,联系人地址) As 联系人地址,工作单位,就诊卡号,健康号,门诊号,住院号,费别,当前床号 From 病人信息 Where 就诊卡号 = 查找内容_In;
    ElsIf 查找方式_In=5 Then
      Open Cursor_Out For
        Select 病人id,姓名,性别,年龄,出生日期,replace(身份证号,'未带','') As  身份证号,婚姻状况,民族,国籍,/*职业,籍贯,*/学历,联系人姓名,nvl(联系人电话,家庭电话) As 联系人电话,nvl(家庭地址,联系人地址) As 联系人地址,工作单位,就诊卡号,健康号,门诊号,住院号,费别,当前床号 From 病人信息 Where 身份证号 = 查找内容_In;
    ElsIf 查找方式_In=6 Then
      Open Cursor_Out For
        Select 病人id,姓名,性别,年龄,出生日期,replace(身份证号,'未带','') As  身份证号,婚姻状况,民族,国籍,/*职业,籍贯,*/学历,联系人姓名,nvl(联系人电话,家庭电话) As 联系人电话,nvl(家庭地址,联系人地址) As 联系人地址,工作单位,就诊卡号,健康号,门诊号,住院号,费别,当前床号 From 病人信息 Where 健康号 = 查找内容_In;
    ElsIf 查找方式_In=7 Then
      Open Cursor_Out For
        Select 病人id,姓名,性别,年龄,出生日期,replace(身份证号,'未带','') As  身份证号,婚姻状况,民族,国籍,/*职业,籍贯,*/学历,联系人姓名,nvl(联系人电话,家庭电话) As 联系人电话,nvl(家庭地址,联系人地址) As 联系人地址,工作单位,就诊卡号,健康号,门诊号,住院号,费别,当前床号 From 病人信息 Where 姓名 Like '%'||查找内容_In||'%';
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetPatient;


  ---------------------------------------------------------------------------------------------------------------
  -- 功    能：获取医嘱申请状态
  ---------------------------------------------------------------------------------------------------------------
  Procedure GetRequestStatus
  (
    Cursor_Out  Out  t_Refcur,
    医嘱id_In  In  病人医嘱记录.ID%Type
  )is
  begin
    Open Cursor_Out For
      Select a.医嘱状态,b.执行状态, b.执行过程
      From 病人医嘱记录 a,病人医嘱发送 b
      Where a.ID=b.医嘱id and a.ID=医嘱id_In And RowNum<2;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetRequestStatus;


  ---------------------------------------------------------------------------------------------------------------
  -- 功    能：获取检查申请信息
  ---------------------------------------------------------------------------------------------------------------
  Procedure GetRequestInfo
  (
    Cursor_Out  Out  t_Refcur,
    查找方式_In  In  Number,
    查找内容_In  In  Varchar2,
    过滤条件_In  In  Varchar2:=null
  ) is
    strSql varchar2(2000);
    v_病历摘要要素ID 病人医嘱附件.要素ID%Type;
    v_临床诊断要素ID 病人医嘱附件.要素ID%Type;
  Begin


    select ID into v_病历摘要要素ID from 诊治所见项目 where  诊治所见项目.中文名='病人主诉';
    select ID Into v_临床诊断要素ID from 诊治所见项目 where  诊治所见项目.中文名='最后诊断';

    If 查找方式_In=1 Then
      strSql :=  ' c.病人id =' || zl_to_number(查找内容_In);
    ElsIf 查找方式_In=2 Then
      strSql := ' c.住院号 =' || zl_to_number(查找内容_In);
    ElsIf 查找方式_In=3 Then
      strSql := ' c.门诊号 =' || zl_to_number(查找内容_In);
    ElsIf 查找方式_In=4 Then
      strSql := ' c.就诊卡号 =''' || 查找内容_In || '''';
    ElsIf 查找方式_In=5 Then
      strSql := ' c.身份证号 =''' || 查找内容_In || '''';
    ElsIf 查找方式_In=6 Then
      strSql := ' c.健康号 =''' || 查找内容_In || '''';
    ElsIf 查找方式_In=7 Then
      strSql := ' c.姓名 like ''%' || 查找内容_In  || '%''';
    ElsIf 查找方式_In=8 Then
      strSql := ' a.Id =' || zl_to_number(查找内容_In);
    End If;

    if Nvl(过滤条件_In,'')<>'' then
      strSql :='And '||过滤条件_In;
    end if;
--过滤前两天未执行的检查项目
    strSql := 'Select Distinct nvl(a.相关ID,a.Id ) As 医嘱ID,c.姓名,c.门诊号,c.住院号,c.性别,c.年龄
           From 病人医嘱记录 a, 病人医嘱发送 b, 病人信息 c ,
                (Select 部门ID From 部门性质说明 Where 工作性质 =''检查'') d
           Where b.执行部门ID = d.部门ID
                 And (b.执行过程 Is Null Or b.执行过程=1 Or b.执行过程 = 0) And b.执行状态 = 0
                 And b.发送时间 > To_Date(To_Char(Sysdate - 3,''yyyy-mm-dd'') || ''23:59:59'',''yyyy-mm-dd hh24:mi:ss'')
                 And a.诊疗类别 = ''D'' And a.Id = b.医嘱ID
                 And a.病人ID = c.病人ID and ' || strSql;


    strSql :='Select /*+ RULE */
        k.医嘱id,m.主页id,m.开嘱科室ID As 申请科室ID,p.名称 As 申请科室,m.开嘱医生 As 申请人,
        m.开嘱时间 As 申请时间,replace(m.医嘱内容,'','',''|'') as 医嘱内容,m.诊疗项目id,m.执行科室ID As 执行部门ID ,n.名称 As 执行部门,
        m.病人id,k.姓名,k.门诊号,k.住院号,k.性别,k.年龄,
        Decode(m.病人来源, 1, ''门诊'', 2, ''住院'', 3, ''外来'', 4, ''体检'')  as 病人来源,
        Decode(m.紧急标志,1,1,nvl((Select 急诊 From 病人挂号记录 Where No = m.挂号单),0)) As 紧急标志,
        (select 内容 from 病人医嘱附件 where 要素ID=' || v_病历摘要要素ID || ' and 医嘱ID= m.id ) as 病历摘要,
        (select 内容 from 病人医嘱附件 where 要素ID=' || v_临床诊断要素ID || ' and 医嘱ID=  m.id) as 临床诊断
      From ( ' || strSql || ' )  k ,病人医嘱记录 m ,部门表 n,部门表 p
      Where k.医嘱ID = m.Id And m.执行科室ID = n.Id And m.开嘱科室ID = p.Id';

   Open Cursor_Out For strSql;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetRequestInfo;


  ---------------------------------------------------------------------------------------------------------------
	-- 功    能：获取心电检查的申请信息
	---------------------------------------------------------------------------------------------------------------
	Procedure GetRequestInfo1
	(
		Cursor_Out	Out	t_Refcur,
		开始日期_In	In	Varchar2,
		结束日期_In	In	Varchar2,
    检查类别_In In  Varchar2
	)is
    strSql varchar2(2000);
    v_病历摘要要素ID 病人医嘱附件.要素ID%Type;
    v_临床诊断要素ID 病人医嘱附件.要素ID%Type;
  Begin


    select ID into v_病历摘要要素ID from 诊治所见项目 where  诊治所见项目.中文名='病人主诉';
    select ID Into v_临床诊断要素ID from 诊治所见项目 where  诊治所见项目.中文名='最后诊断';

    strSql := 'Select Distinct nvl(a.相关ID,a.Id ) As 医嘱ID,c.姓名,c.门诊号,c.住院号,c.性别,c.年龄
           From 病人医嘱记录 a, 病人医嘱发送 b, 病人信息 c ,诊疗项目目录 e,
                (Select 部门ID From 部门性质说明 Where 工作性质 =''检查'') d
           Where b.执行部门ID = d.部门ID
                 And (b.执行过程 Is Null Or b.执行过程=1 Or b.执行过程 = 0) And b.执行状态 = 0
                 And a.诊疗类别 = ''D'' And a.Id = b.医嘱ID and a.诊疗项目ID=e.id and e.操作类型 like ''%' || 检查类别_In || '%''
                 And a.病人ID = c.病人ID and b.发送时间 between to_date(''' || 开始日期_In || ''', ''yyyy-mm-dd hh24:mi:ss'')  and  to_date(''' || 结束日期_In || ''', ''yyyy-mm-dd hh24:mi:ss'')';


    strSql :='Select /*+ RULE */
        k.医嘱id,m.主页id,m.开嘱科室ID As 申请科室ID,p.名称 As 申请科室,m.开嘱医生 As 申请人,
        m.开嘱时间 As 申请时间,m.医嘱内容,m.诊疗项目id,m.执行科室ID As 执行部门ID ,n.名称 As 执行部门,
        m.病人id,k.姓名,k.门诊号,k.住院号,k.性别,k.年龄,
        Decode(m.病人来源, 1, ''门诊'', 2, ''住院'', 3, ''外来'', 4, ''体检'') 病人来源,
        Decode(m.紧急标志,1,1,nvl((Select 急诊 From 病人挂号记录 Where No = m.挂号单),0)) As 紧急标志,
        (select 内容 from 病人医嘱附件 where 要素ID=' || v_病历摘要要素ID || ' and 医嘱ID= m.id ) as 病历摘要,
        (select 内容 from 病人医嘱附件 where 要素ID=' || v_临床诊断要素ID || ' and 医嘱ID=  m.id) as 临床诊断
      From ( ' || strSql || ' )  k ,病人医嘱记录 m ,部门表 n,部门表 p
      Where k.医嘱ID = m.Id And m.执行科室ID = n.Id And m.开嘱科室ID = p.Id';


   Open Cursor_Out For strSql;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetRequestInfo1;


  -----------------------------------------------------------------------------
  -- 功    能：取消检验/检查申请
  -----------------------------------------------------------------------------
  Procedure CancelRequest
  (
    医嘱id_In  In  病人医嘱发送.医嘱ID%Type,
    单独执行_In   Number := 0,
    执行部门ID_IN 部门表.id%Type := 0
    --参数：医嘱ID_IN=单独执行的医嘱ID。
    --      单独执行_In=检查医嘱组合是否采用对每个项目分散单独执行的方式
  ) is
    v_发送号 影像检查记录.发送号%Type;
  Begin

    select 发送号 into V_发送号 from 病人医嘱发送 where 医嘱ID=医嘱id_In;

    Zl_影像检查_Cancel(医嘱id_In, v_发送号,单独执行_In,执行部门ID_IN);

  EXCEPTION
    WHEN OTHERS THEN
      zl_ErrorCenter (SQLCODE, SQLERRM);
  end CancelRequest;


  -----------------------------------------------------------------------------
  -- 功    能：删除报告信息
  -----------------------------------------------------------------------------
  PROCEDURE DeleteReport
  (
    医嘱id_IN 病人医嘱发送.医嘱ID%TYPE
  )Is
     v_Count         Number;
     v_主医嘱ID      病人医嘱发送.医嘱ID%Type;
  Begin
    -- 提取主医嘱ID ，防止因为传入部位医嘱，导致报告保存出错
    Select Nvl(相关id, ID) As 组id Into v_主医嘱ID From 病人医嘱记录 Where ID = 医嘱id_In;

    --先清除报告
    ClearPacsReport(v_主医嘱ID);

    Zl_影像报告标记_Clear(v_主医嘱ID);

    --先检查是否已经出院的住院病人，已经预出院的检查申请，删除报告后不更改执行状态
    Select Count(*) Into v_Count From 病人医嘱记录 a, 病案主页 b
    Where  a.病人ID=b.病人ID And a.主页ID = b.主页ID And b.出院日期 Is Not Null And a.Id = v_主医嘱ID;

    If v_Count =0 Then
       --删除报告，则取消医嘱完成状态
       Update 病人医嘱发送
       Set 执行状态 = 0, 执行过程 = 2
       Where 医嘱id In (Select ID From 病人医嘱记录 Where (ID = v_主医嘱ID Or 相关id = v_主医嘱ID))
             And 执行状态 = 1;
    End If;

  EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
  END DeleteReport;



  -----------------------------------------------------------------------------
  -- 功    能：清除报告数据
  -----------------------------------------------------------------------------
  PROCEDURE ClearPacsReport
  (
    医嘱id_IN 病人医嘱发送.医嘱ID%TYPE
  )IS
  BEGIN
    --其它相关表有级联删除功能,会随电子病历记录一并删除
    Delete 电子病历记录 Where Id In (Select 病历ID From 病人医嘱报告 Where 医嘱ID=医嘱id_IN);
  EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
  END ClearPacsReport;


  -----------------------------------------------------------------------------
  -- 功    能：接收检验/检查申请
  -----------------------------------------------------------------------------
  Procedure RecevieRequest
  (
    医嘱id_IN 病人医嘱发送.医嘱ID%TYPE,
    执行间_IN 病人医嘱发送.执行间%TYPE:=Null,
    检查号_IN 影像检查记录.检查号%TYPE:=NULL,
    检查设备_IN 影像检查记录.检查设备%TYPE:=Null,
    身高_IN 影像检查记录.身高%TYPE:=Null,
    体重_IN 影像检查记录.体重%TYPE:=Null,
    检查技师_IN 影像检查记录.检查技师%TYPE:=Null,
    执行时间_IN 病人医嘱发送.安排时间%TYPE:=Null,
    执行说明_IN 病人医嘱发送.执行说明%TYPE:=NULL,
    单独执行_In   Number := 0,
    执行部门ID_IN 部门表.id%Type := 0
    --参数：医嘱ID_IN=单独执行的医嘱ID。
    --      单独执行_In=检查医嘱组合是否采用对每个项目分散单独执行的方式
  ) Is

    Cursor c_AdviceInfo Is
       Select ID, 相关id, Nvl(相关id, ID) As 组id, 诊疗类别, 病人来源,执行科室ID From 病人医嘱记录 Where ID = 医嘱id_In;
    r_AdviceInfo c_AdviceInfo%Rowtype;

    v_原检查号 影像检查记录.检查号%Type;
    v_新检查号 影像检查记录.检查号%Type;
    v_姓名   影像检查记录.姓名%Type;
    v_英文名 影像检查记录.英文名%Type;
    v_影像类别 影像检查记录.影像类别%Type;
    v_主发送号 影像检查记录.发送号%Type;
    v_发送号 影像检查记录.发送号%Type;
    v_病人来源 病人医嘱记录.病人来源%Type;
    v_人员编号 人员表.编号%Type;
    v_人员姓名 人员表.姓名%Type;
    v_Count Number;
    v_Error Varchar2(255);
    Err_Custom Exception;

  Begin
    --提取医嘱的主医嘱ID，及组ID
    Open c_AdviceInfo;
         Fetch c_AdviceInfo
               Into r_AdviceInfo;
    Close c_AdviceInfo;

    --先检查是否已经出院的住院病人，已经预出院或者出院的检查申请，不允许开始检查和审核费用
    Select Count(*) Into v_Count From 病人医嘱记录 a, 病案主页 b
    Where  a.病人ID=b.病人ID And a.主页ID = b.主页ID And (b.出院日期 Is Not Null Or b.状态 = 3)
       And a.Id = r_AdviceInfo.组id;

    If v_Count >0 Then
      v_Error := '住院病人已经出院或者预出院，无法开始检查。';
      Raise Err_Custom;
    End If;

    --开始执行医嘱
    If Nvl(单独执行_In, 0) = 1 Then
       -- 单个部位医嘱单独执行
       Update 病人医嘱发送
       Set 首次时间 = Sysdate, 末次时间 = Sysdate,执行状态 =3,执行间 = 执行间_In, 安排时间 = 执行时间_IN,
           执行说明 = 执行说明_IN
       Where 医嘱ID = 医嘱id_In;
    Else
       Update 病人医嘱发送
       Set 首次时间 = Sysdate,末次时间 = Sysdate, 执行状态 = 3,执行间 = 执行间_In,安排时间 = 执行时间_IN,
           执行说明 = 执行说明_IN
       Where 医嘱ID In (Select ID From 病人医嘱记录 Where (ID = r_AdviceInfo.组ID Or 相关ID = r_AdviceInfo.组ID));
    End If;

    --处理操作员姓名和编号，如果 检查技师_IN 为空，则填写 user
    If 检查技师_IN Is Null Then
       v_人员姓名 := User;
       v_人员编号 := User;
    Else
       Begin
            Select 编号,姓名 Into v_人员编号,v_人员姓名 From 人员表 a,部门人员 b
            Where a.Id = b.人员ID And b.部门ID=r_AdviceInfo.执行科室ID And a.别名=检查技师_IN And Rownum =1;
       Exception
            When Others Then
                 v_人员姓名 := User;
                 v_人员编号 := User;
       End;
    End If;
    --处理费用
    Select 发送号 Into v_发送号 From 病人医嘱发送 Where 医嘱ID = 医嘱id_IN;
    zl_影像费用执行(医嘱id_IN, v_发送号, 2,单独执行_In,v_人员编号,v_人员姓名,执行部门ID_IN);

    --提取主医嘱相关信息
    Select A.发送号,C.姓名,zlspellcode(C.姓名) 英文名,D.操作类型,B.病人来源
    Into v_主发送号,v_姓名,v_英文名,v_影像类别,v_病人来源
    From 病人医嘱发送 A,病人医嘱记录 B,病人信息 C, 诊疗项目目录 D
    Where A.医嘱ID=B.id And B.id = r_AdviceInfo.组ID And B.病人ID = C.病人ID And B.诊疗项目ID = D.ID;

    --处理检查号
    If 检查号_IN Is Null Then --没传检查号则由HIS跟据类别生成新检查号
      begin
        Select /*+ rule */ 检查号 Into v_原检查号 From 影像检查记录 Where 医嘱id = r_AdviceInfo.组ID;
      Exception
        When Others Then
          Select 最大号码+1 Into v_新检查号 From 影像检查类别 Where 编码=v_影像类别;
      End;
    End If;

    Update /*+ RULE */ 影像检查记录
    Set 影像类别 = v_影像类别, 检查号 = NVL(Nvl(检查号_In, v_原检查号),v_新检查号), 姓名 = v_姓名, 英文名 = v_英文名, 身高 = 身高_In,
        体重 = 体重_In, 检查设备 = 检查设备_In, 检查技师 = 检查技师_In
    Where 医嘱id = r_AdviceInfo.组ID;

    If Sql%Rowcount = 0 Then
      Insert Into 影像检查记录(医嘱id, 发送号, 影像类别, 检查号, 姓名, 英文名, 身高, 体重, 检查设备, 检查技师)
      Values(r_AdviceInfo.组ID, v_主发送号, v_影像类别, NVL(Nvl(检查号_In, v_原检查号),v_新检查号),
             v_姓名, v_英文名, 身高_In, 体重_In, 检查设备_In, 检查技师_In);
    End If;

    If v_新检查号 Is NOT Null Then
      Update 影像检查类别 Set 最大号码 = v_新检查号 Where 编码 = v_影像类别;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end RecevieRequest;


  -----------------------------------------------------------------------------
  -- 功    能：发送报告文本信息
  -----------------------------------------------------------------------------
  procedure SendReport
  (
    医嘱id_IN       病人医嘱发送.医嘱ID%TYPE,
    报告所见_IN     电子病历内容.内容文本%TYPE,
    报告建议_IN     电子病历内容.内容文本%TYPE,
    报告医生_IN     电子病历记录.创建人%TYPE,
    审核医生_IN     影像检查记录.复核人%TYPE := Null,
    执行部门ID_IN 部门表.id%Type := 0
  )Is

    --提取病人医嘱及报告的相关信息
    CURSOR c_Advice(v_组ID Number) IS
        Select E.Id,E.病人来源,E.病人ID,E.主页ID,E.婴儿,E.病人科室ID,E.文件id, E.病历种类,E.病历名称,F.病历ID,E.执行科室ID
        From (Select C.ID,C.病人来源,C.病人ID,C.主页ID,C.婴儿,C.病人科室ID,C.文件id, D.种类 病历种类, D.名称 病历名称,C.执行科室ID
          From (Select A.ID,A.病人来源,A.病人ID,A.主页ID,A.婴儿,A.病人科室ID, B.病历文件id 文件id,A.执行科室ID
                     From 病人医嘱记录 A, 病历单据应用 B
                     Where A.Id=v_组ID And A.诊疗项目id = B.诊疗项目id(+) And B.应用场合(+) = Decode(A.病人来源, 2, 2, 4, 4, 1)) C,病历文件列表 D
          Where C.文件id = D.Id(+)) E,病人医嘱报告 F
        Where E.Id=F.医嘱ID(+);

    --查找文件的组成元素
    CURSOR c_File(v_File number) IS
        Select A.Id, A.文件id, A.父id, A.对象序号, A.对象类型, A.对象标记, A.保留对象, A.对象属性, A.内容行次,
               A.内容文本, A.是否换行, A.预制提纲id, A.复用提纲, A.使用时机, A.诊治要素id, A.替换域, A.要素名称,
               A.要素类型, A.要素长度, A.要素小数, A.要素单位, A.要素表示, A.输入形态, A.要素值域
        From 病历文件结构 A
        Where A.文件id = v_File
        Order By A.对象序号;

    Cursor c_Report(v_电子病历记录ID Number) Is
        Select /*+ rule */ B.Id, A.内容文本
               From 电子病历内容 A, 电子病历内容 B
               Where A.文件id = v_电子病历记录ID And Nvl(A.定义提纲id, 0) <> 0 And
                     (A.内容文本 like '%所见%' Or A.内容文本 like '%建议%' Or A.内容文本 like '%报告医生%' ) And
                     B.父id = A.Id And B.是否换行 = 1;

    Cursor c_ExecutAdvice(v_组ID Number) Is
         Select 医嘱ID,发送号 From 病人医嘱记录 a,病人医嘱发送 b
         Where a.ID=b.医嘱ID And (a.id =v_组ID Or a.相关ID =v_组ID ) And b.执行状态 = 3;
    r_ExecutAdvice c_ExecutAdvice%Rowtype;

    r_Advice      c_Advice%Rowtype;
    v_病历id      电子病历内容.文件ID%Type;
    v_病历内容id  电子病历内容.Id%Type;
    v_病历内容idNew  电子病历内容.Id%Type;
    v_对象序号    电子病历内容.对象序号%Type;
    v_父ID        电子病历内容.父ID%Type;
    v_内容文本    电子病历内容.内容文本%Type;
    v_定义提纲ID  电子病历内容.定义提纲ID%Type;
    --v_格式内容    电子病历格式.内容%Type;
    v_Error         Varchar2(255);
    Err_Custom      Exception;
    v_Count         Number;
    v_主医嘱ID      病人医嘱发送.医嘱ID%Type;
    v_人员编号      人员表.编号%Type;
    v_人员姓名      人员表.姓名%Type;
    v_报告时间     电子病历记录.保存时间%Type;
  Begin

    -- 提取主医嘱ID ，防止因为传入部位医嘱，导致报告保存出错
    Select Nvl(相关id, ID) As 组id Into v_主医嘱ID From 病人医嘱记录 Where ID = 医嘱id_In;

    Open c_Advice(v_主医嘱ID);
      Fetch c_Advice Into r_Advice;

    If Nvl(r_Advice.文件ID,0)=0 Then
        v_Error:='本次检查项目没有对应相关的检查报告，请与管理员联系！';
        Raise Err_Custom;
    Else
        If Nvl(r_Advice.病历id,0)>0 Then  ----产生过报告
            --找出检查已填写的报告提纲中含有'%所见%','%描述%,'%建议%','%意见%',并用传入的参数更新
            For r_Report In c_Report(r_Advice.病历id) Loop
                If r_Report.内容文本 like '%所见%' Then
                    Update 电子病历内容 Set 内容文本=报告所见_IN Where ID=r_Report.Id;
                Elsif r_Report.内容文本 like '%建议%' Then
                    Update 电子病历内容 Set 内容文本=报告建议_IN Where ID=r_Report.Id;
                Elsif r_Report.内容文本 like '%报告医生%' Then
                    Update 电子病历内容 Set 内容文本=报告医生_IN Where ID=r_Report.Id;
                --Elsif r_Report.内容文本 like '%报告时间%' Then
                    --Update 电子病历内容 Set 内容文本=报告建议_IN Where ID=r_Report.Id;
                End If;
            End Loop;
            --更新保存时间
            Update 电子病历记录 Set 完成时间=Sysdate,保存人=报告医生_IN,保存时间=Sysdate Where ID=r_Advice.病历id;
        Else
            --产生电子病历记录
            Select 电子病历记录_ID.Nextval Into v_病历id From Dual;
            Insert Into 电子病历记录
              (Id, 病人来源, 病人id, 主页id, 婴儿, 科室id, 病历种类, 文件id, 病历名称, 创建人, 创建时间, 完成时间,
               保存人, 保存时间, 最后版本, 签名级别)
            Values
              (v_病历id, r_Advice.病人来源, r_Advice.病人id, r_Advice.主页id, r_Advice.婴儿, r_Advice.病人科室id,
               r_Advice.病历种类, r_Advice.文件id, r_Advice.病历名称, 报告医生_IN, Sysdate, Sysdate, 报告医生_IN, Sysdate, 1, 2);

            --产生医嘱报告记录
            Insert Into 病人医嘱报告 (医嘱ID,病历ID) Values(v_主医嘱ID,v_病历ID);
            --插入报告时间
            Select to_char(a.保存时间,'yyyy-mm-dd hh24:mi')  Into v_报告时间 From 电子病历记录 a,病人医嘱报告 b Where a.id=b.病历id And b.医嘱id=医嘱id_In;
            --新产生报告内容
            For r_File In c_File(r_Advice.文件ID) Loop
                Select 电子病历内容_ID.Nextval Into v_病历内容id From Dual;
                If nvl(v_对象序号,0)=0 Then
                   v_对象序号:=r_File.对象序号;
                Else
                   v_对象序号:=v_对象序号+1;
                End If;

                If NVL(r_File.父ID,0)<>0 And (r_File.内容文本 like '%所见%' Or r_File.内容文本 like '%描述%') Then--所见定义行(非提纲)
                     v_内容文本:=chr(32)||chr(32)||chr(32)||报告所见_IN || Chr(13) || Chr(13);
                     v_定义提纲ID:=0;
                Elsif NVL(r_File.父ID,0)<>0 And (r_File.内容文本 like '%建议%' Or r_File.内容文本 like '%意见%') Then--建议定义行(非提纲)
                     v_内容文本:=chr(32)||chr(32)||chr(32)||报告建议_IN || Chr(13) || Chr(13);
                     v_定义提纲ID:=0;
                Elsif Nvl(r_File.父id, 0) <> 0 And (r_File.内容文本 Like '%报告医生%') Then--报告医生定义行(非提纲)
 				            v_内容文本   := '报告医生: ' || 报告医生_IN || Chr(13) || Chr(13);
					          v_定义提纲id := 0;
                Elsif Nvl(r_File.父id, 0) <> 0 And (r_File.内容文本 Like '%报告时间%') Then--报告时间定义行(非提纲)
 				            v_内容文本   := '报告时间: ' || v_报告时间 || Chr(13) || Chr(10)||'【此报告仅供临床科室查看结果,具体以打印的纸质报告单为准！】';
					          v_定义提纲id := 0;
                Elsif nvl(r_File.对象类型,0)=1 And NVL(r_File.父ID,0)=0 Then--提纲定义行
                     v_父ID:=v_病历内容id;
                     v_内容文本:=r_File.内容文本;
                     v_定义提纲ID:=r_File.id;
                Elsif nvl(r_File.对象类型,0)=4 And r_File.要素名称 Is Not Null Then  --自动替换要素
                     v_内容文本:=zl_replace_element_value(r_File.要素名称,r_Advice.病人ID,r_Advice.主页ID,r_Advice.病人来源,r_Advice.Id);
                     v_定义提纲ID:=0;
                Else
                    v_内容文本:=r_File.内容文本;
                    v_定义提纲ID:=0;
                End If;

                --报告内容单独写一行
                If NVL(r_File.父ID,0)<>0 And (r_File.内容文本 like '%所见%' Or r_File.内容文本 like '%建议%') Then--先写提纲显示名称，再写内容，同时对象序号发生变化
                   Select 电子病历内容_ID.Nextval Into v_病历内容idNew From Dual;
                   v_对象序号 := v_对象序号 + 1;
                    Insert Into 电子病历内容
                      (Id, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次,
                       内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域, 要素名称, 要素类型,
                       要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
                    Values
                      (v_病历内容idNew, v_病历id, 0, 0, Decode(v_定义提纲id, 0, v_父id, Null), v_对象序号, r_File.对象类型,
                       r_File.对象标记, r_File.保留对象, 0, Null, v_内容文本, r_File.是否换行,
                       r_File.预制提纲id, r_File.复用提纲, r_File.使用时机, r_File.诊治要素id, r_File.替换域,
                       r_File.要素名称, r_File.要素类型, r_File.要素长度, r_File.要素小数, r_File.要素单位,
                       r_File.要素表示, r_File.输入形态, r_File.要素值域, Decode(v_定义提纲id, 0, Null, v_定义提纲id));
                    v_对象序号 := v_对象序号 - 1;
                    v_内容文本:=r_File.内容文本;
                End If;

                Insert Into 电子病历内容
                  (Id, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次,
                   内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度,
                   要素小数, 要素单位, 要素表示, 输入形态, 要素值域, 定义提纲id)
                Values
                  (v_病历内容id, v_病历id, 1, 0, Decode(v_定义提纲id, 0, v_父id, Null), v_对象序号, r_File.对象类型,
                   r_File.对象标记, r_File.保留对象, r_File.对象属性, Null, v_内容文本, r_File.是否换行, r_File.预制提纲id,
                   r_File.复用提纲, r_File.使用时机, r_File.诊治要素id, r_File.替换域, r_File.要素名称, r_File.要素类型,
                   r_File.要素长度, r_File.要素小数, r_File.要素单位, r_File.要素表示, r_File.输入形态, r_File.要素值域,
                   Decode(v_定义提纲id, 0, Null, v_定义提纲id));
             End Loop;

        /* 因电子病历格式中含了内容文字格式，此种方法导入之后内容文字将不可见
        Select 内容 Into v_格式内容 From 病历文件格式 Where 文件ID=r_Advice.文件ID;
        Insert Into 电子病历格式 (文件ID,内容) Values (v_病历id,v_格式内容);
        */

        End If;

        --先检查是否已经出院的住院病人，已经预出院的检查申请，添加报告后不更改执行状态
        Select Count(*) Into v_Count From 病人医嘱记录 a, 病案主页 b
        Where  a.病人ID=b.病人ID And a.主页ID = b.主页ID And b.出院日期 Is Not Null And a.Id = v_主医嘱ID;

        If v_Count =0 Then
           --只对已经接收申请，正在执行的医嘱才更新，更新为 完成状态，审核过程
           Update 病人医嘱发送 Set 执行状态=1, 执行过程=6, 完成时间=sysdate
           Where 医嘱id in(select id from 病人医嘱记录 where id= v_主医嘱ID or 相关id=v_主医嘱ID);
                 --And 执行状态 = 3 ;--不需要执行“接收申请”，因此不需要判断执行状态，直接更新已完成

           --处理操作员姓名和编号，如果 检查技师_IN 为空，则填写 user
           If 报告医生_IN Is Null Then
              v_人员姓名 := User;
              v_人员编号 := User;
           Else
               Begin
                    Select 编号,姓名 Into v_人员编号,v_人员姓名 From 人员表 a,部门人员 b
                    Where a.Id = b.人员ID And b.部门ID=r_Advice.执行科室ID And a.别名=报告医生_IN And Rownum =1;
               Exception
                    When Others Then
                         v_人员姓名 := User;
                         v_人员编号 := User;
               End;
           End If;

           --处理费用
           For r_ExecutAdvice In c_ExecutAdvice(v_主医嘱ID) Loop
               zl_影像费用执行(r_ExecutAdvice.医嘱ID,r_ExecutAdvice.发送号 , 6,1,v_人员编号,v_人员姓名,执行部门ID_IN);
           End Loop;
        End If;

        Update 影像检查记录 set 报告人=报告医生_IN, 复核人=审核医生_IN where 医嘱id=v_主医嘱ID;
      End If;
      Close c_Advice;
    Exception
      When Err_Custom Then
        Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
      When Others Then
        Zl_ErrorCenter(Sqlcode, Sqlerrm);
  end SendReport;


  -----------------------------------------------------------------------------
  -- 功    能：删除心电报告信息
  -----------------------------------------------------------------------------
  PROCEDURE DeleteElectrocardioReport
  (
    医嘱id_IN  病人医嘱发送.医嘱ID%TYPE
  )Is
     v_Count         Number;
     v_主医嘱ID      病人医嘱发送.医嘱ID%Type;
  Begin
    -- 提取主医嘱ID ，防止因为传入部位医嘱，导致报告保存出错
    Select Nvl(相关id, ID) As 组id Into v_主医嘱ID From 病人医嘱记录 Where ID = 医嘱id_In;

    --先清除报告
    Delete 电子病历记录 Where Id In (Select 病历ID From 病人医嘱报告 Where 医嘱ID=v_主医嘱ID);


    --先检查是否已经出院的住院病人，已经预出院的检查申请，删除报告后不更改执行状态
    Select Count(*) Into v_Count From 病人医嘱记录 a, 病案主页 b
    Where  a.病人ID=b.病人ID And a.主页ID = b.主页ID And b.出院日期 Is Not Null And a.Id = v_主医嘱ID;

    If v_Count =0 Then
       --删除报告，则取消医嘱完成状态
       Update 病人医嘱发送
       Set 执行状态 = 3, 执行过程 = 2
       Where 医嘱id In (Select ID From 病人医嘱记录 Where (ID = v_主医嘱ID Or 相关id = v_主医嘱ID))
             And 执行状态 = 1;
    End If;

  EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
  END DeleteElectrocardioReport;



  -----------------------------------------------------------------------------
  -- 功    能：发送心电报告信息
  -----------------------------------------------------------------------------
  procedure SendElectrocardioReport
  (
    医嘱id_IN       病人医嘱发送.医嘱ID%TYPE,
    报告标题_IN     电子病历内容.内容文本%TYPE,
    诊断结果_IN     电子病历内容.内容文本%TYPE,
    诊断建议_IN     电子病历内容.内容文本%TYPE,
    报告医生_IN     电子病历记录.创建人%TYPE,
    审核医生_IN     影像检查记录.复核人%TYPE := Null
  )is
    cursor c_AdviceInf(v_组ID Number) is
           select A.病人来源,A.病人ID,A.主页ID,A.开嘱科室ID,A.姓名,A.性别,A.年龄,B.门诊号,B.住院号
           from 病人医嘱记录 A,病人信息 B
           where A.病人ID = b.病人id and a.id =v_组ID;

    r_AdviceInf  c_AdviceInf%RowType;
    v_病历ID     病人医嘱报告.病历ID%Type;
    v_StudyInf   Varchar2(2048);

    v_Count         Number;
    v_主医嘱ID      病人医嘱发送.医嘱ID%Type;
    v_格式ID        病历文件列表.id%Type;
    v_编号          病历文件列表.编号%type;
    v_格式名     varchar2(255);

    v_Error      varchar2(255);
    Err_Custom   Exception;
    v_父ID   number;
  begin
    -- 提取主医嘱ID ，防止因为传入部位医嘱，导致报告保存出错
    Select Nvl(相关id, ID) As 组id Into v_主医嘱ID From 病人医嘱记录 Where ID = 医嘱id_In;

    open c_AdviceInf(v_主医嘱ID);
    fetch c_AdviceInf into r_AdviceInf;

    if c_AdviceInf%Rowcount = 0 then
       close c_AdviceInf;
       v_Error := '病人医嘱记录获取失败，请检查传递的医嘱ID是否正确。';
       raise Err_Custom;
    end if;


    begin
      select Id into v_格式ID from 病历文件列表 where 名称='心电报告格式';
    exception
      When Others Then
        v_格式ID := 0;
    end;

    --构造心电报告格式
    if v_格式ID = 0 then
       select 病历文件列表_ID.NEXTVAL into v_格式ID from dual;
       select max(编号)+1 into v_编号 from 病历文件列表;

       v_格式名 :=  '心电报告格式';

       insert into 病历文件列表(id,种类,编号,名称,页面)
       values(v_格式ID, 7, v_编号, v_格式名, v_编号);


       --格式说明： 纸张大小;方向;纸张高度;纸张宽度;左边距;右边距;上边距;下边距;文字背景色;纸张背景色;显示页号
       --原始格式：9;1;16840;11907;849;849;1587;1417;10070188;16777215;1;
       --当需要将病历页面增大时，请在plSql中单独执行如下语句：
       --update 病历页面格式 set 格式='256;1;16840;16442;283;283;482;283;10070188;16777215;1' where 名称='心电报告格式'
       insert into 病历页面格式(种类,编号,名称,格式)
       values(7, v_编号,v_格式名,'256;1;20840;11907;240;240;1587;1417;10070188;16777215;1');
    end if;

    --生成病历ID
    select 电子病历记录_ID.Nextval into v_病历ID from dual;



    --插入电子病历记录
    insert into 电子病历记录(ID,病人来源,病人ID,主页ID,科室ID,病历种类,文件ID,病历名称,完成时间,保存人,保存时间,创建人,创建时间)
      select v_病历ID,r_AdviceInf.病人来源,r_AdviceInf.病人ID,r_AdviceInf.主页ID,r_AdviceInf.开嘱科室ID,
             7,v_格式ID,'心电检查报告单',sysdate,报告医生_IN,sysdate,报告医生_IN,sysdate
      from 病人医嘱记录
      where ID=医嘱ID_IN;



    --构造电子病历内容
    v_StudyInf :=lpad(' ',trunc((60 - length(报告标题_IN))/2), ' ') || 报告标题_IN;
    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行)
    values(电子病历内容_ID.NEXTVAL,v_病历ID,1,2, v_StudyInf ,1); --报告标题


    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行) values(电子病历内容_ID.NEXTVAL,v_病历ID,2,2, '' ,1); --空行
    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行) values(电子病历内容_ID.NEXTVAL,v_病历ID,3,2, '' ,1); --空行

----  吴泽宁 2011/10/18 插入提纲，解决体检无法提取诊断结果的问题
--   -- 对象类型=1,对象标记=6,内容文本='诊断意见',预制提纲=-8 复用提纲=0
--    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,对象标记,内容文本,预制提纲ID,复用提纲,是否换行)
--    values(电子病历内容_ID.NEXTVAL,v_病历ID,16,1,6,'诊断意见' ,-8,0,1);
--    commit;
--****--
    v_StudyInf := '  姓名： ' || rpad(nvl(to_char( r_AdviceInf.姓名), ' '),15, ' ') || '  性别： ' || rpad(nvl(to_char( r_AdviceInf.性别),' '),15,' ') || ' 年龄： ' || r_AdviceInf.年龄;
    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行)
    values(电子病历内容_ID.NEXTVAL,v_病历ID,4,2, v_StudyInf ,1); --检查信息


    v_StudyInf := '门诊号： ' || rpad(nvl(to_char(r_AdviceInf.门诊号), ' '),15, ' ') || '住院号： ' || rpad(nvl(to_char(r_AdviceInf.住院号), ' '),15,' ') || ' 日期： ' || to_char(sysdate, 'yyyy-mm-dd');
    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行)
    values(电子病历内容_ID.NEXTVAL,v_病历ID,5,2, v_StudyInf ,1); --检查信息


    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行) values(电子病历内容_ID.NEXTVAL,v_病历ID,6,2, '' ,1); --空行
    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行) values(电子病历内容_ID.NEXTVAL,v_病历ID,7,2, '' ,1); --空行


    if not(诊断结果_IN is null) then
      v_StudyInf := '诊断结果：';
      insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行)
      values(电子病历内容_ID.NEXTVAL,v_病历ID,8,2, v_StudyInf ,1); --检查信息

      insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行)
      values(电子病历内容_ID.NEXTVAL,v_病历ID,9,2, 诊断结果_IN ,1); --检查信息
    end if;

    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行) values(电子病历内容_ID.NEXTVAL,v_病历ID,10,2, '' ,1); --空行


--- 吴泽宁 2011/10/18 修改  增加诊断结果 父ID 关联
    begin
  Select id  into v_父ID from  电子病历内容 where 文件id=v_病历ID and 内容文本='诊断意见';
    exception
      when others then
        v_父ID := null;
    end;


    if not(诊断建议_IN is null) then
      v_StudyInf := '诊断建议：';
      insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行)
      values(电子病历内容_ID.NEXTVAL,v_病历ID,11,2, v_StudyInf ,1); --检查信息

      insert into 电子病历内容(ID, 文件ID,终止版,父ID,对象序号,对象类型,对象属性,内容文本,是否换行)
      values(电子病历内容_ID.NEXTVAL,v_病历ID,0,v_父ID,12,2,0, 诊断建议_IN ,1); --检查信息
    end if;

    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行) values(电子病历内容_ID.NEXTVAL,v_病历ID,13,2, '' ,1); --空行

    v_StudyInf := '------------------------------------------------------------------------';
    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,内容文本,是否换行) values(电子病历内容_ID.NEXTVAL,v_病历ID,14,2, v_StudyInf ,1); --图形分割


    --关联医嘱报告
    insert into 病人医嘱报告(医嘱ID,病历ID) values(v_主医嘱ID,v_病历ID);



    --先检查是否已经出院的住院病人，已经预出院的检查申请，添加报告后不更改执行状态
    Select Count(*) Into v_Count From 病人医嘱记录 a, 病案主页 b
    Where  a.病人ID=b.病人ID And a.主页ID = b.主页ID And b.出院日期 Is Not Null And a.Id = v_主医嘱ID;

    If v_Count =0 Then
        --只对已经接收申请，正在执行的医嘱才更新，更新为 完成状态，审核过程
        Update 病人医嘱发送 Set 执行状态=1, 执行过程=6, 完成时间=sysdate
        Where 医嘱id in(select id from 病人医嘱记录 where id= v_主医嘱ID or 相关id=v_主医嘱ID) ;  --And 执行状态 = 3 --由于不需要执行“接收申请”的过程，因此不需要判断执行状态
    End If;

    --更新影像检查记录的报告信息
    update 影像检查记录 set 报告人=报告医生_IN where 医嘱ID=医嘱ID_IN;

    close c_AdviceInf;

    exception
      when others then
        zl_ErrorCenter(sqlCode, sqlErrm);
  end SendElectrocardioReport;



    -----------------------------------------------------------------------------
  -- 功    能：添加心电图像
  -----------------------------------------------------------------------------
  function AddElectrocardioReportImage
  (
    医嘱id_IN       病人医嘱发送.医嘱ID%TYPE
  )return number is
  PRAGMA AUTONOMOUS_TRANSACTION;
    v_主医嘱ID      病人医嘱发送.医嘱ID%Type;
    v_病历ID        电子病历记录.ID%type;
    v_内容ID        电子病历内容.ID%Type;
    v_序号          电子病历内容.对象序号%Type;
  begin
    -- 提取主医嘱ID ，防止因为传入部位医嘱，导致报告保存出错
    Select Nvl(相关id, ID) As 组id Into v_主医嘱ID From 病人医嘱记录 Where ID = 医嘱id_In;

    select 病历ID into v_病历ID from 病人医嘱报告 where 医嘱ID= v_主医嘱ID;

    select 电子病历内容_ID.NEXTVAL into v_内容ID from dual;
    select nvl(max(对象序号),0) + 1 into v_序号 from 电子病历内容 where 文件ID=v_病历ID;

    insert into 电子病历内容(ID, 文件ID,对象序号,对象类型,对象属性,是否换行) values(v_内容ID,v_病历ID,v_序号,5, '2;0;0;0;0;12150;13020;1;1;1;0' ,1);

    commit;

    return v_内容ID;

    exception
      when others then
        zl_ErrorCenter(sqlCode, sqlErrm);

  end AddElectrocardioReportImage;



  -----------------------------------------------------------------------------
  -- 功    能：清除报告附件
  -----------------------------------------------------------------------------
  procedure ClearReportAffix
  (
    病历id_In       电子病历附件.病历ID%TYPE,
    附件标记_IN     电子病历附件.创建人%TYPE
  )is
  begin
      Delete From 电子病历附件 Where 病历id = 病历id_In And 创建人 = 附件标记_IN;
    Exception
      When Others Then
        Zl_ErrorCenter(Sqlcode, Sqlerrm);
  end ClearReportAffix;

  -----------------------------------------------------------------------------
  -- 功    能：添加报告附件
  -----------------------------------------------------------------------------
  Procedure AddReportAffix
  (
    病历id_In In 电子病历附件.病历id%Type,
    文件名_In In 电子病历附件.文件名%Type,
    大小_In   In 电子病历附件.大小%Type,
    附件标记_IN in  电子病历附件.创建人%TYPE
  )is
  begin
    Insert Into 电子病历附件(病历id, 序号, 文件名, 大小, 创建人, 日期)
    Values(病历id_In, 10000, 文件名_In, 大小_In, 附件标记_IN, Sysdate);

    Exception
      When Others Then
        Zl_ErrorCenter(Sqlcode, Sqlerrm);
  end;


end b_PacsInterface;
/
