CREATE SEQUENCE 电子票据异常记录_ID START WITH 1;
Create Table 电子票据异常记录(
	ID Number(18),
	操作场景 number(2),
	业务类型 number(2),
	记录标志 number(2),
	单据号 varchar2(20),
	业务ID number(18),
	电子票据id number(18),
	病人ID number(18),
	姓名 varchar2(100),
	性别 varchar2(4),
	年龄 varchar2(20),
	门诊号 number(18),
	住院号 number(18),
	是否换开 number(2),
	票据信息 CLOB,
	操作员编号 varchar2(6),
	操作员姓名 varchar2(50),
	登记时间 Date)
 TABLESPACE zl9Expense;

Alter Table 电子票据异常记录 Add Constraint 电子票据异常记录_PK Primary Key(ID) Using Index Tablespace zl9Indexhis;
Alter table 电子票据异常记录 Add Constraint 电子票据异常记录_UQ_业务ID Unique(业务ID,记录标志,操作场景,业务类型)  Using Index Tablespace zl9Indexhis; 
Alter Table 电子票据异常记录 Add Constraint 电子票据异常记录_FK_病人ID Foreign Key (病人ID) References 病人信息(病人ID);
CREATE INDEX 电子票据异常记录_IX_登记时间 ON 电子票据异常记录(登记时间) TABLESPACE zl9Indexhis; 
CREATE INDEX 电子票据异常记录_IX_病人ID ON 电子票据异常记录(病人ID) TABLESPACE zl9Indexhis;
CREATE INDEX 电子票据异常记录_IX_电子票据id ON 电子票据异常记录(电子票据id) TABLESPACE zl9Indexhis;
CREATE INDEX 电子票据异常记录_IX_单据号 ON 电子票据异常记录(单据号,操作场景) TABLESPACE zl9Indexhis; 



Create Or Replace Procedure Zl_电子票据使用记录_Insert
(
  Id_In         In 电子票据使用记录.Id%Type,
  票种_In       In 电子票据使用记录.票种%Type,
  结算id_In     In 电子票据使用记录.结算id%Type,
  病人id_In     In 电子票据使用记录.病人id%Type,
  姓名_In       In 电子票据使用记录.姓名%Type,
  性别_In       In 电子票据使用记录.性别%Type,
  年龄_In       In 电子票据使用记录.年龄%Type,
  门诊号_In     In 电子票据使用记录.门诊号%Type,
  住院号_In     In 电子票据使用记录.住院号%Type,
  票据金额_In   In 电子票据使用记录.票据金额%Type,
  开票点_In     In 电子票据使用记录.开票点%Type,
  系统来源_In   In 电子票据使用记录.系统来源%Type,
  生成时间_In   In 电子票据使用记录.生成时间%Type,
  备注_In       In 电子票据使用记录.备注%Type,
  操作员编号_In In 电子票据使用记录.操作员编号%Type,
  操作员姓名_In In 电子票据使用记录.操作员姓名%Type,
  登记时间_In   In 电子票据使用记录.登记时间%Type,
  原票据id_In   In 电子票据使用记录.原票据id%Type := Null,
  退款id_In     In 电子票据使用记录.原票据id%Type := Null,
  代码_In       In 电子票据使用记录.代码%Type := Null,
  号码_In       In 电子票据使用记录.号码%Type := Null,
  检验码_In     In 电子票据使用记录.检验码%Type := Null,
  凭证代码_In   In 电子票据使用记录.凭证代码%Type := Null,
  凭证号码_In   In 电子票据使用记录.凭证号码%Type := Null,
  凭证检验码_In In 电子票据使用记录.凭证检验码%Type := Null,
  Url内网_In    In 电子票据使用记录.Url内网%Type := Null,
  Url外网_In    In 电子票据使用记录.Url外网%Type := Null
) As
  n_记录状态 电子票据使用记录.记录状态%Type;
Begin
  n_记录状态 := 1;

  Insert Into 电子票据使用记录
    (ID, 票种, 记录状态, 结算id, 病人id, 姓名, 性别, 年龄, 门诊号, 住院号, 票据金额, 生成时间, 原票据id, 退款id, 开票点, 系统来源, 代码, 号码, 检验码, 凭证代码, 凭证号码, 凭证检验码,
     Url内网, Url外网, 备注, 操作员编号, 操作员姓名, 登记时间)
  Values
    (Id_In, 票种_In, n_记录状态, 结算id_In, Decode(Nvl(病人id_In, 0), 0, Null, 病人id_In), 姓名_In, 性别_In, 年龄_In,
     Decode(Nvl(门诊号_In, 0), 0, Null, 门诊号_In), Decode(Nvl(住院号_In, 0), 0, Null, 住院号_In), 票据金额_In, 生成时间_In, 原票据id_In,
     退款id_In, 开票点_In, 系统来源_In, 代码_In, 号码_In, 检验码_In, 凭证代码_In, 凭证号码_In, 凭证检验码_In, Url内网_In, Url外网_In, 备注_In, 操作员编号_In,
     操作员姓名_In, Nvl(登记时间_In, Sysdate));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据使用记录_Insert;
/


Create Or Replace Procedure Zl_电子票据使用记录_Delete
(
  Id_In           In 电子票据使用记录.Id%Type,
  开票点_In       In 电子票据使用记录.开票点%Type,
  系统来源_In     In 电子票据使用记录.系统来源%Type,
  生成时间_In     In 电子票据使用记录.生成时间%Type,
  备注_In         In 电子票据使用记录.备注%Type,
  操作员编号_In   In 电子票据使用记录.操作员编号%Type,
  操作员姓名_In   In 电子票据使用记录.操作员姓名%Type,
  登记时间_In     In 电子票据使用记录.登记时间%Type,
  原电子票据id_In In 电子票据使用记录.Id%Type,
  代码_In         In 电子票据使用记录.代码%Type := Null,
  号码_In         In 电子票据使用记录.号码%Type := Null,
  检验码_In       In 电子票据使用记录.检验码%Type := Null,
  凭证代码_In     In 电子票据使用记录.凭证代码%Type := Null,
  凭证号码_In     In 电子票据使用记录.凭证号码%Type := Null,
  凭证检验码_In   In 电子票据使用记录.凭证检验码%Type := Null,
  Url内网_In      In 电子票据使用记录.Url内网%Type := Null,
  Url外网_In      In 电子票据使用记录.Url外网%Type := Null
) As
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_是否换开 电子票据使用记录.是否换开%Type;
Begin

  Update 电子票据使用记录 Set 记录状态 = 3 Where ID = 原电子票据id_In Returning Nvl(是否换开, 0) Into n_是否换开;
  If Sql%NotFound Then
    v_Err_Msg := '未找到原始的电子票据信息，不能作废操作!';
    Raise Err_Item;
  End If;
  If Nvl(n_是否换开, 0) = 1 Then
    --当前电子票据已经换开纸质票据
    v_Err_Msg := '当前电子票据已经换开纸质票据,需要先冲红纸质票据后才能作废电子发票!';
    Raise Err_Item;
  End If;

  Insert Into 电子票据使用记录
    (ID, 票种, 记录状态, 结算id, 病人id, 姓名, 性别, 年龄, 门诊号, 住院号, 代码, 号码, 检验码, 票据金额, 生成时间, 原票据id, 打印id, 是否换开, 纸质发票号, 开票点, 系统来源, 备注,
     操作员编号, 操作员姓名, 登记时间, 退款id, 凭证代码, 凭证号码, 凭证检验码, Url内网, Url外网)
    Select Id_In, 票种, 2, 结算id, 病人id, 姓名, 性别, 年龄, 门诊号, 住院号, Nvl(代码_In, 代码), Nvl(号码_In, 号码), Nvl(检验码_In, 检验码), 票据金额,
           生成时间_In, 原电子票据id_In, 打印id, 是否换开, 纸质发票号, Nvl(开票点_In, 开票点) As 开票点, Nvl(系统来源_In, 系统来源) As 系统来源,
           Nvl(备注_In, 备注) As 备注, 操作员编号_In, 操作员姓名_In, 登记时间_In, 退款id, Nvl(凭证代码_In, 凭证代码), Nvl(凭证号码_In, 凭证号码),
           Nvl(凭证检验码_In, 凭证检验码), Nvl(Url内网_In, Url内网), Nvl(Url外网_In, Url外网)
    From 电子票据使用记录
    Where ID = 原电子票据id_In;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据使用记录_Delete;
/
 
Create Or Replace Procedure Zl_电子票据异常记录_Insert
(
  异常id_In     电子票据异常记录.Id%Type,
  操作场景_In   电子票据异常记录.操作场景%Type,
  业务类型_In   电子票据异常记录.业务类型%Type,
  记录标志_IN   电子票据异常记录.记录标志%Type,
  单据号_In     电子票据异常记录.单据号%Type,
  业务id_In     电子票据异常记录.业务id%Type,
  电子票据id_In 电子票据异常记录.电子票据id%Type,
  病人id_In     电子票据异常记录.病人id%Type,
  姓名_In       电子票据异常记录.姓名%Type,
  性别_In       电子票据异常记录.性别%Type,
  年龄_In       电子票据异常记录.年龄%Type,
  门诊号_In     电子票据异常记录.门诊号%Type,
  住院号_In     电子票据异常记录.住院号%Type,
  操作员编号_In 电子票据异常记录.操作员编号%Type,
  操作员姓名_In 电子票据异常记录.操作员姓名%Type,
  登记时间_In   电子票据异常记录.登记时间%Type,
  是否换开_In   电子票据异常记录.是否换开%Type := 0
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:插入电子票据异常记录
  -- 入参 
  --   操作场景_In:1-医疗卡发卡;2-病人信息登记;3-病人入院登记
  --   业务类型_In:1-医疗卡,2-预交
  --   记录标志_IN:0-开具电子票据;1-冲红电子票据;2-纸质票据;3-作废纸质票据
  --   单据号_In:业务类型=1:表示医疗卡费用NO,业务类型=2:表示预交款NO
  --   业务ID:原结算ID或原预交ID
  ------------------------------------------------------------------------------------------------------------------------------
  n_异常id   Number(18);
  d_登记时间 电子票据异常记录.登记时间%Type;
Begin

  n_异常id   := 异常id_In;
  d_登记时间 := 登记时间_In;
  If d_登记时间 Is Null Then
    d_登记时间 := Sysdate;
  End If;

  If Nvl(n_异常id, 0) = 0 Then
    Select 电子票据异常记录_Id.Nextval Into n_异常id From Dual;
  End If;

  Insert Into 电子票据异常记录
    (ID, 操作场景, 业务类型, 记录标志, 单据号, 业务id, 电子票据id, 病人id, 姓名, 性别, 年龄, 门诊号, 住院号, 是否换开, 操作员编号, 操作员姓名, 登记时间)
  Values
    (n_异常id, 操作场景_In, 业务类型_In, 记录标志_IN, 单据号_In, 业务id_In, 电子票据id_In, 病人id_In, 姓名_In, 性别_In, 年龄_In, 门诊号_In, 住院号_In,
     是否换开_In, 操作员编号_In, 操作员姓名_In, d_登记时间);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据异常记录_Insert;
/

Create Or Replace Procedure Zl_电子票据异常记录_Modify
(
  异常id_In   电子票据异常记录.Id%Type, 
  票据信息_In Clob, 
  是否换开_In 电子票据异常记录.是否换开%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:更新电子票据异常信息
  -- 入参 
  --     票据信息_In:更新类别_in=0时，更新电子票据信息字段;更新类别_in=1时更新纸质票据字段 
  --     是否换开_In:NULL-表示不更新是否换开字段;否则更新成当前传入值
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin 
    Update 电子票据异常记录
    Set 是否换开 = Nvl(是否换开_In, 是否换开), 票据信息 = 票据信息_In
    Where ID = Nvl(异常id_In, 0);
    If Sql%NotFound Then
      v_Err_Msg := '未找到需要更新的票据信息，请检查!';
      Raise Err_Item;
    End If; 
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据异常记录_Modify;
/

Create Or Replace Procedure Zl_电子票据异常记录_Delete(异常id_In 电子票据异常记录.Id%Type) As
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:删除异常记录
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin
  Delete 电子票据异常记录 Where ID = Nvl(异常id_In, 0);
  If Sql%NotFound Then
    v_Err_Msg := '未找到需要删除的电子票据信息，请检查!';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_电子票据异常记录_Delete;
/

Create Or Replace Procedure Zl_Exsesvr_Addeinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：增加电子票据信息
  --入参：Json_In:格式
  --  input      
  --    balance_id          N  1  结算ID
  --    balance_delid       N     退款ID:退款开具红票时有效：目前只有预交款有效,填写的是退款预交ID
  --    einvoice_id         N  1  电子票据ID
  --    operator_code       C  1  操作员编号
  --    operator_name       C  1  操作员姓名
  --    create_time         C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  --    pati_info           C     病人信息
  --      pati_id           N  1  病人ID
  --      pati_pageid       N     主页ID
  --      pati_name         C  1  姓名
  --      pati_sex          C  1  性别
  --      pati_age          C  1  年龄
  --      outpatient_num    C  1  门诊号
  --      inpatient_num     C  1  住院号
  --    einvoce_info        C     电子票据信息
  --      invoice_type      N  1  票种：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
  --      placeCode         C  1  开票点编码
  --      inv_total         N  1  开票金额
  --      inv_oldid         N     原票据ID
  --      sys_source        C  1  系统来源
  --      demo              C  1  备注
  --      einvoice_code     C  1  电子票据代码
  --      einvoice_no       C  1  电子票据号码
  --      einvoice_random   C  1  电子校验码
  --      voucher_code      C  1  预交金凭证代码
  --      voucher_no        C  1  预交金凭证号码
  --      voucher_random    C  1  预交金凭证校验码
  --      create_time       C  1  电子票据生成时间:yyyymmddhh24miss
  --      picture_url       C  1  电子票据H5页面URL
  --      picture_neturl    C  1  电子票据外网H5页面URL
  --      qrcode            C  1  电子票据二维码图片数据:该值已Base64编码，解析时需要Base64解码,图片格式为:PNG
  --    --出参: Json_Out,格式如下
  --  output
  --    code                N 1 应答吗：0-失败；1-成功
  --    message             C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  n_Id         电子票据使用记录.Id%Type;
  n_票种       电子票据使用记录.票种%Type;
  n_记录状态   电子票据使用记录.记录状态%Type;
  n_结算id     电子票据使用记录.结算id%Type;
  n_病人id     电子票据使用记录.病人id%Type;
  v_姓名       电子票据使用记录.姓名%Type;
  v_性别       电子票据使用记录.性别%Type;
  v_年龄       电子票据使用记录.年龄%Type;
  n_门诊号     电子票据使用记录.门诊号%Type;
  n_住院号     电子票据使用记录.住院号%Type;
  v_代码       电子票据使用记录.代码%Type;
  v_号码       电子票据使用记录.号码%Type;
  v_检验码     电子票据使用记录.检验码%Type;
  v_凭证代码   电子票据使用记录.凭证代码%Type;
  v_凭证号码   电子票据使用记录.凭证号码%Type;
  v_凭证检验码 电子票据使用记录.凭证检验码%Type;
  n_票据金额   电子票据使用记录.票据金额%Type;
  v_生成时间   电子票据使用记录.生成时间%Type;
  v_Url内网    电子票据使用记录.Url内网%Type;
  v_Url外网    电子票据使用记录.Url外网%Type;
  c_二维码     Clob;
  n_原票据id   电子票据使用记录.原票据id%Type;
  n_退款id     电子票据使用记录.退款id%Type;
  v_备注       电子票据使用记录.备注%Type;
  v_开票点     电子票据使用记录.开票点%Type;
  v_系统来源   电子票据使用记录.系统来源%Type;
  v_操作员编号 电子票据使用记录.操作员编号%Type;
  v_操作员姓名 电子票据使用记录.操作员姓名%Type;
  d_登记时间   电子票据使用记录.登记时间%Type;

  n_记录状态 Number(2);
  j_Input    PLJson;
  j_Json     PLJson;
  j_Temp     PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_结算id     := j_Json.Get_Number('balance_id');
  n_退款id     := j_Json.Get_Number('balance_delid');
  n_Id         := j_Json.Get_Number('einvoice_id');
  v_操作员编号 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  d_登记时间   := Nvl(To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  --读取病人信息

  If Not j_Json.Exist('pati_info') Then
  
    Json_Out := zlJsonOut('无病人信息，不能增加电子票据。');
    Return;
  End If;

  j_Temp   := j_Json.Get_Pljson('pati_info');
  n_病人id := j_Temp.Get_Number('pati_id');
  --n_主页id := j_Temp.Get_Number('pati_pageid');

  v_姓名   := j_Temp.Get_String('pati_name');
  v_性别   := j_Temp.Get_String('pati_sex');
  v_年龄   := j_Temp.Get_String('pati_age');
  n_门诊号 := j_Temp.Get_Number('outpatient_num');
  n_住院号 := j_Temp.Get_Number('inpatient_num');

  If Not j_Json.Exist('einvoce_info') Then
  
    Json_Out := zlJsonOut('无电子票据信息,不能增加电子票据。');
    Return;
  End If;

  --读取电子票据信息
  j_Temp := PLJson();
  j_Temp := j_Json.Get_Pljson('einvoce_info');
  --票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  n_票种       := Nvl(j_Temp.Get_Number('invoice_type'), 1);
  v_开票点     := j_Temp.Get_String('placeCode');
  n_原票据id   := j_Temp.Get_Number('inv_oldid');
  v_系统来源   := j_Temp.Get_String('sys_source');
  v_备注       := j_Temp.Get_String('demo');
  v_代码       := j_Temp.Get_String('einvoice_code');
  v_号码       := j_Temp.Get_String('einvoice_no');
  v_检验码     := j_Temp.Get_String('einvoice_random');
  v_凭证代码   := j_Temp.Get_String('voucher_code');
  v_凭证号码   := j_Temp.Get_String('voucher_no');
  v_凭证检验码 := j_Temp.Get_String('voucher_random');
  n_票据金额   := j_Temp.Get_Number('inv_total');
  v_生成时间   := j_Temp.Get_String('happen_time');
  v_Url内网    := j_Temp.Get_String('picture_url');
  v_Url外网    := j_Temp.Get_String('picture_neturl');
  c_二维码     := j_Temp.Get_Clob('qrcode');

  --增加电子票据信息 
  Zl_电子票据使用记录_Insert(n_Id, n_票种, n_结算id, n_病人id, v_姓名, v_性别, v_年龄, n_门诊号, n_住院号, n_票据金额, v_开票点, v_系统来源, v_生成时间, v_备注,
                     v_操作员编号, v_操作员姓名, d_登记时间, n_原票据id, n_退款id, v_代码, v_号码, v_检验码, v_凭证代码, v_凭证号码, v_凭证检验码, v_Url内网,
                     v_Url外网);
  --更新二维码
  Insert Into 电子票据二维码 (使用记录id, 二维码) Values (n_Id, c_二维码);
  Json_Out := zlJsonOut('成功', 1);
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Addeinvoice;
/



Create Or Replace Procedure Zl_Exsesvr_Deleinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：删除电子票据信息
  --入参：Json_In:格式
  -- input      
  --  einvoice_id  N  1  电子票据ID
  --  operator_code  C  1  操作员编号
  --  operator_name  C  1  操作员姓名
  --  create_time  C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  --  einvoce_info  C    电子票据信息
  --    placeCode  C  1  开票点编码
  --    sys_source  C  1  系统来源
  --    demo  C  1  备注
  --    inv_oldid  N    原票据ID
  --    einvoice_code  C  1  电子票据代码
  --    einvoice_no  C  1  电子票据号码
  --    einvoice_random  C  1  电子校验码
  --    voucher_code  C  1  预交金凭证代码
  --    voucher_no  C  1  预交金凭证号码
  --    voucher_random  C  1  预交金凭证校验码
  --    happen_time  C  1  电子票据生成时间:yyyymmddhh24miss
  --    picture_url  C  1  电子票据H5页面URL
  --    picture_neturl  C  1  电子票据外网H5页面URL
  --    qrcode  C  1  电子票据二维码图片数据:该值已Base64编码，解析时需要Base64解码,图片格式为:PNG
  --    --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  n_Id         电子票据使用记录.Id%Type;
  v_代码       电子票据使用记录.代码%Type;
  v_号码       电子票据使用记录.号码%Type;
  v_检验码     电子票据使用记录.检验码%Type;
  v_凭证代码   电子票据使用记录.凭证代码%Type;
  v_凭证号码   电子票据使用记录.凭证号码%Type;
  v_凭证检验码 电子票据使用记录.凭证检验码%Type;
  v_生成时间   电子票据使用记录.生成时间%Type;
  v_Url内网    电子票据使用记录.Url内网%Type;
  v_Url外网    电子票据使用记录.Url外网%Type;
  c_二维码     Clob;
  n_原票据id   电子票据使用记录.原票据id%Type;
  v_备注       电子票据使用记录.备注%Type;
  v_开票点     电子票据使用记录.开票点%Type;
  v_系统来源   电子票据使用记录.系统来源%Type;
  v_操作员编号 电子票据使用记录.操作员编号%Type;
  v_操作员姓名 电子票据使用记录.操作员姓名%Type;
  d_登记时间   电子票据使用记录.登记时间%Type;

  j_Input PLJson;
  j_Json  PLJson;
  j_Temp  PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Id         := j_Json.Get_Number('einvoice_id');
  v_操作员编号 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  d_登记时间   := Nvl(To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  If Not j_Json.Exist('einvoce_info') Then
  
    Json_Out := zlJsonOut('无电子票据信息,不能增加电子票据。');
    Return;
  End If;

  --读取电子票据信息
  j_Temp     := PLJson();
  j_Temp     := j_Json.Get_Pljson('einvoce_info');
  v_开票点   := j_Temp.Get_String('placeCode');
  v_系统来源 := j_Temp.Get_String('sys_source');
  v_备注     := j_Temp.Get_String('demo');
  n_原票据id := j_Temp.Get_Number('inv_oldid');

  v_代码       := j_Temp.Get_String('einvoice_code');
  v_号码       := j_Temp.Get_String('einvoice_no');
  v_检验码     := j_Temp.Get_String('einvoice_random');
  v_凭证代码   := j_Temp.Get_String('voucher_code');
  v_凭证号码   := j_Temp.Get_String('voucher_no');
  v_凭证检验码 := j_Temp.Get_String('voucher_random');
  v_生成时间   := j_Temp.Get_String('happen_time');
  v_Url内网    := j_Temp.Get_String('picture_url');
  v_Url外网    := j_Temp.Get_String('picture_neturl');
  c_二维码     := j_Temp.Get_Clob('qrcode');

  Zl_电子票据使用记录_Delete(n_Id, v_开票点, v_系统来源, v_生成时间, v_备注, v_操作员编号, v_操作员姓名, d_登记时间, n_原票据id, v_代码, v_号码, v_检验码, v_凭证代码,
                     v_凭证号码, v_凭证检验码, v_Url内网, v_Url外网);
  --更新二维码
  Insert Into 电子票据二维码 (使用记录id, 二维码) Values (n_Id, c_二维码);
  Json_Out := zlJsonOut('成功', 1);
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Deleinvoice;
/


Create Or Replace Procedure Zl_Exsesvr_Savepaperinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --功能：增加纸质票据使用信息
  --入参：Json_In:格式
  --   input      
  --    oper_mode           N  1  操作方式:0-换开;1-重新换开;2-作废票据;3-回收票据
  --    einvoice_id         N  1  电子票据ID
  --    operator_code       C  1  操作员编号
  --    operator_name       C  1  操作员姓名
  --    create_time         C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  --    paperinv_info       C     纸质票据信息:存在多条时，请按操作顺序上传(避免数据错误)
  --      inv_occasion      N  1  应用场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
  --      invoice_type      N  1  票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --      inv_red           N     是否红票:1-红票;0-非红票
  --      invoice_no        C  1  发票号
  --      inv_total         N  1  发票金额
  --      recv_id           N     领用id
  --    --出参: Json_Out,格式如下
  --  output
  --    code          N 1 应答吗：0-失败；1-成功
  --    message       C 1 应答消息：失败时返回具体的错误信息
  ---------------------------------------------------------------------------
  n_Id 电子票据使用记录.Id%Type;
  -- v_操作员编号 电子票据使用记录.操作员编号%Type;
  v_操作员姓名 电子票据使用记录.操作员姓名%Type;
  d_登记时间   电子票据使用记录.登记时间%Type;
  n_结算id     电子票据使用记录.结算id%Type;
  v_发票号     票据使用明细.号码%Type;
  n_发票金额   票据使用明细.票据金额%Type;
  n_领用id     票据使用明细.领用id%Type;
  n_操作方式   Number(2);
  n_应用场合   Number(2);
  n_票种       Number(2);
  n_是否红票   Number(2);
  j_Input      PLJson;
  j_Json       PLJson;
  j_Temp       PLJson;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  --操作方式:0-换开;1-重新换开;2-作废票据;3-回收票据
  n_操作方式 := j_Json.Get_Number('oper_mode');
  n_Id       := j_Json.Get_Number('einvoice_id');
  --v_操作员编号 := j_Json.Get_String('operator_code');
  v_操作员姓名 := j_Json.Get_String('operator_name');
  d_登记时间   := Nvl(To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  If Not j_Json.Exist('paperinv_info') Then
  
    Json_Out := zlJsonOut('无纸质票据信息。');
    Return;
  End If;
  Select Max(结算id) Into n_结算id From 电子票据使用记录 Where ID = n_Id;
  If Nvl(n_结算id, 0) = 0 Then
  
    Json_Out := zlJsonOut('传入的电子票据无效!');
    Return;
  End If;

  j_Temp := j_Json.Get_Pljson('paperinv_info');
  --应用场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
  n_应用场合 := j_Temp.Get_Number('inv_occasion');
  --票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  n_票种 := j_Temp.Get_Number('invoice_type');
  --是否红票:1-红票;0-非红票
  n_是否红票 := Nvl(j_Temp.Get_Number('inv_red'), 0);
  v_发票号   := j_Temp.Get_String('invoice_no');
  n_发票金额 := j_Temp.Get_Number('inv_total');
  n_领用id   := j_Temp.Get_Number('recv_id');

  --纸质票据处理
  --操作方式_In:0-换开;1-重新换开;2-作废票据;3-回收票据
  Zl_纸质票据使用_Update(n_应用场合, n_票种, n_结算id, n_Id, v_发票号, n_发票金额, n_领用id, v_操作员姓名, d_登记时间, n_操作方式, 0, n_是否红票);

  Json_Out := zlJsonOut('成功', 1);
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Savepaperinvoice;
/

Create Or Replace Procedure Zl_Exsesvr_Getstarteinvoices
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取启用电子票据业务
  --入参：Json_In:NULL
  --     
  --出参: Json_Out,格式如下
  --output      
  --  code  C  1  应答码：0-失败；1-成功
  --  message  C  1  应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  data[]      启用站点列表
  --    occasion  N  1  场合:1-收费,2-预交,3-结帐,4-挂号
  --    client_name  C  1  站点名
  ---------------------------------------------------------------------------

  v_Output Varchar2(32767);
  c_Output Clob;
Begin
  --解析入参
  For c_卡类别 In (Select 场合, 站点 From 电子票据站点控制) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'occasion', c_卡类别.场合, 1, 1);
    zlJsonPutValue(v_Output, 'client_name', c_卡类别.站点, 0, 2);
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"成功","data":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"成功","data":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getstarteinvoices;
/
CREATE OR REPLACE Procedure Zl_Exsesvr_Geteinvoicecode
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --功能：获取开票点编号
  --入参：json格式
  --input
  --   operator_id    N  1  操作员ID
  --   ssite          C  1  客户端
  --出参：json格式
  --Json_Out
  --  code            C  1  应答码：0-失败；1-成功
  --  message         C  1  应答消息： 成功时返回处方No，[数组] 失败时返回具体的错误信息
  --  einvoice_code   C  1  开票点编码
  -------------------------------------------------------------------------------------------------
  n_操作员id   票据开票点对照.人员id%Type;
  v_客户端     票据开票点对照.客户端%Type;
  v_开票点编码 电子票据开票点.编码%Type;
  j_Input      PLJson;
  j_Json       PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_操作员id := j_Json.Get_Number('operator_id');
  v_客户端   := j_Json.Get_String('ssite');

  --按收费员+客户端对码
  For r_开票点 In (Select b.编码 As 开票点编码
                From 票据开票点对照 A, 电子票据开票点 B
                Where a.开票点id = b.Id And Nvl(b.撤档时间, Sysdate + 1) >= Sysdate And a.人员id = n_操作员id And a.客户端 = v_客户端) Loop
    v_开票点编码 := r_开票点.开票点编码;
    Json_Out     := '{"output":{"code":1,"message": "成功","einvoice_code":"' || Nvl(v_开票点编码, '') || '"}}';
    Return;
  End Loop;

  --按收费员对码
  For r_开票点 In (Select Nvl(a.人员id, 0) As 人员id, Nvl(a.客户端, '-') As 客户端, b.编码 As 开票点编码
                From 票据开票点对照 A, 电子票据开票点 B
                Where a.开票点id = b.Id And Nvl(b.撤档时间, Sysdate + 1) >= Sysdate And a.人员id = n_操作员id And a.客户端 = '-') Loop
    v_开票点编码 := r_开票点.开票点编码;
    Json_Out     := '{"output":{"code":1,"message": "成功","einvoice_code":"' || Nvl(v_开票点编码, '') || '"}}';
  End Loop;

  --按客户端对码
  For r_开票点 In (Select Nvl(a.人员id, 0) As 人员id, Nvl(a.客户端, '-') As 客户端, b.编码 As 开票点编码
                From 票据开票点对照 A, 电子票据开票点 B
                Where a.开票点id = b.Id And Nvl(b.撤档时间, Sysdate + 1) >= Sysdate And a.人员id = 0 And a.客户端 = v_客户端) Loop
    v_开票点编码 := r_开票点.开票点编码;
    Json_Out     := '{"output":{"code":1,"message": "成功","einvoice_code":"' || Nvl(v_开票点编码, '') || '"}}';
  End Loop;

  Json_Out := '{"output":{"code":1,"message": "成功","einvoice_code":"' || Null || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoicecode;
/
CREATE OR REPLACE Procedure Zl_Exsesvr_Addeinvoiceerrdata
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:插入电子票据异常记录
  --入参：Json_In:
  --input
  --  err_id              N 1 异常ID
  --  business_type       N 1 业务类型:1-医疗卡,2-预交 
  --  business_id         N 1 业务id:结算ID或预交ID  
  --  occasion            N 1 操作场景:1-医疗卡发卡;2-病人信息登记;3-病人入院登记
  --  record_sign         N 1 记录标志:0-开具电子票据;1-冲红电子票据;2-纸质票据;3-作废纸质票据  
  --  einvoice_id         N 1 电子票据id
  --  pati_id             N 1 病人id
  --  pati_name           C 1 姓名
  --  pati_sex            C 1 性别
  --  pati_age            C 1 年龄
  --  outpatient_num      C 1 门诊号
  --  inpatient_num       C 1 住院号
  --  err_no              C 1 单据号
  --  operator_code       C 1 操作员编号
  --  operator_name       C 1 操作员姓名
  --  create_time         C   登记时间  格式为:yyyy-mm-dd hh24:mi:ss
  --  is_turn             N   是否换开
  --出参: Json_Out,格式如下
  --output      
  --  code                C  1 应答码：0-失败；1-成功
  --  message             C  1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  ------------------------------------------------------------------------------------------------------------------------------

  j_Input      PLJson;
  j_Json       PLJson;
  n_异常id     电子票据异常记录.Id%Type;
  n_操作场景   电子票据异常记录.操作场景%Type;
  n_业务类型   电子票据异常记录.业务类型%Type;
  n_记录标志   电子票据异常记录.记录标志%Type;
  v_单据号     电子票据异常记录.单据号%Type;
  n_业务id     电子票据异常记录.业务id%Type;
  n_电子票据id 电子票据异常记录.电子票据id%Type;
  n_病人id     电子票据异常记录.病人id%Type;
  v_姓名       电子票据异常记录.姓名%Type;
  v_性别       电子票据异常记录.性别%Type;
  v_年龄       电子票据异常记录.年龄%Type;
  n_门诊号     电子票据异常记录.门诊号%Type;
  n_住院号     电子票据异常记录.住院号%Type;
  v_操作员编号 电子票据异常记录.操作员编号%Type;
  v_操作员姓名 电子票据异常记录.操作员姓名%Type;
  d_登记时间   电子票据异常记录.登记时间%Type;
  n_是否换开   电子票据异常记录.是否换开%Type;
Begin

  j_Input      := PLJson(Json_In);
  j_Json       := j_Input.Get_Pljson('input');
  n_异常id     := Pljson_Ext.Get_Number(j_Json, 'err_id');
  n_业务类型   := Pljson_Ext.Get_Number(j_Json, 'business_type');
  n_业务id     := Pljson_Ext.Get_Number(j_Json, 'business_id');
  n_操作场景   := Pljson_Ext.Get_Number(j_Json, 'occasion');
  n_记录标志   := Pljson_Ext.Get_Number(j_Json, 'record_sign');
  v_操作员编号 := Pljson_Ext.Get_String(j_Json, 'operator_code');
  v_操作员姓名 := Pljson_Ext.Get_String(j_Json, 'operator_name');
  d_登记时间   := To_Date(Pljson_Ext.Get_String(j_Json, 'create_time'), 'yyyy-mm-dd hh24:mi:ss');
  If d_登记时间 Is Null Then
    d_登记时间 := Sysdate;
  End If;
  v_单据号     := Pljson_Ext.Get_String(j_Json, 'err_no');
  n_电子票据id := Pljson_Ext.Get_Number(j_Json, 'einvoice_id');
  n_病人id     := Pljson_Ext.Get_Number(j_Json, 'pati_id');
  v_姓名       := Pljson_Ext.Get_String(j_Json, 'pati_name');
  v_性别       := Pljson_Ext.Get_String(j_Json, 'pati_sex');
  v_年龄       := Pljson_Ext.Get_String(j_Json, 'pati_age');
  n_门诊号     := Pljson_Ext.Get_Number(j_Json, 'outpatient_num');
  n_住院号     := Pljson_Ext.Get_Number(j_Json, 'inpatient_num');
  n_是否换开   := Pljson_Ext.Get_Number(j_Json, 'is_turn');

  Zl_电子票据异常记录_Insert(n_异常id, n_操作场景, n_业务类型, n_记录标志, v_单据号, n_业务id, n_电子票据id, n_病人id, v_姓名, v_性别, v_年龄, n_门诊号, n_住院号,
                     v_操作员编号, v_操作员姓名, d_登记时间, n_是否换开);
  Json_Out := '{"output":{"code":1,"message": "成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Addeinvoiceerrdata;
/
CREATE OR REPLACE Procedure Zl_Exsesvr_Modifyeinvoerrdata
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:插入电子票据异常记录
  --入参：Json_In:
  --input
  --  err_id              N 1 异常id
  --  einvoice_info       C 1 票据信息
  --  is_turn             N   是否换开
  --出参: Json_Out,格式如下
  --output      
  --  code                C  1 应答码：0-失败；1-成功
  --  message             C  1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  ------------------------------------------------------------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  n_异常id   电子票据异常记录.Id%Type;
  c_票据信息 电子票据异常记录.票据信息%Type;
  n_是否换开 电子票据异常记录.是否换开%Type;
Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_异常id := Pljson_Ext.Get_Number(j_Json, 'err_id');
  Begin
    c_票据信息 := j_Json.Get_Clob('einvoice_info');
  Exception
    When Others Then
      Json_Out := '{"output":{"code":0,"message": "失败,未传入票据信息,请检查!"}}';
      Return;
  End;
  n_是否换开 := Pljson_Ext.Get_Number(j_Json, 'is_turn');
  Zl_电子票据异常记录_Modify(n_异常id, c_票据信息, n_是否换开);

  Json_Out := '{"output":{"code":1,"message": "成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Modifyeinvoerrdata;
/
CREATE OR REPLACE Procedure Zl_Exsesvr_Deleteeinvoerrdata
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ------------------------------------------------------------------------------------------------------------------------------
  --功能:插入电子票据异常记录
  --入参：Json_In:
  --input
  --  err_id              N 1 异常id
  --出参: Json_Out,格式如下
  --output      
  --  code                C  1 应答码：0-失败；1-成功
  --  message             C  1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  ------------------------------------------------------------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  n_异常id 电子票据异常记录.Id%Type;
Begin

  j_Input  := PLJson(Json_In);
  j_Json   := j_Input.Get_Pljson('input');
  n_异常id := Pljson_Ext.Get_Number(j_Json, 'err_id');

  Zl_电子票据异常记录_Delete(n_异常id);
  Json_Out := '{"output":{"code":1,"message": "成功"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Deleteeinvoerrdata;
/
Create Or Replace Procedure Zl_Exsesvr_Geteinvoiceerrdata
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取电子票据异常记录
  --入参：Json_In:
  --input
  --  business_type       N 1 业务类型:1-医疗卡,2-预交 
  --  business_id         N 1 业务id:结算ID或预交ID  
  --  occasion            N 1 操作场景:1-医疗卡发卡;2-病人信息登记;3-病人入院登记
  --  record_sign         N 0 记录标志:0-开具电子票据;1-冲红电子票据;2-纸质票据;3-作废纸质票据  
  --出参: Json_Out,格式如下
  --output      
  --  code                C  1 应答码：0-失败；1-成功
  --  message             C  1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  err_id              N  1 电子票据异常id
  --  record_sign         N  0 记录标志:0-开具电子票据;1-冲红电子票据;2-纸质票据;3-作废纸质票据  
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_异常id       电子票据异常记录.Id%Type;
  n_操作场景     电子票据异常记录.操作场景%Type;
  n_业务类型     电子票据异常记录.业务类型%Type;
  n_记录标志     电子票据异常记录.记录标志%Type;
  n_业务id       电子票据异常记录.业务id%Type;
  n_记录标志_Out 电子票据异常记录.记录标志%Type;
Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_业务类型 := Pljson_Ext.Get_Number(j_Json, 'business_type');
  n_业务id   := Pljson_Ext.Get_Number(j_Json, 'business_id');
  n_操作场景 := Pljson_Ext.Get_Number(j_Json, 'occasion');
  n_记录标志 := Pljson_Ext.Get_Number(j_Json, 'record_sign');

  If Nvl(n_业务类型, 0) = 0 Or Nvl(n_业务id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "失败,传入的业务类型或业务id为0"}}';
    Return;
  End If;

  If Nvl(n_操作场景, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "失败,传入的操作场景为0"}}';
    Return;
  End If;

  If n_记录标志 Is Null Then
    Select Max(ID), Max(记录标志)
    Into n_异常id, n_记录标志_Out
    From 电子票据异常记录
    Where 业务类型 = n_业务类型 And 业务id = n_业务id And 操作场景 = n_操作场景;
  Else
    Select Max(ID)
    Into n_异常id
    From 电子票据异常记录
    Where 业务类型 = n_业务类型 And 业务id = n_业务id And 记录标志 = n_记录标志 And 操作场景 = n_操作场景;
  End If;

  Json_Out := '{"output":{"code":1,"message": "成功","err_id":' || Nvl(n_异常id, 0) || ',"record_sign":' ||
              Nvl(n_记录标志_Out, 0) || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoiceerrdata;
/
Create Or Replace Procedure Zl_Exsesvr_Geteinvoicedata
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:根据结帐ID,获取有效的电子票据ID
  --入参：Json_In:
  --input
  --  fun_oper            N 1 操作类型：0-根据票种和结算id获取电子票据ID；1-根据电子票据ID获取 是否换开、纸质发票号、结算id
  --  blnc_id             N   结算ID(电子票据使用记录.结算id)
  --  inv_type            N   票种:1-收费, 2-预交, 3-结帐, 4-挂号;5-医疗发卡 
  --  einvoice_id         N   电子票据ID
  --出参: Json_Out,格式如下
  --output      
  --  code                C  1 应答码：0-失败；1-成功
  --  message             C  1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  einvoice_id         N    有效的电子票据ID(操作类型=0时返回)
  --  blnc_id             N    结算ID(操作类型=1时返回)
  --  is_turn             N    是否换开(操作类型=1时返回)
  --  inv_no              N    纸质票号(操作类型=1时返回)
  ---------------------------------------------------------------------------
  j_Input          PLJson;
  j_Json           PLJson;
  n_操作类型       Number(2);
  n_票种           电子票据使用记录.票种%Type;
  n_是否换开       电子票据使用记录.是否换开%Type;
  v_纸质发票号     电子票据使用记录.纸质发票号%Type;
  n_结算id         电子票据使用记录.结算id%Type;
  n_结算id_Out     电子票据使用记录.结算id%Type;
  n_电子票据id     电子票据使用记录.Id%Type;
  n_电子票据id_Out 电子票据使用记录.Id%Type;
Begin
  --解析入参

  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_操作类型 := Nvl(Pljson_Ext.Get_Number(j_Json, 'fun_oper'), 0);

  If n_操作类型 = 0 Then
    --根据票种和结算id获取电子票据ID
    n_结算id := Pljson_Ext.Get_Number(j_Json, 'blnc_id');
    n_票种   := Pljson_Ext.Get_Number(j_Json, 'inv_type');
    If (Nvl(n_结算id, 0) = 0 Or Nvl(n_票种, 0) = 0) Then
      Json_Out := '{"output":{"code":0,"message": "失败,传入的结算id或场合为0"}}';
      Return;
    End If;
  
    Select Max(ID)
    Into n_电子票据id_Out
    From 电子票据使用记录
    Where 结算id = n_结算id And 票种 = n_票种 And 记录状态 = 1 And Nvl(原票据id, 0) = 0;
  
    Json_Out := '{"output":{"code":1,"message": "成功","einvoice_id":' || Nvl(n_电子票据id_Out, 0) || '}}';
  Else
    --根据电子票据ID获取 是否换开、纸质发票号、结算id
    n_电子票据id := Pljson_Ext.Get_Number(j_Json, 'einvoice_id');
    If Nvl(n_电子票据id, 0) = 0 Then
      Json_Out := '{"output":{"code":0,"message": "失败,传入的电子票据ID为0"}}';
      Return;
    End If;
  
    Select Max(是否换开), Max(纸质发票号), Max(结算id)
    Into n_是否换开, v_纸质发票号, n_结算id_Out
    From 电子票据使用记录
    Where ID = n_电子票据id;
  
    Json_Out := '{"output":{"code":1,"message": "成功","blnc_id":' || Nvl(n_结算id_Out, 0) || ',"is_turn":' ||
                Nvl(n_是否换开, 0) || ',"inv_no":"' || v_纸质发票号 || '"}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoicedata;
/
Create Or Replace Procedure Zl_Exsesvr_Checkiseinvoice
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:根据传入的票种和结算ID,检查当前结算是否启用了有电子票据
  --入参：Json_In:
  --input
  --  blnc_id             N 1 结算ID(电子票据使用记录.id)
  --  inv_type            N 1 票种:1-收费, 2-预交, 3-结帐, 4-挂号;5-医疗发卡 
  --出参: Json_Out,格式如下
  --output      
  --  code                C  1 应答码：0-失败；1-成功
  --  message             C  1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --  is_einvoice         N  1 是否启用电子票据:1-启用;0:未启用
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_票种     电子票据使用记录.票种%Type;
  n_结算id   电子票据使用记录.Id%Type;
  n_Einvoice Number(2);
Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_结算id := Pljson_Ext.Get_Number(j_Json, 'blnc_id');
  n_票种   := Pljson_Ext.Get_Number(j_Json, 'inv_type');

  If Nvl(n_结算id, 0) = 0 Or Nvl(n_票种, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "失败,传入的结算id或场合为0"}}';
    Return;
  End If;

  If n_票种 = 2 Then
    --预交记录
    Select Max(预交电子票据) Into n_Einvoice From 病人预交记录 Where Mod(记录性质, 10) = 1 And ID = n_结算id;
  Else
    Select Max(是否电子票据) Into n_Einvoice From 病人预交记录 Where 结帐id = n_结算id;
  End If;

  Json_Out := '{"output":{"code":1,"message": "成功","is_einvoice":' || Nvl(n_Einvoice, 0) || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkiseinvoice;
/
Create Or Replace Procedure Zl_Exsesvr_Geteinvoiceinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --功能:获取电子票据信息
  --入参：Json_In:
  --input
  --err_id              N 1 异常ID
  --出参: Json_Out,格式如下
  --output      
  --code                C  1 应答码：0-失败；1-成功
  --message             C  1 应答消息：成功时返回成功信息，失败时返回具体的错误信息
  --记录标志=0 时 返回----------------------------------------
  --data
  --  input          
  --    balance_id        N  1  结算ID
  --    balance_delid     N     退款ID:退款开具红票时有效：目前只有预交款有效,填写的是退款预交ID
  --    einvoice_id       N  1  电子票据ID
  --    operator_code     C  1  操作员编号
  --    operator_name     C  1  操作员姓名
  --    create_time       C  1  登记时间:yyyy-mm-dd hh24:mi:ss
  --    pati_info         C     病人信息
  --      pati_id         N  1  病人ID
  --      pati_pageid     N     主页ID
  --      pati_name       C  1  姓名
  --      pati_sex        C  1  性别
  --      pati_age        C  1  年龄
  --      outpatient_num  C  1  门诊号
  --      inpatient_num   C  1  住院号
  --    einvoce_info      C     电子票据信息
  --      invoice_type    N  1  票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --      placeCode       C  1  开票点编码
  --      inv_total       N  1  开票金额
  --      inv_oldid       N    原票据ID
  --      sys_source      C  1  系统来源
  --      demo            C  1  备注
  --      einvoice_code   C  1  电子票据代码
  --      einvoice_no     C  1  电子票据号码
  --      einvoice_random C  1  电子校验码
  --      voucher_code    C  1  预交金凭证代码
  --      voucher_no      C  1  预交金凭证号码
  --      voucher_random  C  1  预交金凭证校验码
  --      happen_time     C  1  电子票据生成时间:yyyymmddhh24miss
  --      picture_url     C  1  电子票据H5页面URL
  --      picture_neturl  C  1  电子票据外网H5页面URL
  --      qrcode          C  1  电子票据二维码图片数据:该值已Base64编码，解析时需要Base64解码,图片格式为:PNG
  --记录标志=1 时 返回-----------------------------------------
  --data
  --  input         
  --    einvoice_id       N 1 电子票据ID
  --    operator_code     C 1 操作员编号
  --    operator_name     C 1 操作员姓名
  --    create_time       C 1 登记时间:yyyy-mm-dd hh24:mi:ss
  --    einvoce_info      C   电子票据信息
  --      placeCode       C 1 开票点编码
  --      sys_source      C 1 系统来源
  --      demo            C 1 备注
  --      inv_oldid       N   原票据ID
  --      einvoice_code   C 1 电子票据代码
  --      einvoice_no     C 1 电子票据号码
  --      einvoice_random C 1 电子校验码
  --      voucher_code    C 1 预交金凭证代码
  --      voucher_no      C 1 预交金凭证号码
  --      voucher_random  C 1 预交金凭证校验码
  --      happen_time     C 1 电子票据生成时间:yyyymmddhh24miss
  --      picture_url     C 1 电子票据H5页面URL
  --      picture_neturl  C 1 电子票据外网H5页面URL
  --      qrcode          C 1 电子票据二维码图片数据:该值已Base64编码，解析时需要Base64解码,图片格式为:PNG
  --记录标志=2,3 时 返回-------------------------------------------
  --data
  --  input         
  --    oper_mode         N 1 操作方式:0-换开;1-重新换开;2-作废票据;3-回收票据
  --    einvoice_id       N 1 电子票据ID
  --    operator_code     C 1 操作员编号
  --    operator_name     C 1 操作员姓名
  --    create_time       C 1 登记时间:yyyy-mm-dd hh24:mi:ss
  --    paperinv_info     C   纸质票据信息:存在多条时，请按操作顺序上传(避免数据错误)
  --      inv_occasion    N 1 应用场合:1-收费,2-预交(包含押金),3-结帐,4-挂号,5-就诊卡
  --      invoice_type    N 1 票种:1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
  --      inv_red         N   是否红票:1-红票;0-非红票
  --      invoice_no      C 1 发票号
  --      inv_total       N 1 发票金额
  --      recv_id         N   领用id
  ---------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  c_Output Clob;
  n_异常id 电子票据异常记录.Id%Type;

Begin
  --解析入参

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_异常id := Pljson_Ext.Get_Number(j_Json, 'err_id');

  If Nvl(n_异常id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "失败,传入的异常id为0"}}';
    Return;
  End If;

  Begin
    Select 票据信息 Into c_Output From 电子票据异常记录 Where ID = n_异常id;
  Exception
    When Others Then
      Json_Out := '{"output":{"code":0,"message": "失败,根据传入的异常id未找到数据"}}';
      Return;
  End;

  If c_Output Is Null Then
    Json_Out := '{"output":{"code":0,"message": "失败,根据传入的异常id未找到数据"}}';
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message": "成功","data":' || c_Output || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoiceinfo;
/
Create Or Replace Procedure Zl_病人预交记录_Insert_s
(
  Id_In           病人预交记录.Id%Type,
  单据号_In       病人预交记录.No%Type,
  票据号_In       票据使用明细.号码%Type,
  病人id_In       病人预交记录.病人id%Type,
  主页id_In       病人预交记录.主页id%Type,
  姓名_In         病人预交记录.姓名%Type,
  性别_In         病人预交记录.性别%Type,
  年龄_In         病人预交记录.年龄%Type,
  门诊号_In       病人预交记录.门诊号%Type,
  住院号_In       病人预交记录.住院号%Type,
  付款方式名称_In 病人预交记录.付款方式名称%Type,
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
  结帐id_In       病人预交记录.结帐id%Type := Null,
  结算性质_In     病人预交记录.结算性质%Type := Null,
  更新交款余额_In Number := 1,
  操作状态_In     Number := 0,
  关联交易id_In   病人预交记录.关联交易id%Type := Null,
  校对标志_In     病人预交记录.校对标志%Type := Null,
  预交电子票据_In 病人预交记录.预交电子票据%Type := Null,
  险类_In         保险结算记录.险类%Type := Null
) As
  --------------------------------------------------------------------------------------------
  --功能：产生预交记录
  --操作状态_In:0-正常结算,1-保存为异常单据或未生效的单据,2-完成异常结算,3-修正预交记录数据
  --结帐ID_IN:>0时,表示某次结帐时,同步产生的预交记录(结帐终款多余存为预交)
  --更新交款余额_In:0-在 zl_人员缴款余额_Update 中更新(主要是自助充值时防止汇总表锁表)；1-在本过程中更新
  --------------------------------------------------------------------------------------------

  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  v_性质   结算方式.性质%Type;
  v_打印id 票据打印内容.Id%Type;

  v_Date Date;

  n_返回值       病人余额.预交余额%Type;
  n_组id         财务缴款分组.Id%Type;
  n_险类         保险结算记录.险类%Type;
  n_预交电子票据 Number(2);
Begin
  v_Date := 收款时间_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  n_组id := Zl_Get组id(操作员姓名_In);

  n_预交电子票据 := 预交电子票据_In;
  If n_预交电子票据 Is Null Then
    n_险类 := 险类_In;
    If 险类_In Is Null Then
      Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = Id_In And 性质 = 3;
    End If;
    n_预交电子票据 := Zl_Fun_Isstarteinvoice(2, n_险类, 预交类别_In);
  End If;

  Select Max(性质) Into v_性质 From 结算方式 Where 名称 = 结算方式_In;

  --操作状态_In：0-正常结算,1-保存为异常单据,2-完成异常结算,3-修正预交记录数据
  If Nvl(操作状态_In, 0) < 2 Then
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 付款方式名称, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行,
       单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 冲预交, 结算性质, 校对标志, 关联交易id, 交易时间,
       交易人员, 预交电子票据)
    Values
      (Id_In, 单据号_In, Decode(操作状态_In, 1, Null, 票据号_In), 1, Decode(操作状态_In, 1, 0, 1), 病人id_In,
       Decode(主页id_In, 0, Null, 主页id_In), 姓名_In, 性别_In, 年龄_In, Decode(门诊号_In, 0, Null, 门诊号_In),
       Decode(住院号_In, 0, Null, 住院号_In), 付款方式名称_In, Decode(科室id_In, 0, Null, 科室id_In), 金额_In, 结算方式_In, 结算号码_In, v_Date,
       缴款单位_In, 单位开户行_In, 单位帐号_In, 操作员编号_In, 操作员姓名_In, 摘要_In, n_组id, 预交类别_In, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In,
       交易说明_In, 合作单位_In, 结帐id_In, Decode(Nvl(结帐id_In, 0), 0, Null, 0), 结算性质_In, Decode(操作状态_In, 1, 1, Null),
       Decode(关联交易id_In, Null, Id_In, 0, Null, 关联交易id_In), 收款时间_In, 操作员姓名_In, n_预交电子票据);
    If Nvl(卡类别id_In, 0) <> 0 Then
      --自定义过程调用
      Zl_Custom_Balance_Update(Id_In);
    End If;
  End If;

  --操作状态_In：0-正常结算,1-保存为异常单据,2-完成异常结算,3-修正预交记录数据
  If Nvl(操作状态_In, 0) = 1 Then
    --保存为异常单据
    Return;
  End If;
  --操作状态_In：0-正常结算,1-保存为异常单据,2-完成异常结算,3-修正预交记录数据
  If Nvl(操作状态_In, 0) = 0 Or Nvl(操作状态_In, 0) = 2 Then
    Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (Id_In, 病人id_In, 预交类别_In, 金额_In);
    --病人余额(预交余额现收)
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
  End If;
  --更新异常单据或修正预交数据
  --操作状态_In：0-正常结算,1-保存为异常单据,2-完成异常结算,3-修正预交记录数据
  If Nvl(操作状态_In, 0) = 2 Or Nvl(操作状态_In, 0) = 3 Then
    --更新并处理余额
    Update 病人预交记录
    Set 记录状态 = 1, 校对标志 = Decode(操作状态_In, 2, Null, 校对标志_In), 实际票号 = Nvl(票据号_In, 实际票号), 收款时间 = Nvl(v_Date, 收款时间),
        操作员编号 = Nvl(操作员编号_In, 操作员编号), 操作员姓名 = Nvl(操作员姓名_In, 操作员姓名), 缴款组id = Nvl(n_组id, 缴款组id), 交易时间 = Nvl(v_Date, 交易时间),
        交易人员 = Nvl(操作员姓名_In, 交易人员), 科室id = Nvl(科室id_In, 科室id), 金额 = Nvl(金额_In, 金额), 结算方式 = Nvl(结算方式_In, 结算方式),
        结算号码 = Nvl(结算号码_In, 结算号码), 缴款单位 = Nvl(缴款单位_In, 缴款单位), 单位开户行 = Nvl(单位开户行_In, 单位开户行), 单位帐号 = Nvl(单位帐号_In, 单位帐号),
        摘要 = Nvl(摘要_In, 摘要), 卡类别id = Nvl(卡类别id_In, 卡类别id), 结算卡序号 = Nvl(结算卡序号_In, 结算卡序号), 卡号 = Nvl(卡号_In, 卡号),
        交易流水号 = Nvl(交易流水号_In, 交易流水号), 交易说明 = Nvl(交易说明_In, 交易说明), 预交电子票据 = Nvl(n_预交电子票据, 预交电子票据)
    Where ID = Id_In;
    --自定义过程调用
    Zl_Custom_Balance_Update(Id_In);
    If 操作状态_In = 3 Then
      Return;
    End If;
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

  --相关汇总表处理：人员缴款余额(现收)
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
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_Insert_s;
/
Create Or Replace Procedure Zl_病人预交记录_余额退款_s
(
  Id_In           病人预交记录.Id%Type,
  单据号_In       病人预交记录.No%Type,
  票据号_In       票据使用明细.号码%Type,
  病人id_In       病人预交记录.病人id%Type,
  主页id_In       病人预交记录.主页id%Type,
  姓名_In         病人预交记录.姓名%Type,
  性别_In         病人预交记录.性别%Type,
  年龄_In         病人预交记录.年龄%Type,
  门诊号_In       病人预交记录.门诊号%Type,
  住院号_In       病人预交记录.住院号%Type,
  付款方式名称_In 病人预交记录.付款方式名称%Type,
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
  关联交易id_In   病人预交记录.关联交易id%Type := Null,
  交易流水号_In   病人预交记录.交易流水号%Type := Null,
  交易说明_In     病人预交记录.交易说明%Type := Null,
  合作单位_In     病人预交记录.合作单位%Type := Null,
  收款时间_In     病人预交记录.收款时间%Type := Null,
  结算信息_In     Varchar2 := Null,
  仅更新数据_In   Number := 0,
  操作状态_In     Number := 0,
  结帐id_In       病人预交记录.结帐id%Type := Null,
  结算序号_In     病人预交记录.结算序号%Type := Null,
  预交电子票据_In 病人预交记录.预交电子票据%Type := Null,
  险类_In         保险结算记录.险类%Type := Null
) As
  ----------------------------------------------
  --余额退款操作
  --结算信息_In:原预交ID|金额||....
  --仅更新数据_IN:0-表示需要插入预交记录及更新病人余额;1-表示只更新结算信息中的消费数据
  --操作状态_IN:0-表示完成结算;1-表示未完成结算; (操作状态_IN=1时,生成的预交记录的校对标志为1)
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  n_打印id       票据打印内容.Id%Type;
  d_收款时间     Date;
  n_返回值       病人余额.预交余额%Type;
  n_组id         财务缴款分组.Id%Type;
  n_结帐id       病人预交记录.结帐id%Type;
  n_结算序号     病人预交记录.结算序号%Type;
  n_Count        Number(18);
  n_预交余额     病人余额.预交余额%Type;
  n_预交电子票据 Number(2);
  n_险类         保险结算记录.险类%Type;
Begin
  n_预交电子票据 := 预交电子票据_In;
  If n_预交电子票据 Is Null Then
    n_险类 := 险类_In;
    If 险类_In Is Null Then
      Select Nvl(Max(险类), 0) Into n_险类 From 保险结算记录 Where 记录id = Id_In And 性质 = 3;
    End If;
    n_预交电子票据 := Zl_Fun_Isstarteinvoice(2, n_险类, 预交类别_In);
  End If;

  n_组id := Zl_Get组id(操作员姓名_In);
  If 仅更新数据_In = 0 Then
    d_收款时间 := 收款时间_In;
    If d_收款时间 Is Null Then
      Select Sysdate Into d_收款时间 From Dual;
    End If;
    n_结算序号 := 结算序号_In;
    n_结帐id   := 结帐id_In;
    If Nvl(n_结帐id, 0) = 0 Then
      Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
    End If;
    If Nvl(n_结算序号, 0) = 0 Then
      n_结算序号 := -1 * n_结帐id;
    End If;
    --为了并发，先锁定病人余额(金额_In为负数)
    Update 病人余额
    Set 预交余额 = Nvl(预交余额, 0) + 金额_In
    Where 病人id = 病人id_In And 类型 = 预交类别_In And 性质 = 1
    Returning 预交余额 Into n_预交余额;
    If Sql%RowCount = 0 Then
      Insert Into 病人余额
        (病人id, 性质, 类型, 预交余额, 费用余额)
      Values
        (病人id_In, 1, Nvl(预交类别_In, 0), 金额_In, 0);
      n_预交余额 := 金额_In;
    End If;
  
    Insert Into 病人预交记录
      (ID, NO, 实际票号, 记录性质, 记录状态, 病人id, 主页id, 姓名, 性别, 年龄, 门诊号, 住院号, 付款方式名称, 科室id, 金额, 结算方式, 结算号码, 收款时间, 缴款单位, 单位开户行,
       单位帐号, 操作员编号, 操作员姓名, 摘要, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结帐id, 结算序号, 冲预交, 结算性质, 校对标志, 关联交易id,
       交易时间, 交易人员, 附加标志, 预交电子票据)
    Values
      (Id_In, 单据号_In, Decode(操作状态_In, 0, 票据号_In, Null), 1, 0, 病人id_In, Decode(主页id_In, 0, Null, 主页id_In), 姓名_In, 性别_In,
       年龄_In, Decode(门诊号_In, 0, Null, 门诊号_In), Decode(住院号_In, 0, Null, 住院号_In), 付款方式名称_In,
       Decode(科室id_In, 0, Null, 科室id_In), 金额_In, 结算方式_In, 结算号码_In, d_收款时间, 缴款单位_In, 单位开户行_In, 单位帐号_In, 操作员编号_In,
       操作员姓名_In, 摘要_In, n_组id, 预交类别_In, 卡类别id_In, 结算卡序号_In, 卡号_In, 交易流水号_In, 交易说明_In, 合作单位_In, n_结帐id, n_结算序号, Null,
       Null, 操作状态_In, Decode(Nvl(关联交易id_In, 0), 0, Id_In, 关联交易id_In), 收款时间_In, 操作员姓名_In, 1, n_预交电子票据);
  
    --更新预交单据余额
    Insert Into 预交单据余额 (预交id, 病人id, 预交类别, 预交余额) Values (Id_In, 病人id_In, 预交类别_In, 金额_In);
  End If;

  If 仅更新数据_In = 1 Then
    Select Max(结帐id), Max(收款时间), Max(1) Into n_结帐id, d_收款时间, n_Count From 病人预交记录 Where ID = Id_In;
    If n_Count = 0 Then
      v_Err_Msg := '未找到退款记录，请检查！';
      Raise Err_Item;
    End If;
  End If;

  If 结算信息_In Is Not Null Then
    Zl_病人预交记录_Relevance(病人id_In, Id_In, 结算信息_In, n_结帐id, 操作员编号_In, 操作员姓名_In, 收款时间_In, 操作状态_In, n_组id);
  End If;

  If 操作状态_In = 1 Then
    Return;
  End If;

  --更新记录状态1
  Update 病人预交记录
  Set 记录状态 = 1, 校对标志 = 0, 实际票号 = 票据号_In
  Where NO = 单据号_In And 记录性质 = 1 And 记录状态 = 0
  Returning 结帐id Into n_结帐id;
  If Sql%NotFound Then
    v_Err_Msg := '未找到指定的单据(' || 单据号_In || ',可能因为并发原因被他人退款，请检查！';
    Raise Err_Item;
  End If;
  Update 病人预交记录 Set 校对标志 = 0 Where 结帐id = n_结帐id And Nvl(校对标志, 0) <> 0;

  --处理票据
  If 票据号_In Is Not Null Then
    Select 票据打印内容_Id.Nextval Into n_打印id From Dual;
  
    --发出票据
    Insert Into 票据打印内容 (ID, 数据性质, NO) Values (n_打印id, 2, 单据号_In);
  
    Insert Into 票据使用明细
      (ID, 票种, 号码, 性质, 原因, 领用id, 打印id, 使用时间, 使用人, 票据金额)
    Values
      (票据使用明细_Id.Nextval, 2, 票据号_In, 1, 1, 领用id_In, n_打印id, d_收款时间, 操作员姓名_In, 金额_In);
    --状态改动
    Update 票据领用记录
    Set 当前号码 = 票据号_In, 剩余数量 = Decode(Sign(剩余数量 - 1), -1, 0, 剩余数量 - 1), 使用时间 = Sysdate
    Where ID = Nvl(领用id_In, 0);
  
    Update 病人预交记录 Set 实际票号 = 票据号_In Where 病人id = 病人id_In And 记录性质 = 11 And NO = 单据号_In;
  
  End If;

  --人员缴款余额(现收)
  Update 人员缴款余额
  Set 余额 = Nvl(余额, 0) + 金额_In
  Where 性质 = 1 And 收款员 = 操作员姓名_In And 结算方式 = 结算方式_In
  Returning 余额 Into n_返回值;

  If Sql%RowCount = 0 Then
    Insert Into 人员缴款余额 (收款员, 结算方式, 性质, 余额) Values (操作员姓名_In, 结算方式_In, 1, 金额_In);
    n_返回值 := 金额_In;
  End If;
  If Nvl(n_返回值, 0) = 0 Then
    Delete From 人员缴款余额 Where 收款员 = 操作员姓名_In And 性质 = 1 And 结算方式 = 结算方式_In And Nvl(余额, 0) = 0;
  End If;

  If Nvl(n_预交余额, 0) = 0 Then
    Delete From 病人余额
    Where 病人id = 病人id_In And 类型 = 预交类别_In And Nvl(预交余额, 0) = 0 And Nvl(费用余额, 0) = 0 And 性质 = 1;
  End If;

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
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人预交记录_余额退款_s;
/