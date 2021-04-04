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
--125489:焦博,2018-05-28,增加参数预交公共模块参数预交款分站点显示
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1103, 0, 0, 0, 0, 0, 0, 25, '预交款分站点显示', Null, '0',
         '在预交款管理中控制预交款是否分站点来显示,如果分站点显示,则将只能查询和处理本站点缴款的预交款（余额退款除外）,否则允许查询和操作其他站点的预交款。', '1-分站点显示,0-不分站点显示', Null,
         '适用于总分院形式，严格区分预交分站点显示预交款业务', Null
  From Dual;




-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--126591:刘兴洪,2018-06-01,增加误差费的处理
--126587:刘兴洪,2018-06-01,无结帐数据及消费卡及冲预交的问题
Create Or Replace Procedure Zl_Third_Settlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --------------------------------------------------------------------------------------------------
  --功能:三方接口支付
  --入参:Xml_In:
  --<IN>
  --        <BRID>病人ID</BRID>         //病人ID
  --        <XM>姓名</XM>               //姓名
  --        <SFZH>身份证号</SFZH>       //身份证号
  --        <ZYID>主页ID</ZYID>         //主页ID
  --        <JSLX>2</JSLX>         //结算类型,1-门诊,2-住院，默认为 2
  --        <JE></JE>         //本次结算总金额
  --        <NO></NO>         //结帐的费用单据号(门诊记帐单),目前仅结算类型=1时候使用
  --        <JZKNO></JZKNO>   //结帐的就诊卡单据号,目前仅结算类型=1时候使用
  --        <JZSJ></JZSJ>     //结帐时间
  --       <JSLIST>
  --         <JS>
  --              <JSKLB>支付卡类别</JSKLB >
  --              <JSKH>支付卡号</ JSKH >
  --              <JSFS>支付方式</JSFS> //支付方式:现金;支票,如果是三方卡,可以传空
  --              <JSJE>结算金额</JSJE> //结算金额(正金额：个人补款，负金额：医院退款)<SFCYJ>为1时为冲预交金额
  --              <JYLSH>交易流水号</JYLSH>
  --              <ZY>摘要</ZY>
  --              <SFCYJ>是否冲预交</SFCYJ>  //是否冲预交，0-结算，1-冲预交.允冲预交时,只填JSJE节点
  --              <SFXFK>是否消费卡</SFXFK>  //(1-是消费卡),消费卡时,传入结算卡类别,结算卡号,结算金额等接点
  --              <EXPENDLIST>  //扩展交易信息
  --                  <EXPEND>
  --                        <JYMC>交易名称</JYMC> //交易名称   退款时,传入冲预交的流水号
  --                        <JYLR>交易内容</JYLR> //交易内容   退款时,传入冲预交的金额
  --                  </EXPEND>
  --              </EXPENDLIST>
  --         </JS>
  --       </JSLIST >
  --</IN>

  --出参:Xml_Out
  --  <OUT>
  --    <CZSJ>操作时间</CZSJ>          //HIS的登记时间
  --    ――如无下列错误结点则说明正确执行
  --    <ERROR>
  --      <MSG>错误信息</MSG>
  --    </ERROR>
  --  </OUT>
  --------------------------------------------------------------------------------------------------
  n_主页id     病案主页.主页id%Type;
  n_病人id     病案主页.病人id%Type;
  v_姓名       病人信息.姓名%Type;
  v_身份证号   病人信息.身份证号%Type;
  n_结帐总额   病人预交记录.冲预交%Type;
  n_待结帐金额 病人预交记录.冲预交%Type;
  n_结算类型   Number(3);
  v_操作员编码 病人结帐记录.操作员编号%Type;
  v_操作员姓名 病人结帐记录.操作员姓名%Type;
  n_结帐id     病人结帐记录.Id%Type;
  n_冲预交金额 病人预交记录.冲预交%Type;
  d_结帐时间   Date;
  d_开始日期   Date;
  d_结束日期   Date;
  d_最小日期   Date;
  d_最大日期   Date;

  n_结算卡序号   消费卡类别目录.编号%Type;
  n_时间类型     Number(3);
  v_No           病人结帐记录.No%Type;
  n_卡类别id     医疗卡类别.Id%Type;
  v_结算方式     病人预交记录.结算方式%Type;
  v_Temp         Varchar2(500);
  v_Ids          Varchar2(20000);
  x_Templet      Xmltype; --模板XML
  v_Err_Msg      Varchar2(200);
  v_单据号       Varchar2(20000);
  v_就诊卡单据号 Varchar2(20000);
  Err_Item    Exception;
  Err_Special Exception;

  v_卡类别     三方交易记录.类别%Type;
  v_消费卡结算 Varchar2(20000);
  n_Number     Number(2);
  n_费用id     门诊费用记录.Id%Type;
  n_记录性质   门诊费用记录.记录性质%Type;
  v_费用no     门诊费用记录.No%Type;
  n_序号       门诊费用记录.序号%Type;
  n_记录状态   门诊费用记录.记录状态%Type;
  n_执行状态   门诊费用记录.执行状态%Type;
  n_未结金额   门诊费用记录.实收金额%Type;
  n_结帐金额   门诊费用记录.实收金额%Type;
  n_误差费     门诊费用记录.实收金额%Type;

  Type t_费用结算明细 Is Ref Cursor;
  c_费用结算明细 t_费用结算明细;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZYID'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/JSLX')),
         To_Number(Extractvalue(Value(A), 'IN/NO')), To_Number(Extractvalue(Value(A), 'IN/JZKNO')),
         To_Date(Extractvalue(Value(A), 'IN/JZSJ'), 'yyyy-mm-dd hh24:mi:ss'), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM')
  Into n_主页id, n_病人id, n_结帐总额, n_结算类型, v_单据号, v_就诊卡单据号, d_结帐时间, v_身份证号, v_姓名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  n_结算类型 := Nvl(n_结算类型, 2);
  If n_结算类型 = 1 And Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
  --0.相关检查
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '不能有效识别病人身份,不允许结算!';
    Raise Err_Item;
  End If;

  --人员id,人员编号,人员姓名
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := '系统不能认别有效的操作员,不允许结算!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员编码 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_操作员姓名 := v_Temp;
  v_Err_Msg    := Null;

  For c_交易记录 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If Not (c_交易记录.结算卡类别 Is Null Or Nvl(c_交易记录.是否消费卡, '0') = '1' Or Nvl(c_交易记录.是否冲预交, 0) = 1) Then
    
      Select Decode(Translate(Nvl(c_交易记录.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
      Into n_Number
      From Dual;
    
      If Nvl(n_Number, 0) = 1 Then
        Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_交易记录.结算卡类别);
      Else
        Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_交易记录.结算卡类别;
      End If;
      If v_卡类别 Is Null Then
        v_Err_Msg := '不支持的结算方式,请检查！';
        Raise Err_Item;
      End If;
    
      If Zl_Fun_三方交易记录_Locked(v_卡类别, c_交易记录.交易流水号, c_交易记录.结算卡号, c_交易记录.摘要, 2) = 0 Then
        v_Err_Msg := '交易流水号为:' || c_交易记录.交易流水号 || '的交易正在进行中，不允许再次提交此交易!';
        Raise Err_Special;
      End If;
    
    End If;
  End Loop;
  n_时间类型 := Zl_Getsysparameter('结帐费用时间', 1137);

  If n_结算类型 = 2 Then
    Open c_费用结算明细 For
      Select Max(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
             Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
             Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
      From 住院费用记录
      Where 病人id = n_病人id And 记录状态 <> 0 And 主页id = n_主页id And 记帐费用 = 1
      Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
      Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
      Order By NO, 序号;
  Else
  
    If v_单据号 Is Null And v_就诊卡单据号 Is Null Then
      Open c_费用结算明细 For
        Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
               Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
               Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
        From 门诊费用记录
        Where 病人id = n_病人id And 记录状态 <> 0 And 记帐费用 = 1
        Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
        Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
        
        Union All
        Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
               Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
               Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
        From 住院费用记录
        Where 病人id = n_病人id And 记录状态 <> 0 And 记帐费用 = 1 And Mod(记录性质, 10) = 5
        Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
        Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
        Order By NO, 序号;
    
    Elsif v_单据号 Is Not Null And v_就诊卡单据号 Is Not Null Then
      Open c_费用结算明细 For
        Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
               Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
               Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
        From 门诊费用记录
        Where 病人id + 0 = n_病人id And 记录状态 <> 0 And Mod(记录性质, 10) = 2 And 记帐费用 = 1 And
              NO In (Select /*+cardinality(b,10) */
                      Column_Value
                     From Table(f_Str2list(v_单据号)) B)
        Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
        Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
        Union All
        Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
               Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
               Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
        From 住院费用记录
        Where 病人id + 0 = n_病人id And 记录状态 <> 0 And 记帐费用 = 1 And Mod(记录性质, 10) = 5 And
              NO In (Select /*+cardinality(b,10) */
                      Column_Value
                     From Table(f_Str2list(v_就诊卡单据号)) B)
        Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
        Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
        Order By 记录性质, NO, 序号;
    Elsif v_单据号 Is Not Null Then
      Open c_费用结算明细 For
        Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
               Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
               Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
        From 门诊费用记录
        Where 病人id + 0 = n_病人id And 记录状态 <> 0 And Mod(记录性质, 10) = 2 And 记帐费用 = 1 And
              NO In (Select /*+cardinality(b,10) */
                      Column_Value
                     From Table(f_Str2list(v_单据号)) B)
        Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
        Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
        Order By NO, 序号;
    Else
      Open c_费用结算明细 For
        Select Min(Decode(结帐id, Null, ID, Null)) As ID, Mod(记录性质, 10) As 记录性质, NO, 序号, 记录状态, 执行状态,
               Trunc(Min(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最小时间, Trunc(Max(Decode(n_时间类型, 0, 登记时间, 发生时间))) As 最大时间,
               Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) As 金额, Sum(Nvl(结帐金额, 0)) As 结帐金额
        From 住院费用记录
        Where 病人id + 0 = n_病人id And 记录状态 <> 0 And 记帐费用 = 1 And Mod(记录性质, 10) = 5 And
              NO In (Select /*+cardinality(b,10) */
                      Column_Value
                     From Table(f_Str2list(v_就诊卡单据号)) B)
        Group By Mod(记录性质, 10), NO, 序号, 记录状态, 执行状态
        Having(Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0) Or (Sum(Nvl(实收金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Sum(Nvl(结帐金额, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(结帐金额, 0)) = 0 And Sum(Nvl(应收金额, 0)) <> 0 And Mod(Count(*), 2) = 0
        Order By NO, 序号;
    End If;
  End If;

  Select 病人结帐记录_Id.Nextval, Sysdate, Nextno(15) Into n_结帐id, d_结帐时间, v_No From Dual;
  n_待结帐金额 := 0;
  Loop
    Fetch c_费用结算明细
      Into n_费用id, n_记录性质, v_费用no, n_序号, n_记录状态, n_执行状态, d_最小日期, d_最大日期, n_未结金额, n_结帐金额;
    Exit When c_费用结算明细%NotFound;
  
    n_待结帐金额 := n_待结帐金额 + Nvl(n_未结金额, 0);
    If d_开始日期 Is Null Then
      d_开始日期 := d_最小日期;
    Elsif d_开始日期 > d_最小日期 Then
      d_开始日期 := d_最小日期;
    End If;
    If d_结束日期 Is Null Then
      d_结束日期 := d_最大日期;
    Elsif d_结束日期 < d_最大日期 Then
      d_结束日期 := d_最大日期;
    End If;
  
    If Nvl(n_结帐金额, 0) = 0 Then
      If n_费用id Is Not Null Then
        If Length(v_Ids || ',' || n_费用id) > 4000 Then
          v_Ids := Substr(v_Ids, 2);
          Zl_结帐费用记录_Batch(v_Ids, n_病人id, n_结帐id);
          v_Ids := '';
        End If;
        v_Ids := v_Ids || ',' || n_费用id;
      Else
        Zl_结帐费用记录_Insert(0, v_费用no, n_记录性质, n_记录状态, n_执行状态, n_序号, n_未结金额, n_结帐id);
      End If;
    Else
      Zl_结帐费用记录_Insert(0, v_费用no, n_记录性质, n_记录状态, n_执行状态, n_序号, n_未结金额, n_结帐id);
    End If;
  End Loop;

  If v_Ids Is Not Null Then
    v_Ids := Substr(v_Ids, 2);
    Zl_结帐费用记录_Batch(v_Ids, n_病人id, n_结帐id);
  End If;

  n_待结帐金额 := Round(n_待结帐金额, 6);

  If n_待结帐金额 <> Nvl(n_结帐总额, 0) Then
    v_Err_Msg := '传入的结帐金额与实际结帐金额不符,不允许结算!';
    Raise Err_Item;
  End If;

  Zl_病人结帐记录_Insert(n_结帐id, v_No, n_病人id, d_结帐时间, d_开始日期, d_结束日期, 0, 0, n_主页id, Null, n_结算类型, Null, n_结算类型, 0, n_主页id,
                   n_结帐总额);

  For r_结算方式 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As 是否消费卡,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    v_卡类别   := r_结算方式.结算方式;
    n_结帐金额 := n_结帐金额 + Nvl(r_结算方式.结算金额, 0);
  
    If Nvl(r_结算方式.是否冲预交, 0) = 0 Then
      --付款
      n_卡类别id := Null;
      If r_结算方式.结算卡类别 Is Not Null Then
        Select Decode(Translate(Nvl(r_结算方式.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Number
        From Dual;
        If Nvl(r_结算方式.是否消费卡, 0) = 1 Then
          If Nvl(n_Number, 0) = 1 Then
            Select Max(编号), Max(结算方式), Max(名称)
            Into n_结算卡序号, v_结算方式, v_卡类别
            From 消费卡类别目录
            Where 编号 = n_卡类别id And Nvl(启用, 0) = 1;
          Else
            Select Max(编号), Max(结算方式), Max(名称)
            Into n_结算卡序号, v_结算方式, v_卡类别
            From 消费卡类别目录
            Where 名称 = r_结算方式.结算卡类别 And Nvl(启用, 0) = 1;
          End If;
          If n_结算卡序号 Is Null Then
            v_Err_Msg := '未找到对应的消费卡信息';
            Raise Err_Item;
          End If;
          n_卡类别id := Null;
        Else
          If Nvl(n_Number, 0) = 1 Then
            Select Max(ID), Max(结算方式), Max(名称)
            Into n_卡类别id, v_结算方式, v_卡类别
            From 医疗卡类别
            Where ID = n_卡类别id And Nvl(是否启用, 0) = 1;
          Else
            Select Max(ID), Max(结算方式), Max(名称)
            Into n_卡类别id, v_结算方式, v_卡类别
            From 医疗卡类别
            Where 名称 = r_结算方式.结算卡类别 And Nvl(是否启用, 0) = 1;
          End If;
        
          If n_卡类别id Is Null Then
            v_Err_Msg := '未找到对应的医疗卡信息!';
            Raise Err_Item;
          End If;
        End If;
      End If;
    
      If n_卡类别id Is Not Null Then
        --三方卡
        v_结算方式 := v_结算方式 || '|' || r_结算方式.结算金额 || '|';
        Zl_病人结帐结算_Modify(1, n_病人id, n_结帐id, v_结算方式, Null, 0, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明, 0, 0, 0,
                         n_结算类型, Null, v_操作员编码, v_操作员姓名, d_结帐时间, Null, 0);
        For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                              Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                       From Table(Xmlsequence(Extract(r_结算方式.Expend, '/EXPENDLIST/EXPEND'))) B) Loop
          Zl_三方结算交易_Insert(n_卡类别id, 0, r_结算方式.结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
        End Loop;
      Else
        If n_结算卡序号 Is Not Null Then
          --消费卡
          v_消费卡结算 := Nvl(v_消费卡结算, '') || '||' || n_结算卡序号 || '|' || r_结算方式.结算卡号 || '|0|' || r_结算方式.结算金额;
        Else
          --其他结算
          v_结算方式 := r_结算方式.结算方式 || '|' || r_结算方式.结算金额 || '||';
          Zl_病人结帐结算_Modify(0, n_病人id, n_结帐id, v_结算方式, Null, 0, n_卡类别id, r_结算方式.结算卡号, r_结算方式.交易流水号, r_结算方式.交易说明, 0, 0, 0,
                           n_结算类型, Null, v_操作员编码, v_操作员姓名, d_结帐时间, Null, 0);
        End If;
      End If;
    Else
      --冲预交,目前默认全冲
      n_冲预交金额 := r_结算方式.结算金额;
      Zl_病人结帐结算_Modify(0, n_病人id, n_结帐id, Null, n_冲预交金额, 0, Null, Null, Null, Null, 0, 0, 0, n_结算类型, Null, v_操作员编码,
                       v_操作员姓名, d_结帐时间, Null, 0);
    End If;
  
    Update 三方交易记录
    Set 业务结算id = n_结帐id
    Where 流水号 = Nvl(r_结算方式.交易流水号, '-') And 类别 = v_卡类别 And 业务类型 = 2;
  
  End Loop;

  --消费卡处理
  If v_消费卡结算 Is Not Null Then
    v_消费卡结算 := Substr(v_消费卡结算, 3);
    Zl_病人结帐结算_Modify(3, n_病人id, n_结帐id, v_消费卡结算, Null, 0, Null, Null, Null, Null, 0, 0, 0, n_结算类型, Null, v_操作员编码,
                     v_操作员姓名, d_结帐时间, Null, 0);
  End If;

  n_误差费 := Round(Nvl(n_结帐总额, 0) - Nvl(n_结帐金额, 0), 6);

  If Abs(Nvl(n_误差费, 0)) > 1 Then
    v_Err_Msg := '计算的误差金额大于了1.00或小于-1.00元,不允许结帐操作,请检查!';
    Raise Err_Item;
  End If;
   
  Zl_病人结帐结算_Modify(0, n_病人id, n_结帐id, '', Null, 0, Null, Null, Null, Null, 0, 0, n_误差费, n_结算类型, Null, v_操作员编码, v_操作员姓名,
                   d_结帐时间, Null, 1);

  Update 病人预交记录 Set 校对标志 = 0 Where 结帐id = n_结帐id And Nvl(校对标志, 0) <> 0;
  v_Temp := '<CZSJ>' || To_Char(d_结帐时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Third_Settlement;
/

--125779:殷瑞,2018-05-28,退药按药品id排序处理
Create Or Replace Procedure Zl_输液配药记录_销帐审核
(
  配药id_In   In Varchar2, --ID串:ID1,审核标志1,ID2,审核标志2....
  操作人员_In In 输液配药记录.操作人员%Type,
  操作时间_In In 输液配药记录.操作时间%Type
) Is
  v_Tansid     Varchar2(20);
  v_Tansids    Varchar2(4000);
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
    Select /*+ rule*/
    Distinct a.Id As 退药id, c.收发id, c.数量, a.药品id, a.批次,c.记录id as 配药id
    From 药品收发记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where c.记录id = d.Column_Value And b.Id = c.收发id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And
          a.药品id + 0 = b.药品id And a.序号 = b.序号 And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
    Order By a.药品id, a.批次;

  v_退药记录 c_退药记录%RowType;

  Cursor c_费用销帐 Is
    Select /*+ rule*/
     a.No, a.序号 || ':' || c.数量 || ':' || c.记录id As 费用序号
    From 住院费用记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.Id = b.费用id And b.Id = c.收发id And Mod(b.记录状态, 3) = 1 And c.记录id = d.Column_Value;

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
  
    --统计审核确认的输液单(n_审核标志 = 1)
    If n_审核标志 = 1 Then
      If v_Tansids Is Null Then
        v_Tansids := v_Tansid;
      Else
        v_Tansids := v_Tansids || ',' || v_Tansid;
      End If;
    End If;
  
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
  End Loop;

  --先退药
  For v_退药记录 In c_退药记录 Loop
    Zl_药品收发记录_部门退药(v_退药记录.退药id, 操作人员_In, 操作时间_In, Null, Null, Null, v_退药记录.数量, Null, 操作人员_In);
  
    --取退药待发id
    Select a.Id
    Into v_发药id
    From 药品收发记录 A, 药品收发记录 B
    Where b.Id = v_退药记录.退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
          a.序号 = b.序号 And Mod(a.记录状态, 3) = 1 And a.审核日期 Is Null;
  
    --输液配药内容中的收发ID更新为退药待发的收发ID
    Update 输液配药内容 Set 收发id = v_发药id Where 记录id = v_退药记录.配药id And 收发id = v_退药记录.收发id;
  
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
      Select 记录id, v_原始id, 数量 From 输液配药内容 Where 记录id = v_退药记录.配药id And 收发id = v_发药id;
  
    v_收发ids := v_收发ids || ',' || v_原始id;
  End Loop;

  --费用销帐
  For v_费用销帐 In c_费用销帐 Loop
    Zl_住院记帐记录_Delete(v_费用销帐.No, v_费用销帐.费用序号, v_Usercode, Zl_Username, 2, 1, 1, d_审核时间);
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_销帐审核;
/

--126046:殷瑞,2018-05-28,退药按药品id排序处理
Create Or Replace Procedure Zl_输液配药记录_摆药
(
  部门id_In   In 输液配药记录.部门id%Type,
  配药id_In   In Varchar2, --ID串:ID1,ID2....
  摆药单号_In In 输液配药记录.摆药单号%Type,
  操作人员_In In 输液配药状态.操作人员%Type := Null,
  操作时间_In In 输液配药状态.操作时间%Type := Null,
  移动操作_In In Number := 0
) Is
  v_Tansid Varchar2(20);
  v_Tmp    Varchar2(4000);

  v_收发ids  Varchar2(4000);
  v_Error    Varchar2(255);
  n_是否打包 输液配药记录.是否打包%Type;
  n_操作状态 输液配药记录.操作状态%Type;
  v_摆药人   Varchar2(20);
  v_配药台   Varchar2(20);
  n_配药台id Number(4);
  n_部门id   Number(18);
  n_批次     Number(2);
  d_日期     Date;
  Err_Custom Exception;
  Cursor c_收发记录 Is
    Select /*+ rule*/
     a.Id, Nvl(a.批次, 0) As 批次
    From 药品收发记录 A,
         (Select Distinct 收发id
           From 输液配药内容 A, Table(Cast(f_Num2list(配药id_In) As Zltools.t_Numlist)) B
           Where a.记录id = b.Column_Value) B
    Where a.Id = b.收发id And a.审核人 Is Null
    Order By a.药品id, a.批次;

  v_收发记录 c_收发记录%RowType;
Begin
  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;

  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --检查当前输液单的状态是否为待摆药状态
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Tansid;
    
      If n_操作状态 > 1 Then
        v_Error := '该数据已被操作，不能进行发药！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    Begin
      Select 是否打包 Into n_是否打包 From 输液配药记录 Where ID = v_Tansid For Update Nowait;
    Exception
      When Others Then
        v_Error := '已有其他用户在执行发药，不能重复操作！';
        Raise Err_Custom;
    End;
  
    v_配药台   := '';
    n_配药台id := 0;
    n_部门id   := 0;
    v_摆药人   := '';
    Begin
      Select 名称, ID, 部门id, 配药批次, 执行时间
      Into v_配药台, n_配药台id, n_部门id, n_批次, d_日期
      From (Select f.名称, f.Id, a.部门id, a.配药批次, a.执行时间
             From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C, 配液台药品对照 D, 配液台 F
             Where a.Id = b.记录id And b.收发id = c.Id And c.药品id = d.药品id And d.配药台id = f.Id And c.库房id = d.部门id And
                   a.Id = v_Tansid
             Order By d.配药台id)
      Where Rownum = 1;
    
      Select 摆药人
      Into v_摆药人
      From 配液工作安排
      Where 部门id = n_部门id And 配药台id = n_配药台id And 批次 = n_批次 And
            日期 = To_Date(To_Char(Sysdate, 'yyyy-mm-dd'), 'yyyy-mm-dd');
    Exception
      When Others Then
        Null;
    End;
  
    Update 输液配药记录
    Set 操作状态 = 2, 操作人员 = 操作人员_In, 操作时间 = 操作时间_In, 摆药单号 = 摆药单号_In, 配药台 = v_配药台
    Where ID = v_Tansid;
  
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 实际工作人员)
    Values
      (v_Tansid, 2, 操作人员_In, 操作时间_In, v_摆药人);
    If n_是否打包 <> 0 And 移动操作_In = 0 Then
      Update 输液配药记录 Set 操作状态 = 4, 操作人员 = 操作人员_In, 操作时间 = 操作时间_In Where ID = v_Tansid;
      Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (v_Tansid, 4, 操作人员_In, 操作时间_In);
    End If;
  End Loop;

  For v_收发记录 In c_收发记录 Loop
    If v_收发ids Is Null Then
      v_收发ids := v_收发记录.Id || ',' || v_收发记录.批次;
    Else
      If Length(v_收发ids || '|' || v_收发记录.Id || ',' || v_收发记录.批次) > 3950 Then
        Zl_药品收发记录_批量发药(v_收发ids, 部门id_In, 操作人员_In, 操作时间_In, 4, 操作人员_In, 摆药单号_In);
        v_收发ids := v_收发记录.Id || ',' || v_收发记录.批次;
      Else
        v_收发ids := v_收发ids || '|' || v_收发记录.Id || ',' || v_收发记录.批次;
      End If;
    End If;
  End Loop;

  If Not v_收发ids Is Null Then
    Zl_药品收发记录_批量发药(v_收发ids, 部门id_In, 操作人员_In, 操作时间_In, 4, 操作人员_In, 摆药单号_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_摆药;
/

--126046:殷瑞,2018-05-28,退药按药品id排序处理
Create Or Replace Procedure Zl_输液配药记录_取消摆药(配药id_In In Varchar2 --ID串:配药ID1,配药ID2....
                                           ) Is
  v_Id       Varchar2(20);
  v_发药id   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_Date     Date;
  v_操作人员 输液配药记录.操作人员%Type;
  d_操作时间 输液配药记录.操作时间%Type;
  n_操作状态 输液配药记录.操作状态%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;

  Cursor c_配药内容 Is
    Select /*+ rule*/
    Distinct c.记录id, a.Id As 退药id, c.收发id, a.批号, a.效期, a.产地, c.数量 As 退药数, a.药品id, a.批次
    From 药品收发记录 A, 药品收发记录 B, 输液配药内容 C, Table(Cast(f_Num2list(配药id_In) As Zltools.t_Numlist)) D
    Where c.记录id = d.Column_Value And b.Id = c.收发id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And
          a.药品id + 0 = b.药品id And a.序号 = b.序号 And (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)
    Order By a.药品id, a.批次;

  v_配药内容 c_配药内容%RowType;
Begin
  If 配药id_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := 配药id_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --分解单据ID串
    v_Id  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp := Replace(',' || v_Tmp, ',' || v_Id || ',');
  
    --检查当前输液单的状态是否为待摆药状态
    Begin
      Select 操作状态 Into n_操作状态 From 输液配药记录 Where ID = v_Id;
    
      If n_操作状态 != 2 Then
        v_Error := '该数据已被操作，不能进行取消摆药操作！';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '该数据已被操作！';
        Raise Err_Custom;
    End;
  
    Select 操作人员, 操作时间
    Into v_操作人员, d_操作时间
    From 输液配药状态
    Where 配药id = v_Id And 操作类型 = 1 And Rownum = 1;
  
    Update 输液配药记录 Set 操作状态 = 1, 操作人员 = v_操作人员, 操作时间 = d_操作时间 Where ID = v_Id;
  
    --向[输液配药状态]表中记录“取消摆药”的操作
    Insert Into 输液配药状态
      (配药id, 操作类型, 操作人员, 操作时间, 操作说明)
    Values
      (v_Id, 1, v_操作人员, Sysdate, '取消摆药');
  
  End Loop;

  Select Sysdate Into v_Date From Dual;

  For v_配药内容 In c_配药内容 Loop
    --处理退药
    Zl_药品收发记录_部门退药(v_配药内容.退药id, Zl_Username, v_Date, v_配药内容.批号, v_配药内容.效期, v_配药内容.产地, v_配药内容.退药数, Null, Zl_Username);
  
    Select Max(a.Id)
    Into v_发药id
    From 药品收发记录 A, 药品收发记录 B
    Where b.Id = v_配药内容.退药id And a.单据 = b.单据 And a.No = b.No And a.库房id + 0 = b.库房id And a.药品id + 0 = b.药品id And
          a.序号 = b.序号 And Mod(a.记录状态, 3) = 1 And a.审核日期 Is Null;
  
    --替换输液配药内容中的收发ID
    Update 输液配药内容 Set 收发id = v_发药id Where 记录id = v_配药内容.记录id And 收发id = v_配药内容.收发id;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_输液配药记录_取消摆药;
/

------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0014' Where 编号=&n_System;
Commit;
