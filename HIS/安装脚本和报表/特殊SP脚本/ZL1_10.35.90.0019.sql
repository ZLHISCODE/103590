----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--126717:胡俊勇,2018-07-06,微生物PDF报告打印
Alter Table 医嘱报告内容 Add 打印次数 Number(5); 




------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--127912:余伟节,2018-07-03,预约中心入住或入住取消服务地址配置
Insert Into 三方服务配置目录 (系统标识, 服务名称) Values ('预约中心', '入住或入住取消');

--127796,董露露,2018-07-04,病理诊断录入后可以不录入病理号
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval,&n_System,-Null,-Null,-Null,-Null,-Null,A.* From (
  Select 部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0
  Union All Select 0, 0, 307, '不填写病理号','0', '0','控制首页整理时填写病理诊断后是否可以不填写病理号', '0-填写病理诊断后必须填写病理号；1-填写病理诊断后可以不填写病理号', '', '适用于病人首页整理时', Null
  From Dual)A;


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--126717:胡俊勇,2018-07-06,微生物PDF报告打印
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1252,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select 'Zl_医嘱报告内容_Print','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select &n_System,1253,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0 Union All 
    Select 'Zl_医嘱报告内容_Print','EXECUTE' From Dual Union All
Select 对象,权限 From zlProgPrivs Where 1 = 0) A;




-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--128177:冉俊明,2018-07-06,三方平台挂号接口支持使用预交款
Create Or Replace Procedure Zl_Third_Regist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --功能:HIS挂号 
  --入参:Xml_In: 
  --<IN> 
  --   <CZFS>3</CZFS>    //操作方式 
  --   <CZJLID>1</CZJLID>    //出诊记录ID 
  --   <HM>号码</HM>    //号码 
  --   <HX>号序</HX>     //号序 
  --   <JKFS>0</JKFS>  //缴款方式,0-挂号或预约缴款;1-预约不缴款 
  --   <YYSJ>2014-10-21 </YYSJ>    //预约日期 YYYY-MM-DD,分时段非序号控制需要传入时间 
  --   <JE>金额</JE>     //金额 
  --   <JSLIST> 
  --     <JS>            //结算信息，挂号非医保结算目前仅支持一个，结构与收费一致 
  --       <JSKLB>结算卡类别</JSKLB>    //结算卡类别 
  --       <JSKH>支付宝帐号</JSKH>           //结算卡号(支付宝帐号) 
  --       <JYSM>交易说明</JYSM>            //说明，固定传支付宝 
  --       <JYLSH>流水号</JYLSH>           //流水号，传订单号 
  --       <JSFS>结算方式</JSFS>            //结算方式:现金、支票，如果是三方卡,可以传空 
  --       <JSJE>结算金额</JSJE>            //结算金额 
  --       <ZY>摘要</ZY>                  //摘要 
  --       <SFCYJ></SFCYJ>              //是否冲预交，挂号目前不传 
  --       <SFXFK></SFXFK>              //是否消费卡,挂号目前不传 
  --       <EXPENDLIST>                 //扩展信息 
  --         <EXPEND> 
  --           <JYMC>交易名称</JYMC>        //交易名称 
  --           <JYLR>交易内容<JYLR>         //交易内容 
  --         </EXPEND> 
  --         <EXPEND> 
  --           ... 
  --         </EXPEND> 
  --       </EXPENDLIST> 
  --     </JS> 
  --   </JSLIST> 
  --   <HZDW>合作单位</HZDW>        //合作单位名称 
  --   <YYFS>支付宝<YYFS>    //预约方式,如自助机，支付宝 
  --   <BRID>病人ID</BRID>     //病人ID 
  --   <SFZH>身份证号</SFZH>     //身份证号 
  --   <XM>姓名</XM>            //姓名 
  --   <BRLX></BRLX>             //医保病人类型 
  --   <FB>普通</FB>               //病人费别，可以不传 
  --   <JQM>机器名</JQM>            //机器名 
  --</IN> 

  --出参:Xml_Out 
  --<OUTPUT> 
  -- <GHDH>挂号单号</GHDH>          //挂号单号 
  -- <CZSJ>操作时间</CZSJ>          //HIS的登记时间 
  -- <JZID>结帐ID</JZID>          //本次结帐ID 
  -- <ERROR><MSG>错误信息</MSG></ERROR>  //出错时返回 
  --</ OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  v_号码     挂号安排.号码%Type;
  d_发生时间 Date;
  d_原始时间 Date;
  d_登记时间 Date;

  n_应收金额   门诊费用记录.应收金额%Type;
  v_流水号     病人预交记录.交易流水号%Type;
  v_说明       门诊费用记录.摘要%Type;
  n_病人id     病人信息.病人id%Type;
  v_身份证号   病人信息.身份证号%Type;
  v_预约方式   预约方式.名称%Type;
  v_卡类别名称 医疗卡类别.名称%Type;
  v_结算卡号   病人预交记录.卡号%Type;
  n_门诊号     门诊费用记录.标识号%Type;
  v_姓名       门诊费用记录.姓名%Type;
  v_性别       门诊费用记录.性别%Type;
  v_年龄       门诊费用记录.年龄%Type;
  v_付款方式   门诊费用记录.付款方式%Type;
  v_费别       门诊费用记录.费别%Type;
  v_No         病人挂号记录.No%Type;
  v_结算方式   医疗卡类别.结算方式%Type;
  n_收费细目id 门诊费用记录.收费细目id%Type;
  n_病人科室id 门诊费用记录.病人科室id%Type;
  n_开单部门id 门诊费用记录.开单部门id%Type;
  v_操作员编号 门诊费用记录.操作员编号%Type;
  v_操作员姓名 门诊费用记录.操作员姓名%Type;
  v_医生姓名   挂号安排.医生姓名%Type;
  n_医生id     挂号安排.医生id%Type;
  n_结帐id     门诊费用记录.结帐id%Type;
  n_卡类别id   医疗卡类别.Id%Type;
  v_排班       挂号安排.周日%Type;
  n_安排id     挂号安排.Id%Type;
  n_计划id     挂号安排计划.Id%Type;
  n_预交id     病人预交记录.Id%Type;
  n_序号控制   挂号安排.序号控制%Type;
  n_号序       挂号序号状态.序号%Type;
  v_星期       挂号安排限制.限制项目%Type;
  v_病人类型   病人信息.病人类型%Type;
  n_存在       Number(3);
  v_现金       结算方式.名称%Type;
  n_分时段     Number(3);
  v_结算内容   Varchar2(3000);
  v_合作单位   病人挂号记录.合作单位%Type;
  v_机器名     挂号序号状态.机器名%Type;
  n_缴款方式   Number(3);
  n_挂号模式   Number(3);
  n_Exists     Number(3);
  v_保险结算   Varchar2(1000);
  n_记录id     临床出诊记录.Id%Type;
  v_Temp       Varchar2(32767); --临时XML 
  x_Templet    Xmltype; --模板XML 
  v_Err_Msg    Varchar2(200);
  d_启用时间   Date;
  n_Count      Number(3);
  v_卡类别     三方交易记录.类别%Type;
  n_冲预交     病人预交记录.冲预交%Type;
  v_Para       Varchar2(2000);
  Err_Item Exception;
  Err_Special Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/HX'),
         To_Date(Extractvalue(Value(A), 'IN/YYSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/JE'),
         Extractvalue(Value(A), 'IN/YYFS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/BRID'),
         Extractvalue(Value(A), 'IN/BRLX'), Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/JQM'),
         Extractvalue(Value(A), 'IN/JKFS'), Extractvalue(Value(A), 'IN/CZJLID'), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM')
  Into v_号码, n_号序, d_原始时间, n_应收金额, v_预约方式, v_合作单位, n_病人id, v_病人类型, v_费别, v_机器名, n_缴款方式, n_记录id, v_身份证号, v_姓名
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_病人id, 0) = 0 And Not v_身份证号 Is Null And Not v_姓名 Is Null Then
    n_病人id := Zl_Third_Getpatiid(v_身份证号, v_姓名);
  End If;
  If Nvl(n_病人id, 0) = 0 Then
    v_Err_Msg := '无法确定病人信息,请检查!';
    Raise Err_Item;
  End If;

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  If Sysdate - 10 > Nvl(d_启用时间, Sysdate - 30) Then
    If n_挂号模式 = 1 And Nvl(d_原始时间, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_记录id Is Null Then
      v_Err_Msg := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
      Raise Err_Item;
    End If;
  Else
    If n_挂号模式 = 1 And Nvl(d_原始时间, Sysdate) > Nvl(d_启用时间, Sysdate - 30) And n_记录id Is Null Then
      Begin
        Select a.Id
        Into n_记录id
        From 临床出诊记录 A, 临床出诊号源 B
        Where a.号源id = b.Id And b.号码 = v_号码 And Nvl(d_原始时间, Sysdate) Between a.开始时间 And a.终止时间;
      Exception
        When Others Then
          v_Err_Msg := '系统当前处于出诊表排班模式，传入的参数无法确定挂号安排，请重试！';
          Raise Err_Item;
      End;
    End If;
  End If;

  d_登记时间 := Sysdate;
  d_发生时间 := Trunc(d_原始时间);

  For c_交易记录 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号, Extractvalue(b.Column_Value, '/JS/ZY') As 摘要
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    --冲预交不需要三方交易锁 
    If Nvl(c_交易记录.是否冲预交, 0) = 0 Then
      If c_交易记录.结算卡类别 Is Null Then
        v_卡类别 := c_交易记录.结算方式;
      Else
        Select Decode(Translate(Nvl(c_交易记录.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Count
        From Dual;
      
        If Nvl(n_Count, 0) = 1 Then
          Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(c_交易记录.结算卡类别);
        Else
          Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = c_交易记录.结算卡类别;
        End If;
      End If;
    
      If v_卡类别 Is Null Then
        v_Err_Msg := '不支持的结算方式,请检查！';
        Raise Err_Item;
      End If;
    
      If Zl_Fun_三方交易记录_Locked(v_卡类别, c_交易记录.交易流水号, c_交易记录.结算卡号, c_交易记录.摘要, 4) = 0 Then
        v_Err_Msg := '交易流水号为:' || c_交易记录.交易流水号 || '的交易正在进行中，不允许再次提交此交易!';
        Raise Err_Special;
      End If;
    End If;
  End Loop;

  If v_病人类型 Is Not Null Then
    Begin
      Select 1 Into n_存在 From 病人类型 Where 名称 = v_病人类型;
    Exception
      When Others Then
        v_Err_Msg := '没有发现为(' || v_病人类型 || ')的病人类型';
        Raise Err_Item;
    End;
    Update 病人信息 Set 病人类型 = Nvl(病人类型, v_病人类型) Where 病人id = n_病人id;
  End If;

  Select a.门诊号, a.姓名, a.性别, a.年龄, Nvl(b.编码, c.编码)
  Into n_门诊号, v_姓名, v_性别, v_年龄, v_付款方式
  From 病人信息 A, 医疗付款方式 B, (Select 编码 From 医疗付款方式 Where 缺省标志 = '1' And Rownum < 2) C
  Where a.病人id = n_病人id And a.医疗付款方式 = b.名称(+);

  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_操作员编号 From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_操作员姓名 From Dual;
  v_Temp := Zl_Identity(2);
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into n_开单部门id From Dual;

  v_No := Nextno(12);
  Select 病人结帐记录_Id.Nextval Into n_结帐id From Dual;
  Select 病人预交记录_Id.Nextval Into n_预交id From Dual;

  If n_记录id Is Null Then
    For r_结算 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_结算.是否冲预交, 0) = 0 Then
        If r_结算.结算方式 Is Null Then
          Begin
            Select b.结算方式, b.Id
            Into v_结算方式, n_卡类别id
            From 医疗卡类别 B
            Where b.名称 = r_结算.结算卡类别 And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '没有发现该结算卡的相关信息';
              Raise Err_Item;
          End;
        Else
          Select Nvl(Max(1), 0) Into n_Exists From 结算方式 Where 名称 = r_结算.结算方式 And 性质 In (3, 4);
          If n_Exists = 1 Then
            v_保险结算 := v_保险结算 || '||' || r_结算.结算方式 || '|' || r_结算.结算金额;
          Else
            If v_结算方式 Is Null Then
              v_结算方式 := r_结算.结算方式;
            Else
              v_Err_Msg := '目前计划排班挂号不支持非医保外的多种结算方式,请检查!';
              Raise Err_Item;
            End If;
          End If;
        End If;
      
        If r_结算.结算卡类别 Is Not Null Then
          v_卡类别名称 := r_结算.结算卡类别;
          v_结算卡号   := r_结算.结算卡号;
          v_流水号     := r_结算.交易流水号;
          v_说明       := r_结算.交易说明;
        
          If n_卡类别id Is Null Then
            Begin
              Select b.Id Into n_卡类别id From 医疗卡类别 B Where b.名称 = r_结算.结算卡类别 And Rownum < 2;
            Exception
              When Others Then
                v_Err_Msg := '没有发现该结算卡的相关信息';
                Raise Err_Item;
            End;
          End If;
        
          Select Decode(Translate(Nvl(r_结算.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(r_结算.结算卡类别);
          Else
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = r_结算.结算卡类别;
          End If;
        Else
          v_卡类别 := r_结算.结算方式;
        End If;
      
        Update 三方交易记录
        Set 业务结算id = n_结帐id
        Where 流水号 = r_结算.交易流水号 And 类别 = v_卡类别 And 业务类型 = 4;
      Else
        n_冲预交 := r_结算.结算金额;
      End If;
    End Loop;
  
    If v_保险结算 Is Not Null Then
      v_保险结算 := Substr(v_保险结算, 3);
    End If;
  
    Select Decode(To_Char(d_原始时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',
                   Null)
    Into v_星期
    From Dual;
  
    Begin
      Select ID
      Into n_计划id
      From (Select ID
             From 挂号安排计划
             Where 号码 = v_号码 And d_原始时间 Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                   Nvl(失效时间, To_Date('3000-01-01', 'YYYY-MM-DD')) And 审核时间 Is Not Null
             Order By 生效时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Select ID Into n_安排id From 挂号安排 Where 号码 = v_号码;
    End;
  
    If Nvl(n_计划id, 0) <> 0 Then
      --从计划读取信息 
      Select a.项目id, b.科室id, a.医生姓名, a.医生id,
             Decode(To_Char(d_发生时间, 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7', a.周六,
                     Null), Nvl(a.序号控制, 0)
      Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, v_排班, n_序号控制
      From 挂号安排计划 A, 挂号安排 B
      Where a.Id = n_计划id And b.Id = a.安排id;
      Select Count(Rownum) Into n_分时段 From 挂号计划时段 Where 星期 = v_星期 And 计划id = n_计划id And Rownum < 2;
    
      --合作单位检查 
      If v_合作单位 Is Not Null Then
        Begin
          Select 1 Into n_存在 From 合作单位计划控制 Where 计划id = n_计划id And 数量 = 0 And 合作单位 = v_合作单位;
        Exception
          When Others Then
            n_存在 := 0;
        End;
      End If;
    
      If n_存在 = 1 Then
        v_Err_Msg := '传入的合作单位在此号码上被禁用！';
        Raise Err_Item;
      End If;
    
      If n_分时段 = 1 And n_序号控制 = 0 Then
        d_发生时间 := d_原始时间;
        Select 序号
        Into n_号序
        From 挂号计划时段
        Where 计划id = n_计划id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi:ss') = To_Char(d_发生时间, 'hh24:mi:ss');
      Else
        Begin
          Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
          Into d_发生时间
          From 挂号计划时段
          Where 计划id = n_计划id And 星期 = v_星期 And 序号 = Nvl(n_号序, 0);
        Exception
          When Others Then
            If n_分时段 = 1 And n_序号控制 = 1 Then
              Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(Max(结束时间), 'hh24:mi:ss'),
                              'YYYY-MM-DD hh24:mi:ss')
              Into d_发生时间
              From 挂号计划时段
              Where 计划id = n_计划id And 星期 = v_星期;
            Else
              Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
              Into d_发生时间
              From 时间段
              Where 时间段 = v_排班;
            End If;
            If d_发生时间 < d_登记时间 Then
              d_发生时间 := d_登记时间;
            End If;
        End;
      End If;
    Else
      --从安排读取信息 
      Select b.项目id, b.科室id, b.医生姓名, b.医生id,
             Decode(To_Char(d_发生时间, 'D'), '1', b.周日, '2', b.周一, '3', b.周二, '4', b.周三, '5', b.周四, '6', b.周五, '7', b.周六,
                     Null), Nvl(b.序号控制, 0)
      Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, v_排班, n_序号控制
      From 挂号安排 B
      Where b.Id = n_安排id;
      Select Count(Rownum) Into n_分时段 From 挂号安排时段 Where 星期 = v_星期 And 安排id = n_安排id And Rownum < 2;
    
      --合作单位检查 
      If v_合作单位 Is Not Null Then
        Begin
          Select 1 Into n_存在 From 合作单位安排控制 Where 安排id = n_安排id And 数量 = 0 And 合作单位 = v_合作单位;
        Exception
          When Others Then
            n_存在 := 0;
        End;
      End If;
    
      If n_存在 = 1 Then
        v_Err_Msg := '传入的合作单位在此号码上被禁用！';
        Raise Err_Item;
      End If;
    
      If n_分时段 = 1 And n_序号控制 = 0 Then
        d_发生时间 := d_原始时间;
        Select 序号
        Into n_号序
        From 挂号安排时段
        Where 安排id = n_安排id And 星期 = v_星期 And To_Char(开始时间, 'hh24:mi:ss') = To_Char(d_发生时间, 'hh24:mi:ss');
      Else
        Begin
          Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
          Into d_发生时间
          From 挂号安排时段
          Where 安排id = n_安排id And 星期 = v_星期 And 序号 = Nvl(n_号序, 0);
        Exception
          When Others Then
            If n_分时段 = 1 And n_序号控制 = 1 Then
              Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(Max(结束时间), 'hh24:mi:ss'),
                              'YYYY-MM-DD hh24:mi:ss')
              Into d_发生时间
              From 挂号安排时段
              Where 安排id = n_安排id And 星期 = v_星期;
            Else
              Select To_Date(To_Char(d_发生时间, 'yyyy-mm-dd') || '' || To_Char(开始时间, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
              Into d_发生时间
              From 时间段
              Where 时间段 = v_排班;
            End If;
            If d_发生时间 < d_登记时间 Then
              d_发生时间 := d_登记时间;
            End If;
        End;
      End If;
    End If;
  
    If Trunc(d_发生时间) <> Trunc(Sysdate) Then
      If Nvl(n_缴款方式, 0) = 0 Then
        Zl_三方机构挂号_Insert(3, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1);
      Else
        Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1);
      End If;
    Else
      If Nvl(n_缴款方式, 0) = 0 Then
        Zl_三方机构挂号_Insert(1, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1);
      Else
        Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算方式, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, Null, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1);
      End If;
    End If;
  
    For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_结帐id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --出诊表排班模式 
    For r_结算 In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As 结算卡类别,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As 结算卡号,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As 是否冲预交,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As 结算方式,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As 交易流水号,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As 交易说明,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As 结算金额
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_结算.是否冲预交, 0) = 0 Then
        If r_结算.结算方式 Is Null Then
          Begin
            Select b.结算方式, b.Id
            Into v_结算方式, n_卡类别id
            From 医疗卡类别 B
            Where b.名称 = r_结算.结算卡类别 And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '没有发现该结算卡的相关信息';
              Raise Err_Item;
          End;
          v_结算内容 := v_结算内容 || '|' || v_结算方式 || ',' || r_结算.结算金额 || ',,';
        Else
          v_结算内容 := v_结算内容 || '|' || r_结算.结算方式 || ',' || r_结算.结算金额 || ',,';
        End If;
      
        If r_结算.结算卡类别 Is Not Null Then
          v_结算内容   := v_结算内容 || '1';
          v_卡类别名称 := r_结算.结算卡类别;
          v_结算卡号   := r_结算.结算卡号;
          v_流水号     := r_结算.交易流水号;
          v_说明       := r_结算.交易说明;
          If n_卡类别id Is Null Then
            Begin
              Select b.Id Into n_卡类别id From 医疗卡类别 B Where b.名称 = r_结算.结算卡类别 And Rownum < 2;
            Exception
              When Others Then
                v_Err_Msg := '没有发现该结算卡的相关信息';
                Raise Err_Item;
            End;
          End If;
        
          Select Decode(Translate(Nvl(r_结算.结算卡类别, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where ID = To_Number(r_结算.结算卡类别);
          Else
            Select Max(名称) Into v_卡类别 From 医疗卡类别 Where 名称 = r_结算.结算卡类别;
          End If;
        Else
          v_结算内容 := v_结算内容 || '0';
          v_卡类别   := r_结算.结算方式;
        End If;
      
        Update 三方交易记录
        Set 业务结算id = n_结帐id
        Where 流水号 = r_结算.交易流水号 And 类别 = v_卡类别 And 业务类型 = 4;
      Else
        n_冲预交 := r_结算.结算金额;
      End If;
    End Loop;
  
    If v_结算内容 Is Not Null Then
      v_结算内容 := Substr(v_结算内容, 2);
    Else
      Begin
        Select 名称 Into v_现金 From 结算方式 Where 性质 = 1;
      Exception
        When Others Then
          v_现金 := '现金';
      End;
      v_结算内容 := v_现金 || ',' || 0 || ',,0';
    End If;
  
    Select 项目id, 科室id, 医生姓名, 医生id, 是否序号控制, 是否分时段
    Into n_收费细目id, n_病人科室id, v_医生姓名, n_医生id, n_序号控制, n_分时段
    From 临床出诊记录
    Where ID = n_记录id;
  
    Begin
      Select 开始时间 Into d_发生时间 From 临床出诊序号控制 Where 记录id = n_记录id And 序号 = n_号序;
    Exception
      When Others Then
        d_发生时间 := d_原始时间;
    End;
  
    If Trunc(d_发生时间) <> Trunc(Sysdate) Then
      If Nvl(n_缴款方式, 0) = 0 Then
        Zl_三方机构挂号_Insert(3, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, v_保险结算, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1, 0, n_记录id);
      Else
        Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, v_保险结算, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1, 0, n_记录id);
      End If;
    Else
      If Nvl(n_缴款方式, 0) = 0 Then
        Zl_三方机构挂号_Insert(1, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, v_保险结算, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1, 0, n_记录id);
      Else
        Zl_三方机构挂号_Insert(2, n_病人id, v_号码, n_号序, v_No, Null, v_结算内容, Null, d_发生时间, d_登记时间, v_合作单位, n_应收金额, Null, Null,
                         v_流水号, v_说明, v_预约方式, n_预交id, n_卡类别id, Null, 1, n_结帐id, 0, v_保险结算, n_冲预交, v_结算卡号, 1, v_费别, Null,
                         v_机器名, 1, 0, n_记录id);
      End If;
    End If;
  
    For c_扩展信息 In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As 交易名称,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As 交易内容
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_三方结算交易_Insert(n_卡类别id, 0, v_结算卡号, n_结帐id, c_扩展信息.交易名称 || '|' || c_扩展信息.交易内容, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_登记时间, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_结帐id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Regist;
/

--126717:胡俊勇,2018-07-06,微生物PDF报告打印
Create Or Replace Procedure Zl_医嘱报告内容_Print
(
  报告id_In In 医嘱报告内容.Id%Type,
  类型_In   In Number
) Is
  --类型_In:0-表示打印告
Begin
  If 类型_In = 0 Then
    Update 医嘱报告内容 Set 打印次数 = Nvl(打印次数, 0) + 1 Where ID = 报告id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_医嘱报告内容_Print;
/

--128391:陈龙,2018-07-05,增加审核状态为2的情况
CREATE OR REPLACE PROCEDURE Zl_医嘱审核管理_Cancel
(
  医嘱ids_In  VARCHAR2,
  审核对象_In NUMBER := 1, --1=手术医嘱，2=输血医嘱
  执行类别_In NUMBER := 0 --0=老版血库流程；不为0时，则为目标审核状态：1=待审核；7=待签发；4-已签发；3-已拒绝；
) IS
  --取消审核
  CURSOR c_Advice IS
    SELECT * FROM TABLE(CAST(f_Num2list(医嘱ids_In) AS t_Numlist));
  Err_Item EXCEPTION;
  v_Err_Msg  VARCHAR2(200);
  n_Count    NUMBER;
  n_医嘱状态 NUMBER;
  n_审核状态 NUMBER;
  n_操作类型 NUMBER;
BEGIN
  FOR r_Advice IN c_Advice LOOP
    SELECT COUNT(1), MAX(医嘱状态), Nvl(MAX(审核状态), 0)
    INTO n_Count, n_医嘱状态, n_审核状态
    FROM 病人医嘱记录
    WHERE Id = r_Advice.Column_Value;
  
    IF n_Count = 0 THEN
      v_Err_Msg := '有医嘱已经删除,请查证。';
      RAISE Err_Item;
    END IF;
  
    IF n_医嘱状态 <> 1 THEN
      v_Err_Msg := '您选择的医嘱中包含有校对的医嘱，不能取消审核。';
      RAISE Err_Item;
    END IF;
  
    IF n_审核状态 = 1 THEN
      n_操作类型 := 19;
    ELSIF n_审核状态 = 7 THEN
      n_操作类型 := 18;
    ELSIF n_审核状态 = 3 THEN
      n_操作类型 := 12;
    ELSIF n_审核状态 = 4 OR n_审核状态 = 2 THEN
      n_操作类型 := 11;
    END IF;
  
    IF 审核对象_In = 1 OR 执行类别_In = 0 THEN
      UPDATE 病人医嘱记录 SET 审核状态 = 1 WHERE Id = r_Advice.Column_Value OR 相关id = r_Advice.Column_Value;
      DELETE FROM 病人医嘱状态
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = r_Advice.Column_Value OR 相关id = r_Advice.Column_Value) AND
            操作类型 IN (11, 12) AND
            操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = r_Advice.Column_Value AND 操作类型 IN (11, 12));
    ELSIF 审核对象_In = 2 AND 执行类别_In <> 0 THEN
      UPDATE 病人医嘱记录
      SET 审核状态 = 执行类别_In
      WHERE Id = r_Advice.Column_Value OR 相关id = r_Advice.Column_Value;
      DELETE FROM 病人医嘱状态
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = r_Advice.Column_Value OR 相关id = r_Advice.Column_Value) AND
            操作类型 = n_操作类型 AND
            操作时间 =
            (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = r_Advice.Column_Value AND 操作类型 = n_操作类型);
    END IF;
  
  END LOOP;
EXCEPTION
  WHEN Err_Item THEN
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  WHEN OTHERS THEN
    Zl_Errorcenter(SQLCODE, SQLERRM);
END Zl_医嘱审核管理_Cancel;
/

--128391:陈龙,2018-07-05,增加审核状态为2的情况
CREATE OR REPLACE PROCEDURE Zl_医嘱审核管理_Update
(
  医嘱id_In   病人医嘱状态.医嘱id%TYPE,
  操作时间_In 病人医嘱状态.操作时间%TYPE,
  操作说明_In 病人医嘱状态.操作说明%TYPE := NULL,
  审核对象_In NUMBER := 1, --1=手术医嘱，2=输血医嘱
  操作人员_In VARCHAR2 := NULL
) IS
  --修改只适用于审核不通过的医嘱，修改其审核说明
  Err_Item EXCEPTION;
  v_Err_Msg  VARCHAR2(200);
  n_Count    NUMBER;
  n_审核状态 NUMBER;
BEGIN
  SELECT COUNT(1), Nvl(MAX(审核状态), 0) INTO n_Count, n_审核状态 FROM 病人医嘱记录 WHERE Id = 医嘱id_In;
  IF n_Count = 0 THEN
    v_Err_Msg := '有医嘱已经删除,请查证。';
    RAISE Err_Item;
  END IF;

  IF 审核对象_In = 1 THEN
    UPDATE 病人医嘱状态
    SET 操作时间 = 操作时间_In, 操作说明 = 操作说明_In
    WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In) AND 操作类型 = 12 AND
          操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = 医嘱id_In AND 操作类型 = 12);
  ELSE
    IF n_审核状态 = 1 THEN
      UPDATE 病人医嘱状态
      SET 操作时间 = 操作时间_In, 操作说明 = 操作说明_In
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In) AND 操作类型 = 19 AND
            操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = 医嘱id_In AND 操作类型 = 19);
      IF SQL%NOTFOUND THEN
        INSERT INTO 病人医嘱状态
          (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
          SELECT Id, 19, 操作人员_In, 操作时间_In, 操作说明_In
          FROM 病人医嘱记录
          WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In;
      END IF;
    ELSIF n_审核状态 = 7 THEN
      UPDATE 病人医嘱状态
      SET 操作时间 = 操作时间_In, 操作说明 = 操作说明_In
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In) AND 操作类型 = 18 AND
            操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = 医嘱id_In AND 操作类型 = 18);
      IF SQL%NOTFOUND THEN
        INSERT INTO 病人医嘱状态
          (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
          SELECT Id, 18, 操作人员_In, 操作时间_In, 操作说明_In
          FROM 病人医嘱记录
          WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In;
      END IF;
    ELSIF n_审核状态 = 3 THEN
      UPDATE 病人医嘱状态
      SET 操作时间 = 操作时间_In, 操作说明 = 操作说明_In
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In) AND 操作类型 = 12 AND
            操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = 医嘱id_In AND 操作类型 = 12);
      IF SQL%NOTFOUND THEN
        INSERT INTO 病人医嘱状态
          (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
          SELECT Id, 12, 操作人员_In, 操作时间_In, 操作说明_In
          FROM 病人医嘱记录
          WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In;
      END IF;
    ELSIF n_审核状态 = 4 OR n_审核状态 = 2 THEN
      UPDATE 病人医嘱状态
      SET 操作时间 = 操作时间_In, 操作说明 = 操作说明_In
      WHERE 医嘱id IN (SELECT Id FROM 病人医嘱记录 WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In) AND 操作类型 = 11 AND
            操作时间 = (SELECT MAX(操作时间) FROM 病人医嘱状态 WHERE 医嘱id = 医嘱id_In AND 操作类型 = 11);
      IF SQL%NOTFOUND THEN
        INSERT INTO 病人医嘱状态
          (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
          SELECT Id, 11, 操作人员_In, 操作时间_In, 操作说明_In
          FROM 病人医嘱记录
          WHERE Id = 医嘱id_In OR 相关id = 医嘱id_In;
      END IF;
    END IF;
  END IF;
EXCEPTION
  WHEN Err_Item THEN
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  WHEN OTHERS THEN
    Zl_Errorcenter(SQLCODE, SQLERRM);
END Zl_医嘱审核管理_Update;
/

--127912:余伟节,2018-07-03,预约系统调用
Create Or Replace Procedure Zl_Third_Outpatireg
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：用于产生预入院记录/取消预入院    数据写入
  --入参：xml_in
  --<IN>
  -- <TYPE>1</TYPE>   --操作类型：1-产生预入院记录；0-取消预入院
  -- <GHID>1162695</GHID>       --挂号id
  -- <RYKSID>202704</RYKSID>    --入院科室ID
  -- <RYBQID>202704</RYBQID>    --入院病区ID
  -- <CH>5</CH>   --床号
  -- <YZID>3</YZID> --医嘱id
  -- <CZYBH></CZYBH> --操作员编号
  -- <CZYXM></CZYXM> --操作员姓名
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

  n_医嘱id 病人医嘱记录.Id%Type;
  Cursor c_Advice Is
    Select Nvl(a.相关id, a.Id) As 组id, a.相关id, a.序号, a.病人id, a.挂号单, a.婴儿, a.姓名, c.操作类型, a.诊疗类别, a.医嘱状态, a.医嘱内容, a.开嘱医生,
           a.开始执行时间, a.执行时间方案, a.频率次数, a.频率间隔, a.间隔单位, Nvl(a.紧急标志, 0) As 紧急标志, a.诊疗项目id, a.收费细目id
    From 病人医嘱记录 A, 诊疗项目目录 C
    Where a.诊疗项目id = c.Id And a.诊疗类别 = 'Z' And c.操作类型 = '2' And a.Id = n_医嘱id;
  r_Advice c_Advice%RowType;

  Cursor c_Pati(v_病人id 病人信息.病人id%Type) Is
    Select a.病人id, a.住院号, a.姓名, a.性别, a.年龄, a.费别, a.出生日期, a.国籍, a.民族, a.学历, a.婚姻状况, a.职业, a.身份, a.身份证号, a.出生地点, a.家庭地址,
           a.家庭地址邮编, a.家庭电话, a.户口地址, a.户口地址邮编, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, a.工作单位, a.合同单位id, a.单位电话, a.单位邮编,
           a.单位开户行, a.单位帐号, a.担保人, a.担保额, a.担保性质, a.籍贯, a.区域, a.医疗付款方式, a.险类
    From 病人信息 A
    Where a.病人id = v_病人id;
  r_Pati c_Pati%RowType;

  n_Type   Number;
  n_挂号id 病人医嘱记录.Id%Type;
  n_科室id 病人医嘱记录.Id%Type;
  n_病区id 病人医嘱记录.Id%Type;
  v_床号   病案主页.入院病床%Type;

  n_病人id 病案主页.病人id%Type;
  v_No     病人挂号记录.No%Type;
  n_Count  Number;

  v_入院方式 病案主页.入院方式%Type;
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_Temp     Varchar2(4000);
  v_Error    Varchar2(2000);

  Err_Custom Exception;
Begin
  Select Extractvalue(Value(A), 'IN/TYPE'), Extractvalue(Value(A), 'IN/GHID') As 挂号id,
         Extractvalue(Value(A), 'IN/RYKSID') As 入院科室id, Extractvalue(Value(A), 'IN/RYBQID') As 入院病区id,
         Extractvalue(Value(A), 'IN/CH') As 床号, Extractvalue(Value(A), 'IN/CZYBH') As 编号,
         Extractvalue(Value(A), 'IN/CZYXM') As 姓名, Extractvalue(Value(A), 'IN/YZID') As 医嘱id
  Into n_Type, n_挂号id, n_科室id, n_病区id, v_床号, v_人员编号, v_人员姓名, n_医嘱id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If n_Type = 1 Then
    --住院预约登记
    Select a.病人id, a.No, Decode(a.急诊, 1, '急诊', Null)
    Into n_病人id, v_No, v_入院方式
    From 病人挂号记录 A
    Where a.Id = n_挂号id;
  
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
  
    If r_Advice.紧急标志 = 1 Then
      v_入院方式 := '急诊';
    End If;
  
    Open c_Pati(n_病人id);
    Fetch c_Pati
      Into r_Pati;
  
    --当前操作人员
    If v_人员编号 Is Null Or v_人员姓名 Is Null Then
      v_Temp     := Zl_Identity;
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
      v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    End If;
  
    --删除留观记录和住院预约记录不能并存
    Begin
      Select Count(1) Into n_Count From 病案主页 Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = 0;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      Zl_入院病案主页_Delete(r_Advice.病人id, 0, 0, 0);
      n_Count := 0;
    End If;
  
    If n_Count = 0 Then
      Select Count(1) Into n_Count From 病案主页 Where 病人id = r_Advice.病人id And 出院日期 Is Null;
    End If;
    If n_Count = 0 Then
      Select Count(1)
      Into n_Count
      From 病案主页
      Where 病人id = r_Advice.病人id And (入院日期 >= r_Advice.开始执行时间 Or 出院日期 >= r_Advice.开始执行时间);
    End If;
  
    If n_Count = 0 Then
      Zl_入院病案主页_Insert(1, 0, r_Pati.病人id, r_Pati.住院号, Null, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.费别, r_Pati.出生日期,
                       r_Pati.国籍, r_Pati.民族, r_Pati.学历, r_Pati.婚姻状况, r_Pati.职业, r_Pati.身份, r_Pati.身份证号, r_Pati.出生地点,
                       r_Pati.家庭地址, r_Pati.家庭地址邮编, r_Pati.家庭电话, r_Pati.户口地址, r_Pati.户口地址邮编, r_Pati.联系人姓名, r_Pati.联系人关系,
                       r_Pati.联系人地址, r_Pati.联系人电话, r_Pati.工作单位, r_Pati.合同单位id, r_Pati.单位电话, r_Pati.单位邮编, r_Pati.单位开户行,
                       r_Pati.单位帐号, r_Pati.担保人, r_Pati.担保额, r_Pati.担保性质, n_科室id, Null, Null, v_入院方式, Null, Null,
                       r_Advice.开嘱医生, r_Pati.籍贯, r_Pati.区域, r_Advice.开始执行时间, Null, Null, r_Pati.医疗付款方式, Null, Null, Null,
                       Null, Null, Null, r_Pati.险类, v_人员编号, v_人员姓名, 0, Null, Null, 0, Null, Null, Null, Null, Null, Null,
                       Null, n_挂号id);
    End If;
  
    --更新病区和床号
    Update 病案主页
    Set 入院病床 = v_床号, 出院病床 = v_床号, 入院病区id = n_病区id, 当前病区id = n_病区id
    Where 病人id = r_Pati.病人id And 主页id = 0;
    --换床检查床位是否为空
    Select Count(*) Into n_Count From 床位状况记录 Where 病区id = n_病区id And 床号 = v_床号 And 状态 = '空床';
    If n_Count = 0 Then
      v_Error := '操作失败,床位 ' || v_床号 || ' 不是空床！' || Chr(13) || Chr(10) || '这可能是由于网络并发操作引起的,请刷新后再试！';
      Raise Err_Custom;
    End If;
    --将床位进行占用
    Update 床位状况记录
    Set 状态 = '占用', 病人id = r_Pati.病人id, 科室id = Decode(共用, 1, n_科室id, 科室id)
    Where 病区id = n_病区id And 床号 = v_床号;
  Else
    --取消登记
    Select b.病人id Into n_病人id From 病案主页 B Where b.挂号id = n_挂号id;
    Zl_入院病案主页_Delete(n_病人id, 0);
  End If;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Err_Custom Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || v_Error || '</MSG></ERROR></OUTPUT>');
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Outpatireg;
/

--127912:余伟节,2018-07-03,预约病人接收
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
  v_床号      病案主页.入院病床%Type;
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
  v_床号 := 床号_In;
  If 登记模式_In = 2 And v_床号 Is Null Then
    Select Count(1) Into v_Count From 病案主页 Where 病人id = 病人id_In And 主页id = 0;
    If v_Count = 0 Then
      v_Error := '病人预约记录不存在,不能继续操作!';
      Raise Err_Custom;
    End If;
    Select 入院病床 Into v_床号 From 病案主页 Where 病人id = 病人id_In And 主页id = 0;
  End If;
  If 登记模式_In <> 1 Then
    Update 病人信息
    Set 主页id = v_主页id, 当前病区id = 入院病区id_In, 当前科室id = 入院科室id_In, 当前床号 = Decode(v_床号, '家庭病床', Null, v_床号), 入院时间 = 入院时间_In,
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
  If v_床号 Is Null Then
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
        入院方式 = 入院方式_In, 入院属性 = 入院属性_In, 二级院转入 = 二级院转入_In, 住院目的 = 住院目的_In, 入院病床 = Decode(v_床号, '家庭病床', Null, v_床号),
        是否陪伴 = 是否陪伴_In, 当前病况 = 入院病况_In, 当前病区id = 入院病区id_In, 护理等级id = Decode(护理等级id_In, 0, Null, 护理等级id_In),
        出院科室id = 入院科室id_In, 出院病床 = Decode(v_床号, '家庭病床', Null, v_床号), 门诊医师 = 门诊医师_In, 编目员编号 = 操作员编号_In, 编目员姓名 = 操作员姓名_In,
        姓名 = 姓名_In, 性别 = 性别_In, 年龄 = 年龄_In, 婚姻状况 = 婚姻状况_In, 职业 = 职业_In, 国籍 = 国籍_In, 学历 = 学历_In, 单位电话 = 单位电话_In,
        单位邮编 = 单位邮编_In, 单位地址 = 工作单位_In, 区域 = 区域_In, 家庭地址 = 家庭地址_In, 家庭电话 = 家庭电话_In, 家庭地址邮编 = 家庭地址邮编_In, 户口地址 = 户口地址_In,
        户口地址邮编 = 户口地址邮编_In, 联系人姓名 = 联系人姓名_In, 联系人关系 = 联系人关系_In, 联系人地址 = 联系人地址_In, 联系人身份证号 = 联系人身份证号_In, 联系人电话 = 联系人电话_In,
        医疗付款方式 = 付款方式_In, 备注 = 备注_In, 险类 = 险类_In, 状态 = Decode(v_床号, Null, 1, 0), 登记人 = 操作员姓名_In, 登记时间 = v_Date,
        再入院 = 再入院_In, 病人类型 = 病人类型_In
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
       Decode(v_床号, '家庭病床', Null, v_床号), 是否陪伴_In, 入院病况_In, 入院病区id_In, Decode(护理等级id_In, 0, Null, 护理等级id_In), 入院科室id_In,
       Decode(v_床号, '家庭病床', Null, v_床号), 门诊医师_In, 操作员编号_In, 操作员姓名_In, Decode(v_床号, Null, 1, 0), 姓名_In, 性别_In, 年龄_In,
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
    If v_床号 Is Not Null And v_床号 <> '家庭病床' Then
      Select 等级id Into v_等级id From 床位状况记录 Where 病区id = 入院病区id_In And 床号 = v_床号;
    End If;
  
    --如果同时入科,则入院和入科填写到一条入院变动
    Insert Into 病人变动记录
      (ID, 病人id, 主页id, 开始时间, 开始原因, 附加床位, 病区id, 科室id, 护理等级id, 床位等级id, 床号, 病情, 操作员编号, 操作员姓名)
    Values
      (病人变动记录_Id.Nextval, 病人id_In, v_主页id, 入院时间_In, 1, 0, 入院病区id_In, 入院科室id_In, Decode(护理等级id_In, 0, Null, 护理等级id_In),
       v_等级id, Decode(v_床号, '家庭病床', Null, v_床号), 入院病况_In, 操作员编号_In, 操作员姓名_In);
  
    Insert Into 病人自动计算
      (ID, 病人id, 主页id, 开始时间, 开始原因, 性质, 病区id, 科室id, 护理等级id, 操作员编号, 操作员姓名)
    Values
      (病人自动计算_Id.Nextval, 病人id_In, v_主页id, 入院时间_In, 1, 1, 入院病区id_In, 入院科室id_In, Decode(护理等级id_In, 0, Null, 护理等级id_In),
       操作员编号_In, 操作员姓名_In);
    Insert Into 病人自动计算
      (ID, 病人id, 主页id, 开始时间, 开始原因, 性质, 附加床位, 病区id, 科室id, 床位等级id, 床号, 操作员编号, 操作员姓名)
    Values
      (病人自动计算_Id.Nextval, 病人id_In, v_主页id, 入院时间_In, 1, 2, 0, 入院病区id_In, 入院科室id_In, v_等级id,
       Decode(v_床号, '家庭病床', Null, v_床号), 操作员编号_In, 操作员姓名_In);
    Insert Into 病人自动计算
      (ID, 病人id, 主页id, 开始时间, 开始原因, 性质, 附加床位, 病区id, 科室id, 床位等级id, 床号, 操作员编号, 操作员姓名)
    Values
      (病人自动计算_Id.Nextval, 病人id_In, v_主页id, 入院时间_In, 1, 3, 0, 入院病区id_In, 入院科室id_In, v_等级id,
       Decode(v_床号, '家庭病床', Null, v_床号), 操作员编号_In, 操作员姓名_In);
  
    --同时入科且非家庭病床时床位被占用
    If v_床号 Is Not Null And v_床号 <> '家庭病床' And 登记模式_In <> 2 Then
      Select Count(*) Into v_Count From 床位状况记录 Where 病区id = 入院病区id_In And 床号 = v_床号 And 状态 = '空床';
    
      If v_Count = 0 Then
        v_Error := '操作失败,床位 ' || v_床号 || ' 不是空床！';
        Raise Err_Custom;
      End If;
    
      Update 床位状况记录
      Set 状态 = '占用', 病人id = 病人id_In, 科室id = Decode(共用, 1, 入院科室id_In, 科室id)
      Where 病区id = 入院病区id_In And 床号 = v_床号;
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
  
    If 登记模式_In = 0 And v_床号 Is Not Null Then
      If 再入院_In = 0 Then
        Zl_电子病历时机_Insert(病人id_In, 主页id_In, 2, '入院', 入院科室id_In, Null, 入院时间_In, 入院时间_In);
      Else
        Zl_电子病历时机_Insert(病人id_In, 主页id_In, 2, '再次入院', 入院科室id_In, Null, 入院时间_In, 入院时间_In);
      End If;
    End If;
  
    If v_床号 Is Not Null Then
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

--127912:余伟节,2018-07-03,预约病人处理
CREATE OR REPLACE Procedure Zl_入院病案主页_Delete
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
  v_入院病区   病案主页.入院病区id%Type;
  v_床号       病案主页.入院病床%Type;

  n_病人性质 病案主页.病人性质%Type;
  n_主页id   病案主页.主页id%Type;

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
      --预约中心的预约病人存在床位记录 
      Select 入院病区id, 入院病床 Into v_入院病区, v_床号 From 病案主页 Where 病人id = 病人id_In And 主页id = 0;
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

  b_Message.Zlhis_Patient_006(病人id_In, 主页id_In, '入院登记');

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
    --预约中心的预约病人需要释放床位
    If v_床号 Is Not Null Then
      Update 床位状况记录 Set 状态 = '空床', 病人id = Null Where 病区id = v_入院病区 And 床号 = v_床号;
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
                主页id = (Select Max(主页id) From 病案主页 Where 病人id = 病人id_In And 主页id < 主页id_In);
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





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0019' Where 编号=&n_System;
Commit;
