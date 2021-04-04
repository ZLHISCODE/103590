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
--134222:焦博,2018-11-23,调整Oracle过程Zl_Third_Getregistalter,分段创建返回的XML字符串
Create Or Replace Procedure Zl_Third_Getregistalter
(
  Xml_In  Xmltype,
  Xml_Out Out Xmltype
) Is
  -----------------------------------------------
  --功能：获取当天操作的停换诊安排
  --入参：XML_IN
  --<IN>
  --  <JSKLB>结算卡类别</JSKLB>
  --  <RQ>日期</RQ>
  --</IN>
  --出参:XML_OUT
  --<OUTPUT>
  --  <TZLISTS>          //停诊列表
  --    <ITEM>
  --      <HM>号码</HM>
  --      <YSID>医生ID</YSID>
  --      <YS>医生姓名</YS>
  --      <KSSJ>停诊开始时间</KSSJ>
  --      <JSSJ>停诊结束时间</JSSJ>
  --      <BRLIST>
  --        <INFO>
  --          <YYNO>预约单据号</YYNO>
  --          <BRID>病人ID</BRID>
  --          <YYSJ>预约时间</YYSJ>
  --          <CZSJ>操作时间</CZSJ>
  --          <YYKS>预约科室</YYKS>
  --          <GHLX>号类</GHLX>
  --          <YSXM>医生姓名</YSXM>
  --        </INFO>
  --      </BRLIST>
  --    </ITEM>
  --  </TZLISTS>
  --  <HZLISTS>          //换诊列表
  --    <ITEM>
  --      <BRID>病人ID</BRID>
  --      <YYSJ>预约的操作时间</YYSJ>
  --      <YSJ>原预约时间</YSJ>
  --      <YHM>原号码</YHM>
  --      <YYS>原医生</YYS>
  --      <YZC>原医生的职称</YZC>
  --      <XSJ>现预约时间</XSJ>
  --      <XHM>现号码</XHM>
  --      <XYS>现医生</XYS>
  --      <XZC>现医生的职称</XZC>
  --    </ITEM>
  --  </HZLIST>
  --</OUTPUT>
  -----------------------------------------------------

  d_Date     Date;
  v_Jsklb    Varchar2(100);
  n_卡类别id 医疗卡类别.Id%Type;
  n_Cnt      Number(3);
  v_Temp     Clob;
  v_Brinfo   Varchar2(4000);
  d_启用时间 Date;
  v_Para     Varchar2(2000);
  n_Exists   Number(3);
  n_挂号模式 Number(3);
  x_Templet  Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/JSKLB') Into v_Jsklb From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'yyyy-mm-dd')
  Into d_Date
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select b.Id Into n_卡类别id From 医疗卡类别 B Where b.名称 = v_Jsklb And Rownum < 2;
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  v_Para     := zl_GetSysParameter(256);
  n_挂号模式 := Substr(v_Para, 1, 1);
  Begin
    d_启用时间 := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_启用时间 := Null;
  End;

  If n_挂号模式 = 1 And Nvl(d_Date, Sysdate) > Nvl(d_启用时间, Sysdate - 30) Then
    --出诊表排班模式
    --获取停诊安排
    For r_停诊 In (Select a.Id As 记录id, b.号码, a.医生id, a.医生姓名, a.停诊开始时间, a.停诊终止时间
                 From 临床出诊记录 A, 临床出诊号源 B, 临床出诊停诊记录 C
                 Where a.Id = c.记录id And a.号源id = b.Id And a.停诊开始时间 Is Not Null And c.审批时间 Between d_Date And
                       d_Date + 1 - 1 / 24 / 60 / 60) Loop
      v_Temp := v_Temp || '<ITEM><HM>' || r_停诊.号码 || '</HM><YSID>' || r_停诊.医生id || '</YSID><YS>' || r_停诊.医生姓名 ||
                '</YS><KSSJ>' || r_停诊.停诊开始时间 || '</KSSJ><JSSJ>' || r_停诊.停诊终止时间 || '</JSSJ><BRLIST>';
      For r_停诊病人 In (Select a.记录性质, a.No, a.病人id, To_Char(a.发生时间, 'yyyy-mm-dd') As 发生时间,
                            To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, b.名称, d.号类, c.医生姓名 As 医生姓名
                     From 病人挂号记录 A, 部门表 B, 临床出诊记录 C, 临床出诊号源 D
                     Where a.执行部门id = b.Id And a.出诊记录id = c.Id And c.号源id = d.Id And 记录状态 = 1 And
                           发生时间 Between r_停诊.停诊开始时间 And r_停诊.停诊终止时间 And a.出诊记录id = r_停诊.记录id And Not Exists
                      (Select 1 From 就诊变动记录 Where 挂号单 = a.No)) Loop
        --停诊病人列表，不包含已经换诊和取消了的病人
        If r_停诊病人.记录性质 = 2 Then
          v_Brinfo := '<INFO><YYNO>' || r_停诊病人.No || '</YYNO><BRID>' || r_停诊病人.病人id || '</BRID><YYSJ>' || r_停诊病人.发生时间 ||
                      '</YYSJ><CZSJ>' || r_停诊病人.登记时间 || '</CZSJ>' || '<YYKS>' || r_停诊病人.名称 || '</YYKS><GHLX>' ||
                      r_停诊病人.号类 || '</GHLX><YSXM>' || r_停诊病人.医生姓名 || '</YSXM></INFO>';
          v_Temp   := v_Temp || v_Brinfo;
        Else
          Begin
            Select 1
            Into n_Exists
            From 病人预交记录
            Where NO = r_停诊病人.No And 记录性质 = 4 And 卡类别id = n_卡类别id;
          Exception
            When Others Then
              n_Exists := 0;
          End;
          If n_Exists = 1 Then
            v_Brinfo := '<INFO><YYNO>' || r_停诊病人.No || '</YYNO><BRID>' || r_停诊病人.病人id || '</BRID><YYSJ>' || r_停诊病人.发生时间 ||
                        '</YYSJ><CZSJ>' || r_停诊病人.登记时间 || '</CZSJ>' || '<YYKS>' || r_停诊病人.名称 || '</YYKS><GHLX>' ||
                        r_停诊病人.号类 || '</GHLX><YSXM>' || r_停诊病人.医生姓名 || '</YSXM></INFO>';
            v_Temp   := v_Temp || v_Brinfo;
          End If;
        End If;
        v_Brinfo := '';
      End Loop;
      v_Temp := v_Temp || '</BRLIST></ITEM>';
    End Loop;
    v_Temp := '<TZLISTS>' || v_Temp || '</TZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    --换取换诊列表
    v_Temp := '';
    For r_换诊 In (Select d.记录性质, d.No, a.病人id, To_Char(d.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 预约时间,
                        To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, a.原号码, a.原医生姓名, b.专业技术职务 As 原职务, a.现号码, a.现医生姓名,
                        c.专业技术职务 As 现职务
                 From 就诊变动记录 A, 人员表 B, 人员表 C, 病人挂号记录 D
                 Where a.登记时间 Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60 And a.原医生id = b.Id And a.现医生id = c.Id And
                       a.挂号单 = d.No) Loop
      --只返回该卡类别挂号的病人         
      If r_换诊.记录性质 = 2 Then
        v_Temp := v_Temp || '<ITEM><BRID>' || r_换诊.病人id || '</BRID><YYSJ>' || r_换诊.登记时间 || '</YYSJ>';
        v_Temp := v_Temp || '<YSJ>' || r_换诊.预约时间 || '</YSJ><YHM>' || r_换诊.原号码 || '</YHM><YYS>' || r_换诊.原医生姓名 ||
                  '</YYS><YZC>' || r_换诊.原职务 || '</YZC>';
        v_Temp := v_Temp || '<XSJ>' || r_换诊.预约时间 || '</XSJ><XHM>' || r_换诊.现号码 || '</XHM><XYS>' || r_换诊.现医生姓名 ||
                  '</XYS><XZC>' || r_换诊.现职务 || '</XZC></ITEM>';
      Else
        Begin
          Select 1 Into n_Exists From 病人预交记录 Where NO = r_换诊.No And 记录性质 = 4 And 卡类别id = n_卡类别id;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists = 1 Then
          v_Temp := v_Temp || '<ITEM><BRID>' || r_换诊.病人id || '</BRID><YYSJ>' || r_换诊.登记时间 || '</YYSJ>';
          v_Temp := v_Temp || '<YSJ>' || r_换诊.预约时间 || '</YSJ><YHM>' || r_换诊.原号码 || '</YHM><YYS>' || r_换诊.原医生姓名 ||
                    '</YYS><YZC>' || r_换诊.原职务 || '</YZC>';
          v_Temp := v_Temp || '<XSJ>' || r_换诊.预约时间 || '</XSJ><XHM>' || r_换诊.现号码 || '</XHM><XYS>' || r_换诊.现医生姓名 ||
                    '</XYS><XZC>' || r_换诊.现职务 || '</XZC></ITEM>';
        End If;
      End If;
    End Loop;
    v_Temp := '<HZLISTS>' || v_Temp || '</HZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --计划排班模式
    --获取停诊安排
    For Rs In (Select b.号码, b.医生id, b.医生姓名, To_Char(a.开始停止时间, 'yyyy-mm-dd hh24:mi:ss') As 开始停止时间,
                      To_Char(a.结束停止时间, 'yyyy-mm-dd hh24:mi:ss') As 结束停止时间
               From 挂号安排停用状态 A, 挂号安排 B
               Where a.安排id = b.Id And a.制订日期 Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60) Loop
      v_Temp := v_Temp || '<ITEM><HM>' || Rs.号码 || '</HM><YSID>' || Rs.医生id || '</YSID><YS>' || Rs.医生姓名 ||
                '</YS><KSSJ>' || Rs.开始停止时间 || '</KSSJ><JSSJ>' || Rs.结束停止时间 || '</JSSJ><BRLIST>';
      ----2015/7/28
      For Rs_Br In (Select a.No, a.病人id, To_Char(a.发生时间, 'yyyy-mm-dd') As 发生时间,
                           To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, b.名称, c.号类, a.执行人 As 医生姓名
                    From 病人挂号记录 A, 部门表 B, 挂号安排 C
                    Where a.号别 = Rs.号码 And a.执行状态 = 0 And a.执行部门id = b.Id And b.Id = c.科室id And a.号别 = c.号码 And
                          Trunc(发生时间) Between Trunc(To_Date(Rs.开始停止时间, 'yyyy-mm-dd hh24:mi:ss')) And
                          Trunc(To_Date(Rs.结束停止时间, 'yyyy-mm-dd hh24:mi:ss'))) Loop
        --只返回该卡类别挂号的病人
        Select Count(*)
        Into n_Cnt
        From (Select 1
               From 病人预交记录 A
               Where a.No = Rs_Br.No And a.记录性质 = 4 And a.记录状态 = 1 And a.病人id = Rs_Br.病人id And 卡类别id = n_卡类别id
               Union All
               Select 1 From 病人挂号记录 Where NO = Rs_Br.No And 记录状态 = 1 And 交易说明 = v_Jsklb);
        If n_Cnt > 0 Then
          v_Brinfo := '<INFO><YYNO>' || Rs_Br.No || '</YYNO><BRID>' || Rs_Br.病人id || '</BRID><YYSJ>' || Rs_Br.发生时间 ||
                      '</YYSJ><CZSJ>' || Rs_Br.登记时间 || '</CZSJ>' || '<YYKS>' || Rs_Br.名称 || '</YYKS><GHLX>' || Rs_Br.号类 ||
                      '</GHLX><YSXM>' || Rs_Br.医生姓名 || '</YSXM></INFO>';
          v_Temp   := v_Temp || v_Brinfo;
        End If;
        v_Brinfo := '';
      End Loop;
      v_Temp := v_Temp || '</BRLIST></ITEM>';
    End Loop;
    v_Temp := '<TZLISTS>' || v_Temp || '</TZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    --获取换诊记录
    v_Temp := '';
    For Rs In (Select d.No, a.病人id, To_Char(d.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 预约时间,
                      To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间, a.原号码, a.原医生姓名, b.专业技术职务 As 原职务, a.现号码, a.现医生姓名,
                      c.专业技术职务 As 现职务
               From 就诊变动记录 A, 人员表 B, 人员表 C, 病人挂号记录 D
               Where a.登记时间 Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60 And a.原医生id = b.Id And a.现医生id = c.Id And
                     a.挂号单 = d.No) Loop
      --只返回该卡类别挂号的病人         
      Select Count(*)
      Into n_Cnt
      From (Select 1
             From 病人预交记录 A
             Where a.No = Rs.No And a.记录性质 = 4 And a.记录状态 = 1 And a.病人id = Rs.病人id And 卡类别id = n_卡类别id
             Union All
             Select 1 From 病人挂号记录 Where NO = Rs.No And 记录状态 = 1 And 交易说明 = v_Jsklb);
      If n_Cnt > 0 Then
        v_Temp := v_Temp || '<ITEM><BRID>' || Rs.病人id || '</BRID><YYSJ>' || Rs.登记时间 || '</YYSJ>';
        v_Temp := v_Temp || '<YSJ>' || Rs.预约时间 || '</YSJ><YHM>' || Rs.原号码 || '</YHM><YYS>' || Rs.原医生姓名 || '</YYS><YZC>' ||
                  Rs.原职务 || '</YZC>';
        v_Temp := v_Temp || '<XSJ>' || Rs.预约时间 || '</XSJ><XHM>' || Rs.现号码 || '</XHM><XYS>' || Rs.现医生姓名 || '</XYS><XZC>' ||
                  Rs.现职务 || '</XZC></ITEM>';
      End If;
    End Loop;
    v_Temp := '<HZLISTS>' || v_Temp || '</HZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getregistalter;
/

--134551:胡俊勇,2018-11-22,三方接口过程产生锚点
Create Or Replace Procedure Zl_Third_Buildpatient
(
  Patiinfo_In  In Xmltype,
  Patiinfo_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------
  --参数说明:
  -- 入参 Patiinfo_In:
  --<IN>
  --  <ZJH></ZJH>                 //证件号，目前仅支持身份证号
  --  <ZJLX></ZJLX>                       //证件类型(目前仅支持身份证,为空时默认为身份证)
  --  <XM></XM>                       //姓名
  --  <SJH></SJH>                      //手机号
  --</IN>

  --出参 Patiinfo_Out：
  --<OUTPUT>
  --       <BRID></BRID>                //病人ID
  --       <MZH></MZH>                  //门诊号
  --     <ERROR></ERROR>         //如果有错误返回该节点
  --</OUTPUT>
  -------------------------------------------------------------------------------
  n_Pati_Id      病人信息.病人id%Type;
  n_Card_Type_Id 医疗卡类别.Id%Type;
  n_Count        Number(5);
  n_Sum          Number(5);
  v_校验位       Varchar2(50);

  v_姓名         病人信息.姓名%Type;
  v_身份证号     病人信息.身份证号%Type;
  v_手机号       病人信息.家庭电话%Type;
  v_性别         病人信息.性别%Type;
  v_年龄         病人信息.年龄%Type;
  v_操作员       人员表.姓名%Type;
  v_医疗付款方式 病人信息.医疗付款方式%Type;
  n_门诊号       病人信息.门诊号%Type;
  v_证件类型     医疗卡类别.名称%Type;
  v_证件号       病人医疗卡信息.卡号%Type;

  v_Pattern Varchar2(500);
  v_Temp    Varchar2(32767); --临时XML
  v_Err_Msg Varchar2(2000);
  n_存在    Number(2);

  d_出生日期  病人信息.出生日期%Type;
  d_Curr_Time Date;

  Err_Item Exception;
Begin
  Patiinfo_Out := Xmltype('<OUTPUT></OUTPUT>');
  Select Sysdate Into d_Curr_Time From Dual;

  --新建病人：姓名、身份证号、手机号（存在家庭电话中）、出生日期、性别、年龄(后面三项可从身份证中获取)。
  Select Extractvalue(Value(I), 'IN/XM'), Extractvalue(Value(I), 'IN/ZJH'), Extractvalue(Value(I), 'IN/SJH'),
         Extractvalue(Value(I), 'IN/ZJLX')
  Into v_姓名, v_证件号, v_手机号, v_证件类型
  From Table(Xmlsequence(Extract(Patiinfo_In, 'IN'))) I;

  Begin
    If v_证件类型 Is Null Then
      Select 病人id
      Into n_Pati_Id
      From 病人医疗卡信息
      Where 卡号 = v_证件号 And 卡类别id In (Select ID From 医疗卡类别 Where 名称 Like '%身份证%') And Rownum < 2;
    Else
      Select 病人id
      Into n_Pati_Id
      From 病人医疗卡信息
      Where 卡号 = v_证件号 And 卡类别id In (Select ID From 医疗卡类别 Where 名称 = v_证件类型) And Rownum < 2;
    End If;
    n_存在 := 1;
  Exception
    When Others Then
      n_存在 := 0;
  End;

  If Nvl(n_存在, 0) = 1 Then
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    Select 门诊号 Into n_门诊号 From 病人信息 Where 病人id = n_Pati_Id;
    If n_门诊号 Is Null Then
      n_门诊号 := Nextno(3);
      Update 病人信息 Set 门诊号 = n_门诊号 Where 病人id = n_Pati_Id;
    End If;
    v_Temp := '<MZH>' || n_门诊号 || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
  Else
    If v_姓名 Is Null Then
      v_Err_Msg := '传入姓名为空!';
      Raise Err_Item;
    End If;
    If v_证件类型 Like '%身份证%' Or v_证件类型 Is Null Then
      v_身份证号 := v_证件号;
    Else
      v_Err_Msg := '目前不支持身份证以外的方式建档！';
      Raise Err_Item;
    End If;
  
    If v_身份证号 Is Null Then
      v_Err_Msg := '传入身份证号为空!';
      Raise Err_Item;
    Else
      --身份证合法验证
      v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,83,91';
    
      --地区检验
      If Instr(v_Pattern, Substr(v_身份证号, 1, 2)) = 0 Then
        v_Err_Msg := '身份证前两位地区码不正确!';
        Raise Err_Item;
      End If;
      --身份证长度检查
      If Length(v_身份证号) = 15 Then
        --检查身份证号:15位身份证号要求全部为数字
        v_Pattern := '^\d{15}$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_身份证号, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中包含非法字符，请检查!';
          Raise Err_Item;
        End If;
        --获取性别
        If Mod(To_Number(Substr(v_身份证号, 15, 1)), 2) = 1 Then
          v_性别 := '男';
        Else
          v_性别 := '女';
        End If;
        --出生日期的合法性检查
      
        v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(v_身份证号, 7, 6), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Raise Err_Item;
        Else
          --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
          If Instr(',0229,0230,', ',' || Substr(v_身份证号, 9, 4) || ',') > 0 Then
            v_Temp     := '19' || Substr(v_身份证号, 7, 2) || '0301';
            d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_出生日期 := To_Date('19' || Substr(v_身份证号, 7, 6), 'yyyy-mm-dd');
          End If;
          If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '身份证中的出生日期无效，请检查!';
            Raise Err_Item;
          End If;
        End If;
      Elsif Length(v_身份证号) = 18 Then
        -- 18 位身份证号前17 位全部为数字，最后1位可为数字或x
        v_Pattern := '^\d{17}[0-9Xx]$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_身份证号, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中包含非法字符!';
          Raise Err_Item;
        End If;
        --获取性别
        If Mod(To_Number(Substr(v_身份证号, 17, 1)), 2) = 1 Then
          v_性别 := '男';
        Else
          v_性别 := '女';
        End If;
        --出生日期的合法性检查
        v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(v_身份证号, 7, 8), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '身份证中的出生日期无效，请检查!';
          Raise Err_Item;
        Else
          --以前的老身份证没有区分闰年的情况兼容处理将出生日期改为2月28号，如：19470229这种情况
          If Instr(',0229,0230,', ',' || Substr(v_身份证号, 11, 4) || ',') > 0 Then
            v_Temp     := Substr(v_身份证号, 7, 4) || '0301';
            d_出生日期 := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_出生日期 := To_Date(Substr(v_身份证号, 7, 8), 'yyyy-mm-dd');
          End If;
          If d_出生日期 > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '身份证中的出生日期无效，请检查!';
            Raise Err_Item;
          End If;
          --计算校验位
          n_Sum     := (To_Number(Substr(v_身份证号, 1, 1)) + To_Number(Substr(v_身份证号, 11, 1))) * 7 +
                       (To_Number(Substr(v_身份证号, 2, 1)) + To_Number(Substr(v_身份证号, 12, 1))) * 9 +
                       (To_Number(Substr(v_身份证号, 3, 1)) + To_Number(Substr(v_身份证号, 13, 1))) * 10 +
                       (To_Number(Substr(v_身份证号, 4, 1)) + To_Number(Substr(v_身份证号, 14, 1))) * 5 +
                       (To_Number(Substr(v_身份证号, 5, 1)) + To_Number(Substr(v_身份证号, 15, 1))) * 8 +
                       (To_Number(Substr(v_身份证号, 6, 1)) + To_Number(Substr(v_身份证号, 16, 1))) * 4 +
                       (To_Number(Substr(v_身份证号, 7, 1)) + To_Number(Substr(v_身份证号, 17, 1))) * 2 +
                       To_Number(Substr(v_身份证号, 8, 1)) * 1 + To_Number(Substr(v_身份证号, 9, 1)) * 6 +
                       To_Number(Substr(v_身份证号, 10, 1)) * 3;
          n_Count   := Mod(n_Sum, 11);
          v_Pattern := '10X98765432';
          v_校验位  := Substr(v_Pattern, n_Count + 1, 1);
          If v_校验位 <> Upper(Substr(v_身份证号, 18, 1)) Then
            v_Err_Msg := '身份证号码不正确，请检查。';
            Raise Err_Item;
          End If;
        End If;
      Else
        v_Err_Msg := '身份证长度不对,请检查。';
        Raise Err_Item;
      End If;
    
      If Nvl(v_年龄, '_') = '_' Then
        v_年龄 := Zl_Age_Calc(0, d_出生日期, d_Curr_Time);
      End If;
    End If;
  
    Select 名称 Into v_医疗付款方式 From 医疗付款方式 Where 缺省标志 = 1;
    n_Pati_Id := Nextno(1);
    n_门诊号  := Nextno(3);
    Insert Into 病人信息
      (病人id, 姓名, 身份证号, 家庭电话, 出生日期, 性别, 年龄, 登记时间, 门诊号, 医疗付款方式, 手机号)
      Select n_Pati_Id, v_姓名, v_身份证号, v_手机号, d_出生日期, v_性别, v_年龄, d_Curr_Time, n_门诊号, v_医疗付款方式, v_手机号




      From Dual;
    --病人信息保存完后，完成医疗卡绑定（二代身份证卡类别的绑定）
    Begin
      If v_证件类型 Is Null Then
        Select ID Into n_Card_Type_Id From 医疗卡类别 Where 名称 Like '%身份证%' And Rownum < 2;
      Else
        Select ID Into n_Card_Type_Id From 医疗卡类别 Where 名称 = v_证件类型 And Rownum < 2;
      End If;
    Exception
      When No_Data_Found Then
        v_Err_Msg := '身份证卡类别不存在！';
        Raise Err_Item;
    End;
    Select b.姓名 Into v_操作员 From 上机人员表 A, 人员表 B Where a.人员id = b.Id And a.用户名 = User;
  
    Zl_医疗卡变动_Insert(11, n_Pati_Id, n_Card_Type_Id, Null, v_身份证号, '创建虚拟卡', Null, v_操作员, d_Curr_Time);
  
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    v_Temp := '<MZH>' || n_门诊号 || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
	 b_Message.Zlhis_Patient_015(n_Pati_Id);
  End If;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Buildpatient;
/

--129503:陈刘,2018-11-22,护理项目保存接口
Create Or Replace Procedure Zl_Third_Tendfile_Itemsave
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  n_Fileid       Number(18);
  n_格式id       Number(18);
  n_Xh           Number(5);
  n_Brid         Number(18);
  n_Zyid         Number(5);
  n_Babby        Number(1);
  v_Czy          Varchar2(20);
  n_Newadd       Number(1);
  n_Kind         Number(1);
  n_Num          Number(1);
  Intins         Number(1);
  n_归档         Number(1);
  v_科室id       Number(18);
  v_Name         Varchar2(20);
  d_婴儿出院时间 Date;
  v_Error        Varchar2(255);
  v_Temp         Varchar2(32767);
  x_Templet      Xmltype; --模板XML
  Err_Custom Exception;
Begin

  Select To_Number(Extractvalue(Value(A), 'IN/BRID'))
  Into n_Brid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/ZYID'))
  Into n_Zyid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/YEXH'))
  Into n_Babby
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Char(Extractvalue(Value(A), 'IN/CZY')) 
  Into v_Czy 
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Begin
    Select Max(1)
    Into n_归档
    From 病人护理文件
    Where 病人id = n_Brid And 主页id = n_Zyid And 婴儿 = n_Babby And 归档时间 Is Null;
  End;

  For r_Input In (Select 发生时间, Lx, Ly, Mc, Nr, Bw, Wj
                  From Xmltable('$a/IN/ITEMLIST/ITEM' Passing Xmlfilelist_In As "a" Columns 发生时间 Varchar2(20) Path
                                 'TIME', Lx Number(1) Path 'LX', Ly Number(2) Path 'LY', Mc Varchar2(20) Path 'MC',
                                 Nr Varchar2(20) Path 'NR', Bw Varchar2(10) Path 'BW', Wj Varchar2(4000) Path 'Wj') B) Loop
    If r_Input.Mc Is Null Then
      v_Error := '未录入数据，不允许操作，请检查！';
      Raise Err_Custom;
    Else
      Select Max(项目序号) Into n_Xh From 护理记录项目 Where 项目名称 = r_Input.Mc;
    End If;
  
    If n_Babby <> 0 Then
      Begin
        Select 开始执行时间
        Into d_婴儿出院时间
        From 病人医嘱记录 B, 诊疗项目目录 C
        Where b.诊疗项目id + 0 = c.Id And b.医嘱状态 = 8 And Nvl(b.婴儿, 0) <> 0 And c.类别 = 'Z' And
              Instr(',3,5,11,', ',' || c.操作类型 || ',', 1) > 0 And b.病人id = n_Brid And b.主页id = n_Zyid And b.婴儿 = n_Babby;
      Exception
        When Others Then
          d_婴儿出院时间 := Null;
      End;
    End If;
  
    If d_婴儿出院时间 Is Null Then
      v_科室id := 0;
      Begin
        Select a.科室id
        Into v_科室id
        From 病人变动记录 A
        Where a.科室id Is Not Null And a.病人id = n_Brid And a.主页id = n_Zyid And
              (To_Date(To_Char(To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss'), 'YYYY-MM-DD HH24:MI') || '59',
                       'YYYY-MM-DD HH24:MI:SS') >= a.开始时间 And
              (To_Date(To_Char(To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss'), 'YYYY-MM-DD HH24:MI') || '00',
                        'YYYY-MM-DD HH24:MI:SS') <= Nvl(a.终止时间, Sysdate) Or a.终止时间 Is Null)) And Rownum < 2;
      Exception
        When Others Then
          v_科室id := 0;
      End;
      If v_科室id = 0 Then
        v_Error := '数据发生时间 ' || To_Date(r_Input.发生时间, 'YYYY-MM-DD HH24:MI:SS') || ' 不在病人有效变动时间范围内，不允许操作！';
        Raise Err_Custom;
      End If;
    End If;
  
    Select Max(a.Id)
    Into n_Fileid
    From 病人护理文件 A, 病历文件结构 B, 病历文件列表 C
    Where a.格式id = c.Id And 保留 = 0 And 病人id = n_Brid And 主页id = n_Zyid And 婴儿 = n_Babby And a.格式id = b.文件id And
          要素名称 = r_Input.Mc And b.文件id = c.Id And a.开始时间 < To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss') And
          (结束时间 > To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss') Or 结束时间 Is Null)
    Order By a.开始时间;
  
    If n_Fileid Is Null Then
      Select Max(a.Id), Max(c.子类), Max(c.Id)
      Into n_Fileid, n_Kind, n_格式id
      From 病人护理文件 A, 病历文件列表 C
      Where a.格式id = c.Id And 保留 = -1 And 病人id = n_Brid And 主页id = n_Zyid And 婴儿 = n_Babby And
            a.开始时间 < To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss') And
            (结束时间 > To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss') Or 结束时间 Is Null)
      Order By a.开始时间;
      If n_Fileid <> 0 Then
        v_Name := r_Input.Mc;
        If n_Kind = '1' Then
          If n_Xh = 4 Or n_Xh = 5 Then
            Select Max(内容文本)
            Into n_Num
            From 病人护理文件 A, 病历文件结构 B
            Where a.格式id = b.文件id And a.Id = n_Fileid And 要素名称 = '婴儿体温单';
            If Not (n_Num = 1) Then
              v_Name := '血压';
            End If;
          End If;
          Begin
            Select 1
            Into Intins
            From (Select To_Char(f.记录名) As 项目名称, g.项目性质
                   From 体温记录项目 F, 护理记录项目 G
                   Where f.项目序号 = g.项目序号 And g.项目性质 = 2 And
                         (g.适用科室 = 1 Or
                         (g.适用科室 = 2 And Exists
                          (Select 1 From 护理适用科室 D Where g.项目序号 = d.项目序号 And d.科室id = v_科室id))) And Nvl(g.应用方式, 0) <> 0 And
                         (Nvl(g.适用病人, 0) = 0 Or Nvl(g.适用病人, 0) = Decode(n_Babby, 0, 1, 2))
                   Union All
                   Select b.要素名称 As 项目名称, 1 As 项目性质
                   From 病历文件结构 A, 病历文件结构 B
                   Where a.文件id = n_格式id And a.父id Is Null And a.对象序号 In (2, 3) And b.父id = a.Id) H
            Where Instr(',' || h.项目名称 || ',', ',' || v_Name || ',', 1) > 0;
          
          Exception
            When Others Then
              Intins := 0;
          End;
        Else
          Begin
            Select 1
            Into Intins
            From 体温记录项目 F, 护理记录项目 G
            Where f.项目序号 = g.项目序号 And Nvl(g.应用方式, 0) <> 0 And g.护理等级 >= 0 And
                  (Nvl(g.适用病人, 0) = 0 Or Nvl(g.适用病人, 0) = Decode(n_Babby, 0, 1, 2)) And f.项目序号 = n_Xh And
                  (g.适用科室 = 1 Or (g.适用科室 = 2 And Exists
                   (Select 1 From 护理适用科室 D Where g.项目序号 = d.项目序号 And d.科室id = v_科室id)));
          Exception
            When Others Then
              Intins := 0;
          End;
        End If;
      
        If Intins = 1 Then
        
          Zl_体温单数据_Update(n_Fileid, To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss'), 1, n_Xh, r_Input.Nr, r_Input.Bw, 0,
                          Null, 1, r_Input.Ly, Null, 0, 0, Null, Null, v_Czy);
        End If;
      End If;
    Else
      Select Max(1)
      Into n_Newadd
      From 病人护理数据
      Where 发生时间 = To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss') And 文件id = n_Fileid;
      If n_Newadd = 1 Then
        Zl_病人护理数据_Update(n_Fileid, To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss'), 1, n_Xh, r_Input.Nr, r_Input.Bw, 1,
                         r_Input.Ly, 0, v_Czy, Null, Null, Null);
      Else
        Select Max(1)
        Into n_Num
        From 病人护理数据
        Where 发生时间 = To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss') And 签名人 Is Not Null And 文件id = n_Fileid;
        If n_Num = 1 Then
          v_Error := '当前病人的护理文件已签名，不允许修改，请先回退签名！';
          Raise Err_Custom;
        Else
          Zl_病人护理数据_Update(n_Fileid, To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss'), 1, n_Xh, r_Input.Nr, r_Input.Bw, 1,
                           r_Input.Ly, 0, v_Czy, Null, Null, Null);
          Zl_病人护理打印_Update(n_Fileid, To_Date(r_Input.发生时间, 'yyyy-mm-dd hh24:mi:ss'), 1);
        End If;
      End If;
    End If;
  End Loop;
  v_Temp := '<RESULT>True</RESULT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xmlfilelist_Out := x_Templet;

Exception
  When Err_Custom Then
    v_Temp := '<RESULT>False</RESULT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<ERROR><MSG>' || v_Error || '</MSG></ERROR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xmlfilelist_Out := x_Templet;
  
  When Others Then
    v_Temp := '<RESULT>False</RESULT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<ERROR><MSG>' || SQLErrM || '</MSG></ERROR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xmlfilelist_Out := x_Templet;
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Itemsave;
/





------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0038' Where 编号=&n_System;
Commit;
