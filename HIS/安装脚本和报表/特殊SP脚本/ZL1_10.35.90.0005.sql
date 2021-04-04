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
--117772:蒋廷中,2018-04-02,增加系统参数传染病报告卡强制填写
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, -null, -null, 300, '传染病报告卡强制填写', '0',
         '0', '若启用参数，则如果是填写诊断弹出的传染病报告卡则不显示退出按钮且点击关闭X按钮时不关闭窗体。若不启用参数则不控制',
         '0-表示不启用,1-表示启用', Null, '适用于某些医院可以要求医生强制填写传染病报告卡', Null
  From Dual;


--123734:刘鹏飞,2018-03-31,新版护士站病区基本概况信息显示处理
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1265, 1, 0, 0, 0, 0, 0, 14, '按床位编制显示床位状况', 1, NULL,
         '控制新版护士工作站主界面病区基本信息栏床位信息是否按床位编制显示床位使用状况', '0-显示占用床位总数和空床总数；1-显示每种床位编制分类的床位数和空床数', Null, '适用于要查看详细的床位使用情况', Null
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
--122954:余伟节,2018-04-08,中联合理用药
Create Or Replace Procedure Zl_中联合理用药参数_Update
(
  病人id_In In 病人信息.病人id%Type,
  主页id_In In 病案主页.主页id%Type,
  挂号id_In In 病人挂号记录.Id%Type := Null
) As

  --------------------------------------------------------------------------------------------------
  --功能:合理用药监测传入值；返回合理用药传入值
  --出参:Xml_Return  返回补充的XML串
  -- <details_xml>
  --    <patient_info>
  --      <info name="年龄数字" value="28114.45"/>
  --     <info name="年龄周期" value="成人"/>
  --      <info name="性别" value="女"/>
  --      <info name="职业" value="运动员"/>
  --      <info name="妊娠" value="1"/>
  --      <info name="哺乳" value="1"/>
  --      <info name="肝功能不全" value="1">
  --      <info name="严重肝功能不全" value="1">
  --      <info name="肾功能不全" value="1">
  --      <info name="严重肾功能不全" value="1">
  --      <info name="诊断" value="J18.000"/> --诊断传编码，多个诊断以逗号分隔
  --    </patient_info>
  --    <medicine_info>
  --      <medicine>
  --        <info name="医嘱ID" value="1"/>
  --        <info name="本位码" value="86900967000160" main="46d64420-8319-4768-9a11-f4b0f5e4ce7a"/> --main值是固定的
  --        <info name="诊疗项目ID" value="67232" main="4e19df1c-c1b9-4a43-a83d-0741a19961ab"/>
  --        <info name="输液组号" value="1"/>
  --        <info name="计量单位" value="ml"/>
  --        <info name="单次量" value="250"/>
  --        <info name="单次量-按体重" value="5.21"/>--单次量-按体重= trunc(单次量/病人体重,2)
  --        <info name="单次量-按体表" value="170.3"/>--单次量-按体表= trunc(单次量/(0.0061*病人身高+0.0128*病人体重-0.1529),2)
  --        <info name="每日量" value="250"/>
  --        1.每日量=单次量*日频次
  --        2.日频次计算：
  --            a.适用范围=-1，日频次=1
  --            b.间隔单位=天 and 频率间隔=1，日频次=频率次数
  --            c.间隔单位=天 and 频率间隔>1 and 频率次数=1，日频次=1
  --            d.间隔单位=小时 and 频率间隔<=24,日频次=24/频率间隔*频率次数
  --            e.间隔单位=小时 and 频率间隔>24 and 频率次数=1，日频次=1
  --            f.间隔单位=周 and 频率次数=1，日频次=1
  --        <info name="每日量-按体重" value="5.21"/>  --trunc(每日量/病人体重,2)
  --        <info name="每日量-按体表" value="170.3"/>  --每日量-按体表= trunc(每日量/(0.0061*病人身高+0.0128*病人体重-0.1529),2)
  --        <info name="给药频次" value="每日一次"/>
  --        <info name="给药途径" value="001"/>
  --      </medicine>
  --    </medicine_info>
  --  </details_xml>
  --------------------------------------------------------------------------------------------------
  Xml_Ret             Xmltype;
  Xml_Document        Xmldom.Domdocument;
  Xml_Nodelist        Xmldom.Domnodelist;
  Xml_Domelement      Xmldom.Domelement;
  Xml_Domnamednodemap Xmldom.Domnamednodemap;
  Xml_Node_Med        Xmldom.Domnode;
  Xml_Node            Xmldom.Domnode;
  Xml_Node_New        Xmldom.Domnode;
  ----------------------------------
  n_身高 Number(10, 2); --单位:cm
  n_体重 Number(10, 2); --体重:KG

  l_Clob    Clob;
  v_Err_Msg Varchar2(2000);
  v_Temp    Varchar2(200);
  v_Value   Varchar2(200);
  n_Nodenum Number(5);
  Err_Item Exception;
Begin
  --：
  --将CLOB数据提取到v_XML中
  Select 参数内容 Into l_Clob From 中联合理用药参数;
  Xml_Ret        := Xmltype(l_Clob); --缓存函数返回值
  Xml_Document   := Xmldom.Newdomdocument(Xml_Ret);
  Xml_Domelement := Xmldom.Getdocumentelement(Xml_Document);
  Xml_Nodelist   := Xmldom.Getelementsbytagname(Xml_Domelement, 'patient_info');
  --获取patient_info/INfo节点
  Xml_Nodelist := Xmldom.Getchildnodes(Xmldom.Item(Xml_Nodelist, 0));
  n_Nodenum    := Xmldom.Getlength(Xml_Nodelist);
  For I In 0 .. n_Nodenum - 1 Loop
    Xml_Domnamednodemap := Xmldom.Getattributes(Xmldom.Item(Xml_Nodelist, I));
    v_Temp              := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'name'));
    If v_Temp = '身高' Then
      n_身高 := Nvl(To_Number(Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'))), 0);
    End If;
    If v_Temp = '体重' Then
      n_体重 := Nvl(To_Number(Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'))), 0);
    End If;
  End Loop;
  --获取medicine/INfo节点

  Xml_Nodelist := Xmldom.Getelementsbytagname(Xml_Domelement, 'medicine');
  n_Nodenum    := Xmldom.Getlength(Xml_Nodelist);
  For I In 0 .. n_Nodenum - 1 Loop
    Xml_Node_Med := Xmldom.Item(Xml_Nodelist, I); --取第一个孩子节点medicine
    Xml_Nodelist := Xmldom.Getchildnodes(Xml_Node_Med); --infos
    Xml_Node     := Xmldom.Getfirstchild(Xml_Node_Med); --取第一个孩子节点
    While Not Xmldom.Isnull(Xml_Node) Loop
      Xml_Domnamednodemap := Xmldom.Getattributes(Xml_Node);
      v_Temp              := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'name'));
      v_Value             := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'));
      If v_Temp = '单次量' Then
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '单次量-按体重');
        If n_体重 > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              To_Char(To_Number(v_Value) / n_体重, 'FM9999990.09'));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
        --单次量-按体表trunc(每日量/(0.0061*病人身高+0.0128*病人体重-0.1529),2)
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '单次量-按体表');
        If n_体重 > 0 And n_身高 > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              To_Char(To_Number(v_Value) / (0.0061 * n_身高 + 0.0128 * n_体重 - 0.1529), 'FM9999990.09'));
        
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      End If;
    
      If v_Temp = '每日量' Then
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '每日量-按体重');
        If n_体重 > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              To_Char(To_Number(v_Value) / n_体重, 'FM9999990.09'));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      
        --每日量-按体表
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '每日量-按体表');
        If n_体重 > 0 And n_身高 > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              To_Char(Trunc(To_Number(v_Value) / (0.0061 * n_身高 + 0.0128 * n_体重 - 0.1529), 2),
                                       'FM9999990.09'));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      End If;
      --取下一个兄弟节点
      Xml_Node := Xmldom.Getnextsibling(Xml_Node);
    End Loop;
  End Loop;

  --将函数返回值存入临时表,ZLHIS在事物结束前提取（因为受驱动限制返回值不能超过4000限制）
  Update 中联合理用药参数 Set 参数内容 = Xml_Ret.Getclobval();

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_中联合理用药参数_Update;
/

--123754:冉俊明,2018-04-09,医生工作站预约挂号性能问题
Create Or Replace Procedure Zl1_Auto_Buildingregisterplan
(
  挂号时间_In In Date := Null,
  号源id_In   临床出诊号源.Id%Type := Null
) As
  -------------------------------------------------------------------------
  --功能说明：自动生成临床出诊记录
  --          1、根据号源自动生成预约数内的临床出诊记录;
  --          2、预约天数的确定:号源预约天数-->预约方式的天数（取最大)-->系统预约天数
  --入参:挂号时间_IN:NULL时，自动生成;否则只检查指定日期是否生成了出诊记录没有
  --    号源id_In:NULL时处理所有号源，否则只处理指定号源
  -------------------------------------------------------------------------
  n_缺省预约天数 临床出诊号源.预约天数%Type;
  v_操作员姓名   临床出诊安排.操作员姓名%Type;
  d_登记日期     临床出诊安排.登记时间%Type;
  n_安排id       临床出诊安排.Id%Type;
  n_项目id       临床出诊安排.项目id %Type;

  n_记录id   临床出诊记录.Id%Type;
  d_当前日期 临床出诊记录.出诊日期%Type;

  l_固定时段 t_Strlist := t_Strlist();
  n_Count    Number(18);

  n_加预约天数 Number := 0;
  d_开始时间   临床出诊记录.开始时间%Type;
Begin

  Select Max(预约天数) Into n_缺省预约天数 From 预约方式;
  If Nvl(n_缺省预约天数, 0) = 0 Then
    n_缺省预约天数 := To_Number(Nvl(zl_GetSysParameter('挂号允许预约天数'), '0'));
  End If;
  If Nvl(n_缺省预约天数, 0) = 0 Then
    n_缺省预约天数 := 7;
  End If;

  --以半天为单位,如果参数“号源开放时间”在12:00:00-23:59:59期间的，则开放预约天数+1天
  n_加预约天数 := Zl_Fun_Getappointmentdays;

  d_当前日期   := Trunc(Nvl(挂号时间_In, Sysdate));
  d_登记日期   := Sysdate;
  v_操作员姓名 := Zl_Username;

  --第一层循环，号源信息
  For c_号源 In (Select c.Id, c.号类, c.号码, c.项目id, c.科室id, c.医生姓名,
                      Decode(Nvl(c.预约天数, 0), 0, n_缺省预约天数, c.预约天数) + n_加预约天数 As 预约天数, Nvl(b.站点, '-') As 站点,
                      Nvl(c.是否假日换休, 0) As 是否假日换休, Nvl(c.假日控制状态, 0) As 假日控制状态, Nvl(c.排班方式, 0) As 排班方式
               From 临床出诊号源 C, 部门表 B, 人员表 A, 收费项目目录 D
               Where c.科室id = b.Id And c.医生id = a.Id(+) And c.项目id = d.Id And Nvl(c.是否删除, 0) = 0 And
                     Nvl(c.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(b.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(a.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(d.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (号源id_In Is Null Or c.Id = 号源id_In)
                    --
                     And Exists (Select 1
                      From 临床出诊安排 M, 临床出诊表 N
                      Where m.出诊id = n.Id And m.号源id = c.Id And Nvl(n.排班方式, 0) = 0 And n.发布时间 Is Not Null And
                            m.审核时间 Is Not Null And d_当前日期 <= m.终止时间)) Loop
  
    --检查当前日期所在的安排的收费项目是否为号源中的收费项目，如果不是，则更新号源中的收费项目
    Begin
      Select 项目id
      Into n_项目id
      From (Select a.项目id
             From 临床出诊安排 A, 临床出诊表 B
             Where a.出诊id = b.Id And a.号源id = c_号源.Id And a.审核时间 Is Not Null And d_当前日期 Between a.开始时间 And a.终止时间 And
                   Nvl(b.排班方式, 0) = 0 And b.发布时间 Is Not Null
             Order By a.登记时间 Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        n_项目id := Null;
    End;
    If Nvl(n_项目id, 0) <> 0 Then
      If Nvl(c_号源.项目id, 0) <> n_项目id Then
        Update 临床出诊号源 Set 项目id = n_项目id Where ID = c_号源.Id;
        Commit;
      End If;
    End If;
  
    --第二层循环，出诊日期
    --从头一天开始生成，避免如全日(8:00-7:59)在0:00-7:59没有出诊记录
    --1.未指定号源ID，则是正常生成出诊记录，有出诊记录的日期将不再处理
    --2.指定了号源ID，肯定是发布后新增了临时安排重新生成出诊记录
    For c_日期 In (Select m.日期,
                        Decode(To_Char(m.日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7',
                                '周六', Null) As 星期
                 From (Select Trunc(d_当前日期) + 天数 As 日期
                        From (Select Level - 1 As 天数 From Dual Connect By Level <= c_号源.预约天数 + 1)
                        Where 号源id_In Is Not Null
                        Union All
                        Select Trunc(d_当前日期 - 1) + 天数 As 日期
                        From (Select Level - 1 As 天数 From Dual Connect By Level <= c_号源.预约天数 + 2)
                        Where 号源id_In Is Null And Not Exists
                         (Select 1
                               From 临床出诊记录 A
                               Where a.号源id = c_号源.Id And a.出诊日期 = Trunc(d_当前日期 - 1) + 天数)) M
                 Where 挂号时间_In Is Null Or Trunc(挂号时间_In) = m.日期) Loop
    
      l_固定时段 := t_Strlist();
      --检查当日是否在月/周出诊表中,若在，则不生成出诊记录
      Select Count(1)
      Into n_Count
      From 临床出诊安排 A, 临床出诊表 B
      Where a.出诊id = b.Id And a.号源id = c_号源.Id And c_日期.日期 Between Trunc(a.开始时间) And Trunc(a.终止时间) And
            Nvl(b.排班方式, 0) In (1, 2) And Rownum < 2;
    
      --当前号源为按月/周排班，且当前日期之前已有按月/周排班的出诊记录就不再按固定安排生成出诊记录了
      If n_Count = 0 And Nvl(c_号源.排班方式, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From 临床出诊安排 A, 临床出诊表 B
        Where a.出诊id = b.Id And Nvl(b.排班方式, 0) In (1, 2) And a.号源id = c_号源.Id And a.开始时间 < c_日期.日期 And Rownum < 2;
      End If;
    
      If n_Count = 0 Then
        If 号源id_In Is Null Then
          --出诊安排,取最后登记的一个
          Begin
            Select 安排id
            Into n_安排id
            From (Select a.Id As 安排id
                   From 临床出诊安排 A, 临床出诊表 B
                   Where a.号源id = c_号源.Id And a.出诊id = b.Id And Nvl(b.排班方式, 0) = 0 And b.发布时间 Is Not Null And
                         a.审核时间 Is Not Null And c_日期.日期 Between a.开始时间 And a.终止时间
                   Order By a.登记时间 Desc)
            Where Rownum < 2;
          Exception
            When Others Then
              n_安排id := 0;
          End;
        Else
          --如果指定了号源ID，肯定是发布后新增了临时安排重新生成出诊记录，最后登记的一个肯定是本次新增的，
          --只需要处理这个安排即可，不在这个安排有效时间范围内的就不处理
          Begin
            Select 安排id
            Into n_安排id
            From (Select a.Id As 安排id, a.开始时间, a.终止时间, Row_Number() Over(Order By a.登记时间 Desc) As 行号
                   From 临床出诊安排 A, 临床出诊表 B
                   Where a.号源id = c_号源.Id And a.出诊id = b.Id And Nvl(b.排班方式, 0) = 0 And b.发布时间 Is Not Null And
                         a.审核时间 Is Not Null And c_日期.日期 Between 开始时间 And 终止时间)
            Where 行号 = 1;
          Exception
            When Others Then
              n_安排id := 0;
          End;
        End If;
      
        If Nvl(n_安排id, 0) <> 0 Then
          If 号源id_In Is Not Null Then
            --2.指定了号源ID，肯定是发布后新增了临时安排重新生成出诊记录
            --当日有出诊记录，需要做如下处理
            For c_记录 In (Select a.安排id, a.Id As 记录id, a.出诊日期, a.上班时段, a.是否分时段, a.是否序号控制
                         From 临床出诊记录 A
                         Where a.号源id = c_号源.Id And a.出诊日期 = c_日期.日期) Loop
            
              Select Count(1) Into n_Count From 病人挂号记录 Where 出诊记录id = c_记录.记录id;
              If n_Count = 0 Then
                --2.2.1如果时段不存在预约挂号数据，则删除重新生成
                Zl_临床出诊上班时段_Delete(c_记录.安排id, To_Char(c_记录.出诊日期, 'yyyy-mm-dd'), 1, c_记录.上班时段);
              Else
                --2.2.2如果时段存在预约挂号数据，则只需调整出诊记录的安排ID即可
                Update 临床出诊记录 Set 安排id = n_安排id Where ID = c_记录.记录id;
                l_固定时段.Extend();
                l_固定时段(l_固定时段.Count) := c_记录.上班时段;
              End If;
            End Loop;
          End If;
        
          --检查这天是否出诊
          Select Count(1) Into n_Count From 临床出诊限制 Where 安排id = n_安排id And 限制项目 = c_日期.星期;
          If n_Count = 0 Then
            --如果不存在临床出诊记录，则增加临床出诊记录(时间段为NULL 的空记录)
            Insert Into 临床出诊记录
              (ID, 安排id, 号源id, 出诊日期, 登记人, 登记时间)
              Select 临床出诊记录_Id.Nextval, n_安排id, a.Id As ID, c_日期.日期, v_操作员姓名, d_登记日期 As 登记时间
              From 临床出诊号源 A, 临床出诊安排 B
              Where a.Id = b.号源id And b.Id = n_安排id And Not Exists
               (Select 1 From 临床出诊记录 Where 号源id = a.Id And 出诊日期 = c_日期.日期);
          Else
            For c_记录 In (With c_时间段 As
                            (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间
                            From (Select 时间段, 开始时间, 终止时间, 号类, 站点, 缺省时间, 提前时间,
                                          Row_Number() Over(Partition By 时间段 Order By 时间段, 站点 Asc, 号类 Asc) As 组号
                                   From 时间段
                                   Where Nvl(站点, c_号源.站点) = c_号源.站点 And Nvl(号类, c_号源.号类) = c_号源.号类)
                            Where 组号 = 1)
                           Select n_安排id As 安排id, B1.号源id, c_日期.日期 As 出诊日期, m.上班时段, m.Id As 限制id,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.开始时间, 'hh24:mi:ss'),
                                           'yyyy-mm-dd hh24:mi:ss') As 开始时间,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.终止时间, 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.终止时间 <= j.开始时间 Then
                                     1
                                    Else
                                     0
                                  End As 终止时间, Null As 停诊开始时间, Null As 停诊终止时间, Null As 停诊原因,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.缺省时间, j.开始时间), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.缺省时间 < j.开始时间 Then
                                     1
                                    Else
                                     0
                                  End As 缺省预约时间,
                                  To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(Nvl(j.提前时间, j.开始时间), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.开始时间 < j.提前时间 Then
                                     -1
                                    Else
                                     0
                                  End As 提前挂号时间, m.限号数, 0 As 已挂数, m.限约数, 0 As 已约数, 0 As 其中已接收, m.是否序号控制, m.是否分时段, m.预约控制,
                                  m.是否独占, B1.项目id, B1.医生id, B1.医生姓名, Null As 替诊医生id, Null As 替诊医生姓名, m.分诊方式, m.诊室id,
                                  0 As 是否锁定, 0 As 是否临时出诊, v_操作员姓名 As 操作员姓名, d_登记日期 As 登记时间, c_日期.星期 As 限制项目
                           From 临床出诊安排 B1, 临床出诊限制 M, c_时间段 J
                           Where B1.Id = n_安排id And B1.Id = m.安排id And m.限制项目 = c_日期.星期 And m.上班时段 = j.时间段 And
                                 To_Date(To_Char(c_日期.日期, 'yyyy-mm-dd ') || To_Char(j.开始时间, 'hh24:mi:ss'),
                                         'yyyy-mm-dd hh24:mi:ss') >= B1.开始时间 And Not Exists
                            (Select 1 From Table(l_固定时段) Where Column_Value = m.上班时段)) Loop
            
              Select 临床出诊记录_Id.Nextval Into n_记录id From Dual;
              Insert Into 临床出诊记录
                (ID, 安排id, 号源id, 出诊日期, 上班时段, 开始时间, 终止时间, 停诊开始时间, 停诊终止时间, 停诊原因, 缺省预约时间, 提前挂号时间, 限号数, 已挂数, 限约数, 已约数,
                 其中已接收, 是否序号控制, 是否分时段, 预约控制, 是否独占, 项目id, 科室id, 医生id, 医生姓名, 替诊医生id, 替诊医生姓名, 分诊方式, 诊室id, 是否锁定, 是否临时出诊, 登记人,
                 登记时间, 是否发布)
              Values
                (n_记录id, c_记录.安排id, c_记录.号源id, c_记录.出诊日期, c_记录.上班时段, c_记录.开始时间, c_记录.终止时间, c_记录.停诊开始时间, c_记录.停诊终止时间,
                 c_记录.停诊原因, c_记录.缺省预约时间, c_记录.提前挂号时间, c_记录.限号数, c_记录.已挂数, c_记录.限约数, c_记录.已约数, c_记录.其中已接收, c_记录.是否序号控制,
                 c_记录.是否分时段, c_记录.预约控制, c_记录.是否独占, c_记录.项目id, c_号源.科室id, c_记录.医生id, c_记录.医生姓名, c_记录.替诊医生id, c_记录.替诊医生姓名,
                 c_记录.分诊方式, c_记录.诊室id, c_记录.是否锁定, c_记录.是否临时出诊, c_记录.操作员姓名, d_登记日期, 1);
            
              d_开始时间 := c_记录.开始时间;
              --插入临床出诊序号控制
              If Nvl(c_记录.是否分时段, 0) = 1 And Nvl(c_记录.是否序号控制, 0) = 1 Then
                --分时段且启用序号控制，使用"预约顺序号"记录"是否预约"
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约, 预约顺序号)
                  Select n_记录id, 序号,
                         To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'),
                                  'yyyy-mm-dd hh24:mi:ss') + Case
                            When d_开始时间 > To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                             1
                            Else
                             0
                          End,
                         To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'),
                                  'yyyy-mm-dd hh24:mi:ss') + Case
                            When d_开始时间 >= To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                             1
                            Else
                             0
                          End, 限制数量, 是否预约, 是否预约
                  From 临床出诊时段
                  Where 限制id = c_记录.限制id;
              Else
                Insert Into 临床出诊序号控制
                  (记录id, 序号, 开始时间, 终止时间, 数量, 是否预约)
                  Select n_记录id, 序号,
                         To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'),
                                 'yyyy-mm-dd hh24:mi:ss') + Case
                           When d_开始时间 > To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(开始时间, 'hh24:mi:ss'),
                                                 'yyyy-mm-dd hh24:mi:ss') Then
                            1
                           Else
                            0
                         End,
                         To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'),
                                 'yyyy-mm-dd hh24:mi:ss') + Case
                           When d_开始时间 >= To_Date(To_Char(c_记录.出诊日期, 'yyyy-mm-dd ') || To_Char(终止时间, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') Then
                            1
                           Else
                            0
                         End, 限制数量, 是否预约
                  From 临床出诊时段
                  Where 限制id = c_记录.限制id;
              End If;
            
              --插入合作单位挂号控制记录
              Insert Into 临床出诊挂号控制记录
                (类型, 性质, 名称, 记录id, 序号, 控制方式, 数量)
                Select 类型, 性质, 名称, n_记录id, 序号, 控制方式, 数量
                From 临床出诊挂号控制
                Where 限制id = c_记录.限制id;
            
              --插入临床出诊诊室记录
              Insert Into 临床出诊诊室记录
                (记录id, 诊室id)
                Select n_记录id, 诊室id From 临床出诊诊室 Where 限制id = c_记录.限制id;
            End Loop;
          
            --根据停诊安排和法定节假日调整出诊记录的出诊/预约情况
            Zl_Clinicvisitmodify(c_号源.Id, n_安排id, c_日期.日期, v_操作员姓名, d_登记日期);
          End If;
        End If;
      End If;
      --一天一提交
      Commit;
    End Loop;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Auto_Buildingregisterplan;
/

--123726:冉俊明,2018-03-31,分批卫材费用销帐时报错
Create Or Replace Procedure Zl_病人费用销帐_Audit
(
  Id_In       病人费用销帐.费用id%Type,
  申请时间_In 病人费用销帐.申请时间%Type,
  审核人_In   病人费用销帐.审核人%Type,
  审核时间_In 病人费用销帐.审核时间%Type,
  状态_In     病人费用销帐.状态%Type,
  Int自动退料 Integer := 1,
  申请类别_In 病人费用销帐.申请类别%Type := 1 --对药品和卫材有效,缺省为已执行的药品或卫材 
) As
  n_执行状态       住院费用记录.执行状态%Type;
  n_申请类别       病人费用销帐.申请类别%Type;
  v_收费类别       住院费用记录.收费类别%Type;
  v_No             住院费用记录.No%Type;
  n_实际数量       药品收发记录.实际数量%Type;
  n_数量           病人费用销帐.数量%Type;
  n_收发id         药品收发记录.Id%Type;
  n_医嘱id         住院费用记录.Id%Type;
  v_跟踪在用       材料特性.跟踪在用%Type;
  n_收费细目id     住院费用记录.收费细目id%Type;
  n_审核部门id     病人费用销帐.审核部门id%Type;
  n_执行部门id     住院费用记录.执行部门id%Type;
  n_病人id         住院费用记录.病人id%Type;
  n_主页id         住院费用记录.主页id%Type;
  n_审核标志       病案主页.审核标志%Type;
  n_住院状态       病案主页.状态%Type;
  n_病人审核方式   Number(2);
  n_未入科禁止记账 Number(2);

  n_Cnt     Number(18);
  n_Temp    Number(18);
  v_Err_Msg Varchar2(300);
  Err_Item Exception;
Begin

  n_申请类别 := 0;
  Select a.执行状态, a.收费类别, a.收费细目id, a.执行部门id, a.No, Nvl(b.跟踪在用, 0), a.医嘱序号, 病人id, 主页id
  Into n_执行状态, v_收费类别, n_收费细目id, n_执行部门id, v_No, v_跟踪在用, n_医嘱id, n_病人id, n_主页id
  From 住院费用记录 A, 材料特性 B
  Where a.Id = Id_In And a.收费细目id = b.材料id(+);

  If Nvl(n_主页id, 0) <> 0 Then
  
    n_病人审核方式   := Nvl(zl_GetSysParameter(185), 0);
    n_未入科禁止记账 := Nvl(zl_GetSysParameter(215), 0);
    If n_病人审核方式 = 1 Or n_未入科禁止记账 = 1 Then
      Begin
        Select 审核标志, 状态
        Into n_审核标志, n_住院状态
        From 病案主页
        Where 病人id = Nvl(n_病人id, 0) And 主页id = Nvl(n_主页id, 0);
      Exception
        When Others Then
          n_审核标志 := 0;
          n_住院状态 := 0;
      End;
      If n_未入科禁止记账 = 1 And n_住院状态 = 1 Then
        v_Err_Msg := '病人未入科,禁止对病人相关费用的操作!';
        Raise Err_Item;
      End If;
    
      If n_病人审核方式 = 1 Then
        If Nvl(n_审核标志, 0) = 1 Then
          v_Err_Msg := '该病人目前正在审核费用,不能进行费用相关调整!';
          Raise Err_Item;
        End If;
        If Nvl(n_审核标志, 0) = 2 Then
          v_Err_Msg := '该病人目前已经完成了费用审核,不能进行费用相关调整!';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
  End If;
  If Instr(',5,6,7', ',' || v_收费类别) > 0 Or (v_收费类别 = '4' And Nvl(v_跟踪在用, 0) = 1) Then
    n_申请类别 := 申请类别_In;
  End If;

  Update 病人费用销帐
  Set 审核人 = 审核人_In, 审核时间 = 审核时间_In, 状态 = 状态_In
  Where 费用id = Id_In And 申请类别 = n_申请类别 And 申请时间 = 申请时间_In And 状态 = 0
  Returning 数量, 审核部门id Into n_数量, n_审核部门id;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '销帐审核失败,当前操作的记录可能因为并发操作已经被他人处理,请先刷新信息!';
    Raise Err_Item;
  End If;

  If n_申请类别 = 0 And (Instr(',5,6,7', ',' || v_收费类别) > 0 Or (v_收费类别 = '4' And Nvl(v_跟踪在用, 0) = 1)) Then
    --需要检查未执行的数量必须全部申请,才会通过 
    Select Sum(Nvl(付数, 0) * Nvl(实际数量, 0))
    Into n_实际数量
    From 药品收发记录
    Where 审核日期 Is Null And 费用id = Id_In And Instr(',8,9,10,21,24,25,26,', ',' || 单据 || ',') > 0;
    If Nvl(n_实际数量, 0) < Nvl(n_数量, 0) Then
      Select '在单据号<<' || v_No || '>>中' || Decode(v_收费类别, '4', '卫材', '药品') || '为:' || Chr(13) || 编码 || '-' || 名称 ||
              Chr(13) || '的申请数量(' || LTrim(To_Char(n_数量, '9999999990.99')) || ')大于了待发' || Decode(v_收费类别, '4', '料', '药') ||
              '数量(' || LTrim(To_Char(Nvl(n_实际数量, 0), '9999999990.99')) || '),不允许审核!'
      Into v_Err_Msg
      From 收费项目目录
      Where ID = n_收费细目id;
      Raise Err_Item;
    End If;
  
    If n_医嘱id <> 0 Then
      Select Nvl(Max(d.Id), 0)
      Into n_Cnt
      From 病人医嘱记录 A, 病人医嘱发送 B, 输液配药记录 D
      Where a.Id = n_医嘱id And a.Id = b.医嘱id And b.No = v_No And a.相关id = d.医嘱id And b.发送号 = d.发送号 And b.记录性质 = 2 And
            d.操作时间 = 申请时间_In And d.操作状态 = 9;
    
      If n_Cnt <> 0 Then
        Select Count(1)
        Into n_Temp
        From 输液配药状态
        Where 配药id = n_Cnt And 操作类型 = 10 And 操作时间 = 审核时间_In;
        If n_Temp = 0 Then
          Insert Into 输液配药状态 (配药id, 操作类型, 操作人员, 操作时间) Values (n_Cnt, 10, 审核人_In, 审核时间_In);
        End If;
        Update 输液配药记录 Set 操作人员 = 审核人_In, 操作时间 = 审核时间_In, 操作状态 = 10 Where ID = n_Cnt;
      End If;
    End If;
  End If;

  If n_执行状态 <> 0 Then
    If Instr(',5,6,7,', ',' || v_收费类别 || ',') > 0 And n_申请类别 = 1 Then
      If n_执行部门id <> n_审核部门id Then
        Begin
          Select '[' || 编码 || ']' || 名称 Into v_Err_Msg From 收费项目目录 Where ID = n_收费细目id;
        Exception
          When Others Then
            v_Err_Msg := '';
        End;
        v_Err_Msg := '在销帐审核时,药品为' || v_Err_Msg || ' 的已经被执行科室执行,不能再进行销帐审核,请取消审核!';
        Raise Err_Item;
      End If;
    End If;
  
    If v_收费类别 = '4' Then
      If v_跟踪在用 = 1 Then
        If n_执行部门id <> n_审核部门id And n_申请类别 = 1 And Int自动退料 <> 1 Then
          Begin
            Select '[' || 编码 || ']' || 名称 Into v_Err_Msg From 收费项目目录 Where ID = n_收费细目id;
          Exception
            When Others Then
              v_Err_Msg := '';
          End;
          v_Err_Msg := '在销帐审核时,卫材为' || v_Err_Msg || ' 的已经被执行科室执行,不能再进行销帐审核,请取消审核!';
          Raise Err_Item;
        End If;
      
        If n_申请类别 = 1 And Int自动退料 = 1 Then
          n_收发id := -1;
          --可能来自于多个批次 
          For c_收发记录 In (Select ID, 批号, Nvl(Sum(Nvl(付数, 1) * 实际数量), 0) As 数次
                         From 药品收发记录
                         Where 费用id = Id_In And 单据 In (25, 26) And (记录状态 = 1 Or Mod(记录状态, 3) = 0)
                         Group By ID, 批号) Loop
            n_收发id := c_收发记录.Id;
            If n_数量 = 0 Then
              Exit;
            End If;
          
            If n_数量 > c_收发记录.数次 Then
              n_Temp := c_收发记录.数次;
              n_数量 := n_数量 - c_收发记录.数次;
            Else
              n_Temp := n_数量;
              n_数量 := 0;
            End If;
            Zl_材料收发记录_部门退料(c_收发记录.Id, 审核人_In, 审核时间_In, c_收发记录.批号, Null, Null, n_Temp, 0);
          End Loop;
          If n_收发id = -1 Then
            v_Err_Msg := '在销帐审核时,卫材为' || v_Err_Msg || ' 的未找到相关的药品收发信息,可能是因为中途' || Chr(13) ||
                         '更改了卫材的跟踪属性,不能再进行销帐审核,请取消审核!';
            Raise Err_Item;
          End If;
        End If;
      Else
        --不是跟踪的卫材 
        Update 住院费用记录 Set 执行状态 = 0 Where ID = Id_In;
      End If;
    Elsif Instr(',5,6,7,', ',' || v_收费类别 || ',') = 0 Then
      --可能存在部分消帐,所以先将非药品的处理成部分执行,再在销帐审核过程(ZL_住院记帐记录_Delete)中处理,处理规则如下: 
      --在调用本过程时: 
      --   1.如果是已经执行的,则改为部分执行(执行状态=2);再在销帐过程中处理这部分数据(ZL_住院记帐记录_Delete):即:如果执行状态=2,并且部分销帐的,则改为1(已执行) 
      --      原因是因为非药品类只能存在两种状态.已执行;2-未执行 
      --   2.如果是未执行的,则执行状态还是为0,而在销帐过程中记录状态保持不变 
      Update 住院费用记录 Set 执行状态 = Decode(Nvl(执行状态, 0), 0, 0, 2) Where ID = Id_In; --非药品由于没有取消执行的操作,所以对已执行的要先改状态才能调销帐 
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人费用销帐_Audit;
/

--123942:李业庆,2018-04-04,卫材先判断不跟踪在用，不处理价格
Create Or Replace Procedure Zl_药品收发记录_销售出库
(
  Id_In           In 门诊费用记录.Id%Type,
  药品摘要_In     药品收发记录.摘要%Type := Null,
  频次_In         药品收发记录.频次%Type := Null,
  单量_In         药品收发记录.单量%Type := Null,
  用法_In         药品收发记录.用法%Type := Null,
  煎法_In         药品收发记录.外观%Type := Null,
  期效_In         药品收发记录.扣率%Type := Null,
  计价特性_In     药品收发记录.扣率%Type := Null,
  主页id_In       未发药品记录.主页id%Type := Null,
  备货材料_In     Number := 0,
  备货材料批次_In 药品收发记录.批次%Type := Null,
  领药部门_In     药品收发记录.对方部门id%Type := Null
) Is
  ----------------------------------
  --功能：收费、划价时按照参数设置分解药品并产生相应的收发记录
  --规则：
  --      1、循环游标判断总出库数量与游标中每条记录数量是否充足，如果充足就是总数量，不充足挨个遍历直到数量直到遍历完并退出
  --      2、金额计算方式：定价取收费价目表现价，时价分批取库存表零售价，时价不分批，零售金额/实际数量，并将所有批次的金额累加起来为总出库金额
  --参数：
  --      Id_In：门诊费用记录或者住院费用记录ID
  --      备货材料_In：只有高值卫材才需要传入，非0表示是高值卫材模式
  --      备货材料批次_In：支持高值卫材扫码确定批次出库，所以35.70支持材料非备货材料模式按批次出库；药品不支持这种模式，即药品批次都传空，做兼容性判断，即使传了非空，只要是药品都不管批次
  --      药品摘要_In：可选参数
  --      频次_In；单量_In；用法_In；煎法_In；期效_In；计价特性_In，可选参数，医嘱记录产生
  -----------------------------------
  Cursor c_Stock
  (
    n_Outmode  Number,
    n_库房id   药品收发记录.库房id%Type,
    n_药品id   药品收发记录.药品id%Type,
    n_备货批次 药品收发记录.批次%Type,
    n_类别     Number --0-卫材,1-药品
  ) Is
    Select 库房id, 药品id, Nvl(批次, 0) 批次, 效期, 性质, 可用数量, 实际数量, 实际金额, 实际差价, 上次供应商id, 上次采购价, 上次批号, 上次生产日期, 上次产地, 灭菌效期, 批准文号,
           平均成本价, 零售价, 上次扣率, 商品条码, 内部条码, 原产地
    From 药品库存 A
    Where 药品id = n_药品id And 库房id = n_库房id And 性质 = 1 And Decode(n_类别, 0, Decode(n_备货批次, Null, 0, Nvl(批次, 0)), 0) =
          Decode(n_类别, 0, Decode(n_备货批次, Null, 0, Nvl(n_备货批次, 0)), 0) And
          (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate)) And Nvl(可用数量, 0) > 0
    Order By Decode(n_Outmode, 1, 效期, Null), Nvl(批次, 0);
  r_Stock c_Stock%RowType;

  n_Outmode      Number;
  n_分批         药品规格.药房分批%Type;
  n_时价         收费项目目录.是否变价%Type;
  n_当前数量     药品库存.实际数量%Type;
  n_费用金额小数 Number;
  n_费用单价小数 Number;
  n_流通金额小数 Number;
  n_流通单价小数 Number;
  n_标准单价     收费价目.现价%Type;
  n_当前单价     收费价目.现价%Type;
  n_类别         药品单据性质.类别id%Type;
  n_总金额       Number;
  n_总数量       药品库存.实际数量%Type;
  n_单据         药品单据性质.单据%Type;
  n_跟踪在用     材料特性.跟踪在用%Type;
  n_序号         门诊费用记录.序号%Type;
  v_名称         收费项目目录.名称%Type;
  n_虚拟库房id   部门表.Id%Type;
  n_库房id       部门表.Id%Type;
  n_优先级       身份.优先级%Type;
  n_Count        Number;
  Err_Custom Exception;
  v_Rust     Varchar2(300);
  v_Error    Varchar2(255);
  v_部门名称 部门表.名称%Type;

  v_单据类别   Varchar2(10);
  v_No         药品收发记录.No%Type;
  n_对方部门id 药品收发记录.对方部门id%Type;
  n_收费细目id 药品收发记录.药品id%Type;
  n_总出库数量 药品库存.实际数量%Type;
  n_发药库房id 药品收发记录.库房id%Type;
  n_记录性质   门诊费用记录.记录性质%Type;
  v_收费类别   门诊费用记录.收费类别%Type;
  n_多病人单   住院费用记录.多病人单%Type;
  n_医嘱序号   门诊费用记录.医嘱序号%Type;
  v_姓名       门诊费用记录.姓名%Type;
  n_付数       门诊费用记录.付数%Type;
  v_操作员     门诊费用记录.操作员姓名%Type;
  d_登记时间   门诊费用记录.登记时间%Type;
  n_门诊标志   门诊费用记录.门诊标志%Type;
  n_病人科室id 门诊费用记录.病人科室id%Type;
  n_标识号     门诊费用记录.标识号%Type;
  v_性别       门诊费用记录.性别%Type;
  n_年龄       门诊费用记录.年龄%Type;
  n_病人id     门诊费用记录.病人id%Type;
  v_发药窗口   门诊费用记录.发药窗口%Type;
  n_记录状态   门诊费用记录.记录状态%Type;

  --药品收发记录
  n_收发id   药品收发记录.Id%Type;
  n_扣率     药品收发记录.扣率%Type;
  d_灭菌效期 药品收发记录.灭菌效期%Type;
  d_灭菌日期 药品收发记录.灭菌日期%Type;

  v_其他出库no 药品收发记录.No%Type;
  n_出库序号   药品收发记录.序号%Type;
  n_定价售价   收费价目.现价%Type;
  n_出库检查   Number(1);
Begin
  Begin
    Select 类别, NO, 序号, 对方部门id, 收费细目id, 总出库数量, 发药库房id, 记录性质, 收费类别, 多病人单, 医嘱序号, 姓名, 付数, 划价人, 登记时间, 门诊标志, 病人科室id, 标识号, 性别,
           年龄, 病人id, 发药窗口, 记录状态, 标准单价
    Into v_单据类别, v_No, n_序号, n_对方部门id, n_收费细目id, n_总出库数量, n_发药库房id, n_记录性质, v_收费类别, n_多病人单, n_医嘱序号, v_姓名, n_付数, v_操作员,
         d_登记时间, n_门诊标志, n_病人科室id, n_标识号, v_性别, n_年龄, n_病人id, v_发药窗口, n_记录状态, n_标准单价
    From (Select '门诊' As 类别, NO, 序号, 病人科室id As 对方部门id, 收费细目id, 付数 * 数次 As 总出库数量, 执行部门id As 发药库房id, 记录性质, 收费类别, 0 As 多病人单,
                  医嘱序号, 姓名, 付数, 划价人, 登记时间, 门诊标志, 病人科室id, 标识号, 性别, 年龄, 病人id, 发药窗口, 记录状态, 标准单价
           From 门诊费用记录
           Where ID = Id_In
           Union All
           Select '住院' As 类别, NO, 序号, 病人科室id As 对方部门id, 收费细目id, 付数 * 数次 As 总出库数量, 执行部门id As 发药库房id, 记录性质, 收费类别, 多病人单, 医嘱序号,
                  姓名, 付数, Nvl(划价人, 操作员姓名) As 划价人, 登记时间, 门诊标志, 病人科室id, 标识号, 性别, 年龄, 病人id, 发药窗口, 记录状态, 标准单价
           From 住院费用记录
           Where ID = Id_In);
  Exception
    When Others Then
      v_No         := Null;
      n_对方部门id := 0;
      n_总出库数量 := 0;
  End;

  Zl_药品库存_可用数量异常处理(n_发药库房id, n_收费细目id);

  n_跟踪在用 := 0;
  If v_收费类别 = '4' Then
    --跟踪在用
    Select 跟踪在用 Into n_跟踪在用 From 材料特性 Where 材料id = n_收费细目id;
  End If;

  --药品或跟踪在用卫材才继续下面处理
  If v_收费类别 In ('5', '6', '7') Or (v_收费类别 = '4' And Nvl(n_跟踪在用, 0) = 1) Then
  
    --住院领药部门确认
    If v_单据类别 = '住院' Then
      n_对方部门id := 领药部门_In;
    End If;
  
    --只处理有数量的
    If n_总出库数量 <> 0 Then
      If v_收费类别 = '4' Then
        --卫材分批出库方式 
        Select Zl_To_Number(Nvl(zl_GetSysParameter(156), 0)) Into n_Outmode From Dual;
      Else
        --药品分批出库方式
        Select Zl_To_Number(Nvl(zl_GetSysParameter(150), 0)) Into n_Outmode From Dual;
      End If;
    
      --金额小数位数
      Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
      Into n_费用金额小数, n_费用单价小数
      From Dual;
    
      --取流通业务精度位数
      --类别:1-药品 2-卫材
      --内容：2-零售价 4-金额
      --单位：药品:1-售价 5-金额单位
      If v_收费类别 = '4' Then
        Select 精度 Into n_流通单价小数 From 药品卫材精度 Where 类别 = 2 And 内容 = 2 And 单位 = 1;
        Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 2 And 内容 = 4 And 单位 = 5;
      Else
        Select 精度 Into n_流通单价小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 2 And 单位 = 1;
        Select 精度 Into n_流通金额小数 From 药品卫材精度 Where 类别 = 1 And 内容 = 4 And 单位 = 5;
      End If;
    
      n_总数量 := n_总出库数量;
    
      If v_收费类别 = '4' Then
        --收费类别=4表示是卫材单据
        If v_单据类别 = '门诊' Then
          If n_记录性质 = 1 Then
            n_单据 := 24;
          Else
            n_单据 := 25;
          End If;
        Elsif v_单据类别 = '住院' Then
          If n_多病人单 = 1 Then
            n_单据 := 26;
          Else
            n_单据 := 25;
          End If;
        End If;
      
        Select Nvl(a.在用分批, 0), Nvl(b.是否变价, 0), b.名称, c.现价
        Into n_分批, n_时价, v_名称, n_定价售价
        From 材料特性 A, 收费项目目录 B, 收费价目 C
        Where a.材料id = b.Id And b.Id = n_收费细目id And b.Id = c.收费细目id And Sysdate Between c.执行日期 And c.终止日期;
      
        --备货卫材需要判断是否设置了虚拟库房对照
        If Nvl(备货材料_In, 0) = 1 Then
          Begin
            Select 虚拟库房id Into n_虚拟库房id From 虚拟库房对照 Where 科室id = n_发药库房id And Rownum <= 1;
          Exception
            When Others Then
              n_虚拟库房id := 0;
          End;
          If Nvl(n_虚拟库房id, 0) = 0 Then
            Begin
              Select 名称 Into v_Error From 部门表 Where ID = n_发药库房id;
            Exception
              When Others Then
                v_Error := '';
            End;
            v_Error := '执行部门"' || Nvl(v_Error, '') || '"未设置虚拟部门,请在卫材参数设置中设置.';
            Raise Err_Custom;
          End If;
        End If;
      Else
        --收费类别<>4表示是药品单据，收费类别有"5，6，7"
        If v_单据类别 = '门诊' Then
          If n_记录性质 = 1 Then
            n_单据 := 8;
          Else
            n_单据 := 9;
          End If;
        Elsif v_单据类别 = '住院' Then
          If n_多病人单 = 1 Then
            n_单据 := 10;
          Else
            n_单据 := 9;
          End If;
        End If;
      
        Select Nvl(a.药房分批, 0), Nvl(b.是否变价, 0), b.名称, c.现价
        Into n_分批, n_时价, v_名称, n_定价售价
        From 药品规格 A, 收费项目目录 B, 收费价目 C
        Where a.药品id = b.Id And b.Id = n_收费细目id And b.Id = c.收费细目id And Sysdate Between c.执行日期 And c.终止日期;
      End If;
    
      --可能分批时价药品分解的批次变了
      If n_时价 = 1 Then
        --只有一个批次时,直接取该批次的单价
        --按照最小单位进行格式化
      
        If Nvl(备货材料_In, 0) = 1 And v_收费类别 = '4' Then
          v_Rust := Zl_Fun_Getprice(n_收费细目id, n_虚拟库房id, n_总出库数量, 备货材料_In, 备货材料批次_In);
        Else
          v_Rust := Zl_Fun_Getprice(n_收费细目id, n_发药库房id, n_总出库数量, 备货材料_In, 备货材料批次_In);
        End If;
        n_当前单价 := To_Number(Substr(v_Rust, 1, Instr(v_Rust, '|') - 1));
      
        If Round(n_当前单价, n_费用单价小数) <> Round(n_标准单价, n_费用单价小数) Then
          If n_医嘱序号 Is Null Then
            If v_收费类别 = '4' Then
              v_Error := '第 ' || n_序号 || ' 行的时价卫生材料"' || v_名称 || '"当前计算单价不一致,请重新输入数量计算！';
            Else
              v_Error := '第 ' || n_序号 || ' 行的时价药品"' || v_名称 || '"当前计算单价不一致,请重新输入数量计算！';
            End If;
          Else
            If v_收费类别 = '4' Then
              v_Error := '在处理病人"' || v_姓名 || '"时发现时价卫生材料"' || v_名称 || '"当前计算的单价发生变化。' || Chr(13) || Chr(10) ||
                         '请检查该病人是否同时使用了两笔相同的"' || v_名称 || '"！';
            Else
              v_Error := '在处理病人"' || v_姓名 || '"时发现时价药品"' || v_名称 || '"当前计算的单价发生变化。' || Chr(13) || Chr(10) ||
                         '请检查该病人是否同时使用了两笔相同的"' || v_名称 || '"！';
            End If;
          End If;
          Raise Err_Custom;
        End If;
      End If;
    
      If v_收费类别 In ('5', '6', '7') Or (v_收费类别 = '4' And Nvl(n_跟踪在用, 0) = 1) Then
        If Nvl(备货材料_In, 0) = 1 And v_收费类别 = '4' Then
          n_库房id := n_虚拟库房id;
        Else
          n_库房id := n_发药库房id;
        End If;
      
        Begin
          If v_收费类别 In ('5', '6', '7') Then
            Select 检查方式 Into n_出库检查 From 药品出库检查 Where 库房id = n_库房id;
          Else
            Select 检查方式 Into n_出库检查 From 材料出库检查 Where 库房id = n_库房id;
          End If;
        Exception
          When Others Then
            n_出库检查 := 0;
        End;
      
        If v_收费类别 = '4' Then
          Select 类别id Into n_类别 From 药品单据性质 Where 单据 = n_单据 + 16;
        Else
          Select 类别id Into n_类别 From 药品单据性质 Where 单据 = n_单据;
        End If;
      
        n_总金额 := 0;
        --打开游标
        If v_收费类别 = '4' Then
          Open c_Stock(n_Outmode, n_库房id, n_收费细目id, 备货材料批次_In, 0);
        Else
          Open c_Stock(n_Outmode, n_库房id, n_收费细目id, 备货材料批次_In, 1);
        End If;
        --循环遍历
        While n_总出库数量 <> 0 Loop
          Fetch c_Stock
            Into r_Stock;
          If c_Stock%NotFound Then
            --第一次就没有库存,分批或时价都不允许。
            --分批药品数量分解不完,也就是库存不足。
            If n_分批 = 1 Or n_时价 = 1 Then
              Close c_Stock;
              If n_单据 = 8 Or n_单据 = 24 Then
                If v_收费类别 = '4' Then
                  v_Error := '第 ' || n_序号 || ' 行的分批或时价卫生材料"' || v_名称 || '"没有可用的库存！';
                Else
                  v_Error := '第 ' || n_序号 || ' 行的分批或时价药品"' || v_名称 || '"没有可用的药品库存！';
                End If;
              Else
                --单据=9，10，25，26是记账单提示不一样
                If n_医嘱序号 Is Null Then
                  If v_收费类别 = '4' Then
                    If Nvl(备货材料_In, 0) = 1 And Not (n_分批 = 1 Or n_时价 = 1) Then
                      v_Error := '第 ' || n_序号 || ' 行的卫生材料"' || v_名称 || '"没有足够的材料库存,不能进行备货记帐！';
                    Else
                      v_Error := '第 ' || n_序号 || ' 行的分批或时价卫生材料"' || v_名称 || '"没有足够的材料库存' || Case
                                   When Nvl(备货材料_In, 0) = 0 Then
                                    '！'
                                   Else
                                    ',不能进行备货记帐！'
                                 End;
                    End If;
                  Else
                    v_Error := '第 ' || n_序号 || ' 行的分批或时价药品"' || v_名称 || '"没有足够的库存！';
                  End If;
                Else
                  If v_收费类别 = '4' Then
                    If Nvl(备货材料_In, 0) = 1 And Not (n_分批 = 1 Or n_时价 = 1) Then
                      v_Error := '在处理病人"' || v_姓名 || '"时发现卫生材料"' || v_名称 || '"没有足够的材料库存,不能进行备货记帐！';
                    Else
                      v_Error := '在处理病人"' || v_姓名 || '"时发现分批或时价卫生材料"' || v_名称 || '"没有足够的材料库存' || Case
                                   When Nvl(备货材料_In, 0) = 0 Then
                                    '！'
                                   Else
                                    ',不能进行备货记帐！'
                                 End;
                    End If;
                  Else
                    v_Error := '在处理病人"' || v_姓名 || '"时发现分批或时价药品"' || v_名称 || '"没有足够的库存！';
                  End If;
                End If;
              End If;
              Raise Err_Custom;
            End If;
          Elsif (n_分批 = 1 And Nvl(r_Stock.批次, 0) = 0) Or (n_分批 = 0 And Nvl(r_Stock.批次, 0) <> 0) Then
            Close c_Stock;
            If n_医嘱序号 Is Null Then
              If v_收费类别 = '4' Then
                v_Error := '第 ' || n_序号 || ' 行卫生材料"' || v_名称 || '"的在用分批属性与库存记录不相符,请检查材料数据的正确性！';
              Else
                v_Error := '第 ' || n_序号 || ' 行药品"' || v_名称 || '"的分批属性与库存记录不相符,请检查药品数据的正确性！';
              End If;
            Else
              If v_收费类别 = '4' Then
                v_Error := '在处理病人"' || v_姓名 || '"时发现卫生材料"' || v_名称 || '"的分批属性与库存记录不相符,请检查材料数据的正确性！';
              Else
                v_Error := '在处理病人"' || v_姓名 || '"时发现药品"' || v_名称 || '"的分批属性与库存记录不相符,请检查药品数据的正确性！';
              End If;
            End If;
            Raise Err_Custom;
          End If;
        
          If c_Stock%Found Then
            If Nvl(r_Stock.实际数量, 0) = 0 And (n_总出库数量 > 0 Or n_时价 = 1) And n_出库检查 = 2 Then
              --实际数量为零时，如果严格控制库存，不允许出库
              --实际数量不为零，金额为零，可能是正常的零价格管理。
              --负数的情况相当于入库,这种情况应是允许的；但时价需要计算价格，必须要有实际数量。
              Close c_Stock;
              If n_医嘱序号 Is Null Then
                If v_收费类别 = '4' Then
                  v_Error := '第 ' || n_序号 || ' 行的卫生材料"' || v_名称 || '"当前无库存实际数量，可能存在尚未退料的记录，当前不能出库。';
                Else
                  v_Error := '第 ' || n_序号 || ' 行药品"' || v_名称 || '"当前无库存实际数量，可能存在尚未退药的记录，当前不能出库。';
                End If;
              Else
                If v_收费类别 = '4' Then
                  v_Error := '在处理病人"' || v_姓名 || '"时发现卫生材料"' || v_名称 || '"当前无库存实际数量，可能存在尚未退料的记录，当前不能出库。';
                Else
                  v_Error := '在处理病人"' || v_姓名 || '"时发现药品"' || v_名称 || '"当前无库存实际数量，可能存在尚未退药的记录，当前不能出库。';
                End If;
              End If;
              Raise Err_Custom;
            End If;
          End If;
        
          If n_分批 = 1 Or n_时价 = 1 Then
            --对于不分批的时价只可能分解一次,分解不完上面判断了.它分解是为了计算单价.
            --每次分解取小者,库存不够分解不完在上面判断.
            If n_总出库数量 <= Nvl(r_Stock.可用数量, 0) Then
              n_当前数量 := n_总出库数量;
            Else
              n_当前数量 := Nvl(r_Stock.可用数量, 0);
            End If;
            If n_时价 = 1 Then
              n_当前单价 := Nvl(r_Stock.零售价, Nvl(r_Stock.实际金额 / r_Stock.实际数量, 0));
            Elsif n_分批 = 1 Then
              n_当前单价 := n_定价售价;
            End If;
          Else
            --定价不分批
            --非门诊单据且是高值卫材需要检查库存
            If n_单据 <> 8 Or n_单据 <> 24 Then
              If Nvl(备货材料_In, 0) = 1 And v_收费类别 = '4' Then
                If n_总出库数量 > Nvl(r_Stock.可用数量, 0) Then
                  --不分批, 但又是备货卫材方式出库的,则需要检查当前库存是否充足.
                  v_Error := '第 ' || n_序号 || ' 行的卫生材料"' || v_名称 || '"没有足够的材料库存,不能进行备货记帐！';
                  Raise Err_Custom;
                End If;
              End If;
            End If;
            n_当前数量 := n_总出库数量;
            n_当前单价 := n_定价售价;
          End If;
        
          --药品收发记录
          If c_Stock%Found Then
            --卫材灭菌效期:一次性材料且有效期
            If v_收费类别 = '4' Then
              n_Count := 0;
              Begin
                Select 灭菌效期 Into n_Count From 材料特性 Where Nvl(一次性材料, 0) = 1 And 材料id = n_收费细目id;
              Exception
                When Others Then
                  Null;
              End;
              If Nvl(n_Count, 0) > 0 Then
                d_灭菌效期 := r_Stock.灭菌效期;
                d_灭菌日期 := d_灭菌效期 - n_Count * 30;
              End If;
            End If;
          End If;
        
          Select Nvl(Max(序号), 0) + 1
          Into n_序号
          From 药品收发记录
          Where 单据 = n_单据 And 记录状态 = 1 And NO = v_No;
        
          n_扣率 := Null;
          If 期效_In Is Not Null Or 计价特性_In Is Not Null Then
            n_扣率 := Nvl(期效_In, 0) || Nvl(计价特性_In, 0);
          End If;
        
          --分批药品,如果是只使用了一个批次,则要填写付数
          If n_分批 = 1 And n_当前数量 <> n_总数量 Then
            n_Count := 1;
          Else
            n_Count := 0;
          End If;
        
          Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
          --修改的原单据号存放在摘要中
          Insert Into 药品收发记录
            (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 零售价, 零售金额, 摘要, 填制人,
             填制日期, 费用id, 频次, 发药窗口, 单量, 用法, 外观, 扣率, 灭菌效期, 灭菌日期, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码, 原产地)
          Values
            (n_收发id, 1, n_单据, v_No, n_序号, n_发药库房id, n_对方部门id, n_类别, -1, n_收费细目id, Nvl(r_Stock.批次, 0), r_Stock.上次产地,
             r_Stock.上次批号, r_Stock.效期, Decode(n_Count, 1, 1, n_付数), Decode(n_Count, 1, n_当前数量, n_当前数量 / n_付数),
             Decode(n_Count, 1, n_当前数量, n_当前数量 / n_付数), n_当前单价, Round(n_当前单价 * n_当前数量, n_流通金额小数), 药品摘要_In, v_操作员, d_登记时间,
             Id_In, 频次_In, v_发药窗口, 单量_In, 用法_In, 煎法_In, n_扣率, d_灭菌效期, d_灭菌日期, r_Stock.上次供应商id, r_Stock.上次生产日期,
             r_Stock.批准文号, r_Stock.商品条码, r_Stock.内部条码, r_Stock.原产地);
        
          Zl_未审药品记录_Insert(n_收发id);
        
          --药品库存(普通情况可能没有记录)
          Zl_药品库存_Update(n_收发id, 0, 1);
        
          --产生其他出库单 ，只有高值卫材才需要处理
          If v_收费类别 = '4' And Nvl(备货材料_In, 0) = 1 Then
            Begin
              Select Max(a.No), Max(a.序号)
              Into v_其他出库no, n_出库序号
              From 药品收发记录 A, 住院费用记录 B
              Where a.费用id = b.Id And b.No = v_No And 记录性质 = 2 And b.门诊标志 = n_门诊标志 And
                    Instr(',8,9,10,21,24,25,26,', ',' || a.单据 || ',') > 0;
            Exception
              When Others Then
                v_其他出库no := Null;
            End;
            If v_其他出库no Is Null Then
              v_其他出库no := Nextno(74, n_虚拟库房id, Null, 1);
            End If;
            If v_其他出库no Is Null Then
              v_Error := '在生成卫生材料的其他出库单时,获取相关的出库NO有误,请检查出库单的规则是否有误!';
              Raise Err_Custom;
            End If;
            If Nvl(n_病人科室id, 0) <> 0 Then
              Select 名称 Into v_部门名称 From 部门表 Where ID = n_病人科室id;
            End If;
            v_Error := LPad(' ', 4);
            v_Error := Substr('病人姓名:' || v_姓名 || v_Error || '性别:' || v_性别 || v_Error || '年龄' || n_年龄 || v_Error ||
                              '门诊号:' || Nvl(n_标识号, '') || v_Error || '病人科室:' || v_部门名称, 1, 100);
          
            n_出库序号 := Nvl(n_出库序号, 0) + 1;
            Select 药品收发记录_Id.Nextval Into n_收发id From Dual;
          
            --高值卫材类别id默认19是为了方便统计，因为其他出库可以设置很多类别，所以默认19
            Insert Into 药品收发记录
              (ID, 记录状态, 单据, NO, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 批号, 效期, 付数, 填写数量, 实际数量, 零售价, 零售金额, 摘要,
               填制人, 填制日期, 费用id, 频次, 发药窗口, 单量, 用法, 外观, 扣率, 灭菌效期, 灭菌日期, 供药单位id, 生产日期, 批准文号, 商品条码, 内部条码, 原产地)
            Values
              (n_收发id, 1, 21, v_其他出库no, n_出库序号, n_虚拟库房id, n_对方部门id, 19, -1, n_收费细目id, Nvl(r_Stock.批次, 0), r_Stock.上次产地,
               r_Stock.上次批号, r_Stock.效期, 1, n_当前数量, n_当前数量, n_当前单价, Round(n_当前单价 * n_当前数量, n_流通金额小数), v_Error, v_操作员,
               d_登记时间, Id_In, 频次_In, v_发药窗口, 单量_In, 用法_In, 煎法_In, n_扣率, d_灭菌效期, d_灭菌日期, r_Stock.上次供应商id, r_Stock.上次生产日期,
               r_Stock.批准文号, r_Stock.商品条码, r_Stock.内部条码, r_Stock.原产地);
          
            Zl_未审药品记录_Insert(n_收发id);
          
            --药品库存(普通情况可能没有记录)
            Zl_药品库存_Update(n_收发id, 0, 1);
          End If;
        
          v_Error      := '';
          n_总出库数量 := n_总出库数量 - n_当前数量;
          n_总金额     := n_总金额 + n_当前数量 * n_当前单价;
        End Loop;
      
        --未发药品记录
        Update 未发药品记录
        Set 病人id = n_病人id, 姓名 = v_姓名, 发药窗口 = v_发药窗口, 主页id = 主页id_In
        Where 单据 = n_单据 And NO = v_No And Nvl(库房id, 0) = Nvl(n_发药库房id, 0);
        If Sql%RowCount = 0 Then
          --取身份优先级
          Begin
            Select b.优先级 Into n_优先级 From 病人信息 A, 身份 B Where a.身份 = b.名称(+) And a.病人id = n_病人id;
          Exception
            When Others Then
              Null;
          End;
          Insert Into 未发药品记录
            (单据, NO, 病人id, 主页id, 姓名, 优先级, 库房id, 对方部门id, 填制日期, 已收费, 打印状态, 发药窗口)
          Values
            (n_单据, v_No, n_病人id, 主页id_In, v_姓名, n_优先级, n_发药库房id, n_对方部门id, d_登记时间, n_记录状态, 0, v_发药窗口);
        End If;
      
        --处理未发药记录状态
        Zl_Prescription_Type_Update(v_No, n_记录性质, n_收费细目id, v_收费类别);
      
        Close c_Stock;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_药品收发记录_销售出库;
/


------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0005' Where 编号=&n_System;
Commit;
