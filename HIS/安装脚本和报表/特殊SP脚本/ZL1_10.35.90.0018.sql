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
--127820:余伟节,2018-06-26,合理用药监测提供中联审方接口
Insert Into 三方服务配置目录 (系统标识, 服务名称) Values ('药师处方审查', '审查结果查询');

Insert Into 三方服务配置目录 (系统标识, 服务名称) Values ('药师处方审查', '回写医生拒绝理由');

--127571:蒋廷中,2018-06-25,用于记录电子病案查阅窗口上次选择项
Insert Into zlParameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值, 影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明)
  Select Zlparameters_Id.Nextval, &n_System, 1259, 1, 1, 0, 0, 0, 0, 1, '缺省显示信息', Null, Null,
         '用于记录电子病案查阅窗口上次选择节点项的Key值,便于下次打开电子病案查阅时恢复上次选择的节点', '记录电子病案查阅窗口上次选择节点项的Key值', '', '便于下次打开电子病案查阅时恢复上次选择的节点', Null
  From Dual;

--126645:胡俊勇,2018-06-22,门诊病人预约入院配合修改
Insert Into Zlprocedure(Id, 类型, 名称, 状态, 所有者, 说明) Values (Zlprocedure_Id.Nextval,2,'Zl_Third_Outpatireg',3,User,'用于产生预入院记录/取消预入院。具体入参、出参、返回值说明，详见《vssData/DataStructure/中联三方接口说明(Oracle).xlsx》');

--126645:胡俊勇,2018-06-22,门诊病人预约入院配合修改
Insert Into 三方服务配置目录(系统标识,服务名称) 
Select '预约中心','住院申请' From Dual Union All
Select '预约中心','住院申请取消' From Dual;

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------
--124567:胡俊勇,2018-06-26,门诊医生站号类显示
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限)
Select 100,1260,'基本',User,A.* From (
Select 对象,权限 From zlProgPrivs Where 1 = 0
Union All Select '临床出诊记录','SELECT' From Dual
Union All Select '临床出诊号源','SELECT' From Dual) A;


-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--127819:余伟节,2018-06-27,中联合理用药处方上传
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
  Xml_Node_Pati       Xmldom.Domnode;
  Xml_Node            Xmldom.Domnode;
  Xml_Node_New        Xmldom.Domnode;
  ----------------------------------
  n_身高 Number(10, 2); --单位:cm
  n_体重 Number(10, 2); --体重:KG
  v_Type Varchar2(200);

  l_Clob    Clob;
  v_Err_Msg Varchar2(2000);
  v_Temp    Varchar2(200);
  v_Value   Varchar2(200);
  n_Nodenum Number(5);
  Err_Item Exception;

  Procedure Addpatiinfo
  (
    Nodeparent Xmldom.Domnode,
    Nodecopy   Xmldom.Domnode,
    Nodename   Varchar2,
    Nodevalue  Varchar2
  ) Is
    Nodenew Xmldom.Domnode;
  Begin
    Nodenew := Xmldom.Appendchild(Nodeparent, Xmldom.Clonenode(Nodecopy, False));
    Xmldom.Setattribute(Xmldom.Makeelement(Nodenew), 'name', Nodename);
    Xmldom.Setattribute(Xmldom.Makeelement(Nodenew), 'value', Nodevalue);
  End;
Begin

  --：
  --将CLOB数据提取到v_XML中
  Select 参数内容 Into l_Clob From 中联合理用药参数;
  Xml_Ret        := Xmltype(l_Clob); --缓存函数返回值
  Xml_Document   := Xmldom.Newdomdocument(Xml_Ret);
  Xml_Domelement := Xmldom.Getdocumentelement(Xml_Document);
  Xml_Nodelist   := Xmldom.Getelementsbytagname(Xml_Domelement, 'patient_info');
  Xml_Node_Pati  := Xmldom.Item(Xml_Nodelist, 0);
  --获取patient_info/INfo节点
  Xml_Nodelist := Xmldom.Getchildnodes(Xml_Node_Pati);
  n_Nodenum    := Xmldom.Getlength(Xml_Nodelist);
  For I In 0 .. n_Nodenum - 1 Loop
    Xml_Node            := Xmldom.Item(Xml_Nodelist, I);
    Xml_Domnamednodemap := Xmldom.Getattributes(Xml_Node);
    v_Temp              := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'name'));
    If v_Temp = '提交类型' Then
      v_Type := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'));
      If v_Type = '2' Then
        --1-新开未保存;2-保存医嘱后
        Addpatiinfo(Xml_Node_Pati, Xml_Node, '病人ID', 病人id_In);
        If Nvl(挂号id_In, 0) = 0 Then
          Addpatiinfo(Xml_Node_Pati, Xml_Node, '就诊ID', 主页id_In);
        Else
          Addpatiinfo(Xml_Node_Pati, Xml_Node, '就诊ID', 挂号id_In);
        End If;
      End If;
    End If;
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

--126645:胡俊勇,2018-06-22,门诊病人预约入院配合修改
Create Or Replace Procedure Zl_门诊医嘱发送_Insert
(
  医嘱id_In     In 病人医嘱发送.医嘱id%Type,
  发送号_In     In 病人医嘱发送.发送号%Type,
  记录性质_In   In 病人医嘱发送.记录性质%Type,
  No_In         In 病人医嘱发送.No%Type,
  记录序号_In   In 病人医嘱发送.记录序号%Type,
  发送数次_In   In 病人医嘱发送.发送数次%Type,
  首次时间_In   In 病人医嘱发送.首次时间%Type,
  末次时间_In   In 病人医嘱发送.末次时间%Type,
  发送时间_In   In 病人医嘱发送.发送时间%Type,
  执行状态_In   In 病人医嘱发送.执行状态%Type,
  执行部门id_In In 病人医嘱发送.执行部门id%Type,
  计费状态_In   In 病人医嘱发送.计费状态%Type,
  First_In      In Number := 0,
  样本条码_In   In 病人医嘱发送.样本条码%Type := Null,
  操作员编号_In In 人员表.编号%Type := Null,
  操作员姓名_In In 人员表.姓名%Type := Null,
  原液皮试_In   In Varchar2 := Null,
  预约中心_In   In Number := 0
  --功能：填写病人医嘱发送记录
  --参数：First_IN=表示是否一组医嘱的第一医嘱行,以便处理医嘱相关内容(如成药,配方的第一行,因为给药途径,配方煎法,用法可能为叮嘱不发送)
  --      源液皮试_In 原液皮试医嘱ID，需求号7107/bug115972用于关联药品医嘱行和皮试医嘱行。关联字段为 病人医嘱发送.标本发送批号 存入药品行的医嘱ID值
  --      预约中心_in 是否启用入院预约中心，由程序外部传入
  --      格式：1医嘱ID,2医嘱ID 前面一个为皮试医嘱的医嘱ID，第二个为药品行医嘱的医嘱ID  
) Is
  --包含病人及医嘱(一组医嘱中第一行)相关信息的游标
  Cursor c_Advice Is
    Select Nvl(a.相关id, a.Id) As 组id, a.相关id, a.序号, a.病人id, a.挂号单, a.婴儿, b.姓名, c.操作类型, a.诊疗类别, a.医嘱状态, a.医嘱内容, a.开嘱医生,
           a.开始执行时间, a.执行时间方案, a.频率次数, a.频率间隔, a.间隔单位, Nvl(a.紧急标志, 0) As 紧急标志, a.诊疗项目id, a.收费细目id
    From 病人医嘱记录 A, 病人信息 B, 诊疗项目目录 C
    Where a.病人id = b.病人id And a.诊疗项目id = c.Id And a.Id = 医嘱id_In
    Group By a.相关id, a.Id, a.序号, a.病人id, a.挂号单, a.婴儿, b.姓名, c.操作类型, a.诊疗类别, a.医嘱状态, a.医嘱内容, a.开嘱医生, a.开始执行时间, a.执行时间方案,
             a.频率次数, a.频率间隔, a.间隔单位, a.紧急标志, a.诊疗项目id, a.收费细目id;
  r_Advice c_Advice%RowType;

  Cursor c_Pati(v_病人id 病人信息.病人id%Type) Is
    Select * From 病人信息 Where 病人id = v_病人id;
  r_Pati c_Pati%RowType;

  --其它临时变量
  v_Temp       Varchar2(255);
  v_Count      Number;
  v_病人性质   病案主页.病人性质%Type;
  v_人员编号   人员表.编号%Type;
  v_人员姓名   人员表.姓名%Type;
  v_入院方式   入院方式.名称%Type;
  n_挂号id     病人挂号记录.Id%Type;
  d_开始时间   病人医嘱记录.开始执行时间%Type;
  n_医嘱状态   病人医嘱记录.医嘱状态%Type;
  n_皮试标号   病人医嘱发送.医嘱id%Type;
  n_皮试医嘱id 病人医嘱发送.医嘱id%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;
Begin
  --当前操作人员
  If 操作员编号_In Is Not Null And 操作员姓名_In Is Not Null Then
    v_人员编号 := 操作员编号_In;
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;
  --如果首次时间为空则填入开始执行时间
  Select 开始执行时间, 医嘱状态 Into d_开始时间, n_医嘱状态 From 病人医嘱记录 Where ID = 医嘱id_In;

  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;
  --是一组医嘱的第一行时处理医嘱内容
  If Nvl(First_In, 0) = 1 Or n_医嘱状态 = 1 Then
    --并发操作检查
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.医嘱状态, 0) <> 1 Then
      v_Error := '"' || r_Advice.姓名 || '"的医嘱"' || r_Advice.医嘱内容 || '"已经被其他人发送。' || Chr(13) || Chr(10) ||
                 '该病人的医嘱发送失败。请重新读取发送清单再试。';
      Raise Err_Custom;
    End If;
  
    --发送后的医嘱处理:临嘱发送后自动停止
    ---------------------------------------------------------------------------------------
    Update 病人医嘱记录
    Set 医嘱状态 = 8, 执行终止时间 = 末次时间_In,
        --可能没有
        停嘱时间 = 发送时间_In,
        --要作为发送时间显示
        停嘱医生 = v_人员姓名 --要作为发送人显示,不同于住院,门诊医嘱无护士操作
    Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
  
    Insert Into 病人医嘱状态
      (医嘱id, 操作类型, 操作人员, 操作时间)
      Select ID, 8, v_人员姓名, 发送时间_In From 病人医嘱记录 Where ID = r_Advice.组id Or 相关id = r_Advice.组id;
  
    --特殊医嘱的处理
    ---------------------------------------------------------------------------------------
    If r_Advice.诊疗类别 = 'Z' And Nvl(r_Advice.操作类型, '0') <> '0' And Nvl(r_Advice.婴儿, 0) = 0 Then
      --1-留观;2-住院;    
      --住院医嘱受预约中心影响先进行条件判断
      If r_Advice.操作类型 = '1' And 执行部门id_In Is Not Null Then
        v_Count := 1;
      Elsif r_Advice.操作类型 = '2' And 执行部门id_In Is Not Null Then
        v_Count := 1;
        If 预约中心_In = 1 Then
          v_Count := 0;
        End If;
      Else
        v_Count := 0;
      End If;
    
      If v_Count = 1 Then
        --满足产生新的预约登记的条件：1.当前无预约,2.当前不在院,3-无要求预约时间内的住院记录
      
        --删除超过挂号有效天数的预约登记
        Begin
          Select Count(*) Into v_Count From 病案主页 Where 病人id = r_Advice.病人id And Nvl(主页id, 0) = 0;
        Exception
          When Others Then
            v_Count := 0;
        End;
        If Nvl(v_Count, 0) > 0 Then
          Zl_入院病案主页_Delete(r_Advice.病人id, 0, 0, 0);
          v_Count := 0;
        End If;
      
        If v_Count = 0 Then
          Select Count(*) Into v_Count From 病案主页 Where 病人id = r_Advice.病人id And 出院日期 Is Null;
        End If;
        If v_Count = 0 Then
          Select Count(*)
          Into v_Count
          From 病案主页
          Where 病人id = r_Advice.病人id And (入院日期 >= r_Advice.开始执行时间 Or 出院日期 >= r_Advice.开始执行时间);
        End If;
        If v_Count = 0 Then
          If r_Advice.操作类型 = '1' Then
            --留观医嘱,将病人在"开始时间"留观到临床执行科室
            Begin
              v_病人性质 := 2;
              Select Decode(服务对象, 1, 1, 2)
              Into v_病人性质
              From 部门性质说明
              Where 工作性质 = '临床' And 部门id = 执行部门id_In;
            Exception
              When Others Then
                Null;
            End;
          Elsif r_Advice.操作类型 = '2' Then
            --住院医嘱,将病人在"开始时间"登记到临床执行科室
            v_病人性质 := 0;
          End If;
        
          Open c_Pati(r_Advice.病人id);
          Fetch c_Pati
            Into r_Pati;
        
          v_入院方式 := Null;
          If r_Advice.紧急标志 = 1 Then
            v_入院方式 := '急诊';
            Select Max(ID)
            Into n_挂号id
            From 病人挂号记录
            Where NO = r_Advice.挂号单 And 记录性质 = 1 And 记录状态 = 1;
          Else
            Select Decode(急诊, 1, '急诊', Null), ID
            Into v_入院方式, n_挂号id
            From 病人挂号记录
            Where NO = r_Advice.挂号单 And 记录性质 = 1 And 记录状态 = 1;
          End If;
          If v_病人性质 = 1 Then
            Zl_入院病案主页_Insert(1, v_病人性质, r_Pati.病人id, r_Pati.门诊号, Null, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.费别,
                             r_Pati.出生日期, r_Pati.国籍, r_Pati.民族, r_Pati.学历, r_Pati.婚姻状况, r_Pati.职业, r_Pati.身份,
                             r_Pati.身份证号, r_Pati.出生地点, r_Pati.家庭地址, r_Pati.家庭地址邮编, r_Pati.家庭电话, r_Pati.户口地址,
                             r_Pati.户口地址邮编, r_Pati.联系人姓名, r_Pati.联系人关系, r_Pati.联系人地址, r_Pati.联系人电话, r_Pati.工作单位,
                             r_Pati.合同单位id, r_Pati.单位电话, r_Pati.单位邮编, r_Pati.单位开户行, r_Pati.单位帐号, r_Pati.担保人, r_Pati.担保额,
                             r_Pati.担保性质, 执行部门id_In, Null, Null, v_入院方式, Null, Null, r_Advice.开嘱医生, r_Pati.籍贯, r_Pati.区域,
                             r_Advice.开始执行时间, Null, Null, r_Pati.医疗付款方式, Null, Null, Null, Null, Null, Null, r_Pati.险类,
                             v_人员编号, v_人员姓名, 0, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, n_挂号id);
          Else
            Zl_入院病案主页_Insert(1, v_病人性质, r_Pati.病人id, r_Pati.住院号, Null, r_Pati.姓名, r_Pati.性别, r_Pati.年龄, r_Pati.费别,
                             r_Pati.出生日期, r_Pati.国籍, r_Pati.民族, r_Pati.学历, r_Pati.婚姻状况, r_Pati.职业, r_Pati.身份,
                             r_Pati.身份证号, r_Pati.出生地点, r_Pati.家庭地址, r_Pati.家庭地址邮编, r_Pati.家庭电话, r_Pati.户口地址,
                             r_Pati.户口地址邮编, r_Pati.联系人姓名, r_Pati.联系人关系, r_Pati.联系人地址, r_Pati.联系人电话, r_Pati.工作单位,
                             r_Pati.合同单位id, r_Pati.单位电话, r_Pati.单位邮编, r_Pati.单位开户行, r_Pati.单位帐号, r_Pati.担保人, r_Pati.担保额,
                             r_Pati.担保性质, 执行部门id_In, Null, Null, v_入院方式, Null, Null, r_Advice.开嘱医生, r_Pati.籍贯, r_Pati.区域,
                             r_Advice.开始执行时间, Null, Null, r_Pati.医疗付款方式, Null, Null, Null, Null, Null, Null, r_Pati.险类,
                             v_人员编号, v_人员姓名, 0, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, n_挂号id);
          End If;
          Close c_Pati;
        End If;
      End If;
    End If;
  End If;
  Close c_Advice;

  If 原液皮试_In Is Not Null Then
    v_Count      := Instr(原液皮试_In, ',');
    n_皮试医嘱id := Substr(原液皮试_In, 1, v_Count - 1);
    n_皮试标号   := Substr(原液皮试_In, v_Count + 1);
    Update 病人医嘱发送 Set 标本发送批号 = n_皮试标号 Where 医嘱id = n_皮试医嘱id;
  End If;
  --填写发送记录
  ---------------------------------------------------------------------------------------
  Insert Into 病人医嘱发送
    (医嘱id, 发送号, 记录性质, NO, 记录序号, 发送数次, 发送人, 发送时间, 执行状态, 执行部门id, 计费状态, 首次时间, 末次时间, 样本条码, 门诊记帐, 标本发送批号)
  Values
    (医嘱id_In, 发送号_In, 记录性质_In, No_In, 记录序号_In, 发送数次_In, v_人员姓名, 发送时间_In, 执行状态_In, 执行部门id_In, 计费状态_In,
     Nvl(首次时间_In, d_开始时间), Nvl(末次时间_In, d_开始时间), 样本条码_In, Decode(记录性质_In, 2, 1, Null), n_皮试标号);

  --手术和检查医嘱同步更新主医嘱的计费状态
  If 计费状态_In = 1 And r_Advice.组id <> 医嘱id_In And (r_Advice.诊疗类别 = 'D' Or r_Advice.诊疗类别 = 'F') Then
    Update 病人医嘱发送 Set 计费状态 = 1 Where 医嘱id = r_Advice.组id And 发送号 = 发送号_In;
  End If;

  --自动填为已执行时，需要同步处理费用执行状态及审核划价状态
  If 执行状态_In = 1 Then
    Zl_病人医嘱执行_Finish(医嘱id_In, 发送号_In, Null, Null, v_人员编号, v_人员姓名, 执行部门id_In);
  End If;
  --消息推送
  Begin
    Execute Immediate 'Begin zl_服务窗消息_发送(:1,:2); End;'
      Using 3, 发送号_In;
  Exception
    When Others Then
      Null;
  End;

  If r_Advice.诊疗类别 = 'E' And r_Advice.操作类型 = '6' Then
    --检验项目
    b_Message.Zlhis_Cis_016(r_Advice.病人id, Null, r_Advice.挂号单, 发送号_In, r_Advice.组id, 1);
  Elsif r_Advice.诊疗类别 = 'D' And r_Advice.相关id Is Null Then
    b_Message.Zlhis_Cis_017(r_Advice.病人id, Null, r_Advice.挂号单, 发送号_In, r_Advice.组id, 1);
  Elsif r_Advice.诊疗类别 = 'F' And r_Advice.相关id Is Null Then
    b_Message.Zlhis_Cis_018(r_Advice.病人id, Null, r_Advice.挂号单, 发送号_In, r_Advice.组id);
  Elsif r_Advice.诊疗类别 = 'K' Then
    b_Message.Zlhis_Cis_019(r_Advice.病人id, Null, r_Advice.挂号单, 发送号_In, r_Advice.组id);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_门诊医嘱发送_Insert;
/

--126645:胡俊勇,2018-06-28,门诊病人预约入院配合修改

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
    Where 病人id = r_Pati.病人id;
  
    --将床位进行占用
    Update 床位状况记录
    Set 状态 = '占用', 病人id = r_Pati.病人id, 科室id = Decode(共用, 1, n_科室id, 科室id)
    Where 病区id = n_病区id And 床号 = v_床号;
  Else
    --取消登记
  
    Select b.病人id, b.入院科室id, b.入院病床, b.入院病区id
    Into n_病人id, n_科室id, v_床号, n_病区id
    From 病案主页 B
    Where b.挂号id = n_挂号id;
  
    --更新病区和床号
    Update 病案主页
    Set 入院病床 = Null, 出院病床 = Null, 入院病区id = Null, 当前病区id = Null
    Where 病人id = r_Pati.病人id;
  
    --将床位进行取消占用
    Update 床位状况记录
    Set 状态 = '空床', 病人id = Null, 科室id = Decode(共用, 1, Null, 科室id)
    Where 病区id = n_病区id And 床号 = v_床号;
  
    Zl_入院病案主页_Delete(n_病人id, 0);
  End If;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Outpatireg;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0018' Where 编号=&n_System;
Commit;
