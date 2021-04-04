----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.35.90升级到 v10.35.90
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--125796:刘涛,2018-08-22,增加冲销记录字段
alter table 药品收发记录 add 冲销原因 varchar2(200);

--128856:董露露,2018-08-23,解决病案编目时希望病人多次住院案号均和第一次病案号一致的问题
Alter Table 住院病案记录 Drop Constraint 住院病案记录_UQ_档案号 Cascade Drop Index;
Create Index 住院病案记录_IX_档案号 On 住院病案记录(档案号) PCTFREE 5 Tablespace zl9IndexMdr;
alter index 住院病案记录_IX_病人ID rename to 住院病案记录_IX_档案号;
------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--130541:胡俊勇,2018-08-23,集成平台消息问题
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note) 
Select '临床', 'ZLHIS_CIS_059', '确认停止患者医嘱', '<root><病人ID></病人ID><主页ID></主页ID><ID></ID></root>', '住院护士工作站:确认停止患者医嘱时'  From Dual;

--130471:余伟节,2018-08-21,预约安排查询
Insert Into 三方服务配置目录 (系统标识, 服务名称) Values ('预约中心', '预约安排查询');

--130469:陈龙,2018-08-21,增加消息
--ZLMSG_LISTS
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select 'BLOOD', 'ZLHIS_BLOOD_003', '配血审核完成', '<root><医嘱ID></医嘱ID><收发ID></收发ID></root>', '科室配血管理:单袋血完成配血时'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_004', '取消配血审核', '<root><医嘱ID></医嘱ID><收发ID></收发ID><审核人></审核人><审核时间></审核时间></root>', '科室配血管理:单袋血取消配血时'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_005', '完成血液发放', '<root><医嘱ID></医嘱ID><收发ID></收发ID></root>', '科室发血管理:完成取血申请的血液发放时'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_006', '取消血液发放', '<root><医嘱ID></医嘱ID><收发ID></收发ID><原收发ID></原收发ID></root>', '科室发血管理:取消已发血液时'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_007', '输血执行登记', '<root><医嘱ID></医嘱ID><收发ID></收发ID></root>', '医技工作站:输血医嘱执行登记时'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_008', '输血执行登记删除', '<root><医嘱ID></医嘱ID><收发ID></收发ID><执行人></执行人><执行时间></执行时间><核对人></核对人><核对时间></核对时间><复查人></复查人></root>', '医技工作站;住院护士工作站:取消输血医嘱执行登记时'  From Dual Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_009', '输血医嘱接收状态', '<root><医嘱ID></医嘱ID><操作类型></操作类型></root>', '科室配血管理:接收等待配血的医嘱时;将正在配血的申请转为为等待配血时'  From Dual  Union All 
Select 'BLOOD', 'ZLHIS_BLOOD_010', '输血医嘱审核完成', '<root><医嘱ID></医嘱ID></root>', '输血审核管理:输血医嘱审核或拒绝审核时'  From Dual;



-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--130541:胡俊勇,2018-08-23,集成平台消息问题
Create Or Replace Procedure Zl_病人医嘱记录_回退
(
  医嘱id_In     In 病人医嘱记录.Id%Type,
  Flag_In       In Number := 0,
  医嘱内容_In   In 病人医嘱记录.医嘱内容%Type := Null,
  操作类型_In   In 病人医嘱状态.操作类型%Type := Null,
  操作员编号_In In 人员表.编号%Type := Null,
  操作员姓名_In In 人员表.姓名%Type := Null
  --功能：回退住院医嘱的状态操作或发送操作(回退重整操作通过调用Zl_病人医嘱记录_批量回退来进行) 
  --参数：医嘱ID_IN=一组医嘱ID 
  --      FLAG_IN=附加数据。回退停止：0=清除执行终止时间,1=保留现有的执行终止时间。 
  --      医嘱内容_IN=该过程被批量回退调用时才用，用于错误提示。 
  --      操作类型_IN=该过程被批量回退调用时才用，用于核对回退数据。0-回退发送,n=回退具体医嘱操作 
) Is
  --包含指定医嘱的操作记录,第一条为要回退的内容(状态操作优先) 
  --临嘱不回退发送后的自动停止,在回退发送时自动回退停止操作 
  Cursor c_Rolladvice Is
    Select b.操作人员, b.操作时间, 0 As 发送号, Null As NO, b.操作类型, 0 As 执行状态, Sysdate + Null As 首次时间, Sysdate + Null As 末次时间,
           a.上次执行时间, a.医嘱期效, a.诊疗类别 As 类别, a.诊疗项目id, Null As 类型, a.病人id, a.主页id, a.婴儿, 0 As 记录性质, 0 As 门诊记帐, 0 As 开嘱科室id,
           a.审核标记, a.开嘱医生, a.执行科室id, Nvl(a.相关id, a.Id) As 组id, a.相关id, a.Id As 医嘱id, -null As 发送数次, Null As 样本条码
    From 病人医嘱记录 A, 病人医嘱状态 B
    Where a.Id = b.医嘱id And (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) And
          (Nvl(a.医嘱期效, 0) = 0 And b.操作类型 Not In (1, 2, 3) Or Nvl(a.医嘱期效, 0) = 1 And b.操作类型 Not In (1, 2, 3, 8))
    Union
    Select b.发送人 As 操作人员, b.发送时间 As 操作时间, b.发送号, b.No, -null As 操作类型, b.执行状态, b.首次时间, b.末次时间, a.上次执行时间, a.医嘱期效, c.类别,
           a.诊疗项目id, c.操作类型 As 类型, a.病人id, a.主页id, a.婴儿, b.记录性质, b.门诊记帐, a.开嘱科室id, a.审核标记, a.开嘱医生, a.执行科室id,
           Nvl(a.相关id, a.Id) As 组id, a.相关id, a.Id As 医嘱id, b.发送数次, b.样本条码
    From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C
    Where a.Id = b.医嘱id And a.诊疗项目id = c.Id And (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In)
    Order By 操作时间 Desc, 发送号;
  r_Rolladvice c_Rolladvice%RowType;

  --方式同c_Rolladvice，只取发送部份用了自动回退处理 
  Cursor c_Rollsend(v_发送号 病人医嘱发送.发送号%Type) Is
    Select Distinct b.医嘱id, b.发送时间 As 操作时间, b.发送号, b.执行状态, a.诊疗类别 As 类别, c.当前病区id As 病人病区id, a.病人科室id,
                    b.执行部门id As 执行科室id
    From 病人医嘱记录 A, 病人医嘱发送 B, 病案主页 C
    Where a.Id = b.医嘱id And b.发送号 = v_发送号 And a.病人id = c.病人id And a.主页id = c.主页id And
          (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In)
    Order By b.发送时间 Desc, b.发送号;

  --根据医嘱及发送NO求出本次回退要销帐的费用记录 
  --一组医嘱并不是都填写了发送记录,且可能NO不同(药品有,用法煎法不一定有) 
  --不管发送记录的计费状态(可能无需计费),有费用记录自然关联出来 
  --费用只求价格父号为空的,以便取序号销帐 
  --只管记录状态为1的费用,对于已销帐或部份销帐的记录,不再处理；其中"记录状态=3"的读取出来仅用于判断，不处理。 
  Cursor c_Rollmoneyout
  (
    v_发送号    病人医嘱发送.发送号%Type,
    v_医嘱id    病人医嘱记录.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.记录状态, a.No, a.序号, a.收费类别, a.执行状态, d.跟踪在用, a.执行部门id, a.记录性质
    From 门诊费用记录 A, Table(t_Adviceids) B, 病人医嘱发送 C, 材料特性 D
    Where c.医嘱id = b.Column_Value And c.发送号 = v_发送号 And a.医嘱序号 = b.Column_Value And
          (a.医嘱序号 = v_医嘱id Or Nvl(v_医嘱id, 0) = 0) And a.记录状态 In (0, 1, 3) And a.No = c.No And a.记录性质 = c.记录性质 And
          a.价格父号 Is Null And a.收费细目id = d.材料id(+)
    Order By a.No, a.序号;

  Cursor c_Rollmoneyin
  (
    v_发送号    病人医嘱发送.发送号%Type,
    v_医嘱id    病人医嘱记录.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.记录状态, a.No, a.序号, a.收费类别, a.执行状态, d.跟踪在用, a.执行部门id, a.记录性质
    From 住院费用记录 A, Table(t_Adviceids) B, 病人医嘱发送 C, 材料特性 D
    Where c.医嘱id = b.Column_Value And c.发送号 = v_发送号 And a.医嘱序号 = b.Column_Value And
          (a.医嘱序号 = v_医嘱id Or Nvl(v_医嘱id, 0) = 0) And a.记录状态 In (0, 1, 3) And a.No = c.No And a.记录性质 = c.记录性质 And
          a.价格父号 Is Null And a.收费细目id = d.材料id(+)
    Order By a.No, a.序号;

  --取发送住院记帐时自动发放的卫材(还没有退料的) 
  Cursor c_Stuff_Drug(v_费用id 药品收发记录.费用id%Type) Is
    Select ID
    From 药品收发记录
    Where 费用id = v_费用id And (记录状态 = 1 Or Mod(记录状态, 3) = 0) And 审核人 Is Not Null
    Order By 药品id;

  --用于处理特殊医嘱的回退 
  Cursor c_Patilog
  (
    v_病人id 病人变动记录.病人id%Type,
    v_主页id 病人变动记录.主页id%Type
  ) Is
    Select *
    From 病人变动记录
    Where 病人id = v_病人id And 主页id = v_主页id And 终止时间 Is Null
    Order By 开始时间 Desc;
  r_Patilog c_Patilog%RowType;

  Cursor c_Adviceids Is
    Select ID From 病人医嘱记录 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
  t_Adviceids t_Numlist;

  v_医嘱状态     病人医嘱记录.医嘱状态%Type;
  v_医嘱期效     病人医嘱记录.医嘱期效%Type;
  v_费用no       病人医嘱发送.No%Type;
  v_费用序号     Varchar2(255);
  v_末次时间     病人医嘱发送.末次时间%Type;
  v_重整时间     病人医嘱状态.操作时间%Type;
  v_操作类型     诊疗项目目录.操作类型%Type;
  v_执行频率     诊疗项目目录.执行频率%Type;
  v_上次时间     病人医嘱记录.上次执行时间%Type;
  v_执行时间     病人医嘱记录.执行时间方案%Type;
  v_开始执行时间 病人医嘱记录.开始执行时间%Type;
  v_上次打印时间 病人医嘱记录.上次打印时间%Type;
  v_频率间隔     病人医嘱记录.频率间隔%Type;
  v_间隔单位     病人医嘱记录.间隔单位%Type;
  v_发送号       病人医嘱发送.发送号%Type;
  n_护理等级id   病人变动记录.护理等级id%Type;
  d_开始时间     病人变动记录.开始时间%Type;
  d_操作时间     病人医嘱状态.操作时间%Type;
  v_Tmp发送号    病人医嘱发送.发送号%Type;
  n_执行         Number;

  Intdigit   Number(3);
  v_Update   Number(1);
  v_Count    Number(5);
  v_Temp     Varchar2(2000);
  v_人员编号 人员表.编号%Type;
  v_人员姓名 人员表.姓名%Type;
  v_Time     Varchar2(4000);
  n_Blndo    Number;

  v_Error Varchar2(2000);
  Err_Custom Exception;

  Function Checkmoneyundo
  (
    v_No       住院费用记录.No%Type,
    v_记录性质 住院费用记录.记录性质%Type,
    v_序号     住院费用记录.序号%Type,
    n_场合     Number := 0 --0住院，1门诊 
  ) Return Number Is
    n_Num      Number;
    n_执行状态 Number;
  Begin
    n_Num := 0;
    If n_场合 = 0 Then
      Select Nvl(Sum(Nvl(付数, 1) * 数次), 0) As 数量
      Into n_Num
      From 住院费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 In (2, 3);
      Select Nvl(执行状态, 0)
      Into n_执行状态
      From 住院费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 = 3;
    Else
      Select Nvl(Sum(Nvl(付数, 1) * 数次), 0) As 数量
      Into n_Num
      From 门诊费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 In (2, 3);
      Select Nvl(执行状态, 0)
      Into n_执行状态
      From 门诊费用记录
      Where NO = v_No And 记录性质 = v_记录性质 And 序号 = v_序号 And 记录状态 = 3;
    End If;
    If n_Num <> 0 Then
      n_Num := 1;
    End If;
    --如果主记录是已执行（部分执行的）则不自动退。 
    If n_执行状态 <> 0 Then
      n_Num := 0;
    End If;
    Return(n_Num);
  End;
Begin
  v_Tmp发送号 := -1;
  Open c_Rolladvice;
  Loop
    Fetch c_Rolladvice
      Into r_Rolladvice;
    If c_Rolladvice%RowCount = 0 Then
      Close c_Rolladvice;
      v_Error := Nvl(医嘱内容_In, '该医嘱') || '当前没有可以回退的内容。';
      Raise Err_Custom;
    End If;
    Exit When c_Rolladvice%NotFound;
    Exit When d_操作时间 <> r_Rolladvice.操作时间 And d_操作时间 Is Not Null;
    d_操作时间 := r_Rolladvice.操作时间;
  
    --批量回退调用时判断 
    If 医嘱内容_In Is Not Null Then
      If Nvl(r_Rolladvice.操作类型, 0) <> Nvl(操作类型_In, 0) Then
        v_Error := Nvl(医嘱内容_In, '该医嘱') || '不能与当前医嘱一起回退，可能该医嘱已经执行了其他操作。';
        Raise Err_Custom;
      End If;
    End If;
  
    --一组发送号只执行一次 
    If v_Tmp发送号 <> r_Rolladvice.发送号 Then
      v_Tmp发送号 := r_Rolladvice.发送号;
      n_执行      := 1;
    Else
      n_执行 := 0;
    End If;
  
    If n_执行 = 1 Then
      Open c_Adviceids;
      Fetch c_Adviceids Bulk Collect
        Into t_Adviceids;
      Close c_Adviceids;
    
      If r_Rolladvice.发送号 = 0 Then
        --回退医嘱状态操作(以时间关键字) 
        --4-作废；5-重整；6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果;13-停嘱申请 
        ------------------------------------------------------------------ 
        --最多只能退回到校对状态 
        If r_Rolladvice.操作类型 = 3 Then
          v_Error := Nvl(医嘱内容_In, '该医嘱') || '当前处于通过校对状态，不能再回退。';
          Raise Err_Custom;
        Elsif r_Rolladvice.操作类型 = 4 And Nvl(r_Rolladvice.婴儿, 0) = 0 Then
          If r_Rolladvice.类别 = 'H' Then
            Select 操作类型, 执行频率 Into v_操作类型, v_执行频率 From 诊疗项目目录 Where ID = r_Rolladvice.诊疗项目id;
            If v_操作类型 = '1' And v_执行频率 = '2' Then
              v_Error := '护理等级作废后不能再回退。';
              Raise Err_Custom;
            End If;
          End If;
        End If;
      
        --检查是否回退最近次重整之前的操作 
        If r_Rolladvice.操作类型 <> 5 Then
          --取最后重整时间 
          Select Nvl(医嘱重整时间, To_Date('1900-01-01', 'YYYY-MM-DD'))
          Into v_重整时间
          From 病案主页
          Where 病人id = r_Rolladvice.病人id And 主页id = r_Rolladvice.主页id;
        
          If r_Rolladvice.操作时间 < v_重整时间 Then
            v_Error := '该病人最近次重整之前的操作不能再回退。';
            Raise Err_Custom;
          End If;
        End If;
      
        --删除(该组医嘱)最近的状态操作记录 
        Delete /*+ Rule*/
        From 病人医嘱状态
        Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And 操作时间 = r_Rolladvice.操作时间;
      
        --取删除后应恢复的医嘱状态 
        Select 操作类型
        Into v_医嘱状态
        From 病人医嘱状态
        Where 操作时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = 医嘱id_In) And 医嘱id = 医嘱id_In;
      
        --恢复(该组医嘱)回退后的状态 
        Update 病人医嘱记录 Set 医嘱状态 = v_医嘱状态 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
      
        --其它额外的处理 
        If r_Rolladvice.操作类型 = 8 Then
          --被超期发送收回过的医嘱 ，如果是销帐申请模式，则判断对应的“病人费用销帐”申请是否取消，是则允许回退，否则不允许， 
          --                       如果是产生负数费用模式，则不允许再回退。 
          --可能超期发送收回时被全部收回(无上次执行时间) 
          Select /*+ Rule*/
           Nvl(Count(*), 0)
          Into v_Count
          From 病人医嘱记录 A, 病人医嘱发送 B
          Where b.医嘱id = a.Id And (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) And
                b.发送号 =
                (Select Max(发送号) From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(t_Adviceids))) And
                a.执行终止时间 Is Not Null And ((a.上次执行时间 < b.末次时间) Or (a.上次执行时间 Is Null And b.末次时间 Is Not Null));
          If v_Count > 0 Then
            If zl_GetSysParameter('超期收回产生负数费用', 1254) = '1' Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '已被超期发送收回，不能再撤消停止操作。';
              Raise Err_Custom;
            Else
              --如果已经取消销帐申请，则允许回退. 
              Select Count(1)
              Into v_Count
              From 病人费用销帐 A, 住院费用记录 B, 病人医嘱记录 C
              Where a.费用id = b.Id And c.Id = b.医嘱序号 And (c.Id = 医嘱id_In Or c.相关id = 医嘱id_In);
              If v_Count > 0 Then
                v_Error := Nvl(医嘱内容_In, '该医嘱') || '已被超期发送收回，不能再撤消停止操作。';
                Raise Err_Custom;
              Else
                --得到上次执行时间等信息 
                Select 上次执行时间, 执行时间方案, 开始执行时间, 上次打印时间, 频率间隔, 间隔单位
                Into v_上次时间, v_执行时间, v_开始执行时间, v_上次打印时间, v_频率间隔, v_间隔单位
                From 病人医嘱记录
                Where ID = 医嘱id_In;
                v_上次时间 := To_Date(To_Char(v_上次时间 + 1 / 24 / 60 / 60, 'yyyy-MM-dd hh24:mi:ss'), 'yyyy-MM-dd hh24:mi:ss');
              
                --修改上次执行时间为收回后的末次执行时间。 
                v_末次时间 := Null;
                Begin
                  --一组医嘱的发送首末时间相同,一并给药是取最小的 
                  --取相关ID为NULL的医嘱的发送记录的时间 
                  --但给药途径或中药用法可能未填写发送记录 
                  Select /*+ Rule*/
                   末次时间, 发送号
                  Into v_末次时间, v_发送号
                  From 病人医嘱发送
                  Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And
                        发送号 = (Select Max(发送号)
                               From 病人医嘱发送
                               Where 医嘱id In (Select Column_Value From Table(t_Adviceids))) And Rownum = 1;
                Exception
                  When Others Then
                    Null;
                End;
                Update 病人医嘱记录 Set 上次执行时间 = v_末次时间 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
              
                Select Count(1) Into v_Count From 病人医嘱发送 Where 医嘱id = 医嘱id_In And 发送号 = v_发送号;
                If v_Count > 0 Then
                  --还原医嘱执行时间 
                  Select Zl_Adviceexetimes(医嘱id_In, v_上次时间, v_末次时间, v_执行时间, v_开始执行时间, v_上次打印时间, v_频率间隔, v_间隔单位, 0)
                  Into v_Time
                  From Dual;
                  Insert Into 医嘱执行时间
                    (要求时间, 医嘱id, 发送号)
                    Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), 医嘱id_In, v_发送号
                    From Table(f_Str2list(v_Time));
                End If;
              End If;
            End If;
          End If;
        
          --护理等级变动，后续有其他变动时，不允许回退 
          If r_Rolladvice.类别 = 'H' And Nvl(r_Rolladvice.婴儿, 0) = 0 Then
            Select 操作类型, 执行频率 Into v_操作类型, v_执行频率 From 诊疗项目目录 Where ID = r_Rolladvice.诊疗项目id;
            If v_操作类型 = '1' And v_执行频率 = '2' Then
              Select Count(*), Max(a.护理等级id), Max(a.开始时间)
              Into v_Count, n_护理等级id, d_开始时间
              From 病人变动记录 A
              Where a.病人id = r_Rolladvice.病人id And a.主页id = r_Rolladvice.主页id And a.开始原因 = 6 And a.终止时间 Is Null And
                    a.附加床位 = 0;
              --如果没有找到最后一条是护理等级变动则禁止 
              If v_Count = 0 Then
                --医嘱护理等级和入住时候的护理等级一致时要单独判断 
                Select Count(*)
                Into v_Count
                From 病人变动记录 A
                Where a.病人id = r_Rolladvice.病人id And a.主页id = r_Rolladvice.主页id And a.开始原因 = 6;
                If v_Count > 0 Then
                  v_Error := '由于护理等级医嘱停止后该病人已经产生了其他变动记录,不能回退该医嘱的停止操作。';
                  Raise Err_Custom;
                End If;
              Else
                --如果n_护理等级ID为Null，则检查是否是当前回退的医嘱对应的变动记录,目的是有多个护理等级医嘱时要求按顺序回退。 
                --如果n_护理等级ID不为Null，则有可能是校对下一条护理等级时，自动停止的，未产生变动记录， 
                --     则需要检查当前最后一条变动的护理等级ID是否是当前医嘱的护理等级ID,目的是有多个护理等级医嘱时要求按顺序回退，如果是则不需要再撤销最后一次变动，直接回退医嘱即可。 
                If n_护理等级id Is Null Then
                  Select Count(*)
                  Into v_Count
                  From 病人变动记录 B, 病人医嘱计价 C
                  Where b.病人id = r_Rolladvice.病人id And b.主页id = r_Rolladvice.主页id And c.医嘱id = 医嘱id_In And
                        c.收费细目id = b.护理等级id And b.终止时间 = d_开始时间 And b.终止原因 = 6 And b.附加床位 = 0;
                Else
                  --开始时间只取分钟对比，校对的时候护理等级的开始时间是医嘱开始时间+当前时间的秒钟 
                  Select Count(*)
                  Into v_Count
                  From 病人医嘱计价 C, 病人医嘱记录 A
                  Where a.Id = c.医嘱id And a.Id = 医嘱id_In And c.收费细目id = n_护理等级id And
                        a.开始执行时间 = To_Date(To_Char(d_开始时间, 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi');
                End If;
                If v_Count = 0 Then
                  v_Error := '您回退的医嘱不是最后一条护理等级医嘱，请将后面的护理等级医嘱作废后再回退本条医嘱。';
                  Raise Err_Custom;
                End If;
              
                If n_护理等级id Is Null Then
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
                
                  Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, '1', Null, Null, '护理等级变动');
                End If;
              End If;
            End If;
          End If;
        
          If r_Rolladvice.类别 = 'Z' And Instr(',9,10,', ',' || r_Rolladvice.类型 || ',') > 0 And
             Nvl(r_Rolladvice.婴儿, 0) = 0 Then
            --回退病况医嘱时，调用变动记录回退 
            Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, Null, Null, Null, '病况变动');
          End If;
        
          --回退医嘱停止时,清空停嘱医生和时间,如果是实习医师申请后审核的，则恢复待审核状态 
          Update 病人医嘱记录
          Set 执行终止时间 = Decode(Flag_In, 1, 执行终止时间, Null), 停嘱医生 = Null, 停嘱时间 = Null,
              审核标记 = Decode(r_Rolladvice.审核标记, 3, 2, r_Rolladvice.审核标记)
          Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        Elsif r_Rolladvice.操作类型 = 9 Then
          --回退医嘱确认停止时,检查是否已打印停嘱时间 
          Select /*+ Rule*/
           Count(*)
          Into v_Count
          From 病人医嘱打印
          Where 打印标记 = 1 And 医嘱id In (Select Column_Value From Table(t_Adviceids));
          If v_Count > 0 Then
            v_Error := Nvl(医嘱内容_In, '该医嘱') || '的停嘱时间已经打印，不能再撤消确认停止操作。';
            Raise Err_Custom;
          End If;
        
          --回退医嘱确认停止时,清空停嘱医生和时间 
          Update 病人医嘱记录 Set 确认停嘱时间 = Null, 确认停嘱护士 = Null Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        Elsif r_Rolladvice.操作类型 = 10 Then
          --回退标注皮试结果,同时删除过敏登记(+)或(-),根据记录时间 
          --过敏的记录与医嘱操作无观，不需要处理 
          Delete From 病人过敏记录
          Where 病人id = r_Rolladvice.病人id And Nvl(主页id, 0) = Nvl(r_Rolladvice.主页id, 0) And 记录时间 = r_Rolladvice.操作时间 And
                Nvl(结果, 0) = 0;
        
          Update 病人医嘱记录 Set 皮试结果 = Null Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        Elsif r_Rolladvice.操作类型 = 13 Then
          If Instr(r_Rolladvice.开嘱医生, '/') > 0 Then
            Update 病人医嘱记录 Set 审核标记 = 1 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
          Else
            Update 病人医嘱记录 Set 审核标记 = Null Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
          End If;
        End If;
        --回退术后医嘱作废操作 
        --术后医嘱 
        If r_Rolladvice.操作类型 = 4 And r_Rolladvice.类别 = 'Z' Then
          Select Count(1) Into v_Count From 诊疗项目目录 Where ID = r_Rolladvice.诊疗项目id And 操作类型 = '4';
          If v_Count = 1 Then
            b_Message.Zlhis_Cis_004(r_Rolladvice.病人id, r_Rolladvice.主页id, 医嘱id_In);
          End If;
        End If;
      Else
        --回退医嘱发送(以发送号关键字) 
        ------------------------------------------------------------------ 
        --当前操作人员 
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_人员编号 := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      
        --检查是否是输液配液记录，并是否已经锁定，如果查询有数据说明是配液记录 
        Begin
          Select Decode(Max(是否锁定), 1, 1, 0)
          Into v_Count
          From 输液配药记录
          Where 医嘱id = 医嘱id_In And 发送号 = r_Rolladvice.发送号;
        Exception
          When Others Then
            v_Count := -1;
        End;
      
        If v_Count = 1 Then
          v_Error := '医嘱"' || 医嘱内容_In || '"是输液药品，已经被输液配置中心锁定，不能回退发送。';
          Raise Err_Custom;
        Elsif v_Count = 0 Then
          Zl_输液配药记录_医嘱回退(医嘱id_In, r_Rolladvice.发送号, v_人员姓名, Sysdate);
        End If;
      
        --检查是否存在未审核的销帐申请 
        Select Count(*)
        Into v_Count
        From 病人医嘱记录 A, 病人医嘱发送 B, 住院费用记录 C, 病人费用销帐 D
        Where (a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) And a.Id = b.医嘱id And b.医嘱id = c.医嘱序号 And c.Id = d.费用id And
              c.记录状态 In (0, 1, 3) And d.状态 = 0;
      
        If v_Count > 0 Then
          v_Error := '医嘱"' || 医嘱内容_In || '"存在未审核的销帐申请，请取消或审核销帐申请后再回退发送。';
          Raise Err_Custom;
        End If;
      
        --检查医嘱是否存在有效的医嘱附费 
        Select Count(*)
        Into v_Count
        From 病人医嘱附费 A, 住院费用记录 B
        Where a.医嘱id = b.医嘱序号 And a.No = b.No And b.记录状态 = 1 And b.实收金额 <> 0 And a.发送号 = r_Rolladvice.发送号 And
              a.医嘱id In (Select Column_Value From Table(t_Adviceids));
        If v_Count > 0 Then
          v_Error := '该医嘱下还存在附费项目，请先冲销。';
          Raise Err_Custom;
        End If;
      
        --本科发送自动执行时，回退也自动回退执行(仅护士站有此功能) 
        --非跟踪在用的卫材医嘱，同普通医嘱执行处理 
        Select 医嘱期效 Into v_医嘱期效 From 病人医嘱记录 Where ID = 医嘱id_In;
        If Substr(zl_GetSysParameter('本科执行自动完成', 1254), v_医嘱期效 + 1, 1) = '1' Then
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into Intdigit From Dual;
        
          For r_Rollsend In c_Rollsend(r_Rolladvice.发送号) Loop
            If Nvl(r_Rollsend.执行状态, 0) = 1 And
               (Nvl(r_Rollsend.执行科室id, 0) = Nvl(r_Rollsend.病人病区id, 0) Or
                Nvl(r_Rollsend.执行科室id, 0) = Nvl(r_Rollsend.病人科室id, 0)) Then
            
              --医嘱的执行状态 
              Update 病人医嘱发送 Set 执行状态 = 0 Where 发送号 = r_Rollsend.发送号 And 医嘱id = r_Rollsend.医嘱id;
              v_Update := 1;
            
              If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 0 Then
                --费用的执行状态 
                For r_Rollmoney In c_Rollmoneyin(r_Rollsend.发送号, r_Rollsend.医嘱id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.记录状态 <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1) And
                       Not r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --普通费用直接取消执行状态，不含药品和跟踪在用的卫材 
                      Update 住院费用记录
                      Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
                      Where NO = r_Rollmoney.No And 记录性质 = r_Rolladvice.记录性质 And 记录状态 = r_Rollmoney.记录状态 And
                            Nvl(价格父号, 序号) = r_Rollmoney.序号 And 医嘱序号 = r_Rollsend.医嘱id;
                    Elsif r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1 Then
                      --跟踪在用的卫材，才自动退料 
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_材料收发记录_部门退料(r_Stuff.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, 0, v_人员姓名);
                      End Loop;
                    Elsif r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --住院科室发药的药品自动退药 
                      If r_Rollmoney.执行部门id = r_Rollsend.病人病区id Or r_Rollmoney.执行部门id = r_Rollsend.病人科室id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_药品收发记录_部门退药(r_Drug.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 2);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              Else
                --住院病人费用发送到门诊的情况，病人来源都是住院的 
                For r_Rollmoney In c_Rollmoneyout(r_Rollsend.发送号, r_Rollsend.医嘱id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.记录状态 <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号, 1);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1) And
                       Not r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --普通费用直接取消执行状态，不含药品和跟踪在用的卫材 
                      Update 门诊费用记录
                      Set 执行状态 = 0, 执行时间 = Null, 执行人 = Null
                      Where NO = r_Rollmoney.No And 记录性质 = r_Rolladvice.记录性质 And 记录状态 = r_Rollmoney.记录状态 And
                            Nvl(价格父号, 序号) = r_Rollmoney.序号 And 医嘱序号 = r_Rollsend.医嘱id;
                    Elsif r_Rollmoney.收费类别 = '4' And Nvl(r_Rollmoney.跟踪在用, 0) = 1 Then
                      --跟踪在用的卫材，才自动退料 
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_材料收发记录_部门退料(r_Stuff.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, 0, v_人员姓名);
                      End Loop;
                    Elsif r_Rollmoney.收费类别 In ('5', '6', '7') Then
                      --本科室发药的药品自动退药 
                      If r_Rollmoney.执行部门id = r_Rollsend.病人病区id Or r_Rollmoney.执行部门id = r_Rollsend.病人科室id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_药品收发记录_部门退药(r_Drug.Id, v_人员姓名, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 1);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              End If;
            End If;
          End Loop;
        End If;
        ------------------------------------------------------------------ 
        --被超期收回的长期药品医嘱不允许回退(再退费用就多退了) 
        If Nvl(r_Rolladvice.医嘱期效, 0) = 0 Then
          If r_Rolladvice.上次执行时间 Is Not Null And r_Rolladvice.末次时间 Is Not Null Then
            If r_Rolladvice.上次执行时间 < r_Rolladvice.末次时间 Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '最近超期发送的内容已被收回，不能再回退。';
              Raise Err_Custom;
            End If;
          Elsif r_Rolladvice.上次执行时间 Is Null And r_Rolladvice.末次时间 Is Not Null Then
            --长嘱可能被全部超期收回 
            v_Error := Nvl(医嘱内容_In, '该医嘱') || '未被发送，或发送的内容已被全部超期收回，不能再回退。';
            Raise Err_Custom;
          End If;
        End If;
      
        If Nvl(r_Rolladvice.执行状态, 0) In (1, 3) And v_Update <> 1 Then
          --1-完全执行;3-正在执行 
          v_Error := Nvl(医嘱内容_In, '该医嘱') || '最近发送的内容已经执行或正在执行，不能回退。';
          Raise Err_Custom;
        Else
          --如果相关医嘱已执行，则也要限制回退（例如：检验的采集方式） 
          Select /*+ Rule*/
           Count(1)
          Into v_Count
          From 病人医嘱发送
          Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And 执行状态 In (1, 3) And
                发送号 =
                (Select Max(发送号) From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(t_Adviceids)));
          If v_Count > 0 Then
            v_Error := Nvl(医嘱内容_In, '该医嘱') || '最近发送的内容已经执行或正在执行，不能回退。';
            Raise Err_Custom;
          End If;
        End If;
      
        ------------------------------------------------------------------ 
        --将该组医嘱的费用销帐(按一组医嘱可能有不同NO处理) 
        --如果原始费用已被销帐(或部分销帐),调用过程中有判断 
        v_费用no   := Null;
        v_费用序号 := Null;
        If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 0 Then
          For r_Rollmoney In c_Rollmoneyin(r_Rolladvice.发送号, Null, t_Adviceids) Loop
            --对应的费用已执行 
            If Nvl(r_Rollmoney.执行状态, 0) <> 0 Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '发送的费用单据"' || r_Rollmoney.No || '"中的内容已被部分或完全执行，不能回退。';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.记录状态 <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号);
            End If;
            If n_Blndo > 0 Then
              --这种仅用于判断部分退药 
              If v_费用no <> r_Rollmoney.No And v_费用序号 Is Not Null Then
                Zl_住院记帐记录_Delete(v_费用no, Substr(v_费用序号, 2), v_人员编号, v_人员姓名, 2, 0, 0);
                v_费用序号 := Null;
              End If;
              v_费用no   := r_Rollmoney.No;
              v_费用序号 := v_费用序号 || ',' || r_Rollmoney.序号;
            End If;
          End Loop;
        Else
          For r_Rollmoney In c_Rollmoneyout(r_Rolladvice.发送号, Null, t_Adviceids) Loop
            --对应的费用已执行 
            If Nvl(r_Rollmoney.执行状态, 0) <> 0 And Not (Nvl(r_Rollmoney.执行状态, 0) = -1 And Nvl(r_Rollmoney.记录状态, 0) = 0) Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '发送的费用单据"' || r_Rollmoney.No || '"中的内容已被部分或完全执行，不能回退。';
              Raise Err_Custom;
            End If;
            --收费单据已收费 
            If r_Rollmoney.记录状态 = 1 And Not (r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 1) Then
              v_Error := Nvl(医嘱内容_In, '该医嘱') || '发送的门诊单据"' || r_Rollmoney.No || '"已收费，不能回退。';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.记录状态 <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.记录性质, r_Rollmoney.序号, 1);
            End If;
            If n_Blndo > 0 Then
              --这种仅用于判断部分退药 
              If v_费用no <> r_Rollmoney.No And v_费用序号 Is Not Null Then
                If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 1 Then
                  --住院发送为门诊记帐(如果是门诊医生发送为门诊记帐，门诊医嘱没有回退功能) 
                  Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名);
                Else
                  Zl_门诊划价记录_Delete(v_费用no, Substr(v_费用序号, 2));
                End If;
                v_费用序号 := Null;
              End If;
              v_费用no   := r_Rollmoney.No;
              v_费用序号 := v_费用序号 || ',' || r_Rollmoney.序号;
            End If;
          End Loop;
        End If;
        If v_费用序号 Is Not Null And v_费用no Is Not Null Then
          v_费用序号 := Substr(v_费用序号, 2);
          If r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 0 Then
            Zl_住院记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名, 2, 0, 0);
          Elsif r_Rolladvice.记录性质 = 2 And Nvl(r_Rolladvice.门诊记帐, 0) = 1 Then
            --住院发送为门诊记帐(如果是门诊医生发送为门诊记帐，门诊医嘱没有回退功能) 
            Zl_门诊记帐记录_Delete(v_费用no, v_费用序号, v_人员编号, v_人员姓名);
          Else
            Zl_门诊划价记录_Delete(v_费用no, v_费用序号);
          End If;
        End If;
      
        --通过发送记录产生消息，对于组合类医嘱要提前      
        For R In (Select a.病人id, a.主页id, b.No, b.发送号, b.发送数次, b.首次时间, b.末次时间, b.样本条码, a.Id, a.诊疗类别, c.操作类型, a.执行科室id
                  From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C
                  Where a.Id = b.医嘱id And a.诊疗项目id = c.Id And b.发送号 = r_Rolladvice.发送号 And
                        b.医嘱id In (Select Column_Value From Table(t_Adviceids))) Loop
          If r.诊疗类别 = 'E' And r.操作类型 = '6' Then
            --检验
            b_Message.Zlhis_Cis_036(r.病人id, r.主页id, Null, r.发送号, r.Id, r.No, 2);
          End If;
          Select Count(1) Into v_Count From 部门性质说明 A Where a.部门id = r.执行科室id And a.工作性质 = '护理';
          If v_Count > 0 Then
            --病区执行医嘱回退发送
            b_Message.Zlhis_Cis_044(r.病人id, r.主页id, r.发送号, r.Id, r.No, r.发送数次, r.首次时间, r.末次时间, r.样本条码);
          End If;
        End Loop;
      
        --输血医嘱先删除病人医嘱附费 
        Delete From 病人医嘱附费 Where 发送号 = r_Rolladvice.发送号 And 医嘱id = 医嘱id_In;
      
        --删除医嘱执行时间 (仅主医嘱ID才产生了记录) 
        Delete From 医嘱执行时间 Where 发送号 = r_Rolladvice.发送号 And 医嘱id = 医嘱id_In;
      
        --删除发送记录(该组医嘱的) 
        Delete /*+ Rule*/
        From 病人医嘱发送
        Where 发送号 = r_Rolladvice.发送号 And 医嘱id In (Select Column_Value From Table(t_Adviceids));
      
        --标记(该组医嘱)上次执行时间(以上次发送的末次执行时间) 
        --所有长嘱(包括持续性长嘱)发送时都填写了末次时间 
        --临嘱可能没有，且只可能发送了一次。 
        v_末次时间 := Null;
        Begin
          --一组医嘱的发送首末时间相同,一并给药是取最小的 
          --取相关ID为NULL的医嘱的发送记录的时间 
          --但给药途径或中药用法可能未填写发送记录 
          Select /*+ Rule*/
           末次时间
          Into v_末次时间
          From 病人医嘱发送
          Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And
                发送号 =
                (Select Max(发送号) From 病人医嘱发送 Where 医嘱id In (Select Column_Value From Table(t_Adviceids))) And
                Rownum = 1;
        Exception
          When Others Then
            Null;
        End;
        Update 病人医嘱记录 Set 上次执行时间 = v_末次时间 Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
      
        --回退临嘱发送时，同时自动回退停止 
        If Nvl(r_Rolladvice.医嘱期效, 0) = 1 Then
          --删除(该组临嘱)最近的停止状态操作记录 
          Delete /*+ Rule*/
          From 病人医嘱状态
          Where 医嘱id In (Select Column_Value From Table(t_Adviceids)) And
                操作时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = 医嘱id_In) And 操作类型 = 8;
          --r_RollAdvice.操作时间:因发送时间可能不与自动停止时间相同。 
        
          --取删除后应恢复的医嘱状态 
          Select 操作类型
          Into v_医嘱状态
          From 病人医嘱状态
          Where 操作时间 = (Select Max(操作时间) From 病人医嘱状态 Where 医嘱id = 医嘱id_In) And 医嘱id = 医嘱id_In;
        
          --恢复(该组医嘱)回退后的状态 
          Update 病人医嘱记录
          Set 医嘱状态 = v_医嘱状态, 执行终止时间 = Null, 停嘱医生 = Null, 停嘱时间 = Null
          Where ID = 医嘱id_In Or 相关id = 医嘱id_In;
        End If;
      
        --住院特殊医嘱发送后的回退(3-转科;5-出院;6-转院,11-死亡) 
        If r_Rolladvice.类别 = 'Z' And Instr(',3,5,6,11,', ',' || r_Rolladvice.类型 || ',') > 0 And
           Nvl(r_Rolladvice.婴儿, 0) = 0 Then
          Open c_Patilog(r_Rolladvice.病人id, r_Rolladvice.主页id);
          Fetch c_Patilog
            Into r_Patilog;
          If c_Patilog%Found Then
            If r_Rolladvice.类型 = '3' And r_Patilog.开始原因 = 3 Then
              --取消病人转科状态 
              If r_Patilog.开始时间 Is Null Then
                --转科医嘱的特殊处理，当一个病人有两条转科医嘱时，只能回退最近的一条,70443 
                Select Count(1)
                Into v_Count
                From 病人医嘱记录 A, 诊疗项目目录 B
                Where a.诊疗项目id = b.Id And a.病人id = r_Rolladvice.病人id And a.主页id = r_Rolladvice.主页id And a.诊疗类别 = 'Z' And
                      b.操作类型 = '3' And a.医嘱状态 = 8 And
                      a.开始执行时间 > (Select 开始执行时间 From 病人医嘱记录 Where ID = 医嘱id_In);
                If v_Count = 0 Then
                  Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, Null, Null, Null, '转科');
                Else
                  v_Error := '病人转科已经入科，不能再回退。';
                  Raise Err_Custom;
                End If;
              Else
                v_Error := '病人转科已经入科，不能再回退。';
                Raise Err_Custom;
              End If;
            Elsif r_Rolladvice.类型 In ('5', '6', '11') And r_Patilog.开始原因 = 10 Then
              --取消病人预出院状态 
              Zl_病人变动记录_Undo(r_Rolladvice.病人id, r_Rolladvice.主页id, v_人员编号, v_人员姓名, Null, Null, Null, '预出院');
            End If;
          End If;
          Close c_Patilog;
        End If;
      
        --回退病历时机 
        --1.特殊事件(只有一条医嘱记录)：手术，7-会诊,8-抢救,11-死亡 
        If r_Rolladvice.类别 = 'F' Or r_Rolladvice.类别 = 'Z' And Instr(',7,8,11,', ',' || r_Rolladvice.类型 || ',') > 0 Then
          Zl_电子病历时机_Delete(r_Rolladvice.病人id, r_Rolladvice.主页id, '医嘱', r_Rolladvice.开嘱科室id, 医嘱id_In);
        End If;
      
        --2.额外处理：知情同意书(手术相关的知情同意需再次调用，因为附加手术或麻醉项目可能有关联的知情同意书) 
        If Instr('C,D,E,F,G,K,L', r_Rolladvice.类别) > 0 Then
          For R In (Select a.Id, a.诊疗类别 From 病人医嘱记录 A Where a.Id = 医嘱id_In Or a.相关id = 医嘱id_In) Loop
            --相关id的一组医嘱不一定是这个类别的，所以要再判断一次类别 
            If Instr('C,D,E,F,G,K,L', r.诊疗类别) > 0 Then
              Zl_电子病历时机_Delete(r_Rolladvice.病人id, r_Rolladvice.主页id, '医嘱', r_Rolladvice.开嘱科室id, r.Id);
            End If;
          End Loop;
        End If;
      
        --此处添加消息触发
        If r_Rolladvice.类别 = 'D' And r_Rolladvice.相关id Is Null Then
          --检查 
          b_Message.Zlhis_Cis_037(r_Rolladvice.病人id, r_Rolladvice.主页id, Null, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No, 2);
        Elsif r_Rolladvice.类别 = 'F' And r_Rolladvice.相关id Is Null Then
          --手术 
          b_Message.Zlhis_Cis_038(r_Rolladvice.病人id, r_Rolladvice.主页id, Null, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.类别 = 'K' And r_Rolladvice.相关id Is Null Then
          --输血 
          b_Message.Zlhis_Cis_039(r_Rolladvice.病人id, r_Rolladvice.主页id, Null, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.类别 = 'Z' And r_Rolladvice.操作类型 = '6' Then
          --会诊 
          b_Message.Zlhis_Cis_040(r_Rolladvice.病人id, r_Rolladvice.主页id, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.类别 = 'Z' And r_Rolladvice.操作类型 = '8' Then
          --抢救 
          b_Message.Zlhis_Cis_041(r_Rolladvice.病人id, r_Rolladvice.主页id, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.类别 = 'Z' And r_Rolladvice.操作类型 = '11' Then
          --死亡 
          b_Message.Zlhis_Cis_042(r_Rolladvice.病人id, r_Rolladvice.主页id, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.类别 = 'E' And r_Rolladvice.操作类型 = '5' Then
          --特殊治疗 
          b_Message.Zlhis_Cis_043(r_Rolladvice.病人id, r_Rolladvice.主页id, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.类别 = 'H' And Nvl(r_Rolladvice.操作类型, '0') = '0' Then
          --护理常规 
          b_Message.Zlhis_Cis_007(r_Rolladvice.病人id, r_Rolladvice.主页id, r_Rolladvice.发送号, r_Rolladvice.组id,
                                  r_Rolladvice.No);
        End If;
      End If;
    End If;
    Exit When r_Rolladvice.发送号 = 0;
  End Loop;
  Close c_Rolladvice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_回退;
/

--130541:胡俊勇,2018-08-23,集成平台消息问题
Create Or Replace Procedure Zl_病人医嘱记录_确认停止
(
  --功能：确认停止指定的医嘱 
  --说明：一并给药的只能调用一次 
  --参数：医嘱ID=相关ID为NULL的医嘱的ID(给药途径,中药用法,检查项目,主要手术,及独立医嘱) 
  医嘱id_In           In 病人医嘱记录.Id%Type,
  确认时间_In         In 病人医嘱记录.确认停嘱时间%Type,
  操作员姓名_In       In 人员表.姓名%Type := Null,
  自动确认护理等级_In In Number := 0
) Is
  v_状态     病人医嘱记录.医嘱状态%Type;
  v_医嘱内容 病人医嘱记录.医嘱内容%Type;
  n_病人id   病人医嘱记录.病人id%Type;
  n_主页id   病人医嘱记录.主页id%Type;

  v_Temp     Varchar2(255);
  v_人员姓名 病人医嘱状态.操作人员%Type;
  n_Count    Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --检查医嘱状态是否正确:并发操作 
  Select 医嘱状态, 医嘱内容, 病人id, 主页id
  Into v_状态, v_医嘱内容, n_病人id, n_主页id
  From 病人医嘱记录
  Where ID = 医嘱id_In;
  If v_状态 <> 8 Then
    v_Error := '医嘱"' || v_医嘱内容 || '"当前不处于停止状态。';
    Raise Err_Custom;
  End If;
  --检查是否是输液配液记录，并是否已经锁定 
  Select Count(1)
  Into n_Count
  From 输液配药记录 A, 病人医嘱记录 B
  Where a.医嘱id = b.Id And 医嘱id = 医嘱id_In And a.执行时间 > b.执行终止时间 And a.是否锁定 = 1;
  If n_Count > 0 Then
    v_Error := '医嘱"' || v_医嘱内容 || '"是输液药品，已经被输液配置中心锁定，不能确认停止。';
    Raise Err_Custom;
  End If;

  --当前操作人员 
  If 操作员姓名_In Is Not Null Then
    v_人员姓名 := 操作员姓名_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_人员姓名 := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  Update 病人医嘱记录
  Set 医嘱状态 = 9, 确认停嘱时间 = 确认时间_In, 确认停嘱护士 = v_人员姓名
  Where ID = 医嘱id_In Or 相关id = 医嘱id_In;

  Insert Into 病人医嘱状态
    (医嘱id, 操作类型, 操作人员, 操作时间)
    Select ID, 9, v_人员姓名, Sysdate + 自动确认护理等级_In / 24 / 60 / 60
    From 病人医嘱记录
    Where ID = 医嘱id_In Or 相关id = 医嘱id_In;

  b_Message.Zlhis_Cis_059(n_病人id, n_主页id, 医嘱id_In);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_病人医嘱记录_确认停止;
/

--130471:余伟节,2018-08-22,预约安排
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

--125796:刘涛,2018-08-22,增加冲销记录字段的应用
Create Or Replace Procedure Zl_药品领用_Strike
(
  行次_In       In Integer,
  原记录状态_In In 药品收发记录.记录状态%Type,
  No_In         In 药品收发记录.No%Type,
  序号_In       In 药品收发记录.序号%Type,
  药品id_In     In 药品收发记录.药品id%Type,
  冲销数量_In   In 药品收发记录.实际数量%Type,
  填制人_In     In 药品收发记录.填制人%Type,
  填制日期_In   In 药品收发记录.填制日期%Type,
  冲销方式_In   In Integer := 0, --0－正常冲销方式；1－产生冲销申请单据；2－审核已产生的冲销申请单据  
  冲销原因_In   In 药品收发记录.冲销原因%Type := Null
) Is
  Err_Isstriked Exception;
  Err_Isoutstock Exception;
  Err_Isnonum Exception;
  Err_Isbatch Exception;
  v_Druginf Varchar2(50); --原不分批现在分批的药品信息 

  v_库房id       药品收发记录.库房id%Type;
  v_对方部门id   药品收发记录.对方部门id%Type;
  v_入出类别id   药品收发记录.入出类别id%Type;
  v_产地         药品收发记录.产地%Type;
  v_原产地       药品收发记录.原产地%Type;
  v_批次         药品收发记录.批次%Type;
  v_批号         药品收发记录.批号%Type;
  v_效期         药品收发记录.效期%Type;
  v_成本价       药品收发记录.成本价%Type;
  v_成本金额     药品收发记录.成本金额%Type;
  v_扣率         药品收发记录.扣率%Type;
  v_零售价       药品收发记录.零售价%Type;
  v_零售金额     药品收发记录.零售金额%Type;
  v_差价         药品收发记录.差价%Type;
  v_摘要         药品收发记录.摘要%Type;
  v_剩余数量     药品收发记录.实际数量%Type;
  v_剩余成本金额 药品收发记录.成本金额%Type;
  v_剩余零售金额 药品收发记录.零售金额%Type;
  v_入出系数     药品收发记录.入出系数%Type;

  v_收发id   药品收发记录.Id%Type;
  v_领用人   药品收发记录.领用人%Type;
  v_批准文号 药品收发记录.批准文号%Type;
  v_发药方式 药品收发记录.发药方式%Type;

  v_是否变价     收费项目目录.是否变价%Type;
  Intdigit       Number;
  n_上次供应商id 药品库存.上次供应商id%Type;
  d_上次生产日期 药品库存.上次生产日期%Type;
  v_按月留存领用 Varchar2(4000);
Begin
  --获取金额小数位数 
  Select Nvl(精度, 2) Into Intdigit From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 内容 = 4 And 单位 = 5;
  Select Nvl(是否变价, 0) Into v_是否变价 From 收费项目目录 Where Id = 药品id_In;
  Select Zl_Getsysparameter('按月留存领用', 1305) Into v_按月留存领用 From Dual;

  If 冲销方式_In = 1 Then
    --产生冲销申请单据，不填写审核人、审核日期，不更新库存记录 
    If 行次_In = 1 Then
      Update 药品收发记录
      Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
      Where No = No_In And 单据 = 7 And 记录状态 = 原记录状态_In;
      If Sql%Rowcount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --主要针对原不分批现在分批的药品，不能对其冲销
    Begin
      Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
      Into v_Druginf
      From 药品收发记录 a, 药品规格 b, 收费项目目录 i, 收费项目别名 n
      Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 7 And
            Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) = 0 And a.药品id + 0 = 药品id_In And
            ((Nvl(b.药库分批, 0) = 1 And
            a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or
            Nvl(b.药房分批, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(实际数量) As 剩余数量, Sum(成本金额) As 剩余成本金额, Sum(零售金额) As 剩余零售金额, 库房id, 对方部门id, 入出类别id, 入出系数, 批次, 产地, 原产地, 批号, 效期, 成本价,
           扣率, 零售价, 摘要, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期
    Into v_剩余数量, v_剩余成本金额, v_剩余零售金额, v_库房id, v_对方部门id, v_入出类别id, v_入出系数, v_批次, v_产地, v_原产地, v_批号, v_效期, v_成本价, v_扣率, v_零售价,
         v_摘要, v_领用人, v_批准文号, v_发药方式, n_上次供应商id, d_上次生产日期
    From 药品收发记录
    Where No = No_In And 单据 = 7 And 药品id = 药品id_In And 序号 = 序号_In
    Group By 库房id, 对方部门id, 入出类别id, 入出系数, 批次, 产地, 原产地, 批号, 效期, 成本价, 扣率, 零售价, 摘要, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期;
  
    --冲销数量大于剩余数量，不允许 
    If v_剩余数量 < 冲销数量_In Then
      Raise Err_Isnonum;
    End If;
  
    v_成本金额 := Round(冲销数量_In / v_剩余数量 * v_剩余成本金额, Intdigit);
    v_零售金额 := Round(冲销数量_In / v_剩余数量 * v_剩余零售金额, Intdigit);
    v_差价     := v_零售金额 - v_成本金额;
  
    Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
  
    Insert Into 药品收发记录
      (Id, 记录状态, 单据, No, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 原产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
       填制人, 填制日期, 审核人, 审核日期, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期, 扣率, 冲销原因)
    Values
      (v_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 7, No_In, 序号_In, v_库房id, v_对方部门id, v_入出类别id, v_入出系数, 药品id_In, v_批次,
       v_产地, v_原产地, v_批号, v_效期, -冲销数量_In, -冲销数量_In, v_成本价, -v_成本金额, v_零售价, -v_零售金额, -v_差价, v_摘要, 填制人_In, 填制日期_In, Null, Null,
       v_领用人, v_批准文号, v_发药方式, n_上次供应商id, d_上次生产日期, v_扣率, 冲销原因_In);
    
    Zl_未审药品记录_Insert(v_收发id);
    
  Elsif 冲销方式_In = 2 Then
    --审核已产生的冲销申请单据，填写审核人、审核日期，更新库存记录 
  
    --填写审核人、审核日期 
    Update 药品收发记录
    Set 审核人 = 填制人_In, 审核日期 = 填制日期_In
    Where 单据 = 7 And No = No_In And 序号 = 序号_In And 记录状态 = 原记录状态_In;
  
    --查询当前行记录的对应ID
    Select Id
    Into v_收发id
    From 药品收发记录
    Where 单据 = 7 And No = No_In And 序号 = 序号_In And 记录状态 = 原记录状态_In;
  
    --更新库存信息 领用冲销相当于入库 
    Zl_药品库存_Update(v_收发id, 3, 0);
    
    Zl_未审药品记录_Delete(v_收发id);

    --科室药品留存处理 
    If v_发药方式 = 1 Then
      Update 药品留存
      Set 可用数量 = Nvl(可用数量, 0) + 冲销数量_In, 实际数量 = Nvl(实际数量, 0) + 冲销数量_In, 实际金额 = Nvl(实际金额, 0) + v_零售金额
      Where 期间 = To_Char(Sysdate, Decode(v_按月留存领用, '1', 'yyyymm', 'yyyy')) And 科室id = v_对方部门id And 库房id = v_库房id And
            药品id = 药品id_In;
      --将金额和数量等于0的记录删除掉 
      Delete From 药品留存 Where Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0;
    End If;
  
    --处理调价后冲销 
    Zl_药品收发记录_调价修正(v_收发id);
  Else
    --正常冲销方式，产生冲销记录，填写审核人、审核日期，更新库存记录      
    If 行次_In = 1 Then
      Update 药品收发记录
      Set 记录状态 = Decode(原记录状态_In, 1, 3, 原记录状态_In + 3)
      Where No = No_In And 单据 = 7 And 记录状态 = 原记录状态_In;
      If Sql%Rowcount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --主要针对原不分批现在分批的药品，不能对其冲销
    Begin
      Select Distinct '(' || i.编码 || ')' || Nvl(n.名称, i.名称) As 药品信息
      Into v_Druginf
      From 药品收发记录 a, 药品规格 b, 收费项目目录 i, 收费项目别名 n
      Where a.药品id = b.药品id And a.药品id = i.Id And a.药品id = n.收费细目id(+) And n.性质(+) = 3 And a.No = No_In And a.单据 = 7 And
            Mod(a.记录状态, 3) = 0 And Nvl(a.批次, 0) = 0 And a.药品id + 0 = 药品id_In And
            ((Nvl(b.药库分批, 0) = 1 And
            a.库房id Not In (Select 部门id From 部门性质说明 Where (工作性质 Like '%药房') Or (工作性质 Like '制剂室'))) Or
            Nvl(b.药房分批, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(实际数量) As 剩余数量, Sum(成本金额) As 剩余成本金额, Sum(零售金额) As 剩余零售金额, 库房id, 对方部门id, 入出类别id, 入出系数, 批次, 产地, 原产地, 批号, 效期, 成本价,
           扣率, 零售价, 摘要, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期
    Into v_剩余数量, v_剩余成本金额, v_剩余零售金额, v_库房id, v_对方部门id, v_入出类别id, v_入出系数, v_批次, v_产地, v_原产地, v_批号, v_效期, v_成本价, v_扣率, v_零售价,
         v_摘要, v_领用人, v_批准文号, v_发药方式, n_上次供应商id, d_上次生产日期
    From 药品收发记录
    Where No = No_In And 单据 = 7 And 药品id = 药品id_In And 序号 = 序号_In
    Group By 库房id, 对方部门id, 入出类别id, 入出系数, 批次, 产地, 原产地, 批号, 效期, 成本价, 扣率, 零售价, 摘要, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期;
  
    --冲销数量大于剩余数量，不允许 
    If v_剩余数量 < 冲销数量_In Then
      Raise Err_Isnonum;
    End If;
  
    v_成本金额 := Round(冲销数量_In / v_剩余数量 * v_剩余成本金额, Intdigit);
    v_零售金额 := Round(冲销数量_In / v_剩余数量 * v_剩余零售金额, Intdigit);
    v_差价     := v_零售金额 - v_成本金额;
  
    Select 药品收发记录_Id.Nextval Into v_收发id From Dual;
    Insert Into 药品收发记录
      (Id, 记录状态, 单据, No, 序号, 库房id, 对方部门id, 入出类别id, 入出系数, 药品id, 批次, 产地, 原产地, 批号, 效期, 填写数量, 实际数量, 成本价, 成本金额, 零售价, 零售金额, 差价, 摘要,
       填制人, 填制日期, 审核人, 审核日期, 领用人, 批准文号, 发药方式, 供药单位id, 生产日期, 扣率, 冲销原因)
    Values
      (v_收发id, Decode(原记录状态_In, 1, 2, 原记录状态_In + 2), 7, No_In, 序号_In, v_库房id, v_对方部门id, v_入出类别id, v_入出系数, 药品id_In, v_批次,
       v_产地, v_原产地, v_批号, v_效期, -冲销数量_In, -冲销数量_In, v_成本价, -v_成本金额, v_零售价, -v_零售金额, -v_差价, v_摘要, 填制人_In, 填制日期_In, 填制人_In,
       填制日期_In, v_领用人, v_批准文号, v_发药方式, n_上次供应商id, d_上次生产日期, v_扣率, 冲销原因_In);
    
    --更新库存信息 领用冲销相当于入库 
    Zl_药品库存_Update(v_收发id, 3, 0);
    
    --科室药品留存处理 
    If v_发药方式 = 1 Then
      Update 药品留存
      Set 可用数量 = Nvl(可用数量, 0) + 冲销数量_In, 实际数量 = Nvl(实际数量, 0) + 冲销数量_In, 实际金额 = Nvl(实际金额, 0) + v_零售金额
      Where 期间 = To_Char(Sysdate, Decode(v_按月留存领用, '1', 'yyyymm', 'yyyy')) And 科室id = v_对方部门id And 库房id = v_库房id And
            药品id = 药品id_In;
      --将金额和数量等于0的记录删除掉 
      Delete From 药品留存 Where Nvl(实际数量, 0) = 0 And Nvl(实际金额, 0) = 0;
    End If;
  
    --处理调价后冲销 
    Zl_药品收发记录_调价修正(v_收发id);
  End If;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]该单据已经被他人冲销！[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102,
                            '[ZLSOFT]该单据中包含有一条原来不分批，现在分批的药品[' || v_Druginf || ']，不能冲销！[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_药品领用_Strike;
/

--119442:廖思奇,2018-08-22,签名后更新完成时间
CREATE OR REPLACE Procedure ZL_影像报告签名_保存(
	报告文件ID_In  In   电子病历内容.文件ID%Type,
    开始版_In  In       电子病历内容.开始版%Type,
    终止版_In  In       电子病历内容.终止版%Type,
    对象属性_In In      电子病历内容.对象属性%Type,
    姓名_In In          电子病历内容.内容文本%Type,
    前置文字_In In      电子病历内容.要素名称%Type,
    时间戳_In  In       电子病历内容.要素单位%Type,
    签名级别_In In      电子病历内容.要素表示%Type,
    签名信息_In In      电子病历内容.要素值域%Type
) Is
	 n_Nextid     电子病历内容.Id%Type;
     n_序号       电子病历内容.对象序号%Type;
     n_对象标记   电子病历内容.对象标记%Type;
Begin
     Select max(对象序号) +1 Into n_序号 From 电子病历内容 Where 文件ID = 报告文件ID_In;
     Select nvl(Max(对象标记),0)+1 Into n_对象标记 From 电子病历内容 Where 文件ID = 报告文件ID_In And 对象类型=8;

     Select 电子病历内容_Id.Nextval Into n_Nextid From Dual;
     Insert Into 电子病历内容(ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性,
            内容行次, 内容文本, 是否换行, 定义提纲id, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域, 要素名称,
            要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域)
          Values (n_Nextid ,报告文件ID_In,开始版_In,终止版_In,Null,n_序号,8,n_对象标记,1,对象属性_In,Null,姓名_In,
                 0,Null,Null,Null,Null,Null,Null,前置文字_In,1,50,0,时间戳_In,签名级别_In,0,签名信息_In);
     If n_对象标记=1 Then
        Update 电子病历记录 Set 完成时间 = Sysdate ,签名级别 = 签名级别_In Where id = 报告文件ID_In;
     Else
        Update 电子病历记录 Set 签名级别 = 签名级别_In Where id = 报告文件ID_In;
     End If;
	 Update 电子病历记录 Set 完成时间 = Sysdate, 签名级别 = 签名级别_In Where ID = 报告文件id_In;
Exception
	When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm);
End ZL_影像报告签名_保存;
/

--130471:余伟节,2018-08-22,预约安排
CREATE OR REPLACE Procedure Zl_Third_Outpatireg
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
                       Null, Null, Null, r_Pati.险类, v_人员编号, v_人员姓名, 0, Null, n_病区id, 0, Null, Null, Null, Null, Null,
                       Null, Null, n_挂号id);
    End If;
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

--130471:余伟节,2018-08-21,预约安排查询
Create Or Replace Procedure Zl_Third_Patiinfo_Update
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --功能：用于预约中心更新病人信息
  --入参：xml_in
  --<IN>
  -- <REGID>1162695</REGID>   --挂号id
  -- <PATIID>5</PATIID>     --病人ID
  -- <HOME_TEL>3</HOME_TEL>   --家庭电话
  -- <CONTACT_TEL></CONTACT_TEL> --联系人电话
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
  n_挂号id     病人挂号记录.Id%Type;
  n_病人id     病人信息.病人id%Type;
  v_家庭电话   病人信息.家庭电话%Type;
  v_联系人电话 病人信息.联系人电话%Type;

  v_Error Varchar2(2000);

  Err_Custom Exception;
Begin
  Select Extractvalue(Value(A), 'IN/REGID') As 挂号id, Extractvalue(Value(A), 'IN/PATIID') As 病人id,
         Extractvalue(Value(A), 'IN/HOME_TEL') As 家庭电话, Extractvalue(Value(A), 'IN/CONTACT_TEL') As 联系人电话
  Into n_挂号id, n_病人id, v_家庭电话, v_联系人电话
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If n_病人id = 0 Then
    v_Error := '病人ID不允许为空!';
    Raise Err_Custom;
  End If;
  Update 病人信息 Set 家庭电话 = v_家庭电话, 联系人电话 = v_联系人电话 Where 病人id = n_病人id;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Err_Custom Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || v_Error || '</MSG></ERROR></OUTPUT>');
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Patiinfo_Update;
/

--130469:陈龙,2018-08-17,输血申请完成锚点(Zlhis_Blood_010)
Create Or Replace Procedure Zl_医嘱审核管理_Audit
(
  医嘱id_In   病人医嘱状态.医嘱id%Type,
  结果_In     Number,
  操作人员_In 病人医嘱状态.操作人员%Type,
  操作时间_In 病人医嘱状态.操作时间%Type,
  操作说明_In 病人医嘱状态.操作说明%Type := Null,
  审核对象_In Number := 1 --1=手术医嘱，2=输血医嘱 
) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(200);
  n_审核状态 Number;
  n_Count    Number;
  n_结果     Number;
Begin
  Select Nvl(Max(审核状态), 0), Count(1) Into n_审核状态, n_Count From 病人医嘱记录 Where Id = 医嘱id_In;
  If n_审核状态 Not In (1, 7) And n_Count <> 0 Then
    v_Err_Msg := '有医嘱已经审核或不需审核,请查证。';
    Raise Err_Item;
  Elsif n_Count = 0 Then
    v_Err_Msg := '有医嘱已经删除,请查证。';
    Raise Err_Item;
  End If;

  Update 病人医嘱记录 Set 审核状态 = 结果_In + 1 Where Id = 医嘱id_In Or 相关id = 医嘱id_In;
  If 审核对象_In = 2 And 结果_In = 3 Then
    --启用血库系统时，特殊处理 
    n_结果 := 11;
  Elsif 审核对象_In = 2 And 结果_In = 6 Then
    n_结果 := 18;
  Else
    n_结果 := 结果_In + 10;
  End If;
  Insert Into 病人医嘱状态
    (医嘱id, 操作类型, 操作人员, 操作时间, 操作说明)
    Select Id, n_结果, 操作人员_In, 操作时间_In, 操作说明_In
    From 病人医嘱记录
    Where Id = 医嘱id_In Or 相关id = 医嘱id_In;
  --输血医嘱审核完成，抛出锚点
  If 审核对象_In = 2 And n_结果 = 11 Then
    EXECUTE IMMEDIATE 'b_Message_Blood.Zlhis_Blood_010(:1)' USING 医嘱id_In;
  End If;	
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_医嘱审核管理_Audit;
/




------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.35.90.0026' Where 编号=&n_System;
Commit;
