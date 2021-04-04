--用zlhis用户登录助产士的数据库，
--创建dblink连接
create database link ZLSOL_DBL  connect to ZLSOL identified by &zlsol连接助产士库的密码  using '(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = &助产士库IP)(PORT = &助产士库端口))
    (CONNECT_DATA =
      (SERVICE_NAME = &助产士库实例名)
    )
  )';
--导入在院的产科产妇到助产士工作站中
Insert into Sol_Inf_Puerpera@ZLSOL_DBL( Pid, Tid, Name, Old, Bedno, Pno, Diagnosis, Status, Outtime) 
Select b.病人id, b.主页id, a.姓名, a.年龄, a.当前床号, a.住院号, c.诊断描述, 0 As Status,To_Date('3000-01-01', 'yyyy-mm-dd')
From 病人信息 A, 在院病人 B, 病人诊断记录 C
Where a.病人id = b.病人id And a.主页id = b.主页id And b.病区id = &产科id And b.病人id = c.病人id(+) And b.主页id = c.主页id(+)；
--产妇同步触发器及过程
CREATE OR REPLACE Trigger t_Apex_产妇状态同步
  After Insert Or Delete Or Update On 病人变动记录
  For Each Row
Declare
  n_科室id  病人变动记录.科室id%Type;
  v_Err_Msg Varchar2(255);
  Err_Item Exception; --1.入科、入院入科、转入产科 2.撤销入科 3.出院 4.撤销出院 5、换床 6.撤销换床
Begin
  If Inserting Then
    --入院入科
    If :New.开始原因 = 1 And Nvl(:New.床号, 0) <> 0 Then
      Zl_Apex_产妇状态同步(:New.病人id, :New.主页id, :New.科室id, 1, :New.床号);
    End If ;
    --入科，转入产科
    If :New.开始原因 In (2, 3) And :New.附加床位 = 0 Then
      Zl_Apex_产妇状态同步(:New.病人id, :New.主页id, :New.科室id, 1, :New.床号);
    End If;
    --换床
    If :New.开始原因 = 4 Then
      Zl_Apex_产妇状态同步(:New.病人id, :New.主页id, :New.科室id, 5, :New.床号);
    End If;
  Elsif Deleting Then
    --撤销入科
    If :Old.开始原因 In (2, 3) Then
      Zl_Apex_产妇状态同步(:Old.病人id, :Old.主页id, :Old.科室id, 2);
    End If;
    --撤销换床
    If :Old.开始原因 = 4 Then
      Zl_Apex_产妇状态同步(:Old.病人id, :Old.主页id, :Old.科室id, 6, :New.床号);
    End If;
  Elsif Updating Then
    --出院
    If :New.终止原因 = 1 And :Old.终止时间 Is Null And :Old.附加床位 = 0 Then
      Zl_Apex_产妇状态同步(:New.病人id, :New.主页id, :New.科室id, 3, :New.床号, :New.终止时间);
    End If;
    --撤销出院
    If :Old.终止原因 = 1 And :New.终止时间 Is Null Then
      Zl_Apex_产妇状态同步(:New.病人id, :New.主页id, :New.科室id, 4);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End t_Apex_产妇状态同步;

/


CREATE OR REPLACE Procedure Zl_Apex_产妇状态同步
(
  病人id_In   病人变动记录.病人id%Type,
  主页id_In   病人变动记录.主页id%Type,
  科室id_In   病人变动记录.科室id%Type,
  操作类型_In Number, --1.入科、入院入科、转入产科 2.撤销入科 3.出院 4.撤销出院 5、换床 6.撤销换床
  床号_In     病人变动记录.床号%Type:= Null,
  出院时间_In 病人变动记录.终止时间%Type := Null
) As
  n_Count   Number(5);
  v_Bedno   Varchar(10);
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  If 科室id_In In (205) Then
    --1.入科、转科入住、入院入科
    If 操作类型_In = 1 Then
      Select Count(1) Into n_Count From Sol_Inf_Puerpera@Zlsol_Dbl Where Pid = 病人id_In And Tid = 主页id_In;

      If n_Count = 0 Then
        Insert Into Sol_Inf_Puerpera@Zlsol_Dbl
          (Pid, Tid, Name, Old, Bedno, Pno, Diagnosis, Status, Outtime)
          Select a.病人id, a.主页id, a.姓名, a.年龄, a.当前床号, a.住院号, c.诊断描述, 0, To_Date('3000-01-01', 'yyyy-mm-dd')
          From 病人信息 a, 病人诊断记录 c
          Where a.病人id = 病人id_In And a.主页id = 主页id_In And a.病人id = c.病人id(+) And a.主页id = c.主页id(+);
      Else
        --转科时已存在病人
        Update Sol_Inf_Puerpera@Zlsol_Dbl a Set Status = 0 Where a.Pid = 病人id_In And a.Tid = 主页id_In;
      End If;
    --2.撤销入科
    Elsif 操作类型_In = 2 Then
      Update Sol_Inf_Puerpera@Zlsol_Dbl a Set Status = 4 Where a.Pid = 病人id_In And a.Tid = 主页id_In;
    --3.出院
    Elsif 操作类型_In = 3 Then
      Update Sol_Inf_Puerpera@Zlsol_Dbl a
      Set Status = 3, Outtime = 出院时间_In
      Where a.Pid = 病人id_In And a.Tid = 主页id_In;
    --4.撤销出院
    Elsif 操作类型_In = 4 Then
      Update Sol_Inf_Puerpera@Zlsol_Dbl a
      Set Status = 2, Outtime = To_Date('3000-01-01', 'yyyy-mm-dd')
      Where a.Pid = 病人id_In And a.Tid = 主页id_In And a.Status = 3;
    --5.换床
    Elsif 操作类型_In = 5 Then
      Update Sol_Inf_Puerpera@Zlsol_Dbl a
      Set Bedno = Nvl(床号_In, '家庭病床')
      Where a.Pid = 病人id_In And a.Tid = 主页id_In;
    --6.撤销换床
    Elsif 操作类型_In = 6 Then
      Select 当前床号 Into v_Bedno From 病人信息 Where 病人id = 病人id_In And 主页id = 主页id_In;
      Update Sol_Inf_Puerpera@Zlsol_Dbl a Set Bedno = v_Bedno Where a.Pid = 病人id_In And a.Tid = 主页id_In;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Apex_产妇状态同步;




/
--病人诊断同步
CREATE OR REPLACE Trigger t_Apex_产妇状态同步_诊断
  After Insert Or Delete Or Update On 病人诊断记录
  For Each Row
Declare

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
Begin
  --新增和修改
  If Inserting Or Updating Then
    If :New.记录来源 In (2, 3) And :New.诊断类型 = 2 And :New.诊断次序 = 1 Then
      Zl_Apex_产妇状态同步_诊断(:New.病人id, :New.主页id, 1, :New.诊断描述);
    End If;
    --删除
  Elsif Deleting Then
    If :Old.记录来源 In (2, 3) And :New.诊断类型 = 2 And :New.诊断次序 = 1 Then
      Zl_Apex_产妇状态同步_诊断(:Old.病人id, :Old.主页id, 2);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End t_Apex_产妇状态同步_诊断;

/





CREATE OR REPLACE Procedure Zl_Apex_产妇状态同步_诊断
(
  病人id_In   病人诊断记录.病人id%Type,
  主页id_In   病人诊断记录.主页id%Type,
  操作类型_In Number, --1.修改诊断；2.删除诊断
  诊断描述_In 病人诊断记录.诊断描述%Type := Null
) As
  n_Count Number(5);
Begin
  Select Count(1)
  Into n_Count
  From 在院病人 B, 部门性质说明 C
  Where b.科室id = c.部门id And c.工作性质 = '产科' And b.病人id = 病人id_In And b.主页id = 主页id_In;
  If n_Count = 1 Then
    If 操作类型_In = 1 Then
      Update Sol_Inf_Puerpera@Zlsol_Dbl Set Diagnosis = 诊断描述_In Where Pid = 病人id_In And Tid = 主页id_In;
    Elsif 操作类型_In = 2 Then
      update Sol_Inf_Puerpera@Zlsol_Dbl Set Diagnosis = Null Where Pid = 病人id_In And Tid = 主页id_In;
    End If;
  End If;
End Zl_Apex_产妇状态同步_诊断;
/





--助产士用户同步触发器
--注意：zl_人员表_delete 过程中需要在delete from...后面加一句 Delete From Sol_User@Zlsol_Dbl Where Code = v_User;
CREATE OR REPLACE Trigger t_Apex_人员变动
  After Insert Or Delete Or Update On 上机人员表
  For Each Row
Declare
  v_Name    Varchar2(20);
  n_Count   Number(5);
  v_Err_Msg Varchar2(255);
  Err_Item Exception; ----1.新增人员；2.删除人员；3.修改人员的实质是先删除再新增
Begin
  If Inserting Then
    --新增人员
    Select Max(Distinct b.姓名) 姓名
    Into v_Name
    From 部门人员 A, 人员表 B
    Where a.人员id = b.Id And a.人员id = :New.人员id And a.部门id = &需创建助产士用户的科室ID;
    If v_Name Is Not Null Then
      Insert Into Sol_User@Zlsol_Dbl (Code, Name, State) Values (:New.用户名, v_Name, 1);
    End If;
    --删除人员
  Elsif Deleting Then
    Select Count(*) Into n_Count From 部门人员 A Where a.人员id = :Old.人员id And a.部门id = &需创建助产士用户的科室ID;
    If n_Count > 0 Then
      Delete From Sol_User@Zlsol_Dbl Where Code = :Old.用户名;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End t_Apex_人员变动;
/
CREATE OR REPLACE Trigger t_Apex_人员调整
  After Insert Or Delete Or Update On 人员表
  For Each Row
Declare
  v_Code    Varchar2(20);
Begin
  If Updating Then    
    Select max(Distinct b.用户名) 用户名
    Into v_Code
    From 部门人员 a, 上机人员表 b
    Where a.人员id = b.人员id And a.人员id = :Old.Id And a.部门id = &需创建助产士用户的科室ID;
    If v_Code Is Not Null Then
      --修改姓名
      If :New.姓名 <> :Old.姓名 Then
        Update Sol_User@Zlsol_Dbl Set Name = :New.姓名 Where Code = v_Code;
      End If;
      --启用、停用用户
      If :New.撤档时间 = To_Date('3000-1-1', 'yyyy-mm-dd') Then
        Update Sol_User@Zlsol_Dbl Set State = 1 Where Code = v_Code;
      Else
        Update Sol_User@Zlsol_Dbl Set State = 0 Where Code = v_Code;
      End If;
    End If;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End t_Apex_人员调整;
/