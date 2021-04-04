--在ZLHIS库执行（修改zlsol的密码，ip，实例名[SERVICE_NAME]）
create database link ZLSOL_DBL  connect to ZLSOL identified by ZLSOL_PASSWORD  using '(DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.0.60)(PORT = 1521))
    (CONNECT_DATA =
      (SERVICE_NAME = orcl12)
    )
  )';


--导入在院病人数据(修改病区ID参数7748)
insert into Sol_Inf_Puerpera@ZLSOL_DBL( Pid, Tid, Name, Old, Bedno, Pno, Diagnosis, Status, Outtime) 
Select b.病人id, b.主页id, a.姓名, a.年龄, a.当前床号, a.住院号, c.诊断描述, 1 As Status,To_Date('3000-01-01', 'yyyy-mm-dd')
From 病人信息 A, 在院病人 B, 病人诊断记录 C
Where a.病人id = b.病人id And a.主页id = b.主页id And b.病区id = 7748 And b.病人id = c.病人id(+) And b.主页id = c.主页id(+) And c.记录来源(+) = 3 And
    c.诊断类型(+) = 2 And c.诊断次序(+) = 1

CREATE OR REPLACE Procedure Zl_Apex_产妇状态同步
(
  病人id_In   病人变动记录.病人id%Type,
  主页id_In   病人变动记录.主页id%Type,
  科室id_In   病人变动记录.科室id%Type,
  操作类型_In Number, --1.入科，2.撤销入科，3.出院,4-撤销出院,5换床 6.撤销换床,7.入院诊断更新,8,删除入院诊断
  床号_in     病人变动记录.床号%Type,
  出院时间_In 病人变动记录.终止时间%Type := Null
) As
  n_Count Number(5);
  n_Mid   Number(18);
  v_bedno varchar(10);
Begin
  Select Count(1) Into n_Count From 部门性质说明 Where 部门id = 科室id_In And 工作性质 = '产科';

  If n_Count = 1 Then
    --1.入科
    If 操作类型_In = 1 Then
      Select Count(1) Into n_Count From Sol_Inf_Puerpera@ZLSOL_DBL Where Pid = 病人id_In And Tid = 主页id_In;

      If n_Count = 0 Then
        Insert Into Sol_Inf_Puerpera@ZLSOL_DBL
          (Pid, Tid, Name, Old, Bedno, Pno, Diagnosis,status,outtime)
          Select a.病人id, a.主页id, a.姓名, a.年龄, a.当前床号, a.住院号, c.诊断描述,0,To_Date('3000-01-01', 'yyyy-mm-dd')
          From 病人信息 A, 病人诊断记录 C
          Where a.病人id = 病人id_In And a.主页id = 主页id_In And a.病人id = c.病人id(+) And a.主页id = c.主页id(+) And c.记录来源(+) = 3 And
                c.诊断类型(+) = 2 And c.诊断次序(+) = 1;
      End If;
    Elsif 操作类型_In = 2 Then
      Select Nvl(Max(Mid), 0)
      Into n_Mid
      From Sol_Inf_Puerpera@ZLSOL_DBL A
      Where a.Pid = 病人id_In And a.Tid = 主页id_In And a.Status = 0;

      --未入房，并且未填写待产记录、临产记录，则删除,待产记录可能已存在，由病区填写
      If n_Mid > 0 Then
        /*Select Count(1) Into n_Count From Sol_Rs_Expectant@ZLSOL_DBL Where Mid = n_Mid;
        If n_Count = 0 Then
          Select Count(1) Into n_Count From Sol_Rs_Birth@ZLSOL_DBL Where Mid = n_Mid;
          If n_Count = 0 Then*/
            Delete Sol_Inf_Puerpera@ZLSOL_DBL Where Mid = n_Mid;
         /* End If;
        End If;*/
      End If;

    Elsif 操作类型_In = 3 Then
      --出院
      Update Sol_Inf_Puerpera@ZLSOL_DBL A
      Set Status = 3, Outtime = 出院时间_In
      Where a.Pid = 病人id_In And a.Tid = 主页id_In;

    Elsif 操作类型_In = 4 Then
      --撤销出院
      Update Sol_Inf_Puerpera@ZLSOL_DBL A
      Set Status = 2, Outtime = To_Date('3000-01-01', 'yyyy-mm-dd')
      Where a.Pid = 病人id_In And a.Tid = 主页id_In And a.Status = 3;
    Elsif 操作类型_In = 5 Then
      --换床
      Update Sol_Inf_Puerpera@ZLSOL_DBL A
      Set  bedno=nvl(床号_in,'家庭病床')
      Where a.Pid = 病人id_In And a.Tid = 主页id_In;
    Elsif 操作类型_In = 6 Then
      --撤销换床
      select 当前床号 into v_bedno from 病人信息 where 病人id = 病人id_In And 主页id = 主页id_In;
      Update Sol_Inf_Puerpera@ZLSOL_DBL A
      Set bedno=v_bedno
      Where a.Pid = 病人id_In And a.Tid = 主页id_In ;
    End If;
  End If;
End Zl_Apex_产妇状态同步;
/
CREATE OR REPLACE Trigger t_Apex_产妇状态同步
  After Insert Or Delete Or Update On 病人变动记录
  For Each Row
Declare

  v_Err_Msg Varchar2(255);
  Err_Item Exception;  ----1.入科 2。撤销入科 3.出院 4.撤销出院 5、换床 6.撤销换床
Begin

  If Inserting Then
    --入院入科
    If :New.开始原因 =1 And Nvl(:new.床号,0) <>0Then
      Zl_Apex_产妇状态同步(:New.病人id, :New.主页id, :New.科室id, 1,:new.床号);
    End If;    
    --入科
    If :New.开始原因 In (2, 3) And :New.附加床位 = 0 Then
      Zl_Apex_产妇状态同步(:New.病人id, :New.主页id, :New.科室id, 1,:new.床号);
    End If;
    -----换床
    If :New.开始原因 =4 Then
      Zl_Apex_产妇状态同步(:New.病人id, :New.主页id, :New.科室id, 5,:new.床号);
    End If;
  Elsif Deleting Then
    --撤销入科
    If :Old.开始原因  In (2, 3) Then
      Zl_Apex_产妇状态同步(:Old.病人id, :Old.主页id, :Old.科室id, 2,:new.床号);
    End If;
    --撤销换床
    If :Old.开始原因 = 4 Then
      Zl_Apex_产妇状态同步(:Old.病人id, :Old.主页id, :Old.科室id, 6,:new.床号);
    End If;
    --出院
  Elsif Updating Then
    If :New.终止原因 = 1 And :Old.终止时间 Is Null And :Old.附加床位 = 0 Then
      Zl_Apex_产妇状态同步(:New.病人id, :New.主页id, :New.科室id, 3,:new.床号, :New.终止时间);
    End If;
    --撤销出院
    If :Old.终止原因 = 1 And :New.终止时间 Is Null Then
      Zl_Apex_产妇状态同步(:New.病人id, :New.主页id, :New.科室id, 4,:new.床号);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End t_Apex_产妇状态同步;
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
  From 病人诊断记录 a, 在院病人 b, 部门性质说明 c
  Where a.病人id = b.病人id And a.主页id = b.主页id And b.科室id = c.部门id And c.工作性质 = '产科' And a.病人id = 病人id_In And
        a.主页id = 主页id_In;
  If n_Count = 1 Then
    If 操作类型_In = 1 Then
      Update Sol_Inf_Puerpera@ZLSOL_DBL Set Diagnosis = 诊断描述_In Where Pid = 病人id_In And Tid = 主页id_In;
    Elsif 操作类型_In = 2 Then
      Update Sol_Inf_Puerpera@ZLSOL_DBL Set Diagnosis = Null Where Pid = 病人id_In And Tid = 主页id_In;
    End If;
  End If;
End Zl_Apex_产妇状态同步_诊断;
/
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