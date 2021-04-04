----10.35.20---》10.35.30
--100680:刘硕,2016-10-13,清除本机界面异常
Create Or Replace Procedure ZLTOOLS.Zl_zluserparas_Clear
(
  用户名_In In zluserparas.用户名%Type,
  机器名_In In zluserparas.机器名%Type
) Is
Begin
  Delete From zluserparas Where 用户名=用户名_In or 机器名=机器名_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_zluserparas_Clear;
/
--100258:余智勇,2016-09-05,解决zlReports锁表问题
Alter Table zlTools.zlReports Add (执行人员 Varchar2(20), 最后执行时间 Date);
Alter Table zlTools.Zlrptrunhistory Add (执行人员 Varchar2(20));

Begin  
  For r_Alter In (Select 'alter index ' || Index_Name || ' initrans 20' Line_
                  From All_Indexes
                  Where Table_Name = 'ZLREPORTS' And ini_trans < 20) Loop
    Execute Immediate r_Alter.Line_;
  End Loop;
End;
/

Alter Table zlTools.zlReports Move;
Begin  
  For r_Alter In (Select 'alter index ' || Index_Name || ' rebuild' Line_
                  From All_Indexes
                  Where Table_Name = 'ZLREPORTS') Loop
    Execute Immediate r_Alter.Line_;
  End Loop;
End;
/

--100258:余智勇,2016-09-05,解决zlReports锁表问题
CREATE OR REPLACE Procedure zlTools.Zl_Rptrun_Update
(
  Id_In       In Zlreports.Id%Type,
  执行人员_In In Zlreports.执行人员%Type
) Is
  Pragma Autonomous_Transaction;
  n_Count Number(1);
Begin
  Select Nvl(Count(1), 0)
  Into n_Count
  From zlReports
  Where ID = Id_In And (最后执行时间 < Sysdate - 5 / 24 / 60 Or 最后执行时间 Is Null);

  If n_Count > 0 Then
    Update zlReports Set 执行人员 = 执行人员_In, 最后执行时间 = Sysdate Where ID = Id_In;
    Commit;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Rptrun_Update;
/

--100258:余智勇,2016-09-05,解决zlReports锁表问题
Create Or Replace Procedure zlTools.Zl_Rptrunhistory_Update
(
  报表id_In       In Zlreports.Id%Type,
  执行人员_In     In Zlrptrunhistory.执行人员%Type,
  执行开始时间_In In Zlrptrunhistory.执行开始时间%Type,
  执行结束时间_In In Zlrptrunhistory.执行结束时间%Type
) Is
Begin
  Insert Into Zlrptrunhistory
    (ID, 报表id, 执行人员, 执行开始时间, 执行结束时间)
  Values
    (Zlrptrunhistory_Id.Nextval, 报表id_In, 执行人员_In, 执行开始时间_In, 执行结束时间_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Rptrunhistory_Update;
/

--100994:张永康,2016-09-30,玉溪医院历史数据转出测试集中修改
Alter Table zltools.Zldatamovelog modify 当前进度 varchar2(200);

--100994:张永康,2016-09-30,玉溪医院历史数据转出测试集中修改
Create Table zltools.zlBakTableIndex(
	系统 NUMBER(5) ,
	表名 Varchar2(30),
	索引名 VARCHAR2(30));
Alter Table zltools.zlBakTableIndex Add Constraint zlBakTableIndex_PK Primary Key (系统,表名,索引名) USING INDEX PCTFREE 5;
Alter Table zltools.zlBakTableIndex Add Constraint zlBakTableIndex_FK_系统 Foreign Key (系统,表名) References zlBakTables(系统,表名) On Delete Cascade;

--100412:刘硕,2016-10-26,自动升级改进
Alter Table zltools.zlFilesUpgrade Drop Constraint zlFilesUpgrade_CK_自动注册;
Alter Table zltools.ZLCLIENTS Drop Constraint ZLCLIENTS_CK_预升完成;
Alter Table zltools.zlFilesUpgrade Modify 文件名 Varchar2(100);
--100412:刘硕,2016-10-26,自动升级改进
create table ZLTOOLS.ZLKillProcess
(
序号     number(5),
名称     varchar2(50),
类型     number(1),--0-进程，1-服务
描述     varchar2(200)
);
alter table  ZLTOOLS.ZLKillProcess add constraint ZLKillProcess_UQ_名称 unique(名称) using index;


--101981:刘硕,2016-10-27,清空预升级时点报错
Create Or Replace Procedure zltools.Zl_Zlclients_Control
(
  n_Mode_In       Number,
  v_工作站_In     Zlclients.工作站%Type := Null,
  v_Ip_In         Zlclients.Ip%Type := Null,
  n_升级标志_In   Zlclients.升级标志%Type := Null,
  n_升级服务器_In Zlclients.升级服务器%Type := Null,
  d_预升时点_In   Zlclients.预升时点%Type := Null,
  n_预升完成_In   Zlclients.预升完成%Type := Null,
  n_Ftp服务器_In  Zlclients.Ftp服务器%Type := Null,
  n_收集标志_In   Zlclients.收集标志%Type := Null,
  n_禁止使用_In   Zlclients.禁止使用%Type := Null,
  v_说明_In       Zlclients.说明%Type := Null
  --对客户端进行控制
  --N_Mode_In：0-禁用或启用客户端(IP做为主要条件）,1-预升级设置,2 -升级信息保存(IP做为主要条件）
  --3-取消预升级标志,4-将所有站点设置为升级,5-部件搜集（设置搜集标志）,6-重置升级状态
) Is
  v_Timeset Varchar2(300);
  v_Err     Varchar2(500);
  Err_Custom Exception;
Begin
  --0-禁用或启用客户端(IP做为主要条件）
  If n_Mode_In = 0 Then
    If v_工作站_In Is Not Null Then
      Update zlClients Set 禁止使用 = n_禁止使用_In Where Ip = v_Ip_In;
    End If;
    --1-预升级设置,不需要传其他参数
  Elsif n_Mode_In = 1 Then
    Select Max(内容) Into v_Timeset From zlRegInfo Where 项目 = '客户端预升级时间点';
    If v_Timeset Is Not Null Then
      For r_Ip In (Select To_Date(Today || ' ' || Date_d, 'yyyy-mm-dd HH24:mi:ss') 预升时点, 工作站, Ip
                   From (Select 工作站, Ip, Rownum Rn_c From zlClients) A,
                        (Select To_Char(Sysdate, 'yyyy-mm-dd') Today, Column_Value Date_d, Rownum Rn_d, Count(1) Over() Sn
                          From Table(f_Str2list(v_Timeset, ','))) B
                   Where Mod(a.Rn_c, Sn) + 1 = Rn_d) Loop
      
        Update zlClients Set 预升时点 = r_Ip.预升时点 Where 工作站 = r_Ip.工作站 And Ip = r_Ip.Ip;
      End Loop;
    Else
      Update zlClients Set 预升时点 = NuLL;
    End If;
    --2 -升级信息保存(IP做为主要条件）
  Elsif n_Mode_In = 2 Then
    If n_Ftp服务器_In Is Null Then
      Update zlClients
      Set 升级标志 = n_升级标志_In, 升级服务器 = n_升级服务器_In, 预升时点 = d_预升时点_In, 预升完成 = n_预升完成_In
      Where Ip = v_Ip_In;
    
    Else
      Update zlClients
      Set 升级标志 = n_升级标志_In, Ftp服务器 = n_Ftp服务器_In, 预升时点 = d_预升时点_In, 预升完成 = n_预升完成_In
      Where Ip = v_Ip_In;
    End If;
    --3-取消预升级标志
  Elsif n_Mode_In = 3 Then
    Update zlClients Set 预升完成 = n_预升完成_In;
    --4-将所有站点设置为升级
  Elsif n_Mode_In = 4 Then
    Update zlClients Set 升级标志 = n_升级标志_In;
    --5-部件搜集（设置搜集标志）
  Elsif n_Mode_In = 5 Then
    If v_工作站_In Is Null Then
      Update zlClients Set 收集标志 = n_收集标志_In;
    Else
      Update zlClients Set 收集标志 = n_收集标志_In Where 工作站 = v_工作站_In;
    End If;
  Elsif n_Mode_In = 6 Then
    Update zlClients Set 升级情况 = 0 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 7 Then
    --7未升级
    Update zlClients Set 升级情况 = 1 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 8 Then
    --8已升级
    Update zlClients Set 升级情况 = 2 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 9 Then
    --9修改说明
    Update zlClients Set 说明 = v_说明_In Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 10 Then
    --10修改说明和收集标志
    Update zlClients Set 说明 = v_说明_In, 收集标志 = 0 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 11 Then
    --11修改说明和升级标志
    Update zlClients Set 说明 = v_说明_In, 升级标志 = 0 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 12 Then
    --12更改站点的预升级出错状态
    Update zlClients Set 说明 = v_说明_In, 预升完成 = 0 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 13 Then
    --13更改站点的预升级完成状态
    Update zlClients Set 预升完成 = 1 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 14 Then
    --14定时升级
    Update zlClients Set 预升时点 = Null, 预升完成 = Null Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 15 Then
    --15升级成功
    Update zlClients Set 升级情况 = 1 Where 工作站 = v_工作站_In;
  Elsif n_Mode_In = 16 Then
    --16升级失败
    Update zlClients Set 升级情况 = 2 Where 工作站 = v_工作站_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Control;
/
