----10.35.60---》10.35.70
--000000:刘硕,2017-9-11,规范检查
ALTER TABLE ZLTOOLS.ZLUPGRADESERVER Rename CONSTRAINT ZLUPGRADESERVER_PK_编号 to ZLUPGRADESERVER_PK;
alter index  ZLTOOLS.ZLUPGRADESERVER_PK_编号 rename to ZLUPGRADESERVER_PK;
--000000:蒋敏,2017-8-07,修改数据
Delete from ZLTools.zlTables Where 表名='ZLTools.zlTables';
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLTABLES','ZLTOOLSTBS','A1');
--111526:高腾,2017-7-5,日志启停管理
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLLOGCONFIG','ZLTOOLSTBS','A2');
--112138:高腾,2017-8-31,关闭锁定导航台
Insert Into ZLTools.zlOptions(参数号,参数名,参数值,缺省值,参数说明) Values(24, '允许关闭锁定的导航台', '0','0', '控制当导航台被锁定的时候，能否关闭导航台。');
--104763:高腾,2017-9-7,显示停用报表
Insert Into ZLTools.zlParameters(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)
Select zlParameters_ID.Nextval,-Null,-Null,1,-Null,-Null,-Null,A.* From (
Select 部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From zlParameters Where 1 = 0 Union All
Select 0,0,29,'显示停用报表',Null,'0','是否显示已经被停用的报表','0或NUll：不显示停用报表，1：显示停用报表',NULL,NULL,NULL From Dual Union All
Select 部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 From ZLPARAMETERS Where 1 = 0) A;
--000000:刘硕,2017-8-07,无效表删除
Drop table ZLTools.ZLPROCEDURENOTE;
Drop public synonym ZLPROCEDURENOTE;
Drop Procedure zlTools.Zl_zlProcedureNote_Delete;
Drop public synonym Zl_zlProcedureNote_Delete;
Drop Procedure zlTools.Zl_zlProcedureNote_Update;
Drop public synonym Zl_zlProcedureNote_Update;
Create Or Replace Procedure zlTools.Zl_Zlprocedure_Delete
(
  Id_In           In Zlprocedure.Id%Type
) Is
Begin
  Delete zlProcedureText Where 过程ID=Id_In;
  Delete zlProcedure Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlprocedure_Delete;
/
--111911:刘硕,2017-7-28,7z调整附带重整部分客户端升级代码
Drop Procedure Zltools.Zl_Zlupgradeserver_Delete;
Drop public synonym Zl_Zlupgradeserver_Delete;
Drop Procedure Zltools.Zl_Zlupgradeserver_Insert;
Drop public synonym Zl_Zlupgradeserver_Insert;
Drop Procedure Zltools.Zlreginfo_Defaultserver;
Drop public synonym Zlreginfo_Defaultserver;
Alter table ZLTOOLS.ZLUPGRADESERVER drop constraint ZLUPGRADESERVER_UQ_位置;
Create Or Replace Procedure Zltools.Zl_Zlupgradeserver_Update
(
  模式_In     In Number,
  编号_In     In Zlupgradeserver.编号%Type,
  类型_In     In Zlupgradeserver.类型%Type := Null,
  位置_In     In Zlupgradeserver.位置%Type := Null,
  用户名_In   In Zlupgradeserver.用户名%Type := Null,
  密码_In     In Zlupgradeserver.密码%Type := Null,
  端口_In     In Zlupgradeserver.端口%Type := Null,
  是否升级_In In Zlupgradeserver.是否升级%Type := Null,
  是否缺省_In In Zlupgradeserver.是否缺省%Type := Null,
  是否收集_In In Zlupgradeserver.是否收集%Type := Null,
  收集类型_In In Zlupgradeserver.收集类型%Type := Null,
  密码ex_In   In Zlupgradeserver.密码%Type := Null
) Is
  --模式_IN=0-新增，1-修改,11-只修改是否升级，是否缺省，是否收集，收集类型字段 ,2-删除
  v_收集类型 Zlupgradeserver.收集类型%Type;
  n_是否升级 Zlupgradeserver.是否升级%Type;
  n_编号     Zlupgradeserver.编号%Type;
Begin
  --若设置新的缺省，则清除以前的缺省
  If 是否缺省_In = 1 Then
    Update Zlupgradeserver Set 是否缺省 = 0 Where Nvl(是否缺省, 0) = 1;
  End If;
  If 是否收集_In = 1 Then
    Select Max(v_收集类型) Into v_收集类型 From Zlupgradeserver Where Nvl(是否收集, 0) = 1;
    Update Zlupgradeserver Set 是否收集 = 0, 收集类型 = Null Where Nvl(是否收集, 0) = 1;
  End If;
  If Nvl(模式_In, 0) = 0 Or Nvl(编号_In, 0) = 0 Then
    Select Nvl(Max(编号), 0) + 1 Into n_编号 From Zlupgradeserver;
    Insert Into Zlupgradeserver
      (编号, 类型, 位置, 用户名, 密码, 端口, 是否升级, 是否缺省, 是否收集, 收集类型, 批次)
    Values
      (n_编号, 类型_In, 位置_In, 用户名_In, 密码_In, 端口_In, 是否升级_In, 是否缺省_In, 是否收集_In, 收集类型_In, 0);
  Elsif Nvl(模式_In, 0) = 2 Then
    Delete From Zlupgradeserver Where 编号 = 编号_In;
    Update Zlclients Set 升级文件服务器 = Null Where 升级文件服务器 = 编号_In;
  Else
    Select Max(是否升级) Into n_是否升级 From Zlupgradeserver Where 编号 = 编号_In;
    If Nvl(模式_In, 0) = 1 Then
      Update Zlupgradeserver
      Set 类型 = 类型_In, 位置 = 位置_In, 用户名 = 用户名_In, 密码 = 密码_In, 端口 = 端口_In, 是否升级 = 是否升级_In, 是否缺省 = 是否缺省_In, 是否收集 = 是否收集_In,
          收集类型 = 收集类型_In
      Where 编号 = 编号_In;
    Else
      Update Zlupgradeserver
      Set 是否升级 = 是否升级_In, 是否缺省 = 是否缺省_In, 是否收集 = 是否收集_In, 收集类型 = v_收集类型
      Where 编号 = 编号_In;
    End If;
    --升级服务器不再作为升级服务器，则清除升级服务器设置
    If Nvl(n_是否升级, 0) = 1 And Nvl(是否升级_In, 0) = 0 Then
      Update Zlclients Set 升级文件服务器 = Null Where 升级文件服务器 = 编号_In;
    End If;
  End If;
  --自动插入ZLRegINFO中数据，保证和以前兼容
  If Nvl(是否缺省_In, 0) = 1 Then
    Insert Into Zltools.Zlreginfo
      (项目, 内容)
      Select '升级类型', Null From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '升级类型');
    If Nvl(类型_In, 0) = 0 Then
      Insert Into Zltools.Zlreginfo
        (项目, 内容)
        Select '服务器目录0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '服务器目录0')
        Union All
        Select '访问用户0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '访问用户0')
        Union All
        Select '访问密码0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '访问密码0')
        Union All
        Select '服务器目录', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '服务器目录')
        Union All
        Select '访问用户', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '访问用户')
        Union All
        Select '访问密码', '' From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = '访问密码');
      Update Zltools.Zlreginfo Set 内容 = '0' Where 项目 = '升级类型';
      Update Zltools.Zlreginfo Set 内容 = 位置_In Where 项目 = '服务器目录0';
      Update Zltools.Zlreginfo Set 内容 = 用户名_In Where 项目 = '访问用户0';
      Update Zltools.Zlreginfo Set 内容 = 密码ex_In Where 项目 = '访问密码0';
      Update Zltools.Zlreginfo Set 内容 = 位置_In Where 项目 = '服务器目录';
      Update Zltools.Zlreginfo Set 内容 = 用户名_In Where 项目 = '访问用户';
      Update Zltools.Zlreginfo Set 内容 = 密码ex_In Where 项目 = '访问密码';
    Else
      Insert Into Zltools.Zlreginfo
        (项目, 内容)
        Select 'FTP服务器0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP服务器0')
        Union All
        Select 'FTP用户0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP用户0')
        Union All
        Select 'FTP密码0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP密码0')
        Union All
        Select 'FTP端口0', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP端口0')
        Union All
        Select 'FTP服务器', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP服务器')
        Union All
        Select 'FTP用户', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP用户')
        Union All
        Select 'FTP密码', Null
        From Dual
        Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP密码')
        Union All
        Select 'FTP端口', Null From Dual Where Not Exists (Select 1 From Zltools.Zlreginfo Where 项目 = 'FTP端口');
      Update Zltools.Zlreginfo Set 内容 = '1' Where 项目 = '升级类型';
      Update Zltools.Zlreginfo Set 内容 = 位置_In Where 项目 = 'FTP服务器0';
      Update Zltools.Zlreginfo Set 内容 = 用户名_In Where 项目 = 'FTP用户0';
      Update Zltools.Zlreginfo Set 内容 = 密码ex_In Where 项目 = 'FTP密码0';
      Update Zltools.Zlreginfo Set 内容 = 端口_In Where 项目 = 'FTP端口0';
      Update Zltools.Zlreginfo Set 内容 = 位置_In Where 项目 = 'FTP服务器';
      Update Zltools.Zlreginfo Set 内容 = 用户名_In Where 项目 = 'FTP用户';
      Update Zltools.Zlreginfo Set 内容 = 密码ex_In Where 项目 = 'FTP密码';
      Update Zltools.Zlreginfo Set 内容 = 端口_In Where 项目 = 'FTP端口';
    End If;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlupgradeserver_Update;
/
--102323:高腾,2017-7-13,为后台作业管理添加其它时间单位.周.月.季度
Alter Table Zltools.Zlautojobs Add 时间单位 Varchar2(5);

--102323:高腾,2017-7-13,为后台作业管理添加其它时间单位.周.月.季度
CREATE OR REPLACE Procedure zltools.Zl_Jobsubmit
(
  Job_System In Integer,
  Job_Kind   In Integer,
  Job_Odd    In Integer
) Authid Current_User Is
  v_Content   Varchar2(200);
  v_Parameter Varchar2(200);
  v_Paraitem  Varchar2(200);
  v_Starttime Date;
  v_Cyclekeep Integer;
  v_Jobnum    Number := 0;
  v_What      Varchar2(1000);
  v_Nextdate  Date;
  v_Interval  Varchar2(1000);
  v_Timeunit  Varchar2(5);
  v_Week      Varchar2(10);
  v_Day       Varchar2(1);
Begin
  Select 内容, 参数, 执行时间, 间隔时间, 时间单位
  Into v_Content, v_Parameter, v_Starttime, v_Cyclekeep, v_Timeunit
  From Zlautojobs
  Where Nvl(系统, 0) = Nvl(Job_System, 0) And 类型 = Job_Kind And 序号 = Job_Odd;
  v_What := '';

  If Length(v_Parameter) > 0 Then
    Loop
      If Instr(v_Parameter, ';') > 0 Then
        v_Paraitem  := Substr(v_Parameter, 1, Instr(v_Parameter, ';') - 1);
        v_Parameter := Substr(v_Parameter, Instr(v_Parameter, ';') + 1);
      Else
        v_Paraitem := v_Parameter;
      End If;
    
      v_What := v_What || ',' || Substr(v_Paraitem, Instr(v_Paraitem, ',') + 1);
      Exit When Instr(v_Parameter, ';') = 0;
    End Loop;
  End If;

  If Length(v_What) <> 0 Then
    v_What := v_Content || '(' || Substr(v_What, 2) || ');';
  Else
    v_What := v_Content || ';';
  End If;

  If v_Timeunit = '天' Then
    If To_Char(Sysdate, 'HH24:MI:SS') >= To_Char(v_Starttime, 'HH24:MI:SS') Then
      v_Nextdate := v_Starttime + 1;
    Else
      v_Nextdate := v_Starttime;
    End If;
    v_Interval := 'trunc(Sysdate)+' || v_Cyclekeep || '+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                  To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
  Elsif v_Timeunit = '周' Then
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := v_Starttime + 7;
    Else
      v_Nextdate := v_Starttime;
    End If;
    Select To_Char(v_Starttime, 'DY') Into v_Week From Dual;
    v_Interval := 'TRUNC(next_day(sysdate,''' || v_Week || '''))+7*(' || v_Cyclekeep || '-1)' || '+' ||
                  To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                  To_Char(v_Starttime, 'SS') || '/(24*60*60)';
  Elsif v_Timeunit = '月' Then
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := Add_Months(v_Starttime, 1);
    Else
      v_Nextdate := v_Starttime;
    End If;
    If To_Char(v_Starttime, 'dd') <= 28 Then
      v_Interval := 'TRUNC(ADD_MONTHS(sysdate,' || v_Cyclekeep || '))+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                    To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    Else
      Select To_Char(Last_Day(Sysdate), 'dd') - To_Char(v_Starttime, 'dd') Into v_Day From Dual;
      v_Interval := 'TRUNC(last_day(ADD_MONTHS(sysdate,' || v_Cyclekeep || '))-' || v_Day || ')+' ||
                    To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                    To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    End If;
  Else
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := Add_Months(v_Starttime, 3);
    Else
      v_Nextdate := v_Starttime;
    End If;
    If To_Char(v_Starttime, 'dd') <= 28 Then
      v_Interval := 'TRUNC(ADD_MONTHS(SYSDATE,3*' || v_Cyclekeep || '))+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                    To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    Else
      Select To_Char(Last_Day(Sysdate), 'dd') - To_Char(v_Starttime, 'dd') Into v_Day From Dual;
      v_Interval := 'TRUNC(last_day(ADD_MONTHS(SYSDATE,3*' || v_Cyclekeep || '))-' || v_Day || ')+' ||
                    To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                    To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    End If;
  End If;

  --提交作业 
  Dbms_Job.Submit(v_Jobnum, v_What, v_Nextdate, v_Interval);

  Update Zlautojobs
  Set 作业号 = v_Jobnum
  Where Nvl(系统, 0) = Nvl(Job_System, 0) And 类型 = Job_Kind And 序号 = Job_Odd;
End Zl_Jobsubmit;
/

--102323:高腾,2017-7-13,为后台作业管理添加其它时间单位.周.月.季度
Create Or Replace Procedure Zltools.Zl_Jobremove
(
  Job_System In Integer,
  Job_Kind   In Integer,
  Job_Odd    In Integer
) Authid Current_User Is
  v_Jobnum Number := 0;
Begin
  Select 作业号
  Into v_Jobnum
  From Zlautojobs
  Where Nvl(系统, 0) = Nvl(Job_System, 0) And 类型 = Job_Kind And 序号 = Job_Odd;
  --删除作业 
  Dbms_Job.Remove(v_Jobnum);

  Update Zlautojobs Set 作业号 = Null Where Nvl(系统, 0) = Nvl(Job_System, 0) And 类型 = Job_Kind And 序号 = Job_Odd;
End Zl_Jobremove;
/

--102323:高腾,2017-7-13,为后台作业管理添加其它时间单位.周.月.季度
Create Or Replace Procedure Zltools.Zl_Jobrun
(
  Job_System In Integer,
  Job_Kind   In Integer,
  Job_Odd    In Integer
) Authid Current_User Is
  v_Jobnum Number := 0;
Begin
  Select 作业号
  Into v_Jobnum
  From Zlautojobs
  Where Nvl(系统, 0) = Nvl(Job_System, 0) And 类型 = Job_Kind And 序号 = Job_Odd;
  --执行作业 
  Dbms_Job.Run(v_Jobnum);
End Zl_Jobrun;
/

--102323:高腾,2017-7-13,为后台作业管理添加其它时间单位.周.月.季度
CREATE OR REPLACE Procedure zltools.Zl_Jobchange
(
  Job_System In Integer,
  Job_Kind   In Integer,
  Job_Odd    In Integer
) Authid Current_User Is
  v_Content   Varchar2(200);
  v_Parameter Varchar2(200);
  v_Paraitem  Varchar2(200);
  v_Starttime Date;
  v_Cyclekeep Integer;
  v_Jobnum    Number := 0;
  v_What      Varchar2(1000);
  v_Nextdate  Date;
  v_Interval  Varchar2(1000);
  v_Timeunit  Varchar2(5);
  v_Week      Varchar2(10);
  v_Day       Varchar2(1);
Begin
  Select 内容, 参数, 执行时间, 间隔时间, 作业号, 时间单位
  Into v_Content, v_Parameter, v_Starttime, v_Cyclekeep, v_Jobnum, v_Timeunit
  From Zlautojobs
  Where Nvl(系统, 0) = Nvl(Job_System, 0) And 类型 = Job_Kind And 序号 = Job_Odd;
  v_What := '';

  If Length(v_Parameter) > 0 Then
    Loop
      If Instr(v_Parameter, ';') > 0 Then
        v_Paraitem  := Substr(v_Parameter, 1, Instr(v_Parameter, ';') - 1);
        v_Parameter := Substr(v_Parameter, Instr(v_Parameter, ';') + 1);
      Else
        v_Paraitem := v_Parameter;
      End If;
    
      v_What := v_What || ',' || Substr(v_Paraitem, Instr(v_Paraitem, ',') + 1);
      Exit When Instr(v_Parameter, ';') = 0;
    End Loop;
  End If;

  If Length(v_What) <> 0 Then
    v_What := v_Content || '(' || Substr(v_What, 2) || ');';
  Else
    v_What := v_Content || ';';
  End If;

  If v_Timeunit = '天' Then
    If To_Char(Sysdate, 'HH24:MI:SS') >= To_Char(v_Starttime, 'HH24:MI:SS') Then
      v_Nextdate := v_Starttime + 1;
    Else
      v_Nextdate := v_Starttime;
    End If;
    v_Interval := 'trunc(Sysdate)+' || v_Cyclekeep || '+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                  To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
  Elsif v_Timeunit = '周' Then
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := v_Starttime + 7;
    Else
      v_Nextdate := v_Starttime;
    End If;
    Select To_Char(v_Starttime, 'DY') Into v_Week From Dual;
    v_Interval := 'TRUNC(next_day(sysdate,''' || v_Week || '''))+7*(' || v_Cyclekeep || '-1)' || '+' ||
                  To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                  To_Char(v_Starttime, 'SS') || '/(24*60*60)';
  Elsif v_Timeunit = '月' Then
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := Add_Months(v_Starttime, 1);
    Else
      v_Nextdate := v_Starttime;
    End If;
    If To_Char(v_Starttime, 'dd') <= 28 Then
      v_Interval := 'TRUNC(ADD_MONTHS(sysdate,' || v_Cyclekeep || '))+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                    To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    Else
      Select To_Char(Last_Day(Sysdate), 'dd') - To_Char(v_Starttime, 'dd') Into v_Day From Dual;
      v_Interval := 'TRUNC(last_day(ADD_MONTHS(sysdate,' || v_Cyclekeep || '))-' || v_Day || ')+' ||
                    To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                    To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    End If;
  Else
    If To_Char(Sysdate, 'yyyy-MM-dd HH24:MI:SS') >= To_Char(v_Starttime, 'yyyy-MM-dd HH24:MI:SS') Then
      v_Nextdate := Add_Months(v_Starttime, 3);
    Else
      v_Nextdate := v_Starttime;
    End If;
    If To_Char(v_Starttime, 'dd') <= 28 Then
      v_Interval := 'TRUNC(ADD_MONTHS(SYSDATE,3*' || v_Cyclekeep || '))+' || To_Char(v_Starttime, 'HH24') || '/24' || '+' ||
                    To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' || To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    Else
      Select To_Char(Last_Day(Sysdate), 'dd') - To_Char(v_Starttime, 'dd') Into v_Day From Dual;
      v_Interval := 'TRUNC(last_day(ADD_MONTHS(SYSDATE,3*' || v_Cyclekeep || '))-' || v_Day || ')+' ||
                    To_Char(v_Starttime, 'HH24') || '/24' || '+' || To_Char(v_Starttime, 'MI') || '/(24*60)' || '+' ||
                    To_Char(v_Starttime, 'SS') || '/(24*60*60)';
    End If;
  End If;

  --修改作业 
  Dbms_Job.Change(v_Jobnum, v_What, v_Nextdate, v_Interval);
End Zl_Jobchange;
/

--111523:高腾,2017-7-20,将记录错误日志的SQL语句改为过程,避免硬解析
Create Or Replace Procedure Zltools.Zl_Zlerrorlog_Insert
(
  工作站_In    Zlerrorlog.工作站%Type,
  类型_In      Zlerrorlog.类型%Type,
  错误序号_In  Zlerrorlog.错误序号%Type,
  错误信息_In  Zlerrorlog.错误信息%Type,
  Sessionid_In Number := Null
) Is
Begin
  If Sessionid_In Is Null Then
    Insert Into Zlerrorlog
      (会话号, 用户名, 工作站, 时间, 类型, 错误序号, 错误信息)
      Select Sid, User, 工作站_In, Sysdate, 类型_In, 错误序号_In, 错误信息_In
      From V$session
      Where Audsid = Userenv('SessionID');
  Else
    Insert Into Zlerrorlog
      (会话号, 用户名, 工作站, 时间, 类型, 错误序号, 错误信息)
      Select Sid, User, 工作站_In, Sysdate, 类型_In, 错误序号_In, 错误信息_In
      From GV$session
      Where Audsid = Sessionid_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlerrorlog_Insert;
/

--111523:高腾,2017-7-20,将记录运行日志SQL语句改为过程,避免硬解析
Create Or Replace Procedure Zltools.Zl_Zldiarylog_Insert
(
  工作站_In    Zldiarylog.工作站%Type,
  部件名_In    Zldiarylog.部件名%Type,
  窗体名_In    Zldiarylog.窗体名%Type,
  工作内容_In  Zldiarylog.工作内容%Type,
  Sessionid_In Number := Null
) Is
Begin
  If Sessionid_In Is Null Then
    Insert Into Zldiarylog
      (会话号, 用户名, 工作站, 部件名, 窗体名, 工作内容, 进入时间)
      Select Sid + Serial#, User, 工作站_In, 部件名_In, 窗体名_In, 工作内容_In, Sysdate
      From V$session
      Where Audsid = Userenv('SessionID') And Machine Is Not Null;
  Else
    Insert Into Zldiarylog
      (会话号, 用户名, 工作站, 部件名, 窗体名, 工作内容, 进入时间)
      Select Sid + Serial#, User, 工作站_In, 部件名_In, 窗体名_In, 工作内容_In, Sysdate
      From GV$session
      Where Audsid = Sessionid_In And Machine Is Not Null;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zldiarylog_Insert;
/

--111523:高腾,2017-7-20,将更新运行日志SQL语句改为过程,避免硬解析
Create Or Replace Procedure Zltools.Zl_Zldiarylog_Update
(
  工作站_In    Zldiarylog.工作站%Type,
  部件名_In    Zldiarylog.部件名%Type,
  窗体名_In    Zldiarylog.窗体名%Type,
  退出原因_In  Zldiarylog.退出原因%Type,
  Sessionid_In Number := Null
) Is
  n_会话号 zldiarylog.会话号%type;
  v_用户名 zldiarylog.用户名%type;
Begin
  If Sessionid_In Is Null Then
    Select Sid + Serial#, User
    Into n_会话号, v_用户名
    From V$session
    Where Audsid = Userenv('SessionID') And Machine Is Not Null;
  Else
    Select Sid + Serial#, User
    Into n_会话号, v_用户名
    From GV$session
    Where Audsid = Sessionid_In And Machine Is Not Null;
  End If;
  Update Zldiarylog
  Set 退出原因 = 退出原因_In, 退出时间 = Sysdate
  Where 退出原因 Is Null And 用户名 = v_用户名 And 工作站 = 工作站_In And 会话号 = n_会话号 And 部件名 = 部件名_In And 窗体名 = 窗体名_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zldiarylog_Update;
/

--100682:高腾,2017-8-29,重要操作日志记录
Create Table Zltools.zlAuditLog(
用户名 Varchar2(30),
工作站 Varchar2(50),
操作时间 Date,
操作类型 Number(2),
操作模块编号 Varchar2(18),
操作内容 Varchar2(1024),
操作说明 Varchar2(256));
Alter Table zltools.zlAuditLog Add Constraint zlAuditLog_PK Primary Key (用户名,工作站,操作时间,操作模块编号) Using Index;
Insert into ZLTools.zlTables(系统,表名,表空间,分类) Values(0,'ZLAUDITLOG','ZLTOOLSTBS','B3');

--100682:高腾,2017-8-29,重要操作日志记录
Create Or Replace Procedure Zltools.Zl_Zlauditlog_Insert
(
  用户名_In       Zlauditlog.用户名%Type,
  工作站_In       Zlauditlog.工作站%Type,
  操作类型_In     Zlauditlog.操作类型%Type, --1-新增，2-修改，3-删除
  操作模块编号_In Zlauditlog.操作模块编号%Type,
  操作内容_In     Zlauditlog.操作内容%Type,
  操作说明_In     Zlauditlog.操作说明%Type --用来记录界面提供给操作员输入的备注信息
) Is
Begin
  Insert Into Zlauditlog
    (用户名, 工作站, 操作时间, 操作类型, 操作模块编号, 操作内容, 操作说明)
  Values
    (用户名_In, 工作站_In, Sysdate, 操作类型_In, 操作模块编号_In, 操作内容_In, 操作说明_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Zlauditlog_Insert;
/

--111526:高腾,2017-7-5,日志启停管理
Create Table Zltools.ZLLogConfig(
    编号 number(4),
    名称 varchar(20),
    说明 varchar2(500));
Alter Table Zltools.ZLLogConfig Add Constraint ZLLogConfig_Pk Primary Key(编号) Using Index;

--111538:高腾,2017-7-25,配合RAC环境将Zl_Autologprocess中的v$Session改为Gv$Session
Create Or Replace Procedure Zltools.Zl_Autologprocess As
  --功能： 
  --   1.对多余的运行日志和错误日志进行清除 
  --   2.对异常的运行日志进行标记 
  v_Count Number;
  v_Limit Number;
Begin
  --删除多余的运行日志 
  Select Count(*) Into v_Count From Zldiarylog;
  Begin
    Select Nvl(To_Number(参数值), 0) Into v_Limit From Zloptions Where 参数号 = 2;
  Exception
    When Others Then
      v_Limit := 10000;
  End;
  If v_Count > v_Limit Then
    Delete From Zldiarylog
    Where Rowid In (Select Id
                    From (Select Rowid As Id From Zldiarylog Group By 进入时间, Rowid)
                    Where Rownum < v_Count - v_Limit + 1);
  End If;

  --对异常退出的运行日志记录进行处理 
  Update Zldiarylog
  Set 退出原因 = 2, 退出时间 = Sysdate
  Where 退出原因 Is Null And 会话号 Not In (Select Sid + Serial# From Gv$session Where User# <> 0);

  --删除多余的错误日志 
  Select Count(*) Into v_Count From Zlerrorlog;
  Begin
    Select Nvl(To_Number(参数值), 0) Into v_Limit From Zloptions Where 参数号 = 4;
  Exception
    When Others Then
      v_Limit := 10000;
  End;
  If v_Count > v_Limit Then
    Delete From Zlerrorlog
    Where Rowid In
          (Select Id From (Select Rowid As Id From Zlerrorlog Group By 时间, Rowid) Where Rownum < v_Count - v_Limit + 1);
  End If;
End Zl_Autologprocess;
/

--111538:高腾,2017-7-25,配合RAC环境将getClient中的v$Session改为Gv$Session
CREATE OR REPLACE Package Zltools.b_Runmana Is
 
  Type t_Refcur Is Ref Cursor; 
 
  Procedure Get_Parameters 
  ( 
    Cursor_Out Out t_Refcur, 
    系统_In    In Number := 0 
  ); 
 
  Procedure Get_Parameter 
  ( 
    Cursor_Out Out t_Refcur, 
    参数id_In  In Zlparameters.Id%Type 
  ); 
 
  Procedure Get_Parachangedlog 
  ( 
    Cursor_Out Out t_Refcur, 
    参数id_In  In Zlparachangedlog.参数id%Type 
  ); 
 
  Procedure Get_Job_Number 
  ( 
    Cursor_Out Out t_Refcur, 
    系统_In    In Number 
  ); 
 
  Procedure Get_Depict 
  ( 
    Cursor_Out Out t_Refcur, 
    系统_In    In Zldatamove.系统%Type, 
    组号_In    In Zldatamove.组号%Type 
  ); 
 
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur); 
 
  Procedure Get_Client 
  ( 
    Cur_Out   Out t_Refcur, 
    工作站_In In Zlclients.工作站%Type := Null 
  ); 
 
  Procedure Get_Client_Station(Cur_Out Out t_Refcur); 
 
  Procedure Get_Project_No(Cur_Out Out t_Refcur); 
 
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur); 
 
  Procedure Get_Resile 
  ( 
    Cur_Out   Out t_Refcur, 
    方案号_In In Zlclientparaset.方案号%Type, 
    类型_In   In Number := 0 
  ); 
 
  Procedure Get_Zldatamove 
  ( 
    Cur_Out Out t_Refcur, 
    系统_In In Zldatamove.系统%Type 
  ); 
 
  Procedure Get_Log 
  ( 
    Cur_Out     Out t_Refcur, 
    日志类型_In In Varchar2, 
    Where_In    In Varchar2 
  ); 
 
  Procedure Get_Log_Count 
  ( 
    Cur_Out     Out t_Refcur, 
    日志类型_In In Varchar2 
  ); 
 
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur); 
 
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur); 
 
  Procedure Get_Zloption 
  ( 
    Cur_Out   Out t_Refcur, 
    参数号_In In Zloptions.参数号%Type 
  ); 
 
End b_Runmana;
/

--111538:高腾,2017-7-25,配合RAC环境将getClient中的v$Session改为Gv$Session
CREATE OR REPLACE Package Body Zltools.b_Runmana Is
 
  --功能：取参数信息 
  --frmParameters 
  Procedure Get_Parameters 
  ( 
    Cursor_Out Out t_Refcur, 
    系统_In    In Number := 0 
  ) Is 
  Begin 
    If Nvl(系统_In, 0) = 0 Then 
      Open Cursor_Out For 
        Select a.Id, Nvl(a.系统, 0) 系统, Nvl(a.模块, 0) 模块, Nvl(a.私有, 0) 私有, a.参数号, a.参数名, a.参数值, a.缺省值, Nvl(a.性质, 0) 性质, 
               a.影响控制说明, a.参数值含义, a.关联说明, a.适用说明, a.警告说明, Nvl(a.本机, 0) 本机, Nvl(a.授权, 0) 授权, Nvl(a.固定, 0) 固定, 
               Nvl(a.部门, 0) 部门, b.标题 As 模块名称, zlSpellCode(b.标题) As 模块简码 
        From zlParameters A, zlPrograms B 
        Where Nvl(a.系统, 0) = 0 And Nvl(a.系统, 0) = b.系统(+) And Nvl(a.模块, 0) = b.序号(+); 
    Else 
      Open Cursor_Out For 
        Select a.Id, Nvl(a.系统, 0) 系统, Nvl(a.模块, 0) 模块, Nvl(a.私有, 0) 私有, a.参数号, a.参数名, a.参数值, a.缺省值, Nvl(a.性质, 0) 性质, 
               a.影响控制说明, a.参数值含义, a.关联说明, a.适用说明, a.警告说明, Nvl(a.本机, 0) 本机, Nvl(a.授权, 0) 授权, Nvl(a.固定, 0) 固定, 
               Nvl(a.部门, 0) 部门, b.标题 As 模块名称, zlSpellCode(b.标题) As 模块简码 
        From zlParameters A, zlPrograms B, 
             --处理权限部分，只有授权的才能显示 
             (Select Distinct f.序号 
               From zlProgFuncs F, zlRegFunc R 
               Where Trunc(f.系统 / 100) = r.系统(+) And f.序号 = r.序号(+) And f.功能 = r.功能(+) And 
                     (r.功能 Is Not Null Or r.功能 Is Null And (f.序号 Between 10000 And 19999)) And f.系统 = 系统_In And 
                     1 = (Select 1 From Zlregaudit A Where a.项目 = '授权证章') 
               Union All 
               Select 0 As 序号 
               From Dual) M 
        Where a.系统 = Nvl(系统_In, 0) And Nvl(a.系统, 0) = b.系统(+) And Nvl(a.模块, 0) = b.序号(+) And Nvl(a.模块, 0) = m.序号; 
    End If; 
  End Get_Parameters; 
 
  --功能：根据指定的参数ID取参数信息 
  --调用列表：frmParameters;frmParaChangeSet 
  Procedure Get_Parameter 
  ( 
    Cursor_Out Out t_Refcur, 
    参数id_In  In Zlparameters.Id%Type 
  ) Is 
  Begin 
    Open Cursor_Out For 
      Select a.Id, Nvl(a.系统, 0) 系统, Nvl(a.模块, 0) 模块, Nvl(a.私有, 0) 私有, a.参数号, a.参数名, a.参数值, a.缺省值, Nvl(a.性质, 0) 性质, 
             a.影响控制说明, a.参数值含义, a.关联说明, a.适用说明, a.警告说明, Nvl(a.本机, 0) 本机, Nvl(a.授权, 0) 授权, Nvl(a.固定, 0) 固定, 
             Nvl(a.部门, 0) 部门, b.标题 As 模块名称, zlSpellCode(b.标题) As 模块简码 
      From zlParameters A, zlPrograms B 
      Where a.Id = Nvl(参数id_In, 0) And Nvl(a.系统, 0) = b.系统(+) And Nvl(a.模块, 0) = b.序号(+); 
  End Get_Parameter; 
  --功能：取参数修改信息 
  --调用列表：frmParameters 
  Procedure Get_Parachangedlog 
  ( 
    Cursor_Out Out t_Refcur, 
    参数id_In  In Zlparachangedlog.参数id%Type 
  ) Is 
  Begin 
    Open Cursor_Out For 
      Select 参数id, 序号, 变动说明, 变动内容, 变动人, 变动时间, 变动原因 
      From Zlparachangedlog 
      Where 参数id = Nvl(参数id_In, 0); 
 
  End; 
  --功能：取ZlAutoJob序列号 
  Procedure Get_Job_Number 
  ( 
    Cursor_Out Out t_Refcur, 
    系统_In    In Number 
  ) Is 
  Begin 
    Open Cursor_Out For 
      Select 序号 + 1 As 序号 
      From zlAutoJobs 
      Where Nvl(系统, 0) = 系统_In And 类型 = 3 And 
            序号 + 1 Not In (Select 序号 From zlAutoJobs Where Nvl(系统, 0) = 系统_In And 类型 = 3); 
  End Get_Job_Number; 
 
  --功能：取ZlDataMove描述 
  Procedure Get_Depict 
  ( 
    Cursor_Out Out t_Refcur, 
    系统_In    In Zldatamove.系统%Type, 
    组号_In    In Zldatamove.组号%Type 
  ) Is 
  Begin 
    Open Cursor_Out For 
      Select 转出描述 From zlDataMove Where Nvl(系统, 0) = 系统_In And 组号 = 组号_In; 
  End Get_Depict; 
 
  --功能：取zlClients的MAX IP 
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur) Is 
  Begin 
    Open Cur_Out For 
      Select Max(Ip) As Ip From zlClients; 
  End Get_Client_Maxip; 
 
  --功能：取zlClients的记录 
  Procedure Get_Client 
  ( 
    Cur_Out   Out t_Refcur, 
    工作站_In In Zlclients.工作站%Type := Null 
  ) Is 
    v_Sql Varchar2(1000); 
  Begin 
    If Nvl(工作站_In, '空') = '空' Then 
      v_Sql := 'Select a.Ip, a.工作站, a.Cpu, a.内存, a.硬盘, a.操作系统, a.部门, a.用途, a.说明, a.升级标志, a.禁止使用,
                             a.连接数, Decode(b.Terminal, Null, 0, 1) As 状态, a.收集标志,a.升级服务器,a.站点,a.启用视频源
                From Zlclients a, (Select Distinct Terminal From GV$session) b
                Where Upper(a.工作站) = Upper(b.Terminal(+))
                Order By a.Ip'; 
      Open Cur_Out For v_Sql; 
    Else 
      Open Cur_Out For 
        Select Ip, 工作站, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级标志, 收集标志, 禁止使用, 连接数, 升级服务器, 站点, 启用视频源 
        From zlClients 
        Where Upper(工作站) = 工作站_In; 
    End If; 
  End Get_Client; 
 
  --功能：取zlClients的站点 
  Procedure Get_Client_Station(Cur_Out Out t_Refcur) Is 
  Begin 
    Open Cur_Out For 
      Select Distinct Upper(工作站) || '[' || Ip || ']' As 站点, Upper(工作站) 工作站 From zlClients; 
  End Get_Client_Station; 
 
  --功能：取方案号 
  Procedure Get_Project_No(Cur_Out Out t_Refcur) Is 
  Begin 
    Open Cur_Out For 
      Select 方案号 From Zlclientparaset Where Rownum = 1; 
  End Get_Project_No; 
 
  --功能：取方案 
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur) Is 
  Begin 
    Open Cur_Out For 
      Select 方案号, 方案号 || '-' || 方案名称 As 方案名称, 方案描述, 工作站, 用户名 From Zlclientscheme; 
  End Get_Client_Scheme; 
 
  --功能：取恢复信息 
  Procedure Get_Resile 
  ( 
    Cur_Out   Out t_Refcur, 
    方案号_In In Zlclientparaset.方案号%Type, 
    类型_In   In Number := 0 
  ) Is 
  Begin 
    If 类型_In = 0 Then 
      Open Cur_Out For 
        Select Distinct a.工作站 || Decode(m.工作站, Null, ' ', '[' || m.Ip || ']') As 工作站, a.用户名, a.恢复标志, 
                        '[' || b.方案号 || ']' || b.方案名称 As 方案名称 
        From Zlclientparaset A, Zlclientscheme B, zlClients M 
        Where a.方案号 = b.方案号 And a.工作站 = m.工作站(+) And a.方案号 = 方案号_In; 
    End If; 
 
    If 类型_In = 1 Then 
      Open Cur_Out For 
        Select Distinct Upper(工作站) 工作站, Min(恢复标志) 恢复标志 
        From Zlclientparaset A 
        Where a.方案号 = 方案号_In 
        Group By 工作站; 
    End If; 
 
    If 类型_In = 2 Then 
      Open Cur_Out For 
        Select Distinct Upper(用户名) 用户名, Max(工作站) 工作站, Min(Decode(恢复标志, 2, 0, 恢复标志)) 恢复标志 
        From Zlclientparaset A 
        Where a.方案号 = 方案号_In 
        Group By 用户名 
        Order By 用户名; 
    End If; 
 
  End Get_Resile; 
 
  --功能：取zldataMove数据 
  Procedure Get_Zldatamove 
  ( 
    Cur_Out Out t_Refcur, 
    系统_In In Zldatamove.系统%Type 
  ) Is 
  Begin 
    Open Cur_Out For 
      Select 组号, 组名, 说明, 日期字段, 转出描述, 上次日期 From zlDataMove Where 系统 = 系统_In Order By 组号; 
  End Get_Zldatamove; 
 
  --功能：取日志数据 
  Procedure Get_Log 
  ( 
    Cur_Out     Out t_Refcur, 
    日志类型_In In Varchar2, 
    Where_In    In Varchar2 
  ) Is 
    v_Sql Varchar2(1000); 
  Begin 
    If 日志类型_In = '错误日志' Then 
      v_Sql := 'Select 会话号,工作站,用户名,错误序号,错误信息,To_char(时间,''yyyy-MM-dd hh24:mi:ss'') 时间
                     ,Decode(类型,1,''存储过程错误'',2,''数据联结层错误'',3,''应用程序层错误'',''客户端升级错误'') 错误类型
                        From ZlErrorLog Where ' || Where_In; 
      Open Cur_Out For v_Sql; 
    End If; 
    If 日志类型_In = '运行日志' Then 
      v_Sql := 'Select 会话号,工作站,用户名,部件名,工作内容,To_char(进入时间,''yyyy-MM-dd hh24:mi:ss'') 进入时间
                                 ,To_char(退出时间,''yyyy-MM-dd hh24:mi:ss'') 退出时间,Decode(退出原因,1,''正常退出'',''异常退出'') 退出原因
                                    From ZlDiaryLog Where ' || Where_In; 
      Open Cur_Out For v_Sql; 
    End If; 
  End Get_Log; 
 
  --功能：取日志记录数 
  Procedure Get_Log_Count 
  ( 
    Cur_Out     Out t_Refcur, 
    日志类型_In In Varchar2 
  ) Is 
  Begin 
    If 日志类型_In = '错误日志' Then 
      Open Cur_Out For 
        Select Count(*) 数量 
        From zlErrorLog 
        Union All 
        Select Nvl(To_Number(参数值), 0) 
        From zlOptions 
        Where 参数号 = 4; 
    End If; 
    If 日志类型_In = '运行日志' Then 
      Open Cur_Out For 
        Select Count(*) 数量 
        From zlDiaryLog 
        Union All 
        Select Nvl(To_Number(参数值), 0) 
        From zlOptions 
        Where 参数号 = 2; 
 
    End If; 
  End Get_Log_Count; 
 
  --功能：取zlfilesupgradeg数据 
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur) Is 
  Begin 
    Open Cur_Out For 
      Select 序号, 文件名, 版本号, 修改日期, 文件说明 As 说明, 
             Decode(文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '') As 类型, 安装路径 As 安装路径, 
             Md5 As Md5, 加入日期 
      From zlFilesUpgrade 
      Order By 序号; 
  End Get_Zlfilesupgrade; 
 
  --功能：取非注册项目 
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur) Is 
  Begin 
    Open Cur_Out For 
      Select 项目, 内容 
      From zlRegInfo 
      Where 项目 Not In ('发行码', '版本号', '服务器目录', '访问用户', '访问密码', '收集目录', '收集类型', '注册码', '授权证章', '授权工具', '授权邮戳'); 
  End Get_Not_Regist; 
 
  --功能：取参数值 
  Procedure Get_Zloption 
  ( 
    Cur_Out   Out t_Refcur, 
    参数号_In In Zloptions.参数号%Type 
  ) Is 
  Begin 
    Open Cur_Out For 
      Select Nvl(参数值, 缺省值) Option_Value From zlOptions Where 参数号 = 参数号_In; 
  End Get_Zloption; 
 
End b_Runmana;
/

--110944:余智勇,2017-07-25,表格自适应行高
Alter Table zlTools.zlRPTItems Add 自适应行高 Number(1);

--104724:余智勇,2017-08-08,报表启停功能
Alter Table zlTools.zlReports Add 是否停用 Number(1);
Alter Table zlTools.zlRPTGroups Add 是否停用 Number(1);

--113763:杨周一,2017-08-31,管理工具中添加DBA管理工具
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('06',Null,'DBA工具','D',Null,Null);
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0601','06','数据库性能','M',Null,1);
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0602','06','SQL性能','T',Null,2);
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0603','06','SQL跟踪','S',Null,3);
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0604','06','会话解锁','B',Null,4);
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0605','06','外键索引','F',Null,5);
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0606','06','空间管理','R',Null,6);