--批量导入产科护士账号，可根据实际情况调整SQL查询，ZLSOL为助产士的工作区名称；123为登录密码
--测试导入100人需要1分钟的时间
--以zlsol登录执行以下脚本；conn zlsol/his@ORA_SOL
Declare
  Cursor c_Sql Is
    Select Distinct a.用户名, Substr(d.姓名, 1, 1) 姓, Substr(d.姓名, 2) 名, d.简码
    From 上机人员表@Zlhis_Dbl a, 部门性质说明@Zlhis_Dbl b, 部门人员@Zlhis_Dbl c, 人员表@Zlhis_Dbl d
    Where b.部门id = c.部门id And c.人员id = d.Id And a.人员id = c.人员id And b.工作性质 = '产科';
  n_group_id number(18);
Begin
  For r In c_Sql Loop
    n_group_id:=Apex_Util.find_security_group_id(p_workspace => 'ZLSOL');
    Apex_Util.set_security_group_id(p_security_group_id => n_group_id);
    Apex_Util.Create_User(p_User_Name => r.用户名, p_First_Name => r.名, p_Last_Name => r.姓, p_Web_Password => '123',p_Change_Password_On_First_Use => 'N');
  End Loop;
End;
