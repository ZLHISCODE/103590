以admin用户登录管理界面：http://IP:8080/apex/apex_admin（注意选择简体中文）
1.菜单：管理工作区创建工作区,工作区名称：ZLSOL
2.是否重用现有方案（是）,方案名：ZLSOL
3.管理员用户名：admin，密码：his,名：理员，姓:管

创建成功后，以管理员用户admin登录ZLSOL工作区来执行apex脚本的导入
http://IP:8080/apex
工作区：ZLSOL，用户名:admin,密码：his

点菜单：应用程序构建器
再按钮：导入
   不要直接点下拉菜单中的导入，否则会报错：ERR-1002 在应用程序 "4000" 中未找到项 "F4000_P56_CREATE_OPTION" 的项 ID。）
导入文件：选择文件：6.ZLSOL_APEX_yyyymmdd.sql
其他的保持缺省项，一直点下一步即可完成导入。