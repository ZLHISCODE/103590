----10.35.80---》10.35.90
--117919:蒋敏,2018-1-2,将Long Raw或者long类型转换为BLOB或者CLOB类型
Alter Table zltools.zlRPTGraphs Modify 图片 Blob;
Alter Table zltools.zlXlsVerify Modify 审核信息 Clob;

Declare  
  Cursor c_失效索引 Is
    Select 'alter index ' || Index_Name || ' rebuild' As 索引
    From user_Indexes
    Where Table_Name In ('ZLRPTGRAPHS', 'ZLXLSVERIFY') And Status = 'UNUSABLE';
Begin
For r_失效索引 In c_失效索引 Loop
    Execute Immediate r_失效索引.索引;
  End Loop;
End;
/

--119766:刘硕,2017-12-06,进程清单处理
alter table zltools.zlkillprocess add  是否固定 number(1);

Alter Table zltools.ZLKillProcess Add Constraint ZLKillProcess_PK Primary Key(序号) Using Index;

Insert Into zltools.ZLKILLPROCESS(序号,名称,类型,是否固定,描述) 
Select 1   ,'7Z.EXE'                  ,0  ,1  ,'7-ZIP压缩程序' From Dual Union All
Select 2   ,'WINCMP3.EXE'             ,0  ,1  ,'文件比对工具，用于用户自定义过程收集。' From Dual Union All
Select 3   ,'ZL9LABPRINTSVR.EXE'      ,0  ,1  ,'新版LIS打印服务，主要处理批量打印报告。' From Dual Union All
Select 4   ,'ZL9LABRECEIV.EXE'        ,0  ,1  ,'新版LIS通讯程序部件，主要处理与仪器接口之间数据交互。' From Dual Union All
Select 5   ,'ZL9LABTCPSVR.EXE'        ,0  ,1  ,'新版LIS检验消息转发部件，处理实验室和通讯程序间的消息转发。' From Dual Union All
Select 6   ,'ZL9LISCOMM.EXE'          ,0  ,1  ,'老版检验通讯程序，处理仪器回传数据，加工成检验系统能够认识的数据格式。' From Dual Union All
Select 7   ,'ZL9PACSCAPTURE.EXE'      ,0  ,1  ,'因影像采集系统的视频采集方式优化调整,而独立出来的ActivexExe项目' From Dual Union All
Select 8   ,'ZL9WIZARDMAIN.EXE'       ,0  ,1  ,'病人自助系统后台管理程序，完成自助系统的所有后台设置，包括资源配置、动态页面设计、静态页面参数等。' From Dual Union All
Select 9   ,'ZL9XLS.EXE'              ,0  ,1  ,'Excel报表工具。' From Dual Union All
Select 10  ,'ZLACTMAIN.EXE'           ,0  ,1  ,'BH融合中的虚拟导航台，BH调用各个模块均通过该程序进行导航。' From Dual Union All
Select 11  ,'ZLBAEXPORT.EXE'          ,0  ,1  ,'完成病案数据dbf文件的生成和上传至FTP' From Dual Union All
Select 12  ,'ZLCDOPEN.EXE'            ,0  ,1  ,'打开PACS刻录到光盘上的检查图像，辅助工具，打开PACS刻录到光盘上的检查图像。' From Dual Union All
Select 13  ,'ZLCISAUDITPRINT.EXE'     ,0  ,1  ,'用于电子病案审查中,文件-输出到PDF，避免连续PDF输出引起系统GDI超量，导致系统假死。' From Dual Union All
Select 14  ,'ZLDBATOOLS.EXE'          ,0  ,1  ,'DBA管理工具单独执行文件。' From Dual Union All
Select 15  ,'ZLDRUGMACHINEMANAGE.EXE' ,0  ,1  ,'药房自动发药系统接口配置和业务管理。' From Dual Union All
Select 16  ,'ZLEXINSTALL.EXE'         ,0  ,1  ,'附加组件安装程序，现仅支持对OO4O组件安装。' From Dual Union All
Select 17  ,'ZLGETIMAGE.EXE'          ,0  ,1  ,'提供影像检查图像下载支持，后台下载影像检查图像。' From Dual Union All
Select 18  ,'ZLGETIMAGEEX.EXE'        ,0  ,1  ,'zlGetImageEx是使用ActiveExe的方式实现后台进程加载及上传图像，是一个ActiveExe部件' From Dual Union All
Select 19  ,'ZLHEALTHSERVICE.EXE'     ,0  ,1  ,'实现健康体检中心后台服务运行的启动程序' From Dual Union All
Select 20  ,'ZLHIS+.EXE'              ,0  ,1  ,'ZLHIS+启动程序，登录该程序才能进入导航台，进行业务操作。' From Dual Union All
Select 21  ,'ZLHISCRUST.EXE'          ,0  ,1  ,'客户端自动升级工具，通过该工具对各个客户端进行文件升级。' From Dual Union All
Select 22  ,'ZLHQMSDCOLLECT.EXE'      ,0  ,1  ,'完成医院质量监测数据的采集上报工作' From Dual Union All
Select 23  ,'ZLLISMESSAGE.EXE'        ,0  ,1  ,'LIS消息部件，在大屏幕上显示某些检验科内部的情况。' From Dual Union All
Select 24  ,'ZLLISRECEIVESEND.EXE'    ,0  ,1  ,'部件功能:主要与检验仪器直接通讯，记录仪器回传的检验结果，并保存文本为LIS认识的检验结果。' From Dual Union All
Select 31  ,'ZLNEWQUERY.EXE'          ,0  ,1  ,'老版自助系统,自助挂号、Lis打印、费用查询。' From Dual Union All
Select 32  ,'ZLORCLCONFIG.EXE'        ,0  ,1  ,'用于快速配置ORACLE配置文件的工具。' From Dual Union All
Select 33  ,'ZLPACSBROWSERSTATION.EXE',0  ,1  ,'独立观片站。' From Dual Union All
Select 34  ,'ZLPACSFTPTOOLS.EXE'      ,0  ,1  ,'对FTP进行测试，排查FTP相关操作错误。' From Dual Union All
Select 35  ,'ZLPACSSERVERCENTER.EXE'  ,1  ,1  ,'PACS服务中心' From Dual Union All
Select 36  ,'ZLPACSSRV.EXE'           ,0  ,1  ,'接受Dicom设备发送的检查图像，PACS网关服务，监听影像DICOM设备请求。' From Dual Union All
Select 37  ,'ZLPEISAUTOANALYSE.EXE'   ,0  ,1  ,'体检自动分析服务，实现非标准的仪器数据接口。' From Dual Union All
Select 38  ,'ZLQUEUESHOW.EXE'         ,0  ,1  ,'新版排队显示，pacs排队情况显示。' From Dual Union All
Select 39  ,'ZLRISDUMPTOOL.EXE'       ,0  ,1  ,'基础数据，用户，诊疗项目，数据字典等初始化，初始化ris接口数据。' From Dual Union All
Select 40  ,'ZLRPTSQLADJUST.EXE'      ,0  ,1  ,'10.26病人费用表拆分配套工具。进行大表拆分后的涉及病人费用记录的报表的调整。' From Dual Union All
Select 41  ,'ZLRUNAS.EXE'             ,0  ,1  ,'该文件在自动升级zlhisCrust.exe中使用。主要功能,在USER权限下可以使用管理员权限来进行登录执行管理操作' From Dual Union All
Select 42  ,'ZLSCREENKEYBOARD.EXE'    ,0  ,1  ,'屏幕键盘小程序，在门诊医生工作站中用到，强制续诊，门诊医嘱下达。' From Dual Union All
Select 43  ,'ZLSOFTSHOWARCHIVE.EXE'   ,0  ,1  ,'显示病历查阅,医嘱，pacs历史报告等，ris中调用查看病历内容。' From Dual Union All
Select 44  ,'ZLSOFTSHOWHISFORMS.EXE'  ,0  ,1  ,'显示病历查阅,医嘱，pacs历史报告等，ris中调用查看病历内容。' From Dual Union All
Select 45  ,'ZLSQLTRACE.EXE'          ,0  ,1  ,'zlSQL跟踪工具。' From Dual Union All
Select 46  ,'ZLSVRNOTICE.EXE'         ,0  ,1  ,'自动提醒服务，进行消息提醒的提示与阅读。' From Dual Union All
Select 47  ,'ZLSVRSTUDIO.EXE'         ,0  ,1  ,'ZLHIS系统的后台管理工具。提供了系统的升级、安装、授权以及其他的实用功能，可以方便的进行后台管理。' From Dual Union All
Select 48  ,'ZLUPGRADEREADER.EXE'     ,0  ,1  ,'升级说明阅读器。进行重大功能的核对以及培训事宜的处理。' From Dual Union All
Select 49  ,'ZLWIZARDSTART.EXE'       ,0  ,1  ,'自助系统前台查询启动程序，启动自助系统前台功能。' From Dual;

--117980:高腾,2017-12-06,重要操作变动日志管理界面
Insert Into Zltools.Zlsvrtools (编号, 上级, 标题, 快键, 说明, 次序) Values ('0314', '03', '操作日志管理', 'T', Null, 17);
Alter Table Zltools.zlauditlog add 日志ID Number(18);

--117980:高腾,2017-12-06,重要操作变动日志管理界面
Create Table Zltools.ZlauditlogConfig(
       ID Number(18),
       系统 Number(5),
       模块 Varchar2(18),
       功能 Varchar2(50),
       说明 Varchar2(250),
       是否启用 Number(1),
       是否需审核 Number(1)
);
Alter Table Zltools.zlauditlogconfig Add Constraint ZlauditlogConfig_PK Primary Key(ID) Using Index;
Alter Table Zltools.ZlauditlogConfig Add Constraint ZlauditlogConfig_UQ_系统 Unique(系统, 模块,功能) Using Index;
Alter Table Zltools.ZlauditlogConfig Add Constraint ZlauditlogConfig_FK_系统 Foreign Key(系统) References Zlsystems(编号) On Delete Cascade;
Insert Into Zltools.zlTables(系统,表名,表空间,分类) Values(0,'ZLAUDITLOGCONFIG','ZLTOOLSTBS','A2');
Insert Into Zltools.Zloptions(参数号, 参数名, 参数值, 缺省值, 参数说明) Values(25, '操作日志保存最大天数', '365', '365', '操作日志最多能保存的天数，超过时系统将其自动删除。至少保存90天，天数为0时表示永久保存');
Create Sequence Zltools.ZlauditlogConfig_ID start with 1;
CREATE INDEX Zltools.zlClients_IX_IP ON zlClients(IP);

--119259:高腾,2018-1-2,管理工具其它模块添加重要操作变动日志
Insert Into Zltools.Zlauditlogconfig(ID, 模块, 功能, 说明, 是否启用, 是否需审核)
Select 1, '0201', '拆卸','拆卸历史数据空间，若选择的数据空间是当前正在使用的，则不能拆卸',1,1 From Dual Union All
Select 2, '0201', '切换','切换当前历史数据空间，并将重新创建H表视图指向切换后的历史数据空间',1,0 From Dual Union All
Select 3, '0201', '合并','将至少两个历史数据空间合并为一个，数据会自动合并到编号最小的空间中',1,0 From Dual Union All
Select 4, '0202', '执行','将表数据、约束、索引或权限导出到本地文件',1,0 From Dual Union All
Select 5, '0203', '执行','将表数据、约束、索引或权限由本地文件导入数据库',1,0 From Dual Union All
Select 6, '0206', '执行','将指定表中的数据全部清空',1,1 From Dual Union All
Select 7, '0207', '新增','添加一个数据连接，用于客户端连接其他数据库来查询数据',1,0 From Dual Union All
Select 8, '0207', '修改','修改一个数据连接，可能导致原来使用该连接的功能查询数据有误',1,0 From Dual Union All
Select 9, '0207', '删除','删除一个数据连接，删除后，客户端将不能使用该连接来连接其对应数据库',1,1 From Dual Union All
Select 10, '0302', '结束会话','强行断开一个用户与数据库之间的连接，可能导致该用户未保存的数据丢失',1,0 From Dual Union All
Select 11, '0303', '增加','添加一个自动作业项目，用于定时定周期完成指定任务',1,0 From Dual Union All
Select 12, '0303', '删除','删除一个自动作业项目，若该自动作业已启用定时执行，删除后将会导致该定时任务无法执行',1,1 From Dual Union All
Select 13, '0303', '运行设置','修改一个自动作业项目，包括名称，任务内容，循环时间及执行时间等',1,0 From Dual Union All
Select 14, '0304', '删除','删除指定条件下或全部的运行日志',1,1 From Dual Union All
Select 15, '0305', '删除','删除指定条件下或全部的错误日志',1,1 From Dual Union All
Select 16, '0306', '保存','保存一些重要系统参数的修改',1,0 From Dual Union All
Select 17, '0307', '文件服务器配置-新增','添加一个文件服务器，用于保存客户端升级时需要的升级文件',1,0 From Dual Union All
Select 18, '0307', '文件服务器配置-修改','修改一个文件服务器的基础信息以及启停该服务器',1,0 From Dual Union All
Select 19, '0307', '文件服务器配置-删除','删除一个文件服务器，若该服务器为缺省服务器，则不能删除',1,1 From Dual Union All
Select 20, '0307', '文件升级管理-增加','添加一个第三方部件到数据库中，用于客户端升级时同步将一些第三方部件一起升级',1,0 From Dual Union All
Select 21, '0307', '文件升级管理-修改','修改一个第三方部件的信息，主要是设置其所属系统以及其是否需要注册',1,0 From Dual Union All
Select 22, '0307', '文件升级管理-删除','在数据库中删除一个第三方部件，其并不会影响升级服务器中的文件',1,1 From Dual Union All
Select 23, '0307', '文件升级管理-弃用','在数据库中弃用一个第三方部件，其并不会影响升级服务器中的文件',1,0 From Dual Union All
Select 24, '0307', '上传新的文件','将已经登记的本地有的但是服务器上没有的文件上传到服务器',1,0 From Dual Union All
Select 25, '0307', '上传所有文件','将已经登记的本地所有的文件都上传到服务器',1,0 From Dual Union All
Select 26, '0307', '升级/取消升级','为某个客户端执行升级或取消升级操作',1,0 From Dual Union All
Select 27, '0307', '预升级/取消预升级','为某个客户端执行预升级或取消预升级操作',1,0 From Dual Union All
Select 28, '0307', '全部升级/取消全部升级','对所有客户端执行升级或取消升级操作',1,0 From Dual Union All
Select 29, '0307', '全部预升级/取消全部预升级','对所有客户端进行预升级或取消预升级操作',1,0 From Dual Union All
Select 30, '0308', '修改','修改一个客户端的各项参数',1,1 From Dual Union All
Select 31, '0308', '删除','删除一个指定的客户端，删除后关于该客户端的一切设置都将被清除',1,1 From Dual Union All
Select 32, '0308', '禁用/启用','禁用或启用一个客户端，禁用后该客户端将不能登录本产品',1,0 From Dual Union All
Select 33, '0308', '全部禁用/全部启用','禁用或启用全部客户端，禁用后所有客户端都将不能登录本产品',1,0 From Dual Union All
Select 34, '0308', '清理3个月未登录客户端','清理超过三个月未登录过的客户端',1,1 From Dual Union All
Select 35, '0312', '新增项目','添加一个医院公共信息名称，主要用于医院信息项目的定义',1,0 From Dual Union All
Select 36, '0312', '调整项目','修改一个医院公共信息名称，主要用于医院信息项目的调整',1,1 From Dual Union All
Select 37, '0312', '删除项目','删除一个医院公共信息名称，删除后可能导致医院公共信息的缺失',1,1 From Dual Union All
Select 38, '0314', '日志清理','清理指定天数以外的所有重要操作变动日志',1,1 From Dual Union All
Select 39, '0401', '增加角色','新增一个空白权限的角色',1,0 From Dual Union All
Select 40, '0401', '角色授权','对一个角色进行授权，使其获得一些指定的权限',1,0 From Dual Union All
Select 41, '0401', '删除角色','对角色进行删除操作，该角色以及该角色拥有的权限都将被删除',1,1 From Dual Union All
Select 42, '0401', '复制角色','根据一个角色复制出另一个角色，复制出的角色将拥有和原角色一样的权限信息',1,0 From Dual Union All
Select 43, '0401', '重整所有角色','清除本产品中保存的所有角色，根据用户在数据库中实际拥有的角色重新产生所有角色数据',1,0 From Dual Union All
Select 44, '0401', '恢复所有角色及权限','根据应用系统中保存的所有角色，在数据库中检查并补充创建角色，授予应用系统的公共基础对象权限，以及授予相关数据库对象的访问权限',1,0 From Dual Union All
Select 45, '0401', '修改模块的使用权限','修改应用系统中某些模块的增删改等使用权限',1,0 From Dual Union All
Select 46, '0401', '修改角色的授权用户','将选中角色批量授予某些用户',1,0 From Dual Union All
Select 47, '0402', '批量创建用户','根据所选部门的人员批量创建上机用户，用户默认不授予任何角色权限',1,0 From Dual Union All
Select 48, '0402', '新增用户','新增一个用户，为其绑定人员信息，并授予角色权限',1,0 From Dual Union All
Select 49, '0402', '修改用户','对一个用户绑定的人员以及权限信息进行修改',1,0 From Dual Union All
Select 50, '0402', '删除用户','删除一个用户，同时将释放与该用户绑定的人员及权限',1,1 From Dual Union All
Select 51, '0402', '启停用户','启用或停用一个用户，停用后，该用户将暂时不可使用',1,0 From Dual Union All
Select 52, '0402', '修改密码','修改一个用户的登陆密码',1,0 From Dual Union All
Select 53, '0402', '根据上机人员恢复用户','根据上机用户进行数据恢复，恢复数据之后创建以前的用户并授权，将以用户名作为初始密码',1,0 From Dual Union All
Select 54, '0402', '恢复所有用户角色','根据用户在应用系统中的记录角色重新进行角色授权，恢复角色和用户之后，重建用户的角色',1,0 From Dual Union All
Select 55, '0402', '重整所有用户角色','清除本产品中保存的所有用户的角色，根据用户在数据库中实际拥有的角色重新产生本产品的所有用户的角色数据',1,0 From Dual Union All
Select 56, '0403', '删除','删除一个菜单组，且正在使用该菜单组的导航台将会使用缺省菜单组',1,1 From Dual Union All
Select 57, '0504', '新增','添加一个提醒信息，主要用于在指定时间向一些用户推送某个消息',1,0 From Dual Union All
Select 58, '0504', '修改','修改一个提醒信息，主要是修改提醒内容、提醒条件以及提醒方式等内容',1,0 From Dual Union All
Select 59, '0504', '删除','删除一个提醒信息，删除后对应用户将接收不到该提醒',1,1 From Dual;
Select Zltools.ZlauditlogConfig_ID.Nextval From Dual Connect By Rownum <= (Select Nvl(Max(ID), 0) From Zltools.ZlauditlogConfig);
--117980:高腾,2017-12-26,修正zlauditlog表中的历史数据
Declare
  Cursor c_Auditlog Is
    Select 操作模块编号, 操作内容, Rowid From Zltools.Zlauditlog; --定义游标
Begin
  For c_Item In c_Auditlog Loop
    If To_Number(c_Item.操作模块编号) = 401 Then
      Case
        When Instr(c_Item.操作内容, '新增角色') > 0 Then
          Update Zltools.Zlauditlog Set 日志ID = 1 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.操作内容, '修改角色') > 0 And Instr(c_Item.操作内容, '的权限') > 0 Then
          Update Zltools.Zlauditlog Set 日志ID = 2 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.操作内容 = '对所有角色重新授权' Then
          Update Zltools.Zlauditlog Set 日志ID = 6 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.操作内容, '删除角色') > 0 Then
          Update Zltools.Zlauditlog Set 日志ID = 3 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.操作内容 = '执行操作：修改角色的授权用户' Then
          Update Zltools.Zlauditlog Set 日志ID = 8 Where Zlauditlog.Rowid = c_Item.Rowid;
        Else
          Update Zltools.Zlauditlog Set 日志ID = 1 Where Zlauditlog.Rowid = c_Item.Rowid;
      End Case;
    Else
      Case 
        When Instr(c_Item.操作内容, '新增用户') > 0 Then
          Update Zltools.Zlauditlog Set 日志ID = 10 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.操作内容, '修改用户') > 0 Then
          Update Zltools.Zlauditlog Set 日志ID = 11 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.操作内容, '删除用户') > 0 Then
          Update Zltools.Zlauditlog Set 日志ID = 12 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.操作内容 = '执行操作：批量创建用户' Or c_Item.操作内容 = '执行操作：批量创建用户' Then
          Update Zltools.Zlauditlog Set 日志ID = 9 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.操作内容 = '执行操作：根据上机人员恢复用户' Or c_Item.操作内容 = '执行操作：恢复所有上机人员' Then
          Update Zltools.Zlauditlog Set 日志ID = 15 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.操作内容 = '执行操作：恢复所有用户角色' Then
          Update Zltools.Zlauditlog Set 日志ID = 16 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.操作内容 = '执行操作：重整所有用户角色' Or c_Item.操作内容 = '执行操作：记录所有用户角色' Then
          Update Zltools.Zlauditlog Set 日志ID = 17 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.操作内容, '启用用户') > 0 Or Instr(c_Item.操作内容, '禁用用户') > 0 Then
          Update Zltools.Zlauditlog Set 日志ID = 13 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.操作内容, '修改用户') > 0 And Instr(c_Item.操作内容, '的密码') > 0 Then
          Update Zltools.Zlauditlog Set 日志ID = 14 Where Zlauditlog.Rowid = c_Item.Rowid;
        Else
          Update Zltools.Zlauditlog Set 日志ID = 10 Where Zlauditlog.Rowid = c_Item.Rowid;
      End Case;
    End If;
  End Loop;
End;
/

--117980:高腾,2017-12-06,重要操作变动日志管理界面清理日志过程
Create Or Replace Procedure Zltools.Zl_Zlauditlog_Delete(日志保留天数_In In Number) Is
Begin
  Delete Zlauditlog Where 操作时间 < Sysdate - 日志保留天数_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlauditlog_Delete;
/

--117980:高腾,2017-12-21,自动清理重要操作变动日志
Create Or Replace Procedure Zltools.Zl_Autologprocess As
  --功能： 
  --   对多余的运行日志和错误日志和重要操作变动日志进行清除 
  v_Limit Number;
Begin
  --删除多余的运行日志 
  Select Nvl(Max(To_Number(参数值)), 0) Into v_Limit From zlOptions Where 参数号 = 2;
  Delete From zlDiaryLog Where 进入时间 < Sysdate - v_Limit;

  --删除多余的错误日志 

  Select Nvl(Max(To_Number(参数值)), 0) Into v_Limit From zlOptions Where 参数号 = 4;
  Delete From zlErrorLog Where 时间 < Sysdate - v_Limit;

  --删除多余的重要操作变动日志
  Select Nvl(参数值, 缺省值) Into v_Limit From zlOptions Where 参数号 = 25;
  If v_Limit <> 0 Then
    Delete Zlauditlog Where 操作时间  < Sysdate - v_Limit;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Autologprocess;
/

--117980:高腾,2017-12-06,重要操作变动日志管理界面修改模块记录日志的启停状态
CREATE OR REPLACE Procedure zltools.Zl_Zlauditlogconfig_Update
(
  Id_In         In Zlauditlogconfig.Id%Type,
  是否启用_In   In Zlauditlogconfig.是否启用%Type,
  是否需审核_In In Zlauditlogconfig.是否需审核%Type := Null
) Is
Begin
  If 是否需审核_In Is Null Then
    Update Zlauditlogconfig Set 是否启用 = 是否启用_In Where Id = Id_In;
  Else
    Update Zlauditlogconfig Set 是否需审核 = 是否需审核_In Where Id = Id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlauditlogconfig_Update;
/

--117980:高腾,2017-12-06,重要操作变动日志管理界面插入日志
CREATE OR REPLACE Procedure zltools.Zl_Zlauditlog_Insert
(
  用户名_In   Zlauditlog.用户名%Type,
  工作站_In   Zlauditlog.工作站%Type,
  操作类型_In Zlauditlog.操作类型%Type, --1-新增，2-修改，3-删除
  系统_In     Zlauditlogconfig.系统%Type,
  模块_In     Zlauditlogconfig.模块%Type,
  功能_In     Zlauditlogconfig.功能%Type,
  操作内容_In Zlauditlog.操作内容%Type,
  操作说明_In Zlauditlog.操作说明%Type --用来记录界面提供给操作员输入的备注信息
) Is
  n_是否启用 Zlauditlogconfig.是否启用%Type;
  n_日志Id Zlauditlogconfig.Id%Type;
Begin
  --根据系统编号，模块编号和功能名称查找出当前功能是否开启了记录重要操作变动日志
  If 系统_In Is Null Then
    Select Max(是否启用), Max(Id)
    Into n_是否启用, n_日志Id
    From Zlauditlogconfig
    Where 系统 Is Null And 模块 = 模块_In And 功能 = 功能_In;
  Else
    Select Max(是否启用), Max(Id)
    Into n_是否启用, n_日志Id
    From Zlauditlogconfig
    Where 系统 = 系统_In And 模块 = 模块_In And 功能 = 功能_In;
  End If;
  If n_是否启用 = 1 Then
    Insert Into Zlauditlog
      (用户名, 工作站, 操作时间, 操作类型, 日志Id, 操作内容, 操作说明)
    Values
      (用户名_In, 工作站_In, Sysdate, 操作类型_In, n_日志Id, 操作内容_In, 操作说明_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlauditlog_Insert;
/

--118170:高腾,2017-12-14,杀进程管理界面
Create Or Replace Procedure Zltools.Zl_Zlkillprocess_Edit
(
  操作_In In Number, --1:增加;2:修改;3:删除
  序号_In In Zlkillprocess.序号%Type,
  名称_In In Zlkillprocess.名称%Type := Null,
  类型_In In Zlkillprocess.类型%Type := Null,
  描述_In In Zlkillprocess.描述%Type := Null
) As
  n_序号 Zlkillprocess.序号%Type;
Begin
  If 操作_In = 1 Then
    --获取最大序号
    Select Max(序号) + 1 Into n_序号 From Zlkillprocess;
    --插入数据
    Insert Into Zlkillprocess (序号, 名称, 类型, 描述, 是否固定) Values (n_序号, 名称_In, 类型_In, 描述_In, 0);
  Elsif 操作_In = 2 Then
    Update Zlkillprocess Set 名称 = 名称_In, 类型 = 类型_In, 描述 = 描述_In Where 序号 = 序号_In;
  Else
    Delete Zlkillprocess Where 名称 = 名称_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlkillprocess_Edit;
/

--117980:高腾,2017-12-26,废除操作模块编号字段并添加主键即外键约束
Alter Table Zltools.zlauditlog Rename Column 操作模块编号 To 操作模块编号_bak;
Alter Table Zltools.zlAuditLog Drop Constraint zlAuditLog_PK Cascade Drop Index;
Alter Table Zltools.zlAuditLog Add Constraint zlAuditLog_PK Primary Key (操作时间,用户名,工作站,日志ID) Using Index;
Alter Table Zltools.zlAuditLog Add Constraint zlAuditLog_FK_日志ID Foreign Key(日志ID) References ZlauditlogConfig(ID) On Delete Cascade;
Create Index Zltools.Zlauditlog_IX_日志ID On Zlauditlog(日志ID);

--116688:杨周一,2017-12-25,登录控制相关表结构修正(用户名字符长度修改)
Alter Table zltools.zlAppPermission Modify (用户名 varchar2(20));
Alter Table zltools.zlLoginLimit Modify (用户名 varchar2(20));

--117833;杨周一,2017-12-25,管理工具对象审计管理
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明,次序) Values('0610','06','对象审计管理','O',Null,10);
--111882:余智勇,2017-12-27,标签水平反转显示
Alter Table zlTools.zlRPTItems Add 水平反转 Number(1);


--119139:杨周一,2017-12-28,远程用户和密码保存
Create Or Replace Procedure Zltools.Zl_Zlclients_Set
(
  n_Mode_In       Number,
  n_Rowid_In      Varchar2 := Null,
  v_工作站_In     Zlclients.工作站%Type := Null,
  v_Ip_In         Zlclients.Ip%Type := Null,
  v_Cpu_In        Zlclients.Cpu%Type := Null,
  v_内存_In       Zlclients.内存%Type := Null,
  v_硬盘_In       Zlclients.硬盘%Type := Null,
  v_操作系统_In   Zlclients.操作系统%Type := Null,
  v_部门_In       Zlclients.部门%Type := Null,
  v_用途_In       Zlclients.用途%Type := Null,
  v_说明_In       Zlclients.说明%Type := Null,
  n_升级服务器_In Zlclients.升级服务器%Type := Null,
  n_升级标志_In   Zlclients.升级标志%Type := 0,
  n_连接数_In     Zlclients.连接数%Type := 0,
  v_站点_In       Zlclients.站点%Type := Null,
  n_Apply_In      Number := 0,
  v_Ipbegin_In    Varchar2 := Null,
  v_Ipend_In      Varchar2 := Null,
  n_启用视频源    Zlclients.启用视频源%Type := Null,
  v_管理员用户_In Zlclients.管理员用户%Type := Null,
  v_管理员密码_In Zlclients.管理员密码%Type := Null
  --功能：新增客户端或站点 或者更新客户端属性
  --应用：1、管理工具：新增或修改站点 （修改时以IP与客户端做判断条件，不需传入N_Rowid_In）
  --      2：应用系统：登录时根据当前登录的客户短来判断是否
  --                   新增站点或修改站点参数（更新时N_Rowid_In需传入）
  --站点设置:0-新增站点，1-更新站点
  --N_Apply_In,站点参数应用范围，0-本站点，1，本部门，2，所有站点，3，固定IP段
  --V_Ipbegin_In,V_Ipend_In:在固定IP断应用时传入,两者在一个IP断上，即前面部分相同
) Is
  n_Pos         Number(3);
  n_Ipbegin_Num Number;
  n_Ipend_Num   Number;
  n_Ip_Num      Number;
  n_Count       Number;

  v_Err Varchar2(500);
  Err_Custom Exception;

  Function Get_Ipnum(v_Ip_Input Varchar2) Return Number Is
    v_Ip_Num  Varchar2(20);
    n_Pos_Tmp Number;
    v_Ip_Tmp  Varchar2(20);
  Begin
    n_Pos_Tmp := Length(v_Ip_Input);
    n_Pos_Tmp := n_Pos_Tmp - Length(Replace(v_Ip_Input, '.', ''));
    If n_Pos_Tmp <> 3 Then
      Return Null;
    Else
      v_Ip_Tmp := v_Ip_Input;
      Loop
        n_Pos_Tmp := Instr(v_Ip_Tmp, '.');
        Exit When(Nvl(n_Pos_Tmp, 0) = 0);
        --将每一断数字转化为3位数
        v_Ip_Num := v_Ip_Num || Trim(To_Char(Substr(v_Ip_Tmp, 1, n_Pos_Tmp - 1), '099'));
        v_Ip_Tmp := Substr(v_Ip_Tmp, n_Pos_Tmp + 1);
      End Loop;
      v_Ip_Num := v_Ip_Num || Trim(To_Char(v_Ip_Tmp, '099'));
      n_Ip_Num := To_Number(Trim(v_Ip_Num));
      Return n_Ip_Num;
    End If;
  End;
Begin
  If n_Mode_In = 0 Then
  
    Select Count(1) Into n_Count From zlClients Where 工作站 = v_工作站_In;
    If n_Count = 0 Then
      Insert Into zlClients
        (Ip, 工作站, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级服务器, 升级标志, 连接数, 站点, 启用视频源, 最近登陆时间, 管理员用户, 管理员密码)
      Values
        (v_Ip_In, v_工作站_In, v_Cpu_In, v_内存_In, v_硬盘_In, v_操作系统_In, v_部门_In, v_用途_In, v_说明_In, n_升级服务器_In, n_升级标志_In,
         n_连接数_In, v_站点_In, n_启用视频源, Sysdate, v_管理员用户_In, v_管理员密码_In);
    Else
      v_Err := '已经设置了相同IP地址或工作站,不能再设!';
      Raise Err_Custom;
    End If;
  Else
    If n_Rowid_In Is Null Then
      Update zlClients
      Set Cpu = v_Cpu_In, 内存 = v_内存_In, 硬盘 = v_硬盘_In, 操作系统 = v_操作系统_In, 部门 = v_部门_In, 用途 = v_用途_In, 说明 = v_说明_In,
          连接数 = n_连接数_In, 站点 = v_站点_In, 启用视频源 = n_启用视频源, 升级服务器 = n_升级服务器_In, 升级标志 = n_升级标志_In, 最近登陆时间 = Sysdate,
          管理员用户 = Decode(v_管理员用户_In, '空空', Null, Nvl(v_管理员用户_In, 管理员用户)),
          管理员密码 = Decode(v_管理员密码_In, '空空', Null, Nvl(v_管理员密码_In, 管理员密码))
      Where 工作站 = v_工作站_In And Ip = v_Ip_In;
    Else
      Update zlClients
      Set 工作站 = v_工作站_In, Ip = v_Ip_In, Cpu = Decode(v_Cpu_In, Null, Cpu, v_Cpu_In),
          内存 = Decode(v_内存_In, Null, 内存, v_内存_In), 硬盘 = Decode(v_硬盘_In, Null, 硬盘, v_硬盘_In),
          操作系统 = Decode(v_操作系统_In, Null, 操作系统, v_操作系统_In), 部门 = v_部门_In, 站点 = v_站点_In, 启用视频源 = n_启用视频源, 最近登陆时间 = Sysdate,
          管理员用户 = Decode(v_管理员用户_In, '空空', Null, Nvl(v_管理员用户_In, 管理员用户)),
          管理员密码 = Decode(v_管理员密码_In, '空空', Null, Nvl(v_管理员密码_In, 管理员密码))
      Where Rowid = n_Rowid_In;
    End If;
  End If;
  --本部门
  If n_Apply_In = 1 Then
    Update zlClients
    Set 连接数 = n_连接数_In, 站点 = v_站点_In
    Where Nvl(部门, 'NONE') = Nvl(v_部门_In, 'NONE') And Ip <> v_Ip_In;
  Elsif n_Apply_In = 2 Then
    Update zlClients Set 连接数 = n_连接数_In, 站点 = v_站点_In Where Ip <> v_Ip_In;
  Elsif n_Apply_In = 3 Then
    n_Pos := Length(v_Ipbegin_In);
    n_Pos := n_Pos - Length(Replace(v_Ipbegin_In, '.', ''));
    If n_Pos <> 3 Then
      v_Err := '起始IP格式有误！';
      Raise Err_Custom;
    End If;
    n_Pos := Length(v_Ipend_In);
    n_Pos := n_Pos - Length(Replace(v_Ipend_In, '.', ''));
    If n_Pos <> 3 Then
      v_Err := '结束IP格式有误！';
      Raise Err_Custom;
    End If;
  
    n_Ipbegin_Num := Get_Ipnum(v_Ipbegin_In);
    n_Ipend_Num   := Get_Ipnum(v_Ipend_In);
    For r_Ip In (Select 工作站, Ip From zlClients) Loop
      n_Ip_Num := Get_Ipnum(r_Ip.Ip);
      If n_Ip_Num >= n_Ipbegin_Num And n_Ip_Num <= n_Ipend_Num Then
        Update zlClients Set 连接数 = n_连接数_In, 站点 = v_站点_In Where 工作站 = r_Ip.工作站 And Ip = r_Ip.Ip;
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Set;
/
--119449:刘硕,2018-01-05,基本平台消息完善
Insert Into ZLTOOLS.zlOptions(参数号,参数名,参数值,缺省值,参数说明) Values(26, '集成平台消息保留天数', '7', '7', '供集成平台使用的业务消息数据最多能保留的天数。当设置为0时，自动保留最近7天的消息数据');

--120029:高腾,2018-01-12,限时方案管理功能
Insert Into Zltools.Zlsvrtools (编号, 上级, 标题, 快键, 说明, 次序) Values ('0315', '03', '功能限时管理', 'B', Null, 20);

--120029:高腾,2018-01-12,限时方案管理功能
Create Table Zltools.ZlRunLimit(
       序号 Number(3),
       名称 Varchar2(50),
       是否启用 Number(1),
       描述 Varchar2(250)
);
Alter Table Zltools.ZlRunLimit Add Constraint ZlRunLimit_PK Primary Key(序号) Using Index;
Alter Table Zltools.ZlRunLimit Add Constraint ZlRunLimit_UQ_名称 Unique(名称) Using Index;
Insert Into Zltools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRUNLIMIT','ZLTOOLSTBS','A2');

--120029:高腾,2018-01-12,限时方案管理功能
Create Table Zltools.ZlRunLimitTime(
       ID Number(18),
       方案 Number(3),
       星期 Number(1),
       开始时间 date,
       结束时间 date
);
Alter Table Zltools.ZlRunLimitTime Add Constraint ZlRunLimitTime_PK Primary Key(ID) Using Index;
Alter Table Zltools.ZlRunLimitTime Add Constraint ZlRunLimitTime_UQ_方案时段 Unique(方案,星期,开始时间,结束时间) Using Index;
Alter Table Zltools.ZlRunLimitTime Add Constraint ZlRunLimitTime_FK_方案 Foreign Key(方案) References ZlRunLimit(序号) On Delete Cascade;
Create Sequence Zltools.ZlRunLimitTime_ID start with 1;
Insert Into Zltools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRUNLIMITTIME','ZLTOOLSTBS','A2');

--120029:高腾,2018-01-12,限时方案管理功能
Create Table Zltools.ZlRunLimitSet(
       序号 Number(5),
       系统 Number(5),
       模块 Varchar2(18),
       功能 Varchar2(50),
       操作选项 Number(1),
       方案序号 Number(3),
       限时原因 Varchar2(250)
);
Alter Table Zltools.ZlRunLimitSet Add Constraint ZlRunLimitSet_PK Primary Key(序号) Using Index;
Alter Table Zltools.ZlRunLimitSet Add Constraint ZlRunLimitSet_UQ_模块功能 Unique(系统,模块,功能) Using Index;
Alter Table Zltools.ZlRunLimitSet Add Constraint ZlRunLimitSet_FK_方案序号 Foreign Key(方案序号) References ZlRunLimit(序号) On Delete Cascade;
Alter Table Zltools.ZlRunLimitSet Add Constraint ZlRunLimitSet_FK_系统 Foreign Key(系统) References ZlSystems(编号) On Delete Cascade;
CREATE INDEX Zltools.ZlRunLimitSet_IX_方案序号 ON ZlRunLimitSet(方案序号);
Insert Into Zltools.zlTables(系统,表名,表空间,分类) Values(0,'ZLRUNLIMITSET','ZLTOOLSTBS','A2');

--120029:高腾,2018-01-12,限时方案管理功能ZlRunLimit预设数据
Insert Into Zltools.ZlRunLimit(序号,名称,是否启用,描述) Values(1,'预设方案',1,'');

--120029:高腾,2018-01-12,限时方案管理功能ZlRunLimitTime预设数据
Insert Into Zltools.ZlRunLimitTime(ID,方案,星期,开始时间,结束时间) 
Select 1,1,0,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 2,1,1,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 3,1,2,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 4,1,3,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 5,1,4,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 6,1,5,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 7,1,6,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual;
Select Zltools.ZlRunLimitTime_ID.Nextval From Dual Connect By Rownum <= (Select Nvl(Max(ID), 0) From Zltools.ZlRunLimitTime);

--120029:高腾,2018-01-12,限时方案管理功能ZlRunLimitSet预设数据
Insert Into Zltools.ZlRunLimitSet(序号,模块,功能,操作选项,方案序号,限时原因)
Select 1,'0102','提前升迁',1,1,'提前升迁在一定程度会影响正常业务开展' From Dual Union All
Select 2,'0401','角色授权',1,1,'角色授权会对基础或公共对象重新授权，导致相关SQL重新解析，影响正常业务运行' From Dual Union All
Select 3,'0401','恢复所有角色及权限',1,1,'恢复所有角色及权限会对基础或公共对象重新授权，导致相关SQL重新解析，影响正常业务运行' From Dual Union All
Select 4,'0401','复制角色',1,1,'复制角色会对基础或公共对象重新授权，导致相关SQL重新解析，影响正常业务运行' From Dual Union All
Select 5,'0402','恢复所有用户角色',1,1,'恢复所有用户角色会对基础或公共对象重新授权，导致相关SQL重新解析，影响正常业务运行' From Dual Union All
Select 6,'0402','重整所有用户角色',1,1,'重整用户角色会重新生成角色权限控制数据，影响业务的正常运行' From Dual;

--120029:高腾,2018-01-12,限时方案管理功能修改模块功能对应方案及操作选项的过程
CREATE OR REPLACE Procedure zltools.Zl_Zlrunlimitset_Update
(
  序号_In     In Zlrunlimitset.序号%Type,
  方案_In     In Zlrunlimitset.方案序号%Type := Null,
  操作选项_In In Zlrunlimitset.操作选项%Type := Null,
  限时原因_In In Zlrunlimitset.限时原因%Type := Null
) As
Begin
  If 方案_In Is Null Then
    Update Zlrunlimitset Set 方案序号 = 方案_In Where 序号 = 序号_In;
  Else
    Update Zlrunlimitset Set 方案序号 = 方案_In, 操作选项 = 操作选项_In, 限时原因 = 限时原因_In Where 序号 = 序号_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlrunlimitset_Update;
/

--120029:高腾,2018-01-12,限时方案管理功能修改时间段
CREATE OR REPLACE Procedure zltools.Zl_Zlrunlimittime_Update
(
  操作_In     In Number, --0:新增，1:修改,2:删除
  Id_In       In Zlrunlimittime.Id%Type,
  方案_In     In Zlrunlimittime.方案%Type := Null,
  星期_In     In Zlrunlimittime.星期%Type := Null,
  开始时间_In In Zlrunlimittime.开始时间%Type := Null,
  结束时间_In In Zlrunlimittime.结束时间%Type := Null
) As
  n_Count   Number;
  d_Maxtime Date;
  d_Mintime Date;
Begin
  --检查当前设置时间是否与已有时间有冲突
  --若有，则记录冲突最大区间，并删掉冲突时间段
  If 操作_In = 0 Or 操作_In = 1 Then
    Select Count(1)
    Into n_Count
    From Zlrunlimittime
    Where 开始时间 <= 开始时间_In And 结束时间 >= 结束时间_In And 方案 = 方案_In And 星期 = 星期_In And ID <> Id_In;
    If n_Count = 0 Or (n_Count <> 0 And Id_In <> 0) Then
      Select Min(开始时间), Max(结束时间), Count(1)
      Into d_Mintime, d_Maxtime, n_Count
      From Zlrunlimittime
      Where (开始时间 >= 开始时间_In And 开始时间 <= 结束时间_In Or 结束时间 <= 结束时间_In And 结束时间 >= 开始时间_In Or
            结束时间 >= 结束时间_In And 开始时间 <= 开始时间_In) And 方案 = 方案_In And 星期 = 星期_In And ID <> Id_In;
      If 开始时间_In < d_Mintime Then
        d_Mintime := 开始时间_In;
      End If;
      If 结束时间_In > d_Maxtime Then
        d_Maxtime := 结束时间_In;
      End If;
      If n_Count > 0 Then
        --说明有冲突的字段
        --先将冲突字段删除，再插入新字段
        If Id_In <> 0 Then
          Delete Zlrunlimittime Where ID = Id_In;
        End If;
        Delete Zlrunlimittime
        Where (开始时间 >= 开始时间_In And 开始时间 <= 结束时间_In Or 结束时间 <= 结束时间_In And 结束时间 >= 开始时间_In Or
              结束时间 >= 结束时间_In And 开始时间 <= 开始时间_In) And 方案 = 方案_In And 星期 = 星期_In;
        Insert Into Zlrunlimittime
          (ID, 方案, 星期, 开始时间, 结束时间)
        Values
          (Zlrunlimittime_Id.Nextval, 方案_In, 星期_In, d_Mintime, d_Maxtime);
      Else
        --说明没有冲突，将直接对数据进行插入或更新操作
        If 操作_In = 0 Then
          --新增
          Insert Into Zlrunlimittime
            (ID, 方案, 星期, 开始时间, 结束时间)
          Values
            (Zlrunlimittime_Id.Nextval, 方案_In, 星期_In, 开始时间_In, 结束时间_In);
        Else
          --修改
          Update Zlrunlimittime Set 开始时间 = 开始时间_In, 结束时间 = 结束时间_In Where ID = Id_In;
        End If;
      End If;
    End If;
  Else
    --删除
    Delete Zlrunlimittime Where ID = Id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlrunlimittime_Update;
/

--120029:高腾,2018-01-12,限时方案管理功能修改方案
Create Or Replace Procedure Zltools.Zl_Zlrunlimit_Update
(
  操作_In     In Number, --0:新增，1:修改,2:删除
  序号_In     In Zlrunlimit.序号%Type,
  名称_In     In Zlrunlimit.名称%Type := Null,
  是否启用_In In Zlrunlimit.是否启用%Type := Null,
  描述_In     In Zlrunlimit.描述%Type := Null
) As
  n_序号 Zlrunlimit.序号%Type;
Begin
  If 操作_In = 0 Then
    --新增
    Select Max(序号) + 1 Into n_序号 From Zlrunlimit;
    Insert Into Zlrunlimit (序号, 名称, 是否启用, 描述) Values (n_序号, 名称_In, 是否启用_In, 描述_In);
    Insert Into Zlrunlimittime
      (ID, 方案, 星期, 开始时间, 结束时间)
      Select Zlrunlimittime_Id.Nextval, n_序号, 星期, 开始时间, 结束时间 From Zlrunlimittime Where 方案 = 1;
  Elsif 操作_In = 1 Then
    --修改
    If 是否启用_In Is Null Then
      --修改方案信息功能
      Update Zlrunlimit Set 名称 = 名称_In, 描述 = 描述_In Where 序号 = 序号_In;
    Else
      --启停方案功能
      Update Zlrunlimit Set 是否启用 = 是否启用_In Where 序号 = 序号_In;
    End If;
  Else
    --删除
    Delete Zlrunlimit Where 序号 = 序号_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlrunlimit_Update;
/

--118267:杨周一,2017-12-15,管理工具LIS图片转存功能
Insert Into zlTools.Zlsvrtools(编号,上级,标题,快键,说明,次序) Values('0208','02','检验图片数据转移','U',Null,22);

--121074:高腾,2017-1-26,调整模块名称
Update zlTools.zlSvrTools Set 标题 = '历史数据空间管理' Where 编号 = '0201';

--116852:杨周一,2018-02-27,删除原有DBA工具
Insert Into Zltools.Zlfilesexpired(文件名, 安装路径, 系统编号, 系统版本, 说明)Values('ZLDBAToolsEXE.exe', '[APPSOFT]', Null, '10.35.90', '名称修正,删除原有部件');
Delete From zlFilesUpgrade Where Upper(文件名) = 'ZLDBATOOLSEXE.EXE';
Delete From Zlfiles Where Upper(名称) = 'ZLDBATOOLSEXE.EXE';