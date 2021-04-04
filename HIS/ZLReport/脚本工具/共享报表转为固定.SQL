--使用说明：
--1.在共享系统中编制好报表(组)，并根据规范进行相应的命名
--2.如果是导航台报表(组)，则直接发布到相应的导航台菜单位置
--  如果是模块内使用的报表或票据，则将想要的内部授权功能填写到报表说明中，再发布到该模块所在的导航台菜单位置
--3.以ZLTOOLS用户登录执行该脚本，该脚本根据命名区别不同类型的报表，并对权限等作相应的调整
--4.运行该脚本后，报表在数据库中即和正式安装的情况一致，可直接生成报表安装脚本
--5.如果某些报表需要在多个模块内使用，又不想做成重复的报表，除第一个模块按以上方法操作外，第二个之后的模块在执行前面的步骤后，再按以下操作：
--  a.在具体的某个系统中，选中已调整为固定报表的，要在其它模块内重复使用的报表
--  b.输入"publish report"，将自动调出"发布到模块菜单"功能，在该功能中选择相应的模块进行发布
--  c.如果操作有误，可输入"unpublish report"自动调出"从模块菜单取消"发布功能
--  d.此时再去生成报表安装脚本，就会包含多个模块的权限等相应数据。

Create Or Replace Procedure AdjustReport(SYS_IN zlSystems.编号%TYPE) as
--功能：对共享系统中的报表进行处理，包含特殊属性，发布了的权限部分。
--      某些部分(如具体功能)或能需要人为再调整。
	Cursor c_Report is 
		Select * From zlReports Where Upper(编号) Like 'ZL%\_%\_%' Escape '\' And Nvl(系统,0)=0;
	Cursor c_Group is 
		Select * From zlRPTGroups Where Upper(编号) Like 'ZL%\_%\_%' Escape '\' And Nvl(系统,0)=0;

	v_序号	Number;
	v_编号	zlReports.编号%TYPE;
Begin
	--标题为固定元素
	Update zlRPTItems  Set 系统=1 Where 类型=2 And (名称='标题' OR 内容 Like '%[单位名称]%');
	--票据
	Update zlReports Set 票据=1 Where Upper(编号) Like '%\_BILL\_%' Escape '\';

	--编号规则
	--1．菜单独立表：ZL1_Report_程序号
	--2．菜单报表组：ZL1_Group_程序号
	--3．报表组子表：ZL1_Sub_程序号_序号
	--4．模块内报表：ZL1_Inside_程序号(有多个则加”_序号”)
	--5．模块内票据：ZL1_Bill_程序号(有多张则加”_序号”)

	--报表
	For r_Report In c_Report Loop
		--根据报表编号确定系统、程序号、报表类型
		v_编号:=Substr(r_Report.编号,1,Instr(r_Report.编号,'_')-1);
		
		v_编号:=Substr(r_Report.编号,Instr(r_Report.编号,'_')+1);
		v_编号:=Substr(v_编号,Instr(v_编号,'_')+1);
		IF Instr(v_编号,'_')>0 Then
			v_序号:=To_Number(Substr(v_编号,1,Instr(v_编号,'_')-1));
		Else
			v_序号:=To_Number(v_编号);
		End IF;

		Update zlReports Set 系统=SYS_IN Where ID=r_Report.ID;

		--已发布的(不含Sub)
		IF r_Report.程序ID is Not NULL Then
			Update zlReports Set 程序ID=v_序号 Where ID=r_Report.ID;

			IF Not (Upper(r_Report.编号) Like '%INSIDE%' OR Upper(r_Report.编号) Like '%BILL%') Then
				--zlPrograms
				Update zlPrograms Set 系统=SYS_IN,序号=v_序号 Where 序号=r_Report.程序ID and 系统 is NULL;
				--zlProgFuncs
				Update zlProgFuncs Set 系统=SYS_IN,序号=v_序号 Where 序号=r_Report.程序ID and 系统 is NULL;
				--zlProgPrivs
				Update zlProgPrivs Set 系统=SYS_IN,序号=v_序号 Where 序号=r_Report.程序ID and 系统 is NULL;
				--zlMenus
				Update zlMenus Set 系统=SYS_IN,模块=v_序号 Where 模块=r_Report.程序ID And 系统 is NULL And 组别='缺省';

				--！菜单独立项需要增加其它功能的在之后人为处理,如收费日报的:所有操作员,全院收入的:所有科室
				IF r_Report.名称 Like '%日报' Then
					Insert Into zlProgFuncs(系统,序号,功能) Values(SYS_IN,v_序号,'所有操作员');
				End IF;
			Else
				--模块内部报表或票据的"功能"定为其说明
				Update zlReports Set 功能=说明 Where ID=r_Report.ID;
				--zlProgFuncs
				Update zlProgFuncs Set 系统=SYS_IN,序号=v_序号,功能=r_Report.说明 Where 序号=r_Report.程序ID And 系统 is NULL;
				--zlProgPrivs
				Update zlProgPrivs Set 系统=SYS_IN,序号=v_序号,功能=r_Report.说明 Where 序号=r_Report.程序ID And 系统 is NULL;
				
				--模块内部报表或票据已有固定程序项
				Delete From zlPrograms Where 序号=r_Report.程序ID And 系统 is NULL;

				--模块内部报表或票据不需要菜单
				Delete From zlMenus Where 模块=r_Report.程序ID And 系统 is NULL And 组别='缺省';				
			End IF;
		End IF;
	End Loop;

	--报表组
	For r_Group In c_Group Loop
		v_编号:=Substr(r_Group.编号,1,Instr(r_Group.编号,'_')-1);
		
		v_编号:=Substr(r_Group.编号,Instr(r_Group.编号,'_')+1);
		v_编号:=Substr(v_编号,Instr(v_编号,'_')+1);
		IF Instr(v_编号,'_')>0 Then
			v_序号:=To_Number(Substr(v_编号,1,Instr(v_编号,'_')-1));
		Else
			v_序号:=To_Number(v_编号);
		End IF;

		Update zlRPTGroups Set 系统=SYS_IN Where ID=r_Group.ID;

		--已发布的组
		IF r_Group.程序ID is Not NULL Then
			Update zlRPTGroups Set 程序ID=v_序号 Where ID=r_Group.ID;

			--zlPrograms
			Update zlPrograms Set 系统=SYS_IN,序号=v_序号 Where 序号=r_Group.程序ID And 系统 is NULL;
			--zlProgFuncs
			Update zlProgFuncs Set 系统=SYS_IN,序号=v_序号 Where 序号=r_Group.程序ID And 系统 is NULL;
			--zlProgPrivs
			Update zlProgPrivs Set 系统=SYS_IN,序号=v_序号 Where 序号=r_Group.程序ID And 系统 is NULL;
			--zlMenus
			Update zlMenus Set 模块=v_序号,系统=SYS_IN Where 模块=r_Group.程序ID And 系统 is NULL And 组别='缺省';
		End IF;
	End Loop;
End;
/

--删除处键
ALTER TABLE zlProgFuncs Drop CONSTRAINT zlProgFuncs_FK_序号;
ALTER TABLE zlProgPrivs Drop CONSTRAINT zlProgPrivs_FK_序号;
ALTER TABLE zlMenus Drop CONSTRAINT zlMenus_FK_模块;

Execute AdjustReport(100);
Drop Procedure AdjustReport;

--恢复处键
ALTER TABLE zlProgFuncs ADD CONSTRAINT zlProgFuncs_FK_序号 FOREIGN KEY (系统,序号) REFERENCES zlPrograms(系统,序号) ON DELETE CASCADE;
ALTER TABLE zlProgPrivs ADD CONSTRAINT zlProgPrivs_FK_序号 FOREIGN KEY (系统,序号,功能) REFERENCES zlProgFuncs(系统,序号,功能) ON DELETE CASCADE;
ALTER TABLE zlMenus ADD CONSTRAINT zlMenus_FK_模块 FOREIGN KEY (系统,模块) REFERENCES zlPrograms(系统,序号) ON DELETE CASCADE;
