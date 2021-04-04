--使用前将XXXXX替换为要删除的固定报表(组)编号

--删除zlPrograms->zlProgFuncs->zlProgPrivs
Delete From zlPrograms
Where Upper(部件)=Upper('zl9Report') 
	And (系统,序号) IN(
		Select 系统,程序ID From zlReports A Where 系统=100 And A.编号='XXXXX'
		Union ALL
		Select 系统,程序ID From zlRPTGroups A Where 系统=100 And A.编号='XXXXX'
		);

--删除zlMenus
Delete From zlMenus
Where 系统=100 And 模块 IN(
	Select A.程序ID From zlReports A,zlPrograms B
	Where A.系统=B.系统 And A.程序ID=B.序号 And Upper(B.部件)=Upper('zl9Report') And A.系统=100 And A.编号='XXXXX'
	Union ALL
	Select A.程序ID From zlRPTGroups A,zlPrograms B
	Where A.系统=B.系统 And A.程序ID=B.序号 And Upper(B.部件)=Upper('zl9Report') And A.系统=100 And A.编号='XXXXX'
	);

--仅模块内报表：删除zlProgFuncs->zlProgPrivs
Delete From zlProgFuncs
Where (系统,序号,功能) IN(
	Select A.系统,A.程序ID,A.说明 From zlReports A,zlPrograms B
	Where A.系统=B.系统 And A.程序ID=B.序号 And Upper(B.部件)<>Upper('zl9Report') And A.系统=100 And A.编号='XXXXX'
	Union ALL
	Select B.系统,B.程序ID,B.功能 From zlReports A,zlRPTPuts B Where A.ID=B.报表ID And A.系统=100 And A.编号='XXXXX'
	);

--删除zlReport->zlRPTSubs、zlRPTPuts、zlRPTItems、zlRPTDatas ...
Delete From zlReports A Where 系统=100 And A.编号='XXXXX';
Delete From zlRPTGroups A Where 系统=100 And A.编号='XXXXX';