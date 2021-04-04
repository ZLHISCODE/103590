Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1288, 0, 0, 0, 0, 19,'XW关键图像地址', 0, 'http://127.0.0.1:8080/KeyImage.aspx?colid0=22&'||'colvalue0=[@STU_NO]', 'XW PACS WEB服务器的地址。'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1288 And 参数名 = 'XW关键图像地址');
  
--85463:许华峰,2015-07-01,检查列表按影像类别和检查部位过滤
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1290, 1, 1, 0, 1, 49,'影像类别过滤', 0, 0, '检查列表数据按影像类别过滤'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1290 And 参数名 = '影像类别过滤');
  
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1291, 1, 1, 0, 1, 53,'影像类别过滤', 0, 0, '检查列表数据按影像类别过滤'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1291 And 参数名 = '影像类别过滤');
  
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1294, 1, 1, 0, 1, 109,'影像类别过滤', 0, 0, '检查列表数据按影像类别过滤'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1294 And 参数名 = '影像类别过滤');
  
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1290, 1, 1, 0, 1, 50,'检查部位过滤', 0, 0, '检查列表数据按检查部位过滤'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1290 And 参数名 = '检查部位过滤');
  
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1291, 1, 1, 0, 1, 54,'检查部位过滤', 0, 0, '检查列表数据按检查部位过滤'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1291 And 参数名 = '检查部位过滤');
  
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1294, 1, 1, 0, 1, 110,'检查部位过滤', 0, 0, '检查列表数据按检查部位过滤'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1294 And 参数名 = '检查部位过滤'); 
  
--84345:许华峰,2015-06-30,是否发送WorkList
Alter Table 影像检查记录 Add 是否安排 Number(1);

--84345:许华峰,2015-06-30,是否发送WorkList
Insert Into zlProgPrivs
  (系统, 序号, 功能, 所有者, 对象, 权限)
  Select &n_System, 1290, '基本', User, 'ZL_影像检查记录_发送安排', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where 系统 = &n_System And 序号 = 1290 And 功能 = '基本' And Upper(对象) = Upper('ZL_影像检查记录_发送安排'));
		 
--84345:许华峰,2015-06-30,是否发送WorkList
CREATE OR REPLACE Procedure ZL_影像检查记录_发送安排
( 
  医嘱ID_In       影像检查记录.医嘱ID%Type,
  发送号_In       影像检查记录.发送号%Type,
  是否安排_In     影像检查记录.是否安排%Type,
  检查技师_In     影像检查记录.检查技师%Type, 
  检查技师二_In   影像检查记录.检查技师二%Type, 
  执行间_In       病人医嘱发送.执行间%Type
) As 
Begin 
 
  Update 影像检查记录 
  Set    检查技师 = 检查技师_In, 检查技师二 = 检查技师二_In, 是否安排 = 是否安排_In
  Where  医嘱ID = 医嘱ID_In and 发送号 =发送号_In; 
 
  Update 病人医嘱发送 
  Set 执行间 = 执行间_In
  Where 医嘱ID=医嘱ID_In and 发送号=发送号_In; 
Exception 
  When Others Then 
    Zl_Errorcenter(Sqlcode, Sqlerrm); 
End ZL_影像检查记录_发送安排;
/

--85773:许华峰,2015-06-29,XW3D观片
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1288, 0, 0, 0, 0, 17,'XWWEB观片地址', 0, 'http://127.0.0.1:8080/imageweb/imageAction.action?ColID0=22&ColValue0=[@STU_NO]', 'XW PACS WEB服务器的地址。'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1288 And 参数名 = 'XWWEB观片地址');
  
Insert Into Zlparameters
  (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)
  Select Zlparameters_Id.Nextval, &n_System, 1288, 0, 0, 0, 0, 18,'XW3D观片类型', 0, 'Study3D', 'XW PACS 3D观片时的观片类型，“Study3D”为直接打开该检查的图像，第一个序列加载3D，“SeriesList3D” 为打开该检查的序列，由用户选择加载某个序列的3D后打开'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where 系统 = &n_System And 模块 = 1288 And 参数名 = 'XW3D观片类型');