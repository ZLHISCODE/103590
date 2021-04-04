DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 新部件='zl9TemperatureChartJX',启用=1 WHERE 部件='zl9BodyEditorJX';
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('江西体温部件','适用于江西省','zl9BodyEditorJX',1,'zl9TemperatureChartJX');
    END IF;
  END;
END; 
/
Update Zlparameters
Set 参数值=20
Where 参数名 = '体温曲线固定添加行数' And Nvl(模块, 0) = 1255 And Nvl(系统, 0) = 100 and 参数号 = 76 ;
/
Update 体温记录项目
Set 记录法 = 1, 记录符 = '○', 记录色 = '16744448', 刻度间隔 = '10.00000', 单位值 = '1.00000', 记录频次 = 2, 单位 = '次/分', 最高行 = '0'
Where 项目序号 = 3;
/ 
COMMIT;

