DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 启用=1 WHERE 新部件='zl9TemperatureChartSD';
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('山东地区专用部件','适用于山东地区','zl9BodyEditor',1,'zl9TemperatureChartSD');
    END IF;
  END;
END; 
/
COMMIT;


