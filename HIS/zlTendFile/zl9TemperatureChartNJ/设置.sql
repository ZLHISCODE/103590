DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 启用=1 WHERE 新部件='zl9TemperatureChartNJ';
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('江苏地区专用部件','适用于江苏地区','zl9BodyEditor',1,'zl9TemperatureChartNJ');
    END IF;
  END;
END; 
/
COMMIT;


