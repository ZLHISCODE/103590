DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 新部件='zl9TemperatureChartGD',启用=1 WHERE 部件='zl9BodyEditorGD';
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('广东体温部件','适用于广东省','zl9BodyEditorGD',1,'zl9TemperatureChartGD');
    END IF;
  END;
END; 
/
COMMIT;



