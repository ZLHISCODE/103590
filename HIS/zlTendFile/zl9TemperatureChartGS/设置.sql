DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 新部件='zl9TemperatureChartGS',启用=1 WHERE 部件='zl9BodyEditorGS';
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('甘肃中医院专用部件','适用于甘肃中医院','zl9BodyEditorGS',1,'zl9TemperatureChartGS');
    END IF;
  END;
END; 
/
COMMIT;


