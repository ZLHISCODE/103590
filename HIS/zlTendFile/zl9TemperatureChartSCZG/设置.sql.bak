DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 启用=1 WHERE Upper(新部件)=Upper('zl9TemperatureChartSC');
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('四川地区通用体温部件','适用于四川地区','zl9BodyEditor',1,'zl9TemperatureChartSC');
    END IF;
  END;
END; 
/
COMMIT;


