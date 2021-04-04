DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 启用=1,新部件='zl9TemperatureChartYDEY' WHERE Upper(部件)=Upper('zl9BodyEditorYDEY');
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('体温部件','体温部件','zl9BodyEditorYDEY',1,'zl9TemperatureChartYDEY',');
    END IF;
  END;
END; 
/
COMMIT;



