DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 新部件='zl9TemperatureChartYN',启用=1 WHERE 部件='zl9BodyEditorYN';
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('云南大理专用体温部件','适用于云南大理','zl9BodyEditor',1,'zl9TemperatureChartYN');
    END IF;
  END;
END; 
/
COMMIT;
