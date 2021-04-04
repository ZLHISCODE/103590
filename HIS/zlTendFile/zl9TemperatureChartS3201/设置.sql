DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 新部件='zl9TemperatureChartS3201',启用=1 WHERE 部件='zl9BodyEditorSXHZ';
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('陕西3201医院专用部件','适用于陕西3201医院','zl9BodyEditorSXHZ',1,'zl9TemperatureChartS3201');
    END IF;
  END;
END; 
/
COMMIT;


