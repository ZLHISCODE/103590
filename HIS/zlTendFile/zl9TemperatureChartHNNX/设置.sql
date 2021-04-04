DECLARE 
BEGIN
  UPDATE 体温部件 SET 启用=0;
  BEGIN 
    UPDATE 体温部件 SET 新部件='zl9TemperatureChartHNNX',启用=1 WHERE 部件='zl9BodyEditor';
    IF SQL%notfound then
      INSERT INTO 体温部件 (名称,适用地区,部件,启用,新部件)
      VALUES ('湖南宁乡人民医院专用体温部件','湖南宁乡人民医院','zl9BodyEditorHNNX',1,'zl9TemperatureChartHNNX');
    END IF;
  END;
END; 
/
COMMIT;



