DECLARE 
BEGIN
  UPDATE ���²��� SET ����=0;
  BEGIN 
    UPDATE ���²��� SET �²���='zl9TemperatureChartS3201',����=1 WHERE ����='zl9BodyEditorSXHZ';
    IF SQL%notfound then
      INSERT INTO ���²��� (����,���õ���,����,����,�²���)
      VALUES ('����3201ҽԺר�ò���','����������3201ҽԺ','zl9BodyEditorSXHZ',1,'zl9TemperatureChartS3201');
    END IF;
  END;
END; 
/
COMMIT;


