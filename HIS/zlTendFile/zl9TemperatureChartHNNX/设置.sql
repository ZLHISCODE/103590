DECLARE 
BEGIN
  UPDATE ���²��� SET ����=0;
  BEGIN 
    UPDATE ���²��� SET �²���='zl9TemperatureChartHNNX',����=1 WHERE ����='zl9BodyEditor';
    IF SQL%notfound then
      INSERT INTO ���²��� (����,���õ���,����,����,�²���)
      VALUES ('������������ҽԺר�����²���','������������ҽԺ','zl9BodyEditorHNNX',1,'zl9TemperatureChartHNNX');
    END IF;
  END;
END; 
/
COMMIT;



