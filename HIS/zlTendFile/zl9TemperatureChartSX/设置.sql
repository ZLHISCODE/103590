DECLARE 
BEGIN
  UPDATE ���²��� SET ����=0;
  BEGIN 
    UPDATE ���²��� SET �²���='zl9TemperatureChartSX',����=1 WHERE ����='zl9BodyEditor';
    IF SQL%notfound then
      INSERT INTO ���²��� (����,���õ���,����,����,�²���)
      VALUES ('ɽ������ר�ò���','������ɽ������','zl9BodyEditor',1,'zl9TemperatureChartSX');
    END IF;
  END;
END; 
/
COMMIT;


