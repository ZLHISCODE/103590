DECLARE 
BEGIN
  UPDATE ���²��� SET ����=0;
  BEGIN 
    UPDATE ���²��� SET ����=1 WHERE �²���='zl9TemperatureChartNJ';
    IF SQL%notfound then
      INSERT INTO ���²��� (����,���õ���,����,����,�²���)
      VALUES ('���յ���ר�ò���','�����ڽ��յ���','zl9BodyEditor',1,'zl9TemperatureChartNJ');
    END IF;
  END;
END; 
/
COMMIT;


