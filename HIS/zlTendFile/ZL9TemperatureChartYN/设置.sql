DECLARE 
BEGIN
  UPDATE ���²��� SET ����=0;
  BEGIN 
    UPDATE ���²��� SET �²���='zl9TemperatureChartYN',����=1 WHERE ����='zl9BodyEditorYN';
    IF SQL%notfound then
      INSERT INTO ���²��� (����,���õ���,����,����,�²���)
      VALUES ('���ϴ���ר�����²���','���������ϴ���','zl9BodyEditor',1,'zl9TemperatureChartYN');
    END IF;
  END;
END; 
/
COMMIT;
