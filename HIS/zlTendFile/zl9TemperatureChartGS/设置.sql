DECLARE 
BEGIN
  UPDATE ���²��� SET ����=0;
  BEGIN 
    UPDATE ���²��� SET �²���='zl9TemperatureChartGS',����=1 WHERE ����='zl9BodyEditorGS';
    IF SQL%notfound then
      INSERT INTO ���²��� (����,���õ���,����,����,�²���)
      VALUES ('������ҽԺר�ò���','�����ڸ�����ҽԺ','zl9BodyEditorGS',1,'zl9TemperatureChartGS');
    END IF;
  END;
END; 
/
COMMIT;


