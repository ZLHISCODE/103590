DECLARE 
BEGIN
  UPDATE ���²��� SET ����=0;
  BEGIN 
    UPDATE ���²��� SET �²���='zl9TemperatureChartGD',����=1 WHERE ����='zl9BodyEditorGD';
    IF SQL%notfound then
      INSERT INTO ���²��� (����,���õ���,����,����,�²���)
      VALUES ('�㶫���²���','�����ڹ㶫ʡ','zl9BodyEditorGD',1,'zl9TemperatureChartGD');
    END IF;
  END;
END; 
/
COMMIT;



