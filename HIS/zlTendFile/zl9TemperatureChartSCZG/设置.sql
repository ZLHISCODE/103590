DECLARE 
BEGIN
  UPDATE ���²��� SET ����=0;
  BEGIN 
    UPDATE ���²��� SET ����=1 WHERE Upper(�²���)=Upper('zl9TemperatureChartSCZG');
    IF SQL%notfound then
      INSERT INTO ���²��� (����,���õ���,����,����,�²���)
      VALUES ('�Ĵ��Թ��е���ͨ�����²���','�������Ĵ��Թ���','zl9BodyEditor',1,'zl9TemperatureChartSCZG');
    END IF;
  END;
END; 
/

--�޸Ķ�Ӧ����ֵ
UPDATE zlparameters SET ����ֵ='1' WHERE ϵͳ=100 AND ģ��=1255 AND ������='���������(����/����)��ʽ¼��';
UPDATE zlparameters SET ����ֵ='0' WHERE ϵͳ=100 AND ģ��=1255 AND ������='��������ȱʡ��ʽ';
COMMIT;


