DECLARE 
BEGIN
  UPDATE ���²��� SET ����=0;
  BEGIN 
    UPDATE ���²��� SET �²���='zl9TemperatureChartJX',����=1 WHERE ����='zl9BodyEditorJX';
    IF SQL%notfound then
      INSERT INTO ���²��� (����,���õ���,����,����,�²���)
      VALUES ('�������²���','�����ڽ���ʡ','zl9BodyEditorJX',1,'zl9TemperatureChartJX');
    END IF;
  END;
END; 
/
Update Zlparameters
Set ����ֵ=20
Where ������ = '�������߹̶��������' And Nvl(ģ��, 0) = 1255 And Nvl(ϵͳ, 0) = 100 and ������ = 76 ;
/
Update ���¼�¼��Ŀ
Set ��¼�� = 1, ��¼�� = '��', ��¼ɫ = '16744448', �̶ȼ�� = '10.00000', ��λֵ = '1.00000', ��¼Ƶ�� = 2, ��λ = '��/��', ����� = '0'
Where ��Ŀ��� = 3;
/ 
COMMIT;

