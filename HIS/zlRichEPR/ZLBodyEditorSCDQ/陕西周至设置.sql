Update ���¼�¼��Ŀ Set ���ֵ=42,��Сֵ=35,��λֵ=0.1,�����=2 Where ��Ŀ���=1;
Update ���¼�¼��Ŀ Set ���ֵ=200,��Сֵ=40,��λֵ=2,�����=2 Where ��Ŀ���=2;
Update ���¼�¼��Ŀ Set ���ֵ=200,��Сֵ=40,��λֵ=2,�����=2 Where ��Ŀ���=-1;
Update �����¼��Ŀ Set ��Ŀֵ��='35;42;' Where ��Ŀ���=1;
Update �����¼��Ŀ Set ��Ŀֵ��='40;200;' Where ��Ŀ���=2;
Update �����¼��Ŀ Set ��Ŀֵ��='40;200;' Where ��Ŀ���=-1;

UPDATE ���²��� SET ����=0;

DECLARE 
  nCount number;
BEGIN
  BEGIN 
    SELECT 1 INTO nCount From ���²��� WHERE upper(����)='ZL9BODYEDITORSCDQ';
    EXCEPTION 
      WHEN OTHERS THEN nCount:=0;
  END;
  IF nCount=0 THEN 
    INSERT INTO ���²��� (����,���õ���,����,����)
      VALUES ('�Ĵ�ʡͨ�����²���','�Ĵ�ʡ���������û�','ZL9BODYEDITORSCDQ',1);
  ELSE 
    UPDATE ���²��� SET ����=1 WHERE upper(����)='ZL9BODYEDITORSCDQ';
  END if;
end;
/
commit;
