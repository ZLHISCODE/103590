BEGIN
	UPDATE ���²��� SET ����=0;
	UPDATE ���²��� SET ����=1 WHERE upper(����)='ZL9BODYEDITORHN';
	IF SQL%NOTFOUND THEN 
		INSERT INTO ���²��� (����,���õ���,����,����)
		VALUES ('���ϵ���ר�����²���','���ú��ϵ���','ZL9BODYEDITORHN',1);
	END IF;
END;
/
COMMIT;