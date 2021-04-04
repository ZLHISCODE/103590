BEGIN
	UPDATE 体温部件 SET 启用=0;
	UPDATE 体温部件 SET 启用=1 WHERE upper(部件)='ZL9BODYEDITORHN';
	IF SQL%NOTFOUND THEN 
		INSERT INTO 体温部件 (名称,适用地区,部件,启用)
		VALUES ('河南地区专用体温部件','适用河南地区','ZL9BODYEDITORHN',1);
	END IF;
END;
/
COMMIT;