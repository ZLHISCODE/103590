
Define n_System=100;

CREATE OR REPLACE PROCEDURE ZL_体温单开始日期_UPDATE(
	病人ID_IN IN 病案主页.病人ID%TYPE, 
	主页ID_IN IN 病案主页.主页ID%TYPE, 
	开始日期_IN IN 病案主页从表.信息值%TYPE) 
AS  
BEGIN  
	UPDATE 病案主页从表 
	SET 信息值=开始日期_IN 
	WHERE 病人ID=病人ID_IN AND 主页ID=主页ID_IN AND 信息名='体温单开始日期'; 
	IF SQL%ROWCOUNT =0 THEN  
		INSERT INTO 病案主页从表(病人ID,主页ID,信息名,信息值) 
		VALUES (病人ID_IN ,主页ID_IN ,'体温单开始日期',开始日期_IN); 
	END IF ; 
END ;
/

--1255
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(&n_System,1255,'体温单作图',User,'ZL_体温单开始日期_UPDATE','EXECUTE');