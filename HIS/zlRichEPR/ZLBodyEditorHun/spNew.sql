
Define n_System=100;

CREATE OR REPLACE PROCEDURE ZL_���µ���ʼ����_UPDATE(
	����ID_IN IN ������ҳ.����ID%TYPE, 
	��ҳID_IN IN ������ҳ.��ҳID%TYPE, 
	��ʼ����_IN IN ������ҳ�ӱ�.��Ϣֵ%TYPE) 
AS  
BEGIN  
	UPDATE ������ҳ�ӱ� 
	SET ��Ϣֵ=��ʼ����_IN 
	WHERE ����ID=����ID_IN AND ��ҳID=��ҳID_IN AND ��Ϣ��='���µ���ʼ����'; 
	IF SQL%ROWCOUNT =0 THEN  
		INSERT INTO ������ҳ�ӱ�(����ID,��ҳID,��Ϣ��,��Ϣֵ) 
		VALUES (����ID_IN ,��ҳID_IN ,'���µ���ʼ����',��ʼ����_IN); 
	END IF ; 
END ;
/

--1255
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(&n_System,1255,'���µ���ͼ',User,'ZL_���µ���ʼ����_UPDATE','EXECUTE');