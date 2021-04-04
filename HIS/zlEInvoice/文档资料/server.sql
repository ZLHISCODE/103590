CREATE SEQUENCE ����Ʊ���쳣��¼_ID START WITH 1;
Create Table ����Ʊ���쳣��¼(
	ID Number(18),
	�������� number(2),
	ҵ������ number(2),
	��¼��־ number(2),
	���ݺ� varchar2(20),
	ҵ��ID number(18),
	����Ʊ��id number(18),
	����ID number(18),
	���� varchar2(100),
	�Ա� varchar2(4),
	���� varchar2(20),
	����� number(18),
	סԺ�� number(18),
	�Ƿ񻻿� number(2),
	Ʊ����Ϣ CLOB,
	����Ա��� varchar2(6),
	����Ա���� varchar2(50),
	�Ǽ�ʱ�� Date)
 TABLESPACE zl9Expense;

Alter Table ����Ʊ���쳣��¼ Add Constraint ����Ʊ���쳣��¼_PK Primary Key(ID) Using Index Tablespace zl9Indexhis;
Alter table ����Ʊ���쳣��¼ Add Constraint ����Ʊ���쳣��¼_UQ_ҵ��ID Unique(ҵ��ID,��¼��־,��������,ҵ������)  Using Index Tablespace zl9Indexhis; 
Alter Table ����Ʊ���쳣��¼ Add Constraint ����Ʊ���쳣��¼_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID);
CREATE INDEX ����Ʊ���쳣��¼_IX_�Ǽ�ʱ�� ON ����Ʊ���쳣��¼(�Ǽ�ʱ��) TABLESPACE zl9Indexhis; 
CREATE INDEX ����Ʊ���쳣��¼_IX_����ID ON ����Ʊ���쳣��¼(����ID) TABLESPACE zl9Indexhis;
CREATE INDEX ����Ʊ���쳣��¼_IX_����Ʊ��id ON ����Ʊ���쳣��¼(����Ʊ��id) TABLESPACE zl9Indexhis;
CREATE INDEX ����Ʊ���쳣��¼_IX_���ݺ� ON ����Ʊ���쳣��¼(���ݺ�,��������) TABLESPACE zl9Indexhis; 



Create Or Replace Procedure Zl_����Ʊ��ʹ�ü�¼_Insert
(
  Id_In         In ����Ʊ��ʹ�ü�¼.Id%Type,
  Ʊ��_In       In ����Ʊ��ʹ�ü�¼.Ʊ��%Type,
  ����id_In     In ����Ʊ��ʹ�ü�¼.����id%Type,
  ����id_In     In ����Ʊ��ʹ�ü�¼.����id%Type,
  ����_In       In ����Ʊ��ʹ�ü�¼.����%Type,
  �Ա�_In       In ����Ʊ��ʹ�ü�¼.�Ա�%Type,
  ����_In       In ����Ʊ��ʹ�ü�¼.����%Type,
  �����_In     In ����Ʊ��ʹ�ü�¼.�����%Type,
  סԺ��_In     In ����Ʊ��ʹ�ü�¼.סԺ��%Type,
  Ʊ�ݽ��_In   In ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type,
  ��Ʊ��_In     In ����Ʊ��ʹ�ü�¼.��Ʊ��%Type,
  ϵͳ��Դ_In   In ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type,
  ����ʱ��_In   In ����Ʊ��ʹ�ü�¼.����ʱ��%Type,
  ��ע_In       In ����Ʊ��ʹ�ü�¼.��ע%Type,
  ����Ա���_In In ����Ʊ��ʹ�ü�¼.����Ա���%Type,
  ����Ա����_In In ����Ʊ��ʹ�ü�¼.����Ա����%Type,
  �Ǽ�ʱ��_In   In ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type,
  ԭƱ��id_In   In ����Ʊ��ʹ�ü�¼.ԭƱ��id%Type := Null,
  �˿�id_In     In ����Ʊ��ʹ�ü�¼.ԭƱ��id%Type := Null,
  ����_In       In ����Ʊ��ʹ�ü�¼.����%Type := Null,
  ����_In       In ����Ʊ��ʹ�ü�¼.����%Type := Null,
  ������_In     In ����Ʊ��ʹ�ü�¼.������%Type := Null,
  ƾ֤����_In   In ����Ʊ��ʹ�ü�¼.ƾ֤����%Type := Null,
  ƾ֤����_In   In ����Ʊ��ʹ�ü�¼.ƾ֤����%Type := Null,
  ƾ֤������_In In ����Ʊ��ʹ�ü�¼.ƾ֤������%Type := Null,
  Url����_In    In ����Ʊ��ʹ�ü�¼.Url����%Type := Null,
  Url����_In    In ����Ʊ��ʹ�ü�¼.Url����%Type := Null
) As
  n_��¼״̬ ����Ʊ��ʹ�ü�¼.��¼״̬%Type;
Begin
  n_��¼״̬ := 1;

  Insert Into ����Ʊ��ʹ�ü�¼
    (ID, Ʊ��, ��¼״̬, ����id, ����id, ����, �Ա�, ����, �����, סԺ��, Ʊ�ݽ��, ����ʱ��, ԭƱ��id, �˿�id, ��Ʊ��, ϵͳ��Դ, ����, ����, ������, ƾ֤����, ƾ֤����, ƾ֤������,
     Url����, Url����, ��ע, ����Ա���, ����Ա����, �Ǽ�ʱ��)
  Values
    (Id_In, Ʊ��_In, n_��¼״̬, ����id_In, Decode(Nvl(����id_In, 0), 0, Null, ����id_In), ����_In, �Ա�_In, ����_In,
     Decode(Nvl(�����_In, 0), 0, Null, �����_In), Decode(Nvl(סԺ��_In, 0), 0, Null, סԺ��_In), Ʊ�ݽ��_In, ����ʱ��_In, ԭƱ��id_In,
     �˿�id_In, ��Ʊ��_In, ϵͳ��Դ_In, ����_In, ����_In, ������_In, ƾ֤����_In, ƾ֤����_In, ƾ֤������_In, Url����_In, Url����_In, ��ע_In, ����Ա���_In,
     ����Ա����_In, Nvl(�Ǽ�ʱ��_In, Sysdate));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ��ʹ�ü�¼_Insert;
/


Create Or Replace Procedure Zl_����Ʊ��ʹ�ü�¼_Delete
(
  Id_In           In ����Ʊ��ʹ�ü�¼.Id%Type,
  ��Ʊ��_In       In ����Ʊ��ʹ�ü�¼.��Ʊ��%Type,
  ϵͳ��Դ_In     In ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type,
  ����ʱ��_In     In ����Ʊ��ʹ�ü�¼.����ʱ��%Type,
  ��ע_In         In ����Ʊ��ʹ�ü�¼.��ע%Type,
  ����Ա���_In   In ����Ʊ��ʹ�ü�¼.����Ա���%Type,
  ����Ա����_In   In ����Ʊ��ʹ�ü�¼.����Ա����%Type,
  �Ǽ�ʱ��_In     In ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type,
  ԭ����Ʊ��id_In In ����Ʊ��ʹ�ü�¼.Id%Type,
  ����_In         In ����Ʊ��ʹ�ü�¼.����%Type := Null,
  ����_In         In ����Ʊ��ʹ�ü�¼.����%Type := Null,
  ������_In       In ����Ʊ��ʹ�ü�¼.������%Type := Null,
  ƾ֤����_In     In ����Ʊ��ʹ�ü�¼.ƾ֤����%Type := Null,
  ƾ֤����_In     In ����Ʊ��ʹ�ü�¼.ƾ֤����%Type := Null,
  ƾ֤������_In   In ����Ʊ��ʹ�ü�¼.ƾ֤������%Type := Null,
  Url����_In      In ����Ʊ��ʹ�ü�¼.Url����%Type := Null,
  Url����_In      In ����Ʊ��ʹ�ü�¼.Url����%Type := Null
) As
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_�Ƿ񻻿� ����Ʊ��ʹ�ü�¼.�Ƿ񻻿�%Type;
Begin

  Update ����Ʊ��ʹ�ü�¼ Set ��¼״̬ = 3 Where ID = ԭ����Ʊ��id_In Returning Nvl(�Ƿ񻻿�, 0) Into n_�Ƿ񻻿�;
  If Sql%NotFound Then
    v_Err_Msg := 'δ�ҵ�ԭʼ�ĵ���Ʊ����Ϣ���������ϲ���!';
    Raise Err_Item;
  End If;
  If Nvl(n_�Ƿ񻻿�, 0) = 1 Then
    --��ǰ����Ʊ���Ѿ�����ֽ��Ʊ��
    v_Err_Msg := '��ǰ����Ʊ���Ѿ�����ֽ��Ʊ��,��Ҫ�ȳ��ֽ��Ʊ�ݺ�������ϵ��ӷ�Ʊ!';
    Raise Err_Item;
  End If;

  Insert Into ����Ʊ��ʹ�ü�¼
    (ID, Ʊ��, ��¼״̬, ����id, ����id, ����, �Ա�, ����, �����, סԺ��, ����, ����, ������, Ʊ�ݽ��, ����ʱ��, ԭƱ��id, ��ӡid, �Ƿ񻻿�, ֽ�ʷ�Ʊ��, ��Ʊ��, ϵͳ��Դ, ��ע,
     ����Ա���, ����Ա����, �Ǽ�ʱ��, �˿�id, ƾ֤����, ƾ֤����, ƾ֤������, Url����, Url����)
    Select Id_In, Ʊ��, 2, ����id, ����id, ����, �Ա�, ����, �����, סԺ��, Nvl(����_In, ����), Nvl(����_In, ����), Nvl(������_In, ������), Ʊ�ݽ��,
           ����ʱ��_In, ԭ����Ʊ��id_In, ��ӡid, �Ƿ񻻿�, ֽ�ʷ�Ʊ��, Nvl(��Ʊ��_In, ��Ʊ��) As ��Ʊ��, Nvl(ϵͳ��Դ_In, ϵͳ��Դ) As ϵͳ��Դ,
           Nvl(��ע_In, ��ע) As ��ע, ����Ա���_In, ����Ա����_In, �Ǽ�ʱ��_In, �˿�id, Nvl(ƾ֤����_In, ƾ֤����), Nvl(ƾ֤����_In, ƾ֤����),
           Nvl(ƾ֤������_In, ƾ֤������), Nvl(Url����_In, Url����), Nvl(Url����_In, Url����)
    From ����Ʊ��ʹ�ü�¼
    Where ID = ԭ����Ʊ��id_In;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ��ʹ�ü�¼_Delete;
/
 
Create Or Replace Procedure Zl_����Ʊ���쳣��¼_Insert
(
  �쳣id_In     ����Ʊ���쳣��¼.Id%Type,
  ��������_In   ����Ʊ���쳣��¼.��������%Type,
  ҵ������_In   ����Ʊ���쳣��¼.ҵ������%Type,
  ��¼��־_IN   ����Ʊ���쳣��¼.��¼��־%Type,
  ���ݺ�_In     ����Ʊ���쳣��¼.���ݺ�%Type,
  ҵ��id_In     ����Ʊ���쳣��¼.ҵ��id%Type,
  ����Ʊ��id_In ����Ʊ���쳣��¼.����Ʊ��id%Type,
  ����id_In     ����Ʊ���쳣��¼.����id%Type,
  ����_In       ����Ʊ���쳣��¼.����%Type,
  �Ա�_In       ����Ʊ���쳣��¼.�Ա�%Type,
  ����_In       ����Ʊ���쳣��¼.����%Type,
  �����_In     ����Ʊ���쳣��¼.�����%Type,
  סԺ��_In     ����Ʊ���쳣��¼.סԺ��%Type,
  ����Ա���_In ����Ʊ���쳣��¼.����Ա���%Type,
  ����Ա����_In ����Ʊ���쳣��¼.����Ա����%Type,
  �Ǽ�ʱ��_In   ����Ʊ���쳣��¼.�Ǽ�ʱ��%Type,
  �Ƿ񻻿�_In   ����Ʊ���쳣��¼.�Ƿ񻻿�%Type := 0
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�������Ʊ���쳣��¼
  -- ��� 
  --   ��������_In:1-ҽ�ƿ�����;2-������Ϣ�Ǽ�;3-������Ժ�Ǽ�
  --   ҵ������_In:1-ҽ�ƿ�,2-Ԥ��
  --   ��¼��־_IN:0-���ߵ���Ʊ��;1-������Ʊ��;2-ֽ��Ʊ��;3-����ֽ��Ʊ��
  --   ���ݺ�_In:ҵ������=1:��ʾҽ�ƿ�����NO,ҵ������=2:��ʾԤ����NO
  --   ҵ��ID:ԭ����ID��ԭԤ��ID
  ------------------------------------------------------------------------------------------------------------------------------
  n_�쳣id   Number(18);
  d_�Ǽ�ʱ�� ����Ʊ���쳣��¼.�Ǽ�ʱ��%Type;
Begin

  n_�쳣id   := �쳣id_In;
  d_�Ǽ�ʱ�� := �Ǽ�ʱ��_In;
  If d_�Ǽ�ʱ�� Is Null Then
    d_�Ǽ�ʱ�� := Sysdate;
  End If;

  If Nvl(n_�쳣id, 0) = 0 Then
    Select ����Ʊ���쳣��¼_Id.Nextval Into n_�쳣id From Dual;
  End If;

  Insert Into ����Ʊ���쳣��¼
    (ID, ��������, ҵ������, ��¼��־, ���ݺ�, ҵ��id, ����Ʊ��id, ����id, ����, �Ա�, ����, �����, סԺ��, �Ƿ񻻿�, ����Ա���, ����Ա����, �Ǽ�ʱ��)
  Values
    (n_�쳣id, ��������_In, ҵ������_In, ��¼��־_IN, ���ݺ�_In, ҵ��id_In, ����Ʊ��id_In, ����id_In, ����_In, �Ա�_In, ����_In, �����_In, סԺ��_In,
     �Ƿ񻻿�_In, ����Ա���_In, ����Ա����_In, d_�Ǽ�ʱ��);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ���쳣��¼_Insert;
/

Create Or Replace Procedure Zl_����Ʊ���쳣��¼_Modify
(
  �쳣id_In   ����Ʊ���쳣��¼.Id%Type, 
  Ʊ����Ϣ_In Clob, 
  �Ƿ񻻿�_In ����Ʊ���쳣��¼.�Ƿ񻻿�%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --����:���µ���Ʊ���쳣��Ϣ
  -- ��� 
  --     Ʊ����Ϣ_In:�������_in=0ʱ�����µ���Ʊ����Ϣ�ֶ�;�������_in=1ʱ����ֽ��Ʊ���ֶ� 
  --     �Ƿ񻻿�_In:NULL-��ʾ�������Ƿ񻻿��ֶ�;������³ɵ�ǰ����ֵ
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin 
    Update ����Ʊ���쳣��¼
    Set �Ƿ񻻿� = Nvl(�Ƿ񻻿�_In, �Ƿ񻻿�), Ʊ����Ϣ = Ʊ����Ϣ_In
    Where ID = Nvl(�쳣id_In, 0);
    If Sql%NotFound Then
      v_Err_Msg := 'δ�ҵ���Ҫ���µ�Ʊ����Ϣ������!';
      Raise Err_Item;
    End If; 
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ���쳣��¼_Modify;
/

Create Or Replace Procedure Zl_����Ʊ���쳣��¼_Delete(�쳣id_In ����Ʊ���쳣��¼.Id%Type) As
  ------------------------------------------------------------------------------------------------------------------------------
  --����:ɾ���쳣��¼
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin
  Delete ����Ʊ���쳣��¼ Where ID = Nvl(�쳣id_In, 0);
  If Sql%NotFound Then
    v_Err_Msg := 'δ�ҵ���Ҫɾ���ĵ���Ʊ����Ϣ������!';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ���쳣��¼_Delete;
/

Create Or Replace Procedure Zl_Exsesvr_Addeinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ӵ���Ʊ����Ϣ
  --��Σ�Json_In:��ʽ
  --  input      
  --    balance_id          N  1  ����ID
  --    balance_delid       N     �˿�ID:�˿�ߺ�Ʊʱ��Ч��Ŀǰֻ��Ԥ������Ч,��д�����˿�Ԥ��ID
  --    einvoice_id         N  1  ����Ʊ��ID
  --    operator_code       C  1  ����Ա���
  --    operator_name       C  1  ����Ա����
  --    create_time         C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    pati_info           C     ������Ϣ
  --      pati_id           N  1  ����ID
  --      pati_pageid       N     ��ҳID
  --      pati_name         C  1  ����
  --      pati_sex          C  1  �Ա�
  --      pati_age          C  1  ����
  --      outpatient_num    C  1  �����
  --      inpatient_num     C  1  סԺ��
  --    einvoce_info        C     ����Ʊ����Ϣ
  --      invoice_type      N  1  Ʊ�֣�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
  --      placeCode         C  1  ��Ʊ�����
  --      inv_total         N  1  ��Ʊ���
  --      inv_oldid         N     ԭƱ��ID
  --      sys_source        C  1  ϵͳ��Դ
  --      demo              C  1  ��ע
  --      einvoice_code     C  1  ����Ʊ�ݴ���
  --      einvoice_no       C  1  ����Ʊ�ݺ���
  --      einvoice_random   C  1  ����У����
  --      voucher_code      C  1  Ԥ����ƾ֤����
  --      voucher_no        C  1  Ԥ����ƾ֤����
  --      voucher_random    C  1  Ԥ����ƾ֤У����
  --      create_time       C  1  ����Ʊ������ʱ��:yyyymmddhh24miss
  --      picture_url       C  1  ����Ʊ��H5ҳ��URL
  --      picture_neturl    C  1  ����Ʊ������H5ҳ��URL
  --      qrcode            C  1  ����Ʊ�ݶ�ά��ͼƬ����:��ֵ��Base64���룬����ʱ��ҪBase64����,ͼƬ��ʽΪ:PNG
  --    --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  n_Id         ����Ʊ��ʹ�ü�¼.Id%Type;
  n_Ʊ��       ����Ʊ��ʹ�ü�¼.Ʊ��%Type;
  n_��¼״̬   ����Ʊ��ʹ�ü�¼.��¼״̬%Type;
  n_����id     ����Ʊ��ʹ�ü�¼.����id%Type;
  n_����id     ����Ʊ��ʹ�ü�¼.����id%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_�Ա�       ����Ʊ��ʹ�ü�¼.�Ա�%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  n_�����     ����Ʊ��ʹ�ü�¼.�����%Type;
  n_סԺ��     ����Ʊ��ʹ�ü�¼.סԺ��%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_������     ����Ʊ��ʹ�ü�¼.������%Type;
  v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
  v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
  v_ƾ֤������ ����Ʊ��ʹ�ü�¼.ƾ֤������%Type;
  n_Ʊ�ݽ��   ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type;
  v_����ʱ��   ����Ʊ��ʹ�ü�¼.����ʱ��%Type;
  v_Url����    ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����    ����Ʊ��ʹ�ü�¼.Url����%Type;
  c_��ά��     Clob;
  n_ԭƱ��id   ����Ʊ��ʹ�ü�¼.ԭƱ��id%Type;
  n_�˿�id     ����Ʊ��ʹ�ü�¼.�˿�id%Type;
  v_��ע       ����Ʊ��ʹ�ü�¼.��ע%Type;
  v_��Ʊ��     ����Ʊ��ʹ�ü�¼.��Ʊ��%Type;
  v_ϵͳ��Դ   ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type;
  v_����Ա��� ����Ʊ��ʹ�ü�¼.����Ա���%Type;
  v_����Ա���� ����Ʊ��ʹ�ü�¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type;

  n_��¼״̬ Number(2);
  j_Input    PLJson;
  j_Json     PLJson;
  j_Temp     PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id     := j_Json.Get_Number('balance_id');
  n_�˿�id     := j_Json.Get_Number('balance_delid');
  n_Id         := j_Json.Get_Number('einvoice_id');
  v_����Ա��� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');
  d_�Ǽ�ʱ��   := Nvl(To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  --��ȡ������Ϣ

  If Not j_Json.Exist('pati_info') Then
  
    Json_Out := zlJsonOut('�޲�����Ϣ���������ӵ���Ʊ�ݡ�');
    Return;
  End If;

  j_Temp   := j_Json.Get_Pljson('pati_info');
  n_����id := j_Temp.Get_Number('pati_id');
  --n_��ҳid := j_Temp.Get_Number('pati_pageid');

  v_����   := j_Temp.Get_String('pati_name');
  v_�Ա�   := j_Temp.Get_String('pati_sex');
  v_����   := j_Temp.Get_String('pati_age');
  n_����� := j_Temp.Get_Number('outpatient_num');
  n_סԺ�� := j_Temp.Get_Number('inpatient_num');

  If Not j_Json.Exist('einvoce_info') Then
  
    Json_Out := zlJsonOut('�޵���Ʊ����Ϣ,�������ӵ���Ʊ�ݡ�');
    Return;
  End If;

  --��ȡ����Ʊ����Ϣ
  j_Temp := PLJson();
  j_Temp := j_Json.Get_Pljson('einvoce_info');
  --Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  n_Ʊ��       := Nvl(j_Temp.Get_Number('invoice_type'), 1);
  v_��Ʊ��     := j_Temp.Get_String('placeCode');
  n_ԭƱ��id   := j_Temp.Get_Number('inv_oldid');
  v_ϵͳ��Դ   := j_Temp.Get_String('sys_source');
  v_��ע       := j_Temp.Get_String('demo');
  v_����       := j_Temp.Get_String('einvoice_code');
  v_����       := j_Temp.Get_String('einvoice_no');
  v_������     := j_Temp.Get_String('einvoice_random');
  v_ƾ֤����   := j_Temp.Get_String('voucher_code');
  v_ƾ֤����   := j_Temp.Get_String('voucher_no');
  v_ƾ֤������ := j_Temp.Get_String('voucher_random');
  n_Ʊ�ݽ��   := j_Temp.Get_Number('inv_total');
  v_����ʱ��   := j_Temp.Get_String('happen_time');
  v_Url����    := j_Temp.Get_String('picture_url');
  v_Url����    := j_Temp.Get_String('picture_neturl');
  c_��ά��     := j_Temp.Get_Clob('qrcode');

  --���ӵ���Ʊ����Ϣ 
  Zl_����Ʊ��ʹ�ü�¼_Insert(n_Id, n_Ʊ��, n_����id, n_����id, v_����, v_�Ա�, v_����, n_�����, n_סԺ��, n_Ʊ�ݽ��, v_��Ʊ��, v_ϵͳ��Դ, v_����ʱ��, v_��ע,
                     v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��, n_ԭƱ��id, n_�˿�id, v_����, v_����, v_������, v_ƾ֤����, v_ƾ֤����, v_ƾ֤������, v_Url����,
                     v_Url����);
  --���¶�ά��
  Insert Into ����Ʊ�ݶ�ά�� (ʹ�ü�¼id, ��ά��) Values (n_Id, c_��ά��);
  Json_Out := zlJsonOut('�ɹ�', 1);
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Addeinvoice;
/



Create Or Replace Procedure Zl_Exsesvr_Deleinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�ɾ������Ʊ����Ϣ
  --��Σ�Json_In:��ʽ
  -- input      
  --  einvoice_id  N  1  ����Ʊ��ID
  --  operator_code  C  1  ����Ա���
  --  operator_name  C  1  ����Ա����
  --  create_time  C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --  einvoce_info  C    ����Ʊ����Ϣ
  --    placeCode  C  1  ��Ʊ�����
  --    sys_source  C  1  ϵͳ��Դ
  --    demo  C  1  ��ע
  --    inv_oldid  N    ԭƱ��ID
  --    einvoice_code  C  1  ����Ʊ�ݴ���
  --    einvoice_no  C  1  ����Ʊ�ݺ���
  --    einvoice_random  C  1  ����У����
  --    voucher_code  C  1  Ԥ����ƾ֤����
  --    voucher_no  C  1  Ԥ����ƾ֤����
  --    voucher_random  C  1  Ԥ����ƾ֤У����
  --    happen_time  C  1  ����Ʊ������ʱ��:yyyymmddhh24miss
  --    picture_url  C  1  ����Ʊ��H5ҳ��URL
  --    picture_neturl  C  1  ����Ʊ������H5ҳ��URL
  --    qrcode  C  1  ����Ʊ�ݶ�ά��ͼƬ����:��ֵ��Base64���룬����ʱ��ҪBase64����,ͼƬ��ʽΪ:PNG
  --    --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  n_Id         ����Ʊ��ʹ�ü�¼.Id%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_������     ����Ʊ��ʹ�ü�¼.������%Type;
  v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
  v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
  v_ƾ֤������ ����Ʊ��ʹ�ü�¼.ƾ֤������%Type;
  v_����ʱ��   ����Ʊ��ʹ�ü�¼.����ʱ��%Type;
  v_Url����    ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����    ����Ʊ��ʹ�ü�¼.Url����%Type;
  c_��ά��     Clob;
  n_ԭƱ��id   ����Ʊ��ʹ�ü�¼.ԭƱ��id%Type;
  v_��ע       ����Ʊ��ʹ�ü�¼.��ע%Type;
  v_��Ʊ��     ����Ʊ��ʹ�ü�¼.��Ʊ��%Type;
  v_ϵͳ��Դ   ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type;
  v_����Ա��� ����Ʊ��ʹ�ü�¼.����Ա���%Type;
  v_����Ա���� ����Ʊ��ʹ�ü�¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type;

  j_Input PLJson;
  j_Json  PLJson;
  j_Temp  PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Id         := j_Json.Get_Number('einvoice_id');
  v_����Ա��� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');
  d_�Ǽ�ʱ��   := Nvl(To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  If Not j_Json.Exist('einvoce_info') Then
  
    Json_Out := zlJsonOut('�޵���Ʊ����Ϣ,�������ӵ���Ʊ�ݡ�');
    Return;
  End If;

  --��ȡ����Ʊ����Ϣ
  j_Temp     := PLJson();
  j_Temp     := j_Json.Get_Pljson('einvoce_info');
  v_��Ʊ��   := j_Temp.Get_String('placeCode');
  v_ϵͳ��Դ := j_Temp.Get_String('sys_source');
  v_��ע     := j_Temp.Get_String('demo');
  n_ԭƱ��id := j_Temp.Get_Number('inv_oldid');

  v_����       := j_Temp.Get_String('einvoice_code');
  v_����       := j_Temp.Get_String('einvoice_no');
  v_������     := j_Temp.Get_String('einvoice_random');
  v_ƾ֤����   := j_Temp.Get_String('voucher_code');
  v_ƾ֤����   := j_Temp.Get_String('voucher_no');
  v_ƾ֤������ := j_Temp.Get_String('voucher_random');
  v_����ʱ��   := j_Temp.Get_String('happen_time');
  v_Url����    := j_Temp.Get_String('picture_url');
  v_Url����    := j_Temp.Get_String('picture_neturl');
  c_��ά��     := j_Temp.Get_Clob('qrcode');

  Zl_����Ʊ��ʹ�ü�¼_Delete(n_Id, v_��Ʊ��, v_ϵͳ��Դ, v_����ʱ��, v_��ע, v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��, n_ԭƱ��id, v_����, v_����, v_������, v_ƾ֤����,
                     v_ƾ֤����, v_ƾ֤������, v_Url����, v_Url����);
  --���¶�ά��
  Insert Into ����Ʊ�ݶ�ά�� (ʹ�ü�¼id, ��ά��) Values (n_Id, c_��ά��);
  Json_Out := zlJsonOut('�ɹ�', 1);
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Deleinvoice;
/


Create Or Replace Procedure Zl_Exsesvr_Savepaperinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ֽ��Ʊ��ʹ����Ϣ
  --��Σ�Json_In:��ʽ
  --   input      
  --    oper_mode           N  1  ������ʽ:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
  --    einvoice_id         N  1  ����Ʊ��ID
  --    operator_code       C  1  ����Ա���
  --    operator_name       C  1  ����Ա����
  --    create_time         C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    paperinv_info       C     ֽ��Ʊ����Ϣ:���ڶ���ʱ���밴����˳���ϴ�(�������ݴ���)
  --      inv_occasion      N  1  Ӧ�ó���:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
  --      invoice_type      N  1  Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --      inv_red           N     �Ƿ��Ʊ:1-��Ʊ;0-�Ǻ�Ʊ
  --      invoice_no        C  1  ��Ʊ��
  --      inv_total         N  1  ��Ʊ���
  --      recv_id           N     ����id
  --    --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  n_Id ����Ʊ��ʹ�ü�¼.Id%Type;
  -- v_����Ա��� ����Ʊ��ʹ�ü�¼.����Ա���%Type;
  v_����Ա���� ����Ʊ��ʹ�ü�¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type;
  n_����id     ����Ʊ��ʹ�ü�¼.����id%Type;
  v_��Ʊ��     Ʊ��ʹ����ϸ.����%Type;
  n_��Ʊ���   Ʊ��ʹ����ϸ.Ʊ�ݽ��%Type;
  n_����id     Ʊ��ʹ����ϸ.����id%Type;
  n_������ʽ   Number(2);
  n_Ӧ�ó���   Number(2);
  n_Ʊ��       Number(2);
  n_�Ƿ��Ʊ   Number(2);
  j_Input      PLJson;
  j_Json       PLJson;
  j_Temp       PLJson;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  --������ʽ:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
  n_������ʽ := j_Json.Get_Number('oper_mode');
  n_Id       := j_Json.Get_Number('einvoice_id');
  --v_����Ա��� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');
  d_�Ǽ�ʱ��   := Nvl(To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  If Not j_Json.Exist('paperinv_info') Then
  
    Json_Out := zlJsonOut('��ֽ��Ʊ����Ϣ��');
    Return;
  End If;
  Select Max(����id) Into n_����id From ����Ʊ��ʹ�ü�¼ Where ID = n_Id;
  If Nvl(n_����id, 0) = 0 Then
  
    Json_Out := zlJsonOut('����ĵ���Ʊ����Ч!');
    Return;
  End If;

  j_Temp := j_Json.Get_Pljson('paperinv_info');
  --Ӧ�ó���:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
  n_Ӧ�ó��� := j_Temp.Get_Number('inv_occasion');
  --Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  n_Ʊ�� := j_Temp.Get_Number('invoice_type');
  --�Ƿ��Ʊ:1-��Ʊ;0-�Ǻ�Ʊ
  n_�Ƿ��Ʊ := Nvl(j_Temp.Get_Number('inv_red'), 0);
  v_��Ʊ��   := j_Temp.Get_String('invoice_no');
  n_��Ʊ��� := j_Temp.Get_Number('inv_total');
  n_����id   := j_Temp.Get_Number('recv_id');

  --ֽ��Ʊ�ݴ���
  --������ʽ_In:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
  Zl_ֽ��Ʊ��ʹ��_Update(n_Ӧ�ó���, n_Ʊ��, n_����id, n_Id, v_��Ʊ��, n_��Ʊ���, n_����id, v_����Ա����, d_�Ǽ�ʱ��, n_������ʽ, 0, n_�Ƿ��Ʊ);

  Json_Out := zlJsonOut('�ɹ�', 1);
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Savepaperinvoice;
/

Create Or Replace Procedure Zl_Exsesvr_Getstarteinvoices
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���õ���Ʊ��ҵ��
  --��Σ�Json_In:NULL
  --     
  --����: Json_Out,��ʽ����
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  data[]      ����վ���б�
  --    occasion  N  1  ����:1-�շ�,2-Ԥ��,3-����,4-�Һ�
  --    client_name  C  1  վ����
  ---------------------------------------------------------------------------

  v_Output Varchar2(32767);
  c_Output Clob;
Begin
  --�������
  For c_����� In (Select ����, վ�� From ����Ʊ��վ�����) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'occasion', c_�����.����, 1, 1);
    zlJsonPutValue(v_Output, 'client_name', c_�����.վ��, 0, 2);
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","data":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getstarteinvoices;
/
CREATE OR REPLACE Procedure Zl_Exsesvr_Geteinvoicecode
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ���ȡ��Ʊ����
  --��Σ�json��ʽ
  --input
  --   operator_id    N  1  ����ԱID
  --   ssite          C  1  �ͻ���
  --���Σ�json��ʽ
  --Json_Out
  --  code            C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message         C  1  Ӧ����Ϣ�� �ɹ�ʱ���ش���No��[����] ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  einvoice_code   C  1  ��Ʊ�����
  -------------------------------------------------------------------------------------------------
  n_����Աid   Ʊ�ݿ�Ʊ�����.��Աid%Type;
  v_�ͻ���     Ʊ�ݿ�Ʊ�����.�ͻ���%Type;
  v_��Ʊ����� ����Ʊ�ݿ�Ʊ��.����%Type;
  j_Input      PLJson;
  j_Json       PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����Աid := j_Json.Get_Number('operator_id');
  v_�ͻ���   := j_Json.Get_String('ssite');

  --���շ�Ա+�ͻ��˶���
  For r_��Ʊ�� In (Select b.���� As ��Ʊ�����
                From Ʊ�ݿ�Ʊ����� A, ����Ʊ�ݿ�Ʊ�� B
                Where a.��Ʊ��id = b.Id And Nvl(b.����ʱ��, Sysdate + 1) >= Sysdate And a.��Աid = n_����Աid And a.�ͻ��� = v_�ͻ���) Loop
    v_��Ʊ����� := r_��Ʊ��.��Ʊ�����;
    Json_Out     := '{"output":{"code":1,"message": "�ɹ�","einvoice_code":"' || Nvl(v_��Ʊ�����, '') || '"}}';
    Return;
  End Loop;

  --���շ�Ա����
  For r_��Ʊ�� In (Select Nvl(a.��Աid, 0) As ��Աid, Nvl(a.�ͻ���, '-') As �ͻ���, b.���� As ��Ʊ�����
                From Ʊ�ݿ�Ʊ����� A, ����Ʊ�ݿ�Ʊ�� B
                Where a.��Ʊ��id = b.Id And Nvl(b.����ʱ��, Sysdate + 1) >= Sysdate And a.��Աid = n_����Աid And a.�ͻ��� = '-') Loop
    v_��Ʊ����� := r_��Ʊ��.��Ʊ�����;
    Json_Out     := '{"output":{"code":1,"message": "�ɹ�","einvoice_code":"' || Nvl(v_��Ʊ�����, '') || '"}}';
  End Loop;

  --���ͻ��˶���
  For r_��Ʊ�� In (Select Nvl(a.��Աid, 0) As ��Աid, Nvl(a.�ͻ���, '-') As �ͻ���, b.���� As ��Ʊ�����
                From Ʊ�ݿ�Ʊ����� A, ����Ʊ�ݿ�Ʊ�� B
                Where a.��Ʊ��id = b.Id And Nvl(b.����ʱ��, Sysdate + 1) >= Sysdate And a.��Աid = 0 And a.�ͻ��� = v_�ͻ���) Loop
    v_��Ʊ����� := r_��Ʊ��.��Ʊ�����;
    Json_Out     := '{"output":{"code":1,"message": "�ɹ�","einvoice_code":"' || Nvl(v_��Ʊ�����, '') || '"}}';
  End Loop;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","einvoice_code":"' || Null || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoicecode;
/
CREATE OR REPLACE Procedure Zl_Exsesvr_Addeinvoiceerrdata
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�������Ʊ���쳣��¼
  --��Σ�Json_In:
  --input
  --  err_id              N 1 �쳣ID
  --  business_type       N 1 ҵ������:1-ҽ�ƿ�,2-Ԥ�� 
  --  business_id         N 1 ҵ��id:����ID��Ԥ��ID  
  --  occasion            N 1 ��������:1-ҽ�ƿ�����;2-������Ϣ�Ǽ�;3-������Ժ�Ǽ�
  --  record_sign         N 1 ��¼��־:0-���ߵ���Ʊ��;1-������Ʊ��;2-ֽ��Ʊ��;3-����ֽ��Ʊ��  
  --  einvoice_id         N 1 ����Ʊ��id
  --  pati_id             N 1 ����id
  --  pati_name           C 1 ����
  --  pati_sex            C 1 �Ա�
  --  pati_age            C 1 ����
  --  outpatient_num      C 1 �����
  --  inpatient_num       C 1 סԺ��
  --  err_no              C 1 ���ݺ�
  --  operator_code       C 1 ����Ա���
  --  operator_name       C 1 ����Ա����
  --  create_time         C   �Ǽ�ʱ��  ��ʽΪ:yyyy-mm-dd hh24:mi:ss
  --  is_turn             N   �Ƿ񻻿�
  --����: Json_Out,��ʽ����
  --output      
  --  code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message             C  1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ------------------------------------------------------------------------------------------------------------------------------

  j_Input      PLJson;
  j_Json       PLJson;
  n_�쳣id     ����Ʊ���쳣��¼.Id%Type;
  n_��������   ����Ʊ���쳣��¼.��������%Type;
  n_ҵ������   ����Ʊ���쳣��¼.ҵ������%Type;
  n_��¼��־   ����Ʊ���쳣��¼.��¼��־%Type;
  v_���ݺ�     ����Ʊ���쳣��¼.���ݺ�%Type;
  n_ҵ��id     ����Ʊ���쳣��¼.ҵ��id%Type;
  n_����Ʊ��id ����Ʊ���쳣��¼.����Ʊ��id%Type;
  n_����id     ����Ʊ���쳣��¼.����id%Type;
  v_����       ����Ʊ���쳣��¼.����%Type;
  v_�Ա�       ����Ʊ���쳣��¼.�Ա�%Type;
  v_����       ����Ʊ���쳣��¼.����%Type;
  n_�����     ����Ʊ���쳣��¼.�����%Type;
  n_סԺ��     ����Ʊ���쳣��¼.סԺ��%Type;
  v_����Ա��� ����Ʊ���쳣��¼.����Ա���%Type;
  v_����Ա���� ����Ʊ���쳣��¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ����Ʊ���쳣��¼.�Ǽ�ʱ��%Type;
  n_�Ƿ񻻿�   ����Ʊ���쳣��¼.�Ƿ񻻿�%Type;
Begin

  j_Input      := PLJson(Json_In);
  j_Json       := j_Input.Get_Pljson('input');
  n_�쳣id     := Pljson_Ext.Get_Number(j_Json, 'err_id');
  n_ҵ������   := Pljson_Ext.Get_Number(j_Json, 'business_type');
  n_ҵ��id     := Pljson_Ext.Get_Number(j_Json, 'business_id');
  n_��������   := Pljson_Ext.Get_Number(j_Json, 'occasion');
  n_��¼��־   := Pljson_Ext.Get_Number(j_Json, 'record_sign');
  v_����Ա��� := Pljson_Ext.Get_String(j_Json, 'operator_code');
  v_����Ա���� := Pljson_Ext.Get_String(j_Json, 'operator_name');
  d_�Ǽ�ʱ��   := To_Date(Pljson_Ext.Get_String(j_Json, 'create_time'), 'yyyy-mm-dd hh24:mi:ss');
  If d_�Ǽ�ʱ�� Is Null Then
    d_�Ǽ�ʱ�� := Sysdate;
  End If;
  v_���ݺ�     := Pljson_Ext.Get_String(j_Json, 'err_no');
  n_����Ʊ��id := Pljson_Ext.Get_Number(j_Json, 'einvoice_id');
  n_����id     := Pljson_Ext.Get_Number(j_Json, 'pati_id');
  v_����       := Pljson_Ext.Get_String(j_Json, 'pati_name');
  v_�Ա�       := Pljson_Ext.Get_String(j_Json, 'pati_sex');
  v_����       := Pljson_Ext.Get_String(j_Json, 'pati_age');
  n_�����     := Pljson_Ext.Get_Number(j_Json, 'outpatient_num');
  n_סԺ��     := Pljson_Ext.Get_Number(j_Json, 'inpatient_num');
  n_�Ƿ񻻿�   := Pljson_Ext.Get_Number(j_Json, 'is_turn');

  Zl_����Ʊ���쳣��¼_Insert(n_�쳣id, n_��������, n_ҵ������, n_��¼��־, v_���ݺ�, n_ҵ��id, n_����Ʊ��id, n_����id, v_����, v_�Ա�, v_����, n_�����, n_סԺ��,
                     v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��, n_�Ƿ񻻿�);
  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Addeinvoiceerrdata;
/
CREATE OR REPLACE Procedure Zl_Exsesvr_Modifyeinvoerrdata
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�������Ʊ���쳣��¼
  --��Σ�Json_In:
  --input
  --  err_id              N 1 �쳣id
  --  einvoice_info       C 1 Ʊ����Ϣ
  --  is_turn             N   �Ƿ񻻿�
  --����: Json_Out,��ʽ����
  --output      
  --  code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message             C  1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ------------------------------------------------------------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  n_�쳣id   ����Ʊ���쳣��¼.Id%Type;
  c_Ʊ����Ϣ ����Ʊ���쳣��¼.Ʊ����Ϣ%Type;
  n_�Ƿ񻻿� ����Ʊ���쳣��¼.�Ƿ񻻿�%Type;
Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�쳣id := Pljson_Ext.Get_Number(j_Json, 'err_id');
  Begin
    c_Ʊ����Ϣ := j_Json.Get_Clob('einvoice_info');
  Exception
    When Others Then
      Json_Out := '{"output":{"code":0,"message": "ʧ��,δ����Ʊ����Ϣ,����!"}}';
      Return;
  End;
  n_�Ƿ񻻿� := Pljson_Ext.Get_Number(j_Json, 'is_turn');
  Zl_����Ʊ���쳣��¼_Modify(n_�쳣id, c_Ʊ����Ϣ, n_�Ƿ񻻿�);

  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Modifyeinvoerrdata;
/
CREATE OR REPLACE Procedure Zl_Exsesvr_Deleteeinvoerrdata
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�������Ʊ���쳣��¼
  --��Σ�Json_In:
  --input
  --  err_id              N 1 �쳣id
  --����: Json_Out,��ʽ����
  --output      
  --  code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message             C  1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ------------------------------------------------------------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  n_�쳣id ����Ʊ���쳣��¼.Id%Type;
Begin

  j_Input  := PLJson(Json_In);
  j_Json   := j_Input.Get_Pljson('input');
  n_�쳣id := Pljson_Ext.Get_Number(j_Json, 'err_id');

  Zl_����Ʊ���쳣��¼_Delete(n_�쳣id);
  Json_Out := '{"output":{"code":1,"message": "�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Deleteeinvoerrdata;
/
Create Or Replace Procedure Zl_Exsesvr_Geteinvoiceerrdata
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ����Ʊ���쳣��¼
  --��Σ�Json_In:
  --input
  --  business_type       N 1 ҵ������:1-ҽ�ƿ�,2-Ԥ�� 
  --  business_id         N 1 ҵ��id:����ID��Ԥ��ID  
  --  occasion            N 1 ��������:1-ҽ�ƿ�����;2-������Ϣ�Ǽ�;3-������Ժ�Ǽ�
  --  record_sign         N 0 ��¼��־:0-���ߵ���Ʊ��;1-������Ʊ��;2-ֽ��Ʊ��;3-����ֽ��Ʊ��  
  --����: Json_Out,��ʽ����
  --output      
  --  code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message             C  1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  err_id              N  1 ����Ʊ���쳣id
  --  record_sign         N  0 ��¼��־:0-���ߵ���Ʊ��;1-������Ʊ��;2-ֽ��Ʊ��;3-����ֽ��Ʊ��  
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_�쳣id       ����Ʊ���쳣��¼.Id%Type;
  n_��������     ����Ʊ���쳣��¼.��������%Type;
  n_ҵ������     ����Ʊ���쳣��¼.ҵ������%Type;
  n_��¼��־     ����Ʊ���쳣��¼.��¼��־%Type;
  n_ҵ��id       ����Ʊ���쳣��¼.ҵ��id%Type;
  n_��¼��־_Out ����Ʊ���쳣��¼.��¼��־%Type;
Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_ҵ������ := Pljson_Ext.Get_Number(j_Json, 'business_type');
  n_ҵ��id   := Pljson_Ext.Get_Number(j_Json, 'business_id');
  n_�������� := Pljson_Ext.Get_Number(j_Json, 'occasion');
  n_��¼��־ := Pljson_Ext.Get_Number(j_Json, 'record_sign');

  If Nvl(n_ҵ������, 0) = 0 Or Nvl(n_ҵ��id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "ʧ��,�����ҵ�����ͻ�ҵ��idΪ0"}}';
    Return;
  End If;

  If Nvl(n_��������, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "ʧ��,����Ĳ�������Ϊ0"}}';
    Return;
  End If;

  If n_��¼��־ Is Null Then
    Select Max(ID), Max(��¼��־)
    Into n_�쳣id, n_��¼��־_Out
    From ����Ʊ���쳣��¼
    Where ҵ������ = n_ҵ������ And ҵ��id = n_ҵ��id And �������� = n_��������;
  Else
    Select Max(ID)
    Into n_�쳣id
    From ����Ʊ���쳣��¼
    Where ҵ������ = n_ҵ������ And ҵ��id = n_ҵ��id And ��¼��־ = n_��¼��־ And �������� = n_��������;
  End If;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","err_id":' || Nvl(n_�쳣id, 0) || ',"record_sign":' ||
              Nvl(n_��¼��־_Out, 0) || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoiceerrdata;
/
Create Or Replace Procedure Zl_Exsesvr_Geteinvoicedata
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:���ݽ���ID,��ȡ��Ч�ĵ���Ʊ��ID
  --��Σ�Json_In:
  --input
  --  fun_oper            N 1 �������ͣ�0-����Ʊ�ֺͽ���id��ȡ����Ʊ��ID��1-���ݵ���Ʊ��ID��ȡ �Ƿ񻻿���ֽ�ʷ�Ʊ�š�����id
  --  blnc_id             N   ����ID(����Ʊ��ʹ�ü�¼.����id)
  --  inv_type            N   Ʊ��:1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-ҽ�Ʒ��� 
  --  einvoice_id         N   ����Ʊ��ID
  --����: Json_Out,��ʽ����
  --output      
  --  code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message             C  1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  einvoice_id         N    ��Ч�ĵ���Ʊ��ID(��������=0ʱ����)
  --  blnc_id             N    ����ID(��������=1ʱ����)
  --  is_turn             N    �Ƿ񻻿�(��������=1ʱ����)
  --  inv_no              N    ֽ��Ʊ��(��������=1ʱ����)
  ---------------------------------------------------------------------------
  j_Input          PLJson;
  j_Json           PLJson;
  n_��������       Number(2);
  n_Ʊ��           ����Ʊ��ʹ�ü�¼.Ʊ��%Type;
  n_�Ƿ񻻿�       ����Ʊ��ʹ�ü�¼.�Ƿ񻻿�%Type;
  v_ֽ�ʷ�Ʊ��     ����Ʊ��ʹ�ü�¼.ֽ�ʷ�Ʊ��%Type;
  n_����id         ����Ʊ��ʹ�ü�¼.����id%Type;
  n_����id_Out     ����Ʊ��ʹ�ü�¼.����id%Type;
  n_����Ʊ��id     ����Ʊ��ʹ�ü�¼.Id%Type;
  n_����Ʊ��id_Out ����Ʊ��ʹ�ü�¼.Id%Type;
Begin
  --�������

  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_�������� := Nvl(Pljson_Ext.Get_Number(j_Json, 'fun_oper'), 0);

  If n_�������� = 0 Then
    --����Ʊ�ֺͽ���id��ȡ����Ʊ��ID
    n_����id := Pljson_Ext.Get_Number(j_Json, 'blnc_id');
    n_Ʊ��   := Pljson_Ext.Get_Number(j_Json, 'inv_type');
    If (Nvl(n_����id, 0) = 0 Or Nvl(n_Ʊ��, 0) = 0) Then
      Json_Out := '{"output":{"code":0,"message": "ʧ��,����Ľ���id�򳡺�Ϊ0"}}';
      Return;
    End If;
  
    Select Max(ID)
    Into n_����Ʊ��id_Out
    From ����Ʊ��ʹ�ü�¼
    Where ����id = n_����id And Ʊ�� = n_Ʊ�� And ��¼״̬ = 1 And Nvl(ԭƱ��id, 0) = 0;
  
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","einvoice_id":' || Nvl(n_����Ʊ��id_Out, 0) || '}}';
  Else
    --���ݵ���Ʊ��ID��ȡ �Ƿ񻻿���ֽ�ʷ�Ʊ�š�����id
    n_����Ʊ��id := Pljson_Ext.Get_Number(j_Json, 'einvoice_id');
    If Nvl(n_����Ʊ��id, 0) = 0 Then
      Json_Out := '{"output":{"code":0,"message": "ʧ��,����ĵ���Ʊ��IDΪ0"}}';
      Return;
    End If;
  
    Select Max(�Ƿ񻻿�), Max(ֽ�ʷ�Ʊ��), Max(����id)
    Into n_�Ƿ񻻿�, v_ֽ�ʷ�Ʊ��, n_����id_Out
    From ����Ʊ��ʹ�ü�¼
    Where ID = n_����Ʊ��id;
  
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","blnc_id":' || Nvl(n_����id_Out, 0) || ',"is_turn":' ||
                Nvl(n_�Ƿ񻻿�, 0) || ',"inv_no":"' || v_ֽ�ʷ�Ʊ�� || '"}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoicedata;
/
Create Or Replace Procedure Zl_Exsesvr_Checkiseinvoice
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:���ݴ����Ʊ�ֺͽ���ID,��鵱ǰ�����Ƿ��������е���Ʊ��
  --��Σ�Json_In:
  --input
  --  blnc_id             N 1 ����ID(����Ʊ��ʹ�ü�¼.id)
  --  inv_type            N 1 Ʊ��:1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-ҽ�Ʒ��� 
  --����: Json_Out,��ʽ����
  --output      
  --  code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message             C  1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  is_einvoice         N  1 �Ƿ����õ���Ʊ��:1-����;0:δ����
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Ʊ��     ����Ʊ��ʹ�ü�¼.Ʊ��%Type;
  n_����id   ����Ʊ��ʹ�ü�¼.Id%Type;
  n_Einvoice Number(2);
Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := Pljson_Ext.Get_Number(j_Json, 'blnc_id');
  n_Ʊ��   := Pljson_Ext.Get_Number(j_Json, 'inv_type');

  If Nvl(n_����id, 0) = 0 Or Nvl(n_Ʊ��, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "ʧ��,����Ľ���id�򳡺�Ϊ0"}}';
    Return;
  End If;

  If n_Ʊ�� = 2 Then
    --Ԥ����¼
    Select Max(Ԥ������Ʊ��) Into n_Einvoice From ����Ԥ����¼ Where Mod(��¼����, 10) = 1 And ID = n_����id;
  Else
    Select Max(�Ƿ����Ʊ��) Into n_Einvoice From ����Ԥ����¼ Where ����id = n_����id;
  End If;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","is_einvoice":' || Nvl(n_Einvoice, 0) || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkiseinvoice;
/
Create Or Replace Procedure Zl_Exsesvr_Geteinvoiceinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ����Ʊ����Ϣ
  --��Σ�Json_In:
  --input
  --err_id              N 1 �쳣ID
  --����: Json_Out,��ʽ����
  --output      
  --code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --message             C  1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --��¼��־=0 ʱ ����----------------------------------------
  --data
  --  input          
  --    balance_id        N  1  ����ID
  --    balance_delid     N     �˿�ID:�˿�ߺ�Ʊʱ��Ч��Ŀǰֻ��Ԥ������Ч,��д�����˿�Ԥ��ID
  --    einvoice_id       N  1  ����Ʊ��ID
  --    operator_code     C  1  ����Ա���
  --    operator_name     C  1  ����Ա����
  --    create_time       C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    pati_info         C     ������Ϣ
  --      pati_id         N  1  ����ID
  --      pati_pageid     N     ��ҳID
  --      pati_name       C  1  ����
  --      pati_sex        C  1  �Ա�
  --      pati_age        C  1  ����
  --      outpatient_num  C  1  �����
  --      inpatient_num   C  1  סԺ��
  --    einvoce_info      C     ����Ʊ����Ϣ
  --      invoice_type    N  1  Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --      placeCode       C  1  ��Ʊ�����
  --      inv_total       N  1  ��Ʊ���
  --      inv_oldid       N    ԭƱ��ID
  --      sys_source      C  1  ϵͳ��Դ
  --      demo            C  1  ��ע
  --      einvoice_code   C  1  ����Ʊ�ݴ���
  --      einvoice_no     C  1  ����Ʊ�ݺ���
  --      einvoice_random C  1  ����У����
  --      voucher_code    C  1  Ԥ����ƾ֤����
  --      voucher_no      C  1  Ԥ����ƾ֤����
  --      voucher_random  C  1  Ԥ����ƾ֤У����
  --      happen_time     C  1  ����Ʊ������ʱ��:yyyymmddhh24miss
  --      picture_url     C  1  ����Ʊ��H5ҳ��URL
  --      picture_neturl  C  1  ����Ʊ������H5ҳ��URL
  --      qrcode          C  1  ����Ʊ�ݶ�ά��ͼƬ����:��ֵ��Base64���룬����ʱ��ҪBase64����,ͼƬ��ʽΪ:PNG
  --��¼��־=1 ʱ ����-----------------------------------------
  --data
  --  input         
  --    einvoice_id       N 1 ����Ʊ��ID
  --    operator_code     C 1 ����Ա���
  --    operator_name     C 1 ����Ա����
  --    create_time       C 1 �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    einvoce_info      C   ����Ʊ����Ϣ
  --      placeCode       C 1 ��Ʊ�����
  --      sys_source      C 1 ϵͳ��Դ
  --      demo            C 1 ��ע
  --      inv_oldid       N   ԭƱ��ID
  --      einvoice_code   C 1 ����Ʊ�ݴ���
  --      einvoice_no     C 1 ����Ʊ�ݺ���
  --      einvoice_random C 1 ����У����
  --      voucher_code    C 1 Ԥ����ƾ֤����
  --      voucher_no      C 1 Ԥ����ƾ֤����
  --      voucher_random  C 1 Ԥ����ƾ֤У����
  --      happen_time     C 1 ����Ʊ������ʱ��:yyyymmddhh24miss
  --      picture_url     C 1 ����Ʊ��H5ҳ��URL
  --      picture_neturl  C 1 ����Ʊ������H5ҳ��URL
  --      qrcode          C 1 ����Ʊ�ݶ�ά��ͼƬ����:��ֵ��Base64���룬����ʱ��ҪBase64����,ͼƬ��ʽΪ:PNG
  --��¼��־=2,3 ʱ ����-------------------------------------------
  --data
  --  input         
  --    oper_mode         N 1 ������ʽ:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
  --    einvoice_id       N 1 ����Ʊ��ID
  --    operator_code     C 1 ����Ա���
  --    operator_name     C 1 ����Ա����
  --    create_time       C 1 �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    paperinv_info     C   ֽ��Ʊ����Ϣ:���ڶ���ʱ���밴����˳���ϴ�(�������ݴ���)
  --      inv_occasion    N 1 Ӧ�ó���:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
  --      invoice_type    N 1 Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --      inv_red         N   �Ƿ��Ʊ:1-��Ʊ;0-�Ǻ�Ʊ
  --      invoice_no      C 1 ��Ʊ��
  --      inv_total       N 1 ��Ʊ���
  --      recv_id         N   ����id
  ---------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  c_Output Clob;
  n_�쳣id ����Ʊ���쳣��¼.Id%Type;

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�쳣id := Pljson_Ext.Get_Number(j_Json, 'err_id');

  If Nvl(n_�쳣id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "ʧ��,������쳣idΪ0"}}';
    Return;
  End If;

  Begin
    Select Ʊ����Ϣ Into c_Output From ����Ʊ���쳣��¼ Where ID = n_�쳣id;
  Exception
    When Others Then
      Json_Out := '{"output":{"code":0,"message": "ʧ��,���ݴ�����쳣idδ�ҵ�����"}}';
      Return;
  End;

  If c_Output Is Null Then
    Json_Out := '{"output":{"code":0,"message": "ʧ��,���ݴ�����쳣idδ�ҵ�����"}}';
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","data":' || c_Output || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoiceinfo;
/
Create Or Replace Procedure Zl_����Ԥ����¼_Insert_s
(
  Id_In           ����Ԥ����¼.Id%Type,
  ���ݺ�_In       ����Ԥ����¼.No%Type,
  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ��ҳid_In       ����Ԥ����¼.��ҳid%Type,
  ����_In         ����Ԥ����¼.����%Type,
  �Ա�_In         ����Ԥ����¼.�Ա�%Type,
  ����_In         ����Ԥ����¼.����%Type,
  �����_In       ����Ԥ����¼.�����%Type,
  סԺ��_In       ����Ԥ����¼.סԺ��%Type,
  ���ʽ����_In ����Ԥ����¼.���ʽ����%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ���_In         ����Ԥ����¼.���%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type,
  �������_In     ����Ԥ����¼.�������%Type,
  �ɿλ_In     ����Ԥ����¼.�ɿλ%Type,
  ��λ������_In   ����Ԥ����¼.��λ������%Type,
  ��λ�ʺ�_In     ����Ԥ����¼.��λ�ʺ�%Type,
  ժҪ_In         ����Ԥ����¼.ժҪ%Type,
  ����Ա���_In   ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In   ����Ԥ����¼.����Ա����%Type,
  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
  Ԥ�����_In     ����Ԥ����¼.Ԥ�����%Type := Null,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In     ����Ԥ����¼.������λ%Type := Null,
  �տ�ʱ��_In     ����Ԥ����¼.�տ�ʱ��%Type := Null,
  ����id_In       ����Ԥ����¼.����id%Type := Null,
  ��������_In     ����Ԥ����¼.��������%Type := Null,
  ���½������_In Number := 1,
  ����״̬_In     Number := 0,
  ��������id_In   ����Ԥ����¼.��������id%Type := Null,
  У�Ա�־_In     ����Ԥ����¼.У�Ա�־%Type := Null,
  Ԥ������Ʊ��_In ����Ԥ����¼.Ԥ������Ʊ��%Type := Null,
  ����_In         ���ս����¼.����%Type := Null
) As
  --------------------------------------------------------------------------------------------
  --���ܣ�����Ԥ����¼
  --����״̬_In:0-��������,1-����Ϊ�쳣���ݻ�δ��Ч�ĵ���,2-����쳣����,3-����Ԥ����¼����
  --����ID_IN:>0ʱ,��ʾĳ�ν���ʱ,ͬ��������Ԥ����¼(�����տ�����ΪԤ��)
  --���½������_In:0-�� zl_��Ա�ɿ����_Update �и���(��Ҫ��������ֵʱ��ֹ���ܱ�����)��1-�ڱ������и���
  --------------------------------------------------------------------------------------------

  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  v_����   ���㷽ʽ.����%Type;
  v_��ӡid Ʊ�ݴ�ӡ����.Id%Type;

  v_Date Date;

  n_����ֵ       �������.Ԥ�����%Type;
  n_��id         ����ɿ����.Id%Type;
  n_����         ���ս����¼.����%Type;
  n_Ԥ������Ʊ�� Number(2);
Begin
  v_Date := �տ�ʱ��_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  n_��id := Zl_Get��id(����Ա����_In);

  n_Ԥ������Ʊ�� := Ԥ������Ʊ��_In;
  If n_Ԥ������Ʊ�� Is Null Then
    n_���� := ����_In;
    If ����_In Is Null Then
      Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = Id_In And ���� = 3;
    End If;
    n_Ԥ������Ʊ�� := Zl_Fun_Isstarteinvoice(2, n_����, Ԥ�����_In);
  End If;

  Select Max(����) Into v_���� From ���㷽ʽ Where ���� = ���㷽ʽ_In;

  --����״̬_In��0-��������,1-����Ϊ�쳣����,2-����쳣����,3-����Ԥ����¼����
  If Nvl(����״̬_In, 0) < 2 Then
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����, �Ա�, ����, �����, סԺ��, ���ʽ����, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������,
       ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������, У�Ա�־, ��������id, ����ʱ��,
       ������Ա, Ԥ������Ʊ��)
    Values
      (Id_In, ���ݺ�_In, Decode(����״̬_In, 1, Null, Ʊ�ݺ�_In), 1, Decode(����״̬_In, 1, 0, 1), ����id_In,
       Decode(��ҳid_In, 0, Null, ��ҳid_In), ����_In, �Ա�_In, ����_In, Decode(�����_In, 0, Null, �����_In),
       Decode(סԺ��_In, 0, Null, סԺ��_In), ���ʽ����_In, Decode(����id_In, 0, Null, ����id_In), ���_In, ���㷽ʽ_In, �������_In, v_Date,
       �ɿλ_In, ��λ������_In, ��λ�ʺ�_In, ����Ա���_In, ����Ա����_In, ժҪ_In, n_��id, Ԥ�����_In, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In,
       ����˵��_In, ������λ_In, ����id_In, Decode(Nvl(����id_In, 0), 0, Null, 0), ��������_In, Decode(����״̬_In, 1, 1, Null),
       Decode(��������id_In, Null, Id_In, 0, Null, ��������id_In), �տ�ʱ��_In, ����Ա����_In, n_Ԥ������Ʊ��);
    If Nvl(�����id_In, 0) <> 0 Then
      --�Զ�����̵���
      Zl_Custom_Balance_Update(Id_In);
    End If;
  End If;

  --����״̬_In��0-��������,1-����Ϊ�쳣����,2-����쳣����,3-����Ԥ����¼����
  If Nvl(����״̬_In, 0) = 1 Then
    --����Ϊ�쳣����
    Return;
  End If;
  --����״̬_In��0-��������,1-����Ϊ�쳣����,2-����쳣����,3-����Ԥ����¼����
  If Nvl(����״̬_In, 0) = 0 Or Nvl(����״̬_In, 0) = 2 Then
    Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (Id_In, ����id_In, Ԥ�����_In, ���_In);
    --�������(Ԥ���������)
    If Nvl(v_����, 1) <> 5 Then
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
      Where ���� = 1 And ����id = ����id_In And Nvl(����, 0) = Nvl(Ԥ�����_In, 0)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (����id_In, 1, Nvl(Ԥ�����_In, 0), ���_In, 0);
        n_����ֵ := ���_In;
      End If;
      If Nvl(���_In, 0) = 0 Then
        Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End If;
  End If;
  --�����쳣���ݻ�����Ԥ������
  --����״̬_In��0-��������,1-����Ϊ�쳣����,2-����쳣����,3-����Ԥ����¼����
  If Nvl(����״̬_In, 0) = 2 Or Nvl(����״̬_In, 0) = 3 Then
    --���²��������
    Update ����Ԥ����¼
    Set ��¼״̬ = 1, У�Ա�־ = Decode(����״̬_In, 2, Null, У�Ա�־_In), ʵ��Ʊ�� = Nvl(Ʊ�ݺ�_In, ʵ��Ʊ��), �տ�ʱ�� = Nvl(v_Date, �տ�ʱ��),
        ����Ա��� = Nvl(����Ա���_In, ����Ա���), ����Ա���� = Nvl(����Ա����_In, ����Ա����), �ɿ���id = Nvl(n_��id, �ɿ���id), ����ʱ�� = Nvl(v_Date, ����ʱ��),
        ������Ա = Nvl(����Ա����_In, ������Ա), ����id = Nvl(����id_In, ����id), ��� = Nvl(���_In, ���), ���㷽ʽ = Nvl(���㷽ʽ_In, ���㷽ʽ),
        ������� = Nvl(�������_In, �������), �ɿλ = Nvl(�ɿλ_In, �ɿλ), ��λ������ = Nvl(��λ������_In, ��λ������), ��λ�ʺ� = Nvl(��λ�ʺ�_In, ��λ�ʺ�),
        ժҪ = Nvl(ժҪ_In, ժҪ), �����id = Nvl(�����id_In, �����id), ���㿨��� = Nvl(���㿨���_In, ���㿨���), ���� = Nvl(����_In, ����),
        ������ˮ�� = Nvl(������ˮ��_In, ������ˮ��), ����˵�� = Nvl(����˵��_In, ����˵��), Ԥ������Ʊ�� = Nvl(n_Ԥ������Ʊ��, Ԥ������Ʊ��)
    Where ID = Id_In;
    --�Զ�����̵���
    Zl_Custom_Balance_Update(Id_In);
    If ����״̬_In = 3 Then
      Return;
    End If;
  End If;

  --����Ʊ��
  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
  
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 2, ���ݺ�_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 2, Ʊ�ݺ�_In, 1, 1, ����id_In, v_��ӡid, v_Date, ����Ա����_In, ���_In);
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  End If;

  --��ػ��ܱ�����Ա�ɿ����(����)
  If Nvl(���½������_In, 1) = 1 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ���_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = ���㷽ʽ_In
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, ���㷽ʽ_In, 1, ���_In);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
    End If;
  End If;

  If ���_In < 0 Then
    b_Message.Zlhis_Charge_006(Id_In, ���ݺ�_In);
  Else
    b_Message.Zlhis_Charge_005(Id_In, ���ݺ�_In);
  End If;
  --��Ϣ����;
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 11, Id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_Insert_s;
/
Create Or Replace Procedure Zl_����Ԥ����¼_����˿�_s
(
  Id_In           ����Ԥ����¼.Id%Type,
  ���ݺ�_In       ����Ԥ����¼.No%Type,
  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ��ҳid_In       ����Ԥ����¼.��ҳid%Type,
  ����_In         ����Ԥ����¼.����%Type,
  �Ա�_In         ����Ԥ����¼.�Ա�%Type,
  ����_In         ����Ԥ����¼.����%Type,
  �����_In       ����Ԥ����¼.�����%Type,
  סԺ��_In       ����Ԥ����¼.סԺ��%Type,
  ���ʽ����_In ����Ԥ����¼.���ʽ����%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ���_In         ����Ԥ����¼.���%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type,
  �������_In     ����Ԥ����¼.�������%Type,
  �ɿλ_In     ����Ԥ����¼.�ɿλ%Type,
  ��λ������_In   ����Ԥ����¼.��λ������%Type,
  ��λ�ʺ�_In     ����Ԥ����¼.��λ�ʺ�%Type,
  ժҪ_In         ����Ԥ����¼.ժҪ%Type,
  ����Ա���_In   ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In   ����Ԥ����¼.����Ա����%Type,
  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
  Ԥ�����_In     ����Ԥ����¼.Ԥ�����%Type := Null,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ��������id_In   ����Ԥ����¼.��������id%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In     ����Ԥ����¼.������λ%Type := Null,
  �տ�ʱ��_In     ����Ԥ����¼.�տ�ʱ��%Type := Null,
  ������Ϣ_In     Varchar2 := Null,
  ����������_In   Number := 0,
  ����״̬_In     Number := 0,
  ����id_In       ����Ԥ����¼.����id%Type := Null,
  �������_In     ����Ԥ����¼.�������%Type := Null,
  Ԥ������Ʊ��_In ����Ԥ����¼.Ԥ������Ʊ��%Type := Null,
  ����_In         ���ս����¼.����%Type := Null
) As
  ----------------------------------------------
  --����˿����
  --������Ϣ_In:ԭԤ��ID|���||....
  --����������_IN:0-��ʾ��Ҫ����Ԥ����¼�����²������;1-��ʾֻ���½�����Ϣ�е���������
  --����״̬_IN:0-��ʾ��ɽ���;1-��ʾδ��ɽ���; (����״̬_IN=1ʱ,���ɵ�Ԥ����¼��У�Ա�־Ϊ1)
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  n_��ӡid       Ʊ�ݴ�ӡ����.Id%Type;
  d_�տ�ʱ��     Date;
  n_����ֵ       �������.Ԥ�����%Type;
  n_��id         ����ɿ����.Id%Type;
  n_����id       ����Ԥ����¼.����id%Type;
  n_�������     ����Ԥ����¼.�������%Type;
  n_Count        Number(18);
  n_Ԥ�����     �������.Ԥ�����%Type;
  n_Ԥ������Ʊ�� Number(2);
  n_����         ���ս����¼.����%Type;
Begin
  n_Ԥ������Ʊ�� := Ԥ������Ʊ��_In;
  If n_Ԥ������Ʊ�� Is Null Then
    n_���� := ����_In;
    If ����_In Is Null Then
      Select Nvl(Max(����), 0) Into n_���� From ���ս����¼ Where ��¼id = Id_In And ���� = 3;
    End If;
    n_Ԥ������Ʊ�� := Zl_Fun_Isstarteinvoice(2, n_����, Ԥ�����_In);
  End If;

  n_��id := Zl_Get��id(����Ա����_In);
  If ����������_In = 0 Then
    d_�տ�ʱ�� := �տ�ʱ��_In;
    If d_�տ�ʱ�� Is Null Then
      Select Sysdate Into d_�տ�ʱ�� From Dual;
    End If;
    n_������� := �������_In;
    n_����id   := ����id_In;
    If Nvl(n_����id, 0) = 0 Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    End If;
    If Nvl(n_�������, 0) = 0 Then
      n_������� := -1 * n_����id;
    End If;
    --Ϊ�˲������������������(���_InΪ����)
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
    Where ����id = ����id_In And ���� = Ԥ�����_In And ���� = 1
    Returning Ԥ����� Into n_Ԥ�����;
    If Sql%RowCount = 0 Then
      Insert Into �������
        (����id, ����, ����, Ԥ�����, �������)
      Values
        (����id_In, 1, Nvl(Ԥ�����_In, 0), ���_In, 0);
      n_Ԥ����� := ���_In;
    End If;
  
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����, �Ա�, ����, �����, סԺ��, ���ʽ����, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������,
       ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��Ԥ��, ��������, У�Ա�־, ��������id,
       ����ʱ��, ������Ա, ���ӱ�־, Ԥ������Ʊ��)
    Values
      (Id_In, ���ݺ�_In, Decode(����״̬_In, 0, Ʊ�ݺ�_In, Null), 1, 0, ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In), ����_In, �Ա�_In,
       ����_In, Decode(�����_In, 0, Null, �����_In), Decode(סԺ��_In, 0, Null, סԺ��_In), ���ʽ����_In,
       Decode(����id_In, 0, Null, ����id_In), ���_In, ���㷽ʽ_In, �������_In, d_�տ�ʱ��, �ɿλ_In, ��λ������_In, ��λ�ʺ�_In, ����Ա���_In,
       ����Ա����_In, ժҪ_In, n_��id, Ԥ�����_In, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, n_����id, n_�������, Null,
       Null, ����״̬_In, Decode(Nvl(��������id_In, 0), 0, Id_In, ��������id_In), �տ�ʱ��_In, ����Ա����_In, 1, n_Ԥ������Ʊ��);
  
    --����Ԥ���������
    Insert Into Ԥ��������� (Ԥ��id, ����id, Ԥ�����, Ԥ�����) Values (Id_In, ����id_In, Ԥ�����_In, ���_In);
  End If;

  If ����������_In = 1 Then
    Select Max(����id), Max(�տ�ʱ��), Max(1) Into n_����id, d_�տ�ʱ��, n_Count From ����Ԥ����¼ Where ID = Id_In;
    If n_Count = 0 Then
      v_Err_Msg := 'δ�ҵ��˿��¼�����飡';
      Raise Err_Item;
    End If;
  End If;

  If ������Ϣ_In Is Not Null Then
    Zl_����Ԥ����¼_Relevance(����id_In, Id_In, ������Ϣ_In, n_����id, ����Ա���_In, ����Ա����_In, �տ�ʱ��_In, ����״̬_In, n_��id);
  End If;

  If ����״̬_In = 1 Then
    Return;
  End If;

  --���¼�¼״̬1
  Update ����Ԥ����¼
  Set ��¼״̬ = 1, У�Ա�־ = 0, ʵ��Ʊ�� = Ʊ�ݺ�_In
  Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 0
  Returning ����id Into n_����id;
  If Sql%NotFound Then
    v_Err_Msg := 'δ�ҵ�ָ���ĵ���(' || ���ݺ�_In || ',������Ϊ����ԭ�������˿���飡';
    Raise Err_Item;
  End If;
  Update ����Ԥ����¼ Set У�Ա�־ = 0 Where ����id = n_����id And Nvl(У�Ա�־, 0) <> 0;

  --����Ʊ��
  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
  
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 2, ���ݺ�_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 2, Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_�տ�ʱ��, ����Ա����_In, ���_In);
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  
    Update ����Ԥ����¼ Set ʵ��Ʊ�� = Ʊ�ݺ�_In Where ����id = ����id_In And ��¼���� = 11 And NO = ���ݺ�_In;
  
  End If;

  --��Ա�ɿ����(����)
  Update ��Ա�ɿ����
  Set ��� = Nvl(���, 0) + ���_In
  Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = ���㷽ʽ_In
  Returning ��� Into n_����ֵ;

  If Sql%RowCount = 0 Then
    Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, ���㷽ʽ_In, 1, ���_In);
    n_����ֵ := ���_In;
  End If;
  If Nvl(n_����ֵ, 0) = 0 Then
    Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
  End If;

  If Nvl(n_Ԥ�����, 0) = 0 Then
    Delete From �������
    Where ����id = ����id_In And ���� = Ԥ�����_In And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0 And ���� = 1;
  End If;

  If ���_In < 0 Then
    b_Message.Zlhis_Charge_006(Id_In, ���ݺ�_In);
  Else
    b_Message.Zlhis_Charge_005(Id_In, ���ݺ�_In);
  End If;

  --��Ϣ����;
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 11, Id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_����˿�_s;
/