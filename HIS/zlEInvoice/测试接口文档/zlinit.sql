

INSERT INTO ����Ʊ����� (���,����,����,�Ƿ�����,����,������) values(2,'����ģ��ӿ�(V1.0.0)','BSDZBJPT',1,'zlEInvoice.clsEInvoice_Test','b_Einvoice_Request_Test');
 
insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',1,upper('URL_Type'),'HTTP',NULL);
insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',2,upper('URL_Address'),'','<ip>:<port>/<service>/api/medical/�ӿڷ����ʶ');
insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',3,'Ӧ���ʺ�','','��Appid');
insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',4,'ǩ��˽Կ','','��KEYֵ');
insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',5,'֧�ְ汾','V2.0.3','Ŀǰֻ֧��:V2.0.3��V3.1.0');
insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',6,'���ݴ��䷽ʽ','application/json','�ύ�ͷ������ݿ���ΪJSON��ʽ��Content-Type: application/json��');
insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',7,'�ַ�����','UTF-8','ͳһ����UTF-8�ַ�����');

insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',8,'ȱʡ�����ID','','ȱʡ��ȡ�Ŀ����');
insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',9,'���֤�������ͱ��','999998','ʹ�����֤��Ϊ�ϴ��Ŀ����͵ı��');
insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',10,'�����޿��Ŀ������','999999','�������κο�ʱ�ϴ��Ŀ����ͱ��');
insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',11,'�����޿��Ŀ���','-','�������κο�ʱ�ϴ��Ŀ���');

--insert into �����ӿ����� ( �ӿ���,������,������,����ֵ,˵��) values ('����ģ��ӿ�(V1.0.0)',12,'�����Ŀ','','��С����HIS��ƥ��ʱ���������ѵ��շ���Ŀ');



CREATE TABLE �շ���������(
	���㷽ʽ varchar2(20),
	�����ID Number(18),
	�������� varchar2(20))
TABLESPACE zl9Expense;
Alter Table �շ��������� Add Constraint �շ���������_UQ_��Ʊ��ID  Unique(���㷽ʽ,�����ID, ��������)  Using Index Tablespace zl9Indexhis;
Alter Table �շ��������� Add Constraint �շ���������_FK_�����ID Foreign Key (�����ID) References ҽ�ƿ����(ID) on delete cascade;
 

--����ͳ����
CREATE TABLE ��Ʊ�������(
	���㷽ʽ varchar2(20), --HIS���㷽ʽ
	��Ʊ���㷽ʽ varchar2(20)) --ֻ������:�����˻�֧��,ҽ��ͳ�����֧��,����ҽ��֧��
TABLESPACE zl9Expense;
Alter Table ��Ʊ������� Add Constraint ��Ʊ�������_UQ_���㷽ʽ Unique(���㷽ʽ,��Ʊ���㷽ʽ )  Using Index Tablespace zl9Indexhis; 


--ҩƷ֧��������
CREATE TABLE ֧��������(
	���մ���ID number(18),  
	������� varchar2(20),
	�������� varchar2(50)) --1�����Ը�/��;2�����Ը�/��;3��ȫ�Ը�/��
TABLESPACE zl9Expense;
Alter Table ֧�������� Add Constraint ֧��������_UQ_���մ���ID Unique(���մ���ID,�������)  Using Index Tablespace zl9Indexhis; 
Alter Table ֧�������� Add Constraint ֧��������FK_���մ���ID Foreign Key (���մ���ID) References ����֧������(ID) on delete cascade;


Create Table ����Ʊ�ݺ˶Լ�¼(
 ҵ������ Date,
 ��Ʊ�� Varchar2(100),
 His��Ʊ�� Number(18),
 His��Ʊ��� Number(18),
 ƽ̨��Ʊ�� Number(16, 5),
 ƽ̨��Ʊ��� Number(16, 5),
 �˶����� Number(1),
 �˶��� Varchar2(50),
 �˶�ʱ�� Date,
 �˶Խ�� Number(1),
 �˶�˵�� Varchar2(4000)
) Tablespace Zl9Expense;

Alter Table ����Ʊ�ݺ˶Լ�¼ Add Constraint ����Ʊ�ݺ˶Լ�¼_Uq_ҵ������ Unique(ҵ������, ��Ʊ��, �˶�����, �˶���) Using Index Tablespace Zl9indexhis;
Alter Table ����Ʊ�ݺ˶Լ�¼ Modify ҵ������ Constraint ����Ʊ�ݺ˶Լ�¼_NN_ҵ������ Not Null;
Alter Table ����Ʊ�ݺ˶Լ�¼ Modify �˶����� Constraint ����Ʊ�ݺ˶Լ�¼_NN_�˶����� Not Null;
Alter Table ����Ʊ�ݺ˶Լ�¼ Modify �˶��� Constraint ����Ʊ�ݺ˶Լ�¼_NN_�˶��� Not Null;
Alter Table ����Ʊ�ݺ˶Լ�¼ Modify �˶�ʱ�� Constraint ����Ʊ�ݺ˶Լ�¼_NN_�˶�ʱ�� Not Null;
Alter Table ����Ʊ�ݺ˶Լ�¼ Modify �˶Խ�� Constraint ����Ʊ�ݺ˶Լ�¼_NN_�˶Խ�� Not Null;
 
Create Table ����Ʊ��������¼(
 ҵ������ Date,
 ����Ʊ��ID Number(18),
 ҵ����ˮ�� Varchar2(50),
 HIS��Ʊ�� Varchar2(100),
 HIS��Ʊ��� Number(16,5),
 HISƱ��״̬ Number(1),
 ƽ̨��Ʊ�� Varchar2(100),
 ƽ̨��Ʊ��� Number(16,5),
 ƽ̨Ʊ��״̬ Number(1),
 ������ʽ NUmber(1),
 ������ Varchar2(50),
 ����ʱ�� Date,
 ������� Number(1),
 ����˵�� Varchar2(4000)
) Tablespace Zl9Expense;

Alter Table ����Ʊ��������¼ Add Constraint ����Ʊ��������¼_Uq_ҵ������ Unique(ҵ������, ����Ʊ��ID, ҵ����ˮ��, ƽ̨Ʊ��״̬) Using Index Tablespace Zl9indexhis;
Alter Table ����Ʊ��������¼ Modify ҵ������ Constraint ����Ʊ��������¼_NN_ҵ������ Not Null;
Alter Table ����Ʊ��������¼ Modify ������ʽ Constraint ����Ʊ��������¼_NN_������ʽ Not Null;
Alter Table ����Ʊ��������¼ Modify ������ Constraint ����Ʊ��������¼_NN_������ Not Null;
Alter Table ����Ʊ��������¼ Modify ����ʱ�� Constraint ����Ʊ��������¼_NN_����ʱ�� Not Null;
Alter Table ����Ʊ��������¼ Modify ������� Constraint ����Ʊ��������¼_NN_������� Not Null;

CONNECT sys@system AS sysdba;
grant execute on dbms_crypto to zlhis

CONNECT zlhis@his ;

Create Or Replace Package b_Einvoice_Request_Test Is
  ------------------------------------------------------------------
  --����Ʊ��ҵ����(��Ҫ����Բ�˼����Ʊ�ݵĴ���)
  --һ.��������˵��
  --  0.get_charset-��ȡ��ǰ���ݿ���ַ���
  --  1.zlJsonStr-����Ҫ��ϳ�Json�����ַ����е��������
  --  2.MD5_Clob-���Clob����MD5����
  --  3.Request-����ҵ������   
  --  3.GetRequestData_Encode-��ȡ��������(����)
  --  4.Getrequestdata_Decode-��ȡҵ������(����)
  --  5.Get_ParaInfo-��ȡ��˼��ز�������
  --  6.Get_Version-��ȡ��ǰ��˼�İ汾��
  --  7 Get_IDENTITY-��ȡ����Ա��Ϣ
  --��.����Ʊ�ݴ���
  --1.��ȡҵ���������
  --1.1 Get_Einvoice_Node-��ȡ��Ʊ�㺯��
  --1.2 Get_Chargedata_Create-��ȡ�շѿ�ƱƱ��
  --1.3 Get_SendCarddata_Create-��ȡ������ƱƱ��
  --1.4 Get_Registerdata_Create-��ȡ�Һſ�ƱƱ��
  --1.5 Get_MZBalancedata_Create-��ȡ������ʿ�ƱƱ��
  --1.6 Get_ZYBalancedata_Create-��ȡסԺ���ʿ�ƱƱ��
  --1.7 Get_Depositdata_Create-��ȡԤ����ƱƱ��
  --2.����Ʊ�ݲ������
  --2.1.Einvoice_Start-����Ʊ���Ƿ�����
  --2.2.EInvoice_Create-����Ʊ�ݿ���
  --2.3.Einvoice_Cancel_Check-����Ʊ������ǰ���(����:1-�Ϸ�;0-���Ϸ�)
  --2.4.Einvoice_Cancel-����Ʊ������(����1-�ɹ�;0-ʧ��)
  ------------------------------------------------------------------
  Function Get_Charset Return Number;

  Function zlJsonStr
  (
    Str_In  In Varchar2,
    Type_In Number := 0
  ) Return Varchar2;

  Procedure Get_Identity
  (
    ��Աid_Out     Out ��Ա��.Id%Type,
    ��Ա���_Out   Out ��Ա��.���%Type,
    ����Ա����_Out Out ��Ա��.����%Type
  );

  Function Md5_Clob(Souce_In In Clob) Return Varchar2;

  Function Getrequestdata_Encode
  (
    Reqdata_In Clob,
    Appid_In   Varchar2,
    Key_In     Varchar2,
    Version_In Varchar2 := '1.0'
  ) Return Clob;

  Function Getrequestdata_Decode
  (
    Datasouce_In   Clob,
    Appid_In       Varchar2,
    Key_In         Varchar2,
    Datadecode_Out Out Clob,
    Errmsg_Out     Out Varchar2,
    Version_In     Varchar2 := '1.0'
  ) Return Number;

  Procedure Get_Parainfo
  (
    Version_Out     Out Varchar2,
    Url_Out         Out Varchar2,
    Url_Type_Out    Out Varchar2,
    Appid_Out       Out Varchar2,
    Key_Out         Out Varchar2,
    Contenttype_Out Out Varchar2,
    Charset_Out     Out Varchar2
  );
  Function Get_Version Return Varchar2;

  Function Request
  (
    Reqdata_In    Clob,
    Servername_In Varchar2,
    Respdata_Out  Out Clob,
    Errmsg_Out    Out Varchar2,
    Version_In    Varchar2 := '1.0'
  ) Return Number;

  Function Get_Einvoice_Node
  (
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ��Աid_In     ��Ա��.Id%Type := Null
    
  ) Return Varchar2;

  Function Einvoice_Start
  (
    ҵ�񳡾�_In Integer,
    ����_In     ���ս����¼.����%Type,
    ����_In     Integer := Null
  ) Return Number;

  Procedure Get_Chargedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Procedure Get_Sendcarddata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Procedure Get_Registerdata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Procedure Get_Zybalancedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Procedure Get_Mzbalancedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Procedure Get_Depositdata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  );

  Function Einvoice_Create
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number;

  --����Ʊ�����ϼ��
  Function Einvoice_Cancel_Check
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number;

  --����Ʊ������
  Function Einvoice_Cancel
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number;

End b_Einvoice_Request_Test;
/


Create Or Replace Package Body b_Einvoice_Request_Test Is
  ------------------------------------------------------------------
  --����Ʊ��ҵ����(��Ҫ����Բ�˼����Ʊ�ݵĴ���)
  --һ.��������˵��
  --  0.get_charset-��ȡ��ǰ���ݿ���ַ���
  --  1.zlJsonStr-����Ҫ��ϳ�Json�����ַ����е��������
  --  2.MD5_Clob-���Clob����MD5����
  --  3.Request-����ҵ������   
  --  3.GetRequestData_Encode-��ȡ��������(����)
  --  4.Getrequestdata_Decode-��ȡҵ������(����)
  --  5.Get_ParaInfo-��ȡ��˼��ز�������
  --  6.Get_Version-��ȡ��ǰ��˼�İ汾��
  --  7 Get_IDENTITY-��ȡ����Ա��Ϣ
  --��.����Ʊ�ݴ���
  --1.��ȡҵ���������
  --1.1 Get_Einvoice_Node-��ȡ��Ʊ�㺯��
  --1.2 Get_Chargedata_Create-��ȡ�շѿ�ƱƱ��
  --1.3 Get_SendCarddata_Create-��ȡ������ƱƱ��
  --1.4 Get_Registerdata_Create-��ȡ�Һſ�ƱƱ��
  --1.5 Get_MZBalancedata_Create-��ȡ������ʿ�ƱƱ��
  --1.6 Get_ZYBalancedata_Create-��ȡסԺ���ʿ�ƱƱ��
  --1.7 Get_Depositdata_Create-��ȡԤ����ƱƱ��
  --2.����Ʊ�ݲ������
  --2.1.Einvoice_Start-����Ʊ���Ƿ�����
  --2.2.EInvoice_Create-����Ʊ�ݿ���
  --2.3.Einvoice_Cancel_Check-����Ʊ������ǰ���(����:1-�Ϸ�;0-���Ϸ�)
  --2.4.Einvoice_Cancel-����Ʊ������(����1-�ɹ�;0-ʧ��)
  ------------------------------------------------------------------
  Mv_��Ʊ��     ����Ʊ��ʹ�ü�¼.��Ʊ��%Type;
  Mv_��Աid     ��Ա��.Id%Type;
  Mv_����Ա��� ��Ա��.���%Type;
  Mv_����Ա���� ��Ա��.����%Type;
  Mv_Url_Type   �����ӿ�����.����ֵ%Type;

  Mv_Url         �����ӿ�����.����ֵ%Type;
  Mv_Appid       �����ӿ�����.����ֵ%Type;
  Mv_Key         �����ӿ�����.����ֵ%Type;
  Mv_Version     �����ӿ�����.����ֵ%Type;
  Mv_Contenttype �����ӿ�����.����ֵ%Type;
  Mv_Charset     �����ӿ�����.����ֵ%Type;
  Mn_Charset     Number(2); --1.ZHS16GBK;0-AL32UTF8

  Function Get_Charset Return Number As
    --���ܣ���ȡ��ǰ���ݿ���ַ���
    --����:1.ZHS16GBK;0-AL32UTF8
    n_Charset Number(2);
  Begin
    If Mn_Charset Is Not Null Then
      Return Mn_Charset;
    End If;
    Begin
      Select Nvl(Max(1), 0)
      Into n_Charset
      From Nls_Database_Parameters
      Where Parameter = 'NLS_CHARACTERSET' And Value Like '%ZHS16GBK%';
    Exception
      When Others Then
        n_Charset := 1;
    End;
    Mn_Charset := n_Charset;
  End Get_Charset;

  Function zlJsonStr
  (
    Str_In  In Varchar2,
    Type_In Number := 0
  ) Return Varchar2 Is
    --���ܣ�����Ҫ��ϳ�Json�����ַ����е��������
    --������
    --     Type_In=��������:0-�ַ�,1-��ֵ
    v_Temp Varchar2(32767);
  Begin
    If Str_In Is Null Then
      If Nvl(Type_In, 0) = 1 Then
        Return '0';
      Else
        Return Null;
      End If;
    Elsif Nvl(Type_In, 0) = 1 Then
      If Instr(Str_In, '.') = 0 Then
        Return Str_In;
      Else
        Return To_Char(Str_In, 'FM9999999990.0999999999');
      End If;
    Else
      --ע��˳��
      v_Temp := Str_In;
      v_Temp := Replace(v_Temp, '\', '\\');
      v_Temp := Replace(v_Temp, '"', '\"');
      v_Temp := Replace(v_Temp, Chr(13), '\r');
      v_Temp := Replace(v_Temp, Chr(10), '\n');
      v_Temp := Replace(v_Temp, Chr(9), '\t');
      Return v_Temp;
    End If;
  End zlJsonStr;
  Procedure Get_Identity
  (
    ��Աid_Out     Out ��Ա��.Id%Type,
    ��Ա���_Out   Out ��Ա��.���%Type,
    ����Ա����_Out Out ��Ա��.����%Type
  ) Is
  Begin
    If Mv_����Ա��� Is Not Null Then
      ��Աid_Out     := Mv_��Աid;
      ��Ա���_Out   := Mv_����Ա���;
      ����Ա����_Out := Mv_����Ա����;
      Return;
    End If;
    Select p.Id, p.���, p.����
    Into ��Աid_Out, ��Ա���_Out, ����Ա����_Out
    From �ϻ���Ա�� O, ��Ա�� P
    Where o.�û��� = User And o.��Աid = p.Id And Rownum < 2;
  
    Mv_��Աid     := ��Աid_Out;
    Mv_����Ա��� := ��Ա���_Out;
    Mv_����Ա���� := ����Ա����_Out;
  Exception
    When Others Then
      Null;
  End Get_Identity;

  Function Md5_Clob(Souce_In In Clob) Return Varchar2 Is
  Begin
    If Souce_In Is Null Then
      Return Null;
    End If;
    Return Rawtohex(Dbms_Crypto.Hash(Souce_In, Dbms_Crypto.Hash_Md5));
  End Md5_Clob;

  Function Getrequestdata_Encode
  (
    Reqdata_In Clob,
    Appid_In   Varchar2,
    Key_In     Varchar2,
    Version_In Varchar2 := '1.0'
  ) Return Clob As
    ------------------------------------------------------------------
    --ҵ�����ݼ��ܴ���
    --���:
    --   ReqData_In-���������
    --   Appid_In-Ӧ���ʺ�
    --   Key_In-ǩ��˽Կ
    --   version_In-����İ汾�ţ�ȱʡΪ:1.0
    --����:
    --    ���ܺ������
    ------------------------------------------------------------------
  
    v_Guid    Varchar2(1000);
    v_Sign    Varchar2(32676);
    c_Data    Clob;
    c_Temp    Clob;
    n_Charset Number(2);
  Begin
  
    v_Guid := Sys_Guid();
  
    n_Charset := Get_Charset;
  
    --base64����
    c_Data := Replace(Zltools.Zlbase64.Decode(Reqdata_In, n_Charset), Chr(13), '');
  
    --stringA=appid+app1+data=ewogI&noise=ibuaiVcKdpRxkhJA+key=192006250b4c09247ec02f6a2d+version=1.0
    c_Temp := To_Clob('appid=' || Appid_In || 'data=') || c_Data ||
              To_Clob('noise=' || v_Guid || 'key=' || Key_In || 'version=' || Version_In);
  
    --MD5����
    v_Sign := Upper(Md5_Clob(c_Temp));
  
    --ҵ���������
    c_Temp := To_Clob('{');
    c_Temp := c_Temp || To_Clob('"appid":"' || Appid_In || '",'); --Ӧ���ʺ�
    c_Temp := c_Temp || To_Clob('"data":"' || c_Data || '",'); --����ҵ�����
    c_Temp := c_Temp || To_Clob('"noise":"' || v_Guid || '",'); --���������ʶ
    c_Temp := c_Temp || To_Clob('"version":"' || Version_In || '",'); --�汾
    c_Temp := c_Temp || To_Clob('"sign":"' || v_Sign || '"'); --ǩ��
    c_Temp := c_Temp || To_Clob('}');
  
    Return c_Temp;
  End Getrequestdata_Encode;

  Function Getrequestdata_Decode
  (
    Datasouce_In   Clob,
    Appid_In       Varchar2,
    Key_In         Varchar2,
    Datadecode_Out Out Clob,
    Errmsg_Out     Out Varchar2,
    Version_In     Varchar2 := '1.0'
  ) Return Number As
    ------------------------------------------------------------------
    --����ҵ������(����)
    --���:
    --   dataSouce_In-��Ҫ�����ԭ����
  
    --����:
    --    DataDecode_Out-���ؽ���������
    --    ErrMsg_out-���صĴ�������
    ------------------------------------------------------------------
    j_Json PLJson;
    c_Data       Clob;
    v_Temp       Varchar2(32767);
    c_Temp       Clob;
    v_Result     Varchar2(1000);
    n_Charset    Number(2);
    v_Guid       Varchar2(32767);
    v_Sign       Varchar2(100);
    v_Sign_Check Varchar2(100);
  Begin
    j_Json := PLJson(Datasouce_In);
  
    --  ���ؽ������  data  String  ��  ��API���÷��ص����ݲ�ͬ��ʵ�������Ը��ӿ�APIΪ׼������ֵΪJSON��ʽ��base64���룬�����ַ���UTF-8 
    --  ���������ʶ  noise  String  ��  ÿ�����󷵻�һ��Ψһ��ţ�ȫ��Ψһ���������UUID����
    --  ǩ��  sign  String  ��  MD5ժҪ���ת���ɴ�д
    c_Data := j_Json.Get_Clob('data');
    v_Guid := j_Json.Get_String('noise');
    v_Sign := j_Json.Get_String('sign');
  
    --stringA=appid+app1+data=ewogI&noise=ibuaiVcKdpRxkhJA+key=192006250b4c09247ec02f6a2d+version=1.0
    c_Temp := To_Clob('appid=' || Appid_In || 'data=') || c_Data ||
              To_Clob('noise=' || v_Guid || 'key=' || Key_In || 'version=' || Version_In);
    --MD5����
    v_Sign_Check := Upper(b_Einvoice_Request_Test.Md5_Clob(c_Temp));
    If Nvl(v_Sign_Check, '') <> Nvl(v_Sign, '') Then
      --��һ�£����ܱ����ģ�����0
      Errmsg_Out := 'ǩ����Ϣ����ȷ�������Ʒ�ṩ����ϵ��';
      Return 0;
    End If;
  
    n_Charset := b_Einvoice_Request_Test.Get_Charset;
    c_Data    := Zltools.Zlbase64.Decode(c_Data, n_Charset); --base64����
  
    j_Json   := PLJson(c_Data);
    v_Result := j_Json.Get_String('result'); --��S0000��ʾ�ɹ���ʶ��������Ϊ�����ʶ
    If Nvl(v_Result, '') <> 'S0000' Then
      --{"result":"E0001","message":"BASE64(������Ϣ)"}��
      v_Temp         := j_Json.Get_String('message');
      v_Temp         := Zltools.Zlbase64.Decode(v_Temp, n_Charset);
      Errmsg_Out     := v_Result || '-' || v_Temp;
      Datadecode_Out := Null;
      Return 0;
    End If;
  
    --����ɹ�:
    --{"result":"S0000","message":"BASE64(��Ӧҵ�����)"}
    Datadecode_Out := j_Json.Get_Clob('message');
    --base64����
    Datadecode_Out := Zltools.Zlbase64.Decode(Datadecode_Out, n_Charset);
    Return 1;
  Exception
    When Others Then
      Errmsg_Out := 'JSON���ݴ���' || To_Char(SQLCode) || ':' || SQLErrM;
      Return 0;
  End Getrequestdata_Decode;

  Procedure Get_Parainfo
  (
    Version_Out     Out Varchar2,
    Url_Out         Out Varchar2,
    Url_Type_Out    Out Varchar2,
    Appid_Out       Out Varchar2,
    Key_Out         Out Varchar2,
    Contenttype_Out Out Varchar2,
    Charset_Out     Out Varchar2
  ) Is
  Begin
    If Mv_Url Is Null Then
    
      For c_���� In (Select ������, ������, ����ֵ From �����ӿ����� Where �ӿ��� = '����ģ��ӿ�(V1.0.0)') Loop
        If c_����.������ = Upper('URL_Type') Then
          Mv_Url_Type := c_����.����ֵ;
        Elsif c_����.������ = Upper('URL_Address') Then
          Mv_Url := c_����.����ֵ;
        Elsif c_����.������ = Upper('Ӧ���ʺ�') Then
          Mv_Appid := c_����.����ֵ;
        Elsif c_����.������ = Upper('ǩ��˽Կ') Then
          Mv_Key := c_����.����ֵ;
        Elsif c_����.������ = Upper('֧�ְ汾') Then
          Mv_Version := c_����.����ֵ;
        Elsif c_����.������ = Upper('���ݴ��䷽ʽ') Then
          Mv_Contenttype := c_����.����ֵ;
        Elsif c_����.������ = Upper('�ַ�����') Then
          Mv_Charset := c_����.����ֵ;
        End If;
      End Loop;
    
      If Mv_Charset Is Null Then
        Mv_Charset := 'UTF-8';
      End If;
      If Mv_Url_Type Is Null Then
        Mv_Url_Type := 'HTTP';
      End If;
      Mv_Url := Mv_Url_Type || '://' || Mv_Url;
      If Mv_Contenttype Is Null Then
        Mv_Contenttype := 'application/json';
      End If;
    End If;
    Version_Out     := Mv_Version;
    Url_Out         := Mv_Url;
    Url_Type_Out    := Mv_Url_Type;
    Appid_Out       := Mv_Appid;
    Key_Out         := Mv_Key;
    Contenttype_Out := Mv_Contenttype;
    Charset_Out     := Mv_Charset;
  End Get_Parainfo;

  Function Get_Version Return Varchar2 As
    v_Version     �����ӿ�����.����ֵ%Type;
    v_Url         �����ӿ�����.����ֵ%Type;
    v_Url_Type    �����ӿ�����.����ֵ%Type;
    v_Appid       �����ӿ�����.����ֵ%Type;
    v_Key         �����ӿ�����.����ֵ%Type;
    v_Contenttype �����ӿ�����.����ֵ%Type;
    v_Charset     �����ӿ�����.����ֵ%Type;
  Begin
    Get_Parainfo(v_Version, v_Url, v_Url_Type, v_Appid, v_Key, v_Contenttype, v_Charset);
    Return v_Version;
  End Get_Version;

  Function Request
  (
    Reqdata_In    Clob,
    Servername_In Varchar2,
    Respdata_Out  Out Clob,
    Errmsg_Out    Out Varchar2,
    Version_In    Varchar2 := '1.0'
  ) Return Number As
    ------------------------------------------------------------------
    --����ҵ������
    --���:
    --   ReqData_In-���͵ı�������
    --   ServerName_in-���������
    --   version_In-����İ汾�ţ�ȱʡΪ:1.0
    --����:
    --    RespData_Out-��Ӧ��������
    --    ErrMsg_out-���صĴ�������
    ------------------------------------------------------------------
  
    o_Http_Req  Utl_Http.Req; --http�������
    o_Http_Resp Utl_Http.Resp; --http��Ӧ����
  
    Err_Item Exception;
    v_Err_Msg Varchar2(255);
    c_Temp    Clob; --��Ӧ����
  
    v_Buffer_Text Varchar2(32767); --����
    n_Deftimeout  Integer Default 3600;
    --HttP��ر���
    v_Url         Varchar2(4000); --���͵�Http
    v_Url_Type    �����ӿ�����.����ֵ%Type;
    v_Appid       �����ӿ�����.����ֵ%Type;
    v_Key         �����ӿ�����.����ֵ%Type;
    v_Version     �����ӿ�����.����ֵ%Type;
    v_Charset     �����ӿ�����.����ֵ%Type;
    v_Contenttype Varchar2(1000);
    v_Temp        Varchar2(32767);
    n_Amount      Pls_Integer := 3900;
    n_Offset      Pls_Integer := 1;
    n_Count       Number(18);
    c_Data        Clob;
  Begin
  
    Get_Parainfo(v_Version, v_Url, v_Url_Type, v_Appid, v_Key, v_Contenttype, v_Charset);
    --URL��http://[ip]:[port]/[service]/api/medical/ [�ӿڷ����ʶ]
    v_Url     := v_Url || '/' || Servername_In;
    v_Version := Version_In;
  
    --��֯ҵ������ļ�������
    c_Data := Getrequestdata_Encode(Reqdata_In, v_Appid, v_Key, v_Version);
  
    -- ��ʼ��HTTP�������.
    Utl_Http.Set_Transfer_Timeout(n_Deftimeout);
    o_Http_Req := Utl_Http.Begin_Request(v_Url, 'POST');
    Utl_Http.Set_Header(o_Http_Req, 'Content-Type', v_Contenttype);
    Utl_Http.Set_Header(o_Http_Req, 'Content-Length', Lengthb(c_Data));
    Utl_Http.Set_Body_Charset(o_Http_Req, v_Charset);
  
    n_Count := Dbms_Lob.Getlength(c_Data);
  
    If n_Count > 30000 Then
      --�ֿ鷢��HTTP����
      Utl_Http.Set_Header(o_Http_Req, 'Transfer-Encoding', 'chunked'); --�ֿ�
      While (n_Offset < n_Count) Loop
        Dbms_Lob.Read(c_Data, n_Amount, n_Offset, v_Buffer_Text);
      
        Utl_Http.Write_Text(o_Http_Req, v_Buffer_Text);
        n_Offset := n_Offset + n_Amount;
      End Loop;
    Else
      --����HTTP����
      v_Temp := c_Data;
      Utl_Http.Write_Text(o_Http_Req, v_Temp);
    End If;
    --������Ӧ
    o_Http_Resp := Utl_Http.Get_Response(o_Http_Req);
  
    Begin
      c_Temp := Null;
      v_Temp := Null;
      Loop
        Utl_Http.Read_Text(o_Http_Resp, v_Buffer_Text, 30000);
        If Length(Nvl(v_Temp, '') || v_Buffer_Text) >= 30000 Then
          c_Temp := c_Temp || To_Clob(Nvl(v_Temp, '') || v_Buffer_Text);
          v_Temp := Null;
        End If;
        v_Temp := Nvl(v_Temp, '') || v_Buffer_Text;
      End Loop;
      If v_Temp Is Not Null And c_Temp Is Not Null Then
        c_Temp := c_Temp || To_Clob(Nvl(v_Temp, ''));
        v_Temp := Null;
      End If;
      --�ر�HTTP����
      Utl_Http.End_Response(o_Http_Resp);
    Exception
      When Utl_Http.Request_Failed Then
        v_Err_Msg := 'HTTP����ʧ�ܣ�' || To_Char(SQLCode) || ':' || Substr(SQLErrM, 1, 128);
        Raise Err_Item;
      When Utl_Http.Transfer_Timeout Then
        v_Err_Msg := 'HTTP����ʱʧ�ܣ�' || To_Char(SQLCode) || ':' || Substr(SQLErrM, 1, 128);
        Raise Err_Item;
      When Utl_Http.End_Of_Body Then
        Utl_Http.End_Response(o_Http_Resp);
      When Others Then
        Errmsg_Out := 'HTTP�������' || To_Char(SQLCode) || ':' || Substr(SQLErrM, 1, 128);
        Raise Err_Item;
    End;
  
    --��������
    If v_Temp Is Null Then
      If Getrequestdata_Decode(c_Temp, v_Appid, v_Key, Respdata_Out, Errmsg_Out, v_Version) = 1 Then
        Return 1;
      End If;
    Else
      If Getrequestdata_Decode(v_Temp, v_Appid, v_Key, Respdata_Out, Errmsg_Out, v_Version) = 1 Then
        Return 1;
      End If;
    End If;
    Return 0;
    --�ͷ�clob
    Dbms_Lob.Freetemporary(c_Temp);
    Return 1;
  Exception
    When Err_Item Then
      Utl_Http.End_Response(o_Http_Resp);
      Dbms_Lob.Freetemporary(c_Temp);
      Errmsg_Out := v_Err_Msg;
      Return 0;
    When Others Then
      Utl_Http.End_Response(o_Http_Resp);
      Dbms_Lob.Freetemporary(c_Temp);
      Errmsg_Out := 'JSON���ݴ���' || To_Char(SQLCode) || ':' || Substr(SQLErrM, 1, 128);
      Return 0;
  End Request;

  Function Get_Einvoice_Node
  (
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ��Աid_In     ��Ա��.Id%Type := Null
  ) Return Varchar2 Is
    n_Count        Number(18);
    v_������       Varchar2(100);
    v_����Ա��Ʊ�� ����Ʊ�ݿ�Ʊ��.����%Type;
    v_��������Ʊ�� ����Ʊ�ݿ�Ʊ��.����%Type;
    v_����Ա���   ��Ա��.���%Type;
    v_����Ա����   ��Ա��.����%Type;
    n_��Աid       ��Ա��.Id%Type;
  Begin
    --ȱʡ���շ�Ա���Ϊ����
    If Mv_��Ʊ�� Is Not Null Then
      Return Mv_��Ʊ��;
    End If;
  
    Select Max(1) Into n_Count From Ʊ�ݿ�Ʊ����� Where Rownum < 2;
    If Nvl(n_Count, 0) = 0 Then
      If ����Ա���_In Is Null Then
        If Nvl(��Աid_In, 0) <> 0 Then
          Select Max(���) Into v_����Ա��� From ��Ա�� Where ID = ��Աid_In;
          If v_����Ա��� Is Not Null Then
            Return v_����Ա���;
          End If;
        End If;
      
        If ����Ա����_In Is Not Null Then
          Select Max(���) Into v_����Ա��� From ��Ա�� Where ���� = ����Ա����_In;
          If v_����Ա��� Is Not Null Then
            Return v_����Ա���;
          End If;
        End If;
        --�Ե�½��Ϊ׼
        If Mv_����Ա��� Is Not Null Then
          Return Mv_����Ա���;
        End If;
      
        Get_Identity(n_��Աid, v_����Ա���, v_����Ա����);
        If v_����Ա��� Is Not Null Then
          Return v_����Ա���;
        End If;
      
      End If;
      Return ����Ա���_In;
    End If;
    n_��Աid := ��Աid_In;
    If ��Աid_In = 0 Then
      If ����Ա���_In Is Not Null Or ����Ա����_In Is Not Null Then
        Select Max(ID)
        Into n_��Աid
        From ��Ա��
        Where ��� = Nvl(����Ա���_In, 's-') Or ���� = Nvl(����Ա����_In, '--');
      End If;
    End If;
  
    If Nvl(n_��Աid, 0) = 0 Then
      Get_Identity(n_��Աid, v_����Ա���, v_����Ա����);
    End If;
    Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  
    For c_��Ʊ�� In (Select Nvl(a.��Աid, 0) As ��Աid, Nvl(a.�ͻ���, '-') As �ͻ���, a.��Ʊ��id, b.���� As ��Ʊ�����, b.����
                  From Ʊ�ݿ�Ʊ����� A, ����Ʊ�ݿ�Ʊ�� B
                  Where a.��Ʊ��id = b.Id And Nvl(b.����ʱ��, Sysdate + 1) >= Sysdate And (a.��Աid = n_��Աid Or a.�ͻ��� = v_������)
                  Order By ��Աid, �ͻ���) Loop
      If Nvl(c_��Ʊ��.��Աid, 0) = 0 And c_��Ʊ��.�ͻ��� = v_������ Then
        v_��������Ʊ�� := c_��Ʊ��.��Ʊ�����;
      End If;
    
      If Nvl(c_��Ʊ��.��Աid, 0) = n_��Աid And c_��Ʊ��.�ͻ��� = '-' Then
        v_����Ա��Ʊ�� := c_��Ʊ��.��Ʊ�����;
      End If;
    
      If Nvl(c_��Ʊ��.��Աid, 0) = n_��Աid And c_��Ʊ��.�ͻ��� = v_������ Then
        Return c_��Ʊ��.��Ʊ�����;
      End If;
    End Loop;
    If v_����Ա��Ʊ�� Is Not Null Then
      Return v_����Ա��Ʊ��;
    End If;
    Return v_��������Ʊ��;
  End Get_Einvoice_Node;

  Function Einvoice_Start
  (
    ҵ�񳡾�_In Integer,
    ����_In     ���ս����¼.����%Type,
    ����_In     Integer := Null
  ) Return Number Is
    ------------------------------------------------------------------
    --����:�жϵ���Ʊ���Ƿ�����
    --���:ҵ�񳡾�_In-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --     ����_In��null-��ʾ������;���ʺ�Ԥ��:1-����;2-סԺ 
    --����:������Ϣ_Out-���صĴ�����Ϣ
    --����:1-����;0-δ����
    -------------------------------------------------------------------
    n_Return Number(2);
  Begin
    n_Return := Zl_Fun_Isstarteinvoice(ҵ�񳡾�_In, ����_In, ����_In);
    Return n_Return;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Start;

  Procedure Get_Chargedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    --
    ---------------------------------------------------------------------------
    --����:��ȡ�շѿ�Ʊ����
    --���:
    --    Json_In,��ʽ����
    --  input
    --    occasion N 1  Ӧ�ó���:1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --    einvoice_id  N,1 ��ǰ����Ʊ��ID
    --    balance_id N 1  ����ID  occasion=2ʱ��Ԥ��id;occasion<>2��ʾ����id
    --    writeoff_id  N 1  ����ID  occasion=2ʱ������Ԥ��id;occasion<>2��ʾ����id
    --����:
    --  ReqData_Out-���ص�ҵ����������
    --  Totalmoney_Out-Ʊ���ܶ�
    --  Code_Out-��ȡ�Ƿ�ɹ���0-ʧ�ܣ�1-�ɹ�
    --  Message_Out ������Ϣ
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_Ӧ�ó��� Number(2);
    n_����id   ����Ԥ����¼.����id%Type;
    --n_����id   ����Ԥ����¼.����id%Type;
  
    v_ҵ���ʶ   Varchar2(20);
    v_ҵ����ˮ�� Varchar2(50);
  
    v_��Ʊ��       Varchar2(100);
    v_�ɷ�         Varchar2(32767);
    v_Ʊ����Ϣ     Varchar2(32767);
    v_������Ϣ     Varchar2(32767);
    v_֪ͨ         Varchar2(32767);
    v_�ɷ�����     Varchar2(32767);
    v_����         Varchar2(32767);
    v_������չ��Ϣ Varchar2(32767);
    v_����ҽ����Ϣ Varchar2(32767);
    c_��ϸ         Clob;
    v_��ϸ         Varchar2(32767);
    c_������ϸ     Clob;
    v_������ϸ     Varchar2(32767);
    c_������Ϣ     Clob; --���շ��صĽ�����Ϣ��
  
    n_�����       ������Ϣ.�����%Type;
    n_����id       ����Ԥ����¼.����id%Type;
    v_��������     ������ü�¼.����%Type;
    v_�����Ա�     ������ü�¼.�Ա�%Type;
    v_��������     ������ü�¼.����%Type;
    d_ҵ����ʱ�� ������ü�¼.�Ǽ�ʱ��%Type;
    v_�շ�Ա       ������ü�¼.����Ա����%Type;
  
    n_ȱʡ�����id     Number(18);
    v_����ֵ           Varchar2(100);
    n_Ʊ���ܽ��       ������ü�¼.���ʽ��%Type;
    n_����ܶ�         ������ü�¼.���ʽ��%Type;
    n_�û�id           ��Ա��.Id%Type;
    v_����Ա���       ��Ա��.���%Type;
    v_����Ա����       ��Ա��.����%Type;
    v_Temp             Varchar2(32767);
    v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
    v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
    n_����             ���ս����¼.����%Type;
    v_���ջ�������     �������.���ջ�������%Type;
    n_ҽ�����         ������ü�¼.ҽ�����%Type;
    n_�Һ�id           ������ü�¼.�Һ�id%Type;
    v_��������         ���ղ���.����%Type;
    v_��������         Varchar2(20);
    v_������ұ���     ���ű�.����%Type;
    v_�����������     ���ű�.����%Type;
    v_������         Varchar2(50);
    v_����ids          Varchar2(32767);
    v_ҽ����           �����ʻ�.ҽ����%Type;
    l_����id           t_NumList := t_NumList();
    v_�汾��           Varchar2(30);
    n_����Ʊ��id       ����Ʊ��ʹ�ü�¼.Id%Type;
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_Ӧ�ó��� := Nvl(j_Json.Get_Number('occasion'), 0);
    n_����id   := j_Json.Get_Number('balance_id');
    --n_����id     := Nvl(j_Json.Get_Number('writeoff_id'), 0);
    n_����Ʊ��id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
    If Nvl(n_Ӧ�ó���, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '��Ч��Ӧ�ó���';
      Return;
    End If;
  
    Select Nvl(Max(����ֵ), 'V2.0.3')
    Into v_�汾��
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '֧�ְ汾';
  
    b_Einvoice_Request_Test.Get_Identity(n_�û�id, v_����Ա���, v_����Ա����);
  
    --v_ҵ���ʶ:01  סԺ,02  ����,03  ����,04  ����,05  �������,06  �Һ�,07  סԺԤ����,08  ���Ԥ����
    Select Decode(n_Ӧ�ó���, 1, '02', 2, '07', 3, '01', 4, '06', 5, '02', '02') Into v_ҵ���ʶ From Dual;
  
    n_Ʊ���ܽ��   := 0;
    d_ҵ����ʱ�� := Null;
    v_����ids      := Null;
    c_��ϸ         := Null;
    v_��ϸ         := Null;
    For c_�շ�ϸĿ In (Select Min(a.Id) As ����id, a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, Max(a.���㵥λ) As ���㵥λ,
                          Sum(a.��׼����) As �۸�, Avg(Nvl(a.����, 1) * Nvl(a.����, 0)) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��,
                          Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.ʵ�ս��) - Sum(a.ͳ����) As �Էѽ��,
                          Max(t.����) As ҽ����Ŀ����, Max(t.����) As ҽ����Ŀ����, Max(t.ͳ��ȶ�) As ҽ����������, Max(a.ժҪ) As ��ע,
                          Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����) As ����,
                          Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��,
                          Max(a.���ʽ) As ���ʽ����, Max(a.�վݷ�Ŀ) As �վݷ�Ŀ, Max(c.����) As �վݷ�Ŀ����, Max(a.ҽ�����) As ҽ�����,
                          Max(a.�Һ�id) As �Һ�id, Max(d.����) As ������, Max(d.���) As �������, Max(b.����) As ��Ŀ����, Max(b.����) As ��Ŀ����,
                          Max(b.���) As ���, Max(q.ҩƷ����) As ҩƷ����
                   From ������ü�¼ A, �շ���ĿĿ¼ B, �վݷ�Ŀ C, �շ���� D, ҩƷ��� M, ҩƷ���� Q, ������ĿĿ¼ J, ����֧������ T, ֧�������� S
                   Where a.No In (Select Distinct NO From ������ü�¼ Where ����id = n_����id) And Mod(a.��¼����, 10) = 1 And
                         a.�շ���� = d.����(+) And a.�շ�ϸĿid = b.Id And a.�վݷ�Ŀ = c.����(+) And a.�շ�ϸĿid = m.ҩƷid(+) And
                         m.ҩ��id = q.ҩ��id(+) And q.ҩ��id = j.Id(+) And a.���մ���id = t.Id(+) And t.����(+) = 1 And
                         a.���մ���id = s.���մ���id(+)
                   Group By a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, c.����, c.����, j.����, j.����
                   Order By NO, ���) Loop
      If v_�������� Is Null Then
        v_�������� := c_�շ�ϸĿ.����;
        v_�����Ա� := c_�շ�ϸĿ.�Ա�;
        v_�������� := c_�շ�ϸĿ.����;
        n_����id   := c_�շ�ϸĿ.����id;
      End If;
      If d_ҵ����ʱ�� Is Null And Nvl(c_�շ�ϸĿ.��¼״̬, 0) = 1 Then
        --ȡԭʼҵ����ʱ��
        d_ҵ����ʱ�� := c_�շ�ϸĿ.�Ǽ�ʱ��;
        v_�շ�Ա       := c_�շ�ϸĿ.����Ա����;
      End If;
      If v_ҽ�Ƹ��ʽ���� Is Null Then
        v_ҽ�Ƹ��ʽ���� := c_�շ�ϸĿ.���ʽ����;
      End If;
      If Nvl(n_ҽ�����, 0) = 0 Then
        n_ҽ����� := c_�շ�ϸĿ.ҽ�����;
      End If;
      If Nvl(n_�Һ�id, 0) = 0 Then
        n_�Һ�id := c_�շ�ϸĿ.�Һ�id;
      End If;
    
      If Instr(Nvl(v_����ids, '') || ',', ',' || c_�շ�ϸĿ.����id || ',') = 0 Then
        l_����id.Extend;
        l_����id(l_����id.Count) := c_�շ�ϸĿ.����id;
      End If;
    
      --listDetailNo  ��ϸ��ˮ��  String  60  ��  ��ϸ��ˮ��
      v_Temp := '{' || '"listDetailNo":"' || b_Einvoice_Request_Test.Zljsonstr(LPad(c_�շ�ϸĿ.����id, 20, '0')) || '"';
      --chargeCode  �շ���Ŀ����  String  50  ��  ��дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���,�磺��λ�ѡ�����
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�վݷ�Ŀ����) || '"';
      --chargeName  �շ���Ŀ����  String  100  ��
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�վݷ�Ŀ) || '"';
      --prescribeCode  ��������  String  60  ��
      v_Temp := v_Temp || ',' || '"prescribeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.No) || '"';
      --listTypeCode  ҩƷ������  String  50  ��  ��ҩƷ�������01��������д
      v_Temp := v_Temp || ',' || '"listTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.������) || '"';
      --listTypeName  ҩƷ�������  String  50  ��  ��ҩƷ�������ƣ��������࿹��Ⱦҩ��
      v_Temp := v_Temp || ',' || '"listTypeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�������) || '"';
      --code  ����  String  50  ��  ��ҩƷ���룬������д
      v_Temp := v_Temp || ',' || '"code":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��Ŀ����) || '"';
      --name  ҩƷ����  String  50  ��  ��ҩƷ���ƣ��������Ƶ�
      v_Temp := v_Temp || ',' || '"name":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��Ŀ����) || '"';
      --form  ����  String  50  ��
      v_Temp := v_Temp || ',' || '"form":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҩƷ����) || '"';
      --specification  ���  String  50  ��
      v_Temp := v_Temp || ',' || '"specification":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���) || '"';
      --unit  ������λ   String  20  ��
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���㵥λ) || '"';
      --std  ����  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�۸�, 1);
      --number  ����  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.����, 1);
      --amt  ���  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ʵ�ս��, 1);
      --selfAmt  �Էѽ��  Number  14,6  ��  ���޽���д0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�Էѽ��, 1);
      --receivableAmt  Ӧ�շ���  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"receivableAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.Ӧ�ս��, 1);
      --medicalCareType  ҽ��ҩƷ����  String  1  ��  1�����Ը�/��
      --          2�����Ը�/��
      --          3��ȫ�Ը�/��
      v_Temp := v_Temp || ',' || '"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����Ŀ����) || '"';
      --medCareItemType  ҽ����Ŀ����  String  100  ��
      v_Temp := v_Temp || ',' || '"medCareItemType":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����Ŀ����) || '"';
      --medReimburseRate  ҽ����������  Number  3,2  ��
      v_Temp := v_Temp || ',' || '"medReimburseRate":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����������, 1);
      --remark  ��ע  String  200  ��
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��ע) || '"';
      --sortNo  ���  Integer  ����  ��  ���
      v_Temp := v_Temp || ',' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���, 1);
      --chrgtype  ��������  String  50  ��
      v_Temp := v_Temp || ',' || '"chrgtype":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��������) || '"}';
    
      If Length(Nvl(v_��ϸ, '') || v_Temp) > 32700 Then
        If c_��ϸ Is Null Then
          c_��ϸ := To_Clob(v_��ϸ);
        Else
          c_��ϸ := c_��ϸ || To_Clob(',' || v_��ϸ);
        End If;
        v_��ϸ := Null;
      End If;
    
      If v_��ϸ Is Null Then
        v_��ϸ := v_Temp;
      Else
        v_��ϸ := v_��ϸ || ',' || v_Temp;
      End If;
    End Loop;
  
    If v_��ϸ Is Not Null And c_��ϸ Is Not Null Then
      --listDetail  �嵥��Ŀ��ϸ  String  ����  ��  ���A-2,JSON��ʽ�б�
      c_��ϸ := c_��ϸ || ',' || To_Clob(v_��ϸ);
      c_��ϸ := To_Clob(',"listDetail":[') || c_��ϸ || To_Clob(']');
    
      v_��ϸ := Null;
    Elsif v_��ϸ Is Not Null Then
      v_��ϸ := ',"listDetail":[' || v_��ϸ || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --������ϸ
    v_������ϸ := Null;
    c_������ϸ := Null;
    For c_����ͳ�� In (Select Rownum As ���, �վݷ�Ŀ����, �վݷ�Ŀ����, ����, ���㵥λ, Round(����, 2) As ����, Round(���ʽ��, 2) As ���ʽ��,
                          Round(�Էѽ��, 2) As �Էѽ��, ��ע, ���ʽ�� - Round(���ʽ��, 2) As ����
                   From (Select /*+cardinality(b,10)*/
                           c.���� As �վݷ�Ŀ����, a.�վݷ�Ŀ As �վݷ�Ŀ����, 1 As ����, '' As ���㵥λ, Sum(a.���ʽ��) As ����, a.�վݷ�Ŀ,
                           Sum(a.���ʽ��) As ���ʽ��, Sum(a.���ʽ��) - Sum(a.ͳ����) As �Էѽ��, '' As ��ע
                          From ������ü�¼ A, Table(l_����id) B, �վݷ�Ŀ C
                          Where a.����id = b.Column_Value And a.�վݷ�Ŀ = c.����(+)
                          Group By c.����, a.�վݷ�Ŀ)) Loop
      --sortNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���շ���Ŀ���ֵ����1�����β������ظ�
      v_Temp := '{' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���, 1);
      --chargeCode  �շ���Ŀ����  String  50  ��  ��дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�վݷ�Ŀ����) || '"';
      --chargeName  �շ���Ŀ����  String  100  ��  ��дҵ��ϵͳ�ڲ���Ŀ����
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�վݷ�Ŀ����) || '"';
      --unit  ������λ  String  20  ��
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���㵥λ) || '"';
      --std  �շѱ�׼  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.����, 1);
      --number  ����  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.����, 1);
      --amt  ���  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���ʽ��, 1);
      --selfAmt  �Էѽ��  Number  14,2  ��  ���޽���д0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�Էѽ��, 1);
      --remark  ��ע  String  200  ��
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.��ע) || '"}';
    
      If Length(Nvl(v_������ϸ, '') || v_Temp) > 32700 Then
        If c_������ϸ Is Null Then
          c_������ϸ := To_Clob(v_������ϸ);
        Else
          c_������ϸ := c_������ϸ || To_Clob(',' || v_������ϸ);
        End If;
        v_������ϸ := Null;
      End If;
    
      If v_������ϸ Is Null Then
        v_������ϸ := v_Temp;
      Else
        v_������ϸ := v_������ϸ || ',' || v_Temp;
      End If;
    
      n_Ʊ���ܽ�� := Nvl(n_Ʊ���ܽ��, 0) + Nvl(c_����ͳ��.���ʽ��, 0);
      n_����ܶ�   := Nvl(n_����ܶ�, 0) + Nvl(c_����ͳ��.����, 0);
    End Loop;
    Totalmoney_Out := n_Ʊ���ܽ��;
    If v_������ϸ Is Not Null And c_������ϸ Is Not Null Then
      c_������ϸ := c_������ϸ || ',' || To_Clob(v_������ϸ);
      ----chargeDetail �շ���Ŀ��ϸ  String  ����  ��  ���A-1,JSON��ʽ�б�
      c_������ϸ := To_Clob(',"chargeDetail":[') || c_������ϸ || To_Clob(']');
      v_������ϸ := Null;
    Elsif v_������ϸ Is Not Null Then
      v_������ϸ := ',"chargeDetail":[' || v_������ϸ || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --Ʊ����Ϣ
    --Select Count(1) Into n_���ϴ��� From ����Ʊ��ʹ�ü�¼ Where Ʊ�� = n_Ӧ�ó��� And ����id = n_����id;
    --����ID||_||����Ʊ��ID
    v_ҵ����ˮ�� := n_����id || '_' || Nvl(n_����Ʊ��id, 0);
    v_��Ʊ��     := b_Einvoice_Request_Test.Get_Einvoice_Node(v_����Ա���, v_����Ա����, n_�û�id);
  
    v_Ʊ����Ϣ := '"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ����ˮ��) || '"'; --ҵ����ˮ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"busType":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ���ʶ) || '"'; --ҵ���ʶ
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --��������
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"busDateTime":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(d_ҵ����ʱ��, 'yyyymmddHH24miss') || '000') || '"'; --ҵ����ʱ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_��Ʊ��) || '"'; --��Ʊ�����:ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(v_�շ�Ա) || '"'; --�շ�Ա
  
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_����Ա����) || '"'; --Ʊ�ݱ�����
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"totalAmt":' || b_Einvoice_Request_Test.Zljsonstr(n_Ʊ���ܽ��, 1); --��Ʊ�ܽ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"remark":"' || '' || '"'; --��ע
    -------------------------------------------------------------------------------------------
  
    --ȡ�ɷ���Ϣ
    v_�ɷ� := Null;
    For c_�ɷ� In (Select Max(Decode(��Ϣ��, '������', ��Ϣֵ, '֧��������', ��Ϣֵ, '')) As ֧��������,
                        Max(Decode(��Ϣ��, 'ҽ��֧��������', ��Ϣֵ, 'ҽ��������', ��Ϣֵ, '')) As ҽ��֧��������,
                        Max(Decode(Upper(��Ϣ��), '֧�������ں�USERID', ��Ϣֵ, '')) As ֧�������ں�userid,
                        Max(Decode(Upper(��Ϣ��), '֧����С����USERID', ��Ϣֵ, '')) As ֧����С����userid,
                        Max(Decode(Upper(��Ϣ��), '΢�Ź��ں�OPENID', ��Ϣֵ, '')) As ΢�Ź��ں�openid,
                        Max(Decode(Upper(��Ϣ��), '΢��С����OPENID', ��Ϣֵ, '')) As ΢��С����openid
                 From (Select ��Ϣ��, ��Ϣֵ
                        From ������Ϣ�ӱ�
                        Where ����id = n_����id And ��Ϣ�� In ('֧�������ں�USERID', '֧����С����USERID', '΢�Ź��ں�OPENID', '΢��С����OPENID')
                        Union All
                        Select ������Ŀ, ��������
                        From �������㽻��
                        Where ����id In (Select ID From ����Ԥ����¼ Where ����id = n_����id) And ������Ŀ Like '%������')) Loop
      v_�ɷ� := ',"alipayCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.֧�������ں�userid) || '"'; --����֧�����˻�
      v_�ɷ� := v_�ɷ� || ',"weChatOrderNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.֧��������) || '"'; --΢��֧��������
      If v_�汾�� = 'V3.1.0' Then
        --�ð汾���д˽ӵ�
        v_�ɷ� := v_�ɷ� || ',"weChatMedTransNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.ҽ��֧��������) || '"'; --΢��ҽ��֧��������
      End If;
    
      If c_�ɷ�.΢�Ź��ں�openid Is Not Null Then
        v_�ɷ� := v_�ɷ� || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.΢�Ź��ں�openid) || '"'; --΢�Ź��ںŻ�С�����û�ID
      Else
        v_�ɷ� := v_�ɷ� || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.΢��С����openid) || '"'; --΢�Ź��ںŻ�С�����û�ID
      End If;
      Exit;
    End Loop;
  
    -------------------------------------------------------------------------------------------
    --ȡ֪ͨ��Ϣ
    Select To_Number(Max(����ֵ))
    Into n_ȱʡ�����id
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = 'ȱʡ�����ID';
    v_֪ͨ := Null;
    For c_֪ͨ In (Select Max(a.����id) As ����id, Max(a.����) As ����, Max(a.�ֻ���) As �ֻ���, Max(a.Email) As Email, Max(1) As �ɿ�����,
                        Max(a.���֤��) As ���֤��, Max(m.����) As �����, Max(m.����) As ����, Max(a.�����) As �����
                 From ������Ϣ A,
                      (
                        
                        Select ����id, ����, ����, ����
                        From (Select b.����id, c.����, c.����, b.����, Decode(b.�����id, n_ȱʡ�����id, 2, c.ȱʡ��־) As ȱʡ��־
                                From ����ҽ�ƿ���Ϣ B, ҽ�ƿ���� C
                                Where b.�����id = c.Id And b.����id = n_����id
                                Order By ȱʡ��־)
                        Where Rownum < 2) M
                 Where a.����id = m.����id(+)) Loop
    
      v_֪ͨ := ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.�ֻ���) || '"'; --�����ֻ�����
      v_֪ͨ := v_֪ͨ || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.Email) || '"'; --���������ַ
      If v_�汾�� = 'V3.1.0' Then
        --�ð汾���д˽ӵ�
        v_֪ͨ := v_֪ͨ || ',"payerType":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.�ɿ�����) || '"'; --����������
      End If;
      v_֪ͨ := v_֪ͨ || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.���֤��) || '"'; --ͳһ������ô���
    
      If c_֪ͨ.����� Is Not Null Then
        v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.�����) || '"'; --������
        v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.����) || '"'; --����
      Elsif c_֪ͨ.���֤�� Is Not Null Then
        Select Nvl(Max(����ֵ), '99998')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '���֤�������ͱ��';
        v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
        v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.���֤��) || '"'; --����
      Else
        --û��һ�ſ����̶�һ�ֿ����
        Select Nvl(Max(����ֵ), '99999')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ������';
        v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
        Select Nvl(Max(����ֵ), '-')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ���';
        v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --����
      End If;
      If Nvl(n_�����, 0) = 0 Then
        n_����� := c_֪ͨ.�����;
      
      End If;
    
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --������Ϣ
    Select Max(����) Into v_Temp From zlRegInfo Where ��Ŀ = 'ҽ�ƻ�������';
  
    --����:1-�շ�;2-���㣨����סԺ���㡢����������㣩��3-Ԥ��
    Select Max(a.����), Max(b.���ջ�������), Max(Nvl(a.��������, c.����))
    Into n_����, v_���ջ�������, v_��������
    From ���ս����¼ A, ������� B, ���ղ��� C
    Where a.���� = b.��� And a.����id = c.Id(+) And a.��¼id = n_����id And a.���� = Decode(n_Ӧ�ó���, 2, 3, 3, 2, 1);
  
    Select Max(����) Into v_ҽ�Ƹ��ʽ���� From ҽ�Ƹ��ʽ Where ���� = v_ҽ�Ƹ��ʽ����;
    If Nvl(n_����, 0) <> 0 Then
      Select Max(ҽ����) Into v_ҽ���� From �����ʻ� Where ����id = n_����id And ���� = n_����;
    End If;
  
    v_������ := Null;
    If Nvl(n_ҽ�����, 0) <> 0 Then
      Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')), Max(b.����), Max(b.����), Max(a.No)
      Into v_��������, v_������ұ���, v_�����������, v_������
      From ���˹Һż�¼ A, ���ű� B
      Where a.ִ�в���id = b.Id And
            a.No = (Select Max(�Һŵ�) From ����ҽ����¼ Where ID = n_ҽ����� Or ���id = n_ҽ�����);
    Elsif Nvl(n_�Һ�id, 0) <> 0 Then
      Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')), Max(b.����), Max(b.����), Max(a.No)
      Into v_��������, v_������ұ���, v_�����������, v_������
      From ���˹Һż�¼ A, ���ű� B
      Where a.ִ�в���id = b.Id And a.Id = n_�Һ�id;
    End If;
    If v_������ Is Null Then
      --ȡ���һ�ιҺ�ID
      Select Max(a.Id), Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')), Max(b.����), Max(b.����), Max(a.No)
      Into n_�Һ�id, v_��������, v_������ұ���, v_�����������, v_������
      From ���˹Һż�¼ A, ���ű� B
      Where a.ִ�в���id = b.Id And
            a.Id = (Select ID
                    From (Select ID, ����ʱ�� From ���˹Һż�¼ Where ����id = n_����id Order By ����ʱ�� Desc)
                    Where Rownum < 2);
    End If;
  
    If v_�������� Is Null And Nvl(n_����, 0) <> 0 Then
    
      Select Max(��������)
      Into v_��������
      From (
             
             Select Distinct a.���� As ��������
             From ���ղ��� A, ������׼��Ŀ B
             Where a.���� = n_���� And a.Id = b.����id And
                   b.�շ�ϸĿid In (Select Distinct �շ�ϸĿid From ������ü�¼ Where ����id = n_����id)
             Union All
             Select Distinct a.���� As ��������
             From ���ղ��� A, ������׼��Ŀ B
             Where a.���� = n_���� And a.Id = b.����id And
                   b.���� In (Select Distinct ���մ���id From ������ü�¼ Where ����id = n_����id))
      Where Rownum < 2;
    End If;
    --medicalInstitution  ҽ�ƻ�������  String  60  ��  ���ա�ҽ�ƻ�����������ʵʩϸ�򡷺͡������������޶�<ҽ�ƻ�����������ʵʩϸ��>�������й����ݵ�֪ͨ��ȷ����ҽ�������������
    v_������Ϣ := ',"medicalInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"';
    --medCareInstitution  ҽ����������  String  60  ��  ҽ��������Ψһ����
    v_������Ϣ := v_������Ϣ || ',"medCareInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_���ջ�������) || '"';
    --medCareTypeCode  ҽ�����ͱ���  String  60  ��  
    v_������Ϣ := v_������Ϣ || ',"medCareTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ�Ƹ��ʽ����) || '"';
    --medicalCareType  ҽ����������  String  60  ��  ȡֵ��Χ����ְ������ҽ�Ʊ��ա�����������ҽ�Ʊ��գ�����������ҽ�Ʊ��ա�����ũ�����ҽ�Ʊ��գ�������ҽ�Ʊ��ա���ҽ����
    v_������Ϣ := v_������Ϣ || ',"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ�Ƹ��ʽ����) || '"';
    --medicalInsuranceID  ����ҽ�����  String  60  ��  �α�����ҽ��ϵͳ�е�Ψһ��ʶ(ҽ����)
    v_������Ϣ := v_������Ϣ || ',"medicalInsuranceID":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ����) || '"';
    --consultationDate  ��������  String  10  ��  ���߾�ҽʱ��
    --          ��ʽ:yyyy-MM-dd
    v_������Ϣ := v_������Ϣ || ',"consultationDate":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"';
    --category  �������  String  200  ��  
    v_������Ϣ := v_������Ϣ || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(v_�����������) || '"';
    --patientCategoryCode  ������ұ���  String  60  ��  
    v_������Ϣ := v_������Ϣ || ',"patientCategoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_������ұ���) || '"';
    --patientNo  ���߾�����  String  20  ��  ����ÿ�ξ���һ�ξ����ɵ�һ���µı�š������ߵǼǺţ�
    v_������Ϣ := v_������Ϣ || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_������) || '"';
    --patientId  ����ΨһID  String  50  ��  ������ҵ��ϵͳ�е�Ψһ��ʶID���������֤���롣
    v_������Ϣ := v_������Ϣ || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(n_����id) || '"';
    --sex  �Ա�  String  2  ��  
    v_������Ϣ := v_������Ϣ || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(v_�����Ա�) || '"';
    --age  ����  String  10  ��  
    v_������Ϣ := v_������Ϣ || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"';
    --caseNumber  ������  String  50  ��  
    v_������Ϣ := v_������Ϣ || ',"caseNumber":"' || b_Einvoice_Request_Test.Zljsonstr(n_�����) || '"';
    --specialDiseasesName  ���ⲡ������  String  200  ��  
    v_������Ϣ := v_������Ϣ || ',"specialDiseasesName":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"';
    -------------------------------------------------------------------------------------------
    --������Ϣ
    v_���� := Null;
    For c_���� In (Select �ֽ�Ԥ��, ֧ƱԤ��, ת��Ԥ��, �����ʻ�֧��, ҽ��ͳ�����֧��, ����ҽ��֧��, �����ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, �ֽ�֧��, 0) As �ֽ��˿�,
                        Decode(Sign(�ֽ�֧��), -1, ֧Ʊ֧��, 0) As ֧Ʊ�˿�, Decode(Sign(�ֽ�֧��), -1, ת��֧��, 0) As ת���˿�,
                        Decode(Sign(�ֽ�֧��), -1, 0, �ֽ�֧��) As �ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, 0, ֧Ʊ֧��) As ֧Ʊ֧��,
                        Decode(Sign(�ֽ�֧��), -1, 0, ת��֧��) As ת��֧��,
                        Nvl(�����ʻ�֧��, 0) + Nvl(ҽ��ͳ�����֧��, 0) + Nvl(����ҽ��֧��, 0) As �����ܶ�,
                        Nvl(�����ܶ�, 0) - Nvl(�����ʻ�֧��, 0) - Nvl(ҽ��ͳ�����֧��, 0) - Nvl(����ҽ��֧��, 0) As �Էѽ��, �����ܶ�, ҽ���������,
                        0 As �����ʻ����
                 From (Select /*+cardinality(b,10)*/
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '�ֽ�', 1, 0), 0) * a.��Ԥ��) As �ֽ�Ԥ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0), 0) * a.��Ԥ��) As ֧ƱԤ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 0, '�ֽ�', 0, 1), 0) * a.��Ԥ��) As ת��Ԥ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '�����˻�֧��', 1, 0)) * a.��Ԥ��) As �����ʻ�֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, 'ҽ��ͳ�����֧��', 1, 0)) * a.��Ԥ��) As ҽ��ͳ�����֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 1, 0)) * a.��Ԥ��) As ����ҽ��֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 0, '�����˻�֧��', 0, 'ҽ��ͳ�����֧��', 0, 1)) *
                              a.��Ԥ��) As �����ֽ�֧��,
                         Max(Decode(Mod(a.��¼����, 10), 1, 0,
                                     Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', �������, '�����˻�֧��', �������, 'ҽ��ͳ�����֧��', �������, ''))) As ҽ���������,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '�ֽ�', 1, 0)) * a.��Ԥ��) As �ֽ�֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0)) * a.��Ԥ��) As ֧Ʊ֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '�ֽ�', 0, '֧Ʊ', 0, 1) * a.��Ԥ��)) As ת��֧��,
                         Sum(��Ԥ��) As �����ܶ�
                        From ����Ԥ����¼ A, Table(l_����id) B, ��Ʊ������� C
                        Where a.����id = b.Column_Value And a.���㷽ʽ = c.���㷽ʽ(+)))
    
     Loop
      --accountPay  �����˻�֧��  Number  14,2  ��  �����߹涨�ø����˻�֧���α��˵�ҽ�Ʒ��ã�������ҽ�Ʊ���Ŀ¼��Χ�ں�Ŀ¼��Χ��ķ��ã���
      --          ���޽���д0
      v_���� := ',"accountPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.�����ʻ�֧��, 0), 1);
      --fundPay  ҽ��ͳ�����֧��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а��涨�ɻ���ҽ�Ʊ���ͳ�����֧���Ľ�
      --          ���޽���д0
      v_���� := v_���� || ',"fundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.ҽ��ͳ�����֧��, 0), 1);
      --otherfundPay  ����ҽ��֧��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а��涨�ɴ󲡱��ա�ҽ�ƾ���������Աҽ�Ʋ��������䡢��ҵ����Ȼ�����ʽ�֧���Ľ�
      v_���� := v_���� || ',"otherfundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.����ҽ��֧��, 0), 1);
      --ownPay  �Էѽ��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а����йع涨�����ڻ���ҽ�Ʊ���Ŀ¼��Χ��ȫ���ɸ���֧���ķ��ã�
      --          ���޽���д0
      v_���� := v_���� || ',"ownPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�Էѽ��, 1);
      --selfConceitedAmt  �����Ը�  Number  14,2  ��  ҽ�������𸶱�׼�ڸ���֧�����ã�
      --          ���޽���д0
      v_���� := v_���� || ',"selfConceitedAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfPayAmt  �����Ը�  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ������ɸ��˸��������ڻ���ҽ�Ʊ���Ŀ¼��Χ���Ը����ֵĽ���չ�����֡����顢���յȴ�����ѷ�ʽ���ɻ��߶���ѵķ��á�����Ϊ��������˰��ҽ��ר��ӿ۳��ţ�Ϣ�����޽���д0
      v_���� := v_���� || ',"selfPayAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfCashPay  �����ֽ�֧��  Number  14,2  ��  ����ͨ���ֽ����п���΢�š�֧����������֧���Ľ�
      --          ���޽���д0
      v_���� := v_���� || ',"selfCashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ֽ�֧��, 1);
      --cashPay  �ֽ�Ԥ������  Number  14,2  ��
      v_���� := v_���� || ',"cashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ�Ԥ��, 1);
      --chequePay  ֧ƱԤ������  Number  14,2  ��
      v_���� := v_���� || ',"chequePay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧ƱԤ��, 1);
      --transferAccountPay  ת��Ԥ������  Number  14,2  ��
      v_���� := v_���� || ',"transferAccountPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת��Ԥ��, 1);
      --cashRecharge  �������(�ֽ�)  Number  14,2  ��
      v_���� := v_���� || ',"cashRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ�֧��, 1);
      --chequeRecharge  �������(֧Ʊ)  Number  14,2  ��
      v_���� := v_���� || ',"chequeRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧Ʊ֧��, 1);
      --transferRecharge  ������ת�ˣ�  Number  14,2  ��
      v_���� := v_���� || ',"transferRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת��֧��, 1);
      --cashRefund  �˻����(�ֽ�)  Number  14,2  ��
      v_���� := v_���� || ',"cashRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ��˿�, 1);
      --chequeRefund  �˽����(֧Ʊ)  Number  14,2  ��
      v_���� := v_���� || ',"chequeRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧Ʊ�˿�, 1);
      --transferRefund  �˽����(ת��)  Number  14,2  ��
      v_���� := v_���� || ',"transferRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת���˿�, 1);
      --ownAcBalance  �����˻����  Number  14,2  ��
      v_���� := v_���� || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ʻ����, 1);
      --reimbursementAmt  �����ܽ��  Number  14,2  ��  ҽ������󷵻ص��ܽ��
      v_���� := v_���� || ',"reimbursementAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ܶ�, 1);
      --balancedNumber  �����  String  100  ��  ҽ����������ɵĺ���/����Ψһֵ
      v_���� := v_���� || ',"balancedNumber":"' || b_Einvoice_Request_Test.Zljsonstr(c_����.ҽ���������) || '"';
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --��������
    v_�ɷ����� := Null;
    For c_���� In (Select /*+cardinality(b,10)*/
                  Nvl(c.��������, Nvl(d.��������, '-')) As ��������, Sum(��Ԥ��) As �����ܶ�
                 From ����Ԥ����¼ A, Table(l_����id) B, �շ��������� C,
                      (Select ���㷽ʽ, �������� From �շ��������� D Where �����id Is Null) D
                 Where a.����id = b.Column_Value And a.�����id = c.�����id(+) And a.���㷽ʽ = c.���㷽ʽ(+) And a.���㷽ʽ = d.���㷽ʽ(+)
                 Group By Nvl(c.��������, Nvl(d.��������, '-'))
                 Order By ��������)
    
     Loop
      --payChannelCode  ������������  String  10  ��
      If v_�ɷ����� Is Null Then
        v_�ɷ����� := Nvl(v_�ɷ�����, '') || '{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.��������, 0)) || '"';
      Else
        v_�ɷ����� := v_�ɷ����� || ',{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.��������, 0)) || '"';
      End If;
      --payChannelValue  �����������  Number  14,2  ��
      v_�ɷ����� := v_�ɷ����� || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.�����ܶ�, 0), 1) || '}';
    End Loop;
  
    If v_�ɷ����� Is Not Null Then
      --payChannelDetail  ���������б�  String  ����  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö����磺��¼4���������б�
      --        ���A-5,JSON��ʽ�б�
      v_�ɷ����� := ',"payChannelDetail":[' || v_�ɷ����� || ']';
    Else
      v_�ɷ����� := ',"payChannelDetail":[]';
    End If;
  
    -------------------------------------------------------------------------------------------
    --����ҽ����Ϣ
    v_����ҽ����Ϣ := Null;
    --otherMedicalList  ����ҽ����Ϣ�б�  String  ����  ��  ��д����δ֪ҽ����Ϣ���ڵ���Ʊ����������ƴ�ӷ�ʽ��ʾ��
    --            ���A-4,JSON��ʽ�б�
    --  infoNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���������ֵ����1�����β������ظ�
    --  infoName  ҽ����Ϣ����  String  100  ��  ����ñ������ͱ��룬�ɲο���¼7ҽ�����������б�
    --  infoValue  ҽ����Ϣֵ  String  100  ��  ����ñ������
    --  infoOther  ҽ��������Ϣ  String  100  ��  ��ҽ������������
  
    -------------------------------------------------------------------------------------------
    --������չ��Ϣ
    v_������չ��Ϣ := Null;
    --otherInfo  ������չ��Ϣ�б�  String  ����  ��  ��д��Ϣ��Ҫ�ڵ���Ʊ���ϵ�����ʾ��������չ��Ϣ��δ֪��Ϣ��
    --          ���A-3,JSON��ʽ�б�
    --  infoNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���������ֵ����1�����β������ظ�
    --  infoName  ��չ��Ϣ����  String  100  ��
    --  infoValue  ��չ��Ϣֵ  String  500  ��
  
    c_������Ϣ := To_Clob('{' || v_Ʊ����Ϣ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_�ɷ�);
    c_������Ϣ := c_������Ϣ || To_Clob(v_֪ͨ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_������Ϣ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_����);
    c_������Ϣ := c_������Ϣ || To_Clob(v_�ɷ�����);
  
    If v_������չ��Ϣ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_������չ��Ϣ);
    End If;
    If v_����ҽ����Ϣ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_����ҽ����Ϣ);
    End If;
    --  eBillRelateNo  ҵ��Ʊ�ݹ�����  String  32  ��  ��һ��ҵ��������Ҫ����N�ŵ���Ʊ�ݣ���N�ŵ���Ʊ��Ӧ��ֵ����һ�£����ں��ڹ�����ѯ
    c_������Ϣ := c_������Ϣ || To_Clob(',"eBillRelateNo":""');
    If v_������ϸ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_������ϸ);
    Else
      c_������Ϣ := c_������Ϣ || c_������ϸ;
    End If;
  
    If v_��ϸ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_��ϸ);
    Else
      c_������Ϣ := c_������Ϣ || c_��ϸ;
    End If;
    c_������Ϣ  := c_������Ϣ || To_Clob('}');
    Reqdata_Out := c_������Ϣ;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Chargedata_Create;

  Procedure Get_Sendcarddata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    --
    ---------------------------------------------------------------------------
    --����:��ȡ������Ʊ����
    --���:
    --    Json_In,��ʽ����
    --  input
    --    occasion N 1  Ӧ�ó���:5-���￨
    --    einvoice_id  N,1 ��ǰ����Ʊ��ID
    --    balance_id N 1  ����ID  occasion=2ʱ��Ԥ��id;occasion<>2��ʾ����id
    --    writeoff_id  N 1  ����ID  occasion=2ʱ������Ԥ��id;occasion<>2��ʾ����id
    --����:
    --  ReqData_Out-���ص�ҵ����������
    --  Totalmoney_Out-Ʊ���ܶ�
    --  Code_Out-��ȡ�Ƿ�ɹ���0-ʧ�ܣ�1-�ɹ�
    --  Message_Out ������Ϣ
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_Ӧ�ó��� Number(2);
    n_����id   ����Ԥ����¼.����id%Type;
    --n_����id   ����Ԥ����¼.����id%Type;
  
    v_ҵ���ʶ   Varchar2(20);
    v_ҵ����ˮ�� Varchar2(50);
  
    v_��Ʊ��       Varchar2(100);
    v_�ɷ�         Varchar2(32767);
    v_Ʊ����Ϣ     Varchar2(32767);
    v_������Ϣ     Varchar2(32767);
    v_֪ͨ         Varchar2(32767);
    v_�ɷ�����     Varchar2(32767);
    v_����         Varchar2(32767);
    v_������չ��Ϣ Varchar2(32767);
    v_����ҽ����Ϣ Varchar2(32767);
    c_��ϸ         Clob;
    v_��ϸ         Varchar2(32767);
    c_������ϸ     Clob;
    v_������ϸ     Varchar2(32767);
    c_������Ϣ     Clob; --���շ��صĽ�����Ϣ��
  
    n_�����       ������Ϣ.�����%Type;
    n_����id       ����Ԥ����¼.����id%Type;
    v_��������     ������ü�¼.����%Type;
    v_�����Ա�     ������ü�¼.�Ա�%Type;
    v_��������     ������ü�¼.����%Type;
    d_ҵ����ʱ�� ������ü�¼.�Ǽ�ʱ��%Type;
    v_�շ�Ա       ������ü�¼.����Ա����%Type;
  
    n_ȱʡ�����id     Number(18);
    v_����ֵ           Varchar2(100);
    n_Ʊ���ܽ��       ������ü�¼.���ʽ��%Type;
    n_����ܶ�         ������ü�¼.���ʽ��%Type;
    n_�û�id           ��Ա��.Id%Type;
    v_����Ա���       ��Ա��.���%Type;
    v_����Ա����       ��Ա��.����%Type;
    v_Temp             Varchar2(32767);
    v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
    v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
    n_����             ���ս����¼.����%Type;
    v_���ջ�������     �������.���ջ�������%Type;
    n_ҽ�����         ������ü�¼.ҽ�����%Type;
    n_�Һ�id           ������ü�¼.�Һ�id%Type;
    v_��������         ���ղ���.����%Type;
    v_��������         Varchar2(20);
    v_������ұ���     ���ű�.����%Type;
    v_�����������     ���ű�.����%Type;
    v_������         Varchar2(50);
    v_����ids          Varchar2(32767);
    v_ҽ����           �����ʻ�.ҽ����%Type;
    l_����id           t_NumList := t_NumList();
    v_�汾��           Varchar2(30);
    n_����Ʊ��id       ����Ʊ��ʹ�ü�¼.Id%Type;
  
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_Ӧ�ó��� := Nvl(j_Json.Get_Number('occasion'), 0);
    n_����id   := j_Json.Get_Number('balance_id');
    --n_����id     := Nvl(j_Json.Get_Number('writeoff_id'), 0);
    n_����Ʊ��id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
  
    If Nvl(n_Ӧ�ó���, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '��Ч��Ӧ�ó���';
      Return;
    End If;
  
    Select Nvl(Max(����ֵ), 'V2.0.3')
    Into v_�汾��
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '֧�ְ汾';
  
    b_Einvoice_Request_Test.Get_Identity(n_�û�id, v_����Ա���, v_����Ա����);
  
    --v_ҵ���ʶ:01  סԺ,02  ����,03  ����,04  ����,05  �������,06  �Һ�,07  סԺԤ����,08  ���Ԥ����
    Select Decode(n_Ӧ�ó���, 1, '02', 2, '07', 3, '01', 4, '06', 5, '02', '02') Into v_ҵ���ʶ From Dual;
  
    n_Ʊ���ܽ��   := 0;
    d_ҵ����ʱ�� := Null;
    v_����ids      := Null;
    c_��ϸ         := Null;
    v_��ϸ         := Null;
    For c_�շ�ϸĿ In (Select Min(a.Id) As ����id, a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, Max(a.���㵥λ) As ���㵥λ,
                          Sum(a.��׼����) As �۸�, Avg(Nvl(a.����, 1) * Nvl(a.����, 0)) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��,
                          Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.ʵ�ս��) - Sum(a.ͳ����) As �Էѽ��,
                          Max(t.����) As ҽ����Ŀ����, Max(t.����) As ҽ����Ŀ����, Max(t.ͳ��ȶ�) As ҽ����������, Max(a.ժҪ) As ��ע,
                          Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����) As ����,
                          Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, Max('') As ���ʽ����,
                          Max(a.�վݷ�Ŀ) As �վݷ�Ŀ, Max(c.����) As �վݷ�Ŀ����, Max(a.ҽ�����) As ҽ�����, Max(0) As �Һ�id,
                          Max(d.����) As ������, Max(d.���) As �������, Max(b.����) As ��Ŀ����, Max(b.����) As ��Ŀ����, Max(b.���) As ���,
                          Max(q.ҩƷ����) As ҩƷ����
                   From סԺ���ü�¼ A, �շ���ĿĿ¼ B, �վݷ�Ŀ C, �շ���� D, ҩƷ��� M, ҩƷ���� Q, ������ĿĿ¼ J, ����֧������ T, ֧�������� S
                   Where a.No In (Select Distinct NO From סԺ���ü�¼ Where ����id = n_����id) And a.��¼���� = 5 And
                         a.�շ���� = d.����(+) And a.�շ�ϸĿid = b.Id And a.�վݷ�Ŀ = c.����(+) And a.�շ�ϸĿid = m.ҩƷid(+) And
                         m.ҩ��id = q.ҩ��id(+) And q.ҩ��id = j.Id(+) And a.���մ���id = t.Id(+) And t.����(+) = 1 And
                         a.���մ���id = s.���մ���id(+)
                   Group By a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, c.����, c.����, j.����, j.����
                   Order By NO, ���) Loop
      If v_�������� Is Null Then
        v_�������� := c_�շ�ϸĿ.����;
        v_�����Ա� := c_�շ�ϸĿ.�Ա�;
        v_�������� := c_�շ�ϸĿ.����;
        n_����id   := c_�շ�ϸĿ.����id;
      End If;
      If d_ҵ����ʱ�� Is Null And Nvl(c_�շ�ϸĿ.��¼״̬, 0) = 1 Then
        --ȡԭʼҵ����ʱ��
        d_ҵ����ʱ�� := c_�շ�ϸĿ.�Ǽ�ʱ��;
        v_�շ�Ա       := c_�շ�ϸĿ.����Ա����;
      End If;
      If v_ҽ�Ƹ��ʽ���� Is Null Then
        v_ҽ�Ƹ��ʽ���� := c_�շ�ϸĿ.���ʽ����;
      End If;
      If Nvl(n_ҽ�����, 0) = 0 Then
        n_ҽ����� := c_�շ�ϸĿ.ҽ�����;
      End If;
      If Nvl(n_�Һ�id, 0) = 0 Then
        n_�Һ�id := c_�շ�ϸĿ.�Һ�id;
      End If;
    
      If Instr(Nvl(v_����ids, '') || ',', ',' || c_�շ�ϸĿ.����id || ',') = 0 Then
        l_����id.Extend;
        l_����id(l_����id.Count) := c_�շ�ϸĿ.����id;
      End If;
    
      --listDetailNo  ��ϸ��ˮ��  String  60  ��  ��ϸ��ˮ��
      v_Temp := '{' || '"listDetailNo":"' || b_Einvoice_Request_Test.Zljsonstr(LPad(c_�շ�ϸĿ.����id, 20, '0')) || '"';
      --chargeCode  �շ���Ŀ����  String  50  ��  ��дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���,�磺��λ�ѡ�����
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�վݷ�Ŀ����) || '"';
      --chargeName  �շ���Ŀ����  String  100  ��
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�վݷ�Ŀ) || '"';
      --prescribeCode  ��������  String  60  ��
      v_Temp := v_Temp || ',' || '"prescribeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.No) || '"';
      --listTypeCode  ҩƷ������  String  50  ��  ��ҩƷ�������01��������д
      v_Temp := v_Temp || ',' || '"listTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.������) || '"';
      --listTypeName  ҩƷ�������  String  50  ��  ��ҩƷ�������ƣ��������࿹��Ⱦҩ��
      v_Temp := v_Temp || ',' || '"listTypeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�������) || '"';
      --code  ����  String  50  ��  ��ҩƷ���룬������д
      v_Temp := v_Temp || ',' || '"code":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��Ŀ����) || '"';
      --name  ҩƷ����  String  50  ��  ��ҩƷ���ƣ��������Ƶ�
      v_Temp := v_Temp || ',' || '"name":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��Ŀ����) || '"';
      --form  ����  String  50  ��
      v_Temp := v_Temp || ',' || '"form":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҩƷ����) || '"';
      --specification  ���  String  50  ��
      v_Temp := v_Temp || ',' || '"specification":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���) || '"';
      --unit  ������λ   String  20  ��
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���㵥λ) || '"';
      --std  ����  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�۸�, 1);
      --number  ����  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.����, 1);
      --amt  ���  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ʵ�ս��, 1);
      --selfAmt  �Էѽ��  Number  14,6  ��  ���޽���д0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�Էѽ��, 1);
      --receivableAmt  Ӧ�շ���  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"receivableAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.Ӧ�ս��, 1);
      --medicalCareType  ҽ��ҩƷ����  String  1  ��  1�����Ը�/��
      --          2�����Ը�/��
      --          3��ȫ�Ը�/��
      v_Temp := v_Temp || ',' || '"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����Ŀ����) || '"';
      --medCareItemType  ҽ����Ŀ����  String  100  ��
      v_Temp := v_Temp || ',' || '"medCareItemType":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����Ŀ����) || '"';
      --medReimburseRate  ҽ����������  Number  3,2  ��
      v_Temp := v_Temp || ',' || '"medReimburseRate":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����������, 1);
      --remark  ��ע  String  200  ��
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��ע) || '"';
      --sortNo  ���  Integer  ����  ��  ���
      v_Temp := v_Temp || ',' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���, 1);
      --chrgtype  ��������  String  50  ��
      v_Temp := v_Temp || ',' || '"chrgtype":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��������) || '"}';
    
      If Length(Nvl(v_��ϸ, '') || v_Temp) > 32700 Then
        If c_��ϸ Is Null Then
          c_��ϸ := To_Clob(v_��ϸ);
        Else
          c_��ϸ := c_��ϸ || To_Clob(',' || v_��ϸ);
        End If;
        v_��ϸ := Null;
      End If;
    
      If v_��ϸ Is Null Then
        v_��ϸ := v_Temp;
      Else
        v_��ϸ := v_��ϸ || ',' || v_Temp;
      End If;
    End Loop;
  
    If v_��ϸ Is Not Null And c_��ϸ Is Not Null Then
      --listDetail  �嵥��Ŀ��ϸ  String  ����  ��  ���A-2,JSON��ʽ�б�
      c_��ϸ := c_��ϸ || ',' || To_Clob(v_��ϸ);
      c_��ϸ := To_Clob(',"listDetail":[') || c_��ϸ || To_Clob(']');
    
      v_��ϸ := Null;
    Elsif v_��ϸ Is Not Null Then
      v_��ϸ := ',"listDetail":[' || v_��ϸ || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --������ϸ
    v_������ϸ := Null;
    c_������ϸ := Null;
    For c_����ͳ�� In (Select Rownum As ���, �վݷ�Ŀ����, �վݷ�Ŀ����, ����, ���㵥λ, Round(����, 2) As ����, Round(���ʽ��, 2) As ���ʽ��,
                          Round(�Էѽ��, 2) As �Էѽ��, ��ע, ���ʽ�� - Round(���ʽ��, 2) As ����
                   From (Select /*+cardinality(b,10)*/
                           c.���� As �վݷ�Ŀ����, a.�վݷ�Ŀ As �վݷ�Ŀ����, 1 As ����, '' As ���㵥λ, Sum(a.���ʽ��) As ����, a.�վݷ�Ŀ,
                           Sum(a.���ʽ��) As ���ʽ��, Sum(a.���ʽ��) - Sum(a.ͳ����) As �Էѽ��, '' As ��ע
                          From סԺ���ü�¼ A, Table(l_����id) B, �վݷ�Ŀ C
                          Where a.����id = b.Column_Value And a.�վݷ�Ŀ = c.����(+)
                          Group By c.����, a.�վݷ�Ŀ)) Loop
      --sortNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���շ���Ŀ���ֵ����1�����β������ظ�
      v_Temp := '{' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���, 1);
      --chargeCode  �շ���Ŀ����  String  50  ��  ��дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�վݷ�Ŀ����) || '"';
      --chargeName  �շ���Ŀ����  String  100  ��  ��дҵ��ϵͳ�ڲ���Ŀ����
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�վݷ�Ŀ����) || '"';
      --unit  ������λ  String  20  ��
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���㵥λ) || '"';
      --std  �շѱ�׼  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.����, 1);
      --number  ����  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.����, 1);
      --amt  ���  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���ʽ��, 1);
      --selfAmt  �Էѽ��  Number  14,2  ��  ���޽���д0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�Էѽ��, 1);
      --remark  ��ע  String  200  ��
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.��ע) || '"}';
    
      If Length(Nvl(v_������ϸ, '') || v_Temp) > 32700 Then
        If c_������ϸ Is Null Then
          c_������ϸ := To_Clob(v_������ϸ);
        Else
          c_������ϸ := c_������ϸ || To_Clob(',' || v_������ϸ);
        End If;
        v_������ϸ := Null;
      End If;
    
      If v_������ϸ Is Null Then
        v_������ϸ := v_Temp;
      Else
        v_������ϸ := v_������ϸ || ',' || v_Temp;
      End If;
    
      n_Ʊ���ܽ�� := Nvl(n_Ʊ���ܽ��, 0) + Nvl(c_����ͳ��.���ʽ��, 0);
      n_����ܶ�   := Nvl(n_����ܶ�, 0) + Nvl(c_����ͳ��.����, 0);
    End Loop;
    Totalmoney_Out := n_Ʊ���ܽ��;
    If v_������ϸ Is Not Null And c_������ϸ Is Not Null Then
      c_������ϸ := c_������ϸ || ',' || To_Clob(v_������ϸ);
      ----chargeDetail �շ���Ŀ��ϸ  String  ����  ��  ���A-1,JSON��ʽ�б�
      c_������ϸ := To_Clob(',"chargeDetail":[') || c_������ϸ || To_Clob(']');
      v_������ϸ := Null;
    Elsif v_������ϸ Is Not Null Then
      v_������ϸ := ',"chargeDetail":[' || v_������ϸ || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --Ʊ����Ϣ
    --Select Count(1) Into n_���ϴ��� From ����Ʊ��ʹ�ü�¼ Where Ʊ�� = n_Ӧ�ó��� And ����id = n_����id;
    --����ID||_||����Ʊ��ID
    v_ҵ����ˮ�� := n_����id || '_' || Nvl(n_����Ʊ��id, 0);
  
    v_��Ʊ�� := b_Einvoice_Request_Test.Get_Einvoice_Node(v_����Ա���, v_����Ա����, n_�û�id);
  
    v_Ʊ����Ϣ := '"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ����ˮ��) || '"'; --ҵ����ˮ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"busType":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ���ʶ) || '"'; --ҵ���ʶ
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --��������
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"busDateTime":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(d_ҵ����ʱ��, 'yyyymmddHH24miss') || '000') || '"'; --ҵ����ʱ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_��Ʊ��) || '"'; --��Ʊ�����:ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(v_�շ�Ա) || '"'; --�շ�Ա
  
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_����Ա����) || '"'; --Ʊ�ݱ�����
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"totalAmt":' || b_Einvoice_Request_Test.Zljsonstr(n_Ʊ���ܽ��, 1); --��Ʊ�ܽ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"remark":"' || '' || '"'; --��ע
    -------------------------------------------------------------------------------------------
  
    --ȡ�ɷ���Ϣ
    v_�ɷ� := Null;
    For c_�ɷ� In (Select Max(Decode(��Ϣ��, '������', ��Ϣֵ, '֧��������', ��Ϣֵ, '')) As ֧��������,
                        Max(Decode(��Ϣ��, 'ҽ��֧��������', ��Ϣֵ, 'ҽ��������', ��Ϣֵ, '')) As ҽ��֧��������,
                        Max(Decode(Upper(��Ϣ��), '֧�������ں�USERID', ��Ϣֵ, '')) As ֧�������ں�userid,
                        Max(Decode(Upper(��Ϣ��), '֧����С����USERID', ��Ϣֵ, '')) As ֧����С����userid,
                        Max(Decode(Upper(��Ϣ��), '΢�Ź��ں�OPENID', ��Ϣֵ, '')) As ΢�Ź��ں�openid,
                        Max(Decode(Upper(��Ϣ��), '΢��С����OPENID', ��Ϣֵ, '')) As ΢��С����openid
                 From (Select ��Ϣ��, ��Ϣֵ
                        From ������Ϣ�ӱ�
                        Where ����id = n_����id And ��Ϣ�� In ('֧�������ں�USERID', '֧����С����USERID', '΢�Ź��ں�OPENID', '΢��С����OPENID')
                        Union All
                        Select ������Ŀ, ��������
                        From �������㽻��
                        Where ����id In (Select ID From ����Ԥ����¼ Where ����id = n_����id) And ������Ŀ Like '%������')) Loop
      v_�ɷ� := ',"alipayCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.֧�������ں�userid) || '"'; --����֧�����˻�
      v_�ɷ� := v_�ɷ� || ',"weChatOrderNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.֧��������) || '"'; --΢��֧��������
      If v_�汾�� = 'V3.1.0' Then
        --�ð汾���д˽ӵ�
        v_�ɷ� := v_�ɷ� || ',"weChatMedTransNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.ҽ��֧��������) || '"'; --΢��ҽ��֧��������
      End If;
    
      If c_�ɷ�.΢�Ź��ں�openid Is Not Null Then
        v_�ɷ� := v_�ɷ� || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.΢�Ź��ں�openid) || '"'; --΢�Ź��ںŻ�С�����û�ID
      Else
        v_�ɷ� := v_�ɷ� || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.΢��С����openid) || '"'; --΢�Ź��ںŻ�С�����û�ID
      End If;
      Exit;
    End Loop;
  
    -------------------------------------------------------------------------------------------
    --ȡ֪ͨ��Ϣ
    Select To_Number(Max(����ֵ))
    Into n_ȱʡ�����id
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = 'ȱʡ�����ID';
    v_֪ͨ := Null;
    For c_֪ͨ In (Select Max(a.����id) As ����id, Max(a.����) As ����, Max(a.�ֻ���) As �ֻ���, Max(a.Email) As Email, Max(1) As �ɿ�����,
                        Max(a.���֤��) As ���֤��, Max(m.����) As �����, Max(m.����) As ����, Max(a.�����) As �����
                 From ������Ϣ A,
                      (
                        
                        Select ����id, ����, ����, ����
                        From (Select b.����id, c.����, c.����, b.����, Decode(b.�����id, n_ȱʡ�����id, 2, c.ȱʡ��־) As ȱʡ��־
                                From ����ҽ�ƿ���Ϣ B, ҽ�ƿ���� C
                                Where b.�����id = c.Id And b.����id = n_����id
                                Order By ȱʡ��־)
                        Where Rownum < 2) M
                 Where a.����id = m.����id(+)) Loop
    
      v_֪ͨ := ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.�ֻ���) || '"'; --�����ֻ�����
      v_֪ͨ := v_֪ͨ || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.Email) || '"'; --���������ַ
      If v_�汾�� = 'V3.1.0' Then
        --�ð汾���д˽ӵ�
        v_֪ͨ := v_֪ͨ || ',"payerType":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.�ɿ�����) || '"'; --����������
      End If;
      v_֪ͨ := v_֪ͨ || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.���֤��) || '"'; --ͳһ������ô���
    
      If c_֪ͨ.����� Is Not Null Then
        v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.�����) || '"'; --������
        v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.����) || '"'; --����
      Elsif c_֪ͨ.���֤�� Is Not Null Then
        Select Nvl(Max(����ֵ), '99998')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '���֤�������ͱ��';
        v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
        v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.���֤��) || '"'; --����
      Else
        --û��һ�ſ����̶�һ�ֿ����
        Select Nvl(Max(����ֵ), '99999')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ������';
        v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
        Select Nvl(Max(����ֵ), '-')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ���';
        v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --����
      End If;
      If Nvl(n_�����, 0) = 0 Then
        n_����� := c_֪ͨ.�����;
      
      End If;
    
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --������Ϣ
    Select Max(����) Into v_Temp From zlRegInfo Where ��Ŀ = 'ҽ�ƻ�������';
  
    --����:1-�շ�;2-���㣨����סԺ���㡢����������㣩��3-Ԥ��
    Select Max(a.����), Max(b.���ջ�������), Max(Nvl(a.��������, c.����))
    Into n_����, v_���ջ�������, v_��������
    From ���ս����¼ A, ������� B, ���ղ��� C
    Where a.���� = b.��� And a.����id = c.Id(+) And a.��¼id = n_����id And a.���� = Decode(n_Ӧ�ó���, 2, 3, 3, 2, 1);
  
    Select Max(����) Into v_ҽ�Ƹ��ʽ���� From ҽ�Ƹ��ʽ Where ���� = v_ҽ�Ƹ��ʽ����;
    If Nvl(n_����, 0) <> 0 Then
      Select Max(ҽ����) Into v_ҽ���� From �����ʻ� Where ����id = n_����id And ���� = n_����;
    End If;
  
    v_������ := Null;
    If Nvl(n_ҽ�����, 0) <> 0 Then
      Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')), Max(b.����), Max(b.����), Max(a.No)
      Into v_��������, v_������ұ���, v_�����������, v_������
      From ���˹Һż�¼ A, ���ű� B
      Where a.ִ�в���id = b.Id And
            a.No = (Select Max(�Һŵ�) From ����ҽ����¼ Where ID = n_ҽ����� Or ���id = n_ҽ�����);
    Elsif Nvl(n_�Һ�id, 0) <> 0 Then
      Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')), Max(b.����), Max(b.����), Max(a.No)
      Into v_��������, v_������ұ���, v_�����������, v_������
      From ���˹Һż�¼ A, ���ű� B
      Where a.ִ�в���id = b.Id And a.Id = n_�Һ�id;
    End If;
    If v_������ Is Null Then
      --ȡ���һ�ιҺ�ID
      Select Max(a.Id), Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')), Max(b.����), Max(b.����), Max(a.No)
      Into n_�Һ�id, v_��������, v_������ұ���, v_�����������, v_������
      From ���˹Һż�¼ A, ���ű� B
      Where a.ִ�в���id = b.Id And
            a.Id = (Select ID
                    From (Select ID, ����ʱ�� From ���˹Һż�¼ Where ����id = n_����id Order By ����ʱ�� Desc)
                    Where Rownum < 2);
    End If;
  
    If v_�������� Is Null And Nvl(n_����, 0) <> 0 Then
    
      Select Max(��������)
      Into v_��������
      From (
             
             Select Distinct a.���� As ��������
             From ���ղ��� A, ������׼��Ŀ B
             Where a.���� = n_���� And a.Id = b.����id And
                   b.�շ�ϸĿid In (Select Distinct �շ�ϸĿid From ������ü�¼ Where ����id = n_����id)
             Union All
             Select Distinct a.���� As ��������
             From ���ղ��� A, ������׼��Ŀ B
             Where a.���� = n_���� And a.Id = b.����id And
                   b.���� In (Select Distinct ���մ���id From ������ü�¼ Where ����id = n_����id))
      Where Rownum < 2;
    End If;
    v_������Ϣ := ',"medicalInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"'; --ҽ�ƻ�������
    v_������Ϣ := v_������Ϣ || ',"medCareInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_���ջ�������) || '"'; --ҽ��������Ψһ����
    v_������Ϣ := v_������Ϣ || ',"medCareTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ�Ƹ��ʽ����) || '"'; --ҽ�����ͱ���
    v_������Ϣ := v_������Ϣ || ',"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ�Ƹ��ʽ����) || '"'; --ȡֵ��Χ����ְ������ҽ�Ʊ��ա�����������ҽ�Ʊ��գ�����������ҽ�Ʊ��ա�����ũ�����ҽ�Ʊ��գ�������ҽ�Ʊ��ա���ҽ����
    v_������Ϣ := v_������Ϣ || ',"medicalInsuranceID":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ����) || '"';
    v_������Ϣ := v_������Ϣ || ',"consultationDate":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --���߾�ҽʱ��
    v_������Ϣ := v_������Ϣ || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(v_�����������) || '"'; --�������
    v_������Ϣ := v_������Ϣ || ',"patientCategoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_������ұ���) || '"'; --������ұ���
    --patientNo  ���߾�����  String  20  ��  ����ÿ�ξ���һ�ξ����ɵ�һ���µı�š������ߵǼǺţ�
    v_������Ϣ := v_������Ϣ || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_������) || '"';
    v_������Ϣ := v_������Ϣ || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(n_����id) || '"'; --������ҵ��ϵͳ�е�Ψһ��ʶID���������֤���롣
    v_������Ϣ := v_������Ϣ || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(v_�����Ա�) || '"'; --�Ա�
    v_������Ϣ := v_������Ϣ || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --����
    v_������Ϣ := v_������Ϣ || ',"caseNumber":"' || b_Einvoice_Request_Test.Zljsonstr(n_�����) || '"'; --������
    v_������Ϣ := v_������Ϣ || ',"specialDiseasesName":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --���ⲡ������
  
    -------------------------------------------------------------------------------------------
    --������Ϣ
    v_���� := Null;
    For c_���� In (Select �ֽ�Ԥ��, ֧ƱԤ��, ת��Ԥ��, �����ʻ�֧��, ҽ��ͳ�����֧��, ����ҽ��֧��, �����ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, �ֽ�֧��, 0) As �ֽ��˿�,
                        Decode(Sign(�ֽ�֧��), -1, ֧Ʊ֧��, 0) As ֧Ʊ�˿�, Decode(Sign(�ֽ�֧��), -1, ת��֧��, 0) As ת���˿�,
                        Decode(Sign(�ֽ�֧��), -1, 0, �ֽ�֧��) As �ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, 0, ֧Ʊ֧��) As ֧Ʊ֧��,
                        Decode(Sign(�ֽ�֧��), -1, 0, ת��֧��) As ת��֧��,
                        Nvl(�����ʻ�֧��, 0) + Nvl(ҽ��ͳ�����֧��, 0) + Nvl(����ҽ��֧��, 0) As �����ܶ�,
                        Nvl(�����ܶ�, 0) - Nvl(�����ʻ�֧��, 0) - Nvl(ҽ��ͳ�����֧��, 0) - Nvl(����ҽ��֧��, 0) As �Էѽ��, �����ܶ�, ҽ���������,
                        0 As �����ʻ����
                 From (Select /*+cardinality(b,10)*/
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '�ֽ�', 1, 0), 0) * a.��Ԥ��) As �ֽ�Ԥ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0), 0) * a.��Ԥ��) As ֧ƱԤ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 0, '�ֽ�', 0, 1), 0) * a.��Ԥ��) As ת��Ԥ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '�����˻�֧��', 1, 0)) * a.��Ԥ��) As �����ʻ�֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, 'ҽ��ͳ�����֧��', 1, 0)) * a.��Ԥ��) As ҽ��ͳ�����֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 1, 0)) * a.��Ԥ��) As ����ҽ��֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 0, '�����˻�֧��', 0, 'ҽ��ͳ�����֧��', 0, 1)) *
                              a.��Ԥ��) As �����ֽ�֧��,
                         Max(Decode(Mod(a.��¼����, 10), 1, 0,
                                     Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', �������, '�����˻�֧��', �������, 'ҽ��ͳ�����֧��', �������, ''))) As ҽ���������,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '�ֽ�', 1, 0)) * a.��Ԥ��) As �ֽ�֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0)) * a.��Ԥ��) As ֧Ʊ֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '�ֽ�', 0, '֧Ʊ', 0, 1) * a.��Ԥ��)) As ת��֧��,
                         Sum(��Ԥ��) As �����ܶ�
                        From ����Ԥ����¼ A, Table(l_����id) B, ��Ʊ������� C
                        Where a.����id = b.Column_Value And a.���㷽ʽ = c.���㷽ʽ(+)))
    
     Loop
      --accountPay  �����˻�֧��  Number  14,2  ��  �����߹涨�ø����˻�֧���α��˵�ҽ�Ʒ��ã�������ҽ�Ʊ���Ŀ¼��Χ�ں�Ŀ¼��Χ��ķ��ã���
      --          ���޽���д0
      v_���� := ',"accountPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.�����ʻ�֧��, 0), 1);
      --fundPay  ҽ��ͳ�����֧��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а��涨�ɻ���ҽ�Ʊ���ͳ�����֧���Ľ�
      --          ���޽���д0
      v_���� := v_���� || ',"fundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.ҽ��ͳ�����֧��, 0), 1);
      --otherfundPay  ����ҽ��֧��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а��涨�ɴ󲡱��ա�ҽ�ƾ���������Աҽ�Ʋ��������䡢��ҵ����Ȼ�����ʽ�֧���Ľ�
      v_���� := v_���� || ',"otherfundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.����ҽ��֧��, 0), 1);
      --ownPay  �Էѽ��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а����йع涨�����ڻ���ҽ�Ʊ���Ŀ¼��Χ��ȫ���ɸ���֧���ķ��ã�
      --          ���޽���д0
      v_���� := v_���� || ',"ownPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�Էѽ��, 1);
      --selfConceitedAmt  �����Ը�  Number  14,2  ��  ҽ�������𸶱�׼�ڸ���֧�����ã�
      --          ���޽���д0
      v_���� := v_���� || ',"selfConceitedAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfPayAmt  �����Ը�  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ������ɸ��˸��������ڻ���ҽ�Ʊ���Ŀ¼��Χ���Ը����ֵĽ���չ�����֡����顢���յȴ�����ѷ�ʽ���ɻ��߶���ѵķ��á�����Ϊ��������˰��ҽ��ר��ӿ۳��ţ�Ϣ�����޽���д0
      v_���� := v_���� || ',"selfPayAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfCashPay  �����ֽ�֧��  Number  14,2  ��  ����ͨ���ֽ����п���΢�š�֧����������֧���Ľ�
      --          ���޽���д0
      v_���� := v_���� || ',"selfCashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ֽ�֧��, 1);
      --cashPay  �ֽ�Ԥ������  Number  14,2  ��
      v_���� := v_���� || ',"cashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ�Ԥ��, 1);
      --chequePay  ֧ƱԤ������  Number  14,2  ��
      v_���� := v_���� || ',"chequePay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧ƱԤ��, 1);
      --transferAccountPay  ת��Ԥ������  Number  14,2  ��
      v_���� := v_���� || ',"transferAccountPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת��Ԥ��, 1);
      --cashRecharge  �������(�ֽ�)  Number  14,2  ��
      v_���� := v_���� || ',"cashRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ�֧��, 1);
      --chequeRecharge  �������(֧Ʊ)  Number  14,2  ��
      v_���� := v_���� || ',"chequeRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧Ʊ֧��, 1);
      --transferRecharge  ������ת�ˣ�  Number  14,2  ��
      v_���� := v_���� || ',"transferRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת��֧��, 1);
      --cashRefund  �˻����(�ֽ�)  Number  14,2  ��
      v_���� := v_���� || ',"cashRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ��˿�, 1);
      --chequeRefund  �˽����(֧Ʊ)  Number  14,2  ��
      v_���� := v_���� || ',"chequeRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧Ʊ�˿�, 1);
      --transferRefund  �˽����(ת��)  Number  14,2  ��
      v_���� := v_���� || ',"transferRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת���˿�, 1);
      --ownAcBalance  �����˻����  Number  14,2  ��
      v_���� := v_���� || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ʻ����, 1);
      --reimbursementAmt  �����ܽ��  Number  14,2  ��  ҽ������󷵻ص��ܽ��
      v_���� := v_���� || ',"reimbursementAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ܶ�, 1);
      --balancedNumber  �����  String  100  ��  ҽ����������ɵĺ���/����Ψһֵ
      v_���� := v_���� || ',"balancedNumber":"' || b_Einvoice_Request_Test.Zljsonstr(c_����.ҽ���������) || '"';
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --��������
    v_�ɷ����� := Null;
    For c_���� In (Select /*+cardinality(b,10)*/
                  Nvl(c.��������, Nvl(d.��������, '-')) As ��������, Sum(��Ԥ��) As �����ܶ�
                 From ����Ԥ����¼ A, Table(l_����id) B, �շ��������� C,
                      (Select ���㷽ʽ, �������� From �շ��������� D Where �����id Is Null) D
                 Where a.����id = b.Column_Value And a.�����id = c.�����id(+) And a.���㷽ʽ = c.���㷽ʽ(+) And a.���㷽ʽ = d.���㷽ʽ(+)
                 Group By Nvl(c.��������, Nvl(d.��������, '-'))
                 Order By ��������)
    
     Loop
      --payChannelCode  ������������  String  10  ��
      If v_�ɷ����� Is Null Then
        v_�ɷ����� := Nvl(v_�ɷ�����, '') || '{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.��������, 0)) || '"';
      Else
        v_�ɷ����� := v_�ɷ����� || ',{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.��������, 0)) || '"';
      End If;
      --payChannelValue  �����������  Number  14,2  ��
      v_�ɷ����� := v_�ɷ����� || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.�����ܶ�, 0), 1) || '}';
    End Loop;
  
    If v_�ɷ����� Is Not Null Then
      --payChannelDetail  ���������б�  String  ����  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö����磺��¼4���������б�
      --        ���A-5,JSON��ʽ�б�
      v_�ɷ����� := ',"payChannelDetail":[' || v_�ɷ����� || ']';
    Else
      v_�ɷ����� := ',"payChannelDetail":[]';
    End If;
  
    -------------------------------------------------------------------------------------------
    --����ҽ����Ϣ
    v_����ҽ����Ϣ := Null;
    --otherMedicalList  ����ҽ����Ϣ�б�  String  ����  ��  ��д����δ֪ҽ����Ϣ���ڵ���Ʊ����������ƴ�ӷ�ʽ��ʾ��
    --            ���A-4,JSON��ʽ�б�
    --  infoNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���������ֵ����1�����β������ظ�
    --  infoName  ҽ����Ϣ����  String  100  ��  ����ñ������ͱ��룬�ɲο���¼7ҽ�����������б�
    --  infoValue  ҽ����Ϣֵ  String  100  ��  ����ñ������
    --  infoOther  ҽ��������Ϣ  String  100  ��  ��ҽ������������
  
    -------------------------------------------------------------------------------------------
    --������չ��Ϣ
    v_������չ��Ϣ := Null;
    --otherInfo  ������չ��Ϣ�б�  String  ����  ��  ��д��Ϣ��Ҫ�ڵ���Ʊ���ϵ�����ʾ��������չ��Ϣ��δ֪��Ϣ��
    --          ���A-3,JSON��ʽ�б�
    --  infoNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���������ֵ����1�����β������ظ�
    --  infoName  ��չ��Ϣ����  String  100  ��
    --  infoValue  ��չ��Ϣֵ  String  500  ��
  
    c_������Ϣ := To_Clob('{' || v_Ʊ����Ϣ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_�ɷ�);
    c_������Ϣ := c_������Ϣ || To_Clob(v_֪ͨ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_������Ϣ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_����);
    c_������Ϣ := c_������Ϣ || To_Clob(v_�ɷ�����);
  
    If v_������չ��Ϣ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_������չ��Ϣ);
    End If;
    If v_����ҽ����Ϣ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_����ҽ����Ϣ);
    End If;
    --  eBillRelateNo  ҵ��Ʊ�ݹ�����  String  32  ��  ��һ��ҵ��������Ҫ����N�ŵ���Ʊ�ݣ���N�ŵ���Ʊ��Ӧ��ֵ����һ�£����ں��ڹ�����ѯ
    c_������Ϣ := c_������Ϣ || To_Clob(',"eBillRelateNo":""');
    If v_������ϸ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_������ϸ);
    Else
      c_������Ϣ := c_������Ϣ || c_������ϸ;
    End If;
  
    If v_��ϸ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_��ϸ);
    Else
      c_������Ϣ := c_������Ϣ || c_��ϸ;
    End If;
    c_������Ϣ  := c_������Ϣ || To_Clob('}');
    Reqdata_Out := c_������Ϣ;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Sendcarddata_Create;

  Procedure Get_Registerdata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    --
    ---------------------------------------------------------------------------
    --����:��ȡ�Һſ�Ʊ����
    --���:
    --    Json_In,��ʽ����
    --  input
    --    occasion N 1  Ӧ�ó���: 4-�Һ�
    --    einvoice_id  N,1 ��ǰ����Ʊ��ID
    --    balance_id N 1  ����ID  occasion=2ʱ��Ԥ��id;occasion<>2��ʾ����id
    --    writeoff_id  N 1  ����ID  occasion=2ʱ������Ԥ��id;occasion<>2��ʾ����id
    --����:
    --  ReqData_Out-���ص�ҵ����������
    --  Totalmoney_Out-Ʊ���ܶ�
    --  Code_Out-��ȡ�Ƿ�ɹ���0-ʧ�ܣ�1-�ɹ�
    --  Message_Out ������Ϣ
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_Ӧ�ó��� Number(2);
    n_����id   ����Ԥ����¼.����id%Type;
    --n_����id   ����Ԥ����¼.����id%Type;
  
    v_ҵ���ʶ   Varchar2(20);
    v_ҵ����ˮ�� Varchar2(50);
  
    v_��Ʊ��       Varchar2(100);
    v_�ɷ�         Varchar2(32767);
    v_Ʊ����Ϣ     Varchar2(32767);
    v_������Ϣ     Varchar2(32767);
    v_֪ͨ         Varchar2(32767);
    v_�ɷ�����     Varchar2(32767);
    v_����         Varchar2(32767);
    v_������չ��Ϣ Varchar2(32767);
    v_����ҽ����Ϣ Varchar2(32767);
    c_��ϸ         Clob;
    v_��ϸ         Varchar2(32767);
    c_������ϸ     Clob;
    v_������ϸ     Varchar2(32767);
    c_������Ϣ     Clob; --���շ��صĽ�����Ϣ��
  
    n_�����       ������Ϣ.�����%Type;
    n_����id       ����Ԥ����¼.����id%Type;
    v_��������     ������ü�¼.����%Type;
    v_�����Ա�     ������ü�¼.�Ա�%Type;
    v_��������     ������ü�¼.����%Type;
    d_ҵ����ʱ�� ������ü�¼.�Ǽ�ʱ��%Type;
    v_�շ�Ա       ������ü�¼.����Ա����%Type;
  
    n_ȱʡ�����id     Number(18);
    v_����ֵ           Varchar2(100);
    n_Ʊ���ܽ��       ������ü�¼.���ʽ��%Type;
    n_����ܶ�         ������ü�¼.���ʽ��%Type;
    n_�û�id           ��Ա��.Id%Type;
    v_����Ա���       ��Ա��.���%Type;
    v_����Ա����       ��Ա��.����%Type;
    v_Temp             Varchar2(32767);
    v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
    v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
    n_����             ���ս����¼.����%Type;
    v_���ջ�������     �������.���ջ�������%Type;
    n_ҽ�����         ������ü�¼.ҽ�����%Type;
    n_�Һ�id           ������ü�¼.�Һ�id%Type;
    v_��������         ���ղ���.����%Type;
    v_��������         Varchar2(20);
    v_������ұ���     ���ű�.����%Type;
    v_�����������     ���ű�.����%Type;
    v_������         Varchar2(50);
    v_����ids          Varchar2(32767);
    v_ҽ����           �����ʻ�.ҽ����%Type;
    l_����id           t_NumList := t_NumList();
    v_�汾��           Varchar2(30);
    n_����Ʊ��id       ����Ʊ��ʹ�ü�¼.Id%Type;
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_Ӧ�ó��� := Nvl(j_Json.Get_Number('occasion'), 0);
    n_����id   := j_Json.Get_Number('balance_id');
    --n_����id     := Nvl(j_Json.Get_Number('writeoff_id'), 0);
    n_����Ʊ��id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
  
    If Nvl(n_Ӧ�ó���, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '��Ч��Ӧ�ó���';
      Return;
    End If;
  
    Select Nvl(Max(����ֵ), 'V2.0.3')
    Into v_�汾��
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '֧�ְ汾';
  
    b_Einvoice_Request_Test.Get_Identity(n_�û�id, v_����Ա���, v_����Ա����);
  
    --v_ҵ���ʶ:01  סԺ,02  ����,03  ����,04  ����,05  �������,06  �Һ�,07  סԺԤ����,08  ���Ԥ����
    Select Decode(n_Ӧ�ó���, 1, '02', 2, '07', 3, '01', 4, '06', 5, '02', '02') Into v_ҵ���ʶ From Dual;
  
    n_Ʊ���ܽ��   := 0;
    d_ҵ����ʱ�� := Null;
    v_����ids      := Null;
    c_��ϸ         := Null;
    v_��ϸ         := Null;
    For c_�շ�ϸĿ In (Select Min(a.Id) As ����id, a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, Max(a.���㵥λ) As ���㵥λ,
                          Sum(a.��׼����) As �۸�, Avg(Nvl(a.����, 1) * Nvl(a.����, 0)) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��,
                          Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.ʵ�ս��) - Sum(a.ͳ����) As �Էѽ��,
                          Max(t.����) As ҽ����Ŀ����, Max(t.����) As ҽ����Ŀ����, Max(t.ͳ��ȶ�) As ҽ����������, Max(a.ժҪ) As ��ע,
                          Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����) As ����,
                          Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��,
                          Max(a.���ʽ) As ���ʽ����, Max(a.�վݷ�Ŀ) As �վݷ�Ŀ, Max(c.����) As �վݷ�Ŀ����, Max(a.ҽ�����) As ҽ�����,
                          Max(B1.Id) As �Һ�id, Max(d.����) As ������, Max(d.���) As �������, Max(b.����) As ��Ŀ����, Max(b.����) As ��Ŀ����,
                          Max(b.���) As ���, Max(q.ҩƷ����) As ҩƷ����
                   From ������ü�¼ A, ���˹Һż�¼ B1, �շ���ĿĿ¼ B, �վݷ�Ŀ C, �շ���� D, ҩƷ��� M, ҩƷ���� Q, ������ĿĿ¼ J, ����֧������ T, ֧�������� S
                   Where a.No = B1.No And a.No In (Select Distinct NO From ������ü�¼ Where ����id = n_����id) And a.��¼���� = 4 And
                         a.�շ���� = d.����(+) And a.�շ�ϸĿid = b.Id And a.�վݷ�Ŀ = c.����(+) And a.�շ�ϸĿid = m.ҩƷid(+) And
                         m.ҩ��id = q.ҩ��id(+) And q.ҩ��id = j.Id(+) And a.���մ���id = t.Id(+) And t.����(+) = 1 And
                         a.���մ���id = s.���մ���id(+)
                   Group By a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, c.����, c.����, j.����, j.����
                   Order By NO, ���) Loop
      If v_�������� Is Null Then
        v_�������� := c_�շ�ϸĿ.����;
        v_�����Ա� := c_�շ�ϸĿ.�Ա�;
        v_�������� := c_�շ�ϸĿ.����;
        n_����id   := c_�շ�ϸĿ.����id;
      End If;
      If d_ҵ����ʱ�� Is Null And Nvl(c_�շ�ϸĿ.��¼״̬, 0) = 1 Then
        --ȡԭʼҵ����ʱ��
        d_ҵ����ʱ�� := c_�շ�ϸĿ.�Ǽ�ʱ��;
        v_�շ�Ա       := c_�շ�ϸĿ.����Ա����;
      End If;
      If v_ҽ�Ƹ��ʽ���� Is Null Then
        v_ҽ�Ƹ��ʽ���� := c_�շ�ϸĿ.���ʽ����;
      End If;
      If Nvl(n_ҽ�����, 0) = 0 Then
        n_ҽ����� := c_�շ�ϸĿ.ҽ�����;
      End If;
      If Nvl(n_�Һ�id, 0) = 0 Then
        n_�Һ�id := c_�շ�ϸĿ.�Һ�id;
      End If;
    
      If Instr(Nvl(v_����ids, '') || ',', ',' || c_�շ�ϸĿ.����id || ',') = 0 Then
        l_����id.Extend;
        l_����id(l_����id.Count) := c_�շ�ϸĿ.����id;
      End If;
    
      --listDetailNo  ��ϸ��ˮ��  String  60  ��  ��ϸ��ˮ��
      v_Temp := '{' || '"listDetailNo":"' || b_Einvoice_Request_Test.Zljsonstr(LPad(c_�շ�ϸĿ.����id, 20, '0')) || '"';
      --chargeCode  �շ���Ŀ����  String  50  ��  ��дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���,�磺��λ�ѡ�����
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�վݷ�Ŀ����) || '"';
      --chargeName  �շ���Ŀ����  String  100  ��
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�վݷ�Ŀ) || '"';
      --prescribeCode  ��������  String  60  ��
      v_Temp := v_Temp || ',' || '"prescribeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.No) || '"';
      --listTypeCode  ҩƷ������  String  50  ��  ��ҩƷ�������01��������д
      v_Temp := v_Temp || ',' || '"listTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.������) || '"';
      --listTypeName  ҩƷ�������  String  50  ��  ��ҩƷ�������ƣ��������࿹��Ⱦҩ��
      v_Temp := v_Temp || ',' || '"listTypeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�������) || '"';
      --code  ����  String  50  ��  ��ҩƷ���룬������д
      v_Temp := v_Temp || ',' || '"code":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��Ŀ����) || '"';
      --name  ҩƷ����  String  50  ��  ��ҩƷ���ƣ��������Ƶ�
      v_Temp := v_Temp || ',' || '"name":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��Ŀ����) || '"';
      --form  ����  String  50  ��
      v_Temp := v_Temp || ',' || '"form":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҩƷ����) || '"';
      --specification  ���  String  50  ��
      v_Temp := v_Temp || ',' || '"specification":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���) || '"';
      --unit  ������λ   String  20  ��
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���㵥λ) || '"';
      --std  ����  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�۸�, 1);
      --number  ����  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.����, 1);
      --amt  ���  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ʵ�ս��, 1);
      --selfAmt  �Էѽ��  Number  14,6  ��  ���޽���д0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�Էѽ��, 1);
      --receivableAmt  Ӧ�շ���  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"receivableAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.Ӧ�ս��, 1);
      --medicalCareType  ҽ��ҩƷ����  String  1  ��  1�����Ը�/��
      --          2�����Ը�/��
      --          3��ȫ�Ը�/��
      v_Temp := v_Temp || ',' || '"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����Ŀ����) || '"';
      --medCareItemType  ҽ����Ŀ����  String  100  ��
      v_Temp := v_Temp || ',' || '"medCareItemType":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����Ŀ����) || '"';
      --medReimburseRate  ҽ����������  Number  3,2  ��
      v_Temp := v_Temp || ',' || '"medReimburseRate":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����������, 1);
      --remark  ��ע  String  200  ��
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��ע) || '"';
      --sortNo  ���  Integer  ����  ��  ���
      v_Temp := v_Temp || ',' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���, 1);
      --chrgtype  ��������  String  50  ��
      v_Temp := v_Temp || ',' || '"chrgtype":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��������) || '"}';
    
      If Length(Nvl(v_��ϸ, '') || v_Temp) > 32700 Then
        If c_��ϸ Is Null Then
          c_��ϸ := To_Clob(v_��ϸ);
        Else
          c_��ϸ := c_��ϸ || To_Clob(',' || v_��ϸ);
        End If;
        v_��ϸ := Null;
      End If;
    
      If v_��ϸ Is Null Then
        v_��ϸ := v_Temp;
      Else
        v_��ϸ := v_��ϸ || ',' || v_Temp;
      End If;
    End Loop;
  
    If v_��ϸ Is Not Null And c_��ϸ Is Not Null Then
      --listDetail  �嵥��Ŀ��ϸ  String  ����  ��  ���A-2,JSON��ʽ�б�
      c_��ϸ := c_��ϸ || ',' || To_Clob(v_��ϸ);
      c_��ϸ := To_Clob(',"listDetail":[') || c_��ϸ || To_Clob(']');
    
      v_��ϸ := Null;
    Elsif v_��ϸ Is Not Null Then
      v_��ϸ := ',"listDetail":[' || v_��ϸ || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --������ϸ
    v_������ϸ := Null;
    c_������ϸ := Null;
    For c_����ͳ�� In (Select Rownum As ���, �վݷ�Ŀ����, �վݷ�Ŀ����, ����, ���㵥λ, Round(����, 2) As ����, Round(���ʽ��, 2) As ���ʽ��,
                          Round(�Էѽ��, 2) As �Էѽ��, ��ע, ���ʽ�� - Round(���ʽ��, 2) As ����
                   From (Select /*+cardinality(b,10)*/
                           c.���� As �վݷ�Ŀ����, a.�վݷ�Ŀ As �վݷ�Ŀ����, 1 As ����, '' As ���㵥λ, Sum(a.���ʽ��) As ����, a.�վݷ�Ŀ,
                           Sum(a.���ʽ��) As ���ʽ��, Sum(a.���ʽ��) - Sum(a.ͳ����) As �Էѽ��, '' As ��ע
                          From ������ü�¼ A, Table(l_����id) B, �վݷ�Ŀ C
                          Where a.����id = b.Column_Value And a.�վݷ�Ŀ = c.����(+)
                          Group By c.����, a.�վݷ�Ŀ)) Loop
      --sortNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���շ���Ŀ���ֵ����1�����β������ظ�
      v_Temp := '{' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���, 1);
      --chargeCode  �շ���Ŀ����  String  50  ��  ��дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�վݷ�Ŀ����) || '"';
      --chargeName  �շ���Ŀ����  String  100  ��  ��дҵ��ϵͳ�ڲ���Ŀ����
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�վݷ�Ŀ����) || '"';
      --unit  ������λ  String  20  ��
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���㵥λ) || '"';
      --std  �շѱ�׼  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.����, 1);
      --number  ����  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.����, 1);
      --amt  ���  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���ʽ��, 1);
      --selfAmt  �Էѽ��  Number  14,2  ��  ���޽���д0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�Էѽ��, 1);
      --remark  ��ע  String  200  ��
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.��ע) || '"}';
    
      If Length(Nvl(v_������ϸ, '') || v_Temp) > 32700 Then
        If c_������ϸ Is Null Then
          c_������ϸ := To_Clob(v_������ϸ);
        Else
          c_������ϸ := c_������ϸ || To_Clob(',' || v_������ϸ);
        End If;
        v_������ϸ := Null;
      End If;
    
      If v_������ϸ Is Null Then
        v_������ϸ := v_Temp;
      Else
        v_������ϸ := v_������ϸ || ',' || v_Temp;
      End If;
    
      n_Ʊ���ܽ�� := Nvl(n_Ʊ���ܽ��, 0) + Nvl(c_����ͳ��.���ʽ��, 0);
      n_����ܶ�   := Nvl(n_����ܶ�, 0) + Nvl(c_����ͳ��.����, 0);
    End Loop;
    Totalmoney_Out := n_Ʊ���ܽ��;
    If v_������ϸ Is Not Null And c_������ϸ Is Not Null Then
      c_������ϸ := c_������ϸ || ',' || To_Clob(v_������ϸ);
      --chargeDetail �շ���Ŀ��ϸ  String  ����  ��  ���A-1,JSON��ʽ�б�
      c_������ϸ := To_Clob(',"chargeDetail":[') || c_������ϸ || To_Clob(']');
      v_������ϸ := Null;
    Elsif v_������ϸ Is Not Null Then
      v_������ϸ := ',"chargeDetail":[' || v_������ϸ || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --Ʊ����Ϣ
    --Select Count(1) Into n_���ϴ��� From ����Ʊ��ʹ�ü�¼ Where Ʊ�� = n_Ӧ�ó��� And ����id = n_����id;
    --lpad(����Ʊ�����ϴ���,5) & Lpad(ԭ����ID,20)
    --����ID||_||����Ʊ��ID
    v_ҵ����ˮ�� := n_����id || '_' || Nvl(n_����Ʊ��id, 0);
    v_��Ʊ��     := b_Einvoice_Request_Test.Get_Einvoice_Node(v_����Ա���, v_����Ա����, n_�û�id);
  
    v_Ʊ����Ϣ := '"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ����ˮ��) || '"'; --ҵ����ˮ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"busType":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ���ʶ) || '"'; --ҵ���ʶ
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --��������
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"busDateTime":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(d_ҵ����ʱ��, 'yyyymmddHH24miss') || '000') || '"'; --ҵ����ʱ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_��Ʊ��) || '"'; --��Ʊ�����:ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(v_�շ�Ա) || '"'; --�շ�Ա
  
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_����Ա����) || '"'; --Ʊ�ݱ�����
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"totalAmt":' || b_Einvoice_Request_Test.Zljsonstr(n_Ʊ���ܽ��, 1); --��Ʊ�ܽ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"remark":"' || '' || '"'; --��ע
    -------------------------------------------------------------------------------------------
  
    --ȡ�ɷ���Ϣ
    v_�ɷ� := Null;
    For c_�ɷ� In (Select Max(Decode(��Ϣ��, '������', ��Ϣֵ, '֧��������', ��Ϣֵ, '')) As ֧��������,
                        Max(Decode(��Ϣ��, 'ҽ��֧��������', ��Ϣֵ, 'ҽ��������', ��Ϣֵ, '')) As ҽ��֧��������,
                        Max(Decode(Upper(��Ϣ��), '֧�������ں�USERID', ��Ϣֵ, '')) As ֧�������ں�userid,
                        Max(Decode(Upper(��Ϣ��), '֧����С����USERID', ��Ϣֵ, '')) As ֧����С����userid,
                        Max(Decode(Upper(��Ϣ��), '΢�Ź��ں�OPENID', ��Ϣֵ, '')) As ΢�Ź��ں�openid,
                        Max(Decode(Upper(��Ϣ��), '΢��С����OPENID', ��Ϣֵ, '')) As ΢��С����openid
                 From (Select ��Ϣ��, ��Ϣֵ
                        From ������Ϣ�ӱ�
                        Where ����id = n_����id And ��Ϣ�� In ('֧�������ں�USERID', '֧����С����USERID', '΢�Ź��ں�OPENID', '΢��С����OPENID')
                        Union All
                        Select ������Ŀ, ��������
                        From �������㽻��
                        Where ����id In (Select ID From ����Ԥ����¼ Where ����id = n_����id) And ������Ŀ Like '%������')) Loop
      v_�ɷ� := ',"alipayCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.֧�������ں�userid) || '"'; --����֧�����˻�
      v_�ɷ� := v_�ɷ� || ',"weChatOrderNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.֧��������) || '"'; --΢��֧��������
      If v_�汾�� = 'V3.1.0' Then
        --�ð汾���д˽ӵ�
        v_�ɷ� := v_�ɷ� || ',"weChatMedTransNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.ҽ��֧��������) || '"'; --΢��ҽ��֧��������
      End If;
    
      If c_�ɷ�.΢�Ź��ں�openid Is Not Null Then
        v_�ɷ� := v_�ɷ� || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.΢�Ź��ں�openid) || '"'; --΢�Ź��ںŻ�С�����û�ID
      Else
        v_�ɷ� := v_�ɷ� || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.΢��С����openid) || '"'; --΢�Ź��ںŻ�С�����û�ID
      End If;
      Exit;
    End Loop;
  
    -------------------------------------------------------------------------------------------
    --ȡ֪ͨ��Ϣ
    Select To_Number(Max(����ֵ))
    Into n_ȱʡ�����id
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = 'ȱʡ�����ID';
    v_֪ͨ := Null;
    For c_֪ͨ In (Select Max(a.����id) As ����id, Max(a.����) As ����, Max(a.�ֻ���) As �ֻ���, Max(a.Email) As Email, Max(1) As �ɿ�����,
                        Max(a.���֤��) As ���֤��, Max(m.����) As �����, Max(m.����) As ����, Max(a.�����) As �����
                 From ������Ϣ A,
                      (
                        
                        Select ����id, ����, ����, ����
                        From (Select b.����id, c.����, c.����, b.����, Decode(b.�����id, n_ȱʡ�����id, 2, c.ȱʡ��־) As ȱʡ��־
                                From ����ҽ�ƿ���Ϣ B, ҽ�ƿ���� C
                                Where b.�����id = c.Id And b.����id = n_����id
                                Order By ȱʡ��־)
                        Where Rownum < 2) M
                 Where a.����id = m.����id(+)) Loop
    
      v_֪ͨ := ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.�ֻ���) || '"'; --�����ֻ�����
      v_֪ͨ := v_֪ͨ || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.Email) || '"'; --���������ַ
      If v_�汾�� = 'V3.1.0' Then
        --�ð汾���д˽ӵ�
        v_֪ͨ := v_֪ͨ || ',"payerType":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.�ɿ�����) || '"'; --����������
      End If;
      v_֪ͨ := v_֪ͨ || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.���֤��) || '"'; --ͳһ������ô���
    
      If c_֪ͨ.����� Is Not Null Then
        v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.�����) || '"'; --������
        v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.����) || '"'; --����
      Elsif c_֪ͨ.���֤�� Is Not Null Then
        Select Nvl(Max(����ֵ), '99998')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '���֤�������ͱ��';
        v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
        v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.���֤��) || '"'; --����
      Else
        --û��һ�ſ����̶�һ�ֿ����
        Select Nvl(Max(����ֵ), '99999')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ������';
        v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
        Select Nvl(Max(����ֵ), '-')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ���';
        v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --����
      End If;
      If Nvl(n_�����, 0) = 0 Then
        n_����� := c_֪ͨ.�����;
      
      End If;
    
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --������Ϣ
    Select Max(����) Into v_Temp From zlRegInfo Where ��Ŀ = 'ҽ�ƻ�������';
  
    --����:1-�շ�;2-���㣨����סԺ���㡢����������㣩��3-Ԥ��
    Select Max(a.����), Max(b.���ջ�������), Max(Nvl(a.��������, c.����))
    Into n_����, v_���ջ�������, v_��������
    From ���ս����¼ A, ������� B, ���ղ��� C
    Where a.���� = b.��� And a.����id = c.Id(+) And a.��¼id = n_����id And a.���� = Decode(n_Ӧ�ó���, 2, 3, 3, 2, 1);
  
    Select Max(����) Into v_ҽ�Ƹ��ʽ���� From ҽ�Ƹ��ʽ Where ���� = v_ҽ�Ƹ��ʽ����;
    If Nvl(n_����, 0) <> 0 Then
      Select Max(ҽ����) Into v_ҽ���� From �����ʻ� Where ����id = n_����id And ���� = n_����;
    End If;
  
    v_������ := Null;
    If Nvl(n_�Һ�id, 0) <> 0 Then
      Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')), Max(b.����), Max(b.����), Max(a.No)
      Into v_��������, v_������ұ���, v_�����������, v_������
      From ���˹Һż�¼ A, ���ű� B
      Where a.ִ�в���id = b.Id And a.Id = n_�Һ�id;
    End If;
  
    If v_�������� Is Null And Nvl(n_����, 0) <> 0 Then
      Select Max(��������)
      Into v_��������
      From (Select Distinct a.���� As ��������
             From ���ղ��� A, ������׼��Ŀ B
             Where a.���� = n_���� And a.Id = b.����id And
                   b.�շ�ϸĿid In (Select Distinct �շ�ϸĿid From ������ü�¼ Where ����id = n_����id)
             Union All
             Select Distinct a.���� As ��������
             From ���ղ��� A, ������׼��Ŀ B
             Where a.���� = n_���� And a.Id = b.����id And
                   b.���� In (Select Distinct ���մ���id From ������ü�¼ Where ����id = n_����id))
      Where Rownum < 2;
    End If;
  
    v_������Ϣ := ',"medicalInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"'; --ҽ�ƻ�������
    v_������Ϣ := v_������Ϣ || ',"medCareInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_���ջ�������) || '"'; --ҽ��������Ψһ����
    v_������Ϣ := v_������Ϣ || ',"medCareTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ�Ƹ��ʽ����) || '"'; --ҽ�����ͱ���
    v_������Ϣ := v_������Ϣ || ',"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ�Ƹ��ʽ����) || '"'; --ȡֵ��Χ����ְ������ҽ�Ʊ��ա�����������ҽ�Ʊ��գ�����������ҽ�Ʊ��ա�����ũ�����ҽ�Ʊ��գ�������ҽ�Ʊ��ա���ҽ����
    v_������Ϣ := v_������Ϣ || ',"medicalInsuranceID":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ����) || '"';
  
    v_������Ϣ := v_������Ϣ || ',"consultationDate":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --���߾�ҽʱ��
    v_������Ϣ := v_������Ϣ || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(v_�����������) || '"'; --�������
    v_������Ϣ := v_������Ϣ || ',"patientCategory":"' || b_Einvoice_Request_Test.Zljsonstr(v_������ұ���) || '"'; --������ұ���
    --patientNo  ���߾�����  String  20  ��  ����ÿ�ξ���һ�ξ����ɵ�һ���µı�š������ߵǼǺţ�
    v_������Ϣ := v_������Ϣ || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_������) || '"';
    v_������Ϣ := v_������Ϣ || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(n_����id) || '"'; --������ҵ��ϵͳ�е�Ψһ��ʶID���������֤���롣
    v_������Ϣ := v_������Ϣ || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(v_�����Ա�) || '"'; --�Ա�
    v_������Ϣ := v_������Ϣ || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --����
    -------------------------------------------------------------------------------------------
    --������Ϣ
    v_���� := Null;
    For c_���� In (Select �ֽ�Ԥ��, ֧ƱԤ��, ת��Ԥ��, �����ʻ�֧��, ҽ��ͳ�����֧��, ����ҽ��֧��, �����ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, �ֽ�֧��, 0) As �ֽ��˿�,
                        Decode(Sign(�ֽ�֧��), -1, ֧Ʊ֧��, 0) As ֧Ʊ�˿�, Decode(Sign(�ֽ�֧��), -1, ת��֧��, 0) As ת���˿�,
                        Decode(Sign(�ֽ�֧��), -1, 0, �ֽ�֧��) As �ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, 0, ֧Ʊ֧��) As ֧Ʊ֧��,
                        Decode(Sign(�ֽ�֧��), -1, 0, ת��֧��) As ת��֧��,
                        Nvl(�����ʻ�֧��, 0) + Nvl(ҽ��ͳ�����֧��, 0) + Nvl(����ҽ��֧��, 0) As �����ܶ�,
                        Nvl(�����ܶ�, 0) - Nvl(�����ʻ�֧��, 0) - Nvl(ҽ��ͳ�����֧��, 0) - Nvl(����ҽ��֧��, 0) As �Էѽ��, �����ܶ�, ҽ���������,
                        0 As �����ʻ����
                 From (Select /*+cardinality(b,10)*/
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '�ֽ�', 1, 0), 0) * a.��Ԥ��) As �ֽ�Ԥ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0), 0) * a.��Ԥ��) As ֧ƱԤ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 0, '�ֽ�', 0, 1), 0) * a.��Ԥ��) As ת��Ԥ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '�����˻�֧��', 1, 0)) * a.��Ԥ��) As �����ʻ�֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, 'ҽ��ͳ�����֧��', 1, 0)) * a.��Ԥ��) As ҽ��ͳ�����֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 1, 0)) * a.��Ԥ��) As ����ҽ��֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 0, '�����˻�֧��', 0, 'ҽ��ͳ�����֧��', 0, 1)) *
                              a.��Ԥ��) As �����ֽ�֧��,
                         Max(Decode(Mod(a.��¼����, 10), 1, 0,
                                     Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', �������, '�����˻�֧��', �������, 'ҽ��ͳ�����֧��', �������, ''))) As ҽ���������,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '�ֽ�', 1, 0)) * a.��Ԥ��) As �ֽ�֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0)) * a.��Ԥ��) As ֧Ʊ֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '�ֽ�', 0, '֧Ʊ', 0, 1) * a.��Ԥ��)) As ת��֧��,
                         Sum(��Ԥ��) As �����ܶ�
                        From ����Ԥ����¼ A, Table(l_����id) B, ��Ʊ������� C
                        Where a.����id = b.Column_Value And a.���㷽ʽ = c.���㷽ʽ(+)))
    
     Loop
      --accountPay  �����˻�֧��  Number  14,2  ��  �����߹涨�ø����˻�֧���α��˵�ҽ�Ʒ��ã�������ҽ�Ʊ���Ŀ¼��Χ�ں�Ŀ¼��Χ��ķ��ã���
      --          ���޽���д0
      v_���� := ',"accountPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.�����ʻ�֧��, 0), 1);
      --fundPay  ҽ��ͳ�����֧��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а��涨�ɻ���ҽ�Ʊ���ͳ�����֧���Ľ�
      --          ���޽���д0
      v_���� := v_���� || ',"fundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.ҽ��ͳ�����֧��, 0), 1);
      --otherfundPay  ����ҽ��֧��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а��涨�ɴ󲡱��ա�ҽ�ƾ���������Աҽ�Ʋ��������䡢��ҵ����Ȼ�����ʽ�֧���Ľ�
      v_���� := v_���� || ',"otherfundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.����ҽ��֧��, 0), 1);
      --ownPay  �Էѽ��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а����йع涨�����ڻ���ҽ�Ʊ���Ŀ¼��Χ��ȫ���ɸ���֧���ķ��ã�
      --          ���޽���д0
      v_���� := v_���� || ',"ownPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�Էѽ��, 1);
      --selfConceitedAmt  �����Ը�  Number  14,2  ��  ҽ�������𸶱�׼�ڸ���֧�����ã�
      --          ���޽���д0
      v_���� := v_���� || ',"selfConceitedAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfPayAmt  �����Ը�  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ������ɸ��˸��������ڻ���ҽ�Ʊ���Ŀ¼��Χ���Ը����ֵĽ���չ�����֡����顢���յȴ�����ѷ�ʽ���ɻ��߶���ѵķ��á�����Ϊ��������˰��ҽ��ר��ӿ۳��ţ�Ϣ�����޽���д0
      v_���� := v_���� || ',"selfPayAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfCashPay  �����ֽ�֧��  Number  14,2  ��  ����ͨ���ֽ����п���΢�š�֧����������֧���Ľ�
      --          ���޽���д0
      v_���� := v_���� || ',"selfCashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ֽ�֧��, 1);
      --�Ժ�����漰��Ԥ��,�ݱ���
      --cashPay  �ֽ�Ԥ������  Number  14,2  ��
      --v_���� := v_���� || ',"cashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ�Ԥ��, 1);
      --chequePay  ֧ƱԤ������  Number  14,2  ��
      --v_���� := v_���� || ',"chequePay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧ƱԤ��, 1);
      --transferAccountPay  ת��Ԥ������  Number  14,2  ��
      --v_���� := v_���� || ',"transferAccountPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת��Ԥ��, 1);
      --cashRecharge  �������(�ֽ�)  Number  14,2  ��
      --v_���� := v_���� || ',"cashRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ�֧��, 1);
      --chequeRecharge  �������(֧Ʊ)  Number  14,2  ��
      --v_���� := v_���� || ',"chequeRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧Ʊ֧��, 1);
      --transferRecharge  ������ת�ˣ�  Number  14,2  ��
      --v_���� := v_���� || ',"transferRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת��֧��, 1);
      --cashRefund  �˻����(�ֽ�)  Number  14,2  ��
      --v_���� := v_���� || ',"cashRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ��˿�, 1);
      --chequeRefund  �˽����(֧Ʊ)  Number  14,2  ��
      --v_���� := v_���� || ',"chequeRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧Ʊ�˿�, 1);
      --transferRefund  �˽����(ת��)  Number  14,2  ��
      --v_���� := v_���� || ',"transferRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת���˿�, 1);
      --ownAcBalance  �����˻����  Number  14,2  ��
      v_���� := v_���� || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ʻ����, 1);
      --reimbursementAmt  �����ܽ��  Number  14,2  ��  ҽ������󷵻ص��ܽ��
      v_���� := v_���� || ',"reimbursementAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ܶ�, 1);
      --balancedNumber  �����  String  100  ��  ҽ����������ɵĺ���/����Ψһֵ
      v_���� := v_���� || ',"balancedNumber":"' || b_Einvoice_Request_Test.Zljsonstr(c_����.ҽ���������) || '"';
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --��������
    v_�ɷ����� := Null;
    For c_���� In (Select /*+cardinality(b,10)*/
                  Nvl(c.��������, Nvl(d.��������, '-')) As ��������, Sum(��Ԥ��) As �����ܶ�
                 From ����Ԥ����¼ A, Table(l_����id) B, �շ��������� C,
                      (Select ���㷽ʽ, �������� From �շ��������� D Where �����id Is Null) D
                 Where a.����id = b.Column_Value And a.�����id = c.�����id(+) And a.���㷽ʽ = c.���㷽ʽ(+) And a.���㷽ʽ = d.���㷽ʽ(+)
                 Group By Nvl(c.��������, Nvl(d.��������, '-'))
                 Order By ��������)
    
     Loop
      --payChannelCode  ������������  String  10  ��
      If v_�ɷ����� Is Null Then
        v_�ɷ����� := Nvl(v_�ɷ�����, '') || '{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.��������, 0)) || '"';
      Else
        v_�ɷ����� := v_�ɷ����� || ',{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.��������, 0)) || '"';
      End If;
      --payChannelValue  �����������  Number  14,2  ��
      v_�ɷ����� := v_�ɷ����� || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.�����ܶ�, 0), 1) || '}';
    End Loop;
  
    If v_�ɷ����� Is Not Null Then
      --payChannelDetail  ���������б�  String  ����  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö����磺��¼4���������б�
      --        ���A-5,JSON��ʽ�б�
      v_�ɷ����� := ',"payChannelDetail":[' || v_�ɷ����� || ']';
    Else
      v_�ɷ����� := ',"payChannelDetail":[]';
    End If;
  
    -------------------------------------------------------------------------------------------
    --����ҽ����Ϣ
    v_����ҽ����Ϣ := Null;
    --otherMedicalList  ����ҽ����Ϣ�б�  String  ����  ��  ��д����δ֪ҽ����Ϣ���ڵ���Ʊ����������ƴ�ӷ�ʽ��ʾ��
    --            ���A-4,JSON��ʽ�б�
    --  infoNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���������ֵ����1�����β������ظ�
    --  infoName  ҽ����Ϣ����  String  100  ��  ����ñ������ͱ��룬�ɲο���¼7ҽ�����������б�
    --  infoValue  ҽ����Ϣֵ  String  100  ��  ����ñ������
    --  infoOther  ҽ��������Ϣ  String  100  ��  ��ҽ������������
  
    -------------------------------------------------------------------------------------------
    --������չ��Ϣ
    v_������չ��Ϣ := Null;
    --otherInfo  ������չ��Ϣ�б�  String  ����  ��  ��д��Ϣ��Ҫ�ڵ���Ʊ���ϵ�����ʾ��������չ��Ϣ��δ֪��Ϣ��
    --          ���A-3,JSON��ʽ�б�
    --  infoNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���������ֵ����1�����β������ظ�
    --  infoName  ��չ��Ϣ����  String  100  ��
    --  infoValue  ��չ��Ϣֵ  String  500  ��
  
    c_������Ϣ := To_Clob('{' || v_Ʊ����Ϣ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_�ɷ�);
    c_������Ϣ := c_������Ϣ || To_Clob(v_֪ͨ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_������Ϣ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_����);
    c_������Ϣ := c_������Ϣ || To_Clob(v_�ɷ�����);
  
    If v_������չ��Ϣ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_������չ��Ϣ);
    End If;
    If v_����ҽ����Ϣ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_����ҽ����Ϣ);
    End If;
    --  eBillRelateNo  ҵ��Ʊ�ݹ�����  String  32  ��  ��һ��ҵ��������Ҫ����N�ŵ���Ʊ�ݣ���N�ŵ���Ʊ��Ӧ��ֵ����һ�£����ں��ڹ�����ѯ
    --isArrears  �Ƿ����ͨ  String  1  ��  0-��1-�ǣ���Ƿ���������ҽԺҵ��Ҫ���Ʊ���Ƿ����ͨ��
    --arrearsReason  ������ͨԭ��  String  200  ��  isArrears=0����д������ͨ��ԭ��
  
    c_������Ϣ := c_������Ϣ || To_Clob(',"eBillRelateNo":"","isArrears":"1","arrearsReason":""');
    If v_������ϸ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_������ϸ);
    Else
      c_������Ϣ := c_������Ϣ || c_������ϸ;
    End If;
  
    If v_��ϸ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_��ϸ);
    Else
      c_������Ϣ := c_������Ϣ || c_��ϸ;
    End If;
    c_������Ϣ  := c_������Ϣ || To_Clob('}');
    Reqdata_Out := c_������Ϣ;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Registerdata_Create;

  Procedure Get_Zybalancedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    --
    ---------------------------------------------------------------------------
    --����:��ȡסԺ���ʿ�Ʊ����
    --���:
    --    Json_In,��ʽ����
    --  input
    --    occasion N 1  Ӧ�ó���:1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨���̶���3
    --    einvoice_id  N,1 ��ǰ����Ʊ��ID
    --    balance_id N 1  ����ID  occasion=2ʱ��Ԥ��id;occasion<>2��ʾ����id
    --    writeoff_id  N 1  ����ID  occasion=2ʱ������Ԥ��id;occasion<>2��ʾ����id
    --����:
    --  ReqData_Out-���ص�ҵ����������
    --  Totalmoney_Out-Ʊ���ܶ�
    --  Code_Out-��ȡ�Ƿ�ɹ���0-ʧ�ܣ�1-�ɹ�
    --  Message_Out ������Ϣ
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_Ӧ�ó��� Number(2);
    n_����id   ����Ԥ����¼.����id%Type;
  
    v_ҵ����ˮ�� Varchar2(50);
  
    v_��Ʊ��       Varchar2(100);
    v_�ɷ�         Varchar2(32767);
    v_Ʊ����Ϣ     Varchar2(32767);
    v_������Ϣ     Varchar2(32767);
    v_֪ͨ         Varchar2(32767);
    v_�ɷ�����     Varchar2(32767);
    v_����         Varchar2(32767);
    v_������չ��Ϣ Varchar2(32767);
    v_����ҽ����Ϣ Varchar2(32767);
    c_��ϸ         Clob;
    v_��ϸ         Varchar2(32767);
    c_������ϸ     Clob;
    v_������ϸ     Varchar2(32767);
    c_������Ϣ     Clob; --���շ��صĽ�����Ϣ��
  
    c_Ԥ�� Clob;
    v_Ԥ�� Varchar2(32767);
  
    n_����id ����Ԥ����¼.����id%Type;
  
    n_ȱʡ�����id     Number(18);
    v_����ֵ           Varchar2(100);
    n_Ʊ���ܽ��       ������ü�¼.���ʽ��%Type;
    n_����ܶ�         ������ü�¼.���ʽ��%Type;
    n_�û�id           ��Ա��.Id%Type;
    v_����Ա���       ��Ա��.���%Type;
    v_����Ա����       ��Ա��.����%Type;
    v_Temp             Varchar2(32767);
    n_����             ���ս����¼.����%Type;
    v_���ջ�������     �������.���ջ�������%Type;
    v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
    v_��������         ���ղ���.����%Type;
    v_ҽ����           �����ʻ�.ҽ����%Type;
    v_�汾��           Varchar2(30);
    v_סԺ����         Varchar2(4000);
    n_����Ʊ��id       ����Ʊ��ʹ�ü�¼.Id%Type;
    Cursor c_Balance_Record Is
      Select a.No, a.�շ�ʱ��, a.��������, a.����Ա���, a.����Ա����, a.����id, a.��ҳid,
             Decode(Nvl(a.����id, 0), 0, a.ԭ��, Nvl(b.����, c.����)) As ����, Nvl(b.�Ա�, c.�Ա�) As �Ա�, Nvl(b.����, c.����) As ����, c.�����,
             Nvl(b.סԺ��, c.סԺ��) As סԺ��, a.��ʼ����, a.��������, a.��ע, a.���ʽ��, Decode(Nvl(a.����id, 0), 0, q.�����ʼ�, c.Email) As Email,
             q.��ϵ��, Decode(Nvl(a.����id, 0), 0, q.������ô���, c.���֤��) As ���֤��,
             Decode(Nvl(a.����id, 0), 0, Nvl(q.�绰, To_Char(j.�ƶ��绰)), c.�ֻ���) As �ֻ���,
             Decode(Nvl(a.����id, 0), 0, 2, 1) As �ɿ�����, Decode(Nvl(a.��������, 0), 1, '02', '01') As ҵ���ʶ, b.��Ժ����, b.��Ժ����,
             m.���� As ��Ժ���ұ���, m.���� As ��Ժ��������, p.���� As ��Ժ���ұ���, p.���� As ��Ժ��������, b.��Ժ���� As ����, t.���� As ��������,
             Nvl(b.������, b.סԺ��) As ������, Nvl(b.ҽ�Ƹ��ʽ, c.ҽ�Ƹ��ʽ) As ҽ�Ƹ��ʽ, Nvl(b.��Ժ����, Sysdate) - b.��Ժ���� As סԺ����
      From ���˽��ʼ�¼ A, ������ҳ B, ������Ϣ C, ��Լ��λ Q, ��Ա�� J, ���ű� M, ���ű� P, ���ű� T
      Where a.Id = n_����id And a.����id = b.����id(+) And a.��ҳid = b.��ҳid(+) And a.����id = c.����id(+) And a.ԭ�� = q.����(+) And
            b.��Ժ����id = m.Id(+) And b.��Ժ����id = p.Id(+) And b.��ǰ����id = t.Id(+)
           
            And q.��ϵ�� = j.����(+);
    r_Balance_Record c_Balance_Record%RowType;
  
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_Ӧ�ó��� := Nvl(j_Json.Get_Number('occasion'), 0);
    n_����id   := j_Json.Get_Number('balance_id');
    --n_����id     := Nvl(j_Json.Get_Number('writeoff_id'), 0);
    n_����Ʊ��id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
  
    If Nvl(n_Ӧ�ó���, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '��Ч��Ӧ�ó���';
      Return;
    End If;
  
    Select Nvl(Max(����ֵ), 'V2.0.3')
    Into v_�汾��
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '֧�ְ汾';
  
    b_Einvoice_Request_Test.Get_Identity(n_�û�id, v_����Ա���, v_����Ա����);
  
    --v_ҵ���ʶ:01  סԺ,02  ����,03  ����,04  ����,05  �������,06  �Һ�,07  סԺԤ����,08  ���Ԥ����
  
    n_Ʊ���ܽ�� := 0;
    c_��ϸ       := Null;
    v_��ϸ       := Null;
  
    Open c_Balance_Record;
    Fetch c_Balance_Record
      Into r_Balance_Record;
    If c_Balance_Record%NotFound Then
      Code_Out    := 0;
      Message_Out := 'δ�ҵ�ָ���Ľ�������';
      Return;
    End If;
  
    v_סԺ���� := Null;
    For c_�շ�ϸĿ In (Select Min(a.Id) As ����id, a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, Max(a.���㵥λ) As ���㵥λ,
                          Sum(a.��׼����) As �۸�, Avg(Nvl(a.����, 1) * Nvl(a.����, 0)) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��,
                          Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.ʵ�ս��) - Sum(a.ͳ����) As �Էѽ��,
                          Max(t.����) As ҽ����Ŀ����, Max(t.����) As ҽ����Ŀ����, Max(t.ͳ��ȶ�) As ҽ����������, Max(a.ժҪ) As ��ע,
                          Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����id) As ����id,
                          Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, Max(a.�վݷ�Ŀ) As �վݷ�Ŀ, Max(c.����) As �վݷ�Ŀ����, Max(a.��ҳid) As ��ҳid,
                          Max(d.����) As ������, Max(d.���) As �������, Max(b.����) As ��Ŀ����, Max(b.����) As ��Ŀ����, Max(b.���) As ���,
                          Max(q.ҩƷ����) As ҩƷ����
                   From סԺ���ü�¼ A, �շ���ĿĿ¼ B, �վݷ�Ŀ C, �շ���� D, ҩƷ��� M, ҩƷ���� Q, ������ĿĿ¼ J, ����֧������ T, ֧�������� S
                   Where a.����id = n_����id And a.���ʷ��� = 1 And a.�շ���� = d.����(+) And a.�շ�ϸĿid = b.Id And a.�վݷ�Ŀ = c.����(+) And
                         a.�շ�ϸĿid = m.ҩƷid(+) And m.ҩ��id = q.ҩ��id(+) And q.ҩ��id = j.Id(+) And a.���մ���id = t.Id(+) And
                         t.����(+) = 1 And a.���մ���id = s.���մ���id(+)
                   Group By a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, c.����, c.����, j.����, j.����
                   Order By NO, ���) Loop
    
      --listDetailNo  ��ϸ��ˮ��  String  60  ��  ��ϸ��ˮ��
      v_Temp := '{' || '"listDetailNo":"' || b_Einvoice_Request_Test.Zljsonstr(LPad(c_�շ�ϸĿ.����id, 20, '0')) || '"';
      --chargeCode  �շ���Ŀ����  String  50  ��  ��дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���,�磺��λ�ѡ�����
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�վݷ�Ŀ����) || '"';
      --chargeName  �շ���Ŀ����  String  100  ��
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�վݷ�Ŀ) || '"';
      --prescribeCode  ��������  String  60  ��
      v_Temp := v_Temp || ',' || '"prescribeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.No) || '"';
      --listTypeCode  ҩƷ������  String  50  ��  ��ҩƷ�������01��������д
      v_Temp := v_Temp || ',' || '"listTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.������) || '"';
      --listTypeName  ҩƷ�������  String  50  ��  ��ҩƷ�������ƣ��������࿹��Ⱦҩ��
      v_Temp := v_Temp || ',' || '"listTypeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�������) || '"';
      --code  ����  String  50  ��  ��ҩƷ���룬������д
      v_Temp := v_Temp || ',' || '"code":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��Ŀ����) || '"';
      --name  ҩƷ����  String  50  ��  ��ҩƷ���ƣ��������Ƶ�
      v_Temp := v_Temp || ',' || '"name":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��Ŀ����) || '"';
      --form  ����  String  50  ��
      v_Temp := v_Temp || ',' || '"form":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҩƷ����) || '"';
      --specification  ���  String  50  ��
      v_Temp := v_Temp || ',' || '"specification":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���) || '"';
      --unit  ������λ   String  20  ��
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���㵥λ) || '"';
      --std  ����  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�۸�, 1);
      --number  ����  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.����, 1);
      --amt  ���  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ʵ�ս��, 1);
      --selfAmt  �Էѽ��  Number  14,6  ��  ���޽���д0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�Էѽ��, 1);
      --receivableAmt  Ӧ�շ���  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"receivableAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.Ӧ�ս��, 1);
      --medicalCareType  ҽ��ҩƷ����  String  1  ��  1�����Ը�/��
      --          2�����Ը�/��
      --          3��ȫ�Ը�/��
      v_Temp := v_Temp || ',' || '"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����Ŀ����) || '"';
      --medCareItemType  ҽ����Ŀ����  String  100  ��
      v_Temp := v_Temp || ',' || '"medCareItemType":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����Ŀ����) || '"';
      --medReimburseRate  ҽ����������  Number  3,2  ��
      v_Temp := v_Temp || ',' || '"medReimburseRate":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����������, 1);
      --remark  ��ע  String  200  ��
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��ע) || '"';
      --sortNo  ���  Integer  ����  ��  ���
      v_Temp := v_Temp || ',' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���, 1);
      --chrgtype  ��������  String  50  ��
      v_Temp := v_Temp || ',' || '"chrgtype":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��������) || '"}';
    
      If Instr(Nvl(v_סԺ����, '') || ',', ',' || Nvl(c_�շ�ϸĿ.��ҳid, 0) || ',') = 0 Then
        v_סԺ���� := Nvl(v_סԺ����, '') || ',' || Nvl(c_�շ�ϸĿ.��ҳid, 0);
      End If;
      If Length(Nvl(v_��ϸ, '') || v_Temp) > 32700 Then
        If c_��ϸ Is Null Then
          c_��ϸ := To_Clob(v_��ϸ);
        Else
          c_��ϸ := c_��ϸ || To_Clob(',' || v_��ϸ);
        End If;
        v_��ϸ := Null;
      End If;
    
      If v_��ϸ Is Null Then
        v_��ϸ := v_Temp;
      Else
        v_��ϸ := v_��ϸ || ',' || v_Temp;
      End If;
    End Loop;
    If v_סԺ���� Is Not Null Then
      v_סԺ���� := Substr(v_סԺ����, 2);
    End If;
    If v_��ϸ Is Not Null And c_��ϸ Is Not Null Then
      --listDetail  �嵥��Ŀ��ϸ  String  ����  ��  ���A-2,JSON��ʽ�б�
      c_��ϸ := c_��ϸ || ',' || To_Clob(v_��ϸ);
      c_��ϸ := To_Clob(',"listDetail":[') || c_��ϸ || To_Clob(']');
    
      v_��ϸ := Null;
    Elsif v_��ϸ Is Not Null Then
      v_��ϸ := ',"listDetail":[' || v_��ϸ || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --������ϸ
    v_������ϸ := Null;
    c_������ϸ := Null;
    For c_����ͳ�� In (Select Rownum As ���, �վݷ�Ŀ����, �վݷ�Ŀ����, ����, ���㵥λ, Round(����, 2) As ����, Round(���ʽ��, 2) As ���ʽ��,
                          Round(�Էѽ��, 2) As �Էѽ��, ��ע, ���ʽ�� - Round(���ʽ��, 2) As ����
                   From (Select /*+cardinality(b,10)*/
                           c.���� As �վݷ�Ŀ����, a.�վݷ�Ŀ As �վݷ�Ŀ����, 1 As ����, '' As ���㵥λ, Sum(a.���ʽ��) As ����, a.�վݷ�Ŀ,
                           Sum(a.���ʽ��) As ���ʽ��, Sum(a.���ʽ��) - Sum(a.ͳ����) As �Էѽ��, '' As ��ע
                          From סԺ���ü�¼ A, �վݷ�Ŀ C
                          Where a.����id = n_����id And a.�վݷ�Ŀ = c.����(+)
                          Group By c.����, a.�վݷ�Ŀ)) Loop
      --sortNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���շ���Ŀ���ֵ����1�����β������ظ�
      v_Temp := '{' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���, 1);
      --chargeCode  �շ���Ŀ����  String  50  ��  ��дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�վݷ�Ŀ����) || '"';
      --chargeName  �շ���Ŀ����  String  100  ��  ��дҵ��ϵͳ�ڲ���Ŀ����
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�վݷ�Ŀ����) || '"';
      --unit  ������λ  String  20  ��
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���㵥λ) || '"';
      --std  �շѱ�׼  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.����, 1);
      --number  ����  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.����, 1);
      --amt  ���  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���ʽ��, 1);
      --selfAmt  �Էѽ��  Number  14,2  ��  ���޽���д0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�Էѽ��, 1);
      --remark  ��ע  String  200  ��
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.��ע) || '"}';
    
      If Length(Nvl(v_������ϸ, '') || v_Temp) > 32700 Then
        If c_������ϸ Is Null Then
          c_������ϸ := To_Clob(v_������ϸ);
        Else
          c_������ϸ := c_������ϸ || To_Clob(',' || v_������ϸ);
        End If;
        v_������ϸ := Null;
      End If;
    
      If v_������ϸ Is Null Then
        v_������ϸ := v_Temp;
      Else
        v_������ϸ := v_������ϸ || ',' || v_Temp;
      End If;
    
      n_Ʊ���ܽ�� := Nvl(n_Ʊ���ܽ��, 0) + Nvl(c_����ͳ��.���ʽ��, 0);
      n_����ܶ�   := Nvl(n_����ܶ�, 0) + Nvl(c_����ͳ��.����, 0);
    End Loop;
    Totalmoney_Out := n_Ʊ���ܽ��;
    If v_������ϸ Is Not Null And c_������ϸ Is Not Null Then
      c_������ϸ := c_������ϸ || ',' || To_Clob(v_������ϸ);
      --chargeDetail  chargeDetail  �շ���Ŀ��ϸ  �շ���Ŀ��ϸ
      c_������ϸ := To_Clob(',"chargeDetail":[') || c_������ϸ || To_Clob(']');
      v_������ϸ := Null;
    Elsif v_������ϸ Is Not Null Then
      v_������ϸ := ',"chargeDetail":[' || v_������ϸ || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --Ʊ����Ϣ
    --Select Count(1) Into n_���ϴ��� From ����Ʊ��ʹ�ü�¼ Where Ʊ�� = n_Ӧ�ó��� And ����id = n_����id;
    --����ID||_||����Ʊ��ID
    v_ҵ����ˮ�� := n_����id || '_' || Nvl(n_����Ʊ��id, 0);
    v_��Ʊ��     := b_Einvoice_Request_Test.Get_Einvoice_Node(v_����Ա���, v_����Ա����, n_�û�id);
  
    v_Ʊ����Ϣ := '"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ����ˮ��) || '"'; --ҵ����ˮ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"busType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.ҵ���ʶ) || '"'; --ҵ���ʶ
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.����) || '"'; --��������
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"busDateTime":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Balance_Record.�շ�ʱ��, 'yyyymmddHH24miss') || '000') || '"';
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_��Ʊ��) || '"'; --��Ʊ�����:ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.����Ա����) || '"'; --�շ�Ա
  
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_����Ա����) || '"'; --Ʊ�ݱ�����
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"totalAmt":' || b_Einvoice_Request_Test.Zljsonstr(n_Ʊ���ܽ��, 1); --��Ʊ�ܽ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"remark":"' || Nvl(r_Balance_Record.��ע, '') || '"'; --��ע
    -------------------------------------------------------------------------------------------
  
    --ȡ�ɷ���Ϣ
    v_�ɷ� := Null;
    For c_�ɷ� In (Select Max(Decode(��Ϣ��, '������', ��Ϣֵ, '֧��������', ��Ϣֵ, '')) As ֧��������,
                        Max(Decode(��Ϣ��, 'ҽ��֧��������', ��Ϣֵ, 'ҽ��������', ��Ϣֵ, '')) As ҽ��֧��������,
                        Max(Decode(Upper(��Ϣ��), '֧�������ں�USERID', ��Ϣֵ, '')) As ֧�������ں�userid,
                        Max(Decode(Upper(��Ϣ��), '֧����С����USERID', ��Ϣֵ, '')) As ֧����С����userid,
                        Max(Decode(Upper(��Ϣ��), '΢�Ź��ں�OPENID', ��Ϣֵ, '')) As ΢�Ź��ں�openid,
                        Max(Decode(Upper(��Ϣ��), '΢��С����OPENID', ��Ϣֵ, '')) As ΢��С����openid
                 From (Select ��Ϣ��, ��Ϣֵ
                        From ������Ϣ�ӱ�
                        Where ����id = n_����id And ��Ϣ�� In ('֧�������ں�USERID', '֧����С����USERID', '΢�Ź��ں�OPENID', '΢��С����OPENID')
                        Union All
                        Select ������Ŀ, ��������
                        From �������㽻��
                        Where ����id In (Select ID From ����Ԥ����¼ Where ����id = n_����id) And ������Ŀ Like '%������')) Loop
      v_�ɷ� := ',"alipayCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.֧�������ں�userid) || '"'; --����֧�����˻�
      v_�ɷ� := v_�ɷ� || ',"weChatOrderNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.֧��������) || '"'; --΢��֧��������
      If v_�汾�� = 'V3.1.0' Then
        --�ð汾���д˽ӵ�
        v_�ɷ� := v_�ɷ� || ',"weChatMedTransNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.ҽ��֧��������) || '"'; --΢��ҽ��֧��������
      End If;
    
      If c_�ɷ�.΢�Ź��ں�openid Is Not Null Then
        v_�ɷ� := v_�ɷ� || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.΢�Ź��ں�openid) || '"'; --΢�Ź��ںŻ�С�����û�ID
      Else
        v_�ɷ� := v_�ɷ� || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.΢��С����openid) || '"'; --΢�Ź��ںŻ�С�����û�ID
      End If;
      Exit;
    End Loop;
  
    -------------------------------------------------------------------------------------------
    --ȡ֪ͨ��Ϣ
    Select To_Number(Max(����ֵ))
    Into n_ȱʡ�����id
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = 'ȱʡ�����ID';
  
    v_֪ͨ := ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.�ֻ���) || '"'; --�����ֻ�����
    v_֪ͨ := v_֪ͨ || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.Email) || '"'; --���������ַ
    If v_�汾�� = 'V3.1.0' Then
      --�ð汾���д˽ӵ�
      v_֪ͨ := v_֪ͨ || ',"payerType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.�ɿ�����) || '"'; --����������
    End If;
    v_֪ͨ := v_֪ͨ || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.���֤��) || '"'; --ͳһ������ô���
  
    If Nvl(r_Balance_Record.����id, 0) = 0 Then
      --û��һ�ſ����̶�һ�ֿ����
      Select Nvl(Max(����ֵ), '99999')
      Into v_����ֵ
      From �����ӿ�����
      Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ������';
      v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
    
      Select Nvl(Max(����ֵ), '-')
      Into v_����ֵ
      From �����ӿ�����
      Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ���';
      v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --����
    
    Else
      v_Temp := Null;
    
      For c_֪ͨ In (Select ����, ����, ����
                   From (Select b.����id, c.����, c.����, b.����, Decode(b.�����id, n_ȱʡ�����id, 2, c.ȱʡ��־) As ȱʡ��־
                          From ����ҽ�ƿ���Ϣ B, ҽ�ƿ���� C
                          Where b.�����id = c.Id And b.����id = r_Balance_Record.����id
                          Order By ȱʡ��־)
                   Where Rownum < 2) Loop
      
        If c_֪ͨ.���� Is Not Null Then
        
          v_Temp := ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.����) || '"'; --������
          v_Temp := v_Temp || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.����) || '"'; --����
        End If;
        Exit;
      End Loop;
      If v_Temp Is Not Null Then
        v_֪ͨ := v_֪ͨ || v_Temp;
      Else
        If r_Balance_Record.���֤�� Is Not Null Then
          Select Nvl(Max(����ֵ), '99998')
          Into v_����ֵ
          From �����ӿ�����
          Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '���֤�������ͱ��';
          v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
          v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.���֤��) || '"'; --����
        Else
          --û��һ�ſ����̶�һ�ֿ����
          Select Nvl(Max(����ֵ), '99999')
          Into v_����ֵ
          From �����ӿ�����
          Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ������';
          v_֪ͨ := v_֪ͨ || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
          Select Nvl(Max(����ֵ), '-')
          Into v_����ֵ
          From �����ӿ�����
          Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ���';
          v_֪ͨ := v_֪ͨ || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --����
        End If;
      End If;
    End If;
  
    -------------------------------------------------------------------------------------------
    --������Ϣ
    Select Max(����) Into v_Temp From zlRegInfo Where ��Ŀ = 'ҽ�ƻ�������';
  
    --����:1-�շ�;2-���㣨����סԺ���㡢����������㣩��3-Ԥ��
    Select Max(a.����), Max(b.���ջ�������), Max(Nvl(a.��������, c.����))
    Into n_����, v_���ջ�������, v_��������
    From ���ս����¼ A, ������� B, ���ղ��� C
    Where a.���� = b.��� And a.����id = c.Id(+) And a.��¼id = n_����id And a.���� = 2;
  
    If Nvl(n_����, 0) <> 0 Then
      Select Max(ҽ����) Into v_ҽ���� From �����ʻ� Where ����id = n_����id And ���� = n_����;
    End If;
    Select Max(����) Into v_ҽ�Ƹ��ʽ���� From ҽ�Ƹ��ʽ Where ���� = Nvl(r_Balance_Record.ҽ�Ƹ��ʽ, '-');
  
    --medicalInstitution  ҽ�ƻ�������  String  60  ��  ���ա�ҽ�ƻ�����������ʵʩϸ�򡷺͡������������޶�<ҽ�ƻ�����������ʵʩϸ��>�������й����ݵ�֪ͨ��ȷ����ҽ�������������
    v_������Ϣ := ',"medicalInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"';
    --medCareInstitution  ҽ����������  String  60  ��  ҽ��������Ψһ����
    v_������Ϣ := v_������Ϣ || ',"medCareInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_���ջ�������) || '"';
    --medCareTypeCode  ҽ�����ͱ���  String  60  ��
    v_������Ϣ := v_������Ϣ || ',"medCareTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ�Ƹ��ʽ����) || '"';
    --medicalCareType  ҽ����������  String  60  ��  �ɳ���ְ������ҽ�Ʊ��ա�����������ҽ�Ʊ��ա�����ũ�����ҽ�ơ�����ҽ�Ʊ��յȹ���
    v_������Ϣ := v_������Ϣ || ',"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.ҽ�Ƹ��ʽ) || '"';
    --medicalInsuranceID  ����ҽ�����  String  60  ��  �α�����ҽ��ϵͳ�е�Ψһ��ʶ(ҽ����)
    v_������Ϣ := v_������Ϣ || ',"medicalInsuranceID":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ����) || '"';
    --category  ��Ժ��������  String  60  ��
    v_������Ϣ := v_������Ϣ || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.��Ժ��������) || '"';
    --categoryCode  ��Ժ���ұ���  String  60  ��
    v_������Ϣ := v_������Ϣ || ',"categoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.��Ժ���ұ���) || '"';
    --leaveCategory  ��Ժ��������  String  60  ��
    v_������Ϣ := v_������Ϣ || ',"leaveCategory":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.��Ժ��������) || '"';
    --leaveCategoryCode  ��Ժ���ұ���  String  60  ��
    v_������Ϣ := v_������Ϣ || ',"leaveCategoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.��Ժ���ұ���) || '"';
    --hospitalNo  ����סԺ��  String  20  ��  ����Ժ����Ժ�������������̵�Ψһ��
    v_������Ϣ := v_������Ϣ || ',"hospitalNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.סԺ��) || '"';
    --visitNo  סԺ������  String  20  ��  סԺ�ڼ䣬���ڶ�ν��㣬��������������һ��סԺ�����ţ����޾����ţ��ɵ��ڻ���סԺ��
    v_������Ϣ := v_������Ϣ || ',"visitNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.סԺ��) || '"';
    --consultationDate  ��������  String  10  ��  ���߾�ҽʱ��
    --          ��ʽ:yyyy-MM-dd
    v_������Ϣ := v_������Ϣ || ',"consultationDate":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Balance_Record.��Ժ����, 'yyyy-mm-dd')) || '"';
    --patientId  ����ΨһID  String  50  ��  ������ҵ��ϵͳ�е�Ψһ��ʶID���������֤���롣
    v_������Ϣ := v_������Ϣ || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.����id) || '"';
    --patientNo  ���߾�����  String  20  ��  ����ÿ�ξ���һ�ξ����ɵ�һ���µı�š������ߵǼǺţ�
    v_������Ϣ := v_������Ϣ || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.��ҳid) || '"';
    --sex  �Ա�  String  2  ��
    v_������Ϣ := v_������Ϣ || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.�Ա�) || '"';
    --age  ����  String  10  ��
    v_������Ϣ := v_������Ϣ || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.����) || '"';
    --hospitalArea  ����  String  60  ��
    v_������Ϣ := v_������Ϣ || ',"hospitalArea":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.��������) || '"';
    --bedNo  ����  String  20  ��
    v_������Ϣ := v_������Ϣ || ',"bedNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.����) || '"';
    --caseNumber  ������  String  50  ��
    v_������Ϣ := v_������Ϣ || ',"caseNumber":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.������) || '"';
  
    If Instr(v_סԺ����, ',') > 0 Then
      --���ν�����סԺ��
      For c_��ҳ In (Select Min(��Ժ����) As ��Ժ����, Max(��Ժ����) As ��Ժ����, Sum(Nvl(��Ժ����, Sysdate) - ��Ժ����) As סԺ����
                   From ������ҳ
                   Where ����id = r_Balance_Record.����id And
                         ��ҳid In (Select /*+cardinality(A,10)*/
                                   Column_Value
                                  From Table(f_Num2List(v_סԺ����)) A)) Loop
      
        --inHospitalDate  סԺ����  String  10  ��  ��ʽ:yyyy-MM-dd
        v_������Ϣ := v_������Ϣ || ',"inHospitalDate":"' || b_Einvoice_Request_Test.Zljsonstr(To_Char(c_��ҳ.��Ժ����, 'yyyy-mm-dd')) || '"';
        --outHospitalDate  ��Ժ����  String  10  ��  ��ʽ:yyyy-MM-dd
        v_������Ϣ := v_������Ϣ || ',"outHospitalDate":"' || b_Einvoice_Request_Test.Zljsonstr(To_Char(c_��ҳ.��Ժ����, 'yyyy-mm-dd')) || '"';
        --hospitalDays  סԺ����  Number  6,2  ��
        v_������Ϣ := v_������Ϣ || ',"hospitalDays":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_��ҳ.סԺ����, 0), 1);
        Exit;
      
      End Loop;
    Else
      --inHospitalDate  סԺ����  String  10  ��  ��ʽ:yyyy-MM-dd
      v_������Ϣ := v_������Ϣ || ',"inHospitalDate":"' ||
                b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Balance_Record.��Ժ����, 'yyyy-mm-dd')) || '"';
      --outHospitalDate  ��Ժ����  String  10  ��  ��ʽ:yyyy-MM-dd
      v_������Ϣ := v_������Ϣ || ',"outHospitalDate":"' ||
                b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Balance_Record.��Ժ����, 'yyyy-mm-dd')) || '"';
      --hospitalDays  סԺ����  Number  6,2  ��
      v_������Ϣ := v_������Ϣ || ',"hospitalDays":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(r_Balance_Record.סԺ����, 0), 1);
    End If;
  
    -------------------------------------------------------------------------------------------
    --������Ϣ
    v_���� := Null;
  
    --Ԥ���б� 
    For c_��Ԥ�� In (
                  
                  Select q.ƾ֤����, q.ƾ֤����, a.No, Max(a.��Ԥ��) As ��Ԥ��
                  From (Select NO, Sum(��Ԥ��) As ��Ԥ��
                          From ����Ԥ����¼
                          Where ����id = n_����id And Mod(��¼����, 10) = 1) A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ Q
                  Where a.No = b.No And b.��¼���� = 1 And b.Id = q.����id And q.Ʊ�� = 2 And q.�˿�id Is Null) Loop
      --    voucherBatchCode  Ԥ����ƾ֤����  String  50  ��
      v_Temp := '{voucherBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_��Ԥ��.ƾ֤����) || '"';
      --    voucherNo  Ԥ����ƾ֤����  String  20  ��
      v_Temp := v_Temp || ',"voucherNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_��Ԥ��.ƾ֤����) || '"';
      --    voucherAmt  Ԥ����ƾ֤���  Number  14,2  ��  �������Ľ��
      --          ע:��Ԥȫ����㣬�����ܽ��粿�ֽ����㣬����ʵ�ʲ��������
      v_Temp := v_Temp || ',"voucherAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_��Ԥ��.��Ԥ��, 1) || '}';
    
      If Length(Nvl(v_Ԥ��, '') || v_Temp) > 32700 Then
        If v_Ԥ�� Is Null Then
          c_Ԥ�� := To_Clob(v_������ϸ);
        Else
          c_Ԥ�� := c_Ԥ�� || To_Clob(',' || v_Ԥ��);
        End If;
        v_Ԥ�� := Null;
      End If;
    
      If v_Ԥ�� Is Null Then
        v_Ԥ�� := v_Temp;
      Else
        v_Ԥ�� := v_Ԥ�� || ',' || v_Temp;
      End If;
    End Loop;
  
    If v_Ԥ�� Is Not Null And c_��ϸ Is Not Null Then
      --payMentVoucher  Ԥ����ƾ֤���ѿۿ��б�  String  ����  ��  ���㿪��סԺ����Ʊ��ʱ���������ѿۿ��ӦԤ����ƾ֤��Ϣ
      c_Ԥ�� := c_Ԥ�� || ',' || To_Clob(v_Ԥ��);
      c_Ԥ�� := To_Clob(',"payMentVoucher":[') || c_Ԥ�� || To_Clob(']');
    
      v_Ԥ�� := Null;
    Elsif v_Ԥ�� Is Not Null Then
      v_Ԥ�� := ',"payMentVoucher":[' || v_Ԥ�� || ']';
    End If;
  
    --    payMentVoucher  Ԥ����ƾ֤���ѿۿ��б�  String  ����  ��  ���㿪��סԺ����Ʊ��ʱ���������ѿۿ��ӦԤ����ƾ֤��Ϣ
    --          ���A-6,JSON��ʽ�б�
  
    For c_���� In (Select �ֽ�Ԥ��, ֧ƱԤ��, ת��Ԥ��, �����ʻ�֧��, ҽ��ͳ�����֧��, ����ҽ��֧��, �����ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, �ֽ�֧��, 0) As �ֽ��˿�,
                        Decode(Sign(�ֽ�֧��), -1, ֧Ʊ֧��, 0) As ֧Ʊ�˿�, Decode(Sign(�ֽ�֧��), -1, ת��֧��, 0) As ת���˿�,
                        Decode(Sign(�ֽ�֧��), -1, 0, �ֽ�֧��) As �ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, 0, ֧Ʊ֧��) As ֧Ʊ֧��,
                        Decode(Sign(�ֽ�֧��), -1, 0, ת��֧��) As ת��֧��,
                        Nvl(�����ʻ�֧��, 0) + Nvl(ҽ��ͳ�����֧��, 0) + Nvl(����ҽ��֧��, 0) As �����ܶ�,
                        Nvl(�����ܶ�, 0) - Nvl(�����ʻ�֧��, 0) - Nvl(ҽ��ͳ�����֧��, 0) - Nvl(����ҽ��֧��, 0) As �Էѽ��, �����ܶ�, ҽ���������,
                        0 As �����ʻ����
                 From (Select /*+cardinality(b,10)*/
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '�ֽ�', 1, 0), 0) * a.��Ԥ��) As �ֽ�Ԥ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0), 0) * a.��Ԥ��) As ֧ƱԤ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 0, '�ֽ�', 0, 1), 0) * a.��Ԥ��) As ת��Ԥ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '�����˻�֧��', 1, 0)) * a.��Ԥ��) As �����ʻ�֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, 'ҽ��ͳ�����֧��', 1, 0)) * a.��Ԥ��) As ҽ��ͳ�����֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 1, 0)) * a.��Ԥ��) As ����ҽ��֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 0, '�����˻�֧��', 0, 'ҽ��ͳ�����֧��', 0, 1)) *
                              a.��Ԥ��) As �����ֽ�֧��,
                         Max(Decode(Mod(a.��¼����, 10), 1, 0,
                                     Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', �������, '�����˻�֧��', �������, 'ҽ��ͳ�����֧��', �������, ''))) As ҽ���������,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '�ֽ�', 1, 0)) * a.��Ԥ��) As �ֽ�֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0)) * a.��Ԥ��) As ֧Ʊ֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '�ֽ�', 0, '֧Ʊ', 0, 1) * a.��Ԥ��)) As ת��֧��,
                         Sum(��Ԥ��) As �����ܶ�
                        From ����Ԥ����¼ A, ��Ʊ������� C
                        Where a.����id = n_����id And a.���㷽ʽ = c.���㷽ʽ(+))) Loop
      --accountPay  �����˻�֧��  Number  14,2  ��  �����߹涨�ø����˻�֧���α��˵�ҽ�Ʒ��ã�������ҽ�Ʊ���Ŀ¼��Χ�ں�Ŀ¼��Χ��ķ��ã���
      --          ���޽���д0
      v_���� := ',"accountPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.�����ʻ�֧��, 0), 1);
      --fundPay  ҽ��ͳ�����֧��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а��涨�ɻ���ҽ�Ʊ���ͳ�����֧���Ľ�
      --          ���޽���д0
      v_���� := v_���� || ',"fundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.ҽ��ͳ�����֧��, 0), 1);
      --otherfundPay  ����ҽ��֧��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а��涨�ɴ󲡱��ա�ҽ�ƾ���������Աҽ�Ʋ��������䡢��ҵ����Ȼ�����ʽ�֧���Ľ�
      v_���� := v_���� || ',"otherfundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.����ҽ��֧��, 0), 1);
      --ownPay  �Էѽ��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а����йع涨�����ڻ���ҽ�Ʊ���Ŀ¼��Χ��ȫ���ɸ���֧���ķ��ã�
      --          ���޽���д0
      v_���� := v_���� || ',"ownPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�Էѽ��, 1);
      --selfConceitedAmt  �����Ը�  Number  14,2  ��  ҽ�������𸶱�׼�ڸ���֧�����ã�
      --          ���޽���д0
      v_���� := v_���� || ',"selfConceitedAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfPayAmt  �����Ը�  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ������ɸ��˸��������ڻ���ҽ�Ʊ���Ŀ¼��Χ���Ը����ֵĽ���չ�����֡����顢���յȴ�����ѷ�ʽ���ɻ��߶���ѵķ��á�����Ϊ��������˰��ҽ��ר��ӿ۳��ţ�Ϣ�����޽���д0
      v_���� := v_���� || ',"selfPayAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfCashPay  �����ֽ�֧��  Number  14,2  ��  ����ͨ���ֽ����п���΢�š�֧����������֧���Ľ�
      --          ���޽���д0
      v_���� := v_���� || ',"selfCashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ֽ�֧��, 1);
      --cashPay  �ֽ�Ԥ������  Number  14,2  ��
      v_���� := v_���� || ',"cashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ�Ԥ��, 1);
      --chequePay  ֧ƱԤ������  Number  14,2  ��
      v_���� := v_���� || ',"chequePay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧ƱԤ��, 1);
      --transferAccountPay  ת��Ԥ������  Number  14,2  ��
      v_���� := v_���� || ',"transferAccountPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת��Ԥ��, 1);
      --cashRecharge  �������(�ֽ�)  Number  14,2  ��
      v_���� := v_���� || ',"cashRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ�֧��, 1);
      --chequeRecharge  �������(֧Ʊ)  Number  14,2  ��
      v_���� := v_���� || ',"chequeRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧Ʊ֧��, 1);
      --transferRecharge  ������ת�ˣ�  Number  14,2  ��
      v_���� := v_���� || ',"transferRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת��֧��, 1);
      --cashRefund  �˻����(�ֽ�)  Number  14,2  ��
      v_���� := v_���� || ',"cashRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ��˿�, 1);
      --chequeRefund  �˽����(֧Ʊ)  Number  14,2  ��
      v_���� := v_���� || ',"chequeRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧Ʊ�˿�, 1);
      --transferRefund  �˽����(ת��)  Number  14,2  ��
      v_���� := v_���� || ',"transferRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת���˿�, 1);
      --ownAcBalance  �����˻����  Number  14,2  ��
      v_���� := v_���� || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ʻ����, 1);
      --reimbursementAmt  �����ܽ��  Number  14,2  ��  ҽ������󷵻ص��ܽ��
      v_���� := v_���� || ',"reimbursementAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ܶ�, 1);
      --balancedNumber  �����  String  100  ��  ҽ����������ɵĺ���/����Ψһֵ
      v_���� := v_���� || ',"balancedNumber":"' || b_Einvoice_Request_Test.Zljsonstr(c_����.ҽ���������) || '"';
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --��������
    v_�ɷ����� := Null;
    For c_���� In (Select /*+cardinality(b,10)*/
                  Nvl(c.��������, Nvl(d.��������, '-')) As ��������, Sum(��Ԥ��) As �����ܶ�
                 From ����Ԥ����¼ A, �շ��������� C, (Select ���㷽ʽ, �������� From �շ��������� D Where �����id Is Null) D
                 Where a.����id = n_����id And a.�����id = c.�����id(+) And a.���㷽ʽ = c.���㷽ʽ(+) And a.���㷽ʽ = d.���㷽ʽ(+)
                 Group By Nvl(c.��������, Nvl(d.��������, '-'))
                 Order By ��������)
    
     Loop
      --payChannelCode  ������������  String  10  ��
      If v_�ɷ����� Is Null Then
        v_�ɷ����� := Nvl(v_�ɷ�����, '') || '{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.��������, 0)) || '"';
      Else
        v_�ɷ����� := v_�ɷ����� || ',{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.��������, 0)) || '"';
      End If;
      --payChannelValue  �����������  Number  14,2  ��
      v_�ɷ����� := v_�ɷ����� || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.�����ܶ�, 0), 1) || '}';
    End Loop;
  
    If v_�ɷ����� Is Not Null Then
      --payChannelDetail  ���������б�  String  ����  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö����磺��¼4���������б�
      --        ���A-5,JSON��ʽ�б�
      v_�ɷ����� := ',"payChannelDetail":[' || v_�ɷ����� || ']';
    Else
      v_�ɷ����� := ',"payChannelDetail":[]';
    End If;
  
    -------------------------------------------------------------------------------------------
    --����ҽ����Ϣ
    v_����ҽ����Ϣ := Null;
    --otherMedicalList  ����ҽ����Ϣ�б�  String  ����  ��  ��д����δ֪ҽ����Ϣ���ڵ���Ʊ����������ƴ�ӷ�ʽ��ʾ��
    --            ���A-4,JSON��ʽ�б�
    --  infoNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���������ֵ����1�����β������ظ�
    --  infoName  ҽ����Ϣ����  String  100  ��  ����ñ������ͱ��룬�ɲο���¼7ҽ�����������б�
    --  infoValue  ҽ����Ϣֵ  String  100  ��  ����ñ������
    --  infoOther  ҽ��������Ϣ  String  100  ��  ��ҽ������������
  
    -------------------------------------------------------------------------------------------
    --������չ��Ϣ
    v_������չ��Ϣ := Null;
    --otherInfo  ������չ��Ϣ�б�  String  ����  ��  ��д��Ϣ��Ҫ�ڵ���Ʊ���ϵ�����ʾ��������չ��Ϣ��δ֪��Ϣ��
    --          ���A-3,JSON��ʽ�б�
    --  infoNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���������ֵ����1�����β������ظ�
    --  infoName  ��չ��Ϣ����  String  100  ��
    --  infoValue  ��չ��Ϣֵ  String  500  ��
  
    c_������Ϣ := To_Clob('{' || v_Ʊ����Ϣ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_�ɷ�);
    c_������Ϣ := c_������Ϣ || To_Clob(v_֪ͨ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_������Ϣ);
  
    If v_Ԥ�� Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_Ԥ��);
    Else
      c_������Ϣ := c_������Ϣ || c_Ԥ��;
    End If;
  
    c_������Ϣ := c_������Ϣ || To_Clob(v_����);
    c_������Ϣ := c_������Ϣ || To_Clob(v_�ɷ�����);
  
    If v_������չ��Ϣ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_������չ��Ϣ);
    End If;
    If v_����ҽ����Ϣ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_����ҽ����Ϣ);
    End If;
  
    --eBillRelateNo  ҵ��Ʊ�ݹ�����  String  32  ��  ��һ��ҵ��������Ҫ����N�ŵ���Ʊ�ݣ���N�ŵ���Ʊ��Ӧ��ֵ����һ�£����ں��ڹ�����ѯ
    --isArrears  �Ƿ����ͨ  String  1  ��  0-��1-�ǣ���Ƿ���������ҽԺҵ��Ҫ���Ʊ���Ƿ����ͨ��
    --arrearsReason  ������ͨԭ��  String  200  ��  isArrears=0����д������ͨ��ԭ��
    c_������Ϣ := c_������Ϣ || To_Clob(',"eBillRelateNo":"","isArrears":"1","arrearsReason":""');
    If v_������ϸ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_������ϸ);
    Else
      c_������Ϣ := c_������Ϣ || c_������ϸ;
    End If;
  
    If v_��ϸ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_��ϸ);
    Else
      c_������Ϣ := c_������Ϣ || c_��ϸ;
    End If;
    c_������Ϣ  := c_������Ϣ || To_Clob('}');
    Reqdata_Out := c_������Ϣ;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Zybalancedata_Create;

  Procedure Get_Mzbalancedata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    --
    ---------------------------------------------------------------------------
    --����:��ȡ������ʿ�Ʊ����
    --���:
    --    Json_In,��ʽ����
    --  input
    --    occasion N 1  Ӧ�ó���:1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --    einvoice_id  N,1 ��ǰ����Ʊ��ID
    --    balance_id N 1  ����ID  occasion=2ʱ��Ԥ��id;occasion<>2��ʾ����id
    --    writeoff_id  N 1  ����ID  occasion=2ʱ������Ԥ��id;occasion<>2��ʾ����id
    --����:
    --  ReqData_Out-���ص�ҵ����������
    --  Totalmoney_Out-Ʊ���ܶ�
    --  Code_Out-��ȡ�Ƿ�ɹ���0-ʧ�ܣ�1-�ɹ�
    --  Message_Out ������Ϣ
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_Ӧ�ó��� Number(2);
    n_����id   ����Ԥ����¼.����id%Type;
  
    v_ҵ����ˮ�� Varchar2(50);
  
    v_��Ʊ��       Varchar2(100);
    v_�ɷ�         Varchar2(32767);
    v_Ʊ����Ϣ     Varchar2(32767);
    v_������Ϣ     Varchar2(32767);
    v_֪ͨ         Varchar2(32767);
    v_�ɷ�����     Varchar2(32767);
    v_����         Varchar2(32767);
    v_������չ��Ϣ Varchar2(32767);
    v_����ҽ����Ϣ Varchar2(32767);
    c_��ϸ         Clob;
    v_��ϸ         Varchar2(32767);
    c_������ϸ     Clob;
    v_������ϸ     Varchar2(32767);
    c_������Ϣ     Clob; --���շ��صĽ�����Ϣ��
  
    v_�������� ������ü�¼.����%Type;
    v_�����Ա� ������ü�¼.�Ա�%Type;
    v_�������� ������ü�¼.����%Type;
  
    n_ȱʡ�����id     Number(18);
    v_����ֵ           Varchar2(100);
    n_Ʊ���ܽ��       ������ü�¼.���ʽ��%Type;
    n_����ܶ�         ������ü�¼.���ʽ��%Type;
    n_�û�id           ��Ա��.Id%Type;
    v_����Ա���       ��Ա��.���%Type;
    v_����Ա����       ��Ա��.����%Type;
    v_Temp             Varchar2(32767);
    v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
    v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
    n_����             ���ս����¼.����%Type;
    v_���ջ�������     �������.���ջ�������%Type;
    n_ҽ�����         ������ü�¼.ҽ�����%Type;
    n_�Һ�id           ������ü�¼.�Һ�id%Type;
    v_��������         ���ղ���.����%Type;
    v_��������         Varchar2(20);
    v_������ұ���     ���ű�.����%Type;
    v_�����������     ���ű�.����%Type;
    v_������         Varchar2(50);
    v_ҽ����           �����ʻ�.ҽ����%Type;
    v_�汾��           Varchar2(30);
    n_����Ʊ��id       ����Ʊ��ʹ�ü�¼.Id%Type;
  
    Cursor c_Balance_Record Is
      Select a.No, a.�շ�ʱ��, a.��������, a.����Ա���, a.����Ա����, a.����id, a.��ҳid, Decode(Nvl(a.����id, 0), 0, a.ԭ��, c.����) As ����,
             '' As �Ա�, '' As ����, c.�����, a.��ע, a.���ʽ��, Decode(Nvl(a.����id, 0), 0, q.�����ʼ�, c.Email) As Email, q.��ϵ��,
             Decode(Nvl(a.����id, 0), 0, q.������ô���, c.���֤��) As ���֤��,
             Decode(Nvl(a.����id, 0), 0, Nvl(q.�绰, To_Char(j.�ƶ��绰)), c.�ֻ���) As �ֻ���,
             Decode(Nvl(a.����id, 0), 0, 2, 1) As �ɿ�����, Decode(Nvl(a.��������, 0), 1, '02', '01') As ҵ���ʶ, c.����� As ������
      From ���˽��ʼ�¼ A, ������Ϣ C, ��Լ��λ Q, ��Ա�� J
      Where a.Id = n_����id And a.����id = c.����id(+) And a.ԭ�� = q.����(+) And q.��ϵ�� = j.����(+);
    r_Balance_Record c_Balance_Record%RowType;
  
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_Ӧ�ó��� := Nvl(j_Json.Get_Number('occasion'), 0);
    n_����id   := j_Json.Get_Number('balance_id');
    --n_����id   := Nvl(j_Json.Get_Number('writeoff_id'), 0);
  
    n_����Ʊ��id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
  
    If Nvl(n_Ӧ�ó���, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '��Ч��Ӧ�ó���';
      Return;
    End If;
  
    Select Nvl(Max(����ֵ), 'V2.0.3')
    Into v_�汾��
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '֧�ְ汾';
  
    b_Einvoice_Request_Test.Get_Identity(n_�û�id, v_����Ա���, v_����Ա����);
  
    --v_ҵ���ʶ:01  סԺ,02  ����,03  ����,04  ����,05  �������,06  �Һ�,07  סԺԤ����,08  ���Ԥ����
  
    n_Ʊ���ܽ�� := 0;
    c_��ϸ       := Null;
    v_��ϸ       := Null;
  
    Open c_Balance_Record;
    Fetch c_Balance_Record
      Into r_Balance_Record;
    If c_Balance_Record%NotFound Then
      Code_Out    := 0;
      Message_Out := 'δ�ҵ�ָ���Ľ�������';
      Return;
    End If;
  
    For c_�շ�ϸĿ In (Select Min(a.Id) As ����id, a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���) As ���, a.�շ�ϸĿid, Max(a.���㵥λ) As ���㵥λ,
                          Sum(a.��׼����) As �۸�, Avg(Nvl(a.����, 1) * Nvl(a.����, 0)) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��,
                          Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.���ʽ��) As ���ʽ��, Sum(a.ʵ�ս��) - Sum(a.ͳ����) As �Էѽ��,
                          Max(t.����) As ҽ����Ŀ����, Max(t.����) As ҽ����Ŀ����, Max(t.ͳ��ȶ�) As ҽ����������, Max(a.ժҪ) As ��ע,
                          Max(a.��������) As ��������, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.����) As ����,
                          Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��,
                          Max(a.���ʽ) As ���ʽ����, Max(a.�վݷ�Ŀ) As �վݷ�Ŀ, Max(c.����) As �վݷ�Ŀ����, Max(a.ҽ�����) As ҽ�����,
                          Max(a.�Һ�id) As �Һ�id, Max(d.����) As ������, Max(d.���) As �������, Max(b.����) As ��Ŀ����, Max(b.����) As ��Ŀ����,
                          Max(b.���) As ���, Max(q.ҩƷ����) As ҩƷ����
                   From ������ü�¼ A, �շ���ĿĿ¼ B, �վݷ�Ŀ C, �շ���� D, ҩƷ��� M, ҩƷ���� Q, ������ĿĿ¼ J, ����֧������ T, ֧�������� S
                   Where ����id = n_����id And a.���ʷ��� = 1 And a.�շ���� = d.����(+) And a.�շ�ϸĿid = b.Id And a.�վݷ�Ŀ = c.����(+) And
                         a.�շ�ϸĿid = m.ҩƷid(+) And m.ҩ��id = q.ҩ��id(+) And q.ҩ��id = j.Id(+) And a.���մ���id = t.Id(+) And
                         t.����(+) = 1 And a.���մ���id = s.���մ���id(+)
                   Group By a.No, a.��¼״̬, a.����id, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid, c.����, c.����, j.����, j.����
                   Order By NO, ���) Loop
      If v_�������� Is Null Then
        v_�������� := c_�շ�ϸĿ.����;
        v_�����Ա� := c_�շ�ϸĿ.�Ա�;
        v_�������� := c_�շ�ϸĿ.����;
      End If;
    
      If v_ҽ�Ƹ��ʽ���� Is Null Then
        v_ҽ�Ƹ��ʽ���� := c_�շ�ϸĿ.���ʽ����;
      End If;
      If Nvl(n_ҽ�����, 0) = 0 Then
        n_ҽ����� := c_�շ�ϸĿ.ҽ�����;
      End If;
      If Nvl(n_�Һ�id, 0) = 0 Then
        n_�Һ�id := c_�շ�ϸĿ.�Һ�id;
      End If;
    
      --listDetailNo  ��ϸ��ˮ��  String  60  ��  ��ϸ��ˮ��
      v_Temp := '{' || '"listDetailNo":"' || b_Einvoice_Request_Test.Zljsonstr(LPad(c_�շ�ϸĿ.����id, 20, '0')) || '"';
      --chargeCode  �շ���Ŀ����  String  50  ��  ��дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���,�磺��λ�ѡ�����
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�վݷ�Ŀ����) || '"';
      --chargeName  �շ���Ŀ����  String  100  ��
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�վݷ�Ŀ) || '"';
      --prescribeCode  ��������  String  60  ��
      v_Temp := v_Temp || ',' || '"prescribeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.No) || '"';
      --listTypeCode  ҩƷ������  String  50  ��  ��ҩƷ�������01��������д
      v_Temp := v_Temp || ',' || '"listTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.������) || '"';
      --listTypeName  ҩƷ�������  String  50  ��  ��ҩƷ�������ƣ��������࿹��Ⱦҩ��
      v_Temp := v_Temp || ',' || '"listTypeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�������) || '"';
      --code  ����  String  50  ��  ��ҩƷ���룬������д
      v_Temp := v_Temp || ',' || '"code":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��Ŀ����) || '"';
      --name  ҩƷ����  String  50  ��  ��ҩƷ���ƣ��������Ƶ�
      v_Temp := v_Temp || ',' || '"name":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��Ŀ����) || '"';
      --form  ����  String  50  ��
      v_Temp := v_Temp || ',' || '"form":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҩƷ����) || '"';
      --specification  ���  String  50  ��
      v_Temp := v_Temp || ',' || '"specification":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���) || '"';
      --unit  ������λ   String  20  ��
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���㵥λ) || '"';
      --std  ����  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�۸�, 1);
      --number  ����  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.����, 1);
      --amt  ���  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ʵ�ս��, 1);
      --selfAmt  �Էѽ��  Number  14,6  ��  ���޽���д0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.�Էѽ��, 1);
      --receivableAmt  Ӧ�շ���  Number  14,6  ��
      v_Temp := v_Temp || ',' || '"receivableAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.Ӧ�ս��, 1);
      --medicalCareType  ҽ��ҩƷ����  String  1  ��  1�����Ը�/��
      --          2�����Ը�/��
      --          3��ȫ�Ը�/��
      v_Temp := v_Temp || ',' || '"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����Ŀ����) || '"';
      --medCareItemType  ҽ����Ŀ����  String  100  ��
      v_Temp := v_Temp || ',' || '"medCareItemType":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����Ŀ����) || '"';
      --medReimburseRate  ҽ����������  Number  3,2  ��
      v_Temp := v_Temp || ',' || '"medReimburseRate":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.ҽ����������, 1);
      --remark  ��ע  String  200  ��
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��ע) || '"';
      --sortNo  ���  Integer  ����  ��  ���
      v_Temp := v_Temp || ',' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.���, 1);
      --chrgtype  ��������  String  50  ��
      v_Temp := v_Temp || ',' || '"chrgtype":"' || b_Einvoice_Request_Test.Zljsonstr(c_�շ�ϸĿ.��������) || '"}';
    
      If Length(Nvl(v_��ϸ, '') || v_Temp) > 32700 Then
        If c_��ϸ Is Null Then
          c_��ϸ := To_Clob(v_��ϸ);
        Else
          c_��ϸ := c_��ϸ || To_Clob(',' || v_��ϸ);
        End If;
        v_��ϸ := Null;
      End If;
    
      If v_��ϸ Is Null Then
        v_��ϸ := v_Temp;
      Else
        v_��ϸ := v_��ϸ || ',' || v_Temp;
      End If;
    End Loop;
  
    If v_��ϸ Is Not Null And c_��ϸ Is Not Null Then
      --listDetail  �嵥��Ŀ��ϸ  String  ����  ��  ���A-2,JSON��ʽ�б�
      c_��ϸ := c_��ϸ || ',' || To_Clob(v_��ϸ);
      c_��ϸ := To_Clob(',"listDetail":[') || c_��ϸ || To_Clob(']');
    
      v_��ϸ := Null;
    Elsif v_��ϸ Is Not Null Then
      v_��ϸ := ',"listDetail":[' || v_��ϸ || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --������ϸ
    v_������ϸ := Null;
    c_������ϸ := Null;
    For c_����ͳ�� In (Select Rownum As ���, �վݷ�Ŀ����, �վݷ�Ŀ����, ����, ���㵥λ, Round(����, 2) As ����, Round(���ʽ��, 2) As ���ʽ��,
                          Round(�Էѽ��, 2) As �Էѽ��, ��ע, ���ʽ�� - Round(���ʽ��, 2) As ����
                   From (Select /*+cardinality(b,10)*/
                           c.���� As �վݷ�Ŀ����, a.�վݷ�Ŀ As �վݷ�Ŀ����, 1 As ����, '' As ���㵥λ, Sum(a.���ʽ��) As ����, a.�վݷ�Ŀ,
                           Sum(a.���ʽ��) As ���ʽ��, Sum(a.���ʽ��) - Sum(a.ͳ����) As �Էѽ��, '' As ��ע
                          From ������ü�¼ A, �վݷ�Ŀ C
                          Where a.����id = n_����id And a.�վݷ�Ŀ = c.����(+)
                          Group By c.����, a.�վݷ�Ŀ)) Loop
      --sortNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���շ���Ŀ���ֵ����1�����β������ظ�
      v_Temp := '{' || '"sortNo":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���, 1);
      --chargeCode  �շ���Ŀ����  String  50  ��  ��дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
      v_Temp := v_Temp || ',' || '"chargeCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�վݷ�Ŀ����) || '"';
      --chargeName  �շ���Ŀ����  String  100  ��  ��дҵ��ϵͳ�ڲ���Ŀ����
      v_Temp := v_Temp || ',' || '"chargeName":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�վݷ�Ŀ����) || '"';
      --unit  ������λ  String  20  ��
      v_Temp := v_Temp || ',' || '"unit":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���㵥λ) || '"';
      --std  �շѱ�׼  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"std":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.����, 1);
      --number  ����  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"number":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.����, 1);
      --amt  ���  Number  14,2  ��
      v_Temp := v_Temp || ',' || '"amt":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.���ʽ��, 1);
      --selfAmt  �Էѽ��  Number  14,2  ��  ���޽���д0
      v_Temp := v_Temp || ',' || '"selfAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.�Էѽ��, 1);
      --remark  ��ע  String  200  ��
      v_Temp := v_Temp || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(c_����ͳ��.��ע) || '"}';
    
      If Length(Nvl(v_������ϸ, '') || v_Temp) > 32700 Then
        If c_������ϸ Is Null Then
          c_������ϸ := To_Clob(v_������ϸ);
        Else
          c_������ϸ := c_������ϸ || To_Clob(',' || v_������ϸ);
        End If;
        v_������ϸ := Null;
      End If;
    
      If v_������ϸ Is Null Then
        v_������ϸ := v_Temp;
      Else
        v_������ϸ := v_������ϸ || ',' || v_Temp;
      End If;
    
      n_Ʊ���ܽ�� := Nvl(n_Ʊ���ܽ��, 0) + Nvl(c_����ͳ��.���ʽ��, 0);
      n_����ܶ�   := Nvl(n_����ܶ�, 0) + Nvl(c_����ͳ��.����, 0);
    End Loop;
    Totalmoney_Out := n_Ʊ���ܽ��;
    If v_������ϸ Is Not Null And c_������ϸ Is Not Null Then
      c_������ϸ := c_������ϸ || ',' || To_Clob(v_������ϸ);
      --chargeDetail �շ���Ŀ��ϸ  String  ����  ��  ���A-1,JSON��ʽ�б�
      c_������ϸ := To_Clob(',"chargeDetail":[') || c_������ϸ || To_Clob(']');
      v_������ϸ := Null;
    Elsif v_������ϸ Is Not Null Then
      v_������ϸ := ',"chargeDetail":[' || v_������ϸ || ']';
    End If;
  
    -------------------------------------------------------------------------------------------
    --Ʊ����Ϣ
    --Select Count(1) Into n_���ϴ��� From ����Ʊ��ʹ�ü�¼ Where Ʊ�� = n_Ӧ�ó��� And ����id = n_����id;
    --lpad(����Ʊ�����ϴ���,5) & Lpad(ԭ����ID,20)
    --����ID||_||����Ʊ��ID
    v_ҵ����ˮ�� := n_����id || '_' || Nvl(n_����Ʊ��id, 0);
    v_��Ʊ��     := b_Einvoice_Request_Test.Get_Einvoice_Node(v_����Ա���, v_����Ա����, n_�û�id);
  
    v_Ʊ����Ϣ := '"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ����ˮ��) || '"'; --ҵ����ˮ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"busType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.ҵ���ʶ) || '"'; --ҵ���ʶ
    If Nvl(r_Balance_Record.����id, 0) = 0 Then
      v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.����) || '"'; --��������
    Else
      v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payer":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --��������
    End If;
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"busDateTime":"' ||
              b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Balance_Record.�շ�ʱ��, 'yyyymmddHH24miss') || '000') || '"';
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_��Ʊ��) || '"'; --��Ʊ�����:ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.����Ա����) || '"'; --�շ�Ա
  
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_����Ա����) || '"'; --Ʊ�ݱ�����
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"totalAmt":' || b_Einvoice_Request_Test.Zljsonstr(n_Ʊ���ܽ��, 1); --��Ʊ�ܽ��
    v_Ʊ����Ϣ := v_Ʊ����Ϣ || ',"remark":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.��ע) || '"'; --��ע
    -------------------------------------------------------------------------------------------
  
    --ȡ�ɷ���Ϣ
    v_�ɷ� := Null;
    For c_�ɷ� In (Select Max(Decode(��Ϣ��, '������', ��Ϣֵ, '֧��������', ��Ϣֵ, '')) As ֧��������,
                        Max(Decode(��Ϣ��, 'ҽ��֧��������', ��Ϣֵ, 'ҽ��������', ��Ϣֵ, '')) As ҽ��֧��������,
                        Max(Decode(Upper(��Ϣ��), '֧�������ں�USERID', ��Ϣֵ, '')) As ֧�������ں�userid,
                        Max(Decode(Upper(��Ϣ��), '֧����С����USERID', ��Ϣֵ, '')) As ֧����С����userid,
                        Max(Decode(Upper(��Ϣ��), '΢�Ź��ں�OPENID', ��Ϣֵ, '')) As ΢�Ź��ں�openid,
                        Max(Decode(Upper(��Ϣ��), '΢��С����OPENID', ��Ϣֵ, '')) As ΢��С����openid
                 From (Select ��Ϣ��, ��Ϣֵ
                        From ������Ϣ�ӱ�
                        Where ����id = r_Balance_Record.����id And
                              ��Ϣ�� In ('֧�������ں�USERID', '֧����С����USERID', '΢�Ź��ں�OPENID', '΢��С����OPENID')
                        Union All
                        Select ������Ŀ, ��������
                        From �������㽻��
                        Where ����id In (Select ID From ����Ԥ����¼ Where ����id = n_����id) And ������Ŀ Like '%������')) Loop
      v_�ɷ� := ',"alipayCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.֧�������ں�userid) || '"'; --����֧�����˻�
      v_�ɷ� := v_�ɷ� || ',"weChatOrderNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.֧��������) || '"'; --΢��֧��������
      If v_�汾�� = 'V3.1.0' Then
        --�ð汾���д˽ӵ�
        v_�ɷ� := v_�ɷ� || ',"weChatMedTransNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.ҽ��֧��������) || '"'; --΢��ҽ��֧��������
      End If;
    
      If c_�ɷ�.΢�Ź��ں�openid Is Not Null Then
        v_�ɷ� := v_�ɷ� || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.΢�Ź��ں�openid) || '"'; --΢�Ź��ںŻ�С�����û�ID
      Else
        v_�ɷ� := v_�ɷ� || ',"openID":"' || b_Einvoice_Request_Test.Zljsonstr(c_�ɷ�.΢��С����openid) || '"'; --΢�Ź��ںŻ�С�����û�ID
      End If;
      Exit;
    End Loop;
  
    -------------------------------------------------------------------------------------------
    --ȡ֪ͨ��Ϣ
    Select To_Number(Max(����ֵ))
    Into n_ȱʡ�����id
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = 'ȱʡ�����ID';
  
    v_֪ͨ := ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.�ֻ���) || '"'; --�����ֻ�����
    v_֪ͨ := v_֪ͨ || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.Email) || '"'; --���������ַ
    If v_�汾�� = 'V3.1.0' Then
      --�ð汾���д˽ӵ�
      v_֪ͨ := v_֪ͨ || ',"payerType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.�ɿ�����) || '"'; --����������
    End If;
    v_֪ͨ := v_֪ͨ || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.���֤��) || '"'; --ͳһ������ô���
  
    v_Temp := Null;
    If Nvl(r_Balance_Record.����id, 0) <> 0 Then
    
      For c_֪ͨ In (Select Max(����) As �����, Max(����) As ����
                   From (
                          
                          Select ����id, ����, ����, ����
                          From (Select b.����id, c.����, c.����, b.����, Decode(b.�����id, n_ȱʡ�����id, 2, c.ȱʡ��־) As ȱʡ��־
                                  From ����ҽ�ƿ���Ϣ B, ҽ�ƿ���� C
                                  Where b.�����id = c.Id And b.����id = Nvl(r_Balance_Record.����id, 0)
                                  Order By ȱʡ��־)
                          Where Rownum < 2)) Loop
      
        If c_֪ͨ.����� Is Not Null Then
          v_Temp := ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.�����) || '"'; --������
          v_Temp := v_Temp || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_֪ͨ.����) || '"'; --����
        End If;
        Exit;
      End Loop;
      If r_Balance_Record.���֤�� Is Not Null And v_Temp Is Null Then
        Select Nvl(Max(����ֵ), '99998')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '���֤�������ͱ��';
        v_Temp := ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
        v_Temp := v_Temp || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.���֤��) || '"'; --����
      End If;
    End If;
    If v_Temp Is Null Then
      --û��һ�ſ����̶�һ�ֿ����
      Select Nvl(Max(����ֵ), '99999')
      Into v_����ֵ
      From �����ӿ�����
      Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ������';
      v_Temp := ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --������
      Select Nvl(Max(����ֵ), '-')
      Into v_����ֵ
      From �����ӿ�����
      Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ���';
      v_Temp := v_Temp || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_����ֵ) || '"'; --����
    End If;
    v_֪ͨ := v_֪ͨ || v_Temp;
  
    -------------------------------------------------------------------------------------------
    --������Ϣ
    Select Max(����) Into v_Temp From zlRegInfo Where ��Ŀ = 'ҽ�ƻ�������';
  
    --����:1-�շ�;2-���㣨����סԺ���㡢����������㣩��3-Ԥ��
    Select Max(a.����), Max(b.���ջ�������), Max(Nvl(a.��������, c.����))
    Into n_����, v_���ջ�������, v_��������
    From ���ս����¼ A, ������� B, ���ղ��� C
    Where a.���� = b.��� And a.����id = c.Id(+) And a.��¼id = n_����id And a.���� = 2;
  
    Select Max(����) Into v_ҽ�Ƹ��ʽ���� From ҽ�Ƹ��ʽ Where ���� = v_ҽ�Ƹ��ʽ����;
    If Nvl(n_����, 0) <> 0 Then
      Select Max(ҽ����) Into v_ҽ���� From �����ʻ� Where ����id = Nvl(r_Balance_Record.����id, 0) And ���� = n_����;
    End If;
  
    v_������ := Null;
    If Nvl(n_ҽ�����, 0) <> 0 Then
      Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')), Max(b.����), Max(b.����), Max(a.No)
      Into v_��������, v_������ұ���, v_�����������, v_������
      From ���˹Һż�¼ A, ���ű� B
      Where a.ִ�в���id = b.Id And
            a.No = (Select Max(�Һŵ�) From ����ҽ����¼ Where ID = n_ҽ����� Or ���id = n_ҽ�����);
    Elsif Nvl(n_�Һ�id, 0) <> 0 Then
      Select Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')), Max(b.����), Max(b.����), Max(a.No)
      Into v_��������, v_������ұ���, v_�����������, v_������
      From ���˹Һż�¼ A, ���ű� B
      Where a.ִ�в���id = b.Id And a.Id = n_�Һ�id;
    End If;
    If v_������ Is Null And Nvl(r_Balance_Record.����id, 0) <> 0 Then
      --ȡ���һ�ιҺ�ID
      Select Max(a.Id), Max(To_Char(a.����ʱ��, 'yyyy-mm-dd')), Max(b.����), Max(b.����), Max(a.No)
      Into n_�Һ�id, v_��������, v_������ұ���, v_�����������, v_������
      From ���˹Һż�¼ A, ���ű� B
      Where a.ִ�в���id = b.Id And a.Id = (Select ID
                                        From (Select ID, ����ʱ��
                                               From ���˹Һż�¼
                                               Where ����id = Nvl(r_Balance_Record.����id, 0)
                                               Order By ����ʱ�� Desc)
                                        Where Rownum < 2);
    End If;
  
    If v_�������� Is Null And Nvl(n_����, 0) <> 0 Then
    
      Select Max(��������)
      Into v_��������
      From (
             
             Select Distinct a.���� As ��������
             From ���ղ��� A, ������׼��Ŀ B
             Where a.���� = n_���� And a.Id = b.����id And
                   b.�շ�ϸĿid In (Select Distinct �շ�ϸĿid From ������ü�¼ Where ����id = n_����id)
             Union All
             Select Distinct a.���� As ��������
             From ���ղ��� A, ������׼��Ŀ B
             Where a.���� = n_���� And a.Id = b.����id And
                   b.���� In (Select Distinct ���մ���id From ������ü�¼ Where ����id = n_����id))
      Where Rownum < 2;
    End If;
    v_������Ϣ := ',"medicalInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"'; --ҽ�ƻ�������
    v_������Ϣ := v_������Ϣ || ',"medCareInstitution":"' || b_Einvoice_Request_Test.Zljsonstr(v_���ջ�������) || '"'; --ҽ��������Ψһ����
    v_������Ϣ := v_������Ϣ || ',"medCareTypeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ�Ƹ��ʽ����) || '"'; --ҽ�����ͱ���
    v_������Ϣ := v_������Ϣ || ',"medicalCareType":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ�Ƹ��ʽ����) || '"'; --ȡֵ��Χ����ְ������ҽ�Ʊ��ա�����������ҽ�Ʊ��գ�����������ҽ�Ʊ��ա�����ũ�����ҽ�Ʊ��գ�������ҽ�Ʊ��ա���ҽ����
    v_������Ϣ := v_������Ϣ || ',"medicalInsuranceID":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҽ����) || '"';
    v_������Ϣ := v_������Ϣ || ',"consultationDate":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --���߾�ҽʱ��
    v_������Ϣ := v_������Ϣ || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(v_�����������) || '"'; --�������
    v_������Ϣ := v_������Ϣ || ',"patientCategoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_������ұ���) || '"'; --������ұ���
    --patientNo  ���߾�����  String  20  ��  ����ÿ�ξ���һ�ξ����ɵ�һ���µı�š������ߵǼǺţ�
    v_������Ϣ := v_������Ϣ || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_������) || '"';
  
    v_������Ϣ := v_������Ϣ || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(r_Balance_Record.����id, 0)) || '"'; --������ҵ��ϵͳ�е�Ψһ��ʶID���������֤���롣
    If Nvl(r_Balance_Record.����id, 0) = 0 Then
      v_������Ϣ := v_������Ϣ || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.�Ա�) || '"'; --�Ա�
      v_������Ϣ := v_������Ϣ || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.����) || '"'; --����
    Else
      v_������Ϣ := v_������Ϣ || ',"sex":"' || b_Einvoice_Request_Test.Zljsonstr(v_�����Ա�) || '"'; --�Ա�
      v_������Ϣ := v_������Ϣ || ',"age":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --����
    End If;
    v_������Ϣ := v_������Ϣ || ',"caseNumber":"' || b_Einvoice_Request_Test.Zljsonstr(r_Balance_Record.������) || '"'; --������
    v_������Ϣ := v_������Ϣ || ',"specialDiseasesName":"' || b_Einvoice_Request_Test.Zljsonstr(v_��������) || '"'; --���ⲡ������
    -------------------------------------------------------------------------------------------
  
    --������Ϣ
    v_���� := Null;
    For c_���� In (Select �ֽ�Ԥ��, ֧ƱԤ��, ת��Ԥ��, �����ʻ�֧��, ҽ��ͳ�����֧��, ����ҽ��֧��, �����ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, �ֽ�֧��, 0) As �ֽ��˿�,
                        Decode(Sign(�ֽ�֧��), -1, ֧Ʊ֧��, 0) As ֧Ʊ�˿�, Decode(Sign(�ֽ�֧��), -1, ת��֧��, 0) As ת���˿�,
                        Decode(Sign(�ֽ�֧��), -1, 0, �ֽ�֧��) As �ֽ�֧��, Decode(Sign(�ֽ�֧��), -1, 0, ֧Ʊ֧��) As ֧Ʊ֧��,
                        Decode(Sign(�ֽ�֧��), -1, 0, ת��֧��) As ת��֧��,
                        Nvl(�����ʻ�֧��, 0) + Nvl(ҽ��ͳ�����֧��, 0) + Nvl(����ҽ��֧��, 0) As �����ܶ�,
                        Nvl(�����ܶ�, 0) - Nvl(�����ʻ�֧��, 0) - Nvl(ҽ��ͳ�����֧��, 0) - Nvl(����ҽ��֧��, 0) As �Էѽ��, �����ܶ�, ҽ���������,
                        0 As �����ʻ����
                 From (Select /*+cardinality(b,10)*/
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '�ֽ�', 1, 0), 0) * a.��Ԥ��) As �ֽ�Ԥ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0), 0) * a.��Ԥ��) As ֧ƱԤ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, Decode(a.���㷽ʽ, '֧Ʊ', 0, '�ֽ�', 0, 1), 0) * a.��Ԥ��) As ת��Ԥ��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '�����˻�֧��', 1, 0)) * a.��Ԥ��) As �����ʻ�֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, 'ҽ��ͳ�����֧��', 1, 0)) * a.��Ԥ��) As ҽ��ͳ�����֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 1, 0)) * a.��Ԥ��) As ����ҽ��֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', 0, '�����˻�֧��', 0, 'ҽ��ͳ�����֧��', 0, 1)) *
                              a.��Ԥ��) As �����ֽ�֧��,
                         Max(Decode(Mod(a.��¼����, 10), 1, 0,
                                     Decode(c.��Ʊ���㷽ʽ, '����ҽ��֧��', �������, '�����˻�֧��', �������, 'ҽ��ͳ�����֧��', �������, ''))) As ҽ���������,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '�ֽ�', 1, 0)) * a.��Ԥ��) As �ֽ�֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '֧Ʊ', 1, 0)) * a.��Ԥ��) As ֧Ʊ֧��,
                         Sum(Decode(Mod(a.��¼����, 10), 1, 0, Decode(a.���㷽ʽ, '�ֽ�', 0, '֧Ʊ', 0, 1) * a.��Ԥ��)) As ת��֧��,
                         Sum(��Ԥ��) As �����ܶ�
                        From ����Ԥ����¼ A, ��Ʊ������� C
                        Where a.����id = n_����id And a.���㷽ʽ = c.���㷽ʽ(+)))
    
     Loop
      --accountPay  �����˻�֧��  Number  14,2  ��  �����߹涨�ø����˻�֧���α��˵�ҽ�Ʒ��ã�������ҽ�Ʊ���Ŀ¼��Χ�ں�Ŀ¼��Χ��ķ��ã���
      --          ���޽���д0
      v_���� := ',"accountPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.�����ʻ�֧��, 0), 1);
      --fundPay  ҽ��ͳ�����֧��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а��涨�ɻ���ҽ�Ʊ���ͳ�����֧���Ľ�
      --          ���޽���д0
      v_���� := v_���� || ',"fundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.ҽ��ͳ�����֧��, 0), 1);
      --otherfundPay  ����ҽ��֧��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а��涨�ɴ󲡱��ա�ҽ�ƾ���������Աҽ�Ʋ��������䡢��ҵ����Ȼ�����ʽ�֧���Ľ�
      v_���� := v_���� || ',"otherfundPay":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.����ҽ��֧��, 0), 1);
      --ownPay  �Էѽ��  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ����а����йع涨�����ڻ���ҽ�Ʊ���Ŀ¼��Χ��ȫ���ɸ���֧���ķ��ã�
      --          ���޽���д0
      v_���� := v_���� || ',"ownPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�Էѽ��, 1);
      --selfConceitedAmt  �����Ը�  Number  14,2  ��  ҽ�������𸶱�׼�ڸ���֧�����ã�
      --          ���޽���д0
      v_���� := v_���� || ',"selfConceitedAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfPayAmt  �����Ը�  Number  14,2  ��  ���߱��ξ�ҽ��������ҽ�Ʒ������ɸ��˸��������ڻ���ҽ�Ʊ���Ŀ¼��Χ���Ը����ֵĽ���չ�����֡����顢���յȴ�����ѷ�ʽ���ɻ��߶���ѵķ��á�����Ϊ��������˰��ҽ��ר��ӿ۳��ţ�Ϣ�����޽���д0
      v_���� := v_���� || ',"selfPayAmt":' || b_Einvoice_Request_Test.Zljsonstr(0, 1);
      --selfCashPay  �����ֽ�֧��  Number  14,2  ��  ����ͨ���ֽ����п���΢�š�֧����������֧���Ľ�
      --          ���޽���д0
      v_���� := v_���� || ',"selfCashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ֽ�֧��, 1);
      --cashPay  �ֽ�Ԥ������  Number  14,2  ��
      v_���� := v_���� || ',"cashPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ�Ԥ��, 1);
      --chequePay  ֧ƱԤ������  Number  14,2  ��
      v_���� := v_���� || ',"chequePay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧ƱԤ��, 1);
      --transferAccountPay  ת��Ԥ������  Number  14,2  ��
      v_���� := v_���� || ',"transferAccountPay":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת��Ԥ��, 1);
      --cashRecharge  �������(�ֽ�)  Number  14,2  ��
      v_���� := v_���� || ',"cashRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ�֧��, 1);
      --chequeRecharge  �������(֧Ʊ)  Number  14,2  ��
      v_���� := v_���� || ',"chequeRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧Ʊ֧��, 1);
      --transferRecharge  ������ת�ˣ�  Number  14,2  ��
      v_���� := v_���� || ',"transferRecharge":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת��֧��, 1);
      --cashRefund  �˻����(�ֽ�)  Number  14,2  ��
      v_���� := v_���� || ',"cashRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�ֽ��˿�, 1);
      --chequeRefund  �˽����(֧Ʊ)  Number  14,2  ��
      v_���� := v_���� || ',"chequeRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.֧Ʊ�˿�, 1);
      --transferRefund  �˽����(ת��)  Number  14,2  ��
      v_���� := v_���� || ',"transferRefund":' || b_Einvoice_Request_Test.Zljsonstr(c_����.ת���˿�, 1);
      --ownAcBalance  �����˻����  Number  14,2  ��
      v_���� := v_���� || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ʻ����, 1);
      --reimbursementAmt  �����ܽ��  Number  14,2  ��  ҽ������󷵻ص��ܽ��
      v_���� := v_���� || ',"reimbursementAmt":' || b_Einvoice_Request_Test.Zljsonstr(c_����.�����ܶ�, 1);
      --balancedNumber  �����  String  100  ��  ҽ����������ɵĺ���/����Ψһֵ
      v_���� := v_���� || ',"balancedNumber":"' || b_Einvoice_Request_Test.Zljsonstr(c_����.ҽ���������) || '"';
      Exit;
    End Loop;
    -------------------------------------------------------------------------------------------
    --��������
    v_�ɷ����� := Null;
    For c_���� In (Select /*+cardinality(b,10)*/
                  Nvl(c.��������, Nvl(d.��������, '-')) As ��������, Sum(��Ԥ��) As �����ܶ�
                 From ����Ԥ����¼ A, �շ��������� C, (Select ���㷽ʽ, �������� From �շ��������� D Where �����id Is Null) D
                 Where a.����id = n_����id And a.�����id = c.�����id(+) And a.���㷽ʽ = c.���㷽ʽ(+) And a.���㷽ʽ = d.���㷽ʽ(+)
                 Group By Nvl(c.��������, Nvl(d.��������, '-'))
                 Order By ��������)
    
     Loop
      --payChannelCode  ������������  String  10  ��
      If v_�ɷ����� Is Null Then
        v_�ɷ����� := Nvl(v_�ɷ�����, '') || '{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.��������, 0)) || '"';
      Else
        v_�ɷ����� := v_�ɷ����� || ',{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.��������, 0)) || '"';
      End If;
      --payChannelValue  �����������  Number  14,2  ��
      v_�ɷ����� := v_�ɷ����� || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(Nvl(c_����.�����ܶ�, 0), 1) || '}';
    End Loop;
  
    If v_�ɷ����� Is Not Null Then
      --payChannelDetail  ���������б�  String  ����  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö����磺��¼4���������б�
      --        ���A-5,JSON��ʽ�б�
      v_�ɷ����� := ',"payChannelDetail":[' || v_�ɷ����� || ']';
    Else
      v_�ɷ����� := ',"payChannelDetail":[]';
    End If;
  
    -------------------------------------------------------------------------------------------
    --����ҽ����Ϣ
    v_����ҽ����Ϣ := Null;
    --otherMedicalList  ����ҽ����Ϣ�б�  String  ����  ��  ��д����δ֪ҽ����Ϣ���ڵ���Ʊ����������ƴ�ӷ�ʽ��ʾ��
    --            ���A-4,JSON��ʽ�б�
    --  infoNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���������ֵ����1�����β������ظ�
    --  infoName  ҽ����Ϣ����  String  100  ��  ����ñ������ͱ��룬�ɲο���¼7ҽ�����������б�
    --  infoValue  ҽ����Ϣֵ  String  100  ��  ����ñ������
    --  infoOther  ҽ��������Ϣ  String  100  ��  ��ҽ������������
  
    -------------------------------------------------------------------------------------------
    --������չ��Ϣ
    v_������չ��Ϣ := Null;
    --otherInfo  ������չ��Ϣ�б�  String  ����  ��  ��д��Ϣ��Ҫ�ڵ���Ʊ���ϵ�����ʾ��������չ��Ϣ��δ֪��Ϣ��
    --          ���A-3,JSON��ʽ�б�
    --  infoNo  ���  Integer  ����  ��  Ĭ�ϴ�1��ʼ��ÿ���������ֵ����1�����β������ظ�
    --  infoName  ��չ��Ϣ����  String  100  ��
    --  infoValue  ��չ��Ϣֵ  String  500  ��
  
    c_������Ϣ := To_Clob('{' || v_Ʊ����Ϣ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_�ɷ�);
    c_������Ϣ := c_������Ϣ || To_Clob(v_֪ͨ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_������Ϣ);
    c_������Ϣ := c_������Ϣ || To_Clob(v_����);
    c_������Ϣ := c_������Ϣ || To_Clob(v_�ɷ�����);
  
    If v_������չ��Ϣ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_������չ��Ϣ);
    End If;
    If v_����ҽ����Ϣ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_����ҽ����Ϣ);
    End If;
    --  eBillRelateNo  ҵ��Ʊ�ݹ�����  String  32  ��  ��һ��ҵ��������Ҫ����N�ŵ���Ʊ�ݣ���N�ŵ���Ʊ��Ӧ��ֵ����һ�£����ں��ڹ�����ѯ
    c_������Ϣ := c_������Ϣ || To_Clob(',"eBillRelateNo":""');
    If v_������ϸ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_������ϸ);
    Else
      c_������Ϣ := c_������Ϣ || c_������ϸ;
    End If;
  
    If v_��ϸ Is Not Null Then
      c_������Ϣ := c_������Ϣ || To_Clob(v_��ϸ);
    Else
      c_������Ϣ := c_������Ϣ || c_��ϸ;
    End If;
    c_������Ϣ  := c_������Ϣ || To_Clob('}');
    Reqdata_Out := c_������Ϣ;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Mzbalancedata_Create;

  Procedure Get_Depositdata_Create
  (
    Json_In        Varchar2,
    Reqdata_Out    Out Clob,
    Totalmoney_Out Out Number,
    Code_Out       Out Integer,
    Message_Out    Out Varchar2
  ) Is
    ---------------------------------------------------------------------------
    --����:��ȡԤ����Ʊ����
    --���:
    --    Json_In,��ʽ����
    --  input
    --    occasion N 1  Ӧ�ó���:1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨���̶���2
    --    einvoice_id  N,1 ��ǰ����Ʊ��ID
    --    deposit_id N 1  Ԥ��ID
    --    writeoff_id  N 1  ����ID  occasion=2ʱ������Ԥ��id;occasion<>2��ʾ����id
    --����:
    --  ReqData_Out-���ص�ҵ����������
    --  Totalmoney_Out-Ʊ���ܶ�
    --  Code_Out-��ȡ�Ƿ�ɹ���0-ʧ�ܣ�1-�ɹ�
    --  Message_Out ������Ϣ
    ---------------------------------------------------------------------------
    j_Input PLJson;
    j_Json  PLJson;
  
    n_Ӧ�ó��� Number(2);
    n_����id   ����Ԥ����¼.����id%Type;
    n_����id   ����Ԥ����¼.����id%Type;
  
    v_ҵ����ˮ�� Varchar2(50);
  
    v_��Ʊ��   Varchar2(100);
    v_�ɷ����� Varchar2(32767);
    c_������Ϣ Clob; --���շ��صĽ�����Ϣ��
  
    v_Ԥ��     Varchar2(32767);
    v_�����   ҽ�ƿ����.����%Type;
    v_����     ����ҽ�ƿ���Ϣ.����%Type;
    n_Ԥ����� �������.Ԥ�����%Type;
  
    n_ȱʡ�����id Number(18);
    v_����ֵ       Varchar2(100);
    n_Ʊ���ܽ��   ������ü�¼.���ʽ��%Type;
    n_�û�id       ��Ա��.Id%Type;
    v_����Ա���   ��Ա��.���%Type;
    v_����Ա����   ��Ա��.����%Type;
    v_Temp         Varchar2(32767);
    n_���ϴ���     Number(2);
    v_�汾��       Varchar2(30);
    n_����Ʊ��id   ����Ʊ��ʹ�ü�¼.Id%Type;
    Cursor c_Deposit_Rec Is
      Select a.No, a.�տ�ʱ��, a.Ԥ�����, a.�����id, a.����id, a.��ҳid, a.����id, a.�ɿλ, a.��λ������, a.��λ�ʺ�, a.ժҪ, a.���㷽ʽ, a.�������, a.����,
             a.������ˮ��, a.����˵��, a.������λ, a.���, a.����Ա���, a.����Ա����, Nvl(b.����, c.����) As ����, Nvl(b.�Ա�, c.�Ա�) As �Ա�,
             Nvl(b.����, c.����) As ����, c.�����, Nvl(b.סԺ��, c.סԺ��) As סԺ��, c.Email, c.���֤��, c.�ֻ���, 1 As �ɿ�����,
             Decode(Nvl(a.Ԥ�����, 0), 1, '07', '07') As ҵ���ʶ, d.���� As ��Ժ���ұ���, d.���� As ��Ժ��������, e.���� As ��Ժ���ұ���,
             e.���� As ��Ժ��������, b.��Ժ����, b.��Ժ����, Nvl(b.������, b.סԺ��) As ������, j.���� As ҽ�ƿ�����
      From ����Ԥ����¼ A, ������ҳ B, ������Ϣ C, ���ű� D, ���ű� E, ҽ�ƿ���� J
      Where a.Id = n_����id And a.����id = b.����id(+) And a.��ҳid = b.��ҳid(+) And a.����id = c.����id(+) And b.��Ժ����id = d.Id(+) And
            b.��Ժ����id = e.Id(+) And a.�����id = j.Id(+);
    r_Deposit_Rec c_Deposit_Rec%RowType;
  
  Begin
    j_Input := PLJson(Json_In);
    j_Json  := j_Input.Get_Pljson('input');
  
    n_Ӧ�ó���   := Nvl(j_Json.Get_Number('occasion'), 0);
    n_����id     := j_Json.Get_Number('deposit_id');
    n_����id     := Nvl(j_Json.Get_Number('writeoff_id'), 0);
    n_����Ʊ��id := Nvl(j_Json.Get_Number('einvoice_id'), 0);
  
    If Nvl(n_Ӧ�ó���, 0) = 0 Then
      Code_Out    := 0;
      Message_Out := '��Ч��Ӧ�ó���';
      Return;
    End If;
  
    Select Nvl(Max(����ֵ), 'V2.0.3')
    Into v_�汾��
    From �����ӿ�����
    Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '֧�ְ汾';
  
    b_Einvoice_Request_Test.Get_Identity(n_�û�id, v_����Ա���, v_����Ա����);
  
    --v_ҵ���ʶ:01  סԺ,02  ����,03  ����,04  ����,05  �������,06  �Һ�,07  סԺԤ����,08  ���Ԥ����
  
    n_Ʊ���ܽ�� := 0;
    Open c_Deposit_Rec;
    Fetch c_Deposit_Rec
      Into r_Deposit_Rec;
  
    If c_Deposit_Rec%NotFound Then
      Code_Out    := 0;
      Message_Out := 'δ�ҵ�ָ����Ԥ����������';
      Return;
    End If;
    Select Count(1) Into n_���ϴ��� From ����Ʊ��ʹ�ü�¼ Where Ʊ�� = n_Ӧ�ó��� And ����id = n_����id;
  
    Begin
      Select ����, ����
      Into v_�����, v_����
      From (Select b.����id, c.����, c.����, b.����, Decode(b.�����id, n_ȱʡ�����id, 2, c.ȱʡ��־) As ȱʡ��־
             From ����ҽ�ƿ���Ϣ B, ҽ�ƿ���� C
             Where b.�����id = c.Id And b.����id = r_Deposit_Rec.����id
             Order By ȱʡ��־)
      Where Rownum < 2;
    Exception
      When Others Then
        v_���� := Null;
    End;
  
    If v_����� Is Null Then
      If r_Deposit_Rec.���֤�� Is Not Null Then
        Select Nvl(Max(����ֵ), '99998')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '���֤�������ͱ��';
      
        v_����� := v_����ֵ;
        v_����   := r_Deposit_Rec.���֤��;
      Else
        --û��һ�ſ����̶�һ�ֿ����
        Select Nvl(Max(����ֵ), '99999')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ������';
        v_����� := v_����ֵ;
      
        Select Nvl(Max(����ֵ), '-')
        Into v_����ֵ
        From �����ӿ�����
        Where �ӿ��� = '����ģ��ӿ�(V1.0.0)' And ������ = '�����޿��Ŀ���';
        v_���� := v_����ֵ;
      End If;
    End If;
  
    If Nvl(n_����id, 0) = 0 Then
      n_Ʊ���ܽ�� := r_Deposit_Rec.���;
    Else
      Select -1 * Max(��Ԥ��) Into n_Ʊ���ܽ�� From ����Ԥ����¼ Where ID = n_����id;
    End If;
  
    Select Max(Ԥ�����)
    Into n_Ԥ�����
    From �������
    Where ����id = r_Deposit_Rec.����id And ���� = 1 And ���� = r_Deposit_Rec.Ԥ�����;
  
    Totalmoney_Out := n_Ʊ���ܽ��;
  
    If Nvl(n_����id, 0) <> 0 Then
      v_��Ʊ�� := b_Einvoice_Request_Test.Get_Einvoice_Node(v_����Ա���, v_����Ա����, n_�û�id);
      v_Ԥ��   := Null;
      For c_Ʊ�� In (Select ����, ����, ƾ֤����, ƾ֤����
                   From ����Ʊ��ʹ�ü�¼
                   Where ID = n_����id And �˿�id Is Null And ��¼״̬ = 1) Loop
        --�˿Ʊ
        --busType  ҵ���ʶ  String  20  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö��գ����磺��¼5 ҵ���ʶ�б�
        --          07:��ʶסԺԤ����
        --          08:��ʶ���Ԥ����
        v_Ԥ�� := '{"busType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.ҵ���ʶ) || '"';
        --billBatchCode  ����Ʊ�ݴ���  String  50  ��  
        v_Ԥ�� := v_Ԥ�� || ',"billBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_Ʊ��.����) || '"';
        --billNo  ����Ʊ�ݺ���  String  20  ��  
        v_Ԥ�� := v_Ԥ�� || ',"billNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_Ʊ��.����) || '"';
        --reason  ���ԭ��  String  200  ��  
        v_Ԥ�� := v_Ԥ�� || ',"reason":"' || b_Einvoice_Request_Test.Zljsonstr('�˿�') || '"';
        --operator  ������  String  60  ��  
        v_Ԥ�� := v_Ԥ�� || ',"operator":"' || b_Einvoice_Request_Test.Zljsonstr(v_����Ա����) || '"';
        --busDateTime  ҵ����ʱ��  String  17  ��  yyyyMMddHHmmssSSS
        v_Ԥ�� := v_Ԥ�� || ',"busDateTime":"' || To_Char(r_Deposit_Rec.�տ�ʱ��, 'yyyymmddhh24miss') || '000' || '"';
        --placeCode  ��Ʊ�����  String  50  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
        v_Ԥ�� := v_Ԥ�� || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_��Ʊ��) || '"';
        --voucherBatchCode  Ԥ����ƾ֤����  String  50  ��  
        v_Ԥ�� := v_Ԥ�� || ',"voucherBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(c_Ʊ��.ƾ֤����) || '"';
        --voucherNo  Ԥ����ƾ֤����  String  20  ��  
        v_Ԥ�� := v_Ԥ�� || ',"voucherNo":"' || b_Einvoice_Request_Test.Zljsonstr(c_Ʊ��.ƾ֤����) || '"';
        --amt  Ԥ�ɽ��˿���  Number  14,2  ��  
        v_Ԥ�� := v_Ԥ�� || ',"amt":' || b_Einvoice_Request_Test.Zljsonstr(-1 * n_Ʊ���ܽ��, 1);
        --ownAcBalance  Ԥ�ɽ��˻����  Number  14,2  ��  �����˿�֮ǰ���˻����
        v_Ԥ�� := v_Ԥ�� || ',"amt":' || b_Einvoice_Request_Test.Zljsonstr(n_Ԥ�����, 1);
        --remark  ��ע  String  600  ��  
        v_Ԥ�� := v_Ԥ�� || ',"voucherBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.ժҪ) || '"}';
        Exit;
      End Loop;
      If v_Ԥ�� Is Null Then
        Message_Out := 'ԭʼԤ����δ���ߵ���Ʊ��ƾ֤�������п����˿�Ʊ��';
        Code_Out    := 0;
      End If;
      Reqdata_Out := To_Clob(v_Ԥ��);
      Code_Out    := 1;
      Return;
    End If;
  
    -------------------------------------------------------------------------------------------
    --��������
    Select Max(c.��������)
    Into v_Temp
    From �շ��������� C
    Where c.�����id = r_Deposit_Rec.�����id And c.���㷽ʽ = r_Deposit_Rec.���㷽ʽ;
  
    If v_Temp Is Null Then
      Select Max(��������)
      Into v_Temp
      From �շ��������� D
      Where �����id Is Null And ���㷽ʽ = r_Deposit_Rec.���㷽ʽ;
    End If;
  
    --payChannelCode  ������������  String  10  ��
    v_�ɷ����� := ',"payChannelDetail":[{"payChannelCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_Temp) || '"';
    --payChannelValue  �����������  Number  14,2  ��
    v_�ɷ����� := v_�ɷ����� || ',"payChannelValue":' || b_Einvoice_Request_Test.Zljsonstr(n_Ʊ���ܽ��, 1) || '}]';
  
    --����ID||_||����Ʊ��ID
    v_ҵ����ˮ�� := n_����id || '_' || Nvl(n_����Ʊ��id, 0);
    v_��Ʊ��     := b_Einvoice_Request_Test.Get_Einvoice_Node(v_����Ա���, v_����Ա����, n_�û�id);
  
    --busType  ҵ���ʶ  String  20  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö��գ����磺��¼5 ҵ���ʶ�б�
    --          ֵ��
    --          07:��ʶסԺԤ����
    --          08:��ʶ���Ԥ����
    v_Ԥ�� := '"busType":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.ҵ���ʶ) || '"';
    --busNo  Ԥ����ҵ����ˮ��  String  50  ��  ��λ�ڲ�Ψһ
    v_Ԥ�� := v_Ԥ�� || ',"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ����ˮ��) || '"';
    --payer  ��������  String  100  ��
    v_Ԥ�� := v_Ԥ�� || ',"busNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.����) || '"';
    --busDateTime  ҵ����ʱ��  String  17  ��  ��ʽ��yyyyMMddHHmmssSSS
    v_Ԥ�� := v_Ԥ�� || ',"busDateTime":"' ||
            b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Deposit_Rec.�տ�ʱ��, 'yyyymmddhh24miss') || '000') || '"';
    --placeCode  ��Ʊ�����  String  50  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
    v_Ԥ�� := v_Ԥ�� || ',"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_��Ʊ��) || '"';
    --payee  �տ���  String  50  ��
    v_Ԥ�� := v_Ԥ�� || ',"payee":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.����Ա����) || '"';
    --drawee  �ɿ���  String  50  ��  �ɷ�������
    v_Ԥ�� := v_Ԥ�� || ',"drawee":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.����) || '"';
    --author  ������  String  100  ��
    v_Ԥ�� := v_Ԥ�� || ',"author":"' || b_Einvoice_Request_Test.Zljsonstr(v_����Ա����) || '"';
    --tel  �����ֻ�����  String  13  ��  �����ֻ��ţ�����Ҫ����Ԥ����ƾ֤�鼯������֪ͨ�����
    v_Ԥ�� := v_Ԥ�� || ',"tel":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.�ֻ���) || '"';
    --email  ���������ַ  String  100  ��  ���������ַ������Ԥ����ƾ֤�鼯������֪ͨ�����
    v_Ԥ�� := v_Ԥ�� || ',"email":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.Email) || '"';
    --idCardNo  �������֤����  String  20  ��
    v_Ԥ�� := v_Ԥ�� || ',"idCardNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.���֤��) || '"';
    --cardType  ������  String  10  ��  ����Ԥ����ɽɴ��Ӧ�Ŀ����ͣ�����￨���籣����
    --          ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
    --          ���磺��¼3�������б�
    v_Ԥ�� := v_Ԥ�� || ',"cardType":"' || b_Einvoice_Request_Test.Zljsonstr(v_�����) || '"';
    --cardNo  ����  String  50  ��  ���ݿ�������д
    v_Ԥ�� := v_Ԥ�� || ',"cardNo":"' || b_Einvoice_Request_Test.Zljsonstr(v_����) || '"';
    --amt  Ԥ�ɽ���  Number  14,2  ��
    v_Ԥ�� := v_Ԥ�� || ',"amt":' || b_Einvoice_Request_Test.Zljsonstr(n_Ʊ���ܽ��, 1);
    --ownAcBalance  Ԥ�ɽ��˻����  Number  14,2  ��  ���νɴ�֮ǰ���˻����
    v_Ԥ�� := v_Ԥ�� || ',"ownAcBalance":' || b_Einvoice_Request_Test.Zljsonstr(n_Ԥ�����, 1);
    --category  ��Ժ��������  String  200  ��
    v_Ԥ�� := v_Ԥ�� || ',"category":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.��Ժ��������) || '"';
    --categoryCode  ��Ժ���ұ���  String  100  ��
    v_Ԥ�� := v_Ԥ�� || ',"categoryCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.��Ժ���ұ���) || '"';
    --inHospitalDate  ��Ժ����  String  10  ��  ��ʽ:yyyy-MM-dd
    v_Ԥ�� := v_Ԥ�� || ',"inHospitalDate":"' || b_Einvoice_Request_Test.Zljsonstr(To_Char(r_Deposit_Rec.��Ժ����, 'yyyy-mm-dd')) || '"';
    --hospitalNo  ����סԺ��  String  20  ��  ����Ժ����Ժ�������������̵�Ψһ��
    v_Ԥ�� := v_Ԥ�� || ',"hospitalNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.סԺ��) || '"';
    --visitNo  סԺ������  String  20  ��  סԺ�ڼ䣬���ڶ�ν��㣬��������������һ��סԺ�����ţ����޾����ţ��ɵ��ڻ���סԺ��
    v_Ԥ�� := v_Ԥ�� || ',"visitNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.סԺ��) || '"';
    --patientId  ����ΨһID  String  50  ��  ������ҵ��ϵͳ�е�Ψһ��ʶID���������֤���롣
    v_Ԥ�� := v_Ԥ�� || ',"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.����id) || '"';
    --patientNo  ���߾�����  String  20  ��  ����ÿ�ξ���һ�ξ����ɵ�һ���µı�š������ߵǼǺţ�
    v_Ԥ�� := v_Ԥ�� || ',"patientNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.��ҳid) || '"';
    --caseNumber  ������  String  50  ��  �������
    v_Ԥ�� := v_Ԥ�� || ',"caseNumber":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.������) || '"';
    --payChannelDetail  ���������б�  String  ����  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö����磺��¼4���������б�
    --          ���A-1,JSON��ʽ�б�
    --payChannelCode  ������������  String  10  ��
    --payChannelValue  �����������  Number  14,2  ��
    v_Ԥ�� := v_Ԥ�� || v_�ɷ�����;
    --accountName  �˻�����  String  200  ��  ������д����ɷ����������п�
    v_Ԥ�� := v_Ԥ�� || ',"accountName":"' || b_Einvoice_Request_Test.Zljsonstr(Nvl(r_Deposit_Rec.ҽ�ƿ�����, '')) || '"';
    --accountNo  �˻�����  String  200  ��  ������д����ɷ����������п�
    v_Ԥ�� := v_Ԥ�� || ',"accountNo":"' ||
            b_Einvoice_Request_Test.Zljsonstr(Nvl(r_Deposit_Rec.����, Nvl(r_Deposit_Rec.��λ�ʺ�, ''))) || '"';
    --accountBank  �˻�������  String  200  ��  ������д����ɷ����������п�
    v_Ԥ�� := v_Ԥ�� || ',"accountBank":"' ||
            b_Einvoice_Request_Test.Zljsonstr(Nvl(r_Deposit_Rec.ҽ�ƿ�����, Nvl(r_Deposit_Rec.��λ������, ''))) || '"';
    --remark  ��ע  String  600  ��
    v_Ԥ�� := v_Ԥ�� || ',"remark":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.ժҪ) || '"';
    If v_�汾�� = 'V3.1.0' Then
      --workUnit  ������λ���ַ      String  200  ��  �ɿ��˵Ĺ�����λ���ַ
      v_Ԥ�� := v_Ԥ�� || ',"workUnit":"' || b_Einvoice_Request_Test.Zljsonstr(r_Deposit_Rec.�ɿλ) || '"';
    
    End If;
    c_������Ϣ  := To_Clob('{' || v_Ԥ�� || '}');
    Reqdata_Out := c_������Ϣ;
    Code_Out    := 1;
  Exception
    When Others Then
      Message_Out := SQLCode || ':' || SQLErrM;
      Code_Out    := 0;
  End Get_Depositdata_Create;

  Function Einvoice_Create
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number Is
    Pragma Autonomous_Transaction;
    ---------------------------------------------------------------------------
    --����:���е���Ʊ�ݿ���
    --���:  
    --    ҵ�񳡾�_In- 1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨ 
    --    ����id_In-ҵ�񳡾�_In=2,Ԥ��ID;ҵ�񳡾�_In<>2:����ID
    --    ����ID_In- ����ID  ҵ�񳡾�_In=2ʱ������Ԥ��id;ҵ�񳡾�_In<>2��ʾ����id
    --����: 
    --  ������Ϣ_Out-����=0ʱ�����ش��� 
    --����:
    --   1-��Ʊ�ɹ�;0-ʧ��
    ---------------------------------------------------------------------------
  
    n_����Ʊ��id ����Ʊ��ʹ�ü�¼.Id%Type;
    v_����       Varchar2(100);
  
    n_����id     ������Ϣ.����id%Type;
    v_�Ա�       ������Ϣ.�Ա�%Type;
    v_����       ������Ϣ.���� %Type;
    n_�����     ������Ϣ.�����%Type;
    n_סԺ��     ������Ϣ.סԺ��%Type;
    n_Find       Number(2);
    n_Ʊ�ݽ��   ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type;
    v_��Ʊ��     ����Ʊ��ʹ�ü�¼.��Ʊ��%Type;
    v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
    v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
    v_ƾ֤У���� ����Ʊ��ʹ�ü�¼.ƾ֤������%Type;
    v_Ʊ�ݴ���   ����Ʊ��ʹ�ü�¼.����%Type;
    v_Ʊ�ݺ���   ����Ʊ��ʹ�ü�¼.����%Type;
    v_Ʊ��У���� ����Ʊ��ʹ�ü�¼.������%Type;
    v_ϵͳ��Դ   ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type;
    v_����ʱ��   Varchar2(20);
    c_��ά��     Clob;
    v_Url        ����Ʊ��ʹ�ü�¼.Url����%Type;
    v_����url    ����Ʊ��ʹ�ü�¼.Url����%Type;
  
    v_����Ա���   ��Ա��.���%Type;
    v_����Ա����   ��Ա��.����%Type;
    n_�û�id       ��Ա��.Id%Type;
    v_Req_Json     Varchar2(32767);
    c_Req_Data     Clob;
    v_Err_Msg      Varchar2(4000);
    n_Code         Number(2);
    n_�Ƿ�����     Number(2);
    v_Service_Name Varchar2(100);
    v_Version      Varchar2(20);
    v_Respdata     Varchar2(32767); --��Ӧ����
    j_Json         PLJson;
  Begin
  
    If Nvl(ҵ�񳡾�_In, 0) < 1 Or Nvl(ҵ�񳡾�_In, 0) > 5 Then
      ������Ϣ_Out := '����ʶ���ҵ��!';
      Return 0;
    End If;
    n_Find := 1;
    If ҵ�񳡾�_In = 1 Or ҵ�񳡾�_In = 4 Then
      --�շѼ��Һ�
      Begin
        Select a.����id, Nvl(a.����, b.����) As ����, Nvl(a.����, b.����) As ����, Nvl(a.�Ա�, b.�Ա�) As �Ա�, b.�����, b.סԺ�� As סԺ��
        Into n_����id, v_����, v_�Ա�, v_����, n_�����, n_סԺ��
        From ������ü�¼ A, ������Ϣ B
        Where a.����id = ����id_In And a.����id = b.����id(+) And Rownum < 2;
      Exception
        When Others Then
          n_Find := 0;
      End;
    End If;
    If ҵ�񳡾�_In = 2 Then
      --Ԥ��
      Begin
        Select a.����id, Nvl(c.����, b.����) As ����, Nvl(c.����, b.����) As ����, Nvl(c.�Ա�, b.�Ա�) As �Ա�, b.�����,
               Nvl(c.סԺ��, b.סԺ��) As סԺ��, Nvl(Ԥ�����, 2)
        Into n_����id, v_����, v_�Ա�, v_����, n_�����, n_סԺ��, n_�Ƿ�����
        From ����Ԥ����¼ A, ������Ϣ B, ������ҳ C
        Where a.Id = ����id_In And a.����id = b.����id And a.����id = c.����id(+) And a.��ҳid = c.��ҳid(+);
      Exception
        When Others Then
          n_Find := 0;
      End;
    End If;
    --�ݲ�֧����Ԥ��
    If Nvl(n_�Ƿ�����, 0) = 1 Then
      Return 1;
    End If;
    If ҵ�񳡾�_In = 3 Then
      --����
      Begin
        Select a.����id, Decode(Nvl(a.����id, 0), 0, a.ԭ��, Nvl(c.����, b.����)) As ����, Nvl(c.����, b.����) As ����,
               Nvl(c.�Ա�, b.�Ա�) As �Ա�, b.�����, Nvl(c.סԺ��, b.סԺ��) As סԺ��, Nvl(��������, 2) As n_�Ƿ�����
        Into n_����id, v_����, v_�Ա�, v_����, n_�����, n_סԺ��, n_�Ƿ�����
        From ���˽��ʼ�¼ A, ������Ϣ B, ������ҳ C
        Where a.Id = ����id_In And a.����id = b.����id And a.����id = c.����id(+) And a.��ҳid = c.��ҳid(+);
      Exception
        When Others Then
          n_Find := 0;
      End;
    End If;
  
    If ҵ�񳡾�_In = 5 Then
      --ҽ�ƿ�
      Begin
        Select a.����id, Nvl(a.����, b.����) As ����, Nvl(a.����, b.����) As ����, Nvl(a.�Ա�, b.�Ա�) As �Ա�, b.�����, b.סԺ�� As סԺ��
        Into n_����id, v_����, v_�Ա�, v_����, n_�����, n_סԺ��
        From סԺ���ü�¼ A, ������Ϣ B
        Where a.����id = ����id_In And a.����id = b.����id(+) And Rownum < 2;
      Exception
        When Others Then
          n_Find := 0;
      End;
    End If;
  
    n_Ʊ�ݽ�� := 0;
    If Nvl(n_Find, 0) = 0 Then
      --δ�ҵ�ԭʼ��������
      ������Ϣ_Out := 'δ�ҵ���Ҫ���ߵ���Ʊ�ݵĽ�������!';
      Return 0;
    End If;
  
    Select ����Ʊ��ʹ�ü�¼_Id.Nextval Into n_����Ʊ��id From Dual;
    If Nvl(n_����id, 0) = 0 Then
      n_����id := Null;
    End If;
    b_Einvoice_Request_Test.Get_Identity(n_�û�id, v_����Ա���, v_����Ա����);
    v_��Ʊ�� := b_Einvoice_Request_Test.Get_Einvoice_Node(v_����Ա���, v_����Ա����, n_�û�id);
  
    --1.�ȴ������Ʊ��
    Zl_����Ʊ��ʹ�ü�¼_Insert(n_����Ʊ��id, ҵ�񳡾�_In, ����id_In, n_����id, v_����, v_�Ա�, v_����, n_�����, n_סԺ��, n_Ʊ�ݽ��, v_��Ʊ��, v_ϵͳ��Դ, Null,
                       '', v_����Ա���, v_����Ա����, Sysdate);
  
    --    occasion N 1  Ӧ�ó���:1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    v_Req_Json := '"occasion":' || b_Einvoice_Request_Test.Zljsonstr(ҵ�񳡾�_In, 1);
    --    einvoice_id  N,1 ��ǰ����Ʊ��ID
    v_Req_Json := v_Req_Json || ',"einvoice_id":' || b_Einvoice_Request_Test.Zljsonstr(n_����Ʊ��id, 1);
    If ҵ�񳡾�_In = 2 Then
      --deposit_id N 1  Ԥ��ID
      v_Req_Json := v_Req_Json || ',"deposit_id":' || b_Einvoice_Request_Test.Zljsonstr(����id_In, 1);
    Else
      --balance_id N 1  ����ID  occasion=2ʱ��Ԥ��id;occasion<>2��ʾ����id
      v_Req_Json := v_Req_Json || ',"balance_id":' || b_Einvoice_Request_Test.Zljsonstr(����id_In, 1);
    End If;
    --    writeoff_id  N 1  ����ID  occasion=2ʱ������Ԥ��id;occasion<>2��ʾ����id
    v_Req_Json := v_Req_Json || ',"writeoff_id":' || b_Einvoice_Request_Test.Zljsonstr(����id_In, 1);
    v_Req_Json := '{"input":{' || v_Req_Json || '}}';
  
    --2.��ȡ����Ʊ��
    If ҵ�񳡾�_In = 1 Then
      --�շ�
      b_Einvoice_Request_Test.Get_Chargedata_Create(v_Req_Json, c_Req_Data, n_Ʊ�ݽ��, n_Code, v_Err_Msg);
      v_Service_Name := 'invoiceEBillOutpatient';
      v_Version      := '1.0';
    Elsif ҵ�񳡾�_In = 2 Then
      --Ԥ��
    
      b_Einvoice_Request_Test.Get_Depositdata_Create(v_Req_Json, c_Req_Data, n_Ʊ�ݽ��, n_Code, v_Err_Msg);
      If Nvl(����id_In, 0) <> 0 Then
        --�˿�
        v_Service_Name := 'writeOffPayMentVoucher';
      Else
        v_Service_Name := 'invoicePayMentVoucher';
      End If;
      v_Version := '1.0';
    Elsif ҵ�񳡾�_In = 3 And Nvl(n_�Ƿ�����, 0) = 1 Then
      --�������
      b_Einvoice_Request_Test.Get_Mzbalancedata_Create(v_Req_Json, c_Req_Data, n_Ʊ�ݽ��, n_Code, v_Err_Msg);
      v_Service_Name := 'invoiceEBillOutpatient';
      v_Version      := '1.0';
    Elsif ҵ�񳡾�_In = 3 And Nvl(n_�Ƿ�����, 0) <> 1 Then
      --סԺ����
      b_Einvoice_Request_Test.Get_Zybalancedata_Create(v_Req_Json, c_Req_Data, n_Ʊ�ݽ��, n_Code, v_Err_Msg);
      v_Service_Name := 'invEBillHospitalized';
      v_Version      := '1.0';
    Elsif ҵ�񳡾�_In = 4 Then
      --�Һ�
      b_Einvoice_Request_Test.Get_Sendcarddata_Create(v_Req_Json, c_Req_Data, n_Ʊ�ݽ��, n_Code, v_Err_Msg);
      v_Service_Name := 'invEBillRegistration';
      v_Version      := '1.0';
    Elsif ҵ�񳡾�_In = 5 Then
      --����
      b_Einvoice_Request_Test.Get_Registerdata_Create(v_Req_Json, c_Req_Data, n_Ʊ�ݽ��, n_Code, v_Err_Msg);
      v_Service_Name := 'invoiceEBillOutpatient';
      v_Version      := '1.0';
    End If;
  
    If n_Code = 0 Then
      Rollback;
      ������Ϣ_Out := v_Err_Msg;
      Return 0;
    End If;
  
    --����ҵ������
    n_Code := b_Einvoice_Request_Test.Request(c_Req_Data, v_Service_Name, v_Respdata, v_Err_Msg, v_Version);
    If n_Code = 0 Then
      ������Ϣ_Out := v_Err_Msg;
      Rollback;
      Return 0;
    End If;
    --��������
    j_Json := PLJson(v_Respdata);
    If Nvl(ҵ�񳡾�_In, 0) = 2 Then
      If Nvl(����id_In, 0) <> 0 Then
      
        --voucherBatchCode  Ԥ�����Ʊƾ֤����  String  50  ��  
        v_ƾ֤���� := j_Json.Get_String('voucherBatchCode');
        --voucherNo  Ԥ�����Ʊƾ֤����  String  20  ��  
        v_ƾ֤���� := j_Json.Get_String('voucherNo');
        --voucherRandom  Ԥ�����Ʊƾ֤У����  String  20  ��  
        v_ƾ֤У���� := j_Json.Get_String('voucherRandom');
        --eScarletBillBatchCode  ���Ӻ�ƱƱ�ݴ���  String  50  ��  
        v_Ʊ�ݴ��� := j_Json.Get_String('eScarletBillBatchCode');
        --eScarletBillNo  ���Ӻ�ƱƱ�ݺ���  String  20  ��  
        v_Ʊ�ݺ��� := j_Json.Get_String('eScarletBillNo');
        --eScarletRandom  ���Ӻ�ƱƱ��У����  String  20  ��  
        v_Ʊ��У���� := j_Json.Get_String('eScarletRandom');
        --createTime  ���Ӻ�Ʊ����ʱ��  String  17  ��  ����ʱ�䣺ʱ���ʽ��ȷ������yyyyMMddHHmmssSSS
        v_����ʱ�� := j_Json.Get_String('createTime');
        --billQRCode  ���Ӻ�Ʊ��ά��ͼƬ����  String  ����  ��  ��ֵ��Base64���룬����ʱ��ҪBase64���룬ͼƬ��ʽΪPNG
        c_��ά�� := j_Json.Get_Clob('billQRCode');
        --pictureUrl  ����Ʊ��H5ҳ��URL  String  ����  ��  
        v_Url := j_Json.Get_String('pictureUrl');
      Else
        --Ԥ��
        --voucherBatchCode  Ԥ����ƾ֤����  String  50  ��  
        v_ƾ֤���� := j_Json.Get_String('voucherBatchCode');
        --voucherNo  Ԥ����ƾ֤����  String  20  ��  
        v_ƾ֤���� := j_Json.Get_String('voucherNo');
        --voucherRandom  Ԥ����ƾ֤У����  String  20  ��  
        v_ƾ֤У���� := j_Json.Get_String('voucherRandom');
        --billBatchCode  ����Ʊ�ݴ���  String  50  ��  
        v_Ʊ�ݴ��� := j_Json.Get_String('billBatchCode');
        --billNo  ����Ʊ�ݺ���  String  20  ��  
        v_Ʊ�ݺ��� := j_Json.Get_String('billNo');
        --random  ����У����  String  20  ��  
        v_Ʊ��У���� := j_Json.Get_String('random');
        --createTime  ����Ʊ������ʱ��  String  17  ��  ����ʱ�䣺ʱ���ʽ��ȷ������yyyyMMddHHmmssSSS
        v_����ʱ�� := j_Json.Get_String('createTime');
        --billQRCode  ����Ʊ�ݶ�ά��ͼƬ����  String  ����  ��  ��ֵ��Base64���룬����ʱ��ҪBase64���룬ͼƬ��ʽΪPNG
        c_��ά�� := j_Json.Get_Clob('billQRCode');
        --pictureUrl  ����Ʊ��H5ҳ��URL  String  ����  ��  
        v_Url := j_Json.Get_String('pictureUrl');
      End If;
    Else
      --����
      --billBatchCode  ����Ʊ�ݴ���  String  50  ��  
      v_Ʊ�ݴ��� := j_Json.Get_String('billBatchCode');
      --billNo  ����Ʊ�ݺ���  String  20  ��  
      v_Ʊ�ݺ��� := j_Json.Get_String('billNo');
      --random  ����У����  String  20  ��  
      v_Ʊ��У���� := j_Json.Get_String('random');
      --createTime  ����Ʊ������ʱ��  String  17  ��  ����ʱ�䣺ʱ���ʽ��ȷ������yyyyMMddHHmmssSSS
      v_����ʱ�� := j_Json.Get_String('createTime');
      --billQRCode  ����Ʊ�ݶ�ά��ͼƬ����  String  ����  ��  ��ֵ��Base64���룬����ʱ��ҪBase64���룬ͼƬ��ʽΪPNG
      c_��ά�� := j_Json.Get_Clob('billQRCode');
      --pictureUrl  ����Ʊ��H5ҳ��URL  String  ����  ��  
      v_Url := j_Json.Get_String('pictureUrl');
      --pictureNetUrl  ����Ʊ������H5ҳ��URL  String  ����  ��  ��������
      v_����url := j_Json.Get_String('pictureNetUrl');
    End If;
  
    --���µ���Ʊ����Ϣ
    Update ����Ʊ��ʹ�ü�¼
    Set ���� = v_Ʊ�ݴ���, ���� = v_Ʊ�ݺ���, ������ = v_Ʊ��У����, ����ʱ�� = v_����ʱ��, Url���� = v_Url, Url���� = v_����url, ϵͳ��Դ = '', Ʊ�ݽ�� = n_Ʊ�ݽ��,
        ƾ֤���� = v_ƾ֤����, ƾ֤���� = v_ƾ֤����, ƾ֤������ = v_ƾ֤У����
    Where ID = n_����Ʊ��id;
    --�����ά��
    Insert Into ����Ʊ�ݶ�ά�� (ʹ�ü�¼id, ��ά��) Values (n_����Ʊ��id, c_��ά��);
    Commit;
    Return 1;
  Exception
    When Others Then
      ������Ϣ_Out := SQLCode || ':' || SQLErrM;
      Rollback;
      Return 0;
  End Einvoice_Create;

  --����Ʊ�����ϼ��
  Function Einvoice_Cancel_Check
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number Is
    ---------------------------------------------------------------------------
    --����:���е���Ʊ�ݳ����
    --���:  
    --    ҵ�񳡾�_In- 1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨ 
    --    ����id_In-ҵ�񳡾�_In=2,Ԥ��ID;ҵ�񳡾�_In<>2:����ID 
    --����: 
    --  ������Ϣ_Out-����=0ʱ�����ش��� 
    --����:
    --   1-��Ʊ�Ϸ�;0-��Ʊ���Ϸ�
    ---------------------------------------------------------------------------
  
    v_Req_Data     Varchar2(32767);
    v_Err_Msg      Varchar2(4000);
    n_Code         Number(2);
    v_Service_Name Varchar2(100);
    v_Version      Varchar2(20);
    v_Respdata     Varchar2(32767); --��Ӧ����
    j_Json         PLJson;
  
    Cursor c_Einvoice Is
      Select a.Id, Nvl(a.�Ƿ񻻿�, 0) As �Ƿ񻻿�, a.ֽ�ʷ�Ʊ��, a.����, a.����, a.������, a.����ʱ��
      From ����Ʊ��ʹ�ü�¼ A
      Where a.Id = ����id_In And a.��¼״̬ = 1 And Ʊ�� = ҵ�񳡾�_In;
    r_Einvoice c_Einvoice%RowType;
  
  Begin
  
    If Nvl(ҵ�񳡾�_In, 0) < 1 Or Nvl(ҵ�񳡾�_In, 0) > 5 Then
      ������Ϣ_Out := '����ʶ���ҵ��!';
      Return 0;
    End If;
    If ҵ�񳡾�_In = 2 Then
      --Ԥ����û����ؼ��ӿڣ�ֱ�ӷ���1
      Return 1;
    End If;
    Open c_Einvoice;
    Fetch c_Einvoice
      Into r_Einvoice;
    If c_Einvoice%NotFound Then
      --�޵���Ʊ���������;�����ˣ�ֱ�ӷ���1
    
      Return 1;
    End If;
  
    If r_Einvoice.�Ƿ񻻿� = 1 Then
      --�Ѿ�����ֽ��Ʊ�ݣ�������������
      ������Ϣ_Out := '�Ѿ�����ֽ�ʷ�Ʊ' || Nvl(r_Einvoice.ֽ�ʷ�Ʊ��, '') || '����ֹ�Ե���Ʊ�ݽ��г�����!';
      Return 0;
    End If;
  
    v_Service_Name := 'getEBillStatesByBillInfo';
    v_Version      := '1.0';
  
    --billBatchCode  ����Ʊ�ݴ���  String  50  ��  
    v_Req_Data := '{' || '"billBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����) || '"';
    --billNo  ����Ʊ�ݺ���  String  20  ��  
    v_Req_Data := v_Req_Data || ',' || '"billNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����) || '"}';
  
    --����ҵ������
    n_Code := b_Einvoice_Request_Test.Request(v_Req_Data, v_Service_Name, v_Respdata, v_Err_Msg, v_Version);
    If n_Code = 0 Then
      ������Ϣ_Out := v_Err_Msg;
      Return 0;
    End If;
    --��������
    j_Json := PLJson(v_Respdata);
    --state  ״̬  String  1  ��  ״̬��1������2����
    If j_Json.Get_String('state') = '2' Then
      --�����˵ģ��������¿���
      Return 1;
    End If;
    --isScarlet  �Ƿ��ѿ���Ʊ  String  1  ��  0δ����Ʊ��1�ѿ���Ʊ
    If j_Json.Get_String('isScarlet') = '1' Then
      --�Ѿ����ߺ�Ʊ�������ٽ��п���
      Return 1;
    End If;
    --isPrtPaper  �Ƿ��ӡֽ��Ʊ��  String  1  ��  0δ��ӡ��1�Ѵ�ӡ
    If j_Json.Get_String('state') = '1' Then
      ������Ϣ_Out := '�Ѿ���ӡֽ��Ʊ�ݣ����������ϲ���!';
      Return 0;
    End If;
    If b_Einvoice_Request_Test.Get_Version <> '3.1.0' Then
      --�����ʽӿ�
      Return 1;
    End If;
  
    --4.1.16  ��ѯ����Ʊ������״̬�ӿ�
    v_Service_Name := 'getEBillStatesByBillInfo';
    v_Version      := '1.0';
  
    --billBatchCode  ����Ʊ�ݴ���  String  50  ��  ֵΪ���߽ӿڷ��صĵ���Ʊ�ݴ���(�������)
    v_Req_Data := '{' || '"billBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����) || '"';
  
    --billNo  ����Ʊ�ݺ�  String  20  ��  
    v_Req_Data := v_Req_Data || ',' || '"billNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����) || '"';
    --random  ����У����  String  20  ��  
    v_Req_Data := v_Req_Data || ',' || '"random":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����) || '"';
    --createTime  ����Ʊ������ʱ��  String  17  ��  ���ߵ���Ʊ�ݷ��ص�����ʱ�䣺yyyyMMddHHmmssSSS
    v_Req_Data := v_Req_Data || ',' || '"createTime":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����ʱ��) || '"}';
  
    --����ҵ������
    n_Code := b_Einvoice_Request_Test.Request(v_Req_Data, v_Service_Name, v_Respdata, v_Err_Msg, v_Version);
    If n_Code = 0 Then
      ������Ϣ_Out := v_Err_Msg;
      Return 0;
    End If;
    --��������
    j_Json := PLJson(v_Respdata);
    --state  ����״̬  String  1  ��  0δ���ˣ�1������
  
    If j_Json.Get_String('state') = '1' Then
      ������Ϣ_Out := '�õ���Ʊ���Ѿ����ʣ������������ϲ���';
      Return 0;
    End If;
    Return 1;
  Exception
    When Others Then
      ������Ϣ_Out := SQLCode || ':' || SQLErrM;
      Return 0;
  End Einvoice_Cancel_Check;

  --����Ʊ������
  Function Einvoice_Cancel
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number Is
    ---------------------------------------------------------------------------
    --����:���е���Ʊ�ݳ��
    --���:  
    --    ҵ�񳡾�_In- 1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨ 
    --    ����id_In-ҵ�񳡾�_In=2,Ԥ��ID;ҵ�񳡾�_In<>2:����ID 
    --����: 
    --  ������Ϣ_Out-����=0ʱ�����ش��� 
    --����:
    --   1-��Ʊ�Ϸ�;0-��Ʊ���Ϸ�
    ---------------------------------------------------------------------------
  
    v_Req_Data     Varchar2(32767);
    v_Err_Msg      Varchar2(4000);
    n_Code         Number(2);
    v_Service_Name Varchar2(100);
    v_Version      Varchar2(20);
    v_Respdata     Varchar2(32767); --��Ӧ����
    j_Json         PLJson;
    n_��Աid       ��Ա��.Id%Type;
    v_����Ա���   ��Ա��.���%Type;
    v_����Ա����   ��Ա��.����%Type;
    v_ҵ����ʱ�� Varchar2(30);
  
    v_��Ʊ����     ����Ʊ��ʹ�ü�¼.����%Type;
    v_��Ʊ����     ����Ʊ��ʹ�ü�¼.����%Type;
    v_��ƱУ����   ����Ʊ��ʹ�ü�¼.������%Type;
    v_ϵͳ��Դ     ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type;
    c_��Ʊ��ά��   Clob;
    v_��Ʊurl      ����Ʊ��ʹ�ü�¼.Url����%Type;
    v_��Ʊ����url  ����Ʊ��ʹ�ü�¼.Url����%Type;
    v_��Ʊ��       ����Ʊ��ʹ�ü�¼.��Ʊ��%Type;
    v_��Ʊ����ʱ�� ����Ʊ��ʹ�ü�¼.����ʱ��%Type;
    v_ԭ��         Varchar2(50);
    v_ժҪ         ����Ԥ����¼.ժҪ%Type;
    n_����id       ����Ʊ��ʹ�ü�¼.Id%Type;
    Cursor c_Einvoice Is
      Select a.Id, Nvl(a.�Ƿ񻻿�, 0) As �Ƿ񻻿�, a.ֽ�ʷ�Ʊ��, a.����, a.����, a.������, a.����ʱ��, a.����id, a.סԺ��
      From ����Ʊ��ʹ�ü�¼ A
      Where a.Id = ����id_In And a.��¼״̬ = 1 And Ʊ�� = ҵ�񳡾�_In;
    r_Einvoice c_Einvoice%RowType;
  
  Begin
  
    If Nvl(ҵ�񳡾�_In, 0) < 1 Or Nvl(ҵ�񳡾�_In, 0) > 5 Then
      ������Ϣ_Out := '����ʶ���ҵ��!';
      Return 0;
    End If;
  
    Open c_Einvoice;
    Fetch c_Einvoice
      Into r_Einvoice;
    If c_Einvoice%NotFound Then
      --�޵���Ʊ���������;�����ˣ�ֱ�ӷ���1
      Return 1;
    End If;
  
    n_Code := b_Einvoice_Request_Test.Einvoice_Cancel_Check(ҵ�񳡾�_In, ����id_In, ������Ϣ_Out);
    If n_Code = 0 Then
      --ʧ�ܣ�ֱ���˳�
      Return n_Code;
    End If;
    b_Einvoice_Request_Test.Get_Identity(n_��Աid, v_����Ա���, v_����Ա����);
    v_��Ʊ�� := b_Einvoice_Request_Test.Get_Einvoice_Node(v_����Ա���, v_����Ա����, n_��Աid);
    n_Code   := 1;
  
    If ҵ�񳡾�_In = 1 Then
      v_ԭ�� := '�˷�';
      Begin
        Select To_Char(�Ǽ�ʱ��, 'yyyymmddhh24miss') || '000'
        Into v_ҵ����ʱ��
        From ������ü�¼
        Where ����id = ����id_In And Rownum < 2;
      Exception
        When Others Then
          n_Code := 0;
      End;
    Elsif ҵ�񳡾�_In = 2 Then
      v_ԭ�� := '��Ԥ��';
      Begin
        Select To_Char(�տ�ʱ��, 'yyyymmddhh24miss') || '000'
        Into v_ҵ����ʱ��
        From ����Ԥ����¼
        Where ID = ����id_In And Rownum < 2;
      Exception
        When Others Then
          n_Code := 0;
      End;
    Elsif ҵ�񳡾�_In = 3 Then
      v_ԭ�� := '��������';
      Begin
        Select To_Char(�շ�ʱ��, 'yyyymmddhh24miss') || '000'
        Into v_ҵ����ʱ��
        From ���˽��ʼ�¼
        Where ID = ����id_In And Rownum < 2;
      Exception
        When Others Then
          n_Code := 0;
      End;
    
    Elsif ҵ�񳡾�_In = 4 Then
      v_ԭ�� := '�˺�';
      Begin
        Select To_Char(�Ǽ�ʱ��, 'yyyymmddhh24miss') || '000'
        Into v_ҵ����ʱ��
        From ������ü�¼
        Where ����id = ����id_In And Rownum < 2;
      Exception
        When Others Then
          n_Code := 0;
      End;
    Elsif ҵ�񳡾�_In = 5 Then
      v_ԭ�� := '�˿�';
      Begin
        Select To_Char(�Ǽ�ʱ��, 'yyyymmddhh24miss') || '000'
        Into v_ҵ����ʱ��
        From סԺ���ü�¼
        Where ����id = ����id_In And Rownum < 2;
      Exception
        When Others Then
          n_Code := 0;
      End;
    End If;
  
    If n_Code = 0 Then
      ������Ϣ_Out := 'δ�ҵ�ԭʼ��������!';
      Return n_Code;
    End If;
  
    If ҵ�񳡾�_In = 2 Then
      --����Ԥ��
      v_Service_Name := 'cancelPayMentVoucherBalance';
      v_Version      := '1.0';
    Else
      v_Service_Name := 'writeOffEBill';
      v_Version      := '1.0';
    End If;
  
    If ҵ�񳡾�_In = 2 Then
      Select Max(ժҪ) Into v_ժҪ From ����Ԥ����¼ Where ID = ����id_In;
    
      --billBatchCode  ����Ʊ�ݴ���  String  50  ��  
      v_Req_Data := '{' || '"billBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����) || '"';
      --billNo  ����Ʊ�ݺ���  String  20  ��  
      v_Req_Data := v_Req_Data || ',' || '"billNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����) || '"';
      --reason  ���ԭ��  String  200  ��  
      v_Req_Data := v_Req_Data || ',' || '"reason":"' || b_Einvoice_Request_Test.Zljsonstr(v_ԭ��) || '"';
      --operator  ������  String  60  ��  
      v_Req_Data := v_Req_Data || ',' || '"operator":"' || b_Einvoice_Request_Test.Zljsonstr(v_����Ա����) || '"';
      --busDateTime  ҵ����ʱ��  String  17  ��  yyyyMMddHHmmssSSS
      v_Req_Data := v_Req_Data || ',' || '"busDateTime":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ����ʱ��) || '"';
      --placeCode  ��Ʊ�����  String  50  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
      v_Req_Data := v_Req_Data || ',' || '"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_��Ʊ��) || '"';
      --patientId  ����ΨһID  String  50  ��  
      v_Req_Data := v_Req_Data || ',' || '"patientId":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����id) || '"';
      --hospitalNo  ����סԺ��  String  20  ��  
      v_Req_Data := v_Req_Data || ',' || '"hospitalNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.סԺ��) || '"';
      --remark  ��ע  String  600  ��  
      v_Req_Data := v_Req_Data || ',' || '"remark":"' || b_Einvoice_Request_Test.Zljsonstr(v_ժҪ) || '"}';
    Else
      --billBatchCode  ����Ʊ�ݴ���  String  50  ��  
      v_Req_Data := '{' || '"billBatchCode":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����) || '"';
      --billNo  ����Ʊ�ݺ���  String  20  ��  
      v_Req_Data := v_Req_Data || ',' || '"billNo":"' || b_Einvoice_Request_Test.Zljsonstr(r_Einvoice.����) || '"';
      --reason  ���ԭ��  String  200  ��  
      v_Req_Data := v_Req_Data || ',' || '"reason":"' || b_Einvoice_Request_Test.Zljsonstr(v_ԭ��) || '"';
      --operator  ������  String  60  ��  
      v_Req_Data := v_Req_Data || ',' || '"operator":"' || b_Einvoice_Request_Test.Zljsonstr(v_����Ա����) || '"';
    
      --busDateTime  ҵ����ʱ��  String  17  ��  yyyyMMddHHmmssSSS
      v_Req_Data := v_Req_Data || ',' || '"busDateTime":"' || b_Einvoice_Request_Test.Zljsonstr(v_ҵ����ʱ��) || '"';
      --placeCode  ��Ʊ�����  String  50  ��  ֱ����дҵ��ϵͳ�ڲ�����ֵ����ҽ��ƽ̨���ö���
      v_Req_Data := v_Req_Data || ',' || '"placeCode":"' || b_Einvoice_Request_Test.Zljsonstr(v_��Ʊ��) || '"}';
    End If;
  
    Select ����Ʊ��ʹ�ü�¼_Id.Nextval Into n_����id From Dual;
    --�ȳ���
    Zl_����Ʊ��ʹ�ü�¼_Delete(n_����id, v_��Ʊ��, v_ϵͳ��Դ, Null, '', v_����Ա���, v_����Ա����, Sysdate, r_Einvoice.Id);
  
    --����ҵ������
    n_Code := b_Einvoice_Request_Test.Request(v_Req_Data, v_Service_Name, v_Respdata, v_Err_Msg, v_Version);
    If n_Code = 0 Then
      Rollback;
      ������Ϣ_Out := v_Err_Msg;
      Return 0;
    End If;
  
    --��������
    j_Json := PLJson(v_Respdata);
    --  eScarletBillBatchCode  ���Ӻ�ƱƱ�ݴ���  String  20  ��  
    v_��Ʊ���� := j_Json.Get_String('eScarletBillBatchCode');
    --  eScarletBillNo  ���Ӻ�ƱƱ�ݺ���  String  20  ��  
    v_��Ʊ���� := j_Json.Get_String('eScarletBillNo');
    --  eScarletRandom  ���Ӻ�ƱУ����  String  20  ��  
    v_��ƱУ���� := j_Json.Get_String('eScarletRandom');
    --  createTime  ���Ӻ�Ʊ����ʱ��  String  17  ��  yyyyMMddHHmmssSSS
    v_��Ʊ����ʱ�� := j_Json.Get_String('createTime');
    --  billQRCode  ����Ʊ�ݶ�ά��ͼƬ����  String  ����    ��ֵ��Base64���룬����ʱ��ҪBase64����
    c_��Ʊ��ά�� := j_Json.Get_String('billQRCode');
    --  pictureUrl  ����Ʊ��H5ҳ��URL  String  ����    
    v_��Ʊurl := j_Json.Get_String('pictureUrl');
    --  pictureNetUrl  ����Ʊ������H5ҳ��URL��ַ  String  ����    ��������
    v_��Ʊ����url := j_Json.Get_String('pictureNetUrl');
    --���µ���Ʊ����Ϣ
    Update ����Ʊ��ʹ�ü�¼
    Set ���� = v_��Ʊ����, ���� = v_��Ʊ����, ������ = v_��ƱУ����, ����ʱ�� = v_��Ʊ����ʱ��, Url���� = v_��Ʊurl, Url���� = v_��Ʊ����url
    Where ID = n_����id;
  
    --�����ά��
    Insert Into ����Ʊ�ݶ�ά�� (ʹ�ü�¼id, ��ά��) Values (n_����id, c_��Ʊ��ά��);
    Commit;
    Return 1;
  Exception
    When Others Then
      ������Ϣ_Out := SQLCode || ':' || SQLErrM;
      Return 0;
  End Einvoice_Cancel;
End b_Einvoice_Request_Test;
/

Create Or Replace Procedure Zl_����Ʊ�ݺ˶Լ�¼_Update
(
  �˶�����_In     ����Ʊ�ݺ˶Լ�¼.�˶�����%Type,
  ҵ������_In     ����Ʊ�ݺ˶Լ�¼.ҵ������%Type,
  ��Ʊ��_In       ����Ʊ�ݺ˶Լ�¼.��Ʊ��%Type,
  His��Ʊ��_In    ����Ʊ�ݺ˶Լ�¼.His��Ʊ��%Type,
  His��Ʊ���_In  ����Ʊ�ݺ˶Լ�¼.His��Ʊ���%Type,
  ƽ̨��Ʊ��_In   ����Ʊ�ݺ˶Լ�¼.ƽ̨��Ʊ��%Type,
  ƽ̨��Ʊ���_In ����Ʊ�ݺ˶Լ�¼.ƽ̨��Ʊ���%Type,
  �˶���_In       ����Ʊ�ݺ˶Լ�¼.�˶���%Type,
  �˶�ʱ��_In     ����Ʊ�ݺ˶Լ�¼.�˶�ʱ��%Type,
  �˶Խ��_In     ����Ʊ�ݺ˶Լ�¼.�˶Խ��%Type,
  �˶�˵��_In     ����Ʊ�ݺ˶Լ�¼.�˶�˵��%Type
) As
  --���ܣ�����/�������Ʊ�ݺ˶Լ�¼
  --��Σ�
  --  �˶�����_In 1-�˶Կ�Ʊ����Ʊ��2-���˶���Ʊ
  --  �˶Խ��_In 1-�˶Գɹ���0-�˶�ʧ��
Begin
  Update ����Ʊ�ݺ˶Լ�¼
  Set His��Ʊ�� = His��Ʊ��_In, His��Ʊ��� = His��Ʊ���_In, ƽ̨��Ʊ�� = ƽ̨��Ʊ��_In, ƽ̨��Ʊ��� = ƽ̨��Ʊ���_In, �˶�ʱ�� = �˶�ʱ��_In, �˶Խ�� = �˶Խ��_In,
      �˶�˵�� = �˶�˵��_In
  Where ҵ������ = ҵ������_In And �˶��� = �˶���_In And �˶����� = �˶�����_In;

  If Sql%RowCount = 0 Then
    Insert Into ����Ʊ�ݺ˶Լ�¼
      (ҵ������, ��Ʊ��, His��Ʊ��, His��Ʊ���, ƽ̨��Ʊ��, ƽ̨��Ʊ���, �˶�����, �˶���, �˶�ʱ��, �˶Խ��, �˶�˵��)
    Values
      (ҵ������_In, ��Ʊ��_In, His��Ʊ��_In, His��Ʊ���_In, ƽ̨��Ʊ��_In, ƽ̨��Ʊ���_In, �˶�����_In, �˶���_In, �˶�ʱ��_In, �˶Խ��_In, �˶�˵��_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ�ݺ˶Լ�¼_Update;
/
Create Or Replace Procedure Zl_����Ʊ��������¼_Update
(
  ҵ������_In     ����Ʊ��������¼.ҵ������%Type,
  ����Ʊ��id_In   ����Ʊ��������¼.����Ʊ��id%Type,
  ҵ����ˮ��_In   ����Ʊ��������¼.ҵ����ˮ��%Type,
  His��Ʊ��_In    ����Ʊ��������¼.His��Ʊ��%Type,
  His��Ʊ���_In  ����Ʊ��������¼.His��Ʊ���%Type,
  HisƱ��״̬_In  ����Ʊ��������¼.HisƱ��״̬%Type,
  ƽ̨��Ʊ��_In   ����Ʊ��������¼.ƽ̨��Ʊ��%Type,
  ƽ̨��Ʊ���_In ����Ʊ��������¼.ƽ̨��Ʊ���%Type,
  ƽ̨Ʊ��״̬_In ����Ʊ��������¼.ƽ̨Ʊ��״̬%Type,
  ������ʽ_In     ����Ʊ��������¼.������ʽ%Type,
  ������_In       ����Ʊ��������¼.������%Type,
  ����ʱ��_In     ����Ʊ��������¼.����ʱ��%Type,
  �������_In     ����Ʊ��������¼.�������%Type,
  ����˵��_In     ����Ʊ��������¼.����˵��%Type
) As
  --���ܣ�����/�������Ʊ��������¼
  --��Σ�
  --  HISƱ��״̬_In\ƽ̨Ʊ��״̬_In 1-������2-��죬3-����
  --  ������ʽ_In 1-����HIS���ݣ�2-����ƽ̨���ݣ�3-����HIS��ƽ̨�����ؿ�Ʊ�ݣ�4-�����������
  --  �������_In 1-�����ɹ���0-����ʧ��
Begin
  If ����Ʊ��id_In Is Null Then
    Update ����Ʊ��������¼
    Set ������ʽ = ������ʽ_In, ������ = ������_In, ����ʱ�� = ����ʱ��_In, ������� = �������_In, ����˵�� = ����˵��_In
    Where ҵ������_In = ҵ������_In And ����Ʊ��id Is Null And ҵ����ˮ�� = ҵ����ˮ��_In;
  Elsif ҵ����ˮ��_In Is Null Then
    Update ����Ʊ��������¼
    Set ������ʽ = ������ʽ_In, ������ = ������_In, ����ʱ�� = ����ʱ��_In, ������� = �������_In, ����˵�� = ����˵��_In
    Where ҵ������_In = ҵ������_In And ����Ʊ��id = ����Ʊ��id_In And ҵ����ˮ�� Is Null;
  Else
    Update ����Ʊ��������¼
    Set ������ʽ = ������ʽ_In, ������ = ������_In, ����ʱ�� = ����ʱ��_In, ������� = �������_In, ����˵�� = ����˵��_In
    Where ҵ������_In = ҵ������_In And ����Ʊ��id = ����Ʊ��id_In And ҵ����ˮ�� = ҵ����ˮ��_In;
  End If;

  If Sql%RowCount = 0 Then
    Insert Into ����Ʊ��������¼
      (ҵ������, ����Ʊ��id, ҵ����ˮ��, His��Ʊ��, His��Ʊ���, HisƱ��״̬, ƽ̨��Ʊ��, ƽ̨��Ʊ���, ƽ̨Ʊ��״̬, ������ʽ, ������, ����ʱ��, �������, ����˵��)
    Values
      (ҵ������_In, ����Ʊ��id_In, ҵ����ˮ��_In, His��Ʊ��_In, His��Ʊ���_In, HisƱ��״̬_In, ƽ̨��Ʊ��_In, ƽ̨��Ʊ���_In, ƽ̨Ʊ��״̬_In, ������ʽ_In, ������_In,
       ����ʱ��_In, �������_In, ����˵��_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ʊ��������¼_Update;
/