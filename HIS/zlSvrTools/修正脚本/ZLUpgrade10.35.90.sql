----10.35.80---��10.35.90
--117919:����,2018-1-2,��Long Raw����long����ת��ΪBLOB����CLOB����
Alter Table zltools.zlRPTGraphs Modify ͼƬ Blob;
Alter Table zltools.zlXlsVerify Modify �����Ϣ Clob;

Declare  
  Cursor c_ʧЧ���� Is
    Select 'alter index ' || Index_Name || ' rebuild' As ����
    From user_Indexes
    Where Table_Name In ('ZLRPTGRAPHS', 'ZLXLSVERIFY') And Status = 'UNUSABLE';
Begin
For r_ʧЧ���� In c_ʧЧ���� Loop
    Execute Immediate r_ʧЧ����.����;
  End Loop;
End;
/

--119766:��˶,2017-12-06,�����嵥����
alter table zltools.zlkillprocess add  �Ƿ�̶� number(1);

Alter Table zltools.ZLKillProcess Add Constraint ZLKillProcess_PK Primary Key(���) Using Index;

Insert Into zltools.ZLKILLPROCESS(���,����,����,�Ƿ�̶�,����) 
Select 1   ,'7Z.EXE'                  ,0  ,1  ,'7-ZIPѹ������' From Dual Union All
Select 2   ,'WINCMP3.EXE'             ,0  ,1  ,'�ļ��ȶԹ��ߣ������û��Զ�������ռ���' From Dual Union All
Select 3   ,'ZL9LABPRINTSVR.EXE'      ,0  ,1  ,'�°�LIS��ӡ������Ҫ����������ӡ���档' From Dual Union All
Select 4   ,'ZL9LABRECEIV.EXE'        ,0  ,1  ,'�°�LISͨѶ���򲿼�����Ҫ�����������ӿ�֮�����ݽ�����' From Dual Union All
Select 5   ,'ZL9LABTCPSVR.EXE'        ,0  ,1  ,'�°�LIS������Ϣת������������ʵ���Һ�ͨѶ��������Ϣת����' From Dual Union All
Select 6   ,'ZL9LISCOMM.EXE'          ,0  ,1  ,'�ϰ����ͨѶ���򣬴��������ش����ݣ��ӹ��ɼ���ϵͳ�ܹ���ʶ�����ݸ�ʽ��' From Dual Union All
Select 7   ,'ZL9PACSCAPTURE.EXE'      ,0  ,1  ,'��Ӱ��ɼ�ϵͳ����Ƶ�ɼ���ʽ�Ż�����,������������ActivexExe��Ŀ' From Dual Union All
Select 8   ,'ZL9WIZARDMAIN.EXE'       ,0  ,1  ,'��������ϵͳ��̨��������������ϵͳ�����к�̨���ã�������Դ���á���̬ҳ����ơ���̬ҳ������ȡ�' From Dual Union All
Select 9   ,'ZL9XLS.EXE'              ,0  ,1  ,'Excel�����ߡ�' From Dual Union All
Select 10  ,'ZLACTMAIN.EXE'           ,0  ,1  ,'BH�ں��е����⵼��̨��BH���ø���ģ���ͨ���ó�����е�����' From Dual Union All
Select 11  ,'ZLBAEXPORT.EXE'          ,0  ,1  ,'��ɲ�������dbf�ļ������ɺ��ϴ���FTP' From Dual Union All
Select 12  ,'ZLCDOPEN.EXE'            ,0  ,1  ,'��PACS��¼�������ϵļ��ͼ�񣬸������ߣ���PACS��¼�������ϵļ��ͼ��' From Dual Union All
Select 13  ,'ZLCISAUDITPRINT.EXE'     ,0  ,1  ,'���ڵ��Ӳ��������,�ļ�-�����PDF����������PDF�������ϵͳGDI����������ϵͳ������' From Dual Union All
Select 14  ,'ZLDBATOOLS.EXE'          ,0  ,1  ,'DBA�����ߵ���ִ���ļ���' From Dual Union All
Select 15  ,'ZLDRUGMACHINEMANAGE.EXE' ,0  ,1  ,'ҩ���Զ���ҩϵͳ�ӿ����ú�ҵ�����' From Dual Union All
Select 16  ,'ZLEXINSTALL.EXE'         ,0  ,1  ,'���������װ�����ֽ�֧�ֶ�OO4O�����װ��' From Dual Union All
Select 17  ,'ZLGETIMAGE.EXE'          ,0  ,1  ,'�ṩӰ����ͼ������֧�֣���̨����Ӱ����ͼ��' From Dual Union All
Select 18  ,'ZLGETIMAGEEX.EXE'        ,0  ,1  ,'zlGetImageEx��ʹ��ActiveExe�ķ�ʽʵ�ֺ�̨���̼��ؼ��ϴ�ͼ����һ��ActiveExe����' From Dual Union All
Select 19  ,'ZLHEALTHSERVICE.EXE'     ,0  ,1  ,'ʵ�ֽ���������ĺ�̨�������е���������' From Dual Union All
Select 20  ,'ZLHIS+.EXE'              ,0  ,1  ,'ZLHIS+�������򣬵�¼�ó�����ܽ��뵼��̨������ҵ�������' From Dual Union All
Select 21  ,'ZLHISCRUST.EXE'          ,0  ,1  ,'�ͻ����Զ��������ߣ�ͨ���ù��߶Ը����ͻ��˽����ļ�������' From Dual Union All
Select 22  ,'ZLHQMSDCOLLECT.EXE'      ,0  ,1  ,'���ҽԺ����������ݵĲɼ��ϱ�����' From Dual Union All
Select 23  ,'ZLLISMESSAGE.EXE'        ,0  ,1  ,'LIS��Ϣ�������ڴ���Ļ����ʾĳЩ������ڲ��������' From Dual Union All
Select 24  ,'ZLLISRECEIVESEND.EXE'    ,0  ,1  ,'��������:��Ҫ���������ֱ��ͨѶ����¼�����ش��ļ��������������ı�ΪLIS��ʶ�ļ�������' From Dual Union All
Select 31  ,'ZLNEWQUERY.EXE'          ,0  ,1  ,'�ϰ�����ϵͳ,�����Һš�Lis��ӡ�����ò�ѯ��' From Dual Union All
Select 32  ,'ZLORCLCONFIG.EXE'        ,0  ,1  ,'���ڿ�������ORACLE�����ļ��Ĺ��ߡ�' From Dual Union All
Select 33  ,'ZLPACSBROWSERSTATION.EXE',0  ,1  ,'������Ƭվ��' From Dual Union All
Select 34  ,'ZLPACSFTPTOOLS.EXE'      ,0  ,1  ,'��FTP���в��ԣ��Ų�FTP��ز�������' From Dual Union All
Select 35  ,'ZLPACSSERVERCENTER.EXE'  ,1  ,1  ,'PACS��������' From Dual Union All
Select 36  ,'ZLPACSSRV.EXE'           ,0  ,1  ,'����Dicom�豸���͵ļ��ͼ��PACS���ط��񣬼���Ӱ��DICOM�豸����' From Dual Union All
Select 37  ,'ZLPEISAUTOANALYSE.EXE'   ,0  ,1  ,'����Զ���������ʵ�ַǱ�׼���������ݽӿڡ�' From Dual Union All
Select 38  ,'ZLQUEUESHOW.EXE'         ,0  ,1  ,'�°��Ŷ���ʾ��pacs�Ŷ������ʾ��' From Dual Union All
Select 39  ,'ZLRISDUMPTOOL.EXE'       ,0  ,1  ,'�������ݣ��û���������Ŀ�������ֵ�ȳ�ʼ������ʼ��ris�ӿ����ݡ�' From Dual Union All
Select 40  ,'ZLRPTSQLADJUST.EXE'      ,0  ,1  ,'10.26���˷��ñ������׹��ߡ����д���ֺ���漰���˷��ü�¼�ı���ĵ�����' From Dual Union All
Select 41  ,'ZLRUNAS.EXE'             ,0  ,1  ,'���ļ����Զ�����zlhisCrust.exe��ʹ�á���Ҫ����,��USERȨ���¿���ʹ�ù���ԱȨ�������е�¼ִ�й������' From Dual Union All
Select 42  ,'ZLSCREENKEYBOARD.EXE'    ,0  ,1  ,'��Ļ����С����������ҽ������վ���õ���ǿ���������ҽ���´' From Dual Union All
Select 43  ,'ZLSOFTSHOWARCHIVE.EXE'   ,0  ,1  ,'��ʾ��������,ҽ����pacs��ʷ����ȣ�ris�е��ò鿴�������ݡ�' From Dual Union All
Select 44  ,'ZLSOFTSHOWHISFORMS.EXE'  ,0  ,1  ,'��ʾ��������,ҽ����pacs��ʷ����ȣ�ris�е��ò鿴�������ݡ�' From Dual Union All
Select 45  ,'ZLSQLTRACE.EXE'          ,0  ,1  ,'zlSQL���ٹ��ߡ�' From Dual Union All
Select 46  ,'ZLSVRNOTICE.EXE'         ,0  ,1  ,'�Զ����ѷ��񣬽�����Ϣ���ѵ���ʾ���Ķ���' From Dual Union All
Select 47  ,'ZLSVRSTUDIO.EXE'         ,0  ,1  ,'ZLHISϵͳ�ĺ�̨�����ߡ��ṩ��ϵͳ����������װ����Ȩ�Լ�������ʵ�ù��ܣ����Է���Ľ��к�̨����' From Dual Union All
Select 48  ,'ZLUPGRADEREADER.EXE'     ,0  ,1  ,'����˵���Ķ����������ش��ܵĺ˶��Լ���ѵ���˵Ĵ���' From Dual Union All
Select 49  ,'ZLWIZARDSTART.EXE'       ,0  ,1  ,'����ϵͳǰ̨��ѯ����������������ϵͳǰ̨���ܡ�' From Dual;

--117980:����,2017-12-06,��Ҫ�����䶯��־�������
Insert Into Zltools.Zlsvrtools (���, �ϼ�, ����, ���, ˵��, ����) Values ('0314', '03', '������־����', 'T', Null, 17);
Alter Table Zltools.zlauditlog add ��־ID Number(18);

--117980:����,2017-12-06,��Ҫ�����䶯��־�������
Create Table Zltools.ZlauditlogConfig(
       ID Number(18),
       ϵͳ Number(5),
       ģ�� Varchar2(18),
       ���� Varchar2(50),
       ˵�� Varchar2(250),
       �Ƿ����� Number(1),
       �Ƿ������ Number(1)
);
Alter Table Zltools.zlauditlogconfig Add Constraint ZlauditlogConfig_PK Primary Key(ID) Using Index;
Alter Table Zltools.ZlauditlogConfig Add Constraint ZlauditlogConfig_UQ_ϵͳ Unique(ϵͳ, ģ��,����) Using Index;
Alter Table Zltools.ZlauditlogConfig Add Constraint ZlauditlogConfig_FK_ϵͳ Foreign Key(ϵͳ) References Zlsystems(���) On Delete Cascade;
Insert Into Zltools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLAUDITLOGCONFIG','ZLTOOLSTBS','A2');
Insert Into Zltools.Zloptions(������, ������, ����ֵ, ȱʡֵ, ����˵��) Values(25, '������־�����������', '365', '365', '������־����ܱ��������������ʱϵͳ�����Զ�ɾ�������ٱ���90�죬����Ϊ0ʱ��ʾ���ñ���');
Create Sequence Zltools.ZlauditlogConfig_ID start with 1;
CREATE INDEX Zltools.zlClients_IX_IP ON zlClients(IP);

--119259:����,2018-1-2,����������ģ�������Ҫ�����䶯��־
Insert Into Zltools.Zlauditlogconfig(ID, ģ��, ����, ˵��, �Ƿ�����, �Ƿ������)
Select 1, '0201', '��ж','��ж��ʷ���ݿռ䣬��ѡ������ݿռ��ǵ�ǰ����ʹ�õģ����ܲ�ж',1,1 From Dual Union All
Select 2, '0201', '�л�','�л���ǰ��ʷ���ݿռ䣬�������´���H����ͼָ���л������ʷ���ݿռ�',1,0 From Dual Union All
Select 3, '0201', '�ϲ�','������������ʷ���ݿռ�ϲ�Ϊһ�������ݻ��Զ��ϲ��������С�Ŀռ���',1,0 From Dual Union All
Select 4, '0202', 'ִ��','�������ݡ�Լ����������Ȩ�޵����������ļ�',1,0 From Dual Union All
Select 5, '0203', 'ִ��','�������ݡ�Լ����������Ȩ���ɱ����ļ��������ݿ�',1,0 From Dual Union All
Select 6, '0206', 'ִ��','��ָ�����е�����ȫ�����',1,1 From Dual Union All
Select 7, '0207', '����','���һ���������ӣ����ڿͻ��������������ݿ�����ѯ����',1,0 From Dual Union All
Select 8, '0207', '�޸�','�޸�һ���������ӣ����ܵ���ԭ��ʹ�ø����ӵĹ��ܲ�ѯ��������',1,0 From Dual Union All
Select 9, '0207', 'ɾ��','ɾ��һ���������ӣ�ɾ���󣬿ͻ��˽�����ʹ�ø��������������Ӧ���ݿ�',1,1 From Dual Union All
Select 10, '0302', '�����Ự','ǿ�жϿ�һ���û������ݿ�֮������ӣ����ܵ��¸��û�δ��������ݶ�ʧ',1,0 From Dual Union All
Select 11, '0303', '����','���һ���Զ���ҵ��Ŀ�����ڶ�ʱ���������ָ������',1,0 From Dual Union All
Select 12, '0303', 'ɾ��','ɾ��һ���Զ���ҵ��Ŀ�������Զ���ҵ�����ö�ʱִ�У�ɾ���󽫻ᵼ�¸ö�ʱ�����޷�ִ��',1,1 From Dual Union All
Select 13, '0303', '��������','�޸�һ���Զ���ҵ��Ŀ���������ƣ��������ݣ�ѭ��ʱ�估ִ��ʱ���',1,0 From Dual Union All
Select 14, '0304', 'ɾ��','ɾ��ָ�������»�ȫ����������־',1,1 From Dual Union All
Select 15, '0305', 'ɾ��','ɾ��ָ�������»�ȫ���Ĵ�����־',1,1 From Dual Union All
Select 16, '0306', '����','����һЩ��Ҫϵͳ�������޸�',1,0 From Dual Union All
Select 17, '0307', '�ļ�����������-����','���һ���ļ������������ڱ���ͻ�������ʱ��Ҫ�������ļ�',1,0 From Dual Union All
Select 18, '0307', '�ļ�����������-�޸�','�޸�һ���ļ��������Ļ�����Ϣ�Լ���ͣ�÷�����',1,0 From Dual Union All
Select 19, '0307', '�ļ�����������-ɾ��','ɾ��һ���ļ������������÷�����Ϊȱʡ������������ɾ��',1,1 From Dual Union All
Select 20, '0307', '�ļ���������-����','���һ�����������������ݿ��У����ڿͻ�������ʱͬ����һЩ����������һ������',1,0 From Dual Union All
Select 21, '0307', '�ļ���������-�޸�','�޸�һ����������������Ϣ����Ҫ������������ϵͳ�Լ����Ƿ���Ҫע��',1,0 From Dual Union All
Select 22, '0307', '�ļ���������-ɾ��','�����ݿ���ɾ��һ���������������䲢����Ӱ�������������е��ļ�',1,1 From Dual Union All
Select 23, '0307', '�ļ���������-����','�����ݿ�������һ���������������䲢����Ӱ�������������е��ļ�',1,0 From Dual Union All
Select 24, '0307', '�ϴ��µ��ļ�','���Ѿ��Ǽǵı����еĵ��Ƿ�������û�е��ļ��ϴ���������',1,0 From Dual Union All
Select 25, '0307', '�ϴ������ļ�','���Ѿ��Ǽǵı������е��ļ����ϴ���������',1,0 From Dual Union All
Select 26, '0307', '����/ȡ������','Ϊĳ���ͻ���ִ��������ȡ����������',1,0 From Dual Union All
Select 27, '0307', 'Ԥ����/ȡ��Ԥ����','Ϊĳ���ͻ���ִ��Ԥ������ȡ��Ԥ��������',1,0 From Dual Union All
Select 28, '0307', 'ȫ������/ȡ��ȫ������','�����пͻ���ִ��������ȡ����������',1,0 From Dual Union All
Select 29, '0307', 'ȫ��Ԥ����/ȡ��ȫ��Ԥ����','�����пͻ��˽���Ԥ������ȡ��Ԥ��������',1,0 From Dual Union All
Select 30, '0308', '�޸�','�޸�һ���ͻ��˵ĸ������',1,1 From Dual Union All
Select 31, '0308', 'ɾ��','ɾ��һ��ָ���Ŀͻ��ˣ�ɾ������ڸÿͻ��˵�һ�����ö��������',1,1 From Dual Union All
Select 32, '0308', '����/����','���û�����һ���ͻ��ˣ����ú�ÿͻ��˽����ܵ�¼����Ʒ',1,0 From Dual Union All
Select 33, '0308', 'ȫ������/ȫ������','���û�����ȫ���ͻ��ˣ����ú����пͻ��˶������ܵ�¼����Ʒ',1,0 From Dual Union All
Select 34, '0308', '����3����δ��¼�ͻ���','������������δ��¼���Ŀͻ���',1,1 From Dual Union All
Select 35, '0312', '������Ŀ','���һ��ҽԺ������Ϣ���ƣ���Ҫ����ҽԺ��Ϣ��Ŀ�Ķ���',1,0 From Dual Union All
Select 36, '0312', '������Ŀ','�޸�һ��ҽԺ������Ϣ���ƣ���Ҫ����ҽԺ��Ϣ��Ŀ�ĵ���',1,1 From Dual Union All
Select 37, '0312', 'ɾ����Ŀ','ɾ��һ��ҽԺ������Ϣ���ƣ�ɾ������ܵ���ҽԺ������Ϣ��ȱʧ',1,1 From Dual Union All
Select 38, '0314', '��־����','����ָ�����������������Ҫ�����䶯��־',1,1 From Dual Union All
Select 39, '0401', '���ӽ�ɫ','����һ���հ�Ȩ�޵Ľ�ɫ',1,0 From Dual Union All
Select 40, '0401', '��ɫ��Ȩ','��һ����ɫ������Ȩ��ʹ����һЩָ����Ȩ��',1,0 From Dual Union All
Select 41, '0401', 'ɾ����ɫ','�Խ�ɫ����ɾ���������ý�ɫ�Լ��ý�ɫӵ�е�Ȩ�޶�����ɾ��',1,1 From Dual Union All
Select 42, '0401', '���ƽ�ɫ','����һ����ɫ���Ƴ���һ����ɫ�����Ƴ��Ľ�ɫ��ӵ�к�ԭ��ɫһ����Ȩ����Ϣ',1,0 From Dual Union All
Select 43, '0401', '�������н�ɫ','�������Ʒ�б�������н�ɫ�������û������ݿ���ʵ��ӵ�еĽ�ɫ���²������н�ɫ����',1,0 From Dual Union All
Select 44, '0401', '�ָ����н�ɫ��Ȩ��','����Ӧ��ϵͳ�б�������н�ɫ�������ݿ��м�鲢���䴴����ɫ������Ӧ��ϵͳ�Ĺ�����������Ȩ�ޣ��Լ�����������ݿ����ķ���Ȩ��',1,0 From Dual Union All
Select 45, '0401', '�޸�ģ���ʹ��Ȩ��','�޸�Ӧ��ϵͳ��ĳЩģ�����ɾ�ĵ�ʹ��Ȩ��',1,0 From Dual Union All
Select 46, '0401', '�޸Ľ�ɫ����Ȩ�û�','��ѡ�н�ɫ��������ĳЩ�û�',1,0 From Dual Union All
Select 47, '0402', '���������û�','������ѡ���ŵ���Ա���������ϻ��û����û�Ĭ�ϲ������κν�ɫȨ��',1,0 From Dual Union All
Select 48, '0402', '�����û�','����һ���û���Ϊ�����Ա��Ϣ���������ɫȨ��',1,0 From Dual Union All
Select 49, '0402', '�޸��û�','��һ���û��󶨵���Ա�Լ�Ȩ����Ϣ�����޸�',1,0 From Dual Union All
Select 50, '0402', 'ɾ���û�','ɾ��һ���û���ͬʱ���ͷ�����û��󶨵���Ա��Ȩ��',1,1 From Dual Union All
Select 51, '0402', '��ͣ�û�','���û�ͣ��һ���û���ͣ�ú󣬸��û�����ʱ����ʹ��',1,0 From Dual Union All
Select 52, '0402', '�޸�����','�޸�һ���û��ĵ�½����',1,0 From Dual Union All
Select 53, '0402', '�����ϻ���Ա�ָ��û�','�����ϻ��û��������ݻָ����ָ�����֮�󴴽���ǰ���û�����Ȩ�������û�����Ϊ��ʼ����',1,0 From Dual Union All
Select 54, '0402', '�ָ������û���ɫ','�����û���Ӧ��ϵͳ�еļ�¼��ɫ���½��н�ɫ��Ȩ���ָ���ɫ���û�֮���ؽ��û��Ľ�ɫ',1,0 From Dual Union All
Select 55, '0402', '���������û���ɫ','�������Ʒ�б���������û��Ľ�ɫ�������û������ݿ���ʵ��ӵ�еĽ�ɫ���²�������Ʒ�������û��Ľ�ɫ����',1,0 From Dual Union All
Select 56, '0403', 'ɾ��','ɾ��һ���˵��飬������ʹ�øò˵���ĵ���̨����ʹ��ȱʡ�˵���',1,1 From Dual Union All
Select 57, '0504', '����','���һ��������Ϣ����Ҫ������ָ��ʱ����һЩ�û�����ĳ����Ϣ',1,0 From Dual Union All
Select 58, '0504', '�޸�','�޸�һ��������Ϣ����Ҫ���޸��������ݡ����������Լ����ѷ�ʽ������',1,0 From Dual Union All
Select 59, '0504', 'ɾ��','ɾ��һ��������Ϣ��ɾ�����Ӧ�û������ղ���������',1,1 From Dual;
Select Zltools.ZlauditlogConfig_ID.Nextval From Dual Connect By Rownum <= (Select Nvl(Max(ID), 0) From Zltools.ZlauditlogConfig);
--117980:����,2017-12-26,����zlauditlog���е���ʷ����
Declare
  Cursor c_Auditlog Is
    Select ����ģ����, ��������, Rowid From Zltools.Zlauditlog; --�����α�
Begin
  For c_Item In c_Auditlog Loop
    If To_Number(c_Item.����ģ����) = 401 Then
      Case
        When Instr(c_Item.��������, '������ɫ') > 0 Then
          Update Zltools.Zlauditlog Set ��־ID = 1 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.��������, '�޸Ľ�ɫ') > 0 And Instr(c_Item.��������, '��Ȩ��') > 0 Then
          Update Zltools.Zlauditlog Set ��־ID = 2 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.�������� = '�����н�ɫ������Ȩ' Then
          Update Zltools.Zlauditlog Set ��־ID = 6 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.��������, 'ɾ����ɫ') > 0 Then
          Update Zltools.Zlauditlog Set ��־ID = 3 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.�������� = 'ִ�в������޸Ľ�ɫ����Ȩ�û�' Then
          Update Zltools.Zlauditlog Set ��־ID = 8 Where Zlauditlog.Rowid = c_Item.Rowid;
        Else
          Update Zltools.Zlauditlog Set ��־ID = 1 Where Zlauditlog.Rowid = c_Item.Rowid;
      End Case;
    Else
      Case 
        When Instr(c_Item.��������, '�����û�') > 0 Then
          Update Zltools.Zlauditlog Set ��־ID = 10 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.��������, '�޸��û�') > 0 Then
          Update Zltools.Zlauditlog Set ��־ID = 11 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.��������, 'ɾ���û�') > 0 Then
          Update Zltools.Zlauditlog Set ��־ID = 12 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.�������� = 'ִ�в��������������û�' Or c_Item.�������� = 'ִ�в��������������û�' Then
          Update Zltools.Zlauditlog Set ��־ID = 9 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.�������� = 'ִ�в����������ϻ���Ա�ָ��û�' Or c_Item.�������� = 'ִ�в������ָ������ϻ���Ա' Then
          Update Zltools.Zlauditlog Set ��־ID = 15 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.�������� = 'ִ�в������ָ������û���ɫ' Then
          Update Zltools.Zlauditlog Set ��־ID = 16 Where Zlauditlog.Rowid = c_Item.Rowid;
        When c_Item.�������� = 'ִ�в��������������û���ɫ' Or c_Item.�������� = 'ִ�в�������¼�����û���ɫ' Then
          Update Zltools.Zlauditlog Set ��־ID = 17 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.��������, '�����û�') > 0 Or Instr(c_Item.��������, '�����û�') > 0 Then
          Update Zltools.Zlauditlog Set ��־ID = 13 Where Zlauditlog.Rowid = c_Item.Rowid;
        When Instr(c_Item.��������, '�޸��û�') > 0 And Instr(c_Item.��������, '������') > 0 Then
          Update Zltools.Zlauditlog Set ��־ID = 14 Where Zlauditlog.Rowid = c_Item.Rowid;
        Else
          Update Zltools.Zlauditlog Set ��־ID = 10 Where Zlauditlog.Rowid = c_Item.Rowid;
      End Case;
    End If;
  End Loop;
End;
/

--117980:����,2017-12-06,��Ҫ�����䶯��־�������������־����
Create Or Replace Procedure Zltools.Zl_Zlauditlog_Delete(��־��������_In In Number) Is
Begin
  Delete Zlauditlog Where ����ʱ�� < Sysdate - ��־��������_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlauditlog_Delete;
/

--117980:����,2017-12-21,�Զ�������Ҫ�����䶯��־
Create Or Replace Procedure Zltools.Zl_Autologprocess As
  --���ܣ� 
  --   �Զ����������־�ʹ�����־����Ҫ�����䶯��־������� 
  v_Limit Number;
Begin
  --ɾ�������������־ 
  Select Nvl(Max(To_Number(����ֵ)), 0) Into v_Limit From zlOptions Where ������ = 2;
  Delete From zlDiaryLog Where ����ʱ�� < Sysdate - v_Limit;

  --ɾ������Ĵ�����־ 

  Select Nvl(Max(To_Number(����ֵ)), 0) Into v_Limit From zlOptions Where ������ = 4;
  Delete From zlErrorLog Where ʱ�� < Sysdate - v_Limit;

  --ɾ���������Ҫ�����䶯��־
  Select Nvl(����ֵ, ȱʡֵ) Into v_Limit From zlOptions Where ������ = 25;
  If v_Limit <> 0 Then
    Delete Zlauditlog Where ����ʱ��  < Sysdate - v_Limit;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Autologprocess;
/

--117980:����,2017-12-06,��Ҫ�����䶯��־��������޸�ģ���¼��־����ͣ״̬
CREATE OR REPLACE Procedure zltools.Zl_Zlauditlogconfig_Update
(
  Id_In         In Zlauditlogconfig.Id%Type,
  �Ƿ�����_In   In Zlauditlogconfig.�Ƿ�����%Type,
  �Ƿ������_In In Zlauditlogconfig.�Ƿ������%Type := Null
) Is
Begin
  If �Ƿ������_In Is Null Then
    Update Zlauditlogconfig Set �Ƿ����� = �Ƿ�����_In Where Id = Id_In;
  Else
    Update Zlauditlogconfig Set �Ƿ������ = �Ƿ������_In Where Id = Id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlauditlogconfig_Update;
/

--117980:����,2017-12-06,��Ҫ�����䶯��־������������־
CREATE OR REPLACE Procedure zltools.Zl_Zlauditlog_Insert
(
  �û���_In   Zlauditlog.�û���%Type,
  ����վ_In   Zlauditlog.����վ%Type,
  ��������_In Zlauditlog.��������%Type, --1-������2-�޸ģ�3-ɾ��
  ϵͳ_In     Zlauditlogconfig.ϵͳ%Type,
  ģ��_In     Zlauditlogconfig.ģ��%Type,
  ����_In     Zlauditlogconfig.����%Type,
  ��������_In Zlauditlog.��������%Type,
  ����˵��_In Zlauditlog.����˵��%Type --������¼�����ṩ������Ա����ı�ע��Ϣ
) Is
  n_�Ƿ����� Zlauditlogconfig.�Ƿ�����%Type;
  n_��־Id Zlauditlogconfig.Id%Type;
Begin
  --����ϵͳ��ţ�ģ���ź͹������Ʋ��ҳ���ǰ�����Ƿ����˼�¼��Ҫ�����䶯��־
  If ϵͳ_In Is Null Then
    Select Max(�Ƿ�����), Max(Id)
    Into n_�Ƿ�����, n_��־Id
    From Zlauditlogconfig
    Where ϵͳ Is Null And ģ�� = ģ��_In And ���� = ����_In;
  Else
    Select Max(�Ƿ�����), Max(Id)
    Into n_�Ƿ�����, n_��־Id
    From Zlauditlogconfig
    Where ϵͳ = ϵͳ_In And ģ�� = ģ��_In And ���� = ����_In;
  End If;
  If n_�Ƿ����� = 1 Then
    Insert Into Zlauditlog
      (�û���, ����վ, ����ʱ��, ��������, ��־Id, ��������, ����˵��)
    Values
      (�û���_In, ����վ_In, Sysdate, ��������_In, n_��־Id, ��������_In, ����˵��_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlauditlog_Insert;
/

--118170:����,2017-12-14,ɱ���̹������
Create Or Replace Procedure Zltools.Zl_Zlkillprocess_Edit
(
  ����_In In Number, --1:����;2:�޸�;3:ɾ��
  ���_In In Zlkillprocess.���%Type,
  ����_In In Zlkillprocess.����%Type := Null,
  ����_In In Zlkillprocess.����%Type := Null,
  ����_In In Zlkillprocess.����%Type := Null
) As
  n_��� Zlkillprocess.���%Type;
Begin
  If ����_In = 1 Then
    --��ȡ������
    Select Max(���) + 1 Into n_��� From Zlkillprocess;
    --��������
    Insert Into Zlkillprocess (���, ����, ����, ����, �Ƿ�̶�) Values (n_���, ����_In, ����_In, ����_In, 0);
  Elsif ����_In = 2 Then
    Update Zlkillprocess Set ���� = ����_In, ���� = ����_In, ���� = ����_In Where ��� = ���_In;
  Else
    Delete Zlkillprocess Where ���� = ����_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlkillprocess_Edit;
/

--117980:����,2017-12-26,�ϳ�����ģ�����ֶβ�������������Լ��
Alter Table Zltools.zlauditlog Rename Column ����ģ���� To ����ģ����_bak;
Alter Table Zltools.zlAuditLog Drop Constraint zlAuditLog_PK Cascade Drop Index;
Alter Table Zltools.zlAuditLog Add Constraint zlAuditLog_PK Primary Key (����ʱ��,�û���,����վ,��־ID) Using Index;
Alter Table Zltools.zlAuditLog Add Constraint zlAuditLog_FK_��־ID Foreign Key(��־ID) References ZlauditlogConfig(ID) On Delete Cascade;
Create Index Zltools.Zlauditlog_IX_��־ID On Zlauditlog(��־ID);

--116688:����һ,2017-12-25,��¼������ر�ṹ����(�û����ַ������޸�)
Alter Table zltools.zlAppPermission Modify (�û��� varchar2(20));
Alter Table zltools.zlLoginLimit Modify (�û��� varchar2(20));

--117833;����һ,2017-12-25,�����߶�����ƹ���
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0610','06','������ƹ���','O',Null,10);
--111882:������,2017-12-27,��ǩˮƽ��ת��ʾ
Alter Table zlTools.zlRPTItems Add ˮƽ��ת Number(1);


--119139:����һ,2017-12-28,Զ���û������뱣��
Create Or Replace Procedure Zltools.Zl_Zlclients_Set
(
  n_Mode_In       Number,
  n_Rowid_In      Varchar2 := Null,
  v_����վ_In     Zlclients.����վ%Type := Null,
  v_Ip_In         Zlclients.Ip%Type := Null,
  v_Cpu_In        Zlclients.Cpu%Type := Null,
  v_�ڴ�_In       Zlclients.�ڴ�%Type := Null,
  v_Ӳ��_In       Zlclients.Ӳ��%Type := Null,
  v_����ϵͳ_In   Zlclients.����ϵͳ%Type := Null,
  v_����_In       Zlclients.����%Type := Null,
  v_��;_In       Zlclients.��;%Type := Null,
  v_˵��_In       Zlclients.˵��%Type := Null,
  n_����������_In Zlclients.����������%Type := Null,
  n_������־_In   Zlclients.������־%Type := 0,
  n_������_In     Zlclients.������%Type := 0,
  v_վ��_In       Zlclients.վ��%Type := Null,
  n_Apply_In      Number := 0,
  v_Ipbegin_In    Varchar2 := Null,
  v_Ipend_In      Varchar2 := Null,
  n_������ƵԴ    Zlclients.������ƵԴ%Type := Null,
  v_����Ա�û�_In Zlclients.����Ա�û�%Type := Null,
  v_����Ա����_In Zlclients.����Ա����%Type := Null
  --���ܣ������ͻ��˻�վ�� ���߸��¿ͻ�������
  --Ӧ�ã�1�������ߣ��������޸�վ�� ���޸�ʱ��IP��ͻ������ж����������贫��N_Rowid_In��
  --      2��Ӧ��ϵͳ����¼ʱ���ݵ�ǰ��¼�Ŀͻ������ж��Ƿ�
  --                   ����վ����޸�վ�����������ʱN_Rowid_In�贫�룩
  --վ������:0-����վ�㣬1-����վ��
  --N_Apply_In,վ�����Ӧ�÷�Χ��0-��վ�㣬1�������ţ�2������վ�㣬3���̶�IP��
  --V_Ipbegin_In,V_Ipend_In:�ڹ̶�IP��Ӧ��ʱ����,������һ��IP���ϣ���ǰ�沿����ͬ
) Is
  n_Pos         Number(3);
  n_Ipbegin_Num Number;
  n_Ipend_Num   Number;
  n_Ip_Num      Number;
  n_Count       Number;

  v_Err Varchar2(500);
  Err_Custom Exception;

  Function Get_Ipnum(v_Ip_Input Varchar2) Return Number Is
    v_Ip_Num  Varchar2(20);
    n_Pos_Tmp Number;
    v_Ip_Tmp  Varchar2(20);
  Begin
    n_Pos_Tmp := Length(v_Ip_Input);
    n_Pos_Tmp := n_Pos_Tmp - Length(Replace(v_Ip_Input, '.', ''));
    If n_Pos_Tmp <> 3 Then
      Return Null;
    Else
      v_Ip_Tmp := v_Ip_Input;
      Loop
        n_Pos_Tmp := Instr(v_Ip_Tmp, '.');
        Exit When(Nvl(n_Pos_Tmp, 0) = 0);
        --��ÿһ������ת��Ϊ3λ��
        v_Ip_Num := v_Ip_Num || Trim(To_Char(Substr(v_Ip_Tmp, 1, n_Pos_Tmp - 1), '099'));
        v_Ip_Tmp := Substr(v_Ip_Tmp, n_Pos_Tmp + 1);
      End Loop;
      v_Ip_Num := v_Ip_Num || Trim(To_Char(v_Ip_Tmp, '099'));
      n_Ip_Num := To_Number(Trim(v_Ip_Num));
      Return n_Ip_Num;
    End If;
  End;
Begin
  If n_Mode_In = 0 Then
  
    Select Count(1) Into n_Count From zlClients Where ����վ = v_����վ_In;
    If n_Count = 0 Then
      Insert Into zlClients
        (Ip, ����վ, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ����������, ������־, ������, վ��, ������ƵԴ, �����½ʱ��, ����Ա�û�, ����Ա����)
      Values
        (v_Ip_In, v_����վ_In, v_Cpu_In, v_�ڴ�_In, v_Ӳ��_In, v_����ϵͳ_In, v_����_In, v_��;_In, v_˵��_In, n_����������_In, n_������־_In,
         n_������_In, v_վ��_In, n_������ƵԴ, Sysdate, v_����Ա�û�_In, v_����Ա����_In);
    Else
      v_Err := '�Ѿ���������ͬIP��ַ����վ,��������!';
      Raise Err_Custom;
    End If;
  Else
    If n_Rowid_In Is Null Then
      Update zlClients
      Set Cpu = v_Cpu_In, �ڴ� = v_�ڴ�_In, Ӳ�� = v_Ӳ��_In, ����ϵͳ = v_����ϵͳ_In, ���� = v_����_In, ��; = v_��;_In, ˵�� = v_˵��_In,
          ������ = n_������_In, վ�� = v_վ��_In, ������ƵԴ = n_������ƵԴ, ���������� = n_����������_In, ������־ = n_������־_In, �����½ʱ�� = Sysdate,
          ����Ա�û� = Decode(v_����Ա�û�_In, '�տ�', Null, Nvl(v_����Ա�û�_In, ����Ա�û�)),
          ����Ա���� = Decode(v_����Ա����_In, '�տ�', Null, Nvl(v_����Ա����_In, ����Ա����))
      Where ����վ = v_����վ_In And Ip = v_Ip_In;
    Else
      Update zlClients
      Set ����վ = v_����վ_In, Ip = v_Ip_In, Cpu = Decode(v_Cpu_In, Null, Cpu, v_Cpu_In),
          �ڴ� = Decode(v_�ڴ�_In, Null, �ڴ�, v_�ڴ�_In), Ӳ�� = Decode(v_Ӳ��_In, Null, Ӳ��, v_Ӳ��_In),
          ����ϵͳ = Decode(v_����ϵͳ_In, Null, ����ϵͳ, v_����ϵͳ_In), ���� = v_����_In, վ�� = v_վ��_In, ������ƵԴ = n_������ƵԴ, �����½ʱ�� = Sysdate,
          ����Ա�û� = Decode(v_����Ա�û�_In, '�տ�', Null, Nvl(v_����Ա�û�_In, ����Ա�û�)),
          ����Ա���� = Decode(v_����Ա����_In, '�տ�', Null, Nvl(v_����Ա����_In, ����Ա����))
      Where Rowid = n_Rowid_In;
    End If;
  End If;
  --������
  If n_Apply_In = 1 Then
    Update zlClients
    Set ������ = n_������_In, վ�� = v_վ��_In
    Where Nvl(����, 'NONE') = Nvl(v_����_In, 'NONE') And Ip <> v_Ip_In;
  Elsif n_Apply_In = 2 Then
    Update zlClients Set ������ = n_������_In, վ�� = v_վ��_In Where Ip <> v_Ip_In;
  Elsif n_Apply_In = 3 Then
    n_Pos := Length(v_Ipbegin_In);
    n_Pos := n_Pos - Length(Replace(v_Ipbegin_In, '.', ''));
    If n_Pos <> 3 Then
      v_Err := '��ʼIP��ʽ����';
      Raise Err_Custom;
    End If;
    n_Pos := Length(v_Ipend_In);
    n_Pos := n_Pos - Length(Replace(v_Ipend_In, '.', ''));
    If n_Pos <> 3 Then
      v_Err := '����IP��ʽ����';
      Raise Err_Custom;
    End If;
  
    n_Ipbegin_Num := Get_Ipnum(v_Ipbegin_In);
    n_Ipend_Num   := Get_Ipnum(v_Ipend_In);
    For r_Ip In (Select ����վ, Ip From zlClients) Loop
      n_Ip_Num := Get_Ipnum(r_Ip.Ip);
      If n_Ip_Num >= n_Ipbegin_Num And n_Ip_Num <= n_Ipend_Num Then
        Update zlClients Set ������ = n_������_In, վ�� = v_վ��_In Where ����վ = r_Ip.����վ And Ip = r_Ip.Ip;
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Set;
/
--119449:��˶,2018-01-05,����ƽ̨��Ϣ����
Insert Into ZLTOOLS.zlOptions(������,������,����ֵ,ȱʡֵ,����˵��) Values(26, '����ƽ̨��Ϣ��������', '7', '7', '������ƽ̨ʹ�õ�ҵ����Ϣ��������ܱ�����������������Ϊ0ʱ���Զ��������7�����Ϣ����');

--120029:����,2018-01-12,��ʱ����������
Insert Into Zltools.Zlsvrtools (���, �ϼ�, ����, ���, ˵��, ����) Values ('0315', '03', '������ʱ����', 'B', Null, 20);

--120029:����,2018-01-12,��ʱ����������
Create Table Zltools.ZlRunLimit(
       ��� Number(3),
       ���� Varchar2(50),
       �Ƿ����� Number(1),
       ���� Varchar2(250)
);
Alter Table Zltools.ZlRunLimit Add Constraint ZlRunLimit_PK Primary Key(���) Using Index;
Alter Table Zltools.ZlRunLimit Add Constraint ZlRunLimit_UQ_���� Unique(����) Using Index;
Insert Into Zltools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRUNLIMIT','ZLTOOLSTBS','A2');

--120029:����,2018-01-12,��ʱ����������
Create Table Zltools.ZlRunLimitTime(
       ID Number(18),
       ���� Number(3),
       ���� Number(1),
       ��ʼʱ�� date,
       ����ʱ�� date
);
Alter Table Zltools.ZlRunLimitTime Add Constraint ZlRunLimitTime_PK Primary Key(ID) Using Index;
Alter Table Zltools.ZlRunLimitTime Add Constraint ZlRunLimitTime_UQ_����ʱ�� Unique(����,����,��ʼʱ��,����ʱ��) Using Index;
Alter Table Zltools.ZlRunLimitTime Add Constraint ZlRunLimitTime_FK_���� Foreign Key(����) References ZlRunLimit(���) On Delete Cascade;
Create Sequence Zltools.ZlRunLimitTime_ID start with 1;
Insert Into Zltools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRUNLIMITTIME','ZLTOOLSTBS','A2');

--120029:����,2018-01-12,��ʱ����������
Create Table Zltools.ZlRunLimitSet(
       ��� Number(5),
       ϵͳ Number(5),
       ģ�� Varchar2(18),
       ���� Varchar2(50),
       ����ѡ�� Number(1),
       ������� Number(3),
       ��ʱԭ�� Varchar2(250)
);
Alter Table Zltools.ZlRunLimitSet Add Constraint ZlRunLimitSet_PK Primary Key(���) Using Index;
Alter Table Zltools.ZlRunLimitSet Add Constraint ZlRunLimitSet_UQ_ģ�鹦�� Unique(ϵͳ,ģ��,����) Using Index;
Alter Table Zltools.ZlRunLimitSet Add Constraint ZlRunLimitSet_FK_������� Foreign Key(�������) References ZlRunLimit(���) On Delete Cascade;
Alter Table Zltools.ZlRunLimitSet Add Constraint ZlRunLimitSet_FK_ϵͳ Foreign Key(ϵͳ) References ZlSystems(���) On Delete Cascade;
CREATE INDEX Zltools.ZlRunLimitSet_IX_������� ON ZlRunLimitSet(�������);
Insert Into Zltools.zlTables(ϵͳ,����,��ռ�,����) Values(0,'ZLRUNLIMITSET','ZLTOOLSTBS','A2');

--120029:����,2018-01-12,��ʱ����������ZlRunLimitԤ������
Insert Into Zltools.ZlRunLimit(���,����,�Ƿ�����,����) Values(1,'Ԥ�跽��',1,'');

--120029:����,2018-01-12,��ʱ����������ZlRunLimitTimeԤ������
Insert Into Zltools.ZlRunLimitTime(ID,����,����,��ʼʱ��,����ʱ��) 
Select 1,1,0,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 2,1,1,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 3,1,2,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 4,1,3,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 5,1,4,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 6,1,5,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual Union All
Select 7,1,6,To_Date('1899-12-30 8:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('1899-12-30 12:00:00','YYYY-MM-DD HH24:MI:SS') From Dual;
Select Zltools.ZlRunLimitTime_ID.Nextval From Dual Connect By Rownum <= (Select Nvl(Max(ID), 0) From Zltools.ZlRunLimitTime);

--120029:����,2018-01-12,��ʱ����������ZlRunLimitSetԤ������
Insert Into Zltools.ZlRunLimitSet(���,ģ��,����,����ѡ��,�������,��ʱԭ��)
Select 1,'0102','��ǰ��Ǩ',1,1,'��ǰ��Ǩ��һ���̶Ȼ�Ӱ������ҵ��չ' From Dual Union All
Select 2,'0401','��ɫ��Ȩ',1,1,'��ɫ��Ȩ��Ի����򹫹�����������Ȩ���������SQL���½�����Ӱ������ҵ������' From Dual Union All
Select 3,'0401','�ָ����н�ɫ��Ȩ��',1,1,'�ָ����н�ɫ��Ȩ�޻�Ի����򹫹�����������Ȩ���������SQL���½�����Ӱ������ҵ������' From Dual Union All
Select 4,'0401','���ƽ�ɫ',1,1,'���ƽ�ɫ��Ի����򹫹�����������Ȩ���������SQL���½�����Ӱ������ҵ������' From Dual Union All
Select 5,'0402','�ָ������û���ɫ',1,1,'�ָ������û���ɫ��Ի����򹫹�����������Ȩ���������SQL���½�����Ӱ������ҵ������' From Dual Union All
Select 6,'0402','���������û���ɫ',1,1,'�����û���ɫ���������ɽ�ɫȨ�޿������ݣ�Ӱ��ҵ�����������' From Dual;

--120029:����,2018-01-12,��ʱ�����������޸�ģ�鹦�ܶ�Ӧ����������ѡ��Ĺ���
CREATE OR REPLACE Procedure zltools.Zl_Zlrunlimitset_Update
(
  ���_In     In Zlrunlimitset.���%Type,
  ����_In     In Zlrunlimitset.�������%Type := Null,
  ����ѡ��_In In Zlrunlimitset.����ѡ��%Type := Null,
  ��ʱԭ��_In In Zlrunlimitset.��ʱԭ��%Type := Null
) As
Begin
  If ����_In Is Null Then
    Update Zlrunlimitset Set ������� = ����_In Where ��� = ���_In;
  Else
    Update Zlrunlimitset Set ������� = ����_In, ����ѡ�� = ����ѡ��_In, ��ʱԭ�� = ��ʱԭ��_In Where ��� = ���_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlrunlimitset_Update;
/

--120029:����,2018-01-12,��ʱ�����������޸�ʱ���
CREATE OR REPLACE Procedure zltools.Zl_Zlrunlimittime_Update
(
  ����_In     In Number, --0:������1:�޸�,2:ɾ��
  Id_In       In Zlrunlimittime.Id%Type,
  ����_In     In Zlrunlimittime.����%Type := Null,
  ����_In     In Zlrunlimittime.����%Type := Null,
  ��ʼʱ��_In In Zlrunlimittime.��ʼʱ��%Type := Null,
  ����ʱ��_In In Zlrunlimittime.����ʱ��%Type := Null
) As
  n_Count   Number;
  d_Maxtime Date;
  d_Mintime Date;
Begin
  --��鵱ǰ����ʱ���Ƿ�������ʱ���г�ͻ
  --���У����¼��ͻ������䣬��ɾ����ͻʱ���
  If ����_In = 0 Or ����_In = 1 Then
    Select Count(1)
    Into n_Count
    From Zlrunlimittime
    Where ��ʼʱ�� <= ��ʼʱ��_In And ����ʱ�� >= ����ʱ��_In And ���� = ����_In And ���� = ����_In And ID <> Id_In;
    If n_Count = 0 Or (n_Count <> 0 And Id_In <> 0) Then
      Select Min(��ʼʱ��), Max(����ʱ��), Count(1)
      Into d_Mintime, d_Maxtime, n_Count
      From Zlrunlimittime
      Where (��ʼʱ�� >= ��ʼʱ��_In And ��ʼʱ�� <= ����ʱ��_In Or ����ʱ�� <= ����ʱ��_In And ����ʱ�� >= ��ʼʱ��_In Or
            ����ʱ�� >= ����ʱ��_In And ��ʼʱ�� <= ��ʼʱ��_In) And ���� = ����_In And ���� = ����_In And ID <> Id_In;
      If ��ʼʱ��_In < d_Mintime Then
        d_Mintime := ��ʼʱ��_In;
      End If;
      If ����ʱ��_In > d_Maxtime Then
        d_Maxtime := ����ʱ��_In;
      End If;
      If n_Count > 0 Then
        --˵���г�ͻ���ֶ�
        --�Ƚ���ͻ�ֶ�ɾ�����ٲ������ֶ�
        If Id_In <> 0 Then
          Delete Zlrunlimittime Where ID = Id_In;
        End If;
        Delete Zlrunlimittime
        Where (��ʼʱ�� >= ��ʼʱ��_In And ��ʼʱ�� <= ����ʱ��_In Or ����ʱ�� <= ����ʱ��_In And ����ʱ�� >= ��ʼʱ��_In Or
              ����ʱ�� >= ����ʱ��_In And ��ʼʱ�� <= ��ʼʱ��_In) And ���� = ����_In And ���� = ����_In;
        Insert Into Zlrunlimittime
          (ID, ����, ����, ��ʼʱ��, ����ʱ��)
        Values
          (Zlrunlimittime_Id.Nextval, ����_In, ����_In, d_Mintime, d_Maxtime);
      Else
        --˵��û�г�ͻ����ֱ�Ӷ����ݽ��в������²���
        If ����_In = 0 Then
          --����
          Insert Into Zlrunlimittime
            (ID, ����, ����, ��ʼʱ��, ����ʱ��)
          Values
            (Zlrunlimittime_Id.Nextval, ����_In, ����_In, ��ʼʱ��_In, ����ʱ��_In);
        Else
          --�޸�
          Update Zlrunlimittime Set ��ʼʱ�� = ��ʼʱ��_In, ����ʱ�� = ����ʱ��_In Where ID = Id_In;
        End If;
      End If;
    End If;
  Else
    --ɾ��
    Delete Zlrunlimittime Where ID = Id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlrunlimittime_Update;
/

--120029:����,2018-01-12,��ʱ�����������޸ķ���
Create Or Replace Procedure Zltools.Zl_Zlrunlimit_Update
(
  ����_In     In Number, --0:������1:�޸�,2:ɾ��
  ���_In     In Zlrunlimit.���%Type,
  ����_In     In Zlrunlimit.����%Type := Null,
  �Ƿ�����_In In Zlrunlimit.�Ƿ�����%Type := Null,
  ����_In     In Zlrunlimit.����%Type := Null
) As
  n_��� Zlrunlimit.���%Type;
Begin
  If ����_In = 0 Then
    --����
    Select Max(���) + 1 Into n_��� From Zlrunlimit;
    Insert Into Zlrunlimit (���, ����, �Ƿ�����, ����) Values (n_���, ����_In, �Ƿ�����_In, ����_In);
    Insert Into Zlrunlimittime
      (ID, ����, ����, ��ʼʱ��, ����ʱ��)
      Select Zlrunlimittime_Id.Nextval, n_���, ����, ��ʼʱ��, ����ʱ�� From Zlrunlimittime Where ���� = 1;
  Elsif ����_In = 1 Then
    --�޸�
    If �Ƿ�����_In Is Null Then
      --�޸ķ�����Ϣ����
      Update Zlrunlimit Set ���� = ����_In, ���� = ����_In Where ��� = ���_In;
    Else
      --��ͣ��������
      Update Zlrunlimit Set �Ƿ����� = �Ƿ�����_In Where ��� = ���_In;
    End If;
  Else
    --ɾ��
    Delete Zlrunlimit Where ��� = ���_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlrunlimit_Update;
/

--118267:����һ,2017-12-15,������LISͼƬת�湦��
Insert Into zlTools.Zlsvrtools(���,�ϼ�,����,���,˵��,����) Values('0208','02','����ͼƬ����ת��','U',Null,22);

--121074:����,2017-1-26,����ģ������
Update zlTools.zlSvrTools Set ���� = '��ʷ���ݿռ����' Where ��� = '0201';

--116852:����һ,2018-02-27,ɾ��ԭ��DBA����
Insert Into Zltools.Zlfilesexpired(�ļ���, ��װ·��, ϵͳ���, ϵͳ�汾, ˵��)Values('ZLDBAToolsEXE.exe', '[APPSOFT]', Null, '10.35.90', '��������,ɾ��ԭ�в���');
Delete From zlFilesUpgrade Where Upper(�ļ���) = 'ZLDBATOOLSEXE.EXE';
Delete From Zlfiles Where Upper(����) = 'ZLDBATOOLSEXE.EXE';