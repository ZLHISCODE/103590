VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPassDefine_DTBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------------------
' --��ͨBS--- �ӿڶ���<CRMs-MR-ITF><CRMs/ǰ̨�ӿ�>�汾 2.2���Ϻ���ͨҽҩ��Ϣ�������޹�˾��
'--------------------------------------------------------------------------------------------------------------------------------------
'<-begin------------------------------------------------------------------------base_xml----------------------------------------------------->
'<base_xml>
'    <source>HIS</source>
'    <hosp_code>ҽԺ����</hosp_code>
'    <dept_code>���Ҵ���</dept_code>
'    <dept_name>��������</dept_name>
'    <doct>
'        <code>ҽ������</code>
'        <name>ҽ������</name>
'        <type>ҽ���������</type>
'        <type_name>ҽ����������</type_name >
'    </doct>
'</base_xml>
Public Type DTBS_BASE
    strHIS As String
    strҽԺ���� As String
    str���Ҵ��� As String
    str�������� As String
    strҽ������ As String
    strҽ������ As String
    strҽ��������� As String
    strҽ���������� As String
End Type
'<-end------------------------------------------------------------------------base_xml----------------------------------------------------->

'<-begin------------------------------------------------------------------------details_xml ��ʽ����:----------------------------------------------------->
'HIS���� ��1-��¼
'details_xml ��ʽ����:
'<details_xml>
'<doct_pwd>����</doct_pwd>
'</details_xml>

'HIS���� ��2 �˳�
'details_xml ��ʽ����:
'<details_xml></details_xml>

'HIS���� ��3 ��ʼ��UI
'details_xml ��ʽ����:
'<details_xml></details_xml>

'
'HIS���� ��5 Ҫ����ʾ
'details_xml ��ʽ����:
'<details_xml>
'<hosp_flag>����סԺ��ʶ</hosp_flag>
'<medicine>
'        <his_code>ҩƷ����</his_code>
'        <his_name>ҩƷ����</his_name>
'</medicine>
'</details_xml>

'HIS���� ��6 ��������
'details_xml ��ʽ����:
'<details_xml is_upload ="�Ƿ��ϴ�0/1">
'   <his_time>HISϵͳʱ�䣨YYYY-MM-DD HH:mm:SS��</his_time>
'   <hosp_flag>����סԺ��ʶ op/ip</hosp_flag>
'   <treat_type>��������</treat_type>
'   <treat_code>�����</treat_code>
'   <lis_adm_no>��������</lis_adm_no>
'   <bed_no>��λ��</bed_no>
'   <area_code>������</area_code>
'   <patient>
'       <name>����</name>
'       <is_infant>�Ƿ�Ӥ�׶�</is_infant>
'       <birth>��������(YYYY-MM-DD)</birth>
'       <sex>�Ա���/Ů/δ֪��</sex>
'       <weight>���أ���λ��ǧ�ˣ�</weight>
'       <height>���ߣ���λ�����ף�</height>
'       <id_card>����֤��</id_card>
'       <card_type>������</card_type>
'       <card_code>����</card_code>
'       <pregnant_unit>ʱ�䵥λ���졢�ܡ��£�</pregnant_unit>
'       <pregnant >����ʱ��</pregnant>
'       <allergic_data>
'           <!-�����ڵ㣬��ѭ��'
'           <allergic>
'               <type>��������</type>
'               <name>����Դ����</name>
'               <code>����Դ����</code>
'           </allergic>
'       </allergic_data>
'       <diagnose_data>
'        -!--��Ͻڵ㣬��ѭ��-->
'           <diagnose>
'               <type>�������</type>
'               <name>�������</name>
'               <code>��ϴ���</code>
'           </diagnose>
'       </diagnose_data >
'       <lis_data>
'       <!-�����ⵥ�ڵ㣬��ѭ��-->
'           <form>
'               <no>���顢��鵥��</no>
'               <project_name>������Ŀ����</project_name>
'               <lis_flag>���顢����� 0-����, 1-���</lis_flag>
'               <result_date>���顢���������ʱ��</result_date>
'               <sample_code>������������</sample_code>
'               <sample_name>������������</sample_name>
'               <mac_flag>΢�����ͼ��ʶ 0-��, 1-��</mac_flag>
'               <!-��������Ŀ�ڵ㣬��ѭ��-->
'               <item>
'                   <code>���顢������</code>
'                   <name>���顢�������</name>
'                   <value>���顢�����</value>
'                   <uom>���ֵ�ĵ�λ</uom>
'                   <upper>����ο���Χ����</upper>
'                   <lower>����ο���Χ����</lower>
'               </item>
'           </form>
'       </lis_data>
'   </patient>
'   <prescription_data>
'       -!--������Ϣ�ڵ㣬��ѭ��-->
'       <prescription>
'           <id>������</id>
'           <reason>��������</reason>
'           <is_urgent>�Ƿ��������(0:��1:��)</ is_urgent>
'           <is_new>�Ƿ��¿�����(0:��1:��)</ is_new>
'           <is_current>�Ƿ�ǰ������0/1��</is_current>
'           <doct_code>����ҽ������</doct_code>
'           <doct_name>����ҽ������</doct_name>
'           <dept_code>�������Ҵ���</dept_code>
'           <dept_name>������������</dept_name>
'           <pres_type>����ҽ��L/��ʱҽ��T</pres_type>
'           <pres_time>����ʱ�䣨YYYY-MM-DD HH:mm:SS��</pres_time>
'           <medicine_data>
'               -!--ҩƷ��Ϣ�ڵ㣬��ѭ��-->
'               <medicine>
'                   <name>��Ʒ��</name>
'                   <his_code>ҽԺҩƷ����</his_code>
'                   <insur_code>ҽ������</insur_code>
'                   <pyd_code>��Һ����</pyd_code>
'                   <link_group>��Һ�����</ link_group>
'                   <spec>���</spec>
'                   <group>���</group>
'                   <reason>��ҩ����</reason>
'                   <dose_unit>��������λ</dose_unit>
'                   <dose>������</dose>
'                   <freq>Ƶ�δ���</freq>
'                   <administer>��ҩ;������</administer>
'                   <begin_time>��סԺ����ҩ��ʼʱ��(YYYY-MM-DD HH:mm:SS)</begin_time>
'                   <end_time>��סԺ����ҩ����ʱ��(YYYY-MM-DD HH:mm:SS)</end_time>
'                   <days>��ҩ���������</days>
'                   < preventiveflag >��סԺ���Ƿ�Ԥ����ҩ��1�ǣ�0��        </preventiveflag >
'                   <otno>��סԺ����������</otno>
'                   <signer_code>ǩ��ҽʦ����</signer_code>
'                   <accredit_date>��Ȩʱ��</accredit_date>
'                   <accredit_hours>������ҩʱ�䣨Сʱ��</accredit_hours >
'                   <accredit_times>������ҩ����</accredit_times>
'               </medicine>
'           </medicine_data>
'       </prescription>
'   </prescription_data>
'</details_xml>

'����Դ����
Public Type DTBS_ALLERGIC
    str�������� As String
    str����Դ���� As String
    str����Դ���� As String
End Type

'��϶���
Public Type DTBS_DIAGNOSE
    str������� As String
    str������� As String
    str��ϴ��� As String
End Type
'ҩƷ��Ϣ
Public Type DTBS_MEDICINE
    str��Ʒ�� As String
    strҽԺҩƷ���� As String
    strҽ������ As String
    str��Һ���� As String
    str��Һ����� As String
    str��� As String
    str��� As String
    str��ҩ���� As String
    str��������λ As String
    str������ As String
    strƵ�δ��� As String
    str��ҩ;������ As String
    str��ҩ��ʼʱ�� As String 'סԺ
    str��ҩ����ʱ�� As String 'סԺ
    str��ҩ���� As String      '����
    str�Ƿ�Ԥ����ҩ As String 'סԺ
    str�������� As String     'סԺ
    strǩ��ҽʦ���� As String
    str��Ȩʱ�� As String
    str������ҩʱ�� As String
    str������ҩ���� As String
End Type
'������Ϣ
Public Type DTBS_PATIENT
    str���� As String
    str�Ƿ�Ӥ�� As String
    str�������� As String '��ʽ��YYYY - MM - DD
    str�Ա� As String   'ֵ��Χ:�� / Ů / δ֪
    str���� As String '��λ kg
    str���� As String '��λ cm
    str����֤�� As String
    str������ As String
    str���� As String
    str����ʱ�䵥λ As String '��,��,��
    str����ʱ�� As String
    col����Դs  As Collection
    col���s As Collection
    col������ As Collection
End Type
'������Ϣ
Public Type DTBS_PRESCRIPTION
    str������ As String
    str�������� As String
    str�Ƿ�ǰ���� As String   '��0/1��
    str�Ƿ�������� As String
    str�Ƿ��¿����� As String
    str����ҽ������ As String
    str����ҽ������ As String
    str�������Ҵ��� As String
    str������������ As String
    Strҽ������ As String    '����ҽ��L/��ʱҽ��T
    str����ʱ�� As String    '��YYYY-MM-DD HH:mm:SS��
    colҩƷ��Ϣ As Collection     'ҩƷ��
End Type
'DETAILS_XML
Public Type DTBS_DETAILS
    str�Ƿ��ϴ� As String
    strHISϵͳʱ�� As String  '��YYYY-MM-DD HH:mm:SS��
    str����סԺ��ʶ As String ' op/ip
    str�������� As String
    str����� As String
    str�������� As String
    str��λ�� As String
    str������ As String
    udt������Ϣ As DTBS_PATIENT      '������Ϣ
    col������Ϣ  As Collection
    strҩƷ���� As String
    strҩƷ���� As String
End Type
'FORM_ITEM
Public Type DTBS_LISITEM
    str���� As String  '��YYYY-MM-DD HH:mm:SS��
    str���� As String ' op/ip
    str��� As String
    str���ֵ��λ As String
    str�ο���Χ���� As String
    str�ο���Χ���� As String
End Type

'LIS_FORM
Public Type DTBS_LISFORM
    str���� As String
    str��Ŀ���� As String
    str��� As String  '���顢����� 0-����, 1-���
    str�������ʱ�� As String
    str������������ As String
    str������������ As String
    str΢�����ͼ��ʶ As String
    col��Ŀ�ڵ� As Collection    'DTBS_LISITEM �ļ���
End Type

Public Enum DTBS_Func_NUM
    DTBS_��¼ = 1
    DTBS_�˳� = 2
    DTBS_��ʼUI = 3
    DTBS_Ҫ����ʾ = 5
    DTBS_�������� = 6
    DTBS_���ϴ��� = 13
    DTBS_��ȡ��״̬ = 14
End Enum
'<-end------------------------------------------------------------------------details_xml----------------------------------------------------->
