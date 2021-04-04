VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPassDefine_YWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------------------
' --ҩ��ʿ--- �ӿڶ���     �������б�������Ƽ����޹�˾��
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
Public Type YWS_BASE
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
'4)  ���ܺ� 5 Ҫ����ʾ
'HIS���� ��5
'<details_xml>
'<hosp_flag>����סԺ��ʶ</hosp_flag>
'<medicine>
'        <his_code>ҩƷ����</his_code>
'        <his_name>ҩƷ����</his_name>
'</medicine>
'</details_xml>

'5)  ���ܺ� 6 ��8��9��������
'HIS���� ��6��8��9
'<details_xml >
'    <his_time>HISϵͳʱ�䣨YYYY-MM-DD HH:mm:SS��</his_time>
'    <hosp_flag>����סԺ��ʶ op/ip</hosp_flag>
'    <treat_type>��������</treat_type>
'    <treat_code>�����</treat_code>
'    <bed_no>��λ��</bed_no>
'    <patient>
'        <name>����</name>
'        <birth>��������(YYYY-MM-DD)</birth>
'        <sex>�Ա���/Ů/δ֪��</sex>
'        <weight>���أ���λ��ǧ�ˣ�</weight>
'        <height>���ߣ���λ�����ף�</height>
'        <id_card>����֤��</id_card>
'        <medical_record>��������</medical_record>
'        <card_type>������</card_type>
'        <card_code>����</card_code>
'        <pregnant_unit>ʱ�䵥λ���졢�ܡ��£�</pregnant_unit>
'        <pregnant >����ʱ��</pregnant>
'        <allergic_data>
'            <!--�����ڵ㣬��ѭ��-->
'            <allergic>
'                <type>��������</type>
'                <name>����Դ����</name>
'                <code>����Դ����</code>
'            </allergic>
'        </allergic_data>
'        <diagnose_data>
'            <!--��Ͻڵ㣬��ѭ��-->
'            <diagnose>
'                <type>�������</type>
'                <name>�������</name>
'                <code>��ϴ���</code>
'            </diagnose>
'       </diagnose_data >
'    </patient>
'    <prescription_data>
'        <prescription>
'            <id>������</id>
'            <reason>��������</reason>
'            <is_current>�Ƿ�ǰ������0/1��</is_current>
'            <pres_type>����ҽ��L/��ʱҽ��T</pres_type>
'            <pres_time>����ʱ�䣨YYYY-MM-DD HH:mm:SS��</pres_time>
'            <medicine_data>
'           <!--ҩƷ��Ϣ�ڵ㣬��ѭ��-->
'                <medicine>
'                     <zxy_type>��ҩ/�г�ҩ/��ҩ</zxy_type>
'                     <oeridid>������-��ˮ��(Ψһ)</oeridid>
'                     <pres_type>����ҽ��L/��ʱҽ��T</pres_type>
'                     <pres_time>����ʱ�䣨YYYY-MM-DD HH:mm:SS��</pres_time>
'                     <name>��Ʒ��</name>
'                     <his_code>ҽԺҩƷ����</his_code>
'                     <insur_code>ҽ������</insur_code>
'                     <approval>��׼�ĺ�</approval>
'                     <spec>���</spec>
'                     <group>���</group>
'                     <reason>��ҩ����</reason>
'                     <dose_unit>��������λ</dose_unit>
'                     <dose>������</dose>
'                     <freq>Ƶ�δ���</freq>
'                     <administer>��ҩ;������</administer>
'                     <begin_time>��סԺ����ҩ��ʼʱ��(YYYY-MM-DD HH:mm:SS)</begin_time>
'                     <end_time>��סԺ����ҩ����ʱ��(YYYY-MM-DD HH:mm:SS)</end_time>
'                     <days>��ҩ���������</days>
'                </medicine>
'           </medicine_data>
'       </prescription>
'   </prescription_data>
'</details_xml>
'
'����Դ����
Public Type YWS_ALLERGIC
    str�������� As String
    str����Դ���� As String
    str����Դ���� As String
End Type

'��϶���
Public Type YWS_DIAGNOSE
    str������� As String
    str������� As String
    str��ϴ��� As String
End Type
'ҩƷ��Ϣ
Public Type YWS_MEDICINE
    strҩƷ���� As String
    str������ As String
    Strҽ������ As String
    str����ʱ�� As String
    str��Ʒ�� As String
    strҽԺҩƷ���� As String
    strҽ������ As String
    str��׼�ĺ� As String
    str��� As String
    str��� As String
    str��ҩ���� As String
    str��������λ As String
    str������ As String
    strƵ�δ��� As String
    str��ҩ;������ As String
    str��ҩ��ʼʱ�� As String
    str��ҩ����ʱ�� As String
    str��ҩ���� As String      '����
End Type
'������Ϣ
Public Type YWS_PATIENT
    str���� As String
    str�������� As String '��ʽ��YYYY - MM - DD
    str�Ա� As String   'ֵ��Χ:�� / Ů / δ֪
    str���� As String '��λ kg
    str���� As String '��λ cm
    str����֤�� As String
    str�������� As String
    str������ As String
    str���� As String
    str����ʱ�䵥λ As String '��,��,��
    str����ʱ�� As String
    col����Դs  As Collection
    col���s As Collection
End Type
'������Ϣ
Public Type YWS_PRESCRIPTION
    str������ As String
    str�������� As String
    str�Ƿ�ǰ���� As String   '��0/1��
    Strҽ������ As String
    str����ʱ�� As String    '��YYYY-MM-DD HH:mm:SS��
    colҩƷ��Ϣ As Collection     'ҩƷ��
End Type
'DETAILS_XML
Public Type YWS_DETAILS
    str����סԺ��ʶ As String ' op/ip
    strҩƷ���� As String
    strҩƷ���� As String
    strHISϵͳʱ�� As String  '��YYYY-MM-DD HH:mm:SS��
    str�������� As String
    str����� As String
    str��λ�� As String
    udt������Ϣ As YWS_PATIENT      '������Ϣ
    udt������Ϣ As YWS_PRESCRIPTION      '������Ϣ
End Type

Public Enum YWS_Func_NUM
    YWS_��¼ = 1
    YWS_�˳� = 3
    YWS_��ʼ�ͻ��� = 4
    YWS_Ҫ����ʾ = 5
    YWS_�������� = 6
    YWS_�ϴ����� = 8
    YWS_�������������� = 9
    YWS_���Ե��� = 10
    YWS_�������������ؽ�� = 11
End Enum
'<-end------------------------------------------------------------------------details_xml----------------------------------------------------->