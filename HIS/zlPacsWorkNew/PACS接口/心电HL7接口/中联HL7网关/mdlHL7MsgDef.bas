Attribute VB_Name = "mdlHL7MsgDef"
Option Explicit

'��ģ����Ҫ��HL7��Ϣ�ṹ�Ķ���

'HL7�ֶζ���
Public Type THL7Field
    intNo As Integer            '��ţ������������
    strDataType As String       '�������ͣ�HL7�ж��������
    strRecDataDef As String     '�������ݶ���
    strSendDataDef As String    '�������ݶ���
    strRecDataValue As String   '��������ֵ
    strSendDataValue As String  '��������ֵ
    strElementName As String    'Ԫ������
    blnEnable As Boolean        '������
End Type

'HL7��Ϣ�ζ���
Public Type THL7Segment
    intNo As Integer            '��Ϣ�ε���ţ������������
    arrFields() As THL7Field    '�����ֶζ���
    strName As String           '��Ϣ������
    strText As String           '��Ϣ���ı������ջ��߷��͵��ı�
    blnEnable As Boolean        '������
End Type

'HL7��Ϣ
Public Type THL7Message
    lngID As Long                   '��ϢID
    arrSegments() As THL7Segment    '��Ϣ�ζ���
    strMsgName As String            '��Ϣ����
    lngServiceID As Long            '����ID
    strActionType As String         '��������
    strMsgType As String            '��Ϣ����
    strMsgSegmentDef As String      '��Ϣ����϶���
    strText As String               '��Ϣ�ı�
    strIP As String                 '������Ϣ��IP��ַ
    lngPort As Long                 '������Ϣ�Ķ˿ں�
    blnSendOK As Boolean            '������Ϣ�ɹ������յ�AA����Ӧ
End Type

Public Type THl7Messages
    arrMsgs() As THL7Message        '��Ϣ����
    lngActionID As Long             'HL7������Ϣ��¼��ID
End Type

'��������
Public Const HL7_MSG_SEND_NEW_ORDER = "������ҽ��"
Public Const HL7_MSG_SEND_CANCEL_ORDER = "����ȡ��ҽ��"
Public Const HL7_MSG_SEND_DEL_ORDER = "����ɾ��ҽ��"

