Attribute VB_Name = "mdlPublic"
Option Explicit

Public gcnOracle As ADODB.Connection        'HIS���ñ�����ʱ����Ĺ������ݿ�����
Public gstrUserCode As String   '��ǰ����Ա���
Public gstrUserName As String   '��ǰ����Ա����

Public Enum ModulNO
    FOutBillPrint = 1121    '�����շ�
    FInBillPrint = 1137     'סԺ����
End Enum
Public glngSys As Long          '��ǰ����ϵͳ��ţ�100=ZLHIS��׼��
Public glngModul As ModulNO        '��ǰ����ģ��ţ�1121=�����շ�,1137=סԺ����

