Attribute VB_Name = "mdlPubPar"
Option Explicit
Public gcnOracle As ADODB.Connection      'ȫ������
Public gstrSQL As String                  'ȫ�ֹ���
Public gstrDbOwner As String              '���ݿ�ӵ����
Public glngSys As Long                    'ϵͳ���
Public gstrProductName As String          '��������
Public gstrSysName As String              'ϵͳ����
Public gstrAviPath As String              'AVI·��
Public gstrVersion As String              '�汾
Public gstrMatch As String                'ƥ��ģʽ
Public gobjFSO As New Scripting.FileSystemObject    'FSO����
Public gbytEsign As Byte              '�Ƿ����õ���ǩ�� 0-���룻1������                '
Public gAllFont As Collection
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO
