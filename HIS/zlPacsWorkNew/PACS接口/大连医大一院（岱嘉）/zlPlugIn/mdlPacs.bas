Attribute VB_Name = "mdlPacs"
Option Explicit

''''''''���˵��''''''''''''''''''''''
'1��CallPACSView.dll��CallPACSView.lib�����ļ����Ϻ�᷼��ṩ��HISҽ��վ���õĽӿ��ļ���
'2���������ӿ��ļ���Ҫ�ŵ�Windows/System32Ŀ¼�£�����Ҫע�ᡣ
'3�����ӿ���ֱ��д��᷼��ڴ���ҽ��һԺ��RIS������IP��192.9.200.6�� WEB������IP��192.9.200.9���û���������ȣ����ֻ����ҽԺ�������������С�
''''''''''''''''''''''''''''''''''''''''''''''''''

'��ʼ��pacs����
Public Declare Function InitPACSConnection Lib "CallPACSView.dll" ( _
    ByVal strRisIp As String, _
    ByVal strRisUser As String, _
    ByVal strRisPwd As String, _
    ByVal strRisDbName As String, _
    ByVal strPacsIp As String, _
    ByVal strPacsUser As String, _
    ByVal strPacsPwd As String, _
    ByVal strPacsDbName As String _
    ) As String


'����pacsӰ��
Public Declare Function CallPACSView Lib "CallPACSView.dll" ( _
    ByVal strAdviceId As String, _
    ByVal strWebIp As String, _
    ByVal strWebUser As String, _
    ByVal strWebPwd As String, _
    ByVal blnIsOpenImage As Boolean _
    ) As String

Public Const gstrFunc_PACSӰ����� = "PACSӰ�����"

Private Const ᷼�RIS������IP = "192.9.200.6"
Private Const ᷼�RIS�û��� = "zlhis"
Private Const ᷼�RIS���� = "his"
Private Const ᷼�RIS���ݿ��� = "UniRISCDB"
'RIS�����������û��� sa,���� ats

Private Const ᷼�PACS������IP = "192.9.200.6"
Private Const ᷼�PACS�û��� = "zlhis"
Private Const ᷼�PACS���� = "his"
Private Const ᷼�PACS���ݿ��� = "DICOMDB"
'PACS�����������û��� sa,���� ats

Private Const ᷼�WEB������IP = "192.9.200.9"
Private Const ᷼�WEB�û��� = "user"
Private Const ᷼�WEB���� = "1"


Private blnInitPacsConnection As Boolean        '�Ƿ���Ҫ��ʼ��PACS����

Public Function InitPacs() As Boolean
'��ʼ��᷼ε�PACS���ݿ�����

    Dim strErr As String
    
    On Error GoTo err
    
    InitPacs = False
    
    strErr = InitPACSConnection(᷼�RIS������IP, ᷼�RIS�û���, ᷼�RIS����, ᷼�RIS���ݿ���, _
                                ᷼�PACS������IP, ᷼�PACS�û���, ᷼�PACS����, ᷼�PACS���ݿ���)


    If strErr <> "�ɹ�" Then
        MsgBox strErr, vbOKOnly, "PACSӰ��ӿ�"
        Exit Function
    End If
    
    InitPacs = True
    Exit Function
err:
    MsgBox err.Description, vbOKOnly, "PACSӰ��ӿ�"
    err.Clear
End Function

Public Function ShowPacsViewer(ByVal varKeyId As Variant) As Boolean
'����᷼ε�CallPACSView��������IE��ʾWEB�汾��PACSͼ�������
    Dim strErr As String
    
    On Error GoTo err
    
    ShowPacsViewer = False
    
    '�ȳ�ʼ��
    'ֻ�������ҽ��վ��סԺҽ��վ�ų�ʼ��PACS����ͼ��Ĳ��
    If blnInitPacsConnection = False Then
        blnInitPacsConnection = InitPacs
    End If
        
    If blnInitPacsConnection = True Then
        strErr = CallPACSView(CStr(varKeyId), ᷼�WEB������IP, ᷼�WEB�û���, ᷼�WEB����, False)
        If strErr <> "�ɹ�" Then
            MsgBox strErr, vbOKOnly, "PACSӰ��ӿ�"
            Exit Function
        End If
    
        ShowPacsViewer = True
    End If
    
    Exit Function
err:
    MsgBox err.Description, vbOKOnly, "PACSӰ��ӿ�"
    err.Clear
End Function
