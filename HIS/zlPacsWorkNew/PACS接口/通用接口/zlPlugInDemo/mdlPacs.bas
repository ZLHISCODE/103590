Attribute VB_Name = "mdlPacs"
Option Explicit

''''''''���˵��''''''''''''''''''''''
'''˵����''''''''''''''''''''''''''''''''''''''''''
'''1�������ӳ����У�����PACS�Ĳ�����Ҫ��mdlPacsģ����ʵ�֡�
'''2��ͨ������PACS�����漰���������֣���ʼ������PACS���򣨹�Ƭ���߱��棩���������ͷ���Դ��
'''3������������ҽ�ṩ��PACS�ӿ�Ϊ��





'''''''''''''''''''''''''''''''''''''''''''''''''''''
'1��XEFORHIS.dll����ҽ�ṩ��HISҽ��վ���õĽӿ��ļ���
''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''˵�����ڴ˴�����PACS�ṩ��DLL����д��Ӧ�Ķ������
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function XePacsInit Lib "XEFORHIS.dll" () As Boolean
Public Declare Function XePacsCall Lib "XEFORHIS.dll" ( _
    ByVal nPatientIDType As Long, _
    ByVal lpszID As String, _
    ByVal nCallType As Long _
) As Boolean

Public Declare Function XePacsRelease Lib "XEFORHIS.dll" ()



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''˵�����˴�����PACS�Ĺ��ܣ������Ӧ�Ĺ�����������������ģ��clsPlugIn�еĵ���
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const gstrFunc_PACSӰ����� = "PACSӰ�����"
Public Const gstrFunc_PACS������� = "PACS�������"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''˵�����˴�������Ҫʹ�õĹ�������
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private blnInitPacsConnection As Boolean        '�Ƿ���Ҫ��ʼ��PACS���ӣ�PACS����ֻ��Ҫ��ʼ��һ��



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''˵�����˴���д����PACS�ӿڵľ�����̺ͺ���
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InitPacs() As Boolean
'��ʼ����PACS�����ݿ�����

'��ʼ����ҽ��PACS���ݿ�����

    Dim blnErr As Boolean
    
    On Error GoTo err
    
    InitPacs = False
    
    blnErr = XePacsInit


    If blnErr = False Then
        MsgBox "��ʼ�����ݴ���", vbOKOnly, "PACSӰ��ӿ�"
        Exit Function
    End If
    
    InitPacs = True
    Exit Function
err:
    MsgBox err.Description, vbOKOnly, "PACSӰ��ӿ�"
    err.Clear
End Function

Public Function ShowPacsViewer(ByVal varKeyId As Variant, lngViewType As Integer) As Boolean
'˵������ʾPACS��ͼ��������ͱ��������������PACS�ṩ�Ĳ������ñ����̵Ĳ�������������

'������ҽ��XePacsCall��������IE��ʾWEB�汾��PACSͼ�������
'������ varKeyId --- ҽ��ID
'       lngViewType --- �����ʽ��1Ϊ�鿴ͼ��2Ϊ�鿴����
    Dim blnErr As Boolean
    
    
    On Error GoTo err
    
    ShowPacsViewer = False
    
    '�ȳ�ʼ��
    'ֻ�������ҽ��վ��סԺҽ��վ�ų�ʼ��PACS����ͼ��Ĳ��
    If blnInitPacsConnection = False Then
        Dim lngWait As Long
        blnInitPacsConnection = InitPacs
            
        'ѭ��ֻ��Ϊ����ʱ����ҽ�Ľӿڳ�ʼ��֮��ֱ�ӵ���ͼ�񣬻���ʾ������Ҫ��һ����ʱ
        For lngWait = 1 To 6000
        
        Next lngWait
        
    End If
        
        
    'XePacsCall ����˵����  nPatintIDType ������ͣ�1������ţ�2��סԺ�ţ�3�����뵥��
    '                       nCallType �������ͣ�1���鿴ͼ��2���鿴����
    '���ù�XePacsInit�󣬼��ɵ��ñ��������鿴ͼ��򱨸�
    
    If blnInitPacsConnection = True Then
        blnErr = XePacsCall(3, CStr(varKeyId), lngViewType)
        If blnErr = False Then
            MsgBox IIf(lngViewType = 1, "����ͼ��������", "���ı��淢������"), vbOKOnly, "PACSӰ��ӿ�"
            Exit Function
        End If
    
        ShowPacsViewer = True
    End If
    
    Exit Function
err:
    MsgBox err.Description, vbOKOnly, "PACSӰ��ӿ�"
    err.Clear
End Function

Public Function PacsRelease()
'˵���� ����PACS������ͷ����ӣ�ֻ��Ҫ����һ��

'������ҽXePacsRelease�������ͷ����ݿ�����
    On Error GoTo err
    
    If blnInitPacsConnection = True Then
        XePacsRelease
    End If
    
    Exit Function
err:
   MsgBox err.Description, vbOKOnly, "PACSӰ��ӿ�"
    err.Clear
End Function


