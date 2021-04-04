VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaptureParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public VideoDirverType As TVideoDriverType      '��Ƶ��������
Public IsSountHint As Boolean                   '������ʾ
Public IsWindowHint As Boolean                  '������ʾ
Public ScanDeviceTmpDir As String               'ɨ���豸��ʱĿ¼
Public CustomDevicePath As String                'ר����Ƶ�ɼ�����·��

Public ComPortType As String                    '�˿�����
Public ComPortName As String                    'ʹ�õĶ˿�����
Public ComInterval As Long                      '��̤�ɼ����ʱ��
Public CaptureWay As Long                       '��̤�ɼ���ʽ

Public ShowBigImage As Integer                '�Ƿ���ʾ��ͼ��0-����ʾ��1-����ƶ�ʱ��ʾ��2-����ʱ��ʾ
Public IsAllowChangeSize As Boolean              '�Ƿ������ı�ɼ������С
Public ImagePreview As Integer                 '����ͼԤ����0-��Ԥ����1-�ƶ����Ԥ����2-����Ԥ��
Public PreviewTime As Long                      '����ͼԤ����ʱ�ر�ʱ��

Public IsUseAfterCapture As Boolean             '�Ƿ����ú�̨�ɼ�
Public AfterCaptureModality As String           '��̨Ӱ�����
Public CurStorageDeviceNo As String             '�洢�豸��
Public BakStorageDeviceNo As String             '�����豸��
Public IsUseCaptureLock As Boolean              '�Ƿ����òɼ�����

Public strCaptureHot As String                  '�ɼ��ȼ�


Public Sub ReadParameter()
'��ʼ����������
    Dim rsTmp As New ADODB.Recordset
    Dim intVideoCapture As Integer
    Dim strSQL As String
    
BUGEX "ReadParameter 1", True
    IsUseAfterCapture = False
    AfterCaptureModality = "OT"

    
BUGEX "ReadParameter 2", True
    '��ȡ��������
    VideoDirverType = zlDatabase.GetPara("��Ƶ��������", glngSys, glngModule, "0")
    
    
    '��ȡ��ʾ����
    IsSountHint = Val(zlDatabase.GetPara("�ɼ���������ʾ", glngSys, glngModule, True))
    IsWindowHint = Val(zlDatabase.GetPara("�ɼ��󵯴���ʾ", glngSys, glngModule, True))
    
    '��ȡɨ���豸��ʱ�洢��ͼ��Ŀ¼
    ScanDeviceTmpDir = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ɨ���豸��ʱĿ¼", "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA")
  
    'ר����Ƶ�ɼ�
    CustomDevicePath = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "ר����Ƶ�ɼ�")
    
    '��ȡ���ڵĲ���
    ComPortName = zlDatabase.GetPara("��̤�˿�", glngSys, glngModule, "1")
    If IsNumeric(ComPortName) Then
        'ʹ��Com�ɼ��˿�
        ComPortName = CLng(ComPortName)
        ComPortType = "COM"
        
        CaptureWay = zlDatabase.GetPara("��̤�ɼ���ʽ", glngSys, glngModule, "1")
        If CaptureWay < 0 Or CaptureWay > 2 Then CaptureWay = 1
        
        '��ȡ��̤���ʱ��
        ComInterval = zlDatabase.GetPara("��̤ʱ����", glngSys, glngModule, "1")
    ElseIf ComPortName = "��" Or Trim(ComPortName) = "" Then
        '��ʹ�òɼ��˿�
        ComPortType = "��"
    Else
        'ʹ��HID�ɼ��豸
        ComPortType = "HID"
    End If
BUGEX "ReadParameter 3", True

    'strCaptureHot = zlDatabase.GetPara("�ɼ��ȼ�", glngSys, glngModule, "F8")
    strCaptureHot = GetSetting("ZLSOFT", "����ģ��", "�ɼ��ȼ�", "F8")

    
        
BUGEX "ReadParameter 4", True
    '��ʼ�����Ҽ�����==============================================================================
    If glngModule = 1294 Then
        IsUseAfterCapture = False
    Else
        IsUseAfterCapture = GetDeptPara(glngDepartId, "���ú�̨�ɼ�", 0)
    End If
    AfterCaptureModality = GetDeptPara(glngDepartId, "��̨Ӱ�����", "OT")
    
    '��ȡ�����洢�豸��
    CurStorageDeviceNo = GetDeptPara(glngDepartId, "�洢�豸��")
    BakStorageDeviceNo = GetDeptPara(glngDepartId, "�����豸��")
    
    '���òɼ�����
    IsUseCaptureLock = GetDeptPara(glngDepartId, "���òɼ�����", True)
    '�ı�ɼ������С
    IsAllowChangeSize = GetDeptPara(glngDepartId, "�����ı�ɼ������С", True)
    
    ImagePreview = Val(GetDeptPara(glngDepartId, "����ͼԤ����ʽ", "0", True))
    PreviewTime = Val(GetDeptPara(glngDepartId, "�ƶ�Ԥ����ʱ", "0", True))
    
BUGEX "ReadParameter 5", True
End Sub