Attribute VB_Name = "mdlPass"
Option Explicit
'PASS�ӿں���������˵���μ�PASS�ӿ��ĵ�����
'˵����ShellRunAs.dll��Ҫ��װ�ڳ����ϵͳĿ¼
'      DIFPassDll.dllΪPassϵͳ�Զ���ȡ��ע��·��
'ע�������
Public Declare Function RegisterServer Lib "ShellRunAs.dll" () As Integer
'PASS��ʼ��
Public Declare Function PassInit Lib "DIFPassDll.dll" ( _
    ByVal UserName As String, _
    ByVal DepartMentName As String, _
    ByVal WorkstationType As Integer) As Integer
'PASS����ģʽ����
Public Declare Function PassSetControlParam Lib "DIFPassDll.dll" ( _
    ByVal SaveCheckResult As Integer, _
    ByVal AllowAllegen As Integer, _
    ByVal CheckMode As Integer, _
    ByVal DisqMode As Integer, _
    ByVal UseDiposeIdea As Integer) As Integer
'�����˻�����Ϣ
Public Declare Function PassSetPatientInfo Lib "DIFPassDll.dll" ( _
    ByVal PatientID As String, _
    ByVal VisitID As String, _
    ByVal Name As String, _
    ByVal Sex As String, _
    ByVal Birthday As String, _
    ByVal Weight As String, _
    ByVal cHeight As String, _
    ByVal DepartMentName As String, _
    ByVal Doctor As String, _
    ByVal LeaveHospitalDate As String) As Integer
'������ҩƷ��Ϣ
Public Declare Function PassSetRecipeInfo Lib "DIFPassDll.dll" ( _
    ByVal OrderUniqueCode As String, _
    ByVal DrugCode As String, _
    ByVal DrugName As String, _
    ByVal SingleDose As String, _
    ByVal DoseUnit As String, _
    ByVal Frequency As String, _
    ByVal StartOrderDate As String, _
    ByVal StopOrderDate As String, _
    ByVal RouteName As String, _
    ByVal GroupTag As String, _
    ByVal OrderType As String, _
    ByVal OrderDoctor As String) As Integer
'������Ҫ���е�ҩ�����ҩƷ
Public Declare Function PassSetWarnDrug Lib "DIFPassDll.dll" (ByVal DrugUniqueCode As String) As Integer
'��Ϣ��ѯҩƷ����
Public Declare Function PassSetQueryDrug Lib "DIFPassDll.dll" ( _
    ByVal DrugCode As String, _
    ByVal DrugName As String, _
    ByVal DoseUnit As String, _
    ByVal RouteName As String) As Integer
'��ȡ�Ҽ��˵��Ƿ����ֵ
Public Declare Function PassGetState Lib "DIFPassDll.dll" (ByVal QueryItemNo As String) As Integer
'PASS���ܵ���
Public Declare Function PassDoCommand Lib "DIFPassDll.dll" (ByVal CommandNo As Integer) As Integer
'��ȡҩƷ��ʾ����
Public Declare Function PassGetWarn Lib "DIFPassDll.dll" (ByVal DrugUniqueCode As String) As Integer
'����ҩƷ��������λ��
Public Declare Function PassSetFloatWinPos Lib "DIFPassDll.dll" ( _
    ByVal Left As Integer, ByVal Top As Integer, _
    ByVal Right As Integer, ByVal Bottom As Integer) As Integer
'PASS�˳�����
Public Declare Function PassQuit Lib "DIFPassDll.dll" () As Integer
'------------------------------------------------------------------
'ZLHIS���Ƿ�ʹ��PASSϵͳ
Public gblnPass As Boolean

Public Function PassInitialize() As Boolean
'���ܣ���PASS�ӿڽ���ע��ͳ�ʼ����ͬʱ���PASS�ӿ�DLL�Ƿ���ȷ��װ
    On Error GoTo errH
    
    'PASS���ܺ���ע��(����ͻ���ģʽ)
    If RegisterServer <> 0 Then
        MsgBox "PASS�ͻ���ע��ʧ�ܣ���ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'PASS��ʼ��
    If PassInit(UserInfo.��� & "/" & UserInfo.�û���, UserInfo.������ & "/" & UserInfo.������, 10) <> 1 Then
        MsgBox "PASSϵͳ��ʼ��ʧ�ܣ���ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
            
    'PASS�Ƿ���ü��
    If PassGetState("PassEnable") = 0 Then
        MsgBox "��ǰ������ҩ���ϵͳ�����ã�������������Ƿ���ȷ��", vbInformation, gstrSysName
        Call PassQuit: Exit Function
    End If
    
    'PASSӦ��ģʽ����(��Ĭ��ֵ)
    Call PassSetControlParam(1, 2, 0, 2, 1)
    
    PassInitialize = True
    Exit Function
errH:
    If Err.Number = 53 And InStr(UCase(Err.Description), UCase("ShellRunAs.dll")) > 0 Then
        MsgBox "PASS�ӿ��ļ� ShellRunAs.dll ������,���ܺ�����ҩ���ϵͳδ��ȷ��װ�����á�" & _
            vbCrLf & "����ȷ��װ�����ú�����ҩ���ϵͳ֮ǰ����Ӧ�Ĺ��ܲ���ʹ�á�", vbInformation, gstrSysName
    ElseIf Err.Number = 53 And InStr(UCase(Err.Description), UCase("DIFPassDll.dll")) > 0 Then
        MsgBox "PASS�ӿ��ļ� DIFPassDll.dll ������,��������Ϊ����ԭ��" & vbCrLf & _
            vbCrLf & "1.PASS�ͻ����ǵ�һ�ε�¼�����˳�֮�������µ�¼��������ʹ�á�" & _
            vbCrLf & "2.������ҩ���ϵͳδ��ȷ��װ�����ã�����ϸ�����ٵ�¼���ԡ�", vbInformation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function
