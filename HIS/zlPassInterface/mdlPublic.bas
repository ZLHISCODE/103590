Attribute VB_Name = "mdlPublic"
Option Explicit

'����---------------------
Public Const G_STR_PASS As String = "������ҩ����"
Public Const G_STR_MATCH As String = "abcdefghigklmnopkrstuvwxyzABCDEFGHIGKLMNOPKRSTUVWXYZ0123456789"" </>_="
Public Const G_INT_MODEL_0 As Integer = 0
Public Const G_INT_MODEL_1 As Integer = 1
Public Const G_STR_SPLIT As String = "&&"
Public Const SW_SHOWNORMAL = 1
'API����
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'ȫ�ֱ���-------------------------------
Public gfrmMain As Object                   '������
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gobjComLib As Object                    '������������ZL9ComLib
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��
Public gstrSysName As String                'ϵͳ����
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngSys As Long
Public gbytUseType As Byte                  '0-ҽ���´�
                                            '1-�ٴ�·����Ŀ��ҽ������
                                            '2-�ٴ�·�����·������Ŀ��ҽ����������ѡ����
                                            '3-ҽ��˳�����(������ʾ��ֹͣ��ҽ������Ϊ�ƶ�ʱ������Щҽ�������Ҫһ�����)
Public glngObject As Long                   '��¼�������
Public gobjPlugIn   As Object
Public gsngWaitTime   As Single               '���ʵȴ����������
Public gsngAutoLinkTime As Single              'ÿ��5���Ӽ������
Public gblnBreak     As Boolean             'T-�Ͽ�����;F-������
Public gsngCheckLinkTime As Single            '
Public mstrLike As String                   '����ƥ�䷽ʽ
Public mint���� As Integer                  '����ƥ�䷽ʽ��0-ƴ��,1-���
Public gstrMatchMode As String '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
'------------------------------------------------------------------
'������ҩ������ò���
'------------------------------------------------------------------

Public gbytPass As Byte             'ZLHIS��ʹ��PASS�ӿ�����,0-δʹ��,1-����,2-��ͨ,3-̫Ԫͨ,4-ҩ��ʿ,5-��������,6-������Ϣ
Public gbytBlackLamp As Byte        '�Ƿ��������ҩƷ
Public gbytReason As Byte           '����ҩƷҪ����дԭ��
Public gbytSuperVolume As Byte      '�Ƿ��ֹ������ҩƷ
Public gbytOutBlackLamp As Byte     '�Ƿ�����Ժ��ִ�еĽ���ҩƷҽ��
Public gobjPass As Object           '3-̫Ԫͨ�ӿڶ���,4-ҩ��ʿ
Public gbytOpenLog As Byte          '������ͨ�ӿڵ�����־ 0-�����ã�1-����
Public gbytSysSet As Byte           '��������ʹ��ϵͳ���� 1-��ʾ��0-����
Public gstrVersion As String        '��ʶ�ӿڰ汾��
Public gblnPharmReview As Boolean   '��������ҩʦ�󷽸�Ԥϵͳ
Public gblnPrePregnancy As Boolean  '����ѡ������ Preparation of pregnancy
Public gblnTEST As Boolean          '���þ�Ĭʽ��� T-����MDC_DoCheck(0,1)����������������ⲻ���������ֻ�����������ݵĲɼ�����ʹ��Ĭʽ��������ҩ���⣬Ҳ������ҽ����Ϊ�����أ��Ա���ȫԺ��ʼʵʩ�׶δ����û����ݹ�����ϴ���ã�������Ч��Ϣ��ҽ��ҵ��ĸ��š�
'---------------
Public gstrIP           As String           '������IP
Public gstrPort         As String           '�������˿ں�
Public gstrDrugIP       As String           'ҩƷ˵����IP
Public gstrDrugPort     As String           'ҩƷ˵����˿ں�
Public gstrUser         As String           '�û���
Public gstrPWD          As String           '�û�����
Public gstrPortPlus     As String           '�������˿ں�
Public gstrHOSCODE      As String           'ҽԺ����
Public gstrStatusEdit   As String           '�༭����״̬
Public gstrStatusGet    As String           '��ȡ����״̬   http://192.168.0.231:8080/ords/patstatus/pat/getpatstatus
Public gstrStatusSave As String           '���没��״̬   http://192.168.0.231:8080/ords/patstatus/pat/saverecord

Public gbytType         As Byte             '��������:0-����;1-�ǹ���

Public gint�����Ǽ���Ч���� As Integer
Public gblnInitOK As Boolean         '���ڱ�ǳ�ʼ��ִ��״̬ 'T-ִ�й���ʼ��;F-δִ�й���ʼ��
Public gblnPassOK   As Boolean         'T-��ʼ���ɹ�,�������;F=��ʼ��ʧ�� ��ֹPass�ӿڵ���

Public glngPatiID As Long              '��¼��ǰ����ID
Public glng��ҳID As Long
Public gblnTip As Boolean             '����4.0���ڱ����ظ���ӿ��д�����ͬ��ҩƷ��Ϣ
Public gbytOpen As Byte   '�����������

'��¼�û��ṹ
Public Type TYPE_USER_INFO
    id As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    ����ID As Long
    ������ As String
    ������ As String
    רҵ����ְ�� As String
    רҵ�������� As String
    ��ҩ���� As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum DataEnum
    responseText = 1
    responseBody = 2
End Enum

Public UserInfo As TYPE_USER_INFO

Public gobjCOL As clsVSCOL           '��ǰҽ����ӳ��
Public gobjAdvice As Object         '��ǰҽ���б���� vsAdvice
Public gobjCmdAlley As Object           '��ǰPASS����ʷ��ť

Public glngModel As Long                '��ǰ����gbytModel 0-����༭,1-סԺ�༭��2-סԺҽ���嵥,3-��ʿУ��,4-����ҽ���嵥
Public gobjDiags As clsDiags              '����
Public gint���� As Integer              ' ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
Public gcolPASSExe As Collection        '�˵�����ӳ��
Public gcolPASSState As Collection      '�����˵�״ֵ̬ӳ��


Public gobjMap As clsPassMap  'ӳ�����
Public gobjPati As clsPatient

Public gblnOpen As Boolean    '��Ҫ��Ϣ�Ƿ��
Public glngDrugID As Long    '��¼����һ�δ��˵�ҩƷID

'�������ܺ�
Public Enum G_PASS_MK
    MK_���PASS�˵�״̬ = 0
    MK_סԺ������� = 1
    MK_סԺ�ύ��� = 2
    MK_�ֹ�������� = 3
    MK_��ҩ���� = 6
    MK_ϵͳ���� = 11
    MK_��ҩ�о� = 12
    MK_ҩƷ�����Ϣ = 13
    MK_��ҩ;�������Ϣ = 14
    MK_����״̬����ʷ�鿴 = 21
    MK_����״̬����ʷ = 22
    MK_���ﱣ����� = 33
    MK_ҩ���ٴ���Ϣ�ο� = 101
    MK_ҩƷ˵���� = 102
    MK_������ҩ���� = 103
    MK_����ֵ = 104
    MK_ҽԺҩƷ��Ϣ = 105
    MK_ҽҩ��Ϣ���� = 106
    MK_�й�ҩ�� = 107
    MK_ҩ��_ҩ���໥���� = 201
    MK_ҩ��_ʳ���໥ʹ�� = 202
    MK_����ע������� = 203
    MK_����ע������� = 204
    MK_����֢ = 205
    MK_������ = 206
    MK_��������ҩ = 207
    MK_��ͯ��ҩ = 208
    MK_��������ҩ = 209
    MK_��������ҩ = 210
    MK_�رո������� = 402 '�رյ�ǰ���и�������
    MK_��ʾ��ʾ���� = 403  '��ʾ��ʾ��ʾ����
End Enum

Public Enum G_PASS_MK4
    MK4_���PASS�˵�״̬ = 0
    MK4_���
    MK4_�Զ����
    MK4_ҩƷ˵���� = 11
    MK4_ҩ��ר�� = 21
    MK4_������ҩ���� = 31
    MK4_�й�ҩ�� = 41
    MK4_ҩƷ��Ҫ��Ϣ = 51
    MK4_ҩ���໥���� = 61
    MK4_ҩʳ�໥���� = 62
    MK4_�������� = 63
    MK4_����Ũ�� = 64
    MK4_ҩ�����֢ = 65
    MK4_ҩ����Ӧ֢ = 66
    MK4_������Ӧ = 67
    MK4_���𺦼��� = 68
    MK4_���𺦼��� = 69
    MK4_��ͯ��ҩ = 70
    MK4_������ҩ = 71
    MK4_������ҩ = 72
    MK4_������ҩ = 73
    MK4_������ҩ = 74
    MK4_�Ա���ҩ = 75
    MK4_ϸ����ҩ�� = 76
End Enum

'������3.0�˵�����ֵ
Public Enum G_MK_INDEX
    MK_IX_ҩ���ٴ���Ϣ�ο� = 0
    MK_IX_ҩƷ˵���� = 1
    MK_IX_�й�ҩ��
    MK_IX_������ҩ����
    MK_IX_����ֵ
    MK_IX_ר����Ϣ
    MK_IX_ҩ���໥����
    MK_IX_ҩʳ�໥����
    MK_IX_����ע�������
    MK_IX_����ע�������
    MK_IX_����֢
    MK_IX_������
    MK_IX_��������ҩ
    MK_IX_��ͯ��ҩ
    MK_IX_��������ҩ
    MK_IX_��������ҩ
    MK_IX_ҽҩ��Ϣ����
    MK_IX_ҩƷ�����Ϣ
    MK_IX_��ҩ;�������Ϣ
    MK_IX_ҽԺҩƷ��Ϣ
    MK_IX_ϵͳ����
    MK_IX_��ҩ�о�
    MK_IX_����
    MK_IX_���
End Enum
'������4.0�˵�����ֵ
Public Enum G_MK4_INDEX
    MK4_IX_��� = 0
End Enum

'̫Ԫͨ ���ܺ�
Public Enum G_PASS_TYT
    TYT_��ҩ�淶 = 0
    TYT_ҩ����� = 1
    TYT_ҩƷ��ʾ = 2
    TYT_ҽҩ֪ʶ�� = 3
    TYT_ϵͳ���� = 4
    TYT_������� = 5
End Enum

'�������� ���ܺ�
Public Enum G_PASS_HZYY
    HZYY_ҩƷ˵���� = 0
    HZYY_ҩ����� = 1
End Enum
'������Ϣ ���ܺ�
Public Enum G_PASS_ZL
    ZL_ҩ����� = 0
    ZL_����״̬
End Enum

Public Enum G_PASS_UseStation
    US_InDoctor = 0     'סԺҽ��վ
    US_InNurse = 1      'סԺ��ʿվ
    US_Intech = 2       'סԺҽ��վ
End Enum

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    PҩƷ������ҩ = 1341        '1341    ҩƷ������ҩ
    PҩƷ���ŷ�ҩ = 1342        '1342    ҩƷ���ŷ�ҩ
    PPIVA���� = 1345        '1345    PIVA����
End Enum

Public Enum G_TYPE_FUN
    FUN_ҽ����Ϣ = 1
    FUN_������� = 3
    FUN_ҽ����Ϣ_DTBS = 4
    FUN_����� = 5
    FUN_ҽ����Ϣ_HZYY = 6
    FUN_�����_HZYY = 7
    FUN_ҽ����Ϣ_ZL = 8
    FUN_�����_ZL = 9
    FUN_�����_YWS = 10
    FUN_��������_ZL = 11
    FUN_ҩʦ���_ZL = 12
    FUN_����״̬_ZL = 13
End Enum

Public Enum G_TYPE_FLOATWIN
    FLOATWIN_CLOSE = 0   '�ر�
    FLOATWIN_DRUG = 1    'ҩƷ��Ϣ��ʾ��
    FLOATWIN_WARN = 2    '��ʾ����
End Enum

'������ָ������Ļ�����ϵ�λ��
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'��ô�������Ļ�����е�λ��
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'�ж�ָ���ĵ��Ƿ���ָ���ľ����ڲ�
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long
'׼������ʹ����ʼ������ǰ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'�����ƶ�����
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'��ȡ����״̬
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'HWND hwnd, // ָ���ֲ㴰�ھ��
'COLORREF crKey, // ָ����Ҫ͸���ı�����ɫֵ������RGB()��
'BYTE bAlpha, // ����͸���ȣ�0��ʾ��ȫ͸����255��ʾ��͸��
'DWORD dwFlags // ͸����ʽ
'       ���У�dwFlags������ȡ����ֵ��
'       LWA_ALPHA=&H2ʱ��crKey������Ч��bAlpha������Ч��
'       LWA_COLORKEY=&H1�������е�������ɫΪcrKey�ĵط�����Ϊ͸����bAlpha������Ч���䳣��ֵΪ1��
'       LWA_ALPHA | LWA_COLORKEY��crKey�ĵط�����Ϊȫ͸�����������ط�����bAlpha����ȷ��͸���ȡ�
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = -4&
Public Const WM_MOUSEWHEEL = &H20A
 
Public glngOldWindowProc As Long '��������ϵͳĬ�ϵĴ�����Ϣ�������ĵ�ַ

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Const SWP_NOACTIVATE = &H10 '�������
Public Const GWL_EXSTYLE  As Long = (-20)
Public Const WS_EX_TOPMOST As Long = &H8
Public Const HWND_TOPMOST As Long = -1
Public Const SW_SHOWMAXIMIZED = 3
'API:GetSystemMetrics
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
'���أ���-����ɹ�
    Dim strPara As String
    Dim arrTemp As Variant
    
    gbytPass = Val(zlDatabase.GetPara(30, glngSys))  '�ӿ�����
    If gbytPass = UNPASS Then Exit Function
    gstrVersion = zlDatabase.GetPara(228, glngSys) '��ʶ�ӿڰ汾��
    '��ʼ�ɹ��������ظ���ȡ����ֵ��gbytPass����ģ����û�Ȩ�޽��õ�ԭ�����Ϊ:0-UNPASS����ÿ����Ҫ���¶�ȡ��
    If gbytPass = MK Or gbytPass = YWS Then
        gbytSysSet = Val(zlDatabase.GetPara(226, glngSys))
        If gbytPass = MK And gstrVersion = "4.0" Then
            If Not MK_GetPara Then
                MsgBox "������ҩ��������������,�뵽:" & vbCrLf & _
                    "���ٴ��������á�->��ҵ�����̿��ơ�->��������ҩ�ӿڡ�->�����á������á�" & vbCrLf & _
                    "����ȷ����֮ǰ����Ӧ�Ĺ��ܲ���ʹ�á�", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
    ElseIf gbytPass = DT Then  '��ͨ
        gbytOpenLog = Val(zlDatabase.GetPara(225, glngSys))
        If gstrVersion = "4.0" Then
            gstrHOSCODE = zlDatabase.GetPara(90001, glngSys, , "1513")
        End If
    ElseIf gbytPass = HZYY Then '��������
        Call HZYY_GetPara
    ElseIf gbytPass = ZL Then
        If Not ZL_GetPara Then
            MsgBox "������ҩ��������������,�뵽:" & vbCrLf & _
                "���ٴ��������á�->��ҵ�����̿��ơ�->��������ҩ�ӿڡ�->�����á������á�" & vbCrLf & _
                "����ȷ����֮ǰ����Ӧ�Ĺ��ܲ���ʹ�á�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    gbytBlackLamp = Val(zlDatabase.GetPara(161, glngSys))  '�Ƿ��������ҩƷ
    gbytReason = Val(zlDatabase.GetPara(249, glngSys)) '249    ����ҩƷҪ����дԭ��
    gbytSuperVolume = Val(zlDatabase.GetPara(182, glngSys)) '�Ƿ��ֹ������ҩƷ
    
    gbytOutBlackLamp = Val(zlDatabase.GetPara(189, glngSys)) '�Ƿ�����Ժ��ִ�еĽ���ҩƷҽ��
    
    'Ƥ�Խ����Чʱ��
    gint�����Ǽ���Ч���� = Val(zlDatabase.GetPara(70, glngSys))

    InitSysPar = True
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
'���ܣ���ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    If str���� <> "" Then
        strSQL = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str����)
    Else
        strSQL = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.id)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���˹�����¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal bytFunc As Byte = 1) As ADODB.Recordset
'���ܣ���ȡ���˹�����¼
'������bytFunc=1 ���벡���������й�����¼;0=���벡�˱��ξ��������¼
    Dim strSQL As String
    
    On Error GoTo errH
    If bytFunc = 0 Then
        If lng��ҳID = 0 Then
            strSQL = "Select Distinct ҩ��ID,ҩ����,����Դ����,������Ӧ,��¼ʱ�� From ���˹�����¼ Where ����ID=[1] And ���=1 And Nvl(����ʱ��,��¼ʱ��)>Trunc(Sysdate-[3])"
        Else
            strSQL = "Select Distinct ҩ��ID,ҩ����,����Դ����,������Ӧ,��¼ʱ�� From ���˹�����¼ Where ����ID=[1] And ��ҳID=[2] And ���=1"
        End If
    Else
        strSQL = "Select Distinct ҩ��ID,ҩ����,����Դ����,������Ӧ,��¼ʱ�� From ���˹�����¼ Where ����ID=[1] And ���=1"
    End If
    Set Get���˹�����¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng��ҳID, gint�����Ǽ���Ч����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get������ϼ�¼(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str���� As String) As ADODB.Recordset
'���ܣ���ȡ������ϼ�¼
'������lng����ID�����ﲡ�˴��Һ�ID��סԺ���˴���ҳID
'       �������-1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;
'        11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���
'       ��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����(����ҽ��վ,���ժҪ);
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select a.ID,a.����id, a.���id, a.�������, a.��ϴ���, Nvl(b.����, c.����) As ����, NVL(Nvl(b.����, c.����),a.�������) ����" & vbNewLine & _
             ",a.��¼����,a.��¼�� " & vbNewLine & _
             "From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C" & vbNewLine & _
             "Where a.����id = [1] And a.��ҳid = [2] And ȡ��ʱ�� Is Null And ��¼��Դ IN (1, 3) And Instr(',' ||[3]|| ',', ',' || ������� || ',') > 0 And a.����id = b.Id(+) And" & vbNewLine & _
             "      a.���id = c.Id(+)" & vbNewLine & _
             "Order By ��¼��Դ, �������, ��ϴ���"
    Set Get������ϼ�¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng����ID, str����)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���˲��������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ����ݲ���ID����ҳID��ȡ���˲��������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH

    If lng��ҳID = 0 Then
        lng��ҳID = Val(zlDatabase.GetPara(21, glngSys))
        strSQL = "Select ���������" & vbNewLine & _
                 "From ���˹Һż�¼" & vbNewLine & _
                 "Where ����id = [1] And �Ǽ�ʱ�� > Trunc(Sysdate-[2]) And ��������� Is Not Null And Rownum = 1"
    Else
        strSQL = "Select ��Ϣֵ As ���������" & vbNewLine & _
                 "From ������ҳ�ӱ� Where ����id = [1] And ��ҳid = [2] And ��Ϣ�� = '���������'"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        While Not rsTmp.EOF
            Get���˲�������� = Get���˲�������� & "," & rsTmp!���������
            rsTmp.MoveNext
         Wend
        Get���˲�������� = Mid(Get���˲��������, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get���������¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˹�����¼
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ��������ID,��������,������ʼʱ��,��������ʱ�� From ���������¼ Where ����ID=[1] And ��ҳID=[2] "

    Set Get���������¼ = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetPatiOperation(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal str�Һŵ� As String) As ADODB.Recordset
'���ܣ���ȡ���˹�����¼
    Dim strSQL As String
    
    On Error GoTo errH
    If str�Һŵ� = "" Then
        strSQL = " And a.����id = [1] And a.��ҳid = [2] "
    Else
        strSQL = "  And a.�Һŵ� = [3] "
    End If
    strSQL = "Select a.Id, a.����ʱ��, c.����, c.����" & vbNewLine & _
               "From ����ҽ����¼ A, ������϶��� B, ��������Ŀ¼ C" & vbNewLine & _
               "Where a.������Ŀid = b.����id And b.����id = c.Id And a.������� = 'F' And a.ҽ��״̬  In (1,2,3,5,8) " & strSQL
    Set GetPatiOperation = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng��ҳID, str�Һŵ�)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiSymptom(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ����ݲ���ID����ҳID��ȡ����֢״��̫Ԫͨ�ӿ�ʹ�ã�
'lng��ҳId :���ﴫ�Һ�ID
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select a.����,a.���� From ����֢״��¼ a " & vbNewLine & _
            "Where a.����ID=[1] And a.��ҳID=[2] "
    Set GetPatiSymptom = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ����ݲ���ID����ҳID��ȡ���˻�����Ϣ
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select A.סԺ��, A.��ǰ����, A.��������, Nvl(B.����, A.����) ����, Nvl(B.�Ա�, A.�Ա�) �Ա�, Nvl(B.����, A.����) ����, A.�����, A.������,A.���֤��,B.���,B.����" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B" & vbNewLine & _
            "Where A.����id = B.����id And A.����id = [1] And B.��ҳid = [2]"

    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetƵ����Ϣ_����(ByVal strƵ�� As String, intƵ�ʴ��� As Integer, _
    intƵ�ʼ�� As Integer, str�����λ As String, str��Χ As String, Optional strƵ�ʱ��� As String) As Boolean
'���ܣ�����Ƶ�ʵ������Ϣ
'������strƵ��=Ƶ������
'      str��Χ=1-��ҽ,2-��ҽ,-1-һ����,-2-������
'���أ���������ȡ��ʱ������True�����򷵻�False
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    
    On Error GoTo errH
    
    intƵ�ʴ��� = 0
    intƵ�ʼ�� = 0
    str�����λ = ""
    
    strSQL = "Select Ƶ�ʴ���,Ƶ�ʼ��,�����λ,���� From ����Ƶ����Ŀ Where ����=[1] And Instr([2],','||���÷�Χ||',')>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strƵ��, "," & str��Χ & ",")
    If Not rsTmp.EOF Then
        intƵ�ʴ��� = NVL(rsTmp!Ƶ�ʴ���, 0)
        intƵ�ʼ�� = NVL(rsTmp!Ƶ�ʼ��, 0)
        str�����λ = NVL(rsTmp!�����λ)
        strƵ�ʱ��� = "" & rsTmp!����
        GetƵ����Ϣ_���� = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDoctorTitleType(ByVal strDoctTitle As String) As String
'���ܣ�����ҽ��ְ�Ʒ���ְ�����
'����ֵ��
'C --�����ڣ����ڣ�������ҽʦ������ҽʦ��ר��
'B������ҽʦ����ʦ
'A�������ϵ�����ְ��

    If InStr(";������;����;������ҽʦ;����ҽʦ;ר��;", ";" & strDoctTitle & ";") > 0 Then
        GetDoctorTitleType = "C"
    ElseIf InStr(";����ҽʦ;��ʦ;", ";" & strDoctTitle & ";") > 0 Then
        GetDoctorTitleType = "B"
    Else
        GetDoctorTitleType = "A"
    End If

End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.id = rsTmp!id
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = NVL(rsTmp!����)
            UserInfo.���� = NVL(rsTmp!����)
            UserInfo.����ID = NVL(rsTmp!����ID, 0)
            UserInfo.������ = NVL(rsTmp!������)
            UserInfo.������ = NVL(rsTmp!������)
            UserInfo.���� = Get��Ա����
            UserInfo.רҵ����ְ�� = NVL(rsTmp!רҵ����ְ��)
            UserInfo.רҵ�������� = Sys.RowValue("רҵ����ְ��", UserInfo.רҵ����ְ��, "����", "����")
            GetUserInfo = True
        End If
    End If
    gstrDBUser = UserInfo.�û���
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function PassCheckPrivs(ByVal lngModel As Long, Optional ByVal blnInit As Byte = False) As Boolean
'����:����ģ��Ż�ȡģ����е�Ȩ��
'����:blnInit -�Ƿ��ʼ��(סԺҽ��վ��ʼ��ʱ��Ҫ�ж�סԺҽ���´��סԺҽ�����͵ĺ�����ҩ���Ȩ��)
    Dim blnDo As Boolean
    WriteLog "clsPass", "PassCheckPrivs", "PassCheckPrivs_Begin"
    Select Case lngModel
    
    Case PM_����༭, PM_����ҽ���嵥
        If InStr(GetInsidePrivs(p����ҽ���´�), "������ҩ���") > 0 Then blnDo = True
    Case PM_סԺҽ���嵥
        If blnInit Then
            If InStr(GetInsidePrivs(pסԺҽ���´�) & GetInsidePrivs(pסԺҽ������), "������ҩ���") > 0 Then blnDo = True
        Else
            If InStr(GetInsidePrivs(pסԺҽ���´�), "������ҩ���") > 0 Then blnDo = True
        End If
    Case PM_סԺ�༭
        If InStr(GetInsidePrivs(pסԺҽ���´�), "������ҩ���") > 0 Then blnDo = True
    Case PM_��ʿУ��
        If InStr(GetInsidePrivs(pסԺҽ������), "������ҩ���") > 0 Then blnDo = True
    Case PM_סԺ��ҳ
        blnDo = True
    Case PM_������ҩ, PM_���ŷ�ҩ, PM_PIVA����
        If InStr(GetInsidePrivs(lngModel), "������ҩ���") > 0 Then blnDo = True
    End Select
    
    PassCheckPrivs = blnDo
    WriteLog "clsPass", "PassCheckPrivs", "PassCheckPrivs_End"
End Function

Public Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵��:��frmDockInAdviceһ����ҩ����һ��
    Dim i As Long, blnTmp As Boolean
    With gobjAdvice
        If .TextMatrix(lngRow, gobjCOL.intCOL�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, gobjCOL.intCOL�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, gobjCOL.intCOL���ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, gobjCOL.intCOL���ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, gobjCOL.intCOL���ID)) = Val(.TextMatrix(lngRow, gobjCOL.intCOL���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Public Function InitAdviceRS(Optional ByVal bytFunc As Byte = 1) As ADODB.Recordset
'����:����ҽ����¼
    Dim rs As ADODB.Recordset
    Dim strFields As String
    Dim strFieldName As String
    Dim lngLen As Long
    Dim FieldType As DataTypeEnum
    Dim i As Long, j As Long
    
    Dim arrField As Variant
    Dim arrSubFeld As Variant '�ֶ�����|�ֶ�����|�ֶγ��� ȱʡ�ֶ����� ΪadVarChar
    Select Case bytFunc
    
    Case FUN_ҽ����Ϣ
        strFields = "ҽ��ID||18,���ID||18,ҽ����Ч||1,ҽ�����||5,ҽ��״̬||3,�������||3,��������||100,��������ID||18,����ҽ������||10,����ҽ��||100," & _
        "ҩƷID||18,ҩƷ����||100,��������||16,������λ||20,Ƶ��||50,�÷�||100,�÷�ID||18,����ʱ��||20,��ʼʱ��||20,����ʱ��||20,����||16,������λ||20," & _
        "��ҩĿ��||1,ҽ������||100,��ʾ|adInteger|1,����||100,���״̬|adInteger|1,������||30,�������||18,ִ�п���ID||18,����||16"  '���﷢��
    Case FUN_�������
        strFields = "ҽ��ID|adBigInt|18,ҩƷ����||1000,�Ƿ����|adInteger|1,����ҩƷ˵��||100,״̬|adInteger|1"
    Case FUN_ҽ����Ϣ_DTBS
        strFields = "ҽ��ID||18,���ID||18,ҽ����Ч||1,ҽ�����||5,ҽ��״̬||3,�������||3,��������||100,��������ID||18,����ҽ������||10,����ҽ��||100," & _
        "������ĿID||18,ҩƷID||18,ҩƷ����||100,��������||16,������λ||20,Ƶ��||50,�÷�||100,�÷�ID||18,����ʱ��||20,��ʼʱ��||20,����ʱ��||20,����||16,������λ||20," & _
        "��ҩĿ��||1,ҽ������||100,��ʾ|adInteger|1,����||16,���||100,Ƶ�ʱ���||5,��ҩ����||1000,��־||1,��Ժ��ҩ|adInteger|1"
    Case FUN_ҽ����Ϣ_HZYY
        strFields = "ҽ��ID||18,���ID||18,ҽ����Ч||1,ҽ�����||5,ҽ��״̬||3,�������||3,��������||100,��������ID||18,����ҽ��ID||10,����ҽ��||100," & _
        "������ĿID||18,ҩƷID||18,ҩƷ����||100,��������||16,������λ||20,Ƶ��||50,��ҩ�巨||100,��ҩ�巨ID||18,�÷�||100,�÷�ID||18,����ʱ��||20,��ʼʱ��||20,����ʱ��||20,����||16,������λ||20," & _
        "��ҩĿ��||50,ҽ������||100,��ʾ|adInteger|1,����||16,���||100,Ƶ�ʱ���||5,��ҩ����||1000,��־||1,��Ժ��ҩ|adInteger|1,����ID|adBigInt|18,����||100," & _
        "רҵ����ְ��||50,��Һ||3"
    Case FUN_�����
        strFields = "��ʾֵ||3,ҽ��ID||18"
    Case FUN_�����_HZYY
        strFields = "DrugName||100,DrugID||18,advice||1000,source||100,GroupNo||18,Type||200,Message||1000,Severity|adInteger|2,recipeId||18"
    Case FUN_ҽ����Ϣ_ZL
        strFields = "ҽ��ID||18,������ĿID||18,ҩƷID||18,��λ��||50,��Һ���||18,������λ||20,������||20," & _
        "ÿ����||20,��ҩƵ��||50,��ҩƵ������||50,��ҩ;��||18,����|adInteger|2,��ҩ;������||100,ҽ����Ч||1,����ʱ��||20,������־||3," & _
        "����ҽ��||100,ҽ��ְ��||100,ҽ������ҩ��ȼ�||50,ҽ������||1000,��ҩ����||16,����||50,ҩƷ����ҩ��ȼ�||50," & _
        "�������||50,����˵��||500,��ҩĿ��||50,ҩƷ���ɵȼ�||50,ҩƷ��������||50,ҩƷ����˵��||500,�������||300,ҽ��״̬||3"
    Case FUN_�����_ZL
        strFields = "OrderId||18,Type||100,Level||50,DrugCode||50,Describ||2000,Remaks||4000,Light|adInteger|1,Tag|adInteger|2,WarnLevel|adInteger|1,Category|adInteger|1 "
    Case FUN_�����_YWS
        strFields = "Title||200,Detail||1000"
    Case FUN_��������_ZL
        strFields = "Name||500,Type||100,Index||200,Value||1000,Class||200,Obsid||50,Default||200,ControlIndex|adInteger|3,Proid||18"
    Case FUN_ҩʦ���_ZL
        strFields = "ҽ��ID|adBigInt|18,���ID|adBigInt|18,ҽ������||1000,�������||1000,����||1000,Tag|adInteger|1"
    Case FUN_����״̬_ZL
        strFields = "STATUS_ID||50,STATUS_NAME||100,STATUS_SITUATION||5"
    End Select
    
    Set rs = New ADODB.Recordset
    '-----------------------------------------
    With rs.Fields
        arrField = Split(strFields, ",")
        For i = LBound(arrField) To UBound(arrField)
            arrSubFeld = Split(arrField(i), "|")
            strFieldName = arrSubFeld(0)
            If UCase(arrSubFeld(1) & "") = UCase("adVarChar") Then
                FieldType = adVarChar
            ElseIf UCase(arrSubFeld(1) & "") = UCase("adBigInt") Then
                FieldType = adBigInt
            ElseIf UCase(arrSubFeld(1) & "") = UCase("adInteger") Then
                FieldType = adInteger
            Else
                FieldType = adVarChar
            End If
            lngLen = Val(arrSubFeld(2))
            .Append strFieldName, FieldType, lngLen
        Next
    End With
    '---------------------------------------
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    '----------------------------------
    Set InitAdviceRS = rs
End Function

Public Function GetDrugID(ByVal str������ĿID As String) As Variant
'����:����ҩƷID���¼
    Dim strSQL As String
    Dim arrTmp As Variant
    Dim rs As ADODB.Recordset
    
    On Error GoTo errH
    arrTmp = Split(str������ĿID, ",")
    If UBound(arrTmp) = 0 Then
        strSQL = "Select ҩ��ID,ҩƷID from ҩƷ��� where ҩ��id=[1] and rownum <2"
    ElseIf UBound(arrTmp) > 0 Then
        strSQL = "Select a.ҩ��id, Max(a.ҩƷid) As ҩƷid" & vbNewLine & _
        "From ҩƷ��� A" & vbNewLine & _
        "Where a.ҩ��id In (Select * From Table(f_Num2list([1])))" & vbNewLine & _
        "Group By a.ҩ��id"
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlPass", str������ĿID)
    
    If UBound(arrTmp) = 0 Then
        If Not rs.EOF Then
            GetDrugID = rs!ҩƷID & ""
        Else
            GetDrugID = ""
        End If
    Else
        Set GetDrugID = rs
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��ҩ�䷽(ByVal str��IDs As String) As ADODB.Recordset
'����:����ҩƷID���¼
    Dim strSQL As String
    Dim arrTmp As Variant
    Dim rs As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = "Select a.Id, a.���id, a.ҽ����Ч, a.ҽ��״̬,a.�������,a.������Ŀid,a.�շ�ϸĿID as ҩƷID,a.ҽ������ As ҩƷ����,a.���,a.��������, d.���㵥λ As ������λ,a.ִ��Ƶ�� as Ƶ��, a.�����λ, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.��ʼִ��ʱ�� As ��ʼʱ��," & vbNewLine & _
            "       a.ִ����ֹʱ�� As ��ֹʱ��, a.����ʱ��, a.ͣ��ʱ��, c.���� As �÷�, c.Id As �÷�id, a.ִ������, b.ִ������ As ��ִ������,a.����ҽ��,a.��ҩĿ��,a.����, " & vbNewLine & _
            "       a.�ܸ�����,f.סԺ��λ As ������λ,f.���ﵥλ, a.ҽ������, a.��������id,a.ִ�п���ID " & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ����¼ B, ������ĿĿ¼ C, ������ĿĿ¼ D, ҩƷ��� F" & vbNewLine & _
            "Where a.���id = b.Id And b.������Ŀid = c.Id And a.������Ŀid = d.Id And a.�շ�ϸĿid = f.ҩƷid(+) And" & vbNewLine & _
            "      a.���id in (Select * From Table(f_Num2list([1]))) And a.������� = '7'"


    Set rs = zlDatabase.OpenSQLRecord(strSQL, "��ҩ��ID", str��IDs)

    Set Get��ҩ�䷽ = rs

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����(ByVal str��IDs As String) As ADODB.Recordset
'����:����ҩƷID���¼
    Dim strSQL As String
    Dim arrTmp As Variant
    Dim rs As ADODB.Recordset
    
    On Error GoTo errH

    strSQL = "Select a.Id, a.ҽ������" & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B " & vbNewLine & _
            "Where A.������ĿID = B.ID And A.������� ='E' And B.�������� = '2' And b.ִ�з��� = 1 And NVL(a.ҽ������,'��') <> '��' And " & vbNewLine & _
            "      a.ID in (Select /*+cardinality(A,10)*/ * From Table(f_Num2list([1])) A) "


    Set rs = zlDatabase.OpenSQLRecord(strSQL, "ҽ������", str��IDs)

    Set Get���� = rs

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRS(ByVal strTableName As String, ByVal strFileds As String, ByVal strInput As String, _
        Optional ByVal strWhere As String = "ID", Optional ByVal bytModel As Byte = 0, Optional ByVal bytType As Byte = 0) As Variant
'����:����ָ����ָ���ֶεļ�¼��
'������strTableName-����
'     strFileds
'     strInput ��ʽ1(1����������)��ID1,ID2,...
'              ��ʽ2(2����������)������1,��Χ1;����2,��Χ2;...
'             strSQL = "Select ����, ����, ���÷�Χ" & vbNewLine & _
'                "From ����Ƶ����Ŀ" & vbNewLine & _
'                "Where (����, ���÷�Χ) In (Select /*+cardinality(B,10)*/" & vbNewLine & _
'                "                      C1, C2" & vbNewLine & _
'                "                     From Table(f_Str2list2('ÿ�����,1|ÿ������,1', ';', ',')) B)"
'    bytModel=1 ��������Ϊ����
'    ��bytModel=1ʱ�� bytType=0-����� C1,C2 ͬΪ�ַ��� =1-C1(Number),C2(Number);=2-C1(char),C2(Number);=3-C1(Number),C2(Char)
'    ��bytModel=0ʱ�� bytType=0-f_num2list; bytType=1 f_Str2list


    Dim strSQL As String
    Dim strSub As String
    Dim strFun As String
    Dim arrTmp As Variant
    
    On Error GoTo errH
    
    If bytModel = 1 Then
        If bytType = 0 Then
            strSub = " C1,C2 "
            strFun = "f_Str2list2"
        ElseIf bytType = 1 Then
            strSub = " C1,C2 "
            strFun = "f_num2list2"
        ElseIf bytType = 2 Then
            strSub = "C1,To_Number(C2) As C2 "
            strFun = "f_Str2list2"
        ElseIf bytType = 3 Then
            strSub = " To_Number(C1) As C1,C2 "
            strFun = "f_Str2list2"
        End If
        strSQL = " Select  " & strFileds & vbNewLine & _
                " From  " & strTableName & vbNewLine & _
                " Where (" & strWhere & ") In (Select /*+cardinality(B,10)*/" & vbNewLine & _
                "                    " & strSub & vbNewLine & _
                "                     From Table(" & strFun & "([1], ';', ',')) B)"
    Else
        If bytType = 0 Then
            strFun = "f_num2list"
        ElseIf bytType = 1 Then
            strFun = "f_Str2list"
        End If
        arrTmp = Split(strInput, ",")
        If UBound(arrTmp) = 0 Or strInput = "" Then
            strSQL = "Select " & strFileds & "  From " & strTableName & " Where " & strWhere & " = [1]"
        ElseIf UBound(arrTmp) > 0 Then
            strSQL = "Select " & strFileds & vbNewLine & _
            "From " & strTableName & vbNewLine & _
            "Where " & strWhere & " In (Select /*+cardinality(A,10)*/ * From Table(" & strFun & "([1]))A )"
        End If
    End If
    Set GetRS = zlDatabase.OpenSQLRecord(strSQL, "mdlPass", strInput)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function AddDrugReason(ByRef objMap As Object, ByRef rsOut As ADODB.Recordset) As Boolean
'------------------------------------------------------------------------
'����:����ҩƷ��ӽ���˵��
'����:
'objMap-���������
'rsOut-�������
'����:True-����ҽ�����棨�����ڽ���ҩƷ,������д����˵��;���ڽ���ҩƷ��������д����˵����,False-��ֹҽ�����棨���ڽ���ҩƷ�ҽ���ҩƷ˵��δ������д��
'˵��:��ҩ�䷽����˵����������ҩ
'-----------------------------------------------------------------------
    Dim i As Long
    Dim strReason As String
    
    If rsOut Is Nothing Then AddDrugReason = True: Exit Function
    
    rsOut.Filter = "�Ƿ����=1"
    
    For i = 1 To rsOut.RecordCount
        strReason = rsOut!����ҩƷ˵�� & ""
        Call zlCommFun.ShowMsgBox("����˵��", "^��鷢�ֽ�����ҩ:��" & rsOut!ҩƷ���� & "��" & _
            vbCrLf & vbCrLf & "����¼�������ҩ˵����������ҽ����^", "!ȷ��(&O),?ȡ��(&C)", objMap.frmMain, vbInformation, , , , , , "����˵����", 99, strReason)
        If strReason = "" Then
            Exit Function
        Else
            rsOut!����ҩƷ˵�� = strReason
        End If
        rsOut.MoveNext
    Next
    AddDrugReason = True
End Function

Public Function ReadXML(ByVal strXML As String) As ADODB.Recordset
'����:���ص���ҩƷ���ʾֵ
'xmlģ��
'    <his_results_xml fun_id="1006">
'    <result>
'       <type>ZDXGYWSY</type>
'       <level>2</level>
'       <prescA_hiscode>669</prescA_hiscode>
'       <mediA_hiscode>14686</mediA_hiscode>
'       <mediA_name>�����Ƭ</mediA_name>
'       <groupA>669</groupA>
'       <prescB_hiscode /><mediB_hiscode />
'       <mediB_name />
'        <groupB />
'    </result>
'    <result>
'    <type>XHZYWT</type>
'    <level>2</level>
'    <prescA_hiscode>669</prescA_hiscode>
'    <mediA_hiscode>14686</mediA_hiscode>
'    <mediA_name>�����Ƭ</mediA_name>
'    <groupA>669</groupA>
'    <prescB_hiscode>671</prescB_hiscode><mediB_hiscode>14250</mediB_hiscode>
'    <mediB_name>ά����CƬ</mediB_name>
'    <groupB>671</groupB>
'   </result>
'   <types>;ZDXGYWSY;XHZYWT;YHGXHCGYFYLWT_PC;YHGXHCGYFYLWT_DR;</types>
'</his_results_xml>


    Dim xmlDoc As DOMDocument
    Dim xmlRoot As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodes As IXMLDOMNodeList
    Dim rsRet As ADODB.Recordset
    
    Dim str��ʾֵ As String
    Dim strҽ��ID As String
    
    On Error GoTo errH
    
    Set xmlDoc = New DOMDocument
    xmlDoc.loadXML (strXML)
    '����������κ�Ԫ�أ����˳�
    If xmlDoc.documentElement Is Nothing Then
        Set xmlDoc = Nothing
        Exit Function
    End If
    
    Set rsRet = InitAdviceRS(FUN_�����)
    '��ȡXML����
    Set xmlRoot = xmlDoc.selectSingleNode("his_results_xml")
    Set xmlNodes = xmlRoot.selectNodes("result")

    If Not xmlNodes Is Nothing Then
        For Each xmlNode In xmlNodes
            str��ʾֵ = xmlNode.selectSingleNode("level").Text
            If Val(str��ʾֵ) > 0 Then
                strҽ��ID = xmlNode.selectSingleNode("prescA_hiscode").Text
                If Val(strҽ��ID) <> 0 Then
                    rsRet.Filter = "ҽ��ID ='" & strҽ��ID & "'"
                    If Not rsRet.EOF Then
                        If Val(rsRet!��ʾֵ & "") < Val(str��ʾֵ) Then
                            rsRet!��ʾֵ = str��ʾֵ
                        End If
                    Else
                        rsRet.AddNew
                        rsRet!��ʾֵ = str��ʾֵ
                        rsRet!ҽ��ID = strҽ��ID
                        rsRet.Update
                    End If
                End If
                strҽ��ID = xmlNode.selectSingleNode("prescB_hiscode").Text
                If Val(strҽ��ID) > 0 Then
                    rsRet.Filter = "ҽ��ID ='" & strҽ��ID & "'"
                    
                    If Not rsRet.EOF Then
                        If Val(rsRet!��ʾֵ & "") < Val(str��ʾֵ) Then
                            rsRet!��ʾֵ = str��ʾֵ
                        End If
                    Else
                        rsRet.AddNew
                        rsRet!��ʾֵ = str��ʾֵ
                        rsRet!ҽ��ID = strҽ��ID
                        rsRet.Update
                    End If
                End If
            End If
        Next
    End If
    
    If rsRet.RecordCount > 0 Then rsRet.Filter = ""
    
    Set ReadXML = rsRet
    Exit Function
errH:
    MsgBox "ReadXML �����:" & Err.Number & "��������:" & Err.Description, vbOKOnly, gstrSysName
End Function

Public Function FuncGetDripInfo(ByVal lngIndex As Long, ByVal strDrip As String, ByVal lngPharmacyCode As Long, ByVal strPharmacyName As String, ByVal strDuration As String) As String
'����:����ָ����JSON��
'�ַ�������:
'{ "type":"druginfo","index":"drug001","driprate":"60","driptime":"120","pharmacycode":"ҩ������","pharmacyname":"ҩ������","duration":"��ҩ����"}
'driprate��   60   ��ʾ  ÿ����60��
'driptime����ʾ��������Ҫ��ʱ�䣬���û�оʹ��մ�
'��������Ǹ�����ֵ���ʹ����ġ�
'��λΪ���� ���㵥λ1����=20��
    Dim strRet As String
    Dim arrTmp As Variant
    
    If InStr(strDrip, "��/����") > 0 Then
        strDrip = Replace(strDrip, "��/����", "")
        arrTmp = Split(strDrip, "-")
        If UBound(arrTmp) = 1 Then
            strDrip = arrTmp(1)
        Else
            strDrip = arrTmp(0)
        End If
         
    ElseIf InStr(strDrip, "����/Сʱ") > 0 Then
        strDrip = Replace(strDrip, "����/Сʱ", "")
        arrTmp = Split(strDrip, "-")
        If UBound(arrTmp) = 1 Then
            strDrip = arrTmp(1)
        Else
            strDrip = arrTmp(0)
        End If
        strDrip = (Val(strDrip) \ 60) * 20
    Else
        strDrip = ""
    End If
    strRet = "{""type"":""druginfo"",""index"":""" & lngIndex & """,""driprate"":""" & strDrip & """,""driptime"":""""," & _
            """pharmacycode"":""" & lngPharmacyCode & """,""pharmacyname"":""" & strPharmacyName & """,""duration"":""" & _
            strDuration & """}"
    FuncGetDripInfo = strRet
End Function

Public Function FuncGetOtherRecipInfo(ByVal strAdviceID As String, ByVal strRecipNo As String, ByVal strDrugCode As String, _
    ByVal strDrugName As String, ByVal strRouteName As String, ByVal strfrequency As String, ByVal strDoseunit As String, _
    ByVal strDosepertime As String, ByVal strNum As String, ByVal strNumUnit As String, ByVal strDuration As String) As String
    '����:������Ϣ��ʷҽ������,����ָ����JSON��
    Dim strRet As String
' //��ʷҽ����Ϣ
'        {
'            "type":"otherrecipinfo",
'            "hiscode":"his001",//�ַ������ͣ���������
'            "index":" drug001",//�ַ�������ҽ��Ψһ��
'            "recipno":"MZ12376",//�ַ������ͣ�������
'            "drugsource":"USER",//�ַ������ͣ�ҩƷ����
'            "druguniquecode":"123456",//�ַ�������ҽ��Ψһ��
'            "drugname":"��Ī���ֽ���",//�ַ������ͣ�ҩƷ����
'            "routeCode"��"�ڷ�" //�ַ������ͣ���ҩ;�������ø�ҩ;������
'            "routeName"��"�ڷ�"//�ַ������ͣ���ҩ;������
'            "routesource":"USER"
'            "frequency"��"bid"//�ַ������ͣ���ҩƵ��
'            "doseunit":"g"//�ַ������ͣ���ʾÿ��ʹ�ü�������ҩ��λ
'            "dosepertime":"5"//�ַ������ͣ���ʾÿ��ʹ�ü��������ֲ���
'            "num":"2"//�ַ������ͣ�ҩƷ�������������ﴦ�����ר�ã�סԺ���ա�
'            "numunit":"Ƭ"//�ַ������ͣ�ҩƷ����������λ�����ﴦ�����ר�ã�סԺ���ա�
'            "duration":"7" // ������ҩ����
'        }

     strRet = "{""type"":""otherrecipinfo"",""hiscode"":""" & gstrHOSCODE & """,""index"":""" & strAdviceID & """,""recipno"":""" & strRecipNo & """," & _
            """drugsource"":""USER"",""druguniquecode"":""" & strDrugCode & """,""drugname"":""" & strDrugName & """,""routeCode"":""" & strRouteName & """," & _
            """routeName"":""" & strRouteName & """,""routesource"":""USER"",""frequency"":""" & strfrequency & """,""doseunit"":""" & strDoseunit & """," & _
            """dosepertime"":""" & strDosepertime & """,""num"":""" & strNum & """,""duration"":""" & strDuration & """}"
    FuncGetOtherRecipInfo = strRet
End Function


Public Function StrConvToNormal(ByVal strIn As String) As String
'���ܣ���StrConv(str,vbFromUnicode)ת��ʱ,��ʱ����Ϊ�������뵼��ת����xml������һ��������Ч��XML����
    Dim strChar As String
    Dim strRet As String
    Dim i As Long
    
    For i = 1 To Len(strIn)
        strChar = Mid(strIn, i, 1)
        If InStr(G_STR_MATCH & "=", strChar) > 0 Then
            strRet = strRet & strChar
        End If
    Next
    StrConvToNormal = strRet
End Function

Public Sub WriteLog(ByVal strModule As String, ByVal strFunction As String, ByVal strLog As String)
'------------------------------------------------
'���ܣ�д����־
'������
'      strModule  ��ģ����
'      strFunction��������
'      strLog     ����־����
'��ע����ϵͳѡ���п�����־����д����־ʱ������־�ļ���ÿ������ÿ����־����ֻ����һ��
'------------------------------------------------
    LogWrite "������ҩ�ӿڵ�����־", strModule, strFunction, strLog
End Sub


'* ************************************** *
'* ģ�����ƣ�modCharset.bas
'* ģ�鹦�ܣ�GB2312��UTF8�໥ת������
'* ���ߣ�lyserver
'* ************************************** *

'- ------------------------------------------- -
'  ����˵����GB2312ת��ΪUTF8
'- ------------------------------------------- -

Public Function GB2312ToUTF8(strIn As String, Optional ByVal ReturnValueType As VbVarType = vbString) As Variant
    Dim adoStream As Object

    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 2 'adTypeText
    adoStream.Open
    adoStream.WriteText strIn
    adoStream.Position = 0
    adoStream.Type = 1 'adTypeBinary
    GB2312ToUTF8 = adoStream.Read()
    adoStream.Close

    If ReturnValueType = vbString Then GB2312ToUTF8 = Mid(GB2312ToUTF8, 1)
End Function

'- ------------------------------------------- -
'  ����˵����UTF8ת��ΪGB2312
'- ------------------------------------------- -
Public Function UTF8ToGB2312(ByVal varIn As Variant) As String
    Dim bytesData() As Byte
    Dim adoStream As Object

    bytesData = varIn
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 1 'adTypeBinary
    adoStream.Open
    adoStream.Write bytesData
    adoStream.Position = 0
    adoStream.Type = 2 'adTypeText
    UTF8ToGB2312 = adoStream.ReadText()
    adoStream.Close
End Function

Public Function WinHttpPost(ByVal strURL As String, ByVal strData As String, ByVal DataStic As DataEnum, Optional ByVal strHeader As String, Optional ByVal strMethod As String = "POST") As Variant
'֧��HTTPS����
'����:strHeader ��ֵ��ʽ��HeaderName:HeaderValue ����:CONTENT-TYPE:application/json
    Dim XMLHTTP As WinHttp.WinHttpRequest
    Dim DataS As String
    Dim DataB() As Byte
    Dim varHeader As Variant
    Dim varHeaderItem As Variant
    Dim i As Long

    On Error GoTo errH:
       
8      Set XMLHTTP = New WinHttpRequest
9      XMLHTTP.Open strMethod, strURL
10      If strHeader <> "" Then
            varHeader = Split(strHeader, ",")
            For i = LBound(varHeader) To UBound(varHeader)
                varHeaderItem = Split(varHeader(i), ":")
                XMLHTTP.setRequestHeader varHeaderItem(0), varHeaderItem(1)
            Next
        End If

13     XMLHTTP.send strData

110     Do Until XMLHTTP.Status = 200
112         DoEvents
        Loop

    '-----------------------------��������
114 Select Case DataStic
    Case responseText
        '--------------------------------ֱ�ӷ����ַ���
116     DataS = XMLHTTP.responseText
118     WinHttpPost = DataS
120 Case responseBody
        '--------------------------------ֱ�ӷ��ض�����
122     DataB = XMLHTTP.responseBody
124     WinHttpPost = DataS
126 Case responseBody + responseText
        '---------------------------������ת�ַ���[ֱ�ӷ����ִ���������ʱ����]
128     DataS = BytesToStr(XMLHTTP.responseBody)
130     WinHttpPost = DataS
132 Case Else
        '--------------------------------��Ч�ķ���
134     WinHttpPost = ""
    End Select

    '------------------------------------�ͷſռ�
136     Set XMLHTTP = Nothing

    Exit Function

errH:
138     WinHttpPost = ""
140     MsgBox "WinHttpPostʧ�ܣ�" & vbNewLine & "�����:" & Err.Number & vbCrLf & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "�������"
End Function

'==========================================================
'| ģ �� �� | XMLHTTP
'| ˵    �� | ���Inet�ؼ���ʵ������ͨѶ
'---------------------------------------------------------------------------����Begin����---------------------------------------------------------------------------------------
'==========================================================
Public Function HttpGet(ByVal Url As String, ByVal DataStic As DataEnum, Optional ByVal sngWaitTime As Single, Optional ByRef blnBreak As Boolean) As Variant
    Dim XMLHTTP As Object
    Dim DataS As String
    Dim DataB() As Byte
    Dim lngTime As Long
    On Error GoTo errH:
    
100 Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
102 XMLHTTP.Open "get", Url, True
104 XMLHTTP.send
    lngTime = Timer
106 Do While XMLHTTP.readyState <> 4
        If sngWaitTime = 0 Then
            If Timer - lngTime > gsngWaitTime Then blnBreak = True: Exit Function
        Else
            If Timer - lngTime > sngWaitTime Then blnBreak = True: Exit Function
        End If
108     DoEvents
    Loop
    blnBreak = False
    '--------------------------------------��������
110 Select Case DataStic
    Case responseText
        '--------------------------------ֱ�ӷ����ַ���
112     DataS = XMLHTTP.responseText
114     HttpGet = DataS
116 Case responseBody
        '--------------------------------ֱ�ӷ��ض�����
118     DataB = XMLHTTP.responseBody
120     HttpGet = DataB
122 Case responseBody + responseText
        '------------------------------������ת�ַ���[ֱ�ӷ����ִ���������ʱ����]
124     DataS = BytesToStr(XMLHTTP.responseBody)
126     HttpGet = DataS
128 Case Else
        '--------------------------------��Ч�ķ���
130     HttpGet = ""
    End Select

    '--------------------------------------�ͷſռ�
132 Set XMLHTTP = Nothing

    Exit Function

errH:
134 HttpGet = ""
136 MsgBox "HttpGetʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "�������"
End Function

Public Function HttpPost(ByVal strURL As String, ByVal strData As String, ByVal DataStic As DataEnum, _
    Optional ByVal strCONTENTTYPE As String, Optional ByVal strAUthorization As String, Optional ByVal sngWaitTime As Single, _
    Optional ByRef blnBreak As Boolean) As Variant
    Dim XMLHTTP As Object
    Dim DataS As String
    Dim DataB() As Byte
    Dim lngTime As Long
    
    On Error GoTo errH:

100 Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
102 XMLHTTP.Open "POST", strURL, True
104 XMLHTTP.setRequestHeader "Content-Length", Len(HttpPost)
    If strCONTENTTYPE = "" Then
106     XMLHTTP.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
    Else
        XMLHTTP.setRequestHeader "CONTENT-TYPE", strCONTENTTYPE  '"application/x-www-form-urlencoded; charset=utf-8"
    End If
    If strAUthorization <> "" Then
        XMLHTTP.setRequestHeader "AUthorization", strAUthorization
    End If
108 XMLHTTP.send (strData)
    lngTime = Timer
110 Do Until XMLHTTP.readyState = 4
        If sngWaitTime = 0 Then
            If Timer - lngTime > gsngWaitTime Then blnBreak = True: Exit Function
        Else
            If Timer - lngTime > sngWaitTime Then blnBreak = True: Exit Function
        End If
112     DoEvents
    Loop
    blnBreak = False
    '-----------------------------��������
114 Select Case DataStic
    Case responseText
        '--------------------------------ֱ�ӷ����ַ���
116     DataS = XMLHTTP.responseText
118     HttpPost = DataS
120 Case responseBody
        '--------------------------------ֱ�ӷ��ض�����
122     DataB = XMLHTTP.responseBody
124     HttpPost = DataS
126 Case responseBody + responseText
        '---------------------------������ת�ַ���[ֱ�ӷ����ִ���������ʱ����]
128     DataS = BytesToStr(XMLHTTP.responseBody)
130     HttpPost = DataS
    Case 6
        HttpPost = XMLHTTP.responseXML
132 Case Else
        '--------------------------------��Ч�ķ���
134     HttpPost = ""
    End Select

    '------------------------------------�ͷſռ�
136     Set XMLHTTP = Nothing

    Exit Function

errH:
138     HttpPost = ""
140     MsgBox "HttpPostʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, "�������"
End Function

Private Function BytesToStr(ByVal vInput As Variant) As String
    
    Dim strReturn       As String
    Dim i               As Long
    Dim intPrevCharCode As Integer
    Dim intNextCharCode As Integer

    For i = 1 To LenB(vInput)
        intPrevCharCode = AscB(MidB(vInput, i, 1))
        If intPrevCharCode < &H80 Then
            strReturn = strReturn & Chr(intPrevCharCode)
        Else
            intNextCharCode = AscB(MidB(vInput, i + 1, 1))
            strReturn = strReturn & Chr(CLng(intPrevCharCode) * &H100 + CInt(intNextCharCode))
            i = i + 1
        End If
    Next

    BytesToStr = strReturn
End Function

Public Function CreatePlugInOK() As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, glngModel)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub


'�Զ������Ϣ������
Public Function NewWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'����:��������¼����д���,�ǹ����¼�����Ĭ�ϴ�����Ϣ������
'����:vsc-VScrollBar ����
'     OldWindowProc Ĭ�ϴ�����Ϣ��������ַ
    On Error Resume Next
    If msg = WM_MOUSEWHEEL Then
        '���������¼����д���
        If wParam = -7864320 Then '���¹���
            If frmPassAsk.vsc.Value - 10 < frmPassAsk.vsc.Max Then
                frmPassAsk.vsc.Value = frmPassAsk.vsc.Max
            Else
                frmPassAsk.vsc.Value = frmPassAsk.vsc.Value - 10
            End If
        ElseIf wParam = 7864320 Then '���Ϲ���
            If frmPassAsk.vsc.Value + 10 > frmPassAsk.vsc.Min Then
                frmPassAsk.vsc.Value = frmPassAsk.vsc.Min
            Else
                frmPassAsk.vsc.Value = frmPassAsk.vsc.Value + 10
            End If
        End If
    Else
        '����Ĭ�ϴ�����Ϣ������
        NewWindowProc = CallWindowProc(glngOldWindowProc, hWnd, msg, wParam, lParam)
    End If
End Function
