Attribute VB_Name = "mdlLedVoice"
Option Explicit 'Ҫ���������

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gobjLED As Object
'-----------------------------------------------------------------------------------------------------------------------
'SYC XII
Public Declare Function dsbdllNt Lib "7CADSBNT.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer
Public Declare Function dsbdll98 Lib "7CACKY95.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer
Public Declare Function dsbdll16 Lib "7CACKY16.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer

'SYC Q9
Public Declare Function SYC_Q9 Lib "CKY95H.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer

'SHY-II
Public Declare Function shydsbdllNt Lib "CKY32.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer

'TDKJ_BJ_I/II
Public Declare Function TDKJ_BJ_FUN Lib "CKY95H.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer

'TDKJ_BJ_2008��
Public Declare Function TDKJ_BJ_2008 Lib "TdBjq.dll" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer


'Dev_MDT_SD01 �����SD-01������ʾ��

Public Declare Function InitService Lib "Service.dll" () As Boolean '������������������ʼ�������ӿڣ���������������������ǰ����
Public Declare Function InitDevice Lib "Service.dll" (ByVal lPort As Long) As Boolean '��ʼ�����ں������ڵ���"��������"�����Ժ󡢵�������������ǰʹ�øú�������ڲ���port��1-4��
Public Declare Function CloseService Lib "Service.dll" () As Boolean '�رշ������������ر������ӿڣ������еĵ��ý�������á�
Public Declare Function CloseDevice Lib "Service.dll" () As Boolean '�رմ��ں���
Public Declare Function Clear_Screen Lib "Service.dll" () As Boolean '����
Private Declare Function Display Lib "Service.dll" (ByVal lStrNum As Long) As Boolean
                                 '����ģ������,���Զ���Ϊ˽��,���°�װ,��ʾ������������ʾԤ�õĺ��ֵ���
Public Declare Function Voices Lib "Service.dll" (ByVal sCommand As String) As Boolean
                                 '������������������Ԥ��¼�Ƶ��������С���ڲ���Command�������������ַ�����ÿ��λ����һ������
Public Declare Function Price Lib "Service.dll" (ByVal sMoney As String) As Boolean
                                  'Ӧ�պ������������Ų���ʾӦ�յĽ���ڲ���Money����������ַ�����
Public Declare Function GetPrice Lib "Service.dll" Alias "Get" (ByVal sMoney As String) As Boolean
                                 'ʵ�պ������������Ų���ʾӦ�յĽ���ڲ���Money����������ַ�����
Public Declare Function Check Lib "Service.dll" (ByVal sMoney As String) As Boolean
                                 '���㺯�����������Ų���ʾӦ�յĽ���ڲ���Money����������ַ�����
Public Declare Function Medincine Lib "Service.dll" (ByVal sNumber As String) As Boolean
                                  'ȡҩ����, ��ʾ��������Ա��ĳ����ȡҩ����ڲ���Number�����ں�,����Ϊ����
Public Declare Function Display_Line Lib "Service.dll" (ByVal sTest As String, ByVal lSize As Long, ByVal lRow As Long) As Boolean
                                  '����ʾ������������ָ������ʾָ������ʹ�С�ĺ��֡�sTest������Ҫ��ʾ�ĺ��֣�lSize=�������ֺ�(0-6),lRow=�к�(0-3)

'SURPASS������Ϫ������ҽԺ�����շ���������LED(��֮ԣ)
Public Declare Function SetComNo Lib "Fgc01" (ByVal No As Long) As Long
Public Declare Sub SetQuickSwitch Lib "Fgc01" (ByVal Switch As Long)
Public Declare Sub SetHandleType Lib "Fgc01" (ByVal Handle As Long)
Public Declare Sub AllClear Lib "Fgc01" ()
Public Declare Sub PartClear Lib "Fgc01" (ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long)
Public Declare Sub AllDisplay Lib "Fgc01" (ByVal Handle As Long)
Public Declare Sub PartDisplay Lib "Fgc01" (ByVal Handle As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long)
Public Declare Sub StringDisplay Lib "Fgc01" (ByVal Str As String, ByVal Mode As Long, ByVal Delay As Long)
Public Declare Sub SetFontName Lib "Fgc01" (ByVal Name As String)
Public Declare Sub SetFontSize Lib "Fgc01" (ByVal Size As Long)
Public Declare Sub SetFontStyle Lib "Fgc01" (ByVal Style As Long)
Public Declare Sub LocStringDisplay Lib "Fgc01" (ByVal X As Long, ByVal Y As Long, ByVal Str As String)
Public Declare Sub PictureDisplay Lib "Fgc01" (ByVal Handle As Long, ByVal Length As Long, ByVal Mode As Long, ByVal Delay As Long)
Public Declare Sub MagicDisplay Lib "Fgc01" (ByVal Handle As Long, ByVal Mode As Long)
Public Declare Sub MagicClear Lib "Fgc01" (ByVal Mode As Long)
Public Declare Sub PickDisplay Lib "Fgc01" (ByVal Handle As Long, ByVal X0 As Long, ByVal Y0 As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long)
Public Declare Function PlayWaves Lib "Fgc01" (ByVal filenames As String) As Long
Public Declare Function RMB2Wav Lib "Fgc01" (ByVal VALDGT As Double) As Boolean
Public Declare Function Val2RMB Lib "Fgc01" (ByVal VALDGT As Double) As String
Public Declare Sub ClearWaves Lib "Fgc01" ()

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'**Begin********��˳FS-YL01��LED����+������ʾ��(����ʡ��ɽ��)***************
'**********2009-05-25 ZHQ                               ***************
Public Declare Sub opencomm Lib "datasend" (ByVal port As Integer)
'extern "C" __declspec(dllimport) int __stdcall opencomm(int port)
'���ܣ������û���Ҫ����Ӧ�Ĵ��ڡ����Ҫ�������ݴ�����Ҫ�ȴ򿪴��ڷ����޷��������ݡ���opencomm(1)���Ǵ򿪴��� 1

Public Declare Sub SendPray Lib "datasend" (ByVal money As Double)
'extern "C" __declspec(dllimport) int __stdcall SendPray(double money)
'���ܣ�����Ӧ�ս����շ�Ա��Ҫ��ȡ�Ľ���
'Ҫ�󣺽�ΧΪ0.00-9999999Ԫ

Public Declare Sub SendYs Lib "datasend" (ByVal money As Double)
'extern "C" __declspec(dllimport) int __stdcall SendYs(double money)
'����: ����ʵ���ֽ�
'Ҫ�󣺽�ΧΪ0.00-9999999Ԫ

Public Declare Sub SendChange Lib "datasend" (ByVal money As Double)
'extern "C" __declspec(dllimport) int __stdcall SendChange(double money)
'����: ����������
'Ҫ�󣺽�ΧΪ0.00-9999999Ԫ

Public Declare Sub SendName Lib "datasend" (ByVal bufName As String, ByVal Length As Integer)
'extern "C" __declspec(dllimport) int __stdcall SendName(unsigned char *buf,int Length)
'����: ��������
'Ҫ��: �����ĳ��Ȳ��ܳ���24���ֽ�

'Public Declare Sub SendCard Lib "datasend" (ByVal Handle As Long)
'extern "C" __declspec(dllimport) int __stdcall sendcard(void)
'����: ������ʾ "��ʾ���ƿ�"    (����ѯ�Է�����ʦ�����豸��֧�֣���Ҫȡ��)
'Ҫ��Ҫ��ʾ��ʱ��ֱ�ӵ��÷���

'Public Declare Sub SendWid Lib "datasend" (ByVal wid As Integer)
'extern "C" __declspec(dllimport) int __stdcall sendwid(int wid)
'����: ���ʹ��ں�               (����ѯ�Է�����ʦ�������豸��֧������,���Է�ҩ���ڲ����ô˳���)
'Ҫ����ʾ�뵽XX�Ŵ���ȡҩ������Ҫ���ͣ���ֵ��Χ1~99

'**End**********��˳FS-YL01��LED����+������ʾ��(����ʡ��ɽ��)***************

'2010-02-24 ZHQ һ����ҽԺ TDKJ_BJ_IV����������
Public Declare Function TDKJ_BJ_IV Lib "TdBjq.dll" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer
Public gBlnPic As Boolean   '��¼�Ƿ��״γ�ʼ��

'-----------------------------------------------------------------------------------------------------------------------
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128           'Ϊ�� PSS ����;��ά����
End Type
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Public gintOS As Integer                   '����ϵͳ  0��Win32��1:Win98/95��2��Win2000/NT
'-----------------------------------------------------------------------------------------------------------------------
Public Enum LEDDevice
    Dev_SYC_XII = 1      'SYC XII ������ʾ��
    Dev_LK822 = 2        'LK822 ����Һ����ʾ�ն�
    Dev_SHY_II = 3       'SHY-II �������Ա�����
    Dev_NJF_VH = 4       'NJF-VH ����������ʾ��
    Dev_TDKJ_BJ = 5      'TDKJ_BJ_I/II����������
    Dev_MDT_SD01 = 6     '�����SD-01������ʾ��
    Dev_surpass = 7      'SURPASS��������������ʾ��
    Dev_FS_YL01 = 8      'FS-YL01��LED����+������ʾ��
    Dev_TDKJ_BJ_2008 = 9 'TDKJ-BJ_2008�� ����������
    Dev_TDKJ_BJ_IV = 10  'TDKJ_BJ_IV ����������
    Dev_SYC_Q9 = 11      'SYC-Q9����������
    Dev_DDisplay = 99    '˫����ʾ��
End Enum

Public ctlComm As Object                   'MsComm�ؼ�
Public gintDevice As LEDDevice             '�豸��
Public gintPort As Long                    '�˿�      1��COM1;2��COM2;3��COM3;4��COM4
Public gstrSpeed As String                 '��������
Public gblnDDisplay As Boolean             'ʹ��˫����ʾ��

Public gblnHaveBottom As Integer           '��ʾ���еı�־ 1����ʾ��0������ʾ
Public gstrBottom As String                '������ʾ������
'��������2005-10-13��
Public gblnNewDev As Boolean               '�Ƿ�ʹ������SHY-II���豸
Public gbln������� As Boolean             '�Ƿ�Բ��˵ĸ����ʻ�������������ʾ

Public Function MoveObj(lngHwnd As Long) As RECT
'���ܣ��ڶ����MouseDown�¼��е���,����������Hwnd����
'���أ������Ļ������ֵ
   
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function
Public Sub Dev_MDT_SD01_Display(ByVal strShow As Long)
    'display ��ʾԤ��������
    '11  ����!����һ��    '12  ���Եȣ��ڶ���    '13  �뵱����壬��һ��    '18  ף�����տ������ڶ���
    Display (strShow)
End Sub

Public Sub Dev_MDT_SD01_Speak(ByVal strSpeak As String)
'���ܣ�����Dev_MDT_SD01��������ϵͳ����������,�������ţ�ͬʱLED��ʾ
'������strSpeak=����ָ���ʽ����Ҫת��Ϊ���豸֧�ֵ�����
'˵����
    'Display_Line(char *String,int m,int n) ��ָ������ʾָ������ʹ�С�ĺ���
    'm�ֺ�               n�к�
    '0   16X16   ����    4(0-3)
    '1   24X24   ������  2(0-1)
    '2   24X24   ����    2(0-1)
    '3   24X24   ����    2(0-1)
    '4   24X24   ����    2(0-1)
    '5   40X40   ����    1(0)
    '6   48X48   ����    1(0)
    '24*24����ÿ�����10������,40*40����ÿ�����5������
    
    'Voices ����ָ��  ����32������
    '01  ����        '02  ��ӭ����    '04  ���Ե�      '08  ף�����տ���
    '10  лл        '20  Ӧ��        '40  ʵ��        '80  ����
    '03  �뵱�����  '06  �ټ�        '0c  ����        '18  �뵽
    '30  ��ҩ��      '60  ��ҩ��      'C0  �Ŵ���ȡҩ  '81  �����
    '07 0e 1c 38 70 E0 C1 83 0f 1e 3c 78 F0 E1 C3 87
    'Ԫ �� �� ǧ �� ʮ 0  1  2  3  4  5  6  7  8  9
    
    Dim strMoney As String
    
    On Error Resume Next
    Clear_Screen '��Ҫ������,��������ȫ������ʷ��ʾ
    
    Select Case strSpeak
    Case "#0"
        Display_Line "����������", 5, 0 ''�޴˹��ܵķ���,����������ʾ
    Case "#1"  '--����,���Ե�"
        Display 11
        Display 12
        Voices "0104"
    Case "#2"
        Display_Line "лл", 5, 0
        Voices 10
    Case "#3"  '--�뵱�����, лл!"
        Display 13
        Display_Line "лл", 5, 0
        Voices "0310"
    Case "#4"
            Display_Line "������������", 4, 0
    Case "#5"
            Display_Line "������ʾ�ſ�", 4, 0
    Case "#6"
            Display_Line "��������ҩ������", 4, 0
    Case "#7"
            Display_Line "������X��������", 4, 0
    Case "#8"
            Display_Line "������ע������Ƥ��", 4, 0
    Case "#9"
            Display_Line "����������칫��", 4, 0
            Display_Line "��˸���", 4, 1
    Case "#10"
            Display_Line "�������Һ���", 4, 0
            Display_Line "���������", 4, 1
    Case "#11"
            Display_Line "������ʾ���֤��", 4, 0
            Display_Line "ҽ��ƾ֤", 4, 1
    Case "#12"
            Display_Line "������ʾ���֤��", 4, 0
            Display_Line "����ҽ��ƾ֤", 4, 1
    Case "#13"
            Display_Line "������ʾҽ��ƾ֤��", 4, 0
            Display_Line "����ҽ��ƾ֤", 4, 1
    Case "#14"
            Display_Line "��������ʲô��", 4, 0
    Case "#15"
            Display_Line "�������ǳ��ﻹ�Ǹ���", 4, 0
    Case "#16"
            Display_Line "��������ר������", 4, 0
            Display_Line "������ͨ����?", 4, 1
    Case "#17"
            Display_Line "������Ԥ��", 4, 0
            Display_Line "Ȼ���ٹҺ�", 4, 1
    Case "#18"
            Display_Line "��������ò�����", 4, 0
    Case "#19"
            Display_Line "������ʾ������", 4, 0
    Case "#20"
            Display_Line "������B��������", 4, 0
    Case "#50"
            Display_Line "������ʾҽ��ƾ֤", 4, 0
    Case Else
        '#21 1234.56   --��������һǧ������ʮ�ĵ�����Ԫ  J
        '#22 1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
        '#23 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
    '07 0e 1c 38 70 E0 C1 83 0f 1e 3c 78 F0 E1 C3 87
    'Ԫ �� �� ǧ �� ʮ 0  1  2  3  4  5  6  7  8  9
        strMoney = Trim(Mid(strSpeak, 4))
        strSpeak = Left(strSpeak, 3)
        If strSpeak = "#21" Then 'Ӧ��
            Price strMoney
        ElseIf strSpeak = "#22" Then 'ʵ��
            GetPrice strMoney
        ElseIf strSpeak = "#23" Then '����
            Check strMoney
        End If
    End Select

End Sub

Public Sub Dev_TDKJ_BJ_2008_Speak(ByVal strSpeak As String)
'���ܣ�����TDKJ_BJ_2008��������ϵͳ����������
'������strSpeak=����ָ���ʽ����Ҫת��Ϊ���豸֧�ֵ�����
'      bytType-0-�Һ�,1-�շ�
'˵�������������ļ���Ӧ��ģʽ��ͬ,��ͬ��������ò�ͬ

    Dim strMoney As String
    Dim strInFor As String
    
    On Error Resume Next
    Select Case strSpeak
    Case "#0"  '--����������
        '�޴˹���,��������ʾ
        Call TDKJ_BJ_2008(gintPort, "&Sc$")
        Call TDKJ_BJ_2008(gintPort, "&C21����������...$")
    Case "#1"  '--����,���Ե�"
        Call TDKJ_BJ_2008(gintPort, "W")
    Case "#2"  '--лл"
        Call TDKJ_BJ_2008(gintPort, "X")
    Case "#3"  '--�뵱�����, лл!"
        Call TDKJ_BJ_2008(gintPort, "D")
    Case "#4"  '--������������"
        strInFor = Trim(GetSetting("ZLSOFT", "����ȫ��", "�Һ���ʾ", "������������"))
        If strInFor = "������������" Then
            Call TDKJ_BJ_2008(gintPort, "k")
        ElseIf strInFor = "���������ӵ�����" Then
            Call TDKJ_BJ_2008(gintPort, "w")
        Else
            Call TDKJ_BJ_2008(gintPort, "&Sc$")
            Call TDKJ_BJ_2008(gintPort, "&C21" & strInFor & "$")
        End If
    Case "#5"  '--������ʾ�ſ�"
        Call TDKJ_BJ_2008(gintPort, "b")
    Case "#6"  '--��������ҩ������"
        Call TDKJ_BJ_2008(gintPort, "i")
    Case "#7"  '--������X��������"
        Call TDKJ_BJ_2008(gintPort, "j")
    Case "#8"  '--������ע������Ƥ��"
        Call TDKJ_BJ_2008(gintPort, "e")
    Case "#9"  '--����������칫����˸���"
        Call TDKJ_BJ_2008(gintPort, "g")
    Case "#10" '--�������Һ������������"
        Call TDKJ_BJ_2008(gintPort, "h")
    Case "#11" '--������ʾ���֤��ҽ��ƾ֤"
        Call TDKJ_BJ_2008(gintPort, "d")
    Case "#12" '--������ʾ���֤�͹���ҽ��ƾ֤"
        '--
    Case "#13" '--������ʾҽ��ƾ֤�͹���ҽ��ƾ֤"
        '--
    Case "#14" '--��������ʲô��"
        Call TDKJ_BJ_2008(gintPort, "e")
    Case "#15" '--�������ǳ��ﻹ�Ǹ���"
        Call TDKJ_BJ_2008(gintPort, "g")
    Case "#16" '--��������ר�����ﻹ����ͨ����"
        Call TDKJ_BJ_2008(gintPort, "h")
    Case "#17" '--������Ԥ��, Ȼ���ٹҺ�"
        Call TDKJ_BJ_2008(gintPort, "i")
    Case "#18" '--��������ò�����"
        
    Case "#19" '--������ʾ������"
        Call TDKJ_BJ_2008(gintPort, "j")
    Case "#20" '--������B��������"
        '--
    Case "#50"
        Call TDKJ_BJ_2008(gintPort, "a")
    Case "#51"  '�����շѣ������������
        strInFor = Trim(GetSetting("ZLSOFT", "����ȫ��", "�շ���ʾ", "������������"))
        If strInFor = "���ʾ���ĹҺ�Ʊ" Then
            Call TDKJ_BJ_2008(gintPort, "x")
        ElseIf strInFor = "������������" Then
            Call TDKJ_BJ_2008(gintPort, "k")
        ElseIf strInFor = "���������ӵ�����" Then
            Call TDKJ_BJ_2008(gintPort, "w")
        ElseIf strInFor = "������" Then
        Else
            Call TDKJ_BJ_2008(gintPort, "&Sc$")
            Call TDKJ_BJ_2008(gintPort, "&C21" & strInFor & "$")
        End If
    Case Else
        strMoney = Trim(Mid(strSpeak, 4))
        If Left(strSpeak, 3) = "#21" Then '��������
            Call TDKJ_BJ_2008(gintPort, strMoney & "J")
        ElseIf Left(strSpeak, 3) = "#22" Then 'Ԥ��
            Call TDKJ_BJ_2008(gintPort, strMoney & "Y")
        ElseIf Left(strSpeak, 3) = "#23" Then '����
            Call TDKJ_BJ_2008(gintPort, strMoney & "Z")
        End If
    End Select
End Sub

Public Sub Dev_TDKJ_BJ_Speak(ByVal strSpeak As String)
'���ܣ�����TDKJ_BJ_I/II��������ϵͳ����������
'������strSpeak=����ָ���ʽ����Ҫת��Ϊ���豸֧�ֵ�����
'˵�������������ļ���Ӧ��ģʽ��ͬ,��ͬ��������ò�ͬ
    Dim strMoney As String
    
    On Error Resume Next
    
    Select Case strSpeak
    Case "#0"  '--����������
        '�޴˹���,��������ʾ
        Call TDKJ_BJ_FUN(gintPort, "&Sc$")
        Call TDKJ_BJ_FUN(gintPort, "&C21����������...$")
    Case "#1"  '--����,���Ե�"
        DoEvents
        Call TDKJ_BJ_FUN(gintPort, "W")
    Case "#2"  '--лл"
        DoEvents
        Call TDKJ_BJ_FUN(gintPort, "X")
    Case "#3"  '--�뵱�����, лл!"
        Call TDKJ_BJ_FUN(gintPort, "D")
    Case "#4"  '--������������"
        DoEvents
        Call TDKJ_BJ_FUN(gintPort, "k")
    Case "#5"  '--������ʾ�ſ�"
        DoEvents
        Call TDKJ_BJ_FUN(gintPort, "b")
    Case "#6"  '--��������ҩ������"
        Call TDKJ_BJ_FUN(gintPort, "i")
    Case "#7"  '--������X��������"
        Call TDKJ_BJ_FUN(gintPort, "j")
    Case "#8"  '--������ע������Ƥ��"
        Call TDKJ_BJ_FUN(gintPort, "e")
    Case "#9"  '--����������칫����˸���"
        Call TDKJ_BJ_FUN(gintPort, "g")
    Case "#10" '--�������Һ������������"
        Call TDKJ_BJ_FUN(gintPort, "h")
    Case "#11" '--������ʾ���֤��ҽ��ƾ֤"
        Call TDKJ_BJ_FUN(gintPort, "d")
    Case "#12" '--������ʾ���֤�͹���ҽ��ƾ֤"
        '--
    Case "#13" '--������ʾҽ��ƾ֤�͹���ҽ��ƾ֤"
        '--
    Case "#14" '--��������ʲô��"
        Call TDKJ_BJ_FUN(gintPort, "e")
    Case "#15" '--�������ǳ��ﻹ�Ǹ���"
        Call TDKJ_BJ_FUN(gintPort, "g")
    Case "#16" '--��������ר�����ﻹ����ͨ����"
        Call TDKJ_BJ_FUN(gintPort, "h")
    Case "#17" '--������Ԥ��, Ȼ���ٹҺ�"
        Call TDKJ_BJ_FUN(gintPort, "i")
    Case "#18" '--��������ò�����"
        
    Case "#19" '--������ʾ������"
        Call TDKJ_BJ_FUN(gintPort, "j")
    Case "#20" '--������B��������"
        '--
    Case "#50"
        Call TDKJ_BJ_FUN(gintPort, "a")
    Case Else
        strMoney = Trim(Mid(strSpeak, 4))
        If Left(strSpeak, 3) = "#21" Then '��������
            Call TDKJ_BJ_FUN(gintPort, strMoney & "J")
        ElseIf Left(strSpeak, 3) = "#22" Then 'Ԥ��
            Call TDKJ_BJ_FUN(gintPort, strMoney & "Y")
        ElseIf Left(strSpeak, 3) = "#23" Then '����
            Call TDKJ_BJ_FUN(gintPort, strMoney & "Z")
        End If
    End Select
End Sub

Public Sub Dev_TDKJ_BJ_IV_Speak(ByVal strSpeak As String)
'���ܣ�����TDKJ_BJ_IV ��������ϵͳ����������
'������strSpeak=����ָ���ʽ����Ҫת��Ϊ���豸֧�ֵ�����
'˵�������������ļ���Ӧ��ģʽ��ͬ,��ͬ��������ò�ͬ
    Dim strMoney As String
    
    On Error Resume Next
    
    Select Case strSpeak
    Case "#0"  '--����������
        '�޴˹���,��������ʾ
        Call TDKJ_BJ_IV(gintPort, "&Sc$")
        Call TDKJ_BJ_IV(gintPort, "&C21����������...$")
    Case "#1"  '--����,���Ե�"
        DoEvents
        Call TDKJ_BJ_IV(gintPort, "W")
    Case "#2"  '--лл"
        DoEvents
        Call TDKJ_BJ_IV(gintPort, "X")
    Case "#3"  '--�뵱�����, лл!"
        Call TDKJ_BJ_IV(gintPort, "D")
    Case "#4"  '--������������"
        DoEvents
        Call TDKJ_BJ_IV(gintPort, "k")
    Case "#5"  '--������ʾ�ſ�"
        DoEvents
        Call TDKJ_BJ_IV(gintPort, "b")
    Case "#6"  '--��������ҩ������"
        Call TDKJ_BJ_IV(gintPort, "i")
    Case "#7"  '--������X��������"
        Call TDKJ_BJ_IV(gintPort, "j")
    Case "#8"  '--������ע������Ƥ��"
        Call TDKJ_BJ_IV(gintPort, "e")
    Case "#9"  '--����������칫����˸���"
        Call TDKJ_BJ_IV(gintPort, "g")
    Case "#10" '--�������Һ������������"
        Call TDKJ_BJ_IV(gintPort, "h")
    Case "#11" '--������ʾ���֤��ҽ��ƾ֤"
        Call TDKJ_BJ_IV(gintPort, "d")
    Case "#12" '--������ʾ���֤�͹���ҽ��ƾ֤"
        '--
    Case "#13" '--������ʾҽ��ƾ֤�͹���ҽ��ƾ֤"
        '--
    Case "#14" '--��������ʲô��"
        Call TDKJ_BJ_IV(gintPort, "e")
    Case "#15" '--�������ǳ��ﻹ�Ǹ���"
        Call TDKJ_BJ_IV(gintPort, "g")
    Case "#16" '--��������ר�����ﻹ����ͨ����"
        Call TDKJ_BJ_IV(gintPort, "h")
    Case "#17" '--������Ԥ��, Ȼ���ٹҺ�"
        Call TDKJ_BJ_IV(gintPort, "i")
    Case "#18" '--��������ò�����"
        
    Case "#19" '--������ʾ������"
        Call TDKJ_BJ_IV(gintPort, "j")
    Case "#20" '--������B��������"
        '--
    Case "#50"
        Call TDKJ_BJ_IV(gintPort, "a")
    Case Else
        strMoney = Trim(Mid(strSpeak, 4))
        If Left(strSpeak, 3) = "#21" Then '��������
            Call TDKJ_BJ_IV(gintPort, strMoney & "J")
        ElseIf Left(strSpeak, 3) = "#22" Then 'Ԥ��
            Call TDKJ_BJ_IV(gintPort, strMoney & "Y")
        ElseIf Left(strSpeak, 3) = "#23" Then '����
            Call TDKJ_BJ_IV(gintPort, strMoney & "Z")
        End If
    End Select
End Sub

Public Sub Contrast_NJF_VH(ByVal strCommand As String)
'���ܣ�����NJF-VH ����������ʾ��
'������strCommand=SYC-X����������(���ʼ��������������),��Ҫת��ΪNJF-VH������
'˵����"#���"��"#��� ���"
    Dim intNum As Integer, strMoney As String
    Dim strDisp As String, strVoice As String
        
    If InStr(strCommand, " ") > 0 Then
        strCommand = Replace(strCommand, "  ", " ")
        strCommand = Replace(strCommand, "  ", " ")
        intNum = Val(Mid(Split(strCommand, " ")(0), 2))
        strMoney = Split(strCommand, " ")(1)
    Else
        intNum = Val(Mid(strCommand, 2))
    End If
    
    On Error Resume Next
    Set gobjLED = CreateObject("CTSVR.Bjq")
    If Not gobjLED Is Nothing Then
        Select Case intNum
            Case 1
                strDisp = "~����,���Ե�!"
                strVoice = "_H"
            Case 2
                strDisp = "~лл!"
                strVoice = "_T"
            Case 3
                strDisp = "~�����뵱�����,лл!"
                strVoice = "_C"
            Case 4
                strDisp = "~������������?"
                strVoice = "eY"
            Case 5
                strDisp = "~������ʾ�ſ�!"
                strVoice = "gY"
            Case 6
                strDisp = "~��������ҩ������!"
                strVoice = "bX"
            Case 7
                strDisp = "~������X��������!"
                strVoice = "eX"
            Case 8
                strDisp = "~������ע������Ƥ��!"
                strVoice = ""
            Case 9
                strDisp = "~����������칫����˸���!"
                strVoice = "pY"
            Case 10
                strDisp = "~�������Һ������������!"
                strVoice = "aX"
            Case 11
                strDisp = "~������ʾ���֤��ҽ��ƾ֤!"
                strVoice = "hY" '������ʾ���֤
            Case 12
                strDisp = "~������ʾ���֤�͹���ҽ��ƾ֤!"
                strVoice = "hY" '������ʾ���֤
            Case 13
                strDisp = "~������ʾҽ��ƾ֤�͹���ҽ��ƾ֤!"
                strVoice = "hY" '������ʾ���֤
            Case 14
                strDisp = "~��������ʲô��?"
                strVoice = "bY"
            Case 15
                strDisp = "~�������ǳ��ﻹ�Ǹ���?"
                strVoice = "cY"
            Case 16
                strDisp = "~��������ר�����ﻹ����ͨ����?"
                strVoice = "dY"
            Case 17
                strDisp = "~������Ԥ�Ȼ���ٹҺ�!"
                strVoice = "mY"
            Case 18
                strDisp = "~��������ò�����!"
                strVoice = "lY"
            Case 19
                strDisp = "~������ʾ������!"
                strVoice = "jY"
            Case 20
                strDisp = "~������B��������!"
                strVoice = "gX"
            Case 21
                strDisp = "~��������:" & strMoney & "Ԫ"
                strVoice = strMoney & "_P"
            Case 22
                strDisp = "~����:" & strMoney & "Ԫ"
                strDisp = "" '�ѵ���������ʾ
                strVoice = strMoney & "_k"
            Case 23
                If Val(strMoney) <> 0 Then
                    strDisp = "~����:" & strMoney & "Ԫ"
                    strDisp = "" '�ѵ���������ʾ
                    strVoice = strMoney & "_b"
                Else
                    strDisp = ""
                    strVoice = ""
                End If
            Case Else
        End Select
        If strDisp <> "" Or strVoice <> "" Then
            gobjLED.Comport = gintPort
            gobjLED.DispMode = 0
            If strDisp <> "" Then gobjLED.Display strDisp
            If strVoice <> "" Then gobjLED.stdSpeak strVoice
        End If
        Set gobjLED = Nothing
    End If
End Sub
Public Sub ContrastSYC_Q9(ByVal strCommand As String)
    '���ܣ���SYC-XII������ʾ���������Ӧ����
    '������strcommand ʵ�ʵ�����
    '���ߣ�Ƚ��
    Dim strFront As String, strLast As String, strMoney As String
    Dim intLocation As Integer
    
    If Left(strCommand, 1) = "~" Then
        SycVoice Mid(strCommand, 2)
        Exit Sub
    End If
    intLocation = InStr(1, strCommand, " ")
    If intLocation <> 0 Then
        strFront = Left(strCommand, intLocation - 1)
        strLast = Trim(Mid(strCommand, intLocation + 1))
    Else
        strFront = strCommand
        strLast = ""
    End If
    
    Select Case strFront
        Case "#0"  '--����������
            '�޴˹���,��������ʾ
            Call SYC_Q9(gintPort, "*")
            Call SYC_Q9(gintPort, "#����������...#")
        Case "#1"   ': W --����, ���Ե�
            Call SYC_Q9(gintPort, "*")
            Call SYC_Q9(gintPort, "W")
        Case "#2"   ':X  --лл
            Call SYC_Q9(gintPort, "X")
        Case "#3"   'D  --�뵱�����, лл!
            Call SYC_Q9(gintPort, "D")
        Case "#4"   'j  --������������
            Call SYC_Q9(gintPort, "j")
        Case "#5"   'b  --������ʾ�ſ�
            Call SYC_Q9(gintPort, "b")
        Case "#6"   'h  --��������ҩ������
            Call SYC_Q9(gintPort, "h")
        Case "#7"   'i  --������X��������
            Call SYC_Q9(gintPort, "i")
        Case "#8"   'e  --������ע������Ƥ��
            Call SYC_Q9(gintPort, "e")
        Case "#9"   'g  --����������칫����˸���
            Call SYC_Q9(gintPort, "g")
        Case "#10"  'h --�������Һ������������
            Call SYC_Q9(gintPort, "h")
        Case "#11"  '# c --������ʾ���֤��ҽ��ƾ֤
             Call SYC_Q9(gintPort, "c")
        Case "#12"  'c --������ʾ���֤�͹���ҽ��ƾ֤
            Call SYC_Q9(gintPort, "c")
        Case "#13"  'k --������ʾҽ��ƾ֤�͹���ҽ��ƾ֤
            Call SYC_Q9(gintPort, "k")
        Case "#14"  'l --��������ʲô��
            Call SYC_Q9(gintPort, "l")
        Case "#15"  'm --�������ǳ��ﻹ�Ǹ���
            Call SYC_Q9(gintPort, "m")
        Case "#16"  'n --��������ר�����ﻹ����ͨ����
             Call SYC_Q9(gintPort, "n")
        Case "#17"  'o --������Ԥ��, Ȼ���ٹҺ�
            Call SYC_Q9(gintPort, "o")
        Case "#18"  'p --��������ò�����
            Call SYC_Q9(gintPort, "p")
        Case "#19"  'q --������ʾ������
            Call SYC_Q9(gintPort, "q")
        Case "#20"  'q --������B��������
            Call SYC_Q9(gintPort, "q")
        Case "#21"  '1234.56   --��������һǧ������ʮ�ĵ�����Ԫ  J
            Call SYC_Q9(gintPort, strLast & "J")
        Case "#22"  '1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
            Call SYC_Q9(gintPort, strLast & "Y")
        Case "#23"  ' 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
            Call SYC_Q9(gintPort, strLast & "Z")
        Case "#30"   '��ʾ���￨
            Call SYC_Q9(gintPort, "#���ʾ���￨��#")
        Case "#50" '���ʾ��ᱣ�Ͽ�
            Call SYC_Q9(gintPort, "a")
        Case "#51" '������������
            Call SYC_Q9(gintPort, "j")
        Case Else
    End Select
End Sub
Public Sub ContrastSYC_XII(ByVal strCommand As String)
    '���ܣ���SYC-XII������ʾ���������Ӧ����
    '������strcommand ʵ�ʵ�����
        
    Dim strFront As String, strLast As String
    Dim intLocation As Integer
    
    If Left(strCommand, 1) = "~" Then
        SycVoice Mid(strCommand, 2)
        Exit Sub
    End If
    intLocation = InStr(1, strCommand, " ")
    If intLocation <> 0 Then
        strFront = Left(strCommand, intLocation - 1)
        strLast = Trim(Mid(strCommand, intLocation + 1))
    Else
        strFront = strCommand
        strLast = ""
    End If
    
    Select Case strFront
        Case "#1"   ': W --����, ���Ե�
            SycVoice "W"
        Case "#2"   ':X  --лл
            SycVoice "X"
        Case "#3"   'D  --�뵱�����, лл!
            SycVoice "D"
        Case "#4"   'a  --������������
            SycVoice "a"
        Case "#5"   'b  --������ʾ�ſ�
            SycVoice "b"
        Case "#6"   'c  --��������ҩ������
            SycVoice "c"
        Case "#7"   'd  --������X��������
            SycVoice "d"
        Case "#8"   'e  --������ע������Ƥ��
            SycVoice "e"
        Case "#9"   'g  --����������칫����˸���
            SycVoice "g"
        Case "#10"  'h --�������Һ������������
            SycVoice "h"
        Case "#11"  '# i --������ʾ���֤��ҽ��ƾ֤
            SycVoice "i"
        Case "#12"  'j --������ʾ���֤�͹���ҽ��ƾ֤
            SycVoice "j"
        Case "#13"  'k --������ʾҽ��ƾ֤�͹���ҽ��ƾ֤
            SycVoice "k"
        Case "#14"  'l --��������ʲô��
            SycVoice "l"
        Case "#15"  'm --�������ǳ��ﻹ�Ǹ���
            SycVoice "m"
        Case "#16"  'n --��������ר�����ﻹ����ͨ����
            SycVoice "n"
        Case "#17"  'o --������Ԥ��, Ȼ���ٹҺ�
            SycVoice "o"
        Case "#18"  'p --��������ò�����
            SycVoice "p"
        Case "#19"  'q --������ʾ������
            SycVoice "q"
        Case "#20"  'r --������B��������
            SycVoice "r"
        Case "#21"  '1234.56   --��������һǧ������ʮ�ĵ�����Ԫ  J
            SycVoice strLast & "J"
        Case "#22"  '1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
            SycVoice strLast & "Y"
        Case "#23"  ' 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
            SycVoice strLast & "Z"
        Case "#30"   '  ���ʾ���￨(��ҽҪ��:):32663
            SycVoice "p"
        Case Else
    End Select
End Sub

Public Sub ContrastSHY_II(ByVal strCommand As String)
    '���ܣ���SHY-II�������Ա������������Ӧ����
    '������strcommand ʵ�ʵ�����
        
    Dim strFront As String, strLast As String
    Dim intLocation As Integer
    
    If Left(strCommand, 1) = "~" Then
        SHYVoice Mid(strCommand, 2)
        Exit Sub
    End If
    intLocation = InStr(1, strCommand, " ")
    If intLocation <> 0 Then
        strFront = Left(strCommand, intLocation - 1)
        strLast = Trim(Mid(strCommand, intLocation + 1))
    Else
        strFront = strCommand
        strLast = ""
    End If
    '��������2005-10-12�� ���ֹ���ֻ�����豸֧�֣����Զ��������ж�
    If gblnNewDev Then 'ʹ�������豸��֧��ҽ������
        Select Case strFront
            Case "#0" 'g  --����������
                SHYVoice "g"
            Case "#1"   ': W --����, ���Ե�
                SHYVoice "W"
            Case "#2"   ':X  --лл
                SHYVoice "X"
            Case "#3"   'D  --�뵱�����, лл!
                SHYVoice "D"
            Case "#4"   'a  --������������
                SHYVoice "a"
            Case "#5"   'b  --����ˢ��
                SHYVoice "b"
            Case "#6"   'b  --��������ҩ������
                'SHYVoice "b"
            Case "#7"   'c  --������X��������
                'SHYVoice "c"
            Case "#8"   'd  --������ע������Ƥ��
                'SHYVoice "d"
            Case "#9"   'e  --����������칫����˸���
                'SHYVoice "e"
            Case "#10"  'f --�������Һ������������
                'SHYVoice "f"
            Case "#11"  '# i --������ʾ���֤��ҽ��ƾ֤
                'SHYVoice "i"
            Case "#12"  'j --������ʾ���֤�͹���ҽ��ƾ֤
                'SHYVoice "j"
            Case "#13"  'k --������ʾҽ��ƾ֤�͹���ҽ��ƾ֤
                'SHYVoice "k"
            Case "#14"  'd --��������ʲô��
                SHYVoice "d"
            Case "#15"  'h --�������ǳ��ﻹ�Ǹ���
                SHYVoice "h"
            Case "#16"  'f --��������ר�����ﻹ����ͨ����
                SHYVoice "f"
            Case "#17"  'o --������Ԥ��, Ȼ���ٹҺ�
                'SHYVoice "o"
            Case "#18"  'p --��������ò�����
                'SHYVoice "p"
            Case "#19"  'h --������ʾ������
                'SHYVoice "h"
            Case "#20"  'g --������B��������
                'SHYVoice "g"
            Case "#21"  '1234.56   --��������һǧ������ʮ�ĵ�����Ԫ  J
                SHYVoice strLast & "J"
            Case "#22"  '1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
                SHYVoice strLast & "Y"
            Case "#23"  ' 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
                SHYVoice strLast & "Z"
            Case "#24"   'c --�����ʾ�籣��
                SHYVoice "c"
            Case "#25"  'e --��ķ���Ϊ***Ԫ
                If gbln������� Then SHYVoice strLast & "e"
            Case "#26"  'f --�������***Ԫ
                If gbln������� Then SHYVoice strLast & "f"
            Case "#27"  'i --��Ŀ��������븶�ֽ�**Ԫ
                If gbln������� Then SHYVoice strLast & "i"
            Case "#28"  'X --��������ҽ����ݼ���
                SHYVoice "X"
            Case Else
            
        End Select
    Else
        Select Case strFront
            Case "#0" 'g  --����������
                ' SHYVoice "g"
            Case "#1"   ': W --����, ���Ե�
                SHYVoice "W"
            Case "#2"   ':X  --лл
                SHYVoice "X"
            Case "#3"   'D  --�뵱�����, лл!
                SHYVoice "D"
            Case "#4"   'a  --������������
                SHYVoice "a"
            Case "#5"   'b  --������ʾ�ſ�
                ' SHYVoice "b"
            Case "#6"   'b  --��������ҩ������
                SHYVoice "b"
            Case "#7"   'c  --������X��������
                SHYVoice "c"
            Case "#8"   'd  --������ע������Ƥ��
                SHYVoice "d"
            Case "#9"   'e  --����������칫����˸���
                SHYVoice "e"
            Case "#10"  'f --�������Һ������������
                SHYVoice "f"
            Case "#11"  '# g --������ʾ���֤��ҽ��ƾ֤
                SHYVoice "g"
            Case "#12"  'j --������ʾ���֤�͹���ҽ��ƾ֤
                'SHYVoice "j"
            Case "#13"  'k --������ʾҽ��ƾ֤�͹���ҽ��ƾ֤
                'SHYVoice "k"
            Case "#14"  'b --��������ʲô��
                SHYVoice "b"
            Case "#15"  'c --�������ǳ��ﻹ�Ǹ���
                SHYVoice "c"
            Case "#16"  'd --��������ר�����ﻹ����ͨ����
                SHYVoice "d"
            Case "#17"  'e --������Ԥ��, Ȼ���ٹҺ�
                SHYVoice "e"
            Case "#18"  'p --��������ò�����
               ' SHYVoice "p"
            Case "#19"  'h --������ʾ������
                SHYVoice "h"
            Case "#20"  'g --������B��������
                ' SHYVoice "g"
            Case "#21"  '1234.56   --��������һǧ������ʮ�ĵ�����Ԫ  J
                SHYVoice strLast & "J"
            Case "#22"  '1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
                SHYVoice strLast & "Y"
            Case "#23"  ' 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
                SHYVoice strLast & "Z"
            Case Else
            
        End Select
    End If
End Sub
    
Public Function SycVoice(ByVal OutString As String) As Integer
    '���ܣ�SYC XII������ʾ����������ʾ
    '������
    'outstring:���� F,W,X,J,Y,Z,D�����д
    'F --��λ����
    'W          --����,���Ե�
    'X  --лл
    '1234.56J   --��������һǧ������ʮ�ĵ�����Ԫ
    '1234.56Y   --Ԥ��һǧ������ʮ�ĵ�����Ԫ
    '1234.56Z   --����һǧ������ʮ�ĵ�����Ԫ
    'D - -�뵱�����, лл!
    'a - -������������
    'b - -������ʾ�ſ�
    'c - -��������ҩ������
    'd - -������X��������
    'e - -������ע������Ƥ��
    'g - -����������칫����˸���
    'h - -�������Һ������������
    'i - -������ʾ���֤��ҽ��ƾ֤
    'j - -������ʾ���֤�͹���ҽ��ƾ֤
    'k - -������ʾҽ��ƾ֤�͹���ҽ��ƾ֤
    'l - -��������ʲô��
    'm - -�������ǳ��ﻹ�Ǹ���
    'n - -��������ר�����ﻹ����ͨ����
    'o - -������Ԥ��, Ȼ���ٹҺ�
    'p - -��������ò�����
    'q - -������ʾ������
    'r - -������B��������
    '��--����
    '*--����
    '�����Ҫ�ڵ�һ����ʾ���Ż�����������:
    '$1  ($2Ϊ�ڶ���)
    '#__���յ���Һ�:1234#(__Ϊ�ո�)
    
    '���أ�-1:���ɹ����������ɹ�
        
    Dim intresult As Integer
'    On Error goto errhandel
    
    intresult = 0
    Select Case gintOS
        Case 0
            intresult = dsbdll16(gintPort, OutString)
        Case 1
            intresult = dsbdll98(gintPort, OutString)
        Case 2
            intresult = dsbdllNt(gintPort, OutString)
    End Select
    
    SycVoice = intresult
    Exit Function
errHandle:
    SycVoice = -1
End Function

Public Function SHYVoice(ByVal OutString As String) As Integer
    '���ܣ�SYC XII������ʾ����������ʾ
    '������
    'outstring:���� F,W,X,J,Y,Z,D�����д
    'F --��λ����
    'W          --����,���Ե�
    'X  --лл
    '1234.56J   --��������һǧ������ʮ�ĵ�����Ԫ
    '1234.56Y   --Ԥ��һǧ������ʮ�ĵ�����Ԫ
    '1234.56Z   --����һǧ������ʮ�ĵ�����Ԫ
    'D - -�뵱�����, лл!
    'a - -������������
    'b - -��������ҩ������
    'c - -������X��������
    'd - -������ע������Ƥ��
    'e - -����������칫����˸���
    'f - -�������Һ������������
    'g - -����������
    'i - -������ʾ���֤��ҽ��ƾ֤
    'j - -������ʾ���֤�͹���ҽ��ƾ֤
    'k - -������ʾҽ��ƾ֤�͹���ҽ��ƾ֤
    'l - -��������ʲô��
    'm - -�������ǳ��ﻹ�Ǹ���
    'n - -��������ר�����ﻹ����ͨ����
    'o - -������Ԥ��, Ȼ���ٹҺ�
    'p - -��������ò�����
    'q - -������ʾ������
    'r - -������B��������
    
    'DSBDLL(1,'a')  --��������������
'    DSBDLL(1,'b')  --��������ҩ�����ۡ�
'    DSBDLL(1,'c')  --������X�������ۡ�
'    DSBDLL(1,'d')  --������ע������Ƥ�ԡ�
'    DSBDLL(1,'e')  --����������칫����˸��¡�
'    DSBDLL(1,'f')  --�������Һ�����������š�
'    DSBDLL(1,'g')  --���������롣
'    DSBDLL(1,'h')  --�����Ѳ������ó�����
'    DSBDLL(1,'i')  --�����뵱�����,лл��
    
    '���أ�-1:���ɹ����������ɹ�
        
    Dim intresult As Integer
    
    intresult = 0
    Select Case gintOS
        Case 0
            intresult = shydsbdllNt(gintPort, OutString)
        Case 1
            intresult = shydsbdllNt(gintPort, OutString)
        Case 2
            intresult = shydsbdllNt(gintPort, OutString)
    End Select
    
    SHYVoice = intresult
    Exit Function
errHandle:
    SHYVoice = -1
End Function


Public Sub Dev_surpass_speak(ByVal strSpeak As String)
'���ܣ�����Dev_surpass��������ϵͳ����������
'������strSpeak=����ָ���ʽ����Ҫת��Ϊ���豸֧�ֵ�����
'˵�������������ļ���Ӧ��ģʽ��ͬ,��ͬ��������ò�ͬ
    Dim filenames As String
    Dim strMoney As String
    Dim strӦ�� As String, strʵ�� As String, str���� As String, str���� As String, str�븶�� As String
    Dim dbl�ϼ� As Double
    On Error Resume Next

    strӦ�� = "Ӧ��.wav"
    strʵ�� = "Ԥ��.wav"
    str���� = "����.wav"
    str�븶�� = "��������.wav"
    str���� = "�����뵱�����лл.wav"
    
    Select Case strSpeak
           Case "#50"
                Call AllClear
                Call LocStringDisplay(2, 22, "���ã����Եȣ�")
                Call PlayWaves(App.Path & "\���ʾҽ����.wav")
           Case "#1"
                Call AllClear
                Call PlayWaves(App.Path & "\���Ե�.wav")
           Case Else
                strMoney = Trim(Mid(strSpeak, 4))
                If Left(strSpeak, 3) = "#21" Then '��������
                    Call AllClear
                    Call LocStringDisplay(2, 2, "Ӧ�գ�" & Format(strMoney, "0.00") & "Ԫ" + Chr(0))
                    str�븶�� = "��������.wav"
                    Call PlayWaves(App.Path & "\" & str�븶��)
                    Call RMB2Wav(strMoney)
                ElseIf Left(strSpeak, 3) = "#22" Then 'Ԥ��
                    Call LocStringDisplay(2, 22, "Ԥ�գ�" & Format(strMoney, "0.00") & "Ԫ" + Chr(0))
                    Call PlayWaves(App.Path & "\" & strʵ��)
                    Call RMB2Wav(strMoney)
                ElseIf Left(strSpeak, 3) = "#23" Then '����
                    If strMoney > 0 Then
                        Call LocStringDisplay(2, 42, "���㣺" & Format(strMoney, "0.00") & "Ԫ" + Chr(0))
                        Call PlayWaves(App.Path & "\" & str����)
                        Call RMB2Wav(strMoney)
                        Call PlayWaves(App.Path & "\" & str����)
                    End If
                End If
                             
    End Select
End Sub

Public Sub Dev_FS_YL01_Voice(ByVal varTemp As Variant, ByVal intType As Byte, ByVal lngSec As Long)
'���ܣ�����Dev_FS_YL01��������ϵͳ���������������
'������varTemp  �������ַ�(����),Ҳ��������(���)
'      intType  ��������(0-����;1-Ӧ�ս��;2-ʵ�ս��;3-�Ҳ����
'      lngSec   ���ʱ��,����Ϊ��λ,Ϊ0��ʾ��ͣ��
'˵�������ݴ��������ͬ��ʾ������,���ڸ��豸������ֱ�����,
'      ������������֮�����û��ͣ��,�����֮ǰû��˵������ݻᱻ��������ض�,����������������֮����Ҫ�����⴦��.
'���ƣ�2009-05-25 ZHQ

    Dim dtNow As Variant
    
    Select Case intType
    Case 0  '����
        Call SendName(varTemp, LenB(varTemp))
    Case 1  'Ӧ�ս��
        Call SendPray(Round(varTemp, 2))
    Case 2  'ʵ�ս��
        Call SendYs(Round(varTemp, 2))
    Case 3  '�Ҳ����
        Call SendChange(Round(varTemp, 2))
    End Select
    
    dtNow = Time
    Do While True
        If Int((Time - dtNow) * 24 * 60 * 60) >= lngSec Then Exit Do
    Loop
End Sub

Public Sub ShowLED(ByVal strRow1 As String, ByVal strRow2 As String, ByVal strRow3 As String, ByVal strRow4 As String)
'---------------------------------------------------------------------
'�����:�ܺ�ȫ
'   1999-8-22
'---------------------------------------------------------------------
'���ܣ����ݴ��������ֵ������ʾ��LED��
'������strRow1-strRow4:������Ϣ(30���ַ�,15������)
'���أ�
'---------------------------------------------------------------------
    On Error Resume Next
    With ctlComm
        .output = Chr(27) + "@"
        .output = Chr(27) + "CLR"
        
        .output = Chr(27) + "l" + Chr(1) + Chr(1)
        .output = strRow1
        .output = Chr(27) + "l" + Chr(1) + Chr(2)
        .output = strRow2
        .output = Chr(27) + "l" + Chr(1) + Chr(3)
        .output = strRow3
        If gblnHaveBottom = 1 Then
            .output = Chr(27) + "l" + Chr(1) + Chr(4)
            .output = strRow4
        End If
    End With
End Sub

Public Function SetLength(ByVal strText As String, ByVal lngLen As Long) As String
'���ܣ������ַ�������󳤶�
'������lngLen=���ֽ�Ϊ��λ����󳤶�
    Dim strTmp As String, i As Long
    
    If zlCommFun.ActualLen(strText) <= lngLen Then
        SetLength = strText
    Else
        For i = 1 To Len(strText)
            If zlCommFun.ActualLen(strTmp & Mid(strText, i, 1)) <= lngLen Then
                strTmp = strTmp & Mid(strText, i, 1)
            End If
        Next
        SetLength = strTmp
    End If
End Function
