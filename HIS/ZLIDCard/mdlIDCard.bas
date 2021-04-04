Attribute VB_Name = "mdlIDCard"
Option Explicit 'Ҫ���������
Public gblnIDCard As Boolean    '�豸�Ƿ����
Public glngType As Long         '�豸���ͱ��

Public editPort As Long         '��¼ͨ�ýӿڵĶ˿ں�
Public blnUsbPort As Boolean    ' ��¼ͨ�ýӿڵĶ˿�����

'Ŀǰ����,�����豸��ʹ����ͬ�Ľӿ��ļ�termb.dll,��������ͬ��,��������ֻ֧�������Ͽɵ��豸,��������δ��������Ȩ���豸�̽���ϵͳ
Public Enum IDCardType
    CVR100U = 1                 '���ڻ���CVR-100U
    CVR100D = 2
    SS_V1 = 3                   '��˼
    XZX_KDQ = 4                 '������
    GTICR100 = 5                '�ɶ�����ʵҵ�������޹�˾ GTICR100
    HX_FDX9 = 6                 '����𿨹ɷ����޹�˾
    GTICR100_1 = 7              '�ɶ�����ʵҵ����         GTICR100������Ʒ����
    CVR100U_1 = 8               '���ڻ���CVR-100U������Ʒ����
    CVR100D_1 = 9               '���ڻ���CVR-100D������Ʒ����
    DKQ_116D = 10               '������DKQ116D
    GTICR100_01 = 11            '�ɶ�����ʵҵ�������޹�˾ GTICR100-01
    CVR100 = 12                 '���ڻ���CVR-100
    COMMON = 13                 'ͨ�ýӿ�
    SS728M01_B01C = 14          '��˼SS72801 ���B01C
End Enum

Public Type MIDCardInfor    '����ṹ��������������֤��Ϣ
    Name As String * 32  '����
    sex As String * 4  '�Ա�
    nation As String * 6  '����
    born As String * 18  '��������
    address As String * 72  'סַ
    IDcardno As String * 38  '���֤��
    grantdept As String * 32  '��֤����
    UserLifeBegin As String * 18  '��Ч��ʼ����
    UserLifeEnd As String * 18  '��Ч��ֹ����
    reserved As String * 38  '����
    PhotoFileName As String * 255  '��Ƭ·��
End Type

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public IDCardInfor As MIDCardInfor

Public Declare Function InitComm Lib "termb.dll" (ByVal Port As Integer) As Integer
Public Declare Function InitCommExt Lib "termb.dll" () As Integer
Public Declare Function Authenticate Lib "termb.dll" () As Integer
Public Declare Function AuthenticateExt Lib "termb.dll" () As Integer
Public Declare Function Read_Content_Path Lib "termb.dll" (ByVal fileName As String, ByVal Active As Integer) As Integer
Public Declare Function Read_Content Lib "termb.dll" (ByVal Active As Integer) As Integer
Public Declare Function CloseComm Lib "termb.dll" () As Integer
Public Declare Function GetSAMID Lib "termb.dll" () As String

Public Declare Function SDT_GetCOMBaud Lib "sdtapi.dll" (ByVal iPort As Integer, puiBaudRate As Long) As Integer
Public Declare Function SDT_SetCOMBaud Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal uiCurrBaud As Long, ByVal uiSetBaud) As Integer
Public Declare Function SDT_OpenPort Lib "sdtapi.dll" (ByVal iPort As Integer) As Integer
Public Declare Function SDT_ClosePort Lib "sdtapi.dll" (ByVal iPort As Integer) As Integer

Public Declare Function SDT_ResetSAM Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_SetMaxRFByte Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal ucByte As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_GetSAMStatus Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_GetSAMID Lib "sdtapi.dll" (ByVal iPort As Integer, pucSAMID As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_GetSAMIDToStr Lib "sdtapi.dll" (ByVal iPort As Integer, pcSAMID As String, ByVal iIfOpen As Integer) As Integer

Public Declare Function SDT_StartFindIDCard Lib "sdtapi.dll" (ByVal iPort As Integer, CardPUCIIN As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_SelectIDCard Lib "sdtapi.dll" (ByVal iPort As Integer, pucManaMsg As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_ReadMngInfo Lib "sdtapi.dll" (ByVal iPort As Integer, pucManageMsg As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_ReadBaseMsg Lib "sdtapi.dll" (ByVal iPort As Integer, pucCHMsg As String, puiCHMsgLen As Long, _
         pucPHMsg As String, puiPHMsgLen As Long, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_ReadBaseMsgToFile Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal pucCHFile As String, puiCHMsgLen As Long, _
         ByVal pucPHFile As String, puiPHMsgLen As Long, ByVal iIfOpen As Integer) As Integer


Public Declare Function GetBmp Lib "WltRS.dll" (ByVal WLTFile As String, ByVal intf As Integer) As Integer

'���ڻ���CVR-100U������Ʒ������̬�⺯��
Public Declare Function CVR_InitComm Lib "termb.dll" (ByVal Port As Long) As Integer
Public Declare Function CVR_CloseComm Lib "termb.dll" () As Integer
Public Declare Function CVR_Authenticate Lib "termb.dll" () As Integer
Public Declare Function CVR_Read_Content Lib "termb.dll" (ByVal Active As Long) As Integer
Public Declare Function CVR_Ant Lib "termb.dll" (ByVal mode As Long) As Integer

Public Declare Function GetPeopleName Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
Public Declare Function GetPeopleAddress Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
Public Declare Function GetPeopleIDCode Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'������DKQ����API
'********************** �˿���API *************************
Public Declare Function Syn_OpenPort Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer) As Integer
Public Declare Function Syn_ClosePort Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer) As Integer

'********************** SAM��API **************************
Public Declare Function Syn_GetSAMStatus Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer, ByVal iIfOpen As Integer) As Integer
Public Declare Function Syn_ResetSAM Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer, ByVal iIfOpen As Integer) As Integer

'******************* ���֤����API ************************
Public Declare Function Syn_StartFindIDCard Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer, pucManaInfo As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function Syn_SelectIDCard Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer, pucManaMsg As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function Syn_ReadMsg Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer, ByVal iIfOpen As Integer, pIDCardData As MIDCardInfor) As Integer
Public Declare Function Syn_ReadCard Lib "Syn_IDCardRead.dll" (pIDCardData As MIDCardInfor, ByVal Rmode As Integer) As Integer
'Rmode ������1��Ϊ����������2��Ϊһ��һ�ζ�����
'******************* ������API ************************
Public Declare Function Syn_SendSound Lib "Syn_IDCardRead.dll" (ByVal iCmdNo As Integer) As Integer '˵��: ��������
Public Declare Sub Syn_DelPhotoFileLib Lib "Syn_IDCardRead.dll" () '˵��: ɾ����ʱ��Ƭ�ļ�

Public Declare Function SendMessage Lib "user32" _
            Alias "SendMessageA" (ByVal hwnd As Long, _
            ByVal wMsg As Long, ByVal wParam As Long, _
            lParam As Any) As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Function GetPID() As Long
   GetWindowThreadProcessId GetForegroundWindow, GetPID
End Function
Public Function Init(ByVal objIDCard As clsIDCard) As Boolean
    Dim i As Long, lngTmp As Long
    Dim bytReadType     As Byte
    Dim strPort         As String
    Dim strServerIP     As String
    Dim lngReturn       As Long
    Dim pszVerInfo      As String
    Dim szDevCom        As String
    Dim icdev           As Long
    Dim szDevIP         As String
    Dim Amount          As Long
    Dim Msec            As Long
    Dim pszParam        As String
    Dim strSFZH         As String
    Dim pszText         As String
On Error GoTo ErrH

    On Error GoTo ErrH
    gblnIDCard = Val(GetSetting("ZLSOFT", "����ȫ��\IDCard", "����", 0)) = 1
    If gblnIDCard Then

        glngType = Val(GetSetting("ZLSOFT", "����ȫ��\IDCard", "�豸����", 0))
        Select Case glngType
            Case IDCardType.CVR100U, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100, IDCardType.GTICR100_1, IDCardType.GTICR100_01

                i = InitComm(1001) '��ʼ���˿�
                If i <> 1 Then i = InitCommExt
                gblnIDCard = (i = 1)
            Case IDCardType.CVR100U_1, IDCardType.CVR100D_1
                i = CVR_InitComm(1001)
                If i <> 1 Then i = CVR_InitComm(1)
                gblnIDCard = (i = 1)
            Case IDCardType.CVR100D, IDCardType.HX_FDX9
                i = InitComm(1)
                If i <> 1 Then i = InitCommExt
                gblnIDCard = (i = 1)

                If gblnIDCard = False Then
                    i = SDT_OpenPort(1)
                    gblnIDCard = (i = CByte(&H90))
                End If
            Case IDCardType.DKQ_116D
                i = Syn_OpenPort(1001)
                gblnIDCard = (i = 0)
                Call Syn_ClosePort(1001)   '��ֹ�ര�ڳ�ʼ��ʧ��
            Case IDCardType.CVR100
                i = SDT_OpenPort(1001)
                If i <> CByte(&H90) Then i = SDT_OpenPort(1)
                gblnIDCard = (i = CByte(&H90))

            Case IDCardType.COMMON
                gblnIDCard = False
                '���usb�ڵĻ������ӣ������ȼ��usb
                Dim iPort As Long
                For iPort = 1001 To 1016
                    i = SDT_OpenPort(iPort)
                    If i = CByte(&H90) Then
                        editPort = iPort
                        blnUsbPort = True
                        gblnIDCard = True
                        Exit For
                    End If
                Next

                '��⴮�ڵĻ�������
                If Not blnUsbPort Then
                    For iPort = 1 To 16
                        i = SDT_OpenPort(iPort)
                        If i = CByte(&H90) Then
                            editPort = iPort
                            blnUsbPort = False
                            gblnIDCard = True
                            Exit For
                        End If
                    Next
                End If
            Case IDCardType.SS728M01_B01C
                bytReadType = IIf(UCase(GetSetting("ZLSOFT", "����ȫ��\IDCard\SS728M01", "LinkType", "Local")) = "NET", 2, 1)
                strPort = UCase(GetSetting("ZLSOFT", "����ȫ��\IDCard\SS728M01", "Port", "AUTO"))
                strServerIP = GetSetting("ZLSOFT", "����ȫ��\IDCard\SS728M01", "NetIp", "192.168.31.169")

                '����豸�����Ƿ�����
                '���豸
                If bytReadType = 1 Then
                    '2.2.2�򿪱����ն�
                    'long __stdcall ss_dev_open(char* szDevCom);
                    szDevCom = "AUTO"
                    lngReturn = ss_dev_open(szDevCom)
                    If lngReturn < 0 Then
                        If ss_error(lngReturn, False) Then
                            Exit Function
                        End If
                    End If
                ElseIf bytReadType = 2 Then
                    '2.2.4ע�������ն�
                    szDevIP = strServerIP
                    lngReturn = ss_dev_login(szDevIP)
                    If lngReturn < 0 Then
                        If ss_error(lngReturn, False) Then
                            Exit Function
                        End If
                    End If
                End If
                icdev = lngReturn
                '�ر��豸
                If bytReadType = 1 Then
                    '2.2.3�رձ����ն�
                    'long __stdcall ss_dev_close(long icdev);
                    lngReturn = ss_dev_close(icdev)
                    If ss_error(lngReturn, False) Then
                        Exit Function
                    End If
                Else
                    '2.2.5ע�������ն�
                    lngReturn = ss_dev_logout(icdev)
                    If ss_error(lngReturn, False) Then
                        Exit Function
                    End If
                End If
                gblnIDCard = True
        End Select

        If gblnIDCard Then
            Init = True
        Else
            MsgBox "���֤�豸�˿ڳ�ʼ��ʧ��,����!", vbInformation, "ZLSoft"
            Init = False
        End If
    End If
    Exit Function

ErrH:
    gblnIDCard = False
End Function

Public Sub Terminate()
    If gblnIDCard Then
        Select Case glngType
            Case IDCardType.CVR100U, IDCardType.CVR100D, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100, IDCardType.GTICR100_1, IDCardType.GTICR100_01
                CloseComm
            Case IDCardType.CVR100U_1, IDCardType.CVR100D_1
                CVR_CloseComm
            Case IDCardType.HX_FDX9
                Call SDT_ClosePort(1)
            Case IDCardType.DKQ_116D
                Call Syn_ClosePort(1001)
            Case IDCardType.CVR100
                Call SDT_ClosePort(1001)
            Case IDCardType.COMMON
                Call SDT_ClosePort(editPort)
            Case IDCardType.SS728M01_B01C
                Call ClosSS728M01
        End Select
    End If
End Sub

Public Sub SetAutoRead(ByVal timer As timer, blnEnabled As Boolean)
    If gblnIDCard Then
        Select Case glngType
            Case IDCardType.CVR100, IDCardType.CVR100U, IDCardType.CVR100D, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100, IDCardType.GTICR100_01, IDCardType.GTICR100_1, IDCardType.HX_FDX9, IDCardType.CVR100U_1, IDCardType.CVR100D_1, IDCardType.DKQ_116D, IDCardType.COMMON
                timer.Enabled = blnEnabled
            Case IDCardType.SS728M01_B01C
                timer.Enabled = blnEnabled
                '������ȡ�����ҿ�
                ReadCardSS728M01 = blnEnabled
        End Select
    End If
End Sub

Public Function TrimStr(ByVal str As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�
Dim Nstr As String
    If InStr(str, Chr(0)) > 0 Then
        Nstr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        Nstr = Trim(str)
    End If
    TrimStr = Replace(Replace(Replace(Nstr, Chr(13), vbCr), vbLf, ""), vbTab, "")
End Function

Public Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Sub ClearIDCardInfor()
    IDCardInfor.address = ""
    IDCardInfor.born = ""
    IDCardInfor.grantdept = ""
    IDCardInfor.IDcardno = ""
    IDCardInfor.Name = ""
    IDCardInfor.nation = ""
    IDCardInfor.PhotoFileName = ""
    IDCardInfor.reserved = ""
    IDCardInfor.sex = ""
    IDCardInfor.UserLifeBegin = ""
    IDCardInfor.UserLifeEnd = ""
End Sub

Public Function GetParentHwnd(ByVal ChildHwnd As Long) As Long
    Dim i As Long, lngHwnd As Long
    On Error GoTo errHand
    i = ChildHwnd

    Do While i > 0
        lngHwnd = i
        i = GetParent(i)
    Loop
    GetParentHwnd = lngHwnd
    Exit Function
errHand:
    Err.Clear
End Function


Public Sub PressKey(bytKey As Byte)
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub
