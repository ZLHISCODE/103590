Attribute VB_Name = "mdlKeyboard"
Option Explicit
Public gobjCom As MSComm
Public Type Ty_Com_Property
    int�˿ں� As Integer   '�˿ں�
    lng������ As Long
    str��ż����λ As String
    intֹͣλ As Integer
    int����λ As Integer
End Type
Public g_Com_Property As Ty_Com_Property
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public gblnStartKeyboard As Boolean '�Ƿ������������

Public Sub InitComProperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2011-07-28 14:46:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With g_Com_Property
        .int�˿ں� = Val(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "�˿�", 0)) + 1
        .int����λ = Val(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "����λ", "6"))
        .intֹͣλ = Val(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "ֹͣλ", "1"))
        .lng������ = Val(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "������", "9600"))
        .str��ż����λ = Trim(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "��ż����λ", "��"))
    End With
End Sub

Public Sub PressKey(bytKey As Byte)
    '���ܣ�����̷���һ����,����SendKey
    '������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub
