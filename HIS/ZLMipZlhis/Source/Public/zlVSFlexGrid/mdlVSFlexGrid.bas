Attribute VB_Name = "mdlVSFlexGrid"
Option Explicit

'######################################################################################################################
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Enum Color
    ��ɫ = &H80000005
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E4E7
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
    
    ԭʼ���� = 0
    ������¼ = &HFF
    ͣ����Ŀ = &H8000000C
    ������Ŀ = 0
    
    ����ģ��ɫ = &HC00000
    
    ��������ɫ = &H40C0&
    ����ǰ��ɫ = &H8000000E
    ���걳��ɫ = &H80C0FF
    �ͱ걳��ɫ = &H80FFFF
    ����ǰ��ɫ = &H80000012
    Ĭ��ǰ��ɫ = &H80000008
    ��ɫ = &HF5F5F5
    ����ɫ = 0
    ͣ��ɫ = 255
End Enum


Private Type POINTAPI
     X As Long
     Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Const SB_TOP = 6
Public Const WM_VSCROLL = &H115

Private mlngTXTProc As Long

Public Sub PressKey(bytKey As Byte)
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Sub TxtSelAll(objTxt As Object)
'���ܣ����༭��ĵ��ı�ȫ��ѡ��
'������objTxt=��Ҫȫѡ�ı༭�ؼ�,�ÿؼ�����SelStart,SelLength����
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    If TypeName(objTxt) = "TextBox" Then
        If objTxt.MultiLine Then
            SendMessage objTxt.hwnd, WM_VSCROLL, SB_TOP, 0
        End If
    End If
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function


Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByVal hwnd As Long = 0, Optional str��Ŀ As String) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If str��Ŀ = "" Then str��Ŀ = "����������"
    
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        MsgBox str��Ŀ & "���зǷ��ַ���", vbExclamation, ""
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox str��Ŀ & "���ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, ""
            If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
            Exit Function
        End If
    End If
    
    StrIsValid = True
End Function

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '������
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '��С��
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function SetNewWindowLong(ByVal lngHwnd As Long, ByVal dwNewLong As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mlngTXTProc = GetWindowLong(lngHwnd, GWL_WNDPROC)
    Call SetWindowLong(lngHwnd, GWL_WNDPROC, dwNewLong)
        
    SetNewWindowLong = True
    
End Function

Public Function RestoreWindowLong(ByVal lngHwnd As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Call SetWindowLong(lngHwnd, GWL_WNDPROC, mlngTXTProc)
    
    RestoreWindowLong = True
End Function

Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    '******************************************************************************************************************
    '���ܣ�ȥ��TextBox��Ĭ���Ҽ��˵�
    '������
    '���أ�
    '˵���������Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    '******************************************************************************************************************
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(mlngTXTProc, hwnd, msg, wp, lp)
End Function

Public Function ErrCenter() As Byte
    MsgBox Err.Description
End Function

