Attribute VB_Name = "mdlPublic"
Option Explicit 'Ҫ���������
Public gclsInsure As New clsInsure          'ҽ���ӿڶ���
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrPrivsStation As String '��ǰ�û���ҽ������վ��Ȩ��  ֻ��ͨ���ӿڵ���ʱ,�Ŵ���
Public gstrSysName As String                'ϵͳ����
Public gstrUnitName As String
Public glngSys As Long
Public glngModul As Long
Public gstrProductName As String

Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"
Public gbytDec As Byte '���ý���С����λ��
Public gbyt���������Ϣ As Byte '0-�����;1-���;2-��ʾ���
Public gblnOk As Boolean
Public gstrDBUser As String '��ǰ�û���
Public gfrmMain As Object
'�û���Ϣ------------------------
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    �������� As String
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

'ϵͳ����
Public Type TY_Reg_Para  '�Һ���ز���
    bytNODaysGeneral As Byte    '��ͨ�Һ���Ч����
    bytNoDayseMergency As Byte '����Һ���Ч����
End Type
Public Type TY_SysPara
    Sy_Reg  As TY_Reg_Para
End Type
Public gSysPara As TY_SysPara       'ϵͳ�������;�Ժ������չ(���˺�)

Public gstrLike As String   '����ƥ�䷽ʽ
Public glngInterval As Long '�ҺŰ��ű��Զ�ˢ�¼��,0��ʾ���Զ�ˢ��

'Public gblnShowCard As Boolean '�Ƿ�������ʾ����

Public gblnSharedInvoice As Boolean '�Һ�ʹ���շ�Ʊ��
Public gblnBill�Һ� As Boolean '�Ƿ��ϸ����Ʊ��

Public gbytFactLength As Byte '�Һ�Ʊ�ݺ��볤��
Public glng�Һ�ID As Long '�Һ�����ID
Public gdblԤ��������鿨 As Double 'Ԥ�������ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
Public gbytԤ����˷��鿨 As Byte 'Ԥ����˷�ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
Public gbln���ѿ��˷��鿨 As Boolean '���ѿ��˷�ʱ�Ƿ�ˢ����֤
Public gbln������� As Boolean

Public gstr�ſ�ID As String  '���￨����ID
'Public gblnBill�ſ� As Boolean '�Ƿ��ϸ����Ʊ��
'Public gbyt�ſ� As Byte '���￨�ų���
Public gstrCardPass As String 'ˢ��ʱҪ����������,'0000000000'��λ˳���ʾ��������,�ֱ�Ϊ:1.����Һ�,2.���ﻮ��,3.�����շ�,4.�������,5.��Ժ�Ǽ�,6.סԺ����,7.���˽���,8.����Ԥ����,9.���鼼ʦվ,10.Ӱ��ҽ��վ.'
Public gblnPrePayPriority As Boolean '����ʹ��Ԥ����

Public gintԤԼ���� As Integer '�Һ������ԤԼ����
Public gstr�ϰ�ʱ�� As String

Public gstr�Һſ���ID As String   '������վ����ҺŵĿ���ID
Public gstrIme As String '�Զ����������뷨

Public gbytRegistMode As Byte '�Һ�ģʽ
Public gdatRegistTime As Date '�����ģʽ����ʱ��

Public Type TY_VisitPlan_ModulePara '�ٴ����ﰲ��ģ�����
    byt������ӡ��ʽ As Byte
    str��Դά��վ�� As String 'δ����վ��Ŀ��Һ�Դ��ά��վ��
    byt����ȽϷ�ʽ  As Byte '��Դ���밴���ֱȽϷ�ʽ��������0-���ַ��Ƚϣ�1-����ֵ�Ƚ�
End Type
Public gVisitPlan_ModulePara As TY_VisitPlan_ModulePara

'��ѡ������Ŀ
Public gbln���� As Boolean '����
Public gbln�Ա� As Boolean  '�Ա�
Public gbln���� As Boolean  '����
Public gbln��ͥ��ַ As Boolean  '��ͥ��ַ
Public gbln���ʽ As Boolean  '���ʽ
Public gbln�ѱ� As Boolean '�ѱ�
Public gbln���㷽ʽ As Boolean '���㷽ʽ
Public gblnҽ�� As Boolean 'ҽ��
Public gbln�绰 As Boolean

'ȱʡֵ
Public gstr���ʽ As String 'ȱʡ���ʽ
Public gstr�ѱ� As String 'ȱʡ�ѱ�
Public gstr�Ա� As String 'ȱʡ�Ա�
Public gstr���㷽ʽ As String 'ȱʡ���㷽ʽ
'���˺� ����:????    ����:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '����С������
Public gstrFeePrecisionFmt As String '����С����ʽ:0.00000

'��������
Public gbln�ɿ���� As Boolean
Public gbln�Զ������ As Boolean
Public gblnAutoAddName As Boolean '����ʱ�Զ�������ʱ����
Public gblnNewCardNoPop As Boolean '����ʱ����������������
Public gbln���ѽ����� As Boolean
Public gbln�˷��ش� As Boolean '�˺Ų��˿�ʱ�Ƿ��ش�Ʊ
Public gint�ų� As Integer '�ű𳤶�
Public gblnLED As Boolean
Public gblnPrintFree As Boolean
Public gblnPrintCase As Boolean '��ӡ������ǩ
Public gbytInvoice As Byte   '��Ʊ��ӡ��ʽ
Public gByt��ӡ�������� As Byte '�������� ��ӡ��ʽ
Public gblnPrice As Boolean     '�������˹ҺŴ�Ϊ���۵�
Public gintNameDays As Integer  '������������N���ڵĲ���
Public gblnSeekName As Boolean

Public glngOld As Long
Public glngMinW As Long, glngMaxW As Long
Public glngMinH As Long, glngMaxH As Long
Public gbln���֤Ψһ As Boolean
'WIN32����

'API����
Public Const CB_ADDSTRING = &H143
Public Const CB_FINDSTRING = &H14C
Public Const CB_SHOWDROPDOWN = &H14F

Public Const TVM_SETBKCOLOR = 4381&
Public Const TVM_GETBKCOLOR = 4383&
Public Const TVS_HASLINES = 2&
'�ؼ�����λ�û�ȡת��
Public Const EM_EXGETSEL = (&H400 + 52)
Public Const EM_POSFROMCHAR = &HD6

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Public Enum mTextAlign
    taLeftAlign = 0
    taCenterAlign = 1
    taRightAlign = 2
End Enum

Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 'ǳ����
Public Const BDR_RAISEDINNER = &H4 'ǳ͹��
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000



Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_F4 = vbKeyF4

Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Dim p As KBDLLHOOKSTRUCT
Public p1 As KBDLLHOOKSTRUCT
Public gblnBegin As Boolean
Public gblnLen As Boolean
Public gblnCard As Boolean
Public gsngStartTime As Single

Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Type POINTAPI
     X As Long
     Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24

Public Const GWL_STYLE = (-16)              'Set the window style
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000        '��߿�
Public Const WS_SYSMENU = &H80000           '�ڱ������Ƿ�߱�ϵͳ�˵�
Public Const WS_MINIMIZEBOX = &H20000       '�߱���С����ť
Public Const WS_MAXIMIZEBOX = &H10000       '�߱���󻯰�ť
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_RESETCONTENT = &H14B

'�ƶ��ؼ����ޱ߿���
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal Hwnd As Long) As Long
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MOVE = &HF010&
Public Const HTCAPTION = 2

'�л���ָ�������뷨��
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long

'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long

'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'''''''''''''''''''''
'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Type Ty_CardProperty
       lng�����ID      As Long
       str������        As String
       str������        As String
       lng���ų���      As Long
       lng���㷽ʽ      As String
       bln���ƿ�        As Boolean
       bln�ϸ����      As Boolean
       lng����ID        As Long
       lng��������      As Long
       bln���          As Boolean
       int���볤��      As Integer
       int���볤������  As Integer
       int�������      As Integer
       bln���￨        As Boolean
       str��������      As String
       str��׼��Ŀ      As String
       blnȱʡ��־      As Boolean
       blnOneCard       As Boolean '  '�Ƿ�������һ��ͨ�ӿ�,��ģʽ�£�Ʊ���ϸ����Ʊ�ŷ�Χ��ķ�����󶨿����շ�
       rs����           As ADODB.Recordset
       dblӦ�ս��      As Double
       dblʵ�ս��      As Double
       bln�Ƿ��ƿ�      As Boolean
       bln�Ƿ񷢿�      As Boolean
       bln�Ƿ�д��      As Boolean
       lng��������      As Long '0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�����������ſ���������ʾ;ȱʡΪ0 �����:57326
       bln�ظ�ʹ��      As Boolean
       str��������      As String
       byt��������      As Byte
       lng�շ�ϸĿID    As Long 'ҽԺ����������ѷ��ص��շ�ϸĿID,���뵱ǰ���ѵ��շ�ϸĿIDͬ��
End Type
Public gCurSendCard As Ty_CardProperty
Public gstrSQL  As String
Public glngMax��ͥ��ַ As Long       '��ͥ��ַ�������¼�볤��
Public glngMax���ڵ�ַ As Long       '���ڵ�ַ�������¼�볤��
Public glngMax�����ص� As Long       '�����ص��������¼�볤��
Public glngMax��ϵ�˵�ַ As Long    '��ϵ�˵�ַ�������¼�볤��

Public Function WndMessage(ByVal Hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
'���ܣ�ȥ��TextBox��Ĭ���Ҽ��˵�
    If msg <> WM_CONTEXTMENU Then
        WndMessage = CallWindowProc(glngTXTProc, Hwnd, msg, wp, lp)
    End If
End Function

Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������lngHwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�-2=δ�Ӵ���,����=ƥ�������(����ƥ�������)
'˵�����뽫�ú�����KeyPress�¼��е��á�

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '������(ȱʡΪ0.5��)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 'ʹComboBox����ĵ���ƥ�书��ʧЧ
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '������Իس���������
    End If
End Function

Public Function FindName(cbo As ComboBox) As String
'���ܣ�ȡ����ǰComboBox��ֵ(�����Ϊ�����-���ơ�)
'˵������ҪΪSQL���ʹ��
    If cbo.ListIndex = -1 Then
        FindName = "Null"
    Else
        FindName = "'" & Mid(cbo.Text, InStr(1, cbo.Text, "-") + 1) & "'"
    End If
End Function

Public Function FindText(txt As TextBox) As String
'���ܣ�����ǰTextBox��ֵת��Ϊ��׼SQL���
'˵������ҪΪSQL���ʹ��
    If Len(Trim(txt.Text)) = 0 Then
        FindText = "Null"
    Else
        FindText = "'" & txt.Text & "'"
    End If
End Function

Public Function NeedName(strList As String, Optional ByVal blnLast As Boolean = False, _
Optional strSplit As String = "-") As String
    If Not blnLast Then
        NeedName = Mid(strList, InStr(strList, strSplit) + 1)
    Else
        NeedName = strList
        Do While (InStr(NeedName, strSplit)) > 0
            NeedName = Mid(NeedName, InStr(NeedName, strSplit) + 1)
        Loop
    End If
End Function

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim fEatKeystroke As Boolean
    Dim sngTime As Single
    Dim sngPreTime As Timer
    
    gblnCard = False
    
    sngTime = Timer
    If (nCode = HC_ACTION) Then
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            
            CopyMemory p, ByVal lParam, Len(p)
            gblnCard = (sngTime - gsngStartTime) < 0.6
            If gblnCard = False Then gblnLen = False
             
            gsngStartTime = sngTime
            fEatKeystroke = _
            ((p.vkCode = VK_TAB) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
            ((p.vkCode = VK_ESCAPE) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
            ((p.vkCode = VK_ESCAPE) And ((GetKeyState(VK_CONTROL) And &H8000) <> 0)) Or _
            ((p.vkCode = 91) Or (p.vkCode = 92) Or (p.vkCode = 93)) Or _
            ((p.vkCode = VK_F4) And (p.flags And LLKHF_ALTDOWN) <> 0) '�������д�������Alt+F4
            If p.vkCode = Asc(";") Then fEatKeystroke = True
        End If
        
        If p.vkCode = vbKeyBack Then
            LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
            Exit Function
        End If
    End If
    If (fEatKeystroke Or gblnLen) Then
        LowLevelKeyboardProc = -1
    Else
        LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
    End If
End Function

Public Function NeedCode(strList As String) As String
    If InStr(strList, "-") = 0 Then NeedCode = strList: Exit Function
    NeedCode = Mid(strList, 1, InStr(strList, "-") - 1)
End Function
Public Function Custom_WndMessage(ByVal Hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngMinW \ 15
        MinMax.ptMinTrackSize.Y = glngMinH \ 15
        MinMax.ptMaxTrackSize.X = glngMaxW \ 15
        MinMax.ptMaxTrackSize.Y = glngMaxH \ 15
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, Hwnd, msg, wp, lp)
End Function

Public Function SetCboDefault(cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = 1 Then
            cbo.ListIndex = i
            SetCboDefault = i: Exit Function
        End If
    Next
    If cbo.ListCount > 0 And cbo.ListIndex = -1 Then cbo.ListIndex = 0
End Function

Public Sub CheckLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function OpenIme(Optional strIme As String) As Boolean
'����:�����ƴ��������뷨,��ָ������ʱ�ر��������뷨��֧�ֲ������ơ�
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    
    If strIme = "���Զ�����" Then OpenIme = True: Exit Function
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function


Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function GetBaseDict() As ADODB.Recordset
'���ܣ����ֵ��ж�ȡ����
    Dim strSQL As String, strTmp As String, arrTmp As Variant, i As Integer
    strTmp = "����,����,����״��,ְҵ,����ϵ"
    arrTmp = Split(strTmp, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = arrTmp(i)
        If strSQL = "" Then
            strSQL = "Select '" & strTmp & "' ���,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strTmp
        Else
            strSQL = strSQL & " Union all Select '" & strTmp & "' ���,����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strTmp
        End If
    Next
    strSQL = strSQL & " Order by ���,����"
    
    On Error GoTo errH
    Set GetBaseDict = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����,����,����״��,ְҵ,����ϵ")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlInitMEPIPati(ByRef rsPati As ADODB.Recordset) As Boolean
    Set rsPati = New ADODB.Recordset
    With rsPati
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "����ID", adBigInt, , adFldIsNullable
            .Append "��ҳID", adBigInt, , adFldIsNullable
            .Append "�Һ�ID", adBigInt, , adFldIsNullable
            .Append "�����", adVarChar, 18, adFldIsNullable
            .Append "סԺ��", adVarChar, 18, adFldIsNullable
            .Append "ҽ����", adVarChar, 30, adFldIsNullable
            .Append "���֤��", adVarChar, 18, adFldIsNullable
            .Append "����֤��", adVarChar, 20, adFldIsNullable
            .Append "����", adVarChar, 100, adFldIsNullable
            .Append "�Ա�", adVarChar, 4, adFldIsNullable
            .Append "��������", adVarChar, 20, adFldIsNullable
            .Append "�����ص�", adVarChar, 100, adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����", adVarChar, 20, adFldIsNullable
            .Append "ѧ��", adVarChar, 10, adFldIsNullable
            .Append "ְҵ", adVarChar, 80, adFldIsNullable
            .Append "������λ", adVarChar, 100, adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����״��", adVarChar, 4, adFldIsNullable
            .Append "��ͥ�绰", adVarChar, 20, adFldIsNullable
            .Append "��ϵ�˵绰", adVarChar, 20, adFldIsNullable
            .Append "��λ�绰", adVarChar, 20, adFldIsNullable
            .Append "��ͥ��ַ", adVarChar, 100, adFldIsNullable
            .Append "��ͥ��ַ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "���ڵ�ַ", adVarChar, 100, adFldIsNullable
            .Append "���ڵ�ַ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "��λ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "��ϵ�˵�ַ", adVarChar, 100, adFldIsNullable
            .Append "��ϵ�˹�ϵ", adVarChar, 30, adFldIsNullable
            .Append "��ϵ������", adVarChar, 64, adFldIsNullable
        End With
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    zlInitMEPIPati = True
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub
