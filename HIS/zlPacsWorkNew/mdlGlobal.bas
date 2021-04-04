Attribute VB_Name = "mdlGlobal"
Option Explicit
 

Private mstrInstitution As String
Private mstrSysRootPath As String
Private mobjDcmGlobal As DicomGlobal

Private hHook As Long
Private hFormhWnd As Long

'ȫ�����Է���

'ע��ĵ�λ����
Property Get RegInstitution() As String
    If Len(mstrInstitution) <= 0 Then
        mstrInstitution = zlRegInfo("��λ����")
        
        If Len(mstrInstitution) <= 0 Then mstrInstitution = "δע��"
    End If
    
    RegInstitution = mstrInstitution
End Property

'ϵͳ·��
Property Get SysRootPath() As String
    If Len(mstrSysRootPath) <= 0 Then mstrSysRootPath = GetAppRootPath
    
    SysRootPath = mstrSysRootPath
End Property

Property Let SysRootPath(value As String)
    mstrSysRootPath = value
End Property


'��ȡ˽��ע���·��
Public Function GetPrivateRegPath(ByVal strItemName As String) As String
    GetPrivateRegPath = "˽��ģ��\" & UserInfo.�û��� & "\" & App.EXEName & "\��������\" & strItemName
End Function

'��ȡ����ע���·������
Public Function GetPublicRegPath(ByVal strItemName As String) As String
    GetPublicRegPath = "����ģ��\" & App.EXEName & "\" & strItemName
End Function



'����UID
Public Function CreateUID() As String
    If mobjDcmGlobal Is Nothing Then
        Set mobjDcmGlobal = New DicomGlobal
        mobjDcmGlobal.RegString("UIDRoot") = "1"
    End If
    
    CreateUID = mobjDcmGlobal.NewUID
End Function


'�������Ƿ����У�exeName ������Ҫ���Ľ��� exe ���֣����� VB6.EXE
Public Function CheckExeIsRun(ByVal strExeName As String) As Boolean
    Dim objWMIService As Object
    Dim colProcessList As Object
    
On Error Resume Next

    CheckExeIsRun = False
    
    Set objWMIService = VBA.GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select name from Win32_Process Where Name='" & strExeName & "'")
    
    CheckExeIsRun = IIf(colProcessList.Count > 0, True, False)
    
    Set colProcessList = Nothing
    Set objWMIService = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
''Ϊ�˴���˫��ʱ�Ի������ȷ��ʾλ�ã���API������д��һ��MsgBox����
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgBoxD(objParent As Object, ByVal strPrompt As String, Optional ByVal dwStyle As VbMsgBoxStyle = MB_OK, Optional strTitle As String = "") As Long

    Dim lngHwnd As Long
 
    If objParent Is Nothing Then
        lngHwnd = GetActiveWindow
    Else
        lngHwnd = objParent.hwnd
    End If

    If lngHwnd = GetDesktopWindow Or lngHwnd = 0 Then
        lngHwnd = GetForegroundWindow
    End If
 

    MsgBoxD = MsgboxH(lngHwnd, strPrompt, dwStyle, strTitle)

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''
'���VB�е�Msgbox����
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgboxH(hwnd As Long, strPrompt As String, Optional ByVal dwStyle As VbMsgBoxStyle = MB_OK, Optional strTitle As String) As Long

    Dim hInstance As Long
    Dim hThreadId As Long

    hInstance = App.hInstance
    hThreadId = App.ThreadID

    If dwStyle = 0 Then dwStyle = vbOKOnly
    If Len(strTitle) = 0 Then strTitle = App.EXEName

    '����ǰ���ڵľ����������
    hFormhWnd = hwnd

    '���ù���
    hHook = SetWindowsHookEx(WH_CBT, AddressOf BoxPro, hInstance, hThreadId)
    
    '����MessageBox API
    MsgboxH = MessageBox(hwnd, strPrompt, strTitle, dwStyle)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''
'HOOK����
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BoxPro(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    '��������
    Dim rc As RECT
    Dim rcFrm As RECT

    Dim newLeft As Long
    Dim newTop As Long
    Dim dlgWidth As Long
    Dim dlgHeight As Long
    Dim scrWidth As Long
    Dim scrHeight As Long
    Dim frmLeft As Long
    Dim frmTop As Long
    Dim frmWidth As Long
    Dim frmHeight As Long
    Dim hwndMsgBox As Long

    '��MessageBox����ʱ����Msgbox�Ի�����������ڵĴ���
    If nCode = HCBT_ACTIVATE Then
    
        '��ϢΪHCBT_ACTIVATEʱ������wParam��������MessageBox�ľ��
        hwndMsgBox = wParam
        
        '�õ�MessageBox�Ի����Rect
        Call GetWindowRect(hwndMsgBox, rc)
        Call GetWindowRect(hFormhWnd, rcFrm)
        
        If rcFrm.Right = 0 Or rcFrm.Bottom = 0 Then
            Call GetWindowRect(GetDesktopWindow, rcFrm)
        End If
        
        'ʹMessageBox����
        frmLeft = rcFrm.Left
        frmTop = rcFrm.Top
        frmWidth = rcFrm.Right - rcFrm.Left
        frmHeight = rcFrm.Bottom - rcFrm.Top
        dlgWidth = rc.Right - rc.Left
        dlgHeight = rc.Bottom - rc.Top
    
        scrWidth = Screen.Width \ Screen.TwipsPerPixelX
        scrHeight = Screen.Height \ Screen.TwipsPerPixelY
    
        newLeft = frmLeft + ((frmWidth - dlgWidth) \ 2)
        newTop = frmTop + ((frmHeight - dlgHeight) \ 2)
        
'        '�޸�ȷ����ť������
'        Call SetDlgItemText(hwndMsgBox, IDOK, "����ȷ����ť")
        SetWindowPos hwndMsgBox, -1, rcFrm.Left, rcFrm.Top, dlgWidth, dlgHeight, 3 '�������ö�
        SetForegroundWindow hwndMsgBox
        
        'Msgbox����
        Call MoveWindow(hwndMsgBox, newLeft, newTop, dlgWidth, dlgHeight, True)
        
        'ж�ع���
        UnhookWindowsHookEx hHook
    End If
    
    BoxPro = False
End Function


Public Function MainForm() As Object
    Dim objForm As Object
    
    Set MainForm = Nothing
    
    If Forms.Count <= 0 Then Exit Function
    
    For Each objForm In Forms
        If InStr(objForm.Name, "PacsMain") > 0 Then
            Set MainForm = objForm
            Exit Function
        End If
    Next
    
    Set MainForm = Forms(0)
End Function
