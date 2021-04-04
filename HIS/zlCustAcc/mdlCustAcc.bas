Attribute VB_Name = "mdlCustAcc"
Option Explicit

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type POINTAPI
     X As Long
     Y As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

'Windows���----------------------------------
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
'Public Const WS_CHILD = &H40000000
'Public Const WS_POPUP = &H80000000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'ϵͳ��������----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'--
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Declare Function GetROP2 Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const PS_DOT = 2
'Public Const PS_DASH = 1
Public Const R2_XORPEN = 7

Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24

Private Const VK_LCONTROL = &HA2
Private Const VK_RCONTROL = &HA3
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = LVM_FIRST + 30

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public gcnOracle As New ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼

Public gstrDbUser As String                 '��ǰ���ݿ��û�

Public gstr��λ���� As String
Public gstrSQL As String
Public gstrMenuSys As String                '��ǰ�û�ʹ�õĲ˵�ϵͳ
Public glngSys As Long
Public glngModul As Long

Public glngOldProc As Long                  'ԭ��Ϣ�������
'������ϵͳҪ�õ���ȫ�ֱ���
Public gfrmMain As Object                   '����̨���ڣ���Ҫ��������Ϣ�༭���ڵĸ�����
Public glngMain As Long                     'BH��������

Public Function CtrlIsPress() As Boolean
'���ܣ��жϵ�ǰ��Ctrl���Ƿ���
    If (GetKeyState(VK_LCONTROL) And &H80) <> 0 Or (GetKeyState(VK_RCONTROL) And &H80) <> 0 Then
        CtrlIsPress = True
    End If
End Function

Public Sub SetFont(objTarget As Object, objSource As Object)
'����:��һ��������������ȫ��������һ������
    With objSource.Font
        objTarget.Font.Name = .Name
        objTarget.Font.Size = .Size
        objTarget.Font.Bold = .Bold
        objTarget.Font.Italic = .Italic
        objTarget.Font.Underline = .Underline
    End With
End Sub


Public Function wndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = 641
        MinMax.ptMinTrackSize.Y = 422
        MinMax.ptMaxTrackSize.X = Screen.Width / 15
        MinMax.ptMaxTrackSize.Y = Screen.Height / 15
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        wndProc = 1
        Exit Function
    End If
    wndProc = CallWindowProc(glngOldProc, hwnd, Msg, wp, lp)
End Function

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function VsfGetColNum(vsf As VSFlexGrid, strColName As String) As Long
'����:������������vsfFlexGrid�ؼ��е������,û���ҵ�ʱ����-1(ʹ��vsfFee.ColIndex������Ч)
'����:strColName-����
    Dim i As Long
    
    For i = 0 To vsf.Cols - 1
        If vsf.TextMatrix(0, i) = strColName Then VsfGetColNum = i: Exit Function
    Next
    VsfGetColNum = -1
End Function

Public Sub RegBillFile()
'���ܣ�ע�����������ļ�
    Dim strSys As String * 255
    
    GetSystemDirectory strSys, 255
    
    RegSetValue HKEY_CLASSES_ROOT, ".zlb", REG_SZ, "zlBill", 7
    RegSetValue HKEY_CLASSES_ROOT, "zlBill", REG_SZ, "ר����ʵ�", 7
    RegSetValue HKEY_CLASSES_ROOT, "zlBill\DefaultIcon", REG_SZ, Left(strSys, InStr(strSys, Chr(0)) - 1) & "\zl9AppTool.dll,15", 24
End Sub

Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNoBeginTrans:û������ʼ
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zldatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub

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

