VERSION 5.00
Begin VB.UserControl ucHook 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   ScaleHeight     =   330
   ScaleWidth      =   540
   Begin VB.Label Label1 
      Caption         =   "HOOK"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   420
   End
End
Attribute VB_Name = "ucHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mhHook As Long
Private mlngActiveHwnd As Long        '��ǰ���ھ��
Private mblnIsOnlyActive As Boolean   '�Ƿ����Ϊ��ǰ���ھ��ʱ��ִ��hook����

Public Event OnKeyBoardLHook(ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)


Public Sub EnableHook()
'************************************************************************************************
'ע���̹߳���'
'************************************************************************************************
    If mhHook = 0 Then
        'mhHook = SetWindowsHookEx(HOOK_WH_MOUSE_LL, AddressOf HookProc, App.hInstance, 0) '��깳��
        'mhHook = SetWindowsHookEx(HOOK_WH_KEYBOARD, AddressOf HookProc, App.hInstance, 0)  '���̹���
        
        Set hookKey = Me
        
        '�������н���
        mhHook = SetWindowsHookEx(HOOK_WH_KEYBOARD_LL, AddressOf HookProc, App.hInstance, 0)  '���̹���
        
    End If
End Sub

Public Sub FreeHook()
'************************************************************************************************
'�ͷ��̹߳���'
'************************************************************************************************
    If mhHook <> 0 Then
        Call UnhookWindowsHookEx(mhHook)
        mhHook = 0
        
        Set hookKey = Nothing
    End If
End Sub
    
    
    
Private Function GetKeyState(ByVal lParam As Long) As KBDLLHOOKSTRUCT
'���ذ���״̬
    Dim ks As KBDLLHOOKSTRUCT
    
    Call CopyMemory(VarPtr(ks), ByVal lParam, Len(ks))
    
    GetKeyState = ks
End Function

    
Public Function HookProcess(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'************************************************************************************************
'�̹߳��ӻص�����
'
'nCode:
'wParam:��Ϣ����
'lParam:����
'
'************************************************************************************************
    On Error GoTo errHandle
    Dim pid As Long
    Dim ks As KBDLLHOOKSTRUCT
    
    If (nCode <> 0) Then
        HookProcess = CallNextHookEx(mhHook, nCode, wParam, lParam)
        Exit Function
    End If

    '�жϴ�����Ϣ���Ƿ�Ϊ��ǰ����
    Call GetWindowThreadProcessId(GetActiveWindow(), pid)
    If GetCurrentProcessId = pid Then
        ks = GetKeyState(lParam)
        
        RaiseEvent OnKeyBoardLHook(ks.VKCode, ks.scanCode, ks.flags)
    End If
    
errHandle:
    HookProcess = CallNextHookEx(mhHook, nCode, wParam, lParam)
End Function



'---------------------------------------------------------------


Public Property Let ActiveHwnd(lngHwnd As Long)
    mlngActiveHwnd = lngHwnd
End Property

Public Property Get ActiveHwnd() As Long
    ActiveHwnd = mlngActiveHwnd
End Property

Public Property Let IsOnlyActive(blnIsOnlyActive As Boolean)
    mblnIsOnlyActive = blnIsOnlyActive
End Property

Public Property Get IsOnlyActive() As Boolean
    IsOnlyActive = mblnIsOnlyActive
End Property


