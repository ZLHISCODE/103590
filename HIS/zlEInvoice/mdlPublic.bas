Attribute VB_Name = "mdlPublic"
Option Explicit

Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ

Public Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Function GetGUID() As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim udtGUID As GUID
    
    On Error GoTo ErrHand
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
                String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
                String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
                IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
                IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
                IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
                IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
                IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
                IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
                IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
                IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
    
    Exit Function
ErrHand:
    'MsgBox Err.Description
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function CDateEx(ByVal strDate As String) As Date
    '---------------------------------------------------------------------------------------
    ' ���� : ��һ���������ڸ�ʽʱ���ַ�������ʽ��Ϊ��Ч��ʽ������ʱ��
    ' ��� :
    ' ���� :
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2020/5/21 17:26
    '---------------------------------------------------------------------------------------
    Dim intLimit As Integer
    Dim strDateTime As String, strTmp As String
    On Error GoTo ErrHand
    If Not IsNumeric(strDate) Then CDateEx = CDate(0): Exit Function
    
    strDateTime = Left(strDate, 4)  '��
    strDate = Mid(strDate, 5)
    If Val(strDateTime) < 1900 Or Val(strDateTime) > 3000 Then CDateEx = CDate(0): Exit Function
    
    If Val(Left(strDate, 2)) > 12 Then '��
        strDateTime = strDateTime & "-" & Left(strDate, 1)
        strDate = Mid(strDate, 2)
    Else
        strDateTime = strDateTime & "-" & Left(strDate, 2)
        strDate = Mid(strDate, 3)
    End If
    
    If IsDate(strDateTime & "-" & Left(strDate, 2)) Then '��
        strDateTime = strDateTime & "-" & Left(strDate, 2)
        strDate = Mid(strDate, 3)
    Else
        strDateTime = strDateTime & "-" & Left(strDate, 1)
        strDate = Mid(strDate, 2)
    End If
    
    If Val(Left(strDate, 2)) > 24 Then 'ʱ
        strDateTime = strDateTime & " " & Left(strDate, 1)
        strDate = Mid(strDate, 2)
    Else
        strDateTime = strDateTime & " " & Left(strDate, 2)
        strDate = Mid(strDate, 3)
    End If
    
    If Val(Left(strDate, 2)) > 60 Then '��
        strDateTime = strDateTime & ":" & Left(strDate, 1)
        strDate = Mid(strDate, 2)
    Else
        strDateTime = strDateTime & ":" & Left(strDate, 2)
        strDate = Mid(strDate, 3)
    End If
    
    If Val(Left(strDate, 2)) > 60 Then '��
        strDateTime = strDateTime & ":" & Left(strDate, 1)
        strDate = Mid(strDate, 2)
    Else
        strDateTime = strDateTime & ":" & Left(strDate, 2)
        strDate = Mid(strDate, 3)
    End If
    
    CDateEx = strDateTime
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
