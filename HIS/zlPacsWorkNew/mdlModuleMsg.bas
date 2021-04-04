Attribute VB_Name = "mdlModuleMsg"
Option Explicit

Public Enum TMsgModuleType
    mtImage = 0
    mtVideo
    mtPathol
End Enum


Public Const WM_XWREPORT_IMG As Long = 5120         '���ձ���ͼ����Ϣ��API


'�б������Ϣ
Public Const WM_LIST_SYNCROW As Long = 5001         'ͬ���б�ѡ����
Public Const WM_LIST_REFRESH As Long = 5002         'ˢ�������б�
Public Const WM_LIST_MOVEUP As Long = 5003          '����
Public Const WM_LIST_MOVEDOWN As Long = 5004        '����
Public Const WM_LIST_GETLASTADVICE As Long = 5005     '��ȡ��һ��ҽ��
Public Const WM_LIST_GETNEXTADVICE As Long = 5006   '��ȡ��һ��ҽ��

Public Const WM_IMG_OPENVIEW As Long = 5101         '�򿪹�Ƭ
Public Const WM_IMG_CONTRASTVIEW As Long = 5102         '�Աȹ�Ƭ

Public Const WM_REPORT_VIEW As Long = 5201          '����Ԥ��
Public Const WM_REPORT_PRINT As Long = 5202          '�����ӡ


'Public Const WM_VIEW_REPORT As Long = 0             'Ԥ������
'Public Const WM_VIEW_IMAGE As Long = 0              'Ԥ��ͼ��
'
'Public Const WM_EDITOR_LOCK As Long = 0             '�����༭
'Public Const WM_EDITOR_UNLOCK As Long = 0           '�����༭

'���˵�ִ��
Public Const BM_SYS__EVENT_MENU As Long = 1001

'RIS�����Ϣ
Public Const BM_RIS_EVENT_REGISTER As Long = 4001         '���Ǽ�
Public Const BM_RIS_EVENT_RECEVIE As Long = 4002          '��鱨��
Public Const BM_RIS_EVENT_COMPLETE  As Long = 4003      '������
Public Const BM_RIS_EVENT_CANCELREG As Long = 4004      'ȡ���Ǽ�
Public Const BM_RIS_EVENT_CANCELREC As Long = 4005      'ȡ������
Public Const BM_RIS_EVENT_CANCELCOMP As Long = 4006 'ȡ�����

'���������Ϣ
Public Const BM_REPORT_EVENT_PRINT As Long = 6101           '�����ӡ�¼�(plugin...)
Public Const BM_REPORT_EVENT_SAVE As Long = 6102            '���汣���¼�(plugin...)
Public Const BM_REPORT_EVENT_POPUPEXIT As Long = 6103   '�������ڴ������˳��¼�
Public Const BM_REPORT_EVENT_SIGN As Long = 6104            '����ǩ���¼�(plugin...)
Public Const BM_REPORT_EVENT_AUDIT As Long = 6105           'Ԥ�����������
Public Const BM_REPORT_EVENT_REJECT As Long = 6106         '���沵���¼�(plugin...)
Public Const BM_REPORT_EVENT_DELETE As Long = 6107          '����ɾ���¼�(plugin...)
Public Const BM_REPORT_EVENT_BACK As Long = 6108            '��������¼�(plugin...)
Public Const BM_REPORT_EVENT_Verify As Long = 6109          '������֤�¼�(plugin...)
Public Const BM_REPORT_EVENT_REJHISTORY As Long = 6110      '������ʷ�鿴(plugin...)
Public Const BM_REPORT_EVENT_OPEN As Long = 6111            '������¼�
Public Const BM_REPORT_EVENT_IMGCHANGE As Long = 6112       '����ͼ�ı��¼�
Public Const BM_REPORT_EVENT_QUALITY As Long = 6113         '������������¼�
Public Const BM_REPORT_EVENT_ADDIMG As Long = 6114          '����ͼ����¼�
Public Const BM_REPORT_EVENT_CLOSEEPR As Long = 6115        '���洰�ڹر��¼�
Public Const BM_REPORT_EVENT_REFWCHR As Long = 6116         'ˢ�³��ôʾ��ַ�
Public Const BM_REPORT_EVENT_DELREPIMG As Long = 6117       '����ͼɾ���¼�
Public Const BM_REPORT_EVENT_REFFRAGMENT As Long = 6118     'ˢ�´ʾ�Ƭ��

'ͼ�������Ϣ
Public Const BM_IMAGE_EVENT_DEL As Long = 6200              'ɾ��ͼ��
Public Const BM_IMAGE_EVENT_CAPTURE As Long = 6201    '�ɼ�ͼ��
Public Const BM_IMAGE_EVENT_FIRST As Long = 6202   '�ɼ�����ͼ��

Public Const BM_IMAGE_EVENT_QUALITYTAG As Long = 6203   '���Ӱ������
Public Const BM_IMAGE_EVENT_XWFILMPRINT   As Long = 6204  '��Ƭ��ӡ
Public Const BM_IMAGE_EVENT_GETIMAGE    As Long = 6205  '��ȡӰ��
Public Const BM_IMAGE_EVENT_TECHDO      As Long = 6206  '��ʦִ��
Public Const BM_IMAGE_EVENT_CHANGEDEVICE As Long = 6207 '�����豸



'���������Ϣ��������ǰ�汾����
Public Const BM_PATHOL_EVENT_BASE As Long = 7000

Private mlngImageProcHwnd As Long
Private mlngVideoProcHwnd As Long
Private mlngPatholProcHwnd As Long


Public gobjImageMainWindow As Object                  '�������ձ���ͼ��Ϣ�Ĵ���ָ��
Public gobjVideoMainWindow As Object
Public gobjPatholMainWindow As Object


Public Sub AttachModuleMsgProc(moduleType As TMsgModuleType, objMainWindow As Object)
    'ָ���Զ���Ĵ��ڹ���
    '���ز�����ԭ��Ĭ�ϵĴ��ڹ���ָ��
    Dim lngOldProcHwnd As Long
On Error GoTo errhandle:
        
    If App.LogMode = 0 Then Exit Sub
    
    lngOldProcHwnd = SetWindowLong(objMainWindow.hwnd, GWL_WNDPROC, AddressOf MainWindowProc)
    
    Select Case moduleType
        Case mtImage
            mlngImageProcHwnd = lngOldProcHwnd
            Set gobjImageMainWindow = objMainWindow
        Case mtVideo
            mlngVideoProcHwnd = lngOldProcHwnd
            Set gobjVideoMainWindow = objMainWindow
        Case mtPathol
            mlngPatholProcHwnd = lngOldProcHwnd
            Set gobjPatholMainWindow = objMainWindow
    End Select
     
    Exit Sub
errhandle:
    
End Sub

Public Sub UnAttachModuleMsgProc(ByVal hwnd As Long, moduleType As TMsgModuleType)
On Error GoTo errhandle
    Dim temp As Long
    Dim lpWndProc As Long
    
    If hwnd = 0 Then Exit Sub
        
    Select Case moduleType
        Case mtImage
            lpWndProc = mlngImageProcHwnd
        Case mtVideo
            lpWndProc = mlngVideoProcHwnd
        Case mtPathol
            lpWndProc = mlngPatholProcHwnd
    End Select

    temp = SetWindowLong(hwnd, GWL_WNDPROC, lpWndProc)
    
    Exit Sub
errhandle:

End Sub


Function MainWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'��Ϣ�������
Dim lngProcHwnd As Long
Dim objProc As Object

On Error GoTo errhandle
    lngProcHwnd = 0
    Set objProc = Nothing
    
    If Not gobjImageMainWindow Is Nothing Then
        If hw = gobjImageMainWindow.hwnd Then
            Set objProc = gobjImageMainWindow
            lngProcHwnd = mlngImageProcHwnd
        End If
    End If
    
    If Not gobjVideoMainWindow Is Nothing Then
        If hw = gobjVideoMainWindow.hwnd Then
            Set objProc = gobjVideoMainWindow
            lngProcHwnd = mlngVideoProcHwnd
        End If
    End If
    
    If Not gobjPatholMainWindow Is Nothing Then
        If hw = gobjPatholMainWindow.hwnd Then
            Set objProc = gobjPatholMainWindow
            lngProcHwnd = mlngPatholProcHwnd
        End If
    End If
    
    If Not objProc Is Nothing Then
        Call objProc.MainWindowProc(hw, uMsg, wParam, lParam)
    End If
 
    '����ԭ���Ĵ��ڹ���
    MainWindowProc = CallWindowProc(lngProcHwnd, hw, uMsg, wParam, lParam)
Exit Function
errhandle:
    If lngProcHwnd <> 0 Then
        MainWindowProc = CallWindowProc(lngProcHwnd, hw, uMsg, wParam, lParam)
    End If
    
End Function

