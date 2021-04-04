VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ��¼���ݿ��û���

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If

    Set gfrmMain = frmMain
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    glngModul = lngModul
    glngSys = lngSys
    
    gstrUnitName = GetUnitName
    gstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
'    gblnInsure = (UCase(GetSetting("ZLSOFT", "����ȫ��", "�Ƿ�֧��ҽ��", "")) = UCase("Yes"))
'    gintInsure = Val(GetSetting("ZLSOFT", "����ȫ��", "ҽ�����", 0))
    gblnInsure = True
    
    '-------------------------------------------------
    On Error Resume Next
    
    Call InitCISCore(gcnOracle)
    
    Select Case lngModul
    Case 1850               '�������
        frmKind.Show , frmMain
    Case 1851               '�����Ͻ���
        frmDiagnoseAdvice.Show , frmMain
    Case 1852               '�����Ŀ����
        frmMedicalItems.Show , frmMain
    Case 1860               '���ԤԼ����
        frmSchedule.Show , frmMain
    Case 1861               '��칤������
'
'        If gblnInsure Then
'            If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
'        Else
'            Call gclsInsure.InitOracle(gcnOracle)
'        End If
        
        frmMedicalStation.Show , frmMain
    Case 1862               '����������
        frmBalance.Show , frmMain
    Case 1863               '�����ù���
        frmLaterVisit.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ��¼���ݿ��û���

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If

    Set gfrmMain = Nothing
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    glngModul = lngModul
    glngSys = lngSys
    
    gstrUnitName = GetUnitName
    gstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    gblnInsure = True
    
    '-------------------------------------------------
    On Error Resume Next
    
    Call InitCISCore(gcnOracle)
    
    Select Case lngModul
    Case 1850               '�������
        zlCommFun.ShowChildWindow frmKind.hWnd, lngMain
    Case 1851               '�����Ͻ���
        zlCommFun.ShowChildWindow frmDiagnoseAdvice.hWnd, lngMain
    Case 1852               '�����Ŀ����
        zlCommFun.ShowChildWindow frmMedicalItems.hWnd, lngMain
    Case 1860               '���ԤԼ����
        zlCommFun.ShowChildWindow frmSchedule.hWnd, lngMain
    Case 1861               '��칤������
        zlCommFun.ShowChildWindow frmMedicalStation.hWnd, lngMain, True
    Case 1862               '����������
        zlCommFun.ShowChildWindow frmBalance.hWnd, lngMain
    Case 1863               '�����ù���
        zlCommFun.ShowChildWindow frmLaterVisit.hWnd, lngMain
    End Select
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    Load frmPrint
End Sub


Private Sub Class_Terminate()
    Unload frmPrint
End Sub