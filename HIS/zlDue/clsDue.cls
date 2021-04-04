VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDue"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, FrmMain As Object, strDbUser As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngSys : ϵͳ���
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '���أ�
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    
    gstrDBUser = strDbUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    Set gfrmMain = FrmMain
    Call ��ʼС��λ��
    
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    glngSys = lngSys
    gstrUnitName = GetUnitName
    
    '��ʼϵͳ����ֵ
    Call zlinitSystemPara
    '-------------------------------------------------
    On Error Resume Next
    Select Case lngModul
        Case 1025
            frm��Ӧ�̹���.Show , FrmMain
        Case 1322
            frmӦ�������.Show , FrmMain
        Case 1323
            frm�������.Show , FrmMain
        Case 1324
            frm�嵥����.Show , FrmMain
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, strDbUser As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngSys : ϵͳ���
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   lngMain:��������
    '���أ�
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    
    gstrDBUser = strDbUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    Set gfrmMain = Nothing
    Call ��ʼС��λ��
    
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    glngSys = lngSys
    gstrUnitName = GetUnitName
    
    '��ʼϵͳ����ֵ
    Call zlinitSystemPara
    '-------------------------------------------------
    On Error Resume Next
    Select Case lngModul
        Case 1025
            zlCommFun.ShowChildWindow frm��Ӧ�̹���.hwnd, lngMain
        Case 1322
            zlCommFun.ShowChildWindow frmӦ�������.hwnd, lngMain
        Case 1323
            zlCommFun.ShowChildWindow frm�������.hwnd, lngMain
        Case 1324
            zlCommFun.ShowChildWindow frm�嵥����.hwnd, lngMain
    End Select
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function


