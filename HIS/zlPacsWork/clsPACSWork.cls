VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPACSWork"
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
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call InitCISCore(gcnOracle)
    Call InitTechCore(gcnOracle)
    
    '-------------------------------------------------
    'PASS�ӿڳ�ʼ��:ֻ��ʼ��һ��
    '��Ϊ����ģ�����ͬʱʹ��,�ҹ��ñ���gblnPass����˻���Ҫ�ھ���ģ�����ж�Ȩ��
    '-------------------------------------------------
    On Error Resume Next
    Select Case lngModul
        Case 1056 'Ӱ������Ŀ
            frmRadLists.Show , gfrmMain
        Case 1057 'Ӱ���豸Ŀ¼
            frmPacsDev.Show , gfrmMain
        Case 1207 'Ӱ��ʦ����վ
            frmPACStation.Show , gfrmMain
'            frmVideoStation.Show , gfrmMain
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
End Sub
