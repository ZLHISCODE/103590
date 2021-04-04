VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not CheckValid Then Exit Sub

    Set gfrmMain = frmMain
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    glngSys = lngSys
    
    gstrUnitName = GetUnitName
    gblnInsure = (UCase(GetSetting("ZLSOFT", "����ȫ��", "�Ƿ�֧��ҽ��", "")) = UCase("Yes"))
    gintInsure = Val(GetSetting("ZLSOFT", "����ȫ��", "ҽ�����", 0))
    '-------------------------------------------------
    Select Case lngModul
        Case 1530           '�û�������
            frmDefTable.Show , frmMain
        Case 1531           '��ѯͼ������
            frmPicture.Show , frmMain
        Case 1532           'ר�ҽ����嵥
            frmDoctor.Show , frmMain
        Case 1533           '���Ź������
            frmAdvice.Show , frmMain
        Case 1534           '��ѯҳ�涨��
            frmDefQuery.Show , frmMain
        Case 1535           '��ѯĿ¼�滯
            frmDefTree.Show , frmMain
        Case 1536           '��ѯ������ʾ
                        'ҽ���ӿ�
            If gblnInsure Then
                If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
            End If
            
            frmMainQuery.Show , frmMain
        Case 1537           '��ѯ��������
            frmParameter.Show 1, frmMain
        Case 1538           '�����ҺŹ���
            Call InitLocPar
            Call InitSysPar
            On Error Resume Next
            frmselectinfo.Show , frmMain
    End Select
End Sub

Private Sub Class_Terminate()
    '
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
