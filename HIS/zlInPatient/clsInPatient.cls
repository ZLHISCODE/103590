VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'Ҫ���������
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   frmMain:������
'   strDBUser:��ǰ��¼���ݿ��û���
    Set gfrmMain = frmMain
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call InitLocPar(glngModul)
    Call CreateSquareCardObject(frmMain, lngModul)
    Select Case lngModul
        Case P������λ���� '������λ����
            Call InitSysPar
            frmManageBeds.Show , frmMain
        Case P������Ժ���� '������Ժ����
            Call InitSysPar
            
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            
            If gfrmMain Is Nothing Then
                frmManageHosReg.Show
            Else
                frmManageHosReg.Show , frmMain
            End If
        Case P����������� '�����������
            Call InitSysPar
            
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)

            If gfrmMain Is Nothing Then
                frmManageCourse.Show
            Else
                frmManageCourse.Show , frmMain
            End If
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'���ܣ� ����������ָ�����ܣ�����ִ����س���
'������
'   lngModul:��Ҫִ�еĹ������
'   cnMain:����������ݿ�����
'   lngMain:��������
'   strDBUser:��ǰ��¼���ݿ��û���
    Set gfrmMain = Nothing
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstrUnitName = GetUnitName
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call InitLocPar(glngModul)
    Select Case lngModul
        Case P������λ���� '������λ����
            Call InitSysPar
            Call CreateSquareCardObject(frmManageBeds, lngModul)
            ZLCommFun.ShowChildWindow frmManageBeds.hWnd, lngMain
        Case P������Ժ���� '������Ժ����
            Call InitSysPar
            
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(frmManageHosReg, lngModul)
            ZLCommFun.ShowChildWindow frmManageHosReg.hWnd, lngMain
        Case P����������� '�����������
            Call InitSysPar
            
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreateSquareCardObject(frmManageCourse, lngModul)
            ZLCommFun.ShowChildWindow frmManageCourse.hWnd, lngMain
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

Public Sub InitCallByNurse(frmMain As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, lngSys As Long)
'���ܣ��ɻ�ʿ����վ����ʱ��ʼ������
    
    Set gfrmMain = frmMain
    'If Not CheckValid Then Exit Sub    '���ڻ�ʿվ��ͬʱ��������������,�ڵ���������ʱ��ͨѶԭ��ɾ����,�������ﲻ���ټ��
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = P�����������
    
    gstrPrivs = GetPrivFunc(glngSys, glngModul)
    gstrUnitName = GetUnitName
    
    Call GetUserInfo
    Call InitSysPar
    Call InitLocPar(glngModul)
    
    'ҽ���ӿ�
    Call gclsInsure.InitOracle(gcnOracle)
End Sub
Public Function zl_ExecPatiChange(ByVal bytFun As Byte, ByRef frmParent As Object, ByRef strPrivs As String, ParamArray arrPar() As Variant) As Boolean
'����:ִ�в��˱䶯��ع���
'����:bytFun:0-���
'     arrPar:���ݲ�ͬ�Ĺ��ܵ��ã����벻ͬ�Ĳ���,Ŀǰ���֧��10������
    Select Case UBound(arrPar) + 1
    Case 1
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0))
    Case 2
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1))
    Case 3
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2))
    Case 4
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3))
    Case 5
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4))
    Case 6
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4), arrPar(5))
    Case 7
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4), arrPar(5), arrPar(6))
    Case 8
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4), arrPar(5), arrPar(6), arrPar(7))
    Case 9
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4), arrPar(5), arrPar(6), arrPar(7), arrPar(8))
    Case 10
        zl_ExecPatiChange = ExecPatiChange(bytFun, frmParent, strPrivs, arrPar(0), arrPar(1), arrPar(2), arrPar(3), arrPar(4), arrPar(5), arrPar(6), arrPar(7), arrPar(8), arrPar(9))
    End Select
End Function
Public Sub zlOutFeeToInFee(ByVal frmMain As Object, ByVal cnMain As ADODB.Connection, _
    ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, strDBUser As String, _
    ByVal lng����ID As Long, intPatientRange As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������תסԺ����(�������ֻ��ת����ǰ��Ժ��������)
    '���:frmMain-���õ�������
    '       strPrivs-Ȩ�޴�
    '       lngModule-ģ���
    '       intPatientRange-(0-���в���,1-�κη���δ���岡��;2-���δ����Ĳ���;3-סԺδ����Ĳ���;4-����δ����Ĳ���)
    '����:
    '����:
    '����:���˺�
    '����:2010-11-10 11:09:57
    '����:33635
    '����ģ��:
    '     סԺ���ʹ���-->�������תסԺ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set gfrmMain = frmMain
    If UserInfo.ID = 0 Or glngSys = 0 Or gcnOracle Is Nothing Then
        Set gcnOracle = cnMain
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        gstrDBUser = strDBUser
        glngSys = lngSys: glngModul = lngModule
        gstrPrivs = GetPrivFunc(glngSys, glngModul)
        gstrUnitName = GetUnitName
        Call GetUserInfo
        Call InitSysPar
        Call InitLocPar(glngModul)
        'ҽ���ӿ�
        Call gclsInsure.InitOracle(gcnOracle)
    End If
    If gobjSquare Is Nothing Then Call CreateSquareCardObject(frmMain, lngModule) 'ֻ����һ��,�ڶ������е���
    Call frmChargeTurn.ShowMe(frmMain, lng����ID, "", True, intPatientRange, strPrivs, lngModule)
End Sub


Public Sub zlPatiMemoReadAndWrite(ByVal frmParent As Object, ByVal cnMain As ADODB.Connection, _
    ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String, Optional ByVal blnEdit As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˱�ע��Ϣ�༭/�鿴
    '���:frmParent-���õ�������
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If UserInfo.ID = 0 Or glngSys = 0 Or gcnOracle Is Nothing Then
        Set gcnOracle = cnMain
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        gstrUnitName = GetUnitName
        Call GetUserInfo
    End If
    Call ExecPatiChange(EFun.E���˱�ע�༭, frmParent, strPrivs, lng����ID, lng��ҳID)
End Sub

Private Sub Class_Terminate()
    '62646:������,2013-06-18
    If Forms.Count = 0 Then
        CloseSquareCardObject '�رտ����㲿��
    End If
End Sub