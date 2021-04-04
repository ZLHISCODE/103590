VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInExse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'Ҫ���������
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
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
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    gblnShowInTaskBar = blnShowInTaskBar
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call InitLocPar(lngModul)
    Call CreateSquareCardObject(frmMain, lngModul)
    Call CreatePublicExpenseObject(lngModul)
    Select Case lngModul
        Case 1133 'סԺ���ʹ���
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            Err.Clear
            On Error Resume Next
            If gfrmMain Is Nothing Then
                frmManageBilling.Show
            Else
                frmManageBilling.Show , frmMain
            End If
        Case 1134 '���ҷ�ɢ����
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            Err.Clear
            On Error Resume Next
            If gfrmMain Is Nothing Then
                frmDeptBilling.Show
            Else
                frmDeptBilling.Show , frmMain
            End If
        Case 1135 'ҽ�����Ҽ���
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            Err.Clear
            On Error Resume Next
            If gfrmMain Is Nothing Then
                frmTechnoBilling.Show
            Else
                frmTechnoBilling.Show , frmMain
            End If
        Case 1136 '�Զ����ʼ���
            frmCompute.Show 1, frmMain
        Case 1137 '���˽��ʴ���
            Call InitSysPar
                        
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            If gfrmMain Is Nothing Then
                frmManageBalance.Show
            Else
                frmManageBalance.Show , frmMain
            End If
        Case 1139 '���˷��ò�ѯ
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
            
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            '����ȡ��:gobjKernel,ԭ�����ڷ��ò�ѯ��Ҫ�õ�zlCISKernel.clsDockInAdvices
            '��Ϊ;zlCISKernel.clsDockInAdvicesδ����initCisKernel
            Set gobjKernel = CreateObject("zlCISKernel.clsCISKernel")
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs)
            gblnNurseStation = False
            If gfrmMain Is Nothing Then
                frmPatiFeeQuery.Show
            Else
                frmPatiFeeQuery.Show , frmMain
            End If
        Case 1141 'һ�շ����嵥
            Call InitSysPar
            frmDailyList.Show , frmMain
        Case 1142 'ִ�еǼǹ���
            Call InitSysPar
            On Error Resume Next
            frmManageExecute.Show , frmMain
        Case 1143 '������˹���
            Call InitSysPar
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            On Error Resume Next
            frmFeeVerifyManger.Show , frmMain
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
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    gblnShowInTaskBar = True
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call InitLocPar(lngModul)
    Call CreateSquareCardObject(gfrmMain, lngModul)
    Call CreatePublicExpenseObject(lngModul)
    Select Case lngModul
        Case 1133 'סԺ���ʹ���
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmManageBilling.hWnd, lngMain)
        Case 1134 '���ҷ�ɢ����
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmDeptBilling.hWnd, lngMain)
        Case 1135 'ҽ�����Ҽ���
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmTechnoBilling.hWnd, lngMain)
        Case 1136 '�Զ����ʼ���
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmCompute.hWnd)
        Case 1137 '���˽��ʴ���
            Call InitSysPar
                        
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmManageBalance.hWnd, lngMain)
        Case 1139 '���˷��ò�ѯ
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
            
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            '����ȡ��:gobjKernel,ԭ�����ڷ��ò�ѯ��Ҫ�õ�zlCISKernel.clsDockInAdvices
            '��Ϊ;zlCISKernel.clsDockInAdvicesδ����initCisKernel
            Set gobjKernel = CreateObject("zlCISKernel.clsCISKernel")
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs)
                         
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmPatiFeeQuery.hWnd, lngMain)
        Case 1141 'һ�շ����嵥
            Call InitSysPar
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmDailyList.hWnd, lngMain)
        Case 1142 'ִ�еǼǹ���
            Call InitSysPar
            
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmManageExecute.hWnd, lngMain)
        Case 1143 '������˹���
            Call InitSysPar
            'ҽ���ӿ�
            Call gclsInsure.InitOracle(gcnOracle)
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmFeeVerifyManger.hWnd, lngMain)
    End Select
End Sub

Public Function CallByNurse(frmMain As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, lngSys As Long, _
    ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal lng���ͺ� As Long, ByVal lngҽ��ID As Long, ByVal strNO As String, ByVal bln���� As Boolean) As Object
'���ܣ��ɻ�ʿ����վ���ò��Զ�������ʾ��ָ��ҽ�������ķ��õ��ݣ�Ŀ���ǳ�����Щ����
'���أ����ӷ�ģ̬���ڷ���ʱ,���ؼ��ʹ�������,���ڸ��ٹر��¼�(��ģ̬��ʾʱ��ˢ������)
    Static blnNoFirst As Boolean  '�Ƿ��һ�ε���,��ʼ��ȫ�ֱ���
    
    If Not blnNoFirst Then
        blnNoFirst = True
    
        Set gfrmMain = frmMain
        If Not CheckValid Then Exit Function
        
        Set gcnOracle = cnMain
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        gstrDBUser = strDBUser
        glngSys = lngSys
        glngModul = 1133
        
        gstrPrivs = ";" & GetPrivFunc(glngSys, glngModul) & ";"
        gstrUnitName = GetUnitName
        
        If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Function
        Call GetUserInfo
        Call InitSysPar
        Call InitLocPar(glngModul)
        Call CreateSquareCardObject(frmMain, glngModul)
        Call CreatePublicExpenseObject(glngModul)
        
        'ҽ���ӿ�
        Call gclsInsure.InitOracle(gcnOracle)
        Call CreatePlugIn(glngModul)
    End If
    
    Set CallByNurse = frmManageBilling.ShowMeByNurse(frmMain, lng����ID, lng����ID, lng���ͺ�, lngҽ��ID, strNO, bln����)
End Function

Public Function CallReCharge(ByVal frmParent As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, ByVal lngSys As Long, _
    ByVal bytFun As Integer, ByVal bytUseType As Byte, ByVal lngDeptID As Long, ByVal strPrivs As String, Optional ByVal lngPatientID As Long, _
    Optional ByVal strNO As String, Optional ByVal lngAdviceID As Long) As Boolean
'���ܣ����������������˹���(Ŀǰ��ҽ��վ����ʿվ��ҽ������վ�ڵ���)
'������bytFun=0-����,1-���
'      bytUseType=0-��ʿվ����,1-ҽ�����ҵ���,2-ҽ��վ����
'      lngDeptID=��������ʱ���뵱ǰ�����Ĳ��˲���ID,ҽ�����ҵ���ʱ����ҽ������ID
'      strPrivs=���ó��������е�Ȩ��,Ŀǰ��ʱδʹ�ã���Ϊ�������봰�����Լ�ȡ����1150ģ���Ȩ��
    Static blnNoFirst As Boolean  '�Ƿ��һ�ε���,��ʼ��ȫ�ֱ���
        
    If Not blnNoFirst Then
        blnNoFirst = True
        
        Set gfrmMain = frmParent
        If Not CheckValid Then Exit Function
        
        Set gcnOracle = cnMain
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        gstrDBUser = strDBUser
        glngSys = lngSys
        glngModul = 1133
                
        '����Ҫȡ1133סԺ���ʹ���ģ���Ȩ��
                
        Call GetUserInfo
        Call InitSysPar
        Call InitLocPar(glngModul)
        Call CreateSquareCardObject(frmParent, glngModul)
        Call gclsInsure.InitOracle(gcnOracle)
        Call CreatePlugIn(glngModul)
        Call CreatePublicExpenseObject(glngModul)
    End If
    
    gblnOK = False
    With frmReCharge
        .mbytFun = bytFun
        .mbytUseType = bytUseType
        .mlngDeptID = lngDeptID
        .mlngPatientID = lngPatientID
        .mstrPrivs = strPrivs
        .mstrInNO = strNO
        .mlngAdviceID = lngAdviceID
        If frmParent Is Nothing Then
            .Show 1
        Else
            .Show 1, frmParent
        End If
    End With
    CallReCharge = gblnOK
End Function

Public Function CallMzFeeTOZyFeeDel(ByVal frmMain As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, ByVal lngSys As Long, _
    ByVal lngModule As Long, ByVal strPrivs As String, ByVal int���� As Integer, Optional ByRef lng����ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����תסԺ�������ʻ��˷�
    '���:int����-1-�����շ�;2-�������
    '����:ֻҪ��һ������ת���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-25 11:36:41
    '˵��:Ŀǰ�ýӿ���Ҫ�����շѹ���������תסԺ�˷ѹ����е���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static blnNoFirst As Boolean  '�Ƿ��һ�ε���,��ʼ��ȫ�ֱ���
    If Not blnNoFirst Then
        blnNoFirst = True
        Set gfrmMain = frmMain
        Set gcnOracle = cnMain
        gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
        gstrDBUser = strDBUser: glngSys = lngSys: glngModul = lngModule
        Call GetUserInfo
        Call InitSysPar
        Call InitLocPar(glngModul)
        Call gclsInsure.InitOracle(gcnOracle)
        Call CreateSquareCardObject(frmMain, glngModul)
    End If
    If frmFeeRefundment.zlShowEdit(frmMain, int����, glngModul, strPrivs, lng����ID) = False Then Exit Function
    CallMzFeeTOZyFeeDel = True
End Function
Public Function FetchUnBalance(cnMain As ADODB.Connection, strPrivs As String, _
    int���� As Integer, lng����ID As Long, _
    Optional intҵ������ As Integer = 2) As ADODB.Recordset
    '���ܣ���ȡָ������δ����ϸĿ��ϸ(���շ�ϸĿ)
    '������lng����ID=����ID,
    '      int��������-1-����;2-סԺ(ȱʡ)
    '���أ��ɹ�=��¼��,ʧ��=Nothing
    Call InitSysPar
    Set gcnOracle = cnMain
    If intҵ������ = 1 Then
        Set FetchUnBalance = GetVBalance(0, strPrivs, int����, lng����ID)
    Else
        Set FetchUnBalance = GetVBalance(1, strPrivs, int����, lng����ID)
    End If

End Function

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    If Not gobjKernel Is Nothing Then
        blnChildren = gobjKernel.CloseWindows()
    Else
        blnChildren = True
    End If
    CloseWindows = blnChildren And (Forms.Count = 0)
End Function
Private Sub Class_Initialize()
    glngInstanceCount = glngInstanceCount + 1
    Version = App.Major & "." & App.Minor & "." & App.Revision
    Load frmPrint
End Sub
Private Sub Class_Terminate()
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
    
    Set grsҽ�Ƹ��ʽ = Nothing
    Unload frmPrint
    'Call CloseSquareCardObject: ����ʱ , ���ܹر�, ����Ҫ����, ���˳�ʱ, ������Զ��ر�, 59760
End Sub