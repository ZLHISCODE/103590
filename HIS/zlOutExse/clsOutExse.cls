VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutExse"
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
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    
    If glngSys Like "8??" Then
        App.Title = "����ҩ�����"
        gstrCustomerAppellation = "�ͻ�"
    Else
        gstrCustomerAppellation = "����"
    End If
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���ز�����ģ��Ĺ������弤��ʱ���ٴζ�ȡ(�����ȶ�����ΪForm_Loadʱ���ܻ��õ�),��Ϊģ��֮�䴰���л���ı�ȫ�ֱ�����ֵ
    Call InitLocPar(lngModul)
    
    If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
    Call CreateSquareCardObject(frmMain, lngModul)
    Call CreatePublicExpenseObject(lngModul)
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreatePlugIn(lngModul)
    Call InitSysPar
    
    Select Case glngModul
        Case 1120 'ҩ�����۹���
            frmManagePrice.Show , frmMain
        Case 1121 '�����շѹ���
            If gfrmMain Is Nothing Then
                frmManageCharge.Show
            Else
                frmManageCharge.Show , frmMain
            End If
        Case 1122 '������ʹ���
            frmManageBilling.Show , frmMain
        Case 1124 '���ղ������
            frmReplenishTheBalanceManage.Show , frmMain
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
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    
    If glngSys Like "8??" Then
        App.Title = "����ҩ�����"
        gstrCustomerAppellation = "�ͻ�"
    Else
        gstrCustomerAppellation = "����"
    End If
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���ز�����ģ��Ĺ������弤��ʱ���ٴζ�ȡ(�����ȶ�����ΪForm_Loadʱ���ܻ��õ�),��Ϊģ��֮�䴰���л���ı�ȫ�ֱ�����ֵ
    Call InitLocPar(lngModul)
    
    If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(gfrmMain, lngModul)
    Call CreatePublicExpenseObject(lngModul)
    Call CreatePlugIn(lngModul)
    Call InitSysPar
    
    Select Case glngModul
        Case 1120 'ҩ�����۹���
            OS.ShowChildWindow frmManagePrice.hWnd, lngMain
        Case 1121 '�����շѹ���
            OS.ShowChildWindow frmManageCharge.hWnd, lngMain
        Case 1122 '������ʹ���
            OS.ShowChildWindow frmManageBilling.hWnd, lngMain
        Case 1124 '���ղ������
            OS.ShowChildWindow frmReplenishTheBalanceManage.hWnd, lngMain
    End Select
End Sub

Public Function Charge(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, ByVal strDBUser As String, _
    ByVal bytInFun As Byte, ByVal bytInState As Byte, _
    Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, _
    Optional ByVal lngDeptID As Long, Optional ByVal lngUnitID As Long, _
    Optional ByVal bln���� As Boolean, Optional ByVal lng����ҽ�� As Long, _
    Optional ByVal str���ת��ʱ�� As String, Optional ByVal strInNO As String) As Boolean
    '���ܣ�ֱ�ӵ����շ�/���۹���
    '��Σ�
    '   bytInFun 0-�շ�,1-����,2-�������
    '   bytInState 0-ִ��(���޸�),1-���,2-����,3-�˷�(�շѡ����ʲ����˷�),4-�����շ�;5-�쳣��������;11-���Ƶ���
    '   lngUnitID As Long '��ǰ���ʲ���,Ϊ0ʱ��ʾ���в���
    '   lngDeptID As Long '��ǰ���ʿ���,Ϊ0ʱ��ʾ���п���
    '   bln���� As Boolean '33744
    '   strInNO ���뵥�ݣ����˺͸��Ƶ���ʱ���루���������ʱ��Ч��
    Dim lngModule As Long, strPrivs As String
    
    gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    
    lngModule = Choose(bytInFun + 1, 1121, 1120, 1122)
    
    Set gfrmMain = frmMain
    If Not CheckValid Then Exit Function
    
    Set gcnOracle = cnMain
    glngSys = lngSys
    gstrDBUser = strDBUser
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModule) & ";"
    
    Select Case lngModule
    Case 1120
        strPrivs = "����"
    Case 1121
        strPrivs = "�����շ�"
    Case 1122
        strPrivs = IIf(bytInState = 0, "�������", IIf(bytInState = 3, "��������", ""))
    Case Else
        Exit Function
    End Select
    If zlStr.IsHavePrivs(gstrPrivs, strPrivs) = False Then
        MsgBox "��û�б�����ʹ��" & strPrivs & "���ܵ�Ȩ�ޣ�����ʹ�ã�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�û���Ϣ
    Call GetUserInfo

    '��ʼ������
    Call InitLocPar(lngModule)
    Call InitSysPar
    
    If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Function
    Call CreateSquareCardObject(gfrmMain, lngModule)
    Call CreatePublicExpenseObject(lngModule)
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreatePlugIn(lngModule)
    
    '���û���
    Err.Clear
    On Error Resume Next
    frmCharge.mlngModul = lngModule
    frmCharge.mstrPrivs = gstrPrivs
    frmCharge.mbytInFun = bytInFun
    If bytInState = 11 Then '���Ƶ���
        frmCharge.mbytInState = 0
        frmCharge.mblnCopyBill = True
    Else
        frmCharge.mbytInState = bytInState
    End If
    frmCharge.mstrInNO = strInNO
    
    frmCharge.mlng����ID = lng����ID
    frmCharge.mlng��ҳID = lng��ҳID
    frmCharge.mlngUnitID = lngUnitID
    frmCharge.mlngDeptID = lngDeptID
    frmCharge.mbln���� = bln����
    frmCharge.mlng����ҽ�� = lng����ҽ��
    frmCharge.mstr���ת��ʱ�� = str���ת��ʱ��
    If gfrmMain Is Nothing Then
        frmCharge.Show
    Else
        frmCharge.Show 1, frmMain
    End If
    
    Charge = gblnOK
End Function
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
    Load frmPrint
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
    CloseSquareCardObject
    Set grsҽ�Ƹ��ʽ = Nothing
    Unload frmPrint
End Sub
Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Function zlCheckExcuteItemValied(frmMain As Object, cnOracle As ADODB.Connection, ByVal strDBUser As String, _
                                         ByVal lngSys As Long, ByVal lngModule As Long, ByVal lng����ID As Long, _
                                         ByVal str�շ���� As String, ByVal strNos As String, ByVal str�շ�ϸĿIDs As String) _
                                         As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '����:�����ƺ���㷽ʽ�����ִ����Ŀ�ĺϷ���
    '���:frmMain-���õ�������
    '     cnOracle-����Oracle����
    '     strDBUser-���ݿ��û���
    '     lngSys-ϵͳ��
    '     lngModule-���õ�ģ���(��Ҫ�ǿ��Ǻ������ܴ��ڸ���ģ���������ж�)
    '     lng����ID-��ǰ���Ĳ���ID
    '     str�շ����-����ִ�е��շ����,���Դ�����,���ʱ�ö��ŷ���:����:4,5,6,7��,Ϊ��ʱ���������շ������м��
    '     strNOs-����ִ�еĵ��ݺ�(���ʵ���),���Դ����������ʱ�ö��ŷ���,����:4,5�ȣ�Ϊ��ʱ�������ݵ��ݺŽ��м��
    '     str�շ�ϸĿIDs-��ʾ����ִ�е��շ���Ŀ,����Ϊ������ö��ŷ���,���磺123,456�ȣ�Ϊ��ʱ�������շ�ϸĿIDs���м��
    '����:True��ʾ���ݺϷ�,����ִ�У�False-��ʾ���ݲ��Ϸ���������ִ�и���Ŀ
    '����:������
    '����:2013-12-24
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim rsCheck As New ADODB.Recordset, strSql As String
    Dim strReturn As String, strArray() As String
    Set gcnOracle = cnOracle
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrDBUser = strDBUser: glngSys = lngSys: glngModul = lngModule
    strSql = "Select zl_Get_ExcuteItem_Infor([1],[2],[3],[4],[5]) As ������Ϣ From Dual"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSql, "���ִ����Ŀ", lngModule, lng����ID, str�շ����, strNos, str�շ�ϸĿIDs)
    strReturn = Nvl(rsCheck!������Ϣ)
    If strReturn <> "" And InStr(strReturn, "|") > 0 Then
        strArray = Split(strReturn, "|")
        If Val(strArray(0)) = 2 Then
            MsgBox "���ݲ��Ϸ�����ִֹ�У�" & vbCrLf & "��ʾ��Ϣ:" & strArray(1), vbInformation, gstrSysName
            zlCheckExcuteItemValied = False
            Exit Function
        End If
        If Val(strArray(0)) = 1 Then
            If MsgBox("ȷ��Ҫִ����Ŀ��" & vbCrLf & "��ʾ��Ϣ:" & strArray(1), vbYesNo + vbQuestion, gstrSysName) = vbYes Then
                zlCheckExcuteItemValied = True
                Exit Function
            Else
                zlCheckExcuteItemValied = False
                Exit Function
            End If
        End If
    End If
    If strReturn = "" Then
        zlCheckExcuteItemValied = True
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function