VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'------------------------------------------------------------------------------
'˵����ҩƷ�Զ����豸�ӿ���
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'ע�⣺Init������Operation������Version���ԣ�����ZLHIS�ĵ��ýӿڣ�����������
'------------------------------------------------------------------------------

Private mstrVersion As String
Private mblnInit As Boolean
Private mobjComLib As Object
Private mclsLog As New clsLog
Private mtypParams As TYPE_PARAMS
Private mintType As Integer
Private mstrUser As String
Private mblnDrugMachine As Boolean          '����ҩƷ�Զ���
Private mblnIIP As Boolean                  '������Ϣ����ƽ̨
Private mstrIIP_URL As String               '��Ϣ����ƽ̨��ַ
Private mstrKey As String                   '��Ϣ����ƽ̨��Կ
Private mstrToken As String                 '��Ϣ����ƽ̨����

Public Property Get Version() As String
Attribute Version.VB_UserMemId = 0
    mstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    Version = mstrVersion
End Property

Public Sub SaveLog(ByVal strContent As String)
'���ܣ�����־���ݱ��浽zlDrugMachine.log�ĵ���
'������
'  strContent����־����
    
    Dim strMessage As String

    If mblnInit = False Then
        strMessage = "δ���ó�ʼ���������ʼ��δ�ɹ�"
        mclsLog.Add strMessage
    End If
    
    mclsLog.Add strContent
    mclsLog.Save
End Sub

Public Function Operation(ByVal strUser As String, ByVal intType As Integer, ByVal strData As String, _
    ByRef strMessage As String) As Boolean
'���ܣ������������ӿ�
'������
'  strUser����¼ZLHIS���û���
'  intType��ҵ�����
'           1-������Ϣ��
'           2-��Ա��Ϣ��
'           3-ҩƷĿ¼��
'           4-ҩƷ������λ��
'           5-��ҩ���ڣ�
'           21-��ҩ[�����סԺ������ϸ�ϴ�]��
'           22-��ʼ��ҩ��
'           23-��ɷ�ҩ��
'           24-����(����/����)��ҩ��
'           25-����������ҩ��
'  strData��ҵ������
'           ҵ�����=1���ӿڱ��|��������1;��������2;��
'           ҵ�����=2���ӿڱ��|��Ա����1;��Ա����2;��
'           ҵ�����=3���ӿڱ��|���ͱ���1;���ͱ���2;��
'           ҵ�����=4���ӿڱ��|�ⷿid1;�ⷿid2;��
'           ҵ�����=5���ӿڱ��|�ⷿid1;�ⷿid2;��
'           ҵ�����=21 to 23��
'               ���1|����1,������1;����2,������2;��  ˵��������=8���շѵ�������=9�����ʵ��� ����1|8,I0000001;8,I0000002
'               סԺ��2|�շ�ID1;�շ�ID2;��
'           ҵ�����=24���շ�ID1,��ҩ����1;�շ�ID2,��ҩ����2;��   ��ҩ�����ĵ�λΪZLHISҩƷ��С��λ
'           ҵ�����=25������1,������1,�ⷿid1;����2,������2,�ⷿid2;��
'  strMessage��������Ϣ
'
'���أ�True�ɹ���Falseʧ��

    Const INT_LEN As Integer = 5     'strData���ȳ����󣬵ݹ鴦��

    Dim objINF As clsINF
    Dim rsData As ADODB.Recordset
    Dim arrCode As Variant
    Dim strCode As String
    Dim l As Long
    Dim blnOK As Boolean
    Dim intIO As Integer    '1-���2-סԺ��
    
'�ֽ�������������
    Dim strIDs As String, strTmp As String
    Dim i As Long

'    If intType = 21 And Val(strData) = 2 Then
'        strIDs = Split(strData, "|")(1)     '�շ�id
'        If Len(strIDs) > INT_LEN Then
'            '�ֽ��շ�id
'            For i = INT_LEN To 1 Step -1
'                If Mid(strIDs, i, 1) = ";" Then
'                    strTmp = "2|" & Mid(strIDs, i + 1)
'                    Exit For
'                End If
'            Next
'
'            '�ݹ�
'            Call Operation(strUser, intType, strTmp, strMessage)
'        Else
'            '�ݹ���ײ�
'            With mdlDefine.gtypYUYAMA
'                .BottomLevel = True
'                .MacNO = 0
'                .SendIDs = ""
'            End With
'        End If
'
'        '�ų��ѷ��͵�ID
'        If mdlDefine.gtypYUYAMA.SendIDs <> "" And InStrRev(strData, mdlDefine.gtypYUYAMA.SendIDs) > 0 Then
'            strData = Left(strData, InStrRev(strData, mdlDefine.gtypYUYAMA.SendIDs) - 2)
'        End If
'        mdlDefine.gtypYUYAMA.SendIDs = Split(strData, "|")(1)
'        Debug.Print strData
'    End If
    mclsLog.AddTime
    mstrUser = strUser

    If mblnInit = False Then
        strMessage = "δ���ó�ʼ���������ʼ��δ�ɹ�"
        mclsLog.Add strMessage
        mclsLog.Save
        Exit Function
    End If
    
    If mblnDrugMachine = False Then
        mclsLog.Add "δ����ҩƷ�Զ����豸�ӿ�"
        mclsLog.Save
        Exit Function
    End If
    
    mclsLog.Add "��ʼ�����������ӿ�"
    mclsLog.Add "����1:" & strUser, 1, 1
    mclsLog.Add "����2:" & intType, 1, 1
    mclsLog.Add "����3:" & strData, 1, 1
    
    '��ʼ��
    Set objINF = New clsINF
    With objINF
        Set .mobjComLib = mobjComLib
        Set .mobjLog = mclsLog
        
'        .mobjLog.Path = mclsLog.Path
'        .mobjLog.FileName = mclsLog.FileName
'        .mobjLog.Action = mclsLog.Action
'        .mobjLog.SaveDays = mclsLog.SaveDays
'        .mobjLog.Type_ = mclsLog.Type_
    End With
    
    '��ȡZLHIS�ı�׼����
    Set rsData = objINF.GetData(intType, strData, blnOK)      '��¼�����ӿ�����
    If rsData Is Nothing Then
        If blnOK = False Then
            strMessage = "��ȡZLHIS��׼����ʧ��"
            Operation = False
        Else
            strMessage = ""
            Operation = True
        End If
        
        Exit Function
    Else
        If rsData.RecordCount <= 0 Then
            mclsLog.Add "��ȡZLHIS��׼������ʱ��"
            strMessage = ""
            Operation = True
            Exit Function
        End If
    End If
    
    If intType > 20 Then
        'ҵ������
        
        If intType >= 21 And intType <= 23 Then     '21-��ҩ��22-��ʼ��ҩ��23-��ɷ�ҩ
            '��ȡ��1-���2-סԺ����Ϣ
            intIO = Val(strData)
        End If
    
        '��ȡ���ݵĽӿڱ��
        arrCode = Split(GetINFCodes(rsData), ";")
        
        '�����ӿ�
        For l = LBound(arrCode) To UBound(arrCode)
            If Trim(arrCode(l)) <> "" Then
                Call OperationEx(intType, intIO, strData, rsData, objINF, arrCode(l), strMessage)
            Else
                mclsLog.Add "ҵ�����ݵĽӿڱ����"
            End If
        Next
        
        If UBound(arrCode) < 0 Then mclsLog.Add "ҵ�����ݵĽӿڱ����"
        
    Else
        '��������
        
        '��ȡ�ӿڱ��
        strCode = Split(strData, "|")(0)
        
        If Trim(strCode) <> "" Then
            If OperationEx(intType, intIO, "", rsData, objINF, strCode, strMessage) = False Then
                Operation = False
                If Not rsData Is Nothing Then
                    If rsData.State = adStateOpen Then rsData.Close
                    Set rsData = Nothing
                End If
                Exit Function
            End If
        Else
            mclsLog.Add "�������ݵĽӿڱ����"
        End If
        
    End If
    
    '�رռ�¼��
    If Not rsData Is Nothing Then
        If rsData.State = adStateOpen Then rsData.Close
        Set rsData = Nothing
    End If
    
    '��ɴ��Ͳ���
    mdlDefine.gtypYUYAMA.BottomLevel = False
    mclsLog.Add "��ɴ����������ӿ�"
    mclsLog.Save
    strMessage = ""
    Operation = True
    
    Exit Function
    
hErr:
    strMessage = Err.Number & "��" & Err.Description
    mclsLog.Add strMessage
    mclsLog.Add "�����������ӿ�ʧ��"
    mclsLog.Save
End Function

Private Function OperationEx(ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strData As String, ByVal rsData As ADODB.Recordset, _
    ByVal objBase As clsINF, ByVal strCode As String, ByRef strMessage As String) As Boolean
'������
'  intType��ҵ������
'  intIO��1-���2-סԺ
'  strData��ҵ������
'  rsData��ҵ�����ݵļ�¼������
'  objBase���ӿڻ���
'  strCode���ӿڱ��

    Dim objThird As clsINF
    Dim rsFinish As ADODB.Recordset
    Dim blnOK As Boolean

    '���ֽӿ�
    
    If mblnIIP Then
        '������Ϣ����ƽ̨
        Set objThird = New clsINF_IIP
        With objThird
            Set .mobjComLib = objBase.mobjComLib
            Set .mobjLog = objBase.mobjLog
            .URL = mstrIIP_URL
            .KEY = mstrKey
            .Token = mstrToken
            .UserName = mstrUser
        End With
    Else
        '�������ӿ�
        Select Case GetINFType(strCode)
        Case 1          'Τ�ֺ���
            '��ע�⡰Τ�ֺ��ġ�Ŀǰ��������ͬ�汾����ģ��
            Set objThird = New clsINF_WillachHeise
            With objThird
                Set .mobjComLib = objBase.mobjComLib
                Set .mobjLog = objBase.mobjLog
                .URL = mdlDrugMachine.GetInterfaceLink(objBase, strCode)    'WebService��ַ
                .UserName = mstrUser
                
                '���Ҫ��ȡΤ�ֺ��ĵ���������
                'If rsData.State = adStateOpen Then rsData.Close
                'Set rsData = .GetData(intType, strData, blnOK)
            End With
        Case 2          'TOSHO
            Set objThird = New clsINF_TOSHO
            With objThird
                Set .mobjComLib = objBase.mobjComLib
                Set .mobjLog = objBase.mobjLog
                .URL = mdlDrugMachine.GetInterfaceLink(objBase, strCode)    'OLEDB���Ӵ�
                .UserName = mstrUser
                
                '���Ҫ��ȡTOSHO����������
                'If rsData.State = adStateOpen Then rsData.Close
                'Set rsData = .GetData(intType, strData, blnOK)
            End With
        Case 3          '����
            Set objThird = New clsINF_DIH
            With objThird
                Set .mobjComLib = objBase.mobjComLib
                Set .mobjLog = objBase.mobjLog
                .URL = mdlDrugMachine.GetInterfaceLink(objBase, strCode)    'WebService��ַ
                .UserName = mstrUser
            
                '���»�ȡ���͵���������
                If intType = Val("22-��ʼ��ҩ") Or intType = Val("23-��ɷ�ҩ") Then
                    If rsData.State = adStateOpen Then rsData.Close
                    Set rsData = .GetData(intType, strData, blnOK)
                    If rsData Is Nothing Then
                        If blnOK = False Then
                            strMessage = "��ȡZLHIS��׼���ݣ����ͣ�ʧ��"
                            OperationEx = False
                        Else
                            strMessage = ""
                            OperationEx = True
                        End If
                        Exit Function
                    Else
                        If rsData.RecordCount <= 0 Then
                            mclsLog.Add "��ȡZLHIS��׼���ݣ����ͣ���ʱ��"
                            strMessage = ""
                            OperationEx = True
                            Exit Function
                        End If
                    End If
                End If
            End With
        Case 4      '������
            Set objThird = New clsINF_Grani
            With objThird
                Set .mobjComLib = objBase.mobjComLib
                Set .mobjLog = objBase.mobjLog
                .URL = mdlDrugMachine.GetInterfaceLink(objBase, strCode)    'WebService��ַ
                .UserName = mstrUser
                
                '���Ҫ��ȡ���������������
                'If rsData.State = adStateOpen Then rsData.Close
                'Set rsData = .GetData(intType, strData, blnOK)
            End With
        Case 5      'YUYAMA
            Set objThird = New clsINF_YUYAMA
            With objThird
                Set .mobjComLib = objBase.mobjComLib
                Set .mobjLog = objBase.mobjLog
                .URL = mdlDrugMachine.GetInterfaceLink(objBase, strCode)    'OLEDB���Ӵ�
                .UserName = mstrUser
                
                '���Ҫ��ȡYUYAMA����������
                'If rsData.State = adStateOpen Then rsData.Close
                'Set rsData = .GetData(intType, strData, blnOK)
            End With
        Case Else
            strMessage = "δ֪�Ľӿ�����"
            mclsLog.Add strMessage, 1
            Exit Function
        End Select
    End If

    '���浥����Ϣ
    Select Case intType
    Case 21, 23
        objThird.mstrNOs = Split(strData, "|")(1)
    Case 24, 25
        objThird.mstrNOs = strData
    Case Else
        objThird.mstrNOs = ""
    End Select
    
    '�ӿڿ���
    objThird.Action = True
    If objThird.Action = False Then Exit Function
    
    '���˶�Ӧ�ӿڱ�ŵ�����
    If objThird.FilterData(strCode, intType, rsData, rsFinish) = False Then Exit Function
    
    '��������   mintType��Ӧ�����intType��ҵ������
    OperationEx = objThird.SendData(mintType, intType, intIO, strCode, rsFinish, strMessage)
    
    '����
    objThird.Action = False
    
End Function

Public Function Init(ByVal intType As Byte, ByVal objControl As Object, ByRef strMessage As String) As Boolean
'���ܣ���ʼ��
'������
'  intType��Ӧ�����1-ZLHIS��2-�ӿڹ������ߣ���ʱ��������ӿڴ������ݣ���3-�ƶ���֧��������4-��������ƽ̨
'  objControl��clsComLib�������Ӧ�����Ϊ��3-֧���������봫��ADO�����Ӷ���
'  strMessage��������Ϣ
'���أ�True�ɹ���Falseʧ��

    Dim strFile As String, strTmp As String

    mblnInit = False
    mintType = intType
    strMessage = ""
    
    '�������ļ�������ר�ã�
    If LCase(App.Path) Like "*\apply" Then
        strFile = App.Path & "\" & GSTR_CONFIG_FILE
    ElseIf LCase(App.Path) Like "*\apply\*" Then
        strFile = Left(App.Path, InStr(LCase(App.Path), "\apply\") + Len("\apply\") - 1) & GSTR_CONFIG_FILE
    ElseIf LCase(App.Path) Like "*zldrugmachinemanage*" Or LCase(App.Path) Like "*zldrugmachine\*" Or LCase(App.Path) Like "*zldrugmachine" Then
        strFile = Replace(App.Path, "\" & App.EXEName, "") & "\" & App.EXEName & "\zlDrugMachineManage\" & GSTR_CONFIG_FILE
    Else
        strMessage = "��" & App.Path & "\" & GSTR_CONFIG_FILE & "���ļ�·��������"
        Exit Function
    End If
    
    If mdlDrugMachine.VerifyConfigFile(strFile) = False Then
        strMessage = "������" & GSTR_CONFIG_FILE & "�������ļ�ʧ��"
    End If
    If strMessage <> "" Then Exit Function
    
    '��ȡ����
    Call mdlDrugMachine.ReadParams(mtypParams)
    
    '������־���������
    With mclsLog
        .Action = mtypParams.�����־
        .Type_ = IIf(mtypParams.��ϸ��־, 1, 0)
        .SaveDays = mtypParams.������־����
    End With
    
    '��־
    mclsLog.AddTime
    If strMessage <> "" Then mclsLog.Add strMessage
    mclsLog.Add "���²������"
    
    '������
    If intType < 1 Or intType > 4 Then
        strMessage = "Ӧ�����δ����"
        mclsLog.Add strMessage
        mclsLog.Save
        Exit Function
    End If
    If objControl Is Nothing Then
        strMessage = "����δ����"
        mclsLog.Add strMessage
        mclsLog.Save
        Exit Function
    End If
    
    '��ʼ������
    On Error GoTo hErr
    If intType = Val("3-֧����") Then
        Set mobjComLib = New clsPublic
        If TypeName(objControl) <> "Connection" Then
            strMessage = "֧����δ�����������Ӷ���"
            mclsLog.Add strMessage
            mclsLog.Save
            Exit Function
        End If
        strMessage = "���á�zl9CommLib.InitCommon������ʧ�ܣ�����zl9ComLib�����汾��"
        mobjComLib.InitCommon objControl
    Else
        If TypeName(objControl) <> "clsComLib" Then
            strMessage = "δ����clsComLib����"
            mclsLog.Add strMessage
            mclsLog.Save
            Exit Function
        End If
    
        Set mobjComLib = objControl
        strMessage = "���á�zl9CommLib.CurrentConn������ʧ�ܣ�����zl9ComLib�����汾��"
        If mobjComLib.CurrentConn Is Nothing Then
            strMessage = "zlDrugMachine�������������Ӵ���"
            mclsLog.Add strMessage
            mclsLog.Save
            Exit Function
        End If
    End If
    strMessage = ""
    
    '��ȡϵͳ��ģ�����
    mclsLog.Add "��ȡģ�����"
    If intType = Val("3-֧����") Then
        strTmp = mobjComLib.GetPara("������Ϣ����ƽ̨", GLNG_SYS, GLNG_MODULE)
        mblnDrugMachine = Val(mobjComLib.GetPara("����ҩƷ�Զ����豸�ӿ�", GLNG_SYS, GLNG_MODULE)) = 1
        mstrKey = mobjComLib.GetPara("��Ϣ����ƽ̨��Կ", GLNG_SYS, GLNG_MODULE)
        mstrToken = mobjComLib.GetPara("��Ϣ����ƽ̨����", GLNG_SYS, GLNG_MODULE)
    Else
        strTmp = mobjComLib.zlDatabase.GetPara("������Ϣ����ƽ̨", GLNG_SYS, GLNG_MODULE)
        mblnDrugMachine = Val(mobjComLib.zlDatabase.GetPara("����ҩƷ�Զ����豸�ӿ�", GLNG_SYS, GLNG_MODULE)) = 1
        mstrKey = mobjComLib.zlDatabase.GetPara("��Ϣ����ƽ̨��Կ", GLNG_SYS, GLNG_MODULE)
        mstrToken = mobjComLib.zlDatabase.GetPara("��Ϣ����ƽ̨����", GLNG_SYS, GLNG_MODULE)
    End If
    mblnIIP = Val(strTmp) = 1
    mstrIIP_URL = Mid(strTmp, InStr(strTmp, "|") + 1)
    
    If mblnDrugMachine = False Then
        strMessage = "δ����ҩƷ�Զ����豸�ӿ�"
        mclsLog.Add strMessage
        mclsLog.Save
        Exit Function
    End If
    
    '��ʼ�����
    mblnInit = True
    Init = True
    
    mclsLog.Add "��ʼ�����"
    mclsLog.Save
    
    Exit Function
    
hErr:
    mclsLog.Add IIf(strMessage = "", Err.Number & "��" & Err.Description, strMessage)
    mclsLog.Save
End Function

Private Sub Class_Initialize()
    With mclsLog
        .Path = App.Path
        .FileName = App.EXEName
    End With
End Sub

Private Sub Class_Terminate()
    Dim frmThis As Form
    
    mclsLog.Save
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next

    Set mclsLog = Nothing
    Set mobjComLib = Nothing
    Set gcnThird = Nothing
End Sub

Private Function GetINFCodes(ByVal rsData As ADODB.Recordset) As String
'���ܣ��ӱ�׼��¼������õ����нӿڱ��
'������
'  rsData����׼��¼������
'���أ��ӿڱ��1[;�ӿڱ��2[...]]

    Dim l As Long
    Dim strTmp As String, strResult As String
    
    With rsData
        If .State <> adStateOpen Then Exit Function
        
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            If mdlDrugMachine.NVL(!�ӿڱ��) <> "" Then
                strTmp = Trim(!�ӿڱ��)
                If Not ";" & strResult & ";" Like "*;" & strTmp & ";*" Then
                    strResult = strResult & ";" & strTmp
                End If
            End If
            
            .MoveNext
        Loop
        If Left(strResult, 1) = ";" Then strResult = Mid(strResult, 2)
    End With
    
    GetINFCodes = strResult

End Function

Private Function GetINFType(ByVal strCode As String) As Integer
'���ܣ���ȡָ���ӿڱ�ŵĽӿ�����
'������
'  strCode���ӿڱ��
'���أ��ӿ����ͣ�1-Τ�ֺ��ģ�2-TOSHO��3-���ͣ�4-�����᣻5-YUYAMA

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    
    strSQL = "Select ���� From ҩƷ�豸�ӿ� Where ��� = [1] "
    
    If mintType = Val("3-֧����") Then
        Set rsTmp = mobjComLib.OpenSQLRecord(strSQL, "��ȡ�ӿڵ�����", strCode)
    Else
        Set rsTmp = mobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ӿڵ�����", strCode)
    End If
    If rsTmp.EOF = False Then
        GetINFType = IIf(IsNull(rsTmp!����), 0, rsTmp!����)
    End If
    rsTmp.Close
    
    Exit Function

hErr:
    mclsLog.Add Err.Number & "��" & Err.Description
    mclsLog.Save
End Function