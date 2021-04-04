VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicPeis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
'��;������ⲿ�ӿ�
'������Initialize,WriteLisResult,GetLastError
'######################################################################################################################

Private mstrError As String
Private mobjPeisPDFReport As clsPeisPDFReport
Private mobjTaskPersonReport As frmTaskPersonReport

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'���ԣ�PeisPDFReport
Public Property Get PeisPDFReport() As clsPeisPDFReport
    If mobjPeisPDFReport Is Nothing Then
        Set mobjPeisPDFReport = New clsPeisPDFReport
        Call mobjPeisPDFReport.Initialize(2100, 2121)
    End If
    Set PeisPDFReport = mobjPeisPDFReport
End Property

Public Function Initialize(ByVal cnOracle As ADODB.Connection) As Boolean
    '******************************************************************************************************************
    '���ܣ������ĳ�ʼ��
    '������cnOracle:����������ݿ�����
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    ParamInfo.ϵͳ���� = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    ParamInfo.��Ʒ���� = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    mstrError = ""
    Set gobjComLib = Nothing
    Set gobjDatabase = Nothing
    
    Set gcnOracle = cnOracle
    
    On Error Resume Next
    Err = 0
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    On Error GoTo errHand
    
    If Err <> 0 Then
        mstrError = "������������(zl9ComLib.clsComLib)ʧ�ܡ�"
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    Set gobjComFun = CreateObject("zl9ComLib.clsCommFun")
    On Error GoTo errHand
    If Err <> 0 Then
        mstrError = "������������(zl9ComLib.clsCommFun)ʧ�ܡ�"
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    Set gobjDatabase = CreateObject("zl9ComLib.clsDatabase")
    On Error GoTo errHand
    
    If Err <> 0 Then
        mstrError = "������������(zl9ComLib.clsDatabase)ʧ�ܡ�"
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    Set gobjReport = CreateObject("zl9Report.clsReport")
    On Error GoTo errHand
    
    If Err <> 0 Then
        mstrError = "������������(zl9Report.clsReport)ʧ�ܡ�"
        Exit Function
    End If
    
    
    Call gobjComLib.InitCommon(cnOracle)
    
    Initialize = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

'******************************************************************************************************************
'���ܣ�����һ������id���ӿڷ���,����������˽����ж��Ƿ��б���
'������lngPatientKey ����id
'���أ�true--�ò����б���  false--�ò���û�б���
'******************************************************************************************************************
Public Function HasExaminationReport(ByVal lngPatientKey As Long) As Boolean
    Dim rsData As ADODB.Recordset
    Dim strPrivs As String
    
    On Error GoTo errHandle
    
    '�жϵ�ǰ����Ա�Ƿ��б������Ȩ��
    strPrivs = gobjComLib.GetPrivFunc(2100, 2150)
    If Not IsPrivs(strPrivs, "��챨�����") Then
       HasExaminationReport = False
       Exit Function
    End If
    Set rsData = gclsPackage.Get_PeisPersonReportBill(lngPatientKey)
    If rsData.BOF = False Then
        HasExaminationReport = True
        Exit Function
    End If
    HasExaminationReport = False
    
    Exit Function
errHandle:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

'******************************************************************************************************************
'���ܣ����˱������
'������frmParent �������
'      lngPatientKey ����id
'���أ�
'******************************************************************************************************************
Public Function OpenExaminationReport(ByVal frmParent As Object, ByVal lngPatientKey As Long) As Boolean
    Dim objPersonReport As Object
    Dim strPrivs As String
    
    '�жϵ�ǰ����Ա�Ƿ��б������Ȩ��
    strPrivs = gobjComLib.GetPrivFunc(2100, 2150)
    If Not IsPrivs(strPrivs, "��챨�����") Then
        
        Call ShowSimpleMsg("��û����챨����ĵ�Ȩ��,���ܽ��б�����ġ�")
        OpenExaminationReport = False
        Exit Function
    End If
    
    Set objPersonReport = New frmPersonReport
    
    OpenExaminationReport = objPersonReport.ShowReportForm(frmParent, lngPatientKey)
End Function

'******************************************************************************************************************
'���ܣ���ȡ���������¼��
'������lngPatientKey ����id
'���أ�
'******************************************************************************************************************
Public Function GetTaskExaminationReportRecord(ByVal lngPatientKey As Long) As ADODB.Recordset
    Dim objPersonReport As Object
    Dim strPrivs As String
    
    '�жϵ�ǰ����Ա�Ƿ��б������Ȩ��
    strPrivs = gobjComLib.GetPrivFunc(2100, 2150)
    If Not IsPrivs(strPrivs, "��챨�����") Then
        
        Call ShowSimpleMsg("��û����챨����ĵ�Ȩ��,���ܽ��б�����ġ�")
        Set GetTaskExaminationReportRecord = Nothing
        Exit Function
    End If
    Set GetTaskExaminationReportRecord = gclsPackage.Get_PeisPersonReportBill(lngPatientKey)
    
End Function


'******************************************************************************************************************
'���ܣ���ȡ������Ա���洰�����
'������
'���أ��������
'******************************************************************************************************************
Public Function GetTaskPersonExaminationReportForm() As Object
    Dim strPrivs As String
    
    '�жϵ�ǰ����Ա�Ƿ��б������Ȩ��
    strPrivs = gobjComLib.GetPrivFunc(2100, 2150)
    If Not IsPrivs(strPrivs, "��챨�����") Then
        
        Call ShowSimpleMsg("��û����챨����ĵ�Ȩ��,���ܽ��б�����ġ�")
        Set GetTaskPersonExaminationReportForm = Nothing
        Exit Function
    End If
    
    If mobjTaskPersonReport Is Nothing Then
       Set mobjTaskPersonReport = New frmTaskPersonReport
       Call mobjTaskPersonReport.InitData
    End If
    Set GetTaskPersonExaminationReportForm = mobjTaskPersonReport
    
End Function

'******************************************************************************************************************
'���ܣ�����������Ա������Ϣ
'������lngTaskKey ����ID
'      lngPatientKey ����id
'���أ�
'******************************************************************************************************************
Public Function LoadTaskPersonExaminationReportDetail(ByVal lngTaskKey As Long, ByVal lngPatientKey As Long) As Boolean
    Dim strPrivs As String
    
    '�жϵ�ǰ����Ա�Ƿ��б������Ȩ��
    strPrivs = gobjComLib.GetPrivFunc(2100, 2150)
    If Not IsPrivs(strPrivs, "��챨�����") Then
        
        Call ShowSimpleMsg("��û����챨����ĵ�Ȩ��,���ܽ��б�����ġ�")
        LoadTaskPersonExaminationReportDetail = False
        Exit Function
    End If
    
    If mobjTaskPersonReport Is Nothing Then
       Set mobjTaskPersonReport = New frmTaskPersonReport
       Call mobjTaskPersonReport.InitData
    End If
    
    LoadTaskPersonExaminationReportDetail = mobjTaskPersonReport.LoadData(lngTaskKey, lngPatientKey)
    
End Function
    
Public Function WriteLisResult(ByVal strVal As String, ByVal intType As Integer, Optional ByRef strErr As String) As Boolean
    '******************************************************************************************************************
    '���ܣ��Ѽ�������д�������
    '����           strVal ר��ı걾�������
    '               ����(1=��ͨ)<split2>����ID<split2>������Դ<split2>����ʱ��<split2>������<split2>�����<split2>���ʱ��<split2>����Ŀ����<split2>�걾����<split2>
    '                   ָ��1<split4>������1<split4>��λ1<split4>�����־1<split4>�������1<split4>�������1<split4>��˽��Ŀ1<split4>ָ�����1<split3>
    '                   ָ��2<split4>������2<split4>��λ2<split4>�����־2<split4>�������2<split4>�������2<split4>��˽��Ŀ2<split4>ָ�����2<split3>
    '                   ָ��3<split4>������3<split4>��λ3<split4>�����־3<split4>�������3<split4>�������3<split4>��˽��Ŀ3<split4>ָ�����3<split1>
    '
    '               ����(2=΢����)<split2>����ID<split2>������Դ<split2>����ʱ��<split2>������<split2>�����<split2>���ʱ��<split2>����Ŀ����<split2>�걾����<split2>
    '               ϸ����1<split3>����1<split3>��ҩ����1<split3>
    '                   ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
    '                   ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split2>
    '               ϸ����2<split3>����2<split3>��ҩ����2<split3>
    '                   ������1<split4>�����ؽ��1<split4>��ҩ��1<split4>ҩ������1<split4>�÷�����11<split4>�÷�����21<split4>ѪҩŨ��11<split4>ѪҩŨ��21<split4>��ҩŨ��11<split4>��ҩŨ��21<split3>
    '                   ������2<split4>�����ؽ��2<split4>��ҩ��2<split4>ҩ������2<split4>�÷�����12<split4>�÷�����22<split4>ѪҩŨ��12<split4>ѪҩŨ��22<split4>��ҩŨ��12<split4>��ҩŨ��22<split1>
    '
    '               intType 0=��� 1=ȡ�����
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Dim intSample As Integer                    'ѭ������
    Dim astrSample() As String                  'ѭ���걾����
    Dim strSQL As String
    Dim lngState As Long '��ȡ����״̬
    Dim lngҽ��id As Long
    Dim strSQLtmp As String
    Dim rs As ADODB.Recordset
    Dim varMsg As VbMsgBoxResult
    Dim blnBegin As Boolean

    If intType = 1 Then
        lngState = Val(gobjDatabase.GetPara("LISȡ���������", 2100, 0, "0"))
    End If
    strSQLtmp = "Select Sum(����) As ����" & vbNewLine & _
                "From (Select Count(1) As ����" & vbNewLine & _
                "       From ���������Ա A, ��������� B" & vbNewLine & _
                "       Where a.����id = b.����id And a.����id = b.����id And a.���״̬ = 3 And b.�ɼ�ҽ��id = [1] And Rownum < 2" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Count(1) As ����" & vbNewLine & _
                "       From ���������� A, ��������� B" & vbNewLine & _
                "       Where a.����id = b.����id And a.����id = b.����id And a.�嵥id = b.�嵥id And a.�����Ŀid = b.�����Ŀid And a.�������� Is Not Null And" & vbNewLine & _
                "             b.�ɼ�ҽ��id = [1] And Rownum < 2" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Count(1) As ����" & vbNewLine & _
                "       From ���������� A, ��������� B" & vbNewLine & _
                "       Where a.Id = b.����id And b.�ɼ�ҽ��id = [1] And Rownum < 2)"
    astrSample = Split(strVal, "<split1>")
    For intSample = 0 To UBound(astrSample)
        If Trim(astrSample(intSample)) <> "" Then
            If intType = 1 Then
                lngҽ��id = Val(Split(astrSample(intSample), "<split2>")(1))
                Set rs = gobjDatabase.OpenSQLRecord(strSQLtmp, "��ȡ���������Ϣ", lngҽ��id)
                If rs.BOF = False And blnBegin = False Then
                    If IIf(IsNull(rs("����").Value), 0, rs("����").Value) Then
                        Select Case lngState
                            Case 0 '��ʾ
                                varMsg = MsgBox("��ǰ��Ա���������ܼ�������ܼ�״̬�����ߵ�ǰ��Ա�Ѿ������ɣ�Ҳ�п��ܵ�ǰ��Ա����Ӧ�ļ�����Ŀ����д����ĿС�ᣬ��ȷ��Ҫȡ��������ˣ�" & vbCrLf & _
                                              "(ע:ȡ��������ˣ�����ܵ�����ĿС������Ŀ�����ƥ�䣡Ҳ�п��ܵ�������д���ܼ���۲�׼ȷ��)", vbYesNo + vbQuestion + vbDefaultButton2)
                                If varMsg = vbNo Then
                                    Exit Function
                                Else
                                    blnBegin = True
                                End If
                            Case 1 '��ֹ
                                MsgBox ("��ǰ��Ա�����ܼ죬�������ܼ죬���ߵ�ǰ��Ա�Ѿ������ɣ����߸���Ŀ�Ѿ�����С�ᡣ�޷�ȡ��������ˣ�" & vbCrLf & _
                                        "(��ʾ��ȷʵ��Ҫȡ��������ˣ�����Ҫ�������²���:" & vbCrLf & _
                                        "       ���鵱ǰ��Ա�����״̬�Ƿ�Ϊ������,�����,��ȡ��������;" & vbCrLf & _
                                        "       ���鵱ǰ��Ա�Ƿ������ܼ�������ܼ�,�����,����Ҫ����ǰ��Ա�ָ������ܼ�������;" & vbCrLf & _
                                        "       ���鵱ǰ��Ա����Ӧ�ļ�����Ŀ�Ƿ���д����Ӧ����ĿС��,�����,��������)")
                                Exit Function
                        End Select
                    End If
                End If
            End If
            strSQL = "zl_���������_FROMLISTOTJ('" & astrSample(intSample) & "'," & intType & ")"
           Call gobjDatabase.ExecuteProcedure(strSQL, "����")
        End If
    Next
    WriteLisResult = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    strErr = Err.Description
    mstrError = Err.Description
End Function

Public Function GetLastError() As String
    '******************************************************************************************************************
    '���ܣ���ȡ������Ϣ
    '������
    '���أ�
    '******************************************************************************************************************
        
    GetLastError = mstrError

End Function
'
Private Sub Class_Initialize()
    mstrError = ""
End Sub

Private Sub Class_Terminate()
     If Not (gcnOracle Is Nothing) Then Set gcnOracle = Nothing
     If Not (gobjComLib Is Nothing) Then Set gobjComLib = Nothing
     If Not (gobjComFun Is Nothing) Then Set gobjComFun = Nothing
     If Not (gobjDatabase Is Nothing) Then Set gobjDatabase = Nothing
     If Not (gclsPackage Is Nothing) Then Set gclsPackage = Nothing
     If Not (gobjReport Is Nothing) Then Set gobjReport = Nothing
     
     Set mobjPeisPDFReport = Nothing
     Set mobjTaskPersonReport = Nothing
End Sub

