Attribute VB_Name = "mdlɽ��"
Option Explicit

'EmployeeInfo
Private Type ���˻�����Ϣ
    ���˱��00     As String  '01
    ���֤��01     As String  '02
    ����02         As String  '03
    �Ա�03         As String  '04
    ��������04     As String  '05
    ����05         As String  '06
    ҽ��֤��06     As String  '07
    ��λ���07     As String  '08
    ҽ����Ա���08 As String  '09
    ����Ա��־09   As String  '10
    �չ���Ա��־10 As String  '11
  ' ��չ12       As String
  ' ��չ13       As String
  ' ��չ14       As String
  ' ��չ15       As String
    ��Ժ״̬16     As String  '16
    ����          As String
End Type

 '(AccountInfo):
Private Type �ʻ�������Ϣ
    �ʻ�������00           As Currency   '01
    �ʻ����01               As String  '02
    ����סԺ����02           As Integer    '03
    �����ܷ���֧���ۼ�03     As Currency   '04
'   ��չ05                 As String
    �����Է��ۼ�05           As Currency   '06
    ���������ۼ�06           As Currency   '07
    �������ͳ���ۼ�07       As Currency   '08
    �����ʻ�֧���ۼ�08       As Currency   '09
'   ��չ10                 As String
    ����ͳ��֧���ۼ�10       As Currency   '11
    �����ֽ�֧���ۼ�11       As Currency   '12
'   ��չ13                 As String
    ���깫��Ա����֧���ۼ�13 As Currency    '14
'   ��չ15                 As String
'   ��չ16                 As String
'   ��չ17                 As String
'   ��չ18                 As String
'   ��չ19                 As String
'   ��չ20                 As String
'   ��չ21                 As String
'   ��չ22                 As String
End Type

Type ������Ϣ_ɽ��

'|�����ܶ�|�����ʻ�֧��|ͳ��֧��|�ֽ�֧��|����Ա����֧��
'|������|�Էѽ��|סԺ�˴�|�𸶱�׼|תԺ���Ʒ���
'|�𸶱�׼�Ը�|�𸶱�׼����Ա֧��|�ֶ�1ͳ��֧��|�ֶ�1����Ա֧��|�ֶ�1�����Ը�
'|�ֶ�2ͳ��֧��|�ֶ�2����Ա֧��|�ֶ�2�����Ը�|�ֶ�3ͳ��֧��|�ֶ�3����Ա֧��
'|�ֶ�3�����Ը�|���ⶥ����Ա֧��|����Ա�����Ը�|������Ա�ⶥ�����Ը�|������ϵ
'|��λ����|

    ����ID        As Long
    ҽ����        As Long
    �����ܶ�      As Double
    �����ʻ����  As Double
    �����ʻ�      As Double
    ͳ�����      As Double
    ����Ա����    As Double
    ������      As Double
    �Էѽ��      As Double
    
    
End Type

Public g���˻�����Ϣ As ���˻�����Ϣ
Public g�ʻ�������Ϣ As �ʻ�������Ϣ
Public gcnSxDr As New ADODB.Connection 'ҽ������
Public g������Ϣ_ɽ�� As ������Ϣ_ɽ��

Private clsDR As Object   '������
Private mblnInit As Boolean '�Ƿ��ʼ��
Private mlngReturn As Long  '�ӿڷ���ֵ
Private mstrInput As String  '���
Private mstrOutput As String '����
Private mrsTMP As New ADODB.Recordset  '��ʱ��¼��
Private mstrSQL As String    '��ʱ���SQL���

Public Declare Function InitDLL Lib "DBLIB.DLL" (ByVal intType As Long) As Long
Public Declare Function CommitTrans Lib "DBLIB.DLL" () As Long
Public Declare Function RollbackTrans Lib "DBLIB.DLL" () As Long

Public Declare Function CheckMTQ Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal strPersonNO As String, _
    ByVal strWorkunitNo As String, ByVal strMedKind As String, ByVal strSysdate As String, ByVal strDataBuffer As String) As Long
    
Public Declare Function ReadCard Lib "DBLIB.DLL" (ByVal strEmployeeInfo As String, ByVal strAccountInfo As String, _
    ByVal strDataBuffer As String, ByVal strPin As String, Optional ByVal strDestCardNO As String = vbNullString) As Long
    
Public Declare Function Registration Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal intRegType As Long, _
    ByVal strInHosNo As String, ByVal strApprNO As String, ByVal strMedType As String, _
    ByVal strDiseaseNO As String, ByVal strDiseaseName As String, ByVal strLHStatus As String, ByVal MainDocName As String, _
    ByVal strApprPerson As String, ByVal strTransactor As String, ByVal strTransDate As String, _
    ByVal strDataBuffer As String) As Long
    
Public Declare Function TreatInfoEntry Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal intRegType As Long, _
    ByVal strMedType As String, ByVal strInHosNo As String, ByVal strApprNO As String, _
    ByVal strTreatDate As String, ByVal strLeaveHosDt As String, _
    ByVal strDiseaseNO As String, ByVal strDiseaseName As String, _
    ByVal strLHDiseaseNO As String, ByVal strLHDiseaseName As String, ByVal strLHStatus As String, _
    ByVal strTrunHosKind As String, ByVal strmainDocName As String, _
    ByVal strApprPerson As String, ByVal strTransactor As String, ByVal strTransDate As String, _
    ByVal strDataBuffer As String) As Long
    
Public Declare Function FormularyEntry Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal strInHosNo As String, _
    ByVal intTransKind As Long, ByVal intItemKind As Long, ByVal strInternalCode As String, _
    ByVal strFormularyNo As String, ByVal strSysdate As String, ByVal strCenterCode As String, _
    ByVal strItemName As String, ByVal dblUnitPrice As Double, ByVal dblQuantity As Double, _
    ByVal dblAmount As Double, ByVal strDoseType As String, ByVal strDosage As String, _
    ByVal strFrequency As String, ByVal strUsage As String, ByVal strKeBie As String, _
    ByVal strExecDays As String, ByVal strFeeType As String, ByVal strDoctName As String, _
    ByVal strTransactor As String, ByVal strApprPerson As String, ByVal intIsOwnExpenses As Long, _
    ByVal strDataBuffer As String) As Long
    
Public Declare Function ExpenseCalc Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal intTransType As Long, _
    ByVal intInvoiceKind As Long, ByVal strInHosNo As String, ByVal strMedType As String, _
    ByVal strInvoiceNo As String, ByVal strUserName As String, ByVal dblAccCashPay As Double, _
    ByVal strDataBuffer As String) As Long
    
Public Declare Function PreExpenseCalc Lib "DBLIB.DLL" (ByVal strCardNO As String, ByVal strInHosNo As String, _
    ByVal strMedType As String, ByVal strDataBuffer As String) As Long
    
Public Declare Function ChangePinEx Lib "DBLIB.DLL" (ByVal strszOldPin As String, ByVal strszNewPin As String, _
    ByVal strDataBuffer As String) As Long

Public Function ҽ����ʼ��_ɽ��() As Boolean
  Dim strUser As String, strServer As String, strPass As String  'ҽ��������
    On Error GoTo errHand
  
    If mblnInit = False Then
        mstrSQL = "Select * From ���ղ��� Where ����=" & TYPE_ɽ��
        Set mrsTMP = gcnOracle.Execute(mstrSQL)
        Do Until mrsTMP.EOF
            Select Case mrsTMP!������
                Case "ҽ���û���"
                    strUser = IIf(IsNull(mrsTMP("����ֵ")), "", mrsTMP("����ֵ"))
                Case "ҽ��������"
                    strServer = IIf(IsNull(mrsTMP("����ֵ")), "", mrsTMP("����ֵ"))
                Case "ҽ���û�����"
                    strPass = IIf(IsNull(mrsTMP("����ֵ")), "", mrsTMP("����ֵ"))
                Case "ҽԺ�ȼ�"
            End Select
            mrsTMP.MoveNext
        Loop
        
        If OraDataOpen(gcnSxDr, strServer, strUser, strPass, False) = False Then
            MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        'Set clsDR = CreateObject("sxdr.clssxdr")
        
        mlngReturn = InitDLL(1)
        Call WriteBusinessLOG("InitDll", "", "")
        
        If mlngReturn = 0 Then
            mblnInit = True
            ҽ����ʼ��_ɽ�� = True
        Else
            MsgBox "��ʼ��ʧ��!", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        ҽ����ʼ��_ɽ�� = True
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �������(strInPin As String, Optional ByVal strCarNO As String = "NULL") As Boolean
'EmployeeInfo    OUT ���˻�����Ϣ
'AccountInfo     OUT �ʻ�������Ϣ
'DataBuffer      OUT ������Ϣ
'Pin IN  У��Pin��������
'DestCardNo  IN [OPTION] ���ţ� �� (Ĭ��ֵ):��ʵ���������ض�����Ϣ
'                              �ǿ�:ģ����������˿��Ŵӿ��ж��������Ϣ��
'                              ����Ҫ���ڲ�����ʱ��סԺת����Ԥ�ᣩ
    'ҽ����Ա���
    
    '����  ��Ա���
    '11     ��ְ
    '21     ����
    '33     �����Ҽ��˲о���
    '91     ������Ա
    
    '����˵�������ޣ��Զ��壩
    '0   ��
    '1   Ů
    
    '�˴�����ʧ�ܲ��ع����ⲿ����ʱ���Ѵ˺������ɿ�ʼ�����ⲿ���ô����ݷ��صĽ�����лع�

If mblnInit = False Then
    MsgBox "���ȳ�ʼ��ҽ���ӿڣ�", vbInformation, gstrSysName
     Exit Function
End If

Dim retEmpInfo As String
Dim retAccInfo As String
Dim str�������� As String
retEmpInfo = Space(600)
retAccInfo = Space(600)
mstrOutput = Space(600)

If strCarNO = "" Or strCarNO = "NULL" Then strCarNO = vbNullString

mlngReturn = ReadCard(retEmpInfo, retAccInfo, mstrOutput, strInPin, strCarNO)
Call WriteBusinessLOG("ReadCard", strInPin & "��" & strCarNO, Trim(mstrOutput))

If mlngReturn = 0 Then
    g���˻�����Ϣ.���˱��00 = Split(retEmpInfo, "|")(1)
    g���˻�����Ϣ.���֤��01 = Split(retEmpInfo, "|")(2)
    g���˻�����Ϣ.����02 = Split(retEmpInfo, "|")(3)
    g���˻�����Ϣ.�Ա�03 = IIf(Split(retEmpInfo, "|")(4) = "0", "��", "Ů")
    g���˻�����Ϣ.��������04 = Split(retEmpInfo, "|")(5)
    g���˻�����Ϣ.����05 = Split(retEmpInfo, "|")(6)
    g���˻�����Ϣ.ҽ��֤��06 = Split(retEmpInfo, "|")(7)
    g���˻�����Ϣ.��λ���07 = Split(retEmpInfo, "|")(8)
    g���˻�����Ϣ.ҽ����Ա���08 = Split(retEmpInfo, "|")(9)
    g���˻�����Ϣ.����Ա��־09 = Split(retEmpInfo, "|")(10)
    g���˻�����Ϣ.�չ���Ա��־10 = Split(retEmpInfo, "|")(11)
    g���˻�����Ϣ.��Ժ״̬16 = Split(retEmpInfo, "|")(17)
    g���˻�����Ϣ.���� = strInPin
    
    g�ʻ�������Ϣ.�ʻ�������00 = Split(retAccInfo, "|")(1)
    g�ʻ�������Ϣ.�ʻ����01 = Split(retAccInfo, "|")(2)
    g�ʻ�������Ϣ.����סԺ����02 = Split(retAccInfo, "|")(3)
    g�ʻ�������Ϣ.�����ܷ���֧���ۼ�03 = Split(retAccInfo, "|")(4)
    g�ʻ�������Ϣ.�����Է��ۼ�05 = Split(retAccInfo, "|")(6)
    g�ʻ�������Ϣ.�������ͳ���ۼ�07 = Split(retAccInfo, "|")(8)
    g�ʻ�������Ϣ.�����ʻ�֧���ۼ�08 = Split(retAccInfo, "|")(9)
    g�ʻ�������Ϣ.����ͳ��֧���ۼ�10 = Split(retAccInfo, "|")(11)
    g�ʻ�������Ϣ.�����ֽ�֧���ۼ�11 = Split(retAccInfo, "|")(12)
    g�ʻ�������Ϣ.���깫��Ա����֧���ۼ�13 = Split(retAccInfo, "|")(14)
    
   
    With g���˻�����Ϣ
        mstrOutput = Space(600)
        mlngReturn = CheckMTQ(.����05, .���˱��00, .��λ���07, .ҽ����Ա���08, Format(zlDatabase.Currentdate, "yyyyMMdd"), mstrOutput)
        Call WriteBusinessLOG("CheckMTQ", .����05 & "," & .���˱��00 & "," & .��λ���07 & "," & .ҽ����Ա���08 & "," & Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
    End With
    If mlngReturn = 0 Then
        str�������� = Split(mstrOutput, "|")(1)
        If Val(str��������) = 0 Then
            ������� = True
        ElseIf (Val(str��������) >= 1 And Val(str��������) <= 30) Or Val(str��������) = 42 Then
            ������� = False
            MsgBox "��������������ִ��ҽ�����ף�" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
            Call ����_ɽ��
        ElseIf Val(str��������) > 30 And Val(str��������) <> 42 Then
            ������� = True
            MsgBox "�������ַ�����������ȫ����ҽ��������" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
        End If
    Else
        ������� = False
        MsgBox "�������ʧ�ܣ�" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
        Call ����_ɽ��
    End If
Else
    ������� = False
    MsgBox "����ʧ�ܣ�" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
End If

End Function

Public Function ��ݱ�ʶ_ɽ��(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim str������� As String
    Dim str����� As String
    Dim str���� As String
    
    If Not (bytType = 1 Or bytType = 0) Then Exit Function  '�������շѣ���Ժ�Ǽǲŵ���
    
    Dim strIdeReturn As String  '���շ�����Ϣ
    
    strIdeReturn = frmIdentifyɽ��.��ݱ�ʶ(bytType, lng����ID) '''����ӿ�����,readcard
    
    If strIdeReturn = "-1" Then
        ��ݱ�ʶ_ɽ�� = ""
    Else
        mstrSQL = "Select * from �����ʻ� where ����ID=[1] and ����=[2]"
        Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "������Ϣ", lng����ID, TYPE_ɽ��)
        
        If mrsTMP.EOF Then
            MsgBox "����ɽ��ҽ�����ˣ�������ִ�н��ס�", vbInformation, gstrSysName
            Exit Function
        Else
            If bytType = 0 Then
                str������� = Nvl(mrsTMP!�������, "11")
                str����� = mrsTMP!����ID & "_0_" & IIf(IsNumeric(Nvl(mrsTMP!˳���, 0)), mrsTMP!˳���, 0) + 1
                
                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_ɽ�� & ",'˳���','''" & IIf(IsNumeric(Nvl(mrsTMP!˳���, 0)), Nvl(mrsTMP!˳���, 0), 0) + 1 & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "����˳���")
            Else
                str������� = Nvl(mrsTMP!�������, "21")
            End If
        End If
        
        mstrSQL = "select * from ���ղ��� where ID=(" & _
                        "Select ����ID from �����ʻ� where ����ID=[1]" & _
                                                     " and ����=[2]) " & _
                                           " and ����=[2]"
        Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "������Ϣ", lng����ID, TYPE_ɽ��)
        
        If mrsTMP.EOF Then
            MsgBox "����Ŀ¼����ȷ��������ִ�н��ס�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If bytType = 0 Then
        
                    '������ùҺŵǼǽ���
                    mstrOutput = Space(600)
                    str���� = Format(zlDatabase.Currentdate, "yyyyMMdd")
                    mlngReturn = Registration(g���˻�����Ϣ.����05, _
                                                                        1, _
                                                              str�����, _
                                                            "", _
                                                              str�������, _
                                                               mrsTMP!����, _
                                                               mrsTMP!����, _
                                                                "", "", _
                                                                "", _
                                                             UserInfo.����, _
                                                                str����, _
                                                                mstrOutput)
                                                                
                    Call WriteBusinessLOG("Registration", g���˻�����Ϣ.����05 & "," & _
                                                                             "1," & _
                                                                    str����� & "," & _
                                                                          "," & _
                                                              str������� & "," & _
                                                               mrsTMP!���� & "," & _
                                                               mrsTMP!���� & "," & _
                                                                            "0," & _
                                                                           "," & _
                                                             UserInfo.���� & "," & _
                                Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
                
            If mlngReturn = 0 Then
                �ύ_ɽ��
            Else
                ����_ɽ��
                MsgBox "�ҺŵǼ�ʧ��" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            Call �ύ_ɽ��
        End If
    
        '�ύ_ɽ��
        ��ݱ�ʶ_ɽ�� = strIdeReturn
    End If

End Function

Public Function �ύ_ɽ��()
    CommitTrans
    Call WriteBusinessLOG("ComitTrans", "", "")
End Function

Public Function ����_ɽ��()
    RollbackTrans
    Call WriteBusinessLOG("RollbackTrans ", "", "")
End Function

Public Function �������_ɽ��(strSelfNo As String) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '����: ��ȡ�α����˸����ʻ����
    '����: strSelfNO-���˸��˱��
    '����: ���ظ����ʻ����Ľ��
    '�����������ؼ�ͥ�ʻ���סԺ���ظ����ʻ����
    gstrSQL = "Select Nvl(�ʻ����,0) AS �����ʻ� From �����ʻ� " & _
              " Where ҽ����=[1] and ����=[2]"
              
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", strSelfNo, TYPE_ɽ��)
    �������_ɽ�� = rsTemp!�����ʻ�
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_ɽ��(rs��ϸ��¼ As ADODB.Recordset, str���㷽ʽ As String, Optional str���� As String) As Boolean

Dim lng����ID As Long
Dim str������� As String
Dim str����� As String
Dim str������ϸ��ˮ�� As String
Dim dbl�����ʻ� As Double, dblͳ����� As Double, dbl����Ա���� As Double
Dim rsTmpcd As New ADODB.Recordset '�¶��õ���ʱ���ݼ�
Dim str������Ŀ���� As String
On Error GoTo errHandle
If rs��ϸ��¼.RecordCount = 0 Then
    MsgBox "û�в��˷�����ϸ�����ܽ���ҽ������", vbInformation, gstrSysName
    Exit Function
End If

'����������
lng����ID = rs��ϸ��¼!����ID

mstrSQL = "Select * from �����ʻ� where ����ID=[1] and ����=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "������Ϣ", CLng(rs��ϸ��¼!����ID), TYPE_ɽ��)

If mrsTMP.EOF Then
    MsgBox "����ɽ��ҽ�����ˣ�������ִ�н��ס�", vbInformation, gstrSysName
    Exit Function
Else
    str������� = Nvl(mrsTMP!�������, "11")
    str����� = mrsTMP!����ID & "_0_" & IIf(IsNumeric(Nvl(mrsTMP!˳���, 0)), mrsTMP!˳���, 0)
End If

'������ò�����
If Nvl(str����, 0) <> 9 Then
    If �������(mrsTMP!����, mrsTMP!����) = False Then
        Exit Function
    End If
End If

'�����ϸ
rs��ϸ��¼.MoveFirst
Do Until rs��ϸ��¼.EOF
    'δ��������
    mstrSQL = "select * from ����֧����Ŀ where ����=[1] and �շ�ϸĿID=[2]"
    Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "ҽ����ϸ", TYPE_ɽ��, CLng(rs��ϸ��¼!�շ�ϸĿID))
    If mrsTMP.EOF Then
       ' Call ����_ɽ��
        mstrSQL = "Select * from �շ���ĿĿ¼ where ID=[1]"
        Set rsTmpcd = zlDatabase.OpenSQLRecord(mstrSQL, "�շ���ĿĿ¼", CLng(rs��ϸ��¼!�շ�ϸĿID))
        MsgBox rsTmpcd!���� & "(" & rsTmpcd!���� & ")" & "δ���룡" & vbCrLf & "��������ʹ�ô˹��ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    '�Ե���ҽ�������Ҳ��������
    mstrSQL = "Select aka060 ҽ������,aka061  ��Ŀ����,aka065  ��Ŀ�ȼ�,aka069  �Ը�����,1 ��������,aka063  �շ����" & _
              " From ka02 where aka060='" & mrsTMP!��Ŀ���� & "' and aka061='" & mrsTMP!��Ŀ���� & "' and " & Nvl(mrsTMP!��ע, 0) & "=1" & _
              " union all " & _
              " Select aka090   ҽ������,aka091 ��Ŀ����,aka065 ��Ŀ�ȼ�,aka069 �Ը�����,2 ��������,aka063  �շ����" & _
              " From ka03 where aka090='" & mrsTMP!��Ŀ���� & "' and aka091='" & mrsTMP!��Ŀ���� & "' and " & Nvl(mrsTMP!��ע, 0) & "=2" & _
              " union all " & _
              "Select aka100   ҽ������,aka102 ��Ŀ����,aka103 �����ȼ�,0,3,aka063 �շ���� " & _
              " From ka04 where aka100='" & mrsTMP!��Ŀ���� & "' and aka102='" & mrsTMP!��Ŀ���� & "' and " & Nvl(mrsTMP!��ע, 0) & "=3"
    str������Ŀ���� = mrsTMP!��Ŀ����
    Call OpenRecordset_OtherBase(mrsTMP, "ҽ����ϸ", mstrSQL, gcnSxDr)
    If mrsTMP.EOF Then
      
      '  Call ����_ɽ��
        mstrSQL = "Select * from �շ���ĿĿ¼ where ID=[1]"
        Set rsTmpcd = zlDatabase.OpenSQLRecord(mstrSQL, "�շ���ĿĿ¼", CLng(rs��ϸ��¼!�շ�ϸĿID))
        MsgBox rsTmpcd!���� & "(" & rsTmpcd!���� & ")" & "�������" & vbCrLf & "��˶Ժ���ʹ�ô˹��ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    rs��ϸ��¼.MoveNext
Loop

'����ϸ
'

rs��ϸ��¼.MoveFirst

Do Until rs��ϸ��¼.EOF
    str������ϸ��ˮ�� = zlDatabase.GetNextID("��Ա��")
    mstrSQL = "select * from ����֧����Ŀ where ����=[1] and �շ�ϸĿID=[2]"
    Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "ҽ����ϸ", TYPE_ɽ��, CLng(rs��ϸ��¼!�շ�ϸĿID))
    
    mstrSQL = "Select aka060 ҽ������,aka061  ��Ŀ����,aka065  ��Ŀ�ȼ�,aka069  �Ը�����,1 ��������,aka063  �շ����" & _
              " From ka02 where aka060='" & mrsTMP!��Ŀ���� & "' and aka061='" & mrsTMP!��Ŀ���� & "' and " & Nvl(mrsTMP!��ע, 0) & "=1" & _
              " union all " & _
              "Select aka090   ҽ������,aka091 ��Ŀ����,aka065 ��Ŀ�ȼ�,aka069 �Ը�����,2 ��������,aka063  �շ����" & _
              " From ka03 where aka090='" & mrsTMP!��Ŀ���� & "' and aka091='" & mrsTMP!��Ŀ���� & "' and " & Nvl(mrsTMP!��ע, 0) & "=2" & _
              " union all " & _
              "Select aka100   ҽ������,aka102 ��Ŀ����,aka103 �����ȼ�,0,3,aka063 �շ���� " & _
              " From ka04 where aka100='" & mrsTMP!��Ŀ���� & "' and aka102='" & mrsTMP!��Ŀ���� & "' and " & Nvl(mrsTMP!��ע, 0) & "=3"
  
    Call OpenRecordset_OtherBase(mrsTMP, "ҽ����ϸ", mstrSQL, gcnSxDr)
    
'CardNo ���� N ,InHosNo IN  סԺ��,TransKind,  IN  ��������(1  ����¼�� -1 �˴����ϵ�һ����ϸ)   N
'ItemKind    IN  ��Ŀ���(1:ҩƷ, 2:����, 3������ʩ),InternalCode    IN  �շ���ĿҽԺ�ڱ���  N,FormularyNo IN  �����ţ�ͬһ����Ҫ��֤Ψһ��    N
'SysDate IN  ��������(yyyymmdd)  N,CenterCode  IN  �շ���Ŀҽ�����ı���    N,ItemName    IN  �շ���Ŀ����    N
'UnitPrice   IN  ����    N,Quantity    IN  ����    N,Amount  IN  ���    N,DoseType    IN  ����,Dosage  IN  ����
'Frequency   IN  Ƶ��,Usage   IN  �÷�,KeBie   IN  �Ʊ�����,ExecDays    IN  ִ������,FeeType IN  ҽ�������շ���𣨼���¼��  N
'DoctName    IN  ����ҽ��,Transactor  IN  ������  N,ApprPerson  IN  �����ˣ�����������,IsOwnExpenses   IN  ȫ���Էѱ�־( 0����ȫ���Է� 1��ȫ���Է�)    N

'DataBuffer  OUT ������Ϣ/������Ϣ
'
'����str���� =9���ж��Ƿ��ǽ��㣬�����ǰ���¼����+NO+�����Ϊ��ˮ��
If Nvl(str����, 0) = 9 Then
   str������ϸ��ˮ�� = rs��ϸ��¼!��¼���� & rs��ϸ��¼!NO & rs��ϸ��¼!���
End If
    mstrOutput = Space(600)
    mlngReturn = FormularyEntry(g���˻�����Ϣ.����05, _
                                                 str�����, _
                                                         1, _
                                           mrsTMP!��������, _
                                     rs��ϸ��¼!�շ�ϸĿID, _
                                         str������ϸ��ˮ��, _
                                       Format(rs��ϸ��¼!����ʱ��, "yyyyMMdd"), _
                                           mrsTMP!ҽ������, _
                                           mrsTMP!��Ŀ����, _
                 rs��ϸ��¼!ʵ�ս�� / (rs��ϸ��¼!����), _
                                 rs��ϸ��¼!����, _
                                           rs��ϸ��¼!ʵ�ս��, _
                                                        "", _
                                                        "", _
                                                        "", _
                                                        "", _
                                                        "", _
                                                        "", _
                                           mrsTMP!�շ����, _
                                                        "", _
                                             UserInfo.����, _
                                                        "", _
                             IIf(mrsTMP!�Ը����� = 1, 1, 0), _
                                                mstrOutput)
    
    Call WriteBusinessLOG("FormularyEntry", g���˻�����Ϣ.����05 & "," & _
                                                       str����� & "," & _
                                                            "1," & _
                                           mrsTMP!�������� & "," & _
                                     rs��ϸ��¼!�շ�ϸĿID & "," & _
                                               str������ϸ��ˮ�� & "," & _
                                       Format(rs��ϸ��¼!����ʱ��, "yyyyMMdd") & "," & _
                                           mrsTMP!ҽ������ & "," & _
                                           mrsTMP!��Ŀ���� & "," & _
               rs��ϸ��¼!ʵ�ս�� / (rs��ϸ��¼!����) & "," & _
                                 rs��ϸ��¼!���� & "," & _
                                           rs��ϸ��¼!ʵ�ս�� & "," & _
                                                             "," & _
                                                             "," & _
                                                             "," & _
                                                             "," & _
                                                             "," & _
                                                             "," & _
                                           mrsTMP!�շ���� & "," & _
                                                             "," & _
                                             UserInfo.���� & "," & _
                                                             "," & _
                             IIf(mrsTMP!�Ը����� = 1, 1, 0) & "," _
                                                , Trim(mstrOutput))
    If mlngReturn = -1 Then
        Call ����_ɽ��
        MsgBox "�ϴ�" & mrsTMP!ҽ������ & " " & mrsTMP!��Ŀ���� & "ʧ�ܣ�������ֹ��" & _
                vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
        Exit Function
    End If
    '>beging ������ֵ
                '�ڷ��ü�¼�м�¼����ͳ����
                '��Ŀ�����б�����Ŀ���ͣ�ҩƷ�����ƣ�,ժҪ�б����Ը�����,�ɸ��ݱ����õ����࣬����
        If Nvl(str����, 0) = 9 Then
            mstrOutput = Replace(mstrOutput, "|", ";")
            
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rs��ϸ��¼!ID & "," & _
                    Split(mstrOutput, ";")(1) - Split(mstrOutput, ";")(3) - Split(mstrOutput, ";")(4) & _
                    ",NULL,1,'" & str������Ŀ���� & "',NULL,NULL)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ���ֶ�")
            
        End If

    '>end  ������ֵ
    
    rs��ϸ��¼.MoveNext
Loop

'��Ԥ���㽻��
'CardNo  IN  ����    N��InHosNo IN  סԺ�ţ�����ţ�    N��MedType IN  ҽ�����ͬ����Һţ�  N
'DataBuffer  OUT ������(����ִ�гɹ�)�����ԭ��(����ִ��ʧ��)  ���鳤��600����

mstrOutput = Space(600)
'
If Nvl(str����, 0) = 9 Then
    ''����
    �����������_ɽ�� = True
    str���㷽ʽ = str����� & "|" & str�������
    Exit Function
Else
    ''Ԥ����
    mlngReturn = PreExpenseCalc(g���˻�����Ϣ.����05, str�����, str�������, mstrOutput)
    Call WriteBusinessLOG("PreExpenseCalc", g���˻�����Ϣ.����05 & "," & str����� & "," & str�������, Trim(mstrOutput))
End If

If mlngReturn = 0 Then
    '��ȷ����
    
    dbl�����ʻ� = Val(Split(mstrOutput, "|")(2))
    dblͳ����� = Val(Split(mstrOutput, "|")(3))
    dbl����Ա���� = Val(Split(mstrOutput, "|")(5))
 '�����ʻ�����ǰ̨�޸ģ����Դ�Ϊ1λ���ڼ�ͷ��ָ V
    str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";1|ͳ�����;" & dblͳ����� & ";0|����Ա����;" & dbl����Ա���� & ";0"
    �����������_ɽ�� = True
Else
    �����������_ɽ�� = False
End If
    
If Nvl(str����, 0) <> 9 Then
    Call ����_ɽ��
End If
    
    Exit Function
errHandle:
    �����������_ɽ�� = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call ����_ɽ��
End Function

Public Function �������_ɽ��(lng����ID As Long, cur����֧�� As Currency, strҽ���� As String) As Boolean
'
    '���ݽ���ID��ȡ��¼��,���ݸ��������ӿڣ������и��ݴ����"str����"����ʶ�ж��Ƿ��ύ��ϸ��
    Dim rs������ϸ As New ADODB.Recordset  ''���ݽ�����ϸ��¼
    Dim strԤ����Ϣ As String   '����Ԥ���㷵�ص���Ϣ
    Dim lng����ID As Long
    Dim cur��� As Currency
On Error GoTo ErrH
    mstrSQL = "Select ID,NO,���,��¼����,�Ǽ�ʱ�� as ����ʱ��,����ID,�շ����,�վݷ�Ŀ,���㵥λ,������, " & _
                     "�շ�ϸĿID,nvl(����,0)*nvl(����,0) as ����,��׼���� as ����, " & _
                     "ʵ�ս��,ͳ����,���մ���ID ����֧������ID, " & _
                     " ժҪ,�Ƿ��� " & _
            "from ������ü�¼ " & _
            "where ����ID=[1]"
    
    Set rs������ϸ = zlDatabase.OpenSQLRecord(mstrSQL, "���������ϸ", lng����ID)
    lng����ID = rs������ϸ!����ID
    
    
    If �������(g���˻�����Ϣ.����, vbNullString) = False Then
        Exit Function
    End If
    
    If �����������_ɽ��(rs������ϸ, strԤ����Ϣ, 9) Then
        '����ExpenseCalc
        mstrOutput = Space(600)
        rs������ϸ.MoveFirst
        
        
        ' �����                     ,ҽ�����
        'ExpenseCalc( char* CardNo,  //����
        '                      int   TransType,      //��������
        '                      int   InvoiceKind,    //��Ʊ���� 0: ���� 1:סԺ
        '                      char* InHosNo,        //סԺ�����
        '                      char* MedType,        //ҽ�����
        '                      char* InvoiceNo,      //���ݺ�
        '                      char* UserName,       //������
        '                      double AccCashPay,    //�ֽ�������
        '                      char* DataBuffer );   //������
        mstrOutput = Space(600)
        mlngReturn = ExpenseCalc(g���˻�����Ϣ.����05, 1, 0, Split(strԤ����Ϣ, "|")(0), Split(strԤ����Ϣ, "|")(1), _
                                       "1" & rs������ϸ!NO, UserInfo.����, cur����֧��, mstrOutput)
                                       
        Call WriteBusinessLOG("ExpenseCalc", g���˻�����Ϣ.����05 & "," & 1 & "," & 0 & "," & Split(strԤ����Ϣ, "|")(0) & "," & Split(strԤ����Ϣ, "|")(1) & "," & _
                                       "1" & rs������ϸ!NO & "," & UserInfo.���� & "," & cur����֧��, Trim(mstrOutput))
        '�ɹ����׺��ύ
        If mlngReturn = -1 Then
            Call ����_ɽ��
            �������_ɽ�� = False
            Err.Raise 9000, gstrSysName, Trim(mstrOutput)
        Else
        g������Ϣ_ɽ��.�����ʻ� = cur����֧��
        g������Ϣ_ɽ��.ͳ����� = Val(Split(mstrOutput, "|")(3))
        g������Ϣ_ɽ��.����Ա���� = Val(Split(mstrOutput, "|")(5))
        g������Ϣ_ɽ��.�����ܶ� = Val(Split(mstrOutput, "|")(1))
        g������Ϣ_ɽ��.�Էѽ�� = Val(Split(mstrOutput, "|")(7))
        g������Ϣ_ɽ��.������ = Val(Split(mstrOutput, "|")(6))
            
        '���ս����¼
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_ɽ�� & "," & _
                lng����ID & "," & g�ʻ�������Ϣ.�ʻ����01 & ",0,0, " & _
                "" & _
                g�ʻ�������Ϣ.����ͳ��֧���ۼ�10 + g�ʻ�������Ϣ.���깫��Ա����֧���ۼ�13 & "," & g�ʻ�������Ϣ.����סԺ����02 & ",NULL,NULL,NULL,0," & _
                g������Ϣ_ɽ��.�����ܶ� & "," & g������Ϣ_ɽ��.�Էѽ�� & "," & g������Ϣ_ɽ��.������ & ",NULL," & g������Ϣ_ɽ��.ͳ����� + g������Ϣ_ɽ��.����Ա���� & ",NULL,NULL," & _
                cur����֧�� & ",'" & Split(strԤ����Ϣ, "|")(0) & "',NULL,NULL,'" & Split(strԤ����Ϣ, "|")(1) & "')"
                '                  �����Ϊstr������ˮ��                           ҽ�����.
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ɽ��ҽ��")
        
        cur��� = �������_ɽ��(g���˻�����Ϣ.���˱��00) - cur����֧��
        gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_ɽ�� & ",'�ʻ����','" & cur��� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ɽ��ҽ��")
            
        Call �ύ_ɽ��
            �������_ɽ�� = True
        End If
    End If
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_ɽ��(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean

    Dim lng����ID As Long
    Dim str���� As String
    Dim str���� As String
    Dim StrInput As String
    Dim rsTemp As New ADODB.Recordset
    Dim str����� As String
    Dim strҽ�����  As String
    Dim strNO As String
    
    On Error GoTo errHand
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    'ֻ�������һ��
    'ȡ������¼�Ľ���ID�����ݺ�
    
    'ȡ��֤����
    gstrSQL = "Select ����,���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����������", TYPE_ɽ��, lng����ID)
    str���� = Nvl(rsTemp!����)
    str���� = Nvl(rsTemp!����)
    gstrSQL = "select distinct A.����ID,A.NO from ������ü�¼ A,������ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng����ID = rsTemp!����ID
    strNO = rsTemp!NO
    'ȡ������ˮ��
    gstrSQL = "Select * From ���ս����¼ Where ����=1 And ��¼ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������ˮ��", lng����ID, TYPE_ɽ��)
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "û���ҵ�ԭʼ�����¼���޷�����������������"
        Exit Function
    End If
    str����� = Nvl(rsTemp!֧��˳���)
    strҽ����� = Nvl(rsTemp!��ע)
    
    If str����� = "" Or strҽ����� = "" Then
        Err.Raise 9000, gstrSysName, "��ԭʼ�����¼���׺Ų�ȫ���޷�����������������"
        ����������_ɽ�� = False
        Exit Function
    End If
    
    ''��ʵ����  �������������Ҫ��������¼�봰��,�����ݲ�����
    
    If �������(str����) = False Then
        ����������_ɽ�� = False
        Exit Function
    End If
    
    '������������ҺŽ��ף����ڲ��������Ҫ���������
    
    '���ý������
    mstrOutput = Space(600)
    mlngReturn = ExpenseCalc(str����, -1, 0, str�����, strҽ�����, _
                                   strNO, UserInfo.����, cur�����ʻ�, mstrOutput)
    Call WriteBusinessLOG("ExpenseCalc", str���� & ", -1, 0," & str����� & "," & strҽ����� & "," & _
                                   strNO & "," & UserInfo.���� & "," & cur�����ʻ�, Trim(mstrOutput))
    '�ɹ���,���汾�ν������
    
    If mlngReturn = 0 Then
        Call �ύ_ɽ��
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_ɽ�� & "," & lng����ID & "," & _
            Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
            0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
            -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & -1 * Nvl(rsTemp!����ͳ����, 0) & "," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & ",0,0," & _
            -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & Nvl(rsTemp!֧��˳���) & "',null,null,'" & Nvl(rsTemp!��ע) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����������")
        
        ����������_ɽ�� = True
    Else
        Call ����_ɽ��
        ����������_ɽ�� = False
        MsgBox "�˷ѽ���ʧ�ܣ�" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
    End If
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function


Public Function ȡ������_ɽ��()

Dim str����� As String, str������� As String
Dim lng����ID As Long



mstrSQL = "select * from �����ʻ� where ����=[1] and ����=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "�����ʻ�", g���˻�����Ϣ.����05, TYPE_ɽ��)

str����� = mrsTMP!����ID & "_0_" & IIf(IsNumeric(Nvl(mrsTMP!˳���, 0)), mrsTMP!˳���, 0)
str������� = mrsTMP!�������
lng����ID = mrsTMP!����ID

If �������(mrsTMP!����, g���˻�����Ϣ.����05) = False Then
    Exit Function
End If

mstrSQL = "select * from ���ղ��� where ID=[1] and ����=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "���ղ���", lng����ID, TYPE_ɽ��)

    '��������˺ŵǼǽ���
            mstrOutput = Space(600)
            mlngReturn = Registration(g���˻�����Ϣ.����05, _
                                                                -1, _
                                                      str�����, _
                                                               "", _
                                                      str�������, _
                                                       mrsTMP!����, _
                                                       mrsTMP!����, _
                                                                  "", "", _
                                                                 "", _
                                                     UserInfo.����, _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), _
                                                        mstrOutput)
                                                        
            Call WriteBusinessLOG("Registration", g���˻�����Ϣ.����05 & "," & _
                                                                     "-1," & _
                                                            str����� & "," & _
                                                                  "," & _
                                                      str������� & "," & _
                                                       mrsTMP!���� & "," & _
                                                       mrsTMP!���� & "," & _
                                                                    "0," & _
                                                                   "," & _
                                                     UserInfo.���� & "," & _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
    If mlngReturn = 0 Then
        �ύ_ɽ��
    Else
        ����_ɽ��
        Exit Function
    End If

End Function


Public Function ��Ժ�Ǽ�_ɽ��(lngPatiID As Long, lngPageID As Long, strҽ���� As String) As Boolean
'    ����      ����/���   ������  �Ƿ�ɿ�    ����
'CardNo           IN       ����        16
'RegType          IN       �Ǽ�����
'                                -1  �޷���Ժ
'                                 1   ��Ժ�Ǽ�
'                                 2   �Ǽ���Ϣ�޸�
'                                 3   ��Ժ�Ǽ�
'MedType          IN       ҽ�����(����¼)        3
'InHosNo          IN       סԺ��      15
'ApprNo           IN       �������        15

'Returns:
'   0 - SUCCESS
'   -1 - FAILURE
'Remarks:
'    ��Ժ�Ǽ�ǰ�������Ƚ�����ʵ������Ȼ����ô����β��⺯����
' ����û���δ��ȫ�����������ɽ�����ԺסԺ����
Dim str���� As String
Dim str���� As String
Dim str������� As String
Dim str��Ժ���� As String
Dim str���ֱ��� As String
Dim str�������� As String

'>Beging ��ȡ������Ժ��Ϣ
mstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
          "C.����,D.���� As ���ұ���,C.������� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
          "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
          "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]" & _
          " and C.����=[3]"

Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "���ղ���", lngPatiID, lngPageID, TYPE_ɽ��)

If mrsTMP.EOF Then
    ��Ժ�Ǽ�_ɽ�� = False
    MsgBox "�ò���δͨ�������֤�����ܰ���ҽ����Ժ��", vbInformation, gstrSysName
    Exit Function
End If

str���� = mrsTMP!����
str������� = mrsTMP!�������
str���� = mrsTMP!����
str��Ժ���� = Format(mrsTMP!��Ժ����, "yyyyMMdd")
'>End

'>beging ��ȡ����
mstrSQL = "select * from ���ղ��� where ID=(" & _
                "Select ����ID from �����ʻ� where ����ID=[1]" & _
                                             " and ����=[2]) " & _
                                   " and ����=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "������Ϣ", lngPatiID, TYPE_ɽ��)

If mrsTMP.EOF Then
    MsgBox "����Ŀ¼����ȷ��������ִ�н��ס�", vbInformation, gstrSysName
    Exit Function
End If
str���ֱ��� = mrsTMP!����
str�������� = mrsTMP!����
'>End ��ȡ����

'>Beging ������ʵ����
If �������(str����) Then
    
    '>>beging ����Ժ�Ǽ�
    mstrOutput = Space(600)
'    TreatInfoEntry ( char* CardNo,  //����
'                             int   RegType,         //�Ǽ�����(-1.�޷���Ժ 1.סԺ�Ǽ� 2.��Ϣ�޸� 3. ��Ժ�Ǽ�)
'                             char* MedType,         //ҽ�����
'                             char* InHosNo,         //סԺ�����
'                             char* ApprNo,          //�������
'                             char* TreatDate,       //��Ժ����(yyyymmdd)
'                             char* LeaveHosDt,      //��Ժ����(yyyymmdd)
'                             char* DiseaseNo,       //��Ժ��������
'                             char* DiseaseName,     //��Ժ��������
'                             char* LHDiseaseNo,     //��Ժ��������
'                             char* LHDiseaseName,   //��Ժ��������
'                             char* LHStatus,        //��Ժ״̬(1: ���� 2: ��ת 3: ���� 9:����)
'                             char* TrunHosKind,     //תԺ��־
'                             char* MainDocName,     //����ҽʦ
'                             char* ApprPerson,      //������
'                             char* Transactor,      //������
'                             char* TransDate ,      //��������
'                             char* DataBuffer);     //������Ϣ

    mlngReturn = TreatInfoEntry(str����, 1, str�������, lngPatiID & "_1_" & lngPageID, "", _
                         str��Ժ����, "", str���ֱ���, str��������, "" _
                         , "", "", 0, "", "", UserInfo.����, _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), mstrOutput)
                         
      'תԺ��־������¼��      3 ������Ŀǰ��Ϊ0
     Call WriteBusinessLOG("TreatInfoEntry", str���� & ",1," & str������� & ", " & lngPatiID & "_" & lngPageID & ",," & _
                         str��Ժ���� & ", ," & str���ֱ��� & "," & str�������� & ", " & _
                         ", , , 0, ,," & UserInfo.���� & "," & _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
    '>>End ����Ժ�Ǽ�
    If mlngReturn = -1 Then
        Call ����_ɽ��
        ��Ժ�Ǽ�_ɽ�� = False
        MsgBox "ҽ����Ժ�Ǽ�ʧ�ܣ�" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
        Exit Function
    Else
        Call �ύ_ɽ��
        ��Ժ�Ǽ�_ɽ�� = True
        gstrSQL = "zl_�����ʻ�_��Ժ(" & lngPatiID & "," & TYPE_ɽ�� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ҽ����Ժ")
    
    End If
Else
    ��Ժ�Ǽ�_ɽ�� = False
End If
'>End ������ʵ����


End Function

Public Function ������Ժ�Ǽ�_ɽ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim blnTrans As Boolean
    Dim str���� As String, str���� As String, str������� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '�Ƿ����δ����ã�����δ����ò���������Ժ
    If ����δ�����(lng����ID, lng��ҳID) Then
        MsgBox "�ò����ѷ������ã�����������Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�������ò���������Ժ
    gstrSQL = "Select 1 From סԺ���ü�¼ Where ����ID=[1] And ��ҳID=[2] and Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������ò���������Ժ", lng����ID, lng��ҳID)
    If rsTemp.RecordCount <> 0 Then
        MsgBox "�ò����ѷ������ã�����������Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȡ���ղ��������Ϣ
    gstrSQL = "Select * From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ղ��������Ϣ", TYPE_ɽ��, lng����ID)
    str���� = rsTemp!����
    str���� = Nvl(rsTemp!����)
    str������� = Nvl(rsTemp!�������, "21")
    
    '��ʵ������У�鲡�����
    If Not �������(str����) Then Exit Function
    '    TreatInfoEntry ( char* CardNo,  //����
    '                             int   RegType,         //�Ǽ�����(-1.�޷���Ժ 1.סԺ�Ǽ� 2.��Ϣ�޸� 3. ��Ժ�Ǽ�)
    '                             char* MedType,         //ҽ�����
    '                             char* InHosNo,         //סԺ�����
    '                             char* ApprNo,          //�������
    '                             char* TreatDate,       //��Ժ����(yyyymmdd)
    '                             char* LeaveHosDt,      //��Ժ����(yyyymmdd)
    '                             char* DiseaseNo,       //��Ժ��������
    '                             char* DiseaseName,     //��Ժ��������
    '                             char* LHDiseaseNo,     //��Ժ��������
    '                             char* LHDiseaseName,   //��Ժ��������
    '                             char* LHStatus,        //��Ժ״̬(1: ���� 2: ��ת 3: ���� 9:����)
    '                             char* TrunHosKind,     //תԺ��־
    '                             char* MainDocName,     //����ҽʦ
    '                             char* ApprPerson,      //������
    '                             char* Transactor,      //������
    '                             char* TransDate ,      //��������
    '                             char* DataBuffer);     //������Ϣ
    blnTrans = True
    mstrOutput = Space(600)
    mlngReturn = TreatInfoEntry(str����, -1, str�������, lng����ID & "_1_" & lng��ҳID, "", _
                         "", "", "", "", "" _
                         , "", "", 0, "", "", UserInfo.����, _
                        Format(zlDatabase.Currentdate, "yyyyMMddHHmmss"), mstrOutput)
                         
      'תԺ��־������¼��      3 ������Ŀǰ��Ϊ0
     Call WriteBusinessLOG("TreatInfoEntry", str���� & ",1," & str������� & ", " & lng����ID & "_" & lng��ҳID & ",," & _
                         "" & ", ," & "" & "," & "" & ", " & _
                         ", , , 0, ,," & UserInfo.���� & "," & _
                        Format(zlDatabase.Currentdate, "yyyyMMddHHmmss"), Trim(mstrOutput))
    If mlngReturn = -1 Then
        MsgBox mstrOutput, vbInformation, gstrSysName
        Call ����_ɽ��
        Exit Function
    End If
    
    Call �ύ_ɽ��
    blnTrans = False
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_ɽ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    ������Ժ�Ǽ�_ɽ�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call ����_ɽ��
End Function


Public Function ��Ժ�Ǽ�_ɽ��(lngPatiID As Long, lngPageID As Long) As Boolean
'    ����      ����/���   ������  �Ƿ�ɿ�    ����
'CardNo           IN       ����        16
'RegType          IN       �Ǽ�����
'                                -1  �޷���Ժ
'                                 3   ��Ժ�Ǽ�
'MedType          IN       ҽ�����(����¼)        3
'InHosNo          IN       סԺ��      15
'ApprNo           IN       �������        15

'Returns:
'   0 - SUCCESS
'   -1 - FAILURE
'Remarks:
'    ��Ժ�Ǽ�ǰ�������Ƚ�����ʵ������Ȼ����ô����β��⺯����
    ' ����û���δ��ȫ�����������ɽ�����ԺסԺ����
    Dim bln���� As Boolean
    Dim blnTrans As Boolean
    Dim str���� As String
    Dim str���� As String
    Dim str������� As String
    Dim str��Ժ���� As String
    Dim str��Ժ���� As String
    Dim str���ֱ��� As String
    Dim str�������� As String
    Dim str��Ժ���ֱ��� As String
    Dim str��Ժ�������� As String
    Dim lng����ѡ����� As Long '1 ��Ժѡ���ˣ�2���Ժѡ���ˣ�3���Ժ��ûѡ��,
    Dim lngRegType As Long  '�Ǽ�����  (ҽ���Ĳ���)
    Dim str����ѡ�񷵻�ֵ As String
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
    
    '> Beging �Ƿ����޷���Ժ
    lngRegType = 3
    If Not ����δ�����(lngPatiID, lngPageID) Then
        '�Ƿ��ѽ��ʣ��ѽ��ʲ��˲���������Ժ
        bln���� = False
        gstrSQL = "Select 1 From סԺ���ü�¼ Where ����ID=[1] And ��ҳID=[2] And Nvl(����ID,0)<>0 and Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�õ��þ���Ǽǳ���", lngPatiID, lngPageID)
        If Not rsTemp.EOF Then
            bln���� = True
        End If
        
        If Not bln���� Then
            '�޷���Ժ=������Ժ�Ǽ�
            lngRegType = -1
        End If
    End If
    '> End �Ƿ����޷���Ժ
    
    '>Beging ��ȡ������Ժ��Ϣ
    mstrSQL = "select A.��Ժ����,A.��Ժ����,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
              "C.����,D.���� As ���ұ���,C.������� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
              "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
              "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]" & _
              " and C.����=[3]"
    Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "���ղ���", lngPatiID, lngPageID, TYPE_ɽ��)
    If mrsTMP.EOF Then
        MsgBox "�ò���δͨ�������֤�����ܰ���ҽ����Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    str���� = mrsTMP!����
    str������� = mrsTMP!�������
    str���� = mrsTMP!����
    str��Ժ���� = Format(mrsTMP!��Ժ����, "yyyyMMdd")
    str��Ժ���� = Format(mrsTMP!��Ժ����, "yyyyMMdd")
    '>End ��ȡ������Ժ��Ϣ
    '>
    
    '>beging ��ȡ����
    mstrSQL = "select * from ���ղ��� where ID=(" & _
                    "Select ����ID from �����ʻ� where ����ID=[1]" & _
                                                 " and ����=[2]) " & _
                                       " and ����=[2]"
    Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "������Ϣ", lngPatiID, TYPE_ɽ��)
    
    If mrsTMP.EOF Then
        lng����ѡ����� = 0
    Else
        lng����ѡ����� = 1
        str���ֱ��� = mrsTMP!����
        str�������� = mrsTMP!����
    End If
    
    mstrSQL = "select * from ���ղ��� where ID=(" & _
                    "Select ��Ժ����ID from �����ʻ� where ����ID=[1]" & _
                                                 " and ����=[2]) " & _
                                       " and ����=[2]"
    Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "������Ϣ", lngPatiID, TYPE_ɽ��)
    If mrsTMP.EOF Then
        lng����ѡ����� = lng����ѡ����� - 2
    Else
        lng����ѡ����� = lng����ѡ����� + 2
        str��Ժ���ֱ��� = mrsTMP!����
        str��Ժ�������� = mrsTMP!����
    End If
    '>>Beging �ж��Ƿ��в���ûѡ�����û�У���ǿ��ѡ��,Ȼ����ݷ���ֵ����
    
    If lng����ѡ����� <> 3 Then
         If Not frm����ѡ��_ɽ��.Select����(lngPatiID, str���ֱ���, str��������, str��Ժ���ֱ���, str��Ժ��������) Then
             ��Ժ�Ǽ�_ɽ�� = False
             MsgBox "��ѡ���ֺ��ٰ����Ժ�Ǽǣ�", vbInformation, gstrSysName
             Exit Function
         End If
    End If
    '>>End �ж��Ƿ��в���ûѡ�����û�У���ǿ��ѡ��,Ȼ����ݷ���ֵ����
    
    '>End ��ȡ����
    
    '>Beging �����������
    If �������(str����) Then
        blnTrans = True
        '>>beging ����Ժ�Ǽ�
        mstrOutput = Space(600)
        mlngReturn = TreatInfoEntry(str����, lngRegType, str�������, lngPatiID & "_1_" & lngPageID, "", _
                             str��Ժ����, str��Ժ����, str���ֱ���, str��������, str��Ժ���ֱ��� _
                             , str��Ժ��������, "", 0, "", "", UserInfo.����, _
                            Format(zlDatabase.Currentdate, "yyyyMMdd"), mstrOutput)
                             
          'תԺ��־������¼��      3 ������Ŀǰ��Ϊ0
         Call WriteBusinessLOG("TreatInfoEntry", str���� & "," & lngRegType & "," & str������� & ", " & lngPatiID & "_1_" & lngPageID & ",," & _
                             str��Ժ���� & "," & str��Ժ���� & "," & str���ֱ��� & "," & str�������� & ", " & str��Ժ���ֱ��� & _
                             "," & str��Ժ�������� & ", , 0, , ," & UserInfo.���� & "," & _
                            Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
        '>>End ����Ժ�Ǽ�
        If mlngReturn = -1 Then
            Call ����_ɽ��
            ��Ժ�Ǽ�_ɽ�� = False
            MsgBox "ҽ����Ժ�Ǽ�ʧ�ܣ�" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
            Exit Function
        Else
            Call �ύ_ɽ��
            blnTrans = False
            ��Ժ�Ǽ�_ɽ�� = True
            gstrSQL = "zl_�����ʻ�_��Ժ(" & lngPatiID & "," & TYPE_ɽ�� & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
        End If
    Else
        ��Ժ�Ǽ�_ɽ�� = False
    End If
    '>End �����������

    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call ����_ɽ��
End Function

Public Sub ���²���_ɽ��(lngPatiID As Long, lngPageID As Long)
Dim str��Ժ���ֱ��� As String, str��Ժ�������� As String
Dim str��Ժ���ֱ��� As String, str��Ժ�������� As String
   
   If Not frm����ѡ��_ɽ��.Select����(lngPatiID, str��Ժ���ֱ���, str��Ժ��������, _
    str��Ժ���ֱ���, str��Ժ��������) Then Exit Sub
    
End Sub



Public Function סԺ�������_ɽ��(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim strPin As String                'IC������
    Dim str������� As String           '�������
    Dim strҽ���� As String             'ҽ��֤��
    Dim strסԺ�� As String             'סԺ��
    Dim lng��ҳID As Long               '��ҳID
    Dim blnTrans As Boolean             '�Ƿ�ʼҽ������
    Dim blnOut As Boolean               '�����Ƿ��ѳ�Ժ����������;���㻹�ǳ�Ժ����
    Dim dbl�����ʻ� As Double, dblҽ������ As Double, dbl����Ա���� As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset
    Dim cur�������� As Currency
    
    '��ϸ�ϴ���ر���
    Dim str��ˮ�� As String, str��� As String, str������� As String
    Dim int��Ŀ��� As Integer          '������Ŀ���
    Dim dbl�Ը����� As Double
    Dim strҽԺ���� As String, strҽԺ���� As String
    Dim strҽ������ As String, strҽ������ As String
    Dim strƵ�� As String, str�÷� As String, str���� As String, str���� As String
    Dim str���� As String
    
    Const int�����ܶ� As Integer = 1
    Const int�����ʻ� As Integer = 2
    Const intҽ������ As Integer = 3
    Const int����Ա���� As Integer = 5
    On Error GoTo errHand

    '��ȡ����IC������
    gstrSQL = "Select ҽ����,����,����,������� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����IC������", TYPE_ɽ��, lng����ID)
    strPin = Nvl(rsTemp!����)
    str������� = Nvl(rsTemp!�������, "11")
    strҽ���� = rsTemp!ҽ����
    str���� = rsTemp!����
    
    '��ȡ������ҳID����Ժ����
    gstrSQL = " Select A.��ҳID,A.��Ժ���� From ������ҳ A,������Ϣ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ҳID����Ժ����", lng����ID)
    blnOut = Not (IsNull(rsTemp!��Ժ����))
    lng��ҳID = rsTemp!��ҳID
    strסԺ�� = lng����ID & "_1_" & lng��ҳID
    
    '�ȶ���������ǲ���ͬһ�����˵Ŀ�
    If Not �������(strPin, str����) Then Exit Function
    
    blnTrans = True
    If strҽ���� <> g���˻�����Ϣ.���˱��00 Then
        MsgBox "��ǰIC�����Ǹò��˵ģ�", vbInformation, gstrSysName
        Call ����_ɽ��
        Exit Function
    End If
    
    'ȡ���ܷ���
    Do Until rsExse.EOF
        cur�������� = cur�������� + rsExse("���")
        rsExse.MoveNext
    Loop
    cur�������� = Val(Format(cur��������, "#####0.00"))

    '�ϴ�δ�ϴ��ķ�����ϸ
    gstrSQL = "  Select A.ID,A.NO,A.����ID,A.�շ����,A.��¼����,A.��¼״̬,A.���,A.�շ�ϸĿID,C.��Ŀ���� AS ҽ����Ŀ����,B.����,B.����,A.ʵ�ս�� AS ���" & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸�,A.������ AS ҽ��,A.�Ǽ�ʱ��,E.���� AS �������� " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C,���ű� E" & _
              "  Where A.����ID=[1] and A.��ҳID=[2] and A.���ʷ���=1 And A.����Ա���� is not null AND A.ʵ�ս�� IS NOT NULL " & _
              "        and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����= " & TYPE_ɽ�� & _
              "        and A.��������ID=E.ID " & _
              "  Order by A.����ID,A.����ʱ��"
    Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���η�����ϸ", lng����ID, lng��ҳID)
    With rsDetail
        Do While Not .EOF
            '���˵��
            'DLLFUNC int WINAPI FormularyEntry(
            '    char* CardNo,          //����          '    char* InHosNo,         //סԺ�����            '    int TransKind,         //��������          '    int ItemKind,          //��Ŀ���
            '    char* InternalCode,    //�շ���ĿҽԺ����          '    char* FormularyNo,     //������            '    char* SysDate,         //��������(yyyymmdd)            '    char* CenterCode,      //�շ���Ŀ���ı���
            '    char* ItemName,        //�շ���Ŀ����          '    double UnitPrice,      //����          '    double Quantity,       //����          '    double Amount,         //���
            '    char* DoseType,        //����          '    char* Dosage,          //����          '    char* Frequency,       //Ƶ��          '    char* Usage,           //�÷�
            '    char* KeBie,           //�Ʊ�          '    float ExecDays,        //ִ������          '    char* FeeType,         //ҽ�������շ����          '    char* DoctName,        //����ҽ��
            '    char* Transactor,      //������            '    char* ApprPerson,      //������            '    int  IsOwnExpenses,    //ȫ���Էѱ�־          '    char* DataBuffer   );

            '��ȡ�շ���Ŀ��Ϣ
            gstrSQL = " Select A.���,A.ID AS �շ�ϸĿID,A.���� As ҽԺ����,A.���� AS ҽԺ����,B.��Ŀ���� As ҽ������,B.��Ŀ���� AS ҽ������,B.��ע" & _
                    " From �շ�ϸĿ A,(Select * From ����֧����Ŀ Where ����=[1]) B" & _
                    " Where A.ID=[2] And A.ID=B.�շ�ϸĿID(+) "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ���Ŀ��Ϣ", TYPE_ɽ��, CLng(!�շ�ϸĿID))
            If IsNull(rsTemp!ҽ������) Then
                MsgBox "��Ŀ[" & rsTemp!ҽԺ���� & "]" & rsTemp!ҽԺ���� & "δ���룡��������Ŀ����", vbInformation, gstrSysName
                Call �ύ_ɽ��
                Exit Function
            End If
            str��� = rsTemp!���
            int��Ŀ��� = Nvl(rsTemp!��ע, 0)
            strҽԺ���� = rsTemp!ҽԺ����: strҽԺ���� = rsTemp!ҽԺ����
            strҽ������ = Nvl(rsTemp!ҽ������): strҽ������ = Nvl(rsTemp!ҽ������)

            '��ȡ��Ŀ���������
            If int��Ŀ��� = 1 Then
                gstrSQL = "Select aka063 ���,aka069 �Ը����� From ka02 where aka060='" & strҽ������ & "'"
            ElseIf int��Ŀ��� = 2 Then
                gstrSQL = "Select aka063 ���,aka069 �Ը����� From ka03 where aka090='" & strҽ������ & "'"
            Else
                gstrSQL = "Select aka063 ���,0 �Ը����� From ka04 where aka100='" & strҽ������ & "'"
            End If
            Call OpenRecordset_OtherBase(rsTemp, "��ȡ������Ŀ���������", gstrSQL, gcnSxDr)
            If rsTemp.RecordCount = 0 Then
                MsgBox "������Ϊ��Ŀ[" & strҽԺ���� & "]" & strҽԺ���� & "���ж��룬������Ŀ��ɾ����ȡ����", vbInformation, gstrSysName
                Call �ύ_ɽ��
                Exit Function
            End If
            dbl�Ը����� = Nvl(rsTemp!�Ը�����, 0)
            str������� = rsTemp!���
            
            '�����ҩƷ������ȡ��ҩƷ�����Ϣ��Ƶ�Ρ��÷������͡�������
            strƵ�� = "": str�÷� = "": str���� = "": str���� = ""
            If InStr(1, ",5,6,7,", rsTemp!���) <> 0 Then
                gstrSQL = " Select A.Ƶ��,A.�÷�,D.���� AS ����,C.������λ " & _
                        " From (" & _
                        "   Select * from ҩƷ�շ���¼ " & _
                        "   Where ���� in (9,10) and NO=[2]) A,ҩƷĿ¼ B,ҩƷ��Ϣ C,ҩƷ���� D" & _
                        "   Where A.����ID=[1] And A.ҩƷID=B.ҩƷID And B.ҩ��ID=C.ҩ��ID And C.����=D.����"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ�����Ϣ", CLng(rsTemp!ID), CStr(rsTemp!NO))
                strƵ�� = Nvl(rsTemp!Ƶ��)
                str�÷� = Nvl(rsTemp!�÷�)
                str���� = Nvl(rsTemp!����)
                str���� = Nvl(rsTemp!������λ)
            End If

            mstrOutput = Space(600)
            str��ˮ�� = !��¼���� & !NO & 1 & !���
            mlngReturn = FormularyEntry(g���˻�����Ϣ.����05, strסԺ��, IIf(!��¼״̬ = 2, -1, 1), int��Ŀ���, _
                    strҽԺ����, str��ˮ��, Format(!�Ǽ�ʱ��, "yyyyMMdd"), strҽ������, _
                    strҽ������, !�۸�, Abs(!����), Abs(!���), _
                    str����, str����, strƵ��, str�÷�, _
                    !��������, 0, str�������, Nvl(!ҽ��), _
                    UserInfo.����, "", IIf(dbl�Ը����� = 1, 1, 0), mstrOutput)
            Call WriteBusinessLOG("FormularyEntry", g���˻�����Ϣ.����05 & "," & strסԺ�� & "," & IIf(!��¼״̬ = 2, -1, 1) & "," & int��Ŀ��� & "," & _
                    strҽԺ���� & "," & str��ˮ�� & "," & Format(!�Ǽ�ʱ��, "yyyyMMdd") & "," & strҽ������ & "," & _
                    strҽ������ & "," & !�۸� & "," & Abs(!����) & "," & Abs(!���) & "," & _
                    str���� & "," & str���� & "," & strƵ�� & "," & str�÷� & "," & _
                    !�������� & "," & 0 & "," & str������� & "," & Nvl(!ҽ��) & "," & _
                    UserInfo.���� & ",," & IIf(dbl�Ը����� = 1, 1, 0), mstrOutput)
            If mlngReturn = -1 Then
                MsgBox "�ϴ�����[" & !NO & "]��" & !��� & "����ϸʱ����" & vbCrLf & "��ϸ��Ϣ��" & vbCrLf & mstrOutput, vbInformation, gstrSysName
                Call �ύ_ɽ��
                Exit Function
            End If

            '���·�����ϸ�е�ͳ�������Ϣ
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & !ID & "," & _
                    Split(mstrOutput, "|")(1) - Split(mstrOutput, "|")(3) - Split(mstrOutput, "|")(4) & _
                    ",NULL,1,'" & strҽ������ & "',1,'NULL')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ���ֶ�")
            .MoveNext
        Loop
    End With

    'סԺԤ����
    'DLLFUNC int WINAPI PreExpenseCalc ( char* CardNo,   //����,char* InHosNo,          //סԺ����� ,char* MedType,            //ҽ�����,char* DataBuffer );  //������Ϣ
    '���ز���˵��
    '|�����ܶ�|�����ʻ�֧��|ͳ��֧��|�ֽ�֧��|����Ա����֧��
    '|������|�Էѽ��|סԺ�˴�|�𸶱�׼|תԺ���Ʒ���
    '|�𸶱�׼�Ը�|�𸶱�׼����Ա֧��|�ֶ�1ͳ��֧��|�ֶ�1����Ա֧��|�ֶ�1�����Ը�
    '|�ֶ�2ͳ��֧��|�ֶ�2����Ա֧��|�ֶ�2�����Ը�|�ֶ�3ͳ��֧��|�ֶ�3����Ա֧��
    '|�ֶ�3�����Ը�|���ⶥ����Ա֧��|����Ա�����Ը�|������Ա�ⶥ�����Ը�|������ϵ|��λ����|
    mstrOutput = Space(600)
    mlngReturn = PreExpenseCalc(g���˻�����Ϣ.����05, strסԺ��, str�������, mstrOutput)
    Call WriteBusinessLOG("PreExpenseCalc", g���˻�����Ϣ.����05 & "," & strסԺ�� & "," & str�������, mstrOutput)
    
    If mlngReturn = -1 Then
        MsgBox mstrOutput, vbInformation, gstrSysName
        Call �ύ_ɽ��
        Exit Function
    Else
        If cur�������� <> Val(Format(Split(mstrOutput, "|")(1), "#####0.00")) Then
            If MsgBox("ҽԺ�ķ����ܽ��(" & cur�������� & ")��ҽ�����ĵķ����ܶ�(" & Val(Format(Split(mstrOutput, "|")(1), "#####0.00")) & ")���ȣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Call �ύ_ɽ��
                Exit Function
            End If
        End If

    End If

    'ȡ������Ϣ
    dbl�����ʻ� = Val(Split(mstrOutput, "|")(int�����ʻ�))
    dblҽ������ = Val(Split(mstrOutput, "|")(intҽ������))
    dbl����Ա���� = Val(Split(mstrOutput, "|")(int����Ա����))

    '���ؽ�����Ϣ
    blnTrans = False
    סԺ�������_ɽ�� = "�����ʻ�;" & dbl�����ʻ� & ";1"
    סԺ�������_ɽ�� = סԺ�������_ɽ�� & "|ͳ�����;" & dblҽ������ & ";1"
    סԺ�������_ɽ�� = סԺ�������_ɽ�� & "|����Ա����;" & dbl����Ա���� & ";1"
    
    Call �ύ_ɽ��
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then Call �ύ_ɽ��
End Function

Public Function סԺ����_ɽ��(ByVal lng����ID As Long) As Boolean
    Dim strNO As String, strPin As String
    Dim str������� As String, strҽ���� As String, strסԺ�� As String
    Dim dbl�����ʻ� As Double, dblҽ������ As Double, dbl����Ա���� As Double, dbl�ܷ��� As Double, dbl�ֽ� As Double
    Dim lng����ID As Long, lng��ҳID As Long
    Dim blnTrans As Boolean
    Dim blnOut As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    Const int�����ܶ� As Integer = 1
    Const int�����ʻ� As Integer = 2
    Const intҽ������ As Integer = 3
    Const int�ֽ� As Integer = 4
    Const int����Ա���� As Integer = 5
    On Error GoTo errHand
    
    '��ȡ���ʵ���
    gstrSQL = "Select NO,����ID From ���˽��ʼ�¼ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ʵ���", lng����ID)
    strNO = "2" & rsTemp!NO
    lng����ID = rsTemp!����ID
    
    '��ȡ�ʻ�ʵ��֧����
    gstrSQL = "Select Nvl(��Ԥ��,0) AS �����ʻ� From ����Ԥ����¼ Where ����ID=[1] And ��¼���� Not In (1,11) And ���㷽ʽ='�����ʻ�'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ�ʵ��֧����", lng����ID)
    If Not rsTemp.EOF Then
        dbl�����ʻ� = Nvl(rsTemp!�����ʻ�, 0)
    End If

    '��ȡ����IC������
    gstrSQL = "Select ҽ����,����,������� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����IC������", TYPE_ɽ��, lng����ID)
    strPin = Nvl(rsTemp!����)
    str������� = Nvl(rsTemp!�������, "11")
    strҽ���� = rsTemp!ҽ����

    '��ȡ������ҳID����Ժ����
    gstrSQL = " Select A.��ҳID,A.��Ժ���� From ������ҳ A,������Ϣ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ҳID����Ժ����", lng����ID)
    blnOut = Not (IsNull(rsTemp!��Ժ����))
    lng��ҳID = rsTemp!��ҳID
    strסԺ�� = lng����ID & "_1_" & lng��ҳID

    '�ȶ���������ǲ���ͬһ�����˵Ŀ�
    If Not �������(strPin) Then Exit Function
    blnTrans = True
    If strҽ���� <> g���˻�����Ϣ.���˱��00 Then
        Err.Raise 9000, gstrSysName, "��ǰIC�����Ǹò��˵ģ�"
        Call ����_ɽ��
        Exit Function
    End If

    '���˵��
    'DLLFUNC int WINAPI ExpenseCalc(
        'char* CardNo,          //����      'int   TransType,       //��������      'int   InvoiceKind,     //��Ʊ���� 0: ���� 1:סԺ
        'char* InHosNo,         //סԺ�����        'char* MedType,         //ҽ�����      'char* InvoiceNo,       //���ݺ�
        'char* UserName,        //������        'double AccCashPay,     //�ʻ�֧�����      'char* DataBuffer );    //������
    mstrOutput = Space(600)
    mlngReturn = ExpenseCalc(g���˻�����Ϣ.����05, IIf(blnOut, 1, 2), 1, _
        strסԺ��, str�������, strNO, _
        UserInfo.����, dbl�����ʻ�, mstrOutput)
    Call WriteBusinessLOG("ExpenseCalc", g���˻�����Ϣ.����05 & "," & IIf(blnOut, 1, 2) & ",1," & _
        strסԺ�� & "," & str������� & "," & strNO & "," & _
        UserInfo.���� & "," & dbl�����ʻ�, mstrOutput)
    If mlngReturn = -1 Then
        Err.Raise 9000, gstrSysName, mstrOutput
        Call ����_ɽ��
        Exit Function
    End If

    'ȡ������Ϣ
    dbl�ܷ��� = Val(Split(mstrOutput, "|")(int�����ܶ�))
    dbl�ֽ� = Val(Split(mstrOutput, "|")(int�ֽ�))
    dblҽ������ = Val(Split(mstrOutput, "|")(intҽ������))
    dbl����Ա���� = Val(Split(mstrOutput, "|")(int����Ա����))
    
    '�����ܶ�Ƚϣ�
    
    Call �ύ_ɽ��
    blnTrans = False
    
    '���汣�ս����¼
    '���Ը�=�󲡲���;�����Ը�=����Ա����
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_ɽ�� & "," & lng����ID & "," & _
        Year(zlDatabase.Currentdate()) & ",0,0,0,0,0,0,0,0," & dbl�ܷ��� & "," & dbl�ֽ� & ",0," & _
        dblҽ������ & "," & dblҽ������ & ",0," & dbl����Ա���� & "," & dbl�����ʻ� & ",'" & strסԺ�� & "'," & lng��ҳID & "," & IIf(blnOut, 1, 0) & ",'" & str������� & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    סԺ����_ɽ�� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then Call ����_ɽ��
End Function


Public Function ��Ժ�Ǽǳ���_ɽ��(lng����ID As Long, lng��ҳID As Long) As Boolean

Dim str���� As String
Dim str���� As String
Dim str������� As String
Dim str��Ժ���� As String
Dim str��Ժ���� As String
Dim str���ֱ��� As String
Dim str�������� As String
Dim str��Ժ���ֱ��� As String
Dim str��Ժ�������� As String
Dim lng����ѡ����� As Long '1 ��Ժѡ���ˣ�2���Ժѡ���ˣ�3���Ժ��ûѡ��,
Dim lngRegType As Long  '�Ǽ�����  (ҽ���Ĳ���)
Dim str����ѡ�񷵻�ֵ As String

'//////////////////////////  ��Ժ�Ǽǣ�����޷��ã������޷��˷ѣ���Ϊ��ҽ��,�����Ƿ���Ժ��
'                              ���з��ã������ѳ�Ժ������Ժ�Ǽ�
lngRegType = 2   '���־Ϊ�޸ļ�¼

'>Beging ��ȡ������Ժ��Ϣ
mstrSQL = "select A.��Ժ����,A.��Ժ����,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
          "C.����,D.���� As ���ұ���,C.������� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
          "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
          "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]" & _
          " and C.����=[3]"

Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "���ղ���", lng����ID, lng��ҳID, TYPE_ɽ��)
If mrsTMP.EOF Then
    ��Ժ�Ǽǳ���_ɽ�� = False
    MsgBox "�ò���δͨ�������֤�����ܰ�������Ժ��", vbInformation, gstrSysName
    Exit Function
End If

str���� = mrsTMP!����
str������� = mrsTMP!�������
str���� = mrsTMP!����
str��Ժ���� = Format(mrsTMP!��Ժ����, "yyyyMMdd")
str��Ժ���� = "" '������Ժ�����Գ�Ժ���ڸ�Ϊ��
'>End ��ȡ������Ժ��Ϣ



'>

'>beging ��ȡ����
mstrSQL = "select * from ���ղ��� where ID=(" & _
                "Select ����ID from �����ʻ� where ����ID=[1]" & _
                                             " and ����=[2]) " & _
                                   " and ����=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "������Ϣ", lng����ID, TYPE_ɽ��)

If mrsTMP.EOF Then
    lng����ѡ����� = 0
Else
    lng����ѡ����� = 1
    str���ֱ��� = mrsTMP!����
    str�������� = mrsTMP!����
End If

mstrSQL = "select * from ���ղ��� where ID=(" & _
                "Select ��Ժ����ID from �����ʻ� where ����ID=[1]" & _
                                             " and ����=[2]) " & _
                                   " and ����=[2]"
Set mrsTMP = zlDatabase.OpenSQLRecord(mstrSQL, "������Ϣ", lng����ID, TYPE_ɽ��)
If mrsTMP.EOF Then
    lng����ѡ����� = lng����ѡ����� - 2
Else
    lng����ѡ����� = lng����ѡ����� + 2
    str��Ժ���ֱ��� = mrsTMP!����
    str��Ժ�������� = mrsTMP!����
End If
'>>Beging �ж��Ƿ��в���ûѡ�����û�У���ǿ��ѡ��,Ȼ����ݷ���ֵ����

If lng����ѡ����� <> 3 Then
     If Not frm����ѡ��_ɽ��.Select����(lng����ID, str���ֱ���, str��������, str��Ժ���ֱ���, str��Ժ��������) Then
         ��Ժ�Ǽǳ���_ɽ�� = False
         MsgBox "��ѡ���ֺ��ٰ����Ժ�Ǽǣ�", vbInformation, gstrSysName
         Exit Function
     End If
End If
'>>End �ж��Ƿ��в���ûѡ�����û�У���ǿ��ѡ��,Ȼ����ݷ���ֵ����

'>End ��ȡ����

'>Beging ����������� ҽ��������ʵ����
If �������(str����) Then
    
    '>>beging ���ӿ�
    mstrOutput = Space(600)
    mlngReturn = TreatInfoEntry(str����, lngRegType, str�������, lng����ID & "_1_" & lng��ҳID, "", _
                         str��Ժ����, str��Ժ����, str���ֱ���, str��������, str��Ժ���ֱ��� _
                         , str��Ժ��������, "", 0, "", "", UserInfo.����, _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), mstrOutput)
                         
      'תԺ��־������¼��      3 ������Ŀǰ��Ϊ0
     Call WriteBusinessLOG("TreatInfoEntry", str���� & "," & lngRegType & "," & str������� & ", " & lng����ID & "_1_" & lng��ҳID & ",," & _
                         str��Ժ���� & "," & str��Ժ���� & "," & str���ֱ��� & "," & str�������� & ", " & str��Ժ���ֱ��� & _
                         "," & str��Ժ�������� & ", , 0, , ," & UserInfo.���� & "," & _
                        Format(zlDatabase.Currentdate, "yyyyMMdd"), Trim(mstrOutput))
    '>>End ���ӿ�
    If mlngReturn = -1 Then
        Call ����_ɽ��
        ��Ժ�Ǽǳ���_ɽ�� = False
        MsgBox "ҽ��������Ժʧ�ܣ�" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
        Exit Function
    Else
        Call �ύ_ɽ��
        ��Ժ�Ǽǳ���_ɽ�� = True
        gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_ɽ�� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    End If
Else
    ��Ժ�Ǽǳ���_ɽ�� = False
End If
'>End �����������

End Function


Public Function �����ϴ�_ɽ��(int���� As Integer, int״̬ As Integer, str���ݺ� As String) As Boolean

    Dim strPin As String                'IC������
    Dim str������� As String           '�������
    'Dim strҽ���� As String             'ҽ��֤��
    Dim strסԺ�� As String             'סԺ��
    Dim lng��ҳID As Long               '��ҳID
    Dim blnOut As Boolean               '�����Ƿ��ѳ�Ժ����������;���㻹�ǳ�Ժ����
    Dim lng����ID As Long               '����ID
    Dim dbl�����ʻ� As Double, dblҽ������ As Double, dbl����Ա���� As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset
    Dim rsCd As New ADODB.Recordset
    
    '��ϸ�ϴ���ر���
    Dim str��ˮ�� As String, str��� As String, str������� As String
    Dim int��Ŀ��� As Integer          '������Ŀ���
    Dim dbl�Ը����� As Double
    Dim strҽԺ���� As String, strҽԺ���� As String
    Dim strҽ������ As String, strҽ������ As String
    Dim strƵ�� As String, str�÷� As String, str���� As String, str���� As String
    
    Const int�����ܶ� As Integer = 1
    Const int�����ʻ� As Integer = 2
    Const intҽ������ As Integer = 3
    Const int����Ա���� As Integer = 5
    On Error GoTo errHand
    
    ' �����¼״̬Ϊ1�ĵ��ݣ��и�����¼���������浥��
    If int״̬ = 1 Then
        gstrSQL = "Select distinct  A.����ID from סԺ���ü�¼ A,�����ʻ� B " & _
            "where A.����ID=B.����ID And A.��¼����=[1]" & _
            " And A.��¼״̬=[2] And A.NO=[3] " & _
            " And B.����=[4] And A.ʵ�ս��<0"
        Set rsCd = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ��и�����¼", int����, int״̬, str���ݺ�, TYPE_ɽ��)
        If Not rsCd.EOF Then
            MsgBox "��ҽ����֧�ָ�����¼���������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '����NO��,��ȡ����ID
    gstrSQL = "Select distinct  A.����ID from סԺ���ü�¼ A,�����ʻ� B " & _
            "where A.����ID=B.����ID And A.��¼����=[1]" & _
            " And A.��¼״̬=[2] And A.NO=[3]" & _
            " And B.����=[4]"
    Set rsCd = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", int����, int״̬, str���ݺ�, TYPE_ɽ��)
    
    
    �����ϴ�_ɽ�� = True
    '> Beging �����˴���ϸ.
    Do Until rsCd.EOF
        '>> Beging �������
        lng����ID = rsCd!����ID
        gstrSQL = "Select * from �����ʻ� where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", lng����ID)
        If �������(rsTemp!����, rsTemp!����) Then
            gstrSQL = "Select A.*,nvl(A.����,1)*nvl(A.����,0) as ����,A.ʵ�ս�� as ���," & _
                              "nvl(A.ʵ�ս��,0)/(nvl(A.����,1)*nvl(A.����,0)) as �۸�,A.������ as ҽ��,C.���� as ��������,B.* " & _
                      " from סԺ���ü�¼ A,�����ʻ� B,���ű� C" & _
                      " where A.NO=[1]" & _
                            " And A.��¼����=[2]" & _
                            " And A.��¼״̬=[3]" & _
                            " And nvl(A.�Ƿ��ϴ�,0)=0 " & _
                            " And A.����ID=B.����ID " & _
                            " and B.����=[4]" & _
                            " and A.��������ID=C.ID " & _
                            " ANd A.����ID=[5]"
        
            Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���η�����ϸ", str���ݺ�, int����, int״̬, TYPE_ɽ��, lng����ID)
            With rsDetail
                Do While Not .EOF
                    '���˵��
                    'DLLFUNC int WINAPI FormularyEntry(
                    '    char* CardNo,          //����          '    char* InHosNo,         //סԺ�����            '    int TransKind,         //��������          '    int ItemKind,          //��Ŀ���
                    '    char* InternalCode,    //�շ���ĿҽԺ����          '    char* FormularyNo,     //������            '    char* SysDate,         //��������(yyyymmdd)            '    char* CenterCode,      //�շ���Ŀ���ı���
                    '    char* ItemName,        //�շ���Ŀ����          '    double UnitPrice,      //����          '    double Quantity,       //����          '    double Amount,         //���
                    '    char* DoseType,        //����          '    char* Dosage,          //����          '    char* Frequency,       //Ƶ��          '    char* Usage,           //�÷�
                    '    char* KeBie,           //�Ʊ�          '    float ExecDays,        //ִ������          '    char* FeeType,         //ҽ�������շ����          '    char* DoctName,        //����ҽ��
                    '    char* Transactor,      //������            '    char* ApprPerson,      //������            '    int  IsOwnExpenses,    //ȫ���Էѱ�־          '    char* DataBuffer   );
        
                    '��ȡ�շ���Ŀ��Ϣ
                    gstrSQL = " Select A.���,A.ID AS �շ�ϸĿID,A.���� As ҽԺ����,A.���� AS ҽԺ����,B.��Ŀ���� As ҽ������,B.��Ŀ���� AS ҽ������,B.��ע" & _
                            " From �շ�ϸĿ A,(Select * From ����֧����Ŀ Where ����=[1]) B" & _
                            " Where A.ID=[2] And A.ID=B.�շ�ϸĿID(+) "
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ���Ŀ��Ϣ", TYPE_ɽ��, CLng(!�շ�ϸĿID))
                    If IsNull(rsTemp!ҽ������) Then
                        MsgBox "��Ŀ[" & rsTemp!ҽԺ���� & "]" & rsTemp!ҽԺ���� & "δ���룡��������Ŀ����", vbInformation, gstrSysName
                        Call �ύ_ɽ��
                        Exit Function
                    End If
                    str��� = rsTemp!���
                    int��Ŀ��� = Nvl(rsTemp!��ע, 0)
                    strҽԺ���� = rsTemp!ҽԺ����: strҽԺ���� = rsTemp!ҽԺ����
                    strҽ������ = Nvl(rsTemp!ҽ������): strҽ������ = Nvl(rsTemp!ҽ������)
        
                    '��ȡ��Ŀ���������
                    If int��Ŀ��� = 1 Then
                        gstrSQL = "Select aka063 ���,aka069 �Ը����� From ka02 where aka060='" & strҽ������ & "'"
                    ElseIf int��Ŀ��� = 2 Then
                        gstrSQL = "Select aka063 ���,aka069 �Ը����� From ka03 where aka090='" & strҽ������ & "'"
                    Else
                        gstrSQL = "Select aka063 ���,0 �Ը����� From ka04 where aka100='" & strҽ������ & "'"
                    End If
                    Call OpenRecordset_OtherBase(rsTemp, "��ȡ������Ŀ���������", gstrSQL, gcnSxDr)
                    If rsTemp.RecordCount = 0 Then
                        MsgBox "������Ϊ��Ŀ[" & strҽԺ���� & "]" & strҽԺ���� & "���ж��룬������Ŀ��ɾ����ȡ����", vbInformation, gstrSysName
                        Call �ύ_ɽ��
                        Exit Function
                    End If
                    dbl�Ը����� = Nvl(rsTemp!�Ը�����, 0)
                    str������� = rsTemp!���
                    
                    '�����ҩƷ������ȡ��ҩƷ�����Ϣ��Ƶ�Ρ��÷������͡�������
                    strƵ�� = "": str�÷� = "": str���� = "": str���� = ""
                    If InStr(1, ",5,6,7,", rsTemp!���) <> 0 Then
                        gstrSQL = " Select A.Ƶ��,A.�÷�,D.���� AS ����,C.������λ " & _
                                " From (" & _
                                "   Select * from ҩƷ�շ���¼ " & _
                                "   Where ���� in (9,10) and NO='" & rsTemp!NO & "') A,ҩƷĿ¼ B,ҩƷ��Ϣ C,ҩƷ���� D" & _
                                "   Where A.����ID=" & rsTemp!ID & " And A.ҩƷID=B.ҩƷID And B.ҩ��ID=C.ҩ��ID And C.����=D.����"
                        Call OpenRecordset(rsTemp, "��ȡҩƷ�����Ϣ")
                        strƵ�� = Nvl(rsTemp!Ƶ��)
                        str�÷� = Nvl(rsTemp!�÷�)
                        str���� = Nvl(rsTemp!����)
                        str���� = Nvl(rsTemp!������λ)
                    End If
        
        
                    mstrOutput = Space(600)
                    str��ˮ�� = !��¼���� & !NO & 1 & !���
                    strסԺ�� = !����ID & "_1_" & !��ҳID
                    mlngReturn = FormularyEntry(g���˻�����Ϣ.����05, strסԺ��, IIf(!��¼״̬ = 2, -1, 1), int��Ŀ���, _
                            strҽԺ����, str��ˮ��, Format(!�Ǽ�ʱ��, "yyyyMMdd"), strҽ������, _
                            strҽ������, !�۸�, Abs(!����), Abs(!���), _
                            str����, str����, strƵ��, str�÷�, _
                            !��������, 0, str�������, Nvl(!ҽ��), _
                            UserInfo.����, "", IIf(dbl�Ը����� = 1, 1, 0), mstrOutput)
                    Call WriteBusinessLOG("FormularyEntry", g���˻�����Ϣ.����05 & "," & strסԺ�� & "," & IIf(!��¼״̬ = 2, -1, 1) & "," & int��Ŀ��� & "," & _
                            strҽԺ���� & "," & str��ˮ�� & "," & Format(!�Ǽ�ʱ��, "yyyyMMdd") & "," & strҽ������ & "," & _
                            strҽ������ & "," & !�۸� & "," & Abs(!����) & "," & Abs(!���) & "," & _
                            str���� & "," & str���� & "," & strƵ�� & "," & str�÷� & "," & _
                            !�������� & "," & 0 & "," & str������� & "," & Nvl(!ҽ��) & "," & _
                            UserInfo.���� & ",," & IIf(dbl�Ը����� = 1, 1, 0), mstrOutput)
                    If mlngReturn = -1 Then
                        MsgBox "�ϴ�����[" & !NO & "]��" & !��� & "����ϸʱ����" & vbCrLf & "��ϸ��Ϣ��" & vbCrLf & mstrOutput, vbInformation, gstrSysName
                        Call �ύ_ɽ��
                        Exit Function
                    End If
        
                    '���·�����ϸ�е�ͳ�������Ϣ
                    
                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & !ID & "," & _
                            Split(mstrOutput, "|")(1) - Split(mstrOutput, "|")(3) - Split(mstrOutput, "|")(4) & _
                            ",NULL,1,'" & strҽ������ & "',1,NULL)"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ���ֶ�")
                    
                    '>>>�����ϴ���־
                    
                    .MoveNext
                Loop
            End With
            Call �ύ_ɽ��
        End If
        '>> End �������
        
        rsCd.MoveNext
    Loop
    '> END �����˴���ϸ.
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    Call �ύ_ɽ��
Resume
End Function


Public Function סԺ�������_ɽ��(lng����ID As Long) As Boolean
    Dim lng����ID As Long
    Dim str���� As String
    Dim str���� As String
    Dim StrInput As String
    Dim rsTemp As New ADODB.Recordset
    Dim strסԺ�� As String  'ҽ��סԺ��,��ʽΪ ����ID_1_��ҳID
    Dim strҽ�����  As String
    Dim strNO As String
    Dim lng����ID As Long, lng��ҳID As Long
    Dim dbl�����ʻ� As Double
    
    On Error GoTo errHand
    
    'ȡ������ˮ��
    gstrSQL = "Select * From ���ս����¼ Where ����=2 And ��¼ID=" & lng����ID & " and ����=" & TYPE_ɽ��
    Call OpenRecordset(rsTemp, "ȡ������ˮ��")
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "û���ҵ�ԭʼ�����¼���޷�����סԺ���������"
        Exit Function
    End If
    
    strҽ����� = Nvl(rsTemp!��ע)
    lng����ID = rsTemp!����ID
    
    '��ȡ������ҳID����Ժ����
    gstrSQL = " Select A.��ҳID,A.��Ժ���� From ������ҳ A,������Ϣ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� and B.����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "��ȡ������ҳID����Ժ����")
    
    lng��ҳID = rsTemp!��ҳID
    strסԺ�� = lng����ID & "_1_" & lng��ҳID
    
    If strסԺ�� = "" Or strҽ����� = "" Then
        Err.Raise 9000, gstrSysName, "ԭʼ�����¼���׺Ų�ȫ���޷�����������������"
        סԺ�������_ɽ�� = False
        Exit Function
    End If
    
    'ȡ��֤����
    gstrSQL = "Select ����,���� From �����ʻ� Where ����=" & TYPE_ɽ�� & " And ����ID=" & lng����ID
    Call OpenRecordset(rsTemp, "����������")
    str���� = Nvl(rsTemp!����)
    str���� = Nvl(rsTemp!����)
    
    'ȡ������¼�Ľ���ID�����ݺ�
    gstrSQL = "select distinct A.NO,A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B where A.NO=B.NO and  A.��¼״̬=2 and B.ID=" & lng����ID
    Call OpenRecordset(rsTemp, "���²����Ľ���ID")
    lng����ID = rsTemp!ID
    strNO = "2" & rsTemp!NO
    
    
        '��ȡ�ʻ�ʵ��֧����
    gstrSQL = "Select Nvl(��Ԥ��,0) AS �����ʻ� From ����Ԥ����¼ Where ����ID=" & lng����ID & " And ��¼���� Not In (1,11) And ���㷽ʽ='�����ʻ�'"
    Call OpenRecordset(rsTemp, "��ȡ�ʻ�ʵ��֧����")
    If Not rsTemp.EOF Then
        dbl�����ʻ� = Nvl(rsTemp!�����ʻ�, 0)
    End If

    ''��ʵ����  �������������Ҫ��������¼�봰��,�����ݲ�����
    
    If �������(str����) = False Then
        סԺ�������_ɽ�� = False
        Exit Function
    End If
    
  
    '���ý������
    mstrOutput = Space(600)
    mlngReturn = ExpenseCalc(str����, -2, 1, strסԺ��, strҽ�����, _
                                   strNO, UserInfo.����, dbl�����ʻ�, mstrOutput)
    Call WriteBusinessLOG("ExpenseCalc", str���� & ", -2, 1," & strסԺ�� & "," & strҽ����� & "," & _
                                   strNO & "," & UserInfo.���� & "," & dbl�����ʻ�, Trim(mstrOutput))
    '�ɹ���,���汾�ν������
    
    If mlngReturn = 0 Then
        Call �ύ_ɽ��
        gstrSQL = "Select * From ���ս����¼ Where ����=2 And ��¼ID=" & lng����ID & " and ����=" & TYPE_ɽ��
        Call OpenRecordset(rsTemp, "ȡ������Ϣ")
        
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_ɽ�� & "," & lng����ID & "," & _
            Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
            0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
            -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & -1 * Nvl(rsTemp!����ͳ����, 0) & "," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & ",0,0," & _
            -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & Nvl(rsTemp!֧��˳���) & "'," & lng��ҳID & "," & rsTemp!��;���� & ",'" & Nvl(rsTemp!��ע) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "סԺ�������")
        
        סԺ�������_ɽ�� = True
    Else
        Call ����_ɽ��
        סԺ�������_ɽ�� = False
        Err.Raise 9000, gstrSysName, "�˷ѽ���ʧ�ܣ�" & vbCrLf & Trim(mstrOutput)
    End If
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function �޸�����_ɽ��(strOldPwd As String, strNewPwd As String) As Boolean
    �޸�����_ɽ�� = False
    mstrOutput = Space(600)
    
    mlngReturn = ChangePinEx(strOldPwd, strNewPwd, mstrOutput)
    If mlngReturn = 0 Then
        �޸�����_ɽ�� = True
        MsgBox "�����޸ĳɹ�!" & vbCrLf, vbInformation, gstrSysName
    Else
        �޸�����_ɽ�� = False
        MsgBox "�����޸�ʧ��!" & vbCrLf & Trim(mstrOutput), vbInformation, gstrSysName
    End If
End Function


