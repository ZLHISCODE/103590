Attribute VB_Name = "mdlPubExpence"
Option Explicit
'*********************************************************************************************************************************************
'����:������ù��������ķ���������(zlPublicExpense�����Ĵ���)
'�ӿ�˵��:
'    1. zlGetPubExpenseObject-��ȡ���ù�����������
'    2. zlInitPriceGrade-��ʼ���۸�ȼ���Ϣ
'    3. zlGetPriceGrade:��ȡ�۸�ȼ���Ϣ
'    4. zlPatiIdentify:���������֤(����ˢ����֤)
'    5. zlVerifyPassWord:����������֤��
'����:���˺�
'����:2019-01-25 09:51:46
'*********************************************************************************************************************************************
Public gobjPubExpense As Object  '���ù�������
Public gintPriceGradeStartType As Integer
Public gstrҩƷ�۸�ȼ� As String
Public gstr���ļ۸�ȼ� As String
Public gstr��ͨ�۸�ȼ� As String
Public Function zlGetPubExpenseObject(ByRef objPubExpense As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ù�����������
    '����:objPubExpense-���ع������ò�������
    '����:��ȡ����true,���򷵻�False
    '����:���˺�
    '����:2019-01-25 09:57:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    If Not gobjPubExpense Is Nothing Then Set objPubExpense = gobjPubExpense: zlGetPubExpenseObject = True: Exit Function
    
    Err = 0: On Error Resume Next
    If gobjPubExpense Is Nothing Then
        Set gobjPubExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Err <> 0 Then
            MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Err.Clear: On Error GoTo 0
            Exit Function
        End If
    End If
    
    Err.Clear:  On Error GoTo errHandle
    If gobjPubExpense Is Nothing Then Exit Function
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If gobjPubExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
         Exit Function
    End If
    Set objPubExpense = gobjPubExpense: zlGetPubExpenseObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Public Function zlInitPriceGrade() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����õļ۸�ȼ�
    '���:
    '����:��ʼ���ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-01-25 10:00:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExpence As Object
    
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    
    gintPriceGradeStartType = gobjPubExpense.zlGetPriceGradeStartType()
    If gintPriceGradeStartType = 0 Then zlInitPriceGrade = True: Set objPubExpence = Nothing: Exit Function
    '��ȡվ��۸�ȼ�
    Call gobjPubExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    Set objPubExpence = Nothing:
    zlInitPriceGrade = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
 
Public Function zlPatiIdentify(ByVal lngModlue As Long, ByVal frmMain As Object, ByVal lng����ID As Long, ByVal curMoney As Currency, _
    Optional ByVal bln�˷� As Boolean = False, Optional ByVal bytDepositShowMode As Byte = 0, Optional ByVal lngDefaultCardTypeID As Long = 0, _
    Optional ByVal blnFamilyMoney As Boolean, Optional ByVal blnOlnyFamilyIDs As Boolean, Optional strFamilyPatiIDs_Out As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ����֤
    '���:lngModlue-ģ���
    '     dblMoney-���
    '     lng����ID-����ID
    '     bln�˷�-��ǰ�Ƿ��˷Ѳ���
    '     bytDepositShowMode- Ԥ����ʾ��ʽ(0-��������ʾ;1-ֻ��ʾ�������;2-ֻ��ʾסԺ���)
    '     lngDefaultCardTypeID-ȱʡ��ˢ�����
    '     blnFamilyMoney-�Ƿ��ȡ����Ԥ�����
    '     blnOlnyFamilyIDs-true:���鿨��ֻ��ȡ����IDs;False-��Ҫ��ȡ���鿨(��Ч���������ּ��ݲ�����ɾ��)
    '����:strFamilyPatiIDs-���˼���ID,����ö��ŷָ���79868
    '����:�����֤�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-10-24 14:55:59
    '˵��:
    '   һ�������鿨�����������bln�˷�=falseʱ):
    '       1.������ˢ����֤,ֱ�ӷ���True
    '       2.��������ʱ,��Ҫ����ˢ����֤,ͬʱ��Ҫ��������(������ʱ,���Ҫ���������)
    '       3.��������ʱ,�������������(ֻҪ����һ�ſ��������,�ʹ��������������)�������ˢ������������,�������,��ֻˢ����֤
    '       4.NԪ������֧��,��ʾ����������NԪ��ֻˢ����֤,����������;�������ˢ������������
    '  �����˷��鿨��bln�˷�=trueʱ):
    '       1.������ˢ�����ƣ�ֱ�ӷ���true
    '       2.�����˷�ʱ,��Ҫ����ˢ����֤,ͬʱ��Ҫ��������(������ʱ,���Ҫ���������)
    '       3.�����˷�ʱ,�������������(ֻҪ����һ�ſ��������,�ʹ��������������)�������ˢ������������,�������,��ֻˢ����֤
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExpence As Object
    
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    
    zlPatiIdentify = objPubExpence.zlPatiIdentify(lngModlue, frmMain, lng����ID, curMoney, bln�˷�, bytDepositShowMode, lngDefaultCardTypeID, _
                                               blnFamilyMoney, blnOlnyFamilyIDs, strFamilyPatiIDs_Out)
    Set objPubExpence = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlGetPriceGrade(ByVal strվ�� As String, _
    ByVal lng����ID As Long, ByVal lng��ҳid As Long, _
    Optional ByVal strҽ�Ƹ��ʽ As String, _
    Optional ByRef strҩƷ�۸�ȼ�_Out As String, _
    Optional ByRef str���ļ۸�ȼ�_Out As String, _
    Optional ByRef str��ͨ��Ŀ�۸�ȼ�_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ����ҽ�Ƹ��ʽ��վ�㣬��ȡ��Ӧ�ļ۸�ȼ�
    '���:strվ��-��½��վ�㣬���봫�룬����NULLʱ���۸�ȼ�Ϊ���ؿ�
    '     lng����ID-����ID
    '     lng��ҳID-��ҳID
    '     strҽ�Ƹ��ʽ:�������ǿգ����Դ���ҽ�Ƹ��ʽ_In��ʽ����ȡ�۸�ȼ�;�����Բ���ID_In����ҳID����ȡ��Ӧ�Ĳ��˵�ҽ�Ƹ��ʽ��
    
    '����:strҩƷ�۸�ȼ�_out-����ҩƷ�۸�ȼ�
    '     str���ļ۸�ȼ�_out-�������ļ۸�ȼ�
    '     str��ͨ��Ŀ�۸�ȼ�_out-������ͨ�շ���Ŀ�۸�ȼ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2016-07-29 16:10:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExpence As Object
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    zlGetPriceGrade = objPubExpence.zlGetPriceGrade(strվ��, lng����ID, lng��ҳid, strҽ�Ƹ��ʽ, strҩƷ�۸�ȼ�_Out, str���ļ۸�ȼ�_Out, str��ͨ��Ŀ�۸�ȼ�_out)
    Set objPubExpence = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlVerifyPassWord(frmParent As Object, ByVal strPass As String, _
    Optional ByVal strName As String, Optional ByVal strSex As String, _
    Optional ByVal strOld As String, Optional blnPassEncode As Boolean = True) As Boolean
    '���ܣ������������֤
    '������frmParent=��ʾ�ĸ�����
    '      strPass=��ȷ������
    '      strName,strSex,strOld=��ѡ�����������������Ա����䣬��������ʱ����ʾ�������
    '      blnPassEncode-strPass�Ƿ���ļ��ܴ�
    '���أ�True=������֤ͨ��,False=ȡ�����룬������3��������������
    Dim objPubExpence As Object
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    zlVerifyPassWord = objPubExpence.zlVerifyPassWord(frmParent, strPass, strName, strSex, strOld, blnPassEncode)
    Set objPubExpence = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Public Function zlGetErrSwapInfoByJsonString(ByVal strJson As String, ByRef cllSwapInfo_out As Collection, ByRef cllExpends_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����쳣������Ϣ��Json������ȡ�쳣��Ϣ
    '���:
    '����:cllSwapInfo_out-���صĽ�����Ϣ:����,�����ID,������ˮ��,����˵��,���׽��,��ά��,֧����ʽ,����ժҪ
    '     cllExpend_out
    '          |-cllExpend:-��������,��������
    '           ��ʽ:array(����,ֵ),"_����"

    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExpence As Object
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    zlGetErrSwapInfoByJsonString = objPubExpence.zlGetErrSwapInfoByJsonString(strJson, cllSwapInfo_out, cllExpends_Out)
    Set objPubExpence = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function
Public Function zlGetErrSwapInfoByErrID(ByVal lng�쳣ID As String, ByRef rsErrData_Out As ADODB.Recordset, _
    ByRef cllSwapInfo_out As Collection, ByRef cllExpends_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����쳣ID��ȡ��ȡ�쳣��Ϣ
    '���:
    '����:
    '     rsErrData_Out-�쳣���ݼ�
    '     cllSwapInfo_out-����,�����ID,������ˮ��,����˵��,���׽��,��ά��,֧����ʽ,����ժҪ
    '     cllExpend_out
    '          |-cllExpend:-��������,��������
    '           ��ʽ:array(����,ֵ),"_����"
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 19:04:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPubExpence As Object
    On Error GoTo errHandle
    If zlGetPubExpenseObject(objPubExpence) = False Then Exit Function
    zlGetErrSwapInfoByErrID = objPubExpence.zlGetErrSwapInfoByErrID(lng�쳣ID, rsErrData_Out, cllSwapInfo_out, cllExpends_Out)
    Set objPubExpence = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
 

