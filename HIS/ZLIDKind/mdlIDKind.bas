Attribute VB_Name = "mdlIDKind"
Option Explicit
Public gobjSquare As Object '���Ჿ��
Public gobjCardDatabase As Object  '�������е�clsDataBase��
Public gobjCards As Cards    '���еĿ�
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long, glngSys As Long
Public gstrAviPath As String, gstrVersion As String
Public gstrProductName As String
Public gstrDBUser As String   '��ǰ���ݿ��û�
Public gstrUnitName As String '�û���λ����
Public gobjParent As Object

'��������(zl9ComLib)
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjControl As Object

'ˢ������ȫ�ֱ���
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_F4 = vbKeyF4
Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20
Public glngInstanceCount As Long

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Function zlGetFromBedNumberToPatiID(ByVal cnOracle As ADODB.Connection, _
    ByVal lng����ID As Long, _
    ByVal str���� As String, Optional ByRef lng��ҳID As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݴ��Ż�ȡ����ID
    '����:lng��ҳID-���ص�ǰ���ŵ���ҳID
    '����:�ɹ����ز���ID,���򷵻�False
    '����:���˺�
    '����:2012-09-19 15:50:18
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim objDatabase As Object
    
    On Error GoTo errHandle
    Set objDatabase = GetCardSquareDataBaseObject(cnOracle)
    
    lng��ҳID = 0
    strSQL = _
    "   Select  A.����ID,A.��ҳID" & _
    "   From ������Ϣ A,��λ״����¼ C" & _
    "   Where  A.����ID=C.����ID And A.ͣ��ʱ�� is NULL " & _
    "           And C.����ID=[1] And C.����=[2] "
    
    If objDatabase Is Nothing Then
        Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID, str����)
    Else
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID, str����)
    End If
    If rsTemp.EOF Then zlGetFromBedNumberToPatiID = 0: Exit Function
    lng��ҳID = Val(Nvl(rsTemp!��ҳID))
    zlGetFromBedNumberToPatiID = Val(Nvl(rsTemp!����ID))
    Exit Function
errHandle:
    If Not objDatabase Is Nothing Then
        If objDatabase.ErrCenter() = 1 Then Resume
    Else
        If gobjComLib.ErrCenter() = 1 Then Resume
    End If
End Function

Private Function zlInitComponents(Optional lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ӿڲ���
    '����:���˺�
    '����:2012-08-16 11:09:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpand As String
    strExpand = lngCardTypeID
    If gobjSquare Is Nothing Then Exit Function
    '��ʼ�������㲿��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '���: frmMain-���õ�������
    '        lngModule-HIS����ģ���
    '       lngSys-�����ϵͳ��
    '       strDBUser-���ݿ��û���
    '       cnOracle -HIS/��������
    '       blnDeviceSet-�豸���õ��ó�ʼ��
    '       strExpand-��չ��Ϣ(��ѡ:ת�뿨���ID)
    zlInitComponents = gobjSquare.zlInitComponents(gobjParent, _
     glngModul, glngSys, gstrDBUser, _
      gcnOracle, False, strExpand)
End Function

Public Function zlInitCards(ByVal cnOracle As ADODB.Connection, ByVal RegType As gRegType) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-15 16:43:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objCard As Card, blnģ������ As Boolean
    Dim strValue As String
    On Error GoTo errHandle
    If zlCreateSquare(cnOracle) = False Then Exit Function
    'zlGetCards(ByVal BytType As Byte) As Cards
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ч�Ŀ�����
    '���:bytType-0-����ҽ�ƿ�;
    '             1-���õ�ҽ�ƿ�,
    '             2-���д��������˻���������
    '             3-���õ������˻���ҽ�ƿ�
    'Set rsTemp = gobjSquare.zlGetYLCards
    Set gobjCards = gobjSquare.zlGetCards(0)
    blnģ������ = False
    For Each objCard In gobjCards
        Call GetRegInFor(RegType, "ҽ�ƿ����\" & objCard.����, "�س���", strValue)
        Select Case strValue
            Case "����"
                objCard.���ų��� = objCard.���ų��� + IIf(objCard.�豸�Ƿ����ûس�, 0, 1)
            Case "����"
                objCard.���ų��� = objCard.���ų��� - IIf(objCard.�豸�Ƿ����ûس�, 1, 0)
        End Select
        If objCard.�Ƿ�ģ������ And objCard.���� And Not blnģ������ Then blnģ������ = True
    Next
    gobjCards.��ȱʡ������ = Not blnģ������
    zlInitCards = True
    Exit Function
errHandle:
    
    MsgBox Err.Description
End Function
Public Function zlGetKindCards(Optional strIDKindStr As String = "", Optional blnOnlyAccouct As Boolean = False, _
                                Optional NotAutoAppendKind As Boolean = False, Optional OnlyThreeCard As Boolean = False) As Cards
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ч�Ŀ�����
    '����: �ɹ�,������
    '����:���˺�
    '����:2012-08-15 16:58:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCards As Cards, objCard As Card
    Dim varData As Variant, i As Long, varTemp As Variant
    Dim blnFind As Boolean, j As Long
    Dim strKinds As String
    
    On Error GoTo errHandle
    If strIDKindStr = "" Then
        'ȱʡ���
        strIDKindStr = "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|���￨|0;��|�ֻ���|0"
    End If
    Set objCards = New Cards
    varData = Split(strIDKindStr, ";")
    j = 1
    strKinds = ""
    If Not OnlyThreeCard Then
        For i = 0 To UBound(varData)
            '����
            varTemp = Split(varData(i) & "||||||||||||", "|")
            If Trim(varTemp(1)) <> "" Then
                blnFind = False
                If Not gobjCards Is Nothing Then
                    For Each objCard In gobjCards
                        '76243,Ƚ����,2014-8-5,������������Ա����ȫ����ĸICʱ,Ĭ�Ͻ��䴦��Ϊϵͳ��Ĭ�ϵ�IC�����
                        If objCard.���� = Trim(varTemp(1)) _
                            Or (objCard.���� Like "*IC��*" And (varTemp(1) = "IC��" Or varTemp(1) = "IC����" Or varTemp(1) Like "*�ɣÿ�*") And objCard.ϵͳ) _
                            Or (objCard.���� Like "*���֤*" And (varTemp(1) = "�������֤" Or varTemp(1) = "���֤" Or varTemp(1) = "���֤��") And objCard.ϵͳ) Then
                            blnFind = True
                            If InStr(strKinds & ",", "," & objCard.�ӿ���� & ",") = 0 Then
                                strKinds = strKinds & "," & objCard.�ӿ����
                                If objCard.���� And Not objCard.���ѿ� Then
                                   objCards.Add objCard, "K" & objCard.�ӿ����
                                End If
                            End If
                            Exit For
                        End If
                    Next
                End If
               If blnFind = False Then
                    '����
                    Set objCard = New Card
                    '����1|ȫ��1|�Ƿ�ˢ��1|�����ID1|���ų���1|ȱʡ��־1(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�1(1-�����ʻ�;0-�������ʻ�)|
                    '��������1(�ڼ�λ���ڼ�λ����,��Ϊ������)|�Ƿ�ɨ��|�Ƿ�Ӵ�ʽ����|�Ƿ�ǽӴ�ʽ����
                    With objCard
                        .�ӿڱ��� = "-"
                        .���� = varTemp(1)
                        .���� = varTemp(0)
                        .�Ƿ�ˢ�� = Val(varTemp(2)) <> 1
                        .�ӿ���� = 0 ' IIf(Val(varTemp(3)) = 0, -j, Val(varTemp(3)))
                        .ȱʡ��־ = Val(varTemp(4)) = 1
                        .�Ƿ�����ʻ� = Val(varTemp(5)) = 1
                        .�������Ĺ��� = Trim(varTemp(6))
                        '85565,���ϴ�,2015/7/10:�������ʣ�ȱʡΪFasle
                        .�Ƿ�ɨ�� = Val(varTemp(7)) = 1
                        .�Ƿ�Ӵ�ʽ���� = Val(varTemp(8)) = 1
                        .�Ƿ�ǽӴ�ʽ���� = Val(varTemp(9)) = 1
                    End With
                    Err = 0: On Error Resume Next
                    objCards.Add objCard, "M" & objCard.����
                    If Err <> 0 Then Err = 0: On Error GoTo 0
                    j = j + 1
               End If
            End If
        Next
    End If
    'δ����ģ��������
    If NotAutoAppendKind = False Or OnlyThreeCard Then
        If Not gobjCards Is Nothing Then
            For Each objCard In gobjCards
                If InStr(1, strKinds & ",", "," & objCard.�ӿ���� & ",") = 0 And objCard.���� And Not objCard.���ѿ� Then
                    strKinds = strKinds & "," & objCard.�ӿ����
                    objCards.Add objCard, "K" & objCard.�ӿ����
                End If
            Next
        End If
    End If
    
    If Not gobjCards Is Nothing Then
        objCards.��ȱʡ������ = gobjCards.��ȱʡ������
        objCards.������ʾ = gobjCards.������ʾ
    End If
    Set zlGetKindCards = objCards
    
    Err = 0: On Error Resume Next
    Erase varData '�������
    
    Exit Function
errHandle:
    
    MsgBox Err.Description
End Function

Public Function zlGetPati(ByVal cnOracle As ADODB.Connection, _
    ByVal lng����ID As Long, ByRef objPati As PatiInfor, _
    ByRef strErrMsg As String, Optional strOtherName As String = "", _
    Optional strOtherValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID,���»�ȡ����
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-04-06 18:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere As String
    Dim objDatabase As Object
    
    On Error GoTo errHandle
    
    Set objDatabase = GetCardSquareDataBaseObject(cnOracle)
    Set objPati = New PatiInfor
    
    '�������֤���д˲���û��
    If strOtherName = "" Then
        strWhere = " And ����ID=[1]"
    ElseIf strOtherName = "�����" Then
        strWhere = " And �����=[2]"
    ElseIf strOtherName = "סԺ��" Then
        strWhere = " And ����ID=(Select Max(����ID) From ������ҳ Where סԺ�� = [2])"
    Else
        strWhere = " And " & strOtherName & "=[3]"
    End If
    strSQL = "" & _
    "   Select a.����id, a. �����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ,p.���� as ҽ�Ƹ��ʽ����, a. ����, a.�Ա�, a. ����, a.��������, a.�����ص�, a.���֤��, a.����֤��, a.���, " & _
    "        a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.�໤��, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, " & _
    "        a.��ͬ��λid, a.������λ, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.������, a.������, a.��������, a.����ʱ��, a.����״̬, a.��������, a.��Ժ, a.Ic����, " & _
    "        a.������, a.ҽ����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, '' as ����, 0As ��״̬,'' as ����, '' as ��ʧ��ʽ, " & _
    "        a.�������� as ��������,sysdate as ��ʧʱ��, 0  as ��ʧ��Ч����,sysdate as ��ǰʱ��,a.�ֻ���,a.����,B.���� ��������" & _
    "   From ������Ϣ A,������� B,ҽ�Ƹ��ʽ P" & _
    "   Where A.���� = B.���(+) And a.ҽ�Ƹ��ʽ=P.����(+)  " & strWhere
    
    
    If Not objDatabase Is Nothing Then
        Set rsTemp = objDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID, Val(strOtherValue), strOtherValue)
    Else
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID, Val(strOtherValue), strOtherValue)
    End If
    If rsTemp.EOF Then Exit Function
    objPati.����ID = rsTemp!����ID
    objPati.����� = IIf(Val(Nvl(rsTemp!�����)) = 0, "", Nvl(rsTemp!�����))
    objPati.���� = Nvl(rsTemp!����)
    objPati.�Ա� = Nvl(rsTemp!�Ա�)
    objPati.���� = Nvl(rsTemp!����)
    objPati.�������� = Format(rsTemp!��������, "yyyy-mm-dd")
    objPati.������ַ = Nvl(rsTemp!�����ص�)
    objPati.���֤�� = Nvl(rsTemp!���֤��)
    objPati.����֤�� = Nvl(rsTemp!����֤��)
    objPati.ְҵ = Nvl(rsTemp!ְҵ)
    objPati.�ѱ� = Nvl(rsTemp!�ѱ�)
    objPati.���� = Nvl(rsTemp!����)
    objPati.���� = Nvl(rsTemp!����)
    objPati.ѧ�� = Nvl(rsTemp!ѧ��)
    objPati.����״�� = Nvl(rsTemp!����״��)
    objPati.���� = Nvl(rsTemp!����״��)
    objPati.��ͥ��ַ = Nvl(rsTemp!��ͥ��ַ)
    objPati.��ͥ�绰 = Nvl(rsTemp!��ͥ�绰)
    objPati.��ͥ�ʱ� = Nvl(rsTemp!��ͥ��ַ�ʱ�)
    objPati.�໤�� = Nvl(rsTemp!�໤��)
    objPati.��ϵ�� = Nvl(rsTemp!��ϵ������)
    objPati.��ϵ�˹�ϵ = Nvl(rsTemp!��ϵ�˹�ϵ)
    objPati.��ϵ�˵�ַ = Nvl(rsTemp!��ϵ�˵�ַ)
    objPati.��ϵ�˵绰 = Nvl(rsTemp!��ϵ�˵绰)
    objPati.������λ = Nvl(rsTemp!������λ)
    objPati.������λ�绰 = Nvl(rsTemp!��λ�绰)
    objPati.������λ�ʱ� = Nvl(rsTemp!��λ�ʱ�)
    objPati.������λ������ = Nvl(rsTemp!��λ������)
    objPati.������λ�������ʻ� = Nvl(rsTemp!��λ�ʺ�)
    objPati.���ڵ�ַ = Nvl(rsTemp!���ڵ�ַ)
    objPati.���ڵ�ַ�ʱ� = Nvl(rsTemp!���ڵ�ַ�ʱ�)
    objPati.���� = Nvl(rsTemp!����)
    objPati.���� = Nvl(rsTemp!����)
    objPati.ҽ�Ƹ��ʽ���� = Nvl(rsTemp!ҽ�Ƹ��ʽ����)
    objPati.ҽ�Ƹ��ʽ = Nvl(rsTemp!ҽ�Ƹ��ʽ)
    objPati.�������� = Nvl(rsTemp!��������)
    objPati.���￨�� = Nvl(rsTemp!���￨��)
    objPati.�ֻ��� = Nvl(rsTemp!�ֻ���)
    objPati.���� = Val(Nvl(rsTemp!����))
    objPati.�������� = Trim(Nvl(rsTemp!��������))
    zlGetPati = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function

Public Function FromObjectToCard(ByVal objCard As Object) As Card
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Object���󻻳�Card����
    '����:�ɹ�Card����
    '����:���˺�
    '����:2013-10-23 18:03:52
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As New Card
    With objCard
        objTemp.�ӿ���� = .�ӿ����
        objTemp.�ӿڱ��� = .�ӿڱ���
        objTemp.���� = .����
        objTemp.���� = .����
        objTemp.ǰ׺�ı� = .ǰ׺�ı�
        objTemp.���ų��� = .���ų���
        objTemp.ȱʡ��־ = .ȱʡ��־
        objTemp.ϵͳ = .ϵͳ
        objTemp.�Ƿ��ϸ���� = .�Ƿ��ϸ����
        objTemp.�Ƿ��Զ���ȡ = .�Ƿ��Զ���ȡ
        objTemp.�Զ���ȡ��� = .�Զ���ȡ���
        objTemp.���ƿ� = .���ƿ�
        objTemp.�Ƿ�����ʻ� = .�Ƿ�����ʻ�
        objTemp.�Ƿ�ȫ�� = .�Ƿ�ȫ��
        objTemp.�����ظ�ʹ�� = .�����ظ�ʹ��
        objTemp.���㷽ʽ = .���㷽ʽ
        objTemp.�ӿڳ����� = .�ӿڳ�����
        objTemp.�ض���Ŀ = .�ض���Ŀ
        objTemp.���� = .����
        objTemp.��ע = .��ע
        objTemp.�������Ĺ��� = .�������Ĺ���
        objTemp.�Ƿ����� = .�Ƿ�����
        objTemp.���볤�� = .���볤��
        objTemp.���볤������ = .���볤������
        objTemp.������� = .�������
        objTemp.������������ = .������������
        objTemp.�Ƿ�ȱʡ���� = .�Ƿ�ȱʡ����
        objTemp.�Ƿ��ƿ� = .�Ƿ��ƿ�
        objTemp.�Ƿ񷢿� = .�Ƿ񷢿�
        objTemp.�Ƿ�д�� = .�Ƿ�д��
        objTemp.�������� = .��������
        '77872,���ϴ�,2014/9/15:�Ƿ�֧��ת�ʼ�����
        objTemp.�Ƿ�ת�ʼ����� = .�Ƿ�ת�ʼ�����
        objTemp.�Ƿ�ˢ�� = .�Ƿ�ˢ��    '85565,���ϴ�,2015/7/10:��������
        objTemp.�Ƿ�ɨ�� = .�Ƿ�ɨ��
        objTemp.�Ƿ�Ӵ�ʽ���� = .�Ƿ�Ӵ�ʽ����
        objTemp.�Ƿ�ǽӴ�ʽ���� = .�Ƿ�ǽӴ�ʽ����
        objTemp.�Ƿ�ֿ����� = .�Ƿ�ֿ�����
        objTemp.�Ƿ��˿��鿨 = .�Ƿ��˿��鿨
        objTemp.�Ƿ�ȱʡ���� = .�Ƿ�ȱʡ����
    End With
    Set FromObjectToCard = objTemp
End Function

Public Function FromXMLPati(ByVal strPatiXml As String, ByRef strErrMsg As String) As PatiInfor
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��XML�л�ȡ������Ϣ
    '����:������Ϣ����
    '����:���˺�
    '����:2012-08-22 11:43:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strOutCardNO As String, strOutPatiInforXML As String
    Dim objNode As MSXML2.IXMLDOMElement, strExpand As String
    Dim objTempNode As MSXML2.IXMLDOMElement
    Dim strTmp As String, strValue As String
    Dim objPati As New PatiInfor
    Dim objXML As New objXML
    On Error GoTo errHandle
    
   If strPatiXml = "" Then Exit Function
   '���ز�����Ϣ
    If objXML.zlXML_Init = False Then Exit Function
    If objXML.zlXML_LoadXMLToDOMDocument(strPatiXml, False, strErrMsg) = False Then Exit Function
    '    ��ʶ    ��������    ����    ����    ˵��
    '    ����    Varchar2    20
    Call objXML.zlXML_GetNodeValue("����", , strValue)
    objPati.���� = strValue
    '    ����    Varchar2    64
    Call objXML.zlXML_GetNodeValue("����", , strValue)
    objPati.���� = strValue
    '    �Ա�    Varchar2    4
    Call objXML.zlXML_GetNodeValue("�Ա�", , strValue)
    objPati.�Ա� = strValue
    '    ����    Varchar2    10
    Call objXML.zlXML_GetNodeValue("����", , strValue)
    objPati.���� = strValue
    '    ��������    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    Call objXML.zlXML_GetNodeValue("��������", , strValue)
    objPati.�������� = strValue
    '    �����ص�    Varchar2    50
    Call objXML.zlXML_GetNodeValue("�����ص�", , strValue)
    objPati.������ַ = strValue
    '    ���֤��    VARCHAR2    18
    Call objXML.zlXML_GetNodeValue("���֤��", , strValue)
    objPati.���֤�� = strValue
    If objPati.�������� = "" And strValue <> "" Then
        strTmp = gobjCommFun.GetIDCardDate(strValue)
        If IsDate(strTmp) Then objPati.�������� = strTmp
    End If
    '    ����֤��    Varchar2    20
    Call objXML.zlXML_GetNodeValue("����֤��", , strValue)
    objPati.����֤�� = strValue
    '    ְҵ    Varchar2    80
    Call objXML.zlXML_GetNodeValue("ְҵ", , strValue)
    objPati.ְҵ = strValue
    '    ����    Varchar2    20
    Call objXML.zlXML_GetNodeValue("����", , strValue)
    objPati.���� = strValue
    '    ����    Varchar2    30
    Call objXML.zlXML_GetNodeValue("����", , strValue)
    objPati.���� = strValue
    '    ѧ��    Varchar2    10
    Call objXML.zlXML_GetNodeValue("ѧ��", , strValue)
    objPati.ѧ�� = strValue
    '    ����״��    Varchar2    4
    Call objXML.zlXML_GetNodeValue("����״��", , strValue)
    objPati.����״�� = strValue
    
    '    ����    Varchar2    30
    Call objXML.zlXML_GetNodeValue("����", , strValue)
    objPati.���� = strValue
    '    ��ͥ��ַ    Varchar2    50
    Call objXML.zlXML_GetNodeValue("��ͥ��ַ", , strValue)
    objPati.��ͥ��ַ = strValue
     '    ���ڵ�ַ    Varchar2    50
    Call objXML.zlXML_GetNodeValue("���ڵ�ַ", , strValue)
    objPati.���ڵ�ַ = strValue
    '    ��ͥ�绰    Varchar2    20
    Call objXML.zlXML_GetNodeValue("��ͥ�绰", , strValue)
    objPati.��ͥ�绰 = strValue
    '    ��ͥ��ַ�ʱ�    Varchar2    6
    Call objXML.zlXML_GetNodeValue("��ͥ��ַ�ʱ�", , strValue)
    objPati.��ͥ�ʱ� = strValue
    '    �໤��  Varchar2    64
    Call objXML.zlXML_GetNodeValue("�໤��", , strValue)
    objPati.�໤�� = strValue
    
    '    ��ϵ������  Varchar2    64
    Call objXML.zlXML_GetNodeValue("��ϵ������", , strValue)
    objPati.��ϵ�� = strValue
    '    ��ϵ�˹�ϵ  Varchar2    30
    Call objXML.zlXML_GetNodeValue("��ϵ�˹�ϵ", , strValue)
    objPati.��ϵ�˹�ϵ = strValue
    '    ��ϵ�˵�ַ  Varchar2    50
    Call objXML.zlXML_GetNodeValue("��ϵ�˵�ַ", , strValue)
    objPati.��ϵ�˵�ַ = strValue
    '    ��ϵ�˵绰  Varchar2    20
    Call objXML.zlXML_GetNodeValue("��ϵ�˵绰", , strValue)
    objPati.��ϵ�˵绰 = strValue
    '    ������λ    Varchar2    100
    Call objXML.zlXML_GetNodeValue("������λ", , strValue)
    objPati.������λ = strValue
    '    ��λ�绰    Varchar2    20
    Call objXML.zlXML_GetNodeValue("��λ�绰", , strValue)
    objPati.������λ�绰 = strValue
    '    ��λ�ʱ�    Varchar2    6
    Call objXML.zlXML_GetNodeValue("��λ�ʱ�", , strValue)
    objPati.������λ�ʱ� = strValue
    '    ��λ������  Varchar2    50
    Call objXML.zlXML_GetNodeValue("��λ������", , strValue)
    objPati.������λ������ = strValue
    '    ��λ�ʺ�    Varchar2    20
    Call objXML.zlXML_GetNodeValue("��λ�ʺ�", , strValue)
   'txt��λ�ʺ�.Text = strValue
    objPati.������λ�������ʻ� = strValue
    '    �ֻ���    Varchar2    20
    Call objXML.zlXML_GetNodeValue("�ֻ���", , strValue)
    objPati.�ֻ��� = strValue
    '    ��Ƭ�ļ�    Varchar2    20
    Call objXML.zlXML_GetNodeValue("��Ƭ�ļ�", , strValue)
    objPati.��Ƭ�ļ� = strValue
    '    ��Ƭ
    If Trim(strValue) <> "" Then
        Err = 0: On Error Resume Next
        objPati.��Ƭ = LoadPicture(strValue)
        If objPati.��Ƭ = 0 Then objPati.��Ƭ = Nothing
        Err = 0: On Error GoTo errHandle
    End If
    Set FromXMLPati = objPati
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlTranErrInfor(strErrMsg) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Դ�����Ϣ���и�ʽ��
    '����: ���ر���ʽ���Ĵ�����Ϣ
    '����:���˺�
    '����:2012-08-22 14:47:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlTranErrInfor = strErrMsg
End Function

Public Function zlCreateSquare(ByVal cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '����:���˺�
    '����:2012-08-15 16:40:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If Not gobjSquare Is Nothing Then zlCreateSquare = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjSquare = CreateObject("zl9CardSquare.clsCardsquare")
    If Err <> 0 Then Err = 0: Exit Function
    Call gobjSquare.zlInitComponents(gobjParent, glngModul, glngSys, gstrDBUser, cnOracle, False, strExpend)
    '��ʼ�������ɹ�,����Ϊ�����ڴ���
    zlCreateSquare = True
End Function

Public Function zlCreateSquareDataBaseObject(ByVal cnOracle As ADODB.Connection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����clsDataBase����(zlCardSquare����)
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-06-03 11:02:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If Not gobjCardDatabase Is Nothing Then zlCreateSquareDataBaseObject = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjCardDatabase = CreateObject("zl9CardSquare.clsDataBase")
    If Err <> 0 Then Err = 0: Exit Function
    Call gobjCardDatabase.InitCommon(gcnOracle)
    zlCreateSquareDataBaseObject = True
    Err = 0: On Error GoTo 0
End Function
Public Function GetCardSquareDataBaseObject(cnOracle) As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨�����е�CardSquare����
    '����:���˺�
    '����:2015-06-03 11:22:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjCardDatabase Is Nothing Then
        If zlCreateSquareDataBaseObject(cnOracle) = False Then Exit Function
    End If
    Call gobjCardDatabase.InitCommon(cnOracle)
    Set GetCardSquareDataBaseObject = gobjCardDatabase
End Function

Public Function zlCloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    zlCloseWindows = Forms.Count = 0
End Function

Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ͷ���Դ
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-02-13 10:30:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'ʵ����Ϊ0ʱ���ŷ���Դ
    If glngInstanceCount > 0 Then Exit Function
    Call zlCloseWindows '�ͷŴ�����Դ
    If Not gobjSquare Is Nothing Then Set gobjSquare = Nothing
    If Not gobjCardDatabase Is Nothing Then Set gobjCardDatabase = Nothing
    If Not gobjCards Is Nothing Then Set gobjCards = Nothing
    If Not gobjParent Is Nothing Then Set gobjParent = Nothing
    If Not gobjComLib Is Nothing Then Set gobjComLib = Nothing
    If Not gobjCommFun Is Nothing Then Set gobjCommFun = Nothing
    If Not gobjDatabase Is Nothing Then Set gobjDatabase = Nothing
    If Not gobjControl Is Nothing Then Set gobjControl = Nothing
    zlReleaseResources = True
End Function
