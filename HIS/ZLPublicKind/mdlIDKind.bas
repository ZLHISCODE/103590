Attribute VB_Name = "mdlIDKind"
Option Explicit
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
Public gobjPubOneCard As clsPublicOneCard   'һ��ͨ����
Public gblnIsObjRegisterAlone As Boolean

Public Function zlGetPubOneCard(ByRef cnOracle As ADODB.Connection, ByRef objPubOneCard_Out As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡһ��ͨ���ݷ��ʶ���
    '���:
    '����:objOneDataObject_Out-����һ��ͨ���ݷ��ʶ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-12-04 14:10:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    
    On Error GoTo errHandle
    
    If Not gobjPubOneCard Is Nothing Then Set objPubOneCard_Out = gobjPubOneCard: zlGetPubOneCard = True:  Exit Function
    Set objPubOneCard_Out = New clsPublicOneCard
    zlGetPubOneCard = objPubOneCard_Out.zlInitComponents(gobjParent, glngModul, glngSys, gstrDBUser, cnOracle, False, strExpend, gblnIsObjRegisterAlone)
    Set gobjPubOneCard = objPubOneCard_Out
    Exit Function
errHandle:
        If ErrCenter = 1 Then Resume
End Function

Public Function ErrCenter() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '����:���˺�
    '����:2018-12-05 11:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(gcnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    ErrCenter = gobjPubOneCard.ErrCenter
End Function

Public Sub WritLog(ByVal strDev As String, strInput As String, strOutPut As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��¼��־
    '����:���˺�
    '����:2018-12-05 11:35:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(gcnOracle, gobjPubOneCard) = False Then Exit Sub
    End If
    Call gobjPubOneCard.WritDebugLog(strDev, strInput, strOutPut)
End Sub

Public Function zlGetPatiIDFromBedNumber(ByVal lng����ID As Long, ByVal str���� As String, Optional ByRef lng��ҳID As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݴ��Ż�ȡ����ID
    '����:lng��ҳID-���ص�ǰ���ŵ���ҳID
    '����:�ɹ����ز���ID,���򷵻�False
    '����:���˺�
    '����:2012-09-19 15:50:18
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(gcnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    zlGetPatiIDFromBedNumber = gobjPubOneCard.zlGetPatiIDFromBedNumber(lng����ID, str����, lng��ҳID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function zlInitComponents(Optional lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ӿڲ���
    '����:���˺�
    '����:2012-08-16 11:09:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpand As String
    strExpand = lngCardTypeID
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(gcnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    
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
    zlInitComponents = gobjPubOneCard.zlInitComponents(gobjParent, glngModul, glngSys, gstrDBUser, gcnOracle, False, strExpand)
End Function


Public Function zlInitCards(ByVal cnOracle As ADODB.Connection, ByVal RegType As gRegType) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-15 16:43:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnģ������ As Boolean, strValue As String, objCard As Card
    
    On Error GoTo errHandle
    
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(cnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    
    'zlGetCards(ByVal BytType As Byte) As Cards
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ч�Ŀ�����
    '���:bytType-0-����ҽ�ƿ�;
    '             1-���õ�ҽ�ƿ�,
    '             2-���д��������˻���������
    '             3-���õ������˻���ҽ�ƿ�
    'Set rsTemp = gobjSquare.zlGetYLCards
    Set gobjCards = gobjPubOneCard.zlGetCards(0)
    
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
    If ErrCenter = 1 Then Resume
End Function
Public Function GetPatiInforFromPatiID(ByVal cnOracle As ADODB.Connection, ByVal lng����ID As Long, ByRef objPati As clsPatiInfor, _
    ByRef strErrMsg As String, Optional strOtherName As String = "", Optional strOtherValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID,���»�ȡ����
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-04-06 18:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(cnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    GetPatiInforFromPatiID = gobjPubOneCard.zlGetPatiInforFromPatiID(lng����ID, objPati, strErrMsg, strOtherName, strOtherValue)
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function
Public Function zlGetPatiInforFromXML(ByVal cnOracle As ADODB.Connection, ByVal strPatiXml As String, _
    ByRef objPatiInfor_Out As clsPatiInfor, ByRef strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��XML�л�ȡ������Ϣ
    '���:strPatiXml-������ϢXML
    '
    '����:objPatiInfor_Out-���ز�����Ϣ����
    '      strErrMsg_Out-���ش�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-12-05 14:29:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(cnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    zlGetPatiInforFromXML = gobjPubOneCard.zlGetPatiInforFromXML(strPatiXml, strErrMsg_out, objPatiInfor_Out)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatiIDFromCardType(ByVal cnOracle As ADODB.Connection, ByVal strCardType As String, ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef lng����ID As Long, _
    Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, Optional objCtl As Object = Nothing, Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0, Optional ByVal blnNotCheckValidDate As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����ҽ�����Ϳ���,��ȡ��Ӧ�Ĳ���ID
    '���:strCardType-�����,���Ϊ����,��Ϊ�����ID,���Ϊ�ַ�,��Ϊ�������
    '       strCardNo-����
    '       blnNotShowErrMsg-����ʾ�������ʾ��Ϣ
    '       frmMain-���õ�������
    '       objCtl-���õĿؼ�
    '       blnShowMergePati-�����ֶ�����������Ĳ���ʱ,�Ƿ���ʾ�ϲ����ܰ�ť
    '       blnOnlyContractPati-ǩԼ����
    '       blnUserCancel-ѡ�����У��û�ѡ����ȡ��
    '       lngShowCardNoTypeID-���˳���������Ϣʱ������ѡ��������ʾ�Ŀ��ŵĿ����ID,0-��ʾ����ʾ���ţ�>0��ʾ��ʾָ����������ID
    '       blnNotCheckValidDate-�Ƿ�Կ���ֹʹ��ʱ����м��,true-�������ֹʹ��ʱ��,false-���
    '����:strErrMsg-���صĴ�����Ϣ
    '       lng����ID-���صĲ���ID
    '       strCardPass-���ؿ��ŵ�����
    '       lngCardTypeID-���ؿ����ID(0��ʾ����ȷ�������ID)
    '����:��ȡ����ID�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-14 17:07:51
    '˵��:ֻ�д���ҽ�����Ĳŵ��ô˺���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjPubOneCard Is Nothing Then
        If zlGetPubOneCard(cnOracle, gobjPubOneCard) = False Then Exit Function
    End If
    If gobjPubOneCard.zlIsExistOraConnect = False Then Exit Function
    
    GetPatiIDFromCardType = gobjPubOneCard.zlGetPatiID(strCardType, strCardNo, blnNotShowErrMsg, lng����ID, _
        strCardPassWord, strErrMsg, lngCardTypeID, objCtl, frmMain, blnShowMergePati, blnOnlyContractPati, _
        blnCertificate, blnUserCancel, lngShowCardNoTypeID, blnNotCheckValidDate)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
Public Function zlTranErrInfor(strErrMsg) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Դ�����Ϣ���и�ʽ��
    '����: ���ر���ʽ���Ĵ�����Ϣ
    '����:���˺�
    '����:2012-08-22 14:47:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlTranErrInfor = strErrMsg
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
    If Not gobjCards Is Nothing Then Set gobjCards = Nothing
    If Not gobjParent Is Nothing Then Set gobjParent = Nothing
    If Not gobjPubOneCard Is Nothing Then Set gobjPubOneCard = Nothing
    zlReleaseResources = True
End Function
