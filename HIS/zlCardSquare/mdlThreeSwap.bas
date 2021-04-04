Attribute VB_Name = "mdlThreeSwap"
Option Explicit

Public Function CheckThreeBalanceToCash(frmMain As Object, ByVal lngModule As Long, _
    cllThreeSwapCards As Collection, ByVal objCard As Card) As Boolean
    '���������ּ��
    Dim str����Ա As String
    
    On Error GoTo errHandle
    If Not (objCard.�ӿ���� > 0 And Not objCard.���ѿ�) Then CheckThreeBalanceToCash = True: Exit Function
    If CardDelCash(cllThreeSwapCards, objCard.�ӿ����, objCard) Then CheckThreeBalanceToCash = True: Exit Function
    If CardDefaultCash(cllThreeSwapCards, objCard.�ӿ����, objCard) = False Then '���������֣�ͬʱȱʡ�����֣�������ǿ������
        ShowMsgbox objCard.���� & "������ǿ����Ϊ�������㷽ʽ��"
        Exit Function
    End If
    
    If zlstr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "�����˿�ǿ������") Then
        If MsgBox(objCard.���� & "��֧�����֣���ȷ��Ҫ����ǿ��������", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        str����Ա = zlDatabase.UserIdentifyByUser(frmMain, objCard.���� & "ǿ�����֣�Ȩ����֤��", _
            glngSys, lngModule, "�����˿�ǿ������", , True)
        If str����Ա = "" Then Exit Function
    End If
    CheckThreeBalanceToCash = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CardDelCash(objThreeSwapCards As Collection, _
    ByVal lng�����ID As Long, Optional objCard As Card) As Boolean
    'ҽ�ƿ��Ƿ�����
    'Array(�����ID,��������,ȱʡ����,ȱʡ���ַ�ʽ)
    Dim i As Long
    
    If Not objThreeSwapCards Is Nothing Then
        For i = 1 To objThreeSwapCards.Count
            If objThreeSwapCards(i)(0) = lng�����ID Then
                CardDelCash = objThreeSwapCards(i)(1)
                Exit Function
            End If
        Next
    End If
    
    If Not objCard Is Nothing Then
        CardDelCash = objCard.�Ƿ�����
    End If
End Function

Public Function CardDefaultCash(objThreeSwapCards As Collection, _
    ByVal lng�����ID As Long, Optional objCard As Card) As Boolean
    'ҽ�ƿ��Ƿ�ȱʡ����
    'Array(�����ID,��������,ȱʡ����,ȱʡ���ַ�ʽ)
    Dim i As Long
    
    If Not objThreeSwapCards Is Nothing Then
        For i = 1 To objThreeSwapCards.Count
            If objThreeSwapCards(i)(0) = lng�����ID Then
                CardDefaultCash = objThreeSwapCards(i)(2)
                Exit Function
            End If
        Next
    End If
    
    If Not objCard Is Nothing Then
        CardDefaultCash = objCard.�Ƿ�ȱʡ����
    End If
End Function

Public Function CardDefaultBalance(objThreeSwapCards As Collection, ByVal lng�����ID As Long) As String
    'ҽ�ƿ�ȱʡ���ַ�ʽ
    'Array(�����ID,��������,ȱʡ����,ȱʡ���ַ�ʽ)
    Dim i As Long
    
    If Not objThreeSwapCards Is Nothing Then
        For i = 1 To objThreeSwapCards.Count
            If objThreeSwapCards(i)(0) = lng�����ID Then
                CardDefaultBalance = objThreeSwapCards(i)(3)
                Exit Function
            End If
        Next
    End If
End Function

Public Function CheckDelToCash(objThreeSwap As clsThreeSwap, cllThreeSwapCards As Collection, ByVal lng�����ID As Long, _
    ByVal str���㷽ʽ As String, ByVal dblMoney As Double, ByVal lng������� As Long, _
    ByVal str���� As String, ByVal str������ˮ�� As String, ByVal str����˵�� As String) As Boolean
    '���������ּ��
    'Array(�����ID,��������,ȱʡ����,ȱʡ���ַ�ʽ)
    Dim strXMLExpend As String, bln�������� As Boolean
    Dim blnȱʡ����  As Boolean, strȱʡ���ַ�ʽ  As String
    
    On Error GoTo ErrHandler
    If cllThreeSwapCards Is Nothing Then Set cllThreeSwapCards = New Collection
    If CollExitsValue(cllThreeSwapCards, "K" & lng�����ID) Then CheckDelToCash = True: Exit Function
    
    strXMLExpend = _
        "<INPUT>" & vbCrLf & _
        "  <TKLIST>" & vbCrLf & _
        "    <TK>" & vbCrLf & _
        "      <TKFS>" & str���㷽ʽ & "</TKFS>" & vbCrLf & _
        "      <TKJE>" & dblMoney & "</TKJE>" & vbCrLf & _
        "      <JYLSH>" & str������ˮ�� & "</JYLSH>" & vbCrLf & _
        "      <JYSM>" & str����˵�� & "</JYSM>" & vbCrLf & _
        "    </TK>" & vbCrLf & _
        "  </TKLIST>" & vbCrLf & _
        "</INPUT>"
    
    bln�������� = objThreeSwap.CheckDelToCash(Val("7-���ѿ��տ�"), _
        lng�������, lng�����ID, str����, str������ˮ��, str����˵��, dblMoney, _
        strXMLExpend, blnȱʡ����, strȱʡ���ַ�ʽ)
    
    cllThreeSwapCards.Add Array(lng�����ID, bln��������, blnȱʡ����, strȱʡ���ַ�ʽ), "K" & lng�����ID
    
    CheckDelToCash = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function zlGetCorrectCardSql(ByVal strNO As String, ByVal str���� As String, ByVal lng����ID As Long, _
                ByVal lng����ID As Long, ByVal str���㷽ʽ As String, ByVal dbl������ As Double, _
                ByVal str����ʱ�� As String, Optional ByVal lng������ID As Long, Optional ByVal bln���ѿ� As Boolean, _
                Optional ByVal str֧������ As String) As String
    Dim strSQL As String
    
    strSQL = "zl_ҽ�ƿ���¼_����У��("
    '���ݺ�_In         ����Ԥ����¼.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    'ʵ��Ʊ��_In       סԺ���ü�¼.ʵ��Ʊ��%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '����id_In         סԺ���ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '����id_In         סԺ���ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '������Ϣ_In       Varchar2,
    strSQL = strSQL & "'" & IIf(lng������ID > 0, str���㷽ʽ, "") & "',"
    '������_In       ����Ԥ����¼.��Ԥ��%Type,
    strSQL = strSQL & "" & dbl������ & ","
    '����ʱ��_In       סԺ���ü�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "" & str����ʱ�� & ","
    '����Ա���_In     ����Ԥ����¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '����Ա����_In     ����Ԥ����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '�����id_In       ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(lng������ID > 0 And Not bln���ѿ�, lng������ID, "NULL") & ","
    '���㿨���_In     ����Ԥ����¼.���㿨���%Type := Null,
    strSQL = strSQL & "" & IIf(lng������ID > 0 And bln���ѿ�, lng������ID, "NULL") & ","
    '����_In           ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "'" & str֧������ & "',"
    '������ˮ��_In     ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '����˵��_In       ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ")"
    
    zlGetCorrectCardSql = strSQL
End Function



Public Function zlGetUpdateSql(ByVal strNO As String, ByVal lng����ID As Long, _
    Optional ByVal str���㷽ʽ As String, Optional ByVal dbl������ As Double, _
    Optional ByVal int��ɱ�־ As Integer, Optional ByVal intУ�Ա�־ As Integer = 2, _
    Optional ByVal lng�����ID As Long, Optional ByVal bln���ѿ� As Boolean, _
    Optional ByVal str���� As String, Optional ByVal str������ˮ�� As String, _
    Optional ByVal str����˵�� As String, Optional ByVal bln��ͨ���� As Boolean, _
    Optional ByVal str������� As String, Optional ByVal str����ժҪ As String) As String

    Dim strSQL As String

    'Zl_ҽ�ƿ�����_Modify
    strSQL = "Zl_ҽ�ƿ�����_Modify("
    '      ���ݺ�_In     סԺ���ü�¼.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '      ����id_In     סԺ���ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '      ���㷽ʽ_In       ����Ԥ����¼.���㷽ʽ%Type := NULL,
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '      ������_In       ����Ԥ����¼.��Ԥ��%Type := 0,
    strSQL = strSQL & "" & dbl������ & ","
    '      ��ɱ�־_In       Number := 0,
    strSQL = strSQL & "" & int��ɱ�־ & ","
    '      �����ID_In       ����Ԥ����¼.�����ID%Type := Null,
    strSQL = strSQL & "" & ZVal(lng�����ID) & ","
    '      ���ѿ�_In         Number := 0,
    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
    '      ����_In           ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "'" & str���� & "',"
    '      ������ˮ��_In     ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "'" & str������ˮ�� & "',"
    '      ����˵��_In       ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & str����˵�� & "',"
    '      ��ͨ����_In Number:=0
    strSQL = strSQL & "" & IIf(bln��ͨ����, 1, 0) & ","
    '      �������_In       ����Ԥ����¼.�������%Type := Null,
    strSQL = strSQL & "'" & str������� & "',"
    '      ժҪ_In           ����Ԥ����¼.ժҪ%Type := Null
    strSQL = strSQL & "'" & str����ժҪ & "',"
    '      У�Ա�־_In       ����Ԥ����¼.У�Ա�־%Type := 2
    strSQL = strSQL & "" & intУ�Ա�־ & ")"
    zlGetUpdateSql = strSQL
End Function
Public Function GetThirdUpdateSQL(ByVal lngԤ��ID As Long, ByVal strCardNo As String, ByVal str���㷽ʽ As String, ByVal db��� As Double, ByVal str������� As String, _
                            ByVal strSwapGlideNO As String, ByVal strSwapMemo As String, ByVal strժҪ As String, ByVal intNormal As Integer, cllThird As Collection, Optional ByVal blnRetrun As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������Ϣ
    '����: blnRetrun-�Ƿ��˿Ϊtrueֻ���½��㷽ʽ�����Ƿ���ͨ����
    '����:
    '����:2018-09-28
    '˵��:
    '����:132256
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String

    strSQL = "Zl_����Ԥ����¼_Modify("
    '  Id_In         ����Ԥ����¼.Id%Type,
    strSQL = strSQL & "" & lngԤ��ID & ","
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  ������_In   ����Ԥ����¼.���%Type,
    strSQL = strSQL & "" & IIf(blnRetrun, "Null", db���) & ","
    '  �������_In   ����Ԥ����¼.�������%Type,
    strSQL = strSQL & IIf(blnRetrun, "Null", "'" & str������� & "'") & ","
    '  ����_In       ����Ԥ����¼.����%Type,
    strSQL = strSQL & IIf(blnRetrun, "Null", "'" & strCardNo & "'") & ","
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & IIf(blnRetrun, "Null", "'" & strSwapGlideNO & "'") & ","
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & IIf(blnRetrun, "Null", "'" & strSwapMemo & "'") & ","
    '  ����ժҪ_In   ����Ԥ����¼.ժҪ%Type,
    strSQL = strSQL & IIf(blnRetrun, "Null", "'" & strժҪ & "'") & ","
    '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ��ͨ����_In Number:=0
    strSQL = strSQL & "" & intNormal & ")"

    zlAddArray cllThird, strSQL
    GetThirdUpdateSQL = True

End Function


Public Function zlAddUpdateSwapSQL(ByVal blnԤ�� As Boolean, ByVal strIDs As String, _
    ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByRef str���� As String, ByVal str������ˮ�� As String, ByVal str����˵�� As String, _
    ByRef cllPro As Collection, Optional ByVal intУ�Ա�־ As Integer = 0, _
    Optional ByVal int���ͱ�־ As Integer = 0, Optional ByVal bln���ѿ����� As Boolean, _
    Optional ByVal lng��������ID As Long, Optional ByVal strExpend As String, _
    Optional dbl��� As Double, Optional ByVal str���ݺ� As String, _
    Optional ByVal bln�˷� As Boolean, Optional strErrMsg As String, Optional ByVal dbl�ܽ�� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������ˮ�ź���ˮ˵��
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    '����:cllPro-����SQL��
    '����:���˺�
    '����:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strValue As String, strNont As String
    Dim str���㷽ʽ As String, strNO As String
    Dim str������� As String, str����ժҪ As String, str֧������ As String
    Dim dbl������ As Double, dblTotalMoney As Double
    Dim i As Long, lngRow As Long, blnNotFisrt As Boolean, bln��ͨ���� As Boolean
    On Error GoTo errH
    
    If Not bln���ѿ� Then
        If zlXML_ExistNode(strExpend, "OUTPUT") Then
            If dbl�ܽ�� = 0 Then dbl�ܽ�� = dbl���
            If bln�˷� Then
                strNont = "TK"
                Call zlXML_GetChildRows("TKLIST", "TK", lngRow)
            Else
                strNont = "JY"
                Call zlXML_GetChildRows("JYLIST", "JY", lngRow)
            End If
            For i = 0 To lngRow - 1
                '||���׷�ʽ,���׽��,������ˮ��,����˵��,���ݺ�,��ͨ����||...
                If blnNotFisrt Then
                    strErrMsg = "����Ŀǰ��֧��ʹ�ö���֧����ʽ���㣬�������������ϵ�˲�ʹ����ⲿ�����ݡ�"
                    Exit Function
                End If
                If bln�˷� Then
                    Call zlXML_GetChildNodeValue(strNont, "TKFS", i, 0, strValue)
                    str���㷽ʽ = strValue: strValue = ""
                    Call zlXML_GetChildNodeValue(strNont, "TKJE", i, 0, strValue)
                    dbl������ = Val(strValue): strValue = ""
                    dblTotalMoney = dblTotalMoney + dbl������
                Else
                    Call zlXML_GetChildNodeValue(strNont, "JYFS", i, 0, strValue)
                    str���㷽ʽ = strValue: strValue = ""
                    Call zlXML_GetChildNodeValue(strNont, "JYJE", i, 0, strValue)
                    dbl������ = Val(strValue): strValue = ""
                    dblTotalMoney = dblTotalMoney + dbl������
                    Call zlXML_GetChildNodeValue(strNont, "JSHM", i, 0, strValue)
                    str������� = strValue: strValue = ""
                    Call zlXML_GetChildNodeValue(strNont, "JSZY", i, 0, strValue)
                    str����ժҪ = strValue: strValue = ""
                    Call zlXML_GetChildNodeValue(strNont, "KH", i, 0, strValue)
                    str֧������ = IIf(strValue <> "", strValue, str����)
                End If
                Call zlXML_GetNodeValue("JYLSH", i, strValue)
                str������ˮ�� = strValue
                Call zlXML_GetNodeValue("JYSM", i, strValue)
                str����˵�� = strValue
                
                Call zlXML_GetNodeValue("DJH", i, strValue)
                strNO = strValue
                Call zlXML_GetNodeValue("SFPTJS", i, strValue)
                bln��ͨ���� = Val(strValue) = 1
            
                If blnԤ�� Then
                    Call GetThirdUpdateSQL(Val(strIDs), str֧������, str���㷽ʽ, dbl���, str�������, str������ˮ��, str����˵��, str����ժҪ, IIf(bln��ͨ����, 1, 0), cllPro)
                Else
                    strSQL = zlGetUpdateSql(str���ݺ�, Val(strIDs), str���㷽ʽ, dbl���, , , lng�����ID, , str֧������, str������ˮ��, str����˵��, bln��ͨ����, str�������, str����ժҪ)
                    zlAddArray cllPro, strSQL
                End If
                blnNotFisrt = True
            Next
            str���� = str֧������
            If RoundEx(dblTotalMoney, 6) <> RoundEx(dbl�ܽ��, 6) And dbl�ܽ�� <> 0 Then
                strErrMsg = "������:" & dbl�ܽ�� & "Ԫ����ʵ��֧���Ľ��:" & dblTotalMoney & "Ԫ��һ�¡�" & vbCrLf & _
                            "�������������ϵ�˲�ʹ����ⲿ������!"
                Exit Function
            End If
            zlAddUpdateSwapSQL = True
            Exit Function
        End If
    End If
    
    strSQL = "Zl_�����ӿڸ���_Update("
    '  �����id_In   ����Ԥ����¼.�����id%Type,
    strSQL = strSQL & "" & lng�����ID & ","
    '  ���ѿ�_In     Number,
    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
    '  ����_In       ����Ԥ����¼.����%Type,
    strSQL = strSQL & "'" & str���� & "',"
    '  ����ids_In    Varchar2,
    strSQL = strSQL & "'" & strIDs & "',"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    strSQL = strSQL & "'" & str������ˮ�� & "',"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type
    strSQL = strSQL & "'" & str����˵�� & "',"
    '  Ԥ����ɿ�_In Number := 0,--1-����Ԥ����ɿ�;0-�������ѿۿ�
    strSQL = strSQL & "" & IIf(blnԤ��, 1, 0) & ","
    '  �˷ѱ�־_In   Number := 0,--1-�����˷Ѵ���;0-֧������
    strSQL = strSQL & "0,"
    '  У�Ա�־_In   Number := Null,
    strSQL = strSQL & "" & intУ�Ա�־ & ","
    '  ���ͱ�־_In   Number := 0,
    strSQL = strSQL & "" & int���ͱ�־ & ","
    '  ���ѿ�����_In Number := 0 --1-���ѿ�������ã���ʱ ���ѿ�_IN �϶�Ϊ0
    strSQL = strSQL & "" & IIf(bln���ѿ�����, 1, 0) & ")"
    zlAddArray cllPro, strSQL
    zlAddUpdateSwapSQL = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlAddThreeSwapSQLToCollection(ByVal blnԤ���� As Boolean, _
    ByVal strIDs As String, ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByVal str���� As String, strExpend As String, ByRef cllPro As Collection, _
    Optional ByVal lngԤ��ID As Long, Optional ByVal int���� As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������
    '���: blnԤ����-�Ƿ�Ԥ����
    '       lngID-�����Ԥ����,����Ԥ��ID,�������ID
    ' ����:cllPro-����SQL��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, varData As Variant, varTemp As Variant, i As Long, lngRow As Long
    Dim str������Ϣ As String, strTemp As String, strValue As String
     
    Err = 0: On Error GoTo Errhand:
    '���ύ,�����������,�ٸ�����صĽ�����Ϣ
    'strExpend:������չ��Ϣ,��ʽ:��Ŀ����|��Ŀ����||...
    If zlXML_Init("OUTPUT") Then
        If zlXML_LoadXMLToDOMDocument(strExpend, False, True) Then
            Call zlXML_GetChildRows("Expends", "Expend", lngRow)
            For i = 0 To lngRow - 1
                Call zlXML_GetNodeValue("XMMC", i, strValue)
                strTemp = strTemp & "||" & strValue
                Call zlXML_GetNodeValue("XMNR", i, strValue)
                strTemp = strTemp & "|" & strValue
            Next
            If strTemp <> "" Then strTemp = Mid(strTemp, 3)
        Else
            strTemp = strExpend
        End If
    Else
        strTemp = strExpend
    End If
    
    varData = Split(strTemp, "||")

    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If zlCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
                    str������Ϣ = Mid(str������Ϣ, 3)
                    'Zl_�������㽻��_Insert
                    strSQL = "Zl_�������㽻��_Insert("
                    '�����id_In ����Ԥ����¼.�����id%Type,
                    strSQL = strSQL & "" & lng�����ID & ","
                    '���ѿ�_In   Number,
                    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
                    '����_In     ����Ԥ����¼.����%Type,
                    strSQL = strSQL & "'" & str���� & "',"
                    '����ids_In  Varchar2,
                    strSQL = strSQL & "'" & strIDs & "',"
                    '������Ϣ_In Varchar2:������Ŀ|��������||...
                    strSQL = strSQL & "'" & str������Ϣ & "',"
                    'Ԥ����ɿ�_In Number := 0
                    strSQL = strSQL & IIf(blnԤ����, "1", "0") & ","
                    '���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
                    strSQL = strSQL & "NULL" & ","
                    'Ԥ��id_In     ����Ԥ����¼.Id%Type := Null,
                    strSQL = strSQL & IIf(lngԤ��ID = 0, "NULL", lngԤ��ID) & ","
                    '����_In       �������㽻��.����%Type := Null
                    strSQL = strSQL & int���� & ")"
                    zlAddArray cllPro, strSQL
                    str������Ϣ = ""
                End If
                str������Ϣ = str������Ϣ & "||" & strTemp
            End If
        End If
    Next
    If str������Ϣ <> "" Then
        str������Ϣ = Mid(str������Ϣ, 3)
        'Zl_�������㽻��_Insert
        strSQL = "Zl_�������㽻��_Insert("
        '�����id_In ����Ԥ����¼.�����id%Type,
        strSQL = strSQL & "" & lng�����ID & ","
        '���ѿ�_In   Number,
        strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
        '����_In     ����Ԥ����¼.����%Type,
        strSQL = strSQL & "'" & str���� & "',"
        '����ids_In  Varchar2,
        strSQL = strSQL & "'" & strIDs & "',"
        '������Ϣ_In Varchar2:������Ŀ|��������||...
        strSQL = strSQL & "'" & str������Ϣ & "',"
        'Ԥ����ɿ�_In Number := 0
        strSQL = strSQL & IIf(blnԤ����, "1", "0") & ","
        '���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
        strSQL = strSQL & "NULL" & ","
        'Ԥ��id_In     ����Ԥ����¼.Id%Type := Null,
        strSQL = strSQL & IIf(lngԤ��ID = 0, "NULL", lngԤ��ID) & ","
        '����_In       �������㽻��.����%Type := Null
        strSQL = strSQL & int���� & ")"
        zlAddArray cllPro, strSQL
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

