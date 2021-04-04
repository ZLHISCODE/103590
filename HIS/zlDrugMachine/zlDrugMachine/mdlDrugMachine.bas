Attribute VB_Name = "mdlDrugMachine"
Option Explicit

'------------------------------------------------------------------------------
'˵����ҩƷ�Զ����豸�ӿ�ģ��
'���ƣ�������
'------------------------------------------------------------------------------

Public Sub ReadParams(ByRef typVar As TYPE_PARAMS)
'���ܣ���ȡ�����������浽����
    
    Dim objXML As New clsXML
    Dim strFile As String

    '����
    If LCase(App.Path) Like "*\apply" Then
        strFile = App.Path & "\" & GSTR_CONFIG_FILE
    ElseIf LCase(App.Path) Like "*\apply\*" Then
        strFile = Left(App.Path, InStr(LCase(App.Path), "\apply\") + Len("\apply\") - 1) & GSTR_CONFIG_FILE
    ElseIf LCase(App.Path) Like "*zldrugmachinemanage*" Or LCase(App.Path) Like "*zldrugmachine\*" Or LCase(App.Path) Like "*zldrugmachine" Then
        strFile = Replace(App.Path, "\" & App.EXEName, "") & "\" & App.EXEName & "\zlDrugMachineManage\zlDrugMachine.cfg"
    Else
        Exit Sub
    End If
    
    If objXML.OpenXMLFile(strFile) = False Then
        With typVar
            .�����־ = True
            .��ϸ��־ = False
            .������־���� = 7
        End With
        Exit Sub
    End If

    With typVar
        .�����־ = Val(GetParameter(objXML, "output", "0")) = 1
        .��ϸ��־ = Val(GetParameter(objXML, "detailed", "0")) = 1
        .������־���� = Val(GetParameter(objXML, "savedays", "7"))
    End With
    
    objXML.CloseXMLDocument
    Set objXML = Nothing
End Sub

Public Function VerifyConfigFile(ByVal strFile As String) As Boolean
'���ܣ���������ĵ��Ƿ���ڣ������ھ��Զ�����
'������
'���أ�True���ɹ���False���ʧ��

    Dim fsoFile As New FileSystemObject
    Dim tsmFile As TextStream
    
    On Error GoTo hErr
    
    If fsoFile.FileExists(strFile) = False Then
        '���������ĵ�
        Set tsmFile = fsoFile.CreateTextFile(strFile)
        
        'Ĭ�������ĵ�����
        With tsmFile
            .WriteLine "<root>"
            .WriteLine "    <log>"
            .WriteLine "        <output>0</output>"
            .WriteLine "        <detailed>0</detailed>"
            .WriteLine "        <savedays>7</savedays>"
            .WriteLine "    </log>"
            .WriteLine "    <timer>"
            .WriteLine "        <enabled>0</enabled>"
            .WriteLine "        <businessdata></businessdata>"
            .WriteLine "        <cycle>5</cycle>"
            .WriteLine "        <validdays>2</validdays>"
            .WriteLine "        <viewlines>200</viewlines>"
            .WriteLine "    </timer>"
            .WriteLine "</root>"
        End With
        tsmFile.Close
    End If
    
    VerifyConfigFile = True
    Exit Function
    
hErr:
End Function

Private Function GetParameter(ByVal objXML As clsXML, ByVal strName As String, Optional ByVal strDefaultVal As String) As String
'���ܣ���zlDrugMachine.cfg�ļ��л�ȡָ��������ֵ
'������
'  objXML��cfg�ļ������ݼ��غ��XML����
'  strName���������ƣ�����XML�������
'���أ�����ֵ

    Dim strValue As String

    If objXML Is Nothing Then
        GetParameter = strDefaultVal
        Exit Function
    End If
    
    strName = LCase(strName)
    
    If objXML.GetSingleNodeValue(strName, strValue) Then
        GetParameter = strValue
    Else
        GetParameter = strDefaultVal
    End If

End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun���ڸú���
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub CreateSOAP(ByRef objSOAP As Object, ByVal objBase As Object)
'����SOAP����
        
    On Error Resume Next
    
    Set objSOAP = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        Err.Clear
        objBase.mobjLog.Add "������SoapClient30������ʧ�ܣ�", 1
        
        Set objSOAP = CreateObject("MSSOAP.SoapClient")
        If Err.Number <> 0 Then
            Err.Clear
            objBase.mobjLog.Add "�����Դ�����SoapClient20������ʧ�ܣ�", 1
        Else
            objBase.mobjLog.Add "������SoapClient20��������ɣ�", 1
        End If
    Else
        objBase.mobjLog.Add "������SoapClient30��������ɣ�", 1
    End If
    
    On Error GoTo 0
End Sub

Public Sub CreateHTTP(ByRef objHTTP As Object, ByVal objBase As Object)
    On Error Resume Next
    Set objHTTP = Nothing
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number <> 0 Then
        Err.Clear
        objBase.mobjLog.Add "������WinHttp������ʧ�ܣ�����ϵ������Ա", 1
    End If
    On Error GoTo 0
End Sub

Public Function TransmitFlag(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, ByVal rsData As ADODB.Recordset, _
    ByVal objPub As Object, ByVal blnFinish As Boolean) As Boolean

'���ܣ����ͺ��ZLHIS��������־
'������
'  intType��ҵ������
'  intIO��������סԺ
'  rsData����¼������
'  objPub����������
'  blnFinish��True��ɱ�־��Falseʧ�ܱ�־
'���أ�True�ɹ���Faseʧ��
    
    Dim strSQL As String, strInfo As String
    Dim lngStockID As Long
    Dim objDB As Object
    
    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error GoTo hErr
    
    If intAppType = Val("3-֧����") Then
        Set objDB = objPub.mobjComLib
    Else
        Set objDB = objPub.mobjComLib.zlDatabase
    End If
    
    With rsData
        .MoveFirst
        Do
            If intIO = 1 Then
                strInfo = strInfo & ";" & !���� & "," & !������
            Else
                strInfo = strInfo & ";" & !�շ�id
            End If
            lngStockID = !�ⷿid
            
            .MoveNext
            
            If .EOF = False Then
                If lngStockID <> !�ⷿid Then
                    GoTo makProc
                End If
            Else
makProc:
                If Left(strInfo, 1) = ";" Then strInfo = Mid(strInfo, 2)
                If intIO = 1 Then
                    strSQL = "ZL_ҩƷ�շ������־_FLAG(" & _
                        IIf(intType >= 20, intType - 20, intType) & "," & _
                        lngStockID & ",'" & strInfo & "'," & IIf(blnFinish, 1, 0) & ")"
                    objPub.mobjLog.Add strSQL, 2, 1
                    Call objDB.ExecuteProcedure(strSQL, "ҩƷ�շ������־")
                Else
                    strSQL = "ZL_ҩƷ�շ�סԺ��־_FLAG(" & _
                        IIf(intType >= 20, intType - 20, intType) & "," & _
                        "'" & strInfo & "'," & IIf(blnFinish, 1, 0) & ")"
                    objPub.mobjLog.Add strSQL, 2, 1
                    Call objDB.ExecuteProcedure(strSQL, "ҩƷ�շ�סԺ��־")
                    strInfo = ""
                End If
            End If
        Loop While .EOF = False
    End With
    
    objPub.mobjLog.Save
    TransmitFlag = True
    
    Exit Function
    
hErr:
    objPub.mobjLog.Add Err.Number & ":" & Err.Description, 2
    objPub.mobjLog.Save
End Function

Private Function GetWinName(ByVal lngDeptID As Long, ByVal lngWinCode As Long, ByVal objDB As Object, ByVal objLog As Object) As String
'���ܣ������ڱ���ת�ɴ�������
'������
'  lngDeptID���ⷿID
'  lngWinCode�����ڱ���
'���أ���������
    
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    
    strSQL = "Select ���� From ��ҩ���� Where ҩ��id = [1] And ���� = [2] "
    Set rsTemp = objDB.OpenSQLRecord(strSQL, "�����ڱ���ת�ɴ�������", lngDeptID, CStr(lngWinCode))
    If rsTemp.EOF = False Then
        GetWinName = NVL(rsTemp!����)
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    objLog.Add "�����ڱ���ת�ɴ�������ʧ��", 2
    objLog.Add Err.Number & ":" & Err.Description, 2
    objLog.Save
End Function

Public Function UpdateDispenseWindow(ByVal rsData As ADODB.Recordset, ByVal strWin As String, ByVal objDB As Object, ByVal objLog As Object) As Boolean
'���ܣ��������ݿ�Ĵ�����Ϣ
'������
'  rsData������Դ
'  strWin������
'���أ�True�ɹ���Falseʧ��

    Dim lngStockID As Long
    Dim strNO As String, strSQL As String
    Dim intBill As Integer
    Dim strWinName As String

    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error GoTo hErr
    
    With rsData
        .MoveFirst
        Do
            'ͬ�ⷿ��ͬ���ݡ�ͬ������ֻ�ܸ���һ������
            lngStockID = !�ⷿid
            strNO = Trim(!������)
            intBill = !����
            
            '���ڱ���ת��������
            strWinName = GetWinName(lngStockID, Val(strWin), objDB, objLog)
            
            .MoveNext
            
            If .EOF = False Then
                If Not (lngStockID = !�ⷿid And strNO = Trim(!������) And intBill = !����) Then
                    GoTo makProc
                End If
            Else
makProc:
                strSQL = "Zl_δ��ҩƷ��¼_���䷢ҩ����(" & _
                         "'" & strNO & "'," & _
                         intBill & "," & _
                         lngStockID & "," & _
                         IIf(Trim(strWinName) = "", "Null", "'" & strWinName & "'") & ")"
                objLog.Add strSQL, 2, 1
                Call objDB.ExecuteProcedure(strSQL, "���·�ҩ����")
            End If
        Loop While .EOF = False
    End With
    
    objLog.Save
    UpdateDispenseWindow = True
    Exit Function
      
hErr:
    objLog.Add Err.Number & ":" & Err.Description, 2
    objLog.Save
End Function

Public Function CopyStructure(ByVal fdsSource As ADODB.Fields) As ADODB.Recordset
'���ܣ�
'������
'���أ�

    Dim i As Integer

    On Error GoTo hErr

    Set CopyStructure = New ADODB.Recordset
    
    With CopyStructure
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        
        '�ṹ����
        For i = 0 To fdsSource.Count - 1
            .Fields.Append fdsSource(i).Name, IIf(fdsSource(i).Type = adNumeric, adDouble, fdsSource(i).Type), fdsSource(i).DefinedSize, adFldIsNullable
        Next
        
        .Open
    End With
    
    Exit Function

hErr:
    Set CopyStructure = Nothing
End Function

Public Function CopyRecord(ByVal fdsSource As ADODB.Fields, ByRef rsTarget As ADODB.Recordset) As String
'���ܣ�
'������
'���أ�

    Dim i As Integer
    
    On Error GoTo hErr
    
    rsTarget.AddNew
    For i = 0 To fdsSource.Count - 1
        rsTarget.Fields(i).Value = fdsSource(i).Value
    Next
    rsTarget.Update
    
    Exit Function
    
hErr:
    CopyRecord = Err.Number & ":" & Err.Description
End Function

Public Sub ClearRecord(ByRef rsSource As ADODB.Recordset)
    With rsSource
        .MoveLast
        Do While .BOF = False
            .Delete
            .MovePrevious
        Loop
    End With
End Sub

Public Function IP(Optional ByRef strErr As String) As String
    '���ܣ�ͨ��API��ȡ��ʱIP
    
    Dim ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    Dim strTmpErr As String, strALLErr As String
    
    strErr = ""
    On Error GoTo Errhand
    GetIpAddrTable ByVal 0&, ret, True
    If ret <= 0 Then Exit Function
    ReDim bBytes(0 To ret - 1) As Byte
    ReDim TempList(0 To ret - 1) As String
    'retrieve the data
    GetIpAddrTable bBytes(0), ret, False
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    For Tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr, strTmpErr)
        If strTmpErr <> "" Then strALLErr = strALLErr & IIf(strALLErr = "", "", "|") & strTmpErr
    Next Tel
    'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        IP = TempIP 'Return The TempIP
    Exit Function
    strErr = strALLErr
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strALLErr & IIf(strALLErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Private Function ConvertAddressToString(longAddr As Long, Optional ByRef strErr As String) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    
    strErr = ""
    On Error GoTo errH
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errH:
    strErr = Err.Description
    Err.Clear
End Function

Public Function GetUserInfo(ByVal strDBUser As String, ByVal objComLib As Object, ByVal objLog As Object, ByRef typUserInfo As TYPE_USER_INFO) As Boolean
'���ܣ���ȡ��ǰ�û��Ļ�����Ϣ
'���أ�����Ado��¼��
    Dim strSQL As String, strDefault As String
    Dim rsTemp As ADODB.Recordset
    
    If objComLib Is Nothing Then
        objLog.Add "GetUserInfo��objComLib����ΪNothing������ֹ����ִ��", 1
        objLog.Save
        Exit Function
    End If
    
    On Error GoTo hErr
    
    objLog.Add "��ȡ�û���Ϣ", 1
    objLog.Add "�û�����" & UCase(strDBUser), 2
    
    strDefault = " And C.ȱʡ = 1"
    strSQL = "Select User,A.Id, A.���, A.����, A.����, A.רҵ����ְ��,B.�û���, C.����id, D.���� As ������, D.���� As ������ " & vbNewLine & _
             "From ��Ա�� A, �ϻ���Ա�� B, ������Ա C, ���ű� D " & vbNewLine & _
             "Where A.Id = B.��Աid And A.Id = C.��Աid And C.����id = D.Id And B.�û��� = [1] "
    If TypeName(objComLib) = "clsPublic" Then
        Set rsTemp = objComLib.OpenSQLRecord(strSQL & strDefault, "GetUserInfo", UCase(strDBUser))
    Else
        Set rsTemp = objComLib.zlDatabase.OpenSQLRecord(strSQL & strDefault, "GetUserInfo", UCase(strDBUser))
    End If
    objLog.Add strSQL & strDefault, 2
    
    If rsTemp.RecordCount = 0 Then
        strDefault = " And Rownum < 2"
        Set rsTemp = objComLib.zlDatabase.OpenSQLRecord(strSQL & strDefault, "GetUserInfo", UCase(strDBUser))
        objLog.Add strSQL & strDefault, 2
    End If
    
    If rsTemp.RecordCount > 0 Then
        typUserInfo.ID = rsTemp!ID
        typUserInfo.��� = rsTemp!���
        typUserInfo.����ID = mdlDrugMachine.NVL(rsTemp!����ID, 0)
        typUserInfo.���� = mdlDrugMachine.NVL(rsTemp!����)
        typUserInfo.���� = mdlDrugMachine.NVL(rsTemp!����)
        typUserInfo.�û��� = rsTemp!�û���
        GetUserInfo = True
        objLog.Add "��ȡ�û���Ϣ�ɹ�", 1
    Else
        typUserInfo.ID = 0
        typUserInfo.��� = ""
        typUserInfo.����ID = 0
        typUserInfo.���� = ""
        typUserInfo.���� = ""
        typUserInfo.�û��� = ""
        objLog.Add "��ȡ�û���Ϣʧ��", 1
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    objLog.Add "��ȡ�û���Ϣʧ��", 1
    objLog.Add Err.Number & ":" & Err.Description, 1
    objLog.Save
End Function

Public Function SpecialChar(ByVal strVal As Variant) As String
'���ܣ������ַ�ת��
'˵����
' < ת &lt;
' > ת &gt;
' & ת &amp;
' ' ת &apos;
' " ת &quot;
    Dim strReturn As String
    
    If IsNull(strVal) Then
        strVal = ""
        GoTo errHandle
    End If
    If strVal = "" Then
        GoTo errHandle
    End If
    On Error GoTo errHandle
    strReturn = strVal
    strReturn = Replace(strReturn, "<", "&lt;")
    strReturn = Replace(strReturn, ">", "&gt;")
    strReturn = Replace(strReturn, "&", "&amp;")
    strReturn = Replace(strReturn, "'", "&apos;")
    strReturn = Replace(strReturn, """", "&quot;")
    SpecialChar = strReturn
    Exit Function
    
errHandle:
    SpecialChar = strVal
End Function

Public Function GetInterfaceLink(ByVal objPublic As Object, ByVal strCode As String) As String
'���ܣ���ȡ�ӿ�������Ϣ
'������
'  objPublic����������
'  strCode���ӿڱ��
'���أ����ܺ�����Ӵ�

    Dim strSQL As String, strLink As String
    Dim rsTemp As ADODB.Recordset
    Dim objEncrypt As Object
    
    On Error Resume Next
    Set objEncrypt = CreateObject("zlEncryptPub.clsEncrypt")
    If Err.Number <> 0 Then
        objPublic.mobjLog.Add "zlEncryptPub����δע�ᣬӰ��ӿ�������Ϣ�Ľ���", 1
    End If
    Err.Clear
    
    On Error GoTo hErr
    
    strSQL = "Select ������Ϣ From ҩƷ�豸�ӿ� Where ��� = [1] And ͣ������ Is Null And �������� Is Not Null "
    If LCase(TypeName(objPublic.mobjComLib)) = "clspublic" Then
        Set rsTemp = objPublic.mobjComLib.OpenSQLRecord(strSQL, "��ȡ�ӿ�������Ϣ", strCode)
    Else
        Set rsTemp = objPublic.mobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ӿ�������Ϣ", strCode)
    End If
    If rsTemp.EOF = False Then
        If Not IsNull(rsTemp!������Ϣ) Then
            strLink = objEncrypt.Base64Decode(rsTemp!������Ϣ)
        End If
    End If
    rsTemp.Close
    
    objPublic.mobjLog.Add "��ȡ�ӿ�������Ϣ���", 1
    objPublic.mobjLog.Save
    
    GetInterfaceLink = strLink
    Exit Function
    
hErr:
    objPublic.mobjLog.Add Err.Number & ":" & Err.Description, 1
    objPublic.mobjLog.Save
End Function

Public Sub ExecuteProcedureBeach(ByVal cllProcs As Variant, ByVal strCaption As String, ByVal cnThird As ADODB.Connection, _
    ByRef objLog As Object, Optional blnTrans As Boolean = True, Optional blnCommit As Boolean = True)
'---------------------------------------------------------------------------------------------
'����:ִ����ص�Oracle���̼�
'����:cllProcs-oracle���̼�������Ϊ���飬Ҳ����Ϊ���ϣ�����Ϊ��������
'     strCaption -ִ�й��̵ĸ����ڱ���
'     blnTrans-�Ƿ��������
'     blnCommit-ִ������̺�,�ύ����(ǰ��:blnTrans=true)
'---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo hErr
    
    If blnTrans Then cnThird.BeginTrans
    
    If TypeName(cllProcs) = "Collection" Then '������ʽ
        For i = 1 To cllProcs.Count
            strSQL = cllProcs(i)
            objLog.Add strSQL, 2, 1
            Call ExecuteProcedure(cnThird, strSQL, strCaption)
        Next
    ElseIf Not IsObject(cllProcs) Then
        If VarType(cllProcs) = vbArray + vbVariant Or VarType(cllProcs) = vbArray + vbString Then  '������ʽ
            For i = LBound(cllProcs) To UBound(cllProcs)
                strSQL = cllProcs(i)
                objLog.Add strSQL, 2, 1
                Call ExecuteProcedure(cnThird, strSQL, strCaption)
            Next
        End If
    End If
    
    If blnCommit And blnTrans Then
        cnThird.CommitTrans
    End If
    objLog.Save
    Exit Sub
    
hErr:
    If blnCommit And blnTrans Then
        cnThird.RollbackTrans
    End If
    objLog.Add strSQL, 2, 1
    objLog.Add Err.Number & ":" & Err.Description, 2
    objLog.Add "ExecuteProcedureBeach()", 2
    objLog.Save
End Sub

Public Sub ExecuteProcedure(ByVal cnThird As ADODB.Connection, ByRef strSQL As String, ByVal strFormCaption As String)
'���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
'������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
'˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
'  1.���������Ǳ��ʽ,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
'  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
'  3.��Ϊ�ù������Զ�����,����һ��ʹ�ð󶨱���,�Դ�"'"���ַ�����,��Ҫʹ��"''"��ʽ��
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        'ִ�еĹ�����
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        'ִ�й��̲���
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '�Ƿ����ַ����ڣ��Լ����ʽ��������
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '����
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '�ַ���
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle���ӷ�����:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '˫"''"�İ󶨱�������
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '���Ӳ�������LOBʱ������ð󶨱���ת��ΪRAWʱ����2000���ַ�Ҫ��adLongVarChar
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '����
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULLֵ�������ִ���ɼ�����������
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '����
                        If datCur = CDate(0) Then datCur = Now()
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULLֵ�����ַ�����ɼ�����������
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '��ѡ��������NULL������ܸı���ȱʡֵ:��˿�ѡ��������д���м�
                        GoTo NoneVarLine
                    Else '�������������ӵı��ʽ���޷�����
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '����Ա���ù���ʱ��д����
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "���� Oracle ����""" & strProc & """ʱ�����Ż�������д��ƥ�䡣ԭʼ������£�" & vbCrLf & vbCrLf & strSQL
            Exit Sub
        End If
        
        '����?��
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        Set cmdData.ActiveConnection = cnThird      '���Ƚ���
        cmdData.CommandType = adCmdText
        cmdData.CommandText = strProc
        
        Call cmdData.Execute
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
    
NoneVarLine:
    '˵����Ϊ�˼��������ӷ�ʽ
    '1.��������adCmdStoredProc��ʽ��8i����������
    '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    cnThird.Execute strSQL, , adCmdText
End Sub

Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub

