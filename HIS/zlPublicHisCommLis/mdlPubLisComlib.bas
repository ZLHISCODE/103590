Attribute VB_Name = "mdlPubLisComlib"
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'ģ�鹦��:�ӿ��й���������
'---------------------------------------------------------------------------------------

Option Explicit

Public Const Sel_Lis_DB As Integer = 1
Public Const Sel_His_DB As Integer = 2

Public intLis_Setup As Integer                                     '�ж�LIS�Ƿ�װ 0=δ��װ  1=�Ѱ�װ
Public intEMR_Setup As Integer                                     '�ж�EMR�Ƿ�װ 0=δ��װ 1=�Ѱ�װ ����δʹ�ã��°���Ӳ�����û��û�б�Ÿ���жϰ�װʱͨ�����������ͳ�ʼ�������Ƿ�ɹ�ȷ���Ƿ��жϰ�װ



Public Function InitDBConn(cnHisOracle As Connection, Optional strErr As String) As Boolean
      '����       ��ʹLIS��HIS�����ݿ�����
      '����       1=lis 2=his 3=tj 4=xk 5=EMR
      '����       ���ӳɹ�������,���Ӳ��ɹ����ؼ�

          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim strConn As String
          Dim strCode As String
          Dim astrItem() As String

1         On Error GoTo InitDBConn_Error

2         If cnHisOracle.State <> 1 Then
3             strErr = "�����HIS����״̬������������!"
4             InitDBConn = False
5             Exit Function
6         End If

          '�����ݿ��ȡ���õ�HIS����
7         strSQL = "Select ����ֵ From zlOptions Where  ������ ='LISϵͳ��������'"
8         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "zlGetSymbol")
9         If rsTmp.RecordCount > 0 Then
10            strConn = rsTmp("����ֵ") & ""
11        End If
          'û���������ӱ�ʾ����װ
12        If strConn = "" Then
13            Set gcnLisOracle = cnHisOracle
14        Else
              'ͨ�����õ�����������HIS��
15            strCode = gobjHisComLib.zlStr.Sm4DecryptEcb(strConn)
16            astrItem = Split(strCode, "<SP 1>")
17            If OraDataOpen(gcnLisOracle, astrItem(2), astrItem(0), IIf(UCase(astrItem(1)) = "SYS" Or UCase(astrItem(1)) = "SYSTEM", _
                                                                         astrItem(1), astrItem(1))) = False Then
18                InitDBConn = False
19                Exit Function
20            End If
21        End If

          '---------------------------�жϸ���ϵͳ�Ƿ�װ-----------------------------------------

          'LIS
22        strSQL = "select count(*) count from zlsystems where ��� = 2500 "
23        Set rsTmp = OpenSQLRecord(Sel_Lis_DB, strSQL, "��ʼ��")
24        If rsTmp("count") > 0 Then
25            intLis_Setup = 1
26        End If
          
          '�ж�LISϵͳ�Ƿ�װ��δ��װ��ֱ���˳�
27        If intLis_Setup <> 1 Then Exit Function
          
          '--------------------��ȡ��װ�汾--------------------------------------
           '��ȡLIS�汾
28        strSQL = "select �汾��  from zlsystems where ��� = 2500 "
29        Set rsTmp = OpenSQLRecord(Sel_Lis_DB, strSQL, "��ʼ��")
30        If rsTmp.RecordCount > 0 Then
31            gSysInfo.VersionLIS = rsTmp("�汾��")
32        End If
          '��ȡHIS�汾
33        strSQL = "select �汾��  from zlsystems where ��� = 100 "
34        Set rsTmp = OpenSQLRecord(Sel_His_DB, strSQL, "��ʼ��")
35        If rsTmp.RecordCount > 0 Then
36            gSysInfo.VersionHIS = rsTmp("�汾��")
37        End If
          
          '�°���Ӳ���
          'δ�������Ϸ�ʽ�жϰ�װԭ��
          '1��������װ,û��zlsystems����ͨ�����Ϸ�ʽ�ж�
          '2���жϰ�װʱֱ��ͨ������EMR�����Ƿ�ɹ�,�ͳ�ʼ�������Ƿ�ɹ����ж��Ƿ�װ����������жϣ�ʹ��ʱ���ж�

          '�°���Ӳ���
38        intEMR_Setup = getERPSetupType
          
          '------------------------------------------------------------------------------------------

39        InitDBConn = True

40        Exit Function
InitDBConn_Error:
41        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(InitDBConn)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
42        Err.Clear
End Function

Private Function getERPSetupType() As Integer
    '����           �ж��°���Ӳ����Ƿ�װ
    '�����°���Ӳ���û�б��,����ֻ�ܳ��Դ�������,��������ܹ�����,��˵���Ѿ���װ,�����ʾû�а�װ
    '����           1=�Ѱ�װ
            
    On Error GoTo getERPSetupType_Error

    If gobjEmrInterface Is Nothing Then
        Set gobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
    End If
    
    getERPSetupType = 1
    
    Exit Function
getERPSetupType_Error:
    getERPSetupType = 0
End Function

Public Function ComGetUserInfo(ByRef strErr As String) As Boolean
      '���ܣ���ȡ��½�û���Ϣ
      '       intType         1=lis 2=his 3=���� 4=Ѫ��
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim objLogin As Object

1         On Error GoTo ComGetUserInfo_Error

2         gstrDBUser = GetUserDB(Sel_His_DB)

3         strSQL = "Select SYS_CONTEXT('USERENV','TERMINAL') as MName From Dual"
4         Set rsTmp = OpenSQLRecord(Sel_His_DB, strSQL, "��ʼ��")
5         gUserInfo.ComputerName = rsTmp("MName")


          '��ȡ��½վ��
6         Set objLogin = CreateObject("ZLLogin.clsLogin")
7         gUserInfo.NodeNo = objLogin.NodeNo
8         If gUserInfo.NodeNo = "" Then gUserInfo.NodeNo = "-"


9         Set rsTmp = GetUserInfo(Sel_His_DB)

10        If Not rsTmp.EOF Then
11            gUserInfo.ID = Val("" & rsTmp!ID)
12            gUserInfo.No = Trim("" & rsTmp!���)
13            gUserInfo.DeptID = Val("" & rsTmp!����ID)
14            gUserInfo.DeptName = Trim("" & rsTmp!������)
15            gUserInfo.Code = Trim("" & rsTmp!����)
16            gUserInfo.Name = Trim("" & rsTmp!����)
17            ComGetUserInfo = True
18        End If


19        Exit Function
ComGetUserInfo_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(ComGetUserInfo)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
21        Err.Clear

End Function
Public Function ComGetSysParameter(strErr As String) As Boolean
          '��ȡϵͳ����

1         On Error GoTo ComGetSysParameter_Error

2         ComGetSysParameter = False
          
3         gSysParameter.BuffDir = App.Path & "\Buffer"
4         gSysParameter.InvaidWord = "`#@$%&|\{}[]?;""'"
5         ComGetSysParameter = True


6         Exit Function
ComGetSysParameter_Error:
7         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(ComGetSysParameter)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
8         Err.Clear

End Function

Public Function OraDataOpen(cnOracle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strError As String
    Dim strSysName As String
    
    strSysName = "�ӿ�����"
    
    '����������,�����zlRegister������µ�ע�᷽ʽ,��������ϵ�ע�᷽ʽ
    On Error GoTo errhandOld
   
    Set cnOracle = FunGetConnection(strServerName, strUserName, strUserPwd, True, , strError, False)
    If strError <> "" Then
        MsgBox strError, vbInformation, strSysName
        Exit Function
    End If
    OraDataOpen = True
    Exit Function
errhandOld:
    
    On Error Resume Next
    Err = 0
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, TranPasswd(strUserPwd)
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, strSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, strSysName
            Else
                MsgBox strError, vbInformation, strSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    
    OraDataOpen = True

End Function


Public Function TranPasswd(strOld As String) As String

    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim iBit As Integer, StrBit As String
    Dim strNew As String
    On Error GoTo TranPasswd_Error

    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        StrBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(StrBit = "0", "W", StrBit = "1", "I", StrBit = "2", "N", StrBit = "3", "T", StrBit = "4", "E", StrBit = "5", "R", StrBit = "6", "P", StrBit = "7", "L", StrBit = "8", "U", StrBit = "9", "M", _
                   StrBit = "A", "H", StrBit = "B", "T", StrBit = "C", "I", StrBit = "D", "O", StrBit = "E", "K", StrBit = "F", "V", StrBit = "G", "A", StrBit = "H", "N", StrBit = "I", "F", StrBit = "J", "J", _
                   StrBit = "K", "B", StrBit = "L", "U", StrBit = "M", "Y", StrBit = "N", "G", StrBit = "O", "P", StrBit = "P", "W", StrBit = "Q", "R", StrBit = "R", "M", StrBit = "S", "E", StrBit = "T", "S", _
                   StrBit = "U", "T", StrBit = "V", "Q", StrBit = "W", "L", StrBit = "X", "Z", StrBit = "Y", "C", StrBit = "Z", "X", True, StrBit)
        Case 2
            strNew = strNew & _
                Switch(StrBit = "0", "7", StrBit = "1", "M", StrBit = "2", "3", StrBit = "3", "A", StrBit = "4", "N", StrBit = "5", "F", StrBit = "6", "O", StrBit = "7", "4", StrBit = "8", "K", StrBit = "9", "Y", _
                   StrBit = "A", "6", StrBit = "B", "J", StrBit = "C", "H", StrBit = "D", "9", StrBit = "E", "G", StrBit = "F", "E", StrBit = "G", "Q", StrBit = "H", "1", StrBit = "I", "T", StrBit = "J", "C", _
                   StrBit = "K", "U", StrBit = "L", "P", StrBit = "M", "B", StrBit = "N", "Z", StrBit = "O", "0", StrBit = "P", "V", StrBit = "Q", "I", StrBit = "R", "W", StrBit = "S", "X", StrBit = "T", "L", _
                   StrBit = "U", "5", StrBit = "V", "R", StrBit = "W", "D", StrBit = "X", "2", StrBit = "Y", "S", StrBit = "Z", "8", True, StrBit)
        Case 0
            strNew = strNew & _
                Switch(StrBit = "0", "6", StrBit = "1", "J", StrBit = "2", "H", StrBit = "3", "9", StrBit = "4", "G", StrBit = "5", "E", StrBit = "6", "Q", StrBit = "7", "1", StrBit = "8", "X", StrBit = "9", "L", _
                   StrBit = "A", "S", StrBit = "B", "8", StrBit = "C", "5", StrBit = "D", "R", StrBit = "E", "7", StrBit = "F", "M", StrBit = "G", "3", StrBit = "H", "A", StrBit = "I", "N", StrBit = "J", "F", _
                   StrBit = "K", "O", StrBit = "L", "4", StrBit = "M", "K", StrBit = "N", "Y", StrBit = "O", "D", StrBit = "P", "2", StrBit = "Q", "T", StrBit = "R", "C", StrBit = "S", "U", StrBit = "T", "P", _
                   StrBit = "U", "B", StrBit = "V", "Z", StrBit = "W", "0", StrBit = "X", "V", StrBit = "Y", "I", StrBit = "Z", "W", True, StrBit)
        End Select
    Next
    TranPasswd = strNew


    Exit Function
TranPasswd_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlPubLisComlib", "ִ��(TranPasswd)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
    Err.Clear

End Function

Public Sub ExecuteProcedure(ByVal selDB As Integer, strSQL As String, ByVal strFormCaption As String)
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
        '���ԭ�в���:��Ȼ�����ظ�ִ��
        '        cmdData.CommandText = "" '��Ϊ����ʱ�����������
        '        Do While cmdData.Parameters.Count > 0
        '            cmdData.Parameters.Delete 0
        '        Loop

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
                    If IsNumeric(strPar) Then    '����
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then    '�ַ���
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
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then    '����
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
                    ElseIf UCase(strPar) = "SYSDATE" Then    '����
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then    'NULLֵ�����ַ�����ɼ�����������
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then    '��ѡ��������NULL������ܸı���ȱʡֵ:��˿�ѡ��������д���м�
                        GoTo NoneVarLine
                    Else    '�������������ӵı��ʽ���޷�����
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

        If selDB = Sel_Lis_DB Then
            Set cmdData.ActiveConnection = gcnLisOracle    '���Ƚ���(���ִ��1000��Լ0.5x��)
        ElseIf selDB = Sel_His_DB Then
            Set cmdData.ActiveConnection = gcnHisOracle    '���Ƚ���(���ִ��1000��Լ0.5x��)
        End If
        
        Call ExportLog(selDB, False, "ExecuteProcedure", strFormCaption, strSQL)
        'ִ�й���
        'If cmdData.ActiveConnection Is Nothing Then
        '            Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
        cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc

        '        Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
        '        Call gobjComLib.SQLTest
        Call ExportLog(selDB, True, "ExecuteProcedure", strFormCaption, "")
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
    '    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)

    '˵����Ϊ�˼��������ӷ�ʽ
    '1.��������adCmdStoredProc��ʽ��8i����������
    '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"

    If selDB = Sel_Lis_DB Then
        gcnLisOracle.Execute strSQL, , adCmdText
    ElseIf selDB = Sel_His_DB Then
        gcnHisOracle.Execute strSQL, , adCmdText
    End If


    '    Call gobjComLib.SQLTest

End Sub

'Ϊ���������ӽӿں����������벿���ļ�����Ӱ�죬ר���϶���ΪPrivate����
Private Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    TrimEx = strText
End Function


Public Function OpenSQLRecord(ByVal selDB As Integer, ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'               ��Ϊʹ�ð󶨱���,�Դ�"'"���ַ�����,����Ҫʹ��"''"��ʽ��
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLtmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    
    '������ʹ���˶�̬�ڴ������û��ʹ��/*+ XXX*/����ʾ��ʱ�Զ�����

    strSQLtmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLtmp, 7)), 1, 2) <> "/*" And Mid(strSQLtmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLtmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '�ж�ǰ���Ƿ�����IN �����򲻼�Rule
                '���ҵ����һ��SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  'ȡ����3���ַ�
                
                If strTmp = "IN(" Then '����in(select��������������ѭ�������Ƿ����û��ʹ������д����������̬�ڴ溯��
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
        End If
    End If
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
'    cmdData.CommandText = "" '��Ϊ����ʱ�����������
'    Do While cmdData.Parameters.Count > 0
'        cmdData.Parameters.Delete 0
'    Loop
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next
    Call ExportLog(selDB, False, "OpenSQLRecord", strTitle, strSQL, arrInput)
    'ִ�з��ؼ�¼��
    'If cmdData.ActiveConnection Is Nothing Then
    If selDB = Sel_Lis_DB Then
        Set cmdData.ActiveConnection = gcnLisOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
    ElseIf selDB = Sel_His_DB Then
        Set cmdData.ActiveConnection = gcnHisOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
    End If
    
'     Set cmdData.ActiveConnection = gcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
    'End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
    Call ExportLog(selDB, True, "OpenSQLRecord", strTitle, "")
'    Call gobjComLib.SQLTest

End Function


Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnLisOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).value
    rsTemp.Close
    Exit Function
errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
    Currentdate = 0
    Err = 0
End Function

Public Function SetPara(ByVal selDB As Integer, ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
      '���ܣ�����ָ���Ĳ���ֵ
      '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
      '      strValue=Ҫ���õĲ���ֵ
      '      lngSys=ʹ�øò�����ϵͳ��ţ���100
      '      lngModual=ʹ�øò�����ģ��ţ���1230
      '      blnSetup=����ģ���Ƿ��в�������Ȩ��
      '���أ������Ƿ�ɹ�
          Dim strSQL As String
          Dim strResFilter As String
          '������ֵ�����û�б仯�򲻴���
1         On Error GoTo SetPara_Error

2         strSQL = GetPara(selDB, varPara, lngSys, lngModual)
3         If strSQL = strValue Then SetPara = True: Exit Function
          
4         SetPara = True
5         strSQL = "zl_Parameters_Update('" & varPara & "','" & strValue & "'," & lngSys & "," & lngModual & "," & IIf(blnSetup, 1, 0) & ")"
6         Call ExecuteProcedure(selDB, strSQL, "SetPara")
          
          '���»����¼�����߼���zl_Parameters_Update����һ��
          '��������
7         If TypeName(varPara) = "String" Then
8             strResFilter = "������='" & CStr(varPara) & "' And ģ��=" & lngModual & " And ϵͳ=" & lngSys
9         Else
10            strResFilter = "������=" & Val(varPara) & " And ģ��=" & lngModual & " And ϵͳ=" & lngSys
11        End If
          
12        grsParas.Filter = strResFilter
13        If grsParas.EOF Then Exit Function
          'Ȩ���ж�
14        If Not blnSetup Then
              '����ȫ�ֲ���,�̶���ҪȨ��
15            If grsParas!ϵͳ <> 0 And grsParas!ģ�� = 0 And grsParas!˽�� = 0 And grsParas!���� = 0 Then
16                Exit Function
              '����ģ�����,�̶���ҪȨ��
17            ElseIf grsParas!ģ�� = 0 And grsParas!˽�� = 0 And grsParas!���� = 0 Then
18                Exit Function
              'Ҫ��Ȩ���Ƶı�������ģ��
19            ElseIf grsParas!ϵͳ <> 0 And grsParas!ģ�� <> 0 And grsParas!˽�� = 0 And grsParas!���� = 1 And grsParas!��Ȩ = 1 Then
20                Exit Function
21            End If
22        End If
          
23        If grsParas!˽�� = 1 Or grsParas!���� = 1 Then
24            grsUserParas.Filter = "����ID=" & grsParas!ID & _
                          IIf(grsParas!˽�� = 1, " And �û���='" & grsParas!�û��� & "'", " And �û���='NullUser'") & _
                          IIf(grsParas!���� = 1, " And ������='" & grsParas!������ & "'", " And ������='NullMachine'")
              
25            If grsUserParas.EOF Then
26                grsUserParas.AddNew
27                grsUserParas!����id = grsParas!ID
28                grsUserParas!�û��� = IIf(grsParas!˽�� = 1, grsParas!�û���, "NullUser")
29                grsUserParas!������ = IIf(grsParas!���� = 1, grsParas!������, "NullMachine")
30                grsUserParas!����ֵ = strValue
31                grsUserParas.Update
32            Else
33                grsUserParas!����ֵ = strValue
34                grsUserParas.Update
35            End If
36        Else
37            grsParas!����ֵ = strValue
38            grsParas.Update
39        End If
          
40        Exit Function
SetPara_Error:
41        SetPara = False
42        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(SetPara)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
43        Err.Clear
End Function


Public Function GetPara(ByVal selDB As Integer, ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer) As String
      '���ܣ���ȡָ���Ĳ���ֵ
      '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
      '      lngSys=ʹ�øò�����ϵͳ��ţ���100
      '      lngModual=ʹ�øò�����ģ��ţ���1230
      '      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
      '      blnNotCache=�Ƿ񲻴ӻ����ж�ȡ
      '      arrControl=�ؼ����飬��Array(Me.Text1, Me.CheckBox1)�����ں����ڲ��Զ������Ӧ�ؼ�����ʾ��ɫ���Ƿ��ֹ���á�
      '      blnSetup=����ģ���Ƿ��в�������Ȩ��
      '      intType=���ز��������ز�������
      '���أ�����ֵ���ַ�����ʽ
          Dim strSQL As String, i As Integer
          Dim blnNew As Boolean, blnEnabled As Boolean, blnNewRow As Boolean, blnNotExists As Boolean
          Dim strSqlFilter As String, strResFilter As String
          Dim rsTmp As ADODB.Recordset
          Dim strDBUser As String

1         On Error GoTo GetPara_Error

2         strDBUser = GetUserDB(selDB)
3         intType = 0
          
          '��������
4         If TypeName(varPara) = "String" Then
5             strResFilter = "������='" & CStr(varPara) & "' And ģ��=" & lngModual & " And ϵͳ=" & lngSys
6             strSqlFilter = "������=[5] And Nvl(ģ��,0)=[3] And Nvl(ϵͳ,0)= [4] "
7         Else
8             strResFilter = "������=" & Val(varPara) & " And ģ��=" & lngModual & " And ϵͳ=" & lngSys
9             strSqlFilter = "������=[6] And Nvl(ģ��,0)=[3] And Nvl(ϵͳ,0)=[4] "
10        End If
          
          '���������ж�
11        If grsParas Is Nothing Then
12            blnNew = True
13        ElseIf grsParas.State = 0 Then
14            blnNew = True
15        Else
16            grsParas.Filter = strResFilter
17            blnNewRow = grsParas.EOF
18        End If
          
19        If blnNew Or blnNewRow Then
              '��������ȡ��������
20            strSQL = "Select ID,Nvl(ϵͳ,0) as ϵͳ,Nvl(ģ��,0) as ģ��,Nvl(˽��,0) as ˽��,Nvl(����,0) as ����,Nvl(��Ȩ,0) as ��Ȩ,������,������," & _
                  " Nvl(����ֵ,ȱʡֵ) as ����ֵ,[1] as �û���,[2] as ������ From zlParameters Where " & strSqlFilter
21            Set rsTmp = OpenSQLRecord(selDB, strSQL, "GetPara", strDBUser, gUserInfo.ComputerName, lngModual, lngSys, CStr(varPara), Val(varPara))
          
22            If rsTmp.EOF Then
23                blnNotExists = True
24            Else
25                If blnNewRow Then
26                    grsParas.AddNew
27                    For i = 0 To rsTmp.Fields.Count - 1
28                        grsParas.Fields(i) = rsTmp.Fields(i).value
29                    Next
30                    grsParas.Update
31                Else
32                    Set grsParas = New ADODB.Recordset
33                    Set grsParas = CopyNewRec(rsTmp)
34                End If
                  '��ȡ�û��򱾻�����
35                If grsParas!˽�� = 1 Or grsParas!���� = 1 Then
36                    strSQL = "Select ����id, Nvl(�û���, 'NullUser') As �û���, Nvl(������, 'NullMachine') As ������, ����ֵ" & vbNewLine & _
                              "From zlUserParas" & vbNewLine & _
                              "Where ����id = [3]"
                              
37                    If grsParas!˽�� = 1 And grsParas!���� = 1 Then
38                        strSQL = strSQL & " And �û���=[1] And ������=[2]"
39                    ElseIf grsParas!˽�� = 1 Then
40                        strSQL = strSQL & " And �û���=[1] "
41                    Else
42                        strSQL = strSQL & " And ������=[2]"
43                    End If
                      
44                    Set rsTmp = OpenSQLRecord(selDB, strSQL, "GetPara", strDBUser, gUserInfo.ComputerName, Val(rsTmp!ID))
                      
45                    If grsUserParas Is Nothing Then
46                        Set grsUserParas = New ADODB.Recordset
47                        Set grsUserParas = CopyNewRec(rsTmp)
48                    ElseIf grsUserParas.State = 0 Then
49                        Set grsUserParas = New ADODB.Recordset
50                        Set grsUserParas = CopyNewRec(rsTmp)
51                    End If
                      
52                    Do While Not rsTmp.EOF
53                        grsUserParas.AddNew
54                        For i = 0 To rsTmp.Fields.Count - 1
55                            grsUserParas.Fields(i) = rsTmp.Fields(i).value
56                        Next
57                        grsUserParas.Update
58                        rsTmp.MoveNext
59                    Loop
60                End If
61            End If
62        End If

63        If blnNotExists Then
64            GetPara = strDefault
65        Else
              '��ȡ����ֵ
66            If grsParas!˽�� = 1 Or grsParas!���� = 1 Then
67                grsUserParas.Filter = "����ID=" & grsParas!ID & _
                      IIf(grsParas!˽�� = 1, " And �û���='" & grsParas!�û��� & "'", " And �û���='NullUser'") & _
                      IIf(grsParas!���� = 1, " And ������='" & grsParas!������ & "'", " And ������='NullMachine'")
68                If Not grsUserParas.EOF Then
69                    GetPara = NVL(grsUserParas!����ֵ, strDefault)
70                Else
71                    GetPara = NVL(grsParas!����ֵ, strDefault)
72                End If
73            Else
74                GetPara = NVL(grsParas!����ֵ, strDefault)
75            End If
              
              '���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
76            If grsParas!ϵͳ <> 0 And grsParas!ģ�� = 0 And grsParas!˽�� = 0 And grsParas!���� = 0 Then
77                intType = 1
78            ElseIf grsParas!ģ�� = 0 And grsParas!˽�� = 1 And grsParas!���� = 0 Then
79                intType = 2
80            ElseIf grsParas!ϵͳ <> 0 And grsParas!ģ�� <> 0 And grsParas!˽�� = 0 And grsParas!���� = 0 Then
81                intType = 3
82            ElseIf grsParas!ϵͳ <> 0 And grsParas!ģ�� <> 0 And grsParas!˽�� = 1 And grsParas!���� = 0 Then
83                intType = 4
84            ElseIf grsParas!ϵͳ <> 0 And grsParas!ģ�� <> 0 And grsParas!˽�� = 0 And grsParas!���� = 1 Then
85                intType = IIf(grsParas!��Ȩ = 1, 15, 5)
86            ElseIf grsParas!ϵͳ <> 0 And grsParas!ģ�� <> 0 And grsParas!˽�� = 1 And grsParas!���� = 1 Then
87                intType = 6
88            End If
              
              '�����Ӧ�Ŀؼ���ɫ���ɿ�״̬
89            If IsArray(arrControl) And (intType = 3 Or (intType Mod 10) = 5) Then
90                blnEnabled = Not ((intType = 3 Or (intType Mod 10) = 5 And grsParas!��Ȩ = 1) And Not blnSetup)
91                For i = 0 To UBound(arrControl)
92                    Select Case TypeName(arrControl(i))
                      Case "Label"
93                        arrControl(i).ForeColor = vbBlue
94                    Case "TextBox", "MaskEdBox", "CheckBox", "OptionButton", "ComboBox", "ListBox", "Frame", "PictureBox", "ListView"
95                        arrControl(i).ForeColor = vbBlue
96                        If Not blnEnabled Then arrControl(i).Enabled = False
97                    Case "CommandButton", "DTPicker"
98                        If Not blnEnabled Then arrControl(i).Enabled = False
99                    Case "MSHFlexGrid"
100                       arrControl(i).ForeColor = vbBlue
101                       arrControl(i).ForeColorFixed = vbBlue
102                       If Not blnEnabled Then arrControl(i).Enabled = False
103                   Case "VSFlexGrid"
104                       arrControl(i).ForeColor = vbBlue
105                       arrControl(i).ForeColorFixed = vbBlue
106                       If Not blnEnabled Then arrControl(i).Editable = 0
107                   Case Else
108                       On Error Resume Next
109                       arrControl(i).ForeColor = vbBlue
110                       If Not blnEnabled Then arrControl(i).Enabled = False
111                       Err.Clear
112                   End Select
113               Next
114           End If
115       End If
          

116       Exit Function
GetPara_Error:
117       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(GetPara)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
118       Err.Clear
End Function

Public Function GetPrivFunc(ByVal selDB As Integer, lngSys As Long, lngProgId As Long) As String
'���ܣ����ص�ǰ�û����е�ָ������Ĺ��ܴ�
'������lngSys     ����ǹ̶�ģ�飬��Ϊ0
'      lngProgId  �������
'���أ��ֺż���Ĺ��ܴ�,Ϊ�ձ�ʾû��Ȩ��
    Dim rsTmp As ADODB.Recordset, blnNew As Boolean
    Dim strSQL As String, strPrivs As String
    Dim blnRegCheck As Boolean
        
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
        blnNew = True
    Else
        On Error Resume Next
        strPrivs = gcolPrivs("_" & lngSys & "_" & lngProgId)
        If Err.Number > 0 Then blnNew = True: Err.Clear
    End If
    
    If blnNew Then
        strSQL = "Select Text as ���� From Table(Cast(zltools.f_Reg_Func([1],[2]) as zlTools.t_Reg_Rowset))"
        
Beging:
        Set rsTmp = OpenSQLRecord(selDB, strSQL, "GetPrivFunc", lngSys, lngProgId)
        On Error GoTo errH
        
        Do While Not rsTmp.EOF
            strPrivs = strPrivs & ";" & rsTmp!����
            rsTmp.MoveNext
        Loop
        strPrivs = Mid(strPrivs, 2)
        gcolPrivs.Add strPrivs, "_" & lngSys & "_" & lngProgId
    End If
    On Error GoTo 0
    
    GetPrivFunc = strPrivs
    Exit Function
errH:
    If Not blnRegCheck Then
        '�������,����������û�е���zlRegCheck���,�Զ�����һ��,����ٳ���,����ʾ.
        If selDB = 1 Then
            If initRegister = True Then
               zlRegister.zlRegInit gcnLisOracle
            End If
            If FunzlRegCheck(, gcnLisOracle) <> "" Then Exit Function
        Else
'            If initRegister = True Then
'               zlRegister.zlRegInit gcnHisOracle
'            End If
            If FunzlRegCheck(, gcnHisOracle) <> "" Then Exit Function
        End If
'        GetPrivFunc = zlRegister.zlRegFunc(lngSys, lngProgId)
        blnRegCheck = True
        GoTo Beging
    End If
End Function

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset) As ADODB.Recordset
      '������:����
      '��������:2000-11-02
      '���Ƽ�¼��
      '�ڳ����У��������漰���໥���ݼ�¼������ʹ��ADO��Clone���Ʋ����ļ�¼����������һ����¼�������ݷ����仯��ʱ�����и�������������ͬ�ı仯��ͨ��ָ�޸Ļ�ɾ����������������ϣ����Щ��¼���໥�䱣�ֶ���
          Dim rsClone As New ADODB.Recordset
          Dim rsTarget As New ADODB.Recordset
          Dim intFields As Integer
          
1         On Error GoTo CopyNewRec_Error

2         Set rsClone = rsSource.Clone
3         rsClone.Filter = rsSource.Filter
4         Set rsTarget = New ADODB.Recordset
5         With rsTarget
6             For intFields = 0 To rsClone.Fields.Count - 1
7                 .Fields.Append rsClone.Fields(intFields).Name, IIf(rsClone.Fields(intFields).Type = adNumeric, adDouble, rsClone.Fields(intFields).Type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
8             Next
              
9             .CursorLocation = adUseClient
10            .CursorType = adOpenStatic
11            .LockType = adLockOptimistic
12            .Open
              
13            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
14            Do While Not rsClone.EOF
15                .AddNew
16                For intFields = 0 To rsClone.Fields.Count - 1
17                    .Fields(intFields) = rsClone.Fields(intFields).value
18                Next
19                .Update
20                rsClone.MoveNext
21            Loop
22        End With
          
23        Set CopyNewRec = rsTarget


24        Exit Function
CopyNewRec_Error:
25        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "ִ��(CopyNewRec)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
26        Err.Clear
End Function

Public Function ReplaseSpecial(strTmp As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����               �滻�����ַ�
    '����
    '                   ���滻���ַ�
    '����               ���滻�������ַ�����ִ�
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intloop As Integer
    Dim strSpecial As String
    Dim astrtmp() As String
    strSpecial = "'^��^��^;^��^:^��^?^��^|^,^��^.^��^"""
    astrtmp = Split(strSpecial, "^")
    For intloop = 0 To UBound(astrtmp)
        strTmp = Replace$(strTmp, astrtmp(intloop), "")
    Next
    ReplaseSpecial = strTmp
    
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "Select zlWBcode([1]) From Dual"
    Else
        strSQL = "Select zlSpellcode([1]) From Dual"
    End If
    On Error GoTo Errhand
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "zlGetSymbol", strInput)
    zlGetSymbol = IIf(IsNull(rsTmp.Fields(0).value), "", rsTmp.Fields(0).value)
    Exit Function
Errhand:
'    If gobjComLib.ErrCenter() = 1 Then Resume
'    Call gobjComLib.SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/27
'��    ��:�ṩ�����ط����ò������ã����棬ִ�й��̣���ѯ����
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function ComSetPara(ByVal selDB As Integer, ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '���ò���
    ComSetPara = SetPara(selDB, varPara, strValue, lngSys, lngModual, blnSetup)
End Function

Public Function ComGetPara(ByVal selDB As Integer, ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer) As String
    'ȡ����
    
    ComGetPara = GetPara(selDB, varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType)
    
End Function

Public Function ComGetPrivs(ByVal selDB As Integer, ByVal lngSys As Long, ByVal lngModul As Long) As String
    '��ȡģ��Ȩ��
   ComGetPrivs = GetPrivFunc(selDB, lngSys, lngModul)
End Function

Public Function StringFormatDate(strDate, Optional MinOrMax As Integer) As String
    '����               ��ʽ�����浽���ݿ�����ڸ�ʽ
    '����               strDate �������ڸ�ʽ
    '                   MinOrMax 1=��С 2=���
    '����               ��ʽ���õ����ڸ�ʽ
    Select Case MinOrMax
        Case 0
            StringFormatDate = "TO_DATE('" & Format(strDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
        Case 1
            StringFormatDate = "TO_DATE('" & Format(strDate, "yyyy-MM-dd 00:00:00") & "','yyyy-mm-dd hh24:mi:ss')"
        Case 2
            StringFormatDate = "TO_DATE('" & Format(strDate, "yyyy-MM-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')"
    End Select
End Function

Public Function GetUserInfo(ByVal intSelDB As Integer) As ADODB.Recordset
      '���ܣ���ȡ��ǰ�û��Ļ�����Ϣ
      '���أ�����Ado��¼��
          Dim strSQL As String
          Dim strDefault As String
          Dim strDBUser As String

1         On Error GoTo GetUserInfo_Error
2         strDBUser = GetUserDB(intSelDB)
3         strDefault = " And C.ȱʡ = 1"
4         strSQL = "Select User,A.Id, A.���, A.����, A.����, B.�û���, C.����id, D.���� As ������, D.���� As ������" & vbNewLine & _
                   "From ��Ա�� A, �ϻ���Ա�� B, ������Ա C, ���ű� D" & vbNewLine & _
                   "Where A.Id = B.��Աid And A.Id = C.��Աid And C.����id = D.Id And B.�û��� = [1]"

5         Set GetUserInfo = OpenSQLRecord(intSelDB, strSQL & strDefault, "GetUserInfo", strDBUser)
6         If GetUserInfo.RecordCount = 0 Then
7             strDefault = " And Rownum < 2"
8             Set GetUserInfo = OpenSQLRecord(intSelDB, strSQL & strDefault, "GetUserInfo", strDBUser)
9         End If

10        Exit Function


11        Exit Function
GetUserInfo_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "mdlPubLisComlib", "ִ��(GetUserInfo)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
13        Err.Clear

End Function

Public Function GetUserDB(ByVal selDB As Integer) As String
          Dim strTmp As String
          Dim strConnStr As String

1         On Error GoTo GetUserDB_Error

2         If selDB = Sel_Lis_DB Then
3             strConnStr = gcnLisOracle.ConnectionString
4         ElseIf selDB = Sel_His_DB Then
5             strConnStr = gcnHisOracle.ConnectionString
6         End If
7         strTmp = Mid(strConnStr, InStr(strConnStr, "User ID="))
8         strTmp = Mid(strTmp, 9, InStr(strTmp, ";") - 9)
9         GetUserDB = UCase(strTmp)


10        Exit Function
GetUserDB_Error:
11        Call WriteErrLog("zlPublicHisCommLis", "mdlPubLisComlib", "ִ��(GetUserDB)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
12        Err.Clear
End Function

'--------------------------------------------------
'���ܣ���֤ϵͳע����Ȩ����ȷ��
'������blnTemp-�Ƿ��δ�������ʱע����Ϣ��֤
'���أ���ȷ����"";���󷵻ش�����Ϣ
'--------------------------------------------------
Public Function HiszlRegCheck(ByVal selDB As Integer, Optional blnTemp As Boolean) As String
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim arrMd5_1(5) As String
    Dim arrMd5_2(5) As String
    Dim arrMd5_3(5) As String
    Dim arrMd5_4(5) As String
    Dim arrMd5_5(5) As String
    Dim strMD5 As String
    Dim intLine As Integer
    On Error GoTo Errhand
    
    '---------------------------------Beging ��֤ F_Reg_Audit �Ƿ��滻
    '-- ������ 9.25 HIS 10.15
    arrMd5_1(0) = "6746B20191FD2AA9B0E08AFB44E80D4B"
    arrMd5_1(1) = "93C94497A547C10EC3B5C95F5188BA5D"
    arrMd5_1(2) = "A5596EA1AB4F6D4939CBD9599CBFBA0F"
    arrMd5_1(3) = "07069FF5FF76C204EEFCC88366F6A495"
    arrMd5_1(4) = "73C7DB3F742EBC654FAC289B4D37A7B0"
        
    '-- ������ 9.35 HIS 10.24
    arrMd5_2(0) = "10E1A9794EF861981C7F53D887990B1F"
    arrMd5_2(1) = "C4A92BE1F6882A57564206E9B391A600"
    arrMd5_2(2) = "F4878F9061BFC4357DC4545EAC326CD2"
    arrMd5_2(3) = "4BBF3E2A0D667A50B8CBC443A1110EA2"
    arrMd5_2(4) = "07BC27215593F6ED86C9905C0D215BD9"
        
    '-- ������ 9.37 HIS 10.26
    arrMd5_3(0) = "4D1B31CCB39BDCCE4EE61357555DAD9D"
    arrMd5_3(1) = "F544A3A12A833F6EE10CEA514D65782C"
    arrMd5_3(2) = "5CEF0276B15026C1D5546A85F9A3BE1F"
    arrMd5_3(3) = "487CC8AD6D5F2E0DC337677D02EA702F"
    arrMd5_3(4) = "20AD16738F21A228D962E59DAECB0D84"
    
    '-- ������ 9.41 HIS 10.30
    arrMd5_4(0) = "01322819F7B38E12BCAA8525895EF288"
    arrMd5_4(1) = "75E62456DB5F6742B9140DFB73D094FE"
    arrMd5_4(2) = "4270A613EA65B66BF4200BA42F205319"
    arrMd5_4(3) = "64FD2D54E72F9F647DD01D14116988AE"
    arrMd5_4(4) = "D7A22AF77FAC34E04086B800570BCB37"
        
    '-- ������ 9.45 HIS 10.34
    arrMd5_5(0) = "01322819F7B38E12BCAA8525895EF288"
    arrMd5_5(1) = "02AC74A017BEE67D26051B4BA5DA98E8"
    arrMd5_5(2) = "9D1143BA317F835426BB8ED2F319A8CA"
    arrMd5_5(3) = "E2718B7863EB402205FAC8CDD348D649"
    arrMd5_5(4) = "39A9E549EAB1EDD396230AD61DC559B0"
    '������������һ��ִ�У�RowNum���Ǻ�Line�����Ӧ�ģ��ڶ���ִ���Ժ����������������Ӳ�ѯ
    strSQL = "Select Դ��, Rownum As Line" & vbNewLine & _
            "From (Select Substr(Text, 1, 512) As Դ��" & vbNewLine & _
            "       From All_Source" & vbNewLine & _
            "       Where Owner = 'ZLTOOLS' And Name = 'F_REG_AUDIT' And Line In (3, 5, 7, 9, 11)" & vbNewLine & _
            "       Order By Line)"

    Set rsTemp = OpenSQLRecord(selDB, strSQL, "zlRegCheck")
    Do Until rsTemp.EOF
        strMD5 = Md5_String_Calc("" & rsTemp!Դ��)
        intLine = Val("" & rsTemp!Line)
        If Not (arrMd5_1(intLine - 1) = strMD5 Or arrMd5_2(intLine - 1) = strMD5 _
            Or arrMd5_3(intLine - 1) = strMD5 Or arrMd5_4(intLine - 1) = strMD5 _
            Or arrMd5_5(intLine - 1) = strMD5) Then
            HiszlRegCheck = "ע����֤������ȷ����ʹ����ȷ��ע�����"
            Exit Do
        End If
        rsTemp.MoveNext
    Loop
    If HiszlRegCheck <> "" Then Exit Function
    '---------------------------------          End  ��֤ F_Reg_Audit �Ƿ��滻
    
    strSQL = "Select zltools.f_Reg_Audit([1]) As Stamp From zltools.zlRegInfo r Where ��Ŀ='��Ȩ֤��'"
    Set rsTemp = OpenSQLRecord(selDB, strSQL, "zlRegCheck", IIf(blnTemp, 1, 0))
    If rsTemp.RecordCount > 0 Then
        If Left(rsTemp.Fields(0).value, 6) <> "ERROR-" Then
            HiszlRegCheck = ""
        Else
            HiszlRegCheck = rsTemp.Fields(0).value
        End If
    Else
        HiszlRegCheck = "ע����Ϣ��ʧ,������ע��ǰ"
    End If
    Exit Function
Errhand:
    HiszlRegCheck = Err.Description
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/10/11
'��    ��:�����ض���������µĺ���,����������ֻ����ZLHIS10������ҪOracle 8i(8.1.5)���ϰ汾֧��
'֧�ִ��벻ͬ���Ӻ����ɲ�ͬ���ݿ�����ӣ�����HIS,LIS
'������
'int���=��Ŀ���:
'  1   ����ID ����
'  2   סԺ�� ����
'  3   ����� ����
'  10  ҽ�����ͺ� ����,˳��������
'  x   �������ݺ� �ַ�,���ݱ�Ź���˳��������,���Զ���ȱ
'lng����ID=�����Һ����Ź������Ŀ��Ҫ
'���أ�������
'˵����
'  ��Ź���0-����˳����,1-����˳����,2-��ִ�п��ҷ��±��(��Ҫ��ȡ���Һ����)
'            ������ţ�0-˳����,1-������(YYMMDD)+˳���(0000)
'            ��סԺ�ţ�0-˳����,1-����(YYMM)+˳���(0000),2-��(YYYY)+˳���(00000)
'  ���λȷ������1990Ϊ���������������������0��9/A��Z��˳����Ϊ��ȱ���
'  ������-10���������Ʊ�,���ڲ�������²�ȱ��(ȡ�˺�,��δʹ��)
'  For Update�ڲ��������������,����Waitѡ���Ա���������߷��ؿ�
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function GetNextNo(ByVal selDB As Integer, ByVal int��� As Integer, Optional ByVal lng����ID As Long, Optional ByVal strTag As String, Optional ByVal intStep As Integer = 1) As Variant

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo GetNextNo_Error

    GetNextNo = Null
    
    
    strSQL = "Select NextNO([1],[2],[3],[4]) as NO From Dual"
    Set rsTmp = OpenSQLRecord(selDB, strSQL, "GetNextNo", int���, lng����ID, strTag, intStep)
    
'    If gcnOracle.Errors.Count > 0 Then 'Select�к�������ʱ,��VB�в��Զ���������
'        Err.Raise gcnOracle.Errors(0).Number, , gcnOracle.Errors(0).Description
'    End If
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!No) Then GetNextNo = rsTmp!No
    End If

    Exit Function
GetNextNo_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlPubLisComlib", "ִ��(GetNextNo)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, False)
    Err.Clear

End Function


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/10/11
'���ܣ���ȡָ��������Ӧ������(���淶������������Ϊ��������_id��)����һ��ֵ
'������
'   strTable��������
'���أ�
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function GetNextId(strTable As String) As Long
    Dim rsTmp           As New ADODB.Recordset
    Dim strSQL As String, strtab As String

    '�����ô������,ԭ��������ʧЧ��û������ʱ,Ӧ�÷��ش���,��Ȼ������,��������!
    'On Error GoTo errH
    strtab = Trim(strTable)
    If strtab = "������ü�¼" Or strtab = "סԺ���ü�¼" Then strtab = "���˷��ü�¼"

    strSQL = "Select " & strtab & "_ID.Nextval From Dual"
    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "����ID����")
    If Not rsTmp.EOF Then
        GetNextId = rsTmp.Fields(0).value
    End If
    '    Exit Function
    'errH:
    '    If gobjComLib.ErrCenter() = 1 Then Resume
End Function


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/10/15
'��    �ܣ��Ƚ������汾��,�ȵ�ǰ�汾��С������1����ȷ���0���ȵ�ǰ�汾�Ŵ󷵻�-1
'��    ����strVerCur=��ǰ�汾��
'         strVerCom=�Աȵİ汾��
'��    �أ��ԱȰ汾�űȵ�ǰ�汾��С������1����ȷ���0���ȵ�ǰ�汾�Ŵ󷵻�-1
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function VerCompare(ByVal strVerCur As String, Optional ByVal strVerCom As String) As Integer

    If VerFull(strVerCur) < VerFull(strVerCom) Then
        VerCompare = -1
    ElseIf VerFull(strVerCur) > VerFull(strVerCom) Then
        VerCompare = 1
    Else
        VerCompare = 0
    End If
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
'���ܣ�����VB���֧�ֵİ汾����ʽ:9999.9999.9999.9999,��С�汾��0000.0000.0000.0000
'������strVer=��ǰ�汾��
'           blnMax=True,����Ϊ�գ��򷵻����֧�ְ汾��False=����Ϊ�գ��򷵻���С֧�ְ汾
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999.9999", "0000.0000.0000.0000")
        Exit Function
    End If
    '����һ�Σ��Լ�������SP�汾��
    arrVer = Split(strVer & ".0", ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000") & "." & Format(arrVer(3), "0000")
End Function

Public Function IsVerSion(ByVal strVer As String) As Boolean
'���ܣ��ж��ַ����Ƿ��ǰ汾��
    Dim arrVer As Variant
    Dim i As Integer
    If Not strVer Like "*.*.*" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    
    For i = LBound(arrVer) To UBound(arrVer)
        If Not IsNumeric(arrVer(i)) Then Exit Function
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If i = 3 Then
            If Format(Val(arrVer(i)), "0000") <> Format(Trim(arrVer(i)), "0000") Then Exit Function
        Else
            If Val(arrVer(i)) & "" <> Trim(arrVer(i)) Then Exit Function
        End If
    Next
    
    IsVerSion = True
End Function


'---------��������ʹ��
Public Function ReadLob(ByVal lngSys As Long, ByVal Action As Long, ByVal KeyWord As String, _
                        Optional ByVal strFile As String, Optional ByVal bytFunc As Byte = 0, _
                        Optional bytMoved As Byte = 0) As String
'���ܣ���ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'������
'lngSys:ϵͳ���
'Action:�������ͣ����������ǲ����ĸ���
'---ϵͳ100,Zl_Lob_Read
'0-�������ͼ��;1-�����ļ���ʽ;2-�����ļ�ͼ��;3-�������ĸ�ʽ;4-��������ͼ��;
'5-���Ӳ�����ʽ;6-���Ӳ���ͼ��;7-����ҳ���ʽ(ͼ��)��8-���Ӳ�������;9-�����ص����
'10-�ٴ�·���ļ�,11-�ٴ�·��ͼ��;14-��Ա֤���¼;15-��Ա��;16-��Ա��Ƭ;
'17-ҩƷ���(ʹ��˵��);18-ҩƷ���(ͼƬ);
'19-������չ��Ϣ;20-��Ա��չ��Ϣ;22-ҽ����������;23-��Ӧ����Ƭ;24-�Զ������뵥�ļ�;25-ҽ�����뵥�ļ�
'26-����·���ļ�,27-������Ƭ,28-��ѯͼƬԪ��,29-��ѯ����Ŀ¼
'--ϵͳ600��ZL6_Lob_Read
'0-�豸��Ƭ
'---ϵͳ2400,Zl24_Lob_Read
'���鳣��ͼ��,��Action
'---ϵͳ2100,Zl21_Lob_Read
'1-�������͵���;2-���������(��ͼƬֻ�ж�ȡ��û�б���);3-����걨��¼;4-���������Ա,5-���������
'---ϵͳ2500,ZL25_Lob_Read
'0-΢����ͿƬ����
'---ϵͳ2600,Zl26_Lob_Read
'14-����ؼ�Ŀ¼,15-������ԴĿ¼
'      KeyWord:ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'      strFile:�û�ָ����ŵ��ļ�������ָ��ʱ���Զ�ȡ��ʱ�ļ���
'bytFunc-0-BLOB,1-CLOB
'bytMoved=0������¼,1��ȡת���󱸱��¼
'���أ�������ݵ��ļ�����ʧ���򷵻��㳤��""
    Const conChunkSize As Long = 10240
    Dim rsLOB       As ADODB.Recordset
    Dim lngFileNum  As Long, lngCount       As Long, lngBound       As Long
    Dim aryChunk()  As Byte, strText        As String
    Dim strSQL      As String
    Dim objFile     As New FileSystemObject
    Dim lngCurSize  As Long
    
    Err = 0: On Error GoTo Errhand
    Select Case lngSys \ 100
        Case 1
            strSQL = "Select Zl_Lob_Read([1],[2],[3],[4],[5]) as Ƭ�� From Dual"
        Case 6
            strSQL = "Select Zl6_Lob_Read([1],[2],[3],[4],[5]) as Ƭ�� From Dual"
        Case 24
            strSQL = "Select Zl24_Lob_Read([2],[3]) as Ƭ�� From Dual"
        Case 21
            strSQL = "Select Zl21_Lob_Read([1],[2],[3]) as Ƭ�� From Dual"
        Case 25
            strSQL = "Select Zl25_Lob_Read([1],[2],[3],[4],[5]) as Ƭ�� From Dual"
        Case 26
            strSQL = "Select Zl26_Lob_Read([1],[2],[3]) as Ƭ�� From Dual"
    End Select
    If strSQL = "" Then strFile = "": Exit Function
    If bytFunc = 0 Then 'BLOB
        If strFile = "" Then
            strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & objFile.GetTempName
        End If
        lngFileNum = FreeFile
        Open strFile For Binary As lngFileNum
        lngCount = 0
        lngCurSize = 0
        Do
            Set rsLOB = OpenSQLRecord(Sel_His_DB, strSQL, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLOB.EOF Then Exit Do
            If IsNull(rsLOB.Fields(0).value) Then Exit Do
            strText = rsLOB.Fields(0).value
            If lngCurSize = 0 Then
                lngCurSize = Len(strText) / 2
                If lngCurSize = 0 Then Exit Do
                ReDim aryChunk(lngCurSize - 1) As Byte
            ElseIf lngCurSize <> Len(strText) / 2 Then '��ֹ�ظ������ڴ�
                lngCurSize = Len(strText) / 2
                If lngCurSize = 0 Then Exit Do
                ReDim aryChunk(lngCurSize - 1) As Byte
            End If
            For lngBound = LBound(aryChunk) To UBound(aryChunk)
                aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
            Next
            Put lngFileNum, , aryChunk()
            lngCount = lngCount + 1
        Loop
        Close lngFileNum
        If lngCount = 0 Then Kill strFile: strFile = ""
    Else  'CLOB
        lngCount = 0
        strFile = ""
        Do
            Set rsLOB = OpenSQLRecord(Sel_His_DB, strSQL, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLOB.EOF Then Exit Do
            If IsNull(rsLOB.Fields(0).value) Then Exit Do
            strText = rsLOB.Fields(0).value
            strFile = strFile & strText
            lngCount = lngCount + 1
        Loop
    End If
    ReadLob = strFile
    Exit Function
Errhand:
    If bytFunc = 0 Then
        Close lngFileNum
        Kill strFile: ReadLob = ""
    End If
    Err.Clear
End Function
