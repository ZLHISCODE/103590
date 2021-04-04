Attribute VB_Name = "mdlProc"
Option Explicit
'ģ��˵��:���������ģ��
'���صĹ��̼��϶���Ϊ
'       "P_Name", adVarChar, 32 ��������
'       "P_Define", adLongVarChar, 9999999#  ���̶���(�����Ĺ����ı�)
'       "P_System", adVarChar, 20   ϵͳ����
'       "P_SysNum", adInteger, 5 ϵͳ���
'       "P_Owner", adVarChar, 20   ϵͳ������
'       "P_Ver", adVarChar, 20  �ű��ļ��汾

Public gstrBCode As New clsStringBulider

Public Sub GetProceduresByFile(ByVal strFile As String, rsProcedure As ADODB.Recordset, _
                                            Optional ByVal strFileVer As String, Optional ByVal lngSysNum As Long, _
                                            Optional ByVal strSysName As String, Optional ByVal strOwner As String)
    '���ݴ�����ļ�����,���ؼ�¼��
    '����:strVer �ļ���Ӧ�汾
    Dim objTxt As TextStream
    Dim arrTxt() As String, dblRow As Double
    Dim strLine As String, strFMT As String
    Dim blnBegin As Boolean, strPName As String
    Dim arrDelete() As String, strProcName As String
    Dim i As Long
    
    On Error GoTo errH
    If Not gobjFile.FileExists(strFile) Then Exit Sub
    If gobjFile.GetFile(strFile).Size = 0 Then Exit Sub '�ļ�Ϊ��
    
    If strFileVer = "" Then
        strFileVer = Mid(strFile, InStrRev(strFile, "\") + 1)
    End If
    If rsProcedure Is Nothing Then
        Set rsProcedure = New ADODB.Recordset
        With rsProcedure
            .Fields.Append "P_Name", adVarChar, 32 '��������
            .Fields.Append "P_Define", adLongVarChar, 9999999#  '���̶���
            
            .Fields.Append "P_System", adVarChar, 20   'ϵͳ����
            .Fields.Append "P_SysNum", adInteger, 5 'ϵͳ���
            .Fields.Append "P_Owner", adVarChar, 20   'ϵͳ������
            .Fields.Append "P_Ver", adVarChar, 50  '�ű��ļ��汾
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
    End If
    
    'һ�ν��ı��ļ��е����ݶ���ȡ����,��������arrTxt��
    Set objTxt = gobjFile.OpenTextFile(strFile)
    arrTxt = Split(objTxt.ReadAll, vbNewLine)
    objTxt.Close
    
    gstrBCode.Clear
    'ѭ��,��ÿһ�εĹ������ƺ͹��̶��屣�浽��¼����
    ReDim arrDelete(0)
    For dblRow = 0 To UBound(arrTxt)
        strLine = RTrim(arrTxt(dblRow))
        strFMT = UCase(TrimComment(TrimEx(strLine)))
        
        '������к���Drop Procedure��� ,�Ͱѹ������Ƽ�¼����,�����Ӽ�¼�аѸù���ɾ��
        If InStr(1, strFMT, "DROP PROCEDURE") > 0 Then
            strProcName = Mid(strFMT, InStr(1, strFMT, "DROP PROCEDURE") + Len("DROP PROCEDURE "))  '��ȡ
            strProcName = Split(strProcName, " ")(0)    'ȡ��һ���ո�֮ǰ
            If InStr(1, strProcName, ".") > 0 Then strProcName = Split(strProcName, ".")(1) '�ж��Ƿ���������
            If InStr(1, strProcName, ";") > 0 Then strProcName = Left(strProcName, Len(strProcName) - 1) '����ǷֺŽ�β,Ӧ��ȥ���ֺ�
            arrDelete(UBound(arrDelete)) = strProcName
            ReDim Preserve arrDelete(UBound(arrDelete) + 1)
        End If
        
        '��ʼ��¼����
        If strFMT Like "CREATE*PROCEDURE *" Or strFMT Like "CREATE*FUNCTION *" Then
            strPName = Split(strFMT, " ")(4)
            If InStr(1, strPName, "(") > 0 Then strPName = Left(strPName, InStr(1, strPName, "(") - 1)
            If InStr(1, strPName, ".") > 0 Then strPName = Split(strPName, ".")(1)  '�п��ܽű��еĹ�����ǰ�� ������. ��: zltools.zl_xxx

            blnBegin = True
            
            gstrBCode.Append Replace(strLine, """", "") '�����������������" Ӧ��ȥ��
        Else
            '������¼����
            If (strFMT = "/" Or UBound(arrTxt) = dblRow) And blnBegin Then
                    rsProcedure.Filter = "P_Name = '" & strPName & "'"
                    If rsProcedure.RecordCount = 0 Then
                        rsProcedure.AddNew
                        rsProcedure!P_Name = strPName
                    End If
                
                    rsProcedure!P_Define = gstrBCode.ToString
                    rsProcedure!P_Ver = strFileVer
                    
                    
                    If lngSysNum <> 0 Then
                        rsProcedure!P_SysNum = lngSysNum
                    End If
                    If strSysName <> "" Then
                        rsProcedure!P_System = strSysName
                    End If
                    If strOwner <> "" Then
                        rsProcedure!P_Owner = strOwner
                    End If
                    
                    rsProcedure.Update
                    
                    blnBegin = False
                    gstrBCode.Clear
            ElseIf blnBegin Then
                gstrBCode.Append vbNewLine
                gstrBCode.Append Left(strLine, 4000)
            End If
        End If
    Next
    
    '����ű�����Drop Procedure��� ,�ʹӼ�¼���аѹ���ɾ��
    For i = 0 To UBound(arrDelete)
        rsProcedure.Filter = "P_Name  = '" & arrDelete(i) & "'"
        If rsProcedure.RecordCount <> 0 Then
            rsProcedure.Delete
        End If
    Next
    
    rsProcedure.Filter = 0
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "��ȡ�ļ�ʧ��" & err.Description, , gstrSysName
End Sub

Public Function LoadBaseProcs(ByVal strProcName As String) As String
    '���ܣ��������ݿ�洢����
    Dim rsSource As ADODB.Recordset, strSQL As String
    Dim strProcText As String, i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    '�洢�����ռ����ռ����ݿ���Ϊ�����洢����
    strSQL = "Select Name, Type, Text, Line ��� From User_Source Where Type In ('PROCEDURE', 'FUNCTION') And Name =[1] Order By  Line"
    Set rsSource = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ���ݿ����Դ��", strProcName)
    
    gstrBCode.Clear
    
    If Not rsSource.EOF Then
        Do While Not rsSource.EOF
            strTmp = rsSource!Text
            strTmp = Replace(strTmp, vbCr, "")
            strTmp = Replace(strTmp, vbLf, "")
            strTmp = Replace(strTmp, vbNewLine, "")
            
            If rsSource!��� = 1 Then
                '���ݿ�Դ��û��CREATE OR REPLACE
                gstrBCode.Append "CREATE OR REPLACE "
            Else
                gstrBCode.Append vbNewLine
            End If
            
            If UCase(strTmp) Like "*" & """" & UCase(strProcName) & """" & "*" Then
                    strTmp = Replace(UCase(strTmp), """" & UCase(strProcName) & """", strProcName)
            End If
            
            gstrBCode.Append strTmp
            rsSource.MoveNext
        Loop
    End If
    LoadBaseProcs = gstrBCode.ToString
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function UpdateProc2DB(rsProc As ADODB.Recordset, intType As Integer, Optional strErr As String) As Boolean
    '�����̼��ϱ��������ݿ�
    '����:rsProc-���̼���  intType-��������(1-�䶯���� 2-�������޸ĵĹ���)
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim lngID As Long
    Dim arrTxt() As String, i As Long
    Dim lngSysNum As Long, strIDs As String, arrIds As Variant
    
    On Error GoTo errH
    strErr = ""
    If rsProc Is Nothing Then
        UpdateProc2DB = True
        Exit Function
    End If
    If rsProc.RecordCount = 0 Then
        UpdateProc2DB = True
        Exit Function
    End If
    
    With rsProc
        .Filter = 0
        
        Do While Not .EOF
            lngID = GetProcIdByName(!P_Name)
            gcnOracle.BeginTrans
            '����������zlProcedure
            If lngID = 0 Then
                If intType = 1 Then
                    strSQL = "Insert Into Zlprocedure (ID, ����, ����, ״̬, ������, ϵͳ���, ����ǰ�汾) Values" & vbNewLine & _
                                 "(Zlprocedure_Id.Nextval,1,'" & !P_Name & "',1,'" & !P_Owner & "'," & !P_SysNum & ",'" & !P_Ver & "')"
                Else
                    strSQL = "Insert Into Zlprocedure (ID, ����, ����, ״̬, ������, ϵͳ���, ������汾) Values" & vbNewLine & _
                                 "(Zlprocedure_Id.Nextval,1,'" & !P_Name & "',1,'" & !P_Owner & "'," & !P_SysNum & ",'" & !P_Ver & "')"
                End If
            Else
                'ɾ����ת��������
                gcnOracle.Execute "Delete from zlProcedureText where ����=3 and ����ID = (Select ID From zlProcedure where ״̬ = 4 And ID = " & lngID & ")"
                gcnOracle.Execute "Update zlProcedure Set ״̬ = 1 Where ״̬ = 4 And ID = " & lngID    'ֻ�޸���ת�����̵�״̬
                
                '��������
                If intType = 1 Then
                    strSQL = "Update zlProcedure Set ���� = 1,������='" & !P_Owner & "',ϵͳ���=" & !P_SysNum & ",����ǰ�汾='" & !P_Ver & "'" & vbNewLine & _
                                 "Where Id = " & lngID
                Else
                    strSQL = "Update zlProcedure Set ���� = 1,������='" & !P_Owner & "',ϵͳ���=" & !P_SysNum & ",������汾='" & !P_Ver & "'" & vbNewLine & _
                                 "Where Id = " & lngID
                End If
            End If
            gcnOracle.Execute strSQL
            
            'ɾ��zlProcedureText�е�����
            If lngID = 0 Then
                lngID = GetProcIdByName(!P_Name)
            End If
            
            If intType = 1 Then
                gcnOracle.Execute "Delete from zlProcedureText where ����=1 and ����ID = " & lngID
            Else
                gcnOracle.Execute "Delete from zlProcedureText where ����=4 and ����ID = " & lngID
            End If
            
            '������̶��嵽zlProcedureText
            arrTxt = Split(!P_Define, vbNewLine)
            strSQL = "Insert Into zlProcedureText(����ID,����,���,����) "
            For i = 0 To UBound(arrTxt)
                If i = UBound(arrTxt) Then
                    strSQL = strSQL & vbNewLine & "Select " & lngID & "," & IIf(intType = 1, "1", "4") & "," & (i + 1) & ",'" & Replace(arrTxt(i), "'", "''") & "' From Dual "
                Else
                    strSQL = strSQL & vbNewLine & "Select " & lngID & "," & IIf(intType = 1, "1", "4") & "," & (i + 1) & ",'" & Replace(arrTxt(i), "'", "''") & "' From Dual Union All "
                End If
            Next
            gcnOracle.Execute strSQL
            
            
            If strIDs = "" Then
                lngSysNum = !P_SysNum
                strIDs = lngID
            Else
                strIDs = strIDs & "," & lngID 'ƴ������ID
            End If
            
            gcnOracle.CommitTrans
            .MoveNext
        Loop
    End With
    
    'ɾ���Ǹ�ϵͳ����������,��Ϊ�еĿ�zlProcedureText��������Ǽ���ɾ��,���Ҫ��ɾ���ӱ�
    If intType = 1 Then
        gcnOracle.BeginTrans
        arrIds = TranStr2Var(strIDs, ",", 2000) '��ֹ�ַ�����
        For i = 0 To UBound(arrIds)
            strSQL = "Delete From zlProcedureText Where ����ID In  " & vbNewLine & _
                        "(Select ID from Zlprocedure Where ���� = 1 And ϵͳ��� = " & lngSysNum & " And  ID Not In (Select Column_Value From Table(f_Str2list('" & arrIds(i) & "', ','))))"
            gcnOracle.Execute strSQL
        
            strSQL = "Delete From zlProcedure Where ���� = 1 And ϵͳ��� = " & lngSysNum & " And  ID Not In (Select Column_Value From Table(f_Str2list('" & arrIds(i) & "', ',')))"
            gcnOracle.Execute strSQL
        Next
        
        gcnOracle.CommitTrans
    End If
    UpdateProc2DB = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    strErr = err.Description
End Function


Public Function GetProcIdByName(ByVal strName As String, Optional ByVal intProp As Integer, Optional ByVal intStat As Integer) As Long
    '�������Ʒ��ع���ID
    '����˵��:
    'strName -����
    'intPorc-����-1-�û��䶯����;2-�հ׹���;3-�û�����
    'intStat-״̬:1-������;2-���Զ�����;3-���˹�����;4-�ѵ���
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim lngID As Long
    
    On Error GoTo errH
    strSQL = "Select Id From zlProcedure Where ���� = [1]" & IIf(intProp = 0, "", " And ���� = [2]") & IIf(intStat = 0, "", "And ״̬ = [3]")
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡID", strName, intProp, intStat)
    
    
    If rsTmp.RecordCount = 0 Then
        lngID = 0
    Else
        lngID = rsTmp!id
    End If
    
    GetProcIdByName = lngID
    Exit Function
errH:
    MsgBox "��ȡ����ID����" & vbNewLine & err.Description, , gstrSysName
End Function

Public Function GetPorcTxtByName(ByVal strName As String, ByVal intType As Integer) As String
    '���ݹ������ƺ��ı����ͷ��ع����ı�
    'strName:��������  intType:�ı����� 1-�ϴζ������;2-�ϴα�׼����;3-�����Զ�����;4-���α�׼����
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strResult As String
    
    On Error GoTo errH
    
    strSQL = "Select ����  From zlProcedureText Where ���� = [2]  And ����ID = (Select ID From zlProcedure Where ����=[1] ) Order by ��� "
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ�����ı�", strName, intType)
    
    If rsTmp.RecordCount = 0 Then
        Exit Function
    End If
    
    Do While Not rsTmp.EOF
        If strResult = "" Then
            strResult = rsTmp!����
        Else
            strResult = strResult & vbNewLine & rsTmp!����
        End If
        rsTmp.MoveNext
    Loop
    
    GetPorcTxtByName = strResult
    Exit Function
errH:
    MsgBox "��ȡ�����ı����ִ���." & vbNewLine & err.Description, , "����"
End Function


Public Function CheckProcManage() As Boolean
    '����:����û��䶯���̹���ģ���Ƿ��Ѿ�����
    '˵��:�û��䶯������������ǰʹ�õĹ���,����ͨ���ű����ύ,����Ҫ�ڳ����н����жϺ���ʱ���\�޸�
    '��Ҫ���\�޸ĵĲ���:1.��������ģ������;2.zlProcedure\zlProcedureText��ṹ���޸�
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    On Error Resume Next
    
    '1.���ģ��
    strSQL = "Select 1 From zlSvrTools Where �ϼ� = '01' And ���� In ('�䶯������������','�䶯�����ճ�����')"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "���䶯����ģ��")
    
    If rsTmp.RecordCount <> 2 Then
        gcnOracle.Execute "Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0106','01','�䶯������������','B',Null,16)"
        gcnOracle.Execute "Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��,����) Values('0107','01','�䶯�����ճ�����','U',Null,17)"
    End If
    
    '2.�޸ĽṹzlProcedure�������������ֶ�  ����ǰ�汾\������汾\ϵͳ���
    err.Clear
    strSQL = "Select ����ǰ�汾,������汾,ϵͳ��� From zlTools.zlProcedure where 1=0"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "���䶯���̽ṹ")
    
    '������ִ���,������ֶ�
    If err.Number <> 0 Then
        
        If gcnTools Is Nothing Then
            Set gcnTools = GetConnection("ZLTOOLS")
        End If
        
        gcnTools.Execute "Alter Table Zltools.Zlprocedure Add ����ǰ�汾 Varchar2(50)"
        gcnTools.Execute "Alter Table Zltools.Zlprocedure Add ������汾 Varchar2(50)"
        gcnTools.Execute "Alter Table Zltools.zlProcedure Add ϵͳ��� Number(5)"
        gcnTools.Execute "Alter Table Zltools.Zlprocedure Modify ˵�� Varchar2(2000)"
    End If
    
    CheckProcManage = True
End Function

Public Function ConvertStr(ByVal strSource As String) As String
    '����:ȥ���ַ����Ŀո�\���з�,��ת��Ϊ��д
    
    strSource = UCase(strSource)
    strSource = Replace(strSource, " ", "")
    strSource = Replace(strSource, vbNewLine, "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    strSource = Replace(strSource, vbTab, "")
    strSource = Replace(strSource, vbBack, "")
    ConvertStr = strSource
End Function

Public Function GetSqlColor() As String
    '��������:��ȡ�﷨�ؼ���SQL�﷨������ʾ����
    '��ȡ��ֱ�ӽ��﷨�ؼ���SyntaxScheme������Ϊ����ֵ����
    Dim strColor As String, strPath As String
    
    If Not gblnInIDE Then '���Ӷ໷��֧��
        strPath = App.Path & "\PUBLIC\_sql.schclass"
    Else
        strPath = gobjFSO.GetParentFolderName(GetSetting("ZLSOFT", "����ȫ��", "����·��")) & "\PUBLIC\_sql.schclass"
    End If
    If Not gobjFSO.FileExists(strPath) Then
        strPath = "C:\Appsoft\PUBLIC\_sql.schclass"
    End If
    
    If gobjFSO.FileExists(strPath) Then
        strColor = ReadFileToString(strPath)
    End If
    GetSqlColor = strColor
End Function
