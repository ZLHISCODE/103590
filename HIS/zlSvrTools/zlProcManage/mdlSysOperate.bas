Attribute VB_Name = "mdlSysOperate"
Option Explicit
'�ļ�����,������˳�����ļ�ִ��˳����ͬ
'X.X.X����Ϊ4λ�汾��X.X.X.X,��ʱΪ����SP�ű���
Public Enum FileType
    'FT_Before �ű���FT_DBA�ű�ִ��ִ��˳����Ի���
    FT_DBA = 0 '��ҪDBA�û�ִ�еĽű�(System�û�):ZLUPgradeX.X.X_DBA.sql,ZL*_X.X.X_DBA.sql
    FT_Before = 1 '��ǰִ�нű���ZLUPgradeX.X.X_Before.sql.sql(�����ߣ�,ZL*_X.X.X_History_Before.sql (Ӧ��ϵͳ��ʷ��)ZL*_X.X.X_Before.sql(Ӧ��ϵͳ���߿�) *����ϵͳ��\100
    FT_Standard = 2 '��ͨ�����ű���ZLUPgradeX.X.X.sql,ZLUPgradeX.X.X(����).sql,ZL*_X.X.X.sql ,ZL*_X.X.X(����).sql,ZL*_X.X.X_History.sql
    FT_Optional = 3 '��ѡִ�нű�:ZLUPgradeX.X.X_Optional.sql,ZL*_X.X.X_Optional.sql��ZL*_X.X.X__HISTORY_Optional.sql
    FT_Deferred = 4 '�ӳ�ִ�нű�:ZL*_X.X.X_Deferred.sql,ZL*_X.X.X__HISTORY_DEFERRED
End Enum
'�ļ�����ϵͳ
Public Enum SysType
    ST_Tools = 0 '�����߽ű�,�����ļ����ͣ�FT_Before,FT_DBA,FT_Standard,FT_Optional
    ST_App = 1 'Ӧ��ϵͳ���߿�,�����ļ����ͣ�FT_Before,FT_DBA,FT_Standard,FT_Optional��FT_Deferred
    ST_History = 2 'Ӧ��ϵͳ��ʷ�⣬�����ļ����ͣ�FT_Before,FT_Standard,FT_Deferred��FT_Optional
End Enum
'�汾����
Public Enum VersionType
    VT_Normal = 0 '�����汾
    VT_Supple = 1 '���䷢���汾����һ����汾������ǰһ���汾�·�����SP���ǲ���汾
End Enum

Public Enum UserCheckType
    UCT_ZLTOOLS = 0 '�������û���֤
    UCT_DBAUser = 1 'DBA�û���֤
    '��ǰ��������Ϊ1�����ڵ���Ϊ2����Ҫ���������⼸�����Ͷ���ͨ��ֱ�ӵ��ô�����ʹ�õ�
    UCT_CurZLBAK = 2 '��ǰ��ʷ����֤
    UCT_NormalUser = 3 '��ͨ�û���֤
    UCT_SysOwner = 4 '����Ա��¼��֤
    UCT_RACInsUser = 5 'RACʵ���û���֤
End Enum

Public gcllMustObj As Collection '��Ҫ������
Public gobjLog As TextStream
Private mstrStSysOwner As String '��׼��������
Public Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'���ܣ���ָ��INI�����ļ������ݶ�ȡ����¼����
'���أ�Nothing�����"��Ŀ,����"�ļ�¼��,����ͬһ��Ŀ�����ж�������
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As Scripting.TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.Fields.Append "��Ŀ", adVarChar, 100
    rsTmp.Fields.Append "����", adVarChar, 4000, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        If Left(Trim(strLine), 1) = "[" And InStr(strLine, "]") > InStr(strLine, "[") Then
            If strItem <> "" And strText = "" Then
                rsTmp.AddNew
                rsTmp!��Ŀ = strItem
                rsTmp!���� = Null
                rsTmp.Update
            End If
            strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
            strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))

            If strItem <> "" And strText <> "" Then
                rsTmp.AddNew
                rsTmp!��Ŀ = strItem
                rsTmp!���� = strText
                rsTmp.Update
            End If
        ElseIf Trim(strLine) <> "" And strItem <> "" Then
            strText = Trim(strLine)
            rsTmp.AddNew
            rsTmp!��Ŀ = strItem
            rsTmp!���� = strText
            rsTmp.Update
        End If
    Loop
    
    If strItem <> "" And strText = "" Then
        rsTmp.AddNew
        rsTmp!��Ŀ = strItem
        rsTmp!���� = Null
        rsTmp.Update
    End If
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function

Public Function CheckINIValid(rsINI As ADODB.Recordset, ByVal strItem As String) As Boolean
'���ܣ�����Ӧ�������ļ���ʽ�Ƿ���ȷ
'������rsINI=��������ļ����ݵļ�¼��������"��Ŀ,����"�ֶ�
'      strItem=�����ļ��б���Ҫ�������ݵ���Ŀ��,��"��Ŀ1|��Ŀ2|..."
    Dim arrItem As Variant, i As Long
    
    arrItem = Split(strItem, "|")
    For i = 0 To UBound(arrItem)
        rsINI.Filter = "��Ŀ='" & arrItem(i) & "'"
        If rsINI.EOF Then Exit Function
        If rsINI!���� & "" = "" Then Exit Function
        If arrItem(i) Like "*�汾��" Then
            If Not IsVerSion(rsINI!����) Then Exit Function
        End If
    Next
    CheckINIValid = True
End Function

Public Function SplitLine(ByVal strSQL As String) As Variant
'���ܣ���SQL���л��в�֣�ͬʱ��¼���з�
    Dim arrLine As Variant, arrReturn() As Variant
    Dim i As Long, j As Long, lngStart As Long, lngEx As Long, lngCur As Long
    Dim strTmp As String
    arrReturn = Array()
    If strSQL = "" Then SplitLine = arrReturn: Exit Function
    arrLine = Split(Replace(Replace(strSQL, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    ReDim Preserve arrReturn(UBound(arrLine) * 2)
    lngStart = 1
    For i = LBound(arrLine) To UBound(arrLine)
        If i <> 0 Then
            strTmp = Mid(strSQL, lngStart, 2)
            If strTmp = vbCrLf Then
                arrReturn(i * 2 - 1) = vbCrLf
                lngStart = lngStart + 2
            Else
                arrReturn(i * 2 - 1) = Mid(strSQL, lngStart, 1)
                lngStart = lngStart + 1
            End If
        End If
        arrReturn(i * 2) = arrLine(i)
        lngStart = lngStart + Len(arrLine(i))
    Next
    SplitLine = arrReturn
End Function

Public Function TrimCommentLossless(ByVal strSQL As String) As String
'���ܣ�����ȥ��ע�ͣ���TrimComment�Ƚϣ����㷨��������ʵ���ݡ�
    Dim arrLine As Variant, arrTmp As Variant
    Dim i As Long, j As Long
    Dim blnStr As Boolean, blnMultiCom As Boolean
    Dim lngPos1 As Long, lngPos2 As Long, lngPos3 As Long
    Dim blnAddLine As Boolean
    Dim strTmp As String, strFMT As String
    
    On Error GoTo errH
    'ȥ������ע�͡�
    arrTmp = Split(strSQL, "'")
    strFMT = "": blnStr = False: blnMultiCom = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            arrLine = SplitLine(arrTmp(i))
            blnAddLine = True
            For j = LBound(arrLine) To UBound(arrLine) Step 2
                strTmp = arrLine(j)
                blnAddLine = j <> UBound(arrLine)
                If blnMultiCom Then '�Ѿ����ڶ���ע�ͷ�Χ�������Ȳ��ҽ�����
                    lngPos2 = InStr(strTmp, "*/")
                    If lngPos2 > 0 Then
                        strTmp = Mid(strTmp, lngPos2 + 2)
                        blnMultiCom = False
                    Else
                        strTmp = "": blnAddLine = False
                    End If
                End If
                If Not blnMultiCom Then '���/* -- */ ��/*   */--����
                    lngPos2 = InStr(strTmp, "/*")
                    lngPos1 = InStr(strTmp, "--")
                    'ȥ����Ч�Ķ���ע������'/* --*/ ,/* */ ����� --/* */
                    '1������--,����--�ڶ��п�ʼ��֮��
                    '2��������--�����ڶ��п�ʼ��
                    Do While Not blnMultiCom And (lngPos2 > 0 And lngPos2 < lngPos1 Or lngPos1 = 0 And lngPos2 > 0)
                        lngPos3 = InStr(lngPos2, strTmp, "*/")
                        If lngPos3 > 0 Then
                            strTmp = Left(strTmp, lngPos2 - 1) & Mid(strTmp, lngPos3 + 2)
                        Else
                            strTmp = Left(strTmp, lngPos2 - 1)
                            blnMultiCom = True
                        End If
                        lngPos2 = InStr(strTmp, "/*")
                        lngPos1 = InStr(strTmp, "--")
                    Loop
                End If
                'ע���еĿ��У���������
                If blnAddLine Then
                    strFMT = strFMT & strTmp & arrLine(j + 1)
                Else
                    strFMT = strFMT & strTmp
                End If
            Next
        Else
            strTmp = ""
            '��� "'B''C''D'"�����ַ�������ʶ��
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '���ڿմ�����Ϊ�������ַ�
                        strTmp = strTmp & "''"
                    Else '�����ڣ���ô�Ϊ�ַ������һ��
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            strFMT = strFMT & "'" & strTmp & "'"
        End If
        If Not blnMultiCom Then '�Ƕ���ע�ͣ�������ַ����߽�
            blnStr = Not blnStr '��ʼ�����ַ����߽�
        End If
    Next
    
    'ȥ������ע��
    arrTmp = Split(strFMT, "'")
    strFMT = "": blnStr = False: blnMultiCom = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            arrLine = SplitLine(arrTmp(i))
'            blnMultiCom = False
            For j = LBound(arrLine) To UBound(arrLine) Step 2
                strTmp = arrLine(j)
                If j = LBound(arrLine) And blnMultiCom Then
                    blnMultiCom = UBound(arrLine) = 0
                Else
                    blnAddLine = j <> UBound(arrLine)
                    lngPos1 = InStr(strTmp, "--")
                    If lngPos1 > 0 Then
                        strTmp = Left(strTmp, lngPos1 - 1)
                        blnMultiCom = UBound(arrLine) = j
                    End If
                    If blnAddLine Then
                        strFMT = strFMT & strTmp & arrLine(j + 1)
                    Else
                        strFMT = strFMT & strTmp
                    End If
                End If
            Next
        Else
            strTmp = ""
            '��� "'B''C''D'"�����ַ�������ʶ��
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '���ڿմ�����Ϊ�������ַ�
                        strTmp = strTmp & "''"
                    Else '�����ڣ���ô�Ϊ�ַ������һ��
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            strFMT = strFMT & "'" & strTmp & "'"
        End If
        If Not blnMultiCom Then
            blnStr = Not blnStr '��ʼ�����ַ����߽�
        End If
    Next
    TrimCommentLossless = strFMT
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function GetFMTSQLStr(ByVal strSQL As String, ByRef cllStrs As Collection) As String
'���ܣ���ȡSQL�е��ַ���������ռλ��ռλ�����ظ�ʽ����SQL
    Dim arrTmp As Variant
    Dim i As Long, j As Long, intIndex As Integer
    Dim strFMT As String, strTmp As String
    Dim blnStr As Boolean
    
    Set cllStrs = New Collection
    arrTmp = Split(strSQL, "'")
    strFMT = "": blnStr = False
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Not blnStr Then
            strFMT = strFMT & arrTmp(i)
        Else
            strTmp = ""
            '��� "'B''C''D'"�����ַ�������ʶ��
            For j = i To UBound(arrTmp) Step 2
                strTmp = strTmp & arrTmp(j)
                If j + 1 <= UBound(arrTmp) Then
                    If arrTmp(j + 1) = "" Then '���ڿմ�����Ϊ�������ַ�
                        strTmp = strTmp & "''"
                    Else '�����ڣ���ô�Ϊ�ַ������һ��
                        i = j: Exit For
                    End If
                Else
                    i = j: Exit For
                End If
            Next
            intIndex = intIndex + 1
            '����ַ���
            strFMT = strFMT & "[S" & intIndex & "]"
            cllStrs.Add strTmp, "S" & intIndex
        End If
        blnStr = Not blnStr '��ʼ�����ַ����߽�
    Next
    arrTmp = SplitLine(strFMT)
    strFMT = "": blnStr = False
    For i = LBound(arrTmp) To UBound(arrTmp) Step 2
        strTmp = TrimEx(arrTmp(i))
        If strTmp <> "" Then
            If Right(strTmp, 1) = ";" And i <> UBound(arrTmp) Then
                strFMT = strFMT & " " & strTmp & vbCrLf
            Else
                strFMT = strFMT & " " & strTmp
            End If
        End If
    Next
    'ȥ���������еĿո�
    arrTmp = SplitLine(strFMT)
    strFMT = ""
    For i = LBound(arrTmp) To UBound(arrTmp) Step 2
        strTmp = TrimEx(TrimBesideOperator(arrTmp(i)))
        If strTmp <> "" Then
            If Right(strTmp, 1) = ";" And i <> UBound(arrTmp) Then
                strFMT = strFMT & " " & strTmp & vbCrLf
            Else
                strFMT = strFMT & " " & strTmp
            End If
        End If
    Next
    GetFMTSQLStr = UCase(strFMT)
End Function

Public Function TrimBesideOperator(ByVal strText As String) As String
'���ܣ�ȥ��TAB�ַ������߿ո񣬻س������ֻ�ɵ��ո�ָ���
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim i As Long
    
    strText = Replace(Replace(strText, " :", ":"), ": ", ":")
    strText = Replace(Replace(strText, " =", "="), "= ", "=")
    strText = Replace(Replace(strText, " .", "."), ". ", ".")
    strText = Replace(Replace(strText, " )", ")"), ") ", ")")
    strText = Replace(Replace(strText, " (", "("), "( ", "(")
    strText = Replace(Replace(strText, " %", "("), "% ", "%")
    strText = Replace(Replace(strText, " \", "\"), "\ ", "\")
    TrimBesideOperator = strText
End Function

Public Function GetInfoInsideBracket(ByVal strInfo As String, Optional ByVal strLeftChar As String, Optional ByVal strRightChar As String) As String
'����������ȡ����
'����������������ݣ�ֻȡ�����
    Dim lngSart As Long, lngEnd As Long
    If strRightChar = "" Then strRightChar = ")"
    If strLeftChar = "" Then strLeftChar = "("
    lngEnd = InStrRev(strInfo, strRightChar) - Len(strRightChar) + 1 '��ͷ����β�����Բ���һ
    lngSart = InStr(strInfo, strLeftChar) + Len(strLeftChar)
    If lngEnd < lngSart Then
        GetInfoInsideBracket = ""
    Else
        GetInfoInsideBracket = Mid(strInfo, lngSart, lngEnd - lngSart)
    End If
End Function

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
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

Public Function TrimComment(ByVal strSQL As String) As String
'���ܣ�ȥ��д�ڵ���strSQL�������"--"ע��
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim blnStr As Boolean
    Dim i As Long, K As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                K = i: Exit For
            End If
        Next
        If K > 0 Then strSQL = RTrim(Left(strSQL, K - 1))
    End If
    TrimComment = strSQL
End Function

Public Function SplitSQL(ByVal strSQL As String) As String
'���ܣ�ȡ";"��βǰ��ĵ�SQL���,����";"�ź���"--"ע�͡�
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim i As Long, K As Long
    
    '��ȥ��ע�Ͳ���
    strSQL = TrimComment(strSQL)
    
    For i = Len(strSQL) To 1 Step -1
        If Mid(strSQL, i, 1) = ";" Then
            K = i: Exit For
        End If
    Next
    If K > 0 Then strSQL = Left(strSQL, K - 1)
    
    SplitSQL = strSQL
End Function

Public Function RemoveMark(ByVal strText As String) As String
'���ܣ�ȥ��һ�������е�ǰ��"--"ע�ͱ��
    Dim arrText As Variant, strTemp As String, i As Long
    
    arrText = Split(strText, vbCrLf)
    
    strText = ""
    For i = 0 To UBound(arrText)
        strTemp = arrText(i)
        If Left(strTemp, 2) = "--" And Replace(strTemp, "-", "") <> "" Then
            strText = strText & vbCrLf & Mid(strTemp, 3)
        End If
    Next
    RemoveMark = Mid(strText, 3)
End Function


Public Function CheckInitFile(ByVal lngSys As Long, ByVal strFile As String, Optional ByVal blnOnlyCheck As Boolean, Optional ByRef rsReturnINI As ADODB.Recordset, Optional ByVal blnUpgradeCheck As Boolean = True) As Boolean
'������blnUpgradeCheck=�����Ǩ����ļ�
   Dim strSysPath As String, strTmp As String
   Dim rsINI As ADODB.Recordset
   If Not gobjFile.FileExists(strFile) Then
        If Not blnOnlyCheck Then MsgBox "��װ�����ļ�""" & strFile & """�����ڡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    If UCase(gobjFile.GetFileName(strFile)) <> IIf(lngSys = 0, "ZLSERVER.SQL", "ZLSETUP.INI") Then
        If Not blnOnlyCheck Then MsgBox "��װ�����ļ�������ȷ��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If lngSys = 0 Then '������
        '��������������麯���ļ��Ƿ���ڡ�
        If blnUpgradeCheck Then
            strSysPath = gobjFile.GetParentFolderName(strFile)
            strTmp = strSysPath & "\zlUpgradeCheck.sql"
            If Not gobjFile.FileExists(strTmp) Then
                If Not blnOnlyCheck Then MsgBox "��������������ļ�""" & strTmp & """�����ڡ�", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    Else 'Ӧ��ϵͳ
        Set rsINI = ReadINIToRec(strFile)
        If Not CheckINIValid(rsINI, "ϵͳ��|�汾��|��ռ�|�����߰汾��") Then
            If Not blnOnlyCheck Then MsgBox "��װ�����ļ���ʽ����ȷ��", vbExclamation, gstrSysName
            Exit Function
        End If
        '�����ļ�ϵͳ�Ų�ƥ��
        rsINI.Filter = "��Ŀ='ϵͳ��'"
        If Val(rsINI!����) <> lngSys \ 100 Then
            If Not blnOnlyCheck Then MsgBox "��ѡ�����ļ����Ǳ�ϵͳ�İ�װ�����ļ���", vbExclamation, gstrSysName
            Exit Function
        End If
        strSysPath = gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(strFile))
        'ϵͳ��ǨĿ¼���
        If Not gobjFile.FolderExists(strSysPath & "\�����ű�") Then
            If Not blnOnlyCheck Then MsgBox "ϵͳ��ǨĿ¼""" & strSysPath & "\�����ű�""�����ڡ�", vbExclamation, gstrSysName
            Exit Function
        End If
        If blnUpgradeCheck Then
            '���Ӧ��ϵͳ������麯���ļ��Ƿ���ڡ�
            strTmp = strSysPath & "\�����ű�\zl" & lngSys \ 100 & "_UpgradeCheck.sql"
            If Not gobjFile.FileExists(strTmp) Then
                If Not blnOnlyCheck Then MsgBox "ϵͳ��������ļ�""" & strTmp & """�����ڡ�", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
        '��Ӧ�İ�װ�ű��ļ��Ƿ����,����Ҫ��飬��Ϊ�Ѿ�ȡ���˿�ѡ�ű�ִ��
    End If
    Set rsReturnINI = rsINI
    CheckInitFile = True
End Function

Public Function GetUpgradeFiles(ByVal rsUpgradeFiles As ADODB.Recordset, ByVal lngSys As Long, ByVal strCurVer As String, ByVal strIniPath As String, _
                                                        Optional ByVal strNoramlBreak As String, Optional ByVal strBeforeBreak As String, _
                                                        Optional ByRef strMaxVer As String, Optional ByRef strCurMaxVer As String, Optional ByVal strBakDB As String, _
                                                        Optional ByVal blnReadByMax As Boolean, Optional ByVal blnDeleteSpfile As Boolean = True) As ADODB.Recordset
'���ܣ���ȡ����Ҫִ�е��ļ�
'������rsUpgradeFiles=�����ļ���¼���������Ƕ��ϵͳ�������ļ���¼��
'          lngSys=ϵͳ��,=-1��ʾֻ��ʼ����¼��
'          strIniPath=��װ�����ļ�
'          strBreakVers=��Ǩ�����ļ��Ķϵ�汾
'          strBakDB=��ʷ������
'          strMaxVer=���İ汾
'          strCurMaxVer=������Ǩ��Ŀ��汾
'          blnReadByMax=�������汾strMaxVer��ȡ�ű�����Ҫ����ϵͳ��װʱ�����߰汾�ϵ͹����ߵ�������ʱʹ�ã�
'                                   �ò���ΪTrueʱ��������жϵ㴦�����������Ӧ��ϵͳ�ű�����һ��
'          blnDeleteSpfile=�Ƿ�ɾ������SP�ļ�,Ture-ֻ��ȡ������Ŀ��汾������SP�ű� False-��ȡ���а�װ��������SP�ű�
'����:�����ļ���¼
'        strMaxVer=����Ŀ��汾,����ǰ�ű�������Ǩ��������汾
'        strCurMaxVer=������Ǩ��Ŀ��汾��ϵͳ��Ǩ��������ĳЩ�汾����������Ǩ��������Ҫ�ֶ����Ǩ���ܵ�����Ŀ��汾��
'                               û�в���������Ǩ�İ汾ʱ,�ð汾��strMaxVer��ͬ
'˵����
'        strBakDB="":��ȡ���нű�����ʱ���²�������
'                            strNoramlBreak�����߿⣨lngSys=0��Ϊ�����ߣ�����������ֹ��Ϣ
'                            strBeforeBreak:���߿⣨lngSys=0��Ϊ�����ߣ���ǰ������ֹ��Ϣ
'                            strMaxVer:������������Ǩ������Ŀ��汾
'                            strCurMaxVer:������������Ǩ�ı���Ŀ��汾
'                            ���ص��ļ���¼���и��ڱ�����ǨĿ��汾�Ľű�ȫ���޳���
'        strBakDB<>"":��ȡ����strCurVer���Ҳ�����strMaxVer�Ľű�������������ʷ��Ľű��ļ���¼����
'                             ����ʷ��ǵ�����Ǩʱ�����ɵĽű��ļ���¼����Ҫ��������Ӧ��ϵͳ��ǰ�汾��Ӧ��ϵͳ����Ŀ��汾֮�����ʷ��ű�
'                             ��ʱ���²������壺
'                            strNoramlBreak����ʷ�ⳣ��������ֹ��Ϣ
'                            strBeforeBreak:��ʷ����ǰ������ֹ��Ϣ
'                            strMaxVer:���߿�ĵ�ǰ�汾
    Dim rsCurFiles As ADODB.Recordset, arrFields As Variant, blnNew As Boolean
    Dim strCurPriFull As String, strCurFull As String, strMaxFull As String, strMaxPriFull As String
    Dim cllFolder As New Collection, objFolder As Folder, objFile As File
    Dim strBreak As String, strTmp As String, arrTmp As Variant, strFilter As String
    Dim strFileVer As String, stFile As SysType, ftFile As FileType, vtFile As VersionType, strSetupVer As String, blnSpecial As Boolean
    Dim strFileNameRule As String, stJudge As SysType
    Dim cllSuppleVers As New Collection, Item As Variant
    Dim i As Long
    Dim strFirstBreak As String, strSecdBreak As String
    Dim strBaseSupple As String
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strBanner As String, intSpVer As Integer
    
    On Error GoTo errH
    
    strCurPriFull = VerFull(GetPrimaryVer(strCurVer))
    strCurFull = VerFull(strCurVer)
    strMaxFull = VerFull(strMaxVer, True) '�մ�������9999.9999.9999.9999
    strMaxPriFull = VerFull(GetPrimaryVer(strMaxFull)) '��ֹ�մ�����ʧ�ܣ���˲���strMaxVer����
    If rsUpgradeFiles Is Nothing Then
        blnNew = True
    ElseIf rsUpgradeFiles.State = adStateClosed Then
        blnNew = True
    End If
    
    If blnNew Or lngSys = -1 Then
        '���ð汾:����ǰִ�нű�Ϊ���Ҫ��汾����ӦӦ��ϵͳ���߿���ͨ�����ű�Ϊ��Ӧ�����߽ű�
        Set rsUpgradeFiles = CopyNewRec(Nothing, True, , _
                                                                Array("ϵͳ���", adInteger, 5, Empty, "������", adVarChar, 100, Empty, "SysType", adInteger, 1, Empty, _
                                                                        "FileName", adVarChar, 50, Empty, "FilePath", adVarChar, 1000, Empty, "FileType", adInteger, 1, Empty, _
                                                                        "SPVer", adVarChar, 20, Empty, "FullSPVer", adVarChar, 20, Empty, "VerType", adInteger, 1, Empty, _
                                                                        "Optional", adVarChar, 2000, Empty, "AbortLine", adInteger, 10, Empty, "Special", adInteger, 1, Empty, _
                                                                        "���ð汾", adVarChar, 20, Empty, "�ϵ�", adInteger, 1, Empty))
    End If
    If lngSys = -1 Then Set GetUpgradeFiles = rsUpgradeFiles: Exit Function
    '��ȡ��ǰϵͳ�Ľű�
    rsUpgradeFiles.Filter = "ϵͳ���=" & lngSys & IIf(strBakDB <> "", " And ������='" & UCase(strBakDB) & "'", "")
    '�ű��Ѿ����ڣ��������¶�ȡ��
    '��ʷ���ȡ���������汾��Ϊ�ա���Ϊ��ʷ�ⵥ����Ǩ��Ŀ��汾Ϊ���߿⵱ǰ�汾���ǵ�������ʱ�����߿⵱ǰ�汾֮�ϵ���ʷ�ű��Ѿ���ȡ
    If Not rsUpgradeFiles.EOF Or strBakDB <> "" And strMaxVer = "" Then Set GetUpgradeFiles = rsUpgradeFiles: Exit Function
    Set rsCurFiles = CopyNewRec(rsUpgradeFiles, strBakDB = "")
    '////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////          1����Ǩ�ļ���ȡ            ///////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    '��ȡ��Ҫ�Ѽ��ű����ļ���
    If lngSys = 0 Then
        cllFolder.Add gobjFile.GetFile(strIniPath).ParentFolder
        strFileNameRule = "ZLUPGRADE*.*.*.SQL"
    Else
        strFileNameRule = "ZL" & lngSys \ 100 & "_*.*.*.SQL"
        For Each objFolder In gobjFile.GetFolder(gobjFile.GetParentFolderName(gobjFile.GetParentFolderName(strIniPath)) & "\�����ű�\").SubFolders
            If IsVerSion(objFolder.Name) And objFolder.Name Like "*.*.0" Then
                If VerFull(objFolder.Name) >= strCurPriFull And VerFull(objFolder.Name) <= strMaxPriFull Then
                    cllFolder.Add objFolder
                End If
            End If
        Next
    End If
    arrFields = Array("ϵͳ���", "SysType", "FileName", "FilePath", "FileType", "SPVer", "FullSPVer", "VerType", "Special", "���ð汾")
    '����,��ȡ�ļ�
    For Each objFolder In cllFolder
        If lngSys <> 0 And strBakDB = "" Then '��ȡzlUpgrade.ini
            '��ȡ��Ч�Ķϵ�汾
            strTmp = GetUpgradeIniBreak(objFolder.Path & "\zlUpgrade.ini", IIf(VerFull(objFolder.Name) >= strCurPriFull, strCurVer, objFolder.Name), GetPrimaryVer(objFolder.Name, True))
            If strTmp <> "" Then
                strBreak = strBreak & "," & strTmp
            End If
        End If
        '��ȡ�ļ�
        For Each objFile In objFolder.Files
            If UCase(objFile.Name) Like strFileNameRule Then '�����ļ��Ĺ���ĲŽ������ƽ���
                If AnalysisFileName(objFile.Name, lngSys, strFileVer, ftFile, stFile, vtFile, blnSpecial) Then
                    If VerFull(strFileVer) > strCurFull And VerFull(strFileVer) <= strMaxFull Then
                        If vtFile = VT_Supple Then
                            On Error Resume Next
                            'ȷ�ϸô�汾�Ѿ���ǵĲ���汾
                            strBaseSupple = cllSuppleVers("K_" & GetPrimaryVer(strFileVer))
                            If Err.Number <> 0 Then
                                Err.Clear
                                cllSuppleVers.Add strFileVer, "K_" & GetPrimaryVer(strFileVer)
                            '�Ѿ���ǵĲ���汾С�ڵ�ǰ�汾���򽲱���޸�Ϊ��ǰ�汾
                            ElseIf VerFull(strBaseSupple) > VerFull(strFileVer) Then
                                cllSuppleVers.Remove "K_" & GetPrimaryVer(strFileVer)
                                cllSuppleVers.Add strFileVer, "K_" & GetPrimaryVer(strFileVer)
                            End If
                            On Error GoTo errH
                        End If
                        '��ȡ���ð汾
                        If ftFile = FT_Before Or ftFile = FT_Standard And stFile = ST_App And VerFull(strFileVer) > VerFull("10.32.0") Then
                            arrTmp = Split(GetUpgradeCtrolInfo(objFile.Path, ftFile = FT_Before) & "|", "|")
                            strSetupVer = VerFull(arrTmp(IIf(ftFile = FT_Before, 0, 1))) '����Ϊ��׼�汾������Ƚ�;    ��ǰִ�з��أ����Ҫ��汾�����������ű����أ���������|��Ӧ�����߰汾
                            '10.34.0֮�󣬹����ߣ�Ӧ��ϵͳ�汾�Ѿ�һһ��Ӧ����û�нű��İ汾�ÿ��ļ�����
                            If ftFile = FT_Standard Then
                                 If VerFull(strFileVer) >= VerFull("10.34.0") Then
                                    strSetupVer = VerFull(strFileVer) '����Ϊ��׼�汾������Ƚ�
                                ElseIf strSetupVer = VerFull("0") Then  '��ȡӦ�ö�Ӧ���߰汾ʧ�ܣ����Զ�����һ��
                                    strSetupVer = VerFull(GetContractVersion(strFileVer, True))
                                End If
                            End If
                            If Val(arrTmp(0)) <> 1 And ftFile = FT_Standard And strBakDB = "" Then strBreak = strBreak & "," & strFileVer
                        Else
                            strSetupVer = ""
                        End If
                        
                        rsCurFiles.AddNew arrFields, Array(lngSys, stFile, objFile.Name, objFile.Path, ftFile, strFileVer, VerFull(strFileVer), vtFile, IIf(blnSpecial, 1, 0), strSetupVer)
                    End If
                End If
            End If
        Next
    Next
    '////////////////////////////////////////////////////////////////////////////////////
    '////////////////////   2.�ϴ���Ǩ��Ϣ���޳�������汾�ϵ���  ///////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    '��ǲ���汾
    For Each Item In cllSuppleVers
        '���ڸô�汾����С�Ĳ���汾����С����һ���汾
        Call RecUpdate(rsCurFiles, "FullSPVer>='" & VerFull(Item) & "' And FullSPVer<'" & VerFull(GetPrimaryVer(Item, True)) & "'", "VerType", VT_Supple)
    Next
    stJudge = IIf(lngSys = 0, ST_Tools, IIf(strBakDB = "", ST_App, ST_History))
    strFilter = "SysType=" & stJudge & " And FileType<>" & FT_Deferred
    '�޳���ǰ��ֹ���֮ǰ���ļ�
    arrTmp = Split(strBeforeBreak & "||", "|")
    'û����ֹ�ļ�����С�ڵ�����ֹ�汾����ǰִ�нű���Ҫɾ��������ֻɾ��С����ֹ�汾����ǰ�ű�
    Call RecDelete(rsCurFiles, strFilter & " And FileType=" & FT_Before & " And FullSPVer<" & IIf(arrTmp(1) = "", "=", "") & "'" & VerFull(arrTmp(0)) & "'")
    If arrTmp(1) <> "" Then '����ֹ�ļ�����¼��ֹ��
        Call RecUpdate(rsCurFiles, strFilter & "And FileType=" & FT_Before & " And SPVer='" & arrTmp(0) & "'", "AbortLine", Val(arrTmp(2)))
    End If
    arrTmp = Split(strNoramlBreak & "||", "|")
    '�޳�������ֹ���֮ǰ���ļ�
    Call RecDelete(rsCurFiles, strFilter & " And FullSPVer<" & IIf(arrTmp(1) = "", "=", "") & "'" & VerFull(arrTmp(0)) & "'")
    If arrTmp(1) <> "" Then '����ֹ�ļ�
        'ɾ����ֹ��ֹ�汾��ִ��˳������ֹ�ļ�֮ǰ���ļ�
        Call RecDelete(rsCurFiles, strFilter & " And SPVer='" & arrTmp(0) & "' And FileType<" & Val(arrTmp(1)))
        '��¼��ֹ��
        Call RecUpdate(rsCurFiles, strFilter & " And SPVer='" & arrTmp(0) & "' And FileType=" & Val(arrTmp(1)), "AbortLine", Val(arrTmp(2)))
    End If
    '����������Ǩ�汾�ı��
    strBreak = Mid(strBreak, 2): arrTmp = Split(strBreak, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        Call RecUpdate(rsCurFiles, "SPVer='" & arrTmp(i) & "'", "�ϵ�", 1)
    Next
    
    '�޳�����汾�����汾���򣬵�һ���ǲ���汾֮ǰ�����в���汾ȫ��ɾ����
    rsCurFiles.Filter = "VerType=" & VT_Normal: rsCurFiles.Sort = "FullSPVer Desc"
    If Not rsCurFiles.EOF Then Call RecDelete(rsCurFiles, "VerType=" & VT_Supple & " And FullSPVer<'" & rsCurFiles!fullspver & "'")
    
    rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
    If blnDeleteSpfile Then
        '�޳�����SP�ű������汾���򣬵�һ��������SP�汾֮ǰ����������SPȫ��ɾ��
        '�����ж��и����⣬����һ���汾û����ʽ�ű�������������SP�ű�����˲������ִ���
        If Not rsCurFiles.EOF Then
            strTmp = VerFull(VerSpecialNormal(rsCurFiles!SPVer))
            Call RecDelete(rsCurFiles, "Special=1 And FullSPVer<'" & strTmp & "'")
        End If
    Else
        '���Ҫ��������SP�ű�,����Ҫ�жϸ�����SP�ű��Ƿ�װ
        rsCurFiles.Filter = "Special  =1"
        Do While Not rsCurFiles.EOF
            strTmp = Mid(rsCurFiles!SPVer, 1, InStrRev(rsCurFiles!SPVer, ".") - 1) '��¼��ǰС�汾
            If strBanner <> strTmp Then '�����ļ�¼,С�汾��ͬ,��Ҫ���»�ȡ��С�汾���������SP�汾
                strBanner = strTmp
                strSQL = "Select Nvl(Max(Substr(����汾, Instr(����汾, '.', 1, 3) + 1)), 0) ���sp" & vbNewLine & _
                            "From zlUpGrade A" & vbNewLine & _
                            "Where ϵͳ = [1] And ����汾 Like '" & strTmp & ".%'"
                Set rsTmp = OpenSQLRecord(strSQL, "��ȡ���SP�汾", lngSys)
                intSpVer = Val(rsTmp!���sp)
            End If
            
            If Val(Split(rsCurFiles!fullspver, ".")(3)) > intSpVer Then
                rsCurFiles.Delete adAffectCurrent
            End If
            
            rsCurFiles.MoveNext
        Loop
    End If
    '////////////////////////////////////////////////////////////////////////////////////
    '/////////////// 3������Ŀ��汾������Ŀ��汾���Լ���ʷ��ű��Ķ�ȡ ////////////
    '///////////////////////////////////////////////////////////////////////////////////
    If strBakDB = "" Then
        If blnReadByMax Then '�������汾��ȡ
            '��ȡʵ�ʿ��������������汾
            rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
            strCurMaxVer = ""
            If Not rsCurFiles.EOF Then
                strCurMaxVer = rsCurFiles!SPVer & ""
            End If
        Else
            '��ȡ����Ŀ��汾�Լ�����Ŀ��汾
            rsCurFiles.Filter = "": rsCurFiles.Sort = "FullSPVer Desc"
            strMaxVer = "": strCurMaxVer = ""
            If Not rsCurFiles.EOF Then
                strMaxVer = rsCurFiles!SPVer & ""
                rsCurFiles.Filter = "�ϵ�=1": rsCurFiles.Sort = "FullSPVer"
                If Not rsCurFiles.EOF Then
                    strFirstBreak = rsCurFiles!SPVer
                    If rsCurFiles.RecordCount > 1 Then
                        rsCurFiles.MoveNext: strSecdBreak = rsCurFiles!SPVer
                    End If
                    rsCurFiles.Filter = "FullSPVer<'" & VerFull(strFirstBreak) & "'"
                    strCurMaxVer = IIf(rsCurFiles.EOF, strSecdBreak, strFirstBreak)
                End If
            End If
            If strCurMaxVer = "" Then
                strCurMaxVer = strMaxVer
            Else 'ɾ������Ҫ������Ǩ����Ҫִ�еĽű�
                Call RecDelete(rsCurFiles, "FullSPVer>'" & VerFull(strCurMaxVer) & "'")
            End If
        End If
    Else
    '��ȡ��ʷ����Ǩ��¼
        'ɾ��С����ʷ�⵱ǰ�汾�Ľű�����ʷ��汾���ܸ������߿⣬�����Ҫ��������
        Call RecDelete(rsCurFiles, "FullSPVer<='" & VerFull(strCurVer) & "'")
        'ɾ�����߿�ű�
        Call RecDelete(rsCurFiles, "SysType<>" & ST_History)
        '�����ļ���¼����������
        Call RecUpdate(rsCurFiles, "", "������", UCase(strBakDB))
    End If
    '�ϲ���¼���������ζ�ȡ���ļ��ϲ������м�¼����
    rsCurFiles.Filter = ""
    Call RecDataAppend(rsUpgradeFiles, rsCurFiles)
    Set GetUpgradeFiles = rsUpgradeFiles
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Public Function FormatUpgradeBreak(ByVal lngSys As Long, ByVal strResultVer As String, Optional ByVal strUpgradeBreak As String) As String
'���ܣ�������ֹ��Ϣ������ֹ����׼�� ��ʽ���ļ��汾|�ļ�����|�����к�
'������
'     strResultVer:ZLUpgrade�еĽ���汾
'     strUpgradeBreak=��Ǩ��ֹ���
'���أ��ļ��Ĳ���·�����ļ���
    Dim arrTmp As Variant
    Dim lngPos As Long
    Dim strTmp As String
    Dim strFileName As String
    Dim lngAbort As Long
    Dim strFileVer As String '���ļ����϶�ȡ�İ汾��Ϣ
    Dim ftReturn As FileType
    Dim strReturn As String
    
    strReturn = strResultVer & "||"
    If strUpgradeBreak <> "" Then
        '��ʷ�����ֹ������Ϊ�汾��
        If Not IsVerSion(strUpgradeBreak) Then
            strUpgradeBreak = strUpgradeBreak & "||"
            arrTmp = Split(strUpgradeBreak, "|")
            If gobjFile.FileExists(arrTmp(0)) Then
                strFileName = gobjFile.GetFileName(arrTmp(0))
            Else '�����ǲ���汾�Ѿ�ɾ����
                strTmp = StrReverse(arrTmp(0))
                lngPos = InStr(strTmp, "\")
                '��ȡ���һ��\�������
                If lngPos <> 0 Then
                    strFileName = StrReverse(Mid(strTmp, lngPos - 1))
                Else
                    strFileName = ""
                End If
            End If
            lngAbort = Val(arrTmp(1))
            If strFileName <> "" Then
                If AnalysisFileName(strFileName, lngSys, strFileVer, ftReturn) Then
                    strReturn = strFileVer & "|" & ftReturn & "|" & lngAbort
                End If
            End If
        Else '��ʷ����ǰ������ŵ��ǰ汾��
            strReturn = strUpgradeBreak & "||"
        End If
    End If
    FormatUpgradeBreak = strReturn
End Function

Public Function GetUpgradeIniBreak(ByVal strFile As String, Optional ByVal strMinVer As String, Optional ByVal strMaxVer As String)
'���ܣ���ȡ��Ǩ�����ļ��Ķϵ�
'������strFile=��Ǩ�����ļ�·��
'          strMinVer=��Ǩ�����ļ�Ŀ��汾����Сֵ
'          strMaxVer=��Ǩ�����ļ�Ŀ��汾�����ֵ
    Dim rsSub As ADODB.Recordset
    Dim strBreakVer As String
    
    If Not gobjFile.FileExists(strFile) Then Exit Function
    Set rsSub = ReadINIToRec(strFile)
    If rsSub Is Nothing Then Exit Function
    rsSub.Filter = "��Ŀ='��������'" '���������ļ���Ŀ��汾�Ƿ�����������
    If rsSub.EOF Then Exit Function
    If Val(rsSub!���� & "") = 1 Then Exit Function '�����������ô���
    rsSub.Filter = "��Ŀ='Ŀ��汾'" '���������ļ���Ŀ��汾
    If rsSub.EOF Then Exit Function
    strBreakVer = Trim(rsSub!���� & "")
    If Not IsVerSion(strBreakVer) Then Exit Function
    If strMinVer <> "" Then 'С����С�汾����öϵ���Ч
        If VerFull(strBreakVer) <= VerFull(strMinVer) Then Exit Function
    End If
    If strMaxVer <> "" Then '������С�汾����öϵ���Ч
        If VerFull(strBreakVer) > VerFull(strMaxVer) Then Exit Function
    End If
    GetUpgradeIniBreak = strBreakVer
End Function

Public Function GetUpgradeCtrolInfo(ByVal strFile As String, Optional ByVal blnBefore As Boolean) As String
'���ܣ���ȡ�ļ��еĿ�����Ϣ
'      strFile=�����жϵĽű��ļ�·��
'      blnBefore=�ļ��Ƿ�������ִ�нű�
'����: blnBefore=false: ��������|�����߰汾��
'        blnBefore=True: ��Ͱ汾��

    Dim objStream As Scripting.TextStream
    Dim strLine As String, arrFind() As Variant, i As Long, strTmp As String, arrTmp As Variant
    Dim strContinue As String, strToolVer As String, strBreakVer As String, strReqVer As String
    Dim rsSub As ADODB.Recordset
    
    On Error GoTo errH
    
    Set objStream = gobjFile.OpenTextFile(strFile, ForReading)
    If blnBefore Then
        arrFind = Array("[[]��Ͱ汾��[]]")
    Else
        arrFind = Array("[[]��������[]]", "[[]�����߰汾��[]]")
    End If
    Do While Not objStream.AtEndOfStream
        strLine = TrimEx(objStream.ReadLine, True)
        If strLine Like "--" & arrFind(i) & "*" Then
            strTmp = Trim(Mid(strLine, Len("--" & arrFind(i)) - 4 + 1))
            If Not blnBefore Then
                If i = 0 Then
                    strContinue = strTmp
                Else
                    strToolVer = strTmp
                End If
            Else
                strReqVer = strTmp
            End If
        End If
        If i = UBound(arrFind) Then Exit Do
        i = i + 1
    Loop
    objStream.Close
    
    If blnBefore Then
        GetUpgradeCtrolInfo = Trim(strReqVer)
    Else
        If Trim(strContinue) = "" Then strContinue = "1"
        GetUpgradeCtrolInfo = Trim(strContinue) & "|" & Trim(strToolVer)
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
'    Debug.Print err.Source & "\" & Me.name & "\GetCtrolInfo:" & err.Description
End Function


Public Function AnalysisFileName(ByVal strFileName As String, ByVal lngSys As Long, Optional ByRef strVersion As String, Optional ByRef ftReturn As FileType, _
                                                        Optional ByRef stReturn As SysType, Optional ByRef vtReturn As VersionType = VT_Normal, Optional ByRef blnSpecial As Boolean) As Boolean
'����:tͨ���ļ�����ȡ�ļ���Ϣ
'������
'   strFile=������·�����ļ���,����չ��
'   lngSys=ϵͳ��
'����:
'       True=�ɹ���ȡ��False=��ȡʧ�ܣ��ļ�����ϵͳ�����ű���
'       strVerReturn=�ļ��汾
'       ftReturn=�ļ�����
'       stReturn=ϵͳ����
'       vtReturn=�汾����
    Dim strSysString As String, strSuffix As String
    Dim arrVer As Variant
    vtReturn = VT_Normal
    blnSpecial = False
    strVersion = ""
    ftReturn = FT_Before
    stReturn = ST_Tools
    If Not UCase(strFileName) Like "*.SQL" Then Exit Function
    strFileName = UCase(Left(strFileName, Len(strFileName) - 4))
    arrVer = Split(strFileName, ".")
    '�汾�ļ����ļ�������2������(����SP����3����
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    '��ȡ�ű�ϵͳǰ׺
    If arrVer(0) Like "ZLUPGRADE*" Then
        strSysString = "ZLUPGRADE"
        stReturn = ST_Tools
    ElseIf arrVer(0) Like "ZL" & lngSys \ 100 & "_*" Then
        strSysString = "ZL" & lngSys \ 100 & "_"
        stReturn = ST_App
    Else
        Exit Function 'û��ϵͳ��ʶǰ׺������ϵͳ�ű�
    End If
    'ϵͳ��ʶ����������ǰ汾
    arrVer(0) = Mid(arrVer(0), Len(strSysString) + 1) '��ȡ���屾
    arrVer(UBound(arrVer)) = GetPrefixNumber(arrVer(UBound(arrVer)), strSuffix) '��ȡ�μ��汾
    '��ȡ�����屾����汾�Լ��μ��汾����Ϊ���֣����˳�
    If Not IsNumeric(arrVer(0)) Or Not IsNumeric(arrVer(1)) Or Not IsNumeric(arrVer(2)) Or Not IsNumeric(arrVer(UBound(arrVer))) Then Exit Function
    strVersion = arrVer(0) & "." & arrVer(1) & "." & arrVer(2) & IIf(UBound(arrVer) = 2, "", "." & arrVer(UBound(arrVer)))
    If Not IsVerSion(strVersion) Then Exit Function
    '��λ�汾�ž�������SP
    blnSpecial = strVersion Like "*.*.*.*"
    '�汾�����ļ�������Ϣ
    If stReturn = ST_App And strSuffix Like "_HISTORY*" Then
        stReturn = ST_History
        strSuffix = Mid(strSuffix, Len("_HISTORY") + 1)
    End If
    If strSuffix Like "*(����)" Then
        vtReturn = VT_Supple
        strSuffix = Replace(strSuffix, "(����)", "") '��ֹ������Ϣλ�ò��̶�
    End If
    Select Case strSuffix
        Case ""
            ftReturn = FT_Standard
        Case "_DBA"
            If stReturn = ST_History Then Exit Function '��ʷ�ⲻ֧��DBA�ű�
            ftReturn = FT_DBA
        Case "_OPTIONAL"
            ftReturn = FT_Optional
        Case "_BEFORE"
            ftReturn = FT_Before
        Case "_DEFERRED"
            If stReturn = ST_Tools Then Exit Function '�����߲�֧���ӳ�ִ�нű�
            ftReturn = FT_Deferred
        Case Else '������������Χ�ڵģ����ȡʧ��
            Exit Function
    End Select
    AnalysisFileName = True
End Function

Public Function GetPrefixNumber(ByVal strInput As String, Optional ByRef strOther As String) As String
'���ܣ���ȡһ���ַ���������ǰ׺���Լ�ʣ�ಿ��
'������strInput=������ַ���
'          strOther =ȥ������ǰ׺��ʣ�ಿ��
    Dim i As Long
    
    For i = 1 To Len(strInput)
        If Not IsNumeric(Mid(strInput, i, 1)) Then
            Exit For
        End If
    Next
    strOther = Mid(strInput, i)
    GetPrefixNumber = Mid(strInput, 1, i - 1)
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

Public Function VerPAD(ByVal strVer As String) As String
'���ܣ�ʹ�汾�ŵ����汾�������Ϊ4λ����֤���汾��ԭ������������汾�Ŷ���
'������strVer=��ǰ�汾��
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        Exit Function
    End If
    arrVer = Split(strVer & ".", ".")
    VerPAD = RPAD(Lpad(arrVer(0), 2) & "." & arrVer(1) & "." & arrVer(2) & IIf(Val(arrVer(3)) = 0, "", "." & Format(Val(arrVer(3)), "0000")), 20)
End Function

Public Function GetPrimaryVer(ByVal strVer As String, Optional ByVal blnNext As Boolean)
'���ܣ���ȡһ���汾�����汾
'������strVer=��ǰ�汾
'          blnNext=�Ƿ��ȡ��һ�����汾
'���أ����汾
    Dim arrVer As Variant
    
    arrVer = Split(strVer & "..", ".")
    If blnNext Then
        GetPrimaryVer = Val(arrVer(0)) & "." & (Val(arrVer(1)) + 1) & "." & 0
        '������û��9.45.0��ֱ�Ӻ�Ӧ��ϵͳͬһ��ţ�Ϊ10.34.0
        If GetPrimaryVer = "9.45.0" Then GetPrimaryVer = "10.34.0"
    Else
        GetPrimaryVer = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & 0
    End If
End Function

Public Function GetContractVersion(ByVal strVer As String, Optional ByVal blnGetTools As Boolean = True)
'���ܣ���ȡӦ��ϵͳ��Ӧ�����ߵ����汾�����߹����߶�ӦӦ��ϵͳ�汾����Ҫ����
'������strVer=��ǰӦ��ϵͳ�汾
'          blnGetTools=True-��ȡ��Ӧ�Ĺ����߰汾,False-��ȡ��Ӧ��Ӧ��ϵͳ�汾
'���أ���Ӧ�汾��Ӧ��ϵͳ10.34.0֮ǰ��ֻ���Ӧ��汾�������嵽SP�汾
'                          ������10.34.0֮ǰ��ֻ���Ӧ��汾�������嵽SP�汾
    Dim arrVer As Variant
    Dim lngDistance As Long
    If strVer = "" Then strVer = "9.1.0"
    If blnGetTools Then
        If VerFull(strVer) >= VerFull("10.34.0") Then '10.34.0  �Ժ�����ߺ�Ӧ��ϵͳ�汾ͳһ
            GetContractVersion = strVer
        Else
            arrVer = Split(strVer & "...", ".")
            lngDistance = 33 - Val(arrVer(1)) '��ȡӦ��ϵͳ��10.33.0�汾�Ĵ�汾���
            '������9.44.0��ȥ��Ӧ��汾�����Ϊ��Ӧ�����߰汾
            GetContractVersion = "9." & (44 - lngDistance) & ".0"
        End If
    Else
        If VerFull(strVer) >= VerFull("10.34.0") Then  '  �Ժ�����ߺ�Ӧ��ϵͳ�汾ͳһ
            GetContractVersion = strVer
        Else
            arrVer = Split(strVer & "...", ".")
            lngDistance = 44 - Val(arrVer(1)) '��ȡ��������9.44.0�汾�Ĵ�汾���
            'Ӧ��ϵͳ10.33.0��ȥ��Ӧ��汾�����Ϊ��ӦӦ��ϵͳ�İ汾
            GetContractVersion = "10." & (33 - lngDistance) & ".0"
        End If
    End If
End Function

Public Function VerNormal(ByVal strVer As String) As String
'���ܣ���VB���֧�ֵİ汾����ʽ:9999.9999.9999ת��Ϊ�����汾����ʽ����0010.0034.0000.0000��ת��Ϊ10.34.0
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then Exit Function
    arrVer = Split(strVer & ".", ".")
    VerNormal = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2)) & IIf(Val(arrVer(3)) = 0, "", "." & Format(Val(arrVer(3)), "0000"))
End Function

Public Function VerSpecialNormal(ByVal strVer As String) As String
'��ȡһ������sp��Ӧ����ʽ�汾�������һ����ʽ�汾���򷵻�������
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then Exit Function
    arrVer = Split(strVer & ".", ".")
    VerSpecialNormal = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2))
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

