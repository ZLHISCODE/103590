Attribute VB_Name = "mdlUpgrade"
Option Explicit
'�ļ�����,������˳�����ļ�ִ��˳����ͬ
Public Enum FileType
    FT_BefUp = 1 '��ǰִ�нű���ZLUPgradeX.X.X_Before.sql.sql(�����ߣ�,ZL*_X.X.X_History_Before.sql (Ӧ��ϵͳ��ʷ��)ZL*_X.X.X_Before.sql(Ӧ��ϵͳ���߿�) *����ϵͳ��\100
    FT_DBAUp = 2 '��ҪDBA�û�ִ�еĽű�(System�û�):ZLUPgradeX.X.X_DBA.sql,ZL*_X.X.X_DBA.sql
    FT_StUp = 3 '��ͨ�����ű���ZLUPgradeX.X.X.sql,ZLUPgradeX.X.X(����).sql,ZL*_X.X.X.sql ,ZL*_X.X.X(����).sql,ZL*_X.X.X_History.sql
    FT_OptUp = 4 '��ѡִ�нű�:ZLUPgradeX.X.X_Optional.sql,ZL*_X.X.X_Optional.sql��ZL*_X.X.X__HISTORY_Optional.sql
    FT_DefUp = 5 '�ӳ�ִ�нű�:ZL*_X.X.X_Deferred.sql,ZL*_X.X.X__HISTORY_DEFERRED
End Enum
'�ļ�����ϵͳ
Public Enum SysType
    ST_Tools = 1 '�����߽ű�,�����ļ����ͣ�FT_BefUp,FT_DBAUp,FT_StUp,FT_OptUp
    ST_App = 2 'Ӧ��ϵͳ���߿�,�����ļ����ͣ�FT_BefUp,FT_DBAUp,FT_StUp,FT_OptUp��FT_DefUp
    ST_AppHis = 3 'Ӧ��ϵͳ��ʷ�⣬�����ļ����ͣ�FT_BefUp,FT_StUp,FT_DefUp��FT_OptUp
End Enum
'�汾����
Public Enum VersionType
    VT_Normal = 1 '�����汾
    VT_Supple = 2 '���䷢���汾����һ����汾������ǰһ���汾�·�����SP���ǲ���汾
End Enum

'Public Enum FileTypeSys
'    FT_toolsUp = 1 '�����߽ű��� ��ʽ�� ZLUPgradeX.X.X.sql����ļ�
'    FT_toolsUpDbA = 2 '�����߽ű��� ��ʽ�� ZLUPgradeX.X.X_DBA.sql����ļ�
'    FT_toolsUpOpt = 3 '�����߽ű��� ��ʽ�� ZLUPgradeX.X.X_Optional.sql����ļ�
'    FT_toolsUpBef = 4 '�����߽ű��� ��ʽ�� ZLUPgradeX.X.X_Before.sql.sql����ļ�
'
'    FT_SysUp = 1 'ϵͳ�����ű��� ��ʽ��ZL*_X.X.X.sql ����ļ�  *����ϵͳ��\100
'    FT_SysUpDBA = 2 'ϵͳ�����ű��� ��ʽ��ZL*_X.X.X_DBA.sql ����ļ� *����ϵͳ��\100
'    FT_SysUpOpt = 3 'ϵͳ�����ű��� ��ʽ��ZL*_X.X.X_Optional.sql ����ļ� *����ϵͳ��\100
'    FT_SysUpHis = 4 'ϵͳ�����ű��� ��ʽ��ZL*_X.X.X_History.sql ����ļ� *����ϵͳ��\100
'    FT_SysUpBef = 5 'ϵͳ�����ű��� ��ʽ��ZL*_X.X.X_Before.sql ����ļ�  *����ϵͳ��\100
'    FT_SysUpHisBef = 6 'ϵͳ�����ű��� ��ʽ��ZL*_X.X.X_History_Before.sql ����ļ�  *����ϵͳ��\100
'    FT_SysUpDef = 7 'ϵͳ�����ű��� ��ʽ��ZL*_X.X.X_Deferred.sql ����ļ� *����ϵͳ��\100
'    FT_SysUpHisDef = 8 'ϵͳ�����ű��� ��ʽ��ZL*_X.X.X__HISTORY_DEFERRED ����ļ� *����ϵͳ��\100
'    FT_SysUpOther = 9 'ϵͳ�����ű��� ��ʽ��ZL*_X.X.X(����).sql ����ļ� *����ϵͳ��\100
'End Enum

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'������:����
'�޸��ˣ���˶
'�޸����ڣ�2014-1-6
'�޸ĵ㣺���Ӹ��Ƽ�¼���Ĳ����ֶι���
'��������:2000-11-02
'���Ƽ�¼��
'������strFields=��Ҫ���Ƶļ�¼�����ֶε���˳����ֶ�����ɵ��ַ���
'          �磺1 ����1,3 ����2,7 ����3...��ʾ���Ƽ�¼���ĵ�1,3,7..�ֶ���ɼ�¼��������
'              ID ����1,���� ����2,....��ʾ���Ƽ�¼����ID,����...�ֶ���ɼ�¼������
'              ����*Ϊ�µļ�¼��������
'              �������ͻ�����׳���������ͬ�����⣬��ע��
'           arrAppFields=׷�ӵ��ֶ���Ϣ������,����,����,Ĭ��ֵ,û��Ĭ��ֵ��Empty,û��ָ�����ȴ�Empty
'      blnOnlyStructure=�Ƿ�ֻ���ƽṹ
'�ڳ����У��������漰���໥���ݼ�¼������ʹ��ADO��Clone���Ʋ����ļ�¼����������һ����¼�������ݷ����仯��ʱ�����и�������������ͬ�ı仯��ͨ��ָ�޸Ļ�ɾ����������������ϣ����Щ��¼���໥�䱣�ֶ���
  
    Dim rsClone As New ADODB.Recordset
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    Dim arrFieldsName As Variant, strFieldName As String, strFieldNameAlias As String
    Dim arrTmp As Variant
    Dim i As Long
    
    If Not rsSource Is Nothing Then
        Set rsClone = rsSource.Clone
        rsClone.Filter = rsSource.Filter
    End If
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        '������¼���ṹ
        If strFields = "" Then '��¼��ȫ����ģʽ
            arrFieldsName = Array()
            If Not rsClone Is Nothing Then
                ReDim arrFieldsName(rsClone.Fields.Count - 1)
                For intFields = 0 To rsClone.Fields.Count - 1
                    arrFieldsName(intFields) = rsClone.Fields(intFields).name & ""
                    .Fields.Append rsClone.Fields(intFields).name, IIf(rsClone.Fields(intFields).Type = adNumeric, adDouble, rsClone.Fields(intFields).Type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
                Next
            End If
        Else '��¼�����ָ���ģʽ
            arrFieldsName = Split(strFields, ",")
            For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                '�а�������
                arrTmp = Split(arrFieldsName(intFields) & " ", " ")
                strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).name & ""
                '��ȡ�ֶ�ԭ������������
                arrFieldsName(intFields) = strFieldName
                '����ֶ�,�������ڱ������������е�����Ϊ����
                .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).Type = adNumeric, adDouble, rsClone.Fields(strFieldName).Type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:��ʾ����
            Next
        End If
        '׷���ֶ����
        If TypeName(arrAppFields) = "Variant()" Then
            For i = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(i + 2) = Empty Then
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable, arrAppFields(i + 3)
                    End If
                Else
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable, arrAppFields(i + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '��������
        If Not blnOnlyStructure Then
            If rsClone Is Nothing Then Exit Function
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '�¼�¼�����а�˳����ӣ���˿�������
                    .Fields(intFields).value = rsClone.Fields(arrFieldsName(intFields)).value
                Next
                .Update
                rsClone.MoveNext
            Loop
            If rsClone.RecordCount <> 0 Then .Filter = "": .MoveFirst
        End If
    End With
    
    Set CopyNewRec = rsTarget
End Function

Public Function RecDelete(ByRef rsInput As ADODB.Recordset, Optional ByVal strFilter As String) As Boolean
'���ܣ�ɾ��ָ�������ļ�¼���ļ�¼
'������rsInput=��¼��
'      strFilter=����
'���أ��Ƿ�ɹ�
'      rsInput=����ɾ����ļ�¼��
    rsInput.Filter = strFilter
    If rsInput.RecordCount > 0 Then
        rsInput.MoveFirst
        Do While Not rsInput.EOF
            Call rsInput.Delete
            rsInput.MoveNext
        Loop
        Call rsInput.UpdateBatch
    End If
    RecDelete = True
End Function

Public Function RecUpdate(ByRef rsInput As Recordset, ByVal strFilter As String, ParamArray arrInput() As Variant) As Boolean
'���ܣ�����ָ�������ļ�¼���ļ�¼
'������rsInput=��¼��
'      strFilter=����
'      arrInput=������ֶ����Լ�ֵ����ʽ���ֶ���1,ֵ1, �ֶ���2,ֵ2,....
'���أ��Ƿ�ɹ�
'      rsInput=�������º�ļ�¼��
'˵����arrInput���ֶ�ֵ�����ü�¼���е������ֶ������¸��ֶΣ���ʱ��ʽΪ��!�ֶ���
    Dim strFiledName As String, strFileValue As String
    Dim blnFiled As Boolean, i As Long

    On Error GoTo errH
    With rsInput
        .Filter = strFilter
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            For i = LBound(arrInput) To UBound(arrInput) Step 2
                strFiledName = arrInput(i)
                If IsNull(arrInput(i + 1)) Then
                    rsInput(strFiledName).value = Null
                Else
                    If arrInput(i + 1) Like "!?*" Then
                        blnFiled = True
                        On Error Resume Next
                        strFileValue = rsInput(Mid(arrInput(i + 1), 2)).value & ""
                        If err.Number <> 0 Then err.Clear: blnFiled = False
                        On Error GoTo errH
                    End If
                    If Not blnFiled Then
                        rsInput(strFiledName).value = arrInput(i + 1)
                    Else
                        rsInput(strFiledName).value = rsInput(Mid(arrInput(i + 1), 2)).value
                    End If
                End If
                blnFiled = False
            Next
            .MoveNext
        Loop
        Call rsInput.UpdateBatch
    End With
    RecUpdate = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function RecDataAppend(ByRef rsSource As ADODB.Recordset, ByVal rsAppend As ADODB.Recordset, ParamArray arrInput() As Variant) As Boolean
'���ܣ���ָ����¼����������ӵ���һ����¼����
'������rsSource=Ŀ���¼��
'      rsAppend=���ݼ�¼��
'      arrInput=�ֶζ�Ӧ���򣬸ò�������ʱ��Ĭ������¼���ṹ��ͬ����ʽ��arrInput(0):[��¼��1].�ֶ�1,�ֶ�2...��arrInput(1)��[��¼��2].�ֶ�1,�ֶ�2...
'���أ��Ƿ�ɹ�
'      rsSource=������ݺ�ļ�¼��
    Dim arrSource As Variant, arrAppend As Variant
    Dim i As Long, arrValues() As Variant
    Dim strTmp As String
    
    If rsAppend Is Nothing Then RecDataAppend = True: Exit Function
    If rsAppend.RecordCount = 0 Then RecDataAppend = True: Exit Function
    If rsSource Is Nothing Then Exit Function
    On Error GoTo errH
    If LBound(arrInput) = 2 Then
        arrSource = Split(arrInput(LBound(arrInput)), ",")
        arrAppend = Split(arrInput(UBound(arrInput)), ",")
        If UBound(arrSource) <> UBound(arrAppend) Then Exit Function
        ReDim arrValues(UBound(arrAppend)): rsAppend.MoveFirst
        Do While Not rsAppend.EOF
            For i = LBound(arrAppend) To UBound(arrAppend)
                arrValues(i) = rsAppend(arrAppend(i)).value
            Next
            rsSource.AddNew arrSource, arrValues
            Erase arrValues
            rsAppend.MoveNext
        Loop
    ElseIf LBound(arrInput) = 0 Then
        Do While Not rsAppend.EOF
            rsSource.AddNew
            For i = 0 To rsSource.Fields.Count - 1
                rsSource.Fields(i).value = rsAppend.Fields(i).value
            Next
            rsSource.Update
            rsAppend.MoveNext
        Loop
    End If
    
    RecDataAppend = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
    
End Function

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
        If IsNull(rsINI!����) Then Exit Function
        If arrItem(i) Like "*�汾��" Then
            If Not IsVerSion(rsINI!����) Then Exit Function
        End If
    Next
    CheckINIValid = True
End Function

Public Function VerCompare(ByVal strVer1 As String, ByVal strVer2 As String, Optional ByVal blnPrimary As Boolean) As Integer
'���ܣ��Ƚ������ַ�����ʾ�İ汾�ŵĴ�С
'������blnPrimary=�Ƿ�ֻ�Ƚ�"���汾.�ΰ汾",���ܸ��汾
'���أ�1=strVer1>strVer1,-1=strVer1<strVer1,0=strVer1=strVer1
'˵����VB���֧�ֵİ汾��Ϊ9999.9999.9999
    Dim arrVer As Variant
    
    If strVer1 Like "*.*.*" And strVer2 Like "*.*.*" Then
        arrVer = Split(strVer1, ".")
        strVer1 = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & IIf(blnPrimary, "", "." & Format(arrVer(2), "0000"))
        
        arrVer = Split(strVer2, ".")
        strVer2 = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & IIf(blnPrimary, "", "." & Format(arrVer(2), "0000"))
    End If
    If strVer1 > strVer2 Then
        VerCompare = 1
    ElseIf strVer1 < strVer2 Then
        VerCompare = -1
    End If
End Function

Public Function GetCurPriVersion(ByVal strVer As String) As String
'���ܣ���ȡ��ǰ�汾�Ĵ�汾
'������strVer ��ǰ�汾
'���أ� GetNextVersion ��ǰ�汾�Ĵ�汾
    Dim arrVer As Variant
    
    If IsVerSion(strVer) Then
        If Not strVer Like "*.*.0" Then
            arrVer = Split(strVer, ".")
            strVer = arrVer(0) & "." & arrVer(1) & ".0"
        End If
    Else
        Exit Function
    End If
    
    GetCurPriVersion = strVer
End Function

Public Function GetNextVersion(ByVal strVer As String, Optional ByVal blnPrimary As Boolean) As String
'���ܣ���ȡ��ǰSP�汾����һ���汾
'������strVer ��ǰ�汾
'     blnPrimary �Ƿ��ȡ��һ����汾
'���أ� GetNextVersion blnPrimary=true:��һ����汾,blnPrimary=false :��һ��SP�汾
    Dim arrVer As Variant
    
    If IsVerSion(strVer) Then
        arrVer = Split(strVer, ".")
        If blnPrimary Then
            strVer = Val(arrVer(0)) & "." & Val(arrVer(1)) + 1 & ".0"
        Else
            strVer = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2)) + 10
        End If
    Else
        Exit Function
    End If
    
    GetNextVersion = strVer
    
End Function

Public Function GetPreVersion(ByVal strVer As String, Optional ByVal blnPrimary As Boolean) As String
'���ܣ���ȡ��ǰSP�汾����һ���汾
'������strVer ��ǰ�汾
'     blnPrimary �Ƿ��ȡ��һ����汾
'���أ� GetNextVersion blnPrimary=true:��һ����汾,blnPrimary=false :��һ��SP�汾
    Dim arrVer As Variant
    
    If IsVerSion(strVer) Then
        arrVer = Split(strVer, ".")
        If blnPrimary Then
            If Val(arrVer(1)) - 1 < 0 Then
               arrVer(1) = "*"
               arrVer(0) = Val(arrVer(0)) - 1
            Else
                arrVer(1) = Val(arrVer(1)) - 1
            End If
            strVer = Val(arrVer(0)) & "." & arrVer(1) - 1 & ".0"
        Else
            If Val(arrVer(2)) - 10 < 0 Then
                arrVer(2) = "*"
                If Val(arrVer(1)) - 1 < 0 Then
                   arrVer(1) = "*"
                   arrVer(0) = Val(arrVer(0)) - 1
                Else
                    arrVer(1) = Val(arrVer(1)) - 1
                End If
            Else
                arrVer(2) = Val(arrVer(2)) - 10
            End If
            strVer = Val(arrVer(0)) & "." & arrVer(1) & "." & arrVer(2)
        End If
    Else
        Exit Function
    End If
    
    GetPreVersion = strVer
End Function

Public Function GetFileInfo(ByVal strFile As String, ByVal lngSys As Long, Optional ByRef strVerReturn As String, Optional ByRef ftReturn As FileType, _
                                    Optional ByRef stReturn As SysType, Optional ByRef vtReturn As VersionType) As Boolean
'����:��ȡ�ļ���Ϣ
'������
'   strFile=������·�����ļ���,����չ��
'   lngSys=ϵͳ��
'����:
'       True=�ɹ���ȡ��False=��ȡʧ�ܣ��ļ�����ϵͳ�����ű���
'       strVerReturn=�ļ��汾
'       ftReturn=�ļ�����
'       stReturn=ϵͳ����
'       vtReturn=�汾����
    Dim strSysString, strSuffix As String
    Dim arrVer As Variant, strVerTmp As String
    '��ʼ������
    strVerReturn = "": ftReturn = 0: stReturn = 0: vtReturn = VT_Normal
    If Not UCase(strFile) Like "*.SQL" Then Exit Function
    strFile = UCase(Left(strFile, Len(strFile) - 4))
    '��ȡ�ű�ϵͳǰ׺
    If strFile Like "ZLUPGRADE*.*.*" Then
        strSysString = "ZLUPGRADE"
        stReturn = ST_Tools
    ElseIf strFile Like "ZL" & lngSys \ 100 & "_*.*.*" Then
        strSysString = "ZL" & lngSys \ 100 & "_"
        stReturn = ST_App
    Else
        Exit Function 'û��ϵͳ��ʶǰ׺������ϵͳ�ű�
    End If
    'ϵͳ��ʶ����������ǰ汾
    strSuffix = Mid(strFile, Len(strSysString) + 1)
    arrVer = Split(strSuffix, ".")
    If UBound(arrVer) <> 2 Then Exit Function '�����汾�Ľű�������ϵͳ�ű�
    strVerTmp = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2))
    If Not IsVerSion(strVerTmp) Then Exit Function '���ǰ汾��Ϣ
    If Not strSuffix Like strVerTmp & "*" Then Exit Function
    '�汾�����ļ�������Ϣ
    strSuffix = Mid(strSuffix, Len(strVerTmp) + 1)
    If InStr(strSuffix, "(����)") > 0 Then
        vtReturn = VT_Supple
        strSuffix = Replace(strSuffix, "(����)", "") '��ֹ������Ϣλ�ò��̶�
    End If
    If stReturn = ST_App And strSuffix Like "_HISTORY*" Then
        stReturn = ST_AppHis
        strSuffix = Mid(strSuffix, Len("_HISTORY") + 1)
    End If
    Select Case strSuffix
        Case ""
            ftReturn = FT_StUp
        Case "_DBA"
            If stReturn <> ST_AppHis Then ftReturn = FT_DBAUp
        Case "_OPTIONAL"
            ftReturn = FT_OptUp
        Case "_BEFORE"
            ftReturn = FT_BefUp
        Case "_DEFERRED"
            If stReturn <> ST_Tools Then ftReturn = FT_DefUp
    End Select
    If ftReturn = 0 Then Exit Function
    strVerReturn = strVerTmp
    GetFileInfo = True
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean = True) As String
'���ܣ�����VB���֧�ֵİ汾����ʽ:9999.9999.9999,��С�汾��0000.0000.0000
'������strVer=��ǰ�汾��
'           blnMax=True,����Ϊ�գ��򷵻����֧�ְ汾��False=����Ϊ�գ��򷵻���С֧�ְ汾
    Dim arrVer As Variant
    If strVer = "" Then
        VerFull = IIf(blnMax, "9999.9999.9999", "0000.0000.0000")
        Exit Function
    End If
    arrVer = Split(strVer, ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000")
End Function

Public Function VerNormal(ByVal strVer As String) As String
'���ܣ���VB���֧�ֵİ汾����ʽ:9999.9999.9999ת��Ϊ�����汾����ʽ����0010.0034.0000��ת��Ϊ10.34.0
    Dim arrVer As Variant
    If strVer = "" Then
        VerNormal = "0.0.0"
        Exit Function
    End If
    arrVer = Split(strVer, ".")
    VerNormal = Val(arrVer(0)) & "." & Val(arrVer(1)) & "." & Val(arrVer(2))
End Function

Public Function IsVerSion(ByVal strVer As String) As Boolean
'���ܣ��ж��ַ����Ƿ��ǰ汾��
    Dim arrVer As Variant
    Dim i As Integer
    If strVer = "" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) <> 2 Then Exit Function
    
    For i = LBound(arrVer) To UBound(arrVer)
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If Val(arrVer(i)) & "" <> arrVer(i) Then Exit Function
    Next
    
    IsVerSion = True
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
'���ܣ�ȡָ���ַ������ֽ���ĳ���
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function ActualStr(ByVal strAsk As String, ByVal lngLen As Long) As String
'���ܣ�ȡָ���ַ������ָ���ֽڳ��ȵ�����
    Dim strTemp As String, i As Long
    
    strTemp = StrConv(LeftB(StrConv(strAsk, vbFromUnicode), lngLen), vbUnicode)
    If InStr(strTemp, Chr(0)) > 0 Then
        strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
    End If
    ActualStr = strTemp
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
    Dim i As Long, k As Long
    
    If Left(strSQL, 2) <> "--" And InStr(strSQL, "--") > 0 Then
        For i = 1 To Len(strSQL)
            If Mid(strSQL, i, 1) = "'" Then blnStr = Not blnStr
            If Mid(strSQL, i, 2) = "--" And Not blnStr Then
                k = i: Exit For
            End If
        Next
        If k > 0 Then strSQL = RTrim(Left(strSQL, k - 1))
    End If
    TrimComment = strSQL
End Function

Public Function SplitSQL(ByVal strSQL As String) As String
'���ܣ�ȡ";"��βǰ��ĵ�SQL���,����";"�ź���"--"ע�͡�
'˵������Ҫ��RunSQLFile���Ӻ���
    Dim i As Long, k As Long
    
    '��ȥ��ע�Ͳ���
    strSQL = TrimComment(strSQL)
    
    For i = Len(strSQL) To 1 Step -1
        If Mid(strSQL, i, 1) = ";" Then
            k = i: Exit For
        End If
    Next
    If k > 0 Then strSQL = Left(strSQL, k - 1)
    
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

Public Function GetLogSQL(objSQL As clsSQLInfo) As String
'���ܣ���ȡ��ҪSQL��䣬������д��־
    Dim strSQL As String
    
    If objSQL.Block Then
        If objSQL.BlockName <> "" Then
            strSQL = Trim(Split(objSQL.SQL, vbCrLf)(0))
            If InStr(strSQL, "(") > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(strSQL, "(") - 1))
            End If
            If InStr(1, strSQL, " as", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " as", vbTextCompare) - 1))
            End If
            If InStr(1, strSQL, " is", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " is", vbTextCompare) - 1))
            End If
            If InStr(1, strSQL, " Return", vbTextCompare) > 0 Then
                strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " Return", vbTextCompare) - 1))
            End If
        Else '������
            strSQL = ActualStr(TrimEx(objSQL.SQL, True), 150)
        End If
    ElseIf UCase(LTrim(objSQL.SQL)) Like "CREATE * VIEW" Then
        '��ͼ���⴦��
        strSQL = Split(objSQL.SQL, vbCrLf)(0)
        If InStr(1, strSQL, " as", vbTextCompare) > 0 Then '��ͼֻ����as
            strSQL = RTrim(Left(strSQL, InStr(1, strSQL, " as", vbTextCompare) - 1))
        End If
    Else
        If InStr(objSQL.SQL, vbCrLf) > 0 Then
            '����SQL
            strSQL = ActualStr(TrimEx(objSQL.SQL, True), 150)
        Else
            strSQL = ActualStr(objSQL.SQL, 150)
        End If
    End If
    GetLogSQL = strSQL
End Function


Public Function CheckHavHistory(ByVal lngSys As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:����Ƿ���Ҫ������ʷ�ռ�
    '����:lngSys-ϵͳ��
    '����:��Ҫ����,��true,����False
    '--------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 1 from zltools.zlbakTables where ϵͳ=" & lngSys & " and rownum<=1"
    OpenRecordset rsTemp, gstrSQL, "��ȡbak����", , , gcnOracle
    If rsTemp.EOF Then
       '����False,��ʾ��ϵͳû����ʷ���ݿռ�,û��Ҫ������ʷ���ݿռ�
       Exit Function
    End If
    CheckHavHistory = True
End Function

Public Function GrantBakToUser(ByVal cnOracle As ADODB.Connection, ByVal strToOwner As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ƿ����
    '����:strTableName-����
    '     cnoracle-���ݿ�������
    '     strOwNer-������
    '����:���ڸñ���true,����False
    '-----------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    err = 0: On Error GoTo errHand:
    strSQL = "Select TABLE_NAME from user_all_tables" & _
            " Union All Select View_Name From User_Views"
    Call OpenRecordset(rsTemp, strSQL, "������Ȩ", , , cnOracle)
    With rsTemp
        Do While Not .EOF
            strSQL = "Grant ALL on " & Nvl(!Table_Name) & " to " & strToOwner & " With Grant Option"
            cnOracle.Execute strSQL
            .MoveNext
        Loop
    End With
    GrantBakToUser = True
    Exit Function
errHand:
    If MsgBox("����Ȩʱ�������´���,����!" & vbCrLf & " (" & err.Number & ") " & err.Description, vbRetryCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbRetry Then
        Resume
    End If
    GrantBakToUser = False
End Function


Public Function IsNetServer(ByVal strPath As String, ByVal strUser As String, ByVal strPassword As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '--����:���������Ƿ�����������
    '--����:strPath -����·��
    '       strUser-�û���
    '       strPassWord -��������
    '����:����˳��,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/06
    '----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
      
    '���˺�:���ܴ���windows��Դ�������Ѿ��з��ʵ���
    '
    If objFile.FolderExists(strPath) Then
        IsNetServer = True: Exit Function
    End If
    
    Dim NetR As NETRESOURCE
    With NetR
        .dwScope = RESOURCE_GLOBALNET
        .dwType = RESOURCETYPE_DISK
        .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
        .dwUsage = RESOURCEUSAGE_CONNECTABLE
        .lpLocalName = "" 'ӳ���������
        .lpRemoteName = strPath  '������·��
    End With
    
    err = 0
    On Error GoTo errHand:
    If WNetAddConnection2(NetR, strPassword, strUser, CONNECT_UPDATE_PROFILE) = NO_ERROR Then
       IsNetServer = True
    Else
       IsNetServer = False
    End If
    Exit Function
errHand:
       IsNetServer = False
End Function
Public Function CancelNetServer(ByVal strPath As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '����:�Ͽ�����������
    '����:
    '����:���ҳɹ�,����true,���򷵻�False
    '----------------------------------------------------------------------------------------------------------
    err = 0
    On Error Resume Next
    If WNetCancelConnection2(strPath, CONNECT_UPDATE_PROFILE, True) = 0 Then
        CancelNetServer = True
    Else
        CancelNetServer = False
    End If
    err = 0
End Function

Public Sub ReGrantForTools(ByVal cnTools As ADODB.Connection, Optional ByVal strSysOwner As String, Optional ByVal rsToolsObjs As ADODB.Recordset, Optional ByVal blnSysGrant As Boolean, Optional ByVal blnALLSysGrant As Boolean)
    '----------------------------------------------------------------------------------------------------------
    '����:�Թ����ߵĶ������������Ȩ������ͬ���
    '����:cnTools�����������ӡ�strSysOwnerΪ��ʱ�����Դ�Ӧ��ϵͳ���ӣ���ʱΪӦ��ϵͳת��Ȩ�ޡ�
    '     strSysOwner:Ӧ��ϵͳ�����ߡ�Ϊ���Ƿ������������ã�ֻ��������ͬ����Լ���Public��Ȩ
    '     rsToolsObjs�������߶����¼��������ʱ��ȡ���ݿ�
    '     blnSysGrant:ϵͳ������ת�ڹ�����Ȩ�ޣ����ǰ׺ZLTOOLS.
    '     blnALLSysGrant:������ϵͳ������Ȩ,��ʱstrSysOwner�����Ч
    '����:
    '----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset, rsSys As New ADODB.Recordset
    Dim arrObjects As Variant
    Dim i As Long
    Dim strObjectName As String
    Dim arrOwners() As Variant
    
    arrOwners = Array()
    '��ȡ������
    If blnALLSysGrant Then
        '�����߶���������Ȩ����������ͬ���
        gstrSQL = "Select Distinct ������ FROM zlsystems"
        OpenRecordset rsSys, gstrSQL, "��������ͬ���", , , cnTools
        Do While Not rsSys.EOF
            ReDim Preserve arrOwners(UBound(arrOwners) + 1)
            arrOwners(UBound(arrOwners)) = rsSys!������ & ""
            rsSys.MoveNext
        Loop
    ElseIf strSysOwner <> "" Then
        arrOwners = Array(strSysOwner)
    End If
    
    If rsToolsObjs Is Nothing Then
        strSQL = "Select Object_Name, Object_Type" & vbNewLine & _
                "From User_Objects" & vbNewLine & _
                "Where Object_Type In ('FUNCTION', 'PROCEDURE', 'TYPE', 'PACKAGE', 'SEQUENCE', 'TABLE', 'VIEW') And" & vbNewLine & _
                "      Instr(Object_Name, 'BIN$') <= 0"
        Call OpenRecordset(rsTemp, strSQL, "������Ȩ", , , cnTools)
    Else
        Set rsTemp = rsToolsObjs
    End If
    
    On Error Resume Next
    With rsTemp
        
        '�����߶���170�����ң�ͨ��ѭ����ִ��SQL��Լ900�����ң���ʱ��2-3��
        Do While Not .EOF
            '��Ӧ��ϵͳ���������������Ȩ��
            If blnSysGrant Then
                strObjectName = "ZLTOOLS." & !OBJECT_NAME
            Else
                strObjectName = !OBJECT_NAME & ""
            End If
            
            For i = 0 To UBound(arrOwners)
                Select Case !OBJECT_TYPE
                    Case "FUNCTION", "PROCEDURE", "TYPE", "PACKAGE"
                        strSQL = "grant execute on " & strObjectName & " to " & arrOwners(i) & " With GRANT Option"
                    Case "VIEW"
                        strSQL = "grant select on " & strObjectName & " to " & arrOwners(i) & " With GRANT Option"
                    Case "SEQUENCE"
                        strSQL = "grant select,alter on " & strObjectName & " to " & arrOwners(i) & " With GRANT Option"
                    Case "TABLE"
                        strSQL = "grant select,insert,update,delete on " & strObjectName & " to " & arrOwners(i) & " With GRANT Option"
                End Select
                cnTools.Execute strSQL
            Next
            'ͬ�����������ɾ��ͬ��ʣ������´���
            cnTools.Execute "drop synonym " & !OBJECT_NAME: err.Clear
            cnTools.Execute "drop public synonym " & !OBJECT_NAME
            cnTools.Execute "create public synonym " & !OBJECT_NAME & " for " & strObjectName
            '������Ȩ������PUBLIC
            Select Case !OBJECT_TYPE
                Case "FUNCTION", "PROCEDURE", "TYPE", "PACKAGE"
                    If Not ",B_ROLEGROUPMGR,ZL_ZLROLEGRANT_BATCHDELETE,ZL_ZLROLEGRANT_BATCHINSERT," _
                         Like "*," & UCase(!OBJECT_NAME & "") & ",*" Then
                        strSQL = "grant execute on " & strObjectName & " to Public"
                    End If
                Case "SEQUENCE", "TABLE", "VIEW"
                    strSQL = "grant select on " & strObjectName & " to Public"
            End Select
            cnTools.Execute strSQL
            err.Clear
            .MoveNext
        Loop
    End With
End Sub

Public Function GrantSpecialToRole(ByVal cnOracle As ADODB.Connection, ByVal strRoleNames As String, ByVal blnGrantBase As Boolean, strOwners() As String) As Boolean
    '----------------------------------------------------------------------------------------------------------
    '����:�Թ����ߵĶ����Ӧ�ó���һЩ���������Ȩ������Ķ���
    '����:cnOracle��Ӧ��ϵͳ����
    '     strRoleNames:����Ȩ�Ľ�ɫ�������ɫ�Զ��ŷָһ�㲻����15����ɫ
    '     blnGrantBase:�Ƿ��Ӧ��ϵͳ�����������Ȩ
    '     strOwners��Ӧ��ϵͳ������
    '����:
    '----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim blnsysSt As String
    Dim strStSysOwner As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select ������ From zlSystems Where Floor(��� / 100) = 1"
    If UBound(strOwners) <> -1 Then
        OpenRecordset rsTmp, strSQL, "��ȡ��׼��ϵͳ������", , , cnOracle
        Do While Not rsTmp.EOF
            strStSysOwner = strStSysOwner & "," & rsTmp!������
            rsTmp.MoveNext
        Loop
        If strStSysOwner <> "" Then strStSysOwner = strStSysOwner & ","
    End If
    
    For i = LBound(strOwners) To UBound(strOwners)
        If strOwners(i) <> "" Then
            cnOracle.Execute "grant select on " & strOwners(i) & ".���ű� to " & strRoleNames
            cnOracle.Execute "grant select on " & strOwners(i) & ".��Ա�� to " & strRoleNames
            cnOracle.Execute "grant select on " & strOwners(i) & ".������Ա to " & strRoleNames
            cnOracle.Execute "grant select on " & strOwners(i) & ".�ϻ���Ա�� to " & strRoleNames
            cnOracle.Execute "grant select on " & strOwners(i) & ".��Ա����˵�� to " & strRoleNames
            cnOracle.Execute "grant select on " & strOwners(i) & ".��Ա���ʷ��� to " & strRoleNames
            If InStr(strStSysOwner, "," & strOwners(i) & ",") > 0 Then
                '��Ϣƽ̨����
                cnOracle.Execute "grant select on " & strOwners(i) & ".ҵ����Ϣ���� to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".ҵ����Ϣ�嵥 to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".ҵ����Ϣ���Ѳ��� to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".ҵ����Ϣ������Ա to " & strRoleNames
                cnOracle.Execute "grant select on " & strOwners(i) & ".ҵ����Ϣ״̬ to " & strRoleNames
                cnOracle.Execute "grant execute on " & strOwners(i) & ".Zlpub_ҵ����Ϣ�嵥_insert to " & strRoleNames
                cnOracle.Execute "grant execute on " & strOwners(i) & ".Zl_ҵ����Ϣ�嵥_insert to " & strRoleNames
                cnOracle.Execute "grant execute on " & strOwners(i) & ".Zl_ҵ����Ϣ�嵥_read to " & strRoleNames
            End If
            If blnGrantBase Then
                cnOracle.Execute "grant execute on " & strOwners(i) & ".zl_�ֵ����_execute to " & strRoleNames
            End If
        End If
    Next
    '�Է������ļ��������������Ȩ
    '------------------------------------------------------------------------------------------------------------------
    cnOracle.Execute "grant insert,update         on ZLTOOLS.zlDiaryLog to " & strRoleNames
    cnOracle.Execute "grant insert                on ZLTOOLS.zlErrorLog to " & strRoleNames
    cnOracle.Execute "grant update,delete         on ZLTOOLS.zlMessages to " & strRoleNames
    cnOracle.Execute "grant update,delete         on ZLTOOLS.zlMsgState to " & strRoleNames
    cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientScheme to " & strRoleNames
    cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientParaSet to " & strRoleNames
    cnOracle.Execute "grant insert,update,delete  on ZLTOOLS.zlClientparaList to " & strRoleNames
    cnOracle.Execute "grant Select on sys.dba_role_privs to " & strRoleNames
    GrantSpecialToRole = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function
