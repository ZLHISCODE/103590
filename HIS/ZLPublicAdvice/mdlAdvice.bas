Attribute VB_Name = "mdlAdvice"
Option Explicit

Private mobjVBA As Object
Private mobjScript As clsScript
Private mrsDefine As Recordset

Public Enum Enum_Inside_Program
    pסԺ���ʲ��� = 1150
    p���ﲡ������ = 1250
    pסԺ�������� = 1251
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    p�����¼���� = 1255
    p�ٴ�·��Ӧ�� = 1256
    p�����¼���� = 1256
    pҽ�����ѹ��� = 1257
    p���Ʊ������ = 1258
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    p������ϲο� = 1270
    pҩƷ���Ʋο� = 1271
    p���˲������� = 1273
    p��Ƭ���߹��� = 1289
    p��Һ�������� = 1345
End Enum

Public Function Get������Ŀ��¼(ByVal lngID As Long, Optional ByVal strIDs As String) As ADODB.Recordset
'���ܣ���ȡָ��������ĿID�ļ�¼
'������
    Dim StrSQL As String
    
    StrSQL = "Select /*+ rule*/ �������,վ��,���,����ID,ID,����,����,�걾��λ,���㵥λ,���㷽ʽ,ִ��Ƶ��,�����Ա�,����Ӧ��,�����Ŀ,��������,ִ�а���,ִ�п���,�������,�Ƽ�����,�ο�Ŀ¼ID,��ԱID,����ʱ��,����ʱ��,¼������,�Թܱ���,ִ�з���,ִ�б��" & _
            " From ������ĿĿ¼ Where ID"
    On Error GoTo errH
    If strIDs <> "" Then
        StrSQL = StrSQL & " IN(Select Column_Value From Table(f_Num2list([1])))"
        Set Get������Ŀ��¼ = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", strIDs)
    Else
        StrSQL = StrSQL & " = [1]"
        Set Get������Ŀ��¼ = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lngID)
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetMaxAdviceNO(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngӤ�� As Long) As Long
'���ܣ���ȡ��ǰ���˵����ҽ�����
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    If lng��ҳID = 0 Then
        StrSQL = "Select Nvl(Max(���),1) as ��� From ����ҽ����¼ Where ����ID=[1] And ��ҳID Is Null"
    Else
        StrSQL = "Select Nvl(Max(���),1) as ��� From ����ҽ����¼ Where ����ID=[1] And ��ҳID=[2] And Nvl(Ӥ��,0)=[3]"
    End If
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lng����ID, lng��ҳID, lngӤ��)
    If Not rsTmp.EOF Then GetMaxAdviceNO = rsTmp!���

    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function InitAdviceDefine() As Recordset
'���ܣ���ȡҽ�����ݶ����¼��
'������blnNew-�Ƿ񴴽�objVBA��objScript����
'˵����
    Dim StrSQL As String
    Dim rsDefine As Recordset
    

    On Error GoTo errH
    StrSQL = "Select �������,ҽ������ From ҽ�����ݶ��� Order by �������"
    Set rsDefine = New ADODB.Recordset
    Call gobjComlib.zlDatabase.OpenRecordset(rsDefine, StrSQL, "InitAdviceDefine")
    Set InitAdviceDefine = rsDefine
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function FormatExamineAdvice(ByVal strAdvicePro As String, _
    ByVal strAdvicePart As String, ByVal lngExeType As Long) As String
'��ʽ��ҽ���������
    Dim strReturn As String
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        Err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
    If mrsDefine Is Nothing Then Set mrsDefine = InitAdviceDefine
    mrsDefine.Filter = "�������='D'"
    If mrsDefine.RecordCount > 0 Then
        strReturn = mrsDefine!ҽ������ & ""
    End If

    If strReturn = "" Then
        strReturn = strAdvicePro & "," & _
                            Decode(lngExeType, 1, ",����ִ��", 2, ",����ִ��", "") & IIF(strAdvicePart <> "", ":" & get��λ����(strAdvicePart), "")
    Else
        If InStr(strReturn, "[�����Ŀ]") > 0 Then
            strReturn = Replace(strReturn, "[�����Ŀ]", _
                                            """" & strAdvicePro & Decode(lngExeType, 1, ",����ִ��", 2, ",����ִ��", "") & _
                                            """")
        End If

        '�滻��λ����
        If InStr(strReturn, "[��鲿λ]") > 0 Then
            strReturn = Replace(strReturn, "[��鲿λ]", _
                                            """" & get��λ����(strAdvicePart) & """")
        End If

        strReturn = mobjVBA.Eval(strReturn)
    End If

    FormatExamineAdvice = strReturn
End Function

Public Function FormatInspectionAdvice(ByVal str���� As String, ByVal str�ɼ� As String, ByVal str�걾 As String) As String
'���ܣ���������ҽ����ҽ������
    Dim i As Long, strText As String, strField As String, blnDefine As Boolean
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        Err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
    If mrsDefine Is Nothing Then Set mrsDefine = InitAdviceDefine
               
    'ȷ���Ƿ���
    blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
    If blnDefine Then
        mrsDefine.Filter = "�������='C'"
        If mrsDefine.EOF Then
            blnDefine = False
        ElseIf Trim(Nvl(mrsDefine!ҽ������)) = "" Then
            blnDefine = False
        End If
    End If
    
    If Not blnDefine Then
        strText = str���� & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
    Else
        strText = mrsDefine!ҽ������
        If InStr(strText, "[������Ŀ]") > 0 Then
            strField = str����
            strText = Replace(strText, "[������Ŀ]", """" & strField & """")
        End If
        If InStr(strText, "[����걾]") > 0 Then
            strField = str�걾
            strText = Replace(strText, "[����걾]", """" & strField & """")
        End If
        If InStr(strText, "[�ɼ�����]") > 0 Then
            strField = str�ɼ�
            strText = Replace(strText, "[�ɼ�����]", """" & strField & """")
        End If
        
        '����ҽ������
        On Error Resume Next
        strText = mobjVBA.Eval(strText)
        If mobjVBA.Error.Number <> 0 Then
            strText = str���� & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
        End If
        Err.Clear: On Error GoTo 0
    End If
        
    FormatInspectionAdvice = strText
End Function

Public Function FormatOperationAdvice(ByVal str���� As String, ByVal str���� As String, ByVal str���� As String, ByVal str����ʱ�� As String, ByVal str������λ As String) As String
'���ܣ���������ҽ����ҽ������
    Dim i As Long, strText As String, strField As String, blnDefine As Boolean
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        Err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
    If mrsDefine Is Nothing Then Set mrsDefine = InitAdviceDefine
               
    'ȷ���Ƿ���
    blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
    If blnDefine Then
        mrsDefine.Filter = "�������='F'"
        If mrsDefine.EOF Then
            blnDefine = False
        ElseIf Trim(Nvl(mrsDefine!ҽ������)) = "" Then
            blnDefine = False
        End If
    End If
    If Not blnDefine Then
        strText = Format(str����ʱ��, "MM��dd��HH:mm")
        If str���� <> "" Then
            strText = strText & IIF(str���� <> "", " �� " & str���� & " ���� ", " �� ")
        End If
        strText = strText & str���� & IIF(str������λ = "", "", "(��λ:" & str������λ & ")")
        If str���� <> "" Then
            strText = strText & " �� " & str����
        End If
    Else
        strText = mrsDefine!ҽ������
        If InStr(strText, "[����ʱ��]") > 0 Then
            strField = str����ʱ��
            strText = Replace(strText, "[����ʱ��]", """" & strField & """")
        End If
        If InStr(strText, "[��Ҫ����]") > 0 Then
            strField = str���� & IIF(str������λ = "", "", "(��λ:" & str������λ & ")")
            strText = Replace(strText, "[��Ҫ����]", """" & strField & """")
        End If
        If InStr(strText, "[��������]") > 0 Then
            strField = str����
            strText = Replace(strText, "[��������]", """" & strField & """")
        End If
        If InStr(strText, "[������]") > 0 Then
            strField = str����
            strText = Replace(strText, "[������]", """" & strField & """")
        End If
        '����ҽ������
        On Error Resume Next
        strText = mobjVBA.Eval(strText)
        If mobjVBA.Error.Number <> 0 Then
            strText = Format(str����ʱ��, "MM��dd��HH:mm")
            If str���� <> "" Then
                strText = strText & IIF(str���� <> "", " �� " & str���� & " ���� ", " �� ")
            End If
            strText = strText & str���� & IIF(str������λ = "", "", "(��λ:" & str������λ & ")")
            If str���� <> "" Then
                strText = strText & " �� " & str����
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
            
    FormatOperationAdvice = strText
End Function

Private Function get��λ����(ByVal strExtData As String) As String
'��:��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����
'��:��λ��1(������1,������2),��λ��2(������1,������2)-----
Dim i As Integer, strReturn As String, Arr��λ
    If strExtData = "" Then Exit Function
    Arr��λ = Split(Split(strExtData, Chr(9))(0), "|")

    For i = 0 To UBound(Arr��λ)
        strReturn = strReturn & "," & Split(Arr��λ(i), ";")(0) & "(" & Split(Arr��λ(i), ";")(1) & ")"
    Next

    get��λ���� = Mid(strReturn, 2)
End Function

Public Function Getִ������(ByVal lng���ͺ� As Long, ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal str��� As String _
       , ByVal strҽ������ As String, ByVal blnMove As Boolean) As String
'���ܣ�����ָ����ҽ��ID,����ִ��ҽ�����ݹ���ʾ
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, strTmp As String
    Dim bln��ҩ;�� As Boolean, i As Integer
    Dim strƤ�Խ�� As String

    On Error GoTo errH
    
    '��ȡҽ������
    If (str��� = "C" And lng���ID <> 0) Or str��� = "D" Then
        strTmp = strҽ������
        
    ElseIf str��� <> "E" Or lng���ID <> 0 Then
        '�䷽�巨,��������,��Ѫ;��,������ҽ��,ֱ����ʾҽ������
        StrSQL = "Select ҽ������ From ����ҽ����¼ Where ID=[1]"
        If blnMove Then
            StrSQL = Replace(StrSQL, "����ҽ����¼", "H����ҽ����¼")
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Getִ������", IIF(str��� = "E", lng���ID, lngҽ��ID))
        If Not rsTmp.EOF Then strTmp = rsTmp!ҽ������ & ""
    Else
        '���ΪE,�����ID=0
        StrSQL = "Select A.ID,A.���ID,A.�������,A.ҽ������,A.Ƥ�Խ��,A.��������,B.���㵥λ,B.��������,A.ִ��Ƶ��,A.ִ��ʱ�䷽��,B.����" & _
            " From ����ҽ����¼ A,������ĿĿ¼ B" & _
            " Where Not (A.�������='E' And ���ID is Not NULL) And A.������ĿID=B.ID" & _
            " And (A.���ID=[1] Or A.ID=[1])" & _
            " Order by A.���"
        If blnMove Then
            StrSQL = Replace(StrSQL, "����ҽ����¼", "H����ҽ����¼")
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Getִ������", lngҽ��ID)
        rsTmp.Filter = "���ID=" & lngҽ��ID
        If Not rsTmp.EOF Then bln��ҩ;�� = InStr(",5,6,", rsTmp!�������) > 0
        
        If Not bln��ҩ;�� Then
            'һ��������Ŀ����ҩ�÷�����ɼ�����
            rsTmp.Filter = 0
            If Not rsTmp.EOF Then
                If rsTmp!������� = "E" And rsTmp!�������� = "1" Then
                    strƤ�Խ�� = "��Ƥ�Խ����" & rsTmp!Ƥ�Խ��
                    
                    StrSQL = "Select b.������Ӧ, b.����ʱ�� From ����ҽ����¼ A, ���˹�����¼ B, ������ĿĿ¼ C, �����÷����� D" & _
                        " Where a.����id = b.����id And a.������Ŀid = d.�÷�id And d.��Ŀid = c.Id And c.��� In ('5', '6') And d.��Ŀid = b.ҩ��id And" & _
                        " Nvl(d.����, 0) = 0 And b.��¼ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = a.id And �������� = 10) And a.Id = [1] And RowNum<2"

                    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Getִ������", lngҽ��ID)
                    
                    If Not rsTmp.EOF Then
                        strƤ�Խ�� = strƤ�Խ�� & ",����ʱ�䣺" & Format(rsTmp!����ʱ��, "yyyy-MM-dd") & IIF(rsTmp!������Ӧ & "" = "", "", ",������Ӧ��" & rsTmp!������Ӧ)
                    End If
                End If
            End If
            
            StrSQL = "Select ҽ������ From ����ҽ����¼ Where ID=[1]"
            If blnMove Then
                StrSQL = Replace(StrSQL, "����ҽ����¼", "H����ҽ����¼")
            End If
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Getִ������", lngҽ��ID)
            If Not rsTmp.EOF Then strTmp = rsTmp!ҽ������ & ""
        Else
            '��ҩ;��
            For i = 1 To rsTmp.RecordCount
                strTmp = strTmp & vbCrLf & IIF(i = rsTmp.RecordCount, "��", "��") & rsTmp!ҽ������ & IIF(Not IsNull(rsTmp!��������), " " & FormatEx(rsTmp!��������, 5) & rsTmp!���㵥λ, "")
                rsTmp.MoveNext
            Next
            rsTmp.Filter = "ID=" & lngҽ��ID
            strTmp = rsTmp!���� & "," & rsTmp!ִ��Ƶ�� & "(" & rsTmp!ִ��ʱ�䷽�� & "):ÿ" & rsTmp!���㵥λ & " " & Mid(strTmp, 2)
        End If
    End If
    
    Getִ������ = strTmp & strƤ�Խ��
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Sub GetAdvicePartSaveSql(ByVal lngAdviceID As Long, ByVal strProjectName As String, _
    ByRef curAdviceInf As clsExamineAdvice, ByRef arySql As Variant, ByVal lng��� As Long, ByVal lng������� As Long)
'��ȡ��λҽ���ı���sql
'������lng���=��ҽ����¼�����
    Dim i As Long, j As Long
    Dim str��λ As String
    Dim strTmp���� As String
    Dim str���� As String
    Dim lngҽ����� As Long
    Dim lngTmpID As Long
    Dim rsData As ADODB.Recordset
    Dim StrSQL As String

    lngҽ����� = lng���

    StrSQL = "select id from ����ҽ����¼ where ���id=[1]"
    Set rsData = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "��ѯ��λҽ��", lngAdviceID)

    If rsData.RecordCount > 0 Then
        While Not rsData.EOF
            ReDim Preserve arySql(UBound(arySql) + 1)

            arySql(UBound(arySql)) = " ZL_����ҽ����¼_Delete(" & Val(Nvl(rsData!ID)) & ")"

            Call rsData.MoveNext
        Wend
    End If


    '��֯��λ�������
    For i = 0 To UBound(Split(curAdviceInf.��λ����, "|")) '��λ1;����1,����2,����3|��λn;����1,����2,����3---

        str��λ = Split(Split(curAdviceInf.��λ����, "|")(i), ";")(0)
        strTmp���� = Split(Split(curAdviceInf.��λ����, "|")(i), ";")(1)

        For j = 0 To UBound(Split(strTmp����, ","))
            lngҽ����� = lngҽ����� + 1     '����ҽ����¼.��ţ�����
            str���� = Split(strTmp����, ",")(j)
            lngTmpID = gobjComlib.zlDatabase.GetNextID("����ҽ����¼")

            ReDim Preserve arySql(UBound(arySql) + 1)

            arySql(UBound(arySql)) = "ZL_����ҽ����¼_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                 lngҽ����� & "," & curAdviceInf.������Դ & "," & curAdviceInf.����ID & "," & IIF(curAdviceInf.��ҳID = 0, "NULL", curAdviceInf.��ҳID) & "," & _
                 curAdviceInf.Ӥ�� & ",1,1,'D'," & curAdviceInf.�����ĿID & ",NULL,NULL,NULL,1," & _
                 "'" & strProjectName & "',NULL," & _
                 "'" & str��λ & "','һ����',NULL,NULL,NULL,NULL,0," & _
                 curAdviceInf.ִ�п���ID & "," & IIF(curAdviceInf.ִ�п���ID <= 0, "5", curAdviceInf.ִ�п�������) & "," & curAdviceInf.������־ & ",to_date('" & Format(curAdviceInf.��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),NULL," & _
                 curAdviceInf.���˿���ID & "," & curAdviceInf.��������ID & _
                 ",'" & curAdviceInf.����ҽ�� & "',to_date('" & Format(curAdviceInf.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),'" & curAdviceInf.�Һŵ� & "',Null,'" & str���� & "'," & curAdviceInf.ִ������ & ",NULL,NULL,'',NULL,NULL,NULL,NULL," & lng������� & ")"
        Next
    Next

End Sub

Public Sub GetAdviceAffixSaveSql(ByVal lngAdviceID As Long, ByRef arrSQL As Variant, ByVal str���� As String)
'��ȡҽ�������Ĵ洢sql
    Dim arrAppend As Variant
    Dim j As Long
    
    arrAppend = Array()
    If str���� <> "" Then
        arrAppend = Split(str����, "<Split1>")
        For j = 0 To UBound(arrAppend)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & lngAdviceID & "," & _
                "'" & Split(arrAppend(j), "<Split2>")(0) & "'," & Val(Split(arrAppend(j), "<Split2>")(1)) & "," & _
                j + 1 & "," & ZVal(Split(arrAppend(j), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(j), "<Split2>")(3), "'", "''") & "'" & _
                IIF(j = 0, ",1", "") & ")"
        Next
    End If
End Sub

Public Function Check�ϰల��(ByVal blnҩ�� As Boolean) As Boolean
'���ܣ����ҽԺ�Ŀ����Ƿ�ʹ�����ϰల��
'������blnҩ��=�Ǽ��ҩ���ϰ໹����������
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    Static blnҩ��Load As Boolean
    Static blnҩ��Last As Boolean
    Static bln��ҩLoad As Boolean
    Static bln��ҩLast As Boolean
    
    If blnҩ�� Then '�Ƿ��а���ֻ���ȡһ��
        If blnҩ��Load Then Check�ϰల�� = blnҩ��Last: Exit Function
    Else
        If bln��ҩLoad Then Check�ϰల�� = bln��ҩLast: Exit Function
    End If
    
    On Error GoTo errH
    
    If blnҩ�� Then
        StrSQL = "Select 1 From ��������˵�� A,���Ű��� B" & _
            " Where A.����ID=B.����ID And A.�������� IN('��ҩ��','��ҩ��','��ҩ��') And Rownum<2"
    Else
        StrSQL = "Select 1 From ��������˵�� A,���Ű��� B" & _
            " Where A.����ID=B.����ID And A.�������� Not IN('��ҩ��','��ҩ��','��ҩ��') And Rownum<2"
    End If
    Call gobjComlib.zlDatabase.OpenRecordset(rsTmp, StrSQL, "Check�ϰల��")
    Check�ϰల�� = rsTmp.RecordCount > 0
    
    If blnҩ�� Then
        blnҩ��Load = True: blnҩ��Last = Check�ϰల��
    Else
        bln��ҩLoad = True: bln��ҩLast = Check�ϰల��
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get����Ա����ID(ByVal int������� As Integer) As Long
'���ܣ�ȡ����Ա���������ָ������Ĳ��ţ�ȱʡ��������
    Static rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    
    If blnNew Then
        StrSQL = "Select Distinct B.����ID,Nvl(B.ȱʡ,0) as ȱʡ,C.������� From ������Ա B,��������˵�� C" & _
            " Where B.��ԱID = [1] And B.����ID=C.����ID" & _
            " Order by ȱʡ Desc"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", UserInfo.ID)
    End If
    rsTmp.Filter = "������� = 3 or ������� = " & int�������
    
    If Not rsTmp.EOF Then
        Get����Ա����ID = rsTmp!����ID
    Else
        Get����Ա����ID = UserInfo.����ID
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
'���ܣ���ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
        
    On Error GoTo errH
    If str���� <> "" Then
        StrSQL = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", str����)
    Else
        StrSQL = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckPatiDataMoved(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��ж�ָ�����˵������Ƿ���ת��
    Dim rsTmp As ADODB.Recordset, StrSQL As String
 
    StrSQL = "Select ����ת�� From ������ҳ Where ����ID = [1] And ��ҳID = [2]"
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "���ת��", lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        CheckPatiDataMoved = Val("" & rsTmp!����ת��) = 1
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub InitSendRecordset(rsExec As Recordset, rsBill As Recordset, rsRXKey As Recordset, rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, rsUpload As ADODB.Recordset, _
    rsNumber As ADODB.Recordset, rsMoneyNow As ADODB.Recordset, rsItems As ADODB.Recordset)
'���ܣ���ʼ��ҽ����������Ķ�̬��¼��
    '��ʼ��ҽ���Ƽۼ�¼��
    Set rsExec = New ADODB.Recordset
    
    rsExec.Fields.Append "ҽ��ID", adBigInt
    rsExec.Fields.Append "���ͺ�", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "Ҫ��ʱ��", adDate, , adFldIsNullable
    rsExec.Fields.Append "����", adDouble, , adFldIsNullable
    rsExec.Fields.Append "��������", adInteger, , adFldIsNullable
    
    rsExec.CursorLocation = adUseClient
    rsExec.LockType = adLockOptimistic
    rsExec.CursorType = adOpenStatic
    rsExec.Open
    
    '��ʼ��ҽ�����ʵ������ɼ�¼��
    Set rsBill = New ADODB.Recordset
    
    rsBill.Fields.Append "Key", adVarChar, 100
    rsBill.Fields.Append "NO", adVarChar, 30
    rsBill.Fields.Append "�������", adBigInt
    rsBill.Fields.Append "�������", adBigInt
    rsBill.CursorLocation = adUseClient
    rsBill.LockType = adLockOptimistic
    rsBill.CursorType = adOpenStatic
    rsBill.Open
        
    Set rsRXKey = New ADODB.Recordset
    rsRXKey.Fields.Append "Key", adVarChar, 200
    rsRXKey.Fields.Append "ҽ��ID", adVarChar, 200
    rsRXKey.Fields.Append "����", adBigInt
    rsRXKey.Fields.Append "����", adBigInt
    rsRXKey.CursorLocation = adUseClient
    rsRXKey.LockType = adLockOptimistic
    rsRXKey.CursorType = adOpenStatic
    rsRXKey.Open
    
    'SQL��¼��
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "����", adInteger '1-�Ƽ�,2-ǩ��,3-У��,4-����,5-����,6-����
    rsSQL.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsSQL.Fields.Append "��ĿID", adBigInt '�շ�ϸĿID
    rsSQL.Fields.Append "���", adBigInt '��������
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.Fields.Append "NO", adVarChar, 30, adFldIsNullable '����NO�滻����ʱ����
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    '�Ƽ������ۼƼ�¼��
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsTotal.Fields.Append "��ĿID", adBigInt
    rsTotal.Fields.Append "�ⷿID", adBigInt
    rsTotal.Fields.Append "����", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    'ҽ���ϴ����ʵ�
    Set rsUpload = New ADODB.Recordset
    rsUpload.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsUpload.Fields.Append "NO", adVarChar, 30
    rsUpload.CursorLocation = adUseClient
    rsUpload.LockType = adLockOptimistic
    rsUpload.CursorType = adOpenStatic
    rsUpload.Open
    
    '��¼�Թܱ���
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "����", adVarChar, 18
    rsNumber.Fields.Append "���ID", adBigInt
    rsNumber.Fields.Append "��������", adVarChar, 18
    rsNumber.Fields.Append "ִ�п���ID", adVarChar, 18
    rsNumber.Fields.Append "������ĿID", adVarChar, 18
    rsNumber.Fields.Append "Ӥ��", adBigInt
    rsNumber.Fields.Append "������־", adBigInt
    rsNumber.Fields.Append "�걾", adVarChar, 18
    rsNumber.Fields.Append "�ɼ�����ID", adBigInt
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
    '��ǰ���˱���Ҫ���͵ķ���
    Set rsMoneyNow = New ADODB.Recordset
    rsMoneyNow.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsMoneyNow.Fields.Append "������ĿID", adBigInt
    rsMoneyNow.Fields.Append "�շ���ĿID", adBigInt
    rsMoneyNow.Fields.Append "�Թܱ���", adVarChar, 18, adFldIsNullable
    rsMoneyNow.Fields.Append "�շѷ�ʽ", adInteger
    rsMoneyNow.Fields.Append "�շ�ʱ��", adVarChar, 10
    rsMoneyNow.Fields.Append "ִ�в���ID", adBigInt
    rsMoneyNow.CursorLocation = adUseClient
    rsMoneyNow.LockType = adLockOptimistic
    rsMoneyNow.CursorType = adOpenStatic
    rsMoneyNow.Open
    
    '��ǰ���˱��η��͵ķ�����Ŀ����
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "����ID", adBigInt
    rsItems.Fields.Append "��ҳID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "ҽ��ID", adBigInt
    rsItems.Fields.Append "�շ����", adVarChar, 1
    rsItems.Fields.Append "�շ�ϸĿID", adBigInt
    rsItems.Fields.Append "����", adDouble
    rsItems.Fields.Append "����", adDouble
    rsItems.Fields.Append "ʵ�ս��", adDouble
    rsItems.Fields.Append "������", adVarChar, 100, adFldIsNullable
    rsItems.Fields.Append "��������", adVarChar, 100, adFldIsNullable
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open
    
End Sub

Public Function GetTubeMaterial(ByVal str�Թܱ��� As String) As Long
'���ܣ����ݹ����ȡ��Ӧ���Թܲ���ID
    Dim StrSQL As String, rsTube As Recordset
    
    On Error GoTo errH
    
    StrSQL = "Select ����,����ID From ��Ѫ������ Where ����ID is Not NULL and ����=[1]"
    Set rsTube = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "GetTubeMaterial", str�Թܱ���)
    
    If Not rsTube.EOF Then GetTubeMaterial = Nvl(rsTube!����ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get�շ�ִ�п���ID(ByVal lng����ID As Long, lng��ҳID As Long, _
    ByVal str��� As String, ByVal lng��ĿID As Long, ByVal intִ�п��� As Integer, _
    ByVal lng���˿���ID As Long, ByVal lng��������ID As Long, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal lngִ�п���ID As Long, _
    Optional ByVal bytMode As Byte, Optional ByVal bytCallBy As Byte, _
    Optional ByVal int���ó��� As Integer = 1, _
    Optional lng����ȱʡִ�п��� As Long = 0) As Long
'���ܣ���ȡ��ҩ�շ���Ŀ��ִ�п���
'������int��Χ=1.����,2-סԺ
'      lngִ�п���ID=ָ����ȱʡִ�п���ID(����ҩƷ������)
'      bytMode=1-Ҫ����ȱʡֵ,0-����
'      bytCallBy=0-ҽ���������,1-���ѳ������
'      int���ó���=1-����,2-סԺ
'      lng����ȱʡִ�п���-ȱʡִ�п���ID
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Integer
    Dim strҩ�� As String, lngҩ�� As Long
    Dim lng���˲���ID As Long, bytDay As Byte
    
    On Error GoTo errH
    
    If str��� = "4" Then
        lngҩ�� = Val(gobjComlib.zlDatabase.GetPara(IIF(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ���ϲ���", glngSys, _
            IIF(bytCallBy = 1, pҽ�����ѹ���, IIF(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�))))
        
        '��ִ�п�������ʱ
        StrSQL = _
            " Select Distinct" & _
            "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
            " And B.������� IN([1],3) And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And (A.������Դ is NULL Or A.������Դ=[1])" & _
            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And A.�շ�ϸĿID=[3]" & _
            " Order by B.�������,C.����"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", int��Χ, lng���˿���ID, lng��ĿID)
        If Not rsTmp.EOF Then
            If bytMode = 1 Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID  '�����û�У��򷵻ص�һ�����õ�ִ�п���
            
            '1:ȱʡΪָ����(ҽ����)ִ�п���,�����Ƿ�����ڲ��˿���
            rsTmp.Filter = "ִ�п���ID=" & lngִ�п���ID
            
            '2.ȱʡΪ����ָ����ȱʡ����
            If rsTmp.EOF Then rsTmp.Filter = "ִ�п���ID=" & lngҩ��
            
            '3:�����ɷ����ڲ��˿��ҵ�ִ�п���
            If rsTmp.EOF Then
                '2.0 ��������д���ȱʡ��ִ�п���,��ȱʡΪ����ָ����ȱʡ����
                If lng����ȱʡִ�п��� <> 0 Then
                    rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                    If Not rsTmp.EOF Then
                            Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                    End If
                End If
                '2.1:����ȱʡΪ���˿���
                If lngִ�п���ID <> lng���˿���ID And lngҩ�� <> lng���˿���ID Then
                    rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lng���˿���ID
                End If
                '3.2:����ȱʡΪ���˲���
                If rsTmp.EOF And lng��ҳID <> 0 Then
                    lng���˲���ID = GetPatiUnitID(lng����ID, lng��ҳID)
                    If lng���˲���ID <> 0 And lng���˲���ID <> lng���˿���ID And lng���˲���ID <> lngִ�п���ID And lng���˲���ID <> lngҩ�� Then
                        rsTmp.Filter = "��������ID=" & lng���˿���ID & " And ִ�п���ID=" & lng���˲���ID
                    End If
                End If
            End If
            '3.3:�ɷ����ڲ��˿��ҵ�һ��ִ�п���
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            
            '3.4�ɷ��������п��ҵĵ�ǰ���˿���ִ��
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=0 And ִ�п���ID=" & lng���˿���ID
            
            '4:�����û�У��򷵻�0���ڼ��
            If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
        End If
    ElseIf InStr(",5,6,7,", str���) > 0 Then
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(gobjComlib.zlDatabase.GetPara(IIF(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, _
                IIF(bytCallBy = 1, pҽ�����ѹ���, IIF(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�)), , , , , lng���˿���ID))
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(gobjComlib.zlDatabase.GetPara(IIF(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, _
                IIF(bytCallBy = 1, pҽ�����ѹ���, IIF(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�)), , , , , lng���˿���ID))
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(gobjComlib.zlDatabase.GetPara(IIF(int��Χ = 2 Or int���ó��� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, _
                IIF(bytCallBy = 1, pҽ�����ѹ���, IIF(int��Χ = 2 Or int���ó��� = 2, pסԺҽ���´�, p����ҽ���´�)), , , , , lng���˿���ID))
        End If
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If Not Check�ϰల��(True) Then
            StrSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                " And A.�շ�ϸĿID=[4]" & _
                " Order by B.�������,C.����"
        Else
            bytDay = Weekday(gobjComlib.zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            StrSQL = _
                " Select Distinct" & _
                "   B.�������,C.����,Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
                " And B.������� IN([2],3) And B.����ID=C.ID" & _
                " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                " And D.����ID=C.ID And D.����=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                " And (A.������Դ is NULL Or A.������Դ=[2])" & _
                " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                " And A.�շ�ϸĿID=[4]" & _
                " Order by B.�������,C.����"
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", strҩ��, int��Χ, lng���˿���ID, lng��ĿID, bytDay)
        If Not rsTmp.EOF Then
            If lng����ȱʡִ�п��� <> 0 Then
                rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                If Not rsTmp.EOF Then
                        Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                End If
            End If
            Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
            rsTmp.Filter = "ִ�п���ID=" & lngִ�п���ID
            If rsTmp.EOF Then rsTmp.Filter = "ִ�п���ID=" & lngҩ��
            If rsTmp.EOF Then rsTmp.Filter = "��������ID=" & lng���˿���ID
            If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
        End If
    Else
        Select Case intִ�п���
            Case 0 '0-����ȷ����
                Get�շ�ִ�п���ID = Get����Ա����ID(int��Χ)
            Case 1 '1-�������ڿ���
                Get�շ�ִ�п���ID = lng���˿���ID
            Case 2 '2-�������ڲ���
                If int��Χ = 1 Then
                    Get�շ�ִ�п���ID = lng���˿���ID
                Else
                    Get�շ�ִ�п���ID = GetPatiUnitID(lng����ID, lng��ҳID)
                End If
            Case 3 '3-����Ա���ڿ���
                Get�շ�ִ�п���ID = Get����Ա����ID(int��Χ)
            Case 4 '4-ָ������
                StrSQL = "Select Distinct Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID,Decode(A.������Դ,Null,2,1) as ����" & _
                    " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                    " Where A.�շ�ϸĿID=[1] And A.ִ�п���ID=B.����ID" & _
                    " And B.������� IN([2],3) And (A.������Դ is NULL Or A.������Դ=[2])" & _
                    " And (A.��������ID is NULL Or A.��������ID=[3])" & _
                    " And A.ִ�п���ID=C.ID And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                    " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " Order by ����" 'Ĭ�Ͽ�������
                Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lng��ĿID, int��Χ, lng���˿���ID)
                If Not rsTmp.EOF Then
                    If lng����ȱʡִ�п��� <> 0 Then
                         rsTmp.Filter = "ִ�п���ID=" & lng����ȱʡִ�п���
                         If Not rsTmp.EOF Then
                                 Get�շ�ִ�п���ID = rsTmp!ִ�п���ID: Exit Function
                         End If
                     End If
                    Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                    rsTmp.Filter = "��������ID=" & lng���˿���ID
                    If Not rsTmp.EOF Then Get�շ�ִ�п���ID = rsTmp!ִ�п���ID
                End If
            Case 6 '6-���������ڿ���
                Get�շ�ִ�п���ID = lng��������ID
        End Select
        If Get�շ�ִ�п���ID = 0 Then Get�շ�ִ�п���ID = Get����Ա����ID(int��Χ)
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function CalcDrugPrice(ByVal lngҩƷID As Long, lngҩ��ID As Long, ByVal dbl���� As Double, _
    Optional ByVal str�ѱ� As String, Optional ByVal blnNone�Ӱ�Ӽ� As Boolean, Optional ByVal strҩƷ�۸�ȼ� As String, Optional ByVal str���ļ۸�ȼ� As String, Optional ByVal str��ͨ��Ŀ�۸�ȼ� As String) As Double
'���ܣ�����ҩƷʵ��(��ȻҪ����ʵ��,ҩƷ��϶�Ϊ���)������ѱ�ʱ�������ʵ�ս��
'������dbl����=�ۼ�����,���ѱ����ʱ�������ʵ�ս��
'      str�ѱ�=�Ƿ񰴷ѱ������۵ļ۸�,��Ҫ��ֱ�Ӽ���ҩƷ�Ľ�������ʾ����ʱ��
'      gbln�Ӱ�Ӽ�=����ʱ���������,�����ط���ΪFalse
'      blnNone�Ӱ�Ӽ�=Ϊ��ʱ,"gbln�Ӱ�Ӽ�"��Ч
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long
    Dim dbl������ As Double, dbl��ǰ���� As Double
    Dim dbl�ܽ�� As Double, dblʱ�� As Double
    Dim dbl�Ӱ�Ӽ��� As Double
    Dim dbl����ʱ�� As Double, intCount As Integer
        
    If dbl���� = 0 Then Exit Function
    
    On Error GoTo errH
    
    StrSQL = _
        " Select Nvl(����,0) as ����,Nvl(��������,0) as ���," & _
        " Nvl(���ۼ�,Nvl(Decode(Nvl(ʵ������,0),0,0,ʵ�ʽ��/ʵ������),0)) as ʱ��" & _
        " From ҩƷ���" & _
        " Where �ⷿID=[1] And ҩƷID=[2] And Nvl(��������,0)>0" & _
        " And ����=1 And (Nvl(����,0)=0 Or Ч�� is NULL Or Ч��>Trunc(Sysdate))" & _
        " Order by " & IIF(gbytMediOutMode = 1, "Ч��,", "") & "Nvl(����,0)"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "CalcDrugPrice", lngҩ��ID, lngҩƷID)
    
    dbl�ܽ�� = 0: dbl������ = dbl����
    For i = 1 To rsTmp.RecordCount
        '��һ�����ε�ʱ��
        intCount = intCount + 1
        If intCount = 1 Then
            dbl����ʱ�� = Format(rsTmp!ʱ��, gstrDecPrice)
        End If
        If dbl������ = 0 Then Exit For 'Ϊ��ʼ��ȡ������ʱ��
        
        If dbl������ <= rsTmp!��� Then
            dbl��ǰ���� = dbl������
        Else
            dbl��ǰ���� = rsTmp!���
        End If
        dbl�ܽ�� = dbl�ܽ�� + Format(dbl��ǰ���� * Format(rsTmp!ʱ��, gstrDecPrice), gstrDec)
        dbl������ = Val(dbl������) - Val(dbl��ǰ����)
        If dbl������ = 0 Then Exit For
        
        rsTmp.MoveNext
    Next
    
    If dbl������ <> 0 Then
        '��治��,ֻ�漰һ������ʱ������ʱ��Ϊ׼�������Ե�һ������ƽ���۶�������
        dblʱ�� = IIF(intCount = 1, dbl����ʱ��, 0)
    Else
        dblʱ�� = IIF(intCount = 1, dbl����ʱ��, Format(dbl�ܽ�� / dbl����, gstrDecPrice))
        
        '���зѱ����ʱ���ǽ�������������ʵ�ս��
        If str�ѱ� <> "" Then
            dblʱ�� = Format(dblʱ�� * dbl����, gstrDec)
            
            StrSQL = _
                " Select A.���ηѱ�,B.������ĿID" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B" & _
                " Where A.ID=B.�շ�ϸĿID And A.ID=[1]" & _
                GetPriceGradeSQL(strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ��Ŀ�۸�ȼ�, "A", "B", "2", "3", "4") & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))"
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "CalcDrugPrice", lngҩƷID, strҩƷ�۸�ȼ�, str���ļ۸�ȼ�, str��ͨ��Ŀ�۸�ȼ�)
            If rsTmp.EOF Then Exit Function
            
            '���ݷѱ����¼���ʵ�ս��
            If Not (Nvl(rsTmp!���ηѱ�, 0) = 1) Then
                dblʱ�� = ActualMoney(str�ѱ� & IIF(gstr��̬�ѱ� <> "", "," & gstr��̬�ѱ�, ""), rsTmp!������ĿID, dblʱ��, lngҩƷID, lngҩ��ID, dbl����)
            End If
        End If
    End If
    CalcDrugPrice = dblʱ��
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Calc�����ֽ�ʱ��(lng���� As Long, ByVal dat��ʼʱ�� As Date, dat��ֹʱ�� As Date, strPause As String, _
    ByVal strִ��ʱ�� As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, _
    Optional ByVal dat�������� As Date) As String
'���ܣ�������������εķֽ�ִ��ʱ��,Ҫ��<=��ֹʱ�估������ͣʱ�����
'������dat��ʼʱ��=ҽ���Ŀ�ʼִ��ʱ��
'      dat��ֹʱ��=ҽ����ִ����ֹʱ��,û��ʱ����"3000-01-01"
'      strPause=ҽ������ͣʱ���
'      dat��������=��������ʱ��������
'���أ�1."ʱ��1,ʱ��2,...."(yyyy-MM-dd HH:mm:ss)
'      2.lng����=ʵ���ܹ��ֽ�Ĵ���
'˵����1.��Ϊ��ֹʱ�������,��˷ֽ������ʱ���������С��Ҫ�ֽ�Ĵ���
'      2.�������Ǽٶ���ִ��ʱ�估Ƶ��������ȫ��ȷ������¼��㡣
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrFirst As Variant, arrNormal As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    If InStr(strִ��ʱ��, ",") > 0 Then
        arrNormal = Split(Split(strִ��ʱ��, ",")(1), "-")
        arrFirst = Split(Split(strִ��ʱ��, ",")(0), "-")
    Else
        arrNormal = Split(strִ��ʱ��, "-")
        arrFirst = Array()
    End If
    
    vCurTime = dat��ʼʱ��
    
    If str�����λ = "��" Then
        vCurTime = gobjComlib.ZLCommFun.GetWeekBase(dat��ʼʱ��)
        
        Do While lng���� > 0
            blnFirst = (gobjComlib.ZLCommFun.GetWeekBase(vCurTime) = gobjComlib.ZLCommFun.GetWeekBase(dat��������)) And dat�������� <> Empty And UBound(arrFirst) <> -1
            arrTime = IIF(blnFirst, arrFirst, arrNormal)

            '1/8:00-3/15:00-5/9:00
            For i = 1 To intƵ�ʴ���
                If i - 1 <= UBound(arrTime) Then '���ܿ��ܴ�������
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime > dat��ֹʱ�� Then
                        Exit Do
                    ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                        Exit Do
                    ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        lng���� = lng���� - 1
                        If lng���� = 0 Then Exit Do
                    End If
                End If
            Next
            vCurTime = vCurTime + 7
        Loop
    ElseIf str�����λ = "��" Then
        Do While lng���� > 0
            blnFirst = (Int(vCurTime) = Int(dat��������)) And dat�������� <> Empty And UBound(arrFirst) <> -1
            arrTime = IIF(blnFirst, arrFirst, arrNormal)
        
            If intƵ�ʼ�� = 1 Then
                '8:00-12:00-14:00��8-12-14
                For i = 1 To intƵ�ʴ���
                    If i - 1 <= UBound(arrTime) Then '���տ��ܴ�������
                        If InStr(arrTime(i - 1), ":") = 0 Then
                            strTmp = arrTime(i - 1) & ":00"
                        Else
                            strTmp = arrTime(i - 1)
                        End If
                        vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        
                        If vTmpTime > dat��ֹʱ�� Then
                            Exit Do
                        ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                            Exit Do
                        ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            lng���� = lng���� - 1
                            If lng���� = 0 Then Exit Do
                        End If
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To intƵ�ʴ���
                    If i - 1 <= UBound(arrTime) Then '���տ��ܴ�������
                        vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                        If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                            strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                        Else
                            strTmp = Split(arrTime(i - 1), "/")(1)
                        End If
                        vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime > dat��ֹʱ�� Then
                            Exit Do
                        ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                            Exit Do
                        ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            lng���� = lng���� - 1
                            If lng���� = 0 Then Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = vCurTime + intƵ�ʼ��
        Loop
    ElseIf str�����λ = "Сʱ" Then
        '10:00-20:00-40:00��10-20-40��02:30
        arrTime = arrNormal
        Do While lng���� > 0
            For i = 1 To intƵ�ʴ���
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime > dat��ֹʱ�� Then
                    Exit Do
                ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                    Exit Do
                ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    lng���� = lng���� - 1
                    If lng���� = 0 Then Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + intƵ�ʼ�� / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str�����λ = "����" Then
        '��ִ��ʱ��
        Do While lng���� > 0
            vTmpTime = vCurTime
            
            If vTmpTime > dat��ֹʱ�� Then
                Exit Do
            ElseIf TimeisLastPause(vTmpTime, strPause) And dat��ֹʱ�� = CDate("3000-01-01") Then
                Exit Do
            ElseIf vTmpTime >= dat��ʼʱ�� And Not TimeIsPause(vTmpTime, strPause) Then
                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                lng���� = lng���� - 1
                If lng���� = 0 Then Exit Do
            End If

            vCurTime = Format(vCurTime + intƵ�ʼ�� / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If

    lng���� = UBound(Split(Mid(strDetailTime, 2), ",")) + 1
    Calc�����ֽ�ʱ�� = Mid(strDetailTime, 2)
End Function

Public Function AdviceMoneyMake(ByVal lng����ID As Long, ByVal lng��ҳID As Long, rsMoneyNow As Recordset, rsMoneyDay As ADODB.Recordset, _
    ByVal lngҽ��ID As Long, ByVal lng������ĿID, ByVal lng�շ���ĿID As Long, ByVal lngִ�в���id As Long, ByVal str�Թܱ��� As String, _
    ByVal str�շ���� As String, ByVal int�շѷ�ʽ As Integer, ByVal str�ֽ�ʱ�� As String, ByVal byt��Դ As Byte, ByRef lng���ô��� As Long, ByVal dbl���� As Double, _
    Optional ByVal lng��ǰҽ��ID As Long, Optional ByVal lng���ͺ� As Long, Optional ByVal dbl�Ƽ����� As Double, Optional rsExec As Recordset, _
    Optional ByVal lng���㷽ʽ As Long, Optional ByVal strƵ�� As String, Optional ByVal dbl���� As Double, Optional ByVal int��Ч As Integer = 1, _
    Optional ByVal int�������� As Integer, Optional ByVal str������� As String, Optional ByVal str�������� As String) As Boolean
'���ܣ��ж�ָ����ҽ�������Ƿ�Ӧ�ò���
'������lng��ҳID=סԺ���˲�ʹ�ã����ﲡ�˴���0���־���Һ�
'      rsMoneyNow=��ǰ���˱���Ҫ���͵ķ���,��̬��¼��(�շѷ�ʽ=-1,��ʾ�״β���ʱ��һ��ֻ��һ�ε���Ŀ�ļ�¼)
'      rsMoneyDay=��ǰ���˵����ѷ��͵ķ���,��̬��¼��
'      lngҽ��ID=һ��ҽ����ID
'      str�ֽ�ʱ��=���η��͵�ִ��ʱ�䴮���Զ��ŷָ��������ų�����ͣ��ʱ���
'      byt��Դ:1-���2-סԺ
'      dbl�Ƽ�����=�շ���Ŀ�ļƼ�����
'      ����=��ǰ�з���ҽ����������Ϣ
'      lng��ǰҽ��ID=��ǰ��ҽ��id
'      str��������=����ҽ��������������
'�����Ǽ�����ʱҽ����������֯����
'1��������ѡƵ�ʡ������ԡ���Ҫʱ�Ͳ���ʱ�Ե�����Ϊ���Ρ�
'2������һ���Ժ���ҪʱƵ�ʵ�ҽ��ȡ������Ϊ���Ρ�
'3��������ѡƵ��ȡ������Ϊ���Σ����һ��ȡ�������Ե���ȡĩ��Ϊ���Σ���������������ƣ�����80������25��ÿ��4�Σ���ôִ�еǼ�ʱ����ִ���ĴΣ�ǰ���α�������Ϊ25�����Ĵ�Ϊ80����25ȡģ=5��
'4������ִ�еǼ�ҳ��ҽ���嵥�����������У��������Σ�������ʾ�������Ρ�
'5��ҽ���༭ʱ������¼���״�������
'���أ�
'      lng���ô���=һ��ֻ��һ��ʱ��3,4,5,6,7�������ر��η���Ҫ��ȡ�Ĵ���
'      dbl����=�ܵķ��ʹ���������
'      rsExec=ҽ��ִ�мƼ۵�����
    Dim lng����ID As Long, blnMakeMoney As Boolean
    Dim rsDays As ADODB.Recordset, i As Long
    Dim arrTmp As Variant
    Dim dbl���� As Double
    Dim strDate As String
    Dim dbl����Tmp As Double
    Dim StrSQL As String, rsTmp As Recordset, strTmp As String
    
    blnMakeMoney = True
    lng���ô��� = 1
    
    If int�շѷ�ʽ = 9 Then
        '�Զ���
        On Error GoTo errH
        
        StrSQL = "Select zl_fun_CustomExpenses([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17]) as ���ؽ�� From Dual"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "AdviceMoneyMake", lng����ID, lng��ҳID, byt��Դ, lng��ǰҽ��ID, lngҽ��ID, int��Ч, strƵ��, lng������ĿID, lng�շ���ĿID, _
                                            lngִ�в���id, str�������, str�շ����, dbl����, dbl����, dbl�Ƽ�����, int��������, lng���㷽ʽ)
        If rsTmp.RecordCount > 0 Then
            strTmp = rsTmp!���ؽ�� & ""
            
            If Val(Split(strTmp, ":")(0)) = 0 Then
                '����ȡ
                blnMakeMoney = False
            Else
                'Ҫ��ȡ
                If InStr(strTmp, ":") > 0 Then
                    If Val(Split(strTmp, ":")(1)) > 0 Then lng���ô��� = Val(Split(strTmp, ":")(1))
                End If
            End If
        End If
    End If
    
    If int�շѷ�ʽ = 0 Then
        '�����շѵģ�����ڱ��η����С���ҽ�����Ƿ��ų�
        rsMoneyNow.Filter = "(ҽ��ID=" & lngҽ��ID & " And ������ĿID=" & lng������ĿID & " And �շѷ�ʽ=5)" & _
            " Or (ҽ��ID=" & lngҽ��ID & " And ������ĿID=" & lng������ĿID & " And �շѷ�ʽ=6)" 'Or��ʹ��
        If Not rsMoneyNow.EOF Then blnMakeMoney = False
    ElseIf int�շѷ�ʽ = 1 Then '�����Թܷ���(һ�η���ֻ��ȡһ��)
        If str�Թܱ��� <> "" Then
            '��ͬ����(�Թ�)ֻ��ȡһ��
            rsMoneyNow.Filter = "�Թܱ���='" & str�Թܱ��� & "' And ��������='" & str�������� & "' And �շ���ĿID=" & lng�շ���ĿID & " And �շѷ�ʽ<>-1"
            If Not rsMoneyNow.EOF Then blnMakeMoney = False
            
            'ֻ��ȡ�Թܶ�Ӧ�����ķ���
            If blnMakeMoney And str�շ���� = "4" Then
                lng����ID = GetTubeMaterial(str�Թܱ���)
                If lng����ID <> 0 And lng�շ���ĿID <> lng����ID Then blnMakeMoney = False
            End If
        End If
    ElseIf int�շѷ�ʽ = 2 Then 'һ�η���ֻ��ȡһ��
        rsMoneyNow.Filter = "������ĿID=" & lng������ĿID & " And �շ���ĿID=" & lng�շ���ĿID & " And �շѷ�ʽ<>-1"
        If Not rsMoneyNow.EOF Then blnMakeMoney = False
    ElseIf InStr(",3,4,5,6,7,", int�շѷ�ʽ) > 0 Then
        '3-����ֻ��ȡһ�Σ�4-����δִ����ȡһ�Σ�5-����ֻ��ȡһ�Σ��ų�������Ŀ��6-����δִ����ȡһ�Σ��ų�������Ŀ
        
        '�����շѵģ�����ڱ��η����С���ҽ�����Ƿ��ų�
        If int�շѷ�ʽ = 7 Then
            rsMoneyNow.Filter = "(ҽ��ID=" & lngҽ��ID & " And ������ĿID=" & lng������ĿID & " And �շѷ�ʽ=5)" & _
                " Or (ҽ��ID=" & lngҽ��ID & " And ������ĿID=" & lng������ĿID & " And �շѷ�ʽ=6)" 'Or��ʹ��
            If Not rsMoneyNow.EOF Then blnMakeMoney = False
        End If
        
        If blnMakeMoney Then
            Set rsDays = GetExecDays(str�ֽ�ʱ��)
                        
            '�ȴӱ��η����е���(Ƶ��Ϊһ��һ����û���յģ��ж�ʱ��������ȡ,�Ա����������ҽ��"�״β���"ʱ������Ϊ���״�)
            For i = 1 To rsDays.RecordCount
                rsMoneyNow.Filter = "�շ�ʱ��='" & rsDays!�շ�ʱ�� & "' And ������ĿID=" & lng������ĿID & " And �շ���ĿID=" & lng�շ���ĿID & _
                    IIF(int�շѷ�ʽ = 7, "", " And �շѷ�ʽ<>-1") & _
                    IIF((int�շѷ�ʽ = 4 Or int�շѷ�ʽ = 6) And lngִ�в���id <> 0, " And ִ�в���ID=" & lngִ�в���id, "")
                If rsMoneyNow.RecordCount > 0 Then rsDays!���� = 1
                rsDays.MoveNext
            Next
            '�ٴ��ѷ����е���(���켰����ִ�е�)
            rsDays.Filter = "����=0"
            For i = 1 To rsDays.RecordCount
                If i = 1 Then
                    If rsMoneyDay Is Nothing Then
                        Call GetPatiDayMoneyDetail(rsMoneyDay, lng����ID, lng��ҳID, byt��Դ, CDate(rsDays!�շ�ʱ�� & ""))
                    End If
                End If
                rsMoneyDay.Filter = "�շ�ʱ��='" & rsDays!�շ�ʱ�� & "' And ������ĿID=" & lng������ĿID & " And �շ���ĿID=" & lng�շ���ĿID & _
                    IIF(int�շѷ�ʽ = 7, "", " And �շѷ�ʽ<>-1") & _
                    IIF((int�շѷ�ʽ = 4 Or int�շѷ�ʽ = 6) And lngִ�в���id <> 0, " And ִ�з�=0 And ִ�в���ID=" & lngִ�в���id, "")
                If rsMoneyDay.RecordCount > 0 Then rsDays!���� = 1
                rsDays.MoveNext
            Next
        End If
    End If
                            
    '��¼�����η�����ϸ��Ŀ��¼��
    If InStr(",3,4,5,6,7,", int�շѷ�ʽ) > 0 Then
        If int�շѷ�ʽ = 7 Then
            If blnMakeMoney Then
                rsDays.Filter = "����=0"    'û�չ�����Щ��(Ƶ��Ϊһ��һ�ε�δ�յĵ����չ���)���״β���
                lng���ô��� = dbl���� - rsDays.RecordCount
                blnMakeMoney = lng���ô��� > 0
            End If
        Else
            rsDays.Filter = "����=0"
            blnMakeMoney = rsDays.RecordCount > 0
            lng���ô��� = rsDays.RecordCount    'һ��һ�Σ��ж�����Ҫ�վ��ж��ٴ�
        End If
        If blnMakeMoney Or int�շѷ�ʽ = 7 And lng���ô��� = 0 Then
            For i = 1 To rsDays.RecordCount
                rsMoneyNow.AddNew
                rsMoneyNow!ҽ��ID = lngҽ��ID
                rsMoneyNow!������ĿID = lng������ĿID
                rsMoneyNow!�շ���ĿID = lng�շ���ĿID
                rsMoneyNow!�Թܱ��� = str�Թܱ���
                rsMoneyNow!�������� = str��������
                
                '�״β���ʱ�����Ƶ��Ϊһ��һ�Σ�������ķ��ô���Ϊ0,Ϊ���ñ��κ������͵�����ҽ����ȷ�������Ƿ���ȡ����Ҫ������¼�����շѷ�ʽ�����¼Ϊ-1
                rsMoneyNow!�շѷ�ʽ = IIF(int�շѷ�ʽ = 7 And lng���ô��� = 0, -1, int�շѷ�ʽ)
                rsMoneyNow!�շ�ʱ�� = rsDays!�շ�ʱ��
                rsMoneyNow!ִ�в���ID = lngִ�в���id
                rsMoneyNow.Update
            
                rsDays.MoveNext
            Next
        End If
    ElseIf blnMakeMoney Then
        rsMoneyNow.AddNew
        rsMoneyNow!ҽ��ID = lngҽ��ID
        rsMoneyNow!������ĿID = lng������ĿID
        rsMoneyNow!�շ���ĿID = lng�շ���ĿID
        rsMoneyNow!�Թܱ��� = str�Թܱ���
        rsMoneyNow!�������� = str��������
        rsMoneyNow!�շѷ�ʽ = int�շѷ�ʽ
        If str�ֽ�ʱ�� <> "" Then
            rsMoneyNow!�շ�ʱ�� = Format(Split(str�ֽ�ʱ��, ",")(0), "yyyy-MM-dd")  '��ʱ����ʱû���ô�
        Else
            rsMoneyNow!�շ�ʱ�� = ""
        End If
        rsMoneyNow!ִ�в���ID = lngִ�в���id
        rsMoneyNow.Update
    End If
    '��ȡҽ��ִ�мƼ�(��ҩƷ����ҽ����ĲŴ洢)
    If InStr(",5,6,7,", "," & str������� & ",") = 0 Then
        If str�ֽ�ʱ�� <> "" And Not rsExec Is Nothing Then
            arrTmp = Split(str�ֽ�ʱ��, ",")
            dbl����Tmp = dbl����
            For i = 0 To UBound(arrTmp)
                rsExec.AddNew
                rsExec!ҽ��ID = lng��ǰҽ��ID
                rsExec!���ͺ� = lng���ͺ�
                rsExec!Ҫ��ʱ�� = Format(arrTmp(i), "yyyy-MM-dd HH:mm:ss")
                rsExec!�շ�ϸĿID = lng�շ���ĿID
                rsExec!�������� = int��������
                If blnMakeMoney Then
                    '����Ҳ�������뵥������
                    If strƵ�� <> "" And (lng���㷽ʽ = 0 And dbl���� > 0 Or lng���㷽ʽ = 1 Or lng���㷽ʽ = 2 Or str������� = "4") Then
                        '�����ͼ�ʱ����Ҫ��������
                        If int��Ч = 0 Then
                            '1��������ѡƵ�ʡ������ԡ���Ҫʱ�Ͳ���ʱ�Ե�����Ϊ���Ρ�
                            dbl���� = dbl�Ƽ����� * dbl����
                        ElseIf InStr("һ����,��Ҫʱ", strƵ��) Then
                            '2������һ���Ժ���ҪʱƵ�ʵ�ҽ��ȡ������Ϊ���Ρ�
                            dbl���� = dbl�Ƽ����� * dbl����
                        Else
                            '3��������ѡƵ��ȡ������Ϊ���Σ����һ��ʣ�����������������������ƣ�����80������25��ÿ��4�Σ���ôִ�еǼ�ʱ����ִ���ĴΣ�ǰ���α�������Ϊ25�����Ĵ�Ϊ80����25ȡģ=5��
                            '�����п���û��¼��ִ��ʱ��,�ֽ�ʱ���ֻ��һ������������Ϊ����
                            If UBound(arrTmp) = 0 Then
                                dbl���� = dbl�Ƽ����� * dbl����
                            Else
                                If i = UBound(arrTmp) Then
                                    dbl���� = dbl����Tmp
                                Else
                                    If dbl����Tmp >= dbl���� Then
                                        dbl���� = dbl�Ƽ����� * dbl����
                                    Else
                                        dbl���� = dbl����Tmp
                                    End If
                                    dbl����Tmp = dbl����Tmp - dbl����
                                End If
                            End If
                        End If
                    Else
                        dbl���� = dbl�Ƽ�����
                    End If
                    If i <> 0 Then
                        strDate = Format(arrTmp(i - 1), "yyyy-MM-dd")
                    End If
                    'һ�η�����ȡһ�Σ���ֻ�е�һ����ȡ
                    If InStr(",1,2,", int�շѷ�ʽ) > 0 Then
                        If i <> 0 Then dbl���� = 0
                    ElseIf InStr(",3,4,5,6,", int�շѷ�ʽ) > 0 Then
                        '3456����ֻ��ȡһ�εģ�����=0����ȡ��Ĭ�ϵ�һ��������
                        rsDays.Filter = "����=0 And �շ�ʱ��='" & Format(arrTmp(i), "yyyy-MM-dd") & "'"
                        If Not (rsDays.RecordCount > 0 And Format(arrTmp(i), "yyyy-MM-dd") <> strDate) Then
                            dbl���� = 0
                        End If
                    ElseIf int�շѷ�ʽ = 7 Then
                        '�����״β���ȡ�ģ�����=1����ȡ������=0��Ϊ�״�
                        rsDays.Filter = "����=1 And �շ�ʱ��='" & Format(arrTmp(i), "yyyy-MM-dd") & "'"
                        If rsDays.RecordCount = 0 And Format(arrTmp(i), "yyyy-MM-dd") <> strDate Then
                            dbl���� = 0
                        End If
                    End If
                Else
                    '�������ȡ��������Ϊ0
                    dbl���� = 0
                End If
                rsExec!���� = dbl����
                rsExec.Update
            Next
        End If
    End If
    AdviceMoneyMake = blnMakeMoney
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub GetCurBillSet(rsBill As Recordset, ByVal strKey As String, strNO As String, lng������� As Long, lng������� As Long, bln�շѵ� As Boolean, ByRef lngNOSequence As Long)
'���ܣ���ȡ��ǰ���õ��ݵ�NO�����
'������lng�������=���ü�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'      lng�������=���ͼ�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'˵����strKey=���ݼ��ʵ������ɹ��򶨵�Ψһ�ؼ���
'1.������ҩ��"����(����ID,�Һŵ�)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
'2.һ���䷽�е����в�ҩ����һ���������ݺ�
'3.����ҽ�����ҩ�ֺŹ�����ͬ��
'4.������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷�)
'5.��鲿λ�͸�����������Ҫҽ��������ͬ���ݺţ�����������䵥���ĵ��ݺš�
'6.һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
    rsBill.Filter = "Key='" & strKey & "'"
    If rsBill.EOF Then
        rsBill.AddNew
        rsBill!Key = strKey
        
        'ȡ���ݺ�
        'rsBill!NO = gobjComlib.zldatabase.GetNextNo(IIF(bln�շѵ�, 13, 14)),������ʵ�Ҳ��14
        lngNOSequence = lngNOSequence + 1
        rsBill!NO = "TemporaryNO=" & IIF(bln�շѵ�, 13, 14) & Format(lngNOSequence, "00000")
        
        rsBill!������� = IIF(lng������� = -1, 0, 1)
        rsBill!������� = IIF(lng������� = -1, 0, 1)
        rsBill.Update
    Else
        If lng������� <> -1 Then
            rsBill!������� = rsBill!������� + 1
        End If
        If lng������� <> -1 Then
            rsBill!������� = rsBill!������� + 1
        End If
        rsBill.Update
    End If
    strNO = rsBill!NO
    If lng������� <> -1 Then lng������� = rsBill!�������
    If lng������� <> -1 Then lng������� = rsBill!�������
End Sub

Public Function GetAuditName(ByVal strName As String) As String
'���ܣ���"���ҽ��/ʵϰҽ��"��ȡ���ҽ����
    GetAuditName = Mid(strName, 1, IIF(InStr(strName, "/") > 0, InStr(strName, "/") - 1, Len(strName)))
End Function

Public Function GetPatiUnitID(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Long
'���ܣ����ݲ��˻�ȡ��Ӧ�Ĳ���ID
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    
    StrSQL = "Select ��ǰ����ID as ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    GetPatiUnitID = Nvl(rsTmp!����ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Load��̬�ѱ�(lng����id As Long) As String
'���ܣ�Ȩ��ָ�����Ҷ�ȡ��ǰ��Ч�Ķ�̬�ѱ�(Ŀǰֻ��������)
'���أ��ѱ�="���˽�,��һ��"
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, strTmp As String
    
    On Error GoTo errH
    
    StrSQL = _
        " Select ����,����,���� From �ѱ�" & _
        " Where Nvl(����,1)=2 And Nvl(���ÿ���,1)=1 And Nvl(�������,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(��Ч��ʼ,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(��Ч����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Union ALL" & _
        " Select Distinct A.����,A.����,A.����" & _
        " From �ѱ� A,�ѱ����ÿ��� B" & _
        " Where A.����=B.�ѱ� And B.����ID=[1]" & _
        " And Nvl(A.����,1)=2 And Nvl(A.���ÿ���,1)=2 And Nvl(A.�������,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(A.��Ч��ʼ,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(A.��Ч����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by ����"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Load��̬�ѱ�", lng����id)
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    Load��̬�ѱ� = Mid(strTmp, 2)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function ActualMoney(str�ѱ� As String, ByVal lng������ĿID As Long, ByVal curӦ�ս�� As Currency, _
    Optional ByVal lng�շ�ϸĿID As Long, Optional ByVal lng�ⷿID As Long, Optional ByVal dbl���� As Double, Optional ByVal dbl�Ӱ�Ӽ��� As Double) As Currency
'���ܣ������շ�ϸĿID��������ĿID(ǰ������),Ӧ�ս��,���ѱ����õķֶα������۹������ʵ�ս�
'       ���ҩƷ���ɱ����ձ����������ʵ�ս��
'������str�ѱ�=���˷ѱ�����ǰ���̬�ѱ�,�����ʽΪ"���˷ѱ�,��̬�ѱ�1,��̬�ѱ�2,..."
'      lng�ⷿID,dbl����,��ҩƷ����Ŀ���ɱ��ۼ��մ���ʱ����Ҫ����
'      dbl����=�����������ڵ��ۼ�����
'      dbl�Ӱ�Ӽ���=С������,�����Ӧ�ս���Ѱ��Ӱ�Ӽۼ���ʱ��Ҫ�����ڻ�ԭ������
'���أ������۹���ͱ��������ʵ�ս��,����Ƕ�̬�ѱ�,��"str�ѱ�"�������Żݷѱ�(ע�����δ���ۼ���,����ԭ������,Ҳ���ܷ��ص�һ��)
'˵����
'���ɱ��ۼ��ձ������۵����ּ��㷽��(ʵ����һ��)��
'1.���۽�� = �ɱ���� * (1 + ���ձ���)
'2.���۽�� = �ɱ��� * (1 + ���ձ���) * ��������
'��صļ��㹫ʽ��
'      �ɱ��� = ҩƷ�ۼ� * (1 - �����)
'      �ɱ���� = �ۼ۽�� * (1 - �����) = �ɱ��� * ��������
'      �п����ʱ:����� = ����� / �����,����:����� = ָ�������
'      ���ڷ���ҩƷ��Ӧÿ���������ηֱ����ɱ��ۺͳɱ����
'      ����ʱ�۷�����"ҩƷ�ۼ�=Nvl(���ۼ�,ʵ�ʽ��/ʵ������)"��������ʱ��ҩƷ��治��ʱ��������ۼ��㡣
    Dim rsTmp As ADODB.Recordset, StrSQL As String
    
    On Error GoTo errH
    StrSQL = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) as Actualmoney From Dual"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, App.ProductName, str�ѱ�, lng�շ�ϸĿID, lng������ĿID, curӦ�ս�� / (1 + dbl�Ӱ�Ӽ���), dbl����, lng�ⷿID)
        
    str�ѱ� = Split(rsTmp!ActualMoney, ":")(0)
    ActualMoney = Format(Split(rsTmp!ActualMoney, ":")(1) * (1 + dbl�Ӱ�Ӽ���), gstrDec)
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function TimeisLastPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ��ʱ���Ƿ������һ����ͣ��ʱ����,�����һ����ͣû������
'˵������Ϊ���������,�������û����ֹʱ��,ĳЩ�������ѭ��
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    
    For i = UBound(arrPause) To 0 Step -1
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then
            strEnd = "3000-01-01 00:00:00"
            If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
                TimeisLastPause = True: Exit Function
            End If
        End If
    Next
End Function

Public Function TimeIsPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ��ʱ���Ƿ�����ͣ��ʱ�����
'������strPause="��ͣʱ��,��ʼʱ��;...."
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then strEnd = "3000-01-01 00:00:00" '������δ���û���ͣ��ʱ��ֹͣ
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeIsPause = True: Exit Function
        End If
    Next
End Function

Public Function GetExecDays(ByVal str�ֽ�ʱ�� As String) As ADODB.Recordset
'���ܣ����ݵ�ǰҽ����ִ��ʱ�䴮���ز��ظ���ִ��������¼��
    Dim rsTmp As ADODB.Recordset
    Dim arrTmp As Variant, i As Long, strTmp As String
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "�շ�ʱ��", adVarChar, 10
    rsTmp.Fields.Append "����", adInteger '���ھ����Ƿ�����Ѵ��ڵ��б�
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    arrTmp = Split(str�ֽ�ʱ��, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = Format(arrTmp(i), "yyyy-MM-dd")
        rsTmp.Filter = "�շ�ʱ��='" & strTmp & "'"
        If rsTmp.EOF Then
            rsTmp.AddNew
            rsTmp!�շ�ʱ�� = strTmp
            rsTmp!���� = 0
            rsTmp.Update
        End If
    Next
    rsTmp.Filter = ""
    Set GetExecDays = rsTmp
End Function

Private Function GetPatiDayMoneyDetail(rsMoneyDay As ADODB.Recordset, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal byt��Դ As Byte, _
         Optional ByVal lng������ĿID As Long, Optional ByVal lng�շ�ϸĿID As Long, Optional ByVal date���ղ���ȡ As Date) As Boolean
'���ܣ���ȡָ�����˵��켰֮��ҽ�������ķ�����Ŀ��ϸ
'������lng��ҳID=סԺ���˲�ʹ��
'      byt��Դ:1-����(��סԺ�������͵�����)��2-סԺ
'      str�״�ʱ��=����ҽ�����ͣ��״�ִ�е�ʱ��
'      date���ղ���ȡ=����������ղ���ȡ����Ŀ��������Ƶ���ֲ���ÿ��һ�εģ�ʵ����ÿ��һ�εģ��������һ�Σ�ÿ24Сʱһ�ε�
'���أ�rsMoneyDay������"������ĿID,�շ���ĿID,ִ�в���ID,ִ�з�,�շ�ʱ��"�ֶ�
'      ����Ƿ��͵���֮ǰ��ҽ�����򱾹�����ʱû�п��������������鵱���Ƿ���ִ��ʱ���鲻��
    Dim StrSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long
    Dim strToDay As String, strDay As String
        
    On Error GoTo errH
    
    If lng������ĿID = 0 Then
        Set rsMoneyDay = New ADODB.Recordset '�������Filter����
        strToDay = Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd")
        'ִ���жϣ�
        '1.������ǽ�������ü�¼�е�ִ�в��ţ����Ҳ�Է��ü�¼�е�ִ�в���Ϊ׼�жϡ�
        '2.���͸��������⣬ҽ�����õ�ִ�п�����ҽ��ִ�п�����ͬ���Ժ������ͬ�ˣ��ú���Ҳ������Ӧ
        '3.ҽ��ִ��ʱ����Ӧ���õ�ִ��״̬Ҳ��ͬ����ǡ�
        '4.�״β��յ���Ŀ�����Ƶ����һ��ֻ��һ�Σ���û�в������ü�¼������ҽ�����ͼ�¼��,��Ҫ���������������ɵģ��Ա������״β��յ���Ŀ�ж�
        If byt��Դ = 1 Then
            StrSQL = "Select A.������ĿID,C.�շ�ϸĿID as �շ���ĿID,C.ִ�в���ID,Decode(Nvl(C.ִ��״̬,0),0,0,1) as ִ�з�,To_Char(C.����ʱ��,'yyyy-mm-dd') as �շ�ʱ��,0 as �շѷ�ʽ" & _
                " From ����ҽ����¼ A,����ҽ������ B,������ü�¼ C" & _
                " Where A.����ID=[1] And Nvl(A.��ҳID,0) = [2] And a.ҽ����Ч = 1 And A.ID=B.ҽ��ID And B.��¼����=C.��¼���� And B.NO=C.NO" & _
                " And B.ҽ��ID=C.ҽ����� And C.��¼״̬ IN(0,1) And C.����ʱ��>=[3]" & _
                " Union " & _
                " Select A.������ĿID,D.�շ�ϸĿid,D.ִ�п���ID as ִ�в���ID,0 as ִ�з�,To_Char(B.�״�ʱ��,'yyyy-mm-dd') as �շ�ʱ��,-1 as �շѷ�ʽ" & _
                " From ����ҽ����¼ A,����ҽ������ B,����ҽ���Ƽ� D" & _
                " Where A.����ID=[1] And Nvl(A.��ҳID,0) = [2] And a.ҽ����Ч = 1 " & _
                " And A.ID=B.ҽ��ID And NVL(B.�״�ʱ��,a.��ʼִ��ʱ��)>=[3] And A.ID=D.ҽ��ID And D.�շѷ�ʽ=7" & vbNewLine & _
                " And Not Exists (Select 1 From ������ü�¼ C Where c.�շ�ϸĿid=d.�շ�ϸĿid  And b.��¼���� = c.��¼���� And b.No = c.No And a.Id = c.ҽ�����)" & vbNewLine & _
                " Order by ������ĿID,�շ���ĿID"
            Set rsMoneyDay = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "��ȡ���켰������ҽ��", lng����ID, lng��ҳID, CDate(strToDay))
            Set rsMoneyDay = gobjComlib.zlDatabase.CopyNewRec(rsMoneyDay)
        Else
            '����������ҽ����¼.�ϴ�ִ��ʱ��Ϊ��
            '����������ҽ������ͬ���ã����ܲ�ͬʱ���η���,Unionȥ�����ظ���¼
            '�״β��յ���Ŀ�����Ƶ����һ��ֻ��һ�Σ���û�в������ü�¼������ҽ�����ͼ�¼��,��Ҫ���������������ɵģ��Ա������״β��յ���Ŀ�ж�
            StrSQL = "Select a.������Ŀid, c.�շ�ϸĿid As �շ���Ŀid, c.ִ�в���id, Decode(Nvl(c.ִ��״̬, 0), 0, 0, 1) As ִ�з�," & vbNewLine & _
                "     Decode(a.ҽ����Ч, 0, b.�״�ʱ��, c.����ʱ��) As �״�ʱ��, Decode(b.�״�ʱ��,null, 1,Trunc(b.ĩ��ʱ��) - Trunc(b.�״�ʱ��) + 1) As ����,0 as �շѷ�ʽ" & vbNewLine & _
                "From ����ҽ����¼ A, ����ҽ������ B, סԺ���ü�¼ C" & vbNewLine & _
                "Where a.����id = [1] And a.��ҳid = [2] And a.Id = b.ҽ��id And b.��¼���� = c.��¼���� And b.No = c.No And b.ҽ��id = c.ҽ����� And" & vbNewLine & _
                "      c.��¼״̬ In (0, 1) And ((b.�״�ʱ�� > [3] Or b.ĩ��ʱ�� > [3]) Or a.ҽ����Ч = 1 And C.����ʱ�� >= [3])" & vbNewLine & _
                " Union " & vbNewLine & _
                "Select a.������Ŀid, D.�շ�ϸĿid, D.ִ�п���ID as ִ�в���id, 0 As ִ�з�," & vbNewLine & _
                "     b.�״�ʱ��, Decode(a.ҽ����Ч, 0, Trunc(b.ĩ��ʱ��) - Trunc(b.�״�ʱ��) + 1, 1) As ����,-1 as �շѷ�ʽ" & vbNewLine & _
                "From ����ҽ����¼ A, ����ҽ������ B, ����ҽ���Ƽ� D" & vbNewLine & _
                "Where a.����id = [1] And a.��ҳid = [2]" & vbNewLine & _
                "   And a.Id = b.ҽ��id And ((b.�״�ʱ�� > [3] Or b.ĩ��ʱ�� > [3]) Or (a.ҽ����Ч = 1 And b.�״�ʱ�� is null and a.��ʼִ��ʱ�� >= [3]))" & vbNewLine & _
                "   And A.ID=D.ҽ��ID And D.�շѷ�ʽ=7" & vbNewLine & _
                " And Not Exists (Select 1 From סԺ���ü�¼ C Where c.�շ�ϸĿid=d.�շ�ϸĿid  And b.��¼���� = c.��¼���� And b.No = c.No And a.Id = c.ҽ�����)" & vbNewLine & _
                "Order By ������Ŀid, �շ���Ŀid"
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "��ȡ���켰������ҽ��", lng����ID, lng��ҳID, CDate(strToDay))
            '���ݿ�ʼʱ�������������¼����ִ��ʱ��ֳɶ�����¼
            Set rsMoneyDay = InitPatiExecDays
                    
            For i = 1 To rsTmp.RecordCount
                For j = 1 To rsTmp!����
                    If j = 1 Then
                        strDay = Format(rsTmp!�״�ʱ��, "yyyy-MM-dd")
                    Else
                        strDay = Format(DateAdd("d", j - 1, CDate(rsTmp!�״�ʱ��)), "yyyy-MM-dd")
                    End If
                    If strDay >= strToDay Then
                        rsMoneyDay.Filter = "������ĿID=" & Val("" & rsTmp!������ĿID) & " And �շ���ĿID=" & Val("" & rsTmp!�շ���ĿID) & _
                                            " And �շ�ʱ��='" & strDay & "' And ִ�з�=" & Val("" & rsTmp!ִ�з�) & " And �շѷ�ʽ=" & Val("" & rsTmp!�շѷ�ʽ)
                        If rsMoneyDay.RecordCount = 0 Then
                            rsMoneyDay.AddNew
                            rsMoneyDay!������ĿID = Val("" & rsTmp!������ĿID)
                            rsMoneyDay!�շ���ĿID = Val("" & rsTmp!�շ���ĿID)
                            rsMoneyDay!ִ�в���ID = Val("" & rsTmp!ִ�в���ID)
                            rsMoneyDay!ִ�з� = Val("" & rsTmp!ִ�з�)
                            rsMoneyDay!�շѷ�ʽ = Val("" & rsTmp!�շѷ�ʽ)
                            rsMoneyDay!�շ�ʱ�� = strDay
                            rsMoneyDay.Update
                        End If
                    End If
                Next
                rsTmp.MoveNext
            Next
            rsMoneyDay.Filter = ""
        End If
    Else
        '���﷢��ʱ�����ж�ÿ���״β���ȡ����Ŀ�����Ƿ�ִ�д���=1,���=1��û���շѣ�˵�������״��Ѿ�û����ȡ��
        StrSQL = "Select d.ִ�п���id As ִ�в���id" & vbNewLine & _
                "From ����ҽ����¼ A,����ҽ������ B, ����ҽ���Ƽ� D" & vbNewLine & _
                "Where A.����ID=[1] And Nvl(A.��ҳID,0) = [2] And a.Id = b.ҽ��id And A.id = d.ҽ��id And A.������ĿID = [6] And d.�շѷ�ʽ = 7 And d.�շ�ϸĿid = [3] And Not Exists" & vbNewLine & _
                " (Select 1" & vbNewLine & _
                "       From " & IIF(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼") & " C" & vbNewLine & _
                "       Where c.�շ�ϸĿid = d.�շ�ϸĿid And b.��¼���� = c.��¼���� And b.No = c.No And d.ҽ��id = c.ҽ�����) And" & vbNewLine & _
                "      Zl_Adviceexecount(d.ҽ��id, [4], [5],1) = 1"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "��ȡ���켰������ҽ��", lng����ID, lng��ҳID, lng�շ�ϸĿID, CDate(Format(date���ղ���ȡ, "yyyy-MM-dd")), CDate(Format(date���ղ���ȡ, "yyyy-MM-dd 23:59:59")), lng������ĿID)
        If rsTmp.RecordCount > 0 Then
            rsMoneyDay.Filter = "������ĿID=" & lng������ĿID & " And �շ���ĿID=" & lng�շ�ϸĿID & _
                                " And �շ�ʱ��='" & Format(date���ղ���ȡ, "yyyy-MM-dd") & "' And ִ�з�=0" & " And �շѷ�ʽ=-1"
            If rsMoneyDay.RecordCount = 0 Then
                rsMoneyDay.AddNew
                rsMoneyDay!������ĿID = lng������ĿID
                rsMoneyDay!�շ���ĿID = lng�շ�ϸĿID
                rsMoneyDay!ִ�в���ID = Val("" & rsTmp!ִ�в���ID)
                rsMoneyDay!ִ�з� = 0
                rsMoneyDay!�շѷ�ʽ = -1
                rsMoneyDay!�շ�ʱ�� = Format(date���ղ���ȡ, "yyyy-MM-dd")
                rsMoneyDay.Update
            End If
        End If
    End If
    
    GetPatiDayMoneyDetail = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function InitPatiExecDays() As ADODB.Recordset
'���ܣ���ʼ��ҽ����ط���ִ�еļ�¼��
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "������ĿID", adBigInt
    rsTmp.Fields.Append "�շ���ĿID", adBigInt
    rsTmp.Fields.Append "ִ�в���ID", adBigInt
    rsTmp.Fields.Append "�շѷ�ʽ", adInteger
    rsTmp.Fields.Append "ִ�з�", adInteger
    rsTmp.Fields.Append "�շ�ʱ��", adVarChar, 10
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set InitPatiExecDays = rsTmp
End Function

Public Function GetCuvetteNumber(rsNumber As ADODB.Recordset, ByVal str���� As String, ByVal lngҽ��ID As Long, _
    ByVal lng���ID As Long, ByVal str��� As String, ByVal int�������� As Integer, ByVal lngִ�п���ID As Long, _
    ByVal intӤ�� As Integer, ByVal lng������ĿID As Long, ByVal int���� As Integer, ByVal str�걾 As String, ByVal lng�ɼ�����ID As Long) As String
    '���ܣ��Լ���ҽ��������������
    '      1.һ���ɼ���ͬһ����ҽ��ʹ����ͬ����������
    '      2.��ͬ����ļ���ʹ����ͬ����������
    '      3.У���������:12λ��"����+ҽ��ID"
    '������rsNumber=��̬��¼��������"���롢���ID����������"���ֶ�
    Dim strTmp���� As String, strTmp���� As String
    
    If str��� = "C" And str���� <> "" Then '������Ŀ���й���
        rsNumber.Filter = "���ID=" & lng���ID
        If rsNumber.EOF Then
            rsNumber.Filter = "������Ŀid=" & lng������ĿID
            If rsNumber.EOF Then
                rsNumber.Filter = "����='" & str���� & "' And ִ�п���ID=" & lngִ�п���ID & " And Ӥ��=" & intӤ�� & _
                    " And ������־=" & int���� & " And �걾='" & str�걾 & "' And �ɼ�����ID=" & lng�ɼ�����ID
                If rsNumber.EOF Then
                    '�����µ�����
                    rsNumber.AddNew
                    rsNumber!���� = str����
                    rsNumber!���ID = lng���ID
'                    rsNumber!�������� = str���� & Format(lngҽ��ID, Replace(Space(12 - Len(str����)), " ", "0"))
                    rsNumber!�������� = gobjComlib.zlDatabase.GetNextNo(125, lngҽ��ID)
                    rsNumber!������ĿID = lng������ĿID
                    rsNumber!ִ�п���ID = lngִ�п���ID
                    rsNumber!Ӥ�� = intӤ��
                    rsNumber!������־ = int����
                    rsNumber!�걾 = str�걾
                    rsNumber!�ɼ�����ID = lng�ɼ�����ID
                    rsNumber.Update
                    
                    strTmp���� = rsNumber!��������
                Else
                    '��ͬ���롢ִ�п��ҡ�Ӥ���ļ���ʹ����ͬ����������
                    strTmp���� = Nvl(rsNumber!����)
                    strTmp���� = Nvl(rsNumber!��������)
                    
                    rsNumber.AddNew
                    rsNumber!���� = strTmp����
                    rsNumber!���ID = lng���ID
                    rsNumber!�������� = strTmp����
                    rsNumber!������ĿID = lng������ĿID
                    rsNumber!ִ�п���ID = lngִ�п���ID
                    rsNumber!Ӥ�� = intӤ��
                    rsNumber!������־ = int����
                    rsNumber!�걾 = str�걾
                    rsNumber!�ɼ�����ID = lng�ɼ�����ID
                    rsNumber.Update
                End If
            Else
                '�����µ����룺��ͬ�����ҽ��ʹ��"��ͬ��"����
                rsNumber.AddNew
                rsNumber!���� = str����
                rsNumber!���ID = lng���ID
'                rsNumber!�������� = str���� & Format(lngҽ��ID, Replace(Space(12 - Len(str����)), " ", "0"))
                rsNumber!�������� = gobjComlib.zlDatabase.GetNextNo(125, lngҽ��ID)
                rsNumber!������ĿID = lng������ĿID
                rsNumber!ִ�п���ID = lngִ�п���ID
                rsNumber!Ӥ�� = intӤ��
                rsNumber!������־ = int����
                rsNumber!�걾 = str�걾
                rsNumber!�ɼ�����ID = lng�ɼ�����ID
                rsNumber.Update
                
                strTmp���� = rsNumber!��������
            End If
        Else
            'һ���ɼ��ļ�����Ŀʹ����ͬ������
            strTmp���� = Nvl(rsNumber!����)
            strTmp���� = Nvl(rsNumber!��������)
            
            rsNumber.AddNew
            rsNumber!���� = strTmp����
            rsNumber!���ID = lng���ID
            rsNumber!�������� = strTmp����
            rsNumber!������ĿID = lng������ĿID
            rsNumber!ִ�п���ID = lngִ�п���ID
            rsNumber!Ӥ�� = intӤ��
            rsNumber!������־ = int����
            rsNumber!�걾 = str�걾
            rsNumber!�ɼ�����ID = lng�ɼ�����ID
            rsNumber.Update
        End If
        ElseIf str��� = "E" And int�������� = 6 Then
        '�ɼ���ʽʹ����ҽ����ͬ(���)������
        If Not rsNumber.EOF Then
            If Nvl(rsNumber!���ID, 0) = lngҽ��ID Then
                strTmp���� = Nvl(rsNumber!��������)
            End If
        End If
    End If
    
    GetCuvetteNumber = strTmp����
End Function

Public Function GetAuditRecord(lng����ID As Long, lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡָ�����˵ķ���������Ŀ
    Dim StrSQL As String
    
    On Error GoTo errH
    StrSQL = "Select ��ĿId,ʹ������,��������,ʹ������-�������� �������� From ����������Ŀ Where ����ID=[1] And ��ҳID=[2]"
    Set GetAuditRecord = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlInExse", lng����ID, lng��ҳID)
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Getʵ�ս��(ByVal StrSQL As String) As Currency
    Dim lngPos As Long, strMatch As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    StrSQL = Mid(StrSQL, InStr(StrSQL, strMatch) + Len(strMatch))
    strMatch = "End" & Chr(0) & Chr(1)
    StrSQL = Left(StrSQL, InStr(StrSQL, strMatch) - 1)
    Getʵ�ս�� = CCur(StrSQL)
End Function

Public Function Setʵ�ս��(ByVal StrSQL As String, ByVal cur��� As Currency) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strLeft = Mid(StrSQL, 1, InStr(StrSQL, strMatch) - 1)
    strMatch = "End" & Chr(0) & Chr(1)
    strRight = Mid(StrSQL, InStr(StrSQL, strMatch) + Len(strMatch))
    
    Setʵ�ս�� = strLeft & cur��� & strRight
End Function


Public Function Get������ϼ�¼(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str���� As String) As ADODB.Recordset
'���ܣ���ȡ������ϼ�¼
'������lng����ID�����ﲡ�˴��Һ�ID��סԺ���˴���ҳID
'       �������-1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;
'        11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���
'       ��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����(����ҽ��վ,���ժҪ);
    Dim StrSQL As String

    On Error GoTo errH
    StrSQL = "Select a.����id, a.���id, a.�������, a.��ϴ���, Nvl(b.����, c.����) As ����, Nvl(b.����, c.����) ����" & vbNewLine & _
             "From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C" & vbNewLine & _
             "Where a.����id = [1] And a.��ҳid = [2] And NVL(A.�������,1) = 1  And ȡ��ʱ�� Is Null And ��¼��Դ IN (1, 3) And Instr(',' ||[3]|| ',', ',' || ������� || ',') > 0 And a.����id = b.Id(+) And" & vbNewLine & _
             "      a.���id = c.Id(+)" & vbNewLine & _
             "Order By ��¼��Դ, �������, ��ϴ���"
    Set Get������ϼ�¼ = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlPublic", lng����ID, lng����ID, str����)

    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub ReplaceTrueNO(rsSQL As ADODB.Recordset, rsUpload As ADODB.Recordset)
'���ܣ�����ʱ������NO�滻�����ձ������ʵNO
    Dim strNO As String, strCur As String, strPre As String
    
    rsSQL.Filter = 0
    rsSQL.Sort = "NO"
    Do While Not rsSQL.EOF
        If Not IsNull(rsSQL!NO) Then
            strCur = Split(rsSQL!NO, "=")(1)
            If strCur <> strPre Then
                strPre = strCur
                strNO = gobjComlib.zlDatabase.GetNextNo(Val(Left(strCur, 2)))
                            
                'rsUpload��һ��NOֻ��һ����¼
                If Not rsUpload Is Nothing Then
                    rsUpload.Filter = "NO='" & rsSQL!NO & "'"
                    If Not rsUpload.EOF Then
                        rsUpload!NO = strNO
                        rsUpload.Update
                    End If
                End If
            End If
            
            rsSQL!sql = Replace(rsSQL!sql, rsSQL!NO, strNO)
            'rsSQL!NO = strNO '��������£����⵼��Sort��˳������
            rsSQL.Update
        End If
        rsSQL.MoveNext
    Loop
End Sub

Public Function Get��Һ��������() As String
'���ܣ���ȡ��Һ�������ĵĿ���IDs
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Integer
    Dim strReturn As String
    
    On Error GoTo errH

    StrSQL = "Select ����id From ��������˵�� Where �������� = '��������' Order by ����id"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "Get��Һ��������")
    
    For i = 1 To rsTmp.RecordCount
        strReturn = strReturn & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    Get��Һ�������� = Mid(strReturn, 2)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Sub GetTestLabel(ByVal strScript As String, ByVal strSelect As String, intResult As Integer)
'���ܣ���ȡƤ�Ա�ע�ͽ��
'������strScript=Ƥ�Խ������������"����(+),������(++);����(-)"
'      strSelect=��ѡ���Ƥ�Խ��
'���أ�strLabel = Ƥ�Խ����ע����"(+)"
'      intResult=Ƥ�Խ����0-���ԣ�1-����
    Dim arr���� As Variant, arr���� As Variant
    Dim i As Integer
    
    intResult = 0
    
    arr���� = Split(Split(strScript, ";")(0), ",")
    arr���� = Split(Split(strScript, ";")(1), ",")
    
    For i = 0 To UBound(arr����)
        If arr����(i) Like "*" & strSelect & "*" Then
            intResult = 1: Exit Sub
        End If
    Next
    For i = 0 To UBound(arr����)
        If arr����(i) Like "*" & strSelect & "*" Then
            intResult = 0: Exit Sub
        End If
    Next
End Sub

Public Function GetStockCheck(ByVal bytType As Byte) As Collection
'���ܣ���ȡҩƷ�����ĳ�����ļ���
'������bytType:0-ҩƷ��1-����
    Dim rsTmp As ADODB.Recordset, StrSQL As String
    Dim colStock As Collection, i As Long
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '�������
    
    StrSQL = _
        " Select Distinct A.ID,C.��鷽ʽ" & _
        " From ���ű� A,��������˵�� B," & IIF(bytType = 0, "ҩƷ������", "���ϳ�����") & " C" & _
        " Where B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� " & IIF(bytType = 0, "IN('��ҩ��','��ҩ��','��ҩ��')", "='���ϲ���'") & _
        " And C.�ⷿID(+)=A.ID"
        
    On Error GoTo errH
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "GetStockCheck")
    For i = 1 To rsTmp.RecordCount
        colStock.Add Nvl(rsTmp!��鷽ʽ, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Set GetStockCheck = colStock
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'���ܣ��ж��Ƿ����ָ�������������͵�������
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    
    StrSQL = "Select ���ID From ҩƷ�������� Where ����=[1]"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!���ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetOutPatiInfo(rsPati As Recordset, ByVal lng����ID As Long, ByVal lng�Һ�id As Long) As Boolean
'���ܣ���ȡ������Ϣ
    Dim StrSQL As String
    
    On Error GoTo errH
    
    'ִ�в���(�ű����)�����˿���
    StrSQL = "Select ����ID,Ԥ�����,������� From ������� Where ����=1 And ���� = 1 And ����ID=[1]"
    StrSQL = "Select Decode(A.��ͬ��λID,NULL,NULL,Nvl(A.������λ,D.����)) as ��λ,Nvl(c.����,A.����) ����,Nvl(c.�Ա�,A.�Ա�) �Ա� ,Nvl(c.����,A.����) ���� ,A.�����,C.No as �Һŵ�," & _
        " A.�ѱ�,A.����,A.����ģʽ,zl_PatiWarnScheme(A.����ID) as ���ò���,A.������,Nvl(B.Ԥ�����,0)-Nvl(B.�������,0) as ʣ���" & _
        " From ������Ϣ A,(" & StrSQL & ") B,���˹Һż�¼ C,��Լ��λ D" & _
        " Where A.����ID=B.����ID(+) And A.��ͬ��λID=D.ID(+)" & _
        " And A.����id = C.����id(+) And A.����� = C.�����(+) " & _
        " And A.����ID=[1] And c.id(+)=[2]"
    'Set mrsPati = New ADODB.Recordset
    Set rsPati = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "GetOutPatiInfo", lng����ID, lng�Һ�id)

    GetOutPatiInfo = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetMergeIDs(ByRef vsAdvice As VSFlexGrid, ByVal lngRow As Long, ByVal COL_���ID As Long, ByVal COL_ID As Long) As String
'���ܣ���ȡָ��һ����ҩ��ҽ��ID��(��һ����ҩ���ص�ǰҽ��ID)
'������lngRow=һ����ҩ�Ŀ�ʼҩƷ��
    Dim lng���ID As Long, i As Long
    Dim strҽ��ID As String
    
    With vsAdvice
        lng���ID = Val(.TextMatrix(lngRow, COL_���ID))
        For i = lngRow To .Rows - 1
            If Val(.TextMatrix(i, COL_���ID)) = lng���ID Then
                strҽ��ID = strҽ��ID & "," & Val(.TextMatrix(i, COL_ID))
            Else
                Exit For
            End If
        Next
    End With
    
    GetMergeIDs = Mid(strҽ��ID, 2)
End Function

Public Function GetRXKey(ByRef rsRXKey As ADODB.Recordset, ByVal strKey As String, ByVal strҽ��ID As String) As String
'���ܣ�����ҩƷ�����������ƹؼ���,���ڴ���NO����
'������strKey=��ǰ����NO��Key,������������������Key����
'      strҽ��ID=��ǰҩƷ��ҽ��ID����һ����ҩ�������ID��"ID1,ID2,..."
'                һ����ҩ��ʼ�л����ҩƷ�вŴ���,һ����ҩ�м��д����
    Dim intNextCount As Integer
    Dim strNextID As String
    
    rsRXKey.Filter = "Key='" & strKey & "'"
    If rsRXKey.EOF Then
        strNextID = gobjComlib.zlStr.Listminus(strҽ��ID, "")
        intNextCount = UBound(Split(strNextID, ",")) + 1
        
        rsRXKey.AddNew
        rsRXKey!Key = strKey
        rsRXKey!ҽ��ID = strNextID
        rsRXKey!���� = intNextCount
        rsRXKey!���� = 1
        rsRXKey.Update
    ElseIf strҽ��ID <> "" Then
        strNextID = gobjComlib.zlStr.Listminus(strҽ��ID, rsRXKey!ҽ��ID)
        intNextCount = UBound(Split(strNextID, ",")) + 1
        
        rsRXKey!ҽ��ID = rsRXKey!ҽ��ID & "," & strNextID
        rsRXKey!���� = rsRXKey!���� + intNextCount
        rsRXKey.Update
    
        If rsRXKey!���� > gintRXCount Then
            strNextID = gobjComlib.zlStr.Listminus(strҽ��ID, "")
            intNextCount = UBound(Split(strNextID, ",")) + 1
            
            rsRXKey!���� = rsRXKey!���� + 1
            rsRXKey!ҽ��ID = strNextID
            rsRXKey!���� = intNextCount
            rsRXKey.Update
        End If
    ElseIf strҽ��ID = "" Then
        'һ����ҩ�м���,���ֵ�һ�еĹؼ���
    End If

    GetRXKey = rsRXKey!����
End Function

Public Function GetClinicBillID(ByVal lng��ĿID As Long, ByVal int���� As Integer) As Long
'���ܣ���ȡ������Ŀ��Ӧ�����Ƶ���(���ܸ���,�������ɷ���NO)
'������int����=1-����,2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
    
    On Error GoTo errH
    
    StrSQL = "Select �����ļ�ID From ��������Ӧ�� Where ������ĿID=[1] And Ӧ�ó���=[2]"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lng��ĿID, int����)
    If Not rsTmp.EOF Then GetClinicBillID = Nvl(rsTmp!�����ļ�ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetStock(ByVal lngҩƷID As Long, Optional ByVal lng�ⷿID As Long, Optional ByVal int��Χ As Integer = 2, _
        Optional ByVal strDepartments As String, Optional ByVal lng���� As Double) As Double
'���ܣ���ȡָ���ָⷿ��ҩƷ���������(�������סԺ��λ)
'������int��Χ=1-����,2-סԺ(ȱʡ),0-��ʾ���ۼ�
'      strDepartments����ִ�п����ַ���������������ѯ���
'      lng���� ���lng������Ϊ�գ����ѯ�Ƿ��п������������
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, strTmp As String
    
    On Error GoTo errH
    '��ȡҩƷ���(�����������ҩƷ),ҩ��������ҩƷ����Ч��
    If int��Χ = 0 Or int��Χ = 3 Then
        StrSQL = _
            " Select Nvl(Sum(A.��������),0) as ���" & _
            " From ҩƷ��� A" & _
            " Where A.����=1" & _
            " And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
            " And A.ҩƷID=[1] And Instr([2],',' || a.�ⷿid || ',')>0 Group By A.�ⷿID"
    Else
        strTmp = IIF(int��Χ = 1, "����", "סԺ")
        StrSQL = _
            " Select Nvl(Sum(A.��������),0)/Nvl(B." & strTmp & "��װ,1) as ���" & _
            " From ҩƷ��� A,ҩƷ��� B" & _
            " Where A.ҩƷID=B.ҩƷID(+) And A.����=1" & _
            " And (Nvl(A.����,0)=0 Or A.Ч�� is NULL Or A.Ч��>Trunc(Sysdate))" & _
            " And A.ҩƷID=[1] And Instr([2],',' || a.�ⷿid || ',')>0" & _
            " Group by Nvl(B." & strTmp & "��װ,1),A.�ⷿID"
    End If
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(StrSQL, "mdlCISKernel", lngҩƷID, IIF(strDepartments = "", "," & lng�ⷿID & ",", "," & strDepartments & ","))
    
    Do While Not rsTmp.EOF
    
        If strDepartments = "" Then
            GetStock = Format(rsTmp!���, "0.00000")
            Exit Function
        Else
            If Val(rsTmp!���) & "" > lng���� Then
                GetStock = Format(rsTmp!���, "0.00000")
                Exit Function
            End If
        End If
        rsTmp.MoveNext
    
    Loop
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Trim�ֽ�ʱ��(ByVal lng���� As Long, ByVal str�ֽ�ʱ�� As String) As String
'���ܣ���ҽ��ִ�еķֽ�ʱ�䰴�������нض�
    Dim arrTime() As String, strTmp As String, i As Long
    
    arrTime = Split(str�ֽ�ʱ��, ",")
    For i = 0 To lng���� - 1
        strTmp = strTmp & "," & arrTime(i)
    Next
    Trim�ֽ�ʱ�� = Mid(strTmp, 2)
End Function

Public Function GetPriceGradeSQL(ByVal strҩƷ�۸�ȼ� As String, ByVal str���ļ۸�ȼ� As String, ByVal str��ͨ��Ŀ�۸�ȼ� As String, ByVal strTableTmpA As String, ByVal strTableTmpB As String, _
           ByVal strParNumҩƷ As String, ByVal strParNum���� As String, ByVal strParNum��ͨ��Ŀ As String) As String
'���ܣ����˼۸�ȼ����������ȡ�۸��SQL
'������strҩƷ�۸�ȼ�  '���˵�ҩƷ�۸�ȼ�
'      str���ļ۸�ȼ�  '���˵����ļ۸�ȼ�
'      str��ͨ��Ŀ�۸�ȼ�  '���˵���ͨ��Ŀ�۸�ȼ�
'     strTableTmpA   �շ���ĿĿ¼ ���as ��־,strTableTmpB  �շѼ�Ŀ�� ��As��־��
'     strParNumҩƷ  ҩƷ�۸�ȼ�SQL�������,strParNum����  ���ļ۸�ȼ�SQL�������,strParNum��ͨ��Ŀ  ��ͨ��Ŀ�۸�ȼ�SQL�������
    Dim StrSQL As String
    
    If strҩƷ�۸�ȼ� = "" And str���ļ۸�ȼ� = "" And str��ͨ��Ŀ�۸�ȼ� = "" Then
        StrSQL = " And " & strTableTmpB & ".�۸�ȼ� is Null "
    Else
        StrSQL = " And" & vbNewLine & _
                "      ((Instr(';5;6;7;', ';' || " & strTableTmpA & ".��� || ';') > 0 And " & strTableTmpB & ".�۸�ȼ� = [" & strParNumҩƷ & "]) Or" & vbNewLine & _
                "      (Instr(';4;', ';' || " & strTableTmpA & ".��� || ';') > 0 And " & strTableTmpB & ".�۸�ȼ� = [" & strParNum���� & "]) Or" & vbNewLine & _
                "      (Instr(';4;5;6;7;', ';' || " & strTableTmpA & ".��� || ';') = 0 And " & strTableTmpB & ".�۸�ȼ� = [" & strParNum��ͨ��Ŀ & "]) Or" & vbNewLine & _
                "      (" & strTableTmpB & ".�۸�ȼ� Is Null And Not Exists" & vbNewLine & _
                "       (Select 1" & vbNewLine & _
                "         From �շѼ�Ŀ" & vbNewLine & _
                "         Where " & strTableTmpA & ".Id = �շ�ϸĿid  And" & vbNewLine & _
                "               ((Instr(';5;6;7;', ';' || " & strTableTmpA & ".��� || ';') > 0 And �۸�ȼ� = [" & strParNumҩƷ & "]) Or" & vbNewLine & _
                "               (Instr(';4;', ';' || " & strTableTmpA & ".��� || ';') > 0 And �۸�ȼ� = [" & strParNum���� & "]) Or" & vbNewLine & _
                "               (Instr(';4;5;6;7;', ';' || " & strTableTmpA & ".��� || ';') = 0 And �۸�ȼ� = [" & strParNum��ͨ��Ŀ & "]))))) "

    End If
    
    GetPriceGradeSQL = StrSQL
End Function