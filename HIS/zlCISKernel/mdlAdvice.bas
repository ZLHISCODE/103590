Attribute VB_Name = "mdlAdvice"
Option Explicit

'ˢ������ʱ����Ĳ���״̬
Public Enum TYPE_PATI_State
    ps��Ժ = 0
    psԤ�� = 1
    ps��Ժ = 2
    ps���� = 3          'ҽ��վ:�����ﲡ��(��Ժ)
    ps���� = 4          'ҽ��վ:�ѻ��ﲡ��
    ps���ת�� = 5      'ҽ��վ:���ת�ƻ�ת�����Ĳ���(��Ժ)
    ps��ת�� = 6        'ҽ��վ:��ƴ���ס��ת��������������
End Enum

Public Function BillExpend(ByVal strNO As String) As Boolean
'���ܣ��жϹҺŵ��Ƿ��Ѿ�������Ч�Һ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    '������������������Һ���Ч�����Ĳ��˲��������ʾ�������
    If Val(zlDatabase.GetPara(210, glngSys)) = 1 Then Exit Function
    '��ʱ����
    strSQL = "Select  Sysdate-����ʱ�� as ���,���� From ���˹Һż�¼ Where NO=[1] And ��¼����=1 And ��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", strNO)
    If Not rsTmp.EOF Then
        BillExpend = Val(rsTmp!���) > IIF(Val("" & rsTmp!����) = 1, IIF(gint����Һ����� = 0, 1, gint����Һ�����), IIF(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����))
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckOutAdvice(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ���鲡���Ƿ����´��˳�Ժҽ��
    Dim strSQL As String, rsTmp As Recordset
    
    strSQL = "Select 1 from ����ҽ����¼ A,������ĿĿ¼ B Where a.������ĿID=b.ID And a.����ID=[1] And a.��ҳID=[2] And b.���='Z' And b.��������='5' And a.ҽ��״̬ Not In(-1,4) and nvl(A.Ӥ��,0)=0"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng����ID, lng��ҳID)
    CheckOutAdvice = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExeTimeValid(ByVal strTime As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, Optional ByVal bln���� As Boolean) As Boolean
'���ܣ����ָ����ִ��ʱ���Ƿ�Ϸ�
    Dim arrTime() As String, strTmp As String, i As Integer
    Dim strPreTime As String, intPreDay As Long, intCurDay As Long
    
    If strTime = "" Then
        If str�����λ = "����" Then ExeTimeValid = True
        Exit Function
    End If
    
    If str�����λ = "��" Then
        '1/8:00-3/15:00-5/9:00��1/8:00-3/15-5/9:00
        If Not zlStr.CheckCharScope(strTime, "0123456789:-/") Then Exit Function
        
        arrTime = Split(strTime, "-")
        If bln���� Then
            If Not Between(UBound(arrTime) + 1, 1, intƵ�ʴ���) Then Exit Function
        Else
            If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
        End If
        
        For i = 0 To UBound(arrTime)
            If UBound(Split(arrTime(i), "/")) <> 1 Then Exit Function
            '���ڲ���
            strTmp = Split(arrTime(i), "/")(0)
            If InStr(strTmp, ":") > 0 Or strTmp = "" Then Exit Function
            intCurDay = Val(strTmp)
            If intCurDay < 1 Or intCurDay > 7 Then Exit Function
            If intPreDay <> 0 Then
                If intCurDay < intPreDay Then Exit Function
            End If
            
            '����ʱ�䲿��
            strTmp = Split(arrTime(i), "/")(1)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
            If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then Exit Function
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
            If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
            End If
            
            strPreTime = Format(strTmp, "HH:mm")
            intPreDay = intCurDay
        Next
    ElseIf str�����λ = "��" Then
        If intƵ�ʼ�� = 1 Then
            '8:00-12:00-14:00��8:00-12-14:00
            If Not zlStr.CheckCharScope(strTime, "0123456789:-") Then Exit Function
            
            arrTime = Split(strTime, "-")
            If bln���� Then
                If Not Between(UBound(arrTime) + 1, 1, intƵ�ʴ���) Then Exit Function
            Else
                If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
            End If
            
            For i = 0 To UBound(arrTime)
                strTmp = arrTime(i)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then Exit Function
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
                If strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
                End If
                strPreTime = Format(strTmp, "HH:mm")
            Next
        Else
            '1/8:00-1/15:00-2/9:00��1/8:00-1/15-2/9:00
            If Not zlStr.CheckCharScope(strTime, "0123456789:-/") Then Exit Function
            
            arrTime = Split(strTime, "-")
            If bln���� Then
                If Not Between(UBound(arrTime) + 1, 1, intƵ�ʴ���) Then Exit Function
            Else
                If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
            End If
            
            For i = 0 To UBound(arrTime)
                If UBound(Split(arrTime(i), "/")) <> 1 Then Exit Function
                '�����������
                strTmp = Split(arrTime(i), "/")(0)
                If InStr(strTmp, ":") > 0 Or strTmp = "" Then Exit Function
                intCurDay = Val(strTmp)
                If intCurDay < 1 Or intCurDay > intƵ�ʼ�� Then Exit Function
                If intPreDay <> 0 Then
                    If intCurDay < intPreDay Then Exit Function
                End If
                
                '����ʱ�䲿��
                strTmp = Split(arrTime(i), "/")(1)
                If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
                If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
                If Val(Split(strTmp, ":")(0)) >= 24 Or Split(strTmp, ":")(0) = "" Or Len(Split(strTmp, ":")(0)) > 2 Then Exit Function
                If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
                If intPreDay <> 0 And intPreDay = intCurDay And strPreTime <> "" Then
                    If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
                End If
                
                strPreTime = Format(strTmp, "HH:mm")
                intPreDay = intCurDay
            Next
        End If
    ElseIf str�����λ = "Сʱ" Then
        '1:30-2-3:30
        If Not zlStr.CheckCharScope(strTime, "0123456789:-") Then Exit Function
        
        arrTime = Split(strTime, "-")
        If UBound(arrTime) + 1 <> intƵ�ʴ��� Then Exit Function
        
        For i = 0 To UBound(arrTime)
            strTmp = arrTime(i)
            If InStr(strTmp, ":") = 0 Then strTmp = strTmp & ":00"
            If UBound(Split(strTmp, ":")) <> 1 Then Exit Function
            If Val(Split(strTmp, ":")(0)) < 1 Or Val(Split(strTmp, ":")(0)) > intƵ�ʼ�� Or Split(strTmp, ":")(0) = "" Then Exit Function
            If Val(Split(strTmp, ":")(1)) >= 60 Or Split(strTmp, ":")(1) = "" Or Len(Split(strTmp, ":")(1)) > 2 Then Exit Function
            If strPreTime <> "" Then
                If Format(strTmp, "HH:mm") <= strPreTime Then Exit Function
            End If
            strPreTime = Format(strTmp, "HH:mm")
        Next
    End If
    
    ExeTimeValid = True
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


Public Function GetMaxBedLen(Optional lng����ID As Long, Optional bln���� As Boolean) As Integer
'���ܣ���ȡָ�����ŵĴ�λ�ŵ���󳤶�
'������lng����ID=����ID�����ID,Ϊ0��ʾ���в��������
'      blnռ��=�Ƿ�ֻ�ܱ�ռ�õĴ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not bln���� Or lng����ID = 0 Then
        strSQL = "Select Max(LengthB(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIF(lng����ID = 0, " is Not NULL", "= [1] ")
    Else
        strSQL = "Select Max(LengthB(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIF(lng����ID = 0, " is Not NULL", "= [1] ")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID)
    
    If Not rsTmp.EOF Then GetMaxBedLen = IIF(IsNull(rsTmp!����), 0, rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DateIsPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ�������Ƿ�����ͣ��ʱ�����
'������strPause="��ͣʱ��,��ʼʱ��;...."
'˵��������ʱ���ж�,����ͣ���ڰ���ʼ����ֹ�����ж�
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String
    
    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Format(Split(arrPause(i), ",")(0), "yyyy-MM-dd")
        strEnd = Format(Split(arrPause(i), ",")(1), "yyyy-MM-dd")
        If strEnd = "" Then strEnd = "3000-01-01" '������δ���û���ͣ��ʱ��ֹͣ
        If strEnd > strBegin Then
            If Between(Format(vDate, "yyyy-MM-dd"), strBegin, _
                Format(DateAdd("d", -1, CDate(strEnd)), "yyyy-MM-dd")) Then
                DateIsPause = True: Exit Function
            End If
        End If
    Next
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
        vCurTime = zlCommFun.GetWeekBase(dat��ʼʱ��)
        
        Do While lng���� > 0
            blnFirst = (zlCommFun.GetWeekBase(vCurTime) = zlCommFun.GetWeekBase(dat��������)) And dat�������� <> Empty And UBound(arrFirst) <> -1
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

Public Function Calc���ڷֽ�ʱ��(ByVal datBegin As Date, ByVal datEnd As Date, ByVal strPause As String, _
    ByVal strִ��ʱ�� As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, _
    Optional ByVal dat�������� As Date) As String
'���ܣ���ʱ��μ�����εķֽ�ִ��ʱ�估����
'������datBegin-datEnd=Ҫ�����ʱ���,����datBeginӦΪÿ�����ڵĿ�ʼ��׼ʱ��
'      strPause=��ͣ��ʱ���
'      dat��������=��������ʱ��������
'���أ�"ʱ��1,ʱ��2,...."(yyyy-MM-dd HH:mm:ss),ʱ�������Ϊ����
'˵����1.ʱ�����Ҫ�ų���ͣ��ʱ���,����������˶�����
'      2.�������Ǽٶ���ִ��ʱ�估Ƶ��������ȫ��ȷ������¼��㡣
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrNormal As Variant, arrFirst As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer
    
    If InStr(strִ��ʱ��, ",") > 0 Then
        arrNormal = Split(Split(strִ��ʱ��, ",")(1), "-")
        arrFirst = Split(Split(strִ��ʱ��, ",")(0), "-")
    Else
        arrNormal = Split(strִ��ʱ��, "-")
        arrFirst = Array()
    End If
        
    vCurTime = datBegin
    
    If str�����λ = "��" Then
        vCurTime = zlCommFun.GetWeekBase(datBegin)
        If dat�������� <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (vCurTime = zlCommFun.GetWeekBase(dat��������))
        Else
            blnFirst = False
        End If

        Do While vCurTime <= datEnd
            arrTime = IIF(blnFirst, arrFirst, arrNormal)
            blnFirst = False
                        
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
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                End If
            Next
            vCurTime = Format(vCurTime + 7, "yyyy-MM-dd") '������
        Loop
    ElseIf str�����λ = "��" Then
        If dat�������� <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (Int(vCurTime) = Int(dat��������))
        Else
            blnFirst = False
        End If
        
        Do While vCurTime <= datEnd
            arrTime = IIF(blnFirst, arrFirst, arrNormal)
            blnFirst = False
            
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
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
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
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = Format(vCurTime + intƵ�ʼ��, "yyyy-MM-dd") '������
        Loop
    ElseIf str�����λ = "Сʱ" Then
        '10:00-20:00-40:00��10-20-40��02:30
        arrTime = arrNormal
        Do While vCurTime <= datEnd
            For i = 1 To intƵ�ʴ���
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime >= Format(datBegin, "yyyy-MM-dd HH:mm:ss") And vTmpTime <= Format(datEnd, "yyyy-MM-dd HH:mm:ss") Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + intƵ�ʼ�� / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str�����λ = "����" Then
        '��ִ��ʱ��
        Do While vCurTime <= datEnd
            vTmpTime = vCurTime
            
            If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                If Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                End If
            ElseIf vTmpTime > datEnd Then
                Exit Do
            End If

            vCurTime = Format(vCurTime + intƵ�ʼ�� / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If
    
    Calc���ڷֽ�ʱ�� = Mid(strDetailTime, 2)
End Function

Public Function CalcȱʡҩƷ����(ByVal dbl���� As Double, ByVal int�Ƴ� As Integer, _
    ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, Optional ByVal strִ��ʱ�� As String, _
    Optional ByVal dbl����ϵ�� As Double, Optional ByVal dbl��װϵ�� As Double, Optional ByVal int���� As Integer, Optional ByVal dbl�״����� As Double) As Double
'���ܣ����Ƴ̼��������Լ���ҩƷ������ȱʡ����(���䷽ȱʡ����)
'������dbl����=��������λ��һ������
'      int�Ƴ�=һ���Ƴ̵�����
'      int����=0-�ɷ���,1-������,2-һ����(��ʱʧЧ),-N-N���ڷ���ʹ����Ч
'      dbl��װϵ��=�����װ��סԺ��װ
'���أ���סԺ��λ�����ҩƷ����
'˵����
'     1.ҩƷ������������������סԺ��װ���Եġ�
'     2.dbl����ϵ��,dbl��װϵ��,int����=��ҩ������,ֻ���㸶��
    Dim dbl��� As Double, dbl���� As Double
    Dim dblʣ�� As Double, dblOne As Double
    Dim intStep As Integer, dblEnd As Double
    Dim arrTime() As String, strBegin As String
    Dim strTime As String, i As Integer, j As Integer
    Dim dblһ������ As Double
    
    '�Ƴ̲���һ��Ƶ������ʱ�Ͳ����Ƴ�
    If str�����λ = "��" Then
        If int�Ƴ� < 7 Then int�Ƴ� = 1
    ElseIf str�����λ = "��" Then
        If int�Ƴ� < intƵ�ʼ�� Then int�Ƴ� = 1
    ElseIf str�����λ = "Сʱ" Then
        If int�Ƴ� < intƵ�ʼ�� / 24 Then int�Ƴ� = 1
    ElseIf str�����λ = "����" Then
        If int�Ƴ� < intƵ�ʼ�� / (24 * 60) Then int�Ƴ� = 1
    End If
    
    'һ��Ƶ�����ڵĴ���(����)
    If str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / 7
    ElseIf str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / intƵ�ʼ��
    ElseIf str�����λ = "Сʱ" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * 24
    ElseIf str�����λ = "����" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * (24 * 60)
    End If
    
    If dbl����ϵ�� = 0 And dbl��װϵ�� = 0 Then
        '��ҩ����(����) = ����*�Ƴ�*(Ƶ�ʴ���/Ƶ�ʼ��)
        dbl���� = IntEx(int�Ƴ� * dbl���)
    Else
        'ҩƷ�������� = ����/סԺ��װ(����*�Ƴ�*(Ƶ�ʴ���/Ƶ�ʼ��))
        If int���� = 0 Then
            '�ɷ���
            dbl���� = dbl���� * int�Ƴ� * dbl��� / dbl����ϵ�� / dbl��װϵ��
        ElseIf int���� = 1 Then
            '������
            dbl���� = IntEx(dbl���� * int�Ƴ� * dbl��� / dbl����ϵ�� / dbl��װϵ��)
        ElseIf int���� = 2 Then
            'һ����(��ʱʧЧ)
            dbl���� = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * IntEx(int�Ƴ� * dbl���)
        ElseIf int���� < 0 Then
            'ABS(int����)���ڷ���ʹ����Ч(�����������)
            If strִ��ʱ�� <> "" Then
                'һ������/סԺ��װ�ļ���
                dblOne = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * (dbl����ϵ�� * dbl��װϵ��)
                dblһ������ = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��)
                'ȱʡִ�еĴ�����ʱ��ֽ�
                strTime = Calc�����ֽ�ʱ��(IntEx(int�Ƴ� * dbl���), datE, CDate("3000-01-01"), "", strִ��ʱ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                If strTime <> "" Then
                    arrTime = Split(strTime, ",")
                    dblʣ�� = dblOne: dbl���� = dblһ������
                    strBegin = arrTime(0)
                    
                    '��������
                    For i = 0 To UBound(arrTime)
                        If dblʣ�� < dbl���� Or CDate(arrTime(i)) - CDate(strBegin) >= Abs(int����) Then
                            If CDate(arrTime(i)) - CDate(strBegin) >= Abs(int����) Then
                                dblʣ�� = dblOne
                            Else
                                dblʣ�� = dblʣ�� + dblOne
                            End If
                            dbl���� = dbl���� + dblһ������
                            strBegin = arrTime(i)
                        End If
                        dblʣ�� = dblʣ�� - dbl����
                        If dblʣ�� >= dbl����ϵ�� * dbl��װϵ�� Then
                            '���ʣ����Ǵ�����ƿ�����ȥ����
                            dbl���� = dbl���� - Int(dblʣ�� / dbl����ϵ�� / dbl��װϵ��)
                            dblʣ�� = dblʣ�� Mod (dbl����ϵ�� * dbl��װϵ��)
                        End If
                    Next
                End If
            End If
        End If
    End If
    If dbl���� > 0 And dbl�״����� > 0 Then
        dbl���� = dbl���� + (dbl�״����� - dbl����) / dbl����ϵ�� / dbl��װϵ��
    End If
    CalcȱʡҩƷ���� = dbl����
End Function

Public Function CalcȱʡҩƷ����(ByVal dbl���� As Double, ByVal dbl���� As Double, _
    ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, _
    Optional ByVal dbl����ϵ�� As Double, Optional ByVal dbl��װϵ�� As Double, Optional ByVal int���� As Integer) As Long
'���ܣ�������������������ҩƷ���Լ�����ҩ����
'������dbl����=�û����������
'      dbl����=��������λ��һ������
'      int����=0-�ɷ���,1-������,2-һ����(��ʱʧЧ),-N-N���ڷ���ʹ����Ч
'      dbl��װϵ��=�����װ��סԺ��װ
'���أ���ҩ����(��ҩ����������)
    Dim dbl��� As Double
    Dim lng���� As Long
    
    'һ��Ƶ�����ڵĴ���(����)
    If str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / 7
    ElseIf str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / intƵ�ʼ��
    ElseIf str�����λ = "Сʱ" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * 24
    ElseIf str�����λ = "����" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * (24 * 60)
    End If
    
    If int���� = 0 Then
        '�ɷ���
        'dbl���� = dbl���� * int�Ƴ� * dbl��� / dbl����ϵ�� / dbl��װϵ��
        lng���� = Format(dbl���� * dbl��װϵ�� * dbl����ϵ�� / dbl���� / dbl���, "0")
    ElseIf int���� = 1 Then
        '������
        'dbl���� = IntEx(dbl���� * int�Ƴ� * dbl��� / dbl����ϵ�� / dbl��װϵ��)
        lng���� = Format(dbl���� * dbl��װϵ�� * dbl����ϵ�� / dbl���� / dbl���, "0")
    ElseIf int���� = 2 Then
        'һ����(��ʱʧЧ)
        'dbl���� = IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * IntEx(int�Ƴ� * dbl���)
        lng���� = Format(dbl���� / IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) / dbl���, "0")
    ElseIf int���� < 0 Then
        'ABS(int����)���ڷ���ʹ����Ч(�����������)
        lng���� = Format(dbl���� * dbl��װϵ�� * dbl����ϵ�� / dbl���� / dbl���, "0")
    End If

    CalcȱʡҩƷ���� = lng����
End Function

Public Function Calc����ҩƷ����(ByVal dat��ʼִ��ʱ�� As Date, lng���� As Long, str�ֽ�ʱ�� As String, _
    ByVal dbl���� As Double, ByVal dbl����ϵ�� As Double, ByVal dbl��װϵ�� As Double, _
    ByVal int���� As Integer, ByVal dat��ֹʱ�� As Date, ByVal strPause As String, ByVal strִ��ʱ�� As String, _
    ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, _
    Optional ByVal blnLimit As Boolean, Optional ByVal dbl�״����� As Double, Optional ByVal dat�ϴ�ִ��ʱ�� As Date) As Double
'���ܣ������ʹ������������Լ����ҩ����
'������dat��ʼִ��ʱ��=ҽ���Ŀ�ʼִ��ʱ��,���ڼ�����һִ�����ڿ�ʼ��׼ʱ��
'      lng����=���μƻ�Ҫ���͵Ĵ���
'      dbl����=��������λ��һ������
'      int����=0-�ɷ���,1-������,2-һ����(��ʱʧЧ),-N-N���ڷ���ʹ����Ч(��24Сʱ����)
'      dbl��װϵ��=�����װ��סԺ��װ
'      blnLimit=�Ƿ�ʱ�����Ƽ����ҩ;��������ʣ�ಿ��
'���в������ڲ�����ҩƷ����(����-N��)��
'      str�ֽ�ʱ��=���η��ͼƻ�ִ�еķֽ�ʱ��,�������Ӧ
'      strPause=ҽ������ͣʱ���
'      dat��ֹʱ��=ҽ����ִ����ֹʱ��,û��ʱ����"3000-01-01"
'���أ�1.������/סԺ��λ�����ҩƷ����
'      2.lng����=������ҩƷ(����-N�ͷ���ҩƷ)������ʵ��ִ�д���(����)
'      3.str�ֽ�ʱ��=������ҩƷ(����-N�ͷ���ҩƷ)�����ķֽ�ʱ��(����)
'˵����ҩƷ������������������סԺ��װ���Եġ�
    Dim dbl���� As Double, dblʣ�� As Double
    Dim arrTime() As String, dblOne As Double
    Dim strBegin As String, datBase As Date
    Dim strTmp As String, i As Long
    Dim blnIsFirst As Boolean
    
    'ע��һЩ�ط���Val����Ϊ��������Double��ĳЩ�ط��ж�ʱ���ڲ����������⣬���±���0.9<>0.9
    If int���� = 0 Then
        '�ɷ���
        dbl���� = Val(dbl���� * lng���� / dbl����ϵ�� / dbl��װϵ��)
        '����ϴ�ִ��ʱ��ΪNULL��˵�������״�
        If dat�ϴ�ִ��ʱ�� = CDate(0) And dbl�״����� > 0 Then
            dbl���� = Val(dbl���� + (dbl�״����� - dbl����) / dbl����ϵ�� / dbl��װϵ��)
        End If
    ElseIf int���� = 1 Then
        '������
        dbl���� = Val(dbl���� * lng���� / dbl����ϵ�� / dbl��װϵ��)
        '����ϴ�ִ��ʱ��ΪNULL��˵�������״�
        If dat�ϴ�ִ��ʱ�� = CDate(0) And dbl�״����� > 0 Then
            dbl���� = Val(dbl���� + (dbl�״����� - dbl����) / dbl����ϵ�� / dbl��װϵ��)
        End If
        dbl���� = Val(IntEx(dbl����))
        '�����������ʱ,����ľ�����ʹ��,�Ӷ�ʹ���ʹ�������
        If Not blnLimit Then
            dblʣ�� = Val(dbl���� * dbl��װϵ�� * dbl����ϵ�� - (dbl���� * (lng���� - 1) + IIF(0 = dbl�״�����, dbl����, dbl�״�����)))
            If dblʣ�� >= dbl���� And dbl���� <> 0 Then
                'ʣ�����ۿ���ִ�еĴ���
                i = Int(Val(dblʣ�� / dbl����))
                'ʣ��ʵ�ʿ���ִ�еĴ�����ʱ��ֽ�(����ֹʱ������)
                arrTime = Split(str�ֽ�ʱ��, ",")
                datBase = Calc�����ڿ�ʼʱ��(dat��ʼִ��ʱ��, CDate(arrTime(UBound(arrTime))), intƵ�ʼ��, str�����λ)
                
                '��������չʱ��ʱ,���һ����������ִ�е�ʱ�䲻�ټ���,����ͣ����
                strPause = strPause & ";" & Format(datBase, "yyyy-MM-dd HH:mm:ss") & "," & arrTime(UBound(arrTime))
                If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                
                strTmp = Calc�����ֽ�ʱ��(i, datBase, dat��ֹʱ��, strPause, strִ��ʱ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ, dat��ʼִ��ʱ��)
                If strTmp <> "" Then
                    lng���� = lng���� + i
                    str�ֽ�ʱ�� = str�ֽ�ʱ�� & "," & strTmp
                End If
            End If
        End If
    ElseIf int���� = 2 Then
        'һ����(��ʱʧЧ)
        dbl���� = Val(IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * lng����)
        '����ϴ�ִ��ʱ��ΪNULL��˵�������״�
        If dat�ϴ�ִ��ʱ�� = CDate(0) And dbl�״����� > 0 Then
            dbl���� = Val(dbl���� + IntEx(dbl�״����� / dbl����ϵ�� / dbl��װϵ��) - IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��))
        End If
    ElseIf int���� < 0 Then
        'ABS(int����)���ڷ���ʹ����Ч(�����������)
        arrTime = Split(str�ֽ�ʱ��, ",")
        strBegin = arrTime(0)
        
        'һ������/סԺ��װ�ļ���(������λ)
        dblOne = Val(IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * (dbl����ϵ�� * dbl��װϵ��))
        'һ������/סԺ��װ�ļ���(��װ��λ)
        dbl���� = Val(IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��))
        '����ϴ�ִ��ʱ��ΪNULL��˵�������״�
        If dat�ϴ�ִ��ʱ�� = CDate(0) And dbl�״����� > 0 Then
            dbl���� = Val(IntEx(dbl�״����� / dbl����ϵ�� / dbl��װϵ��))
            dblOne = Val(IntEx(dbl�״����� / dbl����ϵ�� / dbl��װϵ��) * (dbl����ϵ�� * dbl��װϵ��))
            blnIsFirst = True
        End If
         '��������
        dblʣ�� = dblOne
        For i = 0 To UBound(arrTime)
            '��һ��ѭ���϶���,���Բ���������
            If dblʣ�� < IIF(blnIsFirst, dbl�״�����, dbl����) Or CDate(arrTime(i)) - CDate(strBegin) >= Abs(int����) Then
                If CDate(arrTime(i)) - CDate(strBegin) >= Abs(int����) Then
                    dblʣ�� = dblOne
                    dbl���� = dbl���� + IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��)
                Else
                    If dblʣ�� + dbl����ϵ�� * dbl��װϵ�� >= IIF(blnIsFirst, dbl�״�����, dbl����) Then
                        'ֻ��ʣ���һ����װ��λ����
                        dblʣ�� = dblʣ�� + dbl����ϵ�� * dbl��װϵ��
                        dbl���� = dbl���� + 1
                    Else
                        '��Ҫʣ���һ�ΰ�װ��λ�Ź�
                        dblʣ�� = dblʣ�� + dblOne
                        dbl���� = dbl���� + IntEx(IIF(blnIsFirst, dbl�״�����, dbl����) / dbl����ϵ�� / dbl��װϵ��)
                    End If
                End If
                strBegin = arrTime(i)
            End If
            dblʣ�� = dblʣ�� - IIF(blnIsFirst, dbl�״�����, dbl����)
            If blnIsFirst Then
                blnIsFirst = False
                dblOne = Val(IntEx(dbl���� / dbl����ϵ�� / dbl��װϵ��) * (dbl����ϵ�� * dbl��װϵ��))
            End If
        Next
        
        'ʣ�ಿ�ּ�������Ч���ڰ����������,�Ӷ�ʹ���ʹ�������
        If Not blnLimit Then
            If dblʣ�� >= dbl���� And dbl���� <> 0 Then
                'ʣ�����ۿ���ִ�еĴ���
                i = Int(Val(dblʣ�� / dbl����))
                'ʣ��ʵ�ʿ���ִ�еĴ�����ʱ��ֽ�(����ֹʱ������)
                datBase = Calc�����ڿ�ʼʱ��(dat��ʼִ��ʱ��, CDate(arrTime(UBound(arrTime))), intƵ�ʼ��, str�����λ)
                
                '��������չʱ��ʱ,���һ����������ִ�е�ʱ�䲻�ټ���,����ͣ����
                strPause = strPause & ";" & Format(datBase, "yyyy-MM-dd HH:mm:ss") & "," & arrTime(UBound(arrTime))
                If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                
                strTmp = Calc�����ֽ�ʱ��(i, datBase, dat��ֹʱ��, strPause, strִ��ʱ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ, dat��ʼִ��ʱ��)
                If strTmp <> "" Then
                    arrTime = Split(strTmp, ",")
                    For i = 0 To UBound(arrTime)
                        If dblʣ�� < dbl���� Or CDate(arrTime(i)) - CDate(strBegin) >= Abs(int����) Then
                            Exit For
                        End If
                        lng���� = lng���� + 1
                        str�ֽ�ʱ�� = str�ֽ�ʱ�� & "," & arrTime(i)
                        dblʣ�� = dblʣ�� - dbl����
                    Next
                End If
            End If
        End If
    End If
    
    Calc����ҩƷ���� = dbl����
End Function

Public Function Calc�����ڿ�ʼʱ��(ByVal dat��ʼִ��ʱ�� As Date, ByVal datĳ��ִ��ʱ�� As Date, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String) As Date
'���ܣ����ݳ�����ĳ��ִ��ʱ�䣬�õ����ڸ������ڵĿ�ʼ��׼ʱ��
    Dim datBegin As Date, datCurr As Date
    
    datCurr = dat��ʼִ��ʱ��
    datBegin = datCurr
    If str�����λ = "��" Then datCurr = zlCommFun.GetWeekBase(datCurr)
    
    If str�����λ = "" Then Exit Function
    Do While datCurr <= datĳ��ִ��ʱ��
        datBegin = datCurr
        If str�����λ = "��" Then
            datCurr = datCurr + 7
        ElseIf str�����λ = "��" Then
            datCurr = datCurr + intƵ�ʼ��
        ElseIf str�����λ = "Сʱ" Then
            datCurr = DateAdd("h", intƵ�ʼ��, datCurr)
        ElseIf str�����λ = "����" Then
            datCurr = DateAdd("n", intƵ�ʼ��, datCurr)
        End If
    Loop
    Calc�����ڿ�ʼʱ�� = datBegin
End Function

Public Function CalcÿСʱ����ӵ��췢�Ϳ�ʼʱ��(ByVal dat��ʼִ��ʱ�� As Date, ByVal datCurr As Date, ByVal intƵ�ʼ�� As Integer) As Date
'���ܣ���Ҫ����ÿNСʱһ�ε���������ҹ�ѡ�˳����ӵ��쿪ʼ���͵ļ������⣻89561
    Dim datBegin As Date

    datBegin = dat��ʼִ��ʱ��
    Do While CDate(Format(datBegin, "yyyy-mm-dd")) < CDate(Format(datCurr, "yyyy-mm-dd"))
        datBegin = DateAdd("h", intƵ�ʼ��, datBegin)
    Loop
    CalcÿСʱ����ӵ��췢�Ϳ�ʼʱ�� = datBegin
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

Public Function Calc�����Գ�������(ByVal datBegin As Date, ByVal datEnd As Date, _
    ByVal str�ϴ�ִ��ʱ�� As String, ByVal strִ����ֹʱ�� As String, _
    ByVal strPause As String, Optional str�״�ʱ�� As String, _
    Optional strĩ��ʱ�� As String, Optional str�ֽ�ʱ�� As String) As Long
'���ܣ��Գ����Է�ҩ��������������Ӧ�÷��͵Ĵ���,����ĩʱ��
'������str�ϴ�ִ��ʱ��=��һ�����ڱ��η��͵Ŀ�ʼʱ��
'      strִ����ֹʱ��=
'���أ����θ�ҽ�����͵Ĵ���
'      str�״�ʱ��,strĩ��ʱ��=����yyyy-MM-dd HH:mm:ss
'˵���������Գ���������ÿ�췢��һ�δ���,��������봲λ������(��ͣʱ����ʼ����ֹ)
    Dim curDate As Date, lng���� As Long, blnSend As Boolean
    
    str�״�ʱ�� = "": strĩ��ʱ�� = "": str�ֽ�ʱ�� = ""
    curDate = CDate(Format(datBegin, "yyyy-MM-dd"))
    Do While curDate <= CDate(Format(datEnd, "yyyy-MM-dd"))
        If Not DateIsPause(curDate, strPause) Then
            blnSend = True
            If str�ϴ�ִ��ʱ�� <> "" Then
                If Format(curDate, "yyyy-MM-dd") <= Format(str�ϴ�ִ��ʱ��, "yyyy-MM-dd") Then
                    blnSend = False 'Ӧ�����ϴ�ִ��ʱ���ִ��
                End If
            End If
            If strִ����ֹʱ�� <> "" Then
                If Format(curDate, "yyyy-MM-dd") > Format(strִ����ֹʱ��, "yyyy-MM-dd") Then
                    blnSend = False 'ӦС�ڵ���ִ����ֹʱ���ִ��
                End If
            End If
            If blnSend Then
                lng���� = lng���� + 1
                If lng���� = 1 Then
                    str�״�ʱ�� = Format(curDate, "yyyy-MM-dd 00:00:00") '��Ϊ���ִ��
                    If str�״�ʱ�� < Format(datBegin, "yyyy-MM-dd HH:mm:ss") Then
                        str�״�ʱ�� = Format(datBegin, "yyyy-MM-dd HH:mm:ss")
                    End If
                    strĩ��ʱ�� = str�״�ʱ��
                    str�ֽ�ʱ�� = str�״�ʱ��
                Else
                    strĩ��ʱ�� = Format(curDate, "yyyy-MM-dd 00:00:00")
                    str�ֽ�ʱ�� = str�ֽ�ʱ�� & "," & strĩ��ʱ��
                End If
            End If
        End If
        curDate = curDate + 1
    Loop
    
    Calc�����Գ������� = lng����
End Function

 Public Function Calc������������(ByVal dbl���� As Double, ByVal dbl���� As Double, ByVal dbl����ϵ�� As Double, ByVal dbl��װϵ�� As Double, _
    ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String) As Double
'���ܣ�����ָ����������������Ƶ�ʼ���ҩƷ����ʹ�õ�����
    Dim dbl��� As Double
    Dim dbl�ܵ��� As Double
    
    'һ��Ƶ�����ڵĴ���(����)
    If str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / 7
    ElseIf str�����λ = "��" Then
        dbl��� = intƵ�ʴ��� / intƵ�ʼ��
    ElseIf str�����λ = "Сʱ" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * 24
    ElseIf str�����λ = "����" Then
        dbl��� = (intƵ�ʴ��� / intƵ�ʼ��) * (24 * 60)
    End If
    
    dbl�ܵ��� = dbl���� * dbl��װϵ�� * dbl����ϵ��
    
    Calc������������ = dbl�ܵ��� / dbl���� / dbl���
End Function

Public Sub GetPatiLastChange(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByRef lng����ID As Long, ByRef lng����id As Long, Optional ByVal int���� As Integer = -1, Optional ByRef strTurnDate As String)
'���ܣ���ȡ���������ת�ƻ�ת������Ϣ
'������int���� -1-����������0-ҽ��վ��1-��ʿվ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    If int���� = -1 Or int���� = 1 Then
        strSQL = " And (��ֹԭ�� = 3 Or ��ֹԭ�� = 15)"
    ElseIf int���� = 0 Or int���� = 2 Then
        strSQL = " And (��ֹԭ�� = 3 )"
    End If
    
    strSQL = "Select ����id, ����id,��ֹʱ��" & vbNewLine & _
        "From (Select ����id, ����id,��ֹʱ��" & vbNewLine & _
        "       From ���˱䶯��¼" & vbNewLine & _
        "       Where ����id = [1] And ��ҳid = [2]  And ��ֹʱ�� Is Not Null" & strSQL & vbNewLine & _
        "       Order By ��ֹʱ�� Desc)" & vbNewLine & _
        "Where Rownum = 1"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiLastChange", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        lng����ID = Val("" & rsTmp!����ID)
        lng����id = Val("" & rsTmp!����ID)
        strTurnDate = Format(rsTmp!��ֹʱ��, "yyyy-MM-dd HH:mm:ss")
    Else
        lng����ID = 0
        lng����id = 0
        strTurnDate = ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetPatiDayMoney(lng����ID As Long) As Currency
'���ܣ���ȡָ�����˵��췢���ķ����ܶ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = NVL(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiCanBilling(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String, Optional ByVal lngModual As Long) As Boolean
'���ܣ����ָ�������Ƿ�������Ȩ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    PatiCanBilling = True
    
    If InStr(strPrivs, "��Ժδ��ǿ�Ƽ���") > 0 And InStr(strPrivs, "��Ժ����ǿ�Ƽ���") > 0 Then
        Exit Function
    End If
    On Error GoTo errH
    strSQL = "Select NVL(B.����,A.����) ����,B.��Ժ����,B.״̬,X.�������" & _
        " From ������Ϣ A,������ҳ B,������� X" & _
        " Where A.����ID=B.����ID And A.����ID=X.����ID(+) And X.����(+) = 2" & _
        " And A.����ID=[1] And B.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!��Ժ����) And NVL(rsTmp!״̬, 0) <> 3 Then Exit Function
        If InStr(strPrivs, "��Ժδ��ǿ�Ƽ���") = 0 Then
            If NVL(rsTmp!�������, 0) <> 0 Then
                strMsg = """" & rsTmp!���� & """�ķ���δ���壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If
        If InStr(strPrivs, "��Ժ����ǿ�Ƽ���") = 0 Then
            If NVL(rsTmp!�������, 0) = 0 Then
                strMsg = """" & rsTmp!���� & """�ķ����ѽ��壬��ǰ�Ѿ���Ժ(��Ԥ��Ժ)���㲻���жԸò��˼��ʵ�Ȩ�ޡ�"
            End If
        End If

        If strMsg <> "" Then
            PatiCanBilling = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckEPRReport(ByVal lngҽ��ID As Long, Optional lng����ID As Long, Optional blnBySign As Boolean, Optional ByVal intִ��״̬ As Integer = -999, Optional ByVal blnMoved As Boolean) As Integer
'���ܣ�����Ӧ��Ŀ�ı�����д���
'������lngҽ��ID=�ɼ��е�ҽ��ID
'      lng����ID=���Դ��룬��Ҫ���ڷ��ر��没��ID
'      intִ��״̬=���ڼ������ʱ�������ۺϵ�ִ��״̬
'������blnBySign=�����Ƿ����ͨ��ǩ�������ж�(����ҽ������վ)
'���أ�0-���滹û����д
'      1-��������д���(��ǩ��,�����޶���ǩ��,����ִ�����)
'      2-����δ��д���(δǩ��,���޶���δǩ��,��δִ�����)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str��鱨��ID As String
    
    On Error GoTo errH
    
    '��鱨���Ƿ�����д
    If lng����ID = 0 Then
        strSQL = "Select ����ID,��鱨��ID || ''  as ��鱨��ID From ����ҽ������ Where ҽ��ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lngҽ��ID)
        If Not rsTmp.EOF Then lng����ID = Val(rsTmp!����id & ""): str��鱨��ID = rsTmp!��鱨��ID & ""
    End If
    If lng����ID = 0 And str��鱨��ID = "" Then
        CheckEPRReport = 0: Exit Function
    End If
    
    If Not blnBySign Then
        '��鱨��ִ�й���(5-���;6-�������)��״̬(1-���)
        '���鱨���ǹ������ɼ���ʽ����ģ����ɼ���ʽ����Ϊ����δ�������ͼ�¼
        strSQL = _
            " Select 2 as ����,ҽ��ID,ִ�й���,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
            " Union ALL" & _
            " Select ����,ҽ��ID,ִ�й���,Decode([2],-999,ִ��״̬,[2]) as ִ��״̬,����ʱ��" & _
            " From (" & _
                " Select 1 as ����,B.ҽ��ID,B.ִ�й���,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
                " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                    " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
                " Order by A.���" & _
            " ) Where Rownum=1" & _
            " Order by ����,����ʱ�� Desc"
        If blnMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lngҽ��ID, intִ��״̬)
        If NVL(rsTmp!ִ�й���, 0) >= 5 Or NVL(rsTmp!ִ��״̬, 0) = 1 Then
            CheckEPRReport = 1
        Else
            CheckEPRReport = 2
        End If
    Else
        'ͨ��ǩ���汾�жϱ�����ɵķ�ʽ
        strSQL = "Select B.�ļ�ID,Max(B.��ʼ��) as ǩ���汾 From ���Ӳ������� B Where B.�ļ�ID=[1] And B.��������=8 Group by B.�ļ�ID"
        strSQL = "Select B.���ʱ��,B.���汾,C.ǩ���汾 From ���Ӳ�����¼ B,(" & strSQL & ") C Where B.ID=[1] And B.ID=C.�ļ�ID(+)"
        If blnMoved Then
            strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
            strSQL = Replace(strSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng����ID)
            
        '(ǩ������ֱ���޸ģ������޶������ǩ�������汾Ӧ��ǩ���汾һ��)
        If IsNull(rsTmp!���ʱ��) Or NVL(rsTmp!���汾, 0) <> NVL(rsTmp!ǩ���汾, 0) Then
            '���ҽ�������Ѿ�ִ��,��ʹû��ǩ���򲻷�Ҳ��ͬ���
            strSQL = _
                " Select 2 as ����,ҽ��ID,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
                " Union ALL" & _
                " Select ����,ҽ��ID,Decode([2],-999,ִ��״̬,[2]) as ִ��״̬,����ʱ��" & _
                " From (" & _
                    " Select 1 as ����,B.ҽ��ID,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
                    " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                        " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
                    " Order by A.���" & _
                " ) Where Rownum=1" & _
                " Order by ����,����ʱ�� Desc"
            If blnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lngҽ��ID, intִ��״̬)
            If NVL(rsTmp!ִ��״̬, 0) = 1 Then
                CheckEPRReport = 1
            Else
                CheckEPRReport = 2
            End If
        Else
            CheckEPRReport = 1
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub GetTestLabel(ByVal strScript As String, ByVal strSelect As String, strLabel As String, intResult As Integer)
'���ܣ���ȡƤ�Ա�ע�ͽ��
'������strScript=Ƥ�Խ������������"����(+),������(++);����(-)"
'      strSelect=��ѡ���Ƥ�Խ������������"����"
'���أ�strLabel = Ƥ�Խ����ע����"(+)"
'      intResult=Ƥ�Խ����0-���ԣ�1-����
    Dim arr���� As Variant, arr���� As Variant
    Dim i As Integer
    
    strLabel = "": intResult = 0
    
    arr���� = Split(Split(strScript, ";")(0), ",")
    arr���� = Split(Split(strScript, ";")(1), ",")
    
    For i = 0 To UBound(arr����)
        If arr����(i) Like strSelect & "(*)" Then
            strLabel = Mid(arr����(i), Len(strSelect) + 1)
            intResult = 1: Exit Sub
        End If
    Next
    For i = 0 To UBound(arr����)
        If arr����(i) Like strSelect & "(*)" Then
            strLabel = Mid(arr����(i), Len(strSelect) + 1)
            intResult = 0: Exit Sub
        End If
    Next
End Sub


Public Function ItemHaveCash(ByVal int������Դ As Integer, ByVal bln����ִ�� As Boolean, ByVal lngҽ��ID As Long, ByVal lng���ID As Long, _
    ByVal lng���ͺ� As Long, ByVal str��� As String, ByVal str���ݺ� As String, ByVal int��¼���� As Integer, ByVal int������� As Integer, ByVal int��ʽ As Integer, _
    Optional ByVal blnMove As Boolean, Optional ByVal dat����ʱ�� As Date, Optional ByRef strҽ��IDs As String, Optional ByRef strNOs As String, Optional ByRef blnIsAbnormal As Boolean) As Boolean
'���ܣ��жϵ�ǰ��ִ��ҽ���Ƿ����շѻ���ʻ��۵��Ƿ������
'������int������Դ=1-����,2-סԺ
'      str���=����������ڴ�һ��ҽ�������ַֿ�ִ�е�����
'      int��ʽ=0-����Ƿ����δ�շѼ�¼
'              1-����Ƿ�������շѼ�¼
'      int�������=1=סԺ���͵��������
'      ���أ�strҽ��IDs=��ҽ������ص�ҽ��ID,NOs=ҽ�����͵ĵ��ݺźͲ��ĸ����еĵ��ݺ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTab As String
    
    If int������Դ = 2 And int��¼���� = 2 And int������� = 0 Then
        strTab = "סԺ���ü�¼"
    Else
        strTab = "������ü�¼"
    End If
    ItemHaveCash = True
    strҽ��IDs = ""
    strNOs = ""
    
    '��Ӧ�ķ������Ƿ����δ�շ�[��������]������
    '���嵥ֻ��ʾ���շѲ�ͬ��
    '1.�����ҽ������(���Ӽ�¼���ʵ���������Ϊ���ܲ��շѵ�����ʵ�)
    '2.���ʻ���Ҳ��ʾΪδ��(�嵥��Ҫ���Գ���ִ�к����)
    '3.��NO��Ӧ�����ҽ���ķ��ü��(�嵥�ǰ���ʾ��ҽ��ID)
    strSQL = _
        " Select A.��¼״̬,Nvl(B.���ID,B.ID) as ҽ��ID,B.�������,A.ִ��״̬,A.NO" & IIF(strTab = "סԺ���ü�¼", ",0 as ����״̬", ",NVL(A.����״̬,0) as ����״̬") & _
        " From " & strTab & " A,����ҽ����¼ B" & _
        " Where A.NO=[4] And A.��¼״̬ IN(0,1,3) And A.ҽ�����+0=B.ID And MOD(A.��¼����,10)=[5]" & IIF(bln����ִ��, " And B.ID=[2]", "") & _
        " Union ALL " & _
        " Select B.��¼״̬,Nvl(C.���ID,C.ID) as ҽ��ID,C.�������,B.ִ��״̬,A.NO" & IIF(strTab = "סԺ���ü�¼", ",0 as ����״̬", ",NVL(b.����״̬,0) as ����״̬") & _
        " From ����ҽ����¼ C," & strTab & " B,����ҽ������ A" & _
        " Where A.NO=B.NO And A.��¼����=MOD(B.��¼����,10) And A.ҽ��ID=B.ҽ�����+0" & IIF(bln����ִ��, " And A.ҽ��ID=[2]", _
            " And A.ҽ��ID IN (Select ID From ����ҽ����¼ Where (ID=[1] Or ���ID=[1]) And �������=[6])") & _
        " And A.���ͺ�=[3] And B.��¼״̬ IN(0,1,3) And A.ҽ��ID=C.ID And A.��¼����=[5]"
    If blnMove Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, strTab, "H" & strTab)
    ElseIf zlDatabase.DateMoved(dat����ʱ��) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, strTab, "H" & strTab)
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ItemHaveCash", IIF(lng���ID <> 0, lng���ID, lngҽ��ID), lngҽ��ID, lng���ͺ�, str���ݺ�, int��¼����, str���)
    If Not rsTmp.EOF Then
        If int��ʽ = 0 Then
            rsTmp.Filter = "ҽ��ID=" & IIF(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ����״̬=1"
            If Not rsTmp.EOF Then
                blnIsAbnormal = True
                ItemHaveCash = False
            Else
                rsTmp.Filter = "ҽ��ID=" & IIF(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ��¼״̬=0"
                If Not rsTmp.EOF Then ItemHaveCash = False
            End If
            
            While Not rsTmp.EOF
                If InStr("," & strҽ��IDs & ",", "," & rsTmp!ҽ��ID & ",") = 0 Then
                    strҽ��IDs = strҽ��IDs & "," & rsTmp!ҽ��ID
                End If
                If InStr("," & strNOs & ",", "," & rsTmp!NO & ",") = 0 Then
                    strNOs = strNOs & "," & rsTmp!NO
                End If
                rsTmp.MoveNext
            Wend
            strNOs = Mid(strNOs, 2)
            strҽ��IDs = Mid(strҽ��IDs, 2)
        ElseIf int��ʽ = 1 Then
            rsTmp.Filter = "ҽ��ID=" & IIF(lng���ID <> 0, lng���ID, lngҽ��ID) & " And �������='" & str��� & "' And ��¼״̬<>1 And ����״̬<>1"
            If Not rsTmp.EOF Then ItemHaveCash = False
        End If
    ElseIf int��ʽ = 1 Then
        ItemHaveCash = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceMoney(ByVal str��ID As String, ByVal strҽ��ID As String, ByVal str���ͺ� As String, _
    str��� As String, str����� As String, ByVal bln����ִ�� As Boolean, ByVal byt��Դ As Byte) As Currency
'���ܣ�����ָ����ҽ��ID������ȡҽ����Ӧδ��˵ļ��ʷ��úϼ�
'������str��ID,strҽ��ID,str���ͺ�="ID1,ID2,..."
'      bln����ִ��=������Ŀ����ִ�У���ʱֻ��һ��ҽ��ID
'      byt��Դ��1:���2-סԺ
'���أ�str���,str�����=���ڱ�����ʾ
'˵������ϵͳ����Ϊִ�к���˷���ʱ�ŷ��ء�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, curMoney As Currency
    Dim strTab As String
    
    str��� = "": str����� = ""
    
    On Error GoTo errH
     
    strTab = IIF(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
    
    If bln����ִ�� Then
        strSQL = _
            " Select B.����,B.����,Sum(A.ʵ�ս��) as ���" & _
            " From " & strTab & " A,�շ���Ŀ��� B" & _
            " Where A.ҽ����� + 0 = [2] And (A.��¼����, A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3]" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3])" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����=B.����" & _
            " Group by B.����,B.����"
    Else
        strSQL = _
            " Select /*+ RULE */ B.����,B.����,Sum(A.ʵ�ս��) as ���" & _
            " From " & strTab & " A,�շ���Ŀ��� B" & _
            " Where A.ҽ����� + 0 In" & _
            "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
            "       Union All" & _
            "       Select ID From ����ҽ����¼" & _
            "       Where ���id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "  And (A.��¼����, A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������" & _
            "       Where ҽ��id In" & _
                "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
                "       Union All" & _
                "       Select ID From ����ҽ����¼" & _
                "       Where ���id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "         And ���ͺ� + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)))" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������" & _
            "       Where ҽ��id In (Select Column_Value From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)))" & _
            "         And ���ͺ� + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))))" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����=B.����" & _
            " Group by B.����,B.����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceMoney", str��ID, strҽ��ID, str���ͺ�, glngSys)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + NVL(rsTmp!���, 0)
        str��� = str��� & rsTmp!����
        str����� = str����� & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    
    str����� = Mid(str�����, 2)
    GetAdviceMoney = curMoney
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceStuffMoney(ByVal str��ID As String, ByVal strҽ��ID As String, _
    ByVal str���ͺ� As String, ByVal bln����ִ�� As Boolean, ByVal int������Դ As Integer, ByVal int��¼���� As Integer, ByVal int������� As Integer) As Currency
'���ܣ�����ָ����ҽ��ID������ȡҽ����Ӧδ��˵ĸ����������ļ��ʷ��úϼ�
'������str��ID,strҽ��ID,str���ͺ�="ID1,ID2,..."
'      bln����ִ��=������Ŀ����ִ�У���ʱֻ��һ��ҽ��ID
'      int������Դ��1:���2-סԺ
'      int�������=1=סԺ���͵��������
'���أ�str���,str�����=���ڱ�����ʾ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    
    On Error GoTo errH
    If int������Դ = 2 And int��¼���� = 2 And int������� = 0 Then
        strTab = "סԺ���ü�¼"
    Else
        strTab = "������ü�¼"
    End If
    
    If bln����ִ�� Then
        strSQL = _
            " Select Sum(A.ʵ�ս��) as ���" & _
            " From " & strTab & " A,�������� B" & _
            " Where A.ҽ����� + 0 = [2] And (A.��¼����, A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3]" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������ Where ҽ��id = [2] And ���ͺ� + 0 = [3])" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����='4' And A.�շ�ϸĿID=B.����ID And B.��������=1"
    Else
        strSQL = _
            " Select /*+ RULE */ Sum(A.ʵ�ս��) as ���" & _
            " From " & strTab & " A,�������� B" & _
            " Where A.ҽ����� + 0 In" & _
            "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
            "       Union All" & _
            "       Select ID From ����ҽ����¼" & _
            "       Where ���id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "  And (A.��¼����, A.NO) In" & _
            "      (Select ��¼����, NO From ����ҽ������" & _
            "       Where ҽ��id In" & _
                "      (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))" & _
                "       Union All" & _
                "       Select ID From ����ҽ����¼" & _
                "       Where ���id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            "         And ���ͺ� + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)))" & _
            "       Union All" & _
            "       Select ��¼����, NO From ����ҽ������" & _
            "       Where ҽ��id In (Select Column_Value From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)))" & _
            "         And ���ͺ� + 0 In (Select Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))))" & _
            "  And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ����='4' And A.�շ�ϸĿID=B.����ID And B.��������=1"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceStuffMoney", str��ID, strҽ��ID, str���ͺ�, glngSys)
    If Not rsTmp.EOF Then GetAdviceStuffMoney = NVL(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ItemCanCancel(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal lng��ID As Long, str������� As String, _
    ByVal bln����ִ�� As Boolean, ByVal blnMove As Boolean, ByVal byt��Դ As Byte) As Boolean
'���ܣ��ж�ָ����Ŀ�Ƿ����ȡ��ִ��
'������byt��Դ=1:���2-סԺ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If gbytBillOpt = 0 Then ItemCanCancel = True: Exit Function
    
    On Error GoTo errH
    
    If bln����ִ�� Then
        strSQL = _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ҽ��ID=[1] And ���ͺ�=[2]" & _
            " Union ALL " & _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ҽ��ID=[1] And ���ͺ�=[2]"
    Else
        strSQL = _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ҽ��ID=[1] And ���ͺ�=[2]" & _
            " Union ALL " & _
            " Select Distinct NO From ����ҽ������ Where ��¼����=2 And ���ͺ�=[2]" & _
            " And ҽ��ID IN(Select ID From ����ҽ����¼ Where (ID=[3] Or ���ID=[3]) And �������=[4])"
    End If
    If blnMove Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ItemCanCancel", lngҽ��ID, lng���ͺ�, lng��ID, str�������)
    
    Do While Not rsTmp.EOF
        '�������ſ��˽��ʽ��Ϊ0�ģ�����ķ��õǼ�
        If HaveBilling(rsTmp!NO, True, "", IIF(bln����ִ��, lngҽ��ID, 0), byt��Դ) <> 0 Then
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("����Ŀ�����Ѿ����ʵķ���,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Case 2
                    MsgBox "����Ŀ�����Ѿ����ʵķ���,�������ܼ�����", vbExclamation, gstrSysName
                    Exit Function
            End Select
        End If
        rsTmp.MoveNext
    Loop
    ItemCanCancel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetUser����IDs(Optional ByVal bln���� As Boolean) As String
'���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
'�������Ƿ�ȡ���������µĿ���
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, blnNew As Boolean
    
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    'û��ǿ�������ٴ�,����ҽ��������
    If blnNew Then
        strSQL = "Select 1 as ���,����ID From ������Ա Where ��ԱID=[1] Union" & _
                " Select Distinct 2 as ���,B.����ID From ������Ա A,�������Ҷ�Ӧ B" & _
                " Where A.����ID=B.����ID And A.��ԱID=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISJob", UserInfo.ID)
    End If
    If bln���� = False Then
        rsTmp.Filter = "��� = 1"
    Else
        rsTmp.Filter = ""
    End If
    
    For i = 1 To rsTmp.RecordCount
        If InStr("," & GetUser����IDs & ",", "," & rsTmp!����ID & ",") = 0 Then
            GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser����IDs() As String
'���ܣ���ȡ����Ա�����Ĳ���(ֱ�����ڲ��������ڿ��������Ĳ���),�����ж��
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, blnNew As Boolean
        
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    If blnNew Then
        strSQL = _
            "Select Distinct ����ID From (" & _
            " Select A.����ID as ����ID" & _
            " From ��������˵�� A,������Ա B" & _
            " Where A.����ID=B.����ID And B.��ԱID=[1]" & _
            " And A.������� in(1,2,3) And A.��������='����'" & _
            " Union" & _
            " Select A.����ID From �������Ҷ�Ӧ A,������Ա B" & _
            " Where A.����ID=B.����ID And B.��ԱID=[1])"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
    ElseIf rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
    End If
    For i = 1 To rsTmp.RecordCount
        GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveBilling(ByVal strNO As String, ByVal blnALL As Boolean, _
     ByVal strTime As String, ByVal lngҽ��ID As Long, ByVal byt��Դ As Byte) As Integer
'���ܣ��ж�һ�ż��ʵ�/���Ƿ��Ѿ�����
'������strNO=���ʵ��ݺ�,�������ＰסԺ
'      blnALL=�Ƿ�����ŵ������ݽ����ж�,����ֻ��δ���ʲ��ֽ����ж�(����ʱ)
'      byt��Դ=1:���2-סԺ
'���أ�0-δ����,1=��ȫ������,2-�Ѳ��ֽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long
    Dim strTab As String
    
    On Error GoTo errH
    strTab = IIF(byt��Դ = 1, "������ü�¼", "סԺ���ü�¼")
        
    '��δ���ϵķ�����
    strSQL = _
        " Select ��� From (" & _
        " Select ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���) as ���," & _
        " Avg(Nvl(����, 1) * ����) As ����" & _
        " From " & strTab & "" & _
        " Where NO=[1] And ��¼����=2" & _
        " Group by ��¼״̬,ִ��״̬,Nvl(�۸񸸺�,���))" & _
        " Group by ��� Having Sum(����)<>0"
    
    '��ÿ�еĽ������
    strSQL = _
        "Select Nvl(�۸񸸺�,���) as ���,Sum(Nvl(���ʽ��,0)) as ���ʽ��" & _
        " From " & strTab & "" & _
        " Where NO=[1] And ��¼���� IN(2,12)" & _
        IIF(Not blnALL, " And Nvl(�۸񸸺�,���) IN(" & strSQL & ")", "") & _
        IIF(strTime <> "", " And �Ǽ�ʱ��=[2]", "") & _
        IIF(lngҽ��ID <> 0, " And ҽ�����+0=[3]", "") & _
        " Group by Nvl(�۸񸸺�,���)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "HaveBilling", strNO, CDate(IIF(strTime = "", "1990-01-01", strTime)), lngҽ��ID)
    If Not rsTmp.EOF Then
        lngTmp = rsTmp.RecordCount '��������
        rsTmp.Filter = "���ʽ��<>0"
        If rsTmp.EOF Then
            HaveBilling = 0 '�޽�����
        ElseIf rsTmp.RecordCount = lngTmp Then
            HaveBilling = 1 'ȫ�����ѽ���
        ElseIf rsTmp.RecordCount > 0 Then
            HaveBilling = 2 '�������ѽ���
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckKSSPrivilege(ByVal int���� As Integer) As Boolean
'���ܣ����ϵͳ�Ƿ���ڿ���ҩ����Ȩ����Ա���������õ�ǰ����Ա����ҩ����UserInfo����
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    UserInfo.��ҩ���� = 0
    
    On Error GoTo errH
    strSQL = "Select ���� From ��Ա����ҩ��Ȩ�� Where ��¼״̬=1 and ��ԱID = [1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID, int����)
    If rsTmp.RecordCount > 0 Then
        UserInfo.��ҩ���� = Val("" & rsTmp!����)
        CheckKSSPrivilege = True
    Else
        strSQL = "Select 1 From ��Ա����ҩ��Ȩ�� Where ��¼״̬=1 and Rownum<2 And ����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", int����)
        CheckKSSPrivilege = rsTmp.RecordCount > 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitAdviceDefine() As Recordset
'���ܣ���ȡҽ�����ݶ����¼��
'������blnNew-�Ƿ񴴽�objVBA��objScript����
'˵����
    Dim strSQL As String
    Dim rsDefine As Recordset
    

    On Error GoTo errH
    strSQL = "Select �������,ҽ������ From ҽ�����ݶ��� Order by �������"
    Set rsDefine = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsDefine, strSQL, "InitAdviceDefine")
    Set InitAdviceDefine = rsDefine
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckSign(ByVal intǩ������ As Long, ByVal lng��������ID As Long, Optional ByVal lngҽ������ID As Long, Optional ByVal lng���˿���ID As Long, _
    Optional ByVal int���˷�Χ As Integer = 2, Optional ByVal blnCheckCert As Boolean = True, Optional ByRef objESign As Object, Optional ByVal str����ҽ�� As String) As Boolean
'���ܣ��ж�һ�����Ż���һ�鲿�����Ƿ���������˵���ǩ�����Ƶ�
'������int���˷�Χ=1-����,2-סԺ(ȱʡ)
'     intǩ������:0-����ҽ���Ͳ�����1-סԺҽ��ҽ���Ͳ�����2-סԺ��ʿҽ����3-ҽ��ҽ���ͱ��棻4-�����¼�ͻ�������5-ҩƷ��ҩ��6-LIS;7-PACS;
'     lng��������ID=���lng��������ID=0������Ҫ���ݴ����ҽ�����ң����˿���ID���Ӧ��Ĭ�Ͽ�������
'                   ��ʿվУ�Ժ�ȷ��ֹͣʱ������Ĳ���ID�����жϲ����Ƿ������˵���ǩ��
'                   ����-1������ҩ�����ʱ������ж��Ƿ�ֿ������ã�
'     blnCheckCert=true ���֤���Ƿ�ͣ�ã�=false��ʾ�����
'     str����ҽ��=ͬʱ���ĳ������Ա�Ĳ���Ա֤���Ƿ�ͣ�ã����磺��ʿУ��ʱ�����ҽ���Ƿ���Ҫǩ��
    Dim strSQL As String, intTmp As Integer
    Dim rsTmp As Recordset
    
    '������϶�δ���ã��򷵻�false
    If intǩ������ = 0 Or intǩ������ = 1 Then
        intTmp = intǩ������ + 1
    ElseIf intǩ������ > 1 And intǩ������ <= 7 Then
        intTmp = intǩ������
    End If
    If Mid(gstrESign, intTmp, 1) <> "1" Then Exit Function
    If lng��������ID = 0 And (lng���˿���ID <> 0 Or lngҽ������ID <> 0) Then
        'ȡ��������
        lng��������ID = Get��������ID(UserInfo.ID, lngҽ������ID, lng���˿���ID, int���˷�Χ)
        If lng��������ID = 0 Then Exit Function
    End If
    grsSign.Filter = "����ID=" & lng��������ID & " and ����=" & intǩ������
    If grsSign.RecordCount = 0 Then
        strSQL = "Select Zl_Fun_Getsignpar([1],[2]) as �Ƿ����� From dual"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlAdvice", intǩ������, lng��������ID)
        If rsTmp.RecordCount > 0 Then
            CheckSign = Val(rsTmp!�Ƿ����� & "") = 1
            grsSign.AddNew
            grsSign!����ID = lng��������ID
            grsSign!���� = intǩ������
            grsSign!�Ƿ����� = Val(rsTmp!�Ƿ����� & "")
        End If
    Else
        grsSign.MoveFirst
        CheckSign = Val(grsSign!�Ƿ����� & "") = 1
    End If
    If CheckSign = True And blnCheckCert Then
        If objESign Is Nothing Then
            On Error Resume Next
            Set objESign = CreateObject("zl9ESign.clsESign")
            err.Clear: On Error GoTo 0
            If Not objESign Is Nothing Then
                Call objESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        '���֤���Ƿ�ͣ��
        If objESign.CertificateStoped(UserInfo.����) Then CheckSign = False
        If str����ҽ�� <> "" Then If objESign.CertificateStoped(str����ҽ��) Then CheckSign = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��������ID(ByVal lngҽ��ID As Long, ByVal lngҽ������ID As Long, ByVal lng���˿���ID As Long, _
    Optional ByVal int��Χ As Integer = 2, Optional ByVal lngִ�п���ID As Long, Optional ByVal lng�������ID As Long) As Long
'���ܣ���ҽ��ȷ����������
'������int��Χ=1-����,2-סԺ(ȱʡ)
'˵������ҽ���������ҷ�Χ��,����˳�����£�
'      1��ҽ������(ҽ������)
'      2���������
'      3�����˿���
'      4������������/סԺ���˵�ĳЩ����ҽ����ִ�п���
'      5������������/סԺ���˵Ŀ�����ΪĬ�Ͽ���
'      6������������/סԺ���˵Ŀ���
'      7��Ĭ�Ͽ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr����ID(1 To 7) As Long
    
    '�������ű������ٴ���ҽ��
    strSQL = "Select Distinct A.����ID,Nvl(A.ȱʡ,0) as ȱʡ" & _
        " From ������Ա A,��������˵�� B,���ű� C" & _
        " Where A.����ID=C.ID And A.����ID=B.����ID" & _
        " And B.������� IN([2],3) And A.��ԱID=[1]" & _
        " And B.�������� IN('�ٴ�','���','����','����','����','Ӫ��','����')" & _
        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҽ��ID, int��Χ)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!����ID = lngҽ������ID Then
            arr����ID(1) = rsTmp!����ID
        ElseIf rsTmp!����ID = lng�������ID Then
            arr����ID(2) = rsTmp!����ID
        ElseIf rsTmp!����ID = lng���˿���ID Then
            arr����ID(3) = rsTmp!����ID
        ElseIf rsTmp!����ID = lngִ�п���ID Then
            arr����ID(4) = rsTmp!����ID
        ElseIf rsTmp!ȱʡ = 1 Then
            arr����ID(5) = rsTmp!����ID
        ElseIf arr����ID(5) = 0 Then
            arr����ID(6) = rsTmp!����ID
        End If
        rsTmp.MoveNext
    Next
    arr����ID(7) = UserInfo.����ID
    
    For i = LBound(arr����ID) To UBound(arr����ID)
        If arr����ID(i) <> 0 Then
            Get��������ID = arr����ID(i)
            Exit For
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Function CreatePlugInOK(ByVal lngMod As Long, Optional ByVal int���� As Integer) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod, int����)
        Call zlPlugInErrH(err, "Initialize")
        err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
    
End Function

Public Function Get��Һ��������() As String
'���ܣ���ȡ��Һ�������ĵĿ���IDs
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strReturn As String
    
    On Error GoTo errH

    strSQL = "Select ����id From ��������˵�� Where �������� = '��������' Order by ����id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get��Һ��������")
    
    For i = 1 To rsTmp.RecordCount
        strReturn = strReturn & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    Get��Һ�������� = Mid(strReturn, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HavePath(ByVal lng����ID As Long) As Boolean
'���ܣ����ָ�����һ����Ƿ��п��õ��ٴ�·��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "Select a.Id" & vbNewLine & _
            "From �ٴ�·��Ŀ¼ A, �ٴ�·���汾 B, �ٴ�·������ C," & vbNewLine & _
            "     (Select ����id From �������Ҷ�Ӧ Where ����id = [1]" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select ID From ���ű� Where ID = [1]) D" & vbNewLine & _
            "Where a.Id = b.·��id And a.���°汾 = b.�汾�� And a.Id = c.·��id(+) And (c.����id = d.����id or c.����id is null) And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID)
    HavePath = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get����ID(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����id As Long, Optional ByRef bln��ҽ As Boolean = False) As ADODB.Recordset
'˵��:��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����
'    �������:1-��ҽ�������;2-��ҽ��Ժ���;11-��ҽ�������;12-��ҽ��Ժ���
'    �ж����ϵ�����£�������ϴ���ֻȡ��һ����Ҫ���
'    ���������������ȣ���Ҫ��Ϊ��֧��������ϡ�
'--- �������ҽ��ʱ, ���ȼ�����ҽ��Ժ����ҽ��Ժ����ҽ�����ҽ����
'    �ų�����¼������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    bln��ҽ = Sys.DeptHaveProperty(lng����id, "��ҽ��")
    
    If bln��ҽ Then
        strSQL = "Select ����id, ���id, �������, �������, ��¼��Դ" & vbNewLine & _
                "From ������ϼ�¼" & vbNewLine & _
                "Where ��¼��Դ In (1, 2, 3) And ������� In (1, 2, 11, 12) And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� = 1 And" & vbNewLine & _
                "      Nvl(�Ƿ�����, 0) = 0 and Not (NVl(����ID,0)=0 and NVl(���ID,0)=0) " & vbNewLine & _
                "Order By Decode(�������, 12, 1, 2, 2, 11, 3, 1, 4), Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc"
    Else
        strSQL = "Select ����id, ���id, �������" & vbNewLine & _
        "From ������ϼ�¼" & vbNewLine & _
        "Where ��¼��Դ In (1, 2, 3) And ������� In (1, 2, 11, 12) And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� = 1 And" & vbNewLine & _
        "   Nvl(�Ƿ�����,0) = 0 and Not (NVl(����ID,0)=0 and NVl(���ID,0)=0) " & vbNewLine & _
        "Order By Sign(�������-10),������� Desc, Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc"
    End If
                
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", lng����ID, lng��ҳID)
    Set Get����ID = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPathTable(lng����ID As Long, lng���ID As Long, lng����id As Long) As ADODB.Recordset
    Dim strSQL As String
 
    strSQL = "Select a.Id, a.����, a.����, a.����, a.˵��, Nvl(a.���ò���,'ͨ��') ���ò���, a.�����Ա�, a.��������, a.���°汾, c.��׼סԺ��,Nvl(a.��������,'��') as ��������,Nvl(a.ȷ������,0) as ȷ������" & vbNewLine & _
            "From �ٴ�·��Ŀ¼ A, �ٴ�·������ B,�ٴ�·���汾 C" & vbNewLine & _
            "Where a.Id = b.·��id And (b.����id = [1] Or b.���id = [2]) And a.���°汾 is not null And a.id = b.·��ID And a.���°汾 = c.�汾��" & vbNewLine & _
            "And a.Id = c.·��id And b.����=0 And (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From �ٴ�·������ D Where a.Id = d.·��id And d.����id = [3]))"
    On Error GoTo errH
    Set GetPathTable = zlDatabase.OpenSQLRecord(strSQL, "��ȡ·��Ŀ¼", lng����ID, lng���ID, lng����id)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim dtCurDate As Date, strMaxNo As String
    Dim strYearStr As String
    
    err = 0: On Error GoTo errH:
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = zlStr.PrefixNO & strNO
        Exit Function
    End If
'    ElseIf intNum = 0 Then
'        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
'        Exit Function
'    End If
    GetFullNO = strNO
    
    strSQL = "Select ��Ź���,Sysdate as ����,������ From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, intNum)
    dtCurDate = datE
    If Not rsTmp.EOF Then
        intType = Val("" & rsTmp!��Ź���)
        dtCurDate = rsTmp!����
        strMaxNo = NVL(rsTmp!������)
    End If
    strYearStr = zlStr.PrefixNO
    If strMaxNo = "" Then strMaxNo = strYearStr & "000001"
    If intType = 1 Then
        '���ձ��
        strSQL = Format(CDate(Format(dtCurDate, "YYYY-MM-dd")) - CDate(Format(dtCurDate, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = zlStr.PrefixNO & strSQL & Format(Right(strNO, 4), "0000")
        Exit Function
    End If
    '������
    If Len(strNO) = 6 Then
        GetFullNO = Left(strMaxNo, 2) & strNO: Exit Function
    End If
    GetFullNO = Left(strMaxNo, 2) & zlStr.LPAD(Right(strNO, 6), 6, "0", True)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CreateObjectPacs(objPublicPACS As Object) As Boolean
    If objPublicPACS Is Nothing Then
        On Error Resume Next
        Set objPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        err.Clear: On Error GoTo 0
        If Not objPublicPACS Is Nothing Then
            Call objPublicPACS.InitInterface(gcnOracle, UserInfo.�û���)
        End If
        If objPublicPACS Is Nothing Then
            MsgBox "PACS��������δ�����ɹ���", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreateObjectPacs = True
End Function

Public Function PatiFeeUsable(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ����˵ĵ�ǰ�����Ƿ���Ч������true������ǰ�ѱ����
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnʧЧ As Boolean
    
    On Error GoTo errH
    strSQL = "Select  Sysdate as ��ǰ,Nvl(b.��Ч��ʼ, To_Date('1900-01-01', 'yyyy-mm-dd')) as ��ʼ,Nvl(b.��Ч����, To_Date('3000-01-01', 'yyyy-mm-dd')) as ����  From ������ҳ A, �ѱ� B Where a.�ѱ�=b.���� And a.����id=[1] And a.��ҳid=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiFeeUsable", lng����ID, lng��ҳID)
    PatiFeeUsable = True
    
    If rsTmp.EOF Then
        blnʧЧ = True
    Else
        If Not Between(Format(rsTmp!��ǰ, "YYYY-MM-DD"), Format(rsTmp!��ʼ, "YYYY-MM-DD"), Format(rsTmp!����, "YYYY-MM-DD")) Then
            blnʧЧ = True
        End If
    End If
    
    If blnʧЧ Then
        MsgBox "�ò��˵ĵ�ǰ�ѱ��Ѿ�ʧЧ�����ܷ���ҽ�������ڲ�����Ϣ�е������˷ѱ�", vbInformation, gstrSysName
        PatiFeeUsable = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitObjPublicExpense() As Boolean
    If gobjPublicExpense Is Nothing Then
        On Error Resume Next
        Set gobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Not gobjPublicExpense Is Nothing Then
            Call gobjPublicExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
        End If
        err.Clear: On Error GoTo 0
    End If
    InitObjPublicExpense = Not gobjPublicExpense Is Nothing
End Function

Public Function GetPatiAge(ByVal lng����ID As Long) As String
'���ܣ���������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    strSQL = "select zl_age_calc([1]) as ���� from dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiAge", lng����ID)
    GetPatiAge = rsTmp!���� & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveRIS(Optional ByVal blnMsg As Boolean) As Boolean
'���ܣ��ж� RIS�ӿڲ��� �Ƿ���ڣ�������
'������blnMsg������ʧ��ʱ�Ƿ���ʾ
    If Not gbln����Ӱ����Ϣϵͳ�ӿ� Then Exit Function
    If gobjRis Is Nothing Then
        On Error Resume Next
        Set gobjRis = CreateObject("zl9XWInterface.clsHISInner")
        err.Clear: On Error GoTo 0
        If Not gobjRis Is Nothing Then
            gbln����Ӱ����ϢϵͳԤԼ = gobjRis.HISSchedulingjudge = 0
        End If
    End If
    If gobjRis Is Nothing Then
        If blnMsg Then
            MsgBox "RIS�ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    HaveRIS = True
End Function

Public Function GetPriceGradeSQL(ByVal strҩƷ�۸�ȼ� As String, ByVal str���ļ۸�ȼ� As String, ByVal str��ͨ��Ŀ�۸�ȼ� As String, ByVal strTableTmpA As String, ByVal strTableTmpB As String, _
           ByVal strParNumҩƷ As String, ByVal strParNum���� As String, ByVal strParNum��ͨ��Ŀ As String) As String
'���ܣ����˼۸�ȼ����������ȡ�۸��SQL
'������strҩƷ�۸�ȼ�  '���˵�ҩƷ�۸�ȼ�
'      str���ļ۸�ȼ�  '���˵����ļ۸�ȼ�
'      str��ͨ��Ŀ�۸�ȼ�  '���˵���ͨ��Ŀ�۸�ȼ�
'     strTableTmpA   �շ���ĿĿ¼ ���as ��־,strTableTmpB  �շѼ�Ŀ�� ��As��־��
'     strParNumҩƷ  ҩƷ�۸�ȼ�SQL�������,strParNum����  ���ļ۸�ȼ�SQL�������,strParNum��ͨ��Ŀ  ��ͨ��Ŀ�۸�ȼ�SQL�������
    Dim strSQL As String
    
    If strҩƷ�۸�ȼ� = "" And str���ļ۸�ȼ� = "" And str��ͨ��Ŀ�۸�ȼ� = "" Then
        strSQL = " And " & strTableTmpB & ".�۸�ȼ� is Null "
    Else
        strSQL = " And" & vbNewLine & _
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
    
    GetPriceGradeSQL = strSQL
End Function

Public Function GetTimPointsInRange(ByVal strTimRange As String, ByVal strTimPoints As String) As String
'���ܣ���ȡ��ָ��ʱ�䷶Χ�ڵ�ʱ���
'������
'      strTimRange ʱ�䷶Χ �� "2013-08-01 9:36:40,2013-08-02 23:59:59"
'      strTimPoints ʱ���  �� "2013-08-01 08:00:00,2013-08-01 12:00:00,2013-08-01 16:00:00"
'���أ��ڷ�Χ�ڵ�ʱ���     �� "2013-08-01 12:00:00,2013-08-01 16:00:00"
'ע�⣺����ʱӦ�ȱ�֤��������ʽ��ȷ��ʱ���֮�䶼���� ���� �ָ�
    Dim arrTimPoints() As String, i As Long
    Dim strBegin As String, strEnd As String
    Dim strResult As String
    
    strBegin = Split(strTimRange, ",")(0)
    strEnd = Split(strTimRange, ",")(1)
    
    arrTimPoints = Split(strTimPoints, ",")
    
    For i = 0 To UBound(arrTimPoints)
        If Between(Format(arrTimPoints(i), "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            strResult = IIF(strResult = "", "", strResult & ",") & arrTimPoints(i)
        End If
    Next
    
    GetTimPointsInRange = strResult
End Function

Public Function InitObjPublicPatient() As Boolean
'���ܣ���ʼ������Ϣ��������
    If gobjPublicPatient Is Nothing Then
        On Error Resume Next
        Set gobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        If Not gobjPublicPatient Is Nothing Then
            Call gobjPublicPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser)
        End If
        err.Clear: On Error GoTo 0
    End If
    InitObjPublicPatient = Not gobjPublicPatient Is Nothing
End Function
