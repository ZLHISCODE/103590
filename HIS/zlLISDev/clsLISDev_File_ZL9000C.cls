VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Zl9000c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Ѫ����
'�����ͺţ�ZL9000C
'�û�������������ҽԺ
'��ս�У�������
'���ڣ�2014-07-08
'
Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo ErrHand '�򿪴�������
    '------------------------------------------------
    '��׼���������ļ���ȡ������
    '   strFile������·�����ļ���
    '   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
    '   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
    '   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
    '   ���صļ���������(������������밴���±�׼��֯���)
    '   Ԫ��֮����||�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i + 1��Ԫ��Ϊ������

    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, yp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String, strlog As String
    Dim testid As Long
    Dim blnUnion As Boolean
    Dim lngCount As Long, str�� As String, varTmp As Variant
    Dim str��_��1 As String, str��_��2 As String, strŮ_��1 As String, strŮ_��2 As String
    Dim str��_��1 As String, str��_��2 As String, strŮ_��1 As String, strŮ_��2 As String
    Dim strIMG As String
    Dim lngRecord As String
    'д��־
    strlog = "�Ѿ����ò���"
    Call writelislog(strlog)
    ReDim Preserve aResults(0)
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        'д��־
        strlog = "�Ѿ����ļ�"
        Call writelislog(strlog)

        strSQL = "Select ShearRate as X,MaleNormal As ��,FemaleNormal As Ů From ShearRate order by ShearRate"
        Set rsTmp = cnAccess.Execute(strSQL)
        lngRecord = 0
        Do Until rsTmp.EOF
            lngRecord = lngRecord + 1
            rsTmp.MoveNext
        Loop
        lngCount = 0
        If lngRecord > 0 Then rsTmp.MoveFirst
        Do Until rsTmp.EOF
            lngCount = lngCount + 1
            If lngRecord > 3 Then
                If lngCount = 1 Then
                    str��_��1 = "" & Split("" & rsTmp!��, "-")(0) & "," & rsTmp!x
                    str��_��2 = "" & Split("" & rsTmp!��, "-")(1) & "," & rsTmp!x
                    strŮ_��1 = "" & Split("" & rsTmp!Ů, "-")(0) & "," & rsTmp!x
                    strŮ_��2 = "" & Split("" & rsTmp!Ů, "-")(1) & "," & rsTmp!x
                End If
            Else
                If lngCount = 1 Then
                    str��_��1 = "" & Split("" & rsTmp!��, "-")(0) & "," & rsTmp!x
                    str��_��2 = "" & Split("" & rsTmp!��, "-")(1) & "," & rsTmp!x
                    strŮ_��1 = "" & Split("" & rsTmp!Ů, "-")(0) & "," & rsTmp!x
                    strŮ_��2 = "" & Split("" & rsTmp!Ů, "-")(1) & "," & rsTmp!x
                End If
            End If
            If lngCount = lngRecord Then
                str��_��1 = "" & Split("" & rsTmp!��, "-")(0) & "," & rsTmp!x
                str��_��2 = "" & Split("" & rsTmp!��, "-")(1) & "," & rsTmp!x
                strŮ_��1 = "" & Split("" & rsTmp!Ů, "-")(0) & "," & rsTmp!x
                strŮ_��2 = "" & Split("" & rsTmp!Ů, "-")(1) & "," & rsTmp!x
            End If
            rsTmp.MoveNext
        Loop

        strSQL = "": blnUnion = False
        'If dtStart < date Then 'Ҫ��ѯ����

        '�����ں���������ж�λ
        strSQL = " select testdataid,���,�Ա�,���� as Rq,����ҽ�� as bg,'ѪҺ' as �걾,0 as �ʿ�," & _
                 " '01|' +  cstr(iif(isnull(ȫѪճ��),0,ȫѪճ��)) + '|02|' + cstr(iif(isnull(Ѫ��ճ��),0,Ѫ��ճ��)) + '|03|' + cstr(iif(isnull(ѹ��),0,ѹ��)) + '|04|' + cstr(iif(isnull(Ѫ��),0,Ѫ��)) + '|05|' +  cstr(format(ȫѪ�������ָ��,'0.00')) " & _
                 " + '|06|' + cstr(format(ȫѪ�������ָ��,'0.00')) + '|07|' + cstr(format(Ѫ������Kֵ,'0.00')) + '|08|' + cstr(format(��ϸ���ۼ�ָ��,'0.00'))" & _
                 " + '|09|' + cstr(iif(isnull(��ϸ������ָ��),0,��ϸ������ָ��)) + '|10|' + cstr(format(ȫѪ���л�ԭճ��,'0.00')) + '|11|' + cstr(format(ȫѪ���л�ԭճ��,'0.00'))" & _
                 " + '|12|' + cstr(format(��ϸ������ָ��TK,'0.00')) + '|13|' + cstr(format(��ϸ������ָ��,'0.00'))as ���1," & _
                 "'|14|' + cstr(format(����ճ��,'0.00')) + '|15|' + cstr(iif(isnull(��ά����ԭ),0,��ά����ԭ))" & _
                 " + '|16|' + cstr(iif(isnull(Ѫ���̴�),0,Ѫ���̴�)) + '|17|' + cstr(iif(isnull(������֬),0,������֬)) + '|18|' + cstr(iif(isnull(����֬����),0,����֬����)) + '|19|' + cstr(iif(isnull(Ѫ��),0,Ѫ��)) + '|20|' + cstr(iif(isnull(ѪС��ճ����),0,ѪС��ճ����))" & _
                 " + '|21|' + cstr(iif(isnull(����Ѫ˨����),0,����Ѫ˨����)) + '|22|' + cstr(iif(isnull(��ϸ����Ӿ),0,��ϸ����Ӿ)) + '|23|' + cstr(iif(isnull(ѪС��ۼ���),0,ѪС��ۼ���)) + '|24|' + cstr(iif(isnull(����Ѫ˨����),0,����Ѫ˨����))" & _
                 " + '|25|' + cstr(format(��������,'0.00')) + '|26|' + cstr(format(��������,'0.00')) + '|27|' + cstr(format(��ϸ���ۼ�ϵ��,'0.00')) + '|28|' + cstr(format(��ϸ����ճ��,'0.00')) " & _
                 " + '|29|' + cstr(iif(isnull(��ϸ������),0,��ϸ������)) + '|30|' + cstr(format(��������,'0.00')) + '|31|' + cstr(iif(isnull(����Ӧ��),0,����Ӧ��)) + '|32|' + cstr(iif(isnull(ȫѪ���л�ԭճ��),0,ȫѪ���л�ԭճ��)) as ���2" & _
                 " from testdata " & _
                 IIf(lngSampleNO = -1, "where ", " where ���=" & lngSampleNO & " and ") & _
                 " ����>=cdate('" & Format(dtStart, "yyyy-mm-dd") & "')" & _
                 " and ����<=cdate('" & Format(dtEnd, "yyyy-mm-dd 23:59:59") & "')"
        '" + '|25|' + cstr(iif(isnull(Ѫ����������),0,Ѫ����������)) + '|26|' + cstr(iif(isnull(Ѫ����������ʱ��),0,Ѫ����������ʱ��)) as ���2" &

        strSQL = strSQL & IIf(blnUnion, "", " Order By ���� Desc")

        Set rsTmp = cnAccess.Execute(strSQL)
        'cnAccess.

        If rsTmp.EOF Then
            .Close
            'д��־
            strlog = "û���ҵ�����"
            Call writelislog(strlog)
            ResultFromFile = aResults
            Exit Function
        End If
        'д��־
        strlog = "yp:�Ѿ��ҵ�����"
        Call writelislog(strlog)

        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("���") Then '�±걾

                lngResultCount = lngResultCount + 1
                If lngResultCount > 0 Then
                    ReDim Preserve aResults(lngResultCount)
                End If
                aResults(lngResultCount) = Format(rsTmp("Rq"), "yyyy-mm-dd") & " " & Format(Now, "HH:mm:ss") '����ʱ��
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         Val(rsTmp("���")) '�������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         IIf(Trim("" & rsTmp("Bg")) = "", " ", rsTmp("Bg")) '������
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         IIf(IsNull(rsTmp("�걾")), " ", rsTmp("�걾")) '�걾
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         Val(rsTmp("�ʿ�")) '�ʿ�

                strBh = rsTmp("���")

            End If
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                     Replace(rsTmp("���1"), Chr(0), "") + Replace(rsTmp("���2"), Chr(0), "") '���
            '�õ�����id���Դ�visc�еõ�shearRate��Viscֵ
            testid = rsTmp("testdataid")
            strSQL = "select testdataid,shearRate,Visc from visc where testdataid=" & testid & ""
            ' If .State = 1 Then .Close
            Set yp = cnAccess.Execute(strSQL)
            str�� = ""
            Do While Not yp.EOF
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         yp("shearRate") & "|" & yp("visc")

                str�� = str�� & "," & yp("shearRate") & "-" & yp("visc")
                yp.MoveNext
            Loop
            If str�� <> "" Then
                strIMG = vbCrLf & "ճ����������;2;20,200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;"
                If Trim("" & rsTmp("�Ա�")) = True Then
                    '��
                    strIMG = strIMG & str��_��1 & "," & str��_��1 & "|" & str��_��2 & "," & str��_��2
                Else
                    strIMG = strIMG & strŮ_��1 & "," & strŮ_��1 & "|" & strŮ_��2 & "," & strŮ_��2
                End If
                strIMG = strIMG & "~" & Mid(str��, 2)
                strIMG = strIMG & ";VIS(mPa.s),30,20~SHR(1/S),195,1" '������
            End If
            If strIMG <> "" Then aResults(lngResultCount) = aResults(lngResultCount) & strIMG
            '                If str�� <> "" Then
            '                     varTmp = Split(str��, ",")
            '                     If UBound(varTmp) < 4 Then
            '                        str�� = varTmp(0) & "," & varTmp(1) & "," & varTmp(2) & "," & varTmp(3) & "," & varTmp(3)
            '
            '                     End If
            '                    '���ݹ�ʽ��ͼ
            '                    If Trim("" & rsTmp!�Ա�) = True Then    '��
            '                        strImg = "ճ����������;110;1;200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20,;"
            '
            '                        strImg = strImg & str��_��1 & "," & str��_��2 & "|" & str��_��2 & "," & str��_��2
            '                    Else
            '                        strImg = "ճ����������;110;1;200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20,;"
            '
            '                        strImg = strImg & strŮ_��1 & "," & strŮ_��2 & "|" & strŮ_��2 & "," & strŮ_��2
            '                    End If
            '
            '                    strImg = strImg & "~" & Mid(str��, 2)
            '                    strImg = strImg & ";VIS(mPa.s),30,20~SHR(1/S),195,1" '������
            '
            '                End If
            '                If strImg <> "" Then aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & strImg
            rsTmp.MoveNext
        Loop
    End With
    'д��־
    strlog = "���һ�����ݣ�" & aResults(lngResultCount)
    Call writelislog(strlog)

    ResultFromFile = aResults
    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",writelislog,clsLISDev_File_Zl9000c"
    Close #lngErrFileNo
End Function

Public Function writelislog(strlog As String)
    On Error GoTo ErrHand '�򿪴�������
    '------------------------------------------------
    Const strFile As String = "c:\lislog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFileName As String
    strFileName = strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strlog)
    objStream.Close
    Set objStream = Nothing

    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",writelislog,clsLISDev_File_Zl9000c"
    Close #lngErrFileNo
End Function