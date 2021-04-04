VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ZL6000C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Ѫ����
'�����ͺţ�ZL6000C
'�û���̫ԭ���õ��ҽԺ
'��ս�У�������
'���ڣ�2013-03-22
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
    Dim lngResultCount As Long, strBH As String, strLog As String
    Dim testid As Long
    Dim blnUnion As Boolean
    Dim lngCount As Long, str�� As String, varTmp As Variant
    Dim str��_��1 As String, str��_��2 As String, strŮ_��1 As String, strŮ_��2 As String
    Dim str��_��1 As String, str��_��2 As String, strŮ_��1 As String, strŮ_��2 As String
    Dim strImg As String
    Dim lngRecord As String
    'д��־
    strLog = "�Ѿ����ò���"
    Call writelislog(strLog)
    ReDim Preserve aResults(0)
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        'д��־
        strLog = "�Ѿ����ļ�"
        Call writelislog(strLog)
        
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
                      " '01|' +  cstr(iif(isnull(Ѫ��ճ��),' ',Ѫ��ճ��)) + '|02|' + cstr(iif(isnull(ѹ��),' ',ѹ��)) + '|03|' + cstr(iif(isnull(Ѫ��),' ',Ѫ��)) + '|04|' + cstr(iif(isnull(ȫѪ�������ָ��),' ',ȫѪ�������ָ��)) + '|05|' + cstr(iif(isnull(ȫѪ�������ָ��),' ',ȫѪ�������ָ��))" & _
                      " + '|06|' + cstr(iif(isnull(Ѫ������Kֵ),' ',Ѫ������Kֵ)) + '|07|' + cstr(iif(isnull(��ϸ���ۼ�ָ��),' ',��ϸ���ۼ�ָ��)) + '|08|' + cstr(iif(isnull(��ϸ���ۼ�ϵ��),' ',��ϸ���ۼ�ϵ��))" & _
                      " + '|09|' + cstr(iif(isnull(ȫѪ���л�ԭճ��),' ',ȫѪ���л�ԭճ��)) + '|10|' + cstr(iif(isnull(ȫѪ���л�ԭճ��),' ',ȫѪ���л�ԭճ��)) + '|11|' + cstr(iif(isnull(��ϸ������ָ��TK),' ',��ϸ������ָ��TK))" & _
                      " + '|12|' + cstr(iif(isnull(��ϸ������ָ��),' ',��ϸ������ָ��)) + '|13|' + cstr(iif(isnull(����ճ��),' ',����ճ��))as ���1," & _
                      "'|14|' + cstr(iif(isnull(Ѫ�쵰��),' ',Ѫ�쵰��)) + '|15|' + cstr(iif(isnull(��ϸ����ճ��),' ',��ϸ����ճ��))" & _
                      " + '|16|' + cstr(iif(isnull(��������),' ',��������)) + '|17|' + cstr(iif(isnull(��������),' ',��������)) + '|18|' + cstr(iif(isnull(��������),' ',��������)) + '|19|' + cstr(iif(isnull(��ά����ԭ),' ',��ά����ԭ)) + '|20|' + cstr(iif(isnull(ȫѪ���л�ԭճ��),' ',ȫѪ���л�ԭճ��))" & _
                      " + '|21|' + cstr(iif(isnull(����Ӧ��),' ',����Ӧ��)) + '|22|' + cstr(iif(isnull(��ϸ����Ӿָ��),' ',��ϸ����Ӿָ��)) + '|23|' + cstr(iif(isnull(ȫѪ�������ָ��),' ',ȫѪ�������ָ��)) + '|24|' + cstr(iif(isnull(��ϸ������),' ',��ϸ������))as ���2 " & _
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
        strLog = "û���ҵ�����"
        Call writelislog(strLog)
            ResultFromFile = aResults
            Exit Function
        End If
        'д��־
        strLog = "yp:�Ѿ��ҵ�����"
        Call writelislog(strLog)

        lngResultCount = -1
        strBH = "***"
        Do While Not rsTmp.EOF
            If strBH <> rsTmp("���") Then '�±걾
                
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

                strBH = rsTmp("���")

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
                     varTmp = Split(str��, ",")
                     If UBound(varTmp) < 4 Then
                        str�� = varTmp(0) & "," & varTmp(1) & "," & varTmp(2) & "," & varTmp(3) & "," & varTmp(3)
                     
                     End If
                    '���ݹ�ʽ��ͼ
                    If Trim("" & rsTmp!�Ա�) = True Then    '��
                        strImg = "ճ����������;110;1;" & frmGraph.Draw_ZL6000C(str��_��1 & ";" & str��_��2, str��, str��_��1 & ";" & str��_��2, App.Path, strBH)
                    Else
                        strImg = "ճ����������;110;1;" & frmGraph.Draw_ZL6000C(strŮ_��1 & ";" & strŮ_��2, str��, strŮ_��1 & ";" & strŮ_��2, App.Path, strBH)
                    End If
                End If
                If strImg <> "" Then aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & strImg
            rsTmp.MoveNext
        Loop
    End With
   'д��־
    strLog = "���һ�����ݣ�" & aResults(lngResultCount)
    Call writelislog(strLog)
    
    ResultFromFile = aResults
End Function

Public Function writelislog(strLog As String)
    Const strFile As String = "c:\lislog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing

End Function





