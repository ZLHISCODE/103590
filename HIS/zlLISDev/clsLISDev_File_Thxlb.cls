VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Thxlb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�캣Ѫ������
'����:�����캣
'����:�¶�
'�û�:��ͩ
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
'   Ԫ��֮����|�ָ�
'   ��0��Ԫ�أ�����ʱ��
'   ��1��Ԫ�أ��������
'   ��2��Ԫ�أ�������
'   ��3��Ԫ�أ��걾
'   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
'   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    
    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
    Dim strBh As String, strAnti As String, str��� As String
    Dim lngResultCount As Long
    Dim str���� As String, arr����() As String
    Dim curNd0 As Currency, curNd1 As Currency, curNd2 As Currency, curNd3 As Currency
    Dim cur���� As Currency, cur���� As Currency, cur���� As Currency, str���� As String, strImg As String
    Dim str��ע As String
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    '�����ļ��õ��������̵������ļ����ڼ�ʦվ�Ĳ���ѡ����ָ���ļ�ʱ��ָ���������ļ��������ļ�������Ҫ�õ����������£�
    '[BDE_MSQL_XLB]
    'Database Name = th_xlb2035
    'SERVER Name = PC081128
    'USER Name = sa
    'PASSWORD=
    '--------------------------------------------------------------------------------------------------------------------
    '�������ļ�
    Call ReadConf(strFile, "SERVER NAME,DATABASE NAME,USER NAME,PASSWORD", str����)
    
    arr���� = Split(str����, ",")
    If UBound(arr����) <> 3 Then Exit Function
    With cnAccess
        .ConnectionString = _
        "Provider=SQLOLEDB.1;Initial Catalog=" & arr����(1) & ";Password=" & arr����(3) & _
        ";Persist Security Info=True;User ID=" & arr����(2) & ";Data Source=" & arr����(0)
        .Open
        
        If .State = adStateClosed Then Exit Function
        strSQL = "SELECT A.SXH as �걾��, A.SEX as �Ա�,B.SXH as ͨ����, B.VAL as ���, A.RQ as ����,A.BZ as ��ע,B.V_Max,B.V_Min " & _
                 " FROM JZXX as A INNER JOIN JZ_ITEM as B ON A.ID = B.ID AND A.RQ >= '" & Format(dtStart, "yyyy-MM-dd") & "' And A.RQ <= '" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                  IIf(lngSampleNO = -1, "", " And A.SXH =" & lngSampleNO) & _
                  " order by A.RQ,A.SXH,B.SXH"
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        lngResultCount = 0
        strBh = "***"
        
        'rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If Val("" & rsTmp("�걾��")) > 0 Then '�±걾
                

                
                If strBh <> "" & rsTmp("�걾��") And strBh <> "***" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str����
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(strBh) '�������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|Ѫ��"
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & str��� & IIf(str��ע <> "", "|BZ|" & str��ע, "")
                    
                    
                    strImg = vbCrLf & "ճ����������;2;20,220|30-30,50-50,100-100,150-150,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;"
                    
                    'ND0 = 9.58: QB0 = 10: ND1 = 4.6: QB1 = 150
                    'ND2 = 6.8: QB2 = 10: ND3 = 3.73: QB3 = 150
                    
                    strImg = strImg & curNd0 & ",30," & curNd1 & ",200|" & curNd2 & ",30," & curNd3 & " ,200"
                    
                    strImg = strImg & "~30-" & cur���� & ",100-" & cur���� & ",200-" & cur����  '���
                    strImg = strImg & ";VIS(mPa.s),35,20~SHR(1/S),200,1" '������
                    
                    If strImg <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & strImg
                    
                    strImg = ""
                    str��� = "": curNd2 = 0: curNd0 = 0: curNd1 = 0: curNd3 = 0
                    cur���� = 0: cur���� = 0: cur���� = 0
                End If
                strBh = Val("" & rsTmp("�걾��"))
                str��� = str��� & "|" & rsTmp("ͨ����") & "|" & rsTmp("���")
                str���� = Format("" & rsTmp("����"), "yyyy-mm-dd hh:mm:ss") '����ʱ��
                str��ע = Trim("" & rsTmp("��ע"))
                'ͼ�ι�ʽ
                'b = ln(����ֵ / ����ֵ) / ln(200 / 3)
                'a = ����ֵ/power(3,b)     //����power(3,b)Ϊ��3��b����
                '
                '//a,bΪ������ϵ��
                '
                '���� = a * power(ʱ��ֵ, b)
                '=========================================================
                'ln: ��EΪ�׵���Ȼ����
                'jz_item :��ȡ����,����,���е�ֵ,�����������.
                '=========================================================
                
                
                If rsTmp("ͨ����") = 1 Then
                    curNd2 = Val("" & rsTmp!v_max): curNd3 = Val("" & rsTmp!v_min)
                    cur���� = Val("" & rsTmp("���"))
                End If
                If rsTmp("ͨ����") = 3 Then
                    curNd0 = Val("" & rsTmp!v_max): curNd1 = Val("" & rsTmp!v_min)
                    cur���� = Val("" & rsTmp("���"))
                End If
                If rsTmp("ͨ����") = 2 Then
                    cur���� = Val("" & rsTmp("���"))
                End If
                

            End If
            rsTmp.MoveNext
        Loop
        
        If strBh <> "" And str��� <> "" Then
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            aResults(UBound(aResults)) = str����
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(strBh) '�������
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|Ѫ��"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & str��� & IIf(str��ע <> "", "|BZ|" & str��ע, "")
            
            
            strImg = vbCrLf & "ճ����������;2;20,220|30-30,50-50,100-100,150-150,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;"
            
            'ND0 = 9.58: QB0 = 10: ND1 = 4.6: QB1 = 150
            'ND2 = 6.8: QB2 = 10: ND3 = 3.73: QB3 = 150
            
            strImg = strImg & curNd0 & ",30," & curNd1 & ",200|" & curNd2 & ",30," & curNd3 & " ,200"
            
            strImg = strImg & "~30-" & cur���� & ",100-" & cur���� & ",200-" & cur����  '���
            strImg = strImg & ";VIS(mPa.s),35,20~SHR(1/S),200,1" '������
            
            If strImg <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & strImg
            
            strImg = ""
            str��� = "": curNd2 = 0: curNd0 = 0: curNd1 = 0: curNd3 = 0
            cur���� = 0: cur���� = 0: cur���� = 0
        End If
    End With
    
    ResultFromFile = aResults
   
End Function

Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '�������ļ�
    'strIniFile :ini�����ļ���
    'strItems   :Ҫ��ȡ���ֶΣ��磺DATABASE NAME=th_xlb�е� DATABASE NAME,����ֶ���,�ָ�
    'strReturn  :��ȡ�ķ���ֵ����:DATABASE NAME=th_xlb�е� th_xlb ���������ֵ��,�ָ�
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As TextStream

    Dim intloop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '�������ļ��ж�ȡ�������ļ��е�����cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intloop = 0 To intCount
        arrItems(intloop, 0) = Split(strItems, ",")(intloop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intloop = 0 To intCount
                    If arrLine(0) = arrItems(intloop, 0) Then
                        arrItems(intloop, 1) = arrLine(1)
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intloop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intloop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If
    
    
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub