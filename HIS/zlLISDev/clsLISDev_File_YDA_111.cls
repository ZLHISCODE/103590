VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_YDA_111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:YDA-111 Ѫ����
'����:
'����:���
'�û�:������˾-��ΰ(˼��������ҽԺ)
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
    Dim i As Integer
    Dim aResults() As String
    Dim frmPic As New frmGraph
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsAnti As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim str��� As String, str�걾�� As String
    Dim lngResultCount As Long
    Dim str���� As String, arr����() As String
    Dim str���� As String, iField As Integer
    Dim aPic() As String
    Dim strPic As String
    Dim strPicPath As String               'ͼ�����·��
    Dim arrHigh(2) As Double, arrVAL(2) As Double, arrLow(2) As Double
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    '�����ļ��õ��������̵������ļ����ڼ�ʦվ�Ĳ���ѡ����ָ���ļ�ʱ��ָ���������ļ��������ļ�������Ҫ�õ����������£�
    '[BDE_MSQL_XLB]
    'Database Name = th_xlb2035
    'SERVER Name = PC081128
    'USER Name = sa
    'PASSWORD=
    '--------------------------------------------------------------------------------------------------------------------
'
    strPicPath = ReadIni("YDA-111", "ͼƬ·��", App.Path & "\YDA-111.ini")
    If strPicPath = "" Then strPicPath = "E:": Call WriteIni("YDA-111", "ͼƬ·��", strPicPath, App.Path & "\YDA-111.ini")
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
        strSQL = "SELECT  * from persons" & _
                " where check_date>='" & Format(dtStart, "yyyy-MM-dd") & "'" & _
                " And check_date<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                IIf(lngSampleNO = -1, "", " And person_ID=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If

        ReDim aResults(0)
        lngResultCount = 0

        Do Until rsTmp.EOF
            If Val("" & rsTmp("person_ID")) > 0 Then '�±걾
                str�걾�� = "" & rsTmp("person_ID")
                str���� = Format(rsTmp("check_Date"), "YYYY-MM-DD")        '����ʱ��
                str��� = ""
                
                For i = 0 To 2
                    arrHigh(i) = rsTmp("id" & i + 2 & "_high")
                    arrVAL(i) = rsTmp("id" & i + 2 & "_val")
                    arrLow(i) = rsTmp("id" & i + 2 & "_low")
                Next
                Call GetResult(arrHigh()): Call GetResult(arrVAL()): Call GetResult(arrLow())
                
                For iField = 12 To 41
                    If "" & rsTmp.Fields(iField).Value <> "" Then
                        str��� = str��� & "|" & rsTmp.Fields(iField).Name & "|" & rsTmp.Fields(iField + 30).Value
                    End If
                Next
            
                If str�걾�� <> "" And str��� <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str���� & "|" & Val(str�걾��) & "| |Ѫ��|0" & str���
                    strPic = "ճ����������;110;1;" & frmPic.Draw_YDA_111(arrHigh(), arrVAL(), arrLow(), strPicPath, str�걾��)
                    If strPic <> "" Then
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & strPic
                    End If
                End If
            End If
            rsTmp.MoveNext
        Loop
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

    Dim intLoop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '�������ļ��ж�ȡ�������ļ��е�����cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intLoop = 0 To intCount
        arrItems(intLoop, 0) = Split(strItems, ",")(intLoop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intLoop = 0 To intCount
                    If Trim(arrLine(0)) = Trim(arrItems(intLoop, 0)) Then
                        arrItems(intLoop, 1) = Trim(arrLine(1))
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intLoop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intLoop, 1)
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

'��˹��Ԫ������Ԫһ�η���
Public Sub GetResult(arrPara() As Double)
    Dim arrA(2, 3) As Double
    Dim intRows As Integer: intRows = 3
    Dim intL As Integer: intL = intRows - 1
    Dim i As Integer, j As Integer, l As Integer, n As Integer, m As Integer, k As Integer
    Dim arrTemp(2) As Double
    
    arrA(0, 0) = 1 / 120 ^ 2: arrA(0, 1) = 1 / 120: arrA(0, 2) = 1: arrA(0, 3) = arrPara(0)
    arrA(1, 0) = 1 / 70 ^ 2: arrA(1, 1) = 1 / 70: arrA(1, 2) = 1: arrA(1, 3) = arrPara(1)
    arrA(2, 0) = 1 / 30 ^ 2: arrA(2, 1) = 1 / 30: arrA(2, 2) = 1: arrA(2, 3) = arrPara(2)

    Do
        n = 0
        For l = k To intL - 1
            arrTemp(n) = arrA(l + 1, k) / arrA(k, k)
            n = n + 1
        Next
        
        For i = k + 1 To intRows - 1
            For j = k To intRows
                arrA(i, j) = arrA(i, j) - arrTemp(m) * arrA(k, j)
            Next
            m = m + 1
        Next
        m = 0
        
        k = k + 1
    Loop While (k < intRows)
    k = l - 1

    Do
        n = 0
        For l = k To 0 Step -1
            arrTemp(n) = arrA(k - l, k + 1) / arrA(k + 1, k + 1)
            n = n + 1
        Next
        
        m = 0
        For i = k To 0 Step -1
            For j = k To intRows
                arrA(k - i, j) = arrA(k - i, j) - arrTemp(m) * arrA(k + 1, j)
            Next
            m = m + 1
        Next

        k = k - 1
    Loop While (k >= 0)
  
    For i = 0 To intRows - 1
        arrPara(i) = arrA(i, intRows) / arrA(i, i)
    Next
End Sub

