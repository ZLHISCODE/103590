VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AVE764B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�������ң������Ƽ��ɷ����޹�˾
'�ͺţ�AVE-764B����AVE-752
'���ͣ���Һ����
'�û��������е�һ����ҽԺ
'���ߣ�����
'ʱ�䣺2014/4/3 15:17:08

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo hErr '�򿪴�������
    '------------------------------------------------
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
    Dim objStream As textStream, objFileSystem As New FileSystemObject
    Dim strFilePath As String
    Dim SERVERNAME As String, DATABASENAME As String, USERNAME As String, PassWord As String, PicPath As String, imgPath As String, imgName() As String
    Dim ITYP As Integer, i As Integer, j As Integer, int_name As Integer
    Dim strSQL As String, strResults() As String, strBth As String, str_img As String
    Dim cnAccess As New ADODB.Connection, rsSick As New ADODB.Recordset, rsChem As New ADODB.Recordset, rsUri As New ADODB.Recordset
    If Len(strFile) = 0 Then strFilePath = App.Path
    strFilePath = objFileSystem.GetParentFolderName(strFile)
    If Not objFileSystem.FileExists(App.Path & "\" & "AVE764B.ini") Then '��ȡ���ݿ���Ϣ
        Set objStream = objFileSystem.CreateTextFile(App.Path & "\" & "AVE764B.ini")
            objStream.WriteLine "[AVE764B]"
            objStream.WriteLine "SERVERName = (local)"
            objStream.WriteLine "DatabaseName = AVE2"
            objStream.WriteLine "USERName = sa"
            objStream.WriteLine "PASSWORD= shuangfeng"
            objStream.WriteLine "ITYP= 0   ;0=�������쳣��������ϸ�� 1=����"
            objStream.Close
            Set objStream = Nothing
    End If
    SERVERNAME = ReadIni("AVE764B", "SERVERName", App.Path & "\" & "AVE764B.ini")
    DATABASENAME = ReadIni("AVE764B", "DatabaseName", App.Path & "\" & "AVE764B.ini")
    USERNAME = ReadIni("AVE764B", "USERName", App.Path & "\" & "AVE764B.ini")
    PassWord = ReadIni("AVE764B", "PASSWORD", App.Path & "\" & "AVE764B.ini")
    ITYP = Val(ReadIni("AVE764B", "ITYP", App.Path & "\" & "AVE764B.ini"))
    If Not objFileSystem.FileExists(App.Path & "\" & "PicAVE764B.ini") Then '��ȡͼƬ����·��
        Set objStream = objFileSystem.CreateTextFile(App.Path & "\" & "PicAVE764B.ini")
        objStream.WriteLine "[AVE764B]"
        objStream.WriteLine "PICPATH="
        objStream.Close
        Set objStream = Nothing
    End If
    PicPath = ReadIni("AVE764B", "PICPATH", App.Path & "\" & "PicAVE764B.ini")
    
    If (Trim(SERVERNAME) = "") Or (Trim(DATABASENAME) = "") Or (Trim(USERNAME) = "") Then Exit Function
    With cnAccess
        .ConnectionString = _
        "Provider=SQLOLEDB.1;Initial Catalog=" & DATABASENAME & ";Password=" & PassWord & _
        ";Persist Security Info=True;User ID=" & USERNAME & ";Data Source=" & SERVERNAME
            .Open
    End With
    If cnAccess.State = adStateClosed Then Exit Function
    ReDim Preserve strResults(0) '��ʼ������������±�Խ��
    strResults(0) = ""
    i = -1
    strSQL = "select a.SickSampleID, a.SickExamineID,a.ColorSet ,a.MuddySet,convert(varchar(100),a.CheckDate,20) as checktime from AVE2.dbo.SickInfo a " & vbNewLine   '����������Ϣ
    strSQL = strSQL & "where convert(varchar(100),a.CheckDate,23) >= " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & " and " & " convert(varchar(100),a.CheckDate,23) <= " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'" & vbNewLine
    strSQL = strSQL & IIf(lngSampleNO = -1, "", " and a.SickExamineID = " & Format(lngSampleNO, "0000"))
    Set rsSick = cnAccess.Execute(strSQL) '��ȡ��Һ������Ϣ
    Do Until rsSick.EOF
'        If strBth <> rsSick("SickSampleID").Value Then 'ÿ���±걾����һ����¼
            i = i + 1
            ReDim Preserve strResults(i)
            strResults(i) = Format(rsSick("checktime").Value, "yyyy-mm-dd hh:mm:ss") & "|" & Format(rsSick("SickExamineID").Value, "0000") & "| |��Һ|0" & "|Color|" & " " & rsSick("ColorSet").Value & "|Muddy|" & " " & rsSick("MuddySet").Value
'        End If
        rsSick.MoveNext
    Loop
    rsSick.Close
    strSQL = "select b.ChemSampleID,b.ChemExamineID,convert(varchar(100),b.ChemDate,23) as chemtime,b.GLU,b.BLD,b.LEU,b.PRO,b.NIT,b.URO,b.BIL,b.KET,b.PH,b.SG,b.VC from AVE2.dbo.ChemResult b " & vbNewLine
    strSQL = strSQL & "where convert(varchar(100),b.ChemDate,23) >= " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & " and " & " convert(varchar(100),b.ChemDate,23) <= " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'" & vbNewLine
    strSQL = strSQL & IIf(lngSampleNO = -1, "", " and b.ChemExamineID = " & Format(lngSampleNO, "0000"))
    Set rsChem = cnAccess.Execute(strSQL) '��ȡ�ɻ�ѧ��������Ϣ
    Do Until rsChem.EOF
        i = i + 1
        ReDim Preserve strResults(i)
        strResults(i) = Format(rsChem("chemtime").Value, "yyyy-mm-dd") & " " & Format(Now, "hh:mm:ss") & "|" & Format(rsChem("ChemExamineID").Value, "0000") & "| |��Һ|0"
        For j = 3 To rsChem.Fields.count - 1
            strResults(i) = strResults(i) & "|" & rsChem.Fields.Item(j).Name & "|" & " " & rsChem(rsChem.Fields.Item(j).Name).Value
        Next
        rsChem.MoveNext
    Loop
    rsChem.Close
    strSQL = "select c.UriSampleID,c.UriExamineID,convert(varchar(100),c.UriDate,23) as uritime,c.Rbc,c.Wbc,c.Crystal,c.Epiphyte,c.Duct,c.Epithelia,c.WhackRbc,c.SingularRbc,c.Infusorial," & _
                 "c.Mucous,c.Pyoid,c.NCrystal,c.LCrystal,c.HCrystalRef,c.CCrystal,c.WCrystal,c.RbcDuct,c.WbcDuct,c.KDuct,c.TDuct,c.LDuct,c.LEpithelia," & _
                 "c.XEpithelia,c.JEpithelia,c.BigRbc,c.SmallRbc,c.JRbc,c.ZRbc,c.QRbc,c.Neutrophil,c.Lymphocyte,c.Acidophil,c.Basophil,c.Monocyte," & _
                 "c.Bacterial , c.Cocci, c.Bacilli, c.Mildew, c.Yeast, c.Parasites, c.Sperm, c.PrintPath, c.PrintImage " & _
                 "from AVE2.dbo.UriResult c " & vbNewLine
    strSQL = strSQL & "where convert(varchar(100),c.UriDate,23) >= " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & " and " & " convert(varchar(100),c.UriDate,23) <= " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'" & vbNewLine
    strSQL = strSQL & IIf(lngSampleNO = -1, "", " and c.UriExamineID = " & Format(lngSampleNO, "0000"))
    Set rsUri = cnAccess.Execute(strSQL) '��ȡ��������Ϣ
    Do Until rsUri.EOF
        i = i + 1
        ReDim Preserve strResults(i)
        strResults(i) = Format(rsUri("uritime").Value, "yyyy-mm-dd") & " " & Format(Now, "hh:mm:ss") & "|" & Format(rsUri("UriExamineID").Value, "0000") & "| |��Һ|0"
        For j = 3 To rsUri.Fields.count - 1
            If j >= 3 And j <= rsUri.Fields.count - 3 Then 'һ���������
                strResults(i) = strResults(i) & "|" & rsUri.Fields.Item(j).Name & "|" & " " & rsUri(rsUri.Fields.Item(j).Name).Value
            End If
            If j = rsUri.Fields.count - 2 Then '��ȡͼ�񱣴��·��
                imgPath = rsUri(rsUri.Fields.Item(j).Name).Value
            End If
            If j = rsUri.Fields.count - 1 Then '��ȡͼ�������
                imgName = Split(rsUri(rsUri.Fields.Item(j).Name).Value, ",")
                For int_name = LBound(imgName) To UBound(imgName)
                    If imgName(int_name) <> "" Then
                        If imgName(int_name) = "H0001" Then
                            str_img = str_img & "^" & imgName(int_name) & ";113;0;" & imgPath & imgName(int_name) & ".jpg"
                            str_img = str_img & "^" & "H0002" & ";113;0;" & imgPath & "H0002" & ".jpg"
                            str_img = str_img & "^" & "MorphA" & ";113;0;" & imgPath & "MorphA" & ".jpg"
                            str_img = str_img & "^" & "MorphC" & ";113;0;" & imgPath & "MorphC" & ".jpg"
                            str_img = str_img & "^" & "MorphS" & ";113;0;" & imgPath & "MorphS" & ".jpg"
                            str_img = str_img & "^" & "MorphT" & ";113;0;" & imgPath & "MorphT" & ".jpg"
                        End If
                    End If
                Next
                If str_img = "" Then
                    str_img = str_img & "^" & "L0001" & ";113;0;" & imgPath & "L0001" & ".jpg"
                    str_img = str_img & "^" & "L0002" & ";113;0;" & imgPath & "L0002" & ".jpg"
                    str_img = str_img & "^" & "L0003" & ";113;0;" & imgPath & "L0003" & ".jpg"
                    str_img = str_img & "^" & "L0004" & ";113;0;" & imgPath & "L0004" & ".jpg"
                    str_img = str_img & "^" & "L0005" & ";113;0;" & imgPath & "L0005" & ".jpg"
                    str_img = str_img & "^" & "L0006" & ";113;0;" & imgPath & "L0006" & ".jpg"
                End If
                If str_img <> "" Then strResults(i) = strResults(i) & vbCrLf & Mid(str_img, 2)
                str_img = ""
            End If
        Next
        rsUri.MoveNext
    Loop
    rsUri.Close
    If UBound(strResults) = 0 Then Exit Function
   
    cnAccess.Close
    ResultFromFile = strResults
    '------------------------------------------------
    Exit Function
    '----------------
hErr:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrMsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_File_AVE764B"
    Close #lngErrFileNo
End Function