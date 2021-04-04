VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_FILE_TMS2024I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:������TMS2024
'����:
'����:����
'�û�:Ϭ�ֹ�·��ҽԺ
'ʱ��:2014/11/17

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo errHand '�򿪴�������
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
    Dim OBJFILE As New FileSystemObject, objStream As textStream
    Dim cnAccess As New ADODB.Connection, rsTmp As New ADODB.Recordset
    Dim strLine As String, strFilePath As String, strSQL As String
    Dim strBth As String, intLoop As Integer, j As Integer
    Dim rsResults() As String
    
    ReDim Preserve rsResults(0)
    rsResults(0) = ""
    If Len(strFile) = 0 Then Exit Function
    strFilePath = OBJFILE.GetParentFolderName(strFile)
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
    End With
    If cnAccess.State = adStateClosed Then Exit Function
    strSQL = "select B.SAMP_NO,A.SAMP_ID,A.ITEM_NAME,A.CONC_DATA,A.RST_DATE,A.FDATE from SAMP_DAT A,SAMP_ODR B where A.SAMP_ID =B.SAMP_ID "
    strSQL = strSQL & "and format(A.RST_DATE,'yyyy-mm-dd hh:mm:ss') >= format('" & dtStart & "','yyyy-mm-dd hh:mm:ss') "
    strSQL = strSQL & "and format(A.RST_DATE,'yyyy-mm-dd hh:mm:ss') <= format('" & dtEnd & "','yyyy-mm-dd hh:mm:ss') "
    strSQL = strSQL & IIf(lngSampleNO = -1, "", "and B.SAMP_NO=" & lngSampleNO)
    strSQL = strSQL & "order by A.FDATE,B.SAMP_NO"
    Set rsTmp = cnAccess.Execute(strSQL)
    strBth = "***"
    intLoop = -1
    Do Until rsTmp.EOF
        If strBth <> rsTmp("SAMP_NO").Value Then '�µı걾
            intLoop = intLoop + 1
            ReDim Preserve rsResults(intLoop)
            strBth = rsTmp("SAMP_NO").Value
            If InStr(rsTmp("SAMP_ID").Value, "C") > 0 Then
                rsResults(intLoop) = Format(rsTmp("RST_DATE").Value, "yyyy-MM-dd hh:mm:ss") & "|" & Val(rsTmp("SAMP_NO").Value) & "|" & " |" & "Ѫ��|1|" & rsTmp("ITEM_NAME").Value & "|" & Format(Val(rsTmp("CONC_DATA").Value), "0.00")
            ElseIf InStr(rsTmp("SAMP_ID").Value, "E") > 0 Then
                rsResults(intLoop) = Format(rsTmp("RST_DATE").Value, "yyyy-MM-dd hh:mm:ss") & "|" & Val(rsTmp("SAMP_NO").Value) & "^1^" & "|" & " |" & "Ѫ��|0|" & rsTmp("ITEM_NAME").Value & "|" & Format(Val(rsTmp("CONC_DATA").Value), "0.00")
            Else
                rsResults(intLoop) = Format(rsTmp("RST_DATE").Value, "yyyy-MM-dd hh:mm:ss") & "|" & Val(rsTmp("SAMP_NO").Value) & "|" & " |" & "Ѫ��|0|" & rsTmp("ITEM_NAME").Value & "|" & Format(Val(rsTmp("CONC_DATA").Value), "0.00")
            End If
        Else
            rsResults(intLoop) = rsResults(intLoop) & "|" & rsTmp("ITEM_NAME").Value & "|" & Format(Val(rsTmp("CONC_DATA").Value), "0.00")
        End If
        rsTmp.MoveNext
    Loop
    If UBound(rsResults) > 0 Then
        For j = 0 To UBound(rsResults)
            Call writelislog(rsResults(j))
        Next
    End If
    ResultFromFile = rsResults
    rsTmp.Close
    cnAccess.Close
    '------------------------------------------------
    Exit Function
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_LIAISON"
    Close #lngErrFileNo
End Function
Public Function writelislog(strLog As String)
    Const strFile As String = "LisLog_TMS2024i_"
    Dim OBJFILE As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim STRFILENAME As String
    STRFILENAME = App.Path & "\" & strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not OBJFILE.FileExists(STRFILENAME) Then Call OBJFILE.CreateTextFile(STRFILENAME)
    Set objStream = OBJFILE.OpenTextFile(STRFILENAME, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function