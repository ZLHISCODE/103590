VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_GC1200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Gc1200 �����߷�����
'����:
'����:�¶�
'�û�:�Ͼ���������ҽԺ
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
    Dim conDBF As New ADODB.Connection
    Dim rsDbf As New ADODB.Recordset
    Dim aResults() As String
    Dim lngResultCount As Long
    Dim TableName As String
    Dim objFile As New Scripting.FileSystemObject
    Dim strSQL As String, strFilePath As String
    Dim str���� As String, str�걾�� As String, str������Ŀ As String, str���Խ�� As String
    
    ReDim aResults(0)
    
    If Trim(Dir(strFile)) = "" Then Exit Function
    
    strFilePath = objFile.GetParentFolderName(strFile)
    
    TableName = UCase(Replace(strFile, strFilePath & "\", ""))
    If conDBF.State = 1 Then conDBF.Close
    conDBF.ConnectionString = "Provider=microsoft.jet.oledb.4.0;data source=" & strFilePath & ";" & _
                          "Extended Properties=dBase 5.0"
    conDBF.Open
    strSQL = "Select * From " & TableName & " Where mid(BRXMCLBM,1,8) >=  '" & Format(dtStart, "yyyyMMdd") & "' And " & _
             "  mid(BRXMCLBM,1,8) <= '" & Format(dtEnd, "yyyyMMdd") & "' " & _
              IIf(lngSampleNO = -1, "", " And  BRID=" & lngSampleNO)

    Set rsDbf = conDBF.Execute(strSQL)
    lngResultCount = -1
    Do Until rsDbf.EOF
        If Trim("" & rsDbf!result) <> "" Then
            str���� = Mid(rsDbf!BRXMCLBM, 1, 8)
            str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7)
            str������Ŀ = Trim("" & rsDbf!xmmc)
            str���Խ�� = Trim("" & rsDbf!result)
            str�걾�� = Trim("" & rsDbf!BRID)
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            If IsDate(CDate(str����)) And Val(str�걾��) > 0 And str������Ŀ <> "" Then
                aResults(lngResultCount) = str���� & "|" & Val(str�걾��) & "| |ѪҺ|0|" & str������Ŀ & "|" & str���Խ��
            End If
        End If
        rsDbf.MoveNext
    Loop
    ResultFromFile = aResults
End Function
