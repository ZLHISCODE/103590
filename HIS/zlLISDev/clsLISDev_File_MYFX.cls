VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_myfx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
    'Dim conDBF As New ADODB.Connection
    Dim rsDbf As New ADODB.Recordset
    Dim cnAccess As New ADODB.Connection
    Dim aResults() As String
    Dim lngResultCount As Long
    Dim TableName As String
    Dim objFile As New Scripting.FileSystemObject
    Dim strSQL As String, strFilePath As String
    Dim str���� As String, str�걾�� As String, str������Ŀ As String, str���Խ�� As String

    ReDim aResults(0)
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .CursorLocation = adUseClient
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        
        If .State = adStateClosed Then Exit Function
        
 
    strSQL = "Select * From testtable Where Format(datetested,'yyyy-mm-dd') Between " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                     " and " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'"

    Set rsDbf = cnAccess.Execute(strSQL)
        
    If rsDbf.EOF Then
        .Close
        Exit Function
    End If
        
    lngResultCount = -1
    Do Until rsDbf.EOF
            str���� = rsDbf!datetested
            str�걾�� = Trim(rsDbf!samplecode)
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            If IsDate(CDate(str����)) And Val(str�걾��) > 0 Then
          
           aResults(lngResultCount) = str���� & "|" & Val(str�걾��) & "| |ѪҺ|0|" & _
                  rsDbf!Name1 & "|" & rsDbf!d1 & "|" & rsDbf!Name2 & "|" & rsDbf!d2 & "|" & rsDbf!Name3 & "|" & rsDbf!d3 & "|" & _
                  rsDbf!Name4 & "|" & rsDbf!d4 & "|" & rsDbf!Name5 & "|" & rsDbf!d5 & "|" & rsDbf!Name6 & "|" & rsDbf!d6 & "|" & _
                  rsDbf!Name7 & "|" & rsDbf!d7 & "|" & rsDbf!Name8 & "|" & rsDbf!d8 & "|" & rsDbf!Name9 & "|" & rsDbf!d9 & "|" & _
                  rsDbf!Name10 & "|" & rsDbf!d10 & "|" & rsDbf!Name11 & "|" & rsDbf!d11 & "|" & rsDbf!Name12 & "|" & rsDbf!d12 & "|" & _
                  rsDbf!Name13 & "|" & rsDbf!d13
                  
            End If
        rsDbf.MoveNext
    Loop
    End With
    ResultFromFile = aResults
End Function


