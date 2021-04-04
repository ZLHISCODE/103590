VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BHP9504"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BHP9504��ѧ������
'����:�������ɹ��Ӽ����ɷ����޹�˾
'����:�µ���
'�û�:��Դ������ҽԺ

Private FileNames As String
Dim FileDate As Date
Dim fso As New FileSystemObject
Dim fld As Folder

Private Function Trunc(ByVal aDate As Date) As Date
  Dim Tmp As String
  Tmp = Format(aDate, "YYYY-MM-DD")
  Trunc = CDate(Tmp)
End Function

Private Function FindFile(ByVal sFol As String, ByVal sFile As String, ByVal aDate As Date, ByVal bDate As Date) As Currency
   Dim tFld As Folder, tFil As File, Filename As String
    
   On Error GoTo Catch
   Set fld = fso.GetFolder(sFol)
   Filename = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or _
                  vbHidden Or vbSystem Or vbReadOnly)
   While Len(Filename) <> 0
      FindFile = FindFile + FileLen(fso.BuildPath(fld.Path, _
      Filename))
      Set tFil = fso.GetFile(fld.Path & "\" & Filename)
      FileDate = tFil.DateLastModified
      If Trunc(FileDate) >= Trunc(aDate) And Trunc(FileDate) <= Trunc(bDate) Then
        If FileNames <> "" Then FileNames = FileNames + "|"
        FileNames = FileNames & fso.BuildPath("", Filename)
      End If
      Filename = Dir() ' Get next file
      DoEvents
   Wend
   If fld.SubFolders.count > 0 Then
      For Each tFld In fld.SubFolders
         DoEvents
         FindFile = FindFile + FindFile(tFld.Path, sFile, aDate, bDate)
      Next
   End If
   Exit Function
Catch: Filename = ""
       Resume Next
End Function


Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strChenel As String, str����Ա As String, str���� As String
    Dim strFilePath As String, strLog As String
    
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim str_Return As String, str_Tmp As String
    Dim lngRow As Integer, varLine As Variant
    Dim ExportFile As String
    Dim aFiles() As String, intloop As Integer, UseMode As Integer
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream2 As Scripting.TextStream
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strIniFile = App.Path & "\BHP9504.ini"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream2 = objFileSystem.CreateTextFile(strIniFile)
       objStream2.WriteLine "[Options]"
       objStream2.WriteLine "ReadMode=0 ;0=��ȡ�����ļ� 1=��ȡѡ�����ڷ�Χ�ڵ��ļ�"
       objStream2.Close
       Set objStream2 = Nothing
    End If
    UseMode = Val(ReadIni("Options", "ReadMode", strIniFile))
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    If UseMode = 0 Then
       FileNames = Format(Now, "YYYY-MM-DD") & ".txt"
    ElseIf UseMode = 1 Then
       Call FindFile(strFilePath, "*.txt", dtStart, dtEnd)
       If FileNames = "" Then
          FileNames = Format(dtEnd, "YYYY-MM-DD") & ".txt"
       End If
    End If
    If Right(strFilePath, 1) <> "\" Then strFilePath = strFilePath & "\"
    aFiles = Split(FileNames, "|")
    For intloop = LBound(aFiles) To UBound(aFiles)
      Set objStream = objFile.OpenTextFile(strFilePath & aFiles(intloop), ForReading)
      
      Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        varLine = Split(strLine, Chr(&HD) & Chr(&HA))
        
        For lngRow = LBound(varLine) To UBound(varLine)
           str_Return = varLine(lngRow)
           str_Tmp = Trim(Mid(str_Return, 1, 12))
           
           If IsNumeric(str_Tmp) Then
              str���� = Trim(Mid(str_Return, 14, 20))
              str�걾�� = Trim(Mid(str_Tmp, 9))
              strChenel = Trim(Mid(str_Return, 38, 12))
              str������ = Trim(Mid(str_Return, 52))
              If Trim(strChenel) <> "" And Trim(str������) <> "" Then
                 If aResults(UBound(aResults)) <> "" Then
                  ReDim Preserve aResults(UBound(aResults) + 1)
                 End If
              End If
              aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|0|" & strChenel & "|" & str������
              strLog = strLog & str���� & "|" & str�걾�� & "| |Ѫ��|0|" & strChenel & "|" & str������
           End If
        Next
        DoEvents
      Loop
    Next
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_BHP9504.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(Date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function
