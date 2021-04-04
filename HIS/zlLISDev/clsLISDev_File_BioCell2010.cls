VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BioCell2010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����2010ø����
'����:֣�ݲ������﹤���������ι�˾
'����:�µ���
'�û�:��������ҽԺ

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
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strChenel As String, str����Ա As String, str���� As String
    Dim strFilePath As String, strLog As String
    
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim ItemName As Integer, ItemResult As Integer, ItemOD As Integer, ItemSCO As Integer '��INI��ȡͨ���ź���Ŀ�����λ��
    Dim lngRow As Long, varline As Variant
    Dim ExportFile As String, OD As String, SCO As String
    Dim aFiles() As String, intloop As Integer, UseMode As Integer, ResultMode As Integer, CheckDate As Integer
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject
    
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strIniFile = App.Path & "\BioCell2010.ini"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "[Options]"
       objStream.WriteLine "ItemName=12       ;��Ŀ������������"
       objStream.WriteLine "ItemResult=15     ;��Ŀֵ��������"
       objStream.WriteLine "ItemOD=14         ;ODֵ��������"
       objStream.WriteLine "ItemSCO=16        ;SCOֵ��������"
       objStream.WriteLine "CheckDate=0       ;����������������"
       objStream.WriteLine "ReadMode=1        ;��ȡ�ļ���ʽ 0=ֻ��ȡ�����ļ� 1=��ȡָ�����ڷ�Χ�ļ�"
       objStream.WriteLine "ResultMode=0      ;���������ʽ 0=������ 1=���⴦��"
       objStream.Close
       Set objStream = Nothing
    End If
    
    ItemName = Val(ReadIni("Options", "ItemName", strIniFile))
    ItemResult = Val(ReadIni("Options", "ItemResult", strIniFile))
    CheckDate = Val(ReadIni("Options", "CheckDate", strIniFile))
    UseMode = Val(ReadIni("Options", "ReadMode", strIniFile))
    ResultMode = Val(ReadIni("Options", "ResultMode", strIniFile))
    If ItemName = 0 Then ItemName = 12
    If ItemResult = 0 Then ItemResult = 15
    If ItemOD = 0 Then ItemOD = 14
    If ItemSCO = 0 Then ItemSCO = 16
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    If UseMode = 0 Then
       FileNames = Format(Now, "YYMMDD") & "dat.txt"
    ElseIf UseMode = 1 Then
       Call FindFile(strFilePath, "*.txt", dtStart, dtEnd)
       If FileNames = "" Then
          FileNames = Format(dtEnd, "YYMMDD") & "dat.txt"
       End If
    End If
    If Right(strFilePath, 1) <> "\" Then strFilePath = strFilePath & "\"
    aFiles = Split(FileNames, "|")
    For intloop = LBound(aFiles) To UBound(aFiles)
      Set objStream = objFile.OpenTextFile(strFilePath & aFiles(intloop), ForReading)
      
      Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        varline = Split(strLine, Chr(&H9))
        If UBound(varline) >= ItemResult Then
           If IsNumeric(varline(0)) Then
              If (Len(varline(0)) = 10) Or (Len(varline(0)) = 11) Or (Len(varline(0)) = 9) Then
                 If CheckDate = 0 Then
                     str���� = "20" & Mid(varline(0), 1, 2) & "-" & Mid(varline(0), 3, 2) & "-" & Mid(varline(0), 5, 2) & " " & Format(Now, "HH:MM:SS")
                 Else
                     str���� = Trim(varline(CheckDate))
                 End If
                 str�걾�� = Mid(varline(0), 7)
                 If Trim(varline(ItemName)) <> "" And Trim(varline(ItemResult)) <> "" Then
                    If aResults(UBound(aResults)) <> "" Then
                      ReDim Preserve aResults(UBound(aResults) + 1)
                    End If
                    strChenel = Trim(varline(ItemName))
                    strChenel = Trim(Replace(strChenel, "÷������", "TP"))
                    str������ = Trim(varline(ItemResult))
                    OD = Format(Trim(varline(ItemOD)), "#0.000#")
                    SCO = Format(Trim(varline(ItemSCO)), "#0.000#")
'                    If (ResultMode = 1) Then
'                       If (strChenel = "HBsAg") And (CCur(str������) < 0.1) Then
'                          str������ = "<0.10"
'                       ElseIf (strChenel = "Anti-HBs") And (CCur(str������) < 5) Then
'                          str������ = "<5.00"
'                       ElseIf (strChenel = "HbeAg") And (CCur(str������) < 0.25) Then
'                          str������ = "<0.25"
'                       ElseIf (strChenel = "Anti-HBe") And (CCur(str������) < 1) Then
'                          str������ = "<1.00"
'                       ElseIf (strChenel = "Anti-HBc") And (CCur(str������) < 2) Then
'                          str������ = "<2.00"
'                       End If
'                    End If
                    aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|0|" & strChenel & "|" & str������ & "^" & OD & "^0.00^" & SCO
                End If
              End If
           End If
        End If
        DoEvents
      Loop
    Next
    For lngRow = LBound(aResults) To UBound(aResults)
      Debug.Print aResults(lngRow)
    Next
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_BioCell2010.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function

