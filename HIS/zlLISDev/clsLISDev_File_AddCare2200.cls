VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AddCare2200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Addcare2200
'����:��̨���¿�
'����:�µ���
'�û�:��������ҽԺ
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
    Dim strChenel As String, StrQc As String, str���� As String
    Dim strFilePath As String, strLog As String
    Dim objStream As TextStream
    Dim strLine As String, strTmp As String
    Dim str�걾�� As String, str������ As String, QcSampleNo As String
    Dim lngRow As Integer, varLine As Variant, iPos As Integer
    Dim strOD As String             'ԭʼ���(OD)
    Dim strCO As String             'OD/CUTOFF
    Dim strCutOff As String         '�ж�ֵ(Cut Off)
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject

    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strIniFile = App.Path & "\AddCare2000.ini"

    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "[Options]"
       objStream.WriteLine "QcSampleNo = 1000 ;�ʿر걾��Ϊ 1000"
       objStream.Close
       Set objStream = Nothing
    End If
    QcSampleNo = ReadIni("Options", "QcSampleNo", strIniFile)
    QcSampleNo = Replace(QcSampleNo, ";�ʿر걾��Ϊ 1000", "")
    QcSampleNo = Trim(QcSampleNo)

    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    lngRow = 0: StrQc = "": strTmp = ""
    Do Until objStream.AtEndOfStream
       strLine = objStream.ReadLine
       iPos = InStr(strLine, "������Ŀ:")
       If iPos > 0 Then
         strChenel = Trim(Mid(strLine, iPos + 5, 6))
         strChenel = Replace(strChenel, "��", "")
         strChenel = Replace(strChenel, "��", "")
       End If
       iPos = InStr(strLine, "���ʱ��:")
       If iPos > 0 Then
         str���� = Trim(Mid(strLine, iPos + 5, 20))
       End If
       If lngRow >= 5 Then
         If Trim(strChenel) <> "" And Trim(str����) <> "" Then
             If aResults(UBound(aResults)) <> "" Then
                ReDim Preserve aResults(UBound(aResults) + 1)
             End If
             str�걾�� = Trim(Mid(strLine, 1, 8))
             If (str�걾�� = "QC") Or Val(str�걾��) > 0 Then
                If str�걾�� = "QC" Then
                  StrQc = "1"
                  str�걾�� = QcSampleNo + Val(Replace(str�걾��, "QC", ""))
                Else
                  StrQc = "0"
                End If
                strTmp = Trim(Mid(strLine, 28, 8))
                strOD = Trim(Mid(strLine, 37, 8))
                strCutOff = Trim(Mid(strLine, 67, 8))
                strCO = Trim(Mid(strLine, 57, 8))
                If StrQc = "1" Then
                  str������ = strCO & "^" & strOD & "^" & strCutOff & "^" & strCO
                Else
                  str������ = strTmp & "^" & strOD & "^" & strCutOff & "^" & strCO
                End If
                
                aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|" & StrQc & "|" & IIf(StrQc = "1", "q" & strChenel, strChenel) & "|" & str������
             End If
        End If
       End If
       lngRow = lngRow + 1
       DoEvents
    Loop
 
    For lngRow = LBound(aResults) To UBound(aResults)
      Debug.Print aResults(lngRow)
    Next
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_AddCare2000.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function




