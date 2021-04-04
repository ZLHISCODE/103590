VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ARIS2X_Swin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:
'����:
'����:����
'�û�:��������ҵ��ҽԺ
Private Declare Function CharUpper Lib "user32.dll" Alias "CharUpperW" (ByVal lpsz As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

Private Function ReadUniFile(ByVal sFile As String) As String '����unicode���������
    Dim a As Long
    a = FileLen(sFile)
    ReDim buff(a - 1) As Byte
    ReDim buff1(a - 3) As Byte
    Open sFile For Binary As #1
    Get #1, , buff
    Close #1
    CopyMemory buff1(0), buff(2), a - 2
    Dim s As String
    s = StrConv(buff1, vbNarrow)
    ReadUniFile = s
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
    Dim objFile As New FileSystemObject, strStream As textStream
    Dim strFileName As String, strFilePath As String
    Dim intDay As Integer, intLoop As Integer, i As Integer, j As Integer
    Dim strResult() As String, strDate As String, strItem As String, strGerm As String, strAnti As String, strValue As String
    Dim strLine As String, arrLine() As String, strSample As String, strFnc As String
    
    On Error GoTo ErrorHand
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    ReDim Preserve strResult(0)
    strResult(0) = ""
    j = 2
    intDay = Abs(DateDiff("d", dtStart, dtEnd))
    For intLoop = 0 To intDay
        strFileName = Dir(strFilePath & "\*_" & Format(dtStart + intLoop, "yyyymmdd") & "_*.txt")
        Do While strFileName <> ""
            strLine = ReadUniFile(strFilePath & "\" & strFileName)
            If strLine <> "" Then
                arrLine = Split(strLine, Chr(&H9))
                If UBound(arrLine) > 0 Then
                    For i = 0 To UBound(arrLine)
                        If i = 1 Then
                            strSample = arrLine(i)
                        End If
                        If i = 9 Then
                            strGerm = arrLine(i)
                        End If
                        If i = 39 Then
                            strDate = Format(arrLine(i), "yyyy-mm-dd hh:mm:ss")
                            If Not IsDate(strDate) Then
                                strDate = Format(Now, "yyyy-mm-dd hh:mm:ss")
                            End If
                        End If
                        If i >= 40 Then
                            If j = 2 Then 'ѭ����������
                                strAnti = arrLine(i)
                                If UCase(arrLine(i + 2)) = "SUSC" Then
                                    strItem = strItem & "|" & strAnti & "|" & "S^1^" & Replace(arrLine(i + 1), " ", "")
                                ElseIf UCase(arrLine(i + 2)) = "RESIST" Or UCase(arrLine(i + 2)) = "NOINTP" Then
                                    strItem = strItem & "|" & strAnti & "|" & "R^1^" & Replace(arrLine(i + 1), " ", "")
                                ElseIf UCase(arrLine(i + 2)) = "INTER" Then
                                    strItem = strItem & "|" & strAnti & "|" & "I^1^" & Replace(arrLine(i + 1), " ", "")
                                End If
                                j = -1
                            End If
                            j = j + 1
                        End If
'                        Call WriteLisLog1(arrLine(i))
                    Next
                End If
            End If
            If strItem <> "" And strGerm <> "" Then
                If strResult(UBound(strResult)) <> "" Then ReDim Preserve strResult(UBound(strResult) + 1)
                strResult(UBound(strResult)) = strDate & "|" & strSample & "| |" & strGerm & "|0" & strItem
                strItem = ""
            End If
            strFileName = Dir   '������һ���ļ�
        Loop
    Next
    ResultFromFile = strResult
    Exit Function
ErrorHand:
    Call WriteLog("������־", "", Err.Description)
    Err.Clear
End Function

