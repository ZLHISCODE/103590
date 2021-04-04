VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_FIA200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:FIA200
'����:ʱ��ֱ���
'����:������
'�û�:������ҽԺ
'����:2012-05-3
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

    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strChenel As String, str����Ա As String, str���� As String
    Dim strFilePath As String, intCount As Long, lngLoop As Long

    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim lngRow As Long, varLine As Variant
    Dim strItem As String
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    '
    lngLoop = DateDiff("d", dtStart, dtEnd)

    For intCount = 0 To lngLoop
        str���� = Format(dtStart, "yyyyMMdd")
        strFile = Dir(strFilePath & "\*" & str���� & "*.TXT")
        Do While strFile <> ""
            str���� = Format(Mid(strFile, 1, 4) & "-" & Mid(strFile, 5, 2) & "-" & Mid(strFile, 7, 2) & " " & _
                      Mid(strFile, 9, 2) & ":" & Mid(strFile, 11, 2) & ":" & Mid(strFile, 13, 2), "yyyy-mm-dd hh:mm:ss")
            strItem = Mid(strFile, 16, 5)
            Open strFilePath & "\" & strFile For Input As #1    ' ���ļ���
            Do While Not EOF(1)    ' ѭ�����ļ�β��
                Line Input #1, strLine    ' ����һ�����ݲ����丳��ĳ������

                If strLine <> "" Then
                    str�걾�� = Mid(strLine, 1, InStr(strLine, " ") - 1)
                    If Val(str�걾��) > 0 Then
                        str������ = strItem & "|" & Mid(Replace(strLine, " ", ""), Len(str�걾��) + 1)
                        If str���� <> "" And strItem <> "" And str������ <> "" Then
                                If aResults(UBound(aResults)) <> "" Then
                                    ReDim Preserve aResults(UBound(aResults) + 1)
                                End If
                                aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |ѪҺ|0|" & str������
                        End If
                    End If
                End If
                lngRow = lngRow + 1
                DoEvents
            Loop
            Close #1    ' �ر��ļ���
            strFile = Dir
        Loop

    Next

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_FIA200.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function








