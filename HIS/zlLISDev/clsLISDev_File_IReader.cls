VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_IReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�������߷�����
'�����ͺţ�IReader
'�����
'�û����ٷ�������ҽԺ
'���ߣ�������
'���ڣ�2014-11-28
'

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo ErrHand '�򿪴�������
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

    Dim aResults() As String
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As Scripting.textStream, strLine As String

    Dim str���� As String, str������Ŀ As String, str�걾�� As String, strIMG As String
    Dim strFilePath As String, strFileName As String
    Dim StrValue As String
    Dim str_Item As String, str_Value As String
    Dim lngS As Long, lngE As Long
    Dim j As Integer

    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    strFileName = UCase(Dir(strFilePath & "\*"))
    ReDim aResults(0)

    Do While strFileName <> ""
        str���� = "": str�걾�� = "": str������Ŀ = "": strIMG = ""
        If UCase(strFileName) Like "IREADER.TXT" Then
            If objFile.FileExists(strFilePath & "\" & strFileName) Then
                Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFileName)
                '                    j = 1
                Do While Not objStream.AtEndOfStream
                    strLine = objStream.ReadLine

                    If Len(strLine) > 0 Then

                        str���� = Split(strLine, ",")(3)
                        str���� = Format(str����, "yyyy-mm-dd hh:mm:ss")
                        str_Item = Split(strLine, ",")(1)

                        str�걾�� = Split(strLine, ",")(0)

                        str_Value = Split(strLine, ",")(2)

                        If str_Value = "" Then
                            str_Value = "0"
                        End If

                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "||" & str���� & "|" & str�걾�� & "|������|ѪҺ|0" & "|" & str_Item & "|" & str_Value

                    End If

                Loop

                '                ���˾Ͱ��ļ��ŵ�BAK��
                objStream.Close
                Set objStream = Nothing
                If Not objFile.FolderExists(strFilePath & "\Bak") Then Call objFile.CreateFolder(strFilePath & "\Bak")
                objFile.CopyFile strFilePath & "\" & strFileName, strFilePath & "\Bak\" & strFileName, True
                objFile.DeleteFile strFilePath & "\" & strFileName, True
            End If
        End If
        strFileName = Dir

    Loop
    aResults(UBound(aResults)) = Mid(aResults(UBound(aResults)), 3)
    ResultFromFile = aResults
    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_KJ680"
    Close #lngErrFileNo
End Function