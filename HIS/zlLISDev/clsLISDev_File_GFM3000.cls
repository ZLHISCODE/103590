VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_GFM3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�������߷�����
'�����ͺţ�GF-M3000
'ɽ������
'�û������������ҽԺ
'���ߣ�������
'���ڣ�2014-12-02
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

    Dim str���� As String, str������Ŀ As String, str�걾�� As String, strimg As String
    Dim strFilePath As String, strFilename As String
    Dim StrValue As String
    Dim str_Item As String, str_Value As String
    Dim lngS As Long, lngE As Long
    Dim j As Integer

    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    strFilename = UCase(Dir(strFilePath & "\*"))
    ReDim aResults(0)

    Do While strFilename <> ""
        str���� = "": str�걾�� = "": str������Ŀ = "": strimg = ""
        If UCase(strFilename) Like "RESULT.TXT" Then
            If objFile.FileExists(strFilePath & "\" & strFilename) Then
                Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename)
                '                    j = 1
                Do While Not objStream.AtEndOfStream
                    strLine = objStream.ReadLine

                    If str���� = "" Then

                        str���� = Trim(Mid(Split(strLine, "��������")(1), 1, 10)) & " " & Trim(Mid(Split(strLine, "����ʱ��")(1), 1, 8))

                        str���� = Format(str����, "yyyy-mm-dd hh:mm:ss")

                    Else
                        str_Item = Split(strLine, ",")(0)

                        If Split(strLine, ",")(2) <> "PC" And Split(strLine, ",")(2) <> "NC" And Split(strLine, ",")(2) <> "QC" Then
                            str�걾�� = Split(strLine, ",")(2)

                            str_Value = Split(strLine, ",")(7) & "^" & Split(strLine, ",")(4) & "^" & Split(strLine, ",")(5) & "^" & Split(strLine, ",")(6)

                            If str_Value = "" Then
                                str_Value = "0"
                            End If

                            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "||" & str���� & "|" & str�걾�� & "|������|ѪҺ|0" & "|" & str_Item & "|" & str_Value
                        End If
                    End If

                Loop

                '                ���˾Ͱ��ļ��ŵ�BAK��
                objStream.Close
                Set objStream = Nothing
                If Not objFile.FolderExists(strFilePath & "\Bak") Then Call objFile.CreateFolder(strFilePath & "\Bak")
                objFile.CopyFile strFilePath & "\" & strFilename, strFilePath & "\Bak\" & strFilename, True
                objFile.DeleteFile strFilePath & "\" & strFilename, True
            End If
        End If
        strFilename = Dir

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
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_File_GFM3000"
    Close #lngErrFileNo
End Function