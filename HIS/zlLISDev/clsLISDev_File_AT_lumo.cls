VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AT_lumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ø����
'����:��ͼ
'����:
'�û�:�ܿ����Ÿ���
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

    Dim aResults() As String
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As Scripting.textStream, strLine As String
    Dim str���� As String, str������Ŀ As String, str���Խ�� As String, str������� As String
    Dim blnOk As Boolean '�Ƿ񰴱걾����ȡ����
    Dim strFilePath As String
    Dim strCO As String
    Dim ii As Long
    Dim varcol As Variant
    Dim str_bz As String
    Dim strline1 As String

    strFilePath = UCase(objFile.GetParentFolderName(strFile))

    If Dir(strFile) = "" Then Exit Function
    strFile = UCase(strFile)

    '
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    ii = 0
    str_bz = "End"
    ReDim aResults(0)
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If Trim(strLine) = "" Then
            str_bz = "Begin"
            ii = 1
        ElseIf Trim(strLine) = "" Then
            ii = 0
            str_bz = "End"
        End If
        If str_bz = "Begin" And ii <= 5 Then
            strline1 = strline1 & "|" & strLine
            ii = ii + 1
        End If

        If str_bz = "End" And Len(Trim(strline1)) > 10 Then
            varcol = Split(strline1, "|")
            If Not Nvl(aResults(UBound(aResults))) = "" Then
                ReDim Preserve aResults(UBound(aResults) + 1)
            End If
            str���� = Split(varcol(2), ",")(0)
            If IsDate(str����) Then
                str���� = Format(str����, "yyyy-MM-dd HH:mm:ss")
            Else
                str���� = Format(Now, "yyyy-MM-dd HH:mm:ss")
            End If
            lngSampleNO = Trim(varcol(2))
            str������Ŀ = Trim(varcol(5))
            str���Խ�� = Trim(varcol(6))
            str������� = Trim(varcol(9))
            strCO = Format(Trim(varcol(10)), "0.000")
            strline1 = ""
            '            aResults(UBound(aResults)) = str���� & "|" & lngSampleNO & "| |ѪҺ|0|" & str������Ŀ & "|" & str���Խ�� & "^" & str������� & "^" & strCO & "^0"
            aResults(UBound(aResults)) = str���� & "|" & lngSampleNO & "| |ѪҺ|0|" & str������Ŀ & "|" & str���Խ�� & "^" & str������� & "^" & strCO & "^0"

        End If
    Loop

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