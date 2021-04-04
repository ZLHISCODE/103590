VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AX4030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�������ң���ɭ����
'�ͺţ�sysmex UF-500i ����aution max AX-4030
'���ͣ����������
'�û���üɽ������ҽԺ
'���ߣ�����
'ʱ�䣺2015/4/2

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31"), Optional ByVal lngendSampleNO As Long) As String()
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
    Dim objFile As New FileSystemObject, txtStream As textStream
    Dim strLine As String, varLine() As String, strResult() As String
    Dim intloop As Long, i As Long, j As Long, intDate As Integer
    Dim strSample As String, strDate As String, strQC As String, strChenal As String, strValue As String, strItem As String
    Dim strFileName As String, strFilePath As String
    Dim strImage As String, strImageName As String

    On Error GoTo hErr
    If strFile = "" Then Exit Function
    ReDim strResult(0)
    strResult(0) = ""
    strFilePath = objFile.GetParentFolderName(strFile)
    intDate = DateDiff("d", Format(dtStart, "yyyy-mm-dd"), Format(dtEnd, "yyyy-mm-dd"))    'ѭ��ȡ�������ڵ�����
    For intloop = 0 To intDate
        strFileName = Dir(strFilePath & "\" & Format(dtStart + intloop, "yyyyMdd") & "*.sdf")
        Do Until strFileName = ""
            Set txtStream = objFile.OpenTextFile(strFilePath & "\" & strFileName)
            Do Until txtStream.AtEndOfStream
                strLine = txtStream.ReadLine
                varLine = Split(strLine, ",")
                If UBound(varLine) >= 8 Then
                    If varLine(0) = "0" Then    'ȡʱ�䣬�걾��
                        strSample = varLine(3)
                        strDate = Format(varLine(1) & " " & varLine(7), "yyyy-mm-dd hh:mm:ss")
                    End If
                    If varLine(0) = "1" Then    'ȡ������������
                        strChenal = varLine(1)
                        strValue = IIf(varLine(3) = "", " ", varLine(3))
                        strItem = strItem & "|" & strChenal & "|" & strValue
                    End If
                ElseIf UBound(varLine) >= 4 And varLine(0) = "3" Then    'ȡ��������ͼ������
                    strImageName = varLine(2)
                    strImage = strImage & "^" & strImageName & ";120;1;" & varLine(3)
                End If
            Loop
            If strSample <> "" And strItem <> "" Then    'һ���걾ѭ�����
                If strResult(UBound(strResult)) <> "" Then ReDim Preserve strResult(UBound(strResult) + 1)
                strResult(UBound(strResult)) = strDate & "|" & strSample & "| |��Һ|0" & strItem
                strItem = ""
                If strImage <> "" Then
                    strResult(UBound(strResult)) = strResult(UBound(strResult)) & vbCrLf & Mid$(strImage, 2)
                    strImage = ""
                End If
            End If
            strFileName = Dir
        Loop
    Next
    ResultFromFile = strResult
    Exit Function
hErr:
    MsgBox "����������" & Err.Description, vbInformation, "�����ӿ�"
    Err.Clear
End Function

