VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_CHEMCLIN600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Chemclin600
'����:������������
'����:����
'�û�:��������У��УҽԺ

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    
    On Error GoTo hErr '�򿪴�������
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
        
    
    Dim OBJFILE As New FileSystemObject, objStream As textStream
    Dim strParentFile As String, STRFILENAME As String, strFilePath As String, strFileName1 As String, str_Segment As String
    Dim strResult() As String, intLop As Integer
    Dim intDate As Integer
    Dim var_Filed As Variant, var_Code As Variant, str_���� As String, str_�걾�� As String
    Dim varItem As Variant, lng_count As Integer, str_Return As String, strValue As String, strOD As String
    Dim intloop As Integer, str_Code As String, mstrCode As String, str��Ŀ As String
    Dim i As Integer
    
    If Len(strFile) = 0 Then Exit Function
    ReDim Preserve strResult(0)
    strResult(0) = ""
    strFilePath = OBJFILE.GetParentFolderName(strFile)
    intDate = DateDiff("d", dtStart, dtEnd)
    For intLop = 0 To intDate
        STRFILENAME = Format(DateAdd("d", intLop, dtStart), "yyyymmdd")
        strFileName1 = Dir(strFilePath & "\" & STRFILENAME & "*.res")
        Do Until STRFILENAME = ""
'            strFileName1 = Dir(strFilePath & "\" & strFileName & "*.res")
            Set objStream = OBJFILE.OpenTextFile(strFilePath & "\" & strFileName1)
            Do Until objStream.AtEndOfStream
                str_Segment = objStream.ReadLine
                str_Segment = Trim(str_Segment)
                If InStr(str_Segment, "<MethodName>") > 0 Then
                    str��Ŀ = Trim(Mid(str_Segment, InStr(str_Segment, ">") + 1, InStr(str_Segment, "</") - 1 - Len("<MethodName>")))
                    str_Return = "": str_�걾�� = "": str_���� = "": strValue = "": strOD = ""
                ElseIf InStr(str_Segment, "<report") > 0 Then
                    str_���� = Mid(str_Segment, InStr(str_Segment, "=") + 1, 16)
                    str_���� = Replace(str_����, "-", "")
                    str_���� = Replace(str_����, """", "")
                    str_���� = Format(Left(str_����, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2) & ":" & Right(str_����, 2), "yyyy-mm-dd hh:mm:ss")
                    str_Return = "": strValue = "": strOD = ""
                ElseIf InStr(str_Segment, "<sample") > 0 Then
                    var_Filed = Split(" " & str_Segment, " ")
                    For intloop = LBound(var_Filed) To UBound(var_Filed)
                        If InStr(var_Filed(intloop), "id=") > 0 Then
                           str_�걾�� = Trim(var_Filed(intloop))
                           str_�걾�� = Trim(Replace(str_�걾��, "id=", ""))
                           str_�걾�� = Replace(str_�걾��, """", "")
                        ElseIf InStr(var_Filed(intloop), "result=") > 0 Then
                           strValue = Trim(var_Filed(intloop))
                           strValue = Trim(Replace(strValue, "result=", ""))
                           strValue = Trim(Replace(strValue, "&#60;", "<"))
                           strValue = Replace(strValue, """", "")
                            If Split(strValue, ";")(0) = "&gt" Then
                                strValue = ">" & Split(strValue, ";")(1)
                            End If
                            If Split(strValue, ";")(0) = "&lt" Then
                                strValue = "<" & Split(strValue, ";")(1)
                            End If
                            strOD = ""
                        ElseIf InStr(var_Filed(intloop), "value=") > 0 Then
                            strOD = Trim(var_Filed(intloop))
                            strOD = Trim(Replace(strOD, "value=", ""))
                            str_Return = strValue '& "^" & strOD & "^0^0"
    '                        StrResult = StrResult & "||" & str_���� & "|" & str_�걾�� & "| |Ѫ��|0|" & str��Ŀ & "|" & str_Return
                            If strResult(UBound(strResult)) <> "" Then ReDim Preserve strResult(UBound(strResult) + 1)
                            strResult(UBound(strResult)) = str_���� & "|" & str_�걾�� & "| |Ѫ��|0|" & str��Ŀ & "|" & str_Return
                        End If
                    Next
                End If
            Loop
            strFileName1 = Dir
            If strFileName1 = "" Then Exit Do
        Loop
    Next
    ResultFromFile = strResult

    '------------------------------------------------
    Exit Function
    '----------------
hErr:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrMsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_File_CHEMCLIN600"
    Close #lngErrFileNo
End Function