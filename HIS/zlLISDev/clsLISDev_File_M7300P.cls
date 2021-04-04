VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_M7300P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:���m7300p
'����:
'����:����
'�û�:������ҽԺ
'ʱ��:2014-10

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
    '   ��׼���������ļ���ȡ������
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
    '   �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������

    Dim str����Ա As String, str���� As String, strTime As String, str�걾�� As String, str������ As String
    Dim i As Long

    Dim objFSO As FileSystemObject
    Dim objText As textStream
    Dim strLine As String

    Set objFSO = New FileSystemObject
    Dim aResults() As String
    Dim lngLoop As Long
    Dim strPath As String
    Dim strErr As String             '������ʾ

    ReDim Preserve aResults(0)
    On Error GoTo errHandle
    strPath = Mid(strFile, 1, InStrRev(strFile, "\"))
    If dtEnd = CDate("3000-12-31") Then
        dtEnd = Now
    End If
    lngLoop = DateDiff("d", dtStart, dtEnd)
    If lngLoop < 0 Then
        strErr = "��ʼ���ڲ��ܴ��ڽ������ڣ�"
        ResultFromFile = aResults
        Exit Function
    End If
    For i = 0 To lngLoop    'ѭ����ȡ�������ļ�
        strFile = Dir(strPath & Format(DateAdd("d", i, dtStart), "YYMMDD") & ".txt") 'ʱ����ڵ��ļ�
        Do While strFile <> ""
            Set objText = objFSO.OpenTextFile(strPath & strFile, _
                ForReading, False, -2)
            Do Until objText.AtEndOfStream
                strLine = Trim(objText.ReadLine)
                strLine = Replace(strLine, vbTab, " ")
                strLine = TrimString(strLine)
                If InStr(strLine, "�ӿں�") > 0 And UBound(Split(strLine, " ")) >= 6 Then
                    If str���� = "" Then str���� = "3014-01-01"
                    If Val(str�걾��) <> Val(Split(strLine, " ")(1)) Or _
                       (InStr(Split(strLine, " ")(4), "���") > 0 And Format(str����, "yyyy-mm-dd") <> Format(Split(strLine, " ")(5), "yyyy-mm-dd")) Or _
                       (InStr(Split(strLine, " ")(4), "���") = 0 And Format(str����, "yyyy-mm-dd") <> Format(Split(strLine, " ")(6), "yyyy-mm-dd")) Then
                        str�걾�� = Val(Split(strLine, " ")(1))
                        If InStr(Split(strLine, " ")(4), "���") > 0 Then
                            If UBound(Split(strLine, " ")) >= 6 Then
                                strTime = " " & Split(strLine, " ")(6)
                            Else
                                strTime = ""
                            End If
                            str���� = Format(Split(strLine, " ")(5) & strTime, "yyyy-mm-dd hh:nn:ss")
                            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                            str������ = " "

                        Else
                            If UBound(Split(strLine, " ")) >= 7 Then
                                strTime = " " & Split(strLine, " ")(7)
                            Else
                                strTime = ""
                            End If
                            str���� = Format(Split(strLine, " ")(6) & strTime, "yyyy-mm-dd hh:nn:ss")
                            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                            str������ = Split(strLine, " ")(4)
                            str������ = IIf(str������ = "", " ", str������)

                        End If
                        aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |ѪҺ|0|" _
                                 & UCase(Trim(Replace(Replace(Split(strLine, " ")(3), "�ӿںţ�", ""), ":", ""))) & "|" & str������
                    Else
                        If InStr(Split(strLine, " ")(4), "���") > 0 Then
                            str������ = " "

                        Else
                            str������ = Split(strLine, " ")(4)
                            str������ = IIf(str������ = "", " ", str������)

                        End If
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" _
                                 & UCase(Trim(Replace(Replace(Split(strLine, " ")(3), "�ӿںţ�", ""), ":", ""))) & "|" & str������

                    End If
                Else
                    str���� = "3014-01-01"
                    str�걾�� = ""
                End If

            Loop
            Set objText = Nothing
            strFile = Dir
        Loop
    Next i
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_M7300P" & vbCrLf & strErr, strFile, Err.Number & " " & Err.Description

End Function
  

Private Function TrimString(strIn As String) As String
'�滻�ַ����ж���ո�Ϊ�����ո�
    Dim strTemp As String
    strTemp = strIn
    Do While InStr(1, strTemp, Space(2)) > 0
        strTemp = Replace(strTemp, Space(2), Space(1))
    Loop
    TrimString = strTemp
End Function