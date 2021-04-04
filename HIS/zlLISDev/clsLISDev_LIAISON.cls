VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LIAISON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'��������: LIAISON ��ѧ����
'�����ͺ�: LBYN7500B
'����:������
'ҽԺ:�Ĵ�ʡ��������ҽԺ
'����:����
'����:2014-10-08

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
                   strCmd As String)
    On Error GoTo errHand '�򿪴�������
    '------------------------------------------------
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim lng_S As Long, lng_E As Long, intLoop As Integer
    Dim strCode As String, strSample As String, strDate As String, strLine() As String, strValue As String
    Dim strBuffer As String
    Dim i As Integer

    
    strReserved = strReserved & strSource
    Call writelislog("���գ�" & strReserved)

    If InStr(strReserved, Chr(&H5)) > 0 Then                                         '�յ������� 05 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    If InStr(strReserved, Chr(&HA) & Chr(&HD)) > 0 Then                              '�յ��н����� 03 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    If InStr(strReserved, Chr(&H3)) > 0 Then                                         '�յ���ʼ�� 03 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    If InStr(strReserved, Chr(&H4)) > 0 Then                                         '�յ���ʼ�� 04 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    
    lng_S = InStr(strReserved, Chr(&H5)) '��ʼ��
    lng_E = InStr(strReserved, Chr(&H4)) '������
    i = 1
    If lng_S > 0 And lng_E > 0 Then
        Do Until strReserved = ""
            lng_S = InStr(strReserved, Chr(&H5)) '��ʼ��
            lng_E = InStr(strReserved, Chr(&H4)) '������
            strBuffer = Mid$(strReserved, lng_S + 1, lng_E + lng_S - 1)
            Call writelislog("������" & strBuffer)
            strReserved = Mid$(strReserved, lng_E + 1) '�´δ���
            Call writelislog("����������" & strReserved)
            strLine = Split(strBuffer, Chr(&HD))
            For intLoop = LBound(strLine) To UBound(strLine)
                If InStr(strLine(intLoop), "4R|") > 0 And i = 1 Then
                    strDate = Split(strLine(intLoop), "|")(12)
                    strDate = Format(Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Right(strDate, 2), "yyyy-mm-dd hh:mm:ss")
                    i = i + 1
                End If
                If InStr(strLine(intLoop), "O|") > 0 Then
                    strSample = Split(strLine(intLoop), "|")(2)
                End If
                If InStr(strLine(intLoop), "Liaison") > 0 And InStr(strLine(intLoop), "1H") = 0 Then
                    strValue = strValue & "|" & Replace(Split(strLine(intLoop), "|")(2), "^", "") & "|" & Split(strLine(intLoop), "|")(3)
                End If
            Next
            If Val(strSample) <> 0 And strValue <> "" Then
                strResult = strResult & "||" & strDate & "|" & strSample & "| |Ѫ��|0" & strValue
            End If
        Loop
        If Left(strResult, 2) = "||" Then
            strResult = Mid(strResult, 3)
        End If
    End If
    Exit Sub

    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_LIAISON"
    Close #lngErrFileNo
End Sub

Public Function writelislog(strLog As String)
    Const strFile As String = "LisLog_LIAISON_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFilename As String
    strFilename = App.Path & "\" & strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFilename) Then Call objFile.CreateTextFile(strFilename)
    Set objStream = objFile.OpenTextFile(strFilename, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function