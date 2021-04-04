VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_KJ680"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'������Bio-rad���� ȫ�Զ�ø����
'�ͺţ�KJ680
'ҽԺ����ɽ����ҽԺ

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
    On Error GoTo ErrHand '�򿪴�������
    '------------------------------------------------
    Dim varLine As Variant
    Dim strLine As String, strtmp As String
    Dim intloop As Integer, i As Integer

    If InStr(1, strSource, "end") = 0 Then
        strReserved = strSource
        Exit Sub
    End If

    '��ʼ����
    strSource = Mid(strSource, InStr(1, strSource, "begin"), InStr(1, strSource, "end") - InStr(1, strSource, "begin"))
    varLine = Split(strSource, Chr(&HD))

    For intloop = 1 To 8
        strLine = ""
        strtmp = Split(strSource, Chr(&HD))(intloop)
        For i = 1 To 12
            strLine = strLine & ";" & Trim(Mid(strtmp, 1 + (i - 1) * 6, 6))
        Next
        If Trim(strLine) <> "" Then
            strResult = strResult & "|" & Mid(strLine, 2)
        End If
    Next

    If strResult <> "" Then strResult = Mid(strResult, 2)

    '------------------------------------------------
    Exit Sub
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_KJ680"
    Close #lngErrFileNo
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    'ָ�붨λ����һ������֮��
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '��ȡ�Ӵ���ͬʱָ���ƶ�����ȡ���ַ���֮��
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ���ַ��������ַ�ָ���Ƶ���һ�п�ʼλ��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '��ȡһ��������Ŀ���ƣ������ַ�ָ���Ƶ���Ŀ����֮��
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

'

Public Function CmdAnalyse(strCmdType As String) As String
    '����:  �ֽ�����Ϊ����ִ�е�����
    
'    strControl = Me.cbo����.Text & ";" & Me.cbo�ο�����.Text & ";" & Me.cbo���Ƶ��.Text & ";" & Me.txt���ʱ�� & _
'                 ";" & Me.cbo���巽ʽ.Text & ":" & Me.cbo�հ���ʽ.Text
    
    Dim lngTime As Long
    Select Case strCmdType
        Case "Begin"                   '��ʼ
            CmdAnalyse = ""
        Case "End"                      '����
            CmdAnalyse = ""
        Case "Revert"                   '�ظ�����
            CmdAnalyse = ""
        Case "Play"                     '�������
            CmdAnalyse = ""
        Case "ReadData_TimeOut"
            CmdAnalyse = 90
        Case "ReadData"                  '�������
            CmdAnalyse = "EIA.READER RPLATE 2 0 H 2" & Chr(&HD)
        Case "450"
            CmdAnalyse = ""   '450��Ƭ
        Case "630"
            CmdAnalyse = ""   '2�벨ƬΪ"450"
        Case "����"                     '����
            CmdAnalyse = ""
        Case "����"                     '��
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '���Ƶ��,�������������������е����Ƶ���趨��Ӧ��ָ��
        Case "��"                       '
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '�հ���ʽ,�������������������еĿհ���ʽ�趨��Ӧ��ָ��
        Case "����"
            CmdAnalyse = ""
        Case "�Լ�"
            CmdAnalyse = ""
        Case "Out"                      '����
            CmdAnalyse = ""
        Case "In"                       '����
            CmdAnalyse = ""
        Case "SpecialConnType"
            CmdAnalyse = ""
        Case "Z5_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z10_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z15_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z20_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z25_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z30_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case Else
            If Left(strCmdType, 1) = "Z" Then
                '���ʱ��
                lngTime = Val(Mid(strCmdType, 2))
                If lngTime >= 0 And lngTime <= 60 Then
                    CmdAnalyse = "Z(" & Format(lngTime, "00") & ")"
                End If
            End If
        End Select
End Function