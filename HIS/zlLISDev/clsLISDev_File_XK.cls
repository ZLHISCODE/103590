VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_XK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:ϸ������/ҩ�������� XK�� ΢����
'����:ɽ���ο�����Ƽ����޹�˾
'����:������
'�û�:ɽ���ӽ�����ְ��ҽԺ
'ʱ��:2012-12-26
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()

    Dim aResults() As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim lngResultCount As Long
    Dim objStream As Scripting.TextStream, strLine As String

    Dim str�걾�� As String, strҩ����Ŀ As String
    Dim str���� As String, strϸ�� As String, str���ж� As String
    If Dir(strFile) = "" Then Exit Function
    Set objStream = objFile.OpenTextFile(strFile, ForReading)

    lngResultCount = -1
    ReDim aResults(0)
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If strLine Like "@*" Then

            ''������ϸ�����ϵ������ �ȷ�����һ��ϸ�����
'            If Val(str�걾��) > 0 And strϸ�� <> "" And str���� <> "" And strҩ����Ŀ <> "" Then
'                lngResultCount = lngResultCount + 1
'                ReDim Preserve aResults(lngResultCount)
'                aResults(lngResultCount) = str���� & "|" & Val((str�걾��)) & "| |" & strϸ�� & "|0" & strҩ����Ŀ
'                str���� = "": str�걾�� = "": strϸ�� = "": strҩ����Ŀ = ""
'            End If

            'һ���걾��ʼ,ϸ�����
            If Split(strLine, "|")(1) <> "" Then
                str�걾�� = Trim(Split(strLine, "|")(1))
            End If
            str���� = Split(strLine, "|")(2)
'            If str���� = "2012-11-14" Then
'                MsgBox "-------"
'            End If
            If str���� >= Format(dtStart, "yyyy-MM-dd") And str���� <= Format(dtEnd, "yyyy-MM-dd") Then

                If lngSampleNo <> -1 Then    'Ҫָ���걾��
                    If str�걾�� <> lngSampleNo Then
                        '�걾�Ų�ͬ��������
                        str�걾�� = ""
                        str���� = ""
                        strϸ�� = ""
                    End If
                End If
            Else
                '���ڲ�����������
                str�걾�� = ""
                str���� = ""
                strϸ�� = ""
            End If
        ElseIf InStr(strLine, "#") > 0 Then
            strϸ�� = Split(strLine, "|")(2)
        ElseIf strLine Like "%*" Then
            'ҩ�����
            If str�걾�� <> "" Then
                '�޸�ϸ��������,����r*,blac,esbl ��������ҩby����ǿ.
                str���ж� = Split(strLine, "|")(6)
                Select Case str���ж�
                    Case "����": str���ж� = "S"
                    Case "��ҩ": str���ж� = "R"
                    Case "�н�": str���ж� = "I"
                End Select
                strҩ����Ŀ = strҩ����Ŀ & "|" & Split(strLine, "|")(2) & "|" & str���ж� & "^1" & "^" & Split(strLine, "|")(4)
             End If
        ElseIf strLine Like "$*" Or strLine = "" Then
            'һ���걾����
            If Val(str�걾��) >= 0 And strϸ�� <> "" And str���� <> "" And strҩ����Ŀ <> "" Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = str���� & "|" & Val(str�걾��) & "| |" & strϸ�� & "|0" & strҩ����Ŀ
                str���� = "": str�걾�� = "": strϸ�� = "": strҩ����Ŀ = ""
            End If
            
            If Val(str�걾��) >= 0 And strϸ�� <> "" And str���� <> "" And strҩ����Ŀ = "" Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = str���� & "|" & Val(str�걾��) & "| |ϸ���������|0|ϸ������|" & strϸ��
                str���� = "": str�걾�� = "": strϸ�� = "": strҩ����Ŀ = ""
            End If
            
        End If
    Loop
    ResultFromFile = aResults

End Function








