VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AutoLumo_A2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AutoLumo A2000 ø���� �����Է�������
'����:��ͼ���﹤�����޹�˾
'����:����
'�û�:˫����ҽԺ
'����:2014-07-24
Private mstrCode As String

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    On Error GoTo errHand '�򿪴�������
    '------------------------------------------------
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim strCode As String, STRRESULTS() As String, strValue() As String, strLine As String
    Dim intLoop As Integer, strOD As Double, strSCO As Double, strCutOff As Double
    Dim intID As Long, strSample As String, strNumber As String, strCheal As String
    Dim intSerial As Integer, lngS As Long, lngE As Long

    strCode = strCode & strSource
    strReserved = strReserved & strSource
    lngS = InStr(strReserved, "{")
    lngE = InStr(strReserved, "}")
    Call writelislog("�յ���" & strReserved)
    '    If lngS = 0 Or lngE = 0 Then
    '        Exit Sub
    '    End If
    If lngS > 0 And lngE > 0 Then
        If Len(strReserved) = 0 Then Exit Sub
        strValue = Split(strReserved, "}")
        For intLoop = 0 To UBound(strValue)
            strLine = strValue(intLoop)
            If strLine <> "" Then
                strLine = Replace(strLine, "{", "")
                intSerial = Val(Split(strLine, ",")(0)) '���
                If intSerial = 5 Then '��ͨ����
                    strSample = Mid(Split(strLine, ",")(2), InStr(Split(strLine, ",")(2), "]") + 1) '�걾��
                    strCheal = Mid(Split(strLine, ",")(3), InStr(Split(strLine, ",")(3), "]") + 1) 'ͨ����
                    strNumber = Split(strLine, ",")(4) '������
                    strOD = Val(Split(strLine, ",")(5))
                    strSCO = Val(Split(strLine, ",")(6))
                    strCutOff = Format(strOD / IIf(strSCO = 0, 1, strSCO), "0.000")
                    '                    strResult = strResult & "||" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "|" & strSample & "|" & " |ѪҺ|0|" & strCheal & "|" & strSCO & "^" & strOD & "^" & strCutOff & "^" & strSCO
                    If Len(strSample) = 12 Then
                        strResult = strResult & "||" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "|^^" & Mid(Split(strLine, ",")(2), InStr(Split(strLine, ",")(2), "]") + 1) & "|" & " |ѪҺ|0|" & strCheal & "|" & strSCO & "^" & strOD & "^" & strCutOff & "^" & strSCO
                    Else
                        strResult = strResult & "||" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "|" & Mid(Split(strLine, ",")(2), InStr(Split(strLine, ",")(2), "]") + 1) & "^^|" & " |ѪҺ|0|" & strCheal & "|" & strSCO & "^" & strOD & "^" & strCutOff & "^" & strSCO
                    End If
                    If intSerial = 5 Then '���Է�������
                        strCmd = strCmd & "{6,0," & strNumber & ",[S]" & strSample & ",[S],[S]" & strCheal & "}"
                    End If
                End If
                If intSerial = 9 Then '����˫��
                    strResult = Format(Now, "yyyy-mm-dd") & "|^^" & Mid(Split(strLine, ",")(2), InStr(Split(strLine, ",")(2), "]") + 1)
                    strCmd = "1|"
                    Call writelislog("strResult:" & strResult)
                End If
            End If
        Next
        strReserved = ""
    End If
    If Left(strResult, 2) = "||" Then
        strResult = Mid(strResult, 3)
    End If
    Exit Sub

    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_AutoLumo_A2000"
    Close #lngErrFileNo
End Sub
Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    
    '������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    Dim blnEnd As Boolean
    Dim strBuffer As String, strSampleID As String, strSampleNO As String, strItem As String, sampleno As String
    Dim i As Long, aItem() As String
    Dim strRackNO As String, iPos As Long
    
    blnSuccess = True: SendSample = ""
    Call writelislog("strSample:" & strSample)
    Call writelislog("strResponse:" & strResponse)
    blnSuccess = True
    aItem = Split(strSample, "|")
    If strResponse Like "{*[S]*}" Then
        strSampleID = Mid(Split(strResponse, ",")(2), InStr(Split(strResponse, ",")(2), "]") + 1, 12)
        strSampleNO = Mid(Split(strResponse, ",")(0), InStr(Split(strResponse, ",")(0), "{") + 1)
        sampleno = aItem(1)
        If Len(strSampleID) = 12 And Val(strSampleNO) = 9 Then '��ָ֯��
            For i = 10 To UBound(aItem)
'                SendSample = "{10,0,[S]" & strSampleID & ",1,0,[S],3,"
                strItem = strItem & "[S]" & aItem(i) & ",1,"
            Next
            If Len(strItem) > 0 Then
                strItem = Mid(strItem, 1, Len(strItem) - 1)
                SendSample = "{10,0,[S]" & strSampleID & ",1,0,[S]," & Val(UBound(aItem) - 9) & "," & strItem & "}"
'                SendSample = "{10,0,[S]" & sampleno & ",1,0,[S]," & "3" & "," & strItem & "}"
                Call writelislog("SendSample:" & SendSample)
            End If
            iSendStep = 0
        End If
   End If
End Function
Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_A2000"
    Dim OBJFILE As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim STRFILENAME As String
    STRFILENAME = strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not OBJFILE.FileExists(STRFILENAME) Then Call OBJFILE.CreateTextFile(STRFILENAME)
    Set objStream = OBJFILE.OpenTextFile(STRFILENAME, ForAppending)
    objStream.WriteLine ("--------------------------------------------------------------------------------------------------------------------------------------------------------------")
    objStream.WriteLine (strLog)
    'objStream.WriteLine ("--------------------------------------------------------------------------------------------------------------------------------------------------------------")
    objStream.Close
    Set objStream = Nothing
End Function