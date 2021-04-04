VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_URIT_180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:URIT_180 ��Һ����
'����:�����ص��Ӽ������޹�˾
'����:����
'�û�:�˱��е�������ҽԺ
'���ڣ�2014-12-12

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
    Dim strCode As String, strValue() As String, strLine As String, strData As String
    Dim lngS As Long, lngE As Long
    Dim intLoop As Integer
    Dim strSample As String, strDate As String, strItems As String
    
On Error GoTo hErr
    If strSource <> "" Then strReserved = strReserved & strSource
    Call writelislog("�յ���" & strReserved)
    lngS = InStr(strReserved, Chr(&H2))
    Call writelislog("��ʼ����" & lngS)
    lngE = InStr(strReserved, Chr(&H3))
    Call writelislog("��������" & lngE)
    If lngS > 0 And lngE > 0 Then
        strCode = strReserved
        Do Until strCode = ""
            strData = Mid$(strCode, lngS, lngE - 1)
            strCode = Mid(strCode, lngE + 1)
            strValue = Split(strData, vbCrLf)
            For intLoop = LBound(strValue) To UBound(strValue)
                strLine = strValue(intLoop)
                If InStr(strLine, "NO") Then '��ȡʱ��
                    strSample = Mid(strLine, InStr(strLine, ".") + 1, 6)
                    strDate = Format(Mid(strLine, 12) & " " & Format(Now, "hh:mm:ss"), "yyyy-mm-dd HH:mm:ss")
                    
                End If
                If intLoop >= 4 And strLine <> "" Then '��ȡ��������
                    strLine = Replace(strLine, "mmol/L", "")
                    strLine = Replace(strLine, "Cell/uL", "")
                    strLine = Replace(strLine, "g/L", "")
                    strLine = Replace(strLine, "umol/L", "")
'                    strResult = strDate & "|" & Val(strSample) & "| |��Һ|0"
                    strItems = strItems & "|" & Trim(Mid(strLine, 2, 4)) & "|" & Trim(Mid(strLine, 6))
                End If
            Next
            If strSample <> "" And strItems <> "" Then
                strResult = strResult & "||" & strDate & "|" & Val(strSample) & "| |��Һ|0" & strItems
            End If
        Loop
        If strResult <> "" Then
            strResult = Mid(strResult, 3)
'            strResult = Replace(strResult, "Norma", "Normal")
            Call writelislog("strResult��" & strResult)
        End If
    End If
    Exit Sub
hErr:
    MsgBox "zlCDLISDev.clsLISdev_URIT_180.Analyse" & Err.Description, vbInformation, "�ӿڴ���"
    Err.Clear
End Sub

Private Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_URIT180_"
    Dim OBJFILE As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim STRFILENAME As String
    STRFILENAME = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not OBJFILE.FileExists(STRFILENAME) Then Call OBJFILE.CreateTextFile(STRFILENAME)
    Set objStream = OBJFILE.OpenTextFile(STRFILENAME, ForAppending)
    objStream.WriteLine (Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & strLog)
    objStream.Close
    Set objStream = Nothing
End Function