VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_Everesys_A1800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��ѧ����ⶨ��
'����:everesys a1800 I
'����:���ݳ��⻪ҽ����ҽѧ���� HYBIOME
'����:����
'�û�:�������������ҽԺ
'ʱ��:2014-12-04

Private mstrCode As String    '��������
Private mstrRack As String    'ɨ����ʱ�����̺�,�ܺŵ���Ҫ���ص���Ϣ.
Private mstrWaitSend As String '�����ͱ걾
Private mstrQCNO As String      '��QC��ţ�ҪתΪ�̶��ı걾��

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
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim strCode As String, strLine As String, strCheel As String, strValue As String, strItem As String, strArray() As String
    Dim intLoop As Integer
    Dim lngS As Long, lngE As Long
    Dim strSample As String, str���� As String, strQC As String, strBH  As String
    Dim strDate As String
    Dim objFileSystem As New FileSystemObject, objStream As textStream
    
    strCode = strCode & strSource
    If Not objFileSystem.FileExists(App.Path & "\" & "A1800.ini") Then '��ȡ���ݿ���Ϣ
        Set objStream = objFileSystem.CreateTextFile(App.Path & "\" & "A1800.ini")
            objStream.WriteLine "[A1800]"
            objStream.WriteLine "STAT=3000"
            objStream.Close
            Set objStream = Nothing
    End If
    strBH = ReadIni("A1800", "STAT", App.Path & "\" & "A1800.ini")
    lngS = InStr(strCode, Chr(&H2))
    lngE = InStr(strCode, Chr(&H3))
    If lngS > 0 And lngE > 0 Then
        strArray = Split(strCode, Chr(&HD))
        For intLoop = 0 To UBound(strArray)
            strLine = strArray(intLoop)
            If InStr(strLine, "O|") > 0 Then '���ұ걾�ź�ʱ��
                strSample = Split(Split(strLine, "|")(2), "^")(0)
                str���� = Split(Split(strLine, "|")(2), "^")(1)
                strDate = Split(strLine, "|")(6)
                strDate = Format(Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Right(strDate, 2), "yyyy-mm-dd hh:mm:ss")
'            ElseIf InStr(strLine, "R|") > 0 Then '��ȡ��������
            ElseIf strLine Like "R|*" Then
                strCheel = Split(Split(strLine, "|")(2), "^")(1)
                strValue = Replace(Split(strLine, "|")(3), "^", "")
                strItem = strItem & "|" & strCheel & "|" & strValue
            End If
        Next
        If strSample <> "" And strItem <> "" Then
            strResult = strDate & "|" & IIf(str���� = "0", Val(strBH) + Val(strSample) & "^1^", strSample) & "| |Ѫ��|0" & strItem
            Call writelislog(strResult)
        End If
    End If
End Sub

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_A1800_"
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