VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CS6400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����CS6400������
'����:����
'����:������
'�û�:�⳵�صڶ�����ҽԺ
'����:2015.01.15
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


    Dim varLine As Variant, varData As Variant, str���� As String, str�걾�� As String
    Dim iRow As Integer, jRow As Integer, str�����˱걾�ʿ� As String, str��Ŀ As String, strResultTmp As String
    Dim strTemp As String, strCode As String, strCmdR As String
    Dim lngS As Long, lngE As Long
    On Error GoTo ErrH
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = "": str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": str��Ŀ = ""
    strSource = Replace(strSource, " ", "")
    strSource = Replace(strSource, Chr(&H0), "")
    strCode = Replace(strSource, " ", "")

    If strCode Like "*" & Chr(11) & "MSH|*OBX*" & Chr(13) & Chr(28) & Chr(13) & "*" Then
        Do While strCode <> ""
            iRow = iRow + 1
            lngS = InStr(strCode, "MSH")
            lngE = InStr(lngS + 3, strCode, Chr(13) & Chr(28) & Chr(13))
            If lngE > 0 Then
                strTemp = Mid(strCode, lngS - 1, lngE + 1)
                strTemp = Trim(strTemp)
                strCode = Mid(strCode, lngE + 3)
            Else
                strTemp = Mid(strCode, IIf(lngS = 0, 1, lngS), lngE + 1)
                strTemp = Trim(strTemp)
                strCode = ""
            End If
            varData = Split(strTemp, Chr(13))
            If UBound(varData) > 0 And varData(0) Like "*ORU*" Then
                strCmdR = Replace(varData(0), "ORU", "ACK") & Chr(13) & "MSA|AA|" & Mid(Mid(varData(0), InStr(varData(0), "ORU^R01") + 8), 1, InStr(Mid(varData(0), InStr(varData(0), "ORU^R01") + 8), "|") - 1) & "||||0|" & Chr(13) & Chr(28) & Chr(13)
                strCmd = "0|" & GetStrCmd(strCmdR)
            End If
            For jRow = 0 To UBound(varData)
                If Mid(varData(jRow), 2) Like "MSH*" Then
                    varLine = Split(varData(jRow), "|")
                    If varLine(15) = 0 Then
                        str�����˱걾�ʿ� = " |ѪҺ|0"
                    Else
                        str�����˱걾�ʿ� = " |ѪҺ|1"
                    End If
                End If
                If varData(jRow) Like "OBR*" Then
                    varLine = Split(varData(jRow), "|")
                    str���� = Mid(varLine(7), 1, 4) & "-" & Mid(varLine(7), 5, 2) & "-" & Mid(varLine(7), 7, 2) & " " & Mid(varLine(7), 9, 2) & ":" & Mid(varLine(7), 11, 2) & ":" & Mid(varLine(7), 13, 2)
                    str�걾�� = varLine(3)
                ElseIf varData(jRow) Like "OBX|?|*" Or varData(jRow) Like "OBX|??|*" Then
                    varLine = Split(varData(jRow), "|")
                    If str��Ŀ = "" Then
                        str��Ŀ = Replace(varLine(3), "^1", "") & "|" & varLine(4)
                    Else
                        str��Ŀ = str��Ŀ & "|" & Replace(varLine(3), "^1", "") & "|" & varLine(4)
                    End If
                End If
            Next
            If str�걾�� <> "" And str��Ŀ <> "" And str���� <> "" Then
                strResultTmp = str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & "|" & str��Ŀ
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strResultTmp
                str��Ŀ = "": str�걾�� = "": str���� = ""
            End If
        Loop
    Else
        strReserved = strCode
    End If
'    strCmd = strCmdR
    Exit Sub
ErrH:
    If Error = 1 Then
        Resume
    End If
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function

Private Function GetStrCmd(strCmd As String) As String
    Dim lngCur As Long
    Do While strCmd <> ""
        GetStrCmd = GetStrCmd & Mid(strCmd, 1, 1) & Chr(&H0)
        strCmd = Mid(strCmd, 2)
    Loop
    GetStrCmd = GetStrCmd
End Function

