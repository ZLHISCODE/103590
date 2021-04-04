VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_H700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim blnEndSample As Boolean '�걾����
    Dim iPos As Long
    Dim aDatas() As String

    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""

    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        '����һ�����ݵĿ�ʼλ��
        iStrPoint = InStr(iStrPoint, strSource, "R")    '��ʼλ��
        '�ж�һ�����Ƿ�ʼ
        If iStrPoint > 0 Then
            '��¼һ���������
            iStrPoint = iStrPoint + 1
            iCurrGrpEntry = iStrPoint - 1
            '�����Ƿ�����
            iPos = InStr(iStrPoint, strSource, "H")   '����λ��
            strBuffer = LTrim(ReadSubString(strSource, iStrPoint, iPos))  'һ������
            If iPos > 0 Then
                '                        R#0002#01#13.408 # #H
                '                        R#0003#01#6.724  # #H
                If strBuffer Like "*#*#*#*#*#*" Then
                    '��ǰ�걾��Ϣ��Ч
                    aDatas = Split(strBuffer, "#")
                    '����ʱ��
                    'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                    strData = Format(Now, "yyyy-mm-dd hh:mm:ss") '
                    '������ţ�����Ϊ����
                    strData = strData & "|" & aDatas(1)
                    '������
                    strData = strData & "| " '��������֧��
                    '�걾
                    strData = strData & "|����Ѫ" '��������֧��
                    '�ʿ�Ʒ
                    strData = strData & "|0"
                    strData = strData & "|CRP|" & Trim(aDatas(3))
                    blnReadItem = False
                    '�������ѽ���
                    iCurrGrpEntry = 0
                    strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                End If
            Else
                Exit Do
            End If
            'ָ�붨λ����һ������֮��
        Else
            Exit Do
        End If
    Loop
    '����δ������ԭʼ��
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    End If
    '���豸���͡������յ�������
    strCmd = Chr(&H6&)
    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLisDev_DSI905"
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