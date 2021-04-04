VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MissionU500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Mission U500
'����:����
'����:������
'�û�:�½�����̩��������ҽԺ
'ʱ��:2014-05-29

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
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim varLine As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strLine As String
    Dim iCount As Integer, lngLength As Long
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, Chr(&H3&)) > 0 Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            lngLength = InStr(strReserved, Chr(&H3)) - InStr(strReserved, Chr(&H2))
            strLine = Mid(strReserved, InStr(strReserved, Chr(&H2)), lngLength + 3)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H3)) + 1)

            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = ""
            '����
            iCount = 0
            varLine = Split(strLine, vbCrLf)
            strItem = ""
            For iRow = 0 To UBound(varLine)

                Select Case iCount
                    Case 1 '����
                    str���� = Mid(Trim(varLine(iRow)), 1, 4) & "-" & TrunMonth(Mid(Trim(Mid(Trim(varLine(iRow)), 5)), 1, 3)) & "-" & Mid(Trim(Mid(Trim(Mid(Trim(varLine(iRow)), 5)), 4)), 1, 2) & " " & Trim(Mid(Trim(Mid(Trim(Mid(Trim(varLine(iRow)), 5)), 4)), 3))
                    Case 3 '�걾��
                    If InStr(varLine(iRow), ":") > 0 Then varLine(iRow) = Replace(varLine(iRow), ":", "")
                    str�걾�� = Val(Trim(Left(Mid(varLine(iRow), InStr(varLine(iRow), "-") + 1), 6)))
                    str�����˱걾�ʿ� = " |��Һ|0"
                    Case 4 To 14
                    strItem = strItem & "|" & Trim(MidUni(varLine(iRow), 2, 3)) & "|" & Trim(MidUni(varLine(iRow), 5))
                End Select
                If varLine(iRow) <> "" Then iCount = iCount + 1
            Next
            If strItem <> "" And Val(str�걾��) <> 0 Then
                strResult = IIf(strResult <> "", strResult & "||", "") & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
            End If
            If InStr(strReserved, Chr(&H3&)) = 0 Then
                blnEnd = True
            End If
        Loop
    Else
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If

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


Private Function MidUni(ByVal strTemp As String, ByVal Start As Long, Optional ByVal Length As Long) As String
'���ܣ������ݿ����õ��ַ������Ӽ���Ҳ���Ǻ��ְ������ַ��㣬����ĸ����һ��
    If Length = 0 Then Length = LenB(strTemp)
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    'ȥ�����ܳ��ֵİ���ַ�
    MidUni = Replace(MidUni, Chr(0), "")
    MidUni = Trim(Replace(MidUni, Chr(0), ""))
    MidUni = Trim(Replace(MidUni, "0.56mmol/L", ""))
    MidUni = Trim(Replace(MidUni, "3.5umol/L", ""))
    MidUni = Trim(Replace(MidUni, "neg", ""))
    MidUni = Trim(Replace(MidUni, "g/L", ""))
    MidUni = Trim(Replace(MidUni, "mg/dL", ""))
    MidUni = Trim(Replace(MidUni, "Leu/uL", ""))
End Function

Private Function TrunMonth(ByVal strMonth As String) As String
    'ת��Ӣ�ĵ��·�Ϊ����
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select
End Function