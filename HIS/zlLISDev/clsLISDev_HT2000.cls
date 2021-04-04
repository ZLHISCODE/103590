VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HT2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����� HT2000
'����:
'����:����
'�û�:�Ͼ���������ҽԺ

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
    
    Dim varLine As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    Dim strImage As String  'ͼ������
    Dim varImage As Variant 'ͼ������
    Dim strItemVal As String   '���ֵ
    Dim strData As String
    Dim intloop As Integer
    Dim strTitle As String
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    
    'û���ҵ�������ʱ�˳�
    If InStr(strSource, Chr(&H3)) <= 0 Then strReserved = strSource: Exit Sub
    strLine = Mid(strSource, InStr(strSource, Chr(&H2)), InStr(strSource, Chr(&H3)))
    'ֻ����һ���걾������ȷ���
    strReserved = Mid(strSource, InStr(strSource, Chr(&H3)) + 1)
    If InStr(strReserved, "") <= 0 Then
        strReserved = ""
    End If
    
    varLine = Split(strLine, vbCr)
    str���� = Mid(varLine(1), 10) & Trim(Mid(varLine(2), 12))
    str�걾�� = Trim(Mid(varLine(1), 6, 5))
    
    strTitle = str���� & "|" & str�걾�� & "| |Ѫ��|0"
    
    'д����
    For intloop = 4 To UBound(varLine) - 1
        
        If intloop = 9 Then 'SG
            strItemVal = Trim(Mid(varLine(intloop), 7, 7))
            strData = strData & "|" & Trim(Mid(varLine(intloop), 2, 4)) & "|" & strItemVal
        ElseIf intloop = 5 Then 'NIT
            strItem = Trim(Mid(varLine(intloop), 8, 2))            '���
            strData = strData & "|" & Trim(Mid(varLine(intloop), 2, 4)) & "|" & strItem
        Else    '����
            strItemVal = Trim(Mid(varLine(intloop), 10, 5))            '���ֵ
            strItem = Trim(Mid(varLine(intloop), 8, 2))                '���
            strData = strData & "|" & Trim(Mid(varLine(intloop), 2, 4)) & "|" & _
                        IIf(strItem = "", "", strItem) & "  " & strItemVal
        End If
    Next
    strResult = strTitle & Replace(Replace(strData, "*", ""), "-", "��")
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



