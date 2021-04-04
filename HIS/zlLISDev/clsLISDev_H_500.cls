VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_H_500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:H-500
'����:�������
'����:�¶�
'�û�:��������ҽԺ

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
    
    Dim varline As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String, iCount As Integer, lngLength As Long
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, Chr(&H3&)) > 0 Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            lngLength = InStr(strReserved, Chr(&H3)) - InStr(strReserved, Chr(&H2))
            strLine = Mid(strReserved, InStr(strReserved, Chr(&H2)), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H3)) + 1)
            
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
            '����
            iCount = 0
            varline = Split(strLine, vbCrLf)
            strItem = ""
            For iRow = 0 To UBound(varline)
                
                Select Case iCount
                    Case 1 '����
                       
                        str���� = Trim(Replace(varline(iRow), "Date:", ""))
                        If InStr(str����, ",") > 0 Then
                             '��������:  Date:08,16,2007 15:10 ��ʽ
                            str���� = Mid(str����, 7, 4) & "-" & Mid(str����, 1, 2) & "-" & Mid(str����, 4, 2)
                        End If
                    Case 2 '�걾��
                        If InStr(varline(iRow), ":") > 0 Then varline(iRow) = Replace(varline(iRow), ":", "")
                        str�걾�� = Val(Replace(varline(iRow), "No.", ""))
                        str�����˱걾�ʿ� = " |��Һ|0"
                    Case 3 To 13
                        'If blnReturn(MidUni(varLine(iRow), 8)) Then
                            If Trim(MidUni(varline(iRow), 5)) <> "" And Trim(MidUni(varline(iRow), 2, 3)) <> "" Then
                                strItem = strItem & "|" & Trim(MidUni(varline(iRow), 2, 3)) & "|" & Trim(MidUni(varline(iRow), 5))
                            End If
                        'End If

                End Select
                If varline(iRow) <> "" Then iCount = iCount + 1
            Next
            
            If strItem <> "" And Val(str�걾��) <> 0 Then
                strItem = Replace(strItem, "umol/L", "")
                strItem = Replace(strItem, "mmol/L", "")
                strItem = Replace(strItem, "Ery/uL", "")
                strItem = Replace(strItem, "g/L", "")
                strItem = Replace(strItem, "Leu/uL", "")
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
            End If
            If InStr(strReserved, Chr(&H3&)) = 0 Then
                blnEnd = True
            End If
        Loop
    Else
        '�յ�����Ϣ��ȫ,����,���´��ٽ���
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
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


Private Function MidUni(ByVal strTemp As String, ByVal Start As Long, Optional ByVal Length As Long) As String
'���ܣ������ݿ����õ��ַ������Ӽ���Ҳ���Ǻ��ְ������ַ��㣬����ĸ����һ��
    If Length = 0 Then Length = LenB(strTemp)
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    'ȥ�����ܳ��ֵİ���ַ�
    MidUni = Replace(MidUni, Chr(0), "")
End Function

Private Function blnReturn(ByVal strValues As String) As Boolean
    '�Ƿ񷵻ؼ����������δ�����򲻷���
    blnReturn = True
    If Val(strValues) = 0 Then
        If InStr(strValues, "0") <= 0 Then
            blnReturn = False
        End If
    End If
End Function
