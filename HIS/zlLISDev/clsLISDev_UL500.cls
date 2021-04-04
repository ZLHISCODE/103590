VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UL500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ULT500
'����:������
'�û�: ������ҽԺ

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    'GetStartCmd = Chr(6)
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
    '   ͼ�����ݷ��ڼ��������ûس����з�(vbcrlf)�ָ�
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    
    Dim varLine As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim irow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String, lngLength As Long, iReturn As Integer
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, vbCrLf) <= 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    varLine = Split(strSource, vbCrLf)
    strItem = ""
    For irow = LBound(varLine) To UBound(varLine)
        If IsDate(varLine(irow)) Then
            If UBound(varLine) > irow + 16 Then
                str���� = varLine(irow)
                str�걾�� = Val(varLine(irow + 2))
                str�����˱걾�ʿ� = " |��Һ|0"
                For iReturn = irow + 4 To irow + 17
                    'If InStr(Mid(varLine(iReturn), InStr(varLine(iReturn), vbTab) + 1), "-") <= 0 Then
                    str��Ŀ = Trim(Mid(varLine(iReturn), 1, InStr(varLine(iReturn), vbTab) - 1))
                    strValues = Trim(Mid(varLine(iReturn), InStr(varLine(iReturn), vbTab) + 1))
                    
                    If UCase(strValues) = "NEG." Then
                        strValues = "-"
                    ElseIf UCase(strValues) = "norm." Then
                        strValues = "norm"
                    ElseIf strValues = "-" Then
                        strValues = "-"
                    Else
                        strValues = Split(strValues, " ")(0)
                    End If
                    strItem = strItem & "|" & str��Ŀ & "|" & strValues
                    'End If)
                Next
                If UBound(varLine) > irow + 17 Then
                    For iReturn = irow + 18 To UBound(varLine)
                        If iReturn < UBound(varLine) Then
                            strReserved = strReserved & varLine(iReturn) & vbCrLf
                        Else
                            If varLine(iReturn) <> "" Then
                                strReserved = strReserved & varLine(iReturn)
                            End If
                        End If
                    Next
                End If
                Exit For
            Else
                strReserved = strSource
                Exit Sub
            End If
            
        End If
    Next
    strItem = Replace(strItem, "mmol/L", "")
    strItem = Replace(strItem, "Cell/uL", "")
    strItem = Replace(strItem, "umol/L", "")
    strItem = Replace(strItem, "g/L", "")
    strItem = Replace(strItem, "PRo", "PRO")
    strItem = Replace(strItem, "URo", "URO")
    If Val(str�걾��) > 0 And strItem <> "" Then
        strResult = str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    'strCmd = Chr(6)
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








