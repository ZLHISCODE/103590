VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_H100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:H-100
'����:�������
'����:�¶�
'�û�:������ҽԺ
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
    Dim strCode As String, varTmp As Variant, iRow As Integer, strLine As String
    Dim strDate As String, strValue As String, strItem As String, strSampleNo As String
    
    On Error GoTo errH
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    strCode = PStr_CutCode(mstrCode, Chr(&H2), Chr(&H3))
    Do While strCode <> ""
        varTmp = Split(strCode, Chr(&HD) & Chr(&HA))
        strItem = "": strDate = "": strSampleNo = ""
        For iRow = LBound(varTmp) To UBound(varTmp)
            strLine = varTmp(iRow)
            If Mid(strLine, 2) Like "Date:####-##-##*" Then
                strDate = Mid(strLine, 7)
            ElseIf Mid(strLine, 2) Like "Seq:*" Then
                strSampleNo = Trim(Replace(Mid(strLine, 6), "-", ""))
            ElseIf InStr(",UBG,BIL,KET,BLD,PRO,NIT,LEU,GLU,SG,pH,VC,", "," & Trim(Mid(strLine, 2, 3)) & ",") > 0 Then
                strValue = Trim(Mid(strLine, 5))
                strValue = Replace(strValue, "umol/L", "")
                strValue = Replace(strValue, "mmol/L", "")
                strValue = Replace(strValue, "cells/uL", "")
                strValue = Trim(Replace(strValue, "g/L", ""))
                If strValue <> "" And Trim(Mid(strLine, 2, 3)) <> "" Then
                    strItem = strItem & "|" & Trim(Mid(strLine, 2, 3)) & "|" & strValue
                End If
            End If
        Next
        If strDate <> "" And strSampleNo <> "" And strItem <> "" Then
            strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Boold|0" & strItem
        End If
        strCode = PStr_CutCode(mstrCode, Chr(&H2), Chr(&H3))
    Loop
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    Exit Sub
errH:
    WriteLog "H100", "", Err.Description
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

