VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_S8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'��Ʒ����: S8 ��������������
'��Ʒ�ͺ�: S8(XD-S8)
'����:�ɶ���˹���������޹�˾
'����:���
'����:2009-09-21

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
    
    '�������������д�������������ע�����벻Ҫɾ����
    Dim str_ʱ�� As String, str_������ As String, str_������Ŀ As String
    Dim str_�걾 As String
    Dim lng_S As Long, lng_E As Long, str_Tmp As String
    Dim str_Data As String, mstrCode As String
    Dim varline As Variant, varTemp As Variant
    Dim str_��Ŀ���� As String
    Dim i As Integer, j As Integer, str_Line As String
    If strSource <> "" Then mstrCode = mstrCode & strReserved & strSource
    
    lng_S = InStr(mstrCode, Chr(&H1B) & Chr(&H49) & Chr(&H43))
    If InStr(mstrCode, "��������") Then
        If InStr(mstrCode, "��������") + 21 < Len(mstrCode) Then    'And Mid(mstrCode, InStr(mstrCode, "��������"), 25) Like "*####/##/## ##:##*"
            lng_E = InStr(mstrCode, "��������") + 21
        Else
            lng_E = 0
        End If
    End If
    If Not (lng_S > 0 And lng_E > 0) Then
        strReserved = mstrCode
        Exit Sub
    End If
    Do While lng_S > 0 And lng_E > 0
        str_Data = Mid(mstrCode, lng_S, lng_E)
        mstrCode = Mid(mstrCode, lng_E + 1)
        str_������Ŀ = ""
        
        varline = Split(str_Data, Chr(&HD) & Chr(&HA))
        For i = LBound(varline) To UBound(varline)
            
            str_Line = varline(i)
            If str_Line Like "*��Ʒ���*" Then
                str_������ = Trim(Mid(str_Line, InStr(str_Line, "��Ʒ���") + 5, _
                            (InStr(Mid(str_Line, InStr(str_Line, "��Ʒ���") + 5), "�Ʊ�")) - 1))
                str_�걾 = "Ѫ��"         'Trim(Mid(str_Line, InStr(str_Line, "�걾����") + 5))
            ElseIf InStr(str_Line, "/L") Then
                varTemp = Split(str_Line, "/L")
                For j = LBound(varTemp) To UBound(varTemp) - 1
                    str_Tmp = Trim(varTemp(j))
                    str_��Ŀ���� = Trim(Mid(str_Tmp, 1, 7))
                    Select Case str_��Ŀ����
                        Case "������"
                            str_��Ŀ���� = "PTT"
                        Case "�ȱ�ת��ø"
                            str_��Ŀ���� = "GBZAM"
                        Case "����"
                            str_��Ŀ���� = "NS"
                        Case "�׵���"
                            str_��Ŀ���� = "BDB"
                        Case "�Ȱ���ת��ø"
                            str_��Ŀ���� = "GAXZYM"
                        Case "����"
                            str_��Ŀ���� = "JG"
                        Case "����ø"
                            str_��Ŀ���� = "DFM"
                        Case "������̼"
                            str_��Ŀ���� = "EYHT"
                    End Select
                    
                    str_������Ŀ = str_������Ŀ & "|" & str_��Ŀ���� & "|" & Val(Trim(Mid(str_Tmp, 8, 13)))
                Next
            ElseIf InStr(str_Line, "��������") > 0 Then
                str_ʱ�� = Format(Trim(Mid(str_Line, InStr(str_Line, "��������") + 5)), "yyyy-MM-dd hh:mm")
            End If
        Next
        If Val(str_������) > 0 And IsDate(str_ʱ��) And str_������Ŀ <> "" Then
            strResult = strResult & "||" & str_ʱ�� & "|" & str_������ & "|" & " |" & str_�걾 & "|0" & str_������Ŀ
        End If
        
        lng_S = InStr(mstrCode, Chr(&H1B) & Chr(&H49) & Chr(&H43))
        If InStr(mstrCode, "��������") Then
            If InStr(mstrCode, "��������") + 21 < Len(mstrCode) Then    'And Mid(mstrCode, InStr(mstrCode, "��������"), 25) Like "*####/##/## ##:##*"
                lng_E = InStr(mstrCode, "��������") + 21
            Else
                lng_E = 0
            End If
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)

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



