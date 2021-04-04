VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BS_420"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: BS-420 ȫ�Զ�����������
'�ͺ�: BS-420
'����:������������ҽ�Ƶ��ӹɷ����޹�˾
'����:�µ���
'����:2012-04-17
'      2012-05-14 Ӧ��ָ��ǰ��"0|" by �¶�
Private mstrCode As String  '����������
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
    '2010-01-14|1| |Ѫ��|0|Alb|44|ALT|6|TP|72|T��Bil|10|D��Bil|2|AST|9|ALP|118|IBil|8|GLB|28|A/G|1.57

    Dim str_Code As String '1����������Ϣ
    Dim var_Code As Variant
    Dim lng_count As Long
    
    Dim str_�ֶηָ��� As String
    Dim str_����ָ��� As String
    Dim str_��������� As String
    Dim str_�ظ��ָ��� As String
    Dim str_ת���ַ� As String
    
    Dim str_Segment As String 'һ����Ϣ��(��chr(&HD)�ָ�)
    Dim var_Filed As Variant   '�ֶ�
    Dim str_Ψһ��ʶ As String '���ڱ�ʶһ����Ϣ
    Dim str_�汾 As String
    
    Dim str_��Ϣ���� As String '����BS-380��R01Ӧ��
    Dim str_�����ͺ� As String
    Dim str_��Ϣ��ʶ  As String
    
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_�걾�� As String, str_���� As String, str_������� As String
    Dim str_�ʿ� As String
    Dim str_��Ŀ As String
    Dim str_��� As String
    Dim str_Return As String
    
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
 
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
                 
        var_Code = Split(str_Code, Chr(&HD))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If Mid(str_Segment, 2, 3) = "MSH" Then
                '��Ϣͷ
                str_�ֶηָ��� = Mid(str_Segment, 5, 1)
                str_����ָ��� = Mid(str_Segment, 6, 1)
                str_��������� = Mid(str_Segment, 7, 1)
                str_�ظ��ָ��� = Mid(str_Segment, 8, 1)
                str_ת���ַ� = Mid(str_Segment, 9, 1)
                
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_Ψһ��ʶ = var_Filed(9)
                str_�ʿ� = var_Filed(15)
                str_������� = var_Filed(15)
                str_�汾 = var_Filed(11)
                
                str_�����ͺ� = var_Filed(3)
                str_��Ϣ��ʶ = var_Filed(9)
                str_��Ϣ���� = var_Filed(10)
                If str_�汾 <> "2.3.1" Then
                    WriteLog "�汾����", "", ""
                End If
                str_��Ŀ = "": str_�걾�� = "": str_���� = ""
                
                If var_Filed(8) = "ORU^R01" Then
                    strCmd = "0|" & Chr(&HB) & "MSH|^~\&|LIS-Server|HanYuan Hospital|Mindray|" & str_�����ͺ� & "|" & Format(Now, "yyyyMMddHHmmss") & "||ACK^R01|" & str_��Ϣ��ʶ & "|P|2.3.1||||" & str_������� & "||ASCII|||" & Chr(&HD) _
                           & "MSA|AA|" & str_��Ϣ��ʶ & "|Message accepted|||0|" & Chr(&HD) _
                           & Chr(&H1C) & Chr(&HD)
                End If
            ElseIf InStr(str_Segment, "OBR|") > 0 Then
                '�걾�ţ�����
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_�걾�� = var_Filed(3)
                str_���� = var_Filed(7)
                '20120417095659
                str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2) & ":" & Mid(str_����, 13, 2)
            ElseIf InStr(str_Segment, "OBX|") > 0 Then
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_��Ŀ = str_��Ŀ & "|" & var_Filed(4) & "|" & var_Filed(5)
            End If
        Next
        
        
        If Val(str_�걾��) > 0 And str_��Ŀ <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & IIf(str_�ʿ� = "2", 1, 0) & str_��Ŀ
        End If
        str_��Ŀ = ""
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "BS420_Analyse", strSource, Err.Number & " " & Err.Description
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
