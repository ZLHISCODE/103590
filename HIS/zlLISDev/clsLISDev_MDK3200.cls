VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MDK3200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'��Ʒ����: Ѫ����MDK3200
'��Ʒ�ͺ�: MDK3200
'����:���
'����:2010-09-15
'�û�:�˹�ҵ��һ��ҽԺ�����»�
'ע������ýӿڱ�������ܹ�˾ ReadTxtSend ����ʹ��

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

    Dim str_Code As String '1����������Ϣ
    Dim var_Code As Variant
    Dim var_Line As Variant
    Dim i As Integer, j As Integer
    Dim var_Filed As Variant   '�ֶ�
    Dim lng_S As Long, lng_E As Long
    Dim str_�걾�� As String, str_���� As String, str_�걾���� As String
    Dim str_�ʿ� As String, str_������Ŀ As String
    Dim str_Line As String, str_Data As String, str_Tmp As String
    
    On Error GoTo errHandle

    mstrCode = mstrCode & strSource
 
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = InStr(mstrCode, Chr(&H3))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S + 1, lng_E - 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
        
        var_Code = Split(str_Code, "  " & Chr(&HD) & Chr(&HA))
        For i = LBound(var_Code) To UBound(var_Code)
            str_Data = Replace(Replace(var_Code(i), "[", ""), "]", "")
            If str_Data Like "Base*" Then
                var_Line = Split(str_Data, Chr(&HD) & Chr(&HA))
                
                For j = LBound(var_Line) + 1 To UBound(var_Line) - 1
                    str_Line = var_Line(j)
                    
                    var_Filed = Split(str_Line, "=")
                    
                    If var_Filed(0) = "TESTDATE" Then
                        str_���� = Format(var_Filed(1), "yyyy-MM-dd")
                    ElseIf var_Filed(0) = "SAMPLENO" Then
                        str_�걾�� = Val(var_Filed(1))
                    End If
                    
                Next
            ElseIf str_Data Like "DATA*" Then
                var_Line = Split(str_Data, Chr(&HD) & Chr(&HA))
                
                For j = LBound(var_Line) + 1 To UBound(var_Line) - 1
                    str_Line = var_Line(j)
                    
                    var_Filed = Split(str_Line, "=")
                    
                    If var_Filed(0) = "RESULT" Then
                        str_Tmp = Val(var_Filed(1))
                    ElseIf var_Filed(0) = "NO" Then
                        str_Tmp = Val(var_Filed(1)) & "|" & str_Tmp
                    End If
                    
                Next
                str_������Ŀ = str_������Ŀ & "|" & str_Tmp
            ElseIf str_Data Like "GRAPH*" Then
                
            End If
        Next
        
        str_�걾���� = "Ѫ��"
        
        If Val(str_�걾��) > 0 And str_������Ŀ <> "" Then
            strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |" & str_�걾���� & "|0" & str_������Ŀ
        End If
        
        str_������Ŀ = ""
        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = InStr(mstrCode, Chr(&H3))
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "MDK3200_analyse", strSource, Err.Number & " " & Err.Description
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