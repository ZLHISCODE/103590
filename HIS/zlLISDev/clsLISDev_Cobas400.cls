VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Cobas400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Cobas400
'����:����
'����:�¶�
'�û�:��ú������ҽԺ
Private mstrCode As String '��������

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = Chr(&H1) & Chr(&HA) & "14 COBAS INTEGRA400 09" & Chr(&HA) & Chr(&H2) & Chr(&HA) & "10  7" & Chr(&HA) & Chr(&H3) & Chr(&HA) & Chr(&H4) & Chr(&HA)
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
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant, str_INTEG As String, int_Loop As Integer
    Dim str_Line As String
    Dim str_�걾�� As String, str_���� As String, str_��Ŀ As String, str_��� As String
    
    
    'str_INTEG = Chr(&H1) & Chr(&HA) & "14 COBAS INTEGRA400 00" & Chr(&HA) & Chr(&H2) & Chr(&HA) & Chr(&H3) & Chr(&HA) & "1" & Chr(&HA) & "428" & Chr(&HA) & Chr(&H4) & Chr(&HA)
    
    strResult = ""
    mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, Chr(&H1) & Chr(&HA))
    lng_E = InStr(mstrCode, Chr(&H4) & Chr(&HA))
    
    Do While lng_S > 0 And lng_E > 0
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        
        mstrCode = Mid(mstrCode, lng_E + 2)
        lng_S = InStr(mstrCode, Chr(&H1) & Chr(&HA))
        lng_E = InStr(mstrCode, Chr(&H4) & Chr(&HA))
        
        '���ڱ���ͨѶ����
        Call Pause(1) '��ʱһ����Ӧ��
        strCmd = Chr(&H1) & Chr(&HA) & "14 COBAS INTEGRA400 09" & Chr(&HA) & Chr(&H2) & Chr(&HA) & "10  7" & Chr(&HA) & Chr(&H3) & Chr(&HA) & Chr(&H4) & Chr(&HA)
    
        str_�걾�� = "": str_���� = "": str_��Ŀ = "": str_��� = ""
        var_Code = Split(str_Code, Chr(&HA))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If Left(str_Line, 3) = "53 " Then
                str_�걾�� = Trim(Mid(str_Line, 4, 5))
                str_���� = Mid(str_Line, 26, 4) & "-" & Mid(str_Line, 23, 2) & "-" & Mid(str_Line, 20, 2)
                
            ElseIf Left(str_Line, 3) = "55 " Then
                str_��Ŀ = Trim(Mid(str_Line, 4))
            ElseIf Left(str_Line, 3) = "00 " Then
                If InStr(Trim(Mid(str_Line, 4, 14)), ">") > 0 And InStr(Trim(Mid(str_Line, 4, 14)), "<") > 0 Then
                    str_��� = Trim(Mid(str_Line, 4, 14))
                Else
                    str_��� = Val(Trim(Mid(str_Line, 4, 14)))
                End If
            End If
        Next
        If str_���� = "" Then str_���� = Format(Now, "yyyy-MM-dd")
        If Val(str_�걾��) > 0 And IsDate(CDate(str_����)) And str_��Ŀ <> "" And str_��� <> "" Then
            strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |ѪҺ|0|" & str_��Ŀ & "|" & str_���
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   �걾֮����||�ָ�
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
