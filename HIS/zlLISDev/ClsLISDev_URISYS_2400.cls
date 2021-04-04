VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_URISYS_2400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:URISYS-2400 (��Һ������)
'����:����
'����:�¶�
'�û�:��ú����ҽԺ

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
        
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������^�Ƿ���^�����  ����ֻ���������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
   
    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim str_Code As String 'һ��������¼
    Dim str_Line As String 'һ������
    Dim str_�걾�� As String, str_��Ŀ As String, str_��� As String, str_���� As String
    Dim str_���� As String, str_SendLine As String
    
    
    strResult = ""
    str_in = strSource
    '-------------------------------------------------------------------------------------------------------------
    If InStr(str_in, Chr(&H5)) > 0 Then                                         '�յ���ʼ�� 05 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    If InStr(str_in, Chr(&H3)) > 0 Or InStr(str_in, Chr(&H17)) > 0 Then         '�յ��н����� 03,17 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    If mstrCode Like "*1H|*" & Chr(&H4) & "*" Then                         '�յ�һ�������ļ�¼,��ʼ����
        WriteLog "���� URISYS��2400", mstrCode, ""
        Do While InStr(mstrCode, Chr(&H4)) > 0
            lng_S = InStr(mstrCode, "1H|")
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
            
            str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            
            '���� У����
            Do While InStr(str_Code, Chr(&H17)) > 0
                lng_S = InStr(str_Code, Chr(&H17))
                str_Code = Mid(str_Code, 1, lng_S - 1) & Mid(str_Code, lng_S + 7)
            Loop
            Do While InStr(str_Code, Chr(&H3)) > 0
                lng_S = InStr(str_Code, Chr(&H3))
                str_Code = Mid(str_Code, 1, lng_S - 1) & Mid(str_Code, lng_S + 5)
            Loop
            
            Do While str_Code <> ""
                lng_S = 1
                lng_E = InStr(str_Code, Chr(&HD))
                
                If lng_S > 0 And lng_E > 0 Then
                    str_Line = Mid(str_Code, lng_S, lng_E - lng_S + 1)
                    str_Code = Mid(str_Code, lng_E + 1)
                    
                    If str_Line Like "O|*|*" Then                            '�걾��,��Ŀ ��
                        str_�걾�� = Val(Split(str_Line, "|")(3))
                        str_���� = Split(str_Line, "|")(14)
                        If str_���� <> "" Then
                            str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2)
                        End If
                    ElseIf InStr(str_Line, "1H|") > 0 Then                       '��Ŀ ��� ��
                        str_�걾�� = "": str_��Ŀ = ""
                    ElseIf str_Line Like "R|*|*" Then
                        
                        str_��� = Trim(Split(str_Line, "|")(3))
                        If str_��� <> "" Then
                            str_��Ŀ = str_��Ŀ & "|" & Replace(Split(str_Line, "|")(2), "^^^", "") & "|" & str_���
                        End If
                    End If
                Else
                    str_Code = ""
                End If
            Loop
            If Val(str_�걾��) <> 0 And IsDate(str_����) And str_��Ŀ <> "" Then
                strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |��Һ|0" & str_��Ŀ
            End If
        Loop
    End If
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        WriteLog "���� URISYS��2400 End", strResult, ""
    End If
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
        SendSample = ""
End Function






