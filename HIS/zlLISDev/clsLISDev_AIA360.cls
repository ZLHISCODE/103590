VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AIA360"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��ѧ���������
'�ͺ�:AIA360
'����:�ձ�������ʽ����
'����:�µ���
'�û�:�����о���ҽԺ
'�޸�����:2012-05-25

Private mstrCode As String    '��������

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

    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim var_Code As Variant, int_Loop As Integer
    Dim str_Code As String 'һ��������¼
    Dim str_Line As String 'һ������
    Dim str_�걾�� As String, str_��Ŀ As String, str_��� As String, str_���� As String
    Dim str_���� As String, str_SendLine As String, str_Item As String, cTdh As String
    Dim str_�ʿ� As String, str_Tmp As String, aTmp As String, bTmp As String

    On Error GoTo hErr
    strReserved = "": strResult = "": strCmd = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    If InStr(mstrCode, Chr(&H5)) > 0 Then                                         '�յ������� 05 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&HA)) > 0 Then                                         '�յ��н����� 10 ,Ӧ�� 06
         strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&H3)) > 0 Then                                         '�յ���ʼ�� 03 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&H4)) > 0 Then                                         '�յ���ʼ�� 04 ,Ӧ�� 06
        strCmd = Chr(&H6)
    End If
    '-------------------------------------------------------------------------------------------------------------
 
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = InStr(mstrCode, Chr(&H4))
    
    If lng_S <= 0 And lng_E <= 0 Then
       strReserved = mstrCode
    End If
    
    Do While lng_S > 0 And lng_E > 0    '�յ�һ�������ļ�¼,��ʼ����
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        '--------------------------------------------------------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If (InStr(str_Line, Chr(&H17)) > 0) And (InStr(str_Line, Chr(&H2)) > 0) Then '�����������
               aTmp = Mid(str_Line, 1, InStr(str_Line, Chr(&H17)) - 1)
               bTmp = Mid(str_Line, InStr(str_Line, Chr(&H2)) + 2)
               str_Line = aTmp + bTmp
            End If
            str_Line = Replace(str_Line, Chr(&H2), "")
            
            If Mid(str_Line, 2, 1) = "O" Then
               str_Tmp = Trim(Split(str_Line, "|")(2))
               str_Tmp = Val(str_Tmp)
               
               str_�ʿ� = "0"
               str_�ʿ� = IIf(Trim(Split(str_Line, "|")(11)) = "Q" Or Trim(Split(str_Line, "|")(11)) = "X\Q", "1", "0")
               
               If str_�걾�� <> str_Tmp Then
                    If str_Item <> "" And Val(str_�걾��) > 0 And str_���� <> "" Then
                        strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & str_�ʿ� & str_Item
                    End If
                    str_�걾�� = str_Tmp
                    str_Item = ""
               End If
            ElseIf Mid(str_Line, 2, 1) = "R" Then
               If UBound(Split(str_Line, "|")) >= 3 Then
                 str_���� = Split(str_Line, "|")(12)
                 If str_���� <> "" Then
                    str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2) & ":" & Mid(str_����, 13, 2)
                    If Not IsDate(str_����) Then str_���� = ""
                  End If
                  str_��Ŀ = Split(str_Line, "|")(2)
                  str_��Ŀ = Trim(Replace(str_��Ŀ, "^", ""))
                  str_��� = Trim(Split(str_Line, "|")(3))
                  If str_��� <> "" Then
                    str_Item = str_Item & "|" & Replace(str_��Ŀ, "^", "") & "|" & str_���
                  End If
               End If
            End If
        Next
        If str_Item <> "" And Val(str_�걾��) > 0 And str_���� <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & str_�ʿ� & str_Item
        End If
        lng_S = InStr(mstrCode, Chr(&H5))
        lng_E = InStr(mstrCode, Chr(&H4))
    Loop
    '--------------------------------------------------------------------------------------------------------------
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
hErr:
    WriteLog "AIA360", strSource, Err.Number & " " & Err.Description
End Sub


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
    'iSendStep����ǰ���Ͳ����Ĳ��衣�����������ͽ������򷵻�0�������Զ���1���ء�
    'strResponse����ǰ��������Ӧ����
    'strSampleInfo�����͵ı걾��Ϣ��
    '   �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ���
    '   ��5�����أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��

    Dim var_�걾 As Variant, i_Count As Integer, i_Loop As Integer, str_��¼ As String, str_�걾�� As String
    Dim i_���� As Integer, str_���� As String, str_������Ŀ As String, str_�걾���� As String
    Dim str_SendItem As String, str_Line As String
    
    SendSample = ""

    
End Function
