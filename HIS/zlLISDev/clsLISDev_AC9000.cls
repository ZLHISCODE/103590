VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AC9000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AC9000 ȫ�Զ�����ʷ�����
'����:�򽭰µϿ�ҽ�������������ι�˾
'����:�¶�
'�û�:������ҽԺ��̫ԭ��
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

        Dim lng_S As Long, lng_E As Long, str_Line As String, var_Tmp As Variant, iLoop As Integer
        Dim str_Date As String, str_No As String, str_Items As String, lng_K As Long
        On Error GoTo hErr
        '��ʼ������ֵ
100     strResult = "": strReserved = "": strCmd = ""
    
102     mstrCode = mstrCode & strSource
104     lng_S = InStr(mstrCode, Chr(&H2))
106     If lng_S > 0 Then
108         lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
110         lng_E = InStr(mstrCode, Chr(&H3))
        End If
    
112     Do While lng_S > 0 And lng_E > 0
114         str_Line = Mid(mstrCode, lng_S, lng_E - lng_S)
116         mstrCode = Mid(mstrCode, lng_E + 1)
            '--- ����
118         str_Date = Mid(str_Line, 2, 19) ' ����
119         lng_K = InStr(22, str_Line, "K")
120         str_No = Trim(Mid(str_Line, 22, lng_K - 22)) ' �걾��
122         str_Line = Mid(str_Line, lng_K)    ' ��Ŀ
124         str_Items = ""
126         var_Tmp = Split(str_Line, "*")
128         For iLoop = LBound(var_Tmp) To UBound(var_Tmp)
130             If Len(var_Tmp(iLoop)) > 4 Then
132                 If Trim(Mid(var_Tmp(iLoop), 1, 4)) <> "" And Trim(Mid(var_Tmp(iLoop), 5)) <> "" Then
134                     str_Items = str_Items & "|" & Trim(Mid(var_Tmp(iLoop), 1, 4)) & "|" & Trim(Mid(var_Tmp(iLoop), 5))
                    End If
                End If
            Next
        
136         If IsDate(str_Date) And Val(str_No) > 0 And str_Items <> "" Then
138             strResult = strResult & "||" & str_Date & "|" & str_No & "| |Ѫ��|0" & str_Items
            End If
                
            '-- ����Ƿ�����������
140         lng_S = InStr(mstrCode, Chr(&H2))
142         If lng_S > 0 Then
144             lng_E = InStr(lng_S, mstrCode, Chr(&H3))
            Else
146             lng_E = InStr(mstrCode, Chr(&H3))
            End If
        Loop
148     If strResult <> "" Then
150         If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
        End If
        Exit Sub
hErr:
152     WriteLog "Ac9000 ", CStr(Erl()) & ":" & Err.Description, strSource
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





