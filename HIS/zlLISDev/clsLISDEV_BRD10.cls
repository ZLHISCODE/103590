VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDEV_BRD10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:BIO-RAD D10�ǻ�Ѫ�쵰�׷�����
'����:����
'����:�¶�
'�û�:�Ϻ�һ��ҽԺ

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
        Dim lng_S As Long, lng_E As Long, str_Code As String, int_Loop As Integer
        Dim str_Value As String, str_Date As String, str_NO As String
        Dim var_Code As Variant, str_Line As String
        On Error GoTo hErr
100     mstrCode = mstrCode & strSource
102     If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
104     If InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)
    
106     lng_S = InStr(mstrCode, Chr(&H5))
108     lng_E = 0
110     If lng_S > 0 Then
112         lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        End If
    
114     If lng_S > 0 And lng_E > 0 Then
116         str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
118         mstrCode = Mid(mstrCode, lng_E + 1)
        
120         var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
122         str_Value = ""
124         str_NO = ""
126         str_Date = ""
128         For int_Loop = LBound(var_Code) To UBound(var_Code)
130             str_Line = var_Code(int_Loop)
132             If Mid(str_Line, 3, 2) = "O|" Then
134                 str_NO = Trim(Split(str_Line, "|")(2))
136             ElseIf Mid(str_Line, 3, 2) = "R|" Then
138                 If Split(str_Line, "|")(2) = "^^^A1c^AREA" Then
140                     str_Value = Trim(Split(str_Line, "|")(3))
142                     str_Date = Trim(Split(str_Line, "|")(12))
144                     If Len(str_Date) >= 13 Then
146                         str_Date = Mid(str_Date, 1, 4) & "-" & Mid(str_Date, 5, 2) & "-" & Mid(str_Date, 7, 2) & " " & Mid(str_Date, 9, 2) & ":" & Mid(str_Date, 11, 2)
                        Else
148                         str_Date = ""
                        End If
                    End If
                End If
            Next
150         If Val(str_NO) > 0 And IsDate(str_Date) And Trim(str_Value) <> "" Then
152             strResult = str_Date & "|" & str_NO & "| |Ѫ��|0|HBA1C|" & str_Value
            Else
154              WriteLog "BRD10", "���ݸ�ʽ������" & str_NO & "," & str_Date, str_Value
            End If
        End If
        Exit Sub
hErr:
156     WriteLog "BRD10", CStr(Erl()) & "��," & Err.Description & vbCrLf & mstrCode, strResult
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
