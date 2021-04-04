VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_hsCrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:hsCRP
'����:δ֪
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
        Dim lng_S As Long, lng_E As Long
        Dim str_Code As String, str_Date As String, str_Value As String
        Dim str_NO As String, var_Code As Variant, int_Loop As Integer, str_Line As String
        On Error GoTo hErr
    
100     mstrCode = mstrCode & strSource
102     lng_S = InStr(mstrCode, "hsCRP Test   No:")
104     lng_E = 0
106     If lng_S > 0 Then
108         lng_E = InStr(lng_S + 6, mstrCode, "hsCRP")
110         If lng_E > 0 Then
112             lng_E = InStr(lng_E + 6, mstrCode, "CRP")
114             If lng_E > 0 Then
116                 lng_E = InStr(lng_E, mstrCode, Chr(&HD) & Chr(&HA))
                Else
118                 lng_E = 0
                End If
            Else
120             lng_E = 0
            End If
        End If
    
122     If lng_S > 0 And lng_E > 0 Then
124         str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
126         mstrCode = Mid(mstrCode, lng_E + 1)
        
128         var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
            str_NO = ""
            str_Date = ""
            str_Value = ""
130         For int_Loop = LBound(var_Code) To UBound(var_Code)
132             str_Line = var_Code(int_Loop)
134             If str_Line Like "hsCRP Test   No:*" Then
136                 str_NO = Trim(Mid(str_Line, 17))
138             ElseIf str_Line Like "Date:*/*" Then
140                 str_Date = Mid(str_Line, 6)
142                 str_Date = Mid(str_Date, 8, 4) & "-" & TrunMonth(Mid(str_Date, 4, 3)) & "-" & Mid(str_Date, 1, 2) & " " & Mid(str_Date, 13, 4)
144             ElseIf str_Line Like "hsCRP*mg/L*" Then
146                 str_Value = Split(Trim(Mid(str_Line, 7)), " ")(0)
                    If str_Value <> "" Then
                        str_Value = "|hsCRP|" & str_Value
                    End If
148             ElseIf str_Line Like "CRP*mg/L*" Then
                    If Split(Trim(Mid(str_Line, 7)), " ")(0) <> "" Then
150                     str_Value = str_Value & "|CRP|" & Split(Trim(Mid(str_Line, 7)), " ")(0)
                    End If
                End If
            Next
152         If Val(str_NO) > 0 And IsDate(str_Date) And str_Value <> "" Then
154             strResult = str_Date & "|" & str_NO & "| |Ѫ��|0" & str_Value
            Else
156             WriteLog "hsCrp", "���ݸ�ʽ������" & str_NO & "," & str_Date, str_Value
            End If
        End If
    
        Exit Sub
hErr:
158     WriteLog "hsCRP", CStr(Erl()) & "��," & Err.Description & vbCrLf & mstrCode, strResult
End Sub

Private Function TrunMonth(ByVal strMonth As String) As String
    'ת��Ӣ�ĵ��·�Ϊ����
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select

End Function
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
