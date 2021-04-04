VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_OP155"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:OP-155 ��Һ������
'����:������ŷ�����Ƽ����޹�˾
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
        Dim str_ʱ�� As String, str_������ As String, str_��Ŀ As String, str_��� As String
        Dim lng_S As Long, lng_E As Long, str_Tmp As String
        Dim str_Data As String
        Dim varline As Variant
        Dim i As Integer, str_Line As String
        On Error GoTo errH
    
100     mstrCode = mstrCode & strSource
        
102     lng_S = InStr(mstrCode, "-")
104     lng_S = lng_S - 4
106     If lng_S > 0 Then
108         If IsNumeric(Mid(mstrCode, lng_S, 4)) Then
110             lng_E = InStr(lng_S, mstrCode, "VC")
112             lng_E = InStr(lng_E, mstrCode, Chr(&HD) & Chr(&HA))
            Else
114             lng_S = 0
            End If
        End If
    
116     strResult = ""
118     Do While lng_S > 0 And lng_E > 0
            
120         str_Data = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
122         mstrCode = Mid(mstrCode, lng_E + 2)
124         str_��Ŀ = ""
126         str_ʱ�� = ""
128         varline = Split(str_Data, Chr(&HD) & Chr(&HA))
130         For i = LBound(varline) To UBound(varline)
132             str_Line = varline(i)
134             If str_Line Like "####-##-##*" Then
136                 str_ʱ�� = Format(CDate(str_Line), "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
138             ElseIf UCase(str_Line) Like "###*" Then
140                 str_������ = Val(str_Line)

                
142             ElseIf Not (UCase(str_Line) Like "*:*") Then
144                     str_Tmp = Trim(Mid(str_Line, 6))
146                     If str_Tmp <> "" And Trim(Mid(str_Line, 1, 3)) <> "" Then
148                         str_��Ŀ = str_��Ŀ & "|" & Trim(Mid(str_Line, 1, 3)) & "|" & Trim(str_Tmp)
                        End If
                End If
            
            Next
150         If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
152             strResult = str_ʱ�� & "|" & str_������ & "|" & " |" & "��Һ" & "|0" & str_��Ŀ
            End If
            
154         lng_S = InStr(mstrCode, "-")
156         lng_S = lng_S - 4
158         If lng_S > 0 Then
160             If IsNumeric(Mid(mstrCode, lng_S, 4)) Then
162                 lng_E = InStr(lng_S, mstrCode, "VC")
164                 lng_E = InStr(lng_E, mstrCode, Chr(&HD) & Chr(&HA))
                Else
166                 lng_S = 0
                End If
            End If
            
        Loop
        Exit Sub
errH:
168     WriteLog "op155", CStr(Erl()) & ":" & Err.Description, ""
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





