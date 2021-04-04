VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_FA100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:FA-100 ��Һ������
'����:ɽ����ɭ
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
102     lng_S = InStr(mstrCode, Chr(&H2))
104     lng_E = InStr(mstrCode, Chr(&H3))
    
106     strResult = ""
108     If lng_S > 0 And lng_E > 0 Then
110         str_Data = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
112         mstrCode = Mid(mstrCode, lng_E + 1)
114         str_��Ŀ = ""
116         str_ʱ�� = ""
118         varline = Split(str_Data, Chr(&HD) & Chr(&HA))
120         For i = LBound(varline) To UBound(varline)
122             str_Line = varline(i)
124             If UCase(str_Line) Like "*NO.*" Then
                    '--FA200��ʽ
126                 str_Line = UCase(str_Line)
128                 str_������ = Val(Mid(str_Line, InStr(str_Line, "NO.") + 7))
130             ElseIf UCase(str_Line) Like "*FA-11*" Then
                
132                 str_Line = UCase(str_Line)
134                 str_������ = Mid(str_Line, InStr(str_Line, "FA-11") + 5)
136                 If InStr(str_������, "-") > 0 Then
138                     str_������ = Val(Split(str_������, "-")(1))
                    Else
140                     str_������ = Val(str_������)
                    End If
142                 str_ʱ�� = Format(Now, "yyyy-MM-dd HH:MM")
144             ElseIf str_Line Like "####/##/##*" Then
146                 str_ʱ�� = Format(CDate(str_Line), "yyyy-MM-dd HH:MM")
                
148             ElseIf Not (UCase(str_Line) Like "*:*") Then
150                     str_Tmp = Trim(Mid(str_Line, 6))
152                     str_Tmp = Replace(UCase(str_Tmp), "MMOL/L", "")
153                     str_Tmp = Replace(UCase(str_Tmp), "CELLS/UL", "")
154                     str_Tmp = Replace(UCase(str_Tmp), "CELL/UL", "")
156                     str_Tmp = Replace(UCase(str_Tmp), "UMOL/L", "")
158                     str_Tmp = Replace(UCase(str_Tmp), "MG/L", "")
160                     str_Tmp = Replace(UCase(str_Tmp), "LEU/UL", "")
162                     str_Tmp = Replace(UCase(str_Tmp), "G/L", "")
                        
164                     'If str_Line Like "*+*" Then
166                     '    str_Tmp = Trim(Mid(str_Line, 6, InStr(1, Mid(str_Line, 6), "+")))
                        'End If
168                     If str_Tmp <> "" And Trim(Mid(str_Line, 2, 4)) <> "" Then
170                         str_��Ŀ = str_��Ŀ & "|" & Replace(Trim(Mid(str_Line, 2, 4)), "*", "") & "|" & Trim(str_Tmp)
                        End If
                End If
            
            Next
172         If Val(str_������) > 0 And IsDate(str_ʱ��) And str_��Ŀ <> "" Then
174             strResult = str_ʱ�� & "|" & str_������ & "|" & " |" & "��Һ" & "|0" & str_��Ŀ
            End If
        End If
        Exit Sub
errH:
176     WriteLog "FA100", CStr(Erl()) & ":" & Err.Description, ""
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

