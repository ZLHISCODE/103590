VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Dymind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'��Ʒ����: Dx-CRP Ѫϸ��������
'��Ʒ�ͺ�: Dx-CRP
'����:�����е������＼�����޹�˾
'����:�¶�
'����:2016-01-05

'�û�:����ҽԺ

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
    
        Dim str_Segment As String   'һ����Ϣ��(��chr(&HD)�ָ�)
        Dim var_Filed As Variant    '�ֶ�
        Dim str_Ψһ��ʶ As String  '���ڱ�ʶһ����Ϣ
        Dim str_�汾 As String
    
        Dim str_��Ϣ���� As String  '���� Ӧ��
        Dim str_�����ͺ� As String
        Dim str_��Ϣ��ʶ  As String
    
        Dim lng_S As Long
        Dim lng_E As Long
    
        Dim str_�걾�� As String, str_���� As String
        Dim str_�ʿ� As String
        Dim str_��Ŀ As String
        Dim strItem As String
        Dim strValue As String
        
        On Error GoTo errHandle
100     strResult = ""
102     If strSource <> "" Then mstrCode = mstrCode & strSource
 
104     lng_S = InStr(mstrCode, Chr(&HB)): lng_E = 0
106     If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&HD) & Chr(&H1C))
        
108     Do While lng_S > 0 And lng_E > 0
110         str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
112         mstrCode = Mid(mstrCode, lng_E + 2)
                 
114         var_Code = Split(str_Code, Chr(&HD))
116         For lng_count = LBound(var_Code) To UBound(var_Code)
118             str_Segment = var_Code(lng_count)
120             If Mid(str_Segment, 2, 3) = "MSH" Then
                    '��Ϣͷ
122                 str_�ֶηָ��� = Mid(str_Segment, 5, 1)
124                 str_����ָ��� = Mid(str_Segment, 6, 1)
126                 str_��������� = Mid(str_Segment, 7, 1)
128                 str_�ظ��ָ��� = Mid(str_Segment, 8, 1)
130                 str_ת���ַ� = Mid(str_Segment, 9, 1)
                
132                 var_Filed = Split(str_Segment, str_�ֶηָ���)
134                 str_Ψһ��ʶ = var_Filed(9)
136                 str_�ʿ� = var_Filed(10) 'P-�걾��Q-�ʿ�
138                 str_�汾 = var_Filed(11)
                
140                 strCmd = "MSH|^~\&|LIS|ZLHIS||||" & Format(Now, "yyyyMMddHHmmss") & "||ACK^R01|HIS" & Format(Now, "yyyyMMddHHmmss") & "|P|2.3.1||||||UNICODE" & Chr(&HD)
142                 strCmd = strCmd & "MSA|AA|" & str_Ψһ��ʶ & Chr(&HD) & Chr(&H1C) & Chr(&HD)
                    
'144                 If str_�汾 <> "2.3.1" Then WriteLog "�汾����2.3.1"
                
146                 str_��Ŀ = "": str_�걾�� = "": str_���� = ""
148             ElseIf Mid(str_Segment, 1, 3) = "OBR" Then
                    '�걾�ţ�����
150                 var_Filed = Split(str_Segment, str_�ֶηָ���)
152                 str_�걾�� = var_Filed(3)
154                 str_���� = var_Filed(7)
156                 str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2)
                
158                 If str_�ʿ� = "Q" Then
                        '�ʿ�
160                     str_��Ŀ = "|" & var_Filed(3) & "|" & var_Filed(20)       '��Ŀ���,���
162                     str_�걾�� = var_Filed(12)  '�ʿ�Һ���
 
                        
                    End If
164             ElseIf Mid(str_Segment, 1, 3) = "OBX" Then
166                 var_Filed = Split(str_Segment, str_�ֶηָ���)
                     
168                 If var_Filed(2) = "NM" And var_Filed(5) <> "" Then
170                     strItem = var_Filed(3)
172                     strItem = Trim(Split(strItem, "^")(1))
174                     strValue = Trim$(var_Filed(5))
176                     If strItem <> "" And strValue <> "" Then str_��Ŀ = str_��Ŀ & "|" & strItem & "|" & strValue
                    End If
                End If
            Next
        
 
178         If Val(str_�걾��) > 0 And str_��Ŀ <> "" Then
180             strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Boold|" & IIf(str_�ʿ� = "2", 1, 0) & str_��Ŀ
            End If
      
182         str_��Ŀ = ""
184         lng_S = InStr(mstrCode, Chr(&HB)): lng_E = 0
186         If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&HD) & Chr(&H1C))
        Loop
188     If strResult <> "" Then strResult = Mid(strResult, 3)
        Exit Sub
errHandle:
'190     WriteLog "Dy " & Erl() & "��, " & Err.Description
End Sub


Public Function SendSample(strSampleInfo As String, _
                           iSendStep As Integer, _
                           blnSuccess As Boolean, _
                           Optional ByVal strResponse As String = "", _
                           Optional ByVal blnUndo As Boolean = False, _
                           Optional ByVal iType As Integer = 0) As String
        '������Ҫ���������͵ı걾��Ϣ
        '   Ԫ��֮����|�ָ�
        '   ��0��Ԫ�أ�����ʱ��
        '   ��1��Ԫ�أ��������
        '   ��2��Ԫ�أ�������
        '   ��3��Ԫ�أ��걾
        '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
        '   ��5��9Ԫ�أ�ϵͳ����
        '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    

End Function




