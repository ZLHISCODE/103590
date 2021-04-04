VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LX20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'����:�ٴ���������ϵ DXC800/LX20
'����:��������������� Beckman coulter
'����:�¶�
'�û�:��̨ɽҽԺ

Private mstrCode As String    '��������
Private mlngStep_In As Long   '�����ַ�����������
Private mstrACK As String     '�Ƿ�Ӧ��ACK,����Ӧ��ACK(06),��-Ӧ�� ETX(03)

Private mlngStep_Out As Long  '��˫����Ϣ����
Private mstr������ As String

Private mstrWaitSend As String '�����ͱ걾

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
        '   ��1��Ԫ�أ��������
        '   ��2��Ԫ�أ�������
        '   ��3��Ԫ�أ��걾
        '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
        '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
        '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
        'strReserved���������ɽ�����ԭʼ������Ҫ��������������
        'strCmd�������Ҫ���ɷ������豸���͵�����
        Dim lng_S As Long, lng_E As Long, i_Count As Integer
        Dim str_in As String, str_Line As String, str_���� As String
        Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String, str_��� As String
        Dim str_WaitSendLine As String
    
100     str_in = strSource
102     If InStr(str_in, Chr(&H4) & Chr(&H1)) > 0 Then       'EOT SOH '����ͨѶ���� Ӧ��
104         lng_S = InStr(str_in, Chr(&H4) & Chr(&H1))
106         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 2)
108         mstrACK = Chr(&H3)  '04 01�̶��̶���06��ʼӦ��
110         strCmd = ACKChar
        
112         WriteLog "�յ���ʼ(EOT SOH)����", "", str_in

114     ElseIf InStr(str_in, Chr(&H4)) > 0 Then                 'EOT '����ͨѶ����
116         lng_S = InStr(str_in, Chr(&H4))
118         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
        
120         mlngStep_In = 0
122         Call ACKChar
124         If mstrWaitSend <> "" Then
126             strCmd = "1|"
128             strResult = Format(Now, "YYYY-MM-DD") & "|^^0"
130             WriteLog "�յ�����(EOT)����.cmd=1", str_in, mstrWaitSend
            Else
132             WriteLog "�յ�����(EOT)����.", str_in, ""
            End If
        End If
    
134     If InStr(str_in, Chr(&H3)) > 0 Then                 '�յ�ETX(03)
136         lng_S = InStr(str_in, Chr(&H3))
138         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
140         mlngStep_Out = mlngStep_Out + 1
        End If
    
142     If InStr(str_in, Chr(&H6)) > 0 Then                 '�յ�ACK(06)
144         lng_S = InStr(str_in, Chr(&H6))
146         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
148         mlngStep_Out = mlngStep_Out + 1
        End If
    
150     If InStr(str_in, Chr(&H5)) > 0 Then                 '�յ�ENQ(05) Ӧ�� ETX(03)
152         lng_S = InStr(str_in, Chr(&H5))
154         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
156         mstrACK = Chr(&H6)
158         strCmd = ACKChar
        End If
    
160     If InStr(str_in, Chr(&HA)) > 0 Then                 '�յ�����(0A) Ӧ��
162         lng_S = InStr(str_in, Chr(&HA))
164         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
166         strCmd = ACKChar
        End If
168     If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
170     lng_S = InStr(mstrCode, "[")
172     lng_E = InStr(mstrCode, "]")
174     If lng_S > 0 And lng_E > 0 And InStr(mstrCode, Chr(&HD)) > 0 Then
176         str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
178         lng_E = InStr(mstrCode, Chr(&HD))
180         mstrCode = Mid(mstrCode, lng_E + 1)
        
182         If InStr(str_Line, ",801,06,") > 0 Then         '˫��ͨѶ- ��ʼ
184             mstr������ = ""
186             mlngStep_Out = 1
188             For i_Count = 0 To 3
190                 str_���� = Split(str_Line, ",")(3 + i_Count)
192                 str_���� = Trim(Replace(str_����, "]", ""))
194                 If str_���� <> "" Then
196                     strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & str_����
                    End If
                Next
198             If strResult <> "" Then
200                 mstr������ = Split(str_Line, ",")(0)
202                 strResult = Mid(strResult, 3)
204                 Call ACKChar
206                 strCmd = "1|" & ACKChar      '��������ͨѶ��ʼ���� EOT SOH
                End If
208             WriteLog "�յ� 801,06 ˫��ʼ :", strCmd, mstrWaitSend
210         ElseIf InStr(str_Line, ",802,01,") > 0 Then     '������ʼ���ؽ��
212             WriteLog "�յ� 802,01 �����ʼ :", strCmd, mstrWaitSend
214         ElseIf InStr(str_Line, ",802,05,") > 0 Then     '�������ؽ�����
216             WriteLog "�յ� 802,05 �������:", strCmd, mstrWaitSend
218         ElseIf InStr(str_Line, ",801,02,") > 0 Then     '�������ؽ�����
                'Call ACKChar
220             WriteLog "�յ� 801,02 ����ȷ��", strCmd, mstrWaitSend 'strCmd = "1|" & ACKChar
222         ElseIf InStr(str_Line, ",802,03,") > 0 Then     '������
224             str_���� = Split(str_Line, ",")(3)
226             str_���� = Right(str_����, 4) & "-" & Mid(str_����, 3, 2) & "-" & Left(str_����, 2)
228             If Len(Split(str_Line, ",")(4)) >= 6 And IsNumeric(Split(str_Line, ",")(4)) Then
230                 str_���� = str_���� & " " & Mid(Split(str_Line, ",")(4), 1, 2) & ":" & Mid(Split(str_Line, ",")(4), 3, 2) & ":" & Mid(Split(str_Line, ",")(4), 5, 2)
                End If
232             str_�걾�� = Trim(Split(str_Line, ",")(9))
234             str_��Ŀ = Trim(Split(str_Line, ",")(10))
236             str_��� = Trim(Split(str_Line, ",")(15))
238             If Len(Trim(str_�걾��)) >= 12 Then str_�걾�� = "^^" & str_�걾��
240             strResult = str_���� & "|" & str_�걾�� & "| |Ѫ��|0|" & str_��Ŀ & "|" & str_���
            End If
        End If
242     If strCmd <> "" Then WriteLog "����ָ��:", strCmd, mstrWaitSend
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
        Dim str_SendItem As String, str_WaitSendLine As String, str_Line As String, str_tmp As String
        On Error GoTo hErr
100     blnSuccess = True
102     If iSendStep = 0 Then
104         var_�걾 = Split(strSampleInfo, "||")
106         WriteLog "SendSample(If begin)", strSampleInfo, strResponse
108         SendSample = ""
110         For i_Count = LBound(var_�걾) To UBound(var_�걾)
112             str_��¼ = var_�걾(i_Count)
            
114             If Val(Split(str_��¼, "|")(1)) = 0 Then
116                 str_�걾�� = Trim(Split(str_��¼, "|")(5))
                Else
118                 str_�걾�� = Val(Split(str_��¼, "|")(1))
                End If
120             str_�걾���� = Split(str_��¼, "|")(3)
                'SE Ѫ��,SF �Լ�Һ,UR �����,TU ��ʱ��,PL Ѫ��,OT ����,BL ѪҺ, Ѫ

122             If InStr(str_�걾����, "Ѫ��") > 0 Then
124                 str_�걾���� = "PL"
                ElseIf InStr(str_�걾����, "����Ѫ") > 0 Then
                    str_�걾���� = "SE"
126             ElseIf InStr(str_�걾����, "Ѫ��") > 0 Then
128                 str_�걾���� = "SE"
130             ElseIf InStr(str_�걾����, "Ѫ") > 0 Then
132                 str_�걾���� = "BL"
134             ElseIf InStr(str_�걾����, "�Լ�Һ") > 0 Then
136                 str_�걾���� = "SF"
138             ElseIf InStr(str_�걾����, "��ʱ��") > 0 Then
140                 str_�걾���� = "TU"
142             ElseIf InStr(str_�걾����, "��") > 0 Then
144                 str_�걾���� = "UR"
                Else
146                 str_�걾���� = "SE"
                End If
148             i_���� = Val(Split(str_��¼, "|")(4))
150             str_���� = Trim(Split(str_��¼, "|")(5))
152             str_������Ŀ = ""
154             str_��¼ = Replace(str_��¼, "|NGSP HBA1", "") 'ȥ����Ժ�ļ�����Ŀ
156             For i_Loop = 10 To UBound(Split(str_��¼, "|"))
158                 str_tmp = Trim(Split(str_��¼, "|")(i_Loop))
160                 If Len(str_tmp) <= 3 Then
162                     str_tmp = str_tmp & Space(4 - Len(str_tmp))
                    End If
                    
164                 str_������Ŀ = str_������Ŀ & "," & str_tmp & ",0"
                    
                Next
            
            
166             If str_������Ŀ <> "" And str_���� <> "" Then
168                 str_������Ŀ = Format(i_Loop - 10, "000") & str_������Ŀ
                
170                 str_SendItem = mstr������ & ",801,01,0000,00,0"
172                 str_SendItem = str_SendItem & "," & IIf(i_���� = 1, "ST", "RO")             'ST-����,RO-����
174                 str_SendItem = str_SendItem & "," & str_�걾����                                         'Ѫ�� SE,SF ��Һ,UR ��(���),,
176                 str_SendItem = str_SendItem & "," & str_���� & Space(15 - Len(str_����))    '�����,15λ,����15λ���ո�
178                 str_SendItem = str_SendItem & "," & Space(20)
180                 str_SendItem = str_SendItem & "," & Space(12)
182                 str_SendItem = str_SendItem & "," & Space(25)
184                 str_SendItem = str_SendItem & "," & Space(18)
186                 str_SendItem = str_SendItem & "," & Space(15)
188                 str_SendItem = str_SendItem & "," & Space(1)
190                 str_SendItem = str_SendItem & "," & str_�걾�� & Space(15 - Len(str_�걾��)) '�걾��
192                 str_SendItem = str_SendItem & "," & Space(18)
194                 str_SendItem = str_SendItem & "," & Space(8)
196                 str_SendItem = str_SendItem & "," & Space(4)
198                 str_SendItem = str_SendItem & ",ZLSOFT" & Space(14)                         'λ��
200                 str_SendItem = str_SendItem & "," & Space(3)                                '����
202                 str_SendItem = str_SendItem & ",5"                                          '���䵥λ,Ĭ��5
204                 str_SendItem = str_SendItem & "," & Space(8)                                '����
206                 str_SendItem = str_SendItem & ",M"                                          '�Ա�
208                 str_SendItem = str_SendItem & "," & Space(45)
210                 str_SendItem = str_SendItem & ",111.1  "
212                 str_SendItem = str_SendItem & ",22.2"
214                 str_SendItem = str_SendItem & ",33.3"
216                 str_SendItem = str_SendItem & ",03"
218                 str_SendItem = str_SendItem & ",4.4444"
220                 str_SendItem = str_SendItem & "," & str_������Ŀ & "]"
222                 str_SendItem = str_SendItem & GetCheckCode(str_SendItem) & vbCrLf
                    
224                 mstrWaitSend = mstrWaitSend & "|" & Chr(&H4) & Chr(&H1) & "|" & str_SendItem & "|" & Chr(&H4)
                    
                End If
            Next
            
226         If mstrWaitSend <> "" Then
228             str_Line = strResponse
230             If InStr(str_Line, Chr(&H4) & Chr(&H1)) > 0 Then
232                 SendSample = ""
234                 iSendStep = 0
                    Exit Function
                Else
236                 If Left(mstrWaitSend, 1) = "|" Then
238                     mstrWaitSend = Mid(mstrWaitSend, 2)
                    End If
240                 If InStr(mstrWaitSend, "|") > 0 Then
242                     SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
244                     mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|"))
                    Else
246                     SendSample = mstrWaitSend
248                     mstrWaitSend = ""
                    End If
                End If
                
250             If SendSample <> "" Then iSendStep = 1
                
            End If
252         WriteLog "SendSample(If End)��", SendSample, "״̬��" & iSendStep & vbNewLine & "������:" & mstrWaitSend
        Else
254         WriteLog "SendSample(Else Begin)", strResponse, "״̬��" & iSendStep & vbNewLine & "������:" & mstrWaitSend
256         str_Line = strResponse
258         If InStr(str_Line, Chr(&H4) & Chr(&H1)) > 0 Then
                '����ǰ���04��01 �Ա��´η���
260             If Left(mstrWaitSend, 1) = "|" Then mstrWaitSend = Mid(mstrWaitSend, 2)
262             If InStr(mstrWaitSend, "|") > 0 Then
264                 str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
                Else
266                 str_WaitSendLine = mstrWaitSend
                End If
268             If Left(str_WaitSendLine, 1) = "[" Then mstrWaitSend = Chr(&H4) & Chr(&H1) & "|" & mstrWaitSend
                
270             SendSample = ""
272             iSendStep = 0

274             WriteLog "SendSample(Else End)", strResponse, "����һ���յ����ݣ�" & iSendStep & vbNewLine & "������:" & mstrWaitSend
                Exit Function
            End If
276         str_WaitSendLine = ""
278         If InStr(str_Line, Chr(&H3)) > 0 Or _
               InStr(str_Line, Chr(&H4)) > 0 Or _
               InStr(str_Line, Chr(&H6)) > 0 Then                            '��������ȷ����Ϣ
               
280             If mstrWaitSend <> "" Then
282                 If Left(mstrWaitSend, 1) = "|" Then mstrWaitSend = Mid(mstrWaitSend, 2)
                    
284                 If InStr(mstrWaitSend, "|") > 0 Then
286                     str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
288                     mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|"))
                    Else
290                     str_WaitSendLine = mstrWaitSend
292                     mstrWaitSend = ""
                    End If
294                 SendSample = str_WaitSendLine
                End If
            End If
296         If mstrWaitSend = "" Then iSendStep = 0
298         If InStr(SendSample, Chr(&H4)) > 0 Then iSendStep = 0             '����һ���걾�����˳�˫��״̬��

300         WriteLog "SendSample(Else End)", strResponse, "״̬��" & iSendStep & vbNewLine & "������" & mstrWaitSend
        End If
        Exit Function
hErr:
302     WriteLog "SendSample", CStr(Erl()) & " " & Err.Description, "״̬:" & iSendStep & ",�յ�:" & strResponse & ",�걾��" & strSampleInfo & ",����:" & mstrWaitSend & ",���η�:" & SendSample
End Function


Private Function ACKChar() As String
    If mstrACK = Chr(&H6) Then
        mstrACK = Chr(&H3)
    Else
        mstrACK = Chr(&H6)
    End If
    ACKChar = mstrACK
End Function


Private Function GetCheckCode(ByVal str_in As String) As String
    '����У����
    Dim ls_ch As String
    Dim ll_i As Long
    Dim ll_len As Long
    Dim ll_sum As Long

    If ((IsNull(str_in)) Or (str_in = "")) Then
        GetCheckCode = ""
    End If

    ll_len = Len(str_in)

    For ll_i = 1 To ll_len
        ls_ch = Mid(str_in, ll_i, 1)
        ll_sum = ll_sum + Asc(ls_ch)
    Next

    ll_sum = 256 - (ll_sum Mod 256)
    GetCheckCode = Hex(ll_sum)
    If Len(GetCheckCode) = 1 Then GetCheckCode = "0" & GetCheckCode
    
End Function







