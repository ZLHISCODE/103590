VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_C601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�ٴ���������ϵ C601
'����:Roche ����
'����:�¶�
'�û�:
'2014-10-11 ����ҽԺ�޸�˫��

'ע�⣺ ϡ�����ϡ����Ŀ���ܿ���һ���걾�У�����ƶ����ʹƶ���ϡ��5�����ܿ���һ���걾�С�
'       HCG�������HH��LL�������͡�HHΪ����LLΪ��

Private mstrCode As String    '��������
Private mlngStep_In As Long   '�����ַ�����������

Private mlngStep_Out As Long  '��˫����Ϣ����
Private mstrRack As String    'ɨ����ʱ�����̺�,�ܺŵ���Ҫ���ص���Ϣ.

Private mstrWaitSend As String '�����ͱ걾
Private mstrQCNO As String      '��QC��ţ�ҪתΪ�̶��ı걾��

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
        Dim str_in As String, str_Line As String, str_���� As String, var_line As Variant
        Dim str_���� As String, str_�걾�� As String, str_��Ŀ As String, str_��� As String, str�����־ As String
        Dim str_WaitSendLine As String, str_���� As String, str_�ʿ� As String
        '-----------------------------------------------------------------------------
        Dim varQc As Variant, strIniFile As String, strQC As String, intloop As Integer
        Dim objFileSystem As New Scripting.FileSystemObject
        Dim objStream As Scripting.textStream, strLine As String
    
        '----------------------------
        '������Ŀ
        Dim strTP As String
        Dim strALB As String
        Dim strGLOB As String
        Dim strAST As String
        Dim strALT As String
        Dim strAPOA1 As String
        Dim strAPOB As String
        Dim strALBGLOB As String
        Dim strASTALT As String
        Dim strAPOA1APOB As String
        Dim strTBIL As String
        Dim strDBIL As String
        Dim strTDBIL As String
        Dim strHDL As String
        Dim strLDL As String
        Dim strHL As String
        Dim strTmp As String
    
100     str_in = strSource
        
102     If InStr(str_in, Chr(&H4)) > 0 Then                 'EOT '����ͨѶ����
        
104         mlngStep_In = 0
        End If
    
106     If InStr(str_in, Chr(&H3)) > 0 Then                 '�յ�ETX(03)
108         lng_S = InStr(str_in, Chr(&H3))
110         mlngStep_Out = mlngStep_Out + 1
112         strCmd = Chr(&H6)
        End If
    
'114     If InStr(str_in, Chr(&H6)) > 0 Then                 '�յ�ACK(06)
'116         lng_S = InStr(str_in, Chr(&H6))
'118         str_in = Mid(str_in, 1, lng_S - 1) & Mid(str_in, lng_S + 1)
'120         mlngStep_Out = mlngStep_Out + 1
'
'122         If mstrWaitSend <> "" Then
'124             If Left(mstrWaitSend, 1) = "," Then
'126                 mstrWaitSend = Mid(mstrWaitSend, 2)
'                End If
'128             If InStr(mstrWaitSend, ",") > 0 Then
'130                 str_WaitSendLine = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
'132                 mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
'                Else
'134                 str_WaitSendLine = mstrWaitSend
'136                 mstrWaitSend = ""
'                End If
'
'138             If InStr(str_WaitSendLine, "|TSDWN^REPLY") > 0 Then
'140                 strCmd = Chr(&H5) & str_WaitSendLine
'                End If
'            End If
'        End If
    
142     If InStr(str_in, Chr(&H5)) > 0 Then                 '�յ�ENQ(05) Ӧ�� ETX(03)
144         strCmd = Chr(&H6)
        End If
    
146     If InStr(str_in, Chr(&H17)) > 0 Then                  '�յ�ETB(17) Ӧ��
148         strCmd = Chr(&H6)
        End If
    
    
150     If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
        'ȥ��
152     Do While mstrCode Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
154         lng_S = InStr(mstrCode, Chr(&H17))
156         lng_E = InStr(lng_S, mstrCode, Chr(&H2)) + 1
158         mstrCode = Mid(mstrCode, 1, lng_S - 1) & Mid(mstrCode, lng_E + 1)
        
        Loop
    
160     lng_S = InStr(mstrCode, Chr(&H5))
162     lng_E = InStr(mstrCode, Chr(&H4))
164     If lng_S > 0 And lng_E > 0 Then
166         str_Line = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
168         mstrCode = Mid(mstrCode, lng_E + 1)
        
170         If InStr(str_Line, "|TSREQ^REAL|") > 0 Then         '˫��ͨѶ- ��ʼ
172             mstrRack = ""
174             mlngStep_Out = 1
176             var_line = Split(str_Line, Chr(&HD))
            
178             For i_Count = LBound(var_line) To UBound(var_line)
180                 str_Line = var_line(i_Count)
182                 If str_Line Like "Q|*" Then
184                     str_���� = Trim(Split(str_Line, "|")(2))
                    
186                     mstrRack = Split(str_����, "^")(3) & "^" & Split(str_����, "^")(4) & "^" & Split(str_����, "^")(5) & "^^" & Split(str_����, "^")(7) & "^" & Split(str_����, "^")(8)
188                     str_���� = Split(str_����, "^")(2)
                    
    '1H|\^&|||cobas6000^1|||||host|TSREQ^REAL|P|1
    'Q|1|^^          260000055220^0^00031^4^^S1^SC||ALL||||||||O
    'L|1|N
    '35
    '1H|\^&|||cobas6000^1|||||host|TSREQ^REAL|P|1
    'Q|1|^^                   601^0^00034^1^^S1^SC||ALL||||||||O
    'L|1|N
    '96
    '


190                     If Val(str_����) <> 0 Then
192                         If Len(CStr(Val(str_����))) >= 12 Then
                                '����
194                             strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & Trim(str_����)
                            Else
                                '������ģʽ�£���ʱ���ֹ�����ı걾��
196                             strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|" & Trim(str_����)
                            End If
                        Else
                            '�걾��˫��ģʽ
198                         str_���� = Trim(Split(str_Line, "|")(2))
200                         str_���� = Trim(Split(str_����, "^")(3))
202                         strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|" & Trim(str_����)
                        End If
                    

                    End If
                Next
            
204             If strResult <> "" Then
206                 strResult = Mid(strResult, 3)
208                 strCmd = "1|"       '��������ͨѶ��ʼ���� 05
                End If

210         ElseIf InStr(str_Line, "|RSUPL^") > 0 Then     '������
            '-----------------------------------------------
    '            strIniFile = App.Path & "\Roche_P800.ini"
    '            strQC = ""
    '            If objFileSystem.FileExists(strIniFile) Then
    '                Set objStream = objFileSystem.OpenTextFile(strIniFile)
    '                Do While Not objStream.AtEndOfStream
    '                    strLine = objStream.ReadLine
    '                    If InStr(strLine, "=") > 0 Then
    '                        strQC = strQC & "|" & strLine
    '                    End If
    '                Loop
    '            Else
    '               Call objFileSystem.CreateTextFile(strIniFile)
    '               Set objStream = objFileSystem.OpenTextFile(strIniFile, ForAppending)
    '               objStream.WriteLine "[ˮƽ1]"
    '               objStream.WriteLine "5200 = PPU"
    '               objStream.WriteLine "[ˮƽ2]"
    '               objStream.WriteLine "6100 = PNU"
    '
    '               strQC = "5200 = PPU|6100 = PNU"
    '            End If
    '            If strQC <> "" Then varQC = Split(Mid(strQC, 2), "|")
                '-----------------------------------------------------------------------------
212             var_line = Split(str_Line, Chr(&HD))
            
214             For i_Count = LBound(var_line) To UBound(var_line)
216                 str_Line = var_line(i_Count)
218                 If str_Line Like "O|*" Then
220                     str_���� = Trim(Split(str_Line, "|")(7))
222                     If Len(str_����) <> 14 Then str_���� = Split(str_Line, "|")(22)
224                     str_���� = Left(str_����, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2) & ":" & Mid(str_����, 13, 2)
226                     str_�ʿ� = IIf(Trim(Split(str_Line, "|")(11)) = "Q", "1", "0")
                    
228                     str_�걾�� = Trim("" & Split(str_Line, "|")(2))
230                     If str_�걾�� = "" Then
232                         str_�걾�� = Trim("" & Split(str_Line, "|")(3))
                            '----�޸�
234                         str_�걾�� = Split(str_�걾��, "^")(0)
                        End If
236                     str_���� = Trim(Split(str_Line, "|")(5))
                    
238                     If Trim(str_�걾��) <> "" Then
240                         If Len(Trim(str_�걾��)) < 12 Then
                                '�걾��
242                             str_�걾�� = Trim(str_�걾��) & IIf(str_���� = "S", "^1^", "")
                            Else
                                '����
244                             str_�걾�� = "^" & IIf(str_���� = "S", "^1", "^") & Trim(str_�걾��)
                                '----�޸�
                                'str_�걾�� = Trim(str_�걾��)
                                'str_�걾�� = Trim(str_�걾��) & IIf(str_���� = "S", "^1^", "")
                            End If

                        End If
                    
246                     str_��Ŀ = ""
248                 ElseIf str_Line Like "R|*" Then
250                     str_��� = Trim(Split(str_Line, "|")(3))
252                     str�����־ = Trim(Split(str_Line, "|")(6))
254                     If str�����־ = "HH" Then
256                         str�����־ = "��"
258                     ElseIf str�����־ = "LL" Then
260                         str�����־ = "��"
                        Else
262                         str�����־ = ""
                        End If
                    
264                     If str_��� <> "" Then
266                         If InStr(str_���, "^") > 0 Then str_��� = Split(str_���, "^")(1) '--- 090407 �����ߵĽ����Ϊ -1^0.715 �ĸ�ʽ��
                        
268                         strTmp = Replace(Split(str_Line, "|")(2), "^^^", "")
270                         If InStr(strTmp, "/") > 0 Then
272                             If Trim(Split(strTmp, "/")(1)) <> "" Then
                                    '���� E2����ϡ�ͱ�������Ŀ
                                    '������Ŀ������Ҫ�ڼ���е����г�ָ�ꡣͨ�������⴦��Ϊ 12^5 ������ʽ
274                                 strTmp = Split(strTmp, "/")(0) & "^" & Trim(Split(strTmp, "/")(1))
                                Else
276                                 strTmp = Split(strTmp, "/")(0)
                                End If
                            
                            
                            End If
278                         str_��Ŀ = str_��Ŀ & "|" & strTmp & "|" & str�����־ & str_���
                        End If
                    End If
                Next
280             If IsDate(CDate(str_����)) And str_�걾�� <> "" And str_��Ŀ <> "" Then
282                 strResult = str_���� & "|" & str_�걾�� & "| |Ѫ��|" & str_�ʿ� & str_��Ŀ
'284                 WriteLog "�յ��������صĽ�� " & strResult, "C601"
                    '���� ������Ŀ
286                 var_line = Split(Mid(str_��Ŀ, 2), "|")
                
288                 For intloop = LBound(var_line) To UBound(var_line)
290                     If Trim(var_line(intloop)) = "320" Then 'Cysc
    '                        If Val(var_line(intLoop + 1)) <> 0 Then
    '                            strResult = strResult & "|4513|" & (78 / Val(var_line(intLoop + 1)) + 4)
    '                            WriteLog "������ ", "|4513|" & (78 / Val(var_line(intLoop + 1)) + 4), ""
    '                            Exit For
    '                        End If
                        End If
                    
292                     If Trim(var_line(intloop)) = "413" Then 'ALB
294                         strALB = Val(var_line(intloop + 1))
                        End If
                    
296                     If Trim(var_line(intloop)) = "678" Then 'TP
298                         strTP = Val(var_line(intloop + 1))
                        End If
                    
300                     If Trim(var_line(intloop)) = "687" Then 'AST
302                         strAST = Val(var_line(intloop + 1))
                        End If
                    
304                     If Trim(var_line(intloop)) = "685" Then 'ALT
306                         strALT = Val(var_line(intloop + 1))
                        End If
                    
308                     If Trim(var_line(intloop)) = "168" Then 'APOA1
310                         strAPOA1 = Val(var_line(intloop + 1))
                        End If
                    
312                     If Trim(var_line(intloop)) = "151" Then 'APOB
314                         strAPOB = Val(var_line(intloop + 1))
                        End If
                    
316                     If Trim(var_line(intloop)) = "257" Then 'APOB
318                         strTBIL = Val(var_line(intloop + 1))
                        End If
                    
320                     If Trim(var_line(intloop)) = "293" Then 'APOB
322                         strDBIL = Val(var_line(intloop + 1))
                        End If
                    
                    Next
                
324                 If strALB <> "" And strTP <> "" Then
326                     strGLOB = strTP - strALB
328                     strGLOB = Format(strGLOB, "0.0")
330                     strResult = strResult & "|5501|" & strGLOB
                    End If
                
332                 If strALB <> "" And strGLOB <> "" Then 'ALB/GLOB
334                     strALBGLOB = strALB / strGLOB
336                     strALBGLOB = Format(strALBGLOB, "0.0")
338                     strResult = strResult & "|5502|" & strALBGLOB
                    End If
                
340                 If strALT <> "" And strAST <> "" Then 'AST/ALT
342                     strASTALT = strAST / strALT
344                     strASTALT = Format(strASTALT, "0.00")
346                     strResult = strResult & "|5503|" & strASTALT
                    End If
                
348                 If strAPOA1 <> "" And strAPOB <> "" Then 'APOA1/APOB
350                     strAPOA1APOB = strAPOA1 / strAPOB
352                     strAPOA1APOB = Format(strAPOA1APOB, "0.0")
354                     strResult = strResult & "|5504|" & strAPOA1APOB
                    End If
                
356                 If strTBIL <> "" And strDBIL <> "" Then 'strTBIL-strDBIL
358                     strTDBIL = strTBIL - strDBIL
360                     strTDBIL = Format(strTDBIL, "0.0")
362                     strResult = strResult & "|5505|" & strTDBIL
                    End If
                
364                 If strHDL <> "" And strLDL <> "" Then 'strHDL/strLDL
366                     strHL = strHDL / strLDL
368                     strHL = Format(strHL, "0.00")
370                     strResult = strResult & "|5506|" & strHL
                    End If
                
                End If

            End If
        End If
        Exit Sub
hErr:
372     WriteLog "E601����", Erl() & "��," & Err.Description, "C601"
End Sub


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", _
                           Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
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
        Dim str_SendItem As String, str_EndString As String, str_tmp As String, lngRow As Long
        Dim strSendNull As String 'δ�ҵ��걾ʱ�����ָ�
        Dim strRackNo As String '�ֹ�����ʱָ���ܺţ�λ��
        Dim strJ As String, strW As String
        Dim strTmp As String, strN As String
        
        Dim dateS As Date
        
        On Error GoTo hErr
        
100     WriteLog "˫��(C601 ��ʼ)��", strSampleInfo, "C601"
'102     WriteLog "˫��(C601 ���)��iSendStep=" & iSendStep & ",strResponse=" & strResponse & ",blnSuccess=" & blnSuccess, "C601"
104     blnSuccess = True
106         If iSendStep = 0 Then
                'If strSampleInfo <> "" Then WriteLog "˫��(C601 ��ʼ)��" & strSampleInfo & " (" & mstrWaitSend & ")", "C601"
108             var_�걾 = Split(strSampleInfo, "||")
            
110             SendSample = ""
112             For i_Count = LBound(var_�걾) To UBound(var_�걾)
114                 str_��¼ = var_�걾(i_Count)
                
116                 If Val(Split(str_��¼, "|")(1)) = 0 Then
118                     str_�걾�� = Trim(Split(str_��¼, "|")(5))
                    Else
120                     str_�걾�� = Val(Split(str_��¼, "|")(1))
                    End If
122                 str_�걾���� = Split(str_��¼, "|")(3)
124                 strRackNo = Trim(Split(str_��¼, "|")(6))
126                 i_���� = Val(Split(str_��¼, "|")(4))
                
128                 str_���� = Trim(Split(str_��¼, "|")(5))
130                 If str_���� = "" Then str_���� = str_�걾��
                
132                 str_������Ŀ = ""
                
134                 For i_Loop = 10 To UBound(Split(str_��¼, "|"))
136                     If Val(Split(str_��¼, "|")(i_Loop)) > 0 Then
138                         str_tmp = Split(str_��¼, "|")(i_Loop)
140                         str_tmp = Replace(str_tmp, "not", "")
                        
142                         If InStr(str_������Ŀ, "^^^" & str_tmp & "^\") <= 0 Then
                                '����Ѫ�ǣ�ֻ��һ��
144                             str_������Ŀ = str_������Ŀ & "^^^" & str_tmp & "^\"
                            End If
                        End If
                    Next
                
146                 If str_������Ŀ <> "" And str_���� <> "" Then
148                     str_������Ŀ = Mid(str_������Ŀ, 1, Len(str_������Ŀ) - 1)
150                     If InStr(mstrRack, "^") > 0 Then
152                         str_SendItem = "H|\^&|||H7600^1||||||TSDWN^REPLY" & Chr(&HD) & "P|1" & Chr(&HD)
154                         strN = "N"
                        Else
156                         str_SendItem = "H|\^&|||H7600^1||||||TSDWN^BATCH" & Chr(&HD) & "P|1" & Chr(&HD)
158                         strN = "A"
                        End If
                    
160                     str_SendItem = str_SendItem & "O|1|" & str_���� & Space(24 - Len(str_����))
162                     If InStr(mstrRack, "^") > 0 Then
164                         str_SendItem = str_SendItem & "|" & str_�걾�� & Mid(mstrRack, InStr(mstrRack, "^")) & "|" & str_������Ŀ
                        Else
                            '�ֹ����͵���Rack��Ϣ
166                         strJ = "": strW = ""
168                         If InStr(strRackNo, ",") > 0 Then
170                             strJ = Trim(Split(strRackNo, ",")(0))
172                             strW = Trim(Split(strRackNo, ",")(1))
                            
174                             If strJ <> "" Then strJ = "^" & Format(strJ, "00000")
176                             If strW <> "" Then strW = "^" & strW & "^^S1^SC"
                            End If
178                         str_SendItem = str_SendItem & "|" & str_�걾�� & strJ & strW & "|" & str_������Ŀ
                        End If
180                     str_SendItem = str_SendItem & "|" & IIf(i_���� = 1, "S", "R")               'S-����,R-����
182                     str_SendItem = str_SendItem & "||" & Format(Now, "yyyyMMddHHmmss")
184                     str_SendItem = str_SendItem & "||||" & strN & "||||1||||||||||O" & Chr(&HD)           '2015-03-07 ��A��ΪN����ɽ·��A����˫��
186                     str_SendItem = str_SendItem & "C|1|I|" & str_�걾�� & "^^^^|G" & Chr(&HD)
188                     str_SendItem = str_SendItem & "L|1|N" & Chr(&HD) & Chr(&H3)
                              
190                     str_EndString = ""
192                     lngRow = 0
194                     Do While str_SendItem <> ""
                        
196                         If Len(str_SendItem) <= 230 Then
198                             lngRow = lngRow + 1
200                             str_EndString = str_EndString & Chr(&H2) & lngRow & str_SendItem & GetCheckCode(lngRow & str_SendItem) & vbCrLf
202                             str_SendItem = ""
                            Else
204                             lngRow = lngRow + 1
206                             str_tmp = Mid(str_SendItem, 1, 230) & Chr(&H17)
208                             str_EndString = str_EndString & Chr(&H2) & lngRow & str_tmp & GetCheckCode(lngRow & str_tmp) & vbCrLf
210                             str_SendItem = Mid(str_SendItem, 231)
                            End If
                        Loop
        
212                     mstrWaitSend = mstrWaitSend & "," & str_EndString & "," & Chr(&H4)
214                     SendSample = Chr(&H5)
216                     iSendStep = 1
                    Else
                        'δ�ҵ������Ӧ����Ŀ
218                     If str_���� <> "" Then
                            'strSendNull = GetIniKeyValue(App.Path & "\C601.ini", "Ӧ��", "����Ŀ", Chr(&H4))
220                         SendSample = Chr(&H5)  'strSendNull
                            mstrWaitSend = mstrWaitSend & "," & Chr(&H4)
                            iSendStep = 1
                        End If
                    End If
                Next
222             If SendSample <> "" Then WriteLog "˫��(E601 Send)��", SendSample, "C601"
                
224         ElseIf InStr(strResponse, Chr(&H6)) > 0 Then
226             If mstrWaitSend <> "" Then
                
228                 If Left(mstrWaitSend, 1) = "," Then
230                     mstrWaitSend = Mid(mstrWaitSend, 2)
                    End If
232                 If InStr(mstrWaitSend, ",") > 0 Then
234                     strTmp = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
236                     mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
                    Else
238                     strTmp = mstrWaitSend
240                     mstrWaitSend = ""
                    End If
                    
242                 SendSample = strTmp
244                 WriteLog "˫��(E601 ����)��", SendSample, "C601"
                End If
246             If mstrWaitSend = "" Then
248                 iSendStep = 0
250                 WriteLog "˫��(E601 ���)", "C601", ""
                End If
252         ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
254             WriteLog "˫��(E601 �յ�04 �˳�)", "C601", ""
256             mstrWaitSend = ""
258             iSendStep = 0
260         ElseIf InStr(strResponse, Chr(&H5)) > 0 Then
262             WriteLog "˫��(E601 �յ�05 �˳�)", "C601", ""
264             mstrWaitSend = ""
266             iSendStep = 0
268             SendSample = Chr(&H4)
            End If
270         'If SendSample <> "" Then
                '��ʱӦ��
272             'dateS = Now
                'Do
274             'Loop While DateDiff("s", dateS, Now) < 0.15
            'End If
'276     WriteLog "˫��(C601 ����)��iSendStep=" & iSendStep & ",blnSuccess=" & blnSuccess, "C601"
        Exit Function
hErr:
278     WriteLog "E601-˫��:", Erl() & "��," & Err.Description, "C601"
End Function


Private Function GetCheckCode(ByVal str_in As String) As String

    '��ȡУ����
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function


