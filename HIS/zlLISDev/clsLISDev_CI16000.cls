VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CI16000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: Axsym Ci16000
'����: ����
'����: �¶�
'�û�: ��ɽ������ҽԺ
'����: 2014-11-22
'�޸ļ�¼:
'2014-11-22 ��ɽ������ҽԺ CI16000 ���°�LIS˫��ʱ�������ô��������ֱ�Ӵ�����������ʱ�䲻��23:59:59�롣�϶�����������ǰʱ�䡣����������


Private mstrData As String
Private mstrWaitSend As String

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

        Dim lngPos As Long
        Dim lngSeqi As Long
        Dim lngCount As Long
        Dim lngCount1 As Long
        Dim lngSeqj As Long
        Dim lngSeqjcnt As Long
        Dim lngfn As Long

    
        Dim varReply As Variant
        Dim strReplyMsg As String
        Dim strInstrReply As String
        Dim strChecksum As String
  
    
        Dim varRecord As Variant
        Dim varField As Variant
        Dim lngS As Long, lngE As Long  '��ʼ������λ��
        Dim strData As String
        Dim strLine As String
        Dim strBarcode As String    '����
        Dim strSampleNo As String   '�걾��
        Dim strQC As String         '�Ƿ��ʿ�
        Dim strDate As String       '����
        Dim strValue As String, strChannel As String '�����ͨ����
        Dim strItem As String
    
        On Error GoTo hErr
100     strData = strSource
102     strResult = ""
    
104     If strData <> "" Then
            '---   Ӧ�𲿷� ----
106         Call SaveLog("�յ��ַ�: " & strData)
108         mstrData = mstrData & strData
110         If InStr(strData, Chr(&H5)) > 0 Then
112             strCmd = Chr(&H6)
114             Call SaveLog("���յ��������͹�����ͨѶ�����ENQ��")
116         ElseIf InStr(strData, Chr(&H3)) > 0 Then
118             strCmd = Chr(&H6)
120             Call SaveLog("�յ�ETX��Ӧ��ACK�󣬵ȴ�ͨѶ������EOT��")
122         ElseIf InStr(strData, Chr(&H17)) > 0 Then
124             strCmd = Chr(&H6)
126             Call SaveLog("�յ�ETB��Ӧ��ACK�󣬼����ȴ���һ֡��")
            End If
    
    
128         lngS = InStr(mstrData, Chr(&H5))
130         lngE = 0
132         If lngS > 0 Then lngE = InStr(lngS, mstrData, Chr(&H4))
        
134         Do While lngS > 0 And lngE > 0
        
136             Call SaveLog("���յ�һ����������")
138             strData = Mid(mstrData, lngS, lngE - lngS + 1)
140             mstrData = Mid(mstrData, lngE + 1)
            
142             Do While strData Like "*" & Chr(&H17) & "*" & Chr(&HD) & Chr(&HA) & Chr(&H2) & "*"
144                 lngS = InStr(strData, Chr(&H17))
146                 lngE = InStr(lngS, strData, Chr(&H2)) + 1
148                 strData = Mid(strData, 1, lngS - 1) & Mid(strData, lngE + 1)
                Loop
            
                '------------------------------------------------------------------

150             Call SaveLog("��ʼ����")
152             varRecord = Split(strData, Chr(&HD))
154             strItem = "": strDate = "": strSampleNo = "": strBarcode = ""
            
156             For lngSeqi = LBound(varRecord) To UBound(varRecord)
158                 strLine = varRecord(lngSeqi)
    
160                 If strLine Like "*" & Chr(&H2) & "?Q|*" Then
162                     varField = Split(strLine, "|")
164                     strBarcode = Trim("" & varField(2))
166                     Call SaveLog("�������롰" & strBarcode & "��������ָ���ʼ��ȡ������Ŀ��Ϣ��")
168                     If Left(strBarcode, 1) = "^" Then strBarcode = Mid(strBarcode, 2)
                    
170                     strCmd = "1|"
172                     strResult = Format(Now, "yyyy-MM-dd") & "|^^" & strBarcode
                    
                        Exit Sub
174                 ElseIf strLine Like "*" & Chr(&H2) & "?O|*" Then

176                     If strItem <> "" And strDate <> "" And strSampleNo <> "" Then
178                         strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Serum|" & strQC & strItem
180                         Call SaveLog("���֡�" & strDate & "���걾��Ϊ��" & strSampleNo & "���ļ�������" & strItem & "��")
                           
                        End If
182                     strItem = "": strDate = "": strSampleNo = "": strBarcode = ""
                        
184                     varField = Split(strLine, "|")
186                     strSampleNo = Trim("" & varField(2))
                        
188                     If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
190                     strQC = Trim("" & varField(11))
192                     If strQC = "Q" Then
194                         strQC = "1"
                        Else
196                         strQC = "0"
                        End If
                        
198                 ElseIf strLine Like "*" & Chr(&H2) & "?R|*" Then
200                     varField = Split(strLine, "|")
202                     strChannel = Trim("" & varField(2))
204                     strValue = Trim("" & varField(3))
206                     If strDate = "" Then
208                         strDate = Trim("" & varField(12))
210                         strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                        End If
212                     If Right(strChannel, 2) = "^F" Then
214                         strChannel = Split(strChannel, "^")(3)
216                         If strChannel <> "" And strValue <> "" Then
218                             strItem = strItem & "|" & strChannel & "|" & strValue
                            End If
                        End If
                        
                    End If
            
                Next
            
220             If strItem <> "" And strDate <> "" And strSampleNo <> "" Then
222                 strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Serum|" & strQC & strItem
224                 Call SaveLog("���֡�" & strDate & "���걾��Ϊ��" & strSampleNo & "���ļ�������" & strItem & "��")
226                 strItem = "": strDate = "": strSampleNo = "": strBarcode = ""
                End If
                '------------------------------------------------------------------
                '����Ƿ�������������
228             lngS = InStr(mstrData, Chr(&H5))
230             lngE = 0
232             If lngS > 0 Then lngE = InStr(mstrData, Chr(&H4))
            Loop
        End If
234     If strResult <> "" Then strResult = Mid$(strResult, 3)
    
        Exit Sub
hErr:
236     Call SaveLog(CStr(Erl()) & "��," & Err.Description)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    Dim varInData As Variant, iLoop As Integer, strData As String, i As Integer
    Dim strBarcode As String, strEmergency As String, strTmp As String
    Dim strChannel As String, strSend As String, strItems As String, iRow As Integer
    Dim strSegment As String, strDate As String, strSampleType As String
    Dim strSampeNo As String, intOdx As Integer
    On Error GoTo hErr

100 SendSample = ""
102 blnSuccess = True
104 If iSendStep = 0 Then
        '�ս���˫��ģʽ����û�����ɷ�������
106     varInData = Split(strSample, "||")
108     For iLoop = LBound(varInData) To UBound(varInData)
110         strData = varInData(iLoop)
112         strDate = Format(CDate(Split(strData, "|")(0)), "yyyyMMdd000000")
114         strSampeNo = Trim(Split(strData, "|")(1))
            '--�걾
116         strSampleType = "" 'Trim(Split(strData, "|")(3))
            '--����
118         strBarcode = Trim(Split(strData, "|")(5))
            '---����
120         If Val(Split(strData, "|")(4)) = 1 Then
122             strEmergency = "S"
            Else
124             strEmergency = ""
            End If
            '----ָ��
126         strItems = ""

128         strSend = ""
130         If UBound(Split(strData, "|")) >= 10 And strBarcode <> "" Then
132             If Right(strItems, 1) = "\" Then strItems = Left(strItems, Len(strItems) - 1)
134             iRow = 0
136             strSegment = Chr(&H5)
138             iRow = iRow + 1
140             strSend = iRow & "H|\^&||||||||||P|1" & Chr(&HD) & Chr(&H3)
142             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
            
                'P|1||A43749||Smith^Tom^J||19950315^120^H|�Ա�M/F/U|||||ҽ��||||||||||||λ��<CR>
                'P|1||�걾��|<CR>
144             iRow = iRow + 1
146             strSend = iRow & "P|1||" & strSampeNo & "||" & strSampeNo & "||||||||||||||||||||" & Chr(&HD) & Chr(&H3)
148             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
            
                'O|1|SID_00002^3^1^^1||^^^IGG^^N^2\^^^IGA|S|||19950316135300|||N||||S|||1^1.00|||||||O<CR>
                'Oָ�� ��Ŀ��������ͨѶ���ϣ��ڴ˴���
                intOdx = 0
150             For i = 10 To UBound(Split(strData, "|"))
152                 strChannel = Trim$("" & Split(strData, "|")(i))
154                 strItems = strItems & "\^^^" & strChannel
                    
156                 If (i + 1) Mod 10 = 0 Or i = UBound(Split(strData, "|")) Then
158                     iRow = iRow + 1
                        intOdx = intOdx + 1
                        strItems = Mid(strItems, 2)
160                     strSend = iRow & "O|" & intOdx & "|" & strBarcode & "||" & strItems & "|" & strEmergency & "||" & strDate & _
                                "||||A||||||||||||||Q" & Chr(&HD) & Chr(&H3)
162                     strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                        strItems = ""
                    End If
                Next
                          
                'L
164             iRow = iRow + 1
166             strSend = iRow & "L|1" & Chr(&HD) & Chr(&H3)
168             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA) & Chr(&H4)
170             If strSegment <> "" Then mstrWaitSend = mstrWaitSend & strSegment


            End If
172         SendSample = GetSendData
174         If mstrWaitSend <> "" Then iSendStep = 1    '����δ�������ݣ����˳�˫��ģʽ

        Next
176     If SendSample <> "" Then
178         SaveLog "˫�����ɴ���������,����:" & SendSample & "����:" & mstrWaitSend
        Else
180         SaveLog "˫�����ɴ���������,����:" & mstrWaitSend
        End If
        
    Else
        '�յ�����Ӧ�𣬴ӻ�����ȡ��ָ���

182     If InStr(strResponse, Chr(&H6)) > 0 Then
184         SendSample = GetSendData
186         If mstrWaitSend = "" Then iSendStep = 0
188         SaveLog "����Ӧ��06��ͨѶ�ɹ�������:" & SendSample & "����:" & mstrWaitSend
190     ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
192         SendSample = GetSendData
194         If mstrWaitSend = "" Then iSendStep = 0
196         SaveLog "����Ӧ��04����ʾϣ����ͣ������Э���Ե���06,ͨѶ�ɹ�������:" & SendSample & "����:" & mstrWaitSend
198     ElseIf InStr(strResponse, Chr(&H15)) > 0 Then
200         SendSample = GetSendData
202         If mstrWaitSend = "" Then iSendStep = 0
204         SaveLog "����Ӧ��15��ͨѶʧ�ܡ�����:" & SendSample & "����:" & mstrWaitSend
206     ElseIf strResponse = Chr(&H5) Then
208         If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
210             mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
            Else
212             mstrWaitSend = ""
            End If
214         Call SaveLog("����Ӧ��05��������������Ȩ�����η���ȡ��")
216         iSendStep = 0
218         SendSample = Chr(&H6)
        Else
220         Call SaveLog("����Ӧ��" & strResponse)
        End If
222     If mstrWaitSend = "" Then iSendStep = 0    '�ѷ����꣬�˳�˫��ģʽ
    End If
    Exit Function
hErr:
224     SaveLog "��" & CStr(Erl()) & "�д���" & Err.Description

End Function

'------- �������ڲ�����
Private Function GetCheckCode(ByVal str_in As String) As String
    '����У����
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function

Private Function GetSendData() As String
    '�ӻ����л�ȡһ�����ݣ����ء����������û�����ݣ����ؿմ�
    If mstrWaitSend = "" Then Exit Function
    If Left$(mstrWaitSend, 1) = Chr(&H5) Then
        GetSendData = Chr(&H5)
        mstrWaitSend = Mid(mstrWaitSend, 2)
        Call SaveLog("����ͨѶ������05������ͨѶ��ʼ��")
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H2) Then
        If InStr(1, mstrWaitSend, Chr(&HA)) > 0 Then
            GetSendData = Mid(mstrWaitSend, 1, InStr(1, mstrWaitSend, Chr(&HA)))
            mstrWaitSend = Mid$(mstrWaitSend, InStr(1, mstrWaitSend, Chr(&HA)) + 1)
        End If
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H4) Then
        GetSendData = Chr(&H4)
        mstrWaitSend = Mid(mstrWaitSend, 2)
        Call SaveLog("����ͨѶ������04������ͨѶ������")
    End If
End Function

Private Sub SaveLog(ByVal StrInput As String)
    '------------------------------------------------------
    '--  ����:���ݵ��Ա�־,д��־����ǰĿ¼
    '------------------------------------------------------

    '���±������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As textStream
    Dim objFileSystem As New FileSystemObject

    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    If Dir(App.Path & "\����.TXT") = "" Then Exit Sub
    strFileName = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"

    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (strDate & ":" & StrInput)
    objStream.Close
    Set objStream = Nothing
End Sub




