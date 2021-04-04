VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_DM2_Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'DM2  ��ˮ��
'�ٷ�����ҽԺ


Private mstrCode As String    '��������
Private mlngStep_In As Long   '�����ַ�����������
Private mstrACK As String     '�Ƿ�Ӧ��ACK,����Ӧ��ACK(06),��-Ӧ�� ETX(03)

Private mlngStep_Out As Long  '��˫����Ϣ����
Private mstr������ As String

Private mstrWaitSend As String '�����ͱ걾
Private mblnSendCmd  As Boolean

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
        
        mstrCode = mstrCode & strSource
        Dim blnSendCmd   As Boolean
        Dim lngS  As Long
        Dim strSginSample  As String
        Dim lngE  As Long
        Dim strSample() As String   '�Իس����н���
        Dim strRsultTemp  As String  '����ָ����
        Dim i As Integer
        Dim strDate  As String       '��������
        Dim strNO    As String       '�걾���
        Dim intEmergency   As Integer '����
        Dim strSampleType  As String  '�걾����
'        If InStr(strSource, Chr(&H4) & Chr(&H1)) > 0 Then
'            strCmd = "0|" & Chr(&H6)
'            Exit Sub
'        ElseIf InStr(Mid(strSource, 4, 8), ",702,01,") > 0 Then
'            trCmd = "0|" & Chr(&H3)
'            Exit Sub
'        ElseIf InStr(Mid(strSource, 4, 8), ",702,01,") > 0 Then
'
'        End If
        If Not mblnSendCmd Then
            strCmd = "0|" & Chr(&H6)
            mblnSendCmd = Not mblnSendCmd
        Else
            strCmd = "0|" & Chr(&H3)
            mblnSendCmd = Not mblnSendCmd
        End If
        lngS = InStr(mstrCode, Chr(&H4) & Chr(&H1))
        If lngS > 0 Then
            lngE = InStr(lngS + 2, mstrCode, Chr(&H4))
        End If
        If lngS > 0 And lngE > 0 Then
            strSginSample = Mid(mstrCode, lngS, lngE)
            mstrCode = Mid(mstrCode, lngE + 1)
            strSginSample = Replace(Replace(strSginSample, Chr(&H4), ""), Chr(&H1), "")
            strSample = Split(strSginSample, vbCrLf)
            For i = 0 To UBound(strSample)
                If InStr(Mid(strSample(i), 4, 8), ",702,01,") > 0 Then
                    strDate = "20" & Mid(Split(strSample(i), ",")(3), 5, 2) & "-" & Mid(Split(strSample(i), ",")(3), 3, 2) & "-" & Mid(Split(strSample(i), ",")(3), 1, 2)
                    strNO = Trim(Split(strSample(i), ",")(12))
                    intEmergency = IIf(Split(strSample(i), ",")(9) = "ST", 1, 0)
                    strSampleType = GetSampleClass(Split(strSample(i), ",")(9))
                ElseIf InStr(Mid(strSample(i), 4, 8), ",702,03,") > 0 Then
                    'ָ������Ϣ
                    strRsultTemp = strRsultTemp & Split(strSample(i), ",")(10) & "|" & Split(strSample(i), ",")(15) & "|"
                End If
            Next
            
            strResult = strResult & "||" & strDate & "|^" & intEmergency & "^" & strNO & "| |" & strSampleType & "|0|" & Mid(strRsultTemp, 1, Len(strResult) - 1)
        End If
        
End Sub


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
        'iSendStep����ǰ���Ͳ����Ĳ��衣�����������ͽ������򷵻�0�������Զ���1���ء�
        'strResponse����ǰ��������Ӧ����
        'strSampleInfo�����͵ı걾��Ϣ��
        '   '   ��0��Ԫ�أ�����ʱ��
            '   ��1��Ԫ�أ��������
            '   ��2��Ԫ�أ���������
            '   ��3��Ԫ�أ��걾����
            '   ��4��Ԫ�أ������־
            '   ��5��Ԫ�أ���������
            '   ��6��Ԫ�أ��̺ţ�����
            '   ��7��Ԫ�أ�����ID^�Ա�^��������^��������^���䵥λ^����ȫƴ^ϡ�ͱ���^����
            '   ��8��9Ԫ�أ�ϵͳ����
            '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
        Dim strSampleItem()  As String  '�걾��Ϣ
        Dim strTemp()        As String  'ÿ���걾�����߷��Ķ�
        Dim strPatient()     As String  '���˸����ֶ���Ϣ�������������ң��ٴ���ϵ�
        Dim i                As Integer
        Dim j                As Integer
        Dim intItemCount     As Integer    ' ������Ŀ��
        Dim strItem          As String     ' ������Ŀ��
        Dim strCheck         As String     'У�鴮
        
        blnSuccess = True
        If iSendStep = 0 Then
            strSampleItem = Split(strSampleInfo, "||")
            For i = LBound(strSampleItem) To UBound(strSampleItem)
                strTemp = Split(strSampleItem(i), "|")
                strPatient = Split(strTemp(7), "^")
                If UBound(strTemp) >= 10 Then
                    mstrWaitSend = mstrWaitSend & Chr(&H4) & Chr(&H1) & "|"
                    strCheck = "[00" & ",701,01,0000,00,1," & IIf(Val(strTemp(4)) = 1, "ST", "RO") & "," & _
                                GetSampleClass(strTemp(3)) & "," & strTemp(5) & "," & Space(20) & "���" & Space(25 - Len("���")) & "," & _
                                Space(25) & "," & strTemp(2) & Space(16 - Len(strTemp(2))) & "," & Space(15) & " ," & strPatient(0) & Space(12 - Len(strPatient(0))) & "," & _
                                Space(18) & "," & "000000," & "0000," & "����" & Space(16 - Len("����")) & "," & Space(3 - Len(strPatient(3))) & strPatient(3) & "," & _
                                GetAge(strPatient(4)) & ",000000," & GetSex(strPatient(3)) & "," & Space(25) & "," & Space(7) & ",    ,          ,"
                    intItemCount = UBound(strTemp) - 9
                    For j = 10 To UBound(strTemp)
                        strItem = strItem & GetTDM(strTemp(j)) & ",0"
                    Next
                    strCheck = strCheck & intItemCount & "," & strItem & "]"
                    mstrWaitSend = mstrWaitSend & GetCheckCode(strCheck) & vbCrLf & "|" & Chr(&H4) & "|"
                End If
                intItemCount = 0
                strCheck = ""
            Next
            
            '����˫��
            If InStr(mstrWaitSend, "|") > 0 Then
                SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
                iSendStep = 1
                mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|") + 1)
            End If
        Else
            If InStr(strResponse, Chr(&H6)) > 0 Or InStr(strResponse, Chr(3)) > 0 Then
                SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
                mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|") + 1)
            End If
            If mstrWaitSend = "" Then
                iSendStep = 0
            End If
       End If
        
End Function


Private Function ACKChar() As String
    If mstrACK = Chr(&H6) Then
        mstrACK = Chr(&H3)
    Else
        mstrACK = Chr(&H6)
    End If
    ACKChar = mstrACK
End Function

Private Function GetTDM(ByVal strIn As String) As String
    GetTDM = strIn & Space(4 - Len(strIn))
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


Public Function GetSampleClass(ByVal strClass As String) As String
    Select Case Trim(strClass)
        Case "Ѫ��"
            GetSampleClass = "SE"
        Case "�Լ�Һ"
            GetSampleClass = "SF"
        Case "��Һ"
            GetSampleClass = "UR"
        Case "Ѫ��"
            GetSampleClass = "PL"
        Case "ѪҺ"
            GetSampleClass = "BL"
        Case "����"
            GetSampleClass = "OT"
        Case "SE"
            GetSampleClass = "Ѫ��"
        Case "SF"
            GetSampleClass = "�Լ�Һ"
        Case "UR"
            GetSampleClass = "��Һ"
        Case "PL"
            GetSampleClass = "Ѫ��"
        Case "BL"
            GetSampleClass = "ѪҺ"
        Case "OT"
            GetSampleClass = "����"
    End Select
End Function


Private Function GetSex(ByVal strSex As String) As String
    '�Ա�
    Select Case Trim(strSex)
        Case "��"
            GetSex = "M"
        Case "Ů"
            GetSex = "F"
        Case "������"
            GetSex = "B"
    End Select
End Function

Private Function GetAge(ByVal strAge As String) As Integer
    '�������䵥λתΪ����
    If Mid(strAge, 1, 1) = "" Or Mid(strAge, 1, 1) = "��" Then
        GetAge = 5
    ElseIf Mid(strAge, 1, 1) = "��" Then
        GetAge = 4
    ElseIf Mid(strAge, 1, 1) = "��" Then
        GetAge = 3
    ElseIf Mid(strAge, 1, 1) = "��" Then
        GetAge = 2
    ElseIf Mid(strAge, 1, 1) = "ʱ" Then
        GetAge = 1
    End If
End Function





