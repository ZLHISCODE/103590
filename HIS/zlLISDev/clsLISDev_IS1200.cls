VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_IS1200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:IS 1200 ��ѧ������
'����:�Ĵ���������
'����:��;�
'�û�:������ҽԺ
'ʱ��:2013-12-06
Private str_Count As String '������ʱ
Private mstrSendCmd As String  '������ʱӦ���ϼ�ָ��
Private mstrRack As String    'ɨ����ʱ�����̺�,�ܺŵ���Ҫ���ص���Ϣ.
Private mstrWaitSend As String '�����ͱ걾

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = Chr(6)
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
    '����ͼ��ķ�ʽ��
    '                   1.ͼ�����ݸ���ָ�����ݺ�ʹ�ûس����з����ָ���
    '                   2.�ж��ͼ������ʱʹ��"^"���ָ�
    '                   3.����ͼ�����ݸ�ʽ: ͼ�񻭷� 0=ֱ��ͼ  1=ɢ��ͼ
    '                     a) ֱ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ);X1;X2;X3;X4;X5...
    '                     b) ɢ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ):
    '                        ��:00000100001000010000100010;00000100001000010000100010;
    '                        ˵��:1.ɢ��ͼ�Ե���ʽ����ÿһ��ʹ�÷ֺ����ָ�.
    '                             2.�ж��ٸ��ֺž��ж�����
    '                             3.ÿһ���ж��ٸ�����ÿһ�еĳ�����ȷ��
    '                             3.��ͼ�ķ����Ǵ����ϱ����»�������65*65��ͼ���Ǵ�65�п�ʼ��(���ϱ߿�ʼ��)
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim adata() As String, iDataPoint As Integer, aItem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngLoop As Long, intLoop As Integer
    Dim strDate As String, strItem As String, msStrSource As String
    Dim i As Long, str_Q As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '�걾��
    Dim strBarcode As String            '�����
    Dim strImageData As String          'ͼ����Ϣ
    Dim strNow As String, lngS As Long, lngE As Long
    Dim str_02 As Long '��¼02�Ĵ���
    strItem = ""
    str_02 = 0
    str_Q = 0
    'û�п�ʼ��ʱ�˳�
    lngS = InStr(strSource, "1H")
    lngE = InStr(strSource, "L|1")
    
    On Error GoTo errHand

    If InStr(strSource, Chr(&H5)) > 0 Then
'        'Ӧ��
        strCmd = Chr(&H6)
'        Exit Sub
    End If
    
    'û�н�����ʱ�˳�
    If InStr(strSource, Chr(&H4)) > 0 Then
'        'Ӧ��
        strCmd = Chr(&H6)

    End If
    

    'ȡ��ǰ��������
    '��ʼ����
    If lngS > 0 And lngE > 0 And lngE > lngS Then
      Do While lngE > 0 And lngS > 0
            msStrSource = Trim(Mid(strSource, lngS, lngE - lngS + 4))
            strSource = Trim(Mid(strSource, lngE + 4, Len(strSource) - lngE))
            lngS = InStr(strSource, "1H")
            lngE = InStr(strSource, "L|1")
            ''���Ҳ��滻�������
            adata = Split(msStrSource, Chr(&HD))
            For i = LBound(adata) To UBound(adata)
                strTmp = Trim(adata(i))
                'Debug.Print strTmp
                If InStr(strTmp, Chr(&H2)) > 0 Then str_02 = Mid(strTmp, InStr(strTmp, Chr(&H2)) + 1, IIf(str_02 > 9, 2, 1))
                If strTmp <> "" Then
                    '�걾��
                    If Split(strTmp, "|")(0) = "O" And strSampleNO = "" Then strSampleNO = Split(Split(strTmp, "|")(3), "^")(0)
                    If Split(strTmp, "|")(0) = "O" And strDate = "" Then
                        strDate = Mid(Split(Split(strTmp, "|")(22), "^")(0), 1, 4) & "-" & Mid(Split(Split(strTmp, "|")(22), "^")(0), 5, 2) & "-" & Mid(Split(Split(strTmp, "|")(22), "^")(0), 7, 2) & " " & Mid(Split(Split(strTmp, "|")(22), "^")(0), 9, 2) & ":" & Mid(Split(Split(strTmp, "|")(22), "^")(0), 11, 2) & ":" & Mid(Split(Split(strTmp, "|")(22), "^")(0), 13, 2)
                    End If
                    If Split(strTmp, "|")(0) = "O" Then str_Q = IIf(Split(strTmp, "|")(11) = "Q", 1, 0)
                    If Split(strTmp, "|")(0) = "R" Or Split(strTmp, "|")(0) Like "*R" Then
                        If InStr(strTmp, "") > 0 Then
                             If InStr(adata(i + 1), "") > 0 Then
                                If InStr(adata(i + 1), Chr(&H2)) > 0 Then str_02 = Mid(adata(i + 1), InStr(adata(i + 1), Chr(&H2)) + 1, IIf(str_02 > 9, 2, 1))
                                strTmp = Replace(Replace(strTmp, Mid(strTmp, InStr(strTmp, "")), ""), Chr(&HA) & Chr(&HD), "") & Replace(Trim(adata(i + 1)), Chr(&H2) & str_02, "")
                                strItem = strItem & "|" & Val(Split(Split(strTmp, "|")(2), "^")(3)) & "|" & Trim(Split(strTmp, "|")(3))
                             End If
                        Else
                            strItem = strItem & "|" & Val(Split(Split(strTmp, "|")(2), "^")(3)) & "|" & Trim(Split(strTmp, "|")(3))
                        End If
                    End If
                End If
            Next i
            If strDate <> "" And strSampleNO <> "" And strItem <> "" Then
                strResult = strResult & "||" & strDate & "|" & IIf(str_Q = 1, 8000 + Val(strSampleNO), Val(strSampleNO)) & "| |ȫѪ|" & str_Q & strItem
            End If
            strDate = ""
            strSampleNO = ""
            strItem = ""
            str_Q = 0
        Loop
        strSource = ""
        strReserved = ""
        'Ӧ��
        strCmd = Chr(&H6)
    Else
        strReserved = strReserved & strSource
    End If
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
    
errHand:
    WriteLog "IS1200", Err.Number & ":" & Err.Description, ""
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
Public Function SendSample11111(ByVal strSampleinfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    Dim str_SendItem As String, str_Line As String

    WriteLog "��ʼ", strSampleinfo, ""

    var_�걾 = Split(strSampleinfo, "||")
    SendSample11111 = ""
    For i_Count = LBound(var_�걾) To UBound(var_�걾)
        str_��¼ = var_�걾(i_Count)

        If Val(Split(str_��¼, "|")(1)) = 0 Then
            str_�걾�� = Trim(Split(str_��¼, "|")(5))
        Else
            str_�걾�� = Val(Split(str_��¼, "|")(1))
        End If
        str_�걾���� = Split(str_��¼, "|")(3)
        i_���� = Val(Split(str_��¼, "|")(4))
        str_���� = Trim(Split(str_��¼, "|")(5))
        str_������Ŀ = ""

        For i_Loop = 10 To UBound(Split(str_��¼, "|"))
            If Val(Split(str_��¼, "|")(i_Loop)) > 0 Then
                str_������Ŀ = str_������Ŀ & "^^^" & Split(str_��¼, "|")(i_Loop) & "^1\"
            End If
        Next

        If str_������Ŀ <> "" Then
            str_������Ŀ = Mid(str_������Ŀ, 1, Len(str_������Ŀ) - 1)

            str_Line = "1H|\^&|||IS11200^1|||||host|TSDWN^REAL|P|1" & Chr(&HD) & Chr(&H3)
            str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
            'P|1|||||||U||||||
            str_Line = "2P|1|" & str_���� & "||||||U||||||" & Chr(&HD) & Chr(&H3)
            str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","

            str_Line = "3O|1|" & str_���� & "|" & mstrRack & "|" & str_������Ŀ & "|" & IIf(i_���� = 1, "S", "R")
            str_Line = str_Line & "||" & Split(str_��¼, "|")(0) & "||||N|||1|||||||||||O" & Chr(&HD) & Chr(&H3)
            str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","

            str_Line = "4L|1|N" & Chr(&HD) & Chr(&H3)
            str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","

            If SendSample11111 = "" Then
                SendSample11111 = SendSample11111 & Chr(&H5)                              '���� ��Ϣ������־(EOT)
                mstrWaitSend = mstrWaitSend & "," & str_SendItem & Chr(&H4)
            Else
                mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & str_SendItem & Chr(&H4)
            End If
        End If
    Next
    iSendStep = 0
    WriteLog "����", "���͵���������Ϣ", SendSample11111
End Function

'��ȡУ����
Private Function GetCheckCode(ByVal str_in As String) As String
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 100)), 2)
End Function

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    'ָ�붨λ����һ������֮��
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '��ȡ�Ӵ���ͬʱָ���ƶ�����ȡ���ַ���֮��
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function


Private Function GetImageStr(strType As String, strData As String) As String
    Dim intLoop As Integer
    Dim strTmp As String
    
    Select Case strType
        Case "WBC"
            GetImageStr = GetWBCstr(Trim$(strData))
            GetImageStr = strType & ";1;" & Mid(GetImageStr, 2)
        Case Else
            For intLoop = 1 To Len(strData)
        '        Debug.Print Mid(strData, intloop, 1)
                strTmp = Hex(Asc(Mid(strData, intLoop, 1)))
                If Len(strTmp) >= 3 Then
                    GetImageStr = GetImageStr & ";" & Val("&H" & Mid(strTmp, 1, 2))
                    GetImageStr = GetImageStr & ";" & Val("&H" & Mid(strTmp, 3))
                Else
                    GetImageStr = GetImageStr & ";" & Val("&H" & Mid(strTmp, 1, 2))
                End If
            Next
            
            'д��һЩͼ�ι̶��Ĵ�С�ͱ�־λ��
            Select Case strType
                Case "RBC"
                    GetImageStr = strType & ";0;32,176,11-50,35-100,59-,83-200,104-,128-300,152-;" & Mid(GetImageStr, 2)
                Case "PLT"
                    GetImageStr = strType & ";0;32,170,7-2,27-,54-20,81-,108-40,135-,162-60;" & Mid(GetImageStr, 2)
                Case "BAS"
                    GetImageStr = strType & ";0;32,130,12-,25-100,38-,51-200,64-,77-300,90-,103-400,116-" & Mid(GetImageStr, 2)
            End Select
            
            
    End Select
    
End Function

Private Function GetWBCstr(strData) As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intLoop As Integer
    Dim strVal As String
    
    
    'д������
    For intLoop = 1 To Len(strData)
        '����Y����
        If intLoop <= 32 Then
            intY = 1
        Else
            intY = CInt(intLoop / 32)
        End If
        intY = 128 - intY + 1
'        If intY <> intCheckY Then
'            strData = strData & ";" & vbCrLf
''            Debug.Print strData
'        End If
        
        '����X����
        intX = intLoop Mod 32
        If intX = 0 Then intX = 32
        
        strVal = Mid(strData, intLoop, 1)
        
        Select Case strVal
            Case "1", "A"       '4
                GetWBCstr = GetWBCstr & "0001"
            Case "2", "B"       '3
                GetWBCstr = GetWBCstr & "0010"
            Case "3", "C"       '3,4v
                GetWBCstr = GetWBCstr & "0011"
            Case "4", "D"       '2
                GetWBCstr = GetWBCstr & "0100"
            Case "5", "E"       '2,4
                GetWBCstr = GetWBCstr & "0101"
            Case "6", "F"       '2,3
                GetWBCstr = GetWBCstr & "0110"
            Case "7"            '2,3,4
                GetWBCstr = GetWBCstr & "0111"
            Case "8"            '1
                GetWBCstr = GetWBCstr & "1000"
            Case "9"            '1,4
                GetWBCstr = GetWBCstr & "1001"
            Case "A"            '4
                GetWBCstr = GetWBCstr & "0001"
            Case "0"            '0
                GetWBCstr = GetWBCstr & "0000"
        End Select
        
        If intX = 32 Then
            GetWBCstr = GetWBCstr & ";"
        End If
    Next
End Function



