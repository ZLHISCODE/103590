VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Vitrol350"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'��������: ȫ�Զ�����������
'�����ͺ�: Vitrol350
'����:ǿ��
'����:������
'����:2015-09-11

Private mstrCode As String    '��������
Private mstrRack As String    'ɨ����ʱ�����̺�,�ܺŵ���Ҫ���ص���Ϣ.
Private mstrWaitSend As String '�����ͱ걾
Private strCode As String
Private mstrStrt As String
Private mstrEnd As String


Private mstrSampleInfo As String
Private mstrSampleInfoOne As String

Private mstrbj As String



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
    Dim strTmp  As String
    Dim strLine As String
    Dim strXu As String
    Dim varstrTmp As Variant
    Dim i As Integer, j As Integer
    Dim strNO As String
    Dim strDate As String
    Dim strItem As String
    Dim strLime As String
    Dim varSplit As Variant
    Dim strBJ As String
    Dim strCResult As String
    Dim strCItem As String
    Dim str_Line As String, str_Code As String, strLine2 As String
    Dim strCheck As String, lng_S As Long, lng_E As Long
'    strSource = Replace(Replace(strSource, Chr(1), ""), Chr(13), "")
'    strSource = Replace(Replace(strSource, "#1'", ""), "#13'", "")
    mstrStrt = Chr(1) '#1'
    mstrEnd = Chr(13) '#13'
    
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
'        Debug.Print mstrCode
    End If
    
    '-----Ӧ��
    If InStr(strSource, Chr(&HD)) > 0 Then
        If Mid(mstrCode, 4, 1) = "S" Then
            strTmp = " Y~R @-#N1 "
            strCheck = GetTwoCode(strTmp) & strTmp
            str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
            strCmd = str_Line
        ElseIf Mid(mstrCode, 4, 1) = "F" Then
            strTmp = "!Y"
            strCheck = GetTwoCode(strTmp) & strTmp
            str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
            strCmd = str_Line
        ElseIf Mid(mstrCode, 4, 1) = "D" Then
            strTmp = "Y"
            strXu = Asc(Mid(mstrCode, InStr(1, mstrCode, "D") - 1, 1))
            strCheck = GetTwoCode(strTmp) & Chr(strXu) & strTmp
            str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
            strCmd = str_Line
         ElseIf Mid(mstrCode, 4, 1) = "Z" Then
            strTmp = "Y"
            strXu = Mid(mstrCode, InStr(1, mstrCode, "Z") - 1, 1)
            strCheck = GetTwoCode(strTmp) & strXu & strTmp
            str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
            strCmd = str_Line
        ElseIf Mid(mstrCode, 4, 1) = "B" Then
            strTmp = "Y"
            strXu = Mid(mstrCode, InStr(1, mstrCode, "B") - 1, 1)
            strCheck = GetTwoCode(strTmp) & strXu & strTmp
            str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
            strCmd = str_Line
        End If
      
    End If
   ' ---ȡ����������
    lng_S = InStr(mstrCode, Chr(&H1))   '��&H5��ͷ��������������Ŀ�ʼλ��
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&HD)) '��&H4������������������Ľ���λ��
    Else
        lng_E = 0
    End If

    Do While lng_S > 0 And lng_E > 0 '�����ʼ����������Ϊ��
       str_Code = Mid(mstrCode, lng_S, lng_E - lng_S) '���شӿ�ʼλ�õ�����λ��֮����ַ�
       mstrCode = Mid(mstrCode, lng_E + 1)
       
       If Mid(str_Code, 4, 1) = "D" Then
            str_Code = Mid(Trim(Replace(Replace(str_Code, Chr(&H1), ""), Chr(&HD), "")), 1, Len(Trim(Replace(Replace(str_Code, Chr(&H1), ""), Chr(&HD), ""))) - 1)
            If strCode <> "" Then
                strCode = strCode & Mid(Trim(Replace(Replace(str_Code, Chr(&H1), ""), Chr(&HD), "")), 4, Len(Trim(Replace(Replace(str_Code, Chr(&H1), ""), Chr(&HD), ""))) - 1)
            Else
                strCode = strCode & str_Code
            End If
            Debug.Print strCode
            If InStr(strCode, "]") > 0 Then ' Mid(str_Code, Len(str_Code)) = GetCheckCode(Mid(str_Code, 1, Len(str_Code) - 1)) And
                varSplit = Split(strCode, "}")
                For j = 0 To UBound(varSplit) - 1
                    strLine = varSplit(j)
                    If j = 0 Then
                        strDate = Trim(Mid(varSplit(j), InStr(1, varSplit(j), "D") + 1, 14))
                        strDate = Year(Now) & "-" & Mid(strDate, 7, 2) & "-" & Mid(strDate, 9, 2) & " " & Mid(strDate, 1, 2) & ":" & Mid(strDate, 3, 2) & ":" & Mid(strDate, 5, 2)
                        strDate = Format(strDate, "yyyy-mm-dd mm:hh:ss")
    
                        strNO = Trim(Mid(varSplit(j), InStr(1, varSplit(j), "D") + 25, 15))
                        strBJ = Mid(varSplit(j), InStr(1, varSplit(j), "1.000") + 5, 10)
                        strItem = Mid(strBJ, 1, 1)
                        If strItem = " " Then
                            strItem = "space"
                        ElseIf strItem = ";" Then
                            strItem = "FK"
                        ElseIf strItem = "'" Then
                            strItem = "YO"
                        Else
                            strItem = strItem
                        End If
                        If strItem = "#" Then
                            strItem = Mid(strBJ, 1, 2)
                            strCResult = Mid(strBJ, 3)
                        Else
                            strCResult = Mid(strBJ, 2)
                        End If
                        If Mid(strCResult, 1, 3) = "�޽��" Then
                            strCResult = Val(Mid(strCResult, 1, 3))
                        Else
                            strCResult = Val(strCResult)
                        End If
                      
                    Else
                        strBJ = Mid(varSplit(j), 1, 10)
                        strItem = Mid(strBJ, 1, 1)
                        If strItem = " " Then
                            strItem = "space"
                        ElseIf strItem = ";" Then
                            strItem = "FK"
                        ElseIf strItem = "'" Then
                            strItem = "YO"
                        Else
                            strItem = strItem
                        End If
                        If strItem = "#" Then
                            strItem = Mid(strBJ, 1, 2)
                            strCResult = Mid(strBJ, 3)
                        Else
                            strCResult = Mid(strBJ, 2)
                        End If
                        If Mid(strCResult, 1, 3) = "�޽��" Then
                            strCResult = Val(Mid(strCResult, 1, 3))
                        Else
                            strCResult = Val(strCResult)
                        End If
                        End If
                        strCItem = strCItem & "|" & strItem & "|" & strCResult
'                    End If
                Next
            End If
            If strDate <> "" And strCItem <> "" And strNO <> "" Then
                strResult = strResult & "||" & strDate & "|^^" & strNO & "| |Ѫ��|0" & strCItem
                strCode = ""
            End If
            If strResult <> "" Then
                strResult = Mid(strResult, 3)
            End If
       End If
'        strCode = ""
       lng_S = InStr(mstrCode, Chr(&H1))   '��&H5��ͷ��������������Ŀ�ʼλ��
       If lng_S > 0 Then
           lng_E = InStr(lng_S, mstrCode, Chr(&HD))  '��&H4������������������Ľ���λ��
       Else
           lng_E = 0
       End If
    Loop

    
    
'    Debug.Print "1" & strSource
'    Debug.Print "2" & mstrCode
'    If Mid(strSource, 3, 1) = "S" Then
'        strTmp = " Y~R @-#N1 "
'        strCheck = GetTwoCode(strTmp) & strTmp
'        str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
'        strCmd = str_Line
'    ElseIf Mid(strSource, 3, 1) = "F" Then
'        strTmp = "!Y"
'        strCheck = GetTwoCode(strTmp) & strTmp
'        str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
'        strCmd = str_Line
'    ElseIf Mid(strSource, 3, 1) = "D" Then
'        strTmp = "Y"
'        strXu = Asc(Mid(strSource, InStr(1, strSource, "D") - 1, 1))
'        strCheck = GetTwoCode(strTmp) & Chr(strXu) & strTmp
'        str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
'        strCmd = str_Line
'        mstrCode = mstrCode & strSource
'    ElseIf Mid(strSource, 3, 1) = "Z" Then
'        strTmp = "Y"
'        strXu = Mid(strSource, InStr(1, strSource, "Z") - 1, 1)
'        strCheck = GetTwoCode(strTmp) & strXu & strTmp
'        str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
'        strCmd = str_Line
'    ElseIf Mid(strSource, 3, 1) = "B" Then
'        strTmp = "Y"
'        strXu = Mid(strSource, InStr(1, strSource, "B") - 1, 1)
'        strCheck = GetTwoCode(strTmp) & strXu & strTmp
'        str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
'        strCmd = str_Line
'        lng_S = InStr(mstrCode, "S")
'        If lng_S > 0 Then
'            lng_E = InStr(mstrCode, "B")
'        Else
'            lng_E = 0
'        End If
'        Do While lng_S > 0 And lng_E > 0
'            str_Code = Mid(mstrCode, lng_S - 2, lng_E - lng_S + 2)
'            mstrCode = Mid(mstrCode, lng_E + 5)
'            If str_Code <> "" Then
'                strLime = Mid(str_Code, InStr(str_Code, "D"))
'            End If
'            If InStr(strLime, "D") > 0 Then
'                varSplit = Split(strLime, "}")
'                For j = 0 To UBound(varSplit) - 1
'                    If j = 0 Then
'                        strDate = Trim(Mid(varSplit(j), InStr(1, varSplit(j), "D") + 1, 14))
'                        strDate = Year(Now) & "-" & Mid(strDate, 7, 2) & "-" & Mid(strDate, 9, 2) & " " & Mid(strDate, 1, 2) & ":" & Mid(strDate, 3, 2) & ":" & Mid(strDate, 5, 2)
'                        strDate = Format(strDate, "yyyy-mm-dd mm:hh:ss")
'
'                        strNO = Trim(Mid(varSplit(j), InStr(1, varSplit(j), "D") + 25, 15))
'                        strBJ = Mid(varSplit(j), InStr(1, varSplit(j), "1.000") + 5, 10)
'                        strItem = Mid(strBJ, 1, 1)
'                        strCResult = Val(Trim(Mid(strBJ, 2)))
'                        strCItem = strCItem & "|" & strItem & "|" & strCResult
'                    Else
'                        strBJ = Mid(varSplit(j), 1, 10)
'                        strItem = Mid(strBJ, 1, 1)
'                        strCResult = Val(Mid(strBJ, 2))
'                        strCItem = strCItem & "|" & strItem & "|" & strCResult
'                    End If
'                Next
'
'            Else
'                varSplit = Split(strLime, "}")
'                For j = 1 To UBound(varSplit) - 1
'                    strBJ = Mid(varSplit(j), 1, 10)
'                    strItem = Mid(strBJ, 1, 1)
'                    strCResult = Val(Mid(strBJ, 2))
'
'                    strCItem = strCItem & "|" & strItem & "|" & strCResult
'                Next
'            End If
'            lng_S = InStr(mstrCode, "S")
'            If lng_S > 0 Then
'                lng_E = InStr(mstrCode, "B")
'            Else
'                lng_E = 0
'            End If
'        Loop
'
'    End If
'    If strDate <> "" And strCItem <> "" And strNO <> "" Then
'        strResult = strResult & "||" & strDate & "|^^" & strNO & "| |Ѫ��|0" & strCItem
'        mstrCode = ""
'    End If
'    If strResult <> "" Then
'        strResult = Mid(strResult, 3)
'    End If
'    Exit Sub
hErr:
     WriteLog "Analyse", "��" & CStr(Erl()) & "��", Err.Description
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ���������
    '   ��3��Ԫ�أ��걾����
    '   ��4��Ԫ�أ������־
    '   ��5��Ԫ�أ���������
    '   ��6��Ԫ�أ��̺ţ�����
    '   ��7��Ԫ�أ�����ID^�Ա�^��������^����^����ȫƴ^ϡ�ͱ���
    '   ��8��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    Dim blnEnd As Boolean
    Dim strBuffer As String, strSampleID As String
    Dim i As Long, aItem() As String
    Dim strRackNo As String, iPos As Long
    

    On Error GoTo errHandle
    mstrStrt = Chr(1) '#1'
    mstrEnd = Chr(13) '#13'
'    SendSample = SendSampleTwo(strSample, iSendStep, blnSuccess, strResponse, blnUndo, iType)
    SendSample = SendSampleOne(strSample, iSendStep, blnSuccess, strResponse, blnUndo, iType)
    Exit Function
errHandle:
    Call WriteLog("V350.SendSample ����: ", Err.Number, Err.Description)
End Function


Public Function SendSampleTwo(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    'strSampleInfo:������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
        Dim varSampleInfo As Variant
        Dim strNO As String
        Dim strItems As String
        Dim strLine As String
        Dim intloop As Integer
        Dim str_Line As String
        Dim strSendCmd As String
        Dim i As Integer
        Dim intNO As Integer
        Dim j As Integer
        Dim strTmp As String
        Dim strSplit As String
        Dim strCheck As String
        strSplit = "<SPLIT>"
100     intNO = 2
102     varSampleInfo = Split(strSampleInfo, "||")
104     SendSampleTwo = ""
106     blnSuccess = True

108     If iSendStep = 0 Then
110         For intloop = LBound(varSampleInfo) To UBound(varSampleInfo)
112             strLine = varSampleInfo(intloop)
114             If Val(Split(strLine, "|")(5)) > 0 Then
116                 strNO = Split(strLine, "|")(5)
                Else
118                 strNO = Split(strLine, "|")(1)
                End If

120             If strNO <> "" And UBound(Split(strLine, "|")) >= 10 Then
                    'ָ�ʼ
                    strTmp = "S~R @-#N1 "
                    strCheck = GetTwoCode(strTmp) & " " & strTmp
                    str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd & strSplit
                    strTmp = "FSFILE00"
                    strCheck = GetTwoCode(strTmp) & "!" & strTmp
                    str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd & strSplit
                    strTmp = "D" & strNO & Space(3) & "10 1.000"
'                    strTmp = "D|1              " & Left(Space(15) & strNO, 15) & "10!1.000"   30!0.998
                
132                 For i = 10 To UBound(Split(strLine, "|"))
136                     If Split(strLine, "|")(i) = "space" Then
                            strItems = strItems & " "
                        ElseIf Split(strLine, "|")(i) = "FK" Then
                            strItems = strItems & ";"
                        ElseIf Split(strLine, "|")(i) = "YO" Then
                            strItems = strItems & "'"
                        Else
                            strItems = strItems & Split(strLine, "|")(i)
                        End If
                    Next
                    strTmp = strTmp & strItems & "]"
                    strCheck = GetTwoCode(strTmp) & Chr(34) & strTmp
                    str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd & strSplit
                    strTmp = "Z"
                    strCheck = GetTwoCode(strTmp) & "#" & strTmp
                    str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd & strSplit
                    strTmp = "B"
                    strCheck = GetTwoCode(strTmp) & "$" & strTmp
                    str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd & strSplit
152                 iSendStep = 1    '��ʼͨ��

154                 mstrWaitSend = mstrWaitSend & strSplit & str_Line

                End If
            Next
156         SendSampleTwo = GetSendData
        Else
158         If strResponse <> "" Then SendSampleTwo = GetSendData            '�յ�06
        End If
160     If mstrWaitSend = "" Then iSendStep = 0 '�˳�˫��ͨ
        Exit Function
hErr:
'162     SaveLog "SendSample��" & CStr(Erl()) & "��," & Err.Description
End Function


Public Function SendSampleOne(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    'strSampleInfo:������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
        Dim varSampleInfo As Variant
        Dim strNO As String
        Dim strItems As String
        Dim strLine As String
        Dim intloop As Integer
        Dim str_Line As String
        Dim strSendCmd As String
        Dim i As Integer
        Dim intNO As Integer
        Dim j As Integer
        Dim strTmp As String
        Dim strCheck As String
        Dim strDate As String
        SendSampleOne = ""
        blnSuccess = True


        If iSendStep = 0 Then
            If strSampleInfo <> "" Then
                mstrSampleInfo = strSampleInfo
                If Left(mstrSampleInfo, 2) = "||" Then
                    mstrSampleInfo = Mid(mstrSampleInfo, 3)
                End If
                If InStr(mstrSampleInfo, "||") > 0 Then
                    mstrSampleInfoOne = Mid(mstrSampleInfo, 1, InStr(mstrSampleInfo, "||") - 1)
                    mstrSampleInfo = Mid(mstrSampleInfo, InStr(mstrSampleInfo, "||"))
                Else
                    mstrSampleInfoOne = mstrSampleInfo
                    mstrSampleInfo = ""
                End If
                
                If mstrSampleInfoOne <> "" And UBound(Split(mstrSampleInfoOne, "|")) >= 10 Then
                    iSendStep = 1
                    strTmp = "S~R @-#N1 "
                    strCheck = GetCode(strTmp) & " " & strTmp
                    str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
                End If
                SendSampleOne = str_Line
            End If
        Else
            If InStr(strResponse, " Y") > 0 Then
                strTmp = "FSFILE00"
                strCheck = GetCode(strTmp) & "!" & strTmp
                str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
                SendSampleOne = str_Line
            ElseIf InStr(strResponse, "!Y") > 0 Then

                    strLine = mstrSampleInfoOne
                    If Val(Split(strLine, "|")(5)) > 0 Then
                        strNO = Split(strLine, "|")(5)
                    Else
                        strNO = Split(strLine, "|")(1)
                    End If
'                    strDate = Format(Now, "yyyymmddhhmmss ")
'                    strDate = Mid(strDate, 8, 6) & Mid(strDate, 5, 4)
                    strTmp = "D" & strNO & Space(3) & "10 1.000"
                    
'                   strTmp = "D|1" & Left(strNO & Space(15), 15) & "10!1.000"
                   
                    For i = 10 To UBound(Split(strLine, "|"))
                        If Split(strLine, "|")(i) = "space" Then
                            strItems = strItems & " "
                        ElseIf Split(strLine, "|")(i) = "FK" Then
                            strItems = strItems & ";"
                        ElseIf Split(strLine, "|")(i) = "YO" Then
                            strItems = strItems & "'"
                        Else
                            strItems = strItems & Split(strLine, "|")(i)
                        End If
                        
                    Next
                    strTmp = strTmp & strItems & "]"
                    strCheck = GetCode(strTmp) & Chr(34) & strTmp
                    str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
                    SendSampleOne = str_Line

            ElseIf InStr(strResponse, Chr(34) & "Y") > 0 Then
                strTmp = "Z"
                strCheck = GetCode(strTmp) & "#" & strTmp
                str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
                SendSampleOne = str_Line
            ElseIf InStr(strResponse, "#Y") > 0 Then
                strTmp = "B"
                strCheck = GetCode(strTmp) & "$" & strTmp
                str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
                SendSampleOne = str_Line
            ElseIf InStr(strResponse, "$Y") > 0 Then
                If mstrSampleInfo <> "" Then
                    If Left(mstrSampleInfo, 2) = "||" Then
                        mstrSampleInfo = Mid(mstrSampleInfo, 3)
                    End If
                    mstrSampleInfoOne = Mid(mstrSampleInfo, 1, InStr(mstrSampleInfo, "||") - 1)
                    mstrSampleInfo = Mid(mstrSampleInfo, InStr(mstrSampleInfo, "||"))
                    If mstrSampleInfoOne <> "" And UBound(Split(mstrSampleInfoOne, "|")) >= 10 Then
                        strTmp = "S~R @-#N1 "
                        strCheck = GetCode(strTmp) & " " & strTmp
                        str_Line = mstrStrt & strCheck & GetCheckCode(strCheck) & mstrEnd
                        SendSampleOne = str_Line
                    End If
                Else
                    iSendStep = 0
                End If
            End If
        End If
        Exit Function
hErr:
'    SaveLog "SendSample��" & CStr(Erl()) & "��," & Err.Description
End Function




Private Function GetSendData() As String
    If mstrWaitSend <> "" Then
        If Left(mstrWaitSend, 7) = "<SPLIT>" Then
            mstrWaitSend = Mid(mstrWaitSend, 8)
        End If
        If InStr(mstrWaitSend, "<SPLIT>") > 0 Then
            GetSendData = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "<SPLIT>") - 1)
            mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "<SPLIT>"))
        Else
            GetSendData = mstrWaitSend
            mstrWaitSend = ""
        End If
    End If

End Function

Private Function GetCheckCode(ByVal str_in As String) As String

'��ȡУ����
    Dim lngAsc As Long, i As Integer, lngAsc2 As Long
    
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    lngAsc2 = lngAsc
    lngAsc = lngAsc And 192
    lngAsc = lngAsc / 64
    lngAsc2 = lngAsc2 + lngAsc
    lngAsc2 = lngAsc2 And 63
    lngAsc2 = lngAsc2 + 32
    GetCheckCode = Chr(lngAsc2)
'    GetCheckCode = Chr(((lngAsc + (lngAsc And 192) / 64) And 63) + 32)
    
End Function
Private Function GetTwoCode(ByVal str_in As String) As String
    If Len(str_in) > 1 Then
        GetTwoCode = Chr(Len(str_in) + 1 + 32)
    Else
        GetTwoCode = Chr(Len(str_in) + 2 + 32)
    End If
End Function

Private Function GetCode(ByVal strin As String) As String
    GetCode = Chr(Len(strin) + 2 + 32)
End Function


