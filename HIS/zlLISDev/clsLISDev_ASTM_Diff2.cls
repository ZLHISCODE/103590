VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ASTM_Diff2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strDeviceString '������Ϣ���豸��
Private strLastString As String '�ϴη��͵��豸��

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
    
    Dim iStrPoint As Long, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim aItemCan() As String, str�걾��Ϣ As String
    Dim sng1 As Single, sng2 As Single, iPos As Long
    Dim blnQc As Boolean, blnFirst As Boolean
    Dim aItemData() As String, strFldDelim As String, strRepeatDelim As String, strCompDelim As String
    Dim iMode As Integer '1:Current Mode ��2:New Mode
    Dim iResultType As Integer '������ͣ�1����������2����������
    Dim i As Long, strChannel As String, strSample As String, StrUnit As String '����Ƿ�����
    Dim str_WBC As String, str_RBC As String, str_PLT_R As String, str_PLT_F As String, Str_Img As String
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = "": blnQc = False
    iResultType = 1 '��ʼ�Ƿ��ؽ��
    
    iStrPoint = 1: blnReadItem = False
    'ȡ�������
    strSource = Replace(strSource, Chr(5), "")
    'ȡ��������
   ' strSource = Replace(strSource, Chr(4), "")
    'ȡ����ʼ��
    strSource = Replace(strSource, Chr(2), "")
    'ȡ���س����з�
    'strSource = Replace(strSource, vbCrLf, "")
    'ȡ�����з�
    strSource = Replace(strSource, Chr(10), "")
    'ȡ���Ʊ���
    strSource = Replace(strSource, Chr(9), "")
    'ȡ��ͨ�����ӷ�
    i = 1
    Do While True
        i = InStr(i, strSource, Chr(3))
        If i = 0 Then Exit Do
        If i + 2 <= Len(strSource) Then
            strSource = Mid(strSource, 1, i - 1) & Mid(strSource, i + 3)
        Else
            Exit Do
        End If
    Loop
    
    Do While iStrPoint <= Len(strSource)
        '������Ϣ�Ŀ�ʼλ�ã���H��ʼ
        iPos = iStrPoint
        strBuffer = LTrim(ReadLine(strSource, iPos))
        If Mid(strBuffer, 1, 2) Like "*H*" Then
            '��¼һ���������
            iCurrGrpEntry = iStrPoint
            
            '��¼�ָ���
            If Mid(strBuffer, 1, 2) Like "*H" Then
                strFldDelim = Mid(strBuffer, 3, 1) '�ֶηָ���
                strRepeatDelim = Mid(strBuffer, 4, 1) '�ظ����ݷָ���
                strCompDelim = Mid(strBuffer, 5, 1) '������ݷָ���
            Else
                strFldDelim = Mid(strBuffer, 2, 1) '�ֶηָ���
                strRepeatDelim = Mid(strBuffer, 3, 1) '�ظ����ݷָ���
                strCompDelim = Mid(strBuffer, 4, 1) '������ݷָ���
            End If
            
            '�����Ƿ�����
            If InStr(iStrPoint, strSource, "") > 0 Then
                '��Ϣ���������ٱ���
                iCurrGrpEntry = 0
                
                '���ұ걾��Ϣ
                iStrPoint = iPos: blnReadItem = False: blnFirst = True
                
                Do While iStrPoint <= Len(strSource)
                    If Not blnReadItem Then '��ǰ�Ƿ����ڶ�ȡ������Ŀ
                        Do While iStrPoint <= Len(strSource)
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Mid(strBuffer, 2, 1) = "O" Then Exit Do
                        Loop
                        If iStrPoint > Len(strSource) Then
                            '��ǰ��Ϣ��Ч�����ٴ�����
                            Exit Do
                        End If
                        
                        aItemData = Split(strBuffer, strFldDelim)
                        
                        '������ţ�����Ϊ����
                        strSample = Val(aItemData(2))
                        'aItemData = Split(strSample, strCompDelim)
                        'strSample = Val(Trim(aItemData(2)))
                        
                        blnReadItem = True '�걾�Ž�����ת���
                        '׼������
                    Else
                        Do While iStrPoint <= Len(strSource)
                            iPos = iStrPoint '������ǰλ��
                            strBuffer = LTrim(ReadLine(strSource, iStrPoint))
                            If Mid(strBuffer, 2, 1) = "R" Then
                            '������,��һ��Ҫȡʱ��
                               If blnFirst Then
                                    aItemData = Split(strBuffer, strFldDelim)
                                   '����ʱ��,ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                                    strData = aItemData(12)
                                    strData = Format(Mid(strData, 1, 4) & "-" & Mid(strData, 5, 2) & "-" & Mid(strData, 7, 2) & " " & Mid(strData, 9, 2) & ":" & Mid(strData, 11, 2) & ":" & Mid(strData, 13, 2), "yyyy-mm-dd hh:mm:ss")
                                    strData = strData & "|" & strSample
                                                            '������
                                    strData = strData & "| " '��������֧��
                                    '�걾
                                    strData = strData & "|ѪҺ"
                                    '�ʿ�Ʒ
                                    strData = strData & "|0"
                                    str�걾��Ϣ = strData
                                    aItemCan = Split(aItemData(2), strCompDelim)
                                    strChannel = UCase(Trim(aItemCan(3)))
                                    strData = strData & "|" & strChannel
                                    strData = strData & "|" & Format(Val(aItemData(3)), "0.00")
                                    blnFirst = False
                                Else
                                    aItemData = Split(strBuffer, strFldDelim)
                                     aItemCan = Split(aItemData(2), strCompDelim)
                                    strChannel = UCase(Trim(aItemCan(3)))
                                   ' If strChannel > 24 Then Exit Do
                                   'ͼ������
                                    Select Case strChannel
                                        Case UCase("WBC Histo")
'                                            strData = strData & Chr(13) & Chr(10) & strChannel & "|0| | | |" & to_Hex(aItemData(3))
                                            If InStr(aItemData(3), "") > 0 Then
                                                str_WBC = Mid(aItemData(3), 1, InStr(aItemData(3), "") - 1)
                                            End If
                                        Case UCase("RBC Histo")
'                                            strData = strData & "||" & str�걾��Ϣ & Chr(13) & Chr(10) & strChannel & "|0| | | |" & to_Hex(aItemData(3))
                                            If InStr(aItemData(3), "") > 0 Then
                                                str_RBC = Mid(aItemData(3), 1, InStr(aItemData(3), "") - 1)
                                            End If
                                        Case UCase("Plt Raw Histo")
                                            str_PLT_R = aItemData(3)
                                        Case UCase("Plt Fit Histo")
'                                            strData = strData & "||" & str�걾��Ϣ & Chr(13) & Chr(10) & strChannel & "|0| | | |" & to_Hex(aItemData(3))
                                            If InStr(aItemData(3), "") > 0 Then
                                                str_PLT_F = Mid(aItemData(3), 1, InStr(aItemData(3), "") - 1)
                                            End If
                                        Case Else
                                            strData = strData & "|" & strChannel
                                            StrUnit = UCase(aItemData(4))
                                            strData = strData & "|" & Format(Val(aItemData(3)), "0.00")
                                    End Select
                                   
                                End If
                            Else
                                If Mid(strBuffer, 3, 1) <> "|" And InStr(strBuffer, "|") > 0 Then
                                    If str_WBC <> "" And str_RBC = "" Then str_WBC = str_WBC & Mid(strBuffer, 2, InStr(strBuffer, "|") - 2)
                                    If str_RBC <> "" And str_PLT_F = "" Then str_RBC = str_RBC & Mid(strBuffer, 2, InStr(strBuffer, "|") - 2)
                                    If str_PLT_F <> "" Then str_PLT_F = str_PLT_F & Mid(strBuffer, 2, InStr(strBuffer, "|") - 2)
                                End If
                            End If
                            iCurrGrpEntry = 0
                        Loop
                        'һ���걾����
                        If str_WBC <> "" Then
                            Str_Img = Str_Img & "^WBC;0;200,200,19.90625-50,39.8125-100,59.71875-150,79.625-200,99.53125-250,119.4375-300,139.34375-350,159.25-400,179.15625-" & de_code(str_WBC)
                        End If
                        If str_RBC <> "" Then
                            Str_Img = Str_Img & "^RBC;0;200,185,18.75-50,56.25-100,112.5-200,168.75-250" & de_code(str_RBC)
                        End If
                        If str_PLT_R <> "" And str_PLT_F <> "" Then
                            Str_Img = Str_Img & "^PLT;4;180,180,5-2,12.5-5,25-10,37.5-15,50-,62.5-25,75-30;0,0,0,0,0," & Replace(Mid(de_code(str_PLT_R), 2), ";", ",") & "|" & Replace(Mid(de_code(str_PLT_F), 2), ";", ",")
                        End If
                        
                        strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData & _
                                    IIf(Str_Img <> "", vbCrLf & Mid(Str_Img, 2), "")
                        
                    End If
                Loop
            Else
                iStrPoint = iPos
                iCurrGrpEntry = 1
                Exit Do
            End If
        Else
          iStrPoint = iPos
        End If
    Loop
            
    '����δ������ԭʼ��
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    End If
    '���豸���͡������յ�������
    strCmd = Chr(&H6&)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    blnSuccess = True
    iSendStep = 0
    SendSample = ""
End Function

Private Function GetCheckCode(ByVal strSource As String) As String
'��ȡУ����
    Dim lngAsc As Long, i As Integer
    
    For i = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, i, 1))
    Next
    GetCheckCode = CStr(Hex(lngAsc Mod 256))
End Function

Private Function ReadSubString(ByVal strSource As String, iPoint As Long, Optional ByVal Length As Integer = 0) As String
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

Private Function ReadLine(ByVal strSource, iPoint As Long) As String
    '��ָ��λ�ÿ�ʼ��ȡһ�У�ͬʱָ�붨λ����һ������֮��
    Dim iPos As Long
    
    iPos = iPoint
    Do While iPos <= Len(strSource)
        If Mid(strSource, iPos, 1) = Chr(13) Then
            Exit Do
        End If
        iPos = iPos + 1
    Loop
    If iPos > Len(strSource) Then
        ReadLine = Mid(strSource, iPoint)
        iPoint = iPos
    Else
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function

Private Function to_Hex(str��¼ As String) As String
Dim i  As Integer
    to_Hex = ""
    For i = 1 To Len(str��¼)
            to_Hex = to_Hex & Hex(Asc(Mid(str��¼, i, 1)))
    Next
End Function






