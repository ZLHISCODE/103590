VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UF50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: UF-50
'����: sysmex
'����: ����
'�û�: �Ͼ���������ҽԺ
'�޸���:���
'�޸�����:2010-04-28
'�޸�����:�޸ı걾��ȡ��

Private str_Count As String '������ʱ
Private mstrSendCmd As String  '������ʱӦ���ϼ�ָ��

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
    
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim strSampleType As String '��������
    Dim strTmpReturn As String, blnEnd As Boolean
    Dim lngLength As Long, str���� As String, str�걾�� As String, str�����˱걾�ʿ� As String, strLine As String
    Dim strItem As String, i As Integer, blnLineEnd As Boolean
    Dim strTmp As String, intStart As Integer, intEnd As Integer
    Dim intPosition As Integer                  '��ǰ��
    Dim intloop As Integer
    
    Dim dblWBC_M As Double, dblRBC_M As Double, dblEC_M As Double
    Dim dblCAST_M As Double, dblBACT_M As Double
    Dim strExp As String
    Dim str_Image As String, strImgLine As String
    Dim lng_S As Long, str_O As String
    
    Dim var_image As Variant '��ͼ������
    Dim str_Tmp As String
    Dim str_Path As String, str_File As String, intTmp As Integer
    Dim str_RetuImage As String
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    
    'û���ҵ���ʼλ��ʱ�˳�
    If strSource Like "*O*" & Chr(&H3) & "*" Then
        lng_S = InStr(strSource, "O")
        str_O = Mid(strSource, lng_S, lng_S + 26 - 1)
        
            strCmd = "1|" & Chr(&H6)
            
            mstrSendCmd = "SGO" & Mid(str_O, 11) & "|"
            strReserved = Mid(strSource, 1, lng_S - 1) & Mid(strSource, lng_S + 26)

        If Not strReserved Like "**" Then strReserved = ""
        Exit Sub
    End If
    
    strTmp = Replace(strSource, str_Count, "")
    If InStr(strTmp, Chr(&H6)) > 0 Then
        strCmd = Chr(&H6)
        str_Count = strSource
    End If
    If strTmp Like "**" & Chr(&H3) & "*" Then
        str_Count = strSource
        strCmd = Chr(&H6)
    End If
    
    If Not strSource Like "*DS*DE*" & Chr(&H3) & "*" Then strReserved = strSource: Exit Sub
    'û�е��ﳤ��ʱ�˳�
    If Len(strSource) < 520 Then strReserved = strSource: Exit Sub
    'û���ҵ�������ʱ�˳�
    strSource = Mid(strSource, InStr(strSource, "DS"))
    strTmp = Trim(Mid(strSource, 520, 1)) & Chr(&H3)
    If InStr(strSource, strTmp) <= 0 Then strReserved = strSource: Exit Sub
    intStart = InStr(strSource, "DS")
    'intEnd = InStr(strSource, strTmp) + 2
    intEnd = intStart + 520
    'ֻ������ǰһ���걾
    strLine = Mid(strSource, intStart, intEnd - intStart)
    
    strReserved = Mid(strSource, InStr(strSource, strTmp) + 2)
    
    Do While InStr(strSource, Chr(&H3)) > 0
        
        strSource = Mid(strSource, InStr(strSource, "") + 1)
        strImgLine = Mid(strSource, 1, InStr(strSource, Chr(&H3)) - 1)
        strSource = Mid(strSource, InStr(strSource, Chr(&H3)) + 1)
        
        If Left(strImgLine, 2) = "DI" Then
            If Mid(strImgLine, 5, 2) = "01" Then
                str_Image = str_Image & "|" & Mid(strImgLine, 79, 10) & "^" & Mid(strImgLine, 95)
            Else
                str_Image = str_Image & Mid(strImgLine, 30)
            End If
        End If
        
    Loop
    
    If InStr(strReserved, "DS") <= 0 Then
        'û�п�ʼ��ʱ�˳�
        strReserved = ""
    End If
    
    str���� = Format(Mid(strLine, 11, 2) & "-" & Mid(strLine, 13, 2) & "-" & Mid(strLine, 15, 2), "yyyy-mm-dd")
    str�걾�� = Trim(Mid(strLine, 40, 5))
    
    strData = str���� & "|" & str�걾�� & "| |��|0"
    '��ʼ����
    intPosition = InStr(30, strLine, "0201")
    '��һ��
    For intloop = 0 To 4
        strData = strData & "|" & Mid(strLine, intPosition + (intloop * 12), 4) & "|" & Val(Mid(strLine, intPosition + (intloop * 12) + 5, 8))
        If Mid(strLine, intPosition + (intloop * 12), 4) = "0201" Then
            dblWBC_M = Val(Mid(strLine, intPosition + (intloop * 12) + 5, 8))
            strExp = strExp & "|WBC-M|" & dblWBC_M * 0.18
        End If
        
        If Mid(strLine, intPosition + (intloop * 12), 4) = "0202" Then
            dblRBC_M = Val(Mid(strLine, intPosition + (intloop * 12) + 5, 8))
            strExp = strExp & "|RBC-M|" & dblRBC_M * 0.18
        End If
        
        If Mid(strLine, intPosition + (intloop * 12), 4) = "0100" Then
            dblEC_M = Val(Mid(strLine, intPosition + (intloop * 12) + 5, 8))
            strExp = strExp & "|EC-M|" & dblEC_M * 0.18
        End If
        
        If Mid(strLine, intPosition + (intloop * 12), 4) = "0000" Then
            dblCAST_M = Val(Mid(strLine, intPosition + (intloop * 12) + 5, 8))
            strExp = strExp & "|CAST-M|" & dblCAST_M * 2.92
        End If
        
        If Mid(strLine, intPosition + (intloop * 12), 4) = "0401" Then
            dblBACT_M = Val(Mid(strLine, intPosition + (intloop * 12) + 5, 8))
            strExp = strExp & "|BACT-M|" & dblBACT_M * 0.18
        End If
        
    Next
    '�ڶ���
    intPosition = InStr(strLine, "0A12")
    If intPosition > 0 Then
        For intloop = 0 To 19
            If Trim(Mid(strLine, intPosition + (intloop * 12), 4)) <> "" Then
                strData = strData & "|" & Mid(strLine, intPosition + (intloop * 12), 4) & "|" & Val(Mid(strLine, intPosition + (intloop * 12) + 5, 8))
            End If
        Next
    End If
    'RBCI , 51, 1
    '0-��
    '1-Normocytic?
    '2-Microcytic?
    '3-Non Classified?
    If Mid(strLine, 51, 1) = 0 Then
        strData = strData & "|RBCI|��"
    ElseIf Mid(strLine, 51, 1) = "1" Then
        strData = strData & "|RBCI|Normocytic?"
    ElseIf Mid(strLine, 51, 1) = "2" Then
        strData = strData & "|RBCI|Microcytic?"
    ElseIf Mid(strLine, 51, 1) = "3" Then
        strData = strData & "|RBCI|Non Classified?"
    End If
    
    'COND,52,5  total,92,6
    strData = strData & "|COND|" & Val(Mid(strLine, 52, 5)) & "|TOTAL|" & Val(Mid(strLine, 92, 6)) & strExp
    
    strResult = strData
    
    If str_Image <> "" Then
        
        str_Image = Mid(str_Image, 2)
        var_image = Split(str_Image, "|")
        
        str_Path = App.Path & "\Gram"
        If Dir(str_Path, vbDirectory) = "" Then
            MkDir str_Path
        End If
        
        For i = LBound(var_image) To UBound(var_image)
            str_File = ""
            Do While str_File = ""
                intTmp = intTmp + 1
                str_File = Replace(str����, "-", "") & "_" & Format(Val(str�걾��), "000") & "_" & Format(intTmp, "00") & ".gif"
                If Dir(str_Path & "\" & str_File) <> "" Then str_File = ""
            Loop
            
            str_Tmp = var_image(i)
            If InStr(str_Tmp, "WBC") > 0 Then
                str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                If uf_ufHist(str_Tmp, Len(str_Tmp), "c:\tempfile1.bmp", str_Path & "\" & str_File) = 1 Then
                    str_RetuImage = str_RetuImage & "^WBC Fsc;100;1;" & str_Path & "\" & str_File
                End If
            ElseIf InStr(str_Tmp, "RBC") > 0 Then
                str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                If uf_ufHist(str_Tmp, Len(str_Tmp), "c:\tempfile1.bmp", str_Path & "\" & str_File) = 1 Then
                    str_RetuImage = str_RetuImage & "^RBC Fsc;100;1;" & str_Path & "\" & str_File
                End If
            ElseIf InStr(str_Tmp, "Fscw-Flw") > 0 Then
                str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                If uf_ufscat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 256, 128) = 1 Then
                    str_RetuImage = str_RetuImage & "^Fscw-Flw;100;1;" & str_Path & "\" & str_File
                End If
            ElseIf InStr(str_Tmp, "Fl2-Fsc") > 0 Then
                str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                If uf_ufscat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 256, 128) = 1 Then
                    str_RetuImage = str_RetuImage & "^Fl2-Fsc;100;1;" & str_Path & "\" & str_File
                End If
            ElseIf InStr(str_Tmp, "Fl-Fsc") > 0 Then
                str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                If uf_ufscat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
                    str_RetuImage = str_RetuImage & "^Fl-Fsc;100;1;" & str_Path & "\" & str_File
                End If
            ElseIf InStr(str_Tmp, "Fscw-Fl") > 0 Then
                str_Tmp = Mid(str_Tmp, InStr(str_Tmp, "^") + 1)
                If uf_ufscat(str_Tmp, Len(str_Tmp), "C:\tempfile1.bmp", str_Path & "\" & str_File, 128, 128) = 1 Then
                    str_RetuImage = str_RetuImage & "^Fscw-Fl;100;1;" & str_Path & "\" & str_File
                End If
            End If
        Next
        If str_RetuImage <> "" Then strResult = strResult & vbCrLf & Mid(str_RetuImage, 2)
   End If
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
    If InStr(mstrSendCmd, "|") > 0 Then
        SendSample = Split(mstrSendCmd, "|")(0)
        mstrSendCmd = Mid(mstrSendCmd, InStr(mstrSendCmd, "|") + 1)
    End If
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

