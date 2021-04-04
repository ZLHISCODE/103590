VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABX_P80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:Ѫ����
'����:??
'����:����
'�û�:�����ж�Ժ
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
    Dim aData() As String, iDataPoint As Integer, aItem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngloop As Long, intloop As Integer
    Dim strDate As String
    Dim i As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '�걾��
    Dim strBarcode As String            '�����
    Dim strImageData As String          'ͼ����Ϣ
    
    Dim dbl�ܵ��� As Double, dbl�׵��� As Double, dbl���ܶ�֬���� As Double, dbl���ܶ�֬���� As Double
    Dim dbl��֬����A As Double, dbl��֬����B As Double, dbl�ܵ����� As Double, dblֱ�ӵ����� As Double
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    If InStr(strSource, "") > 0 Then
        
        aData = Split(strSource, "")
        
        For lngloop = 0 To UBound(aData)
            
            aItem = Split(aData(lngloop), vbCr)
            
            'û�д�������ʱ�˳�
            If UBound(aItem) < 71 And lngloop = UBound(aData) Then
                If strData <> "" Then
                    strResult = strData & vbCrLf & Mid(strImageData, 2)
                End If
                strReserved = aData(lngloop)
                Exit Sub
            End If
            
            strTmp = ""

            i = -1
            For intloop = 0 To UBound(aItem)
                'ָ��
                If InStr(aItem(intloop), "! ") > 0 Then
                    i = i + 1
                    '�¼�¼
                    If strData <> "" Then
                        strData = strData & vbCrLf & Mid(strImageData, 3) & "||"
                    End If
                End If
                If i < 30 And i >= 0 Then
                    i = i + 1
                    strTmp = strTmp & "|" & i & "|" & Val(Mid(aItem(intloop), 2))
                End If
                
                'ͼ����Ϣ
                If InStr(aItem(intloop), "X ") > 0 Then         'RBC
                    strImageData = strImageData & "^" & GetImageStr("RBC", Mid(Replace(Replace(aItem(intloop), vbCr, ""), vbLf, ""), 2))
                End If
                
                If InStr(aItem(intloop), "Y ") > 0 Then         'PLT
                    strImageData = strImageData & "^" & GetImageStr("PLT", Mid(Replace(Replace(aItem(intloop), vbCr, ""), vbLf, ""), 2))
                End If

                If InStr(aItem(intloop), "Z ") > 0 Then         'BAS
                    strImageData = strImageData & "^" & GetImageStr("BAS", Mid(Replace(Replace(aItem(intloop), vbCr, ""), vbLf, ""), 2))
                End If

                If InStr(aItem(intloop), "[ ") > 0 Then         'WBC
                    strImageData = strImageData & "^" & GetImageStr("WBC", Mid(Replace(Replace(aItem(intloop), vbLf, ""), vbCr, ""), 2))
                End If
'
                '�걾��Ϣ
                If InStr(aItem(intloop), "q ") > 0 Then
                    i = intloop
                    strSampleNO = Trim(Replace(Replace(Replace(aItem(i + 2), "u", ""), vbCr, ""), vbLf, ""))
                    'strSampleNO = Val("9" & Mid(strSampleNO, 1, 3) & Right(strSampleNO, 2))
                    'strBarcode = Trim(Val(Trim(Replace(Mid(aItem(i + 4), 2), "u", ""))))
                    strData = strData & Now & "|" & strSampleNO & "| |ѪҺ|0" & strTmp

                End If
            Next
        Next
    Else
        strReserved = strSource
    End If
    
    If strData <> "" Then
        strResult = strData & vbCrLf & strImageData
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
    Dim intloop As Integer
    Dim strTmp As String
    
    Select Case strType
        Case "WBC"
            GetImageStr = GetWBCstr(Trim$(strData))
            GetImageStr = strType & ";1;" & Mid(GetImageStr, 2)
        Case Else
            For intloop = 1 To Len(strData)
        '        Debug.Print Mid(strData, intloop, 1)
                strTmp = Hex(Asc(Mid(strData, intloop, 1)))
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
    Dim intloop As Integer
    Dim strVal As String
    
    
    'д������
    For intloop = 1 To Len(strData)
        '����Y����
        If intloop <= 32 Then
            intY = 1
        Else
            intY = CInt(intloop / 32)
        End If
        intY = 128 - intY + 1
'        If intY <> intCheckY Then
'            strData = strData & ";" & vbCrLf
''            Debug.Print strData
'        End If
        
        '����X����
        intX = intloop Mod 32
        If intX = 0 Then intX = 32
        
        strVal = Mid(strData, intloop, 1)
        
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

