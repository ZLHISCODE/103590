VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA_7020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    Dim iStrPoint As Integer, iStrPointBK As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim blnReadItem As Boolean '��ǰ�Ƿ����ڶ�ȡ������Ŀ
    Dim aData() As String, iDataPoint As Integer, aItem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngLoop As Long, intLoop As Integer
    Dim strDate As String

    Dim strLine As String               '��ǰҪ�������ִ�
    Dim lngBegin As Long                '��ʼλ��
    Dim lngEnd As String                '����λ��
    Dim str_����  As String
    Dim str_�� As String                '�����־
    Dim str_�ʿ� As String              '�ʿ�
    Dim str_tmp As String
    Dim curALB As Currency, curGLO As Currency, curTBIL As Currency, curDBIL As Currency, curIBIL As Currency
    Dim curTP As Currency, curAST As Currency, curALT As Currency, cur9 As Currency, cur10 As Currency
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    'û���ҵ���ʼλ��ʱ�˳�
    
    If Not strSource Like "***" Then
        strReserved = strSource
        Exit Sub
    End If
    
    lngBegin = InStr(strSource, Chr(&H2))
    lngEnd = InStr(strSource, Chr(&H3))
    
    'ȡ��ǰһ�εĽ���������ķ���
    strLine = Mid(strSource, lngBegin, lngEnd - lngBegin + 1)
    strReserved = Mid(strSource, lngEnd + 1)
    If InStr(strReserved, Chr(&H2)) <= 0 Then
        strReserved = ""    'û���ҵ���ʼλ��ʱ�˳�
    End If
    '�յ�һ�����ݣ�Ӧ��
    strCmd = Chr(&H6)

    '��ʼ����
    If InStr(",:,1,2,", Mid(strLine, 2, 1)) > 0 And InStr(",N,Q,F,", UCase(Mid(strLine, 3, 1))) > 0 Then
    
        If UCase(Mid(strLine, 3, 1)) = "Q" Then
            str_�� = "1"
        End If
        
        str_�ʿ� = "0"
        If Mid(strLine, 3, 1) = "F" Then
            str_�ʿ� = "1"
        End If
        
        
        If Trim(Mid(strLine, 33, 2)) = "" Or Trim(Mid(strLine, 29, 2)) = "" Or Trim(Mid(strLine, 31, 2)) = "" Then
            str_���� = Format(Now, "yyyy-MM-dd")
        Else
            str_���� = Mid(strLine, 33, 2) & "-" & Mid(strLine, 29, 2) & "-" & Mid(strLine, 31, 2)
        End If
        
        strData = Format(CDate(str_����), "yyyy-MM-dd") & "|" & Val(Mid(strLine, 5, 5)) & IIf(str_�� = "1", "^1^", "") & "| |Ѫ��|" & str_�ʿ�
        
        '��ʼ�������
        strLine = Replace(Mid(strLine, 45), "", "")
        
        For intLoop = 0 To Len(strLine) / 10
            If Len(strLine) < 10 Then Exit For
            str_tmp = Replace(Mid(strLine, 1, 10), "%", "")
            
            strData = strData & "|" & Trim(Mid(str_tmp, 1, 3)) & "|" & Val(Trim(Mid(str_tmp, 4)))
            
            strLine = Mid(strLine, 11)
        Next
        
        strResult = strData
        If strData <> "" Then
            aData = Split(strData, "|")
            For intLoop = 5 To UBound(aData)
                If intLoop Mod 2 = 1 Then
                    If Val(Trim(aData(intLoop))) = 7 Then curTP = Val(Trim(aData(intLoop + 1)))
                    If Val(Trim(aData(intLoop))) = 8 Then curALB = Val(Trim(aData(intLoop + 1)))
                    
                    If Val(Trim(aData(intLoop))) = 5 Then curTBIL = Val(Trim(aData(intLoop + 1)))
                    If Val(Trim(aData(intLoop))) = 6 Then curDBIL = Val(Trim(aData(intLoop + 1)))
                    
                    If Val(Trim(aData(intLoop))) = 1 Then curALT = Val(Trim(aData(intLoop + 1)))
                    If Val(Trim(aData(intLoop))) = 2 Then curAST = Val(Trim(aData(intLoop + 1)))
                    
                    If Val(Trim(aData(intLoop))) = 9 Then cur9 = Val(Trim(aData(intLoop + 1)))
                    If Val(Trim(aData(intLoop))) = 10 Then cur10 = Val(Trim(aData(intLoop + 1)))
                End If
            Next
            
            '--���� �򵰰ף�A/G
            If curALB <> 0 And curTP <> 0 Then
                If curTP >= curALB Then
                    curGLO = curTP - curALB
                    strResult = strResult & "|101|" & Format(curGLO, "0.000")
                End If
                If curGLO <> 0 Then
                    strResult = strResult & "|102|" & Format(curALB / curGLO, "0.000")
                End If
            End If
            '-- ���� ��ӵ����أ�DBIL/TBIL
            If curTBIL <> 0 And curDBIL <> 0 Then
                curIBIL = curTBIL - curDBIL
                strResult = strResult & "|103|" & Format(curIBIL, "0.000")
                strResult = strResult & "|105|" & Format(curDBIL / curTBIL, "0.000")
            End If
            '���� AST/ALT
            If curAST <> 0 And curALT <> 0 Then
                strResult = strResult & "|106|" & Format(curAST / curALT, "0.000")
            End If
            
            '����9/10
            If cur9 <> 0 And cur10 <> 0 Then
                '����/����
                strResult = strResult & "|104|" & Format(cur9 / cur10, "0.000")
            End If
            
        End If
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