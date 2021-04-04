VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA_7080"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  ����7080
'ҽԺ:  ��������ҽԺ
'����:  �¶�
'ʱ��:  2008/8/13
Const str_MOR = ">"
Const str_REP = "?"

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
    Dim aData() As String, iDataPoint As Integer, aitem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngLoop As Long, intloop As Integer
    Dim strDate As String
    Dim i As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '�걾��
    Dim strBarcode As String            '�����
    Dim strImageData As String          'ͼ����Ϣ
    Dim strLine As String               '��ǰҪ�������ִ�
    Dim lngBegin As Long                '��ʼλ��
    Dim lngEnd As String                '����λ��
    Dim str_����  As String
    Dim str_�� As String                '�����־
    Dim str_�ʿ� As String              '�ʿ�
    Dim str_Tmp As String
    Dim curALB As Currency, curGLO As Currency, curAST As Currency, curALT As Currency
    Dim curTP As Currency, curTBIL As Currency, curDBIL As Currency
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    'û���ҵ���ʼλ��ʱ�˳�
'    If InStr(strSource, Chr(&H2)) <= 0 Then
'        strReserved = strSource
'        'strCmd = Chr(&H6)
'        Exit Sub
'    End If
'    'û���ҵ�����λ��ʱ�˳�
'    If InStr(strSource, Chr(&H3)) <= 0 Then
'        strReserved = strSource
'        'strCmd = Chr(&H6)
'        Exit Sub
'    End If
    
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
    
    If strLine = str_MOR Or strLine = str_REP Then
        '˫��ͨ�ŵ���������ԭ��Ӧ�𼴿�
        strCmd = str_MOR
    ElseIf strLine Like ";A1*" And Len(strLine) = 43 Then
        str_���� = Format(CDate(Mid(strLine, 37, 2) & "-" & Mid(strLine, 33, 2) & "-" & Mid(strLine, 35, 2)), "yyyy-MM-dd")
        strResult = str_���� & "|^^" & Trim(Mid(strLine, 15, 13))
        strCmd = "1|" '�յ�ɨ�����������
    Else
        '��ʼ����
        If InStr(",:,1,2,", "," & Mid(strLine, 2, 1) & ",") > 0 Then
            '������ ʱ��
            If IsNumeric(Mid(strLine, 36, 2)) And IsNumeric(Mid(strLine, 32, 2)) And IsNumeric(Mid(strLine, 34, 2)) _
               And IsNumeric(Mid(strLine, 38, 2)) And IsNumeric(Mid(strLine, 40, 2)) Then
                str_���� = "20" & Mid(strLine, 36, 2) & "-" & Mid(strLine, 32, 2) & "-" & Mid(strLine, 34, 2) & " " & Mid(strLine, 38, 2) & ":" & Mid(strLine, 40, 2)
                If IsDate(CDate(str_����)) Then
                    str_���� = Format(CDate(str_����), "yyyy-MM-dd HH:mm")
                Else
                    str_���� = Format(Now, "yyyy-MM-dd HH:mm:ss")
                End If
            Else
                str_���� = Format(Now, "yyyy-MM-dd HH:mm:ss")
            End If
            If UCase(Mid(strLine, 3, 1)) = "Q" Then
                str_�� = "1"
            End If
            
            str_�ʿ� = "0"
            If Mid(strLine, 2, 1) = "1" Then
                str_�ʿ� = "1"
            End If
            If Val(Mid(strLine, 5, 5)) = 0 Then 'ȡ����
                
                If Val(Mid(strLine, 15, 2)) <> 0 Then
                    '����ǰ��λ��Ϊ0 ����������
                    strData = str_���� & "|^" & IIf(str_�� = "1", "1", "") & "^" & Val(Mid(strLine, 15, 13)) & "| |Ѫ��|" & str_�ʿ�
                Else
                    '����ǰ��λΪ0,�Ͱ�������Ϊ�걾�Ŵ���
                    strData = str_���� & "|" & Val(Mid(strLine, 15, 13)) & IIf(str_�� = "1", "^1^", "") & "| |Ѫ��|" & str_�ʿ�
                End If
            Else
                strData = str_���� & "|" & Val(Mid(strLine, 5, 5)) & IIf(str_�� = "1", "^1^", "") & "| |Ѫ��|" & str_�ʿ�
            End If
            '��ʼ�������
            strLine = Replace(Mid(strLine, 50), "", "")
            
            For intloop = 0 To Len(strLine) / 9
                If Len(strLine) < 9 Then Exit For
                str_Tmp = Mid(strLine, 1, 9)
                strData = strData & "|" & Trim(Mid(str_Tmp, 1, 2)) & "|" & Val(Trim(Mid(str_Tmp, 3)))
                
                strLine = Mid(strLine, 10)
            Next
            
            strResult = strData
            If strData <> "" Then
                aData = Split(strData, "|")
                For intloop = 5 To UBound(aData)
                    If intloop Mod 2 = 1 Then
                        If Val(Trim(aData(intloop))) = 7 Then curTP = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 8 Then curALB = Val(Trim(aData(intloop + 1)))
                        
                        If Val(Trim(aData(intloop))) = 1 Then curTBIL = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 2 Then curDBIL = Val(Trim(aData(intloop + 1)))
                        
                        If Val(Trim(aData(intloop))) = 4 Then curAST = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 3 Then curALT = Val(Trim(aData(intloop + 1)))
                    End If
                Next
                If curALB <> 0 And curTP <> 0 Then
                    If curTP >= curALB Then
                        curGLO = curTP - curALB
                        strResult = strResult & "|GLO|" & Format(curGLO, "0.000")
                    End If
                    If curGLO <> 0 Then
                        strResult = strResult & "|A/G|" & Format(curALB / curGLO, "0.000")
                    End If
                End If
                If curAST <> 0 And curALT <> 0 Then
                    strResult = strResult & "|A/A|" & Format(curAST / curALT, "0.000")
                End If
                If curTBIL <> 0 And curDBIL <> 0 Then
                    strResult = strResult & "|IBIL|" & Format(curTBIL - curDBIL, "0.000")
                End If
            
            End If
        End If
        strCmd = str_MOR
    
    End If
    'strCmd = Chr(&H6)
End Sub



Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    Dim varItem As Variant, strItem As String, i As Integer, strSampleNO As String
    blnSuccess = True
    If strSample <> "" Then
        strItem = String(86, "0")
        
        varItem = Split(strSample, "|")
        If strResponse Like ";A1*" And Len(strResponse) = 43 Then
            If varItem(5) = Trim(Mid(strResponse, 15, 13)) Then
                
                strSampleNO = Val(varItem(1))
                strSampleNO = Space(5 - Len(strSampleNO)) & strSampleNO
                
                SendSample = Mid(strResponse, 1, 4) & strSampleNO & Mid(strResponse, 10, 33)
                
                For i = 10 To UBound(varItem)
                    If varItem(i) <= 86 And varItem(i) > 0 Then
                        strItem = Mid(strItem, 1, varItem(i) - 1) & "1" & Mid(strItem, varItem(i) + 1)
                    End If
                Next
                If strItem <> String(86, "0") Then
                    SendSample = SendSample & " 86" & strItem & "01000000"
                End If
            End If
            
        End If

    End If
End Function







