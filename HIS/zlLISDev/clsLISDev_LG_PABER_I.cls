VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LG_PABER_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ѪС��ۼ���Ѫ���ӷ�����
'�ͺ�:LG-PABER-I
'����:
'����:����

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
    
    Dim varIn As Variant '�����������
    Dim StrTime As String
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strLine As String
    Dim str�걾�� As String, str���� As String, str�����˱걾�ʿ� As String, strָ�� As String
    Dim aLine() As String
    Dim strDate As String
    Dim intLoop As Integer
    Dim aItem() As String
    Dim strNO As String
    Dim strNow As String
    Dim strItem As String
    Dim intCount As Integer
    Const StrAllitem As String = ",ADP,PT,TT,APTT,FIB,"
    
    On Error GoTo errHandle
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    'û�п�ʼ�ͽ����ַ�ʱ�˳�
    If InStr(strSource, "@c") = 0 Or InStr(strSource, Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD)) = 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    strNow = Mid(strSource, InStr(strSource, "@c") + 5, InStr(strSource, Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD)) + 5)
    
    If InStr(strNow, "Meter") > 0 Then
        strNow = Mid(strNow, InStr(strNow, "Meter") + 6)
    End If
    
    aLine = Split(strNow, vbCr)
    
    strDate = Replace(Mid(aLine(0), 1, 10), ".", "-") & " " & Time
    
    For intLoop = 3 To UBound(aLine)
        strLine = aLine(intLoop)
        strLine = Replace(strLine, Chr(30), " ")
        strLine = Replace(strLine, Chr(0), " ")
        strLine = Replace(strLine, Chr(9), " ")
        
        strLine = DelSapce(strLine)
        aItem = Split(strLine, " ")
        
        If InStr(strLine, "") > 0 Or InStr(strLine, Chr(&H14) & Chr(&H1E)) > 0 Then
            strNO = aItem(1)
        End If
       
        If UBound(aItem) > 1 Then
        
            If InStr(strLine, " ADP ") Then
                strItem = "ADP"
            End If
            
            If InStr(strLine, " PT ") Then
                strItem = "PT"
                strData = strData & "||" & strDate & "|" & strNO & "| |Ѫ��|0|PT|" & aItem(UBound(aItem) - 2)
            End If
            
            If InStr(strLine, " TT ") Then
                strData = strData & "||" & strDate & "|" & strNO & "| |Ѫ��|0|TT|" & aItem(UBound(aItem) - 1)
                strItem = ""
            End If
            
            If InStr(strLine, " APTT ") Then
                strData = strData & "||" & strDate & "|" & strNO & "| |Ѫ��|0|APTT|" & aItem(UBound(aItem) - 1)
                strItem = ""
            End If
            
            If InStr(strLine, " FIB ") Then
                strItem = ""
                strData = strData & "||" & strDate & "|" & strNO & "| |Ѫ��|0|FIB|" & aItem(UBound(aItem) - 1)
            End If
            
            If strItem = "ADP" Then
                If aItem(1) = "180" Then
                    strData = strData & "||" & strDate & "|" & strNO & "| |Ѫ��|0|ADP|" & aItem(UBound(aItem) - 1)
                End If
            End If
            
            If strItem = "PT" Then
                If aItem(2) = "INR" Then
                    strData = strData & "||" & strDate & "|" & strNO & "| |Ѫ��|0|INR|" & aItem(UBound(aItem) - 1)
                End If
            End If
        End If
        
    Next
    
    strResult = Mid(strData, 3)

    Exit Sub
    

errHandle:
    WriteLog "lg_Paber_i,�������", Err.Description, ""

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











