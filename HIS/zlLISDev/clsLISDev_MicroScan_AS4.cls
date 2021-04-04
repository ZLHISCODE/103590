VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MicroScan_AS4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Microscan AS-4 ΢���������
'����:��������(DADE BEHRING INC)
'����:����
'�û�:����һ����ҽԺ
'ʱ��:20100508
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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    Dim str_Code As String, lng_S As Long, lng_len As Long
    Dim str_Day As String, str_���� As String, str�걾�� As String, strValues As String
    Dim int_i As Integer, strRecord As String
    Dim str_WBC As String, str_RBC As String, str_PLT As String
    Dim str_Data As String
    Dim strNow As String
    Dim strLine() As String
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strItem() As String
    Dim strData As String
    
    
    
    If InStr(strSource, Chr(&H2) & Chr(&H22) & Chr(&H48) & Chr(&H22)) <= 0 Or InStr(strSource, Chr(&H4)) <= 0 Then
        strReserved = strSource
        Exit Sub
    End If
    strNow = Mid$(strSource, InStr(strSource, Chr(&H2) & Chr(&H22) & Chr(&H48) & Chr(&H22)), InStr(strSource, Chr(&H4)))
    strReserved = Mid$(strSource, InStr(strSource, Chr(&H4)))
    
    
    
    '��ʼ����
    strLine = Split(strNow, vbCrLf)
    For intLoop = 0 To UBound(strLine)
        'ȥ���س�����
        strTmp = Replace$(strLine(intLoop), vbLf, "")
        strTmp = Replace$(strTmp, vbCr, "")
        strTmp = Replace$(strTmp, Chr(&H2), "")
        strTmp = Replace$(strTmp, Chr(&H3), "")
        
        If InStr(strTmp, ",") > 0 Then
            strItem = Split(strTmp, ",")
            Select Case Mid$(strTmp, 1, 1)
            
                Case "R"
                    If UBound(strItem) >= 12 Then
                        strData = strData & "||" & Mid(strItem(6), 1, 4) & "-" & Mid(strItem(6), 5, 2) & "-" & Mid(strItem(6), 7, 2) & " " & Format(Now, "hh:mm:ss") & _
                                "|" & Replace$(strItem(3), """", "") & "| |" & Replace$(strItem(12), """", "") & "|0"
                    End If
                Case "M"
                    '��Ŀ
                    If UBound(strItem) >= 6 Then
                        strData = strData & "|" & Replace(strItem(3), """", "") & "|" & Replace(strItem(7), """", "") & "^1^" & Replace(strItem(4), """", "")
                    End If
            End Select
        End If
    Next
    If strData <> "" Then strResult = Mid(strData, 3)
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



Private Function Graph_Text(ByVal str_bin As String) As String
    '����ͼ���ı�����
    Dim lng_len As Long, lng_i As Long, str_bit As String
    Dim str_Return As String
    lng_len = Len_Bin(str_bin)

    If lng_len < 1 Then
        Graph_Text = ""
        Exit Function
    End If

    For lng_i = 1 To lng_len
        str_bit = Replace(Mid_bin(str_bin, lng_i, 1, False), ",", "")
    
        If IsNull(str_bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & ";" & Format(CByte("&H" & str_bit), "000")
        End If
    Next
    Graph_Text = str_Return
End Function
