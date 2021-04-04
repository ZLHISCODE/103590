VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABX_60_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ABX Micros 60
'����:ABX
'����:ŷ��ʤ
'�û�:��ݸݸ��ҽԺ

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
    
    Dim lngS As Long, lngE As Long, strCode As String
    Dim varLine As Variant, intRow As Integer, strLine As String
    Dim str�걾�� As String, str���� As String, strItems As String, lngItem As String
    Dim StrWbc As String, StrRbc As String, StrPLt As String
    Dim strImg As String
     lngS = InStr(strSource, Chr(2))
    lngE = InStr(strSource, Chr(3))
    strReserved = strSource
    If lngS > 0 And lngE > 0 And lngE > lngS Then
        strCode = Mid(strReserved, lngS, lngE - lngS + 1)
        strReserved = Mid(strReserved, lngE + 1)
        varLine = Split(strCode, Chr(&HD))
        For intRow = LBound(varLine) To UBound(varLine)
            strLine = varLine(intRow)
            If InStr("!,2,3,4,5,6,7,8,@,A,B,C,#,%,(,),*,+,-,/,.,',"""",$,&", Mid(strLine, 1, 1)) > 0 Then
                'ͨ�����������ת��
                lngItem = Mid(strLine, 1, 1)
                If lngItem = "@" Then
                   lngItem = "PLT"
                ElseIf lngItem = "'" Then
                   lngItem = "GRANP"
                ElseIf lngItem = """" Then
                   lngItem = "LYMNA"
                ElseIf lngItem = "$" Then
                   lngItem = "MONNA"
                ElseIf lngItem = "&" Then
                   lngItem = "MONNP"
                End If
                strItems = strItems & "|" & lngItem & "|" & Val(Mid(strLine, 3))
           ElseIf Mid(strLine, 1, 1) = "q" Then
                '����
                str���� = Format(CDate(Mid(strLine, 9, 2) & "-" & Mid(strLine, 6, 2) & "-" & Mid(strLine, 3, 2)), "yyyy-MM-dd")
            ElseIf Mid(strLine, 1, 1) = "u" Then
                '�걾��
                str�걾�� = Val(Mid(strLine, 3))
           ElseIf Mid(strLine, 1, 1) = "W" Then    'WBCͼ������
                StrWbc = Mid(strLine, 2)
                StrWbc = Graph_Text(StrWbc)
            ElseIf Mid(strLine, 1, 1) = "X" Then    'RBCͼ������
                StrRbc = Mid(strLine, 2)
                StrRbc = Graph_Text(StrRbc)
            ElseIf Mid(strLine, 1, 1) = "Y" Then    'PLTͼ������
                StrPLt = Mid(strLine, 2)
                StrPLt = Graph_Text(StrPLt)
           
            End If
        Next
        
        If Val(str�걾��) > 0 And strItems <> "" Then
            strResult = str���� & "|" & str�걾�� & "| |ѪҺ|0" & strItems
           If StrWbc <> "" Then strImg = strImg & "^" & "WBC;0;140,135,19-50,32.27-100,68-200,96.75-300,129-400" & StrWbc
            If StrRbc <> "" Then strImg = strImg & "^" & "RBC;0;140,135,15.45-30,55-100,103-200" & StrRbc
            If StrPLt <> "" Then strImg = strImg & "^" & "PLT;0;30,135,14-2,30-5,47-10,86-20,129-30" & StrPLt
            If strImg <> "" Then
                strResult = strResult & vbCrLf & Mid(strImg, 2)
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


Private Function Graph_Text(ByVal str_bin As String) As String
    '����ͼ���ı�����
    Dim lng_len As Long, lng_i As Long, str_bit As String
    Dim str_Return As String, str_in As String
    Dim tmpStr As String, tmpStr2 As String
     '----�ַ�ת���ɶ�������
    tmpStr = str_bin
    tmpStr2 = ""
    Do While tmpStr <> ""
        If Len(Hex(Asc(Mid(tmpStr, 1, 1)))) = 4 Then
            tmpStr2 = tmpStr2 & "," & Mid(Hex(Asc(Mid(tmpStr, 1, 1))), 1, 2) & "," & Mid(Hex(Asc(Mid(tmpStr, 1, 1))), 1, 2)
        Else
            tmpStr2 = tmpStr2 & "," & Hex(Asc(Mid(tmpStr, 1, 1)))
        End If
        
        tmpStr = Mid(tmpStr, 2)
    Loop
    
    
    lng_len = Len(tmpStr2)

    If lng_len < 3 Then
        Graph_Text = ""
        Exit Function
    Else
        str_in = tmpStr2
    End If

    For lng_i = 1 To lng_len / 3
        
        str_bit = Replace(Left(str_in, 3), ",", "")
        str_in = Mid(str_in, 4)
        
        If IsNull(str_bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & ";" & Format((CByte("&H" & str_bit) - 32), "000")
        End If
    Next
    Graph_Text = str_Return
End Function

