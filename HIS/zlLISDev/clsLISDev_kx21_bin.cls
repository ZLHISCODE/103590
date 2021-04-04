VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_KX21_bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:  sysmex KX21
'����:  ϣɭ���� sysmex
'ҽԺ:  �ɶ�����
'����:  �¶�
'ʱ��:  2008/3/31
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
    strReserved = strSource
    lng_len = Len_Bin(strReserved)
    If lng_len < 457 Then Exit Sub
    
    For lng_S = 1 To lng_len - 4
        str_Code = Mid_bin(strReserved, lng_S, 1)
        str_Day = Mid_bin(strReserved, lng_S + 3, 1)
        
        If str_Code = Chr(&H2) And str_Day = "/" Then
            If lng_len - lng_S < 456 Then Exit Sub
            str_Data = Mid_bin(strReserved, lng_S, 457, False)
            strReserved = Mid_bin(strReserved, lng_S + 457, 0, False)
            If strReserved = ",00" Then strReserved = ""
            str_���� = Format(CDate(Trim(Mid_bin(str_Data, 2, 2)) & "-" & Trim(Mid_bin(str_Data, 5, 2)) & "-" & Trim(Mid_bin(str_Data, 8, 2))), "yyyy-MM-dd")
            Exit For
        Else
            If str_Code = "" Then
                str_Day = Mid_bin(strReserved, lng_S + 1, 8)
                If IsDate(Left(str_Day, 4) & "-" & Mid(str_Day, 5, 2) & "-" & Right(str_Day, 2)) Then
                    If lng_len - lng_S < 456 Then Exit Sub
                    str_Data = Mid_bin(strReserved, lng_S, 457, False)
                    strReserved = Mid_bin(strReserved, lng_S + 457, 0, False)
                End If
                str_���� = Format(CDate(Left(str_Day, 4) & "-" & Mid(str_Day, 5, 2) & "-" & Right(str_Day, 2)), "yyyy-MM-dd")
                Exit For
            End If
        End If
        
    Next
    
    
    If lng_S = lng_len - 3 Then Exit Sub
    If str_���� = "" Then Exit Sub
    str�걾�� = Mid_bin(str_Data, 15, 13)
    strRecord = ""
    For int_i = 1 To 21
        strValues = Mid_bin(str_Data, 78 + int_i * 5, 5)
        If ((((int_i = 9) Or (int_i = 10)) Or (int_i = 11)) Or (int_i = 15)) Then
            If InStr(strValues, "--") > 0 Then
                strValues = strValues
            Else
                strValues = Format(Val(strValues) * 100, "0.0")
            End If
        End If
        strRecord = strRecord & "|" & int_i & "|" & IIf(InStr(strValues, "--") > 0, strValues, Val(strValues))
    Next
    
    If Val(str�걾��) > 0 And str_���� <> "" And strRecord <> "" Then
        strResult = str_���� & "|" & Val(str�걾��) & "| |Ѫ��|0" & strRecord & vbCrLf
        
    
        str_WBC = "WBC;0;120,50,8.33-,16.66-100,24.99-,33.32-200,41.65-,49.98-300" & Graph_Text(Mid_bin(str_Data, 243, 50, False))
        str_RBC = "^RBC;0;120,50,10-,20-100,30-,40-200,50-" & Graph_Text(Mid_bin(str_Data, 293, 50, False))
        str_PLT = "^PLT;0;120,50,12.5-10,25-20,37.5-30,50-" & Graph_Text(Mid_bin(str_Data, 343, 40, False))
        strResult = strResult & str_WBC & str_RBC & str_PLT
        
        'strCmd = Chr(&H6)
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