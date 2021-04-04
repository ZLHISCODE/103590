VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_BG_XX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ������ʷ����� BG-XX
'����:÷�ݿ���
'����:�µ���
'�û�:��¤������ҽԺ

Private mstrCode As String  '����������
Const mstrItem  As String = ",����,Hb,FIO2,K,Na,Cl,nCa,TCa,pH,PCO2,PO2,pH(tc),PCO2(tc),PO2(tc),HCO3,TCO2,BEb,SBC,BEecf,sO2,P50,AG,RI,A-aDO2,"

'--------------------------------------------------------------------------------------------------------------------
'�����Ǳ���Ľӿں���
Public Function GetAnswerCmd() As String
    '�Զ�Ӧ��ָ��
    GetAnswerCmd = ""
End Function

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
    Dim str_Code As String '1����������Ϣ
    Dim var_Code As Variant
    Dim lng_count As Long
    
    Dim str_Segment As String 'һ����Ϣ��(��chr(&HD)�ָ�)
    Dim var_Filed As Variant   '�ֶ�
    
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_�걾�� As String, str_���� As String
    Dim str_��Ŀ As String
    Dim str_��� As String
    Dim str_Return As String
        
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
   
    mstrCode = Replace(mstrCode, "Ѫ������ʷ�����", Chr(&HB))
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&HA) & Chr(&HA) & Chr(&HA) & Chr(&HA))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
                 
        var_Code = Split(str_Code, Chr(&HA))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If str_Segment Like "*20*-*-*:*" Then
                str_���� = Format(Trim(str_Segment), "yyyy-MM-dd HH:MM") & ":" & Format(Now, "SS")
            ElseIf InStr(str_Segment, "˳��") > 0 Then
                '�걾�ţ�����
                str_�걾�� = Val(Trim(Mid(str_Segment, InStr(str_Segment, "˳��") + 4)))
                str_Return = ""
            ElseIf InStr(str_Segment, "���£�") > 0 Then
                str_Return = str_Return & "|TW|" & Trim(Mid(str_Segment, InStr(str_Segment, "���£�") + 4)) & "��"
            ElseIf InStr(str_Segment, "Hb  :") > 0 Then
                str_Return = str_Return & "|Hb|" & Trim(Mid(str_Segment, InStr(str_Segment, ":") + 1))
            ElseIf InStr(str_Segment, "FIO2:") > 0 Then
                str_Return = str_Return & "|FIO2|" & Trim(Mid(str_Segment, InStr(str_Segment, ":") + 1))
            ElseIf (InStr(str_Segment, "K") > 0) And (InStr(str_Segment, "TCa") > 0) Then
                str_Return = str_Return & "|K|" & Trim(Mid(str_Segment, InStr(str_Segment, "K") + 2, 6))
                If InStr(str_Segment, "Na") > 0 Then
                   str_Return = str_Return & "|Na|" & Trim(Mid(str_Segment, InStr(str_Segment, "Na") + 3, 6))
                End If
                If InStr(str_Segment, "Cl") > 0 Then
                   str_Return = str_Return & "|Cl|" & Trim(Mid(str_Segment, InStr(str_Segment, "Cl") + 3, 6))
                End If
                If InStr(str_Segment, "nCa") > 0 Then
                   str_Return = str_Return & "|nCa|" & Trim(Mid(str_Segment, InStr(str_Segment, "nCa") + 4, 6))
                End If
                str_Return = str_Return & "|TCa|" & Trim(Replace(Mid(str_Segment, InStr(str_Segment, "TCa") + 5), "mmol/L", ""))
            ElseIf (InStr(str_Segment, "pH") > 0) And (InStr(str_Segment, "PO2") > 0) Then
                str_Return = str_Return & "|pH|" & Trim(Mid(str_Segment, InStr(str_Segment, "pH") + 3, 9))
                If InStr(str_Segment, "PCO2") > 0 Then
                   str_Return = str_Return & "|PCO2|" & Trim(Mid(str_Segment, InStr(str_Segment, "PCO2") + 5, 6))
                End If
                str_Return = str_Return & "|PO2|" & Trim(Replace(Mid(str_Segment, InStr(str_Segment, "PO2") + 4, 7), "mmHg", ""))
                str_Return = Replace(str_Return, Chr(&HA), "")
            ElseIf InStr(str_Segment, "pH(tc)") > 0 Then
                str_Return = str_Return & "|pH(tc)|" & Trim(Mid(str_Segment, InStr(str_Segment, "pH(tc)") + 7))
            ElseIf InStr(str_Segment, "PCO2(tc)") > 0 Then
                str_Return = str_Return & "|PCO2(tc)|" & Trim(Mid(str_Segment, InStr(str_Segment, "PCO2(tc)") + 9))
            ElseIf InStr(str_Segment, "PO2(tc)") > 0 Then
                str_Return = str_Return & "|PO2(tc)|" & Trim(Mid(str_Segment, InStr(str_Segment, "PO2(tc)") + 8))
            ElseIf InStr(str_Segment, "HCO3") > 0 Then
                str_Return = str_Return & "|HCO3|" & Trim(Mid(str_Segment, InStr(str_Segment, "HCO3") + 4))
            ElseIf InStr(str_Segment, "TCO2") > 0 Then
                str_Return = str_Return & "|TCO2|" & Trim(Mid(str_Segment, InStr(str_Segment, "TCO2") + 4))
            ElseIf InStr(str_Segment, "BEb") > 0 Then
                str_Return = str_Return & "|BEb|" & Trim(Mid(str_Segment, InStr(str_Segment, "BEb") + 3))
            ElseIf InStr(str_Segment, "SBC") > 0 Then
                str_Return = str_Return & "|SBC|" & Trim(Mid(str_Segment, InStr(str_Segment, "SBC") + 3))
            ElseIf InStr(str_Segment, "BEecf") > 0 Then
                str_Return = str_Return & "|BEecf|" & Trim(Mid(str_Segment, InStr(str_Segment, "BEecf") + 5))
            ElseIf InStr(str_Segment, "sO2") > 0 Then
                str_Return = str_Return & "|sO2|" & Trim(Mid(str_Segment, InStr(str_Segment, "sO2") + 3))
            ElseIf InStr(str_Segment, "P50") > 0 Then
                str_Return = str_Return & "|P50|" & Trim(Mid(str_Segment, InStr(str_Segment, "P50") + 3))
            ElseIf InStr(str_Segment, "AG") > 0 Then
                str_Return = str_Return & "|AG|" & Trim(Mid(str_Segment, InStr(str_Segment, "AG") + 2))
            ElseIf InStr(str_Segment, "RI") > 0 Then
                str_Return = str_Return & "|RI|" & Trim(Mid(str_Segment, InStr(str_Segment, "RI") + 2))
            ElseIf InStr(str_Segment, "A-aDO2") > 0 Then
                str_Return = str_Return & "|AaDO2|" & Trim(Mid(str_Segment, InStr(str_Segment, "A-aDO2") + 6))
            End If
        Next
        str_Return = Replace(Replace(str_Return, Chr(&HD), ""), Chr(&HA), "")
        str_Return = Replace(str_Return, Chr(&H20), "")
        str_Return = Replace(str_Return, "mmHg", "")
        str_Return = Replace(str_Return, "mmol/L", "")

        If Val(str_�걾��) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|0" & str_Return
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&HA) & Chr(&HA) & Chr(&HA) & Chr(&HA))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "BG_XX_Analyse", strSource, Err.Number & " " & Err.Description
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


