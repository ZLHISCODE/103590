VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_RAC_050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:ȫ�Զ���Ѫ������
'����:RAC-050
'����:������
'�û�:̫ԭ������ҽԺ
'ʱ��:2013-1-30

Private mstrCode As String    '��������
Private mstrLastData As String
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

    '���ݸ�ʽ��"����;ͼ������;Y�߶�,X����;�������ұ߿�����(���ڻ��̶�);X��̶�[|Y�̶�];����1����[|����2����...][;�������]"
    '���� �������ݣ���y�������ݣ���,�ָ�����������������|�ָ�
    '������ݣ���x�������ݣ���,�ŷָ�
    Dim strDate As String
    Dim strNo As String, strNow As String
    Dim strItems As String, strItem As String
    Dim varItem As Variant
    Dim strLine As String
    Dim str_Sample As String
    Dim lng_S As Long
    Dim lng_E As Long

    On Err GoTo errHandle
    If mstrLastData <> "" Then
        mstrCode = mstrLastData & mstrCode & strSource    '���滺�������
        mstrLastData = ""
    Else
        mstrCode = mstrCode & strSource
    End If
    lng_E = 0: lng_S = InStr(mstrCode, "R")
    If lng_S > 0 Then lng_E = InStr(mstrCode, "FINISH")
    If lng_E = 0 Then lng_E = Len(mstrCode)
    Do While lng_S > 0 And lng_E > 0
        lng_E = InStr(mstrCode, "0R")
        If lng_E = 0 Then lng_E = Len(mstrCode)
        str_Sample = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        varItem = Split(str_Sample, ",")
        If IsNumeric(varItem(1)) = True And IsDate(varItem(2)) = True Then
            strNow = varItem(1)
            strDate = varItem(2)
            If InStr(varItem(11), "PT") > 0 And InStr(varItem(11), "APTT") = 0 Then
                strItem = varItem(11) & "|" & varItem(12) & "|INR|" & varItem(14)
            ElseIf InStr(varItem(11), "D-Dimer") > 0 Or InStr(varItem(11), "FIB") > 0 Then
                strItem = varItem(11) & "|" & varItem(15)
            Else
                strItem = varItem(11) & "|" & varItem(12)
            End If
        End If
        
        
        If strNo = "" Or strNo = strNow Then
            strItems = strItems & "|" & strItem
            strNo = strNow
        Else
            strResult = strResult & "||" & strDate & "|" & strNo & "| |Ѫ��|0" & strItems
            strItems = ""
            strItems = strItems & "|" & strItem
            strNo = strNow
        End If
        

        lng_E = 0: lng_S = InStr(mstrCode, "R")
        If lng_S > 0 Then lng_E = InStr(mstrCode, "0R")
        If InStr(mstrCode, "0R") = 0 And mstrCode <> "" Then
            mstrLastData = mstrCode
            mstrCode = ""
            If UBound(Split(mstrLastData, ",")) = 17 Then
                If Split(mstrLastData, ",")(17) <> "" Then
                    varItem = Split(mstrLastData, ",")
                    If IsNumeric(varItem(1)) = True And IsDate(varItem(2)) = True Then
                        strNow = varItem(1)
                        strDate = varItem(2)
                        If InStr(varItem(11), "PT") > 0 And InStr(varItem(11), "APTT") = 0 Then
                            strItem = varItem(11) & "|" & varItem(12) & "|INR|" & varItem(14)
                        ElseIf InStr(varItem(11), "D-Dimer") > 0 Or InStr(varItem(11), "FIB") > 0 Then
                            strItem = varItem(11) & "|" & varItem(15)
                        Else
                            strItem = varItem(11) & "|" & varItem(12)
                        End If
                    End If
                    
                    
                    If strNo = "" Or strNo = strNow Then
                        strItems = strItems & "|" & strItem
                        strNo = strNow
                    End If
                    If strItems <> "" And Val(strNo) > 0 Then
                        strResult = strResult & "||" & strDate & "|" & strNo & "| |Ѫ��|0" & strItems
                    End If
                            
                End If
            End If
        End If
    Loop
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
    End If
    Exit Sub
errHandle:
    WriteLog "Test_analyse", strSource, Err.Number & " " & Err.Description
End Sub









