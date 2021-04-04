VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CA7000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:CA7000 ��Ѫ����
'����:ϣɭ����
'����:������
'�û�:��ú������ҽԺ
'ʱ��:2015--08-18
Private mstrCode As String '��������
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

    Dim strDate As String
    Dim strNo As String
    Dim strItems As String
    Dim varItem As Variant
    Dim strItem As String
    Dim strLine As String
    Dim str_Sample As String
    Dim lng_S As Long
    Dim lng_E As Long
    Dim intLoop As Integer
    Dim strAllItem As String
        
    On Err GoTo errHandle
    mstrCode = mstrCode & strSource    '���滺�������
    'Ӧ����
    
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, vbCrLf) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        strCmd = Chr(&H6)
    End If

    lng_E = 0: lng_S = InStr(mstrCode, Chr(&H5)) '"1H"
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4)) '"07" & Chr(&HD) & Chr(&HA))
    
    Do While lng_S > 0 And lng_E > 0
        str_Sample = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        varItem = Split(str_Sample, Chr(&HD) & Chr(&HA))
        For intLoop = LBound(varItem) To UBound(varItem)
            strLine = Trim(varItem(intLoop))
            
            If strLine <> "" Then
                If InStr(Split(strLine, "|")(0), "R") > 0 Then
                    
                    strDate = Split(strLine, "|")(12)
                    strDate = strDate
                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2)
                    strDate = Format(strDate, "yyyy-mm-dd hh:mm:ss")
                    
                    
                    strItem = Split(strLine, "|")(2)
                    strItem = Trim(Split(strItem, "^")(3))
                    strItem = strItem & "|" & Trim(Split(strLine, "|")(3))
                    
                    If strItem <> "" Then
                        strItems = strItems & "|" & strItem
                    End If
                    
                ElseIf InStr(Split(strLine, "|")(0), "O") > 0 Then
                    strResult = ""
                     strNo = Split(strLine, "|")(3)
                     strNo = Trim(Split(strNo, "^")(2))
                
                    
                End If
            End If
        Next
        
       
            
    lng_E = 0: lng_S = InStr(mstrCode, Chr(&H5)) '"1H"
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    Loop
    
    If strDate <> "" And strNo <> "" And strItems <> "" Then
           strResult = strResult & "||" & strDate & "|" & Val(strNo) & "| |ѪҺ|0" & strItems
        End If
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
    End If
errHandle:
    WriteLog "Test_analyse", strSource, Err.Number & " " & Err.Description
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











