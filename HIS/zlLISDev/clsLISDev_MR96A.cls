VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MR96A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: MR-96A ø����
'����: ����
'����: �¶�
'�û�: ɽ������������
Private mstrCode As String
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
    Dim lng_S As Long, lng_E As Long, str_Code As String
    Dim var_Code As Variant, str_Line As String, i As Integer
    Dim str_NO As String, str_Date As String, str_Item As String, str_Value As String
    Dim str_Record As String, str_QC As String
    On Error GoTo hErr
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    If InStr(strSource, Chr(&H1C) & Chr(&HD)) > 0 Then strCmd = ""

    str_Code = PStr_CutCode(mstrCode, Chr(&HB), Chr(&H1C) & Chr(&HD))
    
    Do While str_Code <> ""
        '--------------------------------
        str_Record = "": str_QC = "0"
        var_Code = Split(str_Code, Chr(&HD))
        For i = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(i)
            If str_Line Like "*MSH|*|*|*|*|*|*|*|ORU^R01|*|*|*|*|*|*|*|*|*" Then
                str_QC = Split(str_Line, "|")(15)
                
                strCmd = Chr(&HB) & "MSH|^~\&|||||" & Format(Now, "yyyyMMddHHmmss") & "||ACK^R01|7|P|2.3.1||||" & str_QC & "||ASCII|||" & Chr(&HD)
                strCmd = strCmd & "MSA|AA|7|OK|||0|" & Chr(&HD) & Chr(&H1C) & Chr(&HD)
                
                If str_QC = "2" Then
                    str_QC = "1"
                ElseIf str_QC <> "0" Then
                    Exit For    '�������ݣ����������˳�
                End If
                
            ElseIf str_Line Like "OBR|*|*|*|*|*" Then
                If str_QC = "0" Then
                    str_NO = Trim("" & Split(str_Line, "|")(2))
                    str_Date = Trim("" & Split(str_Line, "|")(4))
                    If str_Date Like "##############" Then
                        str_Date = Mid$(str_Date, 1, 4) & "-" & Mid(str_Date, 5, 2) & "-" & Mid(str_Date, 7, 2) & " " & Mid(str_Date, 9, 2) & ":" & Mid(str_Date, 11, 2) & ":" & Mid(str_Date, 13, 2)
                        If Not IsDate(str_Date) Then str_Date = Format(Now, "yyyy-MM-dd HH:mm")
                    End If
                    
                    If str_Record <> "" Then strResult = strResult & "||" & str_Record
                    If str_NO <> "" And IsDate(str_Date) Then
                        str_Record = str_Date & "|" & str_NO & "| |Blod|" & str_QC
                    Else
                        str_Record = ""
                    End If
                Else
                    str_Record = ""
                End If
                
            ElseIf str_Line Like "OBX|*|*|*|*|*|*" Then
                If str_QC = "0" Then
                    '��ͨ�걾
                    str_Value = Trim("" & Split(str_Line, "|")(5))
                    str_Item = Trim("" & Split(str_Line, "|")(4))
                    If str_Item <> "" And str_Value <> "" And str_Record <> "" Then str_Record = str_Record & "|" & str_Item & "|" & str_Value
                End If
            
            End If
        Next
        If str_Record <> "" Then strResult = strResult & "||" & str_Record
        
        '--------------------------------
        str_Code = PStr_CutCode(mstrCode, Chr(&HB), Chr(&H1C) & Chr(&HD))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
hErr:
    WriteLog "Mr96a.Analyse", CStr(Erl()) & "��", Err.Description
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




