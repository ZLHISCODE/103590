Attribute VB_Name = "mdlMidWife"
Option Explicit
Public gcnOracle As ADODB.Connection

Public glngInstance As Long  '���ʵ������
Public gstrUser As String    '��ǰ��¼����̨���û�
Public gstrURL As String     '������ҳ��URL
Public gstrURLLogin As String '��¼������Ϣ������ĳ�ʼURL
Public glngPatiID As Long, glngPageID As Long '��¼��һ�β��˵Ĳ���id����ҳid

Public Function GetEncrypt(ByVal strCode As String) As String
'���ܣ�����Url�����ܷ�����
    Dim tmp() As Byte, strResult As String
    Dim i As Integer
    
    tmp = StrConv(strCode, vbFromUnicode)
    strResult = URLEncode(tmp(0) + UBound(tmp) + 1)
    
    For i = 1 To UBound(tmp)
        strResult = strResult & URLEncode(tmp(i) + tmp(i - 1))
    Next
    GetEncrypt = strResult
End Function

Public Function URLEncode(ByVal intValue As Integer) As String
    Dim s As String
    
    If (intValue >= 48 And intValue <= 57) Or (intValue >= 65 And intValue <= 90) Or (intValue >= 97 And intValue <= 122) Then
      s = Chr(intValue)
    ElseIf intValue = 32 Then
      s = "+"
    Else
      s = "%" & Hex(intValue)
    End If

    URLEncode = s
End Function

Public Function GetSysPar(lngPar As Long, lngSys As Long) As String
    Dim strSql As String, rstmp As ADODB.Recordset
    
    On Error GoTo errHandle
    'ִ�У�Set cmdData.ActiveConnection = gcnOracle
    '�ᱨ���������Ͳ���ȷ����������ActiveExe��ʽ�Դ�������Ӷ�����ʲô���ƣ����Բ���ʹ�ù�����������OpenSqlRecord
    
    strSql = "Select Nvl(����ֵ,ȱʡֵ) as ����ֵ From zlParameters Where ������= " & lngPar & " And ϵͳ = " & lngSys & " And Nvl(ģ��,0)=0"
    Set rstmp = gcnOracle.Execute(strSql)
    If rstmp.RecordCount > 0 Then GetSysPar = "" & rstmp!����ֵ

    Exit Function
errHandle:
    MsgBox Err.Description & strSql, vbExclamation, "��ȡ����"
    GetSysPar = ""
End Function




