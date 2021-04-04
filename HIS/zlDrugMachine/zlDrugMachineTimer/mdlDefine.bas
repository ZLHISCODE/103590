Attribute VB_Name = "mdlDefine"
Option Explicit

Public Type TYPE_PARAMS
    ��ʱ���� As Integer
    ��Ч���� As Integer
    ��ʾ������� As Integer
    �����־ As Boolean
    ��ϸ��־ As Boolean
    ������־���� As Integer
    ҵ������ As String
End Type

Public Const GSTR_MSG As String = "��ʱ����"

Public Function GetParameter(ByVal objXML As clsXML, ByVal strName As String, Optional ByVal strDefaultVal As String) As String
'���ܣ���zlDrugMachine.cfg�ļ��л�ȡָ��������ֵ
'������
'  objXML��cfg�ļ������ݼ��غ��XML����
'  strName���������ƣ�����XML�������
'���أ�����ֵ

    Dim strValue As String

    If objXML Is Nothing Then
        GetParameter = strDefaultVal
        Exit Function
    End If
    
    strName = LCase(strName)
    
    If objXML.GetSingleNodeValue(strName, strValue) Then
        GetParameter = strValue
    Else
        GetParameter = strDefaultVal
    End If

End Function
