Attribute VB_Name = "mdlPubJson"
Option Explicit
'JSON�ڵ�����
Public Enum JSON_TYPE
    Json_Text = 0 '�ַ�
    Json_num = 1 '��ֵ
End Enum


Public Function zlGetNodeValueFromCollect(ByVal clldata As Collection, ByVal strKey As String, ByVal strType As String) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ڵ�����ݼ�
    '���:cllData-��ǰ������
    '     strKey-Key
    '     strType-"N"-����;"C"�ַ�
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-14 16:20:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    Err = 0: On Error Resume Next
    varTemp = clldata(strKey)
    If Err <> 0 Then
        Err = 0: On Error GoTo 0
        If strType = "N" Then zlGetNodeValueFromCollect = Empty: Exit Function
        zlGetNodeValueFromCollect = "": Exit Function
    End If
    zlGetNodeValueFromCollect = varTemp
End Function

Public Function zlGetNodeObjectFromCollect(ByVal clldata As Collection, ByVal strKey As String) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ڵ�Ķ���
    '���:cllData-��ǰ������
    '     strKey-Key
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-14 16:20:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllTemp As Collection
    Err = 0: On Error Resume Next
    
    Set cllTemp = clldata(strKey)
    If Err <> 0 Then
        Err = 0: On Error GoTo 0
       Set zlGetNodeObjectFromCollect = cllTemp
       Exit Function
    End If
    Set zlGetNodeObjectFromCollect = cllTemp
End Function


Public Function GetJsonNodeString(ByVal strNodeName As String, ByVal strValue As String, _
    Optional ByVal intType As JSON_TYPE, Optional ByVal blnZeroToNull As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡJson�ӵ㴮
    '���:strNodeName-�ӵ���
    '     strValue-ֵ
    '     intType-����:0-�ַ�;1-����
    '     blnZeroToEmpty-�Ƿ���ֵ0ת��ΪNull��������Ϊ����ʱ��Ч
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-08-09 18:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String
    strJson = Chr(34) & strNodeName & Chr(34)
    If intType = Json_Text Then
        strJson = strJson & ":" & Chr(34) & gobjComLib.zlStr.ToJsonStr(strValue) & Chr(34)
    Else
        If strValue = "" Or (blnZeroToNull And Val(strValue) = 0) Then
            strJson = strJson & ":null"
        Else
            strJson = strJson & ":" & IIf(Mid(strValue, 1, 1) = ".", "0", "") & strValue
        End If
    End If
    GetJsonNodeString = strJson
End Function
Public Function GetCollValue(ByVal colValue As Collection, ByVal varRow As Variant, Optional ByVal strElement As String) As Variant
    '���ܣ���ȡJson���鷵�صļ���������ָ���л�ָ��Ԫ�ص�ֵ
    '������
    '  varRow=���������йؼ���
    '  strElement=Ԫ����
    '���أ�
    '  ��δ����strElement����ʱ������ָ���еļ��϶��󣻵�����strElement����ʱ������ָ����ָ��Ԫ�ص�ֵ
    '  ʧ��ʱ����Nothing��Empty�������ᱨ��
    If strElement <> "" Then
        GetCollValue = Empty
    Else
        Set GetCollValue = Nothing
    End If
    
    If colValue Is Nothing Then Exit Function
    
    On Error Resume Next
    If strElement <> "" Then
        GetCollValue = colValue(varRow)(strElement)
    Else
        Set GetCollValue = colValue(varRow)
    End If
    Err.Clear: On Error GoTo 0
End Function

Public Function CollectionExitsValue(ByVal coll As Collection, _
    ByVal strKey As String) As Boolean
    '���ݹؼ����ж�Ԫ���Ƿ�����ڼ�����
    Dim blnExits As Boolean

    If coll Is Nothing Then Exit Function
    CollectionExitsValue = True
    Err = 0: On Error Resume Next
    blnExits = IsObject(coll(strKey))
    If Err <> 0 Then Err = 0: CollectionExitsValue = False
End Function


Public Function GetNodeString(ByVal strNodeName As String) As String
    GetNodeString = Chr(34) & strNodeName & Chr(34)
End Function


