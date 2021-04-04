Attribute VB_Name = "mdlPubJson"
Option Explicit
'JSON节点类型
Public Enum JSON_TYPE
    Json_Text = 0 '字符
    Json_num = 1 '数值
End Enum


Public Function zlGetNodeValueFromCollect(ByVal clldata As Collection, ByVal strKey As String, ByVal strType As String) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定节点的数据集
    '入参:cllData-当前个集合
    '     strKey-Key
    '     strType-"N"-数字;"C"字符
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-14 16:20:53
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
    '功能:获取指定节点的对象集
    '入参:cllData-当前个集合
    '     strKey-Key
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-14 16:20:53
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
    '功能:获取Json接点串
    '入参:strNodeName-接点名
    '     strValue-值
    '     intType-类型:0-字符;1-数字
    '     blnZeroToEmpty-是否将数值0转换为Null，仅类型为数字时有效
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-09 18:59:04
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
    '功能：获取Json数组返回的集合数据中指定行或指定元素的值
    '参数：
    '  varRow=行索引或行关键字
    '  strElement=元素名
    '返回：
    '  当未传入strElement参数时，返回指定行的集合对象；当传入strElement参数时，返回指定行指定元素的值
    '  失败时返回Nothing或Empty，但不会报错
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
    '根据关键字判断元素是否存在于集合中
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


