VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mobjScriptControl As Object
                 
Public Function OpenJson(ByVal strJson As String, Optional strPreNode As String) As Boolean
'功能：Json处理的Javascript初始化
'参数：strPreNode=需要缺省增加的Json前缀结点，如input,output
    
    On Error GoTo errHand

    If strJson = "" Then Exit Function
    
    mobjScriptControl.Language = "javascript"
    mobjScriptControl.Reset
     
    If strPreNode <> "" Then
        strJson = "{" & """" & strPreNode & """" & ":" & strJson & "};"
    End If

    Call mobjScriptControl.AddCode("var json=" & strJson)
    
    OpenJson = True
    
    Exit Function
errHand:
    MsgBox Err.Number & ":" & Err.Description, vbCritical, App.ProductName
End Function

Public Function SetValue(ByVal strElement As String, ByVal varNewValue As Variant) As Variant
'功能：设置Json结点的值，只支持String，数字和Null值
    
    On Error GoTo errHand
    
    If IsNull(varNewValue) Then
        '注意这种情况返回值isNull
        SetValue = mobjScriptControl.Eval("json." & strElement & "=null")
    ElseIf TypeName(varNewValue) = "String" Then
        SetValue = mobjScriptControl.Eval("json." & strElement & "='" & varNewValue & "'")
    Else
        SetValue = mobjScriptControl.Eval("json." & strElement & "=" & varNewValue)
    End If
    
    Exit Function
errHand:
    MsgBox Err.Number & ":" & Err.Description, vbCritical, App.ProductName
End Function

Public Function GetValue(ByVal strElement As String, Optional ByVal varNullValue As Variant) As Variant
'功能：获取Json指定结点的值
'参数：
'  strElement=结点及路径，如：output.message，output.pati_list[0].phone_number,output.num_list
'  varNullValue=当结点值为为null时，返回的转换值
'返回：对应类型的Json结点的值；如果是值列表数组，则返回值列表串(字符元素没有引号)，如"111,aaa,222"
    Dim arrPath As Variant
    Dim varResult As Variant
    Dim i As Long
    
    varResult = Empty
        
    Err.Clear: On Error Resume Next
    varResult = mobjScriptControl.Eval("json." & strElement)
    If Err.Number <> 0 Then
        If Err.Number <> 1010 Then GoTo errHand
        Err.Clear: On Error GoTo errHand
        
        '对一些js关键字不能直接路径访问，改为方式：json['结点名']['对象名'][x]['结点名']
        arrPath = Split(strElement, ".")
        strElement = ""
        For i = 0 To UBound(arrPath)
            If InStr(arrPath(i), "[") > 0 Then
                strElement = strElement & "['" & Left(arrPath(i), InStr(arrPath(i), "[") - 1) & "']" & Mid(arrPath(i), InStr(arrPath(i), "["))
            Else
                strElement = strElement & "['" & arrPath(i) & "']"
            End If
        Next
        varResult = mobjScriptControl.Eval("json" & strElement)
    Else
        On Error GoTo errHand
    End If
    
    If IsNull(varResult) Then
        'Variant类型的缺省值定义为Empty或不写时，类外部未传时有错(不触发)，原因未知
        'IsMissing函数对Variant类型的可选参数可准确判断是否传了值
        If Not IsMissing(varNullValue) Then
            varResult = varNullValue
        End If
    ElseIf TypeName(varResult) = "String" Then
        'Do Nothing
    ElseIf TypeName(varResult) = "JScriptTypeInfo" Then
        '数组或父结点(object)。非值列表数组时返回"[object Object]"
        varResult = CStr(varResult)
    End If
    
    GetValue = varResult
    
    Exit Function
errHand:
    MsgBox "json值获取错误：" & strElement & vbCrLf & vbCrLf & Err.Number & ":" & Err.Description, vbCritical, App.ProductName
End Function

Public Function GetList(ByVal strList As String, Optional ByVal strKeys As String, Optional ByVal varNullValue As Variant) As Collection
'功能：获取Json中的数组数据或子结点数据到集合中
'参数：
'  strList=Json数组结点或父结点名及路径，如：output，output.pati_list，output.pati_list[0].baby_list
'  strKeys=数组中作为关键字的结点名，可以多个用","号分隔，如"pati_id,pati_pageid"。注意关键字结点的数据不允许存在重复
'  varNullValue=当数组中的结点值为为null时，返回的转换值
'返回：
'  正常的情况下，返回一个集合
'    数组时，如指定了strKeys参数则集合元素以数组中指定结点的值为关键字，格式为"_Key1Value_Key2Value..."
'      集合的每一个元素也是一个集合，该集合以数组结点名为关键字，格式如"_pati_id"，该集合的元素值为对应类型的Json结点值
'         如果数组中包含子数组，如果是值列表数组，则返回到集合中的值为值列表串(字符元素没有引号)，如"111,aaa,222",如果是非值列表数组，则返回到集合中的值为"[object Object]"
'    父结点时，以子结点名作为关键字加入集合元素，如"_message"。
'  失败的情况下返回Nothing
    Dim colList As Collection
    Dim colRow As Collection
    Dim strCode As String
    Dim arrElement As Variant
    Dim arrKey As Variant
    Dim strKey As String
    Dim i As Long, j As Long
    Dim lngLen As Long
    
    On Error GoTo errHand
    
    If mobjScriptControl.Eval("typeof(json." & strList & ")") = "object" Then
        Set colList = New Collection '集合元素个数可能为0
        
        lngLen = -1
        If mobjScriptControl.Eval("json." & strList & " instanceof Array") Then
            lngLen = mobjScriptControl.Eval("json." & strList & ".length")
        End If
        If lngLen >= 0 Then '多结点数组
            '取结点名
            strCode = _
                "var cElements = '';" & vbCrLf & _
                "for(var k in json." & strList & "[0]) {" & vbCrLf & _
                "  cElements = cElements + ',' + k;" & vbCrLf & _
                "};"
            mobjScriptControl.AddCode strCode
            strCode = mobjScriptControl.Eval("cElements.substr(1)")
            arrElement = Split(strCode, ",")
            
            '产生数据
            If strKeys <> "" Then
                arrKey = Split(strKeys, ",")
            End If
            For i = 0 To lngLen - 1
                Set colRow = New Collection
                For j = 0 To UBound(arrElement)
                    colRow.Add GetValue(strList & "[" & i & "]." & arrElement(j), varNullValue), "_" & arrElement(j)
                Next
                
                If strKeys <> "" Then
                    strKey = ""
                    For j = 0 To UBound(arrKey)
                        strKey = strKey & "_" & GetValue(strList & "[" & i & "]." & arrKey(j), varNullValue)
                    Next
                    colList.Add colRow, strKey
                Else
                    colList.Add colRow
                End If
            Next
        Else '多子结点的父结点
            '取结点名
            strCode = _
                "var cElements = '';" & vbCrLf & _
                "for(var k in json." & strList & ") {" & vbCrLf & _
                "  cElements = cElements + ',' + k;" & vbCrLf & _
                "};"
            mobjScriptControl.AddCode strCode
            strCode = mobjScriptControl.Eval("cElements.substr(1)")
            
            If strCode <> "" Then
                arrElement = Split(strCode, ",")
    
                For i = 0 To UBound(arrElement)
                    colList.Add GetValue(strList & "." & arrElement(i), varNullValue), "_" & arrElement(i)
                Next
            End If
        End If
    End If
    
    Set GetList = colList
    
    Exit Function
errHand:
    MsgBox Err.Number & ":" & Err.Description, vbCritical, App.ProductName
End Function

Public Function GetLength(ByVal strElement As String) As Long
'功能：获取指定Json数组的元素个数
'参数：strElement=json数组名路径，如：output.pati_list，output.pati_list[0].baby_list
    
    On Error GoTo errHand
    
    GetLength = mobjScriptControl.Eval("json." & strElement).length
    
    Exit Function
errHand:
    MsgBox Err.Number & ":" & Err.Description, vbCritical, App.ProductName
End Function

Private Sub Class_Initialize()
    Set mobjScriptControl = CreateObject("ScriptControl")
End Sub

Private Sub Class_Terminate()
    Set mobjScriptControl = Nothing
End Sub
