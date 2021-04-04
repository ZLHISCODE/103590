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
'���ܣ�Json������Javascript��ʼ��
'������strPreNode=��Ҫȱʡ���ӵ�Jsonǰ׺��㣬��input,output
    
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
'���ܣ�����Json����ֵ��ֻ֧��String�����ֺ�Nullֵ
    
    On Error GoTo errHand
    
    If IsNull(varNewValue) Then
        'ע�������������ֵisNull
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
'���ܣ���ȡJsonָ������ֵ
'������
'  strElement=��㼰·�����磺output.message��output.pati_list[0].phone_number,output.num_list
'  varNullValue=�����ֵΪΪnullʱ�����ص�ת��ֵ
'���أ���Ӧ���͵�Json����ֵ�������ֵ�б����飬�򷵻�ֵ�б���(�ַ�Ԫ��û������)����"111,aaa,222"
    Dim arrPath As Variant
    Dim varResult As Variant
    Dim i As Long
    
    varResult = Empty
        
    Err.Clear: On Error Resume Next
    varResult = mobjScriptControl.Eval("json." & strElement)
    If Err.Number <> 0 Then
        If Err.Number <> 1010 Then GoTo errHand
        Err.Clear: On Error GoTo errHand
        
        '��һЩjs�ؼ��ֲ���ֱ��·�����ʣ���Ϊ��ʽ��json['�����']['������'][x]['�����']
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
        'Variant���͵�ȱʡֵ����ΪEmpty��дʱ�����ⲿδ��ʱ�д�(������)��ԭ��δ֪
        'IsMissing������Variant���͵Ŀ�ѡ������׼ȷ�ж��Ƿ���ֵ
        If Not IsMissing(varNullValue) Then
            varResult = varNullValue
        End If
    ElseIf TypeName(varResult) = "String" Then
        'Do Nothing
    ElseIf TypeName(varResult) = "JScriptTypeInfo" Then
        '����򸸽��(object)����ֵ�б�����ʱ����"[object Object]"
        varResult = CStr(varResult)
    End If
    
    GetValue = varResult
    
    Exit Function
errHand:
    MsgBox "jsonֵ��ȡ����" & strElement & vbCrLf & vbCrLf & Err.Number & ":" & Err.Description, vbCritical, App.ProductName
End Function

Public Function GetList(ByVal strList As String, Optional ByVal strKeys As String, Optional ByVal varNullValue As Variant) As Collection
'���ܣ���ȡJson�е��������ݻ��ӽ�����ݵ�������
'������
'  strList=Json������򸸽������·�����磺output��output.pati_list��output.pati_list[0].baby_list
'  strKeys=��������Ϊ�ؼ��ֵĽ���������Զ����","�ŷָ�����"pati_id,pati_pageid"��ע��ؼ��ֽ������ݲ����������ظ�
'  varNullValue=�������еĽ��ֵΪΪnullʱ�����ص�ת��ֵ
'���أ�
'  ����������£�����һ������
'    ����ʱ����ָ����strKeys�����򼯺�Ԫ����������ָ������ֵΪ�ؼ��֣���ʽΪ"_Key1Value_Key2Value..."
'      ���ϵ�ÿһ��Ԫ��Ҳ��һ�����ϣ��ü�������������Ϊ�ؼ��֣���ʽ��"_pati_id"���ü��ϵ�Ԫ��ֵΪ��Ӧ���͵�Json���ֵ
'         ��������а��������飬�����ֵ�б����飬�򷵻ص������е�ֵΪֵ�б���(�ַ�Ԫ��û������)����"111,aaa,222",����Ƿ�ֵ�б����飬�򷵻ص������е�ֵΪ"[object Object]"
'    �����ʱ�����ӽ������Ϊ�ؼ��ּ��뼯��Ԫ�أ���"_message"��
'  ʧ�ܵ�����·���Nothing
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
        Set colList = New Collection '����Ԫ�ظ�������Ϊ0
        
        lngLen = -1
        If mobjScriptControl.Eval("json." & strList & " instanceof Array") Then
            lngLen = mobjScriptControl.Eval("json." & strList & ".length")
        End If
        If lngLen >= 0 Then '��������
            'ȡ�����
            strCode = _
                "var cElements = '';" & vbCrLf & _
                "for(var k in json." & strList & "[0]) {" & vbCrLf & _
                "  cElements = cElements + ',' + k;" & vbCrLf & _
                "};"
            mobjScriptControl.AddCode strCode
            strCode = mobjScriptControl.Eval("cElements.substr(1)")
            arrElement = Split(strCode, ",")
            
            '��������
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
        Else '���ӽ��ĸ����
            'ȡ�����
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
'���ܣ���ȡָ��Json�����Ԫ�ظ���
'������strElement=json������·�����磺output.pati_list��output.pati_list[0].baby_list
    
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