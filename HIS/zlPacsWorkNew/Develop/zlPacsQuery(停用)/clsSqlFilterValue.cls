VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSqlFilterValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mobjParDatas() As FilterData

Private mblnIsUpdate As Boolean

Public Property Get Count() As Long
    Count = UBound(mobjParDatas)
End Property

Public Property Get IsUpdate() As Boolean
    IsUpdate = mblnIsUpdate
End Property

Public Property Let IsUpdate(value As Boolean)
    mblnIsUpdate = value
End Property


Public Property Get Item(ByVal lngIndex As Long) As FilterData
    Item = mobjParDatas(lngIndex)
End Property

Public Property Get ParData(ByVal strName As String) As Variant
'获取参数
    Dim i As Long
    
    For i = 1 To UBound(mobjParDatas)
        If mobjParDatas(i).Name = strName Then
            ParData = mobjParDatas(i).value
            Exit Property
        End If
    Next i
    
End Property

Public Property Let ParData(ByVal strName As String, value As Variant)
'设置参数
'只有值改变参数值时mblnIsUpdate会设置为true，首次新增参数时mblnIsUpdate不会进行设置
    Dim i As Long
    Dim lngUbound As Long
    
    For i = 1 To UBound(mobjParDatas)
        If mobjParDatas(i).Name = strName Then
            If mobjParDatas(i).value <> value Then
                mobjParDatas(i).value = value
                mblnIsUpdate = True
            End If
            
            Exit Property
        End If
    Next i
    
    lngUbound = UBound(mobjParDatas) + 1
    ReDim Preserve mobjParDatas(lngUbound)
    
    mobjParDatas(lngUbound).Name = strName
    mobjParDatas(lngUbound).value = value

End Property


Public Sub UpdateParValue(ByVal strName As String, value As Variant)
'更新参数值
'新增参数或者参数值改变均会设置mblnIsUpdate为true
    Dim i As Long
    Dim lngUbound As Long
    
    For i = 1 To UBound(mobjParDatas)
        If mobjParDatas(i).Name = strName Then
            If mobjParDatas(i).value <> value Then
                mobjParDatas(i).value = value
                mblnIsUpdate = True
            End If
            
            Exit Sub
        End If
    Next i
    
    lngUbound = UBound(mobjParDatas) + 1
    ReDim Preserve mobjParDatas(lngUbound)
    
    mobjParDatas(lngUbound).Name = strName
    mobjParDatas(lngUbound).value = value
    
    mblnIsUpdate = True

End Sub


Public Function GetStorageString() As String
'获取存储字符
    Dim i As Long
    Dim strRes As String
    
    GetStorageString = ""
    strRes = ""
    
    For i = 1 To UBound(mobjParDatas)
        If strRes <> "" Then strRes = strRes & "[#]"
        strRes = strRes & mobjParDatas(i).Name & "[-]" & mobjParDatas(i).value
    Next i
    
    GetStorageString = strRes
End Function

Public Function LoadStorageString(ByVal strStorage As String) As Boolean
'载入存储字符
    Dim i As Long
    Dim lngUbound As Long
    Dim arySplit() As String
    
    ReDim mobjParDatas(0)
    arySplit = Split("[#]" & strStorage & "[#]", "[#]")
    
    LoadStorageString = False
    
    For i = 1 To UBound(arySplit)
        If arySplit(i) <> "" Then
            lngUbound = UBound(mobjParDatas) + 1
            ReDim Preserve mobjParDatas(lngUbound)
            
            mobjParDatas(lngUbound).Name = Split(arySplit(i), "[-]")(0)
            mobjParDatas(lngUbound).value = Split(arySplit(i), "[-]")(1)
        End If
    Next i
    
    If UBound(mobjParDatas) >= 1 Then LoadStorageString = True
End Function

Public Function HasName(ByVal strName As String) As Boolean
'判断名称是否存在
    Dim i As Long
    
    HasName = False
    For i = 1 To UBound(mobjParDatas)
        If mobjParDatas(i).Name = strName Then
            HasName = True
            Exit Function
        End If
    Next i
End Function

Public Sub ClearAll()
    ReDim mobjParDatas(0)
End Sub

Private Sub Class_Initialize()
    ReDim mobjParDatas(0)
    mblnIsUpdate = False
End Sub

