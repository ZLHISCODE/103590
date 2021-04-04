VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSqlBracket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Depth As Long            '深度
Public DepthTag As String       '深度标记
        
Public Start As Long       '开始位置
Public Context As String

Public OrCount As Long
Public AndCount As Long
Public StartCount As Long
Public BetweenCount As Long

Public HasBetween As Boolean

Public IsSelect As Boolean
Public HasSubSelect As Boolean

Public IsParameter As Boolean
Public HasSubParameter As Boolean

Private mstrParNames() As String
Private maryParReplace() As String


Public Parent As clsSqlBracket
Private mobjSubItems()  As clsSqlBracket
 
'参数名称
Property Get ParNameCount() As Long
    ParNameCount = UBound(mstrParNames)
End Property

Property Get ParNames(ByVal lngIndex As Long) As String
    ParNames = mstrParNames(lngIndex)
End Property
 
Public Sub AddParName(ByVal strParName As String)
    Dim lngBound As Long
    
    lngBound = UBound(mstrParNames) + 1
    ReDim Preserve mstrParNames(lngBound)
    
    mstrParNames(lngBound) = strParName
End Sub

'参数号
Property Get ParReplaceCount() As Long
    ParReplaceCount = UBound(maryParReplace)
End Property

Property Get ParReplace(ByVal lngIndex As Long) As String
    ParReplace = maryParReplace(lngIndex)
End Property
 
Public Sub AddParReplace(ByVal strParReplace As String)
    Dim lngBound As Long
    
    lngBound = UBound(maryParReplace) + 1
    ReDim Preserve maryParReplace(lngBound)
    
    maryParReplace(lngBound) = strParReplace
End Sub

Public Sub AddParLink(ByVal strParName As String, ByVal strParReplaceValue As String)
    Dim lngBound As Long
    
    lngBound = UBound(mstrParNames) + 1
    ReDim Preserve mstrParNames(lngBound)
    ReDim Preserve maryParReplace(lngBound)
    
    mstrParNames(lngBound) = strParName
    maryParReplace(lngBound) = strParReplaceValue
End Sub
 
'子括弧
Property Get SubItemCount() As Long
    SubItemCount = UBound(mobjSubItems)
End Property

Property Get SubItems(ByVal lngIndex As Long) As clsSqlBracket
    Set SubItems = mobjSubItems(lngIndex)
End Property


Public Sub AddSubItems(objBrack As clsSqlBracket)
    Dim lngBound As Long
    
    lngBound = UBound(mobjSubItems) + 1
    ReDim Preserve mobjSubItems(lngBound)
    
    Set mobjSubItems(lngBound) = objBrack
End Sub


Public Function GetBracket(ByVal strDepthTag As String) As clsSqlBracket
'根据深度标记获取对应的节点
    Dim strPath() As String
    Dim i As Long
    Dim objResult As clsSqlBracket
    
    Set GetBracket = Nothing
    
    strPath = Split(strDepthTag, "#")
    
    Set objResult = Nothing
    For i = 0 To UBound(strPath)
        If i = 0 Then
            Set objResult = Me
        Else
            Set objResult = objResult.SubItems(Val(strPath(i)))
        End If
    Next i
    
    Set GetBracket = objResult
End Function


Private Sub Class_Initialize()
'初始化

    Depth = 0
    DepthTag = "0"
    
    Start = 0
    
    OrCount = 0
    AndCount = 0
    StartCount = 0
    BetweenCount = 0
    
    HasBetween = False
    
    IsSelect = False
    IsParameter = False
    
    HasSubParameter = False
    HasSubSelect = False
    
    Set Parent = Nothing
    
    ReDim mobjSubItems(0)
    ReDim mstrParNames(0)
    ReDim maryParReplace(0)
End Sub
