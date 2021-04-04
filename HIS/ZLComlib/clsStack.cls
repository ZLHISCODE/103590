VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'========================================================
'功能：该类是堆栈的基本类
'编者：lshuo,2017-8-2
'========================================================
'========================================================
'==变量缓存
'========================================================
Private marrStack()                 As Variant  '当前堆栈

Private mlngFrontIndex              As Long     '堆栈的尾部
Private mlngBaseIndex               As Long     '堆栈的头部

Private mlngStackSize               As Long     '堆栈最大尺寸

Private Const M_LNG_AUTOINCREMENT   As Long = 20         '自增长的基础量
'========================================================
'==属性
'========================================================
'StackSize:(R/W)堆栈大小
Public Property Get StackSize() As Long
    StackSize = mlngStackSize
End Property

Public Property Let StackSize(ByVal lngStackSize As Long)
    If lngStackSize > 0 Then
        Call InitStack(lngStackSize)
    End If
End Property
'Count:(R)堆栈元素个数
Public Property Get Count() As Long
    If mlngFrontIndex >= mlngBaseIndex Then
        Count = mlngFrontIndex - mlngBaseIndex + 1
    End If
End Property

'--------------------------------------------------------------------------------------------------
'接口           GetStack
'功能           获得堆栈数组
'返回值         Variant
'-------------------------------------------------------------------------------------------------
Public Property Get Data() As Variant
    Dim arrRet          As Variant
    Dim i               As Long
    
    If mlngFrontIndex >= mlngBaseIndex Then
        ReDim marrRetStack(mlngFrontIndex)
        For i = mlngFrontIndex To mlngBaseIndex Step -1
            Call gobjComLib.zlCommFun.AssignVariant(arrRet(mlngFrontIndex - i), marrStack(i))
        Next
    End If
    
    Data = arrRet
End Property

'--------------------------------------------------------------------------------------------------
'接口           InitStack
'功能
'返回值
'入参列表:
'参数名         类型                    说明
'lngStackSize   Long                    堆栈大小
'-------------------------------------------------------------------------------------------------
Public Sub InitStack(ByVal lngStackSize As Long)
    mlngStackSize = lngStackSize
    ReDim marrStack(mlngStackSize - 1)
    mlngBaseIndex = 0
    mlngFrontIndex = -1
End Sub

'--------------------------------------------------------------------------------------------------
'接口           Push
'功能           元素进入堆栈
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'varStackItem   Variant                 进入堆栈的元素
'-------------------------------------------------------------------------------------------------
Public Function Push(ByVal varStackItem As Variant) As Boolean
    If Me.IsFull Then
        Call StackAutoIncrement
    End If
    mlngFrontIndex = mlngFrontIndex + 1
    Call gobjComLib.zlCommFun.AssignVariant(marrStack(mlngFrontIndex), varStackItem)
    Push = True
End Function
'--------------------------------------------------------------------------------------------------
'接口           Pop
'功能           弹出堆栈顶端元素
'lngTimes       弹出几次，多次时，此时返回最后一次的。
'返回值         Variant
'-------------------------------------------------------------------------------------------------
Public Function Pop(Optional ByVal lngTimes As Long = 1) As Variant
    If Me.IsEmpty Then
        Pop = Empty
        Exit Function
    Else
        If lngTimes < 1 Then Pop = Empty: Exit Function
        If lngTimes >= Me.Count Then
            Pop = Empty
        Else
            Call gobjComLib.zlCommFun.AssignVariant(Pop, marrStack(mlngFrontIndex - lngTimes + 1))
        End If
        mlngFrontIndex = mlngFrontIndex - lngTimes
    End If
End Function
'--------------------------------------------------------------------------------------------------
'接口           Rank
'功能           获取堆栈制定位置的一个元素，从头到尾数,以0索引
'返回值         Variant
'-------------------------------------------------------------------------------------------------
Public Function Rank(ByVal lngIndex As Long) As Variant
    If Me.IsEmpty Then
        Rank = Empty
        Exit Function
    End If
    If lngIndex >= Me.Count Then
        Rank = Empty
        Exit Function
    End If
    Call gobjComLib.zlCommFun.AssignVariant(Rank, marrStack(mlngFrontIndex - lngIndex))
End Function

'--------------------------------------------------------------------------------------------------
'接口           Current
'功能           获取堆栈当前元素
'返回值         Variant
'-------------------------------------------------------------------------------------------------
Public Function Current() As Variant
    If Me.IsEmpty Then
        Current = Empty
        Exit Function
    End If
    Call gobjComLib.zlCommFun.AssignVariant(Current, marrStack(mlngFrontIndex))
End Function

'--------------------------------------------------------------------------------------------------
'接口           IsFull
'功能           判断堆栈是否是满的
'返回值         Boolean
'-------------------------------------------------------------------------------------------------
Public Function IsFull() As Boolean
    IsFull = (mlngFrontIndex + 1 = mlngStackSize)
End Function

'--------------------------------------------------------------------------------------------------
'接口           IsEmpty
'功能           判断堆栈是否是空的
'返回值         Boolean
'-------------------------------------------------------------------------------------------------
Public Function IsEmpty() As Boolean
    IsEmpty = mlngFrontIndex < mlngBaseIndex And Not Me.IsFull
End Function

'--------------------------------------------------------------------------------------------------
'接口           Clear
'功能           清理堆栈(仅移动指针，不销毁数据）
'-------------------------------------------------------------------------------------------------
Public Sub Clear()
    mlngFrontIndex = -1
    mlngBaseIndex = 0
End Sub
'--------------------------------------------------------------------------------------------------
'接口           TearDown
'功能           销毁堆栈（清理堆栈的数据，彻底销毁）
'-------------------------------------------------------------------------------------------------
Public Sub TearDown()
    Erase marrStack
    mlngFrontIndex = -1
    mlngBaseIndex = 0
End Sub

'--------------------------------------------------------------------------------------------------
'接口           ToString
'功能           展示堆栈
'返回值         String
'-------------------------------------------------------------------------------------------------
Public Function ToString() As String
    ToString = gobjComLib.zlStr.DisPlayOneValue(Me.Data)
End Function

'--------------------------------------------------------------------------------------------------
'接口           StackAutoIncrement
'功能           自动扩张队列
'返回值         Boolean
'-------------------------------------------------------------------------------------------------
Private Function StackAutoIncrement() As Boolean
    mlngStackSize = mlngStackSize + M_LNG_AUTOINCREMENT
    ReDim Preserve marrStack(mlngStackSize - 1)
    StackAutoIncrement = True
End Function

Private Sub Class_Initialize()
    mlngFrontIndex = -1
    mlngBaseIndex = 0
End Sub

Private Sub Class_Terminate()
    Erase marrStack
End Sub

