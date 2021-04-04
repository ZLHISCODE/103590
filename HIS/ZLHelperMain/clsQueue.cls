VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@模块 clsQueue-2019/6/26
'@编写 lshuo
'@功能
'   该类是队列的基本类
'@引用
'
'@备注
'
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
Private Const M_LNG_AUTOINCREMENT   As Long = 20            '自增长的基础量
Private Const M_LNG_INIT_QUEUESIZE  As Long = 5             '队列的初始大小
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
Private marrQueue()         As Variant  '当前队列
Private mlngFrontIndex      As Long     '队列的开头
Private mlngRearIndex       As Long     '队列的尾部

'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------
Private mlngQueueSize       As Long     '队列最大尺寸
Private mlngQueueCount      As Long     '当前队列的数据个数

'QueueSize:(R/W)队列大小
Public Property Get QueueSize() As Long
    QueueSize = mlngQueueSize
End Property

Public Property Let QueueSize(ByVal lngQueueSize As Long)
    If lngQueueSize > 0 Then
        Call InitQueue(lngQueueSize)
    End If
End Property
'Count:(R)队列元素个数
Public Property Get Count() As Long
    Count = mlngQueueCount
End Property
'Data(R):获得队列数组
Public Property Get Data() As Variant
    Dim i               As Long
    Dim arrRetQueue     As Variant
    Dim lngIndex        As Long
    
    If mlngQueueCount > 0 Then
        ReDim arrRetQueue(mlngQueueCount - 1)
        i = 0
        Do While i < mlngQueueCount
            lngIndex = (mlngFrontIndex + i) Mod mlngQueueSize
            If IsObject(marrQueue(lngIndex)) Then
                Set arrRetQueue(i) = marrQueue(lngIndex)
            Else
                arrRetQueue(i) = marrQueue(lngIndex)
            End If
            i = i + 1
        Loop
    End If
    Data = arrRetQueue
End Property
'Current(R/W):获取队列当前元素或更换队列的队首元素
Public Property Get Current() As Variant
    If Me.IsEmpty Then
        Current = Empty
        Exit Property
    End If
    If IsObject(marrQueue(mlngFrontIndex)) Then
        Set Current = marrQueue(mlngFrontIndex)
    Else
        Current = marrQueue(mlngFrontIndex)
    End If
End Property

Public Property Let Current(varTmp As Variant)
    If Me.IsEmpty Then
        Exit Property
    End If
    If IsObject(varTmp) Then
        Set marrQueue(mlngFrontIndex) = varTmp
    Else
        marrQueue(mlngFrontIndex) = varTmp
    End If
End Property

Public Property Set Current(objTmp As Object)
    If Me.IsEmpty Then
        Exit Property
    End If
    Set marrQueue(mlngFrontIndex) = objTmp
End Property
'Last(R/W):获取队列最后一个元素或替换最后的元素
Public Property Get Last() As Variant
    Dim lngIndex        As Long
    If Me.IsEmpty Then
        Last = Empty
        Exit Property
    End If
    
    lngIndex = (mlngRearIndex - 1 + mlngQueueSize) Mod mlngQueueSize
    If IsObject(marrQueue(lngIndex)) Then
        Set Last = marrQueue(lngIndex)
    Else
        Last = marrQueue(lngIndex)
    End If
End Property

Public Property Let Last(varTmp As Variant)
    Dim lngIndex        As Long
    If Me.IsEmpty Then
        Exit Property
    End If
    lngIndex = (mlngRearIndex - 1 + mlngQueueSize) Mod mlngQueueSize
    If IsObject(varTmp) Then
        Set marrQueue(lngIndex) = varTmp
    Else
        marrQueue(lngIndex) = varTmp
    End If
End Property

Public Property Set Last(objTmp As Object)
    If Me.IsEmpty Then
        Exit Property
    End If
    Set marrQueue((mlngRearIndex - 1 + mlngQueueSize) Mod mlngQueueSize) = objTmp
End Property
'Last(R/W):获取队列指定位置的元素或替换指定位置的元素，lngIndex 排第几，从头到尾，>0
Public Property Get Rank(ByVal lngIndex As Long) As Variant
    If Me.IsEmpty Then
        Rank = Empty
        Exit Property
    End If
    If lngIndex > mlngQueueCount Then
        Rank = Empty
        Exit Property
    End If
    lngIndex = (mlngFrontIndex + lngIndex - 1) Mod mlngQueueSize
    If IsObject(marrQueue(lngIndex)) Then
        Set Rank = marrQueue(lngIndex)
    Else
        Rank = marrQueue(lngIndex)
    End If
End Property

Public Property Let Rank(ByVal lngIndex As Long, varTmp As Variant)
    If Me.IsEmpty Then
        Exit Property
    End If
    If lngIndex > mlngQueueCount Then
        Exit Property
    End If
    lngIndex = (mlngFrontIndex + lngIndex - 1) Mod mlngQueueSize
    If IsObject(varTmp) Then
        Set marrQueue(lngIndex) = varTmp
    Else
        marrQueue(lngIndex) = varTmp
    End If
End Property


Public Property Set Rank(ByVal lngIndex As Long, objTmp As Object)
    If Me.IsEmpty Then
        Exit Property
    End If
    If lngIndex > mlngQueueCount Then
        Exit Property
    End If
    Set marrQueue((mlngFrontIndex + lngIndex - 1) Mod mlngQueueSize) = objTmp
End Property
'IsFull(R):判断队列是否是满的
Public Property Get IsFull() As Boolean
    IsFull = mlngQueueSize = mlngQueueCount And mlngQueueSize <> 0
End Property
'IsEmpty(R):判断队列是否是空的
Public Property Get IsEmpty() As Boolean
    IsEmpty = mlngQueueCount = 0
End Property
'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------

'@方法    InitQueue
'   初始化队列
'@返回值
'
'@参数:
'lngQueueSize Long In
'   队列大小
'@备注
'
Public Sub InitQueue(ByVal lngQueueSize As Long)
    mlngQueueSize = lngQueueSize
    ReDim Preserve marrQueue(mlngQueueSize - 1)
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
End Sub
'@方法    EnQueue
'   元素进入队列
'@返回值  Boolean
'
'@参数:
'varQueueItem Variant In
'   进入队列的元素
'@备注
'
Public Function EnQueue(ByVal varQueueItem As Variant) As Boolean
    If Me.IsEmpty Then
        mlngFrontIndex = 0
        mlngRearIndex = 0
    ElseIf Me.IsFull Then
        Call QueueAutoIncrement
    End If
    mlngQueueCount = mlngQueueCount + 1
    If IsObject(varQueueItem) Then
        Set marrQueue(mlngRearIndex) = varQueueItem
    Else
        marrQueue(mlngRearIndex) = varQueueItem
    End If
    mlngRearIndex = (mlngRearIndex + 1) Mod mlngQueueSize
    EnQueue = True
End Function
'@方法    DeQueue
'   元素出队列
'@返回值  Variant
'
'@参数:
'@备注
'
Public Function DeQueue() As Variant
    If Me.IsEmpty Then
        DeQueue = Empty
        Exit Function
    Else
        If IsObject(marrQueue(mlngFrontIndex)) Then
            Set DeQueue = marrQueue(mlngFrontIndex)
        Else
            DeQueue = marrQueue(mlngFrontIndex)
        End If
        mlngFrontIndex = (mlngFrontIndex + 1) Mod mlngQueueSize
        mlngQueueCount = mlngQueueCount - 1
    End If
End Function

'@方法    Clear
'   清理队列(仅移动指针，不销毁数据）
'@返回值
'
'@参数:
'@备注
'
Public Sub Clear()
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
End Sub
'@方法    TearDown
'   销毁队列（清理队列的数据，彻底销毁）
'@返回值
'
'@参数:
'@备注
'
Public Sub TearDown()
    Erase marrQueue
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
    mlngQueueSize = 0
End Sub
'@方法    ToString
'   展示队列
'@返回值  String
'
'@参数:
'@备注
'
Public Function ToString() As String
    ToString = DisPlayOneValue(Me.Data)
End Function
'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------
'@方法    QueueAutoIncrement
'   自动增长队列
'@返回值  Boolean
'
'@参数:
'@备注
'
Private Function QueueAutoIncrement() As Boolean
    Dim i                   As Long
    Dim lngQueueSizeNew     As Long
    
    lngQueueSizeNew = mlngQueueSize + M_LNG_AUTOINCREMENT
    ReDim Preserve marrQueue(lngQueueSizeNew - 1)
    '在左边
    If mlngFrontIndex + mlngRearIndex < mlngQueueSize Then
        For i = 0 To mlngFrontIndex - 1
            marrQueue((mlngQueueSize + i) Mod lngQueueSizeNew) = marrQueue(i)
        Next
        mlngRearIndex = (mlngFrontIndex + mlngQueueSize) Mod lngQueueSizeNew
    Else
        For i = mlngFrontIndex To mlngQueueSize - 1
            marrQueue((i + M_LNG_AUTOINCREMENT) Mod lngQueueSizeNew) = marrQueue(i)
        Next
        mlngFrontIndex = (mlngFrontIndex + M_LNG_AUTOINCREMENT) Mod lngQueueSizeNew
    End If
    mlngQueueSize = lngQueueSizeNew
    QueueAutoIncrement = True
End Function
'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------

Private Sub Class_Initialize()
    mlngQueueSize = M_LNG_INIT_QUEUESIZE
    ReDim Preserve marrQueue(mlngQueueSize - 1)
End Sub

Private Sub Class_Terminate()
    Erase marrQueue
End Sub


