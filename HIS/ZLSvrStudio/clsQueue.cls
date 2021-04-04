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
'==================================================================================================
'编写           lshuo
'日期           2018/9/29
'模块           clsQueue
'说明           该类是队列的基本类
'==================================================================================================
Option Explicit
'========================================================
'功能：该类是队列的基本类
'编者：lshuo,2017-8-2
'========================================================
'========================================================
'==变量缓存
'========================================================
Private marrQueue()         As Variant  '当前队列

Private mlngFrontIndex      As Long     '队列的开头
Private mlngRearIndex       As Long     '队列的尾部

Private mlngQueueSize       As Long     '队列最大尺寸
Private mlngQueueCount      As Long     '当前队列的数据个数
Private mlngIncrementTimes  As Long
Private Const M_LNG_AUTOINCREMENT   As Long = 20         '自增长的基础量

'========================================================
'==属性
'========================================================
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

'--------------------------------------------------------------------------------------------------
'接口           Current
'功能           获取队列当前元素
'返回值         Variant
'-------------------------------------------------------------------------------------------------
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

'--------------------------------------------------------------------------------------------------
'接口           Current
'功能           更换队列的队首元素
'返回值         Variant
'-------------------------------------------------------------------------------------------------
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
'--------------------------------------------------------------------------------------------------
'接口           Current
'功能           更换队列的队首元素
'返回值         Variant
'-------------------------------------------------------------------------------------------------
Public Property Set Current(objTmp As Object)
    If Me.IsEmpty Then
        Exit Property
    End If
    Set marrQueue(mlngFrontIndex) = objTmp
End Property
'--------------------------------------------------------------------------------------------------
'接口           Last
'功能           获取队列最后一个元素
'返回值         Variant
'-------------------------------------------------------------------------------------------------
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
'--------------------------------------------------------------------------------------------------
'接口           Last
'功能           替换最后的元素
'返回值         Variant
'-------------------------------------------------------------------------------------------------
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
'--------------------------------------------------------------------------------------------------
'接口           Last
'功能           替换最后的元素
'返回值         Variant
'-------------------------------------------------------------------------------------------------
Public Property Set Last(objTmp As Object)
    If Me.IsEmpty Then
        Exit Property
    End If
    Set marrQueue((mlngRearIndex - 1 + mlngQueueSize) Mod mlngQueueSize) = objTmp
End Property
'--------------------------------------------------------------------------------------------------
'接口           Rank
'功能           获取队列制定位置的一个元素
'返回值         Variant
'lngIndex       Long                排第几，从头到尾，>0
'-------------------------------------------------------------------------------------------------
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
'--------------------------------------------------------------------------------------------------
'接口           Rank
'功能           获取队列制定位置的一个元素
'返回值         Variant
'lngIndex       Long                排第几，从头到尾，>0
'-------------------------------------------------------------------------------------------------
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

'--------------------------------------------------------------------------------------------------
'接口           Rank
'功能           获取队列制定位置的一个元素
'返回值         Variant
'lngIndex       Long                排第几，从头到尾，>0
'-------------------------------------------------------------------------------------------------
Public Property Set Rank(ByVal lngIndex As Long, objTmp As Object)
    If Me.IsEmpty Then
        Exit Property
    End If
    If lngIndex > mlngQueueCount Then
        Exit Property
    End If
    Set marrQueue((mlngFrontIndex + lngIndex - 1) Mod mlngQueueSize) = objTmp
End Property
'--------------------------------------------------------------------------------------------------
'接口           IsFull
'功能           判断队列是否是满的
'返回值         Boolean
'-------------------------------------------------------------------------------------------------
Public Property Get IsFull() As Boolean
    IsFull = mlngQueueSize = mlngQueueCount And mlngQueueSize <> 0
End Property
'--------------------------------------------------------------------------------------------------
'接口           IsEmpty
'功能           判断队列是否是空的
'返回值         Boolean
'-------------------------------------------------------------------------------------------------
Public Property Get IsEmpty() As Boolean
    IsEmpty = mlngQueueCount = 0
End Property
'--------------------------------------------------------------------------------------------------
'接口           InitQueue
'功能           初始化队列
'返回值
'入参列表:
'参数名         类型                    说明
'lngQueueSize   Long                    队列大小
'-------------------------------------------------------------------------------------------------
Public Sub InitQueue(ByVal lngQueueSize As Long)
    mlngQueueSize = lngQueueSize
    ReDim Preserve marrQueue(mlngQueueSize - 1)
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
End Sub

'--------------------------------------------------------------------------------------------------
'接口           EnQueue
'功能           元素进入队列
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'varQueueItem   Variant                 进入队列的元素
'-------------------------------------------------------------------------------------------------
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

'--------------------------------------------------------------------------------------------------
'接口           DeQueue
'功能           元素出队列
'返回值         Variant
'-------------------------------------------------------------------------------------------------
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
'--------------------------------------------------------------------------------------------------
'接口           Clear
'功能           清理队列(仅移动指针，不销毁数据）
'-------------------------------------------------------------------------------------------------
Public Sub Clear()
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
End Sub
'--------------------------------------------------------------------------------------------------
'接口           TearDown
'功能           销毁队列（清理队列的数据，彻底销毁）
'-------------------------------------------------------------------------------------------------
Public Sub TearDown()
    Erase marrQueue
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
    mlngQueueSize = 0
End Sub
'--------------------------------------------------------------------------------------------------
'接口           ToString
'功能           展示队列
'返回值         String
'-------------------------------------------------------------------------------------------------
Public Function ToString() As String
    ToString = gclsBase.DisPlayOneValue(Me.Data)
End Function

'--------------------------------------------------------------------------------------------------
'接口           QueueAutoIncrement
'功能           自动增长队列
'返回值         Boolean
'-------------------------------------------------------------------------------------------------
Private Function QueueAutoIncrement() As Boolean
    Dim i                   As Long
    Dim lngQueueSizeNew     As Long
    
    If mlngIncrementTimes < 50 Then
        lngQueueSizeNew = mlngQueueSize + (mlngIncrementTimes + 1) * M_LNG_AUTOINCREMENT
    Else
        mlngIncrementTimes = mlngIncrementTimes + 1
        lngQueueSizeNew = mlngQueueSize + 1000
    End If
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

Private Sub Class_Terminate()
    Erase marrQueue
End Sub
