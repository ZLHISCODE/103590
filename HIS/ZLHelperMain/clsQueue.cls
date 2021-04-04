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
'@ģ�� clsQueue-2019/6/26
'@��д lshuo
'@����
'   �����Ƕ��еĻ�����
'@����
'
'@��ע
'
Option Explicit
'---------------------------------------------------------------------------
'                0��API�ͳ�������
'---------------------------------------------------------------------------
Private Const M_LNG_AUTOINCREMENT   As Long = 20            '�������Ļ�����
Private Const M_LNG_INIT_QUEUESIZE  As Long = 5             '���еĳ�ʼ��С
'---------------------------------------------------------------------------
'                1���������
'---------------------------------------------------------------------------
Private marrQueue()         As Variant  '��ǰ����
Private mlngFrontIndex      As Long     '���еĿ�ͷ
Private mlngRearIndex       As Long     '���е�β��

'---------------------------------------------------------------------------
'                2�����Ա����붨��
'---------------------------------------------------------------------------
Private mlngQueueSize       As Long     '�������ߴ�
Private mlngQueueCount      As Long     '��ǰ���е����ݸ���

'QueueSize:(R/W)���д�С
Public Property Get QueueSize() As Long
    QueueSize = mlngQueueSize
End Property

Public Property Let QueueSize(ByVal lngQueueSize As Long)
    If lngQueueSize > 0 Then
        Call InitQueue(lngQueueSize)
    End If
End Property
'Count:(R)����Ԫ�ظ���
Public Property Get Count() As Long
    Count = mlngQueueCount
End Property
'Data(R):��ö�������
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
'Current(R/W):��ȡ���е�ǰԪ�ػ�������еĶ���Ԫ��
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
'Last(R/W):��ȡ�������һ��Ԫ�ػ��滻����Ԫ��
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
'Last(R/W):��ȡ����ָ��λ�õ�Ԫ�ػ��滻ָ��λ�õ�Ԫ�أ�lngIndex �ŵڼ�����ͷ��β��>0
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
'IsFull(R):�ж϶����Ƿ�������
Public Property Get IsFull() As Boolean
    IsFull = mlngQueueSize = mlngQueueCount And mlngQueueSize <> 0
End Property
'IsEmpty(R):�ж϶����Ƿ��ǿյ�
Public Property Get IsEmpty() As Boolean
    IsEmpty = mlngQueueCount = 0
End Property
'---------------------------------------------------------------------------
'                3����������
'---------------------------------------------------------------------------

'@����    InitQueue
'   ��ʼ������
'@����ֵ
'
'@����:
'lngQueueSize Long In
'   ���д�С
'@��ע
'
Public Sub InitQueue(ByVal lngQueueSize As Long)
    mlngQueueSize = lngQueueSize
    ReDim Preserve marrQueue(mlngQueueSize - 1)
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
End Sub
'@����    EnQueue
'   Ԫ�ؽ������
'@����ֵ  Boolean
'
'@����:
'varQueueItem Variant In
'   ������е�Ԫ��
'@��ע
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
'@����    DeQueue
'   Ԫ�س�����
'@����ֵ  Variant
'
'@����:
'@��ע
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

'@����    Clear
'   ��������(���ƶ�ָ�룬���������ݣ�
'@����ֵ
'
'@����:
'@��ע
'
Public Sub Clear()
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
End Sub
'@����    TearDown
'   ���ٶ��У��������е����ݣ��������٣�
'@����ֵ
'
'@����:
'@��ע
'
Public Sub TearDown()
    Erase marrQueue
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
    mlngQueueSize = 0
End Sub
'@����    ToString
'   չʾ����
'@����ֵ  String
'
'@����:
'@��ע
'
Public Function ToString() As String
    ToString = DisPlayOneValue(Me.Data)
End Function
'---------------------------------------------------------------------------
'                4��˽�з���
'---------------------------------------------------------------------------
'@����    QueueAutoIncrement
'   �Զ���������
'@����ֵ  Boolean
'
'@����:
'@��ע
'
Private Function QueueAutoIncrement() As Boolean
    Dim i                   As Long
    Dim lngQueueSizeNew     As Long
    
    lngQueueSizeNew = mlngQueueSize + M_LNG_AUTOINCREMENT
    ReDim Preserve marrQueue(lngQueueSizeNew - 1)
    '�����
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
'                5�����󷽷����¼�
'---------------------------------------------------------------------------

Private Sub Class_Initialize()
    mlngQueueSize = M_LNG_INIT_QUEUESIZE
    ReDim Preserve marrQueue(mlngQueueSize - 1)
End Sub

Private Sub Class_Terminate()
    Erase marrQueue
End Sub

