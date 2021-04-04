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
'��д           lshuo
'����           2018/9/29
'ģ��           clsQueue
'˵��           �����Ƕ��еĻ�����
'==================================================================================================
Option Explicit
'========================================================
'���ܣ������Ƕ��еĻ�����
'���ߣ�lshuo,2017-8-2
'========================================================
'========================================================
'==��������
'========================================================
Private marrQueue()         As Variant  '��ǰ����

Private mlngFrontIndex      As Long     '���еĿ�ͷ
Private mlngRearIndex       As Long     '���е�β��

Private mlngQueueSize       As Long     '�������ߴ�
Private mlngQueueCount      As Long     '��ǰ���е����ݸ���
Private mlngIncrementTimes  As Long
Private Const M_LNG_AUTOINCREMENT   As Long = 20         '�������Ļ�����

'========================================================
'==����
'========================================================
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

'--------------------------------------------------------------------------------------------------
'�ӿ�           Current
'����           ��ȡ���е�ǰԪ��
'����ֵ         Variant
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
'�ӿ�           Current
'����           �������еĶ���Ԫ��
'����ֵ         Variant
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
'�ӿ�           Current
'����           �������еĶ���Ԫ��
'����ֵ         Variant
'-------------------------------------------------------------------------------------------------
Public Property Set Current(objTmp As Object)
    If Me.IsEmpty Then
        Exit Property
    End If
    Set marrQueue(mlngFrontIndex) = objTmp
End Property
'--------------------------------------------------------------------------------------------------
'�ӿ�           Last
'����           ��ȡ�������һ��Ԫ��
'����ֵ         Variant
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
'�ӿ�           Last
'����           �滻����Ԫ��
'����ֵ         Variant
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
'�ӿ�           Last
'����           �滻����Ԫ��
'����ֵ         Variant
'-------------------------------------------------------------------------------------------------
Public Property Set Last(objTmp As Object)
    If Me.IsEmpty Then
        Exit Property
    End If
    Set marrQueue((mlngRearIndex - 1 + mlngQueueSize) Mod mlngQueueSize) = objTmp
End Property
'--------------------------------------------------------------------------------------------------
'�ӿ�           Rank
'����           ��ȡ�����ƶ�λ�õ�һ��Ԫ��
'����ֵ         Variant
'lngIndex       Long                �ŵڼ�����ͷ��β��>0
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
'�ӿ�           Rank
'����           ��ȡ�����ƶ�λ�õ�һ��Ԫ��
'����ֵ         Variant
'lngIndex       Long                �ŵڼ�����ͷ��β��>0
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
'�ӿ�           Rank
'����           ��ȡ�����ƶ�λ�õ�һ��Ԫ��
'����ֵ         Variant
'lngIndex       Long                �ŵڼ�����ͷ��β��>0
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
'�ӿ�           IsFull
'����           �ж϶����Ƿ�������
'����ֵ         Boolean
'-------------------------------------------------------------------------------------------------
Public Property Get IsFull() As Boolean
    IsFull = mlngQueueSize = mlngQueueCount And mlngQueueSize <> 0
End Property
'--------------------------------------------------------------------------------------------------
'�ӿ�           IsEmpty
'����           �ж϶����Ƿ��ǿյ�
'����ֵ         Boolean
'-------------------------------------------------------------------------------------------------
Public Property Get IsEmpty() As Boolean
    IsEmpty = mlngQueueCount = 0
End Property
'--------------------------------------------------------------------------------------------------
'�ӿ�           InitQueue
'����           ��ʼ������
'����ֵ
'����б�:
'������         ����                    ˵��
'lngQueueSize   Long                    ���д�С
'-------------------------------------------------------------------------------------------------
Public Sub InitQueue(ByVal lngQueueSize As Long)
    mlngQueueSize = lngQueueSize
    ReDim Preserve marrQueue(mlngQueueSize - 1)
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
End Sub

'--------------------------------------------------------------------------------------------------
'�ӿ�           EnQueue
'����           Ԫ�ؽ������
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'varQueueItem   Variant                 ������е�Ԫ��
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
'�ӿ�           DeQueue
'����           Ԫ�س�����
'����ֵ         Variant
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
'�ӿ�           Clear
'����           ��������(���ƶ�ָ�룬���������ݣ�
'-------------------------------------------------------------------------------------------------
Public Sub Clear()
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
End Sub
'--------------------------------------------------------------------------------------------------
'�ӿ�           TearDown
'����           ���ٶ��У��������е����ݣ��������٣�
'-------------------------------------------------------------------------------------------------
Public Sub TearDown()
    Erase marrQueue
    mlngFrontIndex = 0
    mlngRearIndex = 0
    mlngQueueCount = 0
    mlngQueueSize = 0
End Sub
'--------------------------------------------------------------------------------------------------
'�ӿ�           ToString
'����           չʾ����
'����ֵ         String
'-------------------------------------------------------------------------------------------------
Public Function ToString() As String
    ToString = gclsBase.DisPlayOneValue(Me.Data)
End Function

'--------------------------------------------------------------------------------------------------
'�ӿ�           QueueAutoIncrement
'����           �Զ���������
'����ֵ         Boolean
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

Private Sub Class_Terminate()
    Erase marrQueue
End Sub