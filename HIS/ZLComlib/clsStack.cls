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
'���ܣ������Ƕ�ջ�Ļ�����
'���ߣ�lshuo,2017-8-2
'========================================================
'========================================================
'==��������
'========================================================
Private marrStack()                 As Variant  '��ǰ��ջ

Private mlngFrontIndex              As Long     '��ջ��β��
Private mlngBaseIndex               As Long     '��ջ��ͷ��

Private mlngStackSize               As Long     '��ջ���ߴ�

Private Const M_LNG_AUTOINCREMENT   As Long = 20         '�������Ļ�����
'========================================================
'==����
'========================================================
'StackSize:(R/W)��ջ��С
Public Property Get StackSize() As Long
    StackSize = mlngStackSize
End Property

Public Property Let StackSize(ByVal lngStackSize As Long)
    If lngStackSize > 0 Then
        Call InitStack(lngStackSize)
    End If
End Property
'Count:(R)��ջԪ�ظ���
Public Property Get Count() As Long
    If mlngFrontIndex >= mlngBaseIndex Then
        Count = mlngFrontIndex - mlngBaseIndex + 1
    End If
End Property

'--------------------------------------------------------------------------------------------------
'�ӿ�           GetStack
'����           ��ö�ջ����
'����ֵ         Variant
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
'�ӿ�           InitStack
'����
'����ֵ
'����б�:
'������         ����                    ˵��
'lngStackSize   Long                    ��ջ��С
'-------------------------------------------------------------------------------------------------
Public Sub InitStack(ByVal lngStackSize As Long)
    mlngStackSize = lngStackSize
    ReDim marrStack(mlngStackSize - 1)
    mlngBaseIndex = 0
    mlngFrontIndex = -1
End Sub

'--------------------------------------------------------------------------------------------------
'�ӿ�           Push
'����           Ԫ�ؽ����ջ
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'varStackItem   Variant                 �����ջ��Ԫ��
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
'�ӿ�           Pop
'����           ������ջ����Ԫ��
'lngTimes       �������Σ����ʱ����ʱ�������һ�εġ�
'����ֵ         Variant
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
'�ӿ�           Rank
'����           ��ȡ��ջ�ƶ�λ�õ�һ��Ԫ�أ���ͷ��β��,��0����
'����ֵ         Variant
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
'�ӿ�           Current
'����           ��ȡ��ջ��ǰԪ��
'����ֵ         Variant
'-------------------------------------------------------------------------------------------------
Public Function Current() As Variant
    If Me.IsEmpty Then
        Current = Empty
        Exit Function
    End If
    Call gobjComLib.zlCommFun.AssignVariant(Current, marrStack(mlngFrontIndex))
End Function

'--------------------------------------------------------------------------------------------------
'�ӿ�           IsFull
'����           �ж϶�ջ�Ƿ�������
'����ֵ         Boolean
'-------------------------------------------------------------------------------------------------
Public Function IsFull() As Boolean
    IsFull = (mlngFrontIndex + 1 = mlngStackSize)
End Function

'--------------------------------------------------------------------------------------------------
'�ӿ�           IsEmpty
'����           �ж϶�ջ�Ƿ��ǿյ�
'����ֵ         Boolean
'-------------------------------------------------------------------------------------------------
Public Function IsEmpty() As Boolean
    IsEmpty = mlngFrontIndex < mlngBaseIndex And Not Me.IsFull
End Function

'--------------------------------------------------------------------------------------------------
'�ӿ�           Clear
'����           ������ջ(���ƶ�ָ�룬���������ݣ�
'-------------------------------------------------------------------------------------------------
Public Sub Clear()
    mlngFrontIndex = -1
    mlngBaseIndex = 0
End Sub
'--------------------------------------------------------------------------------------------------
'�ӿ�           TearDown
'����           ���ٶ�ջ��������ջ�����ݣ��������٣�
'-------------------------------------------------------------------------------------------------
Public Sub TearDown()
    Erase marrStack
    mlngFrontIndex = -1
    mlngBaseIndex = 0
End Sub

'--------------------------------------------------------------------------------------------------
'�ӿ�           ToString
'����           չʾ��ջ
'����ֵ         String
'-------------------------------------------------------------------------------------------------
Public Function ToString() As String
    ToString = gobjComLib.zlStr.DisPlayOneValue(Me.Data)
End Function

'--------------------------------------------------------------------------------------------------
'�ӿ�           StackAutoIncrement
'����           �Զ����Ŷ���
'����ֵ         Boolean
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
