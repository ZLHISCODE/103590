VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDxHidDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements DirectXEvent                 'Dx事件接口


Private DX7 As DirectX7
Private DxInput As DirectInput
Private DiDev As DirectInputDevice
Private DiDevEnum As DirectInputEnumDevices
Private EventHandle As Long
Private Js As DIJOYSTATE
Private DiProp_Dead As DIPROPLONG
Private DiProp_Range As DIPROPRANGE
Private DiProp_Saturation As DIPROPLONG


Public Event OnDxKeyPress(ByVal lngButtonNum As Long)


Property Get Handle() As Long
    Handle = EventHandle
End Property


Private Sub DoOnDxKeyPress(ByVal lngButtonNum As Long)
On Error GoTo errHandle
    RaiseEvent OnDxKeyPress(lngButtonNum)
errHandle:
End Sub


Public Function OpenDxDevice(ByVal strInstanceName As String) As Boolean
'打开设备
On Error GoTo errHandle
    Dim i As Integer
    
    OpenDxDevice = False
    
BUGEX "OpenDxDevice 1"
    
    '打开设备前，如果检查到设备没有关闭，则先自动关闭设备
    If EventHandle <> 0 Then Call CloseDxDevice
    
BUGEX "OpenDxDevice 2"
    Set DxInput = DX7.DirectInputCreate()
    Set DiDevEnum = DxInput.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
    
BUGEX "OpenDxDevice 3"
    If DiDevEnum.GetCount = 0 Then Exit Function
    
BUGEX "OpenDxDevice 4"
    For i = 1 To DiDevEnum.GetCount
        If DiDevEnum.GetItem(i).GetInstanceName = strInstanceName Then Exit For
    Next
    
BUGEX "OpenDxDevice 5"
    EventHandle = DX7.CreateEvent(Me)
    
BUGEX "OpenDxDevice 6"
    Set DiDev = DxInput.CreateDevice(DiDevEnum.GetItem(i).GetGuidInstance)
    
BUGEX "OpenDxDevice 7"
    DiDev.SetCommonDataFormat DIFORMAT_JOYSTICK
    DiDev.SetCooperativeLevel 0, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DiDev.SetEventNotification EventHandle
    
BUGEX "OpenDxDevice 8"
    With DiProp_Dead
        .lData = 1000
        .lObj = DIJOFS_X
        .lSize = Len(DiProp_Dead)
        .lHow = DIPH_BYOFFSET
        .lObj = DIJOFS_X
        DiDev.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        DiDev.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        .lObj = DIJOFS_Y
    End With
    
    With DiProp_Saturation
        .lData = 9500
        .lHow = DIPH_BYOFFSET
        .lSize = Len(DiProp_Saturation)
        .lObj = DIJOFS_X
         DiDev.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        .lObj = DIJOFS_Y
         DiDev.SetProperty "DIPROP_SATURATION", DiProp_Saturation
    End With
    
    With DiProp_Range
        .lHow = DIPH_DEVICE
        .lSize = Len(DiProp_Range)
        .lMin = 0
        .lMax = 10000
        DiDev.SetProperty "DIPROP_RANGE", DiProp_Range
    End With

BUGEX "OpenDxDevice 9"
    DiDev.Acquire
    
    OpenDxDevice = True
    
BUGEX "OpenDxDevice 10"
    Exit Function
errHandle:
    OpenDxDevice = False
    
    Debug.Print
End Function


Public Sub CloseDxDevice()
'关闭设备
    If EventHandle <> 0 Then
        DX7.DestroyEvent EventHandle
        EventHandle = 0
    End If
End Sub


Public Sub PollDxDevice()
'轮询设备
    If EventHandle <> 0 Then DiDev.Poll
End Sub



Private Sub Class_Initialize()
BUGEX "clsDxHidDevice_Initialize 1"
    Set DX7 = New DirectX7
BUGEX "clsDxHidDevice_Initialize 2"
End Sub

Private Sub Class_Terminate()
    Call CloseDxDevice
    
    Set DX7 = Nothing
End Sub

Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)
On Error GoTo errHandle
    Dim lngButtonNum As Long
    
    If DiDev Is Nothing Then Exit Sub
    If eventid <> EventHandle Then Exit Sub
            
    DiDev.GetDeviceStateJoystick Js
    If err.Number = DIERR_NOTACQUIRED Or err.Number = DIERR_INPUTLOST Then
        DiDev.Acquire
        Exit Sub
    End If
    
    Select Case 128
        Case Js.buttons(0)
            lngButtonNum = 0
        Case Js.buttons(1)
            lngButtonNum = 1
        Case Js.buttons(2)
            lngButtonNum = 2
        Case Js.buttons(3)
            lngButtonNum = 3
        Case Js.buttons(4)
            lngButtonNum = 4
        Case Js.buttons(5)
            lngButtonNum = 5
        Case Js.buttons(6)
            lngButtonNum = 6
        Case Js.buttons(7)
            lngButtonNum = 7
        Case Js.buttons(8)
            lngButtonNum = 8
        Case Js.buttons(9)
            lngButtonNum = 9
        Case Else
            lngButtonNum = -1
    End Select
    
    If lngButtonNum = -1 Then Exit Sub
    
    Call DoOnDxKeyPress(lngButtonNum)

    Exit Sub
errHandle:
    Debug.Print err.Description
End Sub
