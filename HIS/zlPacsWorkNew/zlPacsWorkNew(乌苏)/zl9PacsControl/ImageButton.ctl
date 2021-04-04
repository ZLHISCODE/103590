VERSION 5.00
Begin VB.UserControl ImageButton 
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   PropertyPages   =   "ImageButton.ctx":0000
   ScaleHeight     =   960
   ScaleWidth      =   960
   ToolboxBitmap   =   "ImageButton.ctx":0026
End
Attribute VB_Name = "ImageButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private nPicture As StdPicture 'normal Picture
Private hPicture As StdPicture 'hot Picture
Private dPicture As StdPicture 'disable Picture
Private wPicture As StdPicture 'down Picture

Private nColor As OLE_COLOR
Private hColor As OLE_COLOR
Private dColor As OLE_COLOR
Private wColor As OLE_COLOR

Private mstrHint As String

Private mintMouseState As Integer 'mouse state 0: Normal, 1:Hot, 2:Down

Private mtoolTip As New clsToolTip



Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event OnKeyDown(KeyCode As Integer, Shift As Integer)
Public Event OnKeyUp(KeyCode As Integer, Shift As Integer)
Public Event OnKeyPress(KeyAscii As Integer)

Public Event OnClick()
Public Event OnDblClick()
Public Event OnResize()

Public Event OnEnterFocus()
Public Event OnExitFocus()

Public Enum AppearanceStyleType
    astFlat = 0
    ast3D = 1
End Enum


Public Enum BackStyleType
    Transparency = 0
    UnTransparency = 1
End Enum


Public Enum BorderStyleType
    None = 0
    FixedSingle = 1
End Enum



'appearance---
Property Get AppearanceStyle() As AppearanceStyleType
    AppearanceStyle = Appearance
End Property


Property Let AppearanceStyle(value As AppearanceStyleType)
    Appearance = value
End Property

'autoredraw---
Property Get IsAutoRedraw() As Boolean
    IsAutoRedraw = AutoRedraw
End Property

Property Let IsAutoRedraw(value As Boolean)
    AutoRedraw = value
End Property

'backstyle---
Property Get BackSty() As BackStyleType
    BackSty = BackStyle
End Property

Property Let BackSty(value As BackStyleType)
    BackStyle = value
End Property

'borderstyle---
Property Get BorderSty() As BorderStyleType
    BorderSty = BorderStyle
End Property

Property Let BorderSty(value As BorderStyleType)
    BorderStyle = value
End Property

'enable---
Property Get IsEnable() As Boolean
    IsEnable = Enabled
End Property

Property Let IsEnable(value As Boolean)
    Enabled = value
    
    If Not Enabled Then
        If dColor >= 0 Then BackColor = dColor
        If Not dPicture Is Nothing Then Set Picture = dPicture
    Else
        If nColor >= 0 Then BackColor = nColor
        If Not nPicture Is Nothing Then Set Picture = nPicture
    End If
End Property


'scaleheight---
Property Get ScaleH() As Long
    ScaleH = ScaleHeight
End Property

Property Let ScaleH(value As Long)
    ScaleHeight = value
End Property

'scalewidth---
Property Get ScaleW() As Long
    ScaleW = ScaleWidth
End Property

Property Let ScaleW(value As Long)
    ScaleWidth = value
End Property

'scaleleft
Property Get ScaleL() As Long
    ScaleL = ScaleLeft
End Property

Property Let ScaleL(value As Long)
    ScaleLeft = value
End Property

'scaletop---
Property Get ScaleT() As Long
    ScaleT = ScaleTop
End Property

Property Let ScaleT(value As Long)
    ScaleTop = value
End Property

'scalemode---
Property Get ScaleType() As ScaleModeConstants
    ScaleType = ScaleMode
End Property

Property Let ScaleType(value As ScaleModeConstants)
    ScaleMode = value
End Property

'mousepointer---
Property Get MousePointerType() As MousePointerConstants
    MousePointerType = MousePointer
End Property

Property Let MousePointerType(value As MousePointerConstants)
    MousePointer = value
End Property

'mouseicon---
Property Get MouseIco() As StdPicture
    Set MouseIco = MouseIcon
End Property

Property Set MouseIco(value As StdPicture)
    Set MouseIcon = value
End Property


'hwnd---
Property Get Handle() As OLE_HANDLE
    Handle = hWnd
End Property


Property Get Hint() As String
    Hint = mstrHint
End Property

Property Let Hint(value As String)
    mstrHint = value
End Property






Property Get NormalColor() As OLE_COLOR
    NormalColor = nColor
End Property

Property Let NormalColor(value As OLE_COLOR)
    nColor = value
    
    If hColor < 0 Then
        hColor = value
    End If
    
    If dColor < 0 Then
        dColor = value
    End If
    
    If wColor < 0 Then
        wColor = value
    End If
    
    BackColor = nColor
End Property



Property Get HotColor() As OLE_COLOR
    HotColor = hColor
End Property

Property Let HotColor(value As OLE_COLOR)
    hColor = value
    
    If nColor < 0 Then
        nColor = value
        
        BackColor = nColor
    End If
    
    If dColor < 0 Then
        dColor = value
    End If
    
    If wColor < 0 Then
        wColor = value
    End If
End Property



Property Get DisableColor() As OLE_COLOR
    DisableColor = dColor
End Property

Property Let DisableColor(value As OLE_COLOR)
    dColor = value
    
    If nColor < 0 Then
        nColor = value
        
        BackColor = nColor
    End If
    
    If hColor < 0 Then
        hColor = value
    End If
    
    If wColor < 0 Then
        wColor = value
    End If
End Property



Property Get DownColor() As OLE_COLOR
    DownColor = wColor
End Property

Property Let DownColor(value As OLE_COLOR)
    wColor = value
    
    If nColor < 0 Then
        nColor = value
        
        BackColor = nColor
    End If
    
    If hColor < 0 Then
        hColor = value
    End If
    
    If dColor < 0 Then
        dColor = value
    End If
End Property


Property Get NormalPicture() As StdPicture
    Set NormalPicture = nPicture
End Property

Property Set NormalPicture(value As StdPicture)
    Set nPicture = value
    
    If hPicture Is Nothing Then
       Set hPicture = value
    End If
    
    If dPicture Is Nothing Then
        Set dPicture = value
    End If
    
    If wPicture Is Nothing Then
        Set wPicture = value
    End If
    
    Set Picture = nPicture
End Property



Property Get HotPicture() As StdPicture
    Set HotPicture = hPicture
End Property

Property Set HotPicture(value As StdPicture)
    Set hPicture = value
    
    If nPicture Is Nothing Then
       Set nPicture = value
       Set Picture = nPicture
    End If
    
    If dPicture Is Nothing Then
        Set dPicture = value
    End If
    
    If wPicture Is Nothing Then
        Set wPicture = value
    End If
End Property



Property Get DisablePicture() As StdPicture
    Set DisablePicture = dPicture
End Property

Property Set DisablePicture(value As StdPicture)
    Set dPicture = value
    
    If hPicture Is Nothing Then
       Set hPicture = value
    End If
    
    If nPicture Is Nothing Then
        Set nPicture = value
        Set Picture = nPicture
    End If
    
    If wPicture Is Nothing Then
        Set wPicture = value
    End If
End Property



Property Get DownPicture() As StdPicture
    Set DownPicture = wPicture
End Property

Property Set DownPicture(value As StdPicture)
    Set wPicture = value
    
    If hPicture Is Nothing Then
       Set hPicture = value
    End If
    
    If dPicture Is Nothing Then
        Set dPicture = value
    End If
    
    If nPicture Is Nothing Then
        Set nPicture = value
        Set Picture = nPicture
    End If
End Property



Private Sub UserControl_Click()
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnClick
End Sub

Private Sub UserControl_DblClick()
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnDblClick
End Sub

Private Sub UserControl_EnterFocus()
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnEnterFocus
End Sub

Private Sub UserControl_ExitFocus()
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnExitFocus
End Sub

Private Sub UserControl_Initialize()
    Set nPicture = Nothing
    Set hPicture = Nothing
    Set dPicture = Nothing
    Set wPicture = Nothing
    
    nColor = -1
    hColor = -1
    dColor = -1
    wColor = -1
    
    mintMouseState = 0
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnKeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnKeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    RaiseEvent OnKeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    mintMouseState = 2
    
    If wColor >= 0 Then
        BackColor = wColor
    End If
    
    If Not wPicture Is Nothing Then
        Set Picture = wPicture
    End If
    
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    Dim blnMouseEnter As Boolean
    
    If Not Enabled Then Exit Sub
        
    blnMouseEnter = (X > 10) And (X < Width - 10) And (Y > 10) And (Y < Height - 10)
    
    If blnMouseEnter Then
        If mintMouseState <> 1 Then
        
            mintMouseState = 1
            
            If hColor >= 0 Then BackColor = hColor
            If Not hPicture Is Nothing Then Set Picture = hPicture
                        
            
            SetCapture hWnd
        
        End If
    Else
        mintMouseState = 0
        
        If nColor >= 0 Then BackColor = nColor
        If Not nPicture Is Nothing Then Set Picture = nPicture
        
        ReleaseCapture
    End If
    
    
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Not Enabled Then Exit Sub
    
    mintMouseState = 0
    
    If nColor >= 0 Then
        BackColor = nColor
    End If
    
    If Not nPicture Is Nothing Then
        Set Picture = nPicture
    End If
    
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    On Error Resume Next
    
    'set control size
    If Not nPicture Is Nothing Then
        Width = ScaleX(nPicture.Width, vbHimetric, vbTwips)
        Height = ScaleY(nPicture.Height, vbHimetric, vbTwips)
    End If
    
    If Not Enabled Then
        If dColor >= 0 Then BackColor = dColor
        If Not dPicture Is Nothing Then Set Picture = dPicture
        
        Exit Sub
    End If
    
    Dim curColor As OLE_COLOR
    
    Select Case mintMouseState
        Case 0: curColor = nColor
        Case 1: curColor = hColor
        Case 2: curColor = wColor
    End Select
    
    If curColor >= 0 Then
        BackColor = curColor
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    Set hPicture = PropBag.ReadProperty("hPicture", Nothing)
    Set nPicture = PropBag.ReadProperty("nPicture", Nothing)
    Set dPicture = PropBag.ReadProperty("dPicture", Nothing)
    Set wPicture = PropBag.ReadProperty("wPicture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    
    hColor = PropBag.ReadProperty("hColor", -1)
    nColor = PropBag.ReadProperty("nColor", -1)
    dColor = PropBag.ReadProperty("dColor", -1)
    wColor = PropBag.ReadProperty("wColor", -1)
    
    Appearance = PropBag.ReadProperty("Appearance", 1)
    AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    BackStyle = PropBag.ReadProperty("BackStyle", 1)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Enabled = PropBag.ReadProperty("Enabled", True)
    
    ScaleHeight = PropBag.ReadProperty("ScaleHeight", 960)
    ScaleWidth = PropBag.ReadProperty("ScaleWidth", 960)
    ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    
    MousePointer = PropBag.ReadProperty("MousePointer", 0)
    mstrHint = PropBag.ReadProperty("Hint", "")
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    RaiseEvent OnResize
End Sub

Private Sub UserControl_Show()
    On Error Resume Next
    
    If Not Enabled Then
        If dColor >= 0 Then BackColor = dColor
        If Not dPicture Is Nothing Then Set Picture = dPicture
        
        Exit Sub
    End If
    
    If Not nPicture Is Nothing Then
        Set Picture = nPicture
        
        Width = ScaleX(nPicture.Width, vbHimetric, vbTwips)
        Height = ScaleY(nPicture.Height, vbHimetric, vbTwips)
        
    End If
    
    If nColor >= 0 Then
        BackColor = nColor
    End If
    
    Call mtoolTip.CreateBalloon(Handle, mstrHint, szClassic, False)
    
End Sub

Private Sub UserControl_Terminate()
    'terminate event...
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    Call PropBag.WriteProperty("hPicture", hPicture, Nothing)
    Call PropBag.WriteProperty("nPicture", nPicture, Nothing)
    Call PropBag.WriteProperty("dPicture", dPicture, Nothing)
    Call PropBag.WriteProperty("wPicture", wPicture, Nothing)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    
    Call PropBag.WriteProperty("hColor", hColor, -1)
    Call PropBag.WriteProperty("nColor", nColor, -1)
    Call PropBag.WriteProperty("dColor", dColor, -1)
    Call PropBag.WriteProperty("wColor", wColor, -1)
    
    Call PropBag.WriteProperty("Appearance", Appearance, 1)
    Call PropBag.WriteProperty("AutoRedraw", AutoRedraw, False)
    Call PropBag.WriteProperty("BackStyle", BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
    
    Call PropBag.WriteProperty("ScaleHeight", ScaleHeight, 960)
    Call PropBag.WriteProperty("ScaleWidth", ScaleWidth, 960)
    Call PropBag.WriteProperty("ScaleLeft", ScaleLeft, 960)
    Call PropBag.WriteProperty("ScaleTop", ScaleTop, 960)
    Call PropBag.WriteProperty("ScaleMode", ScaleMode, 960)
    
    Call PropBag.WriteProperty("MousePointer", MousePointer, 0)
    Call PropBag.WriteProperty("Hwnd", hWnd, 0)
    Call PropBag.WriteProperty("Hint", Hint, "")
End Sub


