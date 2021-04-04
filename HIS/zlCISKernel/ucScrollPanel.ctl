VERSION 5.00
Begin VB.UserControl ucScrollPanel 
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   2850
   ToolboxBitmap   =   "ucScrollPanel.ctx":0000
   Begin VB.PictureBox picFill 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1005
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   690
      Width           =   500
      Begin VB.Shape shpBack 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   5  'Downward Diagonal
         Height          =   390
         Left            =   60
         Top             =   75
         Width           =   375
      End
   End
   Begin VB.VScrollBar scrV 
      Enabled         =   0   'False
      Height          =   1965
      Left            =   2535
      TabIndex        =   1
      Top             =   -15
      Width           =   295
   End
   Begin VB.HScrollBar scrH 
      Enabled         =   0   'False
      Height          =   295
      Left            =   30
      TabIndex        =   0
      Top             =   1935
      Width           =   2520
   End
   Begin VB.Shape shpEnable 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   5  'Downward Diagonal
      Height          =   585
      Left            =   165
      Top             =   705
      Visible         =   0   'False
      Width           =   630
   End
End
Attribute VB_Name = "ucScrollPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const MLNG_VK_LBUTTON = &H1
Private Const MLNG_HTCAPTION = 2
Private Const MLNG_HTCLIENT = 1
Private Const MLNG_WM_MOVE = 3
Private Const MLNG_WM_PAINT = &HF
Private Const MLNG_WM_ERASEBKGND = &H14
Private Const MLNG_WM_MOUSEWHEEL = &H20A
Private Const MLNG_WM_MOUSEMOVE = &H200
Private Const MLNG_WM_LBUTTONDOWN = &H201
Private Const MLNG_WM_LBUTTONUP = &H202
Private Const MLNG_WM_NCHITTEST = &H84
Private Const MLNG_WM_NCPAINT = &H85

Private Const MLNG_STEP As Long = 320   '滚动条滚动时对应的移动距离

Private Type PointAPI
        x As Long
        Y As Long
End Type

Public Enum TBorderStyle
    bsNone = 0
    bsFixed = 1
End Enum


Public Enum TAppearance
    aNone = 0
    a3D = 1
End Enum

Private mlngMaxWidth As Long
Private mlngMaxHeight As Long

Private mblnScrollState As Boolean
Private mblnIsRegMsg As Boolean
Private mobjPosDic As New Scripting.Dictionary

Private WithEvents mobjMsg As clsMsg
Attribute mobjMsg.VB_VarHelpID = -1



Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As PointAPI) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Event OnResize()
Public Event OnPaint()
Public Event OnKeyDown(KeyCode As Integer, Shift As Integer)


'滚动条滚动状态
Property Get UCScrollState() As Boolean
    UCScrollState = mblnScrollState
End Property

Property Let UCScrollState(value As Boolean)
    mblnScrollState = value
End Property

'背景色
Property Get UCBackColor() As OLE_COLOR
    UCBackColor = UserControl.BackColor
End Property

Property Let UCBackColor(value As OLE_COLOR)
    UserControl.BackColor = value
End Property

'边框样式
Property Get UCBorderStyle() As TBorderStyle
    UCBorderStyle = UserControl.BorderStyle
End Property

Property Let UCBorderStyle(value As TBorderStyle)
    UserControl.BorderStyle = value
End Property

'3D样式
Property Get UCAppearance() As TAppearance
    UCAppearance = UserControl.Appearance
End Property


Property Let UCAppearance(value As TAppearance)
    UserControl.Appearance = value
End Property


'可编辑属性
Property Get UCEnabled() As Boolean
    UCEnabled = UserControl.Enabled
End Property


Property Let UCEnabled(value As Boolean)
    UserControl.Enabled = value
    shpEnable.Visible = Not UserControl.Enabled
    
    If Not UserControl.Enabled Then
        Call ConfigDisableFace
        
        Call AdjustScrollPos
    Else
        Call UserControl_Paint
    End If

'    If Enabled Then
'        Call SetFormToAlpha(255)
'    Else
'        Call SetFormToAlpha(100)
'    End If
End Property

'包含的子控件集合
Property Get UCControls() As ContainedControls
    Set UCControls = UserControl.ContainedControls
End Property


'控件句柄
Property Get UCHwnd() As OLE_HANDLE
    UCHwnd = UserControl.Hwnd
End Property


'Tag属性
Property Get UCTag() As String
    UCTag = UserControl.Tag
End Property

Property Let UCTag(value As String)
    UserControl.Tag = value
End Property


'KeyPreview属性
Property Get UCKeyPreview() As Boolean
    UCKeyPreview = UserControl.KeyPreview
End Property


Property Let UCKeyPreview(value As Boolean)
    UserControl.KeyPreview = value
End Property



'需要继承的方法
Public Function UCScaleX(ByVal sigWidth As Single, Optional ByVal smcFromScale As ScaleModeConstants, _
    Optional ByVal smcToScale As ScaleModeConstants) As Single
    
    UCScaleX = UserControl.ScaleX(sigWidth, smcFromScale, smcToScale)
    
End Function

Public Function UCScaleY(ByVal sigHeight As Single, Optional ByVal smcFromScale As ScaleModeConstants, _
    Optional ByVal smcToScale As ScaleModeConstants) As Single
    
    UCScaleY = UserControl.ScaleY(sigHeight, smcFromScale, smcToScale)
    
End Function






Private Sub mobjMsg_OnWindowMessage(lngResult As Long, ByVal lngHwnd As Long, ByVal lngMsg As Long, ByVal lngWParam As Long, ByVal lngLParam As Long)
On Error GoTo errHandle
    Dim curPos As PointAPI
    Dim lngMoveDirection As Long

    Select Case lngMsg
        Case MLNG_WM_ERASEBKGND
            If Not Ambient.UserMode Then
                '当处于设计时状态时，将变量mblnScrollState设置为false后，以便程序能够从新配置滚动条
                Call GetClientPoint(curPos)
                lngMoveDirection = GetMousePosArea(curPos)
                
                If lngMoveDirection <= 0 Then
                    mblnScrollState = False
                End If
                
                lngResult = 1
                Exit Sub
            End If
            
        Case MLNG_WM_LBUTTONDOWN ', MLNG_WM_MOVE
            If Ambient.UserMode Then
                '如果为运行时，则退出消息处理
                lngResult = mobjMsg.CallDefaultWindowProc(lngHwnd, lngMsg, lngWParam, lngLParam)
                
                Exit Sub
            End If
            
            '以下程序为处理vb设计环境时的消息事件
            Call GetClientPoint(curPos)
            lngMoveDirection = GetMousePosArea(curPos)
            
            If lngMoveDirection <= 0 Then
                lngResult = mobjMsg.CallDefaultWindowProc(lngHwnd, lngMsg, lngWParam, lngLParam)
                
                Exit Sub
            End If

            While (GetAsyncKeyState(MLNG_VK_LBUTTON) < 0)
                If Not mblnIsRegMsg Then
                    lngResult = 1
                    Exit Sub
                End If

                Select Case lngMoveDirection

                    Case 1
                        If scrH.Enabled Then
                            Call ChangeHScrollValue(1)

                            scrH.SetFocus
                            OS.Wait 50
                        End If
                        
                    Case 2
                        If scrH.Enabled Then
                            Call ChangeHScrollValue(-1)
                            
                            scrH.SetFocus
                            OS.Wait 50
                        End If
                        
                    Case 3
                        If scrV.Enabled Then
                            Call ChangeVScrollValue(-1)
                            
                            scrV.SetFocus
                            OS.Wait 50
                        End If
                        
                    Case 4
                        If scrV.Enabled Then
                            Call ChangeVScrollValue(1)

                            scrV.SetFocus
                            OS.Wait 50
                        End If
                        
                    Case 5
                        lngResult = 1
                        scrH.SetFocus

                        Exit Sub

                    Case 6
                        lngResult = 1
                        scrV.SetFocus

                        Exit Sub
                        
                End Select

            Wend

            lngResult = 1
            Exit Sub

        Case MLNG_WM_MOUSEWHEEL
            '处理运行是的滚轮消息
            Call ChangeVScrollValue(lngWParam)

            lngResult = 1
            Exit Sub
    End Select
    
    '其他我们不关心的消息自己不处理，必须由 VB 的默认处理函数处理
    lngResult = mobjMsg.CallDefaultWindowProc(lngHwnd, lngMsg, lngWParam, lngLParam)
Exit Sub
errHandle:
End Sub

Private Function GetMousePosArea(curPos As PointAPI) As Long
'获取指针当前所处位置区域
'  _____________________
'  |                  4|
'  |                   |
'  |                   |
'  |                   6
'  |                   |
'  |                   |
'  |                  3|
'  |1               2  |
'  _________5__________|
'
    If curPos.x >= 0 And curPos.x <= 240 And curPos.Y >= scrH.Top Then
        GetMousePosArea = 1
    ElseIf curPos.x >= scrV.Left - 240 And curPos.x <= scrV.Left And curPos.Y >= scrH.Top Then
        GetMousePosArea = 2
    ElseIf curPos.x >= scrV.Left And curPos.Y >= scrH.Top - 240 And curPos.Y <= scrH.Top Then
        GetMousePosArea = 3
    ElseIf curPos.x >= scrV.Left And curPos.Y >= 0 And curPos.Y <= 240 Then
        GetMousePosArea = 4
    ElseIf curPos.Y >= scrH.Top And curPos.x <= scrH.Width Then
        GetMousePosArea = 5
    ElseIf curPos.x >= scrV.Left And curPos.Y <= scrV.Height Then
        GetMousePosArea = 6
    Else
        GetMousePosArea = 0
    End If
End Function

Private Function GetClientPoint(ByRef curPos As PointAPI) As Long
'获取当前指针对应在控件中的位置
    GetClientPoint = GetCursorPos(curPos)
    
    GetClientPoint = ScreenToClient(Hwnd, curPos)
    
    curPos.x = ScaleX(curPos.x, vbPixels, vbTwips)
    curPos.Y = ScaleY(curPos.Y, vbPixels, vbTwips)
End Function

Private Sub HMove()
'横向移动
On Error GoTo errHandle
    Dim objControl As Object
    Dim lngLeft As Long
    Dim strPosTag As String
    
    
    mblnScrollState = True
    
    For Each objControl In UserControl.ContainedControls
        If OS.ObjectHasProperty(objControl, "visible") Then
            strPosTag = mobjPosDic.Item(objControl.Name)
            
'            lngLeft = Val(Replace(objControl.Tag, Val(objControl.Tag) & "-", ""))
            lngLeft = Val(Replace(strPosTag, Val(strPosTag) & "-", ""))
            objControl.Left = lngLeft - scrH.value * MLNG_STEP
        End If
    Next
    
    DoEvents
    
Exit Sub
errHandle:
    MsgBox err.Description
End Sub

Private Sub picFill_Resize()
On Error Resume Next
    shpBack.Left = 0
    shpBack.Top = 0
    shpBack.Width = picFill.ScaleWidth
    shpBack.Height = picFill.ScaleHeight
End Sub

Private Sub scrH_Change()
    Call HMove
End Sub


Private Sub scrH_Scroll()
    Call HMove
End Sub


Private Sub scrV_Change()
    Call VMove
End Sub


Private Sub scrV_Scroll()
    Call VMove
End Sub


Private Sub ChangeVScrollValue(ByVal lngPosState As Long)
On Error GoTo errHandle
    If Not scrV.Enabled Then Exit Sub
    
    If lngPosState < 0 Then
        If scrV.value = scrV.Max Then Exit Sub
        scrV.value = scrV.value + 1
    Else
        If scrV.value = scrV.Min Then Exit Sub
        scrV.value = scrV.value - 1
    End If
Exit Sub
errHandle:
    MsgBox err.Description
End Sub

Private Sub ChangeHScrollValue(ByVal lngPosState As Long)
On Error GoTo errHandle
    If lngPosState < 0 Then
        If scrH.value = scrH.Max Then Exit Sub
        scrH.value = scrH.value + 1
    Else
        If scrH.value = scrH.Min Then Exit Sub
        scrH.value = scrH.value - 1
    End If
Exit Sub
errHandle:
    MsgBox err.Description
End Sub

Private Sub VMove()
'纵向移动
On Error GoTo errHandle
    Dim objControl As Object
    Dim strPosTag As String
    
    mblnScrollState = True
    
    For Each objControl In UserControl.ContainedControls
        If OS.ObjectHasProperty(objControl, "visible") Then
            strPosTag = mobjPosDic.Item(objControl.Name)
            
'            objControl.Top = Val(objControl.Tag) - scrV.value * MLNG_STEP
            objControl.Top = Val(strPosTag) - scrV.value * MLNG_STEP
        End If
    Next
    
    DoEvents
    
Exit Sub
errHandle:
    MsgBox err.Description
End Sub

Private Sub GetHW(ByRef lngWidth As Long, ByRef lngHeight As Long, ByRef lngLeft As Long, ByRef lngTop As Long)
'获取超出控件显示区域外的大小及X,Y坐标
'
' .........
' .    ___.______________
' .   |   .             |
' .........          .........
'     |              .  |    .
'     |              .  |    .
' .........          .  |    .
' .   |___.__________.__|    .
' .       .          .........
' .........
'
    Dim objControl As Object
    
    Dim lngCurHeight As Long
    Dim lngCurWidth As Long

    
    lngCurHeight = 0
    lngCurWidth = 0
    
    lngHeight = 0
    lngWidth = 0
    lngTop = 0
    lngLeft = 0

    For Each objControl In UserControl.ContainedControls
        If OS.ObjectHasProperty(objControl, "visible") Then
            
'            objControl.Tag = objControl.Top & "-" & objControl.Left
            If mobjPosDic.Exists(objControl.Name) Then
                Call mobjPosDic.Remove(objControl.Name)
            End If
            
            Call mobjPosDic.Add(objControl.Name, objControl.Top & "-" & objControl.Left)
            
            lngCurHeight = objControl.Height + objControl.Top
            If lngCurHeight > lngHeight Then lngHeight = lngCurHeight
            
            lngCurWidth = objControl.Width + objControl.Left
            If lngCurWidth > lngWidth Then lngWidth = lngCurWidth
            
            lngTop = IIF(objControl.Top < lngTop, objControl.Top, lngTop)
            lngLeft = IIF(objControl.Left < lngLeft, objControl.Left, lngLeft)
        End If
    Next
    
End Sub


Private Sub CallPaint()
'触发OnPaint事件
On Error GoTo errHandle
    RaiseEvent OnPaint
Exit Sub
errHandle:
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CallKeyDown(KeyCode, Shift)
End Sub

Private Sub CallKeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHandle
    RaiseEvent OnKeyDown(KeyCode, Shift)
Exit Sub
errHandle:
End Sub

Private Sub UserControl_Paint()
On Error GoTo errHandle

    If mblnScrollState Then
        Call CallPaint
        Exit Sub
    End If
    
    If Not UserControl.Enabled Then
        Call CallPaint
        '如果enabeld为false，则直接退出滚动条配置,这里不能直接调用UserControl_Resize事件，否则会造成循环刷新
        Exit Sub
    End If
    
    Call CalcScroll
   
    Call SendControlToTop
    
    Call UserControl_Resize
    
    Call CallPaint
Exit Sub
errHandle:
End Sub


Public Sub CalcScroll()
'计算滚动条的值及显示位置
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim blnHasHScroll As Boolean
    Dim blnHasVScroll As Boolean
    
    Call GetHW(lngWidth, lngHeight, lngLeft, lngTop)
    
    '计算是否需要显示滚动条
    blnHasHScroll = False
    blnHasVScroll = False
    
    Select Case True
        Case lngHeight + Abs(lngTop) <= UserControl.Height And lngWidth + Abs(lngLeft) <= UserControl.Width
        Case lngHeight + Abs(lngTop) > UserControl.Height And lngWidth + Abs(lngLeft) > UserControl.Width
            blnHasHScroll = True
            blnHasVScroll = True
        Case lngHeight + Abs(lngTop) > UserControl.Height
            blnHasVScroll = True
            If lngWidth + Abs(lngLeft) > UserControl.Width - scrV.Width Then blnHasVScroll = True
        Case lngWidth + Abs(lngLeft) > UserControl.Width
            blnHasHScroll = True
            If lngHeight + Abs(lngTop) > UserControl.Height - scrH.Height Then blnHasVScroll = True
    End Select
    
    '显示滚动条后计算最大值
    scrH.Max = Fix((lngWidth - UserControl.ScaleWidth + IIF(scrV.Visible, scrV.Width, 0)) / MLNG_STEP) + 1
    scrV.Max = Fix((lngHeight - UserControl.ScaleHeight + IIF(scrH.Visible, scrH.Height, 0)) / MLNG_STEP) + 1
    
    scrH.Min = Fix(lngLeft / MLNG_STEP) - IIF(lngLeft < 0, 1, 0)
    scrV.Min = Fix(lngTop / MLNG_STEP) - IIF(lngTop < 0, 1, 0)
    
    scrH.value = 0
    scrV.value = 0
    
    '判断滚动条是否能够使用
    scrV.Enabled = IIF(lngHeight + Abs(lngTop) <= UserControl.Height - IIF(blnHasHScroll, scrH.Height, 0), False, True)
    scrH.Enabled = IIF(lngWidth + Abs(lngLeft) <= UserControl.Width - IIF(blnHasVScroll, scrV.Width, 0), False, True)
    
    scrV.Visible = scrV.Enabled
    scrH.Visible = scrH.Enabled
End Sub


Private Sub CallResize()
'触发OnResize事件
On Error GoTo errHandle
    RaiseEvent OnResize
Exit Sub
errHandle:
End Sub


Private Sub UserControl_Resize()
On Error Resume Next

    If UserControl.Enabled Then
        mblnScrollState = False
        
        Call CalcScroll
        
        mblnScrollState = True
    End If
    
    Call AdjustScrollPos
    
    Call CallResize
End Sub

Private Sub AdjustScrollPos()
On Error GoTo errHandle
    scrH.Left = 0
    scrH.Top = ScaleHeight - scrH.Height
    scrH.Width = Width - IIF(scrV.Visible Or Not UserControl.Enabled, scrV.Width, 0) - IIF(BorderStyle = bsFixed, 50, 0)
    
    scrV.Left = ScaleWidth - scrV.Width
    scrV.Top = 0
    scrV.Height = ScaleHeight - IIF(scrH.Visible Or Not UserControl.Enabled, scrH.Height, 0)
    
    picFill.Left = scrH.Width
    picFill.Top = scrV.Height
    
    picFill.Width = ScaleWidth - scrH.Width
    picFill.Height = ScaleHeight - scrV.Height
    

    shpEnable.Left = 0
    shpEnable.Top = 0
    shpEnable.Width = ScaleWidth
    shpEnable.Height = ScaleHeight
Exit Sub
errHandle:
End Sub

Private Sub SendControlToTop()
    Call picFill.ZOrder(0)

    Call scrH.ZOrder(0)
    Call scrV.ZOrder(0)
End Sub


Private Sub ConfigDisableFace()
'配置不可编辑页面状态
    scrV.Enabled = False
    scrV.Visible = True
    
    scrH.Enabled = False
    scrH.Visible = True
End Sub


Private Sub UserControl_Show()
On Error GoTo errHandle

    Call SendControlToTop

'    Call mobjMsg.SetMsgHook(hWnd)
'
'
'    mblnIsRegMsg = True
    
'    Debug.Print "RegMsg"
Exit Sub
errHandle:
'    Debug.Print "RegMsg：" & Err.Description
End Sub


Private Sub UserControl_Hide()
On Error GoTo errHandle
'    mblnIsRegMsg = False
'
'    Call mobjMsg.SetMsgUnHook
    
'    Debug.Print "UnRegMsg"
Exit Sub
errHandle:
'    Debug.Print "UnRegMsg：" & Err.Description
End Sub


Private Sub UserControl_Initialize()
    mblnScrollState = False
    mblnIsRegMsg = False
    
    '注册控件处理消息
    Set mobjMsg = New clsMsg
    Call mobjMsg.SetMsgHook(Hwnd)
    
    mblnIsRegMsg = True
End Sub


Private Sub UserControl_Terminate()
On Error Resume Next
    mblnIsRegMsg = False

    '卸载控件处理消息
    Call mobjMsg.SetMsgUnHook
    
    Set mobjMsg = Nothing
End Sub


Private Sub UserControl_InitProperties()
    UserControl.Appearance = a3D
    UserControl.BorderStyle = bsFixed
    UserControl.BackColor = &H8000000F
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("UCBackColor", UserControl.BackColor, &H8000000F)
    
    Call PropBag.WriteProperty("UCBorderStyle", UserControl.BorderStyle, bsFixed)
    Call PropBag.WriteProperty("UCAppearance", UserControl.Appearance, a3D)
    Call PropBag.WriteProperty("UCEnabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("UCTag", UserControl.Tag, "")
    Call PropBag.WriteProperty("UCKeyPreview", UserControl.KeyPreview, True)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

    UserControl.BorderStyle = PropBag.ReadProperty("UCBorderStyle", bsFixed)
    UserControl.Appearance = PropBag.ReadProperty("UCAppearance", a3D)
    UserControl.Enabled = PropBag.ReadProperty("UCEnabled", True)
    UserControl.Tag = PropBag.ReadProperty("UCTag", "")
    UserControl.KeyPreview = PropBag.ReadProperty("UCKeyPreview", True)
    
    shpEnable.Visible = Not UserControl.Enabled
    
    If Not UserControl.Enabled Then
        Call ConfigDisableFace

        Call AdjustScrollPos
    End If
    
    '颜色设置必须放在最后进行读取，因为设置边框或3d样式时，可能造成颜色恢复默认
    UserControl.BackColor = PropBag.ReadProperty("UCBackColor", &H8000000F)
End Sub


