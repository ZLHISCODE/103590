VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDropDownToolWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'#########################################################################
'##模 块 名：cDropDownToolWindow.cls
'##创 建 人：吴庆伟
'##日    期：2004年12月12日
'##修 改 人：
'##日    期：
'##描    述：下拉式表格选择器类
'#########################################################################

Option Explicit
Private m_lPtr As Long
Private m_hWndParent As Long
Private m_hWndForm As Long
Private m_hWndObject As Long
Private m_bShown As Boolean

'################################################################################################################
'## 用途：  获取指针所指的目标对象
'##
'## 参数：  lPtr    :指针，通过ObjPtr(Obj)来获得的指针。
'##
'## 说明：  注意指针 Ptr 与句柄 hWnd 不同。
'################################################################################################################
Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oT As Object
    ' Turn the pointer into an illegal, uncounted interface
    CopyMemory oT, lPtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set ObjectFromPtr = oT
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory oT, 0&, 4
    ' OK, hit the End button if you must--you'll probably still crash
    ' but it will not be the uncounted reference...
End Property

Public Property Get DropDownObject() As Object
   Set DropDownObject = ObjectFromPtr(m_lPtr)
End Property

Public Sub Create(ByRef picThis As Object)
   If (m_lPtr <> 0) Then
      Destroy
   End If
   On Error Resume Next
   m_lPtr = ObjPtr(picThis)
   If (m_lPtr <> 0) Then
      With DropDownObject
         m_hWndParent = .Container.hWnd
         .BorderStyle = 0
         .Visible = False
      End With
   End If
   If (Err.Number <> 0) Then
      Err.Raise Err.Number, App.EXEName & ".Create", "Invalid object passed to Create"
      m_lPtr = 0
      m_hWndParent = 0
   End If
End Sub

Public Sub Destroy()
   If (m_hWndForm <> 0) Then
'      DetachMessage Me, m_hWndForm, WM_ACTIVATE
      m_hWndForm = 0
   End If
   If (m_lPtr <> 0) Then
      SetParent DropDownObject.hWnd, m_hWndParent
      m_lPtr = 0
   End If
End Sub

Public Sub Show(ByVal x As Long, ByVal y As Long)
Dim tP As POINTAPI
Dim hWndDesktop As Long
Dim lStyle As Long
Dim lHWnd As Long
Dim lParenthWNd As Long
   
   ' Make sure the picture box won't appear in the
   ' task bar by making it into a Tool Window:
   lHWnd = DropDownObject.hWnd
   lStyle = GetWindowLong(lHWnd, GWL_EXSTYLE)
   lStyle = lStyle Or WS_EX_TOOLWINDOW
   lStyle = lStyle And Not (WS_EX_APPWINDOW)
   SetWindowLongL lHWnd, GWL_EXSTYLE, lStyle
   
   ' Determine where to show it in Screen coordinates:
   tP.x = x \ Screen.TwipsPerPixelX: tP.y = y \ Screen.TwipsPerPixelY
   lParenthWNd = DropDownObject.Parent.hWnd
'   ClientToScreen lParenthWNd, tP
   
   ' Make the picture box a child of the desktop (so
   ' it can be fully shown even if it extends beyond
   ' the form boundaries):
   SetParent lHWnd, hWndDesktop
   
   ' Show the form:
   SetWindowPos lHWnd, hWndDesktop, tP.x, tP.y, DropDownObject.Width \ Screen.TwipsPerPixelX, DropDownObject.Height \ Screen.TwipsPerPixelY, SWP_SHOWWINDOW
   
   ' Tell VB it is shown:
   DropDownObject.Visible = True
   DropDownObject.ZOrder
   
   ' Try to set focus:
   SetFocusAPI lHWnd
   
   ' Capture all mouse messages.
   SetCapture lHWnd
   
   ' Start subclassing for Alt-tab
   m_hWndForm = lParenthWNd
   m_hWndObject = lHWnd
'   AttachMessage Me, m_hWndForm, WM_ACTIVATE
   
   ' Store a flag saying we're shown:
   m_bShown = True
End Sub

Public Sub Hide()
   ' Stop subclassing for Alt-tab
   If (m_hWndForm <> 0) Then
'      DetachMessage Me, m_hWndForm, WM_ACTIVATE
   End If
   If (m_hWndObject <> 0) Then
      ' Hide the picturebox:
      DropDownObject.Visible = False
   End If
   m_hWndForm = 0
   m_hWndObject = 0
   
   ' Stop capturing mouse messages:
   ReleaseCapture
   
   ' Store a flag saying we're not shown:
   m_bShown = False
End Sub

Public Property Get IsShown() As Boolean
   ' Return whether we are shown or not.
   IsShown = m_bShown
End Property

Public Property Get InRect(ByVal x As Single, ByVal y As Single) As Boolean
Dim tR As RECT
   If (IsShown()) Then
      GetClientRect DropDownObject.hWnd, tR
      x = x \ Screen.TwipsPerPixelX
      y = y \ Screen.TwipsPerPixelY
      If (PtInRect(tR, x, y) = 1) Then
         InRect = True
      End If
   End If
End Property

Public Sub Resize(ByVal lNewWidth As Long, ByVal lNewHeight As Long)
Dim tWR As RECT, tSR As RECT
Dim lR As Long

   ' Get the size of the window on screen:
   GetWindowRect DropDownObject.hWnd, tWR
   ' Check if it will fit:
   lR = SystemParametersInfo(SPI_GETWORKAREA, 0, tSR, 0)
   If (lR = 0) Then
      ' Call failed - just use standard screen:
      tSR.Left = 0
      tSR.Top = 0
      tSR.Right = Screen.Width \ Screen.TwipsPerPixelX
      tSR.Bottom = Screen.Height \ Screen.TwipsPerPixelY
   End If
   If (tWR.Left + lNewWidth > tSR.Right) Then
      ' too big in x
      lNewWidth = tSR.Right - tWR.Left
   End If
   If (tWR.Top + lNewHeight > tSR.Bottom) Then
      ' too big in y
      lNewHeight = tSR.Bottom - tWR.Top
   End If
   MoveWindow DropDownObject.hWnd, tWR.Left, tWR.Top, lNewWidth, lNewHeight, 1

End Sub

Private Sub Class_Terminate()
   ' Clear up
   Destroy
End Sub

'Private Property Let ISubclass_MsgResponse(ByVal RHS As zlSubTmr.EMsgResponse)
'   ' NR
'End Property
'
'Private Property Get ISubclass_MsgResponse() As zlSubTmr.EMsgResponse
'   ' Respond to the message after windows has done its stuff:
'   ISubclass_MsgResponse = emrPreProcess
'End Property
'
'Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'   Select Case iMsg
'   Case WM_ACTIVATE
'      ' This effectively catches Alt-Tabbing:
'      If (IsShown()) Then
'         Hide
'      End If
'   Case WM_KEYDOWN
''      Debug.Print "KeyDown"
'   End Select
'End Function
'
'
'



