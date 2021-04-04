VERSION 5.00
Begin VB.UserControl TranControl 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "TranControl.ctx":0000
End
Attribute VB_Name = "TranControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Const SW_SHOW = 5
Private Const SW_HIDE = 0
Private Const DT_SINGLELINE = &H20
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const AC_SRC_OVER = &H0


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
  
  
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private Type PointAPI
        X As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Type TBlendProps
        BlendOp   As Byte
        BlendFlags   As Byte
        SourceConstantAlpha   As Byte
        AlphaFormat   As Byte
End Type
  
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
  
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
  
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
  
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
  
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
  
Private m_hMemDC As Long
Private m_hMemBmp As Long, m_hMemBmpPrev As Long
Private m_rcControl As RECT

Private iAlphaValue As Integer


Private mblnGetBmp As Boolean
    
    
Public Event OnResize()
    
    
Property Get TranColor() As OLE_COLOR
    TranColor = UserControl.BackColor
End Property

Property Let TranColor(value As OLE_COLOR)
    UserControl.BackColor = value
    Call Translucence
End Property



Property Get AlphaValue() As Integer
    AlphaValue = iAlphaValue
End Property

Property Let AlphaValue(value As Integer)
    iAlphaValue = value
    Call Translucence
End Property


      

Private Sub UserControl_Initialize()
    mblnGetBmp = False
    
    m_hMemDC = CreateCompatibleDC(UserControl.hdc)
    
    Call Translucence
End Sub
  

Private Sub UserControl_InitProperties()
    iAlphaValue = 128
    BackColor = vbBlack
End Sub


Private Sub UserControl_Resize()
    On Error Resume Next
    
    mblnGetBmp = False
    
    RaiseEvent OnResize
End Sub


Private Sub UserControl_Terminate()
    On Error Resume Next
    
    If m_hMemBmp <> 0 Then
        DeleteObject SelectObject(m_hMemDC, m_hMemBmpPrev)
    End If
    
    DeleteDC m_hMemDC
End Sub
  
Public Sub Translucence()
    On Error Resume Next
    
    Dim bp As TBlendProps
    Dim bpPtr As Long
    Dim hdc As Long
    Dim tPt As PointAPI
    
    '获得控件当前位置和大小
    ClientToScreen UserControl.hWnd, tPt
    
    Call GetClientRect(UserControl.hWnd, m_rcControl)
    
    '创建一幅内存位图
    If m_hMemBmp <> 0 Then
        DeleteObject (SelectObject(m_hMemDC, m_hMemBmpPrev))
    End If
    
    m_hMemBmp = CreateCompatibleBitmap(UserControl.hdc, m_rcControl.Right, m_rcControl.Bottom)
    m_hMemBmpPrev = SelectObject(m_hMemDC, m_hMemBmp)
      
    '隐藏控件
    ShowWindow UserControl.hWnd, SW_HIDE
    DoEvents
      
    '保存控件容器的图像到内存位图中
    Dim hDesktopDC As Long
    hDesktopDC = GetDC(UserControl.hWnd)
    BitBlt m_hMemDC, 0, 0, m_rcControl.Right, m_rcControl.Bottom, hDesktopDC, 0, 0, vbSrcCopy
    ReleaseDC 0, hDesktopDC
      
    '通过alpha效果进行半透明渲染
    UserControl.AutoRedraw = True
    
    bp.BlendOp = AC_SRC_OVER
    bp.BlendFlags = 0
    bp.SourceConstantAlpha = iAlphaValue
    bp.AlphaFormat = 0
    
    CopyMemory bpPtr, bp, 4
    AlphaBlend m_hMemDC, 0, 0, m_rcControl.Right, m_rcControl.Bottom, UserControl.hdc, 0, 0, m_rcControl.Right, m_rcControl.Bottom, bpPtr
    UserControl.AutoRedraw = False
  
    '显示控件
    ShowWindow UserControl.hWnd, SW_SHOW
      
    '将渲染后的结果复制到控件中
    BitBlt UserControl.hdc, 0, 0, m_rcControl.Right, m_rcControl.Bottom, m_hMemDC, 0, 0, vbSrcCopy
    
    mblnGetBmp = True
End Sub
  
Private Sub UserControl_Paint()
    On Error Resume Next
    
    If mblnGetBmp Then
        BitBlt UserControl.hdc, 0, 0, m_rcControl.Right, m_rcControl.Bottom, m_hMemDC, 0, 0, vbSrcCopy
    Else
        Translucence
    End If
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    UserControl.BackColor = PropBag.ReadProperty("TranColor", vbBlack)
    iAlphaValue = PropBag.ReadProperty("AlphaValue", 128)
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next

    Call PropBag.WriteProperty("TranColor", UserControl.BackColor, vbBlack)
    Call PropBag.WriteProperty("AlphaValue", iAlphaValue, 128)
End Sub
