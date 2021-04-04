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
  
Private Type TBlendProps
        BlendOp   As Byte
        BlendFlags   As Byte
        SourceConstantAlpha   As Byte
        AlphaFormat   As Byte
End Type
  
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
    
    '��ÿؼ���ǰλ�úʹ�С
    ClientToScreen UserControl.hWnd, tPt
    
    Call GetClientRect(UserControl.hWnd, m_rcControl)
    
    '����һ���ڴ�λͼ
    If m_hMemBmp <> 0 Then
        DeleteObject (SelectObject(m_hMemDC, m_hMemBmpPrev))
    End If
    
    m_hMemBmp = CreateCompatibleBitmap(UserControl.hdc, m_rcControl.Right, m_rcControl.Bottom)
    m_hMemBmpPrev = SelectObject(m_hMemDC, m_hMemBmp)
      
    '���ؿؼ�
    ShowWindow UserControl.hWnd, SW_HIDE
    DoEvents
      
    '����ؼ�������ͼ���ڴ�λͼ��
    Dim hDesktopDC As Long
    hDesktopDC = GetDC(UserControl.hWnd)
    BitBlt m_hMemDC, 0, 0, m_rcControl.Right, m_rcControl.Bottom, hDesktopDC, 0, 0, vbSrcCopy
    ReleaseDC 0, hDesktopDC
      
    'ͨ��alphaЧ�����а�͸����Ⱦ
    UserControl.AutoRedraw = True
    
    bp.BlendOp = AC_SRC_OVER
    bp.BlendFlags = 0
    bp.SourceConstantAlpha = iAlphaValue
    bp.AlphaFormat = 0
    
    CopyMemory bpPtr, bp, 4
    AlphaBlend m_hMemDC, 0, 0, m_rcControl.Right, m_rcControl.Bottom, UserControl.hdc, 0, 0, m_rcControl.Right, m_rcControl.Bottom, bpPtr
    UserControl.AutoRedraw = False
  
    '��ʾ�ؼ�
    ShowWindow UserControl.hWnd, SW_SHOW
      
    '����Ⱦ��Ľ�����Ƶ��ؼ���
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
