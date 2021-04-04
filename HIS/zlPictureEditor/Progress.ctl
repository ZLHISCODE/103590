VERSION 5.00
Begin VB.UserControl Progress 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   345
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Progress.ctx":0000
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   170
      Left            =   1800
      TabIndex        =   0
      Top             =   90
      Width           =   1005
   End
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'######################################################################################
'##ģ �� ����Progress.ctl
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��5��20��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ����һ���Զ���ķ����Ľ������ؼ���
'##��    ����
'######################################################################################

Option Explicit
Private mvarValue As Single
Private mvarMax As Long

Private Enum GradientFillRectType
   GRADIENT_FILL_RECT_H = 0
   GRADIENT_FILL_RECT_V = 1
End Enum

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type TRIVERTEX
   X As Long
   Y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type
Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GradientFill Lib "msimg32" ( _
   ByVal hdc As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_RECT, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" ( _
   ByVal hdc As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_TRIANGLE, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
'����ָ����ɫ���߼���ˢ
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'ȡ���ڻ�ͼ����Ĳ�ͬ�������ڸ���������������BITMAP, DIBSECTION, EXTLOGPEN, LOGBRUSH, LOGFONT ���� LOGPEN �ṹ
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'��һ������ѡ��ָ�����豸�������������У��ö����Զ��滻��ͬһ���͵�ǰһ����
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'ɾ��һ���߼����ʡ���ˢ�����塢λͼ��������ߵ�ɫ��
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'��ȡ�������ڻ���������Ļ�Ļ����������������ͼ��
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'�ͷű�׼Windows�豸������Դ��
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

'#########################################################################
' ͼ�κ�������

'��ȡ������ʾԪ�صĵ�ǰ��ɫֵ
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'���ƾ��ε�һ�����߶�����
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'��һ�� OLE_COLOR ����ת��Ϊһ�� COLORREF ���͡�
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
'����һ��ͼ�ꡢ��̬������λͼ��
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'ͬ�ϣ������ڶ�����Ϊһ������ֵ��
Private Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'��ʾһ��Windowsλͼ��ʽ��
Private Const CF_BITMAP = 2
'3DЧ����ɫ
Private Const LR_LOADMAP3DCOLORS = &H1000
'ͼƬ���ļ�lpsz�е��룬���Ǵ���Դ�ļ��е��롣
Private Const LR_LOADFROMFILE = &H10
'����͸��ɫ
Private Const LR_LOADTransparent = &H20
'���� �豸�޹� DIB λͼ�������豸���λͼ��
Private Const IMAGE_BITMAP = 0
'ʹ��ָ����ˢ����������
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'��Դ������Ŀ�껭���ı��ؿ鴫�����ɫ����
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'�������洰�壨��Ļ���ľ��
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'��ȡϵͳ������λ��ϵͳ���ã����гߴ���Ե� Pixel ��ʾ
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Property Get Max() As Long
    Max = mvarMax
End Property

Public Property Let Max(vData As Long)
    mvarMax = vData
    PropertyChanged "Max"
End Property

Public Property Get Value() As Single
    Value = mvarValue
End Property

Public Property Let Value(vData As Single)
    mvarValue = vData
    lblProgress.Caption = Format(vData, "0%")
    DrawProgress mvarValue, UserControl.hdc, 0, 0, ScaleWidth / Screen.TwipsPerPixelX, ScaleHeight / Screen.TwipsPerPixelY
    Refresh
    PropertyChanged "Value"
End Property

Public Sub Cls()
    UserControl.Cls
End Sub

Private Sub UserControl_Initialize()
    mvarValue = 0#
End Sub

Private Sub UserControl_Resize()
    lblProgress.Move 0, (ScaleHeight - lblProgress.Height) / 2, ScaleWidth
End Sub

'######################################################################################
'   ���Ʋ�ɫ������
'######################################################################################

Private Sub DrawProgress( _
      lPercent As Single, _
      ByVal lHDC As Long, _
      ByVal lLeft As Long, ByVal lTop As Long, _
      ByVal lRight As Long, ByVal lBottom As Long _
   )
Dim hBr As Long
Dim tR As RECT
Dim tProgR As RECT

   tR.Left = lLeft + 1
   tR.Top = lTop + 1
   tR.Right = lRight - 1
   tR.Bottom = lBottom - 1

   ' Draw the progress bar
   LSet tProgR = tR
   tProgR.Right = tProgR.Left + (tProgR.Right - tProgR.Left) * lPercent
   GradientFillRect lHDC, tProgR, RGB(234, 94, 45), RGB(238, 164, 36), GRADIENT_FILL_RECT_H
   
   ' Draw the text in front of the progress bar
'   DrawTextA lHDC, Format(lPercent, "0%"), -1, tR, DT_CENTER

   ' Frame the progress bar:
   hBr = CreateSolidBrush(&H0&)
   FrameRect lHDC, tR, hBr
   DeleteObject hBr
End Sub

'��ɫת��
Public Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

Private Sub GradientFillRect( _
      ByVal lHDC As Long, _
      tR As RECT, _
      ByVal oStartColor As OLE_COLOR, _
      ByVal oEndColor As OLE_COLOR, _
      ByVal eDir As GradientFillRectType _
   )
Dim hBrush As Long
Dim lStartColor As Long
Dim lEndColor As Long
Dim lR As Long
   
   ' Use GradientFill:
   lStartColor = TranslateColor(oStartColor)
   lEndColor = TranslateColor(oEndColor)

   Dim tTV(0 To 1) As TRIVERTEX
   Dim tGR As GRADIENT_RECT
   
   setTriVertexColor tTV(0), lStartColor
   tTV(0).X = tR.Left
   tTV(0).Y = tR.Top
   setTriVertexColor tTV(1), lEndColor
   tTV(1).X = tR.Right
   tTV(1).Y = tR.Bottom
   
   tGR.UpperLeft = 0
   tGR.LowerRight = 1
   
   GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
      
   If (Err.Number <> 0) Then
      ' Fill with solid brush:
      hBrush = CreateSolidBrush(TranslateColor(oEndColor))
      FillRect lHDC, tR, hBrush
      DeleteObject hBrush
   End If
End Sub

Private Sub setTriVertexColor(tTV As TRIVERTEX, lColor As Long)
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    lRed = (lColor And &HFF&) * &H100&
    lGreen = (lColor And &HFF00&)
    lBlue = (lColor And &HFF0000) \ &H100&
    setTriVertexColorComponent tTV.Red, lRed
    setTriVertexColorComponent tTV.Green, lGreen
    setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent(ByRef iColor As Integer, ByVal lComponent As Long)
    If (lComponent And &H8000&) = &H8000& Then
       iColor = (lComponent And &H7F00&)
       iColor = iColor Or &H8000
    Else
       iColor = lComponent
    End If
End Sub


