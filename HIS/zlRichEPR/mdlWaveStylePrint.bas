Attribute VB_Name = "mdlWaveStylePrint"
Option Explicit

'***************************************************************
'�滭���API���������ṹ����
'***************************************************************
'�ṹ����
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type Size
    W   As Long
    H   As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

'��������
Private Const LF_FACESIZE = 32

Private Type LogFont
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private T_OldPoint   As POINTAPI
Private T_NewPoint   As POINTAPI
Private T_ClientRect As RECT
Private T_LableRect  As RECT      '������ı�����Ч����
Private T_ControlRect As RECT     '����˵���Ч����
Private T_Brush      As LOGBRUSH
Private T_Font       As LogFont
Private T_Size       As Size

'������õ����ж���
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateBitmap _
               Lib "gdi32" (ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            ByVal nPlanes As Long, _
                            ByVal nBitCount As Long, _
                            lpBits As Any) As Long

Private Declare Function CreateCompatibleBitmap _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long) As Long
                            
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

'�������ʡ�ˢ��
Private Declare Function CreatePen _
               Lib "gdi32" (ByVal nPenStyle As Long, _
                            ByVal nWidth As Long, _
                            ByVal crColor As Long) As Long

Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long

Private Declare Function ExtCreatePen _
               Lib "gdi32" (ByVal dwPenStyle As Long, _
                            ByVal dwWidth As Long, _
                            lplb As LOGBRUSH, _
                            ByVal dwStyleCount As Long, _
                            lpStyle As Long) As Long

Private Const PS_SOLID = 0
Private Const PS_DASH = 1                    '  -------
Private Const PS_DOT = 2                     '  .......
Private Const PS_DASHDOT = 3                 '  _._._._
Private Const PS_DASHDOTDOT = 4              '  _.._.._
Private Const PS_NULL = 5                    '������ͼ
Private Const PS_COSMETIC = &H0
Private Const PS_GEOMETRIC = &H10000
Private Const PS_ALTERNATE = 8
Private Const PS_ENDCAP_FLAT = &H200
Private Const PS_ENDCAP_MASK = &HF00
Private Const PS_ENDCAP_ROUND = &H0
Private Const PS_ENDCAP_SQUARE = &H100
Private Const PS_JOIN_BEVEL = &H1000
Private Const PS_JOIN_MASK = &HF000
Private Const PS_JOIN_MITER = &H2000
Private Const PS_JOIN_ROUND = &H0

'CreateSolidBrush ������ɫ��ˢ
'CreateBrushIndirect ͨ�� LOGBRUSH ���ʹ�����ˢ
'CreateHatchBrush ������Ӱ��ˢ
'CreatePatternBrush ����ͼ����ˢ
'GetSysColorBrush ����ϵͳ��׼ɫ��ˢ
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

'//lbStyle����ѡֵ:
Private Const BS_SOLID = 0
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_HATCHED = 2
Private Const BS_PATTERN = 3
Private Const BS_INDEXED = 4
Private Const BS_DIBPATTERN = 5
Private Const BS_DIBPATTERNPT = 6
Private Const BS_PATTERN8X8 = 7
Private Const BS_DIBPATTERN8X8 = 8
Private Const BS_MONOPATTERN = 9

'//lbHatch����ѡֵ:
Private Const HS_HORIZONTAL = 0              '  -----
Private Const HS_VERTICAL = 1                '  |||||
Private Const HS_FDIAGONAL = 2               '  \\\\\
Private Const HS_BDIAGONAL = 3               '  /////
Private Const HS_CROSS = 4                   '  +++++
Private Const HS_DIAGCROSS = 5               '  xxxxx

Private Declare Function CreateHatchBrush _
               Lib "gdi32" (ByVal nIndex As Long, _
                            ByVal crColor As Long) As Long
'nIndex,ͬ���溯����lbHatch
'Private Const HS_HORIZONTAL = 0              '  -----
'Private Const HS_VERTICAL = 1                '  |||||
'Private Const HS_FDIAGONAL = 2               '  \\\\\
'Private Const HS_BDIAGONAL = 3               '  /////
'Private Const HS_CROSS = 4                   '  +++++
'Private Const HS_DIAGCROSS = 5               '  xxxxx

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

'BLACK_BRUSH����ɫ����
'DKGRAY_BRUSH������ɫ����
'GRAY_BRUSH����ɫ����
'HOLLOW_BRUSH���ջ��ʣ��൱��HOLLOW_BRUSH��
'LTGRAY_BRUSH������ɫ����
'NULL_BRUSH���ջ��ʣ��൱��HOLLOW_BRUSH��
'WHITE_BRUSH����ɫ����
'BLACK_PEN����ɫ�ֱ�
'WHITE_PEN����ɫ�ֱ�
Private Const WHITE_BRUSH = 0    '��ɫ����
Private Const LTGRAY_BRUSH = 1   '����ɫ����
Private Const GRAY_BRUSH = 2     '��ɫ����
Private Const DKGRAY_BRUSH = 3   '����ɫ����
Private Const BLACK_BRUSH = 4    '��ɫ����
Private Const NULL_BRUSH = 5
Private Const HOLLOW_BRUSH = NULL_BRUSH
Private Const WHITE_PEN = 6      '��ɫ�ֱ�
Private Const BLACK_PEN = 7      '��ɫ�ֱ�

'����һ������
Private Declare Function CreateEllipticRgn _
               Lib "gdi32" (ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long) As Long

Private Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long

Private Declare Function CreateRectRgn _
               Lib "gdi32" (ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long) As Long

Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long

Private Declare Function CreateRoundRectRgn _
               Lib "gdi32" (ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long, _
                            ByVal X3 As Long, _
                            ByVal Y3 As Long) As Long

'�������ͷŶ�����
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SelectObject _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal hObject As Long) As Long

Private Declare Function ReleaseDC _
               Lib "user32" (ByVal hWnd As Long, _
                             ByVal hDC As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'�����ǹ��ܺ���
Private Declare Function DrawFocusRect _
               Lib "user32" (ByVal hDC As Long, _
                             lpRect As RECT) As Long

Private Declare Function Ellipse _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long) As Long

Private Declare Function CreatePolygonRgn _
               Lib "gdi32" (lpPoint As POINTAPI, _
                            ByVal nCount As Long, _
                            ByVal nPolyFillMode As Long) As Long

Private Const ALTERNATE = 1

Private Const WINDING = 2

Private Declare Function ExtFloodFill _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal crColor As Long, _
                            ByVal wFillType As Long) As Long

Private Const FLOODFILLBORDER = 0

Private Const FLOODFILLSURFACE = 1

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function Rectangle _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long) As Long

Private Declare Function FillRect _
               Lib "user32" (ByVal hDC As Long, _
                             lpRect As RECT, _
                             ByVal hBrush As Long) As Long

Private Declare Function FillRgn _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal hRgn As Long, _
                            ByVal hBrush As Long) As Long

Private Declare Function Polyline _
               Lib "gdi32" (ByVal hDC As Long, _
                            lpPoint As POINTAPI, _
                            ByVal nCount As Long) As Long

Private Declare Function FloodFill _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal crColor As Long) As Long

Private Declare Function GetArcDirection Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function GetBrushOrgEx _
               Lib "gdi32" (ByVal hDC As Long, _
                            lpPoint As POINTAPI) As Long

Private Declare Function GetCurrentObject _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal uObjectType As Long) As Long

Private Declare Function GetCurrentPositionEx _
               Lib "gdi32" (ByVal hDC As Long, _
                            lpPoint As POINTAPI) As Long

Private Declare Function GetPixel _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long) As Long

Private Declare Function InvertRect _
               Lib "user32" (ByVal hDC As Long, _
                             lpRect As RECT) As Long

Private Declare Function LineTo _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long) As Long

Private Declare Function GetClientRect _
               Lib "user32" (ByVal hWnd As Long, _
                             lpRect As RECT) As Long

Private Declare Function GetWindowRect _
               Lib "user32" (ByVal hWnd As Long, _
                             lpRect As RECT) As Long

Private Declare Function BitBlt _
               Lib "gdi32" (ByVal hDestDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            ByVal hSrcDC As Long, _
                            ByVal xSrc As Long, _
                            ByVal ySrc As Long, _
                            ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Private Declare Function MoveToEx _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            lpPoint As POINTAPI) As Long

Private Declare Function Polygon _
               Lib "gdi32" (ByVal hDC As Long, _
                            lpPoint As POINTAPI, _
                            ByVal nCount As Long) As Long

Private Declare Function PtInRegion _
               Lib "gdi32" (ByVal hRgn As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long) As Long

Private Declare Function SetTextColor _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal crColor As Long) As Long

Private Declare Function SetBkMode _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal nBkMode As Long) As Long

Private Const TRANSPARENT = 1

Private Declare Function DrawText _
               Lib "user32" _
               Alias "DrawTextA" (ByVal hDC As Long, _
                                  ByVal lpStr As String, _
                                  ByVal nCount As Long, _
                                  lpRect As RECT, _
                                  ByVal wFormat As Long) As Long

Private Const DT_CENTER = &H1

Private Declare Function TextOut _
               Lib "gdi32" _
               Alias "TextOutA" (ByVal hDC As Long, _
                                 ByVal X As Long, _
                                 ByVal Y As Long, _
                                 ByVal lpString As String, _
                                 ByVal nCount As Long) As Long

Private Declare Function CreateFont _
               Lib "gdi32" _
               Alias "CreateFontA" (ByVal H As Long, _
                                    ByVal W As Long, _
                                    ByVal E As Long, _
                                    ByVal O As Long, _
                                    ByVal W As Long, _
                                    ByVal i As Long, _
                                    ByVal u As Long, _
                                    ByVal s As Long, _
                                    ByVal C As Long, _
                                    ByVal OP As Long, _
                                    ByVal CP As Long, _
                                    ByVal q As Long, _
                                    ByVal PAF As Long, _
                                    ByVal f As String) As Long

Private Declare Function GetUpdateRect _
               Lib "user32" (ByVal hWnd As Long, _
                             lpRect As RECT, _
                             ByVal bErase As Long) As Long

'��ָ�����Դ���һ���߼�����
Private Declare Function CreateFontIndirect _
                Lib "gdi32" _
                Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700

'��ȡ����ĸ߶�,��ȡ���ֵĿ�Ȳ�׼
Private Declare Function GetTextExtentPoint32 _
               Lib "gdi32" _
               Alias "GetTextExtentPoint32A" (ByVal hDC As Long, _
                                              ByVal lpsz As String, _
                                              ByVal cbString As Long, _
                                              lpSize As Size) As Long

'nNumber*nNumerator/nDenominator �Զ��������롣�޷�����ķ���-1
Private Declare Function MulDiv _
               Lib "kernel32" (ByVal nNumber As Long, _
                               ByVal nNumerator As Long, _
                               ByVal nDenominator As Long) As Long
'˵��
'����ָ���豸����������豸�Ĺ��ܷ�����Ϣ
'���� ���ͼ�˵��
'hdc Long��Ҫ��ѯ���豸����Ϣ���豸����
'nIndex Long������GetDeviceCaps��������ʾ����ȷ��������Ϣ������

Private Declare Function GetDeviceCaps _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90

Private RGB_BLACK          As Long
Private RGB_RED            As Long
Private RGB_WRITE          As Long
Private RGB_BLUE          As Long
Private RGB_GRAY          As Long
Private RGB_FleetGRAY     As Long
'-------------------------------------------------------------------------------------------
Private Const mintCurveNullRow As Integer = 2
Private Const mlngBreathRowHeight As Long = 300 '�������߶�(�)
Public Const mlngWaveLeft As Long = 180 '���µ�չʾ��߾�(�)
Public Const mlngWaveTop As Long = 180 '���µ�չʾ�ұ߾�(�)

Private mbln������� As Boolean '�����Ƿ�Ϊ�����Ŀ
'��������
Private Type BasicStyle
    TitleText As String '��������
    TitleFont As String '��������
    TabRowHeight As Long 'һ����Ŀ���߶�
    TabDays As Long '�������
    TabDayTime As Long '������
    TabBeginTime As Long '��ʼʱ��
    TabTimeSplit As Long 'ʱ����
    TabTitleName As String 'һ����Ŀ����ͷ����
    ScaleColWidth As Long '�̶������ܿ��
    CurveColWidth As Long '��ͼ��������
    CurveRowHeight As Long '��ͼ������߶�
    AddCurveNull As Long '��ͼ����������ӵĿ�����(ֻ������ߣ�����Զ�������)
    DownTabRowHeight As Long '������Ŀ�����߶�
    AddTabNull As Long '������Ŀ�������ӵĿ�����
    BlnBaby As Boolean '�Ƿ�Ӥ�����µ�
End Type
Private TBasicStyle As BasicStyle
'���µ�(���߶ȡ��̶ȿ�ȡ����߱��߶ȿ�ȵȱ���)��λ����
Private Type WaveDrawStyle
    �ϱ��߶� As Long
    �±��߶� As Long
    �̶��ܿ�� As Long
    �����п�� As Long
    �����и߶� As Long
    ���������� As Long '��������¼��=3
    ���±�������� As Long
    �������߶� As Long '����Ϊ���ʱ��Ч
End Type
Private TWaveDrawStyle As WaveDrawStyle
'������Ŀ ��ʽ:��Ŀ���|��Ŀ���
Private mstrCurveItem As String
'�����Ŀ ��ʽ:��Ŀ���;��¼Ƶ��|��Ŀ���;��¼Ƶ��
Private mstrTabItem As String
'��ͼ�����DC
Private mlngDC As Long
Private mlngFont As Long
Private mlngOldFont As Long

'��ͼ��ʵ������(�)
Private mlngWaveWidth As Long
Private mlngWaveHeight As Long

Public Property Get WaveWidth() As Long
    WaveWidth = mlngWaveWidth
End Property

Public Property Get WaveHeight() As Long
    WaveHeight = mlngWaveHeight
End Property

'-------------------------------------------------------------------------------------------
'���µ���ʽ��ͼ
'-------------------------------------------------------------------------------------------
Public Function DrawWaveStyle(objPrint As Object, ByVal rsStyle As ADODB.Recordset, Optional ByVal blnExamples As Boolean = False, Optional ByRef mlngHeight As Long) As Boolean
'-----------------------------------------------------------
'���ܣ����ݹ�������µ���ʽ(�����ļ��ṹ)������µ������ͼ����
'����: objPrint ��ͼ�豸����,
'      rsStyle  ���µ���ʽ��¼��
'      blnExamples �Ƿ����"ʾ��"����,ר�����µ�Ĭ�ϲ������
'-----------------------------------------------------------
    Dim strTmp As String, strSQL As String
    Dim lngId As Long
    Dim arrItem() As String, arrCode() As String, lngIndex As Long, lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    Dim objFont As StdFont
    '��ͼ
    Dim lngWidth As Long, lngHeight As Long, lngLeft As Long, lngTop As Long
    Dim lngBrush As Long, lngOldBrush As Long
    Dim lngX As Long, lngY As Long, lngCurX As Long, lngCurY As Long
    Dim intBold As Integer, intFine As Integer
    '���߲���
    Dim lngCurveRows As Long, lngMaxValue As Long, lngMinValue As Long
    '���߱�񲿷�
    Dim lngTabRows As Long
    '������Ŀ
    Dim strCurveItem As String
    Dim rsCurve As New ADODB.Recordset
    
    On Error GoTo errHand
    mbln������� = False
    If rsStyle Is Nothing Or objPrint Is Nothing Then Exit Function
    If rsStyle.State = adStateClosed Then Exit Function
    
    If TypeName(objPrint) = "Printer" Then
        intBold = 4
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    
    RGB_BLACK = RGB(0, 0, 0)
    RGB_RED = RGB(255, 0, 0)
    RGB_WRITE = RGB(255, 255, 255)
    RGB_BLUE = RGB(0, 0, 255)
    RGB_GRAY = &H808080
    RGB_FleetGRAY = &HC0C0C0
    '------------------------------------------------------------------------------------
    '��һ��:��������׼��
    '------------------------------------------------------------------------------------
    With TBasicStyle
        .TitleText = "XX���µ�"
        .TitleFont = "����,9"
        .TabRowHeight = 225
        .TabDays = 7
        .TabDayTime = 6
        .TabBeginTime = 4
        .TabTimeSplit = 4
        .TabTitleName = "��    ��@סԺ����@����������&ʱ    ��"
        .ScaleColWidth = 1035
        .CurveColWidth = 180
        .CurveRowHeight = 90
        .AddCurveNull = 0
        .DownTabRowHeight = 225
        .AddTabNull = 0
    End With
    
    mstrCurveItem = "1"
    mstrTabItem = "1"
    '��ȡ��ʽ��������
    rsStyle.Filter = "��ID=NULL And �������=1 And �����ı�='��ʽ����'"
    If rsStyle.RecordCount > 0 Then
        lngId = rsStyle!ID
        rsStyle.Filter = "��ID=" & lngId
        Do While Not rsStyle.EOF
            Select Case "" & rsStyle!Ҫ������
            Case "�����ı�"
                TBasicStyle.TitleText = "" & rsStyle!�����ı�
            Case "��������"
                TBasicStyle.TitleFont = "" & rsStyle!�����ı�
                If TBasicStyle.TitleFont = "" Then TBasicStyle.TitleFont = "����,9"
            Case "���߶�"
                TBasicStyle.TabRowHeight = Val("" & rsStyle!�����ı�)
                If TBasicStyle.TabRowHeight < 225 Or TBasicStyle.TabRowHeight > 600 Then
                    TBasicStyle.TabRowHeight = 225
                End If
            Case "����"
                TBasicStyle.TabDays = Val("" & rsStyle!�����ı�)
                If TBasicStyle.TabDays = 0 Then TBasicStyle.TabDays = 7
            Case "������"
                If InStr(1, ",2,4,6,8,12,24,", "," & Val("" & rsStyle!�����ı�) & ",") = 0 Then
                    TBasicStyle.TabDayTime = 6
                Else
                    TBasicStyle.TabDayTime = Val("" & rsStyle!�����ı�)
                End If
            Case "��ʼʱ��"
                TBasicStyle.TabBeginTime = Val("" & rsStyle!�����ı�)
            Case "ʱ����"
                TBasicStyle.TabTimeSplit = Val("" & rsStyle!�����ı�)
            Case "��ͷ����"
                TBasicStyle.TabTitleName = "" & rsStyle!�����ı�
            Case "�̶ȿ��"
                TBasicStyle.ScaleColWidth = Val("" & rsStyle!�����ı�)
            Case "�����п�"
                TBasicStyle.CurveColWidth = Val("" & rsStyle!�����ı�)
            Case "�����и�"
                TBasicStyle.CurveRowHeight = Val("" & rsStyle!�����ı�)
            Case "���߿���"
                TBasicStyle.AddCurveNull = Val("" & rsStyle!�����ı�)
                If TBasicStyle.AddCurveNull < 0 Then TBasicStyle.AddCurveNull = 0
            Case "���߶�1"
                TBasicStyle.DownTabRowHeight = Val("" & rsStyle!�����ı�)
                If TBasicStyle.DownTabRowHeight < 225 Or TBasicStyle.TabRowHeight > 600 Then
                    TBasicStyle.DownTabRowHeight = 225
                End If
            Case "������"
                TBasicStyle.AddTabNull = Val("" & rsStyle!�����ı�)
                If TBasicStyle.AddTabNull < 0 Then TBasicStyle.AddTabNull = 0
            Case "Ӥ�����µ�"
                TBasicStyle.BlnBaby = "" & rsStyle!�����ı�
            End Select
        rsStyle.MoveNext
        Loop
    End If
    With TWaveDrawStyle
        .�ϱ��߶� = Fix(objPrint.ScaleX(TBasicStyle.TabRowHeight, vbTwips, vbPixels))
        .�±��߶� = Fix(objPrint.ScaleX(TBasicStyle.DownTabRowHeight, vbTwips, vbPixels))
        .�̶��ܿ�� = Fix(objPrint.ScaleX(TBasicStyle.ScaleColWidth, vbTwips, vbPixels))
        .�����и߶� = Fix(objPrint.ScaleX(TBasicStyle.CurveRowHeight, vbTwips, vbPixels))
        .�����п�� = Fix(objPrint.ScaleX(TBasicStyle.CurveColWidth, vbTwips, vbPixels))
        .�������߶� = Fix(objPrint.ScaleX(mlngBreathRowHeight, vbTwips, vbPixels))
    End With
    '��ȡ������Ŀ��Ϣ
    rsStyle.Filter = "��ID=NULL And �������=2 And �����ı�='������Ŀ����'"
    If rsStyle.RecordCount > 0 Then
        lngId = rsStyle!ID
        rsStyle.Filter = "��ID=" & lngId
        strTmp = ""
        Do While Not rsStyle.EOF
            strTmp = strTmp & "|" & "" & rsStyle!�����ı�
        rsStyle.MoveNext
        Loop
        If Left(strTmp, 1) = "|" Then strTmp = Mid(strTmp, 2)
        If strTmp <> "" Then mstrCurveItem = strTmp
    End If
    '��ȡ�����Ŀ��Ϣ
    rsStyle.Filter = "��ID=NULL And �������=3 And �����ı�='�����Ŀ����'"
    If rsStyle.RecordCount > 0 Then
        lngId = rsStyle!ID
        rsStyle.Filter = "��ID=" & lngId
        strTmp = ""
        Do While Not rsStyle.EOF
                strTmp = strTmp & "|" & "" & rsStyle!�����ı� & ";" & "" & rsStyle!Ҫ�ر�ʾ
        rsStyle.MoveNext
        Loop
        If Left(strTmp, 1) = "|" Then strTmp = Mid(strTmp, 2)
        If strTmp <> "" Then mstrTabItem = strTmp
    End If
    If mstrCurveItem = "" Then mstrCurveItem = "1"
    '��ȡ�󶨵���Ŀ
    strCurveItem = Replace(mstrCurveItem, "|", ",")
    If mstrTabItem = "" Then mstrTabItem = ";"
    arrItem = Split(mstrTabItem, "|")
    For lngIndex = 0 To UBound(arrItem)
        strCurveItem = strCurveItem & "," & Split(arrItem(lngIndex), ";")(0)
    Next lngIndex
    strSQL = _
        " SELECT /*+ RULE */" & vbNewLine & _
        "  A.��Ŀ���, A.�������,DECODE(A.��Ŀ���,4,'Ѫѹ',A.��¼��) ��¼��,A.��¼��,A.��¼��, A.��¼ɫ, NVL(A.���ֵ, 0) ���ֵ, NVL(A.��Сֵ, 0) ��Сֵ, NVL(A.��λֵ, 0) ��λֵ, A.�̶ȼ��, A.��ʾ��, B.��Ŀ��λ ��λ," & vbNewLine & _
        "  Decode(A.��¼��,3,A.�����,nvl(A.�����,2)-2) AS �����,DECODE(NVL(C.��Ŀ���,''),'',B.��Ŀ��ʾ,4) ��Ŀ��ʾ" & vbNewLine & _
        " FROM �����¼��Ŀ B, ���¼�¼��Ŀ A,��������Ŀ C" & vbNewLine & _
        " WHERE B.��Ŀ��� = A.��Ŀ��� And B.��Ŀ���=C.��Ŀ���(+) And B.��Ŀ���<>5  AND B.��Ŀ����=1 AND NOT (NVL(B.Ӧ�÷�ʽ,0)=2 And B.��Ŀ���=-1) AND EXISTS" & vbNewLine & _
        "  (SELECT 1 FROM TABLE(CAST(F_NUM2LIST([1]) AS ZLTOOLS.T_NUMLIST)) WHERE COLUMN_VALUE = B.��Ŀ���)" & vbNewLine & _
        " ORDER BY A.�������"
    Set rsCurve = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����������Ŀ", strCurveItem)
    '�������߱����Ŀ���ж�����
    lngCurveRows = 0
    rsCurve.Filter = "��Ŀ���=1"
    If rsCurve.RecordCount > 0 Then
        lngCurveRows = Val(NVL(rsCurve!�����, 0))
        If lngCurveRows < 0 Then lngCurveRows = 0
        lngMaxValue = Val(NVL(rsCurve!���ֵ, 0))
        If lngMaxValue < 42 Then lngMaxValue = 42
        lngMinValue = Val(NVL(rsCurve!��Сֵ, 0))
        If lngMinValue > 35 Then lngMinValue = 35
        '�̶�����������������Ŀ����
        lngCurveRows = lngCurveRows + ((lngMaxValue - lngMinValue) / 0.1) + mintCurveNullRow + TBasicStyle.AddCurveNull
        TWaveDrawStyle.���������� = lngCurveRows
    End If
    rsCurve.Filter = "��¼��=3 And ��Ŀ���<>1"
    rsCurve.Sort = "�������"
    Do While Not rsCurve.EOF
        lngMaxValue = Val(zlCommFun.NVL(rsCurve!���ֵ, 0))
        lngMinValue = Val(zlCommFun.NVL(rsCurve!��Сֵ, 0))
        lngRow = ((lngMaxValue - lngMinValue) / Val(NVL(rsCurve!��λֵ)))
        If Val(NVL(rsCurve!�����, 0)) > 0 Then lngRow = lngRow + Val(NVL(rsCurve!�����, 0))
        If lngRow Mod 2 = 1 Then lngRow = lngRow + 1
        lngCurveRows = lngCurveRows + lngRow
    rsCurve.MoveNext
    Loop
    '�������߱���ж�����
    lngTabRows = TBasicStyle.AddTabNull
    rsCurve.Filter = "��¼��=2 And ��Ŀ���<>5"
    mstrTabItem = ""
    Do While Not rsCurve.EOF
        For lngIndex = 0 To UBound(arrItem)
            arrCode = Split(arrItem(lngIndex), ";")
            If Val(arrCode(0)) = Val(rsCurve!��Ŀ���) Then
                If Val(rsCurve!��Ŀ���) = 3 Then '˵������Ϊ�����Ŀ
                    lngTabRows = lngTabRows + 1
                    mbln������� = True
                    arrCode(1) = TBasicStyle.TabDayTime
                Else
                    If InStr(1, GetTabFrequency(Val(TBasicStyle.TabDayTime), Val(rsCurve!��Ŀ��ʾ)), Val(arrCode(1))) = 0 Then
                        arrCode(1) = IIf(Val(TBasicStyle.TabDayTime) > 2, 2, 1)
                    End If
                    Select Case Val(arrCode(1))
                    Case 3
                        lngTabRows = lngTabRows + 3
                    Case 4
                        lngTabRows = lngTabRows + 2
                    Case Else
                        lngTabRows = lngTabRows + 1
                    End Select
                End If
                mstrTabItem = mstrTabItem & "|" & Join(arrCode, ";")
                Exit For
            End If
        Next lngIndex
    rsCurve.MoveNext
    Loop
    If Left(mstrTabItem, 1) = "|" Then mstrTabItem = Mid(mstrTabItem, 2)
    If mstrTabItem = "" Then mstrTabItem = ";"
    arrItem = Split(mstrTabItem, "|")
    
    TWaveDrawStyle.���±�������� = lngTabRows
    lngLeft = mlngWaveLeft: lngTop = mlngWaveTop  '��λ�
'    lngLeft=350
    '������
    lngWidth = objPrint.ScaleX(TWaveDrawStyle.�̶��ܿ��, vbPixels, vbTwips) + TBasicStyle.TabDays * (TBasicStyle.TabDayTime * objPrint.ScaleX(TWaveDrawStyle.�����п��, vbPixels, vbTwips))
    lngWidth = lngWidth + lngLeft * 2
    '����߶�
    lngHeight = 4 * objPrint.ScaleY(TWaveDrawStyle.�ϱ��߶�, vbPixels, vbTwips) + lngCurveRows * objPrint.ScaleY(TWaveDrawStyle.�����и߶�, vbPixels, vbTwips) + lngTabRows * objPrint.ScaleY(TWaveDrawStyle.�±��߶�, vbPixels, vbTwips)
    lngHeight = lngHeight - IIf(mbln������� = True, (objPrint.ScaleY(TWaveDrawStyle.�±��߶�, vbPixels, vbTwips) - objPrint.ScaleY(TWaveDrawStyle.�������߶�, vbPixels, vbTwips)), 0)
    Set objFont = New StdFont
    With objFont
        .Name = "����"
        .Size = 9
        .Bold = False: .Italic = False
    End With
    Set objPrint.Font = objFont
    lngHeight = lngHeight + objPrint.TextHeight("��") * 2
    arrItem = Split(TBasicStyle.TitleFont, ",")
    Set objFont = New StdFont
    With objFont
        .Name = arrItem(0)
        .Size = 9
        If UBound(arrItem) > 0 Then .Size = Val(arrItem(1))
        .Bold = False: .Italic = False
        If InStr(1, TBasicStyle.TitleFont, "��") > 0 Then .Bold = True
        If InStr(1, TBasicStyle.TitleFont, "б") > 0 Then .Italic = True
    End With
    Set objPrint.Font = objFont
    lngHeight = lngHeight + objPrint.TextHeight(TBasicStyle.TitleText) + lngTop * 2
    mlngWaveWidth = lngWidth - lngLeft * 2: mlngWaveHeight = lngHeight - lngTop * 2
    objPrint.Width = lngWidth: objPrint.Height = lngHeight
    '��ȡdcʱһ��Ҫע�����������Ŀ�ȸ߶Ⱥ��ڻ�ȡ�������ܻ�ͼ�ɹ�
    mlngDC = objPrint.hDC
    '------------------------------------------------------------------------------------
    '��һ��:��ʼ���л�ͼ����
    '------------------------------------------------------------------------------------
    '--ONE:�����ͼ�������
    T_ClientRect.Left = 0: T_ClientRect.Right = objPrint.Width
    T_ClientRect.Top = 0: T_ClientRect.Bottom = objPrint.Height
    '������ɫˢ��
    lngBrush = GetStockObject(WHITE_BRUSH)
    'ʹ�ø�ˢ����䱳��ɫ��ȫ�ף�
    lngOldBrush = SelectObject(mlngDC, lngBrush)
    Call FillRect(mlngDC, T_ClientRect, lngBrush)
    '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
    Call SelectObject(mlngDC, lngOldBrush)
    Call DeleteObject(lngBrush)
    Call SetTextColor(mlngDC, RGB_BLACK)
    '--TWO:һ����Ŀ����Ϣ�����
    lngX = objPrint.ScaleX(lngLeft, vbTwips, vbPixels)
    lngY = objPrint.ScaleY(lngTop, vbTwips, vbPixels)
    '�������
    Call SetFontIndirect(objPrint, TBasicStyle.TitleFont)
    T_Size.W = objPrint.ScaleX(objPrint.TextWidth(TBasicStyle.TitleText), vbTwips, vbPixels)
    T_Size.H = objPrint.ScaleY(objPrint.TextHeight(TBasicStyle.TitleText), vbTwips, vbPixels)
    lngCurX = 0: lngCurY = lngY + T_Size.H \ 2
    Call GetTextRect(objPrint, lngCurX, lngCurY, TBasicStyle.TitleText, objPrint.ScaleX(objPrint.Width, vbTwips, vbPixels), True)
    Call DrawText(mlngDC, TBasicStyle.TitleText, -1, T_LableRect, DT_CENTER)
    Call ReleaseFontIndirect(objPrint)
    '���˻�����Ϣ���
    lngCurX = lngX: lngCurY = T_LableRect.Bottom + objPrint.ScaleY(objPrint.TextHeight("1"), vbTwips, vbPixels) / 2
    strTmp = "����:'����:'�Ա�:'�Ʊ�:'����:'��Ժ����:'סԺ������:'���:"
    arrItem = Split(strTmp, "'")
    Call SetFontIndirect(objPrint, "����,9,��")
    T_Size.W = objPrint.ScaleX(objPrint.TextWidth(strTmp), vbTwips, vbPixels)
    T_Size.H = objPrint.ScaleY(objPrint.TextHeight(strTmp), vbTwips, vbPixels)
    lngY = lngCurY + T_Size.H
    lngWidth = (objPrint.ScaleX(objPrint.Width, vbTwips, vbPixels) - lngX * 2 - T_Size.W) / (UBound(arrItem) + 1)
    If lngWidth < 0 Then lngWidth = 0
    For lngIndex = 0 To UBound(arrItem)
        Call GetTextRect(objPrint, lngCurX, lngCurY, arrItem(lngIndex), 0, False)
        Call DrawText(mlngDC, arrItem(lngIndex), -1, T_LableRect, DT_CENTER)
        lngCurX = lngCurX + lngWidth + objPrint.ScaleX(objPrint.TextWidth(arrItem(lngIndex)), vbTwips, vbPixels)
    Next lngIndex
    Call ReleaseFontIndirect(objPrint)
    '������
    Call SetFontIndirect(objPrint, "����,9")
    lngCurX = lngX: lngCurY = lngY + objPrint.ScaleY(objPrint.TextHeight("1"), vbTwips, vbPixels) / 2
    lngY = lngCurY
    arrItem = Split(TBasicStyle.TabTitleName, "@")
    '���������
    For lngIndex = 0 To UBound(arrItem)
        T_Size.W = objPrint.ScaleX(objPrint.TextWidth(arrItem(lngIndex)), vbTwips, vbPixels)
        T_Size.H = objPrint.ScaleY(objPrint.TextHeight(arrItem(lngIndex)), vbTwips, vbPixels)
        lngHeight = (TWaveDrawStyle.�ϱ��߶� - T_Size.H) / 2
        If lngHeight < 0 Then lngHeight = 0
        Call GetTextRect(objPrint, lngCurX, lngCurY + lngHeight, arrItem(lngIndex), objPrint.ScaleY(TBasicStyle.ScaleColWidth, vbTwips, vbPixels), False)
        Call DrawText(mlngDC, arrItem(lngIndex), -1, T_LableRect, DT_CENTER)
        lngCurY = lngCurY + TWaveDrawStyle.�ϱ��߶�
    Next lngIndex
    '�ٻ��� (����)
    lngCurX = lngX: lngCurY = lngY
    lngWidth = TWaveDrawStyle.�̶��ܿ�� + TBasicStyle.TabDays * (TBasicStyle.TabDayTime * TWaveDrawStyle.�����п��) + lngX
    For lngIndex = 0 To UBound(arrItem) + 1
        Call WaveDrawLine(mlngDC, lngCurX, lngCurY, lngWidth, lngCurY, PS_SOLID, IIf(lngIndex = 0 Or lngIndex = UBound(arrItem) + 1, intBold, intFine), RGB_BLACK)
        lngCurY = lngCurY + TWaveDrawStyle.�ϱ��߶�
    Next lngIndex
    '(����)
    lngCurX = lngX: lngCurY = lngY
    lngHeight = lngCurY + TWaveDrawStyle.�ϱ��߶� * (UBound(arrItem) + 1)
    lngY = lngHeight
    Call WaveDrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngHeight, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = lngCurX + TWaveDrawStyle.�̶��ܿ��
    For lngIndex = 0 To TBasicStyle.TabDays
        Call WaveDrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngHeight, PS_SOLID, intBold, RGB_BLACK)
        lngCurX = lngCurX + TBasicStyle.TabDayTime * TWaveDrawStyle.�����п��
    Next lngIndex
    T_Size.H = objPrint.ScaleY(objPrint.TextHeight("1"), vbTwips, vbPixels)
    lngCurX = lngX + TWaveDrawStyle.�̶��ܿ��
    If T_Size.H > TWaveDrawStyle.�ϱ��߶� Then
        lngCurY = lngCurY + TWaveDrawStyle.�ϱ��߶� * UBound(arrItem)
    Else
        lngCurY = lngY - T_Size.H
    End If
    '�ڻ���ʱ��
    For lngIndex = 1 To TBasicStyle.TabDays
        For lngRow = 1 To TBasicStyle.TabDayTime
            strTmp = TBasicStyle.TabBeginTime + (lngRow - 1) * TBasicStyle.TabTimeSplit
            Call GetTextRect(objPrint, lngCurX, lngCurY, strTmp, TWaveDrawStyle.�����п��, False)
            Call DrawText(mlngDC, strTmp, -1, T_LableRect, DT_CENTER)
            lngCurX = lngCurX + TWaveDrawStyle.�����п��
            '����
            If Not lngRow = TBasicStyle.TabDayTime Then
                Call WaveDrawLine(mlngDC, lngCurX, lngY - TWaveDrawStyle.�ϱ��߶�, lngCurX, lngY, PS_SOLID, intFine, RGB_BLACK)
            End If
        Next lngRow
    Next lngIndex
    Call ReleaseFontIndirect(objPrint)
    
    lngCurX = lngX: lngCurY = lngY
    '--THREE:�����������Ļ���
    Call DrawCanvas(mlngDC, objPrint, rsCurve, lngCurX, lngCurY)
    lngY = lngY + lngCurveRows * TWaveDrawStyle.�����и߶�
    Call SetTextColor(mlngDC, RGB_BLACK)
    '--FOUR:������Ŀ���Ļ���
    lngCurX = lngX: lngCurY = lngY
    mlngHeight = lngY
    If TBasicStyle.BlnBaby = False Then
        Call DrawDownTab(mlngDC, objPrint, rsCurve, lngCurX, lngCurY)
    End If
    
    '--Five:��׼���µ����ʾ������
    If blnExamples = True Then
        Call SetTextColor(mlngDC, RGB_RED)
        Call SetFontIndirect(objPrint, "����,30")
        strTmp = "ʾ������"
        Call GetTextRect(objPrint, 0, objPrint.ScaleY(mlngWaveHeight, vbTwips, vbPixels) \ 3, strTmp, objPrint.ScaleX(mlngWaveWidth, vbTwips, vbPixels) + 10, True)
        Call DrawText(mlngDC, strTmp, -1, T_LableRect, DT_CENTER)
        Call ReleaseFontIndirect(objPrint)
        Call WaveDrawLine(mlngDC, T_LableRect.Left, T_LableRect.Top, T_LableRect.Right, T_LableRect.Top, PS_SOLID, 2, RGB_RED)
        Call WaveDrawLine(mlngDC, T_LableRect.Left, T_LableRect.Top, T_LableRect.Left, T_LableRect.Bottom, PS_SOLID, 2, RGB_RED)
        Call WaveDrawLine(mlngDC, T_LableRect.Left, T_LableRect.Bottom, T_LableRect.Right, T_LableRect.Bottom, PS_SOLID, 2, RGB_RED)
        Call WaveDrawLine(mlngDC, T_LableRect.Right, T_LableRect.Top, T_LableRect.Right, T_LableRect.Bottom, PS_SOLID, 2, RGB_RED)
        Call SetTextColor(mlngDC, RGB_BLACK)
    End If
    
    DrawWaveStyle = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function DrawCanvas(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsTemp As ADODB.Recordset, ByVal lngX As Long, ByVal lngY As Long) As String
'------------------------------------------------------------------------------------------------------
'����:���̶������������������̶�ֵ��Ϣ
'����:lngDC ��ͼ�����DC��objDraw �滭����.rsTemp:����������Ŀ��¼��(A.��Ŀ���,A.�������,A.��¼��,A.��¼��,A.��¼ɫ,A.���ֵ,A.��Сֵ,A.��λֵ,C.��Ŀ��λ ��λ,A.�����-2 AS �����,B.��λ)
'����:���ظ������ߵľ�����Ϣ����( "��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��ɫ")
'����˵����Ϣ(��Ŀ�ķ���)
'-------------------------------------------------------------------------------------------------------
    Dim str˵�� As String
    Dim lngMaxX     As Long, lngMaxY As Long  '�߽�
    Dim lngCurX As Long, lngCurY As Long
    Dim sinCurAlerY As Single '������
    Dim lngRow      As Long
    Dim intLables   As Integer
    Dim lngCurveRows As Long '���ߵ�����
    Dim bln˫�� As Boolean                  '�˲������û�ָ��,bln˫��=TRUE��ʾֻ��ʾ����;������ʾʮ��
    '���¶��Ǳ�׼�߶�
    Dim intLineMode   As Integer
    Dim sinAlertness  As Single              '������,��������
    Dim lngLableStep  As Long
    Dim lngColStep    As Long
    Dim lngRowStep As Long
    Dim arrTemp()     As String
    Dim intBold As Integer, intFine As Integer
    Dim sinY��λ As Single '���ߵ�λ�����Bottom

    '�������ͼ�������(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
    Dim sin�̶� As Single, bln��ʾ�̶� As Boolean, blnFirst As Boolean
    Dim sin�̶ȼ�� As Single, sinBegin�̶� As Single, dbl��λֵ As Double

    On Error GoTo errHand
    If TypeName(objDraw) = "Printer" Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    '------------------------------------------------------------------------------------------------------------------
    '����ֵ
    str˵�� = ""
    intLineMode = PS_SOLID
    bln˫�� = True
    
    lngColStep = TWaveDrawStyle.�����п��
    lngRowStep = TWaveDrawStyle.�����и߶�
    '��һ��������ɼ�¼��=1(����)�����
    rsTemp.Filter = "��¼��=1"
    intLables = rsTemp.RecordCount
    lngLableStep = TWaveDrawStyle.�̶��ܿ�� \ intLables
    
    lngCurX = lngX: lngCurY = lngY
    lngMaxX = lngCurX + TWaveDrawStyle.�̶��ܿ�� + TBasicStyle.TabDays * TBasicStyle.TabDayTime * TWaveDrawStyle.�����п��
    lngMaxY = TWaveDrawStyle.���������� * lngRowStep + lngCurY
    '�Ȼ��̶�����
    For lngRow = 1 To intLables
        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow = 1, intBold, intFine), RGB_BLACK)
        lngCurX = lngCurX + lngLableStep
        Call WaveDrawLine(lngDC, lngCurX - lngLableStep, lngMaxY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        If lngRow = intLables Then
            lngCurX = TWaveDrawStyle.�̶��ܿ�� + lngX
            Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        End If
    Next
    'Ĭ�����һ��������ʾ��Ŀ����
    lngCurY = lngCurY + lngRowStep * mintCurveNullRow
    lngCurX = lngX
    Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, PS_SOLID, intFine, RGB_BLACK)
    '�����µ�������
    lngCurX = lngX + TWaveDrawStyle.�̶��ܿ��
    lngCurveRows = TWaveDrawStyle.���������� - mintCurveNullRow
    For lngRow = 1 To lngCurveRows
        lngCurY = lngCurY + lngRowStep
        '�����µ���������
        If (bln˫�� And lngRow Mod 2 = 0) Or Not bln˫�� Or lngRow = lngCurveRows Then
            Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0 Or lngRow = lngCurveRows, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 Or lngRow = lngCurveRows, intBold, intFine), RGB_BLACK)
        End If
    Next
    lngCurY = lngY
    '�����µ�������
    For lngRow = 1 To TBasicStyle.TabDays * TBasicStyle.TabDayTime
        lngCurX = lngCurX + lngColStep
        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod TBasicStyle.TabDayTime = 0, intBold, intFine), IIf(lngRow Mod TBasicStyle.TabDayTime = 0, RGB_RED, RGB_BLACK))
    Next
    '���̶ȿ�ı�ߣ��ӹ̶������10�п�ʼ��ʶ��
    intLables = 1
    rsTemp.Filter = "��¼��=1"
    rsTemp.Sort = "�������"
    lngCurX = lngX
    With rsTemp
        Do While Not .EOF
            '��ʾ�̶ȿ���Ŀ�����Ƽ�����,�����¡�
            If intLables = rsTemp.RecordCount Then
                lngLableStep = TWaveDrawStyle.�̶��ܿ�� - ((intLables - 1) * lngLableStep)
            End If
            lngCurX = lngX + ((intLables - 1) * lngLableStep)
            lngCurY = lngY
            '���������Ŀ������
            Call SetFontIndirect(objDraw, "����,9")
            Call SetTextColor(lngDC, NVL(!��¼ɫ, RGB_BLACK))
            Call GetTextRect(objDraw, lngCurX, lngCurY + objDraw.ScaleY(objDraw.TextHeight(NVL(!��¼��)), vbTwips, vbPixels) \ 2, Trim(NVL(!��¼��)), lngLableStep)
            Call DrawText(lngDC, Trim(NVL(!��¼��)), -1, T_LableRect, DT_CENTER)
            Call ReleaseFontIndirect(objDraw)
            '�����Ŀ��λ
            If Trim(NVL(!��λ)) <> "" Then
                Call SetFontIndirect(objDraw, "����,8")
                Call GetTextRect(objDraw, lngCurX, lngCurY + lngRowStep * mintCurveNullRow + objDraw.ScaleY(objDraw.TextHeight(NVL(!��λ)), vbTwips, vbPixels) \ 2, Trim(NVL(!��λ)), lngLableStep)
                Call DrawText(lngDC, Trim(NVL(!��λ, 0)), -1, T_LableRect, DT_CENTER)
                Call ReleaseFontIndirect(objDraw)
                sinY��λ = T_LableRect.Bottom
            Else
                sinY��λ = lngY + lngRowStep * mintCurveNullRow
            End If
            
            intLables = intLables + 1
            Call SetFontIndirect(objDraw, "����,9")
            'ǿ���趨����������Ŀ����ʾģʽ
            Select Case !��Ŀ���
                Case 1  '��������ʱ����̶�
                    sin�̶ȼ�� = NVL(!�̶ȼ��, 1)
                    dbl��λֵ = 0.1
                    sinAlertness = NVL(!��ʾ��, 37)
                    arrTemp = Split(NVL(!��¼��, "��,��,��"), ",")
                    str˵�� = str˵�� & "��" & NVL(!��¼��) & "(����" & arrTemp(0) & ",Ҹ��" & arrTemp(1) & ",����" & arrTemp(2) & ")"
                Case 2, -1  '����/������10�ı�������̶�
                    sin�̶ȼ�� = NVL(!�̶ȼ��, 10)
                    dbl��λֵ = 2
                    sinAlertness = NVL(!��ʾ��, 0)
                    If !��Ŀ��� = 2 Then
                        str˵�� = str˵�� & "��" & NVL(!��¼��) & "(ȱʡ��¼��" & NVL(!��¼��, "+") & ",����H)"
                    Else
                        str˵�� = str˵�� & "��" & NVL(!��¼��) & "(" & NVL(!��¼��, "��") & ")"
                    End If
                Case 3  '������5�ı�������̶�
                    dbl��λֵ = 1
                    sin�̶ȼ�� = NVL(!�̶ȼ��, 5)
                    sinAlertness = NVL(!��ʾ��, 0)
                    str˵�� = str˵�� & "��" & NVL(!��¼��) & "(��������" & NVL(!��¼��, "*") & ",������R)"
                Case Else
                    dbl��λֵ = Val(NVL(!��λֵ, 0))
                    sin�̶ȼ�� = NVL(!�̶ȼ��, Val(NVL(!��λֵ, 0)) * 10)
                    If sin�̶ȼ�� > Val(NVL(!���ֵ)) - Val(NVL(!��Сֵ)) Then
                        sin�̶ȼ�� = Val(NVL(!���ֵ)) - Val(NVL(!��Сֵ))
                    End If
                    sinAlertness = NVL(!��ʾ��, 0)
                    str˵�� = str˵�� & "��" & NVL(!��¼��) & "(" & NVL(!��¼��, "*") & ")"
            End Select
            '����ֵ
            lngCurY = lngCurY + lngRowStep * mintCurveNullRow '�̶�ǰ2�еĸ߶Ȳ�����̶�
            '��������ж�λ����Чλ��
            lngCurY = lngCurY + (TWaveDrawStyle.�����и߶� * Val(NVL(!�����, 0)))
            blnFirst = False
            Do While True
                bln��ʾ�̶� = False
                If blnFirst = False Then    '�ս���ѭ������ʱȡ�����ֵ
                    sin�̶� = NVL(!���ֵ, 0)
                    sinBegin�̶� = sin�̶�
                    blnFirst = True
                Else                    '����õ�ÿ���̶ȵ�ֵ
                    sin�̶� = sin�̶� - dbl��λֵ     '���Ŀǰ��ʾģʽΪ˫������˫���ۼ�
                End If

                '�������õĿ̶ȼ����ʾ�̶�ֵ
                If Val(Format(sin�̶�, "#0.00")) = Val(Format(sinBegin�̶�, "#0.00")) Then bln��ʾ�̶� = True
                If bln��ʾ�̶� = True Or sin�̶� < sinBegin�̶� Then sinBegin�̶� = sinBegin�̶� - sin�̶ȼ��
                If sinBegin�̶� < Val(NVL(!��Сֵ)) Then sinBegin�̶� = Val(NVL(!��Сֵ))

                If bln��ʾ�̶� Then
                    '�������ֵ�������ߵ�λ�ظ�
                    If sin�̶� = Val(NVL(!���ֵ, 0)) And lngCurY <= sinY��λ Then
                        Call GetTextRect(objDraw, lngCurX, sinY��λ + IIf(lngCurY = sinY��λ, (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), 0), Val(Format(sin�̶�, "#0.0")), lngLableStep)
                    ElseIf Format(lngCurY, "#0") = lngMaxY Then
                        Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin�̶�, "#0.0")), lngLableStep)
                    Else
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Val(Format(sin�̶�, "#0.0")), lngLableStep)
                    End If
                    Call DrawText(lngDC, Val(Format(sin�̶�, "#0.0")), -1, T_LableRect, DT_CENTER)
                End If
                If Val(Format(sin�̶�, "#0.00")) <= Val(Format(NVL(!��Сֵ), "#0.00")) Or Format(lngCurY, "#0") > lngMaxY Then
                    '���������
                    If sinAlertness > Val(NVL(!��Сֵ)) And sinAlertness < Val(NVL(!���ֵ)) Then
                        '�������ֵ�뵱ǰֵ֮��Ĳ��,�Լ���Сֵ,����õ������ٸ��̶�,�ٸ��ݵ�λ�̶ȵõ�ʵ������
                        sinCurAlerY = Format((Val(NVL(!���ֵ)) - sinAlertness) / dbl��λֵ * TWaveDrawStyle.�����и߶�, "#0.0")
                        sinCurAlerY = Format(sinCurAlerY + lngY + lngRowStep * mintCurveNullRow + (TWaveDrawStyle.�����и߶� * Val(NVL(!�����, 0))), "#0")
                        Call WaveDrawLine(lngDC, lngX + TWaveDrawStyle.�̶��ܿ��, CLng(sinCurAlerY), lngMaxX, CLng(sinCurAlerY), PS_SOLID, intBold, RGB_RED)
                    End If
         
                    Exit Do
                End If
                lngCurY = lngCurY + TWaveDrawStyle.�����и߶�
            Loop
            Call ReleaseFontIndirect(objDraw)
            sinBegin�̶� = 0
            sin�̶� = 0
            .MoveNext
        Loop
    End With
    '��ɶ������߲��ֵ����
    rsTemp.Filter = "��¼��=3"
    rsTemp.Sort = "�������"
    With rsTemp
        Do While Not .EOF
            lngY = lngMaxY
            lngCurY = lngY
            lngCurX = lngX
            lngCurveRows = ((Val(NVL(!���ֵ, 0)) - Val(NVL(!��Сֵ, 0))) / Val(NVL(!��λֵ)))
            If Val(NVL(!�����, 0)) > 0 Then lngCurveRows = lngCurveRows + Val(NVL(!�����, 0))
            If lngCurveRows Mod 2 = 1 Then lngCurveRows = lngCurveRows + 1
            If lngCurveRows > 0 Then
                lngMaxY = lngCurveRows * lngRowStep + lngCurY
                '��ɿ̶�����Ļ���
                Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                Call WaveDrawLine(lngDC, lngCurX + TWaveDrawStyle.�̶��ܿ��, lngCurY, lngCurX + TWaveDrawStyle.�̶��ܿ��, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                Call WaveDrawLine(lngDC, lngCurX, lngMaxY, lngCurX + TWaveDrawStyle.�̶��ܿ��, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                '��������еĻ���
                lngCurX = lngX + TWaveDrawStyle.�̶��ܿ��
                For lngRow = 1 To lngCurveRows
                    lngCurY = lngCurY + lngRowStep
                    '�����µ���������
                    If (bln˫�� And lngRow Mod 2 = 0) Or Not bln˫�� Or lngRow = lngCurveRows Then
                        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0 Or lngRow = lngCurveRows, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 Or lngRow = lngCurveRows, intBold, intFine), RGB_BLACK)
                    End If
                Next
                lngCurY = lngY
                '��������еĻ���
                For lngRow = 1 To TBasicStyle.TabDays * TBasicStyle.TabDayTime
                    lngCurX = lngCurX + lngColStep
                    Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod TBasicStyle.TabDayTime = 0, intBold, intFine), IIf(lngRow Mod TBasicStyle.TabDayTime = 0, RGB_RED, RGB_BLACK))
                Next
                '�����Ŀ���ƺͿ̶ȵ����
                lngCurX = lngX: lngCurY = lngY
                '���������Ŀ������
                Call SetFontIndirect(objDraw, "����,9")
                Call SetTextColor(lngDC, NVL(!��¼ɫ, RGB_BLACK))
                T_Size.H = objDraw.ScaleY(objDraw.TextHeight("��"), vbTwips, vbPixels)
                If T_Size.H * Len(NVL(!��¼��)) >= lngCurveRows * lngRowStep Then
                    lngCurY = lngY
                Else
                    lngCurY = lngY + ((lngCurveRows * lngRowStep) - (T_Size.H * Len(NVL(!��¼��)))) \ 2
                End If
                For lngRow = 1 To Len(NVL(!��¼��))
                    Call GetTextRect(objDraw, lngCurX, lngCurY, Mid(NVL(!��¼��), lngRow, 1), TWaveDrawStyle.�̶��ܿ�� \ 2, False)
                    Call DrawText(lngDC, Mid(NVL(!��¼��), lngRow, 1), -1, T_LableRect, DT_CENTER)
                    lngCurY = lngCurY + T_Size.H
                Next lngRow
                Call ReleaseFontIndirect(objDraw)
                '�����Ŀ��λ
                lngCurY = lngY: If NVL(!��¼��) <> "" Then lngCurX = T_LableRect.Right
                If Trim(NVL(!��λ)) <> "" And NVL(!��¼��) <> "" Then
                    Call SetFontIndirect(objDraw, "����,8")
                    T_Size.H = objDraw.ScaleY(objDraw.TextHeight("��"), vbTwips, vbPixels)
                    If T_Size.H * Len(Trim(NVL(!��λ))) >= lngCurveRows * lngRowStep Then
                        lngCurY = lngY
                    Else
                        lngCurY = lngY + ((lngCurveRows * lngRowStep) - (T_Size.H * Len(NVL(!��λ)))) \ 2
                    End If
                    For lngRow = 1 To Len(Trim(NVL(!��λ)))
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Mid(Trim(NVL(!��λ)), lngRow, 1), 0, False)
                        Call DrawText(lngDC, Mid(Trim(NVL(!��λ)), lngRow, 1), -1, T_LableRect, DT_CENTER)
                        lngCurY = lngCurY + T_Size.H
                    Next lngRow
                    Call ReleaseFontIndirect(objDraw)
                End If
                Call SetFontIndirect(objDraw, "����,9")
                dbl��λֵ = Val(NVL(!��λֵ, 0))
                sin�̶ȼ�� = NVL(!�̶ȼ��, Val(NVL(!��λֵ, 0)) * 10)
                If sin�̶ȼ�� > Val(NVL(!���ֵ)) - Val(NVL(!��Сֵ)) Then
                    sin�̶ȼ�� = Val(NVL(!���ֵ)) - Val(NVL(!��Сֵ))
                End If
                sinAlertness = NVL(!��ʾ��, 0)
                str˵�� = str˵�� & "��" & NVL(!��¼��) & "(" & NVL(!��¼��, "*") & ")"
                lngCurY = lngY + (TWaveDrawStyle.�����и߶� * Val(NVL(!�����, 0)))
                blnFirst = False
                Do While True
                    bln��ʾ�̶� = False
                    If blnFirst = False Then     '�ս���ѭ������ʱȡ�����ֵ
                        sin�̶� = NVL(!���ֵ, 0)
                        sinBegin�̶� = sin�̶�
                        blnFirst = True
                    Else                    '����õ�ÿ���̶ȵ�ֵ
                        sin�̶� = sin�̶� - dbl��λֵ     '���Ŀǰ��ʾģʽΪ˫������˫���ۼ�
                    End If
    
                    '�������õĿ̶ȼ����ʾ�̶�ֵ
                    If Val(Format(sin�̶�, "#0.00")) = Val(Format(sinBegin�̶�, "#0.00")) Then bln��ʾ�̶� = True
                    If bln��ʾ�̶� = True Or sin�̶� < sinBegin�̶� Then sinBegin�̶� = sinBegin�̶� - sin�̶ȼ��
                    If sinBegin�̶� < Val(NVL(!��Сֵ)) Then sinBegin�̶� = Val(NVL(!��Сֵ))
    
                    If bln��ʾ�̶� Then
                        '�������ֵ�������ߵ�λ�ظ�
                        lngCurX = lngX + TWaveDrawStyle.�̶��ܿ�� - objDraw.ScaleX(objDraw.TextWidth(Val(Format(sin�̶�, "#0.0"))), vbTwips, vbPixels)
                        lngCurX = lngCurX - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 3)
                        If sin�̶� = Val(NVL(!���ֵ, 0)) And lngCurY = lngY Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY + (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin�̶�, "#0.0")))
                        ElseIf Format(lngCurY, "#0") = lngMaxY Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin�̶�, "#0.0")))
                        Else
                            Call GetTextRect(objDraw, lngCurX, lngCurY, Val(Format(sin�̶�, "#0.0")))
                        End If
                        Call DrawText(lngDC, Val(Format(sin�̶�, "#0.0")), -1, T_LableRect, DT_CENTER)
                    End If
                    If Val(Format(sin�̶�, "#0.00")) <= Val(Format(NVL(!��Сֵ), "#0.00")) Or Format(lngCurY, "#0") > lngMaxY Then
                        '���������
                        If sinAlertness > Val(NVL(!��Сֵ)) And sinAlertness < Val(NVL(!���ֵ)) Then
                            '�������ֵ�뵱ǰֵ֮��Ĳ��,�Լ���Сֵ,����õ������ٸ��̶�,�ٸ��ݵ�λ�̶ȵõ�ʵ������
                            sinCurAlerY = Format((Val(NVL(!���ֵ)) - sinAlertness) / dbl��λֵ * TWaveDrawStyle.�����и߶�, "#0.0")
                            sinCurAlerY = Format(sinCurAlerY + lngY + (TWaveDrawStyle.�����и߶� * Val(NVL(!�����, 0))), "#0")
                            Call WaveDrawLine(lngDC, lngX + TWaveDrawStyle.�̶��ܿ��, CLng(sinCurAlerY), lngMaxX, CLng(sinCurAlerY), PS_SOLID, intBold, RGB_RED)
                        End If
                        Exit Do
                    End If
                    lngCurY = lngCurY + TWaveDrawStyle.�����и߶�
                Loop
                Call ReleaseFontIndirect(objDraw)
                sinBegin�̶� = 0
                sin�̶� = 0
            End If
        .MoveNext
        Loop
    End With
    str˵�� = "˵��:" & Mid(str˵��, 2)
    DrawCanvas = str˵��
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub DrawDownTab(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsTemp As ADODB.Recordset, ByVal lngX As Long, ByVal lngY As Long)
'-----------------------------------------------------
'����: ��ɱ��±�����ݵ����
'-----------------------------------------------------
    Dim lngCurX As Long, lngCurY As Long, lngRowTopY As Long
    Dim lngMaxX As Long, lngMaxY As Long
    Dim lngRow As Long, lngHeight As Long, lngWidth As Long
    Dim lngDayTime As Long, lngRecordTime As Long, lngRowTime As Long
    Dim lngTimeCOlWidth As Long 'Ƶ���п�
    Dim strContent As String '��ͷ����
    Dim intBold As Integer, intFine As Integer '��������
    Dim arrItem() As String '�󶨵���Ŀ����
    Dim i As Long, j As Long, k As Long
    If TWaveDrawStyle.���±�������� = 0 Then Exit Sub
    
    On Error GoTo errHand
    If TypeName(objDraw) = "Printer" Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    
    lngCurX = lngX: lngCurY = lngY
    lngMaxX = lngCurX + TWaveDrawStyle.�̶��ܿ�� + TBasicStyle.TabDays * TBasicStyle.TabDayTime * TWaveDrawStyle.�����п��
    lngMaxY = lngCurY + TWaveDrawStyle.���±�������� * TWaveDrawStyle.�±��߶� - IIf(mbln������� = True, (TWaveDrawStyle.�±��߶� - TWaveDrawStyle.�������߶�), 0)
    
    '����ɱ����߿�Ļ���
    '������(��ͷ����)
    Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = lngCurX + TWaveDrawStyle.�̶��ܿ��
    Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
    '������(���岿��)
    For lngRow = 1 To TBasicStyle.TabDays
        lngCurX = lngCurX + TBasicStyle.TabDayTime * TWaveDrawStyle.�����п��
        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
    Next lngRow
    lngCurX = lngX + TWaveDrawStyle.�̶��ܿ��
    lngCurY = lngY
    '������(���岿��)
    For lngRow = 1 To TWaveDrawStyle.���±��������
        If mbln������� = True And lngRow = 1 Then
            lngCurY = lngCurY + TWaveDrawStyle.�������߶�
        Else
            lngCurY = lngCurY + TWaveDrawStyle.�±��߶�
        End If
        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, PS_SOLID, IIf(lngRow = TWaveDrawStyle.���±��������, intBold, intFine), RGB_BLACK)
    Next lngRow
    Call WaveDrawLine(lngDC, lngX, lngMaxY, lngX + TWaveDrawStyle.�̶��ܿ��, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = lngX: lngCurY = lngY
    Call SetFontIndirect(objDraw, "����,9")
    Call SetTextColor(lngDC, RGB_BLACK)
    '������ɺ����������(�������Ϊ�����Ŀ)
    If mbln������� = True Then
        rsTemp.Filter = "��Ŀ���=3"
        lngHeight = TWaveDrawStyle.�������߶�
        lngWidth = TWaveDrawStyle.�̶��ܿ��
        strContent = NVL(rsTemp!��¼��) & IIf(Trim(NVL(rsTemp!��λ)) <> "", "(" & Trim(NVL(rsTemp!��λ)) & ")", "")
        T_Size.H = objDraw.ScaleY(objDraw.TextHeight(strContent), vbTwips, vbPixels)
        T_Size.W = objDraw.ScaleX(objDraw.TextWidth(strContent), vbTwips, vbPixels)
        If T_Size.H < lngHeight Then
            lngCurY = lngCurY + (lngHeight - T_Size.H) \ 2
        End If
        Call GetTextRect(objDraw, lngCurX, lngCurY, strContent, lngWidth, False)
        '������������
        If T_LableRect.Left < lngX Then T_LableRect.Left = lngX
        If T_LableRect.Top < lngY Then T_LableRect.Top = lngY
        If T_LableRect.Right > lngWidth + lngX Then T_LableRect.Right = lngWidth + lngX
        If T_LableRect.Bottom > lngHeight + lngY Then T_LableRect.Bottom = lngHeight + lngY
        Call DrawText(lngDC, strContent, -1, T_LableRect, DT_CENTER)
        If TWaveDrawStyle.���±�������� > 1 Then
            lngCurY = lngY + lngHeight
            Call WaveDrawLine(lngDC, lngX, lngCurY, lngX + TWaveDrawStyle.�̶��ܿ��, lngCurY, PS_SOLID, intFine, RGB_BLACK)
        End If
        '���Ƶ�εĻ���
        lngCurY = lngY
        For lngRow = 1 To TBasicStyle.TabDays
            lngCurX = lngX + TWaveDrawStyle.�̶��ܿ�� + (lngRow - 1) * TBasicStyle.TabDayTime * TWaveDrawStyle.�����п��
            For lngDayTime = 1 To TBasicStyle.TabDayTime - 1
                lngCurX = lngCurX + TWaveDrawStyle.�����п��
                Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngCurY + lngHeight, PS_SOLID, intFine, RGB_BLACK)
            Next lngDayTime
        Next lngRow
    End If
    '����������
    arrItem = Split(mstrTabItem, "|")
    lngCurY = lngY + IIf(mbln������� = True, TWaveDrawStyle.�������߶�, 0)
    lngRowTopY = lngCurY
    lngCurX = lngX
    rsTemp.Filter = "��¼��=2 And ��Ŀ���<>3 And ��Ŀ���<>5"
    rsTemp.Sort = "�������"
    With rsTemp
        Do While Not .EOF
            For lngRow = 0 To UBound(arrItem)
                If Val(Split(arrItem(lngRow), ";")(0)) = Val(!��Ŀ���) Then
                    lngRecordTime = Val(Split(arrItem(lngRow), ";")(1))
                    Select Case lngRecordTime '��¼Ƶ��
                    Case 3
                        lngRowTime = 1 'ÿһ�б������
                        lngDayTime = 3 'ռ�õı������
                    Case 4
                        lngRowTime = 2 'ÿһ�б������
                        lngDayTime = 2 'ռ�õı������
                    Case Else
                        lngRowTime = lngRecordTime 'ÿһ�б������
                        lngDayTime = 1 'ռ�õı������
                    End Select
                    '�����ͷ����
                    lngHeight = TWaveDrawStyle.�±��߶� * lngDayTime
                    lngWidth = TWaveDrawStyle.�̶��ܿ��
                    strContent = NVL(rsTemp!��¼��) & IIf(Trim(NVL(rsTemp!��λ)) <> "", "(" & Trim(NVL(rsTemp!��λ)) & ")", "")
                    T_Size.H = objDraw.ScaleY(objDraw.TextHeight(strContent), vbTwips, vbPixels)
                    T_Size.W = objDraw.ScaleX(objDraw.TextWidth(strContent), vbTwips, vbPixels)
                    If T_Size.H < lngHeight Then
                        lngCurY = lngCurY + (lngHeight - T_Size.H) \ 2
                    End If
                    Call GetTextRect(objDraw, lngCurX, lngCurY, strContent, lngWidth, False)
                    '������������
                    If T_LableRect.Left < lngX Then T_LableRect.Left = lngX
                    If T_LableRect.Top < lngRowTopY Then T_LableRect.Top = lngRowTopY
                    If T_LableRect.Right > lngWidth + lngX Then T_LableRect.Right = lngWidth + lngX
                    If T_LableRect.Bottom > lngHeight + lngRowTopY Then T_LableRect.Bottom = lngHeight + lngRowTopY
                    Call DrawText(lngDC, strContent, -1, T_LableRect, DT_CENTER)
                    If Not rsTemp.RecordCount = rsTemp.AbsolutePosition Then
                        lngCurY = lngRowTopY + lngHeight
                        Call WaveDrawLine(lngDC, lngX, lngCurY, lngX + TWaveDrawStyle.�̶��ܿ��, lngCurY, PS_SOLID, intFine, RGB_BLACK)
                    End If
                    lngCurY = lngRowTopY
                    '����ÿһ��Ƶ��ռ�õ��п�
                    lngTimeCOlWidth = (TBasicStyle.TabDayTime * TWaveDrawStyle.�����п��) \ lngRowTime
                    '���Ʊ����
                    For i = 1 To lngDayTime 'ռ�õı������
                        lngCurX = lngX + TWaveDrawStyle.�̶��ܿ��
                        lngCurY = lngRowTopY + (i - 1) * TWaveDrawStyle.�±��߶�
                        For j = 1 To TBasicStyle.TabDays '��������
                            lngCurX = lngX + TWaveDrawStyle.�̶��ܿ�� + (j - 1) * TBasicStyle.TabDayTime * TWaveDrawStyle.�����п��
                            For k = 1 To lngRowTime - 1 'ÿһ�б������
                                lngCurX = lngCurX + lngTimeCOlWidth
                                Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngCurY + TWaveDrawStyle.�±��߶�, PS_SOLID, intFine, RGB_BLACK)
                            Next k
                        Next j
                    Next i
                    lngCurX = lngX
                    lngRowTopY = lngRowTopY + lngHeight
                    lngCurY = lngRowTopY
                    Exit For
                End If
            Next lngRow
        .MoveNext
        Loop
    End With
    
    '�����ɿ��еĻ���
    lngRecordTime = TWaveDrawStyle.���±�������� - TBasicStyle.AddTabNull '���е���ʼ��
    lngCurX = lngX
    lngCurY = lngY + IIf(mbln������� = True, TWaveDrawStyle.�������߶�, 0) + (lngRecordTime - IIf(mbln������� = True, 1, 0)) * TWaveDrawStyle.�±��߶�
    For lngRow = lngRecordTime To TWaveDrawStyle.���±�������� - 1
        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngX + TWaveDrawStyle.�̶��ܿ��, lngCurY, PS_SOLID, intFine, RGB_BLACK)
        lngCurY = lngCurY + TWaveDrawStyle.�±��߶�
    Next lngRow
    Call ReleaseFontIndirect(objDraw)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
'-------------------------------------------------------------------------------------------
Private Sub SetFontIndirect(ByVal objDraw As Object, ByVal strFontInfo As String)
    '����:������������,��ʹ��
    Dim BFileName() As Byte
    Dim i As Integer
    Dim stdset As StdFont

    On Error GoTo errHand
    Set stdset = New StdFont
    With stdset
        .Name = Split(strFontInfo, ",")(0)
        .Size = 9
        If UBound(Split(strFontInfo, ",")) > 0 Then .Size = Val(Split(strFontInfo, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strFontInfo, "��") > 0 Then .Bold = True
        If InStr(1, strFontInfo, "б") > 0 Then .Italic = True
    End With

    Set objDraw.Font = stdset
    BFileName = StrConv(stdset.Name, vbFromUnicode)
    With T_Font
        For i = 1 To Len(stdset.Name)
            .lfFaceName(i - 1) = BFileName(i - 1)
        Next i
        .lfHeight = -MulDiv(stdset.Size, GetDeviceCaps(mlngDC, LOGPIXELSY), 71)
        .lfWidth = 0
        .lfWeight = IIf(stdset.Bold = True, FW_BOLD, FW_NORMAL)
        .lfCharSet = stdset.Charset
        .lfUnderline = stdset.Underline
        .lfItalic = stdset.Italic
        .lfStrikeOut = stdset.Strikethrough
    End With

    mlngFont = CreateFontIndirect(T_Font)
    mlngOldFont = SelectObject(mlngDC, mlngFont)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ReleaseFontIndirect(ByVal objDraw As Object)
    Dim stdset As StdFont
    Dim strFontInfo As String
    
    strFontInfo = "����,9"
    
    Set stdset = New StdFont
    With stdset
        .Name = Split(strFontInfo, ",")(0)
        .Size = 9
        If UBound(Split(strFontInfo, ",")) > 0 Then .Size = Val(Split(strFontInfo, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strFontInfo, "��") > 0 Then .Bold = True
        If InStr(1, strFontInfo, "б") > 0 Then .Italic = True
    End With
    Set objDraw.Font = stdset
    Call SelectObject(mlngDC, mlngOldFont)
    Call DeleteObject(mlngFont)
End Sub

Private Sub GetTextRect(ByVal objDraw As Object, ByVal lngX As Long, ByVal lngY As Long, ByVal strInput As String, _
    Optional ByVal lngWidth As Long = 0, Optional bln���� As Boolean = True, Optional ByVal lngHeght As Long = 0)
    
    Dim lngInputW As Long, lng1H As Long
    Dim sngSize As Single
        
    T_LableRect.Left = lngX + objDraw.ScaleX(15, vbTwips, vbPixels) '��������߽绮���غ�
    
    If bln���� = True Then
        T_LableRect.Top = lngY - objDraw.ScaleY(objDraw.TextHeight("1") \ 2, vbTwips, vbPixels)
    Else
        T_LableRect.Top = lngY
    End If
    
    T_LableRect.Right = objDraw.ScaleX(objDraw.TextWidth(strInput) + 30, vbTwips, vbPixels) + T_LableRect.Left
    T_LableRect.Bottom = objDraw.ScaleY(objDraw.TextHeight("1") + 30, vbTwips, vbPixels) + T_LableRect.Top
    
    If lngWidth <> 0 And (lngWidth - objDraw.ScaleX(objDraw.TextWidth(strInput) + 30, vbTwips, vbPixels)) \ 2 > 0 Then
        '���ı���ʾ����ʾ��ȵ��м�����
        T_LableRect.Left = T_LableRect.Left + (lngWidth - objDraw.ScaleX(objDraw.TextWidth(strInput) + 30, vbTwips, vbPixels)) \ 2
        T_LableRect.Right = objDraw.ScaleX(objDraw.TextWidth(strInput) + 30, vbTwips, vbPixels) + T_LableRect.Left
    End If
    
    If lngHeght <> 0 And (lngHeght - objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels)) > 0 Then
        T_LableRect.Bottom = T_LableRect.Bottom + (lngHeght - objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels))
    End If
    
End Sub

Public Sub WaveDrawLine(ByVal lngDC As Long, ByVal lngSX As Long, ByVal lngSY As Long, ByVal lngDX As Long, ByVal lngDY As Long, _
    Optional ByVal lngType As Long = PS_SOLID, Optional ByVal intWidth As Integer = 1, Optional ByVal lngRGB As Long = 0, _
    Optional ByVal blnEndRow As Boolean = False)
    
    Dim X As Long
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim sngX As Single, sngY As Single
    On Error GoTo errHand
    '�����»��ʽ��л���
    lngPen = CreatePen(lngType, intWidth, lngRGB)
    lngOldPen = SelectObject(lngDC, lngPen)
    '��ͼ
    Call MoveToEx(lngDC, lngSX, lngSY, T_OldPoint)
    Call LineTo(lngDC, lngDX, lngDY)
    '���������»����¼�ͷ
    If blnEndRow Then
        If lngSY > lngDY Then '������ʧ�ܣ����ϼ�ͷ��
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDC, X, lngSY - sngY, T_OldPoint)
                Call LineTo(lngDC, lngSX, lngSY)
            Next X
        Else '�����³ɹ�(���¼�ͷ)
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDC, X, lngSY + sngY, T_OldPoint)
                Call LineTo(lngDC, lngSX, lngSY)
            Next X
        End If
    End If
    
    '��ԭ���ʲ�����
    Call SelectObject(lngDC, lngOldPen)
    Call DeleteObject(lngPen)
    lngPen = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetTabFrequency(ByVal intFrequency As Integer, ByVal intItemStyle As Integer) As String
'------------------------------------------------------
'����:���ݼ������ȷ����¼Ƶ��
'intFrequency��������
'intItemStyle����Ŀ��ʾ:��Ͳ�����Ŀ��¼Ƶ�����Ϊ2
'------------------------------------------------------
    Dim strFrequency As String
    Dim arrFrequency() As String
    Dim i As Integer
    
    If intItemStyle = 4 Then
        arrFrequency = Split("1|2", "|")
    Else
        arrFrequency = Split("1|2|3|4|6", "|")
    End If
    For i = 0 To UBound(arrFrequency)
        If intFrequency >= Val(arrFrequency(i)) And Not (intFrequency = 8 And Val(arrFrequency(i)) = 6) Then
            strFrequency = strFrequency & "|" & arrFrequency(i)
        End If
    Next i
    
    If Left(strFrequency, 1) = "|" Then strFrequency = Mid(strFrequency, 2)
    If strFrequency = "" Then strFrequency = "1"
    GetTabFrequency = strFrequency
End Function

Public Function GetPipWaveStyle(ByVal lngFileID As Long) As ADODB.Recordset
'-----------------------------------------------------------------------
'����:��ȡ��׼���µ�����ʽ(������װ)
'-----------------------------------------------------------------------
    Dim strSQL As String
    Dim rsSource As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngParentId As Long, lngId As Long
    Dim lngRow As Long, lngRowNO As Long
    Dim intFields As Integer
    Dim lngItemNO As Long '��Ŀ���
    On Error GoTo errHand
    strSQL = "SELECT Id, �ļ�id, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���," & vbNewLine & _
        "       Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
        " FROM �����ļ��ṹ" & vbNewLine & _
        " WHERE �ļ�id = 0"
    Call zlDatabase.OpenRecordset(rsSource, strSQL, "�����ļ��ṹ")
    '��ʼ���Ƽ�¼���ṹ��
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Type = 200 Then       '�����ʹ���Ϊ�ַ���
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
            End If
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '��ȡ��Ŀ��Ϣ
    strSQL = _
        " SELECT Decode(b.��Ŀ���, 3, Decode(b.��¼��, 2, 1, b.�������), b.�������) �������, b.��Ŀ���, Decode(b.��Ŀ���, 4, 'Ѫѹ', b.��¼��) ��Ŀ����, b.��λ," & vbNewLine & _
        "       b.��¼��," & vbNewLine & _
        "       Decode(b.��¼��," & vbNewLine & _
        "               2," & vbNewLine & _
        "               Decode(b.��Ŀ���," & vbNewLine & _
        "                      3," & vbNewLine & _
        "                      6," & vbNewLine & _
        "                      Decode(Decode(c.��Ŀ���, NULL, a.��Ŀ��ʾ, 4)," & vbNewLine & _
        "                             4," & vbNewLine & _
        "                             Decode(Sign(Nvl(b.��¼Ƶ��, 2) - 2), 1, 2, Nvl(b.��¼Ƶ��, 2))," & vbNewLine & _
        "                             Nvl(b.��¼Ƶ��, 2)))," & vbNewLine & _
        "               NULL) ��¼Ƶ��" & vbNewLine & _
        " FROM �����¼��Ŀ a, ���¼�¼��Ŀ b, ��������Ŀ c" & vbNewLine & _
        " WHERE a.��Ŀ��� = b.��Ŀ��� AND a.��Ŀ��� = c.��Ŀ���(+) AND NVL(a.Ӧ�÷�ʽ,0) <> 0 AND a.��Ŀ���� = 1 AND a.��Ŀ��� <> 5" & vbNewLine & _
        " ORDER BY Decode(b.��¼��, 2, 2, 1), Decode(b.��Ŀ���, 3, Decode(b.��¼��, 2, 1, b.�������), b.�������)"
    Call zlDatabase.OpenRecordset(rsSource, strSQL, "������Ŀ")
    '1:���µ��Ļ�����ʽ������
    lngParentId = 100
    With rsTemp
        '������
        .AddNew
        .Fields("ID").Value = lngParentId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = Null
        .Fields("�������").Value = 1: .Fields("��������").Value = 1: .Fields("��������").Value = "���µ��Ļ�����ʽ������"
        .Fields("�����ı�").Value = "��ʽ����"
        .Update
        '�Ӷ���
        lngId = 101
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 1: .Fields("��������").Value = 4: .Fields("��������").Value = "�����ı�"
        .Fields("�����ı�").Value = "���µ�": .Fields("Ҫ������").Value = "�����ı�"
        .Update
        lngId = 102
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 2: .Fields("��������").Value = 4: .Fields("��������").Value = "��������"
        .Fields("�����ı�").Value = "����,20,��": .Fields("Ҫ������").Value = "��������"
        .Update
        lngId = 103
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 3: .Fields("��������").Value = 4: .Fields("��������").Value = "����"
        .Fields("�����ı�").Value = 7: .Fields("Ҫ������").Value = "����"
        .Update
        lngId = 104
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 4: .Fields("��������").Value = 4: .Fields("��������").Value = "������"
        .Fields("�����ı�").Value = 6: .Fields("Ҫ������").Value = "������"
        .Update
        lngId = 105
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 5: .Fields("��������").Value = 4: .Fields("��������").Value = "��ʼʱ��"
        .Fields("�����ı�").Value = 4: .Fields("Ҫ������").Value = "��ʼʱ��"
        .Update
        lngId = 106
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 6: .Fields("��������").Value = 4: .Fields("��������").Value = "ʱ����"
        .Fields("�����ı�").Value = 4: .Fields("Ҫ������").Value = "ʱ����"
        .Update
        lngId = 107
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 7: .Fields("��������").Value = 4: .Fields("��������").Value = "һ����Ŀ�����߶�"
        .Fields("�����ı�").Value = 255: .Fields("Ҫ������").Value = "���߶�"
        .Update
        lngId = 108
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 8: .Fields("��������").Value = 4: .Fields("��������").Value = "һ����Ŀ����ͷ����"
        .Fields("�����ı�").Value = "��       ��@סԺ����@����������@ʱ       ��"
        .Fields("Ҫ������").Value = "��ͷ����"
        .Update
        lngId = 109
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 9: .Fields("��������").Value = 4: .Fields("��������").Value = "�̶������ܿ��(�)"
        .Fields("�����ı�").Value = 1350: .Fields("Ҫ������").Value = "�̶ȿ��"
        .Update
        lngId = 110
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 10: .Fields("��������").Value = 4: .Fields("��������").Value = "��ͼ�������߱���п�(�)"
        .Fields("�����ı�").Value = 225: .Fields("Ҫ������").Value = "�����п�"
        .Update
        lngId = 111
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 11: .Fields("��������").Value = 4: .Fields("��������").Value = "��ͼ�������߱���и�(�)"
        .Fields("�����ı�").Value = 90: .Fields("Ҫ������").Value = "�����и�"
        .Update
        lngId = 112
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 12: .Fields("��������").Value = 4: .Fields("��������").Value = "���߱����ӿ�����(����Զ�������)"
        .Fields("�����ı�").Value = 10: .Fields("Ҫ������").Value = "���߿���"
        .Update
        lngId = 113
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 13: .Fields("��������").Value = 4: .Fields("��������").Value = "������Ŀ�����߶�"
        .Fields("�����ı�").Value = 255: .Fields("Ҫ������").Value = "���߶�1"
        .Update
        lngId = 114
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 14: .Fields("��������").Value = 4: .Fields("��������").Value = "������Ŀ�������ӵĿ�����"
        .Fields("�����ı�").Value = 0: .Fields("Ҫ������").Value = "������"
        .Update
    End With
    '2:���µ�������Ŀ����
    lngParentId = 200
    With rsTemp
        '������
        .AddNew
        .Fields("ID").Value = lngParentId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = Null
        .Fields("�������").Value = 2: .Fields("��������").Value = 1: .Fields("��������").Value = "���µ�������Ŀ����"
        .Fields("�����ı�").Value = "������Ŀ����"
        .Update
        lngRowNO = 1
        rsSource.Filter = "��¼��=1 OR ��¼��=3"
        rsSource.Sort = "��¼��,�������"
        Do While Not rsSource.EOF
            '�Ӷ���
            If Val(NVL(rsSource!��Ŀ���, 0)) <> 0 Then
                lngId = 200 + lngRowNO
                .AddNew
                .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
                .Fields("�������").Value = lngRowNO: .Fields("��������").Value = 4
                .Fields("��������").Value = NVL(rsSource!��Ŀ����)
                .Fields("�����ı�").Value = NVL(rsSource!��Ŀ���)
                .Fields("Ҫ������").Value = NVL(rsSource!��Ŀ����)
                .Update
                lngRowNO = lngRowNO + 1
            End If
        rsSource.MoveNext
        Loop
    End With
    '2:���µ������Ŀ����
    lngParentId = 300
    With rsTemp
        '������
        .AddNew
        .Fields("ID").Value = lngParentId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = Null
        .Fields("�������").Value = 3: .Fields("��������").Value = 1: .Fields("��������").Value = "���µ������Ŀ����"
        .Fields("�����ı�").Value = "�����Ŀ����"
        .Update
        lngRowNO = 1
        rsSource.Filter = "��¼��=2"
        rsSource.Sort = "�������"
        Do While Not rsSource.EOF
            If Val(NVL(rsSource!��Ŀ���, 0)) <> 0 Then
                '�Ӷ���
                lngId = 300 + lngRowNO
                .AddNew
                .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = lngFileID: .Fields("��ID").Value = lngParentId
                .Fields("�������").Value = lngRowNO: .Fields("��������") = 4
                .Fields("��������").Value = NVL(rsSource!��Ŀ����)
                .Fields("�����ı�").Value = NVL(rsSource!��Ŀ���)
                .Fields("Ҫ������").Value = NVL(rsSource!��Ŀ����)
                .Fields("Ҫ�ر�ʾ").Value = NVL(rsSource!��¼Ƶ��)
                .Update
                lngRowNO = lngRowNO + 1
            End If
            rsSource.MoveNext
        Loop
    End With
    rsTemp.Filter = ""
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    rsTemp.Sort = "ID"
    
    Set GetPipWaveStyle = rsTemp
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


