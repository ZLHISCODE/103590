Attribute VB_Name = "mdlPrint"
Option Explicit
'----------------------------------------------------------------------------------------------
'˵��:
'1.��ģ�����������ӡ���ܺ���,Ϊ���סԺ��������
'2.��ģ����Ҫ����һЩ������������(���ȡ����,API��ͼ),���뱣֤�⼸������������Ӻ���������ģ���С�
'3.��ģ�黹�������±������¼���Ĵ�ӡ����
'----------------------------------------------------------------------------------------------
Public Const OFFSET_LEFT = 20
Public Const OFFSET_TOP = 20
Public Const OFFSET_RIGHT = 20
Public Const OFFSET_BOTTOM = 20

Private Const MAXROWS = 46
Private Const ROWHEIGHT = 35
Private Const OPDAYS = 10                           '������������
Private Const HOUR_STEP_Twips = 205                 '������������Сʱ֮��Ŀ�� �������±�

Private Const INTSTEPTwip = 90  '��������5�������Ŀ�� ��������
Private Const STRING_WAY As String = "��"
Private msngScale As Single
Private mstrSQL As String
Private mrsTmp As ADODB.Recordset
Private mblnMoved As Boolean
Private mbln�������� As Boolean
Private mblnӤ�����µ���ʾ��Ժ As Boolean
Private mstrChar(2) As String                       '����Ϊ����,Ҹ��,����
Private mstrBreath As String                        '����
Private mstrPulse As String                         '����
Private mint����Ӧ�� As Integer
Private mstr���ʷ��� As String
Private mlngFirstWidth As Long
Private mintOpDays As Integer
Private mblnStopFlag As Boolean
Private mbyt����() As Byte

Private Enum COLOR
    ��ɫ = 0
    ���ɫ = &H404040
    ��ɫ = &HE0E0E0
    ��ɫ = 200
End Enum

Private Type GRAPHPOINT
    X As Single
    Y As Single
    ���� As String
    ��ɫ As Long
    ��־ As Byte
End Type

Private Type BODYFLAG
    ��Ժ As Byte
    ��� As Byte
    ת�� As Byte
    ���� As Byte
    ���� As Byte
    ��Ժ As Byte
    ���� As Byte
    ���� As Byte
End Type

Private mBodyFlag As BODYFLAG

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWHEELSCROLLLINES = 104
Public WHEEL_SCROLL_LINES As Long
Global glngPrevWndProc As Long
Public Const SW_RESTORE = 9
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

'Window�汾����
Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Const DC_PAPERNAMES = 16 'ֽ������(ÿ64�ַ�Ϊһ��,��Chr(0)����)
Public Const DC_PAPERS = 2 'ֽ�ű��(Array or Word)
Public Const DC_BINNAMES = 12 '��ֽ��ʽ(ÿ24�ַ�Ϊһ��,��Chr(0)����)
Public Const DC_BINS = 6 '��ֽ���(Array or Word)

'��ӡֽ�ų���(256=�Զ���)
Public Const PageSize1 = "�ż㣬 8 1/2 x 11 Ӣ��"
Public Const PageSize2 = "+A611 С���ż㣬 8 1/2 x 11 Ӣ��"
Public Const PageSize3 = "С�ͱ��� 11 x 17 Ӣ��"
Public Const PageSize4 = "�����ʣ� 17 x 11 Ӣ��"
Public Const PageSize5 = "�����ļ��� 8 1/2 x 14 Ӣ��"
Public Const PageSize6 = "�����飬5 1/2 x 8 1/2 Ӣ��"
Public Const PageSize7 = "�����ļ���7 1/2 x 10 1/2 Ӣ��"
Public Const PageSize8 = "A3, 297 x 420 ����"
Public Const PageSize9 = "A4, 210 x 297 ����"
Public Const PageSize10 = "A4С�ţ� 210 x 297 ����"
Public Const PageSize11 = "A5, 148 x 210 ����"
Public Const PageSize12 = "B4, 250 x 354 ����"
Public Const PageSize13 = "B5, 182 x 257 ����"
Public Const PageSize14 = "�Կ����� 8 1/2 x 13 Ӣ��"
Public Const PageSize15 = "�Ŀ����� 215 x 275 ����"
Public Const PageSize16 = "10 x 14 Ӣ��"
Public Const PageSize17 = "11 x 17 Ӣ��"
Public Const PageSize18 = "������8 1/2 x 11 Ӣ��"
Public Const PageSize19 = "#9 �ŷ⣬ 3 7/8 x 8 7/8 Ӣ��"
Public Const PageSize20 = "#10 �ŷ⣬ 4 1/8 x 9 1/2 Ӣ��"
Public Const PageSize21 = "#11 �ŷ⣬ 4 1/2 x 10 3/8 Ӣ��"
Public Const PageSize22 = "#12 �ŷ⣬ 4 1/2 x 11 Ӣ��"
Public Const PageSize23 = "#14 �ŷ⣬ 5 x 11 1/2 Ӣ��"
Public Const PageSize24 = "C �ߴ繤����"
Public Const PageSize25 = "D �ߴ繤����"
Public Const PageSize26 = "E �ߴ繤����"
Public Const PageSize27 = "DL ���ŷ⣬ 110 x 220 ����"
Public Const PageSize28 = "C5 ���ŷ⣬ 162 x 229 ����"
Public Const PageSize29 = "C3 ���ŷ⣬ 324 x 458 ����"
Public Const PageSize30 = "C4 ���ŷ⣬ 229 x 324 ����"
Public Const PageSize31 = "C6 ���ŷ⣬ 114 x 162 ����"
Public Const PageSize32 = "C65 ���ŷ⣬114 x 229 ����"
Public Const PageSize33 = "B4 ���ŷ⣬ 250 x 353 ����"
Public Const PageSize34 = "B5 ���ŷ⣬176 x 250 ����"
Public Const PageSize35 = "B6 ���ŷ⣬ 176 x 125 ����"
Public Const PageSize36 = "�ŷ⣬ 110 x 230 ����"
Public Const PageSize37 = "�ŷ������ 3 7/8 x 7 1/2 Ӣ��"
Public Const PageSize38 = "�ŷ⣬ 3 5/8 x 6 1/2 Ӣ��"
Public Const PageSize39 = "U.S. ��׼��д���� 14 7/8 x 11 Ӣ��"
Public Const PageSize40 = "�¹���׼��д���� 8 1/2 x 12 Ӣ��"
Public Const PageSize41 = "�¹����ɸ�д���� 8 1/2 x 13 Ӣ��"

'ֽ�Ŵ�ӡ�߽����================================================================
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'��ͬ��ӡ���Ĵ�ӡ��Ԫ���Ȳ�ͬ

'Public Const PHYSICALHEIGHT = 111  'Physical Height in device units

'Public Const PHYSICALOFFSETY = 113 'Physical Printable Area y margin
Public Const LOGPIXELSX = 88 'Number of pixels per logical inch along the screen width
Public Const LOGPIXELSY = 90
Public Const SCALINGFACTORX = 114  'Scaling factor x
Public Const SCALINGFACTORY = 115  'Scaling factor y
Public Const DRIVERVERSION = 0     'Device driver version

'WinNT�Զ���ֽ�ſ���================================================================
Public Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal Level As Long, ByRef pForm As Any, ByVal cbBuf As Long, ByRef pcbNeeded As Long, ByRef pcReturned As Long) As Long
Public Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByRef lpString2 As Long) As Long

' Optional functions not used in this sample, but may be useful.
Public Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long

' Constants for DEVMODE
Public Const CCHFORMNAME = 32
Public Const CCHDEVICENAME = 32
Public Const DM_FORMNAME As Long = &H10000
Public Const DM_ORIENTATION = &H1&

' Constants for PRINTER_DEFAULTS.DesiredAccess
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' Constants for DocumentProperties() call
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY

' Custom constants for this sample's SelectForm function
Public Const FORM_NOT_SELECTED = 0
Public Const FORM_SELECTED = 1
Public Const FORM_ADDED = 2


Public Type RECTL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type SIZEL
    cx As Long
    cy As Long
End Type

Public Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    Sacl As Long  ' ACL
    Dacl As Long  ' ACL
End Type

' The two definitions for FORM_INFO_1 make the coding easier.
Public Type FORM_INFO_1
    flags As Long
    pName As Long   ' String
    Size As SIZEL
    ImageableArea As RECTL
End Type

Public Type sFORM_INFO_1
    flags As Long
    pName As String
    Size As SIZEL
    ImageableArea As RECTL
End Type

Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public Type PRINTER_DEFAULTS
    pDatatype As String
    pDevMode As Long    ' DEVMODE
    DesiredAccess As Long
End Type

Public Type PRINTER_INFO_2
    pServerName As String
    pPrinterName As String
    pShareName As String
    pPortName As String
    pDriverName As String
    pComment As String
    pLocation As String
    pDevMode As DEVMODE
    pSepFile As String
    pPrintProcessor As String
    pDatatype As String
    pParameters As String
    pSecurityDescriptor As SECURITY_DESCRIPTOR
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    cJobs As Long
    AveragePPM As Long
End Type

Public gblnOK As Boolean


    
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    
Private Const PS_SOLID = 0

Public gblnPrinted As Boolean           '�Ƿ��ӡ�����µ�
  
Public Sub DrawDcLine(hDC As Long, startpx As Long, startpy As Long, endpx As Long, endpy As Long)
        Dim old     As Long
        Dim p     As Long
        Dim a     As POINTAPI
          
        p = CreatePen(PS_SOLID, 3, vbRed)
        old = SelectObject(hDC, p)
        MoveToEx hDC, startpx, startpy, a
        LineTo hDC, endpx, endpy
        SelectObject hDC, old
        DeleteObject p
End Sub
  
'===================================================================================

Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Public Function GetPaperName(intSize As Integer) As String
    '���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡֽ������
    '���أ� ֽ������
    If intSize = 256 Then
        GetPaperName = "�û��Զ��� ..."
    ElseIf intSize >= 1 And intSize <= 41 Then
        GetPaperName = Switch( _
        intSize = 1, PageSize1, intSize = 2, PageSize2, intSize = 3, PageSize3, intSize = 4, PageSize4, intSize = 5, PageSize5, _
            intSize = 6, PageSize6, intSize = 7, PageSize7, intSize = 8, PageSize8, intSize = 9, PageSize9, intSize = 10, PageSize10, _
            intSize = 11, PageSize11, intSize = 12, PageSize12, intSize = 13, PageSize13, intSize = 14, PageSize14, intSize = 15, PageSize15, _
            intSize = 16, PageSize16, intSize = 17, PageSize17, intSize = 18, PageSize18, intSize = 19, PageSize19, intSize = 20, PageSize20, _
            intSize = 21, PageSize21, intSize = 22, PageSize22, intSize = 23, PageSize23, intSize = 24, PageSize24, intSize = 25, PageSize25, _
            intSize = 26, PageSize26, intSize = 27, PageSize27, intSize = 28, PageSize28, intSize = 29, PageSize29, intSize = 30, PageSize30, _
            intSize = 31, PageSize31, intSize = 32, PageSize32, intSize = 33, PageSize33, intSize = 34, PageSize34, intSize = 35, PageSize35, _
            intSize = 36, PageSize36, intSize = 37, PageSize37, intSize = 38, PageSize38, intSize = 39, PageSize39, intSize = 40, PageSize40, _
            intSize = 41, PageSize41)
    Else
        GetPaperName = "���ɲ��ֽ�� ..."
    End If
End Function

Public Function DrawCell(Dev As Object, ByVal Data As Variant, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, _
                        Optional ByVal TW As Long, Optional ByVal TH As Long, Optional BorderColor As Long, _
                        Optional ForeColor As Long, Optional BackColor As Long = &HFFFFFF, Optional ByVal Font As StdFont, _
                        Optional Border As String = "1111", Optional HAlign As Byte, Optional VAlign As Byte = 1, Optional Warp As Boolean, _
                        Optional Ratio As Single = 1, Optional ByVal sngScale As Single = 1) As Boolean
                        
    '���ܣ���ָ���豸�ϰ�ָ����ʽ��������ֻ�ͼ��
    '������
    '   Dev=����豸,ΪPrinter��PictureBox����
    '   Data=�������,Ϊ����(x)���ַ���("xxx")��ͼ��(stdPicture)���ַ���������vbCrLf,��Data����Ϊ������ʱ,��ʾ�������
    '   TW,TH=������޶���Χ,���������Χ���Զ�ȡ������С,Ϊ0ʱ��Ч
    '   Border=�߿���,��������,"1111"��ʾȫ��
    '   Align=���ֶ���,0=��,1=��,2=��,��ˮƽ���뼰��ֱ����
    '   Warp=���������Ϊ�ַ���ʱ,��ʾ�Ƿ��Զ����С����Զ�����ʱ,�����ݲ������
    '   Ratio=�������,������,���궼��Ӱ��,ȱʡΪ1(100%)
    '˵����1.��ʹ�øú���֮ǰ,Ӧ��û�иı��豸����ͼ��ʼֵ
    '      2.�����λ���λ���ڱ��������Χ�����Ͻ�
    
    Dim i As Long, Text As String, arrText() As String
    Dim LINE_W As Integer, blnW As Boolean, blnH As Boolean
    Dim sglFontSize As Single
    
    On Error GoTo errH
    
    sglFontSize = Dev.Font.Size
    
    DrawCell = True
    
    '��Χ�޶�
    If TW > 0 Then
        If X > TW Then Exit Function
        If X + W > TW Then W = TW - X
    End If
    If TH > 0 Then
        If Y > TH Then Exit Function
        If Y + H > TH Then H = TH - Y
    End If
    
    If TypeName(Data) = "Integer" Then
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        If Val(Data) < 0 Then
            Dev.Line (X, Y)-(X + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), Y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, B '����
        Else
            Dev.Line (X, Y)-(X + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), Y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, BF 'ʵ�ľ���(����)
        End If
    ElseIf TypeName(Data) = "String" Then
        '����
        If Font Is Nothing Then
            Set Font = New StdFont
            Font.Name = "����"
            Font.Size = 9
        End If

        'ǧ��Ҫ��Set Dev.Font=Font,��֪Ϊ��,�õ���ByVal
        Dev.Font.Name = Font.Name
        Dev.Font.Size = Font.Size * sngScale
        Dev.Font.Bold = Font.Bold
        Dev.Font.Underline = Font.Underline
        Dev.Font.Italic = Font.Italic
        
        '�����ź���������������,�ж�ʱ��ԭʼ��СΪ׼
        If H >= Dev.TextHeight(Replace(Data, vbCrLf, "")) Then blnH = True          '�߶��Ƿ���(�ӻس�����һ�и߶�)
        If W >= Dev.TextWidth(Data) Then blnW = True And InStr(Data, vbCrLf) = 0    '����Ƿ���(�ӻس���Ϊ������,�Ա����)
        '����
        LINE_W = 30 * Ratio '���߼�����(���ʱ��,�ж�ʱ����)
        X = -Int(-X * Ratio): Y = -Int(-Y * Ratio)
        W = -Int(-W * Ratio): H = -Int(-H * Ratio)
        Dev.Font.Size = Font.Size * Ratio
        '�������
        Dev.Line (X, Y)-(X + W, Y + H), BackColor, BF
        Dev.ForeColor = ForeColor
        '�������(�߿�֮���ٸ�һ��)
        '�����߶ȷ�Χ�����
        If blnH Then
            If blnW Then
                Select Case HAlign
                Case 0
                    Dev.CurrentX = X + LINE_W
                Case 1
                    Dev.CurrentX = X + (W - Dev.TextWidth(Data)) / 2
                Case 2
                    Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(Data)
                End Select
                Select Case VAlign
                Case 0
                    Dev.CurrentY = Y + LINE_W
                Case 1
                    Dev.CurrentY = Y + (H - Dev.TextHeight(Data)) / 2 + LINE_W / 2
                Case 2
                    Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Data)
                End Select
                Dev.FontTransparent = True
                Dev.Print Data
            Else
                If Not Warp Then
                    '���Զ�����ʱ�����ֲ����
                    
'                    '����ʱ���Զ���С����
'                    If Dev.TextWidth(Data) > W Then
'                        '�����ܵĿ�ȼ��������С
'                        Dev.Font.Size = Dev.Font.Size * (1 - (Dev.TextWidth(Data) - W) / W - 0.05)
'                    End If
'                    Text = Data
                    
                    For i = 1 To Len(Data)
                        If Dev.TextWidth(Text & Mid(Data, i, 1)) > W Then Exit For
                        Text = Text & Mid(Data, i, 1)
                    Next

                    Select Case HAlign
                    Case 0
                        Dev.CurrentX = X + LINE_W
                    Case 1
                        Dev.CurrentX = X + (W - Dev.TextWidth(Text)) / 2
                    Case 2
                        Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(Text)
                    End Select
                    Select Case VAlign
                    Case 0
                        Dev.CurrentY = Y + LINE_W
                    Case 1
                        Dev.CurrentY = Y + (H - Dev.TextHeight(Text)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Text)
                    End Select
                    Dev.FontTransparent = True
                    '�����ȡ����
                    Dev.Print Text
                Else
                    '������ֳɶ���(�ڿ�߷�Χ��)
                    ReDim arrText(0) '�ڴ�,��һ�в����ܳ���
                    Data = Replace(Data, vbCrLf, vbCr)
                    Data = Replace(Data, vbLf, vbCr)
                    For i = 1 To Len(Data)
                        If Mid(Data, i, 1) = vbCr Then
                            '���г������˳�,���߲��ݲ����
                            If Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        ElseIf Dev.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) > W Then
                            '���г������˳�,���߲��ݲ����
                            If Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        End If
                        '�п���һ��һ���ַ���ȶ�����
                        If Dev.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) <= W And Mid(Data, i, 1) <> vbCr Then
                            arrText(UBound(arrText)) = arrText(UBound(arrText)) & Mid(Data, i, 1)
                        End If
                    Next
                    
                    '�����ʼ����
                    Select Case VAlign
                    Case 0
                        Dev.CurrentY = Y + LINE_W
                    Case 1
                        Dev.CurrentY = Y + (H - Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = Y + H - LINE_W - Dev.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)
                    End Select
                    
                    '�������
                    For i = 0 To UBound(arrText)
                        Select Case HAlign
                        Case 0
                            Dev.CurrentX = X + LINE_W
                        Case 1
                            Dev.CurrentX = X + (W - Dev.TextWidth(arrText(i))) / 2
                        Case 2
                            Dev.CurrentX = X + W - LINE_W - Dev.TextWidth(arrText(i))
                        End Select
                        Dev.FontTransparent = True
                        Dev.Print arrText(i)
                    Next
                End If
            End If
        End If
    ElseIf Not Data Is Nothing Then
        LINE_W = 30 * Ratio '���߼�����(���ʱ��,�ж�ʱ����)
        X = X * Ratio: Y = Y * Ratio: W = W * Ratio: H = H * Ratio
        
        'ͼ��(�߿�֮��)
        Dev.PaintPicture Data, X + 15, Y + 15, W - LINE_W, H - LINE_W
    End If
    If TypeName(Data) <> "Integer" Then
        '�����߿�,�ϣ��£�����
        If Mid(Border, 1, 1) Then Dev.Line (X, Y)-(X + W, Y), BorderColor
        If Mid(Border, 2, 1) Then Dev.Line (X, Y + H)-(X + W, Y + H), BorderColor
        If Mid(Border, 3, 1) Then Dev.Line (X, Y)-(X, Y + H), BorderColor
        If Mid(Border, 4, 1) Then Dev.Line (X + W, Y)-(X + W, Y + H), BorderColor
    End If
    
    Dev.Font.Size = sglFontSize
    
    Exit Function
errH:
    DrawCell = False
End Function

Public Function InitPrint(objParent As Object) As Boolean
    '���ܣ�����ע���frmparent.mobjreport���ݳ�ʼ����ӡ������(����->������->��ǰ)
    '���أ�����޴�ӡ����ֽ�Ų���,��ʧ��
    Dim i As Integer, strPName As String
    Dim strPrinter As String  '��ӡ��
    Dim intPage As Integer  'ֽ��
    Dim lngWidth As Long  '�Զ���ֽ�ſ��
    Dim lngHeight As Long  '�Զ���ֽ�Ÿ߶�
    Dim intOrient As Byte  'ֽ��
    Dim intBin As Integer  '��ֽ��ʽ
    If Not ExistsPrinter Then Exit Function
    
    '��ʼ����ӡ����
    
    strPrinter = Trim(zlDatabase.GetPara("���µ���ӡ��", glngSys, 1255, Printer.DeviceName))
    intPage = Val(zlDatabase.GetPara("���µ�ֽ��", glngSys, 1255, Printer.PaperSize))
    lngWidth = Val(zlDatabase.GetPara("���µ����", glngSys, 1255, Printer.Width))
    lngHeight = Val(zlDatabase.GetPara("���µ��߶�", glngSys, 1255, Printer.Height))
    intOrient = Val(zlDatabase.GetPara("���µ�ֽ��", glngSys, 1255, Printer.Orientation))
    intBin = Val(zlDatabase.GetPara("���µ���ֽ", glngSys, 1255, Printer.PaperBin))
    
    '��ӡ��
    If Printer.DeviceName <> strPName Then
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = strPrinter Then Set Printer = Printers(i): Exit For
        Next
    End If
    On Error Resume Next
    'ֽ��
    If intPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    Else
        Printer.PaperSize = intPage
    End If
    'ֽ��
    'ֽ��ֵ��,ֽ�ſ��ֵ����,ֽ��ԭΪ1
    Printer.Orientation = intOrient
    '��ֽ
    Printer.PaperBin = intBin
    '����
    Printer.Copies = 1
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    'WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And intPage = 256 Then
        If AddCustomPaper(objParent.hWnd, lngWidth / 56.7, lngHeight / 56.7) = FORM_NOT_SELECTED Then Exit Function
    End If
    InitPrint = True
End Function

Public Function IsWindowsNT() As Boolean
    '���ܣ��Ƿ�WindowNT����ϵͳ
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Public Function IsWindows95() As Boolean
    '���ܣ��Ƿ�Window95����ϵͳ
    Const dwMask95 = &H1&
    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function

Private Function GetWinPlatform() As Long
    Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

Public Function GetFormName(ByVal PrinterHandle As Long, FormSize As SIZEL, FormName As String) As Integer
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1           ' Working FI1 array
    Dim Temp() As Byte                  ' Temp FI1 array
    Dim FormIndex As Integer
    Dim BytesNeeded As Long
    Dim RetVal As Long
    
    FormName = vbNullString
    FormIndex = 0
    ReDim aFI1(1)
    ' First call retrieves the BytesNeeded.
    RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
    ReDim Temp(BytesNeeded)
    ReDim aFI1(BytesNeeded / Len(FI1))
    ' Second call actually enumerates the supported forms.
    RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
    Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
    For i = 0 To NumForms - 1
        With aFI1(i)
            If .Size.cx = FormSize.cx And .Size.cy = FormSize.cy Then
                ' Found the desired form
                FormName = PtrCtoVbString(.pName)
                FormIndex = i + 1
                Exit For
            End If
        End With
    Next i
    GetFormName = FormIndex  ' Returns non-zero when form is found.
End Function

Public Function AddCustomPaper(ByVal lngHwnd As Long, lngWidth As Long, lngHeight As Long) As Integer
    '���ܣ�����һ��NT��ʹ�õ��Զ���ֽ��
    '����������=mm(����)
    Dim lngSize As Long ' Size of DEVMODE
    Dim vDevMode As DEVMODE
    Dim arrDevMode() As Byte ' Working DEVMODE
    
    Dim lngHandle As Long 'Handle to printer
    Dim lngPrtDC As Long ' Handle to Printer DC
    Dim strPrtName As String
    
    Dim vFormSize As SIZEL
    
    strPrtName = Printer.DeviceName
    lngPrtDC = Printer.hDC
    
    If OpenPrinter(strPrtName, lngHandle, 0&) Then '��ȡ��ӡ�����
        ' Retrieve the size of the DEVMODE.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, 0&, 0&, 0&)
        ' Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
        
        ' Fill the DEVMODE From the printer.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        ' Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        ' If FormName is "zlBillPaper", we must make sure it exists
        ' before using it. Otherwise, it came From our EnumForms list,
        ' And we do not need to check first. Note that we could have
        ' passed in a Flag instead of checking for a literal name.
        
        ' Use form "zlBillPaper", adding it if necessary.
        ' Set the desired size of the form needed.
        ' Given in thousandths of millimeters
        vFormSize.cx = lngWidth * 1000 ' width
        vFormSize.cy = lngHeight * 1000 ' height
        
        If GetFormName(lngHandle, vFormSize, "zlBillPaper") = 0 Then
            'Form not found - Either of the next 2 lines will work.
            'FormName = AddNewForm(lngHandle, vFormSize, "zlBillPaper")
            AddNewForm lngHandle, vFormSize, "zlBillPaper"
            If GetFormName(lngHandle, vFormSize, "zlBillPaper") = 0 Then
                Call ClosePrinter(lngHandle)
                AddCustomPaper = FORM_NOT_SELECTED   ' Selection Failed!
                Exit Function
            Else
                AddCustomPaper = FORM_ADDED  ' Form Added, Selection succeeded!
            End If
        End If
        ' Change the appropriate member in the DevMode.
        ' In this case, you want to change the form name.
        vDevMode.dmFormName = "zlBillPaper" & Chr(0)  ' Must be NULL terminated!
        ' Set the dmFields bit flag to indicate what you are changing.
        vDevMode.dmFields = DM_FORMNAME
        
        ' Copy your changes back, then update DEVMODE.
        Call CopyMemory(arrDevMode(1), vDevMode, Len(vDevMode))
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))   ' Reset the DEVMODE for the DC.
        ' Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
        ' Selection Succeeded! But was Form Added?
        If AddCustomPaper <> FORM_ADDED Then AddCustomPaper = FORM_SELECTED
    Else
        AddCustomPaper = FORM_NOT_SELECTED   ' Selection Failed!
    End If
End Function

Public Function DelCustomPaper() As Boolean
    '���ܣ�ɾ���ղŴ������Զ���ֽ��
    Dim lngHandle As Long
    Dim strName As String
    
    strName = Printer.DeviceName
    If OpenPrinter(strName, lngHandle, 0&) Then
        DelCustomPaper = (DeleteForm(lngHandle, "zlBillPaper" & Chr(0)) <> 0)
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function PtrCtoVbString(ByVal Add As Long) As String
    Dim sTemp As String * 512, X As Long
    
    X = lstrcpy(sTemp, ByVal Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
        PtrCtoVbString = ""
    Else
        PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Public Function AddNewForm(lngPrtHandle As Long, vFormSize As SIZEL, strFormName As String) As String
    Dim FI1 As sFORM_INFO_1
    Dim aFI1() As Byte
    Dim RetVal As Long
    
    With FI1
        .flags = 0
        .pName = strFormName
        With .Size
            .cx = vFormSize.cx
            .cy = vFormSize.cy
        End With
        With .ImageableArea
            .Left = 0
            .Top = 0
            .Right = FI1.Size.cx
            .Bottom = FI1.Size.cy
        End With
    End With
    ReDim aFI1(Len(FI1))
    Call CopyMemory(aFI1(0), FI1, Len(FI1))
    RetVal = AddForm(lngPrtHandle, 1, aFI1(0))
    If RetVal = 0 Then
        If Err.LastDllError = 5 Then
            MsgBox "��û��Ȩ�����ô�ӡ��""" & Printer.DeviceName & """Ϊ�Զ���ߴ磬��ӡ������ܻ᲻������", vbExclamation, App.Title
        Else
            MsgBox "���ô�ӡ��ֽ��ʱ�������󣬱�ţ� " & Err.LastDllError, vbExclamation, App.Title
        End If
        AddNewForm = ""
    Else
        AddNewForm = FI1.pName
    End If
End Function

Public Sub ShowFlash(Optional strInfo As String, Optional sngPer As Single, Optional frmParent As Object)
    '���ܣ���ʾ�����صȴ�����ȴ���(strInfo)
    '����:strInfo=������ʾ��Ϣ
    '     sngPer=����
    Static blnShow As Boolean
    If sngPer > 1 Then sngPer = 1

    If strInfo = "" Then
        Unload frmFlash
        blnShow = False
    Else
        If Not blnShow Then
            On Error Resume Next
            frmFlash.lbl.Top = frmFlash.lbl.Top - frmFlash.lbl.Height / 2
            frmFlash.lblPer.Top = frmFlash.lbl.Top
            frmFlash.lbl.Caption = strInfo
            frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)

            If sngPer > 0 Then
                frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
            Else
                frmFlash.lblPer.Caption = ""
            End If

            If frmParent Is Nothing Then
                SetWindowPos frmFlash.hWnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                ShowWindow frmFlash.hWnd, 5
            Else
                Err.Clear
                frmFlash.Show , frmParent
                If Err.Number <> 0 Then
                    Err.Clear
                    SetWindowPos frmFlash.hWnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                    ShowWindow frmFlash.hWnd, 5
                End If
            End If

            frmFlash.Refresh
            blnShow = True
        Else
            frmFlash.lbl.Caption = strInfo
            If sngPer >= 0 Then
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If sngPer > 0 Then
                    frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                Else
                    frmFlash.lblPer.Caption = ""
                End If
            End If
            frmFlash.Refresh
        End If
    End If
End Sub

Private Function NewPrintPage(objOut As Object, intPage As Long, Optional blnNewPage As Boolean = True) As Object
    '���ܣ���ӡ��Ԥ��һҳ����ʱ�Ե�ǰҳ����������,��������ҳ
    '������blnNewPage=ΪFalseʱ����ӡҳ�ŵ�,һ���ӡ��������������,��˲����������
    '���أ���ҳ����,����Ϊ��ӡ����PictureBox
    On Error GoTo errH
    Dim objDraw As Object, blnPrint As Boolean
    Dim lngWidth As Long, lngHeight As Long, lngOldY As Long
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim strFontName As String, lngFontSize As Long, blnFontBold As Boolean
    Dim blnFontItalic As Boolean, lngFontColor As Long
    
    
    blnPrint = TypeName(objOut) = "Printer"
    '�߽���Ϣ(Twip)
    
    lngLeft = Val(zlDatabase.GetPara("���µ���߾�", glngSys, 1255, OFFSET_LEFT)) * 56.7
    lngRight = Val(zlDatabase.GetPara("���µ��ұ߾�", glngSys, 1255, OFFSET_RIGHT)) * 56.7
    lngTop = Val(zlDatabase.GetPara("���µ��ϱ߾�", glngSys, 1255, OFFSET_TOP)) * 56.7
    lngBottom = Val(zlDatabase.GetPara("���µ��±߾�", glngSys, 1255, OFFSET_BOTTOM)) * 56.7
    
    
    lngWidth = Printer.Width: lngHeight = Printer.Height
    
    'һҳ���������Ĵ���
    If Not blnPrint Then
        Set objDraw = objOut.picPage(objOut.picPage.UBound)
    Else
        Set objDraw = Printer
    End If
    strFontName = objDraw.Font.Name
    lngFontSize = objDraw.Font.Size
    blnFontBold = objDraw.Font.Bold
    blnFontItalic = objDraw.Font.Italic
    lngFontColor = objDraw.ForeColor
    lngOldY = objDraw.CurrentY
    
    '��ӡҳ��(0Ϊ����ӡ)
    If intPage <> 0 Then
        objDraw.ForeColor = 0
        objDraw.Font.Name = "����"
        objDraw.Font.Size = 9
        objDraw.Font.Bold = False
        objDraw.CurrentY = lngHeight - lngBottom - objDraw.TextHeight("��")
        objDraw.CurrentX = lngLeft + (lngWidth - lngLeft - lngRight) * (3 / 4)
        objDraw.FontTransparent = True
        objDraw.Print "���� " & intPage & " ҳ��"
    End If
    
    If Not blnPrint Then
        'Ԥ����ӡ����
        objDraw.DrawStyle = 2
        objDraw.Line (0, lngTop)-(lngWidth, lngTop), &H808080
        objDraw.Line (0, lngHeight - lngBottom)-(lngWidth, lngHeight - lngBottom), &H808080
        objDraw.Line (lngLeft, 0)-(lngLeft, lngHeight), &H808080
        objDraw.Line (lngWidth - lngRight, 0)-(lngWidth - lngRight, lngHeight), &H808080
        objDraw.DrawStyle = 0
    End If
    
    '������ҳ
    If blnNewPage Then
        intPage = intPage + 1
        If blnPrint Then
            Printer.NewPage
            Set objDraw = Printer
        Else
            Load objOut.picPage(objOut.picPage.UBound + 1)
            Set objDraw = objOut.picPage(objOut.picPage.UBound)
            objDraw.Width = Printer.Width
            objDraw.Height = Printer.Height
            objDraw.ZOrder
            objDraw.Cls
            objDraw.AutoRedraw = True
        End If
        objDraw.Font.Name = strFontName
        objDraw.Font.Size = lngFontSize
        objDraw.Font.Bold = blnFontBold
        objDraw.Font.Italic = blnFontItalic
        objDraw.ForeColor = lngFontColor
        '��ҳ���
        objDraw.CurrentX = lngLeft: objDraw.CurrentY = lngTop
    Else
        objDraw.CurrentY = lngOldY
    End If
    
    Set NewPrintPage = objDraw
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DrawPatiInfo(objDraw As Object, X As Long, Y As Long, strInfo As String, ByVal lngPaperWidth As Long) As Boolean
    '����:��ӡ���в�����Ϣ
    '����:strInfo=������סԺ�š����ҡ����������š�����
    
    '˳��:
    '   1��DrawPatiInfo
    '   2��DrawBodyInfo
    '   3��DrawBodyScale
    '   4��DrawBodyTopScale
    '   5��DrawBodyPaper
    '   6��DrawBodyGraph
    '   7��DrawBodyRecordItem
    '   8��DrawBodyTips
    '   9��DrawBodyPageNO
    '��ͼ˳���:=1/9   ���ǻ�����ͼ�εĵ�һ��
    Dim strTmp As String
    Dim W_9pt As Long, H_9pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    
    If Trim(strInfo) = "" Then Exit Function
    If InStr(1, strInfo, "'") < 1 Then Exit Function
    If UBound(Split(strInfo, "'")) < 6 Then Exit Function
    On Error GoTo errHand
    
    lngTmpX = X
    lngTmpY = Y
    
    W_9pt = objDraw.TextWidth("��")
    H_9pt = objDraw.TextHeight("��")
    
    Call DrawText(objDraw, X, Y, "����:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("����:"), Y, Split(strInfo, "'")(0), 16711680)
    X = X + objDraw.TextWidth("����:" & Split(strInfo, "'")(0)) + W_9pt / 3
    
    Call DrawText(objDraw, X, Y, "�Ա�:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("�Ա�:"), Y, Split(strInfo, "'")(5), 16711680)
    X = X + objDraw.TextWidth("�Ա�:" & Split(strInfo, "'")(5)) + W_9pt / 3
    
    Call DrawText(objDraw, X, Y, "����:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("����:"), Y, Split(strInfo, "'")(6), 16711680)
    X = X + objDraw.TextWidth("����:" & Split(strInfo, "'")(6)) + W_9pt / 3
    
    Call DrawText(objDraw, X, Y, "����:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("����:"), Y, Split(strInfo, "'")(4), 16711680)
    X = X + objDraw.TextWidth("����:" & Split(strInfo, "'")(4)) + W_9pt / 3
    
    Call DrawText(objDraw, X, Y, "��Ժ����:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("��Ժ����:"), Y, Split(strInfo, "'")(3), 16711680)
    X = X + objDraw.TextWidth("��Ժ����:" & Split(strInfo, "'")(3)) + W_9pt / 3
    
    Call DrawText(objDraw, X, Y, "סԺ��:", 0)
    Call DrawText(objDraw, X + objDraw.TextWidth("סԺ��:"), Y, Split(strInfo, "'")(1), 16711680)
    X = X + objDraw.TextWidth("סԺ��:" & Split(strInfo, "'")(1)) + W_9pt / 3

    X = lngTmpX
    Y = lngTmpY + H_9pt + H_9pt \ 2
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub DrawBodyRecordItem(objDraw As Object, X As Long, Y As Long, ByVal colFirstWidth As Long, strValue() As String, ByVal rsArrRecordItemInfo As ADODB.Recordset, _
    ByVal strDays As String, ByVal strOPS As String, Optional sngScale As Single = 1, Optional ByVal intGridRows As Integer = 0)
    '******************************************************************************************************************
    '����:���ײ���¼����Ŀ
    '����:intRows=Ҫ��������
    '     colFirstWidth=���еĿ��
    '     strValue()=�ַ��б���������ʾ��ʾ��ֵ  ˵����strValue()����ΪEmpty
    
    '˳��:
    '��ͼ˳���:=7/9   ���ǻ�����ͼ�εĵ��߲�
    '******************************************************************************************************************
    
    Dim intRow As Integer, intCol As Integer 'ѭ��֮��
    Dim H_9pt As Long, W_9pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    Dim strTmp As String
    Dim intAdd As Integer
    Dim blnExistData As Boolean
    Dim blnGrade As Boolean
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim lngColor As Long
    Dim intFactRow As Integer
    Dim intCount As Integer
    Dim bln���� As Boolean
    Dim int������������ʽ As Integer
    Dim int����λ�� As Integer
    On Error GoTo ErrHead
    '�̶��ӵ����п�ʼ,�����������
    
    If UBound(strValue) < 0 Then Exit Sub
    If IsEmpty(strValue) = True Then Exit Sub
    int����λ�� = 1
    int������������ʽ = zlDatabase.GetPara("����������", glngSys, 1255, 0)
    
    objDraw.FontSize = 9 * sngScale
    H_9pt = objDraw.TextHeight("��")
    W_9pt = objDraw.TextWidth("��")
    objDraw.DrawStyle = 0
    lngTmpX = X
    lngTmpY = Y
    
    intAdd = 1
    If mbln�������� Then intAdd = 0
                
    If InStr(1, strValue(0), ";") > 0 Then
        
        intFactRow = LBound(strValue) - 1
        int����λ�� = 2
        For intRow = LBound(strValue) To UBound(strValue)
            
            objDraw.FontSize = 9 * sngScale
            
            If intGridRows = 0 Or intGridRows > intFactRow + IIf(bln����, -1, 0) Then

                If Split(strValue(intRow), ";")(0) = "����" Then
                    bln���� = True
                    intFactRow = intFactRow + 2
                    
                    For intCol = 0 To 42
                        If intCol = 0 Then
                        
                            Call DrawCell(objDraw, Split(strValue(intRow), ";")(intCol) & Split(strValue(intRow), ";")(intCol + 2), lngTmpX, Y + intFactRow * (H_9pt + H_9pt \ 2), colFirstWidth, 2 * (H_9pt + H_9pt \ 2), , , , , , objDraw.Font, , 1, 1)
                            Call DrawLine(objDraw, lngTmpX, Y + intFactRow * (H_9pt + H_9pt \ 2), lngTmpX, Y + intFactRow * (H_9pt + H_9pt \ 2) + 2 * (H_9pt + H_9pt \ 2), , , 2)
                            Call DrawLine(objDraw, lngTmpX, Y + intFactRow * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth, Y + intFactRow * (H_9pt + H_9pt \ 2), , , 1)
                            Call DrawLine(objDraw, lngTmpX + colFirstWidth, Y + intFactRow * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth, Y + intFactRow * (H_9pt + H_9pt \ 2) + 2 * (H_9pt + H_9pt \ 2), , , 1)
                            
                        Else
                            If intCol > UBound(Split(strValue(intRow), ";")) Then
                                strTmp = ""
                            Else
                                strTmp = Split(strValue(intRow), ";")(intCol + 2)
                            End If
                            
                            '��ӡ����ֵ���������ӡ��
                            If int������������ʽ = 0 Then
                                If intCol Mod 2 = 0 Then
                                    Call DrawCell(objDraw, strTmp, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol - 1), Y + intFactRow * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips, (H_9pt + H_9pt \ 2), , , , COLOR.��ɫ, , objDraw.Font, "1000")
                                Else
                                    Call DrawCell(objDraw, strTmp, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol - 1), Y + intFactRow * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), HOUR_STEP_Twips, (H_9pt + H_9pt \ 2), , , , COLOR.��ɫ, , objDraw.Font, "0000")
                                End If
                            Else
                                If int����λ�� = 2 Then
                                    Call DrawCell(objDraw, strTmp, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol - 1), Y + intFactRow * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips, (H_9pt + H_9pt \ 2), , , , COLOR.��ɫ, , objDraw.Font, "1000")
                                Else
                                    Call DrawCell(objDraw, strTmp, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol - 1), Y + intFactRow * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), HOUR_STEP_Twips, (H_9pt + H_9pt \ 2), , , , COLOR.��ɫ, , objDraw.Font, "0000")
                                End If
                                If strTmp <> "" Then
                                    int����λ�� = int����λ�� + 1
                                    If int����λ�� > 2 Then int����λ�� = 1
                                End If
                            End If
                        End If
                    Next
                    
                    For intCol = 0 To 42
                        If (intCol) Mod 6 = 0 Then
                            Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * intCol, Y + intFactRow * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * intCol, Y + intFactRow * (H_9pt + H_9pt \ 2) + 2 * (H_9pt + H_9pt \ 2), IIf(intCol = 0, 0, vbRed), , 2)
                        Else
                            Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * intCol, Y + intFactRow * (H_9pt + H_9pt \ 2) - 15, lngTmpX + colFirstWidth + HOUR_STEP_Twips * intCol, Y + intFactRow * (H_9pt + H_9pt \ 2) + 2 * (H_9pt + H_9pt \ 2) + 15)
                            Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol + 1), Y + intFactRow * (H_9pt + H_9pt \ 2) - 15, lngTmpX + colFirstWidth + HOUR_STEP_Twips * (intCol + 1), Y + intFactRow * (H_9pt + H_9pt \ 2) + 2 * (H_9pt + H_9pt \ 2) + 15)
                            
                        End If
                    Next
                Else
                    blnExistData = False
                    rsArrRecordItemInfo.Filter = ""
                    rsArrRecordItemInfo.Filter = "���=" & intRow
                    If rsArrRecordItemInfo.RecordCount > 0 Then
                        If rsArrRecordItemInfo("��Ŀ����").Value = 2 Then
                            '�ǻ��Ŀ��Ҫ�ж��Ƿ�Ϊ�գ���Ϊ��ֵ���򲻴�ӡ����
                            
                            For intCol = 1 To 14
                                
                                If intCol <= UBound(Split(strValue(intRow), ";")) Then
                                    strTmp = Trim(Split(strValue(intRow), ";")(intCol + 2))
                                    If strTmp <> "" Then
                                        blnExistData = True
                                    End If
                                End If
                                
                            Next
                        Else
                            blnExistData = True
                        End If
                    Else
                        blnExistData = True
                    End If
                    
                    If blnExistData Then
                        
                        intFactRow = intFactRow + 1
                        
                        For intCol = 0 To 14
                            If intCol = 0 Then
                            
                                strTmp = Split(strValue(intRow), ";")(intCol) & Split(strValue(intRow), ";")(intCol + 2)
                                
                                Call DrawCell(objDraw, strTmp, lngTmpX + IIf(intFactRow >= 4 And intFactRow <= 8, 200, 0), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), colFirstWidth, (H_9pt + H_9pt \ 2), , , , , , objDraw.Font, , 1, 1)
                                Call DrawLine(objDraw, lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
                            Else
        
                                Select Case Split(strValue(intRow), ";")(0)
                                Case "Ѫѹ"
                                    objDraw.FontSize = 7.8 * sngScale
                                Case Else
                                      
                                End Select
                                
                                blnGrade = False
                                
                                If intCol Mod 2 = 1 Then
                                    'һ�������
                                    If (Split(strValue(intRow), ";")(intCol + 2) <> "" And Split(strValue(intRow), ";")(intCol + 3) <> "") Or Val(Split(strValue(intRow), ";")(1)) = 2 Then
                                        '�����綼��ֵ
                                        
                                        If intCol > UBound(Split(strValue(intRow), ";")) Then
                                            strTmp = ""
                                        Else
                                            strTmp = Split(strValue(intRow), ";")(intCol + 2)
                                        End If
                                        
                                        If Split(strValue(intRow), ";")(0) = "������" And strTmp <> "" Then
                                            '����Ƿ�Ϊ���������ǣ���ֱ��������������ӡ���ĸ���ݵ�����
                                            blnGrade = AnsyGrade(strTmp, str1, str2, str3)
                                        End If
                                
                                        lngColor = GridTextColor(Split(strValue(intRow), ";")(0), strTmp)
                                        Call DrawCell(objDraw, IIf(blnGrade, "", strTmp), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 3, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1)
                                        If blnGrade Then
                                            objDraw.FontSize = 7.5 * sngScale
                                            Call DrawGrade(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 3, (H_9pt + H_9pt \ 2), str1, str2, str3, 0)
                                            objDraw.FontSize = 9 * sngScale
                                        End If
                                        
                                        If (intCol + 1) Mod 2 = 0 Then
                                            Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
                                        End If
                                        
                                    Else
                                        '�ϲ���ӡ
                                        If intCol > UBound(Split(strValue(intRow), ";")) Then
                                            strTmp = ""
                                        Else
                                            If Split(strValue(intRow), ";")(intCol + 2) <> "" Then
                                                strTmp = Split(strValue(intRow), ";")(intCol + 2)
                                            ElseIf Split(strValue(intRow), ";")(intCol + 3) <> "" Then
                                                strTmp = Split(strValue(intRow), ";")(intCol + 3)
                                            Else
                                                strTmp = ""
                                            End If
                                        End If
                                        
                                        If Split(strValue(intRow), ";")(0) = "������" And strTmp <> "" Then
                                            '����Ƿ�Ϊ���������ǣ���ֱ��������������ӡ���ĸ���ݵ�����
                                            blnGrade = AnsyGrade(strTmp, str1, str2, str3)
                                        End If
                                                                        
                                        lngColor = GridTextColor(Split(strValue(intRow), ";")(0), strTmp)
                                        Call DrawCell(objDraw, IIf(blnGrade, "", strTmp), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 6 * Fix(intCol / 2), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 6, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1)
                                        If blnGrade Then
                                            objDraw.FontSize = 7.5 * sngScale
                                            Call DrawGrade(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 6 * Fix(intCol / 2), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 6, (H_9pt + H_9pt \ 2), str1, str2, str3, 0)
                                            objDraw.FontSize = 9 * sngScale
                                        End If
                                        
                                        Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 6 * Fix(intCol / 2), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), IIf(intCol = 1, 0, vbRed), , 2)
                                        
                                        intCol = intCol + 1
                                    End If
                                Else
                                    'һ�������
                                    If intCol > UBound(Split(strValue(intRow), ";")) Then
                                        strTmp = ""
                                    Else
                                        strTmp = Split(strValue(intRow), ";")(intCol + 2)
                                    End If
            
                                    If Split(strValue(intRow), ";")(0) = "������" And strTmp <> "" Then
                                        '����Ƿ�Ϊ���������ǣ���ֱ��������������ӡ���ĸ���ݵ�����
                                        blnGrade = AnsyGrade(strTmp, str1, str2, str3)
                                    End If
        
                                    lngColor = GridTextColor(Split(strValue(intRow), ";")(0), strTmp)
                                    Call DrawCell(objDraw, IIf(blnGrade, "", strTmp), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 3, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1)
                                    If blnGrade Then
                                        objDraw.FontSize = 7.5 * sngScale
                                        Call DrawGrade(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 3, (H_9pt + H_9pt \ 2), str1, str2, str3, 0)
                                        objDraw.FontSize = 9 * sngScale
                                    End If
                                    
                                    If (intCol + 1) Mod 2 = 0 Then
                                        Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
                                    End If
                                End If
                            End If
                            
                            If intCol = 14 Then
                                Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), vbRed, , 2)
                            End If
                        Next
                    End If
                End If
                
    '            If intRow = UBound(strValue) Then
    '                Call DrawLine(objDraw, lngTmpX, y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
    '            End If
    
            End If
        Next

'        '������
'        If intGridRows > 0 And intGridRows > intFactRow + IIf(bln����, -1, 0) Then
'            intCount = intGridRows - intFactRow
'            For intRow = 1 To intCount
'                intFactRow = intFactRow + 1
'                Call DrawCell(objDraw, "", lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), colFirstWidth, (H_9pt + H_9pt \ 2), , , , , , objDraw.Font, , 1, 1)
'                Call DrawLine(objDraw, lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
'
'                For intCol = 1 To 14
'                    Call DrawCell(objDraw, "", lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 3, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font)
'                    If (intCol + 1) Mod 2 = 0 Then
'                        Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), IIf(intCol = 1, 0, vbRed), , 2)
'                    End If
'                    If intCol = 14 Then
'                        Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), vbRed, , 2)
'                    End If
'                Next
'            Next
'            intFactRow = intGridRows
'        End If
        
'        '����
        objDraw.FontSize = 9 * sngScale     'Ѫѹ�޸��������С,�˴���ԭ
        intFactRow = intFactRow + 1
        Call DrawCell(objDraw, "����������", lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), colFirstWidth, (H_9pt + H_9pt \ 2), , , , , , objDraw.Font, , 1, 1)
        Call DrawLine(objDraw, lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
        For intCol = 1 To 14
            If (intCol + 1) Mod 2 = 0 Then
                Call DrawCell(objDraw, GetSplitStr(strOPS, CInt((intCol - 1) / 2)), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 6, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1)
                Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), IIf(intCol = 1, 0, vbRed), , 2)
            End If
            If intCol = 14 Then
                Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), vbRed, , 2)
            End If
        Next
        
        'סԺ����
        Dim lngStartDay As Long
        intFactRow = intFactRow + 1
        lngStartDay = GetSplitStr(strDays, 0)
        Call DrawCell(objDraw, "סԺ����", lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), colFirstWidth, (H_9pt + H_9pt \ 2), , , , , , objDraw.Font, , 1, 1)
        Call DrawLine(objDraw, lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
        For intCol = 1 To 14
            If (intCol + 1) Mod 2 = 0 Then
                Call DrawCell(objDraw, CStr(lngStartDay + CInt((intCol - 1) / 2)), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), HOUR_STEP_Twips * 6, (H_9pt + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1)
                Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1), Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), IIf(intCol = 1, 0, vbRed), , 2)
            End If
            If intCol = 14 Then
                Call DrawLine(objDraw, lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * (intCol - 1) + HOUR_STEP_Twips * 3, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), vbRed, , 2)
            End If
        Next
        
        Call DrawLine(objDraw, lngTmpX, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), lngTmpX + colFirstWidth + HOUR_STEP_Twips * 3 * 14, Y + (intFactRow + intAdd) * (H_9pt + H_9pt \ 2) + (H_9pt + H_9pt \ 2), , , 2)
         
        '�̶��ӵ����п�ʼ,�����������(���ں���ռ����,ʵ�ʴӵ����п�ʼ),�����м�λ�����,���Դ�7�п�ʼ
        Call DrawRotateText(objDraw, lngTmpX + 30, Y + 7 * (H_9pt + H_9pt \ 2), "��")
        Call DrawRotateText(objDraw, lngTmpX + 30, Y + 7 * (H_9pt + H_9pt \ 2) + 200, "��")
        
        'Ϊ������ͼ��׼��
        X = lngTmpX + W_9pt * 2
        Y = lngTmpY + (intFactRow - LBound(strValue) + 1 + intAdd) * (H_9pt + H_9pt \ 2) + 30
    Else
        'Ϊ������ͼ��׼��
        X = lngTmpX + W_9pt * 2
        Y = lngTmpY + 30
        
    End If
    
    objDraw.FontSize = 9 * sngScale
    Exit Sub
ErrHead:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GridTextColor(ByVal strItemName As String, strResult As String) As Long
    
    GridTextColor = 0
    
    Select Case strItemName
    Case ""
    Case Else

        If InStr(strItemName, "Ƥ��") > 0 Then
            If InStr(strResult, "+") > 0 Then
                GridTextColor = COLOR.��ɫ
            End If
        End If
        
    End Select

End Function

Private Function AnsyGrade(ByVal strText As String, ByRef str1 As String, ByRef str2 As String, ByRef str3 As String) As Boolean
    
    '���ܣ���������
    
    Dim intPos As Integer
    
    If strText = "" Then Exit Function
    If InStr(strText, "/") = 0 Then Exit Function
    intPos = InStr(strText, "+")
    If intPos > 0 Then
        If intPos = 1 Then Exit Function
        str1 = Mid(strText, 1, intPos - 1)
        strText = Mid(strText, intPos + 1)
    End If
    
    intPos = InStr(strText, "/")
    If intPos > 0 Then
        If intPos = 1 Then Exit Function
        If intPos = Len(strText) Then Exit Function
        
        str2 = Mid(strText, 1, intPos - 1)
        str3 = Mid(strText, intPos + 1)
    End If
    
    AnsyGrade = True
    
End Function

Private Function DrawGrade(objDraw As Object, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, ByVal Text1 As String, ByVal Text2 As String, ByVal Text3 As String, Optional ByVal ForeColor As Long = 0) As Boolean
    
    '���ܣ���ӡ����
    
    Dim W0 As Long
    Dim H0 As Long
    Dim X0 As Long
    Dim Y0 As Long
    Dim L0 As Long
    
    '�������ռ�õĿ��
    W0 = objDraw.TextWidth(Text1)
    H0 = objDraw.TextHeight("A")
    
    If Len(Text2) > Len(Text3) Then
        L0 = objDraw.TextWidth(Text2)
    Else
        L0 = objDraw.TextWidth(Text3)
    End If
    
    W0 = W0 + L0
    
    If Text1 <> "" Then
        X0 = X + (W - W0) / 2
        Y0 = Y + (H - H0) / 2
        Call DrawText(objDraw, X0, Y0, Text1, 0)
    End If
    
    X0 = X + (W - W0) / 2 + objDraw.TextWidth(Text1) + (W0 - objDraw.TextWidth(Text2)) / 2
    Y0 = Y - 15
        
    Call DrawText(objDraw, X0, Y0, Text2, 0)
    
    Y0 = Y0 + H0
    X0 = X + (W - W0) / 2 + objDraw.TextWidth(Text1)
    Call DrawLine(objDraw, X0, Y0, X0 + L0 + objDraw.TextWidth("A"), Y0, , , 1)

    Y0 = Y0 + 0
    X0 = X + (W - W0) / 2 + objDraw.TextWidth(Text1) + (W0 - objDraw.TextWidth(Text3)) / 2
    Call DrawText(objDraw, X0, Y0, Text3, 0)
                                     
End Function

Private Sub DrawBodyTips(objDraw As Object, X As Long, Y As Long, ByVal colFirstWidth As Long, ByVal strTips As String, Optional sngScale As Single = 1)
    '���ܣ�������ײ��ļ�¼��˵��
    '����:strTips=˵���ַ���
    
    '˳��:
    '��ͼ˳���:=8/9   ���ǻ�����ͼ�εĵڰ˲�
    Dim H_9pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    lngTmpX = X
    lngTmpY = Y
    H_9pt = objDraw.TextHeight("��")
    Call DrawCell(objDraw, strTips, X, Y, colFirstWidth + HOUR_STEP_Twips * 6 * 7, H_9pt + H_9pt \ 2, , , , , , objDraw.Font, "0000", 0)
    X = lngTmpX
    Y = lngTmpY + H_9pt * 2
End Sub

Private Sub DrawBodyPageFooter(objDraw As Object, X As Long, Y As Long, ByVal intPageNo As Integer, ByVal intBeginPage As Integer)
    '******************************************************************************************************************
    '���ܣ�������ײ�˵��
    '����:intNOPage=ҳ��
    '******************************************************************************************************************
    Dim blnWeek As Boolean
    Dim blnPageNo As Boolean
    Dim H_9pt As Long, W_9pt As Long
    Dim strNOPage As String
    Dim lngX As Long
        
    H_9pt = objDraw.TextHeight("��")
    W_9pt = objDraw.TextWidth("��")
    
    blnWeek = (Val(zlDatabase.GetPara("��ӡ����", glngSys, 1255, "0")) = 1)
    blnPageNo = (Val(zlDatabase.GetPara("��ӡҳ��", glngSys, 1255, "1")) = 1)
    
    '��ӡҳ��
    '------------------------------------------------------------------------------------------------------------------
    If intPageNo > -1 And blnPageNo Then
        intPageNo = intPageNo + intBeginPage - 1
        strNOPage = "��  -- " & CStr(intPageNo) & " --  ҳ"
    End If
    
    If blnWeek Then
        If strNOPage = "" Then
            strNOPage = "��  -- " & CStr(intBeginPage) & " --  ��"
        Else
            strNOPage = strNOPage & "(�� " & CStr(intBeginPage) & " ��)"
        End If
        
    End If
    
    Call DrawCell(objDraw, strNOPage, X + ((6 * W_9pt + HOUR_STEP_Twips * 6 * 7) - objDraw.TextWidth(strNOPage)) / 2, Y, W_9pt * 20, H_9pt + H_9pt \ 2, , , , , , objDraw.Font, "0000", 0, 0)
End Sub

Private Function DrawBodyInfo(objDraw As Object, X As Long, Y As Long, ByVal colFirstWidth As Long, _
    ByVal strDate As String, ByVal strPatiDay As String, ByVal strOPSFate As String, Optional sngScale As Long = 1, Optional ByVal strBeginDate As String) As Boolean
    '����:������ǰҳ���סԺ���������ڵ�
    '����:objDraw=�������
    '     colFirstWidth = ���еĿ��
    '     strDate=��ҳ��סԺ�����ַ���
    '     strPatiDay=סԺ�����ַ���
    '     strOPSFate=�������ַ���
    
    '˳��:
    '��ͼ˳���:=2/9   ���ǻ�����ͼ�εĵڶ���
    
    Dim H_9pt As Long
    Dim intDay As Integer
    Dim lngLeft As Long
    Dim lngTmpX As Long, lngTmpY As Long
    Dim strTmp As String
    Dim lngStartDay As Long
    
    DrawBodyInfo = True
    
    lngTmpX = X
    lngTmpY = Y
    '�ο��߶�
    objDraw.Font.Name = "����"
    objDraw.Font.Size = 9 * sngScale
    objDraw.Font.Bold = False
    H_9pt = objDraw.TextHeight("��")
    If colFirstWidth < H_9pt + H_9pt \ 2 Then DrawBodyInfo = False: Exit Function
    objDraw.DrawStyle = 0
    '������
   
    Call DrawCell(objDraw, "��    ��", X, Y + (H_9pt + H_9pt \ 2) * 0, colFirstWidth, H_9pt + H_9pt \ 2, 0, 0, , , , objDraw.Font, "1111", 1, 1, False)
'    Call DrawCell(objDraw, "סԺ����", X, Y + (H_9pt + H_9pt \ 2) * 1, colFirstWidth, H_9pt + H_9pt \ 2, 0, 0, , , , objDraw.Font, "1111", 1, 1, False)
'    Call DrawCell(objDraw, "��/�������", X, Y + (H_9pt + H_9pt \ 2) * 2, colFirstWidth, H_9pt + H_9pt \ 2, 0, 0, , , , objDraw.Font, "1111", 1, 1, False)
    
    Call DrawLine(objDraw, X, Y, X + colFirstWidth, Y, , , 2)
    Call DrawLine(objDraw, X, Y, X, Y + (H_9pt + H_9pt \ 2) * 2 + H_9pt + H_9pt \ 2, , , 2)
'    Call DrawLine(objDraw, X + colFirstWidth, Y, X + colFirstWidth, Y + (H_9pt + H_9pt \ 2) * 2 + H_9pt + H_9pt \ 2, , , 2)
    
    '������
    lngLeft = lngTmpX + colFirstWidth
    lngStartDay = GetSplitStr(strPatiDay, 0)
    
    For intDay = 0 To 6

        strTmp = GetSplitStr(strDate, intDay)
        If Right(strTmp, 5) = "01-01" Then
            'һ��ĵ�һ��

        ElseIf strTmp = Format(strBeginDate, "yyyy-MM-dd") Then
            '��Ժ��һ�죬д�����

        ElseIf intDay = 0 Then
            strTmp = Right(strTmp, 5)
        ElseIf Right(strTmp, 2) = "01" Then
            strTmp = Right(strTmp, 5)
        Else
            strTmp = Right(strTmp, 2)
        End If

        Call DrawCell(objDraw, strTmp, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y, (HOUR_STEP_Twips * 6), H_9pt + H_9pt \ 2, 0, 0, , &HFF0000, , objDraw.Font, "1111", 1, 1, False)
        
        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y, , , 2)
        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + H_9pt + H_9pt \ 2, IIf(intDay = 0, 0, vbRed), , 2)
        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y + H_9pt + H_9pt \ 2, vbRed, , 2)

'        Call DrawCell(objDraw, CStr(lngStartDay + intDay), lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + H_9pt * 1 + H_9pt \ 2, (HOUR_STEP_Twips * 6), H_9pt + H_9pt \ 2, 0, 0, , &HFF0000, , objDraw.Font, "1111", 1, 1, False)
'
'        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + H_9pt * 1 + H_9pt \ 2, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + H_9pt * 1 + H_9pt \ 2 + H_9pt + H_9pt \ 2, , , 2)
'        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y + H_9pt * 1 + H_9pt \ 2, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y + H_9pt * 1 + H_9pt \ 2 + H_9pt + H_9pt \ 2, , , 2)
'
'        If intDay <= UBound(Split(strOPSFate, "'")) + 1 Then
'            Call DrawCell(objDraw, GetSplitStr(strOPSFate, intDay), lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + (H_9pt + H_9pt \ 2) * 2, (HOUR_STEP_Twips * 6), H_9pt + H_9pt \ 2, 0, 0, , 255, , objDraw.Font, "1111", 1, 1, False)
'        Else
'            Call DrawCell(objDraw, "", lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + (H_9pt + H_9pt \ 2) * 2, (HOUR_STEP_Twips * 6), H_9pt + H_9pt \ 2, 0, 0, , 255, , objDraw.Font, "1111", 1, 1, False)
'        End If
'        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + (H_9pt + H_9pt \ 2) * 2, lngLeft + (HOUR_STEP_Twips * 6) * intDay, Y + (H_9pt + H_9pt \ 2) * 2 + H_9pt + H_9pt \ 2, , , 2)
'        Call DrawLine(objDraw, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y + (H_9pt + H_9pt \ 2) * 2, lngLeft + (HOUR_STEP_Twips * 6) * intDay + (HOUR_STEP_Twips * 6), Y + (H_9pt + H_9pt \ 2) * 2 + H_9pt + H_9pt \ 2, , , 2)
                  
    Next
    X = lngTmpX
    Y = lngTmpY + (H_9pt + H_9pt \ 2)
End Function

Private Sub DrawBodyTopScale(objDraw As Object, X As Long, Y As Long, Optional sngScale As Long = 1)
    '����:��������ı��
    '
    
    '˳��:
    '��ͼ˳���:=4/9   ���ǻ�����ͼ�εĵ��Ĳ�
    Dim i As Integer, j As Integer 'ѭ��֮��
    Dim H_9pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    
    Dim lngHourBegin As Long
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    lngHourBegin = Val(zlDatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4))
    lngTmpX = X: lngTmpY = Y
    
    '�ο��߶�
    objDraw.Font.Name = "����"
    objDraw.Font.Size = 9 * sngScale
    objDraw.Font.Bold = False
    
    objDraw.DrawStyle = 0
    H_9pt = objDraw.TextHeight("��")
'    For i = 1 To 7
'        '��������
'        Call DrawCell(objDraw, "����", X, Y, HOUR_STEP_Twips * 3, (H_9pt * 2), , , , , , , , 1, 1)
'        Call DrawLine(objDraw, X, Y, X, Y + (H_9pt * 2), , , 2)
'
'        X = X + HOUR_STEP_Twips * 3
'
'        Call DrawCell(objDraw, "����", X, Y, HOUR_STEP_Twips * 3, (H_9pt * 2), , , , , , , , 1, 1)
'        Call DrawLine(objDraw, X + HOUR_STEP_Twips * 3, Y, X + HOUR_STEP_Twips * 3, Y + (H_9pt * 2), , , 2)
'
'        X = X + HOUR_STEP_Twips * 3
'
'    Next
'    Y = lngTmpY + H_9pt * 2
'    X = lngTmpX
    
    Dim intCount As Integer
    Dim lngColor As Long
    
    For i = 1 To 7
        
        intCount = 0
        For j = lngHourBegin To 12 - (4 - lngHourBegin) Step 4
            intCount = intCount + 1
            If intCount < 2 Then
                lngColor = RGB(200, 0, 0)
            Else
                lngColor = 0
            End If
            
            If j = 12 - lngHourBegin Then
'                objDraw.FontSize = 8
                Call DrawCell(objDraw, CStr(j), X, Y, HOUR_STEP_Twips, (H_9pt * 2 + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1, 1, True)
                X = X + HOUR_STEP_Twips
            Else
'                objDraw.FontSize = 9
                Call DrawCell(objDraw, CStr(j), X, Y, HOUR_STEP_Twips, (H_9pt * 2 + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1, 1, True)
                X = X + HOUR_STEP_Twips
            End If
        Next
        
        intCount = 0
        For j = lngHourBegin To 12 - (4 - lngHourBegin) Step 4
            intCount = intCount + 1
            If intCount >= 2 Then
                lngColor = RGB(200, 0, 0)
            Else
                lngColor = 0
            End If
            If j = 12 - lngHourBegin Then
'                objDraw.FontSize = 8
                Call DrawCell(objDraw, CStr(j), X, Y, HOUR_STEP_Twips, (H_9pt * 2 + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1, 1, True)
                X = X + HOUR_STEP_Twips
            Else
'                objDraw.FontSize = 9
                Call DrawCell(objDraw, CStr(j), X, Y, HOUR_STEP_Twips, (H_9pt * 2 + H_9pt \ 2), , , , lngColor, , objDraw.Font, , 1, 1, True)
                X = X + HOUR_STEP_Twips
            End If
        Next
        
    Next
    
    X = lngTmpX
    For i = 1 To 7
        X = X + HOUR_STEP_Twips * 6
        Call DrawLine(objDraw, X, Y, X, Y + (H_9pt * 2 + H_9pt \ 2), vbRed, 1, 2)
    Next
    
    objDraw.FontSize = 9
    X = lngTmpX
    Y = lngTmpY + H_9pt * 2 + H_9pt \ 2
End Sub
'
'Private Function DrawBodyGraph(objDraw As Object, _
'                                x As Long, _
'                                y As Long, _
'                                strArrItemInfo() As String, _
'                                strArrValue() As String, _
'                                strArrValueComment() As String, _
'                                strArrValueInOut() As String, _
'                                Optional strIndex As String = "", _
'                                Optional sngScale As Long = 1) As Boolean
'
'    '����:���ڲ���������ݵ�����
'    '����:strArrItemInfo()=���±���Ŀ��Ϣ  ��ʽ��(0)="��¼ɫ'�����'��Ŀ��'���ֵ'��Сֵ'��λֵ'��¼��"
'    '                                                  ... ...
'    '                                            (n)="��¼ɫ'�����'��Ŀ��'���ֵ'��Сֵ'��λֵ'��¼��"
'    '     strArrValue()=ֵ����  ��ʽ��(0)="43'54'65'-2'76'57'657'543 ... ...  65'876'"  <---42(7��*6)�����ݣ�ĳ��ĳʱû������Ϊ-2
'    '                                 (1)="87'987'09'5'36'-2'865'963 ... ...  -2'53'"   <---42(7��*6)�����ݣ�ĳ��ĳʱû������Ϊ-2
'    '                                 ......
'    '                                 (n)="453'445'46'-2'33'-2'865'45 ... ...  3'3'"   <---42(7��*6)�����ݣ�ĳ��ĳʱû������Ϊ-2
'    '                                       ����Ĳ���˵���������Ԫ�ظ���Ӧ�����±���Ŀ��Ϣ��Ԫ�ظ�����ͬ
'    '     strArrValueComment()=˵������ ��ʽ��ֵͬ����
'    '     strIndex=���ַ����������Ҫ����Щ���� ��ʽ��"0'4'2'3"  ��Ҫ���б����ܺͲ��ܴ��ڡ�ֵ���顱Ԫ�صĸ�����
'    '��ͼ˳���:=6/9   ���ǻ�����ͼ�εĵ�����
'
'    Dim lngRow As Long, lngCol As Long 'ѭ��֮��
'    Dim lngTmpX As Long, lngTmpY As Long
'    Dim H_9pt As Long
'    Dim W_9pt As Long
'    Dim lngArrXY() As Long
'    Dim intItemCount As Integer  '��Ŀ����
'    Dim lngMax As Long, lngMin As Long '�����Сֵ
'    Dim lng��Ŀ��� As Long
'    Dim lngColor As Long, lngTopRow As Long, lngStep As Single, strTag As String    '��¼ɫ,�����,��λֵ,��¼��
'    Dim intItemDrawCount As Integer 'Ҫ������Ŀ����
'    Dim lngValue As Double    '��ǰֵ
'    Dim lngValue1 As Double    '��ǰֵ
'    Dim lngValue2 As Double    '��ǰֵ
'    Dim strComment As String  '��ǰҪ����˵��
'    Dim lngUpIndex As Long
'    Dim Y2 As Long, i As Long
'    Dim lngCommentColor As Long
'    Dim aryData() As String
'    Dim aryTmp As Variant
'    Dim blnStop As Boolean
'    Dim rsPoint As New ADODB.Recordset
'
'    Dim X1 As Long
'    Dim Y1 As Long
'    Dim dblHeight As Double         '40-42��֮�����Ч��ӡ�߶�
'    Dim rsTmp As New ADODB.Recordset
'    On Error GoTo errHand
'
'    DrawBodyGraph = False
'    If IsEmpty(strArrItemInfo) Or IsEmpty(strArrValue) Then Exit Function
'    If UBound(strArrItemInfo) <> UBound(strArrValue) Then Exit Function
'    lngTmpY = y
'    lngTmpX = x
'    objDraw.DrawStyle = 0
'    W_9pt = objDraw.TextWidth("��")
''    H_9pt = objDraw.TextHeight("��")
'    H_9pt = ROWHEIGHT * 10 / 3
'
'    intItemCount = UBound(strArrValue) + 1
'    'ReDim lngArrXY(7 * 6 - 1, 4) '��¼����
'    ReDim lngArrXY(7 * 6 - 1, 5) '��¼����      ,�������һά,���ڼ�¼�Ƿ�Ϊ����
'
'    Dim mpt����() As POINTAPI
'    Dim mpt����() As POINTAPI
'    Dim mpt����() As POINTAPI
'
'    ReDim mpt����(0 To 41)
'    ReDim mpt����(0 To 41)
'    ReDim mpt����(0 To 41)
'
'    Dim lngX As Long
'    Dim lngY As Long
'    Dim lngY1 As Long
'    Dim strtmp As String
'    Dim intPointCount As Integer
'    Dim lngYMax As Long
'    Dim intCharNumber As Integer
'    Dim blnPrint As Boolean
'
'    blnPrint = (Val(zlDatabase.GetPara("����ӡ�������ͼ��", glngSys, 1255, "0")) = 0)
'
'    lngYMax = (lngTmpY + 3 * H_9pt / 2) + 40 * 3 * H_9pt / 2 - H_9pt / 2
'    Call PointInit(rsPoint)
'
'    '��ӡ���ת
'    For lngCol = 0 To 41
'
'        lngX = lngTmpX + lngCol * HOUR_STEP_Twips + HOUR_STEP_Twips / 2
'        lngY = (lngTmpY + 3 * H_9pt / 2)
'        lngY1 = (lngTmpY + 3 * H_9pt / 2) + 35 * 3 * H_9pt / 2
'
'        If strArrValueInOut(lngCol) <> "" Then
'
'            strComment = strArrValueInOut(lngCol)
'            aryTmp = Split(strComment, ";")
'
'            Set rsTmp = New ADODB.Recordset
'            '20090926:������40-42�ȼ��ӡ,����һ����Ϣ�����������С����,�ж�����Ϣ���Ӻ���һ���ӡ,��������һ���ֱ��ȫ����ӡ
'            Set rsTmp = New ADODB.Recordset
'            rsTmp.Fields.Append "�к�", adVarChar, 30
'            rsTmp.Fields.Append "ʱ��", adVarChar, 30
'            rsTmp.Fields.Append "���", adVarChar, 50
'            '20090926--
'            rsTmp.Fields.Append "����", adVarChar, 50       '��¼�����ת,������Ժ,����δ��˵��,�ϱ�˵��
'            rsTmp.Fields.Append "��ӡ��", adVarChar, 30
'            rsTmp.Fields.Append "����", adVarChar, 30
'            rsTmp.Fields.Append "�߶�", adVarChar, 30       'δ��˵�����ϱ�˵�����ùܸ߶�
'            rsTmp.Fields.Append "�����С", adVarChar, 50
'            '----------
'            rsTmp.Open
'
'            For i = 0 To UBound(aryTmp)
'                If Trim(aryTmp(i)) <> "" Then
'                    rsTmp.AddNew
'                    rsTmp.Fields("ʱ��").Value = Split(aryTmp(i), "'")(1)
'                    rsTmp.Fields("���").Value = Split(aryTmp(i), "'")(0)
'                End If
'            Next
'            rsTmp.Sort = "ʱ��"
'            If rsTmp.RecordCount > 0 Then
'                rsTmp.MoveFirst
'                strComment = ""
'                Do While Not rsTmp.EOF
'                    strComment = strComment & " " & rsTmp.Fields("���").Value
'                    rsTmp.MoveNext
'                Loop
'
'                strComment = Trim(strComment)
'
'            End If
'
'            intCharNumber = 0
'            For i = 1 To Len(strComment)
'
'                '��42���������1�ǹ̶��������ʽ��
'
'                If lngY < lngYMax Then
'                    strtmp = Mid(strComment, i, 1)
'
'                    If Asc(strtmp) < 0 Then
'                        If intCharNumber Mod 2 = 1 Then lngY = lngY + ROWHEIGHT * 2.5
'                    End If
'
'                    Call DrawRotateText(objDraw, lngX - objDraw.TextWidth(strtmp) / 2 + 15, lngY + 15, strtmp, 255)
'
'                    If Asc(strtmp) < 0 Then
'                        lngY = lngY + ROWHEIGHT * 5
'                        intCharNumber = 0
'                    Else
'                        lngY = lngY + ROWHEIGHT * 2.5
'                        intCharNumber = intCharNumber + 1
'                    End If
'
'
'
'                End If
'            Next
'        End If
'
'        '˵��
'        strComment = Split(strArrValueComment(0), "'")(lngCol)
'        If strComment <> "" Then
'            aryTmp = Split(strComment, ";")
'
'            '�ϱ�
'            If aryTmp(0) <> "" Then
'                If strArrValueInOut(lngCol) <> "" Then aryTmp(0) = " " & aryTmp(0)
'                intCharNumber = 0
'                For i = 1 To Len(aryTmp(0))
'                    If lngY < lngYMax Then
'                        strtmp = Mid(aryTmp(0), i, 1)
'
'                        If Asc(strtmp) < 0 Then
'                            If intCharNumber Mod 2 = 1 Then lngY = lngY + ROWHEIGHT * 2.5
'                        End If
'
'                        Call DrawRotateText(objDraw, lngX - objDraw.TextWidth(strtmp) / 2 + 15, lngY + 15, strtmp, -2147483635)
'
'                        If Asc(strtmp) < 0 Then
'                            lngY = lngY + ROWHEIGHT * 5
'                            intCharNumber = 0
'                        Else
'                            lngY = lngY + ROWHEIGHT * 2.5
'                            intCharNumber = intCharNumber + 1
'                        End If
'
'                    End If
'                Next
'            End If
'
'            '�±�
'            intCharNumber = 0
'            If aryTmp(1) <> "" Then
'                For i = 1 To Len(aryTmp(1))
'                    If lngY1 < lngYMax Then
'                        strtmp = Mid(aryTmp(1), i, 1)
'
'                        If Asc(strtmp) < 0 Then
'                            If intCharNumber Mod 2 = 1 Then lngY = lngY + ROWHEIGHT * 2.5
'                        End If
'
'                        Call DrawRotateText(objDraw, lngX - objDraw.TextWidth(strtmp) / 2 + 15, lngY1 + 15, strtmp, -2147483635)
'
'                        If Asc(strtmp) < 0 Then
'                            lngY1 = lngY1 + ROWHEIGHT * 5
'                            intCharNumber = 0
'                        Else
'                            lngY1 = lngY1 + ROWHEIGHT * 2.5
'                            intCharNumber = intCharNumber + 1
'                        End If
'
'                    End If
'                Next
'            End If
'        End If
'    Next
'
'    Dim int���� As Integer
'    Dim int���� As Integer
'
'    If Trim(strIndex) = "" Then
'
'        '����Ŀ����ѭ��
'
'        For intItemDrawCount = 0 To intItemCount - 1
'
'            intPointCount = -1
'            lngTopRow = CInt(Split(strArrItemInfo(intItemDrawCount), "'")(1))
'            lngStep = Val(Split(strArrItemInfo(intItemDrawCount), "'")(5))
'
'            lngMax = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(3))
'            lngMin = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(4))
'
'            lng��Ŀ��� = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(7))
'            Select Case lng��Ŀ���
'            Case 1
'                int���� = intItemDrawCount
'            Case 2
'                int���� = intItemDrawCount
'            End Select
'            objDraw.ForeColor = lngColor
'
'            'Erase lngArrXY
'            '��ʼ���Ա��ڵ�һ��ʱ����������
'
'            lngUpIndex = -1
'
'            'ѭ��������
'            For lngCol = 0 To 7 * 6 - 1 '��
'
'                lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
'                strTag = Trim(Split(strArrItemInfo(intItemDrawCount), "'")(6))
'
'                x = lngTmpX + (lngCol + 1) * HOUR_STEP_Twips - HOUR_STEP_Twips / 2
'
'                '1�������ǰֵ����ֵ��(��ǰֵ/��λֵ+�����)*(H_9pt+H_9pt\2)������ʱΪ���Y����λ��ֵ
'                '   ������ʱ�����걣���ڱ�����
'                If Val(Split(strArrValue(intItemDrawCount), "'")(lngCol)) = -2 Then
'                    y = lngTmpY
'                    Y2 = lngTmpY
'                    lngArrXY(lngCol, 0) = -1
'                    lngArrXY(lngCol, 1) = x
'                    lngArrXY(lngCol, 2) = y
'                    lngArrXY(lngCol, 3) = Y2
'                    lngArrXY(lngCol, 4) = -1
'                    lngArrXY(lngCol, 5) = 0
'                Else
'                    y = lngTmpY
'                    Y2 = lngTmpY
'
'                    aryData = Split(Split(strArrValue(intItemDrawCount), "'")(lngCol), ",")
'
'                    For i = 0 To UBound(aryData)
'                        lngValue = (lngTopRow - 1) * lngStep + lngMax
'
'                        '������ǰ��Ŀ����Сֵ���ֵʱ,ֱ��ȡ��Сֵ�����ֵ
'                        lngArrXY(lngCol, 5) = IIf(Left(aryData(i), InStr(aryData(i), ";") - 1) = "����", 1, 0)
'
'                        Select Case Val(Left(aryData(i), InStr(aryData(i), ";") - 1))
'                        Case Is < lngMin
'                            aryData(i) = lngMin & ";" & Mid(aryData(i), InStr(aryData(i), ";") + 1)
'                        Case Is > lngMax
'                            aryData(i) = lngMax & ";" & Mid(aryData(i), InStr(aryData(i), ";") + 1)
'                        End Select
'
'                        If InStr(aryData(i), ";") > 0 Then
'                            lngValue1 = lngValue - Val(Left(aryData(i), InStr(aryData(i), ";") - 1))
'                        Else
'                            lngValue1 = lngValue - Val(aryData(i))
'                        End If
'
'                        If i = 0 Then
'                            y = lngTmpY + (lngValue1 / lngStep) * (H_9pt + H_9pt \ 2)
'                            lngArrXY(lngCol, 0) = 2
'                        Else
'                            Y2 = lngTmpY + (lngValue1 / lngStep) * (H_9pt + H_9pt \ 2)
'                            lngArrXY(lngCol, 4) = 2
'                        End If
'
'                        '��Ϊ��3�����ܴ�����ı���˵��������Ŀ�Ĳ�λ��
'                        If i = 1 Then Exit For
'                    Next
''                    lngArrXY(lngCol, 0) = 2
'                    lngArrXY(lngCol, 1) = x
'                    lngArrXY(lngCol, 2) = y
'                    lngArrXY(lngCol, 3) = Y2
'
'                    '2����X-W_9pt\2,Y-H_9pt\2 ����һ����¼��
'                    lngX = x - W_9pt \ 3
'                    lngY = y - H_9pt \ 3
'
'                    lngX = x
'                    lngY = y
'
'                    Select Case lng��Ŀ���
'                    Case 1
'                        If InStr(aryData(0), ";") > 0 Then
'                            aryTmp = Split(aryData(0), ";")
'                            Select Case aryTmp(5)
'                            Case "����"
'                                strTag = mstrChar(0)
'                            Case "Ҹ��"
'                                strTag = mstrChar(1)
'                            Case "����"
'                                strTag = mstrChar(2)
'                            Case Else
'                                strTag = mstrChar(1)
'                            End Select
'
'                            X1 = lngX
'                            Y1 = lngY
'                            Select Case Val(aryTmp(0))
'                            Case Is <= lngMin
'                                strTag = "��"
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                            Case Is >= lngMax
'                                strTag = "��"
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                            End Select
'
'                            If Val(aryTmp(2)) = 1 Then
'                                '���Ժϸ�
'                                Call DrawText(objDraw, X1 - 50, Y1 - 250, "v", lngColor)
'                            End If
'
'                            mpt����(lngCol).x = X1
'                            mpt����(lngCol).y = Y1
'
'                        End If
'                    Case 2
'                        '���Ϊ�ձ�ʾȱʡ�ַ�;����Ϊ��,�ڻ�ͼʱ����Ϊ��,��ȡ��Դ�ļ��е�λͼ
'                        aryTmp = Split(aryData(0), ";")
'                        strTag = IIf(aryTmp(5) = "����", "", mstrPulse)
'
'                        mpt����(lngCol).x = lngX
'                        mpt����(lngCol).y = lngY
'
'                        X1 = lngX
'                        Y1 = lngY
'                        If InStr(aryData(0), ";") > 0 Then
'                            aryTmp = Split(aryData(0), ";")
'                            Select Case Val(aryTmp(0))
'                            Case Is <= lngMin
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                            Case Is >= lngMax
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                            End Select
'                        End If
'
'                    Case -1
'                        mpt����(lngCol).x = lngX
'                        mpt����(lngCol).y = lngY
'
'                        X1 = lngX
'                        Y1 = lngY
'
'                        If InStr(aryData(0), ";") > 0 Then
'                            aryTmp = Split(aryData(0), ";")
'                            Select Case Val(aryTmp(0))
'                            Case Is <= lngMin
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                            Case Is >= lngMax
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                            End Select
'                        End If
'
'                    Case Else
'                        If lng��Ŀ��� = 3 Then
'                            aryTmp = Split(aryData(0), ";")
'                            strTag = IIf(aryTmp(5) = "������", "", mstrBreath)
'                        End If
'
'                        X1 = lngX
'                        Y1 = lngY
'                        Select Case Val(aryData(0))
'                        Case Is <= lngMin
'                            Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                        Case Is >= lngMax
'                            Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                        End Select
'
'                    End Select
'
'                    If lng��Ŀ��� = 1 Then         '����
'                        Call PointAdd(rsPoint, lngX, lngY, lng��Ŀ���, strTag, lngColor, lngCol, CStr(aryTmp(5)))
'                    ElseIf lng��Ŀ��� = 3 Then     '����
'                        Call PointAdd(rsPoint, lngX, lngY, lng��Ŀ���, strTag, lngColor, lngCol, CStr(aryTmp(5)), IIf(strTag = "", "BREATH", ""))
'                    ElseIf lng��Ŀ��� = 2 Then     '����
'                        Call PointAdd(rsPoint, lngX, lngY, lng��Ŀ���, strTag, lngColor, lngCol, CStr(aryTmp(5)), IIf(strTag = "", "PACEMAKER", ""))
'                    Else
'                        Call PointAdd(rsPoint, lngX, lngY, lng��Ŀ���, strTag, lngColor, lngCol, "")
'                    End If
'
'                    lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
'
'                    If Y2 <> lngTmpY Then
'
'                        Select Case lng��Ŀ���
'                        Case 2
'
'                            '��������������ʹ��ʱ
''                            If mint����Ӧ�� = 2 Then
'
'                                strTag = mstr���ʷ���
'                                mpt����(lngCol).x = lngX
'                                mpt����(lngCol).y = (Y2 - H_9pt \ 3)
'                                lngColor = 255
'
'                                If InStr(aryData(1), ";") > 0 Then
'                                    X1 = lngX
'                                    Y1 = Y2
'                                    aryTmp = Split(aryData(1), ";")
'                                    Select Case Val(aryTmp(0))
'                                    Case Is <= lngMin
'                                        Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= lngMax
'                                        Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'                                End If
'
''                            End If
'                        Case 1
'                            strTag = "��"
'                            lngColor = 255
'                        End Select
'
'                        If lng��Ŀ��� = 1 Then
'                            Call PointAdd(rsPoint, lngX, Y2 - H_9pt \ 3, lng��Ŀ���, strTag, lngColor, lngCol, CStr(aryTmp(5)))
'                        Else
'                            Call PointAdd(rsPoint, lngX, Y2 - H_9pt \ 3, lng��Ŀ���, strTag, lngColor, lngCol, "")
'                        End If
'
'                    End If
'
'                    '3���ҵ�ǰһ������Ľ������ߡ������һ������ģ�0��Ԫ��Ϊ0ʱ��ʾ����
'                    '--------------------------------------------------------------------------------------------------
'                    If lngCol > 0 Then
'                        lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
'                        If lngUpIndex > -1 Then
'                            If lngArrXY(lngUpIndex, 0) = 2 Then
'                                If lngArrXY(lngUpIndex, 5) <> 1 And lngArrXY(lngCol, 5) <> 1 Then
'                                    Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 2), x, y, lngColor)
'                                End If
'
'                                If Y2 <> lngTmpY Then
'                                    Select Case lng��Ŀ���
'                                    Case 1          '������Ŀ
'                                        Call DrawLine(objDraw, x, y, x, Y2, 255, IIf(Y2 < y, 0, 2), , True)
'                                    Case 2
'                                        If lngArrXY(lngUpIndex, 4) = 2 Then
'                                            Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 3), x, Y2, lngColor)
'                                        End If
'                                    End Select
'
'                                End If
'                            Else
'                                '��Ȼ�ϵ㣬��������������㣬���Ȼ�����
'                                '--------------------------------------------------------------------------------------
'                                If Y2 <> lngTmpY Then
'                                    Select Case lng��Ŀ���
'                                    Case 1
'                                        Call DrawLine(objDraw, x, y, x, Y2, 255, IIf(Y2 < y, 0, 2), , True)
'                                    End Select
'                                End If
'
'                                '�ҳ����켰������һ����Ч��,ֻ�и��������ݲŲ�����
'                                '--------------------------------------------------------------------------------------
'                                lngUpIndex = lngUpIndex - 1
'                                blnStop = False
'                                Do While lngUpIndex >= (lngCol \ 6) * 6 - 6 And lngUpIndex >= 0
'
'                                    If blnStop = False Then
'                                        If Split(strArrValueComment(0), "'")(lngUpIndex + 1) <> "" Then
'                                            aryTmp = Split(Split(strArrValueComment(0), "'")(lngUpIndex + 1), ";")
'                                        Else
'                                            aryTmp = Split(";;", ";")
'                                        End If
'                                        blnStop = (Val(aryTmp(2)) = 1)
'                                    End If
'                                    If blnStop Then
'                                        Exit Do
'                                    End If
'
'                                    If lngArrXY(lngUpIndex, 0) = 2 And lngArrXY(lngUpIndex, 5) <> 1 And lngArrXY(lngCol, 5) <> 1 Then
'                                        If blnStop = False Then
'                                            blnStop = False
'                                            lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
'
'                                            Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 2), x, y, lngColor)
'
'                                            If Y2 <> lngTmpY Then
'                                                Select Case lng��Ŀ���
'                                                Case 1
'                                                    Call DrawLine(objDraw, x, y, x, Y2, 255, IIf(Y2 < y, 0, 2), , True)
'                                                Case 2                  '���������������
'                                                    If lngArrXY(lngUpIndex, 4) = 2 Then
'                                                        Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 3), x, Y2, lngColor)
'                                                    End If
'                                                End Select
'                                            End If
'                                        End If
'                                        Exit Do
'                                    End If
'                                    lngUpIndex = lngUpIndex - 1
'                                Loop
'
'                                '--------------------------------------------------------------------------------------
'
'                            End If
'                        End If
'                    ElseIf Y2 <> lngTmpY Then
'                        Select Case lng��Ŀ���
'                        Case 1
'                            Call DrawLine(objDraw, x, y, x, Y2, 255, IIf(Y2 < y, 0, 2), , True)
'                        End Select
'                    End If
'                End If
'                lngUpIndex = lngCol
'            Next
'        Next
'
'        '�������������������γɶ���Σ����������ߺ����
'        '--------------------------------------------------------------------------------------------------------------
'        If blnPrint Then Call DrawPoly(objDraw, mpt����, mpt����)
'
'    End If
'
'    '������ַ���ͼ��
'    '--------------------------------------------------------------------------------------------------------------
'    Call DrawPoint(objDraw, rsPoint)
'
'    x = lngTmpX - (UBound(strArrItemInfo) + 1) * (W_9pt + W_9pt \ 2)
'
'    Exit Function
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'End Function

Private Function DrawBodyGraph(objDraw As Object, _
                                X As Long, _
                                Y As Long, _
                                strArrItemInfo() As String, _
                                strArrValue() As String, _
                                strArrValueComment() As String, _
                                strArrValueInOut() As String, _
                                Optional strIndex As String = "", _
                                Optional sngScale As Long = 1) As Boolean
                            
    '����:���ڲ���������ݵ�����
    '����:strArrItemInfo()=���±���Ŀ��Ϣ  ��ʽ��(0)="��¼ɫ'�����'��Ŀ��'���ֵ'��Сֵ'��λֵ'��¼��"
    '                                                  ... ...
    '                                            (n)="��¼ɫ'�����'��Ŀ��'���ֵ'��Сֵ'��λֵ'��¼��"
    '     strArrValue()=ֵ����  ��ʽ��(0)="43'54'65'-2'76'57'657'543 ... ...  65'876'"  <---42(7��*6)�����ݣ�ĳ��ĳʱû������Ϊ-2
    '                                 (1)="87'987'09'5'36'-2'865'963 ... ...  -2'53'"   <---42(7��*6)�����ݣ�ĳ��ĳʱû������Ϊ-2
    '                                 ......
    '                                 (n)="453'445'46'-2'33'-2'865'45 ... ...  3'3'"   <---42(7��*6)�����ݣ�ĳ��ĳʱû������Ϊ-2
    '                                       ����Ĳ���˵���������Ԫ�ظ���Ӧ�����±���Ŀ��Ϣ��Ԫ�ظ�����ͬ
    '     strArrValueComment()=˵������ ��ʽ��ֵͬ����
    '     strIndex=���ַ����������Ҫ����Щ���� ��ʽ��"0'4'2'3"  ��Ҫ���б����ܺͲ��ܴ��ڡ�ֵ���顱Ԫ�صĸ�����
    '��ͼ˳���:=6/9   ���ǻ�����ͼ�εĵ�����
    
    Dim lngRow As Long, lngCol As Long 'ѭ��֮��
    Dim lngTmpX As Long, lngTmpY As Long
    Dim H_9pt As Long
    Dim W_9pt As Long
    Dim lngArrXY() As Long
    Dim intItemCount As Integer  '��Ŀ����
    Dim lngMax As Long, lngMin As Long '�����Сֵ
    Dim lng��Ŀ��� As Long
    Dim lngColor As Long, lngTopRow As Long, lngStep As Single, strTag As String    '��¼ɫ,�����,��λֵ,��¼��
    Dim intItemDrawCount As Integer 'Ҫ������Ŀ����
    Dim lngValue As Double    '��ǰֵ
    Dim lngValue1 As Double    '��ǰֵ
    Dim lngValue2 As Double    '��ǰֵ
    Dim strComment As String  '��ǰҪ����˵��
    Dim lngUpIndex As Long
    Dim Y2 As Long, i As Long
    Dim lngCommentColor As Long
    Dim aryData() As String
    Dim aryTmp As Variant
    Dim blnStop As Boolean
    Dim rsPoint As New ADODB.Recordset
    
    Dim X1 As Long
    Dim Y1 As Long
    Dim dblHeight As Double         '40-42��֮�����Ч��ӡ�߶�
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    DrawBodyGraph = False
    If IsEmpty(strArrItemInfo) Or IsEmpty(strArrValue) Then Exit Function
    If UBound(strArrItemInfo) <> UBound(strArrValue) Then Exit Function
    lngTmpY = Y
    lngTmpX = X
    objDraw.DrawStyle = 0
    W_9pt = objDraw.TextWidth("��")
'    H_9pt = objDraw.TextHeight("��")
    H_9pt = ROWHEIGHT * 10 / 3
    
    intItemCount = UBound(strArrValue) + 1
    'ReDim lngArrXY(7 * 6 - 1, 4) '��¼����
    ReDim lngArrXY(7 * 6 - 1, 5) '��¼����      ,�������һά,���ڼ�¼�Ƿ�Ϊ����
    
    Dim mpt����() As POINTAPI
    Dim mpt����() As POINTAPI
    Dim mpt����() As POINTAPI
    
    ReDim mpt����(0 To 41)
    ReDim mpt����(0 To 41)
    ReDim mpt����(0 To 41)
    
    Dim lngX As Long
    Dim lngY As Long
    Dim lngY1 As Long
    Dim strTmp As String
    Dim intPointCount As Integer
    Dim lngYMax As Long
    Dim intCharNumber As Integer
    Dim blnPrint As Boolean
    
    blnPrint = (Val(zlDatabase.GetPara("����ӡ�������ͼ��", glngSys, 1255, "0")) = 0)
    
    lngYMax = (lngTmpY + 3 * H_9pt / 2) + 44 * 3 * H_9pt / 2 - H_9pt / 2
    Call PointInit(rsPoint)
    
    '��ӡ���ת
            
    Set rsTmp = New ADODB.Recordset
    '20090926:������40-42�ȼ��ӡ,����һ����Ϣ�����������С����,�ж�����Ϣ���Ӻ���һ���ӡ,��������һ���ֱ��ȫ����ӡ
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "�к�", adDouble, 30
    rsTmp.Fields.Append "ʱ��", adVarChar, 30
    rsTmp.Fields.Append "���", adVarChar, 50
    '20090926--
    rsTmp.Fields.Append "����", adVarChar, 50       '��¼�����ת,������Ժ,����δ��˵��,�ϱ�˵��
    rsTmp.Fields.Append "��ӡ��", adVarChar, 30
    rsTmp.Fields.Append "����", adVarChar, 30
    rsTmp.Fields.Append "�߶�", adVarChar, 30       'δ��˵�����ϱ�˵�����ùܸ߶�
    rsTmp.Fields.Append "�����С", adVarChar, 50
    '----------
    rsTmp.Open
    
    For lngCol = 0 To 41
    
        lngX = lngTmpX + lngCol * HOUR_STEP_Twips + HOUR_STEP_Twips / 2
        lngY = (lngTmpY + 3 * H_9pt / 2)
        lngY1 = (lngTmpY + 3 * H_9pt / 2) + 35 * 3 * H_9pt / 2                          '�õ�35�ȵ�����
        dblHeight = (lngTmpY + 3 * H_9pt / 2) + 10 * 3 * H_9pt / 2 - H_9pt / 2          '�õ�40�ȵ�����,�����ǰ�������������
        dblHeight = dblHeight - lngY
        
        If strArrValueInOut(lngCol) <> "" Then
        
            strComment = strArrValueInOut(lngCol)
            aryTmp = Split(strComment, ";")
            
            For i = 0 To UBound(aryTmp)
                If Trim(aryTmp(i)) <> "" Then
                    rsTmp.AddNew
                    rsTmp.Fields("����").Value = 1                              '��ӡ��δ�������ת�����ȵ�����(����δ��˵��,�ϱ�˵��)
                    rsTmp.Fields("����").Value = lngX & ";" & lngY
                    rsTmp.Fields("�к�").Value = lngCol
                    rsTmp.Fields("ʱ��").Value = Split(aryTmp(i), "'")(1)
                    rsTmp.Fields("���").Value = Split(aryTmp(i), "'")(0)
                End If
            Next
            rsTmp.Sort = "ʱ��"
        End If
        
        '˵��
        strComment = Split(strArrValueComment(0), "'")(lngCol)
        If strComment <> "" Then
            aryTmp = Split(strComment, ";")
            
            '�ϱ�
            If aryTmp(0) <> "" Then
                rsTmp.AddNew
                rsTmp.Fields("����").Value = 2                              'ָδ��˵��,�ϱ�˵��
                rsTmp.Fields("����").Value = lngX & ";" & lngY
                rsTmp.Fields("�к�").Value = lngCol
                rsTmp.Fields("ʱ��").Value = ""
                rsTmp.Fields("���").Value = SimplifyString(aryTmp(0))
            End If
            
            '�±�
            intCharNumber = 0
            If aryTmp(1) <> "" Then
                For i = 1 To Len(aryTmp(1))
                    If lngY1 < lngYMax Then
                        strTmp = Mid(aryTmp(1), i, 1)
                        
                        If Asc(strTmp) < 0 Then
                            If intCharNumber Mod 2 = 1 Then lngY = lngY + ROWHEIGHT * 2.5
                        End If
                    
                        Call DrawRotateText(objDraw, lngX - objDraw.TextWidth(strTmp) / 2 + 15, lngY1 + 15, strTmp, vbRed)

                        If Asc(strTmp) < 0 Then
                            lngY1 = lngY1 + ROWHEIGHT * 5
                            intCharNumber = 0
                        Else
                            lngY1 = lngY1 + ROWHEIGHT * 2.5
                            intCharNumber = intCharNumber + 1
                        End If
                        
                    End If
                Next
            End If
        End If
    Next
    
    '������ת����,�Լ�δ��˵��,�ϱ�˵��
    Call OutputNote(objDraw, dblHeight, rsTmp, lngTmpX, lngTmpY)
    
    Dim int���� As Integer
    Dim int���� As Integer
    
    If Trim(strIndex) = "" Then
    
        '����Ŀ����ѭ��
        
        For intItemDrawCount = 0 To intItemCount - 1
        
            intPointCount = -1
            lngTopRow = CInt(Split(strArrItemInfo(intItemDrawCount), "'")(1))
            lngStep = Val(Split(strArrItemInfo(intItemDrawCount), "'")(5))
            
            lngMax = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(3))
            lngMin = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(4))

            lng��Ŀ��� = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(7))
            Select Case lng��Ŀ���
            Case 1
                int���� = intItemDrawCount
            Case 2
                int���� = intItemDrawCount
            End Select
            objDraw.ForeColor = lngColor
            
            'Erase lngArrXY
            '��ʼ���Ա��ڵ�һ��ʱ����������
            
            lngUpIndex = -1
            
            'ѭ��������
            For lngCol = 0 To 7 * 6 - 1 '��
                
                lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
                strTag = Trim(Split(strArrItemInfo(intItemDrawCount), "'")(6))

                X = lngTmpX + (lngCol + 1) * HOUR_STEP_Twips - HOUR_STEP_Twips / 2
                    
                '1�������ǰֵ����ֵ��(��ǰֵ/��λֵ+�����)*(H_9pt+H_9pt\2)������ʱΪ���Y����λ��ֵ
                '   ������ʱ�����걣���ڱ�����
                If Val(Split(strArrValue(intItemDrawCount), "'")(lngCol)) = -2 Then
                    Y = lngTmpY
                    Y2 = lngTmpY
                    lngArrXY(lngCol, 0) = -1
                    lngArrXY(lngCol, 1) = X
                    lngArrXY(lngCol, 2) = Y
                    lngArrXY(lngCol, 3) = Y2
                    lngArrXY(lngCol, 4) = -1
                    lngArrXY(lngCol, 5) = 0
                Else
                    Y = lngTmpY
                    Y2 = lngTmpY
                    
                    aryData = Split(Split(strArrValue(intItemDrawCount), "'")(lngCol), ",")
                    
                    For i = 0 To UBound(aryData)
                        lngValue = (lngTopRow - 1) * lngStep + lngMax
                        
                        '������ǰ��Ŀ����Сֵ���ֵʱ,ֱ��ȡ��Сֵ�����ֵ
                        lngArrXY(lngCol, 5) = IIf(Left(aryData(i), InStr(aryData(i), ";") - 1) = "����", 1, 0)
                        
                        Select Case Val(Left(aryData(i), InStr(aryData(i), ";") - 1))
                        Case Is < lngMin
                            aryData(i) = lngMin & ";" & Mid(aryData(i), InStr(aryData(i), ";") + 1)
                        Case Is > lngMax
                            aryData(i) = lngMax & ";" & Mid(aryData(i), InStr(aryData(i), ";") + 1)
                        End Select
                        
                        If InStr(aryData(i), ";") > 0 Then
                            lngValue1 = lngValue - Val(Left(aryData(i), InStr(aryData(i), ";") - 1))
                        Else
                            lngValue1 = lngValue - Val(aryData(i))
                        End If

                        If i = 0 Then
                            Y = lngTmpY + (lngValue1 / lngStep) * (H_9pt + H_9pt \ 2)
                            lngArrXY(lngCol, 0) = 2
                        Else
                            Y2 = lngTmpY + (lngValue1 / lngStep) * (H_9pt + H_9pt \ 2)
                            lngArrXY(lngCol, 4) = 2
                        End If
                        
                        '��Ϊ��3�����ܴ�����ı���˵��������Ŀ�Ĳ�λ��
                        If i = 1 Then Exit For
                    Next
'                    lngArrXY(lngCol, 0) = 2
                    lngArrXY(lngCol, 1) = X
                    lngArrXY(lngCol, 2) = Y
                    lngArrXY(lngCol, 3) = Y2
                    
                    '2����X-W_9pt\2,Y-H_9pt\2 ����һ����¼��
                    lngX = X - W_9pt \ 3
                    lngY = Y - H_9pt \ 3

                    lngX = X
                    lngY = Y
                    
                    Select Case lng��Ŀ���
                    Case 1
                        If InStr(aryData(0), ";") > 0 Then
                            aryTmp = Split(aryData(0), ";")
                            Select Case aryTmp(5)
                            Case "����"
                                strTag = mstrChar(0)
                            Case "Ҹ��"
                                strTag = mstrChar(1)
                            Case "����"
                                strTag = mstrChar(2)
                            Case Else
                                strTag = mstrChar(1)
                            End Select
                            
                            X1 = lngX
                            Y1 = lngY
                            Select Case Val(aryTmp(0))
                            Case Is <= lngMin
                                strTag = ""
                                lngArrXY(lngCol, 5) = 1
                                Call DrawText(objDraw, X1 - 90, Y1, "��")
                                Call DrawText(objDraw, X1 - 90, Y1 + 180, "��")
'                                strTag = "��"
'                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                            Case Is >= lngMax
                                strTag = "��"
                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                            End Select
                            
                            If Val(aryTmp(2)) = 1 Then
                                '���Ժϸ�
                                Call DrawText(objDraw, X1 - 50, Y1 - 250, "��", vbRed)
                            End If
                            
                            mpt����(lngCol).X = X1
                            mpt����(lngCol).Y = Y1
            
                        End If
                    Case 2
                        '���Ϊ�ձ�ʾȱʡ�ַ�;����Ϊ��,�ڻ�ͼʱ����Ϊ��,��ȡ��Դ�ļ��е�λͼ
                        aryTmp = Split(aryData(0), ";")
                        strTag = IIf(aryTmp(5) = "����", "", mstrPulse)
                        
                        mpt����(lngCol).X = lngX
                        mpt����(lngCol).Y = lngY
                        
                        X1 = lngX
                        Y1 = lngY
                        If InStr(aryData(0), ";") > 0 Then
                            aryTmp = Split(aryData(0), ";")
                            Select Case Val(aryTmp(0))
                            Case Is <= lngMin
                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                            Case Is >= lngMax
                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                            End Select
                        End If
                        
                    Case -1
                        mpt����(lngCol).X = lngX
                        mpt����(lngCol).Y = lngY
                        
                        X1 = lngX
                        Y1 = lngY
                        
                        If InStr(aryData(0), ";") > 0 Then
                            aryTmp = Split(aryData(0), ";")
                            Select Case Val(aryTmp(0))
                            Case Is <= lngMin
                                Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                            Case Is >= lngMax
                                Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                            End Select
                        End If
                        
                    Case Else
                        If lng��Ŀ��� = 3 Then
                            aryTmp = Split(aryData(0), ";")
                            strTag = IIf(aryTmp(5) = "������", "", mstrBreath)
                        End If
                        
                        X1 = lngX
                        Y1 = lngY
                        Select Case Val(aryData(0))
                        Case Is <= lngMin
                            Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                        Case Is >= lngMax
                            Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                        End Select
                        
                    End Select
                    
                    If lng��Ŀ��� = 1 Then         '����
                        Call PointAdd(rsPoint, lngX, lngY, lng��Ŀ���, strTag, lngColor, lngCol, CStr(aryTmp(5)))
                    ElseIf lng��Ŀ��� = 3 Then     '����
                        Call PointAdd(rsPoint, lngX, lngY, lng��Ŀ���, strTag, lngColor, lngCol, CStr(aryTmp(5)), IIf(strTag = "", "BREATH", ""))
                    ElseIf lng��Ŀ��� = 2 Then     '����
                        Call PointAdd(rsPoint, lngX, lngY, lng��Ŀ���, strTag, lngColor, lngCol, CStr(aryTmp(5)), IIf(strTag = "", "PACEMAKER", ""))
                    Else
                        Call PointAdd(rsPoint, lngX, lngY, lng��Ŀ���, strTag, lngColor, lngCol, "")
                    End If

                    lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
                    
                    If Y2 <> lngTmpY Then
                        
                        Select Case lng��Ŀ���
                        Case 2
                        
                            '��������������ʹ��ʱ
'                            If mint����Ӧ�� = 2 Then
                            
                                strTag = mstr���ʷ���
                                mpt����(lngCol).X = lngX
                                mpt����(lngCol).Y = (Y2 - H_9pt \ 3)
                                lngColor = 255

                                If InStr(aryData(1), ";") > 0 Then
                                    X1 = lngX
                                    Y1 = Y2
                                    aryTmp = Split(aryData(1), ";")
                                    Select Case Val(aryTmp(0))
                                    Case Is <= lngMin
                                        Call DrawLine(objDraw, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= lngMax
                                        Call DrawLine(objDraw, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                End If
                                
'                            End If
                        Case 1
                            strTag = "��"
                            lngColor = 255
                        End Select
                                                                        
                        If lng��Ŀ��� = 1 Then
                            Call PointAdd(rsPoint, lngX, Y2 - H_9pt \ 3, lng��Ŀ���, strTag, lngColor, lngCol, CStr(aryTmp(5)))
                        Else
                            Call PointAdd(rsPoint, lngX, Y2 - H_9pt \ 3, lng��Ŀ���, strTag, lngColor, lngCol, "")
                        End If
                    
                    End If

                    '3���ҵ�ǰһ������Ľ������ߡ������һ������ģ�0��Ԫ��Ϊ0ʱ��ʾ����
                    '--------------------------------------------------------------------------------------------------
                    If lngCol > 0 Then
                        lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
                        If lngUpIndex > -1 Then
                            If lngArrXY(lngUpIndex, 0) = 2 Then
                                If lngArrXY(lngUpIndex, 5) <> 1 And lngArrXY(lngCol, 5) <> 1 Then
                                    Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 2), X, Y, lngColor)
                                End If
                                
                                If Y2 <> lngTmpY Then
                                    Select Case lng��Ŀ���
                                    Case 1          '������Ŀ
                                        Call DrawLine(objDraw, X, Y, X, Y2, 255, IIf(Y2 < Y, 0, 2), , True)
                                    Case 2
                                        If lngArrXY(lngUpIndex, 4) = 2 Then
                                            Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 3), X, Y2, lngColor)
                                        End If
                                    End Select
                                    
                                End If
                            Else
                                '��Ȼ�ϵ㣬��������������㣬���Ȼ�����
                                '--------------------------------------------------------------------------------------
                                If Y2 <> lngTmpY Then
                                    Select Case lng��Ŀ���
                                    Case 1
                                        Call DrawLine(objDraw, X, Y, X, Y2, 255, IIf(Y2 < Y, 0, 2), , True)
                                    End Select
                                End If
                                
                                '�ҳ����켰������һ����Ч��,ֻ�и��������ݲŲ�����
                                '--------------------------------------------------------------------------------------
                                lngUpIndex = lngUpIndex - 1
                                blnStop = False
                                Do While lngUpIndex >= (lngCol \ 6) * 6 - 6 And lngUpIndex >= 0
                                    'If lngCol = 9 Then Stop
                                    If blnStop = False Then
                                        If Split(strArrValueComment(0), "'")(lngUpIndex + 1) <> "" Then
                                            aryTmp = Split(Split(strArrValueComment(0), "'")(lngUpIndex + 1), ";")
                                        Else
                                            aryTmp = Split(";;", ";")
                                        End If
                                        blnStop = (Val(aryTmp(2)) = 1)
                                    End If
                                    If lngArrXY(lngUpIndex, 5) = 1 Or lngArrXY(lngCol, 5) = 1 Then Exit Do
                                    If blnStop Then
                                        Exit Do
                                    End If

                                    If lngArrXY(lngUpIndex, 0) = 2 And lngArrXY(lngUpIndex, 5) <> 1 And lngArrXY(lngCol, 5) <> 1 Then
                                        If blnStop = False Then
                                            blnStop = False
                                            lngColor = CLng(Split(strArrItemInfo(intItemDrawCount), "'")(0))
                                            
                                            Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 2), X, Y, lngColor)
                                            
                                            If Y2 <> lngTmpY Then
                                                Select Case lng��Ŀ���
                                                Case 1
                                                    Call DrawLine(objDraw, X, Y, X, Y2, 255, IIf(Y2 < Y, 0, 2), , True)
                                                Case 2                  '���������������
                                                    If lngArrXY(lngUpIndex, 4) = 2 Then
                                                        Call DrawLine(objDraw, lngArrXY(lngUpIndex, 1), lngArrXY(lngUpIndex, 3), X, Y2, lngColor)
                                                    End If
                                                End Select
                                            End If
                                        End If
                                        Exit Do
                                    End If
                                    lngUpIndex = lngUpIndex - 1
                                Loop
                                
                                '--------------------------------------------------------------------------------------
                                
                            End If
                        End If
                    ElseIf Y2 <> lngTmpY Then
                        Select Case lng��Ŀ���
                        Case 1
                            Call DrawLine(objDraw, X, Y, X, Y2, 255, IIf(Y2 < Y, 0, 2), , True)
                        End Select
                    End If
                End If
                lngUpIndex = lngCol
            Next
        Next
        
        '�������������������γɶ���Σ����������ߺ����
        '--------------------------------------------------------------------------------------------------------------
        Call DrawPoly(objDraw, mpt����, mpt����, mbyt����)
    End If
    
    '������ַ���ͼ��
    '--------------------------------------------------------------------------------------------------------------
    Call DrawPoint(objDraw, rsPoint, int����)
    
    X = lngTmpX - (UBound(strArrItemInfo) + 1) * (W_9pt + W_9pt \ 2)
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function InitOffset(ByRef rsOffset As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Set rsOffset = New ADODB.Recordset
    With rsOffset
        .Fields.Append "Line", adInteger
        .Fields.Append "Column", adInteger
        .Fields.Append "Offset", adVarChar, 30
        .Open
    End With
    
    InitOffset = True
    
End Function


Public Function IsCenterValue(ByRef rsOffset As ADODB.Recordset, ByVal lngLine As Long, ByVal lngColumn As Long, ByVal dtDot As Date, ByVal dtCenter As Date) As Boolean
    '******************************************************************************************************************
    '���ܣ����ͬһ�����ж��ֵ����ȡ��е��ֵ��Ϊ���е�ֵ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngSecond As Long
        
    On Error GoTo errHand
    
    IsCenterValue = False
        
    lngSecond = Abs(DateDiff("s", dtCenter, dtDot))
    
    rsOffset.Filter = ""
    rsOffset.Filter = "Line=" & lngLine & " And Column=" & lngColumn
    
    If rsOffset.RecordCount = 0 Then
    
        rsOffset.AddNew
        rsOffset("Line").Value = lngLine
        rsOffset("Column").Value = lngColumn
        rsOffset("Offset").Value = lngSecond
        
        IsCenterValue = True
        
    ElseIf Val(rsOffset("Offset").Value) >= lngSecond Then
    
        rsOffset("Offset").Value = lngSecond
        IsCenterValue = True
        
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
End Function

Public Function PointInit(ByRef rsPoint As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Set rsPoint = New ADODB.Recordset
    With rsPoint
        .Fields.Append "X", adVarChar, 30
        .Fields.Append "Y", adVarChar, 30
        .Fields.Append "Line", adVarChar, 30
        .Fields.Append "Column", adInteger
        .Fields.Append "����", adVarChar, 30
        .Fields.Append "��ɫ", adVarChar, 30
        .Fields.Append "��־", adInteger
        .Fields.Append "ͼ��", adVarChar, 200
        .Fields.Append "�ص���ʶ", adVarChar, 50
        .Fields.Append "��Ŀ���", adBigInt
        .Fields.Append "���²�λ", adVarChar, 30
        .Open
    End With
    
    PointInit = True
    
End Function

Public Function PointAdd(ByRef rsPoint As ADODB.Recordset, ByVal X1 As Single, ByVal Y1 As Single, ByVal lngLine As Long, ByVal strChar As String, ByVal lngColor As Long, ByVal lngColumn As Long, ByVal str���²�λ As String, Optional ByVal strͼ�� As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    With rsPoint
        .AddNew
        .Fields("X").Value = X1
        .Fields("Y").Value = Y1
        .Fields("Line").Value = lngLine
        .Fields("Column").Value = lngColumn
        .Fields("����").Value = strChar
        .Fields("��ɫ").Value = lngColor
        .Fields("��־").Value = 0
        .Fields("�ص���ʶ").Value = ""
        .Fields("ͼ��").Value = strͼ��
        If str���²�λ = "" And lngLine = 1 Then str���²�λ = "Ҹ��"
        .Fields("���²�λ").Value = str���²�λ
    End With
    
    PointAdd = True
End Function

Private Function PointCalc(ByRef rsPoint As ADODB.Recordset, ByVal int���� As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡ�ص�������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strTmp As String
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsConverPoint As ADODB.Recordset
    
    On Error GoTo errHand
    
    Call GetConverPoint(rsPoint, rsConverPoint, int����)
    
    If Not rsConverPoint Is Nothing Then
        If rsConverPoint.RecordCount > 0 Then
            rsConverPoint.MoveFirst
            Do While Not rsConverPoint.EOF
                
                strTmp = rsConverPoint("Lines").Value
                If InStr("," & strTmp & ",", ",1,") > 0 Then

                    strTmp = "0," & strTmp & ",0"
                    strTmp = Replace(strTmp, ",1,", ",")
                    
                    strSQL = "Select a.���,a.��Ƿ���,a.�����ɫ " & _
                            "From �����ص���� a,(Select �ϼ����, Count(1) As ���� From �����ص���� Where ��Ŀ��� In (" & strTmp & ") Or (��Ŀ���=1 And Nvl(���²�λ,'Ҹ��')=[2]) Group By �ϼ����) b " & _
                            "Where a.��� = b.�ϼ���� And b.���� = a.�ص���Ŀ And a.�ص���Ŀ=[1]"
                    
                Else
                
                    strSQL = "Select a.���,a.��Ƿ���,a.�����ɫ " & _
                            "From �����ص���� a,(Select �ϼ����, Count(1) As ���� From �����ص���� Where ��Ŀ��� In (" & strTmp & ") Group By �ϼ����) b " & _
                            "Where a.��� = b.�ϼ���� And b.���� = a.�ص���Ŀ And a.�ص���Ŀ=[1]"
                        
                End If
                Set rs = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", Val(rsConverPoint("�ص���Ŀ").Value), CStr(rsConverPoint("���²�λ").Value))

                If rs.BOF = False Then
                    rsPoint.Filter = ""
                    rsPoint.Filter = "�ص���ʶ='" & rsConverPoint("�ص���ʶ").Value & "'"
                    If rsPoint.RecordCount > 0 Then
                        rsPoint.MoveFirst
                        rsPoint("��־").Value = 1
                        rsPoint("����").Value = zlCommFun.NVL(rs("��Ƿ���").Value)
                        rsPoint("��ɫ").Value = zlCommFun.NVL(rs("�����ɫ").Value, 0)
                        
                        '��ȡ���ͼ�β���ʾ
                        strTmp = App.Path & "\ConverPoint" & zlCommFun.NVL(rs("���").Value, 0) & ".tmp"
                        If Dir(strTmp) <> "" Then Kill strTmp
                        strTmp = zlBlobRead(9, zlCommFun.NVL(rs("���").Value, 0), strTmp)
                        If Dir(strTmp) <> "" And strTmp <> "" Then
                            
                            rsPoint("ͼ��").Value = strTmp
                            
                        End If

                    End If
                End If
                rsConverPoint.MoveNext
            Loop
        End If
    End If
    
    PointCalc = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function DrawPoint(objDraw As Object, ByRef rsPoint As ADODB.Recordset, ByVal int���� As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ����ص����ַ���ͼ��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strChar As String
    Dim lngColor As Long
    Dim X1 As Long
    Dim Y1 As Long
    Dim strTmp As String
    
    On Error GoTo errHand
    
    '�����ص�
    Call PointCalc(rsPoint, int����)
    
    '�����
    rsPoint.Filter = ""
    If rsPoint.RecordCount > 0 Then
        rsPoint.MoveFirst

        Do While Not rsPoint.EOF

            strChar = rsPoint("����").Value
            lngColor = Val(rsPoint("��ɫ").Value)

            X1 = Val(rsPoint("X").Value) - objDraw.TextWidth(strChar) / 2
            Y1 = Val(rsPoint("Y").Value) - objDraw.TextHeight(strChar) / 2

            If X1 > 0 Or Y1 > 0 Then
                Select Case rsPoint("��־").Value
                Case 0                              '�����ĵ�
                    '�����ĵ�Ҳ������ͼ�Σ������������������������ʱΪͼ�Σ�����ͼ��������Դ�ļ���ģ������赥������
                    If rsPoint!ͼ�� <> "" Then  '���������Դ�ļ��е�ID
                        X1 = Val(rsPoint("X").Value)
                        Y1 = Val(rsPoint("Y").Value)
            
                        Call DrawPicture(objDraw, rsPoint!ͼ��, X1 - 90, Y1 - 90, X1 + 90, Y1 + 90, True)
                    Else
                        Call DrawText(objDraw, X1, Y1, strChar, lngColor)
                    End If
                Case 1                              '�ص��ĵ�
                    
                    If strChar <> "" Then
                        Call DrawText(objDraw, X1, Y1, strChar, lngColor)
                    Else
                        strTmp = rsPoint("ͼ��").Value
                        If strTmp <> "" And Dir(strTmp) <> "" Then
                            
                            X1 = Val(rsPoint("X").Value)
                            Y1 = Val(rsPoint("Y").Value)
                
                            Call DrawPicture(objDraw, strTmp, X1 - 90, Y1 - 90, X1 + 90, Y1 + 90)
                        End If
                    End If
                    
                End Select
            End If

            rsPoint.MoveNext
        Loop
    End If
    
    DrawPoint = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

'Public Function DrawPoly(objDraw As Object, pt����() As POINTAPI, pt����() As POINTAPI) As Boolean
'    '******************************************************************************************************************
'    '���ܣ��������������������γɶ���Σ����������ߺ����
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim intCount As Integer
'    Dim intStart As Integer
'    Dim intEnd As Integer
'
'    For intCount = 0 To 41
'        If pt����(intCount).X > 0 Then
'            If pt����(intCount).X = 0 Then
'                If intEnd > 0 Then
'                    intEnd = intCount
'                    Call DrawFillPoly(objDraw, intStart, intEnd, pt����, pt����)
'                    intEnd = 0
'                    intStart = intCount
'                Else
'                    intStart = intCount
'                End If
'            Else
'                intEnd = intCount
'                If intStart = 0 Then intStart = intCount
'            End If
'        Else
'            If intEnd > 0 And intStart > 0 Then
'                Call DrawFillPoly(objDraw, intStart, intEnd, pt����, pt����)
'            End If
'            intStart = 0
'            intEnd = 0
'        End If
'    Next
'    If intEnd > 0 Then
'        '��һ���������
'        Call DrawFillPoly(objDraw, intStart, intEnd, pt����, pt����)
'    End If
'
'    DrawPoly = True
'
'End Function

Public Function DrawPoly(objDraw As Object, pt����() As POINTAPI, pt����() As POINTAPI, byt����() As Byte) As Boolean
    '******************************************************************************************************************
    '���ܣ��������������������γɶ���Σ����������ߺ����
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intCount As Integer
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim intDay As Integer
    
    intStart = -1: intEnd = -1
    For intCount = 0 To 41
        If pt����(intCount).X > 0 Then
            If pt����(intCount).X = 0 Then
                If intEnd >= 0 Then
                    intEnd = intCount
                    Call DrawFillPoly(objDraw, intStart, intEnd, pt����, pt����)
                    intEnd = -1
                    intStart = intCount
                Else
                    intStart = intCount
                End If
            Else
                intEnd = intCount
                If intStart = -1 Then intStart = intCount
            End If
        Else
            '���û�����ݼ�� �Ƿ���δ��˵������������������֮�����ݳ���һ��
            '��δ��˵������������
            '���ݳ���һ�컭��������

            intDay = intCount - (IIf(intEnd = -1, intStart, intEnd) + ((intCount + 1) Mod 6))

            If byt����(intCount) = 1 Or intDay > 6 Then
                If intEnd >= 0 And intStart >= 0 Then
                    Call DrawFillPoly(objDraw, intStart, intEnd, pt����, pt����)
                End If
                intStart = -1
                intEnd = -1
            End If
        End If
    Next
    If intEnd >= 0 Then
        '��һ���������
        Call DrawFillPoly(objDraw, intStart, intEnd, pt����, pt����)
    End If

    DrawPoly = True

End Function

Private Function GetConverPoint(ByVal rsPoint As ADODB.Recordset, ByRef rsConverPoint As ADODB.Recordset, ByVal int���� As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim X0 As Long
    Dim Y0 As Long
    Dim lngLine As Long
    Dim intMax As Integer

    On Error GoTo errHand

    If rsPoint.RecordCount = 0 Then Exit Function

    Set rsConverPoint = New ADODB.Recordset
    With rsConverPoint
        .Fields.Append "�ص���ʶ", adVarChar, 30
        .Fields.Append "Lines", adVarChar, 30
        .Fields.Append "�ص���Ŀ", adInteger
        .Fields.Append "���²�λ", adVarChar, 30
        .Open
    End With


    '------------------------------------------------------------------------------------------------------------------
    rsPoint.Sort = "X,Y"
    rsPoint.MoveFirst
    Do While Not rsPoint.EOF
        If rsPoint("X").Value = X0 And rsPoint("Y").Value = Y0 Then
            rsConverPoint.Filter = ""
            rsConverPoint.Filter = "�ص���ʶ='" & X0 & "," & Y0 & "'"
            If rsConverPoint.RecordCount = 0 Then
                rsConverPoint.AddNew
                rsConverPoint("�ص���ʶ").Value = X0 & "," & Y0
                rsConverPoint("Lines").Value = ""
                rsConverPoint("�ص���Ŀ").Value = 0
            End If
            If rsConverPoint("Lines").Value = "" Then

                rsPoint.MovePrevious
                rsPoint("�ص���ʶ").Value = X0 & "," & Y0
                rsPoint.MoveNext
                rsConverPoint("�ص���Ŀ").Value = 2
                rsConverPoint("Lines").Value = lngLine & "," & rsPoint("Line").Value
            Else
                rsConverPoint("Lines").Value = rsConverPoint("Lines").Value & "," & rsPoint("Line").Value
                rsConverPoint("�ص���Ŀ").Value = rsConverPoint("�ص���Ŀ").Value + 1
            End If

            If rsPoint("���²�λ").Value <> "" Then
                'Ŀǰֻ�����µĲ�λ�ص�����,���Դ˴��Ա���ԭ��
                If InStr(1, rsPoint("���²�λ").Value, ";") <> 0 Then
                    '��ӡ����ȡ�Ĳ�λ�Ǹ������ߵ�,û�кϲ���һ��
                    rsConverPoint("���²�λ").Value = Split(rsPoint("���²�λ").Value, ";")(int���� + 1)
                Else
                    rsConverPoint("���²�λ").Value = rsPoint("���²�λ").Value
                End If
            End If

            rsPoint("�ص���ʶ").Value = X0 & "," & Y0
            rsPoint("��־").Value = 2

            GetConverPoint = True

        End If

        lngLine = rsPoint("Line").Value
        X0 = rsPoint("X").Value
        Y0 = rsPoint("Y").Value

        rsPoint.MoveNext
    Loop

    rsConverPoint.Filter = ""

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
'
'Private Function GetConverPoint(ByVal rsPoint As ADODB.Recordset, ByRef rsConverPoint As ADODB.Recordset) As Boolean
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim X0 As Long
'    Dim Y0 As Long
'    Dim Y1 As Long          '��ʱʹ��,�����ж��Ƿ��غ�
'    Dim lngLine As Long
'    Const dblError As Long = 15  '���ֵ,��ʾ����+/-15�����
'
'    On Error GoTo errHand
'    If rsPoint.RecordCount = 0 Then Exit Function
'
'    Set rsConverPoint = New ADODB.Recordset
'    With rsConverPoint
'        .Fields.Append "���", adVarChar, 30            'ÿһ����ű�ʾһ���غϵĵ�
'        .Fields.Append "Lines", adVarChar, 30           '��ǰ�Ǳ��������غϵ��������,�ָ�Ϊ���浱ǰ������
'        .Fields.Append "����", adInteger                '����ָ����Χ�����,�Դ˽��к������ж�
'        .Fields.Append "ʵ������", adInteger            'ʵ������
'        .Fields.Append "���²�λ", adVarChar, 30        '���²�λ/������ʽ/����
'        .Open
'    End With
'
'
'    '------------------------------------------------------------------------------------------------------------------
'    rsPoint.Sort = "X,Y"
'    rsPoint.MoveFirst
'    Do While Not rsPoint.EOF
'
'        '����趨Ϊ30
'        If rsPoint("X").Value = X0 And Abs(rsPoint("Y").Value - Y0) <= dblError Then
'            rsConverPoint.Filter = ""
'            Y1 = IIf(rsPoint!y >= Y0, Y0, rsPoint!y) + dblError
'            rsConverPoint.Filter = "����='" & X0 & "," & Y0 & "'"
'            If rsConverPoint.RecordCount = 0 Then
'                rsConverPoint.AddNew
'                rsConverPoint("�ص���ʶ").Value = X0 & "," & Y0
'                rsConverPoint("Lines").Value = ""
'                rsConverPoint("�ص���Ŀ").Value = 0
'            End If
'            If rsConverPoint("Lines").Value = "" Then
'
'                rsPoint.MovePrevious
'                rsPoint("�ص���ʶ").Value = X0 & "," & Y0
'                rsPoint.MoveNext
'                rsConverPoint("�ص���Ŀ").Value = 2
'                rsConverPoint("Lines").Value = lngLine & "," & rsPoint("Line").Value
'            Else
'                rsConverPoint("Lines").Value = rsConverPoint("Lines").Value & "," & rsPoint("Line").Value
'                rsConverPoint("�ص���Ŀ").Value = rsConverPoint("�ص���Ŀ").Value + 1
'            End If
'
'            If rsPoint("���²�λ").Value <> "" Then
'                rsConverPoint("���²�λ").Value = rsPoint("���²�λ").Value
'            End If
'
'            rsPoint("�ص���ʶ").Value = X0 & "," & Y0
'            rsPoint("��־").Value = 2
'
'            GetConverPoint = True
'
'        End If
'
'        lngLine = rsPoint("Line").Value
'        X0 = rsPoint("X").Value
'        Y0 = rsPoint("Y").Value
'
'        rsPoint.MoveNext
'    Loop
'
'    rsConverPoint.Filter = ""
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'End Function

Public Function GetDataFromHis(ByVal lng����id As Long, ByVal lng��ҳid As Long, ByVal lngӤ�� As Long, ByVal dtFrom As Date, ByVal dtTo As Date, Optional ByVal bytMode As Byte = 1) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ���ҽ����¼��ȡ��������������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    Dim lng������Ŀid As Long
    Dim rs As New ADODB.Recordset
    
    
    Select Case bytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1              '��ҽ����¼��ȡ��������������
        
'        dtFrom = dtFrom - 14
        
        strSQL = _
                "Select ִ��ʱ��,����,�ε�" & vbNewLine & _
                "From (Select ִ��ʱ��,����, Rownum As �ε�" & vbNewLine & _
                "       From (Select Distinct C.ִ��ʱ��,'����' As ���� " & vbNewLine & _
                "              From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ��ִ�� C" & vbNewLine & _
                "              Where A.����id = [1] And A.��ҳid = [2] And Nvl(A.Ӥ��, 0) = [3] And A.ҽ����Ч = 1 And A.������Ŀid = B.ID And" & vbNewLine & _
                "                    A.������� = 'F' And A.ҽ��״̬ = 8 And C.ҽ��id = A.ID And C.ִ��ʱ�� < =[5] " & vbNewLine & _
                "               Union All Select a.����ʱ�� As ִ��ʱ��,'����' As ���� From ������������¼ a Where a.����id=[1] And a.��ҳid=[2] And a.����ʱ�� Is Not Null And RowNum<2) " & _
                "       Order By ִ��ʱ��)" & vbNewLine & _
                "Where ִ��ʱ�� >= [4] And �ε� <= 12 " & vbNewLine & _
                "Order By ִ��ʱ�� "
                
        Set GetDataFromHis = zlDatabase.OpenSQLRecord(strSQL, "���µ�", lng����id, lng��ҳid, lngӤ��, dtFrom, dtTo)
    '------------------------------------------------------------------------------------------------------------------
    Case 2              '���ת��־(��Ժ,��Ժ,ת��,����)
        
        
        '1-��Ժ��2-��ƣ�3-ת�ƣ�4-������5-��λ�ȼ��䶯��6-����ȼ��䶯��7-����ҽʦ�ı䣻8-���λ�ʿ�ı�,9-���۲���תסԺ,10-����Ԥ��Ժ,11-����ҽʦ�䶯,12-����ҽʦ�䶯,13-����䶯
        
        '�е�ҽԺ���������ж��
'        strSQL = "Select ID From ������ĿĿ¼ Where ���='Z' And ��������='11' "
'        Set rs = zlDatabase.OpenSQLRecord(strSQL, "���µ�")
'        If rs.BOF = False Then lng������Ŀid = zlCommFun.NVL(rs("ID").Value)
        
        strSQL = _
                "   Select b.���� As ����,��ʼʱ�� As ʱ��, Decode(��ʼԭ��, 2,'���',3, 'ת��',4,'����'||Decode(����,Null,'','('||����||')')) As ����,Decode(��ʼԭ��,2,9,3,6,4,7) As �к� " & vbNewLine & _
                "   From ���˱䶯��¼ A,���ű� b" & vbNewLine & _
                "   Where b.id(+)=a.����id and a.��ʼԭ�� In (2,3,4) And A.����id = [1] And A.��ҳid = [2]  And [3]=0 And A.��ʼʱ�� Between [4] And [5] " & vbNewLine & _
                "   Union All" & vbNewLine & _
                "   Select '' As ����,ʱ��,����,�к� From (Select * From (Select ��ʼʱ�� As ʱ��, '��Ժ' As ����,5 As �к� " & vbNewLine & _
                "   From ���˱䶯��¼ A" & vbNewLine & _
                "   Where a.��ʼԭ��=1 And A.����id = [1] And A.��ҳid = [2] And [3]=0 Order By a.��ʼʱ��) Where RowNum=1) Where ʱ�� Between [4] And [5] " & vbNewLine & _
                "   Union All" & vbNewLine & _
                "   Select '' As ����,Nvl(b.��ʼִ��ʱ��,a.��Ժ����) As ʱ��, Decode(��Ժ��ʽ, '����', '��Ժ', ��Ժ��ʽ) As ����,8 As �к� " & vbNewLine & _
                "   From ������ҳ A,(Select x.����id,x.��ҳid,Max(x.��ʼִ��ʱ��) As ��ʼִ��ʱ�� From ����ҽ����¼ x,������ĿĿ¼ z Where x.����id=[1] And x.��ҳid=[2] " & vbNewLine & _
                "   And x.������Ŀid+0=z.ID And x.ҽ��״̬ in (3,8) And z.���='Z' And z.��������='11' Group By x.����id,x.��ҳid) B " & vbNewLine & _
                "   Where A.����id = [1] And A.��ҳid = [2] And A.��Ժ���� Between [4] And [5] And a.����id=b.����id(+) And a.��ҳid=b.��ҳid(+) "
                
        
        strSQL = "Select distinct * From (" & strSQL & ") Order By ʱ��,�к� "
        Set GetDataFromHis = zlDatabase.OpenSQLRecord(strSQL, "���µ�", lng����id, lng��ҳid, lngӤ��, dtFrom, dtTo)
    '------------------------------------------------------------------------------------------------------------------
    Case 3              '����������¼�������/��������
        
        strSQL = "Select '' As ����,a.����ʱ�� As ʱ��,'����' As ����,13 As �к� From ������������¼ a Where a.����id=[1] And a.��ҳid=[2] And a.���=[3] And a.����ʱ�� Is Not Null"
        Set GetDataFromHis = zlDatabase.OpenSQLRecord(strSQL, "���µ�", lng����id, lng��ҳid, lngӤ��)
        
    End Select
    
End Function

'todo:�����������������õ��Ĺ��̻���
Public Sub DrawLine(pic As Object, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1, Optional ByVal blnEndArrow As Boolean)
    
    '��(X1,Y1),(X2,Y2)֮��ʹ��ForeColorɫ��һֱ��
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    Dim lngLoop As Long

    lngSaveForeColor = pic.ForeColor
    bytSaveLineWidth = pic.DrawWidth
    pic.ForeColor = ForeColor
    pic.DrawStyle = DrawStyle
    pic.DrawWidth = LineWidth
    pic.Line (X2, Y2)-(X1, Y1)

    If blnEndArrow Then

        If Y1 < Y2 Then
            For lngLoop = X1 - 40 To X1 + 40
                pic.Line (X2, Y2)-(lngLoop, Y2 - 75)
            Next
        Else

            For lngLoop = X1 - 40 To X1 + 40
                pic.Line (X2, Y2)-(lngLoop, Y2 + 75)
            Next

        End If
    End If

    pic.ForeColor = lngSaveForeColor
    pic.DrawWidth = bytSaveLineWidth

End Sub

Public Sub DrawText(objDraw As Object, ByVal X As Single, ByVal Y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0)
    '��(X,Y)�����Text�ı�
    Dim lngSaveForeColor As Long
    
    With objDraw
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        .CurrentX = X
        .CurrentY = Y
        objDraw.FontTransparent = True
        objDraw.Print Text
        .ForeColor = lngSaveForeColor
    End With
End Sub

Public Sub DrawRotateText(objDraw As Object, ByVal X As Single, ByVal Y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0, Optional ByVal sglScale As Single = 1.5)
    '��(X,Y)�����Text�ı�
    Dim lngSaveForeColor As Long
    Dim lngLoop As Long
    Dim objFont As New clsRotateFont '��ת�������

    With objDraw
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        objDraw.FontTransparent = True

        If Asc(Text) < 0 Then
            'ȫ��
            .CurrentX = X
            .CurrentY = Y
            
            objDraw.Print Text
        Else
            '���,����ת90��

            .CurrentX = X + objDraw.TextWidth("1") * sglScale
            .CurrentY = Y

            Set objFont = New clsRotateFont
            Set objFont.LogFont = New StdFont
            objFont.LogFont.Name = .FontName
            objFont.LogFont.Size = .FontSize
            objFont.Rotation = -90
            objFont.Output objDraw, .CurrentX, .CurrentY, Text

        End If

        .ForeColor = lngSaveForeColor
    End With
End Sub

Private Function DrawFillPoly(objDraw As Object, ByVal intStart As Integer, ByVal intEnd As Integer, pt����() As POINTAPI, pt����() As POINTAPI)
    '******************************************************************************************************************
    '
    '
    '
    '******************************************************************************************************************
    
    Dim intCol As Integer
    Dim ptPoly() As POINTAPI
    Dim intLoop1 As Integer
    Dim intLoop2 As Integer
    Dim intLoop As Integer
    Dim lngBrush As Long
    Dim lngRgn As Long
    Dim lngSvrColor As Long
    
    
    lngSvrColor = objDraw.ForeColor
    ReDim ptPoly(0)
    
    intLoop1 = 0
    intLoop2 = 0
    intLoop = 0
    For intCol = intStart To intEnd
        If pt����(intCol).X > 0 Then
            intLoop1 = intLoop1 + 1
            ReDim Preserve ptPoly(intLoop1)
            ptPoly(intLoop1).X = pt����(intCol).X / Screen.TwipsPerPixelX
            ptPoly(intLoop1).Y = pt����(intCol).Y / Screen.TwipsPerPixelY
        End If
    Next
    For intCol = intEnd To intStart Step -1
        If pt����(intCol).X > 0 Then
            intLoop2 = intLoop2 + 1
            ReDim Preserve ptPoly(intLoop2 + intLoop1)
            ptPoly(intLoop2 + intLoop1).X = pt����(intCol).X / Screen.TwipsPerPixelX
            ptPoly(intLoop2 + intLoop1).Y = pt����(intCol).Y / Screen.TwipsPerPixelY
        End If
    Next

    intLoop = intLoop2 + intLoop1 + 1
    ReDim Preserve ptPoly(intLoop)
    ptPoly(intLoop).X = ptPoly(1).X
    ptPoly(intLoop).Y = ptPoly(1).Y

    lngBrush = CreateHatchBrush(3, vbRed)
    lngRgn = CreatePolygonRgn(ptPoly(1), intLoop, ALTERNATE)
    FillRgn objDraw.hDC, lngRgn, lngBrush

    Call DeleteObject(lngRgn)
    Call DeleteObject(lngBrush)

    objDraw.ForeColor = 255
    Polyline objDraw.hDC, ptPoly(intLoop1), intLoop2 + 2

    objDraw.ForeColor = lngSvrColor
End Function

Private Sub DrawBodyPaper(objDraw As Object, X As Long, Y As Long, intGuageRow As Integer, Optional sngScale As Long = 1)
    '******************************************************************************************************************
    '����:���ڲ����ĵ���
    '����:intGuageRow=Ҫ��������
    '˳��:
    '��ͼ˳���:=5/9   ���ǻ�����ͼ�εĵ��岽
    '******************************************************************************************************************
    
    Dim intRow As Integer, intCol As Integer
    Dim H_9pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    Dim X0 As Long, Y0 As Long
    
    objDraw.Font.Name = "����"
    objDraw.Font.Size = 9 * sngScale
    objDraw.Font.Bold = False
'    H_9pt = objDraw.TextHeight("��")
    H_9pt = ROWHEIGHT * 10 / 3
    
    lngTmpX = X
    lngTmpY = Y
    
    '------------------------------------------------------------------------------------------------------------------
    '����������ͼ��
    For intRow = 0 To intGuageRow - 1
        objDraw.DrawStyle = 2
        Y0 = lngTmpY + intRow * (H_9pt + H_9pt \ 2)
        
'        If (intRow + 1) Mod 5 = 0 Then
        If intRow Mod 5 = 0 Then
            If (intRow - 1) = 24 Then
                Call DrawLine(objDraw, lngTmpX + 10, Y0 + ROWHEIGHT * 5, lngTmpX + HOUR_STEP_Twips * 6 * 7, Y0 + ROWHEIGHT * 5, RGB(200, 0, 0), , 2)
            Else
                Call DrawLine(objDraw, lngTmpX + 10, Y0 + ROWHEIGHT * 5, lngTmpX + HOUR_STEP_Twips * 6 * 7, Y0 + ROWHEIGHT * 5, COLOR.��ɫ, , 2)
            End If
            
        Else
            Call DrawLine(objDraw, lngTmpX + 10, Y0 + ROWHEIGHT * 5, lngTmpX + HOUR_STEP_Twips * 6 * 7, Y0 + ROWHEIGHT * 5, COLOR.��ɫ)
        End If
    Next
    Y = intGuageRow * (H_9pt + H_9pt \ 2) + lngTmpY  '�������µ�һ����

    '------------------------------------------------------------------------------------------------------------------
    '����������ͼ��
    For intCol = 0 To 6
        objDraw.DrawStyle = 0
        X0 = lngTmpX + intCol * HOUR_STEP_Twips * 6
        objDraw.Line (X0 + HOUR_STEP_Twips * 1, lngTmpY)-(X0 + HOUR_STEP_Twips * 1, Y), COLOR.���ɫ
        objDraw.Line (X0 + HOUR_STEP_Twips * 2, lngTmpY)-(X0 + HOUR_STEP_Twips * 2, Y), COLOR.���ɫ
        objDraw.Line (X0 + HOUR_STEP_Twips * 3, lngTmpY)-(X0 + HOUR_STEP_Twips * 3, Y), COLOR.���ɫ
        objDraw.Line (X0 + HOUR_STEP_Twips * 4, lngTmpY)-(X0 + HOUR_STEP_Twips * 4, Y), COLOR.���ɫ
        objDraw.Line (X0 + HOUR_STEP_Twips * 5, lngTmpY)-(X0 + HOUR_STEP_Twips * 5, Y), COLOR.���ɫ

        Call DrawLine(objDraw, X0 + HOUR_STEP_Twips * 6, lngTmpY, X0 + HOUR_STEP_Twips * 6, Y, vbRed, , 2)
    Next
    
    '������һ��ͼ����
    X = lngTmpX
    Y = lngTmpY
End Sub

Private Function DrawBodyScale(objDraw As Object, X As Long, Y As Long, intGuageRow As Integer, arrstrItem() As String, Optional sngScale As Long = 1)
    '******************************************************************************************************************
    '����:�������±���ĿarrstrItem()�������ĵ�ǰҳ�������Ŀ�̶�
    '����:intGuageRow=�̶�����
    '     arrstrItem=���µ���ɫ�����ȸߡ����ơ������С���������ǣ�������������Ϊ�˺�����ͼ�ķ��㣩
    'ע��:�ڵ���ʱӦȷ��intGuageRow���ڵ���1(��Ϊֻ��intGuageRow���ڵ���1�ű�ʾ�����±���Ŀ)
    
    '˳��:
    '��ͼ˳���:=3/9   ���ǻ�����ͼ�εĵ�����
    '******************************************************************************************************************
    Dim aryItem() As String
    Dim intCountItem As Integer '��¼��Ŀ����
    Dim lngColor As Long
    Dim intItemTop As Integer
    Dim strItem As String
    Dim H_9pt As Long, W_9pt As Long
    Dim i As Long, j As Long, l As Long, k As Long 'ѭ��֮��
    Dim lngTmpY As Long, lngTmpX As Long
    Dim lngTmpVal As Single
    Dim lngY As Long
    Dim strTmp As String
    Dim lngPercW As Long
    Dim intLoop As Integer
    
    '�ο��߶�
    objDraw.Font.Name = "����"
    objDraw.Font.Size = 9 * sngScale
    objDraw.Font.Bold = False
    
    W_9pt = objDraw.TextWidth("��")
    H_9pt = ROWHEIGHT * 10 / 3
    
    
    'Ϊ���˳�
    If IsEmpty(arrstrItem) = True Then Exit Function
    '���������˳�
    If UBound(arrstrItem) < 0 Then Exit Function
    '���������˳�
    If intGuageRow < 1 Then Exit Function
    
    DrawBodyScale = True
    objDraw.DrawStyle = 0
    lngTmpY = Y
    lngTmpX = X
    
    intCountItem = UBound(arrstrItem)
    
    ReDim aryItem(intCountItem)
    For i = 0 To intCountItem '��
        aryItem(i) = arrstrItem(i)
    Next
    
    For i = 0 To intCountItem '��
        If GetSplitStr(aryItem(i), 2) = "����" Then
            
            For intLoop = i To intCountItem - 1
                aryItem(intLoop) = aryItem(intLoop + 1)
            Next
            
            intCountItem = intCountItem - 1
            
            Exit For
        
        End If
    Next
    
    For i = 0 To intCountItem '��
        If GetSplitStr(aryItem(i), 2) = "����" Then
            lngPercW = W_9pt * 3
        Else
            lngPercW = W_9pt * 7
        End If
        
        Y = lngTmpY + 30
        lngColor = CLng(GetSplitStr(aryItem(i), 0))
        intItemTop = CInt(GetSplitStr(aryItem(i), 1))
        l = 0
        For j = 0 To intGuageRow '��
            '��������������˳�
            If UBound(Split(aryItem(i), "'")) < 6 Then: DrawBodyScale = False: Exit Function
            
            lngTmpVal = Val(GetSplitStr(aryItem(i), 3)) - Val(GetSplitStr(aryItem(i), 5)) * l
            
            If j = 0 Then
            
                '����Ϊ����
'                If GetSplitStr(aryItem(i), 8) <> "" Then
'                    strItem = GetSplitStr(aryItem(i), 2) & "(" & GetSplitStr(aryItem(i), 8) & ")"
'                Else
                    strItem = GetSplitStr(aryItem(i), 2)
'                End If
                    
                lngY = Y
                Call DrawText(objDraw, X + (lngPercW - objDraw.TextWidth(strItem)) / 2, lngY, strItem, lngColor)
                If strItem = "����" Then
                    Call DrawText(objDraw, X + (lngPercW - objDraw.TextWidth("   F     C   ")) / 2, lngY + 180, "   F     C   ", lngColor)
                End If
'                For k = 1 To Len(strItem)
'
'                    strTmp = Mid(strItem, k, 1)
'
'                    If strTmp = "/" Then
'                        strTmp = "\"
'                    End If
'
'                    Call DrawRotateText(objDraw, X + (lngPercW - objDraw.TextWidth(strTmp)) / 2, lngY, strTmp, lngColor, 1)
'
'                    lngY = lngY + IIf(Asc(strTmp) < 0, objDraw.TextHeight(strTmp), objDraw.TextWidth(strTmp))
'
'                Next

                Y = Y + H_9pt * 2 ' H_9pt * 3  + H_9pt \ 2 + H_9pt + H_9pt \ 2
                
            ElseIf j >= intItemTop And lngTmpVal >= CLng(GetSplitStr(aryItem(i), 4)) - IIf(strItem = "����", 1, 0) Then
                
                '���ݼ����
                If (lngTmpVal - Fix(lngTmpVal)) = 0 Then
                    
                    Select Case GetSplitStr(aryItem(i), 2)
                    Case "����", "����", "����"
                        If lngTmpVal Mod 10 = 0 Then
                            Call DrawText(objDraw, X + (lngPercW - objDraw.TextWidth(CStr(lngTmpVal))) / 2, Y + objDraw.TextHeight(CStr(lngTmpVal)) / 2 - 30, CStr(lngTmpVal), lngColor)
                        End If
                    Case Else
                        strTmp = CStr(lngTmpVal)
                        Call GetValue(strTmp)
                        Call DrawText(objDraw, X + (lngPercW - objDraw.TextWidth(strTmp)) / 2, Y + objDraw.TextHeight(CStr(CStr(lngTmpVal))) / 2 - 30, strTmp, lngColor)
                    End Select
                    
                End If
                Y = Y + H_9pt + H_9pt \ 2
                l = l + 1
            Else
                Y = Y + H_9pt + H_9pt \ 2
            End If
        Next
        X = X + lngPercW
    Next
    
    X = lngTmpX
    Y = lngTmpY

    '����
    Call DrawCell(objDraw, 1, lngTmpX, lngTmpY, (W_9pt + W_9pt \ 2) * (intCountItem + 1), 0)
    For i = 0 To intCountItem + 1
        Call DrawCell(objDraw, 1, X, Y, 0, (intGuageRow) * (H_9pt + H_9pt \ 2) + H_9pt * 7)

        If i = 0 Or i = intCountItem + 1 Then
            Call DrawLine(objDraw, X, Y, X, Y + (intGuageRow) * (H_9pt + H_9pt \ 2) + H_9pt * 7, , , 2)
        End If

        X = X + IIf(i = 0, W_9pt * 3, W_9pt * 7)
    Next
    
    'Ϊ�˽�ֵ������һ����ͼ����ʼ��
    X = lngTmpX + W_9pt * 10
    Y = lngTmpY
End Function

Private Function GetSplitStr(strValue As String, ByVal Index As Long) As String
    '���ܣ����б��еõ�ָ���ַ���
    Dim arrStrTmp As Variant
    If strValue = "" Then Exit Function
    arrStrTmp = Split(strValue, "'")
    If Index >= LBound(arrStrTmp) And Index <= UBound(arrStrTmp) Then
        GetSplitStr = arrStrTmp(Index)
    End If
End Function

Private Sub CloseRs(rs As ADODB.Recordset)
    '���ܣ��ر�Recordset����
    On Error Resume Next
    If rs.State = ADODB.adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Private Sub CalcOPS(lngArryOPSDay() As String, ByVal DayCount As Long, ByVal Index As Long, Optional ByVal lng���� As Long)
    '����:����ÿ�������������
    '����:lngArryOPSDay=Ҫ�������õ���������
    '     DayCount=ΪסԺ��������,�˲������Զ���ʼ��lngArryOPSDay����ĳ���
    '     Index=Ҫ��ĳ�����������Ϊ����
    Dim i As Long
    Dim strTmp As String
    
    If Index > DayCount - 1 Then Exit Sub
    
    If UBound(lngArryOPSDay) < DayCount - 1 Then
        ReDim Preserve lngArryOPSDay(DayCount - 1)
    End If
    For i = Index To Index + mintOpDays
        
        If i <= UBound(lngArryOPSDay) Then
        Select Case i - Index
        Case 0
        
            strTmp = lng����
            If mblnStopFlag Then
                lngArryOPSDay(i) = "0"
            Else
                lngArryOPSDay(i) = IIf(lngArryOPSDay(i) <> "" And lngArryOPSDay(i) <> "-1", lngArryOPSDay(i) & "(" & strTmp & ")", "0")
            End If
            
        Case Else
            
            If (i - Index) <= mintOpDays Then
                
                If mblnStopFlag Then
                    lngArryOPSDay(i) = (i - Index)
                Else
                    lngArryOPSDay(i) = IIf(lngArryOPSDay(i) <> "", (i - Index) & "/" & lngArryOPSDay(i), (i - Index))
                End If
                
            End If
        End Select
        End If
    Next
End Sub

Public Function ConvertTimeToChinese(ByVal strTime As String) As String
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp1 As String
    Dim strTmp2 As String
    
    If InStr(strTime, ":") <= 0 Then Exit Function
    On Error GoTo errHand
    
    strTmp1 = Left(strTime, InStr(strTime, ":") - 1)
    strTmp2 = Mid(strTime, InStr(strTime, ":") + 1)
    
    strTmp1 = Switch(strTmp1 = "00", "��", strTmp1 = "01", "һ", strTmp1 = "02", "��", strTmp1 = "03", "��", _
                    strTmp1 = "04", "��", strTmp1 = "05", "��", strTmp1 = "06", "��", strTmp1 = "07", "��", _
                    strTmp1 = "08", "��", strTmp1 = "09", "��", strTmp1 = "10", "ʮ", strTmp1 = "11", "ʮһ", _
                    strTmp1 = "12", "ʮ��", strTmp1 = "13", "ʮ��", strTmp1 = "14", "ʮ��", strTmp1 = "15", "ʮ��", _
                    strTmp1 = "16", "ʮ��", strTmp1 = "17", "ʮ��", strTmp1 = "18", "ʮ��", strTmp1 = "19", "ʮ��", _
                    strTmp1 = "20", "��ʮ", strTmp1 = "21", "��ʮһ", strTmp1 = "22", "��ʮ��", strTmp1 = "23", "��ʮ��")
    
    strTmp2 = Switch(strTmp2 = "00", "��", strTmp2 = "01", "һ", strTmp2 = "02", "��", strTmp2 = "03", "��", _
                    strTmp2 = "04", "��", strTmp2 = "05", "��", strTmp2 = "06", "��", strTmp2 = "07", "��", _
                    strTmp2 = "08", "��", strTmp2 = "09", "��", strTmp2 = "10", "ʮ", strTmp2 = "11", "ʮһ", _
                    strTmp2 = "12", "ʮ��", strTmp2 = "13", "ʮ��", strTmp2 = "14", "ʮ��", strTmp2 = "15", "ʮ��", _
                    strTmp2 = "16", "ʮ��", strTmp2 = "17", "ʮ��", strTmp2 = "18", "ʮ��", strTmp2 = "19", "ʮ��", _
                    strTmp2 = "20", "��ʮ", strTmp2 = "21", "��ʮһ", strTmp2 = "22", "��ʮ��", strTmp2 = "23", "��ʮ��", _
                    strTmp2 = "24", "��ʮ��", strTmp2 = "25", "��ʮ��", strTmp2 = "26", "��ʮ��", strTmp2 = "27", "��ʮ��", _
                    strTmp2 = "28", "��ʮ��", strTmp2 = "29", "��ʮ��", strTmp2 = "30", "��ʮ", strTmp2 = "31", "��ʮһ", _
                    strTmp2 = "32", "��ʮ��", strTmp2 = "33", "��ʮ��", strTmp2 = "34", "��ʮ��", strTmp2 = "35", "��ʮ��", _
                    strTmp2 = "36", "��ʮ��", strTmp2 = "37", "��ʮ��", strTmp2 = "38", "��ʮ��", strTmp2 = "39", "��ʮ��", _
                    strTmp2 = "40", "��ʮ", strTmp2 = "41", "��ʮһ", strTmp2 = "42", "��ʮ��", strTmp2 = "43", "��ʮ��", _
                    strTmp2 = "44", "��ʮ��", strTmp2 = "45", "��ʮ��", strTmp2 = "46", "��ʮ��", strTmp2 = "47", "��ʮ��", _
                    strTmp2 = "48", "��ʮ��", strTmp2 = "49", "��ʮ��", strTmp2 = "50", "��ʮ", strTmp2 = "51", "��ʮһ", _
                    strTmp2 = "52", "��ʮ��", strTmp2 = "53", "��ʮ��", strTmp2 = "54", "��ʮ��", strTmp2 = "55", "��ʮ��", _
                    strTmp2 = "56", "��ʮ��", strTmp2 = "57", "��ʮ��", strTmp2 = "58", "��ʮ��", strTmp2 = "59", "��ʮ��")
                    
    ConvertTimeToChinese = strTmp1 & "ʱ" & strTmp2 & "��"
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function PrintOrPreviewBodyState(objOut As Object, _
                                        ByVal lng����id As Long, _
                                        ByVal lng��ҳid As Long, _
                                        ByVal intBaby As Integer, _
                                        ByVal lngSectID As Long, _
                                        ByVal lngBeginY As Long, _
                                        ByVal lngLeft As Long, _
                                        ByVal objParent As Object, _
                                        Optional ByVal blnKeepOn As Boolean = False, _
                                        Optional ByVal intBeginPage As Integer = -1, _
                                        Optional ByVal intEndPage As Integer = -1, _
                                        Optional ByVal intPageNo As Integer = -1, _
                                        Optional ByVal sngScale As Single = 1, _
                                        Optional ByVal blnMoved As Boolean) As Boolean
    '******************************************************************************************************************
    '����:��ӡ��Ԥ��ĳ������¶ȱ�
    '����:objOut=�������,����ΪPrinter��һ������(�����а����ؼ�����picPage)
    '      lngCaseRecordID=������¼id
    '      lngBeginY=��ʼ������
    '      blnKeepOn=�Ƿ񱣳�����
    '      objParent=�����ô���
    '      intBeginPage=Ҫ��ʼҳ�����,��Ϊ-1ʱ��ʾ�������.
    '      intEndPage=����ҳ������intEndPage����ʵ��ҳ����ֻ��ӡ��ʵ��ҳ��
    '      intPageNO=��ʼ��ҳ��,���Ϊ-1��ʾ����ʾҳ��
    '      sngScale=�������
    
    '����:���δ�ӡ�����Ƿ�ɹ�
    '******************************************************************************************************************
    
    Dim blnPrint As Boolean
    Dim strInfo As String                                   '�ǲ��Ǵ�ӡ�� ����ʾ��ʾ��Ϣ
    Dim lngPage As Long                                     '��ǰҳ
    Dim intAllOpt As Single
    Dim intCurOpt As Single                                 '�ܽ��ȣ���ǰ����
    Dim i As Long, j As Long, l As Long, lngRecItemRow As Long, lngRecordCount As Long
    Dim X As Long, Y As Long                                'X���꣬Y���꣨Twip����
    Dim objDraw As Object                                   '���л�ͼ�Ķ���
    Dim intDrawLineRows As Integer                          '�����߱��������(����������)�����20��
    Dim intDrawLineCols As Integer                          '�����߱��������
    Dim intDrawGridRows As Integer                          '����ײ����±�¼����Ŀ������
    Dim intRepairRows As Integer
    Dim strBeginDate As String, strEndDate As String        '������˵Ŀ�ʼ����ֹʱ��
    Dim strPatiInfo As String                               '������Ϣ�б�:���� סԺ�� ���� ���� ���� ����
    Dim strStateTips As String                              '�ײ���˵����Ϣ
    Dim strArrItemInfo() As String                          '���±�������Ŀ�б�����
    Dim strArrItemDataInfo() As String                      '���±�������Ŀ�����б�����
    Dim strArrItemDataComment() As String                   '���������Ŀ���ݵ�˵������
    Dim strArrRecordItemInfo() As String                    '���±�¼����Ŀ�б�����
    Dim rsArrRecordItemInfo As New ADODB.Recordset          '���±�¼����Ŀ�б��¼
    Dim lngArr����������() As String
    Dim lngOPSDayCount As Long
    Dim blnTag As Boolean 'ȷ���Ƿ�������������Ŀ������
    Dim blnComment As Boolean 'ȷ���Ƿ�������������Ŀ��˵��
    Dim lngCountPage As Long '���ݲ��˵����������ҳ��
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    'ֽ�ųߴ���Ϣ
    Dim lngTop As Long
    Dim dblSureW As Double, dblSureH As Double
    Dim H_9pt As Long, W_9pt As Long 'һ��С���ֵĸ߶�
    Dim H_16pt As Long, W_16pt As Long
    Dim lngTmpX As Long, lngTmpY As Long
    Dim mlngHourBegin As Long
    Dim strTmp As String
    Dim lngCol As Long
    Dim strTime As String
    Dim strSvrTime As String
    Dim intCount As Integer
    Dim lngTmpDay As Long    '��ʱ������¼��ǰҳĳʱ����϶�Ӧ������
    Dim strTmpDate As String '��ʱ������¼��ǰҳ�������ʱ���
    Dim strTmpDay As String '��ʱ���浱ǰҳ��Ŀ�ʼ����
    Dim strTmpString0 As String, strTmpString1 As String, strTmpString2 As String '��ʱ��
    Dim strNewTmpString As String
    Dim lngNewTmpX As Long, lngNewTmpY As Long
    Dim lngPicPageIndex As Long
    Dim strArrNewTmp() As String, strArrNewTmpComment() As String
    Dim strArrCurLineData(0 To 41) As String
    Dim strArrCurLineComment(0 To 41) As String
    Dim strArrItemDataInOut(0 To 41) As String
    Dim strArrItemTmpInOut() As String
    Dim aryTmp As Variant
    Dim rsTmp As New ADODB.Recordset
    Dim lngValue As Long
    Dim bytδ����ʾλ�� As Byte
    Dim lngƫ����(0 To 41) As Long
    Dim blnAllow As Boolean
    Dim blnShow As Boolean
        
    Dim lngRowCount As Long
    
    ReDim mbyt����(0 To 41) As Byte
    
    '�����Ŀ����ʱ�õ�����Ҫ��Ϊ����� ��λ����
    Dim dbl��� As Double, dbl��С As Double, dbl��λֵ As Double, lng����� As Long, lng��λ���� As Long
    
    
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrPrint
    
    msngScale = 0.85
    
    '��ȡ���±�һ�쿪ʼʱ��
    '------------------------------------------------------------------------------------------------------------------
    mlngHourBegin = Val(zlDatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4))

    mintOpDays = Val(zlDatabase.GetPara("�������ע����", glngSys, 1255, "10"))
    mblnStopFlag = (Val(zlDatabase.GetPara("�ٴ�����ֹͣǰ�α�ע", glngSys, 1255, "0")) = 1)
    bytδ����ʾλ�� = Val(zlDatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0"))
    mblnӤ�����µ���ʾ��Ժ = (zlDatabase.GetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", glngSys, 1255, 1) = 1)
    
    '���˱䶯�����ʾ����
    '------------------------------------------------------------------------------------------------------------------
    strTmp = zlDatabase.GetPara("���µ����", glngSys, 1255, "1;1;1;1;1;1;1;1")
    If UBound(Split(strTmp, ";")) >= 5 Then
        mBodyFlag.��Ժ = Val(Split(strTmp, ";")(0))
        mBodyFlag.��� = Val(Split(strTmp, ";")(1))
        mBodyFlag.ת�� = Val(Split(strTmp, ";")(2))
        mBodyFlag.���� = Val(Split(strTmp, ";")(3))
        mBodyFlag.���� = Val(Split(strTmp, ";")(4))
        mBodyFlag.��Ժ = Val(Split(strTmp, ";")(5))
        If UBound(Split(strTmp, ";")) >= 6 Then mBodyFlag.���� = Val(Split(strTmp, ";")(6))
        If UBound(Split(strTmp, ";")) >= 7 Then mBodyFlag.���� = Val(Split(strTmp, ";")(7))
    End If
    
    blnPrint = TypeName(objOut) = "Printer"
    Screen.MousePointer = 11
    intAllOpt = 6
    
    '������ȴ���
    '------------------------------------------------------------------------------------------------------------------
    strInfo = "����" & IIf(blnPrint, "׼����ӡ���±�", "����Ԥ��") & ",���Ժ�..."
    Call ShowFlash(strInfo, , objParent)
    
    '��ӡǰ�����
    If blnKeepOn = False Then
        If Not blnPrint Then
            For i = objOut.picPage.UBound To 0 Step -1
                If i = 0 Then
                    objOut.picPage(i).Cls
'                    objOut.txtPage.Text = ""
                Else
                    Unload objOut.picPage(i)
                End If
            Next
            Set objDraw = objOut.picPage(0) 'PictureBox
            objDraw.Width = Printer.Width * sngScale
            objDraw.Height = Printer.Height * sngScale
        Else
            Set objDraw = Printer
        End If
    Else
        If Not blnPrint Then
            i = objOut.picPage.UBound + 1
            Load objOut.picPage(i)
'            objOut.SetPages
            Set objDraw = objOut.picPage(objOut.picPage.UBound)
            objDraw.Width = Printer.Width * sngScale
            objDraw.Height = Printer.Height * sngScale
        Else
            Set objDraw = Printer
        End If
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    strSQL = _
        "   Select Decode(b.����ʱ��,Null,a.��ʼ,b.����ʱ��) As ��ʼ,a.��ֹ From (Select ����ID,��ҳid,Min(��ʼʱ��) as ��ʼ,Max(Nvl(��ֹʱ��,sysdate)) as ��ֹ" & _
        "    From ���˱䶯��¼" & _
        "    Where ��ʼʱ�� is Not Null And ����ID=[1] And ��ҳID=[2] Group By ����ID,��ҳid) a, " & _
        "   (Select ����ID,��ҳid,����ʱ�� From ������������¼ Where ����ID = [1] And ��ҳID = [2] And ���=[3]) b Where a.����id=b.����id(+) And a.��ҳid=b.��ҳid(+) "

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng����id, lng��ҳid, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        lngCountPage = DateDiff("d", rsTemp!��ʼ, rsTemp!��ֹ) + 1
        lngCountPage = IIf(lngCountPage / 7 = Fix(lngCountPage / 7), lngCountPage / 7, Fix(lngCountPage / 7) + 1)
        strBeginDate = Format(rsTemp!��ʼ, "YYYY-MM-DD HH:MM:SS")
        strEndDate = Format(rsTemp!��ֹ, "YYYY-MM-DD HH:MM:SS")
    Else
        CloseRs rsTemp
        GoTo ErrPrint '�������˱䶯��Ϣ�˳�
    End If
    
    ReDim lngArr����������(DateDiff("d", CDate(strBeginDate), CDate(strEndDate)))

    intCurOpt = intCurOpt + 1
    
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '------------------------------------------------------------------------------------------------------------------
    '�ڣ����ݣ����˵Ļ�����Ϣ
    '��ȡ���˻�����Ϣ
    Dim varPatiInfo As Variant
    
    '����'סԺ��''��Ժʱ��
    strPatiInfo = "''''''"
    varPatiInfo = Split(strPatiInfo, "'")
    
    strSQL = " Select  b.����,A.סԺ��,b.��Ժʱ��,b.�Ա�,b.���� From ������Ϣ B,������ҳ A Where A.����ID=B.����ID And A.����id=[1] And A.��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng����id, lng��ҳid)
    If rsTemp.BOF = False Then
        varPatiInfo(0) = zlCommFun.NVL(rsTemp("����").Value)
        varPatiInfo(1) = zlCommFun.NVL(rsTemp("סԺ��").Value)
        varPatiInfo(3) = Format(zlCommFun.NVL(rsTemp("��Ժʱ��").Value), "yyyy-MM-dd")
        varPatiInfo(5) = zlCommFun.NVL(rsTemp("�Ա�").Value)
        varPatiInfo(6) = zlCommFun.NVL(rsTemp("����").Value)
    End If
    
    '��Ժʱ��(�����ʱ��Ϊ׼)
    mstrSQL = "select ��ʼʱ�� from ���˱䶯��¼ where ����id=[1] And ��ҳid=[2] and ��ʼԭ��=2 order by ��ʼʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng����id, lng��ҳid)
    If rsTemp.BOF = False Then
        varPatiInfo(3) = Format(zlCommFun.NVL(rsTemp("��ʼʱ��").Value), "yyyy-MM-dd")
    End If
        
        
    Select Case intBaby
    Case 0
        
    Case Else
        
        varPatiInfo(5) = ""
        varPatiInfo(6) = ""
        
        gstrSQL = "Select Decode(a.Ӥ������,Null,b.����||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������,Ӥ���Ա�,����ʱ�� From ������������¼ a,������Ϣ b Where a.����id=[1] And a.��ҳid=[2] And a.����id=b.����id And a.���=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlPrint", lng����id, lng��ҳid, intBaby)
        If rsTemp.BOF = False Then
            varPatiInfo(0) = rsTemp("Ӥ������").Value
            varPatiInfo(5) = zlCommFun.NVL(rsTemp("Ӥ���Ա�").Value)
            varPatiInfo(6) = "������"
            
            If IsNull(rsTemp("����ʱ��").Value) = False Then varPatiInfo(3) = Format(zlCommFun.NVL(rsTemp("����ʱ��").Value), "yyyy-MM-dd")
        End If
        
    End Select
        

    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '------------------------------------------------------------------------------------------------------------------
    '�ڣ����ݣ����˵�������Ϣ
    
    '������������������
    '�����������������Щ��,����ʱ�䴫������
    '�ҳ�������������

    mstrSQL = "Select ʱ��,��Ŀ����,rownum as ���� From (SELECT A.����ʱ�� As ʱ��,c.��Ŀ���� " & _
                "FROM ���˻����¼ A,���˻������� C " & _
                "Where A.ID=C.��¼ID " & _
                    "AND A.����id=[1] And Nvl(a.Ӥ��,0)=[4] " & _
                    "AND A.��ҳid=[2] " & _
                    "AND c.��¼����=4 " & _
                    "AND A.����ʱ��<[3] And c.��ֹ�汾 Is Null Order By A.����ʱ��)"
    If mblnMoved Then
        mstrSQL = Replace(mstrSQL, "���˻����¼", "H���˻����¼")
        mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
    End If
    
    Dim TmplngDayCount As Long
    
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", _
                                        lng����id, _
                                        lng��ҳid, _
                                        CDate(strEndDate), intBaby)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        lngOPSDayCount = rsTmp.RecordCount
        
        TmplngDayCount = DateDiff("d", CDate(strBeginDate), CDate(strEndDate)) + 1
        
        For i = 0 To rsTmp.RecordCount - 1
            '��ʱ���������������������
            If IsNull(rsTmp!ʱ��) = False Then
                Call CalcOPS(lngArr����������, TmplngDayCount, DateDiff("d", CDate(strBeginDate), rsTmp!ʱ��), rsTmp!����)
            End If
            rsTmp.MoveNext
        Next
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '�ڣ����ݣ�����������Ŀ
    
    '1�������ʼ����Ŀ�ַ�������
    '�����±�������Ŀ
    
    mint����Ӧ�� = 2
    mstr���ʷ��� = ""
    mstrSQL = "Select a.Ӧ�÷�ʽ,b.��¼�� From �����¼��Ŀ a,���¼�¼��Ŀ b Where a.��Ŀ���=-1 And a.��Ŀ���=b.��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint")
    If rsTemp.BOF = False Then
        mint����Ӧ�� = zlCommFun.NVL(rsTemp("Ӧ�÷�ʽ").Value, 2)
        mstr���ʷ��� = zlCommFun.NVL(rsTemp("��¼��").Value, "��")
    End If
    
        
    '�õ�����������Ŀ
    mstrSQL = " Select c.��Ŀ���,A.��¼�� as ��Ŀ��,A.��¼��,To_Char(A.��¼ɫ)||''''||To_Char(A.�����)||''''||A.��¼��||''''||To_Char(A.���ֵ)||''''||To_Char(A.��Сֵ)||''''||To_Char(A.��λֵ)||''''||A.��¼��||''''||To_Char(A.��Ŀ���)||''''||c.��Ŀ��λ As �б� " & _
                " From ���¼�¼��Ŀ A,�����¼��Ŀ C " & _
                " Where A.��Ŀ���=C.��Ŀ��� And A.��¼��=[1] AND C.����ȼ�>=[2] And Nvl(c.Ӧ�÷�ʽ,0)=1 And Nvl(c.���ò���,0) In (0,[4]) " & _
                " And (c.���ÿ���=1 Or (c.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=c.��Ŀ��� And D.����id=[3]))) " & _
                " Order by A.�������"
                
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", 1, 0, lngSectID, IIf(intBaby = 0, 1, 2))
    If rsTemp.RecordCount < 1 Then
        CloseRs rsTemp
        MsgBox "���κ����±���Ŀ��", vbExclamation, gstrSysName
        GoTo ErrExit    '�������˳�
    End If
    rsTemp.MoveFirst
                        
    ReDim strArrItemInfo(rsTemp.RecordCount - 1)                    'ȷ����Ŀ��Ϣ�б�Ԫ�ظ���
    ReDim strArrItemDataInfo(rsTemp.RecordCount - 1)                'ȷ����Ŀ�����б�Ԫ�ظ���
    ReDim strArrItemDataComment(rsTemp.RecordCount - 1)             'ȷ��������Ŀ�������е�˵������
    Dim bytShow As Byte
    
    Dim varTmp As Variant

    
    mbln�������� = False
    strStateTips = "˵��:"
    For i = 0 To UBound(strArrItemInfo)
        strArrItemInfo(i) = rsTemp!�б�
        
        Select Case rsTemp("��Ŀ���").Value
        Case -1
            strTmp = rsTemp!��Ŀ�� & "(" & rsTemp!��¼�� & ")"
        Case 1
            varTmp = Split(zlCommFun.NVL(rsTemp("��¼��").Value, "��,��,��"), ",")
            
            mstrChar(0) = CStr(varTmp(0))
            mstrChar(1) = CStr(varTmp(1))
            mstrChar(2) = CStr(varTmp(2))
    
            strTmp = rsTemp!��Ŀ�� & "(����" & mstrChar(0) & ",Ҹ��" & mstrChar(1) & ",����" & mstrChar(2) & ")"
        Case 2
            '������:mint����Ӧ�� = 2
            mstrPulse = rsTemp!��¼��
            If mint����Ӧ�� = 0 Then
                strTmp = rsTemp!��Ŀ�� & "(" & rsTemp!��¼�� & ",���ʡ�)"
            Else
                strTmp = rsTemp!��Ŀ�� & "(" & rsTemp!��¼�� & ")"
            End If
        Case 3
            mstrBreath = rsTemp!��¼��
            mbln�������� = True
            strTmp = rsTemp!��Ŀ�� & "(" & rsTemp!��¼�� & ")"
        Case Else
            strTmp = rsTemp!��Ŀ�� & "(" & rsTemp!��¼�� & ")"
        End Select
        
        If i = 0 Then
            strStateTips = strStateTips & strTmp
        Else
            strStateTips = strStateTips & "��" & strTmp
        End If
        rsTemp.MoveNext
    Next
    intDrawLineCols = UBound(strArrItemInfo) + 1   '������±�������Ŀ�ĸ���
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '�ڣ����ݣ�����¼����Ŀ
    '------------------------------------------------------------------------------------------------------------------
    '�������¼����Ŀ
    mstrSQL = " Select RowNum-1 As ���,A.* From (Select Decode(A.��Ŀ���,4,'Ѫѹ',A.��¼��) as ��Ŀ��,C.��Ŀ��λ As ��λ,A.��Ŀ���,A.��¼Ƶ��,C.��Ŀ���� " & _
                " From ���¼�¼��Ŀ A,�����¼��Ŀ C " & _
                " Where A.��Ŀ���=C.��Ŀ��� And A.��¼��=[1] AND C.����ȼ�>=[2] And A.��Ŀ���<>5 And Nvl(c.Ӧ�÷�ʽ,0)=1 And Nvl(c.���ò���,0) In (0,[4]) " & _
                " And (c.���ÿ���=1 Or (c.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=c.��Ŀ��� And D.����id=[3]))) " & _
                " Order by A.�������) A"
    Set rsArrRecordItemInfo = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", 2, 0, lngSectID, IIf(intBaby = 0, 1, 2))
    If rsArrRecordItemInfo.RecordCount > 0 Then
        rsArrRecordItemInfo.MoveFirst
        ReDim strArrRecordItemInfo(rsArrRecordItemInfo.RecordCount - 1)
        For i = 0 To UBound(strArrRecordItemInfo)
            
            strArrRecordItemInfo(i) = rsArrRecordItemInfo!��Ŀ�� & "'" & IIf(IsNull(rsArrRecordItemInfo!��λ), "", rsArrRecordItemInfo!��λ) & "'" & zlCommFun.NVL(rsArrRecordItemInfo("��¼Ƶ��").Value, 2)
            rsArrRecordItemInfo.MoveNext
        Next
        intDrawGridRows = UBound(strArrRecordItemInfo) + 1 '������±�¼����Ŀ������

    Else
        intDrawGridRows = 0 '������±�¼����Ŀ������
    End If
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    
    '�ڣ����ݣ�ͼ���������
    '------------------------------------------------------------------------------------------------------------------
    '2��ȷ��X��Y������λ��
    '�߽���Ϣ(Twip)
    lngLeft = Val(zlDatabase.GetPara("���µ���߾�", glngSys, 1255, OFFSET_LEFT)) * 56.7 * sngScale
    lngTop = Val(zlDatabase.GetPara("���µ��ϱ߾�", glngSys, 1255, OFFSET_TOP)) * 56.7 * sngScale
    
    
    'ȷ����ʼ����
    X = lngLeft * sngScale: Y = lngTop * sngScale
    objDraw.CurrentX = X: objDraw.CurrentY = Y
    '3����ӡ����ĳ�ʼ��
    '����ο��߶�
    objDraw.Font = "����"
    objDraw.FontSize = 24 * sngScale
    objDraw.FontBold = True
    H_16pt = objDraw.TextHeight("��")
    W_16pt = objDraw.TextWidth("��")
    
    lngTmpX = X
    lngTmpY = Y
    objDraw.Font = "����"
    objDraw.FontSize = 9 * sngScale
    objDraw.FontBold = False
    H_9pt = (ROWHEIGHT * 10 / 3) * sngScale
    W_9pt = objDraw.TextWidth("��")
    
    mlngFirstWidth = W_9pt * 10
    
    '------------------------------------------------------------------------------------------------------------------
    '4��������п�����߱���ܹ��ж�����
    '������±���Ŀ��������
    strSQL = "Select Max((A.���ֵ-A.��Сֵ)/Decode(A.��λֵ,0,1,A.��λֵ)+A.�����) as �ܸ߶� From ���¼�¼��Ŀ A Where A.��¼��=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", 1)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        intDrawLineRows = MAXROWS
    Else
        CloseRs rsTemp
        GoTo ErrPrint
    End If
    
    If intDrawLineRows < 1 Then
        CloseRs rsTemp
        GoTo ErrPrint
    End If

'    lngColFirstWidth = 6 * W_9pt
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    '5��ѭ��������ҳ��ѭ��
    
    intCurOpt = 0
    intAllOpt = 100
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    If blnPrint = False Then
        lngPicPageIndex = objOut.picPage.UBound + 1
    End If
    
    '��ʽ��ʼ���Ĳ���ѭ��ÿһҳ
    '------------------------------------------------------------------------------------------------------------------
    For lngPage = 1 To lngCountPage
    
        strTmpDay = Format(CDate(strBeginDate) + 7 * (lngPage - 1), "YYYY-MM-DD") '��õ�ǰҳ��ĵ�һ��������ʱ��
        If strTmpDay < strBeginDate Then strTmpDay = strBeginDate
        
        strTmpDate = Format(CDate(strTmpDay), "MM-DD") '��¼��ǰ�ڼ��
        strTmpDate = strTmpDate & "��" & Format(IIf(CDate(strBeginDate) + 7 * (lngPage - 1) + 6 > CDate(strEndDate), strEndDate, (CDate(strBeginDate) + 7 * (lngPage - 1) + 6)), "MM-DD")
            
        intCurOpt = lngPage / lngCountPage
        strInfo = "����" & IIf(blnPrint, "��ӡ���±�", "Ԥ��") & ",���Ժ�..."
        Call ShowFlash(strInfo, intCurOpt, objParent)
        
        '��ҳ�Ŵ�ӡ
        If intBeginPage > 0 Then  'ֻ��ӡָ��ҳ���
            If lngPage >= intBeginPage And lngPage <= intEndPage Then
                If lngPage > intBeginPage Then  '���ڶ�ҳʱ��ʼ��ʼ��ֽ�Ż�ҳ��
                    If Not blnPrint Then
                        Load objOut.picPage(lngPicPageIndex)
'                        objOut.cboPage.AddItem "�� " & (lngPage - intBeginPage + 1) & " ҳ"
'                        objOut.txtPage.Text = "��ǰҳ" & Space(17) & "�� " & objOut.picPage.UBound + 1 & " ҳ"
                        Set objDraw = objOut.picPage(lngPicPageIndex) 'PictureBox
                        objDraw.Cls
                        objDraw.Width = Printer.Width * sngScale
                        objDraw.Height = Printer.Height * sngScale
                        lngPicPageIndex = lngPicPageIndex + 1
                    Else
                        Printer.NewPage
                    End If
                End If
            Else
                GoTo NOPageSub
            End If
        Else  '��ӡ����ʱ
            If lngPage > 1 Then
                If Not blnPrint Then
                    Load objOut.picPage(lngPicPageIndex)
'                    objOut.cboPage.AddItem "�� " & lngPage & " ҳ"
'                    objOut.txtPage.Text = "��ǰҳ" & Space(17) & "�� " & objOut.picPage.UBound + 1 & " ҳ"
                    Set objDraw = objOut.picPage(lngPicPageIndex) 'PictureBox
                    objDraw.Cls
                    objDraw.Width = Printer.Width * sngScale
                    objDraw.Height = Printer.Height * sngScale
                    lngPicPageIndex = lngPicPageIndex + 1
                Else
                    Printer.NewPage
                End If
            End If
        End If
        
        X = lngTmpX
        Y = lngTmpY

        objDraw.Font = "����"
        objDraw.FontSize = 9 * sngScale
        objDraw.FontBold = False
        '��ӡ�ʿغ�
        
        strTmp = zlDatabase.GetPara("�ʿغ�", glngSys, 1255, "")
        
        X = lngTmpX + (6 * W_9pt + HOUR_STEP_Twips * 6 * 7) - objDraw.TextWidth(strTmp)

        Call DrawText(objDraw, X, lngTmpY - objDraw.TextHeight(strTmp), strTmp)
        
        X = lngTmpX
        
        objDraw.Font = "����"
        objDraw.FontSize = 18 * sngScale
        objDraw.FontBold = True
        
        '��ӡҽԺ����,��ӡ���µ�����
        strTmpString0 = IIf(GetUnitName = "-", "", GetUnitName) & "���µ�"
        If strTmpString0 <> "" Then
            Call DrawCell(objDraw, strTmpString0, lngTmpX + (((UBound(strArrItemInfo) + 1) * (W_9pt + W_9pt \ 2) + HOUR_STEP_Twips * 6 * 7) - objDraw.TextWidth(strTmpString0)) / 2, lngTmpY, objDraw.TextWidth(strTmpString0), H_16pt + H_16pt \ 2, , , , , , objDraw.Font, "0000", 1, 1)
        End If

        strTmpString0 = ""
        Y = Y + H_16pt + 2 * H_16pt / 3
        
        objDraw.Font = "����"
        objDraw.FontSize = 10 * sngScale
        objDraw.FontBold = True
        

        varPatiInfo(2) = ""
        varPatiInfo(4) = ""
        strTmp = ""
        strTime = ""
        
        strSQL = " Select  c.���� As ����,b.���� As ����,d.����� AS ����,a.��ʼԭ�� " & _
                "From ���˱䶯��¼ a,���ű� b,���ű� c,��λ״����¼ D " & _
                "Where a.����id=[1] And a.��ҳid=[2] And a.����id Is Not Null And a.����id=b.id and a.����id=c.id " & _
                "And a.����=d.���� And a.��ʼʱ��-4/24<=[3] And Nvl(a.��ֹʱ��,Sysdate)>=[4] " & _
                "Order By a.��ʼʱ�� desc"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng����id, lng��ҳid, CDate(strTmpDay) + 7, CDate(strTmpDay))
        If rsTmp.BOF = False Then
            varPatiInfo(2) = zlCommFun.NVL(rsTmp("����").Value)
            varPatiInfo(4) = zlCommFun.NVL(rsTmp("����").Value)
            
'            Do While Not rsTmp.EOF
'
'                If zlCommFun.NVL(rsTmp("����").Value) <> strTmp And zlCommFun.NVL(rsTmp("����").Value) <> "" Then
'
'                    strTmp = zlCommFun.NVL(rsTmp("����").Value)
'
'                    If varPatiInfo(2) = "" Then
'                        varPatiInfo(2) = strTmp
'                    Else
'                        varPatiInfo(2) = varPatiInfo(2) & "->" & strTmp
'                    End If
'
'                End If
'
'                If zlCommFun.NVL(rsTmp("����").Value) <> strTime And zlCommFun.NVL(rsTmp("����").Value) <> "" Then
'
'                    strTime = zlCommFun.NVL(rsTmp("����").Value)
'
'                    If varPatiInfo(4) = "" Then
'                        varPatiInfo(4) = strTime
'                    Else
'                        varPatiInfo(4) = varPatiInfo(4) & "->" & strTime
'                    End If
'
'                End If
'
'                rsTmp.MoveNext
'            Loop
'
'            If Left(varPatiInfo(2), 2) = "->" Then varPatiInfo(2) = Mid(varPatiInfo(2), 3)
'            If Left(varPatiInfo(4), 2) = "->" Then varPatiInfo(4) = Mid(varPatiInfo(4), 3)

        End If
        
        strPatiInfo = Join(varPatiInfo, "'")
    

        mstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2) As ������ From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "���µ�", "������", lng����id, lng��ҳid)
        If rsTmp.BOF = False Then
            If intBaby = 0 Then
                strPatiInfo = strPatiInfo & "'" & zlCommFun.NVL(rsTmp("������").Value)
            Else
                strPatiInfo = strPatiInfo & "'"
            End If
        Else
            strPatiInfo = strPatiInfo & "'"
        End If
    
    
        '���Ȼ��������еĲ�����Ϣ���ע�⣺���滹û�м��� {����} ����ʾ��
        Call DrawPatiInfo(objDraw, X, Y, strPatiInfo & "'" & strTmpDate, lngLeft + 9726)

        objDraw.Font = "����"
        objDraw.FontSize = 9 * sngScale
        objDraw.FontBold = False
                
        '6���������סԺ���ڼ�����ʱ��
        '�����ǰ����ڼ�ε�������סԺ����
        '------------------------------------------------------------------------------------------------------------------
        strTmpString0 = ""
        strTmpString1 = ""
        strTmpString2 = ""
        
        lngValue = 0
        mstrSQL = "Select zl_CalcInDays([1],[2],[3],[4]) As ��ʼ���� From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "���µ�", lng����id, lng��ҳid, intBaby, (Int(CDate(strTmpDay))))
        If rsTmp.BOF = False Then
            lngValue = rsTmp("��ʼ����").Value
        End If
        
        For i = 0 To 6
            strTmpString0 = strTmpString0 & "'" & Format(CDate(strTmpDay) + i, "YYYY-MM-DD")
            If CDate(Format(strTmpDay, "YYYY-MM-DD")) + i > CDate(Format(strEndDate, "YYYY-MM-DD")) Then
                strTmpString1 = strTmpString1 & "'"
                strTmpString2 = strTmpString2 & "'"
            Else
'                strTmpString1 = strTmpString1 & "'" & Format(Int(CDate(strTmpDay)) + i - Int(CDate(strBeginDate)) + 1)
                strTmpString1 = strTmpString1 & "'" & lngValue + i
                If lngOPSDayCount > 0 Then
                    strTmpString2 = strTmpString2 & "'" & IIf(CStr(lngArr����������((lngPage - 1) * 7 + i)) = "-1", "", CStr(lngArr����������((lngPage - 1) * 7 + i)))
                End If
            End If
        Next
        strTmpString0 = Right(strTmpString0, Len(strTmpString0) - 1)
        strTmpString1 = Right(strTmpString1, Len(strTmpString1) - 1)
        If lngOPSDayCount > 0 Then
            strTmpString2 = Right(strTmpString2, Len(strTmpString2) - 1)
        Else
            strTmpString2 = "'''''''''"
        End If
        lngNewTmpX = X
        
        Call DrawBodyInfo(objDraw, X, Y, mlngFirstWidth, strTmpString0, strTmpString1, strTmpString2, , strBeginDate)
        lngNewTmpY = Y + (intDrawLineRows) * (H_9pt + H_9pt \ 2) + H_9pt * 2 + H_9pt / 2 - 30
        
        Call DrawBodyScale(objDraw, X, Y, intDrawLineRows, strArrItemInfo)
        Call DrawBodyTopScale(objDraw, X, Y)
        Call DrawBodyPaper(objDraw, X, Y, intDrawLineRows)
        
        '���ת���
        '------------------------------------------------------------------------------------------------------------------
        Set rsTmp = GetDataFromHis(lng����id, lng��ҳid, intBaby, CDate(strTmpDay), CDate(strTmpDay) + 8, 2)
        If Not (rsTmp Is Nothing) Then
            If rsTmp.BOF = False Then
                
                Do While Not rsTmp.EOF
                    
                    If zlCommFun.NVL(rsTmp("����")) <> "" Then
                        
                        bytShow = 0
                        Select Case Val(rsTmp("�к�").Value)
                        Case 5
                            bytShow = mBodyFlag.��Ժ
                        Case 6
                            bytShow = mBodyFlag.ת��
                        Case 7
                            bytShow = mBodyFlag.����
                        Case 8
                            bytShow = mBodyFlag.��Ժ
                        Case 9
                            bytShow = mBodyFlag.���
                        End Select
                    
                        If bytShow > 0 Then
                            blnShow = True
                            If Val(rsTmp("�к�").Value) = 8 And intBaby > 0 Then
                                blnShow = mblnӤ�����µ���ʾ��Ժ
                            End If
                            
                            If blnShow Then
                                Select Case bytShow
                                Case 1
                                    strTmp = rsTmp("����").Value
                                Case 2
                                    strTmp = rsTmp("����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                                Case 3
                                    strTmp = rsTmp("����").Value & rsTmp("����").Value
                                Case 4
                                    strTmp = rsTmp("����").Value & rsTmp("����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                                End Select
                                strTmp = strTmp & "'" & Format(rsTmp("ʱ��").Value, "HH:mm:ss")
                            Else
                                strTmp = ""
                            End If
                        Else
                            strTmp = ""
                        End If
                        
                        intCount = GetCurveColumn(CDate(rsTmp("ʱ��").Value), CDate(strTmpDay), mlngHourBegin) - 1
                        
                        If intCount >= 0 And intCount <= 41 Then
                            strArrItemDataInOut(intCount) = IIf(Trim(strArrItemDataInOut(intCount)) = "", "", Trim(strArrItemDataInOut(intCount)) & ";") & strTmp
                        End If
                        
                    End If
                    
                    rsTmp.MoveNext
                Loop
            End If
        End If
        
        If intBaby > 0 Then
            
            Set rsTmp = GetDataFromHis(lng����id, lng��ҳid, intBaby, CDate(strTmpDay), CDate(strTmpDay) + 8, 3)
            If Not (rsTmp Is Nothing) Then
                If rsTmp.BOF = False Then
                    
                    Do While Not rsTmp.EOF
                        
                        If zlCommFun.NVL(rsTmp("����")) <> "" Then
                        
                            If mBodyFlag.���� > 0 Then
                            
                                Select Case mBodyFlag.����
                                Case 1
                                    strTmp = rsTmp("����").Value
                                Case 2
                                    strTmp = rsTmp("����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                                Case 3
                                    strTmp = rsTmp("����").Value & rsTmp("����").Value
                                Case 4
                                    strTmp = rsTmp("����").Value & rsTmp("����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                                End Select
                                strTmp = strTmp & "'" & Format(rsTmp("ʱ��").Value, "HH:mm:ss")
                            Else
                                strTmp = ""
                            End If
                            
                            intCount = GetCurveColumn(CDate(rsTmp("ʱ��").Value), CDate(strTmpDay), mlngHourBegin) - 1
                            
                            If intCount >= 0 And intCount <= 41 Then
                                strArrItemDataInOut(intCount) = IIf(Trim(strArrItemDataInOut(intCount)) = "", "", Trim(strArrItemDataInOut(intCount)) & ";") & strTmp
                            End If
                            
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                End If
            End If
        End If
        
        '��ȡ����Ϊ���ת��ʶ
        '------------------------------------------------------------------------------------------------------------------
        mstrSQL = "SELECT A.����ʱ�� As ʱ��,c.��Ŀ���� " & _
                    "FROM ���˻����¼ A,���˻������� C " & _
                    "Where A.ID=C.��¼ID " & _
                        "AND A.����id=[1] " & _
                        "AND A.��ҳid=[2]  And Nvl(a.Ӥ��,0)=[5] " & _
                        "AND c.��¼����=4 " & _
                        "AND A.����ʱ�� Between [3] And [4] And c.��ֹ�汾 Is Null Order By A.����ʱ��"
        If mblnMoved Then
            mstrSQL = Replace(mstrSQL, "���˻����¼", "H���˻����¼")
            mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng����id, lng��ҳid, CDate(strTmpDay), CDate(strTmpDay) + 8, intBaby)
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            
            For i = 0 To rsTmp.RecordCount - 1
                '��ʱ��������������������У�
                If IsNull(rsTmp!ʱ��) = False Then
                    lngCol = GetCurveColumn(CDate(rsTmp("ʱ��").Value), CDate(strTmpDay), mlngHourBegin) - 1
                    
                    Select Case rsTmp("��Ŀ����").Value
                    Case "����"
                        If lngCol >= 0 And lngCol <= 41 And mBodyFlag.���� > 0 Then
                            If mBodyFlag.���� = 2 Then
                                strTmp = rsTmp("��Ŀ����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                            Else
                                strTmp = rsTmp("��Ŀ����").Value
                            End If
                            strTmp = strTmp & "'" & Format(rsTmp("ʱ��").Value, "HH:mm:ss")
                            
                            strArrItemDataInOut(lngCol) = IIf(Trim(strArrItemDataInOut(lngCol)) = "", "", Trim(strArrItemDataInOut(lngCol)) & ";") & strTmp
                        End If
                    
                    Case Else
                        If lngCol >= 0 And lngCol <= 41 And mBodyFlag.���� > 0 Then
                            If mBodyFlag.���� = 2 Then
                                strTmp = rsTmp("��Ŀ����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                            Else
                                strTmp = rsTmp("��Ŀ����").Value
                            End If
                            strTmp = strTmp & "'" & Format(rsTmp("ʱ��").Value, "HH:mm:ss")
                            
                            strArrItemDataInOut(lngCol) = IIf(Trim(strArrItemDataInOut(lngCol)) = "", "", Trim(strArrItemDataInOut(lngCol)) & ";") & strTmp
                        End If
                    End Select
                End If
                rsTmp.MoveNext
            Next
        End If
        
        '7����Forѭ�����ֱ������ǰ��Ŀ
        '����������ߵ����ݲ��� 'strArrItemInfo�ﱣ����������Ŀ����Ŀ��Ϣ
        '����ǰҳ��ʱ����ڣ�strTmpDay��(strTmpDay+6)�������ݰ���Ŀ���ζ�������
        '--------------------------------------------------------------------------------------------------------------
        Dim rsOffset As ADODB.Recordset
        
        Call InitOffset(rsOffset)
        
        For l = 0 To UBound(strArrItemInfo)
        
            dbl��� = Val(Split(strArrItemInfo(l), "'")(3))
            dbl��С = Val(Split(strArrItemInfo(l), "'")(4))
            dbl��λֵ = Val(Split(strArrItemInfo(l), "'")(5))
            lng����� = Val(Split(strArrItemInfo(l), "'")(1))
            lng��λ���� = (dbl��� - dbl��С) / dbl��λֵ
            lng��λ���� = IIf(lng��λ���� + lng����� > (MAXROWS - 1), (MAXROWS - 1) - lng�����, lng��λ���� + lng�����)
            
            '----------------------------------------------------------------------------------------------------------
            '��ȡ��������
            mstrSQL = "SELECT A.����ʱ�� As ʱ��,C.��¼���� As ��ֵ,c.��¼���,c.���²�λ,c.���Ժϸ� " & _
                        "FROM ���˻����¼ A,���˻������� C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
                        "Where A.ID=C.��¼ID " & _
                            "AND A.����id=[1] " & _
                            "AND A.��ҳid=[2]  And Nvl(a.Ӥ��,0)=[7] " & _
                            "AND D.��Ŀ���=C.��Ŀ��� " & _
                            "AND C.��¼����=1 " & _
                            "AND E.��Ŀ���=D.��Ŀ��� And (Nvl(E.Ӧ�÷�ʽ,0)=1 Or ([6]=-1 And E.Ӧ�÷�ʽ=2 And c.��¼���=1)) And Nvl(e.���ò���,0) In (0,[8]) " & _
                            "AND E.����ȼ�>=0  " & _
                            "AND a.����ʱ�� BETWEEN [3] And [4] And c.��ֹ�汾 Is Null And c.δ��˵�� Is Null " & _
                            "AND D.��¼��=1 AND D.��Ŀ��� In ([5],[6]) And c.��¼��� In ([9],[10]) " & _
                        "Order By a.����ʱ��,c.��¼���"
                        
            If mblnMoved Then
                mstrSQL = Replace(mstrSQL, "���˻����¼", "H���˻����¼")
                mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
            End If
            
            Select Case Val(Split(strArrItemInfo(l), "'")(7))
            Case 2
                
                If mint����Ӧ�� = 2 Then
                    '������Ŀ��Ҫ����������Ŀ
                    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng����id, lng��ҳid, CDate(Format(strBeginDate, "YYYY-MM-DD")), CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), Val(Split(strArrItemInfo(l), "'")(7)), -1, intBaby, IIf(intBaby = 0, 1, 2), 0, 1)
                Else
                    '���ʵ���Ӧ��ʱ��������Ŀ����Ҫ����������Ŀ
                    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng����id, lng��ҳid, CDate(Format(strBeginDate, "YYYY-MM-DD")), CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), Val(Split(strArrItemInfo(l), "'")(7)), 0, intBaby, IIf(intBaby = 0, 1, 2), 0, 0)
                End If
            Case -1
                Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng����id, lng��ҳid, CDate(Format(strBeginDate, "YYYY-MM-DD")), CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), Val(Split(strArrItemInfo(l), "'")(7)), 2, intBaby, IIf(intBaby = 0, 1, 2), 1, 1)
            Case Else
                Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", lng����id, lng��ҳid, CDate(Format(strBeginDate, "YYYY-MM-DD")), CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), Val(Split(strArrItemInfo(l), "'")(7)), 0, intBaby, IIf(intBaby = 0, 1, 2), 0, 1)
            End Select

            '�Ƚ�����ԭʼ���浽strArrNewTmp������
            If rsTemp.RecordCount > 0 Then
            
                rsTemp.MoveFirst
                
                ReDim strArrNewTmp(0)
                
                intCount = -1
                strSvrTime = ""
                For i = 0 To rsTemp.RecordCount - 1

                    lngCol = GetCurveColumn(CDate(rsTemp("ʱ��").Value), CDate(strBeginDate), mlngHourBegin) - 1
                    

                    strTime = Format(Int(CDate(strBeginDate)) + ((lngCol + 1) * 4 - (4 - mlngHourBegin)) / 24, "YYYY-MM-DD hh:mm:ss")
                    
                    If strSvrTime <> strTime Or zlCommFun.NVL(rsTemp("��¼���").Value, 0) = 1 Then
                        strSvrTime = strTime
                        
                        intCount = intCount + 1
                        ReDim Preserve strArrNewTmp(intCount)
                        
                        strArrNewTmp(intCount) = strTime & "'" & Trim(rsTemp!��ֵ) & ";" & Trim(zlCommFun.NVL(rsTemp("��¼���").Value)) & ";" & zlCommFun.NVL(rsTemp("���Ժϸ�").Value, 0) & ";" & rsTemp("ʱ��").Value & ";0;" & zlCommFun.NVL(rsTemp("���²�λ").Value)
                        
                    ElseIf strSvrTime <> strTime Then
                    
                        intCount = intCount + 1
                        ReDim Preserve strArrNewTmp(intCount)
                        strArrNewTmp(intCount) = strTime & "';"
                        
                    Else
                        
                        intCount = intCount + 1
                        ReDim Preserve strArrNewTmp(intCount)
                        strArrNewTmp(intCount) = strTime & "'" & Trim(rsTemp!��ֵ) & ";" & Trim(zlCommFun.NVL(rsTemp("��¼���").Value)) & ";" & zlCommFun.NVL(rsTemp("���Ժϸ�").Value, 0) & ";" & rsTemp("ʱ��").Value & ";1;" & zlCommFun.NVL(rsTemp("���²�λ").Value)
                        
                    End If
                    
                    rsTemp.MoveNext
                Next
                lngRowCount = rsTemp.RecordCount
            Else
                lngRowCount = 0
            End If

            '----------------------------------------------------------------------------------------------------------
            '�������˵����ʱ���
            
            mstrSQL = "SELECT c.��Ŀ��� as ItemNO,c.��¼����,A.����ʱ�� As ʱ��,Decode(c.��¼����,1,c.δ��˵��,c.��¼����) As ˵��,Decode(c.��¼����,1,1,c.��¼���) As ��¼��� " & _
                        "FROM ���˻����¼ A,���˻������� C " & _
                        "Where A.ID=C.��¼ID " & _
                            "AND A.����id=[1] " & _
                            "AND A.��ҳid=[2]  And Nvl(a.Ӥ��,0)=[5] " & _
                            "AND (c.��¼���� In (2,6) Or c.��¼����=1 And c.δ��˵�� Is Not Null) " & _
                            "AND a.����ʱ�� BETWEEN [3] And [4] And c.��ֹ�汾 Is Null " & _
                        "Order By a.����ʱ��,��¼����"
                        
            If mblnMoved Then
                mstrSQL = Replace(mstrSQL, "���˻����¼", "H���˻����¼")
                mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
            End If
            
            Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", _
                                                lng����id, _
                                                lng��ҳid, _
                                                CDate(Format(strBeginDate, "YYYY-MM-DD")), _
                                                CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), intBaby)
            If rsTemp.RecordCount > 0 Then
                '������˵����������
                rsTemp.MoveFirst
                ReDim strArrNewTmpComment(rsTemp.RecordCount - 1)
                For i = 0 To rsTemp.RecordCount - 1
                
                    lngCol = GetCurveColumn(CDate(rsTemp("ʱ��").Value), CDate(strBeginDate), mlngHourBegin) - 1
                    
                    'δ��˵��
                    If Val(zlCommFun.NVL(rsTemp!��¼����, 0)) = 1 And Val(zlCommFun.NVL(rsTemp!ItemNO, 0)) = 2 Then
                        '���㵱ǰҳ������
                        If Format(rsTemp("ʱ��").Value, "yyyy-MM-dd") >= strTmpDay And Format(rsTemp("ʱ��").Value, "yyyy-MM-dd") <= Format(DateAdd("d", 6, strTmpDay), "yyyy-MM-dd") Then
                            lngTmpDay = GetCurveColumn(CDate(rsTemp("ʱ��").Value), CDate(strTmpDay), mlngHourBegin) - 1
                            mbyt����(lngTmpDay) = 1
                        End If
                    End If

                    strTime = Format(Int(CDate(strBeginDate)) + ((lngCol + 1) * 4 - (4 - mlngHourBegin)) / 24, "YYYY-MM-DD hh:mm:ss")
                    
                    strArrNewTmpComment(i) = Format(strTime, "YYYY-MM-DD HH:MM:SS") & "'" & Trim(zlCommFun.NVL(rsTemp!˵��)) & "'" & zlCommFun.NVL(rsTemp!��¼����, 0) & "'" & zlCommFun.NVL(rsTemp!��¼���, 0)
                    
                    rsTemp.MoveNext
                Next
            End If
               
            For i = 0 To 41
                blnTag = False
                blnComment = False
                
                '�������ǰ�е�ʱ���
                strTmp = Format(Int(CDate(strTmpDay)) + ((i + 1) * 4 - (4 - mlngHourBegin)) / 24, "yyyy-MM-dd HH:mm:ss")

                '������������
                '------------------------------------------------------------------------------------------------------
                If lngRowCount > 0 Then
                    For j = 0 To UBound(strArrNewTmp)
                    
                        '�����ʱ��������д��ڵ�ǰʱ�䣬��ȡ��ʱ����������ֵ��������-2
                        '�����Ԫ��� 4
                        
                        If Trim(strArrNewTmp(j)) <> "" Then
                                                
                            If Format(Split(strArrNewTmp(j), "'")(0), "yyyy-MM-dd HH:mm:ss") = strTmp Then
                                '������¼��ǰ���ߵ�����

                                '���ͬһ�����ж��ֵ����ȡ��е��ֵ��Ϊ���е�ֵ
                                blnAllow = IsCenterValue(rsOffset, l, i, CDate(Split(Split(strArrNewTmp(j), "'")(1), ";")(3)), CDate(strTmp))
                                                            
                                '��¼ͬһʱ�������������
                                If blnAllow Then
                                    If strArrCurLineData(i) <> "" And strArrCurLineData(i) <> "-2" And Val(Split(Split(strArrNewTmp(j), "'")(1), ";")(4)) = 0 Then
                                        
                                        If Val(Split(strArrItemInfo(l), "'")(7)) = -1 And mint����Ӧ�� = 1 Then
                                            strArrCurLineData(i) = Split(strArrNewTmp(j), "'")(1)
                                        Else
                                            strArrCurLineData(i) = strArrCurLineData(i) & "," & Split(strArrNewTmp(j), "'")(1)
                                        End If
                                        
                                    Else
                                        strArrCurLineData(i) = Split(strArrNewTmp(j), "'")(1)
                                    End If
                                    blnTag = True
                                End If
                            End If
                        End If
                    Next
                End If
                
                '����˵������
                '------------------------------------------------------------------------------------------------------
                If rsTemp.RecordCount > 0 Then
                    For j = 0 To UBound(strArrNewTmpComment)
                        
                        If Format(Split(strArrNewTmpComment(j), "'")(0), "yyyy-MM-dd HH:mm:ss") = strTmp Then
                            '������¼��ǰ�������ݵ�˵��
                            
                            If strArrCurLineComment(i) = "" Then strArrCurLineComment(i) = ";;"
                            
                            aryTmp = Split(strArrCurLineComment(i), ";")
                            
                            Select Case Val(Split(strArrNewTmpComment(j), "'")(2))
                            Case 1
                                If bytδ����ʾλ�� = 0 Then
                                    If aryTmp(0) = "" Then
                                        aryTmp(0) = Split(strArrNewTmpComment(j), "'")(1)
                                    Else
                                        'If InStr(1, aryTmp(0) & " ", Split(strArrNewTmpComment(j), "'")(1) & " ") = 0 Then
                                            aryTmp(0) = aryTmp(0) & " " & Split(strArrNewTmpComment(j), "'")(1)
                                        'End If
                                    End If
                                Else
                                    If aryTmp(1) = "" Then
                                        aryTmp(1) = Split(strArrNewTmpComment(j), "'")(1)
                                    Else
                                        'If InStr(1, aryTmp(1) & " ", Split(strArrNewTmpComment(j), "'")(1) & " ") = 0 Then
                                            aryTmp(1) = aryTmp(1) & " " & Split(strArrNewTmpComment(j), "'")(1)
                                        'End If
                                    End If
                                End If
                            Case 6

                                If aryTmp(1) = "" Then
                                    aryTmp(1) = Split(strArrNewTmpComment(j), "'")(1)
                                Else
                                    aryTmp(1) = aryTmp(1) & " " & Split(strArrNewTmpComment(j), "'")(1)
                                End If
                                
                            Case Else
                                If aryTmp(0) = "" Then
                                    aryTmp(0) = Split(strArrNewTmpComment(j), "'")(1)
                                Else
                                    aryTmp(0) = aryTmp(0) & " " & Split(strArrNewTmpComment(j), "'")(1)
                                End If
                            End Select
                            
                            If Val(aryTmp(2)) = 0 Then aryTmp(2) = Split(strArrNewTmpComment(j), "'")(3)
                            
                            strArrCurLineComment(i) = Join(aryTmp, ";")
                            
                            blnComment = True
                        End If
                    Next
                End If
                '���û�о�дĬ��ֵ
                If Not blnTag Then strArrCurLineData(i) = "-2"
                If Not blnComment Then strArrCurLineComment(i) = ""
            Next
            strArrItemDataInfo(l) = Join(strArrCurLineData, "'")
            strArrItemDataComment(l) = Join(strArrCurLineComment, "'")
            
            '�����������
            For j = 0 To UBound(strArrCurLineData)
                strArrCurLineData(j) = ""
            Next
        Next
        '��ʱ�Ѿ���ȡ�������������Ŀ��������
        
        '�����һ�������м��������ַ����б��ʾ��ʾ��Щ��Ŀ
        Call DrawBodyGraph(objDraw, X, Y, strArrItemInfo, strArrItemDataInfo, strArrItemDataComment, strArrItemDataInOut, "")
        
        '8�������¼��Ŀ��¼�б����飬���¶� X ��Y �ȱ�����ֵ
        X = lngNewTmpX
        Y = lngNewTmpY - H_9pt / 2 - 30    '����/�����ǩ�еĸ߶�=H_9pt * 2
        
        
        '�ڣ����ݣ�����������----------------------------------------------------------------------------------------------------------

        ReDim strArrNewTmp(0)
        Dim varNewTmpString As Variant
        Dim intCol As Integer
        Dim intColTmp As Integer
        Dim intColFirst1 As Integer
        Dim intColFirst2 As Integer
        
        intColFirst1 = 0
        intColFirst2 = 0
        
        If intDrawGridRows > 0 Then
            
            ReDim strArrNewTmp(intDrawGridRows - 1) '��ʼ����ʱ����׼����ȡ����
            
            'Ҫ���±�¼����Ŀ��������ѭ��
            For i = LBound(strArrRecordItemInfo) To UBound(strArrRecordItemInfo)
                '�˴�������ȡ����¼����Ŀ��������
                
                mstrSQL = "SELECT A.����ʱ�� As ��ʱ������ʱ��,C.��¼���� As ˵��,E.������Ŀ,D.��¼��,D.��Ŀ���,D.��¼Ƶ�� " & _
                            "FROM ���˻����¼ A,���˻������� C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
                            "Where A.ID=C.��¼id " & _
                                "AND A.����id=[1] " & _
                                "AND A.��ҳid=[2]  And Nvl(a.Ӥ��,0)=[7] " & _
                                "AND D.��Ŀ���=C.��Ŀ��� " & _
                                "AND C.��¼����=1 " & _
                                "AND E.��Ŀ���=D.��Ŀ��� And Nvl(E.Ӧ�÷�ʽ,0)=1 And Nvl(e.���ò���,0) In (0,[8]) " & _
                                "AND E.����ȼ�>=0  " & _
                                "AND A.����ʱ�� BETWEEN [3] And [4] And c.��ֹ�汾 Is Null " & _
                                "AND D.��¼��=2 AND D.��¼�� In ([5],[6]) " & _
                            "Order By A.����ʱ��,Decode(D.��¼��,'����ѹ',0,1)"
                            
                If mblnMoved Then
                    mstrSQL = Replace(mstrSQL, "���˻����¼", "H���˻����¼")
                    mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
                End If
                
                If CStr(Split(strArrRecordItemInfo(i), "'")(0)) = "Ѫѹ" Then
                    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", _
                                                        lng����id, _
                                                        lng��ҳid, _
                                                        CDate(Format(strBeginDate, "YYYY-MM-DD")), _
                                                        CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), _
                                                        "����ѹ", "����ѹ", intBaby, IIf(intBaby = 0, 1, 2))
                Else
                    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", _
                                                        lng����id, _
                                                        lng��ҳid, _
                                                        CDate(Format(strBeginDate, "YYYY-MM-DD")), _
                                                        CDate(Format(strEndDate, "YYYY-MM-DD") & " 23:59:59"), _
                                                        CStr(Split(strArrRecordItemInfo(i), "'")(0)), "", intBaby, IIf(intBaby = 0, 1, 2))
                End If
                
                If CStr(Split(strArrRecordItemInfo(i), "'")(0)) = "����" Then
                     '������42��
                     strNewTmpString = String(42, ";")
                 Else
                     strNewTmpString = String(14, ";")
                 End If
                
                Dim sgl�ų���(1 To 14) As Single
                Dim sgl������(1 To 14) As Single
                
                If rsTemp.BOF = False Then
                        
                    varNewTmpString = Split(strNewTmpString, ";")
                    
                    Do While Not rsTemp.EOF
                        
                        If CStr(Split(strArrRecordItemInfo(i), "'")(0)) = "����" Then
                            
                            intCol = GetCurveColumn(CDate(rsTemp("��ʱ������ʱ��").Value), CDate(strTmpDay), mlngHourBegin)
                            
                            If intCol >= LBound(varNewTmpString) And intCol <= UBound(varNewTmpString) Then
                                varNewTmpString(intCol) = zlCommFun.NVL(rsTemp!˵��)
                            End If
                        Else
                            
                            If Int((rsTemp!��ʱ������ʱ�� - Int(CDate(strTmpDay))) * 24) < 0 Then
                                intCol = 0
                            Else
                                intCol = 1 + Int((rsTemp!��ʱ������ʱ�� - Int(CDate(strTmpDay))) * 24) \ 12
                            End If
                            
                            If intCol >= 1 And intCol <= 14 Then
                                
                               Select Case rsTemp!��Ŀ���
                                Case 7
                                    
                                    If rsTemp!��¼Ƶ�� = 1 Then
                                        intColTmp = IIf(intCol Mod 2 = 0, intCol, intCol + 1)
                                    Else
                                        intColTmp = intCol
                                    End If
                                    
                                    sgl������(intColTmp) = sgl������(intColTmp) + Val(zlCommFun.NVL(rsTemp!˵��))
                                    
                                    If Right(zlCommFun.NVL(rsTemp!˵��), 2) = "/C" Then
                                    
                                        varNewTmpString(intCol) = sgl������(intColTmp) & "/C"
                                        
                                    ElseIf Right(zlCommFun.NVL(rsTemp!˵��), 1) = "C" Then
                                        varNewTmpString(intColTmp) = "C"
                                    Else
                                        varNewTmpString(intColTmp) = sgl������(intColTmp)
                                    End If
                                Case 9
                                    
                                    If rsTemp!��¼Ƶ�� = 1 Then
                                        intColTmp = IIf(intCol Mod 2 = 0, intCol, intCol + 1)
                                    Else
                                        intColTmp = intCol
                                    End If
                                    
                                    sgl�ų���(intColTmp) = sgl�ų���(intColTmp) + Val(zlCommFun.NVL(rsTemp!˵��))
                                    If Right(zlCommFun.NVL(rsTemp!˵��), 2) = "/C" Then
                                    
                                        varNewTmpString(intColTmp) = sgl�ų���(intColTmp) & "/C"
                                        
                                    ElseIf Right(zlCommFun.NVL(rsTemp!˵��), 1) = "C" Then
                                        varNewTmpString(intColTmp) = "C"
                                    Else
                                        varNewTmpString(intColTmp) = sgl�ų���(intColTmp)
                                    End If
                                    
                                Case Else
                                    'varPatiInfo(3)
                                    
                                    Select Case rsTemp("��¼��").Value
                                    Case "����ѹ"
                                        
                                        If Format(rsTemp!��ʱ������ʱ��, "yyyy-MM-dd") = Format(varPatiInfo(3), "yyyy-MM-dd") Then
                                            If intColFirst1 = 0 Then
                                                varNewTmpString(intCol) = zlCommFun.NVL(rsTemp!˵��)
                                                intColFirst1 = intCol
                                            ElseIf intColFirst1 <> intCol Then
                                                varNewTmpString(intCol) = zlCommFun.NVL(rsTemp!˵��)
                                            End If
                                        Else
                                            intColFirst1 = intCol
                                            varNewTmpString(intCol) = zlCommFun.NVL(rsTemp!˵��)
                                        End If
                                        
'                                        varNewTmpString(intCol) = rsTemp!˵��
                                    Case "����ѹ"
                                        
                                        If Format(rsTemp!��ʱ������ʱ��, "yyyy-MM-dd") = Format(varPatiInfo(3), "yyyy-MM-dd") Then
                                            
                                            If intColFirst2 = 0 Then
                                                If InStr(varNewTmpString(intCol), "/") > 0 Then
                                                    varNewTmpString(intCol) = varNewTmpString(intCol) & zlCommFun.NVL(rsTemp!˵��)
                                                Else
                                                    varNewTmpString(intCol) = varNewTmpString(intCol) & "/" & zlCommFun.NVL(rsTemp!˵��)
                                                End If

                                                intColFirst2 = intCol
                                            ElseIf intColFirst2 <> intCol Then
                                                If InStr(varNewTmpString(intCol), "/") > 0 Then
                                                    varNewTmpString(intCol) = varNewTmpString(intCol) & zlCommFun.NVL(rsTemp!˵��)
                                                Else
                                                    varNewTmpString(intCol) = varNewTmpString(intCol) & "/" & zlCommFun.NVL(rsTemp!˵��)
                                                End If
                                            End If
                                            
                                        Else
                                            intColFirst2 = intCol
                                            If InStr(varNewTmpString(intCol), "/") > 0 Then
                                                varNewTmpString(intCol) = varNewTmpString(intCol) & zlCommFun.NVL(rsTemp!˵��)
                                            Else
                                                varNewTmpString(intCol) = varNewTmpString(intCol) & "/" & zlCommFun.NVL(rsTemp!˵��)
                                            End If
                                        End If
                                        

                                        If varNewTmpString(intCol) = "/" Then varNewTmpString(intCol) = ""
                                    Case Else
                                        varNewTmpString(intCol) = zlCommFun.NVL(rsTemp!˵��)
                                    End Select
                                    
                                End Select
                            End If
                        End If
                        
                        rsTemp.MoveNext
                    Loop
                    
                    strNewTmpString = Join(varNewTmpString, ";")
                    
                End If
                
                strArrNewTmp(i) = strNewTmpString
            Next
            
            
            For j = LBound(strArrRecordItemInfo) To UBound(strArrRecordItemInfo)
                If Split(strArrRecordItemInfo(j), "'")(1) <> "" Then
                    strArrNewTmp(j) = Split(strArrRecordItemInfo(j), "'")(0) & ";" & Split(strArrRecordItemInfo(j), "'")(2) & ";(" & Split(strArrRecordItemInfo(j), "'")(1) & ")" & strArrNewTmp(j)
                Else
                    strArrNewTmp(j) = Split(strArrRecordItemInfo(j), "'")(0) & ";" & Split(strArrRecordItemInfo(j), "'")(2) & ";" & strArrNewTmp(j)
                End If
            Next
            
        End If
        
'        y = y - 90
        
        intRepairRows = zlDatabase.GetPara("���±������", glngSys, 1255, 8)
        
        Call DrawBodyRecordItem(objDraw, X, Y, mlngFirstWidth, strArrNewTmp, rsArrRecordItemInfo, strTmpString1, strTmpString2, sngScale, intRepairRows)
        
        '9��������¼��˵��
        Call DrawBodyTips(objDraw, lngNewTmpX, Y, mlngFirstWidth, strStateTips, sngScale)
        
        '10����󻭳�ҳ�룬Ȼ�������һҳ
        Call DrawBodyPageFooter(objDraw, X, Y, intPageNo, intEndPage)
        
NOPageSub:                                 Next    '����ÿһҳ��ѭ��
        If blnPrint = False Then
            '����Ǵ�ӡԤ��,Ӧ����ӡ���Ŀɴ�ӡ�Ŀ�ʼ����ʼԤ��
            dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
            dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
            objDraw.DrawStyle = 2
            On Error Resume Next
            objDraw.Line (objDraw.Width * dblSureW, objDraw.Height * dblSureH)-(objDraw.Width * (1 - dblSureW) - Screen.TwipsPerPixelX * sngScale, objDraw.Height * (1 - dblSureH) - Screen.TwipsPerPixelY * sngScale), &H808080, B
        End If
        Call ShowFlash
        
        Screen.MousePointer = 0
        PrintOrPreviewBodyState = True
        Exit Function
ErrPrint:
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
ErrExit:
        Call ShowFlash
        Screen.MousePointer = 0
        Err.Clear
        PrintOrPreviewBodyState = False
End Function

Public Function Ceil(ByVal dbValue As Double) As Integer
    '******************************************************************************************************************
    '���ܣ� ת��ʱ��Ϊ��ֵ
    '������
    '���أ�
    '******************************************************************************************************************
    
    Ceil = (0 - Int(0 - dbValue))
    
End Function

'Private Function ConvertTimeToCol(ByVal strFrom As String, ByVal dtDate As Date, ByVal lngHourBegin As Long) As Long
'    '******************************************************************************************************************
'    '���ܣ� ת��ʱ��Ϊ��ֵ
'    '������
'    '���أ�
'    '******************************************************************************************************************
'
'    strFrom = Int(CDate(strFrom)) - (4 - lngHourBegin) / 24
'    ConvertTimeToCol = Int(DateDiff("h", CDate(strFrom), dtDate) / 4)
'
'End Function

Public Function InitDateTimeRange(ByRef varTime As Variant, Optional ByVal intHourBegin As Integer = 4) As Boolean
    '******************************************************************************************************************
    '���ܣ��������µ�һ�������ʱ�䷶Χ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim strTmpBegin As String
    Dim strTmpEnd As String
    
    On Error GoTo errHand
    
    varTime = Split(String(6, ";"), ";")
    
    varTime(0) = "00:00:00," & Format(intHourBegin + 1, "00") & ":59:59"
    varTime(5) = (intHourBegin + 18) & ":00:00,23:59:59"
    
    For intLoop = 1 To 4
        strTmpBegin = Format(DateAdd("s", 1, CDate("2000-01-01 " & Split(varTime(intLoop - 1), ",")(1))), "HH:mm:ss")
        strTmpEnd = Format(DateAdd("h", 4, CDate("2000-01-01 " & Split(varTime(intLoop - 1), ",")(1))), "HH:mm:ss")
        varTime(intLoop) = strTmpBegin & "," & strTmpEnd
    Next
    
    InitDateTimeRange = True
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
        
        
End Function

Public Function GetCurveColumn(ByVal dtDateTime As Date, ByVal dtBeginDateTime As Date, Optional ByVal intHourBegin As Integer = 4) As Integer
    '******************************************************************************************************************
    '���ܣ� ��ʱ��������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim varTime As Variant
    Dim strTmp As String
    Dim intDays As Integer
    Dim intLoop As Integer
    
    On Error GoTo errHand
    
    GetCurveColumn = -1
    
    '��ʼ��ʱ�䷶Χ����
    Call InitDateTimeRange(varTime, intHourBegin)

    '���㵱ǰ���ʱ������һ��ĵڼ���λ����
    strTmp = Format(dtDateTime, "HH:mm:ss")
    For intLoop = 0 To 6
        If strTmp >= Split(varTime(intLoop), ",")(0) And strTmp <= Split(varTime(intLoop), ",")(1) Then
            Exit For
        End If
    Next
    If intLoop < 7 Then
        
        '���㵱���ڵ�ǰ���µ�ҳ���ǵڼ��죨0��ʾ��һ�죻1��ʾ�ڶ���.....��
        intDays = DateDiff("d", Int(dtBeginDateTime), Int(dtDateTime))
        GetCurveColumn = intDays * 6 + intLoop + 1
    
    End If
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
            
End Function

Public Function GetCurveDateTime(ByVal intCol As Integer, ByVal dtBeginDateTime As Date, Optional ByVal intHourBegin As Integer = 4) As String
    '******************************************************************************************************************
    '���ܣ� �������ʱ��
    '������
    '���أ����ص�ǰ�е�ʱ�䷶Χ����ʽΪ��"��ʼʱ��,����ʱ��"
    '******************************************************************************************************************
    Dim varTime As Variant
    Dim strTmp As String
    Dim intDays As Integer
    Dim strDay As String
    
    On Error GoTo errHand
    
    GetCurveDateTime = ""
    
    '��ʼ��ʱ�䷶Χ����
    Call InitDateTimeRange(varTime, intHourBegin)
        
    intDays = intCol \ 6
    intCol = (intCol Mod 6)
    If intCol = 0 Then
        intCol = 6
        If intDays >= 1 Then intDays = intDays - 1
    End If
    
    '�������ڵ�����
    If intCol >= 1 And intCol <= 6 Then
        strDay = Format(DateAdd("d", intDays, Int(dtBeginDateTime)), "yyyy-MM-dd")
        strTmp = strDay & " " & Split(varTime(intCol - 1), ",")(0)
        strTmp = strTmp & "," & strDay & " " & Split(varTime(intCol - 1), ",")(1)
    End If
    
    '����ʱ��
    GetCurveDateTime = strTmp
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Public Function GetEditDateTime(ByVal intCol As Integer, ByVal dtBeginDateTime As Date) As String
    '******************************************************************************************************************
    '���ܣ� �������ʱ��
    '������
    '���أ����ص�ǰ�е�ʱ�䷶Χ����ʽΪ��"��ʼʱ��,����ʱ��"
    '******************************************************************************************************************
    Dim strTmp As String
    Dim intDays As Integer
    Dim strDay As String
    
    On Error GoTo errHand
    
    GetEditDateTime = ""
        
    intDays = intCol \ 2
    intCol = (intCol Mod 2)
    
    If intCol = 0 Then
        If intDays >= 1 Then intDays = intDays - 1
    End If
    
    strDay = Format(DateAdd("d", intDays, Int(dtBeginDateTime)), "yyyy-MM-dd")
    If intCol = 1 Then
        strTmp = strDay & " 00:00:00"
        strTmp = strTmp & "," & strDay & " 11:59:59"
    ElseIf intCol = 0 Then
        strTmp = strDay & " 12:00:00"
        strTmp = strTmp & "," & strDay & " 23:59:59"
    End If
    
    GetEditDateTime = strTmp

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

'=========================================
Public Sub ClearSpecRowCol(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
    '����: ���ָ�������ָ����ָ���е�����
    '����: obj=Ҫ����������ؼ�
    '      intRow=Ҫ������к�
    '      intCol=Ҫ������к��б���Array(1,2,3),������������Ա�ʾΪArray()
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Public Sub SetColumnText(fgd As Object, intRow As Integer, ByVal varColText As Variant)
    '����: ����ָ������ؼ�����ͷ�ı�
    '����: fgd=����ؼ�
    '      intRow=�к�
    '      varColText=��ͷ�ı�����
    Dim i As Integer
    For i = 0 To fgd.Cols - 1
        fgd.TextMatrix(intRow, i) = varColText(i)
    Next
End Sub

Public Sub SetColAlignment(fgd As Object, varColAlignment As Variant)
    '����: ����ָ������ؼ����ж��뷽ʽ
    '����: fgd=����ؼ�
    '      varColAlignment=�ж��뷽ʽ����
    Dim i As Long
    For i = 0 To UBound(varColAlignment)
        fgd.ColAlignment(i) = varColAlignment(i)
    Next
End Sub

Public Sub SetColData(fgd As Object, varColData As Variant)
    '����: ����ָ������ؼ�����������Դ��ʽ
    '����: fgd=����ؼ�
    '      varColData=��������Դ��ʽ����
    Dim i As Long
    For i = 0 To UBound(varColData)
        fgd.ColData(i) = varColData(i)
    Next
End Sub

Public Sub SetFixColAlignment(fgd As Object, varFixColAlignment As Variant)
    '����: ����ָ������ؼ��Ĺ̶��ж��뷽ʽ
    '����: fgd=����ؼ�
    '      varColAlignment=�̶��ж��뷽ʽ����
    Dim i As Long
    For i = 0 To UBound(varFixColAlignment)
        fgd.ColAlignmentFixed(i) = varFixColAlignment(i)
    Next
End Sub

Public Sub SetColumnWidth(fgd As Object, ByVal varColWidth As Variant)
    '����: ����ָ������ؼ����п�
    '����: fgd=����ؼ�
    '      varColWidth=�п�����
    Dim i As Integer
    For i = 0 To fgd.Cols - 1
        fgd.ColWidth(i) = varColWidth(i)
    Next
End Sub

Public Sub SetRowForeColor(mshObject As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    Dim i As Integer
    Dim blnPre As Boolean
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshObject
        blnPre = .Redraw
        intRow = .Row
        intCol = .Col
        .Redraw = False
        .Row = lngRow
        For i = 0 To .Cols - 1
            .Col = i
            .CellForeColor = lngColor
        Next
        
        .Row = intRow
        .Col = intCol
        .Redraw = blnPre
    End With
End Sub

Public Sub CalcXY(objFrm As Object, objMSH As Object, objX As Single, objY As Single, sglX As Single, sglY As Single)
    sglX = objFrm.Left + objX + objMSH.CellLeft + Screen.TwipsPerPixelX
    sglY = objFrm.Top + objFrm.Height - objFrm.ScaleHeight + objY + objMSH.CellTop + objMSH.CellHeight
    If sglX + 5895 > Screen.Width Then
        sglX = Screen.Width - 5895
    End If
    If sglY + 3420 > Screen.Height Then
        sglY = sglY - objMSH.CellHeight - 3420
    End If
End Sub


Public Function Check�Ƿ����(strSource As String, strTarge As String) As Boolean
    '���strSource�е�ÿһ���ַ��Ƿ���strTarge��
    Dim i As Long
    Check�Ƿ���� = False
    
    Select Case strTarge
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "������"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "��С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    Check�Ƿ���� = True
End Function

Public Sub SelectRow(mshObject As Object, Optional ByVal BackColor As Long = &H8000000D, Optional ByVal ForeColor As Long = &H8000000E)
    Dim i As Integer
    Dim blnPre As Boolean
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshObject
        blnPre = .Redraw
        intRow = .Row
        intCol = .Col
        .Redraw = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = BackColor
            .CellForeColor = ForeColor
        Next
        
        .Row = intRow
        .Col = intCol
        .Redraw = blnPre
    End With
End Sub

Public Sub UnSelectRow(mshObject As Object, Optional lngColorSave As Long = 0)
    Dim i As Integer
    Dim blnPre As Boolean
    
    With mshObject
        blnPre = .Redraw
        .Redraw = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = .BackColor
            .CellForeColor = lngColorSave
        Next
        .Redraw = blnPre
    End With
End Sub


Public Sub Hook(ByVal hWnd As Long)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    glngPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)

    '��ȡ"�������"�еĹ�������ֵ

    Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, WHEEL_SCROLL_LINES, 0)

    If WHEEL_SCROLL_LINES > frmCaseTendBody.BodyEdit.ScrollBarY.Max Then WHEEL_SCROLL_LINES = frmCaseTendBody.BodyEdit.ScrollBarY.Max
End Sub

Public Sub UnHook(ByVal hWnd As Long)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngReturnValue As Long

    lngReturnValue = SetWindowLong(hWnd, GWL_WNDPROC, glngPrevWndProc)
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '******************************************************************************************************************
    '���ܣ�����ϵͳ�¼������д���
    '������
    '���أ�
    '******************************************************************************************************************
    Dim pt As POINTAPI
    Dim wzDelta
    Dim wKeys As Integer
    
    Select Case uMsg
    Case WM_MOUSEWHEEL                          '�����¼�
        wzDelta = HIWORD(wParam)
        wKeys = LOWORD(wParam)
        pt.X = LOWORD(lParam)
        pt.Y = HIWORD(lParam)
    
        '����Ļ����ת��ΪfrmCaseTendBody��������
    
        ScreenToClient frmCaseTendBody.hWnd, pt

        With frmCaseTendBody.BodyEdit
        
            '�ж������Ƿ���frmCaseTendBody.BodyEdit������
    
            If pt.X > .Left / Screen.TwipsPerPixelX And pt.X < (.Left + .Width) / Screen.TwipsPerPixelX And pt.Y > .Top / Screen.TwipsPerPixelY And pt.Y < (.Top + .Height) / Screen.TwipsPerPixelY Then
    
                If wKeys = 16 Then
                    'ˮƽ����
                    
                Else
                    '��ֱ����
                    If Sgn(wzDelta) = 1 Then
                        .ScrollBarY.Value = IIf(.ScrollBarY.Value - WHEEL_SCROLL_LINES < .ScrollBarY.Min, .ScrollBarY.Min, .ScrollBarY.Value = .ScrollBarY.Value - WHEEL_SCROLL_LINES)
                    Else
                        .ScrollBarY.Value = IIf(.ScrollBarY.Value + WHEEL_SCROLL_LINES > .ScrollBarY.Max, .ScrollBarY.Max, .ScrollBarY.Value + WHEEL_SCROLL_LINES)
                    End If
                End If
            End If
        End With
    Case Else                                   '�����¼�����ϵͳȱʡ����
        WindowProc = CallWindowProc(glngPrevWndProc, hw, uMsg, wParam, lParam)
    End Select
End Function

Public Function ShowTxtSelDialog(ByVal frmParent As Object, _
                                    ByVal objTXT As Object, _
                                    ByVal strLvw As String, _
                                    ByVal strSavePath As String, _
                                    ByVal strDescrible As String, _
                                    ByVal rsData As ADODB.Recordset, _
                                    ByRef rsResult As ADODB.Recordset, _
                                    Optional ByVal lngCX As Long = 9000, _
                                    Optional ByVal lngCY As Long = 4500, _
                                    Optional blnMuliSel As Boolean = False, _
                                    Optional strInitKey As String = "", _
                                    Optional ByVal WinStyle As Byte = 3, _
                                    Optional ByVal blnSort As Boolean = True) As Boolean
    '******************************************************************************************************************
    '����:������+�б�ṹ
    '����:������2;�ɹ�����1;ȡ������0
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim objPoint As POINTAPI
    Dim lngObjHeight As Long
    
    On Error GoTo errHand
    
    If rsData.BOF Then Exit Function
    
    If objTXT Is Nothing Then
        '��Ļ����
        
        lngX = (Screen.Width - lngCX) / 2
        lngY = (Screen.Height - lngCY) / 2
        
    Else
        Call ClientToScreen(objTXT.hWnd, objPoint)
                    
        lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
        lngY = objTXT.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        
        lngObjHeight = objTXT.Height
    End If
    
    If frmSelectDialog.ShowSelect(frmParent, WinStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, , False, blnMuliSel, , blnSort) Then
                            
        Set rsResult = rsData
        ShowTxtSelDialog = True
        
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "Trans", adTinyInt                   '1��ʾ��ʼ;2��ʾ����
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = gstrSysName
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
            
            If Val(rs("Custom").Value) = 0 Then
                strSQL = CStr(rs("SQL").Value)
                Call zlDatabase.ExecuteProcedure(strSQL, strTitle)
            End If
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
End Function

Public Function SQLRecordSavePicture(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String
    Dim aryTmp As Variant
    Dim strFile As String
    
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = gstrSysName
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
            
            If Val(rs("Custom").Value) = 100 Then
                
                strTmp = rs("Parameter").Value
                aryTmp = Split(strTmp, ";")
                
                If UBound(aryTmp) >= 2 Then
                    If Dir(CStr(aryTmp(2))) <> "" And CStr(aryTmp(2)) <> "" Then
                        
                        Call zlBlobSave(Val(aryTmp(0)), Val(aryTmp(1)), CStr(aryTmp(2)))

                    End If
                End If
                
            End If
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordSavePicture = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
End Function

Public Function DrawPicture(objDraw As Object, ByVal strFile As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional ByVal bln��Դ As Boolean = False) As Boolean
    '******************************************************************************************************************
    '���ܣ���������С�Զ��ȱ���������Ƭ�ļ�
    '����������ǰ����Ƭ�ļ�
    '���أ����ź����Ƭ�ļ�
    '******************************************************************************************************************
    Dim strTmp As String
    Dim objMap As StdPicture
    Dim W As Single
    Dim H As Single
    Dim sglPerW As Single
    Dim sglPerH As Single
    Dim sglPer As Single
    Dim cx As Long
    Dim cy As Long
    
    On Error GoTo errHand
    
    If strFile = "" Then Exit Function
    
    cx = X2 - X1
    cy = Y2 - Y1
    
    If bln��Դ Then
        Set objMap = VB.LoadResPicture(strFile, vbResBitmap)
    Else
        Set objMap = VB.LoadPicture(strFile)
    End If
    
    W = objMap.Width * 0.566950910348006
    H = objMap.Height * 0.566950910348006
    
    If W > cx Then sglPerW = cx / W
    If H > cy Then sglPerH = cy / H
    
    If W > cx Or H > cy Then
        sglPer = IIf(sglPerW > sglPerH, sglPerH, sglPerW)
        W = W * sglPer
        H = H * sglPer
    End If
                    
    objDraw.PaintPicture objMap, X1, Y1, W, H
'    objDraw.PaintPicture objMap, X1, Y1, W, H, vbSrcAnd
    
    DrawPicture = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If bln��Դ Then
        ShowSimpleMsg "δ�ҵ�����Դ(" & strFile & "),���ܸ���Դ������!"
    Else
        ShowSimpleMsg "���ܴ��ļ�(" & strFile & "),���ļ���������ʹ�û��ļ�������!"
    End If
End Function

'######################################################################################################################

Public Function GetGridItem(ByVal byt����ȼ� As Byte, ByVal lng����id As Long, ByVal bytӤ�� As Byte, Optional ByVal byt��Ŀ���� As Byte = 1, Optional ByVal strNotItem As String, Optional ByVal blnBodyItem As Boolean = True) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ�
    '������strNotItem�����1,���2,���3,......
    '���أ�
    '******************************************************************************************************************
    
    If blnBodyItem Then
        mstrSQL = " Select A.�������,A.��¼Ƶ��,A.��¼��,c.��Ŀ����,c.��Ŀ���,c.��Ŀ����,c.��Ŀ��ʾ,c.��Ŀֵ��,c.��Ŀ��λ,A.��¼�� as ����,A.��Ŀ��� as ��Ŀ��,A.��Ŀ��� As ID,Nvl(B.ID,0) as ��ĿID,c.��Ŀ����,C.������Ŀ,1 As ĩ��," & _
                    " C.��Ŀ��λ As ��λ,��¼��,��Сֵ,���ֵ,��¼ɫ,1 as ��¼��,��λֵ,�����,Nvl(C.��Ŀ����,1) as �洢����,c.��Ŀ����,c.��ĿС��,c.������ " & _
                    " From ���¼�¼��Ŀ A,����������Ŀ B,�����¼��Ŀ C " & _
                    " Where C.��ĿID=B.ID(+) And A.��Ŀ���=C.��Ŀ��� And A.��¼��=2 And c.��Ŀ����=[4] And Nvl(C.Ӧ�÷�ʽ,0)=1 And C.����ȼ�>=[1] And Nvl(C.���ò���,0) In (0,[3]) " & _
                    " And (c.���ÿ���=1 Or (c.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=c.��Ŀ��� And D.����id=[2]))) "
                        
        If strNotItem <> "" Then mstrSQL = mstrSQL & " And A.��Ŀ��� Not In (" & strNotItem & ")"
        mstrSQL = mstrSQL & " Order by A.�������"
    Else
    
        mstrSQL = " Select c.��Ŀ����,c.��Ŀ���,c.��Ŀ����,c.��Ŀ��ʾ,c.��Ŀֵ��,c.��Ŀ��λ,c.��Ŀ���� as ����,c.��Ŀ��� as ��Ŀ��,c.��Ŀ��� As ID,Nvl(B.ID,0) as ��ĿID,c.��Ŀ����,C.������Ŀ,1 As ĩ��," & _
                    " C.��Ŀ��λ As ��λ,1 as ��¼��,Nvl(C.��Ŀ����,1) as �洢����,c.��Ŀ����,c.��ĿС��,c.������ " & _
                    " From ����������Ŀ B,�����¼��Ŀ C " & _
                    " Where C.��ĿID=B.ID(+) And c.��Ŀ����=[4] And Nvl(C.Ӧ�÷�ʽ,0)=1 And C.����ȼ�>=[1] And Nvl(C.���ò���,0) In (0,[3]) " & _
                    " And (c.���ÿ���=1 Or (c.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=c.��Ŀ��� And D.����id=[2]))) "
                        
        If strNotItem <> "" Then mstrSQL = mstrSQL & " And c.��Ŀ��� Not In (" & strNotItem & ")"
    
    
    End If
    
    Set GetGridItem = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", byt����ȼ�, lng����id, bytӤ��, byt��Ŀ����)
    
End Function

Public Function GetGridDataItem(ByVal byt����ȼ� As Byte, _
                                ByVal lng����id As Long, _
                                ByVal byt���ò��� As Byte, _
                                ByVal lng����id As Long, _
                                ByVal lng��ҳid As Long, _
                                ByVal dt��ʼʱ�� As Date, _
                                ByVal dt����ʱ�� As Date, ByVal bytӤ�� As Byte, Optional ByVal blnMoved As Boolean) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ���ȡ�����ݵĻ������Ŀ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    
    
    mstrSQL = " Select A.�������,A.��¼Ƶ��,A.��¼��,A.��¼�� as ����,A.��Ŀ��� as ��Ŀ��,A.��Ŀ��� As ID,Nvl(B.ID,0) as ��ĿID,c.��Ŀ����,C.������Ŀ,1 As ĩ��," & _
                " C.��Ŀ��λ As ��λ,��¼��,��Сֵ,���ֵ,��¼ɫ,1 as ��¼��,��λֵ,�����,Nvl(C.��Ŀ����,1) as �洢����,c.��Ŀ����,c.��ĿС��,c.������ " & _
                " From ���¼�¼��Ŀ A,����������Ŀ B,�����¼��Ŀ C " & _
                " Where C.��ĿID=B.ID(+) And A.��Ŀ���=C.��Ŀ��� And A.��¼��=2 And c.��Ŀ����=2 And Nvl(C.Ӧ�÷�ʽ,0)=1 And C.����ȼ�>=[1] And Nvl(C.���ò���,0) In (0,[3]) " & _
                " And (c.���ÿ���=1 Or (c.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=c.��Ŀ��� And D.����id=[2])))"
                
    strSQL = "Select e.��Ŀ��� " & _
                "FROM ���˻����¼ A,���˻������� C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
                "Where A.ID = c.��¼ID " & _
                    "AND A.������Դ=2 " & _
                    "AND Nvl(a.Ӥ��,0)=[8] " & _
                    "AND a.����id=[4] " & _
                    "AND a.��ҳid=[5] " & _
                    "AND d.��Ŀ���=C.��Ŀ��� " & _
                    "AND c.��¼����=1 And E.��Ŀ����=2 " & _
                    "AND E.��Ŀ���=D.��Ŀ��� " & _
                    "AND E.����ȼ�>=[1]  " & _
                    "AND a.����ʱ�� BETWEEN [6] And [7] And c.��ֹ�汾 Is Null " & _
                    "AND d.��¼��=2"
                                            
    If blnMoved Then
        strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
        strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
    End If
    
    mstrSQL = mstrSQL & " And c.��Ŀ��� In (" & strSQL & ")"
    mstrSQL = mstrSQL & " Order by A.�������"
    
    Set GetGridDataItem = zlDatabase.OpenSQLRecord(mstrSQL, "mdlPrint", byt����ȼ�, lng����id, byt���ò���, lng����id, lng��ҳid, dt��ʼʱ��, dt����ʱ��, bytӤ��)
    
End Function

Public Function ExistsPrinter() As Boolean
    Dim lngHDc As Long
    
    If Printers.Count = 0 Then Exit Function
    
    On Error Resume Next
    lngHDc = Printer.hDC
    If Err.Number = 0 Then ExistsPrinter = True
    Err.Clear: On Error GoTo 0
End Function

Private Function GetFontSize(ByVal objDraw As Object, ByVal dblHeight As Double, ByVal strText As String, ByRef Y1 As Single) As Single
    Dim sinFontSize As Single
    Dim sinFontSize_Bak As Single
    Dim intCharNumber As Integer
    Dim intCount As Integer
    Dim strChar As String
    '�������������С
    
    sinFontSize_Bak = objDraw.FontSize
    For sinFontSize = objDraw.FontSize To 5 Step -1
        Y1 = 0
        intCharNumber = 0
        For intCount = 1 To Len(strText)
            strChar = Mid(strText, intCount, 1)
            
            If Asc(strChar) < 0 Then
                If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
            End If
            
            If Asc(strChar) < 0 Then
                intCharNumber = 0
                Y1 = Y1 + ROWHEIGHT * 5
            Else
                Y1 = Y1 + ROWHEIGHT * 2.5
                intCharNumber = intCharNumber + 1
            End If
        Next
        'If Y1 <= dblHeight + 100 Then Exit For
        Exit For            '���ܳ���������,Ϊ�˵õ��߶�,���Ըú�����������뻹�Ǳ�������
    Next
    
    objDraw.FontSize = sinFontSize_Bak
    GetFontSize = sinFontSize
End Function

Private Sub OutputNote(ByVal objDraw As Object, ByVal dblHeight As Double, ByRef rsNote As ADODB.Recordset, ByVal lngTmpX As Long, ByVal lngTmpY As Long)
    '���������Ϣ:��Ժ,���,ת��,��Ժ,��������,δ��˵��,�ϱ�˵��������
    'δ��˵�����ϱ�˵��,��û�����ת�������估��������Ϣʱ,��ӡ��42-40֮��;�����40��ʼ���´�ӡ
    '��δ��˵�����ϱ�˵����,���ת����Ϣ��һ���̶ȷ������ʱ,����д������̶���,�������̶�Ҳ����Ϣ,˳��
    Dim intCol As Integer                   '��¼��ǰ�к�
    Dim intMax As Integer                   '������
    Dim intCur As Integer                   '��ǰ��¼��λ��
    Dim bln�ϱ� As Boolean
    Dim sinX1 As Single, sinY1 As Single, sinHeight As Single, H_9pt As Single, sinMaxY1 As Single
    Dim rsTarget As New ADODB.Recordset

    '����ַ���ر�������
    Dim sinFontSize As Single
    Dim sinFontSize_Bak As Single
    Dim intCharNumber As Integer
    Dim intCount As Integer
    Dim strChar As String

    intMax = 41
    H_9pt = ROWHEIGHT * 10 / 3
    sinFontSize_Bak = objDraw.FontSize
    Set rsTarget = rsNote.Clone
    With rsNote
        If .RecordCount = 0 Then Exit Sub
        .Sort = "�к�,ʱ��"
        intCol = !�к�

        '�������ת��������ѭ��
        Do While Not .EOF
            If Trim(NVL(!���)) <> "" Then
                If !���� = 1 Then   '���ת��������
                    '������ӡ���Ƿ��Ѵ������,���������У������
                    If intCol > intMax Then intCol = intMax

                    '����õ����ʵ������С���߶�
                    !�����С = GetFontSize(objDraw, dblHeight, NVL(!���), sinY1)
                    !�߶� = sinY1
                    !��ӡ�� = IIf(intCol < !�к�, !�к�, intCol)
                    .Update
                    If intCol <= !�к� Then intCol = !�к�
                    intCol = intCol + 1
                Else        '�ϱ�˵��,δ��˵��
                    Call GetFontSize(objDraw, dblHeight, NVL(!���), sinY1)
                    !�߶� = sinY1
                    .Update
                End If
            End If

            .MoveNext
        Loop
        .MoveFirst

        '�������ת�ȵ�������(ֻ�����һ�вŴ���һ���������)
        sinY1 = (lngTmpY + 3 * H_9pt / 2)       '42��
        .Filter = "��ӡ��='" & intMax & "'"
        .Sort = "�к�,ʱ��"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            'ֻ�����ת�����Ÿ����˴�ӡ��
            !���� = Split(!����, ";")(0) & ";" & sinY1
            .Update
            sinY1 = sinY1 + !�߶� + 100

            .MoveNext
        Loop
        .Filter = 0
        .MoveFirst

        '����У��δ��˵���Լ��ϱ�˵���ĸ߶�(δ��˵�����ϱ�˵��,��û�����ת�������估��������Ϣʱ,��ӡ��42-40֮��;�����40��ʼ���´�ӡ)
        Set rsTarget = .Clone
        intCol = 0
        Do While Not .EOF
            If !���� = 2 Then       '�ϱ�˵��,δ��˵��
                Set rsTarget = .Clone
                rsTarget.Filter = "��ӡ��='" & !�к� & "'"
                If rsTarget.RecordCount <> 0 Then
                    '�Ѵ��ڴ�ӡ���ݵĲ�У��������
                    sinMaxY1 = Split(rsTarget!����, ";")(1)
                    Do While Not rsTarget.EOF
                        If bln�ϱ� = False Then
                            bln�ϱ� = (rsTarget!���� = 2)
                        End If
                        sinMaxY1 = sinMaxY1 + rsTarget!�߶� + 100
                        rsTarget.MoveNext
                    Loop
                    sinY1 = (lngTmpY + 3 * H_9pt / 2) + 10 * 3 * H_9pt / 2 - H_9pt / 2      '40�ȵ�����
                    If sinY1 < sinMaxY1 Or bln�ϱ� Then sinY1 = sinMaxY1
                    sinHeight = !�߶�
                    intCol = !�к�
                Else
                    sinY1 = (lngTmpY + 3 * H_9pt / 2)       '42��
                    intCol = !�к�
                    sinHeight = !�߶�
                End If
                rsTarget.Filter = 0

                !���� = Split(!����, ";")(0) & ";" & sinY1
                !��ӡ�� = !�к�                                 '��ʱ���´�ӡ��,�Ա������ѭ������
                .Update
            End If
            .MoveNext
        Loop

        '��ʼ�������������
        .MoveFirst
        Do While Not .EOF
            If Trim(NVL(!���)) <> "" Then
                'If (!���� = 2) Then Stop
                sinX1 = lngTmpX + (IIf(!��ӡ�� = "", Val(!�к�), Val(!��ӡ��))) * HOUR_STEP_Twips + HOUR_STEP_Twips / 2
                sinY1 = Split(!����, ";")(1)
                intCharNumber = 0
                objDraw.FontSize = IIf(!�����С = "", 9, !�����С)

                For intCount = 1 To Len(!���)
                    strChar = Mid(!���, intCount, 1)

                    If Asc(strChar) < 0 Then
                        If intCharNumber Mod 2 = 1 Then sinY1 = sinY1 + ROWHEIGHT * 2.5
                    End If
                    Call DrawRotateText(objDraw, sinX1 - objDraw.TextWidth(strChar) / 2, sinY1 + 15, strChar, vbRed)
                    If Asc(strChar) < 0 Then
                        intCharNumber = 0
                        sinY1 = sinY1 + ROWHEIGHT * 5
                    Else
                        sinY1 = sinY1 + ROWHEIGHT * 2.5
                        intCharNumber = intCharNumber + 1
                    End If
                Next
            End If

            .MoveNext
        Loop
    End With
    objDraw.FontSize = sinFontSize_Bak
End Sub

Private Function SimplifyString(ByVal strText As String) As String
    Dim arrData
    Dim strData As String
    Dim intLen As Integer, intActLen As Integer
    Dim intCol As Integer, intCount As Integer
    '�����ַ���,ȥ���ظ�������
    
    arrData = Split(strText, " ")
    intCount = UBound(arrData)
    strData = ""
    For intCol = 0 To intCount
        If InStr(1, " " & strData & " ", " " & arrData(intCol) & " ") = 0 Then
            strData = strData & " " & arrData(intCol)
        End If
    Next
    SimplifyString = Mid(strData, 2)
End Function

Public Sub GetValue(strValue As String)
    Dim str���� As String
    Dim str���� As String
    
    str���� = strValue & "��"
    str���� = CStr(Val(strValue) * 9 / 5 + 32) & "��"
    strValue = str���� & String(10 - Len(str���� & str����), " ") & str����
End Sub
