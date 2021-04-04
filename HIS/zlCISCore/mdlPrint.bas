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

Private Const HOUR_STEP_Twips = 300 '������������Сʱ֮��Ŀ�� �������±�
Private Const INTSTEPTwip = 90  '��������5�������Ŀ�� ��������
Private Const STRING_WAY As String = "��"

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

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
Public Const PHYSICALWIDTH = 110   'Physical Width in device units
Public Const PHYSICALHEIGHT = 111  'Physical Height in device units
Public Const PHYSICALOFFSETX = 112 'Physical Printable Area x margin
Public Const PHYSICALOFFSETY = 113 'Physical Printable Area y margin
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
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
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
    CX As Long
    CY As Long
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
    Flags As Long
    pName As Long   ' String
    Size As SIZEL
    ImageableArea As RECTL
End Type

Public Type sFORM_INFO_1
    Flags As Long
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

Private Type CellInfoType
    FontSize As Long            '��ӡ�����С
    FontName As String          '��ӡ����
    FontBold As Boolean         '�Ӵ�
    FontItalic As Boolean       'б��
    FontColor As OLE_COLOR      '�ı���ɫ
    FontBackColor As OLE_COLOR  '��Ԫ�񱳾�
    LineColor As OLE_COLOR      '�߿���ɫ
    Text As String              '��ӡ�ı�
    Merge As String             '�ϲ���Ԫ����Ϣ
    Height As Long              '��
    Width As Long               '��
    HAlign As Byte
    VAlign As Byte
End Type

Private mCellArr() As CellInfoType

Dim mPageHeadDep As String                  '����
Dim mPageHeadName As String                 '��������
Dim mPageHeadNo As String                   'סԺ�������
Dim mNewPageTop As Long                     '��ҳ�ĳ�ʹ���߶�
Dim mNewPageInit As Long                    '��ҳ��ʹ���߶�
Dim mPageBedNumber As String                '����
Dim mPrintBegingPage As Long                '��ӡ��ʼҳ
Dim mPrintEndPage As Long                   '��ӡ����ҳ
Dim mPageNumber As Long                     '��ӡҳ��


'===================================================================================

Private Function GetCellWH(arrTmp() As CellInfoType, ByVal Row As Long, ByVal Col As Long, ByVal Row1 As Long, ByVal Col1 As Long, Optional blnRC As Boolean) As Long
    '�õ���ĳһ��Ԫ��ĳһ��Ԫ����ο����
    Dim i As Long
    Dim lngWH As Long  '��ʱ��¼����
    
    lngWH = 0
    If Row > Row1 Then Exit Function
    If Col > Col1 Then Exit Function
    If blnRC Then    '���и�
        For i = LBound(arrTmp, 1) To UBound(arrTmp, 1)
            If i >= Row And i <= Row1 Then
                lngWH = lngWH + arrTmp(i, 1).Height
            End If
        Next
        GetCellWH = lngWH
    Else
        '���п�
        For i = LBound(arrTmp, 2) To UBound(arrTmp, 2)
            If i >= Col And i <= Col1 Then
                lngWH = lngWH + arrTmp(1, i).Width
            End If
        Next
        GetCellWH = lngWH
    End If
End Function

Private Sub GridDraw(objOut As Object, objDraw As Object, ByVal lng���� As Long, y As Long, ByVal blnPrintNO As Boolean, lngEndPage As Long, Optional ByVal bytGridAlign As Byte = 0)
    '����:���ݴ���Ĳ���ID�������ж������ݲ����뵽������,�ٴ�����ӡ
    '�ֿ���Ŀ����Ϊ�˽������Թ���ʹ��DrawGridArr����
    Dim arrTmp() As CellInfoType
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHandle
    
    strSQL = "SELECT * From ���˲��������� " & vbCrLf & _
            "Where ����id =  " & lng���� & vbCrLf & _
            "ORDER BY �ؼ���,-��,-��"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������ӡ")
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        '��һ�бض��Ǳ�������˵��
        ReDim arrTmp(1 To rsTmp!��, 1 To rsTmp!��) As CellInfoType
        '����ı������
        rsTmp.MoveNext
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount - 1
                arrTmp(-rsTmp!��, -rsTmp!��).FontBackColor = -1
                arrTmp(-rsTmp!��, -rsTmp!��).FontBold = objDraw.Font.Bold
                arrTmp(-rsTmp!��, -rsTmp!��).FontColor = objDraw.ForeColor
                arrTmp(-rsTmp!��, -rsTmp!��).FontItalic = objDraw.Font.Italic
                arrTmp(-rsTmp!��, -rsTmp!��).FontName = objDraw.Font.Name
                arrTmp(-rsTmp!��, -rsTmp!��).FontSize = objDraw.Font.Size
                strSQL = Format(CStr(zlCommFun.Nvl(rsTmp!�ϲ���, "")), "0000000000000000")
                arrTmp(-rsTmp!��, -rsTmp!��).Merge = IIf(strSQL = "0" Or strSQL = "0000000000000000", "", strSQL)
                arrTmp(-rsTmp!��, -rsTmp!��).Text = zlCommFun.Nvl(rsTmp!��������)
                arrTmp(-rsTmp!��, -rsTmp!��).Width = zlCommFun.Nvl(rsTmp!��, 0)
                arrTmp(-rsTmp!��, -rsTmp!��).Height = zlCommFun.Nvl(rsTmp!��, 0)
                '���ڶ��뷽ʽ��ͬ������Ҫת��
                Select Case zlCommFun.Nvl(rsTmp!����, 1)
                    Case 2: arrTmp(-rsTmp!��, -rsTmp!��).HAlign = 0
                    Case 3: arrTmp(-rsTmp!��, -rsTmp!��).HAlign = 1
                    Case Else: arrTmp(-rsTmp!��, -rsTmp!��).HAlign = 2
                End Select
                arrTmp(-rsTmp!��, -rsTmp!��).VAlign = 1
                rsTmp.MoveNext
            Next
            Call DrawGridArr(objOut, objDraw, arrTmp, y, blnPrintNO, lngEndPage, bytGridAlign)
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DrawGridArr(objOut As Object, objDraw As Object, arrTmp() As CellInfoType, y As Long, ByVal blnPrintNO As Boolean, lngEndPage As Long, Optional ByVal bytGridAlign As Byte = 0)
    '����:���ݴ����������д�ӡ���
    '����:lngEndPage    ��ʼ��ҳ����
    '     ObjOut    ��������Ԥ�����������ӡ������
    '     ObjDraw   ���������ӡ�������
    '     arrTmp()  �����Ѿ�������ֵ�ĵ�Ԫ����Ϣ
    '     Y         ���ÿ�ʼ��Y���꿪ʼ��ͼ
    '     bytGridAlign  ���ñ��������뷽ʽ    (��0 ��1 ��2 )
    Dim i As Long
    Dim j As Long
    Dim m As Long   '�ϲ���ʼ��Ԫ����
    Dim n As Long   '�ϲ���ʼ��Ԫ����
    Dim m1 As Long  '�ϲ���ֹ��Ԫ����
    Dim n1 As Long  '�ϲ���ֹ��Ԫ����
    Dim x As Long
    Dim lngLeft As Long, lngRight As Long, lngTop As Long, lngBottom As Long, lngWidth As Long, lngHeight As Long   'ֽ�Ŵ�С�߽�
    Dim lngGridW As Long, lngGridH As Long  '�����
    Dim lngTmpPageNo As Long    '��ʱ����ҳ��
    Dim TmpFont As New StdFont
    Dim strMerge As String
    On Error GoTo ErrHandle
    
    '�õ�ֽ�ŵı߽�����
    lngLeft = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��߾�", OFFSET_LEFT) * 56.7 + Screen.TwipsPerPixelX * 2
    lngRight = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ұ߾�", OFFSET_RIGHT) * 56.7 - Screen.TwipsPerPixelX * 2
    lngTop = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ϱ߾�", OFFSET_TOP) * 56.7 + Screen.TwipsPerPixelY * 2
    lngBottom = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�±߾�", OFFSET_BOTTOM) * 56.7 - Screen.TwipsPerPixelY * 2
    lngWidth = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "���", Printer.Width)
    lngHeight = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�߶�", Printer.Height)
    
    On Error Resume Next
    Err.Clear
    i = LBound(arrTmp, 1)
    i = LBound(arrTmp, 2)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    '�õ����Ŀ����߶�
    lngGridH = GetCellWH(arrTmp, 1, 1, UBound(arrTmp, 1), UBound(arrTmp, 2), True)
    lngGridW = GetCellWH(arrTmp, 1, 1, UBound(arrTmp, 1), UBound(arrTmp, 2), False)
    y = y + lngGridH
    '�ж��Ƿ���ҳ
    Set objDraw = Nothing
    If blnPrintNO Then
        Set objDraw = CheckNewPage(objOut, lngEndPage, y, lngGridH)
    Else
        lngTmpPageNo = 0
        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, lngGridH)
    End If
    If objDraw Is Nothing Then Exit Sub

    '���ݶ��뷽ʽ����X����
    Select Case bytGridAlign
        Case 1      '�ж���
            lngLeft = lngLeft + (lngWidth - (lngLeft + lngRight) - lngGridW) / 2
        Case 2      '�Ҷ���
            lngLeft = lngWidth - (lngRight + lngGridW)
        Case Else   '�����
            lngLeft = lngLeft
    End Select
    '��ʼ��
    For i = LBound(arrTmp, 1) To UBound(arrTmp, 1)
        For j = LBound(arrTmp, 2) To UBound(arrTmp, 2)
            '��λX
            If j = LBound(arrTmp, 2) Then
                x = lngLeft
            Else
                x = x + arrTmp(1, j - 1).Width
            End If
            '��������
            TmpFont.Bold = arrTmp(i, j).FontBold
            TmpFont.Italic = arrTmp(i, j).FontItalic
            TmpFont.Size = arrTmp(i, j).FontSize
            TmpFont.Name = arrTmp(i, j).FontName
            '�����ϲ���Ϣ���н���
            strMerge = arrTmp(i, j).Merge
            If Len(strMerge) = 16 And IsNumeric(strMerge) And strMerge Like "0###0###0###0###" Then
                '�����ʽ:0006000100060008
                m = CLng(Mid(strMerge, 1, 4))
                n = CLng(Mid(strMerge, 5, 4))
                m1 = CLng(Mid(strMerge, 9, 4))
                n1 = CLng(Mid(strMerge, 13, 4))
                If i = m And j = n Then
                    Call DrawCell(objDraw, arrTmp(m, n).Text, x, y, GetCellWH(arrTmp, m, n, m1, n1, False), GetCellWH(arrTmp, m, n, m1, n1, True), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , arrTmp(i, j).LineColor, arrTmp(i, j).FontColor, arrTmp(i, j).FontBackColor, TmpFont, "1111", arrTmp(i, j).HAlign, arrTmp(i, j).VAlign)
                End If
                arrTmp(i, j).Text = arrTmp(m, n).Text
            Else
                Call DrawCell(objDraw, arrTmp(i, j).Text, x, y, arrTmp(1, j).Width, arrTmp(i, 1).Height, IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , arrTmp(i, j).LineColor, arrTmp(i, j).FontColor, arrTmp(i, j).FontBackColor, TmpFont, "1111", arrTmp(i, j).HAlign, arrTmp(i, j).VAlign)
            End If
        Next
        '��λY
        y = y + arrTmp(i, 1).Height
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

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

Public Function DrawCell(Dev As Object, ByVal Data As Variant, ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long, lngNowPage As Long, _
    Optional ByVal TW As Long, Optional ByVal TH As Long, Optional BorderColor As Long, _
    Optional ForeColor As Long, Optional BackColor As Long = &HFFFFFF, Optional ByVal Font As StdFont, _
        Optional Border As String = "1111", Optional HAlign As Byte, Optional VAlign As Byte = 1, Optional Warp As Boolean, _
        Optional Ratio As Single = 1) As Boolean
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
    
    On Error GoTo errH
    
    DrawCell = True
    
    '��Χ�޶�
    If TW > 0 Then
        If x > TW Then Exit Function
        If x + W > TW Then W = TW - x
    End If
    If TH > 0 Then
        If y > TH Then Exit Function
        If y + H > TH Then H = TH - y
    End If
    
    If TypeName(Data) = "Integer" Then
        x = x * Ratio: y = y * Ratio: W = W * Ratio: H = H * Ratio
        If Val(Data) < 0 Then
            Dev.Line (x, y)-(x + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, B '����
        Else
            Dev.Line (x, y)-(x + W - IIf(W > 0, Screen.TwipsPerPixelX * Ratio, 0), y + H - IIf(H > 0, Screen.TwipsPerPixelY * Ratio, 0)), ForeColor, BF 'ʵ�ľ���(����)
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
        Dev.Font.Size = Font.Size
        Dev.Font.Bold = Font.Bold
        Dev.Font.Underline = Font.Underline
        Dev.Font.Italic = Font.Italic
        SetPrinterFont Dev.Font, Font.Size
        
        '�����ź���������������,�ж�ʱ��ԭʼ��СΪ׼
        If H >= Printer.TextHeight(Replace(Data, vbCrLf, "")) Then blnH = True '�߶��Ƿ���(�ӻس�����һ�и߶�)
        If W >= Printer.TextWidth(Data) Then blnW = True And InStr(Data, vbCrLf) = 0   '����Ƿ���(�ӻس���Ϊ������,�Ա����)
        
        '����
        LINE_W = 30 * Ratio '���߼�����(���ʱ��,�ж�ʱ����)
        x = -Int(-x * Ratio): y = -Int(-y * Ratio)
        W = -Int(-W * Ratio): H = -Int(-H * Ratio)
        Dev.Font.Size = Font.Size * Ratio
        SetPrinterFont Dev.Font, Font.Size
        
        '�������
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
        (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
            Dev.Line (x, y)-(x + W, y + H), BackColor, BF
        End If
        
        Dev.ForeColor = ForeColor
        '�������(�߿�֮���ٸ�һ��)
        '�����߶ȷ�Χ�����
        If blnH Then
            If blnW Then
                Select Case HAlign
                Case 0
                    Dev.CurrentX = x + LINE_W
                Case 1
                    Dev.CurrentX = x + (W - Printer.TextWidth(Data)) / 2
                Case 2
                    Dev.CurrentX = x + W - LINE_W - Printer.TextWidth(Data)
                End Select
                Select Case VAlign
                Case 0
                    Dev.CurrentY = y + LINE_W
                Case 1
                    Dev.CurrentY = y + (H - Printer.TextHeight(Data)) / 2 + LINE_W / 2
                Case 2
                    Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Data)
                End Select
                Dev.FontTransparent = True
                If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
                    Dev.Print Data
                End If
                Select Case VAlign
                Case 0
                    Dev.CurrentY = y + LINE_W + Printer.TextHeight(Date)
                Case 1
                    Dev.CurrentY = y + (H - Printer.TextHeight(Data)) / 2 + LINE_W / 2 + Printer.TextHeight(Date)
                Case 2
                    Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Data) + Printer.TextHeight(Date)
                End Select
            Else
                If Not Warp Then
                    '���Զ�����ʱ�����ֲ����
                    For i = 1 To Len(Data)
                        If Printer.TextWidth(Text & Mid(Data, i, 1)) > W Then Exit For
                        Text = Text & Mid(Data, i, 1)
                    Next
                    Select Case HAlign
                    Case 0
                        Dev.CurrentX = x + LINE_W
                    Case 1
                        Dev.CurrentX = x + (W - Printer.TextWidth(Text)) / 2
                    Case 2
                        Dev.CurrentX = x + W - LINE_W - Printer.TextWidth(Text)
                    End Select
                    Select Case VAlign
                    Case 0
                        Dev.CurrentY = y + LINE_W
                    Case 1
                        Dev.CurrentY = y + (H - Printer.TextHeight(Text)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Text)
                    End Select
                    Dev.FontTransparent = True
                    '�����ȡ����
                    If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                    (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
                        Dev.Print Text
                    End If
                    Select Case VAlign
                    Case 0
                        Dev.CurrentY = y + LINE_W + Printer.TextHeight(Text)
                    Case 1
                        Dev.CurrentY = y + (H - Printer.TextHeight(Text)) / 2 + LINE_W / 2 + Printer.TextHeight(Text)
                    Case 2
                        Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Text) + Printer.TextHeight(Text)
                    End Select
                Else
                    '������ֳɶ���(�ڿ�߷�Χ��)
                    ReDim arrText(0) '�ڴ�,��һ�в����ܳ���
                    Data = Replace(Data, vbCrLf, vbCr)
                    Data = Replace(Data, vbLf, vbCr)
                    For i = 1 To Len(Data)
                        If Mid(Data, i, 1) = vbCr Then
                            '���г������˳�,���߲��ݲ����
                            If Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        ElseIf Printer.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) > W Then
                            '���г������˳�,���߲��ݲ����
                            If Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 2) > H Then Exit For
                            ReDim Preserve arrText(UBound(arrText) + 1)
                        End If
                        '�п���һ��һ���ַ���ȶ�����
                        If Printer.TextWidth(arrText(UBound(arrText)) & Mid(Data, i, 1)) <= W And Mid(Data, i, 1) <> vbCr Then
                            arrText(UBound(arrText)) = arrText(UBound(arrText)) & Mid(Data, i, 1)
                        End If
                    Next
                    
                    '�����ʼ����
                    Select Case VAlign
                    Case 0
                        Dev.CurrentY = y + LINE_W
                    Case 1
                        Dev.CurrentY = y + (H - Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)) / 2 + LINE_W / 2
                    Case 2
                        Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)
                    End Select
                    
                    '�������
                    For i = 0 To UBound(arrText)
                        Select Case HAlign
                        Case 0
                            Dev.CurrentX = x + LINE_W
                        Case 1
                            Dev.CurrentX = x + (W - Printer.TextWidth(arrText(i))) / 2
                        Case 2
                            Dev.CurrentX = x + W - LINE_W - Printer.TextWidth(arrText(i))
                        End Select
                        Dev.FontTransparent = True
                        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                        (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
                            Dev.Print arrText(i)
                        End If
                    Next
                    If UBound(arrText) > 0 Then
                        Select Case VAlign
                        Case 0
                            Dev.CurrentY = y + LINE_W + Printer.TextHeight(arrText(0))
                        Case 1
                            Dev.CurrentY = y + (H - Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1)) / 2 + LINE_W / 2 + Printer.TextHeight(arrText(0))
                        Case 2
                            Dev.CurrentY = y + H - LINE_W - Printer.TextHeight(Replace(Data, vbCr, "")) * (UBound(arrText) + 1) + Printer.TextHeight(arrText(0))
                        End Select
                    End If
                End If
            End If
        End If
    ElseIf Not Data Is Nothing Then
        LINE_W = 30 * Ratio '���߼�����(���ʱ��,�ж�ʱ����)
        x = x * Ratio: y = y * Ratio: W = W * Ratio: H = H * Ratio
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
        (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
            'ͼ��(�߿�֮��)
            Dev.PaintPicture Data, x + 15, y + 15, W - LINE_W, H - LINE_W
        End If
    End If
    
    If TypeName(Data) <> "Integer" Then
        '�����߿�
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
        (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
            If Mid(Border, 1, 1) Then Dev.Line (x, y)-(x + W, y), BorderColor
            If Mid(Border, 2, 1) Then Dev.Line (x, y + H)-(x + W, y + H), BorderColor
            If Mid(Border, 3, 1) Then Dev.Line (x, y)-(x, y + H), BorderColor
            If Mid(Border, 4, 1) Then Dev.Line (x + W, y)-(x + W, y + H), BorderColor
        End If
    End If
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
    
    If Printers.Count = 0 Then Exit Function
    
    '��ʼ����ӡ����
    strPrinter = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��ӡ��", Printer.DeviceName)
    intPage = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "ֽ��", Printer.PaperSize)
    lngWidth = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "���", Printer.Width)
    lngHeight = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�߶�", Printer.Height)
    intOrient = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "ֽ��", Printer.Orientation)
    intBin = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��ֽ", Printer.PaperBin)
    
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
        If AddCustomPaper(objParent.hwnd, lngWidth / 56.7, lngHeight / 56.7) = FORM_NOT_SELECTED Then Exit Function
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
            If .Size.CX = FormSize.CX And .Size.CY = FormSize.CY Then
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
        
        ' Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        ' Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        ' If FormName is "zlBillPaper", we must make sure it exists
        ' before using it. Otherwise, it came from our EnumForms list,
        ' and we do not need to check first. Note that we could have
        ' passed in a Flag instead of checking for a literal name.
        
        ' Use form "zlBillPaper", adding it if necessary.
        ' Set the desired size of the form needed.
        ' Given in thousandths of millimeters
        vFormSize.CX = lngWidth * 1000 ' width
        vFormSize.CY = lngHeight * 1000 ' height
        
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
    Dim sTemp As String * 512, x As Long
    
    x = lstrcpy(sTemp, ByVal Add)
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
        .Flags = 0
        .pName = strFormName
        With .Size
            .CX = vFormSize.CX
            .CY = vFormSize.CY
        End With
        With .ImageableArea
            .Left = 0
            .Top = 0
            .Right = FI1.Size.CX
            .Bottom = FI1.Size.CY
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

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
    '���ܣ�������ʽ���߰�ť�е���һ���˵�
    Dim vRect As RECT, vDot1 As POINTAPI, vDot2 As POINTAPI
    
    Call GetWindowRect(ToolBar.hwnd, vRect)
    vDot1.x = vRect.Left: vDot1.y = vRect.Top
    vDot2.x = vRect.Right: vDot2.y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot2)
    
    vDot1.x = vDot1.x * 15: vDot1.y = vDot1.y * 15
    vDot2.x = vDot2.x * 15: vDot2.y = vDot2.y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.x + Button.Left, vDot2.y
End Sub

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
                SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                ShowWindow frmFlash.hwnd, 5
            Else
                Err.Clear
                frmFlash.Show , frmParent
                If Err.Number <> 0 Then
                    Err.Clear
                    SetWindowPos frmFlash.hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                    ShowWindow frmFlash.hwnd, 5
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

Public Function GetDeptName(lngID As Long) As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select * From ���ű� Where ID=" & lngID
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlPrint")
    If rsTmp.RecordCount > 0 Then GetDeptName = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetLastPrint(ByVal lng������¼ID As Long, ByRef lngEndY As Long, ByRef lngEndPage As Long) As Boolean
    '���ܣ���ȡ�����ϴβ�����ӡ����λ����Ϣ
    '���أ�lngEndY=�ϴδ�ӡ�Ľ���λ��(mm)
    '      intEndPage=�ϴδ�ӡ�Ľ���ҳ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select * From ������ӡ��¼ Where ������¼ID=" & lng������¼ID & " Order By ��ӡʱ�� Desc"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������ӡ")
    
    If rsTmp.RecordCount > 0 Then
        lngEndY = rsTmp!����λ��
        lngEndPage = rsTmp!����ҳ��
    Else
        lngEndY = 0
        lngEndPage = 1
    End If
    GetLastPrint = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function NewPrintPage(objOut As Object, intPage As Long, Optional blnNewPage As Boolean = True) As Object
    '���ܣ���ӡ��Ԥ��һҳ����ʱ�Ե�ǰҳ����������,��������ҳ
    '������blnNewPage=ΪFalseʱ����ӡҳ�ŵ�,һ���ӡ��������������,��˲����������
    '���أ���ҳ����,����Ϊ��ӡ����PictureBox
    Dim objDraw As Object, blnPrint As Boolean
    Dim lngWidth As Long, lngHeight As Long, lngOldY As Long
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim strFontName As String, lngFontSize As Long, blnFontBold As Boolean
    Dim blnFontItalic As Boolean, lngFontColor As Long
    Dim x As Long, y As Long, H_9pt As Long, W_9pt As Long
    Dim strText As String
    Dim objPrinter As Object
    On Error GoTo errH
    
    blnPrint = TypeName(objOut) = "Printer"
    
    '�߽���Ϣ(Twip)
    lngLeft = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��߾�", OFFSET_LEFT) * 56.7
    lngRight = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ұ߾�", OFFSET_RIGHT) * 56.7
    lngTop = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ϱ߾�", OFFSET_TOP) * 56.7 + Screen.TwipsPerPixelY * 2
    lngBottom = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�±߾�", OFFSET_BOTTOM) * 56.7
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
    
    '������ҳ
    If blnNewPage Then
        '��ӡҳ��(0Ϊ����ӡ)
        If intPage <> 0 Then
            objDraw.ForeColor = 0
            objDraw.Font.Name = "����"
            objDraw.Font.Size = 9
            objDraw.Font.Bold = False
            SetPrinterFont objDraw.Font, 9
            objDraw.CurrentY = lngHeight - IIf(lngBottom < 1134, 1134, lngBottom) - (Printer.TextHeight("��") * 2)
            objDraw.CurrentX = lngLeft + (lngWidth - lngLeft - lngRight) * (3 / 4)
            objDraw.FontTransparent = True
            If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
                objDraw.Print "���� " & intPage & " ҳ��"
            End If
        End If
        intPage = intPage + 1
        If intPage - mPageNumber + 1 > mPrintEndPage And mPrintEndPage > 0 Then Set objDraw = Nothing: Exit Function
        If blnPrint Then
            '����ָ��ҳ
            If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
                If (intPage - mPageNumber) >= mPrintBegingPage Then
                    Printer.NewPage
                End If
                Set objDraw = Printer
            Else
'                Printer.KillDoc
'                InitPrint Printer
'                Printer.EndDoc
            End If
        Else
'            If intPage - mPageNumber >= mPrintBegingPage And intPage - mPageNumber <= mPrintEndPage Then
'                Load objOut.picPage(objOut.picPage.UBound + 1)
'            End If
            'Ԥ����ӡ����
            objDraw.DrawStyle = 2
            objDraw.Line (0, lngTop)-(lngWidth, lngTop), &H808080
            objDraw.Line (0, lngHeight - lngBottom)-(lngWidth, lngHeight - lngBottom), &H808080
            objDraw.Line (lngLeft, 0)-(lngLeft, lngHeight), &H808080
            objDraw.Line (lngWidth - lngRight, 0)-(lngWidth - lngRight, lngHeight), &H808080
            objDraw.DrawStyle = 0
    
            If intPage - mPageNumber >= mPrintBegingPage Then
                Load objOut.picPage(objOut.picPage.UBound + 1)
            End If
            Set objDraw = objOut.picPage(objOut.picPage.UBound)
            objDraw.Width = Printer.Width
            objDraw.Height = Printer.Height
            objDraw.ZOrder
            objDraw.Cls
            objDraw.AutoRedraw = True
        End If
        '��ҳ���
        objDraw.CurrentX = lngLeft: objDraw.CurrentY = lngTop
        '--�����޸�
'        'Ԥ����ӡ����
'        objDraw.DrawStyle = 2
'        objDraw.Line (0, lngTop)-(lngWidth, lngTop), &H808080
'        objDraw.Line (0, lngHeight - lngBottom)-(lngWidth, lngHeight - lngBottom), &H808080
'        objDraw.Line (lngLeft, 0)-(lngLeft, lngHeight), &H808080
'        objDraw.Line (lngWidth - lngRight, 0)-(lngWidth - lngRight, lngHeight), &H808080
'        objDraw.DrawStyle = 0
'        objDraw.CurrentY = IIf(mNewPageInit >= lngTop, mNewPageInit, lngTop)
        objDraw.CurrentY = lngTop
        objDraw.Font.Name = "����"
        objDraw.Font.Size = 9
        objDraw.Font.Bold = False
        SetPrinterFont objDraw.Font, 9
        H_9pt = Printer.TextHeight("��")
        W_9pt = Printer.TextWidth("��")
        '��ӡ��ǰ����������Ϣ
        objDraw.ForeColor = 0
        objDraw.Font.Name = "����"
        objDraw.Font.Size = 18
        objDraw.Font.Bold = True
        SetPrinterFont objDraw.Font, 18
        '��������
        strText = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "��λ����")
        '�ж��Ƿ���ҳ
        y = objDraw.CurrentY + H_9pt + Printer.TextHeight("��")  '��ʼ����2���ָ�
        y = y - Printer.TextHeight("��")
        '�õ�����XY����
        x = lngLeft + (lngWidth - (lngLeft + lngRight) - Printer.TextWidth(strText)) / 2
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        objDraw.CurrentY = y + Printer.TextHeight(strText)
        
        '��ӡ��������,ǩ��,����
        objDraw.ForeColor = 0
        objDraw.Font.Name = "����"
        objDraw.Font.Size = 10.5
        objDraw.Font.Bold = False
        SetPrinterFont objDraw.Font, 10.5
        '�ж��Ƿ���ҳ
        y = objDraw.CurrentY + H_9pt * 2 + Printer.TextHeight("��") '��ʼ����2���ָ�
        y = y - Printer.TextHeight("��")
        strText = mPageHeadDep
        
        x = lngLeft
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        
        strText = mPageHeadName
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        
        strText = mPageHeadNo
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (4 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        
        strText = mPageBedNumber
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (7 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        objDraw.CurrentY = y + Printer.TextHeight(strText)
        y = objDraw.CurrentY + H_9pt / 5: x = lngLeft
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Line (lngLeft, y)-(lngWidth - lngRight, y), 0
        End If
        '--
        objDraw.Font.Name = strFontName
        objDraw.Font.Size = lngFontSize
        objDraw.Font.Bold = blnFontBold
        objDraw.Font.Italic = blnFontItalic
        objDraw.ForeColor = lngFontColor
        SetPrinterFont objDraw.Font, Int(Nvl(lngFontSize, 0))
        
        mNewPageTop = y + 100
        
    Else
        objDraw.CurrentY = lngOldY
        If intPage <> 0 Then
            objDraw.ForeColor = 0
            objDraw.Font.Name = "����"
            objDraw.Font.Size = 9
            objDraw.Font.Bold = False
            SetPrinterFont objDraw.Font, 9
            objDraw.CurrentY = lngHeight - IIf(lngBottom < 1134, 1134, lngBottom) - (Printer.TextHeight("��") * 2)
            objDraw.CurrentX = lngLeft + (lngWidth - lngLeft - lngRight) * (3 / 4)
            objDraw.FontTransparent = True
            If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (intPage - mPageNumber >= mPrintBegingPage - 1 And intPage - mPageNumber <= mPrintEndPage) Then
                objDraw.Print "���� " & intPage & " ҳ��"
            End If
        End If
    End If
    Set NewPrintPage = objDraw
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadPatiInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
    '�ڲ�����ӡʱ�õ����˵���Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng��ҳID = 0 Then
        strSQL = "Select A.����ID,A.�����,A.����,A.�Ա�,A.����,A.����,A.����," & _
        " A.����״��,A.ְҵ,A.���֤��,A.������λ,A.��ͥ��ַ" & _
            " From ������Ϣ A Where ����ID=" & lng����ID
    Else
        strSQL = "Select A.����ID,A.סԺ��,A.����,A.�Ա�,A.����,A.����,A.����," & _
        " A.����״��,A.ְҵ,A.���֤��,A.������λ,A.��ͥ��ַ," & _
            " B.��Ժ����,C.���� as ��Ժ����,B.��Ժ����,D.���� as ��Ժ����,B.��Ժ����" & _
            " From ������Ϣ A,������ҳ B,���ű� C,���ű� D" & _
            " Where B.��Ժ����ID=C.ID And B.��Ժ����ID=D.ID And A.����ID=B.����ID" & _
            " And A.����ID=" & lng����ID & " And B.��ҳID=" & lng��ҳID
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������ӡ")
    If rsTmp.RecordCount > 0 Then
        Set ReadPatiInfo = rsTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckNewPage(objOut As Object, lngPage As Long, lngY As Long, Optional lngDefHeight As Long = -1) As Object
    '���ܣ�����ǲ��ǳ����߽磬�����¿�ʼһҳ���󷵻�
    '������ObjOut       �������
    '       lngPage     ҳ��Ϊ0ʱ��ʾ����ӡ
    '       lngDefY     �ɴ�ӡ��������
    '       lngY        ��ǰ��ӡλ��
    Dim lngBottom As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngTop As Long
    Dim lngFontHeigh As Long
    Dim lngTmp As Long
    
    
    '�߽���Ϣ(Twip)
    lngBottom = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�±߾�", OFFSET_BOTTOM) * 56.7
    lngTop = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ϱ߾�", OFFSET_TOP) * 56.7
    
    lngWidth = Printer.Width
    lngHeight = Printer.Height
    If lngPage > 0 Then
        lngTmp = Printer.TextHeight("��") * 2
        If UCase(TypeName(objOut)) = UCase("Printer") Then
'            lngTmp = lngTmp - 50
        End If
    End If
    
    If lngY > lngHeight - IIf(lngBottom < 1134, 1134, lngBottom) - lngTmp Then
        Set CheckNewPage = NewPrintPage(objOut, lngPage, True)
'        CheckNewPage.Width = lngWidth: CheckNewPage.Height = lngHeight
        lngTop = mNewPageTop
        lngY = lngTop
    Else
        If lngDefHeight = -1 Then
            lngFontHeigh = Printer.TextHeight("��")
        Else
            lngFontHeigh = lngDefHeight
        End If
        lngY = lngY - lngFontHeigh
        If UCase(TypeName(objOut)) = UCase("Printer") Then
            Set CheckNewPage = Printer
        Else
            Set CheckNewPage = objOut.picPage(objOut.picPage.UBound)
        End If
    End If
End Function

Private Function PrintLineS(objDraw As Object, ByVal strChars As String, ByVal lngLeft As Long, ByVal lngRight As Long, lngNowPage As Long) As String
    '�������ǹ�������ӡ֮��:  ���ݵ�ǰλ�ô�ӡ�ı�����(���ݻػ�����߽�)����ʣ�µ��ַ����Ա��´δ�ӡ
    '
    Dim lngTmp As Long
    Dim strTmp As String
    Dim i As Long
    Dim lngWidth As Long
    Dim y As Long
    On Error GoTo ErrHandle
    
    lngWidth = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "���", Printer.Width)
    PrintLineS = ""
    
    '��ȥ���س���
    strChars = Replace(strChars, vbCrLf, vbCr)
    strChars = Replace(strChars, vbLf, vbCr)
    strTmp = ""
    lngRight = lngWidth - lngRight
    'ѭ�����ַ�
    For i = 1 To Len(strChars)
        '������λ��
        '��������λ��ֵ
        lngTmp = lngRight - lngLeft
        strTmp = strTmp & Mid(strChars, i, 1)
        '��������һ���ַ��Ŀ�ȳ����߽��Ȼ�����һ���ַ��ǻػ���ʱ�Ϳ�ʼ��ӡ��һ��,����ʣ�µ��ַ�����
        If (Printer.TextWidth(strTmp) > lngTmp Or Mid(strChars, i, 1) = vbCr) And _
            InStr("������������������������!%)}];:,.>?", Mid(strChars, i, 1)) = 0 Then
            If i = 1 Then
                PrintLineS = Mid(strChars, i + 1)
            Else
                strTmp = Left(strChars, i - 1)
                If Mid(strChars, i, 1) = vbCr Then
                    If i + 1 > Len(strChars) Then
                        PrintLineS = ""
                    Else
                        PrintLineS = Mid(strChars, i + 1)
                    End If
                Else
                    PrintLineS = Mid(strChars, i)
                End If
            End If
            '��ʱ�����˳�ѭ����ʼ��ӡ
            Exit For
        End If
    Next
    '��ʼ��ӡ�����ַ���
    objDraw.CurrentX = lngLeft
    y = objDraw.CurrentY
    If strTmp <> "" Then
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
        (lngNowPage - mPageNumber >= mPrintBegingPage - 1 And lngNowPage - mPageNumber <= mPrintEndPage) Then
            objDraw.Print strTmp
        End If
        objDraw.CurrentY = y + Printer.TextHeight(strTmp)
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PrintOutCase(objParent As Object, objOut As Object, ByVal lng�������� As Long, ByVal blnCurCase As Boolean, ByVal lngCurCase As Long, ByVal lng����ID As Long, _
    ByVal var��ҳ�򵥾� As Variant, ByVal blnPatiInfo As Boolean, ByVal lngY As Long, Optional ByVal lngҳ�� As Long = 0, Optional ByVal lng��ʼҳ As Long = 0, Optional ByVal lng����ҳ As Long = 0) As Boolean
    '���ܣ���ӡ���в���
    '������ObjParent        �����߶���
    '       ObjOut          ��������Ǵ�ӡ������Ԥ�����壩
    '       lng��������     ָ����������
    '       blnCurCase      �Ƿ�Ϊֻ��ӡ�����ǰ��ҳ
    '       lngCurCase      ָ����ǰ��ӡ������Ƿݲ�������ӡ���ʱ�ʹ��Ƿ������ӡ���
    '                       ����ʱ��ʾ������¼ID
    '       lng����id       ����Ǵ�ӡ����ʾ����ô,�������IDΪ0,���� var��ҳ�򵥾� ��Ϊ������¼ID
    '       var��ҳ�򵥾�   �����סԺ���˾ͼ�¼��ҳID����������ﲡ�˾ͼ�¼�Һŵ���ͨ�����������ж���סԺ��������
    '       blnPatiInfo     �Ƿ��ӡ������Ϣ
    '       lngY            ��ӡ��ʼ��Y����
    '       lngҳ��         ������ʼ��ҳ��,Ϊ0ʱ��ʾ����ӡҳ��
    
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim rsNewTmp As New ADODB.Recordset
    Dim i As Long
    Dim lngPrintPageNO As Long
    Dim lngBPage As Long    '��ʼҳ
    Dim lngEPage As Long   '����ҳ
    Dim sngBeginY As Single '��ʼλ��
    Dim sngEndY As Single '����λ��
    Dim lngTop As Long
    Dim lngHeight As Long
    Dim lngNewPage As Boolean           '�ϴ��Ƿ�����һҳ��ӡ
    Dim IntOnePage As Boolean           '�Ƿ��һҳ
    Dim bNowPrint As Boolean            '�Ƿ��ӡ��һҳ
        
    mPrintBegingPage = lng��ʼҳ
    mPrintEndPage = lng����ҳ
    mPageNumber = lngҳ��
        
    On Error GoTo ErrHandle
    '�õ����ϱ߾�,Ϊ��ҳ����׼��
    lngTop = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ϱ߾�", OFFSET_TOP) * 56.7
    lngHeight = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�߶�", Printer.Height)
    
    '��ʼ����ӡ����
    If InitPrint(objParent) = False Then
        MsgBox "��ӡ����ʼ��ʧ�ܣ����κο��ô�ӡ�����޷����ô�ӡֽ�ţ���", vbExclamation, gstrSysName
        Exit Function
    End If
    lngHeight = Printer.Height
    
    If lngCurCase < 0 Then
        '����ʱ��ʾ������¼ID
        strSQL = "select a.id,a.��д����,b.��ҳ from ���˲�����¼ a ,�����ļ�Ŀ¼ b where a.�ļ�ID = b.ID And a.ID=" & (-1 * lngCurCase)
        blnCurCase = True: lngCurCase = 1: var��ҳ�򵥾� = ""
    Else
        If lng����ID > 0 Then
            If VarType(var��ҳ�򵥾�) = vbString Then
                strSQL = "select a.id,a.��д����,b.��ҳ from ���˲�����¼ a ,�����ļ�Ŀ¼ b where a.�������� = " & lng�������� & " and a.�ļ�ID = b.ID and a.����ID = " & lng����ID & " AND a.�Һŵ�='" & var��ҳ�򵥾� & "' order by b.��ӡ˳��,a.��д���� "
            ElseIf IsNumeric(var��ҳ�򵥾�) Then
                strSQL = "select a.id,a.��д����,b.��ҳ from ���˲�����¼ a ,�����ļ�Ŀ¼ b where a.�������� = " & lng�������� & " and a.�ļ�ID = b.ID and a.����ID = " & lng����ID & " AND a.��ҳID=" & var��ҳ�򵥾� & "  order  by b.��ӡ˳��,a.��д���� "
            Else
                MsgBox "��ӡ���˲���ʱ��������ȷ�����ܼ�����", vbExclamation, gstrSysName
                Exit Function
            End If
        Else
            '���ﴦ����Щ�Բ���ʾ���Ĵ�ӡ
            If IsNumeric(var��ҳ�򵥾�) Then
                If 1 * var��ҳ�򵥾� > 0 Then
                    strSQL = "Select * from ����ʾ��Ŀ¼ where ID = " & var��ҳ�򵥾�
                Else
                    strSQL = "Select -1*ID As ID,������¼ID from ���˲����޶���¼ where ID = " & -1 * var��ҳ�򵥾�
                End If
            Else
                MsgBox "��ӡ����ʾ���ļ������ڣ����ܼ�����", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������ӡ")
    With rsTmp
        If .RecordCount > 0 Then
            .MoveFirst
            lngPrintPageNO = lngҳ��
            sngEndY = lngY
            IntOnePage = False
            For i = 1 To .RecordCount  '��ÿһ�ݲ�����������������ӡ�����Ĺ��̣������ز�������ӡ��һ�ݲ���
                '�õ���ӡǰ��ҳ����λ��
                If lng����ID > 0 And IsNumeric(var��ҳ�򵥾�) Then
                    'ֻ��סԺ�����ͻ����������ڴ�ӡ����ʱ����һҳ�Ŀ���,�����ڲ�����������ж� "�������� in (2,3)"
                    strSQL = "select a.����ID,a.��ҳID,a.�Һŵ�,nvl(b.��ҳ,0) ��ҳ from ���˲�����¼ a,�����ļ�Ŀ¼ b where a.�ļ�ID=b.id and  a.�������� in (2,3) and a.id=" & zlCommFun.Nvl(rsTmp!ID, 0)
                    Call zlDatabase.OpenRecordset(rsNewTmp, strSQL, "������ӡ")
                    If rsNewTmp.RecordCount > 0 Then
                        'ֻ���Ǹ�������������ҳ��,����,���ǵ�һҳ,���Ҳ���˳����Ǵ���ָ������˳��ŵĲ����ſ�������һҳ
                        If (rsNewTmp!��ҳ = 1 And i > 1 And blnCurCase = False And i > lngCurCase) Or lngNewPage = True Then
                            'Ҫ����һҳ��
                            sngBeginY = lngTop
                            If PrintOutCase = True Then
                                sngEndY = lngHeight - lngTop
                            End If
                            lngBPage = lngPrintPageNO + 1   'ȷ���ڱ���ʱ���Ƿݲ�����ҳ������һҳ
                            IntOnePage = True
                        Else
                            sngBeginY = sngEndY
                            lngBPage = lngPrintPageNO
                            If rsNewTmp!��ҳ = 1 And i = 1 And blnCurCase = False Then IntOnePage = True
                        End If
                    Else
                        sngBeginY = sngEndY
                        lngBPage = lngPrintPageNO
                    End If
                    If Not rsNewTmp.EOF Then
                        If rsNewTmp!��ҳ = 1 And IntOnePage <> False Then
                            lngNewPage = True
                        Else
                            lngNewPage = False
                        End If
                    Else
                        lngNewPage = False
                    End If
                Else
                    '����ļ�ʾ���Ķ�Ҫ������ӡ
                    sngBeginY = sngEndY
                    lngBPage = lngPrintPageNO
                End If
                '=====================================================================================
                '���ֻ���ָ�����Ƿݲ���ʱ���������˳�
                If blnCurCase = True Then
                    If i = lngCurCase Then
                        '��ʼ��ӡָ������
                        If UCase(TypeName(objParent)) = UCase("FRMCASEPRINT") Then
                            zlCommFun.ShowFlash "��" & .RecordCount & "�ݲ���������ӡ��" & i & "�ݣ� ���Ժ�... ..."
                        Else
                            zlCommFun.ShowFlash "��" & .RecordCount & "�ݲ���������ӡ��" & i & "�ݣ� ���Ժ�... ...", objParent
                        End If
                        If lng����ID > 0 Then
                            If PrintOrPreviewCase(objParent, objOut, !ID, blnPatiInfo, lngҳ�� > 0, lngPrintPageNO, sngEndY) Then
                                '�õ���ӡ�������λ�ã�λ�ñ�����Ǵ�ӡ���λ�ã�
                                lngEPage = lngPrintPageNO
                                '������ӡ����ʱ�����没��λ��
                                'If UCase(TypeName(objOut)) = "PRINTER" Then
                                '    strSQL = "zl_������ӡ��¼_insert(" & zlCommFun.NVL(!����ID, 0) & "," & zlCommFun.NVL(!��ҳID, 0) & ",'" & zlCommFun.NVL(!�Һŵ�) & "'," & lngBPage & "," & lngEPage & "," & sngBeginY & "," & sngEndY & ",'" & UserInfo.���� & "')"
                                '    Call zlDatabase.ExecuteProcedure(strSQL, "������ӡ")
                                'End If
                                PrintOutCase = True
                                bNowPrint = True
                            End If
                        Else
                            If PrintOrPreviewCase(objParent, objOut, !ID, False, lngҳ�� > 0, lngPrintPageNO, sngEndY, 1 * var��ҳ�򵥾� > 0) Then
                                '�õ���ӡ�������λ�ã�λ�ñ�����Ǵ�ӡ���λ�ã�
                                lngEPage = lngPrintPageNO
                                PrintOutCase = True
                                bNowPrint = True
                            End If
                        End If
                        zlCommFun.StopFlash
                        Exit Function
                    End If
                ElseIf i >= lngCurCase Then
                    '��ʼ��ӡÿ�ݲ���
                    If UCase(TypeName(objParent)) = UCase("FRMCASEPRINT") Then
                        zlCommFun.ShowFlash "��" & .RecordCount & "�ݲ���������ӡ��" & i & "�ݣ� �����" & Format((i - 1) / .RecordCount, "0.00") * 100 & "%"
                    Else
                        zlCommFun.ShowFlash "��" & .RecordCount & "�ݲ���������ӡ��" & i & "�ݣ� �����" & Format((i - 1) / .RecordCount, "0.00") * 100 & "%", objParent
                    End If
                    If lng����ID > 0 Then
                        If PrintOrPreviewCase(objParent, objOut, !ID, blnPatiInfo, lngҳ�� > 0, lngPrintPageNO, sngEndY) Then
                            '�õ���ӡ�������λ�ã�λ�ñ�����Ǵ�ӡ���λ�ã�
                            lngEPage = lngPrintPageNO
'                            If UCase(TypeName(objOut)) = "PRINTER" Then
                                strSQL = "zl_������ӡ��¼_insert(" & !ID & "," & lngBPage & "," & lngEPage & "," & sngBeginY & "," & sngEndY & ",'" & UserInfo.���� & "')"
                                Call zlDatabase.ExecuteProcedure(strSQL, "������ӡ")
'                            End If
                            PrintOutCase = True
                        Else
                            zlCommFun.StopFlash
                            '��ӡʧ��ֱ���˳�
                            Exit Function
                        End If
                    Else
                        If PrintOrPreviewCase(objParent, objOut, !ID, False, lngҳ�� > 0, lngPrintPageNO, sngEndY, 1 * var��ҳ�򵥾� > 0) Then
                            '�õ���ӡ�������λ�ã�λ�ñ�����Ǵ�ӡ���λ�ã�
                            lngEPage = lngPrintPageNO
                            PrintOutCase = True
                        Else
                            zlCommFun.StopFlash
                            '��ӡʧ��ֱ���˳�
                            Exit Function
                        End If
                    End If
                End If
                .MoveNext
            Next
        Else
            MsgBox "���κοɴ�ӡ�Ĳ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    zlCommFun.StopFlash
    PrintOutCase = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PrintOrPreviewCase(objParent As Object, objOut As Object, ByVal lng������¼ID As Long, ByVal blnPatiInfo As Boolean, _
    ByVal blnPrintNO As Boolean, lngEndPage As Long, sngEndY As Single, Optional ByVal blnDemo As Boolean = False) As Boolean
    '����:��ӡָ����������
    '����:  ObjParent       ��������
    '       objOut          �������
    '       lng������¼ID
    '       blnPatiInfo     �Ƿ��ӡ������Ϣ
    '       blnPrintNO      �Ƿ��ӡҳ��
    '       lngEndPage      �ϴε�ҳ��,�����ر��δ�ӡ�����ҳ��
    '       sngEndY         �ϴεĴ�ӡλ��,�����ر��εĴ�ӡλ��
    '       blnDemo         ������ʾ�ǲ��ǲ���ʾ��,�������ô lng������¼ID �ͱ�ʾ����ʾ��ID
    
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset, rsNewTmp1 As New ADODB.Recordset, rsNewTmp2 As New ADODB.Recordset
    Dim strԪ�ر��� As String
    Dim strSQL As String, i As Long, j As Long, m As Long
    Dim ObjStdPic As New StdPicture, lngStdPicWidth As Long, lngStdPicHeight As Long, dblPic���� As Double   'ȡͼƬ�ı���
    Dim objDraw As Object   '��ͼ�����������Ǵ�ӡ��Ҳ������ͼƬ�ؼ�
    Dim blnPrint As Boolean '�ж��ǲ��Ǵ�ӡ��
    Dim x As Long, y As Long, TmpX As Long, TmpY As Long, H_9pt As Long, W_9pt As Long, Tmp_W As Long, Tmp_H As Long
    Dim strFontName As String, strFontSize As String, strFontBold As String, strFontItalic As String
    Dim strTitleFontName As String, strTitleFontSize As String, strTitleFontItalic As String, strTitleFontBold As String, strTitleAlig As String   '����Ķ��뷽ʽ
    Dim strText As String
    Dim lng����ID As Long, lng��ҳID As Long, blnOutPati As Boolean   '��סԺ��������
    Dim lngLeft As Long, lngRight As Long, lngTop As Long, lngBottom As Long, lngWidth     As Long, lngHeight  As Long
    Dim lngTmpPageNo As Long    '��¼��ʱ��ҳ��
    Dim lngPageTmp  As Long     '���ڼ�¼��ʱ��ҳ�������ж�
    Dim tmpPage As Integer, tmpPrintHeight As Long      '��ʱ��¼����
    Dim blOnePrintText As Boolean                       '�Ƿ��һ�δ�ӡ�ı���
    Dim blnMultiSign As Boolean '�Ƿ��ж��ǩ��
    
    '�õ�ֽ�ŵı߽�����
    lngLeft = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��߾�", OFFSET_LEFT) * 56.7 + Screen.TwipsPerPixelX * 2
    lngRight = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ұ߾�", OFFSET_RIGHT) * 56.7 - Screen.TwipsPerPixelX * 2
    lngTop = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ϱ߾�", OFFSET_TOP) * 56.7 + Screen.TwipsPerPixelY * 2
    lngBottom = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�±߾�", OFFSET_BOTTOM) * 56.7 - Screen.TwipsPerPixelY * 2
    lngWidth = Printer.Width ' GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "���", Printer.Width)
    lngHeight = Printer.Height ' GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�߶�", Printer.Height)
    If blnDemo = False Then
        '�жϲ����ǲ��Ǵ��ڣ�����������ID
        If lng������¼ID > 0 Then
            strSQL = "select a.*,b.���� ��������,c.����,c.סԺ��,c.�����,d.��Ժ���� As ��ǰ���� from ���˲�����¼ a,���ű� b , ������Ϣ c,������ҳ d" & _
                " Where a.����ID=b.id(+) and a.����id = c.����id" & _
                " And a.����ID=d.����ID(+) And a.��ҳID=d.��ҳID(+)" & _
                " And a.id =" & lng������¼ID
        Else
            strSQL = "select a.*,b.���� ��������,d.����,d.סԺ��,d.�����,e.��Ժ���� As ��ǰ���� from ���˲�����¼ a,���ű� b,���˲����޶���¼ c ,������Ϣ d,������ҳ e" & _
                " Where a.����ID=b.id(+) and a.id=c.������¼id and a.����id = d.����id" & _
                " And a.����ID=e.����ID(+) And a.��ҳID=e.��ҳID(+)" & _
                " And c.id =" & -1 * lng������¼ID
        End If
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������ӡ")
        If rsTmp.RecordCount > 0 Then
            lng����ID = zlCommFun.Nvl(rsTmp!����id, 0)
            lng��ҳID = zlCommFun.Nvl(rsTmp!��ҳID, 0)
            '��Ϊ����
            If lng��ҳID = 0 Then: blnOutPati = True: Else blnOutPati = False
        Else
            MsgBox "ָ�����������ڣ�", vbExclamation, gstrSysName: Exit Function
        End If
    Else
        '����ʾ��
        strSQL = "select * from ���˲������� where ����ʾ��ID=" & lng������¼ID & " order by �������"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������ӡ")
        If rsTmp.RecordCount < 1 Then
            MsgBox "ָ������ʾ�����κ����ݣ�", vbExclamation, gstrSysName: Exit Function
        End If
        strSQL = "select a.���� ��������,a.�ƶ��� ��д��,a.�ƶ��� ��д���� ,b.���� ��������  FROM ����ʾ��Ŀ¼ a, ���ű� b where  a.����ID=b.id(+) and  a.id=" & lng������¼ID
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������ӡ")
        If rsTmp.RecordCount < 1 Then
            MsgBox "ָ������ʾ�������ڣ�", vbExclamation, gstrSysName: Exit Function
        End If
    End If
    '��ʼ��һ����ҳ��
    Set objDraw = Nothing
    If blnPrintNO Then
        Set objDraw = NewPrintPage(objOut, lngEndPage, False)
    Else
        lngTmpPageNo = 0
        Set objDraw = NewPrintPage(objOut, lngEndPage, False)
    End If
    '����ӡ��
    If objDraw Is Nothing Then
        MsgBox "��ӡ�������˳���ӡ��", vbExclamation, gstrSysName: Exit Function
    End If
    blnPrint = UCase(TypeName(objDraw)) = "PRINTER"
    If blnPrint = False Then
        objDraw.Width = lngWidth: objDraw.Height = lngHeight
    End If
    mNewPageInit = sngEndY
    objDraw.CurrentY = IIf(sngEndY >= lngTop, sngEndY, lngTop)
    
    
    objDraw.Font.Name = "����"
    objDraw.Font.Size = 9
    objDraw.Font.Bold = False
    SetPrinterFont objDraw.Font, 9
    H_9pt = Printer.TextHeight("��")
    W_9pt = Printer.TextWidth("��")
    '��ӡ��ǰ����������Ϣ
    objDraw.ForeColor = 0
    objDraw.Font.Name = "����"
    objDraw.Font.Size = 18
    objDraw.Font.Bold = True
    SetPrinterFont objDraw.Font, 18
    '��������
    If blnDemo = False Then
        strText = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "��λ����")
    Else
        strText = zlCommFun.Nvl(rsTmp!��������)
    End If
    '�ж��Ƿ���ҳ
    y = objDraw.CurrentY + H_9pt * 2 + Printer.TextHeight("��") '��ʼ����2���ָ�
    tmpPrintHeight = Printer.TextHeight("��")
    Set objDraw = Nothing
    lngPageTmp = lngEndPage
    If blnPrintNO Then
        Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
    Else
        lngTmpPageNo = 0
        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
    End If
    If objDraw Is Nothing Then Exit Function
    
    If lngPageTmp = lngEndPage And objDraw.CurrentY = lngTop Then
        '�õ�����XY����
        x = lngLeft + (lngWidth - (lngLeft + lngRight) - Printer.TextWidth(strText)) / 2
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        objDraw.CurrentY = y + Printer.TextHeight(strText)
        
        '��ӡ��������,ǩ��,����
        objDraw.ForeColor = 0
        objDraw.Font.Name = "����"
        objDraw.Font.Size = 10.5
        objDraw.Font.Bold = False
        SetPrinterFont objDraw.Font, 10.5
        
        '�ж��Ƿ���ҳ
        y = objDraw.CurrentY + H_9pt * 2 + Printer.TextHeight("��") '��ʼ����2���ָ�
        tmpPrintHeight = Printer.TextHeight("��")
        Set objDraw = Nothing
        If blnPrintNO Then
            Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
        Else
            lngTmpPageNo = 0
            Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
        End If
        If objDraw Is Nothing Then Exit Function
        
        If blnDemo = False Then
            strText = "����:" & zlCommFun.Nvl(rsTmp!��������)
        Else
            strText = "���ÿ���:" & IIf(Trim(zlCommFun.Nvl(rsTmp!��������)) = "", "���п���", zlCommFun.Nvl(rsTmp!��������))
        End If
        x = lngLeft
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        mPageHeadDep = strText
        If blnDemo = False Then
            strText = "������:" & zlCommFun.Nvl(rsTmp!����)
        Else
            strText = "���ƶ���:" & zlCommFun.Nvl(rsTmp!��д��)
        End If
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        mPageHeadName = strText
        If blnDemo = False Then
            If blnOutPati Then
                strText = "���������:" & rsTmp!�����
            Else
                strText = "    סԺ��:" & rsTmp!סԺ��
            End If
        Else
            strText = "�ƶ�����:" & IIf(IsNull(rsTmp!��д����), "", Format(rsTmp!��д����, "YYYY��MM��DD��"))
        End If
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (4 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        mPageHeadNo = strText
        If blnDemo = False Then
            strText = "����:" & zlCommFun.Nvl(rsTmp!��ǰ����)
        Else
            strText = "����:"
        End If
        x = lngLeft + (lngWidth - lngLeft - lngRight) * (7 / 9)
        objDraw.CurrentX = x: objDraw.CurrentY = y
        objDraw.FontTransparent = True
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Print strText
        End If
        objDraw.CurrentY = y + Printer.TextHeight(strText)
        mPageBedNumber = strText
        
        y = objDraw.CurrentY + H_9pt / 5: x = lngLeft
        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
            objDraw.Line (lngLeft, y)-(lngWidth - lngRight, y), 0
        End If
    End If
    sngEndY = objDraw.CurrentY
    '�ò�����ҳ�Ƿ��ӡ������Ϣ,����ǲ���ʾ����ô lng����ID һ��Ϊ0 ��ʱ�Ͳ���ӡ������Ϣ
    If blnPatiInfo And lng����ID > 0 Then
        '����������Ϣ
        Set rsTmp = ReadPatiInfo(lng����ID, lng��ҳID)
        If Not rsTmp Is Nothing Then
            '�ֿ���ӡ�����סԺ������Ϣ
            objDraw.Font.Name = "����"
            objDraw.Font.Size = 10.5
            objDraw.Font.Bold = False
            SetPrinterFont objDraw.Font, 10.5
            '�ж��Ƿ���ҳ
            y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("��")  '�����м��Ϊ1/3����ͨ�ָ�
            tmpPrintHeight = Printer.TextHeight("��")
            Set objDraw = Nothing
            If blnPrintNO Then
                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
            Else
                lngTmpPageNo = 0
                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
            End If
            If objDraw Is Nothing Then Exit Function
            
            '����ID,��ʶ��,(����)
            x = lngLeft
            strText = "������ID:" & lng����ID
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            'סԺ������������Ϣ��ͬ
            If blnDemo = False Then
                If blnOutPati Then
                    x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
                    strText = "�������:" & rsTmp!�����
                    Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                Else
                    x = lngLeft + (lngWidth - lngLeft - lngRight) * (1 / 3)
                    strText = "��סԺ��:" & rsTmp!סԺ��
                    Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                    
                    x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
                    strText = "��������:" & zlCommFun.Nvl(rsTmp!��Ժ����)
                    Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                End If
            End If
            '�ж��Ƿ���ҳ
            y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("��")
            tmpPrintHeight = Printer.TextHeight("��")
            Set objDraw = Nothing
            If blnPrintNO Then
                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
            Else
                lngTmpPageNo = 0
                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
            End If
            If objDraw Is Nothing Then Exit Function
            
            '����,�Ա�,����
            x = lngLeft
            strText = "��������:" & zlCommFun.Nvl(rsTmp!����)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            x = lngLeft + (lngWidth - lngLeft - lngRight) * (1 / 3)
            strText = "�����Ա�:" & zlCommFun.Nvl(rsTmp!�Ա�)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
            strText = "��������:" & zlCommFun.Nvl(rsTmp!����)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            '�ж��Ƿ���ҳ
            y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("��")
            tmpPrintHeight = Printer.TextHeight("��")
            Set objDraw = Nothing
            If blnPrintNO Then
                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
            Else
                lngTmpPageNo = 0
                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
            End If
            If objDraw Is Nothing Then Exit Function
            
            '����,ְҵ,����״��
            x = lngLeft
            strText = "��������:" & zlCommFun.Nvl(rsTmp!����)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            x = lngLeft + (lngWidth - lngLeft - lngRight) * (1 / 3)
            strText = "����ְҵ:" & zlCommFun.Nvl(rsTmp!ְҵ)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
            strText = "����״��:" & zlCommFun.Nvl(rsTmp!����״��)
            Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            '�ж��Ƿ���ҳ
            y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("��")
            tmpPrintHeight = Printer.TextHeight("��")
            Set objDraw = Nothing
            If blnPrintNO Then
                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
            Else
                lngTmpPageNo = 0
                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
            End If
            If objDraw Is Nothing Then Exit Function
            
            '������λ
            x = lngLeft
            strText = "������λ:" & zlCommFun.Nvl(rsTmp!������λ)
            Call DrawCell(objDraw, strText, x, y, lngWidth - lngLeft - lngRight, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            '�ж��Ƿ���ҳ
            y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("��")
            sngEndY = y
            tmpPrintHeight = Printer.TextHeight("��")
            Set objDraw = Nothing
            If blnPrintNO Then
                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
            Else
                lngTmpPageNo = 0
                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
            End If
            If objDraw Is Nothing Then Exit Function
            
            '��ͥ��ַ
            x = lngLeft
            strText = "��ͥ��ַ:" & zlCommFun.Nvl(rsTmp!��ͥ��ַ)
            Call DrawCell(objDraw, strText, x, y, lngWidth - lngLeft - lngRight, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            If blnOutPati = False And blnDemo = False Then
                '�ж��Ƿ���ҳ
                y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("��")
                tmpPrintHeight = Printer.TextHeight("��")
                Set objDraw = Nothing
                If blnPrintNO Then
                    Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                Else
                    lngTmpPageNo = 0
                    Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                End If
                If objDraw Is Nothing Then Exit Function
                
                '��Ժ����,��Ժ����
                x = lngLeft
                strText = "��Ժ����:" & zlCommFun.Nvl(rsTmp!��Ժ����)
                Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
                strText = "��Ժ����:" & IIf(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy��MM��dd��"))
                Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                '�ж��Ƿ���ҳ
                y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("��")
                tmpPrintHeight = Printer.TextHeight("��")
                Set objDraw = Nothing
                If blnPrintNO Then
                    Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                Else
                    lngTmpPageNo = 0
                    Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                End If
                If objDraw Is Nothing Then Exit Function
                
                '��Ժ����,��Ժ����
                x = lngLeft
                strText = "��Ժ����:" & zlCommFun.Nvl(rsTmp!��Ժ����)
                Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
                x = lngLeft + (lngWidth - lngLeft - lngRight) * (2 / 3)
                strText = "��Ժ����:" & IIf(IsNull(rsTmp!��Ժ����), "", Format(rsTmp!��Ժ����, "yyyy��MM��dd��"))
                Call DrawCell(objDraw, strText, x, y, (lngWidth - lngLeft - lngRight) / 3, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", 0)
            End If
            y = objDraw.CurrentY + Printer.TextHeight("��") / 5: x = lngLeft
            If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
            (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
            mPageNumber <= mPrintEndPage) Then
                objDraw.Line (lngLeft, y)-(lngWidth - lngRight, y), 0
            End If
            objDraw.CurrentY = y
            sngEndY = y
        End If
        '��ɱ����ӡ��������Ϣ�Ĵ�ӡ
    End If
    '��ʼ׼�������Ĵ�ӡ
    If blnDemo = False Then
        If lng������¼ID > 0 Then
            strSQL = "select * from ���˲������� where ������¼id in (select id from ���˲�����¼ where �������� not in (-1,-2) and id =" & lng������¼ID & ") ORDER BY �������"
        Else
            strSQL = "select * from ���˲������� where �����޶�id=" & -1 * lng������¼ID & " ORDER BY �������"
        End If
    Else
        strSQL = "select * from ���˲������� where ����ʾ��ID=" & lng������¼ID & " order by �������"
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������ӡ")
    With rsTmp
        '�ж��Ƿ��ж��ǩ��
        .Filter = "Ԫ������=-1": blnMultiSign = (.RecordCount > 1): .Filter = ""
        
        If .RecordCount > 0 Then
            .MoveFirst
            For i = 0 To .RecordCount - 1
                '��������Ԫ�ر��룬�Ա�ȷ�����ĸ�����Ԫ��
                strSQL = " SELECT * FROM ����Ԫ��Ŀ¼ where ����=" & !Ԫ������ & " and ����='" & zlCommFun.Nvl(!Ԫ�ر���) & "'"
                Call zlDatabase.OpenRecordset(rsNewTmp2, strSQL, "������ӡ")
                If rsNewTmp2.RecordCount > 0 Then
                    strԪ�ر��� = zlCommFun.Nvl(rsNewTmp2!����) & "_" & zlCommFun.Nvl(rsNewTmp2!����)
                Else
                    strԪ�ر��� = ""
                End If
                '���ȴ������Ĵ�ӡ(��������Щ),��������Ԫ�������Լ��Ĺ����е�������
                objDraw.ForeColor = zlCommFun.Nvl(!������ɫ, 0)
                strTitleFontName = zlCommFun.Nvl(!��������)
                If Trim(strTitleFontName) = "" Then strTitleFontName = "����"
                strTitleFontSize = "10.5": strTitleFontBold = "1": strTitleFontItalic = "0"
                For j = 1 To UBound(Split(strTitleFontName, ","))
                    Select Case j
                    Case 1
                        strTitleFontSize = Val(Split(strTitleFontName, ",")(j))
                    Case 2
                        strTitleFontBold = CLng(Split(strTitleFontName, ",")(j))
                    Case 3
                        strTitleFontItalic = CLng(Split(strTitleFontName, ",")(j))
                    End Select
                Next
                strTitleFontName = Split(strTitleFontName, ",")(0)
                '�õ�λ����ʱ���ڱ����ﱣ��
                strTitleAlig = zlCommFun.Nvl(!����λ��, 1)
                If !������ʾ = 1 And (!Ԫ������ >= 0 Or !Ԫ������ = -5) Then
                    objDraw.Font.Name = strTitleFontName
                    objDraw.Font.Size = Format(strTitleFontSize)
                    objDraw.Font.Bold = IIf(strTitleFontBold = "1", True, False)
                    objDraw.Font.Italic = IIf(strTitleFontItalic = "1", True, False)
                    SetPrinterFont objDraw.Font, Format(strTitleFontSize)
                    '�õ������ı�׼����ӡ
                    strText = zlCommFun.Nvl(!�����ı�)
                    strText = IIf(Trim(strText) <> "", strText & "��", strText)
                    '�Խ����ļ�ð��
                    '                    strText = IIf(CLng(strTitleAlig) = 1 And zlCommFun.NVL(!����λ��, 1) = 1 And zlCommFun.NVL(!Ƕ�뷽ʽ, 2) = 1, strText & "��", strText)
                    '���ݱ���λ���������XY����
                    Select Case CLng(strTitleAlig)
                    Case 1  '��
                        x = lngLeft
                    Case 2  '��
                        x = lngLeft + (lngWidth - (lngLeft + lngRight) - Printer.TextWidth(strText)) / 2
                    Case 3  '��
                        x = lngWidth - lngRight - Printer.TextWidth(strText)
                    End Select
                    '�ж��Ƿ���ҳ
                    y = objDraw.CurrentY + H_9pt / 2 + Printer.TextHeight("��")
                    tmpPrintHeight = Printer.TextHeight("��")
                    Set objDraw = Nothing
                    If blnPrintNO Then
                        Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                    Else
                        lngTmpPageNo = 0
                        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                    End If
                    If objDraw Is Nothing Then Exit Function
                    
                    TmpY = y '������һλ��,�Ա��������ʽ��ӡ
                    sngEndY = y
                    If CLng(strTitleAlig) = 1 Then
                        'ֻ����������������,���Ҷ�������о��ڴ�ʱ���������ַ���������
                        Call DrawCell(objDraw, strText, x, y, lngWidth - lngLeft - lngRight, Printer.TextHeight("��"), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , , , objDraw.Font, "0000", CLng(strTitleAlig) - 1)
                        TmpX = x + Printer.TextWidth(strText) + Printer.TextWidth("��") / 2
                    Else
                        '�����������뷽ʽ
                        objDraw.CurrentX = x: objDraw.CurrentY = y: objDraw.FontTransparent = True
                        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                        (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
                        mPageNumber <= mPrintEndPage) Then
                            objDraw.Print strText
                        End If
                        objDraw.CurrentY = y + Printer.TextHeight(strText)
                        TmpX = lngLeft
                        objDraw.CurrentX = TmpX
                    End If
                    sngEndY = objDraw.CurrentY
                End If
                Select Case True
                    '�ı��Ρ���ת�ı��ĸ��ӱ�����������Ϊ����ͼ�󱨸��ר��ֽ
                Case !Ԫ������ = 0 Or !Ԫ������ = -5 _
                    Or (!Ԫ������ = 1 And IIf(IsNull(!�ı�ת��), 0, !�ı�ת��) = 1) _
                    Or (!Ԫ������ = 2 And IIf(IsNull(!�ı�ת��), 0, !�ı�ת��) = 1) _
                    Or (!Ԫ������ = 4 And Trim(zlCommFun.Nvl(!�����ı�)) <> "���ͼ�󱨸�")
                    'ʵ�ַ���:ֱ�Ӷ����ַ���,����һ����ÿ���ַ���ȡ���ַ���,
                    '���պù�һ�еĳ���ʱ�Ϳ�ʼ��ӡ��һ��,������ȡ�ַ�,���
                    'ͬʱ�ڴ�ӡǰ���ж��ǲ��ǿ�ʼ�µ�һҳ
                    '�����ַ�����
                    objDraw.ForeColor = zlCommFun.Nvl(!������ɫ, 0)
                    strFontName = zlCommFun.Nvl(!��������)
                    strFontSize = "9"
                    strFontBold = "0"
                    strFontItalic = "0"
                    For j = 1 To UBound(Split(strFontName, ","))
                        Select Case j
                        Case 1
                            strFontSize = Val(Split(strFontName, ",")(j))
                        Case 2
                            strFontBold = CLng(Split(strFontName, ",")(j))
                        Case 3
                            strFontItalic = CLng(Split(strFontName, ",")(j))
                        End Select
                    Next
                    If Trim(strFontName) = "" Then strFontName = "����"
                    strFontName = Split(strFontName, ",")(0)
                    objDraw.Font.Name = strFontName
                    objDraw.Font.Size = Format(strFontSize)
                    objDraw.Font.Bold = IIf(strFontBold = "1", True, False)
                    objDraw.Font.Italic = IIf(strFontItalic = "1", True, False)
                    SetPrinterFont objDraw.Font, Format(strFontSize)
                    '��������ı������ݲ����䱣�����ı�������
                    strText = ""
                    strSQL = "select * from ���˲����ı��� where ����ID=" & !ID & " order by �к�"
                    Call zlDatabase.OpenRecordset(rsNewTmp1, strSQL, "������ӡ")
                    If rsNewTmp1.RecordCount > 0 Then
                        For j = 0 To rsNewTmp1.RecordCount - 1
                            strText = strText & zlCommFun.Nvl(rsNewTmp1!����)
                            rsNewTmp1.MoveNext
                        Next
                    End If
                    '����ר��ֻֽ��ӡ����ת�������ı�
                    If !Ԫ������ = 4 Then
                        TmpX = lngLeft
                        objDraw.CurrentY = sngEndY
                        y = objDraw.CurrentY + H_9pt / 2  '����һ��С�������
                        objDraw.CurrentY = y
                        sngEndY = y
                    Else
                        '����Ƕ�뷽ʽ�ֱ���
                        'Ҫ�������⣬��ô��������������ģ������⻹Ӧ��Ҫ��ʾ�ģ����һ���������ʾ��һ���ַ��Ŀ�ȣ���������λ�ñ����Ƕ����
                        If CLng(strTitleAlig) = 1 And zlCommFun.Nvl(!����λ��, 1) = 1 And zlCommFun.Nvl(!Ƕ�뷽ʽ, 2) = 1 And !������ʾ = 1 And Printer.TextWidth(zlCommFun.Nvl(!�����ı�)) < lngWidth - (lngLeft + lngRight) - Printer.TextWidth("��") Then
                            '1 ��������
                            'ֻ����ʾ���Ⲣ��,�����ı��ĳ���С�ڴ�ӡ����Ŀ�ȼ�ȥһ���ֵĿ��ʱ�ſ��Խ��������ӡ
                            objDraw.CurrentY = TmpY
                            If strText = "" Then '������ı���û�����ݾ�����һ��
                                y = objDraw.CurrentY + H_9pt  '����һ��С�������
                                objDraw.CurrentY = y
                                sngEndY = y
                            End If
                        Else
                            '2 ����һ��, 3 �ı�����
                            'Ŀǰ������ʽ������Ϊ����һ������ӡ
                            'todo:�Ժ������ı����Ƶķ�ʽ
                            objDraw.CurrentY = sngEndY
                            TmpX = lngLeft
                            y = objDraw.CurrentY + H_9pt / 2 '����һ��С�������
                            objDraw.CurrentY = y
                            sngEndY = y
                        End If
                    End If
                    blOnePrintText = True
                    Do While strText <> ""
                        '�ж��Ƿ���ҳ
                        If blOnePrintText = True Then
                            y = objDraw.CurrentY + Printer.TextHeight("��") + 30
                        Else
                            y = objDraw.CurrentY + H_9pt / 2 + Printer.TextHeight("��")
                        End If
                        blOnePrintText = False
                        tmpPrintHeight = Printer.TextHeight("��")
                        
                        If (Printer.TextWidth("��") * Len(strText) / 0.7) > (lngWidth - (lngRight * 2)) Then
                            Set objDraw = Nothing
                            If blnPrintNO Then
                                Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                            Else
                                lngTmpPageNo = 0
                                Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                            End If
                            If objDraw Is Nothing Then Exit Function
                            objDraw.CurrentY = y  '+ H_9pt
                        Else
                            objDraw.CurrentY = y - Printer.TextHeight("��")
                        End If
                        sngEndY = y
                        strText = PrintLineS(objDraw, strText, TmpX, lngRight, IIf(blnPrintNO, lngEndPage, lngTmpPageNo))
                        TmpX = lngLeft
                    Loop
                    sngEndY = objDraw.CurrentY
                    '����ת�ı��ĸ��ӱ�
                Case !Ԫ������ = 1 And IIf(IsNull(!�ı�ת��), 0, !�ı�ת��) = 0
                    '�ж��Ƿ���ҳ
                    y = objDraw.CurrentY + H_9pt / 2 + Printer.TextHeight("��")
                    tmpPrintHeight = Printer.TextHeight("��")
                    Set objDraw = Nothing
                    If blnPrintNO Then
                        Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                    Else
                        lngTmpPageNo = 0
                        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                    End If
                    If objDraw Is Nothing Then Exit Function
                    
                    objDraw.CurrentY = y  '+ H_9pt
                    objDraw.ForeColor = zlCommFun.Nvl(!������ɫ, 0)
                    strFontName = zlCommFun.Nvl(!��������)
                    strFontSize = "9"
                    strFontBold = "0"
                    strFontItalic = "0"
                    For j = 1 To UBound(Split(strFontName, ","))
                        Select Case j
                            Case 1
                                strFontSize = Val(Split(strFontName, ",")(j))
                            Case 2
                                strFontBold = CLng(Split(strFontName, ",")(j))
                            Case 3
                                strFontItalic = CLng(Split(strFontName, ",")(j))
                        End Select
                    Next
                    If Trim(strFontName) = "" Then strFontName = "����"
                    strFontName = Split(strFontName, ",")(0)
                    objDraw.Font.Name = strFontName
                    objDraw.Font.Size = Format(strFontSize)
                    objDraw.Font.Bold = IIf(strFontBold = "1", True, False)
                    objDraw.Font.Italic = IIf(strFontItalic = "1", True, False)
                    SetPrinterFont objDraw.Font, Format(strFontSize)
                    '�������ж���������ݲ���ͼ
                    Call GridDraw(objOut, objDraw, !ID, y, blnPrintNO, lngEndPage, zlCommFun.Nvl(!����λ��, 1) - 1)
                    '����һ��С�������
                    y = objDraw.CurrentY + H_9pt / 2
                    objDraw.CurrentY = y:      sngEndY = y
                    '����ת�ı���������
                Case !Ԫ������ = 2 And IIf(IsNull(!�ı�ת��), 0, !�ı�ת��) = 0
                    '�Ѿ�������ָ����ȷ�Χ�ڵ����д�ӡ�����뷽ʽ���Ƿ���ҳ�ȴ���
                    '�����ǰҳ�ܴ��¾��ڵ�ǰҳ��ӡ
                    '�����ǰҳ���¾�����ҳ��ӡ,�����ҳ������˵����ֽ�Ż�ֽ�Ŵ�С���õ������ֻ������ҳ��.
                    m = zlCommFun.Nvl(!������ɫ, 0)
                    strFontName = zlCommFun.Nvl(!��������)
                    strFontSize = "9"
                    strFontBold = "0"
                    strFontItalic = "0"
                    For j = 1 To UBound(Split(strFontName, ","))
                        Select Case j
                        Case 1
                            strFontSize = Val(Split(strFontName, ",")(j))
                        Case 2
                            strFontBold = CLng(Split(strFontName, ",")(j))
                        Case 3
                            strFontItalic = CLng(Split(strFontName, ",")(j))
                        End Select
                    Next
                    If Trim(strFontName) = "" Then strFontName = "����"
                    strFontName = Split(strFontName, ",")(0)
                    '�õ��������ĸ߶�����
                    strSQL = "SELECT nvl(MAX(��+��),0)+30 �������� ,nvl(MAX(��+��),0)+30 �������� FROM ���˲��������� WHERE ����id=" & !ID
                    Call zlDatabase.OpenRecordset(rsNewTmp1, strSQL, "������ӡ")
                    With rsNewTmp1
                        Tmp_W = rsNewTmp1!�������� + W_9pt * 2
                        Tmp_H = rsNewTmp1!�������� + H_9pt * 2
                    End With
                    '����λ��
                    Select Case zlCommFun.Nvl(!����λ��, 1)
                        Case 1  '��
                            TmpX = lngLeft
                        Case 2  '��
                            TmpX = lngLeft + (lngWidth - (lngLeft + lngRight) - Tmp_W) / 2
                        Case 3  '��
                            TmpX = lngWidth - (lngRight + Tmp_W)
                    End Select
                    y = objDraw.CurrentY + H_9pt / 2 '����һ��С�������
                    y = y + Tmp_H '�����µĵ�ǰYλ��Ϊ��ǰYλ�ü����������ĸ߶�,�ں��潫�����ж��ǲ���Ҫ��ҳ��
                    '�ж��Ƿ���ҳ
                    Set objDraw = Nothing
                    If blnPrintNO Then
                        Set objDraw = CheckNewPage(objOut, lngEndPage, y, Tmp_H)
                    Else
                        lngTmpPageNo = 0
                        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, Tmp_H)
                    End If
                    If objDraw Is Nothing Then Exit Function
                    
                    sngEndY = y
                    objDraw.ForeColor = m       'm��ʱ������¼��ɫ
                    objDraw.Font.Name = strFontName
                    objDraw.Font.Size = Format(strFontSize)
                    objDraw.Font.Bold = IIf(strFontBold = "1", True, False)
                    objDraw.Font.Italic = IIf(strFontItalic = "1", True, False)
                    SetPrinterFont objDraw.Font, Format(strFontSize)
                    strSQL = "Select ����,��������,������λ,��,��,��,�� From ���˲��������� where ����ID=" & !ID & " order by ��,��"
                    Call zlDatabase.OpenRecordset(rsNewTmp1, strSQL, "������ӡ")
                    With rsNewTmp1
                        Do While Not .EOF
                            x = TmpX + zlCommFun.Nvl(!��, 0)
                            y = sngEndY + zlCommFun.Nvl(!��, 0)
                            Call DrawCell(objDraw, zlCommFun.Nvl(!����) & " " & zlCommFun.Nvl(!��������) & " " & zlCommFun.Nvl(!������λ), x, y, zlCommFun.Nvl(!��, 0) + W_9pt, zlCommFun.Nvl(!��, 0), IIf(blnPrintNO, lngEndPage, lngTmpPageNo), , , , m, , objDraw.Font, "0000", 0, 1, True)
                            .MoveNext
                        Loop
                    End With
                    y = objDraw.CurrentY + H_9pt / 2 '����һ��С�������
                    objDraw.CurrentY = y:      sngEndY = y
                    'Ϊ����ͼ�󱨸��ר��ֽ
                Case !Ԫ������ = 4 And Trim(zlCommFun.Nvl(!�����ı�)) = "���ͼ�󱨸�"
                    objDraw.Font.Name = "����": objDraw.Font.Size = 9: objDraw.Font.Bold = False: objDraw.Font.Italic = False
                    SetPrinterFont objDraw.Font, 9
                    TmpX = lngLeft
                    objDraw.CurrentY = sngEndY: objDraw.CurrentX = TmpX
                    y = objDraw.CurrentY + H_9pt / 2 + Printer.TextHeight("��") '���ư��С�������
                    tmpPrintHeight = Printer.TextHeight("��")
                    Set objDraw = Nothing
                    If blnPrintNO Then
                        Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                    Else
                        lngTmpPageNo = 0
                        Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                    End If
                    If objDraw Is Nothing Then Exit Function
                    
                    objDraw.FontTransparent = True
                    If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                    (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
                    mPageNumber <= mPrintEndPage) Then
                        objDraw.Print "    �������ͼ��"
                    End If
                    objDraw.CurrentY = y + Printer.TextHeight("    �������ͼ��")
                    y = objDraw.CurrentY + H_9pt    '����һ��С�������
                    objDraw.CurrentY = y:   sngEndY = y
                    '���ͼ
                Case !Ԫ������ = 3
                    Set ObjStdPic = GetMap(!ID, frmFlash.picTmp)
                    If Not (ObjStdPic Is Nothing Or ObjStdPic = 0) Then
                        TmpX = lngLeft
                        objDraw.CurrentY = sngEndY
                        objDraw.CurrentX = TmpX
                        '�õ�����ֽ������
                        m = lngHeight - lngTop - lngBottom - Screen.TwipsPerPixelY * 6
                        '�õ�ͼƬ�ĸ�
                        lngStdPicHeight = objDraw.ScaleY(ObjStdPic.Height, vbHimetric, objDraw.ScaleMode)
                        '�õ�ͼƬ�Ŀ�
                        lngStdPicWidth = objDraw.ScaleX(ObjStdPic.Width, vbHimetric, objDraw.ScaleMode)
                        '�õ�����ߵı�
                        dblPic���� = ObjStdPic.Width / ObjStdPic.Height
                        '������ͼƬ��
                        If lngStdPicHeight > m Then
                            lngStdPicHeight = m
                            '�ٵõ���
                            lngStdPicWidth = lngStdPicHeight * dblPic����
                        End If
                        If lngStdPicWidth > lngWidth - lngLeft - lngRight - Screen.TwipsPerPixelX * 3 Then
                            lngStdPicWidth = lngWidth - lngLeft - lngRight - Screen.TwipsPerPixelX * 3
                            lngStdPicHeight = lngStdPicWidth / dblPic����
                        End If
                        '����ȷ��X����
                        If lngStdPicWidth < lngWidth - (lngLeft + lngRight + Screen.TwipsPerPixelX * 2) Then
                            TmpX = lngLeft + (lngWidth - (lngLeft + lngRight) - lngStdPicWidth) / 2 - Screen.TwipsPerPixelX * 2
                        Else
                            TmpX = lngLeft
                        End If
                        
                        y = objDraw.CurrentY + lngStdPicHeight + H_9pt / 2  '���ư��С�������
                        Set objDraw = Nothing
                        If blnPrintNO Then
                            Set objDraw = CheckNewPage(objOut, lngEndPage, y, lngStdPicHeight)
                        Else
                            lngTmpPageNo = 0
                            Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, lngStdPicHeight)
                        End If
                        If objDraw Is Nothing Then Exit Function
                        
                        objDraw.CurrentY = y + Screen.TwipsPerPixelY * 2
                        y = objDraw.CurrentY
                        '��ʼ��ӡͼƬ
                        If mPrintBegingPage = 0 Or mPrintEndPage = 0 Or _
                        (IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - mPageNumber >= mPrintBegingPage - 1 And IIf(blnPrintNO, lngEndPage, lngTmpPageNo) - _
                        mPageNumber <= mPrintEndPage) Then
                            objDraw.PaintPicture ObjStdPic, TmpX, y, lngStdPicWidth, lngStdPicHeight, 0, 0, objDraw.ScaleX(ObjStdPic.Width, vbHimetric, objDraw.ScaleMode), objDraw.ScaleY(ObjStdPic.Height, vbHimetric, objDraw.ScaleMode)  'lngStdPicWidth, lngStdPicHeight
                        End If
                        objDraw.CurrentY = y + lngStdPicHeight
                        y = objDraw.CurrentY
                        sngEndY = y
                    End If
                    '��д��ǩ������ǰ���ڡ���ǰʱ�䡢����
                Case !Ԫ������ = -1 Or !Ԫ������ = -2 Or !Ԫ������ = -3 Or !Ԫ������ = -4
                    '���뷽ʽΪ2��3�ı����������ӡ
'                    If zlCommFun.Nvl(!����λ��, 1) <> 1 Then
                        objDraw.Font.Name = strTitleFontName
                        objDraw.Font.Size = strTitleFontSize
                        objDraw.Font.Bold = IIf(strTitleFontBold = "1", True, False)
                        objDraw.Font.Italic = IIf(strTitleFontItalic = "1", True, False)
                        SetPrinterFont objDraw.Font, Int(strTitleFontSize)
                        If !Ԫ������ = -4 Then
                            strText = zlCommFun.Nvl(!�����ı�)
                        Else
                            strText = IIf(!������ʾ = 1, zlCommFun.Nvl(!�����ı�) & "��", "")
                        End If
'                    Else
'                        strText = ""
'                    End If
                    '��������ı������ݲ����䱣�����ı������������Щ��Щ�����͵�Ԫ��Ӧ��û���ı��Σ�һ��Ӧֱ������
                    strSQL = "select * from ���˲����ı��� where ����ID=" & !ID & " order by �к�"
                    Call zlDatabase.OpenRecordset(rsNewTmp1, strSQL, "������ӡ")
                    If rsNewTmp1.RecordCount > 0 Then
                        For j = 0 To rsNewTmp1.RecordCount - 1
                            If !Ԫ������ = -1 And Not blnMultiSign Then
                                strText = strText & GetAllName(lng������¼ID)
                            Else
                                strText = strText & zlCommFun.Nvl(rsNewTmp1!����)
                            End If
                            rsNewTmp1.MoveNext
                        Next
                    End If
                    If !Ԫ������ = -4 Or (Printer.TextWidth(zlCommFun.Nvl(!�����ı�)) < lngWidth - (lngLeft + lngRight) - Printer.TextWidth("��")) Then
                        If !Ԫ������ = -4 Then
                            objDraw.CurrentY = sngEndY + H_9pt
                        End If
                        Select Case zlCommFun.Nvl(!����λ��, 1)
                        Case 1  '��
                            TmpX = lngLeft
                        Case 2  '��
                            TmpX = lngLeft + (lngWidth - (lngLeft + lngRight) - Printer.TextWidth(strText)) / 2
                        Case 3  '��
                            TmpX = lngWidth - lngRight - Printer.TextWidth(strText)
                        End Select
                    Else
                        TmpX = lngLeft
                        y = objDraw.CurrentY + H_9pt  '����һ��С�������
                        objDraw.CurrentY = y
                    End If
                    Do While strText <> ""
                        '�ж��Ƿ���ҳ
                        y = objDraw.CurrentY + H_9pt / 3 + Printer.TextHeight("��")
                        tmpPrintHeight = Printer.TextHeight("��")
                        Set objDraw = Nothing
                        If blnPrintNO Then
                            Set objDraw = CheckNewPage(objOut, lngEndPage, y, tmpPrintHeight)
                        Else
                            lngTmpPageNo = 0
                            Set objDraw = CheckNewPage(objOut, lngTmpPageNo, y, tmpPrintHeight)
                        End If
                        If objDraw Is Nothing Then Exit Function
                        
                        objDraw.CurrentY = y  '+ H_9pt
                        sngEndY = y
                        strText = PrintLineS(objDraw, strText, TmpX, lngRight, IIf(blnPrintNO, lngEndPage, lngTmpPageNo))
                        TmpX = lngLeft
                    Loop
                    sngEndY = objDraw.CurrentY
                End Select
                .MoveNext
            Next
        End If
    End With
    If blnPrintNO Then
        Set objDraw = NewPrintPage(objOut, lngEndPage, False)
    Else
        lngTmpPageNo = 0
        Set objDraw = NewPrintPage(objOut, lngTmpPageNo, False)
    End If

    If Not blnPrint And mPrintBegingPage <> 0 Then
        tmpPage = objOut.picPage.UBound + 1 - mPrintEndPage
        If tmpPage > 0 Then
            tmpPage = objOut.picPage.UBound + 1
            For i = tmpPage To mPrintEndPage + 1 Step -1
                Unload objOut.picPage(i - 1)
            Next
        End If
        Set objDraw = objOut.picPage(objOut.picPage.UBound)
    End If
    PrintOrPreviewCase = True
    Set objDraw = Nothing:     Set rsTmp = Nothing:     Set rsNewTmp1 = Nothing:     Set rsNewTmp2 = Nothing:     Set ObjStdPic = Nothing
    Exit Function
ErrHandle:
    If Err.Number = 480 Then
        MsgBox "û���㹻���ڴ�������ڴ����Ԥ����", vbInformation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
    Call SaveErrLog
    Set objDraw = Nothing:     Set rsTmp = Nothing:     Set rsNewTmp1 = Nothing:     Set rsNewTmp2 = Nothing:     Set ObjStdPic = Nothing
End Function
Function GetAllName(CaseHistoryID As Long) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����           �������е�ҽ�����޸ļ�¼
    '����           ������¼ID
    '����           ���е�ҽ������ ��ʽ"�ϼ�/�¼�"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String
    
    
    If CaseHistoryID > 0 Then
        '����
        strTmp = " ������¼ID = " & CaseHistoryID & " Order By  �汾���"
    Else
        gstrSql = "select * from ���˲����޶���¼ where id = " & Abs(CaseHistoryID)
        zlDatabase.OpenRecordset rsTmp, gstrSql, "ҽ��ǩ��"
        '��ǰѡ���
        strTmp = " ������¼ID = " & rsTmp("������¼ID") & "And �汾��� <= " & rsTmp("�汾���") & " Order By  �汾���"
    End If
    
    On Error GoTo errH
    
    gstrSql = "select * from ���˲����޶���¼ where " & strTmp
    
    zlDatabase.OpenRecordset rsTmp, gstrSql, "ҽ��ǩ��"
    Do Until rsTmp.EOF
        If GetAllName = "" Then
            GetAllName = rsTmp("��д��")
        Else
            GetAllName = rsTmp("��д��") & "/" & GetAllName
        End If
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    If CaseHistoryID > 0 Then
        gstrSql = "select * from ���˲�����¼ where id = " & CaseHistoryID
        zlDatabase.OpenRecordset rsTmp, gstrSql, "ҽ��ǩ��"
        If rsTmp.EOF <> True Then
            If GetAllName <> "" Then
                GetAllName = rsTmp("������") & "/" & GetAllName
            Else
                GetAllName = rsTmp("������")
            End If
        End If
'    Else
'        gstrSql = "select * from ���˲�����¼ where �����޶�ID = " & Abs(CaseHistoryID)
    End If
    
    Set rsTmp = Nothing
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
  
End Function

Private Sub SetPrinterFont(ByVal DevFont As StdFont, intFontSize As Integer)
    Printer.Font.Name = DevFont.Name
'    Printer.Font.Size = DevFont.Size
    Printer.Font.Size = intFontSize
    Printer.Font.Bold = DevFont.Bold
    Printer.Font.Underline = DevFont.Underline
    Printer.Font.Italic = DevFont.Italic
    Printer.Font.Strikethrough = DevFont.Strikethrough
    Printer.Font.Weight = DevFont.Weight
End Sub
