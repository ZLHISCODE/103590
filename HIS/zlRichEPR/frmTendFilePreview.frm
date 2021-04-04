VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmTendFilePreview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "preView"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form24"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1995
      LargeChange     =   10
      Left            =   3540
      Max             =   100
      SmallChange     =   2
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   285
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   90
      ScaleHeight     =   2655
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   630
      Width           =   3255
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   1455
         Left            =   570
         TabIndex        =   5
         Top             =   930
         Width           =   2265
         _cx             =   3995
         _cy             =   2566
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTendFilePreview.frx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblDownTable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������ɻ���"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label lblUpTable 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������ɻ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   270
         TabIndex        =   3
         Top             =   600
         Width           =   1125
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1380
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Line lineRight 
         X1              =   1380
         X2              =   1380
         Y1              =   360
         Y2              =   2220
      End
      Begin VB.Line lineLeft 
         X1              =   720
         X2              =   720
         Y1              =   360
         Y2              =   2220
      End
      Begin VB.Line lineBottom 
         X1              =   630
         X2              =   2790
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line lineTop 
         X1              =   630
         X2              =   2790
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "frmTendFilePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngRows As Long
Private arrFormat
Private dblTitle As Double      '�������ĸ߶�
Private dblUpTable As Double    '������ĸ߶�
Private dblDownTable As Double  '������ĸ߶�

Private mlngFile As Long                 '���˻����ļ�.ID
Private mlngFormat As Long               '��ʽID
Private mlngRows As Long
Private mstrSQL As String
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL���� As String
'���ĿSQL
Private mstrSQLActive�� As String
Private mstrSQLActive�� As String
Private mstrSQLActive�� As String
Private mstrSQLActive���� As String

Private mobjParent As Object

'�����ļ���ʽ�������
Private mintTabTiers As Integer     '��ͷ���
Private mintTagFormHour As Integer  '��ʼʱ������
Private mintTagToHour As Integer    '��ֹʱ������
Private mobjTagFont As New StdFont  '������ʽ����
Private mlngTagColor As Long        '������ʽ��ɫ
Private mstrPaperSet As String      '��ʽ
Private mstrPageHead As String      'ҳü
Private mstrPageFoot As String      'ҳ��
Private mblnChildForm As Boolean
Private mlngActiveRows As Long      '��Ч������
Private mstrSubhead As String       '���ϱ�ǩ
Private mstrTabHead As String       '��ͷ��Ԫ
Private mstrPreHead As String       '�账�����,�ı�����Ŀ�����л�󶨶����Ŀ����
Private Const mlngFixedCOL As Long = 2 '�̶��󶨵���,Ŀǰֻ���˻������ͼ�¼ID
Private mstrActivePreHead As String '�账��Ļ��Ŀ
Private mlngActiveColCount As Long '��Ч�Ļ��Ŀ����
Private mstrColWidth As String      '�п����д�
Private mstrColumns As String       '��ǰ�����ļ����ж�Ӧ����Ŀ
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
Private mrsItems As New ADODB.Recordset


Private Const conLineWide As Integer = 30        '������ռ���(��λΪ�)ռ�����߿��
Private Const conLineHigh As Integer = 30        '������ռ�߶�(��λΪ�)ռ�����߸߶�
Private Const conRatemmToTwip As Single = 56.6857142857143      '������羵ı���
Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'WinNT�Զ���ֽ�ſ���================================================================
'ע����dmFields��Long��,as Long��β����&��
Private Const DM_ORIENTATION = &H1&
Private Const DM_PAPERSIZE = &H2&
Private Const DM_PAPERLENGTH = &H4&
Private Const DM_PAPERWIDTH = &H8&
Private Const DM_COPIES = &H100&
Private Const DM_DEFAULTSOURCE = &H200&
Private Const DM_COLLATE = &H8000&
Private Const DM_FORMNAME = &H10000
'Constants for DocumentProperties() call
Private Const DM_COPY = 2
Private Const DM_OUT_BUFFER = DM_COPY
Private Const DM_PROMPT = 4
Private Const DM_IN_PROMPT = DM_PROMPT
Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = DM_MODIFY
'Constants for DocumentProperties() return
Private Const IDOK = 1
Private Const IDCANCEL = 2
'Constants for DEVMODE
Private Const CCHFORMNAME = 32
Private Const CCHDEVICENAME = 32

Private Type DEVMODE
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
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Private Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Function zlGetPrinterSet() As Boolean
    '------------------------------------------------
    '���ܣ���ȡ��ϵͳע���Ĵ�ӡȱʡ����
    '------------------------------------------------
    Dim iCount As Long
    Dim strDeviceName As String
    Dim intPaperSize As Integer
    Dim intPaperBin As Integer
    Dim intOrientation As Long
    
    If Printers.Count = 0 Then
        zlGetPrinterSet = False
        Exit Function
    End If
    
    strDeviceName = GetSetting("ZLSOFT", "����ģ��\" & "zl9PrintMode" & "\Default", "DeviceName", Printer.DeviceName)
    If Printer.DeviceName <> strDeviceName Then
        For iCount = 0 To Printers.Count - 1
            If Printers(iCount).DeviceName = strDeviceName Then
                Set Printer = Printers(iCount)
                Exit For
            End If
        Next
    End If
    
    Err = 0
    On Error Resume Next
    Printer.PaperBin = GetSetting("ZLSOFT", "����ģ��\" & "zl9PrintMode" & "\Default", "PaperBin", Printer.PaperBin)
    Printer.Orientation = arrFormat(1)
    
    intPaperSize = arrFormat(0)
    If intPaperSize = 256 Then
        Dim lngWidth As Long
        Dim lngHeight As Long
        
        lngWidth = arrFormat(3)
        lngHeight = arrFormat(2)
        
        Call SetCustonPager(lngWidth, lngHeight)
    Else
        Printer.PaperSize = intPaperSize
    End If

    zlGetPrinterSet = True
End Function

Private Function SetCustonPager(ByVal lngWidth As Long, ByVal lngHeight As Long) As Integer
'���ܣ��������Զ���ֽ��
'�����������Ϊ��λ
    If IsWindowsNT Then
        '��Ȼ����ʹ�����Ч�����ܸı�PaperSize������ֵ
        Printer.Width = lngWidth
        Printer.Height = lngHeight
        SetCustonPager = SetNTPrinterPaper(Me.hWnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    Else
        'Windows98ϵ�л�����ͨ����������
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    End If
End Function

Private Function IsWindowsNT() As Boolean
'���ܣ��Ƿ�WindowNT����ϵͳ
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Private Function IsWindows95() As Boolean
'��    �ܣ��ж��Ƿ���Windows95�¹���
'��    ������
'��    �أ��Ƿ���True
    Const dwMask95 = &H1&
    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function

Private Function GetWinPlatform() As Long
'��    �ܣ����ص�ǰ��ϵͳ�汾����
'��    ������
'��    �أ�
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

Private Function SetNTPrinterPaper(ByVal lngHwnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
    ByVal intOrient As Integer, ByVal intCopys As Integer, Optional ByVal blnPrompt As Boolean) As Boolean
'���ܣ�NT�����У����ô�ӡ�����Զ���ֽ�ųߴ�
'������lngWidth��lngHeight=mm(����)
'     intOrient=1-����,2-����
'     intCopys=��ӡ����(�����ӡ��֧��,1-9999,��֧��ʱ�������,Ҳ��Ӱ����������)
'˵��������Width,Height�⣬����ͨ�����������õ����Բ�ֱ�ӷ�ӳ��Printer�ϣ�
'      (ȡDevModeҲ��ӳ������������Ҫ��GetJob���ܻ�ȡ����Ĵ�ӡ�ĵ�����)
    Dim vDevMode As DEVMODE
    Dim arrDevMode() As Byte
    Dim lngSize As Long
    
    Dim lngPrtDC As Long
    Dim lngHandle As Long
    Dim strPrtName As String
    
    lngPrtDC = Printer.hDC
    strPrtName = Printer.DeviceName
    
    If OpenPrinter(strPrtName, lngHandle, 0&) Then
        'Retrieve the size of the DEVMODE:fMode=0
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, 0&, 0&, 0&)
        'Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
    
        'Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        'Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        '���ô�ӡ�ĵ�����
        vDevMode.dmOrientation = intOrient
        vDevMode.dmPaperSize = 256
        vDevMode.dmPaperWidth = intWidth * 10 'in tenths of a millimeter
        vDevMode.dmPaperLength = intHeight * 10 'in tenths of a millimeter
        vDevMode.dmCopies = intCopys
        'vDevMode.dmCollate = 0& '�߼���ӡ����(��ȡ��ʱ,Copiesֻ֧��1;����֪��ôȡ����)
        vDevMode.dmFields = DM_ORIENTATION Or DM_PAPERSIZE Or DM_PAPERLENGTH Or DM_PAPERWIDTH Or DM_COPIES 'Or DM_COLLATE
        
        'Copy your changes back, then update DEVMODE.
        Call CopyMemory(arrDevMode(1), vDevMode, Len(vDevMode))
        If blnPrompt Then
            lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_IN_PROMPT Or DM_OUT_BUFFER)
        Else
            lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        End If
        If lngSize = IDOK Then SetNTPrinterPaper = True
        'Reset the DEVMODE for the DC.
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))
        If lngSize = 0 Then SetNTPrinterPaper = False
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function


Private Sub Form_Load()
    Dim lngFixRows As Long                          '�̶�����
    Dim dblRowHeight As Double                      '�и�
    Dim lngParent As Long
    Dim strUpText As String
    Dim lngHeight As Long, lngWidth As Long         '��Ч�߶ȣ����
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim lngOffsetLeft As Long, lngOffsetTop As Long
    Dim rsTemp As New ADODB.Recordset
    Dim arrHead() As String   '��ͷ����
    Dim arrData() As String, arrColWith() As String
    Dim lngMutilRow1 As Long, lngMutilRow2 As Long, lngMutilRow3 As Long
    Dim i As Integer
    Dim lngFixRowsheight As Long
    On Error GoTo errHand
    'arrFormat(ֽ��|ֽ��|��|��|�ϱ߾�|�±߾�|��߾�|�ұ߾�|�и�|�̶�����|������������|�����������С|�����ı�|������������|�����������С|�������ı�|��ͷ��Ŀ����)
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    
    '����ҳ���ʽ
    Call zlGetPrinterSet
    
    '��ȡ��ӡ����ǰ״̬
    picDraw.Height = Printer.Height
    picDraw.Width = Printer.Width
    lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
    picDraw.ScaleHeight = Printer.Height - lngOffsetTop * 2
    picDraw.ScaleWidth = Printer.Width - lngOffsetLeft * 2
    'ҳ�߾�
    lngTop = arrFormat(4)
    lngBottom = arrFormat(5)
    lngLeft = arrFormat(6)
    lngRight = arrFormat(7)
    'ʵ����Ч�߶ȣ����
    lngHeight = picDraw.ScaleHeight - lngTop - lngBottom
    lngWidth = picDraw.ScaleWidth - lngLeft - lngRight
    
    '��,�±߾�(lngTop , lngBottom)
    '��,�ұ߾�(lngLeft , lngRight)
    lineTop.X1 = 0
    lineTop.X2 = picDraw.ScaleWidth
    lineTop.Y1 = lngTop
    lineTop.Y2 = lngTop
    lineBottom.X1 = 0
    lineBottom.X2 = picDraw.ScaleWidth
    lineBottom.Y1 = picDraw.ScaleHeight - lngBottom
    lineBottom.Y2 = lineBottom.Y1
    
    lineLeft.X1 = lngLeft
    lineLeft.X2 = lngLeft
    lineLeft.Y1 = 0
    lineLeft.Y2 = picDraw.ScaleHeight
    lineRight.X1 = picDraw.ScaleWidth - lngRight
    lineRight.X2 = lineRight.X1
    lineRight.Y1 = 0
    lineRight.Y2 = picDraw.ScaleHeight
    
    '1�����������ϱ߾࿪ʼ
    dblRowHeight = arrFormat(8)
    VsfData.RowHeightMin = dblRowHeight
    '�̶�����,���ݱ�ͷ���ݼ���
    '98992,����,2016-12-19
    If UBound(arrFormat) > 15 Then
        lngFixRowsheight = GetFixRowsHeight(arrFormat(16), arrFormat(9))
    Else
        lngFixRowsheight = arrFormat(9)
    End If
    '����������������
    lblTitle.FontName = arrFormat(10)
    lblTitle.FontSize = arrFormat(11)
    lblTitle.Caption = arrFormat(12)
    '���������������
    lblUpTable.FontName = arrFormat(13)
    lblUpTable.FontSize = arrFormat(14)
    
    '���ñ���������
    picDraw.FontName = lblTitle.FontName
    picDraw.FontSize = lblTitle.FontSize
    lblTitle.Left = lngLeft
    lblTitle.Top = lngTop + 30
    lblTitle.Width = lngWidth
    lblTitle.Height = picDraw.TextHeight("a")
    
    '2�����ϱ�ǩ�ӱ������¿�ʼ
    strUpText = arrFormat(11)
    If strUpText <> "" Then
        lblUpTable.Caption = strUpText
        lblUpTable.AutoSize = True
    End If
    '���ñ���������
    picDraw.FontName = lblUpTable.FontName
    picDraw.FontSize = lblUpTable.FontSize
    lblUpTable.Left = lngLeft
    lblUpTable.Top = lblTitle.Top + lblTitle.Height + 30
    lblUpTable.Width = picDraw.ScaleWidth
    
    '3�����ñ��
    lngHeight = lngHeight - lblUpTable.Height - lblTitle.Height
    VsfData.Top = lblUpTable.Top + lblUpTable.Height + 30
    VsfData.Left = lngLeft
    VsfData.Width = lngWidth
    lngHeight = lngHeight + lngTop - VsfData.Top - lngFixRowsheight
    VsfData.Height = lngHeight
    lngRows = CLng(lngHeight \ dblRowHeight)
    VsfData.Rows = lngFixRows + lngRows
    VsfData.FixedRows = lngFixRows
    VsfData.RowHeightMin = dblRowHeight
    
    Call VScroll1_Change
    
    If mrsItems.State = 0 Then
        '���ִ��ڵ����л����¼��Ŀ
        gstrSQL = " Select ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ" & _
                  " From �����¼��Ŀ B" & _
                  " Where B.Ӧ�÷�ʽ<>0 " & _
                  " Order by ��Ŀ���"
        Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjParent Is Nothing Then
        Set mobjParent = Nothing
    End If
End Sub

Private Sub VScroll1_Change()
    picDraw.Top = -1 * VScroll1.Value * (picDraw.Height - Me.Height) / 100
End Sub

Public Function ShowMe(ByVal objParent As Object, ByVal strInput As String) As Long
    '��ȡ�����¼���ĸ�ʽ
    lngRows = 0
    arrFormat = Split(strInput, "|")
'    Me.Show 1, objParent   '�˶�������ȷ��ʱ����Ҫ�ɼ�����
    Unload frmTendFilePreview
    Load frmTendFilePreview
    Unload frmTendFilePreview
    ShowMe = lngRows
End Function

Public Function AnaliseData(ByVal objParent As Object, ByVal lngFormat As Long, ByVal strInput As String) As Boolean
    mlngFormat = lngFormat
    arrFormat = Split(strInput, "|")
    Set mobjParent = objParent
    
    Unload frmTendFilePreview
    Load frmTendFilePreview
    
    If Not ReadStruDef Then
        'û����Ҫ��������,���ֱ�ӷ��ؽ����ɹ�,Ӧ�ò������������
        AnaliseData = (mstrPreHead = "")
        Exit Function
    End If
    If Not ReadData Then Exit Function
    Unload frmTendFilePreview
    AnaliseData = True
End Function

Private Function ReadData() As Boolean
    Dim strCaption As String
    Dim rsPati As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean
    
    On Error GoTo errHand
    
    strCaption = mobjParent.Caption
    '��ȡ����ʹ�øü�¼�ļ��Ĳ����б�(�Ѿ���ӡ���Ļ����ļ�����������������)
    gstrSQL = _
        " Select Id, ����id, ����id, ��ҳid, Ӥ��" & vbNewLine & _
        " From ���˻����ļ� a" & vbNewLine & _
        " Where ��ʽid = [1] And �鵵�� Is Null And Not Exists" & vbNewLine & _
        " (Select 1 From ���˻����ӡ b Where b.�ļ�id = a.Id And b.��ӡ�� Is Not Null)"
    
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡʹ�øû����ļ��Ĳ����б�", mlngFormat)
    
    gcnOracle.BeginTrans
    blnTrans = True
    Do While Not rsPati.EOF
        mobjParent.Caption = strCaption & Space(2) & "һ����" & rsPati.RecordCount & "�ݻ����ļ������ڴ���" & rsPati.AbsolutePosition
        
        '���Ŀ����(�����з������)
        Call PreActiveCOL(rsPati!ID)
        'װ������
        mlngFile = rsPati!ID
        Call SQLCombination
        gstrSQL = mstrSQL
'        If mlngFile = 86 Then Stop
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", CLng(rsPati!ID), CLng(rsPati!����ID), CLng(rsPati!��ҳID), CLng(rsPati!Ӥ��))
        '�����ݲ����û����¼���ĸ�ʽ,ͬʱʵ��һ�����ݷ�����ʾ�Ĺ���
        Call PreTendFormat(rsTemp)
        '����ÿ������
        If Not ParseData Then
            mobjParent.Caption = strCaption
            gcnOracle.RollbackTrans
            Exit Function
        End If
        
        rsPati.MoveNext
    Loop
    mobjParent.Caption = strCaption
    gcnOracle.CommitTrans
    blnTrans = False
    
    ReadData = True
    Exit Function
errHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    mobjParent.Caption = strCaption
End Function

Private Function ParseData() As Boolean
    Dim arrCol, arrData, arrMutilRow
    Dim strTime As String
    Dim lngMutilRow As Long, lngRecord As Long
    Dim lngRow As Long, lngCount As Long
    Dim lngCol As Long, lngMAX As Long
    Dim blnSave As Boolean, i As Long
    Dim strSQLData() As String
    ReDim Preserve strSQLData(1 To 1)
    
    On Error GoTo errHand
    'ѭ����������������(һ�а󶨶����Ŀ,������ĿΪ�ı���)
    
    arrCol = Split(mstrPreHead & mstrActivePreHead, ",")
    lngMAX = UBound(arrCol)
    lngCount = VsfData.Rows - 1
    
    gstrSQL = "ZL_���˻����ӡ_DELETE(" & mlngFormat & "," & mlngFile & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "�����ǰ�ļ��Ĵ�ӡ����"
    
    blnSave = False: arrMutilRow = Array()
    For lngRow = 1 To lngCount
        lngMutilRow = 0
        lngRecord = Val(VsfData.TextMatrix(lngRow, VsfData.Cols - 1))
        If lngRecord <> 0 Then
            blnSave = True
            strTime = Format(VsfData.TextMatrix(lngRow, 1), "YYYY-MM-DD HH:mm:ss")
            
            '������ܴ���:����������ʱ�����������ϸ������
            If Val(VsfData.TextMatrix(lngRow, VsfData.Cols - 2)) < 0 Then
                If lngRow + 1 <= lngCount Then
                    If Val(VsfData.TextMatrix(lngRow + 1, VsfData.Cols - 1)) > 0 And Val(VsfData.TextMatrix(lngRow + 1, VsfData.Cols - 2)) < 0 And _
                        strTime = Format(VsfData.TextMatrix(lngRow + 1, 1), "YYYY-MM-DD HH:mm:ss") Then
                        blnSave = False
                    End If
                End If
            End If
                
            For lngCol = 0 To lngMAX
                If VsfData.TextMatrix(lngRow, arrCol(lngCol)) <> "" Then
                    '׼����ֵ
                    With txtLength
                        .Width = VsfData.ColWidth(arrCol(lngCol))
                        .Text = VsfData.TextMatrix(lngRow, arrCol(lngCol))
                        .FontName = VsfData.FontName
                        .FontSize = VsfData.FontSize
                        .FontBold = VsfData.CellFontBold
                        .FontItalic = VsfData.CellFontItalic
                    End With
                    arrData = GetData(txtLength.Text)
                    If UBound(arrData) > lngMutilRow Then
                        lngMutilRow = UBound(arrData)
                        If Trim(arrData(lngMutilRow)) = "" Then lngMutilRow = lngMutilRow - 1
                    End If
                End If
            Next
            If lngMutilRow < 0 Then lngMutilRow = 0
            ReDim Preserve arrMutilRow(UBound(arrMutilRow) + 1)
            arrMutilRow(UBound(arrMutilRow)) = lngMutilRow + 1
            If blnSave = True Then
                '----�˴���Ҫ���������ܵ�����
                lngMutilRow = 0
                '���������ϸ����������
                For i = 1 To UBound(arrMutilRow)
                    lngMutilRow = lngMutilRow + arrMutilRow(i)
                Next i
                '��������������������ڷ�����ϸ��������+1(1ΪĬ�ϵ���������),��������������Ϊ׼,��������ϸ������+1Ϊ׼
                If lngMutilRow + 1 > Val(arrMutilRow(0)) Then
                    lngMutilRow = lngMutilRow + 1
                Else
                    lngMutilRow = Val(arrMutilRow(0))
                End If
                arrMutilRow = Array()
                
                gstrSQL = "ZL_���˻����ӡ_UPDATE(" & mlngFile & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss')" & "," & lngMutilRow & ")"
                strSQLData(ReDimArray(strSQLData)) = gstrSQL
            End If
        End If
    Next
    
    'ִ�й���
    For i = 1 To UBound(strSQLData)
        If strSQLData(i) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLData(i), "������ӡ��������")
        End If
    Next i
    
    ParseData = True
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim blnTag As Boolean
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    On Error GoTo errHand
    
    '���û����¼���ĸ�ʽ
    With VsfData
        .FixedRows = 3
        .Clear
        Set .DataSource = rsTemp
        
        '��ͷ��д
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '��ɻ��Ŀ�еĴ���(Ҫ����Ļ��Ŀ)�����Ŀ�ֶ���ȡ��SQLĬ�ϰ�����󣬴˴�ͬ�������������̶�����е�ǩ��,�����ܶԺ���Ĵ������Ӱ��
        mstrActivePreHead = ""
        If mlngActiveColCount > 0 Then
            '�ƶ����Ŀ�е��̶��е�ǩ��
            For lngCol = 1 To mlngFixedCOL
                .ColPosition(.Cols - mlngActiveColCount - mlngFixedCOL) = .Cols - 1
                .ColHidden(.Cols - 1) = True
            Next
            For lngCol = .Cols - mlngActiveColCount - mlngFixedCOL To .Cols - mlngFixedCOL - 1
                mstrActivePreHead = mstrActivePreHead & "," & lngCol
                .ColHidden(lngCol) = True
            Next
        Else
            For lngCol = 1 To mlngFixedCOL
                .ColHidden(.Cols - lngCol) = True
            Next
        End If
        
        '������ͷ
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + 1) = strCell
        Next
        
        '�п�����
        Dim blnAlign As Boolean
        aryItem = Split(mstrColWidth, ",")
        For lngCount = 2 To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(0))
                If InStr(1, aryItem(lngCount - 2), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(1))
                End If
            End If
        Next
        
        '�̶��и�ʽΪ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '�ٰ��кϲ�
        For lngCount = 0 To VsfData.Cols - 1
            VsfData.MergeCol(lngCount) = True
        Next
        .AutoSize 0, .Cols - 1
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .ROWHEIGHT(lngCount) < .RowHeightMin Then .ROWHEIGHT(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
    End With
    Exit Sub
errHand:
    MsgBox Err.Description
End Sub

Private Function ReadStruDef() As Boolean
    Dim arrCol
    Dim intCol As Integer, intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '��ȡ�����ļ���ʽ����
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ���ʽ����", mlngFormat)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "��ͷ����": mintTabTiers = Val("" & !�����ı�)
            Case "������":  VsfData.Cols = Val("" & !�����ı�)
            Case "��С�и�": VsfData.RowHeightMin = Val("" & !�����ı�)
            Case "�ı�����"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set lblUpTable.Font = VsfData.Font
                Set Font = lblUpTable.Font
                
            Case "�ı���ɫ": VsfData.ForeColor = Val("" & !�����ı�)
            Case "�����ɫ": VsfData.GridColor = Val("" & !�����ı�): VsfData.GridColorFixed = VsfData.GridColor
            
            Case "�����ı�": lblTitle.Caption = "" & !�����ı�
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False
            
            Case "��ʼʱ��": mintTagFormHour = Val("" & !�����ı�)
            Case "��ֹʱ��": mintTagToHour = Val("" & !�����ı�)
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "������ɫ": mlngTagColor = Val("" & !�����ı�)
            Case "��Ч������": mlngActiveRows = Val(!�����ı�)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select ��ʽ, ҳü, ҳ��,���� From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", mlngFormat)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!��ʽ: mstrPageHead = "" & rsTemp!ҳü: mstrPageFoot = "" & rsTemp!ҳ��
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlngFormat)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ͷ��Ԫ����", mlngFormat)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !�����д� - 1 & "," & !������� & "," & !�����ı�
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '��ѯ�����֯
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql�� As String, str��ʽ As String
    Dim bln���� As Boolean, blnʱ�� As Boolean, bln��ʿ As Boolean
    Dim blnǩ���� As Boolean, blnǩ��ʱ�� As Boolean, blnǩ������ As Boolean
    Dim lngColumn As Long
    
    gstrSQL = "Select d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", mlngFormat)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = ""
        mstrSQL�� = "": mstrSQL�� = "": strSql�� = "": mstrSQL�� = "": mstrSQL���� = ""
        bln���� = False: blnʱ�� = False: bln��ʿ = False
        blnǩ���� = False: blnǩ��ʱ�� = False: blnǩ������ = False
        Do While Not .EOF
            
            If lngColumn <> !������� Then
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", ";1;" & str��ʽ) & "|" & !������� & ";" & !Ҫ������
                mstrColWidth = mstrColWidth & "," & !��������
                str��ʽ = ""
                If !Ҫ������ <> "" Then
                    str��ʽ = "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                    If strSql�� <> "" Then
                        mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
                    End If
                Else
                    If strSql�� <> "" Then
                        mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql�� = ""
                lngColumn = !�������
            Else
                mstrColumns = mstrColumns & "'" & !Ҫ������
                str��ʽ = str��ʽ & "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
            End If
            
            Select Case !Ҫ������
            Case "����"
                bln���� = True
                mstrSQL�� = mstrSQL�� & ",����"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case "ʱ��"
                blnʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ʱ��"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ����"
                blnǩ���� = True
                mstrSQL�� = mstrSQL�� & ",ǩ����"
                mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ��ʱ��"
                blnǩ��ʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ǩ��ʱ��"
                mstrSQL�� = mstrSQL�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����,12,5)) As ǩ��ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ������"
                blnǩ������ = True
                mstrSQL�� = mstrSQL�� & ",ǩ������"
                mstrSQL�� = mstrSQL�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����, 1,11)) As ǩ������"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "��ʿ"
                bln��ʿ = True
                mstrSQL�� = mstrSQL�� & ",��ʿ"
                mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case Else
                If !Ҫ������ <> "" Then
                    mstrSQL�� = mstrSQL�� & ",Max(""" & !Ҫ������ & """) As """ & !Ҫ������ & """"
                    mstrSQL���� = mstrSQL���� & " Or """ & !Ҫ������ & """ Is Not Null"
                    strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                    
                    If Trim("" & !�����ı�) = "" And Trim("" & !Ҫ�ص�λ) = "" Then
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & !Ҫ������ & """"
                    Else
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,Null,'" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')), '') As """ & !Ҫ������ & """"
                    End If
                End If
            End Select
            .MoveNext
        Loop
        
        mstrColWidth = Mid(mstrColWidth, 2)
        '�������һ�еĸ�ʽ
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", ";1;" & str��ʽ) '& "|" & !������� & ";" & !Ҫ������
        mstrColumns = Mid(mstrColumns, 2)     '��ʽ��:�к�;��Ŀ����1,��Ŀ����2|�к�...,ʵ��;1;����|2;����|3...
        If Mid(strSql��, 3) <> "" Then
            mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
        Else
            mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If mstrSQL���� <> "" Then mstrSQL���� = "(" & Mid(mstrSQL����, 5) & ")"
        
        '���û�г������ڣ�ʱ�䣬��ʿ�����ڲ���Ҫ���䣬�Ա�֤�в�����������
        If bln���� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
        If blnʱ�� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
        If bln��ʿ = False Then mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
        
        If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
        If blnǩ������ = False Then mstrSQL�� = mstrSQL�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����,1,11)) As ǩ������"
        If blnǩ��ʱ�� = False Then mstrSQL�� = mstrSQL�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����,12,5)) As ǩ��ʱ��"
        
        If Mid(mstrSQL��, 2) = "" Then
            MsgBox "�Բ�����û�ж��嵱ǰ��������ʾ����Ϣ�����ڲ����ļ������ж��壡"
            Exit Function
        End If
        
        '�����ڲ��������ӹ̶���
        mstrSQL�� = mstrSQL�� & ",MAX(�������) AS �������,MAX(��¼ID) AS ��¼ID"
        mstrSQL�� = mstrSQL�� & ",NVL(L.�������,0) AS �������,C.��¼ID"
        mstrSQL�� = mstrSQL�� & ",�������,��¼ID"
        
        '������Щ�е�������Ҫ���д�ӡ��������
        Dim arrData
        Dim strtodo As String
        Dim intto As Integer, intDo As Integer
        mstrPreHead = ""
        arrCol = Split(mstrColumns, "|")
        intCount = UBound(arrCol)
        For intCol = 0 To intCount
            'If UBound(Split(Split(arrCol(intCol), ";")(3), "]}{[")) > 0 Then
            If UBound(Split(Split(arrCol(intCol), ";")(1), "'")) > 0 Then
                'ֻҪ��һ����������������Ϊ�ı��ʹ���
                
'                strtodo = Split(arrCol(intCol), ";")(3)
'                strtodo = Replace(strtodo, "]}{[", "||")
'                strtodo = Replace(Replace(strtodo, "{[", ""), "]}", "")
'                arrData = Split(strtodo, "||")
                strtodo = Split(arrCol(intCol), ";")(1)
                arrData = Split(strtodo, "'")
                intDo = UBound(arrData)
                For intto = 0 To intDo
                    mrsItems.Filter = "��Ŀ����='" & arrData(intto) & "'"
                    If mrsItems.RecordCount <> 0 Then
                        '����û�������Ŀʱ�������ó��ı���,��ô������20�����ϵ���Ŀ�ż��,�û����ý������͵����ó������Ͳ���ȷ
'                        If mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ��ʾ = 0 And mrsItems!��Ŀ���� >= 10 Then
                            mstrPreHead = mstrPreHead & "," & Val(Split(arrCol(intCol), ";")(0)) + 1    '�����й̶����У�������Ŵ�0��ʼ�����+1
                            Exit For
'                        End If
                    End If
                Next
            Else
                '����Ƿ�Ϊ�ı���
                'mrsItems.Filter = "��Ŀ����='" & Replace(Replace(Split(arrCol(intCol), ";")(3), "{[", ""), "]}", "") & "'"
                mrsItems.Filter = "��Ŀ����='" & Split(arrCol(intCol), ";")(1) & "'"
                If mrsItems.RecordCount <> 0 Then
                    '����û�������Ŀʱ�������ó��ı���,��ô������20�����ϵ���Ŀ�ż��,�û����ý������͵����ó������Ͳ���ȷ
'                    If mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ��ʾ = 0 And mrsItems!��Ŀ���� >= 10 Then
                        mstrPreHead = mstrPreHead & "," & Val(Split(arrCol(intCol), ";")(0)) + 1    '�����й̶����У�������Ŵ�0��ʼ�����+1
'                    End If
                End If
            End If
        Next
        
        mrsItems.Filter = 0
        If mstrPreHead = "" Then Exit Function
        mstrPreHead = Mid(mstrPreHead, 2)
    End With
    
    ReadStruDef = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SQLCombination()
    mstrSQL = "Select ����,����ʱ��," & Mid(mstrSQL��, 12) & mstrSQLActive�� & vbCrLf & _
                " From (Select ��¼���,ʱ�� as ����,����ʱ��," & Mid(mstrSQL��, 2) & mstrSQLActive�� & vbCrLf & _
                "        From (Select NVL(c.��¼���,0) ��¼���,l.����ʱ��," & Mid(mstrSQL��, 2) & mstrSQLActive�� & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻�����ϸ a,���˻����ļ� f " & vbCrLf & _
                "               Where l.Id = c.��¼id And l.�ļ�ID+0=f.ID " & _
                "               And a.��¼id(+)=l.ID And a.��¼����(+)=5 And Nvl(a.��ֹ�汾,0)=0 And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] )" & vbCrLf & _
                IIf(mstrSQL���� <> "", "Where " & mstrSQL���� & mstrSQLActive����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ������,ǩ��ʱ��" & _
                                "       Order By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ������,ǩ��ʱ��)"
End Sub

Private Sub PreActiveCOL(ByVal lngFileID As Long)
'���ܣ���ȡָ���ļ��󶨵Ļ��Ŀ,���󶨵�������ȡSQL��
    Dim rsTemp As New ADODB.Recordset
    Dim strCOLActive As String, StrKey As String, strCOL As String
    Dim i As Integer, j As Integer, blnAdd As Boolean, intMax As Integer, intCol As Integer
    Dim arrCol, arrAc
    Dim strColFormat As String, strCOLNames As String, strCOLPart As String, strCOLCOND As String, strCOLDEF As String, strCOLMID As String, strCOLIN As String
    On Error GoTo errHand
    
    mlngActiveColCount = 0
    mstrActivePreHead = ""
    mstrSQLActive�� = ""
    mstrSQLActive���� = ""
    mstrSQLActive�� = ""
    mstrSQLActive�� = ""
    '1����ȡ�󶨵Ļ��Ŀ��Ϣ
    gstrSQL = " Select   A.�к�,A.ҳ��,A.��ͷ����,A.���,A.��Ŀ���,A.��λ From ���˻�����Ŀ A " & _
              " Where A.�ļ�ID=[1]" & _
              " Order by A.ҳ��,A.�к�,A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������Զ���Ļ��Ŀ", lngFileID)
    If rsTemp.RecordCount <> 0 Then
        Do While Not rsTemp.EOF
            If StrKey <> rsTemp!ҳ�� & "_" & rsTemp!�к� Then
                StrKey = rsTemp!ҳ�� & "_" & rsTemp!�к�
                strCOLActive = strCOLActive & "||" & StrKey & "|" & rsTemp!��Ŀ��� & "," & NVL(rsTemp!��λ)
            Else
                strCOLActive = strCOLActive & ";" & rsTemp!��Ŀ��� & "," & NVL(rsTemp!��λ)
            End If
            rsTemp.MoveNext
        Loop
    End If
    If strCOLActive <> "" Then strCOLActive = Mid(strCOLActive, 3)
    If strCOLActive = "" Then Exit Sub
    '2:���Ŀȡ�ش���
    arrCol = Split(strCOLActive, "||")
    arrAc = Array()
    For i = 0 To UBound(arrCol)
        blnAdd = True
        StrKey = CStr(arrCol(i))
        StrKey = Mid(StrKey, InStr(1, StrKey, "|") + 1)
        For j = 0 To UBound(arrAc)
            If CStr(arrAc(j)) = StrKey Then
                blnAdd = False
                Exit For
            End If
        Next j
        If blnAdd = True Then
            ReDim Preserve arrAc(UBound(arrAc) + 1)
            arrAc(UBound(arrAc)) = StrKey
        End If
    Next i
    '3:��ʼ���л��Ŀ������ȡSQL��װ
    For i = 0 To UBound(arrAc)
        intCol = i + 1
        arrCol = Split(arrAc(i), ";") 'ÿһ�а󶨵���Ŀ
        intMax = UBound(arrCol)
        '�����б�ʾ(ÿ������������Ŀ)
        strCOLPart = "": strCOLNames = "": strColFormat = "": strCOLCOND = "": strCOLMID = "": strCOLIN = "": strCOLDEF = ""
        For j = 0 To intMax
            strCOLPart = Split(arrCol(j), ",")(1)
            mrsItems.Filter = "��Ŀ���=" & Val(Split(arrCol(j), ",")(0))
            If mrsItems.RecordCount > 0 Then
                strCOLNames = strCOLNames & "," & mrsItems!��Ŀ����
                strCOLCOND = strCOLCOND & " OR """ & strCOLPart & mrsItems!��Ŀ���� & """ IS NOT NULL"
                strCOLMID = strCOLMID & ",Max(""" & strCOLPart & mrsItems!��Ŀ���� & """) As """ & strCOLPart & mrsItems!��Ŀ���� & """"
                If j = 0 Then
                    strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.���²�λ||") & "c.��Ŀ����, '" & strCOLPart & mrsItems!��Ŀ���� & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & strCOLPart & mrsItems!��Ŀ���� & """"
                Else
                    strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.���²�λ||") & "c.��Ŀ����, '" & strCOLPart & mrsItems!��Ŀ���� & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'/','/'||c.��¼����||'')), '') As """ & strCOLPart & mrsItems!��Ŀ���� & """"
                End If
                If j = 0 Then
                    If intMax = 0 Then
                        strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """ AS A" & Format(intCol, "00")
                    Else
                        strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """||"
                    End If
                Else
                    strCOLDEF = strCOLDEF & "NVL(""" & strCOLPart & mrsItems!��Ŀ���� & """,'/')"
                    If j = intMax Then
                        strCOLDEF = "Decode(" & strCOLDEF & ",'" & String(intMax, "/") & "',''," & strCOLDEF & ") As A" & Format(intCol, "00")
                    End If
                End If
            End If
        Next j
        '�����Ŀ�м���Ҫ���������
        If strCOLNames <> "" Then
            mlngActiveColCount = mlngActiveColCount + 1
        End If
        '��װ���ĿSQL
        mstrSQLActive�� = mstrSQLActive�� & "," & strCOLDEF
        mstrSQLActive���� = mstrSQLActive���� & strCOLCOND
        mstrSQLActive�� = mstrSQLActive�� & strCOLMID
        mstrSQLActive�� = mstrSQLActive�� & strCOLIN
    Next i
    mrsItems.Filter = ""
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼��������,û�±�
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long, lngLen As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.hWnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        lngLen = SendMessage(txtLength.hWnd, EM_GETLINE, lngRow - 1, strLine(0))
        Call ClearArray(strLine, lngLen)
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte, Optional ByVal lngPos As Long = 0)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = lngPos To intMax
        strLine(intDo) = 0
        If lngPos > 0 Then Exit Sub '��Ϊ��,��ʾ�������ַ���������
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Private Function GetFixRowsHeight(ByVal strTabHead As String, ByVal lngFixRow As Long) As Long
    Dim aryItem() As String
    Dim arrTemp() As String
    Dim strCell As String, StrText As String
    Dim lngCellWith As Long
    Dim lngRow As Long, lngCol As Long
    Dim lngCount As Long

        aryItem = Split(strTabHead, "'")
        VsfData.Cols = (UBound(aryItem) + 1) / lngFixRow
        VsfData.FixedRows = 3
        With VsfData
            .MergeCells = flexMergeRestrictRows
            .MergeCellsFixed = flexMergeFree
            For lngCount = 0 To UBound(aryItem)
                strCell = aryItem(lngCount)
                lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
                lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
                StrText = Left(strCell, InStr(1, strCell, ",") - 1)
                lngCellWith = Mid(strCell, InStr(1, strCell, ",") + 1)
                .TextMatrix(lngRow, lngCol) = StrText
                '�п�����
                    
                .ColWidth(lngCol) = lngCellWith
                
            Next
            '�ٰ��кϲ�
            For lngCount = 0 To .Cols - 1
                .MergeCol(lngCount) = True
            Next
            .MergeRow(-1) = True
            .AutoResize = True
            .WordWrap = True
            .AutoSizeMode = flexAutoSizeRowHeight
            .AutoSize 0, .Cols - 1
            .AutoResize = False
            '�����и�
            For lngCount = 0 To .Rows - 1
                If .ROWHEIGHT(lngCount) < .RowHeightMin Then .ROWHEIGHT(lngCount) = .RowHeightMin
            Next
        End With
        GetFixRowsHeight = VsfData.ROWHEIGHT(0) + VsfData.ROWHEIGHT(1) + VsfData.ROWHEIGHT(2)
End Function
