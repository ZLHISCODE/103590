VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.UserControl usrTendFileReader 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8565
   Begin MSComctlLib.ImageList imgRow 
      Left            =   6150
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileReader.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileReader.ctx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   3930
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3090
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   450
      ScaleHeight     =   3825
      ScaleWidth      =   7485
      TabIndex        =   5
      Top             =   810
      Width           =   7515
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1380
         Picture         =   "usrTendFileReader.ctx":0734
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3420
         Width           =   345
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1710
         Picture         =   "usrTendFileReader.ctx":0ABA
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3420
         Width           =   345
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2040
         Picture         =   "usrTendFileReader.ctx":0E4A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3420
         Width           =   345
      End
      Begin VB.CheckBox chkҳ�� 
         Caption         =   "��ӡҳ��"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   3480
         Width           =   1155
      End
      Begin RichTextLib.RichTextBox rtbHead 
         Height          =   1380
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2434
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"usrTendFileReader.ctx":11A3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbFoot 
         Height          =   1380
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1950
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2434
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"usrTendFileReader.ctx":1240
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   60
      ScaleHeight     =   4515
      ScaleWidth      =   8385
      TabIndex        =   1
      Top             =   510
      Width           =   8385
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   1590
         TabIndex        =   0
         Top             =   930
         Width           =   4305
         _cx             =   7594
         _cy             =   4683
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
         SelectionMode   =   1
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
         FormatString    =   $"usrTendFileReader.ctx":12DD
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
         WordWrap        =   0   'False
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
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ�㻤���¼��"
         Height          =   180
         Left            =   3450
         TabIndex        =   3
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:##"
         Height          =   180
         Left            =   390
         TabIndex        =   2
         Top             =   540
         Width           =   720
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "usrTendFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'��������:
'1.�����¼ͬһʱ��ֻ���ܴ���һ����¼
'2.�����¼�в���Ҫ�����µ����� , ��¼�����Ƿ����, �ܲ������, �����˵����ݲż�¼
'3.¼�뻤���¼����ʱ,�����¼������ݴ�����������, ����ȡ����
'4.�����¼���в���Ҫ¼�������¼�������׾����ȷ��Ҫ��¼���ڻ���ժҪ�������͵�����
'#ʵ��ԭ��:
'1.�����û��޸Ĺ�������,�����ṩ�༭״̬ҳ���л��Ĺ���,���û��޸Ĺ���ҳ���ݽ�����ҳ����,���ٳ���ʵ���Ѷ�
'2.���Ӽ�¼����¼��Щҳ��Щ��Ԫ���û��޸Ĺ�
'3.�κα༭(ճ��,�������),����Ҫ���¼���ÿ�����ݵ�ռ����

Public mblnEditable As Boolean
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream
Private mfrmParent As Object
Private mblnInit As Boolean
Private mblnShow As Boolean                 '�Ƿ���ʾ¼���
Private mintPreDays As Long
Private mstrMaxDate As String
Private mintNORule As Integer             '0-���ļ���ʽ���;1-ͳһ���

Private mlng��ǰҳ�� As Long
Private mlng��ʼҳ�� As Long
Private mint��ǰ��ʼҳ As Integer           '��ǰ�ļ�����ʼҳ(�����Ѵ�ӡ����,�Լ�Ԥ�����Ѵ�ӡҳ��ʼԤ��)
Private mint����ҳ As Integer
Private mintҳ�� As Integer
Private mlng��ǰ�ļ�ID As Long
Private mlng�ϲ��ļ�ID As Long
Private mlng��ӡҳ As Long
Private mlng��ʽID As Long
Private mlng����id As Long
Private mlng��ҳid As Long
Private mlng����ID As Long
Private mlng����ID As Long
Private mintӤ�� As Integer
Private mbln���� As Boolean                 '�Ƿ���Ҫ¼������
Private mstrPrivs As String

Private mintSymbol As Integer               '��ǰ�ؼ�����
Private mstrSymbol As String                '�����ַ�
Private mblnClear As Boolean                '���Ϊ��,�����ش�mrsDataMap��¼��;����ҳʱӦ����,�����û��޸ĵ������Ա���ʾ������ʹ��
Private mstrCollectItems As String          '������Ŀ����
Private mstrColCollect As String            '������Ŀ�м���:col;1|col;4,5
Private mstrCOLNothing As String            'δ�󶨵��м���+���Ŀ��(���ܻ��Ŀ���Ƿ��)
Private mstrCOLActive As String             '��м���
Private mstrCatercorner As String           '�жԽ��߼���
Private mblnEditAssistant As Boolean        '��ǰѡ�����Ŀ�Ƿ�������дʾ�ѡ��
Private mlngPageRows As Long                '���ļ���ʽһҳ����ʾ��������
Private mlngOverrunRows As Long             '����������
Private mlngRowCount As Long                '��ǰ��¼������
Private mlngRowCurrent As Long              '��ǰ��¼�ڱ�ҳ��ʵ������
Private mlngDate As Long                    '����
Private mlngTime As Long                    'ʱ��
Private mlngOperator As Long                '��ʿ
Private mlngSignLevel As Long               'ǩ������
Private mlngSigner As Long                  'ǩ����Ϣ
Private mlngSignName As Long                'ǩ����
Private mlngSignTime As Long                'ǩ��ʱ��
Private mlngRecord As Long                  '��¼ID
Private mlngNoEditor As Long                '��ֹ�༭��,���ڻ�ʿ�����Ի�ʿ��Ϊ׼,�����ڻ�ʿ������ǩ����Ϊ׼
Private mlngCollectType As Long             '�������
Private mlngCollectText As Long             '�����ı�
Private mlngCollectStyle As Long            '���ܱ��
Private mlngCollectDay As Long              '��������:0-����;1-����
Private mlngPrintedPage As Long             '��ӡҳ��
Private mlngPrintedRow As Long              '��ӡ�к�

Private mblnSign As Boolean                 '�Ƿ�ǩ��
Private mblnArchive As Boolean              '�Ƿ�鵵
Private mintType As Integer                 '��¼��ǰ�ı༭ģʽ
Private mblnDateAd As Boolean               '������д?
Private mstr��ʼʱ�� As String              '��ǰ�ļ��Ŀ�ʼʱ��
Private mstr����ʱ�� As String              '��ǰ�ļ��Ľ���ʱ��
Private CellRect As RECT

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '���л����¼��Ŀ�嵥
Private mrsSelItems As New ADODB.Recordset          '��ǰ¼��Ļ����¼��Ŀ�嵥
Private mrsDataMap As New ADODB.Recordset           '��ǰ���ݾ���

Private Enum ColIcon
    ǩ�� = 1
    ��ǩ = 2
End Enum
Private Enum SignLevel
    ���� = 1
    ���� = 2
    �м� = 3
    ʦ�� = 4
    Աʿ = 5
    δ���� = 9
End Enum

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRefresh()
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

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
Private mstrSubhead As String       '���ϱ�ǩ
Private mstrTabHead As String       '��ͷ��Ԫ
Private mstrColWidth As String      '�п����д�
Private mstrColumns As String       '��ǰ�����ļ����ж�Ӧ����Ŀ
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
'����򿪻����¼�ļ���SQL���������ط�Ҳ��ʹ�ã������޸�
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL���� As String
Private mstrSQL As String

'##############################################################################################
'ҳüҳ�Ŵ�ӡ���
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'����
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'�������ڸ�ʽ��ָ���豸�������Ϣ
Private Type FORMATRANGE
    hDC As Long             '��Ⱦ�豸
    hdcTarget As Long       'Ŀ���豸
    rc As RECT              '��Ⱦ���򣬵�λ��羡�
    rcPage As RECT          '��Ⱦ�豸���������򣬵�λ��羡�
    chrg As CHARRANGE       '���ڸ�ʽ�����ı���Χ��
End Type

Private Type PageInfo
    PageNumber As Long      'ҳ��
    Start As Long           '�ַ���ʼλ��
    End As Long             '�ַ���ֹλ��
    ActualHeight As Long    '��ҳʵ�ʴ�ӡ�߶�
End Type
Private AllPages() As PageInfo   'ҳ��Ϣ
Private Const WM_PASTE = &H302&              'ճ��
Private Const WM_USER = &H400                'ͨ���� WM_USER + X ���Զ�����Ϣ
Private Const EM_FORMATRANGE = (WM_USER + 57)    'Ϊĳһ�豸��ʽ��ָ����Χ���ı���
Private Const EM_SETTARGETDEVICE = (WM_USER + 72) '�����������������õ�Ŀ���豸���п�
Private Const EM_HIDESELECTION = (WM_USER + 63)  '��ʾ/�����ı���
Private Const PHYSICALOFFSETX = 112  '���ڴ�ӡ�豸���ԣ���ʾ������ҳ�����Ե���ɴ�ӡ��������Ե�ľ��룬�����豸��λ��
Private Const PHYSICALOFFSETY = 113  '���ڴ�ӡ�豸���ԣ���ʾ������ҳ���ϱ�Ե���ɴ�ӡ������ϱ�Ե�ľ��룬�����豸��λ��
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '��ȡ��Ӣ�Ļ���ַ�������


'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼���ͼ���,û�±�
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        x As Long
        Y As Long
End Type

Private Const WHITE_BRUSH = 0    '��ɫ����
Private Const cdblWidth As Double = 6          'һ��Ӣ���ַ��Ŀ��
Private Const cHideCols = 2         'ǰ׺������:����,ʱ��
Private Const cControlFields = 2    '��¼��������:ҳ��,�к�

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '��VB����ɫת��ΪRGB��ʾ
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Private Function GetSymbolWidth(ByVal strPara As String) As Double
    'ȱʡ������9��,�������Сͬ�ȷŴ�
    Dim sinFontSize As Single
    Dim i As Integer, j As Integer
    
    j = Len(strPara)
    sinFontSize = VsfData.FontSize
    For i = 1 To j
        GetSymbolWidth = GetSymbolWidth + IIf(Asc(Mid(strPara, i, 1)) > 0, 1, 2) * cdblWidth * sinFontSize / 9
    Next
End Function

Private Sub DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim strText As String
    Dim strLeft As String
    Dim strRight As String
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim dblWidth As Double
    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim blnDraw As Boolean
    '��ͼ���
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim t_ClientRect As RECT
    On Error GoTo errHand
    '******************************************
    '�ڴ��¼��в��ܶԵ�Ԫ����κ����Ը�ֵ,����Celldata,�����������¼�����ѭ��,���¹��������ʱ���޷�����������
    '******************************************
    'ʹ��ƥ��ı���ɫ��ǰ��ɫ����������ı������
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False
    
    strText = VsfData.TextMatrix(ROW, COL)
'    If IsDiagonal(Col) And InStr(1, strText, "/") <> 0 Then
    If InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '����ֵ
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        'ȡ�ַ����
        dblWidth = GetSymbolWidth(strRight)
        '�趨�ͻ������С
        With t_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With
        
        '1���������
        '�����뱳��ɫ��ͬ��ˢ��
        If ROW < VsfData.FixedRows Then
            lngBackColor = GetRBGFromOLEColor(VsfData.BackColorFixed)
            lngForeColor = GetRBGFromOLEColor(VsfData.ForeColorFixed)
        Else
            If ROW = VsfData.RowSel Then
                lngBackColor = GetRBGFromOLEColor(VsfData.BackColorSel)
                lngForeColor = RGB(0, 0, 0)
            Else
                lngBackColor = RGB(255, 255, 255)
                lngForeColor = GetRBGFromOLEColor(VsfData.Cell(flexcpForeColor, ROW, COL))
            End If

        End If
        lngBrush = CreateSolidBrush(lngBackColor)
        'ʹ�ø�ˢ����䱳��ɫ
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, t_ClientRect, lngBrush)
        '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)
        
        '2��׼������
        '�����»���
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '����
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '����ı�
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)
        
        '��ԭ���ʲ�����
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)
        
        '�������ͼ
        Done = True
    End If
    
    '3������ǻ����У���������⴦��
    If Val(VsfData.TextMatrix(ROW, mlngCollectType)) < 0 And Val(VsfData.TextMatrix(ROW, mlngCollectStyle)) > 0 _
        And (COL >= mlngDate And COL < mlngNoEditor) Then
        Call DrawCollectCell(hDC, ROW, COL, Left, Top, Right, Bottom)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawCollectCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Dim lngPen As Long, lngOldPen As Long
    Dim lpPoint As POINTAPI
    
    '�����»���
    lngPen = CreatePen(0, 1, vbRed)
    lngOldPen = SelectObject(hDC, lngPen)
    
    If Val(VsfData.TextMatrix(ROW, mlngCollectStyle)) = 1 Then  '���»�����
        '����
        Call MoveToEx(hDC, Left, Top, lpPoint)
        Call LineTo(hDC, Right, Top)
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Bottom - 2)
    Else                                                        '��������˫����
        If InStr(1, "|" & mstrColCollect & ";", "|" & COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then 'And Val(VsfData.TextMatrix(ROW, COL)) <> 0 Then
            '����
            Call MoveToEx(hDC, Left, Bottom - 4, lpPoint)
            Call LineTo(hDC, Right, Bottom - 4)
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    End If
    
    '��ԭ���ʲ�����
    Call SelectObject(hDC, lngOldPen)
    Call DeleteObject(lngPen)
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼��������,û�±�
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngROWS As Long

    GetData = ""
    lngROWS = SendMessage(txtLength.Hwnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngROWS
        Call ClearArray(strLine)
        Call SendMessage(txtLength.Hwnd, EM_GETLINE, lngRow - 1, strLine(0))
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData & IIf(lngRow < lngROWS, vbCrLf, "")
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = 0 To intMax
        strLine(intDo) = 0
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

'**********************************************************************************************************************
'######################################################################################################################



Private Function GetPeriod() As String
    On Error GoTo errHand
    
    '��ȡָ��ҳ������ݷ���ʱ�䷶Χ
    gstrSQL = " Select /*+ RULE */ MIN(����ʱ��) ��ʼʱ��,MAX(����ʱ��) AS ����ʱ�� From ���˻����ӡ Where �ļ�ID=[1] And (��ʼҳ��=[2] OR ����ҳ��=[2])"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡָ��ҳ������ݷ���ʱ�䷶Χ", mlng��ǰ�ļ�ID, mintҳ��)
    If NVL(rsTemp!��ʼʱ��) = "" Then
        If mintӤ�� = 0 Then
            gstrSQL = " Select  /*+ RULE */ ��Ժ���� AS ��ʼʱ��,sysdate AS ����ʱ�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
        Else
            gstrSQL = " Select  /*+ RULE */ ����ʱ�� AS ��ʼʱ��,sysdate AS ����ʱ�� From ������������¼ Where ����ID=[1] And ��ҳID=[2] And ���=[3]"
        End If
        Set rsTemp = OpenSQLRecord(gstrSQL, "ȡ��Ժ���ڻ��������", mlng����id, mlng��ҳid, mintӤ��)
    End If
    GetPeriod = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ReadStruDef()
    Dim lngCOL As Long
    On Error GoTo errHand
    
    '��ȡ�ļ�����
    mblnDateAd = False
    Call GetFileProperty
    
    '��ȡ���Ŀ�������ж���(��ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...)
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""
    gstrSQL = " Select  /*+ RULE */ A.�к�,A.��ͷ����,A.���,A.��Ŀ���,A.��λ From ���˻���ҳ��_���Ŀ A " & _
              " Where A.�ļ�ID=[1] And A.ҳ��=[2] " & _
              " Order by A.�к�,A.���"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ�������Զ���Ļ��Ŀ", mlng��ǰ�ļ�ID, mintҳ��)
    If rsTemp.RecordCount <> 0 Then
        Do While Not rsTemp.EOF
            If lngCOL <> rsTemp!�к� Then
                lngCOL = rsTemp!�к�
                mstrCOLActive = mstrCOLActive & "||" & rsTemp!�к� & ";" & rsTemp!��ͷ���� & "|" & rsTemp!��Ŀ��� & "," & NVL(rsTemp!��λ)
            Else
                mstrCOLActive = mstrCOLActive & ";" & rsTemp!��Ŀ��� & "," & NVL(rsTemp!��λ)
            End If
            rsTemp.MoveNext
        Loop
    End If
    If mstrCOLActive <> "" Then mstrCOLActive = Mid(mstrCOLActive, 3)
    
    '��ȡ�����ļ���ʽ����
    gstrSQL = "Select  /*+ RULE */ d.�������, d.�����ı�, d.Ҫ������" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ�����ļ���ʽ����", mlng��ʽID)
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
                Set lblSubhead.Font = VsfData.Font
                Set Font = lblSubhead.Font
                
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
            Case "��Ч������"
                mlngOverrunRows = 0
                mlngPageRows = Val(!�����ı�)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select  /*+ RULE */ ��ʽ, ҳ��, ����||'-'||��� AS KEY From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", mlng��ʽID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!��ʽ
        If picHead.Tag = "" Then
            '���ǵ�ҽԺ�ڻ����ļ�ҳüҳ�Ÿ�ʽͳһ���˴�ֻ��ȡһ��
            Call ReadPageHead(rtbHead, rsTemp!Key)
            Call ReadPageFoot(rtbFoot, rsTemp!Key)
            picHead.Tag = rsTemp!Key
            chkҳ��.Value = IIf(Val(NVL(rsTemp!ҳ��, 0)) > 0, 1, 0)
            If chkҳ��.Value = 1 Then optPageAlign(Val(NVL(rsTemp!ҳ��, 0)) - 1).Value = True
        End If
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select  /*+ RULE */ d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlng��ʽID)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select  /*+ RULE */ d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
        " Order By d.�������"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ��ͷ��Ԫ����", mlng��ʽID)
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
    Dim bln�Խ��� As Boolean, blnѡ���� As Boolean          '�����һ���ǶԽ�����ѡ����,��ֱ����ȡ��������,ƴ��ͷʱ����ֵ�����/
    Dim lngColumn As Long, blnAddCollect As Boolean
    
    gstrSQL = "Select  /*+ RULE */ d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", mlng��ʽID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = ""
        mstrSQL�� = "": mstrSQL�� = "": strSql�� = "": mstrSQL�� = "": mstrSQL���� = ""
        bln���� = False: blnʱ�� = False: bln��ʿ = False
        blnǩ���� = False: blnǩ��ʱ�� = False: blnǩ������ = False
        Do While Not .EOF
            If lngColumn <> !������� Then
                blnAddCollect = False
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) & "|" & !������� & "'" & !Ҫ������
                mstrColWidth = mstrColWidth & "," & !�������� & "`" & !������� & "`" & !Ҫ�ر�ʾ
                If !Ҫ�ر�ʾ = 1 Then mstrCatercorner = mstrCatercorner & "," & !�������
                str��ʽ = ""
                If !Ҫ������ <> "" Then
                    str��ʽ = "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                    mstrSQL�� = mstrSQL�� & "," & IIf(Mid(strSql��, 3) = "", "''", Mid(strSql��, 3)) & " As C" & Format(lngColumn, "00")
                Else
                    If strSql�� <> "" Then
                        mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql�� = ""
                lngColumn = !�������
                bln�Խ��� = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                blnѡ���� = False
                mrsItems.Filter = "��Ŀ����='" & NVL(!Ҫ������) & "'"
                If mrsItems.RecordCount <> 0 Then
                    blnѡ���� = (mrsItems!��Ŀ��ʾ = 5)
                    If mrsItems!��Ŀ��ʾ = 4 Then   '������Ŀ
                        blnAddCollect = True
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!��Ŀ���
                        mstrColCollect = mstrColCollect & "|" & !������� & ";" & mrsItems!��Ŀ���
                    End If
                End If
                mrsItems.Filter = 0
            Else
                mstrColumns = mstrColumns & "," & !Ҫ������
                str��ʽ = str��ʽ & "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                mrsItems.Filter = "��Ŀ����='" & NVL(!Ҫ������) & "'"
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!��Ŀ��ʾ = 4 Then   '������Ŀ
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!��Ŀ���
                        If blnAddCollect Then
                            mstrColCollect = mstrColCollect & "," & mrsItems!��Ŀ���
                        Else    '�п���һ�а�������Ŀ,��һ����Ŀ���ǻ�����Ŀ,�ڶ�����Ŀ���ǻ�����Ŀ,���,����Ĵ��뱣֤���������
                            mstrColCollect = mstrColCollect & "|" & !������� & ";" & mrsItems!��Ŀ���
                        End If
                    End If
                End If
                mrsItems.Filter = 0
            End If
            
            Select Case !Ҫ������
            Case "����"
                bln���� = True
                mblnDateAd = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                mstrSQL�� = mstrSQL�� & ",����"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As ����"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case "ʱ��"
                blnʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ʱ��"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ����"
                blnǩ���� = True
                mstrSQL�� = mstrSQL�� & ",ǩ����"
                mstrSQL�� = mstrSQL�� & ",l.ǩ����"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ��ʱ��"
                blnǩ��ʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ǩ��ʱ��"
                mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
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
                    
                    If bln�Խ��� And blnѡ���� Then
                        If strSql�� <> "" Then
                            '�ڶ���
                            strSql�� = strSql�� & "||'/'||""" & !Ҫ������ & """"
                        Else
                            '��һ��
                            strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                        End If
                    Else
                        strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                    End If
                    
                    If (Trim("" & !�����ı�) = "" And Trim("" & !Ҫ�ص�λ) = "") Or (bln�Խ��� And blnѡ����) Then
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & !Ҫ������ & """"
                    Else
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')), '') As """ & !Ҫ������ & """"
                    End If
                Else
                    'Ϊ�ձ�ʾδ����,ǿ�Ƽ�,��������滻
                    mstrCOLNothing = mstrCOLNothing & "," & Format(!�������, "00")
                    mstrSQL�� = mstrSQL�� & ",Max(""" & "C" & Format(!�������, "00") & """) As C" & Format(!�������, "00")
                    mstrSQL���� = mstrSQL���� & " Or """ & "C" & Format(!�������, "00") & """ Is Not Null"
                    mstrSQL�� = mstrSQL�� & ", C" & Format(!�������, "00") & " AS C" & Format(!�������, "00")
                End If
            End Select
            .MoveNext
        Loop
        
        If mstrCollectItems <> "" Then
            mstrCollectItems = Mid(mstrCollectItems, 2)
            mstrColCollect = Mid(mstrColCollect, 2)
        End If
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '�������һ�еĸ�ʽ
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) '& "|" & !������� & "'" & !Ҫ������
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
        If blnǩ��ʱ�� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
        
        If Mid(mstrSQL��, 2) = "" Then
            MsgBox "�Բ�����û�ж��嵱ǰ��������ʾ����Ϣ�����ڲ����ļ������ж��壡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�����ڲ��������ӹ̶���
        mstrSQL�� = UCase(mstrSQL�� & ",MAX(ǩ������) AS ǩ������,MAX(ǩ����Ϣ) AS ǩ����Ϣ,MAX(��¼ID) AS ��¼ID,MAX(����) AS ����,MAX(ʵ������) AS ʵ������,MAX(�������) AS �������,MAX(�����ı�) AS �����ı�,MAX(���ܱ��) AS ���ܱ��,MAX(��������) AS ��������,MAX(��ӡҳ��) AS ��ӡҳ��,MAX(��ӡ�к�) AS ��ӡ�к�")
        mstrSQL�� = UCase(mstrSQL�� & ",l.ǩ������,l.ǩ���� AS ǩ����Ϣ,C.��¼ID,P.����||'' AS ����,DECODE(SIGN(P.����ҳ��-P.��ʼҳ��),1,DECODE(SIGN([5]-P.��ʼҳ��),1, P.�����к�,P.����-P.�����к� ),P.����) AS ʵ������,NVL(L.�������,0) AS �������,L.�����ı�,L.���ܱ��,to_char(L.����ʱ��,'yyyy-MM-dd hh24:mi:ss')||'' AS ��������,p.��ӡҳ��,p.��ӡ�к�")
        mstrSQL�� = UCase(mstrSQL�� & ",ǩ������,ǩ����Ϣ,��¼ID,����,ʵ������,�������,�����ı�,���ܱ��,��������,��ӡҳ��,��ӡ�к�")
        
        '�����Ŀ���뵽SQL��
        Call PreActiveCOL
        'Call SQLCombination
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreActiveHead()
    Dim arrData
    Dim intCol As Integer
    Dim strName As String
    Dim intDo As Integer, intCount As Integer
    '���±�ͷ
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        VsfData.TextMatrix(mintTabTiers - 1, intCol + cHideCols + VsfData.FixedCols - 1) = strName
        If mintTabTiers = 3 And VsfData.TextMatrix(1, intCol + cHideCols + VsfData.FixedCols - 1) = "" Then VsfData.TextMatrix(1, intCol + cHideCols + VsfData.FixedCols - 1) = strName
        If mintTabTiers = 2 And VsfData.TextMatrix(0, intCol + cHideCols + VsfData.FixedCols - 1) = "" Then VsfData.TextMatrix(0, intCol + cHideCols + VsfData.FixedCols - 1) = strName
    Next
End Sub

Private Sub PreActiveCOL()
    Dim arrData
    Dim arrCol
    Dim intCol As Integer, strName As String
    Dim strColFormat As String, strCOLNames As String, strCOLPart As String, strCOLCOND As String, strCOLDEF As String, strCOLMID As String, strCOLIN As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '�����Ŀ���뵽��ѯSQL�У���ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...
    '�󶨶����Ŀ�����о��Զ�תΪ�Խ�����
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        
        '�����б�ʾ(ÿ������������Ŀ)
        strCOLPart = ""
        strCOLNames = ""
        strColFormat = ""
        strCOLCOND = ""
        strCOLMID = ""
        strCOLIN = ""
        strCOLDEF = ""
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            strCOLPart = Split(arrCol(intIn), ",")(1)
            mrsItems.Filter = "��Ŀ���=" & Val(Split(arrCol(intIn), ",")(0))
            strCOLNames = strCOLNames & "," & mrsItems!��Ŀ����
            strCOLCOND = strCOLCOND & " OR """ & strCOLPart & mrsItems!��Ŀ���� & """ IS NOT NULL"
            strCOLMID = strCOLMID & ",Max(""" & strCOLPart & mrsItems!��Ŀ���� & """) As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            If intIn = 0 Then
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.���²�λ||") & "c.��Ŀ����, '" & strCOLPart & mrsItems!��Ŀ���� & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            Else
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.���²�λ||") & "c.��Ŀ����, '" & strCOLPart & mrsItems!��Ŀ���� & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'/','/'||c.��¼����||'')), '') As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            End If
            If intIn = 0 Then
                If intMax = 0 Then
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """ AS C" & intCol
                Else
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """||"
                End If
            Else
                strCOLDEF = strCOLDEF & "NVL(" & strCOLPart & mrsItems!��Ŀ���� & ",'/') AS C" & intCol
            End If
            
            strColFormat = strColFormat & "{[" & strCOLPart & mrsItems!��Ŀ���� & "]" & IIf(intMax > 0 And intIn = 0, "/", "") & "}"
        Next
        If strCOLPart <> "" Then
            strCOLPart = Mid(strCOLPart, 2)
        End If
        strCOLNames = Mid(strCOLNames, 2)
        
        '�Խ���
        If intMax > 0 Then
            mstrCatercorner = mstrCatercorner & IIf(mstrCatercorner = "", "", ",") & intCol
        End If
        '�и�ʽ:15'��ʿ'1'{[��ʿ]}
        mstrColumns = Replace(mstrColumns, intCol & "''1'", intCol & "'" & strCOLNames & "'1'" & strColFormat)
        '��
        mstrSQL�� = Replace(mstrSQL��, "'' AS C" & Format(intCol, "00"), strCOLDEF)
        '����
        mstrSQL���� = Replace(UCase(mstrSQL����), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)
        '��
        mstrSQL�� = Replace(mstrSQL��, ",MAX(""" & "C" & Format(intCol, "00") & """) AS C" & Format(intCol, "00"), strCOLMID)
        '��
        mstrSQL�� = Replace(mstrSQL��, ", C" & Format(intCol, "00") & " AS C" & Format(intCol, "00"), strCOLIN)
    Next
    mrsItems.Filter = 0
    
    '��δ�󶨵��е�SQL���������ش�
    If mstrCOLNothing = "" Then Exit Sub
    arrData = Split(mstrCOLNothing, ",")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        '��(����Ҫ����)
'        mstrSQL�� = Replace(mstrSQL��, ",'' AS C" & arrData(intDo), "")
        '����
        mstrSQL���� = Replace(UCase(mstrSQL����), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")
        '��
        mstrSQL�� = Replace(mstrSQL��, ",MAX(""" & "C" & Format(arrData(intDo), "00") & """) AS C" & Format(arrData(intDo), "00"), "")
        '��
        mstrSQL�� = Replace(mstrSQL��, ", C" & Format(arrData(intDo), "00") & " AS C" & Format(arrData(intDo), "00"), "")
    Next
End Sub

Private Sub SQLCombination(ByVal str���� As String)
    mstrSQL = "Select  /*+ RULE */ ����,����ʱ��," & Mid(mstrSQL��, 12) & vbCrLf & _
                " From (Select ��¼���,ʱ�� as ����,����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select c.��¼���,to_char(l.����ʱ��,'yyyy-MM-dd hh24:mi:ss') AS ����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻����ļ� f,���˻����ӡ p " & vbCrLf & _
                "               Where l.ID=p.��¼ID And l.Id = c.��¼id And l.�ļ�ID=f.ID And f.ID=p.�ļ�ID " & _
                "               And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] " & str���� & ")" & vbCrLf & _
                IIf(mstrSQL���� <> "", "Where " & mstrSQL����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ��ʱ��" & _
                                "       Order By ����ʱ��,��¼���,��ʿ,ǩ����,ǩ��ʱ��)"
End Sub

Private Sub zlReadTip(aryPeriod)
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCOL As Long, lngCount As Long, strCell As String
    Dim strTmpSQL As String
    Dim strTmp As String
    
    Err = 0: On Error GoTo errHand
    
    '���ϱ�ǩ��ȡ
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as ��Ϣ From Dual"
    aryItem = Split(mstrSubhead, "|")
        
    For lngCount = 0 To UBound(aryItem)
        strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
        strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
        
        strTmp = strPrefix
        Select Case strItemName
        Case "��ǰ����"
        
            strTmpSQL = "Select  /*+ RULE */ b.����" & vbNewLine & _
                        "From (Select ����id, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,���ű� b " & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����id Is Not Null And b.ID=a.����id" & vbNewLine & _
                        "Order By a.��ʼʱ��"
                        
            Set rsTemp = OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����id, mlng��ҳid, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            
        Case "��ǰ����"
        
            strTmpSQL = "Select  /*+ RULE */ a.����" & vbNewLine & _
                        "From (Select ����, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���� Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"

            Set rsTemp = OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����id, mlng��ҳid, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "��ǰ����"
        
            strTmpSQL = "Select  /*+ RULE */ ���� From ���ű� a Where a.ID=[1]"
            Set rsTemp = OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID)
            
        Case "סԺҽʦ"
            strTmpSQL = "Select  /*+ RULE */ a.����ҽʦ" & vbNewLine & _
                        "From (Select ����ҽʦ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ҽʦ Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = OpenSQLRecord(strTmpSQL, "סԺҽʦ", mlng����id, mlng��ҳid, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
        Case "���λ�ʿ"
        
            strTmpSQL = "Select  /*+ RULE */ a.���λ�ʿ" & vbNewLine & _
                        "From (Select ���λ�ʿ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���λ�ʿ Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = OpenSQLRecord(strTmpSQL, "���λ�ʿ", mlng����id, mlng��ҳid, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case "����ȼ�"
            strTmpSQL = "Select  /*+ RULE */ b.����" & vbNewLine & _
                        "From (Select ����ȼ�ID, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,����ȼ� b" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ȼ�ID Is Not Null And b.���=a.����ȼ�ID" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set rsTemp = OpenSQLRecord(strTmpSQL, "����ȼ�", mlng����id, mlng��ҳid, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If rsTemp.BOF = False Then rsTemp.MoveLast
            
        Case Else
            strTmp = ""
            Set rsTemp = OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, mlng����id, mlng��ҳid, mintӤ��)
        End Select
        
        If rsTemp.BOF = False Then
            If strTmp <> "" Then
                lblSubhead.Tag = lblSubhead.Tag & " " & strTmp & rsTemp.Fields(0).Value
            Else
                lblSubhead.Tag = lblSubhead.Tag & " " & rsTemp.Fields(0).Value
            End If
        End If
    Next
    lblSubhead.Tag = Trim(lblSubhead.Tag)
    
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlRefresh(ByVal str���� As String)
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCOL As Long, lngCount As Long, strCell As String
    Dim strTmpSQL As String
    Dim strTmp As String
    
    Err = 0: On Error GoTo errHand
    
    'װ������
    Call SQLCombination(str����)
    gstrSQL = mstrSQL
    If gblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ��������", mlng��ǰ�ļ�ID, mlng����id, mlng��ҳid, mintӤ��, mintҳ��, mlngPageRows)
    '���ݼ�¼��,���ڿ��ٻָ�
    Set mrsDataMap = CopyNewRec(rsTemp, mrsDataMap)
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, rsTarget As ADODB.Recordset) As ADODB.Recordset
    'ֻ������¼���Ľṹ,ͬʱ����ҳ��,�к��ֶ�
    Dim intFields As Integer
    
    With rsTarget
        If .Fields.Count = 0 Then
            For intFields = 0 To rsSource.Fields.Count - 1
                If rsSource.Fields(intFields).Name = "��������" Then
                    .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:��ʾ����
                ElseIf rsSource.Fields(intFields).Type = 200 Then       '�����ʹ���Ϊ�ַ���
                    .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
                Else
                    .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
                End If
            Next
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End If
        
        If rsSource.RecordCount <> 0 Then rsSource.MoveFirst
        Do While Not rsSource.EOF
            .AddNew
            For intFields = 0 To rsSource.Fields.Count - 1
                .Fields(intFields) = rsSource.Fields(intFields).Value
            Next
            .Update
            rsSource.MoveNext
        Loop
    End With
    
    Set CopyNewRec = rsTarget
End Function

Private Sub PreTendMutilRows()
    Dim lngRowCount As Long, lngRowCurrent As Long  '��ǰ��¼������,��ǰ��¼�ڱ�ҳ��ʵ������
    Dim lngCOL As Long, lngMax As Long
    Dim lngRow As Long, lngStart As Long, lngPrintedRow As Long
    Dim blnDelete As Boolean
    On Error GoTo errHand
    
    Dim arrData
    Dim intData As Integer, intDatas As Integer
    '���һ����ʾ�����������ʾ(���ݵ�ǰ����ռ����������ӿհ��в�����������,Ȼ�������δ���ǰ�е�����)
    'ÿҳֻ��ʾʵ�ʵ�������,��'@��ȡ��ע�ͼ���
    '���µ����������ݵ�ʵ����
    
    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") <> 0 Then Exit Do
        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
'        @ʵ��������
        lngPrintedRow = Val(VsfData.TextMatrix(lngRow, mlngPrintedRow))
        If lngPrintedRow = 0 Then
            lngRowCurrent = VsfData.TextMatrix(lngRow, mlngRowCurrent)
        Else
            If mlngPageRows < (lngRowCount + lngPrintedRow - 1) Then
                'ʼ�յ������У������ҳ���ݵĿ�ҳ���ж���
                VsfData.TextMatrix(lngRow, mlngRowCurrent) = (lngRowCount + lngPrintedRow - mlngPageRows - 1)
            End If
        End If
        
        If lngRowCount > 1 Then
            '�����ӿ���
            VsfData.Rows = VsfData.Rows + lngRowCount - 1
            '�ӵ�ǰ�е���һ�п�ʼ��ÿ�е�λ��+�����ӵĿհ���������֤�����Ŀհ��дӵ�ǰ�е���һ�п�ʼ
            For intData = VsfData.Rows - lngRowCount To lngRow + 1 Step -1
                VsfData.RowPosition(intData) = intData + lngRowCount - 1
            Next
            
            'ѭ������ǰ������
            For lngCOL = 0 To VsfData.Cols - 1
                If VsfData.ColHidden(lngCOL) And lngCOL <> mlngRowCount Then
                    'ѭ����ֵ
                    For intData = 2 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCOL) = VsfData.TextMatrix(lngRow, lngCOL)
                    Next
                ElseIf (lngCOL < mlngNoEditor And lngCOL <> mlngDate And lngCOL <> mlngTime) Then
                    '׼����ֵ
                    With txtLength
                        .Width = VsfData.ColWidth(lngCOL)
                        .Text = VsfData.TextMatrix(lngRow, lngCOL)
                        .FontName = VsfData.CellFontName
                        .FontSize = VsfData.CellFontSize
                    End With
                    arrData = GetData(txtLength.Text)
                    intDatas = UBound(arrData)
                    
                    If intDatas > 0 Then
                        'ѭ����ֵ
                        For intData = 0 To intDatas
                            If VsfData.Rows <= lngRow + intData Then VsfData.Rows = VsfData.Rows + 1
                            VsfData.TextMatrix(lngRow + intData, lngCOL) = arrData(intData)
                        Next
                    End If
                ElseIf lngCOL = mlngNoEditor Then
                        '����ֵ��Ϊ��1��ʼ,������4������,����4|1
                        For intData = 1 To lngRowCount
                            VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                        Next
                    Else
                End If
            Next
            lngRow = lngRow + lngRowCount - 1
        Else
            VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
        End If
        lngRow = lngRow + 1
    Loop
    
    '���ÿҳ��������
    lngRow = VsfData.FixedRows
    Do While True
        '�̶�������ʾ����ʱ����ǩ����
        lngStart = GetStartRow(lngRow)
        
        '���⴦���һ��(��һ�п��ܴ��ڿ�ҳ����)
        If lngRow = VsfData.FixedRows And Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngRow, mlngRowCurrent)) Then
            blnDelete = True
            lngRow = lngRow + Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) - Val(VsfData.TextMatrix(lngRow, mlngRowCurrent))
        End If
        
        If lngStart <> lngRow Then
            If mlngDate > -1 Then VsfData.TextMatrix(lngRow, mlngDate) = VsfData.TextMatrix(lngStart, mlngDate)
            If mlngTime > -1 Then VsfData.TextMatrix(lngRow, mlngTime) = VsfData.TextMatrix(lngStart, mlngTime)
            If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow, mlngOperator) = VsfData.TextMatrix(lngStart, mlngOperator)
            If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow, mlngSignName) = VsfData.TextMatrix(lngStart, mlngSignName)
            If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow, mlngSignTime) = VsfData.TextMatrix(lngStart, mlngSignTime)
        End If
        
        If blnDelete Then
            For lngCOL = lngStart To lngRow - 1
                VsfData.RemoveItem lngStart
            Next
            blnDelete = False
            lngRow = VsfData.FixedRows  'ֻ�����һ�м�¼ɾ�������,���Թ̶�����Ϊ�̶���Ϊ��ʼ��
        End If
        
        lngRow = lngRow + mlngPageRows
        If lngRow > VsfData.Rows - 1 Then Exit Do
    Loop
    
    '������ش�,������ҳ��Ч�����еĲ���ɾ��
    If gintPrintState = 2 Then
        If VsfData.Rows > VsfData.FixedRows + mlngPageRows Then
            VsfData.Rows = VsfData.FixedRows + mlngPageRows
        End If
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat()
    Dim aryItem() As String
    Dim lngRow As Long, lngCOL As Long, lngCount As Long, strCell As String
    On Error GoTo errHand
    
    '���û����¼���ĸ�ʽ
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = mrsDataMap
        
        '��ͷ��д
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '�����ڲ�����������
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngCollectStyle) = True
        .ColHidden(mlngCollectText) = True
        .ColHidden(mlngCollectType) = True
        .ColHidden(mlngCollectDay) = True
        .ColHidden(mlngPrintedPage) = True
        .ColHidden(mlngPrintedRow) = True
        
        '������ͷ
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCOL = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCOL + cHideCols + .FixedCols - 1) = strCell
        Next
        Call PreActiveHead
        
        '�п�����
        Dim blnAlign As Boolean
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
        '�̶��и�ʽΪ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '�ٰ��кϲ�
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) <> .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
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
        
        Call PreTendMutilRows
        
        If gintPrintState <> 2 Then
            mint����ҳ = (VsfData.Rows - VsfData.FixedRows) \ mlngPageRows
            'If (VsfData.Rows - VsfData.FixedRows) Mod mlngPageRows <> 0 Then mint����ҳ = mint����ҳ + 1
            mint����ҳ = mint����ҳ + mint��ǰ��ʼҳ
            If gintPrintState = 1 Then mintҳ�� = mint����ҳ
        End If
        
        Call WriteColor
        Call ShowPage
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function PrintHead() As Boolean
    PrintHead = PrintRTBData(rtbHead, True)
End Function

Public Function PrintFoot() As Boolean
    Dim lngPage As Long
    On Error GoTo errHand
    '���Ҫ��ӡҳ�����ȴ�ӡҳ��,�ٴ�ӡҳ��
    
    If mintNORule = 1 Then
        If gintPrintState = 1 Then
            'ȡ��ǰ�ļ������ҳ,���δ��ӡ��,��һҳ��ȡ���ҳ����
            lngPage = mlng��ʼҳ�� + mintҳ�� - mint��ǰ��ʼҳ
        Else
            lngPage = mintҳ��
        End If
    Else
        lngPage = mintҳ��
    End If
    mlng��ǰҳ�� = lngPage
    
    PrintFoot = PrintRTBData(rtbFoot, False, lngPage)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function PrintRTBData(ByVal objRTB As RichTextBox, ByVal blnHead As Boolean, Optional ByVal lngPage As Long = 0) As Boolean
    Dim fr As FORMATRANGE           '��ʽ�����ı���Χ
    Dim rcDrawTo As RECT            'Ŀ����������
    Dim rcPage As RECT              'Ŀ��ҳ������
    Dim gTargetDC As Long
    Dim lngFoot As Long
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
    Dim lngNextPos As Long, lngLen As Long, lngTmp As Long, lngPageCount As Long
    Dim rsTemp As New ADODB.Recordset
    
    lngLen = lstrlen(objRTB.Text)
    lngOffsetLeft = gobjOutTo.ScaleX(GetDeviceCaps(gobjOutTo.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = gobjOutTo.ScaleY(GetDeviceCaps(gobjOutTo.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
    
    If blnHead Then
        gobjOutTo.Print ""
    Else
        If chkҳ��.Value = 1 Then
            lngFoot = 180
            gobjOutTo.CurrentY = gobjOutTo.Height - gobjOutTo.ScaleX(gobjSend.EmptyDown, vbMillimeters, vbTwips) - 200
            If optPageAlign(0).Value Then
                gobjOutTo.CurrentX = gobjOutTo.ScaleX(gobjSend.EmptyLeft, vbMillimeters, vbTwips) - 30
            ElseIf optPageAlign(1).Value Then
                gobjOutTo.CurrentX = (gobjOutTo.Width - 90 * LenB(StrConv("ҳ��:" & mintҳ��, vbFromUnicode))) / 2
            Else
                gobjOutTo.CurrentX = gobjOutTo.Width - gobjOutTo.ScaleX(gobjSend.EmptyRight, vbMillimeters, vbTwips) - 90 * LenB(StrConv("ҳ��:" & mintҳ��, vbFromUnicode))
            End If
            gobjOutTo.Print "ҳ��:" & lngPage
        Else
            gobjOutTo.Print ""
        End If
    End If
    
    gTargetDC = hDC
    With rcPage
        .Left = 0
        .Top = 0
        .Right = gobjOutTo.Width
        .Bottom = gobjOutTo.Height
    End With
    With rcDrawTo
        If blnHead Then
            .Left = lngOffsetLeft
            .Top = lngOffsetTop
            .Right = gobjOutTo.Width - lngOffsetLeft
            .Bottom = gobjOutTo.ScaleX(gobjSend.EmptyUp, vbMillimeters, vbTwips) - 30
        Else
            .Left = lngOffsetLeft
            .Top = gobjOutTo.Height - gobjOutTo.ScaleX(gobjSend.EmptyDown, vbMillimeters, vbTwips) + lngFoot
            .Right = gobjOutTo.Width - lngOffsetLeft
            .Bottom = gobjOutTo.Height
        End If
    End With
    With fr
        .hDC = gobjOutTo.hDC
        .hdcTarget = gTargetDC
        .rc = rcDrawTo
        .rcPage = rcPage
        .chrg.cpMin = 0
        .chrg.cpMax = -1
    End With
    
    Do
        lngNextPos = SendMessage(objRTB.Hwnd, EM_FORMATRANGE, 0, fr)
        
        lngPageCount = lngPageCount + 1             ' ҳ����1
        '��¼��ҳ��Ϣ
        ReDim Preserve AllPages(1 To lngPageCount) As PageInfo
        AllPages(lngPageCount).PageNumber = lngPageCount
        AllPages(lngPageCount).ActualHeight = fr.rc.Bottom - fr.rc.Top          'ʵ�ʴ�ӡ�߶�
        AllPages(lngPageCount).Start = lngTmp
        AllPages(lngPageCount).End = lngNextPos
        
        fr.chrg.cpMin = lngNextPos
        If lngNextPos <= lngTmp Or lngNextPos >= lngLen Then Exit Do      ' �������ҳ��ķ�ҳ
        lngTmp = lngNextPos
    Loop
    Call SendMessage(objRTB.Hwnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    
    For lngLen = 1 To lngPageCount
        If lngLen > 1 Then Exit For
        With fr
            .hDC = gobjOutTo.hDC
            .hdcTarget = gTargetDC
            .rc = rcDrawTo
            .rcPage = rcPage
            .chrg.cpMin = AllPages(lngLen).Start
            .chrg.cpMax = AllPages(lngLen).End
        End With
        Call SendMessage(objRTB.Hwnd, EM_FORMATRANGE, 1, fr)
        Call SendMessage(objRTB.Hwnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    Next
    
End Function

Public Function PrintPage() As Boolean
    Dim strSQL() As String
    Dim blnTrans As Boolean
    Dim blnSave As Boolean          '�Ѵ�ӡ�����ݲ�����
    Dim strTime As String
    Dim strCurrDate As String
    Dim lngRow As Long, lngROWS As Long
    Dim intMax As Integer, intPos As Integer
    Dim lngCurRow As Long, lngDataLines As Long
    On Error GoTo errHand
    
    '����ʾ�н��д���
    lngROWS = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngROWS
        If Not VsfData.RowHidden(lngRow) Then
            If lngCurRow = 0 Then lngCurRow = 1
            If VsfData.TextMatrix(lngRow, mlngRowCount) Like "*|1" Then
                lngDataLines = Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)
                blnSave = (Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) = 0) Or gintPrintState > 1
                
                If blnSave Then
                    strTime = VsfData.TextMatrix(lngRow, 1)
                    gstrSQL = "ZL_���˻����ӡ_PRINT(" & mlng��ǰ�ļ�ID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'),'" & gstrUserName & "'," & IIf(mlng��ǰҳ�� = 0, mintҳ��, mlng��ǰҳ��) & "," & lngCurRow & ")"
                    Debug.Print gstrSQL
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                End If
            End If
            lngCurRow = lngCurRow + 1
        End If
    Next
    
    On Error Resume Next
    intMax = UBound(strSQL)

    gcnOracle.BeginTrans
    blnTrans = True

    On Error GoTo errHand
    If intMax > 0 Then
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
                gcnOracle.Execute strSQL(intPos), , adCmdStoredProc
            End If
        Next
    End If

    gcnOracle.CommitTrans
    blnTrans = False
    PrintPage = True
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitRecords()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim lngCOL As Long, lngOrder As Long, strName As String, intImmovable As Integer, strFormat As String
    Dim arrColumn, arrItem, strColumns As String
    Dim blnSet As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    strColumns = mstrColumns
    If Not mblnInit Then
        '��ʼ���ڴ��¼��(δ��Ӧ��Ŀ����Ϊ���Ŀ,�����о�Ϊ�̶���)
        strFields = "��," & adDouble & ",18|���," & adDouble & ",2|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",20|�̶�," & adDouble & ",2|��ʽ," & adLongVarChar & ",2000"
        Call Record_Init(mrsSelItems, strFields)
        strFields = "��|���|��Ŀ���|��Ŀ����|�̶�|��ʽ"
    End If
    
    '�����ж���
    If Not mblnInit Then
        arrColumn = Split(strColumns, "|")
        j = UBound(arrColumn)
        For i = 0 To j
            lngCOL = Split(arrColumn(i), "'")(0)
            arrItem = Split(Split(arrColumn(i), "'")(1), ",")
            blnSet = False   '����������Դ���ֵΪ׼'�����Ҳ�����Ŀ���ǻ��Ŀ
            If UBound(Split(arrColumn(i), "'")) > 1 Then
                blnSet = True
                intImmovable = Split(arrColumn(i), "'")(2)
            End If
            If UBound(Split(arrColumn(i), "'")) > 2 Then
                strFormat = Split(arrColumn(i), "'")(3)
            End If
            
            k = UBound(arrItem)
            For l = 0 To k
                strName = arrItem(l)
                mrsItems.Filter = "��Ŀ����='" & strName & "'"
                If mrsItems.RecordCount <> 0 Then
                    lngOrder = mrsItems!��Ŀ���
                    If Not blnSet Then intImmovable = 1   '�̶��������޸�
                Else
                    lngOrder = 0
                    If Not blnSet Then intImmovable = 0
                    
                    '��¼������
                    Select Case strName
                    Case "����"
                        mlngDate = i + cHideCols + VsfData.FixedCols
                    Case "ʱ��"
                        mlngTime = i + cHideCols + VsfData.FixedCols
                    Case "��ʿ"
                        mlngOperator = i + cHideCols + VsfData.FixedCols
                    Case "ǩ����"
                        mlngSignName = i + cHideCols + VsfData.FixedCols
                    Case "ǩ��ʱ��"
                        mlngSignTime = i + cHideCols + VsfData.FixedCols
                    End Select
                End If
                strValues = lngCOL & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, strFields, strValues)
            Next
        Next
        
        'Call OutputRsData(mrsSelItems)
        
        '��������ڲ�������(�����ڶ�ȡ���ݺ��ʱ���ӵ�,��ʱֻ��Ԥ������)
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '����������
        mlngSigner = mlngSignLevel + 1
        mlngRecord = mlngSigner + 1
        mlngRowCount = mlngRecord + 1
        mlngRowCurrent = mlngRowCount + 1
        mlngCollectType = mlngRowCurrent + 1
        mlngCollectText = mlngCollectType + 1
        mlngCollectStyle = mlngCollectText + 1
        mlngCollectDay = mlngCollectStyle + 1
        mlngPrintedPage = mlngCollectDay + 1
        mlngPrintedRow = mlngPrintedPage + 1
        If mlngOperator <> -1 And mlngSignName <> -1 Then
            mlngNoEditor = IIf(mlngOperator < mlngSignName, mlngOperator, mlngSignName)
        Else
            mlngNoEditor = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        End If
    End If
    
    mrsItems.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ShowPage(Optional ByVal intPage As Integer = 0) As Boolean
    '��ʾָ��ҳ�����ݲ����´�ӡ����
    Dim aryPeriod
    Dim strBegin As String, strEnd As String
    Dim lngRow As Long, lngROWS As Long, lngStart As Long
    Dim lngShows As Long
    On Error GoTo errHand
    
    If intPage <> 0 Then mintҳ�� = intPage
    With VsfData
        'С��ҳ����Ч������˵��ֻ��һҳ����
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            lngROWS = .Rows - 1
            For lngRow = .FixedRows To lngROWS
                .RowHidden(lngRow) = True
            Next
        End If
        
        'С��ҳ����Ч������˵��ֻ��һҳ����
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            lngRow = 3 + mlngPageRows * (mintҳ�� - mint��ǰ��ʼҳ)
            lngROWS = 3 + mlngPageRows * (mintҳ�� - mint��ǰ��ʼҳ + 1) - 1
        Else
            lngRow = 3
            lngROWS = .Rows - 1
        End If
        If lngROWS > .Rows - 1 Then lngROWS = .Rows - 1
        '��ȡָ��ҳ��ʱ�䷶Χ
        If lngRow > lngROWS Then
            Exit Function
        End If
        strBegin = .TextMatrix(lngRow, 1)
        lngStart = lngROWS
        lngStart = GetStartRow(lngStart)
        strEnd = .TextMatrix(lngStart, 1)
        aryPeriod = Split(strBegin & "||" & strEnd, "||")
        
        'С��ҳ����Ч������˵��ֻ��һҳ����
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            '��ʾ������
            For lngRow = lngRow To lngROWS
                .RowHidden(lngRow) = False
                lngShows = lngShows + 1
            Next
        End If
        
        ShowPage = True
        Call zlReadTip(aryPeriod)
    End With
    
    '���ô�ӡ�������
    Dim objPrint As New zlPrintTends, objAppRow As zlTabAppRow
    Dim strLable As String, strAppRow As String, lngSpaces As Long
    Dim lngPos As Long, lngMax As Long, lngNumber As Long, blnNumber As Boolean, lngASC As Long
    
    '���ô�ӡ��ʽ
    If UBound(Split(mstrPaperSet, ";")) >= 0 Then SaveSetting "ZLSOFT", "����ģ��\zl9TendFilePrint\Default", "PaperSize", Val(Split(mstrPaperSet, ";")(0))
    If UBound(Split(mstrPaperSet, ";")) >= 1 Then SaveSetting "ZLSOFT", "����ģ��\zl9TendFilePrint\Default", "Orientation", Val(Split(mstrPaperSet, ";")(1))
    If UBound(Split(mstrPaperSet, ";")) >= 2 Then SaveSetting "ZLSOFT", "����ģ��\zl9TendFilePrint\Default", "Height", Val(Split(mstrPaperSet, ";")(2))
    If UBound(Split(mstrPaperSet, ";")) >= 3 Then SaveSetting "ZLSOFT", "����ģ��\zl9TendFilePrint\Default", "Width", Val(Split(mstrPaperSet, ";")(3))
    If UBound(Split(mstrPaperSet, ";")) >= 4 Then objPrint.EmptyLeft = Round(ScaleY(Val(Split(mstrPaperSet, ";")(4)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 5 Then objPrint.EmptyRight = Round(ScaleY(Val(Split(mstrPaperSet, ";")(5)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 6 Then objPrint.EmptyUp = Round(ScaleX(Val(Split(mstrPaperSet, ";")(6)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 7 Then objPrint.EmptyDown = Round(ScaleX(Val(Split(mstrPaperSet, ";")(7)), vbTwips, vbMillimeters), 2)
    
    Set objPrint.Body = VsfData
    objPrint.Title.Text = lblTitle.Caption
    Set objPrint.Title.Font = lblTitle.Font
    Set objPrint.AppFont = lblSubhead.Font
    
    lngSpaces = lblSubhead.Height / 210
    strLable = lblSubhead.Caption
    lngMax = Len(strLable)
    lngNumber = 0
    lngStart = 1
    For lngPos = 1 To lngMax
        '�����ѧ����,��������Ƶ���һ����ʾ
        lngASC = Asc(Mid(strLable, lngPos, 1))

        '����Ƿ񳬿�(���ȳ����п�,���������س����з�)
        If TextWidth(Mid(strLable, lngStart, lngPos - lngStart + 1) & "��") > (Val(Split(mstrPaperSet, ";")(3)) - Val(Split(mstrPaperSet, ";")(4)) - Val(Split(mstrPaperSet, ";")(5)) - 500) Or lngPos = lngMax Or lngASC = 10 Then

            strAppRow = Mid(strLable, lngStart, lngPos - lngStart + 1)
            
            lngStart = lngPos + 1
            
            '���������
            Set objAppRow = New zlTabAppRow
            Call objAppRow.Add(strAppRow)
            Call objPrint.UnderAppRows.Add(objAppRow)
        End If
    Next

    lngMax = Val(Split(mstrPaperSet, ";")(3))
'    If mstrPageHead <> "" Then objPrint.Header = mstrPageHead
'    If mstrPageFoot <> "" Then
'        mstrPageFoot = Replace(mstrPageFoot, "{��ӡʱ��}", Now)
'        mstrPageFoot = Replace(mstrPageFoot, "{ҳ��}", mintҳ�� + mint�ϲ���ʼҳ - 1)
'        mstrPageFoot = Replace(mstrPageFoot, "{��ӡ��}", gstrUserName)
'        objPrint.Footer = LeftB(mstrPageFoot & Space(lngMax), lngMax - objPrint.EmptyLeft - objPrint.EmptyRight)
'    End If
    
    Set gobjSend = objPrint

    '������������
    gstrTabTitle = gobjSend.Title.Text
    gstrTitleFName = gobjSend.Title.Font.Name
    gintTitleFSize = gobjSend.Title.Font.Size
    gblnTitleFItalic = gobjSend.Title.Font.Italic
    gblnTitleFBold = gobjSend.Title.Font.Bold
    glngTitleColor = gobjSend.Title.Color
    '���������Ŀ�������Ŀ������
    gstrAppRowFName = gobjSend.AppFont.Name
    gintAppRowFSize = gobjSend.AppFont.Size
    gblnAppRowFItalic = gobjSend.AppFont.Italic
    gblnAppRowFBold = gobjSend.AppFont.Bold
    glngAppRowColor = gobjSend.AppColor
    gintUpAppRow = gobjSend.UnderAppRows.Count
    gintDownAppRow = gobjSend.BelowAppRows.Count
    
    If gobjSend.FixRow = 0 Then gobjSend.FixRow = gobjSend.Body.FixedRows
    gintFixRow = gobjSend.FixRow
    gintFixCol = gobjSend.FixCol
'    gintRowTotal = gobjSend.Rows
'    gintColTotal = gobjSend.Cols
    gintGroups = 1
    
    gsngDown = gobjSend.EmptyDown
    gsngLeft = gobjSend.EmptyLeft
    gsngRight = gobjSend.EmptyRight
    gsngUp = gobjSend.EmptyUp
    gsngHeader = gobjSend.PageHeader
    gsngFooter = gobjSend.PageFooter
    
    gstrHeader = gobjSend.Header
    gstrHeader = IIf(gstrHeader = "", ";;", gstrHeader)
    gstrFooter = gobjSend.Footer
    gstrFooter = IIf(gstrFooter = "", ";;", gstrFooter)
    
    '��಻��һ�о���һҳ��
    Call GetPrinterSet
    Call CalculateHeight
    Call CalculateRC
    gstr�Խ��� = GetDiagonal
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetStartPage() As Integer
    GetStartPage = mint��ǰ��ʼҳ
End Function

Public Function GetCollectCols() As String
    GetCollectCols = mstrColCollect
End Function

Public Function GetPages() As Integer
    GetPages = mint����ҳ
End Function

Public Function isEndPage() As Boolean
    isEndPage = (mintҳ�� = mint����ҳ)
End Function

Public Sub PrevPage()
    If mintҳ�� > 1 Then
        mintҳ�� = mintҳ�� - 1
        Call ShowPage
    End If
End Sub

Public Function NextPage() As Boolean
    If mintҳ�� < mint����ҳ Then
        mintҳ�� = mintҳ�� + 1
        NextPage = ShowPage
    End If
End Function

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long
    Dim lngRow As Long
    '����Ժ�ɫ��ʾ,��ӡ����ҳ����������Ϊ��ɫ
    
    glngPrintRow = 0
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 1) <> "" Then
                If .TextMatrix(lngCount, mlngPrintedPage) <> "" And gintPrintState = 1 Then
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = &HE0E0E0
                    glngPrintRow = lngCount                 '��¼�¸�����,�Ӵ˿�ʼ���´�ӡ
                Else
                    '�Ե�һ��δ��ӡ������Ϊ��ǰ��ʾҳ
                    If lngRow = 0 And gintPrintState = 1 Then
                        lngRow = lngCount
                        mintҳ�� = (lngCount - VsfData.FixedRows) \ mlngPageRows + mint��ǰ��ʼҳ - 1
                        If (lngCount - VsfData.FixedRows) > mlngPageRows Then
                            If (lngCount - VsfData.FixedRows) Mod mlngPageRows <> 0 Then mintҳ�� = mintҳ�� + 1
                        End If
                        If mintҳ�� < mint��ǰ��ʼҳ Then mintҳ�� = mint��ǰ��ʼҳ
                    End If
                    
                    '����Ժ�ɫ��ʾ
                    blnTag = False
                    If mintTagFormHour < mintTagToHour Then
                        blnTag = (Hour(.TextMatrix(lngCount, 1)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 1)) < mintTagToHour)
                    Else
                        blnTag = (Hour(.TextMatrix(lngCount, 1)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 1)) < mintTagToHour)
                    End If
                    If blnTag Then
                        Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                        .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                    End If
                End If
                '����С�����ʾ
                If VsfData.TextMatrix(lngCount, mlngRowCount) Like "*|1" Then
                    If Val(VsfData.TextMatrix(lngCount, mlngCollectType)) <> 0 Then
                        VsfData.TextMatrix(lngCount, mlngDate) = VsfData.TextMatrix(lngCount, mlngCollectText)
                        VsfData.TextMatrix(lngCount, mlngTime) = VsfData.TextMatrix(lngCount, mlngCollectText)
                    End If
                End If
            End If
        Next
        
        '���δ��ֵ,ȡ���һҳ
        If (lngRow = 0 And gintPrintState = 1) Then
            mintҳ�� = (.Rows - VsfData.FixedRows) \ mlngPageRows + mint��ǰ��ʼҳ - 1
            If (.Rows - VsfData.FixedRows) > mlngPageRows Then
                If (.Rows - VsfData.FixedRows) Mod mlngPageRows <> 0 Then mintҳ�� = mintҳ�� + 1
            End If
            If mintҳ�� < mint��ǰ��ʼҳ Then mintҳ�� = mint��ǰ��ʼҳ
        End If
        If mintҳ�� = 0 Then mintҳ�� = 1
                        
        '���ҳ��>��ǰ��ʼҳ,˵����ʼҳ��Ч
        If gintPrintState = 1 Then
            '�����ǰҳ������ʼҳ,ɾ����Чҳ����
            If mintҳ�� > mint��ǰ��ʼҳ Then
                For lngRow = 1 To mlngPageRows
                    VsfData.RemoveItem VsfData.FixedRows
                Next
                mint��ǰ��ʼҳ = mintҳ��
                glngPrintRow = glngPrintRow - mlngPageRows
                If glngPrintRow < VsfData.FixedRows Then glngPrintRow = 0
            End If
            '�����ʼ�г���һҳ,ɾ����Чҳ����
            If lngRow >= VsfData.FixedRows + mlngPageRows Then
                For lngRow = 1 To mlngPageRows
                    VsfData.RemoveItem VsfData.FixedRows
                Next
                glngPrintRow = glngPrintRow - mlngPageRows
                If glngPrintRow < VsfData.FixedRows Then glngPrintRow = 0
                mint��ǰ��ʼҳ = mint��ǰ��ʼҳ + 1
                mintҳ�� = mintҳ�� + 1
            End If
            If mint����ҳ > mintҳ�� And VsfData.Rows - VsfData.FixedRows <= mlngPageRows Then
                mint����ҳ = mintҳ��
            End If
        End If
    End With
End Sub

Private Sub zlLableBruit()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    
    lblSubhead.Top = lblTitle.Top + lblTitle.Height + 120
    lblSubhead.Width = VsfData.Width
    lblSubhead.Caption = lblSubhead.Tag
    VsfData.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, ScaleWidth - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top
End Sub

Private Sub GetFileProperty()
    '��ȡ�ļ�����
    On Error GoTo errHand
    
    gstrSQL = "Select ����ת�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "�ж������Ƿ�ת��", mlng����id, mlng��ҳid)
    gblnMoved_HL = NVL(rsTemp!����ת��, 0) <> 0
    
    gstrSQL = " Select  /*+ RULE */ ��ʼʱ��,����ʱ��,��ʽID,����ID,�鵵�� From ���˻����ļ� " & _
              " Where ����ID=[1] And ��ҳID=[2] And Ӥ��=[3] And ID=[4] And Rownum<2"
    If gblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
    End If
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", mlng����id, mlng��ҳid, mintӤ��, mlng��ǰ�ļ�ID)
    If rsTemp.RecordCount <> 0 Then
        mlng��ʽID = rsTemp!��ʽID
        mlng����ID = rsTemp!����ID
        mstr��ʼʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
        mstr����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    
    RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitEnv()
    On Error GoTo errHand
    
    '���ִ��ڵ����л����¼��Ŀ
    gstrSQL = " Select  /*+ RULE */ ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ" & _
              " From �����¼��Ŀ B" & _
              " Where B.Ӧ�÷�ʽ<>0 " & _
              " Order by ��Ŀ���"
    Set mrsItems = OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    
    '��ȡ�����ļ���Ź���
    gstrSQL = " Select NVL(����ֵ,0) AS ����ֵ From zlparameters Where ģ��=1255 and ������='�����ļ�ҳ�����'"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ�����ļ���Ź���")
    mintNORule = 0
    If rsTemp.RecordCount <> 0 Then
        mintNORule = rsTemp!����ֵ
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, Optional ByVal intPage As Integer = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngPatiID           ����id
    '       lngPageID           ��ҳid
    '       lngDeptID           Ҫ��ʾ�����¼�Ŀ���
    '       intBaby             Ӥ����־
    '       blnEditable         ���Ϊ��,˵������Ϊ��ѯ�Ӵ�����ʹ��,ȡ����༭��صĹ���
    '       blnClear            ���Ϊ��,�����ش�mrsDataMap��¼��;����ҳʱӦ����,�����û��޸ĵ������Ա���ʾ������ʹ��
    '���أ� ��
    '******************************************************************************************************************
    Dim str�ļ��� As String, lng����ҳ As Long, lng������ As Long
    Dim rsTemp As New ADODB.Recordset
    Dim str���� As String
    Dim blnInitRec As Boolean, blnPrint As Boolean
    On Error GoTo errHand
    Err = 0
    
    mblnInit = False
    mlng��ǰ�ļ�ID = lngFileID
    mintҳ�� = intPage
    mlng����id = lngPatiID
    mlng��ҳid = lngPageId
    mintӤ�� = intBaby
    mlngPageRows = frmAsk.intPageRows
    Set mfrmParent = frmParent
    
    If mrsItems.State = 0 Then
        Call InitEnv            '��ʼ������
    End If
    Call InitVariable
    
    If mintNORule = 1 Then
        '�϶���һ���ļ������˲Ŵ�ӡ��һ���ļ�,����,��ȡ��ǰ�ļ�
        'ȡ����ǰ�ļ����ҳ��
        gstrSQL = " Select MAX(B.��ӡҳ��) AS ҳ��" & _
                  " From ���˻����ӡ B" & _
                  " Where B.�ļ�ID=[1]"
        Set rsTemp = OpenSQLRecord(gstrSQL, "ȡ����ӡҳ��", mlng��ǰ�ļ�ID)
        mlng��ʼҳ�� = NVL(rsTemp!ҳ��, 0)
        
        If mlng��ʼҳ�� = 0 Then
            'ȡ������סԺ�����ļ������ҳ��
            gstrSQL = " Select MAX(B.��ӡҳ��) AS ҳ��" & _
                      " From ���˻����ļ� A,���˻����ӡ B" & _
                      " Where A.ID=B.�ļ�ID And A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3]"
            Set rsTemp = OpenSQLRecord(gstrSQL, "ȡ����ӡҳ��", mlng����id, mlng��ҳid, mintӤ��)
            mlng��ʼҳ�� = NVL(rsTemp!ҳ��, 0) + 1
        Else
            'ȡ��ǰ�ļ����һҳ�����һ�����ݵĴ�ӡ�к�,�������һҳ��+1
            gstrSQL = " Select MAX(B.��ӡ�к�) AS ҳ��" & _
                      " From ���˻����ӡ B" & _
                      " Where B.�ļ�ID=[1] AND B.��ӡҳ��=[2]"
            gstrSQL = " Select ����,��ӡ�к� From ���˻����ӡ Where �ļ�ID=[1] And ��ӡҳ��=[2] And ��ӡ�к�=(" & gstrSQL & ")"
            Set rsTemp = OpenSQLRecord(gstrSQL, "ȡ����ӡҳ��", mlng��ǰ�ļ�ID, mlng��ʼҳ��)
            If rsTemp!���� + rsTemp!��ӡ�к� - 1 > mlngPageRows Then mlng��ʼҳ�� = mlng��ʼҳ�� + 1
        End If
    End If
    
    mlng�ϲ��ļ�ID = 0
    gstrSQL = " Select  /*+ RULE */ MAX(��ӡҳ��) AS ��ӡҳ��,MAX(����ҳ��) AS ҳ�� From ���˻����ӡ Where �ļ�ID=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡָ��ҳ������ݷ���ʱ�䷶Χ", mlng��ǰ�ļ�ID)
    mlng��ӡҳ = NVL(rsTemp!��ӡҳ��, 1)
    mint����ҳ = mlng��ӡҳ
    If mint����ҳ < NVL(rsTemp!ҳ��, 1) Then mint����ҳ = NVL(rsTemp!ҳ��, 1)
    
    '�����һ��δ��ӡ�괦���Ŵ�ӡ
    'If mintҳ�� = 0 Or (mintҳ�� > 0 And gintPrintState <> 3) Then
    If intPage = 0 Then
        gstrSQL = " SELECT /*+ RULE */ ����ҳ��,�����к� FROM ���˻����ӡ " & vbNewLine & _
                  " WHERE �ļ�ID=[1] AND ��ӡ�� IS NOT NULL" & vbNewLine & _
                  "       AND ����ʱ��=(SELECT MAX(����ʱ��) FROM ���˻����ӡ WHERE �ļ�ID=[1] AND ��ӡ�� IS NOT NULL)"
        Set rsTemp = OpenSQLRecord(gstrSQL, "�����һ��δ��ӡ�괦���Ŵ�ӡ", mlng��ǰ�ļ�ID)
        If rsTemp.RecordCount = 0 Then
            intPage = 1
            blnPrint = False
        Else
            intPage = rsTemp!����ҳ��
            If rsTemp!�����к� = mlngPageRows Then intPage = intPage + 1
            blnPrint = True
        End If
    Else
        blnPrint = True
    End If
    mint��ǰ��ʼҳ = IIf(intPage > mlng��ӡҳ, mlng��ӡҳ, intPage)
    
    '��һҳ����������ӡģʽ��
'    If intPage = 1 And (gintPrintState = 1 Or gintPrintState = 3) Then
        '��鵱ǰ�ļ��Ƿ��������ļ�����Ϊ�ϲ���ӡ
        gstrSQL = " Select A.ID,A.�ļ�����" & vbNewLine & _
                  " From ���˻����ļ� A" & vbNewLine & _
                  " Where A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3] And A.����ID=[4]"
        Set rsTemp = OpenSQLRecord(gstrSQL, "��鵱ǰ�ļ��Ƿ��������ļ�����Ϊ�ϲ���ӡ", mlng����id, mlng��ҳid, mintӤ��, mlng��ǰ�ļ�ID)
        If rsTemp.RecordCount <> 0 Then
            mlng�ϲ��ļ�ID = rsTemp!Id
            str�ļ��� = rsTemp!�ļ�����
            '�������ļ������һҳ��ӡ����
            gstrSQL = " SELECT MAX(��ӡҳ��) AS ��ӡҳ��,MAX(��ӡ�к�) AS ��ӡ�к� FROM ���˻����ӡ" & vbNewLine & _
                      " Where �ļ�ID=[1] And ��ӡ�� Is Not NULL AND ��ӡҳ��=" & vbNewLine & _
                      "     (SELECT MAX(��ӡҳ��) AS ��ӡҳ�� FROM ���˻����ӡ WHERE �ļ�ID=[1] AND ��ӡ�� IS NOT NULL)"
            Set rsTemp = OpenSQLRecord(gstrSQL, "�������ļ������һҳ��ӡ����", mlng�ϲ��ļ�ID)
            lng����ҳ = NVL(rsTemp!��ӡҳ��, 0)
            lng������ = NVL(rsTemp!��ӡ�к�, 0)
            If mlng�ϲ��ļ�ID <> 0 And lng����ҳ = 0 Then
                MsgBox "��ǰ�ļ��롰" & str�ļ��� & "������Ϊ�ϲ���ӡ����" & str�ļ��� & "��δ��ӡ��", vbInformation, gstrSysName
                Exit Function
            End If
            If rsTemp!��ӡ�к� = mlngPageRows Then lng����ҳ = lng����ҳ + 1
            If mint��ǰ��ʼҳ < lng����ҳ Then mint��ǰ��ʼҳ = lng����ҳ
            If mint����ҳ < lng����ҳ Then mint����ҳ = mint��ǰ��ʼҳ + mint����ҳ
            mintҳ�� = mint��ǰ��ʼҳ
        End If
'    End If
    If gintPrintState = 2 Then mint����ҳ = mint��ǰ��ʼҳ
    If gintPrintState = 1 And mlng�ϲ��ļ�ID <> 0 And blnPrint Then mint��ǰ��ʼҳ = mint����ҳ: intPage = mint����ҳ: mintҳ�� = mint����ҳ
    
    If mlng�ϲ��ļ�ID <> 0 And mintҳ�� = lng����ҳ Then
        mlng��ǰ�ļ�ID = mlng�ϲ��ļ�ID
        mintҳ�� = lng����ҳ
        Call ReadStruDef
        Call InitRecords
        blnInitRec = True
        str���� = " AND (P.��ӡҳ��=[5] OR (P.��ӡҳ��=[5]-1 AND P.��ӡ�к�+P.����-1>[6]))"
        Call zlRefresh(str����)
    End If
    
    mlng��ǰ�ļ�ID = lngFileID
    mintҳ�� = IIf(intPage > mlng��ӡҳ, mlng��ӡҳ, intPage) 'IIf(gintPrintState = 1, 1, intPage)
    Do While True
        'mintҳ�룺û��ӡǰ��1����ӡ����ʵ�ʵ�ҳ�ţ�������Ҫ�����£���Ȼ���ش��3ҳ��ʼ�ղ�����ʾ��ǰ�ļ���������
        If blnPrint Then
            If mintҳ�� > mint����ҳ Then Exit Do
        Else
            If mintҳ�� > mint����ҳ - mint��ǰ��ʼҳ + 1 Then Exit Do
        End If
        Call ReadStruDef
        If Not blnInitRec Then
            Call InitRecords
            blnInitRec = True
        End If
        
        Select Case gintPrintState
        Case 1  '����
            If mlng�ϲ��ļ�ID = 0 Then
                If mintҳ�� = IIf(intPage > mlng��ӡҳ, mlng��ӡҳ, intPage) Then
                    str���� = " AND (P.��ʼҳ��=[5] OR (P.����ҳ��=[5])) "
                Else
                    str���� = " AND (P.��ʼҳ��=[5]) "
                End If
            Else
                If Not blnPrint Then
                    str���� = " AND P.��ʼҳ��=[5]"
                Else
                    str���� = " AND ((P.��ӡҳ��=[5] OR (P.��ӡҳ��=[5]-1 AND P.��ӡ�к�+P.����-1>[6]))"
                    str���� = str���� & " OR (P.��ӡҳ�� Is NULL))"
                End If
            End If
        Case 2  '�ش�ָ��ҳ
            str���� = " AND (P.��ӡҳ��=[5] OR (P.��ӡҳ��=[5]-1 AND P.��ӡ�к�+P.����-1>[6]))"
        Case 3  '��ָ��ҳ��ʼ�����ش�
            str���� = " AND (P.��ӡҳ��>=[5] OR (P.��ӡҳ��=[5]-1 AND P.��ӡ�к�+P.����-1>[6]))"
        End Select
        
        Call zlRefresh(str����)
        If gintPrintState > 1 Then Exit Do '�ش��ȡָ��ҳ��ֱ���˳�
        mintҳ�� = mintҳ�� + 1
    Loop
    
    '�����ݲ����û����¼���ĸ�ʽ,ͬʱʵ��һ�����ݷ�����ʾ�Ĺ���
    mintҳ�� = IIf(intPage > mlng��ӡҳ, mlng��ӡҳ, intPage)
    Call PreTendFormat
    
    mblnInit = True
    mblnEditable = False
    ShowMe = True
'    Call OutputRsData(mrsSelItems)
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub lblTitle_Click()
    Call NextPage
'    Call PrevPage
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    With picMain
        .Top = 0
        .Left = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    
    Call zlLableBruit
End Sub

Private Sub vsfData_DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call DrawCell(hDC, ROW, COL, Left, Top, Right, Bottom, Done)
End Sub

Private Sub InitVariable()
    '�����ش���
    mlngDate = -1
    mlngTime = -1
    mlngOperator = -1
    mlngSigner = -1
    mlngSignTime = -1
    mlngRecord = -1
    mlngNoEditor = -1
    
    mblnShow = False
    mblnSign = False
    mblnArchive = False
    mblnEditAssistant = False
    
    Set mrsDataMap = New ADODB.Recordset
End Sub

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngROWS As Long
    '��ȡ������ʼ��,������ҳ�򷵻�0
    '�����ҳδ��ʾȫ,��˵��������ҳ,Ҳ����0
    '���������������������в�������
    
    lngROWS = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '������
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '��ǰ��
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    'Ѱ����ʼ��
    For lngRow = lngRow To 3 Step -1
        If VsfData.TextMatrix(lngRow, mlngRowCount) = lngROWS & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    
    GetStartRow = lngStart
End Function

Public Function GetDiagonal() As String
    GetDiagonal = "," & mstrCatercorner & "," & mstrCOLNothing & ","
End Function

Private Function IsDiagonal(ByVal intCol As Integer) As Boolean
    Dim arrCol, arrData
    Dim intDo As Integer, intCount As Integer
    '�ж�ָ�����Ƿ��������жԽ��ߣ�mstrColWidth�ĸ�ʽ��765`11`1`1,765`11`2`1,...����������`�������`�жԽ��ߣ�
    
    IsDiagonal = (InStr(1, "," & mstrCatercorner & "," & mstrCOLNothing & ",", "," & intCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0)
End Function


'######################################################################################################################
'**********************************************************************************************************************
'�����ǻ������������

Private Sub picMain_Resize()
    picMain.Left = 0
    
    lblTitle.Left = 0
    lblTitle.Width = picMain.Width
    
    VsfData.Width = picMain.Width - VsfData.Left * 2
    VsfData.Height = picMain.Height - VsfData.Top
End Sub

Private Sub UserControl_GotFocus()
    On Error Resume Next
    VsfData.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mblnShow = False
    mblnInit = False
    
'    Set objStream = objFileSys.OpenTextFile("C:\WORKLOG.txt", ForAppending, True)
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
End Sub

Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '���¼�¼,���������,������
    'strPrimary:�ֶ���|ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID|5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCol As Integer, intCols As Integer
    Dim strValues As String
    With rsObj
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strValues = ""
            intCols = .Fields.Count - 1
            For intCol = 0 To intCols
                strValues = strValues & "," & .Fields(intCol).Name & ":" & .Fields(intCol).Value
            Next
            Debug.Print Mid(strValues, 2)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Private Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '���ܣ����¶�������
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function
