VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl usrTendFileReader 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
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
      Begin VB.ComboBox cboҳ�� 
         Height          =   300
         Left            =   3405
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3435
         Width           =   1320
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1380
         Picture         =   "usrTendFileReader.ctx":0734
         Style           =   1  'Graphical
         TabIndex        =   13
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
         TabIndex        =   12
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
      Begin VB.Label lblҳ�� 
         AutoSize        =   -1  'True
         Caption         =   "ҳ��λ��"
         Height          =   180
         Left            =   2610
         TabIndex        =   10
         Top             =   3495
         Width           =   720
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
Private mintNORule As Integer               '0-ͳһ���;1-���ļ���ʽ˳����

Private mArrPage                           '��¼����ӡ��ҳ������:��ʽ��ҳ��;��ӡ��ʶ(1-����,2-������ӡ)
Private mlngMinIndex As Long, mlngMaxIndex As Long '������С���������
Private mlng��ǰҳ�� As Long
Private mint��ǰ��ʼҳ As Integer           '��ǰ�ļ�����ʼҳ(�����Ѵ�ӡ����,�Լ�Ԥ�����Ѵ�ӡҳ��ʼԤ��)
Private mint����ҳ As Integer
Private mintҳ�� As Integer
Private mlng��ǰ�ļ�ID As Long
Private mstrMergeID As String  '�ϲ��ļ�
Private mlng��ʽID As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
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
Private mstrColCorrelative As String        '������Ŀ�����м���:COl,3;COl,4|COl,5;COl,6(�����к�,��Ŀ���;������,��Ŀ���),��Ҫ��Է������
Private mstrCOLNothing As String            'δ�󶨵��м���+���Ŀ��(���ܻ��Ŀ���Ƿ��)
Private mstrCOLActive As String             '��м���
Private mstrCatercorner As String           '�жԽ��߼���
Private mblnEditAssistant As Boolean        '��ǰѡ�����Ŀ�Ƿ�������дʾ�ѡ��
Private mlngPageRows As Long                '���ļ���ʽһҳ����ʾ��������
Private mlngOverrunRows As Long             '����������
Private mlngRowCount As Long                '��ǰ��¼������
Private mlngRowCurrent As Long              '��ǰ��¼�ڱ�ҳ��ʵ������
Private mlngStartSpread As Long             '�ж������Ƿ��ڴ�ӡ��ʼҳ���п�ҳ��1-�ǣ�����-��(ʵ�ʿ�ʼ�к�)
Private mlngDate As Long                    '����
Private mlngTime As Long                    'ʱ��
Private mlngOperator As Long                '��ʿ
Private mlngSignLevel As Long               'ǩ������
Private mlngSigner As Long                  'ǩ����Ϣ
Private mlngSignName As Long                'ǩ����
Private mlngSignTime As Long                'ǩ��ʱ��
Private mlngJoinSignName As Long            '����ǩ����
Private mlngRecord As Long                  '��¼ID
Private mlngFileID As Long                  '�ļ�ID����Ҫ���ںϲ��ļ�ʹ��
Private mlngNoEditor As Long                '��ֹ�༭��,���ڻ�ʿ�����Ի�ʿ��Ϊ׼,�����ڻ�ʿ������ǩ����Ϊ׼
Private mlngCollectType As Long             '�������
Private mlngCollectText As Long             '�����ı�
Private mlngCollectStyle As Long            '���ܱ��
Private mlngCollectDay As Long              '��������:0-����;1-����
Private mlngPrintedPage As Long             '��ӡҳ��
Private mlngPrintedRow As Long              '��ӡ�к�
Private mlngPrintedEndPage As Long          '��ӡ����ҳ��,��Ҫ��¼��ҳ���ݵ�ǰ��ӡ����һҳ
Private mlngCollectValue As Long
Private mlngPrintedTag As Long                '��ӡ��ʶ,��¼�ϴδ�ӡ�Ƿ����δ��ҳ��ӡ�հ���
Private mbln����ʱ��ϲ� As Boolean         '������ʱ��ϲ�
Private mblnʱ�������� As Boolean           '����ʱ����(�磺Ѫ�Ǽ�ⵥֻ��Ҫ��ʾ����)
Private Const mlngDemo As Long = 0          '����
Private mlngSingerType As Long              '��ʿ��ǩ������ʾģʽ����������ʾ������β��ʾ�ȣ�
Private mblnSignPic As Boolean            'ǩ������ʾ��ʽ
Private mblnPrintRow As Boolean           '��¼��Ԥ������ӡʱ������δ��ҳ�հײ����Ƿ�������
Private mblnFullPagePrint As Boolean      '��¼��Ԥ������ӡʱ,������ҳ�Ž��д�ӡ
Private mblnOddEvenPagePrint As Boolean   '��¼����ӡʱ������ҳ��ż���
Private mblnDateModel As Boolean          '������ʾ��ʽ����ͬ���ڵ�����ʾһ�Σ�ÿһ����¼����ʾ
Private mlngCollectColor As Long            'С���ʶ��ɫ
Private mblnShowNullCollet As Boolean       'С���Ƿ��ڿ�ֵ�����»�����

Private mblnSign As Boolean                 '�Ƿ�ǩ��
Private mblnArchive As Boolean              '�Ƿ�鵵
Private mintType As Integer                 '��¼��ǰ�ı༭ģʽ
Private mblnDateAd As Boolean               '������д?
Private mstr��ʼʱ�� As String              '��ǰ�ļ��Ŀ�ʼʱ��
Private mstr����ʱ�� As String              '��ǰ�ļ��Ľ���ʱ��
Private CellRect As RECT

Private mrsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '���л����¼��Ŀ�嵥
Private mrsElement As New ADODB.Recordset           '�����ڼ�¼���ı�ǩҪ��
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
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
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
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Const WHITE_BRUSH = 0    '��ɫ����
Private Const cdblWidth As Double = 6          'һ��Ӣ���ַ��Ŀ��
Private Const cHideCols = 3         'ǰ׺������:����,ʱ��,����ʱ��ϲ�ʱ��ʾ������
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
    On Error GoTo ErrHand
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
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawCollectCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Dim lngPen As Long, lngOldPen As Long
    Dim lpPoint As POINTAPI
    
    '�����»���
    lngPen = CreatePen(0, 1, mlngCollectColor)
    lngOldPen = SelectObject(hDC, lngPen)
    
    Select Case Val(VsfData.TextMatrix(ROW, mlngCollectStyle))
    Case 1 '���»�����
        '����
        Call MoveToEx(hDC, Left, Top, lpPoint)
        Call LineTo(hDC, Right, Top)
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Bottom - 2)
    Case 2  '��������˫����
        If IIf(mblnShowNullCollet, True, VsfData.TextMatrix(ROW, COL) <> "") Then
            '����
            Call MoveToEx(hDC, Left, Bottom - 4, lpPoint)
            Call LineTo(hDC, Right, Bottom - 4)
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    Case 3  '�Ϻ���
        '����
        Call MoveToEx(hDC, Left, Top, lpPoint)
        Call LineTo(hDC, Right, Top)
    Case 4 '�������µ�����
        If InStr(1, "|" & mstrColCollect & ";", "|" & COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then 'And Val(VsfData.TextMatrix(ROW, COL)) <> 0 Then
             If IIf(mblnShowNullCollet, True, VsfData.TextMatrix(ROW, COL) <> "") Then
                '����
                Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
                Call LineTo(hDC, Right, Bottom - 2)
            End If
        End If
    End Select
    
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
    Dim lngRow As Long, lngRows As Long, lngLen As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.hWnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        lngLen = SendMessage(txtLength.hWnd, EM_GETLINE, lngRow - 1, strLine(0))
        Call ClearArray(strLine, lngLen)
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData & IIf(lngRow < lngRows, vbCrLf, "")
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte, Optional ByVal lngPos As Long = 0)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = lngPos To intMax
        strLine(intDo) = 0
        If lngPos > 0 Then Exit Sub     '��Ϊ��,��ʾ�������ַ���������
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
    Dim rs As New ADODB.Recordset
    Dim strPeriod As String
    On Error GoTo ErrHand
    
    '53588:������,2013-4-25,�޸����ݵ�ʱ��С�ڲ�����Ժʱ�䣬���ţ�����������ʾ����
    '�磺�������ʱ��Ϊ2013-03-13 11:23:34 �ļ���ʼʱ��������ͬ����ʱ¼������ʱ��Ϊ 2013-03-13 11:23
    '�ͻᵼ���޷���ȡ���ţ�ӦΪ���������ʱ��Ϊ2013-03-13 11:23:00 С���˲������ʱ�䵼���޷���ȡ������
    '��ȡ���˵����ʱ��
    If mintӤ�� = 0 Then
        gstrSQL = "Select ��ʼʱ��, Sysdate As ����ʱ��" & vbNewLine & _
            " From ���˱䶯��¼" & vbNewLine & _
            " Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 2" & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select ��ʼʱ��, Sysdate As ����ʱ��" & vbNewLine & _
            " From ���˱䶯��¼ a" & vbNewLine & _
            " Where a.����id = [1] And a.��ҳid = [2] And a.��ʼԭ�� = 1 And Not Exists" & vbNewLine & _
            " (Select 1 From ���˱䶯��¼ Where ����id = a.����id And ��ҳid = a.��ҳid And ��ʼԭ�� = 2)"
    Else
        gstrSQL = " Select   ����ʱ�� AS ��ʼʱ��,sysdate AS ����ʱ�� From ������������¼ Where ����ID=[1] And ��ҳID=[2] And ���=[3]"
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ���ڻ��������", mlng����ID, mlng��ҳID, mintӤ��)
    
    '��ȡָ��ҳ������ݷ���ʱ�䷶Χ
    gstrSQL = " Select  MIN(����ʱ��) ��ʼʱ��,MAX(����ʱ��) AS ����ʱ�� From ���˻����ӡ Where �ļ�ID=[1] And (��ʼҳ��=[2] OR ����ҳ��=[2])"
    Call SQLDIY(gstrSQL)
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ��ҳ������ݷ���ʱ�䷶Χ", mlng��ǰ�ļ�ID, mintҳ��)
    If NVL(mrsTemp!��ʼʱ��) = "" Then
        strPeriod = Format(rs!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(rs!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    Else
        If Format(mrsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") < Format(rs!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") Then
            strPeriod = Format(rs!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(mrsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
        Else
            strPeriod = Format(mrsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(mrsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    GetPeriod = strPeriod
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadStruDef() As Boolean
    Dim lngCol As Long
    On Error GoTo ErrHand
    
    '��ȡ�ļ�����
    mblnDateAd = False
    Call GetFileProperty
    
    '��ȡ���Ŀ�������ж���(��ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...)
    mbln����ʱ��ϲ� = False
    mblnʱ�������� = False
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""
    mstrColCorrelative = ""
    gstrSQL = " Select   A.�к�,A.��ͷ����,A.���,A.��Ŀ���,A.��λ From ���˻�����Ŀ A " & _
              " Where A.�ļ�ID=[1] And A.ҳ��=[2] " & _
              " Order by A.�к�,A.���"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������Զ���Ļ��Ŀ", mlng��ǰ�ļ�ID, mintҳ��)
    If mrsTemp.RecordCount <> 0 Then
        Do While Not mrsTemp.EOF
            If lngCol <> mrsTemp!�к� Then
                lngCol = mrsTemp!�к�
                mstrCOLActive = mstrCOLActive & "||" & mrsTemp!�к� & ";" & mrsTemp!��ͷ���� & "|" & mrsTemp!��Ŀ��� & "," & NVL(mrsTemp!��λ)
            Else
                mstrCOLActive = mstrCOLActive & ";" & mrsTemp!��Ŀ��� & "," & NVL(mrsTemp!��λ)
            End If
            mrsTemp.MoveNext
        Loop
    End If
    If mstrCOLActive <> "" Then mstrCOLActive = Mid(mstrCOLActive, 3)
    
    '��ȡ�����ļ���ʽ����
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ���ʽ����", mlng��ʽID)
    With mrsTemp
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
                mlngPageRows = Val("" & !�����ı�)
            Case "����ʱ��ϲ�"
                mbln����ʱ��ϲ� = (Val("" & !�����ı�) = 1)
            '65502:������,2013-11-12
            Case "ʱ��������"
                mblnʱ�������� = (Val("" & !�����ı�) = 1)
            End Select
            .MoveNext
        Loop
    End With
    
    If mblnʱ�������� = True Then mbln����ʱ��ϲ� = False
    
    gstrSQL = "Select   ��ʽ,ҳü ,ҳ��, ����||'-'||��� AS KEY From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", mlng��ʽID)
    If Not mrsTemp.EOF Then
        mstrPaperSet = "" & mrsTemp!��ʽ
        If picHead.Tag = "" Then
            '���ǵ�ҽԺ�ڻ����ļ�ҳüҳ�Ÿ�ʽͳһ���˴�ֻ��ȡһ��
            Call ReadPageHead(rtbHead, mrsTemp!Key)
            Call ReadPageFoot(rtbFoot, mrsTemp!Key)
            picHead.Tag = mrsTemp!Key
            chkҳ��.Value = IIf(Val(NVL(mrsTemp!ҳ��, 0)) > 0, 1, 0)
            If chkҳ��.Value = 1 Then
                optPageAlign(Val(NVL(mrsTemp!ҳ��, 0)) - 1).Value = True
                '46251,������,2012-09-11,װ��ҳ�����λ��
                If CInt(Val(NVL(mrsTemp!ҳü, 0))) > 0 And CInt(Val(NVL(mrsTemp!ҳü, 0))) < 5 Then
                    Call zlControl.CboSetIndex(cboҳ��.hWnd, CInt(Val(NVL(mrsTemp!ҳü, 0))) - 1)
                End If
            End If
        End If
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlng��ʽID)
    With mrsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
        " Order By d.�������"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ͷ��Ԫ����", mlng��ʽID)
    With mrsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !�����д� - 1 & "," & !������� & "," & !�����ı�
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '��ѯ�����֯
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql�� As String, str��ʽ As String, strSqlNull As String
    Dim bln���� As Boolean, blnʱ�� As Boolean, bln��ʿ As Boolean
    Dim blnǩ���� As Boolean, blnǩ��ʱ�� As Boolean, blnǩ������ As Boolean
    Dim bln�Խ��� As Boolean, blnѡ���� As Boolean          '�����һ���ǶԽ�����ѡ����,��ֱ����ȡ��������,ƴ��ͷʱ����ֵ�����/
    Dim lngColumn As Long, blnAddCollect As Boolean
    Dim strColCorrelative As String
    Dim str����ֵ As String
    
    gstrSQL = "Select   d.�������,d.������, d.��������, d.�����д�, d.�����ı�, upper(d.Ҫ������) AS Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
        " Order By d.�������, d.�����д�"
        
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", mlng��ʽID)
    With mrsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = "": strColCorrelative = ""
        mstrSQL�� = "": mstrSQL�� = "": strSql�� = "": mstrSQL�� = "": mstrSQL���� = "": strSqlNull = ""
        bln���� = False: blnʱ�� = False: bln��ʿ = False
        blnǩ���� = False: blnǩ��ʱ�� = False: blnǩ������ = False
        Do While Not .EOF
            If lngColumn <> !������� Then
                blnAddCollect = False
                If strColCorrelative <> "" Then
                    mstrColCorrelative = mstrColCorrelative & "|" & strColCorrelative
                End If
                strColCorrelative = ""
                
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) & "|" & !������� & "'" & !Ҫ������
                mstrColWidth = mstrColWidth & "," & !�������� & "`" & !������� & "`" & !Ҫ�ر�ʾ
                If !Ҫ�ر�ʾ = 1 Then mstrCatercorner = mstrCatercorner & "," & !�������
                str��ʽ = ""
                If !Ҫ������ <> "" Then str��ʽ = "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                    
                If Mid(strSqlNull, 3) = "" Then
                    strSqlNull = "''"
                Else
                    strSqlNull = Mid(strSqlNull, 3)
                End If
                mstrSQL�� = mstrSQL�� & "," & IIf(Mid(strSql��, 3) = "", "''", "Decode(" & Mid(strSql��, 3) & "," & strSqlNull & ",''," & Mid(strSql��, 3) & ")") & " As C" & Format(lngColumn, "00")
                
                strSql�� = ""
                strSqlNull = ""
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
                        If Val(NVL(!������)) > 0 And Val(NVL(!�������)) <> Val(NVL(!������)) Then
                            strColCorrelative = Val(NVL(!������)) & ";" & !������� & "," & mrsItems!��Ŀ���
                        End If
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
                            strColCorrelative = ""
                            mstrColCollect = mstrColCollect & "," & mrsItems!��Ŀ���
                        Else    '�п���һ�а�������Ŀ,��һ����Ŀ���ǻ�����Ŀ,�ڶ�����Ŀ���ǻ�����Ŀ,���,����Ĵ��뱣֤���������
                            blnAddCollect = True
                            mstrColCollect = mstrColCollect & "|" & !������� & ";" & mrsItems!��Ŀ���
                            If Val(NVL(!������)) > 0 And Val(NVL(!�������)) <> Val(NVL(!������)) Then
                                strColCorrelative = Val(NVL(!������)) & ";" & !������� & "," & mrsItems!��Ŀ���
                            End If
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
                '51589:������,2013-03-01,��ӽ���ǩ��
                'mstrSQL�� = mstrSQL�� & ",l.ǩ����"
                mstrSQL�� = mstrSQL�� & ",DECODE(TRIM(NVL(L.ǩ����,'')),'',TRIM(L.ǩ����),DECODE(TRIM(NVL(L.����ǩ����,'')),'',TRIM(L.ǩ����), TRIM(L.ǩ����) || '/' || TRIM(L.����ǩ����))) ǩ����"
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
                    
                    strSql�� = strSql�� & "||'" & !�����ı� & "'||""" & !Ҫ������ & """||'" & !Ҫ�ص�λ & "'"
                    strSqlNull = strSqlNull & "||" & "'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "'"
                    mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', c.��¼����, '') As """ & !Ҫ������ & """"
                    
''                    If bln�Խ��� And blnѡ���� Then
''                        If strSql�� <> "" Then
''                            '�ڶ���
''                            strSql�� = strSql�� & "||'/'||""" & !Ҫ������ & """"
''                        Else
''                            '��һ��
''                            strSql�� = strSql�� & "||""" & !Ҫ������ & """"
''                        End If
''                    Else
''                        strSql�� = strSql�� & "||""" & !Ҫ������ & """"
''                        strSqlNull = strSqlNull & "||" & "'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "'"
''                    End If
''
''                    If (Trim("" & !�����ı�) = "" And Trim("" & !Ҫ�ص�λ) = "") Or (bln�Խ��� And blnѡ����) Then
''                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & !Ҫ������ & """"
''                    Else
''                        'mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')), '') As """ & !Ҫ������ & """"
''                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')),  '" & !�����ı� & "'||'" & !Ҫ�ص�λ & "') As """ & !Ҫ������ & """"
''                    End If
                Else
                    'Ϊ�ձ�ʾδ����,ǿ�Ƽ�,��������滻
                    mstrCOLNothing = mstrCOLNothing & "," & Val(Format(!�������, "00"))
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
        '��InitRecords����Ҫ��������Ŀ���е��������������Ŀ���
        If Left(mstrColCorrelative, 1) = "|" Then mstrColCorrelative = Mid(mstrColCorrelative, 2)
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '�������һ�еĸ�ʽ
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) '& "|" & !������� & "'" & !Ҫ������
        mstrColumns = Mid(mstrColumns, 2)     '��ʽ��:�к�;��Ŀ����1,��Ŀ����2|�к�...,ʵ��;1;����|2;����|3...

        If Mid(strSqlNull, 3) = "" Then
            strSqlNull = "''"
        Else
            strSqlNull = Mid(strSqlNull, 3)
        End If
        mstrSQL�� = mstrSQL�� & "," & IIf(Mid(strSql��, 3) = "", "''", "Decode(" & Mid(strSql��, 3) & "," & strSqlNull & ",''," & Mid(strSql��, 3) & ")") & " As C" & Format(lngColumn, "00")
                
                
        
        If mstrSQL���� <> "" Then mstrSQL���� = "(" & Mid(mstrSQL����, 5) & ")"
        
        '���û�г������ڣ�ʱ�䣬��ʿ�����ڲ���Ҫ���䣬�Ա�֤�в�����������
        If bln���� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
        If blnʱ�� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
        If bln��ʿ = False Then mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
        
        '51589:������,2013-03-01,��ӽ���ǩ��
        'If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
        If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",DECODE(TRIM(NVL(L.ǩ����,'')),'',TRIM(L.ǩ����),DECODE(TRIM(NVL(L.����ǩ����,'')),'',TRIM(L.ǩ����), TRIM(L.ǩ����) || '/' || TRIM(L.����ǩ����))) ǩ����"
        If blnǩ��ʱ�� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
        
        If Mid(mstrSQL��, 2) = "" Then
            MsgBox "�Բ�����û�ж��嵱ǰ��������ʾ����Ϣ�����ڲ����ļ������ж��壡", vbInformation, gstrSysName
            Exit Function
        End If
        '50503:������,2012-09-12,���ݴ�ĳһҳ��һ�оͿ�ʼ��ҳ,��ӿ�ʼ�к�
        '56134:������,2012-12-19,���˻����ӡ��Ӵ�ӡ��ʶ
        '46506:������,2012-12-27,���˻����ӡ��Ӵ�ӡ����ҳ�ţ����ڱ�ʶ��ҳ���ݴ�ӡ
        '�����ڲ��������ӹ̶���
        '˵�����Ҫ�������ڡ���ӡ��ʶ����֮ǰ��ӣ��������޸�zlPrintMdl
        str����ֵ = " Decode(Instr('|" & mstrColCollect & "|', ';' || C.��Ŀ��� || '|'), 0," & _
                  " Decode(Instr('|" & mstrColCollect & "|', ',' || C.��Ŀ��� || '|'), 0," & _
                  " Decode(Instr('|" & mstrColCollect & "|', ';' || C.��Ŀ��� || ','), 0, Null, " & _
                  " Substr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ';' || C.��Ŀ��� || ',') - 1) || ';' || C.��Ŀ���," & _
                  " Instr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ';' || C.��Ŀ��� || ',') - 1) || ';' || C.��Ŀ���, '|', -1) + 1)), " & _
                  " Substr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ',' || C.��Ŀ��� || '|') - 1) || ',' || C.��Ŀ���, " & _
                  " Instr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ',' || C.��Ŀ��� || '|') - 1) || ',' || " & _
                  " C.��Ŀ���, '|', -1) + 1)),Substr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ';' || C.��Ŀ��� || '|') - 1) || ';' || C.��Ŀ���, " & _
                  " Instr(Substr('|" & mstrColCollect & "|', 1, Instr('|" & mstrColCollect & "|', ';' || C.��Ŀ��� || '|') - 1) || ';' || C.��Ŀ���, '|', -1) + 1)) ����ֵ "
                  
        mstrSQL�� = UCase(mstrSQL�� & ",MAX(ǩ������) AS ǩ������,MAX(ǩ����Ϣ) AS ǩ����Ϣ,MAX(����ǩ����) AS ����ǩ����,MAX(�ļ�ID) AS �ļ�ID,MAX(��¼ID) AS ��¼ID,MAX(����) AS ����,MAX(ʵ������) AS ʵ������,MAX(��ʼ�к�) AS ��ʼ�к�,MAX(��ӡ����ҳ��) as ��ӡ����ҳ��,f_List2str(Cast(Collect(����ֵ) As t_Strlist), '|') ����ֵ,MAX(��ӡ��ʶ) AS ��ӡ��ʶ,MAX(�������) AS �������,MAX(�����ı�) AS �����ı�,MAX(���ܱ��) AS ���ܱ��,MAX(��������) AS ��������,MAX(��ӡҳ��) AS ��ӡҳ��,MAX(��ӡ�к�) AS ��ӡ�к�")
        mstrSQL�� = UCase(mstrSQL�� & ",l.ǩ������,l.ǩ���� AS ǩ����Ϣ,l.����ǩ����,l.�ļ�ID,C.��¼ID,P.����||'' AS ����,DECODE(SIGN(P.����ҳ��-P.��ʼҳ��),1,DECODE(SIGN([5]-P.��ʼҳ��),1, P.�����к�,P.����-P.�����к� ),P.����) AS ʵ������,DECODE(SIGN(P.����ҳ��-P.��ʼҳ��),1,DECODE(SIGN([5]-P.��ʼҳ��),1,P.��ʼ�к�+P.����-P.�����к�,P.��ʼ�к�),P.��ʼ�к�) ��ʼ�к�,P.��ӡ����ҳ��," & str����ֵ & ",P.��ӡ��ʶ,NVL(L.�������,0) AS �������,L.�����ı�,L.���ܱ��,to_char(L.����ʱ��,'yyyy-MM-dd hh24:mi:ss')||'' AS ��������,p.��ӡҳ��,p.��ӡ�к�")
        mstrSQL�� = UCase(mstrSQL�� & ",ǩ������,ǩ����Ϣ,����ǩ����,�ļ�ID,��¼ID,����,ʵ������,��ʼ�к�,��ӡ����ҳ��,����ֵ,��ӡ��ʶ,�������,�����ı�,���ܱ��,��������,��ӡҳ��,��ӡ�к�")
        
        
        '�����Ŀ���뵽SQL��
        Call DelActiveNoUsed
        Call PreActiveCOL
        'Call SQLCombination
    End With
    
    ReadStruDef = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

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

Private Function DelActiveNoUsed() As Boolean
'------------------------------------------------
'����:ɾ�����ڷǿ����ϵĻ��Ŀ����Ϣ
'����:������,2013-07-16
'�����:63401
'------------------------------------------------
    Dim arrData, arrActive, arrCol
    Dim strSQL As String
    Dim lngCol As Long, intDo As Integer, intCount As Integer
    Dim blnTran As Boolean
    
    If mstrCOLNothing = "" Then DelActiveNoUsed = True: Exit Function
    arrActive = Array()
    arrCol = Array()
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        lngCol = Val(Split(Split(arrData(intDo), "|")(0), ";")(0))
        If InStr(1, "," & mstrCOLNothing & ",", "," & lngCol & ",") <> 0 Then
            '��¼�������������ϵĻ��Ŀ������Ϣ
            ReDim Preserve arrActive(UBound(arrActive) + 1)
            arrActive(UBound(arrActive)) = CStr(arrData(intDo))
        Else
            '��¼��Ҫ�Ƴ��Ļ��Ŀ�к�
            ReDim Preserve arrCol(UBound(arrCol) + 1)
            arrCol(UBound(arrCol)) = lngCol
        End If
    Next
    
    On Error GoTo ErrHand
    
    'ɾ������Ҫ�Ļ��Ŀ��Ϣ(��Ҫ������֮ǰ���������,�����������)
    If UBound(arrCol) > 1 Then
        gcnOracle.BeginTrans
        blnTran = True
    End If
    
    For intDo = 0 To UBound(arrCol)
        If CStr(arrCol(intDo)) <> "" Then
            strSQL = "ZL_���˻���ҳ��_UPDATE(" & mlng��ǰ�ļ�ID & "," & mintҳ�� & "," & Val(arrCol(intDo)) & ",NULL,'" & gstrUserName & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "������Ŀ������")
        End If
    Next
    If blnTran = True Then gcnOracle.CommitTrans
    
    '���¸�����ȡ�Ļ��Ŀ����Ϣ
    If UBound(arrActive) = -1 Then
        mstrCOLActive = ""
    Else
        mstrCOLActive = Join(arrActive, "||")
    End If
    
    DelActiveNoUsed = True
    Exit Function
ErrHand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """ AS C" & Format(intCol, "00")
                Else
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """||"
                End If
            Else
                strCOLDEF = strCOLDEF & "NVL(""" & strCOLPart & mrsItems!��Ŀ���� & """,'/')"
                If intIn = intMax Then
                    strCOLDEF = "Decode(" & strCOLDEF & ",'" & String(intMax, "/") & "',''," & strCOLDEF & ") As C" & Format(intCol, "00")
                End If
            End If
            
            strColFormat = strColFormat & "{[" & strCOLPart & mrsItems!��Ŀ���� & "]" & IIf(intMax > 0 And intIn < intMax, "/", "") & "}"
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
        '77476:LPF:����滻intcolǰ���"|"�ַ�,�����3�к͵�13�ж�Ϊ���Ŀʱ��Ŀ�滻����
        mstrColumns = Replace(mstrColumns, "|" & intCol & "''1'", "|" & intCol & "'" & strCOLNames & "'1'" & strColFormat)
        '��
        mstrSQL�� = Replace(mstrSQL��, "'' AS C" & Format(intCol, "00"), strCOLDEF)
        '����
         '53893:������,2012-09-21,������Ŀ����ʱ���������
        'mstrSQL���� = Replace(UCase(mstrSQL����), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)
        mstrSQL���� = Replace(UCase(Replace(UCase(mstrSQL����), " OR """ & "C" & Format(intCol, "00") & """ IS NOT NULL", strCOLCOND)), """" & "C" & Format(intCol, "00") & """ IS NOT NULL", Mid(strCOLCOND, 5))
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
        'mstrSQL���� = Replace(UCase(mstrSQL����), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")
        mstrSQL���� = Replace(UCase(Replace(UCase(mstrSQL����), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")), """" & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL OR ", "")
        mstrSQL���� = Replace(UCase(mstrSQL����), "(""" & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL)", "")

        '��
        mstrSQL�� = Replace(mstrSQL��, ",MAX(""" & "C" & Format(arrData(intDo), "00") & """) AS C" & Format(arrData(intDo), "00"), "")
        '��
        mstrSQL�� = Replace(mstrSQL��, ", C" & Format(arrData(intDo), "00") & " AS C" & Format(arrData(intDo), "00"), "")
    Next
End Sub

Private Sub SQLCombination(ByVal str���� As String)
    mstrSQL = "Select  '' as ����,����ʱ��,����ʱ�� ����ʱ��1," & Mid(mstrSQL��, 12) & vbCrLf & _
                " From (Select ��¼���,ʱ�� as ����,����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select nvl(c.��¼���,0) ��¼���,to_char(l.����ʱ��,'yyyy-MM-dd hh24:mi:ss') AS ����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻����ļ� f,���˻����ӡ p " & vbCrLf & _
                "               Where l.ID=p.��¼ID And l.Id = c.��¼id And l.�ļ�ID+0=f.ID+0 And f.ID=p.�ļ�ID " & _
                "               And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] " & str���� & ")" & vbCrLf & _
                IIf(mstrSQL���� <> "", "Where " & mstrSQL����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ��ʱ��" & _
                                "       Order By ����ʱ��,��¼���,��ʿ,ǩ����,ǩ��ʱ��)"
End Sub

Private Sub zlReadTip(aryPeriod)
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String, strBed As String
    Dim strTmpSQL As String
    Dim strTmp As String
    Dim blnReplace As Boolean
    
    Err = 0: On Error GoTo ErrHand
    
    '���ϱ�ǩ��ȡ
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    
    '87057,����10:30:20ת����ס���������ļ�ʱ��Ϊ10:30:20,��ʱ¼����������Ϊ10:30(��¼���޷�¼����),�����޷���ʾ�µĿ���
    aryPeriod(0) = Format(aryPeriod(0), "YYYY-MM-DD HH:mm") & ":59"
    
    '��ȡ��ǰҳ֮ǰ��������ID
    gstrSQL = "Select ����ID From ���˱䶯��¼ " & _
        "   Where  ����ID=[1] And ��ҳID=[2] And [3]>=��ʼʱ�� " & _
        " And ��ʼʱ�� IS NOT NULL And ����id IS NOT NULL And NVL(���Ӵ�λ,0)=0 Order by ��ʼʱ�� DESC"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰҳ֮ǰ��������ID", mlng����ID, mlng��ҳID, CDate(aryPeriod(0)))
    If mrsTemp.RecordCount > 0 Then mlng����ID = Val(mrsTemp!����ID)
    
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as ��Ϣ From Dual"
    aryItem = Split(mstrSubhead, "|")
        
    For lngCount = 0 To UBound(aryItem)
        strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
        strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
        
        strTmp = strPrefix
        strCell = ""
        '68336
        blnReplace = True
        mrsElement.Filter = "������='" & strItemName & "'"
        If mrsElement.RecordCount > 0 Then
            blnReplace = Val(NVL(mrsElement!�滻��, 0)) = 1
        End If
        Select Case strItemName
        Case "��ǰ����"
        
            strTmpSQL = "Select   b.����" & vbNewLine & _
                        "From (Select ����id, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]��And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a,���ű� b " & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����id Is Not Null And b.ID=a.����id" & vbNewLine & _
                        "Order By a.��ʼʱ��"
                        
            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If mrsTemp.BOF = False Then mrsTemp.MoveLast
            
        Case "��ǰ����"

            strTmpSQL = "Select   a.����" & vbNewLine & _
                        "From (Select ����, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]��And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���� Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"

            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If mrsTemp.BOF = False Then mrsTemp.MoveLast
        Case "��λ�䶯"
            strTmpSQL = "Select   a.����" & vbNewLine & _
                        "From (Select ����, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3] And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a" & vbNewLine & _
                        "Where (a.��ֹʱ��>=[4] And a.��ʼʱ��<=[5]) And a.���� Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"

            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            strCell = "": strBed = ""
            Do While Not mrsTemp.EOF
                If strBed <> mrsTemp.Fields(0).Value Then
                    strBed = mrsTemp.Fields(0).Value
                    strCell = strCell & "->" & mrsTemp.Fields(0).Value
                End If
            mrsTemp.MoveNext
            Loop
            strCell = Mid(strCell, 3)
            If mrsTemp.RecordCount > 0 Then mrsTemp.MoveFirst
        Case "��ǰ����"
        
            strTmpSQL = "Select   ���� From ���ű� a Where a.ID=[1]"
            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID)
            
        Case "סԺҽʦ"
            strTmpSQL = "Select   a.����ҽʦ" & vbNewLine & _
                        "From (Select ����ҽʦ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]��And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ҽʦ Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "סԺҽʦ", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If mrsTemp.BOF = False Then mrsTemp.MoveLast
        Case "���λ�ʿ"
        
            strTmpSQL = "Select   a.���λ�ʿ" & vbNewLine & _
                        "From (Select ���λ�ʿ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]��And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���λ�ʿ Is Not Null" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "���λ�ʿ", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If mrsTemp.BOF = False Then mrsTemp.MoveLast
            
        Case "����ȼ�"
            strTmpSQL = "Select   b.����" & vbNewLine & _
                        "From (Select ����ȼ�ID, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                        "            From ���˱䶯��¼" & vbNewLine & _
                        "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]��And NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT NULL) a,����ȼ� b" & vbNewLine & _
                        "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ȼ�ID Is Not Null And b.���=a.����ȼ�ID" & vbNewLine & _
                        "Order By a.��ʼʱ��"
            Set mrsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "����ȼ�", mlng����ID, mlng��ҳID, mlng����ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
            If mrsTemp.BOF = False Then mrsTemp.MoveLast
            
        Case "������"
            strTmp = strPrefix
            gstrSQL = " Select f_List2str(Cast(Collect(Rownum || '��' || �������) As t_Strlist), ' ') As ������� from (Select ������� From ( Select  ������� || Decode(Nvl(�Ƿ�����, 0), 0, '', ' (��)') ������� ,Mod(�������, 10) �������, ���ʱ�� " & vbNewLine & _
                    "                   From ���˻������ C" & vbNewLine & _
                    "                      Where ����id = [1] And ��ҳid = [2] And �ļ�id = [3] And c.���ʱ�� Between [4] And [5])" & vbNewLine & _
                    "                       Group By ������� Order By Min(���ʱ��), Min(�������)) "
            Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���õ����", mlng����ID, mlng��ҳID, mlng��ǰ�ļ�ID, CDate(Format(aryPeriod(0), "YYYY-MM-DD hh:mm")), CDate(aryPeriod(1)))
            If NVL(mrsTemp!�������) = "" Then
                gstrSQL = " Select f_List2str(Cast(Collect(Rownum || '��' || �������) As t_Strlist), ' ') As ������� from (Select ������� From ( Select ������� || Decode(Nvl(�Ƿ�����, 0), 0, '', ' (��)') �������, �������,���ʱ��" & vbNewLine & _
                    "                                                  From (Select  Distinct �������,�������, ���ʱ��, �Ƿ�����," & vbNewLine & _
                    "                                                                Rank() Over(Partition By �ļ�id Order By ���ʱ�� Desc) As Top" & vbNewLine & _
                    "                                                         From (Select �ļ�id,�������, ���ʱ��, �������, �Ƿ�����" & vbNewLine & _
                    "                                                                From ���˻������ C" & vbNewLine & _
                    "                                                                Where ����id = [1] And ��ҳid = [2] And �ļ�ID = [3] And C.���ʱ�� < [4] ))" & vbNewLine & _
                    "                                                  Where Top = 1) Order By �������, ���ʱ��) "
                    
                Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���õ����", mlng����ID, mlng��ҳID, mlng��ǰ�ļ�ID, CDate(Format(aryPeriod(0), "YYYY-MM-DD hh:mm")), CDate(aryPeriod(1)))
            End If
            
            If NVL(mrsTemp!�������) = "" Then
                strTmp = ""
                gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as ��Ϣ From Dual"
                Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, mlng����ID, mlng��ҳID, mintӤ��, CDate(aryPeriod(0)))
            End If
        Case Else
            If blnReplace = True Then
                strTmp = ""
                gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as ��Ϣ From Dual"
                Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, mlng����ID, mlng��ҳID, mintӤ��, CDate(aryPeriod(0)))
            Else
                strTmp = strPrefix
                gstrSQL = "Select ���� From ���˻���Ҫ������ Where �ļ�ID=[1] And ҳ��=[2] And ����=[3]"
                Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", mlng��ǰ�ļ�ID, mintҳ��, strItemName)
            End If
        End Select
        
        If mrsTemp.BOF = False Then
            If strCell = "" Then
                If strTmp <> "" Then
                    lblSubhead.Tag = lblSubhead.Tag & " " & strTmp & mrsTemp.Fields(0).Value
                Else
                    lblSubhead.Tag = lblSubhead.Tag & " " & mrsTemp.Fields(0).Value
                End If
            Else
                lblSubhead.Tag = lblSubhead.Tag & " " & strTmp & strCell
            End If
        End If
    Next
    lblSubhead.Tag = Trim(lblSubhead.Tag)
    
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlRefresh(ByVal str���� As String)
    Dim aryRow() As String, aryItem() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    Dim strTmpSQL As String
    Dim strTmp As String
    
    Err = 0: On Error GoTo ErrHand
    
    'װ������
    Call SQLCombination(str����)
    gstrSQL = mstrSQL
    If gblnMoved Then
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng��ǰ�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ��, mint����ҳ)
    '���ݼ�¼��,���ڿ��ٻָ�
    Set mrsDataMap = CopyNewRec(mrsTemp, mrsDataMap)
    
    Exit Sub
ErrHand:
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
    Dim arrData
    Dim intData As Integer, intDatas As Integer
    Dim lngRowCount As Long, lngRowCurrent As Long  '��ǰ��¼������,��ǰ��¼�ڱ�ҳ��ʵ������
    Dim lngCol As Long, lngMax As Long
    Dim lngRow As Long, lngStart As Long, lngPrintedRow As Long, lngLastRow As Long
    Dim str����ʱ�� As String, str����ʱ��_L As String
    Dim blnDelete As Boolean
    Dim strSignName As String
    Dim blnClear As Boolean
    Dim blnCollectType As Boolean  '��¼���������е���һ���Ƿ��ǻ�����
    Dim lngCurrRow As Long, lngCollectMutilRows As Long '�������ݵ�ǰ�С����������ݵ�����
    Dim i As Integer, j As Integer, arrItem, arrCorrelative, arrLastRow, arrMutilRows '���������Ŀ����
    
    On Error GoTo ErrHand
    '���һ����ʾ�����������ʾ(���ݵ�ǰ����ռ����������ӿհ��в�����������,Ȼ�������δ���ǰ�е�����)
    'ÿҳֻ��ʾʵ�ʵ�������,��'@��ȡ��ע�ͼ���
    '���µ����������ݵ�ʵ����
    arrItem = Split(mstrColCorrelative, "|")
    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") <> 0 Then Exit Do
        
        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
        lngRowCurrent = Val(VsfData.TextMatrix(lngRow, mlngRowCurrent))
        
        str����ʱ�� = Format(VsfData.TextMatrix(lngRow, 1), "YYYY-MM-DD HH:mm:ss")
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) < 0 Then
            If blnCollectType = False Then str����ʱ��_L = "": blnCollectType = True
            '�������������ϸ���ݵĴ���(���ݱ��淽ʽ��һ���������ݶ�Ӧ������ϸ,��ϸ�еļ�¼��Ų�ͬ)
            If str����ʱ��_L <> "" And str����ʱ��_L = str����ʱ�� Then
                If UBound(arrItem) < 0 Then '�����ǰû�����û��ܹ�ϵ,��֮ǰ���ݴ��ڷ�����ܵ���������ӷ�������ѭ������
                    lngCurrRow = lngLastRow + lngCollectMutilRows 'ȷ��ÿһ���������������ʼλ��
                    lngCollectMutilRows = 1
                    If lngCurrRow < lngRow Then
                        VsfData.TextMatrix(lngCurrRow, mlngDate) = ""
                        VsfData.TextMatrix(lngCurrRow, mlngTime) = ""
                        
                        For lngCol = mlngTime + 1 To mlngNoEditor - 1
                            If (lngCol <> mlngSignTime And VsfData.ColHidden(lngCol) = False) Then
                                '׼����ֵ
                                With txtLength
                                    .Width = VsfData.ColWidth(lngCol)
                                    '������Ҫע��һ�㣺��ȡ�������ݵ�����Ӧ����lngRow������lngCurrRow����Ϊ�ڴ������������¼ʱ�ᵼ���������ݵ���λ�÷����仯 (���� = ����¼��ʼ�к� + ��������)
                                    .Text = Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                    .FontName = VsfData.CellFontName
                                    .FontSize = VsfData.CellFontSize
                                    .FontBold = VsfData.CellFontBold
                                    .FontItalic = VsfData.CellFontItalic
                                End With
                                arrData = GetData(txtLength.Text)
                                intDatas = UBound(arrData)
                                
                                If intDatas >= 0 Then
                                    'ѭ����ֵ
                                    If intDatas + 1 > lngRow - lngCurrRow Then intDatas = lngRow - lngCurrRow - 1
                                    If lngCollectMutilRows < intDatas + 1 Then lngCollectMutilRows = intDatas + 1
                                    For intData = 0 To intDatas
                                        VsfData.TextMatrix(lngCurrRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                    Next
                                End If
                            End If
                        Next lngCol
                    End If
                    lngLastRow = lngCurrRow
                Else
                    '�����˷�����ܹ�ϵ������ÿ��������Ŀ����չʾ����
                    For i = 0 To UBound(arrItem)
                        lngCurrRow = Val(arrLastRow(i)) + Val(arrMutilRows(i)) '����Ŀ����,ȷ��ÿ�������������ʼλ��
                        lngCollectMutilRows = 1
                        arrMutilRows(i) = lngCollectMutilRows
                        If lngCurrRow < lngRow Then
                            arrCorrelative = Split(arrItem(i), ";")
                            For j = 0 To 1
                                '׼����ֵ
                                    lngCol = Split(arrCorrelative(j), ",")(0) + cHideCols + VsfData.FixedCols - 1
                                    With txtLength
                                        .Width = VsfData.ColWidth(lngCol)
                                        '������Ҫע��һ�㣺��ȡ�������ݵ�����Ӧ����lngRow������lngCurrRow����Ϊ�ڴ������������¼ʱ�ᵼ���������ݵ���λ�÷����仯 (���� = ����¼��ʼ�к� + ��������)
                                        .Text = Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                        .FontName = VsfData.CellFontName
                                        .FontSize = VsfData.CellFontSize
                                        .FontBold = VsfData.CellFontBold
                                        .FontItalic = VsfData.CellFontItalic
                                    End With
                                    arrData = GetData(txtLength.Text)
                                    intDatas = UBound(arrData)
                                    
                                    If intDatas >= 0 Then
                                        If intDatas + 1 > lngRow - lngCurrRow Then intDatas = lngRow - lngCurrRow - 1
                                        If lngCollectMutilRows < intDatas + 1 Then lngCollectMutilRows = intDatas + 1
                                        arrMutilRows(i) = lngCollectMutilRows
                                        For intData = 0 To intDatas
                                            VsfData.TextMatrix(lngCurrRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                                        Next intData
                                    End If
                            Next j
                        End If
                        arrLastRow(i) = lngCurrRow
                    Next i
                End If
                '��ֵ��ɺ��Ƴ�ԭ��������
                VsfData.RowPosition(lngRow) = VsfData.Rows - 1
                VsfData.RemoveItem VsfData.Rows - 1
                GoTo NextData
            Else
                '������Ĭ��Ϊһ��(ֻ����Ի����е�����)
                lngCollectMutilRows = 1
                lngLastRow = lngRow '��¼������������е�λ��
                'ȷ��������������ӷ�������ÿ��������Ŀ����ʼλ��
                arrLastRow = Array(): arrMutilRows = Array()
                For i = 0 To UBound(arrItem)
                    ReDim Preserve arrLastRow(UBound(arrLastRow) + 1)
                    arrLastRow(UBound(arrLastRow)) = lngLastRow
                    ReDim Preserve arrMutilRows(UBound(arrMutilRows) + 1)
                    arrMutilRows(UBound(arrMutilRows)) = lngCollectMutilRows
                Next i
            End If
        Else
            If blnCollectType = True Then str����ʱ��_L = "": blnCollectType = False
            If str����ʱ��_L <> "" And Mid(str����ʱ��_L, 1, 16) = Mid(str����ʱ��, 1, 16) And str����ʱ��_L <> str����ʱ�� Then
                '������ͬ��������ͬ���Ҳ��ǻ��������У���˵����Щ������һ�飬����lngDemo��
                VsfData.TextMatrix(lngRow, mlngDate) = ""
                VsfData.TextMatrix(lngRow, mlngTime) = ""
                VsfData.TextMatrix(lngRow, mlngDemo) = lngRow - lngLastRow + 1
                If lngRow - lngLastRow = Val(VsfData.TextMatrix(lngLastRow, mlngRowCount)) Then
                    VsfData.TextMatrix(lngLastRow, mlngDemo) = 1
                End If
            Else
                lngLastRow = lngRow
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
            For lngCol = 0 To VsfData.Cols - 1
                If VsfData.ColHidden(lngCol) And lngCol <> mlngRowCount And lngCol <> mlngDemo Then
                    'ѭ����ֵ
                    For intData = 2 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = VsfData.TextMatrix(lngRow, lngCol)
                        '46506:������,2012-12-27,��ҳ��ӡ
                        '������ӡ������Ŵ��ڴ�ӡ��ҳ���ݺ�벿������
                        '��ӡ����ҳ��Ϊ��˵��֮ǰδʹ����ҳ��ӡ����ҳ�����Ѿ�ȫ����ӡ
                        '��ӡ��ҳ���ݺ�벿������Ϊ:��ӡҳ��+(��ǰ��+��ӡ�к�-1)\��ҳ������>��ӡ����ҳ��
                        If lngCol = mlngPrintedPage And gintPrintState = 1 And Val(VsfData.TextMatrix(lngRow, mlngRowCount)) <> Val(VsfData.TextMatrix(lngRow, mlngRowCurrent)) Then
                            If Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage)) <> 0 And Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) <> 0 And Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage)) >= Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) Then
                                If Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) + (intData + lngPrintedRow - 2) \ mlngPageRows > Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage)) Then
                                    VsfData.TextMatrix(lngRow + intData - 1, lngCol) = ""
                                End If
                            ElseIf Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage)) <> 0 And Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) = 0 Then
                                '�ϴ�ֻ��ӡ�˿�ҳ���ݿ�ҳ���ֵ�
                                '���磺��3ҳ���ݿ�ҳ����4ҳ��֮ǰֻ��ӡ�˵�4��ҳ�����ݡ��ٴ�����������������ӡ��3ҳ����ӡ��4ҳ
                                If Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage)) > mintҳ�� Then
                                    If Val(VsfData.TextMatrix(lngRow, mlngRowCurrent)) > intData Then
                                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage))
                                    End If
                                Else
                                    If Val(VsfData.TextMatrix(lngRow, mlngRowCount)) - Val(VsfData.TextMatrix(lngRow, mlngRowCurrent)) < intData Then
                                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = Val(VsfData.TextMatrix(lngRow, mlngPrintedEndPage))
                                    End If
                                End If
                            End If
                        End If
                    Next
                ElseIf (lngCol < mlngNoEditor And lngCol <> mlngDate And lngCol <> mlngTime) Then
                    '׼����ֵ
                    With txtLength
                        .Width = VsfData.ColWidth(lngCol)
                        .Text = Replace(Replace(Replace(VsfData.TextMatrix(lngRow, lngCol), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        .FontName = VsfData.CellFontName
                        .FontSize = VsfData.CellFontSize
                        .FontBold = VsfData.CellFontBold
                        .FontItalic = VsfData.CellFontItalic
                    End With
                    arrData = GetData(txtLength.Text)
                    intDatas = UBound(arrData)
                    
                    If intDatas > 0 Then
                        'ѭ����ֵ
                        If intDatas + 1 > lngRowCount Then intDatas = lngRowCount - 1
                        For intData = 0 To intDatas
                            If VsfData.Rows <= lngRow + intData Then VsfData.Rows = VsfData.Rows + 1
                            VsfData.TextMatrix(lngRow + intData, lngCol) = Replace(Replace(Replace(arrData(intData), Chr(10), ""), Chr(13), ""), Chr(1), "")
                        Next
                    End If
                ElseIf lngCol = mlngNoEditor Then
                    '����ֵ��Ϊ��1��ʼ,������4������,����4|1
                    For intData = 1 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                    Next
                    '���һ����Ҫ��д���ǩ��
                    If mlngSignName > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignName) = VsfData.TextMatrix(lngRow, mlngSignName)
                    If mlngSignTime > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignTime) = VsfData.TextMatrix(lngRow, mlngSignTime)
                    '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
                    Call SingerShowType(VsfData, lngRow, lngRow + lngRowCount - 1)
                Else
                
                End If
            Next
            lngRow = lngRow + lngRowCount - 1
        Else
            VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
        End If
        lngRow = lngRow + 1
NextData:
        str����ʱ��_L = str����ʱ��
    Loop
    
    '���ÿҳ��������
    lngRow = VsfData.FixedRows
    
    Do While True
        '�̶�������ʾ����ʱ����ǩ����
        lngStart = GetStartRow(lngRow)
        
        '���⴦���һ��(��һ�п��ܴ��ڿ�ҳ����)
        '50503:������,2012-09-12,ֻ�п�ʼ�к�<>1�ĲŽ��д������⴦��������ݴӼ�¼��ĳҳ��һ�оͿ�ҳ������
        If lngRow = VsfData.FixedRows And Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngRow, mlngRowCurrent)) And Val(VsfData.TextMatrix(lngRow, mlngStartSpread)) > 1 Then
            blnDelete = True
            lngRow = lngRow + Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)) - Val(VsfData.TextMatrix(lngRow, mlngRowCurrent))
        End If
        
        If lngStart <> lngRow Or (Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 1 And lngStart = lngRow) Then
            '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 1 Then
                For lngRowCount = lngStart To VsfData.FixedRows Step -1
                    If Val(VsfData.TextMatrix(lngRowCount, mlngDemo)) = 1 Then
                        If mlngDate > -1 Then VsfData.TextMatrix(lngRow, mlngDate) = VsfData.TextMatrix(lngRowCount, mlngDate)
                        If mlngTime > -1 Then VsfData.TextMatrix(lngRow, mlngTime) = VsfData.TextMatrix(lngRowCount, mlngTime)
                        Exit For
                    End If
                Next
            Else
                If mlngDate > -1 Then VsfData.TextMatrix(lngRow, mlngDate) = VsfData.TextMatrix(lngStart, mlngDate)
                If mlngTime > -1 Then VsfData.TextMatrix(lngRow, mlngTime) = VsfData.TextMatrix(lngStart, mlngTime)
            End If
            If lngStart <> lngRow Then
                '65994:������,2013-09-26,����β��ǩ��,������ݿ�ҳ��ᵼ����ҳ��û��ǩ����ֻ�еڶ�ҳ��ǩ��
                If mlngSingerType <> 3 Then '��β��ǩ������д��ҳ�����ڵڶ�ҳ���ݵ���ʼ��
                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow, mlngOperator) = VsfData.TextMatrix(lngStart, mlngOperator)
                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow, mlngSignName) = VsfData.TextMatrix(lngStart, mlngSignName)
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow, mlngSignTime) = VsfData.TextMatrix(lngStart, mlngSignTime)
                Else 'β��ǩ��,��д��ҳ������ʼҳ����ʼ��(ӦΪ��ʼ�������ʱ,��ʼ���Ѿ����)
                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngStart, mlngOperator) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngOperator)
                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngStart, mlngSignName) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngSignName)
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngStart, mlngSignTime) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngSignTime)
                End If
                '���µ�ǰҳ���һ�У���ҳ���ݵ�ǩ���˺�ǩ��ʱ��
                If lngStart <> lngRow - 1 Then
                    '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngOperator) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngOperator)
                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignName) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngSignName)
                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignTime) = VsfData.TextMatrix(lngStart + Val(VsfData.TextMatrix(lngStart, mlngRowCount)) - 1, mlngSignTime)
                    Call SingerShowType(VsfData, lngStart, lngRow - 1)
                End If
            End If
        End If
        
        If blnDelete Then
            '89208:��֮ǰ���µ��������������ڽ���ɾ��(�Ա㴦�����63760�����е�ǩ������ʾ)
            lngRowCount = Val(VsfData.TextMatrix(lngStart, mlngRowCount))
            lngRowCount = lngRowCount - (lngRow - lngStart)
            If lngRowCount > 0 Then
                VsfData.TextMatrix(lngRow, mlngDemo) = VsfData.TextMatrix(lngStart, mlngDemo)
                For lngCol = 0 To lngRowCount - 1
                    VsfData.TextMatrix(lngCol + lngRow, mlngRowCount) = lngRowCount & "|" & lngCol + 1
                    VsfData.TextMatrix(lngCol + lngRow, mlngRowCurrent) = lngRowCount
                Next
            End If
            
            For lngCol = lngStart To lngRow - 1
                VsfData.RemoveItem lngStart
            Next
            
            blnDelete = False
            lngRow = VsfData.FixedRows  'ֻ�����һ�м�¼ɾ�������,���Թ̶�����Ϊ�̶���Ϊ��ʼ��
        End If
        
        lngRow = lngRow + mlngPageRows
        If lngRow > VsfData.Rows - 1 Then Exit Do
    Loop
    
    '63760:������,�������ݻ�ʿ��ǩ���ˡ�ǩ��ʱ��Ĵ���ͬһ��ǩ����ʼ����ʾһ�Σ�
    If mlngSingerType > 0 And VsfData.FixedRows <= VsfData.Rows - 1 Then
        lngPrintedRow = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        lngRow = VsfData.FixedRows
        Do While True
            lngStart = GetStartRow(lngRow)
            lngRowCount = Val(VsfData.TextMatrix(lngStart, mlngRowCount))
            If lngRowCount <= 0 Then Exit Do
            
            If mlngSingerType = 3 Then 'β��ǩ��
                strSignName = VsfData.TextMatrix(lngStart + lngRowCount - 1, lngPrintedRow)
            Else '����ǩ������βǩ��
                strSignName = VsfData.TextMatrix(lngStart, lngPrintedRow)
            End If
            strSignName = FormatValue(strSignName)
            '����Ƿ��Ƿ������ݣ��ӷ�����ʼ�п�ʼ����
            If Val(VsfData.TextMatrix(lngStart, mlngDemo)) = 1 And lngStart = lngRow And strSignName <> "" Then
                For lngRow = lngStart + lngRowCount To VsfData.Rows - 1
                    If lngRow = lngStart + lngRowCount Then
                    
                        If Val(VsfData.TextMatrix(lngRow, mlngDemo)) <= 1 Then Exit For
                        
                        '���ͬһ���鲻ͬ����֮��Ļ�ʿ��ǩ�����Ƿ���ͬ��������Ӧ�Ĵ���
                        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
                        If lngRowCount = 0 Then Exit For
                        
                        If mlngSingerType = 3 Then 'β��ǩ��
                            '��ʿ��ǩ������ͬ��ֻ�ڱ��������һ���������һ����ʾ��ʿ��ǩ����
                            If strSignName = FormatValue(VsfData.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow)) Then
                                '�������ı������ݸպ���ĳһҳ�������������һҳ���������һ�����ݵĻ�ʿ��ǩ����
                                If (lngRow - VsfData.FixedRows) Mod mlngPageRows > 0 And lngStart <= lngRow - 1 Then
                                    If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngOperator) = ""
                                    If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignName) = ""
                                    If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignTime) = ""
                                End If
                            Else
                                If FormatValue(VsfData.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow)) <> "" Then
                                    strSignName = FormatValue(VsfData.TextMatrix(lngRow + lngRowCount - 1, lngPrintedRow))
                                End If
                            End If
                        Else '����ǩ������βǩ��
                            If strSignName = FormatValue(VsfData.TextMatrix(lngRow, lngPrintedRow)) Then
                                '�������ı������ݸպ���ĳһҳ�������������һҳ���������һ�����ݵĻ�ʿ��ǩ����
                                If (lngRow - VsfData.FixedRows) Mod mlngPageRows > 0 Then
                                    blnClear = True
                                    '��βǩ����Ҫע�⣺�������ĳ�����ݵ�����(����ʼ������)��ĳҳ�����һ�У���ȡ����ʿǩ���˵���ʾ
                                    If mlngSingerType = 2 And lngRowCount = 1 Then
                                        If lngRow + lngRowCount < VsfData.Rows Then
                                            If Val(VsfData.TextMatrix(lngRow + lngRowCount, mlngDemo)) <= 1 Then
                                                blnClear = False
                                            End If
                                        Else
                                            blnClear = False
                                        End If
                                    End If
                                    
                                    If blnClear = True Then
                                        If mlngSingerType = 1 Or (mlngSingerType = 2 And (lngRow + 1 - VsfData.FixedRows) Mod mlngPageRows > 0) Then
                                            blnClear = False
                                            If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow, mlngOperator) = ""
                                            If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow, mlngSignName) = ""
                                            If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow, mlngSignTime) = ""
                                        End If
                                    End If
                                    
                                    If mlngSingerType = 2 And lngStart < lngRow - 1 Then '��βǩ����Ӧ��ȥ����һ�����ݵ�β��(��һ������������Ҫ>1)
                                        '�����һ�����ݿ�ҳ�������ڵ�ǰҳֻ��һ���򲻽�������������ݵ����һ��ǩ������ʿ
                                        If (lngRow - 1 - VsfData.FixedRows) Mod mlngPageRows > 0 Then
                                            If mlngOperator <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngOperator) = ""
                                            If mlngSignName <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignName) = ""
                                            If mlngSignTime <> -1 Then VsfData.TextMatrix(lngRow - 1, mlngSignTime) = ""
                                        End If
                                    End If
                                End If
                            Else
                                If FormatValue(VsfData.TextMatrix(lngRow, lngPrintedRow)) <> "" Then
                                    strSignName = FormatValue(VsfData.TextMatrix(lngRow, lngPrintedRow))
                                End If
                            End If
                        End If
                        
                        lngStart = lngRow
                    End If
                Next lngRow
            Else
                lngRow = lngStart + lngRowCount
            End If
            
            If lngRow > VsfData.Rows - 1 Then Exit Do
        Loop
    End If
    
    '������ش�,������ҳ��Ч�����еĲ���ɾ��
    If gintPrintState = 2 Then
        If VsfData.Rows > VsfData.FixedRows + mlngPageRows Then
            VsfData.Rows = VsfData.FixedRows + mlngPageRows
        End If
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat()
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    Dim blnAlign As Boolean
    
    On Error GoTo ErrHand
    
    '���û����¼���ĸ�ʽ
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = mrsDataMap
        
        '��ͷ��д
        .MergeCells = flexMergeFixedOnly ' = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        '������кϲ����ԣ�.Clear�������֮ǰ�ĺϲ���Ϣ
        For lngCount = .FixedRows To .Rows - 1
            .MergeRow(lngCount) = False
        Next
        '�����ڲ�����������
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngStartSpread) = True
        '51589:������,2013-03-01,��ӽ���ǩ��
        .ColHidden(mlngJoinSignName) = True
        .ColHidden(mlngFileID) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngCollectStyle) = True
        .ColHidden(mlngCollectText) = True
        .ColHidden(mlngCollectType) = True
        .ColHidden(mlngCollectDay) = True
        .ColHidden(mlngPrintedPage) = True
        .ColHidden(mlngPrintedRow) = True
        .ColHidden(mlngPrintedTag) = True
        .ColHidden(mlngPrintedEndPage) = True
        .ColHidden(mlngCollectValue) = True
        '������ͷ
        Dim strCOL As String
        Dim dblWidth As Double
        
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        Call PreActiveHead
        
        '�п�����
        blnAlign = False
        aryItem = Split(mstrColWidth, ",")
        If mbln����ʱ��ϲ� Then strCOL = "," & mlngDate & "," & mlngTime & ","
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0))
                If mbln����ʱ��ϲ� And InStr(1, strCOL, "," & lngCount & ",") > 0 Then
                    dblWidth = dblWidth + .ColWidth(lngCount)
                End If
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        '������ʱ������ʾ����,�п�Ϊ������ʱ���е��ܿ��
        If mbln����ʱ��ϲ� Then
            .ColWidth(2) = IIf(dblWidth < 1600, 1600, dblWidth)
            .TextMatrix(0, 2) = "����ʱ��"
            If mintTabTiers >= 2 Then .TextMatrix(1, 2) = "����ʱ��"
            If mintTabTiers >= 3 Then .TextMatrix(2, 2) = "����ʱ��"
            .ColAlignment(2) = .ColAlignment(mlngDate)
        End If
        
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
        If mbln����ʱ��ϲ� Then
            '��PreTendMutilRows()��Ҫ��������,���Ա��뽫�е��������Է�����������
            .ColHidden(mlngDate) = True
            .ColHidden(mlngTime) = True
            .ColHidden(2) = False
        End If
        
        If mblnʱ�������� = True Then .ColHidden(mlngTime) = True
        
        Call WriteColor
        
        '���̶ܹ��е��и߲���ȷ��Ҫ�Զ�������
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        '���ǹ̶��е��и�����Ϊ��С�и�
        For lngCount = 0 To .FixedRows - 1
            If .RowHeight(lngCount) < .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        For lngCount = .FixedRows To .Rows - 1
            .RowHeight(lngCount) = .RowHeightMin
        Next
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function PrintHead() As Boolean
    Dim lngPage As Long
    On Error GoTo ErrHand
    
    lngPage = mintҳ��
    mlng��ǰҳ�� = lngPage
    PrintHead = PrintRTBData(rtbHead, True, lngPage)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function PrintFoot() As Boolean
    Dim lngPage As Long
    On Error GoTo ErrHand
    
    lngPage = mintҳ��
    mlng��ǰҳ�� = lngPage
    PrintFoot = PrintRTBData(rtbFoot, False, lngPage)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function PrintRTBData(ByVal objRTB As RichTextBox, ByVal blnHead As Boolean, Optional ByVal lngPage As Long = 0) As Boolean
    Dim fr As FORMATRANGE           '��ʽ�����ı���Χ
    Dim rcDrawTo As RECT            'Ŀ����������
    Dim rcPage As RECT              'Ŀ��ҳ������
    Dim rcHeand As RECT
    Dim rcFoot As RECT
    Dim gTargetDC As Long
    Dim lngFoot As Long
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
    Dim lngNextPos As Long, lngLen As Long, lngTmp As Long, lngPageCount As Long
    Dim mrsTemp As New ADODB.Recordset
    Dim lngPageIndex As Long, lngPrintTextY As Long
    
    lngLen = lstrlen(objRTB.Text)
    lngOffsetLeft = gobjOutTo.ScaleX(GetDeviceCaps(gobjOutTo.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = gobjOutTo.ScaleY(GetDeviceCaps(gobjOutTo.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
    
    '46251,������,2012-09-11,װ��ҳ�����λ��
    lngPageIndex = Val(cboҳ��.ItemData(cboҳ��.ListIndex))
    If lngPageIndex <= 0 Or lngPageIndex > 4 Then lngPageIndex = 4
    If blnHead Then
        If chkҳ��.Value = 1 And (lngPageIndex = 1 Or lngPageIndex = 2) Then
            lngFoot = gobjOutTo.TextHeight("��")
            With rcHeand
                .Left = lngOffsetLeft
                .Right = gobjOutTo.Width - lngOffsetLeft
                If lngPageIndex = 1 Then
                    lngPrintTextY = lngOffsetTop + 30
                    .Top = lngOffsetTop + lngFoot + 60
                    .Bottom = gobjOutTo.ScaleX(gobjSend.EmptyUp, vbMillimeters, vbTwips)
                Else
                    .Top = lngOffsetTop
                    .Bottom = gobjOutTo.ScaleX(gobjSend.EmptyUp, vbMillimeters, vbTwips) - lngFoot - 60
                    lngPrintTextY = .Bottom + 30
                End If
                If lngPrintTextY < lngOffsetTop + 30 Then lngPrintTextY = lngOffsetTop + 30
            End With
        Else
            With rcHeand
                .Left = lngOffsetLeft
                .Top = lngOffsetTop
                .Right = gobjOutTo.Width - lngOffsetLeft
                .Bottom = gobjOutTo.ScaleX(gobjSend.EmptyUp, vbMillimeters, vbTwips) - 30
            End With
            gobjOutTo.Print ""
        End If
    Else
        '62436:������,2013-06-20,�޸�ҳ��������꣬��֤�ڴ�ӡ���ɴ�ӡ�����ڡ�
        If chkҳ��.Value = 1 And (lngPageIndex = 3 Or lngPageIndex = 4) Then
            If lngPageIndex = 3 Then
                lngFoot = gobjOutTo.TextHeight("��") + 60
                lngPrintTextY = gobjOutTo.Height - gobjOutTo.ScaleY(gobjSend.EmptyDown, vbMillimeters, vbTwips) - lngOffsetTop * 2
                rcFoot.Bottom = gobjOutTo.Height
            Else
                lngFoot = gobjOutTo.TextHeight("��") + 60
                lngPrintTextY = gobjOutTo.Height - lngOffsetTop * 2 - lngFoot
                rcFoot.Bottom = lngPrintTextY
            End If
            If lngPrintTextY + lngFoot > gobjOutTo.Height - lngOffsetTop * 2 Then lngPrintTextY = gobjOutTo.Height - lngOffsetTop * 2 - lngFoot
            If lngPageIndex = 4 Then lngFoot = 0
        Else
            gobjOutTo.Print ""
            lngFoot = 0
            rcFoot.Bottom = gobjOutTo.Height
        End If
        With rcFoot
            .Left = lngOffsetLeft
            .Top = gobjOutTo.Height - lngOffsetTop * 2 - gobjOutTo.ScaleY(gobjSend.EmptyDown, vbMillimeters, vbTwips) + lngFoot
            .Right = gobjOutTo.Width - lngOffsetLeft
        End With
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
            .Left = rcHeand.Left
            .Top = rcHeand.Top
            .Right = rcHeand.Right
            .Bottom = rcHeand.Bottom
        Else
            .Left = rcFoot.Left
            .Top = rcFoot.Top
            .Right = rcFoot.Right
            .Bottom = rcFoot.Bottom
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
        lngNextPos = SendMessage(objRTB.hWnd, EM_FORMATRANGE, 0, fr)
        
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
    Call SendMessage(objRTB.hWnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    
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
        Call SendMessage(objRTB.hWnd, EM_FORMATRANGE, 1, fr)
        Call SendMessage(objRTB.hWnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    Next
    
    '��������ҳ��
    If chkҳ��.Value = 1 And ((blnHead = True And (lngPageIndex = 1 Or lngPageIndex = 2)) Or (blnHead = False And (lngPageIndex = 3 Or lngPageIndex = 4))) Then
        gobjOutTo.CurrentY = lngPrintTextY
        If optPageAlign(0).Value Then
            gobjOutTo.CurrentX = gobjOutTo.ScaleX(gobjSend.EmptyLeft, vbMillimeters, vbTwips) - 30
        ElseIf optPageAlign(1).Value Then
            gobjOutTo.CurrentX = (gobjOutTo.Width - 90 * LenB(StrConv("�� " & lngPage & " ҳ", vbFromUnicode))) / 2
        Else
            gobjOutTo.CurrentX = gobjOutTo.Width - gobjOutTo.ScaleX(gobjSend.EmptyRight, vbMillimeters, vbTwips) - 90 * LenB(StrConv("ҳ��:" & mintҳ��, vbFromUnicode))
        End If
        gobjOutTo.Print "�� " & lngPage & " ҳ"
    End If
End Function

Public Function PrintPage(Optional blnOddEvenPrint As Boolean = False, Optional ArrSQL As Variant) As Boolean
    Dim strSQL() As String
    Dim blnTrans As Boolean
    Dim blnSave As Boolean          '�Ѵ�ӡ�����ݲ�����
    Dim strTime As String
    Dim strCurrDate As String
    Dim lngRow As Long, lngRows As Long
    Dim intMax As Integer, intPos As Integer
    Dim lngCurRow As Long, lngDataLines As Long
    Dim intTag As Integer
    Dim int����ҳ�� As Integer
    Dim lngFileID As Long
    
    ReDim Preserve strSQL(1 To 1)
    On Error GoTo ErrHand
    
    '56134:������,2012-12-19,���˻����ӡ��Ӵ�ӡ��ʶ
    If mblnPrintRow = True Then
        intTag = 1
    Else
        intTag = 0
        If gintPrintState = 1 And glngPrintRow > 0 And Val(VsfData.TextMatrix(glngPrintRow, VsfData.Cols - 7)) > 0 Then
            intTag = 1
        End If
    End If
    
    '����ʾ�н��д���
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Not VsfData.RowHidden(lngRow) Then
            If lngCurRow = 0 Then lngCurRow = 1
            If FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) Like "*|1" Then
                lngDataLines = Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0)
                blnSave = (Val(VsfData.TextMatrix(lngRow, mlngPrintedPage)) = 0) Or gintPrintState > 1
                '�ش�Ļ�����ԭ�еĽ���ҳ��
                int����ҳ�� = IIf(mlng��ǰҳ�� = 0, mintҳ��, mlng��ǰҳ��)
                If blnSave Then
                    strTime = VsfData.TextMatrix(lngRow, 1)
                    lngFileID = Val(VsfData.TextMatrix(lngRow, mlngFileID))
                    gstrSQL = "ZL_���˻����ӡ_PRINT(" & lngFileID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'),'" & gstrUserName & "'," & _
                        IIf(mlng��ǰҳ�� = 0, mintҳ��, mlng��ǰҳ��) & "," & lngCurRow & "," & intTag & "," & int����ҳ�� & ")"
                    'Debug.Print gstrSQL
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                End If
            End If
            '46506:������,2012-12-28,��¼����ҳ��ӡ
            'ÿһҳ������������Ǳ������ݵ���ʼ��,��˵���ǿ�ҳ����
            If lngCurRow = 1 And Not FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) Like "*|1" Then
                strTime = VsfData.TextMatrix(lngRow, 1)
                lngFileID = Val(VsfData.TextMatrix(lngRow, mlngFileID))
                gstrSQL = "ZL_���˻����ӡ_PRINT(" & lngFileID & ",to_date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'),'" & gstrUserName & "'," & _
                         "NULL,NULL," & intTag & "," & IIf(mlng��ǰҳ�� = 0, mintҳ��, mlng��ǰҳ��) & ",1)"
                strSQL(ReDimArray(strSQL)) = gstrSQL
            End If
            lngCurRow = lngCurRow + 1
        End If
    Next
    
    '�������ż��ӡ�����Ҵ�ӡҳ����Ϊ1,�ͷ�������SQL
    If blnOddEvenPrint = True Then
        ArrSQL = strSQL
        PrintPage = True
        Exit Function
    End If
    
    On Error Resume Next
    intMax = UBound(strSQL)

    gcnOracle.BeginTrans
    blnTrans = True

    On Error GoTo ErrHand
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
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitRecords()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim lngCol As Long, lngOrder As Long, strName As String, intImmovable As Integer, strFormat As String
    Dim arrColumn, arrItem, arrCorrelative(), strColumns As String
    Dim blnSet As Boolean
    
    On Error GoTo ErrHand
    
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
            lngCol = Split(arrColumn(i), "'")(0)
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
                strValues = lngCol & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, strFields, strValues)
            Next
        Next
        
        '���������ܹ�������Ϣ
        arrCorrelative = Array()
        arrColumn = Split(mstrColCorrelative, "|")
        For i = 0 To UBound(arrColumn)
            arrItem = Split(arrColumn(i), ";")
            If UBound(arrItem) = 1 Then
                mrsSelItems.Filter = "��=" & Val(arrItem(0))
                If mrsSelItems.RecordCount = 1 Then
                    ReDim Preserve arrCorrelative(UBound(arrCorrelative) + 1)
                    arrCorrelative(UBound(arrCorrelative)) = Val(arrItem(0)) & "," & mrsSelItems!��Ŀ��� & ";" & CStr(arrItem(1))
                End If
            End If
        Next i
        If UBound(arrCorrelative) = -1 Then
            mstrColCorrelative = ""
        Else
            mstrColCorrelative = Join(arrCorrelative, "|")
        End If
        mrsSelItems.Filter = ""
        
        'Call OutputRsData(mrsSelItems)
        
        '��������ڲ�������(�����ڶ�ȡ���ݺ��ʱ���ӵ�,��ʱֻ��Ԥ������)
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '����������
        mlngSigner = mlngSignLevel + 1
        '51589:������,2013-03-01,��ӽ���ǩ��
        mlngJoinSignName = mlngSigner + 1
        mlngFileID = mlngJoinSignName + 1
        mlngRecord = mlngFileID + 1
        mlngRowCount = mlngRecord + 1
        mlngRowCurrent = mlngRowCount + 1
        mlngStartSpread = mlngRowCurrent + 1 '50503:������,2012-09-12
        mlngPrintedEndPage = mlngStartSpread + 1 '46506:������,2012-12-27
        mlngCollectValue = mlngPrintedEndPage + 1  '����,105302
        mlngPrintedTag = mlngCollectValue + 1 '56134:������,2012-12-19
        mlngCollectType = mlngPrintedTag + 1
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
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ShowPage(Optional ByVal intPage As Integer = 0) As Boolean
    '��ʾָ��ҳ�����ݲ����´�ӡ����
    Dim aryPeriod
    Dim strBegin As String, strEnd As String
    Dim lngRow As Long, lngRows As Long, lngStart As Long
    Dim lngOffsetLeft As Long, lngScaleWidth As Long
    Dim lngShows As Long
    '���Ŀ��ر���
    Dim mrsTemp As New ADODB.Recordset
    Dim blnPrintRow As Boolean
    On Error GoTo ErrHand
    
    If intPage <> 0 Then mlngMinIndex = intPage - Val(mArrPage(0))
    If mlngMinIndex > mlngMaxIndex Then mlngMinIndex = mlngMaxIndex
    
    mintҳ�� = Val(mArrPage(mlngMinIndex))
    If InStr(1, CStr(mArrPage(mlngMinIndex)), ";") <> 0 Then
        gintPrintState = Val(Split(CStr(mArrPage(mlngMinIndex)), ";")(1))
    Else
        gintPrintState = 2
    End If
    
    lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    
    Call LoadPageData '������Ӧҳ��ӡ����
    
    With VsfData
        'С��ҳ����Ч������˵��ֻ��һҳ����
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            lngRows = .Rows - 1
            For lngRow = .FixedRows To lngRows
                .RowHidden(lngRow) = True
            Next
        End If
        
        'С��ҳ����Ч������˵��ֻ��һҳ����
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            lngRow = 3 + mlngPageRows * (mintҳ�� - mint��ǰ��ʼҳ)
            lngRows = 3 + mlngPageRows * (mintҳ�� - mint��ǰ��ʼҳ + 1) - 1
        Else
            lngRow = 3
            lngRows = .Rows - 1
        End If
        If lngRows > .Rows - 1 Then lngRows = .Rows - 1
        '��ȡָ��ҳ��ʱ�䷶Χ
        If lngRow > lngRows Then
            Exit Function
        End If
        strBegin = Format(.TextMatrix(lngRow, 1), "YYYY-MM-DD HH:mm:ss")
        lngStart = lngRows
        lngStart = GetStartRow(lngStart)
        strEnd = .TextMatrix(lngStart, 1)
        If Not IsDate(strEnd) And lngStart <> lngRow Then
            lngStart = lngRow
            strEnd = .TextMatrix(lngStart, 1)
        End If
        strEnd = Format(strEnd, "YYYY-MM-DD HH:mm") & ":59"
        '53588:������,2013-4-25,�޸����ݵ�ʱ��С�ڲ�����Ժʱ�䣬���ţ�����������ʾ����
        '�磺�������ʱ��Ϊ2013-03-13 11:23:34 �ļ���ʼʱ��������ͬ����ʱ¼������ʱ��Ϊ 2013-03-13 11:23
        '�ͻᵼ���޷���ȡ���ţ�ӦΪ���������ʱ��Ϊ2013-03-13 11:23:00 С���˲������ʱ�䵼���޷���ȡ������
        '��ȡ���˵���Ժʱ��
        If mintӤ�� = 0 Then
            gstrSQL = "Select ��ʼʱ��, Sysdate As ����ʱ��" & vbNewLine & _
                " From ���˱䶯��¼" & vbNewLine & _
                " Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 2" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select ��ʼʱ��, Sysdate As ����ʱ��" & vbNewLine & _
                " From ���˱䶯��¼ a" & vbNewLine & _
                " Where a.����id = [1] And a.��ҳid = [2] And a.��ʼԭ�� = 1 And Not Exists" & vbNewLine & _
                " (Select 1 From ���˱䶯��¼ Where ����id = a.����id And ��ҳid = a.��ҳid And ��ʼԭ�� = 2)"
        Else
            gstrSQL = " Select   ����ʱ�� AS ��ʼʱ�� From ������������¼ Where ����ID=[1] And ��ҳID=[2] And ���=[3]"
        End If
        Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ���ڻ��������", mlng����ID, mlng��ҳID, mintӤ��)
        If Format(strBegin, "yyyy-MM-dd HH:mm:ss") < Format(mrsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") Then
            strBegin = Format(mrsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
        End If
        aryPeriod = Split(strBegin & "||" & strEnd, "||")
        
        lngStart = lngRow
        'С��ҳ����Ч������˵��ֻ��һҳ����
        If VsfData.Rows - VsfData.FixedRows > mlngPageRows Then
            '��ʾ������
            For lngRow = lngRow To lngRows
                .RowHidden(lngRow) = False
                lngShows = lngShows + 1
            Next
        End If
        
        '56134:������,2012-12-19
        '����һҳ�����ʣ����(ֻ�����һҳ�Ż��д����)
        '�ش�������ش�������ֻҪmblnPrintRow=True���������ݲ���һҳ������Ҫ������
        '������ӡʱmblnPrintRow=True���������ݲ���һҳ�������ҳδ��ӡ������������Ѿ���ӡ��˵��֮ǰ�Ѿ����������ˡ�
        'Ԥ���������ֻҪmblnPrintRow=True��������
        If mblnPrintRow = True And lngRows - lngStart + 1 < mlngPageRows Then
            blnPrintRow = False
            If gblnPrintMode = False Then 'Ԥ��
                blnPrintRow = True
            Else '��ӡ
                If gintPrintState = 1 Then '������ӡ
                    '�����ҳ֮ǰû�д�ӡ������������
                    If glngPrintRow >= lngStart And glngPrintRow <= lngRows Then
                        blnPrintRow = (Val(.TextMatrix(glngPrintRow, mlngPrintedTag)) = 0)
                    Else
                        blnPrintRow = True
                    End If
                Else '�ش�������ش�
                    blnPrintRow = True
                End If
            End If
            If blnPrintRow = True Then
                VsfData.Rows = VsfData.Rows + mlngPageRows - (lngRows - lngStart + 1)
                For lngRow = lngRows To VsfData.Rows - 1
                    VsfData.RowHeight(lngRow) = VsfData.RowHeightMin
                Next
            End If
        End If
        
        ShowPage = True
        Call zlReadTip(aryPeriod)
    End With
    
    '���ô�ӡ�������
    Dim objPrint As New zlTFPrintTends, objAppRow As zlTFTabAppRow
    Dim strLable As String, strAppRow As String, lngSpaces As Long
    Dim lngPos As Long, lngMax As Long, lngNumber As Long, blnNumber As Boolean, lngASC As Long
    
    '���ô�ӡ��ʽ
    If UBound(Split(mstrPaperSet, ";")) >= 0 Then SaveSetting "ZLSOFT", "����ģ��\zl9TendFile\Default", "PaperSize", Val(Split(mstrPaperSet, ";")(0))
    If UBound(Split(mstrPaperSet, ";")) >= 1 Then SaveSetting "ZLSOFT", "����ģ��\zl9TendFile\Default", "Orientation", Val(Split(mstrPaperSet, ";")(1))
    If UBound(Split(mstrPaperSet, ";")) >= 2 Then SaveSetting "ZLSOFT", "����ģ��\zl9TendFile\Default", "Height", Val(Split(mstrPaperSet, ";")(2))
    If UBound(Split(mstrPaperSet, ";")) >= 3 Then SaveSetting "ZLSOFT", "����ģ��\zl9TendFile\Default", "Width", Val(Split(mstrPaperSet, ";")(3))
    If UBound(Split(mstrPaperSet, ";")) >= 4 Then objPrint.EmptyLeft = Round(ScaleY(Val(Split(mstrPaperSet, ";")(4)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 5 Then objPrint.EmptyRight = Round(ScaleY(Val(Split(mstrPaperSet, ";")(5)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 6 Then objPrint.EmptyUp = Round(ScaleX(Val(Split(mstrPaperSet, ";")(6)), vbTwips, vbMillimeters), 2)
    If UBound(Split(mstrPaperSet, ";")) >= 7 Then objPrint.EmptyDown = Round(ScaleX(Val(Split(mstrPaperSet, ";")(7)), vbTwips, vbMillimeters), 2)
    
    On Error Resume Next
    Printer.PaperSize = Val(Split(mstrPaperSet, ";")(0))
    Printer.Orientation = Val(Split(mstrPaperSet, ";")(1))
    
    If Printer.PaperSize = 256 Then
        Call SetCustonPager(Val(Split(mstrPaperSet, ";")(3)), Val(Split(mstrPaperSet, ";")(2)))
    End If
    
    On Error GoTo ErrHand
    Set objPrint.Body = VsfData
    objPrint.Title.Text = lblTitle.Caption
    Set objPrint.Title.Font = lblTitle.Font
    Set objPrint.AppFont = lblSubhead.Font
    
    lngSpaces = lblSubhead.Height / 210
    strLable = lblSubhead.Caption
    '60333:������,2013-10-14,�޸��ļ���ŵ��´�ӡ��Ϣ��ʧ
    If UBound(Split(mstrPaperSet, ";")) >= 4 Then
        lngScaleWidth = Printer.Width - (lngOffsetLeft + Val(Split(mstrPaperSet, ";")(4))) * 2
    Else
        lngScaleWidth = Printer.Width - (lngOffsetLeft) * 2
    End If
    lngMax = Len(strLable)
    lngNumber = 0
    lngStart = 1
    For lngPos = 1 To lngMax
        '�����ѧ����,��������Ƶ���һ����ʾ
        lngASC = Asc(Mid(strLable, lngPos, 1))

        '����Ƿ񳬿�(���ȳ����п�,���������س����з�)
        If TextWidth(Mid(strLable, lngStart, lngPos - lngStart + 1) & "��") > lngScaleWidth Or lngPos = lngMax Or lngASC = 10 Then
            If lngPos = lngMax Or lngASC = 10 Then
                strAppRow = Mid(strLable, lngStart, lngPos - lngStart + 1)
            Else
                strAppRow = Mid(strLable, lngStart, lngPos - lngStart - 1) & "��"
            End If
            lngStart = lngPos + 1
            
            '���������
            Set objAppRow = New zlTFTabAppRow
            Call objAppRow.Add(strAppRow)
            Call objPrint.UnderAppRows.Add(objAppRow)
            
            If lngPos = lngMax Or lngASC = 10 Then
            Else
                Exit For        '�����¼�����ϱ�ǩ����Ҳֻ��һ�У����ϱ�ǩ�������仯��Ӱ�����ӡ�У����Թ̶����������
            End If
        End If
    Next
    '60333:������,2013-10-14,�޸��ļ���ŵ��´�ӡ��Ϣ��ʧ
    If UBound(Split(mstrPaperSet, ";")) >= 3 Then
        lngMax = Val(Split(mstrPaperSet, ";")(3))
    Else
        lngMax = Printer.Width
    End If
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
    glngHideCols = cHideCols
    glngSignName = IIf(mblnSignPic = True, mlngSignName, -1)
    '64583:������,2013-09-22,��ӡʱͬһ�������Ƿ��ظ���ʾ
    glngDate = IIf(mblnDateModel = True, mlngDate, -1)
    glngCollectColor = mlngCollectColor
    
    ShowPage = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetFixedProperty(ByVal strName As String) As Variant
'�������ƻ�ȡ�̶�����
    Dim varProperty As Variant
    Select Case strName
        Case "��Ч������"
            varProperty = mlngPageRows
        Case "������ɫ"
            varProperty = mlngCollectColor
    End Select
    GetFixedProperty = varProperty
End Function

Public Function GetFixedCol(ByVal strName As String) As Long
'�������ƻ�ȡ�̶�����Ϣ
    Dim lngCol As Long
    Select Case strName
        Case "����"
            lngCol = mlngDate
        Case "ʱ��"
            lngCol = mlngTime
        Case "��ʿ"
            lngCol = mlngOperator
        Case "ǩ����"
            lngCol = mlngSignName
        Case "ǩ��ʱ��"
            lngCol = mlngSignTime
        Case "ǩ������"
            lngCol = mlngSignLevel
        Case "ǩ����Ϣ"
            lngCol = mlngSigner
        Case "����ǩ����"
            lngCol = mlngJoinSignName
        Case "�ļ�ID"
            lngCol = mlngFileID
        Case "��¼ID"
            lngCol = mlngRecord
        Case "����"
            lngCol = mlngRowCount
        Case "ʵ������"
            lngCol = mlngRowCurrent
        Case "��ʼ�к�"
            lngCol = mlngStartSpread
        Case "��ӡ����ҳ��"
            lngCol = mlngPrintedEndPage
        Case "����ֵ"
            lngCol = mlngCollectValue
        Case "��ӡ��ʶ"
            lngCol = mlngPrintedTag
        Case "�������"
            lngCol = mlngCollectType
        Case "�����ı�"
            lngCol = mlngCollectText
        Case "���ܱ��"
            lngCol = mlngCollectStyle
        Case "��������"
            lngCol = mlngCollectDay
        Case "��ӡҳ��"
            lngCol = mlngPrintedPage
        Case "��ӡ�к�"
            lngCol = mlngPrintedRow
        Case "��ֹ�༭"
            lngCol = mlngNoEditor
    End Select
    GetFixedCol = lngCol
End Function

Public Function GetStartPage() As Integer
    If UBound(mArrPage) < 0 Then
        GetStartPage = 1
    Else
        GetStartPage = Val(mArrPage(0))
    End If
End Function

Public Function GetCollectCols(ByVal lngRaw As Long) As String
    GetCollectCols = VsfData.TextMatrix(lngRaw, mlngCollectValue)
End Function

Public Function GetPages() As Integer
    If UBound(mArrPage) < 0 Then
        GetPages = 1
    Else
        GetPages = mlngMaxIndex + Val(mArrPage(0))
    End If
End Function

Public Function isEndPage() As Boolean
    isEndPage = (mlngMinIndex = mlngMaxIndex)
End Function

Public Sub PrevPage()
    If mlngMinIndex > 0 Then
        mlngMinIndex = mlngMinIndex - 1
        If mlngMinIndex <= UBound(mArrPage) Then
            Call ShowPage
        End If
    End If
End Sub

Public Function NextPage() As Boolean
    If mlngMinIndex < mlngMaxIndex Then
        mlngMinIndex = mlngMinIndex + 1
        If mlngMinIndex <= UBound(mArrPage) Then
            NextPage = ShowPage
        End If
    End If
End Function

Public Function AppointPage(ByVal intPage As Integer) As Boolean
    If UBound(mArrPage) >= 0 Then
        If intPage <= mlngMaxIndex + Val(mArrPage(0)) Then
            mlngMinIndex = intPage - Val(mArrPage(0))
            AppointPage = ShowPage
        End If
    End If
End Function

Public Function GetFileName() As String
    GetFileName = lblTitle.Caption
End Function

Public Function blnOddEvenPagePrint() As Boolean
    blnOddEvenPagePrint = mblnOddEvenPagePrint
End Function

Public Function blnShowNullCollet() As Boolean
    blnShowNullCollet = mblnShowNullCollet
End Function

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long
    Dim lngRow As Long, lngCol As Long
    On Error GoTo ErrHand
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
                    
                    If Val(.TextMatrix(lngCount, mlngCollectType)) = 0 Then
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
                End If
                '����С�����ʾ
                '65889:������,2013-11-1,����С���п�ҳ���������֤��һҳС������Ҳ����ȷ��ʾС�����ƣ����������ݷ���ʱ��
                '��� (lngCount - .FixedRows + 1) Mod mlngPageRows = 1
                If FormatValue(VsfData.TextMatrix(lngCount, mlngRowCount)) Like "*|1" Or (lngCount - .FixedRows + 1) Mod mlngPageRows = 1 Then
                    If Val(VsfData.TextMatrix(lngCount, mlngCollectType)) <> 0 Then
                        VsfData.TextMatrix(lngCount, mlngDate) = VsfData.TextMatrix(lngCount, mlngCollectText)
                        If mblnʱ�������� = False Then
                            VsfData.TextMatrix(lngCount, mlngTime) = VsfData.TextMatrix(lngCount, mlngCollectText)
                        End If
            
                        '88967:��ʿ��ǩ����ͬʱ���ڣ�������ͬһ����Ա����Ӧ����ϲ�(��ӡǩ�������ǩ��ͼƬ��ע��ǩ�����лس������)
                        For lngCol = mlngTime + 1 To IIf(mlngNoEditor < mlngSignName, mlngSignName, mlngNoEditor)
                            '52953,������,2012-08-24,��������Ϊ0ҲҪ��ʾ,��������:60792
                            'If .TextMatrix(lngCount, lngCOL) = "0" Then .TextMatrix(lngCount, lngCOL) = ""
                            If Trim(.TextMatrix(lngCount, lngCol)) <> "" And .ColHidden(lngCol) = False Then
                                '66085:������,2012-09-26,�������ڻ����кϲ�,��ԭ����������+�ո�ͬһ�ĳ����к�����chr(13)
                                '������ӿո���п�������������ʾ����ȫ(��Ҫ����Ҷ���)
'                                Select Case .ColAlignment(lngCol)
'                                    Case 6, 7, 8
'                                        .TextMatrix(lngCount, lngCol) = IIf(lngCol Mod 2 = 1, " ", "") & .TextMatrix(lngCount, lngCol)
'                                    Case 3, 4, 5
'                                        .TextMatrix(lngCount, lngCol) = IIf(lngCol Mod 2 = 1, " ", "") & .TextMatrix(lngCount, lngCol) & IIf(lngCol Mod 2 = 1, " ", "")
'                                    Case 0, 1, 2
'                                        .TextMatrix(lngCount, lngCol) = .TextMatrix(lngCount, lngCol) & IIf(lngCol Mod 2 = 1, " ", "")
'                                    Case Else
'                                        .TextMatrix(lngCount, lngCol) = IIf(lngCol Mod 2 = 1, " ", "") & .TextMatrix(lngCount, lngCol)
'                                End Select
                                .TextMatrix(lngCount, lngCol) = .TextMatrix(lngCount, lngCol) & IIf(lngCol Mod 2 = 1, Chr(13), "")
                            End If
                        Next
                        .MergeRow(lngCount) = True
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
        
        '������Ϊ�յ��еķ���ʱ��Ҳ����Ϊ��(��������)
        If mbln����ʱ��ϲ� Then
            lngRow = VsfData.FixedRows
            Do While True
                If lngRow > VsfData.Rows - 1 Then Exit Do
                If VsfData.TextMatrix(lngRow, mlngDate) = "" Then
                    VsfData.TextMatrix(lngRow, 1) = ""
                    VsfData.TextMatrix(lngRow, 2) = ""
                Else
                    If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then
                        VsfData.TextMatrix(lngRow, 2) = VsfData.TextMatrix(lngRow, mlngDate)
                    Else
                        VsfData.TextMatrix(lngRow, 1) = Format(VsfData.TextMatrix(lngRow, 1), "yyyy-MM-dd HH:mm")
                        VsfData.TextMatrix(lngRow, 2) = Format(VsfData.TextMatrix(lngRow, 1), "yyyy-MM-dd HH:mm")
                    End If
                End If
                lngRow = lngRow + 1
            Loop
        End If
        
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
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
    On Error GoTo ErrHand
    
    gstrSQL = " Select   ��ʼʱ��,����ʱ��,��ʽID,����ID,�鵵�� From ���˻����ļ� " & _
              " Where ����ID=[1] And ��ҳID=[2] And Ӥ��=[3] And ID=[4] And Rownum<2"
    If gblnMoved Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
    End If
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", mlng����ID, mlng��ҳID, mintӤ��, mlng��ǰ�ļ�ID)
    If mrsTemp.RecordCount <> 0 Then
        mlng��ʽID = mrsTemp!��ʽID
        mlng����ID = mrsTemp!����ID
        mstr��ʼʱ�� = Format(mrsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
        mstr����ʱ�� = Format(mrsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    
    RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitEnv()
    On Error GoTo ErrHand
    
     '46251,������,2012-09-11,װ��ҳ�����λ��
    With cboҳ��
        .Clear
        .AddItem "ҳü�Ϸ�": .ItemData(.NewIndex) = 1
        .AddItem "ҳü�·�": .ItemData(.NewIndex) = 2
        .AddItem "ҳ���Ϸ�": .ItemData(.NewIndex) = 3
        .AddItem "ҳ���·�": .ItemData(.NewIndex) = 4
        cboҳ��.Tag = 3
        Call zlControl.CboSetIndex(cboҳ��.hWnd, 2)
    End With
    
    '���ִ��ڵ����л����¼��Ŀ
    gstrSQL = " Select   ��Ŀ���,upper(��Ŀ����) AS ��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ" & _
              " From �����¼��Ŀ B" & _
              " Order by ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    
    '��ȡ�����ڼ�¼��������������Ŀ
    gstrSQL = _
        " Select i.����id, i.����, i.������, nvl(i.�滻��,0) �滻��,i.����,i.����,i.С��,i.��λ,i.��ʾ��,i.��ֵ��,i.����" & vbNewLine & _
        " From ����������Ŀ i, ������������ k" & vbNewLine & _
        " Where k.Id = i.����id And ((k.���� In ('02', '05', '06') And i.�滻�� = 1) Or (k.���� = 2 And k.���� = '06' And NVL(i.�滻��,0) = 0))" & vbNewLine & _
        " Order By k.����, k.����, i.����"
    Set mrsElement = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ڼ�¼��������������Ŀ")
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, Optional ByVal strPages As String = "") As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngFileID           �ļ�ID
    '       lngPatiID           ����id
    '       lngPageID           ��ҳid
    '       intBaby             Ӥ����־
    '       strPage             Ϊ��˵���ӵ�һҳ��ʼ���д�ӡ,��Ϊ�ո�ʽΪ��ҳ��;��ʶ(�����������ӡ),ҳ��;��ʶ......
    '���أ� ��
    '******************************************************************************************************************
    Dim mrsTemp As New ADODB.Recordset
    Dim i As Long
    Dim arrTemp() As String
    On Error GoTo ErrHand
    Err = 0
    
    mArrPage = Array(): mlngMinIndex = -1: mlngMaxIndex = -1
    mblnInit = False
    mlng��ǰ�ļ�ID = lngFileID
    mlng����ID = lngPatiID
    mlng��ҳID = lngPageId
    mintӤ�� = intBaby
    mlngPageRows = frmAsk.mintPageRows
    Set mfrmParent = frmParent
    mintNORule = Val(zlDatabase.GetPara("�����ļ�ҳ�����", glngSys, 1255, 0))
    mblnSignPic = (Val(zlDatabase.GetPara("��¼��ǩ������ʾ��ʽ", glngSys, 1255, 0)) = 1)
    '56134:������,2012-12-19,��¼����ӡʱ,����δ��ҳ�հײ���������
    mblnPrintRow = (Val(zlDatabase.GetPara("��¼��δ��ҳ��ӡ���", glngSys, 1255, 0)) = 1)
    '46506:������,2012-12-19,��¼����ӡʱ��������ҳ�Ž������(�ļ�Ϊ����ʱ��Ч)
    mblnFullPagePrint = (Val(zlDatabase.GetPara("��¼����ҳ��ӡ", glngSys, 1255, 0)) = 1)
    '49753:������,2012-12-19,��¼����ӡʱ������ҳ��ż���
    mblnOddEvenPagePrint = (Val(zlDatabase.GetPara("��¼����ż��ӡ", glngSys, 1255, 0)) = 1)
    '--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
    mlngSingerType = Val(zlDatabase.GetPara("��ʿ��ǩ������ʾģʽ", glngSys, 1255, "2"))
    If InStr(1, ",0,1,2,3,", "," & mlngSingerType & ",") = 0 Then mlngSingerType = 2
    '64583:������,2013-09-22,Ԥ������ӡʱͬһҳ��ͬ������ʾ��ʽ:���;һ��
    mblnDateModel = (Val(zlDatabase.GetPara("��¼��������ʾ��ʽ", glngSys, 1255, 0)) = 1)
    '68739:������,2014-1-2,���"С���ʶ��ɫ"
    mlngCollectColor = Val(zlDatabase.GetPara("С���ʶ��ɫ", glngSys, 1255, "255"))
    
    arrTemp = Split(zlDatabase.GetPara("С��ȱʡ��ʽ", glngSys, 1255), ";")
    If UBound(arrTemp) > 0 Then
        mblnShowNullCollet = arrTemp(1) = 0
    Else
        mblnShowNullCollet = True
    End If
    '�жϲ����Ƿ�ת��
    gstrSQL = "Select ����ת�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж������Ƿ�ת��", mlng����ID, mlng��ҳID)
    gblnMoved = NVL(mrsTemp!����ת��, 0) <> 0
    
    '���ʱ������ӡ�������ȫ����ӡ
    If gblnBatch = False Then
        '�жϵ�ǰ�ļ��Ƿ��Ѿ�����
        gstrSQL = " Select  ����ʱ�� From ���˻����ļ� " & _
                  " Where ����ID=[1] And ��ҳID=[2] And Ӥ��=[3] And ID=[4] And Rownum<2"
        Call SQLDIY(gstrSQL)
        Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", mlng����ID, mlng��ҳID, mintӤ��, mlng��ǰ�ļ�ID)
        If mrsTemp.RecordCount > 0 Then
            '����ļ��Ѿ�����,���ܼ�¼���Ƿ���ҳ�����д�ӡ
            If Trim(NVL(mrsTemp!����ʱ��)) <> "" Then mblnFullPagePrint = False
        End If
    Else
        mblnFullPagePrint = False
        mblnOddEvenPagePrint = False
    End If
    If mblnFullPagePrint = True Then mblnPrintRow = False
    
    If mrsItems.State = 0 Then
        Call InitEnv            '��ʼ������
    End If
    Call InitVariable
    
    mstrMergeID = ""
    gstrSQL = " Select MIN(��ʼҳ��) AS ��ʼҳ��,MAX(����ҳ��) AS ����ҳ�� From ���˻����ӡ Where �ļ�ID=[1]"
    Call SQLDIY(gstrSQL)
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ļ���С�����ҳ��", mlng��ǰ�ļ�ID)
    mint����ҳ = Val(NVL(mrsTemp!����ҳ��, 0))
    mintҳ�� = Val(NVL(mrsTemp!��ʼҳ��, 0))
    If mint����ҳ = 0 Then Exit Function
    
    If strPages <> "" Then
        For i = 0 To UBound(Split(strPages, ","))
            If Val(Split(strPages, ",")(i)) >= mintҳ�� And Val(Split(strPages, ",")(i)) <= mint����ҳ Then
                ReDim Preserve mArrPage(UBound(mArrPage) + 1)
                mArrPage(UBound(mArrPage)) = Split(strPages, ",")(i)
            End If
        Next i
    Else
        For i = mintҳ�� To mint����ҳ
            ReDim Preserve mArrPage(UBound(mArrPage) + 1)
            mArrPage(UBound(mArrPage)) = i & ";2"
        Next i
    End If
   
    '�������ҳ��ӡҳ������ѡ��Ĵ�ӡ���һҳ�����ļ����ҳ�ţ������Ƿ���ҳ
    If mblnFullPagePrint = True And mint����ҳ = Val(mArrPage(UBound(mArrPage))) Then
        gstrSQL = "Select Max(�����к�) �����к� From ���˻����ӡ where �ļ�ID=[1] And ����ҳ��=[2] "
        Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ļ����һҳ�Ľ����к�", mlng��ǰ�ļ�ID, mint����ҳ)
        If mrsTemp!�����к� < mlngPageRows Then
            If UBound(mArrPage) <= 0 Then
                mArrPage = Array()
            Else
                ReDim Preserve mArrPage(UBound(mArrPage) - 1)
            End If
        End If
    End If
    If UBound(mArrPage) < 0 Then
        If gblnBatch = False Then
            MsgBox "û������������ɴ�ӡ�Ļ����¼�����ݣ�", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    '��ż��ӡʱ����ҳ���Ƿ��������������򲻽�����ż��ӡ
    If mblnOddEvenPagePrint = True Then
        mintҳ�� = Val(mArrPage(0)) - 1
        For i = 0 To UBound(mArrPage)
            If mintҳ�� + 1 <> Val(mArrPage(i)) Then
                mblnOddEvenPagePrint = False
                Exit For
            End If
            mintҳ�� = Val(mArrPage(i))
        Next i
    End If
    
    mlngMinIndex = 0: mlngMaxIndex = UBound(mArrPage)
    mintҳ�� = Val(mArrPage(mlngMinIndex))
    mint����ҳ = Val(mArrPage(UBound(mArrPage)))
    
    mstrMergeID = ""
    '��ȡ�ϲ��ļ���Ϣ
    gstrSQL = _
        "Select Id From (With ���˻����ļ�_F1 As" & vbNewLine & _
        " (Select a.Id, a.����id From ���˻����ļ� a Where a.����id = [1] And a.��ҳid = [2] And Nvl(a.Ӥ��, 0) = [3])" & vbNewLine & _
        "Select Id From ���˻����ļ�_F1 Start With ����id = [4] Connect By Prior Id = ����id Order By Level Desc)"
    Call SQLDIY(gstrSQL)
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��鵱ǰ�ļ��Ƿ��������ļ�����Ϊ�ϲ���ӡ", mlng����ID, mlng��ҳID, mintӤ��, mlng��ǰ�ļ�ID)
    Do While Not mrsTemp.EOF
        mstrMergeID = mstrMergeID & "," & mrsTemp!ID
    mrsTemp.MoveNext
    Loop
    mstrMergeID = Mid(mstrMergeID, 2)
    Call ShowPage
    mblnInit = True
    mblnEditable = False
    ShowMe = True
'    Call OutputRsData(mrsSelItems)
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadPageData() As Boolean
    Dim str���� As String, lng����ҳ As Long, lngFileID As Long
    Dim blnInitRec As Boolean
    Dim arrCode, i As Integer
    On Error GoTo ErrHand
    
    mint��ǰ��ʼҳ = mintҳ��
    Set mrsDataMap = New ADODB.Recordset
    lngFileID = mlng��ǰ�ļ�ID
    '�ϲ��ļ�����
    blnInitRec = False
    arrCode = Split(mstrMergeID, ",")
    If UBound(arrCode) >= 0 And mintҳ�� = Val(mArrPage(0)) Then
        For i = 0 To UBound(arrCode)
            gstrSQL = "Select MAX(����ҳ��) ����ҳ From ���˻����ӡ Where �ļ�ID=[1]"
            Call SQLDIY(gstrSQL)
            Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ϲ��ļ�������ҳ��", Val(arrCode(i)))
            lng����ҳ = NVL(mrsTemp!����ҳ, 0)
            If lng����ҳ = mintҳ�� Then
                mlng��ǰ�ļ�ID = Val(arrCode(i))
                Call ReadStruDef
                If Not blnInitRec Then
                    Call InitRecords
                    blnInitRec = True
                End If
                str���� = " And P.����ҳ��=[5]"
                Call zlRefresh(str����)
            End If
        Next i
    End If
    mlng��ǰ�ļ�ID = lngFileID
    'Ҫ��ӡ���ļ�����
    Call ReadStruDef
    If Not blnInitRec Then
        Call InitRecords
        blnInitRec = True
    End If
    str���� = " AND (P.��ʼҳ��=[5] OR (P.����ҳ��=[5])) "
    Call zlRefresh(str����)
    Call PreTendFormat
    LoadPageData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
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
    mlngSignName = -1
    mlngFileID = -1
    mlngRecord = -1
    mlngNoEditor = -1
    mlngPrintedEndPage = -1
    mlngCollectValue = -1
    
    mblnShow = False
    mblnSign = False
    mblnArchive = False
    mblnEditAssistant = False
    
    Set mrsDataMap = New ADODB.Recordset
End Sub

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '��ȡ������ʼ��,������ҳ�򷵻�0
    '�����ҳδ��ʾȫ,��˵��������ҳ,Ҳ����0
    '���������������������в�������
    
    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '������
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '��ǰ��
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    'Ѱ����ʼ��
    For lngRow = lngRow To 3 Step -1
        If FormatValue(VsfData.TextMatrix(lngRow, mlngRowCount)) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    
    GetStartRow = lngStart
End Function

Public Function GetDiagonal() As String
    GetDiagonal = "," & mstrCatercorner & "," '& mstrCOLNothing & ","
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
    On Error Resume Next
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
    If Not mrsTemp Is Nothing Then
        If mrsTemp.State = adStateOpen Then mrsTemp.Close
        Set mrsTemp = Nothing
    End If
    If Not mrsItems Is Nothing Then
        If mrsItems.State = adStateOpen Then mrsItems.Close
        Set mrsItems = Nothing
    End If
    If Not mrsElement Is Nothing Then
        If mrsElement.State = adStateOpen Then mrsElement.Close
        Set mrsElement = Nothing
    End If
    If Not mrsSelItems Is Nothing Then
        If mrsSelItems.State = adStateOpen Then mrsSelItems.Close
        Set mrsSelItems = Nothing
    End If
    If Not mrsDataMap Is Nothing Then
        If mrsDataMap.State = adStateOpen Then mrsDataMap.Close
        Set mrsDataMap = Nothing
    End If
    If Not mfrmParent Is Nothing Then Set mfrmParent = Nothing
    If Not mobjTagFont Is Nothing Then Set mobjTagFont = Nothing
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

Private Sub SingerShowType(ByVal vsfObj As VSFlexGrid, ByVal lngStartRow As Long, ByVal lngEndRow As Long)
'-------------------------------------------------
'���ܣ���ʿǩ������ʾ��ʽ
''--58414,������,2013-04-10,��ӻ�ʿ��ǩ������ʾģʽ
'-------------------------------------------------
    Dim lngRow As Integer
    
    Select Case mlngSingerType
        Case 0 '��������ʾ
            For lngRow = lngStartRow To lngEndRow
                If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
            Next
        Case 1 '������ʾ
            For lngRow = lngStartRow To lngEndRow
                If lngRow = lngStartRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
        Case 3 'β����ʾ
            If mlngOperator > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngOperator) = "" Then vsfObj.TextMatrix(lngStartRow, mlngOperator) = vsfObj.TextMatrix(lngEndRow, mlngOperator)
            End If
            If mlngSignName > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngSignName) = "" Then vsfObj.TextMatrix(lngStartRow, mlngSignName) = vsfObj.TextMatrix(lngEndRow, mlngSignName)
            End If
            If mlngSignTime > 0 Then
                If vsfObj.TextMatrix(lngStartRow, mlngSignTime) = "" Then vsfObj.TextMatrix(lngStartRow, mlngSignTime) = vsfObj.TextMatrix(lngEndRow, mlngSignTime)
            End If
            For lngRow = lngEndRow To lngStartRow Step -1
                If lngRow = lngEndRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
        Case Else '��β��ʾ
            '���һ����Ҫ��д���ǩ��
            For lngRow = lngStartRow To lngEndRow
                If lngRow = lngStartRow Or lngRow = lngEndRow Then
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = vsfObj.TextMatrix(lngStartRow, mlngOperator)
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = vsfObj.TextMatrix(lngStartRow, mlngSignName)
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = vsfObj.TextMatrix(lngStartRow, mlngSignTime)
                Else
                    If mlngOperator > 0 Then vsfObj.TextMatrix(lngRow, mlngOperator) = ""
                    If mlngSignName > 0 Then vsfObj.TextMatrix(lngRow, mlngSignName) = ""
                    If mlngSignTime > 0 Then vsfObj.TextMatrix(lngRow, mlngSignTime) = ""
                End If
            Next
    End Select
End Sub

