VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.Form frmPrintPreview 
   Caption         =   "��ӡԤ��"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   9345
   Icon            =   "frmPrintPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9345
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMerge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   4320
      ScaleHeight     =   1335
      ScaleWidth      =   2655
      TabIndex        =   14
      ToolTipText     =   "���ڱ���ͼƬ�����ͼ"
      Top             =   840
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox picOrig 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   6840
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   13
      ToolTipText     =   "���ڱ���ͼƬ�����ͼ"
      Top             =   120
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picPrintBlank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   6120
      ScaleHeight     =   390
      ScaleWidth      =   570
      TabIndex        =   12
      ToolTipText     =   "���ڻ���ʼҳ�����Ŀհ�����"
      Top             =   120
      Visible         =   0   'False
      Width           =   600
   End
   Begin zlSubclass.Subclass Subclass1 
      Left            =   3480
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.PictureBox picZoom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   3375
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   10
      ToolTipText     =   "������ʱ���ͼƬ"
      Top             =   45
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.VScrollBar VS 
      DragIcon        =   "frmPrintPreview.frx":038A
      Height          =   2145
      LargeChange     =   20
      Left            =   8955
      Max             =   100
      MouseIcon       =   "frmPrintPreview.frx":0694
      SmallChange     =   10
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2070
      Width           =   250
   End
   Begin VB.HScrollBar HS 
      DragIcon        =   "frmPrintPreview.frx":07E6
      Height          =   250
      LargeChange     =   20
      Left            =   2835
      Max             =   100
      MouseIcon       =   "frmPrintPreview.frx":0AF0
      SmallChange     =   10
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4410
      Width           =   6105
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   2835
      ScaleHeight     =   2145
      ScaleWidth      =   6060
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2250
      Width           =   6060
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   135
         MouseIcon       =   "frmPrintPreview.frx":0C42
         MousePointer    =   99  'Custom
         ScaleHeight     =   930
         ScaleWidth      =   5790
         TabIndex        =   6
         Top             =   180
         Width           =   5820
         Begin VB.PictureBox picBlank 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   40
            Left            =   0
            MouseIcon       =   "frmPrintPreview.frx":0D94
            MousePointer    =   99  'Custom
            ScaleHeight     =   45
            ScaleWidth      =   825
            TabIndex        =   11
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   225
         ScaleHeight     =   960
         ScaleWidth      =   5820
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   5820
      End
   End
   Begin VB.PictureBox picҳ�� 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   135
      ScaleHeight     =   2220
      ScaleWidth      =   2310
      TabIndex        =   3
      Top             =   1350
      Visible         =   0   'False
      Width           =   2310
      Begin VSFlex8Ctl.VSFlexGrid vfgҳ�� 
         Height          =   1695
         Left            =   45
         TabIndex        =   4
         Top             =   135
         Width           =   2100
         _cx             =   3704
         _cy             =   2990
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   10197915
         ForeColor       =   -2147483640
         BackColorFixed  =   10197915
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483634
         BackColorBkg    =   10197915
         BackColorAlternate=   10197915
         GridColor       =   10197915
         GridColorFixed  =   10197915
         TreeColor       =   16777215
         FloodColor      =   192
         SheetBorder     =   10197915
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPrintPreview.frx":0FA6
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   7
         PicturesOver    =   -1  'True
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
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
         WallPaperAlignment=   1
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picBuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   2655
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   2
      ToolTipText     =   "���ڻ���ÿҳ��ԭʼͼ��������Ԥ������ӡͼƬ��ʧ�棬����Ҫ���»�"
      Top             =   45
      Visible         =   0   'False
      Width           =   645
   End
   Begin XtremeSuiteControls.TabControl tabThis 
      Height          =   1230
      Left            =   90
      TabIndex        =   1
      Top             =   810
      Width           =   2595
      _Version        =   589884
      _ExtentX        =   4577
      _ExtentY        =   2170
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7065
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPrintPreview.frx":0FE9
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13600
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   4770
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlZoom 
      Left            =   2760
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   1
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event PrintEpr()

'�ļ� "File"
Private Const ID_FILE_PRINT = 304       '��ӡ(P)...
Private Const ID_FILE_EXIT = 305        '�˳�(X)
Private Const ID_File_SaveAsPic = 306   '���ΪͼƬ(I)

'��ͼ "View"
Private Const ID_View_ToolBar = 310     '������(T)
Private Const ID_View_StatusBar = 311   '״̬��(S)
Private Const ID_View_ZoomFactor = 312  '���ű���(C)
Private Const ID_View_First = 313       '��һҳ
Private Const ID_View_Prev = 314        'ǰһҳ
Private Const ID_View_Next = 315        '��һҳ
Private Const ID_View_Last = 316        '���һҳ
Private Const ID_View_ActualSize = 317  'ʵ�ʴ�С Ctrl+1
Private Const ID_View_FitSize = 318     '�ʺ�ҳ�� Ctrl+2
Private Const ID_View_FitWidth = 319    '�ʺϿ�� Ctrl+3
Private Const ID_View_FitHeight = 320   '�ʺϸ߶� Ctrl+4
Private Const ID_View_Size_250 = 330    '250%
Private Const ID_View_Size_200 = 331    '200%
Private Const ID_View_Size_150 = 332    '150%
Private Const ID_View_Size_100 = 333    '100%
Private Const ID_View_Size_75 = 334     '75%
Private Const ID_View_Size_50 = 335     '50%
Private Const ID_View_Size_25 = 336     '25%
Private Const ID_View_ZoomIn = 337      '�Ŵ�
Private Const ID_View_ZoomOut = 338     '��С

Private Const ID_View_StartPage = 340   '��ʼҳ��

'���� "Help"
Private Const ID_HELP_CONTENT = 500     '��������
Private Const ID_HELP_CONTACT = 502     '���ͷ���
Private Const ID_HELP_ONLINE = 503      '����ҽҵ
Private Const ID_HELP_ABOUT = 504       '����...

Private mcolMerge As Collection     '�ڱ��ϲ���Ԫ���ϼ�¼������Ԫ������У��Ա����������Ԫ��
Private mcolMergePic As Collection  '��ҳ�ϲ�������Ԫ��ͼƬ����
Private mcolPage As Collection      '����Ԥ����ÿҳͼƬ�ļ��ϣ����ھ������⣬��ӡʱ��������ЩͼƬ����Ҫ������

Private cboStartPage As CommandBarComboBox  '��ʼҳ��
Private mTableThis As cTableEPR
Private mlngCurPage As Long             '��ǰҳ
Private mlngPageCount As Long           '��ҳ��
Private mlngStartPage As Long           '��ʼҳ��
Private mlngBlankHeight As Long         '��ʼҳ���ϲ����׸߶�

Private m_bSubClassing As Boolean
Private mlngX As Long, mlngY As Long, mblnMouseDown As Boolean
Private mdblZoomFactor As Double        '���ű���
Private Const Shadow_W = 60             '��Ӱ���

Private Type tPage
    BRow As Long
    Erow As Long
    BCol As Long
    ECol As Long
End Type
Private mPages() As tPage   '�±��0��ʼ

Private Type tPaper
    PaperWidth As Long  'ֽ�ſ��(�Ѱ���ӡ����ת��)
    PaperHeight As Long 'ֽ�Ÿ߶�
    AvailableWidth As Long  '���ÿ��=ֽ�ſ��-ҳ�߾�
    AvailableHeight As Long '���ø߶�=ֽ�Ÿ߶�-ҳ�߾�
    PaperType As Integer
    Orientation As Integer 'ֽ�ŷ���:1-����,2-����
End Type
Private mPaper As tPaper
Private mRatioX As Single, mRatioY As Single

Private mobjParent As Object    '������

'################################################################################################################
'���ư�͸�����ο���غ���������
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
' BlendOp:
Private Const AC_SRC_OVER = &H0
' AlphaFormat:
Private Const AC_SRC_ALPHA = &H1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function AlphaBlend Lib "MSIMG32.dll" ( _
  ByVal hDCDest As Long, _
  ByVal nXOriginDest As Long, _
  ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, _
  ByVal nHeightDest As Long, _
  ByVal hdcSrc As Long, _
  ByVal nXOriginSrc As Long, _
  ByVal nYOriginSrc As Long, _
  ByVal nWidthSrc As Long, _
  ByVal nHeightSrc As Long, _
  ByVal lBlendFunction As Long _
) As Long

' cAlphaDibSection functions:
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Public Sub OutPut(ByRef frmParent As Object, ByRef TableThis As cTableEPR, Optional ByVal blnPreview As Boolean = True, Optional ByVal strPrintDeviceName As String)
    '******************************************************************************************************************
    '## ���ܣ�  ���һ�ݱ��������ӡԤ��������ӡ��
    '##
    '## ������frmParent       ��������
    '##       TableThis       �����������󼰲����ļ�����ĶԼ�����
    '##       blnPreView      ���Ƿ��ӡԤ��
    '******************************************************************************************************************
    Dim dblZoom As Double, lngPaperW As Long, lngPaperH As Long, i As Long
        
    Set mTableThis = TableThis
    Set mcolMergePic = New Collection
    Set mcolPage = New Collection
    Set mobjParent = frmParent
    
    With mTableThis.EPRFileInfo
        '�����Ѱ������任
        lngPaperW = .PaperWidth
        lngPaperH = .PaperHeight
        
        mPaper.PaperWidth = lngPaperW
        mPaper.PaperHeight = lngPaperH
        mPaper.AvailableWidth = lngPaperW - .MarginLeft - .MarginRight
        mPaper.AvailableHeight = lngPaperH - .MarginTop - .MarginBottom
        mPaper.PaperType = .PaperKind
        mPaper.Orientation = .PaperOrient
    End With
    
    If blnPreview Then
        Call InitCommandBars    '��������ʼ��
    End If
    mdblZoomFactor = 1#
    
    '=================================================================================================
    
    '��ҳ
    Call SplitPage
    Call SetMergeRalation
    mlngCurPage = 1
    mRatioX = 1: mRatioY = 1
    
    If blnPreview Then
        zlCommFun.ShowFlash "���Ժ�..."
        Screen.MousePointer = vbHourglass
    
        '��ÿҳ���ص�����ͼ�б���
        vfgҳ��.Rows = mlngPageCount
        vfgҳ��.ColWidth(0) = 0
        vfgҳ��.ColWidth(1) = 2100
        vfgҳ��.RowHeightMin = 2900
        vfgҳ��.FixedRows = 0
        vfgҳ��.FixedCols = 0
        
        With mPaper
            If .PaperWidth / 2000 > .PaperHeight / 3000 Then
                dblZoom = 2000 / .PaperWidth
                vfgҳ��.RowHeightMin = .PaperHeight * dblZoom + 20
            Else
                dblZoom = 3000 / .PaperHeight
            End If
            picBuff.Width = .PaperWidth
            picBuff.Height = .PaperHeight
            picZoom.Width = .PaperWidth * dblZoom
            picZoom.Height = .PaperHeight * dblZoom
            cboStartPage.Clear
            
            For i = 1 To mlngPageCount
                picBuff.Cls
                DrawPage i, 0, picBuff
                mcolPage.Add picBuff.Image, "K" & i     '���棬��������ʱֱ������������imagelist�ؼ����棬ͼƬ̫��ʱ���ڴ����
                
                '����ͼ�����ð�ɫ������Ч����ã�
                picZoom.Cls
                SetStretchBltMode picZoom.hdc, HALFTONE
                StretchBlt picZoom.hdc, 0, 0, picZoom.Width, picZoom.Height, picBuff.hdc, 0, 0, picBuff.Width, picBuff.Height, SRCCOPY
        '            picZoom.PaintPicture picBuff.Image, 0, 0, .PaperWidth * dblZoom, .PaperHeight * dblZoom
                
                picZoom.Line (0, 0)-(picZoom.ScaleWidth - 15, picZoom.ScaleHeight - 15), RGB(99, 99, 99), B
                picZoom.Line (15, 15)-(picZoom.ScaleWidth - 30, picZoom.ScaleHeight - 30), vbBlack, B
                
                cboStartPage.AddItem "�� " & CStr(i) & " ҳ"
                imlZoom.ListImages.Add 1, "K" & i, picZoom.Image
                vfgҳ��.Cell(flexcpText, i - 1, 0) = i
                vfgҳ��.Cell(flexcpPicture, i - 1, 1) = imlZoom.ListImages("K" & i).Picture
                vfgҳ��.Cell(flexcpPictureAlignment, i - 1, 1) = 3
                imlZoom.ListImages.Clear   'ֻ����ʱʹ�ã�������ͷ��ڴ�
            Next
            vfgҳ��.Row = 0
            vfgҳ��_RowColChange
        End With
        
        zlCommFun.StopFlash
        Screen.MousePointer = vbDefault
    End If
    '=================================================================================================
    
    
    If blnPreview Then
        Me.Show vbModal, mobjParent
    Else
        Call PrintTable(strPrintDeviceName)
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    pAttachMessages
End Sub

Private Sub picBlank_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mlngY = y
    mblnMouseDown = True
End Sub

Private Sub picBlank_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnMouseDown Then
        Dim lngTop As Long
        lngTop = IIf((picBlank.Top + (y - mlngY)) < 0, 0, picBlank.Top + (y - mlngY))
        lngTop = IIf(lngTop > picPage.ScaleHeight, picPage.ScaleHeight - picBlank.Height, lngTop)
        picBlank.Top = lngTop
        mlngBlankHeight = IIf(picBlank.Top > 100, picBlank.Top, 0)
        mlngBlankHeight = mlngBlankHeight / mdblZoomFactor
        'ˢ�°�͸�����ο�
        Call DrawAlphaRect(mlngBlankHeight * mdblZoomFactor)
    End If
End Sub

Private Sub picBlank_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMouseDown = False
End Sub

Private Sub DrawAlphaRect(lngHeight As Long)
    '���ư�͸�����ο�
    Dim lBlend As Long
    Dim bf As BLENDFUNCTION
    
    ' Draw the first picture:
    bf.BlendOp = AC_SRC_OVER
    bf.BlendFlags = 0
    bf.SourceConstantAlpha = 255
    bf.AlphaFormat = 0
    CopyMemory lBlend, bf, 4
    
    picPage.Picture = mcolPage("K" & mlngCurPage)
    
    bf.SourceConstantAlpha = 65
    CopyMemory lBlend, bf, 4
    AlphaBlend picPage.hdc, 0, 0, _
        picPage.ScaleWidth \ Screen.TwipsPerPixelX, _
        lngHeight \ Screen.TwipsPerPixelY, _
        picBlank.hdc, 0, 0, _
        picBlank.ScaleWidth \ Screen.TwipsPerPixelX, _
        picBlank.ScaleHeight \ Screen.TwipsPerPixelY, _
        lBlend
    picPage.Refresh
End Sub

Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    '�Զ������Ϣ������
    Dim tP As POINTAPI
    Dim sngX As Single, sngY As Single   '�������
    Dim intShift As Integer              '��갴��
    Dim bWay As Boolean                  '��귽��
    Dim bMouseFlag As Boolean            '����¼������־

    Select Case Msg
    Case WM_MOUSEWHEEL   '����
        Dim wzDelta, wKeys As Integer
        'wzDelta���ݹ��ֹ����Ŀ�������ֵС�����ʾ���������������û����򣩣�
        '�������ʾ������ǰ����������ʾ������
        wzDelta = HIWORD(wParam)
        'wKeysָ���Ƿ���CTRL=8��SHIFT=4������(��=2����=16����=2������)���£�������
        wKeys = LOWORD(wParam)
        tP.x = LOWORD(lParam)    'pt��������
        tP.y = HIWORD(lParam)
        '--------------------------------------------------
        If wzDelta < 0 Then  '���û�����
           bWay = True
        Else                 '����ʾ������
           bWay = False
        End If
        '--------------------------------------------------
        '����Ļ����ת��ΪForm1.��������
        ScreenToClient hWnd, tP
        sngX = tP.x
        sngY = tP.y
        intShift = wKeys
        bMouseFlag = True  '�ù�����־
        If bMouseFlag = True Then
            bMouseFlag = False
            DoMouseWheel bWay, intShift, sngX, sngY, CLng(wzDelta)
        End If
    End Select
End Sub

Private Sub DoMouseWheel(bBackDirection As Boolean, Shift As Integer, x As Single, y As Single, Value As Single)
    '�������Ĵ���
    If Shift = 8 Then
        '���Ŵ���
        Dim R As Double
        If bBackDirection Then
            '��С
            R = IIf(mdblZoomFactor - 0.25 < 0.25, 0.25, mdblZoomFactor - 0.25)
        Else
            R = IIf(mdblZoomFactor + 0.25 > 1#, 1#, mdblZoomFactor + 0.25)
        End If
        mdblZoomFactor = R
        PreviewPage mlngCurPage
    Else
        Dim lngR As Long
        lngR = VS.Value - IIf(Value < 0, -1, 1) * 50
        lngR = IIf(lngR > VS.Max, VS.Max, lngR)
        lngR = IIf(lngR < VS.Min, VS.Min, lngR)
        VS.Value = lngR
    End If
End Sub

Private Sub picPage_Resize()
    picBlank.Left = 0: picBlank.Width = picPage.ScaleWidth
    picBlank.Top = mlngBlankHeight * mdblZoomFactor
    picShadow.Move picPage.Left + Shadow_W, picPage.Top + Shadow_W, picPage.Width, picPage.Height
End Sub

Private Sub vfgҳ��_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    vfgҳ��.ToolTipText = "��" & vfgҳ��.MouseRow + 1 & "ҳ/��" & vfgҳ��.Rows & "ҳ"
End Sub

Private Sub vfgҳ��_RowColChange()
    vfgҳ��.ShowCell vfgҳ��.Row, 1
    mlngCurPage = vfgҳ��.Row + 1
    PreviewPage mlngCurPage
End Sub


'################################################################################################################
'## ���ܣ�  ���ΪͼƬ�ļ�
'################################################################################################################
Private Function SaveAsPicture() As Boolean
    On Error GoTo LL
    Dim strF As String
    dlgThis.Filename = ""
    dlgThis.Filter = "*.bmp|*.bmp|*.*|*.*"
    dlgThis.ShowSave
    strF = dlgThis.Filename
    If strF <> "" Then
        '���浽�ļ�
        SavePicture picPage.Image, strF
        SaveAsPicture = True
        MsgBox "����ɹ����ļ���:" & vbCrLf & strF, vbOKOnly + vbInformation, gstrSysName
    End If
    Exit Function
LL:
    MsgBox "����ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
    SaveAsPicture = False
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim i As Long
    Select Case Control.ID
    Case ID_File_SaveAsPic
        Call SaveAsPicture
    Case ID_FILE_PRINT
        '��ӡ(P)...
        Call PrintTable
    Case ID_FILE_EXIT
        '�˳�(X)
        Unload Me
    Case ID_View_ToolBar
        '������(T)
    Case ID_View_StatusBar
        '״̬��(S)
        stbThis.Visible = Not stbThis.Visible
        cbsThis.RecalcLayout
    Case ID_View_ZoomFactor
        '���ű���(C)
        Dim R As Double
        R = Val(Control.Text) / 100#
        mdblZoomFactor = R
        PreviewPage mlngCurPage
    Case ID_View_ZoomIn
        '�Ŵ�
        mdblZoomFactor = IIf(mdblZoomFactor + 0.25 > 1#, 1#, mdblZoomFactor + 0.25)
        PreviewPage mlngCurPage
    Case ID_View_ZoomOut
        '��С
        mdblZoomFactor = IIf(mdblZoomFactor - 0.25 < 0.25, 0.25, mdblZoomFactor - 0.25)
        PreviewPage mlngCurPage
    Case ID_View_First
        '��һҳ
        vfgҳ��.Row = 0
    Case ID_View_Prev
        'ǰһҳ
        vfgҳ��.Row = IIf(vfgҳ��.Row - 1 > 0, vfgҳ��.Row - 1, 0)
    Case ID_View_Next
        '��һҳ
        vfgҳ��.Row = IIf(vfgҳ��.Row + 1 > vfgҳ��.Rows, vfgҳ��.Rows, vfgҳ��.Row + 1)
    Case ID_View_Last
        '���һҳ
        vfgҳ��.Row = vfgҳ��.Rows - 1
    Case ID_View_ActualSize
        'ʵ�ʴ�С Ctrl+1
        mdblZoomFactor = 1#
        PreviewPage mlngCurPage
    Case ID_View_FitSize
        '�ʺ�ҳ�� Ctrl+2
        If picBack.ScaleWidth / mPaper.PaperWidth < picBack.ScaleHeight / mPaper.PaperHeight Then
            mdblZoomFactor = (picBack.ScaleWidth - Shadow_W * 4) / mPaper.PaperWidth
        Else
            mdblZoomFactor = (picBack.ScaleHeight - Shadow_W * 4) / mPaper.PaperHeight
        End If
        PreviewPage mlngCurPage
    Case ID_View_FitWidth
        '�ʺϿ�� Ctrl+3
        mdblZoomFactor = (picBack.ScaleWidth - Shadow_W * 4) / mPaper.PaperWidth
        PreviewPage mlngCurPage
    Case ID_View_FitHeight
        '�ʺϸ߶� Ctrl+4
        mdblZoomFactor = (picBack.ScaleHeight - Shadow_W * 4) / mPaper.PaperHeight
        PreviewPage mlngCurPage
    Case ID_View_Size_250
        '250%
        mdblZoomFactor = 2.5
        PreviewPage mlngCurPage
    Case ID_View_Size_200
        '200%
        mdblZoomFactor = 2#
        PreviewPage mlngCurPage
    Case ID_View_Size_150
        '150%
        mdblZoomFactor = 1.5
        PreviewPage mlngCurPage
    Case ID_View_Size_100
        '100%
        mdblZoomFactor = 1#
        PreviewPage mlngCurPage
    Case ID_View_Size_75
        '75%
        mdblZoomFactor = 0.75
        PreviewPage mlngCurPage
    Case ID_View_Size_50
        '50%
        mdblZoomFactor = 0.5
        PreviewPage mlngCurPage
    Case ID_View_Size_25
        '25%
        mdblZoomFactor = 0.25
        PreviewPage mlngCurPage
    Case ID_HELP_CONTENT
        '��������
        ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
    Case ID_HELP_CONTACT
        '���ͷ���
        Call zlMailTo(Me.hWnd)
    Case ID_HELP_ONLINE
        '������ҳ
        Call zlHomePage(Me.hWnd)
    Case ID_HELP_ABOUT
        '����...
        ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
    Case ID_View_StartPage
        mlngStartPage = Val(Mid(Control.Text, 3))
        If mlngStartPage = 0 Or mlngStartPage > mlngPageCount Then Exit Sub
        
        mlngCurPage = mlngStartPage
        vfgҳ��.RowHeightMin = 0
        For i = 0 To mlngStartPage - 2
            vfgҳ��.RowHeight(i) = 0
        Next
        For i = mlngStartPage - 1 To mlngPageCount - 1
            vfgҳ��.RowHeight(i) = 2900
        Next
        vfgҳ��.Row = mlngStartPage - 1
        picBlank.Top = 0
        mlngBlankHeight = 0
        picBlank.Visible = True
        vfgҳ��_RowColChange
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height / Screen.TwipsPerPixelY
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    Me.cbsThis.GetClientRect Left, Top, Right, Bottom
    tabThis.Move (Left + 1) * Screen.TwipsPerPixelX, (Top + 1) * Screen.TwipsPerPixelY, 2500, (Bottom - Top - 2) * Screen.TwipsPerPixelY
    picBack.Move tabThis.Left + tabThis.Width + Screen.TwipsPerPixelX, _
        (Top + 1) * Screen.TwipsPerPixelY, _
        (Right - Left - 2) * Screen.TwipsPerPixelX - 2500 - VS.Width, _
        (Bottom - Top - 2) * Screen.TwipsPerPixelY - HS.Height
    Reposition
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_FILE_PRINT
        '��ӡ(P)...
    Case ID_FILE_EXIT
        '�˳�(X)
    Case ID_View_ToolBar
        '������(T)
    Case ID_View_StatusBar
        '״̬��(S)
        Control.Checked = stbThis.Visible
    Case ID_View_ZoomFactor
        '���ű���(C)
        Control.Text = Format(mdblZoomFactor, "0%")
    Case ID_View_ZoomIn
        '�Ŵ�
        Control.Enabled = (mdblZoomFactor < 1#) And (Abs(mdblZoomFactor - 1#) > 0.00001)
    Case ID_View_ZoomOut
        '��С
        Control.Enabled = (mdblZoomFactor > 0.25) And (Abs(mdblZoomFactor - 0.25) > 0.00001)
    Case ID_View_First
        '��һҳ
        Control.Enabled = (mlngPageCount > 1) And (mlngCurPage > mlngStartPage)
    Case ID_View_Prev
        'ǰһҳ
        Control.Enabled = (mlngPageCount > 1) And (mlngCurPage > mlngStartPage)
    Case ID_View_Next
        '��һҳ
        Control.Enabled = (mlngPageCount > 1) And (mlngCurPage < mlngPageCount)
    Case ID_View_Last
        '���һҳ
        Control.Enabled = (mlngPageCount > 1) And (mlngCurPage < mlngPageCount)
    Case ID_View_ActualSize
        'ʵ�ʴ�С Ctrl+1
        Control.Checked = (Abs(mdblZoomFactor - 1#) < 0.00001)
    Case ID_View_FitSize
        '�ʺ�ҳ�� Ctrl+2
    Case ID_View_FitWidth
        '�ʺϿ�� Ctrl+3
    Case ID_View_FitHeight
        '�ʺϸ߶� Ctrl+4
    Case ID_View_Size_250
        '250%
    Case ID_View_Size_200
        '200%
    Case ID_View_Size_150
        '150%
    Case ID_View_Size_100
        '100%
        Control.Checked = (Abs(mdblZoomFactor - 1#) < 0.00001)
    Case ID_View_Size_75
        '75%
        Control.Checked = (Abs(mdblZoomFactor - 0.75) < 0.00001)
    Case ID_View_Size_50
        '50%
        Control.Checked = (Abs(mdblZoomFactor - 0.5) < 0.00001)
    Case ID_View_Size_25
        '25%
        Control.Checked = (Abs(mdblZoomFactor - 0.25) < 0.00001)
    Case ID_HELP_CONTENT
        '��������
    Case ID_HELP_CONTACT
        '���ͷ���
    Case ID_HELP_ONLINE
        '����ҽҵ
    Case ID_HELP_ABOUT
        '����...
    End Select
End Sub
Private Sub InitCommandBars()
    
    Dim BarPreview As CommandBar
    Dim cbp�ļ� As CommandBarPopup          '�ļ��˵�
    Dim cbp��ͼ As CommandBarPopup          '��ͼ�˵�
    Dim cbp���� As CommandBarPopup          '�����˵�

    '����λ�ûָ�
    Call RestoreWinState(Me, App.ProductName)
    Me.WindowState = vbMaximized
    '## �˵���ʼ��
    Dim cbpPopup As CommandBarPopup                     '��ʱ����
    Dim cbpPopupSub As CommandBarPopup                  '��ʱ����
    Dim objControl As CommandBarControl                 '�������ؼ�
    Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�
    Dim Combo As CommandBarComboBox                     '������������ؼ�
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    cbsThis.ActiveMenuBar.Title = "�˵���"
    Set cbp�ļ� = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�ļ�(&F)")
    With cbp�ļ�.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "��ӡ(&P)..."): objControl.IconId = 103
        Set objControl = .Add(xtpControlButton, ID_File_SaveAsPic, "���ΪͼƬ(&I)...")
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "�˳�(&X)"): objControl.IconId = 191
        objControl.BeginGroup = True
    End With
    
    Set cbp��ͼ = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "��ͼ(&V)")
    With cbp��ͼ.CommandBar.Controls
        Set cbpPopup = .Add(xtpControlPopup, 0, "������(&T)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, "�������б�"
        Set objControl = .Add(xtpControlButton, ID_View_StatusBar, "״̬��(&S)"): objControl.IconId = 702
        
        Set cbpPopup = .Add(xtpControlPopup, 0, "���ű���(&C)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_View_ActualSize, "ʵ�ʴ�С(&A)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_View_Size_75, "75%"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_View_Size_50, "50%"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_View_Size_25, "25%"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_View_ZoomIn, "�Ŵ�"): objControl.IconId = 502
        objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_View_ZoomOut, "��С"): objControl.IconId = 513
        Set objControl = .Add(xtpControlButton, ID_View_First, "��һҳ(&F)   "): objControl.BeginGroup = True: objControl.IconId = 7401
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_View_Prev, "ǰһҳ(&P)   "): objControl.IconId = 7402
        Set objControl = .Add(xtpControlButton, ID_View_Next, "��һҳ(&N)   "): objControl.IconId = 7403
        Set objControl = .Add(xtpControlButton, ID_View_Last, "���һҳ(&L) "): objControl.IconId = 7404
    End With
    
    Set cbp���� = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "����(&H)")
    With cbp����.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "��������(&H)")
        objControl.BeginGroup = True
        Set cbpPopupSub = .Add(xtpControlPopup, 0, "&Web�ϵ�" & gstrProductName)
        objControl.BeginGroup = True
        Set objControl = cbpPopupSub.CommandBar.Controls.Add(xtpControlButton, ID_HELP_ONLINE, gstrProductName & "����(&H)"): objControl.IconId = conMenu_Help_Web_Forum
        Set objControl = cbpPopupSub.CommandBar.Controls.Add(xtpControlButton, ID_HELP_CONTACT, "���ͷ���(&M)"): objControl.IconId = conMenu_Help_Web_Mail
        Set objControl = .Add(xtpControlButton, ID_HELP_ABOUT, "����(&A)..."): objControl.IconId = conMenu_Help_About
        objControl.BeginGroup = True
    End With
    
    Set BarPreview = cbsThis.Add("��ӡԤ��", xtpBarTop)
    With BarPreview.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "��ӡ"): objControl.IconId = 103
        objControl.Style = xtpButtonIconAndCaption
           
        Set objControl = .Add(xtpControlButton, ID_View_ActualSize, "ʵ�ʴ�С")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_View_ZoomIn, "�Ŵ�"): objControl.IconId = 502
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_View_ZoomOut, "��С"): objControl.IconId = 513
        Set Combo = .Add(xtpControlComboBox, ID_View_ZoomFactor, "���ű���")
        Combo.AddItem "100%", 1
        Combo.AddItem "75%", 2
        Combo.AddItem "50%", 3
        Combo.AddItem "25%", 4
        Combo.ListIndex = 1
        Combo.Width = 80
        Combo.DropDownWidth = 80
        Combo.DropDownListStyle = True
        
        Set objControl = .Add(xtpControlButton, ID_View_First, "��һҳ"): objControl.IconId = 7401
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_View_Prev, "ǰһҳ"): objControl.IconId = 7402
        Set objControl = .Add(xtpControlButton, ID_View_Next, "��һҳ"): objControl.IconId = 7403
        Set objControl = .Add(xtpControlButton, ID_View_Last, "���һҳ"): objControl.IconId = 7404
        
        Set objControl = .Add(xtpControlLabel, 0, "��ʼҳ��:")
        objControl.BeginGroup = True
        Set cboStartPage = .Add(xtpControlComboBox, ID_View_StartPage, "��ʼҳ��")
        cboStartPage.AddItem "�� 1 ҳ", 1
        cboStartPage.ListIndex = 1
        cboStartPage.Width = 80
        cboStartPage.DropDownWidth = 80
        cboStartPage.DropDownListStyle = True
        
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "�ر�(&Q)"): objControl.IconId = 191
        objControl.BeginGroup = True
    End With
    
    '�ȼ���
    cbsThis.KeyBindings.Add FCONTROL, Asc("P"), ID_FILE_PRINT
    cbsThis.KeyBindings.Add FCONTROL, Asc("Q"), ID_FILE_EXIT
    cbsThis.KeyBindings.Add FCONTROL, Asc("1"), ID_View_ActualSize
    
    cbsThis.KeyBindings.Add 0, VK_F1, ID_HELP_CONTENT
    cbsThis.KeyBindings.Add 0, vbKeyHome, ID_View_First
    cbsThis.KeyBindings.Add 0, vbKeyEnd, ID_View_Last
    cbsThis.KeyBindings.Add 0, vbKeyPageUp, ID_View_Prev
    cbsThis.KeyBindings.Add 0, vbKeyPageDown, ID_View_Next
    cbsThis.KeyBindings.Add 0, vbKeyAdd, ID_View_ZoomIn
    cbsThis.KeyBindings.Add 0, vbKeySubtract, ID_View_ZoomOut
    
    'TAB�ؼ��ĳ�ʼ��
    tabThis.Icons = zlCommFun.GetPubIcons
    tabThis.InsertItem 0, "ҳ������ͼ ", picҳ��.hWnd, 513
    
    With tabThis.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .ShowIcons = True
        .DisableLunaColors = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '����λ����Ϣ
    Call SaveWinState(Me, App.ProductName)
    pDetachMessages
    '�ֶ��ͷ��ڴ�
'    SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
    EmptyWorkingSet GetCurrentProcess()
    
    mlngCurPage = 0
    mlngStartPage = 0
    mlngPageCount = 0
    mlngBlankHeight = 0
    Set mTableThis = Nothing
    Set mcolMerge = Nothing
    Set mcolMergePic = Nothing
    Set mcolPage = Nothing
End Sub

Private Sub picҳ��_Resize()
    vfgҳ��.Move 0, 0, picҳ��.ScaleWidth, picҳ��.ScaleHeight
End Sub

Private Sub picback_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mlngX = x: mlngY = y
    If Button = 2 Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        Set Popup = cbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            .Add xtpControlButton, ID_View_Size_100, "100%"
            .Add xtpControlButton, ID_View_Size_75, "75%"
            .Add xtpControlButton, ID_View_Size_50, "50%"
            .Add xtpControlButton, ID_View_Size_25, "25%"
            Set Control = .Add(xtpControlButton, ID_View_ZoomIn, "�Ŵ�")
            Control.BeginGroup = True
            .Add xtpControlButton, ID_View_ZoomOut, "��С"
            
            Set Control = .Add(xtpControlButton, ID_File_SaveAsPic, "���ΪͼƬ(&I)...")
            Control.BeginGroup = True
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub picback_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If VS.Enabled Then
            If (y - mlngY) / 15 > 0 Then
                VS.Value = IIf(VS.Value - (y - mlngY) / 15 < VS.Min, VS.Min, VS.Value - (y - mlngY) / 15)
            Else
                VS.Value = IIf(VS.Value - (y - mlngY) / 15 > VS.Max, VS.Max, VS.Value - (y - mlngY) / 15)
            End If
        End If
        If HS.Enabled Then
            If (x - mlngX) / 15 > 0 Then
                HS.Value = IIf(HS.Value - (x - mlngX) / 15 < HS.Min, HS.Min, HS.Value - (x - mlngX) / 15)
            Else
                HS.Value = IIf(HS.Value - (x - mlngX) / 15 > HS.Max, HS.Max, HS.Value - (x - mlngX) / 15)
            End If
        End If
    End If
End Sub

Private Sub picPage_DblClick()
    Dim R As Double
    R = mdblZoomFactor + 0.25
    If R > 1# Then R = 0.25
    mdblZoomFactor = R
    PreviewPage mlngCurPage
End Sub

Private Sub picPage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mlngX = x: mlngY = y
    If Button = 1 Then Set picPage.MouseIcon = HS.MouseIcon
    If Button = 2 Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        Set Popup = cbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            .Add xtpControlButton, ID_View_Size_100, "100%"
            .Add xtpControlButton, ID_View_Size_75, "75%"
            .Add xtpControlButton, ID_View_Size_50, "50%"
            .Add xtpControlButton, ID_View_Size_25, "25%"
            Set Control = .Add(xtpControlButton, ID_View_ZoomIn, "�Ŵ�")
            Control.BeginGroup = True
            .Add xtpControlButton, ID_View_ZoomOut, "��С"
            
            Set Control = .Add(xtpControlButton, ID_File_SaveAsPic, "���ΪͼƬ(&I)...")
            Control.BeginGroup = True
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub picPage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If VS.Enabled Then
            If (y - mlngY) / 15 > 0 Then
                VS.Value = IIf(VS.Value - (y - mlngY) / 15 < VS.Min, VS.Min, VS.Value - (y - mlngY) / 15)
            Else
                VS.Value = IIf(VS.Value - (y - mlngY) / 15 > VS.Max, VS.Max, VS.Value - (y - mlngY) / 15)
            End If
        End If
        If HS.Enabled Then
            If (x - mlngX) / 15 > 0 Then
                HS.Value = IIf(HS.Value - (x - mlngX) / 15 < HS.Min, HS.Min, HS.Value - (x - mlngX) / 15)
            Else
                HS.Value = IIf(HS.Value - (x - mlngX) / 15 > HS.Max, HS.Max, HS.Value - (x - mlngX) / 15)
            End If
        End If
    End If
End Sub

Private Sub Reposition()
    VS.Top = picBack.Top
    VS.Left = ScaleWidth - VS.Width
    VS.Height = picBack.Height
    
    HS.Left = picBack.Left
    HS.Top = picBack.Top + picBack.Height
    HS.Width = picBack.Width
    
    '����Ԥ��ҳ
    
    If picBack.ScaleWidth >= picPage.Width + Shadow_W * 4 Then
        picPage.Left = (picBack.ScaleWidth - (picPage.Width + Shadow_W * 4)) / 2 + Shadow_W * 2
        picShadow.Left = picPage.Left + Shadow_W
        HS.Enabled = False
    Else
        HS.Max = (picPage.Width + Shadow_W * 4 - picBack.ScaleWidth) / 15
        If HS.Max / 3 < HS.SmallChange Then
            HS.LargeChange = HS.SmallChange
        Else
            HS.LargeChange = HS.Max / 3
        End If
        HS.Value = 0
        HS.Enabled = True
        HS_Change
    End If
    If picBack.ScaleHeight >= picPage.Height + Shadow_W * 4 Then
        picPage.Top = (picBack.ScaleHeight - (picPage.Height + Shadow_W * 4)) / 2 + Shadow_W
        picShadow.Top = picPage.Top + Shadow_W
        VS.Enabled = False
    Else
        VS.Max = (picPage.Height + Shadow_W * 4 - picBack.ScaleHeight) / 15
        If VS.Max / 3 < VS.SmallChange Then
            VS.LargeChange = VS.SmallChange
        Else
            VS.LargeChange = VS.Max / 3
        End If
        VS.Value = 0
        VS.Enabled = True
        VS_Change
    End If
End Sub

Private Sub picPage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Set picPage.MouseIcon = VS.MouseIcon
End Sub

Private Sub VS_Change()
    picPage.Top = -VS.Value * 15# + Shadow_W * 2
    picPage_Resize
    Me.Refresh
End Sub

Private Sub VS_Scroll()
    picPage.Top = -VS.Value * 15# + Shadow_W * 2
    picPage_Resize
    Me.Refresh
End Sub

Private Sub HS_Change()
    picPage.Left = -HS.Value * 15# + Shadow_W * 2
    picPage_Resize
    Me.Refresh
End Sub

Private Sub HS_Scroll()
    picPage.Left = -HS.Value * 15# + Shadow_W * 2
    picPage_Resize
    Me.Refresh
End Sub

Private Sub pAttachMessages()
'��Ϣ�����
    Subclass1.hWnd = Me.hWnd
    Subclass1.Messages(WM_MOUSEWHEEL) = True
    m_bSubClassing = True
End Sub

Private Sub pDetachMessages()
'ȡ����Ϣ����
    If (m_bSubClassing) Then
        Subclass1.Messages(WM_MOUSEWHEEL) = False
        m_bSubClassing = False
    End If
End Sub

Private Sub SearchCell(ByVal objCells As cTabCells, ByVal Row As Long, ByVal Col As Long, ByVal Erow As Long, ByVal ECol As Long, _
    W As Long, H As Long, strSkip As String)
'���ܣ���������һ����Ԫ���ʵ�ʿ�ߣ��������ϲ��ĵ�Ԫ��Ŀ�ߣ�
'      ��ҳ�ĺϲ���Ԫ�񣬼�¼������Ԫ��Ķ�Ӧ��ϵ
'������ row,col��ǰ��Ԫ�������
'       Erow,ECol=��ǰҳ���������
'���أ�W,H=�õ�Ԫ��Ŀ��(�����ϲ���Ԫ),strSkip=���ϲ��ĵ�Ԫ��:"[1,3],[3,4]......"
    Dim i As Long, j As Long
    Dim arrTmp As Variant, lngEnd As Long
    
    W = 0: H = 0
    With objCells.Cell(Row, Col)
        If .Merge Then
            If InStr(.MergeRange, ";") > 0 Then
                arrTmp = Split(Split(.MergeRange, ";")(1), ",")             '����"1,1;,3,3"
                For i = Row To Val(arrTmp(0))
                    For j = Col To Val(arrTmp(1))
                        If i <= Erow And j <= ECol Then strSkip = strSkip & ",[" & i & "," & j & "]"  '������ǰҳ�ĵ�Ԫ�����ռ�
                        If i = Row Then W = W + objCells.Cell(1, j).Width
                    Next
                    H = H + objCells.Cell(i, 1).Height
                Next
            Else    '��ҳ�ı��ϲ���Ԫ��,����Ԫ������һҳ,ͬҳ���ѱ�����
                H = .Height
                W = .Width
            End If
        Else
            H = .Height
            W = .Width
        End If
    End With
    If W > mPaper.AvailableWidth Then W = mPaper.AvailableWidth
    If H > mPaper.AvailableHeight Then H = mPaper.AvailableHeight
End Sub

Private Sub SetMergeRalation()
'���ܣ����ÿ�ҳ�ϲ�ʱ�����ϲ���Ԫ�����׵�Ԫ��Ķ�Ӧ��ϵ
    Dim i As Long, j As Long, p As Long, m As Long, N As Long
    Dim arrE As Variant, arrB As Variant, strSkip As String
    
    Set mcolMerge = New Collection
    With mTableThis.Cells
        For i = 1 To .Rows
            For j = 1 To .Cols
                If .Cell(i, j).Merge Then
                    If InStr(.Cell(i, j).MergeRange, ";") > 0 Then
                        p = GetPage(i, j)
                        arrB = Split(Split(.Cell(i, j).MergeRange, ";")(0), ",")
                        arrE = Split(Split(.Cell(i, j).MergeRange, ";")(1), ",")    '��������
                        If Val(arrE(0)) > mPages(p).Erow Then
                            '��ҳ�е�ÿ��
                            For m = mPages(p).Erow + 1 To Val(arrE(0))
                                For N = Val(arrB(1)) To Val(arrE(1))
                                    mcolMerge.Add i & "_" & j, "m" & m & "_" & N
                                    strSkip = strSkip & ",[" & m & "," & N & "]"
                                Next
                            Next
                        End If
                        If Val(arrE(1)) > mPages(p).ECol Then
                            '��ҳ�е�ÿ��
                            For N = mPages(p).ECol + 1 To Val(arrE(1))
                                For m = Val(arrB(0)) To Val(arrE(0))
                                    If InStr(strSkip, "[" & m & "," & N & "]") = 0 Then
                                        If m = 10 And N = 1 Then Stop
                                        mcolMerge.Add i & "_" & j, "m" & m & "_" & N
                                    End If
                                Next
                            Next
                        End If
                    End If
                End If
            Next
        Next
    End With
End Sub

Private Function CutText(objTarget As Object, ByVal strTxt As String, ByVal lngW As Long) As String
'���ܣ��������Ŀ�ȼ����ı�
    Dim strTmp As String, i As Double, ltwd As Double, ltest As Double
    Err.Clear
    On Error Resume Next
    
    ltest = objTarget.TextWidth(String(Len(strTxt), "A")) 'ģ��θ������,�ַ�����ʱ��������������쳣�����������
    
    ltwd = objTarget.TextWidth(strTxt)
    
    If ltwd < ltest Then '��ģ���ȴ��ڼ����ȣ���������ʧʵ
        ltwd = lngW + 1
    End If
    
    If Err.Number <> 0 Then
        ltwd = lngW + 1: Err.Clear
    End If
    If ltwd <= lngW Then
        CutText = strTxt
    Else
        For i = 1 To Len(strTxt)
            strTmp = Mid(strTxt, 1, i)
            If objTarget.TextWidth(strTmp) > lngW + 15 Then '��15��ʾ����һ������֮�ڵ����
                strTmp = Mid(strTxt, 1, i - 1)
                Exit For
            End If
        Next
        CutText = strTmp
    End If
End Function

Private Sub SetLine(objTarget As Object, lngType As Long)
'���ܣ�����������ʽ
'������
    Select Case lngType
        Case 4  '����
            objTarget.DrawWidth = 1
            objTarget.DrawStyle = vbDot 'vbDash
        Case 5  '����
            objTarget.DrawWidth = 2
            objTarget.DrawStyle = vbSolid
        Case Else
            objTarget.DrawWidth = 1
            objTarget.DrawStyle = vbSolid
    End Select
End Sub

Private Function DrawCellLine(objTarget As Object, objCell As cTabCell, ByVal x As Long, ByVal y As Long, _
    ByVal W As Long, ByVal H As Long, Optional ByVal blnMerged As Boolean, Optional ByVal blnMerge As Boolean)
'���ܣ������Ԫ��ı���,�߿�߰������ϲ��ĵ�Ԫ��Ŀ��
'      ������Ԫ��ͬҳ�ı��ϲ���Ԫ���ѱ���������ҳ�ı��ϲ���Ԫ���������Ԫ�����ʽ���
'������blnMerged=�Ƿ��ǿ�ҳ�ı��ϲ���Ԫ��
'      blnMerge=��ҳ�ϲ���Ԫ�������Ԫ��
    Dim arrGrid As Variant, arrRange As Variant, blndo As Boolean
    Dim lngMerge As Long '���ң����£����������ֺϲ���ʽ
    
    If blnMerged Then
        arrGrid = Split(mcolMerge("m" & objCell.Row & "_" & objCell.Col), "_")
        With mTableThis.Cells.Cell(Val(arrGrid(0)), Val(arrGrid(1)))
            arrRange = Split(.MergeRange, ";")                   '����"1,1;,3,3"
            If .CellLineTop <> 0 Then
                If objCell.Row = Split(arrRange(0), ",")(0) Then '�뿪ʼ������ͬ�Ż�����
                    Call SetLine(objTarget, .CellLineTop)
                    objTarget.Line (x, y)-(x + W, y), .CellLineTopColor
                End If
            End If
            If .CellLineBottom <> 0 Then
                If objCell.Row = Split(arrRange(1), ",")(0) Then '�����������ͬ�Ż�����
                    Call SetLine(objTarget, .CellLineBottom)
                    objTarget.Line (x, y + H)-(x + W, y + H), .CellLineBottomColor
                End If
            End If
            If .CellLineLeft <> 0 Then
                If objCell.Col = Split(arrRange(0), ",")(1) Then '�뿪ʼ������ͬ�Ż�����
                    Call SetLine(objTarget, .CellLineLeft)
                    objTarget.Line (x, y)-(x, y + H), .CellLineLeftColor
                End If
            End If
            If .CellLineRight <> 0 Then
                If objCell.Col = Split(arrRange(1), ",")(1) Then '�����������ͬ�Ż�����
                    Call SetLine(objTarget, .CellLineRight)
                    objTarget.Line (x + W, y)-(x + W, y + H), .CellLineRightColor
                End If
            End If
        End With
    Else
        With objCell
            If blnMerge Then  '��ҳ�ϲ���Ԫ�������Ԫ��
                arrRange = Split(Split(.MergeRange, ";")(1), ",")
                If Val(arrRange(0)) = .Row Then
                    If Val(arrRange(1)) > .Col Then lngMerge = 0     '���Һϲ�ʱ�����ұ���
                Else
                    If Val(arrRange(1)) > .Col Then
                        lngMerge = 2    '�����²����ҡ��ױ���
                    Else
                        lngMerge = 1    '���²�������
                    End If
                End If
            Else
                lngMerge = -1
            End If
            
            If .CellLineTop <> 0 Then
                blndo = True
                If .CellLineTop = 4 Then
                    If objCell.Row > 1 And .Merge = False Then
                        '�����Ǳ��ϲ���Ԫ��(�Ҳ����׵�Ԫ��)���򲻴����ߣ��������Ϊ��㲻ͬ���ش���Ϊʵ��
                        If mTableThis.Cells.Cell(objCell.Row - 1, objCell.Col).Merge And InStr(mTableThis.Cells.Cell(objCell.Row - 1, objCell.Col).MergeRange, ";") = 0 Then
                           blndo = False
                        End If
                    End If
                End If
                If blndo Then
                    Call SetLine(objTarget, .CellLineTop)
                    objTarget.Line (x, y)-(x + W, y), .CellLineTopColor
                End If
            End If
            If .CellLineBottom <> 0 And lngMerge < 1 Then
                blndo = True
                If .CellLineBottom = 4 Then
                    If objCell.Row < mTableThis.Cells.Rows And .Merge = False Then
                        '�����Ǳ��ϲ���Ԫ��(�Ҳ����׵�Ԫ��)���򲻴����ߣ��������Ϊ��㲻ͬ���ش���Ϊʵ��
                        If mTableThis.Cells.Cell(objCell.Row + 1, objCell.Col).Merge And InStr(mTableThis.Cells.Cell(objCell.Row + 1, objCell.Col).MergeRange, ";") = 0 Then
                           blndo = False
                        End If
                    End If
                End If
                If blndo Then
                    Call SetLine(objTarget, .CellLineBottom)
                    objTarget.Line (x, y + H)-(x + W, y + H), .CellLineBottomColor
                End If
            End If
            If .CellLineLeft <> 0 Then
                blndo = True
                If .CellLineLeft = 4 Then
                    If objCell.Col > 1 And .Merge = False Then
                        '�ұ��Ǳ��ϲ���Ԫ��(�Ҳ����׵�Ԫ��)���������Ϊ��㲻ͬ���ش���Ϊʵ��
                        If mTableThis.Cells.Cell(objCell.Row, objCell.Col - 1).Merge And InStr(mTableThis.Cells.Cell(objCell.Row, objCell.Col - 1).MergeRange, ";") = 0 Then
                           blndo = False
                        End If
                    End If
                End If
                If blndo Then
                    Call SetLine(objTarget, .CellLineLeft)
                    objTarget.Line (x, y)-(x, y + H), .CellLineLeftColor
                End If
            End If
            If .CellLineRight <> 0 And (lngMerge = -1 Or lngMerge = 1) Then
                blndo = True
                If .CellLineRight = 4 Then
                    If objCell.Col < mTableThis.Cells.Cols And .Merge = False Then
                        '����Ǳ��ϲ���Ԫ��(�Ҳ����׵�Ԫ��)���������Ϊ��㲻ͬ���ش���Ϊʵ��
                        If mTableThis.Cells.Cell(objCell.Row, objCell.Col + 1).Merge And InStr(mTableThis.Cells.Cell(objCell.Row, objCell.Col + 1).MergeRange, ";") = 0 Then
                           blndo = False
                        End If
                    End If
                End If
                If blndo Then
                    Call SetLine(objTarget, .CellLineRight)
                    objTarget.Line (x + W, y)-(x + W, y + H), .CellLineRightColor
                End If
            End If
        End With
    End If
    
End Function

Private Sub GetMergeCellWH(ByVal lngPage As Long, ByVal objCell As cTabCell, ByVal blnAll As Boolean, _
        lngW As Single, lngAW As Single, lngH As Single, lngAH As Single)
'���ܣ���ȡһ����ҳ�ϲ���Ԫ��Ŀ��
'������blnAll=�Ƿ�������е�Ԫ�񣬰�����ҳ��
'���أ���Ԫ���δ��ҳ�Ϳ�ҳ��ȫ�����
    Dim i As Long
    Dim arrMerge As Variant
    
    arrMerge = Split(Split(objCell.MergeRange, ";")(1), ",")
    For i = objCell.Row To Val(arrMerge(0))
        If i > mPages(lngPage).Erow And Not blnAll Then Exit For
        lngH = lngH + mTableThis.Cells.Cell(i, 1).Height
        lngAH = lngAH + mTableThis.Cells.Cell(i, 1).Height * mRatioY
    Next
    For i = objCell.Col To Val(arrMerge(1))
        If i > mPages(lngPage).ECol And Not blnAll Then Exit For
        lngW = lngW + mTableThis.Cells.Cell(1, i).Width
        lngAW = lngAW + mTableThis.Cells.Cell(1, i).Width * mRatioX
    Next
    
    If lngW > mPaper.AvailableWidth Then lngW = mPaper.AvailableWidth
    If lngH > mPaper.AvailableHeight Then lngH = mPaper.AvailableHeight
    If lngAW > mPaper.AvailableWidth * mRatioX Then lngAW = mPaper.AvailableWidth * mRatioX
    If lngAH > mPaper.AvailableHeight * mRatioY Then lngAH = mPaper.AvailableHeight * mRatioY
End Sub

Private Function DrawCell(objTarget As Object, ByVal objCell As cTabCell, ByVal x As Single, ByVal y As Single, _
    ByVal W As Single, ByVal H As Single, ByVal lngPage As Long, Optional blnSimulate As Boolean) As Boolean
'���ܣ���ָ���豸�ϰ�ָ����ʽ��������ֻ�ͼ��
'������
'   objTarget=����豸,ΪPrinter��PictureBox����
'   objCell=����ĵ�Ԫ�����е��������������ͣ��ı���ͼƬ(stdPicture)
'   x,y=������ݵ���ʼ��������
'   w,h=��Ԫ���ʵ�ʿ�ߣ��������ϲ��ĵ�Ԫ������,��λ��ת��Ϊ��ӡ�����λ(�)
'   lngPage=��ǰ��Ԫ��������ҳ�ţ���0��ʼ
'   blnSimulate=ģ����������������ҳ�ϲ����׵�Ԫ��ͼƬ�����漯��
    Dim lngX As Single, lngY As Single
    Dim lngAW As Single, lngAH As Single            '��Ԫ���ȥ���������ֵļ���Ŀ��ÿ��
    Dim lngW As Single, lngH As Single              '��ҳ�ϲ����׵�Ԫ���ڵ�ǰҳ��ʵ�ʿ��
        
    Dim arrMerge As Variant
    Dim blnMerge As Boolean                     '��ҳ�ϲ���Ԫ�������Ԫ��,�������ͼƬ�����з֣�����Ǵ�ӡ����Ŵ������ʵ�ʴ�ӡʱ����С���Ա��ֲ�ʧ��
    Dim blnMerged As Boolean                    '��ҳ�ϲ��ı��ϲ���Ԫ��,û��ҳ���ѱ�skip�˲����������
    Dim lngLineW As Single                        '���������ֵļ��
    Dim i As Long
    Dim picTmp As StdPicture
    
    Dim lngProw As Long, lngPcol As Long, lngPPage As Long  '��ҳ�ϲ��׵�Ԫ����Ϣ
        
    On Error GoTo errH
    
    If objCell.Merge Then
        If InStr(objCell.MergeRange, ";") > 0 Then                                  '��ҳ�ϲ���Ԫ����׵�Ԫ��
            arrMerge = Split(Split(objCell.MergeRange, ";")(1), ",")
            blnMerge = Val(arrMerge(0)) > mPages(lngPage).Erow Or Val(arrMerge(1)) > mPages(lngPage).ECol
        End If
        blnMerged = InStr(objCell.MergeRange, ";") = 0                              '��ҳ�ϲ��ı��ϲ���Ԫ��
    End If
    
        
    '��ҳ�ı��ϲ���Ԫ��ֱ�Ӵӻ���ͼƬ���г�
    If blnMerged Then
        objTarget.CurrentX = x
        objTarget.CurrentY = y
        arrMerge = Split(mcolMerge("m" & objCell.Row & "_" & objCell.Col), "_")
        lngProw = Val(arrMerge(0))
        lngPcol = Val(arrMerge(1))
        
        '�ϲ���Ԫ��ͼƬ�ڵ���ʼ����
        lngX = 0: lngY = 0
        For i = arrMerge(0) To objCell.Row - 1
            lngY = lngY + mTableThis.Cells.Cell(i, 1).Height * mRatioY
        Next
        For i = arrMerge(1) To objCell.Col - 1
            lngX = lngX + mTableThis.Cells.Cell(1, i).Width * mRatioX
        Next
        
        On Error Resume Next
        Set picTmp = mcolMergePic("m" & lngProw & "_" & lngPcol)
        '��ָ��ҳ������������ʼҳ���ǵ�һҳʱ����ҳ�ϲ��ĵ�Ԫ��û�����Ȳ���ͼƬ
        If Err.Number <> 0 Then
            lngPPage = GetPage(lngProw, lngPcol)
            Call GetMergeCellWH(lngPPage, mTableThis.Cells.Cell(lngProw, lngPcol), True, lngW, lngAW, lngH, lngAH)   'lngAW,lngAH��ʹ��
            Call DrawCell(objTarget, mTableThis.Cells.Cell(lngProw, lngPcol), 0, 0, lngW, lngH, lngPPage, True)
            Set picTmp = mcolMergePic("m" & lngProw & "_" & lngPcol)
        End If
        
        If picTmp.Handle <> 0 Then objTarget.PaintPicture picTmp, x, y, W, H, lngX, lngY, W * mRatioX, H * mRatioY
        On Error GoTo 0
    Else
        If TypeName(objTarget) = "Printer" And Not blnMerge Then
            lngLineW = Printer.TwipsPerPixelX * 2   'ÿ�����أ���ĻΪ15�,��ӡ��Ϊ2.4�
        Else
            lngLineW = Screen.TwipsPerPixelX * 2
        End If
        lngAW = W - lngLineW * 2
        lngAH = H - lngLineW * 2
        If blnMerge Then
            picMerge.Cls
            picMerge.Move x, y, W * mRatioX, H * mRatioY
        End If
        
        
        If objCell.�������� = CellTypeEnum.cprCTReportPic Or objCell.�������� = CellTypeEnum.cprCTPicture Then
            If blnMerge Then                                                            '��ҳ�ϲ��ĵ�Ԫ���Ȱ����ֻ�����ʱͼƬ�������и�
                Call DrawCellPic(picMerge, objCell, 0, 0, W * mRatioX, H * mRatioY, lngAW * mRatioX, lngAH * mRatioY, lngLineW * mRatioX, blnMerge)
            Else
                Call DrawCellPic(objTarget, objCell, x, y, W, H, lngAW, lngAH, lngLineW)
            End If
        Else
            If blnMerge Then
                Call DrawCellText(picMerge, objCell, 0, 0, W * mRatioX, H * mRatioY, lngAW * mRatioX, lngAH * mRatioY, lngLineW * mRatioX, blnMerge)
            Else
                Call DrawCellText(objTarget, objCell, x, y, W, H, lngAW, lngAH, lngLineW)
            End If
        End If
        
        '��ҳ�ϲ����׵�Ԫ���и�ͼƬ
        If blnMerge Then
            lngW = 0: lngH = 0      '����ֱ����w,h��Ϊ��ߵ��������Ϊ�����ܿ�ҳ
            lngAW = 0: lngAH = 0
            Call GetMergeCellWH(lngPage, objCell, False, lngW, lngAW, lngH, lngAH)
                        
            If Not blnSimulate Then objTarget.PaintPicture picMerge.Image, x, y, lngW, lngH, 0, 0, lngAW, lngAH
            mcolMergePic.Add picMerge.Image, "m" & objCell.Row & "_" & objCell.Col
        End If
    End If
    
    '�����(��Ϊ��ҳ�ϲ��ĵ�Ԫ��ͼƬ�ǰ���Ԫ������)
    If Not blnSimulate Then
        If blnMerge Then
            Call DrawCellLine(objTarget, objCell, x, y, lngW, lngH, False, blnMerge)
        Else
            Call DrawCellLine(objTarget, objCell, x, y, W, H, blnMerged)
        End If
    End If
        
    DrawCell = True
    Exit Function
errH:
    DrawCell = False
    MsgBox "���" & "[" & objCell.Row & "," & objCell.Col & "]�����쳣��" & vbCrLf & Err.Description
End Function

Private Sub DrawCellPic(objTarget As Object, objCell As cTabCell, ByVal x As Single, ByVal y As Single, _
        ByVal W As Single, ByVal H As Single, ByVal lngAW As Single, ByVal lngAH As Single, ByVal lngLineW As Single, Optional ByVal blnMerge As Boolean)
'���ܣ������Ԫ��ͼƬ
'������lngAW,lngAH=��Ԫ���ȥ���������ֵļ���Ŀ��ÿ��
'      w,h=��Ԫ���ʵ�ʿ�ߣ��������ϲ��ĵ�Ԫ������,��λ��ת��Ϊ��ӡ�����λ(�)
'      lngLineW=�߿������ֵļ��,2����λֵ
'      blnMerge=�Ƿ�Ϊ�ϲ���Ԫ����׵�Ԫ��
    Dim lngX As Single, lngY As Single
    Dim lngW As Single, lngH As Single, lngSpace As Single
    Dim picTmp As StdPicture
    
    picOrig.Cls
    Set picTmp = mTableThis.CellPicture(picOrig, objCell.Key, W, H) 'picOrig����������ͼƬ�ϻ没�����
    If picTmp.Handle = 0 Then Exit Sub
    lngW = Me.ScaleX(picTmp.Width, vbHimetric, vbTwips) * IIf(blnMerge, mRatioX, 1)
    lngH = Me.ScaleY(picTmp.Height, vbHimetric, vbTwips) * IIf(blnMerge, mRatioY, 1)
            
    Select Case objCell.HAlignment
        Case HALignLeft
            lngX = x + lngLineW
        Case HAlignCenter
            lngSpace = lngAW - lngW
            If lngSpace < 0 Then lngSpace = 0
            lngX = x + lngLineW + lngSpace / 2
        Case HALignRight
            lngSpace = W - lngLineW - lngW
            If lngSpace < 0 Then lngSpace = 0
            lngX = x + lngLineW + lngSpace
    End Select
    Select Case objCell.VAlignment
        Case VALignTop
            lngY = y + lngLineW
        Case VAlignCenter
            lngSpace = lngAH - lngH
            If lngSpace < 0 Then lngSpace = 0
            lngY = y + lngLineW + lngSpace / 2
        Case VALignBottom
            lngSpace = H - lngLineW - lngH
            If lngSpace < 0 Then lngSpace = 0
            lngY = y + lngLineW + lngSpace
    End Select
    
    If (lngW > lngAW Or lngH > lngAH) Then  '�������ʱ�Զ�����
        objTarget.PaintPicture picTmp, lngX, lngY, lngAW, lngAH
    Else
        objTarget.PaintPicture picTmp, lngX, lngY, lngW, lngH
    End If
End Sub

Private Sub DrawCellText(objTarget As Object, objCell As cTabCell, ByVal x As Single, ByVal y As Single, _
        ByVal W As Single, ByVal H As Single, ByVal lngAW As Single, ByVal lngAH As Single, ByVal lngLineW As Single, Optional ByVal blnMerge As Boolean)
'���ܣ������Ԫ������
'������lngAW,lngAH=��Ԫ���ȥ���������ֵļ���Ŀ��ÿ��
'      w,h=��Ԫ���ʵ�ʿ�ߣ��������ϲ��ĵ�Ԫ������,��λ��ת��Ϊ��ӡ�����λ(�)
'      lngLineW=�߿������ֵļ��,2����λֵ
'      blnMerge=�Ƿ�Ϊ�ϲ���Ԫ����׵�Ԫ��
    Dim strText As String, strLeave As String
    Dim arrText As Variant, arrText2 As Variant
    
    Dim lngX As Single, lngY As Single
    Dim lngTxtW As Single, lngTxtH As Single, lngSpace As Single
    Dim i As Long, j As Long
    
    With objCell
        strText = mTableThis.CellContent(objCell.Key)
        strText = Replace(strText, vbCrLf, vbCr)          '���ⲿ������ı������е�����vbCr��vbLf�����
        strText = Replace(strText, vbLf, vbCr)
        arrText = Split(strText, vbCr)
        Call SetObjFontFormat(objTarget, objCell, blnMerge)
        
        lngTxtH = objTarget.TextHeight("�ı�")   '�����������и�
                
        '�Զ�����
        arrText2 = Array()
        For i = 0 To UBound(arrText)
            strText = arrText(i)
            
            Do
                j = UBound(arrText2) + 1
                If lngTxtH * (j + 1) > lngAH Then Exit For '���ߺ������
                
                ReDim Preserve arrText2(j)
                strLeave = CutText(objTarget, strText, lngAW)
                If InStr("����������������������", Replace(strText, strLeave, "")) > 0 And strText <> strLeave Then 'Ϊ�˱��ֺͱ༭�ؼ�һ�£���������ֻ��һ�������ŵ����
                    arrText2(j) = Mid(strLeave, 1, Len(strLeave) - 1)
                Else
                    arrText2(j) = strLeave
                End If
                strText = Mid(strText, Len(arrText2(j)) + 1)
            Loop While strText <> ""
        Next
        
        lngSpace = 0
        If .VAlignment = VAlignCenter Then
            lngSpace = lngAH - lngTxtH * (UBound(arrText2) + 1)
        ElseIf .VAlignment = VALignBottom Then
            lngSpace = H - lngLineW - lngTxtH * (UBound(arrText2) + 1)
        End If
        If lngSpace < 0 Then lngSpace = 0
        Select Case .VAlignment
            Case VALignTop
                lngY = y + lngLineW
            Case VAlignCenter
                lngY = y + lngLineW + lngSpace / 2
            Case VALignBottom
                lngY = y + lngLineW + lngSpace
        End Select
        
        For i = 0 To UBound(arrText2)
            strText = arrText2(i)
            lngTxtW = objTarget.TextWidth(strText)
            
            '�ı����뷽ʽ�����������ʼλ��
            Select Case .HAlignment
                Case HALignLeft
                    lngX = x + lngLineW
                Case HAlignCenter
                    lngX = x + lngLineW + (lngAW - lngTxtW) / 2
                Case HALignRight
                    lngX = x + lngLineW + (W - lngLineW - lngTxtW)
            End Select
            objTarget.CurrentX = lngX
            objTarget.CurrentY = lngY
            objTarget.Print strText
            lngY = lngY + lngTxtH
        Next
    End With
End Sub

Private Sub SetObjFontFormat(objTarget As Object, objCell As cTabCell, blnMerge As Boolean)
    With objCell
        objTarget.ForeColor = .FontColor
        objTarget.Font.Name = .FontName
        objTarget.Font.Italic = .FontItalic
        objTarget.Font.Size = Val(.FontSize * IIf(blnMerge, mRatioY, 1))
        objTarget.Font.Underline = .FontUnderline
        objTarget.Font.Bold = .FontBold
        objTarget.Font.Strikethrough = .FontStrikeout
    End With
End Sub

Private Sub DrawPage(ByVal lngPage As Long, ByVal lngBlankHeight As Long, ByRef objTarget As Object)
'���ܣ����ĳҳ���ݵ�ָ���豸
'������ lngPage:ָ�������ҳ��,��һҳ��1����0
'       objTarget:����pic���ӡ��
'       lngBlankHeight:��ʼҳ�Ϸ��Ŀհ�����߶�,��ӡԤ��ʱ�����������PreviewPage�е�������͸������

    Dim i As Long, j As Long
    Dim x As Long, y As Long                        '��ӡ����ʼλ��
    Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, lineW As Long
    Dim lngPicWidth As Long, lngPicHeight As Long   'ҳüͼƬ��ȸ߶�
    Dim strHead As String, strFoot As String, strText As String, arrTxt As Variant, strTmp As String, strlfSplit As String
    Dim blnPrint As Boolean, lngOffsetX As Long, lngOffsetY As Long
    
    Dim strSkip As String                           '���ϲ��ĵ�Ԫ���������,"[1,3],[3,4]......"
    Dim lngW As Long, lngH As Long                  '��Ԫ��ʵ�ʵ������ߣ������ϲ��ĵ�Ԫ��Ŀ�ߣ���������ֽ�ŵ������
        
    Const CGRAY = &H8000000B                        'Ԥ��ʱҳ�߾�Ļ�ɫ
            
    objTarget.ScaleMode = vbTwips                   '���ô�ӡ������ʾ���������λΪ�,ÿ�����أ���ĻΪ15�,��ӡ��Ϊ2.4�
    blnPrint = TypeName(objTarget) = "Printer"
    
    With mTableThis.EPRFileInfo
        '0.�����Ԥ�����򻭳�ҳ�߾�ľ����ͻ�ʵ��
        If Not blnPrint Then
            lineW = Screen.TwipsPerPixelX * 1       '�߿�Ϊһ������,ת��Ϊ���
            X1 = .MarginLeft - lineW
            Y1 = .MarginTop - lineW
            X2 = mPaper.PaperWidth - .MarginRight + lineW
            Y2 = mPaper.PaperHeight - .MarginBottom + lineW
        
            
            objTarget.DrawWidth = 1
            objTarget.DrawStyle = vbDot 'vbSolid
            
            objTarget.Line (0, Y1)-(mPaper.PaperWidth, Y1), CGRAY              '�Ϻ���
            objTarget.Line (0, Y2 - lineW)-(mPaper.PaperWidth, Y2 - lineW), CGRAY     '�º���
            objTarget.Line (X1, 0)-(X1, mPaper.PaperHeight), CGRAY                     '������
            objTarget.Line (X2 - lineW, 0)-(X2 - lineW, mPaper.PaperHeight), CGRAY     '������
            
            '�ٻ���ҳüҳ�ҵ���
            objTarget.DrawWidth = 1
            objTarget.DrawStyle = vbDot
            objTarget.Line (0, .HeadMargin)-(mPaper.PaperWidth, .HeadMargin), CGRAY            '�Ϻ���
            objTarget.Line (0, Y2 + .MarginTop)-(mPaper.PaperWidth, Y2 + .MarginTop), CGRAY    '�º���
        Else
            lngOffsetX = objTarget.ScaleX(GetDeviceCaps(objTarget.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
            lngOffsetY = objTarget.ScaleY(GetDeviceCaps(objTarget.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)
        End If
        
        lngPicWidth = 0: lngPicHeight = 0
        x = IIf(mTableThis.EPRFileInfo.MarginLeft - lngOffsetX >= 0, mTableThis.EPRFileInfo.MarginLeft - lngOffsetX, lngOffsetX) 'ҳü�����ұ߾���ҳ����ͬ
        y = IIf(.HeadMargin - lngOffsetY >= 0, .HeadMargin - lngOffsetY, lngOffsetY)                     'ҳü���ϱ߾�
    
        If lngBlankHeight = 0 Then                   '��������ҳ�հ����򣬱�ʾ���򣬲���ӡҳü
            '1.���ҳü(ҳüͼƬ��Ŀǰ�̶���������Ͻ�)
            If Not (.HeadPic Is Nothing) Then
                If .HeadPic.Handle <> 0 Then
                    lngPicWidth = Me.ScaleX(.HeadPic.Width, vbHimetric, vbTwips)
                    lngPicHeight = Me.ScaleY(.HeadPic.Height, vbHimetric, vbTwips)
                    
                    If y + lngPicHeight > .MarginTop Then lngPicHeight = .MarginTop - y   '�߶Ȳ�����ҳü�߶�
                    If lngPicWidth > mPaper.AvailableWidth Then lngPicWidth = mPaper.AvailableWidth
                    
                    If lngPicWidth > 0 And lngPicHeight > 0 Then
                        objTarget.PaintPicture .HeadPic, x, y, lngPicWidth, lngPicHeight
                        x = x + lngPicWidth
                    End If
                End If
            End If
    
            '2.���ҳü����
            If .HeadConText <> "" Then
                objTarget.ForeColor = .HeadFontColor
                objTarget.Font.Name = .HeadFontName
                objTarget.Font.Italic = .HeadFontItalic
                objTarget.Font.Size = .HeadFontSize
                objTarget.Font.Underline = .HeadFontUnderline
                objTarget.Font.Bold = .HeadFontBold
                objTarget.Font.Strikethrough = .HeadFontStrikethrough
                    
                strHead = Replace(mTableThis.GetHeadFootContent(0), "[ҳ��]", lngPage)
                strHead = Replace(strHead, "[��ҳ��]", mlngPageCount)
                lngW = mPaper.AvailableWidth - lngPicWidth - (objTarget.TextWidth("��") / 3) * 2
                lngH = objTarget.TextHeight("�ı�")
                strHead = Replace(strHead, vbCr, vbLf)                  '���ⲿ������ı������е�����vbCr��vbLf�����
                strHead = Replace(strHead, vbLf & vbLf, vbLf)
                
                For i = 1 To Len(strHead)
                    strTmp = Mid(strHead, 1, i)
                    If objTarget.TextWidth(strTmp) >= lngW Or InStr(strTmp, vbLf) > 0 Then
                        strlfSplit = strlfSplit & Mid(strHead, 1, i) & IIf(InStr(strTmp, vbLf) > 0, "", vbLf)
                        strHead = Mid(strHead, i + 1)
                        i = 0
                    End If
                Next
                strHead = strlfSplit & strHead
                            
                arrTxt = Split(strHead, vbLf)
                '����ҳü
                For i = 0 To UBound(arrTxt)
                    If lngH * (i + 1) > .MarginTop - .HeadMargin Then Exit For '�����������
                    strText = CutText(objTarget, CStr(arrTxt(i)), lngW)
                    objTarget.CurrentX = x
                    objTarget.CurrentY = y + lngH * i
                    objTarget.Print strText
                Next
            End If
        End If
    End With
    
    
    '3.���ÿ����Ԫ��
    '���ȫ������ҳ�հ������ڣ������,������������ڣ�����������������հ����򸲸�)
    If lngBlankHeight < mPaper.PaperHeight - mTableThis.EPRFileInfo.MarginBottom Then
        With mPages(lngPage - 1)
            y = IIf(mTableThis.EPRFileInfo.MarginTop - lngOffsetY >= 0, mTableThis.EPRFileInfo.MarginTop - lngOffsetY, lngOffsetY)    '���Ϊ����ҳ�߾�
            For i = .BRow To .Erow
                x = IIf(mTableThis.EPRFileInfo.MarginLeft - lngOffsetX >= 0, mTableThis.EPRFileInfo.MarginLeft - lngOffsetX, lngOffsetX)
                For j = .BCol To .ECol
                    '1.ͬһҳ�ĺϲ���Ԫ��ֻ�����һ��
                    '2.��ҳ�ĺϲ���Ԫ���ڵ�һ����Ԫ��ʱ�������ͼƬ�����з�
                    If InStr(strSkip, "[" & i & "," & j & "]") = 0 Then
                        lngW = 0: lngH = 0
                        Call SearchCell(mTableThis.Cells, i, j, .Erow, .ECol, lngW, lngH, strSkip)
                        
                        If lngW <> 0 And lngH <> 0 Then
                            If Not DrawCell(objTarget, mTableThis.Cells.Cell(i, j), x, y, lngW, lngH, lngPage - 1) Then Exit Sub
                        End If
                    End If
                    x = x + mTableThis.Cells.Cell(1, j).Width
                Next
                y = y + mTableThis.Cells.Cell(i, 1).Height
            Next
        End With
    End If
    
    '4.���ҳ�ţ�ֻ����û�������ڸ�����ʱ�Ŵ�ӡ(�����������СʱҲ��Ϊû������)
    If lngBlankHeight < 200 Then
        With mTableThis.EPRFileInfo
            If .FootConText <> "" Then
                objTarget.ForeColor = .FootFontColor
                objTarget.Font.Name = .FootFontName
                objTarget.Font.Italic = .FootFontItalic
                objTarget.Font.Size = .FootFontSize
                objTarget.Font.Underline = .FootFontUnderline
                objTarget.Font.Bold = .FootFontBold
                objTarget.Font.Strikethrough = .FootFontStrikethrough
                    
                strFoot = Replace(mTableThis.GetHeadFootContent(1), "[ҳ��]", lngPage)
                strFoot = Replace(strFoot, "[��ҳ��]", mlngPageCount)
                lngW = mPaper.AvailableWidth + lngOffsetX - (objTarget.TextWidth("��") / 3) * 2
                lngH = objTarget.TextHeight("�ı�")
                strFoot = Replace(strFoot, vbCr, vbLf)                  '���ⲿ������ı������е�����vbCr��vbLf�����
                strFoot = Replace(strFoot, vbLf & vbLf, vbLf)
                
                For i = 1 To Len(strFoot)
                    strTmp = Mid(strFoot, 1, i)
                    If objTarget.TextWidth(strTmp) >= lngW Or InStr(strTmp, vbLf) > 0 Then
                        strlfSplit = strlfSplit & Mid(strFoot, 1, i) & IIf(InStr(strTmp, vbLf) > 0, "", vbLf)
                        strFoot = Mid(strFoot, i + 1)
                        i = 0
                    End If
                Next
                strFoot = strlfSplit & strFoot
                
                arrTxt = Split(strFoot, vbLf)
                x = IIf(.MarginLeft - lngOffsetX >= 0, .MarginLeft - lngOffsetX, lngOffsetX)
                y = mPaper.PaperHeight - .MarginBottom - lngOffsetY
                '����ҳ��
                For i = 0 To UBound(arrTxt)
                    If lngH * (i + 1) > .MarginBottom - .FootMargin Then Exit For  '�����������
                    strText = CutText(objTarget, arrTxt(i), lngW)
                    objTarget.CurrentX = x
                    objTarget.CurrentY = y + lngH * i
                    objTarget.Print strText
                Next
            End If
        End With
    End If
    
    '5.�����ʼҳ�����հ����򣬷ŵ�����Ը��ǲ��ֵ�Ԫ��
    If lngBlankHeight > 0 Then
        objTarget.PaintPicture picPrintBlank.Image, 0, 0, objTarget.Width, lngBlankHeight
    End If
End Sub


Private Function PrintTable(Optional ByVal strPrintDeviceName As String) As Boolean
'���ܣ�  ��ӡ��ǰ��񵽴�ӡ��
    If Not ExistsPrinter Then MsgBox "û�а�װ��ӡ�豸�����ܴ�ӡ��", vbExclamation, App.Title: Exit Function
        
    Dim strOldPrinterName As String, lngOldPaperKind As Long
    Dim intPageFrom As Integer, intPageTo As Integer
    Dim bytPageOddEven As Byte                          '��żҳ����
    Dim intCopies As Integer, blnCopyOrder As Boolean   '����,��ݴ�ӡ
    Dim T As Variant, aryPage() As String, i As Long, j As Long, k As Long, l As Long, m As Long
    Dim lngPageCount As Long
    Dim Pages() As Long                                 '��ӡ��Χ�ڵ��������ӡ��ҳ��
    Dim blnRangePrint As Boolean                        '�Ƿ���ҳ�뷶Χ��ӡ
    Dim blnHave As Boolean, blnFirstPrinted As Boolean
    Dim p As Integer, lngNumber As Long
    Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single
    
    X1 = GetDeviceCaps(Printer.hdc, LOGPIXELSX)
    X2 = GetDeviceCaps(picPage.hdc, LOGPIXELSX)
    mRatioX = X1 / X2
    Y1 = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
    Y2 = GetDeviceCaps(picPage.hdc, LOGPIXELSY)
    mRatioY = Y1 / Y2
    Set mcolMergePic = New Collection
    
    intPageFrom = IIf(mlngStartPage > 0, mlngStartPage, 1)
    intPageTo = mlngPageCount
    intCopies = 1
    blnCopyOrder = True
    blnRangePrint = False
    strOldPrinterName = Printer.DeviceName
    ReDim Pages(0 To 0) As Long '�±�Ϊ0��δʹ��
    
    '��ȡ��ӡ������Ϣ
    If strPrintDeviceName = "" Then '�����ӡ������ʾֱ�Ӵ�ӡ
        With frmPrintAsk
            .mstrPageRange = intPageFrom & "-" & intPageTo
            If Me.Visible Then
                .Show vbModal, Me
            Else
                .Show vbModal, mobjParent
            End If
            If .blnOK = False Then Unload frmPrintAsk: Exit Function
            
            If .optPageScope(2).Value = True Then
                'ҳ�뷶Χ
                blnRangePrint = True
                T = Split(.txtPageScope.Tag, ",")
                For i = 0 To UBound(T)
                    aryPage = Split(T(i), "-")
                    If UBound(aryPage) = 0 Then
                        'ֻ��һҳ
                        lngPageCount = UBound(Pages) + 1
                        ReDim Preserve Pages(0 To lngPageCount) As Long
                        Pages(lngPageCount) = Val(T(i))
                    ElseIf UBound(aryPage) = 1 Then
                        l = Val(Split(T(i), "-")(0))
                        m = Val(Split(T(i), "-")(1))
                        For j = l To m Step IIf(m > l, 1, -1)
                            blnHave = False
                            For k = 1 To UBound(Pages)
                                If Pages(k) = j Then blnHave = True
                            Next
                            If blnHave = False Then
                                lngPageCount = UBound(Pages) + 1
                                ReDim Preserve Pages(0 To lngPageCount) As Long
                                Pages(lngPageCount) = j
                            End If
                        Next
                    End If
                Next
            ElseIf .optPageScope(1).Value = True Then
                '��ǰҳ
                intPageFrom = mlngCurPage: intPageTo = mlngCurPage
            Else
                'ȫ����ӡ
                intPageFrom = IIf(mlngStartPage > 0, mlngStartPage, 1): intPageTo = mlngPageCount
            End If
            bytPageOddEven = .cboPageOddEven.ListIndex
            intCopies = Val(.txtCopies.Text)
            blnCopyOrder = IIf(.chkCopyOrder.Value = vbChecked, True, False)
            If Printers(.cboPrinterName.ListIndex).DeviceName <> Printer.DeviceName Then
                Set Printer = Printers(.cboPrinterName.ListIndex)
            End If
            Unload frmPrintAsk
        End With
    Else
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = strPrintDeviceName Then
                Set Printer = Printers(i)
                Exit For
            End If
        Next
    End If
       
    If bytPageOddEven = 1 Then
        '����ҳ
        If intPageFrom Mod 2 = 0 Then intPageFrom = intPageFrom + 1
    ElseIf bytPageOddEven = 2 Then
        'ż��ҳ
        If intPageFrom Mod 2 = 1 Then intPageFrom = intPageFrom + 1
    End If
    If intPageFrom > intPageTo Then Exit Function
    
    
    Err = 0: On Error GoTo PrintErr
    lngOldPaperKind = Printer.PaperSize
    
    '����ֽ�ţ��Զ���ֽ�ŵĵ�������
    If mPaper.PaperType = Split(PageSize42, ",")(9) Then
        Call SetCustomPager(Me.hWnd, mPaper.PaperWidth, mPaper.PaperHeight)
    Else
        Printer.PaperSize = mPaper.PaperType
    End If
    Printer.Orientation = mPaper.Orientation
    '��ֽ��Դ�����ⲿ���������
    
    '��ʼ��ӡ
    Printer.Print Space(1)
    If blnCopyOrder = True Then
        '��ݴ�ӡ
        For lngNumber = 1 To intCopies
            For p = intPageFrom To intPageTo Step IIf(bytPageOddEven = 0, 1, 2)
                If blnRangePrint Then
                    'ҳ�뷶Χ��ӡ
                    For i = 1 To UBound(Pages)
                        If p = Pages(i) Then
                            If lngNumber > 1 Or blnFirstPrinted Then Printer.NewPage
                            DrawPage p, IIf(p = mlngStartPage, mlngBlankHeight, 0), Printer
                            blnFirstPrinted = True
                            Exit For
                        End If
                    Next
                Else
                    If lngNumber > 1 Or blnFirstPrinted = True Then Printer.NewPage
                    DrawPage p, IIf(p = mlngStartPage, mlngBlankHeight, 0), Printer
                    blnFirstPrinted = True
                End If
            Next
        Next
    Else
        For p = intPageFrom To intPageTo Step IIf(bytPageOddEven = 0, 1, 2)
            For lngNumber = 1 To intCopies
                If blnRangePrint Then
                    'ҳ�뷶Χ��ӡ
                    For i = 1 To UBound(Pages)
                        If p = Pages(i) Then
                            If lngNumber > 1 Or blnFirstPrinted = True Then Printer.NewPage
                            DrawPage p, IIf(p = mlngStartPage, mlngBlankHeight, 0), Printer
                            blnFirstPrinted = True
                            Exit For
                        End If
                    Next
                Else
                    If lngNumber > 1 Or blnFirstPrinted = True Then Printer.NewPage
                    DrawPage p, IIf(p = mlngStartPage, mlngBlankHeight, 0), Printer
                    blnFirstPrinted = True
                End If
            Next
        Next
    End If
    
    Printer.EndDoc
    If mPaper.PaperType = Split(PageSize42, ",")(9) Then
        Call SetCustomPager(Me.hWnd, mPaper.PaperWidth, mPaper.PaperHeight)
    Else
        Printer.PaperSize = lngOldPaperKind
    End If
    
    '�ָ�Ĭ�ϴ�ӡ��
    For j = 0 To Printers.Count - 1
        If Printers(j).DeviceName = strOldPrinterName Then
            Set Printer = Printers(j)
            Exit For
        End If
    Next
    
    RaiseEvent PrintEpr
    
    PrintTable = True
    Exit Function
PrintErr:
    MsgBox "��ӡʧ��:" & Err.Description, vbInformation
    
    PrintTable = False
End Function


Private Sub PreviewPage(ByVal lngPage As Long)
'����: Ԥ����lngPageҳ��ҳ��
    picBlank.Visible = (lngPage = mlngStartPage)
    LockWindowUpdate picPage.hWnd
            
    '����ͼƬ
    picPage.Width = mPaper.PaperWidth * mdblZoomFactor
    picPage.Height = mPaper.PaperHeight * mdblZoomFactor
    picPage.Cls
    
    If mdblZoomFactor = 1 Then
        picPage.Picture = mcolPage("K" & lngPage)
    Else
        picBuff.Cls
        picBuff.Width = mPaper.PaperWidth
        picBuff.Height = mPaper.PaperHeight
        picBuff.Picture = mcolPage("K" & lngPage)
        
        '���ð�ɫ������Ч����ã�
        SetStretchBltMode picPage.hdc, HALFTONE
        StretchBlt picPage.hdc, 0, 0, picPage.Width, picPage.Height, picBuff.hdc, 0, 0, picBuff.Width, picBuff.Height, SRCCOPY
    End If
    
    Call Reposition
    If lngPage = mlngStartPage And mlngBlankHeight > 100 Then Call DrawAlphaRect(mlngBlankHeight * mdblZoomFactor)
    LockWindowUpdate 0
    UpdateWindow picPage.hWnd
    stbThis.Panels(2).Text = " �� " & mlngCurPage & " ҳ/ �� " & mlngPageCount & " ҳ"
End Sub

Private Function GetPage(ByVal lngRow As Long, ByVal lngCol As Long) As Long
'���ܣ���ȡָ���������ڵ�ҳ
    Dim i As Long
    
    For i = 0 To UBound(mPages)
        If lngRow >= mPages(i).BRow And lngRow <= mPages(i).Erow Then
            If lngCol >= mPages(i).BCol And lngCol <= mPages(i).ECol Then
                GetPage = i
                Exit Function
            End If
        End If
    Next
End Function

Private Sub SplitPage()
'���ܣ� ��ҳ����,������ҳ�����Լ�ÿҳ����ʼ����
'       �Ȱ��м���,�ٰ��м��㣬���ĳ�и߻��п���һҳ�߻��ֻ��һҳ,�������ֲ���ӡ(��drawcell��ʵ��)
    Dim lngW As Long, lngH As Long, i As Long, j As Long, u As Long
    Dim lngPageCount As Long '���м������ҳ��
    
    '����һҳ
    ReDim mPages(0)
    With mPages(UBound(mPages))
        .BRow = 1
        .Erow = 1
        .BCol = 1
        .ECol = 1
    End With
    
    With mTableThis
        lngW = mPaper.AvailableWidth
        lngH = mPaper.AvailableHeight
                
        '1.�Ȱ��м���
        If .Cells.Rows > 1 Then
            lngH = lngH - .Cells.Cell(1, 1).Height
            For i = 2 To .Cells.Rows
                lngH = lngH - .Cells.Cell(i, 1).Height
                If lngH < 0 Then '��ҳ
                    ReDim Preserve mPages(UBound(mPages) + 1)
                    u = UBound(mPages)
                    mPages(u).BRow = i
                    mPages(u).Erow = i
                    mPages(u).BCol = 1
                    mPages(u).ECol = 1
                    lngH = mPaper.AvailableHeight - .Cells.Cell(i, 1).Height
                Else
                    mPages(UBound(mPages)).Erow = i
                End If
            Next
        End If
        lngPageCount = UBound(mPages) + 1
        
        '2.�ٰ��м���
        If .Cells.Cols > 1 Then
            lngW = lngW - .Cells.Cell(1, 1).Width
            For i = 2 To .Cells.Cols
                lngW = lngW - .Cells.Cell(1, i).Width
                If lngW < 0 Then '��ҳ
                    For j = 0 To lngPageCount - 1
                        ReDim Preserve mPages(UBound(mPages) + 1)
                        u = UBound(mPages)
                        mPages(u).BRow = mPages(j).BRow
                        mPages(u).Erow = mPages(j).Erow
                        mPages(u).BCol = i
                        mPages(u).ECol = i
                    Next
                    lngW = mPaper.AvailableWidth - .Cells.Cell(1, i).Width
                Else
                    For j = UBound(mPages) + 1 - lngPageCount To UBound(mPages)
                        mPages(j).ECol = i
                    Next
                End If
            Next
        End If
    End With
    mlngPageCount = UBound(mPages) + 1
End Sub
