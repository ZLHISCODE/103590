VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPathItemEditOut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ŀ����"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12015
   Icon            =   "frmPathItemEditOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraImportRef 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      TabIndex        =   33
      Top             =   4680
      Width           =   12015
      Begin RichTextLib.RichTextBox rtfImportRef 
         Height          =   1600
         Left            =   120
         TabIndex        =   34
         Top             =   300
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   2831
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmPathItemEditOut.frx":058A
      End
      Begin VB.Label lblImportRef 
         AutoSize        =   -1  'True
         Caption         =   "δ�ɹ������ҽ������"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   60
         Width           =   1800
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12015
      TabIndex        =   5
      Top             =   0
      Width           =   12015
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·����Ŀ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1095
         TabIndex        =   7
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "    ���������ٴ�·������һ���׶��е���Ŀ��Ϣ��������Ӧ��ҽ���������ȣ�������ָ����Ŀ�Ŀ�ѡִ�н����"
         Height          =   180
         Left            =   1095
         TabIndex        =   6
         Top             =   480
         Width           =   9000
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathItemEditOut.frx":0627
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   12840
         Y1              =   825
         Y2              =   825
      End
   End
   Begin VB.Frame fraContent 
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   12255
      Begin VB.Frame fraERPType 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5280
         TabIndex        =   36
         Top             =   383
         Width           =   1575
         Begin VB.OptionButton optEPRType 
            Caption         =   "�°�"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   38
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optEPRType 
            Caption         =   "�ϰ�"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame fraSend 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   750
         Width           =   11775
         Begin VB.OptionButton optExecute 
            Caption         =   "��������(����·�����е�˵�����ֵ�)"
            Height          =   180
            Index           =   0
            Left            =   3600
            TabIndex        =   27
            Top             =   30
            Width           =   3420
         End
         Begin VB.OptionButton optExecute 
            Caption         =   "��������"
            Height          =   180
            Index           =   1
            Left            =   840
            TabIndex        =   26
            Top             =   30
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optExecute 
            Caption         =   "��Ҫʱ����"
            Height          =   180
            Index           =   3
            Left            =   2145
            TabIndex        =   25
            Top             =   30
            Width           =   1200
         End
         Begin VB.Label lblSendKind 
            Caption         =   "���ɷ�ʽ"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   -7
            Width           =   855
         End
      End
      Begin VB.Frame fraKind 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   960
         TabIndex        =   20
         Top             =   413
         Width           =   2895
         Begin VB.OptionButton optType 
            Caption         =   "������"
            Height          =   180
            Index           =   2
            Left            =   1905
            TabIndex        =   23
            Top             =   0
            Width           =   840
         End
         Begin VB.OptionButton optType 
            Caption         =   "������"
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   22
            Top             =   0
            Width           =   840
         End
         Begin VB.OptionButton optType 
            Caption         =   "ҽ����"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   960
         MaxLength       =   1000
         TabIndex        =   16
         Top             =   30
         Width           =   9375
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11400
         MouseIcon       =   "frmPathItemEditOut.frx":2169
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   15
         Width           =   330
      End
      Begin VB.Label lblEPR 
         Caption         =   "�����汾"
         Height          =   180
         Left            =   4440
         TabIndex        =   39
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblIcon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀͼ��"
         Height          =   180
         Left            =   10605
         TabIndex        =   18
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12015
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8310
      Width           =   12015
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   10680
         TabIndex        =   30
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   9480
         TabIndex        =   29
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   0
         X2              =   12720
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   0
         X2              =   12720
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraExecute 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   0
      TabIndex        =   11
      Top             =   6720
      Width           =   12015
      Begin VSFlex8Ctl.VSFlexGrid vsResult 
         Height          =   1395
         Left            =   5205
         TabIndex        =   4
         Top             =   120
         Width           =   6615
         _cx             =   11668
         _cy             =   2461
         Appearance      =   2
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathItemEditOut.frx":22BB
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         OwnerDraw       =   0
         Editable        =   2
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
         AutoSizeMouse   =   -1  'True
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
         Begin MSComctlLib.ImageList imgNature 
            Left            =   540
            Top             =   390
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   6
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":2339
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":28D3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":2E6D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":3407
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":39A1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmPathItemEditOut.frx":3F3B
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   12840
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   12720
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblExePrompt 
         Caption         =   $"frmPathItemEditOut.frx":44D5
         Height          =   1455
         Left            =   960
         TabIndex        =   31
         Top             =   120
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmPathItemEditOut.frx":4577
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblResult 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�н��"
         Height          =   180
         Left            =   4380
         TabIndex        =   12
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   555
      Top             =   435
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraAdvice 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   12015
      Begin zlCISPath.UCAdviceList UcAdvice 
         Height          =   2415
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   4128
      End
      Begin VB.OptionButton optSend 
         Caption         =   "ѡ��ʹ��"
         Height          =   255
         Index           =   1
         Left            =   10695
         TabIndex        =   3
         Top             =   83
         Width           =   1095
      End
      Begin VB.OptionButton optSend 
         Caption         =   "ȫ��ʹ��"
         Height          =   255
         Index           =   0
         Left            =   9510
         TabIndex        =   2
         Top             =   83
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdvice 
         Caption         =   "ҽ���༭(&E)"
         Height          =   350
         Left            =   90
         TabIndex        =   1
         Top             =   60
         Width           =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   4
         X1              =   0
         X2              =   12885
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   5
         X1              =   0
         X2              =   12885
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblSend 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����·����Ŀʱ������ҽ��"
         Height          =   180
         Left            =   7320
         TabIndex        =   14
         Top             =   120
         Width           =   2160
      End
   End
   Begin VB.Frame fraEPR 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   12015
      Begin VSFlex8Ctl.VSFlexGrid vsEPR 
         Height          =   4155
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   11760
         _cx             =   20743
         _cy             =   7329
         Appearance      =   2
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathItemEditOut.frx":4D75
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         OwnerDraw       =   0
         Editable        =   2
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
         AutoSizeMouse   =   -1  'True
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
   End
   Begin XtremeCommandBars.CommandBars cbsIcon 
      Left            =   90
      Top             =   660
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPathItemEditOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event CheckDataValid(PathItem As TYPE_PATH_ITEM, Cancel As Boolean)

Private mrsResult       As ADODB.Recordset          'ִ�н����
Private mrsNature       As ADODB.Recordset          '������ʼ�
Private mrsAdvice       As ADODB.Recordset
Private mvPreItem       As TYPE_PATH_ITEM
Private mvBakItem       As TYPE_PATH_ITEM
Private mvItem          As TYPE_PATH_ITEM
Private mlng·��ID      As Long
Private mlngItemID      As Long
Private mblnAdjust      As Boolean                  '�Ƿ�΢��ģʽ
Private mblnReadOnly    As Boolean                  '�Ƿ�ֻ���鿴ģʽ
Private mblnUseExecute  As Boolean                  '�Ƿ�����ִ�л���
Private mblnReturn      As Boolean
Private mblnChange      As Boolean
Private mblnOK          As Boolean
Private mstrPrivs       As String                   'ģ��Ȩ��

Private Enum CONST_COL_ִ�н��
    colִ��ͼ�� = 0
    colִ�н�� = 1
    col������� = 2
    colȱʡ��� = 3
End Enum

Public Sub ShowView(frmParent As Object, ByVal lngItemID As Long)
'���ܣ��鿴��Ŀ
    mlngItemID = lngItemID
    mblnReadOnly = True

    Me.Show 1, frmParent
End Sub

Public Function ShowEdit(frmParent As Object, rsAdvice As ADODB.Recordset, vItem As TYPE_PATH_ITEM, vPreItem As TYPE_PATH_ITEM, ByVal blnAdjust As Boolean, Optional ByVal lng·��ID As Long, Optional ByVal strPrivs As String) As Boolean
'���ܣ����õ�ǰѡ����Ŀ����ϸ����
'������rsAdvice=(��/��)�Ѿ�����ĵ�ǰ·�����е�ҽ����¼ȫ��
'      vItem=(��/��)��Ҫ���޸�ʱ��ǰ��Ŀ������
'      mvPreItem = (��)ǰһ��ʱ��׶�����ͬ����Ŀ����������ʱ�ο�
'      blnAdjust = �Ƿ����΢��ģʽ
'      lng·��ID = ���·����ʱ�༭��·��ID
    Set mrsAdvice = rsAdvice
    mvItem = vItem
    mvBakItem = vItem
    mvPreItem = vPreItem
    mblnAdjust = blnAdjust
    mlng·��ID = lng·��ID
    mstrPrivs = strPrivs

    Me.Show 1, frmParent

    If mblnOK Then
        vItem = mvItem
    End If
    ShowEdit = mblnOK
End Function

Private Sub cbsIcon_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.ID = -1 Then
        picIcon.Cls
        mvItem.ͼ��ID = 0
    Else
        mvItem.ͼ��ID = Control.ID
        Call DrawPicture(GetPathIcon(mvItem.ͼ��ID))
    End If
    mblnChange = True
End Sub

Private Sub DrawPicture(objPic As StdPicture)
    Dim X As Long, Y As Long, W As Long, H As Long

    W = picIcon.ScaleX(objPic.Width, vbHimetric, vbTwips)
    H = picIcon.ScaleY(objPic.Height, vbHimetric, vbTwips)

    X = (picIcon.ScaleWidth - W) / 2
    Y = (picIcon.ScaleHeight - H) / 2

    picIcon.PaintPicture objPic, X, Y, W, H
End Sub

Private Sub cmdAdvice_Click()
'���ܣ��༭��Ŀ����Ӧ��ҽ��
    Dim rsScheme As ADODB.Recordset
    Dim strFilter As String, lng��� As Long
    Dim colAdviceID As New Collection
    Dim strҽ��IDs As String, lngҽ��ID As Long
    Dim strItem As String, i As Long
    Dim strSql As String, rsTmp As Recordset
    Dim strʹ�ÿ��� As String
    Dim strSelectedID As String
    Dim strSelectedIDAlt As String
    Dim blnUpdate As Boolean

    If mvItem.ҽ��IDs <> "" Then
        If optSend(1).Value Then    '�����Ϲ�ѡ��û�и���"mrsAdvice!�Ƿ�ȱʡ"������ʱ�Ÿ���
            strSelectedID = "," & UCAdvice.GetAdviceIDSelected & ","
            strSelectedIDAlt = "," & UCAdvice.GetAdviceIDSelected(1) & ","
        End If

        Call InitSchemeRecordset(rsScheme)

        strFilter = "": strҽ��IDs = ""
        If mvItem.Edit = 2 And mvItem.�����ҽ��IDs <> "" Then
           '��ǰ�����ҽ����δ����
            strҽ��IDs = mvItem.�����ҽ��IDs
        Else
            strҽ��IDs = mvItem.ҽ��IDs
        End If
        For i = 0 To UBound(Split(strҽ��IDs, ","))
            strFilter = strFilter & " Or ID=" & Split(strҽ��IDs, ",")(i)
        Next
        mrsAdvice.Filter = Mid(strFilter, 5)
        Do While Not mrsAdvice.EOF
            rsScheme.AddNew
            rsScheme!��� = mrsAdvice!ID
            rsScheme!������ = mrsAdvice!���id
            rsScheme!��Ч = mrsAdvice!��Ч
            rsScheme!������ĿID = mrsAdvice!������ĿID
            rsScheme!�շ�ϸĿID = mrsAdvice!�շ�ϸĿID
            rsScheme!ҽ������ = mrsAdvice!ҽ������
            rsScheme!�������� = mrsAdvice!��������
            rsScheme!�ܸ����� = mrsAdvice!�ܸ�����
            rsScheme!ҽ������ = mrsAdvice!ҽ������
            rsScheme!ִ��Ƶ�� = mrsAdvice!ִ��Ƶ��
            rsScheme!Ƶ�ʴ��� = mrsAdvice!Ƶ�ʴ���
            rsScheme!Ƶ�ʼ�� = mrsAdvice!Ƶ�ʼ��
            rsScheme!�����λ = mrsAdvice!�����λ
            rsScheme!ʱ�䷽�� = mrsAdvice!ʱ�䷽��
            rsScheme!ִ�п���ID = mrsAdvice!ִ�п���ID
            rsScheme!ִ������ = mrsAdvice!ִ������
            rsScheme!�걾��λ = mrsAdvice!�걾��λ
            rsScheme!��鷽�� = mrsAdvice!��鷽��
            rsScheme!�Ƿ�ȱʡ = IIf(InStr(strSelectedID, "," & mrsAdvice!ID & ",") > 0, 1, 0)
            rsScheme!�Ƿ�ѡ = IIf(InStr(strSelectedIDAlt, "," & mrsAdvice!ID & ",") > 0, 1, 0)
            rsScheme!�䷽ID = mrsAdvice!�䷽ID
            rsScheme!�����ĿID = mrsAdvice!�����ĿID
            rsScheme!ִ�б�� = mrsAdvice!ִ�б��
            rsScheme.Update
            mrsAdvice.MoveNext
        Loop
        mrsAdvice.Filter = ""
    End If

    On Error GoTo errH
    strSql = "Select ����ID From ����·������ Where ·��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)
    Do Until rsTmp.EOF
        strʹ�ÿ��� = strʹ�ÿ��� & "," & rsTmp!����ID
        rsTmp.MoveNext
    Loop
    On Error GoTo 0

    Set rsScheme = gobjKernel.ShowSchemeEdit(Me, 1, rsScheme, False, optSend(1).Value, Mid(strʹ�ÿ���, 2), 1)
    '·��ҽ������
    If Not rsScheme Is Nothing Then
        'blnUpdate=True �������δͣ�õ�·��ҽ���䶯��¼���浽������·��ҽ���䶯���У�����˺��ٸ��±�·��ҽ����
        blnUpdate = InStr(mstrPrivs, "����·��ҽ������") = 0 And mvItem.ԭҽ��IDs <> ""
        '�Ȳ����µ�ҽ��ID
        strҽ��IDs = ""
        Do While Not rsScheme.EOF
            lngҽ��ID = zlDatabase.GetNextId("����·��ҽ������")
            colAdviceID.Add lngҽ��ID, "_" & rsScheme!���
            strҽ��IDs = strҽ��IDs & "," & lngҽ��ID
            rsScheme.MoveNext
        Loop

        If Not blnUpdate Then
            strҽ��IDs = Mid(strҽ��IDs, 2)
            mvItem.ҽ��IDs = strҽ��IDs
        Else
            strҽ��IDs = Mid(strҽ��IDs, 2)
            mvItem.�����ҽ��IDs = strҽ��IDs
        End If
        '�����µ�ҽ��
        rsScheme.MoveFirst: lng��� = 1
        Do While Not rsScheme.EOF
            lngҽ��ID = colAdviceID("_" & rsScheme!���)
            mrsAdvice.AddNew

            mrsAdvice!ID = lngҽ��ID
            If Not IsNull(rsScheme!������) Then
                mrsAdvice!���id = colAdviceID("_" & rsScheme!������)
            End If
            mrsAdvice!��� = lng���
            mrsAdvice!��Ч = rsScheme!��Ч
            mrsAdvice!������ĿID = rsScheme!������ĿID
            mrsAdvice!�շ�ϸĿID = rsScheme!�շ�ϸĿID
            If IsNull(rsScheme!������ĿID) Then
                mrsAdvice!ҽ������ = rsScheme!ҽ������ '����¼��ҽ���ű���
            End If
            mrsAdvice!�������� = rsScheme!��������
            mrsAdvice!�ܸ����� = rsScheme!�ܸ�����
            mrsAdvice!ҽ������ = rsScheme!ҽ������
            mrsAdvice!ִ��Ƶ�� = rsScheme!ִ��Ƶ��
            mrsAdvice!Ƶ�ʴ��� = rsScheme!Ƶ�ʴ���
            mrsAdvice!Ƶ�ʼ�� = rsScheme!Ƶ�ʼ��
            mrsAdvice!�����λ = rsScheme!�����λ
            mrsAdvice!ʱ�䷽�� = rsScheme!ʱ�䷽��
            mrsAdvice!ִ�п���ID = rsScheme!ִ�п���ID
            mrsAdvice!ִ������ = rsScheme!ִ������
            mrsAdvice!�걾��λ = rsScheme!�걾��λ
            mrsAdvice!��鷽�� = rsScheme!��鷽��
            mrsAdvice!�Ƿ�ȱʡ = rsScheme!�Ƿ�ȱʡ
            mrsAdvice!�Ƿ�ѡ = rsScheme!�Ƿ�ѡ
            mrsAdvice!�䷽ID = rsScheme!�䷽ID
            mrsAdvice!�����ĿID = rsScheme!�����ĿID
            mrsAdvice!ִ�б�� = rsScheme!ִ�б��
            mrsAdvice!����� = IIf(blnUpdate, 1, 0)
            mrsAdvice!��ĿID = IIf(blnUpdate, mvItem.ID, 0)
            mrsAdvice.Update

            lng��� = lng��� + 1
            rsScheme.MoveNext
        Loop
        'ˢ����ʾ
        Call ShowAdvice(strҽ��IDs)

        'ȱʡ��Ŀ����
        If txtItem.Text = "" Then
            txtItem.Text = UCAdvice.GetAdviceTitle
        End If
        mblnChange = True
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnCancel As Boolean
    Dim strFilter As String, strSelectedID As String
    Dim i As Long
    Dim strSelectedAltID As String
    Dim blnIsAllSelect As Boolean
    Dim strTmp As String

    If mblnReadOnly Then
        mblnOK = True: Unload Me: Exit Sub
    End If

    '���ݼ��
    If Trim(txtItem.Text) = "" Then
        MsgBox "������·����Ŀ�����ݡ�", vbInformation, gstrSysName
        txtItem.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txtItem.Text) > txtItem.MaxLength Then
        MsgBox "��Ŀ������������� " & txtItem.MaxLength \ 2 & " �����ֻ��� " & txtItem.MaxLength & "��", vbInformation, gstrSysName
        txtItem.SetFocus: Exit Sub
    End If

    '���ҽ��
    If optType(0).Value Then
        If mvItem.ҽ��IDs = "" Then
            MsgBox "û�ж��嵱ǰ��Ŀ����Ӧ��ҽ�����ݡ�", vbInformation, gstrSysName
            If cmdAdvice.Enabled Then cmdAdvice.SetFocus
            Exit Sub
        End If
        strFilter = ""
        For i = 0 To UBound(Split(mvItem.ҽ��IDs, ","))
            strFilter = strFilter & " Or ID=" & Split(mvItem.ҽ��IDs, ",")(i)
        Next
        strSelectedID = "," & UCAdvice.GetAdviceIDSelected & ","
        strSelectedAltID = "," & UCAdvice.GetAdviceIDSelected(1, blnIsAllSelect) & ","

        mrsAdvice.Filter = Mid(strFilter, 5)
        strFilter = ""
        If blnIsAllSelect Then
            '����Ҫ��һ�����Ǳ�ѡ
            MsgBox "һ��·����Ŀ������һ�����Ǳ�ѡҽ����", vbInformation, Me.Caption
            Exit Sub
        End If
        Do While Not mrsAdvice.EOF
            If InStr(strSelectedID, "," & mrsAdvice!ID & ",") > 0 Then
                mrsAdvice!�Ƿ�ȱʡ = 1
            Else
                mrsAdvice!�Ƿ�ȱʡ = 0
            End If
            If InStr(strSelectedAltID, "," & mrsAdvice!ID & ",") > 0 Then
                mrsAdvice!�Ƿ�ѡ = 1
            Else
                mrsAdvice!�Ƿ�ѡ = 0
            End If
            mrsAdvice.Update
            mrsAdvice.MoveNext
        Loop
        mrsAdvice.Filter = ""
    Else
        mvItem.ҽ��IDs = ""
    End If

    '��鲡��
    If optType(1).Value Then
        With vsEPR
            strFilter = ""
            mvItem.����IDs = "": mvItem.�°没��IDs = "": mvItem.�������� = ""
            For i = .FixedRows To .Rows - 1
                If .RowData(i) <> "" Then
                    strTmp = .RowData(i) '��ʽ��(NEW/OLD)|ID
                    If Split(strTmp, "|")(0) = "OLD" Then
                        mvItem.����IDs = mvItem.����IDs & "," & Split(strTmp, "|")(1)
                    Else
                        mvItem.�°没��IDs = mvItem.�°没��IDs & "," & Split(strTmp, "|")(1)
                    End If
                    '��������:�ļ�ID,ԭ��ID,����,���
                    mvItem.�������� = mvItem.�������� & ";" & IIf(Split(strTmp, "|")(0) = "OLD", Split(strTmp, "|")(1) & ",", "," & Split(strTmp, "|")(1)) & "," & Trim(.TextMatrix(i, 0)) & "," & i
                    If InStr(strFilter & ",", "," & .TextMatrix(i, 0) & ",") = 0 Then
                        strFilter = strFilter & "," & .TextMatrix(i, 0)
                    Else
                        MsgBox "ָ�����ظ��Ĳ����ļ�""" & .TextMatrix(i, 0) & """��", vbInformation, gstrSysName
                        .Row = i: Call .ShowCell(.Row, .Col)
                        .SetFocus: Exit Sub
                    End If
                End If
            Next
            mvItem.����IDs = Mid(mvItem.����IDs, 2)
            mvItem.�������� = Mid(mvItem.��������, 2)
            mvItem.�°没��IDs = Mid(mvItem.�°没��IDs, 2)
            If mvItem.����IDs = "" And mvItem.�°没��IDs = "" Then
                MsgBox "��ָ����Ŀ����Ӧ�Ĳ����ļ���", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
        End With
    Else
        mvItem.����IDs = "": mvItem.�°没��IDs = ""
    End If

    '�����
    If Not optExecute(0).Value And fraExecute.Visible Then
        With vsResult
            strFilter = ""
            mvItem.��Ŀ��� = ""
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, colִ�н��)) <> "" Then
                    If InStr(strFilter & ",", "," & .TextMatrix(i, colִ�н��) & ",") = 0 _
                        And InStr(strFilter, "," & .TextMatrix(i, colִ�н��) & "|") = 0 Then
                        strFilter = strFilter & "," & .TextMatrix(i, colִ�н��)
                        If .TextMatrix(i, col�������) <> "" Then
                            mrsNature.Filter = "����='" & .TextMatrix(i, col�������) & "'"
                            strFilter = strFilter & "|" & mrsNature!����
                        End If
                    Else
                        MsgBox "ָ�����ظ���ִ�н��""" & .TextMatrix(i, colִ�н��) & """��", vbInformation, gstrSysName
                        .Row = i: Call .ShowCell(.Row, .Col)
                        .SetFocus: Exit Sub
                    End If

                    'ȱʡ���
                    If Val(.TextMatrix(i, colȱʡ���)) <> 0 Then
                        mvItem.��Ŀ��� = .TextMatrix(i, colִ�н��)
                    End If
                End If
            Next
            strFilter = Mid(strFilter, 2)
            If strFilter = "" Then
                MsgBox "��ָ����Ŀ����Ӧ��ִ�н����", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
            If mvItem.��Ŀ��� = "" Then
                MsgBox "��ָ����Ŀ����Ӧ��ȱʡ�����", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
            mvItem.��Ŀ��� = strFilter & vbTab & mvItem.��Ŀ���
        End With
    Else
        mvItem.��Ŀ��� = ""
    End If

    '���������ռ�
    If mvItem.ҽ��IDs <> "" Then
        mvItem.����Ҫ�� = IIf(optSend(0).Value, 0, 1) '0-ȫ�����ɣ�1-ѡ������
    Else
        mvItem.����Ҫ�� = 0
    End If
    mvItem.��Ŀ���� = txtItem.Text
    For i = 0 To optExecute.UBound
        If i <> 2 Then
            If optExecute(i).Value Then mvItem.ִ�з�ʽ = i: Exit For
        End If
    Next
    RaiseEvent CheckDataValid(mvItem, blnCancel)
    If blnCancel Then Exit Sub

    If mblnChange Then
        mvItem.������ = 1
    End If

    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If TypeName(ActiveControl) <> "VSFlexGrid" Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strDefault As String, strSql As String
    Dim arrResult As Variant, i As Long
    Dim objControl As Object
    Dim strTmp As String

    On Error GoTo errH

    mblnOK = False
    fraEPR.BackColor = Me.BackColor
    fraAdvice.BackColor = Me.BackColor
    fraExecute.BackColor = Me.BackColor

    mblnUseExecute = Val(zlDatabase.GetPara("�Ƿ�����·��ִ�л���", glngSys, P����·��Ӧ��, 1))
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsIcon.VisualTheme = xtpThemeOffice2003
    cbsIcon.ActiveMenuBar.Visible = False
    With Me.cbsIcon.Options
        .ToolBarAccelTips = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With

    'ֻ���鿴ģʽʱ����ȡһЩ����
    If mblnReadOnly And mlngItemID <> 0 Then
        strSql = "Select ��Ŀ����,ִ�з�ʽ,��Ŀ���,ͼ��ID,����Ҫ�� From ����·����Ŀ Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngItemID)

        '��Ŀ������Ϣ
        mvItem.ID = mlngItemID
        mvItem.��Ŀ���� = rsTmp!��Ŀ����
        mvItem.��Ŀ��� = NVL(rsTmp!��Ŀ���)
        mvItem.ִ�з�ʽ = NVL(rsTmp!ִ�з�ʽ, 0)
        mvItem.ͼ��ID = NVL(rsTmp!ͼ��ID, 0)
        mvItem.����Ҫ�� = Val("" & rsTmp!����Ҫ��)

        '����ҽ����Ϣ
        strSql = "Select ҽ������ID From ����·��ҽ�� Where ·����ĿID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngItemID)
        strSql = ""
        Do While Not rsTmp.EOF
            strSql = strSql & "," & rsTmp!ҽ������ID
            rsTmp.MoveNext
        Loop
        mvItem.ҽ��IDs = Mid(strSql, 2)

        '����������Ϣ
        strSql = "Select �ļ�ID,ԭ��ID From ����·������ Where ��ĿID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngItemID)
        strSql = ""
        Do While Not rsTmp.EOF
            If rsTmp!�ļ�ID & "" <> "" Then
                strSql = strSql & "," & rsTmp!�ļ�ID
            Else
                strTmp = strTmp & "," & rsTmp!ԭ��ID
            End If
            rsTmp.MoveNext
        Loop
        mvItem.����IDs = Mid(strSql, 2)
        mvItem.�°没��IDs = Mid(strTmp, 2)
        'ҽ����¼��
        If mvItem.ҽ��IDs <> "" Then
            strSql = " Select Distinct A.ID,A.���ID,A.���,A.��Ч,A.������ĿID,A.�շ�ϸĿID," & _
                     " A.ҽ������,A.��������,A.�ܸ�����,A.�걾��λ,A.��鷽��,A.ҽ������," & _
                     " A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ִ������,A.ִ�б��,A.ִ�п���ID,A.ʱ�䷽��,A.�Ƿ�ȱʡ,A.�Ƿ�ѡ,A.�䷽ID,A.�����ĿID" & _
                     " From ����·��ҽ������ A,����·��ҽ�� B" & _
                     " Where A.ID=B.ҽ������ID And B.·����ĿID=[1]" & _
                     " Order by A.���,A.ID"
            Set mrsAdvice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngItemID)
        End If
    End If

    '��ȡ������ʼ�
    strSql = "Select ����,���� From ·��������� Order by ����"
    Set mrsNature = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsNature, strSql, Me.Caption)
    strSql = ""
    Do While Not mrsNature.EOF
        strSql = strSql & "|" & mrsNature!���� & "-" & mrsNature!����
        mrsNature.MoveNext
    Loop
    vsResult.ColData(col�������) = Mid(strSql, 2)

    '��ȡ���ý����
    strSql = "Select A.����,A.����,Nvl(����,0) as ����,B.���� as ����" & _
        " From ·��������� A,·��������� B" & _
        " Where A.ĩ��=1 And Nvl(A.����,0)=B.����(+)" & _
        " Order by A.����"
    Set mrsResult = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsResult, strSql, Me.Caption)

    '�༭����ʱ��һЩ����
    If mvItem.ID <> 0 Then
        txtItem.Text = mvItem.��Ŀ����
        If mvItem.ͼ��ID <> 0 Then
            Call DrawPicture(GetPathIcon(mvItem.ͼ��ID))
        End If
        '----
        If mvItem.ҽ��IDs <> "" Then
            '��ʾҽ��
            optType(0).Value = True
            Call ShowAdvice(mvItem.ҽ��IDs)
            
            optSend(0).Value = (mvItem.����Ҫ�� = 0)
            optSend(1).Value = Not optSend(0).Value

            Call UCAdvice.Setѡ���еĿɼ���(optSend(0).Value)
        ElseIf mvItem.����IDs <> "" Or mvItem.�°没��IDs <> "" Then
            '��ʾ����
            optType(1).Value = True
            If mvItem.Edit = 0 Then
                If mvItem.�°没��IDs = "" Then '�ϰ�
                    strSql = "Select /*+ Rule*/ ID as �ļ�ID,����,1 as �汾 From �����ļ��б�" & _
                        " Where ID IN(Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
                        " Order by ���"
                ElseIf mvItem.����IDs <> "" Then '�°�+�ϰ�
                    strSql = "Select A.�ļ�ID,A.ԭ��ID,Nvl(a.����, b.����) as ����,decode(�ļ�ID,NULL,2,1) as �汾 From �ٴ�·������ A, �����ļ��б� B Where a.��Ŀid = [3] And a.�ļ�id = b.Id(+)" & vbNewLine & _
                        "order by a.���"
                Else '�°�
                    strSql = "Select T.ԭ��ID,T.����,2 as �汾 From ����·������ T Where t.��Ŀid = [3] And t.�ļ�id Is Null And t.ԭ��ID IN(Select Column_Value From Table(Cast(f_STR2list([2]) As zlTools.t_Strlist)))" & _
                        " Order by ���"
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mvItem.����IDs, mvItem.�°没��IDs, mvItem.ID)
            Else
                Set rsTmp = FuncGetEMRInfo(mvItem.��������)
            End If
            With vsEPR
                .Rows = .FixedRows
                Do While Not rsTmp.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = rsTmp!����
                    .Cell(flexcpData, .Rows - 1, 0) = .TextMatrix(.Rows - 1, 0)
                    If rsTmp!�汾 & "" = "1" Then
                        .RowData(.Rows - 1) = "OLD|" & rsTmp!�ļ�ID
                    Else
                        .RowData(.Rows - 1) = "NEW|" & rsTmp!ԭ��ID
                    End If
                    rsTmp.MoveNext
                Loop
                If Not mblnReadOnly Then .Rows = .Rows + 1 '����һ������������
            End With
        ElseIf mvItem.������ <> 1 Then
            optType(0).Value = True
        Else
            optType(2).Value = True
        End If
        optType(0).Enabled = True
        optType(1).Enabled = True
        '----
        If mvItem.ִ�з�ʽ = 2 Then mvItem.ִ�з�ʽ = 3
        optExecute(mvItem.ִ�з�ʽ).Value = True
        '----
        With vsResult
            .Rows = .FixedRows
            If mvItem.��Ŀ��� <> "" Then
                If UBound(Split(mvItem.��Ŀ���, vbTab)) >= 1 Then
                    strDefault = Split(mvItem.��Ŀ���, vbTab)(1)
                End If
                arrResult = Split(Split(mvItem.��Ŀ���, vbTab)(0), ",")
                For i = 0 To UBound(arrResult)
                    .Rows = .Rows + 1

                    .TextMatrix(.Rows - 1, colִ�н��) = Split(arrResult(i), "|")(0)
                    .Cell(flexcpData, .Rows - 1, colִ�н��) = .TextMatrix(.Rows - 1, colִ�н��)

                    '����������
                    If UBound(Split(arrResult(i), "|")) > 0 Then
                        Set .Cell(flexcpPicture, .Rows - 1, colִ��ͼ��) = imgNature.ListImages(Val(Split(arrResult(i), "|")(1))).Picture
                        mrsNature.Filter = "����=" & Val(Split(arrResult(i), "|")(1))
                        .TextMatrix(.Rows - 1, col�������) = mrsNature!����
                    End If

                    If Split(arrResult(i), "|")(0) = strDefault Then
                        .TextMatrix(.Rows - 1, colȱʡ���) = 1
                    End If
                Next
            End If
            If Not mblnReadOnly And Not mblnAdjust Then .Rows = .Rows + 1 '����һ������������
        End With
    Else
        '����ʱ��ȡ������ִ�н��
        mvItem.������ = 1
        mrsResult.Filter = "����=1"
        If Not mrsResult.EOF Then
            vsResult.Rows = vsResult.FixedRows + 1
            Call SetResultInput(vsResult.FixedRows, mrsResult)
        End If
        If optType(0).Value = True Then Call UCAdvice.Setѡ���еĿɼ���(optSend(0).Value)
    End If

    If Not mblnReadOnly Then
        vsEPR.Row = 0: vsEPR.Row = 1: vsEPR.Col = 0
        If Not mblnAdjust Then
            vsResult.Row = 0: vsResult.Row = 1: vsResult.Col = colִ�н��
        End If
    End If

    'ֻ���鿴ʱ��һЩ���洦��
    If mblnReadOnly Then
        cmdCancel.Visible = False
        cmdOK.Left = cmdCancel.Left

        cmdAdvice.Visible = False

        vsEPR.Editable = flexEDNone
        vsResult.Editable = flexEDNone

        For Each objControl In Me.Controls
            If TypeName(objControl) = "TextBox" Then
                objControl.Locked = True
            ElseIf TypeName(objControl) = "OptionButton" Then
                objControl.Enabled = False
            End If
        Next
    ElseIf mblnAdjust Then
        txtItem.BackColor = Me.BackColor
        txtItem.TabStop = False

        vsResult.Editable = flexEDNone
        vsResult.BackColor = Me.BackColor
        vsResult.BackColorBkg = Me.BackColor
        vsResult.TabStop = False

        For Each objControl In Me.Controls
            If TypeName(objControl) = "TextBox" Then
                objControl.Locked = True
            ElseIf TypeName(objControl) = "OptionButton" Then
                objControl.Enabled = False
            End If
        Next
    End If
    '����ο�
    rtfImportRef.Text = mvItem.����ο�

    Call SetFormFace

    mblnChange = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOK And mvItem.ID <> 0 And mblnChange Then
        If MsgBox("��·����Ŀ����Ϣ�ѱ����ģ�ȷʵҪ���������˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If

    If Not mrsResult Is Nothing Then
        If mrsResult.State = 1 Then mrsResult.Close
        Set mrsResult = Nothing
    End If
    If Not mrsNature Is Nothing Then
        If mrsNature.State = 1 Then mrsNature.Close
        Set mrsNature = Nothing
    End If

    mlngItemID = 0
    mblnAdjust = False
    mblnReadOnly = False
End Sub

Private Sub optSend_Click(Index As Integer)
    Call UCAdvice.Setѡ���еĿɼ���(optSend(0).Value)

    If Visible Then mblnChange = True
End Sub

Private Sub optType_Click(Index As Integer)
    Call SetFormFace

    If Visible Then
        If Index = 0 Then
            If cmdAdvice.Enabled Then cmdAdvice.SetFocus
            Call optSend_Click(0)
        ElseIf Index = 1 Then
            vsEPR.SetFocus
        End If
    End If
End Sub

Private Sub optExecute_Click(Index As Integer)
    Call SetFormFace

    If Visible Then mblnChange = True
End Sub

Private Sub picBottom_GotFocus()
    If cmdOK.Enabled And cmdOK.Visible Then
        cmdOK.SetFocus
    ElseIf cmdCancel.Enabled And cmdCancel.Visible Then
        cmdCancel.SetFocus
    End If
End Sub

Private Sub SetFormFace()
'���ܣ����������������ý���Ŀɼ����ݺͳߴ�
    Dim lngTop As Long, lngHeight As Long

    On Error Resume Next
    If WindowState = 1 Then Exit Sub
    fraAdvice.Enabled = optType(0).Value: fraAdvice.Visible = fraAdvice.Enabled
    fraEPR.Enabled = optType(1).Value: fraEPR.Visible = fraEPR.Enabled
    fraExecute.Enabled = Not optExecute(0).Value And mblnUseExecute: fraExecute.Visible = fraExecute.Enabled
    If optType(1).Value Then
        If gobjEmr Is Nothing Then
            fraERPType.Visible = False: lblEPR.Visible = False
            optEPRType(0).Value = True
        Else
            fraERPType.Visible = True: lblEPR.Visible = True
            optEPRType(1).Value = True
        End If
    Else
        fraERPType.Visible = False: lblEPR.Visible = False
    End If

    fraImportRef.Enabled = mvItem.������ <> 1
    fraImportRef.Visible = fraImportRef.Enabled And fraAdvice.Enabled
    '����Load�¼��е��øù���ʱ������fraImportRef.Visible=True���������Ч����ֵʼ�ձ���False
    If fraImportRef.Enabled And fraAdvice.Enabled Then
        fraImportRef.BackColor = fraAdvice.BackColor
        fraImportRef.Top = IIf(fraExecute.Enabled, fraExecute.Top, picBottom.Top) - 2000
        fraImportRef.Height = 2000
        rtfImportRef.Top = lblImportRef.Top + lblImportRef.Height + 30
        rtfImportRef.Height = fraImportRef.Height - rtfImportRef.Top
        fraAdvice.Height = fraImportRef.Top - fraAdvice.Top
    Else
        fraAdvice.Height = Me.Height - fraAdvice.Top - IIf(fraExecute.Enabled, fraExecute.Height, 0) - picBottom.Height - 450
        fraEPR.Height = fraAdvice.Height
    End If

    UCAdvice.Height = fraAdvice.Height - cmdAdvice.Height - 60
    vsEPR.Height = fraEPR.Height - 60
End Sub

Private Sub picIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long

    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim vPoint As POINTAPI

    On Error GoTo errH

    If mblnReadOnly Or mblnAdjust Then Exit Sub

    If img16.ListImages.count = 0 Then
        strSql = "Select ID,Nvl(����,0) as ���� From �ٴ�·��ͼ�� Order by ���� Desc,ID"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
        Do While Not rsTmp.EOF
            img16.ListImages.Add , "_" & IIf(rsTmp!���� = 1, 1, -1) * rsTmp!ID, GetPathIcon(rsTmp!ID)
            img16.ListImages(img16.ListImages.count).Tag = CStr(rsTmp!ID) 'ҪCStr
            rsTmp.MoveNext
        Loop
        cbsIcon.AddImageList img16
    End If

    Set objPopup = cbsIcon.Add("Popup", xtpBarPopup)
    objPopup.SetPopupToolBar True
    objPopup.Width = 260
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, -1, "�����Ŀͼ��")
        objControl.Flags = xtpFlagControlStretched
        For i = 1 To img16.ListImages.count
            Set objControl = .Add(xtpControlButton, img16.ListImages(i).Tag, "")
            If i = 1 Then
                objControl.BeginGroup = True
            ElseIf Val(Mid(img16.ListImages(i).Key, 2)) < 0 Then
                If Val(Mid(img16.ListImages(i - 1).Key, 2)) > 0 Then
                    objControl.BeginGroup = True
                End If
            End If
        Next
    End With

    vPoint.X = (fraContent.Left + lblIcon.Left - lblIcon.Width - 120) / Screen.TwipsPerPixelX
    vPoint.Y = (fraContent.Top + picIcon.Height) / Screen.TwipsPerPixelY
    ClientToScreen Me.Hwnd, vPoint
    objPopup.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtItem_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txtItem_GotFocus()
    Call zlControl.TxtSelAll(txtItem)
End Sub

Private Sub InitSchemeRecordset(rsScheme As ADODB.Recordset)
    Set rsScheme = New ADODB.Recordset
    rsScheme.Fields.Append "�Ƿ�ѡ", adSmallInt
    rsScheme.Fields.Append "�Ƿ�ȱʡ", adSmallInt
    rsScheme.Fields.Append "���", adBigInt
    rsScheme.Fields.Append "������", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "��Ч", adSmallInt
    rsScheme.Fields.Append "������ĿID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "ҽ������", adVarChar, 1000, adFldIsNullable
    rsScheme.Fields.Append "����", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "��������", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "�ܸ�����", adSingle, , adFldIsNullable
    rsScheme.Fields.Append "ҽ������", adVarChar, 1000, adFldIsNullable
    rsScheme.Fields.Append "ִ��Ƶ��", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "Ƶ�ʴ���", adSmallInt, , adFldIsNullable
    rsScheme.Fields.Append "Ƶ�ʼ��", adSmallInt, , adFldIsNullable
    rsScheme.Fields.Append "�����λ", adVarChar, 10, adFldIsNullable
    rsScheme.Fields.Append "ʱ�䷽��", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "ִ������", adSmallInt
    rsScheme.Fields.Append "�걾��λ", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "��鷽��", adVarChar, 100, adFldIsNullable
    rsScheme.Fields.Append "�䷽ID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "�����ĿID", adBigInt, , adFldIsNullable
    rsScheme.Fields.Append "ִ�б��", adSingle, , adFldIsNullable

    rsScheme.CursorLocation = adUseClient
    rsScheme.LockType = adLockOptimistic
    rsScheme.CursorType = adOpenStatic
    rsScheme.Open
End Sub

Private Function ShowAdvice(ByVal strҽ��IDs As String) As Boolean
'���ܣ���ʾ·����Ŀ��Ӧ��ҽ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim str��ҩ As String, str�巨 As String
    Dim str���� As String, str�걾 As String
    Dim strFilter As String
    Dim i As Long, j As Long

    If strҽ��IDs = "" Then
        Call UCAdvice.ShowAdvice(0, "", 0, 0, mblnReadOnly)
        ShowAdvice = True: Exit Function
    End If

    '���ɶ�̬SQL
    For i = 0 To UBound(Split(strҽ��IDs, ","))
        strFilter = strFilter & " Or ID=" & Split(strҽ��IDs, ",")(i)
    Next
    With mrsAdvice
        strSql = ""
        .Filter = Mid(strFilter, 5)
        Do While Not .EOF
            strSql = strSql & " Union ALL Select "
            For i = 0 To .Fields.count - 1
                If Not IsNull(.Fields(i).Value) Then
                    If Rec.IsType(.Fields(i).Type, adVarChar) Then
                        strSql = strSql & "'" & Replace(Replace(.Fields(i).Value, "[", "("), "]", ")") & "'"
                    Else
                        strSql = strSql & .Fields(i).Value 'û��������
                    End If
                Else
                    If Rec.IsType(.Fields(i).Type, adBigInt) Or Rec.IsType(.Fields(i).Type, adSmallInt) Or Rec.IsType(.Fields(i).Type, adSingle) Then
                        strSql = strSql & "-Null"
                    Else
                        strSql = strSql & "Null"
                    End If
                End If
                strSql = strSql & " As " & .Fields(i).Name & ","
            Next
            strSql = Left(strSql, Len(strSql) - 1) & " From Dual"
            .MoveNext
        Loop
        .Filter = ""
        strSql = Mid(strSql, 12)
    End With
    If strSql = "" Then
        Call UCAdvice.ShowAdvice(0, "", 0, 0, mblnReadOnly)
    Else
        Call UCAdvice.ShowAdvice(0, strSql, 0, 0, mblnReadOnly)
    End If
    ShowAdvice = True
End Function

Private Sub UcAdvice_DataChange()
    If Visible Then mblnChange = True
End Sub

Private Sub vsEPR_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsEPR_AfterRowColChange(-1, -1, Row, Col)
End Sub

Private Sub vsEPR_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsEPR
        If NewCol = 0 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsEPR_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI

    With vsEPR
        If optEPRType(0).Value Then  '��ʾ�°���Ӳ���δ��װ��װ�����������ϰ����̣�
            strSql = "Select A.ID,Decode(A.����,1,'���ﲡ��',5,'����֤������',6,'֪���ļ�') as ����," & _
                " A.���,A.����,A.˵�� From �����ļ��б� A" & _
                " Where A.���� IN(1,5,6) And Nvl(A.����,0) IN(0,1,2) And A.ͨ�� IN(1,2)" & _
                " Order by A.����,A.���"
        Else
            '�°�����
            If gobjEmr Is Nothing Then Exit Sub

            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
                Set gobjEmr = Nothing
                MsgBox "���Ӳ��������������߻򵼺�̨��¼ʱδ�ܳɹ����ӵ��Ӳ���������!", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            Else
                'gobjEmr.GetAntetypeList(byref strParameter as string) as Adodb.RecordSet
                '��¼�������ֶΣ������ţ��������ƣ��������ƣ�ID����ţ����ƣ�˵��
                On Error Resume Next
                Set rsTmp = gobjEmr.GetAntetypeList("")
                Err.Clear: On Error GoTo 0
                If rsTmp Is Nothing Then Exit Sub
                strSql = Rec.ToSQL(rsTmp)
            End If
        End If
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "�����ļ�", False, "", "", False, False, True, _
            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�в����ļ����ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call SetEPRInput(Row, rsTmp)
            Call EPREnterNextCell(True)
        End If
    End With
End Sub

Private Sub vsEPR_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long

    If mblnReadOnly Then Exit Sub

    With vsEPR
        If KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 0) <> "" Then
                If MsgBox("ȷʵҪ������в����ļ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsEPR_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsEPR_KeyPress(KeyAscii As Integer)
    With vsEPR
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EPREnterNextCell
        ElseIf .Col = 0 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsEPR_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsEPR_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsEPR_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsEPR.EditSelStart = 0
    vsEPR.EditSelLength = zlCommFun.ActualLen(vsEPR.EditText)
End Sub

Private Sub vsEPR_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim i As Long
    Dim strFilter As String, strTag As String

    With vsEPR
        If Col = 0 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call EPREnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call EPREnterNextCell
            Else
                strInput = UCase(.EditText)
                If optEPRType(0).Value Then
                    strSql = "Select A.ID,Decode(A.����,1,'���ﲡ��',4,'������',5,'����֤������',6,'֪���ļ�') as ����," & _
                             " A.���,A.����,A.˵�� From �����ļ��б� A" & _
                             " Where A.���� IN(1,4,5,6) And Nvl(A.����,0) IN(0,1,2)" & _
                             " And A.ͨ�� IN(1,2) And (A.��� Like [1] Or A.���� Like [2] Or zlSpellCode(A.����) Like [2])" & _
                             " Order by A.����,A.���"
                Else
                    '�°�����
                    If gobjEmr Is Nothing Then Exit Sub
                    If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
                        Set gobjEmr = Nothing
                        MsgBox "���Ӳ��������������߻򵼺�̨��¼ʱδ�ܳɹ����ӵ��Ӳ���������!", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    Else
                        'gobjEmr.GetAntetypeList(byref strParameter as string) as Adodb.RecordSet
                        '��¼�������ֶΣ������ţ��������ƣ��������ƣ�ID����ţ����ƣ�˵��
                        On Error Resume Next
                        Set rsTmp = gobjEmr.GetAntetypeList("")
                        Err.Clear: On Error GoTo 0
                        If rsTmp Is Nothing Then Exit Sub
                        strSql = Rec.ToSQL(rsTmp)
                        strSql = "select A.* from (" & strSql & ") A where A.��� Like [1] Or A.���� Like [2] Or zlSpellCode(A.����) Like [2]  order by ������,���"
                    End If

                End If
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "�����ļ�", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "û���ҵ�ƥ��Ĳ����ļ���", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call SetEPRInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call EPREnterNextCell(True)
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Sub SetEPRInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ��������ļ�������
    Dim strItem As String, i As Long
    Dim strTmp As String

    With vsEPR
        For i = 1 To rsInput.RecordCount
            If i > 1 Then
                .AddItem "", lngRow + 1
                lngRow = lngRow + 1
            End If
            If optEPRType(0).Value Then
                strTmp = "OLD" '�ɰ�
            Else
                strTmp = "NEW" '�°�
            End If
            .RowData(lngRow) = strTmp & "|" & rsInput!ID   '�°�ID��32λ�ַ���
            .TextMatrix(lngRow, 0) = rsInput!����
            .Cell(flexcpData, lngRow, 0) = .TextMatrix(lngRow, 0)

            strItem = strItem & "��" & rsInput!����

            rsInput.MoveNext
        Next

        'ȱʡ��Ŀ����
        If txtItem.Text = "" Then txtItem.Text = "��д" & Mid(strItem, 2)

        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If

        mblnChange = True
    End With
End Sub

Private Sub EPREnterNextCell(Optional ByVal blnNewRow As Boolean)
    Dim i As Long, j As Long

    With vsEPR
        If blnNewRow Then
            .Row = .Rows - 1: .Col = 0
            .ShowCell .Row, .Col
        Else
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long

    With vsResult
        If Col = col������� Then
            .TextMatrix(Row, Col) = zlCommFun.GetNeedName(.TextMatrix(Row, Col))
            If .TextMatrix(Row, Col) <> "" Then
                mrsNature.Filter = "����='" & .TextMatrix(Row, Col) & "'"
                Set .Cell(flexcpPicture, .Row, colִ��ͼ��) = imgNature.ListImages(Val(mrsNature!����)).Picture
            Else
                Set .Cell(flexcpPicture, .Row, colִ��ͼ��) = Nothing
            End If
        ElseIf Col = colȱʡ��� Then
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                For i = .FixedRows To .Rows - 1
                    If i <> Row Then .TextMatrix(i, colȱʡ���) = 0
                Next
            End If
        End If
    End With
End Sub

Private Sub vsResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsResult
        If Not ResultCellEditable(NewRow, NewCol) Then
            .FocusRect = flexFocusLight
            .ComboList = ""
        ElseIf NewCol = colִ�н�� Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        ElseIf NewCol = col������� Then
            .ComboList = .ColData(NewCol)
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsResult_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI

    With vsResult
        '������Ӳ�ѯ������������˳�򲻶ԣ���Ҫ�ر�����
        strSql = " Select A.���� As ID,A.�ϼ� As �ϼ�id,A.����,A.����,A.����,Nvl(A.ĩ��,0) As ĩ��,B.���� As ����" & _
                 " From ·��������� A,·��������� B Where Nvl(A.����,0)=B.����(+)" & _
                 " Start With A.�ϼ� Is Null Connect By Prior A.����=A.�ϼ�"
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 2, "�������", True, "", "", False, True, True, _
            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�г���������ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call SetResultInput(Row, rsTmp)
            Call ResultEnterNextCell
        End If
    End With
End Sub

Private Sub vsResult_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long

    If Col = col������� Then
        With vsResult
            If .TextMatrix(Row, Col) <> "" Then
                For i = 0 To .ComboCount - 1
                    If zlCommFun.GetNeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                        .ComboIndex = i: Exit For
                    End If
                Next
            End If
        End With
    End If
End Sub

Private Sub vsResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long

    If mblnReadOnly Or mblnAdjust Then Exit Sub

    With vsResult
        If KeyCode = vbKeyDelete Then
            If .Col = col������� Then
                .TextMatrix(.Row, .Col) = ""
                Set .Cell(flexcpPicture, .Row, colִ��ͼ��) = Nothing
            ElseIf .TextMatrix(.Row, colִ�н��) <> "" Then
                If MsgBox("ȷʵҪ������н����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsResult_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsResult_KeyPress(KeyAscii As Integer)
    With vsResult
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call ResultEnterNextCell
        ElseIf KeyAscii = Asc(",") Then
            KeyAscii = 0
        ElseIf .Col = colִ�н�� Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsResult_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsResult_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    If KeyAscii = Asc(",") Then KeyAscii = 0
End Sub

Private Sub vsResult_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsResult.EditSelStart = 0
    vsResult.EditSelLength = zlCommFun.ActualLen(vsResult.EditText)
End Sub

Private Sub vsResult_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not ResultCellEditable(Row, Col) Then Cancel = True
End Sub

Private Function ResultCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    ResultCellEditable = True

    With vsResult
        If lngCol = colִ��ͼ�� Then
            ResultCellEditable = False
        ElseIf lngCol > colִ�н�� And .TextMatrix(lngRow, colִ�н��) = "" Then
            ResultCellEditable = False
        ElseIf lngCol = col������� And .TextMatrix(lngRow, colִ�н��) <> "" Then
            '�ֵ��еĽ�����ʲ��������,�ֹ�����Ĳ�����
            If .TextMatrix(lngRow, col�������) <> "" Then
                mrsResult.Filter = "����='" & .TextMatrix(lngRow, colִ�н��) & "'"
                If Not mrsResult.EOF Then
                    If NVL(mrsResult!����) = .TextMatrix(lngRow, col�������) Then
                        ResultCellEditable = False
                    End If
                End If
            End If
        End If
    End With
End Function

Private Sub vsResult_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI

    With vsResult
        If Col = colִ�н�� Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call ResultEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call ResultEnterNextCell
            Else
                strInput = UCase(.EditText)
                strSql = "Select A.���� as ID,A.����,A.����,A.����,B.���� as ����" & _
                    " From ·��������� A,·��������� B" & _
                    " Where Nvl(A.����,0)=B.����(+) And A.ĩ��=1" & _
                    " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                    " Order by A.����"
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "�������", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%")
                If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    Cancel = True
                Else
                    Call SetResultInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call ResultEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = col������� Then
            If mblnReturn Then Call ResultEnterNextCell
            mblnReturn = False
        End If
    End With
End Sub

Private Sub SetResultInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������Ŀ���������
    Dim i As Long

    With vsResult
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                End If

                .TextMatrix(lngRow, colִ�н��) = rsInput!����

                '����������
                If Not IsNull(rsInput!����) Then
                    mrsNature.Filter = "����='" & rsInput!���� & "'"
                    Set .Cell(flexcpPicture, lngRow, colִ��ͼ��) = imgNature.ListImages(Val(mrsNature!����)).Picture
                    .TextMatrix(lngRow, col�������) = rsInput!����
                End If

                If i = 1 And lngRow = .FixedRows Then
                    .TextMatrix(lngRow, colȱʡ���) = 1
                    Call vsResult_AfterEdit(lngRow, colȱʡ���)
                End If

                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, colִ�н��) = .EditText
        End If
        .Cell(flexcpData, lngRow, colִ�н��) = .TextMatrix(lngRow, colִ�н��)

        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If

        mblnChange = True
    End With
End Sub

Private Sub ResultEnterNextCell()
    With vsResult
        If .Col + 1 <= .Cols - 1 Then
            .Col = .Col + 1
        ElseIf .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1: .Col = colִ�н��
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub
