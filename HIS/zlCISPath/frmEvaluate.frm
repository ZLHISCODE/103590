VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEvaluate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "·������"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9945
   Icon            =   "frmEvaluate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   9945
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgNature 
      Left            =   8280
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvaluate.frx":617A
            Key             =   "Selected"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvaluate.frx":6514
            Key             =   "UnSelected"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvaluate.frx":68AE
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvaluate.frx":6C48
            Key             =   "UnCheck"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraStart 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   0
      TabIndex        =   27
      Top             =   2640
      Width           =   9855
      Begin VB.OptionButton optStart 
         Caption         =   "���ʱ����Ϊ�뾶��1��,Ŀǰ���뾶��n��"
         Height          =   250
         Index           =   0
         Left            =   960
         TabIndex        =   29
         Top             =   30
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.OptionButton optStart 
         Caption         =   "��ǰʱ����Ϊ�뾶��1��"
         Height          =   250
         Index           =   1
         Left            =   5040
         TabIndex        =   28
         Top             =   30
         Width           =   2895
      End
      Begin VB.Label lblStart 
         Caption         =   "·�����"
         Height          =   230
         Left            =   120
         TabIndex        =   30
         Top             =   45
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   9
         X1              =   120
         X2              =   10000
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   8
         X1              =   120
         X2              =   10000
         Y1              =   315
         Y2              =   315
      End
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   0
      TabIndex        =   21
      Top             =   3000
      Width           =   9855
      Begin VB.OptionButton optDate 
         Caption         =   "��һ�׶���ǰ������"
         Height          =   250
         Index           =   3
         Left            =   4680
         TabIndex        =   36
         Top             =   35
         Width           =   1935
      End
      Begin VB.OptionButton optDate 
         Caption         =   "��һ��Ľ׶��Ӻ󣨼�����ǰ�׶Σ�"
         Height          =   250
         Index           =   2
         Left            =   6720
         TabIndex        =   25
         Top             =   35
         Width           =   3255
      End
      Begin VB.OptionButton optDate 
         Caption         =   "��һ�׶���ǰ������"
         Height          =   250
         Index           =   1
         Left            =   2640
         TabIndex        =   24
         Top             =   35
         Width           =   1935
      End
      Begin VB.OptionButton optDate 
         Caption         =   "����������һ��"
         Height          =   250
         Index           =   0
         Left            =   960
         TabIndex        =   23
         Top             =   35
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   120
         X2              =   10000
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   120
         X2              =   10000
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lblDate 
         Caption         =   "ʱ�����"
         Height          =   230
         Left            =   120
         TabIndex        =   22
         Top             =   45
         Width           =   735
      End
   End
   Begin VB.Frame fraResult 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      TabIndex        =   13
      Top             =   3360
      Width           =   9855
      Begin VB.OptionButton optResult 
         Caption         =   "��������(&3)"
         Height          =   250
         Index           =   3
         Left            =   7680
         TabIndex        =   4
         Top             =   20
         Width           =   1575
      End
      Begin VB.OptionButton optResult 
         Caption         =   "������˳�(&2)"
         Height          =   250
         Index           =   2
         Left            =   5480
         TabIndex        =   3
         Top             =   20
         Width           =   1575
      End
      Begin VB.OptionButton optResult 
         Caption         =   "������(��������)"
         Height          =   250
         Index           =   1
         Left            =   2920
         TabIndex        =   2
         Top             =   20
         Width           =   1935
      End
      Begin VB.OptionButton optResult 
         Caption         =   "����(����)"
         Height          =   250
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   20
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   4
         X1              =   120
         X2              =   10000
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   5
         X1              =   120
         X2              =   10000
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Label lblResult 
         Caption         =   "������"
         Height          =   230
         Left            =   120
         TabIndex        =   14
         Top             =   30
         Width           =   855
      End
   End
   Begin VB.Frame fraRemark 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   7560
      TabIndex        =   20
      Top             =   3800
      Width           =   2295
      Begin VB.TextBox txtRemark 
         Height          =   2175
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   650
         Width           =   2415
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPersonnel 
         Height          =   1305
         Left            =   0
         TabIndex        =   10
         Top             =   2895
         Width           =   2415
         _cx             =   4260
         _cy             =   2302
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   320
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEvaluate.frx":6FE2
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         OwnerDraw       =   0
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblRemark 
         Caption         =   "2012-12-12������ע(&R)"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame fraVariation 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   120
      TabIndex        =   18
      Top             =   3800
      Width           =   7335
      Begin VB.TextBox txtVariation 
         Height          =   300
         Left            =   4245
         MaxLength       =   1000
         TabIndex        =   6
         Top             =   15
         Width           =   2970
      End
      Begin VSFlex8Ctl.VSFlexGrid vsVariation 
         Height          =   3855
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   7215
         _cx             =   12726
         _cy             =   6800
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   320
         ColWidthMin     =   0
         ColWidthMax     =   8000
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmEvaluate.frx":701C
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         OwnerDraw       =   0
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblVariation 
         Caption         =   "����ԭ��"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   60
         Width           =   3375
      End
      Begin VB.Label lblSearch 
         Caption         =   "����(&F)"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   60
         Width           =   735
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
      ScaleWidth      =   9945
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8055
      Width           =   9945
      Begin VB.CommandButton cmdFee 
         Caption         =   "��������(&F)"
         Height          =   350
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdTurn 
         Caption         =   "·����ת(&T)"
         Height          =   350
         Left            =   1440
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdMergeOver 
         Caption         =   "�����ϲ�·��(&M)"
         Height          =   350
         Left            =   2760
         TabIndex        =   33
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8760
         TabIndex        =   12
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   7560
         TabIndex        =   11
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblMsg 
         BackColor       =   &H00EFF0E0&
         Height          =   255
         Left            =   4440
         TabIndex        =   34
         Top             =   215
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsCriterion 
      Height          =   2130
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9735
      _cx             =   17171
      _cy             =   3757
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   8000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEvaluate.frx":7081
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      OwnerDraw       =   0
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9945
      TabIndex        =   15
      Top             =   0
      Width           =   9945
      Begin VB.Label lblNoteOne 
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         Height          =   255
         Left            =   960
         TabIndex        =   35
         Top             =   600
         Width           =   8895
      End
      Begin VB.Label lblPathTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "·��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   120
         Width           =   7695
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "����˵����׶�����˵��"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   400
         Width           =   8895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   10000
         Y1              =   800
         Y2              =   800
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   120
         Picture         =   "frmEvaluate.frx":7108
         Top             =   45
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmEvaluate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ���
Private mlngFun As Long '0-��������,1-�׶�����
Private mlngState As Byte '0-�鿴(��������),1-����,2-�޸�(�׶�����)
                            '�����������ṩ�޸ģ�Ҫ��ֻ��ȡ�����룬���µ��롣�׶������Ĳ鿴ͨ����������ʵ��,�ݲ��ṩȡ������

Private mPP As TYPE_PATH_Pati
Private mPati As TYPE_Pati
Private mstrPath As String '��ǰ�����·��������
Private mlngDiagnosisType As Long '�������:1-��ҽ�������;2-��ҽ��Ժ���;11-��ҽ�������;12-��ҽ��Ժ���
Private mlngDiagnosisSorce As Long '�����Դ1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����
Private mlng����ID As Long
Private mlng���ID As Long
Private mlngType As Long    '=1 �ϲ�·�� =0 ��Ҫ·��
Private mlng��Ҫ·����¼ID As Long    '��Ҫ·����¼��ID
Private mlngMergeID As Long          '�鿴�ϲ�·���ĵ�������
Private mbln��¼���� As Boolean    '=True �ǲ�¼����,False=�ǲ�¼����
Public mblnImp As Boolean                  '��Ϸ���ʱ�Ƿ���������

'ģ�����
Private mrsCondition As ADODB.Recordset
Private mbln��Ŀ������� As Boolean
Private mcol As Collection

Private Enum CNAME
    c��� = 0
    c���� = 1
End Enum

Private Enum CONST_COL_����ԭ��
    col������� = 0
    col����ԭ�� = 1
    col����ѡ�� = 2
End Enum

Private mblnOK As Boolean
Private mobjParent As Object
Private mcolSQL As New Collection
Private mrsMerge As Recordset
Private mblnPathSend As Boolean


Public Function ShowMe(frmParent As Object, ByVal lngFun As Long, ByVal lngState As Long, t_pati As TYPE_Pati, t_pp As TYPE_PATH_Pati, _
    Optional strPath As String, Optional lngDiagnosisType As Long, Optional lngDiagnosisSorce As Long, Optional ByVal lng����ID As Long, _
    Optional ByVal lng���ID As Long, Optional ByVal lngType As Long, Optional ByVal lng��Ҫ·����¼ID As Long, Optional ByVal lngMergeID As Long, _
    Optional ByVal bln��¼ As Boolean = False) As Boolean
    '����:bln��¼  -True ��¼����
    
    mlngFun = lngFun
    mlngState = lngState
    mPati = t_pati
    mlng��Ҫ·����¼ID = lng��Ҫ·����¼ID
    mPP = t_pp
    mstrPath = strPath  '����ʱ����
    mlngDiagnosisType = lngDiagnosisType '����ʱ����
    mlngDiagnosisSorce = lngDiagnosisSorce '����ʱ����
    mlng����ID = lng����ID
    mlng���ID = lng���ID
    mlngType = lngType
    mlngMergeID = lngMergeID
    mbln��¼���� = bln��¼
    
    Set mobjParent = frmParent
        
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Function GetCondition(lng����ID As Long) As ADODB.Recordset
'���ܣ���ȡ·����������
    Dim strSql As String
    
    On Error GoTo errH
    If mlngFun = 0 Then
        strSql = "Select a.ָ��ID,a.��ϵʽ, a.����ֵ, a.�������" & vbNewLine & _
                "From ·���������� A" & vbNewLine & _
                "Where a.����ID = [1]"
        Set GetCondition = zlDatabase.OpenSQLRecord(strSql, "��ȡָ������", lng����ID)
    Else
        strSql = "Select a.ָ��ID, a.��ϵʽ, a.����ֵ, a.�������, Nvl(a.��ĿID,0) as ��ĿID,Nvl(b.ִ�н��,'�޽��') as ִ�н��,B.��Ŀ���� " & vbNewLine & _
            "From ·���������� A, (Select A.��ĿID, A.ִ�н��, B.��Ŀ���� From ����·��ִ�� A,�ٴ�·����Ŀ B" & vbNewLine & _
            "   Where A.·����¼ID = [2] And A.�׶�ID = [3] And A.���� = [4] And A.��ĿId = B.Id) B" & vbNewLine & _
            "Where a.��ĿID = b.��ĿID(+) And a.����ID = [1]"
            
        Set GetCondition = zlDatabase.OpenSQLRecord(strSql, "��ȡָ������", lng����ID, mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetCriterion() As ADODB.Recordset
'���ܣ���ȡ���������ͽ׶�������ָ�궨��
    Dim strSql As String
 
    strSql = "Select a.ID ����ID, b.ID ָ��ID,b.���, b.����ָ��, b.ָ����,b.ָ������" & vbNewLine & _
            "From �ٴ�·������ A, ·������ָ�� B" & vbNewLine & _
            "Where a.·��id = [1] And a.�汾�� = [2] And a.Id = b.����id And a.�������� = [3]" & IIf(mlngFun = 1, " And a.�׶�id = [4]", "") & vbNewLine & _
            "Order by ���"
    On Error GoTo errH
    Set GetCriterion = zlDatabase.OpenSQLRecord(strSql, "��ȡ·��ָ��", mPP.·��ID, mPP.�汾��, mlngFun + 1, mPP.��ǰ�׶�ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatiCriterion() As ADODB.Recordset
'���ܣ�mlngfun=0��ȡ����·�������������,
'      mlngfun=1�޸Ľ׶�����
    Dim strSql As String
    
    On Error GoTo errH
    If mlngFun = 0 Then
        If mlngMergeID = 0 Then
            strSql = "Select a.����˵��,a.δ����ԭ��, a.״̬,b.����ָ��, b.ָ����" & vbNewLine & _
                    "From �����ٴ�·�� A, ����·��ָ�� B" & vbNewLine & _
                    "Where a.id = [1] And a.id = b.·����¼id(+) And b.��������(+)=1"
        Else
            strSql = "Select a.����˵��,Null as δ����ԭ��,1 as ״̬,b.����ָ��, b.ָ����" & vbNewLine & _
                    "From ���˺ϲ�·�� A, ����·��ָ�� B" & vbNewLine & _
                    "Where a.id = [2] And a.id = b.�ϲ�·����¼ID(+) And b.��������(+)=1"
        End If
        Set GetPatiCriterion = zlDatabase.OpenSQLRecord(strSql, "��ȡ����·��ָ��", mPP.����·��ID, mlngMergeID)
    Else
        strSql = "Select a.�������,a.����ԭ��, Nvl(a.ʱ�����,0) as ʱ�����, a.����˵��,a.������,b.����ָ��,b.ָ����" & vbNewLine & _
                "From ����·������ A, ����·��ָ�� B" & vbNewLine & _
                "Where a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3]" & vbNewLine & _
                "And a.·����¼id = b.·����¼id(+) And a.�׶�id=b.�׶�id(+) And a.����=b.����(+) And b.��������(+)=2"
    
        Set GetPatiCriterion = zlDatabase.OpenSQLRecord(strSql, "��ȡ����·��ָ��", mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetMoneyInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡָ�����˵�ʣ���
'������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    strSql = "Select Nvl(�������,0) as �������,Nvl(Ԥ�����,0) as Ԥ�����" & _
            " From ������� Where ����=1 And ���� = 2 And ����ID=[1] " & _
            " Union All Select -1*Nvl(Sum(���),0) as �������,0 as Ԥ�����" & _
            " From ����ģ�����" & _
            " Where ����ID=[1] And ��ҳID=[2]"
    strSql = "Select Sum(�������) as �������,Sum(Ԥ�����) as Ԥ����� From (" & strSql & ")"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then Set GetMoneyInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdFee_Click()
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strInfo As String, strThisInfo As String, lngDay As Long, lngLen As Long, DatIn As Date
    Dim lng�׶�ID As Long, str�׶����� As String
    Dim cur��� As Currency, cur���ϼ� As Currency
    
    Set rsTmp = GetMoneyInfo(mPati.����ID, mPati.��ҳID)
    If Not rsTmp Is Nothing Then
        strInfo = "δ����ã�" & Format(rsTmp!�������, "0.00") & ",Ԥ����" & Format(rsTmp!Ԥ�����, "0.00")
    End If
    
    If mPP.��ǰ�׶η�֧ID = 0 Then
        strSql = "Select ��׼סԺ�� From �ٴ�·���汾 Where ·��id = [1] And �汾�� = [2]"
    Else
        strSql = "Select ��׼סԺ�� From �ٴ�·����֧ Where ·��id = [1] And �汾�� = [2] And ID=[3]"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.·��ID, mPP.�汾��, mPP.��ǰ�׶η�֧ID)
    If InStr(rsTmp!��׼סԺ��, "-") > 0 Then
        lngLen = Split(rsTmp!��׼סԺ��, "-")(0)
    Else
        lngLen = Val(rsTmp!��׼סԺ��)
    End If
    strInfo = strInfo & vbCrLf & "����׼סԺ��" & lngLen & "����㣬���������ķ���(������ѡ��Ŀ)��"
    DatIn = GetPatiInPath(mPati, mPP.����·��ID)
    
    For lngDay = mPP.��ǰ���� + 1 To lngLen
        lng�׶�ID = GetPhaseByDay(mPP.·��ID, mPP.�汾��, lngDay, str�׶�����)
        
        cur��� = GetChargeOfDay(lng�׶�ID, lngDay, DatIn)
        cur���ϼ� = cur���ϼ� + cur���
       
        strThisInfo = "��" & lngDay & "�죺" & IIf(lngDay < 10, Space(2), "") & Format(cur���, "0.00")
        
        If lngLen > 10 And (lngDay Mod 2) = 0 And lngDay <> mPP.��ǰ���� + 1 Then
            strInfo = strInfo & vbTab & vbTab & strThisInfo
        Else
            strInfo = strInfo & vbCrLf & strThisInfo
        End If
    Next
    strInfo = strInfo & vbCrLf & "���ƣ�" & Space(4) & Format(cur���ϼ�, "0.00")
    MsgBox strInfo, vbInformation + vbOKOnly, gstrSysName
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPhaseByDay(ByVal lng·��ID As Long, ByVal lng�汾�� As Long, ByVal lng���� As Long, str�׶����� As String) As Long
'���ܣ���ȡָ��������Ӧ��ȱʡ�׶�ID
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select ID,����" & vbNewLine & _
            "From �ٴ�·���׶�" & vbNewLine & _
            "Where ·��id = [1] And �汾�� = [2] And" & vbNewLine & _
            "      (([3] Between ��ʼ���� And ��������) Or (��ʼ���� = [3] And �������� Is Null) Or (��ʼ���� Is Null And �������� Is Null))" & vbNewLine & _
            "Order By Decode(��ʼ����, Null, 1, 0),���"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ID, lng�汾��, lng����)
    GetPhaseByDay = rsTmp!ID
    str�׶����� = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetChargeOfDay(ByVal lng�׶�ID As Long, ByVal lng���� As Long, ByVal DatIn As Date) As Long
'���ܣ���ȡָ��������Ӧ��ȱʡ�׶�ID
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select Zl_Getpathcharge([1],[2],[3],[4],[5],[6],[7]) as ��� From dual"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.����ID, mPati.��ҳID, mPP.·��ID, mPP.�汾��, lng�׶�ID, lng����, DatIn)
    GetChargeOfDay = Val("" & rsTmp!���)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function MakeMergePath() As ADODB.Recordset
    Set MakeMergePath = New ADODB.Recordset
    
    MakeMergePath.Fields.Append "Id", adBigInt
    MakeMergePath.Fields.Append "·��id", adBigInt, , adFldIsNullable
    MakeMergePath.Fields.Append "�汾��", adBigInt, , adFldIsNullable
    MakeMergePath.Fields.Append "����", adVarChar, 100, adFldIsNullable
    
    MakeMergePath.Fields.Append "����", adVarChar, 50, adFldIsNullable
    MakeMergePath.Fields.Append "˵��", adVarChar, 500, adFldIsNullable
    MakeMergePath.Fields.Append "״̬", adBigInt, , adFldIsNullable   '=0δ�ﵽ��׼סԺ��,=1�ﵽ��׼סԺ�գ���δ�����һ�죬=2��ǰΪ��׼סԺ�����һ��
    MakeMergePath.Fields.Append "ѡ��", adBigInt, , adFldIsNullable   '=1ѡ��
    MakeMergePath.Fields.Append "��ʾ", adBigInt, , adFldIsNullable   '=1��ʾ
    MakeMergePath.Fields.Append "�����޸�", adBigInt, , adFldIsNullable   '0=����=1������
    
    MakeMergePath.CursorLocation = adUseClient
    MakeMergePath.LockType = adLockOptimistic
    MakeMergePath.CursorType = adOpenStatic
    MakeMergePath.Open
End Function

Private Sub cmdMergeOver_Click()
'���ܣ�ѡ����Ҫ��ɵĺϲ�·��
    Dim lngCount As Long
    Dim rsTmp As Recordset
    Dim objPathImport As New frmPathImport
    Dim t_pp As TYPE_PATH_Pati
    'Ϊ�������ʱ���ֶ��·��������⣬
    'ԭ��ֱ����frmPathImport.showme�൱���õ�ȫ�ֶ��󣬶�����newһ��������ʹ�ã����Ǳ���͵���ʱʹ�õĵ��봰����ֳ�ͻ�����������������봰�壬������ʱ�ǵ��봰������������壩��
    
    If mrsMerge.RecordCount > 0 Then
        mrsMerge.MoveFirst
        Set rsTmp = zlDatabase.CopyNewRec(mrsMerge)
        Do While Not mrsMerge.EOF
            '����ﵽ��׼סԺ�գ�����ʾ����
            If Val(mrsMerge!״̬ & "") = 1 Or Val(mrsMerge!״̬ & "") = 2 Or optDate(1).Value Or optDate(1).Enabled = False Then
                mrsMerge!��ʾ = 1
            Else
                mrsMerge!��ʾ = 0
            End If
            
            '����ﵽ���һ�죬��Ϊû��ѡ����һ�׶��Ӻ�����빴ѡ
            If Val(mrsMerge!״̬ & "") = 2 And Not optDate(2).Value Then
                mrsMerge!ѡ�� = 1
                mrsMerge!�����޸� = 1
            Else
                mrsMerge!�����޸� = 0
            End If
            mrsMerge.MoveNext
        Loop
        mrsMerge.MoveFirst
        If Not objPathImport.ShowMe(Me, mPati, 4, t_pp, , , , , , mrsMerge, True) Then
            '��ԭ��Ϣ
            Set mrsMerge = rsTmp
            Exit Sub
        End If
        Do While Not mrsMerge.EOF
            If Val(mrsMerge!ѡ�� & "") = 1 Then lngCount = lngCount + 1
            mrsMerge.MoveNext
        Loop
        mrsMerge.MoveFirst
        lblMSG.Caption = "��ѡ��" & lngCount & "���ϲ�·����"
    Else
        MsgBox "û�пɽ����ĺϲ�·����", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdTurn_Click()
    Dim lngPathID As Long, lngPathVersion As Long
    Dim str����� As String, blnTrnAduit As Boolean
    Dim objPathImport As New frmPathImport
    Dim t_pp As TYPE_PATH_Pati
    
    If InStr(cmdTurn.Tag, ",") > 0 Then
        lngPathID = Split(cmdTurn.Tag, ",")(0)
        lngPathVersion = Split(cmdTurn.Tag, ",")(1)
    End If
    
    
    If objPathImport.ShowMe(Me, mPati, 1, t_pp, mPP.·��ID, lngPathID, lngPathVersion) Then
        '���û��Ȩ�޵Ļ��ȼ��ǰ��׶��Ƿ�����ת��δ��˵�����
        If InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";��ת���;") = 0 Then
            If CheckPathIsTurnAduit Then
                str����� = zlDatabase.UserIdentify(Me, "ǰ��׶δ���δ��˵�·����ת��������˺�����������", glngSys, p�ٴ�·��Ӧ��, "��ת���")
                If str����� = "" Then Exit Sub
                blnTrnAduit = True
            Else
                str����� = zlDatabase.UserIdentify(Me, "��û����ת��˵�Ȩ�ޣ��Ƿ����ھͽ�����ˣ�", glngSys, p�ٴ�·��Ӧ��, "��ת���")
            End If
        Else
            str����� = UserInfo.����
        End If
        cmdTurn.Tag = lngPathID & "," & lngPathVersion & "," & str����� & "," & IIf(blnTrnAduit, "1", "0")
        Call cmdOK_Click
    Else
        cmdTurn.Tag = ""
    End If
End Sub

Private Function CheckPathIsTurnAduit() As Boolean
'���ܣ�����Ƿ����δ��˵���ת�׶Ρ�trueΪ����
     Dim strSql As String, rsTmp As Recordset
     
     strSql = "Select 1 From ����·������ Where ԭ·��id is not null And ��ת����� is null And ·����¼ID=[1]"
     
     On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ת���", mPP.����·��ID)
     
     CheckPathIsTurnAduit = rsTmp.RecordCount > 0
     Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If vsCriterion.Visible And vsCriterion.Enabled And vsCriterion.Rows > vsCriterion.FixedRows Then
        vsCriterion.SetFocus
    Else
        If txtRemark.Visible And txtRemark.Enabled Then txtRemark.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '����������ָ�����������
    End If
End Sub

Private Sub InitFace()
'���ܣ���ʼ����������
    Dim i As Integer, lngMin As Long, lngMax As Long, lngסԺ���� As Long
    Dim strPathin As String, strPatientin As String
    Dim strSql As String, rsTmp As Recordset
    Dim lngState As Long
        
    '1.���������ʼ
    
    fraVariation.BackColor = Me.BackColor
    fraRemark.BackColor = Me.BackColor
    fraDate.BackColor = Me.BackColor
    fraResult.BackColor = Me.BackColor
    fraStart.BackColor = Me.BackColor
    
    lblResult.Tag = "������"
    fraStart.Visible = mlngFun = 0
    fraDate.Visible = mlngFun = 1
    lblNoteOne.Visible = False
    
    If mlngFun = 0 Then
        fraStart.Top = fraDate.Top
        If mlngState = 0 Then
        '�鿴
            strPathin = Format(GetPatiInPath(mPati, mPP.����·��ID), "YYYY-MM-DD")
            strPatientin = Format(GetPatiInDate(mPati, lngסԺ����), "YYYY-MM-DD")
            If strPathin = strPatientin Then
                optStart(0).Value = True
            Else
                optStart(1).Value = True
            End If
            
            optStart(0).Caption = "���ʱ����Ϊ�뾶��1��"
            optStart(1).Caption = "�Ե��뵱����Ϊ�뾶��1��"
            optStart(0).Enabled = False
            optStart(1).Enabled = False
        Else
        '���ǵ�����Ժ�ģ�����ѡ����ʼʱ��
            Call GetPatiInDate(mPati, lngסԺ����)
            If lngסԺ���� <= 1 Or mblnPathSend Then
                fraStart.Visible = False
                fraStart.Tag = "���ɼ�"
                vsCriterion.Height = vsCriterion.Height + fraStart.Height
            Else
                optStart(0).Caption = "���ʱ����Ϊ�뾶��1��,Ŀǰ���뾶��" & lngסԺ���� & "��"
            End If
        End If
        
        If mlngType = 1 And fraStart.Tag <> "���ɼ�" Then
            '�ϲ�·��ֻ���Ե�ǰ������Ϊ��һ��
            fraStart.Visible = False
            fraStart.Tag = "���ɼ�"
            vsCriterion.Height = vsCriterion.Height + fraStart.Height
        End If
        If mlngType = 1 Then txtRemark.Enabled = False: lblVariation.Caption = "�ϲ�·������д������ԭ��"
        
        Me.Caption = "��������"
        lblResult.Caption = "������"
        lblPathTitle.Caption = "����·����" & mstrPath
        lblNote.Caption = "��ѡ����������������ϵ�����������ѡ��ԭ����д˵�����Ա����ͳ�Ʒ�����"
        lblRemark.Caption = "��ע(&R)"
        optResult(0).Caption = "����(&0)": optResult(1).Caption = "������(&1)"
        If (Not mblnImp) And mlngFun = 0 Then
            optResult(1).Enabled = False
        End If
        optResult(2).Visible = False
        optResult(3).Visible = False
        
        cmdFee.Visible = False
        vsPersonnel.Visible = False
        txtRemark.Height = fraRemark.Height - lblRemark.Height - 60
        lblRemark.Top = 0
        txtRemark.Top = lblRemark.Top + lblRemark.Height + 30
        
        cmdMergeOver.Visible = False
        
        If mlngState = 0 Then   '�鿴
            optResult(0).Enabled = False
            optResult(1).Enabled = False
            txtRemark.Enabled = False
            vsVariation.Enabled = False
            txtVariation.Enabled = False
            
            cmdOK.Visible = False
            cmdCancel.Caption = "�˳�(&X)"
        Else
            If mlngType = 0 Then
                cmdOK.Left = cmdCancel.Left
                cmdCancel.Visible = False
            End If
        End If
    Else
        cmdTurn.Visible = mlngState = 1 '����ʱѡ������ת·���������Ʋ������޸ģ�ֻ��ȡ������
        lblPathTitle.Visible = False
        lblNote.Top = lblPathTitle.Top
        lblNote.Height = 400
        lblNote.Caption = "����ݲ��˵ĵ�ǰ��������������Ծ����Ƿ��������·�����ƶ��ļƻ����к�����������������˱��죬��ѡ�����ԭ�򣬲���д����˵�����Ա����ͳ�Ʒ����ͳ����Ľ�·����"
        If mbln��¼���� Then
            '������ǰ���ɵ��µ�ǰ����֮ǰ��¼����ʱ,ֻ��ѡ����������
            optDate(0).Value = True: optDate(1).Enabled = False: optDate(2).Enabled = False: optDate(3).Enabled = False
            
            lblNoteOne.Visible = True
            lblNoteOne.Top = lblNote.Top + lblNote.Height
            lblNoteOne.Caption = "��ǰ�׶�֮��������������Ŀ,Ҫ�����·����ȡ����ǰ���ɵ�·����Ŀ��������������"
            lblNoteOne.ForeColor = vbRed
        Else
            If GetNextPhase(mPP.��ǰ�׶�ID, mPP.��ǰ�׶η�֧ID) = 0 Then 'û�к����׶�ʱ��������ѡ����һ�׶���ǰ
                If optDate(1).Value Then optDate(1).Value = False    '��ǰ������
                optDate(1).Enabled = False
                If optDate(3).Value Then optDate(3).Value = False    '��ǰ������
                optDate(3).Enabled = False
            End If
        End If
        
        If mPP.�ϲ�·������ = 0 Or mlngState <> 1 Then
            cmdMergeOver.Visible = False
        Else
            '�ų������˵ĺϲ�·��
            strSql = "Select a.Id, a.·��id, b.����, b.����, b.˵��,c.�汾��,a.��ǰ�׶�ID,d.��֧ID as ��ǰ�׶η�֧ID,a.��ǰ����" & vbNewLine & _
                    "From ���˺ϲ�·�� A, �ٴ�·��Ŀ¼ B, �ٴ�·���汾 C,�ٴ�·���׶� D" & vbNewLine & _
                    "Where a.·��id = b.Id And a.·��id = c.·��id And a.�汾�� = c.�汾�� And d.id(+)=a.��ǰ�׶�id " & _
                    "  And a.��Ҫ·����¼id = [1] and a.����ʱ�� is Null"

            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
            If rsTmp.RecordCount = 0 Then cmdMergeOver.Visible = False
            If rsTmp.RecordCount > 0 Then
                Set mrsMerge = MakeMergePath
                Do While Not rsTmp.EOF
                    mrsMerge.AddNew
                    mrsMerge!ID = rsTmp!ID
                    mrsMerge!·��ID = Val(rsTmp!·��ID & "")
                    mrsMerge!�汾�� = Val(rsTmp!�汾�� & "")
                    mrsMerge!���� = rsTmp!���� & ""
                    mrsMerge!���� = rsTmp!���� & ""
                    Call IsLastDate(, , , Val(rsTmp!·��ID & ""), Val(rsTmp!�汾�� & ""), Val(rsTmp!��ǰ�׶�ID & ""), Val(rsTmp!��ǰ�׶η�֧ID & ""), Val(rsTmp!��ǰ���� & ""), True, lngState, Val(rsTmp!ID & ""))
                    mrsMerge!״̬ = lngState
                    
                    rsTmp.MoveNext
                Loop
                If mrsMerge.RecordCount > 0 Then mrsMerge.Update: mrsMerge.MoveFirst
            End If
        End If
        
        lblRemark.Caption = Format(mPP.��ǰ����, "YYYY-MM-DD") & "������ע"
        optResult(0).Caption = "����(&0)"
        optResult(1).Caption = "��������(&1)"
        optResult(2).Caption = "������˳�(&2)"
        If mbln��¼���� Then
            optResult(2).Enabled = False
            optResult(3).Enabled = False
            cmdTurn.Enabled = False
        Else
            '�ﵽ��׼סԺ���������ɱ������ɡ�
            '���û�дﵽ��׼סԺ�գ����ṩһ��ѡ���ǰ����
            If IsLastDate(True, lngMin, lngMax, mPP.·��ID, mPP.�汾��, mPP.��ǰ�׶�ID, mPP.��ǰ�׶η�֧ID, mPP.��ǰ����) Then
                optResult(3).Visible = True
                optResult(3).Caption = "��������(&3)"
            Else
                optResult(3).Caption = "��ǰ���(&3)"
                optResult(3).Visible = True
            End If
            If InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";����·��;") = 0 Then
                optResult(2).Visible = False
                optResult(3).Visible = False
            Else
                '89620:�߱�����·������ǰ��ɲ�������ǰ���,�����ֹ��
                If optResult(3).Caption = "��ǰ���(&3)" Then
                    optResult(3).Visible = (InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";��ǰ���;") > 0)
                End If
            End If
            
            '������׼סԺ�պ󣬲���ѡ������
            If mPP.��ǰ���� > lngMax Then
                optResult(1).Value = True
                optResult(0).Enabled = False: optResult(0).Tag = "��ֹѡ������"
            End If
        End If
    End If
    
    '2.����ָ����ʼ(�����ж��������������)
    Call InitVsCriterion
                
    lblResult.Tag = ""
    '3.���ر���ԭ���б�
    For i = 0 To optResult.count - 1
        If optResult(i).Value Then
            Exit For
        End If
    Next
    Call optResult_Click(i)
        
    '4.��ʼ��������
    If mlngFun = 1 Then
    
        With vsPersonnel
            .Redraw = flexRDNone
            .Editable = flexEDKbdMouse
            .Rows = 1
            .Cols = 1
            .TextMatrix(0, 0) = "������"
            
            .Rows = 2
            .TextMatrix(1, 0) = UserInfo.����  'ȱʡΪ��ǰ����Ա
            .Redraw = True
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitVsCriterion()
'���ܣ���ʼ������ָ���
    Dim strcol As String, arrHead As Variant
    Dim i As Long, lng����ID As Long
    Dim rsCriterion As ADODB.Recordset
    Dim blnValue As Boolean, blnThis As Boolean
    
    lng����ID = 0
    Set rsCriterion = GetCriterion
    If rsCriterion.RecordCount > 0 Then
        lng����ID = rsCriterion!����ID
        Set mrsCondition = GetCondition(lng����ID)
        
        strcol = "���,450,4;����ָ��,6800,1;���,900,1;ָ������;ָ����"
        '1.��ʼ������ָ���ͷ
        With vsCriterion
            .Redraw = flexRDNone
            .Clear
            .FixedCols = 1: .FixedRows = 1
            arrHead = Split(strcol, ";")
            .Cols = UBound(arrHead) + 1
            .Rows = .FixedRows
            .Rows = .FixedRows + rsCriterion.RecordCount
            .Editable = flexEDKbdMouse
            Set mcol = New Collection
            
            For i = 0 To UBound(arrHead)
                mcol.Add i, Split(arrHead(i), ",")(0)
                .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
                
                If UBound(Split(arrHead(i), ",")) > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
                    'Ϊ��֧��zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  'Ϊ��֧��zl9PrintMode
                End If
            Next
            
            '2.����ָ���б�
            For i = 1 To rsCriterion.RecordCount
                .RowData(i) = Val(rsCriterion!ָ��ID)
                .TextMatrix(i, mcol("���")) = rsCriterion!���
                .TextMatrix(i, mcol("����ָ��")) = rsCriterion!����ָ��
                .TextMatrix(i, mcol("���")) = Split(rsCriterion!ָ����, vbTab)(1)
                
                .TextMatrix(i, mcol("ָ������")) = rsCriterion!ָ������
                .TextMatrix(i, mcol("ָ����")) = rsCriterion!ָ����
                
                rsCriterion.MoveNext
            Next
            
            .Redraw = flexRDDirect
        End With
    
        '3.���������ָ������������ָ����������ȱʡ���������
        If mlngState = 1 And mrsCondition.RecordCount > 0 Then
            If mlngFun = 0 Then
                Call SetResult
            
            '�׶��������쳣������أ�����ָ�겻��ʱ��ȱʡ��ִ�н��Ϊ�쳣˵��
            ElseIf mlngFun = 1 Then
                With mrsCondition
                    blnValue = False
                    
                    .Filter = "��ĿID<>0"
                    For i = 1 To .RecordCount
                        Select Case !��ϵʽ
                            Case "="
                                blnThis = (!ִ�н�� = !����ֵ)
                            Case "<>"
                                blnThis = (!ִ�н�� <> !����ֵ)
                            Case ">"
                                blnThis = (!ִ�н�� > !����ֵ)
                            Case ">="
                                blnThis = (!ִ�н�� >= !����ֵ)
                            Case "<"
                                blnThis = (!ִ�н�� < !����ֵ)
                            Case "<="
                                blnThis = (!ִ�н�� <= !����ֵ)
                            Case "Like"
                                blnThis = (!ִ�н�� Like "*" & !����ֵ & "*")
                            Case Else
                                blnThis = True
                        End Select
                                        
                        If i = 1 Then
                            blnValue = blnThis
                        Else
                            If !������� = 1 Then
                                blnValue = (blnValue And blnThis)
                            Else
                                blnValue = (blnValue Or blnThis)
                            End If
                        End If
                        
                        .MoveNext
                    Next
                    mbln��Ŀ������� = blnValue
                    
                    If blnValue Or optResult(0).Enabled = False Then '�׶���������������ʱ��ʾ����
                        optResult(1).Value = True   'ȱʡΪ��������
                
                        '�����Ŀִ�н���������������ټ��ָ�������Ƿ����
                        Call SetResult
                    Else
                        optResult(0).Value = True
                    End If
                End With
            End If
        End If
    Else
        'û������ָ��ʱ������ʾָ����
        vsCriterion.Tag = "û������ָ���¼"
        vsCriterion.Visible = False
        If mlngFun = 0 Then
            fraStart.Top = vsCriterion.Top
            fraResult.Top = fraStart.Top + IIf(fraStart.Tag = "���ɼ�", 0, fraStart.Height + 30)
        Else
            fraDate.Top = vsCriterion.Top
            fraResult.Top = fraDate.Top + fraDate.Height + 30
        End If
        fraVariation.Top = fraResult.Top + fraResult.Height
        fraRemark.Top = fraVariation.Top
                
        Me.Height = Me.Height - vsCriterion.Height - 120
    End If
End Sub

Private Sub InitVariation(ByVal lngKind As Long)
'���ܣ���ʼ������ԭ���б�
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    
    strSql = "Select b.���� As ����, a.����, a.����, a.����" & vbNewLine & _
            "From ���쳣��ԭ�� A, ���쳣��ԭ�� B" & vbNewLine & _
            "Where a.ĩ�� = 1 And a.�ϼ� = b.���� and a.����=[1]" & vbNewLine & _
            "order by ����,����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngKind)
        
    With vsVariation
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If rsTmp.RecordCount > 0 Then
            .MergeCol(col�������) = True
            .Rows = .FixedRows + rsTmp.RecordCount
            'ȱʡ��ѡ��
            Set .Cell(flexcpPicture, .FixedRows, col����ѡ��, .Rows - 1, col����ѡ��) = imgNature.ListImages(IIf(mlngFun = 0, "UnSelected", "UnCheck")).Picture
            .Cell(flexcpPictureAlignment, .FixedRows, col����ѡ��, .Rows - 1, col����ѡ��) = flexPicAlignCenterCenter

            For i = .FixedRows To rsTmp.RecordCount
                .Cell(flexcpData, i, col����ѡ��) = 0
                
                .RowData(i) = CStr(rsTmp!����)    '����
                .TextMatrix(i, col�������) = rsTmp!����
                .TextMatrix(i, col����ԭ��) = rsTmp!���� & "-" & rsTmp!����
                .Cell(flexcpData, i, col����ԭ��) = "" & rsTmp!����
                rsTmp.MoveNext
            Next
        End If
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    vsVariation.Redraw = flexRDDirect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mblnPathSend = CheckPathSend(mPati.����ID, mPati.��ҳID)
    mblnImp = InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";��Ϸ�����������;") > 0
    Call InitFace
    Call LoadData
    
End Sub

Private Sub SetFillTableByStr(vstmp As VSFlexGrid, strTmp As String, lngCol As Long)
'���ܣ����ַ�����ֵ���ָ�����䵽����У����ң���δβ����һ����
'������
    Dim i As Long, arrtmp As Variant
    
    arrtmp = Split(strTmp, ",")
    With vstmp
        .Rows = .FixedRows + UBound(arrtmp) + 2 '����һ�п���
        For i = 0 To UBound(arrtmp)
            .TextMatrix(i + .FixedRows, lngCol) = arrtmp(i)
        Next
        .TextMatrix(.Rows - 1, lngCol) = ""
    End With
End Sub

Private Function Get��Ŀ����ԭ��() As ADODB.Recordset
'���ܣ���ȡ·������Ŀ�ı���ԭ��
    Dim strSql As String
    If mlngState = 1 Then
        strSql = "Select distinct ����ԭ�� From (Select ����ԭ�� From ����·��ִ�� " & _
                "Where ·����¼Id = [1] And �׶�ID = [2] And ���� = [3] And ����ԭ�� Is Not Null And Nvl(����ʱ������,0)<2 Order by �Ǽ�ʱ��)"
    ElseIf mlngState = 2 Then
        strSql = "Select ����ԭ�� From ����·������ Where ·����¼Id = [1] And �׶�ID = [2] And ���� = [3] "
    End If
    On Error GoTo errH
    Set Get��Ŀ����ԭ�� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadData()
'���ܣ���������
    Dim i As Long, str����ԭ�� As String
    Dim j As Long
    Dim rsTmp As ADODB.Recordset
                
    If mlngFun = 1 Then
        Set rsTmp = Get��Ŀ����ԭ��
        If rsTmp.RecordCount > 0 Then
            optResult(0).Enabled = False: optResult(0).Tag = "��ֹѡ������"
            optResult(1).Value = True   '��������
        End If
                
        If rsTmp.RecordCount > 0 Then
            For j = 1 To rsTmp.RecordCount
                i = vsVariation.FindRow(CStr(rsTmp!����ԭ��)) '���������rowdata
                If i > 0 Then
                    vsVariation.Row = i
                    vsVariation.TopRow = i
                    Call vsVariation_Click
                End If
                rsTmp.MoveNext
            Next
        End If
    End If
                           
    '1.����ָ����
    '�鿴��������������
    '�����޸�ʱ����ԭ�����������������ָ�꣬Ҳ����û��ָ��
    If mlngFun = 0 And mlngState = 0 Or (mlngFun = 1 And mlngState = 2) Then
        Set rsTmp = GetPatiCriterion
        'һ���м�¼
        If mlngFun = 0 Then
            optResult(0).Value = rsTmp!״̬ = 1
            optResult(1).Value = rsTmp!״̬ <> 1
            
            If Not IsNull(rsTmp!δ����ԭ��) Then
                i = vsVariation.FindRow(CStr(rsTmp!δ����ԭ��)) '���������rowdata
                If i > 0 Then
                    vsVariation.Row = i
                    vsVariation.TopRow = i
                    Call vsVariation_Click
                End If
            End If
            txtRemark.Text = "" & rsTmp!����˵��
        Else
            If rsTmp!ʱ����� = -1 Then
                optDate(2).Value = True     '����click�¼������ù�����optResult�Ŀ�����
            ElseIf rsTmp!ʱ����� = 1 Then
                optDate(1).Value = True
            ElseIf rsTmp!ʱ����� = 2 Then
                optDate(3).Value = True
            Else
                optDate(0).Value = True
            End If
            
            If rsTmp!������� = -1 Then
                If mPP.����·��״̬ = 1 Then
                    optResult(1).Value = True   '���첢������������������ȡ����������ʱ�൱�ڱ���������
                Else
                    optResult(2).Value = True   '�����˳�
                End If
            Else
                optResult(0).Value = True
            End If
            
            txtRemark.Text = "" & rsTmp!����˵��
            Call SetFillTableByStr(vsPersonnel, rsTmp!������, 0)
        End If
            
        '����ָ����
        If vsCriterion.Tag <> "û������ָ���¼" Then
            With vsCriterion
                .Redraw = flexRDNone
                For i = 1 To .Rows - 1
                    rsTmp.Filter = "����ָ��='" & .TextMatrix(i, mcol("����ָ��")) & "'"
                    If rsTmp.RecordCount > 0 Then
                        .TextMatrix(i, mcol("���")) = "" & rsTmp!ָ����
                    Else
                        .TextMatrix(i, mcol("���")) = ""
                    End If
                Next
                .Redraw = flexRDDirect
            End With
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mlngFun = 0 And mblnOK = False And mlngType = 0 And mlngState <> 0 Then
        '��������ʱ��������ȡ����ť��ֻ�ܵ�ȷ��
        Cancel = 1
        Exit Sub
    End If
    
    mbln��Ŀ������� = False
    Set mrsCondition = Nothing
    Set mobjParent = Nothing
    Set mcolSQL = Nothing
End Sub

Private Sub optDate_Click(Index As Integer)
    If Index = 0 Then
        If optResult(0).Tag <> "��ֹѡ������" Then optResult(0).Enabled = True
        optResult(1).Enabled = True
        optResult(2).Enabled = True
        optResult(3).Enabled = True
        If optResult(3).Caption = "��ǰ���(&3)" Then
            'optResult(3).Enabled = False
            If optResult(3).Value Then optResult(0).Value = True
        End If
        
        cmdTurn.Enabled = optResult(1).Value And Not mbln��¼����
    Else
    'ʱ�����ʱ��ֻ��ѡ�����������
        optResult(0).Enabled = False
        optResult(1).Enabled = True
        '���ѡ��ʱ����ǰ������ʹ����ǰ�������ܡ�
        If Index = 1 And optResult(3).Caption = "��ǰ���(&3)" Then
            optResult(3).Enabled = True
            If optResult(0).Value Or optResult(2).Value Then optResult(1).Value = True
        Else
            optResult(3).Enabled = False
            optResult(1).Value = True
        End If
        optResult(2).Enabled = False
        
        
        cmdTurn.Enabled = False
    End If
End Sub

Private Sub optResult_Click(Index As Integer)
    If lblResult.Tag = "������" Then Exit Sub
    
    If mlngFun = 0 Then '����
        Call InitVariation(0)
    Else
        If Index = 1 Or Index = 3 Then '������������
            Call InitVariation(1)
            If Index = 3 And optResult(3).Caption = "��ǰ���(&3)" Then
                optDate(1).Value = True
            End If
        ElseIf Index = 2 Then   '�����˳�
            Call InitVariation(2)
        End If
        
        cmdTurn.Enabled = (Index = 1 And optDate(0).Value) And Not mbln��¼����
    End If
    
    '��������ʱ��ֹ�ñ���ԭ��,�鿴��������ʱҲ����
    If Index = 0 Or mlngState = 0 Or mlngType = 1 Then
        vsVariation.Enabled = False
        vsVariation.BackColor = Me.BackColor
        vsVariation.Row = 0
        txtVariation.Enabled = False
        txtVariation.BackColor = Me.BackColor
    Else
        vsVariation.Enabled = True
        vsVariation.BackColor = &H80000005
        txtVariation.Enabled = True
        txtVariation.BackColor = &H80000005
        
        If vsVariation.Visible And vsVariation.Enabled Then vsVariation.SetFocus
    End If
End Sub

Private Sub txtRemark_GotFocus()
    Call zlControl.TxtSelAll(txtRemark)
End Sub

Private Sub vsCriterion_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = mcol("���") And mlngState = 1 Then    '�޸��������ʱ���ٸ���ָ�������������
        Call SetResult
    End If
End Sub

Private Sub SetResult()
'���ܣ�����ָ����Ŀ�Ľ����������Ľ��
    Dim i As Long, j As Long, strValue As String
    Dim blnValue As Boolean, blnThis As Boolean
    Dim blnFirst As Boolean
        
    blnFirst = True
    If mlngFun = 1 Then
        blnValue = mbln��Ŀ�������
    Else
        blnValue = True
    End If
    For i = 1 To vsCriterion.Rows - 1
        strValue = vsCriterion.TextMatrix(i, mcol("���"))
        If mlngFun = 0 Then
            mrsCondition.Filter = "ָ��ID = " & vsCriterion.RowData(i)
        Else
            mrsCondition.Filter = "ָ��ID = " & vsCriterion.RowData(i) & " And ��ĿID = 0"
        End If
        With mrsCondition
            For j = 1 To .RecordCount
                 Select Case !��ϵʽ
                    Case "="
                        blnThis = (strValue = !����ֵ)
                    Case "<>"
                        blnThis = (strValue <> !����ֵ)
                    Case ">"
                        blnThis = (strValue > !����ֵ)
                    Case ">="
                        blnThis = (strValue >= !����ֵ)
                    Case "<"
                        blnThis = (strValue < !����ֵ)
                    Case "<="
                        blnThis = (strValue <= !����ֵ)
                    Case "Like"
                        blnThis = (strValue Like "*" & !����ֵ & "*")
                    Case Else
                        blnThis = True
                End Select
                
                If blnFirst And mlngFun = 0 Then
                    blnValue = blnThis
                    blnFirst = False
                Else
                    If !������� = 1 Then
                        blnValue = (blnValue And blnThis)
                    Else
                        blnValue = (blnValue Or blnThis)
                    End If
                End If
                .MoveNext
            Next
        End With
    Next
    
    If mlngFun = 0 Then
        If blnValue Then
            optResult(0).Value = True
        Else
            If mblnImp Then
                optResult(1).Value = True
            Else
                optResult(0).Value = True
            End If
        End If
    Else
        If blnValue Then  '�׶���������������ʱ��ʾ����
            optResult(1).Value = True
        Else
            If optResult(0).Enabled Then optResult(0).Value = True  'ѡ������Ӻ����ǰʱ���Լ�������׼סԺ�գ�������ѡ��������
        End If
    End If
End Sub

Private Sub vsCriterion_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Visible Then
        If NewCol = mcol("���") And mlngState <> 0 Then
            Dim arrtmp As Variant
            
            With vsCriterion
                arrtmp = Split(.TextMatrix(NewRow, mcol("ָ����")), vbTab)
                .ColComboList(NewCol) = Replace(arrtmp(0), ",", "|")
            End With
        End If
    End If
End Sub

Private Sub vsCriterion_GotFocus()
'    vsCriterion.ForeColorSel = vbWhite
'    vsCriterion.BackColorSel = &H8000000D
End Sub

Private Sub vsCriterion_LostFocus()
'    vsCriterion.ForeColorSel = vbBlack
'    vsCriterion.BackColorSel = vbWhite
End Sub

Private Sub vsCriterion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call ResultEnterNextCell(vsCriterion)
    End If
End Sub

Private Sub vsCriterion_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mcol("���") Or mlngState = 0 Then
        Cancel = True
    End If
End Sub

Private Sub vsPersonnel_GotFocus()
    If vsPersonnel.Row = vsPersonnel.Rows - 1 Then
        With vsPersonnel
            If .TextMatrix(.Row, .Col) <> "" Then
                Call vsPersonnel_AfterEdit(.Row, .Col)
            End If
        End With
    End If
End Sub

Private Sub vsPersonnel_KeyDown(KeyCode As Integer, Shift As Integer)
'���ܣ�ɾ�����һ�У��������Ԫ������
    If KeyCode = vbKeyDelete Then
        With vsPersonnel
            If .Row = .Rows - 1 And .Row > .FixedRows And .TextMatrix(.Row, 0) = "" Then '������һ��
                .Rows = .Rows - 1
            ElseIf .Row > .FixedRows - 1 Then
                .TextMatrix(.Row, .Col) = ""
            End If
        End With
    End If
End Sub

Private Sub vsPersonnel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ����һ�лس����Զ���һ��
    With vsPersonnel
        If Trim(.TextMatrix(Row, Col)) <> "" And Row = .Rows - 1 Then
            .Rows = .Rows + 1
            .Select .Rows - 1, .Col
        End If
    End With
End Sub

Private Sub vsPersonnel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ResultEnterNextCell(vsPersonnel)
    End If
End Sub

Private Sub vsPersonnel_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strtxt As String, strSql As String, blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset, i As Long
    Dim vPoint As POINTAPI
    
    With vsPersonnel
        strtxt = Trim(.EditText)
        If strtxt = "" Then Exit Sub
        
        If zlCommFun.IsCharAlpha(strtxt) Then
            strtxt = UCase(strtxt)
            strSql = " And a.���� like [1]"
        Else
            strSql = " And a.���� like [1]"
        End If
        strSql = "Select Distinct a.ID,a.��� as ����,a.���� From ��Ա�� a, ��Ա����˵�� b Where a.Id = b.��Աid And b.��Ա���� In ('ҽ��', '��ʿ')" & strSql
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "������", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, False, strtxt & "%")
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "������������δ�ҵ�ƥ���ҽ����ʿ��", vbInformation, gstrSysName
            End If
            Cancel = True
            Exit Sub
        End If
        For i = .FixedCols To .Rows - 1
            If .TextMatrix(i, 0) = rsTmp!���� And i <> .Row Then
                MsgBox "�Ѿ���������ͬ��������Ա��", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
        Next
        
        .EditText = rsTmp!����
    End With
End Sub

Private Sub ResultEnterNextCell(vsthis As VSFlexGrid)
    With vsthis
        If .Col <= .Cols - 1 Then
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    If mlngState = 0 And mlngFun = 0 Then
        mblnOK = True
    Else
        mblnOK = False
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, str������ As String, strTmp As String
    Dim blnOver As Boolean, blnOK As Boolean, lngLen As Long
    Dim strSql As String, str����� As String, strVariation As String
    Dim rsTmp As ADODB.Recordset
    Dim lngMax As Long, lngMin As Long
    Dim str����ԭ�� As String
    Dim str��ת����� As String
    Dim blnTmp As Boolean
    
    '��������ݣ������ѡ��һ������ԭ�򣬱���˵�����Բ���
    If optResult(0).Value = False And vsVariation.Rows > vsVariation.FixedRows Then
        With vsVariation
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, col����ѡ��) = 1 Then
                    strVariation = strVariation & "," & .RowData(i)
                    str����ԭ�� = str����ԭ�� & "," & Mid(.TextMatrix(i, col����ԭ��), InStr(.TextMatrix(i, col����ԭ��), "-") + 1)
                End If
            Next
            strVariation = Mid(strVariation, 2)
            If str����ԭ�� = "" And vsVariation.Enabled Then
                MsgBox "��ѡ��һ�ֱ���ԭ��", vbInformation, gstrSysName
                If vsVariation.Enabled And vsVariation.Visible Then vsVariation.SetFocus
                Exit Sub
            End If
        End With
    End If
    
    '�������ԭ����������Ҫ�������д����˵��
    If InStr(str����ԭ�� & ",", ",����,") > 0 Or InStr(str����ԭ�� & ",", ",����,") > 0 Then
        If Trim(txtRemark.Text) = "" Then
            MsgBox "����ԭ��Ϊ�����ģ�������д������ע��", vbInformation, gstrSysName
            If txtRemark.Enabled Then txtRemark.SetFocus
            Exit Sub
        End If
    End If
    
    If txtRemark.Text <> Trim(txtRemark.Text) Then txtRemark.Text = Trim(txtRemark.Text)
    If mlngFun = 0 Then
        lngLen = Sys.FieldsLength("�����ٴ�·��", "����˵��")
    Else
        lngLen = Sys.FieldsLength("����·������", "����˵��")
    End If
    If zlCommFun.ActualLen(txtRemark.Text) > lngLen Then
        Call MsgBox("��ע��Ϣ���ܳ�����󳤶�" & lngLen, vbInformation, gstrSysName)
        txtRemark.SetFocus
        Exit Sub
    End If
    
    '����ָ��
    If vsCriterion.Visible Then
        With vsCriterion
            For i = .FixedRows To .Rows - 1
                If InStr(.TextMatrix(i, mcol("����ָ��")), "|") > 0 Then
                    MsgBox "��" & i & "�У�����ָ���к��������ַ�:|�����ܱ������ݣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
                    Exit Sub
                End If
                If .TextMatrix(i, mcol("���")) = "" Then
                    MsgBox "��" & i & "�У�����ָ��δ��д�������������д����������", vbInformation, gstrSysName
                    .Select i, mcol("���")
                    Exit Sub
                End If
            Next
        End With
    End If
    
    If mlngFun = 1 Then
        With vsPersonnel
            For i = .FixedRows To .Rows - 1
                strTmp = Trim(.TextMatrix(i, 0))
                If strTmp <> "" Then
                    str������ = str������ & "," & strTmp
                End If
            Next
            str������ = Mid(str������, 2)
        End With
        If str������ = "" Then
            MsgBox "������δ��д������������һ�������ˡ�", vbInformation, gstrSysName
            Exit Sub
        ElseIf LenB(str������) > 50 Then
            MsgBox "������̫�࣬������󳤶�50", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '���������һ���׶ε����һ�죬��������ɺ��Զ�����·��
    If mlngFun = 1 Then
        If optResult(0).Value Or optResult(1).Value Then
            blnOver = IsLastDate(False, lngMin, lngMax, mPP.·��ID, mPP.�汾��, mPP.��ǰ�׶�ID, mPP.��ǰ�׶η�֧ID, mPP.��ǰ����)
            
            If InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";����·��;") = 0 Then
                blnOver = False
            End If
        End If
        
            
        '·����תʱ������
        If cmdTurn.Enabled And cmdTurn.Visible And cmdTurn.Tag <> "" Then
            blnOver = False
        End If
        
        If optResult(0).Value Or optResult(1).Value And optDate(0).Value Then
            '�������ʱ�����������ǰ�׶λ���ǰ������һ�׶Σ��򲻽���·����������������ټ��
            If blnOver Then
                '���жϸ�·���Ƿ�������ϲ�ͬ���·���������������Ժ����Ƿ�͵��������ͬ
                If mPP.����·������ = 0 Then
                    If Not CheckPathOutDiag(mPP.����·��ID, mPati.����ID, mPati.��ҳID) Then
                        MsgBox "��Ժ��ϲ������ò��ַ�Χ�ڣ��������������·����ֻ�ܱ����˳�·����", vbInformation, gstrSysName
                        Exit Sub
                    Else
                        MsgBox "ע�⣺Ŀǰ�Ѵﵽ�򳬹���׼סԺ�գ�����ִ�к��Զ���ɲ���·����", vbInformation, gstrSysName
                    End If
                Else
                    MsgBox "ע�⣺Ŀǰ�Ѵﵽ�򳬹���׼סԺ�գ�����ִ�к��Զ���ɲ���·����", vbInformation, gstrSysName
                End If
            End If
            
        ElseIf optResult(3).Value Then
            blnOver = True
            '���жϸ�·���Ƿ�������ϲ�ͬ���·���������������Ժ����Ƿ�͵��������ͬ
            If mPP.����·������ = 0 Then
                If Not CheckPathOutDiag(mPP.����·��ID, mPati.����ID, mPati.��ҳID) Then
                    MsgBox "��Ժ��ϲ������ò��ַ�Χ�ڣ��������������·����ֻ�ܱ����˳�·����", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Else
            blnOver = False
        End If
        
        '������׼סԺ�ձ������,������˳���Ҫ���
        If optResult(1).Value And mPP.��ǰ���� > lngMax Or optResult(2).Value Then
            If InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";�������;") = 0 Then
                str����� = zlDatabase.UserIdentify(Me, "�����˳����ڼ�����Ҫ��ˡ�", glngSys, p�ٴ�·��Ӧ��, "�������")
                If str����� = "" Then Exit Sub
            Else
                str����� = UserInfo.����
            End If
        End If
    End If
    
    '���ϲ�·��
    If Not mrsMerge Is Nothing And mlngFun = 1 And mlngState = 1 Then
        If mrsMerge.RecordCount > 0 Then
            Set rsTmp = zlDatabase.CopyNewRec(mrsMerge)
            mrsMerge.MoveFirst
            Do While Not mrsMerge.EOF
                '��鲻���Ϲ�ѡҪ���
                'δ�ﵽ��׼סԺ�յģ������ֹ�ѡ�˵�
                If Val(mrsMerge!ѡ�� & "") = 1 Then
                    If Not (Val(mrsMerge!״̬ & "") = 1 Or Val(mrsMerge!״̬ & "") = 2 Or optDate(1).Value Or optDate(1).Enabled = False) Then
                        MsgBox "������ɵĺϲ�·��δ�����׼סԺ�գ�������ǰ��ɣ���ѡ����һ�׶���ǰ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    '���빴ѡ��δ��ѡ
                    If Val(mrsMerge!״̬ & "") = 2 And Not optDate(2).Value Then
                        mrsMerge!ѡ�� = 1
                        mrsMerge.Update
                        blnTmp = True
                    End If
                End If
                mrsMerge.MoveNext
            Loop
            mrsMerge.MoveFirst
            If blnTmp And lblMSG.Caption = "" Then
                '���״̬����ΪNull��δѡ���Ҫֹͣ�ĺϲ�·�����������б���Ҫ��ɵĺϲ�·�����������һ�죩,����ʾ
                If MsgBox("�кϲ�·���ﵽ�˱�׼סԺ�գ��������Զ���ɺϲ�·�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                    Set mrsMerge = rsTmp
                    Exit Sub
                End If
            End If
        End If
    End If
        
    If blnOver Or optResult(2).Value Then
        '�ж��Ƿ���δ��˵Ľ׶�
        If InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";��ת���;") = 0 Then
            If CheckPathIsTurnAduit Then
                str��ת����� = zlDatabase.UserIdentify(Me, "ǰ��׶δ���δ��˵�·����ת��������˺������" & IIf(optResult(2).Value, "�˳�", "���") & "��", glngSys, p�ٴ�·��Ӧ��, "��ת���")
                If str��ת����� = "" Then Exit Sub
            End If
        Else
            str��ת����� = UserInfo.����
        End If
        If CheckPathOutLog Then
            blnOK = frmPathOutLog.ShowMe(Me, mPati.����ID, mPati.��ҳID, 0, mcolSQL, mPP.·��ID, mPP.����·��ID)
            If blnOK = False Then
                i = Val(zlDatabase.GetPara("������д�����ǼǱ�", glngSys, p�ٴ�·��Ӧ��, "0"))
                If i = 1 Then Exit Sub
            End If
        End If
    End If
    
    
    '����ȷ����ȷ������Ϊ�����ã���ִ�����������ã���ֹ���濨���û���ε����
    cmdOK.Enabled = False
    If mlngFun = 1 Then
    '�׶�����ǰ���ýӿ�
        blnTmp = True
        If CreatePlugInOK(p�ٴ�·��Ӧ��) Then
            On Error Resume Next
            blnTmp = gobjPlugIn.PathEvaluateBefore(glngSys, p�ٴ�·��Ӧ��, mPati.����ID, mPati.��ҳID, mPP.����·��ID, mPP.��ǰ�׶�ID)
            '����ӿڲ����ڣ���Ӱ��ԭ���߼�
            If Not blnTmp And Err.Number <> 0 Then blnTmp = True
            Call zlPlugInErrH(Err, "PathEvaluateBefore")
            Err.Clear: On Error GoTo 0
        End If
        If Not blnTmp Then Unload Me: Exit Sub
    End If
    
    Call SaveData(blnOver, str�����, str������, strVariation, str��ת�����)
    
    If mlngFun = 1 Then
    '�׶���������ýӿ�
        If CreatePlugInOK(p�ٴ�·��Ӧ��) Then
            On Error Resume Next
            Call gobjPlugIn.PathEvaluateAfter(glngSys, p�ٴ�·��Ӧ��, mPati.����ID, mPati.��ҳID, mPP.����·��ID, mPP.��ǰ�׶�ID)
            Call zlPlugInErrH(Err, "PathEvaluateAfter")
            Err.Clear: On Error GoTo 0
        End If
    End If
    
    mblnOK = True
    cmdOK.Enabled = True
    Unload Me
End Sub

Private Sub SaveData(ByVal blnOver As Boolean, ByVal str����� As String, ByVal str������ As String, ByVal strVariation As String, ByVal str��ת����� As String)
'����:��������
'����:str�����=�����˳����ڼ����������
'   blnOver=    ���һ������ʱ����·��
'   strVariation=����ԭ��
    Dim strSql As String, str����˵�� As String, lng������� As Long
    Dim strID As String, str���ϵ��� As String, i As Long
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strTotal As String, strThis As String, dateInPath As Date
    Dim strʱ����� As String, str·����ת As String
    Dim rsTmp As Recordset, dateCur As Date
    Dim AddDate As Date
    Dim str�ϲ�·����¼IDs As String   'Ҫ�����ĺϲ�·����¼ID

    If mlngFun = 0 Then
        str���ϵ��� = IIf(optResult(0).Value = True, "1", "0")
        '����Ǻϲ�·�������ϣ�ֱ���˳���
        If str���ϵ��� = "0" And mlngType = 1 Then Exit Sub
        str����˵�� = Trim(txtRemark.Text)
        dateCur = zlDatabase.Currentdate
        If mlngType = 0 Then
            strID = zlDatabase.GetNextId("�����ٴ�·��")
        Else
            strID = zlDatabase.GetNextId("���˺ϲ�·��")
        End If
        If optStart(0).Value Then
            If CheckPathSend(mPati.����ID, mPati.��ҳID) Then
                dateInPath = dateCur
            Else
                dateInPath = GetPatiInDate(mPati)
            End If
            AddDate = dateCur
        Else
            dateInPath = dateCur
        End If
        
        strSql = "Zl_����·������_Insert(" & mPati.����ID & "," & mPati.��ҳID & "," & mPati.����ID & "," & _
                mPP.·��ID & "," & mPP.�汾�� & "," & strID & ",'" & UserInfo.���� & "','" & str����˵�� & "'," & _
                str���ϵ��� & ",To_Date('" & Format(dateInPath, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
                strVariation & "'," & mlngDiagnosisType & "," & mlngDiagnosisSorce & "," & IIf(mlng����ID = 0, "NULL", mlng����ID) & "," & IIf(mlng���ID = 0, "NULL", mlng���ID)
                
        If vsCriterion.Visible = False Then
            colSQL.Add strSql & ",Null," & colSQL.count + 1 & "," & mlngType & "," & mlng��Ҫ·����¼ID & ")", "C" & colSQL.count + 1
        Else
            With vsCriterion
                For i = .FixedRows To .Rows - 1
                    strThis = .TextMatrix(i, mcol("����ָ��")) & "|" & .TextMatrix(i, mcol("���")) & "|" & .TextMatrix(i, mcol("ָ������")) & "||"
                    If LenB(strTotal & strThis) > 4000 Then
                        colSQL.Add strSql & ",'" & strTotal & "'," & colSQL.count + 1 & "," & mlngType & "," & mlng��Ҫ·����¼ID & ")", "C" & colSQL.count + 1
                        strTotal = strThis
                    Else
                        strTotal = strTotal & strThis
                    End If
                Next
                If strTotal <> "" Then
                    colSQL.Add strSql & ",'" & strTotal & "'," & colSQL.count + 1 & "," & mlngType & "," & mlng��Ҫ·����¼ID & ")", "C" & colSQL.count + 1
                Else
                    colSQL.Add strSql & ",Null," & colSQL.count + 1 & "," & mlngType & "," & mlng��Ҫ·����¼ID & ")", "C" & colSQL.count + 1
                End If
            End With
        End If
        If mlngType = 0 And mblnPathSend = False Then
            '�����ʼʱ�䲻�ǵ�ǰʱ�䣬���Զ�ƥ�䲢����֮ǰ��·���׶κ���Ŀ��
            If optStart(0).Value And optResult(0).Value Then
                'ƥ��·����Ŀ
                Call CreatePathItem(dateCur, dateInPath, mPati, mPP, CLng(strID), colSQL)
            End If
        End If
    Else
        str����˵�� = Trim(txtRemark.Text)
        
        If cmdTurn.Enabled And cmdTurn.Visible Then str·����ת = cmdTurn.Tag '·��ID,�汾��
        If str·����ת = "" Then str·����ת = "Null,Null,,Null"  '·��ID,�汾��,��ת�����,�����ʷ��ת
        
        If optDate(0).Value Then
            strʱ����� = "0"
        ElseIf optDate(1).Value Then
            strʱ����� = "1"      '��һ�׶���ǰ������
        ElseIf optDate(3).Value Then
            strʱ����� = "2"     '��һ�׶���ǰ������
        ElseIf optDate(2).Value Then
            strʱ����� = "-1"    '�Ӻ�
        End If
        
        lng������� = 0
        If optResult(1).Value Then
            lng������� = 1
        ElseIf optResult(2).Value Then
            lng������� = 2
        ElseIf optResult(3).Value Then
            lng������� = 3
        End If
        If Not mrsMerge Is Nothing And mlngFun = 1 And mlngState = 1 Then
            If mrsMerge.RecordCount > 0 Then
                mrsMerge.MoveFirst
                Do While Not mrsMerge.EOF
                    '��鲻���Ϲ�ѡҪ���
                    'δ�ﵽ��׼סԺ�յģ������ֹ�ѡ�˵�
                    If Val(mrsMerge!ѡ�� & "") = 1 Then
                        str�ϲ�·����¼IDs = str�ϲ�·����¼IDs & "," & mrsMerge!ID
                    End If
                    mrsMerge.MoveNext
                Loop
                str�ϲ�·����¼IDs = Mid(str�ϲ�·����¼IDs, 2)
            End If
        End If
        
        strSql = "Zl_����·������_Insert(" & mlngState & "," & mPP.����·��ID & "," & mPP.��ǰ�׶�ID & _
            ",To_Date('" & mPP.��ǰ���� & "','YYYY-MM-DD')," & mPP.��ǰ���� & ",'" & _
            str������ & "'," & lng������� & ",'" & str����˵�� & "','" & UserInfo.���� & "','" & str����� & "','" & strVariation & "'," & strʱ����� & "," & Split(str·����ת, ",")(0) & "," & Split(str·����ת, ",")(1)
            
        With vsCriterion
            If .Visible Then    '���Բ�����ָ��
                For i = .FixedRows To .Rows - 1
                    strThis = .TextMatrix(i, mcol("����ָ��")) & "|" & .TextMatrix(i, mcol("���")) & "|" & .TextMatrix(i, mcol("ָ������")) & "||"
                    If LenB(strTotal & strThis) > 4000 Then
                        colSQL.Add strSql & ",'" & strTotal & "'," & colSQL.count + 1 & ",'" & Split(str·����ת, ",")(2) & "'," & Split(str·����ת, ",")(3) & ",'" & str�ϲ�·����¼IDs & "')", "C" & colSQL.count + 1
                        strTotal = strThis
                    Else
                        strTotal = strTotal & strThis
                    End If
                Next
                If strTotal <> "" Then
                    colSQL.Add strSql & ",'" & strTotal & "'," & colSQL.count + 1 & ",'" & Split(str·����ת, ",")(2) & "'," & Split(str·����ת, ",")(3) & ",'" & str�ϲ�·����¼IDs & "')", "C" & colSQL.count + 1
                Else
                    colSQL.Add strSql & ",Null," & colSQL.count + 1 & ",'" & Split(str·����ת, ",")(2) & "'," & Split(str·����ת, ",")(3) & ",'" & str�ϲ�·����¼IDs & "')", "C" & colSQL.count + 1
                End If
            Else
                colSQL.Add strSql & ",Null," & colSQL.count + 1 & ",'" & Split(str·����ת, ",")(2) & "'," & Split(str·����ת, ",")(3) & ",'" & str�ϲ�·����¼IDs & "')", "C" & colSQL.count + 1
            End If
        End With
        If blnOver Then
            strSql = "Zl_����·������_Update(" & mPP.����·��ID & ",'" & str��ת����� & "')"
            colSQL.Add strSql, "C" & colSQL.count + 1
        End If
    End If
    
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        'ִ�г����ǼǱ��SQL
        For i = 1 To mcolSQL.count
            Call zlDatabase.ExecuteProcedure(mcolSQL("C" & i), "�����ǼǱ�")
        Next
        For i = 1 To colSQL.count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "·������")
        Next
    gcnOracle.CommitTrans: blnTrans = False
    '��Ϣ����
    strSql = ""
    For i = 1 To mcolSQL.count
        If InStr(UCase(mcolSQL("C" & i)), "Zl_����·������_INSERT") > 0 Then
            strSql = "do"
            Exit For
        End If
    Next
    
    If strSql <> "" Then
        For i = 1 To colSQL.count
            If InStr(UCase(colSQL("C" & i)), "Zl_����·������_INSERT") > 0 Then
                strSql = "do"
                Exit For
            End If
        Next
    End If
    
    If strSql <> "" Then
        Call ZLHIS_CIS_001(Nothing, mPati.����ID, mPati.��ҳID, mPati.����ID, mPati.����ID)
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsVariation_Click()
    Dim i As Long
    
    With vsVariation
        If .Row >= .FixedRows Then
            .Redraw = flexRDNone
            If mlngFun = 1 Then  '�׶�����
                If .Cell(flexcpData, .Row, col����ѡ��) = 0 Then
                    Set .Cell(flexcpPicture, .Row, col����ѡ��) = imgNature.ListImages("Check").Picture
                    .Cell(flexcpData, .Row, col����ѡ��) = 1
                Else
                    Set .Cell(flexcpPicture, .Row, col����ѡ��) = imgNature.ListImages("UnCheck").Picture
                    .Cell(flexcpData, .Row, col����ѡ��) = 0
                End If
            ElseIf mlngFun = 0 Then '��������
                If .Cell(flexcpData, .Row, col����ѡ��) = 0 Then
                    Set .Cell(flexcpPicture, .Row, col����ѡ��) = imgNature.ListImages("Selected").Picture
                    .Cell(flexcpData, .Row, col����ѡ��) = 1
                    For i = .FixedRows To .Rows - 1
                        If i <> .Row Then
                            If .Cell(flexcpData, i, col����ѡ��) = 1 Then
                                Set .Cell(flexcpPicture, i, col����ѡ��) = imgNature.ListImages("UnSelected").Picture
                                .Cell(flexcpData, i, col����ѡ��) = 0
                            End If
                        End If
                    Next
                Else
                    Set .Cell(flexcpPicture, .Row, col����ѡ��) = imgNature.ListImages("UnSelected").Picture
                    .Cell(flexcpData, .Row, col����ѡ��) = 0
                End If
            End If
            .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub vsVariation_GotFocus()
    If vsVariation.Row < vsVariation.FixedRows And vsVariation.Rows > vsVariation.FixedRows Then vsVariation.Row = vsVariation.FixedRows
End Sub

Private Sub vsVariation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call vsVariation_Click
    End If
End Sub

Private Sub txtVariation_GotFocus()
    Call zlControl.TxtSelAll(txtVariation)
End Sub

Private Sub txtVariation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim i As Long, strtxt As String
        strtxt = "*" & UCase(Trim(txtVariation.Text)) & "*"
        With vsVariation
            For i = .FixedRows To .Rows - 1
                If .RowData(i) Like strtxt Or .TextMatrix(i, col����ԭ��) Like strtxt Or .Cell(flexcpData, i, col����ԭ��) Like strtxt Then
                    .SetFocus
                    .Row = i
                    .TopRow = i
                    Exit Sub
                End If
            Next
        End With
    End If
End Sub

Private Function IsLastDate(Optional ByVal blnEnd As Boolean, Optional ByRef lngMin As Long, Optional ByRef lngMax As Long, Optional ByVal lng·��ID As Long _
                                                                                                                            , Optional ByVal lng�汾�� As Long, Optional ByVal lng��ǰ�׶�ID As Long, Optional ByVal lng��ǰ�׶η�֧ID As Long, Optional ByVal lng��ǰ���� As Long, _
                            Optional ByVal blnBoth As Boolean, Optional ByRef lngState As Long, Optional ByVal lng�ϲ�·����¼ID As Long) As Boolean
'���ܣ��ж��Ƿ��˳�·��
'      blnEnd=false:�жϵ�ǰ�����Ƿ���·�����׶ε����һ�죬��û�к����׶�
'      blnEnd= true:�Ƿ���������˳����ڱ�׼סԺ�շ�Χ�ڶ����˳���
'������blnBoth=�ϲ�·�����ʱ��ͬʱ������һ����Ǵﵽ��׼סԺ�շ�Χ
'     lng��Ҫ·���׶�ID:blnBoth=trueʱ���ϲ�·�������׶�
'���أ�lngMin��lngMax��׼סԺ��
'      lngState :��blnBoth=true  ����0=δ�ﵽ��׼סԺ�գ�1=�ﵽ��׼סԺ�գ���Ϊ�ﵽ���һ�죬2=��׼סԺ�����һ��
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim arrtmp As Variant, lngʵ������ As Long, lng�������� As Long
    Dim blnIsLastDate As Boolean
    
    lngState = 0 'lngStateΪ���ô�ֵ����ʼΪ0��
    
    If lng��ǰ�׶η�֧ID = 0 Then
        strSql = "Select ��׼סԺ�� From �ٴ�·���汾 Where ·��id = [1] And �汾�� = [2]"
    Else
        strSql = "Select ��׼סԺ�� From �ٴ�·����֧ Where ·��id = [1] And �汾�� = [2] And ID=[3]"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ID, lng�汾��, lng��ǰ�׶η�֧ID)
    If Not IsNull(rsTmp!��׼סԺ��) Then
        arrtmp = Split(rsTmp!��׼סԺ��, "-")
        If UBound(arrtmp) > 0 Then
            lngMin = arrtmp(0)
            lngMax = arrtmp(1)
        Else
            lngMin = 1  'С�ڵ���n��
            lngMax = arrtmp(0)
        End If

        If blnEnd Or blnBoth Then
            lng�������� = GetMustDay(mPP.����·��ID, lng��ǰ����, , lng�ϲ�·����¼ID)
            If lng�������� > lngMax Then
                blnIsLastDate = True

            Else
                blnIsLastDate = Between(lng��������, lngMin, lngMax)
            End If
            If blnIsLastDate And blnBoth Then
                lngState = 1
            End If
        End If
        If blnIsLastDate Then IsLastDate = blnIsLastDate
        If Not blnEnd Or blnBoth Then
            lngʵ������ = GetMustDay(mPP.����·��ID, lng��ǰ����, True, lng�ϲ�·����¼ID)
            If lngʵ������ >= lngMax Then
                blnIsLastDate = GetNextPhase(lng��ǰ�׶�ID, lng��ǰ�׶η�֧ID) = 0
                If blnIsLastDate And blnBoth Then
                    lngState = 2
                End If
            End If
        End If
    End If
    If blnIsLastDate Then IsLastDate = blnIsLastDate
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
