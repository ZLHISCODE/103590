VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl ElementEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   11940
   ToolboxBitmap   =   "ElementEdit.ctx":0000
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3690
      Left            =   9570
      ScaleHeight     =   3690
      ScaleWidth      =   2235
      TabIndex        =   15
      Top             =   15
      Visible         =   0   'False
      Width           =   2235
      Begin MSMask.MaskEdBox txtTime 
         Height          =   330
         Left            =   0
         TabIndex        =   24
         Top             =   1815
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.ListBox lstTime 
         Appearance      =   0  'Flat
         Height          =   1290
         Left            =   0
         TabIndex        =   17
         Top             =   2400
         Width           =   2235
      End
      Begin VB.PictureBox picClock 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1770
         Left            =   0
         ScaleHeight     =   1740
         ScaleWidth      =   2205
         TabIndex        =   16
         Top             =   0
         Width           =   2235
         Begin VB.Line linHand 
            BorderWidth     =   2
            X1              =   240
            X2              =   960
            Y1              =   735
            Y2              =   585
         End
         Begin VB.Shape shpCenter 
            Height          =   90
            Left            =   675
            Shape           =   3  'Circle
            Top             =   720
            Width           =   90
         End
         Begin VB.Shape shpDot 
            FillColor       =   &H00FFFFFF&
            Height          =   105
            Index           =   0
            Left            =   690
            Shape           =   3  'Circle
            Top             =   180
            Width           =   135
         End
      End
      Begin VB.Label lblTimeType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʽ����"
         Height          =   180
         Left            =   0
         TabIndex        =   19
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label lblAmOrPm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   1725
         TabIndex        =   18
         Top             =   1890
         Width           =   360
      End
   End
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3690
      Left            =   5460
      ScaleHeight     =   3690
      ScaleWidth      =   4035
      TabIndex        =   11
      Top             =   15
      Visible         =   0   'False
      Width           =   4035
      Begin VB.ListBox lstDate 
         Appearance      =   0  'Flat
         Height          =   1290
         Left            =   0
         TabIndex        =   12
         Top             =   2400
         Width           =   4035
      End
      Begin MSComCtl2.MonthView mvwDate 
         Height          =   2160
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3810
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   59834369
         TitleForeColor  =   16711680
         CurrentDate     =   38549
         MaxDate         =   401769
         MinDate         =   367
      End
      Begin VB.Label lblDateType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʽ����"
         Height          =   180
         Left            =   0
         TabIndex        =   14
         Top             =   2190
         Width           =   720
      End
   End
   Begin VB.TextBox txt�ı� 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   210
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox txt����1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   -15
      TabIndex        =   5
      Text            =   "99999"
      Top             =   225
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F5BE9E&
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   30
      MousePointer    =   5  'Size
      ScaleHeight     =   105
      ScaleWidth      =   5325
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   5325
      Begin VB.Image imgTitle 
         Height          =   45
         Left            =   1350
         MousePointer    =   5  'Size
         Picture         =   "ElementEdit.ctx":0312
         Top             =   30
         Width           =   2250
      End
   End
   Begin VB.TextBox txt����2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Text            =   "99999"
      Top             =   225
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   5415
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3360
      Width           =   5415
      Begin VB.Image imgResize 
         Height          =   270
         Left            =   5175
         MousePointer    =   8  'Size NW SE
         Picture         =   "ElementEdit.ctx":0394
         Top             =   0
         Width           =   225
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "ESC ȡ���˳����س�:�����޸ġ�"
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   60
         Width           =   3810
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg��ѡ 
      Height          =   555
      Left            =   2865
      TabIndex        =   4
      Top             =   225
      Visible         =   0   'False
      Width           =   780
      _cx             =   1376
      _cy             =   979
      Appearance      =   0
      BorderStyle     =   0
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
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
      FormatString    =   $"ElementEdit.ctx":0736
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComCtl2.UpDown ud���� 
      Height          =   300
      Left            =   645
      TabIndex        =   6
      Top             =   165
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      OrigLeft        =   1395
      OrigTop         =   2220
      OrigRight       =   1635
      OrigBottom      =   2520
      Enabled         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg��ѡ 
      Height          =   570
      Left            =   3705
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   690
      _cx             =   1217
      _cy             =   1005
      Appearance      =   0
      BorderStyle     =   0
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16777215
      ForeColorFixed  =   0
      BackColorSel    =   16761024
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
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
      FormatString    =   $"ElementEdit.ctx":0773
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   1395
      Left            =   60
      TabIndex        =   21
      Top             =   1350
      Width           =   2595
      _cx             =   4577
      _cy             =   2461
      Appearance      =   2
      BorderStyle     =   0
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ElementEdit.ctx":07B0
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
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
   Begin MSComctlLib.Toolbar cbrThis 
      Height          =   330
      Left            =   1905
      TabIndex        =   22
      Top             =   945
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ButtonWidth     =   1349
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ils16"
      HotImageList    =   "ils16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "����"
            Object.ToolTipText     =   "����"
            Object.Tag             =   "����"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ѡ��"
            Key             =   "ѡ��"
            Object.ToolTipText     =   "ѡ��"
            Object.Tag             =   "ѡ��"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ȡ��"
            Key             =   "ȡ��"
            Object.ToolTipText     =   "ȡ��"
            Object.Tag             =   "ȡ��"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   4215
      Top             =   1245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ElementEdit.ctx":0885
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ElementEdit.ctx":70E7
            Key             =   "find"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ElementEdit.ctx":7481
            Key             =   "cancel"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFind 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   105
      TabIndex        =   20
      Top             =   975
      Width           =   1755
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   220
      Left            =   3240
      TabIndex        =   23
      Top             =   2655
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Shape shpBorderOut 
      BorderColor     =   &H00E09060&
      Height          =   255
      Left            =   2490
      Top             =   210
      Width           =   270
   End
   Begin VB.Label lblDot 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   300
      Left            =   675
      TabIndex        =   10
      Top             =   202
      Width           =   105
   End
   Begin VB.Shape shpBorder1 
      BorderColor     =   &H00E09060&
      Height          =   255
      Left            =   1785
      Top             =   210
      Width           =   270
   End
   Begin VB.Shape shpBorder2 
      BorderColor     =   &H00E09060&
      Height          =   255
      Left            =   2160
      Top             =   210
      Width           =   270
   End
   Begin VB.Label lbl��λ 
      BackStyle       =   0  'Transparent
      Caption         =   "��λ"
      Height          =   210
      Left            =   1320
      TabIndex        =   9
      Top             =   247
      Width           =   480
   End
   Begin VB.Image imgOpt2 
      Height          =   195
      Left            =   5205
      Picture         =   "ElementEdit.ctx":DCE3
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgOpt1 
      Height          =   195
      Left            =   5205
      Picture         =   "ElementEdit.ctx":DF69
      Top             =   210
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "ElementEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const conPI As Double = 3.14159265358979
Public Event pOk()              '��������
Public Event pChange()          '��ѡ�����ݸı�
Public Event pCancel()          'ȡ���޸�
Public Event TitleMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event TitleMouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event TitleMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public MoveTag As String, MoveOldX As Single, MoveOldY As Single
Public Element As cTabElement
Private mPressKeyAscii As Integer
Private lngX As Long, lngY As Long
Private mblnHandMove As Boolean, mblnHandTime As Boolean, mbEt As Byte
Public Property Get hWnd()
    hWnd = UserControl.hWnd
End Property
'################################################################################################################
'## ���ܣ�  ��ʾ����Ҫ�ر༭��
'##
'## ������  Ele         :���༭������Ҫ��
'################################################################################################################
Public Sub SetElement(ByRef Ele As cTabElement, ByVal KeyAscii As Integer, Optional ByVal bEditType As Byte = 0)
'���ܣ���ʾ�ؼ���ǰ��ֵҪ�أ��ÿؼ���ı�Ҫ��,���������Ҫ��.�ڿ��ı���ʽ����
Dim i As Long, j As Long, T As Variant, strTmp As String, dtInit As Date
    mPressKeyAscii = KeyAscii: MoveTag = "": MoveOldX = 0: MoveOldY = 0: mbEt = bEditType
    Set Element = Ele
    With Element
        If .�滻�� = 2 Then         '�ֵ���
            txtFind.Text = Chr(KeyAscii)
            vgdList.Clear: vgdList.Rows = 2
            If txtFind.Text <> "" Then DoFind
        Else
            Select Case .Ҫ�ر�ʾ       '0-�ı�,1-����,2-��ѡ,3-��ѡ
            Case 0
                Select Case .Ҫ������
                    Case 2                      '������
                        Dim strMinTime As String, strMaxTime As String, strMinDate As String, strMaxDate As String
                        T = Split(.Ҫ��ֵ��, ";")
                        If T(0) = 0 Then
                            strMinDate = Format("1901-01-01", "yyyy-MM-dd"): strMinTime = Format("00:00:00", "HH:mm:ss")
                        Else
                            strMinDate = Format(T(0), "yyyy-MM-dd"): strMinTime = Format(T(0), "HH:mm:ss")
                        End If
                        
                        If T(1) = 0 Then
                            strMaxDate = Format("3000-01-01", "yyyy-MM-dd"): strMaxTime = Format("23:59:59", "HH:mm:ss")
                        Else
                            strMaxDate = Format(T(1), "yyyy-MM-dd"): strMaxTime = Format(T(1), "HH:mm:ss")
                        End If
                        
                        If .������̬ = 0 Then '������ʽ
                            On Error Resume Next
                            mvwDate.MinDate = Format("1901-01-01", "yyyy-MM-dd"): mvwDate.MaxDate = Format("3000-01-01", "yyyy-MM-dd")
                            Err.Clear
                            mvwDate.MinDate = Format(strMinDate, "yyyy-MM-dd"): mvwDate.MaxDate = Format(strMaxDate, "yyyy-MM-dd")
                            txtTime.Tag = Format(strMinTime, "HH:mm:ss") & "|" & Format(strMaxTime, "HH:mm:ss")
'                            dtTime.MinDate = Format(strMinTime, "HH:mm:ss"): dtTime.MaxDate = Format(strMaxTime, "HH:mm:ss")
                            lstDate.Tag = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\ElementEdit\" & .Ҫ������ & .Ҫ�ر�ʾ & .�滻��, "\DateType", 0)
                            lstTime.Tag = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\ElementEdit\" & .Ҫ������ & .Ҫ�ر�ʾ & .�滻��, "\TimeType", 0)
                            If .Ҫ�س��� >= 10 Then '���ںͳ�����
                                If .�����ı� = "" Then
                                    dtInit = Now
                                Else
                                    dtInit = Format(.�����ı�, "yyyy-MM-dd")
                                End If
                                If dtInit < mvwDate.MinDate Then dtInit = mvwDate.MinDate
                                If dtInit > mvwDate.MaxDate Then dtInit = mvwDate.MaxDate
                                mvwDate.Value = dtInit
                                Call mvwDate_SelChange(mvwDate.MinDate, mvwDate.MaxDate, False)
                            End If
                            
                            Call DrawWatch
                            If .Ҫ�س��� <> 10 Then 'ʱ��ͳ�����
                                If .�����ı� <> "" And Format(.�����ı�, "hh:mm:ss") <> CDate("00:00:00") Then
                                    txtTime.Text = Format(.�����ı�, "HH:mm:ss")
                                Else
                                    txtTime.Text = "__:__:__"
                                End If
                            End If
                            Err.Clear
                        Else            'չ����
                            txtDate.Tag = strMinDate & "|" & strMaxDate & "|" & strMinTime & "|" & strMaxTime
                            Select Case .Ҫ�س���
                                Case 8
                                    txtDate.Format = "HH:mm:ss"
                                    txtDate.Mask = "##:##:##"
                                    txtDate.Text = Format(IIf(Trim(.�����ı�) = "", Now, Trim(.�����ı�)), "HH:mm:ss")
                                    If CDate(txtDate) < CDate(strMinTime) Then txtDate.Text = strMinTime
                                    If CDate(txtDate) > CDate(strMaxTime) Then txtDate.Text = strMaxTime
                                Case 10
                                    txtDate.Format = "yyyy-MM-dd"
                                    txtDate.Mask = "####-##-##"
                                    txtDate.Text = Format(IIf(Trim(.�����ı�) = "", Now, Trim(.�����ı�)), "yyyy-MM-dd")
                                    If CDate(txtDate) < CDate(strMinDate) Then txtDate.Text = strMinDate
                                    If CDate(txtDate) > CDate(strMaxDate) Then txtDate.Text = strMaxDate
                                Case 19
                                    txtDate.Format = "yyyy-MM-dd HH:mm:ss"
                                    txtDate.Mask = "####-##-## ##:##:##"
                                    txtDate.Text = Format(IIf(Trim(.�����ı�) = "", Now, Trim(.�����ı�)), "yyyy-MM-dd HH:mm:ss")
                                    If CDate(txtDate) < CDate(strMinDate & " " & strMinTime) Then txtDate.Text = strMinDate & " " & strMinTime
                                    If CDate(txtDate) > CDate(strMaxDate & " " & strMaxTime) Then txtDate.Text = strMaxDate & " " & strMaxTime
                            End Select
                            txtDate.MaxLength = .Ҫ�س���
                        End If
                    Case 3                      '�߼���
                        strTmp = .�����ı�:   T = Split(.Ҫ��ֵ��, ";"):      strTmp = IIf(strTmp = "", T(1), strTmp):      .�����ı� = strTmp
                        vfg��ѡ.RowHeightMax = 240:     vfg��ѡ.Cols = 2:       vfg��ѡ.ColWidth(0) = 250
                        vfg��ѡ.ColWidth(1) = IIf(ScaleWidth > 250, ScaleWidth - 250, 250):             vfg��ѡ.Rows = UBound(T) + 1
                        For i = 0 To UBound(T)
                            vfg��ѡ.Cell(flexcpText, i, 1) = T(i)
                            vfg��ѡ.Cell(flexcpPicture, i, 0) = IIf(T(i) = strTmp, imgOpt2.Picture, imgOpt1.Picture)
                        Next i
                    Case Else
                        txt�ı�.MaxLength = .Ҫ�س���:  txt�ı� = .�����ı� & Chr(mPressKeyAscii)
                        txt�ı�.SelStart = 0: txt�ı�.SelLength = Len(.�����ı�): txt�ı�.Visible = True
                End Select
            Case 1
                T = Split(.Ҫ��ֵ��, ";")    '��ʽ:  0;100000
                If UBound(T) < 1 Then
                    ud����.Min = 0:             ud����.Max = 999999999
                Else
                    ud����.Min = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(1)), CLng(T(0)))
                    ud����.Max = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(0)), CLng(T(1)))
                End If
                txt����1.Tag = "��ֵ...":       i = InStr(1, .�����ı�, ".")
                If i > 0 Then
                    txt����1 = Mid(.�����ı�, 1, i - 1):    txt����1.Visible = True
                    txt����1.SelStart = 0: txt����1.SelLength = Len(txt����1)
                    txt����2 = Mid(.�����ı�, i + 1)
                Else
                    txt����1 = .�����ı�:                   txt����2 = ""
                End If
                txt����1.Tag = "":                          txt����1.MaxLength = .Ҫ�س���
                lbl��λ = .Ҫ�ص�λ
                If Trim(.Ҫ�ص�λ) <> "" Then
                    lbl��λ.Visible = True
                Else
                    lbl��λ.Visible = False
                End If
                If .Ҫ��С�� > 0 Then
                    txt����2.MaxLength = .Ҫ��С��:         txt����2.Visible = True:        lblDot.Visible = True
                Else
                    txt����2.Visible = False:   lblDot.Visible = False
                End If
            Case 2
                vfg��ѡ.Clear:      vfg��ѡ.FocusRect = flexFocusNone:      vfg��ѡ.Editable = flexEDKbdMouse:      T = Split(.Ҫ��ֵ��, ";")
                If .������̬ = 0 Then
                    strTmp = "��" & .�����ı� & "��"
                Else 'չ����ʽ   '���
                    For i = 1 To UBound(Split(.�����ı�, "��"))
                        strTmp = strTmp & "��" & Split(Split(.�����ı�, "��")(i), "��")(0)
                    Next
                    strTmp = strTmp & "��"
                End If
                If .������̬ = 0 Then
                    vfg��ѡ.RowHeightMax = 240:     vfg��ѡ.Cols = 2:       vfg��ѡ.ColWidth(0) = 250
                    vfg��ѡ.ColWidth(1) = IIf(ScaleWidth > 250, ScaleWidth - 250, 250):             vfg��ѡ.Rows = UBound(T) + 1
                    For i = 0 To UBound(T)
                        vfg��ѡ.Cell(flexcpText, i, 1) = T(i)
                        vfg��ѡ.Cell(flexcpPicture, i, 0) = IIf(InStr(strTmp, "��" & T(i) & "��") > 0, imgOpt2.Picture, imgOpt1.Picture)
                    Next i
                Else
                    vfg��ѡ.RowHeightMax = 0:       vfg��ѡ.Rows = 1:               vfg��ѡ.Cols = (UBound(T) + 1) * 2
                    For i = 0 To UBound(T) * 2 Step 2 'ÿ���һ��Ϊͼ��
                        vfg��ѡ.Cell(flexcpText, 0, i + 1) = T(i / 2)
                        vfg��ѡ.Cell(flexcpPicture, 0, i) = IIf(InStr(strTmp, "��" & T(i / 2) & "��") > 0, imgOpt2.Picture, imgOpt1.Picture)
                    Next
                End If
            Case 3
                vfg��ѡ.Clear:      vfg��ѡ.Editable = flexEDKbdMouse:      T = Split(.Ҫ��ֵ��, ";")
                If .������̬ = 0 Then
                    strTmp = "��" & .�����ı� & "��"
                Else 'չ����ʽ
                    For i = 1 To UBound(Split(.�����ı�, "��"))
                        strTmp = strTmp & "��" & Split(Split(.�����ı�, "��")(i), "��")(0)
                    Next
                    strTmp = strTmp & "��"
                End If
                If .������̬ = 0 Then
                    vfg��ѡ.RowHeightMax = 240:         vfg��ѡ.Cols = 1:       vfg��ѡ.Rows = UBound(T) + 1:       vfg��ѡ.ColWidth(0) = 240
                    For i = 0 To UBound(T)
                        vfg��ѡ.Cell(flexcpText, i, 0) = T(i)
                        vfg��ѡ.Cell(flexcpChecked, i, 0) = IIf(InStr(1, strTmp, "��" & vfg��ѡ.Cell(flexcpText, i, 0) & "��") > 0, flexChecked, flexUnchecked)
                    Next
                Else
                    vfg��ѡ.RowHeightMax = 0:           vfg��ѡ.Rows = 1:                   vfg��ѡ.Cols = UBound(T) + 1
                    For i = 0 To UBound(T)
                        vfg��ѡ.Cell(flexcpText, 0, i) = T(i)
                        vfg��ѡ.Cell(flexcpChecked, 0, i) = IIf(InStr(1, strTmp, "��" & vfg��ѡ.Cell(flexcpText, 0, i) & "��") > 0, flexChecked, flexUnchecked)
                    Next
                End If
            End Select
        End If

        If .������̬ = 0 Then
            Width = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\ElementEdit\" & .Ҫ������ & .Ҫ�ر�ʾ & .�滻��, "\MainWidth", 2500)
            Height = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\ElementEdit\" & .Ҫ������ & .Ҫ�ر�ʾ & .�滻��, "\MainHeight", 3870)
        End If
    End With
    
    Call UserControl_Resize
End Sub

Private Sub cbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        Call DoFind
    Case 2
        vgdList_DblClick
    Case 3
        RaiseEvent pOk
    End Select
End Sub

Private Sub dtTime_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub imgResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgResize.Tag = "Down"
    lngX = x: lngY = y
End Sub

Private Sub imgResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If imgResize.Tag = "Down" Then
        If Width + x - lngX >= 1000 And Width + x - lngX <= 12000 Then
            Width = Width + x - lngX
        End If
        If Height + y - lngY >= 1000 And Height + y - lngY <= 9000 Then
            Height = Height + y - lngY
        End If
        DoEvents
    End If
End Sub

Private Sub imgResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgResize.Tag = ""
'    Call SetCtlFocus
End Sub
Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent TitleMouseDown(Button, Shift, UserControl.Extender.Left, UserControl.Extender.Top)
End Sub
Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent TitleMouseDown(Button, Shift, UserControl.Extender.Left, UserControl.Extender.Top)
End Sub

Private Sub lstDate_DblClick()
    If Element Is Nothing Then Exit Sub
    With Element
        Select Case .Ҫ�س���
            Case 19
                .�����ı� = lstDate.Text & " " & lstTime.Text
                RaiseEvent pOk
        End Select
    End With
End Sub

Private Sub lstDate_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Element Is Nothing Then Exit Sub
    
    If lstDate.ListIndex >= 0 Then
        lstDate.Tag = Val(lstDate.ListIndex)
    End If
    
    With Element
        Select Case .Ҫ�س���
            Case 10
                .�����ı� = lstDate.Text
                RaiseEvent pOk
            Case 19
                .�����ı� = lstDate.Text & " " & lstTime.Text
        End Select
    End With
End Sub

Private Sub lstTime_DblClick()
    If Element Is Nothing Then Exit Sub
    With Element
        Select Case .Ҫ�س���
            Case 19
                .�����ı� = lstDate.Text & " " & lstTime.Text
                RaiseEvent pOk
        End Select
    End With
End Sub

Private Sub lstTime_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Element Is Nothing Then Exit Sub
    
    If lstTime.ListIndex >= 0 Then
        lstTime.Tag = Val(lstTime.ListIndex)
    End If
    With Element
        Select Case .Ҫ�س���
            Case 8
                .�����ı� = lstTime.Text
                RaiseEvent pOk
            Case 19
                .�����ı� = lstDate.Text & " " & lstTime.Text
        End Select
    End With
End Sub

Private Sub picClock_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim xLine As Long, yLine As Long
    Dim dblSin As Double, dblValue As Double
    If mblnHandMove = False Then Exit Sub
    xLine = x - (shpCenter.Left + shpCenter.Width / 2)
    yLine = y - (shpCenter.Top + shpCenter.Height / 2)
    
    If xLine = 0 And yLine = 0 Then Exit Sub
    
    dblSin = yLine / Sqr(xLine ^ 2 + yLine ^ 2)
    If dblSin = 1 Then
        dblValue = 6
    ElseIf dblSin = -1 Then
        dblValue = 0
    Else
        If Sgn(xLine) >= 0 Then
            dblValue = Round(Atn(dblSin / Sqr(-dblSin * dblSin + 1)) / conPI * 180 / 30, 1) + 3
        Else
            dblValue = 9 - Round(Atn(dblSin / Sqr(-dblSin * dblSin + 1)) / conPI * 180 / 30, 1)
        End If
    End If
    If lblAmOrPm.Caption = "����" And dblValue < 12 Then dblValue = dblValue + 12
    Call SetTimer(dblValue)
End Sub

Private Sub picStatus_Resize()
imgResize.Move picStatus.ScaleWidth - imgResize.Width, 0
lblInfo.Move 80, 0, picStatus.Width - imgResize.Width
End Sub
Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent TitleMouseDown(Button, Shift, UserControl.Extender.Left, UserControl.Extender.Top)
End Sub
'
'Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent TitleMouseMove(Button, Shift, UserControl.Extender.Left, UserControl.Extender.Top)
'End Sub

'Private Sub picTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent TitleMouseUp(Button, Shift, UserControl.Extender.Left, UserControl.Extender.Top)
'End Sub

Private Sub picTitle_Resize()
    imgTitle.Move (picTitle.ScaleWidth - imgTitle.Width) / 2, 30
End Sub
Private Sub txtDate_Validate(Cancel As Boolean)
'txtdate.tag=��С����|�������|��Сʱ��|���ʱ��
    If Element Is Nothing Then Exit Sub
    With Element
        txtDate.Text = Trim(txtDate.Text)
        If IsDate(txtDate.Text) Then
            Select Case .Ҫ�س���
                Case 8
                    If CDate(txtDate.Text) > CDate(Split(txtDate.Tag, "|")(3)) Then txtDate.Text = Split(txtDate.Tag, "|")(3)
                    If CDate(txtDate.Text) < CDate(Split(txtDate.Tag, "|")(2)) Then txtDate.Text = Split(txtDate.Tag, "|")(2)
                Case 10
                    If CDate(txtDate.Text) > CDate(Split(txtDate.Tag, "|")(1)) Then txtDate.Text = Split(txtDate.Tag, "|")(1)
                    If CDate(txtDate.Text) < CDate(Split(txtDate.Tag, "|")(0)) Then txtDate.Text = Split(txtDate.Tag, "|")(0)
                Case 19
                    If CDate(txtDate.Text) > CDate(Split(txtDate.Tag, "|")(1) & " " & Split(txtDate.Tag, "|")(3)) Then txtDate.Text = Split(txtDate.Tag, "|")(1) & " " & Split(txtDate.Tag, "|")(3)
                    If CDate(txtDate.Text) < CDate(Split(txtDate.Tag, "|")(0) & " " & Split(txtDate.Tag, "|")(2)) Then txtDate.Text = Split(txtDate.Tag, "|")(0) & " " & Split(txtDate.Tag, "|")(2)
            End Select
            .�����ı� = Format(txtDate.Text, txtDate.Format)
        Else
            .�����ı� = Format(Now, txtDate.Format)
        End If
    End With
    UserControl_KeyPress vbKeyReturn
End Sub

Private Sub txtDate_ValidationError(InvalidText As String, StartPosition As Integer)
    StartPosition = 0
End Sub

Private Sub txtFind_GotFocus()
    If Element Is Nothing Then Exit Sub
    txtFind.SelStart = 0: txtFind.SelLength = 1000
    lblInfo = Element.Ҫ������ & " ����ϣ��������Ŀ�ı���/����/����"
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call DoFind
        Exit Sub
    End If
End Sub

Private Sub txtTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call UserControl_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub UserControl_EnterFocus()
    SetCtlFocus
End Sub

Private Sub UserControl_Hide()
    If Not Element Is Nothing Then
        With Element
            If .������̬ = 0 Then
                SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\ElementEdit\" & .Ҫ������ & .Ҫ�ر�ʾ & .�滻��, "\MainWidth", Width
                SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\ElementEdit\" & .Ҫ������ & .Ҫ�ر�ʾ & .�滻��, "\MainHeight", Height
                If .Ҫ������ = 2 Then  '��¼�����ַ���ʽ ��¼ʱ���ַ���ʽ
                    If lstDate.ListIndex >= 0 Then
                        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\ElementEdit\" & .Ҫ������ & .Ҫ�ر�ʾ & .�滻��, "\DateType", lstDate.ListIndex
                    End If
                    If lstTime.ListIndex >= 0 Then
                        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\ElementEdit\" & .Ҫ������ & .Ҫ�ر�ʾ & .�滻��, "\TimeType", lstTime.ListIndex
                    End If
                End If
            Else
                If .Ҫ�ر�ʾ = 3 Then '��ѡ��չ����ʽ,ѡ���Ǽ�ʱ��Ч��
                    UserControl_KeyPress vbKeyReturn
                End If
            End If
        End With
    End If
    Set Element = Nothing
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Element Is Nothing Then Exit Sub
    If KeyCode = vbKeyDelete Then
        If Element.Ҫ������ = 2 Then '������
            Element.�����ı� = ""
            RaiseEvent pOk
        End If
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Element.Ҫ������ = 0 Then
            '��ֵ��
            Dim T As Variant, dblMax As Double, dblMin As Double
            T = Split(Element.Ҫ��ֵ��, ";")    '��ʽ:  0;100000
            If UBound(T) < 1 Then
                dblMin = 0#
                dblMax = 0#
            Else
                dblMin = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(1)), CLng(T(0)))
                dblMax = IIf(CLng(T(0)) > CLng(T(1)), CLng(T(0)), CLng(T(1)))
            End If
            If Element.Ҫ�ر�ʾ = 0 Then
                '�ı���ʾ
                If Trim(txt�ı�) = "" Then
                    Element.�����ı� = ""
                ElseIf Element.Ҫ��ֵ�� <> ";" And Element.Ҫ��ֵ�� <> "0;0" And Element.Ҫ��ֵ�� <> "" Then
                    If Val(txt�ı�) > dblMax Then
                        txt�ı� = dblMax
                    ElseIf Val(txt�ı�) < dblMin Then
                        txt�ı� = dblMin
                    End If
                    Element.�����ı� = IIf(Element.Ҫ��С�� > 0, Format(txt�ı�, "0." & String(Element.Ҫ��С��, "0")), txt�ı�)
                Else
                    Element.�����ı� = IIf(Element.Ҫ��С�� > 0, Format(txt�ı�, "0." & String(Element.Ҫ��С��, "0")), txt�ı�)
                End If
            ElseIf Element.Ҫ�ر�ʾ = 1 Then
                '���±�ʾ
                If Trim(Element.�����ı�) <> "" And Element.Ҫ��ֵ�� <> ";" And Element.Ҫ��ֵ�� <> "0;0" Then
                    If Val(Element.�����ı�) > dblMax Then
                        Element.�����ı� = dblMax
                    ElseIf Val(Element.�����ı�) < dblMin Then
                        Element.�����ı� = dblMin
                    End If
                Else
                    Element.�����ı� = IIf(Element.Ҫ��С�� > 0, Format(Element.�����ı�, "0." & String(Element.Ҫ��С��, "0")), Element.�����ı�)
                End If
            End If
        ElseIf Element.Ҫ������ = 2 Then '����/ʱ����Ҫ��
            If lstDate.Visible And lstTime.Visible = False Then '������
                Element.�����ı� = lstDate.Text
            ElseIf lstDate.Visible = False And lstTime.Visible Then 'ʱ����
                If Not IsDate(txtTime.Text) Then
                    txtTime.Text = Format(Val(Mid(txtTime, 1, 2)) & ":" & Val(Mid(txtTime, 4, 2)) & ":" & Val(Mid(txtTime, 7, 2)), "00:00:00")
                End If
                Element.�����ı� = lstTime.Text
            ElseIf lstDate.Visible And lstTime.Visible Then '����ʱ����
                If Not IsDate(txtTime.Text) Then
                    txtTime.Text = Format(Val(Mid(txtTime, 1, 2)) & ":" & Val(Mid(txtTime, 4, 2)) & ":" & Val(Mid(txtTime, 7, 2)), "00:00:00")
                End If
                Element.�����ı� = lstDate.Text & " " & lstTime.Text
            End If
        End If
        If Element.�滻�� <> 2 Then '�ֵ���Ŀ��ѡ�д���
            RaiseEvent pOk
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        RaiseEvent pCancel
    ElseIf KeyAscii = vbKeySpace Then
        If vfg��ѡ.Visible Then vfg��ѡ_KeyDown KeyAscii, 0
    ElseIf KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Then
        If vfg��ѡ.Visible Then vfg��ѡ_KeyDown KeyAscii, 0
    ElseIf InStr("1234567890", Chr(KeyAscii)) > 0 And (Element.Ҫ�ر�ʾ = 2 Or Element.Ҫ�ر�ʾ = 3) Then
        mPressKeyAscii = KeyAscii
        Call UserControl_Show
    End If
End Sub
Private Sub UserControl_Resize()
Dim lX As Long, lY As Long
    On Error Resume Next
    lX = Screen.TwipsPerPixelX:     lY = Screen.TwipsPerPixelY
    txt����1.Visible = False:       txt����2.Visible = False:       lblDot.Visible = False:         lbl��λ.Visible = False:    shpBorder1.Visible = False
    shpBorder2.Visible = False:     txt�ı�.Visible = False:        ud����.Visible = False:         vfg��ѡ.Visible = False:    vfg��ѡ.Visible = False
    picTime.Visible = False:        picDate.Visible = False:        cbrThis.Visible = False:        txtFind.Visible = False:    vgdList.Visible = False
    shpBorder1.BorderWidth = 1:     shpBorder2.BorderWidth = 1:     txtDate.Visible = False
    
    picTitle.Move 60, 60, ScaleWidth - 120
    picStatus.Move lX, ScaleHeight - picStatus.Height - lY, ScaleWidth - lX * 2
    shpBorderOut.Move 0, 0, Width, Height: lblInfo = "ESC ȡ���˳����س�:�����޸ġ�"
    If Element Is Nothing Then Exit Sub

    With Element
        If .������̬ = 1 Then
            picTitle.Visible = False: picStatus.Visible = False: shpBorderOut.Visible = False
        Else
            picTitle.Visible = True: picStatus.Visible = True: shpBorderOut.Visible = True
        End If
        If .�滻�� = 2 Then
            Dim ltxtWidth As Long, lvgdListHeight As Long
            lblInfo = .Ҫ������ & " ����ϣ��������Ŀ�ı���/����/�����س�"
            cbrThis.Move ScaleWidth - 80 - cbrThis.Width, picTitle.Height + 120
            ltxtWidth = Width - cbrThis.Width - 240: If ltxtWidth < 100 Then ltxtWidth = 100
            txtFind.Move 80, cbrThis.Top, ltxtWidth, cbrThis.Height
            vgdList.Move 80, txtFind.Top + txtFind.Height + lX * 4, ScaleWidth - 160, IIf(ScaleHeight - picStatus.Height - picTitle.Height - txtFind.Height - 250 < 0, 0, ScaleHeight - picStatus.Height - picTitle.Height - txtFind.Height - 250)
            shpBorder1.Move vgdList.Left - lX, vgdList.Top - lY, vgdList.Width + lX * 3, vgdList.Height + lY * 2
            shpBorder2.Move txtFind.Left - lX, txtFind.Top - lY, txtFind.Width + lX * 2, txtFind.Height + lY * 2
            cbrThis.Visible = True:     txtFind.Visible = True:     vgdList.Visible = True: shpBorder1.Visible = True: shpBorder2.Visible = True
            txtFind.ZOrder 0
            If txtFind.Visible And txtFind.Enabled Then txtFind.SetFocus
        Else
            Select Case .Ҫ�ر�ʾ
            Case 0
                Select Case .Ҫ������
                    Case 2      '������
                        If .������̬ = 1 Then
                            If .Ҫ�س��� < 19 Then txtDate.Width = 1000 Else txtDate.Width = 1800
                            If ScaleWidth <= txtDate.Width + lX * 2 Then txtDate.Left = 0 Else txtDate.Left = (ScaleWidth - txtDate.Width - lX * 2) / 2
                            If ScaleHeight <= txtDate.Height + lY * 2 Then txtDate.Top = 0 Else txtDate.Top = (ScaleHeight - txtDate.Height - lY * 2) / 2
                            shpBorder1.Move txtDate.Left - lX, txtDate.Top - lY, txtDate.Width + lX * 2, txtDate.Height + lY * 2
                            shpBorder1.Visible = True: txtDate.Visible = True
                            If txtDate.Visible And txtDate.Enabled Then txtDate.SetFocus
                            txtDate.SelStart = 0: txtDate.SelLength = Len(txtDate)
                        Else
                            imgResize.Tag = ""
                            Select Case .Ҫ�س���
                                Case 8
                                    picTime.Move 80, picTitle + 200, picClock.Width + 80 + lstTime.Width, picClock.Height
                                    txtTime.Move picClock.Width + 80, picClock.Top: lblAmOrPm.Move txtTime.Left + txtTime.Width + 100, txtTime.Top
                                    lblTimeType.Move txtTime.Left, txtTime.Height
                                    lstTime.Move txtTime.Left, lblTimeType.Top + lblTimeType.Height, lstTime.Width, picClock.Height - txtTime.Height - lblTimeType.Height
                                    Width = picTime.Width + 160: Height = picStatus.Height + picTitle.Height + picTime.Height + 200
                                    picTime.Visible = True
                                    If picTime.Visible Then txtTime.SetFocus
                                Case 10
                                    picDate.Move 80, picTitle + 200
                                    Width = picDate.Width + 160: Height = picStatus.Height + picTitle.Height + picDate.Height + 200
                                    picDate.Visible = True
                                    If picDate.Visible Then mvwDate.SetFocus
                                Case Else
                                    picDate.Move 80, picTitle + 200
                                    picTime.Move picDate.Left + picDate.Width + 80, picTitle + 200, picClock.Width, picDate.Height
                                        txtTime.Move picClock.Left, picClock.Height: lblAmOrPm.Move txtTime.Left + txtTime.Width + 100, txtTime.Top
                                        lblTimeType.Move txtTime.Left, txtTime.Height + txtTime.Top
                                        lstTime.Move txtTime.Left, lblTimeType.Top + lblTimeType.Height, lstTime.Width, picTime.Height - picClock.Height - txtTime.Height - lblTimeType.Height
                                    picDate.Visible = True
                                    If picDate.Visible Then txtTime.SetFocus
                                    picTime.Visible = True
                                    Width = picDate.Width + picTime.Width + 240: Height = picStatus.Height + picTitle.Height + picDate.Height + 200
                            End Select
                        End If
                    Case 3      '�߼���
                        vfg��ѡ.Move 80, picTitle.Height + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height
                        shpBorder1.Move vfg��ѡ.Left - lX, vfg��ѡ.Top - lY, vfg��ѡ.Width + lX * 3, vfg��ѡ.Height + lY * 2
                        vfg��ѡ.Cell(flexcpAlignment, 0, 1, vfg��ѡ.Rows - 1, 1) = flexAlignLeftCenter: vfg��ѡ.Visible = True
                        shpBorder1.Visible = True:       vfg��ѡ.BackColorSel = &HFFC0C0
                    Case Else
                        txt�ı�.Move 80, picTitle.Height + 120, ScaleWidth - 160, IIf(ScaleHeight - 200 - picStatus.Height - picTitle.Height < 0, 0, ScaleHeight - 200 - picStatus.Height - picTitle.Height)
                        shpBorder1.Move txt�ı�.Left - lX, txt�ı�.Top - lY, txt�ı�.Width + lX * 2, txt�ı�.Height + lY * 2
                        txt�ı�.Visible = True: shpBorder1.Visible = True
                        If txt�ı�.Visible And txt�ı�.Enabled Then txt�ı�.SetFocus
                End Select
            Case 1
                Dim lW1 As Long, lW2 As Long, lW3 As Long, lW4 As Long, lW5 As Long
                If Trim(Element.Ҫ�ص�λ) <> "" Then
                    lbl��λ.Width = TextWidth(lbl��λ) + lX * 6
                    lbl��λ.Move ScaleWidth - lbl��λ.Width + lX * 3, picTitle.Height + 170
                    lbl��λ.Visible = True
                    lW5 = lbl��λ.Width
                Else
                    lbl��λ.Visible = False
                    lW5 = 0
                End If
                lW4 = ud����.Width + lX * 4
                ud����.Move ScaleWidth - lW4 - lW5 + lX * 3, picTitle.Height + 120
                ud����.Visible = True
                If Element.Ҫ��С�� > 0 Then
                    txt����2.Width = TextWidth(Space(Element.Ҫ��С��)) + lX * 4
                    lW3 = txt����2.Width + lX
                    txt����2.Move ScaleWidth - lW5 - lW4 - lW3 + lX, picTitle.Height + 170
                    shpBorder2.Move txt����2.Left - lX, txt����2.Top - lY - 50, txt����2.Width + lX * 2, txt����2.Height + 50 + lY * 2
                    shpBorder2.Visible = True
                    txt����2.Visible = True
                    lblDot.Width = TextWidth(".") + lX * 2
                    lW2 = lblDot.Width
                    lblDot.Move txt����2.Left - lW2 + lX * 2, picTitle.Height + 170
                    lblDot.BackStyle = 0
                    lblDot.Visible = True
                Else
                    lW2 = 0
                    lW3 = 0
                    shpBorder2.Visible = False
                    txt����2.Visible = False
                    lblDot.Visible = False
                End If
                lW1 = TextWidth(txt����1.Text) + lX * 2
                lW1 = IIf(lW1 < 400, 400, lW1)
                
                If Width < lW1 + lW2 + lW3 + lW4 + lW5 Then Width = lW1 + lW2 + lW3 + lW4 + lW5
                Height = txt����1.Height + lY * 3 + picStatus.Height + picTitle.Height + 180
                
                txt����1.Move 80, picTitle.Height + 170, ScaleWidth - lW5 - lW4 - lW3 - lW2 - lX * 4
                shpBorder1.Move txt����1.Left - lX, txt����1.Top - lY - 50, txt����1.Width + lX * 2, txt����1.Height + 50 + lY * 2
                txt����1.Visible = True
                shpBorder1.Visible = True
                If txt����1.Visible And txt����1.Enabled Then txt����1.SelStart = 0: txt����1.SelLength = Len(txt����1): txt����1.SetFocus
            Case 2
                If .������̬ = 0 Then '����
                    vfg��ѡ.Move 80, picTitle.Height + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height
                    shpBorder1.Move vfg��ѡ.Left - lX, vfg��ѡ.Top - lY, vfg��ѡ.Width + lX * 3, vfg��ѡ.Height + lY * 2
                    vfg��ѡ.Cell(flexcpAlignment, 0, 1, vfg��ѡ.Rows - 1, 1) = flexAlignLeftCenter: vfg��ѡ.Visible = True
                    shpBorder1.Visible = True: vfg��ѡ.BackColorSel = &HFFC0C0: vfg��ѡ.HighLight = flexHighlightAlways
                Else                   'չ��
                    Dim i As Byte
                    vfg��ѡ.BackColorSel = &HFFFFFF: vfg��ѡ.ForeColorSel = &H0&
                    For i = 0 To vfg��ѡ.Cols - 1
                        If i Mod 2 = 0 Then
                            vfg��ѡ.ColWidth(i) = 250
                        Else
                            vfg��ѡ.ColWidth(i) = IIf((ScaleWidth - 50 - vfg��ѡ.Cols / 2 * 250) / (vfg��ѡ.Cols / 2) > 0, (ScaleWidth - 50 - vfg��ѡ.Cols / 2 * 250) / (vfg��ѡ.Cols / 2), 0)
                        End If
                        vfg��ѡ.Cell(flexcpAlignment, 0, i) = flexAlignLeftCenter
                    Next
                    vfg��ѡ.Move lX, lY, ScaleWidth - lX * 2, ScaleHeight - lY * 2: vfg��ѡ.RowHeight(0) = vfg��ѡ.Height - lX * 2
                    vfg��ѡ.Visible = True: vfg��ѡ.HighLight = flexHighlightNever: vfg��ѡ.Col = 0: vfg��ѡ.Refresh
                End If
                If vfg��ѡ.Visible And vfg��ѡ.Enabled Then vfg��ѡ.SetFocus
            Case 3
                If .������̬ = 0 Then '����
                    vfg��ѡ.Move 80, picTitle.Height + 120, ScaleWidth - 160, ScaleHeight - 200 - picStatus.Height - picTitle.Height
                    shpBorder1.Move vfg��ѡ.Left - lX, vfg��ѡ.Top - lY, vfg��ѡ.Width + lX * 3, vfg��ѡ.Height + lY * 2
                    vfg��ѡ.Visible = True: shpBorder1.Visible = True: vfg��ѡ.HighLight = flexHighlightAlways: vfg��ѡ.BackColorSel = &HFFC0C0
                Else                'չ��
                    vfg��ѡ.Move 0, 0, ScaleWidth, ScaleHeight: vfg��ѡ.RowHeight(0) = ScaleHeight
                    vfg��ѡ.BackColorSel = &HFFFFFF: vfg��ѡ.ForeColorSel = &H0&
                    For i = 0 To vfg��ѡ.Cols - 1
                        vfg��ѡ.ColWidth(i) = ScaleWidth / vfg��ѡ.Cols
                        vfg��ѡ.Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
                    Next
                    vfg��ѡ.Visible = True: vfg��ѡ.HighLight = flexHighlightNever: vfg��ѡ.Col = 0: vfg��ѡ.Refresh
                End If
                If vfg��ѡ.Visible And vfg��ѡ.Enabled Then vfg��ѡ.SetFocus
            End Select
        End If
    End With
    Err.Clear
End Sub
Private Sub txt����1_Change()
    If Element Is Nothing Then Exit Sub
    If txt����1.Tag = "" Then
        Element.�����ı� = Trim(txt����1.Text) & IIf(Element.Ҫ��С�� > 0, "." & Format(Trim(txt����2.Text), String(Element.Ҫ��С��, "0")), "")
    End If
End Sub

Private Sub txt�ı�_Change()
    If Element Is Nothing Then Exit Sub
    Element.�����ı� = Trim(txt�ı�.Text)
End Sub
Private Sub SetCtlFocus()
    '���ÿؼ�����
    If txt����1.Visible And txt����1.Enabled Then
        txt����1.SetFocus
    ElseIf txt����2.Visible And txt����2.Enabled Then
        txt����2.SetFocus
    ElseIf txt�ı�.Visible And txt�ı�.Enabled Then
        txt�ı�.SetFocus
    ElseIf vfg��ѡ.Visible And vfg��ѡ.Enabled Then
        vfg��ѡ.SetFocus
    ElseIf vfg��ѡ.Visible And vfg��ѡ.Enabled Then
        vfg��ѡ.SetFocus
    ElseIf txtTime.Visible And txtTime.Enabled Then
        txtTime.SetFocus
    ElseIf lstDate.Visible And lstDate.Enabled And lstTime.Visible = False Then
        mvwDate.SetFocus
    End If
End Sub

Private Sub UserControl_Show()
Dim strInput As String
    If Element Is Nothing Then Exit Sub
    With Element
        '��ѡ��ѡͨ������1234567890ֱ�Ӷ�λ
        If InStr("1234567890", Chr(mPressKeyAscii)) > 0 Then
            Dim PressN As Integer, i As Integer, strValue As String
            PressN = CByte(Chr(mPressKeyAscii))
            Select Case .Ҫ�ر�ʾ
                Case 2
                    vfg��ѡ.Visible = False
                    If .������̬ = 1 Then 'չ����
                        i = PressN * 2 - 1
                        If i < 0 Then i = 0
                        If i > vfg��ѡ.Cols Then i = 1
                        vfg��ѡ.Col = i
                        For i = 0 To vfg��ѡ.Cols - 1 Step 2
                            If i = vfg��ѡ.Col Or i = vfg��ѡ.Col - 1 Then
                                If vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                                    vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                                Else
                                    vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture
                                End If
                            Else
                                vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                            End If
                            
                            If vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                                If Trim(vfg��ѡ.Cell(flexcpText, 0, i + 1)) = "�Զ���" And .��̬�� = 1 Then
                                    strInput = MidUni(Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName)), 1, 200)
                                    strValue = strValue & IIf(strInput = "", "��", "��") & IIf(strInput = "", "�Զ���", strInput)
                                Else
                                    strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i + 1)
                                End If
                            Else
                                strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i + 1)
                            End If
                        Next
                    Else '������
                        i = PressN - 1
                        If i < 0 Then i = 0
                        If i > vfg��ѡ.Rows - 1 Then i = 0
                        vfg��ѡ.Row = i
                        For i = 0 To vfg��ѡ.Rows - 1
                            If i = vfg��ѡ.Row Then
                                If vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt2.Picture Then
                                    vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                                Else
                                    vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt2.Picture
                                End If
                            Else
                                vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                            End If
                        Next
                        
                        If vfg��ѡ.Cell(flexcpPicture, vfg��ѡ.Row, 0) = imgOpt2.Picture Then
                            strValue = vfg��ѡ.Cell(flexcpText, vfg��ѡ.Row, 1)
                            If strValue = "�Զ���" And .��̬�� = 1 Then
                                strValue = MidUni((Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName))), 1, 200)
                            End If
                        Else
                            strValue = ""
                        End If
                    End If
                    .�����ı� = strValue
                    UserControl_KeyPress vbKeyReturn
                Case 3
                    If .������̬ = 1 Then
                        vfg��ѡ.Visible = False
                        PressN = PressN - 1
                        If PressN < 0 Or PressN >= vfg��ѡ.Cols Then PressN = 0
                        If vfg��ѡ.Cell(flexcpChecked, 0, PressN) = flexChecked Then
                            vfg��ѡ.Cell(flexcpChecked, 0, PressN) = flexUnchecked
                        Else
                            vfg��ѡ.Cell(flexcpChecked, 0, PressN) = flexChecked
                        End If
        
                        For i = 0 To vfg��ѡ.Cols - 1
                            If vfg��ѡ.Cell(flexcpChecked, 0, i) = flexChecked Then
                                If Trim(vfg��ѡ.Cell(flexcpText, 0, i)) = "�Զ���" And .��̬�� = 1 Then
                                    strInput = MidUni(Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName)), 1, 200)
                                    strValue = strValue & IIf(strInput = "", "��", "��") & IIf(strInput = "", "�Զ���", strInput)
                                Else
                                    strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i)
                                End If
                            Else
                                strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i)
                            End If
                        Next
                        .�����ı� = strValue
                        UserControl_KeyPress vbKeyReturn
                    Else
                        i = PressN - 1
                        If i < 0 Then i = 0
                        If i >= vfg��ѡ.Rows Then i = 0
                        If vfg��ѡ.Cell(flexcpChecked, i, 0) = flexChecked Then
                            vfg��ѡ.Cell(flexcpChecked, i, 0) = flexUnchecked
                        Else
                            vfg��ѡ.Cell(flexcpChecked, i, 0) = flexChecked
                        End If
                         
                        For i = 0 To vfg��ѡ.Rows - 1
                            If vfg��ѡ.Cell(flexcpChecked, i, 0) = flexChecked Then
                                If vfg��ѡ.Cell(flexcpText, i, 0) = "�Զ���" And Element.��̬�� = 1 Then
                                    strInput = MidUni(Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName)), 1, 200)
                                    strValue = strValue & "��" & strInput
                                Else
                                    strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, i, 0)
                                End If
                            End If
                        Next
                        strValue = Mid(strValue, 2)
                        .�����ı� = strValue
                        RaiseEvent pChange
                    End If
            End Select
        End If
    End With
End Sub
Private Sub txt����1_GotFocus()
    zlCommFun.OpenIme
    txt����1.SelStart = 0:              txt����1.SelLength = Len(txt����1)
    ud����.BuddyControl = txt����1:     ud����.BuddyProperty = "Text"
End Sub

Private Sub txt����1_KeyPress(KeyAscii As Integer)
    If InStr("1234567890. " & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If KeyAscii = vbKeySpace Or InStr(".", Chr(KeyAscii)) = 1 Then
        KeyAscii = 0
        If txt����2.Visible And txt����2.Enabled Then
            txt����2.SelStart = 0
            txt����2.SelLength = Len(txt����2)
            txt����2.SetFocus
        End If
    End If
End Sub

Private Sub txt����2_Change()
    If Element Is Nothing Then Exit Sub
    If txt����1.Tag = "" Then
        If Element.Ҫ��С�� > 0 Then
            Dim lngLen As Long, strR As String
            lngLen = Len(Trim(txt����2))
            If lngLen > Element.Ҫ��С�� Then
                strR = Trim(txt����1.Text) & "." & Trim(txt����2) & String(Element.Ҫ��С�� - Len(Trim(txt����2)), "0")
            Else
                strR = Trim(txt����1.Text) & "." & Left(Trim(txt����2), Element.Ҫ��С��)
            End If
        Else
            strR = Trim(txt����1.Text)
        End If
        Element.�����ı� = IIf(Element.Ҫ��С�� > 0, Format(strR, "0." & String(Element.Ҫ��С��, "0")), strR)
    End If
End Sub

Private Sub txt����2_GotFocus()
    zlCommFun.OpenIme
    txt����2.SelStart = 0:                      txt����2.SelLength = Len(txt����2)
    ud����.BuddyControl = txt����2:             ud����.BuddyProperty = "Text"
    ud����.Move txt����2.Left + txt����2.Width + 15
End Sub

Private Sub txt����2_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txt�ı�_GotFocus()
    If Element Is Nothing Then Exit Sub
    If Element.Ҫ������ = 0 Then
        zlCommFun.OpenIme
    End If
End Sub
Private Sub txt�ı�_KeyPress(KeyAscii As Integer)
    If Element Is Nothing Then Exit Sub
    If Element.Ҫ������ = 0 Then
        '��ֵ�͵Ŀ��ƣ�ֻ���������֣�С����͸��ţ���С����ֻ��Ϊ1���������ڿ�ͷ������ֻ���ڿ�ʼ����
        'Asc(".") = vbKeyDelete = 46
        If Len(txt�ı�.Text) = 0 And KeyAscii = 46 Then KeyAscii = 0
        If InStr(1, txt�ı�.Text, ".") <> 0 And KeyAscii = 46 Then
            KeyAscii = 0
        ElseIf InStr(1, txt�ı�.Text, ".") = 0 And KeyAscii = 46 And txt�ı�.SelLength = Len(txt�ı�) And txt�ı�.SelStart = 0 Then
            KeyAscii = 0
        End If
        If txt�ı�.Text = "-" And KeyAscii = 46 Then KeyAscii = 0
        If KeyAscii = vbKeyBack Or KeyAscii = 46 Then Exit Sub
        If KeyAscii = Asc("-") Then
            If txt�ı�.SelStart <> 0 Then KeyAscii = 0
        Else
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub vfg��ѡ_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strValue As String, PressN As Integer, strInput As String

    On Error Resume Next
    If Element Is Nothing Then Exit Sub
    If Not KeyCode = vbKeySpace Then Exit Sub
    '�ո�ѡ��
    strValue = ""
    If Element.������̬ = 0 Then '����ʽ
        For i = 0 To vfg��ѡ.Rows - 1
            If i = vfg��ѡ.Row Then
                If vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt2.Picture Then
                    vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                Else
                    vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt2.Picture
                End If
            Else
                vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
            End If
        Next

        If vfg��ѡ.Cell(flexcpPicture, vfg��ѡ.Row, 0) = imgOpt2.Picture Then
            strValue = vfg��ѡ.Cell(flexcpText, vfg��ѡ.Row, 1)
        Else
            If Element.Ҫ������ = 3 Then
                strValue = vfg��ѡ.Cell(flexcpText, 1, 1)
                If strValue = "�Զ���" And Element.��̬�� = 1 Then
                    strInput = MidUni((Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName))), 1, 200)
                    strValue = strInput
                End If
            Else
                strValue = ""
            End If
        End If
    Else 'չ��ʽ
        For i = 0 To vfg��ѡ.Cols - 1 Step 2
            If i = vfg��ѡ.Col Or i = vfg��ѡ.Col - 1 Then
                If vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                    vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                Else
                    vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture
                End If
            Else
                vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
            End If
            If vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                If Trim(vfg��ѡ.Cell(flexcpText, 0, i + 1)) = "�Զ���" And Element.��̬�� = 1 Then
                    strInput = MidUni((Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName))), 1, 200)
                    strValue = strValue & IIf(strInput = "", "��", "��") & IIf(strInput = "", "�Զ���", strInput)
                Else
                    strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i + 1)
                End If
            Else
                strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i + 1)
            End If
        Next
    End If
    Element.�����ı� = strValue
    KeyCode = 0
    UserControl_KeyPress vbKeyReturn
    Err.Clear
End Sub

Private Sub vfg��ѡ_KeyPress(KeyAscii As Integer)
Dim i As Long, strValue As String, PressN As Integer, strInput As String
    If Element Is Nothing Then Exit Sub
    If InStr("1234567890", Chr(KeyAscii)) > 0 And Element.������̬ = 1 Then
        PressN = CByte(Chr(KeyAscii))
        With Element
            i = PressN * 2 - 1
            If i < 0 Then i = 0
            If i > vfg��ѡ.Cols Then i = 1
            vfg��ѡ.Col = i
            For i = 0 To vfg��ѡ.Cols - 1 Step 2
                If i = vfg��ѡ.Col Or i = vfg��ѡ.Col - 1 Then
                    If vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                        vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                    Else
                        vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture
                    End If
                Else
                    vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                End If
    
                If vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                    If Trim(vfg��ѡ.Cell(flexcpText, 0, i + 1)) = "�Զ���" And .��̬�� = 1 Then
                        strInput = MidUni((Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName))), 1, 200)
                        strValue = strValue & IIf(strInput = "", "��", "��") & IIf(strInput = "", "�Զ���", strInput)
                    Else
                        strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i + 1)
                    End If
                Else
                    strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i + 1)
                End If
            Next
            KeyAscii = 0
            Element.�����ı� = strValue
            UserControl_KeyPress vbKeyReturn
         End With
    End If
End Sub


Private Sub vfg��ѡ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, strValue As String, strInput As String
    If mbEt >= 2 Then Exit Sub
    If Element Is Nothing Then Exit Sub
    strValue = ""

    If Button = vbLeftButton Then
        If Element.������̬ = 0 Then
            For i = 0 To vfg��ѡ.Rows - 1
                If i = vfg��ѡ.Row Then
                    If vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt2.Picture Then
                        vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                    Else
                        vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt2.Picture
                    End If
                Else
                    vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                End If
            Next

            If vfg��ѡ.Cell(flexcpPicture, vfg��ѡ.Row, 0) = imgOpt2.Picture Then
                strValue = vfg��ѡ.Cell(flexcpText, vfg��ѡ.Row, 1)
                If strValue = "�Զ���" And Element.��̬�� = 1 Then
                    strValue = MidUni((Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName))), 1, 200)
                End If
            Else
                If Element.Ҫ������ = 3 Then
                    strValue = vfg��ѡ.Cell(flexcpText, 1, 1)
                Else
                    strValue = ""
                End If
            End If
        Else
            For i = 0 To vfg��ѡ.Cols - 1 Step 2
                If i = vfg��ѡ.Col Or i = vfg��ѡ.Col - 1 Then
                    If vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                        vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                    Else
                        vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture
                    End If
                Else
                    vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                End If
                If vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                    If Trim(vfg��ѡ.Cell(flexcpText, 0, i + 1)) = "�Զ���" And Element.��̬�� = 1 Then
                        strInput = MidUni((Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName))), 1, 200)
                        strValue = strValue & IIf(strInput = "", "��", "��") & IIf(strInput = "", "�Զ���", strInput)
                    Else
                        strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i + 1)
                    End If
                Else
                    strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i + 1)
                End If
            Next
        End If
    End If
    Element.�����ı� = strValue
    UserControl_KeyPress vbKeyReturn
End Sub

Private Sub vfg��ѡ_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, strValue As String, strInput As String
    If mbEt < 2 Then Exit Sub
    If Element Is Nothing Then Exit Sub
    strValue = ""

    If Button = vbLeftButton Then
        If Element.������̬ = 0 Then
            For i = 0 To vfg��ѡ.Rows - 1
                If i = vfg��ѡ.Row Then
                    If vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt2.Picture Then
                        vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                    Else
                        vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt2.Picture
                    End If
                Else
                    vfg��ѡ.Cell(flexcpPicture, i, 0) = imgOpt1.Picture
                End If
            Next

            If vfg��ѡ.Cell(flexcpPicture, vfg��ѡ.Row, 0) = imgOpt2.Picture Then
                strValue = vfg��ѡ.Cell(flexcpText, vfg��ѡ.Row, 1)
                If strValue = "�Զ���" And Element.��̬�� = 1 Then
                    strValue = MidUni((Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName))), 1, 200)
                End If
            Else
                If Element.Ҫ������ = 3 Then
                    strValue = vfg��ѡ.Cell(flexcpText, 1, 1)
                Else
                    strValue = ""
                End If
            End If
        Else
            For i = 0 To vfg��ѡ.Cols - 1 Step 2
                If i = vfg��ѡ.Col Or i = vfg��ѡ.Col - 1 Then
                    If vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                        vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                    Else
                        vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture
                    End If
                Else
                    vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt1.Picture
                End If
                If vfg��ѡ.Cell(flexcpPicture, 0, i) = imgOpt2.Picture Then
                    If Trim(vfg��ѡ.Cell(flexcpText, 0, i + 1)) = "�Զ���" And Element.��̬�� = 1 Then
                        strInput = MidUni((Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName))), 1, 200)
                        strValue = strValue & IIf(strInput = "", "��", "��") & IIf(strInput = "", "�Զ���", strInput)
                    Else
                        strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i + 1)
                    End If
                Else
                    strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i + 1)
                End If
            Next
        End If
    End If
    Element.�����ı� = strValue
    UserControl_KeyPress vbKeyReturn
End Sub

Private Sub vfg��ѡ_RowColChange()
    If Element Is Nothing Then Exit Sub
    If Element.������̬ = 1 Then
        vfg��ѡ.Cell(flexcpBackColor, 0, 0, 0, vfg��ѡ.Cols - 1) = 0
        vfg��ѡ.Cell(flexcpBackColor, 0, vfg��ѡ.Col) = &HFFC0C0
    End If
End Sub

Private Sub vfg��ѡ_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, strValue As String, strInput As String
    If Element Is Nothing Then Exit Sub
    strValue = ""
    With Element
        If .������̬ = 0 Then
            For i = 0 To vfg��ѡ.Rows - 1
                If vfg��ѡ.Cell(flexcpChecked, i, 0) = flexChecked Then
                    If vfg��ѡ.Cell(flexcpText, i, 0) = "�Զ���" And .��̬�� = 1 Then
                        strInput = MidUni((Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName))), 1, 200)
                        strValue = strValue & "��" & IIf(strInput = "", "�Զ���", strInput)
                    Else
                        strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, i, 0)
                    End If
                End If
            Next
            strValue = Mid(strValue, 2)
        Else
            For i = 0 To vfg��ѡ.Cols - 1
                If vfg��ѡ.Cell(flexcpChecked, 0, i) = flexChecked Then
                    If Trim(vfg��ѡ.Cell(flexcpText, 0, i)) = "�Զ���" And .��̬�� = 1 Then
                        strInput = MidUni((Trim(InputBox("��¼���Զ���Ҫ��ѡ��", gstrSysName))), 1, 200)
                        strValue = strValue & IIf(strInput = "", "��", "��") & IIf(strInput = "", "�Զ���", strInput)
                    Else
                        strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i)
                    End If
                Else
                    strValue = strValue & "��" & vfg��ѡ.Cell(flexcpText, 0, i)
                End If
            Next
        End If
        .�����ı� = strValue
        RaiseEvent pChange
    End With
End Sub
Private Sub vfg��ѡ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Element Is Nothing Then Exit Sub
    If Element.������̬ = 0 Then
        vfg��ѡ.Col = 0
    End If
End Sub

Private Sub vfg��ѡ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    vfg��ѡ.Col = 0
    Cancel = True
End Sub

Private Sub txtTime_Change()
    If mblnHandTime Then Exit Sub
    Call SetTimer(Val(Mid(txtTime, 1, 2)) + Val(Mid(txtTime, 4, 2)) / 60 + Val(Mid(txtTime, 7, 2)) / 60 / 60)
End Sub


Private Sub lblAmOrPm_Click()
    If Not IsDate(txtTime.Text) Then
        txtTime.SetFocus
        Exit Sub
    End If
    If lblAmOrPm.Caption = "����" Then
        lblAmOrPm.Caption = "����"
        txtTime.Text = Format(CDate(txtTime.Text) - 12 / 24, "hh:mm:ss")
    Else
        lblAmOrPm.Caption = "����"
        txtTime.Text = Format(CDate(txtTime.Text) + 12 / 24, "hh:mm:ss")
    End If
    Call SetTimer(Hour(txtTime.Text) + Minute(txtTime.Text) / 60 + Second(txtTime.Text) / 60 / 60)
End Sub

Private Sub mvwDate_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    Dim intYear As Integer, intMonth As Integer, intDay As Integer
    Dim strYear As String, strMonth As String, strDay As String
    
    If Element Is Nothing Then Exit Sub
    If Element.Ҫ�س��� = 8 Then Exit Sub
    intYear = Year(mvwDate.Value)
    intMonth = Month(mvwDate.Value)
    intDay = Day(mvwDate.Value)
    
    strYear = GetChineseNumber(Mid(intYear, 1, 1), True) & GetChineseNumber(Mid(intYear, 2, 1), True) & GetChineseNumber(Mid(intYear, 3, 1), True) & GetChineseNumber(Mid(intYear, 4, 1), True)
    strMonth = GetChineseNumber(intMonth)
    strDay = GetChineseNumber(intDay)
     
    With lstDate
        .Clear
        .AddItem strYear & "��" & strMonth & "��" & strDay & "��"
        .AddItem strYear & "��" & strMonth & "��"
        .AddItem strMonth & "��" & strDay & "��"
        .AddItem intYear & "��" & intMonth & "��" & intDay & "��"
        .AddItem intYear & "��" & intMonth & "��"
        .AddItem intMonth & "��" & intDay & "��"
        .AddItem intYear & "-" & intMonth & "-" & intDay
        .AddItem intMonth & "-" & intDay
        .AddItem WeekdayName(Weekday(mvwDate.Value))
        .AddItem GetSolarTerm(mvwDate.Value)
        If Val(.Tag) >= 0 Then
            .ListIndex = Val(.Tag)
            .TopIndex = .ListIndex
        End If
    End With
    If Element Is Nothing Then Exit Sub
    If Element.Ҫ������ = 2 Then
        If Element.Ҫ�س��� = 10 Then
            Element.�����ı� = lstDate.Text
            RaiseEvent pChange
        ElseIf Element.Ҫ�س��� > 10 Then
            Element.�����ı� = lstDate.Text & " " & lstTime.Text
            RaiseEvent pChange
        End If
    End If
End Sub

Private Sub picClock_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnHandMove = True: mblnHandTime = True
End Sub

Private Sub picClock_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim xLine As Long, yLine As Long
    Dim dblSin As Double, dblValue As Double

    mblnHandMove = False
    xLine = x - (shpCenter.Left + shpCenter.Width / 2)
    yLine = y - (shpCenter.Top + shpCenter.Height / 2)
    
    If xLine = 0 And yLine = 0 Then Exit Sub
    
    dblSin = yLine / Sqr(xLine ^ 2 + yLine ^ 2)
    If dblSin = 1 Then
        dblValue = 6
    ElseIf dblSin = -1 Then
        dblValue = 0
    Else
        If Sgn(xLine) >= 0 Then
            dblValue = Round(Atn(dblSin / Sqr(-dblSin * dblSin + 1)) / conPI * 180 / 30, 1) + 3
        Else
            dblValue = 9 - Round(Atn(dblSin / Sqr(-dblSin * dblSin + 1)) / conPI * 180 / 30, 1)
        End If
    End If
    If lblAmOrPm.Caption = "����" And dblValue < 12 Then dblValue = dblValue + 12
    Call SetTimer(dblValue)
    mblnHandTime = False
End Sub

Private Function GetSolarTerm(ByVal dtAsk As Date) As String
    '���ܣ����ָ�����ڵĽ���
    '������dtAsk����������
    
    Const conYearMinutes As Double = 525948.76   'ÿ��ķ�������һ��ʵ����365.242194444�죬�����Ӽ��������׼ȷ
    Dim dtBaseDate As Date
    Dim aryTermName() As String
    Dim aryTermData() As String
    
    Dim dblMinutes As Double
    Dim dtTermDate As Date
    
    dtAsk = Int(dtAsk) + 2 / 24 + 5 / 24 / 60
    dtBaseDate = Format("1900-01-06 2:05:00", "YYYY-MM-DD hh:mm:ss")
    If dtAsk < dtBaseDate Then GetSolarTerm = "": Exit Function
    aryTermName = Split("С��,��,����,��ˮ,����,����,����,����,����,С��,â��,����,С��,����,����,����,��¶,���,��¶,˪��,����,Сѩ,��ѩ,����", ",")
    aryTermData = Split("0,21208,42467,63836,85337,107014,128867,150921,173149,195551,218072,240693,263343,285989,308563,331033,353350,375494,397447,419210,440795,462224,483532,504758", ",")
      
    Dim intCount As Integer
    For intCount = 0 To UBound(aryTermData)
        dblMinutes = conYearMinutes * (Year(dtAsk) - 1900) + CLng(aryTermData(intCount))
        dtTermDate = DateAdd("n", dblMinutes, dtBaseDate)
        
        If DateDiff("d", dtAsk, dtTermDate) >= 0 Then
            Select Case DateDiff("d", dtAsk, dtTermDate)
            Case 0
                GetSolarTerm = aryTermName(intCount)
            Case Is < 8
                GetSolarTerm = aryTermName(intCount) & "ǰ" & DateDiff("d", dtAsk, dtTermDate) & "��"
            Case Else
                If intCount <> 0 Then
                    dblMinutes = conYearMinutes * (Year(dtAsk) - 1900) + CLng(aryTermData(intCount - 1))
                    dtTermDate = DateAdd("n", dblMinutes, dtBaseDate)
                    GetSolarTerm = aryTermName(intCount - 1) & "��" & Abs(DateDiff("d", dtAsk, dtTermDate)) & "��"
                Else
                    dblMinutes = conYearMinutes * (Year(dtAsk) - 1 - 1900) + CLng(aryTermData(UBound(aryTermData)))
                    dtTermDate = DateAdd("n", dblMinutes, dtBaseDate)
                    GetSolarTerm = aryTermName(UBound(aryTermData)) & "��" & Abs(DateDiff("d", dtAsk, dtTermDate)) & "��"
                End If
            End Select
            Exit Function
        ElseIf intCount = UBound(aryTermData) Then
            GetSolarTerm = aryTermName(intCount) & "��" & Abs(DateDiff("d", dtAsk, dtTermDate)) & "��"
            Exit Function
        End If
    Next
    GetSolarTerm = ""
End Function

Private Sub DrawWatch()
'������
    Dim CenterX As Long, CenterY As Long, lngRadii As Long
    CenterX = picClock.ScaleWidth / 2
    CenterY = picClock.ScaleHeight / 2
    If CenterX < CenterY Then
        lngRadii = CenterX - 60
    Else
        lngRadii = CenterY - 60
    End If
    
    Dim intHour As Integer, x As Long, y As Long
    x = CenterX - shpCenter.Width / 2
    y = CenterY - shpCenter.Height / 2
    shpCenter.Move x, y
    
    linHand.X1 = x + shpCenter.Width / 2
    linHand.Y1 = y + shpCenter.Height / 2
    
    For intHour = 0 To 11
        If intHour > shpDot.Count - 1 Then
            Load shpDot(intHour)
        End If
        If intHour Mod 3 = 0 Then
            shpDot(intHour).Width = 60
            shpDot(intHour).Height = 60
        Else
            shpDot(intHour).Width = 45
            shpDot(intHour).Height = 45
        End If
        x = CenterX + lngRadii * Sin(intHour * 30 / 180 * conPI) - shpDot(intHour).Width / 2
        y = CenterY - lngRadii * Cos(intHour * 30 / 180 * conPI) - shpDot(intHour).Height / 2
        shpDot(intHour).Move x, y
        shpDot(intHour).Visible = True
    Next
End Sub

Private Sub SetTimer(ByVal dblTime As Double)
'����ʱ�䶨ָ��
    Dim CenterX As Long, CenterY As Long, lngRadii As Long
    Dim intCount As Integer

    If Element Is Nothing Then Exit Sub
    CenterX = picClock.ScaleWidth / 2
    CenterY = picClock.ScaleHeight / 2
    If CenterX < CenterY Then
        lngRadii = CenterX - 60
    Else
        lngRadii = CenterY - 60
    End If
    
    If dblTime < 12 Then
        lblAmOrPm.Caption = "����"
    Else
        lblAmOrPm.Caption = "����"
    End If
    
    If mblnHandTime = True Then
        txtTime.Text = Format(dblTime / 24, "hh:mm:ss")
    End If
    linHand.X2 = CenterX + lngRadii * Sin(dblTime * 30 / 180 * conPI)
    linHand.Y2 = CenterY - lngRadii * Cos(dblTime * 30 / 180 * conPI)
    
    For intCount = 0 To 11
        If intCount = IIf(dblTime < 12, dblTime, dblTime - 12) Then
            shpDot(intCount).BorderColor = RGB(255, 0, 0)
        Else
            shpDot(intCount).BorderColor = RGB(0, 0, 0)
        End If
    Next

    If mblnHandMove = True Then Exit Sub

    '���ø�ʽ
    Dim intHour As Integer, intMinute As Integer, intSecond As Integer
    
    intHour = Val(Mid(txtTime.Text, 1, 2))
    intMinute = Val(Mid(txtTime.Text, 4, 2))
    intSecond = Val(Mid(txtTime.Text, 7, 2))
    
    With lstTime
        .Clear
        .AddItem IIf(intHour < 12, "����", "����") & GetChineseNumber(IIf(intHour < 12, intHour, intHour - 12)) & "ʱ" & GetChineseNumber(intMinute) & "��"
        .AddItem GetChineseNumber(intHour) & "ʱ" & GetChineseNumber(intMinute) & "��"
        .AddItem IIf(intHour < 12, "����", "����") & IIf(intHour < 12, intHour, intHour - 12) & "ʱ" & intMinute & "��" & intSecond & "��"
        .AddItem IIf(intHour < 12, "����", "����") & IIf(intHour < 12, intHour, intHour - 12) & "ʱ" & intMinute & "��"
        .AddItem intHour & "ʱ" & intMinute & "��" & intSecond & "��"
        .AddItem intHour & "ʱ" & intMinute & "��"
        .AddItem IIf(intHour < 12, intHour, intHour - 12) & ":" & Format(intMinute, "00") & ":" & Format(intSecond, "00") & IIf(intHour < 12, " AM", " PM")
        .AddItem IIf(intHour < 12, intHour, intHour - 12) & ":" & Format(intMinute, "00") & IIf(intHour < 12, " AM", " PM")
        .AddItem intHour & ":" & Format(intMinute, "00") & ":" & Format(intSecond, "00")
        .AddItem intHour & ":" & Format(intMinute, "00")
        If Val(.Tag) >= 0 Then
            .ListIndex = Val(.Tag)
            .TopIndex = .ListIndex
        End If
    End With

    With Element
        If .Ҫ������ = 2 Then
            Select Case .Ҫ�س���
                Case 8
                    .�����ı� = lstTime.Text
                Case 10
                    .�����ı� = lstDate.Text
                Case 19
                    .�����ı� = lstDate.Text & " " & lstTime.Text
            End Select
            RaiseEvent pChange
        End If
    End With
End Sub

Private Function GetChineseNumber(ByVal bytNumber As Byte, Optional blnZeroCircle As Boolean) As String
    '���ܣ����غ�������
    '������
    '   bytNumber,Ҫ��������֣�������Ҫ�󲻴���99;
    '   blnZeroCircle,�Ƿ��ԡ����0���������Ϊ��
    
    Dim bytBit1 As Byte, bytBit2 As Byte
    Dim strBit1 As String, strBit2 As String
    
    If bytNumber > 99 Then GetChineseNumber = "": Exit Function
    
    bytBit1 = bytNumber \ 10: bytBit2 = bytNumber Mod 10
    
    If bytBit1 = 0 Then
        strBit1 = ""
        If blnZeroCircle = False Then
            strBit2 = Split("��,һ,��,��,��,��,��,��,��,��", ",")(bytBit2)
        Else
            strBit2 = Split("��,һ,��,��,��,��,��,��,��,��", ",")(bytBit2)
        End If
    Else
        strBit1 = Split(",,��,��,��,��,��,��,��,��", ",")(bytBit1) & "ʮ"
        strBit2 = Split(",һ,��,��,��,��,��,��,��,��", ",")(bytBit2)
    End If
    GetChineseNumber = strBit1 & strBit2
End Function
Private Sub DoFind()
Dim lngMatch As Long, rsTemp As ADODB.Recordset, lngCount As Long
    If Element Is Nothing Then Exit Sub
    lngMatch = Val(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0))
    If Trim(txtFind.Text) = "" Then Exit Sub
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ����, ����, ���� From The (Select Cast(Zl_Dic_Search([1], [2], " & lngMatch & ") As " & gstrDbOwner & ".t_Dic_Rowset) From Dual)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ֵ�", Element.Ҫ������, Trim(txtFind.Text))
    Set vgdList.DataSource = rsTemp
    With vgdList
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftCenter
        Next
    End With
    If rsTemp.RecordCount = 0 Then
        cbrThis.Buttons(2).Enabled = False
        lblInfo = Element.Ҫ������ & " û��ƥ�����Ŀ"
        txtFind.SelStart = 0: txtFind.SelLength = 1000: If txtFind.Visible And txtFind.Enabled Then txtFind.SetFocus
    Else
        cbrThis.Buttons(2).Enabled = True
        lblInfo = Element.Ҫ������ & " ��ѡ��ϣ������Ŀ"
        If vgdList.Visible And vgdList.Enabled Then vgdList.SetFocus
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfg��ѡ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Element Is Nothing Then Exit Sub
    If Button = vbLeftButton Then
        If Element.������̬ = 1 Then
            Dim i As Integer, strValue As String
            vfg��ѡ.Cell(flexcpBackColor, 0, 0, 0, vfg��ѡ.Cols - 1) = 0
            vfg��ѡ.Cell(flexcpBackColor, 0, vfg��ѡ.Col) = &HFFC0C0: vfg��ѡ.Refresh
            
            If x > vfg��ѡ.Cell(flexcpLeft, 0, vfg��ѡ.Col) + 200 Then  'ѡ�а�Ť��Χ�㰴AfterEdit�¼�����
                If vfg��ѡ.Cell(flexcpChecked, 0, vfg��ѡ.Col) = flexUnchecked Then
                    vfg��ѡ.Cell(flexcpChecked, 0, vfg��ѡ.Col) = flexChecked
                Else
                    vfg��ѡ.Cell(flexcpChecked, 0, vfg��ѡ.Col) = flexUnchecked
                End If
                Call vfg��ѡ_AfterEdit(0, vfg��ѡ.Col)
            End If
        End If
    End If
End Sub

Private Sub vgdList_DblClick()
Dim i As Long, strReturn As String
    If Element Is Nothing Then Exit Sub
    If vgdList.Row <= 0 Then Exit Sub
    strReturn = ""
    For i = 0 To vgdList.Cols - 1
        strReturn = strReturn & ";" & vgdList.TextMatrix(vgdList.Row, i)
    Next
    If Len(strReturn) > 0 Then strReturn = Mid(strReturn, 2)
    If strReturn = ";;" Then Exit Sub
    Element.�����ı� = Split(strReturn, ";")(1)
    
    RaiseEvent pOk
End Sub

Private Sub vgdList_GotFocus()
    shpBorder1.BorderWidth = 2
End Sub

Private Sub vgdList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then vgdList_DblClick
End Sub

Private Sub vgdList_LostFocus()
    shpBorder1.BorderWidth = 1
End Sub


