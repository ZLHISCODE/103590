VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProcUpgrade 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "�䶯������������"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   12360
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmProcUpgrade.frx":0000
   ScaleHeight     =   7845
   ScaleWidth      =   12360
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid vsfModule 
      Height          =   2535
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   7335
      _cx             =   12938
      _cy             =   4471
      Appearance      =   1
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483636
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsfSp 
      Height          =   1935
      Left            =   1560
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   4935
      _cx             =   8705
      _cy             =   3413
      Appearance      =   1
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483636
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
   Begin VB.PictureBox pctBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   12360
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2640
      Width           =   12360
      Begin VB.CommandButton cmdManual 
         Caption         =   "���̵���(&U)"
         Height          =   345
         Left            =   7320
         TabIndex        =   28
         Top             =   443
         Width           =   1215
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "�����ű�(&I)"
         Height          =   345
         Left            =   8640
         TabIndex        =   27
         Top             =   443
         Width           =   1215
      End
      Begin VB.Frame fra2 
         Height          =   30
         Index           =   1
         Left            =   2640
         TabIndex        =   25
         Top             =   120
         Width           =   9615
      End
      Begin VB.Frame fra1 
         Height          =   30
         Index           =   1
         Left            =   0
         TabIndex        =   24
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtFind 
         ForeColor       =   &H80000000&
         Height          =   270
         Left            =   1800
         TabIndex        =   18
         Text            =   "��������ƻ��޸��˺󰴻س����ж�λ"
         Top             =   480
         Width           =   3735
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAlter 
         Height          =   2415
         Left            =   3960
         TabIndex        =   22
         Top             =   840
         Width           =   7000
         _cx             =   12347
         _cy             =   4260
         Appearance      =   1
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
         ForeColorFixed  =   -2147483636
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   200
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfProc 
         Height          =   2415
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   3735
         _cx             =   6588
         _cy             =   4260
         Appearance      =   1
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483636
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   200
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
         ExplorerBar     =   1
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
      Begin VB.Label lblCheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�û��䶯������������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   26
         Top             =   0
         Width           =   1800
      End
      Begin VB.Label lblFind 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1320
         TabIndex        =   21
         Top             =   525
         Width           =   360
      End
      Begin VB.Label lblAlter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������ᱻ�޸ĵ��û��䶯����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4320
         TabIndex        =   20
         Top             =   525
         Width           =   2520
      End
      Begin VB.Label lblProc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�û��䶯����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   525
         Width           =   1080
      End
   End
   Begin VB.Frame fra1 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "���(&A)"
      Height          =   350
      Left            =   10200
      TabIndex        =   3
      Top             =   1970
      Width           =   1455
   End
   Begin VB.Frame fra2 
      Height          =   30
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   9975
   End
   Begin VB.Label lblSp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ȷ����ǰϵͳ��װĿ¼������SP�ű�û����©."
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   2520
      TabIndex        =   31
      Top             =   2295
      Width           =   4050
   End
   Begin VB.Label lblSp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����SP�ű�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   1
      Left            =   1540
      TabIndex        =   30
      Top             =   2295
      Width           =   900
   End
   Begin VB.Label lblSp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰϵͳִ�й�"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   2295
      Width           =   1260
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   15
   End
   Begin VB.Label lblWarn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ݱ�����װ�������ű��еı�׼��Ʒ���������ݿ��еĹ��̽��жԱ�,�ҳ���Щ���޸��˵Ĺ��̣��Լ��Ƿ����������޸�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   9810
   End
   Begin VB.Label lblVisable 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label lblCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "�û��䶯���̼��"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label lblCurrent 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��ǰ�汾ϵͳ��װĿ¼��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6840
      TabIndex        =   9
      Top             =   2040
      Width           =   1980
   End
   Begin VB.Label lblTarget 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Ŀ��汾ϵͳ��װĿ¼��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2760
      TabIndex        =   8
      Top             =   2040
      Width           =   1980
   End
   Begin VB.Label lblCurPath 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   8880
      TabIndex        =   7
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label lblTargetPath 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C:\AppSoft"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4800
      TabIndex        =   6
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label lblCurCmd 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   9240
      TabIndex        =   5
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label lblTargetCmd 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   6000
      TabIndex        =   4
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   $"frmProcUpgrade.frx":803A
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   10980
   End
   Begin VB.Image Img 
      Height          =   555
      Left            =   240
      Picture         =   "frmProcUpgrade.frx":8114
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�䶯������������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmProcUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmCollect As frmProcCollect
Attribute mfrmCollect.VB_VarHelpID = -1
Private mblnChanged As Boolean
Private Enum txtColor
    ��ɫ = &H80000012
    ��ɫ = &H80000010
End Enum

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���

End Function

Private Sub cmdCheck_Click()
    Dim strSys As String, i As Integer
    Dim strMsg As String, intNum As Integer
    Dim strInitFile As String, strCurInitPath As String

    strCurInitPath = lblCurPath.Caption
    With vsfModule
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                '��ȡ��ǰϵͳ�������ļ�
                strInitFile = lblTargetPath.Caption & "\" & Decode(.TextMatrix(i, .ColIndex("���")) \ 100, 1, "ZLHIS10", 3, "ZLMEDREC10", 4, "ZLMATERIAL10", _
                                                                    6, "ZLDEVICE10", 21, "ZLPEIS10", 22, "ZLBLOOD10", _
                                                                    23, "ZLINFECT10", 24, "ZLOPER10", _
                                                                    25, "ZLLIS10", 26, "ZLPSS10", 27, "ZLHEC10") & "\Ӧ�ýű�\ZLSETUP.INI"
                                                                    
                '���������ʽΪ: "ϵͳ��,ϵͳ����,��ǰ�汾,Ŀ��汾,Ŀ¼",���ϵͳ֮���÷ֺż��
                If strSys = "" Then
                    strSys = .TextMatrix(i, .ColIndex("���")) & "," & .TextMatrix(i, .ColIndex("ϵͳ����")) & "," & _
                                .TextMatrix(i, .ColIndex("��ǰ�汾��")) & "," & .TextMatrix(i, .ColIndex("Ŀ��汾��")) & "," & strInitFile
                Else
                    strSys = strSys & ";" & .TextMatrix(i, .ColIndex("���")) & "," & .TextMatrix(i, .ColIndex("ϵͳ����")) & "," & _
                                                    .TextMatrix(i, .ColIndex("��ǰ�汾��")) & "," & .TextMatrix(i, .ColIndex("Ŀ��汾��")) & "," & strInitFile
                End If
                intNum = intNum + 1
            End If
        Next
    End With
    
    If intNum = 0 Then
        strMsg = "û��ѡ���κ�ϵͳ��������ѡ��"
        MsgBox strMsg, vbApplicationModal, gstrSysName
        Exit Sub
    End If
    
    strMsg = "��ѡ��" & intNum & "��Ӧ��ϵͳ��Ϊ��֤���������ȷ�ԣ���ȷ���ű��ļ��������ԡ�" & vbNewLine & _
                    "���ִ����ɺ󣬽����ϴμ�������е�����ͬʱɾ���ϴε����Ĺ��̴��룬��ȷ��Ҫ������"
                    
    If MsgBox(strMsg, vbYesNo, "���ȷ��") = vbNo Then Exit Sub
    
    vsfAlter.Rows = 1
    vsfProc.Rows = 1
    Set mfrmCollect = New frmProcCollect
    mfrmCollect.ShowMe strSys, strCurInitPath
    
End Sub

Private Sub cmdExport_Click()
    '�ű�����
    Dim strPath As String, i As Long
    Dim blnExp As Boolean, lngNum As Long
    Dim strProc As String, strName As String
    
    On Error GoTo errH

    With vsfAlter
        If .Rows = 1 Then
            MsgBox "��������û�б䶯���̱��޸ģ����赼����", , "��ʾ"
            Exit Sub
        End If
        
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����״̬")) = "�ѵ���" Then
                blnExp = True
            ElseIf .TextMatrix(i, .ColIndex("����״̬")) = "������" Then
                lngNum = lngNum + 1
            End If
        Next
        
        If Not blnExp Then
            MsgBox "���ȶ��������б��޸ĵı䶯���̽��м��������ٽ��е�����", , "��ʾ"
            Exit Sub
        Else
            If lngNum > 0 Then
                MsgBox "Ŀǰϵͳ����" & lngNum & "������δ�����˹����������ò��ֹ��̲���ͨ���ű�������" & _
                            vbNewLine & "Ϊ���������©�������˹�������Ըò��ֹ�����ִ�нű�����", , "��ʾ"
            End If
            
            strPath = OpenFolder(Me, "��ѡ�񵼳��ű�Ŀ¼")
            If strPath = "" Then Exit Sub
            If Right(strPath, 1) <> "\" Then
                strPath = strPath & "\"
            End If
            strPath = strPath & "ProcExport.Sql"
            
            For i = 1 To .Rows - 1
                If i = 1 Then
                    gobjFile.CreateTextFile strPath
                End If
                
                If .TextMatrix(i, .ColIndex("����״̬")) = "�ѵ���" Then
                    strName = .TextMatrix(i, .ColIndex("��������"))
                    strProc = GetPorcTxtByName(strName, 3)
                    
                    '��Ҫת���Ĺ�����������20,�Ͳ�ִ��ת��
                    If .Rows - lngNum > 20 Then
                        ShowFlash "���ڽ�����" & strName & "�������ű�"
                    End If
                    
                    Do While Right(strProc, 2) = vbNewLine
                        strProc = Left(strProc, Len(strProc) - 2)
                    Loop
                    
                    gobjFile.OpenTextFile(strPath, ForAppending).Write strProc & vbNewLine & "/" & vbNewLine '�����ű�
                    gcnOracle.Execute "Update zlProcedure Set ״̬ = 4 Where ���� =  '" & strName & "'" '�޸�״̬
                    .TextMatrix(i, .ColIndex("����״̬")) = "�ѵ���"
                End If
            Next
            gobjFile.OpenTextFile(strPath).Close
            ShowFlash ""
            MsgBox "���̵����ɹ���" & vbNewLine & "�Ѿ������̱�����" & strPath, , "��ʾ"
        End If
    End With
    Exit Sub
errH:
    ShowFlash ""
    If 0 = 1 Then
        Resume
    End If
    MsgBox "�����ű����ִ���" & vbNewLine & Err.Description, , gstrSysName
End Sub

Private Sub cmdManual_Click()
    Dim arrIds() As String, lngIdx As Long
    Dim i As Long
    
    With vsfAlter
        If .Row = 0 Then
            MsgBox "��ѡ�����ڱ��������в��ᱻ�޸ģ�����������", , gstrSysName
            Exit Sub
        End If
        
        '��ΪҪ��������,���԰���Ҫ�����Ĺ���ID�������Ӵ���
        lngIdx = .Row - 1
        ReDim arrIds(.Rows - 2)
        
        For i = 1 To .Rows - 1
            arrIds(i - 1) = .RowData(i) & ":" & .TextMatrix(i, .ColIndex("��������"))
        Next
        
    End With
    
    If frmProcDiff.ShowMe(arrIds, lngIdx) Then
        LoadProc
    End If
End Sub


Private Sub Form_Activate()
    Call LoadSystems
    Call LoadSpVer
    Call LoadProc
End Sub


Private Sub Form_Load()
    Dim strCol As String

    '����ʼ��
    strCol = " ,400,1;���,1000,1;ϵͳ����,2000,1;��ǰ�汾��,1800,1;Ŀ��汾��,1800,1"
    Call InitTable(vsfModule, strCol)
    vsfModule.FixedCols = 1
    vsfModule.ColDataType(0) = flexDTBoolean
    vsfModule.Cell(flexcpChecked, 0, 0) = flexUnchecked
    vsfModule.Cell(flexcpForeColor, 0, 0, 0, vsfModule.Cols - 1) = &H80000008
    
    strCol = " ,350,1;ϵͳ,2000,1;��������,2000,1;����ǰ���½ű�,2000,1;�޸���,2000,1;�޸�ʱ��,2000,1;�޸�˵��,2000,1"
    Call InitTable(vsfProc, strCol)
    vsfProc.FixedCols = 1
    vsfProc.Rows = 1
    vsfProc.Cell(flexcpForeColor, 0, 0, 0, vsfProc.Cols - 1) = &H80000008
    
    strCol = " ,390,1;��������,3000,1;���������½ű�,2400,1;����״̬,500,1"
    Call InitTable(vsfAlter, strCol)
    vsfAlter.FixedCols = 1
    vsfAlter.Rows = 1
    vsfAlter.Cell(flexcpForeColor, 0, 0, 0, vsfAlter.Cols - 1) = &H80000008
    
    strCol = "���,600,1;ϵͳ,2000,1;����SP�汾,2000,1"
    Call InitTable(vsfSp, strCol)
End Sub

Private Sub ResizeLable()
    On Error Resume Next
    
    
    lblSystem.Width = IIf(lblSystem.Width > 4000, 4000, lblSystem.Width)
    lblSystem.Left = lblWarn.Left
    lblVisable.Left = lblSystem.Left + lblSystem.Width + 60
    lblTarget.Left = lblVisable.Left + lblVisable.Width + 240
    lblTargetPath.Left = lblTarget.Left + lblTarget.Width + 60
    lblTargetCmd.Left = lblTargetPath.Left + lblTargetPath.Width + 60
    lblCurrent.Left = lblTargetCmd.Left + lblTargetCmd.Width + 240
    lblCurPath.Left = lblCurrent.Left + lblCurrent.Width + 60
    lblCurCmd.Left = lblCurPath.Left + lblCurPath.Width + 60
    If lblCurCmd.Visible Then
        cmdCheck.Left = lblCurCmd.Left + lblCurCmd.Width + 240
    Else
        cmdCheck.Left = lblTargetCmd.Left + lblTargetCmd.Width + 240
    End If
    
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    fra2(0).Width = Me.ScaleWidth - fra2(0).Left
    fra2(1).Width = Me.ScaleWidth - fra2(1).Left
    
    ResizeLable

    vsfModule.Left = lblWarn.Left
    
    pctBottom.Width = Me.ScaleWidth
    pctBottom.Height = Me.ScaleHeight - pctBottom.Top
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mfrmCollect = Nothing
End Sub

Private Sub lblSp_Click(Index As Integer)
    If Index <> 1 Then Exit Sub
    With vsfSp
        .Visible = Not .Visible
        If .Visible Then .SetFocus  '�ɼ��ͻ�ȡ����
    End With
End Sub

Private Sub lblVisable_Click()
    Dim i As Long, intNum  As Integer
    
    With vsfModule
        .Visible = Not .Visible
        
        If .Visible Then .SetFocus  '�ɼ��ͻ�ȡ����
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                intNum = intNum + 1
            End If
        Next
        lblSystem.Caption = "����" & .Rows - 1 & "��ϵͳ����ѡ" & intNum & "��ϵͳ"
        ResizeLable
    End With
End Sub

Private Sub pctBottom_Resize()
    On Error Resume Next


    
    vsfProc.Width = pctBottom.Width - vsfProc.Left - vsfAlter.Width - 500
    vsfAlter.Left = vsfProc.Width + vsfProc.Left + 360
    vsfProc.Height = pctBottom.ScaleHeight - vsfProc.Top - 200
    vsfAlter.Height = vsfProc.Height
    lblProc.Left = vsfProc.Left
    lblAlter.Left = vsfAlter.Left
    txtFind.Left = vsfProc.Left + vsfProc.Width - txtFind.Width
    lblFind.Left = txtFind.Left - lblFind.Width - 60

    cmdExport.Left = vsfAlter.Width + vsfAlter.Left - cmdExport.Width
    cmdManual.Left = cmdExport.Left - cmdManual.Width - 40
End Sub


Private Sub LoadProc()
    '�������ݿ��б���ı䶯����
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    ShowFlash "���ڼ��ر䶯����..."
    strSQL = "Select a.Id, a.ϵͳ���, a.����, a.����, a.״̬, a.������, a.�޸���Ա, To_Char(a.�޸�ʱ��, 'yyyy-mm-dd hh24:mi') �޸�ʱ��, a.�ϴ��޸���Ա," & vbNewLine & _
                "       To_Char(a.�ϴ��޸�ʱ��, 'yyyy-mm-dd hh24:mi') �ϴ��޸�ʱ��, a.����ǰ�汾, a.������汾, a.����, a.˵��, c.���� ϵͳ" & vbNewLine & _
                "From (Select Distinct a.Id, a.ϵͳ���, a.����, a.����, a.״̬, a.������, a.�޸���Ա, a.�޸�ʱ��, a.�ϴ��޸���Ա, a.�ϴ��޸�ʱ��, a.����ǰ�汾, a.������汾, b.����," & vbNewLine & _
                "                       a.˵��" & vbNewLine & _
                "       From Zlprocedure A, Zlproceduretext B" & vbNewLine & _
                "       Where ���� = 1 And a.Id = b.����id And (b.���� = 1 Or b.���� = 4)) A, zlSystems C" & vbNewLine & _
                "Where a.ϵͳ��� = c.���" & vbNewLine & _
                "Order By a.ϵͳ���, a.����"
    Set rsTmp = OpenSQLRecord(strSQL, "���ر䶯����")
    
    '���ر䶯����
    With vsfProc
        rsTmp.Filter = "���� = 1"
        If rsTmp.RecordCount = 0 Then Exit Sub
        rsTmp.MoveFirst
        
        .Rows = 1
        .Rows = rsTmp.RecordCount + 1
        .MergeCells = flexMergeRestrictRows
        .MergeCol(.ColIndex("ϵͳ")) = True
        
        .Redraw = flexRDNone
        i = 1
        Do While Not rsTmp.EOF
            .TextMatrix(i, 0) = i
            .TextMatrix(i, .ColIndex("ϵͳ")) = rsTmp!ϵͳ & ""
            .TextMatrix(i, .ColIndex("��������")) = rsTmp!���� & ""
            .TextMatrix(i, .ColIndex("�޸���")) = rsTmp!�޸���Ա & ""
            .TextMatrix(i, .ColIndex("�޸�ʱ��")) = rsTmp!�޸�ʱ�� & ""
            .TextMatrix(i, .ColIndex("�޸�˵��")) = rsTmp!˵�� & ""
            .TextMatrix(i, .ColIndex("����ǰ���½ű�")) = rsTmp!����ǰ�汾 & ""
            .RowData(i) = rsTmp!Id & ""
            i = i + 1
            rsTmp.MoveNext
        Loop
        .AutoResize = True
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDDirect
        
    End With
    
    '�����޸ĵı䶯����
    rsTmp.Filter = "���� = 4"
    If rsTmp.RecordCount = 0 Then Exit Sub
    rsTmp.MoveFirst
    
    With vsfAlter
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTmp.RecordCount + 1
        
        i = 1
        Do While Not rsTmp.EOF
            .TextMatrix(i, 0) = i
            .TextMatrix(i, .ColIndex("��������")) = rsTmp!���� & ""
            .TextMatrix(i, .ColIndex("���������½ű�")) = rsTmp!������汾 & ""
            .TextMatrix(i, .ColIndex("����״̬")) = Decode(rsTmp!״̬, 1, "������", 2, "������", 3, "�ѵ���", 4, "�ѵ���") & ""
            If rsTmp!״̬ = 1 Then
                .Cell(flexcpForeColor, i, .ColIndex("����״̬")) = ��ɫ
            Else
                .Cell(flexcpForeColor, i, .ColIndex("����״̬")) = ��ɫ
            End If
            .RowData(i) = rsTmp!Id & ""
            
            rsTmp.MoveNext
            i = i + 1
        Loop
        .Redraw = flexRDDirect
    End With
    
    ShowFlash ""
    Exit Sub
errH:
    ShowFlash ""
    If 0 = 1 Then
        Resume
    End If
    MsgBox Err.Description, , gstrSysName
End Sub

Private Sub LoadSystems()
    '���ذ�װ��ϵͳ
    Dim strSQL As String, rsSys As New ADODB.Recordset
    Dim i As Long, strTmp As String
    
    '���Ȼ�ȡϵͳ��ŵ���Ϣ
    strSQL = "Select ��� ϵͳ���, ���� ϵͳ����, �汾�� ϵͳ�汾��, ������ ϵͳ������, ������װ From Zlsystems where Upper(������)=[1] Order by Nvl(�����,0),���"
    Set rsSys = OpenSQLRecord(strSQL, "��ȡ��װϵͳ", gstrUserName)
    
    If rsSys.RecordCount = 0 Then
        MsgBox "��ʹ��ϵͳ�����ߵ�¼��", , gstrSysName
        Exit Sub
    Else
        With vsfModule
            i = .FixedRows
            .Rows = .FixedRows
            .Rows = rsSys.RecordCount + .FixedRows
            Do While Not rsSys.EOF
                .TextMatrix(i, .ColIndex("���")) = rsSys!ϵͳ��� & ""
                .TextMatrix(i, .ColIndex("ϵͳ����")) = rsSys!ϵͳ���� & ""
                .TextMatrix(i, .ColIndex("��ǰ�汾��")) = rsSys!ϵͳ�汾�� & ""
                .TextMatrix(i, .ColIndex("Ŀ��汾��")) = ""
                rsSys.MoveNext
                i = i + 1
            Loop
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = flexAlignCenterCenter
        End With
    End If
    
    LoadUpdateSystem lblTargetPath.Caption
End Sub

Private Sub LoadUpdateSystem(ByVal strPath As String)
    '��ȡ����ϵͳ������Ŀ��汾
    Dim i As Long, strInitFile As String
    Dim strTarget As String, blnStep As Boolean
    Dim intNum As Integer
    
    With vsfModule
        For i = 1 To .Rows - 1
            strInitFile = strPath & "\" & Decode(.TextMatrix(i, .ColIndex("���")) \ 100, 1, "ZLHIS10", 3, "ZLMEDREC10", 4, "ZLMATERIAL10", _
                                                                                6, "ZLDEVICE10", 21, "ZLPEIS10", 22, "ZLBLOOD10", _
                                                                                23, "ZLINFECT10", 24, "ZLOPER10", _
                                                                                25, "ZLLIS10", 26, "ZLPSS10", 27, "ZLHEC10") & "\Ӧ�ýű�\ZLSETUP.INI"
            If gobjFile.FileExists(strInitFile) Then
                If GetUpgradeFiles(Nothing, Val(.TextMatrix(i, .ColIndex("���"))), .TextMatrix(i, .ColIndex("��ǰ�汾��")), strInitFile, , , , strTarget, , True, False) Is Nothing Then
                    .Cell(flexcpText, 1, .ColIndex("Ŀ��汾��"), .Rows - 1, .ColIndex("Ŀ��汾��")) = ""
                    .Cell(flexcpChecked, i, 0) = flexUnchecked
                    Exit Sub
                End If
                .TextMatrix(i, .ColIndex("Ŀ��汾��")) = strTarget
                
                If strTarget <> "" Then
                    intNum = intNum + 1
                    .Cell(flexcpChecked, i, 0) = flexChecked
                    
                    '����Ƿ��汾����
                    If .TextMatrix(i, .ColIndex("��ǰ�汾��")) <> "" And GetPrimaryVer(.TextMatrix(i, .ColIndex("��ǰ�汾��"))) <> GetPrimaryVer(strTarget) Then
                        blnStep = True
                    End If
                Else
                    .TextMatrix(i, .ColIndex("Ŀ��汾��")) = ""
                    .Cell(flexcpChecked, i, 0) = flexUnchecked
                End If
            Else
                .TextMatrix(i, .ColIndex("Ŀ��汾��")) = ""
                .Cell(flexcpChecked, i, 0) = flexUnchecked
            End If
        Next
        
        lblSystem.Caption = "����" & .Rows - 1 & "��ϵͳ����ѡ" & intNum & "��ϵͳ"
        lblCurrent.Visible = blnStep
        lblCurPath.Visible = blnStep
        lblCurCmd.Visible = blnStep
        ResizeLable
    End With
End Sub

Private Sub LoadSpVer()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long
    
    strSQL = "Select a.ϵͳ, b.����,a.Ŀ��汾 ����SP�汾" & vbNewLine & _
                "From zlUpGrade A, zlSystems B" & vbNewLine & _
                "Where a.����汾 Like '%.%.%.%' And a.ϵͳ = b.��� And" & vbNewLine & _
                "      Substr(a.����汾, 1, Instr(a.����汾, '.', 1, 2) - 1) = Substr(b.�汾��, 1, Instr(b.�汾��, '.', 1, 2) - 1)" & vbNewLine & _
                "Order By a.ϵͳ"
    
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ����SP�汾")
    
    If rsTmp.RecordCount = 0 Then
        lblSp(0).Visible = False
        lblSp(1).Visible = False
        lblSp(2).Visible = False
    Else
        lblSp(0).Visible = True
        lblSp(1).Visible = True
        lblSp(2).Visible = True
        
        With vsfSp
            .Rows = 1: .Rows = rsTmp.RecordCount + 1
            i = 1
            Do While Not rsTmp.EOF
                .TextMatrix(i, .ColIndex("���")) = rsTmp!ϵͳ
                .TextMatrix(i, .ColIndex("ϵͳ")) = rsTmp!����
                .TextMatrix(i, .ColIndex("����SP�汾")) = rsTmp!����SP�汾
                i = i + 1
                rsTmp.MoveNext
            Loop
        End With

    End If
    
End Sub

Private Sub mfrmCollect_ReturnChangedProc(ByVal rsTmp As ADODB.Recordset, ByVal intType As Integer)
    '���յ���¼��������������
    'intType: 1��ʾ�䶯���� 2��ʾ�����ű��еı䶯����
    Dim i As Long
    
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    If intType = 1 Then
        
        With vsfProc
            .Redraw = flexRDNone
            .MergeCells = flexMergeRestrictRows
            .MergeCol(.ColIndex("ϵͳ")) = True
            i = .Rows
            .Rows = rsTmp.RecordCount + .Rows
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                .TextMatrix(i, .ColIndex("ϵͳ")) = rsTmp!P_System & ""
                .TextMatrix(i, 0) = i
                .TextMatrix(i, .ColIndex("��������")) = rsTmp!P_Name & ""
                .TextMatrix(i, .ColIndex("����ǰ���½ű�")) = rsTmp!P_Ver & ""
                rsTmp.MoveNext
                i = i + 1
            Loop
            .AutoResize = True
            .AutoSize 0, .Cols - 1
            .Redraw = flexRDDirect
        End With
    Else
        With vsfAlter
            .Redraw = flexRDNone
            i = .Rows
            .Rows = .Rows + rsTmp.RecordCount
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                .TextMatrix(i, 0) = i
                .TextMatrix(i, .ColIndex("��������")) = rsTmp!P_Name & ""
                .TextMatrix(i, .ColIndex("���������½ű�")) = rsTmp!P_Ver & ""
                .TextMatrix(i, .ColIndex("����״̬")) = "������"
                rsTmp.MoveNext
                i = i + 1
            Loop
            .AutoResize = True
            .AutoSize 0, .Cols - 1
            .Redraw = flexRDDirect
        End With
    End If
    
End Sub

Private Sub lblCurCmd_Click()
    Dim strPath As String
    
    strPath = OpenFolder(Me, "ѡ��ǰ�汾ϵͳ��װĿ¼")
    If strPath = "" Then Exit Sub
    
    lblCurPath.Caption = strPath
    lblCurCmd.Left = lblCurPath.Left + lblCurPath.Width + 60
    
    LoadUpdateSystem strPath
End Sub

Private Sub lblTargetCmd_Click()
    Dim strPath As String
    
    strPath = OpenFolder(Me, "ѡ��Ŀ��汾ϵͳ��װĿ¼")
    If strPath = "" Then Exit Sub
    
    lblTargetPath.Caption = strPath
    lblTargetCmd.Left = lblTargetPath.Left + lblTargetPath.Width + 60
    
    LoadUpdateSystem strPath
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = "��������ƻ��޸��˺󰴻س����ж�λ" Then
        txtFind.Text = ""
        txtFind.ForeColor = ��ɫ
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If txtFind.Text = "" Then Exit Sub
    If KeyAscii <> 13 Then Exit Sub
    
    GetRowPos vsfProc, txtFind.Text, "��������,�޸���"
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "��������ƻ��޸��˺󰴻س����ж�λ"
        txtFind.ForeColor = ��ɫ
    End If
End Sub

Private Sub vsfAlter_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '����ѡ��
    Dim strProc As String, i As Long
    
    On Error Resume Next

    With vsfAlter
        If .Rows = 1 Then Exit Sub
        If .Redraw = flexRDNone Then Exit Sub
        
        .Cell(flexcpForeColor, OldRow, 0) = Color.���ɫ
        .Cell(flexcpFontBold, OldRow, 0) = False
        .Cell(flexcpFontBold, NewRow, 0) = True
        .Cell(flexcpForeColor, NewRow, 0) = Color.��ɫ
        
        If mblnChanged Then '��ֹ�ظ����û����¼�
            mblnChanged = False
            Exit Sub
        End If
        mblnChanged = True
        strProc = .TextMatrix(NewRow, .ColIndex("��������"))

    End With
    
    With vsfProc
        If .Rows = 1 Then Exit Sub
        If .Redraw = flexRDNone Then Exit Sub
        For i = 1 To .Rows - 1
            If strProc = .TextMatrix(i, .ColIndex("��������")) Then
                .Select i, 0
                .TopRow = i - (vsfAlter.Row - vsfAlter.TopRow)
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub vsfSp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If vsfSp.Visible = True Then vsfSp.Visible = False
    End If
End Sub
Private Sub vsfModule_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If vsfModule.Visible = True Then vsfModule.Visible = False
    End If
End Sub

Private Sub vsfModule_LostFocus()
    If vsfModule.Visible Then
        lblVisable_Click
    End If
End Sub

Private Sub vsfProc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '����ѡ��
    Dim strProc As String, i As Long
    
    On Error Resume Next
    With vsfProc
    
        If .Rows = 1 Then Exit Sub
        If .Redraw = flexRDNone Then Exit Sub
        
        .Cell(flexcpForeColor, OldRow, 0) = Color.���ɫ
        .Cell(flexcpFontBold, OldRow, 0) = False
        .Cell(flexcpFontBold, NewRow, 0) = True
        .Cell(flexcpForeColor, NewRow, 0) = Color.��ɫ
    
        If mblnChanged Then
            mblnChanged = False
            Exit Sub
        End If
    
        mblnChanged = True
        strProc = .TextMatrix(NewRow, .ColIndex("��������"))
    End With
    
    With vsfAlter
        If .Rows = 1 Then Exit Sub
        If .Redraw = flexRDNone Then Exit Sub
        
        For i = 1 To .Rows - 1
            If strProc = .TextMatrix(i, .ColIndex("��������")) Then
                .Select i, 0
                .TopRow = i - (vsfProc.Row - vsfProc.TopRow)
                Exit Sub
            End If
        Next
        
        If i = .Rows - 1 And strProc <> .TextMatrix(i, .ColIndex("��������")) Then
            .Select 0, 0
        End If
    End With
    
End Sub

Private Sub vsfModule_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then
        Cancel = True
    End If
    
    'û�������ű�,����ѡ��
    With vsfModule
        If .TextMatrix(Row, .ColIndex("��ǰ�汾��")) = "" Or .TextMatrix(Row, .ColIndex("Ŀ��汾��")) = "" Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfModule_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsfModule
        If .Redraw = flexRDNone Then Exit Sub
        If .Rows = 1 Then Exit Sub
        
        'ȫѡ
        If Row = 0 Then
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                    If .TextMatrix(i, .ColIndex("��ǰ�汾��")) <> "" And .TextMatrix(i, .ColIndex("Ŀ��汾��")) <> "" Then
                        .Cell(flexcpChecked, i, 0) = flexChecked
                    End If
                Else
                    .Cell(flexcpChecked, 1, 0, .Rows - 1, 0) = flexUnchecked
                End If
            Next
        End If
    End With
End Sub


