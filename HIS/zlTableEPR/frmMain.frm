VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "���ʽ�����༭��"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11775
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMain"
   ScaleHeight     =   7440
   ScaleWidth      =   11775
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picAtt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8970
      Left            =   420
      ScaleHeight     =   8970
      ScaleWidth      =   3165
      TabIndex        =   13
      Top             =   1350
      Visible         =   0   'False
      Width           =   3165
      Begin VB.CheckBox cmdAvg 
         Caption         =   "��ֵ"
         Height          =   330
         Left            =   600
         MouseIcon       =   "frmMain.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   5665
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CheckBox cmdSum 
         Caption         =   "�ϼ�"
         Height          =   330
         Left            =   105
         MouseIcon       =   "frmMain.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5665
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   0
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   34
         Text            =   "frmMain.frx":1A5E
         Top             =   6105
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox txtSum 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "�ϼ�Ӧ��"
         Height          =   350
         Left            =   1875
         TabIndex        =   32
         Top             =   5655
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   1
         Left            =   165
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "frmMain.frx":1B69
         Top             =   6210
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "frmMain.frx":1C82
         Top             =   6315
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   3
         Left            =   345
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Text            =   "frmMain.frx":1CD1
         Top             =   6390
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   4
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   28
         Text            =   "frmMain.frx":1D52
         Top             =   6510
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   5
         Left            =   615
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   27
         Text            =   "frmMain.frx":1DB9
         Top             =   6615
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   6
         Left            =   750
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "frmMain.frx":1E10
         Top             =   6705
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   7
         Left            =   855
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "frmMain.frx":1E69
         Top             =   6825
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1800
         Index           =   8
         Left            =   1005
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "frmMain.frx":1EA8
         Top             =   6960
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.Frame fraType 
         Caption         =   "��Ԫ������"
         Height          =   1980
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   2775
         Begin VB.CheckBox chkType 
            Caption         =   "�п�ǩ��"
            Height          =   350
            Index           =   8
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1230
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "�п�ǩ��"
            Height          =   350
            Index           =   7
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "ǩ��λ"
            Height          =   350
            Index           =   6
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   570
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "����ͼ"
            Height          =   350
            Index           =   5
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "�ο�ͼ"
            Height          =   350
            Index           =   4
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1560
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "��ϱ༭"
            Height          =   350
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1230
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "��Ҫ��"
            Height          =   350
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "�����ı�"
            Height          =   350
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   570
            Width           =   1125
         End
         Begin VB.CheckBox chkType 
            Caption         =   "�̶��ı�"
            Height          =   350
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Shape shpTxtSum 
         BorderColor     =   &H00E09060&
         Height          =   255
         Left            =   3300
         Top             =   4035
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.Timer timeTmp 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3165
      Tag             =   "���ڸı��и��п�������¼"
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar Processing 
      Height          =   270
      Left            =   1875
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox picHistory 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   345
      ScaleHeight     =   1800
      ScaleWidth      =   2625
      TabIndex        =   5
      Top             =   465
      Width           =   2625
      Begin VSFlex8Ctl.VSFlexGrid vsHistory 
         Height          =   1275
         Left            =   165
         TabIndex        =   6
         Top             =   105
         Width           =   2070
         _cx             =   3651
         _cy             =   2249
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
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
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
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
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
   Begin VB.PictureBox picMainBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5385
      Left            =   3780
      ScaleHeight     =   5385
      ScaleWidth      =   8415
      TabIndex        =   2
      Top             =   180
      Width           =   8415
      Begin zlTableEPR.Document Doc 
         Height          =   885
         Left            =   2775
         TabIndex        =   12
         Top             =   4260
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1561
         Border          =   0   'False
      End
      Begin VB.PictureBox PicDy 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   885
         Index           =   0
         Left            =   1425
         ScaleHeight     =   885
         ScaleWidth      =   1170
         TabIndex        =   10
         Top             =   4260
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.PictureBox picRulerV 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3825
         Left            =   0
         ScaleHeight     =   3825
         ScaleWidth      =   225
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Y����"
         Top             =   225
         Width           =   225
      End
      Begin VB.PictureBox picRulerH 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   5340
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "X����"
         Top             =   0
         Width           =   5340
      End
      Begin TTF160Ctl.F1Book F1Main 
         Height          =   3825
         Left            =   285
         TabIndex        =   0
         Top             =   240
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   6747
         _0              =   $"frmMain.frx":1EE7
         _1              =   $"frmMain.frx":22F0
         _2              =   $"frmMain.frx":26F9
         _3              =   $"frmMain.frx":2B03
         _4              =   $"frmMain.frx":2F0C
         _count          =   5
         _ver            =   2
      End
      Begin zlTableEPR.ElementEdit elEdit 
         Height          =   2055
         Left            =   5415
         TabIndex        =   9
         Top             =   570
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   3625
      End
      Begin zlTableEPR.PictureEdit PicEdit 
         Height          =   1860
         Left            =   5460
         TabIndex        =   11
         Top             =   630
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   3281
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7080
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":320C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   "msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19403
            MinWidth        =   19403
            Text            =   "����:������� סԺ��:12345678 ����:12345 ����:12�� �Ա�:δ֪ ҽ����:1234567890123"
            TextSave        =   "����:������� סԺ��:12345678 ����:12345 ����:12�� �Ա�:δ֪ ҽ����:1234567890123"
            Key             =   "PatInfo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   794
            MinWidth        =   88
            Text            =   "Ins"
            TextSave        =   "Ins"
            Key             =   "Insert"
            Object.ToolTipText     =   "������Ƿ���"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin zlTableEPR.ColorPicker ColorForeColor 
      Height          =   2190
      Left            =   2205
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   -45
      Top             =   1755
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A9E
            Key             =   "HIGHLIGHT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C13
            Key             =   "FORECOLOR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D60
            Key             =   "FILLCOLOR"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMain.frx":3ED4
      Left            =   480
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type SelRange        'F1Bookѡ������,��ʼ�С���,��ֹ�С���
    lsRow As Long
    lsCol As Long
    leRow As Long
    leCol As Long
End Type
Private Const conPane_SentenceList = 500
Private Const conPane_Attribute = 501
Private Const conPane_History = 502
Private Const conPane_Content = 503
Private Const conPane_PacsPic = 504
Public Document As cTableEPR, editType As Integer, EditMode As Integer
Private DocOld As Collection
'����
Private WithEvents mfrmSentence As frmSentenceList
Attribute mfrmSentence.VB_VarHelpID = -1
Private WithEvents mfrmMainError As frmMainMsg
Attribute mfrmMainError.VB_VarHelpID = -1
Private WithEvents mfrmPacsPic As frmPACSImg
Attribute mfrmPacsPic.VB_VarHelpID = -1
Private WithEvents mfrmEPRModelSaveAs As frmEPRModelSaveAs
Attribute mfrmEPRModelSaveAs.VB_VarHelpID = -1
Private mfrmTipInfo As New frmTipInfo

'���ڱ���
Private SelCell As New cTabCell
Private Undo As New cTabUndos
Private mReadOnly As Byte, mstrSex As String '�Ա�
Private mfrmParent As Object, mstrPrivs As String, mstrModelPrivate As String, mblnCanPrint As Boolean, mblnMoved As Boolean
Private mblnInit As Boolean '��ʼ��������
Private mblnChangeRC As Boolean '�ı��и��п�
Private mblnClickZ As Boolean '�ڵ�0�л��0����갴��
Private mblnShowAtt As Boolean, mblnAdd As Boolean '��ʾ�����С�׷��
Private mblnEditing As Boolean '�ı���̶��ı����ڱ༭״̬
Private mbFunType As Byte '�������������������ָ����������Դ =0 Sum =1 Avg

Private Sub AddUndo(TmpCell As cTabCell)
    If TmpCell.Key = "" Then Exit Sub
    Undo.Add Undo.Count & "_" & TmpCell.Key
    With Undo(Undo.Count)
        .Key = TmpCell.Key
        .CT = TmpCell.��������
        .CTxt = TmpCell.�����ı�
        .Ekey = TmpCell.ElementKey
        .Tkey = TmpCell.TextKey
        .PKey = TmpCell.PictureKey
        If Len(.PKey) <> 0 Then
            Set .OrigPic = Document.Pictures("K" & .PKey).OrigPic
        End If
        .PmKey = TmpCell.PicMarkKey
    End With
End Sub

Public Sub ShowMe(ByVal frmParent As Object, DocTab As cTableEPR, ByVal strModelPrivate As String, ByVal blnMoved As Boolean, Optional blnCanPrint As Boolean = True, Optional ByVal intStyle As Integer)
'## ������  frmParent       :������
'##         Doc             :�ⲿ���򴴽���cTableEPR��,�����ĵ��и�������
'##         strModelPrivate:����ģ��ӵ��Ȩ��
'##         blnMoved        :��ǰ�����Ƿ�ת��
'##         blnCanPrint     :�Ƿ�����Ԥ������ӡ
'################################################################################################################
Dim bfrmMode As Byte
    '���ô�����ʾ״̬
    mblnInit = False: mblnChangeRC = False: mblnClickZ = False: mblnShowAtt = False: mblnAdd = False: mblnEditing = False: mbFunType = 0
    On Error GoTo errHand
    Set mfrmParent = frmParent
    mstrModelPrivate = strModelPrivate
    mstrPrivs = GetPrivFunc(glngSys, 1070)
    mblnCanPrint = blnCanPrint
    mblnMoved = blnMoved
    mblnInit = True
    Set Document = DocTab              '����ֵ
    editType = Document.ET: EditMode = Document.EM
    
    Call OpenDoc(False)                   '���ݱ༭ģʽ���ĵ��������ĵ�
    Call DockPaneState              '����б�״̬
    Call RefreshPatiInfo            'ˢ�²�����Ϣ��
    If editType = TabET_��������� Then Call RereshHistory             'ˢ����ʷ�汾
    If Document.EPRFileInfo.���� = Tab���Ʊ��� Then zlRefreshPacsPic
    Me.Caption = IIf(mReadOnly = 2, "������ʷ�汾----", "") & Document.EPRFileInfo.����
    mblnInit = False: zlCommFun.StopFlash
    If Not Me.Visible Then
       'bfrmMode = 0
        If editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭ Then intStyle = 1
       'If mblnMoved = 2 Then bfrmMode = 1
        Me.Show intStyle, mfrmParent           '��ʾ�༭��,�鿴ģʽ��ģ̬������ʾ
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
    Unload Me
    Call SaveErrLog
End Sub
Private Function SaveDoc(Optional blnSign As Boolean, Optional blnExit As Boolean) As Boolean
'���ܣ���������
'������blnSign��ʾǩ������
Dim arrSQL As Variant, i As Integer, blnTran As Boolean, SignCellKey As String
'1 ���п�ǩ��,�п�ǩ��,��ǩ�� ���д���
'2 ���и�,�п���,����������У��
'2 ���޸�����¶���仯,���ʱ,����ǰ��ÿ����Ԫ��������֤,�䶯������������ֹ��SQL,Ȼ��ID���,��ʼ��=��ֹ��,��ֹ��=0
'3 ����Document.Save������ñ�������SQL,ʧ��ֱ���˳�
'4 �Կ�������ִ��SQL,ͬʱ��ʾ������
'5 �ٴ�ˢ�½���

    On Error GoTo errHand
1    mblnInit = True
2    arrSQL = Array()
3    If blnSign Then
4        If frmSign.Visible Then Exit Function
5    End If
    
    
6    Processing.Value = 0: Processing.Visible = True: Processing.Max = 200
    

7    stbThis.Panels("msg").Text = "��ʼ�������������--------"
8    If Doc.Visible Or mblnEditing Then F1Main_GotFocus '������ڱ༭״̬����Ҫ�����ݸ�������
    
9    If Not ValiCellDate(Not mblnAdd) Then Processing.Visible = False: GoTo lOut
10    Processing.Value = Processing.Value + 10
    
11    stbThis.Panels("msg").Text = "��ʼ���������--------"
12    If editType = TabET_��������� Then Call CompareChange(arrSQL)
13    Processing.Value = Processing.Value + 10
    
14    If blnSign Then
15        If Not AddSign(arrSQL, SignCellKey) Then Processing.Visible = False: GoTo lOut 'ǩ��
16    End If
    
17    zlCommFun.ShowFlash "���ڱ������ݣ����Եȣ�", Me
18    stbThis.Panels("msg").Text = "��ʼ�������ݱ���SQL--------"
19    If Not Document.SaveDoc(arrSQL) Then
20        Processing.Visible = False
21        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���棺" & vbCrLf & "      �������,���ݽ����ᱻ��¼", True, 0
22        stbThis.Panels("msg").Text = "�������,���ݽ����ᱻ��¼"
23        GoTo lOut
24    End If
25    Processing.Value = Processing.Value + 10

26    Err.Clear
27    gcnOracle.BeginTrans '--------------------------д������
28    stbThis.Panels("msg").Text = "��ʼ�ύ����--------"
29    Processing.Max = Processing.Value + UBound(arrSQL) + 1: blnTran = True
30    For i = 0 To UBound(arrSQL)
31        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "д������")
32        Processing.Value = Processing.Value + 1
33    Next
34    stbThis.Panels("msg").Text = "�������"
35    Call gcnOracle.CommitTrans: blnTran = False: Processing.Visible = False
    
36    If (editType = TabET_�������༭ Or editType = TabET_���������) And Not blnExit Then
37        If EditMode = TabEm_���� Then Document.EM = TabEm_�޸�: EditMode = TabEm_�޸�
38        Call Document.EPRPatiRecInfo.GetPatiRecordInfo(Document.EPRPatiRecInfo.ID, mblnMoved) '�ض����Ӳ�����¼
39    End If
    
40    SaveDoc = True: Call RelateFeedback(True)

lOut:   On Error Resume Next
41        mfrmParent.RefreshList
42        Call mfrmParent.Event_Saved(Document.EPRPatiRecInfo.ID) '���Ƶ�����Ҫ����Ϊ�����Ƿ�ģ̬��ʽ���ã��������¼���ʽ
43        Err.Clear
44        mblnInit = False
45        zlCommFun.StopFlash
46        Exit Function
errHand:
    Call MsgBox("SaveDoc�����У�" & Erl(), vbInformation, gstrSysName)
    If ErrCenter = 1 Then
        Resume
    End If
    mblnInit = False: zlCommFun.StopFlash
    If blnTran Then gcnOracle.RollbackTrans
    
    Call SaveErrLog
    mReadOnly = 2 '��Ϊ�鿴ģʽ���������ٴα��棬��Ҫ�˳�����������
    stbThis.Panels("msg").Text = "�������,���ݽ����ᱻ��¼"
    If blnSign Then
        F1Main.TextRC(Document.Cells(SignCellKey).Row, Document.Cells(SignCellKey).Col) = "[ǩ��λ]"
    End If
    Processing.Visible = False
    MsgBox "�������,���ݽ����ᱻ��¼��", vbExclamation, gstrSysName
End Function
Private Sub OpenDoc(Optional ByVal blnNew As Boolean)
'���ܣ���ȡ�ļ��ṹ���ļ����ݣ�ˢ���ĵ�����
'˵��:blnNew ��ʾ�����ϵ��½�
    If blnNew Then mblnInit = True
    
    If blnNew Then
        Document.EM = TabEm_����
        ClearPicture
    End If
    If blnNew And editType = TabET_�����ļ����� Then   '�ļ�����ʱ���½�
        Document.InitEmptyStructure         '��ʼ��һ�����ĵ�
    Else
        If Document.ReadFileStructure Then   '��ȡ�ļ��ṹ
            Document.ReadFileContent mblnMoved  '��ȡ�ļ�����
        Else
            Document.InitEmptyStructure         '��ʼ��һ�����ĵ�
        End If
    End If
    mReadOnly = Document.mReadOnly  '��OpenDoc�����п��ܸı�0-����,1-ǩ������޸�,2-������򿪲��Ļ��������ǩ���汾
    If editType = TabET_��������� Then Set DocOld = New Collection  '��˴�ʱ�ȱ���ԭʼ��¼,�����ڱ���ʱ���жԱ�
    Call RefreshF1Main: mblnInit = False                        '���������
End Sub
Private Sub DockPaneState()
Dim PaneHistory As Pane, PaneSentenceList As Pane, PaneAttribute As Pane, PanePacsPic As Pane, PaneContent As Pane

    On Error GoTo errHand
    
    Set PaneHistory = dkpMain.FindPane(conPane_History)
    Set PaneSentenceList = dkpMain.FindPane(conPane_SentenceList)
    Set PaneAttribute = dkpMain.FindPane(conPane_Attribute)
    Set PanePacsPic = dkpMain.FindPane(conPane_PacsPic)
    Set PaneContent = dkpMain.FindPane(conPane_Content)
    
    If Not PaneSentenceList Is Nothing Then
        dkpMain_AttachPane PaneSentenceList
    End If
    If Not PaneAttribute Is Nothing Then
        dkpMain_AttachPane PaneAttribute
    End If
    If Not PaneHistory Is Nothing Then
        dkpMain_AttachPane PaneHistory
    End If
    If Not PanePacsPic Is Nothing Then
        dkpMain_AttachPane PanePacsPic
    End If
    
    Select Case editType
        Case TabET_�����ļ�����, TabET_ȫ��ʾ���༭
            If Not PaneHistory Is Nothing Then
                PaneHistory.Close
                PanePacsPic.Close
            End If

            dkpMain.ShowPane conPane_Attribute
        Case TabET_�������༭, TabET_���������
            If Not PaneAttribute Is Nothing Then
                PaneAttribute.Close
            End If
            dkpMain.ShowPane conPane_SentenceList
            
            If Document.EPRFileInfo.���� <> Tab���Ʊ��� Then
                If Not PanePacsPic Is Nothing Then PanePacsPic.Close
            End If
            
            If mReadOnly = 2 Then
                PaneSentenceList.Close
                PanePacsPic.Close
                PaneHistory.Close
            End If
            
            If editType = TabET_�������༭ Then PaneHistory.Close
    End Select
    
    If Not PaneContent Is Nothing Then PaneContent.Selected = True
    PostMessage Processing.hWnd, PBM_SETBARCOLOR, 0, &H80FF80
    PostMessage Processing.hWnd, PBM_SETBKCOLOR, 0, vbWhite
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub RefreshPatiInfo()
    If editType <> TabET_�������༭ And editType <> TabET_��������� Then
        stbThis.Panels("PatInfo").Text = "������"
        If stbThis.Panels("PatInfo").Width > 3800 Then
            stbThis.Panels("PatInfo").Width = stbThis.Panels("PatInfo").Width / 3
        End If
        Exit Sub
    End If

    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    If Document.EPRPatiRecInfo.Ӥ�� <> 0 Then
        gstrSQL = "Select '����:' ||  nvl(B.Ӥ������,A.���� || '֮Ӥ' || B.���) || Decode([2], 2, '  ĸ��סԺ��:' || A.סԺ�� || '  ĸ�״���:' || A.��ǰ����, '  ĸ�������:' || A.�����) ||" & vbNewLine & _
                "        '  ����:' || To_Char(B.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') || '  �Ա�:' || Nvl(B.Ӥ���Ա�,'δ֪') || '  ҽ����' || A.ҽ���� As ��Ϣ, Nvl(B.Ӥ���Ա�,'δ֪') �Ա�" & vbNewLine & _
                "From ������Ϣ A, ������������¼ B" & vbNewLine & _
                "Where A.����id = [1] And A.����id = B.����id And B.��ҳid = [3] And B.��� = [4]"
    Else
        gstrSQL = "Select '����:' || ���� || Decode([2], 2, '  סԺ��:' || סԺ�� || '  ����:' || ��ǰ����, '  �����:' || �����) || '  �Ա�:' || �Ա� || '  ����:' || ���� ||" & vbNewLine & _
                "        '  ҽ����' || ҽ���� As ��Ϣ, �Ա�" & vbNewLine & _
                "From ������Ϣ" & vbNewLine & _
                "Where ����id = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.������Դ, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.Ӥ��)
    If rsTemp.EOF Then
        stbThis.Panels("PatInfo").Text = "������Ϣ��": mstrSex = "δ֪"
    Else
        stbThis.Panels("PatInfo").Text = rsTemp!��Ϣ
    End If
    

    If Me.Document.EPRPatiRecInfo.ҽ��id = 0 Then
        Select Case Me.Document.EPRPatiRecInfo.������Դ
        Case TabPF_����
            gstrSQL = "Select r.���� From ���˹Һż�¼ r Where r.Id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.Document.EPRPatiRecInfo.��ҳID)
            If Not rsTemp.EOF Then
                stbThis.Panels("PatInfo").Text = stbThis.Panels("PatInfo").Text & "  ����:" & IIf(NVL(rsTemp!����, 0) = 1, "��", "")
            End If
        Case TabPF_סԺ
            gstrSQL = "Select ��Ժ����, ��Ժ����, ��Ժ��ʽ From ������ҳ Where ����id = [1] And Nvl(��ҳid, 0) = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.��ҳID)
            If Not rsTemp.EOF Then
                If IsNull(rsTemp!��Ժ����) Then
                    stbThis.Panels("PatInfo").Text = stbThis.Panels("PatInfo").Text & "  ����:" & NVL(rsTemp!��Ժ����)
                Else
                    stbThis.Panels("PatInfo").Text = stbThis.Panels("PatInfo").Text & "  ����:" & NVL(rsTemp!��Ժ��ʽ) & "(��Ժ)"
                End If
            End If
        End Select
    Else
        gstrSQL = "Select ������־ From ����ҽ����¼ Where Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.Document.EPRPatiRecInfo.ҽ��id)
        If Not rsTemp.EOF Then
            stbThis.Panels("PatInfo").Text = stbThis.Panels("PatInfo").Text & "  ����:" & IIf(NVL(rsTemp!������־, 0) = 1, "��", "")
        End If
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub RereshHistory()
Dim rsTemp As ADODB.Recordset
On Error GoTo errHand
    With vsHistory
        .Clear: .Rows = 2: .Cols = 4
        .ColWidth(0) = 1000: .ColWidth(1) = 2400: .ColWidth(2) = 1000: .ColWidth(3) = 600
        .TextMatrix(0, 0) = "ǩ����": .TextMatrix(0, 1) = "ǩ��ʱ��": .TextMatrix(0, 2) = "ǩ������": .TextMatrix(0, 3) = "�汾"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    If mReadOnly = 2 Then Exit Sub '����ʷ�汾ʱ������ʾ��ʷ�汾
    
    gstrSQL = "Select Ҫ�ر�ʾ, �����ı�, ��������, ��ֹ��" & vbNewLine & _
                "From ���Ӳ�������" & vbNewLine & _
                "Where �ļ�id = [1] And �������� In (6, 7, 8) And Nvl(��ֹ��, 0)>0 and Nvl(��ֹ��, 0)<=[2]" & vbNewLine & _
                "Order By ��ֹ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Document.EPRPatiRecInfo.ID, Document.EPRPatiRecInfo.���汾)
    If rsTemp.EOF Then
        On Error Resume Next
        dkpMain.FindPane(conPane_History).Close
        Exit Sub
    End If
    
    With vsHistory
        .Rows = rsTemp.RecordCount + 1
        Do Until rsTemp.EOF
            .RowHeight(rsTemp.AbsolutePosition) = 800
            .TextMatrix(rsTemp.AbsolutePosition, 0) = NVL(rsTemp!�����ı�)
            .TextMatrix(rsTemp.AbsolutePosition, 1) = Split(Split(rsTemp!��������, "|")(1), ";")(4)
            .TextMatrix(rsTemp.AbsolutePosition, 2) = Decode(Document.EPRPatiRecInfo.��������, 4, Decode(rsTemp!Ҫ�ر�ʾ, 3, "��ʿ��", "��ʿ"), Decode(rsTemp!Ҫ�ر�ʾ, 3, "����ҽʦ", 2, "����ҽʦ", "����ҽʦ"))
            .TextMatrix(rsTemp.AbsolutePosition, 3) = CInt(rsTemp!��ֹ��)
            rsTemp.MoveNext
        Loop
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 2) = flexAlignLeftCenter
        .Cell(flexcpFontSize, 1, 0, .Rows - 1, .Cols - 1) = 12
    End With

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub MainCommandbarDefine()
'## �˵���ʼ��
    Dim cbpPopup As CommandBarPopup                     '��ʱ����
    Dim subPopup As CommandBarPopup                     '�Ӳ˵�
    Dim objControl As CommandBarControl                 '�������ؼ�
    Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�'
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsMain
        .VisualTheme = xtpThemeOffice2003
        .StatusBar.Visible = False
        .ActiveMenuBar.Title = "�˵���"
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
        .Icons = frmPublicIcon.imgPublic.Icons
        .EnableCustomization (False)
        .Options.IconsWithShadow = True '����VisualTheme����Ч
        .Options.ToolBarAccelTips = True
        .Options.ShowExpandButtonAlways = False '��ʾ��չ��ť
        .Options.UseDisabledIcons = True
        .Options.AlwaysShowFullMenus = False '�Ƿ���ʾ���в˵�
    End With
    
'------------------------------------------------�ļ�-----------------------------------------------------
    Set cbpPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�ļ�(&F)"): cbpPopup.ID = ID_File_Menu
    With cbpPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_CLEAR, "���(&C)")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE, "����(&S)")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE_QUIT, "�����˳�(&Q)")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVEASEPRDEMO, "���Ϊ����(&M)")
        Set objControl = .Add(xtpControlButton, ID_EDIT_SAVEASPHRASE, "���Ϊ�ʾ�(&D)")
        Set objControl = .Add(xtpControlButton, ID_FILE_EXPORTTOXML, "����ΪXML�ļ�(&E)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORTFROMXML, "��XML�ļ�����(&I)")

        Set objControl = .Add(xtpControlButton, ID_FILE_PAGESETUP, "ҳ������(&U)..."): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINTPREVIEW, "��ӡԤ��(&V)")
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "��ӡ(&P)...")
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "�˳�(&X)")
    End With
    
'------------------------------------------------�༭-----------------------------------------------------
    Set cbpPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�༭(&E)"): cbpPopup.ID = ID_Edit_Menu
    With cbpPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_EDIT_UNDO, "����(&U)"): objControl.ToolTipText = "�����Ե�Ԫ��Ϊ��С��λ�����ݱ仯"
        Set objControl = .Add(xtpControlButton, ID_EDIT_REDO, "����(&R)")
        Set objControl = .Add(xtpControlButton, ID_EDIT_CUT, "����(&X)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
        Set objControl = .Add(xtpControlButton, ID_EDIT_PASTE, "ճ��(&V)")
        Set objControl = .Add(xtpControlButton, ID_EDIT_DELETE, "ɾ��(&D)")
        
        Set subPopup = .Add(xtpControlPopup, 0, "ǩ�����޶�(&S)"): subPopup.ID = ID_SIGN: subPopup.BeginGroup = True
        Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_SIGN_QUIT, "ǩ��(&S)")
        Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_UNTREAD, "����(&C)")
    End With

'------------------------------------------------����-----------------------------------------------------
    Set cbpPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "����(&I)"): cbpPopup.ID = ID_Insert_Menu
    With cbpPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_INSERT_DATETIME, "���ں�ʱ��(&D)...")
        Set objControl = .Add(xtpControlButton, ID_INSERT_DATE, "��������")
        Set objControl = .Add(xtpControlButton, ID_INSERT_TIME, "����ʱ��")
        Set objControl = .Add(xtpControlButton, ID_INSERT_SPECIALCHAR, "�������(&S)...")
        Set objControl = .Add(xtpControlButton, ID_INSERT_PICTURE, "ͼƬ(&P)")
        Set objControl = .Add(xtpControlButton, ID_INSERT_ELEMENT, "Ҫ��(&E)")
        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORT, "��ʷ�ļ�(&H)...")
        Set objControl = .Add(xtpControlButton, ID_INSERT_EPRDEMO, "���뷶��(&F)...")
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTINHERITROW, "����̳���(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTINHERITCOL, "����̳���(&C)")
    End With

'------------------------------------------------���-----------------------------------------------------
    Set cbpPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "���(&T)"): cbpPopup.ID = ID_TABLE_INSERTTABLE
    With cbpPopup.CommandBar.Controls
        Set subPopup = .Add(xtpControlPopup, 0, "��(&R)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_FORMATROWHEIGHT, "�и�(&R)...")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_SAMEROWHEIGHT, "��ͬ�и�(&S)")
            
        Set subPopup = .Add(xtpControlPopup, 0, "��(&C)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_FORMATCOLWIDTH, "�п�(&C)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_SAMECOLWIDTH, "��ͬ�п�(&S)")
        
        Set subPopup = .Add(xtpControlPopup, 0, "����(&I)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTCOLLEFT, "��(�����)(&L)"): objControl.BeginGroup = True
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTCOLRIGHT, "��(���Ҳ�)(&T")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTROWUP, "��(���Ϸ�)(&B)"): objControl.BeginGroup = True
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTROWDOWN, "��(���·�)(&A)")

        
        Set subPopup = .Add(xtpControlPopup, 0, "ɾ��(&D)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_DELETECOL, "��(&C)"): objControl.BeginGroup = True
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_DELETEROW, "��(&R)")
        
        Set subPopup = .Add(xtpControlPopup, 0, "��ʽ(&F)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_FONT, "����(&F)...")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_BOLD, "����(&B)"): objControl.BeginGroup = True
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_ITALIC, "б��(&I)")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_UNDERLINE, "�»���")
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_PROTECT, "����(&P)"): objControl.BeginGroup = True
            Set objControl = subPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_MERGE, "�ϲ�(&M)")
            
        Set subPopup = .Add(xtpControlSplitButtonPopup, ID_FORMAT_FORECOLOR, "������ɫ")
            Set objCustControl = subPopup.CommandBar.Controls.Add(xtpControlCustom, 0, ""): objCustControl.Handle = ColorForeColor.hWnd
            
        Set subPopup = .Add(xtpControlPopup, ID_TABLE_CELLALIGNMENT, "���뷽ʽ(&A)")
            subPopup.CommandBar.SetTearOffPopup "���뷽ʽ", ID_TABLE_CELLALIGNMENT, 100
            subPopup.CommandBar.SetPopupToolBar True
            subPopup.BeginGroup = True
            subPopup.CommandBar.Width = 70
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT1, "���������(&1)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT2, "���Ͼ���(&2)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT3, "�����Ҷ���(&3)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT4, "�в������(&4)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT5, "�в�����(&5)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT6, "�в��Ҷ���(&6)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT7, "���������(&7)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT8, "���¾���(&8)"
            subPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT9, "�����Ҷ���(&9)"
        
        Set objControl = .Add(xtpControlButton, ID_TABLE_BORDERSTYLE, "�߿���ʽ(&B)")
    End With

'------------------------------------------------����-----------------------------------------------------
    Set cbpPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "����(&H)")
    With cbpPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "��������(&H)")
        Set cbpPopup = .Add(xtpControlPopup, 0, "&Web�ϵ�" & gstrProductName)
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_HELP_ONLINE, gstrProductName & "��ҳ(&H)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_HELP_WEBFORUM, gstrProductName & "��̳(&F)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_HELP_CONTACT, "���ͷ���(&M)")
            If App.LogMode = 0 Then Call cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_DesignTest, "���԰�Ť")
        Set objControl = .Add(xtpControlButton, ID_HELP_ABOUT, "����(&A)...")
    End With

    '## ��������ʼ��
    Call MainToolbarDefine
    '�˵�������Key��
    Call MainbarHidBinding
End Sub
Private Sub MainToolbarDefine()
Dim Bar���� As CommandBar                           '���ù�����
Dim Bar��ʽ As CommandBar                           '��ʽ������
Dim Bar��� As CommandBar                            '��񹤾���
Dim Barǩ�� As CommandBar                           'ǩ�����޶������
Dim Combo As CommandBarComboBox                     '������������ؼ�
Dim cbpPopup As CommandBarPopup                     '��ʱ����
Dim objControl As CommandBarControl                 '�������ؼ�
Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�'

    Set Bar���� = cbsMain.Add("����", xtpBarTop): Bar����.BarID = ID_Com_Bar
    With Bar����.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_CLEAR, "���")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE, "����")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE_QUIT, "�����˳�")
        
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "��ӡ"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINTPREVIEW, "��ӡԤ��")

        Set objControl = .Add(xtpControlButton, ID_EDIT_CUT, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_EDIT_COPY, "����")
        Set objControl = .Add(xtpControlButton, ID_EDIT_PASTE, "ճ��")


        Set objControl = .Add(xtpControlButton, ID_EDIT_UNDO, "����")
        Set objControl = .Add(xtpControlButton, ID_EDIT_REDO, "����")

        Set objControl = .Add(xtpControlButton, ID_INSERT_DATETIME, "����������ʱ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_INSERT_DATE, "��������")
        Set objControl = .Add(xtpControlButton, ID_INSERT_TIME, "����ʱ��")
        Set objControl = .Add(xtpControlButton, ID_INSERT_SPECIALCHAR, "�����������")

        Set objControl = .Add(xtpControlButton, ID_INSERT_PICTURE, "����ͼ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_INSERT_ELEMENT, "����Ҫ��")

        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORT, "��ʷ�ļ�"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_INSERT_EPRDEMO, "���뷶��")
    End With
    
    Set Barǩ�� = cbsMain.Add("ǩ��", xtpBarTop): Barǩ��.BarID = ID_Sign_Bar
    With Barǩ��.Controls
        Set objControl = .Add(xtpControlButton, ID_SIGN_QUIT, "ǩ��"): objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_UNTREAD, "����"): objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "����(&H)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "�˳�(&Q)"): objControl.Style = xtpButtonIconAndCaption
    End With

    Set Bar��ʽ = cbsMain.Add("��ʽ", xtpBarTop): Bar��ʽ.BarID = ID_Format_Bar
    With Bar��ʽ.Controls
        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_FONTNAME, "����", -1, False): Combo.BeginGroup = True
        Dim FontsCol As New Collection, i As Long
        Set FontsCol = GetAllFonts
        For i = 1 To FontsCol.Count
            Combo.AddItem FontsCol.Item(i)
            If FontsCol.Item(i) = "����" Then Combo.ListIndex = i
        Next
        Combo.Width = 90: Combo.DropDownWidth = 250: Combo.DropDownListStyle = True: Combo.flags = xtpFlagRightAlign
        
        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_FONTSIZE, "����ߴ�", -1, False)
        '�ֺ��б�
        Combo.AddItem "����", 1: Combo.AddItem "С��", 2: Combo.AddItem "һ��", 3: Combo.AddItem "Сһ", 4
        Combo.AddItem "����", 5: Combo.AddItem "С��", 6: Combo.AddItem "����", 7: Combo.AddItem "С��", 8
        Combo.AddItem "�ĺ�", 9: Combo.AddItem "С��", 10: Combo.AddItem "���", 11: Combo.AddItem "С��", 12
        Combo.AddItem "����", 13: Combo.AddItem "С��", 14: Combo.AddItem "�ߺ�", 15: Combo.AddItem "�˺�", 16
        Combo.AddItem 5, 17:    Combo.AddItem 5.5, 18:      Combo.AddItem 6.5, 19:  Combo.AddItem 7.5, 20
        Combo.AddItem 8, 21:    Combo.AddItem 9, 22:        Combo.AddItem 10, 23:   Combo.AddItem 10.5, 24
        Combo.AddItem 11, 25:   Combo.AddItem 12, 26:       Combo.AddItem 14, 27:   Combo.AddItem 16, 28
        Combo.AddItem 18, 29:   Combo.AddItem 20, 30:       Combo.AddItem 22, 31:   Combo.AddItem 24, 32
        Combo.AddItem 26, 33:   Combo.AddItem 28, 34:       Combo.AddItem 36, 35:   Combo.AddItem 48, 36
        Combo.AddItem 72, 37
        Combo.ListIndex = 12: Combo.Width = 50: Combo.DropDownWidth = 80: Combo.DropDownListStyle = True

        Set objControl = .Add(xtpControlButton, ID_FORMAT_BOLD, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FORMAT_ITALIC, "б��")
        Set objControl = .Add(xtpControlButton, ID_FORMAT_UNDERLINE, "�»���")
        Set objControl = .Add(xtpControlButton, ID_FORMAT_PROTECT, "����")
        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_FORMAT_FORECOLOR, "������ɫ")
            Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
                objCustControl.Handle = ColorForeColor.hWnd
    End With

    Set Bar��� = cbsMain.Add("���", xtpBarTop): Bar���.BarID = ID_Table_Bar
    With Bar���.Controls
        Set objControl = .Add(xtpControlButton, ID_TABLE_MERGE, "�ϲ���Ԫ��")
        
        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_TABLE_CELLALIGNMENT, "���뷽ʽ")
        cbpPopup.CommandBar.SetTearOffPopup "��Ԫ����뷽ʽ", ID_TABLE_CELLALIGNMENT, 100
        cbpPopup.CommandBar.SetPopupToolBar True
        cbpPopup.CommandBar.Width = 70
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT1, "���������"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT2, "���Ͼ���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT3, "�����Ҷ���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT4, "�в������"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT5, "�в�����"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT6, "�в��Ҷ���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT7, "���������"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT8, "���¾���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT9, "�����Ҷ���"
        
        Set objControl = .Add(xtpControlButton, ID_TABLE_BORDERSTYLE, "�߿���ʽ(&B)")
        Set objControl = .Add(xtpControlButton, ID_TABLE_FORMATROWHEIGHT, "�и�")
        Set objControl = .Add(xtpControlButton, ID_TABLE_SAMEROWHEIGHT, "��ͬ�и�")
        Set objControl = .Add(xtpControlButton, ID_TABLE_FORMATCOLWIDTH, "�п�")
        Set objControl = .Add(xtpControlButton, ID_TABLE_SAMECOLWIDTH, "��ͬ�п�")
        
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTCOLLEFT, "����������"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTCOLRIGHT, "���Ҳ������")
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTROWUP, "���Ϸ�������")
        Set objControl = .Add(xtpControlButton, ID_TABLE_INSERTROWDOWN, "���·�������")
        Set objControl = .Add(xtpControlButton, ID_TABLE_DELETECOL, "ɾ����")
        Set objControl = .Add(xtpControlButton, ID_TABLE_DELETEROW, "ɾ����")
    End With

    '������λ�õ���
    If Screen.Width / Screen.TwipsPerPixelX > 1024 Then
        DockingRightOf cbsMain, Bar���, Barǩ��
        DockingRightOf cbsMain, Bar��ʽ, Bar���
        DockingRightOf cbsMain, Bar����, Bar��ʽ
    Else
        DockingRightOf cbsMain, Bar����, Barǩ��
        DockingRightOf cbsMain, Bar��ʽ, Bar���
    End If
    Bar����.EnableDocking xtpFlagHideWrap
    Barǩ��.EnableDocking xtpFlagHideWrap
    Bar��ʽ.EnableDocking xtpFlagHideWrap
    Bar���.EnableDocking xtpFlagHideWrap
End Sub
Private Sub MainbarHidBinding()

    'Ctrl�ȼ�
    cbsMain.KeyBindings.Add FCONTROL, Asc("S"), ID_FILE_SAVE                '����
    cbsMain.KeyBindings.Add FCONTROL, Asc("P"), ID_FILE_PRINT               '��ӡ
    cbsMain.KeyBindings.Add FCONTROL, Asc("Z"), ID_EDIT_UNDO                '����
    cbsMain.KeyBindings.Add FCONTROL, Asc("Y"), ID_EDIT_REDO                '����
    cbsMain.KeyBindings.Add FCONTROL, Asc("X"), ID_EDIT_CUT                 '����
    cbsMain.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY                '����
    cbsMain.KeyBindings.Add FCONTROL, Asc("V"), ID_EDIT_PASTE               'ճ��
    cbsMain.KeyBindings.Add FCONTROL, Asc("N"), ID_FILE_CLEAR               '��� �½�
    cbsMain.KeyBindings.Add FCONTROL, Asc("Q"), ID_FILE_EXIT                '�˳�
    cbsMain.KeyBindings.Add FCONTROL, Asc("D"), ID_EDIT_SAVEASPHRASE        '��Ϊ�ʾ�
    cbsMain.KeyBindings.Add FCONTROL, Asc("M"), ID_FILE_SAVEASEPRDEMO       '��Ϊ����
    cbsMain.KeyBindings.Add FCONTROL, Asc("E"), ID_FILE_EXPORTTOXML                 '����XML
    cbsMain.KeyBindings.Add FCONTROL, Asc("R"), ID_FILE_IMPORTFROMXML               '����XML
    
    '���ʱ��Ҫ��ݼ�
    cbsMain.KeyBindings.Add FCONTROL, Asc("B"), ID_FORMAT_BOLD  ' "����"
    cbsMain.KeyBindings.Add FCONTROL, Asc("I"), ID_FORMAT_ITALIC ' "б��")
    cbsMain.KeyBindings.Add FCONTROL, Asc("U"), ID_FORMAT_UNDERLINE ' "�»���")
    cbsMain.KeyBindings.Add FCONTROL, Asc("T"), ID_FORMAT_PROTECT ' "����")
    cbsMain.KeyBindings.Add FCONTROL, Asc("J"), ID_TABLE_MERGE ' "�ϲ���Ԫ��")
    
    'Ctrl+Shift�ȼ�
    cbsMain.KeyBindings.Add FCONTROL Or FSHIFT, Asc("S"), ID_FILE_SAVE_QUIT                   '�����˳�
    cbsMain.KeyBindings.Add FCONTROL Or FSHIFT, Asc("P"), ID_FILE_PAGESETUP                  'ҳ������
    'F�ȼ�
    cbsMain.KeyBindings.Add 0, VK_F1, ID_HELP_CONTENT                           '����
    cbsMain.KeyBindings.Add 0, VK_F2, ID_FILE_PRINTPREVIEW                      '��ӡԤ��
    cbsMain.KeyBindings.Add 0, VK_F4, ID_INSERT_DATETIME                        '���볤ʱ��
    cbsMain.KeyBindings.Add FCONTROL, VK_F4, ID_INSERT_DATE                     '��������
    cbsMain.KeyBindings.Add FALT, VK_F4, ID_INSERT_TIME                         '����ʱ��
    cbsMain.KeyBindings.Add FSHIFT, VK_F4, ID_INSERT_SPECIALCHAR                '���������ַ�
    cbsMain.KeyBindings.Add 0, VK_F6, ID_FILE_IMPORT                            '��ʷ�ļ�
    cbsMain.KeyBindings.Add 0, VK_F5, ID_INSERT_PICTURE                         '����ͼƬ
    cbsMain.KeyBindings.Add 0, VK_F7, ID_INSERT_ELEMENT                         '����Ҫ��
    cbsMain.KeyBindings.Add 0, VK_F9, ID_INSERT_EPRDEMO                         '���뷶��
    cbsMain.KeyBindings.Add 0, VK_F11, ID_SIGN_QUIT                             'ǩ��
    cbsMain.KeyBindings.Add FCONTROL, VK_F11, ID_UNTREAD                               '����
    cbsMain.KeyBindings.Add FCONTROL Or FSHIFT, VK_F1, ID_DesignTest
     
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    On Error GoTo errHand
    If Control.Visible = False Or Control.Enabled = False Or mblnInit Then Exit Sub
    
    Select Case Control.ID
        Case ID_FILE_CLEAR '���(&C)")
            If MsgBox("ȷʵҪʹ�����ĵ��������༭���ĵ���?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then Call OpenDoc(True)
        Case ID_FILE_SAVE '����(&S)")
            If SaveDoc Then
                Call Me.ShowMe(mfrmParent, Document, mstrModelPrivate, mblnMoved, mblnCanPrint)
            End If
        Case ID_FILE_SAVE_QUIT '�����˳�(&Q)")
            If SaveDoc(, True) Then Unload Me
        Case ID_FILE_SAVEASEPRDEMO '���Ϊ����(&D)...")
            Call SaveAsDemo
        Case ID_FILE_EXPORTTOXML '����ΪXML�ļ�(&E)")
            Call ExportXml
        Case ID_FILE_IMPORTFROMXML '��XML�ļ�����(&I)")
            Call ImportXml
        Case ID_FILE_PAGESETUP 'ҳ������(&U)...")
            Call PageSetUp
        Case ID_FILE_PRINTPREVIEW '��ӡԤ��(&V)")
            Call PrintDoc(True)
        Case ID_FILE_PRINT '��ӡ(&P)...")
            Call PrintDoc(False)
        Case ID_FILE_EXIT '�˳�(&X)")
            Unload Me
        Case ID_EDIT_UNDO '����(&U)")
            Call ExeUndo
        Case ID_EDIT_REDO '����(&R)")
        Case ID_EDIT_CUT '����(&X)")
            Call ContentMove("Cut")
        Case ID_EDIT_COPY '����(&C)")
            Call ContentMove("Copy")
        Case ID_EDIT_PASTE 'ճ��(&V)")
            Call ContentMove("Paste")
        Case ID_EDIT_DELETE 'ɾ��(&D)")
            Call F1Main_KeyDown(vbKeyDelete, 0)
        Case ID_SIGN_QUIT 'ǩ��(&S)")
            If SaveDoc(True, True) Then Unload Me
        Case ID_UNTREAD '����(&C)")
            Call RollBack
        Case ID_VIEW_HEADFOOT 'ҳüҳ��(&H)")
        Case ID_INSERT_DATETIME '���ں�ʱ��(&D)...")
            Call InsertOtherText(SelCell.Key, "����ʱ��")
        Case ID_INSERT_DATE '����
            Call InsertOtherText(SelCell.Key, "����")
        Case ID_INSERT_TIME 'ʱ��
            Call InsertOtherText(SelCell.Key, "ʱ��")
        Case ID_INSERT_SPECIALCHAR '�������(&S)...")
            Call InsertOtherText(SelCell.Key, "�������")
        Case ID_INSERT_PICTURE '����ͼƬ(&P)")
            Call InsertPicture(SelCell.Key, SelCell.PictureKey)
        Case ID_INSERT_ELEMENT '����Ҫ��(&E)")
            Call InsertElement(SelCell.Key)
        Case ID_FILE_IMPORT '��ʷ�ļ�(&H)...")
        Case ID_INSERT_EPRDEMO '���뷶��(&F)...")
            Call ImportDemo
        Case ID_TABLE_FORMATROWHEIGHT '�����и�(&R)...")
            Call SetRowCol("�и�")
        Case ID_TABLE_SAMEROWHEIGHT '��ͬ�и�(&S)")
            Call SetRowCol("��ͬ�и�")
        Case ID_TABLE_FORMATCOLWIDTH '�����п�(&C)")
            Call SetRowCol("�п�")
        Case ID_TABLE_SAMECOLWIDTH '��ͬ�п�(&S)")
            Call SetRowCol("��ͬ�п�")
        Case ID_TABLE_FORMATCELL '������Ԫ������(&E)...")
        Case ID_TABLE_INSERTCOLLEFT '������(�����)(&L)")
            Call InsertRowCol("InsertLeftCol")
        Case ID_TABLE_INSERTCOLRIGHT '������(���Ҳ�)(&T")
            Call InsertRowCol("InsertRightCol")
        Case ID_TABLE_INSERTROWUP '������(���Ϸ�)(&A)")
            Call InsertRowCol("InsertUpRow")
        Case ID_TABLE_INSERTROWDOWN '������(���·�)(&B)")
            Call InsertRowCol("InsertDnRow")
        Case ID_TABLE_INSERTINHERITROW '�������̳���(&R)")
            Call InsertInherit("Row")
        Case ID_TABLE_INSERTINHERITCOL '�������̳���(&C)")
            Call InsertInherit("Col")
        Case ID_TABLE_DELETECOL 'ɾ����(&C)")
            Call DeleteRowCol("Col")
        Case ID_TABLE_DELETEROW 'ɾ����(&R)")
            Call DeleteRowCol("Row")
        Case ID_FORMAT_FONT '����(&F)...")
            Call SetCellFont
        Case ID_FORMAT_FONTSIZE '�ֺ�
            Call SetCellFormat("�ֺ�", Control.Text)
        Case ID_FORMAT_FONTNAME '������
            Call SetCellFormat("��������", Control.Text)
        Case ID_FORMAT_BOLD '����(&B)")
            Control.Checked = Not Control.Checked
            Call SetCellFormat("����", Control.Checked)
        Case ID_FORMAT_ITALIC 'б��(&I)")
            Control.Checked = Not Control.Checked
            Call SetCellFormat("б��", Control.Checked)
        Case ID_FORMAT_UNDERLINE '�»���")
            Control.Checked = Not Control.Checked
            Call SetCellFormat("�»���", Control.Checked)
        Case ID_FORMAT_PROTECT '����(&P)")
            Control.Checked = Not Control.Checked
            Call SetCellFormat("����", Control.Checked)
        Case ID_TABLE_MERGE '�ϲ�(&M)")
            Control.Checked = Not Control.Checked
            Call SetCellFormat("�ϲ�", Control.Checked)
        Case ID_FORMAT_FORECOLOR ' "������ɫ")
            Call ColorForeColor_pOK(False)
        Case ID_TABLE_CELLALIGNMENT1 '���������"
            Call SetCellFormat("���������", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT2 '���Ͼ���"
            Call SetCellFormat("���Ͼ���", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT3 '�����Ҷ���"
            Call SetCellFormat("�����Ҷ���", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT4 '�в������"
            Call SetCellFormat("�в������", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT5 '�в�����"
            Call SetCellFormat("�в�����", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT6 '�в��Ҷ���"
            Call SetCellFormat("�в��Ҷ���", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT7 '���������"
            Call SetCellFormat("���������", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT8 '���¾���"
            Call SetCellFormat("���¾���", Control.Checked)
        Case ID_TABLE_CELLALIGNMENT9 '�����Ҷ���"
            Call SetCellFormat("�����Ҷ���", Control.Checked)
        Case ID_TABLE_BORDERSTYLE '�߿���ʽ
            Call SetCellBorder
        Case ID_HELP_CONTENT '��������(&H)")
            ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
        Case ID_HELP_ONLINE ' gstrProductName & "��ҳ(&H)")
            Call zlHomePage(Me.hWnd)
        Case ID_HELP_WEBFORUM ' gstrProductName & "��̳(&F)")
            Call zlWebForum(Me.hWnd)
        Case ID_HELP_CONTACT '���ͷ���(&M)")
            Call zlMailTo(Me.hWnd)
        Case ID_HELP_ABOUT '����(&A)...")
            ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
        Case ID_DesignTest '��ƻ������԰�Ť
            Call DesignTest
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    On Error Resume Next
    Processing.Width = stbThis.Panels("PatInfo").Width
    Processing.Left = stbThis.Panels("PatInfo").Left
    Processing.Top = stbThis.Top + 60
    Err.Clear
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mblnInit = True Then Exit Sub

    Select Case Control.ID
        Case ID_FILE_CLEAR '���(&C)")
            Control.Enabled = editType <> TabET_���������
        Case ID_FILE_SAVE '����(&S)")
            Control.Enabled = mReadOnly = 0
        Case ID_FILE_SAVE_QUIT '�����˳�(&Q)")
            Control.Enabled = mReadOnly = 0
        Case ID_FILE_SAVEASEPRDEMO '���Ϊ����(&D)...")
            Control.Visible = editType <> TabET_�����ļ����� And mReadOnly <> 2
        Case ID_EDIT_SAVEASPHRASE '��Ϊ�ʾ�
            Control.Visible = False
        Case ID_FILE_EXPORTTOXML '����ΪXML�ļ�(&E)")
        Case ID_FILE_IMPORTFROMXML '��XML�ļ�����(&I)")
            Control.Enabled = editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭ Or (editType = TabET_�������༭ And EditMode = TabEm_����)
        Case ID_FILE_PAGESETUP 'ҳ������(&U)...")
            Control.Enabled = mblnCanPrint And mReadOnly <> 2
        Case ID_FILE_PRINTPREVIEW '��ӡԤ��(&V)")
            Control.Enabled = mblnCanPrint
        Case ID_FILE_PRINT '��ӡ(&P)...")
            Control.Enabled = mblnCanPrint
        Case ID_FILE_EXIT '�˳�(&X)")
        Case ID_EDIT_UNDO '����(&U)")
            Control.Enabled = Undo.Count > 0 And mReadOnly <> 2 And (Not Doc.Visible) And (Not elEdit.Visible) And (Not PicEdit.Visible)
            If Control.Enabled Then
                Control.ToolTipText = "���� " & Undo(Undo.Count).Row & "�� " & Undo(Undo.Count).Col & "�� " & _
                        Decode(Undo(Undo.Count).CT, cprCTFixtext, "�ı�", cprCTText, "�ı�", cprCTElement, "Ҫ��", cprCTTextElement, "��ϱ༭", cprCTPicture, "�ο�ͼ", cprCTReportPic, "����ͼ") & "�仯"
            Else
                Control.ToolTipText = "�����Ե�Ԫ��Ϊ��С��λ�����ݱ仯"
            End If
        Case ID_EDIT_REDO '����(&R)")
            Control.Visible = False
        Case ID_EDIT_CUT '����(&X)")
            Control.Enabled = (SelCell.�������� = cprCTText Or SelCell.�������� = cprCTTextElement Or SelCell.�������� = cprCTFixtext)
        Case ID_EDIT_COPY '����(&C)")
            Control.Enabled = (SelCell.�������� = cprCTText Or SelCell.�������� = cprCTTextElement Or SelCell.�������� = cprCTFixtext)
        Case ID_EDIT_PASTE 'ճ��(&V)")
            Control.Enabled = (SelCell.�������� = cprCTText Or SelCell.�������� = cprCTTextElement Or SelCell.�������� = cprCTFixtext)
        Case ID_EDIT_DELETE 'ɾ��(&D)")
            Control.Enabled = (SelCell.�������� = cprCTFixtext Or SelCell.�������� = cprCTText Or (SelCell.�������� = cprCTTextElement And Doc.Visible) Or SelCell.�������� = cprCTPicture Or SelCell.�������� = cprCTReportPic)
        Case ID_SIGN_QUIT 'ǩ��(&S)")
            Control.Visible = Not (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then Control.Enabled = mReadOnly = 0
        Case ID_UNTREAD '����(&C)")
            Control.Visible = Not (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then Control.Enabled = mReadOnly < 2 '0-���� 1-ǩ������޸� 2-�༭�����򿪲���
        Case ID_INSERT_DATETIME '���ں�ʱ��(&D)...")
            Control.Enabled = (SelCell.�������� = cprCTText Or (SelCell.�������� = cprCTTextElement And Doc.Visible))
        Case ID_INSERT_DATE '����
            Control.Enabled = (SelCell.�������� = cprCTText Or (SelCell.�������� = cprCTTextElement And Doc.Visible))
        Case ID_INSERT_TIME 'ʱ��
            Control.Enabled = (SelCell.�������� = cprCTText Or (SelCell.�������� = cprCTTextElement And Doc.Visible))
        Case ID_INSERT_SPECIALCHAR '�������(&S)...")
            Control.Enabled = (SelCell.�������� = cprCTFixtext Or SelCell.�������� = cprCTText Or (SelCell.�������� = cprCTTextElement And Doc.Visible))
        Case ID_INSERT_PICTURE 'ͼƬ(&P)")
            Control.Enabled = SelCell.�������� = cprCTPicture
            If Control.Enabled Then Control.Enabled = editType <> TabET_���������
        Case ID_INSERT_ELEMENT 'Ҫ��(&E)")
            Control.Enabled = (SelCell.�������� = cprCTElement And (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)) Or (SelCell.�������� = cprCTTextElement And Doc.Visible)
        Case ID_FILE_IMPORT '��ʷ�ļ�(&H)...")
             Control.Visible = False '��ʱ��������ʱ���ٴ���
'            Control.Visible = editType = TabET_ȫ��ʾ���༭ Or (editType = TabET_�������༭ And EditMode = TabEm_����)
        Case ID_INSERT_EPRDEMO '���뷶��(&F)...")
            Control.Visible = editType = TabET_ȫ��ʾ���༭ Or (editType = TabET_�������༭ And EditMode = TabEm_����)
        Case ID_EDIT_SAVEASPHRASE '��Ϊ�ʾ�
            Control.Visible = False
        Case ID_TABLE_FORMATROWHEIGHT, ID_TABLE_SAMEROWHEIGHT, ID_TABLE_FORMATCOLWIDTH, ID_TABLE_SAMECOLWIDTH '�и�,��ͬ�и� �п� ��ͬ�п�
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
            End If
        Case ID_TABLE_INSERTCOLLEFT, ID_TABLE_INSERTCOLRIGHT, ID_TABLE_INSERTROWUP, ID_TABLE_INSERTROWDOWN '������
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
            End If
        Case ID_TABLE_DELETECOL, ID_TABLE_DELETEROW
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)  'ɾ������
                If Control.Enabled Then Control.Enabled = Not mblnClickZ
            End If
        Case ID_TABLE_INSERTINHERITROW '����̳���
            Control.Visible = editType = TabET_��������� And mReadOnly <> 2
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
            End If
        Case ID_TABLE_INSERTINHERITCOL '����̳���(&R)")
            Control.Visible = False '��ʱ��������ʱ���ٴ���
        Case ID_FORMAT_FONT '����(&F)...")
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.�������� = cprCTPicture Or SelCell.�������� = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
            End If
        Case ID_FORMAT_FONTNAME '��������
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.�������� = cprCTPicture Or SelCell.�������� = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
            End If
            If Control.Visible And Control.Enabled Then
                Control.Text = SelCell.FontName
            End If
        Case ID_FORMAT_FONTSIZE '�ֺ�
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.�������� = cprCTPicture Or SelCell.�������� = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
            End If
            If Control.Visible And Control.Enabled Then
                Control.Text = GetFontSizeChinese(SelCell.FontSize)
            End If
        Case ID_FORMAT_BOLD '����(&B)")
            If (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭) Then
                Control.Visible = True
                If Control.Visible Then
                    Control.Enabled = Not (SelCell.�������� = cprCTPicture Or SelCell.�������� = cprCTReportPic)
                    If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                    If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
                End If
                If Control.Visible And Control.Enabled Then
                    Control.Checked = SelCell.FontBold
                End If
            Else
                Control.Visible = False
                If Not Control.Parent Is Nothing Then Control.Parent.Visible = False
            End If
        Case ID_FORMAT_ITALIC 'б��(&I)")
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.�������� = cprCTPicture Or SelCell.�������� = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
            End If
            If Control.Visible And Control.Enabled Then
                Control.Checked = SelCell.FontItalic
            End If
        Case ID_FORMAT_UNDERLINE '�»���")
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.�������� = cprCTPicture Or SelCell.�������� = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
            End If
            If Control.Visible And Control.Enabled Then
                Control.Checked = SelCell.FontUnderline
            End If
        Case ID_FORMAT_PROTECT '����(&P)")
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
            End If
            If Control.Visible And Control.Enabled Then
                Control.Checked = SelCell.��������
            End If
        Case ID_TABLE_MERGE '�ϲ�(&M)")
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
                Control.Checked = SelCell.Merge
            End If
        Case ID_FORMAT_FORECOLOR '������ɫ
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Enabled = Not (SelCell.�������� = cprCTPicture Or SelCell.�������� = cprCTReportPic)
                If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
            End If
        Case ID_TABLE_CELLALIGNMENT '���뷽ʽ
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                If (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignLeft) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT1
                ElseIf (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignCenter) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT2
                ElseIf (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignRight) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT3
                ElseIf (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignLeft) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT4
                ElseIf (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignCenter) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT5
                ElseIf (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignRight) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT6
                ElseIf (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignLeft) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT7
                ElseIf (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignCenter) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT8
                ElseIf (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignRight) Then
                    Control.IconId = ID_TABLE_CELLALIGNMENT9
                End If
            End If
            Control.Enabled = True
            If Control.Enabled Then Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
            If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
        Case ID_TABLE_CELLALIGNMENT1 '���������"
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignLeft)
            End If
        Case ID_TABLE_CELLALIGNMENT2 '���Ͼ���"
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignCenter)
            End If
        Case ID_TABLE_CELLALIGNMENT3 '�����Ҷ���"
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignTop And SelCell.HAlignment = F1HAlignRight)
            End If
        Case ID_TABLE_CELLALIGNMENT4 '�в������"
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignLeft)
            End If
        Case ID_TABLE_CELLALIGNMENT5 '�в�����"
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignCenter)
            End If
        Case ID_TABLE_CELLALIGNMENT6 '�в��Ҷ���"
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignCenter And SelCell.HAlignment = F1HAlignRight)
            End If
        Case ID_TABLE_CELLALIGNMENT7 '���������"
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignLeft)
            End If
        Case ID_TABLE_CELLALIGNMENT8 '���¾���"
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignCenter)
            End If
        Case ID_TABLE_CELLALIGNMENT9 '�����Ҷ���"
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
            If Control.Visible Then
                Control.Checked = (SelCell.VAlignment = F1VAlignBottom And SelCell.HAlignment = F1HAlignRight)
            End If
        Case ID_TABLE_BORDERSTYLE '�߿���ʽ
            If (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭) Then
                Control.Visible = True
                Control.Enabled = (Not Doc.Visible) And (Not elEdit.Visible)
                If Control.Enabled Then Control.Enabled = Not mblnClickZ                                  '����ͨ�������0��/��ѡ��,ͨ�������0��/�����Կ���
            Else
                Control.Visible = False
                If Not Control.Parent Is Nothing Then Control.Parent.Visible = False
            End If
        Case ID_HELP_CONTENT '��������(&H)")
        Case ID_HELP_ONLINE ' gstrProductName & "��ҳ(&H)")
        Case ID_HELP_WEBFORUM ' gstrProductName & "��̳(&F)")
        Case ID_HELP_CONTACT '���ͷ���(&M)")
        Case ID_HELP_ABOUT '����(&A)...")
        Case ID_TABLE_INSERTTABLE      '���˵�
            Control.Visible = (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
        Case ID_SIGN        '��ʽ�ͱ������������
            Control.Visible = Not (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭)
    End Select
End Sub
Private Sub MainDockPaneDefine()
    '��ʼ���沼��
    Dim PaneSentence As Pane
    Dim PaneAttribute As Pane
    Dim PaneHistory As Pane
    Dim PaneContent As Pane
    Dim PanePacsPic As Pane
    
    With Me.dkpMain
        .SetCommandBars cbsMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
    
    Set PaneSentence = dkpMain.CreatePane(conPane_SentenceList, 200, 0, DockLeftOf, Nothing)
    PaneSentence.Title = "�ʾ�ʾ��"
    PaneSentence.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set PaneAttribute = dkpMain.CreatePane(conPane_Attribute, 200, 0, DockBottomOf, PaneSentence)
    PaneAttribute.Title = "����"
    PaneAttribute.Options = PaneNoCloseable Or PaneNoFloatable
    dkpMain.AttachPane PaneAttribute, PaneSentence
    
    Set PaneHistory = dkpMain.CreatePane(conPane_History, 200, 0, DockBottomOf, PaneSentence)
    PaneHistory.Title = "��ʷ�汾"
    PaneHistory.Options = PaneNoCloseable Or PaneNoFloatable
    dkpMain.AttachPane PaneHistory, PaneSentence
    
    Set PanePacsPic = dkpMain.CreatePane(conPane_PacsPic, 200, 0, DockBottomOf, PaneSentence)
    PanePacsPic.Title = "PACS����ͼ"
    PanePacsPic.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption
    dkpMain.AttachPane PanePacsPic, PaneSentence
    
    PaneSentence.MaxTrackSize.Width = 200:  PaneAttribute.MaxTrackSize.Width = 200
    PaneHistory.MaxTrackSize.Width = 200:   PanePacsPic.MaxTrackSize.Width = 200

    Set PaneContent = dkpMain.CreatePane(conPane_Content, 1080, 0, DockRightOf, Nothing)
    PaneContent.Title = "���ȫ��"
    PaneContent.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable
End Sub

Private Sub chkType_Click(Index As Integer)
Dim i As Integer, strAttribute As String, blnReturn As Boolean, IntOld As Integer
    If mblnShowAtt Or mblnInit Then Exit Sub
    For i = 0 To chkType.UBound '����ԭ������
        If chkType(i).Value Then IntOld = i: Exit For
    Next
    Call F1Main_GotFocus
    Call SetCellAttribute(Index, strAttribute, blnReturn)
    If Not blnReturn Then
        Index = IntOld
        MsgBox strAttribute, vbInformation, gstrSysName
        strAttribute = ""
    End If
    DoEvents
    Call ShowAttr(Index, strAttribute)
End Sub
Private Sub ShowAttr(ByVal IntType As Integer, ByVal strAttribute As String)
'��ʾ���޸����Լ�˵��
Dim i As Integer
    mblnShowAtt = True
    For i = 0 To chkType.UBound '��������
        If i = IntType Then
            chkType(i).Value = vbChecked: Text1(i).Visible = True
        Else
            chkType(i).Value = vbUnchecked: Text1(i).Visible = False
        End If
    Next

    If (IntType = 0 Or IntType = 1) Then
        cmdApply.Visible = True: txtSum.Visible = True: txtSum.Locked = False: shpTxtSum.Visible = True
'        cmdSum.Visible = True: cmdAvg.Visible = True
        If InStr(strAttribute, ";") > 0 Then '�ϼƵ�Ԫ��
            txtSum.Text = strAttribute
            txtSum.Locked = False
        ElseIf InStr(strAttribute, ",") > 0 Then 'Դ��Ԫ��,����Ƕ�׺ϼ�
            txtSum.Text = strAttribute
            txtSum.Locked = True
        Else                                '�޺ϼ����Ե�Ԫ��
            txtSum.Text = ""
            txtSum.Locked = False
        End If
    Else
        cmdApply.Visible = False: txtSum.Visible = False: shpTxtSum.Visible = False
'        cmdSum.Visible = True: cmdAvg.Visible = True
    End If
    mblnShowAtt = False
End Sub
Private Sub chkType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If chkType(Index).Enabled And (editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭) Then
        If Not SelCell Is Nothing Then
            If SelCell.�������� = cprCTPicture Or SelCell.�������� = cprCTReportPic Then '�Ա����ͼ�ұ�ѡ�е������ת������������
                chkType(Index).SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmdApply_Click()
Dim i As Integer, strAttribute As String, blnReturn As Boolean
    If txtSum.Text <> "" Then
        If UBound(Split(txtSum.Text, ";")) < 1 Then 'ȷ��������һ�����ϵ�Ԫ�����
            mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      ��ʽ����ȷ����ʽ������·���˵������" & vbCrLf & "      ���飡", True, 1
            Exit Sub
        End If
        
        For i = 0 To UBound(Split(txtSum.Text, ";"))
            If UBound(Split(Split(txtSum.Text, ";")(i), ",")) <> 1 Then 'ȷ����ɺϼƵĵ�Ԫ����Ч
                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      ��ʽ����ȷ����ʽ������·���˵������" & vbCrLf & "      ���飡", True, 1
                Exit Sub
            End If
        Next
    End If
    strAttribute = Trim(txtSum.Text) '�����ǿ�,��ʾȡ����Ԫ��ĺϼ�����
    If Not SetSumAtt(strAttribute) Then
        MsgBox strAttribute, vbInformation, gstrSysName
        txtSum.SelStart = 0: txtSum.SelLength = Len(txtSum): txtSum.SetFocus
    Else
        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "��ʾ��" & vbCrLf & "      �ϼ������趨�ɹ�", True, 0
    End If
End Sub

Private Sub cmdSum_Click()
    If cmdSum.Value = vbChecked Then
        F1Main.MousePointer = vbCustom
        F1Main.MouseIcon = cmdSum.MouseIcon
        mbFunType = 0
        mfrmTipInfo.ShowTipInfo txtSum.hWnd, "��ʾ��" & vbCrLf & "      ��ָ����ǰ��Ԫ������Щ��Ԫ��ϼ���ɡ�", True, 0
    End If
End Sub

Private Sub ColorForeColor_pOK(ByVal ControlSelf As Boolean)
    Call SetCellFormat("������ɫ", ColorForeColor.Color)
    Call SetColorIcon(ID_FORMAT_FORECOLOR, ColorForeColor.Color)
    If ControlSelf Then SendKeys "{ESCAPE}"
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    Select Case Item.ID
    Case conPane_SentenceList
        Item.Handle = mfrmSentence.hWnd
    Case conPane_Attribute
        Item.Handle = picAtt.hWnd
    Case conPane_History
        Item.Handle = picHistory.hWnd
    Case conPane_PacsPic
        Item.Handle = mfrmPacsPic.hWnd
    Case conPane_Content
        Item.Handle = picMainBack.hWnd
    End Select
    Err.Clear
End Sub
Private Sub InitForms()
    Set mfrmSentence = New frmSentenceList: mfrmSentence.mstrPrivs = mstrPrivs
    Set mfrmMainError = New frmMainMsg
    Set mfrmPacsPic = New frmPACSImg
End Sub
Private Sub RefreshF1Main()
Dim lngRow As Long, lngCol As Long, lngCell As Long, vCell As F1CellFormat, lngCount As Long, strShow As String
    On Error GoTo errHand
    With F1Main
        .DeleteRange .MinRow, .MinCol, .MaxRow, .MaxCol, F1ShiftRows
        .ShowTabs = F1TabsOff
        .AllowMoveRange = False '�ƶ�ѡ������
        .AllowFillRange = False '�϶���Χ��ֵ,���¼����ɿ���
        .AllowInCellEditing = False '��Ԫ��༭
        .AllowEditHeaders = False '�༭��ͷ
        .AllowDesigner = False  '�������
        .AllowDelete = False '��ʾ��Ӣ�ĵģ���ò�Ҫ���������ͨ��KeyDown����
        .ShowLockedCellsError = False '��������Ԫ����б༭ʱ����Ϣ��ʾ
        .ScrollToLastRC = False '������������һ����Ԫ��
        .ColWidthUnits = F1ColWidthUnitsTwips '�п���㵥λΪ��
        .ShowSelections = F1On   '�����㲻�ڿؼ���ʱ��ֱ�ӵ�����Ԫ��ѡ��
        .DefaultFontName = "����"
        .DefaultFontSize = 9
        .MaxCol = Me.Document.Cells.Cols
        .MaxRow = Me.Document.Cells.Rows
        
        If editType = TabET_�������༭ Or editType = TabET_��������� Then
            .ShowColHeading = False '��ʾ�̶���
            .ShowRowHeading = False '��ʾ�̶���
            .AllowResize = True    '�϶����Զ�����ʱ �ı��и��п�
        Else
            .ShowColHeading = True
            .ShowRowHeading = True
            .HdrHeight = 300
            .HdrWidth = 300
            .AllowResize = True
        End If
        
        F1Main.SetSelection 1, 1, F1Main.MaxRow, F1Main.MaxCol
        F1Main.SetAlignment F1HAlignJustify, True, F1VAlignTop, 0
        '���и��п�
        For lngRow = 1 To .MaxRow
            .RowHeight(lngRow) = Me.Document.Cells.Cell(lngRow, 1).Height
        Next
        For lngCol = 1 To .MaxCol
            .ColWidthTwips(lngCol) = Me.Document.Cells.Cell(1, lngCol).Width
            .ColText(lngCol) = lngCol '��ͷ��ʾ����
        Next
        
        lngCount = Me.Document.Cells.Count
        If Me.Visible Then Processing.Max = lngCount
        For lngCell = 1 To lngCount
            If Me.Visible Then Processing.Value = lngCell: Processing.Visible = True
            
            lngRow = Me.Document.Cells(lngCell).Row: lngCol = Me.Document.Cells(lngCell).Col
            With Me.Document.Cells.Cell(lngRow, lngCol)
                'ָ������
                If .Merge And InStr(.MergeRange, ";") > 0 Then 'MergeRange���ݸ�ʽ (���Ϸ�)��,��;(���·�)��,��
                    F1Main.SetSelection Split(Split(.MergeRange, ";")(0), ",")(0), Split(Split(.MergeRange, ";")(0), ",")(1), Split(Split(.MergeRange, ";")(1), ",")(0), Split(Split(.MergeRange, ";")(1), ",")(1)
                Else
                    F1Main.SetSelection lngRow, lngCol, lngRow, lngCol
                End If
                Set vCell = F1Main.CreateNewCellFormat
                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then 'ֻ�кϲ���Ԫ���׸���Ǻϲ���Ԫ���ˢ��
'                    vCell.ProtectionLocked = .��������  '�Ƿ�����,��������,�п�,�п�,ǩ��������ʱд��Database
                    vCell.MergeCells = .Merge
                    vCell.WordWrap = True
                    vCell.FontName = .FontName          '����>����</����>
                    vCell.FontSize = .FontSize          '<�ֺ�>9</�ֺ�>
                    vCell.FontBold = .FontBold          '<����>False</����>
                    vCell.FontItalic = .FontItalic        '<б��>False</б��>
                    vCell.FontUnderline = .FontUnderline     '<�»���>False</�»���>
                    vCell.FontStrikeout = .FontStrikeout    '<ɾ����>False</ɾ����>
                    vCell.FontColor = .FontColor         '<������ɫ>vbblack</������ɫ>
                    vCell.AlignHorizontal = .HAlignment       '<�������>F1HAlignCenter</�������>
                    vCell.AlignVertical = .VAlignment       '<�������>F1VAlignCenter</�������>

                    Select Case .��������
                        Case cprCTFixtext    '0-�̶��ı�(���ɱ༭)
                            F1Main.TextRC(lngRow, lngCol) = .�����ı�
                        Case cprCTText '1-�ı���(�ɱ༭�����ı�)
                            F1Main.TextRC(lngRow, lngCol) = .�����ı�
                            If editType = TabET_��������� Then DocOld.Add .�����ı�, .Key
                        Case cprCTElement    '2-��Ҫ��
                            If editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭ Then
                                If .ElementKey <> "" Then
                                    If Document.Elements("K" & .ElementKey).������̬ = 1 And Document.Elements("K" & .ElementKey).Ҫ������ <> 2 Then '������̬=չ��
                                        F1Main.TextRC(lngRow, lngCol) = Document.Elements("K" & .ElementKey).�����ı�
                                    Else
                                        F1Main.TextRC(lngRow, lngCol) = "[" & Document.Elements("K" & .ElementKey).Ҫ������ & "]" & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                                    End If
                                End If
                            Else
                                strShow = ""
                                If .�����ı� = "" Then
                                    If Document.Elements("K" & .ElementKey).�滻�� = 1 Then '�Զ��滻Ҫ��
                                        strShow = GetReplaceEleValue(Document.Elements("K" & .ElementKey).Ҫ������, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.������Դ, Document.EPRPatiRecInfo.ҽ��id, Document.EPRPatiRecInfo.Ӥ��)
                                        If strShow = "" And Not Document.Elements("K" & .ElementKey).�Զ�ת�ı� Then 'ûȡ��ֵ���Ƿ��Զ�ת�����ı�(��)
                                            strShow = "[" & Document.Elements("K" & .ElementKey).Ҫ������ & "]" & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                                        Else
                                            Document.Elements("K" & .ElementKey).�����ı� = strShow
                                            .�����ı� = strShow & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                                            strShow = .�����ı�
                                        End If
                                    Else
                                        If Document.Elements("K" & .ElementKey).������̬ = 1 And Document.Elements("K" & .ElementKey).Ҫ������ <> 2 Then '������̬=չ��
                                            .�����ı� = Document.Elements("K" & .ElementKey).�����ı� & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                                            strShow = .�����ı�
                                        Else
                                            strShow = "[" & Document.Elements("K" & .ElementKey).Ҫ������ & "]" & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                                        End If
                                    End If
                                    F1Main.TextRC(lngRow, lngCol) = strShow
                                Else
                                    F1Main.TextRC(lngRow, lngCol) = .�����ı�
                                End If
                            End If
                            If editType = TabET_��������� Then DocOld.Add .�����ı�, .Key
                        Case cprCTTextElement '3-�ı����Ҫ�ػ�ϱ༭
                            GetTextELement .Key     '����Text Element��дF1Main�еĵ�Ԫ����������ı�
                            If editType = TabET_��������� Then DocOld.Add .�����ı�, .Key
                        Case cprCTReportPic, cprCTPicture    '5-����ͼ
                            If Me.Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                Call PaintPictureOnTable(.Key)
                                If editType = TabET_��������� Then DocOld.Add Document.Pictures("K" & .PictureKey).OrigPic.Handle, .Key
                            Else
                                If editType = TabET_��������� Then DocOld.Add 0, .Key
                            End If
                            F1Main.TextRC(lngRow, lngCol) = IIf(.�������� = cprCTPicture, "�ο�ͼ", "����ͼ")
                        Case cprCTSign         '6-ǩ��'ǩ�������ʱ��Ϊռλ,��ʵ����Ϣ��û��ǩ��ʱ��ֹ��=0����ͨǩ�������ʱ����ʾ���Ա��ٴ�ǩ�����п�/�п�ǩ�������ʱҪ��ʾ
                            strShow = ""
                            If editType = TabET_�������༭ Or editType = TabET_��������� Then
                                Select Case mReadOnly 'mReadOnly 0-����,1-ǩ������޸�,2-������򿪲��Ļ��������ǩ���汾
                                    Case 0
                                        If .��ֹ�� <> 0 Then
                                            With Document.Signs("K" & .SignKey)
                                                strShow = .ǰ������ & .���� & IIf(.��ʾ��ǩ, "����ǩ��_____________", "")
                                                strShow = strShow & IIf(Trim(.��ʾʱ��) = "", "", "��" & Format(.ǩ��ʱ��, .��ʾʱ��))
                                            End With
                                        Else
                                            strShow = "[ǩ��λ]"
                                        End If
                                    Case 1, 2
                                        If .��ֹ�� = 0 Then
                                            strShow = "[ǩ��λ]"
                                        Else
                                            With Document.Signs("K" & .SignKey)
                                                strShow = .ǰ������ & .���� & IIf(.��ʾ��ǩ, "����ǩ��_____________", "")
                                                strShow = strShow & IIf(Trim(.��ʾʱ��) = "", "", "��" & Format(.ǩ��ʱ��, .��ʾʱ��))
                                            End With
                                        End If
                                End Select
                            Else
                                strShow = "[ǩ��λ]"
                            End If
                            F1Main.TextRC(lngRow, lngCol) = strShow 'ǰ������ & ���� & ��ʾ��ǩ & ��ʾʱ��<>""(format(ǩ��ʱ��,��ʾʱ��)
                        Case cprCTRowSign, cprCTColSign '7-�п�ǩ�� '8-�п�ǩ��
                            strShow = ""
                            If editType = TabET_�������༭ Or editType = TabET_��������� Then
                                If .��ֹ�� <> 0 Then
                                    With Document.Signs("K" & .SignKey)
                                        strShow = .ǰ������ & .���� & IIf(.��ʾ��ǩ, "����ǩ��_____________", "")
                                        strShow = strShow & IIf(Trim(.��ʾʱ��) = "", "", "��" & Format(.ǩ��ʱ��, .��ʾʱ��))
                                    End With
                                Else
                                    strShow = "[ǩ��λ]"
                                End If
                            Else
                                strShow = "[ǩ��λ]"
                            End If
                            F1Main.TextRC(lngRow, lngCol) = strShow 'ǰ������ & ���� & ��ʾ��ǩ & ��ʾʱ��<>""(format(ǩ��ʱ��,��ʾʱ��)
                    End Select
                    F1Main.SetCellFormat vCell
                    Call F1Main.SetBorder(-1, .CellLineLeft, .CellLineRight, .CellLineTop, .CellLineBottom, 0, -1, .CellLineLeftColor, .CellLineRightColor, .CellLineTopColor, .CellLineBottomColor)
                End If
            End With
        Next
        Processing.Visible = False
    End With
    
'    F1Main.SetSelection F1Main.MaxRow, F1Main.MaxCol, F1Main.MaxRow, F1Main.MaxCol
    F1Main.SetSelection 1, 1, 1, 1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Doc_BeforeKeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long
    If KeyCode = vbKeyReturn And Shift = 0 Then 'ֱ�ӻس���ʾ�˳��༭
        KeyCode = 0
        F1Main_GotFocus
        F1Main.SetFocus
        Exit Sub
    End If
    If KeyCode = vbKeyReturn And Shift = 4 Then 'ALT+�س���ʾ����
'        SendKeys "^~"
    End If
    If Shift <> 0 Then Exit Sub
    
    Select Case KeyCode
    Case 0, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
        vbKeyEscape, vbKeyTab, vbKeyDelete, vbKeyBack, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
        vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital, _
        vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12
        Exit Sub
    End Select

    With Doc
        If .SelLength > 0 Then .Range(Doc.Selection.EndPos, .Selection.EndPos).Selected
        i = .Selection.StartPos
        If i = 0 Then
            .Range(i, i).Font.Protected = False
            .Range(i, i).Font.Hidden = False
        ElseIf .Range(i - 1, i).Font.Hidden And _
            .Range(i, i + 1).Font.Hidden = False And _
            .Range(i, i + 1).Font.Protected = False Then
            'A���⣺�������ı���|��ͨ�ı�
            .Range(i, i).Font.Protected = False
            .Range(i, i).Font.Hidden = False
        ElseIf .Range(i - 1, i).Font.Protected And .Range(i - 1, i).Font.Hidden = False And .Range(i, i + 1).Font.Hidden And .Range(i, i + 3).Text = "EE(" Then
            'B����1�������عؼ��֣�[Ҫ��]|�����عؼ��֣������عؼ��֣�[Ҫ��]�����عؼ��֣�
            If KeyCode = vbKeySpace Then KeyCode = 0
            .Range(i + 16, i + 16).Text = " "
            .Range(i + 16, i + 17).Font.Protected = False
            .Range(i + 16, i + 17).Font.Hidden = False
            .Range(i + 17, i + 17).Selected
        ElseIf .Range(i - 1, i).Font.Hidden And .Range(i, i + 1).Font.Protected And .Range(i, i + 1).Font.Hidden = False And (i - 16 <> 0) Then
            'B����2�������عؼ��֣�[Ҫ��]�����عؼ��֣������عؼ��֣�|[Ҫ��]�����عؼ��֣�
            If KeyCode = vbKeySpace Then KeyCode = 0
            .Range(i - 16, i - 16).Text = " "
            .Range(i - 16, i - 15).Font.Protected = False
            .Range(i - 16, i - 15).Font.Hidden = False
            .Range(i - 15, i - 15).Selected
        ElseIf i - 16 = 0 And .Range(i - 1, i).Font.Hidden And .Range(i - 1, i).Font.Protected And .Range(i, i + 1).Font.Hidden = False Then
            '����2��0�����عؼ��֣�|[Ҫ��]�����عؼ��֣�
            .Range(i - 16, i - 16).Font.Protected = False
            .Range(i - 16, i - 16).Font.Hidden = False
            .Range(i - 16, i - 16).Selected
        End If
    End With
End Sub



Private Sub Doc_Change()
    With Doc
        .Range(0, Len(.Text)).Font.Name = SelCell.FontName
        .Range(0, Len(.Text)).Font.Size = SelCell.FontSize
        .Range(0, Len(.Text)).Font.Italic = SelCell.FontItalic
        .Range(0, Len(.Text)).Font.Bold = SelCell.FontBold
        .Range(0, Len(.Text)).Font.ForeColor = SelCell.FontColor
        .Range(0, Len(.Text)).Font.Underline = IIf(SelCell.FontUnderline, cprHair, cprNone)
        .Range(0, Len(.Text)).Font.Strikethrough = SelCell.FontStrikeout
    End With
End Sub


Private Sub Doc_DblClick()
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
Dim pt As POINTAPI, lHheight As Long, lHwidth As Long
    pt.x = 0: pt.y = 0
    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '�̶��и߶�
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '�̶��п��
    bInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys And sType = "E" Then
        If Document.Elements("K" & lKey).������̬ = 1 Then Exit Sub
        ShowElInDoc lSE, lES, lKey
    End If
End Sub
Private Sub ShowElInDoc(ByVal lSE As Long, ByVal lES As Long, ByVal lKey As Long)
'��RichEditor�� ��ʾҪ�ر༭��
Dim pt As POINTAPI, lHheight As Long, lHwidth As Long
    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '�̶��и߶�
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '�̶��п��
    
    Doc.Range(lSE, lES).Selected
    ClientToScreen Doc.hWnd, pt
    Dim lLeft As Long, lTOp As Long
    '��ȡ��ʼλ������
    Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos + 1).GetPoint cprGPStart + cprGPLeft + cprGPBottom, lLeft, lTOp
    elEdit.SetElement Document.Elements("K" & lKey), 0, editType
    elEdit.Move F1Main.Left + Doc.Left + lLeft, F1Main.Top + Doc.Top + lTOp
    If elEdit.Top + elEdit.Height > F1Main.Top + F1Main.Height Then
        elEdit.Top = F1Main.Top + Doc.Top + lTOp - elEdit.Height - 300 - Screen.TwipsPerPixelY * 2
    End If
    If elEdit.Left + elEdit.Width > F1Main.Left + F1Main.Width Then
        elEdit.Left = F1Main.Left + Doc.Left + lLeft - elEdit.Width - Screen.TwipsPerPixelX * 2
    End If

    elEdit.Visible = True: elEdit.ZOrder 0: elEdit.SetFocus
End Sub

Private Sub Doc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
Dim i As Long
    
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Or KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        i = Doc.Selection.StartPos: If i = 0 Then i = 1
        bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded)  '���ڹؼ��ֶ�֮��
        If bInKeys Then
            Select Case KeyCode
                Case vbKeyDelete, vbKeyBack
                    If Doc.Range(i - 1, i + 3).Text Like ")?S(" And Doc.Range(i - 1, i + 3).Font.Hidden = True Then
                        '�������ı����������ı����������ı���|�������ı����������ı����������ı���
                        If KeyCode = vbKeyBack Then
                            i = i - 16
                        Else
                            i = i + 16
                        End If
                        bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
                    ElseIf Doc.Range(i - 17, i - 13).Text Like ")?S(" And Doc.Range(i - 17, i + 13).Font.Hidden = True Then
                            '�������ı����������ı����������ı����������ı���|�������ı����������ı���
                        If KeyCode = vbKeyBack Then
                            i = i - 32
                            bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
                        End If
                    ElseIf Doc.Range(i + 15, i + 19).Text Like ")?S(" And Doc.Range(i + 15, i + 19).Font.Hidden = True Then
                        '�������ı����������ı���|�������ı����������ı����������ı����������ı���
                        If KeyCode = vbKeyDelete Then
                            i = i + 32
                            bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '��ǰ���Ҫ��ĩλ��Del,Ӧɾ�����Ҫ��
                        End If
                    ElseIf Doc.Range(i - 17, i - 16).Font.Hidden = False And Doc.Range(i - 17, i - 16).Font.Protected = False Then
                        bInKeys = False
                    End If
                    If bInKeys Then
                        KeyCode = 0
                        If editType <> TabET_��������� Or Document.Elements("K" & lKey).ID = 0 Then
                            Document.Elements("K" & lKey).DeleteFromEditor Doc: Exit Sub  'ɾ��Ҫ��
                        End If
                    End If
                Case vbKeySpace, vbKeyReturn
                    KeyCode = 0
                    Call ShowElInDoc(lSE, lES, lKey): Exit Sub
            End Select
        Else
            With Doc
                If .Range(i - 1, i).Font.Hidden And KeyCode = vbKeyBack Then  '�ڹؼ��ֺ�Back
                    If i <= 1 Then
                        i = i + 15
                    Else
                        i = i - 16
                    End If
                    bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '���ڹؼ��ֶ�֮ǰ
                    If bInKeys Then
                        KeyCode = 0
                        If editType <> TabET_��������� Or Document.Elements("K" & lKey).ID = 0 Then
                            Document.Elements("K" & lKey).DeleteFromEditor Doc: Exit Sub 'ɾ��Ҫ��
                        End If
                    End If
                ElseIf .Range(i, i + 1).Font.Hidden And KeyCode = vbKeyDelete Then '�ڹؼ���ǰ��DEL
                    i = i + 16: KeyCode = 0
                    bInKeys = IsBetweenAnyKeys(Doc, i, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '���ڹؼ��ֶ�֮ǰ
                    If editType <> TabET_��������� Or Document.Elements("K" & lKey).ID = 0 Then Document.Elements("K" & lKey).DeleteFromEditor Doc: Exit Sub 'ɾ��Ҫ��
                End If
            End With
        End If
    End If
End Sub

Private Sub Doc_KeyPress(KeyAscii As Integer)
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
bInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '���ڹؼ��ֶ�֮��
If bInKeys Then KeyAscii = 0
End Sub


Private Sub Doc_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = vbKeyZ Then
        Doc.Undo
    ElseIf Shift = 2 And KeyCode = vbKeyY Then
        Doc.Redo
    End If
End Sub

Private Sub Doc_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bFinded As Boolean, bNeeded As Boolean

    bFinded = IsBetweenKeys(Doc, Doc.Selection.StartPos + 1, "E", lSS, lSE, lES, lEE, lKey, bNeeded)
    If bFinded Then '�������Ԫ���ڲ������ʾѡ��ĳ��ѡ��
        If Document.Elements("K" & lKey).������̬ = 1 Then 'չ����ʽ��Ҫ��¼��     '������
            Dim strTmp As String, p As Long, P1 As Long, P2 As Long, blnForce As Boolean, lMax As Long
            With Doc
                .Freeze
                .ForceEdit = True
                strTmp = .Range(lSE, lES).Text
                p = .Selection.StartPos
                If Document.Elements("K" & lKey).Ҫ�ر�ʾ = 2 Then
                    P1 = .Selection.StartPos - lSE + 1
                    P1 = InStrRev(strTmp, "��", P1)
                    P2 = .Selection.StartPos - lSE + 1
                    P2 = InStrRev(strTmp, "��", P2)
                    If P1 > P2 And P1 > 0 Then
                        '��ѡ
                        strTmp = Replace(strTmp, "��", "��")
                        Mid(strTmp, P1, 1) = "��"
                        .Range(lSE, lES).Text = strTmp
                    ElseIf P2 > P1 And P2 > 0 Then
                        strTmp = Replace(strTmp, "��", "��")
                        Mid(strTmp, P2, 1) = "��"
                        .Range(lSE, lES).Text = strTmp
                    End If
                    Document.Elements("K" & lKey).�����ı� = strTmp
                ElseIf Document.Elements("K" & lKey).Ҫ�ر�ʾ = 3 Then
                    P1 = .Selection.StartPos - lSE + 1
                    P1 = InStrRev(strTmp, "��", P1)
                    P2 = .Selection.StartPos - lSE + 1
                    P2 = InStrRev(strTmp, "��", P2)
                    If P1 > P2 And P1 > 0 Then
                        Mid(strTmp, P1, 1) = "��"
                        .Range(lSE, lES).Text = strTmp
                    ElseIf P2 > P1 And P2 > 0 Then
                        Mid(strTmp, P2, 1) = "��"
                        .Range(lSE, lES).Text = strTmp
                    End If
                    Document.Elements("K" & lKey).�����ı� = strTmp
                End If
                Me.Document.Elements("K" & lKey).�����ı� = strTmp
                .Range(p, p).Selected
                .UnFreeze
            End With
        Else
            
        End If
    End If
End Sub

Private Sub Doc_RequestRightMenu(ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim objPopup As CommandBar
Dim objControl As CommandBarControl

    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, ID_EDIT_CUT, "����")
        Set objControl = .Add(xtpControlButton, ID_EDIT_COPY, "����")
        Set objControl = .Add(xtpControlButton, ID_EDIT_PASTE, "ճ��")
    End With
    objPopup.ShowPopup
End Sub

Private Sub elEdit_LostFocus()
    elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0: elEdit.Tag = ""
End Sub

Private Sub elEdit_pCancel()
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    If Document.Cells.Cell(lsRow, lsCol).�������� = cprCTElement Then '��Ҫ��ȡ���༭
        If F1Main.Visible And F1Main.Enabled Then
            F1Main.SetFocus
        End If
    ElseIf Document.Cells.Cell(lsRow, lsCol).�������� = cprCTText Then '��ǰΪ�ı�����������ʱ�䣬�������
    Else
    
    End If
End Sub

Private Sub elEdit_pChange()
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    On Error GoTo errHand
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    If Document.Cells.Cell(lsRow, lsCol).�������� = cprCTElement Then '��Ҫ�ر���༭
        AddUndo Document.Cells.Cell(lsRow, lsCol)
        Document.Cells.Cell(lsRow, lsCol).�����ı� = elEdit.Element.�����ı�
        F1Main.TextRC(lsRow, lsCol) = IIf(elEdit.Element.�����ı� <> "", elEdit.Element.�����ı�, "[" & elEdit.Element.Ҫ������ & "]") & elEdit.Element.Ҫ�ص�λ
        If elEdit.Visible Then
            elEdit.SetFocus
        End If
    ElseIf Document.Cells.Cell(lsRow, lsCol).�������� = cprCTText Then '��ǰΪ�ı�����������ʱ�䣬�������
        '���ʱ������
    Else
    
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub elEdit_pOk()
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, strShow As String
    On Error GoTo errHand
    If UBound(Split(elEdit.Element.����, "|")) > 0 Then
        lsRow = Split(elEdit.Element.����, "|")(0): lsCol = Split(elEdit.Element.����, "|")(1)
    Else
        Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    End If

    If Document.Cells.Cell(lsRow, lsCol).�������� = cprCTElement Then '��Ҫ��ȷ���༭
        AddUndo Document.Cells.Cell(lsRow, lsCol)
        With elEdit.Element
            If .�滻�� = 1 Then
                If Trim(.�����ı�) = "" Then
                    If .�Զ�ת�ı� Then
                        strShow = " " & elEdit.Element.Ҫ�ص�λ
                    Else
                        strShow = "[" & elEdit.Element.Ҫ������ & "]" & elEdit.Element.Ҫ�ص�λ
                    End If
                Else
                    strShow = .�����ı�
                End If
            Else
                strShow = IIf(elEdit.Element.�����ı� <> "", elEdit.Element.�����ı�, "[" & elEdit.Element.Ҫ������ & "]") & elEdit.Element.Ҫ�ص�λ
            End If
        End With
        Document.Cells.Cell(lsRow, lsCol).�����ı� = Trim(elEdit.Element.�����ı�)
        F1Main.TextRC(lsRow, lsCol) = strShow
        elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0
        If F1Main.Enabled Then F1Main.SetFocus
    ElseIf Document.Cells.Cell(lsRow, lsCol).�������� = cprCTText Then '��ǰΪ�ı�����������ʱ�䣬�������
        AddUndo Document.Cells.Cell(lsRow, lsCol)
        strShow = Document.Cells.Cell(lsRow, lsCol).�����ı� & elEdit.Element.�����ı�
        Document.Cells.Cell(lsRow, lsCol).�����ı� = strShow
        F1Main.TextRC(lsRow, lsCol) = strShow
        elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0
    ElseIf Document.Cells.Cell(lsRow, lsCol).�������� = cprCTTextElement Then
        With elEdit.Element
            If .�滻�� = 1 Then
                If Trim(.�����ı�) = "" Then
                    If .�Զ�ת�ı� Then
                        strShow = " " & elEdit.Element.Ҫ�ص�λ
                    Else
                        strShow = "[" & elEdit.Element.Ҫ������ & "]" & elEdit.Element.Ҫ�ص�λ
                    End If
                Else
                    strShow = .�����ı�
                End If
            Else
                strShow = IIf(elEdit.Element.�����ı� <> "", elEdit.Element.�����ı�, "[" & elEdit.Element.Ҫ������ & "]") & elEdit.Element.Ҫ�ص�λ
            End If
        End With
        
        Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
        bInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '���ڹؼ��ֶ�֮��
        If bInKeys Then
            If InStr(Document.Cells.Cell(lsRow, lsCol).ElementKey, lKey) = 0 Then '�ڵ�ǰ��Ԫ���Ҫ��Key�����Ҳ�����Ҫ��Key�����ǹ����������ʱ�䣬�����ı���ʽ���룬����Ϊ˫��Ҫ�ص���
                Doc.Range(lEE, lEE).Selected
                Doc.Range(lEE, lEE).Font.Protected = False
                Doc.Range(lEE, lEE).Font.Hidden = False
                Doc.Range(lEE, lEE).Text = strShow
                Doc.Range(lEE + Len(strShow), lEE + Len(strShow)).Selected
            Else
                Doc.Range(lSE, lES).Text = strShow
            End If
        Else
            Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Selected
            Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Font.Protected = False
            Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Font.Hidden = False
            Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Text = strShow
            Doc.Range(Doc.Selection.StartPos + Len(strShow), Doc.Selection.StartPos + Len(strShow)).Selected
        End If
        elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub elEdit_TitleMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage elEdit.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub F1Main_DblClick(ByVal nRow As Long, ByVal nCol As Long)
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    If nRow = 0 Or nCol = 0 Then Exit Sub
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    Call EnterEdit(lsRow, lsCol, leRow, leCol, 0, True)
End Sub

Private Sub F1Main_EndEdit(EditString As String, Cancel As Integer)
    mblnEditing = False
    EditString = ToVarchar(EditString, 4000)
    '����:�����������ݱ����洢
    If SelCell.Key = "" Then Exit Sub
    Call AddUndo(SelCell)
    With SelCell
        If .�������� = cprCTFixtext Or .�������� = cprCTText Then
        .�����ı� = EditString
            If InStr(.��������, ",") > 0 And InStr(.��������, ";") = 0 Then '�ϼƵ�Ԫ���Դ��Ԫ��
                Dim lsumRow As Long, lsumCol As Long
                lsumRow = Split(.��������, ",")(0): lsumCol = Split(.��������, ",")(1) '�ϼƵ�Ԫ�������
                Call CalcSumRange(lsumRow, lsumCol)
            End If
        End If
    End With
End Sub

Private Sub F1Main_GotFocus()

    If PicEdit.Visible Then PicEdit.Visible = False: PicEdit.Top = 0: PicEdit.Left = 0: PicEdit.Tag = ""
    If mblnEditing Then F1Main.EndEdit
    If elEdit.Visible Then elEdit.Visible = False: elEdit.Top = 0: elEdit.Left = 0: elEdit.Tag = ""
    If Doc.Visible Then GetFromDoc Doc.Tag, True: Doc.Text = "": Doc.ForceEdit = False: Doc.Visible = False: Doc.Top = 0: Doc.Left = 0: Doc.Title = "": Doc.Tag = ""
End Sub

Private Sub F1Main_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyDelete Then Exit Sub
    
    On Error Resume Next
    'ɾ��ͼƬ���ı�
    If F1Main.SelectionCount > 1 Then Exit Sub
    Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    
    With Document.Cells.Cell(lsRow, lsCol)
        If Not (.�������� = cprCTFixtext Or .�������� = cprCTText Or .�������� = cprCTElement Or .�������� = cprCTPicture Or .�������� = cprCTReportPic) Then Exit Sub '��(�̶�����ͨ�ı���ͼƬ)�����Ա����ֱ��ɾ������
        If editType = TabET_�������༭ Or editType = TabET_��������� Then
            If Not AllowEdit(lsRow, lsCol) Then Exit Sub          '������༭ֱ���˳�
        End If
        
        AddUndo Document.Cells(.Key)
        Select Case .��������
            Case cprCTFixtext, cprCTText
                F1Main.TextRC(lsRow, lsCol) = ""
                .�����ı� = ""
                If InStr(.��������, ",") > 0 And InStr(.��������, ";") = 0 Then '�ϼƵ�Ԫ���Դ��Ԫ��
                    Dim lsumRow As Long, lsumCol As Long
                    lsumRow = Split(.��������, ",")(0): lsumCol = Split(.��������, ",")(1) '�ϼƵ�Ԫ�������
                    Call CalcSumRange(lsumRow, lsumCol)
                End If
            Case cprCTElement
                If Document.Elements("K" & .ElementKey).������̬ = 0 And Document.Elements("K" & .ElementKey).Ҫ������ = 2 Then
                    .�����ı� = ""
                    Document.Elements("K" & .ElementKey).�����ı� = ""
                    F1Main.TextRC(lsRow, lsCol) = "[" & Document.Elements("K" & .ElementKey).Ҫ������ & "]" & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                End If
            Case cprCTPicture, cprCTReportPic
                If editType = TabET_��������� Then Exit Sub
                If .PictureKey <> "" Then
                    Document.Pictures("K" & .PictureKey).OrigPic = New StdPicture
                    .PicMarkKey = ""
                    If ChkControl(PicDy(.Index)) Then Unload PicDy(.Index)
                End If
        End Select
    End With
    Err.Clear

End Sub

Private Sub F1Main_KeyPress(KeyAscii As Integer)
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, blnEndEdit As Boolean
    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii = 3 Then KeyAscii = 0: Exit Sub 'Ctrl+C
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub 'Ctrl+v
    If KeyAscii = 24 Then KeyAscii = 0: Exit Sub 'Ctrl+X
    
    If F1Main.SelectionCount > 1 Then Exit Sub  '������������ֵ
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    Call EnterEdit(lsRow, lsCol, leRow, leCol, KeyAscii)
End Sub

Private Sub F1Main_LostFocus()
    F1Main.EndEdit
End Sub

Private Sub F1Main_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngRow As Long, lngCol As Long
    F1Main.TwipsToRC x, y, lngRow, lngCol
    If lngRow = 0 Or lngCol = 0 Then
        mblnClickZ = True
        mblnChangeRC = True
    Else
        mblnClickZ = False
    End If
'    If lngRow > 0 And lngCol > 0 And Shift = 0 Then
'        Call F1Main.SetSelection(lngRow, lngCol, lngRow, lngCol)
'    End If
End Sub

Private Sub F1Main_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'ȱʡ���̫�ѿ���������/��ͷ����Ҫ��ʾResizeͼ�꣬��̬�ı�
Dim lngRow As Long, lngCol As Long, vRect As F1Rect
    On Error Resume Next
    F1Main.TwipsToRC x, y, lngRow, lngCol
    If lngRow = 0 Then
        If lngCol = 0 Then F1Main.MousePointer = F1Arrow: Err.Clear: Exit Sub
        Set vRect = F1Main.RangeToTwipsEx(1, lngCol, 1, lngCol)
        If x < vRect.Left + 20 Or x > vRect.Right - 20 Then
            F1Main.MousePointer = F1SizeWE
        Else
            F1Main.MousePointer = F1Arrow
        End If
    ElseIf lngCol = 0 Then
        If lngRow = 0 Then F1Main.MousePointer = F1Arrow: Err.Clear: Exit Sub
        Set vRect = F1Main.RangeToTwipsEx(lngRow, 1, lngRow, 1)
        If y < vRect.Top + 20 Or y > vRect.Bottom - 20 Then
            F1Main.MousePointer = F1SizeNS
        Else
            F1Main.MousePointer = F1Arrow
        End If
    Else
        F1Main.MousePointer = F1Arrow
    End If
    Err.Clear
End Sub

Private Sub F1Main_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnChangeRC Then timeTmp.Enabled = True
    If mblnInit Then Exit Sub
Dim lngRow As Long, lngCol As Long
    F1Main.TwipsToRC x, y, lngRow, lngCol

    If lngRow = 0 Or lngCol = 0 Then Exit Sub
    With Document.Cells.Cell(lngRow, lngCol)
        '���ı����ڶ��塢����ʱ��������ѡ�У��ڱ༭ʱ��������༭״̬
        If (Not (.�������� = cprCTFixtext Or .�������� = cprCTText)) And (editType = TabET_�������༭ Or editType = TabET_���������) Then
            Call F1Main_DblClick(lngRow, lngCol)
        End If
    End With
End Sub

Private Sub F1Main_SelChange()
Dim vCell As F1CellFormat, lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long
Dim lngRow As Long, lngCol As Long, lngWidth As Long, lngHeight As Long
    On Error GoTo errHand
    If mblnEditing Then mblnEditing = False
    If Not mblnInit Then
        Call F1Main.GetSelection(F1Main.SelectionCount - 1, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
        If lngEndCol > F1Main.MaxCol Then lngEndCol = F1Main.MaxCol: F1Main.SetSelection lngStarRow, lngStarCol, lngEndRow, lngEndCol 'ͨ��ѡ���0��ѡ����������
        If lngEndRow > F1Main.MaxRow Then lngEndRow = F1Main.MaxRow: F1Main.SetSelection lngStarRow, lngStarCol, lngEndRow, lngEndCol 'ͨ��ѡ���0��ѡ����������
        
        Set vCell = F1Main.GetCellFormat
        Set SelCell = Document.Cells.Cell(lngStarRow, lngStarCol)
        Call SetColorIcon(ID_FORMAT_FORECOLOR, SelCell.FontColor)
        
        If F1Main.SelectionCount > 1 Then
            stbThis.Panels("msg").Text = "���ѡȡ�൥Ԫ��"
            Call ShowAttr(-1, "")
        Else
            stbThis.Panels("msg").Text = lngStarRow & "�� " & lngStarCol & "��--" & lngEndRow & "�� " & lngEndCol & "��"
            'ˢ���������
            If vCell.MergeCells Then 'ѡ�е��Ǻϲ���Ԫ��
                ShowAttr Me.Document.Cells.Cell(lngStarRow, lngStarCol).��������, Me.Document.Cells.Cell(lngStarRow, lngStarCol).��������
                stbThis.Panels("msg").Text = stbThis.Panels("msg").Text & " ����:" & Me.Document.Cells.Cell(lngStarRow, lngStarCol).CellTypeName
                For lngRow = lngStarRow To lngEndRow
                    lngHeight = lngHeight + F1Main.RowHeight(lngRow)
                Next
                For lngCol = lngStarCol To lngEndCol
                    lngWidth = lngWidth + F1Main.ColWidth(lngCol)
                Next
                If editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭ Then
                    stbThis.Panels("msg").Text = stbThis.Panels("msg").Text & " �߶�(����):" & Round(Me.ScaleY(lngHeight, vbTwips, vbMillimeters), 2) & " ���(����):" & Round(Me.ScaleX(lngWidth, vbTwips, vbMillimeters), 2)
                End If
            Else
                If lngStarRow <> lngEndRow Or lngStarCol <> lngEndCol Then 'ѡ�еĶ����Ԫ��
                    Call ShowAttr(-1, "")
                Else
                    ShowAttr Me.Document.Cells.Cell(lngStarRow, lngStarCol).��������, Me.Document.Cells.Cell(lngStarRow, lngStarCol).��������
                    stbThis.Panels("msg").Text = stbThis.Panels("msg").Text & " ����:" & Me.Document.Cells.Cell(lngStarRow, lngStarCol).CellTypeName
                    lngHeight = lngHeight + F1Main.RowHeight(lngStarRow): lngWidth = lngWidth + F1Main.ColWidth(lngStarCol)
                    If editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭ Then
                        stbThis.Panels("msg").Text = stbThis.Panels("msg").Text & " �߶�(����):" & Round(Me.ScaleY(lngHeight, vbTwips, vbMillimeters), 2) & " ���(����):" & Round(Me.ScaleX(lngWidth, vbTwips, vbMillimeters), 2)
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub F1Main_StartEdit(EditString As String, Cancel As Integer)
    mblnEditing = True
End Sub

Private Sub F1Main_TopLeftChanged()
Dim i As Integer, strCellKey As String
    On Error GoTo errHand
    If mblnInit Then Exit Sub
    If elEdit.Visible Then Exit Sub
    If Doc.Visible Then Exit Sub
    If PicEdit.Visible Then Exit Sub
    For i = 1 To PicDy.UBound
        If ChkControl(PicDy(i)) Then
            If PicDy(i).Picture.Handle <> 0 Then
                strCellKey = Split(PicDy(i).Tag, "|")(1)
                Call PaintPictureOnTable(strCellKey)
            End If
        End If
    Next
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
'���ܣ������������ģ̬��ʽ�򿪣�Ȼ���ٽ���������ģ̬���壬����������DockPane��ʽ����Ĵ���ᱻĪ������ɲ����ã����´��뱣֤�Ӵ��崦�ڿ���״̬
Dim lngStyle As Long
    On Error Resume Next
    If Not mfrmSentence Is Nothing Then
        If Not mfrmSentence.Enabled Then
            lngStyle = GetWindowLong(mfrmSentence.hWnd, GWL_STYLE)
            SetWindowLong mfrmSentence.hWnd, GWL_STYLE, lngStyle And Not WS_DISABLED
        End If
    End If
    If Not mfrmPacsPic Is Nothing Then
        If Not mfrmPacsPic.Enabled Then
            lngStyle = GetWindowLong(mfrmPacsPic.hWnd, GWL_STYLE)
            SetWindowLong mfrmPacsPic.hWnd, GWL_STYLE, lngStyle And Not WS_DISABLED
        End If
    End If
    Err.Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Select Case KeyCode
            Case vbKeyA
                ContentMove "All"
        End Select
    End If
End Sub

Private Sub Form_Load()
    Call InitForms
    Call MainCommandbarDefine
    Call MainDockPaneDefine
    Call mfrmSentence.zlSubRefClass(Document.EPRFileInfo.����, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.ҽ��id, Me)
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim IntReturn As Integer
    Err.Clear
    On Error Resume Next
1    If stbThis.Panels("msg").Text <> "�������" And mReadOnly = 0 Then
2        'ǩ���������˳�����ǰ���棻�˵��˳����رմ��壬���Ƽ��˳����ᣬ������Ҫ��ʾ.
3        IntReturn = MsgBox("�Ƿ񱣴�����˳���" & vbCrLf & vbCrLf & "    [��]�����˳���[��]ֱ���˳�,[ȡ��]���˳���", vbQuestion + vbYesNoCancel + vbDefaultButton1, gstrSysName)
4        If IntReturn = vbYes Then
5            Call SaveDoc(, True)
6        ElseIf IntReturn = vbCancel Then
7            Cancel = True
8            Exit Sub
9        End If
10    End If
    
11    If Document.EPRPatiRecInfo.ҽ��id <> 0 Then Call Document.frmEditorClosed(Document.EPRPatiRecInfo.ҽ��id)
12    Call SaveWinState(Me, App.ProductName)
13    If Not mfrmPacsPic Is Nothing Then Unload mfrmPacsPic
14    Set mfrmPacsPic = Nothing
15    If Not mfrmSentence Is Nothing Then Unload mfrmSentence
16    Set mfrmSentence = Nothing
17    If Not mfrmMainError Is Nothing Then Unload mfrmMainError
18    Set mfrmMainError = Nothing
19    If Not mfrmEPRModelSaveAs Is Nothing Then Unload mfrmEPRModelSaveAs
20    Set mfrmEPRModelSaveAs = Nothing
21    If Not mfrmTipInfo Is Nothing Then Unload mfrmTipInfo
22    Set mfrmTipInfo = Nothing
23    Set SelCell = Nothing
24    Set Document = Nothing
25    Set DocOld = Nothing
26    Set Undo = Nothing
27    Unload frmPublicIcon
28    Unload frmPicTypeset
      Err.Clear
      Exit Sub
End Sub

Private Sub mfrmEPRModelSaveAs_SaveModels(lngDemoId As Long, blnOK As Boolean)
Dim boldem As Byte, boldet As Byte
Dim arrSQL As Variant, i As Integer, blnBegin As Boolean

    On Error GoTo errHand
    arrSQL = Array()
    boldem = EditMode
    boldet = editType
    
    Document.EM = TabEm_�޸�
    Document.ET = TabET_ȫ��ʾ���༭
    Document.EPRDemoInfo.GetDemoInfo lngDemoId
    Call Document.SaveDoc(arrSQL)
    
    blnBegin = True
    gcnOracle.BeginTrans '--------------------------д������e
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "д������")
    Next
    blnOK = True
    gcnOracle.CommitTrans
    
    Document.EM = boldem
    Document.ET = boldet
    blnBegin = False
    Exit Sub
errHand:
    If blnBegin Then gcnOracle.RollbackTrans
    Document.EM = boldem
    Document.ET = boldet
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmMainError_Location(ByVal strRange As String, ByVal elKey As Long)
Dim lRow As Long, lCol As Long
    On Error GoTo errHand
    If strRange = "" Or InStr(strRange, "|") = 0 Then
        Call F1Main.SetSelection(1, 1, 1, 1)
        F1Main.SetFocus
        Exit Sub
    End If
    
    lRow = Split(strRange, "|")(0): lCol = Split(strRange, "|")(1)
    If lRow <> 0 And lCol <> 0 Then '��Ҫ��ֱ�Ӷ�λ
        F1Main.SetSelection lRow, lCol, lRow, lCol
        If F1Main.Visible And F1Main.Enabled Then F1Main.SetFocus
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmPacsPic_InsertPicture(pic As StdPicture)
Dim lngKey As Long, l As Long
    On Error GoTo errHand
    
    If SelCell.�������� <> cprCTReportPic Then Exit Sub '��ѡ��Ԫ��Ϊ�Ǳ���ͼ
    If Document.Pictures("K" & SelCell.PictureKey).OrigPic.Handle <> 0 Then
        '����ͼƬ���򳤿�
        Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, vR As F1Rect, lWidth As Long, lHeight As Long
        If SelCell.Merge Then
            lsRow = Split(Split(SelCell.MergeRange, ";")(0), ",")(0): leRow = Split(Split(SelCell.MergeRange, ";")(1), ",")(0)
            lsCol = Split(Split(SelCell.MergeRange, ";")(0), ",")(1): leCol = Split(Split(SelCell.MergeRange, ";")(1), ",")(1)
            Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
            lWidth = vR.Width: lHeight = vR.Height
        Else
            lWidth = SelCell.Width: lHeight = SelCell.Height
        End If
        
        frmPicTypeset.ShowTypeset Me, SelCell.Key, Document.EPRPatiRecInfo.ҽ��id, lWidth, lHeight, _
            Document.Pictures("K" & SelCell.PictureKey).OrigPic, pic, Document.EPRFileInfo.lngModule
    Else
        '��ͼ���
        lngKey = Document.Pictures.Add
        Set Document.Pictures("K" & lngKey).OrigPic = pic  '����ͼƬ
        Document.Cells(SelCell.Key).PictureKey = lngKey
        Call PaintPictureOnTable(SelCell.Key)    '�ػ�ͼƬ�ͱ��
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmSentence_RowDblClick(ByVal lngWordId As Long)
Dim rsTemp As ADODB.Recordset, strText As String, lngStart As Long, lngLen As Long, lKey As Long
    
    On Error GoTo errHand
    If SelCell Is Nothing Then Exit Sub
    
    gstrSQL = "Select * From �����ʾ���� Where �ʾ�id = [1] Order By ���д���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngWordId)
    If rsTemp.EOF Then Exit Sub
    
    Select Case SelCell.��������
        Case cprCTText                    '�ı��ͣ�ֱ�ӽ�Ҫ��ת�����ı�
            strText = ""
            With rsTemp
                Do While Not .EOF
                    Select Case !��������
                    Case 0 '��������
                        strText = strText & IIf(IsNull(!�����ı�), " ", !�����ı�)
                    Case 1, 2 '1-��ʱ����Ҫ��,2-�̶�����Ҫ��
                        strText = strText & IIf(IsNull(!�����ı�), "{" & !Ҫ������ & "}" & !Ҫ�ص�λ, "{" & !�����ı� & "}")
                    End Select
                    .MoveNext
                Loop
            End With
            If Trim(strText) = "" Then Exit Sub
            strText = SelCell.�����ı� & strText
            SelCell.�����ı� = strText
            F1Main.TextRC(SelCell.Row, SelCell.Col) = strText
        Case cprCTTextElement           '��ϱ༭��
            If Not Doc.Visible Then Exit Sub '��ǰ���ڷǱ༭״̬,��������
            Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, leKey As Long, bInKeys As Boolean, bNeeded As Boolean
            bInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, leKey, bNeeded)
            If bInKeys Then Doc.Range(lEE, lEE).Selected
            
            Do Until rsTemp.EOF
                Select Case rsTemp!��������
                    Case 0              '�����ı�
                        If NVL(rsTemp!�����ı�) <> "" Then
                            lngStart = Doc.Selection.StartPos
                            strText = rsTemp!�����ı�
                            lngLen = Len(strText)
                            Doc.Range(lngStart, lngStart).Text = strText
                            Doc.Range(lngStart, lngStart + lngLen).Font.Protected = False
                            Doc.Range(lngStart, lngStart + lngLen).Font.Hidden = False
                            Doc.Range(lngStart + lngLen, lngStart + lngLen).Selected
                        End If
                    Case 1, 2           'Ҫ��
                        lngStart = Doc.Selection.StartPos
                        lKey = Me.Document.Elements.Add
                        Me.Document.Elements("K" & lKey).ID = 0
                        Me.Document.Elements("K" & lKey).�����ı� = NVL(rsTemp!�����ı�)
                        Me.Document.Elements("K" & lKey).Ҫ������ = NVL(rsTemp!Ҫ������)
                        Me.Document.Elements("K" & lKey).����Ҫ��ID = NVL(rsTemp!����Ҫ��ID, 0)
                        Me.Document.Elements("K" & lKey).�滻�� = NVL(rsTemp!�滻��, 0)
                        Me.Document.Elements("K" & lKey).Ҫ������ = NVL(rsTemp!Ҫ������, 0)
                        Me.Document.Elements("K" & lKey).Ҫ�س��� = NVL(rsTemp!Ҫ�س���, 0)
                        Me.Document.Elements("K" & lKey).Ҫ��С�� = NVL(rsTemp!Ҫ��С��, 0)
                        Me.Document.Elements("K" & lKey).Ҫ�ص�λ = NVL(rsTemp!Ҫ�ص�λ)
                        Me.Document.Elements("K" & lKey).Ҫ�ر�ʾ = NVL(rsTemp!Ҫ�ر�ʾ, 0)
                        Me.Document.Elements("K" & lKey).Ҫ��ֵ�� = NVL(rsTemp!Ҫ��ֵ��)
                        Me.Document.Elements("K" & lKey).������̬ = NVL(rsTemp!������̬, 0)
                        Me.Document.Elements("K" & lKey).�������� = NVL(rsTemp!��������, "||")
                        Me.Document.Elements("K" & lKey).���� = SelCell.Row & "|" & SelCell.Col
                        If Me.Document.Elements("K" & lKey).�滻�� = 1 And (Me.Document.ET = TabET_�������༭ Or Me.Document.ET = TabET_���������) Then
                            Me.Document.Elements("K" & lKey).�����ı� = GetReplaceEleValue(Me.Document.Elements("K" & lKey).Ҫ������, _
                                Me.Document.EPRPatiRecInfo.����ID, _
                                Me.Document.EPRPatiRecInfo.��ҳID, _
                                Me.Document.EPRPatiRecInfo.������Դ, _
                                Me.Document.EPRPatiRecInfo.ҽ��id, Me.Document.EPRPatiRecInfo.Ӥ��)
                        End If
                        Me.Document.Elements("K" & lKey).InsertIntoEditor Doc, editType, lngStart '��Ҫ�ز��뵱ǰλ�ã���궨λ��Ҫ��ĩ��
                End Select
                rsTemp.MoveNext
            Loop
            If Doc.Enabled And Doc.Visible Then Doc.SetFocus
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub picAtt_Resize()
Dim i As Integer
    On Error Resume Next '&H00FDD6C6&
    fraType.Top = 80: fraType.Left = 120: fraType.BackColor = &HFDD6C6
    txtSum.Left = 80: txtSum.Top = fraType.Height + 80: txtSum.Width = picAtt.Width - 160
    shpTxtSum.Move txtSum.Left - Screen.TwipsPerPixelX, txtSum.Top - Screen.TwipsPerPixelY, txtSum.Width + Screen.TwipsPerPixelX * 2, txtSum.Height + Screen.TwipsPerPixelY * 2
    cmdApply.Move txtSum.Left + txtSum.Width - cmdApply.Width, txtSum.Top + txtSum.Height + 80
    cmdSum.Move txtSum.Left, txtSum.Top + txtSum.Height + 80
    cmdAvg.Move cmdSum.Left + cmdSum.Width - 20, txtSum.Top + txtSum.Height + 80
    For i = 0 To Text1.UBound
        Text1(i).BackColor = &HFDD6C6
        Text1(i).Move 0, cmdApply.Top + cmdApply.Height + 80
        Text1(i).Width = picAtt.Width
        Text1(i).Height = picAtt.Height - Text1(i).Top
    Next
    Err.Clear
End Sub

Private Sub PicDy_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, strKey As String, strRange As String
    On Error GoTo errHand
    If mblnClickZ Then mblnClickZ = False
    If editType = TabET_��������� Then Exit Sub '���ʱ���ܱ༭
    If Button = vbLeftButton Then
        If F1Main.Enabled And F1Main.Visible Then
            strRange = Split(PicDy(Index).Tag, "|")(0)
            If InStr(strRange, ";") > 0 Then
                lsRow = Split(Split(strRange, ";")(0), ",")(0): lsCol = Split(Split(strRange, ";")(0), ",")(1)
                leRow = Split(Split(strRange, ";")(1), ",")(0): leCol = Split(Split(strRange, ";")(1), ",")(1)
            Else
                lsRow = Split(strRange, ",")(0): lsCol = Split(strRange, ",")(1)
                leRow = Split(strRange, ",")(0): leCol = Split(strRange, ",")(1)
            End If
            Call F1Main.SetSelection(lsRow, lsCol, leRow, leCol)
            strKey = Split(PicDy(Index).Tag, "|")(1)
            EditPicture strKey
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PicEdit_LostFocus()
Dim strKey As String
    strKey = PicEdit.mselKey
    PicEdit.Visible = False: PicEdit.Top = 0: PicEdit.Left = 0: PicEdit.Tag = ""
    If strKey <> "" Then PaintPictureOnTable strKey
    If F1Main.Visible And F1Main.Enabled Then F1Main.SetFocus
End Sub

Private Sub picHistory_Resize()
    vsHistory.Width = picHistory.Width
    vsHistory.Height = picHistory.Height
    vsHistory.Top = 0: vsHistory.Left = 0
End Sub

Private Sub picMainBack_GotFocus()
    If F1Main.Visible And F1Main.Enabled Then
        F1Main.SetFocus
    End If
End Sub

Private Sub picMainBack_Resize()
On Error Resume Next
    picRulerH.Top = 0: picRulerH.Left = picRulerV.Width: picRulerH.Width = picMainBack.Width
    picRulerV.Top = picRulerH.Height: picRulerV.Left = 0: picRulerV.Height = picMainBack.Height
'    F1Main.Width = picMainBack.Width - picRulerV.Width: F1Main.Height = picMainBack.Height - picRulerH.Height
'    F1Main.Top = picRulerH.Height: F1Main.Left = picRulerV.Width
    picRulerH.Visible = False: picRulerV.Visible = False
    F1Main.Width = picMainBack.Width: F1Main.Height = picMainBack.Height
    F1Main.Top = 0: F1Main.Left = 0
    picMainBack.BackColor = RGB(255, 255, 255)
Err.Clear
End Sub

Private Sub SaveAsDemo()
    On Error GoTo errHand
    
    If Doc.Visible Or mblnEditing Then F1Main_GotFocus '������ڱ༭״̬����Ҫ�����ݸ�������
    
    If mfrmEPRModelSaveAs Is Nothing Then
        Set mfrmEPRModelSaveAs = New frmEPRModelSaveAs
    End If
    
    If editType = TabET_ȫ��ʾ���༭ Then
        mfrmEPRModelSaveAs.ShowMe 1, Document.EPRDemoInfo.ID
    Else
        mfrmEPRModelSaveAs.ShowMe 2, Document.EPRPatiRecInfo.ID
    End If
    
    Unload mfrmEPRModelSaveAs
    Set mfrmEPRModelSaveAs = Nothing
    
    mfrmTipInfo.ShowTipInfo stbThis.hWnd, "��ʾ��" & vbCrLf & "      �ɹ�����ɷ��ģ�", True, 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub ImportDemo()
Dim lngModId As Long, lngFileID As Long, mfrmImportEPRDemo As New frmImportEPRDemo
    On Error GoTo errHand
    lngModId = frmImportEPRDemo.ShowMe(Me)
    If lngModId = 0 Then Exit Sub
    
    Call ClearPicture
    If editType = TabET_ȫ��ʾ���༭ Then lngFileID = Document.EPRDemoInfo.ID '������Ƿ��ı༭�ȼ���������ID
    Document.EPRDemoInfo.GetDemoInfo lngModId                                 '����ָ������ķ���ID��ȡ�����Ϣ
    Document.EM = TabEm_�޸�: Document.ET = TabET_ȫ��ʾ���༭    'ָ����ǰģʽΪ����ģʽ
    If Not Document.ReadFileStructure Then Exit Sub                           '��ȡ�ļ��ṹ
    Document.ReadFileContent mblnMoved                                        '��ȡ�ļ�����
    Document.EM = EditMode: Document.ET = editType                '�ָ�֮ǰ��ģʽ
    If lngFileID <> 0 Then Document.EPRDemoInfo.GetDemoInfo lngFileID         '�ָ����ı༭ģʽ���������Ϣ
    mblnInit = True
    RefreshF1Main                                                             'ˢ�½���
    mblnInit = False
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub ExportXml()
Dim strFile As String
    On Error GoTo errHand
    strFile = GetSaveFile(Me.hWnd, Document.EPRFileInfo.���� & ".xml", "XML�ĵ�" & Chr(0) & "*.xml" & Chr(0), "���������ļ�")
    
    If strFile = "" Then Exit Sub
    If gobjFSO.FileExists(strFile) Then
        If MsgBox("�Ƿ񸲸ǵ�ǰ�Ѵ��ڵ��ļ�?", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Call ValiCellDate
    If Not Document.BuildXmlFile(strFile, True) Then Exit Sub
    If gobjFSO.FileExists(strFile) Then
        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "��ʾ��" & vbCrLf & "      �ɹ������ļ� <" & strFile & ">", True, 0
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub ImportXml()
Dim strFile As String
'��������
    On Error GoTo errHand
    strFile = GetOpenFile(Me.hWnd, "*.xml", "XML�ļ�" & Chr(0) & "*.xml" & Chr(0), "���벡���ļ�")
    
    If strFile = "" Then Exit Sub
    If MsgBox("ȷʵҪʹ�õ����ĵ��������༭���ĵ���?", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call ClearPicture
    If Not Document.AnalyseFileStrucuture(strFile, True) Then Exit Sub
    mblnInit = True:  RefreshF1Main: mblnInit = False
    Document.EM = TabEm_����
    mfrmTipInfo.ShowTipInfo stbThis.hWnd, "��ʾ��" & vbCrLf & "      �ɹ����ļ�<" & strFile & ">��������", True, 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnInit = False
End Sub
Private Sub CompareChange(arrSQL As Variant)
'���ܣ��Ƚϱ仯�ĵ�Ԫ����ĳЩ�������ֹ��һ�棬������һ��
'���ã���˱༭ʱ���棬�ٴ��޸ı��棬���ǩ��
Dim l As Long, strKey As String, lCount As Long, lastVar As Long, blnChange As Boolean, lngTmp As Long
    lCount = Document.Cells.Count: lastVar = Document.EPRPatiRecInfo.���汾 + 1
    For l = 1 To lCount
        blnChange = False
        With Document.Cells(l)
        If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) And .ID <> 0 Then 'ID=0��ʾ��׷�ӵ��м�¼
        Select Case .��������
            Case cprCTText
                If .�����ı� <> DocOld(.Key) Then 'ԭ��¼���¼�¼���ݲ�ͬ
                    If .��ʼ�� <> lastVar Then 'ǩ��������޸Ļ�ǩ������ǩ��(����޸ĺ����޸Ĳ�������)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�������_Endvar(" & .ID & "," & lastVar & ")"
                        .ID = 0
                        .��ʼ�� = lastVar: .��ֹ�� = 0
                        .������� = Document.mMaxNo + 1: Document.mMaxNo = .�������
                    End If
                End If
            Case cprCTElement
                If .�����ı� <> DocOld(.Key) Then 'ԭ��¼���¼�¼���ݲ�ͬ
                    If Document.Elements("K" & .ElementKey).��ʼ�� <> lastVar Then   'ǩ��������޸Ļ�ǩ������ǩ��(����޸ĺ����޸Ĳ�������)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�������_Endvar(" & .ID & "," & lastVar & ")"
                        .ID = 0
                        Document.Elements("K" & .ElementKey).��ʼ�� = lastVar:  .��ʼ�� = lastVar
                        Document.Elements("K" & .ElementKey).��ֹ�� = 0:        .��ֹ�� = 0
                        .������� = Document.mMaxNo + 1: Document.mMaxNo = .�������
                    End If
                End If
            Case cprCTTextElement '����͵�Ԫ�� �����е��ı���Ҫ���游����һ����ֹ,Ȼ��ֱ������
                If .�����ı� <> DocOld(.Key) Then 'ԭ��¼���¼�¼���ݲ�ͬ
                    If .��ʼ�� <> lastVar Then 'ǩ��������޸Ļ�ǩ������ǩ��(����޸ĺ����޸Ĳ�������)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�������_Endvar(" & .ID & "," & lastVar & ")"
                        .ID = 0
                        .��ʼ�� = lastVar: .��ֹ�� = 0
                        .������� = Document.mMaxNo + 1: Document.mMaxNo = .�������
                        
                        '���ı���Ҫ�ؿ�ʼ�漰��ֹ����д���
                        For lngTmp = 1 To UBound(Split(.TextKey, "|")) '�Ա���Ԫ���������ı����б���
                            With Document.Texts("K" & Split(.TextKey, "|")(lngTmp))
                                .ID = 0: .��ʼ�� = lastVar: .��ֹ�� = 0
                            End With
                        Next
                        For lngTmp = 1 To UBound(Split(.ElementKey, "|"))
                            With Document.Elements("K" & Split(.ElementKey, "|")(lngTmp))
                                .ID = 0: .��ʼ�� = lastVar: .��ֹ�� = 0
                            End With
                        Next
                    End If
                End If
                
            Case cprCTPicture, cprCTReportPic 'ͼƬ�����Ƚ�
        End Select
        End If
        End With
    Next
End Sub
Private Sub PageSetUp()
Dim mfrmPageSetup As New frmPageSetup
    On Error GoTo errHand
    
    If mfrmPageSetup.ShowMe(Me, Document) = False Then Exit Sub
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub PrintDoc(ByVal blnPreview As Boolean)
'�Ƿ���Ԥ��blnPreview
    On Error GoTo errHand
    If Doc.Visible Or mblnEditing Then F1Main_GotFocus '������ڱ༭״̬����Ҫ�����ݸ�������
    If PicEdit.Visible Then Call PicEdit_LostFocus '���ͼƬ���ڱ༭״̬����Ҫ���ػ��ͼƬ
    
    Call Document.PrintDoc(Me, blnPreview)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub SetCellFormat(ByVal strFormat As String, ByVal vData As Variant)
'����:���õ�Ԫ���ʽ
'����:strFormat ��ʽ��,vData ����ֵ
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long
Dim vCell As F1CellFormat, l As Long
    On Error GoTo errHand
    mblnInit = True
    '�ȴ�����洢��Ϊ���������Ҫ�������Ի�ԭ
    For l = 0 To F1Main.SelectionCount - 1 '���ѡ��
        Call F1Main.GetSelection(l, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
        For lngRow = lngStarRow To lngEndRow
            For lngCol = lngStarCol To lngEndCol
                With Document.Cells.Cell(lngRow, lngCol)
                    Select Case strFormat
                        Case "��������"
                            .FontName = vData
                        Case "�ֺ�"
                            .FontSize = GetFontSizeNumber(CStr(vData))
                        Case "����"
                            .FontBold = vData
                        Case "б��"
                            .FontItalic = vData
                        Case "�»���"
                            .FontUnderline = vData
                        Case "ɾ����"
                            .FontStrikeout = vData
                        Case "����"
                            .�������� = vData
                        Case "�ϲ�"
                            .Merge = vData
                            If lngStarRow = lngEndRow And lngStarCol = lngEndCol Then
                                .Merge = False
                            Else
                                Call ClearChildMember(.Key)
                                If .Merge Then '�ϲ�
                                    If lngRow = lngStarRow And lngCol = lngStarCol Then
                                        .MergeRange = lngRow & "," & lngCol & ";" & lngEndRow & "," & lngEndCol
                                    Else
                                        .MergeRange = lngRow & "," & lngCol
                                    End If
                                Else            'ȡ���ϲ�
                                    .MergeRange = lngRow & "," & lngCol
                                    .CellLineTop = F1BorderThin: .CellLineBottom = F1BorderThin
                                    .CellLineLeft = F1BorderThin: .CellLineRight = F1BorderThin
                                    .������� = Document.mMaxNo + 1: Document.mMaxNo = Document.mMaxNo + 1
                                End If
                            End If
                        Case "������ɫ"
                            .FontColor = vData
                        Case "���������"
                            .VAlignment = VALignTop: .HAlignment = HALignLeft
                        Case "���Ͼ���"
                            .VAlignment = VALignTop: .HAlignment = HAlignCenter
                        Case "�����Ҷ���"
                            .VAlignment = VALignTop: .HAlignment = HALignRight
                        Case "�в������"
                            .VAlignment = VAlignCenter: .HAlignment = HALignLeft
                        Case "�в�����"
                            .VAlignment = VAlignCenter: .HAlignment = HAlignCenter
                        Case "�в��Ҷ���"
                            .VAlignment = VAlignCenter: .HAlignment = HALignRight
                        Case "���������"
                            .VAlignment = VALignBottom: .HAlignment = HALignLeft
                        Case "���¾���"
                            .VAlignment = VALignBottom: .HAlignment = HAlignCenter
                        Case "�����Ҷ���"
                            .VAlignment = VALignBottom: .HAlignment = HALignRight
                    End Select
                End With
            Next
        Next
    Next

    If strFormat = "�ϲ�" Then                  '�ϲ���������
'        For l = 0 To F1Main.SelectionCount - 1 '���ѡ��
'            Call F1Main.GetSelection(l, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
'            If lngStarRow = lngEndRow And lngStarCol = lngEndCol Then
'                Set vCell = F1Main.GetCellFormat  '
'                vCell.MergeCells = False
'            Else
'                Set vCell = F1Main.GetCellFormat  '
'                vCell.MergeCells = vData
'                vCell.BorderStyle(F1TopBorder) = F1BorderThin: vCell.BorderStyle(F1BottomBorder) = F1BorderThin
'                vCell.BorderStyle(F1LeftBorder) = F1BorderThin: vCell.BorderStyle(F1RightBorder) = F1BorderThin
'                For lngRow = lngStarRow To lngEndRow
'                    For lngCol = lngStarCol To lngEndCol
'                        F1Main.TextRC(lngRow, lngCol) = ""
'                    Next
'                Next
'            End If
'        Next
        Call RefreshF1Main
        F1Main.SetSelection lngStarRow, lngStarCol, lngStarRow, lngStarCol
    Else
        Dim SelR() As SelRange
        ReDim SelR(F1Main.SelectionCount - 1) As SelRange
        For l = 0 To F1Main.SelectionCount - 1 '���ѡ��        '��ѡ��������ֹ���и��Ʊ���,��Ϊ��������ı�Selection
            Call F1Main.GetSelection(l, SelR(l).lsRow, SelR(l).lsCol, SelR(l).leRow, SelR(l).leCol) '��N�μ��ѡ�����ʼ����
        Next
        
        For l = 0 To UBound(SelR)
            For lngRow = SelR(l).lsRow To SelR(l).leRow
                For lngCol = SelR(l).lsCol To SelR(l).leCol
                With Document.Cells.Cell(lngRow, lngCol)
                    If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then 'ֻ�кϲ���Ԫ���׸��ͷǺϲ���Ԫ����Ч
                        Call F1Main.SetSelection(lngRow, lngCol, lngRow, lngCol) 'ѡ�������
                        Set vCell = F1Main.GetCellFormat  '������ʾ
                        Select Case strFormat
                            Case "��������"
                                vCell.FontName = vData
                            Case "�ֺ�"
                                vCell.FontSize = GetFontSizeNumber(CStr(vData))
                            Case "����"
                                vCell.FontBold = vData
                            Case "б��"
                                vCell.FontItalic = vData
                            Case "�»���"
                                vCell.FontUnderline = vData
                            Case "ɾ����"
                                vCell.FontStrikeout = vData
                            Case "����"
                                vCell.ProtectionLocked = vData
                            Case "������ɫ"
                                vCell.FontColor = vData
                            Case "���������"
                                vCell.AlignVertical = F1VAlignTop: vCell.AlignHorizontal = F1HAlignLeft
                            Case "���Ͼ���"
                                vCell.AlignVertical = F1VAlignTop: vCell.AlignHorizontal = F1HAlignCenter
                            Case "�����Ҷ���"
                                vCell.AlignVertical = F1VAlignTop: vCell.AlignHorizontal = F1HAlignRight
                            Case "�в������"
                                vCell.AlignVertical = F1VAlignCenter: vCell.AlignHorizontal = F1HAlignLeft
                            Case "�в�����"
                                vCell.AlignVertical = F1VAlignCenter: vCell.AlignHorizontal = F1HAlignCenter
                            Case "�в��Ҷ���"
                                vCell.AlignVertical = F1VAlignCenter: vCell.AlignHorizontal = F1HAlignRight
                            Case "���������"
                                vCell.AlignVertical = F1VAlignBottom: vCell.AlignHorizontal = F1HAlignLeft
                            Case "���¾���"
                                vCell.AlignVertical = F1VAlignBottom: vCell.AlignHorizontal = F1HAlignCenter
                            Case "�����Ҷ���"
                                vCell.AlignVertical = F1VAlignBottom: vCell.AlignHorizontal = F1HAlignRight
                        End Select
                        F1Main.SetCellFormat vCell
                    End If
                End With
                Next
            Next
        Next
        For l = 0 To UBound(SelR)
            Call F1Main.AddSelection(SelR(l).lsRow, SelR(l).lsCol, SelR(l).leRow, SelR(l).leCol)
        Next
    End If
    
    mblnInit = False
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub SetCellBorder()
'����:���õ�Ԫ���ʽ
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long
Dim lOutBorder As Integer, lTopBorder As Integer, lBottomBorder As Integer, lLeftBorder As Integer, lRightBorder As Integer, lShade As Integer, lOutColor As Long, lTopColor As Long, lBottomColor As Long, lLeftColor As Long, lRightColor As Long
Dim vCell As F1CellFormat, l As Long, mfrmBorder As New frmBorder
    On Error GoTo errHand
    mblnInit = True '����SelChange
    Call F1Main.GetSelection(0, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
    With Document.Cells.Cell(lngStarRow, lngStarCol)
        lLeftBorder = .CellLineLeft: lRightBorder = .CellLineRight: lTopBorder = .CellLineTop: lBottomBorder = .CellLineBottom
        lLeftColor = .CellLineLeftColor: lRightColor = .CellLineRightColor: lTopColor = .CellLineTopColor: lBottomColor = .CellLineBottomColor
    End With
    lOutBorder = -1: lOutColor = -1
    
    If Not mfrmBorder.ShowMe(lOutBorder, lLeftBorder, lRightBorder, lTopBorder, lBottomBorder, lShade, lOutColor, lLeftColor, lRightColor, lTopColor, lBottomColor, Me) Then mblnInit = False: Exit Sub
    
'   ���洦��ȽϷ�������ϲ���Ԫ���ԭ�����ԣ��޷�����ֱ�Ӵ�����洢��ˢ��

    '������洢��ǰ����
    On Error Resume Next
    For l = 0 To F1Main.SelectionCount - 1 '���ѡ��
        Call F1Main.GetSelection(l, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
        For lngRow = lngStarRow To lngEndRow  '��ʼ��-��ֹ��
            For lngCol = lngStarCol To lngEndCol '��ʼ��-��ֹ��
                With Document.Cells.Cell(lngRow, lngCol)
                    If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then '��Ч��Ԫ��
                        .CellLineTop = lTopBorder: .CellLineBottom = lBottomBorder: .CellLineLeft = lLeftBorder: .CellLineRight = lRightBorder
                        .CellLineTopColor = lTopColor: .CellLineBottomColor = lBottomColor: .CellLineLeftColor = lLeftColor: .CellLineRightColor = lRightColor
                        If lngRow > 1 Then
                        With Document.Cells.Cell(lngRow - 1, lngCol)    '���趨��Ԫ����Ϸ���Ԫ���±���
                            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                .CellLineBottom = lTopBorder
                                .CellLineBottomColor = lTopColor
                            End If
                        End With
                        End If
                        
                        If lngRow < F1Main.MaxRow Then
                        With Document.Cells.Cell(lngRow + 1, lngCol)    '���趨��Ԫ���·���Ԫ����ϱ���
                            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                .CellLineTop = lBottomBorder
                                .CellLineTopColor = lBottomColor
                            End If
                        End With
                        End If
                        
                        If lngCol > 1 Then
                        With Document.Cells.Cell(lngRow, lngCol - 1)    '���趨��Ԫ���󷽵�Ԫ����ұ���
                            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                .CellLineRight = lLeftBorder
                                .CellLineRightColor = lLeftColor
                            End If
                        End With
                        End If
                        
                        If lngCol < F1Main.MaxCol Then
                        With Document.Cells.Cell(lngRow, lngCol + 1)    '���趨��Ԫ���ҷ���Ԫ��������
                            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                .CellLineLeft = lRightBorder
                                .CellLineLeftColor = lRightColor
                            End If
                        End With
                        End If
                    Else                                                   '�ϲ���Ԫ��ķ��׸���Ԫ��Ϊ��Ч��Ԫ��
                        If lngCol <> lngEndCol Then
                            .CellLineTop = 0: .CellLineBottom = 0: .CellLineLeft = 0: .CellLineRight = 0
                            .CellLineTopColor = 0: .CellLineBottomColor = 0: .CellLineLeftColor = 0: .CellLineRightColor = 0
                        Else    '�ϲ���Ԫ�����
                            If lngCol < F1Main.MaxCol Then
                            With Document.Cells.Cell(lngRow, lngCol + 1)    '���趨��Ԫ���ҷ���Ԫ��������
                                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                    .CellLineLeft = lRightBorder
                                    .CellLineLeftColor = lRightColor
                                End If
                            End With
                            End If
                        End If
                    End If
                End With
            Next
        Next
    Next
    RefreshF1Main
    Err.Clear
    mblnInit = False
    F1Main.SetSelection lngStarRow, lngStarCol, lngStarRow, lngStarCol
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnInit = False
End Sub
Private Function SetSumAtt(ByRef vData As String) As Boolean
',��ǰ��Ԫ������Щ��Ԫ��ϼƵ���
'vdataΪSumʱ��ʽΪ ��,��;��,��----,����Ϊ��,
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long, strTmp As String, l As Long
    On Error GoTo errHand
    
    Call F1Main.GetSelection(0, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
    lngRow = lngStarRow: lngCol = lngStarCol
        
    If vData <> "" Then '=""��ʾȡ���ϼ�����
        For l = 0 To UBound(Split(vData, ";"))
            strTmp = Split(vData, ";")(l) '��Ԫ������=0�򳬹����������л�=��ǰ��Ԫ��������Ϊ��Ч
            If Split(strTmp, ",")(0) > F1Main.MaxRow Or Split(strTmp, ",")(1) > F1Main.MaxCol Then
                vData = "�ϼƵ�Ԫ�����Դ��Ԫ�� " & Split(strTmp, ",")(0) & "�� " & Split(strTmp, ",")(1) & "�� �������Χ�����飡": Exit Function
            End If
            If (Split(strTmp, ",")(0) = lngRow And Split(strTmp, ",")(1) = lngCol) Then
                vData = "�ϼƵ�Ԫ�����Դ��Ԫ���������� " & Split(strTmp, ",")(0) & "�� " & Split(strTmp, ",")(1) & "�У����飡": Exit Function
            End If
                 
            If Not (Document.Cells.Cell(Split(strTmp, ",")(0), Split(strTmp, ",")(1)).�������� = cprCTFixtext Or Document.Cells.Cell(Split(strTmp, ",")(0), Split(strTmp, ",")(1)).�������� = cprCTText) Then
                vData = "�ϼƵ�Ԫ�����Դ��Ԫ�� " & Split(strTmp, ",")(0) & "�� " & Split(strTmp, ",")(1) & "�� ���ǹ̶��ı�/�����ı���Ԫ�����飡": Exit Function
            End If
        Next
    End If
    
    
    With Document.Cells.Cell(lngRow, lngCol)
        strTmp = "" '��ȡ��ԭ�кϼ�����
        If .�������� <> "" And (.�������� = cprCTFixtext Or .�������� = cprCTText) Then 'Ŀ�굥Ԫ��
            For l = 0 To UBound(Split(.��������, ";"))
                strTmp = Split(.��������, ";")(l)
                If UBound(Split(strTmp, ",")) > 0 Then 'ȷ������������ȷ��,ֻ�й̶��ı��Ͷ����ı����кϼ�����
                    With Document.Cells.Cell(Split(strTmp, ",")(0), Split(strTmp, ",")(1)) 'Դ��Ԫ��
                        If .�������� = cprCTFixtext Or .�������� = cprCTText Then
                            .�������� = ""
                        End If
                    End With
                End If
            Next
            .�������� = ""
        End If
        
        If vData <> "" Then '�趨�ϼ�����
            For l = 0 To UBound(Split(vData, ";"))
                strTmp = Split(vData, ";")(l)
                Document.Cells.Cell(Split(strTmp, ",")(0), Split(strTmp, ",")(1)).�������� = lngRow & "," & lngCol
            Next
        End If
        .�������� = vData
    End With
    SetSumAtt = True: vData = ""
    CalcSumRange lngRow, lngCol
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetCellAttribute(ByVal strType As String, ByRef vData As String, blnReturn As Boolean)
'����:���õ�Ԫ���ʽ
'����:strType ����0,1,2,3,4,5,6,7,8 �ֱ��ʾ�̶�TXT,����TXT,��Ҫ��,��ϱ༭,�ο�ͼ,����ͼ,�п�ǩ��,�п�ǩ��,ǩ��λ
'    vData �������,ʧ��ʱ������Ϣ;blnReturn�趨�Ƿ�ɹ�
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long
Dim lmsRow As Long, lmsCol As Long, lmeRow As Long, lmeCol As Long, lCellCount As Long
Dim lngTmp As Long, vR As F1Rect, strTmp As String, lS As Long, l As Long, j As Long
    
    
    For lS = 0 To F1Main.SelectionCount - 1
        Call F1Main.GetSelection(lS, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
        If strType = cprCTColSign Or strType = cprCTRowSign Or strType = cprCTSign Then '���С��пء�ǩ��λ���ж�
            lCellCount = Document.Cells.Count
            For l = 1 To lCellCount
                With Document.Cells(l)
                    Select Case strType '�С��пء�ǩ��������ͬһ�ĵ��г���
                        Case cprCTSign
                            If .�������� = cprCTColSign Or .�������� = cprCTRowSign Then
                                vData = "�пء��пء���ͨǩ��������ͬһ�ĵ��ڳ��֣�": Exit Sub
                            End If
                        Case cprCTColSign
                            If .�������� = cprCTSign Or .�������� = cprCTRowSign Then
                                vData = "�пء��пء���ͨǩ��������ͬһ�ĵ��ڳ��֣�": Exit Sub
                            End If
                        Case cprCTRowSign
                            If .�������� = cprCTSign Or .�������� = cprCTColSign Then
                                vData = "�пء��пء���ͨǩ��������ͬһ�ĵ��ڳ��֣�": Exit Sub
                            End If
                    End Select
                    
                    If .Merge And InStr(.MergeRange, ";") > 0 Then '�Ժϲ�����������ж�
                        lmsRow = Val(Split(Split(.MergeRange, ";")(0), ",")(0)): lmsCol = Val(Split(Split(.MergeRange, ";")(0), ",")(1))
                        lmeRow = Val(Split(Split(.MergeRange, ";")(1), ",")(0)): lmeCol = Val(Split(Split(.MergeRange, ";")(1), ",")(1))
                        If strType = cprCTColSign Then
                            If lmsCol <> lmeCol And .Row <> 1 Then '�п��кϲ�������Ҳ��ڵ�һ��,��Ϊ��һ�п����Ǳ���
                            For j = lngStarCol To lngEndCol 'ѡ�е����г��ֱ��ϲ����
                                If j >= lmsCol And j <= lmeCol Then
                                    vData = "�п�ǩ�������в����п��кϲ���Ԫ��": Exit Sub
                                End If
                            Next
                            End If
                        ElseIf strType = cprCTRowSign Then
                            If lmsRow <> lmeRow Then '�п��кϲ������
                            For j = lngStarRow To lngEndRow 'ѡ�е����г��ֱ��ϲ������
                                If j >= lmsRow And j <= lmeRow Then
                                    vData = "�п�ǩ�������в����п��кϲ���Ԫ��": Exit Sub
                                End If
                            Next
                            End If
                        End If
                    End If
                End With
            Next
        End If
        
        For lngRow = lngStarRow To lngEndRow
            For lngCol = lngStarCol To lngEndCol
                With Document.Cells.Cell(lngRow, lngCol)
                    '��ɾ��ԭ�ж������� �ڼ����еļ�¼,���ܼ��ѡ������Ԫ��,ֻ���� ���ǰ���������ͷ����仯��
                    If .�������� <> strType Then
                        Call ClearChildMember(.Key)
                        .�������� = strType: .�������� = "": .�����д� = 0: .�����ı� = ""
                        Select Case .��������
                            Case cprCTFixtext          '0-�̶��ı�(���ɱ༭)
                                .�������� = True
                                F1Main.TextRC(lngRow, lngCol) = .�����ı�
                            Case cprCTText            '1-�ı���(�ɱ༭�����ı�)
                                .�������� = False
                                F1Main.TextRC(lngRow, lngCol) = .�����ı�
                            Case cprCTElement          '2-��Ҫ��
                                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                    .ElementKey = Me.Document.Elements.Add
                                    Me.Document.Elements("K" & .ElementKey).ID = .ID
                                    Me.Document.Elements("K" & .ElementKey).��ID = 0
                                    Me.Document.Elements("K" & .ElementKey).���� = lngRow & "|" & lngCol
                                    InsertElement .Key '����Ҫ��
                                End If
                                .�������� = False
                            Case cprCTTextElement       '3-�ı����Ҫ�ػ�ϱ༭\
                                .�������� = False: .ElementKey = "": .TextKey = ""
                                F1Main.TextRC(lngRow, lngCol) = .�����ı�
                            Case cprCTPicture, cprCTReportPic          '4-�ο�ͼ
                                .�������� = False
                                .�����ı� = IIf(.�������� = cprCTPicture, "�ο�ͼ", "����ͼ")
                                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                    Set vR = F1Main.RangeToTwipsEx(lngStarRow, lngStarCol, lngEndRow, lngEndCol)
                                    .PictureKey = Me.Document.Pictures.Add
                                    Me.Document.Pictures("K" & .PictureKey).PicID = .ID
                                    Me.Document.Pictures("K" & .PictureKey).DesWidth = vR.Width
                                    Me.Document.Pictures("K" & .PictureKey).DesHeight = vR.Height
                                End If
                                F1Main.TextRC(lngRow, lngCol) = .�����ı�
                            Case cprCTSign, cprCTRowSign, cprCTColSign          '6-ǩ��
                                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                                    .SignKey = Me.Document.Signs.Add
                                End If
                                .�����ı� = "[ǩ��λ]"
                                .�������� = False
                                F1Main.TextRC(lngRow, lngCol) = .�����ı�
                        End Select
                    End If
                End With
            Next
        Next
    Next
    blnReturn = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub SetRowCol(ByVal strType As String)
'����:��������
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long
Dim l As Long, lngHeight As Long, lngWidth As Long, mfrmSetRowCol As New frmSetRowCol
    On Error GoTo errHand
    Call F1Main.GetSelection(0, lngStarRow, lngStarCol, lngEndRow, lngEndCol) 'ȡ�����һ��ѡ��������ֹ����
    Select Case strType
        Case "�и�"
            lngHeight = F1Main.RowHeight(lngStarRow)
            Call mfrmSetRowCol.SetRowCol(Me, strType, lngHeight)
        Case "�п�"
            lngWidth = F1Main.ColWidthTwips(lngStarCol)
            Call mfrmSetRowCol.SetRowCol(Me, strType, lngWidth)
        Case "��ͬ�и�"
            lngHeight = F1Main.RowHeight(lngStarRow)
        Case "��ͬ�п�"
            lngWidth = F1Main.ColWidthTwips(lngStarCol)
    End Select
    If lngHeight = -1 Or lngWidth = -1 Then Exit Sub
    
    For l = 0 To F1Main.SelectionCount - 1
        Call F1Main.GetSelection(l, lngStarRow, lngStarCol, lngEndRow, lngEndCol)
        If lngHeight <> 0 Then '�����и�
            For lngRow = lngStarRow To lngEndRow
                F1Main.RowHeight(lngRow) = lngHeight
            Next
        End If
        If lngWidth <> 0 Then '�����п�
            For lngCol = lngStarCol To lngEndCol
                F1Main.ColWidthTwips(lngCol) = lngWidth
            Next
        End If
    Next
    timeTmp.Enabled = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function InsertRowCol(ByVal strType As String) As Boolean
'����:��������
'˵��������ǰ�жϺϲ���Ԫ�񣬲����ı���洢��������洢���ı�չ����ʽ
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long
Dim lmsRow As Long, lmsCol As Long, lmeRow As Long, lmeCol As Long '�ϲ��������ֹ����
Dim lCRow As Long, lCCol As Long, lNRow As Long, lNCol As Long
Dim vCell As F1CellFormat, l As Long, IntInsertType As Integer, j As Integer, strKey As String, lCellCount As Long
Dim lMaxR As Long, lMaxC As Long, lR As Long, lC As Long, TmpCell As cTabCell
    On Error GoTo errHand
    mblnInit = True
    Call F1Main.GetSelection(0, lngStarRow, lngStarCol, lngEndRow, lngEndCol) 'ȡ���׸�ѡ��������ֹ����
    If lngEndRow > F1Main.MaxRow Then lngEndRow = F1Main.MaxRow
    If lngEndCol > F1Main.MaxCol Then lngEndCol = F1Main.MaxCol
    lCellCount = Document.Cells.Count: lMaxR = Document.Cells.Rows: lMaxC = Document.Cells.Cols
    'ȷ�������У�����ǰ����
    Select Case strType
        Case "InsertLeftCol" '�����������
            lNRow = lngStarRow: lNCol = lngStarCol: lCRow = lngStarRow: lCCol = lngStarCol: IntInsertType = F1ShiftCols
        Case "InsertRightCol" '���Ҳ�������
            lNRow = lngEndRow: lNCol = lngEndCol + 1: lCRow = lngEndRow: lCCol = lngEndCol: IntInsertType = F1ShiftCols
        Case "InsertUpRow"  '���ϲ�������
            lNRow = lngStarRow: lNCol = lngStarCol: lCRow = lngStarRow: lCCol = lngStarCol: IntInsertType = F1ShiftRows
        Case "InsertDnRow"  '���²�������
            lNRow = lngEndRow + 1: lNCol = lngEndCol: lCRow = lngEndRow: lCCol = lngEndCol: IntInsertType = F1ShiftRows
    End Select
    
    If Not (lNRow > F1Main.MaxRow Or lNCol > F1Main.MaxCol) Then '�����׷�Ӳ���Ҫ�ж�
    For l = 1 To lCellCount
        With Document.Cells(l)
            If .Merge And InStr(.MergeRange, ";") > 0 Then '�Ժϲ�����������ж�
                lmsRow = Val(Split(Split(.MergeRange, ";")(0), ",")(0)): lmsCol = Val(Split(Split(.MergeRange, ";")(0), ",")(1))
                lmeRow = Val(Split(Split(.MergeRange, ";")(1), ",")(0)): lmeCol = Val(Split(Split(.MergeRange, ";")(1), ",")(1))
                
                If IntInsertType = F1ShiftCols Then '������
                    If lmsCol <> lmeCol Then '�п��кϲ������
                        If strType = "InsertLeftCol" Then '���󷽲����У����в��ܳ�������ǰһ�кϲ������
                            If lCCol > lmsCol And lCCol <= lmeCol Then
                                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      �������У���ǰ��Ԫ�������в����п��кϲ���Ԫ��" & vbCrLf & "      ���飡", True, 1
'                                stbThis.Panels("msg").Text = "��������ʱ����ǰѡ�еĵ�Ԫ�������в����п��кϲ���Ԫ��"
                                mblnInit = False: Exit Function
                            End If
                        Else                              '���ҷ������У����в��ܳ��������һ�кϲ������
                            If lCCol >= lmsCol And lCCol < lmeCol Then
                                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      �������У���ǰ��Ԫ�������в����п��кϲ���Ԫ��" & vbCrLf & "      ���飡", True, 1
'                                stbThis.Panels("msg").Text = "��������ʱ����ǰѡ�еĵ�Ԫ�������в����п��кϲ���Ԫ��"
                                mblnInit = False: Exit Function
                            End If
                        End If
                    End If
                Else                    '������
                    If lmsRow <> lmeRow Then '�п��кϲ������
                        If strType = "InsertUpRow" Then '���Ϸ������У����в��ܳ�������ǰһ�кϲ������
                            If lCRow > lmsRow And lCRow <= lmeRow Then
                                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      �������У���ǰ��Ԫ�������в����п��кϲ���Ԫ��" & vbCrLf & "      ���飡", True, 1
'                                stbThis.Panels("msg").Text = "��������ʱ����ǰѡ�еĵ�Ԫ�������в����п��кϲ���Ԫ��"
                                mblnInit = False: Exit Function
                            End If
                        Else                            '���·������У����в��ܳ��������һ�кϲ������
                            If lCRow >= lmsRow And lCRow < lmeRow Then
                                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      �������У���ǰ��Ԫ�������в����п��кϲ���Ԫ��" & vbCrLf & "      ���飡", True, 1
                                'stbThis.Panels("msg").Text = "��������ʱ����ǰѡ�еĵ�Ԫ�������в����п��кϲ���Ԫ��"
                                mblnInit = False: Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next
    End If
    
    '�����У���
    Call F1Main.InsertRange(lNRow, lNCol, lNRow, lNCol, IntInsertType)
    If IntInsertType = F1ShiftCols Then '������
        F1Main.MaxCol = F1Main.MaxCol + 1
        For l = 1 To F1Main.MaxCol
            F1Main.ColText(l) = l '��ͷ��ʾ����
        Next
    Else                                '������
        F1Main.MaxRow = F1Main.MaxRow + 1
    End If
    
    '�����洢
    Select Case strType
        Case "InsertLeftCol", "InsertRightCol" '��������
            For lC = lMaxC To 1 Step -1
                If lC >= lNCol Then '����֮�����,��ɾ��������
                    For lR = lMaxR To 1 Step -1
                        Set TmpCell = Document.Cells("K" & lR & "_" & lC)
                        With Document
                            .Cells.Remove ("K" & lR & "_" & lC) '��ɾ��
                            .Cells.Add lR, lC + 1               '��������ȷ�������йؼ���Ψһ
                            With .Cells("K" & lR & "_" & lC + 1)
                                .ID = TmpCell.ID
                                .�ļ�ID = TmpCell.�ļ�ID
                                .������� = TmpCell.�������
                                .�������� = TmpCell.��������
                                .�������� = TmpCell.��������
                                .�������� = TmpCell.��������
                                .�����д� = TmpCell.�����д�
                                .�����ı� = TmpCell.�����ı�
                                .��ʼ�� = TmpCell.��ʼ��
                                .��ֹ�� = TmpCell.��ֹ��
                                .Row = TmpCell.Row
                                .Col = TmpCell.Col + 1
                                .Width = TmpCell.Width
                                .Height = TmpCell.Height
                                .FontName = TmpCell.FontName
                                .FontSize = TmpCell.FontSize
                                .FontBold = TmpCell.FontBold
                                .FontItalic = TmpCell.FontItalic
                                .FontUnderline = TmpCell.FontUnderline
                                .FontStrikeout = TmpCell.FontStrikeout
                                .FontColor = TmpCell.FontColor
                                .HAlignment = TmpCell.HAlignment
                                .VAlignment = TmpCell.VAlignment
                                .CellLineTop = TmpCell.CellLineTop
                                .CellLineBottom = TmpCell.CellLineBottom
                                .CellLineLeft = TmpCell.CellLineLeft
                                .CellLineRight = TmpCell.CellLineRight
                                .CellLineTopColor = TmpCell.CellLineTopColor
                                .CellLineBottomColor = TmpCell.CellLineBottomColor
                                .CellLineLeftColor = TmpCell.CellLineLeftColor
                                .CellLineRightColor = TmpCell.CellLineRightColor
                                .Merge = TmpCell.Merge
                                .MergeRange = TmpCell.MergeRange
                                .TextKey = TmpCell.TextKey
                                .ElementKey = TmpCell.ElementKey
                                .PictureKey = TmpCell.PictureKey
                                .SignKey = TmpCell.SignKey
                                .PicMarkKey = TmpCell.PicMarkKey
                                If .Merge And InStr(.MergeRange, ";") > 0 Then '�ϲ���Ԫ���׸�,�в���,��+1
                                    .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) & "," & Split(Split(.MergeRange, ";")(0), ",")(1) + 1 & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) & "," & Split(Split(.MergeRange, ";")(1), ",")(1) + 1
                                Else
                                    .MergeRange = .Row & "," & .Col
                                End If
                                For j = 0 To UBound(Split(.ElementKey, "|")) '�ı�Ԫ��������ı���
                                    If Len(Split(.ElementKey, "|")(j)) > 0 Then
                                        Document.Elements("K" & Split(.ElementKey, "|")(j)).���� = .Row & "|" & .Col
                                    End If
                                Next
                                For j = 0 To UBound(Split(.TextKey, "|")) '�ı�����������ı���
                                    If Len(Split(.TextKey, "|")(j)) > 0 Then
                                        Document.Texts("K" & Split(.TextKey, "|")(j)).���� = .Row & "|" & .Col
                                    End If
                                Next
                                If .PictureKey <> "" Then
                                    If Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                        If PicEdit.Visible Then PicEdit.Visible = False
                                        Unload PicDy(TmpCell.Index)
                                    End If
                                End If
                            End With
                        End With
                    Next
                End If
            Next
            '�����൥Ԫ��,��Ϊ����������Ϊ��������
            For l = 1 To F1Main.MaxRow
                strKey = Document.Cells.Add(l, lNCol)
                With Document.Cells(strKey)
                    .Height = F1Main.RowHeight(l)
                    .Width = F1Main.ColWidthTwips(Decode(lNCol, F1Main.MaxCol, lNCol - 1, F1Main.MinCol, lNCol + 1, lNCol - 1))
                    .������� = Document.mMaxNo + 1: Document.mMaxNo = Document.mMaxNo + 1
                    If lNCol = F1Main.MaxCol Then '׷�Ӽ̳���
                        .Merge = Document.Cells.Cell(l, lNCol - 1).Merge
                        If InStr(Document.Cells.Cell(l, lNCol - 1).MergeRange, ";") > 0 Then '��Ч�ϲ���Ԫ��
                            .MergeRange = Document.Cells.Cell(l, lNCol - 1).MergeRange '׷�ӵ��У�ֻ���ܳ��ֿ��кϲ����������ʱ���м̳����еĺϲ�����,��+1�в���
                            If Val(Split(Split(.MergeRange, ";")(0), ",")(1)) <> Val(Split(Split(.MergeRange, ";")(1), ",")(1)) Then '�п��кϲ������������ϲ�
                                .MergeRange = l & "," & lNCol
                                .Merge = False
                            Else
                                .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) & "," & Split(Split(.MergeRange, ";")(0), ",")(1) + 1 & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) & "," & Split(Split(.MergeRange, ";")(1), ",")(1) + 1
                                .Merge = True
                            End If
                        Else '�Ǻϲ���Ԫ��ͺϲ�����Ч��Ԫ��
                            .Merge = False
                            .MergeRange = l & "," & lNCol
                        End If
                    Else                         '��ͨ�����в��Ժϲ����д���
                        .MergeRange = l & "," & lNCol
                    End If
                End With
            Next
            F1Main.ColWidthTwips(lNCol) = F1Main.ColWidthTwips(Decode(lNCol, F1Main.MaxCol, lNCol - 1, F1Main.MinCol, lNCol + 1, lNCol - 1))
            Document.Cells.Cols = Document.Cells.Cols + 1
        Case "InsertUpRow", "InsertDnRow" '��������
            For lR = lMaxR To 1 Step -1
                If lR >= lNRow Then '����֮�����,��ɾ��������
                    For lC = lMaxC To 1 Step -1
                        Set TmpCell = Document.Cells("K" & lR & "_" & lC)
                        With Document
                            .Cells.Remove ("K" & lR & "_" & lC) '��ɾ��
                            .Cells.Add lR + 1, lC             '��������ȷ�������йؼ���Ψһ
                            With .Cells("K" & lR + 1 & "_" & lC)
                                .ID = TmpCell.ID
                                .�ļ�ID = TmpCell.�ļ�ID
                                .������� = TmpCell.�������
                                .�������� = TmpCell.��������
                                .�������� = TmpCell.��������
                                .�������� = TmpCell.��������
                                .�����д� = TmpCell.�����д�
                                .�����ı� = TmpCell.�����ı�
                                .��ʼ�� = TmpCell.��ʼ��
                                .��ֹ�� = TmpCell.��ֹ��
                                .Row = TmpCell.Row + 1
                                .Col = TmpCell.Col
                                .Width = TmpCell.Width
                                .Height = TmpCell.Height
                                .FontName = TmpCell.FontName
                                .FontSize = TmpCell.FontSize
                                .FontBold = TmpCell.FontBold
                                .FontItalic = TmpCell.FontItalic
                                .FontUnderline = TmpCell.FontUnderline
                                .FontStrikeout = TmpCell.FontStrikeout
                                .FontColor = TmpCell.FontColor
                                .HAlignment = TmpCell.HAlignment
                                .VAlignment = TmpCell.VAlignment
                                .CellLineTop = TmpCell.CellLineTop
                                .CellLineBottom = TmpCell.CellLineBottom
                                .CellLineLeft = TmpCell.CellLineLeft
                                .CellLineRight = TmpCell.CellLineRight
                                .CellLineTopColor = TmpCell.CellLineTopColor
                                .CellLineBottomColor = TmpCell.CellLineBottomColor
                                .CellLineLeftColor = TmpCell.CellLineLeftColor
                                .CellLineRightColor = TmpCell.CellLineRightColor
                                .Merge = TmpCell.Merge
                                .MergeRange = TmpCell.MergeRange
                                .TextKey = TmpCell.TextKey
                                .ElementKey = TmpCell.ElementKey
                                .PictureKey = TmpCell.PictureKey
                                .SignKey = TmpCell.SignKey
                                .PicMarkKey = TmpCell.PicMarkKey
                                If .Merge And InStr(.MergeRange, ";") > 0 Then '�ϲ���Ԫ���׸�,�в���,��+1
                                    .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) + 1 & "," & Split(Split(.MergeRange, ";")(0), ",")(1) & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) + 1 & "," & Split(Split(.MergeRange, ";")(1), ",")(1)
                                Else
                                    .MergeRange = .Row & "," & .Col
                                End If
                                For j = 0 To UBound(Split(.ElementKey, "|")) '�ı�Ԫ��������ı���
                                    If Len(Split(.ElementKey, "|")(j)) > 0 Then
                                        Document.Elements("K" & Split(.ElementKey, "|")(j)).���� = .Row & "|" & .Col
                                    End If
                                Next
                                For j = 0 To UBound(Split(.TextKey, "|")) '�ı�����������ı���
                                    If Len(Split(.TextKey, "|")(j)) > 0 Then
                                        Document.Texts("K" & Split(.TextKey, "|")(j)).���� = .Row & "|" & .Col
                                    End If
                                Next
                                If .PictureKey <> "" Then
                                    If Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                        If PicEdit.Visible Then PicEdit.Visible = False
                                        Unload PicDy(TmpCell.Index)
                                    End If
                                End If
                            End With
                        End With
                    Next
                End If
            Next
            '�����൥Ԫ��,��Ϊ����������Ϊ��������
            For l = 1 To F1Main.MaxCol
                strKey = Document.Cells.Add(lNRow, l)
                With Document.Cells(strKey)
                    .Width = F1Main.ColWidthTwips(l)
                    .Height = F1Main.RowHeight(Decode(lNRow, F1Main.MaxRow, lNRow - 1, F1Main.MinRow, lNRow + 1, lNRow - 1))
                    .������� = Document.mMaxNo + 1: Document.mMaxNo = Document.mMaxNo + 1
                    If lNRow = F1Main.MaxRow Then '׷�Ӽ̳���
                        .Merge = Document.Cells.Cell(lNRow - 1, l).Merge
                        If InStr(Document.Cells.Cell(lNRow - 1, l).MergeRange, ";") > 0 Then '��Ч�ϲ���Ԫ��
                            .MergeRange = Document.Cells.Cell(lNRow - 1, l).MergeRange '׷�ӵ��У�ֻ���ܳ��ֿ��кϲ����������ʱ���м̳����еĺϲ�����,��+1�в���
                            If Val(Split(Split(.MergeRange, ";")(0), ",")(0)) <> Val(Split(Split(.MergeRange, ";")(1), ",")(0)) Then '�п��кϲ������������ϲ�
                                .Merge = False
                                .MergeRange = lNRow & "," & l
                            Else
                                .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) + 1 & "," & Split(Split(.MergeRange, ";")(0), ",")(1) & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) + 1 & "," & Split(Split(.MergeRange, ";")(1), ",")(1)
                                .Merge = True
                            End If
                        Else '�Ǻϲ���Ԫ��ͺϲ�����Ч��Ԫ��
                            .Merge = False
                            .MergeRange = lNRow & "," & l
                        End If
                    Else                         '��ͨ�����в��Ժϲ����д���
                        .MergeRange = lNRow & "," & l
                    End If
                End With
            Next
            F1Main.RowHeight(lNRow) = F1Main.RowHeight(Decode(lNRow, F1Main.MaxRow, lNRow - 1, F1Main.MinRow, lNRow + 1, lNRow - 1))
            Document.Cells.Rows = Document.Cells.Rows + 1
    End Select
    If editType <> TabET_��������� Then Call RefreshF1Main
    mblnInit = False
    Call F1Main.SetSelection(lngStarRow, lngStarCol, lngEndRow, lngEndCol)
    InsertRowCol = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnInit = False
End Function
Private Sub DeleteRowCol(ByVal strType As String)
'���ܣ�ɾ���л���
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long
Dim lmsRow As Long, lmsCol As Long, lmeRow As Long, lmeCol As Long
Dim l As Long, j As Integer, strDelKey As String, lCellCount As Long, lngDel As Long, strChangeKey As String
Dim lMaxR As Long, lMaxC As Long, lR As Long, lC As Long, TmpCell As cTabCell
    On Error GoTo errHand
    
    mblnInit = True
    Call F1Main.GetSelection(0, lngStarRow, lngStarCol, lngEndRow, lngEndCol) 'ȡ�õ�һ��ѡ��������ֹ����
    lCellCount = Document.Cells.Count: lMaxR = Document.Cells.Rows: lMaxC = Document.Cells.Cols
    
    If strType = "Col" Then 'ɾ���������в��ܳ��ֺϲ���Ԫ��
        lngDel = (lngEndCol - lngStarCol) + 1
    Else
        lngDel = (lngEndRow - lngStarRow) + 1
    End If
    
    If (lngDel = F1Main.MaxRow And strType = "Row") Or (lngDel = F1Main.MaxCol And strType = "Col") Then
        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���棺" & vbCrLf & "      ���ٱ���һ�л�һ��" & vbCrLf & "      ���飡", True, 2
        mblnInit = False
        Exit Sub
    End If
        
    For l = 1 To lCellCount
        With Document.Cells(l)
            If .Merge And InStr(.MergeRange, ";") > 0 Then '�Ժϲ�����������ж�
                lmsRow = Val(Split(Split(.MergeRange, ";")(0), ",")(0)): lmsCol = Val(Split(Split(.MergeRange, ";")(0), ",")(1))
                lmeRow = Val(Split(Split(.MergeRange, ";")(1), ",")(0)): lmeCol = Val(Split(Split(.MergeRange, ";")(1), ",")(1))
                If strType = "Col" Then 'ɾ����
                    If lmsCol <> lmeCol Then '�п��кϲ������
                    For j = lngStarCol To lngEndCol 'ѡ�е����г��ֱ��ϲ����
                        If j >= lmsCol And j <= lmeCol Then
                            mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      ɾ����ʱ����ǰ��Ԫ�������в����п��кϲ���Ԫ��" & vbCrLf & "      ���飡", True, 1
                            mblnInit = False: Exit Sub
                        End If
                    Next
                    End If
                Else                    'ɾ����
                    If lmsRow <> lmeRow Then '�п��кϲ������
                    For j = lngStarRow To lngEndRow 'ѡ�е����г��ֱ��ϲ������
                        If j >= lmsRow And j <= lmeRow Then
                            mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      ɾ����ʱ����ǰ��Ԫ�������в����п��кϲ���Ԫ��" & vbCrLf & "      ���飡", True, 1
                            mblnInit = False: Exit Sub
                        End If
                    Next
                    End If
                End If
            End If
        End With
    Next
    
    'ɾ���У���
    Call F1Main.DeleteRange(lngStarRow, lngStarCol, lngEndRow, lngEndCol, Decode(strType, "Row", F1ShiftRows, "Col", F1ShiftCols))
    If strType = "Col" Then '������
        F1Main.MaxCol = F1Main.MaxCol - lngDel
        For l = 1 To F1Main.MaxCol
            F1Main.ColText(l) = l '��ͷ��ʾ����
        Next
    Else                    '������
        F1Main.MaxRow = F1Main.MaxRow - lngDel
    End If
    
    '������ѭ����ɾ�������е����Ա����Ϊ��ı伯������������˳���ȼ���Keyֵ
    For l = 1 To lCellCount
        With Document.Cells(l)
            If strType = "Row" Then
                If .Row >= lngStarRow And .Row <= lngEndRow Then strDelKey = strDelKey & "|" & .Key
            Else
                If .Col >= lngStarCol And .Col <= lngEndCol Then strDelKey = strDelKey & "|" & .Key
            End If
        End With
    Next
    'Ҫɾ�������Ա,ͬʱɾ�����Ա��������Ա
    For l = 1 To UBound(Split(strDelKey, "|"))
        Call ClearChildMember(Split(strDelKey, "|")(l))
        Call Document.Cells.Remove(Split(strDelKey, "|")(l))
    Next
    
    If strType = "Row" Then
        For lR = 1 To lMaxR
            If lR > lngEndRow Then
                For lC = 1 To lMaxC
                    Set TmpCell = Document.Cells("K" & lR & "_" & lC)
                    With Document
                        .Cells.Remove TmpCell.Key
                        .Cells.Add lR - lngDel, lC
                        With .Cells("K" & lR - lngDel & "_" & lC)
                            .ID = TmpCell.ID
                            .�ļ�ID = TmpCell.�ļ�ID
                            .������� = TmpCell.�������
                            .�������� = TmpCell.��������
                            .�������� = TmpCell.��������
                            .�������� = TmpCell.��������
                            .�����д� = TmpCell.�����д�
                            .�����ı� = TmpCell.�����ı�
                            .��ʼ�� = TmpCell.��ʼ��
                            .��ֹ�� = TmpCell.��ֹ��
                            .Row = lR - lngDel
                            .Col = lC
                            .Width = TmpCell.Width
                            .Height = TmpCell.Height
                            .FontName = TmpCell.FontName
                            .FontSize = TmpCell.FontSize
                            .FontBold = TmpCell.FontBold
                            .FontItalic = TmpCell.FontItalic
                            .FontUnderline = TmpCell.FontUnderline
                            .FontStrikeout = TmpCell.FontStrikeout
                            .FontColor = TmpCell.FontColor
                            .HAlignment = TmpCell.HAlignment
                            .VAlignment = TmpCell.VAlignment
                            .CellLineTop = TmpCell.CellLineTop
                            .CellLineBottom = TmpCell.CellLineBottom
                            .CellLineLeft = TmpCell.CellLineLeft
                            .CellLineRight = TmpCell.CellLineRight
                            .CellLineTopColor = TmpCell.CellLineTopColor
                            .CellLineBottomColor = TmpCell.CellLineBottomColor
                            .CellLineLeftColor = TmpCell.CellLineLeftColor
                            .CellLineRightColor = TmpCell.CellLineRightColor
                            .Merge = TmpCell.Merge
                            .MergeRange = TmpCell.MergeRange
                            .TextKey = TmpCell.TextKey
                            .ElementKey = TmpCell.ElementKey
                            .PictureKey = TmpCell.PictureKey
                            .SignKey = TmpCell.SignKey
                            .PicMarkKey = TmpCell.PicMarkKey
                            If .Merge And InStr(.MergeRange, ";") > 0 Then '�ϲ���Ԫ���׸�
                                .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) - lngDel & "," & Split(Split(.MergeRange, ";")(0), ",")(1) & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) - lngDel & "," & Split(Split(.MergeRange, ";")(1), ",")(1)
                            Else
                                .MergeRange = .Row & "," & .Col
                            End If
                            For j = 0 To UBound(Split(.ElementKey, "|")) '�ı�Ԫ��������ı���
                                If Len(Split(.ElementKey, "|")(j)) > 0 Then
                                    Document.Elements("K" & Split(.ElementKey, "|")(j)).���� = .Row & "|" & .Col
                                End If
                            Next
                            For j = 0 To UBound(Split(.TextKey, "|")) '�ı�����������ı���
                                If Len(Split(.TextKey, "|")(j)) > 0 Then
                                    Document.Texts("K" & Split(.TextKey, "|")(j)).���� = .Row & "|" & .Col
                                End If
                            Next
                            If .PictureKey <> "" Then
                                If Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                    If PicEdit.Visible Then PicEdit.Visible = False
                                    Unload PicDy(TmpCell.Index)
                                End If
                            End If
                        End With
                    End With
                Next
            End If
        Next
    Else
        For lC = 1 To lMaxC      '��ɾ���С���֮������е�Ԫ����ɾ������������ȷ�������еĹؼ�����ROW/COL��Ӧ
            If lC > lngEndCol Then
                For lR = 1 To lMaxR
                    Set TmpCell = Document.Cells("K" & lR & "_" & lC)
                    With Document
                        .Cells.Remove TmpCell.Key
                        .Cells.Add lR, lC - lngDel
                        With .Cells("K" & lR & "_" & lC - lngDel)
                            .ID = TmpCell.ID
                            .�ļ�ID = TmpCell.�ļ�ID
                            .������� = TmpCell.�������
                            .�������� = TmpCell.��������
                            .�������� = TmpCell.��������
                            .�������� = TmpCell.��������
                            .�����д� = TmpCell.�����д�
                            .�����ı� = TmpCell.�����ı�
                            .��ʼ�� = TmpCell.��ʼ��
                            .��ֹ�� = TmpCell.��ֹ��
                            .Row = lR
                            .Col = lC - lngDel
                            .Width = TmpCell.Width
                            .Height = TmpCell.Height
                            .FontName = TmpCell.FontName
                            .FontSize = TmpCell.FontSize
                            .FontBold = TmpCell.FontBold
                            .FontItalic = TmpCell.FontItalic
                            .FontUnderline = TmpCell.FontUnderline
                            .FontStrikeout = TmpCell.FontStrikeout
                            .FontColor = TmpCell.FontColor
                            .HAlignment = TmpCell.HAlignment
                            .VAlignment = TmpCell.VAlignment
                            .CellLineTop = TmpCell.CellLineTop
                            .CellLineBottom = TmpCell.CellLineBottom
                            .CellLineLeft = TmpCell.CellLineLeft
                            .CellLineRight = TmpCell.CellLineRight
                            .CellLineTopColor = TmpCell.CellLineTopColor
                            .CellLineBottomColor = TmpCell.CellLineBottomColor
                            .CellLineLeftColor = TmpCell.CellLineLeftColor
                            .CellLineRightColor = TmpCell.CellLineRightColor
                            .Merge = TmpCell.Merge
                            .MergeRange = TmpCell.MergeRange
                            .TextKey = TmpCell.TextKey
                            .ElementKey = TmpCell.ElementKey
                            .PictureKey = TmpCell.PictureKey
                            .SignKey = TmpCell.SignKey
                            .PicMarkKey = TmpCell.PicMarkKey
                            If .Merge And InStr(.MergeRange, ";") > 0 Then '�ϲ���Ԫ���׸�
                                .MergeRange = Split(Split(.MergeRange, ";")(0), ",")(0) & "," & Split(Split(.MergeRange, ";")(0), ",")(1) - lngDel & ";" & Split(Split(.MergeRange, ";")(1), ",")(0) & "," & Split(Split(.MergeRange, ";")(1), ",")(1) - lngDel
                            Else
                                .MergeRange = .Row & "," & .Col
                            End If
                            For j = 0 To UBound(Split(.ElementKey, "|")) '�ı�Ԫ��������ı���
                                If Len(Split(.ElementKey, "|")(j)) > 0 Then
                                    Document.Elements("K" & Split(.ElementKey, "|")(j)).���� = .Row & "|" & .Col
                                End If
                            Next
                            For j = 0 To UBound(Split(.TextKey, "|")) '�ı�����������ı���
                                If Len(Split(.TextKey, "|")(j)) > 0 Then
                                    Document.Texts("K" & Split(.TextKey, "|")(j)).���� = .Row & "|" & .Col
                                End If
                            Next
                            If .PictureKey <> "" Then
                                If Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                    If PicEdit.Visible Then PicEdit.Visible = False
                                    Unload PicDy(TmpCell.Index)
                                End If
                            End If
                        End With
                    End With
                Next
            End If
        Next
    End If
    
    '����༯������������
    If strType = "Row" Then
        Document.Cells.Rows = Document.Cells.Rows - lngDel
        lngStarRow = lngStarRow - lngDel: lngEndRow = lngEndRow - lngDel
    Else
        Document.Cells.Cols = Document.Cells.Cols - lngDel
        lngStarCol = lngStarCol - lngDel: lngEndCol = lngEndCol - lngDel
    End If
    
    If lngStarRow < 1 Then lngStarRow = 1: If lngEndRow < 1 Then lngEndRow = 1
    If lngStarCol < 1 Then lngStarCol = 1: If lngEndCol < 1 Then lngEndCol = 1
    Call RefreshF1Main
    mblnInit = False
    Call F1Main.SetSelection(lngStarRow, lngStarCol, lngEndRow, lngEndCol)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InsertInherit(ByVal strType As String)
'���ܣ�����̳�����
Dim lngStarRow As Long, lngEndRow As Long, lngStarCol As Long, lngEndCol As Long, lngRow As Long, lngCol As Long, cbrObj As CommandBarControl
Dim vCell As F1CellFormat, l As Long, j As Integer, strDelKey As String, lngCellCount As Long, lngDel As Long, strChangeKey As String
    On Error GoTo errHand
    mblnInit = True
    Call F1Main.SetSelection(F1Main.MaxRow, F1Main.MaxCol, F1Main.MaxRow, F1Main.MaxCol)
    If strType = "Row" Then
        If Not InsertRowCol("InsertDnRow") Then Exit Sub    '�Ȳ����У���������·����ұ�
        mblnInit = True
        For l = 1 To F1Main.MaxCol '���ݸ���,�ȸ��Ƶ�Ԫ��(�����¼���Ա��Key),�ٸ����¼���ԱKey�����¼���Ա����
            Call Document.Cells.Cell(F1Main.MaxRow - 1, l).Clone(Document.Cells.Cell(F1Main.MaxRow, l))
            With Document.Cells.Cell(F1Main.MaxRow, l)
                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                    Call CloneChildMember(.Key)
                End If
            End With
'           ��InsertRowCol���Ѵ���� Document.Cells.Cell(F1Main.MaxRow, l).������� = Document.mMaxNo + 1: Document.mMaxNo = Document.mMaxNo '׷�������ڱ༭ʱ���������Ӧȡ���ֵ
        Next
    Else
        If Not InsertRowCol("InsertRightCol") Then Exit Sub '�Ȳ����У���������·����ұ�
        mblnInit = True
        For l = 1 To F1Main.MaxRow '���ݸ���,�ȸ��Ƶ�Ԫ��(�����¼���Ա��Key),�ٸ����¼���ԱKey�����¼���Ա����
            Call Document.Cells.Cell(l, F1Main.MaxCol - 1).Clone(Document.Cells.Cell(l, F1Main.MaxCol))
            With Document.Cells.Cell(l, F1Main.MaxCol)
                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                    Call CloneChildMember(.Key)
                End If
            End With
'           ��InsertRowCol���Ѵ���� Document.Cells.Cell(l, F1Main.MaxCol).������� = Document.mMaxNo + 1: Document.mMaxNo = Document.mMaxNo
        Next
    End If
    mblnInit = False
    Set cbrObj = cbsMain.FindControl(, ID_FILE_SAVE)
    mblnAdd = True
    If cbrObj.Enabled And cbrObj.Visible Then cbsMain_Execute cbrObj
    mblnAdd = False
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnInit = False
End Sub
Private Sub CloneChildMember(ByVal strCellKey As String)
'����:����ָ����Ԫ����¼���Ա��strCellKey��ʾ��Ա���ڵ�Ԫ���Key
Dim i As Integer, lngKey As Long, strNewKey As String, vCell As F1CellFormat, strShow As String
    With Document.Cells(strCellKey)
        For i = 1 To UBound(Split(.TextKey, "|"))
            lngKey = Document.Texts.Add: strNewKey = strNewKey & "|" & lngKey
            Call Document.Texts("K" & Split(.TextKey, "|")(i)).Clone(Document.Texts("K" & lngKey))
            Document.Texts("K" & lngKey).��ID = .ID
            Document.Texts("K" & lngKey).��ʼ�� = 1
            Document.Texts("K" & lngKey).��ֹ�� = 0
            Document.Texts("K" & lngKey).���� = .Row & "|" & .Col
        Next
        .TextKey = strNewKey: strNewKey = ""
        
        If .ElementKey <> "" Then
            For i = 0 To UBound(Split(.ElementKey, "|"))
                If Split(.ElementKey, "|")(i) = "" Then '�����ǻ������,��������1��ʼ
                    lngKey = Document.Elements.Add: strNewKey = strNewKey & "|" & lngKey
                    Call Document.Elements("K" & Split(.ElementKey, "|")(i)).Clone(Document.Elements("K" & lngKey))
                    Document.Elements("K" & lngKey).��ID = .ID
                    Document.Elements("K" & lngKey).���� = .Row & "|" & .Col
                    Document.Elements("K" & lngKey).��ʼ�� = 1
                    Document.Elements("K" & lngKey).��ֹ�� = 0
                End If
            Next
        End If
        .ElementKey = strNewKey: strNewKey = ""
        
        For i = 1 To UBound(Split(.PicMarkKey, "|"))
            lngKey = Document.PicMarks.Add: strNewKey = strNewKey & "|" & lngKey
            Call Document.PicMarks("K" & Split(.PicMarkKey, "|")(i)).Clone(Document.PicMarks("K" & lngKey))
            Document.PicMarks("K" & lngKey).��ID = .ID
            Document.PicMarks("K" & lngKey).��ʼ�� = 1
            Document.PicMarks("K" & lngKey).��ֹ�� = 0
        Next
        .PicMarkKey = strNewKey: strNewKey = ""
        
        If .PictureKey <> "" Then
            lngKey = Document.Pictures.Add: strNewKey = lngKey
            Call Document.Pictures("K" & .PictureKey).Clone(Document.Pictures("K" & lngKey))
            Document.Pictures("K" & lngKey).PicID = .ID
        End If
        .PictureKey = strNewKey: strNewKey = ""
        
        If .SignKey <> "" Then
            lngKey = Document.Signs.Add: strNewKey = lngKey
            Call Document.Signs("K" & .SignKey).Clone(Document.Signs("K" & lngKey))
        End If
        .SignKey = strNewKey: strNewKey = ""
        
        '��Ա���ݻ�ͼƬ
        Select Case .��������
            Case cprCTPicture, cprCTReportPic
                strShow = IIf(.�������� = cprCTReportPic, "����ͼ", "�ο�ͼ")
                .�����ı� = strShow
                PaintPictureOnTable strCellKey
            Case cprCTSign, cprCTColSign, cprCTRowSign
                strShow = "[ǩ��λ]"
            Case Else
                strShow = .�����ı�
                DocOld.Add .�����ı�, .Key
        End Select
        .��ʼ�� = 1: .��ֹ�� = 0
        F1Main.TextRC(.Row, .Col) = strShow
        Call F1Main.SetSelection(.Row, .Col, .Row, .Col)
        Set vCell = F1Main.GetCellFormat
        vCell.FontName = .FontName: vCell.FontSize = .FontSize: vCell.FontBold = .FontBold: vCell.FontColor = .FontColor
        vCell.FontItalic = .FontItalic: vCell.FontUnderline = .FontUnderline: vCell.FontStrikeout = .FontStrikeout
        vCell.BorderStyle(F1TopBorder) = .CellLineTop: vCell.BorderStyle(F1BottomBorder) = .CellLineBottom
        vCell.BorderStyle(F1LeftBorder) = .CellLineLeft: vCell.BorderStyle(F1RightBorder) = .CellLineRight
        vCell.BorderColor(F1TopBorder) = .CellLineTopColor: vCell.BorderColor(F1BottomBorder) = .CellLineBottomColor
        vCell.BorderColor(F1LeftBorder) = .CellLineLeftColor: vCell.BorderColor(F1RightBorder) = .CellLineRightColor
        Call F1Main.SetCellFormat(vCell)
    End With
End Sub

Private Sub ClearChildMember(ByVal strCellKey As String)
'����:���ָ����Ԫ���¼���Ա,��ɾ����Ա��ֻ����Key������Ϊ��ı�����˳��
Dim j As Long
    With Document.Cells(strCellKey)
        .TextKey = ""
        .ElementKey = ""
        .PicMarkKey = ""
        
        If .PictureKey <> "" Then
            If Document.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                If PicEdit.Visible Then PicEdit.Visible = False
                Unload PicDy(.Index)
            End If
            .PictureKey = ""
        End If
        
        .SignKey = ""
        
        .�������� = 0: .�������� = "": .�����д� = 0: .�����ı� = ""
    End With
End Sub
Public Sub PaintPictureOnTable(ByVal strCellKey As String)
'����:��ָ����Ԫ���ͼ
Dim objTmp As Object, vR As F1Rect, i As Integer, lHheight As Long, lHwidth As Long, lpLeft As Long, lpTop As Long 'ͼƬ��,����,�̶��и߶�,�̶��п��,ͼƬ��XY����
Dim lsRow As Long, leRow As Long, lsCol As Long, leCol As Long '������ֹ����
Dim lsPosX As Long, lsPosY As Long, lpHeight As Long, lpWidth As Long 'ͼƬԴ����XY����,ͼƬ�߿�

    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '�̶��и߶�
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '�̶��п��

    With Document.Cells(strCellKey)
        If .PictureKey = "" Then Exit Sub
        If Document.Pictures("K" & .PictureKey).OrigPic.Handle = 0 Then Exit Sub
        
        'ȷ��ͼƬ����������
        If .Merge Then  'MergeRange���ݸ�ʽ (���Ϸ�)��,��;(���·�)��,��
            lsRow = Split(Split(.MergeRange, ";")(0), ",")(0): leRow = Split(Split(.MergeRange, ";")(1), ",")(0)
            lsCol = Split(Split(.MergeRange, ";")(0), ",")(1): leCol = Split(Split(.MergeRange, ";")(1), ",")(1)
        Else
            lsRow = .Row: leRow = .Row: lsCol = .Col: leCol = .Col
        End If
        Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
        'ȷ��ͼƬ���С��λ�ü��ü�����
        If vR.Right - lHwidth <= 0 Or vR.Bottom - lHheight <= 0 Then '���ڿ���ʾ����
            If ChkControl(PicDy(.Index)) Then
                PicDy(.Index).Visible = False
            End If
            Exit Sub
        ElseIf vR.Left >= 0 And vR.Top >= 0 Then '�����ڱ���м�
            lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + vR.Top: lpWidth = vR.Width: lpHeight = vR.Height: lsPosX = 0: lsPosY = 0
        ElseIf vR.Left >= 0 And vR.Top < 0 Then '�����Ϸ���������(��������)
            lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + lHheight: lpWidth = vR.Width: lpHeight = vR.Height + vR.Top - lHheight: lsPosX = 0: lsPosY = vR.Height - lpHeight
        ElseIf vR.Left < 0 And vR.Top >= 0 Then '�����󷽲�������(��������)
            lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + vR.Top: lpWidth = vR.Width + vR.Left - lHwidth: lpHeight = vR.Height: lsPosX = vR.Width - lpWidth: lsPosY = 0
        ElseIf vR.Left < 0 And vR.Top < 0 Then '�����Ϸ��󷽶�����(��������)
            lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + lHheight: lpWidth = vR.Width + vR.Left - lHwidth: lpHeight = vR.Height + vR.Top - lHheight: lsPosX = vR.Width - lpWidth: lsPosY = vR.Height - lpHeight
        Else                                    '���ڿ���ʾ����
            If ChkControl(PicDy(.Index)) Then
                PicDy(.Index).Visible = False
            End If
            Exit Sub
        End If
        
        '��̬����ͼƬ������
        If Not ChkControl(PicDy(.Index)) Then
            Load PicDy(.Index)
        End If
        Set objTmp = PicDy(.Index): objTmp.Cls
        objTmp.Tag = .MergeRange & "|" & strCellKey: objTmp.ToolTipText = IIf(.�������� = cprCTReportPic, "����ͼ", "�ο�ͼ")
        objTmp.AutoRedraw = True: objTmp.BorderStyle = 0: Set objTmp.Container = picMainBack
        
        '�ȕ���ͼƬ��С��������
        LockWindowUpdate Me.hWnd
        objTmp.Width = vR.Width - Screen.TwipsPerPixelX * 2: objTmp.Height = vR.Height - Screen.TwipsPerPixelY * 2
        Set objTmp.Picture = Document.Pictures("K" & .PictureKey).OrigPic
        objTmp.PaintPicture objTmp.Picture, 0, 0, objTmp.Width, objTmp.Height
        If .PicMarkKey <> "" Then '�б��ͼ�Ȼ���
            For i = 1 To UBound(Split(.PicMarkKey, "|"))
                ShowPicMark objTmp, Me.Document.PicMarks("K" & Split(.PicMarkKey, "|")(i))
            Next
        End If
        Set objTmp.Picture = objTmp.Image
        '������ʵ����ʾ��С�������ػ�
        objTmp.Move lpLeft + Screen.TwipsPerPixelX * 2, lpTop + Screen.TwipsPerPixelY * 2, lpWidth - Screen.TwipsPerPixelX * 2, lpHeight - Screen.TwipsPerPixelY * 2
        objTmp.PaintPicture objTmp.Picture, 0, 0, objTmp.Width, objTmp.Height, lsPosX, lsPosY
        objTmp.Visible = True: objTmp.ZOrder
        LockWindowUpdate 0
    End With
End Sub
Private Sub SetCellFont()
Dim tmpFont As New StdFont, tmpColor As Long
'ͨ�����崰����������
    On Error GoTo errHand
    tmpFont.Name = SelCell.FontName: tmpFont.Size = SelCell.FontSize: tmpFont.Bold = SelCell.FontBold
    tmpFont.Italic = SelCell.FontItalic: tmpFont.Underline = SelCell.FontUnderline: tmpFont.Strikethrough = SelCell.FontStrikeout
    tmpColor = SelCell.FontColor
    If SetFont(Me.hWnd, Me.hdc, tmpFont, tmpColor) Then
        Call SetCellFormat("�ֺ�", tmpFont.Size)
        Call SetCellFormat("��������", tmpFont.Name)
        Call SetCellFormat("����", tmpFont.Bold)
        Call SetCellFormat("б��", tmpFont.Italic)
        Call SetCellFormat("�»���", tmpFont.Underline)
        Call SetCellFormat("ɾ����", tmpFont.Strikethrough)
        Call SetCellFormat("������ɫ", tmpColor)
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub DesignTest()
'���ܣ��˹�����������ƺ͵��Խ׶β鿴�ڴ�������������
Dim picTmp As New StdPicture
    On Error GoTo errHand
       Debug.Print SelCell.�����ı�
       
       If App.LogMode = 0 Then MsgBox "����": Stop
       

       
       
       
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function AllowEdit(nRow As Long, nCol As Long) As Boolean
'���ܣ��ж��Ƿ��ڱ���״̬,��ֹ�༭��������ʾ
    On Error GoTo errHand
    Select Case mReadOnly
        Case 1
            mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      ��ǰ�༭״̬�������ڻ���" & vbCrLf & "      ���ܱ����ǩ����", True, 1
            Exit Function
        Case 2
            mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      ��ǰ���ڲ鿴״̬" & vbCrLf & "      ���ܱ༭��", True, 1
            Exit Function
    End Select
    
    If Document.Cells.Cell(nRow, nCol).�������� = True Or Document.Cells.Cell(nRow, nCol).�������� = cprCTFixtext Then
        Call Beep
        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      ��ѡ�����ڱ����Ŀ�������" & vbCrLf & "      ���ܱ༭��", True, 1
        Exit Function
    End If
    
    Dim lRows As Long, lCols As Long, l As Long
    lRows = Document.Cells.Rows '����������Ƿ����п�ǩ����ǩ��
    For l = 1 To lRows
        With Document.Cells.Cell(l, nCol)
            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                If .�������� = cprCTColSign And .��ֹ�� <> 0 Then
                    Call Beep
                    mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      ��ѡ��������ǩ���Ŀ�������" & vbCrLf & "      ���ܱ༭��", True, 1
                    Exit Function
                End If
            End If
        End With
    Next
    
    lCols = Document.Cells.Cols '����������Ƿ����п�ǩ����ǩ��
    For l = 1 To lCols
        With Document.Cells.Cell(nRow, l)
            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                If .�������� = cprCTRowSign And .��ֹ�� <> 0 Then
                    Call Beep
                    mfrmTipInfo.ShowTipInfo stbThis.hWnd, "���ѣ�" & vbCrLf & "      ��ѡ��������ǩ���Ŀ�������" & vbCrLf & "      ���ܱ༭��", True, 1
                    Exit Function
                End If
            End If
        End With
    Next
    
    AllowEdit = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub EnterEdit(ByVal lsRow As Long, ByVal lsCol As Long, ByVal leRow As Long, ByVal leCol As Long, Optional KeyAscii As Integer, Optional DbClick As Boolean)
'���ܣ�����༭״̬
Dim vR As F1Rect
    On Error GoTo errHand
    If editType = TabET_�������༭ Or editType = TabET_��������� Then
        If Not AllowEdit(lsRow, lsCol) Then F1Main.AllowInCellEditing = False: Exit Sub
    End If
    
    If Document.Cells.Cell(lsRow, lsCol).�������� = cprCTText Or Document.Cells.Cell(lsRow, lsCol).�������� = cprCTFixtext Then
        F1Main.AllowInCellEditing = True
    Else
        F1Main.AllowInCellEditing = False
    End If
    
    With Document.Cells.Cell(lsRow, lsCol)
        If .Merge Then
            Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
        Else
            Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, lsRow, lsCol)
        End If
        Select Case .��������
            Case cprCTText, cprCTFixtext
                If DbClick Then
                    Call F1Main.StartEdit(False, True, False)
                Else
                    Call F1Main.StartEdit(True, True, False)
                End If
            Case cprCTElement
                If Document.Elements("K" & .ElementKey).Ҫ������ = "" Then
                    If (editType = TabET_�����ļ����� Or TabET_ȫ��ʾ���༭) Then
                        Call InsertElement(.Key)
                    End If
                Else
                    If (editType = TabET_�����ļ����� Or TabET_ȫ��ʾ���༭) And Not DbClick And KeyAscii = 0 Then
                        Call InsertElement(.Key)
                    Else
                        Call EditElement(.Key, KeyAscii)
                    End If
                End If
            Case cprCTTextElement
                Call PopDoc(.Key, KeyAscii)
            Case cprCTPicture
                If KeyAscii <> 0 Then KeyAscii = 0: Exit Sub '���������ͼƬ��Ч
                If .PictureKey = "" Then
                    Call InsertPicture(.Key)
                Else '����ͼƬ�༭
                    If Document.Pictures("K" & .PictureKey).OrigPic = 0 Then
                        Call InsertPicture(.Key, .PictureKey)
                    End If
                End If
            Case cprCTReportPic
                If KeyAscii <> 0 Then KeyAscii = 0: Exit Sub '���������ͼƬ��Ч
                If Not dkpMain.FindPane(conPane_PacsPic) Is Nothing Then
                    If Not dkpMain.FindPane(conPane_PacsPic).Closed Then dkpMain.ShowPane conPane_PacsPic
                End If
                Call F1Main.SetFocus
            Case cprCTSign, cprCTColSign, cprCTRowSign
                If KeyAscii <> 0 And KeyAscii <> vbKeySpace Then KeyAscii = 0: Exit Sub '���������ͼƬ��Ч
                If Not (editType = TabET_�����ļ����� Or TabET_ȫ��ʾ���༭) And mReadOnly = 0 Then
                    If SaveDoc(True, True) Then Unload Me
                End If
        End Select
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InsertPicture(ByVal strCellKey As String, Optional ByVal strPicKey As String)
'���ܣ�����ͼƬ,��strPicKey<>""ʱ��ʾ������ǰͼƬ(�ɹ���������)
'��ʱ�Ѿ�ȷ��ѡ�е�ֻ��һ��CELL������ͼƬ����
Dim tmpPic As StdPicture, lngKey As Long, vR As F1Rect, l As Long, ary As Variant
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    On Error GoTo errHand
    
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
    If frmInsertPicture.ShowMe(Me, vR.Width, vR.Height, tmpPic) Then
        AddUndo Document.Cells(strCellKey)
        If strPicKey = "" Then
            lngKey = Document.Pictures.Add
        Else
            lngKey = Val(strPicKey) '��ͼʱKeyֵ����
            If ChkControl(PicDy(Document.Cells(strCellKey).Index)) Then Unload PicDy(Document.Cells(strCellKey).Index)
        End If
        Set Document.Pictures("K" & lngKey).OrigPic = tmpPic '����ͼƬ
        Document.Pictures("K" & lngKey).DesHeight = vR.Height
        Document.Pictures("K" & lngKey).DesWidth = vR.Width
        Document.Cells(strCellKey).PictureKey = lngKey
        Call PaintPictureOnTable(strCellKey) '�ػ�ͼƬ�ͱ��
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InsertOtherText(ByVal strCellKey As String, ByVal strType As String)
'���ܣ��������ڣ�ʱ�䣬������ţ�Ŀ�굥Ԫ�������Text�ͻ�ϱ༭����
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, vR As F1Rect, lHheight As Long, lHwidth As Long, elTmpKey As Long
    On Error GoTo errHand
    If editType = TabET_�������༭ Or editType = TabET_��������� Then
        If Not AllowEdit(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) Then Exit Sub
    End If
    If strType = "�������" Then
        Dim strTmp As String
        strTmp = frmInsSymbol.ShowMe(Decode(mstrSex, "��", 1, "Ů", 2, 0))
        If strTmp <> "" Then
            With Document.Cells(strCellKey) 'ֻ���ı��ͻ��������Բ���
                AddUndo Document.Cells(strCellKey)
                If .�������� = cprCTText Or .�������� = cprCTFixtext Then
                    strTmp = .�����ı� & strTmp
                    .�����ı� = strTmp
                    F1Main.TextRC(.Row, .Col) = strTmp
                Else
                    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
                    bInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '���ڹؼ��ֶ�֮��
                    If bInKeys Then
                        Doc.Range(lEE, lEE).Selected
                        Doc.Range(lEE, lEE).Font.Protected = False
                        Doc.Range(lEE, lEE).Font.Hidden = False
                        Doc.Range(lEE, lEE).Text = IIf(Mid(strTmp, 1, 1) <> "��", "��" & strTmp, strTmp)
                        Doc.Range(lEE + Len(strTmp), lEE + Len(strTmp)).Selected
                    Else
                        If Doc.Range(Doc.Selection.StartPos - 1, Doc.Selection.StartPos).Font.Hidden Then strTmp = "��" & strTmp '����ڹؼ�֮��
                        Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Selected
                        Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Font.Protected = False
                        Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Font.Hidden = False
                        Doc.Range(Doc.Selection.StartPos, Doc.Selection.StartPos).Text = strTmp
                        Doc.Range(Doc.Selection.StartPos + Len(strTmp), Doc.Selection.StartPos + Len(strTmp)).Selected
                    End If
                End If
            End With
        End If
    Else
        elTmpKey = Document.Elements.Add '��ʼ��һ������ʱҪ��
        With Document.Elements("K" & elTmpKey)
               .�����ı� = ""
               .�����д� = 0
               .����Ҫ��ID = 0
               .�滻�� = 0
               .Ҫ������ = strType
               .Ҫ������ = 2
               .Ҫ�س��� = Decode(strType, "����ʱ��", 19, "����", 10, "ʱ��", 8, 19)
               .Ҫ��С�� = 0
               .Ҫ�ص�λ = ""
               .Ҫ�ر�ʾ = 0
               .������̬ = 0
               .Ҫ��ֵ�� = "0;0"
               .�������� = False
               .�Զ�ת�ı� = True
               .���� = 0
        End With
        If Doc.Visible Then
            Call ShowElInDoc(Doc.Selection.StartPos, Doc.Selection.StartPos, elTmpKey)
        Else
            Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
            If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '�̶��и߶�
            If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '�̶��п��
            Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
            
            elEdit.Move F1Main.Left + IIf(vR.Left < 0, lHwidth, vR.Left) + Screen.TwipsPerPixelX * 2, F1Main.Top + vR.Bottom + Screen.TwipsPerPixelX * 2: elEdit.Tag = strCellKey
            elEdit.SetElement Document.Elements("K" & elTmpKey), 0, editType
            
            If elEdit.Top + elEdit.Height > F1Main.Top + F1Main.Height Then
                elEdit.Top = vR.Top - elEdit.Height - Screen.TwipsPerPixelY * 2
            End If
            
            If elEdit.Left + elEdit.Width > F1Main.Left + F1Main.Width Then
                elEdit.Left = vR.Left - elEdit.Width - Screen.TwipsPerPixelX * 2
            End If
            elEdit.Visible = True: elEdit.ZOrder 0: elEdit.SetFocus
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub EditPicture(ByVal strKey As String)
'ͼƬ�༭,strKeyͼƬ���ڵ�Ԫ��KEY
Dim tmpPic As StdPicture, lngKey As Long, vR As F1Rect
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long
    On Error GoTo errHand
    If PicEdit.Visible Then PicEdit_LostFocus  '������ͼ֮���л�ʱ���ȶ�ʧ�����Ա��ػ�ԭͼ
    With Document.Cells(strKey)
        If .�������� = True Then Exit Sub
        If .PictureKey = "" Then Exit Sub
        If Document.Pictures("K" & .PictureKey).OrigPic.Handle = 0 Then Exit Sub
    'ȷ��ͼƬ����������
        If .Merge Then  'MergeRange���ݸ�ʽ (���Ϸ�)��,��;(���·�)��,��
            lsRow = Split(Split(.MergeRange, ";")(0), ",")(0): leRow = Split(Split(.MergeRange, ";")(1), ",")(0)
            lsCol = Split(Split(.MergeRange, ";")(0), ",")(1): leCol = Split(Split(.MergeRange, ";")(1), ",")(1)
        Else
            lsRow = .Row: leRow = .Row: lsCol = .Col: leCol = .Col
        End If
    End With
    AddUndo Document.Cells(strKey)
    Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
    If vR.Top < -15 Or vR.Left < -15 Then Exit Sub '���ݳ�����������,��ֹ�༭
    PicEdit.Move F1Main.Left + vR.Left + Screen.TwipsPerPixelX * 2, F1Main.Top + vR.Top + Screen.TwipsPerPixelY * 2, vR.Width - Screen.TwipsPerPixelX * 2, vR.Height - Screen.TwipsPerPixelY * 2
    PicEdit.Tag = IIf(Document.Cells(strKey).�������� = cprCTPicture, "�ο�ͼ", "����ͼ")
    Call PicEdit.EditPic(Document, cbsMain, strKey)
    PicEdit.Visible = True: PicEdit.SetFocus: PicEdit.ZOrder 0
    '�����˵�
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InsertElement(ByVal strCellKey As String)
'���ܣ� ����Ҫ�أ���Ҫ������Ϊ��ʱ������Ҫ�أ��ǿ�ʱ�޸�Ҫ��
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, mfrmInsElement As New frmInsElement, strShow As String
    On Error GoTo errHand
    With Document.Cells(strCellKey)
        If .�������� = cprCTElement Then
             strShow = F1Main.TextRC(.Row, .Col)
            If mfrmInsElement.ShowMe(Me, Document.Elements("K" & .ElementKey), True, True) Then
                If Document.Elements("K" & .ElementKey).������̬ = 1 And Document.Elements("K" & .ElementKey).Ҫ������ <> 2 Then
                    strShow = Document.Elements("K" & .ElementKey).�����ı� & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                    .�����ı� = Document.Elements("K" & .ElementKey).�����ı�
                Else
                    strShow = "[" & Document.Elements("K" & .ElementKey).Ҫ������ & "]" & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                End If
            End If
            F1Main.TextRC(.Row, .Col) = strShow
        Else
            Doc.Tag = strCellKey
            Call InsertElementInRich(strCellKey)
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InsertElementInRich(ByVal strCellKey As String)
'���ܣ��ڻ�ϱ༭���в���Ҫ��,��ǰ����Ҫ�عؼ����м�Ϊ�޸�Ҫ��,��DOC���ɼ���δ���ڱ༭״̬ʱ������뱾����
Dim lngCp As Long, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, loldKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean, mfrmInsElement As New frmInsElement
Dim blnAutoSpace As Boolean
    On Error GoTo errHand
    lngCp = Doc.Selection.StartPos
    bBeteenKeys = IsBetweenAnyKeys(Doc, lngCp + 1, sKeyType, lKSS, lKSE, lKES, lKEE, loldKey, bNeeded)
    If bBeteenKeys And Doc.SelLength = 0 Then
    '��Ҫ��ǰ�����Ҫ�أ���Ҫ����λ��
        With Doc
            If .Range(lngCp - 1, lngCp).Font.Protected And .Range(lngCp - 1, lngCp).Font.Hidden = False And .Range(lngCp, lngCp + 1).Font.Hidden And .Range(lngCp, lngCp + 3).Text = "EE(" And .Range(lngCp + 16, lngCp + 17).Font.Hidden = False Then
            'B����1�������عؼ��֣�[Ҫ��]|�����عؼ��֣���ͨ�ı�
                Call .Range(lngCp + 16, lngCp + 16).Selected
                bBeteenKeys = False
            ElseIf .Range(lngCp - 1, lngCp).Font.Protected And .Range(lngCp - 1, lngCp).Font.Hidden = False And .Range(lngCp, lngCp + 1).Font.Hidden And .Range(lngCp, lngCp + 3).Text = "EE(" And .Range(lngCp + 16, lngCp + 19).Text = "ES(" Then
            'B����1�������عؼ��֣�[Ҫ��]|�����عؼ��֣������عؼ��֣�[Ҫ��]�����عؼ��֣�
                .Range(lngCp + 16, lngCp + 16).Text = " "
                .Range(lngCp + 16, lngCp + 17).Font.Protected = False
                .Range(lngCp + 16, lngCp + 17).Font.Hidden = False
                Call .Range(lngCp + 17, lngCp + 17).Selected
                lngCp = lngCp + 17
                blnAutoSpace = True
                bBeteenKeys = False
            ElseIf .Range(lngCp - 1, lngCp).Font.Hidden And .Range(lngCp, lngCp + 1).Font.Protected And .Range(lngCp, lngCp + 1).Font.Hidden = False And .Range(lngCp - 16, lngCp - 13).Text = "ES(" And .Range(lngCp - 17, lngCp - 16).Font.Hidden = False Then
            'B����2����ͨ�ı������عؼ��֣�|[Ҫ��]�����عؼ��֣�
                Call .Range(lngCp - 16, lngCp - 16).Selected
                bBeteenKeys = False
            ElseIf .Range(lngCp - 1, lngCp).Font.Hidden And .Range(lngCp, lngCp + 1).Font.Protected And .Range(lngCp, lngCp + 1).Font.Hidden = False And .Range(lngCp - 16, lngCp - 13).Text = "ES(" And .Range(lngCp - 32, lngCp - 29).Text = "EE(" Then
            'B����2�������عؼ��֣�[Ҫ��]�����عؼ��֣������عؼ��֣�|[Ҫ��]�����عؼ��֣�
                .Range(lngCp - 16, lngCp - 16).Text = " "
                lngCp = lngCp + 1
                .Range(lngCp - 17, lngCp - 16).Font.Protected = False
                .Range(lngCp - 17, lngCp - 16).Font.Hidden = False
                Call .Range(lngCp - 16, lngCp - 16).Selected
                lngCp = lngCp - 16
                blnAutoSpace = True
                bBeteenKeys = False
            ElseIf .Range(lngCp - 1, lngCp).Font.Hidden And .Range(lngCp, lngCp + 1).Font.Protected And .Range(lngCp, lngCp + 1).Font.Hidden = False And .Range(lngCp - 16, lngCp + 13).Text = "ES(" And lngCp - 16 = 0 Then
                Call .Range(0, 0).Selected
                bBeteenKeys = False
            End If
        End With
    End If
    With Document
        If bBeteenKeys Then '�޸�Ҫ��
            If sKeyType = "E" Then
                mfrmTipInfo.ShowTipInfo stbThis.hWnd, "��ʾ��" & vbCrLf & "      ��ǰ�������λ��ֻ���޸�Ҫ��" & vbCrLf & "      ����������Ҫ��֮�����Ҫ����������ո�", True, 0
                If mfrmInsElement.ShowMe(Me, Document.Elements("K" & loldKey), True, True, False, True) Then
                    If .Elements("K" & loldKey).�滻�� = 1 And (editType = TabET_�������༭ Or editType = TabET_���������) Then '�༭ʱ����Ҫ��
                        .Elements("K" & loldKey).�����ı� = GetReplaceEleValue(.Elements("K" & loldKey).Ҫ������, .EPRPatiRecInfo.����ID, .EPRPatiRecInfo.��ҳID, .EPRPatiRecInfo.������Դ, .EPRPatiRecInfo.ҽ��id, .EPRPatiRecInfo.Ӥ��)
                    End If
                    Call .Elements("K" & loldKey).Refresh(Doc)
                End If
            End If
        Else '����Ҫ��
            Dim NewElement As New cTabElement, lnewKey As Long
            If mfrmInsElement.ShowMe(Me, NewElement, True, True, False, True) Then
                lnewKey = .Elements.Add                                          '����Ҫ��
                Call NewElement.Clone(.Elements("K" & lnewKey))                  '����Ҫ������ȡ��
                If .Elements("K" & lnewKey).�滻�� = 1 And (editType = TabET_�������༭ Or editType = TabET_���������) Then '�༭ʱ����Ҫ��
                    .Elements("K" & lnewKey).�����ı� = GetReplaceEleValue(.Elements("K" & lnewKey).Ҫ������, .EPRPatiRecInfo.����ID, .EPRPatiRecInfo.��ҳID, .EPRPatiRecInfo.������Դ, .EPRPatiRecInfo.ҽ��id, .EPRPatiRecInfo.Ӥ��)
                End If
                .Elements("K" & lnewKey).���� = .Cells(strCellKey).Row & "|" & .Cells(strCellKey).Col
                Call .Elements("K" & lnewKey).InsertIntoEditor(Doc, editType)  'ˢ����ʾ
                If blnAutoSpace Then '��Ҫ��֮�����Ҫ�أ����Զ�׷�ӿո񣬲���Ҫ�غ�ɾ��
                    If FindKey(Doc, "E", lnewKey, lKSS, lKSE, lKES, lKEE, bNeeded) Then
                        If Doc.Range(lKSS - 1, lKSS).Text = " " Then
                            Doc.Range(lKSS - 1, lKSS).Text = ""
                        End If
                    End If
                End If
                Call GetFromDoc(strCellKey, False)
                If Doc.Enabled And Doc.Visible Then Doc.SetFocus
            End If
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub GetTextELement(ByVal strCellKey As String)
'���ܣ�����Text Element��дF1Main�еĵ�Ԫ����������ı�
Dim i As Long, lCount As Long, strTmp As String, ltCount As Long, leCount As Long, cleTmp As cTabElement, strAEl As String
    On Error GoTo errHand
    With Document.Cells(strCellKey)
        ltCount = UBound(Split(.TextKey, "|")): If ltCount < 0 Then ltCount = 0
        leCount = UBound(Split(.ElementKey, "|")): If leCount < 0 Then leCount = 0
        lCount = ltCount + leCount
        For i = 1 To lCount
            Set cleTmp = .clElement(Document.Elements, i)
            If cleTmp Is Nothing Then '�ô���Ϊ�ı�
                strTmp = strTmp & ToVarchar(.clText(Document.Texts, i).�����ı�, 4000)
            Else
                With Document.Elements("K" & cleTmp.Key)
                    If .�滻�� = 1 And (editType = TabET_�������༭ Or editType = TabET_���������) Then
                        If Trim(.�����ı�) = "" Then
                            strAEl = GetReplaceEleValue(.Ҫ������, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.������Դ, Document.EPRPatiRecInfo.ҽ��id, Me.Document.EPRPatiRecInfo.Ӥ��)
                            .�����ı� = strAEl
                            If strAEl = "" Then
                                If .�Զ�ת�ı� Then
                                    strTmp = strTmp & " " & .Ҫ�ص�λ
                                Else
                                    strTmp = strTmp & "[" & .Ҫ������ & "]" & .Ҫ�ص�λ
                                End If
                            Else
                                strTmp = strTmp & strAEl
                            End If
                        Else
                            strTmp = strTmp & .�����ı� & .Ҫ�ص�λ
                        End If
                    Else
                        If .������̬ = 0 Then
                            strTmp = strTmp & IIf(Trim(.�����ı�) = "", "[" & .Ҫ������ & "]", .�����ı�) & .Ҫ�ص�λ
                        Else
                            strTmp = strTmp & .�����ı� & .Ҫ�ص�λ
                        End If
                    End If
                End With
            End If
        Next
        .�����ı� = strTmp
        F1Main.TextRC(.Row, .Col) = strTmp
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub GetFromDoc(ByVal strCellKey As String, ByVal blnRefreshCell As Boolean)
'��DOC��ȡ��TEXT,ELEMENT,����.textkey;.elementkey,֮��ˢ��F1��Ԫ����
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bsInKeys As Boolean, sKeyType As String, bNeeded As Boolean
Dim lngEnd As Long, p As Long, strText As String, strTmp As String, lNo As Long
Dim txtKeys As String, elKeys As String, ltKey As Long, leKey As Long
    
    On Error GoTo errHand
    If Not Doc.Visible Then Exit Sub
    
    AddUndo Document.Cells(strCellKey)
    strText = Doc.Text:             lngEnd = Len(Doc.Text)
    p = 0:                          lNo = 1
    Do While p < lngEnd
        
        bsInKeys = FindNextAnyKey(Doc, p, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bsInKeys Then '�ҵ�Ҫ��
            '�ȴ���Ҫ�عؼ���֮ǰ��TXT
            If p <> lKSS Then   '����Ҫ��֮ǰ���ı�
                strTmp = Mid(strText, p + 1, lKSS - p)
Process:        If strTmp <> "" Then
                    ltKey = Document.Texts.Add
                    Document.Texts("K" & ltKey).�����ı� = ToVarchar(strTmp, 4000)
                    Document.Texts("K" & ltKey).�����д� = lNo
                    txtKeys = txtKeys & "|" & ltKey
                    lNo = lNo + 1
                End If
            End If
            If p > lngEnd Then Exit Do '�����������ı�����
            
            '�ٴ���Ҫ��
            If lKey = 0 Then
                leKey = Document.Elements.Add
            Else
                leKey = lKey
            End If
            Document.Elements("K" & leKey).�����д� = lNo
            elKeys = elKeys & "|" & leKey
            p = lKEE:          lNo = lNo + 1
        Else            '�ı�
            '��Ҳ�Ҳ�����һ��Ҫ�ر������û��Ҫ��
            strTmp = Mid(strText, p + 1)
            p = lngEnd + 1
            GoTo Process
        End If
    Loop
    
    Document.Cells(strCellKey).TextKey = txtKeys 'ȡ���µ��ı�Key��
    Document.Cells(strCellKey).ElementKey = elKeys 'ȡ���µ�Ҫ��Key��
    If blnRefreshCell Then GetTextELement strCellKey
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshToDoc(ByVal strCellKey As String)
'���ܣ����ݵ�ǰ��Ԫ��ˢ��Rich�༭������
Dim i As Long, lCount As Long, strTmp As String, ltCount As Long, leCount As Long, cleTmp As cTabElement
    On Error GoTo errHand
    With Document.Cells(strCellKey)
        ltCount = UBound(Split(.TextKey, "|")): If ltCount < 0 Then ltCount = 0
        leCount = UBound(Split(.ElementKey, "|")): If leCount < 0 Then leCount = 0
        lCount = ltCount + leCount
        For i = 1 To lCount
            Set cleTmp = .clElement(Document.Elements, i)
            If cleTmp Is Nothing Then '�ô���Ϊ�ı�
                .clText(Document.Texts, i).InsertIntoEditor Doc
            Else
                cleTmp.InsertIntoEditor Doc, editType
            End If
        Next
        Doc.ForceEdit = True
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub EditElement(ByVal strKey As String, ByVal KeyAscii As Integer)
'���ܣ��༭Ҫ��
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, vR As F1Rect, lHheight As Long, lHwidth As Long
    On Error GoTo errHand
    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '�̶��и߶�
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '�̶��п��
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
    elEdit.Tag = strKey
    With Document.Elements("K" & Document.Cells(strKey).ElementKey)
        If KeyAscii = vbKeySpace Then KeyAscii = 0
        If (Not (.Ҫ�ر�ʾ = 2 Or .Ҫ�ر�ʾ = 3)) And KeyAscii <> 0 Then Exit Sub  '����\��ѡ,��ֵ�����⣬��������ֻ����ո��˫������༭״̬
        If (.Ҫ�ر�ʾ = 2 Or .Ҫ�ر�ʾ = 3) And InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 0 Then Exit Sub '�ࡢ��ѡҪ������ո��˫������ֵ
        'If .Ҫ������ = 0 And .Ҫ�ر�ʾ = 0 And InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 0 Then Exit Sub '��ֵ��Ҫ������ո��˫������ֵ
        
        If .������̬ = 1 Then
            If vR.Top < -15 Or vR.Left < -15 Then Exit Sub '���ݳ�����������,��ֹ�༭
            elEdit.Top = F1Main.Top + vR.Top + Screen.TwipsPerPixelX * 2: elEdit.Left = F1Main.Left + vR.Left + Screen.TwipsPerPixelX * 2
            elEdit.Width = vR.Width - Screen.TwipsPerPixelX * 2: elEdit.Height = vR.Height - Screen.TwipsPerPixelY * 2
        Else
            elEdit.Top = F1Main.Top + vR.Bottom + Screen.TwipsPerPixelX * 2: elEdit.Left = F1Main.Left + IIf(vR.Left < 0, lHwidth, vR.Left) + Screen.TwipsPerPixelX * 2
        End If
    End With
    elEdit.SetElement Document.Elements("K" & Document.Cells(strKey).ElementKey), KeyAscii, editType
    
    If Document.Elements("K" & Document.Cells(strKey).ElementKey).������̬ = 0 Then
        If elEdit.Top + elEdit.Height > F1Main.Top + F1Main.Height Then
            elEdit.Top = vR.Top - elEdit.Height - Screen.TwipsPerPixelY * 2
        End If
        
        If elEdit.Left + elEdit.Width > F1Main.Left + F1Main.Width Then
            elEdit.Left = vR.Left - elEdit.Width - Screen.TwipsPerPixelX * 2
        End If
    End If

    elEdit.Visible = True: elEdit.ZOrder 0: elEdit.SetFocus
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub PopDoc(ByVal strCellKey As String, ByVal KeyAscii As Integer, Optional ByVal blnNew As Boolean = True)
'���ܣ���ʾRich�ؼ�,blnNew=true ��ʾ��ʼ���ؼ�����������,=false��ʾ�������ݣ�ֻ����ʾ
Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, vR As F1Rect, lHheight As Long, lHwidth As Long
Dim lpLeft As Long, lpTop As Long, lrHeight As Long, lrWidth As Long 'XY����,�߿�
    On Error GoTo errHand
    Doc.Title = "�༭" '��ʾ����༭״̬
    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '�̶��и߶�
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '�̶��п��
    
    Call F1Main.GetSelection(0, lsRow, lsCol, leRow, leCol)
    Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
    If vR.Right - lHwidth <= 0 Or vR.Bottom - lHheight <= 0 Then '���ڿ���ʾ����
        Doc.Title = "": Call F1Main_GotFocus: Exit Sub
    ElseIf vR.Left >= 0 And vR.Top >= 0 Then '�����ڱ���м�
        lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + vR.Top: lrWidth = vR.Width: lrHeight = vR.Height
    ElseIf vR.Left >= 0 And vR.Top < 0 Then '�����Ϸ���������(��������)
        lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + lHheight: lrWidth = vR.Width: lrHeight = vR.Height + vR.Top - lHheight
    ElseIf vR.Left < 0 And vR.Top >= 0 Then '�����󷽲�������(��������)
        lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + vR.Top: lrWidth = vR.Width + vR.Left - lHwidth: lrHeight = vR.Height
    ElseIf vR.Left < 0 And vR.Top < 0 Then '�����Ϸ��󷽶�����(��������)
        lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + lHheight: lrWidth = vR.Width + vR.Left - lHwidth: lrHeight = vR.Height + vR.Top - lHheight
    Else '����,δ֪
        Doc.Title = "": Call F1Main_GotFocus: Exit Sub
    End If
    '�ؼ���λ
    Doc.Move lpLeft + Screen.TwipsPerPixelX * 2, lpTop + Screen.TwipsPerPixelY * 2, lrWidth - Screen.TwipsPerPixelX * 2, lrHeight - Screen.TwipsPerPixelY * 2
    Doc.Tag = strCellKey
    If blnNew Then
        Doc.NewDoc
        RefreshToDoc strCellKey
    End If
    '�ؼ�����
    Doc.Visible = True: Doc.ZOrder 0: Doc.SetFocus: Doc.ForceEdit = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ClearPicture()
'���ܣ��½����ĵ�ǰ����ѵĵ�ͼƬ,�ڱ༭��ͼƬ���ı���Ҫ�ؿؼ�����
Dim l As Long, lCount As Long
    On Error Resume Next
    lCount = Document.Cells.Count
    For l = 1 To lCount
        With Document.Cells.Item(l)
            If .PictureKey <> "" Then
                ClearChildMember .Key
            End If
        End With
    Next
    If PicEdit.Visible Then PicEdit.Visible = False
    If elEdit.Visible Then elEdit.Visible = False
    If Doc.Visible Then Doc.Visible = False
End Sub
Private Sub CalcSumRange(ByVal nRow As Long, ByVal nCol As Long)
'����:����ָ����Ԫ������Щ��Ԫ��ϼƵ���
Dim SumRange As String, subRange As String, SumVal As Double, l As Long

    With Document.Cells.Cell(nRow, nCol) '�ϼƵ�Ԫ��
        SumRange = .�������� '�ϼƵ�Ԫ���Դ��Ԫ��
        If UBound(Split(SumRange, ";")) > 0 Then
            For l = 0 To UBound(Split(SumRange, ";"))
                subRange = Split(SumRange, ";")(l)
                SumVal = SumVal + Val(Document.Cells.Cell(Split(subRange, ",")(0), Split(subRange, ",")(1)).�����ı�)
            Next
            .�����ı� = Format(SumVal, "0.00")
            F1Main.TextRC(nRow, nCol) = Format(SumVal, "0.00")
        End If
    End With
End Sub

Private Function ValiCellDate(Optional DataVerify As Boolean = True) As Boolean
'���ܣ�1 �� ����ͨ���¼�  �����ݱ��浽���е����� ���б���,�϶��ı���иߡ��п�
'      2 ����ǰ�����ݽ��н��� , Ŀǰֻ����Ҫ���ڲ�������ʱ
Dim l As Long, lCount As Long, lngWidth As Long, lngHeight As Long, blnChangeRC As Boolean
    If timeTmp.Enabled Then timeTmp.Enabled = False: blnChangeRC = True
    On Error GoTo errHand
    lCount = Document.Cells.Count
    For l = 1 To lCount
        With Document.Cells(l)
            If (editType = TabET_�����ļ����� Or TabET_ȫ��ʾ���༭) Then
                On Error Resume Next
                .Width = F1Main.ColWidthTwips(.Col)
                .Height = F1Main.RowHeight(.Row)
                Err.Clear
            End If
            
            If DataVerify Then
                On Error GoTo errHand
                Select Case .��������
                    Case cprCTElement
                        If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then
                            If .ElementKey <> "" Then
                                If Document.Elements("K" & .ElementKey).Ҫ������ = "" Then
                                    MsgBox .Row & "��" & .Col & "�� " & "ΪҪ�ص�Ԫ��,��δָ������Ҫ�أ�", vbInformation, gstrSysName
                                    Call F1Main.SetSelection(.Row, .Col, .Row, .Col)
                                    Exit Function
                                End If
                            Else
                                MsgBox .Row & "��" & .Col & "�� " & "ΪҪ�ص�Ԫ��,��δָ������Ҫ�أ�", vbInformation, gstrSysName
                                Call F1Main.SetSelection(.Row, .Col, .Row, .Col)
                                Exit Function
                            End If
                        End If
                    Case Else
                End Select
            End If
        End With
    Next
    
    If DataVerify Then
        If editType = TabET_�������༭ Or editType = TabET_��������� Then
            If Not mfrmMainError.ShowNotice(Me) Then Exit Function
            Me.Refresh
        End If
    End If
    
    If blnChangeRC Then F1Main_SelChange
    
    ValiCellDate = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub ContentMove(strType As String)
'����:ȫѡ,����,����,ճ��,Ŀǰֻ֧���ı�,�ͻ�ϱ༭����
Dim strCellKey As String, strTmp As String
    On Error GoTo errHand
    If SelCell Is Nothing Then Exit Sub
    strCellKey = SelCell.Key
    If strCellKey = "" Then Exit Sub
    If Not (Document.Cells(strCellKey).�������� = cprCTFixtext Or Document.Cells(strCellKey).�������� = cprCTText Or Document.Cells(strCellKey).�������� = cprCTTextElement) Then Exit Sub
    If Doc.Visible Then
'        If UCase(strType) <> "PASTE" And Doc.Selection.StartPos = Doc.Selection.EndPos Then Exit Sub
        Dim sType As String, lsSS As Long, lsSE As Long, lsES As Long, lsEE As Long, leKey As Long, bsInKeys As Boolean, bNeeded As Boolean
        bsInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.StartPos + 1, sType, lsSS, lsSE, lsES, lsEE, leKey, bNeeded)
        Dim leSS As Long, leSE As Long, leES As Long, leEE As Long, beInKeys As Boolean
        beInKeys = IsBetweenAnyKeys(Doc, Doc.Selection.EndPos + 1, sType, leSS, leSE, leES, leEE, leKey, bNeeded)
    End If
    
    Select Case UCase(strType)
        Case "ALL"
            Select Case Document.Cells(strCellKey).��������
                Case cprCTText, cprCTFixtext
                    
                Case cprCTTextElement
                    If Doc.Visible Then
                        Call Doc.SelectAll
                    End If
            End Select
        Case "CUT"
            If editType = TabET_�������༭ Or editType = TabET_��������� Then
                If Not AllowEdit(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) Then Exit Sub '���������п�����
            End If
            Select Case Document.Cells(strCellKey).��������
                Case cprCTText, cprCTFixtext
                    If mblnEditing Then
                        SendKeys "^X"
                    Else
                        Call Clipboard.SetText(Document.Cells(strCellKey).�����ı�)
                        Document.Cells(strCellKey).�����ı� = ""
                        F1Main.TextRC(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) = ""
                    End If
                Case cprCTTextElement
                    If Doc.Visible Then
                        If Doc.Selection.StartPos = Doc.Selection.EndPos Then Exit Sub
                        If bsInKeys And beInKeys Then '��ʼλ����ֹλ�����ڹؼ���֮��
                            strTmp = Doc.Range(lsSS, leEE).Text
                            Doc.Range(lsSS, leEE).Text = ""
                        ElseIf bsInKeys Then          '��ʼλ�ڹؼ���֮��
                            strTmp = Doc.Range(lsSS, Doc.Selection.EndPos).Text
                            Doc.Range(lsSS, Doc.Selection.EndPos).Text = ""
                        ElseIf beInKeys Then          '��ֹλ�ڹؼ���֮��
                            strTmp = Doc.Range(Doc.Selection.StartPos, leEE).Text
                            Doc.Range(Doc.Selection.StartPos, leEE).Text = ""
                        Else
                            strTmp = Doc.Selection.Text
                            Doc.Range(Doc.Selection.StartPos, Doc.Selection.EndPos).Text = ""
                        End If
                        If strTmp = "" Then Exit Sub
                        strTmp = Doc.GetCleanTxt(strTmp)
                        Clipboard.SetText strTmp
                    Else
                        strTmp = Document.Cells(strCellKey).�����ı�
                        If strTmp = "" Then Exit Sub
                        F1Main.TextRC(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) = ""
                        Document.Cells(strCellKey).�����ı� = "": Document.Cells(strCellKey).TextKey = "": Document.Cells(strCellKey).ElementKey = ""
                        Clipboard.SetText strTmp
                    End If
            End Select
        Case "COPY"
            Select Case Document.Cells(strCellKey).��������
                Case cprCTText, cprCTFixtext
                    If mblnEditing Then
                        SendKeys "^C"
                    Else
                        Call Clipboard.SetText(Document.Cells(strCellKey).�����ı�)
                    End If
                Case cprCTTextElement
                    If Doc.Visible Then
                        If Doc.Selection.StartPos = Doc.Selection.EndPos Then Exit Sub
                        If bsInKeys And beInKeys Then '��ʼλ����ֹλ�����ڹؼ���֮��
                            strTmp = Doc.Range(lsSS, leEE).Text
                        ElseIf bsInKeys Then          '��ʼλ�ڹؼ���֮��
                            strTmp = Doc.Range(lsSS, Doc.Selection.EndPos).Text
                        ElseIf beInKeys Then          '��ֹλ�ڹؼ���֮��
                            strTmp = Doc.Range(Doc.Selection.StartPos, leEE).Text
                        Else
                            strTmp = Doc.Selection.Text
                        End If
                        If strTmp = "" Then Exit Sub
                        strTmp = Doc.GetCleanTxt(strTmp)
                        Clipboard.SetText strTmp
                    Else
                        Call Clipboard.SetText(Document.Cells(strCellKey).�����ı�)
                    End If
            End Select
        Case "PASTE"
            If editType = TabET_�������༭ Or editType = TabET_��������� Then
                If Not AllowEdit(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) Then Exit Sub '���������п�����
            End If
            Select Case Document.Cells(strCellKey).��������
                Case cprCTText, cprCTFixtext
                    If mblnEditing Then
                        SendKeys "^V"
                    Else
                        Document.Cells(strCellKey).�����ı� = Clipboard.GetText()
                        F1Main.TextRC(Document.Cells(strCellKey).Row, Document.Cells(strCellKey).Col) = Document.Cells(strCellKey).�����ı�
                    End If
                Case cprCTTextElement
                    If Doc.Visible Then
                        strTmp = Clipboard.GetText
                        If bsInKeys And beInKeys Then '��ʼλ����ֹλ�����ڹؼ���֮��
                            Doc.Range(leEE, leEE).Selected
                        ElseIf bsInKeys Then          '��ʼλ�ڹؼ���֮��
                            Doc.Range(Doc.Selection.EndPos, Doc.Selection.EndPos).Selected
                        ElseIf beInKeys Then          '��ֹλ�ڹؼ���֮��
                            Doc.Range(leEE, leEE).Selected
                        Else
                            Doc.Range(Doc.Selection.EndPos, Doc.Selection.EndPos).Selected
                        End If
                        If strTmp = "" Or strTmp = "GetText" Then Exit Sub
                        Doc.Range(Doc.Selection.EndPos, Doc.Selection.EndPos).Font.Hidden = False
                        Doc.Range(Doc.Selection.EndPos, Doc.Selection.EndPos).Font.Protected = False
                        Doc.Range(Doc.Selection.EndPos, Doc.Selection.EndPos).Text = strTmp
                        Doc.Range(0, Len(Doc.Text)).Font.Name = SelCell.FontName '��������ʽ��ֵ
                        Doc.Range(0, Len(Doc.Text)).Font.Size = SelCell.FontSize
                        Doc.Range(0, Len(Doc.Text)).Font.Bold = SelCell.FontBold
                        Doc.Range(0, Len(Doc.Text)).Font.Italic = SelCell.FontItalic
                        Doc.Range(0, Len(Doc.Text)).Font.Underline = SelCell.FontUnderline
                        Doc.Range(0, Len(Doc.Text)).Font.ForeColor = SelCell.FontColor
                        Doc.Range(0, Len(Doc.Text)).Font.Strikethrough = SelCell.FontStrikeout
                    End If
            End Select
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function AddSign(arrSQL As Variant, SignCellKey As String) As Boolean
Dim SignKey As String, SignTxt As String, oSign As cTabSign
    On Error GoTo errHand
1    SignKey = "": SignCellKey = ""
2    If InStr("6,7,8", SelCell.��������) > 0 Then 'ȷ���Ƿ���ǩ��λ
3        SignKey = SelCell.SignKey
4        SignCellKey = SelCell.Key
5    Else
6        Dim l As Integer, lCount As Long
7        lCount = Document.Cells.Count
8        For l = 1 To lCount
9            If IIf(Document.Cells(l).Merge, InStr(Document.Cells(l).MergeRange, ";") > 0, True) Then
10                If InStr("6,7,8", Document.Cells(l).��������) > 0 Then
11                    If SignKey = "" Then
12                        SignKey = Document.Cells(l).SignKey
13                        SignCellKey = Document.Cells(l).Key
14                    Else
15                        SignKey = "": Exit For '�����ڶ��ǩ����Ԫ��ʱ������ʾ
16                    End If
17                End If
18            End If
19        Next
20    End If
    
21    If SignKey = "" Then
22        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "��ʾ��" & vbCrLf & "      ����ѡ����Ҫǩ���ĵ�Ԫ��" & vbCrLf, True, 0
23        Exit Function
24    End If
25    If InStr("7,8", Document.Cells(SignCellKey).��������) > 0 And Document.Cells(SignCellKey).��ֹ�� <> 0 Then
26        mfrmTipInfo.ShowTipInfo stbThis.hWnd, "��ʾ��" & vbCrLf & "      �õ�Ԫ���ڱ���״̬������ǩ����" & vbCrLf & "      ���飡", True, 0
27        Exit Function
28    End If
    
29    Set oSign = frmSign.ShowMe(SignKey, Me)  'ǩ�������ǩ��Ԫ�ظ�ֵ
30    If Not oSign Is Nothing Then
31        Set Document.Signs("K" & SignKey) = oSign
32    Else
33        Exit Function
34    End If

35    If Document.Cells(SignCellKey).��ֹ�� = 0 Then
36        Document.Cells(SignCellKey).��ʼ�� = 1
37        Document.Cells(SignCellKey).��ֹ�� = IIf(editType = TabET_���������, Document.EPRPatiRecInfo.���汾 + 1, 1)
38    Else
39        Document.Cells(SignCellKey).ID = 0                          'ͬһǩ��λ���ǩ��
40        Document.Cells(SignCellKey).������� = Document.mMaxNo + 1
41        Document.Cells(SignCellKey).��ʼ�� = Document.EPRPatiRecInfo.���汾 + 1
42        Document.Cells(SignCellKey).��ֹ�� = Document.Cells(SignCellKey).��ʼ��
43        Document.mMaxNo = Document.mMaxNo + 1
44    End If
45    With Document.Signs("K" & SignKey)
46        SignTxt = .ǰ������ & .���� & IIf(.��ʾ��ǩ, "����ǩ��_____________", "")
47        SignTxt = SignTxt & IIf(Trim(.��ʾʱ��) = "", "", "��" & Format(.ǩ��ʱ��, .��ʾʱ��))
48    End With
49    Dim lSignRow As Long, lSignCol As Long
50    lSignRow = Document.Cells(SignCellKey).Row: lSignCol = Document.Cells(SignCellKey).Col
51    F1Main.TextRC(lSignRow, lSignCol) = SignTxt
52    AddSign = True
53    Exit Function
errHand:
    Call MsgBox("AddSign������:" & Erl(), vbInformation, gstrSysName)
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function RollBack() As Boolean
Dim mfrmUntread As New frmUntread
    On Error GoTo errHand
1    If mfrmUntread.ShowMe(Me, mstrPrivs) Then
2        On Error Resume Next
3        Call mfrmParent.RefreshList
4        Call mfrmParent.Event_Saved(Document.EPRPatiRecInfo.ID) '���Ƶ�����Ҫ����Ϊ�����Ƿ�ģ̬��ʽ���ã��������¼���ʽ
5        Err.Clear
        
6        On Error GoTo errHand
7        Document.EPRPatiRecInfo.GetPatiRecordInfo Document.EPRPatiRecInfo.ID, mblnMoved '���¶�ȡ����
8        Call Me.ShowMe(mfrmParent, Document, mstrModelPrivate, mblnMoved, mblnCanPrint)
9    End If
    Exit Function
errHand:
    Call MsgBox("RollBack������:" & Erl(), vbInformation, gstrSysName)
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub timeTmp_Timer()
    ValiCellDate False
End Sub
Private Sub txtSum_KeyPress(KeyAscii As Integer)
    If InStr("1234567890,;" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsHistory_DblClick()
    On Error GoTo errHand
    If Val(vsHistory.TextMatrix(vsHistory.Row, vsHistory.Cols - 1)) = 0 Then Exit Sub
    Me.Enabled = False
    zlCommFun.ShowFlash "���Եȣ����������ڴ��ļ�", Me
    Dim DocTmp As New cTableEPR
    DocTmp.InitOpenEPR mfrmParent, TabEm_�޸�, TabET_���������, Document.EPRPatiRecInfo.ID, False, 2, Document.EPRPatiRecInfo.������Դ, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.Ӥ��, UserInfo.����ID, Document.EPRPatiRecInfo.ҽ��id, mstrModelPrivate, mblnMoved, mblnCanPrint, gbytEsign
    DocTmp.EPRPatiRecInfo.���汾 = Val(vsHistory.TextMatrix(vsHistory.Row, vsHistory.Cols - 1))
    DocTmp.frmEditor.ShowMe mfrmParent, DocTmp, mstrModelPrivate, mblnMoved, mblnCanPrint
    Me.Enabled = True: zlCommFun.StopFlash
    Exit Sub
errHand:
    Me.Enabled = True: zlCommFun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Sub zlRefreshPacsPic()
    mfrmPacsPic.zlRefresh Document.EPRPatiRecInfo.ҽ��id, Document.EPRFileInfo.lngModule
End Sub
Private Sub ExeUndo()
'ִ�г�������
'1 �������ֵ 2 �Ա����ʾ����ˢ����ʾ��ͼƬˢ����ʾ 3 ɾ��Undo�������һ����Ա
Dim strShow As String, lRow As Long, lCol As Long
    On Error GoTo errHand
    If Undo.Count < 1 Then Exit Sub
    With Undo(Undo.Count)
        Select Case .CT
            Case cprCTFixtext, cprCTText
                Document.Cells(.Key).�����ı� = .CTxt
                F1Main.TextRC(.Row, .Col) = .CTxt
                
                If InStr(Document.Cells(.Key).��������, ",") > 0 And InStr(Document.Cells(.Key).��������, ";") = 0 Then '�ϼƵ�Ԫ���Դ��Ԫ��
                    lRow = Split(Document.Cells(.Key).��������, ",")(0): lCol = Split(Document.Cells(.Key).��������, ",")(1) '�ϼƵ�Ԫ�������
                    Call CalcSumRange(lRow, lCol)
                End If
            Case cprCTElement
                Document.Cells(.Key).ElementKey = .Ekey: lRow = .Row: lCol = .Col: strShow = .CTxt
                With Document.Cells(.Key)
                    .�����ı� = strShow: strShow = ""
                    If .�����ı� <> "" Then
                        strShow = .�����ı�
                    Else
                        If editType = TabET_�����ļ����� Or editType = TabET_ȫ��ʾ���༭ Then
                            If .ElementKey <> "" Then
                                If Document.Elements("K" & .ElementKey).������̬ = 1 Then
                                    strShow = Document.Elements("K" & .ElementKey).�����ı�
                                Else
                                    strShow = "[" & Document.Elements("K" & .ElementKey).Ҫ������ & "]" & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                                End If
                            Else
                                strShow = ""
                            End If
                        Else
                            If Document.Elements("K" & .ElementKey).�滻�� = 1 Then '�Զ��滻Ҫ��
                                If Document.Elements("K" & .ElementKey).�Զ�ת�ı� Then 'ûȡ��ֵ���Ƿ��Զ�ת�����ı�(��)
                                    strShow = ""
                                Else
                                    strShow = "[" & Document.Elements("K" & .ElementKey).Ҫ������ & "]" & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                                End If
                            Else
                                If Document.Elements("K" & .ElementKey).������̬ = 1 Then '������̬=չ��
                                    strShow = Document.Elements("K" & .ElementKey).�����ı� & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                                Else
                                    strShow = "[" & Document.Elements("K" & .ElementKey).Ҫ������ & "]" & Document.Elements("K" & .ElementKey).Ҫ�ص�λ
                                End If
                            End If
                        End If
                    End If
                    F1Main.TextRC(lRow, lCol) = strShow
                End With
            Case cprCTTextElement
                Document.Cells(.Key).ElementKey = .Ekey
                Document.Cells(.Key).TextKey = .Tkey
                GetTextELement .Key '��ʾ�������
            Case cprCTPicture, cprCTReportPic
                Document.Cells(.Key).PictureKey = .PKey
                If Len(.PKey) <> 0 Then
                    Set Document.Pictures("K" & .PKey).OrigPic = .OrigPic
                End If
                Document.Cells(.Key).PicMarkKey = .PmKey
                PaintPictureOnTable .Key
        End Select
    End With
    Undo.Remove (Undo.Count)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
'   ��;��  ��̬���¹���������ɫ��ͼ�ꡣ
'################################################################################################################
Private Sub SetColorIcon(ID As Long, Color As OLE_COLOR)
    Dim ctlPictureBox As VB.PictureBox
    Set ctlPictureBox = Controls.Add("VB.PictureBox", "ctlPictureBox1")
    Dim ListImage As ListImage
    Set ListImage = imgColor.ListImages("FORECOLOR")

    ctlPictureBox.AutoRedraw = True
    ctlPictureBox.AutoSize = True
    ctlPictureBox.BackColor = imgColor.MaskColor

    ctlPictureBox.Picture = ListImage.ExtractIcon

    If Color = vbWhite Then Color = RGB(254, 254, 254)
    ctlPictureBox.Line (1, ctlPictureBox.Height * 0.6)-(ctlPictureBox.Width, ctlPictureBox.Height), Color, BF
    ctlPictureBox.Refresh

    'Replace icon
    imgColor.ListImages.Remove imgColor.ListImages("FORECOLOR").Index
    imgColor.ListImages.Add 1, "FORECOLOR", ctlPictureBox.Image

    'OK Now replace Tag property
    imgColor.ListImages(1).Tag = ID

    cbsMain.AddImageList imgColor
    cbsMain.RecalcLayout

    Me.Controls.Remove ctlPictureBox
    Set ctlPictureBox = Nothing
End Sub

Private Function RelateFeedback(ByVal isRelated As Boolean) As Boolean
'���ܣ���Ⱦ�����濨���������Խ��������������ȡ������
'������isRelated  true-������false-ȡ������
    Dim strSQL As String
    Dim rsDisease As ADODB.Recordset
    Dim strIDs As String
    Dim arrayID() As String
    Dim i As Long
    Dim objDisease As Object
  
On Error GoTo errHand
    If Me.Document.EPRPatiRecInfo.�������� <> Tab������� Then Exit Function
  
    If isRelated Then   '����
        If Me.Document.EPRPatiRecInfo.������Դ = TabPF_���� Then
            strSQL = "select rowNum as NO,a.ID,c.���� as ����, a.�Ǽ�ʱ�� from  �������Լ�¼ A ,���˹Һż�¼ B ,���ű� C where A.�ļ�ID is NULL  and A.�Һŵ� = B.NO and A.����ID = B.����ID and A.��¼״̬ <> 3 and A.�Ǽǿ���ID = C.ID  and A.����ID = [1] and B.ID = [2]"
        ElseIf Me.Document.EPRPatiRecInfo.������Դ = TabPF_סԺ Then
            strSQL = "select rowNum as NO,a.ID ,c.���� as ����,a.�Ǽ�ʱ�� from  �������Լ�¼ A ,���ű� C  where A.�ļ�ID is NULL  and A.��¼״̬ <> 3  and A.�Ǽǿ���ID = C.ID and A.����ID = [1] and A.��ҳID = [2] "
        End If
        Set rsDisease = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ñ����Ӧ�����Խ��������", Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.��ҳID)
        If rsDisease.RecordCount = 1 Then
            strSQL = "Zl_�������Լ���¼_Update(2," & rsDisease!ID & "," & Me.Document.EPRPatiRecInfo.ID & ",NULL,NULL,NULL,NULL)"
            Call zlDatabase.ExecuteProcedure(strSQL, "����������������Խ��������")
        ElseIf rsDisease.RecordCount > 1 Then
            Set objDisease = CreateObject("zl9Disease.clsDisease")
            If objDisease Is Nothing Then Exit Function
            If objDisease.GetFeedbackList().ShowMe(Me, rsDisease, strIDs) Then
                If strIDs <> "" Then
                    arrayID = Split(strIDs, ",")
                    For i = LBound(arrayID) To UBound(arrayID)
                        If Val(arrayID(i)) <> 0 Then
                            strSQL = "Zl_�������Լ���¼_Update(2," & arrayID(i) & "," & Me.Document.EPRPatiRecInfo.ID & ",NULL,NULL,NULL,NULL)"
                            Call zlDatabase.ExecuteProcedure(strSQL, "����������������Խ��������")
                        End If
                    Next
                End If
            End If
        End If
    Else 'ȡ������
        strSQL = "Zl_�������Լ���¼_Update(3, NULL " & "," & Me.Document.EPRPatiRecInfo.ID & ",NULL,NULL,NULL,NULL)"
        Call zlDatabase.ExecuteProcedure(strSQL, "ȡ��������������Խ���������Ĺ���")
    End If
    
    RelateFeedback = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

