VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.Form frmDocMsg 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҩʦ������"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14400
   Icon            =   "frmDocMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   14400
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   12600
      Picture         =   "frmDocMsg.frx":6852
      ScaleHeight     =   330
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   405
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTmp 
      Height          =   915
      Left            =   14400
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   915
      _cx             =   1614
      _cy             =   1614
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      MouseIcon       =   "frmDocMsg.frx":6DDC
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   10000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDocMsg.frx":76B6
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1920
         ScaleHeight     =   240
         ScaleWidth      =   480
         TabIndex        =   23
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   4320
      ScaleHeight     =   3735
      ScaleWidth      =   9735
      TabIndex        =   17
      Top             =   2160
      Width           =   9735
      Begin zl9CISJob.ucCommandBar cbsChat 
         Height          =   420
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   3015
         _extentx        =   5318
         _extenty        =   741
      End
      Begin VB.PictureBox picChat 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   9735
         TabIndex        =   4
         Top             =   480
         Width           =   9735
         Begin VB.PictureBox pic����A 
            BorderStyle     =   0  'None
            Height          =   855
            Index           =   0
            Left            =   0
            ScaleHeight     =   855
            ScaleWidth      =   3375
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   3375
            Begin VB.TextBox txt����A 
               BackColor       =   &H0080FF80&
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
               Height          =   255
               Index           =   0
               Left            =   960
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   19
               Text            =   "frmDocMsg.frx":7751
               Top             =   400
               Width           =   1335
            End
            Begin VB.Label lbl�Ķ� 
               AutoSize        =   -1  'True
               Caption         =   "�Ѷ�"
               ForeColor       =   &H00808080&
               Height          =   180
               Index           =   0
               Left            =   2520
               TabIndex        =   21
               Top             =   600
               Width           =   360
            End
            Begin VB.Label lbl����A 
               AutoSize        =   -1  'True
               Caption         =   "����Ա  2019-07-31 22:12:03"
               ForeColor       =   &H00808080&
               Height          =   180
               Index           =   0
               Left            =   840
               TabIndex        =   20
               Top             =   30
               Width           =   2430
            End
            Begin VB.Image img����A 
               Appearance      =   0  'Flat
               Height          =   720
               Index           =   0
               Left            =   50
               Picture         =   "frmDocMsg.frx":775E
               Stretch         =   -1  'True
               Top             =   0
               Width           =   720
            End
            Begin VB.Shape shp����A 
               BackColor       =   &H0080FF80&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H8000000F&
               DrawMode        =   9  'Not Mask Pen
               FillColor       =   &H0080FF80&
               FillStyle       =   0  'Solid
               Height          =   495
               Index           =   0
               Left            =   840
               Shape           =   4  'Rounded Rectangle
               Top             =   285
               Width           =   1575
            End
         End
      End
      Begin VB.VScrollBar vsBar 
         Height          =   7575
         LargeChange     =   4
         Left            =   9480
         SmallChange     =   4
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   120
      ScaleHeight     =   8055
      ScaleWidth      =   4095
      TabIndex        =   12
      Top             =   360
      Width           =   4095
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2580
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   3360
         _Version        =   589884
         _ExtentX        =   5927
         _ExtentY        =   4551
         _StockProps     =   0
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   1305
      End
      Begin zl9CISJob.ucCommandBar cbsList 
         Height          =   420
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   3975
         _extentx        =   7011
         _extenty        =   741
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "����(&S)"
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&C)"
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox picAdvice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   4320
      ScaleHeight     =   1695
      ScaleWidth      =   9735
      TabIndex        =   14
      Top             =   360
      Width           =   9735
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   7275
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   12645
         _cx             =   22304
         _cy             =   12832
         Appearance      =   0
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
         MouseIcon       =   "frmDocMsg.frx":8628
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772554
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   16119285
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   10000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDocMsg.frx":8F02
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   1
         BackColorFrozen =   14737632
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   24
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin zl9CISJob.ucCommandBar cbsAdvice 
         Height          =   420
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   3975
         _extentx        =   7011
         _extenty        =   741
      End
   End
   Begin VB.PictureBox picIn 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   4320
      ScaleHeight     =   2175
      ScaleWidth      =   9735
      TabIndex        =   13
      Top             =   6240
      Width           =   9735
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   0
         Left            =   7200
         ScaleHeight     =   345
         ScaleWidth      =   1095
         TabIndex        =   9
         Top             =   1680
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�ر�(C)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   230
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   90
            Width           =   705
         End
      End
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   1
         Left            =   8520
         ScaleHeight     =   345
         ScaleWidth      =   1095
         TabIndex        =   10
         Top             =   1680
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����(S)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   90
            Width           =   705
         End
      End
      Begin VB.TextBox txtIn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   1575
         Left            =   0
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   0
         Width           =   9735
      End
   End
   Begin VB.Timer TmrIcon 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   13680
      Top             =   240
   End
   Begin VB.PictureBox PicNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   13200
      Picture         =   "frmDocMsg.frx":8F9D
      ScaleHeight     =   330
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   14040
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocMsg.frx":9867
            Key             =   "PatiIn"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocMsg.frx":9E01
            Key             =   "Meet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocMsg.frx":A39B
            Key             =   "Msg"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocMsg.frx":10BFD
            Key             =   "PatiOut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocMsg.frx":114D7
            Key             =   "msgno"
         EndProperty
      EndProperty
   End
   Begin zlSubclass.Subclass Subclass 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Image img����B 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   14280
      Picture         =   "frmDocMsg.frx":11A71
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   720
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.ImageManager imgList 
      Left            =   11520
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDocMsg.frx":1293B
   End
End
Attribute VB_Name = "frmDocMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'----------------------------------------------------------------------------------------------------
'-----ϵͳ�����������
'----------------------------------------------------------------------------------------------------
Private Const MAX_TOOLTIP As Integer = 64
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_MOUSEWHEEL = &H20A          '������
Private Const SW_RESTORE = 9
Private Const conCOLOR_BULELIGHT As Long = &HE4B440
Private Const conCOLOR_BULE As Long = &HD48A00

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type

Private Enum colList
    col_��Ϣ״̬ = 0
    col_������ = 1
    col_����ʱ�� = 2
    col_������Ϣ = 3
    col_������Ϣ = 4
    col_��Դ = 5
    col_���� = 6
    col_�Ա� = 7
    col_���� = 8
    COL_��ʶ�� = 9
    col_���� = 10
    
    '������
    col_����ID = 11
    col_����Id = 12
    col_����ID = 13
    COL_ҽ��IDs = 14
    col_ҽ����� = 15
    col_�Ự״̬ = 16
    col_δ��ID = 17
    col_�ỰID = 18
    col_δ��ʱ�� = 19
End Enum

Private Enum COL��ҩ�嵥
    '������
    COLB_ID = 1
    COLB_��� = 2
    COLB_��ҩ��Դ = 3
    COLB_������ĿID = 4
    COLB_�շ�ϸĿID = 5
    COLB_Ƶ�ʼ�� = 6
    COLB_�����λ = 7
    COLB_�÷�id = 8
    COLB_�巨id = 9
    COLB_��ֹʱ�� = 10
    '�ɼ���
    COLB_��Ч = 11
    COLB_��ʼʱ�� = 12
    colB_ҩƷ��� = 13
    colB_��ҩ���� = 14
    COLB_�÷� = 15
    COLB_�������� = 16
    COLB_�ܸ����� = 17
    COLB_���� = 18
    COLB_ִ��Ƶ�� = 19
    
    '������
    COLB_Ƶ�ʴ��� = 20
    COLB_�䷽���� = 21
End Enum


Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

Private mnfIconData As NOTIFYICONDATA
Private mblnIconShow As Boolean
Private mintPreTime As Integer
Public isUnload As Boolean
Private mfrmParent As Object
Private mintType As Integer
Private mdtBegin As Date, mdtEnd As Date
Private mstrδ���Ựids As String
Private mbln��Ϣ���� As Boolean
Private WithEvents mclsNotice As clsNotice
Attribute mclsNotice.VB_VarHelpID = -1


Public Sub SetNotifyIcon(ByVal intType As Integer, ByVal strMsg As String)
    'intType 0-��ʼ��  1-��Ϣ 2-��˸ 3-��ԭ
    'strMsg
    On Error Resume Next
    '����Ĵ�����Խ�ͼ����ӵ�ϵͳͼ��
    If intType = 0 And mnfIconData.hwnd <> 0 Then Call Shell_NotifyIcon(NIM_DELETE, mnfIconData)
    mnfIconData.hwnd = Me.hwnd
    mnfIconData.uID = picMsg.Picture '����ȷ��ʹ���ĸ�ͼ��
    mnfIconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    mnfIconData.uCallbackMessage = WM_MOUSEMOVE
    mnfIconData.hIcon = IIf(intType = 2, PicNo.Picture.Handle, picMsg.Picture.Handle)
    mnfIconData.szTip = strMsg & vbNullChar  '�����ǽ�����Ƶ�ͼ����ʱ������ʾ������
    mnfIconData.cbSize = Len(mnfIconData)
    Call Shell_NotifyIcon(IIf(intType = 0, NIM_ADD, NIM_MODIFY), mnfIconData)
End Sub

Private Sub StartMsg()
    TmrIcon.Enabled = Not TmrIcon.Enabled
    If TmrIcon.Enabled = False Then Call SetNotifyIcon(3, IIf(mintType = 1, "����������", "סԺ������") & vbCrLf & "��ǰ�û���" & UserInfo.����)
End Sub


Private Sub cbsList_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 4 'ˢ��
        Call LoadMsg
        Call ClearChat
        rptList.Tag = ""
    End Select
End Sub


Private Sub Form_Load()
    On Error GoTo errH
    isUnload = False
    '���ؿ�ݰ�ť
    cmdSend.Top = -1000
    cmdClose.Top = -1000
    
    '����ؼ���ɫ
    Me.BackColor = RGB(247, 247, 247)
    picChat.BackColor = RGB(247, 247, 247)
    picBack.BackColor = RGB(247, 247, 247)
    pic����A(0).BackColor = RGB(247, 247, 247)
    lbl����A(0).BackColor = RGB(247, 247, 247)
    lbl�Ķ�(0).BackColor = RGB(247, 247, 247)
    txt����A(0).BackColor = RGB(129, 246, 129)
    picIn.BackColor = RGB(247, 247, 247)
    
    '�����¼���ʼ��
    Subclass.hwnd = Me.hwnd
    Subclass.Messages(WM_MOUSEWHEEL) = True
    
    '�˵���ʼ��
    Call InitCommandBarList
    
    Call InitAdviceTable
    
    Call InitDockPannel '��ʼ�϶����ֳ�ʼ��
    
    Call InitReportColumn
    
    Call SetNotifyIcon(0, IIf(mintType = 1, "����������", "סԺ������") & vbCrLf & "��ǰ�û���" & UserInfo.����)
    
    mstrδ���Ựids = ""
    
    Call LoadMsg

    Call ClearChat

    '��Ϣ��ʼ��
    Set mclsNotice = zl9ComLib.GetClsNotice
    
    '���DCN�����Ƿ�����
    If Not mclsNotice Is Nothing Then
        If mclsNotice.CheckDcnEnable(3) = False Then
            Set mclsNotice = Nothing
        End If
    End If
    
    Me.Caption = IIf(mintType = 1, "����������", "סԺ������")
    Call RestoreWinState(Me, App.ProductName)
    rptList.AllowColumnSort = False
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errH
    
    If isUnload = False Then
        Cancel = 1
        If txtIn.Text <> "" Then
            If MsgBox("���������ڴ���������Ϣ��ȷ�ϹرջỰ����?", vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
                Exit Sub
            End If
        End If
        Me.Hide
        Exit Sub
    End If
    Call Shell_NotifyIcon(NIM_DELETE, mnfIconData)
    Subclass.Messages(WM_MOUSEWHEEL) = False
    'ж����Ϣ����
    If Not mclsNotice Is Nothing Then
        Set mclsNotice = Nothing
    End If
    
    Call SaveWinState(Me, App.ProductName)
    Set mfrmParent = Nothing
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngMsg As Long
    
On Error GoTo errH
    lngMsg = X / Screen.TwipsPerPixelX
    If lngMsg = WM_LBUTTONDBLCLK Then
        If IsWindowEnabled(mfrmParent.hwnd) Then
           Me.Hide
           ShowWindow Me.hwnd, SW_RESTORE
           Call picBack_Resize
            vsBar.Value = vsBar.Max '�Զ���λ�����
            If rptList.SelectedRows.Count > 0 Then
                If InStr(mstrδ���Ựids & ",", "," & rptList.SelectedRows(0).Record(col_�ỰID).Value & ",") > 0 Then
                    Call ReadMsg(Val(rptList.SelectedRows(0).Record(col_�ỰID).Value & ""), rptList.SelectedRows(0).Record(col_������).Value & "")
                End If
            End If
        End If
    End If
    Exit Sub
errH:
    Err.Clear
End Sub

Private Sub picAdvice_Resize()
    On Error Resume Next
    cbsAdvice.Width = picAdvice.Width
    vsAdvice.Top = cbsAdvice.Top + cbsAdvice.Height + 10
    vsAdvice.Width = picAdvice.Width
    vsAdvice.Height = picAdvice.Height - vsAdvice.Top
End Sub

Private Sub picBack_Resize()
    On Error Resume Next
    cbsChat.Width = picBack.Width
    picChat.Height = picBack.Height - cbsChat.Top
    picChat.Width = picBack.Width - 300
    
    vsBar.Top = cbsChat.Height: vsBar.Height = picBack.Height - cbsChat.Height
    vsBar.Left = picBack.Width - vsBar.Width
End Sub

Private Sub picChat_Resize()
    On Error Resume Next
    Call CtlResize(1)
End Sub

Private Sub picIn_Resize()
    On Error Resume Next
    txtIn.Width = picIn.Width - 80
    txtIn.Height = picIn.Height - 600
    
    picBtn(1).Top = txtIn.Height + (picIn.Height - txtIn.Height - picBtn(1).Height) / 2
    picBtn(0).Top = picBtn(1).Top
    
    picBtn(1).Left = picIn.Width - picBtn(1).Width - 100
    picBtn(0).Left = picBtn(1).Left - picBtn(0).Width - 200
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    cbsList.Width = picList.Width
    
    rptList.Left = 0
    rptList.Top = cbsList.Top + cbsList.Height
    rptList.Width = picList.Width
    rptList.Height = picList.Height - rptList.Top
End Sub




Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    '���Զ�λ����
    On Error Resume Next
    If rptList.SelectedRows.Count = 0 Then Exit Sub

    If Not mfrmParent Is Nothing And Val(rptList.SelectedRows(0).Record(col_����Id).Value) <> 0 Then
        Call mfrmParent.LocateMsgPati(Val(rptList.SelectedRows(0).Record(col_����Id).Value), Val(rptList.SelectedRows(0).Record(col_����ID).Value), Val(Split(rptList.SelectedRows(0).Record(COL_ҽ��IDs).Value, ",")(0)))
    End If
End Sub

Private Sub rptList_SelectionChanged()
    On Error GoTo errH
    If rptList.SelectedRows.Count = 0 Then Exit Sub          '���������
    If rptList.SelectedRows.Count > 0 Then
        If Val(rptList.SelectedRows(0).Record(col_�Ự״̬).Value & "") = 1 Then
            rptList.PaintManager.HighlightForeColor = &H808080
        Else
            rptList.PaintManager.HighlightForeColor = vbBlack
        End If
    End If
    
    If Val(rptList.Tag) = Val(rptList.SelectedRows(0).Record(col_�ỰID).Value) Then Exit Sub
    rptList.Tag = Val(rptList.SelectedRows(0).Record(col_�ỰID).Value)

    cbsAdvice.FindControl(999).Caption = "������Ϣ"
    cbsAdvice.RefreshCtl
    
    cbsChat.FindControl(0).Caption = "��ѡ������Ự"
    cbsChat.RefreshCtl
    
    LoadChat
    
    vsBar.Value = vsBar.Max '�Զ���λ�����
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearChat()
    On Error GoTo errH
    cbsAdvice.FindControl(999).Caption = "������Ϣ"
    cbsAdvice.RefreshCtl
    
    cbsChat.FindControl(0).Caption = "��ѡ������Ự"
    cbsChat.RefreshCtl
    
    '��մ�����Ϣ
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Cell(flexcpBackColor, vsAdvice.FixedRows, colB_ҩƷ���, vsAdvice.Rows - 1, colB_ҩƷ���) = &H8000000B      '����ɫ
    vsAdvice.Cell(flexcpBackColor, vsAdvice.FixedRows, COLB_�÷�, vsAdvice.Rows - 1, COLB_�÷�) = &H8000000B      '����ɫ
    vsAdvice.Cell(flexcpBackColor, vsAdvice.FixedRows, 0, vsAdvice.Rows - 1, 0) = &H8000000B
    
    
    '�������ؼ�
    Call SetCtl(0) 'ж�ؿؼ�
    picChat.Visible = False
    
    '�����������
    txtIn.Text = ""
    
     Call CtlResize(1)
     
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub TmrIcon_Timer()
    If Replace(mstrδ���Ựids, ",", "") = "" Then
        StartMsg
        Exit Sub
    End If
    Call SetNotifyIcon(IIf(mblnIconShow, 1, 2), IIf(mintType = 1, "����������", "סԺ������") & vbCrLf & "��ǰ�û���" & UserInfo.���� & vbCrLf & "��ǰ��" & UBound(Split(Mid(mstrδ���Ựids, 2), ",")) + 1 & "���µĻỰ��Ϣ")
    mblnIconShow = Not mblnIconShow
End Sub


'���ͺ͹رտؼ�
Private Sub cmdClose_Click()
    lblBtn_Click 0
End Sub

Private Sub cmdSend_Click()
    lblBtn_Click 1
End Sub

Private Sub picBtn_Click(Index As Integer)
    lblBtn_Click Index
End Sub

Private Sub picBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtn(Index).BackColor = conCOLOR_BULELIGHT
End Sub

Private Sub picBtn_Resize(Index As Integer)
    lblBtn(Index).Move picBtn(Index).ScaleWidth / 2 + lblBtn(Index).Width / 2, picBtn(Index).ScaleHeight / 2 - lblBtn(Index) / 2
End Sub

Private Sub lblBtn_Click(Index As Integer)
    If Index = 0 Then
        Unload Me
    Else
        Call SendMsg
    End If
End Sub

Private Sub picIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picBtn(0).BackColor <> conCOLOR_BULE And picBtn(0).Enabled Then picBtn(0).BackColor = conCOLOR_BULE
    If picBtn(1).BackColor <> conCOLOR_BULE And picBtn(1).Enabled Then picBtn(1).BackColor = conCOLOR_BULE
End Sub


Private Function InitCommandBarList() As Boolean
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim curDate As Date
    
    On Error GoTo errH
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsList.ObjCommandBar)
    With cbsList.ObjCommandBar

        Set .Icons = imgList.Icons

        .ActiveMenuBar.Title = "�˵�"
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
        .ActiveMenuBar.Visible = False
        .Options.LargeIcons = False

        '��Ϣ������
        '------------------------------------------------------------------------------------------------------------------
        Set objBar = .Add("ȱʡ", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagStretched
        Set objControl = NewToolBar(objBar, xtpControlLabel, 1, "��Ϣ�б�")
        objControl.IconId = 5
        
        Set objControl = objBar.Controls.Add(xtpControlLabel, 999, "ʱ��")   'ҽ��ʱ��
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = objBar.Controls.Add(xtpControlCustom, 3, "ʱ��")
            objCustom.Handle = cboTime.hwnd
            objCustom.Flags = xtpFlagRightAlign

        
        Set objControl = NewToolBar(objBar, xtpControlButton, 4, "", , , xtpButtonIcon)
        objControl.IconId = 2
        objControl.Flags = xtpFlagRightAlign


    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsAdvice.ObjCommandBar)
    With cbsAdvice.ObjCommandBar

        Set .Icons = imgList.Icons

        .ActiveMenuBar.Title = "�˵�"
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
        .ActiveMenuBar.Visible = False
        .Options.LargeIcons = False

        '������Ϣ������
        '------------------------------------------------------------------------------------------------------------------
        Set objBar = .Add("ȱʡ", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagStretched
        Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "������Ϣ")
        objControl.IconId = 4
        
        Set objControl = objBar.Controls.Add(xtpControlLabel, 999, "������Ϣ")   'ҽ��ʱ��
        objControl.Flags = xtpFlagRightAlign

    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsChat.ObjCommandBar)
    With cbsChat.ObjCommandBar

        Set .Icons = imgList.Icons

        .ActiveMenuBar.Title = "�˵�"
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
        .ActiveMenuBar.Visible = False
        .Options.LargeIcons = False

        '������Ϣ������
        '------------------------------------------------------------------------------------------------------------------
        Set objBar = .Add("ȱʡ", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagStretched
        Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "��ѡ������Ự")
        objControl.IconId = 1

    End With
    
    
    
    'ȱʡҽ��ʱ��
    cboTime.Clear
    cboTime.AddItem "����"
    cboTime.AddItem "����"
    cboTime.AddItem "����"
    cboTime.AddItem "�������"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "�������"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "[ָ��..]"
    Call zlControl.CboSetIndex(cboTime.hwnd, 6)
    curDate = zlDatabase.Currentdate
    mdtBegin = Format(curDate - 30, "yyyy-MM-dd 00:00:00")
    mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")

    mintPreTime = 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    On Error GoTo errH
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = imgList.Icons
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    On Error GoTo errH
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.ID = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    On Error GoTo errH
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = True 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    dkpMain.Options.LockSplitters = True
    DockPannelInit = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'InitDockPannel��ʼ���򻮷�
Private Sub InitDockPannel()
    Dim objPane As Pane
    On Error GoTo errH
    Set objPane = dkpMain.CreatePane(1, 270, 500, DockLeftOf, objPane)
    objPane.Title = "��Ϣ�б�"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 500, 150, DockRightOf, objPane)
    objPane.Title = "������Ϣ"
    objPane.Options = PaneNoCaption
'
    Set objPane = dkpMain.CreatePane(3, 500, 500, DockBottomOf, objPane)
    objPane.Title = "��Ϣ����"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(4, 500, 160, DockBottomOf, objPane)
    objPane.Title = "��Ϣ����"
    objPane.Options = PaneNoCaption

    Call DockPannelInit(dkpMain)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'�󶨲��ֿؼ�
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
  On Error GoTo errH
    Select Case Item.ID
        Case 1
            Item.Handle = picList.hwnd
        Case 2
            Item.Handle = picAdvice.hwnd
        Case 3
            Item.Handle = picBack.hwnd
        Case 4
            Item.Handle = picIn.hwnd
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub InitReportColumn()
    Dim objCol As ReportColumn
    
    On Error GoTo errH
    
    With rptList
        Set objCol = .Columns.Add(col_��Ϣ״̬, "", 18, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("msgno").Index - 1

        Set objCol = .Columns.Add(col_������, "����ҩʦ", 0, False)
        Set objCol = .Columns.Add(col_����ʱ��, "�Ựʱ��", 0, False)
        Set objCol = .Columns.Add(col_������Ϣ, "������Ϣ", 190, True)
        Set objCol = .Columns.Add(col_������Ϣ, "������Ϣ", 120, True)
        Set objCol = .Columns.Add(col_��Դ, "��Դ", 0, False)
        Set objCol = .Columns.Add(col_����, "��������", 0, False)
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 0, False)
        Set objCol = .Columns.Add(col_����, "����", 0, False)
        Set objCol = .Columns.Add(COL_��ʶ��, "��ʶ��", 0, False)
        Set objCol = .Columns.Add(col_����, "����", 0, False)
        Set objCol = .Columns.Add(col_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(col_����Id, "����ID", 0, False)
        Set objCol = .Columns.Add(col_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(COL_ҽ��IDs, "ҽ��IDs", 0, False)
        Set objCol = .Columns.Add(col_ҽ�����, "ҽ�����", 0, False)
        Set objCol = .Columns.Add(col_�Ự״̬, "�Ự״̬", 0, False)
        Set objCol = .Columns.Add(col_δ��ID, "δ��ID", 0, False)
        Set objCol = .Columns.Add(col_�ỰID, "�ỰID", 0, False)
        Set objCol = .Columns.Add(col_δ��ʱ��, "δ��ʱ��", 0, False)

        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnShaded
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridNoLines
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ������Ϣ..."
            .HighlightBackColor = &HFFEDCA
            .HighlightForeColor = vbBlack
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboTime_Click()
    Dim curDate As Date
    
    On Error GoTo errH
    If cboTime.ListIndex = mintPreTime And mintPreTime <> 7 Then Exit Sub
    
    curDate = zlDatabase.Currentdate
    
    Select Case cboTime.Text
    Case "����"
        mdtBegin = CDate(0)
        mdtEnd = CDate(0)
    Case "����"
        mdtBegin = Format(curDate, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "����"
        mdtBegin = Format(curDate - 1, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate - 1, "yyyy-MM-dd 23:59:59")
    Case "�������"
        mdtBegin = Format(curDate - 2, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "���һ��"
        mdtBegin = Format(curDate - 7, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "�������"
        mdtBegin = Format(curDate - 14, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "���һ��"
        mdtBegin = Format(curDate - 30, "yyyy-MM-dd 00:00:00")
        mdtEnd = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "[ָ��..]"
        If Not frmSelectTime.ShowMe(Me, mdtBegin, mdtEnd, cboTime) Then
            'ȡ��ʱ�ָ�ԭ����ѡ��
            Call zlControl.CboSetIndex(cboTime.hwnd, mintPreTime)
            rptList.SetFocus
            Exit Sub
        Else
            rptList.SetFocus
        End If
    End Select
        
    If mdtBegin = CDate(0) Or mdtEnd = CDate(0) Then
        cboTime.ToolTipText = ""
    Else
        cboTime.ToolTipText = "��Χ��" & Format(mdtBegin, "yyyy-MM-dd HH:mm:ss") & " �� " & Format(mdtEnd, "yyyy-MM-dd HH:mm:ss")
    End If
    mintPreTime = cboTime.ListIndex
    
    Call LoadMsg
    Call ClearChat
    rptList.Tag = ""
    Me.Refresh
    
    rptList.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadMsg()
    Dim strSQL As String, strFilter As String
    Dim rsTmp As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim j As Long, i As Long
    
    On Error GoTo errH
    If cboTime.Text <> "����" Then
        strFilter = " AND (m.�Ķ�ʱ�� IS NULL Or m.����ʱ�� Between [2] And [3]) "
    End If
    
    If mintType = 1 Then '����
        strSQL = "Select b.Id, b.�����ʶ, b.��������, b.����id, b.������, b.����ʱ��, b.״̬, '����' As ��Դ, a.����id, a.����, a.�Ա�, a.����, a.ִ�в���id As ����," & vbNewLine & _
            "              a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��, m.Id As ��Ϣid, m.����ʱ�� As ����ʱ��, m.�Ķ�ʱ��,decode(m.�Ķ�ʱ��,null,0,1) as �Ķ�״̬" & vbNewLine & _
            "       From ���˹Һż�¼ A, ����Ự�� B, ������Ϣ�� M" & vbNewLine & _
            "       Where a.����id = b.����id And a.Id = b.����id And b.Id = m.�Ựid(+) And b.������ = [1] And b.������Դ = 1 And m.������ = [1]" & strFilter
    Else
        strSQL = " Select b.Id, b.�����ʶ, b.��������, b.����id, b.������, b.����ʱ��, b.״̬, 'סԺ' As ��Դ, a.����id, a.����, a.�Ա�, a.����, a.��Ժ����id As ����," & vbNewLine & _
            "              a.��Ժ���� As ����ʱ��, a.סԺ�� As ��ʶ��, m.Id As ��Ϣid, m.����ʱ�� As ����ʱ��, m.�Ķ�ʱ��,decode(m.�Ķ�ʱ��,null,0,1) as �Ķ�״̬" & vbNewLine & _
            "       From ������ҳ A, ����Ự�� B, ������Ϣ�� M" & vbNewLine & _
            "       Where a.����id = b.����id And a.��ҳid = b.����id And b.Id = m.�Ựid(+) And b.������ =[1] And b.������Դ = 2  And m.������ = [1]" & strFilter
    End If

    strSQL = "Select d.Id, d.�����ʶ, d.��������, d.����id, d.������, d.����ʱ��, d.״̬, d.��Դ, d.����id, d.����, d.�Ա�, d.����, g.���� As ����, d.���� As ����id," & vbNewLine & _
            "       d.����ʱ��, d.��ʶ��, Max(d.��Ϣid) As δ��id, Max(d.����ʱ��) As ��Ϣʱ��," & vbNewLine & _
            "       Min(Nvl(d.�Ķ�ʱ��, To_Date('1900-01-01', 'yyyy-mm-dd'))) As �Ƿ�δ��,min(d.�Ķ�״̬)" & vbNewLine & _
            "From (" & strSQL & ") D, ���ű� G" & vbNewLine & _
            "Where d.���� =g.Id " & vbNewLine & _
            "Group By d.Id, d.�����ʶ, d.��������, d.����id, d.������, d.����ʱ��, d.״̬, d.��Դ, d.����id, d.����, d.�Ա�, d.����, g.����, d.����, d.����ʱ��, d.��ʶ��" & vbNewLine & _
            "Order By min(d.�Ķ�״̬),Max(d.��Ϣid) desc,d.״̬"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, mdtBegin, mdtEnd)
    
    rptList.Records.DeleteAll

    With rptList
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                Set objRecord = .Records.Add()
                Set objItem = objRecord.AddItem("")
                    If Format(rsTmp!�Ƿ�δ�� & "", "yyyy-MM-dd") = "1900-01-01" Then
                        objItem.Icon = img16.ListImages("Msg").Index - 1
                    End If
                Set objItem = objRecord.AddItem(rsTmp!������ & "")
                Set objItem = objRecord.AddItem(Format(rsTmp!��Ϣʱ�� & "", "yyyy-MM-dd hh:mm"))
                
                Set objItem = objRecord.AddItem(rsTmp!������ & "  " & Format(rsTmp!��Ϣʱ�� & "", "yyyy-MM-dd hh:mm"))
                objItem.Bold = True
                objItem.ForeColor = vbRed
                objItem.Icon = img16.ListImages.Item("Meet").Index - 1
                
                Set objItem = objRecord.AddItem(rsTmp!���� & "  " & rsTmp!�Ա� & "  " & rsTmp!���� & "  " & rsTmp!����)
                Set objItem = objRecord.AddItem(rsTmp!��Դ & "")
                Set objItem = objRecord.AddItem(rsTmp!���� & "")
                Set objItem = objRecord.AddItem(rsTmp!�Ա� & "")
                Set objItem = objRecord.AddItem(rsTmp!���� & "")
                Set objItem = objRecord.AddItem(rsTmp!��ʶ�� & "")
                Set objItem = objRecord.AddItem(rsTmp!���� & "")

                Set objItem = objRecord.AddItem(rsTmp!����ID & "")
                Set objItem = objRecord.AddItem(rsTmp!����ID & "")
                Set objItem = objRecord.AddItem(rsTmp!����id & "")
                Set objItem = objRecord.AddItem(rsTmp!�����ʶ & "")
                Set objItem = objRecord.AddItem(rsTmp!�������� & "")
                Set objItem = objRecord.AddItem(rsTmp!״̬ & "")
                Set objItem = objRecord.AddItem(rsTmp!δ��ID & "")
                Set objItem = objRecord.AddItem(rsTmp!ID & "")
                Set objItem = objRecord.AddItem(Format(rsTmp!�Ƿ�δ�� & "", "yyyy-MM-dd"))
  
                objRecord.PreviewText = rsTmp!�������� & ""

                '����ɵĻ��ﲡ���û�ɫ��ʾ
                If Val(rsTmp!״̬ & "") = 1 Then
                    For j = 0 To rptList.Columns.Count - 1
                        objRecord.Item(j).ForeColor = &H808080
                    Next
                End If

                rsTmp.MoveNext
            Loop
      
        End If
        .Populate
        
        '��ʼ��δ���ỰID
        For i = 0 To rptList.Rows.Count - 1
            With rptList.Rows(i)
                If Not .GroupRow Then
                    If .Record(col_δ��ʱ��).Value = "1900-01-01" Then
                        If InStr("," & mstrδ���Ựids & ",", "," & Val(.Record(col_�ỰID).Value & "") & ",") = 0 Then
                            mstrδ���Ựids = mstrδ���Ựids & "," & Val(.Record(col_�ỰID).Value & "")
                        End If
                    End If
                End If
            End With
        Next
        
        If mstrδ���Ựids <> "" And TmrIcon.Enabled = False Then Call StartMsg
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadChat()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng���� As Long
    Dim strMsg As String
    
    On Error GoTo errH
    '�߽�ֵ����
    Call SetCtl(0) 'ж�ؿؼ�
    If rptList.SelectedRows.Count < 1 Then Exit Sub
    If Val(rptList.SelectedRows(0).Record(col_�ỰID).Value & "") = 0 Then Exit Sub
    
    picChat.Visible = True
    '��ʾ������Ϣ
    cbsAdvice.FindControl(999).Caption = "��ǰ���ˣ�" & rptList.SelectedRows(0).Record(col_����).Value & " �Ա�" & rptList.SelectedRows(0).Record(col_�Ա�).Value & _
  " ���䣺" & rptList.SelectedRows(0).Record(col_����).Value & " ���ң�" & rptList.SelectedRows(0).Record(col_����).Value & IIf(rptList.SelectedRows(0).Record(col_��Դ).Value = "����", " �����", " סԺ��") & "��" & rptList.SelectedRows(0).Record(COL_��ʶ��).Value
    cbsAdvice.RefreshCtl
    
    '��ʾ��������Ϣ
    cbsChat.FindControl(0).Caption = rptList.SelectedRows(0).Record(col_������).Value
    cbsChat.RefreshCtl
    
    '��ʾ������Ϣ
    Call GetAdvice
    
    strSQL = "select a.id,A.�ỰID,A.������,A.��������,A.����ʱ��,A.������,A.�Ķ�ʱ�� from ������Ϣ�� A WHERE A.�ỰID=[1] order by A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rptList.SelectedRows(0).Record(col_�ỰID).Value & ""))
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            
            'ҩʦ����
                '������������
               lng���� = lng���� + 1
               Load pic����A(lng����)
               Set pic����A(lng����).Container = picChat
               pic����A(lng����).Tag = rsTmp!������ & ""
               
               'ͷ����ʾ
               Load img����A(lng����)
               Set img����A(lng����).Container = pic����A(lng����)
               img����A(lng����).Tag = rsTmp!ID & ""

               '��������
               Load shp����A(lng����)
               Set shp����A(lng����).Container = pic����A(lng����)

               Load txt����A(lng����)
               Set txt����A(lng����).Container = pic����A(lng����)
               
               
               '�����и�ʽ������������
               strMsg = Replace(rsTmp!�������� & "", vbCrLf, "[���д���]")
               strMsg = Replace(strMsg, Chr(10), vbCrLf)
               strMsg = Replace(strMsg, "[���д���]", vbCrLf)
               txt����A(lng����).Text = strMsg
               
               '����������Ϣ
               Load lbl����A(lng����)
               Set lbl����A(lng����).Container = pic����A(lng����)

               '�����Ķ���Ϣ
               Load lbl�Ķ�(lng����)
               Set lbl�Ķ�(lng����).Container = pic����A(lng����)
               lbl�Ķ�(lng����).Caption = "�Ѷ�"


               If rsTmp!������ & "" = UserInfo.���� Then
                    Set img����A(lng����).Picture = img����B.Picture
                    
                    shp����A(lng����).BackColor = RGB(221, 235, 255)
                    shp����A(lng����).FillColor = RGB(221, 235, 255)
                    txt����A(lng����).BackColor = RGB(208, 224, 240)
                    lbl����A(lng����).Caption = UserInfo.���� & "  " & Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm")
                    lbl�Ķ�(lng����).Visible = rsTmp!�Ķ�ʱ�� & "" <> ""
                Else
                    lbl����A(lng����).Caption = rsTmp!������ & "  " & Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm")
                    lbl�Ķ�(lng����).Visible = False
               End If
            rsTmp.MoveNext
        Loop
    End If
    Call CtlResize(1)
    Call SetCtl(1) '��ʾ�ؼ�
    If rptList.SelectedRows(0).Record(col_δ��ʱ��).Value = "1900-01-01" Then
        Call ReadMsg(Val(rptList.SelectedRows(0).Record(col_�ỰID).Value & ""), rptList.SelectedRows(0).Record(col_������).Value & "")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ReadMsg(ByVal Lng�ỰID As Long, str������ As String)
    '�����Ϣ�Ѷ�
    Dim strSQL As String

    On Error GoTo errH
    If Lng�ỰID = 0 Then Exit Sub
    strSQL = "Zl_������Ϣ��_Edit(2," & Lng�ỰID & ",null,'" & str������ & "')"
    
    If InStr("," & mstrδ���Ựids & ",", "," & Lng�ỰID & ",") > 0 Then
        mstrδ���Ựids = Replace(mstrδ���Ựids & ",", "," & Lng�ỰID & ",", "")
        If mstrδ���Ựids <> "" Then mstrδ���Ựids = "," & mstrδ���Ựids
    End If
    
    rptList.SelectedRows(0).Record(col_��Ϣ״̬).Icon = 9999
    rptList.SelectedRows(0).Record(col_δ��ʱ��).Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetCtl(intType As Integer)
    'ж�ؽ���Ŀؼ�
    'intType -0 ж�ؿؼ� -1��ʾ�ؼ�
    Dim obj As Object
    Dim obj���� As Object
    Dim i As Long
    
    On Error Resume Next
    For i = 1 To 6
        Select Case i
                Case 1
                    Set obj���� = shp����A
                Case 2
                    Set obj���� = img����A
                Case 3
                    Set obj���� = lbl����A
                Case 4
                    Set obj���� = txt����A
                Case 5
                    Set obj���� = lbl�Ķ�
                Case 6
                    Set obj���� = pic����A
        End Select
        For Each obj In obj����
            If obj.Index <> 0 Then
                Select Case intType
                    Case 0
                        Unload obj
                    Case 1
                        If i <> 5 Then
                            obj.Visible = True
                        End If
                End Select
                
            End If
        Next
    Next
End Sub


Private Sub CtlResize(lngMax As Long)
    Dim i As Long
    Dim lngTxtW As Long, lngTxtH As Long
    Dim lng������ݿ�� As Long
    
    On Error Resume Next
    
    If picChat.Visible = False Then
        picChat.Height = picBack.Height - cbsChat.Height
        vsBar.Visible = False
        vsBar.Max = (picChat.Height - picBack.Height + cbsChat.Height) / 100
        Exit Sub
    End If
    
    lng������ݿ�� = IIf(picChat.Width / 2 - 880 > 7000, 7000, picChat.Width / 2 - 880) '��ȡ���ı�ǩ���������ʾ�ı���
    
    For i = lngMax To pic����A.Count - 1
        If i = 1 Then
            pic����A(i).Top = 100: pic����A(i).Left = 0: pic����A(i).Width = picChat.Width
        Else
            pic����A(i).Top = pic����A(i - 1).Top + pic����A(i - 1).Height + 100: pic����A(i).Left = 0: pic����A(i).Width = picChat.Width
        End If
        
        lngTxtW = 0: lngTxtH = 0
        If pic����A(i).Tag = UserInfo.���� Then
            img����A(i).Top = 0: img����A(i).Left = picChat.Width - img����A(i).Width - 50
            lbl����A(i).Top = 30: lbl����A(i).Left = picChat.Width - lbl����A(i).Width - 840

            Call GetTextHight(txt����A(i).Text, lngTxtW, lngTxtH)
            If lngTxtH = 0 And lngTxtW <> 0 Then
                txt����A(i).Width = lngTxtW - 100
                txt����A(i).Height = 330
            ElseIf lngTxtH <> 0 And lngTxtW = 0 Then
                txt����A(i).Width = lng������ݿ�� - 100
                txt����A(i).Height = lngTxtH
            Else
                txt����A(i).Width = lngTxtW - 100
                txt����A(i).Height = lngTxtH
            End If
            
            txt����A(i).Top = 480: txt����A(i).Left = picChat.Width - txt����A(i).Width - 960
            If txt����A(i).Width > 4700 And txt����A(i).Height > 1800 Then
                txt����A(i).Top = txt����A(i).Top + 100
                txt����A(i).Left = txt����A(i).Left - 100
            End If
            
            shp����A(i).Width = txt����A(i).Width + 240 + IIf(txt����A(i).Width > 4700 And txt����A(i).Height > 1800, 220, 120)
            shp����A(i).Height = txt����A(i).Top + txt����A(i).Height + IIf(txt����A(i).Width > 4700 And txt����A(i).Height > 1800, 350, IIf(txt����A(i).Height < 500, 120, 180)) - shp����A(i).Top
            shp����A(i).Top = 285: shp����A(i).Left = picChat.Width - shp����A(i).Width - 840

            
            pic����A(i).Height = shp����A(i).Top + shp����A(i).Height + 75
            lbl�Ķ�(i).Top = shp����A(i).Top + shp����A(i).Height - lbl�Ķ�(i).Height - 15
            lbl�Ķ�(i).Left = shp����A(i).Left - lbl�Ķ�(i).Width - 100
        Else
            img����A(i).Top = 0: img����A(i).Left = 50
            lbl����A(i).Top = 30: lbl����A(i).Left = 840
            shp����A(i).Top = 285: shp����A(i).Left = 840
            txt����A(i).Top = 480: txt����A(i).Left = 960
            Call GetTextHight(txt����A(i).Text, lngTxtW, lngTxtH)
            If lngTxtH = 0 And lngTxtW <> 0 Then
                txt����A(i).Width = lngTxtW - 100
                txt����A(i).Height = 330
            ElseIf lngTxtH <> 0 And lngTxtW = 0 Then
                txt����A(i).Width = lng������ݿ�� - 100
                txt����A(i).Height = lngTxtH
            Else
                txt����A(i).Width = lngTxtW - 100
                txt����A(i).Height = lngTxtH
            End If
            
            
            If txt����A(i).Width > 4700 And txt����A(i).Height > 1800 Then
                txt����A(i).Top = txt����A(i).Top + 100
                txt����A(i).Left = txt����A(i).Left + 100
            End If
            
            shp����A(i).Width = txt����A(i).Left + txt����A(i).Width + IIf(txt����A(i).Width > 4700 And txt����A(i).Height > 1800, 220, 120) - shp����A(i).Left
            shp����A(i).Height = txt����A(i).Top + txt����A(i).Height + IIf(txt����A(i).Width > 4700 And txt����A(i).Height > 1800, 350, IIf(txt����A(i).Height < 500, 120, 180)) - shp����A(i).Top
            pic����A(i).Height = shp����A(i).Top + shp����A(i).Height + 75
            lbl�Ķ�(i).Top = shp����A(i).Top + shp����A(i).Height - lbl�Ķ�(i).Height - 15
            lbl�Ķ�(i).Left = shp����A(i).Left + shp����A(i).Width + 100
        End If
    Next
    picChat.Height = pic����A(pic����A.Count - 1).Top + pic����A(pic����A.Count - 1).Height + 100
    vsBar.Visible = picBack.Height - cbsChat.Height < picChat.Height
    vsBar.Max = (picChat.Height - picBack.Height + cbsChat.Height) / 100
End Sub


Private Sub txtIn_GotFocus()
    Call zlControl.TxtSelAll(txtIn)
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    
    If InStr("&'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SendMsg
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub vsBar_Change()
    Dim lngValue As Long
    On Error Resume Next
    
    lngValue = vsBar.Value
    picChat.Top = (-lngValue * 100 + cbsChat.Height)
End Sub


Private Sub Subclass_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    '�Զ������Ϣ������
    Dim tP As POINTAPI
    Dim sngX As Single, sngY As Single   '�������
    Dim intShift As Integer              '��갴��
    Dim bWay As Boolean                  '��귽��
    Dim bMouseFlag As Boolean            '����¼������־
    Dim wzDelta, wKeys As Integer
    On Error Resume Next
    
    If vsBar.Visible = False Then Exit Sub
    Select Case Msg
        Case WM_MOUSEWHEEL   '����
            wzDelta = (wParam And &HFFFF0000) \ &H10000 'ȡ��32λֵ�ĸ�16λ
            If wzDelta > 0 Then
                vsBar.Value = IIf(vsBar.Value - vsBar.LargeChange < 0, 0, vsBar.Value - vsBar.LargeChange)
            Else
                vsBar.Value = IIf(vsBar.Value + vsBar.LargeChange > vsBar.Max, vsBar.Max, vsBar.Value + vsBar.LargeChange)
            End If
    End Select
End Sub


Private Sub SendMsg()
    Dim strSQL As String
    Dim strSend As String
    Dim lng��ϢID As Long
    Dim strDate As String, strDateSQL As String
    
    On Error GoTo errH
    If txtIn.Text = "" Then Exit Sub
    If rptList.SelectedRows.Count < 1 Then Exit Sub
    If rptList.SelectedRows(0) Is Nothing Then Exit Sub
    If rptList.SelectedRows(0).GroupRow Then Exit Sub
    If Val(rptList.SelectedRows(0).Record(col_�ỰID).Value & "") = 0 Then Exit Sub
    If rptList.SelectedRows(0).Record(col_������).Value & "" = "" Then Exit Sub

    lng��ϢID = zlDatabase.GetNextId("������Ϣ��")
    strSend = Replace(txtIn.Text, "'", "")
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    strDateSQL = "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
    strSQL = "Zl_������Ϣ��_Edit(1," & Val(rptList.SelectedRows(0).Record(col_�ỰID).Value & "") & "," & lng��ϢID & ",'" & UserInfo.���� & "','" & strSend & "'," & strDateSQL & ",'" & rptList.SelectedRows(0).Record(col_������).Value & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

    txtIn.Text = ""
    Call AddMsg(strSend, strDate, UserInfo.����, lng��ϢID)
    vsBar.Value = vsBar.Max '�Զ���λ�����
    Call zlControl.ControlSetFocus(txtIn)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ReAddChat(lngMaxID As Long, Lng�ỰID As Long)
    '���ص�ǰ�Ự��������Ϣ
    Dim strMsg As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If Lng�ỰID = 0 Then Exit Sub
    strSQL = "select a.id,A.�ỰID,A.������,A.��������,A.����ʱ��,A.������,A.�Ķ�ʱ�� from ������Ϣ�� A WHERE A.�ỰID=[1] and a.id>[2] order by A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Lng�ỰID, lngMaxID)
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
               '�����и�ʽ������������
               strMsg = Replace(rsTmp!�������� & "", vbCrLf, "[���д���]")
               strMsg = Replace(strMsg, Chr(10), vbCrLf)
               strMsg = Replace(strMsg, "[���д���]", vbCrLf)
               Call AddMsg(strMsg, Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm"), rsTmp!������ & "", Val(rsTmp!ID & ""))
            rsTmp.MoveNext
        Loop
    End If
    vsBar.Value = vsBar.Max '�Զ���λ�����
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub AddMsg(ByVal strSend As String, ByVal strDate As String, ByVal str������ As String, ByVal lngID As Long)
    Dim lng���� As Long

    On Error GoTo errH
     '������������
    lng���� = pic����A.Count
    Load pic����A(lng����)
    Set pic����A(lng����).Container = picChat
    pic����A(lng����).Tag = str������
    pic����A(lng����).Visible = True
    
    'ͷ����ʾ
    Load img����A(lng����)
    Set img����A(lng����).Container = pic����A(lng����)
    img����A(lng����).Tag = lngID
    img����A(lng����).Visible = True
    

    '��������
    Load shp����A(lng����)
    Set shp����A(lng����).Container = pic����A(lng����)
    shp����A(lng����).Visible = True

    Load txt����A(lng����)
    Set txt����A(lng����).Container = pic����A(lng����)
    txt����A(lng����).Visible = True
    txt����A(lng����).Text = strSend
    
    
    '����������Ϣ
    Load lbl����A(lng����)
    Set lbl����A(lng����).Container = pic����A(lng����)
    lbl����A(lng����).Visible = True

    '�����Ķ���Ϣ
    Load lbl�Ķ�(lng����)
    Set lbl�Ķ�(lng����).Container = pic����A(lng����)
    lbl����A(lng����).Visible = True

    lbl����A(lng����).Caption = str������ & "  " & Format(strDate, "yyyy-MM-dd HH:mm")
    lbl�Ķ�(lng����).Caption = "�Ѷ�"
    lbl�Ķ�(lng����).Visible = False
    
    If str������ = UserInfo.���� Then
        Set img����A(lng����).Picture = img����B.Picture
        shp����A(lng����).BackColor = RGB(221, 235, 255)
        shp����A(lng����).FillColor = RGB(221, 235, 255)
        txt����A(lng����).BackColor = RGB(208, 224, 240)
    End If
    
    pic����A(lng����).Top = pic����A(lng���� - 1).Top + pic����A(lng���� - 1).Height + 100

    Call CtlResize(lng����)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetTextHight(ByVal strText As String, lngWidth As Long, lngHeight As Long)
    '���ڼ�������Ӧ�ı����ȸ߶�
    Dim lng������ݿ�� As Long
    Dim lngW As Long, lngH As Long
    
    On Error Resume Next

    lng������ݿ�� = IIf(picChat.Width / 2 - 880 > 7000, 7000, picChat.Width / 2 - 880) '��ȡ���ı�ǩ���������ʾ�ı���
    
    '�ȼ�����
    vsTmp.ColWidthMax = 0: vsTmp.ColWidthMin = 0
    vsTmp.RowHeightMin = 255: vsTmp.RowHeightMax = 255
    vsTmp.TextMatrix(0, 0) = strText
    vsTmp.Redraw = True
    vsTmp.AutoSizeMode = flexAutoSizeColWidth
    vsTmp.AutoSize 0
    vsTmp.Redraw = True
    vsTmp.Refresh
    lngW = vsTmp.ColWidth(0) - 20

    vsTmp.RowHeightMin = 0: vsTmp.RowHeightMax = 0
    vsTmp.ColWidthMax = lng������ݿ��
    vsTmp.ColWidthMin = lng������ݿ��
    vsTmp.TextMatrix(0, 0) = strText
    vsTmp.Redraw = True
    vsTmp.AutoSizeMode = flexAutoSizeRowHeight
    vsTmp.AutoSize 0
    vsTmp.Redraw = True
    vsTmp.Refresh
    lngH = vsTmp.RowHeight(0)

    If lngW > lng������ݿ�� And lngH = 255 Then
        lngWidth = lngW
    Else
        lngHeight = lngH
        lngWidth = IIf(lngW < lng������ݿ��, lngW, lng������ݿ��)
    End If
End Function


Private Sub InitAdviceTable()
'���ܣ���ʼ��ҽ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    On Error GoTo errH
    strHead = "ID;���;��ҩ��Դ;������ĿID;�շ�ϸĿID;Ƶ�ʼ��;�����λ;�÷�ID;�巨ID;��ֹʱ��;" & _
                "��Ч,450,4;��ʼʱ��,1000,1;ҩƷ���,850,4;��ҩ����,2000,1;�÷�,1000,1;����,850,4;����,850,4;����,450,4;ִ��Ƶ��,1000,4;Ƶ�ʴ���;�䷽����"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    'Ϊ��֧��zl9PrintMode
            End If
            .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '��¼ԭʼ�п�������ѡ����
        Next
        .Editable = flexEDNone
        .WordWrap = True
        .AutoSize colB_��ҩ����
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetAdvice()
    '��ȡ���˴�����Ϣ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, i As Long
    
    On Error GoTo errH
    With vsAdvice
         .Rows = .FixedRows
        If rptList.SelectedRows(0).Record(COL_ҽ��IDs).Value & "" = "" Then
            .Rows = .FixedRows + 1
        Else
            strSQL = "Select a.Id, a.���id As ���, a.������� As ҩƷ���, a.ҽ������ As ��ҩ����, a.ҽ������ As ҽ������, a.������Ŀid, a.�շ�ϸĿid, a.����, a.��ʼִ��ʱ�� As ��ʼʱ��," & vbNewLine & _
                    "       a.ִ����ֹʱ�� As ��ֹʱ��, Decode(a.������Դ, 1, a.�ܸ����� / e.�����װ, 2, a.�ܸ����� / e.סԺ��װ, a.�ܸ�����) As �ܸ�����, a.��������, a.ִ��Ƶ��, a.Ƶ�ʴ���," & vbNewLine & _
                    "       a.Ƶ�ʼ��, a.�����λ, b.������Ŀid As ��ҩid, c.���㵥λ, b.ҽ������ As �÷�, d.���� As ��ҩ�÷�, Decode(a.������Դ, 1, e.���ﵥλ, 2, e.סԺ��λ) As סԺ��λ,a.ҽ����Ч" & vbNewLine & _
                    "From ����ҽ����¼ A, ����ҽ����¼ B, ������ĿĿ¼ C, ������ĿĿ¼ D, ҩƷ��� E" & vbNewLine & _
                    "Where a.���id = b.Id And a.������Ŀid = c.Id And a.�շ�ϸĿid = e.ҩƷid(+) And b.������Ŀid = d.Id And a.����id = [1] And" & vbNewLine & _
                    "      (a.Id In (Select Column_Value As ҽ��id From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) Or" & vbNewLine & _
                    "      a.���id In (Select Column_Value As ҽ��id From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))))" & vbNewLine & _
                    "Order By a.����id, a.��ҳid, a.�Һŵ�, a.���, a.��ʼִ��ʱ��"
        
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rptList.SelectedRows(0).Record(col_����Id).Value & ""), rptList.SelectedRows(0).Record(COL_ҽ��IDs).Value & "")
            If Not rsTmp.EOF Then
                 .Redraw = flexRDNone
                 If .TextMatrix(.Rows - 1, colB_��ҩ����) = "" And Val(.TextMatrix(.Rows - 1, COLB_ID)) = 0 Then .Rows = .Rows - 1
                 For i = 1 To rsTmp.RecordCount
                    If (rsTmp!ҩƷ��� & "" = "7" Or rsTmp!ҩƷ��� & "" = "E") And Val(.Cell(flexcpData, .Rows - 1, COLB_���)) = Val(rsTmp!��� & "") Then
                        If rsTmp!ҩƷ��� & "" = "7" Then
                            .TextMatrix(.Rows - 1, COLB_�䷽����) = .TextMatrix(.Rows - 1, COLB_�䷽����) & vbCrLf & rsTmp!��ҩ���� & " " & FormatEx(NVL(rsTmp!��������), 5) & rsTmp!���㵥λ & " " & rsTmp!ҽ������
                        Else
                            .TextMatrix(.Rows - 1, COLB_�巨id) = Val(rsTmp!������ĿID & "")
                        End If
                    Else
                        .Rows = .Rows + 1
                        lngRow = .Rows - 1
                            
                        '������
                        .TextMatrix(lngRow, COLB_��Ч) = IIf(Val(rsTmp!ҽ����Ч & "") = 1, "����", "����")
                        .ColHidden(COLB_��Ч) = rptList.SelectedRows(0).Record(col_��Դ).Value = "����"
               
                        .TextMatrix(lngRow, COLB_������ĿID) = Val(rsTmp!������ĿID & "")
                        .TextMatrix(lngRow, COLB_�շ�ϸĿID) = Val(rsTmp!�շ�ϸĿid & "")
                        .TextMatrix(lngRow, COLB_Ƶ�ʼ��) = Val(rsTmp!Ƶ�ʼ�� & "")
                        .TextMatrix(lngRow, COLB_�����λ) = rsTmp!�����λ & ""
                        .TextMatrix(lngRow, COLB_�÷�id) = Val(rsTmp!��ҩid & "")
                        .TextMatrix(lngRow, COLB_��ֹʱ��) = Format(rsTmp!��ֹʱ�� & "", "yyyy-mm-dd hh:mm")
                        .TextMatrix(lngRow, COLB_��ʼʱ��) = Format(rsTmp!��ʼʱ�� & "", "yyyy-mm-dd hh:mm")
                        .TextMatrix(lngRow, colB_ҩƷ���) = Decode(rsTmp!ҩƷ��� & "", "5", "����ҩ", "6", "�г�ҩ", "�в�ҩ")
                        .TextMatrix(lngRow, colB_��ҩ����) = IIf(.TextMatrix(lngRow, colB_ҩƷ���) = "�в�ҩ", rsTmp!�÷� & "", rsTmp!��ҩ���� & "")
                        .TextMatrix(lngRow, COLB_�÷�) = IIf(.TextMatrix(lngRow, colB_ҩƷ���) = "�в�ҩ", rsTmp!��ҩ�÷� & "", rsTmp!�÷� & "")
                        .TextMatrix(lngRow, COLB_��������) = IIf(.TextMatrix(lngRow, colB_ҩƷ���) = "�в�ҩ", "", FormatEx(NVL(rsTmp!��������), 5)) & IIf(.TextMatrix(lngRow, colB_ҩƷ���) = "�в�ҩ", "", rsTmp!���㵥λ & "")
                        .TextMatrix(lngRow, COLB_�ܸ�����) = FormatEx(NVL(rsTmp!�ܸ�����), 5) & IIf(Val(rsTmp!�ܸ����� & "") = 0, "", IIf(.TextMatrix(lngRow, colB_ҩƷ���) = "�в�ҩ", "��", rsTmp!סԺ��λ & ""))
                        .Cell(flexcpData, lngRow, COLB_���) = Val(rsTmp!��� & "")
                        
                        If .TextMatrix(lngRow, colB_ҩƷ���) = "�в�ҩ" Then
                            .TextMatrix(lngRow, COLB_���) = ""
                        Else
                            If .Cell(flexcpData, lngRow, COLB_���) = .Cell(flexcpData, lngRow - 1, COLB_���) And .Cell(flexcpData, lngRow, COLB_���) <> "" Then
                                If .TextMatrix(lngRow - 1, COLB_���) = "" Then .TextMatrix(lngRow - 1, COLB_���) = -(lngRow - 1)
                                .TextMatrix(lngRow, COLB_���) = .TextMatrix(lngRow - 1, COLB_���)
                            End If
                        End If
    
                        If rsTmp!���� & "" = "" Then
                            If rsTmp!��ֹʱ�� & "" <> "" And rsTmp!��ʼʱ�� & "" <> "" Then
                                .TextMatrix(lngRow, COLB_����) = FormatEx(NVL(DateDiff("d", CDate(rsTmp!��ʼʱ�� & ""), CDate(rsTmp!��ֹʱ�� & ""))), 5)
                            End If
                        Else
                            .TextMatrix(lngRow, COLB_����) = FormatEx(NVL(rsTmp!����), 5)
                        End If
                        .TextMatrix(lngRow, COLB_ִ��Ƶ��) = rsTmp!ִ��Ƶ�� & ""
                        
                        If rsTmp!ҩƷ��� & "" = "7" Then
                            .TextMatrix(lngRow, COLB_�䷽����) = "�䷽��Ϣ��" & .TextMatrix(lngRow, COLB_�䷽����) & vbCrLf & rsTmp!��ҩ���� & " " & FormatEx(NVL(rsTmp!��������), 5) & rsTmp!���㵥λ & " " & rsTmp!ҽ������
                        ElseIf rsTmp!ҩƷ��� & "" = "E" Then
                            .TextMatrix(lngRow, COLB_�巨id) = Val(rsTmp!������ĿID & "")
                        End If
                        
                    End If
                    rsTmp.MoveNext
                 Next
                 .Redraw = flexRDDirect
            Else
                .Rows = .FixedRows + 1
                .TextMatrix(.Rows - 1, colB_��ҩ����) = "ҽ����ɾ��"
                .Cell(flexcpForeColor, .Rows - 1, colB_��ҩ����, .Rows - 1, colB_��ҩ����) = vbRed
                .Cell(flexcpFontBold, .Rows - 1, colB_��ҩ����, .Rows - 1, colB_��ҩ����) = True
            End If
    
        End If
        Call SetTagһ����ҩ
        .WordWrap = True
        '�Զ������и�
        .AutoSize colB_��ҩ����
        .Cell(flexcpBackColor, .FixedRows, colB_ҩƷ���, .Rows - 1, colB_ҩƷ���) = &H8000000B      '����ɫ
        .Cell(flexcpBackColor, .FixedRows, COLB_�÷�, .Rows - 1, COLB_�÷�) = &H8000000B      '����ɫ
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, 0) = &H8000000B
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub SetTagһ����ҩ(Optional ByVal lng��� As Long)
'���ܣ���һ����ҩ��ҽ��ǰ�ӱ�־
    Dim i As Long
    Dim lngUpRow As Long

    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If lng��� = 0 Then .TextMatrix(i, 0) = ""
            If lng��� <> 0 And Val(.TextMatrix(i, COLB_���)) = lng��� Then .TextMatrix(i, 0) = ""
            If Val(.TextMatrix(i, COLB_���)) <> 0 And ((lng��� = Val(.TextMatrix(i, COLB_���)) And lng��� <> 0) Or lng��� = 0) And .RowHidden(i) = False Then
                lngUpRow = GetUpRow(i)
                If lngUpRow = 0 Then
                    .TextMatrix(i, 0) = "��"
                Else
                    If Val(.TextMatrix(i, COLB_���)) = Val(.TextMatrix(lngUpRow, COLB_���)) And i <> lngUpRow Then
                        If .TextMatrix(lngUpRow, 0) = "��" Then
                            .TextMatrix(lngUpRow, 0) = "��"
                        End If
                        .TextMatrix(i, 0) = "��"
                    Else
                        .TextMatrix(i, 0) = "��"
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Function GetUpRow(ByVal lngRow As Long) As Long
'���ܣ�ȡ��һ����Ч��
    Dim i As Long

    With vsAdvice
        lngRow = lngRow - 1
        For i = lngRow To 1 Step -1
            If .RowHidden(i) = False Then
                GetUpRow = i: Exit For
            End If
        Next
    End With
End Function

Public Function ShowMe(frmParent As Object, intType As Integer)
'      intType 1=���� 2=סԺ
    Dim i As Long, blnδ�� As Boolean
    
    Set mfrmParent = frmParent
    mintType = intType
    Me.Show , frmParent
    
    '��ʼ���ж��Ƿ���δ����Ϣ��û��δ����Ϣ���ش���
    For i = 0 To rptList.Records.Count - 1
        If rptList.Records(i)(col_δ��ʱ��).Value & "" = "1900-01-01" Then
            blnδ�� = True
            Exit For
        End If
    Next
    If blnδ�� = False Then Me.Hide
End Function


Private Sub mclsNotice_DataArrival(ByVal lngNoticeCode As Long, ByVal intChangeType As Integer, ByVal strTableOwner As String, _
    ByVal TableName As String, ByVal strRowId As String)
    'DCN�����ݱ䶯֪ͨ
    'lngNoticeCode����Ϣ��ʶ(�̶�ֵ)�������������ĸ���Ϣ
    'intChangeType�����ݱ䶯���ͣ�1-���� 2-���� 3-ɾ��
    'strTableOwner��ע��DCN��������
    'TableName������
    'strRowid�����ص����ݱ䶯rowid
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnδ�� As Boolean
    
    On Error GoTo errH
    If TableName <> "������Ϣ��" Then Exit Sub
    
    If intChangeType = 1 Then '���µ���Ϣ������
        strSQL = "Select a.Id, a.�Ựid, a.������, a.��������, a.����ʱ��, a.������, a.�Ķ�ʱ��, b.�����ʶ, b.��������, b.����id, b.����id, b.������Դ, b.������, b.����ʱ��," & vbNewLine & _
                    "       b.������ As �Ự������, b.״̬" & vbNewLine & _
                    "From ������Ϣ�� A, ����Ự�� B" & vbNewLine & _
                    "Where a.�Ựid = b.Id And a.Rowid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRowId)
        If rptList.SelectedRows.Count > 0 Then
            If Val(rsTmp!�ỰID & "") = Val(rptList.SelectedRows(0).Record(col_�ỰID).Value) Then
                If Val(rsTmp!ID & "") > Val(img����A(img����A.Count - 1).Tag) Then
                    If Me.Visible = False Then
                        '�������˻ỰID
                        Call ReAddChat(Val(img����A(img����A.Count - 1).Tag), Val(rsTmp!�ỰID & ""))
                        If InStr("," & mstrδ���Ựids & ",", "," & Val(rsTmp!�ỰID & "") & ",") = 0 Then
                            mstrδ���Ựids = mstrδ���Ựids & "," & Val(rsTmp!�ỰID & "")
                        End If
                        If TmrIcon.Enabled = False Then StartMsg
                    Else
                        Call ReAddChat(Val(img����A(img����A.Count - 1).Tag), Val(rsTmp!�ỰID & ""))
                        Call ReadMsg(Val(rsTmp!�ỰID & ""), rptList.SelectedRows(0).Record(col_������).Value & "")
                    End If
                End If
            Else
                If Val(rsTmp!������Դ & "") = mintType And rsTmp!������ & "" = UserInfo.���� Then '��������סԺ�ͽ�����
                    '�������˻ỰID
                    If InStr("," & mstrδ���Ựids & ",", "," & Val(rsTmp!�ỰID & "") & ",") = 0 Then
                        mstrδ���Ựids = mstrδ���Ựids & "," & Val(rsTmp!�ỰID & "")
                    End If
                    '������˸
                    If TmrIcon.Enabled = False Then StartMsg
                    
                    'ˢ���б�
                    LoadMsg
    
                    '���¶�λ��ǰ�Ự
                    If Val(rptList.Tag) <> 0 Then
                        For i = 0 To rptList.Rows.Count - 1
                            With rptList.Rows(i)
                                If Not .GroupRow Then
                                    If .Record(col_�ỰID).Value = Val(rptList.Tag) Then
                                        Exit For
                                    End If
                                End If
                            End With
                        Next
                    
                        If i <= rptList.Rows.Count - 1 Then
                            Set rptList.FocusedRow = rptList.Rows(i)
                        End If
                    End If
                End If
            End If
        Else
            If Val(rsTmp!������Դ & "") = mintType And rsTmp!������ & "" = UserInfo.���� Then '��������סԺ�ͽ�����
                '�������˻ỰID
                If InStr("," & mstrδ���Ựids & ",", "," & Val(rsTmp!�ỰID & "") & ",") = 0 Then
                    mstrδ���Ựids = mstrδ���Ựids & "," & Val(rsTmp!�ỰID & "")
                End If
                '������˸
                If TmrIcon.Enabled = False Then StartMsg
                
                'ˢ���б�
                LoadMsg
            End If
        End If
    ElseIf intChangeType = 2 Then '���Ķ����µ���Ϣ������
        If rptList.SelectedRows.Count = 0 Then Exit Sub
        If pic����A.Count = 0 Then Exit Sub
        For i = lbl�Ķ�.Count - 1 To 1 Step -1
            If lbl�Ķ�(i).Visible = False And pic����A(i).Tag = UserInfo.���� Then
                blnδ�� = True
                Exit For
            End If
        Next
        If blnδ�� = False Then Exit Sub

        strSQL = "select a.id,A.�ỰID,A.������,A.��������,A.����ʱ��,A.������,A.�Ķ�ʱ�� from ������Ϣ�� A WHERE a.RowId=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRowId)
        If Not rsTmp Is Nothing Then
            If Not rsTmp.EOF Then
                If Val(rsTmp!�ỰID & "") = Val(rptList.SelectedRows(0).Record(col_�ỰID).Value) And rsTmp!������ & "" = UserInfo.���� Then
                    For i = 1 To lbl�Ķ�.Count - 1
                        If pic����A(i).Tag = UserInfo.���� Then
                            lbl�Ķ�(i).Visible = True
                        End If
                    Next
                End If
            End If
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



