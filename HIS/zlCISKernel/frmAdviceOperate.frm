VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmAdviceOperate 
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12120
   Icon            =   "frmAdviceOperate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   12120
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox pictmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9960
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pic���� 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   12120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4890
      Visible         =   0   'False
      Width           =   12120
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1275
         MaxLength       =   200
         TabIndex        =   1
         Top             =   15
         Width           =   9585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "У������"
         Height          =   180
         Left            =   495
         TabIndex        =   17
         Top             =   75
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   150
         Picture         =   "frmAdviceOperate.frx":058A
         Top             =   45
         Width           =   240
      End
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   714
      BandCount       =   1
      _CBWidth        =   12120
      _CBHeight       =   405
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   345
      Width1          =   3525
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   345
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   609
         ButtonWidth     =   1349
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫѡ"
               Key             =   "ȫѡ"
               Description     =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Ctrl+A)"
               Object.Tag             =   "ȫѡ"
               ImageKey        =   "ȫѡ"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫ��"
               Key             =   "ȫ��"
               Description     =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Ctrl+R)"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "ȫ��"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ִ��"
               Key             =   "ִ��"
               Description     =   "ִ��"
               Object.ToolTipText     =   "ִ��(Ctrl+E)"
               Object.Tag             =   "ִ��"
               ImageKey        =   "ִ��"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "������������(F12)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "ˢ��"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "���¶�ȡ����(F5)"
               Object.Tag             =   "ˢ��"
               ImageKey        =   "ˢ��"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����(F1)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�(ALT+X)"
               Object.Tag             =   "�˳�"
               ImageKey        =   "�˳�"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picPati 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   12120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   405
      Width           =   12120
      Begin VB.Frame fraOper 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   825
         TabIndex        =   29
         Top             =   -30
         Width           =   8895
         Begin VB.ComboBox cboTime 
            Height          =   300
            Index           =   0
            ItemData        =   "frmAdviceOperate.frx":0B14
            Left            =   5085
            List            =   "frmAdviceOperate.frx":0B16
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   150
            Width           =   1100
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Index           =   1
            ItemData        =   "frmAdviceOperate.frx":0B18
            Left            =   7830
            List            =   "frmAdviceOperate.frx":0B1A
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   150
            Width           =   1095
         End
         Begin VB.OptionButton optOper 
            Caption         =   "��ʼʱ��"
            Height          =   180
            Index           =   1
            Left            =   2330
            TabIndex        =   31
            Top             =   200
            Width           =   1050
         End
         Begin VB.OptionButton optOper 
            Caption         =   "��ǰʱ��"
            Height          =   180
            Index           =   0
            Left            =   1250
            TabIndex        =   30
            Top             =   200
            Width           =   1100
         End
         Begin VB.Label lblS 
            AutoSize        =   -1  'True
            Caption         =   "��ʼʱ�����ڿ���ʱ��"
            Height          =   180
            Left            =   3465
            TabIndex        =   36
            Top             =   195
            Width           =   1800
         End
         Begin VB.Label lblB 
            AutoSize        =   -1  'True
            Caption         =   "��ʼʱ�����ڿ���ʱ��"
            Height          =   180
            Left            =   6225
            TabIndex        =   35
            Top             =   210
            Width           =   1800
         End
         Begin VB.Label lblOper 
            Caption         =   "У��ʱ��(&T)"
            Height          =   180
            Left            =   120
            TabIndex        =   32
            Top             =   200
            Width           =   1000
         End
      End
      Begin VB.Frame fraBaby 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7320
         TabIndex        =   19
         Top             =   105
         Visible         =   0   'False
         Width           =   3195
         Begin VB.OptionButton optBaby 
            Caption         =   "Ӥ��ҽ��"
            Height          =   180
            Index           =   2
            Left            =   2175
            TabIndex        =   22
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "����ҽ��"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "����ҽ��"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   20
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdAlley 
         Caption         =   "����ʷ/����״̬"
         Height          =   350
         Left            =   10545
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Frame fraStop 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   0
         TabIndex        =   18
         Top             =   350
         Visible         =   0   'False
         Width           =   12105
         Begin VB.OptionButton optStop 
            Caption         =   "ָ��ʱ��"
            Height          =   180
            Index           =   1
            Left            =   2655
            TabIndex        =   28
            Top             =   90
            Width           =   1110
         End
         Begin VB.OptionButton optStop 
            Caption         =   "�ϴ�ִ��ʱ��"
            Height          =   180
            Index           =   0
            Left            =   1140
            TabIndex        =   27
            Top             =   90
            Value           =   -1  'True
            Width           =   1410
         End
         Begin VB.CheckBox chkRollSend 
            Caption         =   "�ջس��ڵ�"
            Height          =   195
            Left            =   7275
            TabIndex        =   26
            Top             =   90
            Width           =   1200
         End
         Begin VB.CheckBox chkNoSend 
            Caption         =   "����δ����"
            Height          =   195
            Left            =   6015
            TabIndex        =   25
            Top             =   90
            Width           =   1230
         End
         Begin VB.ComboBox cboҽ�� 
            Height          =   300
            Left            =   10550
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   45
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.ComboBox cboʱ�� 
            Height          =   300
            ItemData        =   "frmAdviceOperate.frx":0B1C
            Left            =   3930
            List            =   "frmAdviceOperate.frx":0B2F
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   45
            Width           =   1110
         End
         Begin MSMask.MaskEdBox txtʱ�� 
            Height          =   300
            Left            =   5040
            TabIndex        =   5
            Top             =   45
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblStop 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ֹʱ��(&T)"
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   105
            Width           =   990
         End
         Begin VB.Label lblҽ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ֹͣҽ��(&D)"
            Height          =   180
            Left            =   9540
            TabIndex        =   6
            Top             =   105
            Visible         =   0   'False
            Width           =   990
         End
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����: סԺ��: ����: ����:"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   150
         TabIndex        =   11
         Top             =   120
         Width           =   2250
      End
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6735
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6930
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   1560
      TabIndex        =   15
      Top             =   6885
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Align           =   1  'Align Top
      Height          =   1590
      Left            =   0
      TabIndex        =   2
      Top             =   5265
      Visible         =   0   'False
      Width           =   12120
      _cx             =   21378
      _cy             =   2805
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
      ForeColorFixed  =   -2147483630
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.PictureBox picUD 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   12120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5220
      Visible         =   0   'False
      Width           =   12120
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Align           =   1  'Align Top
      Height          =   3660
      Left            =   0
      TabIndex        =   0
      Top             =   1230
      Width           =   12120
      _cx             =   21378
      _cy             =   6456
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
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16771802
      ForeColorSel    =   -2147483640
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceOperate.frx":0B51
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   8205
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceOperate.frx":0BEC
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16853
            MinWidth        =   25
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceOperate.frx":1480
            Text            =   "ͨ��"
            TextSave        =   "ͨ��"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceOperate.frx":1A6A
            Text            =   "����"
            TextSave        =   "����"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmAdviceOperate.frx":2054
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmAdviceOperate.frx":268E
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAdviceOperate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���ܣ�
'0-ҽ������:
'    ֻ��ѡ��Ҫ���ϵ�ҽ��
'1-ֹͣҽ��:
'    ��Ҫָ����ֹʱ��(ȱʡΪ��ǰ,������ЧȱʡΪ�������,Ԥ���Ĳ���)
'    ��ʿͣʱ��Ҫָ��ֹͣҽ��
'2-ȷ��ֹͣ:
'    ֻ��ѡ����Ҫȷ�ϵ�ҽ��
'3-У��ҽ��:
'    ��¼��ҽ�������޸�У��ʱ��(�ǲ�¼��ȱʡΪ��ǰ���ɸ�,��¼��ȱʡΪ����ʱ��+1m)
'4-�����Ƽ���Ŀ:
'    ��ɾ��ÿ��ҽ���ļƼ���Ŀ
'5-��ͣҽ��
'    ѡ����Ҫ��ͣ��ҽ��
'6-����ҽ��
'    ѡ����Ҫ���õ�ҽ��
Private mfrmParent As Form
Private mMainPrivs As String
Private mlngҽ��ID As Long '����ȱʡ��λ
Private mint���� As Integer '0-ҽ������,1-ֹͣҽ��,2-ȷ��ֹͣ,3-ҽ��У��,4-�����Ƽ���Ŀ,5-��ͣҽ��,6-����ҽ��,7-ͣ�����
Private mbytUseType As Byte '0-ҽ�����ܵ���,1-�ٴ�·����Ŀִ�к����
Private mstrAdviceOfItem As String '���ظ��ٴ�·����·����Ŀ��Ӧ��ҽ��ID�Ĵ�,�ö��Ÿ�
Private mdateStop As Date '�ٴ�·������ʱ����ֹͣʱ��(����ʱ���1��)
                          'ת�ơ���Ժҽ���´�ʱ�����ҽ���Ŀ�ʼִ��ʱ��
Private mblnAutoRead As Boolean   '����ǰ�Զ�У�ԣ���ʱֻ��ȡ����ҽ��������У�ԣ���������������ȼ�,����/Σҽ��,����ҽ��������,��¼�����,ת�ƣ���Ժ��תԺ������
                                  '����ʱ����ȷ��ֹͣ���Զ���ȡ���벡�˵�ҽ��

Private mintҽ������Χ As Integer    'ҽ������Χ   0-����ҽ��,1-����ҽ��,2-Ӥ��ҽ��
Private mstrȱʡУ��ʱ�� As String  '��1λ��0-��ǰϵͳʱ��,1-��ʼʱ�䣻��2λ����ʼʱ����ڿ���ʱ��ʱ��ѡ��0-��ʼʱ�䣬1-����ʱ�䣻��3λ����ʼʱ��С�ڿ���ʱ��ʱ��ѡ��0-��ʼʱ�䣬1-����ʱ�䡣2��3λ���ڵ�1λΪ1ʱ��Ч��
Private mstrȱʡֹͣʱ�� As String '��1λ��0-��ǰϵͳʱ��,1-���һ�η��͵���ִֹ��ʱ�䡣��2λ������δ���͵�Ҫ��������3λ�����ڷ��͵�Ҫ�ջء�
Private mblnOnePati As Boolean     '�����˻��Ƕಡ��ģʽ
Private mbln���͵��� As Boolean

Private mlng����ID As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlngӤ�� As Long    'ת��ҽ���´ﵯ��ҽ��ֹͣ����ʱ�Ŵ���
Private mstrҩƷ�۸�ȼ� As String '���˵�ҩƷ�۸�ȼ�
Private mstr���ļ۸�ȼ� As String '���˵����ļ۸�ȼ�
Private mstr��ͨ��Ŀ�۸�ȼ� As String '���˵���ͨ��Ŀ�۸�ȼ�

Private mbln��ʿվ As Boolean
Private mlng�������� As Long

Private mint���� As Integer
Private mblnRefresh As Boolean
Private mblnOK As Boolean
Private mblnReturn As Boolean

Private mrsPrice As ADODB.Recordset
Private mrsDept As ADODB.Recordset
Private mstrLike As String
Private mint���� As Integer
Private mstrRollNotify As String '������Ҫ���г����ջ����ѵĲ���(����ID:��ҳID,...)

Private mblnҽ������ As Boolean
Private mbln��ʿǩ�� As Boolean
Private mbln���� As Boolean

Private mlng��ҩ�� As Long
Private mlng��ҩ�� As Long
Private mlng��ҩ�� As Long
Private mlng���ϲ��� As Long
Private mblnHaveAudit As Boolean   '�Ƿ����ִҵҽʦ�ʸ�
Private mlngͣ����� As Long       'ʵϰҽ��ֹͣҽ����Ҫ��� ����
Private mlngҽ������ID As Long
Private mlngӤ������ID As Long
Private mlngӤ������ID As Long
'PASS
Private mobjPassMap As Object  'PASS ������Ϣ��ڲ��� �˱������踳ֵ��PASS�еĲ���ID����ҳID������ҽ�������д���
Private mblnPass As Boolean  'PASSȨ��

Private mclsMipModule As zl9ComLib.clsMipModule
Private mstrPatiClsMsg As String  '�����Ϣ�Ĳ��� ��ʽ "����id1,��ҳid1;����id2,��ҳid2;......"
Private mstrPatiKeepMsg As String  '������Ϣ�Ĳ��� ��ʽ "����id1,��ҳid1;����id2,��ҳid2;......"
Private mblnAll As Boolean '�Ƿ��Ǽ�������ҽ�� mstrPatiAll
'��������
Private mblnFirst As Boolean
Private mstr����IDs As String   '����id,��ҳid;����id,��ҳid;......
Private mint��Ч As Integer
Private mint��� As Integer
Private mblnPauseLast As Boolean
Private mblnFirstLoad As Boolean
Private mbytSize As Byte '�����С 0-С���壨9������) 1-�����壨12 �����壩
Private mstrͣ��ԭ�� As String
Private mbln��������ִ�� As Boolean



Private Enum CtlID
    e���� = 0
    e���� = 1
    eӤ�� = 2
    
    e�ϴ�ִ��ʱ�� = 0
    eָ��ʱ�� = 1
    
    e��ǰʱ�� = 0
    e��ʼʱ�� = 1
    
    e���� = 0
    e���� = 1
End Enum

Private Const con_Date = "��ǰʱ��=__:__,�����賿=00:00,�����糿=08:00,��������=12:00,��������=18:00,��������=23:59"

'������
Private Const COL_ID = 0
Private Const COL_���ID = 1
Private Const COL_��ID = 2
Private Const COL_��� = 3
Private Const COL_������� = 4
Private Const COL_������� = 5
Private Const COL_���� = 6 '1-��ҩ�䷽,2-�������
'Pass��ʾ��
Private Const COL_��ʾ = 7
'������
Private Const COL_ѡ�� = 8 '
Private Const COL_���� = 9 '
Private Const COL_��ֹԭ�� = 10
'�ɼ���
Private Const COL_��־ = 11
Private Const COL_���� = 12
Private Const COL_סԺ�� = 13
Private Const COL_���� = 14
Private Const COL_Ӥ�� = 15
Private Const COL_��Ч = 16
Private Const COL_����ʱ�� = 17
Private Const COL_��ʼʱ�� = 18
Private Const col_ҽ������ = 19
Private Const COL_Ƥ�� = 20
Private Const COL_���� = 21
Private Const COL_���� = 22
Private Const COL_Ƶ�� = 23
Private Const COL_�÷� = 24
Private Const COL_ҽ������ = 25
Private Const COL_ִ��ʱ�� = 26
Private Const COL_��ֹʱ�� = 27 '
Private Const COL_ִ�п��� = 28
Private Const COL_ִ������ = 29
Private Const COL_�ϴ�ִ�� = 30
Private Const COL_����ҽ�� = 31
Private Const COL_У�Ի�ʿ = 32 '
Private Const COL_У��ʱ�� = 33 '
Private Const COL_ͣ��ҽ�� = 34 '
Private Const COL_ͣ��ʱ�� = 35 '
'����
Private Const COL_����ID = 36
Private Const COL_��ҳID = 37
Private Const COL_������ĿID = 38
Private Const COL_Ƶ�ʴ��� = 39
Private Const COL_Ƶ�ʼ�� = 40
Private Const COL_�����λ = 41
Private Const COL_ִ�б�� = 42
Private Const COL_�������� = 43
Private Const COL_�Թܱ��� = 44
Private Const COL_ִ�п���ID = 45
Private Const COL_���˿���ID = 46
Private Const COL_�շ�ϸĿID = 47
Private Const COL_������λ = 48
Private Const COL_ǰ��ID = 49
Private Const COL_ǩ��ID = 50
Private Const COL_������Ա = 51
Private Const COL_��������ID = 52
Private Const COL_����˵�� = 53
Private Const COL_ִ�з��� = 54
Private Const COL_�걾��λ = 55  '������ҩ���������ҩ��
Private Const COL_������� = 56
Private Const COL_���״̬ = 57
Private Const COL_��Ժ����ID = 58 '������ҳ.��Ժ����ID
Private Const COL_�������� = 59


'�Ƽ��嵥����ֵ
Private Const COLP_ҽ��ID = 0 '���Ӵ�ű����Ϣ
Private Const COLP_���ID = 1 '���Ӵ�ű����Ϣ
Private Const COLP_������� = 2 '���Ӵ�ű����Ϣ
Private Const COLP_������ĿID = 3
Private Const COLP_�շ�ϸĿID = 4
Private Const COLP_�̶� = 5
Private Const COLP_�Ƽ�ҽ�� = 6
Private Const COLP_��� = 7 '�շ��������
Private Const COLP_�շ���Ŀ = 8
Private Const COLP_��λ = 9
Private Const COLP_���� = 10
Private Const COLP_���� = 11
Private Const COLP_ִ�п��� = 12
Private Const COLP_�������� = 13
Private Const COLP_���� = 14
Private Const COLP_�շѷ�ʽ = 15
Private Const COLP_�շ���� = 16 '������
Private Const COLP_ִ�п���ID = 17
Private Const COLP_�������� = 18
Private Const COLP_�������� = 19

Public Function ShowMe(frmParent As Object, ByVal MainPrivs As String, ByVal int���� As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, _
    Optional ByVal lngҽ��ID As Long, Optional ByVal bln��ʿվ As Boolean, Optional blnRefresh As Boolean, Optional ByVal bytUseType As Byte, Optional ByVal strAdviceOfItem As String, _
    Optional ByVal dateStop As Date, Optional ByVal blnOnePati As Boolean, Optional ByVal strPatis As String, Optional ByVal blnAutoRead As Boolean, Optional ByVal lngӤ�� As Long, _
    Optional ByVal bln���� As Boolean, Optional ByVal lngҽ������ID As Long, Optional ByVal bln���͵��� As Boolean, Optional ByRef objMip As Object, Optional ByRef strPatisOut As String, _
    Optional ByVal bytSize As Byte, Optional ByVal strͣ��ԭ�� As String) As Boolean
'������blnRefresh=�Ƿ�ˢ������������
'      strPatis=����ʱ,��������ҽ���Ĳ���ID��������ʱ��ȷ��ֹͣ����ǰ����ѡ��Ĳ���ID��(����id,��ҳid;����id,��ҳid;......)
'      blnAutoRead=����ʱ������У������ҽ�������߷���ʱ����ȷ��ֹͣ
'      lngӤ��=ת��ҽ���´ﵯ��ҽ��ֹͣ����ʱ�Ŵ���
'      bln���͵���=ҽ������ʱ������У��ģʽ����ʱ����ˢ�������档
'      strPatisOut�������������ڴ���ʿվ��Ϣ���  ��ʽ "����id1,��ҳid1;����id2,��ҳid2;......"
    Set mfrmParent = frmParent
    mMainPrivs = MainPrivs
    mint���� = int����
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlngӤ�� = lngӤ��
    mlng����ID = lng����ID
    mlngҽ��ID = lngҽ��ID
    mbln��ʿվ = bln��ʿվ
    mlngҽ������ID = lngҽ������ID
    mbln���� = bln����
    mbytSize = bytSize
    mbln��������ִ�� = Val(zlDatabase.GetPara("������Ҫ����ִ��", glngSys)) = 1
    
    If gblnҽ����ֹԭ�� Then
        mstrͣ��ԭ�� = strͣ��ԭ��
    End If
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
     
    mblnOnePati = blnOnePati
    mblnAutoRead = blnAutoRead
    mbln���͵��� = bln���͵���
    If strPatis = "" Then
        mstr����IDs = mlng����ID & "," & mlng��ҳID
    Else
        mstr����IDs = strPatis
    End If
    
    mbytUseType = bytUseType
    mstrAdviceOfItem = strAdviceOfItem
    mdateStop = dateStop
    
    Me.Show 1, frmParent
    
    ShowMe = mblnOK
  
    If mblnOK Then blnRefresh = mblnRefresh
    strPatisOut = mstrPatiClsMsg
    
    If mblnOK And (int���� = 1 And Val(zlDatabase.GetPara("ҽ������ӡģʽ", glngSys, pסԺҽ���´�)) = 1 Or int���� = 2 Or int���� = 3) Then
        If Val(zlDatabase.GetPara("�Զ�����ҽ����ӡ", glngSys, pסԺҽ������)) = 1 Then
            Call frmAdvicePrint.ShowMe(frmParent, lng����ID, lng��ҳID, IIF(int���� = 2 Or int���� = 1, "ͣ����ӡ", "������ӡ"))
        End If
    End If
End Function

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.value = vNewValue
        txtPer.Text = CInt(psb.value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Private Sub cboʱ��_Click()
    If cboʱ��.ListIndex <> -1 Then
        txtʱ��.Text = Split(Split(con_Date, ",")(cboʱ��.ListIndex), "=")(1)
        If cboʱ��.ItemData(cboʱ��.ListIndex) = 1 Then
            txtʱ��.Text = Format(zlDatabase.Currentdate, "HH:mm")
        End If
        
        If Visible Then
            Call SetDefaultTime
        End If
    End If
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99   '������ҩ���
        If mblnPass Then
            Call gobjPass.zlPassCommandBarExe(mobjPassMap, Control.ID - conMenu_Edit_MediAudit * 10#)
        End If
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar Is Nothing Then Exit Sub
    '��ʿУ�Խ���˵�������ҽ��վ��һ��
    If mblnPass Then
        Call gobjPass.zlPASSPopupCommandBars(mobjPassMap, CommandBar, conMenu_Edit_MediAudit)
    End If
End Sub

Private Sub InitCommandBar()
'���ܣ���ʼ��������
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim objMenu As CommandBarPopup
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = frmIcons.imgMain.Icons
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҩ�����", -1, False)
    objMenu.ID = conMenu_EditPopup
    'PASS�����˵����ڵ�һ�ε������Ҽ�ʱ���ص�cbsMain_InitCommandsPopup
    
End Sub

Private Sub cmdAlley_Click()
'���ܣ��Բ��˹���ʷ/����״̬���в鿴
    If mblnPass Then
        Call gobjPass.zlPassCmdAlleyManage(mobjPassMap)
    End If
End Sub

Private Function ResetCond() As Boolean
'���ܣ�����У������
    Dim blnSeek As Boolean
    Me.Refresh
    With frmAdviceOperateCond
        .mMainPrivs = mMainPrivs
        .mint���� = mint����
        .mlng����ID = mlng����ID
        If mlngӤ������ID <> 0 Then
            If mlngӤ������ID = mlngҽ������ID Or mlngӤ������ID = mlngҽ������ID Then
                .mlng����ID = mlngӤ������ID
            End If
        End If
        .mlng����ID = mlng����ID
        .Show 1, Me
        If .mblnOK Then
            mlng����ID = .mlng����ID
            mstr����IDs = .mstr����IDs
            mlngҽ������ID = mlng����ID
            mint��Ч = .mint��Ч
            mint��� = .mint���
            mblnPauseLast = .mblnPauseLast
                        
            'ֻѡ���˵�ǰ���˲Ŷ�λ��ǰҽ��
            If UBound(Split(mstr����IDs, ";")) = 0 Then
                If Val(Split(mstr����IDs, ",")(0)) = mlng����ID Then blnSeek = True
            End If
            Call RefreshData(IIF(blnSeek, mlngҽ��ID, 0), True)
        End If
        ResetCond = .mblnOK
    End With
End Function

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        If tbr.Buttons("����").Visible Then
            If Not ResetCond Then Unload Me: Exit Sub
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("ȫѡ"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("ȫ��"))
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("ִ��"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbr_ButtonClick(tbr.Buttons("�˳�"))
    ElseIf KeyCode = vbKeyF1 Then
        Call tbr_ButtonClick(tbr.Buttons("����"))
    ElseIf KeyCode = vbKeyF5 Then
        Call tbr_ButtonClick(tbr.Buttons("ˢ��"))
    ElseIf KeyCode = vbKeyF12 Then
        If tbr.Buttons("����").Visible Then
            Call tbr_ButtonClick(tbr.Buttons("����"))
        End If
    ElseIf KeyCode = vbKeyF7 Then '�л����뷨
        If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
            If stbThis.Panels("WB").Bevel = sbrRaised Then
                Call stbThis_PanelClick(stbThis.Panels("WB"))
            Else
                Call stbThis_PanelClick(stbThis.Panels("PY"))
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, arrTmp As Variant, strTmp As String
    
    On Error GoTo errH
    mblnRefresh = False
    mblnReturn = False
    mblnFirstLoad = True
    Call InitAdviceTable
    Call SetAdviceCol '������һ��������,�Ա���ȷ�ָ����Ի�
    If mint���� = 2 Or mint���� = 3 Or mint���� = 4 Then
        Call InitPriceTable
    End If
    Call zlControl.SetPubFontSize(Me, mbytSize)
    Call RestoreWinState(Me, App.ProductName, mint����)
    
    strSQL = "Select Ӥ������ID,Ӥ������ID From ������ҳ Where ����ID=[1] and ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount > 0 Then
        mlngӤ������ID = Val(rsTmp!Ӥ������ID & "")
        mlngӤ������ID = Val(rsTmp!Ӥ������ID & "")
    End If
    
    '���ù�����ťͼ��
    Set tbr.HotImageList = frmIcons.imgColor
    Set tbr.ImageList = frmIcons.imgGray
    tbr.Buttons("ȫѡ").Image = "ȫѡ"
    tbr.Buttons("ȫ��").Image = "ȫ��"
    tbr.Buttons("ִ��").Image = "ִ��"
    tbr.Buttons("����").Image = "����"
    tbr.Buttons("ˢ��").Image = "ˢ��"
    tbr.Buttons("����").Image = "����"
    tbr.Buttons("�˳�").Image = "�˳�"
    tbr.ButtonHeight = 500
    
    'ȱʡʱ��ģʽ
    If mint���� = 3 Then
        arrTmp = Array(fraOper, lblOper, optOper(e��ǰʱ��), optOper(e��ʼʱ��), lblS, cboTime(e��ǰʱ��), lblB, cboTime(e��ʼʱ��))
        mstrȱʡУ��ʱ�� = zlDatabase.GetPara("ҽ��ȱʡУ��ʱ��", glngSys, pסԺҽ������, "001", arrTmp, InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��ѡ������;") > 0)
        If Len(mstrȱʡУ��ʱ��) = 1 Then mstrȱʡУ��ʱ�� = mstrȱʡУ��ʱ�� & "01"
    ElseIf (mint���� = 1 Or mint���� = 7) Then
        mblnHaveAudit = HaveAuditPriv(UserInfo.����)
        mlngͣ����� = Val(zlDatabase.GetPara("ʵϰҽ��ֹͣҽ����Ҫ���", glngSys, pסԺҽ���´�))
        arrTmp = Array(fraStop, lblStop, optStop(e��ǰʱ��), optStop(e��ʼʱ��), txtʱ��, cboʱ��, chkNoSend, chkRollSend)
        mstrȱʡֹͣʱ�� = zlDatabase.GetPara("ҽ��ȱʡֹͣʱ��", glngSys, pסԺҽ������, "011", arrTmp, InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��ѡ������;") > 0)
        If Len(mstrȱʡֹͣʱ��) = 1 Then mstrȱʡֹͣʱ�� = mstrȱʡֹͣʱ�� & "11"
    End If
    
    mblnOK = False
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    Select Case mint����
        Case 0
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrRaised
        Case 1
            stbThis.Panels("PY").Bevel = sbrRaised
            stbThis.Panels("WB").Bevel = sbrInset
        Case Else
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrInset
    End Select
    If Not (mint���� = 2 Or mint���� = 3 Or mint���� = 4) Or Not gbln����ƥ�䷽ʽ�л� Then
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
    mblnҽ������ = Val(zlDatabase.GetPara("ҽ��ҽ����������", glngSys, pסԺҽ������)) <> 0
    mbln��ʿǩ�� = Val(zlDatabase.GetPara("У��ҽ������ǩ��", glngSys, pסԺҽ������)) <> 0 And gintCA <> 0 And Mid(gstrESign, 2, 1) = "1"
    
    '�������ÿ�����,ȱʡ��������
    mblnFirst = True
    mblnPauseLast = False
    mint��Ч = 0: mint��� = 0
        
    '0-ҽ������,1-ֹͣҽ��,2-ȷ��ֹͣ,3-ҽ��У��,4-�����Ƽ���Ŀ,5-��ͣҽ��,6-����ҽ��
    If mbln��ʿվ And Not mblnAutoRead And InStr(",2,3,5,6,", mint����) > 0 Then
        tbr.Buttons("����").Enabled = Not mblnOnePati
    Else
        tbr.Buttons("����").Enabled = False
    End If
    tbr.Buttons("����").Visible = tbr.Buttons("����").Enabled 'Enabled�����ж�
    
    fraStop.Visible = False
    fraOper.Visible = False
    
    If mint���� = 0 Then
        Caption = "����ҽ������"
        tbr.Buttons("ִ��").Caption = "����"
        tbr.Buttons("ִ��").ToolTipText = "����ѡ���ҽ��(Ctrl+E)"
    ElseIf (mint���� = 1 Or mint���� = 7) Then
        Caption = "����ҽ��ֹͣ"
        If mint���� = 7 Then Caption = "����ͣ�����"
        tbr.Buttons("ִ��").Caption = IIF(mint���� = 7, "���", "ֹͣ")
        tbr.Buttons("ִ��").ToolTipText = IIF(mint���� = 7, "���", "ֹͣ") & "ѡ���ҽ��(Ctrl+E)"
        
        fraStop.Visible = True
        fraOper.Visible = False
        
        If mbln��ʿվ Then
            lblҽ��.Visible = True
            cboҽ��.Visible = True
        End If
        
        arrTmp = Split(con_Date, ",")
        cboʱ��.Clear
        For i = 0 To UBound(arrTmp)
            strTmp = Split(arrTmp(i), "=")(0)
            cboʱ��.AddItem strTmp
            If Split(arrTmp(i), "=")(1) = "__:__" Then
                cboʱ��.ItemData(cboʱ��.NewIndex) = 1
            End If
        Next
        cboʱ��.ListIndex = 0
    ElseIf mint���� = 2 Then
        Caption = "ȷ��ҽ��ֹͣ"
        tbr.Buttons("ִ��").Caption = "ȷ��"
        tbr.Buttons("ִ��").ToolTipText = "ȷ��ѡ���ҽ��(Ctrl+E)"
    
        picUD.Visible = True
        vsPrice.Visible = True
    ElseIf mint���� = 3 Then
        Caption = "����ҽ��У��"
        tbr.Buttons("ִ��").Caption = "У��"
        tbr.Buttons("ִ��").ToolTipText = "ȷ��ѡ���ҽ��(Ctrl+E)"
                
        stbThis.Panels(4).Visible = True
        stbThis.Panels(5).Visible = True
        
        picUD.Visible = True
        vsPrice.Visible = True
        fraStop.Visible = True
                
        '���˹���ʷ/����״̬���ü�� У��ʱ
        Call zlPASSMap
        If mblnPass Then      'Pass
            Call InitCommandBar
            Call gobjPass.zlPassCmdAlleyEnable(mobjPassMap)
        End If
        For i = e���� To e����
            cboTime(i).AddItem "��ʼʱ��"
            cboTime(i).AddItem "����ʱ��"
        Next
        fraStop.Visible = False
        fraOper.Visible = True
    ElseIf mint���� = 4 Then
        Caption = "�����Ƽ���Ŀ"
        tbr.Buttons("ִ��").Caption = "ȷ��"
        tbr.Buttons("ִ��").ToolTipText = "ȷ��ѡ����Ŀ�ļ�Ŀ(Ctrl+E)"
        
        picUD.Visible = True
        vsPrice.Visible = True
    ElseIf mint���� = 5 Then
        Caption = "����ҽ����ͣ"
        tbr.Buttons("ִ��").Caption = "��ͣ"
        tbr.Buttons("ִ��").ToolTipText = "��ͣѡ���ҽ��(Ctrl+E)"
    ElseIf mint���� = 6 Then
        Caption = "����ҽ������"
        tbr.Buttons("ִ��").Caption = "����"
        tbr.Buttons("ִ��").ToolTipText = "����ѡ���ҽ��(Ctrl+E)"
    End If
    
    Call SetFilterTime
    
    '��ȡ������Ϣ
    If mint���� = 2 Or mint���� = 3 Or mint���� = 4 Then
        strSQL = "Select ID,���� From ���ű� Where վ��='" & gstrNodeNo & "' Or վ�� is Null"
        Set mrsDept = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrsDept, strSQL, Me.Caption)
    End If
    
    '��ʾ������Ϣ��һ�����˲��������(����ҽ��У�Կɲ�������ID)
    If mlng����ID = 0 And mint���� = 3 Then
        lblPati.Caption = ""
        mint���� = 0
        mlng�������� = 0
    Else
        strSQL = _
            " Select B.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� ,NVL(B.����,A.����) ����,B.��Ժ����," & _
            " B.סԺҽʦ,B.��Ժ����ID,C.���� as ����,B.����,B.�������� " & _
            " From ������Ϣ A,������ҳ B,���ű� C" & _
            " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
            " And A.����ID=[1] And B.��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        lblPati.Caption = "����:" & rsTmp!���� & "��סԺ��:" & NVL(rsTmp!סԺ��) & _
            "������:" & NVL(rsTmp!��Ժ����) & "������:" & NVL(rsTmp!����)
        mint���� = NVL(rsTmp!����, 0)
        mlng�������� = Val(rsTmp!�������� & "")
        
        '��ѡ��ͣ��ҽ��:ȱʡΪ���˵�סԺҽʦ���˿��ҵĵ�һ��ҽ��
        'Ŀǰ��֧������ֹͣҽ��,��˿϶����Դ���ĵ�ǰ����Ϊ׼��ȡ
        If (mint���� = 1 Or mint���� = 7) And mbln��ʿվ Then
            Call Get����ҽ��(rsTmp!��Ժ����ID, True, NVL(rsTmp!סԺҽʦ), 0, cboҽ��)
            If cboҽ��.ListIndex = -1 And cboҽ��.ListCount > 0 Then cboҽ��.ListIndex = 0
        End If
    End If
    
    '��ʾҽ������
    If Not tbr.Buttons("����").Enabled Then
        Call RefreshData(mlngҽ��ID, True)
        If (mblnAutoRead Or mblnOnePati) And (mint���� = 2 Or mint���� = 3) Then Call tbr_ButtonClick(tbr.Buttons("ȫѡ"))
    End If
    If mint���� = 3 And mblnAutoRead Then
        stbThis.Panels(2).Text = "����ҽ�����ڷ���ǰ��У�ԣ���ΪУ�Ժ��Զ�ֹͣ��������ҽ����"
    End If
    
    '����ȱʡʱ��
    If (mint���� = 1 Or mint���� = 3) And mdateStop = CDate(0) Then
        Call SetDefaultTime
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ReadMsg()
'���ܣ���鲢��������Ϣ���жϵ�ǰ��ѡ�����Ƿ��п���У�Ե�ҽ����
'˵����
    Dim rsTmp As ADODB.Recordset
    Dim rsMsg As ADODB.Recordset
    Dim strSQL As String
    Dim strPatis As String
    Dim strPati As String
    Dim strPatiClsMsg As String
    Dim strCurDate As String
    Dim lng����ID As Long, lng��ҳID As Long
    Dim i As Long, j As Long
    Dim blnTrans As Boolean
    Dim varArr As Variant
    Dim arrSQL As Variant
    Dim lngҽ��ID As Long
    Dim int���� As Integer
    Dim strMsgNo As String
    Dim strWhere As String
    
    If Not (mint���� = 2 Or mint���� = 3) Then Exit Sub
    
    On Error GoTo errH
    
    arrSQL = Array()
    
    strPatis = mstr����IDs
    strPatis = Replace(strPatis, ",", ":")
    strPatis = Replace(strPatis, ";", ",")
    If mint���� = 3 Then
        strMsgNo = "ZLHIS_CIS_001"
        
        If gblnKSSStrict Or gbln�����ּ����� Or gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ Then
            strWhere = strWhere & " And (Nvl(A.���״̬,0) Not in(1,3,7" & IIF(gblnѪ��ϵͳ = True, "", ",4,5") & ") or a.ҽ����Ч=0 and a.���״̬=1 and a.������־=1 and (instr(',5,6,',A.�������)>0 or A.�������='E' and B.��������='2'))"
        End If
        
        strSQL = "select a.id as ҽ��ID,nvl(a.������־,0) as ����,a.����ID,a.��ҳID from ����ҽ����¼ a,������ĿĿ¼ b where a.������Ŀid=b.id(+) and A.ҽ��״̬=1" & strWhere & _
            " And Exists ( Select 1 From ��Ա�� M,ִҵ��� N" & _
            " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
            " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ'))"
    Else
        strMsgNo = "ZLHIS_CIS_002"
        
        strSQL = "select a.id as ҽ��ID,nvl(a.������־,0) as ����,a.����ID,a.��ҳID from ����ҽ����¼ a where A.ҽ��״̬=8 and Nvl(a.ҽ����Ч,0)=0"
    End If
    strSQL = strSQL & " And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3 And (a.����ID,a.��ҳID) In (Select C1,C2 From Table(f_Num2list2([1])))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPatis)
    
    strPatis = mstr����IDs
    varArr = Split(strPatis, ";")
    
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    For i = 0 To UBound(varArr)
        strPati = varArr(i)
        lng����ID = Split(strPati, ",")(0)
        lng��ҳID = Split(strPati, ",")(1)
        
        rsTmp.Filter = "����ID=" & lng����ID & " and ��ҳID=" & lng��ҳID
        
        If rsTmp.EOF Then
            '���ò��˵���Ϣ��Ϊ����
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & strMsgNo & "',3,'" & UserInfo.���� & "'," & mlng����ID & ",To_Date('" & strCurDate & "','YYYY-MM-DD HH24:MI:SS'))"
            strPatiClsMsg = strPatiClsMsg & ";" & lng����ID & "," & lng��ҳID
        Else
            '�����ݸ�����Ϣ�嵥�жϣ��Ƿ����һ����Ϣ
            lngҽ��ID = rsTmp!ҽ��ID: int���� = 1
            rsTmp.Filter = "����ID=" & lng����ID & " and ��ҳID=" & lng��ҳID & " and ����=1"
            If Not rsTmp.EOF Then
                lngҽ��ID = rsTmp!ҽ��ID
                int���� = 2
            End If
            strSQL = "select 1 From ҵ����Ϣ�嵥 A Where a.����id=[1] And a.����id=[2] And a.���ͱ��� =[3] And a.���ȳ̶�=[4] And a.�Ƿ�����=0 And Rownum<2"
            Set rsMsg = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, strMsgNo, int����)
            If rsMsg.EOF Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_ҵ����Ϣ�嵥_Insert(" & lng����ID & "," & lng��ҳID & ",null," & mlng����ID & ",2,'����" & IIF(mint���� = 3, "�´�", "ֹͣ") & "ҽ����','0010','" & strMsgNo & "'," & lngҽ��ID & "," & int���� & ",0,null," & mlng����ID & ")"
            End If
        End If
        rsTmp.Filter = 0
    Next
    
    If UBound(arrSQL) <> -1 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    mstrPatiClsMsg = Mid(strPatiClsMsg, 2)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshData(Optional ByVal lngҽ��ID As Long, Optional ByVal blnNotify As Boolean)
'���ܣ�ˢ������
'������lngҽ��ID=����ҽ����λ
'      blnNotify=�Ƿ���������ҽ��
    Dim blnChange As Boolean, i As Long
    Dim strPatis As String, arrPatis As Variant
    Dim lng����ID As Long, lng��ҳID As Long
    Dim strMsg As String, strTmp As String
    Dim blnSelect As Boolean
    
    '��ʾҽ������
    Call LoadAdvice(strPatis)
    
    '��ȡ�Ƽ�����
    If mint���� = 2 Or mint���� = 3 Or mint���� = 4 Then
        Call InitPriceRecordset
        Screen.MousePointer = 11
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            Progress = i / (vsAdvice.Rows - 1) * 100
            blnChange = False
            Call LoadPrice(i, blnChange)
            If blnChange And mint���� = 4 Then Call SelectRow(i): blnSelect = True
        Next
        Call AppendPriceItem
        Progress = 0: Screen.MousePointer = 0
    End If
    
    If lngҽ��ID <> 0 Then
        i = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
        If i <> -1 Then vsAdvice.Row = i
    End If
    If vsAdvice.Rows = vsAdvice.FixedRows + 1 And blnSelect = False Then
        If Val(vsAdvice.TextMatrix(vsAdvice.Rows - 1, COL_ID)) <> 0 Then
            Call SelectRow(vsAdvice.Rows - 1): blnSelect = True
        End If
    End If
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    '����ҽ������
    If blnNotify And InStr(",3,4,6,", mint����) > 0 And strPatis <> "" Then
        arrPatis = Split(strPatis, ";")
        For i = 0 To UBound(arrPatis)
            lng����ID = Split(arrPatis(i), ",")(0)
            lng��ҳID = Split(arrPatis(i), ",")(1)
            strTmp = ExistsSpecAdvice(lng����ID, lng��ҳID)
            If strTmp <> "" Then
                strTmp = Replace(Replace(strTmp, "��������", ""), vbCrLf & vbCrLf, vbCrLf)
                strMsg = strMsg & vbCrLf & strTmp
            End If
        Next
        If strMsg <> "" Then MsgBox Mid(strMsg, 3), vbInformation, gstrSysName & " - ������"
    End If
End Sub

Private Sub SelectRow(ByVal lngRow As Long)
'���ܣ�ʹָ����ѡ��(����һ����ҩ)
    With vsAdvice
        If mint���� = 3 Then
            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
            .Cell(flexcpData, lngRow, COL_ѡ��) = 1
        Else
            .TextMatrix(lngRow, COL_ѡ��) = -1 'ֱ�Ӷ�TextMatrixʱ,��Ҫ��True
        End If
    End With
    Call vsAdvice_AfterEdit(lngRow, COL_ѡ��)
End Sub

Private Sub Form_Resize()
    Dim lngTmp As Long
    
    On Error Resume Next
    
    lblPati.Left = 150
    lblPati.Top = 120
    
    If InStr(",1,3,7,", mint����) > 0 Then
        picPati.Height = cmdAlley.Height + cboʱ��.Height + 200
    Else
        picPati.Height = cmdAlley.Height + 50
    End If
    
    vsAdvice.Height = Me.ScaleHeight - cbr.Height - stbThis.Height - picPati.Height _
        - IIF(picUD.Visible, picUD.Height + vsPrice.Height, 0) - IIF(pic����.Visible, pic����.Height, 0)
    
    lngTmp = optBaby(e����).Width + optBaby(e����).Width + optBaby(eӤ��).Width + 40
    fraBaby.Width = lngTmp
    fraBaby.Height = optBaby(e����).Height
    optBaby(e����).Left = 10
    Call zlControl.SetPubCtrlPos(False, 0, lblPati, 30, fraBaby, 30, cmdAlley)
    Call zlControl.SetPubCtrlPos(False, 0, optBaby(e����), 10, optBaby(e����), 10, optBaby(eӤ��))
 
    If cmdAlley.Visible Then
        cmdAlley.Left = Me.ScaleWidth - cmdAlley.Width - 200
        lngTmp = cmdAlley.Left - lngTmp
    Else
        lblPati.Width = Me.ScaleWidth - lblPati.Left
        lngTmp = Me.ScaleWidth - lngTmp
    End If
    fraBaby.Left = lngTmp
    
    txt����.Width = pic����.ScaleWidth - txt����.Left - 30
    psb.Top = stbThis.Top + 60
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - 100
    psb.Left = stbThis.Panels(2).Left + 30
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
    
    fraStop.Left = 0
    fraStop.Height = cboʱ��.Height + 80
    fraStop.Top = cmdAlley.Height + 80
    fraStop.Width = Me.ScaleWidth
    
    lblStop.Left = 150
    lblStop.Top = 50
    Call zlControl.SetPubCtrlPos(False, 0, lblStop, 30, optStop(e�ϴ�ִ��ʱ��), 10, optStop(eָ��ʱ��), 10, cboʱ��, 10, txtʱ��, 200, chkNoSend, 10, chkRollSend, 10, lblҽ��, 10, cboҽ��)
        
    cboҽ��.Left = Me.ScaleWidth - cboҽ��.Width - lblStop.Left
    lblҽ��.Left = cboҽ��.Left - lblҽ��.Width - 100
    
    fraOper.Width = Me.ScaleWidth
    fraOper.Top = cmdAlley.Height + 80
    fraOper.Height = cboʱ��.Height + 20
    fraOper.Left = 0
    lblOper.Top = 50
    lblOper.Left = 150
    Call zlControl.SetPubCtrlPos(False, 0, lblOper, 30, optOper(e��ǰʱ��), 10, optOper(e��ʼʱ��), 100, lblS, 15, cboTime(e����), 50, lblB, 15, cboTime(e����))
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnSetup As Boolean
    
    Call SaveWinState(Me, App.ProductName, mint����)
    
    '�������ò���
    blnSetup = InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��ѡ������;") > 0
    If mint���� = 3 Then
        Call zlDatabase.SetPara("ҽ��ȱʡУ��ʱ��", IIF(optOper(e��ǰʱ��).value, 0, 1) & IIF(cboTime(e����).ListIndex = -1, 0, cboTime(e����).ListIndex) & IIF(cboTime(e����).ListIndex = -1, 1, cboTime(e����).ListIndex), glngSys, pסԺҽ������, blnSetup)
    ElseIf mint���� = 1 Or mint���� = 7 Then
        If optStop(e�ϴ�ִ��ʱ��).value Then
            Call zlDatabase.SetPara("ҽ��ȱʡֹͣʱ��", "1", glngSys, pסԺҽ������, blnSetup)
        Else
            Call zlDatabase.SetPara("ҽ��ȱʡֹͣʱ��", "0" & IIF(chkNoSend.value = 1, "1", "0") & IIF(chkRollSend.value = 1, "1", "0"), glngSys, pסԺҽ������, blnSetup)
        End If
    End If
    
    Set mrsPrice = Nothing
    Set mrsDept = Nothing
    mMainPrivs = ""
    mlngҽ��ID = 0
    mint���� = 0
    mlng����ID = 0
    mlng����ID = 0
    mlng��ҳID = 0
    mbln��ʿվ = False
    mblnAll = False
    Set mclsMipModule = Nothing
    Set mobjPassMap = Nothing
End Sub

Private Sub SetDefaultTime()
'���ܣ����ݽ������ã�����У�Ի�ֹͣҽ����ȱʡʱ��
    Dim i As Long, vCurDate As Date
    
    vCurDate = zlDatabase.Currentdate
    
    '����ʱ��ֵ
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ID)) <> 0 Then
                If mint���� = 3 Then
                    If optOper(e��ǰʱ��).value Then
                        If .TextMatrix(i, COL_��־) = "��¼" Then
                            .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_����ʱ��))), "yyyy-MM-dd HH:mm")
                        Else
                            .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                        End If
                    Else
                        If Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") > Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then
                            .TextMatrix(i, COL_����) = .Cell(flexcpData, i, IIF(cboTime(e����).ListIndex = 0, COL_��ʼʱ��, COL_����ʱ��))
                        Else
                            .TextMatrix(i, COL_����) = .Cell(flexcpData, i, IIF(cboTime(e����).ListIndex = 0, COL_��ʼʱ��, COL_����ʱ��))
                        End If
                    End If
                Else
                    If optStop(eָ��ʱ��).value Then
                        '��ǰʱ���ָ��ʱ��
                        '��������ֹͣʱ�䲻Ϊ�գ������ǻ����¼���ͱ߸ı�ʱ�䡣
                        If mdateStop = CDate(0) Or Not (.TextMatrix(i, COL_�������) = "H" And .TextMatrix(i, COL_��������) = "1" Or .TextMatrix(i, COL_�������) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_��������) & ",") > 0) Then
                            .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd " & txtʱ��.Text)
                            
                            If .TextMatrix(i, COL_Ƶ��) <> "������" Then
                                If chkNoSend.value = 0 Then    '���������
                                    If .TextMatrix(i, COL_�ϴ�ִ��) = "" Then
                                        .TextMatrix(i, COL_����) = .TextMatrix(i, COL_��ʼʱ��)
                                    ElseIf .TextMatrix(i, COL_�ϴ�ִ��) < .TextMatrix(i, COL_����) Then
                                        .TextMatrix(i, COL_����) = .TextMatrix(i, COL_�ϴ�ִ��)
                                    End If
                                End If
                                If chkRollSend.value = 0 Then    '������ջ�
                                    If .TextMatrix(i, COL_�ϴ�ִ��) > .TextMatrix(i, COL_����) Then
                                        .TextMatrix(i, COL_����) = .TextMatrix(i, COL_�ϴ�ִ��)
                                    End If
                                End If
                            End If
                        End If
                    Else
                        'ĩ��ʱ�����Ϊ�գ���Ϊ��ʼʱ��
                        If mdateStop = CDate(0) Or Not (.TextMatrix(i, COL_�������) = "H" And .TextMatrix(i, COL_��������) = "1" Or .TextMatrix(i, COL_�������) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_��������) & ",") > 0) Then
                            If .TextMatrix(i, COL_�������) = "H" Or .TextMatrix(i, COL_Ƶ��) = "������" Then
                                '�������ʱ��Ϊ�գ�����ȼ����⴦��Ĭ��Ϊ��ǰʱ�䡣
                                .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            Else
                                If .TextMatrix(i, COL_�ϴ�ִ��) = "" Then
                                    .TextMatrix(i, COL_����) = .TextMatrix(i, COL_��ʼʱ��)
                                Else
                                    .TextMatrix(i, COL_����) = .TextMatrix(i, COL_�ϴ�ִ��)
                                End If
                            End If
                        End If
                    End If
                    
                    .TextMatrix(i, COL_����) = GetValidateStopDate(.TextMatrix(i, COL_����), i)
                End If
                
                .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
            End If
        Next
    End With
End Sub

Private Sub optStop_Click(Index As Integer)
    If Me.Visible Then
        If (mint���� = 1 Or mint���� = 7) Then
            chkNoSend.Visible = optStop(eָ��ʱ��).value
            chkRollSend.Visible = optStop(eָ��ʱ��).value
        End If
        If optStop(eָ��ʱ��).value Then
            If cboʱ��.ListIndex <> -1 Then
                If cboʱ��.ItemData(cboʱ��.ListIndex) = 1 Then
                    txtʱ��.Text = Format(zlDatabase.Currentdate, "HH:mm")
                End If
            End If
        End If
        Call SetDefaultTime
    End If
End Sub

Private Sub optOper_Click(Index As Integer)
    If Index = e��ǰʱ�� Then
        lblS.Visible = False
        lblB.Visible = False
        cboTime(e����).Visible = False
        cboTime(e����).Visible = False
    Else
        lblS.Visible = True
        lblB.Visible = True
        cboTime(e����).Visible = True
        cboTime(e����).Visible = True
    End If
    If Me.Visible Then Call SetDefaultTime
End Sub

Private Sub optBaby_Click(Index As Integer)
    mintҽ������Χ = Index
    If Not mblnFirstLoad Then
    Call RefreshData(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    End If
End Sub

Private Sub txtʱ��_GotFocus()
    Call zlControl.TxtSelAll(txtʱ��)
End Sub

Private Sub txtʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        optStop(eָ��ʱ��).value = True
        If IsDate("2010-06-22 " & txtʱ��.Text) Then
            Call SetDefaultTime
            Call zlControl.TxtSelAll(txtʱ��)
        Else
            txtʱ��.Text = "__:__"
        End If
    End If
End Sub

Private Function GetValidateStopDate(ByVal strDate As String, ByVal lngRow As Long) As String
'���ܣ���ȡ��Ч��ֹͣʱ��
        
    strDate = Format(strDate, "yyyy-MM-dd HH:mm")
    With vsAdvice
        
        '��ӦС�ڿ�ʼִ��ʱ��
        If strDate < Format(.Cell(flexcpData, lngRow, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
            strDate = Format(.Cell(flexcpData, lngRow, COL_��ʼʱ��), "yyyy-MM-dd HH:mm")
        End If
    
    End With
    GetValidateStopDate = strDate
End Function

Private Sub txtʱ��_Validate(Cancel As Boolean)
    If IsDate("2010-06-22 " & txtʱ��.Text) = False Then
        Cancel = True
    End If
End Sub

Private Sub picUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAdvice.Height + Y < 1000 Or vsPrice.Height - Y < 500 Then Exit Sub
        vsAdvice.Height = vsAdvice.Height + Y
        vsPrice.Height = vsPrice.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        Call zlDatabase.SetPara("���뷽ʽ", IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0)))
        mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long, blnAutoRoll As Boolean
    
    Select Case Button.Key
        Case "ȫѡ"
            If vsAdvice.ColHidden(COL_ѡ��) Then Exit Sub
            If vsAdvice.Rows = vsAdvice.FixedRows Then Exit Sub
            If vsAdvice.Rows = vsAdvice.FixedRows + 1 And Val(vsAdvice.TextMatrix(vsAdvice.FixedRows, COL_ID)) = 0 Then Exit Sub
            
            If mint���� = 3 Then
                For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
                    If vsAdvice.Cell(flexcpData, i, COL_ѡ��) = Empty Then '�������ʵĲ���
                        Set vsAdvice.Cell(flexcpPicture, i, COL_ѡ��) = frmIcons.imgTrueFalse.ListImages("T").Picture
                        vsAdvice.Cell(flexcpData, i, COL_ѡ��) = 1
                    End If
                Next
            Else
                'flexcpText��ͬ��.TextMatrix(lngRow, COL_ѡ��) = -1(True)
                vsAdvice.Cell(flexcpText, vsAdvice.FixedRows, COL_ѡ��, vsAdvice.Rows - 1, COL_ѡ��) = "-1"
            End If
        Case "ȫ��"
            If mint���� = 3 Then
                Set vsAdvice.Cell(flexcpPicture, vsAdvice.FixedRows, COL_ѡ��, vsAdvice.Rows - 1, COL_ѡ��) = Nothing
                vsAdvice.Cell(flexcpData, vsAdvice.FixedRows, COL_ѡ��, vsAdvice.Rows - 1, COL_ѡ��) = Empty
            Else
                vsAdvice.Cell(flexcpText, vsAdvice.FixedRows, COL_ѡ��, vsAdvice.Rows - 1, COL_ѡ��) = "0"
            End If
        Case "ˢ��"
            Call RefreshData(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
        Case "����"
            Call ResetCond
        Case "ִ��"
            Dim bln�����ջ� As Boolean
            
            If Not CheckValid(bln�����ջ�) Then Exit Sub
            If Not CheckSignValid Then Exit Sub
            If ExecuteOperate Then
                If mblnHaveAudit Or (mint���� <> 1 And mint���� <> 7) Then
                    'ҽ��У��ʱ��鲢���ѳ����ջ�(�Զ�)ֹͣ��ҽ��
                    If mint���� = 3 And mstrRollNotify <> "" Then
                        Call ShowRollNotify(mstrRollNotify)
                    ElseIf mint���� = 2 And bln�����ջ� Then
                        If PatiCanBilling(mlng����ID, mlng��ҳID, GetInsidePrivs(pסԺҽ������), pסԺҽ������) Then
                            blnAutoRoll = Val(zlDatabase.GetPara("ֹͣ���Զ������ջ�", glngSys, pסԺҽ������)) = 1
                            Me.Hide
                            If frmAdviceRollSend.ShowMe(mfrmParent, mMainPrivs, mlng����ID, mlng����ID, mlng��ҳID, True, blnAutoRoll, mlngҽ������ID, mlngӤ������ID) Then   '������ģʽ������������ѡ����
                                If blnAutoRoll Then
                                    MsgBox "ȷ��ֹͣ�󣬳��ڷ��͵�ҽ�����Զ��ջء�", vbInformation, gstrSysName
                                End If
                            End If
                        End If
                    End If
                End If
                mblnOK = True: Unload Me
            End If
        Case "����"
            ShowHelp App.ProductName, Me.hwnd, Me.Name
        Case "�˳�"
            Unload Me
    End Select
End Sub

Private Sub txt����_Change()
    If Not pic����.Visible Then Exit Sub
    
    With vsAdvice
        .TextMatrix(.Row, COL_����˵��) = txt����.Text
    End With
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If txt����.MaxLength <> 0 And Not (KeyAscii >= 0 And KeyAscii < 32) Then
        If zlCommFun.ActualLen(txt����.Text) > txt����.MaxLength Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If vsAdvice.Col = COL_��ֹԭ�� Then
        vsAdvice.ComboList = "..."
        vsAdvice.Editable = flexEDKbdMouse
    Else
        vsAdvice.ComboList = ""
    End If

    If NewRow = OldRow Then Exit Sub
    'PASS
    If mblnPass And mint���� = 3 Then
        Call gobjPass.zlPassSetDrug(mobjPassMap)
    End If
    
    With vsAdvice
        'У������˵��
        If mint���� = 3 Then
            If .Cell(flexcpData, .Row, COL_ѡ��) = 2 Then
                txt����.Text = .TextMatrix(.Row, COL_����˵��)
                pic����.Visible = True
            Else
                pic����.Visible = False
            End If
            Call Form_Resize
        End If
        
        '��ʾ�Ƽ���Ŀ
        If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
            If (mint���� = 2 Or mint���� = 3 Or mint���� = 4) And Not mrsPrice Is Nothing Then
                Call ShowPrice(NewRow)
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col_ҽ������ Then
        vsAdvice.AutoSize Col
    ElseIf Col = COL_Ƥ�� Then
        If vsAdvice.ColWidth(Col) > 1200 Then vsAdvice.ColWidth(Col) = 1200
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_ѡ�� Or Col = COL_���� Or Col = COL_��ֹԭ�� Or Col = COL_��ʾ Then Cancel = True 'Pass
End Sub

Private Sub vsAdvice_DblClick()
    
    If mblnPass And mint���� = 3 Then
        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap)
    End If
    
    With vsAdvice
        If mint���� = 3 And .MouseCol = COL_ѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_����: lngRight = COL_��ʼʱ��
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_Ƶ��: lngRight = COL_�÷�
        End If
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_Ƥ��: lngRight = COL_Ƥ��
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then
        '���ֱ�����뺺�ֵ�����
        Call vsAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
'���ܣ���λ����һ���뵥Ԫ������У�Ա�־
    Dim blnGroup As Boolean, i As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        With vsAdvice
            If .ColHidden(COL_ѡ��) And .ColHidden(COL_����) Then
                If .Row + 1 <= .Rows - 1 Then
                    .Row = .Row + 1
                Else
                    .Row = .FixedRows
                End If
            Else
                If .Col = COL_ѡ�� Then
                    If Not .ColHidden(COL_����) Then
                        .Col = COL_����
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            .Row = .Row + 1
                        Else
                            .Row = .FixedRows
                        End If
                    End If
                ElseIf .Col = COL_���� Then
                    If Not .ColHidden(COL_��ֹԭ��) Then
                        .Col = COL_��ֹԭ��
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            .Row = .Row + 1
                        Else
                            .Row = .FixedRows
                        End If
                        .Col = COL_ѡ��
                    End If
                ElseIf .Col = COL_��ֹԭ�� Then
                    If .Row + 1 <= .Rows - 1 Then
                        .Row = .Row + 1
                    Else
                        .Row = .FixedRows
                    End If
                    .Col = COL_ѡ��
                Else
                    If .Row + 1 <= .Rows - 1 Then
                        .Row = .Row + 1
                    Else
                        .Row = .FixedRows
                    End If
                    If Not .ColHidden(COL_ѡ��) Then .Col = COL_ѡ��
                End If
            End If
            Call .ShowCell(.Row, .Col)
        End With
    ElseIf KeyAscii = 32 Then
        With vsAdvice
            If mint���� = 3 And .Col = COL_ѡ�� Then
                KeyAscii = 0
                
                If .Cell(flexcpData, .Row, .Col) = Empty Then
                    Set .Cell(flexcpPicture, .Row, .Col) = frmIcons.imgTrueFalse.ListImages("T").Picture
                    .Cell(flexcpData, .Row, .Col) = 1
                ElseIf .Cell(flexcpData, .Row, .Col) = 1 Then
                    Set .Cell(flexcpPicture, .Row, .Col) = frmIcons.imgQuestion.ListImages("Q").Picture
                    .Cell(flexcpData, .Row, .Col) = 2
                    If .TextMatrix(.Row, COL_�������) = "K" And Val(.TextMatrix(.Row, COL_���״̬)) <> 0 Then
                        MsgBox "��ҽ��Ϊ����˵���Ѫҽ��������ΪУ�����ʡ�", vbInformation, gstrSysName
                        Set .Cell(flexcpPicture, .Row, .Col) = Nothing
                        .Cell(flexcpData, .Row, .Col) = Empty
                    End If
                ElseIf .Cell(flexcpData, .Row, .Col) = 2 Then
                    Set .Cell(flexcpPicture, .Row, .Col) = Nothing
                    .Cell(flexcpData, .Row, .Col) = Empty
                End If
                Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
                
                If InStr(",5,6,", .TextMatrix(.Row, COL_�������)) > 0 Then
                    If .Row - 1 >= .FixedRows Then
                        blnGroup = Val(.TextMatrix(.Row - 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID))
                    End If
                    If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                        blnGroup = Val(.TextMatrix(.Row + 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID))
                    End If
                    If blnGroup Then
                        For i = .Row - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then
                                Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                                .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Else
                                Exit For
                            End If
                        Next
                        For i = .Row + 1 To .Rows - 1
                            If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then
                                Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                                .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                'һ��������ŵĻ���ҽ��
                If Val(.TextMatrix(.Row, COL_�������)) <> 0 And .TextMatrix(.Row, COL_�������) = "Z" Then
                    For i = .Row - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(i, COL_�������)) = Val(.TextMatrix(.Row, COL_�������)) Then
                            .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                        Else
                            Exit For
                        End If
                    Next
                    For i = .Row + 1 To .Rows - 1
                        If Val(.TextMatrix(i, COL_�������)) = Val(.TextMatrix(.Row, COL_�������)) Then
                            .Cell(flexcpData, i, .Col) = .Cell(flexcpData, .Row, .Col)
                            Set .Cell(flexcpPicture, i, .Col) = .Cell(flexcpPicture, .Row, .Col)
                        Else
                            Exit For
                        End If
                    Next
                End If
                
            End If
        End With
    Else
        If vsAdvice.Col = COL_��ֹԭ�� Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsAdvice_CellButtonClick(vsAdvice.Row, vsAdvice.Col)
            Else
                vsAdvice.ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End If
End Sub

Private Function AcceptInput(ByVal Row As Long, ByVal Col As Long) As Boolean
    Dim strTmp As String, vPause As Date
    Dim blnDoSame As Boolean
    Dim lngҽ��ID As Long
    Dim strTmpTim As String
    
    AcceptInput = False
    With vsAdvice
        If .EditText <> "" Then .EditText = zlStr.FullDate(.EditText)
        If .EditText = .TextMatrix(Row, Col) Then AcceptInput = True: Exit Function
    
        '����������Ч��
        If Not IsDate(.EditText) Then
            MsgBox "������һ����Ч��" & .TextMatrix(0, Col) & " ��", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
    
        If (mint���� = 1 Or mint���� = 7) Then '�����ֹʱ��
            '������ڿ�ʼִ��ʱ��
            If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                MsgBox "�����ִ����ֹʱ��������ҽ���Ŀ�ʼִ��ʱ�� " & Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
            
            'ͣҽ��ʱ�����ִ�еǼ����
            If mint���� = 1 And Not (.TextMatrix(Row, COL_�������) = "Z" And InStr(",4,14,9,10,12,", "," & .TextMatrix(Row, COL_��������) & ",") > 0) Then
                '��ȡҽ��id
                lngҽ��ID = IIF(InStr(",5,6,", .TextMatrix(Row, COL_�������)) > 0, .TextMatrix(Row, COL_���ID), .TextMatrix(Row, COL_ID))
                '��ȡʱ��
                strTmpTim = GetAdviceStopTime(lngҽ��ID)
                '��Ϣ��ʾ
                If IsDate(strTmpTim) Then
                    If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(strTmpTim, "yyyy-MM-dd HH:mm") Then
                        strTmp = .EditText 'MsgBoxһ��,EditText�Ϳ���,����Ҫ��¼
                        MsgBox "����ֹͣ��ִ��ʱ�� " & strTmpTim & " ֮ǰ�������ֹͣʱ�䣬���ȷʵҪֹͣ��ִ��ʱ��֮ǰ������ȡ��ִ�еǼǡ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
            
            '��ӦС���ϴ�ִ��ʱ��
            If IsDate(.Cell(flexcpData, Row, COL_�ϴ�ִ��)) Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then
                    strTmp = .EditText 'MsgBoxһ��,EditText�Ϳ���,����Ҫ��¼
                    If MsgBox("�����ִ����ֹʱ��С��ҽ�����ϴ�ִ��ʱ�� " & Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                    End If
                End If
            End If
            
        ElseIf mint���� = 2 Then  '���ȷ��ֹͣʱ��
            If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then
                MsgBox "ȷ��ֹͣҽ����ʱ�䲻��С��ҽ����ִ����ֹʱ�� " & Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
        ElseIf mint���� = 3 Then  '���У��ʱ��
            If Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") >= Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then
                    MsgBox "�����У��ʱ�䲻��С��ҽ���Ŀ���ʱ�� " & Format(.Cell(flexcpData, Row, COL_����ʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            Else
                If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                    MsgBox "�����У��ʱ�䲻��С��ҽ���Ŀ�ʼִ��ʱ�� " & Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
        ElseIf mint���� = 5 Then '�����ͣʱ��
            'Ӧ>=��ʼִ��ʱ��,��Ϊ��ʱ�����δִ��
            If Format(.EditText, "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                MsgBox "ҽ������ͣʱ��Ӧ���ڵ��ڿ�ʼִ��ʱ�� " & Format(.Cell(flexcpData, Row, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
            'Ӧ>�ϴ�ִ��ʱ��,��Ϊ��ʱ�����ִ��
            If .TextMatrix(Row, COL_�ϴ�ִ��) <> "" Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then
                    MsgBox "ҽ������ͣʱ��Ӧ�����ϴ�ִ��ʱ�� " & Format(.Cell(flexcpData, Row, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
            'Ӧ<ִ����ֹʱ��,��Ϊ��ʱ���ִ����Ч
            If .TextMatrix(Row, COL_��ֹʱ��) <> "" Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") >= Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then
                    MsgBox "ҽ������ͣʱ��ӦС��ִ����ֹʱ�� " & Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
            'Ӧ>�ϴ���ͣ�������ʱ��(�����,����ʱ�䲻���ظ�,Ӧ>)
            vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 7)
            If vPause <> CDate(0) Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                    MsgBox "ҽ������ͣʱ��Ӧ�����ϴ���ͣ�������ʱ�� " & Format(vPause, "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
        ElseIf mint���� = 6 Then '�������ʱ��
            'Ӧ>��ͣʱ��
            vPause = GetPauseTime(Val(.TextMatrix(Row, COL_ID)), 6)
            If vPause <> CDate(0) Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") <= Format(vPause, "yyyy-MM-dd HH:mm") Then
                    MsgBox "ҽ��������ʱ��Ӧ�����ϴ���ͣʱ�� " & Format(vPause, "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
            
            'Ӧ<=ִ����ֹʱ��
            If .TextMatrix(Row, COL_��ֹʱ��) <> "" Then
                If Format(.EditText, "yyyy-MM-dd HH:mm") > Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then
                    MsgBox "ҽ��������ʱ��ӦС�ڵ���ִ����ֹʱ�� " & Format(.Cell(flexcpData, Row, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
                End If
            End If
        End If
        
            
        .TextMatrix(Row, Col) = IIF(.EditText = "" And strTmp <> "", strTmp, .EditText)
        .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
        
        Call vsAdvice_AfterEdit(Row, Col) 'һ����ҩ��һ������:��ʾ�󲻻��Զ�ִ�и��¼�
        
        '����Ϊ��ͬʱ��(У��,��ͣ,����)
        blnDoSame = InStr(",1,2,3,5,6,", "," & mint���� & ",") > 0
        If blnDoSame Then
            If Not VsfOnlySelOneRow(Row) Then
                Select Case mint����
                Case 1
                    strTmp = "ֹͣ"
                Case 2
                    strTmp = "ȷ��ֹͣ"
                Case 3
                    strTmp = "У��"
                Case 5
                    strTmp = "��ͣ"
                Case 6
                    strTmp = "����"
                End Select
                
                If MsgBox("Ҫ����������ѡ���ҽ���������ʱ��" & strTmp & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call SetSameTime(Row)
                End If
            End If
        End If
    End With
    AcceptInput = True
End Function

Private Function VsfOnlySelOneRow(lngRow As Long) As Boolean
'���ܣ��ж��Ƿ����һ�пɼ���(һ����ҩ��һ��)
    Dim i As Long, k As Long
    Dim lngBegin As Long, lngEnd As Long
    
    Call RowInһ����ҩ(lngRow, lngBegin, lngEnd)
    
    VsfOnlySelOneRow = True
    With vsAdvice
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) And i <> lngRow And (i < lngBegin Or i > lngEnd) Then
                If mint���� = 3 Then
                    If .Cell(flexcpData, i, COL_ѡ��) <> Empty Then
                        VsfOnlySelOneRow = False
                        Exit Function
                    End If
                Else
                    If Val(.TextMatrix(i, COL_ѡ��)) <> 0 Then
                        VsfOnlySelOneRow = False
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
End Function

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strInput As String
    Dim strMatch As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vPoint As PointAPI

    If KeyAscii = 13 Then
        If Col = COL_���� Then
            mblnReturn = True
        End If
        
        With vsAdvice
            If Col = COL_��ֹԭ�� And .EditText <> "" Then
                strInput = UCase(.EditText)
                
                If IsNumeric(strInput) Then
                    strMatch = "    A.���� Like [1]" '����ƥ����
                ElseIf zlCommFun.IsCharAlpha(strInput) Then
                    strMatch = "   a.���� Like [1]" '��ĸʱֻƥ�����
                ElseIf zlCommFun.IsCharChinese(strInput) Then
                    strMatch = "   a.���� Like [1]" '����
                Else
                    strMatch = "  (a.���� Like [1] or a.���� Like [1] or A.���� Like [1])"
                End If
                strSQL = "select a.���� as id, a.����,a.����,a.���� from ͣ��ԭ�� a where " & strMatch & " order by a.����"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ͣ��ԭ��", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        strInput & "%")
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COL_��ֹԭ��) = rsTmp!���� & ""
                    .Cell(flexcpData, Row, COL_��ֹԭ��) = rsTmp!���� & ""
                    .EditText = .TextMatrix(Row, Col)
                    Call SetSameԭ��(Row)
                Else
                    .TextMatrix(Row, COL_��ֹԭ��) = .EditText
                    .Cell(flexcpData, Row, COL_��ֹԭ��) = .EditText
                    .EditText = .TextMatrix(Row, Col)
                    Call SetSameԭ��(Row)
                End If
            End If
        End With
    Else
        If Col = COL_���� Then
            If InStr("0123456789-: " & Chr(8) & Chr(27) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        ElseIf Col = COL_��ֹԭ�� Then
            If KeyAscii = 39 Then KeyAscii = 0 '������
        End If
    End If
End Sub

Private Sub vsAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_���� Then
        vsAdvice.Refresh    '����е�����ʾ����ˢ�µĻ���һ����ҩͨ��Drawcell�������ĵ�Ԫ����ٴ���ʾ
        If Not AcceptInput(Row, Col) Then
            Cancel = True
        Else
            If mblnReturn Then
                Call vsAdvice_KeyPress(13) '��λ��һ�����뵥Ԫ
            End If
        End If
    End If
    mblnReturn = False
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ�һ����ҩ��һ������
    Dim lngBegin As Long, lngEnd As Long, i As Long
        
    With vsAdvice
        'һ����ҩ��һ��ѡ�������
        If (Col = COL_ѡ�� Or Col = COL_���� Or Col = COL_��ֹԭ��) And InStr(",5,6,", .TextMatrix(Row, COL_�������)) > 0 Then
            If RowInһ����ҩ(Row, lngBegin, lngEnd) Then
                For i = lngBegin To lngEnd
                    If i <> Row Then
                        .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                        .Cell(flexcpData, i, Col) = .Cell(flexcpData, Row, Col)
                        Set .Cell(flexcpPicture, i, Col) = .Cell(flexcpPicture, Row, Col)
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlCommFun.ActualLen(vsAdvice.EditText)
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mblnReturn = False
    If Col <> COL_ѡ�� And Col <> COL_���� And Col <> COL_��ֹԭ�� Then
        Cancel = True
    ElseIf Val(vsAdvice.TextMatrix(Row, COL_ID)) = 0 Then
        Cancel = True
    ElseIf mint���� = 3 Then
        If Col = COL_���� And Not (vsAdvice.TextMatrix(Row, COL_��־) = "��¼" Or InStr(GetInsidePrivs(pסԺҽ������), "�޸�У��ʱ��") > 0) Then
            Cancel = True 'У��ҽ��ʱ,�ǲ�¼��У��ʱ�䲻�ɸ���
        ElseIf Col = COL_ѡ�� Then
            Cancel = True '����ֱ�ӱ༭
        End If
    End If
End Sub

Private Sub InitAdviceTable()
'���ܣ���ʼ��ҽ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "ID;���ID;��ID;���;�������;�������;��ҩ;,240,4;,300,4;,1530,1;��ֹԭ��,1530,1;" & _
        "��־,500,4;����,750,1;סԺ��,750,1;����,500,1;Ӥ��,500,1;��Ч,500,4;����ʱ��;��Чʱ��,1530,1;" & _
        "ҽ������,3000,1;,375,1;����,850,1;����,850,1;Ƶ��,1000,1;�÷�,1000,1;ҽ������,1000,1;ִ��ʱ��,1000,1;" & _
        "��ֹʱ��,1530,1;ִ�п���,850,1;ִ������,850,1;�ϴ�ִ��,1530,1;" & _
        "����ҽ��,850,1;У�Ի�ʿ,850,1;У��ʱ��,1530,1;ͣ��ҽ��,850,1;ͣ��ʱ��,1530,1;" & _
        "����ID;��ҳID;������ĿID;Ƶ�ʴ���;Ƶ�ʼ��;�����λ;ִ�б��;��������;�Թܱ���;" & _
        "ִ�п���ID;���˿���ID;�շ�ϸĿID;������λ;ǰ��ID;ǩ��ID;������Ա;��������ID;����˵��;ִ�з���;�걾��λ;�������;���״̬;��Ժ����ID;��������"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .ColHidden(COL_��ʾ) = True 'Pass
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitPriceTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "ҽ��ID;���ID;�������;������ĿID;�շ�ϸĿID;�̶�;" & _
        "�Ƽ�ҽ��,2000,1;���,650,1;�շ���Ŀ,2500,1;��λ,500,4;�Ƽ�����,850,1;����,850,7;" & _
        "ִ�п���,1000,1;��������,850,1;����,450,4;�շѷ�ʽ,1500,1;�շ����;ִ�п���ID;��������;��������"
    arrHead = Split(strHead, ";")
    With vsPrice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub SetAdviceCol()
'���ܣ�����һЩ�ɼ��м��༭����,Ӧ�ڱ������װ������
    With vsAdvice
        .TextMatrix(0, COL_ѡ��) = "ѡ"
        .Editable = flexEDKbdMouse
        
        .ColHidden(COL_��ֹԭ��) = True
        
        '������������еĿɼ���
        If mint���� = 0 Then
            'ҽ������
            .ColHidden(COL_����) = True
            .ColHidden(COL_�ϴ�ִ��) = True
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .ColDataType(COL_ѡ��) = flexDTBoolean
        ElseIf (mint���� = 1 Or mint���� = 7) Then
            'ֹͣҽ��
            .TextMatrix(0, COL_����) = "��ֹʱ��"
            .ColHidden(COL_��ֹʱ��) = True
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            If gblnҽ����ֹԭ�� Then
                .ColHidden(COL_��ֹԭ��) = False
            End If
            .ColDataType(COL_ѡ��) = flexDTBoolean
        ElseIf mint���� = 2 Then
            'ȷ��ֹͣ
            .TextMatrix(0, COL_����) = "ȷ��ʱ��"
            .ColDataType(COL_ѡ��) = flexDTBoolean
        ElseIf mint���� = 3 Then
            'ҽ��У��
            .TextMatrix(0, COL_����) = "У��ʱ��"
            .ColHidden(COL_�ϴ�ִ��) = True
            .ColHidden(COL_У�Ի�ʿ) = True
            .ColHidden(COL_У��ʱ��) = True
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .Cell(flexcpPictureAlignment, .FixedRows, COL_ѡ��, .Rows - 1, COL_ѡ��) = 4
            .Cell(flexcpForeColor, .FixedRows, COL_��ʼʱ��, .Rows - 1, COL_��ʼʱ��) = vbBlue          '��ɫ
            
        ElseIf mint���� = 4 Then
            '�����Ƽ���Ŀ
            .ColHidden(COL_����) = True
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .ColDataType(COL_ѡ��) = flexDTBoolean
        ElseIf mint���� = 5 Then
            '��ͣҽ��
            .TextMatrix(0, COL_����) = "��ͣʱ��"
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .ColDataType(COL_ѡ��) = flexDTBoolean
        ElseIf mint���� = 6 Then
            '����ҽ��
            .TextMatrix(0, COL_����) = "����ʱ��"
            .ColHidden(COL_ͣ��ҽ��) = True
            .ColHidden(COL_ͣ��ʱ��) = True
            .ColDataType(COL_ѡ��) = flexDTBoolean
        End If
        
        '���ö�����
        If Not .ColHidden(COL_����) Then
            If .TextMatrix(0, COL_����) = "��ֹʱ��" Then
                .FrozenCols = COL_��ֹԭ�� + 1 - .FixedCols
            Else
                .FrozenCols = COL_���� + 1 - .FixedCols
            End If
            .SheetBorder = vbBlack
        ElseIf Not .ColHidden(COL_ѡ��) Then
            .FrozenCols = COL_ѡ�� + 1 - .FixedCols
            .SheetBorder = vbBlack
        End If
        
        '�������б�ʶ
        .Cell(flexcpBackColor, .FixedRows, COL_ѡ��, .Rows - 1, COL_��ֹԭ��) = COLEditBackColor       'ǳ��
    End With
End Sub

Private Function GetWhere() As String
'���ܣ����ݴ��幦�ܲ���ҽ��������
'˵��������"����ҽ����¼"����Ϊ"A"
    Dim strSQL As String
    
    If mint���� = 0 Then
        'ҽ������:��У��,��δ���͹�����������������ͣ�ĳ���Ҳ����ֱ�����ϡ�
        '��ʱ����ҽ��У�Ժ��Զ�ֹͣ������Ҳ��������
        strSQL = " And (A.ҽ��״̬ Not IN(1,2,4,8,9) And A.�ϴ�ִ��ʱ�� is NULL Or " & IIF(mbln��������ִ��, "A.������ĿID is Null And A.ҽ��״̬<>4", "A.ҽ����Ч=1 And A.������ĿID is Null And A.ҽ��״̬=8") & ")"
    ElseIf (mint���� = 1 Or mint���� = 7) Then
        'ֹͣҽ��:����,����ͣ��Ҳ����ֱ��ֹͣ,����ҩ�䷽����
        strSQL = " And A.ҽ��״̬ Not IN(1,2,4,8,9) And Nvl(A.ҽ����Ч,0)=0"
    ElseIf mint���� = 2 Then
        'ȷ��ֹͣ:ֹͣ״̬�ĳ���
        strSQL = " And A.ҽ��״̬=8 And Nvl(A.ҽ����Ч,0)=0"
    ElseIf mint���� = 3 Then
        'ҽ��У��:���¿��ģ�����ҽ�������ʸ�Ļ�����˵�ҽ������У�ԡ�
        strSQL = " And A.ҽ��״̬=1 And Exists(" & _
            "Select M.���� From ��Ա�� M,ִҵ��� N" & _
            " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
            " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')" & _
            " )"
    ElseIf mint���� = 4 Then
        '�����Ƽ���Ŀ
        strSQL = " And A.ҽ��״̬ Not IN(1,2,4,8,9)"
    ElseIf mint���� = 5 Then
        '��ͣҽ��:����,����ҩ�䷽����
        strSQL = " And A.ҽ��״̬ IN(3,5,7) And Nvl(A.ҽ����Ч,0)=0"
    ElseIf mint���� = 6 Then
        '����ҽ��
        strSQL = " And A.ҽ��״̬=6"
    End If
    GetWhere = strSQL
End Function

Private Function LoadAdvice(strPatis As String) As Boolean
'���ܣ����ݵ�ǰ�������ö�ȡ����ʾҽ���嵥
'������str����IDs=���ڷ���ʵ�������ݵĲ��˴�:"����ID,��ҳID,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsPause As New ADODB.Recordset
    Dim str��ҩ As String, str��ҩ As String
    Dim strSQL As String, strWhere As String
    Dim bln��ҩ;�� As Boolean, bln��Ѫ;�� As Boolean
    Dim lng����ID As Long, lng��ҳID As Long
    Dim i As Long, j As Long, k As Long
    Dim strӤ�� As String, str����s As String
    Dim vCurDate As Date, strTmp As String, strDepts As String
    Dim intͼ���� As Integer
    Dim str����ҽ��IDs As String
    Dim bln���� As Boolean
    
    Screen.MousePointer = 11
    Me.Refresh
    On Error GoTo errH
        
    '----------------------------------------------------------------------
    strPatis = ""
    With vsAdvice
        .Rows = .FixedRows
        .ColHidden(COL_����) = True
        .ColHidden(COL_סԺ��) = True
        .ColHidden(COL_����) = True
        .ColHidden(COL_Ӥ��) = True
    End With
    
    '----------------------------------------------------------------------
    strDepts = GetUser����IDs(True)
    strWhere = GetWhere
    strWhere = strWhere & IIF(Not mbln��ʿվ Or Not mblnҽ������, " And A.ǰ��ID is NULL", "")
    
    If DeptIsWoman(0, Get����IDs(mlng����ID)) Then
        'ҽ������Χ
        If mblnFirstLoad Then
            fraBaby.Visible = True
            mintҽ������Χ = Val(zlDatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0"))
            optBaby(mintҽ������Χ).value = True
            mblnFirstLoad = False
        End If
    Else
        mblnFirstLoad = True
        fraBaby.Visible = False
        optBaby(e����).value = True
    End If
    
    'У�Ե�ҽ����Χ����
    If mint���� = 3 Then
        If InStr(GetInsidePrivs(pסԺҽ������), "ȫԺҽ��У��") = 0 Then
            If gbln��������´�ҽ������ Then
                strWhere = strWhere & " And (A.��������ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(f_Num2list([4])) X) And nvl(a.����ҽ��id,0)=0 Or instr(','||[7]||',',','||nvl(a.����ҽ��id,0)||',')>0)"
                bln���� = True
            Else
                strWhere = strWhere & " And A.��������ID In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(f_Num2list([4])) X)"
            End If
        End If
        If mblnAutoRead Then
            '��Ϊ�����ƣ�����δУ�ԵĶ���ȡ����Ϊֻ������ͬʱУ������ҽ���������У������ҽ��ʱ�������Զ�ֹͣ��Щҽ������У�ԣ�
        End If
        If gblnKSSStrict Or gbln�����ּ����� Or gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ Then
            strWhere = strWhere & " And (Nvl(A.���״̬,0) Not in(1,3,7" & IIF(gblnѪ��ϵͳ = True, "", ",4,5") & ") or a.ҽ����Ч=0 and a.���״̬=1 and a.������־=1 and (instr(',5,6,',A.�������)>0 or A.�������='E' and B.��������='2'))"
        End If
    End If
    
    If mint���� <> 2 Then
        '��������ʱ���õ�����
        If mint��Ч <> 0 Then
            strWhere = strWhere & " And Nvl(A.ҽ����Ч,0)=" & mint��Ч - 1
        End If
        If mint��� <> 0 Then
            If mint��� = 1 Then
                'ҩƷ��
                strWhere = strWhere & _
                    " And (A.������� IN('5','6','7')" & _
                    " Or (A.�������='E' And A.���ID is Not NULL)" & _
                    " Or Exists(Select ID From ����ҽ����¼ S Where ������� IN('5','6','7') And S.���ID=A.ID And ����ID=[1])" & _
                    " )"
            ElseIf mint��� = 2 Then
                '������
                strWhere = strWhere & _
                    " And Not A.������� IN('5','6','7')" & _
                    " And Not(A.�������='E' And A.���ID is Not NULL)" & _
                    " And Not Exists(Select ID From ����ҽ����¼ S Where ������� IN('5','6','7') And S.���ID=A.ID And ����ID=[1])"
            End If
        End If
    End If
    
    '�ٴ�·����ҽ��
    If mbytUseType = 1 And (mint���� = 1 Or mint���� = 7) Then
        strWhere = strWhere & " And A.ID IN(Select Column_Value From Table(f_Num2list([3])))" & _
            " And Not(Nvl(A.�������,'ZY')='H' And b.��������='1' And b.ִ��Ƶ��=2)" & _
            " And Not(Nvl(A.�������,'ZY')='Z' And b.�������� IN('4','14', '9', '10', '12'))"
        vCurDate = mdateStop
    Else
        If (mint���� = 1 Or mint���� = 7) Then
            strWhere = strWhere & "  And Not(Nvl(A.�������,'ZY')='H' And b.��������='1' And b.ִ��Ƶ��=2)"
            If mlngӤ�� <> 0 Then   'Ӥ����ת��ҽ��ʱ,ֹֻͣ��Ӥ����.��ĸ�׵�ת��ҽ��ʱ��Ӥ������û�ж������ҲӦһ������
                strWhere = strWhere & "  And A.Ӥ�� = " & mlngӤ��
            Else
                If mbln���� Then
                    strWhere = strWhere & "  And NVL(A.Ӥ��,0) = 0 "
                End If
            End If
            strWhere = strWhere & IIF(mint���� = 7, " And A.��˱��=2 ", IIF(Not mblnHaveAudit, " And NVL(A.��˱��,0)<>2 ", ""))
        End If
        If (mint���� = 1 Or mint���� = 7) And mdateStop <> CDate(0) Then
            vCurDate = mdateStop
        Else
            vCurDate = zlDatabase.Currentdate
        End If
    End If
    '����ҽ����������ͣ
    If mint���� = 5 Then
        strWhere = strWhere & " And NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ' And NVL(a.ִ��Ƶ��,'��')<>'��Ҫʱ' "
    End If
    
    mblnAll = (mintҽ������Χ = 0 And mint��Ч = 0 And mint��� = 0)
    
    '----------------------------------------------------------------------
    For k = 0 To UBound(Split(mstr����IDs, ";"))
        lng����ID = Split(Split(mstr����IDs, ";")(k), ",")(0)
        lng��ҳID = Split(Split(mstr����IDs, ";")(k), ",")(1)
        If bln���� Then str����ҽ��IDs = Get����ҽ��IDs(lng����ID, lng��ҳID, strDepts)
        'ҽ����¼��������������,��������,��鲿λ,��ҩ�巨
        '����¼���ҽ�����������ϣ�ֹͣ��ȷ��ֹͣ��У��
        strSQL = _
            "Select /*+ RULE */ A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,A.���,Nvl(A.�������,'*') as �������,C.�������,NULL as ��ҩ," & _
                " A.�����,NULL as ѡ��,NULL as ����,null as ԭ��,Decode(A.������־,1,'����',2,'��¼','��ͨ') as ��־,A.����,P.סԺ��,P.��ǰ���� as ����," & _
                " Decode(Nvl(A.Ӥ��,0),0,'����','Ӥ��'||A.Ӥ��) as Ӥ��,Decode(Nvl(A.ҽ����Ч,0),0,'����','����') as ��Ч," & _
                " To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��,To_Char(A.��ʼִ��ʱ��,'YYYY-MM-DD HH24:MI') as ��ʼʱ��,A.ҽ������,A.Ƥ�Խ�� as Ƥ��," & _
                " Decode(A.�ܸ�����,NULL,NULL,Decode(A.�������,'E',Decode(B.��������,'4',A.�ܸ�����||'��',A.�ܸ�����||B.���㵥λ),'4',A.�ܸ�����||F.���㵥λ,'5',Round(A.�ܸ�����/D.סԺ��װ,5)||D.סԺ��λ,'6',Round(A.�ܸ�����/D.סԺ��װ,5)||D.סԺ��λ,A.�ܸ�����||B.���㵥λ)) as ����," & _
                " Decode(A.��������,NULL,NULL,A.��������||Decode(A.�������,'4',F.���㵥λ,B.���㵥λ)) as ����,A.ִ��Ƶ�� as Ƶ��," & _
                " Decode(A.�������,'E',Decode(Instr('2468',Nvl(B.��������,'0')),0,NULL,B.����),NULL) as �÷�,A.ҽ������," & _
                " A.ִ��ʱ�䷽�� as ִ��ʱ��,To_Char(A.ִ����ֹʱ��,'YYYY-MM-DD HH24:MI') as ��ֹʱ��," & _
                " Nvl(E.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'-')) as ִ�п���," & _
                " Decode(Instr('567E',Nvl(A.�������,'*')),0,NULL,A.ִ������) as ִ������, To_Char(A.�ϴ�ִ��ʱ��,'YYYY-MM-DD HH24:MI') as �ϴ�ִ��," & _
                " A.����ҽ��,A.У�Ի�ʿ,To_Char(A.У��ʱ��,'YYYY-MM-DD HH24:MI') as У��ʱ��," & _
                " A.ͣ��ҽ��,To_Char(A.ͣ��ʱ��,'YYYY-MM-DD HH24:MI') as ͣ��ʱ��,A.����ID,A.��ҳID,A.������ĿID,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ִ�б��," & _
                " B.��������,B.�Թܱ���,A.ִ�п���ID,A.���˿���ID,A.�շ�ϸĿID,B.���㵥λ as ������λ,A.ǰ��ID,A.�¿�ǩ��ID as ǩ��ID,A.����ҽ��,A.��������ID," & IIF(mint���� = 7, "S.����˵��,", "Null as ����˵��,") & _
                " b.ִ�з���,A.�걾��λ,a.�������,a.���״̬,g.��Ժ����ID,g.��������," & IIF(mint���� = 3, "Decode(a.ҽ��״̬,1,Decode(a.У�Ի�ʿ,Null,0,1),0) as ���ʸ���", "Null as ���ʸ���") & ",D.��ΣҩƷ,d.�Ƿ���������"
        strSQL = strSQL & _
            IIF(mint���� = 0, ",J.����� as ���������", ",NULL as ���������") & _
            " From ����ҽ����¼ A,������Ϣ P,������ҳ G,���ű� E,ҩƷ���� C,ҩƷ��� D,������ĿĿ¼ B,�շ���ĿĿ¼ F" & IIF(mint���� = 7, ",����ҽ��״̬ S", "") & _
            IIF(mint���� = 0, ",���������ϸ I,��������¼ J", "") & _
            " Where A.����ID=P.����ID And A.������ĿID=B.ID" & IIF(InStr(",0,1,2,3,", mint����) > 0, "(+)", "") & _
                " And A.ִ�п���ID=E.ID(+) And A.������ĿID=C.ҩ��ID(+) And p.����ID=G.����ID And P.��ҳID=G.��ҳID " & _
                " And A.�շ�ϸĿID=D.ҩƷID(+) And A.�շ�ϸĿID=F.ID(+)" & _
                " And (Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL) Or A.�������='E' And B.��������='8')" & _
                IIF(mint���� = 7, " And A.ID=S.ҽ��ID And S.��������=13", "") & " And A.����ID=[1] And A.��ҳID=[2]" & _
                IIF(mint���� = 0, " And a.ID = i.ҽ��ID(+) And I.��ID = J.ID(+) and (I.����ύ =1 Or I.��ID is NULL)", "") & _
                " And A.��ʼִ��ʱ�� is Not NULL And Nvl(A.ҽ��״̬,0)<>-1" & _
                Decode(mintҽ������Χ, 1, " And nvl(a.Ӥ��,0) = 0 ", 2, " And nvl(a.Ӥ��,0) <> 0 ", "") & _
                " And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3" & strWhere & " And (G.Ӥ������ID is null or G.Ӥ������ID is not null and (G.Ӥ������ID=[6] or G.Ӥ������ID=[6]) and NVL(A.Ӥ��,0)<>0 or G.Ӥ������ID is not null and (G.Ӥ������ID<>[6] and G.Ӥ������ID<>[6]) and NVL(A.Ӥ��,0)=0)" & _
            " Order by Nvl(A.Ӥ��,0),��ID,A.���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, mstrAdviceOfItem, strDepts, mlng����ID, mlngҽ������ID, str����ҽ��IDs)
        
        If Not rsTmp.EOF Then
            strPatis = strPatis & ";" & lng����ID & "," & lng��ҳID
            If InStr(str����s & ",", "," & rsTmp!���˿���id & ",") = 0 Then
                str����s = str����s & "," & rsTmp!���˿���id
            End If
            
            '��ͣҽ��ʱ��ȡҽ�����ϴ�����ʱ��(��һ����)
            '����ҽ��ʱ��ȡҽ������ͣʱ��
            If mint���� = 5 Or mint���� = 6 Then
                strSQL = "Select B.ҽ��ID,Max(B.����ʱ��) as �ϴ�ʱ��" & _
                    " From ����ҽ����¼ A,����ҽ��״̬ B" & _
                    " Where A.ID=B.ҽ��ID And B.��������=" & IIF(mint���� = 5, 7, 6) & _
                    " And Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL)" & _
                    " And A.����ID=[1] And A.��ҳID=[2] And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3" & strWhere & _
                    " Group by B.ҽ��ID"
                Set rsPause = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
            End If
            
            With vsAdvice
                .Redraw = flexRDNone
                Do While Not rsTmp.EOF
                    '�������
                    intͼ���� = 0
                    strTmp = ""
                    For i = 0 To rsTmp.Fields.Count - 2
                        strTmp = strTmp & vbTab & Replace(NVL(rsTmp.Fields(i).value), vbTab, "")
                    Next
                    .AddItem Mid(strTmp, 2): i = .Rows - 1
                    
                    If mdateStop <> CDate(0) And (mint���� = 1 Or mint���� = 7) Then
                        .TextMatrix(i, COL_ѡ��) = "-1"
                    End If
                                        
                    '�Ƿ���ʾӤ����
                    If InStr(strӤ�� & ",", "," & .TextMatrix(i, COL_Ӥ��) & ",") = 0 Then
                        If strӤ�� <> "" Then .ColHidden(COL_Ӥ��) = False
                        strӤ�� = strӤ�� & "," & .TextMatrix(i, COL_Ӥ��)
                    End If
                    
                    '����֮��ļ����
                    If .TextMatrix(i, COL_סԺ��) <> .TextMatrix(i - 1, COL_סԺ��) And i - 1 >= .FixedRows Then
                        .CellBorderRange i - 1, .FixedCols, i - 1, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                    End If
                    
                    '��ҩ����ҩ��һЩ����
                    bln��ҩ;�� = False: bln��Ѫ;�� = False
                    If .TextMatrix(i, COL_�������) = "E" Then
                        If Val(.TextMatrix(i - 1, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                            If InStr(",5,6,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                                bln��ҩ;�� = True
                                For j = i - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                        '��ʾ��ҩ�ĸ�ҩ;��
                                        .TextMatrix(j, COL_�÷�) = .TextMatrix(i, COL_�÷�)
                                        .TextMatrix(j, COL_��������) = .TextMatrix(i, COL_��������)
                                        .TextMatrix(j, COL_ִ�з���) = .TextMatrix(i, COL_ִ�з���)
                                        '��ʾ��ҩ��ִ������
                                        If Val(.TextMatrix(j, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                            If Val(.TextMatrix(j, COL_ִ�б��)) = 2 Then
                                                .TextMatrix(j, COL_ִ������) = "��ȡҩ"
                                            Else
                                                .TextMatrix(j, COL_ִ������) = "�Ա�ҩ"
                                            End If
                                        ElseIf Val(.TextMatrix(j, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                            .TextMatrix(j, COL_ִ������) = "��Ժ��ҩ"
                                        Else
                                            .TextMatrix(j, COL_ִ������) = IIF(Val(.TextMatrix(j, COL_ִ�б��)) = 1, "��ȡҩ", "")
                                        End If
                                        
                                        '����
                                        .TextMatrix(j, COL_Ƥ��) = .TextMatrix(i, COL_ҽ������)
                                    Else
                                        Exit For
                                    End If
                                Next
                            ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                                If .TextMatrix(i - 1, COL_�������) = "7" Then
                                    .TextMatrix(i, COL_����) = "1" '��ҩ�䷽
                                ElseIf .TextMatrix(i - 1, COL_�������) = "C" Then
                                    .TextMatrix(i, COL_����) = "2" '�������
                                    
                                    '�ɼ���ʽ�Ĺ�����һ���ĵ�һ��������ͬ
                                    j = .FindRow(.TextMatrix(i, COL_ID), .FixedRows, COL_���ID)
                                    If j <> -1 Then
                                        .TextMatrix(i, COL_�Թܱ���) = .TextMatrix(j, COL_�Թܱ���)
                                        .TextMatrix(i, COL_��ʼʱ��) = .TextMatrix(j, COL_��ʼʱ��)
                                        .Cell(flexcpData, i, COL_��ʼʱ��) = CStr(.TextMatrix(j, COL_��ʼʱ��))
                                    End If
                                End If
                                
                                '��ʾ��ҩ�䷽�������ϵ�ִ�п���
                                .TextMatrix(i, COL_ִ�п���) = .TextMatrix(i - 1, COL_ִ�п���)
                                
                                If .TextMatrix(i - 1, COL_�������) = "7" Then
                                    '��ʾ��ҩ�䷽ִ������
                                    If Val(.TextMatrix(i - 1, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                        If Val(.TextMatrix(i - 1, COL_ִ�б��)) = 2 Then
                                            .TextMatrix(i, COL_ִ������) = "��ȡҩ"
                                        Else
                                            .TextMatrix(i, COL_ִ������) = "�Ա�ҩ"
                                        End If
                                    ElseIf Val(.TextMatrix(i - 1, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                        .TextMatrix(i, COL_ִ������) = "��Ժ��ҩ"
                                    Else
                                        .TextMatrix(i, COL_ִ������) = IIF(Val(.TextMatrix(i - 1, COL_ִ�б��)) = 1, "��ȡҩ", "")
                                    End If
                                Else
                                    .TextMatrix(i, COL_ִ������) = ""
                                End If
                                
                                'ɾ����ζ��ҩ��,�Լ���������еļ�����Ŀ;ͬʱ�жϼ�������
                                For j = i - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                        .RemoveItem j: i = .Rows - 1
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        ElseIf .TextMatrix(i - 1, COL_�������) = "K" And Val(.TextMatrix(i - 1, COL_ID)) = Val(.TextMatrix(i, COL_���ID)) Then
                            bln��Ѫ;�� = True
                            '��ʾ��Ѫ;��
                            .TextMatrix(i - 1, COL_�÷�) = .TextMatrix(i, COL_�÷�)
                        Else
                            .TextMatrix(i, COL_ִ������) = ""
                        End If
                    End If
                                                                    
                    '����ɼ��еĵ�һЩ��ʶ
                    If Not (bln��ҩ;�� Or bln��Ѫ;��) And .TextMatrix(i, COL_�������) <> "7" Then
                        '����С��������,��δ�뵽�취
                        If Left(.TextMatrix(i, COL_����), 1) = "." Then
                            .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                        End If
                        If Left(.TextMatrix(i, COL_����), 1) = "." Then
                            .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                        End If
                    
                        'ʱ����MM-DD HH:MI��ʽ��ʾ,��CellData�����ж�
                        .Cell(flexcpData, i, COL_��ʼʱ��) = .TextMatrix(i, COL_��ʼʱ��)
                        .Cell(flexcpData, i, COL_����ʱ��) = .TextMatrix(i, COL_����ʱ��)
                        .Cell(flexcpData, i, COL_�ϴ�ִ��) = .TextMatrix(i, COL_�ϴ�ִ��)
                        .Cell(flexcpData, i, COL_��ֹʱ��) = .TextMatrix(i, COL_��ֹʱ��)
                        .Cell(flexcpData, i, COL_У��ʱ��) = .TextMatrix(i, COL_У��ʱ��)
                        .Cell(flexcpData, i, COL_ͣ��ʱ��) = .TextMatrix(i, COL_ͣ��ʱ��)
                        .TextMatrix(i, COL_��ʼʱ��) = Format(.TextMatrix(i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_����ʱ��) = Format(.TextMatrix(i, COL_����ʱ��), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_�ϴ�ִ��) = Format(.TextMatrix(i, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_��ֹʱ��) = Format(.TextMatrix(i, COL_��ֹʱ��), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_У��ʱ��) = Format(.TextMatrix(i, COL_У��ʱ��), "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, COL_ͣ��ʱ��) = Format(.TextMatrix(i, COL_ͣ��ʱ��), "yyyy-MM-dd HH:mm")
                        
                        If (mint���� = 1 Or mint���� = 7) Then
                            'ͣ��ʱȱʡ��ҽ����ֹʱ��
                            If mdateStop <> CDate(0) And (.TextMatrix(i, COL_�������) = "H" And .TextMatrix(i, COL_��������) = "1" Or .TextMatrix(i, COL_�������) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_��������) & ",") > 0) Then
                                .TextMatrix(i, COL_����) = Format(vCurDate - 1 / 24 / 60, "yyyy-MM-dd HH:mm")
                            ElseIf optStop(e�ϴ�ִ��ʱ��).value Then
                                'ĩ��ʱ�����Ϊ�գ���Ϊ��ʼʱ��
                                If .TextMatrix(i, COL_Ƶ��) <> "������" Then
                                    If .TextMatrix(i, COL_�ϴ�ִ��) = "" Then
                                        .TextMatrix(i, COL_����) = CStr(.Cell(flexcpData, i, COL_��ʼʱ��))
                                    Else
                                        .TextMatrix(i, COL_����) = CStr(.Cell(flexcpData, i, COL_�ϴ�ִ��))
                                    End If
                                Else
                                    '��ǰʱ���ָ��ʱ��
                                    .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                                End If
                            Else
                                '��ǰʱ���ָ��ʱ��
                                .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                                                                
                                If .TextMatrix(i, COL_Ƶ��) <> "������" Then
                                    If chkNoSend.value = 0 Then      '���������
                                        If .TextMatrix(i, COL_�ϴ�ִ��) = "" Then
                                            .TextMatrix(i, COL_����) = .TextMatrix(i, COL_��ʼʱ��)
                                        ElseIf .TextMatrix(i, COL_�ϴ�ִ��) < .TextMatrix(i, COL_����) Then
                                            .TextMatrix(i, COL_����) = .TextMatrix(i, COL_�ϴ�ִ��)
                                        End If
                                    End If
                                    If chkRollSend.value = 0 Then   '������ջ�
                                        If .TextMatrix(i, COL_�ϴ�ִ��) > .TextMatrix(i, COL_����) Then
                                            .TextMatrix(i, COL_����) = .TextMatrix(i, COL_�ϴ�ִ��)
                                        End If
                                    End If
                                End If
                            End If
                            If mint���� = 7 Then
                                '��˵�ʱ��Ĭ��ʱ��Ϊ����ֹͣ��ʱ�䣨����ҽ��״̬.����˵����
                                If rsTmp!����˵�� & "" <> "" Then
                                    strTmp = rsTmp!����˵�� & "<T>"
                                    .TextMatrix(i, COL_����) = Format(Split(strTmp, "<T>")(0), "yyyy-MM-dd HH:mm")
                                    .TextMatrix(i, COL_��ֹԭ��) = Split(strTmp, "<T>")(1)
                                    .Cell(flexcpData, i, COL_��ֹԭ��) = Split(strTmp, "<T>")(1)
                                End If
                            Else
                                'ͣ��ԭ��
                                .TextMatrix(i, COL_��ֹԭ��) = mstrͣ��ԭ��
                                .Cell(flexcpData, i, COL_��ֹԭ��) = mstrͣ��ԭ��
                            End If
                            
                            .TextMatrix(i, COL_����) = GetValidateStopDate(.TextMatrix(i, COL_����), i)

                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
                        ElseIf mint���� = 2 Then
                            .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            'Ӧ>=��ֹʱ��
                            If .TextMatrix(i, COL_����) < .Cell(flexcpData, i, COL_��ֹʱ��) Then
                                .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_��ֹʱ��))), "yyyy-MM-dd HH:mm")
                            End If
                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
                        ElseIf mint���� = 3 Then
                            'У��ʱ��ȱʡУ��ʱ��
                             If optOper(e��ǰʱ��).value Then
                                If .TextMatrix(i, COL_��־) = "��¼" Then
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_����ʱ��))), "yyyy-MM-dd HH:mm")
                                Else
                                    .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                                End If
                            Else
                                If Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") > Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then
                                    .TextMatrix(i, COL_����) = .Cell(flexcpData, i, IIF(cboTime(e����).ListIndex = 0, COL_��ʼʱ��, COL_����ʱ��))
                                Else
                                    .TextMatrix(i, COL_����) = .Cell(flexcpData, i, IIF(cboTime(e����).ListIndex = 0, COL_��ʼʱ��, COL_����ʱ��))
                                End If
                            End If
                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
                                                                                
                            '�����޸ĺ�
                            If Val("" & rsTmp!���ʸ���) = 1 Then
                                Set .Cell(flexcpPicture, i, col_ҽ������) = frmIcons.imgFlag.ListImages("M").Picture
                                intͼ���� = 1
                            End If
                            
                            'Pass:�����������ʾ��ʾ��
                            'ֻ�л�ʿУ�Ե�ʱ������ӳ���������ʾ��ʾ��,������ֹͣ��ȷ��ֹͣ...��δ����ӳ����������þ�ʾֵ,
                            If mblnPass Then
                                If .TextMatrix(i, COL_��ʾ) <> "" Then
                                    Call gobjPass.zlPassSetWarnLight(mobjPassMap, i, Val(.TextMatrix(i, COL_��ʾ))) '���ڵ�ҩ����
                                    .TextMatrix(i, COL_��ʾ) = ""
                                End If
                            End If
                            
                        ElseIf mint���� = 5 Then
                            If mblnPauseLast Then
                                If .TextMatrix(i, COL_�ϴ�ִ��) <> "" Then
                                    'ȱʡ���ϴ�ִ��ʱ��֮����ͣ
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��))), "yyyy-MM-dd HH:mm")
                                Else
                                    '�����ϴ�ִ��ʱ�����Կ�ʼʱ��Ϊ׼
                                    .TextMatrix(i, COL_����) = .Cell(flexcpData, i, COL_��ʼʱ��)
                                End If
                            Else
                                '��ͣҽ��ʱ��:��ͣ����,ҽ����ͣ����Ч,���õ���Ч��
                                .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            End If
                            
                            'Ӧ>=��ʼִ��ʱ��,��Ϊ��ʱ�����δִ��
                            If .TextMatrix(i, COL_����) < .Cell(flexcpData, i, COL_��ʼʱ��) Then
                                .TextMatrix(i, COL_����) = .Cell(flexcpData, i, COL_��ʼʱ��)
                            End If
                            'Ӧ>�ϴ�ִ��ʱ��,��Ϊ��ʱ�����ִ��
                            If .TextMatrix(i, COL_�ϴ�ִ��) <> "" Then
                                If .TextMatrix(i, COL_����) <= .Cell(flexcpData, i, COL_�ϴ�ִ��) Then
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��))), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            'Ӧ<ִ����ֹʱ��,��Ϊ��ʱ���ִ����Ч
                            If .TextMatrix(i, COL_��ֹʱ��) <> "" Then
                                If .TextMatrix(i, COL_����) >= .Cell(flexcpData, i, COL_��ֹʱ��) Then
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", -1, CDate(.Cell(flexcpData, i, COL_��ֹʱ��))), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            'Ӧ>�ϴ���ͣ�������ʱ��(�����,����ʱ�䲻���ظ�,Ӧ>)
                            rsPause.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                            If Not rsPause.EOF Then
                                If .TextMatrix(i, COL_����) <= Format(rsPause!�ϴ�ʱ��, "yyyy-MM-dd HH:mm") Then
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, rsPause!�ϴ�ʱ��), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
                        ElseIf mint���� = 6 Then
                            '����ҽ��ʱ��
                            .TextMatrix(i, COL_����) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                            
                            'Ӧ>��ͣʱ��
                            rsPause.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                            If Not rsPause.EOF Then
                                If .TextMatrix(i, COL_����) <= Format(rsPause!�ϴ�ʱ��, "yyyy-MM-dd HH:mm") Then
                                    .TextMatrix(i, COL_����) = Format(DateAdd("n", 1, rsPause!�ϴ�ʱ��), "yyyy-MM-dd HH:mm")
                                End If
                            End If
                            
                            'Ӧ<=ִ����ֹʱ��
                            If .TextMatrix(i, COL_��ֹʱ��) <> "" Then
                                If .TextMatrix(i, COL_����) > .Cell(flexcpData, i, COL_��ֹʱ��) Then
                                    .TextMatrix(i, COL_����) = .Cell(flexcpData, i, COL_��ֹʱ��)
                                End If
                            End If
                            
                            .Cell(flexcpData, i, COL_����) = .TextMatrix(i, COL_����) '��������ָ�
                        End If
                        
                        '�и�
                        If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                        
                        '���龫ҩƷ��ʶ
                        If .TextMatrix(i, COL_�������) <> "" Then
                            If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", .TextMatrix(i, COL_�������)) > 0 Then
                                .Cell(flexcpFontBold, i, col_ҽ������) = True
                            End If
                        End If
                        
                        'Ƥ�Խ����ʶ
                        If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "1" And .TextMatrix(i, COL_Ƥ��) <> "" Then
                            j = GetSkinTestResult(Val(.TextMatrix(i, COL_������ĿID)), .TextMatrix(i, COL_Ƥ��))
                            .Cell(flexcpForeColor, i, COL_Ƥ��) = Decode(j, 1, vbRed, -1, vbBlue, .Cell(flexcpForeColor, i, COL_Ƥ��))
                        End If
                        
                        '����ǩ����ʶ
                        If Val(.TextMatrix(i, COL_ǩ��ID)) <> 0 Then
                            Set .Cell(flexcpPicture, i, col_ҽ������) = frmIcons.imgSign.ListImages("ǩ��").Picture
                            intͼ���� = 1
                        End If
                        
                        If Val(rsTmp!��ΣҩƷ & "") > 0 Then
                            If .Cell(flexcpPicture, i, col_ҽ������) Is Nothing Then
                                Set .Cell(flexcpPicture, i, col_ҽ������) = frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture
                                intͼ���� = 1
                            Else
                                If .Cell(flexcpPicture, i, col_ҽ������) <> frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture Then
                                    pictmp.Cls
                                    pictmp.PaintPicture .Cell(flexcpPicture, i, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
                                    pictmp.PaintPicture frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                                    Set .Cell(flexcpPicture, i, col_ҽ������) = pictmp.Image
                                    intͼ���� = 2
                                End If
                            End If
                        End If
         
                        If Val(rsTmp!��������� & "") = 2 Then '�������δͨ������ͼ��
                            .TextMatrix(i, COL_ѡ��) = 1
                            If intͼ���� = 0 Then
                                Set .Cell(flexcpPicture, i, col_ҽ������) = frmIcons.imgFlag.ListImages("���δͨ��").Picture
                            ElseIf intͼ���� = 1 Then
                                pictmp.Cls
                                pictmp.PaintPicture .Cell(flexcpPicture, i, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
                                pictmp.PaintPicture frmIcons.imgFlag.ListImages("���δͨ��").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                                Set .Cell(flexcpPicture, i, col_ҽ������) = pictmp.Image
                                intͼ���� = 2
                            ElseIf intͼ���� = 2 Then
                                pictmp.Cls
                                pictmp.Width = 720
                                pictmp.PaintPicture .Cell(flexcpPicture, i, col_ҽ������), 0, 0, 480, pictmp.Height
                                pictmp.PaintPicture frmIcons.imgFlag.ListImages("���δͨ��").Picture, 480, 0, 240, pictmp.Height
                                Set .Cell(flexcpPicture, i, col_ҽ������) = pictmp.Image
                                pictmp.Width = 480
                                intͼ���� = 3
                            End If
                        End If
                        
                        '�׵���ͼ��
                        If Val(rsTmp!�Ƿ��������� & "") > 0 Then
                            If intͼ���� = 0 Then
                                Set .Cell(flexcpPicture, i, col_ҽ������) = frmIcons.imgQuestion.ListImages("�׵���").Picture
                            ElseIf intͼ���� = 1 Then
                                pictmp.Cls
                                pictmp.PaintPicture .Cell(flexcpPicture, i, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
                                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("�׵���").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                                Set vsAdvice.Cell(flexcpPicture, i, col_ҽ������) = pictmp.Image
                                intͼ���� = 2
                            ElseIf intͼ���� = 2 Then
                                pictmp.Cls
                                pictmp.Width = 720
                                pictmp.PaintPicture .Cell(flexcpPicture, i, col_ҽ������), 0, 0, 480, pictmp.Height
                                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("�׵���").Picture, 480, 0, 240, pictmp.Height
                                Set .Cell(flexcpPicture, i, col_ҽ������) = pictmp.Image
                                pictmp.Width = 480
                                intͼ���� = 3
                            ElseIf intͼ���� = 3 Then
                                pictmp.Cls
                                pictmp.Width = 960
                                pictmp.PaintPicture .Cell(flexcpPicture, i, col_ҽ������), 0, 0, 720, pictmp.Height
                                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("�׵���").Picture, 720, 0, 240, pictmp.Height
                                Set .Cell(flexcpPicture, i, col_ҽ������) = pictmp.Image
                                pictmp.Width = 480
                                intͼ���� = 4
                            End If
                        End If
                        
                    End If
                    
                    If bln��ҩ;�� Or bln��Ѫ;�� Then .RemoveItem i
                    
                    Progress = rsTmp.AbsolutePosition / rsTmp.RecordCount * 100
                    
                    rsTmp.MoveNext
                Loop
            End With
        End If
    Next
        
    '----------------------------------------------------------------------
    '������Ϣ��ʾ
    If strPatis <> "" Then
        strPatis = Mid(strPatis, 2)
    End If
    If UBound(Split(strPatis, ";")) = 0 Then
        'ֻ��һ�����˵����ݵ����
        lng����ID = Split(strPatis, ",")(0)
        lng��ҳID = Split(strPatis, ",")(1)
        If lng����ID <> mlng����ID Or fraBaby.Visible Then '���ǵ�ǰ��������ȡ����ʾ
            strSQL = _
                " Select B.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� ,NVL(B.����,A.����) ����,B.��Ժ����," & _
                " B.סԺҽʦ,B.��Ժ����ID,C.���� as ����" & _
                " From ������Ϣ A,������ҳ B,���ű� C" & _
                " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID" & _
                " And B.����ID=[1] And B.��ҳID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
            lblPati.Caption = "����:" & rsTmp!���� & "��סԺ��:" & NVL(rsTmp!סԺ��) & _
                "������:" & NVL(rsTmp!��Ժ����) & "������:" & NVL(rsTmp!����)
        End If
    ElseIf UBound(Split(strPatis, ";")) > 0 Then
        '�ж���������ݵ����
        vsAdvice.ColHidden(COL_����) = False
        vsAdvice.ColHidden(COL_סԺ��) = False
        vsAdvice.ColHidden(COL_����) = False
                
        strSQL = "Select ���� From ���ű� Where ID IN(Select Column_Value From Table(f_Num2list([1])))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(str����s, 2))
        str����s = ""
        Do While Not rsTmp.EOF
            str����s = str����s & "," & rsTmp!����
            rsTmp.MoveNext
        Loop
        lblPati.Caption = "(" & Mid(str����s, 2) & ")���� " & UBound(Split(strPatis, ";")) + 1 & " �����˵�ҽ��"
    ElseIf UBound(Split(strPatis, ";")) = -1 Then
        'û���κβ������ݵ����
        lblPati.Caption = "��ǰ������û���κ������Ϣ��"
    End If
    
    '----------------------------------------------------------------------
    If vsAdvice.Rows = vsAdvice.FixedRows Then
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1
        vsPrice.Rows = vsPrice.FixedRows
        vsPrice.Rows = vsPrice.FixedRows + 1
    Else
        '����ǩ��ͼ�����
        vsAdvice.Cell(flexcpPictureAlignment, vsAdvice.FixedRows, col_ҽ������, vsAdvice.Rows - 1, col_ҽ������) = 0
        '�Զ������и�
        vsAdvice.AutoSize col_ҽ������
    End If
    Call SetAdviceCol
    vsAdvice.Row = vsAdvice.FixedRows
    If Not vsAdvice.ColHidden(COL_ѡ��) Then
        vsAdvice.Col = COL_ѡ��
    Else
        vsAdvice.Col = col_ҽ������
    End If
    vsAdvice.Redraw = flexRDDirect
    
    Progress = 0: Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    strPatis = ""
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Progress = 0
End Function

Private Function CheckValid(Optional ByRef bln�����ջ� As Boolean) As Boolean
'���ܣ�ȷ��ǰ���Ϸ���
'������bln�����ջ�=ȷ��ֹͣʱ�����Ƿ������Ҫ�����ջص�ҽ��
    Dim str���� As String, str���� As String
    Dim str���� As String, strTmp As String
    Dim curDate As Date, i As Long, k As Long
    Dim strPatis As String, strMsg As String
    Dim rsDrug As ADODB.Recordset, strUnRoll As String, lngҩƷID As Long, blnDo As Boolean
    Dim lngҽ��ID As Long
    Dim strTmpTim As String
    Dim strִ�еǼ� As String
    Dim lng���ID As Long
    Dim str��ֹԭ�� As String
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNurse As String
    
    mstrRollNotify = ""
    curDate = zlDatabase.Currentdate
    strUnRoll = zlDatabase.GetPara("��ҩ���ջ�", glngSys, pסԺҽ������)
    
    With vsAdvice
        '�Ƿ��п��Բ����ļ�¼
        If .Rows = .FixedRows + 1 And Val(.TextMatrix(.FixedRows, COL_ID)) = 0 Then
            If mint���� = 0 Then
                'ҽ������
                strTmp = "��ǰû�п������ϵ�ҽ����"
            ElseIf mint���� = 1 Then
                'ֹͣҽ��
                strTmp = "��ǰû�п���ֹͣ��ҽ����"
            ElseIf mint���� = 2 Then
                'ȷ��ֹͣ
                strTmp = "��ǰû�б�ֹͣ��ҽ����"
            ElseIf mint���� = 3 Then
                'ҽ��У��
                strTmp = "��ǰû���¿���ҽ����"
            ElseIf mint���� = 4 Then
                '�����Ƽ���Ŀ
                strTmp = "��ǰû��ͨ��У�Ե���Чҽ����"
            ElseIf mint���� = 5 Then
                '��ͣҽ��
                strTmp = "��ǰû�п�����ͣ��ҽ����"
            ElseIf mint���� = 6 Then
                '����ҽ��
                strTmp = "��ǰû����ͣ����Ҫ���õ�ҽ����"
            ElseIf mint���� = 7 Then
                'ͣ�����
                strTmp = "��ǰû�п�����˵�ͣ����"
            End If
            If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
            Exit Function
        End If
        
        On Error GoTo errH
        '�Ƿ���ѡ��
        str���� = "": str���� = "": str���� = ""
        If Not .ColHidden(COL_ѡ��) Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ID)) <> 0 And (Val(.TextMatrix(i, COL_ѡ��)) <> 0 Or .Cell(flexcpData, i, COL_ѡ��) <> Empty) Then
                    k = k + 1
                    If InStr(strPatis & ",", "," & .TextMatrix(i, COL_����ID)) = 0 Then
                        strPatis = strPatis & "," & .TextMatrix(i, COL_����ID)
                    End If
                    
                    If (mint���� = 1 Or mint���� = 7) Then
                        'ͣҽ��ʱ�����ִ�еǼ����
                        If mint���� = 1 And Not (.TextMatrix(i, COL_�������) = "Z" And InStr(",4,14,9,10,12,", "," & .TextMatrix(i, COL_��������) & ",") > 0) Then
                            lngҽ��ID = IIF(InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0, .TextMatrix(i, COL_���ID), .TextMatrix(i, COL_ID))
                            If lngҽ��ID <> lng���ID Then strTmpTim = GetAdviceStopTime(lngҽ��ID)
                            lng���ID = lngҽ��ID
                            If IsDate(strTmpTim) Then
                                If .TextMatrix(i, COL_ִ��ʱ��) = "" _
                                    And (Val(.TextMatrix(i, COL_Ƶ�ʴ���)) = 0 Or Val(.TextMatrix(i, COL_Ƶ�ʼ��)) = 0 Or .TextMatrix(i, COL_�����λ) = "") Then
                                    '"������"����,������
                                    If Format(.TextMatrix(i, COL_����), "yyyy-MM-dd") < Format(CDate(strTmpTim), "yyyy-MM-dd") Then
                                        strִ�еǼ� = strִ�еǼ� & vbCrLf & "��" & .TextMatrix(i, col_ҽ������) & " ִ��ʱ�䣺" & Format(CDate(strTmpTim), "yyyy-MM-dd")
                                    End If
                                Else
                                    If Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") < Format(strTmpTim, "yyyy-MM-dd HH:mm") Then
                                        strִ�еǼ� = strִ�еǼ� & vbCrLf & "��" & .TextMatrix(i, col_ҽ������) & " ִ��ʱ�䣺" & Format(CDate(strTmpTim), "yyyy-MM-dd  HH:mm")
                                    End If
                                End If
                            End If
                        End If
                         
                        '�ռ����ڷ��͵�ҽ��(�ſ�����ҽ����
                        If IsDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)) And .TextMatrix(i, COL_Ƶ��) <> "��Ҫʱ" And .TextMatrix(i, COL_Ƶ��) <> "��Ҫʱ" Then
                            If .TextMatrix(i, COL_ִ��ʱ��) = "" _
                                And (Val(.TextMatrix(i, COL_Ƶ�ʴ���)) = 0 Or Val(.TextMatrix(i, COL_Ƶ�ʼ��)) = 0 Or .TextMatrix(i, COL_�����λ) = "") Then
                                '"������"����,������
                                If Format(.TextMatrix(i, COL_����), "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)), "yyyy-MM-dd") Then
                                    str���� = str���� & vbCrLf & "��" & .TextMatrix(i, col_ҽ������)
                                End If
                            Else
                                If Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, i, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then
                                    '��������ջص�ҩƷ
                                    blnDo = True
                                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 And strUnRoll <> "" Then
                                        If .TextMatrix(i, COL_�շ�ϸĿID) <> "" Then
                                            lngҩƷID = Val(.TextMatrix(i, COL_�շ�ϸĿID))
                                        Else
                                            lngҩƷID = GetLastSendMediCineID(Val(.TextMatrix(i, COL_ID)), CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)), Val(.TextMatrix(i, COL_��������)))
                                        End If
                                        If lngҩƷID <> 0 Then
                                            gstrSQL = "Select ��ҩ���� From ҩƷ��� Where ҩƷID = [1] And ��ҩ���� is Not Null"
                                            Set rsDrug = zlDatabase.OpenSQLRecord(gstrSQL, "�����ջؼ��", lngҩƷID)
                                            If rsDrug.RecordCount > 0 Then
                                                If InStr("," & strUnRoll & ",", "," & rsDrug!��ҩ���� & ",") > 0 Then
                                                    If CheckMedicineSended(Val(.TextMatrix(i, COL_ID)), CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��))) Then
                                                        blnDo = False
                                                    End If
                                                End If
                                            End If
                                        Else '�����ջأ�ҽ��δ�Ƿѣ����Ա�ҩ������ط��ñ�ɾ���ˣ��绮�۵���ɾ����
                                            blnDo = False
                                        End If
                                    End If
                                    If blnDo Then
                                        str���� = str���� & vbCrLf & "��" & .TextMatrix(i, col_ҽ������)
                                    End If
                                End If
                            End If
                        End If
                        
                        '�ռ�����ֹͣ��ҽ��
                        If CDate(.TextMatrix(i, COL_����)) - curDate > 7 Then
                            str���� = str���� & vbCrLf & "��" & .TextMatrix(i, col_ҽ������) & "��ֹͣʱ�䣺" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm")
                        End If
                        
                        'δ��д��ֹԭ��
                        
                        If gblnҽ����ֹԭ�� And InStr(gstr�ɲ���ͣ��ԭ�����, "," & Val(.TextMatrix(i, COL_���˿���ID)) & ",") = 0 Then
                            If .TextMatrix(i, COL_��ֹԭ��) = "" Then
                                .Row = i: .ShowCell .Row, COL_��ֹԭ��
                                MsgBox "��ҽ��δ¼����ֹԭ��", vbInformation, gstrSysName
                                Exit Function
                            Else
                                If zlCommFun.ActualLen(.TextMatrix(i, COL_��ֹԭ��)) > txt����.MaxLength Then
                                    .Row = i: .ShowCell .Row, COL_��ֹԭ��
                                    MsgBox "��ҽ����ֹԭ������̫����������� " & txt����.MaxLength / 2 & " �����ֻ� " & txt����.MaxLength & " ���ַ���", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                        End If
                    ElseIf mint���� = 2 Then
                        '�ռ����ڷ��͵�ҽ��
                        If IsDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)) And .TextMatrix(i, COL_Ƶ��) <> "��Ҫʱ" And .TextMatrix(i, COL_Ƶ��) <> "��Ҫʱ" Then
                            If .TextMatrix(i, COL_ִ��ʱ��) = "" _
                                And (Val(.TextMatrix(i, COL_Ƶ�ʴ���)) = 0 Or Val(.TextMatrix(i, COL_Ƶ�ʼ��)) = 0 Or .TextMatrix(i, COL_�����λ) = "") Then
                                '"������"����,������
                                If Format(.Cell(flexcpData, i, COL_��ֹʱ��), "yyyy-MM-dd") < Format(CDate(.Cell(flexcpData, i, COL_�ϴ�ִ��)), "yyyy-MM-dd") Then
                                    str���� = str���� & vbCrLf & "��" & .TextMatrix(i, col_ҽ������)
                                End If
                            Else
                                If Format(.Cell(flexcpData, i, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, i, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then
                                    str���� = str���� & vbCrLf & "��" & .TextMatrix(i, col_ҽ������)
                                End If
                            End If
                        End If
                    ElseIf mint���� = 3 Then
                        '�ռ�������ҽ��,ͨ��У�ԵĲ��ж�
                        '3-ת��;4-����;5-��Ժ;6-תԺ,11-����,14-��ǰ
                        If .Cell(flexcpData, i, COL_ѡ��) = 1 And _
                            .TextMatrix(i, COL_�������) = "Z" And InStr(",3,4,5,6,11,14,", Val(.TextMatrix(i, COL_��������))) > 0 Then
                            
                            If InStr(str���� & ",", "," & .TextMatrix(i, COL_����ID) & ":" & .TextMatrix(i, COL_��ҳID) & ",") = 0 Then
                                str���� = str���� & "," & .TextMatrix(i, COL_����ID) & ":" & .TextMatrix(i, COL_��ҳID)
                            End If
                            
                            strMsg = strMsg & vbCrLf & .TextMatrix(i, COL_����) & _
                                IIF(.Cell(flexcpData, i, COL_Ӥ��) <> 0, "(Ӥ��" & .Cell(flexcpData, i, COL_Ӥ��) & ")", "") & "��" & .TextMatrix(i, col_ҽ������)
                            
                            'ת��ҽ�����
                            If Val(.TextMatrix(i, COL_��������)) = 3 Then
                                If CheckCanSendAdvice(Val(.TextMatrix(i, COL_����ID)), Val(.TextMatrix(i, COL_��ҳID)), Val(.TextMatrix(i, COL_ID)), Val(.Cell(flexcpData, i, COL_Ӥ��))) Then
                                    Call MsgBox("����ת��ҽ����" & vbCrLf & .TextMatrix(i, COL_����) & IIF(.Cell(flexcpData, i, COL_Ӥ��) <> 0, "(Ӥ��" & .Cell(flexcpData, i, COL_Ӥ��) & ")", "") & "��" & .TextMatrix(i, col_ҽ������) & vbCrLf & vbCrLf & "���뽫���Է��͵ĳ���ҽ����������У�ԡ�", vbInformation, gstrSysName)
                                    Exit Function
                                End If
                            End If
                        ElseIf .Cell(flexcpData, i, COL_ѡ��) = 2 Then
                            If zlCommFun.ActualLen(.TextMatrix(i, COL_����˵��)) > txt����.MaxLength Then
                                .Row = i: .ShowCell .Row, COL_ѡ��
                                MsgBox "��ҽ��У�����ʵ�˵������̫����������� " & txt����.MaxLength / 2 & " �����ֻ� " & txt����.MaxLength & " ���ַ���", vbInformation, gstrSysName
                                If txt����.Visible Then txt����.SetFocus
                                Exit Function
                            End If
                        End If
                        
                        '����ȼ�ҽ�����ж�
                        If .TextMatrix(i, COL_�������) = "H" And .TextMatrix(i, COL_��������) = "1" Then
                            If Check����ȼ��䶯����(.TextMatrix(i, COL_����ID), .TextMatrix(i, COL_��ҳID), .Cell(flexcpData, i, COL_��ʼʱ��)) Then
                                Exit Function
                            End If
                        End If
                    ElseIf mint���� = 0 Then
                        If mbln��������ִ�� Then
                            If .TextMatrix(i, COL_������ĿID) = "" Then
                                strSQL = "select nvl(max(decode(ִ�н��,1,1,0)),0) as ִ��״̬ from ����ҽ��ִ�� where ҽ��id=[1]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ҽ��ִ�м��", .TextMatrix(i, COL_ID))
                                
                                If Not rsTmp.EOF Then
                                    If InStr(",1,3,", NVL(rsTmp!ִ��״̬, 0)) > 0 Then
                                        MsgBox "����¼��ҽ��[" & .TextMatrix(i, col_ҽ������) & "]�Ѿ�ִ��,�������ϣ�", vbInformation, gstrSysName
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If k = 0 Then
                MsgBox "û��ѡ���κ�ҽ������ѡ����Ҫ" & tbr.Buttons("ִ��").Caption & "��ҽ����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
                
        'ҽ��
        If (mint���� = 1 Or mint���� = 7) And mbln��ʿվ Then
            If cboҽ��.ListIndex = -1 Then
                MsgBox "��ѡ��ֹͣҽ����ҽ����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    strTmp = ""
    strPatis = IIF(UBound(Split(Mid(strPatis, 2), ",")) > 0, "��ѡ���˶�����˵�ҽ��������ϸ���м���Ա�����ֲ��" & vbCrLf & vbCrLf, "")
    If mint���� = 0 Then
        'ҽ������
        strTmp = "ȷʵҪ�����Ѿ�ѡ���ҽ����"
    ElseIf (mint���� = 1 Or mint���� = 7) Then
        If strִ�еǼ� <> "" Then
            MsgBox "����ҽ������д��ִ�еǼǣ�" & vbCrLf & strִ�еǼ� & _
                vbCrLf & vbCrLf & "���޸�ֹͣʱ���ȡ��ִ�еǼǡ�", vbInformation, gstrSysName
            Exit Function
        End If
        'ֹͣҽ��
        If str���� <> "" Then '����Ƿ�����Ҫ�˻س�ǰ��ҩ�����
            If MsgBox("����ҽ�������ڷ��ͣ�" & vbCrLf & str���� & _
                vbCrLf & vbCrLf & "��ֹͣȷ�Ϻ����ʹ��""���ڷ����ջ�""���д���" & _
                vbCrLf & "Ҫ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        If str���� <> "" Then
            If MsgBox("����ҽ����ֹͣʱ�䳬����ǰʱ��̫�ã�" & vbCrLf & str���� & _
                vbCrLf & vbCrLf & "���ֹͣʱ�䲻��ȷ�������ҽ���ķ��ͺͼƷѲ���Ӱ�졣" & _
                vbCrLf & "ȷʵҪ��ָ����ʱ��ֹͣ��Щҽ����", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        If str���� = "" And str���� = "" Then
            strTmp = "ȷʵҪ" & IIF(mint���� = 7, "���", "ֹͣ") & "�Ѿ�ѡ���ҽ����"
        End If
    ElseIf mint���� = 2 Then
        'ȷ��ֹͣ
        If str���� <> "" And InStr(GetInsidePrivs(pסԺҽ������), ";���ڷ����ջ�;") > 0 Then
            bln�����ջ� = True
        Else
            strTmp = "ȷ���Ѿ�ѡ���ҽ��ֹͣ��"
        End If
    ElseIf mint���� = 3 Then
        'ҽ��У��
        If strMsg <> "" Then
            mstrRollNotify = Mid(str����, 2)
            
            '��������˵���ǩ����������"��ֹͣ��δȷ��ֹͣ"��ҽ������ʾ��ʿ�Ƚ���ȷ��ֹͣ
            '��Ϊ����ҽ��У��ʱ�Ὣ"��ֹͣ��δȷ��ֹͣ"��ҽ����"ִ����ֹʱ��"����Ϊ����ҽ���Ŀ�ʼִ��ʱ�䣬ҽ��ֹͣ��ǩ��Դ�İ�����"ִ����ֹʱ��"����ᵼ��ǩ����֤�޷�ͨ��
            If Mid(gstrESign, 2, 1) = "1" Then  'סԺҽ��վ�����˵���ǩ���ż��
                strTmp = ""
                '���ж�ʱ���ų�δ����ǩ�������´��ҽ��
                If CheckStopedUnAffirm(mstrRollNotify, strTmp) Then
                    MsgBox "ҪУ�Ե�ҽ���а�����������ҽ����" & vbCrLf & strMsg & _
                        vbCrLf & vbCrLf & "У�Ժ�Ὣδȷ��ֹͣ��ҽ������ֹͣ��Ϊ�˲�Ӱ��ǩ����֤�����ȶ����²��˽���ȷ��ֹͣ������" & strTmp, vbInformation, gstrSysName
                    Exit Function
                End If
                strTmp = ""
            End If
            
            
            If MsgBox(strPatis & "ҪУ�Ե�ҽ���а�����������ҽ����" & vbCrLf & strMsg & _
                vbCrLf & vbCrLf & "��Щҽ��У�Ժ��ֹͣ��������ҽ����ȷʵҪУ�Ե�ǰѡ���ҽ����", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            strTmp = strPatis & "ȷʵҪ���Ѿ�ѡ���ҽ������У�Դ�����"
        End If
    ElseIf mint���� = 5 Then
        '��ͣҽ��
        strTmp = strPatis & "ȷʵҪ��ͣ�Ѿ�ѡ���ҽ����"
    ElseIf mint���� = 6 Then
        '����ҽ��
        strTmp = strPatis & "ȷʵҪ�����Ѿ�ѡ���ҽ����"
    End If
    If strTmp <> "" Then
        If MsgBox(strTmp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    CheckValid = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckSignValid() As Boolean
'���ܣ�1.���δǩ����ҽ�����ܽ���У��
'      2.һ��ǩ����ҽ������һ��ͨ��У��
    Dim colҽ��ID As New Collection, strҽ��ID As String
    Dim colǩ��ID As New Collection, strǩ��ID As String
    Dim strסԺ As String, strҽ�� As String
    Dim lngǩ��id As Long, strTmp As String
    Dim int״̬ As Integer, i As Long, j As Long
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNurse As String
    
    If mint���� <> 3 Then CheckSignValid = True: Exit Function
    
    With vsAdvice
        '��ȡ��ʿ��Ա�б�ֻ�ǻ�ʿ������ҽ��
        If Mid(gstrESign, 2, 1) = "1" Or Mid(gstrESign, 3, 1) = "1" Then
            strNurse = ""
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ID)) <> 0 And Not .RowHidden(i) Then
                    If .Cell(flexcpData, i, COL_ѡ��) = 1 And Val(.TextMatrix(i, COL_ǩ��ID)) = 0 Then
                        If InStr(strNurse & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") = 0 Then
                            strNurse = strNurse & "," & Val(.TextMatrix(i, COL_ID))
                        End If
                    End If
                End If
            Next
            If strNurse <> "" Then
                strSQL = "Select /*+ Rule*/" & vbNewLine & _
                    " a.����,b.ҽ��ID" & vbNewLine & _
                    "From ��Ա�� A," & vbNewLine & _
                    "     (Select Distinct ������Ա,ҽ��ID" & vbNewLine & _
                    "       From ����ҽ��״̬" & vbNewLine & _
                    "       Where ҽ��id In (Select Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) And �������� = 1) B" & vbNewLine & _
                    "Where a.���� = b.������Ա And Exists (Select 1 From ��Ա����˵�� X Where x.��Աid = a.Id And x.��Ա���� = '��ʿ') And Not Exists" & vbNewLine & _
                    " (Select 1 From ��Ա����˵�� Y Where y.��Աid = a.Id And y.��Ա���� = 'ҽ��')" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strNurse, 2))
                On Error GoTo 0
                
                strNurse = ""
                Do While Not rsTmp.EOF
                    strNurse = strNurse & "," & rsTmp!ҽ��ID
                    rsTmp.MoveNext
                Loop
                strNurse = strNurse & ","
            End If
        End If
        
        For i = .FixedRows To .Rows - 1
            'flexcpData:0-������,1-У��,2-����
            If Val(.TextMatrix(i, COL_ID)) <> 0 And Not .RowHidden(i) Then
                '1.�ռ�δǩ����ҽ������
                If .Cell(flexcpData, i, COL_ѡ��) = 1 And Val(.TextMatrix(i, COL_ǩ��ID)) = 0 Then
                    '����Ϊʹ��ǩ���ĳ���
                    If InStr(strNurse, "," & Val(.TextMatrix(i, COL_ID)) & ",") = 0 Then '��ʿ¼���ҽ��������ǩ�����
                        If Val(.TextMatrix(i, COL_ǰ��ID)) = 0 Then
                            If CheckSign(1, Val(.TextMatrix(i, COL_��������ID)), , , , , gobjESign, .TextMatrix(i, COL_����ҽ��)) Then
                                If UBound(Split(strסԺ, vbCrLf)) < 10 Then
                                    strסԺ = strסԺ & vbCrLf & "��" & .TextMatrix(i, col_ҽ������)
                                ElseIf InStr(strסԺ, "�� ��") = 0 Then
                                    strסԺ = strסԺ & vbCrLf & "�� ��"
                                End If
                            End If
                        ElseIf Val(.TextMatrix(i, COL_ǰ��ID)) <> 0 Then
                            If CheckSign(3, Val(.TextMatrix(i, COL_��������ID)), , , , , gobjESign, .TextMatrix(i, COL_����ҽ��)) Then
                                If UBound(Split(strҽ��, vbCrLf)) < 10 Then
                                    strҽ�� = strҽ�� & vbCrLf & "��" & .TextMatrix(i, col_ҽ������)
                                ElseIf InStr(strҽ��, "�� ��") = 0 Then
                                    strҽ�� = strҽ�� & vbCrLf & "�� ��"
                                End If
                            End If
                        End If
                    End If
                End If
                
                '2.�ռ���ǩ��ҽ����У��״̬
                lngǩ��id = Val(.TextMatrix(i, COL_ǩ��ID))
                If lngǩ��id <> 0 Then
                    j = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID))) '��ID
                    int״̬ = .Cell(flexcpData, i, COL_ѡ��)
                    If int״̬ = 2 Then int״̬ = 0 '�������ʵ�ͬ�ڲ�У��
                    If InStr(strǩ��ID & ",", "," & lngǩ��id & ",") > 0 Then
                        '�ռ�����ǩ���ڽ����ϵ�У��״̬
                        strTmp = Split(colǩ��ID("_" & lngǩ��id), "=")(1)
                        If InStr(strTmp, int״̬) = 0 Then
                            colǩ��ID.Remove "_" & lngǩ��id
                            colǩ��ID.Add lngǩ��id & "=" & strTmp & int״̬, "_" & lngǩ��id
                        End If
                        
                        '�ռ�����ǩ���Ѷ��������ҽ��(��ID)
                        strTmp = colҽ��ID("_" & lngǩ��id)
                        If InStr("," & strTmp & ",", "," & j & ",") = 0 Then
                            colҽ��ID.Remove "_" & lngǩ��id
                            colҽ��ID.Add strTmp & "," & j, "_" & lngǩ��id
                        End If
                    Else
                        strǩ��ID = strǩ��ID & "," & lngǩ��id
                        colǩ��ID.Add lngǩ��id & "=" & int״̬, "_" & lngǩ��id
                        colҽ��ID.Add j, "_" & lngǩ��id
                    End If
                End If
            End If
        Next
        
        '�����ǩ��ҽ��У�����
        strTmp = "": strҽ��ID = Mid(strҽ��ID, 2)
        For i = 1 To colǩ��ID.Count
            lngǩ��id = Split(colǩ��ID(i), "=")(0)
            strǩ��ID = Split(colǩ��ID(i), "=")(1)
            
            '����һ��ǩ����δ��������δУ��ҽ��
            strҽ��ID = colҽ��ID("_" & lngǩ��id)
            strҽ��ID = ExistOtherSignAdvice(lngǩ��id, strҽ��ID)
            If strҽ��ID <> "" Then
                If InStr(strǩ��ID, "0") = 0 Then
                    strǩ��ID = strǩ��ID & "0"
                    strTmp = strTmp & strҽ��ID
                End If
            End If
            
            If Not (strǩ��ID = "1" Or strǩ��ID = "0") Then
                '���ǩ�������ݲ���"��Ҫͨ��У�Ի򶼲�ͨ��У��(��������)"�����
                j = .FindRow(CStr(lngǩ��id), , COL_ǩ��ID)
                Do While j <> -1
                    If Val(.TextMatrix(j, COL_ID)) <> 0 And Not .RowHidden(j) Then
                        If InStr(",0,2,", .Cell(flexcpData, j, COL_ѡ��)) > 0 Then
                            strTmp = strTmp & vbCrLf & .TextMatrix(j, COL_����) & "��" & IIF(Len(.TextMatrix(j, col_ҽ������)) > 40, Left(.TextMatrix(j, col_ҽ������), 40) & "...", .TextMatrix(j, col_ҽ������))
                        End If
                    End If
                    j = .FindRow(CStr(lngǩ��id), j + 1, COL_ǩ��ID)
                Loop
                Exit For '��ֻ��ʾ��һ��
            End If
        Next
    End With
    
    '1.û��ǩ����ҽ��������У�ԣ���סԺҽ����ҽ��ҽ���ֱ���м��
    If strסԺ <> "" Then
        MsgBox "����ҽ��ҽ����û��ǩ�������ܽ���У�ԣ�" & vbCrLf & strסԺ, vbInformation, gstrSysName
        Exit Function
    End If
    If strҽ�� <> "" Then
        MsgBox "����ҽ��ҽ����û��ǩ�������ܽ���У�ԣ�" & vbCrLf & strҽ��, vbInformation, gstrSysName
        Exit Function
    End If
    
    '2.һ��ǩ����ҽ������һ��ͨ��У��
    If strTmp <> "" Then
        MsgBox "����ҽ������������Ҫͨ��У�Ե�ҽ��һ��ǩ��������ǰ����Ϊ��У�Ի�У�����ʣ�" & vbCrLf & strTmp & _
            vbCrLf & vbCrLf & "һ��ǩ����ҽ������һ��ͨ��У�ԣ���������ҽ����У��״̬��", vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckSignValid = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExistOtherSignAdvice(ByVal lngǩ��id As Long, ByVal strҽ��ID As String) As String
'���ܣ�����Ƿ����ĳ���¿�ҽ��ǩ���б���û�ж�ȡ�������ϵ�ҽ��(��ΪҪһ��ͨ��У��,�����,��Щҽ��Ҳ��ûУ�Ե�)
'���أ�δ��ȡ�������δУ��ҽ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.����,B.ҽ������ From ����ҽ��״̬ A,����ҽ����¼ B" & _
        " Where A.ҽ��ID=B.ID And A.��������=1 And B.ҽ��״̬ IN(1,2)" & _
        " And (B.���ID is Null Or B.������� IN('5','6'))" & _
        " And Not Exists(Select 1 From ����ҽ����¼ S Where ������� IN('5','6') And S.���ID=B.ID)" & _
        " And Instr([2],','||Nvl(B.���ID,B.ID)||',')=0 And A.ǩ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngǩ��id, "," & strҽ��ID & ",")
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & vbCrLf & NVL(rsTmp!����) & "��" & IIF(Len(NVL(rsTmp!ҽ������)) > 40, Left(NVL(rsTmp!ҽ������), 40) & "...", NVL(rsTmp!ҽ������))
        rsTmp.MoveNext
    Loop
    ExistOtherSignAdvice = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lngҽ��ID As Long, int�������� As Integer, ByVal lng��Ŀid As Long, ByVal lngCol As Long)
'���ܣ���λ������ʾָ��ҽ����ָ���Ƽ���
'������lngRow=ҽ���к�,lngҽ��ID=�Ƽ�ҽ��ID
'      lng��ĿID=�Ƽ���ĿID,lngCol=�Ƽ۱����ʾ��
    Dim k As Long
    
    With vsAdvice
        .Row = lngRow: .Col = col_ҽ������ '�������Զ�ShowPrice,mrsPrice�����仯
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_ҽ��ID)) = lngҽ��ID _
                And Val(vsPrice.TextMatrix(k, COLP_��������)) = int�������� _
                And Val(vsPrice.TextMatrix(k, COLP_�շ�ϸĿID)) = lng��Ŀid Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Private Function ExecuteOperate() As Boolean
    Dim arrSQL As Variant, lng���ID As Long, blnTrans As Boolean
    Dim blnExe As Boolean, i As Long, j As Long
    Dim lngҽ��ID As Long, lngִ�п���ID As Long
    Dim strҽ��ID As String, intRule As Integer
    Dim lngǩ��id As Long, lng֤��ID As Long
    Dim strSource As String, strSign As String
    Dim strOper As String, strTimeStamp As String, strTimeStampCode As String
    Dim colSomeTime As New Collection
    Dim rsAdviceTmp As ADODB.Recordset
    Dim strAdvicesTmp As String
    Dim lngPatiID As Long
    Dim lngPageId As Long
    Dim lngLastRow As Long    '��һ�ι�ѡ������
    Dim strRevokeIDs As String, arrRevokeID() As String
    Dim str��Һ��ʾ As String, str��Һ��ֹ As String, str��Һ��� As String
    Dim strSQL As String, rsTmp As Recordset
    Dim lng����ȼ�ҽ��id As Long '�����������ϵĻ���ȼ�ҽ�����������Զ�ֹͣ�Ļ���ȼ�ҽ��id
    Dim blnPrintBeforeRedo As Boolean '�Ƿ�����Ѿ���ӡ����ҽ����ӡʱ�����������������֮ǰ
    Dim rsMsgRow As ADODB.Recordset
    Dim lngTmp As Long
    Dim strTmp As String
    Dim int���� As Integer '����ֹͣ��ҽ���Ƿ��ǽ���ҽ��
    Dim str������Ѫ��ʾ As String
    Dim str��ҩIDs As String, varTmp As Variant
    
    Dim strMsg As String
    Dim strPrintDel As String
    Dim arrPrintDel As Variant
    Dim lngLastPatiID As Long
    Dim lngLastPageID As Long
    Dim lngLastPatiDeptID As Long
    Dim rs��Ѫ As ADODB.Recordset
    Dim bln��Ѫ As Boolean
    Dim strErr As String
    
    Screen.MousePointer = 11
    
    mstrPatiKeepMsg = ""
    
    Call InitRecordSet(rsAdviceTmp, rsMsgRow, rs��Ѫ)
    
    '����SQL
    arrSQL = Array()
    With vsAdvice
        If mint���� = 3 Then
            If InitObjRecipeAudit(pסԺҽ���´�) Then
                '�������ϵͳ������������
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COL_ѡ��)) <> 0 Or .Cell(flexcpData, i, COL_ѡ��) <> Empty Then
                        If .TextMatrix(i, COL_�������) = "5" Or .TextMatrix(i, COL_�������) = "6" Then
                            If lngLastPatiID <> Val(.TextMatrix(i, COL_����ID)) Then
                                If Mid(str��ҩIDs, 2) <> "" Then
                                    Call gobjRecipeAudit.BuildData(Mid(str��ҩIDs, 2), lngLastPatiDeptID, 1, lngLastPatiID, lngLastPageID, strTmp)
                                    str��ҩIDs = ""
                                End If
                            End If
                            lngLastPatiID = Val(.TextMatrix(i, COL_����ID))
                            lngLastPageID = Val(.TextMatrix(i, COL_��ҳID))
                            lngLastPatiDeptID = Val(.TextMatrix(i, COL_���˿���ID))
                            If InStr("," & str��ҩIDs & ",", "," & .TextMatrix(i, COL_���ID) & ",") = 0 Then str��ҩIDs = str��ҩIDs & "," & .TextMatrix(i, COL_���ID)
                        End If
                    End If
                Next
                If Mid(str��ҩIDs, 2) <> "" Then
                    Call gobjRecipeAudit.BuildData(Mid(str��ҩIDs, 2), lngLastPatiDeptID, 1, lngLastPatiID, lngLastPageID, strTmp)
                End If
            End If
            '��������������ҩ�󷽽���ж�
            Call Check�������
        End If
        If mint���� <> 4 Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_ѡ��)) <> 0 Or .Cell(flexcpData, i, COL_ѡ��) <> Empty Then
                    'һ��ҽ��ֻУ��һ��,��һ����ҩ��,����ҽ��ֻ��һ����ʾ��
                    blnExe = False
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> lng���ID Then blnExe = True
                    Else
                        blnExe = True
                    End If
                    If blnExe Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        '(��ID)ʹ�����IDΪNULL��ҽ����ID(��ҩ;��,��ҩ�÷�,�����Ŀ,��Ҫ����,������ҽ��)
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                            lngҽ��ID = Val(.TextMatrix(i, COL_���ID))
                        Else
                            lngҽ��ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If mint���� = 0 Then      'ҽ������
                            'ҽ������ҽ������ǩ��
                            If Val(.TextMatrix(i, COL_ǩ��ID)) <> 0 And CheckSign(IIF(mbln��ʿվ, 2, 1), mlngҽ������ID, , , , , gobjESign) Then
                                strҽ��ID = strҽ��ID & "," & lngҽ��ID
                            End If
                            '���ϻ���ȼ�
                            If .TextMatrix(i, COL_�������) = "H" Then
                                lng����ȼ�ҽ��id = Get���˻���ȼ�ҽ��id(Val(.TextMatrix(i, COL_����ID)), Val(.TextMatrix(i, COL_��ҳID)), Val(.TextMatrix(i, COL_Ӥ��)), lngҽ��ID)
                            End If
                            
                            '92129:������Ѫ�������Ѫ�������ҽ����������
                            If .TextMatrix(i, COL_�������) = "K" And gblnѪ��ϵͳ Then
                                If InStr(1, ",2,5,6,", "," & Val(.TextMatrix(i, COL_���״̬)) & ",") <> 0 Then
                                    On Error GoTo errH
                                    strSQL = "Select Nvl(ִ�з���,0) as ִ�з��� from ����ҽ����¼ A, ������ĿĿ¼ B  where A.���ID  = [1] and A.������ĿID = B.ID"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������Ŀ��ִ�з���", lngҽ��ID)
                                    If rsTmp.RecordCount > 0 Then
                                        If Val(rsTmp!ִ�з���) = 0 Then
                                            str������Ѫ��ʾ = str������Ѫ��ʾ & vbCrLf & .TextMatrix(i, col_ҽ������)
                                        End If
                                    End If
                                    On Error GoTo 0
                                End If
                                Call rs��Ѫ.AddNew(Array("ҽ��ID", "����"), Array(lngҽ��ID, 4)): bln��Ѫ = True
                            End If
                            
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_����(" & lngҽ��ID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lng����ȼ�ҽ��id & ")"
                            strRevokeIDs = strRevokeIDs & "," & lngҽ��ID
                            rsMsgRow.Filter = "����id=" & Val(.TextMatrix(i, COL_����ID)) & " And ��ҳid=" & Val(.TextMatrix(i, COL_��ҳID)) & " And ��������=2"
                            If rsMsgRow.EOF Then
                                rsMsgRow.AddNew
                                rsMsgRow!����ID = Val(.TextMatrix(i, COL_����ID))
                                rsMsgRow!��ҳID = Val(.TextMatrix(i, COL_��ҳID))
                                rsMsgRow!�к� = i
                                rsMsgRow!�������� = 2
                                rsMsgRow.Update
                            End If
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_����Σ��ֵҽ��_Update(3,null," & lngҽ��ID & ")"   'ɾ��Σ��ֵ��Ӧ��ϵ
                        ElseIf (mint���� = 1 Or mint���� = 7) Then 'ֹͣҽ��
                            If mblnHaveAudit Then
                                'ҽ��ֹͣҽ������ǩ��
                                If Val(.TextMatrix(i, COL_ǩ��ID)) <> 0 And CheckSign(IIF(mbln��ʿվ, 2, 1), mlngҽ������ID, , , , , gobjESign) Then
                                    strҽ��ID = strҽ��ID & "," & lngҽ��ID
                                    '��¼ֹͣҽ����ִ����ֹʱ�䣺��������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                                    colSomeTime.Add Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm:00"), "_" & lngҽ��ID
                                End If
                            End If
                            '�����Һ��Һ��¼
                            '1������ҩ���ѷ��͵ļ�¼�����δ�������δ���ʵģ�������ֹͣ��
                            '2������ҩ���ѷ��͵ļ�¼������Ѵ������δ���ʵģ�����ֹͣ��������Ҫ��ʾ��
                            '3��������Ѱ�ҩ����δ��ҩ�ļ�¼������ֹͣ����Ҫ��ʾ��
                            '4��������Ѿ����ʵļ�¼������ֹͣ������ʾ��
                            If gstr��Һ�������� <> "" And (.TextMatrix(i, COL_�������) = "5" Or .TextMatrix(i, COL_�������) = "6") Then
                                strSQL = "Select Max(Decode(Instr(',4,5,6,7,8,', ',' || B.�������� || ','), 0, Null, To_Char(A.ִ��ʱ��, 'yyyy-MM-dd HH24:MI'))) As ����ִ��ʱ��," & _
                                    " Max(Decode(A.����״̬, 1, Null, To_Char(A.ִ��ʱ��, 'yyyy-MM-dd HH24:MI'))) As ��ʾִ��ʱ��, Min(A.�Ƿ���) As ���" & _
                                    " From ��Һ��ҩ��¼ A,��Һ��ҩ״̬ B Where A.ҽ��id = [1] and A.ID=B.��ҩID And A.ִ��ʱ�� > [2] and A.����״̬<>10"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, CDate(Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm")))
                                If rsTmp.RecordCount > 0 Then
                                    If rsTmp!����ִ��ʱ�� & "" <> "" Then
                                        If Val(rsTmp!��� & "") = 1 Then
                                            str��Һ��� = str��Һ��� & vbCrLf & .TextMatrix(i, col_ҽ������) & "��ִ��ʱ��:" & rsTmp!����ִ��ʱ�� & "��"
                                        Else
                                            '�ѷ��ͺ���ҩ�ģ���ֹֹͣ
                                            str��Һ��ֹ = str��Һ��ֹ & vbCrLf & .TextMatrix(i, col_ҽ������) & "��ִ��ʱ��:" & rsTmp!����ִ��ʱ�� & "��"
                                        End If
                                    End If
                                    If rsTmp!��ʾִ��ʱ�� & "" <> "" Then
                                        str��Һ��ʾ = str��Һ��ʾ & vbCrLf & .TextMatrix(i, col_ҽ������) & "��ִ��ʱ��:" & rsTmp!��ʾִ��ʱ�� & "��"
                                    End If
                                End If
                            End If
                            
                            strSQL = "ZL_����ҽ����¼_ֹͣ(" & lngҽ��ID & ",To_Date('" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                                 "'" & IIF(mbln��ʿվ, zlCommFun.GetNeedName(cboҽ��.Text), UserInfo.����) & "',0," & IIF(mblnHaveAudit, 1, 0) & "," & mlngͣ����� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'"
                            
                            If gblnҽ����ֹԭ�� Then
                                strSQL = strSQL & ",'" & .TextMatrix(i, COL_��ֹԭ��) & "')"
                            Else
                                strSQL = strSQL & ")"
                            End If
                            
                            arrSQL(UBound(arrSQL)) = strSQL
                            
                            '����������ʾ��ص�ҽ��ֹͣ
                            If .TextMatrix(i, COL_�������) = "H" And .TextMatrix(i, COL_��������) = "1" _
                                Or .TextMatrix(i, COL_�������) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_��������) & ",") > 0 Then
                                mblnRefresh = True
                            End If
                            rsMsgRow.Filter = "����id=" & Val(.TextMatrix(i, COL_����ID)) & " And ��ҳid=" & Val(.TextMatrix(i, COL_��ҳID)) & " And ��������=" & mint����
                            If rsMsgRow.EOF Then
                                rsMsgRow.AddNew
                                rsMsgRow!����ID = Val(.TextMatrix(i, COL_����ID))
                                rsMsgRow!��ҳID = Val(.TextMatrix(i, COL_��ҳID))
                                rsMsgRow!�к� = i
                                rsMsgRow!�������� = mint����
                                rsMsgRow.Update
                            End If
                            
                            If int���� = 0 And mint���� = 1 And .TextMatrix(i, COL_��־) = "����" Then int���� = 1
                            
                        ElseIf mint���� = 2 Then  'ȷ��ֹͣ
                            'ȷ��ֹͣҽ������ǩ��
                            If mbln��ʿǩ�� And Val(.TextMatrix(i, COL_ǩ��ID)) <> 0 And CheckSign(2, mlngҽ������ID, , , , , gobjESign) Then
                                If InStr(mstr����IDs, ";") > 0 Then
                                    If rsAdviceTmp.State = adStateClosed Then rsAdviceTmp.Open
                                    If Val(.TextMatrix(i, COL_����ID)) <> 0 And .TextMatrix(i, COL_����ID) & "|" & .TextMatrix(i, COL_��ҳID) <> .TextMatrix(lngLastRow, COL_����ID) & "|" & .TextMatrix(lngLastRow, COL_��ҳID) Then
                                        rsAdviceTmp.AddNew
                                        rsAdviceTmp!����ID = Val(.TextMatrix(i, COL_����ID))
                                        rsAdviceTmp!��ҳID = Val(.TextMatrix(i, COL_��ҳID))
                                    End If
                                    rsAdviceTmp!ҽ��ids = rsAdviceTmp!ҽ��ids & "," & lngҽ��ID
                                    rsAdviceTmp.Update
                                    lngLastRow = i
                                End If
                                
                                strҽ��ID = strҽ��ID & "," & lngҽ��ID
                                '��¼ȷ��ֹͣҽ��ʱ�䣺��������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                                colSomeTime.Add Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm:00"), "_" & lngҽ��ID
                            End If
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_ȷ��ֹͣ(" & lngҽ��ID & "," & _
                            "To_Date('" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & UserInfo.���� & "')"
                        ElseIf mint���� = 3 Then  'ҽ��У��
                            '��ʿУ��ҽ������ǩ����У�����ʲ�ǩ��
                            If mbln��ʿǩ�� And CheckSign(2, mlngҽ������ID, , , , , gobjESign) And Val(.TextMatrix(i, COL_ǩ��ID)) <> 0 And .Cell(flexcpData, i, COL_ѡ��) = 1 Then
                                If InStr(mstr����IDs, ";") > 0 Then
                                    If rsAdviceTmp.State = adStateClosed Then rsAdviceTmp.Open
                                    If Val(.TextMatrix(i, COL_����ID)) <> 0 And .TextMatrix(i, COL_����ID) & "|" & .TextMatrix(i, COL_��ҳID) <> .TextMatrix(lngLastRow, COL_����ID) & "|" & .TextMatrix(lngLastRow, COL_��ҳID) Then
                                        rsAdviceTmp.AddNew
                                        rsAdviceTmp!����ID = Val(.TextMatrix(i, COL_����ID))
                                        rsAdviceTmp!��ҳID = Val(.TextMatrix(i, COL_��ҳID))
                                    End If
                                    rsAdviceTmp!ҽ��ids = rsAdviceTmp!ҽ��ids & "," & lngҽ��ID
                                    rsAdviceTmp.Update
                                    lngLastRow = i
                                End If
                                
                                strҽ��ID = strҽ��ID & "," & lngҽ��ID
                                '��¼У��ҽ��ʱ�䣺��������ִ�й���֮ǰȡǩ��Դ��,��ʱ��δд�����ݿ�
                                colSomeTime.Add Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm:00"), "_" & lngҽ��ID
                            End If
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_У��(" & lngҽ��ID & "," & _
                                IIF(.Cell(flexcpData, i, COL_ѡ��) = 1, 3, 2) & "," & _
                                "To_Date('" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                                "'" & IIF(.Cell(flexcpData, i, COL_ѡ��) = 2, .TextMatrix(i, COL_����˵��), "") & "'," & _
                                "NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
                            
                            '����������ʾ��ص�ҽ��У��
                            If .TextMatrix(i, COL_�������) = "H" And .TextMatrix(i, COL_��������) = "1" _
                                Or .TextMatrix(i, COL_�������) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_��������) & ",") > 0 Then
                            End If
                            If Not mbln���͵��� Then
                                mblnRefresh = True  '���������ʱˢ����������嵥
                            End If
                            
                            '��ʱ������2011-1-13��ҽһԺ������������ҽ��ʱ�Զ�����У�������ĵģ��ĵò��ԣ�����Ϊ�ڷ��ʹ���������mblnRefresh��ֵ
                            'סԺ���߲�������У�Ժ���Ч��Ҫ������Ϣ������У��ͨ��
                            If .TextMatrix(i, COL_�������) = "Z" And InStr(",9,10,", "," & .TextMatrix(i, COL_��������) & ",") > 0 And Val(.Cell(flexcpData, i, COL_ѡ��)) = 1 Then
                                rsMsgRow.AddNew
                                rsMsgRow!����ID = Val(.TextMatrix(i, COL_����ID))
                                rsMsgRow!��ҳID = Val(.TextMatrix(i, COL_��ҳID))
                                rsMsgRow!�к� = i
                                rsMsgRow!�������� = 3
                                rsMsgRow!��ǰ���� = Get���˵�ǰ����(Val(.TextMatrix(i, COL_����ID)), Val(.TextMatrix(i, COL_��ҳID)))
                                rsMsgRow.Update
                            End If
                            
                            If Val(.Cell(flexcpData, i, COL_ѡ��)) = 2 Then
                                'У������
                                rsMsgRow.Filter = "����id=" & Val(.TextMatrix(i, COL_����ID)) & " And ��ҳid=" & Val(.TextMatrix(i, COL_��ҳID)) & " And ��������=4"
                                If rsMsgRow.EOF Then
                                    rsMsgRow.AddNew
                                    rsMsgRow!����ID = Val(.TextMatrix(i, COL_����ID))
                                    rsMsgRow!��ҳID = Val(.TextMatrix(i, COL_��ҳID))
                                    rsMsgRow!�к� = i
                                    rsMsgRow!�������� = 4
                                    rsMsgRow.Update
                                End If
                            End If
                            If .TextMatrix(i, COL_�������) = "K" And gblnѪ��ϵͳ Then
                                Call rs��Ѫ.AddNew(Array("ҽ��ID", "����"), Array(lngҽ��ID, 3)): bln��Ѫ = True
                            End If
                        ElseIf mint���� = 5 Then  '��ͣҽ��
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_��ͣ(" & lngҽ��ID & ",To_Date('" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & UserInfo.���� & "')"
                        ElseIf mint���� = 6 Then  '����ҽ��
                            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_����(" & lngҽ��ID & ",To_Date('" & Format(.TextMatrix(i, COL_����), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),'" & UserInfo.���� & "')"
                        End If
                    End If
                Else
                    '������δ��ѡ��ҽ����
                    If mint���� = 3 Or mint���� = 2 Then 'ҽ��У�Ժ�ȷ��ֹͣ
                        If InStr(";" & mstrPatiKeepMsg & ";", ";" & Val(.TextMatrix(i, COL_����ID)) & "," & Val(.TextMatrix(i, COL_��ҳID)) & ";") = 0 Then
                            mstrPatiKeepMsg = mstrPatiKeepMsg & ";" & Val(.TextMatrix(i, COL_����ID)) & "," & Val(.TextMatrix(i, COL_��ҳID))
                        End If
                    End If
                End If
                lng���ID = Val(.TextMatrix(i, COL_���ID))
            Next
            mstrPatiKeepMsg = Mid(mstrPatiKeepMsg, 2)
            If mint���� = 0 Then strRevokeIDs = Mid(strRevokeIDs, 2)
        End If
        
        If str��Һ��ֹ <> "" Then
            If MsgBox("����ҽ��ֹͣʱ��֮������Ѿ���ҩ���͵ļ�¼���Ƿ����ֹͣҽ��?" & str��Һ��ֹ, vbQuestion + vbYesNo + vbDefaultButton2, "ҽ��ֹͣ") = vbNo Then
                Screen.MousePointer = 0
                Exit Function
            End If
        ElseIf str��Һ��� <> "" Then
            If MsgBox("����ҽ��ֹͣʱ��֮������Ѿ���ҩ���ͣ����Ѿ�����ļ�¼���Ƿ����ֹͣҽ����" & str��Һ���, vbQuestion + vbYesNo + vbDefaultButton2, "ҽ��ֹͣ") = vbNo Then
                Screen.MousePointer = 0
                Exit Function
            End If
        ElseIf str��Һ��ʾ <> "" Then
            If MsgBox("����ҽ��ֹͣʱ��֮������Ѿ���ҩ����δ��ҩ�ļ�¼���Ƿ����ֹͣҽ����" & str��Һ��ʾ, vbQuestion + vbYesNo + vbDefaultButton2, "ҽ��ֹͣ") = vbNo Then
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
        
        If str������Ѫ��ʾ <> "" Then
            MsgBox "�������ϵ���Ѫҽ���Ѿ������Ѫ������ֱ������ҽ������Ҫ����������Ѫ����ϵ��" & str������Ѫ��ʾ, vbInformation, "ҽ������"
            Screen.MousePointer = 0
            Exit Function
        End If
        
        'ҽ����ӡ���
        If strRevokeIDs <> "" Then
            strPrintDel = Get���˴�ӡ��¼DelSQL(2, mlng����ID, mlng��ҳID, , , , strRevokeIDs, fraBaby.Visible, strMsg)
            If strMsg <> "" Then
                MsgBox "�����ϵ�ҽ���а����Ѿ���ӡ��ҽ�������ش�", vbInformation, gstrSysName
                strPrintDel = ""
            End If
        End If
        
        If mblnHaveAudit Or (mint���� <> 1 And mint���� <> 7) Then
            'ҽ���Ƽ۲���
            lng���ID = 0
            If mint���� = 2 Or mint���� = 3 Or mint���� = 4 Then
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COL_ѡ��)) <> 0 Or .Cell(flexcpData, i, COL_ѡ��) = 1 Then
                        'һ����ҩ��ֻ�账��һ��
                        blnExe = False
                        If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                            If Val(.TextMatrix(i, COL_���ID)) <> lng���ID Then blnExe = True
                        Else
                            blnExe = True
                        End If
                        
                        If blnExe Then
                            'ɾ����Ӧ�ļƼ�
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                arrSQL(UBound(arrSQL)) = "zl_����ҽ���Ƽ�_Delete(" & Val(.TextMatrix(i, COL_���ID)) & ")"
                            Else
                                arrSQL(UBound(arrSQL)) = "zl_����ҽ���Ƽ�_Delete(" & Val(.TextMatrix(i, COL_ID)) & ")"
                            End If
                            
                            '�����µļƼ�
                            '������һ����ѭ����Щ,��Ϊ���ж��Ƿ�Ҫ���漰����Ϸ���,������Filter
                            If Val(vsAdvice.TextMatrix(i, COL_���ID)) <> 0 Then
                                mrsPrice.Filter = "ҽ��ID=" & vsAdvice.TextMatrix(i, COL_ID) & _
                                    " Or ҽ��ID=" & Val(vsAdvice.TextMatrix(i, COL_���ID))
                            Else
                                mrsPrice.Filter = "ҽ��ID=" & vsAdvice.TextMatrix(i, COL_ID) & _
                                    " Or ���ID=" & vsAdvice.TextMatrix(i, COL_ID)
                            End If
                            For j = 1 To mrsPrice.RecordCount
                                '֮�д����շ�ϸĿIDΪ�յ����ü�¼(�������ȷ����ѡ�ļƼ�ҽ��)
                                'ҩƷ������ҽ���ļƼ۹̶���Ӧ�����棻�Ǹ������õ�ʱ�����ĵı����Ҫ���룬���Ҫ���浽�Ƽ۱���
                                If Not IsNull(mrsPrice!�շ�ϸĿID) And (InStr(",4,5,6,7,", mrsPrice!�������) = 0 _
                                    Or mrsPrice!������� = "4" And NVL(mrsPrice!����, 0) = 0 And NVL(mrsPrice!���, 0) = 1) Then
                                    If NVL(mrsPrice!����, 0) <> 0 Then '��������Ϊ0���Զ����˵�
                                        '��ͨ��Ŀ�ı�۵���Ҫ�����룬�����Ǹ������õ�ʱ������ҽ��
                                        If NVL(mrsPrice!����, 0) = 0 And NVL(mrsPrice!���, 0) = 1 _
                                            And Not (InStr(",5,6,7,", mrsPrice!�շ����) > 0 Or mrsPrice!�շ���� = "4" And NVL(mrsPrice!����, 0) = 1) Then
                                            Call SeekPriceRow(i, mrsPrice!ҽ��ID, mrsPrice!��������, mrsPrice!�շ�ϸĿID, COLP_����)
                                            Screen.MousePointer = 0
                                            MsgBox "����Ϊ��۵��շ���Ŀȷ��һ���շѼ۸�", vbInformation, gstrSysName
                                            vsPrice.SetFocus: Exit Function
                                        End If
                                        
                                        '�Ƽ�ִ�п���:ֻ�����ҩƷ������ҽ���ģ�ҩƷ�����ļƼ۵�ִ�п���
                                        If InStr(",4,5,6,7,", mrsPrice!�������) = 0 _
                                            And (InStr(",5,6,7,", mrsPrice!�շ����) > 0 Or mrsPrice!�շ���� = "4" And NVL(mrsPrice!����, 0) = 1) Then
                                            lngִ�п���ID = NVL(mrsPrice!ִ�п���ID, 0)
                                            
                                            '���ı�������ִ�п���
                                            If lngִ�п���ID = 0 And mrsPrice!�շ���� = "4" Then
                                                Call SeekPriceRow(i, mrsPrice!ҽ��ID, mrsPrice!��������, mrsPrice!�շ�ϸĿID, COLP_ִ�п���)
                                                Screen.MousePointer = 0
                                                MsgBox "����""" & vsPrice.TextMatrix(vsPrice.Row, COLP_�շ���Ŀ) & """û��ȷ��ִ�п��ң����ֹ�������ȷ��ִ�п��ҡ�" & vbCrLf & _
                                                    "�������ȷ����ȷ��ִ�п��ң��뵽""����Ŀ¼����""�м��洢�ⷿ�����Ƿ���ȷ��", vbInformation, gstrSysName
                                                vsPrice.SetFocus: Exit Function
                                            End If
                                        Else
                                            lngִ�п���ID = 0
                                        End If
                                        
                                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                        arrSQL(UBound(arrSQL)) = "zl_����ҽ���Ƽ�_Insert(" & mrsPrice!ҽ��ID & "," & _
                                            mrsPrice!�շ�ϸĿID & "," & mrsPrice!���� & "," & NVL(mrsPrice!����, 0) & "," & _
                                            NVL(mrsPrice!����, 0) & "," & ZVal(lngִ�п���ID) & "," & NVL(mrsPrice!��������, 0) & "," & NVL(mrsPrice!�շѷ�ʽ, 0) & ")"
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        End If
                    End If
                    lng���ID = Val(.TextMatrix(i, COL_���ID))
                Next
            End If
        
            '���ϻ�ֹͣʱ�ĵ���ǩ��
            If (mint���� = 0 Or (mint���� = 1 Or mint���� = 7) Or mint���� = 3 Or mint���� = 2) And strҽ��ID <> "" Then
                strOper = Decode(mint����, 0, "����", 1, "ֹͣ", 3, "У��", 2, "ȷ��ֹͣ")
            
                If (mint���� = 3 Or mint���� = 2) And rsAdviceTmp.State = adStateOpen Then
                    j = rsAdviceTmp.RecordCount
                    rsAdviceTmp.MoveFirst
                Else
                    j = 1
                End If
                For i = 1 To j
                    If (mint���� = 3 Or mint���� = 2) And rsAdviceTmp.State = adStateOpen Then
                        If rsAdviceTmp.EOF Then Exit For
                        strAdvicesTmp = rsAdviceTmp!ҽ��ids & ""
                        lngPatiID = Val(rsAdviceTmp!����ID & "")
                        lngPageId = Val(rsAdviceTmp!��ҳID & "")
                    Else
                        strAdvicesTmp = strҽ��ID
                        lngPatiID = mlng����ID
                        lngPageId = mlng��ҳID
                    End If
                    
                    '��ʿ�������ϡ�ֹͣҽ����ǩ����ҽ��
                    If mbln��ʿվ And mint���� <> 3 And mint���� <> 2 Then
                        MsgBox "��Ҫ" & strOper & "��ҽ���а���ҽ����ǩ����ҽ����ֻ����ҽ����" & strOper & "��ǩ����", vbInformation, gstrSysName
                        Screen.MousePointer = 0: Exit Function
                    End If
                    
                    'ҽ��ֹͣ,����ʱ����Ҫǩ��
                    If gobjESign Is Nothing Then
                        If gintCA = 0 Then
                            MsgBox strOper & "��ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ�����" & strOper & "��", vbInformation, gstrSysName
                        Else
                            MsgBox strOper & "��ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ������" & strOper & "��", vbInformation, gstrSysName
                        End If
                        Screen.MousePointer = 0: Exit Function
                    End If
                    
                    '��ȡǩ��ҽ��Դ��
                    strAdvicesTmp = Mid(strAdvicesTmp, 2) '��ID,����Ϊ��ϸID
                    intRule = ReadAdviceSignSource(Decode(mint����, 0, 4, 1, 8, 3, 3), lngPatiID, lngPageId, strAdvicesTmp, 0, False, strSource, , colSomeTime)
                    If intRule = 0 Then Screen.MousePointer = 0: Exit Function
                    If strSource = "" Then
                        Screen.MousePointer = 0
                        MsgBox "���ܶ�ȡ��Ҫ" & strOper & "����ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                     
                    strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
                    If strSign <> "" Then
                        If strTimeStamp <> "" Then
                            strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            strTimeStamp = "NULL"
                        End If
                        lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & "," & Decode(mint����, 0, 4, 1, 8, 3, 3, 2, 9) & "," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strAdvicesTmp & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
                    Else
                        Screen.MousePointer = 0: Exit Function
                    End If
                    
                    If (mint���� = 3 Or mint���� = 2) And rsAdviceTmp.State = adStateOpen Then
                        rsAdviceTmp.MoveNext
                    End If
                Next
                If rsAdviceTmp.State = adStateOpen Then rsAdviceTmp.Close
            End If
        End If
    End With
    varTmp = Split(strPrintDel, "|")
    
    If mint���� = 0 Then
        Call CreatePlugInOK(pסԺҽ���´�)
        If Not gobjPlugIn Is Nothing Then '��������ǰ��ҽӿ�
            On Error Resume Next
            arrRevokeID = Split(strRevokeIDs, ",")
            For i = 0 To UBound(arrRevokeID)
                If Val(arrRevokeID(i)) <> 0 Then
                    strMsg = ""
                    blnExe = gobjPlugIn.AdviceRevokedBefore(glngSys, pסԺҽ���´�, mlng����ID, mlng��ҳID, Val(arrRevokeID(i)), -1, strMsg)
                    Call zlPlugInErrH(err, "AdviceRevokedBefore")
                    If 0 = err.Number Then '�ӿ�û�г������������жϽӿڵķ���ֵ
                        If Not blnExe Then
                            MsgBox strMsg, vbInformation, gstrSysName
                            Screen.MousePointer = 0
                            Exit Function
                        End If
                    End If
                End If
            Next
            If err.Number <> 0 Then err.Clear
            On Error GoTo 0
        End If
    End If

    'ִ��SQL
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(varTmp)
        zlDatabase.ExecuteProcedure CStr(varTmp(i)), Me.Caption
    Next
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    If bln��Ѫ And gblnѪ��ϵͳ Then
        If InitObjBlood(True) = True Then
            rs��Ѫ.MoveFirst
            For i = 1 To rs��Ѫ.RecordCount
                If gobjPublicBlood.AdviceOperation(pסԺҽ���´�, Val(rs��Ѫ!ҽ��ID & ""), Val(rs��Ѫ!���� & ""), False, strErr) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False
                    Screen.MousePointer = 0
                    MsgBox "Ѫ��ϵͳ�ӿڵ���ʧ�ܣ�" & strErr, vbInformation, gstrSysName
                    Exit Function
                End If
                rs��Ѫ.MoveNext
            Next
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    ExecuteOperate = True
    With vsAdvice
        If Not rsMsgRow.EOF Then
            rsMsgRow.Filter = "��������=1" 'ֹͣ
            If Not rsMsgRow.EOF Then
                For i = 1 To rsMsgRow.RecordCount
                    j = Val(rsMsgRow!�к�)
                    '�����ʵϰ����ֹͣҽ�������ZLHIS_CIS_027����ͣ�������
                    If Not mblnHaveAudit And mlngͣ����� = 1 Then
                        Call ZLHIS_CIS_027(mclsMipModule, Val(.TextMatrix(j, COL_����ID)), .TextMatrix(j, COL_����), .TextMatrix(j, COL_סԺ��), , _
                            IIF(mlng�������� = 1, 1, 2), Val(.TextMatrix(j, COL_��ҳID)), mlng����ID, "", Val(.TextMatrix(j, COL_���˿���ID)), , , .TextMatrix(j, COL_����), _
                            Val(.TextMatrix(j, COL_ID)), 0, .TextMatrix(j, COL_�������), .TextMatrix(j, COL_��������), UserInfo.����, .TextMatrix(j, COL_ͣ��ʱ��), int����)
                    Else
                        Call ZLHIS_CIS_002(mclsMipModule, Val(.TextMatrix(j, COL_����ID)), .TextMatrix(j, COL_����), .TextMatrix(j, COL_סԺ��), , _
                            IIF(mlng�������� = 1, 1, 2), Val(.TextMatrix(j, COL_��ҳID)), mlng����ID, "", Val(.TextMatrix(j, COL_���˿���ID)), , , .TextMatrix(j, COL_����), _
                            Val(.TextMatrix(j, COL_ID)), 0, .TextMatrix(j, COL_�������), .TextMatrix(j, COL_��������), UserInfo.����, .TextMatrix(j, COL_ͣ��ʱ��), int����)
                    End If
                    rsMsgRow.MoveNext
                Next
            End If
            rsMsgRow.Filter = "��������=7" 'ֹͣ���
            If Not rsMsgRow.EOF Then
                For i = 1 To rsMsgRow.RecordCount
                    j = Val(rsMsgRow!�к�)
                    Call ZLHIS_CIS_002(mclsMipModule, Val(.TextMatrix(j, COL_����ID)), .TextMatrix(j, COL_����), .TextMatrix(j, COL_סԺ��), , _
                        IIF(mlng�������� = 1, 1, 2), Val(.TextMatrix(j, COL_��ҳID)), mlng����ID, "", Val(.TextMatrix(j, COL_���˿���ID)), , , .TextMatrix(j, COL_����), _
                        Val(.TextMatrix(j, COL_ID)), 0, .TextMatrix(j, COL_�������), .TextMatrix(j, COL_��������), UserInfo.����, .TextMatrix(j, COL_ͣ��ʱ��), int����)
                    rsMsgRow.MoveNext
                Next
            End If
            rsMsgRow.Filter = "��������=2" '����
            If Not rsMsgRow.EOF Then
                strTimeStamp = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                For i = 1 To rsMsgRow.RecordCount
                    j = Val(rsMsgRow!�к�)
                    Call ZLHIS_CIS_003(mclsMipModule, Val(.TextMatrix(j, COL_����ID)), .TextMatrix(j, COL_����), .TextMatrix(j, COL_סԺ��), , _
                        IIF(mlng�������� = 1, 1, 2), Val(.TextMatrix(j, COL_��ҳID)), mlng����ID, "", Val(.TextMatrix(j, COL_���˿���ID)), "", , .TextMatrix(j, COL_����), _
                        Val(.TextMatrix(j, COL_ID)), IIF(.TextMatrix(j, COL_��Ч) = "����", 0, 1), .TextMatrix(j, COL_�������), Val(.TextMatrix(j, COL_��������)), Val(.TextMatrix(j, COL_ִ�з���)), _
                        Val(.TextMatrix(j, COL_ִ�п���ID)), UserInfo.����, strTimeStamp)
                    rsMsgRow.MoveNext
                Next
            End If
            rsMsgRow.Filter = "��������=3" 'У��
            If Not rsMsgRow.EOF Then
                For i = 1 To rsMsgRow.RecordCount
                    j = Val(rsMsgRow!�к�)
                    lngTmp = 0: strTmp = ""
                    Call GetPatChange(Val(.TextMatrix(j, COL_ID)), 13, lngTmp, strTmp)
                    Set rsTmp = zlDatabase.OpenSQLRecord("select ���� from ���ű� where id=[1]", Me.Caption, Val(.TextMatrix(j, COL_���˿���ID)))
                    Call ZLHIS_PATIENT_005(mclsMipModule, Val(.TextMatrix(j, COL_����ID)), .TextMatrix(j, COL_��ҳID), .TextMatrix(j, COL_����), "", .TextMatrix(j, COL_סԺ��), _
                        0, , Val(.TextMatrix(j, COL_���˿���ID)), IIF(rsTmp.EOF, "", rsTmp!���� & ""), rsMsgRow!��ǰ���� & "", lngTmp, .TextMatrix(j, COL_����ʱ��), strTmp, .TextMatrix(i, COL_����ҽ��), Val(.TextMatrix(j, COL_ID)))
                    rsMsgRow.MoveNext
                Next
            End If
            
            rsMsgRow.Filter = "��������=4" 'У��������Ϣ
            If Not rsMsgRow.EOF Then
                strTimeStamp = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                For i = 1 To rsMsgRow.RecordCount
                    j = Val(rsMsgRow!�к�)
                    Call ZLHIS_CIS_035(mclsMipModule, Val(.TextMatrix(j, COL_����ID)), .TextMatrix(j, COL_����), .TextMatrix(j, COL_סԺ��), , _
                        IIF(mlng�������� = 1, 1, 2), Val(.TextMatrix(j, COL_��ҳID)), mlng����ID, "", Val(.TextMatrix(j, COL_���˿���ID)), "", , .TextMatrix(j, COL_����), _
                        Val(.TextMatrix(j, COL_ID)), IIF(.TextMatrix(j, COL_��Ч) = "����", 0, 1), .TextMatrix(j, COL_�������), Val(.TextMatrix(j, COL_��������)), Val(.TextMatrix(j, COL_ִ�з���)), _
                        Val(.TextMatrix(j, COL_ִ�п���ID)), UserInfo.����)
                    rsMsgRow.MoveNext
                Next
            End If
            
        End If
    End With
    Call ReadMsg
    Screen.MousePointer = 0
    If mint���� = 0 Then
        'ǰһ�ε�ҽ���Ƿ�������������ˣ�Ҫ����Ƿ񱻴�ӡ��
        If lng����ȼ�ҽ��id <> 0 Then
            strSQL = "Select Ӥ��,��Ч,ҳ�� From ����ҽ����ӡ Where ��ӡ��� In (1, 2) And ҽ��id = [1] And Not Exists" & _
                " (Select 1 From ����ҽ����¼ Where ID = [1] And ҽ��״̬ In (8, 9)) And Rownum < 2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ȼ�ҽ��id)
            If Not rsTmp.EOF Then
                If MsgBox("���������˻���ȼ�ҽ�����Զ�����ǰһ���Զ�ֹͣ�Ļ���ȼ�ҽ�������õĻ���ȼ�ҽ���Ѿ�������ҽ����ֹͣʱ���ȷ��ֹͣʱ����״�" & _
                    "���Ҫ�����ӡ�����ӵ�" & rsTmp!ҳ�� & "��ʼ������Ƿ�Ҫ������ش�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    strSQL = "Zl_����ҽ����ӡ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & Val(rsTmp!Ӥ�� & "") & "," & Val(rsTmp!��Ч & "") & "," & Val(rsTmp!ҳ�� & "") & ")"
                    On Error GoTo errH
                    gcnOracle.BeginTrans: blnTrans = True
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    gcnOracle.CommitTrans: blnTrans = False
                    On Error GoTo 0
                End If
            End If
        End If
        
        Call CreatePlugInOK(pסԺҽ���´�)
        '�������Ϻ���ҽӿ�
        On Error Resume Next
        If Not gobjPlugIn Is Nothing Then
            arrRevokeID = Split(strRevokeIDs, ",")
            For i = 0 To UBound(arrRevokeID)
                If Val(arrRevokeID(i)) <> 0 Then
                    Call gobjPlugIn.AdviceRevoked(glngSys, pסԺҽ���´�, mlng����ID, mlng��ҳID, Val(arrRevokeID(i)))
                    Call zlPlugInErrH(err, "AdviceRevoked")
                End If
            Next
        End If
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
    End If
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitPriceRecordset()
'˵�����༭ʱ,���Ƽ�ҽ�����շ���Ŀ�������,�ż����¼��
    Set mrsPrice = New ADODB.Recordset
    mrsPrice.Fields.Append "ҽ��ID", adBigInt
    mrsPrice.Fields.Append "���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "�������", adVarChar, 1
    mrsPrice.Fields.Append "������ĿID", adBigInt
    
    mrsPrice.Fields.Append "�걾��λ", adVarChar, 100, adFldIsNullable
    mrsPrice.Fields.Append "��鷽��", adVarChar, 100, adFldIsNullable
    mrsPrice.Fields.Append "ִ�б��", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "��������", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "�շѷ�ʽ", adInteger, , adFldIsNullable
    
    mrsPrice.Fields.Append "�շ����", adVarChar, 1, adFldIsNullable
    mrsPrice.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "����", adDouble, , adFldIsNullable
    mrsPrice.Fields.Append "����", adDouble, , adFldIsNullable
    mrsPrice.Fields.Append "����", adInteger '�����Ƿ��������
    mrsPrice.Fields.Append "���", adInteger
    mrsPrice.Fields.Append "����", adInteger
    mrsPrice.Fields.Append "�̶�", adInteger '���е��շѹ�ϵ���Ƿ�̶�����
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub ShowDefaultRow()
'���ܣ����ڿ��ԼƼ۵�ҽ��,ȱʡ����һ�в�����ȱʡ�Ƽ�ҽ��
'˵����ComboList="#ҽ��ID1;�Ƽ�ҽ��1|#ҽ��ID2;�Ƽ�ҽ��2|..."
'      ���ڵ�һ����ʾ�Ƽ۱�ͻس�������ʱ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim arrCombo As Variant, lngRow As Long
    Dim lngҽ��ID As Long, int�������� As String, str�Ƽ�ҽ�� As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If .ColData(COLP_�Ƽ�ҽ��) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_�Ƽ�ҽ��), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_ҽ��ID)) <> 0 _
                And Val(.TextMatrix(.Rows - 1, COLP_�շ�ϸĿID)) <> 0 Then
                '��һ����ʾʱȱʡ����һ��
                blnFirst = True
                .AddItem "", .Rows
                .Row = .Rows - 1
            End If
            lngRow = .Rows - 1
            
            '���ǵ�һ����ʾʱȱʡ�Ƽ�ҽ������һ����ͬ
            If lngRow > 1 And Not blnFirst Then
                If Val(.TextMatrix(lngRow - 1, COLP_�̶�)) = 0 _
                    And Val(.TextMatrix(lngRow - 1, COLP_ҽ��ID)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                int�������� = 0
                lngҽ��ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str�Ƽ�ҽ�� = Replace(arrCombo(i), "#" & lngҽ��ID & ";", "")
                
                If lngҽ��ID < 0 Then
                    int�������� = Val(Left(Abs(lngҽ��ID), 1))
                    lngҽ��ID = Val(Mid(Abs(lngҽ��ID), 2))
                End If
                If blnHave Then
                    If lngҽ��ID = Val(.TextMatrix(lngRow - 1, COLP_ҽ��ID)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            'ģ��ѡ������Ƽ�ҽ��
            strSQL = "Select ���ID,�������,������ĿID From ����ҽ����¼ Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
            If Not rsTmp.EOF Then
                .TextMatrix(lngRow, COLP_ҽ��ID) = lngҽ��ID
                .TextMatrix(lngRow, COLP_��������) = int��������
                .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = str�Ƽ�ҽ��
                .TextMatrix(lngRow, COLP_���ID) = NVL(rsTmp!���ID)
                .TextMatrix(lngRow, COLP_������ĿID) = rsTmp!������ĿID
                .TextMatrix(lngRow, COLP_�������) = rsTmp!�������
                .Cell(flexcpData, lngRow, COLP_�Ƽ�ҽ��) = .TextMatrix(lngRow, COLP_�Ƽ�ҽ��)
                
                'ֻ��һ���Ƽ�ҽ��ʱ����ͣ��
                If UBound(arrCombo) = 0 Then
                    .Col = COLP_�շ���Ŀ
                Else
                    .Col = COLP_�Ƽ�ҽ��
                End If
            End If
        End If
        Call .ShowCell(.Row, .Col)
        If blnFirst Then .TopRow = .Row '��һ����ʾʱ,ShowCell��Ȼ��������
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset, strSQL As String, i As Long
    Dim lngҽ��ID As Long, int�������� As Integer
    Dim lngԭ��ID As Long, intԭ�������� As Integer
    Dim lng�շ�ϸĿID As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_�Ƽ�ҽ�� Then
            '�������ComboData,TextMatrixȡֵ��ΪComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lngҽ��ID = .ComboData
                If lngҽ��ID < 0 Then
                    int�������� = Val(Left(Abs(lngҽ��ID), 1))
                    lngҽ��ID = Val(Mid(Abs(lngҽ��ID), 2))
                End If
                lngԭ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                intԭ�������� = Val(.TextMatrix(Row, COLP_��������))
                lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                
                '���üƼ�ҽ���Ƿ�������ͬ�շ�ϸĿ
                If lng�շ�ϸĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    If Not mrsPrice.EOF Then
                        MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """�Ѿ��������շ���Ŀ""" & .TextMatrix(Row, COLP_�շ���Ŀ) & """��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                                                                
                'ԭ����ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                If lngԭ��ID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And ��������=" & intԭ�������� & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, Row, Col) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                                                                
                '������ݣ�mrsPrice�п�����ɾ��,����Ҫ�����ݿ��
                strSQL = "Select ���ID,�������,������ĿID,�걾��λ,��鷽��,ִ�б�� From ����ҽ����¼ Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                If rsTmp.EOF Then
                    MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """�����Ѿ���������ɾ��,���˳����½��롣", vbInformation, gstrSysName
                    Exit Sub
                End If
                .TextMatrix(Row, COLP_ҽ��ID) = lngҽ��ID
                .TextMatrix(Row, COLP_��������) = int��������
                .TextMatrix(Row, COLP_���ID) = NVL(rsTmp!���ID)
                .TextMatrix(Row, COLP_������ĿID) = rsTmp!������ĿID
                .TextMatrix(Row, COLP_�������) = rsTmp!�������
                
                '��¼������
                If lng�շ�ϸĿID <> 0 Then
                    '��ѡ���ҽ���Ƿ��д�������޸ĺ����Ŀ�Ƿ����
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And ����=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_����) = IIF(blnHaveSub, "��", "")
                    
                    If lngԭ��ID = 0 Then
                        mrsPrice.AddNew '����
                    Else '����
                        mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    End If
                    mrsPrice!ҽ��ID = lngҽ��ID
                    mrsPrice!���ID = rsTmp!���ID
                    mrsPrice!������ĿID = rsTmp!������ĿID
                    mrsPrice!������� = rsTmp!�������
                    
                    mrsPrice!�걾��λ = rsTmp!�걾��λ
                    mrsPrice!��鷽�� = rsTmp!��鷽��
                    mrsPrice!ִ�б�� = NVL(rsTmp!ִ�б��, 0)
                    mrsPrice!�������� = int��������
                    mrsPrice!�շѷ�ʽ = 0
                    
                    If lngԭ��ID = 0 Then
                        mrsPrice!�շ�ϸĿID = lng�շ�ϸĿID
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_����))
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_����))
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_��������))
                        mrsPrice!��� = Val(.Cell(flexcpData, Row, 0))
                        mrsPrice!�̶� = 0
                    End If
                    mrsPrice!���� = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            End If
        ElseIf Col = COLP_�շ���Ŀ Or Col = COLP_ִ�п��� Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
        ElseIf Col = COLP_���� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
            int�������� = Val(.TextMatrix(Row, COLP_��������))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                Call SelectRow(vsAdvice.Row)
            End If
        ElseIf Col = COLP_���� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, 1), .Cell(flexcpData, Row, 2), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gstrDecPrice)
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
            int�������� = Val(.TextMatrix(Row, COLP_��������))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                Call SelectRow(vsAdvice.Row)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vPoint As PointAPI
    On Error GoTo errH
    With vsAdvice
        If Col = COL_��ֹԭ�� Then
            strSQL = "select a.���� as id, a.����,a.����,a.���� from ͣ��ԭ�� a order by a.����"
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ͣ��ԭ��", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COL_��ֹԭ��) = rsTmp!���� & ""
                .Cell(flexcpData, Row, COL_��ֹԭ��) = rsTmp!���� & ""
                Call SetSameԭ��(Row)
            Else
                If Not blnCancel Then
                    MsgBox "û�п��õ�ͣ��ԭ�����ȵ��ֵ���������ã�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str��ĿIDs As String, blnCancel As Boolean
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim int�������� As Integer, vPoint As PointAPI
    Dim strStock As String
    Dim strSQL2 As String
    
    With vsPrice
        If Col = COLP_�շ���Ŀ Then
            '����ѡ�����е���Ŀ
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_ҽ��ID)) = Val(.TextMatrix(Row, COLP_ҽ��ID)) _
                    And Val(.TextMatrix(Row, COLP_ҽ��ID)) <> 0 And i <> Row Then
                    str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                End If
            Next
            str��ĿIDs = Mid(str��ĿIDs, 2)
            
            If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��ҳID)), "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
            
            
            'ҩƷ���Ŀ��
            Call GetDefaultDeptPar(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��Ժ����ID)))
            If mlng��ҩ�� <> 0 Or mlng��ҩ�� <> 0 Or mlng��ҩ�� <> 0 Or mlng���ϲ��� <> 0 Then
                strStock = _
                    "Select A.ҩƷID,Sum(Nvl(A.��������,0)) as ���" & _
                    " From ҩƷ��� A,�շ���ĿĿ¼ B" & _
                    " Where A.���� = 1 And (Nvl(A.����,0)=0 Or A.Ч�� Is Null Or A.Ч��>Trunc(Sysdate))" & _
                        " And A.�ⷿID=Decode(B.���,'5',[3],'6',[4],'7',[5],'4',[6],Null)" & _
                        " And A.ҩƷID=B.ID And B.��� IN('4','5','6','7')" & _
                        " And (b.ִ�п��� <> 4 Or Exists (Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid = b.Id And (w.������Դ=2 or (w.������Դ is Null And w.��������id = [7]))))" & _
                    " Group by A.ҩƷID Having Sum(Nvl(A.��������,0))<>0"
            Else
                strStock = "Select Null as ҩƷID,Null as ��� From Dual"
            End If
            
            strSQL = _
                " Select Distinct 0 as ĩ��,To_Number('999999999'||����) as ID,-NULL as �ϼ�ID," & _
                " CHR(13)||���� as ����,Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ',7,'��������') as ����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as �۸�,NULL as ���,NULL as ��������,NULL as ҽ������," & _
                " NULL as ˵��,-NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-NULL as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,-ID as ID,Nvl(-�ϼ�ID,To_Number('999999999'||����)) as �ϼ�ID,����,����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as �۸�,NULL as ���,NULL as ��������,NULL as ҽ������," & _
                " NULL as ˵��,-NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-NULL as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,ID,�ϼ�ID,����,����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as �۸�,NULL as ���,NULL as ��������,NULL as ҽ������," & _
                " NULL as ˵��,-NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as ȱʡ�۸�ID,-NULL as �Ƿ���ID,Null as ���ID,-NULL as ��������ID" & _
                " From �շѷ���Ŀ¼ Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL2 = _
                " Select A.ĩ��,A.ID,A.�ϼ�ID,A.����,A.����,A.��λ,A.���,A.����,A.���," & _
                " Decode(Nvl(A.�Ƿ���,0),1,Decode(Instr('567',A.���ID),0,Sum(Nvl(A.ԭ��,0))||'-'||Sum(Nvl(A.�ּ�,0))||'/'||��λ,'ʱ��')," & _
                "   Decode(Instr('567',A.���ID),0,Sum(A.�ּ�)||'/'||A.��λ,LTrim(To_Char(Sum(A.�ּ�)*A.סԺ��װ,'999990.0000'))||'/'||A.סԺ��λ)) as �۸�," & _
                " Decode(Instr('4567',A.���ID),0,NULL,1," & _
                "   Decode(S.���,NULL,NULL,LTrim(To_Char(S.���,'999990.0000'))||A.��λ)," & _
                "   Decode(S.���,NULL,NULL,LTrim(To_Char(S.���/Nvl(A.סԺ��װ,1),'999990.0000'))||A.סԺ��λ)) as ���," & _
                " A.��������,A.ҽ������,A.˵��,Sum(A.ԭ��) as ԭ��ID,Sum(A.�ּ�) as �ּ�ID,Sum(A.ȱʡ�۸�) as ȱʡ�۸�ID,A.�Ƿ��� as �Ƿ���ID,A.���ID,A.��������ID" & _
                " From (" & _
                " Select Distinct 1 as ĩ��,A.ID,Decode(Instr('567',A.���),0,A.����ID,-E.����ID) as �ϼ�ID,A.����,A.����," & _
                " A.���㵥λ as ��λ,A.���,A.����,A.��� as ���ID,C.���� as ���,A.��������,N.���� as ҽ������,A.˵��,B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���," & _
                " -NULL as ��������ID,D.סԺ��λ,D.סԺ��װ" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,������ĿĿ¼ E,����֧����Ŀ M,����֧������ N" & _
                " Where A.ID=B.�շ�ϸĿID [ѡ���滻�Ĺ�����1]  And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "8", "9", "10") & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.��� Not IN('4','J','1') And A.���=C.���� And A.ID=D.ҩƷID(+) And D.ҩ��ID=E.ID(+)" & _
                " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[2]" & _
                " And (Nvl(a.ִ�п���,0) <> 4 Or Exists (Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid = a.Id And (w.������Դ=2 or (w.������Դ is Null And Nvl(w.��������id,[7]) = [7]))))" & _
                " And (a.��� Not in ('5','6','7') Or Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[7])=[7]))"
            If DeptExist("���ϲ���", 2) Then
                strSQL2 = strSQL2 & " Union ALL" & _
                    " Select Distinct 1 as ĩ��,A.ID,-E.����ID as �ϼ�ID,A.����,A.����," & _
                    " A.���㵥λ as ��λ,A.���,A.����,A.��� as ���ID,C.���� as ���,A.��������,N.���� as ҽ������,A.˵��," & _
                    " B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���,D.�������� as ��������ID,NULL as סԺ��λ,NULL as סԺ��װ" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,�������� D,������ĿĿ¼ E,����֧����Ŀ M,����֧������ N" & _
                    " Where A.ID=B.�շ�ϸĿID [ѡ���滻�Ĺ�����2]  And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "8", "9", "10") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And A.ID Not IN(" & str��ĿIDs & ")", "") & _
                    " And A.���='4' And A.���=C.���� And A.ID=D.����ID And D.����ID=E.ID And D.�������=0" & _
                    " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[2]" & _
                    " And Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[7])=[7])"
            End If
            strSQL2 = strSQL2 & " ) A,(" & strStock & ") S Where A.ID=S.ҩƷID(+)" & _
            " Group by A.ĩ��,A.ID,A.�ϼ�ID,A.����,A.����,A.��λ,A.���,A.����,A.���,A.��������,A.ҽ������,A.˵��,A.�Ƿ���,A.���ID,A.��������ID,A.סԺ��λ,A.סԺ��װ,S.���"
            '[ѡ���滻�Ĺ�����1],[ѡ���滻�Ĺ�����2],����������ѡ���д����
            'Ҫȷ�� "ռλ����" �����һλ���ò�����ѡ������ƴ�ӣ�Ҫ���4000���ȵ�����
            Set rsTmp = ShowSQLSelectCIS(Me, strSQL, strSQL2, 2, "�շ���Ŀ", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, _
                "," & str��ĿIDs & ",", mint����, mlng��ҩ��, mlng��ҩ��, mlng��ҩ��, mlng���ϲ���, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "ռλ����")
            If Not rsTmp Is Nothing Then
                'ҽ��������
                If CheckItemInsure(rsTmp, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��ҳID))) Then
                    .SetFocus: Exit Sub
                End If
            
                lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                Call SetItemInput(Row, rsTmp, lngҽ��ID, int��������, lngԭ��ĿID)
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û�п��õ��շ���Ŀ�����ȵ��շ���Ŀ���������ã�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        ElseIf Col = COLP_ִ�п��� Then
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            If .TextMatrix(Row, COLP_�շ����) = "4" Then
                '�������õ�����
                strSQL = _
                    " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                    " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                    " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                    " And B.������� IN(2,3) And B.����ID=C.ID" & _
                    " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And (A.������Դ is NULL Or A.������Դ=2)" & _
                    " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                    " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                    " And A.�շ�ϸĿID=[1]" & _
                    " Order by B.�������,C.����"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)))
            ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                'ҩƷ
                'ҩƷ��ϵͳָ���Ĵ���ҩ������
                If Not Check�ϰల��(True) Then
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                Else
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                End If
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                    Decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"))
            End If
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                .TextMatrix(Row, Col) = rsTmp!����
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!ִ�п���ID = rsTmp!ID
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
End Sub

Private Function CheckItemInsure(rsInput As ADODB.Recordset, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��������(ѡ��)�Ƽ���Ŀ�Ƿ�ҽ������
'���أ����δ���룬������ʾѡ�񲻼������򷵻��档
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, int���� As Integer
    
    If gintҽ������ = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckItemInsure", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then int���� = NVL(rsTmp!����, 0)
    If int���� <> 0 Then
        If Not ItemExistInsure(lng����ID, rsInput!ID, int����) Then
            If gintҽ������ = 1 Then
                If MsgBox("��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckItemInsure = True
                End If
            ElseIf gintҽ������ = 2 Then
                MsgBox "��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��", vbInformation, gstrSysName
                CheckItemInsure = True
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lngҽ��ID As Long, int�������� As Integer, ByVal lngԭ��ĿID As Long)
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    Dim lng����ID As Long, lng��ҳID As Long
    Dim lng�к� As Long, dbl���� As Double
    Dim blnHaveSub As Boolean, dbl���� As Double
    Dim rsTmp As ADODB.Recordset
    
    With vsPrice
        '�������
        .TextMatrix(lngRow, COLP_�շ����) = rsInput!���ID
        .TextMatrix(lngRow, COLP_�շ�ϸĿID) = rsInput!ID
        .TextMatrix(lngRow, COLP_���) = rsInput!���
        .TextMatrix(lngRow, COLP_�շ���Ŀ) = rsInput!����
        If Not IsNull(rsInput!����) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & "(" & rsInput!���� & ")"
        End If
        If Not IsNull(rsInput!���) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & " " & rsInput!���
        End If
        
        '����Ǽ���ҩƷ�Ƽ�(��ҩ��),�����۵�λ����
        .TextMatrix(lngRow, COLP_����) = 1 'ȱʡ�Ƽ�����Ϊ1
        .TextMatrix(lngRow, COLP_��λ) = NVL(rsInput!��λ)
                
        '���ۼ��㴦��:ҩ���Ƽ۲����������ﴦ��,��ҩ��ҩƷ�Ƽ۰��ۼ۴���
        .Cell(flexcpData, lngRow, 0) = 0
        .Cell(flexcpData, lngRow, 1) = 0
        .Cell(flexcpData, lngRow, 2) = 0
        
        'ִ�п���
        lng�к� = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
        If lng�к� = -1 Then
            Set rsTmp = Sys.RowValue("����ҽ����¼", lngҽ��ID)
            lng����ID = rsTmp!����ID
            lng��ҳID = NVL(rsTmp!��ҳID, 0)
            lngִ�п���ID = NVL(rsTmp!ִ�п���ID, 0)
            lng���˿���ID = NVL(rsTmp!���˿���id, 0)
            dbl���� = NVL(rsTmp!�ܸ�����, 0)
            If dbl���� = 0 Then dbl���� = 1
        Else
            lng����ID = Val(vsAdvice.TextMatrix(lng�к�, COL_����ID))
            lng��ҳID = Val(vsAdvice.TextMatrix(lng�к�, COL_��ҳID))
            lngִ�п���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))
            lng���˿���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_���˿���ID))
            dbl���� = Val(vsAdvice.TextMatrix(lng�к�, COL_����))
            If dbl���� = 0 Then dbl���� = 1
        End If
            
        '��ҩ���͸������õ�����ר����ִ�п���
        If InStr(",5,6,7,", rsInput!���ID) > 0 Or rsInput!���ID = "4" And NVL(rsInput!��������ID, 0) = 1 Then
            lngִ�п���ID = Get�շ�ִ�п���ID(lng����ID, lng��ҳID, rsInput!���ID, rsInput!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID, , , 2)
            '��¼�����Ƿ��������
            If rsInput!���ID = "4" Then
                .TextMatrix(lngRow, COLP_��������) = NVL(rsInput!��������ID, 0)
            End If
        End If
        If lngִ�п���ID <> 0 Then
            mrsDept.Filter = "ID=" & lngִ�п���ID
            If Not mrsDept.EOF Then
                .TextMatrix(lngRow, COLP_ִ�п���) = mrsDept!����
            End If
        End If
        .TextMatrix(lngRow, COLP_ִ�п���ID) = lngִ�п���ID
        If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, lng����ID, lng��ҳID, "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
        '����
        If InStr(",5,6,7,", rsInput!���ID) > 0 Then
            If NVL(rsInput!�Ƿ���ID, 0) = 0 Then
                dbl���� = NVL(rsInput!�ּ�ID, 0)
            Else 'δȷ���Ƽ�ҽ��ʱ,ҩƷ�޷�����۸�
                dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, dbl����, , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�) '��ȱʡ�Ƽ�����Ϊ1�����۵�λ����
            End If
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, gstrDecPrice)
        ElseIf rsInput!���ID = "4" And NVL(rsInput!��������ID, 0) = 1 And NVL(rsInput!�Ƿ���ID, 0) = 1 Then
            '�������õ�ʱ�����ĺ�ҩƷһ������
            dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, dbl����, , True, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, gstrDecPrice)
        Else
            If NVL(rsInput!�Ƿ���ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_����) = Format(NVL(rsInput!�ּ�ID, 0), gstrDecPrice)
            Else
                .TextMatrix(lngRow, COLP_����) = Format(NVL(rsInput!ȱʡ�۸�ID), gstrDecPrice)
                .Cell(flexcpData, lngRow, 0) = 1
                .Cell(flexcpData, lngRow, 1) = NVL(rsInput!ԭ��ID, 0)
                .Cell(flexcpData, lngRow, 2) = NVL(rsInput!�ּ�ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_��������) = NVL(rsInput!��������)
        .TextMatrix(lngRow, COLP_�̶�) = "0"
        .TextMatrix(lngRow, COLP_�շѷ�ʽ) = "������ȡ"
        
        '��������ָ�
        .Cell(flexcpData, lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ)
        .Cell(flexcpData, lngRow, COLP_����) = .TextMatrix(lngRow, COLP_����)
        .Cell(flexcpData, lngRow, COLP_����) = .TextMatrix(lngRow, COLP_����)
        .Cell(flexcpData, lngRow, COLP_ִ�п���) = .TextMatrix(lngRow, COLP_ִ�п���)
        
        '��¼������
        If lngҽ��ID <> 0 Then
            If lngԭ��ĿID = 0 Then
                '��ǰҽ���Ƿ��д��������������Ŀ�Ƿ����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And ����=1"
                If Not mrsPrice.EOF Then blnHaveSub = True
                .TextMatrix(lngRow, COLP_����) = IIF(blnHaveSub, "��", "")

                mrsPrice.AddNew '����
            Else '����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
            End If
            If lngԭ��ĿID = 0 Then
                If rsTmp Is Nothing Then
                    Set rsTmp = Sys.RowValue("����ҽ����¼", lngҽ��ID)
                End If
                mrsPrice!ҽ��ID = lngҽ��ID
                mrsPrice!���ID = IIF(Val(.TextMatrix(lngRow, COLP_���ID)) = 0, Null, Val(.TextMatrix(lngRow, COLP_���ID)))
                mrsPrice!������� = .TextMatrix(lngRow, COLP_�������)
                mrsPrice!������ĿID = Val(.TextMatrix(lngRow, COLP_������ĿID))
                mrsPrice!���� = IIF(blnHaveSub, 1, 0)
                
                mrsPrice!�걾��λ = rsTmp!�걾��λ
                mrsPrice!��鷽�� = rsTmp!��鷽��
                mrsPrice!ִ�б�� = NVL(rsTmp!ִ�б��, 0)
                mrsPrice!�������� = int��������
            End If
            mrsPrice!�շѷ�ʽ = 0
            mrsPrice!�շ���� = rsInput!���ID
            mrsPrice!�շ�ϸĿID = rsInput!ID
            If lngִ�п���ID <> 0 Then
                mrsPrice!ִ�п���ID = lngִ�п���ID
            Else
                mrsPrice!ִ�п���ID = Null
            End If
            mrsPrice!���� = NVL(rsInput!��������ID, 0)
            mrsPrice!��� = NVL(rsInput!�Ƿ���ID, 0)
            mrsPrice!���� = 1
            mrsPrice!���� = Val(.TextMatrix(lngRow, COLP_����))
            mrsPrice!�̶� = 0
            mrsPrice.Update
            Call SelectRow(vsAdvice.Row)
        End If
    End With
End Sub

Private Sub vsPrice_DblClick()
    Call vsPrice_KeyPress(32)
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsPrice
        If KeyCode = vbKeyF4 Then
            If CellEditable(.Row, .Col) And .Col = COLP_�Ƽ�ҽ�� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Editable And Val(.TextMatrix(.Row, COLP_�̶�)) = 0 Then
                If Val(.TextMatrix(.Row, COLP_ҽ��ID)) <> 0 And Val(.TextMatrix(.Row, COLP_�շ�ϸĿID)) <> 0 Then
                    'ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(.Row, COLP_ҽ��ID)) & " And ��������=" & Val(.TextMatrix(.Row, COLP_��������)) & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_�Ƽ�ҽ��) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If MsgBox("ȷ��Ҫɾ����ǰ�Ƽ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(.Row, COLP_ҽ��ID)) & " And ��������=" & Val(.TextMatrix(.Row, COLP_��������)) & " And �շ�ϸĿID=" & Val(.TextMatrix(.Row, COLP_�շ�ϸĿID))
                    mrsPrice.Delete
                    Call SelectRow(vsAdvice.Row)
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_�Ƽ�ҽ��
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsPrice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        Else
            If CellEditable(.Row, .Col) And (.Col = COLP_�շ���Ŀ Or .Col = COLP_ִ�п���) Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsPrice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str��ĿIDs As String, int�������� As Integer
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim strInput As String, strMatch As String
    Dim vPoint As PointAPI, strStock As String
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Col = COLP_�Ƽ�ҽ�� Then
                '����ʱ�س�
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '��ȻEnterNextCell����Ҫ�˳�
                    Call EnterNextCell(Row, Col)
                End If
            ElseIf Col = COLP_���� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�շ�����������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_���� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�շѵ���������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '��������뷶Χ
                strTmp = CheckScope(.Cell(flexcpData, Row, 1), .Cell(flexcpData, Row, 2), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, gstrDecPrice)
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                int�������� = Val(.TextMatrix(Row, COLP_��������))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call SelectRow(vsAdvice.Row)
                End If
                
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_�շ���Ŀ And .EditText <> "" Then
                '����ѡ�����е���Ŀ
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COLP_ҽ��ID)) = Val(.TextMatrix(Row, COLP_ҽ��ID)) _
                        And Val(.TextMatrix(Row, COLP_ҽ��ID)) <> 0 And i <> Row Then
                        str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                    End If
                Next
                str��ĿIDs = Mid(str��ĿIDs, 2)
                
                '��ͬ������ƥ�䷽ʽ
                strInput = UCase(.EditText)
                strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=[3] Or C.���� Like [2] And C.���� IN([3],3))"
                If IsNumeric(strInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=3)"
                ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.���� Like [2] And C.����=[3]"
                ElseIf zlCommFun.IsCharChinese(strInput) Then
                    strMatch = " And C.���� Like [2] And C.����=[3]"
                End If
                
                'ҩƷ���Ŀ��
                Call GetDefaultDeptPar(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��Ժ����ID)))
                If mlng��ҩ�� <> 0 Or mlng��ҩ�� <> 0 Or mlng��ҩ�� <> 0 Or mlng���ϲ��� <> 0 Then
                    strStock = _
                        "Select A.ҩƷID,Sum(Nvl(A.��������,0)) as ���" & _
                        " From ҩƷ��� A,�շ���ĿĿ¼ B" & _
                        " Where A.���� = 1 And (Nvl(A.����,0)=0 Or A.Ч�� Is Null Or A.Ч��>Trunc(Sysdate))" & _
                            " And A.�ⷿID=Decode(B.���,'5',[6],'6',[7],'7',[8],'4',[9],Null)" & _
                            " And A.ҩƷID=B.ID And B.��� IN('4','5','6','7')" & _
                        " Group by A.ҩƷID Having Sum(Nvl(A.��������,0))<>0"
                Else
                    strStock = "Select Null as ҩƷID,Null as ��� From Dual"
                End If
                If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��ҳID)), "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                
                strSQL = ""
                If Not DeptExist("���ϲ���", 2) Then strSQL = " And A.���<>'4'"
                strSQL = "Select * From (" & _
                    " Select A.ĩ��,A.ID,A.���,A.����,A.����,Decode(Instr('567',A.���ID),0,A.��λ,C.סԺ��λ) as ��λ,A.���,A.����," & _
                    " Decode(Nvl(A.�Ƿ���,0),1,Decode(Instr('567',A.���ID),0,Sum(Nvl(A.ԭ��,0))||'-'||Sum(Nvl(A.�ּ�,0))||'/'||A.��λ,'ʱ��')," & _
                    "   Decode(Instr('567',A.���ID),0,Sum(A.�ּ�)||'/'||A.��λ,LTrim(To_Char(Sum(A.�ּ�)*C.סԺ��װ,'999990.0000'))||'/'||C.סԺ��λ)) as �۸�," & _
                    " Decode(Instr('4567',A.���ID),0,NULL,1," & _
                    "   Decode(S.���,NULL,NULL,LTrim(To_Char(S.���,'999990.0000'))||A.��λ)," & _
                    "   Decode(S.���,NULL,NULL,LTrim(To_Char(S.���/Nvl(C.סԺ��װ,1),'999990.0000'))||C.סԺ��λ)) as ���,A.��������,N.���� as ҽ������,A.˵��," & _
                    " Sum(A.ԭ��) as ԭ��ID,Sum(A.�ּ�) as �ּ�ID,Sum(A.ȱʡ�۸�) as ȱʡ�۸�ID,A.�Ƿ��� as �Ƿ���ID,A.���ID,B.�������� as ��������ID,B.�������" & _
                    " From (" & _
                    " Select Distinct 1 as ĩ��,A.ID,a.ִ�п���,A.��� as ���ID,D.���� as ���,A.����,A.����," & _
                    " A.���㵥λ as ��λ,A.���,A.����,A.��������,A.˵��,B.ԭ��,B.�ּ�,B.ȱʡ�۸�,A.�Ƿ���" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ���� C,�շ���Ŀ��� D" & _
                    " Where A.ID=B.�շ�ϸĿID And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "11", "12", "13") & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And Instr([4],','||A.ID||',')=0", "") & _
                    " And A.ID=C.�շ�ϸĿID And A.���=D.���� And A.��� Not IN('J','1')" & strSQL & strMatch & _
                    " ) A,�������� B,ҩƷ��� C,����֧����Ŀ M,����֧������ N,(" & strStock & ") S" & _
                    " Where A.ID=B.����ID(+) And A.ID=C.ҩƷID(+) And A.ID=S.ҩƷID(+)" & _
                    " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[5]" & _
                    " And (Nvl(a.ִ�п���,0) <> 4 Or Exists (Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid = a.Id And (w.������Դ=2 or (w.������Դ is Null And Nvl(w.��������id,[10]) = [10]))))" & _
                    " And (a.���id not in ('4','5','6','7') Or Exists(Select 1 From �շ�ִ�п��� W Where w.�շ�ϸĿid=a.Id And Nvl(w.��������id,[10])=[10]))" & _
                    " Group by A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����,A.��������,C.סԺ��λ,C.סԺ��װ,S.���,N.����,A.˵��,A.�Ƿ���,A.���ID,B.��������,B.�������" & _
                    " ) Where Nvl(�������,0) = 0 Order by ���,����"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�շ���Ŀ", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", mint���� + 1, "," & str��ĿIDs & ",", mint����, mlng��ҩ��, mlng��ҩ��, mlng��ҩ��, mlng���ϲ���, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                If Not rsTmp Is Nothing Then
                    'ҽ��������
                    If CheckItemInsure(rsTmp, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��ҳID))) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                        .SetFocus: Exit Sub
                    End If
                    
                    lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                    int�������� = Val(.TextMatrix(Row, COLP_��������))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    Call SetItemInput(Row, rsTmp, lngҽ��ID, int��������, lngԭ��ĿID)
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õ��շ���Ŀ��", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            ElseIf Col = COLP_ִ�п��� And .EditText <> "" Then 'ִ�п���
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                If .TextMatrix(Row, COLP_�շ����) = "4" Then
                    '�������õ�����
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                        " And A.�շ�ϸĿID=[1] And (C.���� Like [3] Or C.���� Like [4] Or C.���� Like [4])" & _
                        " Order by B.�������,C.����"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                    'ҩƷ��ϵͳָ���Ĵ���ҩ������
                    If Not Check�ϰల��(True) Then
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(2,3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And (A.������Դ is NULL Or A.������Դ=2)" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    Else
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(2,3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                            " And (A.������Դ is NULL Or A.������Դ=2)" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                        Decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"), _
                        UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                End If
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                    .TextMatrix(Row, Col) = rsTmp!����
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    
                    '���¼�¼��
                    lngҽ��ID = Val(.TextMatrix(Row, COLP_ҽ��ID))
                    int�������� = Val(.TextMatrix(Row, COLP_��������))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ��������=" & int�������� & " And �շ�ϸĿID=" & lngԭ��ĿID
                        mrsPrice!ִ�п���ID = rsTmp!ID
                        mrsPrice.Update
                        Call SelectRow(vsAdvice.Row)
                    End If
                    
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            End If
        Else
            If Col = COLP_���� Or Col = COLP_���� Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlCommFun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Not CellEditable(NewRow, NewCol) Then
        vsPrice.ComboList = ""
        vsPrice.FocusRect = flexFocusLight
    Else
        vsPrice.FocusRect = flexFocusSolid
        If NewCol = COLP_�Ƽ�ҽ�� Then
            vsPrice.ComboList = vsPrice.ColData(NewCol)
        ElseIf NewCol = COLP_�շ���Ŀ Or NewCol = COLP_ִ�п��� Then
            vsPrice.ComboList = "..."
        Else
            vsPrice.ComboList = ""
        End If
    End If
    
    If NewRow <> OldRow Then
        '��ʾҩƷ���������ĵĿ��
        With vsPrice
            stbThis.Panels(2).Text = ""
            If Val(.TextMatrix(NewRow, COLP_�շ�ϸĿID)) <> 0 Then
                If InStr(",5,6,7,", .TextMatrix(NewRow, COLP_�շ����)) > 0 _
                    Or .TextMatrix(NewRow, COLP_�շ����) = "4" And Val(.TextMatrix(NewRow, COLP_��������)) = 1 Then
                    '����Ƽ�ֻ�������ʾ���������ҩ��ҩƷ��סԺ��λ����ҩ��ҩƷ���ۼ۵�λ
                    If InStr(GetInsidePrivs(pסԺҽ������), "��ʾҩƷ���") = 0 Then
                        If GetStock(Val(.TextMatrix(NewRow, COLP_�շ�ϸĿID)), Val(.TextMatrix(NewRow, COLP_ִ�п���ID))) > 0 Then
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "�п��"
                        Else
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "�޿��"
                        End If
                    Else
                        If InStr(",5,6,7,", .TextMatrix(NewRow, COLP_�������)) > 0 Then
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "���ÿ��:" & _
                                FormatEx(GetStock(Val(.TextMatrix(NewRow, COLP_�շ�ϸĿID)), Val(.TextMatrix(NewRow, COLP_ִ�п���ID))), 5) & .TextMatrix(NewRow, COLP_��λ)
                        Else
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "���ÿ��:" & _
                                FormatEx(GetStock(Val(.TextMatrix(NewRow, COLP_�շ�ϸĿID)), Val(.TextMatrix(NewRow, COLP_ִ�п���ID)), 0), 5) & .TextMatrix(NewRow, COLP_��λ)
                        End If
                    End If
                End If
            End If
        End With
        
        '��ʾҽ������
        stbThis.Panels(3).Text = Getҽ������(NewRow)
    End If
End Sub

Private Function Getҽ������(ByVal lngRow As Long) As String
'���ܣ���ȡָ���еķ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str���� As String
    
    With vsPrice
        If Val(.TextMatrix(lngRow, COLP_�շ�ϸĿID)) <> 0 Then
            strSQL = "Select N.���� From ����֧����Ŀ M,����֧������ N Where M.�շ�ϸĿID=[1] And M.����ID=N.ID And M.����=[2]"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COLP_�շ�ϸĿID)), mint����)
            If Not rsTmp.EOF Then str���� = NVL(rsTmp!����)
        End If
    End With
    Getҽ������ = IIF(str���� <> "", "ҽ������:" & str����, "")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Then Cancel = True: Exit Sub
    
    If Not CellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = COLP_���� Or Col = COLP_���� Or Col = COLP_ִ�п��� Then
        If vsPrice.TextMatrix(Row, COLP_�շ���Ŀ) = "" Then
            Cancel = True '������ȷ���շ���Ŀ
        End If
    End If
    
    If Col = COLP_���� Or Col = COLP_���� Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'���ܣ��жϼ۱��е�Ԫ���Ƿ���Ա༭
    CellEditable = vsPrice.Editable
    With vsPrice
        If lngCol = COLP_ִ�п��� Then
            '��ҩƷ������ҽ���ģ�ҩƷ�����ļƼ۵�ִ�п��ҿ����޸�
            If Not ((.TextMatrix(lngRow, COLP_�շ����) = "4" And Val(.TextMatrix(lngRow, COLP_��������)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_�շ����)) > 0) And InStr(",4,5,6,7,", .TextMatrix(lngRow, COLP_�������)) = 0) Then
                CellEditable = False
            End If
            If .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Or .TextMatrix(lngRow, COLP_�������) = "" Then
                CellEditable = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_�̶�)) <> 0 Then
            '�̶������н������޸ı��
            If Not (.Cell(flexcpData, lngRow, 0) = 1 And lngCol = COLP_����) Then
                CellEditable = False
            End If
        Else
            If lngCol = COLP_���� Then
                If .Cell(flexcpData, lngRow, 0) <> 1 Then CellEditable = False
            ElseIf lngCol <> COLP_�Ƽ�ҽ�� And lngCol <> COLP_���� And lngCol <> COLP_�շ���Ŀ Then
                CellEditable = False
            End If
        End If
    End With
End Function

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ���λ���۱�����һ����������ĵ�Ԫ��
    Dim i As Long, j As Long
    
    With vsPrice
        '��ǰ��Ԫ�����δ��������,���˳�
        If CellEditable(lngRow, lngCol) Then
            If lngCol = COLP_���� And Val(.TextMatrix(lngRow, lngCol)) = 0 Then
                Exit Sub
            ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                Exit Sub
            End If
        End If
        
        '����һ��Ԫ��ʼѭ������
        For i = lngRow To .Rows - 1
            For j = IIF(i = lngRow, lngCol + 1, COLP_�Ƽ�ҽ��) To .Cols - 1
                If CellEditable(i, j) Then Exit For
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            '��ǰ�����û���ҵ���һ���ɱ༭��Ԫ,�������Ƽ�ҽ��,������һ����
            If CStr(.ColData(COLP_�Ƽ�ҽ��)) <> "" Then
                '��ǰ��δ��������,��λ����������Ԫ
                If .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = "" Then
                    .Col = COLP_�Ƽ�ҽ��
                ElseIf .TextMatrix(lngRow, COLP_����) = "" Then
                    .Col = COLP_����
                ElseIf .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Then
                    .Col = COLP_�շ���Ŀ
                ElseIf .Cell(flexcpData, lngRow, 0) = 1 And Val(.TextMatrix(lngRow, COLP_����)) = 0 Then
                    .Col = COLP_����
                Else
                    .AddItem "", .Rows
                    .Row = .Rows - 1: .Col = COLP_�Ƽ�ҽ��
                    
                    'ȱʡѡ��Ƽ�ҽ��(�������)
                    Call ShowDefaultRow
                End If
            Else
                If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1 '���ɱ༭ʱ���ⶨһ��
            End If
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Function LoadPrice(ByVal lngRow As Long, Optional blnChange As Boolean) As Boolean
'���ܣ���ȡָ��ҽ���ļƼ�,�����ݵ�ǰ�������շ� ��ϵ���и���
'���أ�blnChange=�Ƿ���ݵ�ǰ�������շ� ��ϵ�����еļƼ����ݽ����˵���
    Dim rsMan As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim rsAdd As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim dblPrice As Double, strSubItem As String
    Dim lngִ�п���ID As Long
    Dim lng����ID As Long, blnLoad As Boolean
    Dim lngҽ��ID As Long, lng���ID As Long
    
    On Error GoTo errH
    
    With vsAdvice
        '�Ѿ���ȡ����,�����ظ���ȡ
        If .TextMatrix(lngRow, COL_ID) = "" Then LoadPrice = True: Exit Function
        If .RowData(lngRow) = 1 Then LoadPrice = True: Exit Function
        
        
        lngҽ��ID = Val(vsAdvice.TextMatrix(lngRow, COL_ID))
        lng���ID = Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
        If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(.TextMatrix(lngRow, COL_����ID)), Val(.TextMatrix(lngRow, COL_��ҳID)), "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
        
                            
        'ҩƷ�����ĵļƼ�(�����������ʾ������Ϊ�������,ҩƷ�̶�Ϊ1��ʵ��ҩƷ������ʾʱ����)
        'ҩƷȱʡ�̶�Ϊ�����Ƽ�,����ҽ��ʱָ����Ϊ�Ա�ҩ(Ժ��ִ��)�Ĳ���ȡ;ҩƷ������Ϊ����
        If .TextMatrix(lngRow, COL_�������) = "4" Then
            '���ļƼ�
            strSQL = _
                " Select A.ID,A.���ID,A.���,A.ҽ��״̬,A.�������,A.������ĿID,Null as �걾��λ,Null as ��鷽��,0 as ִ�б��," & _
                " C.��� as �շ����,A.�շ�ϸĿID,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,Nvl(X.����,D.ȱʡ�۸�),D.�ּ�) as ����," & _
                " 0 as ����,A.ִ�п���ID,B.��������,C.�Ƿ���,C.����ʱ��,0 as ��������,0 as �շѷ�ʽ" & _
                " From ����ҽ����¼ A,�������� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D,����ҽ���Ƽ� X" & _
                " Where Rownum=1 And A.ID=[1] And A.ID=X.ҽ��ID(+)" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "3", "4", "5") & _
                " And A.�շ�ϸĿID=B.����ID And A.�շ�ϸĿID=C.ID And Nvl(A.ִ������,0)<>5" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.������� IN(2,3) And D.�շ�ϸĿID=C.ID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '��,����ҩ:���ܰ������ҽ��,����1��סԺ��װ�ĵ���
            strSQL = _
                " Select A.ID,A.���ID,A.���,A.ҽ��״̬,A.�������,A.������ĿID,Null as �걾��λ,Null as ��鷽��,0 as ִ�б��," & _
                " C.��� as �շ����,C.ID as �շ�ϸĿID,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*B.סԺ��װ as ����," & _
                " 0 as ����,A.ִ�п���ID,0 as ��������,C.�Ƿ���,C.����ʱ��,0 as ��������,0 as �շѷ�ʽ" & _
                " From ����ҽ����¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.������ĿID=B.ҩ��ID And B.ҩƷID=C.ID And Nvl(A.ִ������,0)<>5" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "3", "4", "5") & _
                " And (A.�շ�ϸĿID is NULL Or A.�շ�ϸĿID=B.ҩƷID)" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.������� IN(2,3) And D.�շ�ϸĿID=C.ID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
        ElseIf .TextMatrix(lngRow, COL_����) = "1" Then
            '�в�ҩ:һ����Ӧ�й���¼����д���շ�ϸĿID
            strSQL = _
                " Select A.ID,A.���ID,A.���,A.ҽ��״̬,A.�������,A.������ĿID,Null as �걾��λ,Null as ��鷽��,0 as ִ�б��," & _
                " C.��� as �շ����,C.ID as �շ�ϸĿID,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*B.סԺ��װ as ����," & _
                " 0 as ����,A.ִ�п���ID,0 as ��������,C.�Ƿ���,C.����ʱ��,0 as ��������,0 as �շѷ�ʽ" & _
                " From ����ҽ����¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where A.�������='7' And A.���ID=[1]" & _
                " And A.�շ�ϸĿID=B.ҩƷID And A.�շ�ϸĿID=C.ID And C.������� IN(2,3)" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "3", "4", "5") & _
                " And D.�շ�ϸĿID=C.ID And Nvl(A.ִ������,0)<>5" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
        End If
        
        '��ȡ���мƼۣ���ҩƷ��ļƼ�,�������ҽ���Ƽ�
        blnLoad = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '��ҩ;��:һ����ҩ��ֻ��ȡһ��������
            If InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                If .TextMatrix(lngRow - 1, COL_���ID) = .TextMatrix(lngRow, COL_���ID) Then
                    blnLoad = False
                End If
            End If
        End If
        If blnLoad Then
            '��ҩ�ĸ�ҩ;������ҩ�䷽�ļ巨���÷�����鼰��λ����������������,������Ŀ
            '���Ƽ�,�ֹ��Ƽۣ�����,Ժ��ִ�У���ҽ������ȡ
            '��Union��ʽ������������
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select A.ID,A.���ID,A.���,A.ҽ��״̬,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��," & _
                "   B.��� as �շ����,A.�շ�ϸĿID,A.����,A.����,Nvl(A.����,0) as ����,A.ִ�п���ID," & _
                "   C.��������,B.�Ƿ���,B.����ʱ��,Nvl(A.��������,0) as ��������,Nvl(A.�շѷ�ʽ,0) as �շѷ�ʽ" & _
                " From (" & _
                " Select A.ID,A.���ID,A.���,A.ҽ��״̬,A.�������,A.������ĿID,A.�걾��λ,Decode(A.�������,'E',Decode(Z.��������,'4',Null,A.��鷽��),A.��鷽��) as ��鷽��,A.ִ�б��," & _
                "   B.�շ�ϸĿID,B.����,B.����,B.����,Nvl(B.ִ�п���ID,A.ִ�п���ID) as ִ�п���ID,B.��������,B.�շѷ�ʽ" & _
                " From ����ҽ����¼ A,����ҽ���Ƽ� B,������ĿĿ¼ Z" & _
                " Where Z.id=a.������ĿID And A.������� Not IN('4','5','6','7') And A.ID=B.ҽ��ID(+) And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5)" & _
                " And A.ID=[1]" & _
                " Union ALL" & _
                " Select A.ID,A.���ID,A.���,A.ҽ��״̬,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��," & _
                "   B.�շ�ϸĿID,B.����,B.����,B.����,Nvl(B.ִ�п���ID,A.ִ�п���ID) as ִ�п���ID,B.��������,B.�շѷ�ʽ" & _
                " From ����ҽ����¼ A,����ҽ���Ƽ� B" & _
                " Where A.������� Not IN('4','5','6','7') And A.ID=B.ҽ��ID(+) And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5)" & _
                " And A.ID=[2]" & _
                " Union ALL" & _
                " Select A.ID,A.���ID,A.���,A.ҽ��״̬,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��," & _
                "   B.�շ�ϸĿID,B.����,B.����,B.����,Nvl(B.ִ�п���ID,A.ִ�п���ID) as ִ�п���ID,B.��������,B.�շѷ�ʽ" & _
                " From ����ҽ����¼ A,����ҽ���Ƽ� B" & _
                " Where A.������� Not IN('4','5','6','7') And A.ID=B.ҽ��ID(+) And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5)" & _
                " And A.���ID=[1]" & _
                " ) A,�շ���ĿĿ¼ B,�������� C" & _
                " Where A.�շ�ϸĿID=B.ID(+) And A.�շ�ϸĿID=C.����ID(+)" & _
                " Order by ���,��������,����"
        End If
        Set rsMan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_���ID)), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
        
        '�����շ� ��ϵ���շ����������ж����Ƿ�仯
        '��ͬ�Ŀ��ҿ����в�ͬ���շѶ��գ������������У������ִ�п��ҵ�����
        strSQL = "Select * From (Select C.������ĿID,C.�շ���ĿID,C.�շ�����,C.���ж���,C.������Ŀ," & _
            " Nvl(C.��鲿λ,'None') as ��鲿λ,Nvl(C.��鷽��,'None') as ��鷽��," & _
            " Nvl(C.��������,0) as ��������,Nvl(C.�շѷ�ʽ,0) as �շѷ�ʽ,C.���ÿ���id" & _
            " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
            " From ����ҽ����¼ A,����ҽ���Ƽ� B,�����շѹ�ϵ C" & _
            " Where A.ID=B.ҽ��ID And A.������ĿID+0=C.������ĿID And B.�շ�ϸĿID+0=C.�շ���ĿID" & _
            " And (C.���ÿ���ID is Null or C.���ÿ���ID = Nvl(A.ִ�п���id,[3]) And C.������Դ = 2)" & _
            " And (A.���ID is Null And A.ִ�б�� IN(1,2) And C.��������=1" & _
            "       Or A.�걾��λ=C.��鲿λ And A.��鷽��=C.��鷽�� And Nvl(C.��������,0)=0" & _
            "       Or (A.��鷽�� is Null or a.������� = 'E' And Exists(Select 1 From ������ĿĿ¼ Z Where Z.id=a.������ĿID And Z.��������='4')) And Nvl(C.��������,0)=0 And C.��鲿λ is Null And C.��鷽�� is Null)" & _
            " And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0)<>5 And (A.ID=[1]" & IIF(lng���ID <> 0, " Or A.ID=[2]", "") & " Or A.���ID=[1])" & _
            " ) Where Nvl(���ÿ���id, 0) = Top"
        Set rsCur = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ID, mlng����ID)
        
        '����ҩƷ�����еļƼ�
        For i = 1 To rsMan.RecordCount
            strSubItem = ""
            
            mrsPrice.AddNew '��δ����Ƽ۹�ϵ��ҲҪ��������ȷ���ɼƼ�ҽ��(�ü�¼����)
            mrsPrice!ҽ��ID = rsMan!ID
            mrsPrice!���ID = rsMan!���ID
            mrsPrice!������� = rsMan!�������
            mrsPrice!������ĿID = rsMan!������ĿID
            mrsPrice!�̶� = IIF(InStr(",4,5,6,7,", rsMan!�������) > 0, 1, 0)
            
            '�����Ŀ����չ
            mrsPrice!�걾��λ = rsMan!�걾��λ
            mrsPrice!��鷽�� = rsMan!��鷽��
            mrsPrice!ִ�б�� = NVL(rsMan!ִ�б��, 0)
            mrsPrice!�������� = NVL(rsMan!��������, 0)
            mrsPrice!�շѷ�ʽ = NVL(rsMan!�շѷ�ʽ, 0)
            
            '������ҽ���Ƽ�ʱ,����ԭ������,�����ѳ�������Ŀ,����δ����(�Ա���������)
            If Not IsNull(rsMan!�շ�ϸĿID) _
                And Format(NVL(rsMan!����ʱ��, "3000-01-01"), "yyyy-MM-dd") = "3000-01-01" Then
                mrsPrice!�շ���� = rsMan!�շ����
                mrsPrice!�շ�ϸĿID = rsMan!�շ�ϸĿID
                mrsPrice!ִ�п���ID = rsMan!ִ�п���ID
                mrsPrice!���� = NVL(rsMan!��������, 0)
                mrsPrice!��� = NVL(rsMan!�Ƿ���, 0)
                mrsPrice!���� = rsMan!����
                
                'ҩƷ(��������ʾ)�����Ϊʱ�ۣ���ʾʱ���㣻�������ȡ�����¼۸�
                '���ģ����Ϊ��������ʱ�ۣ���ʾʱ���㣻����ȡ���ۻ���ǰ����(�����)
                '��ҩƷ�����Ϊ���,��ȡ��ǰ����(�����)����������ȡ���¼۸�
                mrsPrice!���� = rsMan!����
                mrsPrice!���� = NVL(rsMan!����, 0)
                        
                '�����շ� ��ϵ���շ����������ж����Ƿ�仯
                If InStr(",4,5,6,7,", rsMan!�������) = 0 Then '������ҩƷ������ҽ����ҩƷ�Ƽ�
                    If rsMan!������� = "D" Then
                        rsCur.Filter = "������ĿID=" & rsMan!������ĿID & " And �շ���ĿID=" & rsMan!�շ�ϸĿID & _
                            " And ��鲿λ='" & NVL(rsMan!�걾��λ, "None") & "' And ��鷽��='" & NVL(rsMan!��鷽��, "None") & "'" & _
                            " And ��������=" & NVL(rsMan!��������, 0)
                    Else
                        rsCur.Filter = "������ĿID=" & rsMan!������ĿID & " And �շ���ĿID=" & rsMan!�շ�ϸĿID & " And ��鲿λ='None' And ��鷽��='None' And ��������=" & NVL(rsMan!��������, 0)
                    End If
                    If Not rsCur.EOF Then
                        If NVL(rsCur!���ж���, 0) <> 0 And NVL(rsMan!����, 0) <> NVL(rsCur!�շ�����, 0) Then
                            mrsPrice!���� = rsCur!�շ����� '����˹��ж��ղ�ȡ�����õ�����
                            blnChange = True
                        End If
                        mrsPrice!���� = NVL(rsCur!������Ŀ, 0)
                        mrsPrice!�̶� = NVL(rsCur!���ж���, 0)
                    End If
                    '�۸�ȡ���µ�(�Ǳ��)
                    dblPrice = CalcPrice(rsMan!�շ�ϸĿID, , , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                    If dblPrice <> 0 Then mrsPrice!���� = Format(dblPrice, gstrDecPrice)
                End If
            End If
            mrsPrice.Update
            
            '���ڴ�����Ŀ��Ҫ�Ƽ�ҽ��
            If mrsPrice!���� = 1 Then
                If InStr(strSubItem & ";", ";" & mrsPrice!ҽ��ID & "," & mrsPrice!�������� & ";") = 0 Then
                    strSubItem = strSubItem & ";" & mrsPrice!ҽ��ID & "," & mrsPrice!��������
                End If
            End If
            
            '�����շ� ��ϵ�������˵Ķ���(��δУ��֮ǰ,����ҽ���Ƽ�û������,��ʱҲ��������ӵ�)
            If InStr(",1,2,", NVL(rsMan!ҽ��״̬, 0)) > 0 And InStr(",4,5,6,7,", rsMan!�������) = 0 Then '������ҩƷ������ҽ����ҩƷ�Ƽ�
                lngҽ��ID = rsMan!ID
                blnLoad = False: rsMan.MoveNext
                If rsMan.EOF Then
                    blnLoad = True
                ElseIf rsMan!ID <> lngҽ��ID Then
                    blnLoad = True
                End If
                rsMan.MovePrevious
                If blnLoad Then
                    lng����ID = 0 '�����Թܷ���,ֻ��ȡ�Թܶ�Ӧ�����ķ���
                    If .TextMatrix(lngRow, COL_�Թܱ���) <> "" Then
                        lng����ID = GetTubeMaterial(.TextMatrix(lngRow, COL_�Թܱ���))
                    End If
                    strSQL = _
                        "Select ������ĿID,�շ����,�շ���ĿID,�շ�����,���ж���,������Ŀ," & _
                        "   ���˿���ID,ִ�п���ID,��������,�Ƿ���,�걾��λ,��鷽��,ִ�б��,��������,�շѷ�ʽ,Sum(����) as ���� From (" & _
                        " Select c.������ĿID,f.��� as �շ����,c.�շ���ĿID,c.�շ�����,c.���ж���,Nvl(c.������Ŀ,0) as ������Ŀ," & _
                        " B.���˿���ID,B.ִ�п���ID,E.��������,f.�Ƿ���,Decode(Nvl(f.�Ƿ���,0),1,D.ȱʡ�۸�,D.�ּ�) as ����," & _
                        " B.�걾��λ,B.��鷽��,B.ִ�б��,Nvl(c.��������,0) as ��������,Nvl(c.�շѷ�ʽ,0) as �շѷ�ʽ,c.���ÿ���id" & _
                        " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                        " From �����շѹ�ϵ C,����ҽ����¼ B,�շ���ĿĿ¼ F,�շѼ�Ŀ D,�������� E" & _
                        " Where c.������ĿID+0=B.������ĿID And B.ID=[1] And f.ID=E.����ID(+)" & _
                        GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "F", "D", "4", "5", "6") & _
                        " And (B.���ID is Null And B.ִ�б�� IN(1,2) And c.��������=1" & _
                        "       Or B.�걾��λ=c.��鲿λ And B.��鷽��=c.��鷽�� And Nvl(c.��������,0)=0" & _
                        "       Or (B.��鷽�� is Null or b.������� = 'E' And Exists(Select 1 From ������ĿĿ¼ Z Where Z.id=b.������ĿID And Z.��������='4')) And Nvl(c.��������,0)=0 And c.��鲿λ is Null And c.��鷽�� is Null)" & _
                        " And c.�շ���ĿID Not IN(Select �շ�ϸĿID From ����ҽ���Ƽ� Where ҽ��ID=[1])" & _
                        " And c.�շ���ĿID=f.ID And c.�շ���ĿID=D.�շ�ϸĿID And f.������� IN(2,3)" & _
                        " And (c.�շѷ�ʽ=1 And f.���='4' And c.�շ���ĿID=[2] Or Not(c.�շѷ�ʽ=1 And f.���='4' And [2]<>0))" & _
                        " And (f.����ʱ�� is NULL Or f.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " And (f.վ��='" & gstrNodeNo & "' Or f.վ�� is Null) And Sysdate Between D.ִ������ and Nvl(D.��ֹ����,Sysdate)" & _
                        " And (C.���ÿ���ID is Null or C.���ÿ���ID = Nvl(B.ִ�п���id,[3]) And c.������Դ = 2)" & _
                        " ) Where Nvl(���ÿ���id, 0) = Top" & _
                        " Group by ������ĿID,�շ����,�շ���ĿID,�շ�����,���ж���,������Ŀ," & _
                        "   ���˿���ID,ִ�п���ID,��������,�Ƿ���,�걾��λ,��鷽��,ִ�б��,��������,�շѷ�ʽ" & _
                        " Order by ��������,������Ŀ"
                    Set rsAdd = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMan!ID), lng����ID, mlng����ID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                    If Not rsAdd.EOF Then
                        For j = 1 To rsAdd.RecordCount
                            '��ҩ���͸������õ�����ר����ִ�п���
                            lngִ�п���ID = NVL(rsAdd!ִ�п���ID, 0)
                            If InStr(",5,6,7,", rsAdd!�շ����) > 0 Or rsAdd!�շ���� = "4" And NVL(rsAdd!��������, 0) = 1 Then
                                lngִ�п���ID = Get�շ�ִ�п���ID(Val(.TextMatrix(lngRow, COL_����ID)), Val(.TextMatrix(lngRow, COL_��ҳID)), rsAdd!�շ����, rsAdd!�շ���ĿID, 4, NVL(rsAdd!���˿���id, 0), 0, 2, lngִ�п���ID, , , 2)
                            End If
                            
                            mrsPrice.AddNew
                            mrsPrice!ҽ��ID = rsMan!ID
                            mrsPrice!���ID = rsMan!���ID
                            mrsPrice!������� = rsMan!�������
                            mrsPrice!������ĿID = rsMan!������ĿID
                            mrsPrice!�շ���� = rsAdd!�շ����
                            mrsPrice!�շ�ϸĿID = rsAdd!�շ���ĿID
                            If lngִ�п���ID <> 0 Then
                                mrsPrice!ִ�п���ID = lngִ�п���ID
                            Else
                                mrsPrice!ִ�п���ID = Null
                            End If
                            mrsPrice!���� = NVL(rsAdd!��������, 0)
                            mrsPrice!��� = NVL(rsAdd!�Ƿ���, 0)
                            mrsPrice!���� = rsAdd!�շ�����
                            mrsPrice!���� = rsAdd!����
                            mrsPrice!���� = NVL(rsAdd!������Ŀ, 0)
                            mrsPrice!�̶� = NVL(rsAdd!���ж���, 0)
                            
                            '�����Ŀ����չ
                            mrsPrice!�걾��λ = rsAdd!�걾��λ
                            mrsPrice!��鷽�� = rsAdd!��鷽��
                            mrsPrice!ִ�б�� = NVL(rsAdd!ִ�б��, 0)
                            mrsPrice!�������� = NVL(rsAdd!��������, 0)
                            mrsPrice!�շѷ�ʽ = NVL(rsAdd!�շѷ�ʽ, 0)
                            
                            mrsPrice.Update
                            
                            '���ڴ�����Ŀ��Ҫ�Ƽ�ҽ��
                            If mrsPrice!���� = 1 Then
                                If InStr(strSubItem & ";", ";" & mrsPrice!ҽ��ID & "," & mrsPrice!�������� & ";") = 0 Then
                                    strSubItem = strSubItem & ";" & mrsPrice!ҽ��ID & "," & mrsPrice!��������
                                End If
                            End If
                            If NVL(mrsPrice!����, 0) <> 0 Then blnChange = True '�б仯
                            
                            rsAdd.MoveNext
                        Next
                    End If
                End If
            End If
            
            '�Դ��ڴ���ļƼ۽��д�����ֻ֤��һ������
            If strSubItem <> "" Then
                If AdjustSubPrice(Mid(strSubItem, 2)) Then blnChange = True
            End If
            rsMan.MoveNext
        Next
                
        .RowData(lngRow) = 1
    End With
    
    LoadPrice = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AdjustSubPrice(ByVal strSubItem As String) As Boolean
'���ܣ��Դ��ڴ���ļƼ۽��д�����ֻ֤��һ������
'������strSubitem=��������Ƽ���Ŀ��ҽ����Ŀ��"ҽ��ID,��������;..."
'���أ��Ƽ������Ƿ��б仯
    Dim rsTmp As ADODB.Recordset
    Dim arrAdvice As Variant, blnChange As Boolean
    Dim intCount As Integer, i As Integer
    Dim strSQL As String
    
    arrAdvice = Split(strSubItem, ";")
    For i = 0 To UBound(arrAdvice)
        intCount = 0
        strSQL = _
            "Select Sum(Decode(������Ŀ,1,1,0)) as ������," & _
            " Max(Decode(������Ŀ,1,NULL,�շ���ĿID)) as ����ID" & _
            " From (Select C.������Ŀ,C.�շ���ĿID,C.���ÿ���id" & _
            " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
            " From ����ҽ����¼ A,�����շѹ�ϵ C" & _
            " Where A.ID=[1] And A.������ĿID+0=C.������ĿID And Nvl(C.��������,0)=[2]" & _
            "       And (A.���ID is Null And A.ִ�б�� IN(1,2) And C.��������=1" & _
            "       Or A.�걾��λ=C.��鲿λ And A.��鷽��=C.��鷽�� And Nvl(C.��������,0)=0" & _
            "       Or (A.��鷽�� is Null or a.������� = 'E' And Exists(Select 1 From ������ĿĿ¼ Z Where Z.id=a.������ĿID And Z.��������='4')) And Nvl(C.��������,0)=0 And C.��鲿λ is Null And C.��鷽�� is Null)" & _
            "       And (C.���ÿ���ID is Null or C.���ÿ���ID = Nvl(A.ִ�п���id,[3]) And C.������Դ = 2)" & _
            ") Where Nvl(���ÿ���id, 0) = Top"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Split(arrAdvice(i), ",")(0)), Val(Split(arrAdvice(i), ",")(1)), mlng����ID)
        If Not rsTmp.EOF Then intCount = NVL(rsTmp!������, 0)
        If intCount = 0 Then
            '������мƼ�û�д�����Ŀ����ȡ�����д�������
            mrsPrice.Filter = "ҽ��ID=" & Val(Split(arrAdvice(i), ",")(0)) & " And ��������=" & Val(Split(arrAdvice(i), ",")(1))
            Do While Not mrsPrice.EOF
                If mrsPrice!���� = 1 Then
                    mrsPrice!���� = 0
                    mrsPrice.Update
                    blnChange = True
                End If
                mrsPrice.MoveNext
            Loop
        Else
            '������ڴ�����Ŀ�������������ȫ������Ϊ����
            mrsPrice.Filter = "ҽ��ID=" & Val(Split(arrAdvice(i), ",")(0)) & " And ��������=" & Val(Split(arrAdvice(i), ",")(1))
            Do While Not mrsPrice.EOF
                If mrsPrice!�շ�ϸĿID = Val(NVL(rsTmp!����ID, 0)) Then 'Ϊʲôһ��Ҫ��Val?
                    If mrsPrice!���� = 1 Then
                        mrsPrice!���� = 0 '����϶�����ֻ��һ��
                        mrsPrice.Update
                        blnChange = True
                    End If
                Else
                    If mrsPrice!���� = 0 Then
                        mrsPrice!���� = 1
                        mrsPrice.Update
                        blnChange = True
                    End If
                End If
                mrsPrice.MoveNext
            Loop
        End If
    Next
    
    AdjustSubPrice = blnChange
End Function

Private Sub AppendPriceItem()
'���ܣ������޶�Ӧ�Ƽ�������Ŀ�ļ�¼
    Dim arrPrice As Variant, strPrice As String, i As Long
    Dim lng���ID As Long, str��� As String
    Dim lng��Ŀid As Long, str��λ As String
    Dim str���� As String, intִ�б�� As Integer

    mrsPrice.Filter = 0
    Do While Not mrsPrice.EOF
        If mrsPrice!������� = "D" And IsNull(mrsPrice!���ID) Then '��鴲�Ի�������һ�ּ������
            '����Ӧ�е�
            If InStr(strPrice, mrsPrice!ҽ��ID & "_") = 0 Then
                If NVL(mrsPrice!ִ�б��, 0) <> 0 Then
                    '��Ϊ���Ի�����ִ��ʱ���ſ������ü��ռƼ�
                    strPrice = strPrice & "," & mrsPrice!ҽ��ID & "_0," & mrsPrice!ҽ��ID & "_1"
                Else
                    strPrice = strPrice & "," & mrsPrice!ҽ��ID & "_0"
                End If
            End If
            'ȥ�����е�
            If InStr(strPrice, "," & mrsPrice!ҽ��ID & "_" & NVL(mrsPrice!��������, 0)) > 0 Then
                strPrice = Replace(strPrice, "," & mrsPrice!ҽ��ID & "_" & NVL(mrsPrice!��������, 0), "")
            End If
        End If
        mrsPrice.MoveNext
    Loop
    
    'ʣ��ľ���û�е�
    If strPrice <> "" Then
        arrPrice = Split(Mid(strPrice, 2), ",")
        For i = 0 To UBound(arrPrice)
            mrsPrice.Filter = "ҽ��ID=" & Split(arrPrice(i), "_")(0) '����ҽ�����ܶ�Ӧ�ж��ַ������ʵļƼۼ�¼
            If Not mrsPrice.EOF Then
                lng���ID = NVL(mrsPrice!���ID, 0)
                str��� = mrsPrice!�������
                lng��Ŀid = mrsPrice!������ĿID
                str��λ = NVL(mrsPrice!�걾��λ)
                str���� = NVL(mrsPrice!��鷽��)
                intִ�б�� = NVL(mrsPrice!ִ�б��, 0)
                
                mrsPrice.AddNew
                mrsPrice!ҽ��ID = Val(Split(arrPrice(i), "_")(0))
                If lng���ID <> 0 Then mrsPrice!���ID = lng���ID
                mrsPrice!������� = str���
                mrsPrice!������ĿID = lng��Ŀid
                If str��λ <> "" Then mrsPrice!�걾��λ = str��λ
                If str���� <> "" Then mrsPrice!��鷽�� = str����
                mrsPrice!ִ�б�� = intִ�б��
                mrsPrice!�������� = Val(Split(arrPrice(i), "_")(1))
                mrsPrice!�̶� = 0
                mrsPrice.Update
            End If
        Next
    End If
End Sub

Private Sub ShowPrice(ByVal lngRow As Long)
'���ܣ���ʾ��ǰҽ���еļƼ�����(�������ҽ���ļƼ���Ŀ),ͬʱ����һЩ�༭����
    Dim rs������Ŀ As New ADODB.Recordset
    Dim rs�շ�ϸĿ As New ADODB.Recordset
    Dim str������ĿIDs As String, str�շ�ϸĿIDs As String
    Dim strSQL As String, strAllow As String
    Dim str�Ƽ�ҽ�� As String, i As Long, j As Long
    Dim blnNoFirst As Boolean, lngBegin As Long
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    Dim lngComboData As Long, strCombo As String
    Dim strPriceType As String
    
    On Error GoTo errH
    
    With vsPrice
        .Redraw = False
        '�����Ŀ���
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        .Editable = flexEDNone
        
        '�Ƿ�һ����ҩ�еķǵ�һҩƷ��
        If RowInһ����ҩ(lngRow, lngBegin, 0) Then
            If lngRow > lngBegin Then blnNoFirst = True
        End If
        
        If Val(vsAdvice.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            If blnNoFirst Then
                'һ����ҩʱ����һ����ʾ��ҩ;���ļƼ�
                mrsPrice.Filter = "ҽ��ID=" & vsAdvice.TextMatrix(lngRow, COL_ID)
            Else
                mrsPrice.Filter = "ҽ��ID=" & vsAdvice.TextMatrix(lngRow, COL_ID) & _
                    " Or ҽ��ID=" & Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
            End If
        Else
            mrsPrice.Filter = "ҽ��ID=" & vsAdvice.TextMatrix(lngRow, COL_ID) & _
                " Or ���ID=" & vsAdvice.TextMatrix(lngRow, COL_ID)
        End If
        
        If Not mrsPrice.EOF Then
'            If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
'                mrsPrice.Sort = "�������" 'һ����ҩʱ��ʾ˳��Ҫ��ҩƷ��ǰ
'            Else
'                mrsPrice.Sort = ""
'            End If
            If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(vsAdvice.TextMatrix(lngRow, COL_����ID)), Val(vsAdvice.TextMatrix(lngRow, COL_��ҳID)), "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                        
            '��ȡ������Ŀ,�շ�ϸĿ,�۸���Ϣ
            For i = 1 To mrsPrice.RecordCount
                str������ĿIDs = str������ĿIDs & "," & mrsPrice!������ĿID
                If Not IsNull(mrsPrice!�շ�ϸĿID) Then
                    str�շ�ϸĿIDs = str�շ�ϸĿIDs & "," & mrsPrice!�շ�ϸĿID
                End If
                
                '�ռ��������üƼ۵�ҽ��
                If Not IsNull(mrsPrice!�շ�ϸĿID) Then
                    lngComboData = mrsPrice!ҽ��ID
                    If NVL(mrsPrice!��������, 0) <> 0 Then '������ʾǰ�渽����һλ���շ�������
                        lngComboData = -1 * Val(mrsPrice!�������� & lngComboData)
                    End If
                    '���:ҽ��ID_�Ƿ�ȫΪ�̶�
                    If InStr(strAllow, "," & lngComboData & "_") = 0 Then
                        strAllow = strAllow & "," & lngComboData & "_" & mrsPrice!�̶�
                    ElseIf mrsPrice!�̶� = 0 Then
                        strAllow = Replace(strAllow, "," & lngComboData & "_1", "," & lngComboData & "_0")
                    End If
                End If
                
                mrsPrice.MoveNext
            Next
            str������ĿIDs = Mid(str������ĿIDs, 2)
            str�շ�ϸĿIDs = Mid(str�շ�ϸĿIDs, 2)
                        
            strSQL = "Select /*+ Rule*/ A.ID,B.���� as �������,A.����" & _
                " From ������ĿĿ¼ A,������Ŀ��� B" & _
                " Where A.���=B.���� And A.ID IN(Select Column_Value From Table(f_Num2list([1])))"
            Set rs������Ŀ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str������ĿIDs)
            
            '��ȡ�Ƿ��ۼ���۷�Χ����Ŀ��Ϣ
            If str�շ�ϸĿIDs <> "" Then
                strSQL = _
                    " Select A.ID,C.���� as �������,A.����,A.����,A.���," & _
                    " A.����,A.���㵥λ,Nvl(D.סԺ��λ,A.���㵥λ) as סԺ��λ," & _
                    " A.��������,A.�Ƿ���,Nvl(D.סԺ��װ,1) as סԺ��װ,A.���" & _
                    " From �շ���ĿĿ¼ A,�շ���Ŀ��� C,ҩƷ��� D" & _
                    " Where A.���=C.���� And A.ID=D.ҩƷID" & _
                    " And A.��� IN('5','6','7') And A.ID IN(Select Column_Value From Table(f_Num2list([1])))"
                '������
                strSQL = strSQL & " Union ALL " & _
                    " Select A.ID,C.���� as �������,A.����,A.����,A.���,A.����," & _
                    " A.���㵥λ,NULL as סԺ��λ,A.��������,A.�Ƿ���,-Null as סԺ��װ,A.���" & _
                    " From �շ���ĿĿ¼ A,�շ���Ŀ��� C" & _
                    " Where A.���=C.���� And A.��� Not IN('5','6','7')" & _
                    " And A.ID IN(Select Column_Value From Table(f_Num2list([1])))"
                
                strSQL = _
                    " Select A.ID,A.�������,A.����,A.����,A.���,A.����,A.���㵥λ,A.סԺ��λ,A.��������," & _
                    " A.�Ƿ���,A.סԺ��װ,Sum(B.ԭ��) as ԭ��,Sum(B.�ּ�) as �ּ�,Sum(B.ȱʡ�۸�) as ȱʡ�۸�" & _
                    " From (" & strSQL & ") A,�շѼ�Ŀ B Where A.ID=B.�շ�ϸĿID" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "3", "4", "5") & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " Group by A.ID,A.�������,A.����,A.����,A.���,A.����,A.���㵥λ,A.סԺ��װ,A.��������,A.�Ƿ���,A.סԺ��λ"

                strSQL = _
                    " Select /*+ Rule*/ A.ID,A.�������,A.����,Nvl(B.����,A.����) as ����,A.���,A.����," & _
                    " A.���㵥λ,A.סԺ��λ,A.��������,A.�Ƿ���,A.ԭ��,A.�ּ�,A.ȱʡ�۸�,A.סԺ��װ" & _
                    " From (" & strSQL & ") A,�շ���Ŀ���� B" & _
                    " Where A.ID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[2]"
                Set rs�շ�ϸĿ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�շ�ϸĿIDs, IIF(gbytҩƷ������ʾ = 0, 1, 3), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�) 'In
            End If
                        
            'ȷ����ʾ����
            If str�շ�ϸĿIDs <> "" Then
                .Rows = .FixedRows + UBound(Split(str�շ�ϸĿIDs, ",")) + 1
            End If
                                    
            '��ʾÿ������
            j = .FixedRows
            mrsPrice.MoveFirst
            For i = 1 To mrsPrice.RecordCount
                'ȷ���Ƽ�ҽ������
                rs������Ŀ.Filter = "ID=" & mrsPrice!������ĿID
                If mrsPrice!������� = "4" Then
                    str�Ƽ�ҽ�� = "��������-" & rs������Ŀ!����
                ElseIf InStr(",5,6,7,", mrsPrice!�������) > 0 Then
                    str�Ƽ�ҽ�� = "ҩƷҽ��-" & rs������Ŀ!����
                ElseIf mrsPrice!������� = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                    str�Ƽ�ҽ�� = "��ҩ;��-" & rs������Ŀ!����
                ElseIf mrsPrice!������� = "E" And vsAdvice.TextMatrix(lngRow, COL_�������) = "K" Then
                    str�Ƽ�ҽ�� = "��Ѫ;��-" & rs������Ŀ!����
                ElseIf mrsPrice!������� = "E" And InStr(",1,2,", Val(vsAdvice.TextMatrix(lngRow, COL_����))) > 0 Then
                    If vsAdvice.TextMatrix(lngRow, COL_����) = "2" Then
                        str�Ƽ�ҽ�� = "�ɼ�����-" & rs������Ŀ!����
                    ElseIf Not IsNull(mrsPrice!���ID) Then
                        str�Ƽ�ҽ�� = "��ҩ�巨-" & rs������Ŀ!����
                    Else
                        str�Ƽ�ҽ�� = "��ҩ�÷�-" & rs������Ŀ!����
                    End If
                ElseIf Not IsNull(mrsPrice!���ID) Then
                    If mrsPrice!������� = "C" Then
                        str�Ƽ�ҽ�� = "������Ŀ-" & rs������Ŀ!����
                    ElseIf mrsPrice!������� = "D" Then
                        '��λ������
                        str�Ƽ�ҽ�� = "��鲿λ-" & NVL(mrsPrice!�걾��λ) & "(" & NVL(mrsPrice!��鷽��) & ")"
                    ElseIf mrsPrice!������� = "F" Then
                        str�Ƽ�ҽ�� = "��������-" & rs������Ŀ!����
                    ElseIf mrsPrice!������� = "G" Then
                        str�Ƽ�ҽ�� = "������Ŀ-" & rs������Ŀ!����
                    End If
                Else
                    If NVL(mrsPrice!��������, 0) = 1 Then
                        '���Ի����м��շ���
                        str�Ƽ�ҽ�� = rs������Ŀ!������� & "ҽ��-" & rs������Ŀ!���� & "(" & Decode(NVL(mrsPrice!ִ�б��, 0), 1, "����", 2, "����", "") & "����)"
                    Else
                        str�Ƽ�ҽ�� = rs������Ŀ!������� & "ҽ��-" & rs������Ŀ!����
                    End If
                End If
                
                '����ѡ��ļƼ�ҽ��
                If InStr(",4,5,6,7,", mrsPrice!�������) = 0 Then
                    lngComboData = mrsPrice!ҽ��ID
                    If NVL(mrsPrice!��������, 0) <> 0 Then
                        lngComboData = -1 * Val(mrsPrice!�������� & lngComboData)
                    End If
                    '����:û�������κ��շ���Ŀ Or ���ڷǹ̶����շ���Ŀ(����ȫ���̶�)
                    If InStr(strAllow, "," & lngComboData & "_") = 0 _
                        Or InStr(strAllow, "," & lngComboData & "_0") > 0 Then
                        If InStr(strCombo, "|#" & lngComboData & ";" & str�Ƽ�ҽ��) = 0 Then
                            strCombo = strCombo & "|#" & lngComboData & ";" & str�Ƽ�ҽ��
                        End If
                    End If
                End If
                
                '��δ�����շѹ�ϵ�Ĳ���ʾ,������ѡ��
                If Not IsNull(mrsPrice!�շ�ϸĿID) Then
                    rs�շ�ϸĿ.Filter = "ID=" & mrsPrice!�շ�ϸĿID
                    
                    '��ʾ�Ƽ۵�ҽ������
                    .TextMatrix(j, COLP_�Ƽ�ҽ��) = str�Ƽ�ҽ��
                    .TextMatrix(j, COLP_ҽ��ID) = mrsPrice!ҽ��ID
                    .TextMatrix(j, COLP_��������) = NVL(mrsPrice!��������, 0)
                    .TextMatrix(j, COLP_�շѷ�ʽ) = getChargeMode(Val(NVL(mrsPrice!�շѷ�ʽ, 0)))
                        .Cell(flexcpData, j, COLP_�շѷ�ʽ) = Val(NVL(mrsPrice!�շѷ�ʽ, 0))
                    .TextMatrix(j, COLP_���ID) = NVL(mrsPrice!���ID)
                    .TextMatrix(j, COLP_�������) = mrsPrice!�������
                    .TextMatrix(j, COLP_������ĿID) = mrsPrice!������ĿID
                        
                    '��ʾ����Ƽ۵���Ŀ
                    .TextMatrix(j, COLP_�շ����) = mrsPrice!�շ����
                    .TextMatrix(j, COLP_�շ�ϸĿID) = mrsPrice!�շ�ϸĿID
                    .TextMatrix(j, COLP_���) = rs�շ�ϸĿ!�������
                    .TextMatrix(j, COLP_�շ���Ŀ) = rs�շ�ϸĿ!����
                    If Not IsNull(rs�շ�ϸĿ!����) Then
                        .TextMatrix(j, COLP_�շ���Ŀ) = .TextMatrix(j, COLP_�շ���Ŀ) & "(" & rs�շ�ϸĿ!���� & ")"
                    End If
                    If Not IsNull(rs�շ�ϸĿ!���) Then
                        .TextMatrix(j, COLP_�շ���Ŀ) = .TextMatrix(j, COLP_�շ���Ŀ) & " " & rs�շ�ϸĿ!���
                    End If
                    
                    If InStr(",5,6,7,", mrsPrice!�������) > 0 Then
                        'ҩƷҽ�������ҩƷ
                        .TextMatrix(j, COLP_��λ) = NVL(rs�շ�ϸĿ!סԺ��λ)
                    Else
                        '��������ҩƷ�����ļƼ�
                        .TextMatrix(j, COLP_��λ) = NVL(rs�շ�ϸĿ!���㵥λ)
                    End If
                    'ҩ��ȱʡΪ1,��ҩ��ҩƷ������(�ۼ۵�λ)
                    .TextMatrix(j, COLP_����) = FormatEx(mrsPrice!����, 5)
                    
                    'ҩ��ҩƷΪ��1��סԺ��λ����ļ۸�
                    .TextMatrix(j, COLP_����) = Format(NVL(mrsPrice!����), gstrDecPrice)
                    
                    If mrsPrice!�շ���� = "4" Then
                        .TextMatrix(j, COLP_��������) = Val(NVL(mrsPrice!����, 0))
                    End If
                    
                    'ִ�п���
                    lngִ�п���ID = NVL(mrsPrice!ִ�п���ID, 0)
                    '��ҩ��ҩƷ��������õ����ļƼۿ�������ִ�п���
                    If InStr(",4,5,6,7,", mrsPrice!�������) = 0 _
                        And (mrsPrice!�շ���� = "4" And NVL(mrsPrice!����, 0) = 1 Or InStr(",5,6,7,", mrsPrice!�շ����) > 0) Then
                        '�Ե�ǰֵ��Ϊȱʡ����ȡ��Ч��ִ�п���
                        lng���˿���ID = Val(vsAdvice.TextMatrix(lngRow, COL_���˿���ID))
                        lngִ�п���ID = Get�շ�ִ�п���ID(Val(vsAdvice.TextMatrix(lngRow, COL_����ID)), Val(vsAdvice.TextMatrix(lngRow, COL_��ҳID)), _
                            mrsPrice!�շ����, rs�շ�ϸĿ!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID, , , 2)
                        '��¼�Ƿ��������
                        .Editable = flexEDKbdMouse
                    End If
                    If lngִ�п���ID <> 0 Then
                        mrsDept.Filter = "ID=" & lngִ�п���ID
                        If Not mrsDept.EOF Then
                            .TextMatrix(j, COLP_ִ�п���) = mrsDept!����
                        End If
                    End If
                    .TextMatrix(j, COLP_ִ�п���ID) = lngִ�п���ID
                                        
                    '��۵Ĵ���
                    If NVL(rs�շ�ϸĿ!�Ƿ���, 0) = 1 Then
                        If InStr(",5,6,7,", mrsPrice!�շ����) > 0 Then
                            If InStr(",5,6,7,", mrsPrice!�������) > 0 Then
                                'ҩ��ҩƷ����1��סԺ��λ��ʱ��
                                .TextMatrix(j, COLP_����) = CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, NVL(rs�շ�ϸĿ!סԺ��װ, 1), , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                                .TextMatrix(j, COLP_����) = Format(Val(.TextMatrix(j, COLP_����)) * NVL(rs�շ�ϸĿ!סԺ��װ, 1), gstrDecPrice)
                            Else
                                '��ҩ��ҩƷ�����۵�λ����
                                .TextMatrix(j, COLP_����) = Format(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, mrsPrice!����, , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                            End If
                        ElseIf mrsPrice!�շ���� = "4" And NVL(mrsPrice!����, 0) = 1 Then
                            'ʱ�����ļ۸��ҩƷһ������
                            .TextMatrix(j, COLP_����) = Format(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, mrsPrice!����, , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                        Else
                            '��¼��������ļ۸�Χ
                            .Cell(flexcpData, j, 0) = 1 '��ʶΪ���(ҩƷ����)
                            .Cell(flexcpData, j, 1) = NVL(rs�շ�ϸĿ!ԭ��, 0)
                            .Cell(flexcpData, j, 2) = NVL(rs�շ�ϸĿ!�ּ�, 0)
                            'Ҳ����ǰ���˱��,���ڱ�۷�Χ����
                            If .TextMatrix(j, COLP_����) = "" Then
                                .TextMatrix(j, COLP_����) = Format(NVL(rs�շ�ϸĿ!ȱʡ�۸�), gstrDecPrice)
                            ElseIf .TextMatrix(j, COLP_����) <> "" Then
                                If CheckScope(NVL(rs�շ�ϸĿ!ԭ��, 0), NVL(rs�շ�ϸĿ!�ּ�, 0), NVL(mrsPrice!����, 0)) <> "" Then
                                    .TextMatrix(j, COLP_����) = Format(NVL(rs�շ�ϸĿ!ȱʡ�۸�), gstrDecPrice)
                                End If
                            End If
                            '��ۼ�ʹ�̶�Ҳ���Ա༭(�����Ǹ������õ�ʱ������ҽ��)
                            .Editable = flexEDKbdMouse
                        End If
                    End If

                    '��ʾҽ����������
                    If Val(mrsPrice!�շ�ϸĿID & "") <> 0 Then
                        strPriceType = GetPriceType(Val(mlng����ID), Val(mrsPrice!�շ�ϸĿID & ""), Val(mint����), mlng�������� = 1)
                    End If
                    '��������
                    If strPriceType = "" Then
                        .TextMatrix(j, COLP_��������) = NVL(rs�շ�ϸĿ!��������)
                    Else
                        .TextMatrix(j, COLP_��������) = strPriceType
                    End If
                    
                    .TextMatrix(j, COLP_�̶�) = mrsPrice!�̶�
                    .TextMatrix(j, COLP_����) = IIF(NVL(mrsPrice!����, 0) = 0, "", "��")
                    
                    '��¼���ڻָ�����
                    .Cell(flexcpData, j, COLP_�Ƽ�ҽ��) = .TextMatrix(j, COLP_�Ƽ�ҽ��)
                    .Cell(flexcpData, j, COLP_�շ���Ŀ) = .TextMatrix(j, COLP_�շ���Ŀ)
                    .Cell(flexcpData, j, COLP_����) = .TextMatrix(j, COLP_����)
                    .Cell(flexcpData, j, COLP_����) = .TextMatrix(j, COLP_����)
                    .Cell(flexcpData, j, COLP_ִ�п���) = .TextMatrix(j, COLP_ִ�п���)
                    
                    '��ʶ�̶�����Ϊ��ɫ
                    If mrsPrice!�̶� <> 0 Then
                        .Cell(flexcpBackColor, j, .FixedCols, j, .Cols - 1) = &HE0E0E0
                    End If
                    
                    j = j + 1
                End If
                
                mrsPrice.MoveNext
            Next
            
            '���ñ༭����
            '------------------------------------------------------------------
            '��Ҫ�Ƽ۵�ҽ��ѡ��
            If strCombo <> "" Then
                .ColData(COLP_�Ƽ�ҽ��) = Mid(strCombo, 2)
                .Editable = flexEDKbdMouse '����ѡ������Ա༭
            Else
                .ColData(COLP_�Ƽ�ҽ��) = ""
            End If
        End If
        .Row = .FixedRows: .Col = COLP_�Ƽ�ҽ��
        
        'ȱʡѡ��Ƽ�ҽ��(�������)
        Call ShowDefaultRow
        .Redraw = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub SetSameTime(ByVal lngRow As Long)
'���ܣ���������ҽ����Ϊ��ͬ��ֹͣ��ȷ��ֹͣ��У��,��ͣ,����ʱ��
    Dim strTime As String, vPause As Date, strCur As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        strTime = Format(.TextMatrix(lngRow, COL_����), "yyyy-MM-dd HH:mm")
        strCur = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        For i = .FixedRows To .Rows - 1
            If i <> lngRow Then
                blnDo = True
                If mint���� = 3 Then
                    blnDo = .Cell(flexcpData, i, COL_ѡ��) <> Empty
                Else
                    blnDo = Val(.TextMatrix(i, COL_ѡ��)) <> 0
                End If
                
                If blnDo Then
                    If (mint���� = 1 Or mint���� = 7) Then  'ֹͣ
                        'Ӧ>��ʼִ��ʱ��
                        If strTime <= Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                        'Ӧ>=����ʱ��
                        If strTime < Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                        
                        'Ӧ>=�ϴ�ִ��ʱ��,��Ϊ��ʱ�����ִ��,�����ջ���ǰֹͣ�ĳ���
                        If blnDo And .TextMatrix(i, COL_�ϴ�ִ��) <> "" Then
                            If strTime < Format(.Cell(flexcpData, i, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") And _
                                strCur > Format(.Cell(flexcpData, i, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                    ElseIf mint���� = 2 Then    'ȷ��ֹͣ
                        'Ӧ>=��ֹʱ��
                        If .TextMatrix(i, COL_��ֹʱ��) <> "" Then
                            If strTime < Format(.Cell(flexcpData, i, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        
                    ElseIf mint���� = 3 Then    'ҽ��У��
                        'Ӧ>=min(����ʱ��,��ʼʱ��)
                        If Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd HH:mm") < Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then
                            If strTime < Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                        Else
                            If strTime < Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        
                    ElseIf mint���� = 5 Then    '��ͣҽ��
                        'Ӧ>=��ʼִ��ʱ��,��Ϊ��ʱ�����δִ��
                        If strTime < Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                        'Ӧ>�ϴ�ִ��ʱ��,��Ϊ��ʱ�����ִ��
                        If .TextMatrix(i, COL_�ϴ�ִ��) <> "" Then
                            If strTime <= Format(.Cell(flexcpData, i, COL_�ϴ�ִ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        'Ӧ<ִ����ֹʱ��,��Ϊ��ʱ���ִ����Ч
                        If .TextMatrix(i, COL_��ֹʱ��) <> "" Then
                            If strTime >= Format(.Cell(flexcpData, i, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        'Ӧ>�ϴ���ͣ�������ʱ��(�����,����ʱ�䲻���ظ�,Ӧ>)
                        vPause = GetPauseTime(Val(.TextMatrix(i, COL_ID)), 7)
                        If vPause <> CDate(0) Then
                            If strTime <= Format(vPause, "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        
                    ElseIf mint���� = 6 Then    '����ҽ��
                        'Ӧ>��ͣʱ��
                        vPause = GetPauseTime(Val(.TextMatrix(i, COL_ID)), 6)
                        If vPause <> CDate(0) Then
                            If strTime <= Format(vPause, "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                        'Ӧ<=ִ����ֹʱ��
                        If .TextMatrix(i, COL_��ֹʱ��) <> "" Then
                            If strTime > Format(.Cell(flexcpData, i, COL_��ֹʱ��), "yyyy-MM-dd HH:mm") Then blnDo = False
                        End If
                    End If
                End If
                If blnDo Then
                    .TextMatrix(i, COL_����) = strTime
                    .Cell(flexcpData, i, COL_����) = strTime
                End If
            End If
        Next
    End With
End Sub

Private Function GetPauseTime(ByVal lngҽ��ID As Long, ByVal int״̬ As Integer) As Date
'���ܣ���ȡָ��ҽ������ͣʱ��(��ҽ����ǰӦ����ͣ)���ϴ�����ʱ��(�����)
'������int״̬=6-��ͣ,7-����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Max(����ʱ��) as �ϴ�ʱ�� From ����ҽ��״̬ Where ��������=[2] And ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, int״̬)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!�ϴ�ʱ��) Then
            GetPauseTime = rsTmp!�ϴ�ʱ��
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAdvice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    
    'Pass
    If Button = 2 Then
        With vsAdvice
            lngRow = .MouseRow
            If lngRow >= .FixedRows And lngRow <= .Rows - 1 Then
                If Not .RowHidden(lngRow) Then .Row = lngRow
            End If
        End With
    End If
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Pass
    Dim objPopup As CommandBarPopup
    Dim blnDo As Boolean
    If Button = 2 And mint���� = 3 Then
        If cbsMain Is Nothing Then Exit Sub
        If mblnPass Then
            blnDo = gobjPass.PassType = G_PASS_TYPE.DT Or gobjPass.PassType = G_PASS_TYPE.YWS Or (gobjPass.PassType = G_PASS_TYPE.MK And gobjPass.PassVersion = "4.0")
        Else
            blnDo = True
        End If
        If blnDo Then Exit Sub
        Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Function Get���˻���ȼ�ҽ��id(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngӤ����� As Long, ByVal lngҽ��ID As Long) As Long
'���ܣ�ȡ�����������ϵĻ���ȼ�ҽ�����������Զ�ֹͣ�Ļ���ȼ�ҽ��id
'˵���������ϻ���ȼ�ʱ���ã�65092����
'������
'      lng����id
'      lng��ҳid
'      lngӤ�����
'      lngҽ��id �������ϵĻ���ȼ�ҽ��id
'���أ�����ȼ�ҽ��id
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    
    On Error GoTo errH
    
    '1.�ж�У�Ե�ǰ����ҽ��ʱ�ǲ����Զ�ֹͣǰһ�εĻ���ȼ�ҽ����
    '�����β�ѯ�м�¼���� ����ȼ�id Ϊ�գ���������ֶ�ֹͣ�������������Ϊ�Զ�ֹͣ��
    strSQL = "Select a.����ȼ�id" & vbNewLine & _
        "From ���˱䶯��¼ A" & vbNewLine & _
        "Where a.����id = [1] And a.��ҳid = [2] And a.���Ӵ�λ = 0 And" & vbNewLine & _
        "      ��ֹʱ�� =" & vbNewLine & _
        "      (Select MIN(c.��ʼʱ��)" & vbNewLine & _
        "       From ����ҽ����¼ B, ���˱䶯��¼ C" & vbNewLine & _
        "       Where b.����id = c.����id And b.��ҳid = c.��ҳid And Trunc(c.��ʼʱ��, 'MI') = b.��ʼִ��ʱ�� And c.��ʼԭ�� = 6 And b.Id = [3])"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, lngҽ��ID)
    If Not rsTmp.EOF Then
        If Val(rsTmp!����ȼ�id & "") = 0 Then Exit Function
    End If
    
    '2.��ȡ�����Խ������õ� ����ȼ�ҽ�� �� ҽ��id
    strSQL = "Select ҽ��id" & vbNewLine & _
        "From (Select a.Id As ҽ��id" & vbNewLine & _
        "       From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
        "       Where a.������Ŀid = b.Id And a.������� = 'H' And b.�������� = '1' And a.����id = [1] And a.��ҳid = [2] And Nvl(a.Ӥ��, 0) = [3] And" & vbNewLine & _
        "             a.ҽ��״̬ In (8, 9) And a.Id <> [4] And a.��ʼִ��ʱ�� < (Select ��ʼִ��ʱ�� From ����ҽ����¼ Where ID = [4])" & vbNewLine & _
        "       Order By a.��ʼִ��ʱ�� Desc)" & vbNewLine & _
        "Where Rownum < 2"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, lngӤ�����, lngҽ��ID)
    If rsTmp.EOF Then Exit Function
   
    Get���˻���ȼ�ҽ��id = Val(rsTmp!ҽ��ID & "")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub zlPASSMap()
'����:����Pass VsAdvie����ӳ��
'ע��:ɾ�����޸�������������ʱ�����������ҩ�����еĹ�������
    If mobjPassMap Is Nothing Then
        Set mobjPassMap = DynamicCreate("zlPassInterface.clsPassMap", "������ҩ���", True)
        mblnPass = Not gobjPass Is Nothing And Not mobjPassMap Is Nothing
    End If
    
    If mblnPass Then
        With mobjPassMap
            .lngModel = PM_��ʿУ��
            Set .frmMain = Me
            Set .vsAdvice = vsAdvice
            Set .objCmdBar = cmdAlley
            Set .VSCOL = .GetVSCOL(COL_ID, COL_���ID, COL_�������, COL_������ĿID, COL_�շ�ϸĿID, col_ҽ������, COL_��Ч, COL_����, COL_������λ, _
                        COL_�÷�, , COL_Ӥ��, COL_����ʱ��, COL_����ҽ��, COL_��ʼʱ��, COL_��������ID, COL_��ֹʱ��, COL_Ƶ��, COL_Ƶ�ʴ���, COL_Ƶ�ʼ��, _
                        COL_�����λ, COL_��ʾ, COL_���, , , COL_����ID, COL_��ҳID, COL_ѡ��, COL_ִ������, COL_�걾��λ)
        End With
        mblnPass = gobjPass.zlPassCheck(mobjPassMap)
    End If
End Sub

Private Sub GetDefaultDeptPar(ByVal lng���˿���ID As Long)
'���ܣ���ȡȱʡ����
    mlng��ҩ�� = Val(zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , lng���˿���ID))
    mlng��ҩ�� = Val(zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , lng���˿���ID))
    mlng��ҩ�� = Val(zlDatabase.GetPara("סԺȱʡ��ҩ��", glngSys, pסԺҽ���´�, , , , , lng���˿���ID))
    mlng���ϲ��� = Val(zlDatabase.GetPara("סԺȱʡ���ϲ���", glngSys, pסԺҽ���´�, , , , , lng���˿���ID))
End Sub

Private Sub SetFilterTime()
'���ܣ����ý�������Ĺ���������ʼ״̬
    Dim strTmp As String
    
    If mint���� = 3 Then
        strTmp = mstrȱʡУ��ʱ��
        If Left(strTmp, 1) = "0" Then
            optOper(e��ǰʱ��).value = True '��ǰʱ��
            lblS.Visible = False
            lblB.Visible = False
            cboTime(e����).Visible = False
            cboTime(e����).Visible = False
        Else
            optOper(e��ʼʱ��).value = True
            lblS.Visible = True
            lblB.Visible = True
            cboTime(e����).Visible = True
            cboTime(e����).Visible = True
        End If
        cboTime(e����).ListIndex = IIF(Mid(strTmp, 2, 1) = "1", 1, 0)
        cboTime(e����).ListIndex = IIF(Mid(strTmp, 3, 1) = "1", 1, 0)
    ElseIf (mint���� = 1 Or mint���� = 7) Then
        strTmp = mstrȱʡֹͣʱ��
        If Left(strTmp, 1) = "1" Then
            optStop(e�ϴ�ִ��ʱ��).value = True '�ϴ�ִ��ʱ��
            chkNoSend.Visible = False
            chkRollSend.Visible = False
        Else
            optStop(eָ��ʱ��).value = True
            chkNoSend.value = IIF(Mid(strTmp, 2, 1) = "1", 1, 0)
            chkRollSend.value = IIF(Mid(strTmp, 3, 1) = "1", 1, 0)
        End If
    End If
End Sub

Private Sub SetSameԭ��(ByVal lngRow As Long)
'���ܣ���������ҽ����Ϊ��ͬ����ֹԭ��
    Dim strԭ�� As String
    Dim i As Long
    
    Call vsAdvice_AfterEdit(lngRow, COL_��ֹԭ��)
    
    If Not VsfOnlySelOneRow(lngRow) Then
    
        strԭ�� = vsAdvice.TextMatrix(lngRow, COL_��ֹԭ��)
        
        If MsgBox("Ҫ����������ѡ���ҽ������Ϊ���ͣ��ԭ��" & strԭ�� & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If i <> lngRow Then
                    If Val(.TextMatrix(i, COL_ѡ��)) <> 0 Then
                         .TextMatrix(i, COL_��ֹԭ��) = strԭ��
                        .Cell(flexcpData, i, COL_��ֹԭ��) = strԭ��
                    End If
                End If
            Next
        End With
    End If
End Sub

Private Sub InitRecordSet(ByRef rsAdviceTmp As ADODB.Recordset, ByRef rsMsgRow As ADODB.Recordset, ByRef rs��Ѫ As ADODB.Recordset)
'���ܣ���ʼ�����ؼ�¼��
    Set rsAdviceTmp = New ADODB.Recordset
    rsAdviceTmp.Fields.Append "����ID", adBigInt
    rsAdviceTmp.Fields.Append "��ҳID", adBigInt
    rsAdviceTmp.Fields.Append "ҽ��IDs", adVarChar, 4000
    rsAdviceTmp.CursorLocation = adUseClient
    rsAdviceTmp.LockType = adLockOptimistic
    rsAdviceTmp.CursorType = adOpenStatic
    
    Set rsMsgRow = New ADODB.Recordset
    rsMsgRow.Fields.Append "����ID", adBigInt
    rsMsgRow.Fields.Append "��ҳID", adBigInt
    rsMsgRow.Fields.Append "�к�", adBigInt
    rsMsgRow.Fields.Append "��������", adBigInt '1��ֹͣ��2�����ϣ�3��У��ͨ����4-У������
    rsMsgRow.Fields.Append "��ǰ����", adVarChar, 4000
    rsMsgRow.CursorLocation = adUseClient
    rsMsgRow.LockType = adLockOptimistic
    rsMsgRow.CursorType = adOpenStatic
    rsMsgRow.Open
    
    Set rs��Ѫ = New ADODB.Recordset
    With rs��Ѫ
        .Fields.Append "ҽ��ID", adBigInt
        .Fields.Append "����", adInteger '3��У�ԣ�4������
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub

Private Sub Check�������()
'���ܣ������󷽽ӿ��жϵ�ǰҽ���ǲ���������
    Dim i As Long
    Dim str��ҩIDs As String '���뵽�ӿ��еĲ���
    Dim strOutҽ��IDs As String '���ܹ����͵���ҽ��ID
    Dim strErr As String
    Dim lngҽ��ID As Long
    Dim strҽ������ As String
    Dim str���� As String
    Dim lngLastPatiID As Long
    Dim lngLastPageID As Long
    Dim rsTmp As ADODB.Recordset
    Dim j As Long
    Dim strҩ��ҽ��IDs As String
    
    On Error GoTo errH
    
    If Not gbln��ϵͳ Then Exit Sub
    
    With vsAdvice
        'У�Դ��ڶಡ��ģʽ
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ѡ��)) <> 0 Or .Cell(flexcpData, i, COL_ѡ��) <> Empty Then
                If .TextMatrix(i, COL_�������) = "5" Or .TextMatrix(i, COL_�������) = "6" Then
                    If lngLastPatiID <> Val(.TextMatrix(i, COL_����ID)) Then
                        lngLastPatiID = Val(.TextMatrix(i, COL_����ID))
                        lngLastPageID = Val(.TextMatrix(i, COL_��ҳID))
                    
                        Set rsTmp = Nothing
                        Call gobjPass.ZLPharmReviewResultView(lngLastPatiID, lngLastPageID, rsTmp, strErr)
                        If Not rsTmp Is Nothing Then
                            If Not rsTmp.EOF Then
                                For j = 1 To rsTmp.RecordCount
                                    If InStr("," & strOutҽ��IDs & ",", "," & rsTmp!���ID & ",") = 0 Then
                                        strOutҽ��IDs = strOutҽ��IDs & "," & rsTmp!���ID
                                    End If
                                    strҩ��ҽ��IDs = strҩ��ҽ��IDs & "," & rsTmp!ҽ��ID
                                    rsTmp.MoveNext
                                Next
                            End If
                        End If
                        
                        
                    End If
                End If
            End If
        Next
   
        If strOutҽ��IDs <> "" Then
            'ȡ��ѡ��
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) <> 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    lngҽ��ID = IIF(0 = Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                    If InStr("," & strOutҽ��IDs & ",", "," & lngҽ��ID & ",") > 0 Then
                        Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                        .Cell(flexcpData, i, COL_ѡ��) = 0
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                            If InStr("," & strҩ��ҽ��IDs & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") > 0 Then
                                strҽ������ = strҽ������ & vbCrLf & .TextMatrix(i, col_ҽ������)
                            End If
                        End If
                    End If
                End If
            Next
            If strҽ������ <> "" Then
                Call MsgBox("����ҽ��δͨ��������飬����У�ԣ�" & strҽ������, vbInformation, Me.Caption)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
