VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMediPriceCard 
   Caption         =   "ҩƷ���۵�"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15150
   Icon            =   "frmMediPriceCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   2775
      TabIndex        =   32
      Top             =   4200
      Width           =   2775
   End
   Begin VB.PictureBox picOtherSelect 
      Height          =   3255
      Left            =   3360
      ScaleHeight     =   3195
      ScaleWidth      =   4875
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdFilterOk 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   2640
         Picture         =   "frmMediPriceCard.frx":6852
         TabIndex        =   28
         Top             =   2760
         Width           =   1100
      End
      Begin VB.CommandButton cmdFilterCan 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   3720
         Picture         =   "frmMediPriceCard.frx":699C
         TabIndex        =   27
         Top             =   2760
         Width           =   1100
      End
      Begin VB.Frame fra����ѡ�� 
         Caption         =   "����ѡ��ɱ��۵�����أ�"
         Height          =   2535
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4695
         Begin VB.CheckBox chk�ӳ��� 
            Caption         =   "ָ���ӳ���"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   1125
            Width           =   1215
         End
         Begin VB.CheckBox chk��Ӧ�� 
            Caption         =   "ָ����Ӧ��"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkӦ����¼ 
            Caption         =   "�����ɱ��۵��۴�����Ӧ����������¼"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txt�ӳ��� 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   19
            Text            =   "15.0000"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txt��Ӧ�� 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   18
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmd��Ӧ�� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   270
            Left            =   4080
            TabIndex        =   17
            Top             =   350
            Width           =   375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
            Height          =   1695
            Left            =   120
            TabIndex        =   23
            Top             =   2280
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   2990
            _Version        =   393216
            FixedCols       =   0
            GridColor       =   32768
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblComment�ӳ��� 
            Caption         =   "��ָ���ӳ��ʣ���ͳһĬ�ϰ��üӳ��ʼ���ɱ��ۣ���ָ������Ĭ����ʾʵ�ʼӳ��ʣ�"
            ForeColor       =   &H00FF0000&
            Height          =   540
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   4260
         End
         Begin VB.Label lblComment��Ӧ�� 
            AutoSize        =   -1  'True
            Caption         =   "��ָ����Ӧ�̣���ֻ�����ù�Ӧ�̵Ŀ��ҩƷ�ɱ��ۣ�"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Width           =   4320
         End
         Begin VB.Label lblPercent 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   180
            Left            =   2415
            TabIndex        =   24
            Top             =   1125
            Width           =   90
         End
      End
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   14175
      TabIndex        =   10
      Top             =   8160
      Width           =   14175
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   600
         TabIndex        =   35
         Top             =   120
         Width           =   1365
      End
      Begin VB.TextBox txtSummary 
         Height          =   300
         Left            =   5040
         MaxLength       =   100
         TabIndex        =   13
         Top             =   120
         Width           =   8835
      End
      Begin VB.TextBox txtValuer 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lblFind 
         BackColor       =   &H80000003&
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lblSummary 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "����˵��"
         Height          =   180
         Left            =   4200
         TabIndex        =   14
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblValuer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "������"
         Height          =   180
         Left            =   2160
         TabIndex        =   12
         Top             =   180
         Width           =   540
      End
   End
   Begin VB.Frame fraCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   16575
      Begin VB.PictureBox picAdjustTime 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3840
         ScaleHeight     =   375
         ScaleWidth      =   5535
         TabIndex        =   39
         Top             =   120
         Width           =   5535
         Begin VB.OptionButton optʱ�� 
            BackColor       =   &H80000003&
            Caption         =   "ָ������"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   41
            Top             =   15
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optʱ�� 
            BackColor       =   &H80000003&
            Caption         =   "����ִ��"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   40
            Top             =   15
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpRunDate 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   3000
            TabIndex        =   42
            Top             =   0
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   127729667
            CurrentDate     =   36846.5833333333
         End
         Begin VB.Label lblִ��ʱ�� 
            BackColor       =   &H80000003&
            Caption         =   "ִ��ʱ��"
            Height          =   180
            Left            =   0
            TabIndex        =   43
            Top             =   45
            Width           =   855
         End
      End
      Begin VB.TextBox txtNO 
         Enabled         =   0   'False
         Height          =   300
         Left            =   14640
         TabIndex        =   37
         Top             =   120
         Width           =   1695
      End
      Begin VB.CheckBox chkAutoPay 
         BackColor       =   &H80000003&
         Caption         =   "�Զ�����Ӧ����䶯��¼"
         Height          =   210
         Left            =   3360
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chkAotuCost 
         BackColor       =   &H80000003&
         Caption         =   "���ۼ�ʱ�Զ����ӳ��ʵ����ɱ���"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   3000
      End
      Begin VB.CommandButton cmdPriceMethod 
         Caption         =   "��"
         Height          =   300
         Left            =   3360
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboPriceMethod 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   2415
      End
      Begin VB.CheckBox chk������ 
         Caption         =   "�ɱ��۰��ⷿ���ε���"
         Height          =   210
         Left            =   10560
         TabIndex        =   3
         Top             =   -225
         Width           =   2175
      End
      Begin VB.CheckBox chk�Զ�����Ӧ����䶯 
         Caption         =   "�Զ�����Ӧ����䶯"
         Height          =   210
         Left            =   12840
         TabIndex        =   2
         Top             =   -225
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.ComboBox cbo�ۼۼ��㷽ʽ 
         Height          =   300
         Left            =   10800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "������ˮ��"
         Height          =   180
         Left            =   13560
         TabIndex        =   38
         Top             =   180
         Width           =   900
      End
      Begin VB.Label lbl���۷�ʽ 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "�ۼۼ��㷽ʽ"
         Height          =   180
         Left            =   9480
         TabIndex        =   9
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label lblMethod 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "���۷�ʽ"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   720
      End
   End
   Begin XtremeSuiteControls.TabControl TabCtlDetails 
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   1815
      _Version        =   589884
      _ExtentX        =   3201
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfStore 
      Height          =   975
      Left            =   2880
      TabIndex        =   30
      Top             =   4680
      Width           =   3495
      _cx             =   6165
      _cy             =   1720
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPay 
      Height          =   975
      Left            =   8040
      TabIndex        =   31
      Top             =   4680
      Width           =   3495
      _cx             =   6165
      _cy             =   1720
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPrice 
      Height          =   2295
      Left            =   480
      TabIndex        =   33
      Top             =   2040
      Width           =   11055
      _cx             =   19500
      _cy             =   4048
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   34
      Top             =   8715
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediPriceCard.frx":6AE6
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20955
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
            Object.ToolTipText     =   "��ǰ���ּ�״̬"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "��ǰ��д��״̬"
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
   Begin XtremeCommandBars.ImageManager imgList 
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMediPriceCard.frx":737A
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
Attribute VB_Name = "frmMediPriceCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'����ȫ�ֱ���
Private Const mlngRowHeight As Long = 300 '����и����и�
Private mintUnit As Integer     '������¼���õ���ʲô��λ
Private mint���� As Integer     '0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
Private mlng��Ӧ��ID As Long  '������¼��Ӧ��id
Private mdbl�ӳ��� As Double
Private mblnӦ����¼ As Boolean '��¼�Ƿ����Ӧ����¼

Private Enum typeAdjust
    AdjustPriceAndCost = 0
    AdjustPrice = 1
    AdjustCost = 2
End Enum

Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��
Private mstrMoneyFormat As String
Private mintSalePriceDigit As Integer
'��ɫ����
Private Const mconlngColor As Long = &HFFFFFF        '�����޸�����ɫΪ��ɫ
Private Const mconlngCanColColor As Long = &HE7CFBA    '���޸�����ɫΪ����ɫ
Private Const mlngBorderColor As Long = &H0&    'ѡ���б߿���ɫ
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' ûѡ���б߿���ɫ

Private mblnʱ��ҩƷ�����ε��� As Boolean 'ʱ��ҩƷ�������ε���
Private mbln�ɱ��۰��ⷿ���ε��� As Boolean '�ɱ��۰��ⷿ���ε���
Private mbln�ּ���ʾ As Boolean         '�޼�ҩƷ��ʾ true-��ʾ false-����ʾ
Private mdbl�ֶμӳ��� As Double    '������¼�ֶμӳ���
Private mdbl�ɱ��� As Double            '��¼�޸�֮ǰ�ĳɱ���
Private mrs�ֶμӳ� As ADODB.Recordset  '��¼�ֶμӳ��ʼ���
Private mstrNo As String            '���۵�No
Private mintModal As Integer        '������ʲô״̬ 0-���� 1-�޸� 2-����
Private mintMethod As Integer   '���۷�ʽ 0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
Private mstr���ۻ��ܺ� As String
Private mblnLoad As Boolean     '�Ƿ�������
Private mrsReturn As ADODB.Recordset '����ѡ�񷵻ص����ݼ�
Private mblnOK As Boolean
Private mrsFindName As ADODB.Recordset '��ѯ�����ݼ�
Private mBlnClick As Boolean
Private mblnUpdateAdd As Boolean    '�޸�����µ���������
Private mlngOldDrugID As Long '���ԭʼ���Ƿ���ҩƷ
Private mdblOldPrice As Double   'ԭ�ۼ�
Private mblnBatchItem As Boolean   '��¼�Ƿ���������ѡ��ť
Private mstrPrivs As String     '����ԱȨ��
Private Const MStrCaption As String = "ҩƷ���۵�"

'���ܰ�ť
Private Const mconMenu_Save = 100 'ȷ��(&A)
Private Const mconMenu_Quit = 101 'ȡ��(&Q)
Private Const mconMenu_PrintStore = 102 '��ӡ���䶯��(&P)
Private Const mconMenu_ClearAll = 103 '����б�(&C)
Private Const mconMenu_BatchSelect = 104 '����ѡ����Ŀ
Private Const mconMenu_Find = 105 '����
Private Const mconMenu_ModifyPrice = 106 '���۷�ʽ
Private Const mconMenu_CostPrice = 107 '�����ɱ���
Private Const mconMenu_RetailPrice = 108  '�����ۼ�
Private Const mconMenu_Together = 109  '�ɱ����ۼ�һ���
Private Enum menuPriceCol
    ҩƷid = 0
    ԭ��id = 1
    ҩƷ = 2
    ��� = 3
    ҩ������ = 4
    �Ƿ���
    ����
    ��λ
    ��װϵ��
    �ӳ���
    ���������
    �Ƿ��п��
    ������ĿID
    ԭ�ɱ���
    �ֳɱ���
    ԭ���ۼ�
    �����ۼ�
    ԭ�ɹ��޼�
    �ֲɹ��޼�
    ԭָ���ۼ�
    ��ָ���ۼ�
    ������
End Enum
Private Enum menuStoreCol
    ҩƷid = 0
    ҩƷ = 1
    ��� = 2
    �ⷿ = 3
    �ⷿid = 4
    ��Ӧ��
    ��Ӧ��id
    ����
    Ч��
    ����
    ����
    ���
    ����
    ��λ
    ��װϵ��
    ԭ�ɱ���
    �ֳɱ���
    �ɱ�ӯ��
    �ӳ���
    ԭ���ۼ�
    �����ۼ�
    �ۼ�ӯ��
    ������
End Enum

Private Enum menuPayCol
    ҩƷid = 0
    ҩƷ = 1
    ��Ʊ�� = 2
    ��Ʊ����
    ��Ʊ���
    ������
End Enum

Public Sub ShowME(ByVal frmParent As Form, ByVal intModal As Integer, ByVal str���ۻ��ܺ� As String, ByVal intMethod As Integer)
    mintModal = intModal
    mstr���ۻ��ܺ� = str���ۻ��ܺ�
    mintMethod = intMethod

    Me.Show vbModal, frmParent
End Sub

Private Sub cboPriceMethod_Click()
    Dim intCol As Integer
    Dim intTemp As Integer

    With cboPriceMethod
        If .Text = "�����ۼ�" Then
            intTemp = 0
            lbl���۷�ʽ.Visible = False
            cbo�ۼۼ��㷽ʽ.Visible = False
        ElseIf .Text = "�����ɱ���" Then
            intTemp = 1
            lbl���۷�ʽ.Visible = False
            cbo�ۼۼ��㷽ʽ.Visible = False
        Else
            intTemp = 2
            lbl���۷�ʽ.Visible = True
            cbo�ۼۼ��㷽ʽ.Visible = True
        End If
    End With


    If mblnLoad = True And intTemp <> Val(lblMethod.Tag) Then
        If vsfPrice.TextMatrix(1, menuPriceCol.ҩƷid) <> "" Then
            If MsgBox("���۷�ʽ�ı佫����б������ݣ��Ƿ������", vbYesNo, gstrSysName) = vbNo Then
                cboPriceMethod.ListIndex = mint����
                Exit Sub
            Else
                vsfPrice.rows = 2
                For intCol = 0 To vsfPrice.Cols - 1
                    vsfPrice.TextMatrix(1, intCol) = ""
                Next
                vsfStore.rows = 1
                vsfPay.rows = 1
            End If
        End If
    End If
    With cboPriceMethod
        If .Text = "�����ۼ�" Then
            mint���� = 0
            lblMethod.Tag = 0
            optʱ��(0).Value = False
            optʱ��(1).Value = True
            optʱ��(0).Enabled = True
            optʱ��(1).Enabled = True
            dtpRunDate.Enabled = True
            chkAutoPay.Visible = False
            chkAutoPay.Value = 0
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "�����ɱ���" Then
            mint���� = 1
            lblMethod.Tag = 1
'            optʱ��(0).Value = True
            optʱ��(0).Enabled = True
            optʱ��(1).Enabled = True
            dtpRunDate.Enabled = True
            If mblnӦ����¼ = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            End If
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "�ۼ۳ɱ���һ�����" Then
            mint���� = 2
            lblMethod.Tag = 2
            optʱ��(0).Value = False
            optʱ��(1).Value = True
            optʱ��(0).Enabled = True
            optʱ��(1).Enabled = True
            dtpRunDate.Enabled = True
            If mblnӦ����¼ = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = True
        End If
        If .Text = "�����ۼ�" Then
            cmdPriceMethod.Visible = False
            picOtherSelect.Visible = cmdPriceMethod.Visible
        Else
            cmdPriceMethod.Visible = True
        End If
    End With
    vsfStore.Cols = menuStoreCol.������
    vsfPay.Cols = menuPayCol.������
    vsfPrice.Cols = menuPriceCol.������
    Call setColEdit
    Call setColHiddenVsf
End Sub

Private Sub cboPriceMethod_DropDown()
    With cboPriceMethod
        If .Text = "�����ۼ�" Then
            mint���� = 0
        ElseIf .Text = "�����ɱ���" Then
            mint���� = 1
        ElseIf .Text = "�ۼ۳ɱ���һ�����" Then
            mint���� = 2
        End If
    End With
End Sub

Private Sub cbo�ۼۼ��㷽ʽ_Click()
    On Error GoTo errHandle
    Set mrs�ֶμӳ� = Nothing
    If cbo�ۼۼ��㷽ʽ.Text = "�ۼ۰��ֶμӳɼ���" Then
        gstrSQL = "select ���, ��ͼ�, ��߼�, �ӳ���, ��۶�, ˵��, ���� from ҩƷ�ӳɷ��� order by ���"
        Set mrs�ֶμӳ� = zlDatabase.OpenSQLRecord(gstrSQL, "ҩƷ�ӳɷ���")
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_Save  '����
            Call Save
        Case mconMenu_PrintStore    '��ӡ���䶯��
            Call PrintStore
        Case mconMenu_ClearAll  '���
            Call ClearAll
        Case mconMenu_Find '����
            txtFind.SetFocus
            If Trim(txtFind.Text) = "" Then Exit Sub
            Call FindGridRow(UCase(Trim(txtFind.Text)))
        Case mconMenu_Quit  'ȡ��
            Call Quit
        Case mconMenu_BatchSelect  '����ѡ����Ŀ
            Call BatchSelect
    End Select
End Sub

Private Sub chkAotuCost_Click()
    If chkAotuCost.Value = 1 Then
        cbo�ۼۼ��㷽ʽ.Visible = False
        cbo�ۼۼ��㷽ʽ.ListIndex = 0
        lbl���۷�ʽ.Visible = False
    Else
        cbo�ۼۼ��㷽ʽ.Visible = True
        lbl���۷�ʽ.Visible = True
    End If
End Sub


Private Sub Chk��Ӧ��_Click()
    If chk��Ӧ��.Value = 1 Then
        cmd��Ӧ��.Enabled = True
        txt��Ӧ��.Enabled = True
        chkӦ����¼.Enabled = True
    Else
        cmd��Ӧ��.Enabled = False
        txt��Ӧ��.Enabled = False
        chkӦ����¼.Enabled = False
        chkӦ����¼.Value = 0
    End If
End Sub

Private Sub chk�ӳ���_Click()
    If chk�ӳ���.Value = 1 Then
        txt�ӳ���.Enabled = True
    Else
        txt�ӳ���.Enabled = False
    End If
End Sub

Private Sub Quit()
    Call ReleaseSelectorRS 'ж�����ݼ�
    Unload Me
End Sub

Private Sub ClearAll()
    Dim intCol As Integer

    If MsgBox("��ȷ��Ҫ����������ݣ�", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        vsfPrice.rows = 2
        For intCol = 0 To vsfPrice.Cols - 1
            vsfPrice.TextMatrix(1, intCol) = ""
        Next
        vsfStore.rows = 1
        vsfPay.rows = 1
    End If
End Sub

Private Sub cmdFilterCan_Click()
    picOtherSelect.Visible = False
End Sub

Private Sub cmdFilterOk_Click()
    Dim i As Integer

    If chk��Ӧ��.Value = 1 Then
        If Val(Split(txt��Ӧ��.Tag, "|")(0)) = 0 Then
            MsgBox "��ѡ��Ӧ�̡�", vbInformation, gstrSysName
            txt��Ӧ��.SetFocus
            Exit Sub
        End If
    End If
    With vsfPrice
        If Val(.TextMatrix(1, menuPriceCol.ҩƷid)) <> 0 Then
            If MsgBox("����ձ���е����ݣ��Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            Else
                vsfPrice.rows = 2
                For i = 0 To vsfPrice.Cols - 1
                    .TextMatrix(1, i) = ""
                Next
                vsfStore.rows = 1
                vsfPay.rows = 1
            End If
        End If
    End With

    mlng��Ӧ��ID = IIf(chk��Ӧ��.Value = 1, Val(Split(txt��Ӧ��.Tag, "|")(0)), 0)
    mdbl�ӳ��� = IIf(chk�ӳ���.Value = 1, Val(Trim(txt�ӳ���.Text)), 0)
    mblnӦ����¼ = (chkӦ����¼.Enabled And chkӦ����¼.Value = 1)
    picOtherSelect.Visible = False
    If mblnӦ����¼ = True Then
        TabCtlDetails.Item(1).Visible = True
    Else
        TabCtlDetails.Item(1).Visible = False
    End If

    With cboPriceMethod
        If .Text = "�����ۼ�" Then
            mint���� = 0
            lblMethod.Tag = 0
            optʱ��(0).Value = False
            optʱ��(1).Value = True
            optʱ��(0).Enabled = True
            optʱ��(1).Enabled = True
            dtpRunDate.Enabled = True
            chkAutoPay.Visible = False
            chkAutoPay.Value = 0
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "�����ɱ���" Then
            mint���� = 1
            lblMethod.Tag = 1
            optʱ��(0).Value = True
            optʱ��(0).Enabled = False
            optʱ��(1).Enabled = False
            dtpRunDate.Enabled = False
            If mblnӦ����¼ = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "�ۼ۳ɱ���һ�����" Then
            mint���� = 2
            lblMethod.Tag = 2
            optʱ��(0).Value = False
            optʱ��(1).Value = True
            optʱ��(0).Enabled = True
            optʱ��(1).Enabled = True
            dtpRunDate.Enabled = True
            If mblnӦ����¼ = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = True
        End If
    End With

End Sub

Private Sub CmdHelp_Click()

End Sub

Private Sub BatchSelect()
    Dim intRow As Integer

    frmBatchSelect.ShowME Me, mrsReturn, mblnOK

    On Error GoTo errHandle
    If mblnOK = False Then Exit Sub
    If mrsReturn.RecordCount = 0 Then Exit Sub

    With vsfPrice
        If .TextMatrix(.rows - 1, menuPriceCol.ҩƷid) = "" Then
            intRow = .rows - 1
        Else
            .rows = .rows + 1
            intRow = .rows - 1
        End If
    End With
    mblnBatchItem = True

    Call GetDrugPirce(mrsReturn, intRow)
    mblnBatchItem = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub deleteNotExecutePirce()
    '���δִ�м۸�
    Dim intRow As Integer
    Dim intɾ������ As Integer
    
    'Private mint���� As Integer     '0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
    'ɾ����ʽ_In   In Number := 0 --0-����;1-�ۼ�;2-�ɱ���
    On Error GoTo errHandle
    
    If mint���� = 0 Then
        intɾ������ = 1
    ElseIf mint���� = 1 Then
        intɾ������ = 2
    Else
        intɾ������ = 0
    End If
    
    With vsfPrice
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, menuPriceCol.ҩƷid) <> "" Then
                gstrSQL = "Zl_ҩƷδִ�м۸�_Delete(" & Val(.TextMatrix(intRow, menuPriceCol.ҩƷid)) & "," & intɾ������ & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
        Next
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Save()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim dtToday As Date
    Dim lngAdjId As Long
    Dim LngCurID As Long
    Dim strID As String
    Dim intCount As Integer
    Dim dbl��װ As Double
    Dim strTmp As String
    Dim lngCurrBatch As Long
    Dim str���μ۸� As String
    Dim blnPrint As Boolean '�Ƿ��ӡ����֪ͨ��
    Dim blnOne As Boolean   '����Ƿ��ǵ�һ��
    Dim n As Integer
    Dim intProc As Integer
    Dim blnIgnore As Boolean
    Dim blnPrice As Boolean '��¼�Ƿ��ۼ۵�����
    Dim blnCost As Boolean  '��¼�Ƿ�ɱ��۵�����
    Dim intUpdateModel As Integer '����ģʽ 0-�ۼ۵��� 1-�ɱ��۵��� 2-�ɱ����ۼ�һ�����
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim ArrayID
    Dim Array���μ۸�
    Dim strUpdate As String

    Dim lng�ⷿID As Long
    Dim lng��Ӧ��ID As Long
    Dim lngҩƷid As Long
    Dim lng����  As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim str���� As String
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim Str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim dbl��Ʊ��� As Double
    Dim strInfo As String
    Dim strMsg As String '��¼��ʾ��Ϣ
    Dim intCount2 As Integer '��������
    Dim lngDouID As Long
    
    Dim strִ��ʱ�� As String
    Dim str��ֹʱ�� As String
    Dim strDrugs As String
    
    If vsfPrice.rows > 1 Then   'ֻ�������ݵ�����²��ܱ���
        If Val(vsfPrice.TextMatrix(1, menuPriceCol.ҩƷid)) = 0 Then Exit Sub
    End If
    If CheckPrice = False Then Exit Sub
    
    On Error GoTo ErrHand
    
    dtToday = Sys.Currentdate()
    If optʱ��(0).Value = True Then
        strִ��ʱ�� = Format(dtToday, "YYYY-MM-DD HH:mm:ss")
        str��ֹʱ�� = Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss")
    Else
        strִ��ʱ�� = Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss")
        str��ֹʱ�� = Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss")
    End If
                    
    gstrSQL = "select �շѼ�Ŀ_ID.nextval from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�շѼ�Ŀ���")
    lngAdjId = rsTemp.Fields(0).Value

    gcnOracle.BeginTrans
    If mintModal = 1 Then '�޸� ���޸�ģʽ����ɾ��ԭ���ĵ�����Ϣ��Ȼ������µĵ�����Ϣ
        Call deleteNotExecutePirce
    End If

    '����Ƿ����δִ�еļ۸�
    If checkNotExecutePrice(, strInfo) = True Then
        MsgBox strInfo, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ȡ����NO
    mstrNo = Sys.GetNextNo(9)
    '��ȡ���ۻ���NO
    gstrSQL = "select nextno(135) as ��ˮ�� from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������ˮ��")
    If rsTemp.RecordCount = 0 Then
        MsgBox "������ˮ��δ�ܳ�ʼ���ɹ����������Ա��ϵ��", vbInformation, gstrSysName
        Exit Sub
    End If
    txtNO.Text = rsTemp!��ˮ��

    With Me.vsfPrice
        '�ۼ۵���
        strID = ""
        For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
            If mint���� <> 1 Then
                LngCurID = Sys.NextId("�շѼ�Ŀ")
                
                strID = strID & IIf(strID = "", "", ",") & LngCurID
                
                If InStr(1, "," & strDrugs & ",", "," & Val(.TextMatrix(intCount, menuPriceCol.ҩƷid)) & ",") = 0 Then
                    strDrugs = IIf(strDrugs = "", "", strDrugs & ",") & Val(.TextMatrix(intCount, menuPriceCol.ҩƷid))
                End If
                
                dbl��װ = Val(.TextMatrix(intCount, menuPriceCol.��װϵ��))

                If .TextMatrix(intCount, menuPriceCol.�Ƿ���) = "1" And mblnʱ��ҩƷ�����ε��� And mint���� <> 1 Then
                    strTmp = ""
                    lngCurrBatch = -1
                    For n = 1 To vsfStore.rows - 1
                        If Val(.TextMatrix(intCount, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                            If InStr(1, "|" & strTmp, "|" & vsfStore.TextMatrix(n, menuStoreCol.����) & ",") = 0 Then
                                lngCurrBatch = vsfStore.TextMatrix(n, menuStoreCol.����)
                                strTmp = strTmp & IIf(strTmp = "", "", "|") & vsfStore.TextMatrix(n, menuStoreCol.����) & "," & vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�) / dbl��װ
                            End If
                        End If
                    Next
                    str���μ۸� = str���μ۸� & strTmp
                End If
                str���μ۸� = str���μ۸� & ";"
                             
                If CLng(.TextMatrix(intCount, menuPriceCol.ԭ��id)) <> 0 Then
                    '������һ�εļ۸��¼��ִֹ��
                    gstrSQL = "zl_�շѼ�Ŀ_stop(" & .TextMatrix(intCount, menuPriceCol.ҩƷid) & ","
                    If optʱ��(0).Value = True Then
                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)

                    '�����۸��¼
                    gstrSQL = "zl_�շѼ�Ŀ_Insert(" & LngCurID & "," & IIf(.TextMatrix(intCount, menuPriceCol.ԭ��id) = "", "NUll", Val(.TextMatrix(intCount, menuPriceCol.ԭ��id))) & _
                              "," & .TextMatrix(intCount, menuPriceCol.ҩƷid) & "," & Val(.TextMatrix(intCount, menuPriceCol.������ĿID)) & "," & _
                              Round(Val(.TextMatrix(intCount, menuPriceCol.ԭ���ۼ�)) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�) & "," & _
                              Round(Val(.TextMatrix(intCount, menuPriceCol.�����ۼ�)) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�) & _
                              ",NULL,NULL,'" & Me.txtSummary.Text & "'," & lngAdjId & ",'" & Trim(Me.txtValuer.Text) & "',"
                    If optʱ��(0).Value = True Then
                        gstrSQL = gstrSQL & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ",0,'" & mstrNo & "'," & intCount & ",Null," & txtNO & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                    blnPrice = True
                    blnPrint = True
                End If
                
                If .TextMatrix(intCount, menuPriceCol.�Ƿ���) = "1" And mint���� <> 1 Then
                    If .TextMatrix(intCount, menuPriceCol.�Ƿ��п��) = "0" Then
                        If Val(.TextMatrix(intCount, menuPriceCol.ԭ���ۼ�)) <> Val(.TextMatrix(intCount, menuPriceCol.�����ۼ�)) Then
                            'ʱ��ҩƷ�޿�����
                            dbl��װ = Val(.TextMatrix(intCount, menuPriceCol.��װϵ��))
                            lngҩƷid = Val(.TextMatrix(intCount, menuPriceCol.ҩƷid))
                            dblOldCost = Val(.TextMatrix(intCount, menuPriceCol.ԭ���ۼ�)) / dbl��װ
                            dblNewCost = Val(.TextMatrix(intCount, menuPriceCol.�����ۼ�)) / dbl��װ
                            
                            gstrSQL = "Zl_ҩƷ�۸��¼_Stop("
                            '�۸�����_In
                            gstrSQL = gstrSQL & 1
                            '�ⷿid_In
                            gstrSQL = gstrSQL & ",Null"
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & lngҩƷid
                            '����_In
                            gstrSQL = gstrSQL & ",0"
                            '��ֹ����_In
                            gstrSQL = gstrSQL & "," & "to_date('" & str��ֹʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                            gstrSQL = gstrSQL & ")"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                        
                            gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                            '��������_In
                            gstrSQL = gstrSQL & 1
                            '�۸�����_In
                            gstrSQL = gstrSQL & ",1"
                            '�ⷿid_In
                            gstrSQL = gstrSQL & ",Null"
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & lngҩƷid
                            '����_In
                            gstrSQL = gstrSQL & ",0"
                            
                            'ԭ��_In
                            gstrSQL = gstrSQL & "," & dblOldCost
                            '�ּ�_In
                            gstrSQL = gstrSQL & "," & dblNewCost
                            'ִ������_In
                            gstrSQL = gstrSQL & "," & "to_date('" & strִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                            '����˵��_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '������_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            
                            '���ۻ��ܺ�_In
                            gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
                            
                            gstrSQL = gstrSQL & ")"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                            
                            blnPrice = True
                         End If
                    Else
                        'ʱ��ҩƷ�п�����
                        For n = 1 To vsfStore.rows - 1
                            If Val(.TextMatrix(intCount, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                                lng�ⷿID = Val(vsfStore.TextMatrix(n, menuStoreCol.�ⷿid))
                                lngҩƷid = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))
                                lng���� = Val(vsfStore.TextMatrix(n, menuStoreCol.����))
                                lng��Ӧ��ID = Val(vsfStore.TextMatrix(n, menuStoreCol.��Ӧ��id))
                                str���� = vsfStore.TextMatrix(n, menuStoreCol.����)
                                strЧ�� = IIf(Trim(vsfStore.TextMatrix(n, menuStoreCol.Ч��)) = "", "", vsfStore.TextMatrix(n, menuStoreCol.Ч��))
                                str���� = vsfStore.TextMatrix(n, menuStoreCol.����)
                                dblOldCost = Val(vsfStore.TextMatrix(n, menuStoreCol.ԭ���ۼ�)) / Val(vsfStore.TextMatrix(n, menuStoreCol.��װϵ��))
                                dblNewCost = Val(vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�)) / Val(vsfStore.TextMatrix(n, menuStoreCol.��װϵ��))
                                
                                gstrSQL = "Zl_ҩƷ�۸��¼_Stop("
                                '�۸�����_In
                                gstrSQL = gstrSQL & 1
                                '�ⷿid_In
                                gstrSQL = gstrSQL & "," & lng�ⷿID
                                'ҩƷid_In
                                gstrSQL = gstrSQL & "," & lngҩƷid
                                '����_In
                                gstrSQL = gstrSQL & "," & lng����
                                '��ֹ����_In
                                gstrSQL = gstrSQL & "," & "to_date('" & str��ֹʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                                gstrSQL = gstrSQL & ")"
                                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                                
                                gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                                '��������_In
                                gstrSQL = gstrSQL & 1
                                '�۸�����_In
                                gstrSQL = gstrSQL & ",1"
                                '�ⷿid_In
                                gstrSQL = gstrSQL & "," & lng�ⷿID
                                'ҩƷid_In
                                gstrSQL = gstrSQL & "," & lngҩƷid
                                '����_In
                                gstrSQL = gstrSQL & "," & lng����
                                
                                'ԭ��_In
                                gstrSQL = gstrSQL & "," & dblOldCost
                                '�ּ�_In
                                gstrSQL = gstrSQL & "," & dblNewCost
                                'ִ������_In
                                gstrSQL = gstrSQL & "," & "to_date('" & strִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                                '����˵��_In
                                gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                                '������_In
                                gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                                
                                '���ۻ��ܺ�_In
                                gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
                                '��ҩ��λid_In
                                gstrSQL = gstrSQL & "," & IIf(lng��Ӧ��ID = 0, "Null", lng��Ӧ��ID)
                                '����_In
                                gstrSQL = gstrSQL & ",'" & str���� & "'"
                                'Ч��_In
                                gstrSQL = gstrSQL & "," & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                                '����_In
                                gstrSQL = gstrSQL & ",'" & str���� & "'"
                                
                                gstrSQL = gstrSQL & ")"
                                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                                
                                blnPrice = True
                                blnPrint = True
                            End If
                        Next
                    End If
                End If
            End If
        Next
    End With

    '�ɱ��۵��۴���
    If mint���� = 1 Or mint���� = 2 Then
        If vsfStore.rows > 1 Then
            If vsfStore.TextMatrix(1, menuStoreCol.ҩƷid) <> "" Then
'                lngDouID = 0
'                For n = 1 To vsfStore.rows - 1
'                    If vsfStore.TextMatrix(n, menuStoreCol.ҩƷid) = "" Then Exit For
'
'                    '���δ��˵���
'                    If CheckUnVerify(Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))) = True And Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) <> lngDouID Then
'                        lngDouID = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))
'                        strMsg = vsfStore.TextMatrix(n, menuStoreCol.ҩƷ) & ","
'                        intCount2 = intCount2 + 1
'                        If intCount2 > 3 Then Exit For 'ֻ�ж�3��
'                    End If
'                Next
'
'                If strMsg <> "" Then
'                    If MsgBox(strMsg & "����δ��˵��ݣ������ɱ��ۿ��ܻ���ɲ����" & _
'                        vbCrLf & Space(4) & "�����ȴ���δ��˵��ݡ��Ƿ񻹼������ۣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                        gcnOracle.RollbackTrans
'                        Exit Sub
'                    End If
'                End If

                For n = 1 To vsfStore.rows - 1
                    For i = 1 To vsfPay.rows - 1
                        If vsfPay.TextMatrix(i, 0) = "" Then Exit For
                        If Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) = Val(vsfPay.TextMatrix(i, menuPayCol.ҩƷid)) Then
                            lng�ⷿID = Val(vsfStore.TextMatrix(n, menuStoreCol.�ⷿid))
                            lng��Ӧ��ID = Val(vsfStore.TextMatrix(n, menuStoreCol.��Ӧ��id))
                            lngҩƷid = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))
                            lng���� = Val(vsfStore.TextMatrix(n, menuStoreCol.����))
                            str���� = vsfStore.TextMatrix(n, menuStoreCol.����)
                            strЧ�� = IIf(Trim(vsfStore.TextMatrix(n, menuStoreCol.Ч��)) = "", "", vsfStore.TextMatrix(n, menuStoreCol.Ч��))
                            str���� = vsfStore.TextMatrix(n, menuStoreCol.����)
                            dblOldCost = zlStr.FormatEx(Val(vsfStore.TextMatrix(n, menuStoreCol.ԭ�ɱ���)) / Val(vsfStore.TextMatrix(n, menuStoreCol.��װϵ��)), gtype_UserDrugDigits.Digit_�ɱ���, , True)
                            dblNewCost = zlStr.FormatEx(Val(vsfStore.TextMatrix(n, menuStoreCol.�ֳɱ���)) / Val(vsfStore.TextMatrix(n, menuStoreCol.��װϵ��)), gtype_UserDrugDigits.Digit_�ɱ���, , True)
                            Str��Ʊ�� = vsfPay.TextMatrix(i, menuPayCol.��Ʊ��)
                            str��Ʊ���� = Format(vsfPay.TextMatrix(i, menuPayCol.��Ʊ����), "yyyy-mm-dd")
                            dbl��Ʊ��� = Val(vsfPay.TextMatrix(i, menuPayCol.��Ʊ���))
                            
'                            gstrSQL = "Zl_�ɱ��۵�����Ϣ_Insert(" & IIf(lng��Ӧ��ID = 0, "Null", lng��Ӧ��ID) & "," & lng�ⷿID & "," & lngҩƷID & "," & lng���� & ",'" & str���� & "'" & _
'                                    "," & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & str���� & "',Null," & dblOldCost & ", " & dblNewCost & "," & _
'                                    IIf(Str��Ʊ�� <> "", "'" & Str��Ʊ�� & "'", "NULL") & "," & IIf(str��Ʊ���� = "", "Null", "to_date('" & Format(str��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ", " & dbl��Ʊ��� & "," & IIf(mblnӦ����¼ = True, 1, 0) & "," & txtNo.Text & ")"
'                            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                            
                            gstrSQL = "Zl_ҩƷ�۸��¼_Stop("
                            '�۸�����_In
                            gstrSQL = gstrSQL & 2
                            '�ⷿid_In
                            gstrSQL = gstrSQL & "," & lng�ⷿID
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & lngҩƷid
                            '����_In
                            gstrSQL = gstrSQL & "," & lng����
                            '��ֹ����_In
                            gstrSQL = gstrSQL & "," & "to_date('" & str��ֹʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                            gstrSQL = gstrSQL & ")"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                            
'                            ��������_In   In ҩƷ�۸��¼.��������%Type,
'                            �۸�����_In   In ҩƷ�۸��¼.�۸�����%Type,
'                            �ⷿid_In     In ҩƷ�۸��¼.�ⷿid%Type,
'                            ҩƷid_In     In ҩƷ�۸��¼.ҩƷid%Type,
'                            ����_In       In ҩƷ�۸��¼.����%Type := Null,
'
'                            ԭ��_In       In ҩƷ�۸��¼.ԭ��%Type := Null,
'                            �ּ�_In       In ҩƷ�۸��¼.�ּ�%Type := Null,
'                            ִ������_In   In ҩƷ�۸��¼.ִ������%Type := Null,
'                            ����˵��_In   In ҩƷ�۸��¼.����˵��%Type := Null,
'                            ������_In     In ҩƷ�۸��¼.������%Type := Null,
'
'                            ���ۻ��ܺ�_In In ҩƷ�۸��¼.���ۻ��ܺ�%Type := Null,
'                            ��ҩ��λid_In In ҩƷ�۸��¼.��ҩ��λid%Type := Null,
'                            ����_In       In ҩƷ�۸��¼.����%Type := Null,
'                            Ч��_In       In ҩƷ�۸��¼.Ч��%Type := Null,
'                            ����_In       In ҩƷ�۸��¼.����%Type := Null
'
'                            ���Ч��_In   In ҩƷ�۸��¼.���Ч��%Type := Null,
'                            ��Ʊ��_In     In ҩƷ�۸��¼.��Ʊ��%Type := Null,
'                            ��Ʊ����_In   In ҩƷ�۸��¼.��Ʊ����%Type := Null,
'                            ��Ʊ���_In   In ҩƷ�۸��¼.��Ʊ���%Type := Null,
'                            Ӧ����䶯_In In ҩƷ�۸��¼.Ӧ����䶯%Type := 0
  
                            
                            gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                            '��������_In
                            gstrSQL = gstrSQL & 1
                            '�۸�����_In
                            gstrSQL = gstrSQL & ",2"
                            '�ⷿid_In
                            gstrSQL = gstrSQL & "," & lng�ⷿID
                            'ҩƷid_In
                            gstrSQL = gstrSQL & "," & lngҩƷid
                            '����_In
                            gstrSQL = gstrSQL & "," & lng����
                            
                            'ԭ��_In
                            gstrSQL = gstrSQL & "," & dblOldCost
                            '�ּ�_In
                            gstrSQL = gstrSQL & "," & dblNewCost
                            'ִ������_In
                            gstrSQL = gstrSQL & "," & "to_date('" & strִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                            '����˵��_In
                            gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                            '������_In
                            gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                            
                            '���ۻ��ܺ�_In
                            gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
                            '��ҩ��λid_In
                            gstrSQL = gstrSQL & "," & IIf(lng��Ӧ��ID = 0, "Null", lng��Ӧ��ID)
                            '����_In
                            gstrSQL = gstrSQL & ",'" & str���� & "'"
                            'Ч��_In
                            gstrSQL = gstrSQL & "," & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                            '����_In
                            gstrSQL = gstrSQL & ",'" & str���� & "'"
                            
                            '���Ч��_In
                            gstrSQL = gstrSQL & ",Null"
                            '��Ʊ��_In
                            gstrSQL = gstrSQL & ",'" & Str��Ʊ�� & "'"
                            '��Ʊ����_In
                            gstrSQL = gstrSQL & "," & IIf(str��Ʊ���� = "", "Null", "to_date('" & Format(str��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                            '��Ʊ���_In
                            gstrSQL = gstrSQL & "," & dbl��Ʊ���
                            'Ӧ����䶯_In
                            gstrSQL = gstrSQL & "," & IIf(mblnӦ����¼ = True, 1, 0)
                            
                            gstrSQL = gstrSQL & ")"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                            
                            blnCost = True
                            blnPrint = True
                        End If
                    Next
                Next
            End If
        End If
    End If

    '�޿��ʱ�����ɱ���
    If mint���� = 1 Or mint���� = 2 Then
        With Me.vsfPrice
            For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
                If .TextMatrix(intCount, menuPriceCol.�Ƿ��п��) = "0" And Val(.TextMatrix(intCount, menuPriceCol.ԭ�ɱ���)) <> Val(.TextMatrix(intCount, menuPriceCol.�ֳɱ���)) Then
                    dbl��װ = Val(.TextMatrix(intCount, menuPriceCol.��װϵ��))

                    lngҩƷid = Val(.TextMatrix(intCount, menuPriceCol.ҩƷid))
                    dblOldCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.ԭ�ɱ���)) / dbl��װ, gtype_UserDrugDigits.Digit_�ɱ���))
                    dblNewCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.�ֳɱ���)) / dbl��װ, gtype_UserDrugDigits.Digit_�ɱ���))

'                    gstrSQL = "Zl_�ɱ��۵�����Ϣ_Insert(Null,Null," & lngҩƷID & ",0,Null,Null,Null,Null," & dblOldCost & ", " & dblNewCost & ",NULL,Null,0,0, " & txtNO.Text & ")"
'                    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                    
                    gstrSQL = "Zl_ҩƷ�۸��¼_Stop("
                    '�۸�����_In
                    gstrSQL = gstrSQL & 2
                    '�ⷿid_In
                    gstrSQL = gstrSQL & ",Null"
                    'ҩƷid_In
                    gstrSQL = gstrSQL & "," & lngҩƷid
                    '����_In
                    gstrSQL = gstrSQL & ",0"
                    '��ֹ����_In
                    gstrSQL = gstrSQL & "," & "to_date('" & str��ֹʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                    
                    gstrSQL = "Zl_ҩƷ�۸��¼_Insert("
                    '��������_In
                    gstrSQL = gstrSQL & 1
                    '�۸�����_In
                    gstrSQL = gstrSQL & ",2"
                    '�ⷿid_In
                    gstrSQL = gstrSQL & ",Null"
                    'ҩƷid_In
                    gstrSQL = gstrSQL & "," & lngҩƷid
                    '����_In
                    gstrSQL = gstrSQL & ",0"
                    
                    'ԭ��_In
                    gstrSQL = gstrSQL & "," & dblOldCost
                    '�ּ�_In
                    gstrSQL = gstrSQL & "," & dblNewCost
                    'ִ������_In
                    gstrSQL = gstrSQL & "," & "to_date('" & strִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS')"
                    '����˵��_In
                    gstrSQL = gstrSQL & ",'" & Me.txtSummary.Text & "'"
                    '������_In
                    gstrSQL = gstrSQL & ",'" & Trim(Me.txtValuer.Text) & "'"
                    
                    '���ۻ��ܺ�_In
                    gstrSQL = gstrSQL & ",'" & txtNO.Text & "'"
                    
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                    
                    blnCost = True
                End If
            Next
        End With
    End If

    '����ִ��
    If mint���� = 1 Then
        '�����ɱ��۵���ʱ
        If optʱ��(0).Value = True Then
            With Me.vsfPrice
                For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
                    gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & Val(.TextMatrix(intCount, menuPriceCol.ҩƷid)) & "," & typeAdjust.AdjustCost & " )"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                Next
            End With
        End If
    Else
        '���ۼ�
        If optʱ��(0).Value = True Then
            ArrayID = Split(strDrugs, ",")
            For intCount = 0 To UBound(ArrayID)
                gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & ArrayID(intCount) & "," & IIf(mint���� = 0, typeAdjust.AdjustPrice, typeAdjust.AdjustPriceAndCost) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            Next
        End If
    End If

    '����ָ���۸�
    With Me.vsfPrice
        For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
            dbl��װ = Val(.TextMatrix(intCount, menuPriceCol.��װϵ��))

            '����ָ�����ۼ�
            If Val(.TextMatrix(intCount, menuPriceCol.ԭָ���ۼ�)) <> Val(.TextMatrix(intCount, menuPriceCol.��ָ���ۼ�)) And Val(.TextMatrix(intCount, menuPriceCol.��ָ���ۼ�)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.��ָ���ۼ�)) / dbl��װ, mintSalePriceDigit))

                gstrSQL = "zl_ҩƷĿ¼_UpdateCustom(" & Val(.TextMatrix(intCount, menuPriceCol.ҩƷid)) & ",'ָ�����ۼ�=" & strUpdate & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If

            '���²ɹ��޼�
            If Val(.TextMatrix(intCount, menuPriceCol.ԭ�ɹ��޼�)) <> Val(.TextMatrix(intCount, menuPriceCol.�ֲɹ��޼�)) And Val(.TextMatrix(intCount, menuPriceCol.�ֲɹ��޼�)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.�ֲɹ��޼�)) / dbl��װ, mintSalePriceDigit))

                gstrSQL = "zl_ҩƷĿ¼_UpdateCustom(" & Val(.TextMatrix(intCount, menuPriceCol.ҩƷid)) & ",'ָ��������=" & strUpdate & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
        Next
    End With

    '�������ۻ��ܼ�¼
    If blnPrice = True And blnCost = True Then
        intUpdateModel = 2
    ElseIf blnPrice = True And blnCost = False Then
        intUpdateModel = 0
    ElseIf blnPrice = False And blnCost = True Then
        intUpdateModel = 1
    End If

    gstrSQL = "Zl_���ۻ��ܼ�¼_Insert(" & txtNO.Text & "," & intUpdateModel & ","
    If optʱ��(0).Value = True Then
        gstrSQL = gstrSQL & "sysdate" & ","
    Else
        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    End If
    gstrSQL = gstrSQL & IIf(txtSummary.Text = "", "Null", "'" & txtSummary.Text & "'") & ",0,'" & UserInfo.�û����� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)

    gcnOracle.CommitTrans

    If blnPrint = True Then
        If MsgBox("����Ҫ��ӡ����֪ͨ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1333", Me, "NO=" & txtNO.Text, "��װ��λ=" & mintUnit, 2)
        End If
    End If

    '����б�������
    With vsfPrice
        .rows = 2
        For intCol = 0 To .Cols - 1
            .TextMatrix(1, intCol) = ""
        Next
    End With
    vsfStore.rows = 1
    vsfPay.rows = 1
    txtNO.Text = ""
    txtSummary.Text = ""

    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Function CheckUnVerify(ByVal lngҩƷid As Long) As Boolean
    '���ҩƷ�Ƿ����δ��˵���
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ҩƷid = [1] And Rownum = 1 And ������� Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ҩƷ�Ƿ����δ��˵���", lngҩƷid)

    If rsTemp.RecordCount > 0 Then
        CheckUnVerify = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function checkNotExecutePrice(Optional ByVal lngDrugID As Long = 0, Optional ByRef strInfo As String) As Boolean
    '���� ������Ƿ����δִ�еļ۸�
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long, IntCheck As Integer

    Err = 0
    On Error GoTo ErrHand

    If lngDrugID = 0 Then
        'ѭ���ж�����ҩƷ
        For IntCheck = 1 To vsfPrice.rows - 1
            LngmediIDThis = Val(vsfPrice.TextMatrix(IntCheck, menuPriceCol.ҩƷid))
            If LngmediIDThis <> 0 Then
                If mint���� = 0 Or mint���� = 2 Then
                    '�ж��Ƿ���δִ�е���ʷ�۸�
                    gstrSQL = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And ִ������ > Sysdate And �շ�ϸĿID=[1]" & _
                            GetPriceClassString("")
                    
                    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, LngmediIDThis)

                    With RecCheck
                        If Not .EOF Then
                            If Not IsNull(!Records) Then
                                If !Records <> 0 Then
                                    strInfo = "ҩƷ" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.ҩƷ) & "����δִ�м۸�δִ��ҩƷ���ܵ��ۣ�"
                                    checkNotExecutePrice = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End With
                End If

                If mint���� = 1 Or mint���� = 2 Then
                    '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
                    gstrSQL = "Select 1 From ҩƷ�۸��¼ Where �۸�����=2 And ��¼״̬=0 And ҩƷid = [1] And Rownum < 2 "
                    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, LngmediIDThis)

                    If RecCheck.RecordCount > 0 Then
                        strInfo = "ҩƷ" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.ҩƷ) & "����δִ�гɱ��ۣ�δִ��ҩƷ���ܵ��ۣ�"
                        checkNotExecutePrice = True
                        Exit Function
                    End If
                End If
            End If
        Next
    Else
        If mint���� = 0 Or mint���� = 2 Then
            '�ж��Ƿ���δִ�е���ʷ�۸�
            gstrSQL = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And ִ������ > Sysdate And �շ�ϸĿID=[1]" & _
                    GetPriceClassString("")
            
            Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID, gstrPriceClass)

            With RecCheck
                If Not .EOF Then
                    If Not IsNull(!Records) Then
                        If !Records <> 0 Then
                            strInfo = "������δִ�е��ۼ۵��ۼ�¼��δִ��ҩƷ���ܵ��ۣ�"
                            checkNotExecutePrice = True
                            Exit Function
                        End If
                    End If
                End If
            End With
        End If

        If mint���� = 1 Or mint���� = 2 Then
            '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
            gstrSQL = "Select 1 From ҩƷ�۸��¼ Where �۸�����=2 And ��¼״̬=0 And ҩƷid = [1] And Rownum < 2 "
            Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

            If RecCheck.RecordCount > 0 Then
                strInfo = "������δִ�еĳɱ��۵��ۣ�δִ��ҩƷ���ܵ��ۣ�"
                checkNotExecutePrice = True
                Exit Function
            End If
        End If
    End If


    checkNotExecutePrice = False
    Exit Function
ErrHand:
    Call ErrCenter
    Call SaveErrLog
    Me.vsfPrice.SetFocus

End Function

Private Function CheckPrice() As Boolean
    Dim IntCheck As Integer
    Dim n As Integer
    Dim strTmp As String
    Dim bln�޿�� As Boolean
    Dim dbl��װ As Double
    Dim bln���޿�� As Boolean
    Dim lngDouID As Long
    Dim strMsg As String '��¼��ʾ��Ϣ
    Dim intCount2 As Integer '��������
    
    '����ִ�м۸��Ƿ���ȷ
    '�Լ�������Ŀ��ͬ��������ּ��Ƿ���ԭ����ͬ
    CheckPrice = False
    With vsfPrice
        For IntCheck = 1 To .rows - 1
            If Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid)) <> 0 Then
                If Not IsNumeric(Trim(.TextMatrix(IntCheck, menuPriceCol.�����ۼ�))) Then
                    MsgBox "��" & IntCheck & "�е�ҩƷ�ۼ��к��зǷ��ַ���", vbInformation, gstrSysName
                    .Row = IntCheck
                    .Col = menuPriceCol.�����ۼ�
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If
                
                '���۸��Ƿ�Ϊ��
                If .TextMatrix(IntCheck, menuPriceCol.�����ۼ�) = "" Or .TextMatrix(IntCheck, menuPriceCol.ԭ���ۼ�) = "" Or .TextMatrix(IntCheck, menuPriceCol.�ֳɱ���) = "" Or .TextMatrix(IntCheck, menuPriceCol.ԭ�ɱ���) = "" Then
                    MsgBox "��" & IntCheck & "�е�ҩƷ�м۸�Ϊ�գ�����ִ�е��ۣ�", vbInformation, gstrSysName
                    .Row = IntCheck
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If
                For n = 1 To vsfStore.rows - 1
                    If Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                        If vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�) = "" Or vsfStore.TextMatrix(n, menuStoreCol.ԭ���ۼ�) = "" Or vsfStore.TextMatrix(n, menuStoreCol.�ֳɱ���) = "" Or vsfStore.TextMatrix(n, menuStoreCol.ԭ�ɱ���) = "" Then
                            MsgBox "��" & IntCheck & "�е�ҩƷ�м۸�Ϊ�գ�����ִ�е��ۣ�", vbInformation, gstrSysName
                            .Row = IntCheck
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                    End If
                Next
                
                '����ۼ��Ƿ���ͬ
                If mint���� = 0 Or mint���� = 2 Then
                    strTmp = ""
                    bln���޿�� = False
                    dbl��װ = Val(.TextMatrix(IntCheck, menuPriceCol.��װϵ��))
                    If .TextMatrix(IntCheck, menuPriceCol.�Ƿ���) = "1" Then
                        For n = 1 To vsfStore.rows - 1
                            If Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                                bln���޿�� = True
                                If InStr(1, "|" & strTmp, "|" & vsfStore.TextMatrix(n, menuStoreCol.����) & ",") = 0 And vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�) <> vsfStore.TextMatrix(n, menuStoreCol.ԭ���ۼ�) Then
                                    strTmp = strTmp & IIf(strTmp = "", "", "|") & vsfStore.TextMatrix(n, menuStoreCol.����) & "," & vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�) / dbl��װ
                                End If
                            End If
                        Next
                        If strTmp = "" And bln���޿�� = True Then
                            MsgBox "��" & IntCheck & "�е�ҩƷ�����ۼ���ԭ���ۼ���ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                            .Row = IntCheck
                            .Col = menuPriceCol.�����ۼ�
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                        If bln���޿�� = False And .TextMatrix(IntCheck, menuPriceCol.�����ۼ�) = .TextMatrix(IntCheck, menuPriceCol.ԭ���ۼ�) Then
                            MsgBox "��" & IntCheck & "�е�ҩƷ�����ۼ���ԭ���ۼ���ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                            .Row = IntCheck
                            .Col = menuPriceCol.�����ۼ�
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                    End If
                    If .TextMatrix(IntCheck, menuPriceCol.�Ƿ���) <> "1" And .TextMatrix(IntCheck, menuPriceCol.�����ۼ�) = .TextMatrix(IntCheck, menuPriceCol.ԭ���ۼ�) Then
                        MsgBox "��" & IntCheck & "�е�ҩƷ�����ۼ���ԭ���ۼ���ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.�����ۼ�
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                End If
                
                '���ɱ����Ƿ���ͬ
                If mint���� = 1 Or mint���� = 2 Then
                    bln���޿�� = False
                    strTmp = ""
                    For n = 1 To vsfStore.rows - 1
                        If Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                            bln���޿�� = True
                            If vsfStore.TextMatrix(n, menuStoreCol.�ֳɱ���) <> vsfStore.TextMatrix(n, menuStoreCol.ԭ�ɱ���) Then
                                strTmp = "�����ɱ���"
                            End If
                        End If
                    Next
                    If bln���޿�� = True And strTmp = "" Then
                        MsgBox "��" & IntCheck & "�е�ҩƷ�ֳɱ�����ԭ�ɱ�����ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.�ֳɱ���
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                    If bln���޿�� = False And .TextMatrix(IntCheck, menuPriceCol.�ֳɱ���) = .TextMatrix(IntCheck, menuPriceCol.ԭ�ɱ���) Then
                        MsgBox "��" & IntCheck & "�е�ҩƷ�ֳɱ�����ԭ�ɱ�����ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.�ֳɱ���
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                End If
                
                '���۹��������ۺ��ۼۺͳɱ����Ƿ�һ��
                If gtype_UserSysParms.P275_���۹���ģʽ > 0 Then
                    If IsPriceAdjustMod(Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid))) = True Then
                        If .TextMatrix(IntCheck, menuPriceCol.�Ƿ��п��) = 0 Then
                            '�޿�棬ֱ�ӱȽϼ۸���е��ۼۺͳɱ���

                            If Val(.TextMatrix(IntCheck, menuPriceCol.�����ۼ�)) <> Val(.TextMatrix(IntCheck, menuPriceCol.�ֳɱ���)) Then
                                MsgBox "��" & IntCheck & "�еĶ���ҩƷ���������۹������ۼ۱���Ϳ��ɱ���һ�£�", vbInformation, gstrSysName
                                .Row = IntCheck
                                .Col = menuPriceCol.�����ۼ�
                                vsfPrice.SetFocus
                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                .TopRow = IntCheck
                                Exit Function
                            End If
       
                        Else
                            If .TextMatrix(IntCheck, menuPriceCol.�Ƿ���) = "0" Then
                                '���ۣ�������б��е������ۼ��Ƿ�Ϳ����е����³ɱ���һ��
                                For n = 1 To vsfStore.rows - 1
                                    If Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                                        If mint���� = 0 Then
                                            '�����ۼ۷�ʽ
                                            If Val(.TextMatrix(IntCheck, menuPriceCol.�����ۼ�)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.ԭ�ɱ���)) Then
                                                MsgBox "��" & IntCheck & "�еĶ���ҩƷ���������۹������ۼ۱���Ϳ��ɱ���һ�£�", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.�����ۼ�
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        ElseIf mint���� = 1 Then
                                            '�����ɱ��۷�ʽ
                                            If Val(.TextMatrix(IntCheck, menuPriceCol.ԭ���ۼ�)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.�ֳɱ���)) Then
                                                MsgBox "��" & IntCheck & "�еĶ���ҩƷ���������۹����³ɱ��۱�����ۼ�һ�£�", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.�����ۼ�
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        Else
                                            '�ۼۺͳɱ���һ�����ʽ
                                            If Val(.TextMatrix(IntCheck, menuPriceCol.�����ۼ�)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.�ֳɱ���)) Then
                                                MsgBox "��" & IntCheck & "�еĶ���ҩƷ���������۹������ۼ۱���Ϳ���³ɱ���һ�£�", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.�����ۼ�
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                'ʱ�ۣ�������б��е������ۼ��Ƿ�Ϳ����е����³ɱ���һ��
                                For n = 1 To vsfStore.rows - 1
                                    If Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                                        If mint���� = 0 Then
                                            '�����ۼ۷�ʽ
                                            If Val(vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.ԭ�ɱ���)) Then
                                                MsgBox "��" & IntCheck & "�е�ʱ��ҩƷ���������۹������ۼ۱���Ϳ��ɱ���һ�£�", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.�����ۼ�
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        ElseIf mint���� = 1 Then
                                            '�����ɱ��۷�ʽ
                                            If Val(vsfStore.TextMatrix(n, menuStoreCol.�ֳɱ���)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.ԭ���ۼ�)) Then
                                                MsgBox "��" & IntCheck & "�е�ʱ��ҩƷ���������۹����³ɱ��۱�����ۼ�һ�£�", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.�����ۼ�
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        Else
                                            '�ۼۺͳɱ���һ�����ʽ
                                            If Val(vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�)) <> Val(vsfStore.TextMatrix(n, menuStoreCol.�ֳɱ���)) Then
                                                MsgBox "��" & IntCheck & "�е�ʱ��ҩƷ���������۹������ۼ۱���Ϳ���³ɱ���һ�£�", vbInformation, gstrSysName
                                                .Row = IntCheck
                                                .Col = menuPriceCol.�����ۼ�
                                                vsfPrice.SetFocus
                                                .Select IntCheck, 0, IntCheck, .Cols - 1
                                                .TopRow = IntCheck
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                Next
        
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With

    '���δ��˵���
    If vsfStore.rows > 1 And (mint���� = 1 Or mint���� = 2) Then
        If vsfStore.TextMatrix(1, menuStoreCol.ҩƷid) <> "" Then
            lngDouID = 0
            For n = 1 To vsfStore.rows - 1
                If vsfStore.TextMatrix(n, menuStoreCol.ҩƷid) = "" Then Exit For
    
                If CheckUnVerify(Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))) = True And Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) <> lngDouID Then
                    lngDouID = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))
                    strMsg = strMsg & vsfStore.TextMatrix(n, menuStoreCol.ҩƷ) & ","
                    intCount2 = intCount2 + 1
                    If intCount2 > 3 Then Exit For 'ֻ�ж�3��
                End If
            Next
    
            If strMsg <> "" Then
                If MsgBox(strMsg & "����δ��˵��ݣ������ɱ��ۿ��ܻ���ɲ����" & _
                    vbCrLf & Space(4) & "�����ȴ���δ��˵��ݡ��Ƿ񻹼������ۣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
                
    CheckPrice = True
End Function


Private Sub cmdPriceMethod_Click()
    If txt��Ӧ��.Tag = "" Then
        Me.txt��Ӧ��.Tag = "0|"
    End If
    picOtherSelect.Visible = True
End Sub

Private Sub PrintStore()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    If vsfStore.rows = 1 Then
        MsgBox "û�п��䶯��¼��", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(Me.vsfStore.TextMatrix(1, menuStoreCol.�ⷿ)) = "" Then Exit Sub

    objPrint.Title.Text = "���ۿ��䶯��"

    Set objRow = New zlTabAppRow
    objRow.Add "����˵��:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "ִ��ʱ��:" & Format(IIf(optʱ��(0).Value = True, Sys.Currentdate, Me.dtpRunDate.Value), "yyyy��MM��DD�� HH:mm:ss")
    objRow.Add "������:" & Me.txtValuer.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡʱ��:" & Format(Sys.Currentdate, "yyyy��MM��DD�� HH:mm:ss")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = Me.vsfStore.Object
    objPrint.PageFooter = 2

    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    gstrSQL = "Select ����,����,����,id" & _
        " From ��Ӧ��" & _
        " where ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
        " Order By ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ӧ����Ϣ")
    If rsTemp.EOF Then
        MsgBox "���ʼ����Ӧ�̣��ֵ������", vbInformation, gstrSysName
        Exit Sub
    End If

    With Me.mshProvider
        .Left = chk��Ӧ��.Left
        .Top = txt��Ӧ��.Top + txt��Ӧ��.Height
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
        .Row = 1: .ColSel = .Cols - 1
        .ZOrder 0: .Visible = True: .SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Activate()
    If mblnLoad = False Then
        vsfPrice.SetFocus
    End If
    If mBlnClick = False Then
        vsfPrice.Row = 1
        vsfPrice.Col = menuPriceCol.ҩƷ
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        picOtherSelect.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim StrToday As String
    Dim intUnitTemp As Integer
    Dim blnOldAjuset As Boolean '�ж�35.70֮ǰ���ϵ���ģʽ
    Dim rsTemp As ADODB.Recordset
    
    Me.Height = 768 * 15
    Me.Width = 1024 * 15
    '��ȡ���õĵ�λ
    mintUnit = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, 1333, "1"))
    mstrPrivs = GetPrivFunc(glngSys, 1333)
    Select Case mintUnit
        Case 0 'ҩ��
            intUnitTemp = 4
        Case 2 'סԺ
            intUnitTemp = 3
        Case 1 '����
            intUnitTemp = 2
        Case 3 '�ۼ�
            intUnitTemp = 1
    End Select
    '��ȡ������λ����
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
    mintNumberDigit = GetDigitTiaoJia(1, 3, intUnitTemp)
    mintMoneyDigit = GetDigitTiaoJia(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    mintSalePriceDigit = GetDigitTiaoJia(1, 2, 1)
    '��ʼ��ʱ��Ϊ��ǰʱ��+1��
    StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")

    If mintModal = 0 Then '������ʱ����Сʱ������Ϊ��ǰʱ��+1��
        Me.dtpRunDate.MinDate = DateAdd("s", 1, CDate(StrToday))
    End If
    Me.dtpRunDate.Value = DateAdd("d", 1, CDate(StrToday))

    mblnʱ��ҩƷ�����ε��� = Val(zlDatabase.GetPara("ʱ��ҩƷ�����ε���", glngSys, 1333, 0))
    mbln�ɱ��۰��ⷿ���ε��� = Val(zlDatabase.GetPara("�ɱ��۰��ⷿ���ε���", glngSys, 1333, 0))
    mbln�ּ���ʾ = Val(zlDatabase.GetPara("�޼���ʾ", glngSys, 1333, 1))

    txtValuer.Text = UserInfo.�û�����  'gstrUserName

    txtNO.Text = IIf(mintModal = 0, "", mstr���ۻ��ܺ�)
    If mintModal = 0 Then
        lblNO.Visible = False
        txtNO.Visible = False
    End If

    Call initComboBox '��ʼ�������ؼ�
    If mintModal = 1 Then '�޸�
        If (InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") = 0) Or (InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") > 0) Then
            cboPriceMethod.ListIndex = 0
        ElseIf (InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") > 0) Then
            cboPriceMethod.ListIndex = mintMethod
        End If
    ElseIf mintModal = 2 Then '����
        cboPriceMethod.ListIndex = mintMethod
    End If

    Call initCommandBars
    
    Call InitTabControl
    Call InitVsfGridFlex

    Call RestoreWinState(Me, App.ProductName, MStrCaption)
    If mblnӦ����¼ = False Then
        TabCtlDetails.Item(1).Visible = False
    End If
    
    If mintModal <> 0 Then
        gstrSQL = "Select 1 from ҩƷ�۸��¼ where ���ۻ��ܺ�=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ۻ��ܺ�)
        
        '�ж��Ƿ�35.70����ǰ���ݣ������ݵ����Ϸ�������ʾ��ֻ���ۼ�Ҳ����������ʾ
        If Not rsTemp.EOF Then
            Call initGrid
        Else
            Call initGrid_Old
        End If
    End If

    If mintModal = 2 Then '����
        Dim cbrControls As CommandBarControl
        Set cbrControls = cbsMain.FindControl(, mconMenu_Save)
        cbrControls.Enabled = False
        Set cbrControls = cbsMain.FindControl(, mconMenu_BatchSelect)
        cbrControls.Enabled = False
        Set cbrControls = cbsMain.FindControl(, mconMenu_ClearAll)
        cbrControls.Enabled = False
        Set cbrControls = cbsMain.FindControl(, mconMenu_Find)
        cbrControls.Enabled = False
    
        cboPriceMethod.Enabled = False
        cmdPriceMethod.Enabled = False
        optʱ��(0).Enabled = False
        optʱ��(1).Enabled = False
        dtpRunDate.Enabled = False
        cbo�ۼۼ��㷽ʽ.Visible = False
        lbl���۷�ʽ.Visible = False
        chkAotuCost.Visible = False
        chkAotuCost.Enabled = False
        chkAutoPay.Enabled = False
        txtSummary.Enabled = False

        vsfPrice.Cell(flexcpBackColor, 1, 0, vsfPrice.rows - 1, vsfPrice.Cols - 1) = mconlngColor
        If vsfStore.rows > 1 Then
            vsfStore.Cell(flexcpBackColor, 1, 0, vsfStore.rows - 1, vsfStore.Cols - 1) = mconlngColor
        End If
        If vsfPay.rows > 1 Then
            vsfPay.Cell(flexcpBackColor, 0, 0, vsfPay.rows - 1, vsfPay.Cols - 1) = mconlngColor
        End If
    End If
    mblnLoad = True
End Sub

Private Sub initGrid_Old()
    '������޸Ļ��߲�������ȡ��Ӧ�ļ�¼����䵽�����
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim i As Long
    Dim lngDrugID As Long
    Dim db��װϵ�� As Double
    Dim strUnit As String
    Dim StrToday As String
    Dim rs���� As ADODB.Recordset

    On Error GoTo errHandle
    '���۷�ʽ 0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
    If mintMethod = 0 Then
        gstrSQL = "Select Distinct p.ԭ��id, i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "                s.�ӳ���/100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, i.���� As ����, i.���㵥λ As ��λ," & vbNewLine & _
            "                s.���ﵥλ, s.�����װ, s.סԺ��λ, s.סԺ��װ, s.ҩ�ⵥλ, Nvl(s.ҩ���װ, 1) ҩ���װ, s.�ɱ��� As ԭ�ɱ���, s.�ɱ��� As �³ɱ���, p.ԭ��, p.�ּ�," & vbNewLine & _
            "                p.������Ŀid, p.������, p.����˵��, s.���������, To_Char(a.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ҩƷid," & vbNewLine & _
            "                Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select s.ҩƷid From ҩƷ��� s where s.����=1 And Not (zl_fun_getbatchpro(s.�ⷿid,s.ҩƷid)=1 And Nvl(S.����,0) = 0 And S.�������� < 0 And S.ʵ������ = 0 And S.ʵ�ʽ�� = 0 And S.ʵ�ʲ�� = 0)) K, ���ۻ��ܼ�¼ A, �շ���Ŀ���� B, ҩƷ��� S, �շ���ĿĿ¼ I, �շѼ�Ŀ P" & vbNewLine & _
            "Where a.���ۺ� = p.���ۻ��ܺ� And b.�շ�ϸĿid(+) = s.ҩƷid And s.ҩƷid = i.Id And i.Id = k.ҩƷid(+) And i.Id = p.�շ�ϸĿid And" & vbNewLine & _
            "      p.���ۻ��ܺ� = [1] And a.���� = 0 And b.����(+) = 3 And a.���ۺ� = [1] " & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            "Order By ҩƷid"
    ElseIf mintMethod = 1 Then
        gstrSQL = "Select Distinct i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "                s.�ӳ���/100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, i.���� As ����, i.���㵥λ As ��λ," & vbNewLine & _
            "                s.���ﵥλ, s.�����װ, s.סԺ��λ, s.סԺ��װ, s.ҩ�ⵥλ, Nvl(s.ҩ���װ, 1) ҩ���װ, m.ԭ�ɱ���, m.�³ɱ���, p.�ּ� as ԭ��, p.�ּ�, p.������Ŀid," & vbNewLine & _
            "                a.������ As ������, a.˵�� As ����˵��, s.���������, To_Char(m.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ҩƷid," & vbNewLine & _
            "                Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select Min(ԭ�ɱ���) As ԭ�ɱ���, Min(�³ɱ���) As �³ɱ���, min(����) as ����,���ۻ��ܺ�,ҩƷid,min(ִ������) as ִ������ From �ɱ��۵�����Ϣ Where ���ۻ��ܺ� = [1] Group By ���ۻ��ܺ�,ҩƷid) M, (Select s.ҩƷid From ҩƷ��� s where s.����=1 And Not (zl_fun_getbatchpro(s.�ⷿid,s.ҩƷid)=1 And Nvl(S.����,0) = 0 And S.�������� < 0 And S.ʵ������ = 0 And S.ʵ�ʽ�� = 0 And S.ʵ�ʲ�� = 0)) K, ���ۻ��ܼ�¼ A, �շ���Ŀ���� B, ҩƷ��� S, �շ���ĿĿ¼ I, �շѼ�Ŀ P" & vbNewLine & _
            "Where m.���ۻ��ܺ�(+) = a.���ۺ� And b.�շ�ϸĿid(+) = s.ҩƷid And s.ҩƷid = i.Id And i.Id = k.ҩƷid(+) And m.ҩƷid = i.Id And" & vbNewLine & _
            "      i.Id = p.�շ�ϸĿid And Sysdate Between p.ִ������ And p.��ֹ���� And m.���ۻ��ܺ� = [1] And a.���� = 0 And b.����(+) = 3 And" & vbNewLine & _
            "      a.���ۺ� = [1] " & IIf(mintModal = 2, "", " And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            "Order By ҩƷid"
    ElseIf mintMethod = 2 Then
        gstrSQL = "Select distinct p.ԭ��id, i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "       s.�ӳ���/100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, i.���� As ����, i.���㵥λ As ��λ, s.���ﵥλ," & vbNewLine & _
            "       s.�����װ, s.סԺ��λ, s.סԺ��װ, s.ҩ�ⵥλ, Nvl(s.ҩ���װ, 1) ҩ���װ, m.ԭ�ɱ���, m.�³ɱ���, p.ԭ��, p.�ּ�, p.������Ŀid, p.������, p.����˵��, s.���������," & vbNewLine & _
            "       To_Char(p.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ҩƷid, Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select ҩƷid,Min(ԭ�ɱ���) As ԭ�ɱ���, Min(�³ɱ���) As �³ɱ���, min(����) as ����,���ۻ��ܺ� From �ɱ��۵�����Ϣ Where ���ۻ��ܺ� = [1] Group By ҩƷid,���ۻ��ܺ�) M, �շѼ�Ŀ P, ���ۻ��ܼ�¼ A, (Select s.ҩƷid From ҩƷ��� s where s.����=1 And Not (zl_fun_getbatchpro(s.�ⷿid,s.ҩƷid)=1 And Nvl(S.����,0) = 0 And S.�������� < 0 And S.ʵ������ = 0 And S.ʵ�ʽ�� = 0 And S.ʵ�ʲ�� = 0)) K, �շ���Ŀ���� B, ҩƷ��� S, �շ���ĿĿ¼ I" & vbNewLine & _
            "Where m.���ۻ��ܺ� = a.���ۺ� and m.ҩƷid=i.id And p.���ۻ��ܺ� = a.���ۺ� And p.�շ�ϸĿid = k.ҩƷid(+) And p.�շ�ϸĿid = b.�շ�ϸĿid(+) And p.�շ�ϸĿid = s.ҩƷid And" & vbNewLine & _
            "      s.ҩƷid = i.Id And a.���ۺ� =[1] And b.����(+) = 3 " & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ۻ��ܺ�)
    If rsTemp.RecordCount = 0 Then
        MsgBox "�õ��ۼ�¼�Ѿ���ɾ���ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If

    With vsfPrice
        .rows = 2
        rsTemp.MoveFirst
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp!ҩƷid <> lngDrugID Then
                Select Case mintUnit
                    Case 0
                        db��װϵ�� = rsTemp!ҩ���װ
                        strUnit = rsTemp!ҩ�ⵥλ
                    Case 2
                        db��װϵ�� = rsTemp!סԺ��װ
                        strUnit = rsTemp!סԺ��λ
                    Case 1
                        db��װϵ�� = rsTemp!�����װ
                        strUnit = rsTemp!���ﵥλ
                    Case 3
                        db��װϵ�� = 1
                        strUnit = rsTemp!��λ
                End Select

                lngDrugID = rsTemp!ҩƷid
                If mintMethod = 0 Or mintMethod = 2 Then
                    .TextMatrix(.rows - 1, menuPriceCol.ԭ��id) = IIf(IsNull(rsTemp!ԭ��id), "", rsTemp!ԭ��id)
                End If
                .TextMatrix(.rows - 1, menuPriceCol.ҩƷid) = rsTemp!ҩƷid

                If gintҩƷ������ʾ = 1 Then
                    .TextMatrix(.rows - 1, menuPriceCol.ҩƷ) = "[" & rsTemp!���� & "]" & IIf(IsNull(rsTemp!��Ʒ��), rsTemp!ͨ����, rsTemp!��Ʒ��)
                Else
                    .TextMatrix(.rows - 1, menuPriceCol.ҩƷ) = "[" & rsTemp!���� & "]" & rsTemp!ͨ����
                End If
                .TextMatrix(.rows - 1, menuPriceCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                .TextMatrix(.rows - 1, menuPriceCol.�Ƿ���) = rsTemp!�Ƿ���
                
'                If mintMethod = 1 Or mintMethod = 2 Then
'                    gstrSQL = "select min(����) as ���� from �ɱ��۵�����Ϣ where ���ۻ��ܺ�=[1] and ҩƷid=[2]"
'                    Set rs���� = zldatabase.OpenSQLRecord(gstrSQL, "���ز�ѯ", mstr���ۻ��ܺ�, rsTemp!ҩƷid)
'                    If rs����.RecordCount > 0 Then
'                        .TextMatrix(.rows - 1, menuPriceCol.����) = IIf(IsNull(rs����!����), "", rs����!����)
'                    End If
'                Else
                    .TextMatrix(.rows - 1, menuPriceCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
'                End If
                
                .TextMatrix(.rows - 1, menuPriceCol.��λ) = strUnit
                .TextMatrix(.rows - 1, menuPriceCol.��װϵ��) = db��װϵ��

                .TextMatrix(.rows - 1, menuPriceCol.�ӳ���) = rsTemp!�ӳ���
                .TextMatrix(.rows - 1, menuPriceCol.���������) = Nvl(rsTemp!���������, 0)
                .TextMatrix(.rows - 1, menuPriceCol.�Ƿ��п��) = rsTemp!�Ƿ��п��
                .TextMatrix(.rows - 1, menuPriceCol.������ĿID) = IIf(IsNull(rsTemp!������ĿID), "", rsTemp!������ĿID)
                .TextMatrix(.rows - 1, menuPriceCol.ԭ�ɱ���) = zlStr.FormatEx(Nvl(rsTemp!ԭ�ɱ���, 0) * db��װϵ��, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.�ֳɱ���) = zlStr.FormatEx(rsTemp!�³ɱ��� * db��װϵ��, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.ԭ���ۼ�) = zlStr.FormatEx(IIf(IsNull(rsTemp!ԭ��), rsTemp!�ּ�, rsTemp!ԭ��) * db��װϵ��, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.�����ۼ�) = zlStr.FormatEx(rsTemp!�ּ� * db��װϵ��, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.ԭ�ɹ��޼�) = zlStr.FormatEx(rsTemp!ָ������ * db��װϵ��, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.�ֲɹ��޼�) = zlStr.FormatEx(rsTemp!ָ������ * db��װϵ��, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.ԭָ���ۼ�) = zlStr.FormatEx(rsTemp!ָ���ۼ� * db��װϵ��, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.��ָ���ۼ�) = zlStr.FormatEx(rsTemp!ָ���ۼ� * db��װϵ��, mintPriceDigit, , True)

                txtValuer.Text = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                txtSummary.Text = IIf(IsNull(rsTemp!����˵��), "", rsTemp!����˵��)
                If mintModal = 1 Then
                    Me.dtpRunDate.MinDate = CDate(rsTemp!ִ������)
                End If
                If IsNull(rsTemp!ִ������) Then
                    StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
                Else
                    StrToday = Format(rsTemp!ִ������, "yyyy-MM-dd hh:mm:ss")
                End If
                Me.dtpRunDate.Value = CDate(StrToday)

                .rows = .rows + 1
                Call setColEdit
                .RowHeight(.rows - 1) = mlngRowHeight
            End If
            rsTemp.MoveNext
        Next
        Call GetDrugStore_Old(Val(.TextMatrix(1, menuPriceCol.ҩƷid)), 1)
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugStore_Old(ByVal lngDrugID As Long, ByVal intRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim dblOldCost As Double
    Dim dblOldPrice As Double
    Dim dblNewCost As Double
    Dim dblNewPrice As Double
    Dim dbl�ӳ��� As Double
    Dim lngCurRow As Long     '��ǰ��
    Dim i As Long
    Dim dbl��Ʊ��� As Double
    Dim strҩƷ���� As String
    Dim str��Ʊ As String
    Dim str��Ʊ���� As String
    Dim rsPirce As ADODB.Recordset
    Dim rsCost As ADODB.Recordset
    Dim dbl��װ���� As Double
    Dim bln��ͬҩƷ As Boolean
    Dim lngҩƷid As Long
    Dim str��λ As String

    '���ܣ�Ϊ����б��������
    '������ҩƷid

    On Error GoTo errHandle
    '�ȼ���Ƿ����ظ������ݣ�����о���������ظ�������
    With vsfStore
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    With vsfPay
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuPayCol.ҩƷid)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
        gstrSQL = "Select s.�ⷿid,s.ҩƷid, d.���� As �ⷿ, '[' || m.���� || ']' || m.���� As ҩƷ, m.���, m.����, m.���㵥λ �ۼ۵�λ, p.ҩ�ⵥλ, s.�ϴ����� As ����, nvl(s.ʵ������,0) As ����," & vbNewLine & _
            "       s.����, Nvl(m.�Ƿ���, 0) ���, m.Id, Decode(Nvl(m.�Ƿ���, 0), 0, e.�ּ�, Nvl(S.���ۼ�,0)) ʱ���ۼ�, p.�ӳ���," & vbNewLine & _
            "       nvl(s.ƽ���ɱ���,p.�ɱ���) As �ɱ���, s.�ϴι�Ӧ��id, n.���� As ��Ӧ��, s.Ч��, s.�ϴβ��� As ����" & vbNewLine & _
            " From ҩƷ��� S, ���ű� D, �շ���ĿĿ¼ M, ҩƷ��� P, ��Ӧ�� N, �շѼ�Ŀ E" & vbNewLine & _
            " Where d.Id = s.�ⷿid And s.ҩƷid = m.Id And m.Id = p.ҩƷid And Nvl(s.�ϴι�Ӧ��id, 0) = n.Id(+) And m.Id = e.�շ�ϸĿid And" & vbNewLine & _
            " s.���� = 1 And s.ҩƷid = [1] And Sysdate Between e.ִ������ And e.��ֹ����  " & vbNewLine & _
            GetPriceClassString("E") & vbNewLine & _
            " Order By  s.ҩƷid,s.�ⷿid, s.�ϴ�����,s.���� "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

        If mlng��Ӧ��ID > 0 Then
            rsTemp.Filter = "�ϴι�Ӧ��ID=" & mlng��Ӧ��ID
        End If
    Else '�޸ģ�����
        If mintModal = 2 Then   '����
            If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                gstrSQL = "select (sysdate-ִ������ ) as �Ƿ�ִ�� from ���ۻ��ܼ�¼ where ���ۺ�=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�ִ��", txtNO.Text)
                If rsTemp!�Ƿ�ִ�� > 0 Then
                    gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.ҩƷid, b.��ҩ��λid As �ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & vbNewLine & _
                            "                b.�³ɱ���, b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���, b.����, b.����, b.����, e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ," & vbNewLine & _
                            "                nvl(a.��д����,0) As ����, f.�ӳ���, b.Ч��" & vbNewLine & _
                            "From ҩƷ�շ���¼ A,�ɱ��۵�����Ϣ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F" & vbNewLine & _
                            "Where a.id=b.�շ�id And a.�ⷿid = c.Id And b.��ҩ��λid = d.Id(+) And" & vbNewLine & _
                            "      a.ҩƷid = e.Id And e.Id = f.ҩƷid And b.���ۻ��ܺ� = [1] and a.���� = 5"
                Else
                    gstrSQL = "Select Distinct a.�ⷿid,c.���� as �ⷿ, b.ҩƷid,a.�ϴι�Ӧ��id, '[' || e.���� || ']' ||e.���� as ҩƷ,e.���,d.���� as ��Ӧ��, b.�³ɱ���, b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���" & _
                            " ,a.�ϴβ��� as ����,a.����,a.�ϴ����� as ����,e.�Ƿ��� as ���,e.���㵥λ as �ۼ۵�λ,f.ҩ�ⵥλ,nvl(a.ʵ������,0) as ����,f.�ӳ���,a.Ч��" & _
                            " From ҩƷ��� A,���ű� C,��Ӧ�� D,�շ���ĿĿ¼ E,ҩƷ��� F," & _
                                 " (Select Distinct ҩƷid, �ⷿid, ����, ����, Ч��, ����, ԭ�ɱ���, �³ɱ���, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ӧ����䶯, ִ������" & _
                                   " From �ɱ��۵�����Ϣ" & _
                                   " Where ���ۻ��ܺ� = [1]) B" & _
                            " Where a.ҩƷid = b.ҩƷid And a.�ⷿid = b.�ⷿid and nvl(a.����,0)=nvl(b.����,0) and a.�ⷿid=c.id and a.�ϴι�Ӧ��id=d.id(+) and a.ҩƷid=e.id and e.id=f.ҩƷid and a.����=1 "
                End If
            ElseIf cboPriceMethod.Text = "�����ۼ�" Then
                gstrSQL = "select (sysdate-ִ������ ) as �Ƿ�ִ�� from ���ۻ��ܼ�¼ where ���ۺ�=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�ִ��", txtNO.Text)
                If rsTemp!�Ƿ�ִ�� > 0 Then
                    gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ҩƷid, a.��ҩ��λid As �ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���," & vbNewLine & _
                            "                d.���� As ��Ӧ��, f.�ɱ��� As �³ɱ���, f.�ɱ��� As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.����, a.����, a.����, e.�Ƿ��� As ���," & vbNewLine & _
                            "                e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ, nvl(a.��д����,0) As ����, f.�ӳ���, a.Ч��" & vbNewLine & _
                            "From ҩƷ�շ���¼ A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F" & vbNewLine & _
                            "Where a.�۸�id = b.Id And a.�ⷿid = c.Id And a.��ҩ��λid = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.ҩƷid And" & vbNewLine & _
                            "      b.���ۻ��ܺ� = [1] and a.����=13 And a.����id Is Null " & GetPriceClassString("B")
                Else
                    gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ҩƷid, a.�ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & _
                                            " nvl(a.ƽ���ɱ���,f.�ɱ���) As �³ɱ���, nvl(a.ƽ���ɱ���,f.�ɱ���) As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.�ϴβ��� As ����, a.����, a.�ϴ����� As ����," & _
                                            " e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ, nvl(a.ʵ������,0) As ����, f.�ӳ���, a.Ч��" & _
                            " From ҩƷ��� A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F" & _
                            " Where a.ҩƷid = b.�շ�ϸĿid And a.�ⷿid = c.Id And a.�ϴι�Ӧ��id = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.ҩƷid And a.���� = 1  And" & _
                                  " b.���ۻ��ܺ� = [1]" & GetPriceClassString("B")
                End If
            End If
        Else '�޸�
            If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                gstrSQL = "Select Distinct a.�ⷿid,c.���� as �ⷿ, b.ҩƷid,a.�ϴι�Ӧ��id, '[' || e.���� || ']' ||e.���� as ҩƷ,e.���,d.���� as ��Ӧ��, b.�³ɱ���, b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���" & _
                            " ,a.�ϴβ��� as ����,a.����,a.�ϴ����� as ����,e.�Ƿ��� as ���,e.���㵥λ as �ۼ۵�λ,f.ҩ�ⵥλ,nvl(a.ʵ������,0) as ����,f.�ӳ���,a.Ч��" & _
                            " From ҩƷ��� A,���ű� C,��Ӧ�� D,�շ���ĿĿ¼ E,ҩƷ��� F," & _
                                 " (Select Distinct ҩƷid, �ⷿid, ����, ����, Ч��, ����, ԭ�ɱ���, �³ɱ���, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ӧ����䶯, ִ������" & _
                                   " From �ɱ��۵�����Ϣ" & _
                                   " Where ���ۻ��ܺ� = [1]) B" & _
                            " Where a.ҩƷid = b.ҩƷid And a.�ⷿid = b.�ⷿid and nvl(a.����,0)=nvl(b.����,0) and a.�ⷿid=c.id and a.�ϴι�Ӧ��id=d.id(+) and a.ҩƷid=e.id and e.id=f.ҩƷid and a.����=1 "
            ElseIf cboPriceMethod.Text = "�����ۼ�" Then
                gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ҩƷid, a.�ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & _
                                            " nvl(a.ƽ���ɱ���,f.�ɱ���) As �³ɱ���, nvl(a.ƽ���ɱ���,f.�ɱ���) As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.�ϴβ��� As ����, a.����, a.�ϴ����� As ����," & _
                                            " e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ, nvl(a.ʵ������,0) As ����, f.�ӳ���, a.Ч��" & _
                            " From ҩƷ��� A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F" & _
                            " Where a.ҩƷid = b.�շ�ϸĿid And a.�ⷿid = c.Id And a.�ϴι�Ӧ��id = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.ҩƷid And a.���� = 1  And" & _
                                  " b.���ۻ��ܺ� = [1]" & GetPriceClassString("B")
            End If
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, txtNO.Text)
    End If
    
    With vsfStore
        Do While Not rsTemp.EOF
            dbl��װ���� = 0
            dbl��Ʊ��� = 0
            dblOldPrice = 0
            dblNewPrice = 0
            For i = 0 To vsfPrice.rows - 1
                If rsTemp!ҩƷid = vsfPrice.TextMatrix(i, menuPriceCol.ҩƷid) Then
                    dbl��װ���� = vsfPrice.TextMatrix(i, menuPriceCol.��װϵ��)
                    dblOldPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.ԭ���ۼ�))
                    dblNewPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.�����ۼ�))
                    str��λ = vsfPrice.TextMatrix(i, menuPriceCol.��λ)
                    Exit For
                End If
            Next
            .rows = .rows + 1
            Call setColEdit
            .RowHeight(.rows - 1) = mlngRowHeight

            '�ӿհ��п�ʼ��������
            .TextMatrix(.rows - 1, menuStoreCol.ҩƷid) = rsTemp!ҩƷid
            .TextMatrix(.rows - 1, menuStoreCol.�ⷿ) = rsTemp!�ⷿ
            .TextMatrix(.rows - 1, menuStoreCol.�ⷿid) = rsTemp!�ⷿid
            .TextMatrix(.rows - 1, menuStoreCol.��Ӧ��) = Nvl(rsTemp!��Ӧ��, "")
            .TextMatrix(.rows - 1, menuStoreCol.��Ӧ��id) = IIf(mlng��Ӧ��ID > 0, mlng��Ӧ��ID, Nvl(rsTemp!�ϴι�Ӧ��ID))
            .TextMatrix(.rows - 1, menuStoreCol.ҩƷ) = rsTemp!ҩƷ
            strҩƷ���� = rsTemp!ҩƷ

            .TextMatrix(.rows - 1, menuStoreCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            .TextMatrix(.rows - 1, menuStoreCol.��λ) = str��λ
            .TextMatrix(.rows - 1, menuStoreCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.rows - 1, menuStoreCol.Ч��) = Format(IIf(IsNull(rsTemp!Ч��), "", rsTemp!Ч��), "YYYY-MM-DD")
            .TextMatrix(.rows - 1, menuStoreCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.rows - 1, menuStoreCol.����) = zlStr.FormatEx(rsTemp!���� / dbl��װ����, mintNumberDigit, , True)
            .TextMatrix(.rows - 1, menuStoreCol.��װϵ��) = dbl��װ����
            .TextMatrix(.rows - 1, menuStoreCol.����) = Nvl(rsTemp!����, 0)
            .TextMatrix(.rows - 1, menuStoreCol.���) = rsTemp!���


            If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
                dblOldCost = IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���) * dbl��װ����

                If mdbl�ӳ��� > 0 Then
                    dbl�ӳ��� = Round(mdbl�ӳ��� / 100, 7)
                ElseIf dblOldCost > 0 Then
                    dbl�ӳ��� = Round(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice) / dblOldCost - 1, 7)
                Else
                    dbl�ӳ��� = Round(rsTemp!�ӳ��� / 100, 2)
                End If
                If 1 + dbl�ӳ��� = 0 Then
                    dblNewCost = 0
                Else
                    dblNewCost = rsTemp!ʱ���ۼ� * dbl��װ���� / (1 + dbl�ӳ���)
                End If
                If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0

                .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = zlStr.FormatEx(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice), mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = zlStr.FormatEx(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice), mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�ۼ�ӯ��) = Format(rsTemp!���� / dbl��װ���� * (Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                .TextMatrix(.rows - 1, menuStoreCol.�ӳ���) = dbl�ӳ��� * 100
                .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɱ���) = zlStr.FormatEx(dblOldCost, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(dblNewCost, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�ɱ�ӯ��) = Format((Val(.TextMatrix(.rows - 1, menuStoreCol.�ֳɱ���)) - Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ�ɱ���))) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat)
                dbl��Ʊ��� = dbl��Ʊ��� + (dblNewCost - dblOldCost) * Val(.TextMatrix(.rows - 1, menuStoreCol.����))
                
                'ΪӦ����¼��ֵ
                If mint���� = 1 Or mint���� = 2 Then
                    If vsfPay.rows > 1 Then
                        bln��ͬҩƷ = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.ҩƷid) = rsTemp!ҩƷid Then
                                bln��ͬҩƷ = True
                                Exit For
                            End If
                        Next
                        If bln��ͬҩƷ = True Then
                            vsfPay.TextMatrix(i, menuPayCol.��Ʊ���) = zlStr.FormatEx(Val(vsfPay.TextMatrix(i, menuPayCol.��Ʊ���)) + dbl��Ʊ���, mintMoneyDigit, , True)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷ) = strҩƷ����
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = zlStr.FormatEx(dbl��Ʊ���, mintMoneyDigit, , True)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷ) = strҩƷ����
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = zlStr.FormatEx(dbl��Ʊ���, mintMoneyDigit, , True)
                    End If
                End If
            Else
                If mintModal = 2 And (cboPriceMethod.Text = "�����ۼ�" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����") Then   '����
                    gstrSQL = "Select a.�ɱ��� As ԭ��, a.���ۼ� As �ּ�" & vbNewLine & _
                        "From ҩƷ�շ���¼ A, �շѼ�Ŀ B" & vbNewLine & _
                        "Where a.�۸�id = b.Id And b.���ۻ��ܺ� = [1] And a.�ⷿid = [2] And a.ҩƷid = [3] And Nvl(a.����, 0) = [4]" & _
                        GetPriceClassString("B")
                        
                    Set rsPirce = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ۼ�", txtNO.Text, rsTemp!�ⷿid, rsTemp!ҩƷid, Nvl(rsTemp!����, 0))
                    
                    If Not rsPirce.EOF Then
                        .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = zlStr.FormatEx(Val(rsPirce!ԭ��) * dbl��װ����, mintPriceDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = zlStr.FormatEx(Val(rsPirce!�ּ�) * dbl��װ����, mintPriceDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.�ۼ�ӯ��) = Format(rsTemp!���� / dbl��װ���� * (Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = zlStr.FormatEx(dblOldPrice, mintPriceDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = zlStr.FormatEx(dblNewPrice, mintPriceDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.�ۼ�ӯ��) = Format(rsTemp!���� / dbl��װ���� * (Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                    End If
                    If cboPriceMethod.Text = "�����ۼ�" Then
                        gstrSQL = "Select �ɱ���" & vbNewLine & _
                                    "      From (Select ƽ���ɱ��� As �ɱ���" & vbNewLine & _
                                    "             From ҩƷ���" & vbNewLine & _
                                    "             Where ����=1 And �ⷿid = [1] And ҩƷid = [2] And nvl(����,0) = [3]" & vbNewLine & _
                                    "             Union All" & vbNewLine & _
                                    "             Select �ɱ��� From ҩƷ��� Where ҩƷid = [2])" & vbNewLine & _
                                    "      Where Rownum <= 1"

                        Set rsCost = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ɱ���", rsTemp!�ⷿid, rsTemp!ҩƷid, Nvl(rsTemp!����, 0))
                        .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɱ���) = zlStr.FormatEx(rsCost!�ɱ��� * dbl��װ����, mintCostDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(rsCost!�ɱ��� * dbl��װ����, mintCostDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.�ɱ�ӯ��) = Format(0, mstrMoneyFormat)
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɱ���) = zlStr.FormatEx(Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����, mintCostDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(rsTemp!�³ɱ��� * dbl��װ����, mintCostDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.�ɱ�ӯ��) = Format((rsTemp!�³ɱ��� * dbl��װ���� - Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat)
                    End If
                Else '�޸Ļ��߳ɱ��۵���
                    '����ֱ�Ӵ��շѼ�Ŀȡ�ּۣ�ʱ�����ȴӿ��ȡ�����û������շѼ�Ŀȡ
                    If Nvl(rsTemp!���, 0) = 1 Then
                        gstrSQL = "Select Nvl(s.���ۼ�, Decode(Nvl(s.ʵ������, 0), 0, 0, Nvl(s.ʵ�ʽ��, 0) / s.ʵ������)) ʱ���ۼ�" & vbNewLine & _
                        "From ҩƷ��� S" & vbNewLine & _
                        "Where s.����=1 And s.�ⷿid = [1] And s.ҩƷid = [2] And nvl(s.����,0) = [3]"
                        
                        Set rsPirce = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ۼ�", rsTemp!�ⷿid, rsTemp!ҩƷid, Nvl(rsTemp!����, 0))
                        If rsPirce.RecordCount > 0 Then
                            If rsPirce!ʱ���ۼ� > 0 Then
                                .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = zlStr.FormatEx(rsPirce!ʱ���ۼ� * dbl��װ����, mintPriceDigit, , True)
                                .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = zlStr.FormatEx(rsPirce!ʱ���ۼ� * dbl��װ����, mintPriceDigit, , True)
                            Else
                                .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = zlStr.FormatEx(dblOldPrice, mintPriceDigit, , True)
                                .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = zlStr.FormatEx(dblNewPrice, mintPriceDigit, , True)
                            End If
                        Else
                            .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = zlStr.FormatEx(dblOldPrice, mintPriceDigit, , True)
                            .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = zlStr.FormatEx(dblNewPrice, mintPriceDigit, , True)
                        End If
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = zlStr.FormatEx(dblOldPrice, mintPriceDigit, , True)
                        .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = zlStr.FormatEx(dblNewPrice, mintPriceDigit, , True)
                    End If
                    .TextMatrix(.rows - 1, menuStoreCol.�ۼ�ӯ��) = Format(rsTemp!���� / dbl��װ���� * (Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                    .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɱ���) = zlStr.FormatEx(Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(rsTemp!�³ɱ��� * dbl��װ����, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.�ɱ�ӯ��) = Format((rsTemp!�³ɱ��� * dbl��װ���� - Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat)
                End If
                 
                If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    If rsTemp!�³ɱ��� = 0 Then
                        dbl�ӳ��� = 0
                    Else
                        dbl�ӳ��� = Round(Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) / (rsTemp!�³ɱ��� * dbl��װ����) - 1, 7)
                    End If
                    .TextMatrix(.rows - 1, menuStoreCol.�ӳ���) = dbl�ӳ��� * 100
                    .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɱ���) = zlStr.FormatEx(Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(rsTemp!�³ɱ��� * dbl��װ����, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.�ɱ�ӯ��) = Format((rsTemp!�³ɱ��� * dbl��װ���� - Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat)
                    dbl��Ʊ��� = dbl��Ʊ��� + (rsTemp!�³ɱ��� * dbl��װ���� - Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����))
                    str��Ʊ = IIf(IsNull(rsTemp!��Ʊ��), "", rsTemp!��Ʊ��)
                    str��Ʊ���� = IIf(IsNull(rsTemp!��Ʊ����), "", rsTemp!��Ʊ����)
                    
                    'Ϊ�����¼�б�ֵ
                    If vsfPay.rows > 1 Then
                        bln��ͬҩƷ = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.ҩƷid) = rsTemp!ҩƷid Then
                                bln��ͬҩƷ = True
                                Exit For
                            End If
                        Next
                        If bln��ͬҩƷ = True Then
                            vsfPay.TextMatrix(i, menuPayCol.��Ʊ���) = zlStr.FormatEx(Val(vsfPay.TextMatrix(i, menuPayCol.��Ʊ���)) + dbl��Ʊ���, mintMoneyDigit, , True)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷ) = strҩƷ����
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = zlStr.FormatEx(dbl��Ʊ���, mintMoneyDigit, , True)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷ) = strҩƷ����
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = zlStr.FormatEx(dbl��Ʊ���, mintMoneyDigit, , True)
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    '�޸ĺͲ���ʱ�������б�ƽ���ɱ��ۣ��ۼ�
    'mintModal 0-���� 1-�޸� 2-����
    If mintModal = 1 Or mintModal = 2 Then
        With vsfStore
            For i = 1 To .rows - 1
                If lngҩƷid <> .TextMatrix(i, menuStoreCol.ҩƷid) Then
                    Call CaluateAverCost(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    Call CaluateAverOldCost(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    Call CaculateAverPirce(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    Call CaculateAverOldPirce(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    lngҩƷid = Val(.TextMatrix(i, menuStoreCol.ҩƷid))
                End If
            Next
        End With
    End If

    If mint���� = 1 Or mint���� = 2 Then
        If rsTemp.RecordCount = 0 Then Exit Sub
        TabCtlDetails.Item(1).Visible = True
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initComboBox()
    With cbo�ۼۼ��㷽ʽ
        .AddItem "�ۼ���ɱ��۲���������"
        .AddItem "�ۼ۰��̶���������"
        .AddItem "�ۼ۰��ֶμӳɼ���"
        .ListIndex = 0
    End With

    With cboPriceMethod
        If mintModal <> 2 Then  '�ǲ���
            If InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") = 0 Then
                .AddItem "�����ɱ���"
                .ListIndex = 0
                lblMethod.Tag = 0
            ElseIf InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") > 0 Then
                .AddItem "�����ۼ�"
                .ListIndex = 0
                lblMethod.Tag = 0
            ElseIf InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") > 0 Then
                .AddItem "�����ۼ�"
                .AddItem "�����ɱ���"
                .AddItem "�ۼ۳ɱ���һ�����"
                .ListIndex = 0
                lblMethod.Tag = 0
            End If
        Else
            .AddItem "�����ۼ�"
            .AddItem "�����ɱ���"
            .AddItem "�ۼ۳ɱ���һ�����"
            .ListIndex = 0
            lblMethod.Tag = 0
        End If
    End With
End Sub

Private Sub InitTabControl()
    '��ʼ��TabControl�ؼ�
    Dim objtabctl As TabControlItem

    picSplit.Left = 0
    picSplit.Top = vsfPrice.Top + vsfPrice.Height + 5
    With TabCtlDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem 0, "���䶯��", vsfStore.hWnd, 0
        .InsertItem 1, "Ӧ����䶯��", vsfPay.hWnd, 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - vsfPrice.Height - vsfPrice.Top - 20
        .Top = picSplit.Height + picSplit.Top + 20
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If

    With fraCondition
        .Width = Me.ScaleWidth
    End With
    txtNO.Move fraCondition.Width - 2000
    lblNO.Move fraCondition.Width - lblNO.Width - 2100
    
    vsfPrice.Move 20, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, 3000
    picSplit.Left = 50
    picSplit.Top = vsfPrice.Top + vsfPrice.Height + 5
    picSplit.Width = Me.ScaleWidth
'    txtSummary.Width = Me.ScaleWidth - lblSummary.Left - lblSummary.Width - 300
    TabCtlDetails.Move 20, picSplit.Height + picSplit.Top, Me.ScaleWidth, Me.ScaleHeight - picSplit.Top - picSplit.Height - picInfo.Height - stbThis.Height
    picInfo.Move 0, TabCtlDetails.Top + TabCtlDetails.Height, Me.ScaleWidth
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    
    With txtSummary
        .Width = picInfo.Width - .Left - 300
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ReleaseSelectorRS
    Call SaveWinState(Me, App.ProductName, MStrCaption)
    mblnLoad = False
    mblnӦ����¼ = False
    mlng��Ӧ��ID = 0
    mblnUpdateAdd = False
End Sub

Private Sub mshProvider_DblClick()
    With Me.mshProvider
        Me.txt��Ӧ��.Text = .TextMatrix(.Row, 1)
        Me.txt��Ӧ��.Tag = .TextMatrix(.Row, 3) & "|" & .TextMatrix(.Row, 1)
        .Visible = False
    End With

    Me.txt��Ӧ��.SetFocus
End Sub

Private Sub optʱ��_Click(Index As Integer)
    If Index = 0 Then
        dtpRunDate.Enabled = False
    Else
        dtpRunDate.Enabled = True
    End If
End Sub

Private Sub InitVsfGridFlex()
    With vsfPrice

        .Cols = menuPriceCol.������
        .rows = 2
        .RowHeight(1) = mlngRowHeight
        .ColWidth(0) = 200
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
'        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMoveRows '�϶�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .Editable = flexEDNone
'        .GridLineWidth = 2
'        .GridLines = flexGridInset
'        .GridColor = &H80000011
'        .GridColorFixed = &H80000011
'        .ForeColorFixed = &H80000012
'        .BackColorSel = &HF4F4EA

        .TextMatrix(0, menuPriceCol.ҩƷid) = "ҩƷID"
        .TextMatrix(0, menuPriceCol.ԭ��id) = "ԭ��id"
        .TextMatrix(0, menuPriceCol.ҩ������) = "ҩ������"
        .TextMatrix(0, menuPriceCol.ҩƷ) = "ҩƷ"
        .TextMatrix(0, menuPriceCol.���) = "���"
        .TextMatrix(0, menuPriceCol.�Ƿ���) = "�Ƿ���"
        .TextMatrix(0, menuPriceCol.����) = "������"
        .TextMatrix(0, menuPriceCol.��λ) = "��λ"
        .TextMatrix(0, menuPriceCol.��װϵ��) = "��װϵ��"
        .TextMatrix(0, menuPriceCol.�ӳ���) = "�ӳ���"
        .TextMatrix(0, menuPriceCol.���������) = "���������"
        .TextMatrix(0, menuPriceCol.�Ƿ��п��) = "�Ƿ��п��"
        .TextMatrix(0, menuPriceCol.������ĿID) = "������Ŀid"
        .TextMatrix(0, menuPriceCol.ԭ�ɱ���) = "ԭ�ɱ���"
        .TextMatrix(0, menuPriceCol.�ֳɱ���) = "�ֳɱ���"
        .TextMatrix(0, menuPriceCol.ԭ���ۼ�) = "ԭ���ۼ�"
        .TextMatrix(0, menuPriceCol.�����ۼ�) = "�����ۼ�"
        .TextMatrix(0, menuPriceCol.ԭ�ɹ��޼�) = "ԭ�ɹ��޼�"
        .TextMatrix(0, menuPriceCol.�ֲɹ��޼�) = "�ֲɹ��޼�"
        .TextMatrix(0, menuPriceCol.ԭָ���ۼ�) = "ԭָ���ۼ�"
        .TextMatrix(0, menuPriceCol.��ָ���ۼ�) = "��ָ���ۼ�"

        '�����п�
        .ColWidth(menuPriceCol.ҩƷid) = 0
        .ColWidth(menuPriceCol.ԭ��id) = 0
        .ColWidth(menuPriceCol.ҩ������) = 1000
        .ColWidth(menuPriceCol.ҩƷ) = 3000
        .ColWidth(menuPriceCol.���) = 1500
        .ColWidth(menuPriceCol.�Ƿ���) = 0
        .ColWidth(menuPriceCol.����) = 2000
        .ColWidth(menuPriceCol.��λ) = 800
        .ColWidth(menuPriceCol.��װϵ��) = 0
        .ColWidth(menuPriceCol.�ӳ���) = 0
        .ColWidth(menuPriceCol.���������) = 0
        .ColWidth(menuPriceCol.�Ƿ��п��) = 0
        .ColWidth(menuPriceCol.������ĿID) = 0
        .ColWidth(menuPriceCol.ԭ�ɱ���) = 1000
        .ColWidth(menuPriceCol.�ֳɱ���) = 1000
        .ColWidth(menuPriceCol.ԭ���ۼ�) = 1000
        .ColWidth(menuPriceCol.�����ۼ�) = 1000
        .ColWidth(menuPriceCol.ԭ�ɹ��޼�) = 0
        .ColWidth(menuPriceCol.�ֲɹ��޼�) = 0
        .ColWidth(menuPriceCol.ԭָ���ۼ�) = 0
        .ColWidth(menuPriceCol.��ָ���ۼ�) = 0
        '���ö��뷽ʽ
        .ColAlignment(menuPriceCol.ҩ������) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.ҩƷ) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.���) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.����) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.��λ) = flexAlignCenterCenter
        .ColAlignment(menuPriceCol.ԭ�ɱ���) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.�ֳɱ���) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.ԭ���ۼ�) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.�����ۼ�) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.ԭ�ɹ��޼�) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.ԭָ���ۼ�) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '��ͷ���ж���
        .ColComboList(menuPriceCol.ҩƷ) = "|..."
    End With

    With vsfStore
        .Editable = flexEDNone
        .Cols = menuStoreCol.������
        .rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
'        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMoveRows '�϶�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        '��������
        .TextMatrix(0, menuStoreCol.ҩƷid) = "ҩƷid"
        .TextMatrix(0, menuStoreCol.�ⷿ) = "�ⷿ"
        .TextMatrix(0, menuStoreCol.�ⷿid) = "�ⷿid"
        .TextMatrix(0, menuStoreCol.��Ӧ��) = "��Ӧ��"
        .TextMatrix(0, menuStoreCol.��Ӧ��id) = "��Ӧ��id"
        .TextMatrix(0, menuStoreCol.ҩƷ) = "ҩƷ"
        .TextMatrix(0, menuStoreCol.���) = "���"
        .TextMatrix(0, menuStoreCol.��λ) = "��λ"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.Ч��) = "Ч��"
        .TextMatrix(0, menuStoreCol.����) = "������"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.��װϵ��) = "��װϵ��"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.���) = "���"
        .TextMatrix(0, menuStoreCol.ԭ���ۼ�) = "ԭ���ۼ�"
        .TextMatrix(0, menuStoreCol.�����ۼ�) = "�����ۼ�"
        .TextMatrix(0, menuStoreCol.�ۼ�ӯ��) = "�ۼ�ӯ��"
        .TextMatrix(0, menuStoreCol.�ӳ���) = "�ӳ���"
        .TextMatrix(0, menuStoreCol.ԭ�ɱ���) = "ԭ�ɱ���"
        .TextMatrix(0, menuStoreCol.�ֳɱ���) = "�ֳɱ���"
        .TextMatrix(0, menuStoreCol.�ɱ�ӯ��) = "�ɱ�ӯ��"
        '�����п�
        .ColWidth(0) = 0
        .ColWidth(menuStoreCol.�ⷿ) = 1500
        .ColWidth(menuStoreCol.�ⷿid) = 0
        .ColWidth(menuStoreCol.��Ӧ��) = 2000
        .ColWidth(menuStoreCol.��Ӧ��id) = 0
        .ColWidth(menuStoreCol.ҩƷ) = 3000
        .ColWidth(menuStoreCol.���) = 1500
        .ColWidth(menuStoreCol.��λ) = 800
        .ColWidth(menuStoreCol.����) = 1500
        .ColWidth(menuStoreCol.Ч��) = 2000
        .ColWidth(menuStoreCol.����) = 1500
        .ColWidth(menuStoreCol.����) = 1500
        .ColWidth(menuStoreCol.��װϵ��) = 0
        .ColWidth(menuStoreCol.����) = 0
        .ColWidth(menuStoreCol.���) = 0
        .ColWidth(menuStoreCol.ԭ���ۼ�) = 1000
        .ColWidth(menuStoreCol.�����ۼ�) = 1000
        .ColWidth(menuStoreCol.�ۼ�ӯ��) = 1000
        .ColWidth(menuStoreCol.�ӳ���) = 1000
        .ColWidth(menuStoreCol.ԭ�ɱ���) = 1000
        .ColWidth(menuStoreCol.�ֳɱ���) = 1000
        .ColWidth(menuStoreCol.�ɱ�ӯ��) = 1000
        '���뷽ʽ
        .ColAlignment(menuStoreCol.�ⷿ) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.��Ӧ��) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.ҩƷ) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.���) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.��λ) = flexAlignCenterCenter
        .ColAlignment(menuStoreCol.����) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.Ч��) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.����) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.����) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.ԭ���ۼ�) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�����ۼ�) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�ۼ�ӯ��) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�ӳ���) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.ԭ�ɱ���) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�ֳɱ���) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�ɱ�ӯ��) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '��ͷ���ж���
    End With

    With vsfPay
        .Editable = flexEDNone
        .Cols = menuPayCol.������
        .rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
'        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMoveRows '�϶�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        .TextMatrix(0, menuPayCol.ҩƷid) = "ҩƷid"
        .TextMatrix(0, menuPayCol.ҩƷ) = "ҩƷ"
        .TextMatrix(0, menuPayCol.��Ʊ��) = "��Ʊ��"
        .TextMatrix(0, menuPayCol.��Ʊ����) = "��Ʊ����"
        .TextMatrix(0, menuPayCol.��Ʊ���) = "��Ʊ���"
        '�����п�
        .ColWidth(menuPayCol.ҩƷid) = 0
        .ColWidth(menuPayCol.ҩƷ) = 2000
        .ColWidth(menuPayCol.��Ʊ��) = 1500
        .ColWidth(menuPayCol.��Ʊ����) = 2000
        .ColWidth(menuPayCol.��Ʊ���) = 1500
        '���뷽ʽ
        .ColAlignment(menuPayCol.ҩƷ) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.��Ʊ��) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.��Ʊ����) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.��Ʊ���) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '��ͷ���ж���
    End With
End Sub

Private Sub initGrid()
    '������޸Ļ��߲�������ȡ��Ӧ�ļ�¼����䵽�����
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim i As Long
    Dim lngDrugID As Long
    Dim db��װϵ�� As Double
    Dim strUnit As String
    Dim StrToday As String
    Dim rs���� As ADODB.Recordset

    On Error GoTo errHandle
    '���۷�ʽ 0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
    If mintMethod = 0 Then
        gstrSQL = "Select Distinct p.ԭ��id, i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "                s.�ӳ���/100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, i.���� As ����, i.���㵥λ As ��λ," & vbNewLine & _
            "                s.���ﵥλ, s.�����װ, s.סԺ��λ, s.סԺ��װ, s.ҩ�ⵥλ, Nvl(s.ҩ���װ, 1) ҩ���װ, s.�ɱ��� As ԭ�ɱ���, s.�ɱ��� As �³ɱ���, p.ԭ��, p.�ּ�," & vbNewLine & _
            "                p.������Ŀid, p.������, p.����˵��, s.���������, To_Char(a.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ҩƷid," & vbNewLine & _
            "                Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select s.ҩƷid From ҩƷ��� s where s.����=1 And Not (zl_fun_getbatchpro(s.�ⷿid,s.ҩƷid)=1 And Nvl(S.����,0) = 0 And S.�������� < 0 And S.ʵ������ = 0 And S.ʵ�ʽ�� = 0 And S.ʵ�ʲ�� = 0)) K, ���ۻ��ܼ�¼ A, �շ���Ŀ���� B, ҩƷ��� S, �շ���ĿĿ¼ I, �շѼ�Ŀ P" & vbNewLine & _
            "Where a.���ۺ� = p.���ۻ��ܺ� And b.�շ�ϸĿid(+) = s.ҩƷid And s.ҩƷid = i.Id And i.Id = k.ҩƷid(+) And i.Id = p.�շ�ϸĿid And" & vbNewLine & _
            "      p.���ۻ��ܺ� = [1] And a.���� = 0 And b.����(+) = 3 And a.���ۺ� = [1] " & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            "Order By ҩƷid"
    ElseIf mintMethod = 1 Then
        gstrSQL = "Select Distinct i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "                s.�ӳ���/100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, i.���� As ����, i.���㵥λ As ��λ," & vbNewLine & _
            "                s.���ﵥλ, s.�����װ, s.סԺ��λ, s.סԺ��װ, s.ҩ�ⵥλ, Nvl(s.ҩ���װ, 1) ҩ���װ, m.ԭ�ɱ���, m.�³ɱ���, p.ԭ��, p.�ּ�, p.������Ŀid," & vbNewLine & _
            "                a.������ As ������, a.˵�� As ����˵��, s.���������, To_Char(m.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ҩƷid," & vbNewLine & _
            "                Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select Min(ԭ��) As ԭ�ɱ���, Min(�ּ�) As �³ɱ���, min(����) as ����,���ۻ��ܺ�,ҩƷid,min(ִ������) as ִ������ From ҩƷ�۸��¼ Where �۸�����=2 and ���ۻ��ܺ� = [1] Group By ���ۻ��ܺ�,ҩƷid) M, (Select s.ҩƷid From ҩƷ��� s where s.����=1 And Not (zl_fun_getbatchpro(s.�ⷿid,s.ҩƷid)=1 And Nvl(S.����,0) = 0 And S.�������� < 0 And S.ʵ������ = 0 And S.ʵ�ʽ�� = 0 And S.ʵ�ʲ�� = 0)) K, ���ۻ��ܼ�¼ A, �շ���Ŀ���� B, ҩƷ��� S, �շ���ĿĿ¼ I, �շѼ�Ŀ P" & vbNewLine & _
            "Where m.���ۻ��ܺ�(+) = a.���ۺ� And b.�շ�ϸĿid(+) = s.ҩƷid And s.ҩƷid = i.Id And i.Id = k.ҩƷid(+) And m.ҩƷid = i.Id And" & vbNewLine & _
            "      i.Id = p.�շ�ϸĿid And Sysdate Between p.ִ������ And p.��ֹ���� And m.���ۻ��ܺ� = [1] And a.���� = 0 And b.����(+) = 3 And" & vbNewLine & _
            "      a.���ۺ� = [1] " & IIf(mintModal = 2, "", " And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            "Order By ҩƷid"
    ElseIf mintMethod = 2 Then
        gstrSQL = "Select distinct p.ԭ��id, i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "       s.�ӳ���/100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, i.���� As ����, i.���㵥λ As ��λ, s.���ﵥλ," & vbNewLine & _
            "       s.�����װ, s.סԺ��λ, s.סԺ��װ, s.ҩ�ⵥλ, Nvl(s.ҩ���װ, 1) ҩ���װ, m.ԭ�ɱ���, m.�³ɱ���, p.ԭ��, p.�ּ�, p.������Ŀid, p.������, p.����˵��, s.���������," & vbNewLine & _
            "       To_Char(p.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ҩƷid, Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select ҩƷid,Min(ԭ��) As ԭ�ɱ���, Min(�ּ�) As �³ɱ���, min(����) as ����,���ۻ��ܺ� From ҩƷ�۸��¼ Where �۸�����=2 and ���ۻ��ܺ� = [1] Group By ҩƷid,���ۻ��ܺ�) M, �շѼ�Ŀ P, ���ۻ��ܼ�¼ A, (Select s.ҩƷid From ҩƷ��� s where s.����=1 And Not (zl_fun_getbatchpro(s.�ⷿid,s.ҩƷid)=1 And Nvl(S.����,0) = 0 And S.�������� < 0 And S.ʵ������ = 0 And S.ʵ�ʽ�� = 0 And S.ʵ�ʲ�� = 0)) K, �շ���Ŀ���� B, ҩƷ��� S, �շ���ĿĿ¼ I" & vbNewLine & _
            "Where m.���ۻ��ܺ� = a.���ۺ� and m.ҩƷid=i.id And p.���ۻ��ܺ� = a.���ۺ� And p.�շ�ϸĿid = k.ҩƷid(+) And p.�շ�ϸĿid = b.�շ�ϸĿid(+) And p.�շ�ϸĿid = s.ҩƷid And" & vbNewLine & _
            "      s.ҩƷid = i.Id And a.���ۺ� =[1] And b.����(+) = 3 " & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & "Order By ҩƷid "
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ۻ��ܺ�)
    If rsTemp.RecordCount = 0 Then
        MsgBox "�õ��ۼ�¼�Ѿ���ɾ���ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If

    With vsfPrice
        .rows = 2
        rsTemp.MoveFirst
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp!ҩƷid <> lngDrugID Then
                Select Case mintUnit
                    Case 0
                        db��װϵ�� = rsTemp!ҩ���װ
                        strUnit = rsTemp!ҩ�ⵥλ
                    Case 2
                        db��װϵ�� = rsTemp!סԺ��װ
                        strUnit = rsTemp!סԺ��λ
                    Case 1
                        db��װϵ�� = rsTemp!�����װ
                        strUnit = rsTemp!���ﵥλ
                    Case 3
                        db��װϵ�� = 1
                        strUnit = rsTemp!��λ
                End Select

                lngDrugID = rsTemp!ҩƷid
                If mintMethod = 0 Or mintMethod = 2 Then
                    .TextMatrix(.rows - 1, menuPriceCol.ԭ��id) = IIf(IsNull(rsTemp!ԭ��id), "", rsTemp!ԭ��id)
                End If
                .TextMatrix(.rows - 1, menuPriceCol.ҩƷid) = rsTemp!ҩƷid

                If gintҩƷ������ʾ = 1 Then
                    .TextMatrix(.rows - 1, menuPriceCol.ҩƷ) = "[" & rsTemp!���� & "]" & IIf(IsNull(rsTemp!��Ʒ��), rsTemp!ͨ����, rsTemp!��Ʒ��)
                Else
                    .TextMatrix(.rows - 1, menuPriceCol.ҩƷ) = "[" & rsTemp!���� & "]" & rsTemp!ͨ����
                End If
                .TextMatrix(.rows - 1, menuPriceCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                .TextMatrix(.rows - 1, menuPriceCol.�Ƿ���) = rsTemp!�Ƿ���
                .TextMatrix(.rows - 1, menuPriceCol.ҩ������) = IIf(rsTemp!�Ƿ��� = 0, "����", "ʱ��")
                
'                If mintMethod = 1 Or mintMethod = 2 Then
'                    gstrSQL = "select min(����) as ���� from �ɱ��۵�����Ϣ where ���ۻ��ܺ�=[1] and ҩƷid=[2]"
'                    Set rs���� = zldatabase.OpenSQLRecord(gstrSQL, "���ز�ѯ", mstr���ۻ��ܺ�, rsTemp!ҩƷid)
'                    If rs����.RecordCount > 0 Then
'                        .TextMatrix(.rows - 1, menuPriceCol.����) = IIf(IsNull(rs����!����), "", rs����!����)
'                    End If
'                Else
                    .TextMatrix(.rows - 1, menuPriceCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
'                End If
                
                .TextMatrix(.rows - 1, menuPriceCol.��λ) = strUnit
                .TextMatrix(.rows - 1, menuPriceCol.��װϵ��) = db��װϵ��

                .TextMatrix(.rows - 1, menuPriceCol.�ӳ���) = rsTemp!�ӳ���
                .TextMatrix(.rows - 1, menuPriceCol.���������) = Nvl(rsTemp!���������, 100)
                .TextMatrix(.rows - 1, menuPriceCol.�Ƿ��п��) = rsTemp!�Ƿ��п��
                .TextMatrix(.rows - 1, menuPriceCol.������ĿID) = IIf(IsNull(rsTemp!������ĿID), "", rsTemp!������ĿID)
                .TextMatrix(.rows - 1, menuPriceCol.ԭ�ɱ���) = zlStr.FormatEx(Nvl(rsTemp!ԭ�ɱ���, 0) * db��װϵ��, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.�ֳɱ���) = zlStr.FormatEx(rsTemp!�³ɱ��� * db��װϵ��, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.ԭ���ۼ�) = zlStr.FormatEx(IIf(IsNull(rsTemp!ԭ��), rsTemp!�ּ�, rsTemp!ԭ��) * db��װϵ��, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.�����ۼ�) = zlStr.FormatEx(rsTemp!�ּ� * db��װϵ��, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.ԭ�ɹ��޼�) = zlStr.FormatEx(rsTemp!ָ������ * db��װϵ��, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.�ֲɹ��޼�) = zlStr.FormatEx(rsTemp!ָ������ * db��װϵ��, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.ԭָ���ۼ�) = zlStr.FormatEx(rsTemp!ָ���ۼ� * db��װϵ��, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuPriceCol.��ָ���ۼ�) = zlStr.FormatEx(rsTemp!ָ���ۼ� * db��װϵ��, mintPriceDigit, , True)

                txtValuer.Text = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                txtSummary.Text = IIf(IsNull(rsTemp!����˵��), "", rsTemp!����˵��)
                If mintModal = 1 Then
                    Me.dtpRunDate.MinDate = CDate(rsTemp!ִ������)
                End If
                If IsNull(rsTemp!ִ������) Then
                    StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
                Else
                    StrToday = Format(rsTemp!ִ������, "yyyy-MM-dd hh:mm:ss")
                End If
                Me.dtpRunDate.Value = CDate(StrToday)

                .rows = .rows + 1
                Call setColEdit
                .RowHeight(.rows - 1) = mlngRowHeight
            End If
            rsTemp.MoveNext
        Next
        
        .colHidden(menuPriceCol.ԭ���ۼ�) = False
        .colHidden(menuPriceCol.ԭ�ɱ���) = False
        If mintMethod = 1 Then
            '���ɱ���
            .colHidden(menuPriceCol.ԭ���ۼ�) = True
        ElseIf mintMethod = 0 Then
            '���ۼ�
            .colHidden(menuPriceCol.ԭ�ɱ���) = True
        End If
        
        Call GetDrugStore(Val(.TextMatrix(1, menuPriceCol.ҩƷid)), 1)
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim strҩ�� As String
    Dim lngRow As Long

    '����ҩƷ
    On Error GoTo errHandle
    If strInput <> txtFind.Tag Then
        '��ʾ�µĲ���
        txtFind.Tag = strInput

        gstrSQL = "Select Distinct A.Id,'[' || A.���� || ']' As ҩƷ����, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                  "Where (A.վ�� = [3] Or A.վ�� is Null) And A.Id =B.�շ�ϸĿid And A.��� In ('5','6','7') " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] ) " & _
                  "Order By ҩƷ���� "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "ȡƥ���ҩƷID", strInput & "%", "%" & strInput & "%", gstrNodeNo)

        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If

    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    For n = 1 To mrsFindName.RecordCount
        '��������ˣ��򷵻ص�1����¼
        If mrsFindName.EOF Then mrsFindName.MoveFirst

        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = mrsFindName!ҩƷ���� & mrsFindName!ͨ����
        Else
            strҩ�� = mrsFindName!ҩƷ���� & IIf(IsNull(mrsFindName!��Ʒ��), mrsFindName!ͨ����, mrsFindName!��Ʒ��)
        End If

        For lngRow = 1 To vsfPrice.rows - 1
            lngFindRow = vsfPrice.FindRow(strҩ��, lngRow, CLng(menuPriceCol.ҩƷ), True, True)
            If lngFindRow > 0 Then
'                vsfPrice.Select lngFindRow, 1, lngFindRow, vsfPrice.Cols - 1
                vsfPrice.Row = lngFindRow
                vsfPrice.TopRow = lngFindRow
                Exit For
            End If
        Next

        If lngFindRow > 0 Then  '��ѯ�����ݺ���ƶ�����һ�����˳����β�ѯ
            mrsFindName.MoveNext
            Exit For
        Else
            mrsFindName.MoveNext 'δ��ѯ���������ƶ�����һ�����ݼ�������ѯ
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    If vsfPrice.Height + y <= 800 Then Exit Sub
    If TabCtlDetails.Height - y <= 1000 Then Exit Sub
    picSplit.Move 0, picSplit.Top + y
    vsfPrice.Move 0, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, vsfPrice.Height + y

    With TabCtlDetails
        .Top = picSplit.Top + picSplit.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = TabCtlDetails.Height - y
    End With
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub

    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub txtSummary_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If LenB(StrConv(txtSummary.Text, vbFromUnicode)) >= 100 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSummary_Validate(Cancel As Boolean)
    If LenB(StrConv(txtSummary.Text, vbFromUnicode)) > 100 Then
        MsgBox "˵��̫����", vbInformation, gstrSysName
        txtSummary.SelStart = 0
        txtSummary.SelLength = LenB(StrConv(txtSummary.Text, vbFromUnicode))
        Cancel = True
    End If
End Sub

Private Sub txt��Ӧ��_GotFocus()
    Me.txt��Ӧ��.SelStart = 0: Me.txt��Ӧ��.SelLength = Len(Me.txt��Ӧ��.Text)
End Sub

Private Sub txt��Ӧ��_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub

    strTmp = UCase(Trim(Me.txt��Ӧ��.Text))

    If strTmp = "" Then
        Me.txt��Ӧ��.Tag = "|"
        Exit Sub
    ElseIf strTmp = Split(Me.txt��Ӧ��.Tag, "|")(1) Then
        Exit Sub
    End If

    gstrSQL = "Select ����,����,����,id" & _
            " From ��Ӧ��" & _
            " where (���� Like [1] " & _
            "       Or ���� Like [2] " & _
            "       Or ���� Like [2])" & _
            " And ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, strTmp & "%", IIf(gstrMatchMethod = "0", "%", "") & strTmp & "%")

    With rsTemp
        If .EOF Then
            MsgBox "û���ҵ�ƥ��Ĺ�Ӧ�̣����ڹ�Ӧ�̹��������ӹ�Ӧ�̣�", vbInformation, gstrSysName
            Me.txt��Ӧ��.Text = Split(Me.txt��Ӧ��.Tag, "|")(1)
            Me.txt��Ӧ��.SelStart = 0: Me.txt��Ӧ��.SelLength = Len(Me.txt��Ӧ��.Text)
            Exit Sub
        End If

        If .RecordCount = 1 Then
            Me.txt��Ӧ��.Text = Trim(rsTemp!����): Me.txt��Ӧ��.Tag = rsTemp!Id & "|" & rsTemp!����
            Exit Sub
        Else
            With Me.mshProvider
                .Left = Me.chk��Ӧ��.Left
                .Top = Me.txt��Ӧ��.Top + Me.txt��Ӧ��.Height
                .Clear
                Set .DataSource = rsTemp
                .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
                .Row = 1: .ColSel = .Cols - 1
                .ZOrder 0: .Visible = True: .SetFocus
            End With
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub get�ֶμӳ��ۼ�(ByVal lngҩƷid As Long, ByVal lng����ϵ�� As Long, ByVal dbl�ɹ��� As Double, ByRef dbl�ۼ� As Double)
'���ܣ�ͨ���ɱ��۰��ֶμӳɷ�ʽ�����ۼ�
'�������ɱ���,�ۼ�
    Dim dbl��۶� As Double
    Dim blnData As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    mdbl�ֶμӳ��� = 0
    dbl��۶� = 0
    
    gstrSQL = "select ��� from  �շ���ĿĿ¼ a where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҩƷ���ʷ���", lngҩƷid)
    If rsTemp!��� = 7 Then
        mrs�ֶμӳ�.Filter = "����=1"
    Else
        mrs�ֶμӳ�.Filter = "����=0"
    End If
    
    If mrs�ֶμӳ�.RecordCount <> 0 Then
        mrs�ֶμӳ�.MoveFirst
        Do While Not mrs�ֶμӳ�.EOF
            With mrs�ֶμӳ�
                If dbl�ɹ��� > !��ͼ� And dbl�ɹ��� <= !��߼� Then
                    mdbl�ֶμӳ��� = IIf(IsNull(!�ӳ���), 0, !�ӳ���) / 100
                    dbl��۶� = IIf(IsNull(!��۶�), 0, !��۶�)
                    blnData = True
                    Exit Do
                End If
            End With
            mrs�ֶμӳ�.MoveNext
        Loop
    End If
    
    If blnData = False Then
        MsgBox "û�����ý���Ϊ��" & dbl�ɹ��� & "  �ķֶμӳ����ݣ�����ҩƷĿ¼�����ֶμӳ��ʣ������ã�", vbInformation, gstrSysName
        dbl�ۼ� = 0
        Exit Sub
    End If
    
    dbl�ۼ� = dbl�ɹ��� * (1 + mdbl�ֶμӳ���) + dbl��۶�
    
    Set rsTemp = Nothing
    gstrSQL = "Select ָ�����ۼ� From ҩƷ��� Where ҩƷID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", lngҩƷid)
    If rsTemp!ָ�����ۼ� * lng����ϵ�� < dbl�ۼ� Then
        dbl�ۼ� = rsTemp!ָ�����ۼ� * lng����ϵ��
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt��Ӧ��_Validate(Cancel As Boolean)
    If Me.txt��Ӧ��.Text = "" Then
        Me.txt��Ӧ��.Tag = "|"
    ElseIf Me.txt��Ӧ��.Text <> Split(Me.txt��Ӧ��.Tag, "|")(1) Then
        txt��Ӧ��_KeyPress (vbKeyReturn)
    End If
End Sub


Private Sub vsfPay_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfPay
        .Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height - 370
    End With
End Sub

Private Sub vsfPay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPay
        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
            Cancel = True
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub


Private Sub vsfPay_DblClick()
    With vsfPay
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPay_EnterCell()
    With vsfPay
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
    End With
End Sub

Private Sub vsfPay_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfPay
        If KeyCode = vbKeyReturn Then
            If .Col = menuPayCol.ҩƷ Then
                .Col = menuPayCol.��Ʊ��
            ElseIf .Col = menuPayCol.��Ʊ�� Then
                .Col = menuPayCol.��Ʊ����
            ElseIf .Col = menuPayCol.��Ʊ���� Then
                .Col = menuPayCol.��Ʊ���
            ElseIf .Col = menuPayCol.��Ʊ��� And .Row <> .rows - 1 Then
                .Col = menuPayCol.ҩƷ
                .Row = .Row + 1
            End If
        End If
    End With
End Sub

Private Sub vsfPay_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPay
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfPay_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfPay
            If Col = menuPayCol.��Ʊ��� Then
                strkey = .EditText
                intDigit = mintMoneyDigit
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Else
                    KeyAscii = 0
                End If
            ElseIf Col = menuPayCol.��Ʊ�� Then
                If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfPay_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strkey As String

    With vsfPay
        If Col = menuPayCol.��Ʊ���� Then
            strkey = .EditText
            If strkey <> "" Then
                If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                    strkey = TranNumToDate(strkey)
                    If strkey = "" Then
                        MsgBox "�Բ��𣬷�Ʊ���ڱ���Ϊ������,��ʽ(20000101����2000-01-01)��", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strkey
                    .TextMatrix(Row, menuPayCol.��Ʊ����) = .EditText
                End If
                
                If Not IsDate(strkey) Then
                    MsgBox "�Բ��𣬷�Ʊ���ڱ���Ϊ������(20000101����2000-01-01)��", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfprice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfPrice
        If Col = menuPriceCol.�ֳɱ��� Then
            If Val(.TextMatrix(Row, Col)) <> Val(.TextMatrix(Row, menuPriceCol.ԭ�ɱ���)) Then
                .Cell(flexcpFontBold, Row, Col, Row, Col) = 10
                .Cell(flexcpForeColor, Row, Col, Row, Col) = vbRed
            End If
        ElseIf Col = menuPriceCol.�����ۼ� Then
            If Val(.TextMatrix(Row, Col)) <> Val(.TextMatrix(Row, menuPriceCol.ԭ���ۼ�)) Then
                .Cell(flexcpFontBold, Row, Col, Row, Col) = 10
                .Cell(flexcpForeColor, Row, Col, Row, Col) = vbRed
            End If
        End If
    End With
End Sub

Private Sub vsfPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
'    Call SetRowHidden(Val(vsfPrice.TextMatrix(NewRow, menuPriceCol.ҩƷid)))
End Sub

Private Sub SetRowHidden(ByVal lngDrugID As Long)
    '���ܣ��е���ʾ������
    '������ҩƷid
    Dim intRow As Integer

    If lngDrugID = 0 Then Exit Sub
    With vsfStore
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With

    With vsfPay
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, menuPayCol.ҩƷid)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With
End Sub

'Private Sub vsfPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    With vsfPrice
'        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
'            Cancel = True
'            .Editable = flexEDNone
'        Else
'            .Editable = flexEDKbdMouse
'        End If
'    End With
'End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim mrsReturn As Recordset
    Dim vRect As RECT
    Dim dblLeft As Double
    Dim dblTop As Double

    mBlnClick = True
    vRect = zlControl.GetControlRect(vsfPrice.hWnd) '��ȡλ��
    dblLeft = vsfPrice.CellLeft
    dblTop = vRect.Top + vsfPrice.CellTop + vsfPrice.CellHeight


    On Error GoTo errHandle
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "", 0, , , , , , , , , True)
    End If
    Set mrsReturn = frmSelector.ShowME(Me, 0, 1, , dblLeft, dblTop, , , , , , , , , False, mstrPrivs)

    If mrsReturn.RecordCount = 0 Then Exit Sub
    mblnUpdateAdd = True
    Call GetDrugPirce(mrsReturn, Row)
    mblnUpdateAdd = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugPirce(ByVal rsReturn As ADODB.Recordset, ByVal Row As Integer)
    '������ȡҩƷ��Ϣ
    Dim rsTemp As Recordset
    Dim lngDrugID As Long
    Dim intRow As Long
    Dim i As Long
    Dim intCurrentPrice As Integer '�Ƿ���ʱ��
    Dim strUnit As String
    Dim db��װϵ�� As Double
    Dim strInfo As String

    On Error GoTo errHandle

    mlngOldDrugID = Val(vsfPrice.TextMatrix(Row, menuPriceCol.ҩƷid))
    Set rsReturn = CheckDoubleDrug(rsReturn)
    If rsReturn.RecordCount = 0 Then Exit Sub

    rsReturn.MoveFirst
    For i = 0 To rsReturn.RecordCount - 1
        With vsfPrice
            lngDrugID = rsReturn!ҩƷid

            '����Ƿ����Ϊִ�еļ۸�
            If checkNotExecutePrice(lngDrugID, strInfo) = True Then
                MsgBox strInfo, vbInformation, gstrSysName
                Exit Sub
            End If

            Select Case mintUnit
                Case 0
                    db��װϵ�� = rsReturn!ҩ���װ
                    strUnit = rsReturn!ҩ�ⵥλ
                Case 2
                    db��װϵ�� = rsReturn!סԺ��װ
                    strUnit = rsReturn!סԺ��λ
                Case 1
                    db��װϵ�� = rsReturn!�����װ
                    strUnit = rsReturn!���ﵥλ
                Case 3
                    db��װϵ�� = 1
                    strUnit = rsReturn!�ۼ۵�λ
            End Select

            .TextMatrix(Row, menuPriceCol.ҩƷid) = lngDrugID

            If gintҩƷ������ʾ = 1 Then
                .TextMatrix(Row, menuPriceCol.ҩƷ) = "[" & rsReturn!ҩƷ���� & "]" & IIf(IsNull(rsReturn!��Ʒ��), rsReturn!ͨ����, rsReturn!��Ʒ��)
            Else
                .TextMatrix(Row, menuPriceCol.ҩƷ) = "[" & rsReturn!ҩƷ���� & "]" & rsReturn!ͨ����
            End If

            .TextMatrix(Row, menuPriceCol.���) = IIf(IsNull(rsReturn!���), "", rsReturn!���)
            .TextMatrix(Row, menuPriceCol.�Ƿ���) = rsReturn!ʱ��
            .TextMatrix(Row, menuPriceCol.ҩ������) = IIf(rsReturn!ʱ�� = 0, "����", "ʱ��")
            intCurrentPrice = rsReturn!ʱ��
            .TextMatrix(Row, menuPriceCol.����) = IIf(IsNull(rsReturn!����), "", rsReturn!����)
            .TextMatrix(Row, menuPriceCol.��λ) = strUnit
            .TextMatrix(Row, menuPriceCol.��װϵ��) = db��װϵ��
            gstrSQL = "select ҩƷid from ҩƷ��� s where s.ҩƷid=[1] and s.����=1 And Not (zl_fun_getbatchpro(s.�ⷿid,[1])=1 And Nvl(S.����,0) = 0 And S.�������� < 0 And S.ʵ������ = 0 And S.ʵ�ʽ�� = 0 And S.ʵ�ʲ�� = 0) "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����", lngDrugID)
            If rsTemp.RecordCount = 0 Then
                .TextMatrix(Row, menuPriceCol.�Ƿ��п��) = 0
            Else
                .TextMatrix(Row, menuPriceCol.�Ƿ��п��) = 1
            End If

            If intCurrentPrice = 0 Then '����ҩƷ
                '��ʾ����ҩƷ���ۣ��ɱ���ȡƽ���۸��ۼ�ȡ�շѼ�Ŀ�ּ�
                gstrSQL = "Select b.Id, Decode(k.�ɱ���, Null, a.�ɱ���*" & db��װϵ�� & ", k.�ɱ���) As �ɱ���, a.ָ��������, a.ָ�����ۼ�, b.�ּ�*" & db��װϵ�� & " as �ּ�, a.���������, a.�ӳ��� / 100 As �ӳ���," & vbNewLine & _
                    "            b.������Ŀid" & vbNewLine & _
                    "     From ҩƷ��� A, �շѼ�Ŀ B," & vbNewLine & _
                    "          (Select Decode(Sum(Nvl(ʵ������, 0)), 0, Null, Sum(Round(ƽ���ɱ���*" & db��װϵ�� & ", " & mintCostDigit & ") * round(ʵ������/" & db��װϵ�� & "," & mintNumberDigit & ")) / Sum(round(ʵ������/" & db��װϵ�� & "," & mintNumberDigit & "))) As �ɱ���" & vbNewLine & _
                    "            From ҩƷ���" & vbNewLine & _
                    "            Where ���� = 1 And ҩƷid = [1] ) K" & vbNewLine & _
                    "     Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = [1] And Sysdate Between ִ������ And ��ֹ����" & GetPriceClassString("B")
            Else 'ʱ��ҩƷ
                '��ʾʱ��ҩƷ���ۣ�ȡ�����/���������Ϊ��۸�
                gstrSQL = "Select p.Id, Nvl(k.�ּ�, Nvl(j.�ϴ��ۼ�*" & db��װϵ�� & ",p.�ּ�*" & db��װϵ�� & ")) as �ּ�, j.�ӳ��� / 100 As �ӳ���, Nvl(k.�ɱ���, j.�ɱ���*" & db��װϵ�� & ") As �ɱ���, j.ָ��������," & vbNewLine & _
                    "       j.ָ�����ۼ�, j.���������, p.������Ŀid, p.ִ������, p.������Ŀid, i.���� As ��������" & vbNewLine & _
                    "From �շѼ�Ŀ P, ������Ŀ I, ҩƷ��� J," & vbNewLine & _
                    "     (Select Decode(Sum(Nvl(ʵ������, 0)), 0, Null, Sum(Round(���ۼ�*" & db��װϵ�� & ", " & mintPriceDigit & ") * round(ʵ������/" & db��װϵ�� & "," & mintNumberDigit & ")) / Sum(round(ʵ������/" & db��װϵ�� & "," & mintNumberDigit & "))) As �ּ�," & vbNewLine & _
                    "              Decode(Sum(Nvl(ʵ������, 0)), 0, Null, Sum(Round(ƽ���ɱ���*" & db��װϵ�� & ", " & mintCostDigit & ") * round(ʵ������/" & db��װϵ�� & "," & mintNumberDigit & ")) / Sum(round(ʵ������/" & db��װϵ�� & "," & mintNumberDigit & "))) As �ɱ���" & vbNewLine & _
                    "       From ҩƷ���" & vbNewLine & _
                    "       Where ���� = 1 And ҩƷid = [1] ) K" & vbNewLine & _
                    "Where p.������Ŀid = i.Id And p.�շ�ϸĿid = j.ҩƷid And p.�շ�ϸĿid = [1] And" & vbNewLine & _
                    "      (p.��ֹ���� Is Null Or Sysdate Between p.ִ������ And p.��ֹ����) " & GetPriceClassString("P")

            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯҩƷ", lngDrugID)
            If rsTemp.RecordCount = 0 Then
                MsgBox "��ҩƷ�����ڣ������½�����ҩƷ��Ƭ��", vbInformation, gstrSysName
                Exit Sub
            End If
            .TextMatrix(Row, menuPriceCol.ԭ��id) = rsTemp!Id
            .TextMatrix(Row, menuPriceCol.������ĿID) = IIf(IsNull(rsTemp!������ĿID), 0, rsTemp!������ĿID)
            .TextMatrix(Row, menuPriceCol.�ӳ���) = zlStr.FormatEx(IIf(IsNull(rsTemp!�ӳ���), 0, rsTemp!�ӳ���), 5, , True)
            .TextMatrix(Row, menuPriceCol.���������) = IIf(IsNull(rsTemp!���������), 100, rsTemp!���������)
            
            '�ɱ��ۣ��ۼ۲��ð�װ���㣬��֮ǰ��SQL���Ѿ�������
            .TextMatrix(Row, menuPriceCol.ԭ�ɱ���) = zlStr.FormatEx(IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���), mintCostDigit, , True)
            .TextMatrix(Row, menuPriceCol.�ֳɱ���) = zlStr.FormatEx(IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���), mintCostDigit, , True)
            .TextMatrix(Row, menuPriceCol.ԭ���ۼ�) = zlStr.FormatEx(IIf(IsNull(rsTemp!�ּ�), 0, rsTemp!�ּ�), mintPriceDigit, , True)
            .TextMatrix(Row, menuPriceCol.�����ۼ�) = zlStr.FormatEx(IIf(IsNull(rsTemp!�ּ�), 0, rsTemp!�ּ�), mintPriceDigit, , True)
            
            .TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�) = zlStr.FormatEx(IIf(IsNull(rsTemp!ָ��������), 0, rsTemp!ָ��������) * db��װϵ��, mintCostDigit, , True)
            .TextMatrix(Row, menuPriceCol.�ֲɹ��޼�) = .TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�)
            .TextMatrix(Row, menuPriceCol.ԭָ���ۼ�) = zlStr.FormatEx(IIf(IsNull(rsTemp!ָ�����ۼ�), 0, rsTemp!ָ�����ۼ�) * db��װϵ��, mintPriceDigit, , True)
            .TextMatrix(Row, menuPriceCol.��ָ���ۼ�) = .TextMatrix(Row, menuPriceCol.ԭָ���ۼ�)

            Call GetDrugStore(lngDrugID, Row)
            If Row = .rows - 1 Then '���һ�в�������
                .rows = .rows + 1
                .RowHeight(.rows - 1) = mlngRowHeight
                Row = Row + 1
            End If
        End With
'        If mint���� = 0 And mblnʱ��ҩƷ�����ε��� = True Then '�ۼ۵���
'            Call GetDrugStore(lngDrugID, db��װϵ��)
'        ElseIf mint���� <> 0 Then

'        End If
'        Call SetRowHidden(lngDrugID)

        rsReturn.MoveNext
    Next
    Call setColEdit

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugStore(ByVal lngDrugID As Long, ByVal intRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim dblOldCost As Double
    Dim dblOldPrice As Double
    Dim dblNewCost As Double
    Dim dblNewPrice As Double
    Dim dbl�ӳ��� As Double
    Dim lngCurRow As Long     '��ǰ��
    Dim i As Long
    Dim dbl��Ʊ��� As Double
    Dim strҩƷ���� As String
    Dim str��Ʊ As String
    Dim str��Ʊ���� As String
    Dim rsPirce As ADODB.Recordset
    Dim rsCost As ADODB.Recordset
    Dim dbl��װ���� As Double
    Dim bln��ͬҩƷ As Boolean
    Dim lngҩƷid As Long
    Dim str��λ As String
    Dim bln�Ƿ�ִ�� As Boolean
    
    '���ܣ�Ϊ����б��������
    '������ҩƷid

    On Error GoTo errHandle
    '�ȼ���Ƿ����ظ������ݣ�����о���������ظ�������
    With vsfStore
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    With vsfPay
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuPayCol.ҩƷid)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
        gstrSQL = "Select s.�ⷿid,s.ҩƷid, d.���� As �ⷿ, '[' || m.���� || ']' || m.���� As ҩƷ, m.���, m.����, m.���㵥λ �ۼ۵�λ, p.ҩ�ⵥλ, s.�ϴ����� As ����, nvl(s.ʵ������,0) As ����," & vbNewLine & _
            "       s.����, Nvl(m.�Ƿ���, 0) ���, m.Id, Decode(Nvl(m.�Ƿ���, 0), 0, e.�ּ�, Decode(s.���ۼ�,null,Decode(Nvl(s.ʵ������, 0), 0, e.�ּ�, s.ʵ�ʽ�� / s.ʵ������),s.���ۼ�)) As ʱ���ۼ�, p.�ӳ���," & vbNewLine & _
            "       Decode(s.ƽ���ɱ���, null, p.�ɱ���, s.ƽ���ɱ���) As �ɱ���, s.�ϴι�Ӧ��id, n.���� As ��Ӧ��, s.Ч��, s.�ϴβ��� As ����" & vbNewLine & _
            " From ҩƷ��� S, ���ű� D, �շ���ĿĿ¼ M, ҩƷ��� P, ��Ӧ�� N, �շѼ�Ŀ E" & vbNewLine & _
            " Where d.Id = s.�ⷿid And s.ҩƷid = m.Id And m.Id = p.ҩƷid And Nvl(s.�ϴι�Ӧ��id, 0) = n.Id(+) And m.Id = e.�շ�ϸĿid And" & vbNewLine & _
            " s.���� = 1 And s.ҩƷid = [1] And Sysdate Between e.ִ������ And e.��ֹ����  " & vbNewLine & _
            " And Not (zl_fun_getbatchpro(s.�ⷿid,[1])=1 And Nvl(S.����,0) = 0 And S.�������� < 0 And S.ʵ������ = 0 And S.ʵ�ʽ�� = 0 And S.ʵ�ʲ�� = 0) " & vbNewLine & _
            GetPriceClassString("E") & vbNewLine & _
            " Order By s.ҩƷid,s.�ⷿid, s.�ϴ�����,s.���� "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

        If mlng��Ӧ��ID > 0 Then
            rsTemp.Filter = "�ϴι�Ӧ��ID=" & mlng��Ӧ��ID
        End If
    Else '�޸ģ�����
        gstrSQL = "select (sysdate-ִ������ ) as �Ƿ�ִ�� from ���ۻ��ܼ�¼ where ���ۺ�=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�ִ��", txtNO.Text)
          
        bln�Ƿ�ִ�� = rsTemp!�Ƿ�ִ�� > 0
        
        If cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
            If bln�Ƿ�ִ�� = True Then
                gstrSQL = "Select Distinct b.�ⷿid, c.���� As �ⷿ, b.ҩƷid, b.��ҩ��λid as �ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��, b.�³ɱ���," & vbNewLine & _
                                "                b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���, b.����, b.����, b.����, e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ," & vbNewLine & _
                                "                Nvl(b.ʵ������, Nvl(b.ʵ������, 0)) As ����, f.�ӳ���, b.Ч��, Decode(Nvl(e.�Ƿ���, 0), 0, h.ԭ��, b.ԭ���ۼ�) As ԭ���ۼ�," & vbNewLine & _
                                "                Decode(Nvl(e.�Ƿ���, 0), 0, h.�ּ�, Nvl(b.�����ۼ�, h.�ּ�)) As �����ۼ�" & vbNewLine & _
                                "From ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F, �շѼ�Ŀ H," & vbNewLine & _
                                "     (Select Distinct b.ҩƷid, b.�ⷿid, b.����, b.����, b.Ч��, b.����, b.ԭ�� As ԭ�ɱ���, b.�ּ� As �³ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���, b.Ӧ����䶯," & vbNewLine & _
                                "                       b.ִ������, g.ԭ�� As ԭ���ۼ�, g.�ּ� As �����ۼ�, b.���ۻ��ܺ�, i.��д���� As ʵ������, b.��ҩ��λid" & vbNewLine & _
                                "       From ҩƷ�۸��¼ B, ҩƷ�۸��¼ G, ҩƷ�շ���¼ I" & vbNewLine & _
                                "       Where b.�۸����� = 2 And b.���ۻ��ܺ� =[1] And" & vbNewLine & _
                                "             Decode(b.�ⷿid, Null, 1, b.�ⷿid) = Decode(b.�ⷿid, Null, 1, g.�ⷿid(+)) And b.ҩƷid = g.ҩƷid(+) And" & vbNewLine & _
                                "             Decode(b.�ⷿid, Null, 1, Nvl(b.����,0)) = Decode(b.�ⷿid, Null, 1, Nvl(g.����(+),0)) And b.���ۻ��ܺ� = g.���ۻ��ܺ�(+) And g.�۸�����(+) = 1 And b.�շ�id = i.Id) B" & vbNewLine & _
                                "Where b.�ⷿid = c.Id And b.��ҩ��λid = d.Id(+) And b.ҩƷid = e.Id And e.Id = f.ҩƷid And b.ҩƷid = h.�շ�ϸĿid And h.���ۻ��ܺ� = b.���ۻ��ܺ�" & vbNewLine & _
                                "Order By b.ҩƷid, b.�ⷿid, b.����, b.����"
            Else
                gstrSQL = "Select Distinct a.�ⷿid,c.���� as �ⷿ, b.ҩƷid,a.�ϴι�Ӧ��id, '[' || e.���� || ']' ||e.���� as ҩƷ,e.���,d.���� as ��Ӧ��, b.�³ɱ���, b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���" & _
                                    " ,a.�ϴβ��� as ����,a.����,a.�ϴ����� as ����,e.�Ƿ��� as ���,e.���㵥λ as �ۼ۵�λ,f.ҩ�ⵥλ,Nvl(b.ʵ������, Nvl(a.ʵ������, 0)) as ����,f.�ӳ���,a.Ч��, " & _
                                    " Decode(Nvl(e.�Ƿ���, 0), 0, h.ԭ��, b.ԭ���ۼ�) As ԭ���ۼ�,Decode(Nvl(e.�Ƿ���, 0), 0, h.�ּ�, Nvl(b.�����ۼ�, h.�ּ�)) As �����ۼ� " & _
                                    " From ҩƷ��� A,���ű� C,��Ӧ�� D,�շ���ĿĿ¼ E,ҩƷ��� F,�շѼ�Ŀ H, " & _
                                         " (Select Distinct b.ҩƷid, b.�ⷿid, b.����, b.����, b.Ч��, b.����, b.ԭ�� as ԭ�ɱ���, b.�ּ� as �³ɱ���, b.��Ʊ��, b.��Ʊ����," & _
                                         " b.��Ʊ���, b.Ӧ����䶯, b.ִ������, g.ԭ�� As ԭ���ۼ�, g.�ּ� As �����ۼ�,b.���ۻ��ܺ�, i.��д���� As ʵ������ " & _
                                           " From ҩƷ�۸��¼ B, ҩƷ�۸��¼ G, ҩƷ�շ���¼ I " & _
                                           " Where B.�۸�����=2 And B.���ۻ��ܺ� = [1] " & _
                                    " And Decode(b.�ⷿid, Null, 1,b.�ⷿid) = Decode(b.�ⷿid, Null, 1,g.�ⷿid(+)) And " & _
                                    " b.ҩƷid = g.ҩƷid(+) And Decode(b.�ⷿid, Null, 1,Nvl(b.����,0)) = Decode(b.�ⷿid, Null, 1,Nvl(g.����(+),0)) And b.���ۻ��ܺ� = g.���ۻ��ܺ�(+) And g.�۸�����(+) = 1 And b.�շ�id = i.Id(+)) B" & _
                                    " Where a.ҩƷid = b.ҩƷid And Decode(b.�ⷿid, Null, 1,a.�ⷿid) = Decode(b.�ⷿid, Null, 1,b.�ⷿid) and " & _
                                    " Decode(b.�ⷿid, Null, 1,nvl(a.����,0))=Decode(b.�ⷿid, Null, 1,nvl(b.����,0)) and a.�ⷿid=c.id and a.�ϴι�Ӧ��id=d.id(+) and " & _
                                    " a.ҩƷid=e.id and e.id=f.ҩƷid and a.����=1 And a.ҩƷid = h.�շ�ϸĿid And h.���ۻ��ܺ� = b.���ۻ��ܺ� " & _
                                    " And Not (zl_fun_getbatchpro(a.�ⷿid,a.ҩƷid)=1 And Nvl(a.����,0) = 0 And a.�������� < 0 And a.ʵ������ = 0 And a.ʵ�ʽ�� = 0 And a.ʵ�ʲ�� = 0) " & _
                                    " Order By b.ҩƷid, a.�ⷿid, a.�ϴ�����,a.���� "
            End If
        ElseIf cboPriceMethod.Text = "�����ɱ���" Then
            If bln�Ƿ�ִ�� = True Then
                '�Ѿ�ִ����ȡ�Ѳ������շ���¼����
                gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.ҩƷid, b.��ҩ��λid As �ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & vbNewLine & _
                        "                b.�ּ� as �³ɱ���, b.ԭ�� as ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���, b.����, b.����, b.����, e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ," & vbNewLine & _
                        "                nvl(a.��д����,0) As ����, f.�ӳ���, b.Ч��, g.ԭ�� As ԭ���ۼ�, g.�ּ� As �����ۼ� " & vbNewLine & _
                        "From ҩƷ�շ���¼ A, ҩƷ�۸��¼ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F, �շѼ�Ŀ G " & vbNewLine & _
                        "Where a.id=b.�շ�id And a.�ⷿid = c.Id And b.��ҩ��λid = d.Id(+) And" & vbNewLine & _
                        "      a.ҩƷid = e.Id And e.Id = f.ҩƷid And b.�۸�����=2 And b.���ۻ��ܺ� = [1] and a.���� = 5 " & vbNewLine & _
                        " And b.ҩƷid = g.�շ�ϸĿid And Sysdate Between g.ִ������ And g.��ֹ���� " & _
                        " Order By b.ҩƷid, a.�ⷿid, b.����,b.���� "
            Else
                'δִ��ȡ�۸��¼���������Ϣ
                gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ҩƷid, a.�ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & _
                        " g.�ּ� As �³ɱ���, g.ԭ�� As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.�ϴβ��� As ����, a.����, a.�ϴ����� As ����," & _
                        " e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ, nvl(a.ʵ������,0) As ����, f.�ӳ���, a.Ч��, " & _
                        " Decode(Nvl(e.�Ƿ���, 0), 0, b.ԭ��, a.���ۼ�) As ԭ���ۼ�,Decode(Nvl(e.�Ƿ���, 0), 0, b.�ּ�, a.���ۼ�)  As �����ۼ� " & _
                        " From ҩƷ��� A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F, ҩƷ�۸��¼ G " & _
                        " Where a.ҩƷid = b.�շ�ϸĿid And a.�ⷿid = c.Id And a.�ϴι�Ӧ��id = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.ҩƷid And a.���� = 1 And" & _
                        " g.���ۻ��ܺ� = [1] " & GetPriceClassString("B") & _
                        " And Decode(g.�ⷿid, Null, 1,a.�ⷿid) = Decode(g.�ⷿid, Null, 1,g.�ⷿid) And " & _
                        " a.ҩƷid = g.ҩƷid And Decode(g.�ⷿid, Null, 1,Nvl(a.����,0)) = Decode(g.�ⷿid, Null, 1,Nvl(g.����,0)) " & _
                        " And Sysdate Between b.ִ������ And b.��ֹ���� And g.�۸����� = 2 " & _
                        " And Not (zl_fun_getbatchpro(a.�ⷿid,a.ҩƷid)=1 And Nvl(a.����,0) = 0 And a.�������� < 0 And a.ʵ������ = 0 And a.ʵ�ʽ�� = 0 And a.ʵ�ʲ�� = 0) " & _
                        " Order By ҩƷid, �ⷿid, ����,���� "
            
            End If
        ElseIf cboPriceMethod.Text = "�����ۼ�" Then
            If bln�Ƿ�ִ�� = True Then
                '�Ѿ�ִ����ȡ�Ѳ������շ���¼����
                gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ҩƷid, a.��ҩ��λid As �ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���," & vbNewLine & _
                        "                d.���� As ��Ӧ��, nvl(h.ƽ���ɱ���,f.�ɱ���) As �³ɱ���, nvl(h.ƽ���ɱ���,f.�ɱ���) As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.����, a.����, a.����, e.�Ƿ��� As ���," & vbNewLine & _
                        "                e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ, nvl(a.��д����,0) As ����, f.�ӳ���, a.Ч��,a.�ɱ��� As ԭ���ۼ�, a.���ۼ� As �����ۼ� " & vbNewLine & _
                        "From ҩƷ�շ���¼ A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F, ҩƷ��� H " & vbNewLine & _
                        "Where a.�۸�id = b.Id And a.�ⷿid = c.Id And a.��ҩ��λid = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.ҩƷid And" & vbNewLine & _
                        "      b.���ۻ��ܺ� = [1] and a.����=13 And a.����id Is Null and a.�ⷿid=h.�ⷿid(+) and a.ҩƷid=h.ҩƷid(+) and Nvl(a.����,0)=nvl(h.����(+),0) and h.����(+)=1 " & GetPriceClassString("B") & _
                        " Order By b.�շ�ϸĿid, a.�ⷿid, a.����,a.���� "
            Else
                'δִ��ȡ�۸��¼���������Ϣ
                gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ҩƷid, a.�ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & _
                                        " nvl(a.ƽ���ɱ���,f.�ɱ���) As �³ɱ���, nvl(a.ƽ���ɱ���,f.�ɱ���) As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.�ϴβ��� As ����, a.����, a.�ϴ����� As ����," & _
                                        " e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ, nvl(a.ʵ������,0) As ����, f.�ӳ���, a.Ч��,b.ԭ�� As ԭ���ۼ�,b.�ּ� As �����ۼ� " & _
                        " From ҩƷ��� A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F " & _
                        " Where a.ҩƷid = b.�շ�ϸĿid And a.�ⷿid = c.Id And a.�ϴι�Ӧ��id = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.ҩƷid And a.���� = 1 And" & _
                              " b.���ۻ��ܺ� = [1] And Nvl(e.�Ƿ���, 0) = 0 " & GetPriceClassString("B") & _
                              " And Not (zl_fun_getbatchpro(a.�ⷿid,a.ҩƷid)=1 And Nvl(a.����,0) = 0 And a.�������� < 0 And a.ʵ������ = 0 And a.ʵ�ʽ�� = 0 And a.ʵ�ʲ�� = 0) "
                gstrSQL = gstrSQL & " Union All " & _
                        "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ҩƷid, a.�ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & _
                                        " nvl(a.ƽ���ɱ���,f.�ɱ���) As �³ɱ���, nvl(a.ƽ���ɱ���,f.�ɱ���) As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.�ϴβ��� As ����, a.����, a.�ϴ����� As ����," & _
                                        " e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ, nvl(a.ʵ������,0) As ����, f.�ӳ���, a.Ч��,g.ԭ�� As ԭ���ۼ�,g.�ּ� As �����ۼ� " & _
                        " From ҩƷ��� A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F, ҩƷ�۸��¼ G " & _
                        " Where a.ҩƷid = b.�շ�ϸĿid And a.�ⷿid = c.Id And a.�ϴι�Ӧ��id = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.ҩƷid And a.���� = 1 And" & _
                        " b.���ۻ��ܺ� = [1] And Nvl(e.�Ƿ���, 0) = 1 " & GetPriceClassString("B") & _
                        " And Decode(g.�ⷿid, Null, 1,a.�ⷿid) = Decode(g.�ⷿid, Null, 1,g.�ⷿid) And " & _
                        " a.ҩƷid = g.ҩƷid And Decode(g.�ⷿid, Null, 1,Nvl(a.����,0)) = Decode(g.�ⷿid, Null, 1,Nvl(g.����,0)) And b.���ۻ��ܺ� = g.���ۻ��ܺ� And g.�۸����� = 1 " & _
                        " And Not (zl_fun_getbatchpro(a.�ⷿid,a.ҩƷid)=1 And Nvl(a.����,0) = 0 And a.�������� < 0 And a.ʵ������ = 0 And a.ʵ�ʽ�� = 0 And a.ʵ�ʲ�� = 0) " & _
                        " Order By ҩƷid, �ⷿid, ����,���� "
            End If
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, txtNO.Text)
    End If

    With vsfStore
        Do While Not rsTemp.EOF
            dbl��װ���� = 0
            dbl��Ʊ��� = 0
            dblOldPrice = 0
            dblNewPrice = 0
            For i = 0 To vsfPrice.rows - 1
                If rsTemp!ҩƷid = vsfPrice.TextMatrix(i, menuPriceCol.ҩƷid) Then
                    dbl��װ���� = vsfPrice.TextMatrix(i, menuPriceCol.��װϵ��)
                    dblOldPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.ԭ���ۼ�))
                    dblNewPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.�����ۼ�))
                    str��λ = vsfPrice.TextMatrix(i, menuPriceCol.��λ)
                    Exit For
                End If
            Next
            .rows = .rows + 1
            .TextMatrix(.rows - 1, menuStoreCol.���) = rsTemp!���
            Call setColEdit
            .RowHeight(.rows - 1) = mlngRowHeight

            '�ӿհ��п�ʼ��������
            .TextMatrix(.rows - 1, menuStoreCol.ҩƷid) = rsTemp!ҩƷid
            .TextMatrix(.rows - 1, menuStoreCol.�ⷿ) = rsTemp!�ⷿ
            .TextMatrix(.rows - 1, menuStoreCol.�ⷿid) = rsTemp!�ⷿid
            .TextMatrix(.rows - 1, menuStoreCol.��Ӧ��) = Nvl(rsTemp!��Ӧ��, "")
            .TextMatrix(.rows - 1, menuStoreCol.��Ӧ��id) = IIf(mlng��Ӧ��ID > 0, mlng��Ӧ��ID, Nvl(rsTemp!�ϴι�Ӧ��ID))
            .TextMatrix(.rows - 1, menuStoreCol.ҩƷ) = rsTemp!ҩƷ
            strҩƷ���� = rsTemp!ҩƷ

            .TextMatrix(.rows - 1, menuStoreCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            .TextMatrix(.rows - 1, menuStoreCol.��λ) = str��λ
            .TextMatrix(.rows - 1, menuStoreCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.rows - 1, menuStoreCol.Ч��) = Format(IIf(IsNull(rsTemp!Ч��), "", rsTemp!Ч��), "YYYY-MM-DD")
            .TextMatrix(.rows - 1, menuStoreCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.rows - 1, menuStoreCol.����) = zlStr.FormatEx(rsTemp!���� / dbl��װ����, mintNumberDigit, , True)
            .TextMatrix(.rows - 1, menuStoreCol.��װϵ��) = dbl��װ����
            .TextMatrix(.rows - 1, menuStoreCol.����) = Nvl(rsTemp!����, 0)
'            .TextMatrix(.rows - 1, menuStoreCol.���) = rsTemp!���


            If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
                dblOldCost = IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���) * dbl��װ����

                If mdbl�ӳ��� > 0 Then
                    dbl�ӳ��� = Round(mdbl�ӳ��� / 100, 7)
                ElseIf dblOldCost > 0 Then
                    dbl�ӳ��� = Round(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice) / dblOldCost - 1, 7)
                Else
                    dbl�ӳ��� = Round(rsTemp!�ӳ��� / 100, 2)
                End If
                If 1 + dbl�ӳ��� = 0 Then
                    dblNewCost = 0
                Else
                    dblNewCost = rsTemp!ʱ���ۼ� * dbl��װ���� / (1 + dbl�ӳ���)
                End If
                If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0

                .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = zlStr.FormatEx(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice), mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = zlStr.FormatEx(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice), mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�ۼ�ӯ��) = Format(Format(rsTemp!���� / dbl��װ���� * Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)), mstrMoneyFormat) - Format(rsTemp!���� / dbl��װ���� * Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�)), mstrMoneyFormat), mstrMoneyFormat)
                
                .TextMatrix(.rows - 1, menuStoreCol.�ӳ���) = zlStr.FormatEx(zlStr.FormatEx(dbl�ӳ���, 5, , True) * 100, 5, , True)
                .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɱ���) = zlStr.FormatEx(dblOldCost, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(dblNewCost, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�ɱ�ӯ��) = Format(Format(Val(.TextMatrix(.rows - 1, menuStoreCol.�ֳɱ���)) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat) - Format(Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ�ɱ���)) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat), mstrMoneyFormat)
                dbl��Ʊ��� = dbl��Ʊ��� + (dblNewCost - dblOldCost) * Val(.TextMatrix(.rows - 1, menuStoreCol.����))
                
                'ΪӦ����¼��ֵ
                If mint���� = 1 Or mint���� = 2 Then
                    If vsfPay.rows > 1 Then
                        bln��ͬҩƷ = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.ҩƷid) = rsTemp!ҩƷid Then
                                bln��ͬҩƷ = True
                                Exit For
                            End If
                        Next
                        If bln��ͬҩƷ = True Then
                            vsfPay.TextMatrix(i, menuPayCol.��Ʊ���) = zlStr.FormatEx(Val(vsfPay.TextMatrix(i, menuPayCol.��Ʊ���)) + dbl��Ʊ���, mintMoneyDigit, , True)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷ) = strҩƷ����
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = zlStr.FormatEx(dbl��Ʊ���, mintMoneyDigit, , True)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷ) = strҩƷ����
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = zlStr.FormatEx(dbl��Ʊ���, mintMoneyDigit, , True)
                    End If
                End If
            Else
                .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = zlStr.FormatEx(Val(rsTemp!ԭ���ۼ�) * dbl��װ����, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = zlStr.FormatEx(Val(rsTemp!�����ۼ�) * dbl��װ����, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�ۼ�ӯ��) = Format(Format(rsTemp!���� / dbl��װ���� * Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)), mstrMoneyFormat) - Format(rsTemp!���� / dbl��װ���� * Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�)), mstrMoneyFormat), mstrMoneyFormat)
                .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɱ���) = zlStr.FormatEx(Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(rsTemp!�³ɱ��� * dbl��װ����, mintCostDigit, , True)
                .TextMatrix(.rows - 1, menuStoreCol.�ɱ�ӯ��) = Format(Format((rsTemp!�³ɱ��� * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat) - Format((Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat), mstrMoneyFormat)
                 
                If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    If rsTemp!�³ɱ��� = 0 Then
                        dbl�ӳ��� = 0
                    Else
                        dbl�ӳ��� = Round(Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) / (rsTemp!�³ɱ��� * dbl��װ����) - 1, 7)
                    End If
                    .TextMatrix(.rows - 1, menuStoreCol.�ӳ���) = zlStr.FormatEx(zlStr.FormatEx(dbl�ӳ���, 5, , True) * 100, 5, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɱ���) = zlStr.FormatEx(Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(rsTemp!�³ɱ��� * dbl��װ����, mintCostDigit, , True)
                    .TextMatrix(.rows - 1, menuStoreCol.�ɱ�ӯ��) = Format(Format((rsTemp!�³ɱ��� * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat) - Format((Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat), mstrMoneyFormat)
                    dbl��Ʊ��� = dbl��Ʊ��� + (rsTemp!�³ɱ��� * dbl��װ���� - Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����))
                    str��Ʊ = IIf(IsNull(rsTemp!��Ʊ��), "", rsTemp!��Ʊ��)
                    str��Ʊ���� = IIf(IsNull(rsTemp!��Ʊ����), "", rsTemp!��Ʊ����)
                    
                    'Ϊ�����¼�б�ֵ
                    If vsfPay.rows > 1 Then
                        bln��ͬҩƷ = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.ҩƷid) = rsTemp!ҩƷid Then
                                bln��ͬҩƷ = True
                                Exit For
                            End If
                        Next
                        If bln��ͬҩƷ = True Then
                            vsfPay.TextMatrix(i, menuPayCol.��Ʊ���) = zlStr.FormatEx(Val(vsfPay.TextMatrix(i, menuPayCol.��Ʊ���)) + dbl��Ʊ���, mintMoneyDigit, , True)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷ) = strҩƷ����
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = zlStr.FormatEx(dbl��Ʊ���, mintMoneyDigit, , True)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷ) = strҩƷ����
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = zlStr.FormatEx(dbl��Ʊ���, mintMoneyDigit, , True)
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
        
    End With
    '�޸ĺͲ���ʱ�������б�ƽ���ɱ��ۣ��ۼ�
    'mintModal 0-���� 1-�޸� 2-����
    If mintModal = 1 Or mintModal = 2 Then
        With vsfStore
            For i = 1 To .rows - 1
                If lngҩƷid <> .TextMatrix(i, menuStoreCol.ҩƷid) Then
                    Call CaluateAverCost(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    Call CaluateAverOldCost(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    
                    If Val(.TextMatrix(i, menuStoreCol.���)) = 1 Then
                        Call CaculateAverPirce(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                        Call CaculateAverOldPirce(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    End If
                    
                    lngҩƷid = Val(.TextMatrix(i, menuStoreCol.ҩƷid))
                End If
            Next
        End With
    End If

    If mint���� = 1 Or mint���� = 2 Then
        If rsTemp.RecordCount = 0 Then Exit Sub
        TabCtlDetails.Item(1).Visible = True
    End If
            
    Call setColHiddenVsf
    
    '�ϲ���Ԫ��
    vsfStore.MergeCol(menuStoreCol.ҩƷ) = True
    vsfStore.MergeCol(menuStoreCol.���) = True
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfPrice_DblClick()
    With vsfPrice
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPrice_EnterCell()
    Dim i As Integer, j As Integer
    Dim intRow As Integer

    With vsfPrice
        .Editable = flexEDNone
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If

        If .Col = menuPriceCol.�����ۼ� Then
            mdblOldPrice = Val(vsfPrice.TextMatrix(.Row, menuPriceCol.�����ۼ�))
        ElseIf .Col = menuPriceCol.�ֳɱ��� Then
            mdblOldPrice = Val(vsfPrice.TextMatrix(.Row, menuPriceCol.�ֳɱ���))
        End If
    End With
    With vsfStore
        If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.ҩƷid)) = 0 Then Exit Sub
        If .rows > 1 Then
            .Select 0, 0, 0, 0
            For i = 1 To .rows - 1
                If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.ҩƷid)) = Val(.TextMatrix(i, menuStoreCol.ҩƷid)) Then
                    If j = 0 Then j = i
                    .Select j, 3, j, .Cols - 1
                    .TopRow = j
                    intRow = intRow + 1
                End If
                .CellBorderRange i, 0, i, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            For i = j To j + intRow - 1
                If i = j Then .CellBorderRange i, 0, i, .Cols - 1, mlngBorderColor, 0, 2, 0, 0, 0, 2
                If i = j + intRow - 1 Then .CellBorderRange i, 0, i, .Cols - 1, mlngBorderColor, 0, 0, 0, 2, 0, 2
                If i = j And i = j + intRow - 1 Then .CellBorderRange i, 0, i, .Cols - 1, mlngBorderColor, 0, 2, 0, 2, 0, 2
            Next
        End If
    End With
    
    Call SetBorder '������ѡ�б߿�
End Sub

Private Sub SetBorder()
    '������ѡ�б߿�
    Dim intRow As Integer
    
    With vsfPrice
        If .rows <> 1 Then
            For intRow = 1 To .rows - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            .CellBorderRange .Row, menuPriceCol.ҩƷ, .Row, menuPriceCol.�����ۼ�, mlngBorderColor, 0, 2, 0, 2, 0, 2
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim lngDrugID As Long
    Dim strRow As String
    Dim intɾ������ As Integer
    
    With vsfPrice
        If KeyCode = vbKeyReturn Then
            If .Col <> menuPriceCol.�����ۼ� Then '�ɱ��۵���
                If .Col = menuPriceCol.ҩƷ And cboPriceMethod.Text = "�����ɱ���" Then
                    .Col = menuPriceCol.�ֳɱ���
'                    .EditCell
                ElseIf .Col = menuPriceCol.ҩƷ And cboPriceMethod.Text = "�����ۼ�" Then
                    .Col = menuPriceCol.�����ۼ�
'                    .EditCell
                ElseIf .Col = menuPriceCol.�ֳɱ��� And cboPriceMethod.Text = "�����ɱ���" Then
                    If .Row = .rows - 1 And Val(.TextMatrix(.Row, menuPriceCol.ҩƷid)) <> 0 Then
                        .rows = .rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.ҩƷ
                        .RowHeight(.rows - 1) = mlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.ҩƷid)) <> 0 Then
                        .ColComboList(menuPriceCol.ҩƷ) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.ҩƷ
                    End If
                ElseIf .Col = menuPriceCol.ҩƷ And cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    .Col = menuPriceCol.�ֳɱ���
'                    .EditCell
                ElseIf .Col = menuPriceCol.�ֳɱ��� And cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    .Col = menuPriceCol.�����ۼ�
'                    .EditCell
                ElseIf .Col = menuPriceCol.�����ۼ� And cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.ҩƷ
                        .RowHeight(.rows - 1) = mlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.ҩƷid)) <> 0 Then
                        .ColComboList(menuPriceCol.ҩƷ) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.ҩƷ
'                        .EditCell
                    End If
                Else
                    .Col = .Col + 1
'                    .EditCell
                End If
            Else
                If Val(.TextMatrix(.Row, menuPriceCol.ҩƷid)) <> 0 And .Row = .rows - 1 Then
                    .ColComboList(menuPriceCol.ҩƷ) = ""
                    .rows = .rows + 1
                    .Row = .Row + 1
                    .Col = menuPriceCol.ҩƷ
                    .RowHeight(.rows - 1) = mlngRowHeight
'                    .EditCell
                    Call setColEdit
                ElseIf Val(.TextMatrix(.Row, menuPriceCol.ҩƷid)) <> 0 Then
                    .ColComboList(menuPriceCol.ҩƷ) = ""
                    .Row = .Row + 1
                    .Col = menuPriceCol.ҩƷ
'                    .EditCell
                End If
            End If
        ElseIf KeyCode = vbKeyDelete Then
            lngDrugID = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.ҩƷid))
            
            '�޸�ģʽʱɾ��һ���۸�������ݣ������δִ�м۸�
            If mintModal = 1 Then
                'Private mint���� As Integer     '0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
                'ɾ����ʽ_In   In Number := 0 --0-����;1-�ۼ�;2-�ɱ���
                If mint���� = 0 Then
                    intɾ������ = 1
                ElseIf mint���� = 1 Then
                    intɾ������ = 2
                Else
                    intɾ������ = 0
                End If
                
                gstrSQL = "Zl_ҩƷδִ�м۸�_Delete(" & lngDrugID & "," & intɾ������ & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
            
            If .rows > 2 Then
                .RemoveItem .Row
            Else
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.Row, intCol) = ""
                Next
            End If

            With vsfStore
                If lngDrugID = 0 Then Exit Sub
                For intRow = .rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With

            With vsfPay
                If lngDrugID = 0 Then Exit Sub
                For intRow = .rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuPayCol.ҩƷid)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim mrsReturn As Recordset
    Dim rsTemp As Recordset
    Dim vRect As RECT
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim strkey As String
    Dim lngDrugID As Long
    Dim intCurrentPirce As Integer '�Ƿ���ʱ��

    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    mBlnClick = True
    vRect = zlControl.GetControlRect(vsfPrice.hWnd) '��ȡλ��
    dblLeft = vRect.Left + vsfPrice.CellLeft
    dblTop = vRect.Top + vsfPrice.CellTop + vsfPrice.CellHeight

    With vsfPrice
        strkey = .EditText
        Select Case Col
        Case menuPriceCol.ҩƷ
            If grsMaster.State = adStateClosed Then
                Call SetSelectorRS(1, "", 0, , , , , , , , , True)
            End If
            Set mrsReturn = frmSelector.ShowME(Me, 1, 1, strkey, dblLeft, dblTop, , , , , , , , , False, mstrPrivs)
            If mrsReturn.RecordCount = 0 Then Exit Sub
            mblnUpdateAdd = True
            Call GetDrugPirce(mrsReturn, Row)
            mblnUpdateAdd = False
        End Select
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckDoubleDrug(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '����Ƿ����ظ���ҩƷ
    'lngDrugId ҩƷid
    '����ֵ true-�����ظ�ֵ false-�������ظ�ֵ
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    Dim strName As String
    Dim intCount As Integer
    Dim intLength As Integer

    If rsTemp.RecordCount = 0 Then Exit Function
    rsTemp.MoveFirst
    With vsfPrice
        For i = 0 To rsTemp.RecordCount - 1
            For j = 1 To .rows - 1
                If Val(.TextMatrix(j, menuPriceCol.ҩƷid)) = rsTemp!ҩƷid Then
                    strTemp = strTemp & " ҩƷid <> " & rsTemp!ҩƷid & " and "
                    intCount = intCount + 1
                    If intCount < 5 Then
                        strName = strName & rsTemp!ͨ���� & " "
                    End If
                End If
            Next
            rsTemp.MoveNext
        Next
    End With

    If strTemp <> "" Then
        intLength = LenB(StrConv(strTemp, vbFromUnicode)) '�õ��ַ�������
        Do Until Mid(strTemp, intLength, 3) = "and" '�Ӻ���ǰ���ҵ�����һ��"and"
           intLength = intLength - 1
        Loop
        strTemp = Left(strTemp, intLength - 1) '������һ��"and"֮ǰ���ַ���

        rsTemp.Filter = strTemp
        MsgBox strName & "��" & intCount & "��ҩƷ���б����Ѿ����ڣ��Ѵ���ҩƷ������ӣ�", vbInformation, gstrSysName
    End If

    Set CheckDoubleDrug = rsTemp
End Function

Private Sub vsfPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPrice
            If .Col = menuPriceCol.ҩƷ Then
                .Editable = flexEDKbdMouse
                Exit Sub
            End If
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer

    With vsfPrice
        strkey = .EditText
        If .Col = menuPriceCol.�ֳɱ��� Then
            mdbl�ɱ��� = Val(.TextMatrix(Row, Col))
        End If
    End With

    If Col = menuPriceCol.�ֳɱ��� Or Col = menuPriceCol.�����ۼ� Then
        If KeyAscii = vbKeyReturn Then Exit Sub
        If KeyAscii <> vbKeyBack Then
            Select Case Col
                Case menuPriceCol.�ֳɱ���
                    intDigit = mintCostDigit
                Case menuPriceCol.�����ۼ�
                    intDigit = mintPriceDigit
            End Select

            If KeyAscii = vbKeyDelete Then
                If InStr(1, strkey, ".") > 0 Then
                    KeyAscii = 0
                End If
            ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                If vsfPrice.EditSelLength = Len(strkey) Then Exit Sub
                If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                    KeyAscii = 0
                    Exit Sub
                End If
                If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf Col = menuPriceCol.ҩƷ Then
        If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsfPrice_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = menuPriceCol.ҩƷ Then
        vsfPrice.ColComboList(menuPriceCol.ҩƷ) = "|..."
    End If
End Sub

Private Sub setColEdit()
    '���ܣ��������Ƿ�����޸�
    '�����޸ĵ�����ɫΪ��ɫ�����޸ĵ�����ɫΪ��ɫ
    Dim intCol As Integer
    Dim intRow As Integer

    With vsfPrice
        .Cell(flexcpBackColor, 1, 1, .rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "�����ۼ�" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.ҩƷ, .rows - 1, menuPriceCol.ҩƷ) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�����ۼ�, .rows - 1, menuPriceCol.�����ۼ�) = mconlngCanColColor
        ElseIf cboPriceMethod.Text = "�����ɱ���" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.ҩƷ, .rows - 1, menuPriceCol.ҩƷ) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�ֳɱ���, .rows - 1, menuPriceCol.�ֳɱ���) = mconlngCanColColor
        Else
            .Cell(flexcpBackColor, 1, menuPriceCol.ҩƷ, .rows - 1, menuPriceCol.ҩƷ) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�ֳɱ���, .rows - 1, menuPriceCol.�ֳɱ���) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�����ۼ�, .rows - 1, menuPriceCol.�����ۼ�) = mconlngCanColColor
        End If

    End With

    With vsfStore
        If .rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "�����ۼ�" Then
            .Cell(flexcpBackColor, 1, menuStoreCol.�����ۼ�, .rows - 1, menuStoreCol.�����ۼ�) = mconlngCanColColor
        ElseIf cboPriceMethod.Text = "�����ɱ���" Then
'            .Cell(flexcpBackColor, 1, menuStoreCol.�ӳ���, .rows - 1, menuStoreCol.�ӳ���) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.�ֳɱ���, .rows - 1, menuStoreCol.�ֳɱ���) = mconlngCanColColor
        Else
            .Cell(flexcpBackColor, 1, menuStoreCol.�ӳ���, .rows - 1, menuStoreCol.�ӳ���) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.�ֳɱ���, .rows - 1, menuStoreCol.�ֳɱ���) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.�����ۼ�, .rows - 1, menuStoreCol.�����ۼ�) = mconlngCanColColor
        End If
        If .rows > 1 Then
            For intRow = 1 To .rows - 1
                If Val(.TextMatrix(intRow, menuStoreCol.���)) = 1 And mblnʱ��ҩƷ�����ε��� = True And mint���� <> 1 Then
                    .Cell(flexcpBackColor, intRow, menuStoreCol.�����ۼ�, intRow, menuStoreCol.�����ۼ�) = mconlngCanColColor
                Else
                    .Cell(flexcpBackColor, intRow, menuStoreCol.�����ۼ�, intRow, menuStoreCol.�����ۼ�) = mconlngColor
                End If
                If mbln�ɱ��۰��ⷿ���ε��� = True And mint���� <> 0 Then
                    .Cell(flexcpBackColor, intRow, menuStoreCol.�ֳɱ���, intRow, menuStoreCol.�ֳɱ���) = mconlngCanColColor
                Else
                    .Cell(flexcpBackColor, intRow, menuStoreCol.�ֳɱ���, intRow, menuStoreCol.�ֳɱ���) = mconlngColor
                End If
            Next
        End If
    End With

    With vsfPay
        If .rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .rows - 1, .Cols - 1) = mconlngColor
        .Cell(flexcpBackColor, 1, menuPayCol.��Ʊ��, .rows - 1, menuPayCol.��Ʊ��) = mconlngCanColColor
        .Cell(flexcpBackColor, 1, menuPayCol.��Ʊ����, .rows - 1, menuPayCol.��Ʊ����) = mconlngCanColColor
        .Cell(flexcpBackColor, 1, menuPayCol.��Ʊ���, .rows - 1, menuPayCol.��Ʊ���) = mconlngCanColColor
    End With
End Sub


Private Sub vsfPrice_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        vsfPrice.Editable = flexEDNone
        If vsfPrice.Col = menuPriceCol.ҩƷ And mintModal <> 2 Then
            vsfPrice.ColComboList(menuPriceCol.ҩƷ) = "|..."
            vsfPrice.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub vsfPrice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngDrugID As Long
    Dim dblSalePrice As Double
    Dim intRow As Integer
    Dim dbl�ӳ��� As Double

    With vsfPrice
        If .EditText = "" Then Exit Sub
        lngDrugID = Val(.TextMatrix(Row, menuPriceCol.ҩƷid))
        If lngDrugID = 0 Then Exit Sub

        Select Case Col
            Case menuPriceCol.�ֳɱ���
                If Val(.EditText) < 0 Then
                    MsgBox "�ɱ��۲���Ϊ������", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If
                If .EditText > 9999999 Then
                    MsgBox "�ɱ��۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                .EditText = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
                If mbln�ּ���ʾ = True Then
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�)) Then
                        If MsgBox("�ֳɱ��۸��ڲɹ����޼�" & Val(.TextMatrix(.Row, menuPriceCol.ԭ�ɹ��޼�)) & "��" & vbCrLf & "������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            .TextMatrix(.Row, menuPriceCol.�ֲɹ��޼�) = zlStr.FormatEx(.EditText, mintCostDigit, , True)
                        End If
                    End If
                Else
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�)) Then
                        .TextMatrix(.Row, menuPriceCol.�ֲɹ��޼�) = zlStr.FormatEx(.EditText, mintCostDigit, , True)
                    End If
                End If

                If cbo�ۼۼ��㷽ʽ.Text = "�ۼ۰��ֶμӳɼ���" And .TextMatrix(.Row, menuPriceCol.�Ƿ���) = "1" And mint���� = 2 Then
                    Call get�ֶμӳ��ۼ�(lngDrugID, Val(.TextMatrix(.Row, menuPriceCol.��װϵ��)), Val(.EditText), dblSalePrice)
                    If dblSalePrice = 0 Then
                        .EditText = mdbl�ɱ���
                        .TextMatrix(vsfPrice.Row, menuPriceCol.�ֳɱ���) = zlStr.FormatEx(.EditText, mintCostDigit, , True)
                        Exit Sub
                    End If
                    dblSalePrice = dblSalePrice + (Val(.TextMatrix(.Row, menuPriceCol.ԭָ���ۼ�)) - dblSalePrice) * (1 - Val(.TextMatrix(.Row, menuPriceCol.���������)) / 100)
                    .TextMatrix(.Row, menuPriceCol.�����ۼ�) = zlStr.FormatEx(dblSalePrice, mintPriceDigit, , True)
                    
                    '�����ۼ�Ӧ��ͬ�����¿���б�۸���Ϣ
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.ҩƷid) = .TextMatrix(.Row, menuPriceCol.ҩƷid) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�) = zlStr.FormatEx(dblSalePrice, mintPriceDigit, , True)
'                                vsfStore.TextMatrix(intRow, menuStoreCol.�ۼ�ӯ��) = Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)) * (Val(vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ۼ�ӯ��) = Format(Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�)), mstrMoneyFormat) - Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�)), mstrMoneyFormat), mstrMoneyFormat)
                                
                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) <> 0 Then
                                    dbl�ӳ��� = zlStr.FormatEx(zlStr.FormatEx(((Val(vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�))) / Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) - 1), 5, , True) * 100, 5, , True)
                                Else
                                    dbl�ӳ��� = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ӳ���) = dbl�ӳ���
                            End If
                        Next
                    End If
                ElseIf cbo�ۼۼ��㷽ʽ = "�ۼ۰��̶���������" And .TextMatrix(.Row, menuPriceCol.�Ƿ���) = "1" And mint���� = 2 Then
                    dblSalePrice = Val(.EditText) * (1 + Val(.TextMatrix(.Row, menuPriceCol.�ӳ���)))
                    If dblSalePrice > Val(.TextMatrix(.Row, menuPriceCol.ԭָ���ۼ�)) Then dblSalePrice = Val(.TextMatrix(.Row, menuPriceCol.ԭָ���ۼ�))
                    .TextMatrix(.Row, menuPriceCol.�����ۼ�) = zlStr.FormatEx(dblSalePrice, mintPriceDigit, , True)
                    
                    '�����ۼ�Ӧ��ͬ�����¿���б�۸���Ϣ
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.ҩƷid) = .TextMatrix(.Row, menuPriceCol.ҩƷid) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�) = zlStr.FormatEx(dblSalePrice, mintPriceDigit, , True)
'                                vsfStore.TextMatrix(intRow, menuStoreCol.�ۼ�ӯ��) = Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)) * (Val(vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ۼ�ӯ��) = Format(Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�)), mstrMoneyFormat) - Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�)), mstrMoneyFormat), mstrMoneyFormat)

                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) <> 0 Then
                                    dbl�ӳ��� = zlStr.FormatEx(zlStr.FormatEx(((Val(vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�))) / Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) - 1), 5, , True) * 100, 5, , True)
                                Else
                                    dbl�ӳ��� = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ӳ���) = dbl�ӳ���
                            End If
                        Next
                    End If
                End If

                Call CaculateCost(lngDrugID, .EditText) '���¼���ɱ���
            Case menuPriceCol.�����ۼ�
                If Val(.EditText) < 0 Then
                    MsgBox "�ۼ۲���Ϊ������", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If

                If .EditText > 9999999 Then
                    MsgBox "���ۼ۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If

                .EditText = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
'                If mdblOldPrice = .EditText Then 'δ���޸�ֱ���˳�
'                    Exit Sub
'                End If

                If mbln�ּ���ʾ = True Then
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.ԭָ���ۼ�)) Then
                        If MsgBox("�����ۼ۸���ָ���ۼ�" & Val(.TextMatrix(.Row, menuPriceCol.ԭָ���ۼ�)) & "��" & vbCrLf & "������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            .TextMatrix(.Row, menuPriceCol.��ָ���ۼ�) = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
                        End If
                    End If
                Else
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.ԭָ���ۼ�)) Then
                        .TextMatrix(.Row, menuPriceCol.��ָ���ۼ�) = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
                    End If
                End If
                
                If chkAotuCost.Value = 1 Then '�޸��ۼۺ��Զ�����ɱ���
                    .TextMatrix(.Row, menuPriceCol.�ֳɱ���) = zlStr.FormatEx(.EditText / (1 + Val(.TextMatrix(.Row, menuPriceCol.�ӳ���))), mintCostDigit, , True)
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.ҩƷid) = .TextMatrix(.Row, menuPriceCol.ҩƷid) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(.TextMatrix(.Row, menuPriceCol.�ֳɱ���), mintCostDigit, , True)
                                
                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) <> 0 Then
                                    dbl�ӳ��� = zlStr.FormatEx((.EditText / Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) - 1), 5, , True)
                                Else
                                    dbl�ӳ��� = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ӳ���) = zlStr.FormatEx(dbl�ӳ��� * 100, 5, , True)
'                                vsfStore.TextMatrix(intRow, menuStoreCol.�ɱ�ӯ��) = Format((Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.ԭ�ɱ���))) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat)
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ɱ�ӯ��) = Format(Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat) - Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.ԭ�ɱ���)) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat), mstrMoneyFormat)
                           
                            End If
                        Next
                    End If
                End If

                Call ChangeDrugStore(Row, lngDrugID, .EditText)
        End Select
    End With
End Sub

Private Sub ChangeDrugStore(ByVal intRow As Integer, ByVal lngDrugID As Long, ByVal dblNewPrice As Double)
    '���ܣ�ͨ���޸ļ۸���е����ۼ��޸Ŀ���б������Ӧ�����ۼ�
    Dim dblOldPrice As Double
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim dblNum As Double
    Dim dbl��װ As Double
    Dim n As Integer
    Dim dbl��Ʊ��� As Double
    Dim dbl�ӳ��� As Double

    If intRow = 0 Or mint���� = 1 Then Exit Sub

    dbl��װ = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.��װϵ��))

    With vsfStore
        For n = 1 To .rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If Val(.TextMatrix(n, menuStoreCol.ҩƷid)) = lngDrugID Then
                    dblNum = Val(.TextMatrix(n, menuStoreCol.����))
                    dblOldPrice = Val(vsfStore.TextMatrix(n, menuStoreCol.ԭ���ۼ�))

                    .TextMatrix(n, menuStoreCol.�����ۼ�) = zlStr.FormatEx(dblNewPrice, mintPriceDigit, , True)
'                    .TextMatrix(n, menuStoreCol.�ۼ�ӯ��) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (dblNewPrice - dblOldPrice), mstrMoneyFormat)
                    .TextMatrix(n, menuStoreCol.�ۼ�ӯ��) = Format(Format(Val(.TextMatrix(n, menuStoreCol.����)) * dblNewPrice, mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.����)) * dblOldPrice, mstrMoneyFormat), mstrMoneyFormat)

                    If Val(.TextMatrix(n, menuStoreCol.�ֳɱ���)) <> 0 Then
                        dbl�ӳ��� = zlStr.FormatEx(((Val(.TextMatrix(n, menuStoreCol.�����ۼ�))) / Val(.TextMatrix(n, menuStoreCol.�ֳɱ���)) - 1), 5, , True)
                    Else
                        dbl�ӳ��� = 0
                    End If
                    .TextMatrix(n, menuStoreCol.�ӳ���) = zlStr.FormatEx(dbl�ӳ��� * 100, 5, , True)
                    
                    If mint���� = 2 And chkAotuCost.Value = 1 Then
                        dblOldCost = .TextMatrix(n, menuStoreCol.ԭ�ɱ���)
                        dblNewCost = dblNewPrice / (1 + Round(Val(.TextMatrix(n, menuStoreCol.�ӳ���)) / 100, 7))
                        .TextMatrix(n, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(dblNewCost, mintCostDigit, , True)
'                        .TextMatrix(n, menuStoreCol.�ɱ�ӯ��) = Format((dblNewCost - dblOldCost) * dblNum, mstrMoneyFormat)
                        .TextMatrix(n, menuStoreCol.�ɱ�ӯ��) = Format(Format(.TextMatrix(n, menuStoreCol.�ֳɱ���) * dblNum, mstrMoneyFormat) - Format(dblOldCost * dblNum, mstrMoneyFormat), mstrMoneyFormat)

                    End If
                    dbl��Ʊ��� = dbl��Ʊ��� + Val(.TextMatrix(n, menuStoreCol.�ɱ�ӯ��))
                End If
            End If
        Next
    End With

    If chkAutoPay.Value = 1 Then
        With vsfPay
            For n = 1 To .rows - 1
                If .TextMatrix(1, 0) <> "" Then
                    If Val(.TextMatrix(n, menuPayCol.ҩƷid)) = lngDrugID Then
                        .TextMatrix(n, menuPayCol.��Ʊ���) = zlStr.FormatEx(dbl��Ʊ���, mintMoneyDigit, , True)
                    End If
                End If
            Next
        End With
    End If

    If mint���� = 2 Then
        CaluateAverCost lngDrugID
    End If
End Sub

Private Sub CaluateAverCost(ByVal lngҩƷid As Long)
    '����ƽ���ɱ���
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.ҩƷid) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = lngҩƷid Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.�ֳɱ���)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.ҩƷid) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.ҩƷid)) = lngҩƷid Then
                        .TextMatrix(i, menuPriceCol.�ֳɱ���) = zlStr.FormatEx(dblSumCost / dblSumNumber, mintCostDigit, , True)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaluateAverOldCost(ByVal lngҩƷid As Long)
    '����ԭʼƽ���ɱ���
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.ҩƷid) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = lngҩƷid Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.ԭ�ɱ���)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.ҩƷid) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.ҩƷid)) = lngҩƷid Then
                        .TextMatrix(i, menuPriceCol.ԭ�ɱ���) = zlStr.FormatEx(dblSumCost / dblSumNumber, mintCostDigit, , True)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaculateCost(ByVal lngҩƷid As Long, ByVal dbl�ֳɱ��� As Double)
    '���ܣ�ͨ���޸ļ۸���еĳɱ����޸Ŀ���б������Ӧ�ĳɱ���

    Dim n As Integer
    Dim dbl��Ʊ��� As Double

    With vsfStore
        For n = 1 To .rows - 1
            If .TextMatrix(n, menuStoreCol.ҩƷid) <> "" Then
                If Val(.TextMatrix(n, menuStoreCol.ҩƷid)) = lngҩƷid Then
                    .TextMatrix(n, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(dbl�ֳɱ���, mintCostDigit, , True)
                    If (cbo�ۼۼ��㷽ʽ.Text = "�ۼ۰��ֶμӳɼ���" Or cbo�ۼۼ��㷽ʽ.Text = "�ۼ۰��̶���������") And vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.�Ƿ���) = "1" And mint���� = 2 Then
                        .TextMatrix(n, menuStoreCol.�����ۼ�) = vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.�����ۼ�)
                    End If
                    If dbl�ֳɱ��� <> 0 Then
                        .TextMatrix(n, menuStoreCol.�ӳ���) = zlStr.FormatEx(zlStr.FormatEx((Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) / dbl�ֳɱ��� - 1), 5, , True) * 100, 5, , True)
                    End If
                    If cbo�ۼۼ��㷽ʽ = "�ۼ۰��ֶμӳɼ���" Then
                        .TextMatrix(n, menuStoreCol.�ӳ���) = zlStr.FormatEx(zlStr.FormatEx(mdbl�ֶμӳ���, 5, , True) * 100, 5, , True)
                    End If
'                    .TextMatrix(n, menuStoreCol.�ɱ�ӯ��) = Format((dbl�ֳɱ��� - Val(.TextMatrix(n, menuStoreCol.ԭ�ɱ���))) * Val(.TextMatrix(n, menuStoreCol.����)), mstrMoneyFormat)
                    .TextMatrix(n, menuStoreCol.�ɱ�ӯ��) = Format(Format(dbl�ֳɱ��� * Val(.TextMatrix(n, menuStoreCol.����)), mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.ԭ�ɱ���)) * Val(.TextMatrix(n, menuStoreCol.����)), mstrMoneyFormat), mstrMoneyFormat)
                    dbl��Ʊ��� = dbl��Ʊ��� + (dbl�ֳɱ��� - .TextMatrix(n, menuStoreCol.ԭ�ɱ���)) * Val(.TextMatrix(n, menuStoreCol.����))
                    .TextMatrix(n, menuStoreCol.�ۼ�ӯ��) = Format(Format(Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) * Val(.TextMatrix(n, menuStoreCol.����)), mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�)) * Val(.TextMatrix(n, menuStoreCol.����)), mstrMoneyFormat), mstrMoneyFormat)
                
                End If
            End If
        Next
    End With

    If chkAutoPay.Value = 1 Then
        For n = 1 To vsfPay.rows - 1
            If vsfPay.TextMatrix(1, 0) <> "" Then
                If Val(vsfPay.TextMatrix(n, menuPayCol.ҩƷid)) = lngҩƷid Then
                    vsfPay.TextMatrix(n, menuPayCol.��Ʊ���) = Format(dbl��Ʊ���, mstrMoneyFormat)
                End If
            End If
        Next
    End If
End Sub


Private Sub vsfStore_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfStore
        .Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height - 370
    End With
End Sub

Private Sub vsfStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfStore
        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
            Cancel = True
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub setColHiddenVsf()
    '��ͬģʽ���棬����ʾ��һ��
    With vsfStore
        If cboPriceMethod.Text = "�����ۼ�" Then
            .colHidden(menuStoreCol.����) = True
            .colHidden(menuStoreCol.���) = True
            .colHidden(menuStoreCol.�ӳ���) = True
            .colHidden(menuStoreCol.ԭ�ɱ���) = True
            .colHidden(menuStoreCol.�ֳɱ���) = False
            .colHidden(menuStoreCol.�ɱ�ӯ��) = True
            .colHidden(menuStoreCol.ԭ���ۼ�) = False
            .colHidden(menuStoreCol.�����ۼ�) = False
        ElseIf cboPriceMethod.Text = "�����ɱ���" Then
            .colHidden(menuStoreCol.ԭ���ۼ�) = True
            .colHidden(menuStoreCol.�����ۼ�) = False
            .colHidden(menuStoreCol.�ۼ�ӯ��) = True
            .colHidden(menuStoreCol.�ӳ���) = False
            .colHidden(menuStoreCol.ԭ�ɱ���) = False
            .colHidden(menuStoreCol.�ֳɱ���) = False
            .colHidden(menuStoreCol.�ɱ�ӯ��) = False
        ElseIf cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
            .colHidden(menuStoreCol.ԭ���ۼ�) = False
            .colHidden(menuStoreCol.�����ۼ�) = False
            .colHidden(menuStoreCol.�ۼ�ӯ��) = False
            .colHidden(menuStoreCol.�ӳ���) = False
            .colHidden(menuStoreCol.ԭ�ɱ���) = False
            .colHidden(menuStoreCol.�ֳɱ���) = False
            .colHidden(menuStoreCol.�ɱ�ӯ��) = False
        End If
    End With
End Sub

Private Sub vsfStore_Click()
    Dim i As Integer
    With vsfStore
        For i = 1 To vsfPrice.rows - 1
            If Val(.TextMatrix(.Row, menuStoreCol.ҩƷid)) = Val(vsfPrice.TextMatrix(i, menuPriceCol.ҩƷid)) Then
                vsfPrice.Tag = i
            End If
        Next
    End With
End Sub

Private Sub vsfStore_DblClick()
    With vsfStore
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfStore_EnterCell()
    With vsfStore
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
        If .Col = menuStoreCol.�ӳ��� Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.�ӳ���))
        ElseIf .Col = menuStoreCol.�ֳɱ��� Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.�ֳɱ���))
        ElseIf .Col = menuStoreCol.�����ۼ� Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.�����ۼ�))
        End If
    End With
End Sub

Private Sub vsfStore_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfStore
        If KeyCode = vbKeyReturn Then
            If .Col < vsfStore.Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row <> .rows - 1 Then
                    .Row = .Row + 1
                    .Col = menuStoreCol.���
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfStore_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfStore
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfStore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer

    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfStore
            If Col = menuStoreCol.�ֳɱ��� Or Col = menuStoreCol.�����ۼ� Or Col = menuStoreCol.�ӳ��� Then
                strkey = .EditText
                Select Case Col
                    Case menuStoreCol.�ֳɱ���
                        intDigit = mintCostDigit
                    Case menuStoreCol.�����ۼ�
                        intDigit = mintPriceDigit
                    Case menuStoreCol.�ӳ���
                        intDigit = 5
                End Select
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Else
                    KeyAscii = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfStore_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strInput As String
    Dim n As Integer
    Dim intRow As Integer
    Dim dbl��Ʊ��� As Double
    Dim Dbl���� As Double
    Dim Dbl��� As Double
    Dim Dbl�ɱ���� As Double
    Dim dbl�ֲɹ��� As Double
    Dim dblTempNum As Double

    With vsfStore
        If .EditText = "" Then Exit Sub
        intRow = .Row
        Select Case .Col
            Case menuStoreCol.�����ۼ�
                If Not IsNumeric(.EditText) Then
                    MsgBox "�������µ��ۼۡ�", vbInformation, gstrSysName
                    Exit Sub
                Else
                    .EditText = zlStr.FormatEx(.EditText, mintPriceDigit, , True)
                End If

                If .EditText > 9999999 Then
                    MsgBox "���ۼ۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If

'                If mdblOldPrice = .EditText Then Exit Sub

                If chkAotuCost.Value = 1 Then '�޸��ۼۺ��Զ�����ɱ���
                    .TextMatrix(intRow, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(.EditText / (1 + Val(.TextMatrix(intRow, menuStoreCol.�ӳ���)) / 100), mintCostDigit, , True)
                    .TextMatrix(intRow, menuStoreCol.�ɱ�ӯ��) = Format(Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * Val(.TextMatrix(intRow, menuStoreCol.�ֳɱ���)), mstrMoneyFormat) - Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * Val(.TextMatrix(intRow, menuStoreCol.ԭ�ɱ���)), mstrMoneyFormat), mstrMoneyFormat)
                End If
                
                .TextMatrix(intRow, menuStoreCol.�ۼ�ӯ��) = Format(Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * Val(.EditText), mstrMoneyFormat) - Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�)), mstrMoneyFormat), mstrMoneyFormat)

'                .TextMatrix(intRow, menuStoreCol.�ۼ�ӯ��) = Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * (Val(.EditText) - Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                .TextMatrix(intRow, menuStoreCol.�����ۼ�) = zlStr.FormatEx(Val(.EditText), mintPriceDigit, , True)
'                .TextMatrix(intRow, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) / (1 + Val(.TextMatrix(intRow, menuStoreCol.�ӳ���)) / 100), mintCostDigit)
'                .TextMatrix(intRow, menuStoreCol.�ɱ�ӯ��) = Format((Val(.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) - Val(.TextMatrix(intRow, menuStoreCol.ԭ�ɱ���))) * Val(.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat)
                If chkAotuCost.Value <> 1 Then
                    If Val(.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) <> 0 Then
                        .TextMatrix(intRow, menuStoreCol.�ӳ���) = zlStr.FormatEx(zlStr.FormatEx((Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) / Val(.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) - 1), 5, , True) * 100, 5, , True)
                    Else
                        .TextMatrix(intRow, menuStoreCol.�ӳ���) = zlStr.FormatEx(0, 5, , True)
                    End If
                End If
                
                For n = 1 To .rows - 1
                    If .TextMatrix(intRow, menuStoreCol.ҩƷid) = .TextMatrix(n, menuStoreCol.ҩƷid) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.����)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.����)) = Val(.TextMatrix(n, menuStoreCol.����)) Then
                            .TextMatrix(n, menuStoreCol.�����ۼ�) = .TextMatrix(intRow, menuStoreCol.�����ۼ�)
'                            .TextMatrix(n, menuStoreCol.�ۼ�ӯ��) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (Val(.EditText) - Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                            .TextMatrix(n, menuStoreCol.�ۼ�ӯ��) = Format(Format(Val(.TextMatrix(n, menuStoreCol.����)) * Val(.EditText), mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.����)) * Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�)), mstrMoneyFormat), mstrMoneyFormat)
                            If chkAotuCost.Value <> 1 Then
                                If Val(.TextMatrix(n, menuStoreCol.�ֳɱ���)) <> 0 Then
                                    .TextMatrix(n, menuStoreCol.�ӳ���) = zlStr.FormatEx(zlStr.FormatEx((Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) / Val(.TextMatrix(n, menuStoreCol.�ֳɱ���)) - 1), 5, , True) * 100, 5, , True)
                                Else
                                    .TextMatrix(n, menuStoreCol.�ӳ���) = zlStr.FormatEx(0, 5, , True)
                                End If
                            End If
                        End If
                        Dbl���� = Dbl���� + .TextMatrix(n, menuStoreCol.����)
                        Dbl��� = Dbl��� + .TextMatrix(n, menuStoreCol.����) * Val(.TextMatrix(n, menuStoreCol.�����ۼ�))
                        Dbl�ɱ���� = Dbl�ɱ���� + .TextMatrix(n, menuStoreCol.����) * Val(.TextMatrix(n, menuStoreCol.�ֳɱ���))
                    End If
                Next
                For n = 1 To vsfPrice.rows - 1
                    If .TextMatrix(intRow, menuStoreCol.ҩƷid) = vsfPrice.TextMatrix(n, menuPriceCol.ҩƷid) Then
                        If Dbl���� <> 0 Then
                            If chkAotuCost.Value = 1 Then
                                vsfPrice.TextMatrix(n, menuPriceCol.�ֳɱ���) = zlStr.FormatEx(Dbl�ɱ���� / Dbl����, mintPriceDigit, , True)
                            End If
                            vsfPrice.TextMatrix(n, menuPriceCol.�����ۼ�) = zlStr.FormatEx(Dbl��� / Dbl����, mintPriceDigit, , True)
                        Else
                            If chkAotuCost.Value = 1 Then
                                vsfPrice.TextMatrix(n, menuPriceCol.�ֳɱ���) = vsfStore.TextMatrix(intRow, menuStoreCol.�ֳɱ���)
                            End If
                            vsfPrice.TextMatrix(n, menuPriceCol.�����ۼ�) = vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�)
                        End If
                    End If
                Next

                If mint���� > 0 Then
                    For n = 1 To .rows - 1
                        If .TextMatrix(n, menuStoreCol.ҩƷid) <> "" Then
                            If Val(.TextMatrix(n, menuStoreCol.ҩƷid)) = Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)) Then
                                dbl��Ʊ��� = dbl��Ʊ��� + (Val(.TextMatrix(n, menuStoreCol.�ֳɱ���)) - Val(.TextMatrix(n, menuStoreCol.ԭ�ɱ���))) * Val(.TextMatrix(n, menuStoreCol.����))
                            End If
                        End If
                    Next

                    If chkAutoPay.Value = 1 Then
                        For n = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(1, 0) <> "" Then
                                If Val(vsfPay.TextMatrix(n, menuPayCol.ҩƷid)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.ҩƷid)) Then
                                    vsfPay.TextMatrix(n, menuPayCol.��Ʊ���) = zlStr.FormatEx(dbl��Ʊ���, mintMoneyDigit, , True)
                                End If
                            End If
                        Next
                    End If
                End If
            Case menuStoreCol.�ӳ���
                If Val(.EditText) < 0 Then Exit Sub
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If
'                If mdblOldPrice = .EditText Then Exit Sub
                .EditText = zlStr.FormatEx(.EditText, 5, , True)
                .TextMatrix(intRow, menuStoreCol.�ӳ���) = zlStr.FormatEx(Val(.EditText), 5, , True)
                .TextMatrix(intRow, menuStoreCol.�����ۼ�) = zlStr.FormatEx(Val(.TextMatrix(intRow, menuStoreCol.�ֳɱ���)) * (1 + Val(.TextMatrix(intRow, menuStoreCol.�ӳ���)) / 100), mintCostDigit, , True)
'                .TextMatrix(intRow, menuStoreCol.�ۼ�ӯ��) = Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * (Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                .TextMatrix(intRow, menuStoreCol.�ۼ�ӯ��) = Format(Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)), mstrMoneyFormat) - Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�)), mstrMoneyFormat), mstrMoneyFormat)

                For n = 1 To .rows - 1
                    If vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.ҩƷid) = .TextMatrix(n, menuStoreCol.ҩƷid) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.���)) = 0 Or mblnʱ��ҩƷ�����ε��� = False Then
                            .TextMatrix(n, menuStoreCol.�ӳ���) = zlStr.FormatEx(Val(.EditText), 5, , True)
                            .TextMatrix(n, menuStoreCol.�����ۼ�) = zlStr.FormatEx(Val(.TextMatrix(n, menuStoreCol.�ֳɱ���)) * (1 + zlStr.FormatEx(Val(.EditText), 5) / 100), mintCostDigit, , True)
'                            .TextMatrix(n, menuStoreCol.�ۼ�ӯ��) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                            .TextMatrix(n, menuStoreCol.�ۼ�ӯ��) = Format(Format(Val(.TextMatrix(n, menuStoreCol.����)) * Val(.TextMatrix(n, menuStoreCol.�����ۼ�)), mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.����)) * Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�)), mstrMoneyFormat), mstrMoneyFormat)
    
                        End If
                        Dbl���� = Dbl���� + .TextMatrix(n, menuStoreCol.����)
                        Dbl��� = Dbl��� + .TextMatrix(n, menuStoreCol.����) * Val(.TextMatrix(n, menuStoreCol.�����ۼ�))
                    End If
                Next
                If Dbl���� <> 0 Then
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.�����ۼ�) = zlStr.FormatEx(Dbl��� / Dbl����, mintPriceDigit, , True)
                Else
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.�����ۼ�) = .TextMatrix(intRow, menuStoreCol.�����ۼ�)
                End If
            Case menuStoreCol.�ֳɱ���
                If Val(.EditText) > Val(.TextMatrix(.Row, menuStoreCol.�����ۼ�)) Then
                    MsgBox "ע�⣬�³ɱ��۴��������ۼۣ�", vbExclamation, gstrSysName
                End If

                If Val(.EditText) < 0 Then
                    MsgBox "�ɱ��۲���Ϊ������", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If .EditText > 9999999 Then
                    MsgBox "�ɹ��۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
'                If mdblOldPrice = .EditText Then Exit Sub
                .EditText = zlStr.FormatEx(.EditText, mintCostDigit, , True)
                .TextMatrix(intRow, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(Val(.EditText), mintCostDigit, , True)
'                If Val(.EditText) <> 0 Then
'                    .TextMatrix(intRow, menuStoreCol.�ӳ���) = zlStr.FormatEx((Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) / Val(.EditText) - 1) * 100, 5)
'                End If
'                .TextMatrix(intRow, menuStoreCol.�ɱ�ӯ��) = Format((Val(.EditText) - .TextMatrix(intRow, menuStoreCol.ԭ�ɱ���)) * Val(.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat)
                .TextMatrix(intRow, menuStoreCol.�ɱ�ӯ��) = Format(Format(Val(.EditText) * Val(.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat) - Format(.TextMatrix(intRow, menuStoreCol.ԭ�ɱ���) * Val(.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat), mstrMoneyFormat)
                
                If Val(.TextMatrix(intRow, menuStoreCol.���)) = 1 And mblnʱ��ҩƷ�����ε��� = True And mint���� <> 1 Then
                    .TextMatrix(intRow, menuStoreCol.�����ۼ�) = zlStr.FormatEx(zlStr.FormatEx(Val(.EditText), mintCostDigit) * (1 + (Val(.TextMatrix(intRow, menuStoreCol.�ӳ���)) / 100)), mintPriceDigit, , True)
'                    .TextMatrix(intRow, menuStoreCol.�ۼ�ӯ��) = Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * (Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                    .TextMatrix(intRow, menuStoreCol.�ۼ�ӯ��) = Format(Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)), mstrMoneyFormat) - Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�)), mstrMoneyFormat), mstrMoneyFormat)

                End If
                
                dbl��Ʊ��� = (Val(.EditText) - .TextMatrix(intRow, menuStoreCol.ԭ�ɱ���)) * Val(.TextMatrix(intRow, menuStoreCol.����))

                For n = 1 To .rows - 1
                    If .TextMatrix(n, menuStoreCol.ҩƷid) <> "" Then
                        If Val(.TextMatrix(n, menuStoreCol.ҩƷid)) = Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)) And n <> intRow Then
                            If mbln�ɱ��۰��ⷿ���ε��� = False Or (Val(.TextMatrix(intRow, menuStoreCol.����)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.����)) = Val(.TextMatrix(n, menuStoreCol.����))) Then
                                dbl�ֲɹ��� = Val(.EditText)
                                .TextMatrix(n, menuStoreCol.�ֳɱ���) = zlStr.FormatEx(dbl�ֲɹ���, mintCostDigit, , True)
'                                If dbl�ֲɹ��� <> 0 Then
'                                    .TextMatrix(n, menuStoreCol.�ӳ���) = zlStr.FormatEx((Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) / dbl�ֲɹ��� - 1) * 100, 5)
'                                End If
'                                .TextMatrix(n, menuStoreCol.�ɱ�ӯ��) = Format((dbl�ֲɹ��� - .TextMatrix(n, menuStoreCol.ԭ�ɱ���)) * Val(.TextMatrix(n, menuStoreCol.����)), mstrMoneyFormat)
                                 .TextMatrix(n, menuStoreCol.�ɱ�ӯ��) = Format(Format(dbl�ֲɹ��� * Val(.TextMatrix(n, menuStoreCol.����)), mstrMoneyFormat) - Format(.TextMatrix(n, menuStoreCol.ԭ�ɱ���) * Val(.TextMatrix(n, menuStoreCol.����)), mstrMoneyFormat), mstrMoneyFormat)
                               
                                If Val(.TextMatrix(intRow, menuStoreCol.���)) = 1 And mblnʱ��ҩƷ�����ε��� = True And mint���� <> 1 Then
                                    .TextMatrix(n, menuStoreCol.�����ۼ�) = zlStr.FormatEx(zlStr.FormatEx(dbl�ֲɹ���, mintCostDigit) * (1 + (Val(.TextMatrix(n, menuStoreCol.�ӳ���)) / 100)), mintPriceDigit, , True)
'                                    .TextMatrix(n, menuStoreCol.�ۼ�ӯ��) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                                    .TextMatrix(n, menuStoreCol.�ۼ�ӯ��) = Format(Format(Val(.TextMatrix(n, menuStoreCol.����)) * Val(.TextMatrix(n, menuStoreCol.�����ۼ�)), mstrMoneyFormat) - Format(Val(.TextMatrix(n, menuStoreCol.����)) * Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�)), mstrMoneyFormat), mstrMoneyFormat)

                                End If
                            Else
                                dbl�ֲɹ��� = Val(.TextMatrix(n, menuStoreCol.�ֳɱ���))
                            End If
                            dbl��Ʊ��� = dbl��Ʊ��� + (dbl�ֲɹ��� - .TextMatrix(n, menuStoreCol.ԭ�ɱ���)) * Val(.TextMatrix(n, menuStoreCol.����))
                        End If
                    End If
                Next

                If chkAutoPay.Value = 1 Then
                    For n = 1 To vsfPay.rows - 1
                        If vsfPay.TextMatrix(1, 0) <> "" Then
                            If Val(vsfPay.TextMatrix(n, menuPayCol.ҩƷid)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.ҩƷid)) Then
                                vsfPay.TextMatrix(n, menuPayCol.��Ʊ���) = Format(dbl��Ʊ���, mstrMoneyFormat)
                            End If
                        End If
                    Next
                End If

                If mbln�ɱ��۰��ⷿ���ε��� = False Then
                    For n = 1 To vsfPrice.rows - 1
                        If Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)) = Val(vsfPrice.TextMatrix(n, menuPriceCol.ҩƷid)) Then
                            vsfPrice.TextMatrix(n, menuPriceCol.�ֳɱ���) = .TextMatrix(intRow, menuStoreCol.�ֳɱ���)
                            Exit For
                        End If
                    Next
                Else
                    CaluateAverCost Val(.TextMatrix(intRow, menuStoreCol.ҩƷid))
                End If
                Call CaculateAverPirce(Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)))  '�۸�䶯������ƽ���ۼ�
        End Select
    End With
End Sub

Private Sub CaculateAverPirce(ByVal lngҩƷid As Long)
    '�Զ�����ƽ���ۼ�
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.ҩƷid) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = lngҩƷid Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.�����ۼ�)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.ҩƷid) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.ҩƷid)) = lngҩƷid Then
                        .TextMatrix(i, menuPriceCol.�����ۼ�) = zlStr.FormatEx(dblSumPrice / dblSumNumber, mintPriceDigit, , True)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaculateAverOldPirce(ByVal lngҩƷid As Long)
    '�Զ�ԭʼ����ƽ���ۼ�
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.ҩƷid) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = lngҩƷid Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.ԭ���ۼ�)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.ҩƷid) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.ҩƷid)) = lngҩƷid Then
                        .TextMatrix(i, menuPriceCol.ԭ���ۼ�) = zlStr.FormatEx(dblSumPrice / dblSumNumber, mintPriceDigit, , True)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub initCommandBars()
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    Dim cbrControlPopu As CommandBarControl
    Dim lngCount As Integer
    
    With CommandBarsGlobalSettings
        .App = App
        .CompanyName = "����������Ϣ��ҵ�������ι�˾" '��˾����
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '��������������Դ�ļ�
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '�ؼ��������ɫ����
    End With

    With cbsMain.Options
        .ShowExpandButtonAlways = False '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False '�����õĲ˵���������
        .UseFadedIcons = True 'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24 '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16 '����Сͼ��ĳߴ�
    End With

    With cbsMain
        .VisualTheme = xtpThemeOffice2003 '���ÿؼ���ʾ���
        .EnableCustomization False '�Ƿ������Զ�������
        Set .Icons = imgList.Icons '���ù�����ͼ��ؼ�
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap '����仯ʱ�������ʾ����˵�Ҳ������
        .ActiveMenuBar.Title = "�˵�"
    End With
    
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.count To 1 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '����������
    Set cbrToolBar = cbsMain.Add("������", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    cbrToolBar.ContextMenuPresent = False

    With cbrToolBar
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_PrintStore, "��ӡ���䶯��")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_ClearAll, "���")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Find, "����")
        cbrControl.Visible = False
    
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_BatchSelect, "����ѡ����Ŀ")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Save, "ȷ��")
        cbrControl.BeginGroup = True
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Quit, "�˳�")
                
    End With

    For Each cbrControl In cbrToolBar.Controls  '�ù������а�ťͬʱ��ʾͼ�������
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F3, mconMenu_Find
    End With

End Sub
