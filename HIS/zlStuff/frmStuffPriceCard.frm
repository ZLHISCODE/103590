VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmStuffPriceCard 
   Caption         =   "���ĵ��۵�"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12810
   Icon            =   "frmStuffPriceCard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   2775
      TabIndex        =   45
      Top             =   4200
      Width           =   2775
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   840
      TabIndex        =   41
      Top             =   7440
      Width           =   1965
   End
   Begin VB.PictureBox picOtherSelect 
      Height          =   3135
      Left            =   3600
      ScaleHeight     =   3075
      ScaleWidth      =   4755
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdFilterOk 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   2400
         TabIndex        =   38
         Top             =   2640
         Width           =   1100
      End
      Begin VB.CommandButton cmdFilterCan 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   3480
         TabIndex        =   37
         Top             =   2640
         Width           =   1100
      End
      Begin VB.Frame fra����ѡ�� 
         Caption         =   "����ѡ��ɱ��۵�����أ�"
         Height          =   2535
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   4695
         Begin VB.CheckBox chk�ӳ��� 
            Caption         =   "ָ���ӳ���"
            Height          =   180
            Left            =   120
            TabIndex        =   32
            Top             =   1125
            Width           =   1215
         End
         Begin VB.CheckBox chk��Ӧ�� 
            Caption         =   "ָ����Ӧ��"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkӦ����¼ 
            Caption         =   "�����ɱ��۵��۴�����Ӧ����������¼"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txt�ӳ��� 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   29
            Text            =   "15.0000"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txt��Ӧ�� 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   28
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmd��Ӧ�� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   270
            Left            =   4080
            TabIndex        =   27
            Top             =   350
            Width           =   375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
            Height          =   1695
            Left            =   120
            TabIndex        =   33
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
            TabIndex        =   36
            Top             =   1440
            Width           =   4260
         End
         Begin VB.Label lblComment��Ӧ�� 
            AutoSize        =   -1  'True
            Caption         =   "��ָ����Ӧ�̣���ֻ�����ù�Ӧ�̵Ŀ�����ĳɱ��ۣ�"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   240
            TabIndex        =   35
            Top             =   720
            Width           =   4320
         End
         Begin VB.Label lblPercent 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   180
            Left            =   2415
            TabIndex        =   34
            Top             =   1125
            Width           =   90
         End
      End
   End
   Begin VB.PictureBox picInfo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   -120
      ScaleHeight     =   495
      ScaleWidth      =   10575
      TabIndex        =   20
      Top             =   6600
      Width           =   10575
      Begin VB.TextBox txtSummary 
         Height          =   300
         Left            =   4320
         MaxLength       =   100
         TabIndex        =   23
         Top             =   120
         Width           =   5565
      End
      Begin VB.TextBox txtValuer 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   120
         Width           =   1965
      End
      Begin VB.Label lblSummary 
         AutoSize        =   -1  'True
         Caption         =   "����˵��"
         Height          =   180
         Left            =   3360
         TabIndex        =   24
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblValuer 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   360
         TabIndex        =   22
         Top             =   180
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "���(&D)"
      Height          =   350
      Left            =   6720
      TabIndex        =   14
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   12360
      TabIndex        =   13
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   13680
      TabIndex        =   12
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ���䶯��(&P)��"
      Height          =   350
      Left            =   9960
      TabIndex        =   11
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "����ѡ����Ŀ(&I)"
      Height          =   350
      Left            =   8160
      TabIndex        =   10
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Frame fraCondition 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   16335
      Begin VB.CheckBox chkAppAllColumn 
         Caption         =   "�޸ļ۸�Ӧ����������"
         Height          =   255
         Left            =   11040
         TabIndex        =   48
         Top             =   23
         Width           =   2295
      End
      Begin VB.CheckBox chkAutoPay 
         Caption         =   "�Զ�����Ӧ����䶯��¼"
         Height          =   210
         Left            =   4560
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkCostBatch 
         Caption         =   "�ɱ��۰��ⷿ���ε���"
         Height          =   210
         Left            =   2160
         TabIndex        =   39
         Top             =   480
         Width           =   2370
      End
      Begin VB.CheckBox Chk���� 
         Caption         =   "ʱ�����ĸ�Ϊ����"
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1770
      End
      Begin VB.CommandButton cmdPriceMethod 
         Caption         =   "��"
         Height          =   300
         Left            =   3360
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboPriceMethod 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   0
         Width           =   2415
      End
      Begin VB.CheckBox chk������ 
         Caption         =   "�ɱ��۰��ⷿ���ε���"
         Height          =   210
         Left            =   10560
         TabIndex        =   7
         Top             =   -225
         Width           =   2175
      End
      Begin VB.CheckBox chk�Զ�����Ӧ����䶯 
         Caption         =   "�Զ�����Ӧ����䶯"
         Height          =   210
         Left            =   12840
         TabIndex        =   6
         Top             =   -225
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.OptionButton optʱ�� 
         Caption         =   "����ִ��"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   5
         Top             =   8
         Width           =   1095
      End
      Begin VB.OptionButton optʱ�� 
         Caption         =   "ָ������ִ��"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   4
         Top             =   8
         Width           =   1455
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
         Left            =   8040
         TabIndex        =   8
         Top             =   0
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   125829123
         CurrentDate     =   36846.5833333333
      End
      Begin VB.Label lblMethod 
         AutoSize        =   -1  'True
         Caption         =   "���۷�ʽ"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblִ��ʱ�� 
         Caption         =   "ִ��ʱ��"
         Height          =   180
         Left            =   4200
         TabIndex        =   9
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.TextBox txtNO 
      Enabled         =   0   'False
      Height          =   300
      Left            =   13200
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin XtremeSuiteControls.TabControl TabCtlDetails 
      Height          =   975
      Left            =   240
      TabIndex        =   17
      Top             =   5040
      Width           =   1815
      _Version        =   589884
      _ExtentX        =   3201
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfStore 
      Height          =   975
      Left            =   2880
      TabIndex        =   43
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPay 
      Height          =   975
      Left            =   8040
      TabIndex        =   44
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
      TabIndex        =   46
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
      TabIndex        =   47
      Top             =   8190
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16828
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
   Begin VB.Label lblFind 
      Caption         =   "����"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   7488
      Width           =   495
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "������ˮ��"
      Height          =   180
      Left            =   12120
      TabIndex        =   1
      Top             =   180
      Width           =   900
   End
   Begin VB.Label lblDrugName 
      AutoSize        =   -1  'True
      Caption         =   "���ĵ��۵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmStuffPriceCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'����ȫ�ֱ���
Private Const mconlngRowHeight As Long = 300 '����и����и�
Private mintUnit As Integer     '������¼���õ���ʲô��λ
Private mint���� As Integer     '0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
Private mlng��Ӧ��ID As Long  '������¼��Ӧ��id
Private mdbl�ӳ��� As Double
Private mblnӦ����¼ As Boolean '��¼�Ƿ����Ӧ����¼
Private mblnʱ�����İ����ε��� As Boolean 'ʱ�����İ������ε���
Private mint���� As Integer  'mint����=1���������Ŀ¼����
Private mlng���ID As Long '������Ŀ¼�����ȡ�Ĺ��ID
Private mstr������ As String
Private mintSalePriceDigit As Integer
'��ɫ����
Private Const mconlngColor As Long = &HFFFFFF        '�����޸�����ɫΪ��ɫ
Private Const mconlngCanColor As Long = &HE7CFBA    '���޸�����ɫΪ����ɫ

Private mbln�ּ���ʾ As Boolean         '�޼�������ʾ true-��ʾ false-����ʾ
Private mdbl�ֶμӳ��� As Double    '������¼�ֶμӳ���
Private mdbl�ɱ��� As Double            '��¼�޸�֮ǰ�ĳɱ���
Private mstrNo As String            '���۵�No
Private mintModal As Integer        '������ʲô״̬ 0-���� 1-�޸� 2-����
Private mintMethod As Integer   '���۷�ʽ 0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
Private mstr���ۻ��ܺ� As String
Private mblnLoad As Boolean     '�Ƿ�������
Private mrsReturn As ADODB.Recordset '����ѡ�񷵻ص����ݼ�
Private mblnOk As Boolean
Private mrsFindName As ADODB.Recordset '��ѯ�����ݼ�
Private mblnClick As Boolean
Private mintType As Integer      '������ʽ
Private mdbl���� As Double      '������ʽ����д�ĵ������
Private mlngPrice As Long       '��¼�۸�
Private mblnUpdateAdd As Boolean    '�޸�����µ���������
Private mlngOldStuffID As Long '���ԭʼ���Ƿ���ҩƷ
Private mdblOldPrice As Double     '��¼ԭʼ�۸�
Private mblnBatchItem As Boolean   '��¼�Ƿ���������ѡ��ť
Private mstrPrivs As String       'ģ��Ȩ��
Private Const mstrCaption As String = "���ĵ��۵�"

Private mFMT As g_FmtString
Private mOraFMT As g_FmtString

Private Enum menuPriceCol
    ����ID = 0
    ԭ��id = 1
    Ʒ�� = 2
    ��� = 3
    �Ƿ���
    ����
    ��λ
    ��װϵ��
    �Ƿ��������
    �ӳ���
    ���������
    �Ƿ��п��
    ������Ŀid
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
    ����ID = 0
    �ⷿ = 1
    �ⷿID = 2
    ��Ӧ��
    ��Ӧ��ID
    ҩƷ
    ���
    ����
    Ч��
    ����
    ����
    ���
    ����
    ��λ
    ��װϵ��
    ԭ���ۼ�
    �����ۼ�
    �������
    �ӳ���
    ԭ�ɹ���
    �ֲɹ���
    ��۲�
    ������
End Enum

Private Enum menuPayCol
    ����ID = 0
    Ʒ�� = 1
    ��Ӧ��
    ��Ӧ��ID
    ���
    ����
    ��Ʊ��
    ��Ʊ����
    ��Ʊ���
    ������
End Enum

Public Sub ShowMe(ByVal frmParent As Form, ByVal intModal As Integer, ByVal str���ۻ��ܺ� As String, ByVal intMethod As Integer, Optional int���� As Integer, Optional lng���ID As Long)
    mintModal = intModal
    mstr���ۻ��ܺ� = str���ۻ��ܺ�
    mintMethod = intMethod
    mstrPrivs = GetPrivFunc(glngSys, 1726)
    mint���� = int����
    mlng���ID = lng���ID
    
    Me.Show vbModal, frmParent
End Sub

Private Sub cboPriceMethod_Click()
    Dim intCol As Integer
    Dim intTemp As Integer

    With cboPriceMethod
        If .Text = "�����ۼ�" Then
            intTemp = 0
        ElseIf .Text = "�����ɱ���" Then
            intTemp = 1
        Else
            intTemp = 2
        End If
    End With

    If mint���� = 1 Then
        If mblnLoad = True And intTemp <> Val(lblMethod.Tag) Then
            If vsfPrice.TextMatrix(1, menuPriceCol.����ID) <> "" Then
                If MsgBox("���۷�ʽ�ı佫�ָ��б����޸ĵļ۸��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    cboPriceMethod.ListIndex = mint����
                    Exit Sub
                Else
                    mdbl���� = 0
                    mstr������ = ""
                    vsfPrice.Rows = 2
                    For intCol = 0 To vsfPrice.Cols - 1
                        vsfPrice.TextMatrix(1, intCol) = ""
                    Next
                    vsfStore.Rows = 1
                    vsfPay.Rows = 1
                    Call CatalogModifyPrice
                End If
            End If
        End If
    Else
        If mblnLoad = True And intTemp <> Val(lblMethod.Tag) Then
            If vsfPrice.TextMatrix(1, menuPriceCol.����ID) <> "" Then
                If MsgBox("���۷�ʽ�ı佫����б������ݣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    cboPriceMethod.ListIndex = mint����
                    Exit Sub
                Else
                    mdbl���� = 0
                    mstr������ = ""
                    vsfPrice.Rows = 2
                    For intCol = 0 To vsfPrice.Cols - 1
                        vsfPrice.TextMatrix(1, intCol) = ""
                    Next
                    vsfStore.Rows = 1
                    vsfPay.Rows = 1
                End If
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
            chkCostBatch.Visible = False
            chkCostBatch.Value = False
            chkAutoPay.Visible = False
            chkAutoPay.Value = 0
            TabCtlDetails.Item(1).Visible = False
        ElseIf .Text = "�����ɱ���" Then
            mint���� = 1
            lblMethod.Tag = 1
            optʱ��(0).Value = True
            optʱ��(0).Enabled = False
            optʱ��(1).Enabled = False
            dtpRunDate.Enabled = False
            chkCostBatch.Visible = True
            If mblnӦ����¼ = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
                TabCtlDetails.Item(1).Visible = True
            End If
        ElseIf .Text = "�ۼ۳ɱ���һ�����" Then
            mint���� = 2
            lblMethod.Tag = 2
            optʱ��(0).Value = False
            optʱ��(1).Value = True
            optʱ��(0).Enabled = True
            optʱ��(1).Enabled = True
            dtpRunDate.Enabled = True
            chkCostBatch.Visible = True
            If mblnӦ����¼ = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
                TabCtlDetails.Item(1).Visible = True
            End If
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

Private Sub Chk��Ӧ��_Click()
    If chk��Ӧ��.Value = 1 Then
        Cmd��Ӧ��.Enabled = True
        txt��Ӧ��.Enabled = True
        chkӦ����¼.Enabled = True
    Else
        Cmd��Ӧ��.Enabled = False
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

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim intCol As Integer

    If MsgBox("��ȷ��Ҫ����������ݣ�", vbYesNo, gstrSysName) = vbYes Then
        mdbl���� = 0
        mstr������ = ""
        vsfPrice.Rows = 2
        For intCol = 0 To vsfPrice.Cols - 1
            vsfPrice.TextMatrix(1, intCol) = ""
        Next
        vsfStore.Rows = 1
        vsfPay.Rows = 1
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
        If Val(.TextMatrix(1, menuPriceCol.����ID)) <> 0 Then
            If MsgBox("����ձ���е����ݣ��Ƿ������", vbYesNo, gstrSysName) = vbNo Then
                Exit Sub
            Else
                vsfPrice.Rows = 2
                For i = 0 To vsfPrice.Cols - 1
                    .TextMatrix(1, i) = ""
                Next
                vsfStore.Rows = 1
                vsfPay.Rows = 1
            End If
        End If
    End With

    mlng��Ӧ��ID = IIf(chk��Ӧ��.Value = 1, Val(Split(txt��Ӧ��.Tag, "|")(0)), 0)
    mdbl�ӳ��� = IIf(chk�ӳ���.Value = 1, Val(Trim(txt�ӳ���.Text)), 0)
    mblnӦ����¼ = (chkӦ����¼.Enabled And chkӦ����¼.Value = 1)
    picOtherSelect.Visible = False
    If mblnӦ����¼ = True Then
        TabCtlDetails.Item(1).Visible = True
        chkAutoPay.Visible = True
        chkAutoPay.Value = 1
    Else
        TabCtlDetails.Item(1).Visible = False
        chkAutoPay.Visible = False
        chkAutoPay.Value = 0
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdItem_Click()
    Dim intRow As Integer

    frmBatchSelect.ShowMe Me, mrsReturn, mblnOk, mintType, mdbl����, mint����, mstr������

    On Error GoTo ErrHandle
    If mblnOk = False Then Exit Sub
    If mrsReturn.RecordCount = 0 Then Exit Sub

    With vsfPrice
        If .TextMatrix(.Rows - 1, menuPriceCol.����ID) = "" Then
            intRow = .Rows - 1
        Else
            .Rows = .Rows + 1
            intRow = .Rows - 1
        End If
    End With
    mblnBatchItem = True
    Call GetDrugPirce(mrsReturn, intRow)
    mblnBatchItem = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub deleteNotExecutePirce()
    '���δִ�м۸�
    Dim intRow As Integer

    On Error GoTo ErrHandle
    With vsfPrice
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, menuPriceCol.����ID) <> "" Then
                gstrSQL = "Zl_ɾ������δִ�м۸�_Delete(" & Val(.TextMatrix(intRow, menuPriceCol.����ID)) & "," & 0 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
        Next
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOk_Click()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim dtToDay As Date
    Dim lngAdjId As Long
    Dim lngId As Long
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
    Dim lng����ID As Long
    Dim lng����  As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim str���� As String
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim dbl��Ʊ��� As Double

    Dim lng��� As Long
    Dim cllProc As Collection
    Dim strTemp As String
    Dim j As Integer
    Dim dbl�ɱ��� As Double

    Set cllProc = New Collection

    If vsfPrice.Rows > 1 Then
        If Val(vsfPrice.TextMatrix(1, menuPriceCol.����ID)) = 0 Then Exit Sub
    End If
    If CheckPrice = False Then Exit Sub

    On Error GoTo ErrHand
    dtToDay = sys.Currentdate()

    gstrSQL = "select �շѼ�Ŀ_ID.nextval from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�շѼ�Ŀ���")
    lngAdjId = rsTemp.Fields(0).Value

    If mintModal = 1 Then '�޸� ���޸�ģʽ����ɾ��ԭ���ĵ�����Ϣ��Ȼ������µĵ�����Ϣ
        Call deleteNotExecutePirce
    End If

    '����Ƿ����δִ�еļ۸�
    If checkNotExecutePrice = True Then Exit Sub
    '��ȡ����NO
    mstrNo = sys.GetNextNo(9)
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
        lng��� = 1
        For intCount = 1 To IIf(Trim(.TextMatrix(.Rows - 1, 0)) = "", .Rows - 2, .Rows - 1)
            lng����ID = Val(.TextMatrix(intCount, menuPriceCol.����ID))
            dbl��װ = Val(.TextMatrix(intCount, menuPriceCol.��װϵ��))
            
            If lng����ID <> 0 Then
                If Val(.TextMatrix(intCount, menuPriceCol.�����ۼ�)) <> Val(.TextMatrix(intCount, menuPriceCol.ԭ���ۼ�)) Then
                    lngId = sys.NextId("�շѼ�Ŀ")
                    If optʱ��(0).Value = True Then
                        strID = strID & "," & lngId
                    ElseIf lng����ID = -1 Then
                        strID = strID & "," & lngId
                    End If
                    
                    If .TextMatrix(intCount, menuPriceCol.�Ƿ���) = "1" And mblnʱ�����İ����ε��� And mint���� <> 1 Then
                        strTmp = ""
                        lngCurrBatch = -1
                        For n = 1 To vsfStore.Rows - 1
                            If Val(.TextMatrix(intCount, menuPriceCol.����ID)) = Val(vsfStore.TextMatrix(n, menuStoreCol.����ID)) Then
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
                        gstrSQL = "zl_�շѼ�Ŀ_stop("
                                    '    �շ�ϸĿID_IN IN �շѼ�Ŀ.�շ�ϸĿID%TYPE,
                                    gstrSQL = gstrSQL & "" & lng����ID & ","
                                    '    ��ֹ����_IN IN �շѼ�Ŀ.��ֹ����%TYPE := NULL
                                    If optʱ��(0).Value Then
                                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToDay), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                    Else
                                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                    End If
                                    gstrSQL = gstrSQL & ")"
                                    AddArray cllProc, gstrSQL

                                    'Zl_�շѼ�Ŀ_Insert
                                    gstrSQL = "zl_�շѼ�Ŀ_Insert("
                                    '  Id_In         In �շѼ�Ŀ.ID%Type,
                                    gstrSQL = gstrSQL & "" & lngId & ","
                                    '  ԭ��id_In     In �շѼ�Ŀ.ԭ��id%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIf(.TextMatrix(intCount, menuPriceCol.ԭ��id) = "", "NUll", Val(.TextMatrix(intCount, menuPriceCol.ԭ��id))) & ","
                                    '  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type := Null,
                                    gstrSQL = gstrSQL & "" & lng����ID & ","
                                    '  ������Ŀid_In In �շѼ�Ŀ.������Ŀid%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(intCount, menuPriceCol.������Ŀid)) & ","
                                    '  ԭ��_In       In �շѼ�Ŀ.ԭ��%Type := Null,
                                    If .TextMatrix(intCount, menuPriceCol.�Ƿ���) = "1" And Val(.TextMatrix(intCount, menuPriceCol.�Ƿ��������)) = 0 Then
                                        '�Ǹ����������ϵ�ʵ�����ģ����Է�Χ�����ģ�����Ҫ��ҽ��Ӧ��),ʼ����Ϊ��
                                        gstrSQL = gstrSQL & "" & 0 & ","
                                    Else
                                        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(intCount, menuPriceCol.ԭ���ۼ�)) / dbl��װ, g_С��λ��.obj_���С��.���ۼ�С��) & ","
                                    End If

                                    '  �ּ�_In       In �շѼ�Ŀ.�ּ�%Type := Null,
                                    gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(intCount, menuPriceCol.�����ۼ�)) / dbl��װ, g_С��λ��.obj_���С��.���ۼ�С��) & ","
                                    '  �����շ���_In In �շѼ�Ŀ.�����շ���%Type := Null,
                                    gstrSQL = gstrSQL & "NULL,"
                                    '  �Ӱ�Ӽ���_In In �շѼ�Ŀ.�Ӱ�Ӽ���%Type := Null,
                                    gstrSQL = gstrSQL & "NULL,"
                                    '  ����˵��_In   In �շѼ�Ŀ.����˵��%Type := Null,
                                    gstrSQL = gstrSQL & "'" & Me.txtSummary.Text & "',"
                                    '  ����id_In     In �շѼ�Ŀ.����id%Type := Null,
                                    gstrSQL = gstrSQL & "" & lngAdjId & ","
                                    '  ������_In     In �շѼ�Ŀ.������%Type := Null,
                                    gstrSQL = gstrSQL & "'" & Me.txtValuer.Text & "',"
                                    '  ִ������_In   In �շѼ�Ŀ.ִ������%Type := Null,
                                    If Me.optʱ��(0).Value Then
                                        gstrSQL = gstrSQL & "to_date('" & Format(dtToDay, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                                    Else
                                        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                                    End If
                                    '  �䶯ԭ��_In   In �շѼ�Ŀ.�䶯ԭ��%Type := 1,
                                    gstrSQL = gstrSQL & "" & 0 & ","
                                    '  No_In         In �շѼ�Ŀ.NO%Type := Null,
                                    gstrSQL = gstrSQL & "'" & mstrNo & "',"
                                    '  ���_In       In �շѼ�Ŀ.���%Type := 1
                                    gstrSQL = gstrSQL & "" & lng��� & ","
                                    'ȱʡ�۸�_In
                                    If .TextMatrix(intCount, menuPriceCol.�Ƿ���) = "1" And Val(.TextMatrix(intCount, menuPriceCol.�Ƿ��������)) = 0 Then
                                            gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(intCount, menuPriceCol.�����ۼ�)) / dbl��װ, g_С��λ��.obj_���С��.���ۼ�С��) & ","
                                    Else
                                            gstrSQL = gstrSQL & "NULL,"
                                    End If
                                    '���ۻ��ܺ�
                                    gstrSQL = gstrSQL & "" & txtNO.Text & ")"
                                    AddArray cllProc, gstrSQL
                                    lng��� = lng��� + 1
                        blnPrice = True
                        blnPrint = True
                    End If
                End If
                If lng����ID <> 0 Then
                    If Val(.TextMatrix(intCount, menuPriceCol.ԭָ���ۼ�)) <> Val(.TextMatrix(intCount, menuPriceCol.��ָ���ۼ�)) Then
                        strTemp = Round(Val(.TextMatrix(intCount, menuPriceCol.��ָ���ۼ�)) / dbl��װ, g_С��λ��.obj_���С��.���ۼ�С��)
                        'zl_��������_UpdateCustom ( ����ID_IN ,SQL_IN)
                        gstrSQL = "zl_��������_UpdateCustom(" & lng����ID & ",'ָ�����ۼ�=" & strTemp & "')"
                        AddArray cllProc, gstrSQL
                    End If
                    '���²ɹ��޼�
                    If Val(.TextMatrix(intCount, menuPriceCol.ԭ�ɹ��޼�)) <> Val(.TextMatrix(intCount, menuPriceCol.�ֲɹ��޼�)) Then
                        strTemp = Round(Val(.TextMatrix(intCount, menuPriceCol.�ֲɹ��޼�)) / dbl��װ, g_С��λ��.obj_���С��.�ɱ���С��)
                        'zl_��������_UpdateCustom ( ����ID_IN ,SQL_IN)
                        gstrSQL = "zl_��������_UpdateCustom(" & lng����ID & ",'ָ��������=" & strTemp & "')"
                        AddArray cllProc, gstrSQL
                    End If
                End If
            End If
        Next
    End With

    '�ɱ��۵��۴���
    If mint���� = 1 Or mint���� = 2 Then
        With vsfStore
            For i = 1 To .Rows - 1
                lng�ⷿID = Val(.TextMatrix(i, menuStoreCol.�ⷿID))
                lng��Ӧ��ID = Val(.TextMatrix(i, menuStoreCol.��Ӧ��ID))
                lng����ID = Val(.TextMatrix(i, menuStoreCol.����ID))
                lng���� = Val(.TextMatrix(i, menuStoreCol.����))
                str���� = .TextMatrix(i, menuStoreCol.����)
                dbl��װ = Val(.TextMatrix(i, menuStoreCol.��װϵ��))
                If lng����ID <> 0 Then
                    str��Ʊ�� = "": str��Ʊ���� = "": dbl��Ʊ��� = 0
                    If chkAutoPay.Value = 1 Then
                        With vsfPay
                            For j = 1 To .Rows - 1
                                If Val(.TextMatrix(j, menuPayCol.����ID)) = lng����ID And _
                                    Val(.TextMatrix(j, menuPayCol.��Ӧ��ID)) = lng��Ӧ��ID Then
                                    '���Ƿ��д��������Ͽ��䶯���
                                    str��Ʊ�� = Trim(.TextMatrix(j, menuPayCol.��Ʊ��))
                                    str��Ʊ���� = Trim(.TextMatrix(j, menuPayCol.��Ʊ����))
                                    dbl��Ʊ��� = Val(.TextMatrix(j, menuPayCol.��Ʊ���))
                                    Exit For
                                End If
                            Next
                        End With
                    End If

                    dbl�ɱ��� = Round(Val(.TextMatrix(i, menuStoreCol.�ֲɹ���)) / dbl��װ, g_С��λ��.obj_���С��.�ɱ���С��)

                    ' Zl_���ϳɱ�����_Insert
                    gstrSQL = "Zl_���ϳɱ�����_Insert("
                    '  ��ҩ��λid_In In �ɱ��۵�����Ϣ.��ҩ��λid%Type,
                    gstrSQL = gstrSQL & IIf(lng��Ӧ��ID = 0, "Null", lng��Ӧ��ID) & ","
                    '  �ⷿid_In     In �ɱ��۵�����Ϣ.�ⷿid%Type,
                    gstrSQL = gstrSQL & "" & lng�ⷿID & ","
                    '  ����id_In     In �ɱ��۵�����Ϣ.ҩƷid%Type,
                    gstrSQL = gstrSQL & "" & lng����ID & ","
                    '  ����_In       In �ɱ��۵�����Ϣ.����%Type := Null,
                    gstrSQL = gstrSQL & "" & lng���� & ","
                    '����_in
                    gstrSQL = gstrSQL & "" & IIf(str���� = "", "NULL", "'" & str���� & "'") & ","
                    '  ԭ�ɱ���_In   In �ɱ��۵�����Ϣ.ԭ�ɱ���%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(i, menuStoreCol.ԭ�ɹ���)) / dbl��װ, g_С��λ��.obj_���С��.�ɱ���С��) & ","
                    '  �³ɱ���_In   In �ɱ��۵�����Ϣ.�³ɱ���%Type := Null,
                    gstrSQL = gstrSQL & "" & dbl�ɱ��� & ","
                    '  ��Ʊ��_In     In �ɱ��۵�����Ϣ.��Ʊ��%Type := Null,
                    gstrSQL = gstrSQL & "'" & str��Ʊ�� & "',"
                    '  ��Ʊ����_In   In �ɱ��۵�����Ϣ.��Ʊ����%Type := Null,
                    gstrSQL = gstrSQL & "" & IIf(str��Ʊ���� = "", "NULL", "to_date('" & str��Ʊ���� & "','yyyy-mm-dd') ") & ","
                    '  ��Ʊ���_In   In �ɱ��۵�����Ϣ.��Ʊ���%Type := Null,
                    gstrSQL = gstrSQL & "" & dbl��Ʊ��� & ","
                    '  Ӧ����䶯_In In �ɱ��۵�����Ϣ.Ӧ����䶯%Type := 0
                    gstrSQL = gstrSQL & "" & IIf(chkAutoPay.Value = 1 And lng��Ӧ��ID <> 0 And dbl��Ʊ��� <> 0, 1, 0) & ","
                    gstrSQL = gstrSQL & "'" & txtNO.Text & "')"
                    AddArray cllProc, gstrSQL
                    blnCost = True
                End If
            Next
        End With
    End If

    '�޿��ʱ�����ɱ���
    If mint���� = 1 Or mint���� = 2 Then
        With Me.vsfPrice
            For intCount = 1 To .Rows - 1
                lng����ID = Val(.TextMatrix(intCount, menuPriceCol.����ID))
                dbl��װ = Val(.TextMatrix(intCount, menuStoreCol.��װϵ��))
                If lng����ID <> 0 Then
                    If .TextMatrix(intCount, menuPriceCol.�Ƿ��п��) = "0" And Val(.TextMatrix(intCount, menuPriceCol.ԭ�ɱ���)) <> Val(.TextMatrix(intCount, menuPriceCol.�ֳɱ���)) Then
                        dbl��װ = Val(.TextMatrix(intCount, menuPriceCol.��װϵ��))
    
                        lng����ID = Val(.TextMatrix(intCount, menuPriceCol.����ID))
                        dblOldCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.ԭ�ɱ���)) / dbl��װ, g_С��λ��.obj_���С��.�ɱ���С��))
                        dblNewCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.�ֳɱ���)) / dbl��װ, g_С��λ��.obj_���С��.�ɱ���С��))
    
                        gstrSQL = "Zl_���ϳɱ�����_Insert(Null,Null," & lng����ID & ",0,NULL" & "," & dblOldCost & ", " & dblNewCost & ",NULL,Null,0,0, '" & txtNO.Text & "')"
                        AddArray cllProc, gstrSQL
                        blnCost = True
                    End If
                End If
            Next
        End With
    End If

   '����������¶Գɱ��۽��е���:
    '1.����Ϊ�ɱ��۵��ۼ�����ִ��ʱ�������Գɱ��۽��е���
    '2.��������ִ�кͷǳɱ���(���ɱ��۵��۷�ʽ)����ʱ�����������ϵ���ʱ����ִ�С�
     '�����ɱ��۵���ʱ
    If mint���� = 1 Then
        If Me.optʱ��(0).Value = True Then
            With vsfPrice
                For i = 1 To .Rows - 1
                    lng����ID = Val(.TextMatrix(i, menuPriceCol.����ID))
                    If lng����ID <> 0 Then
                      ' Zl_�����շ���¼_Adjust
                      gstrSQL = "Zl_�����շ���¼_Adjust("
                      '  ����id_In In Number, --���ۼ�¼��ID
                      gstrSQL = gstrSQL & "" & 0 & ","
                      '  ����_In   In Number := 0, --�Ƿ�תΪ�������ۣ����²������ԡ��շ�ϸĿ�еı�ۣ�
                      gstrSQL = gstrSQL & "" & 0 & ","
                      '  ����id_In In Number := 0 --����Ϊ0ʱ��ʾ�ǳɱ��۵��ۣ��������ۼ��������
                        gstrSQL = gstrSQL & "" & lng����ID & ")"
                      AddArray cllProc, gstrSQL
                    End If
                Next
            End With
        End If
    Else
        '���ۼ�
        If strID <> "" Then strID = Mid(strID, 2)
        ArrayID = Split(strID, ",")
        Array���μ۸� = Split(str���μ۸�, ";")
        For intCount = 0 To UBound(ArrayID)
            If optʱ��(0).Value = True Or vsfPrice.TextMatrix(intCount + 1, menuPriceCol.ԭ��id) = "" Then
                gstrSQL = "zl_�����շ���¼_Adjust(" & ArrayID(intCount) & "," & Me.Chk����.Value & ",0,'" & Array���μ۸�(intCount) & "')"
                AddArray cllProc, gstrSQL
            End If
        Next
    End If

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
    gstrSQL = gstrSQL & IIf(txtSummary.Text = "", "Null", "'" & txtSummary.Text & "'") & ",1,'" & UserInfo.�û��� & "')"

    AddArray cllProc, gstrSQL

'    gcnOracle.BeginTrans
    ExecuteProcedureArrAy cllProc, mstrCaption
'    gcnOracle.CommitTrans

    If blnPrint = True Then
        If MsgBox("����Ҫ��ӡ����֪ͨ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1726_1", Me, "���ۺ�=" & txtNO.Text, "���㵥λ=" & mintUnit, 2)
        End If
    End If

    '����б�������
    With vsfPrice
        .Rows = 2
        For intCol = 0 To .Cols - 1
            .TextMatrix(1, intCol) = ""
        Next
    End With
    vsfStore.Rows = 1
    vsfPay.Rows = 1
    txtNO.Text = ""
    txtSummary.Text = ""
    
    If mint���� = 1 Then
        Unload Me
    End If
    
    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Function CheckUnVerify(ByVal lng����ID As Long) As Boolean
    '��������Ƿ����δ��˵���
    Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandle
    gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ����id = [1] And Rownum = 1 And ������� Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������Ƿ����δ��˵���", lng����ID)

    If rsTemp.RecordCount > 0 Then
        CheckUnVerify = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function checkNotExecutePrice(Optional ByVal lngDrugID As Long = 0) As Boolean
    '���� ������Ƿ����δִ�еļ۸�
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long, IntCheck As Integer

    err = 0
    On Error GoTo ErrHand

    If lngDrugID = 0 Then
        'ѭ���ж���������
        For IntCheck = 1 To vsfPrice.Rows - 1
            LngmediIDThis = Val(vsfPrice.TextMatrix(IntCheck, menuPriceCol.����ID))
            If LngmediIDThis <> 0 Then
                If mint���� = 0 Or mint���� = 2 Then
                    '�ж��Ƿ���δִ�е���ʷ�۸�
                    gstrSQL = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And ִ������ > Sysdate And �շ�ϸĿID=[1]" & _
                            GetPriceClassString("")
                    
                    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, LngmediIDThis)

                    With RecCheck
                        If Not .EOF Then
                            If Not IsNull(!Records) Then
                                If !Records <> 0 Then
                                    MsgBox "����" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.Ʒ��) & "����δִ�м۸�δִ�����Ĳ��ܵ��ۣ�", vbInformation, gstrSysName
                                    checkNotExecutePrice = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End With
                End If

                If mint���� = 1 Or mint���� = 2 Then
                    '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
                    gstrSQL = "Select 1 From �ɱ��۵�����Ϣ Where ҩƷid = [1] And ִ������ Is Null And Rownum = 1 "
                    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, LngmediIDThis)

                    If RecCheck.RecordCount > 0 Then
                        MsgBox "����" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.Ʒ��) & "����δִ�гɱ��ۣ�δִ�����Ĳ��ܵ��ۣ�", vbInformation, gstrSysName
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
            
            Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngDrugID)

            With RecCheck
                If Not .EOF Then
                    If Not IsNull(!Records) Then
                        If !Records <> 0 Then
                            MsgBox "������δִ�е��ۼ۵��ۼ�¼��δִ�����Ĳ��ܵ��ۣ�", vbInformation, gstrSysName
                            checkNotExecutePrice = True
                            Exit Function
                        End If
                    End If
                End If
            End With
        End If

        If mint���� = 1 Or mint���� = 2 Then
            '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
            gstrSQL = "Select 1 From �ɱ��۵�����Ϣ Where ҩƷid = [1] And ִ������ Is Null And Rownum = 1 "
            Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngDrugID)

            If RecCheck.RecordCount > 0 Then
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
    Dim bln�޿�� As Boolean

    '����ִ�м۸��Ƿ���ȷ
    '�Լ�������Ŀ��ͬ��������ּ��Ƿ���ԭ����ͬ
    CheckPrice = False
    With vsfPrice
        For IntCheck = 1 To .Rows - 1
            If Val(.TextMatrix(IntCheck, menuPriceCol.����ID)) <> 0 Then
                If Not IsNumeric(Trim(.TextMatrix(IntCheck, menuPriceCol.�����ۼ�))) Then
                    MsgBox "��" & IntCheck & "�е������ۼ��к��зǷ��ַ���", vbInformation, gstrSysName
                    .Row = IntCheck
                    .Col = menuPriceCol.�����ۼ�
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If

                If mint���� <> 1 Then
                    If Val(.TextMatrix(IntCheck, menuPriceCol.�����ۼ�)) = Val(.TextMatrix(IntCheck, menuPriceCol.ԭ���ۼ�)) Then
                        MsgBox "��" & IntCheck & "�е������ּ���ԭ����ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.�����ۼ�
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                End If

'                If mint���� <> 0 Then
'                    If Val(.TextMatrix(IntCheck, menuPriceCol.�ֳɱ���)) = Val(.TextMatrix(IntCheck, menuPriceCol.ԭ�ɱ���)) Then
'                        MsgBox "��" & IntCheck & "�е�ҩƷ�ֳɱ�����ԭ�ɱ�����ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
'                        .Row = IntCheck
'                        .Col = menuPriceCol.�ֳɱ���
'                        vsfPrice.SetFocus
'                        .Select IntCheck, 0, IntCheck, .Cols - 1
'                        .TopRow = IntCheck
'                        Exit Function
'                    End If
'                End If

                If .TextMatrix(IntCheck, menuPriceCol.�Ƿ���) = "1" And optʱ��(0).Value <> True And mint���� <> 1 Then
                    MsgBox "��" & IntCheck & "��Ϊʱ�����ģ���������Ϊ����ִ�У�", vbInformation, gstrSysName
                    .Row = IntCheck
                    .Col = menuPriceCol.�����ۼ�
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If
            End If
        Next
    End With

    CheckPrice = True
End Function


Private Sub cmdPriceMethod_Click()
    If txt��Ӧ��.Tag = "" Then
        Me.txt��Ӧ��.Tag = "0|"
    End If
    picOtherSelect.Visible = True
End Sub

Private Sub CmdPrint_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim i As Long


    If vsfStore.Rows = 1 Then Exit Sub
    If Trim(vsfStore.TextMatrix(1, menuStoreCol.����ID)) = "" Then Exit Sub

    objPrint.Title.Text = "���ۿ��䶯��"

    Set objRow = New zlTabAppRow
    objRow.Add "����˵��:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "ִ��ʱ��:" & Format(IIf(Me.optʱ��(0).Value, sys.Currentdate, Me.dtpRunDate.Value), "yyyy��MM��DD�� HH:mm:ss")
    objRow.Add "������:" & Me.txtValuer.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡʱ��:" & Format(sys.Currentdate, "yyyy��MM��DD�� HH:mm:ss")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = vsfStore
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

    On Error GoTo ErrHandle
    gstrSQL = "Select ����,����,����,id" & _
        " From ��Ӧ��" & _
        " where ĩ��=1 And substr(����,5,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Activate()
    If mblnLoad = False Then
        vsfPrice.SetFocus
    End If
    If mblnClick = False Then
        vsfPrice.Row = 1
        vsfPrice.Col = menuPriceCol.Ʒ��
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

    Me.Height = 768 * 15
    Me.Width = 1024 * 15
    '��ȡ���õĵ�λ
    mintUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, 1726, 1))
    mblnʱ�����İ����ε��� = Val(zlDatabase.GetPara("ʱ�����İ����ε���", glngSys, 1726, 0))
    
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With

    '��ʼ��ʱ��Ϊ��ǰʱ��+1��
    StrToday = Format(sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")

    If mintModal = 0 Then '������ʱ����Сʱ������Ϊ��ǰʱ��+1��
        Me.dtpRunDate.MinDate = DateAdd("s", 1, CDate(StrToday))
    End If
    Me.dtpRunDate.Value = DateAdd("d", 1, CDate(StrToday))

    txtValuer.Text = gstrUserName

    txtNO.Text = IIf(mintModal = 0, "", mstr���ۻ��ܺ�)
    If mintModal = 0 Then
        LblNo.Visible = False
        txtNO.Visible = False
    End If

    Call InitTabControl
    Call initComboBox '��ʼ�������ؼ�

    If mintModal = 1 Then '�޸�
        If (InStr(1, ";" & gstrPrivs & ";", ";�ɱ��۵���;") > 0 And InStr(1, ";" & gstrPrivs & ";", ";�ۼ۵���;") = 0) Or (InStr(1, ";" & gstrPrivs & ";", ";�ɱ��۵���;") = 0 And InStr(1, ";" & gstrPrivs & ";", ";�ۼ۵���;") > 0) Then
            cboPriceMethod.ListIndex = 0
        ElseIf (InStr(1, ";" & gstrPrivs & ";", ";�ɱ��۵���;") > 0 And InStr(1, ";" & gstrPrivs & ";", ";�ۼ۵���;") > 0) Then
            cboPriceMethod.ListIndex = mintMethod
        End If
    ElseIf mintModal = 2 Then '����
        cboPriceMethod.ListIndex = mintMethod
    End If

    Call InitVsfGridFlex

    Call RestoreWinState(Me, App.ProductName, mstrCaption)
    If mblnӦ����¼ = False Then
        TabCtlDetails.Item(1).Visible = False
    End If
    If mintModal <> 0 Then
        Call initGrid
    End If

    If mintModal = 2 Then '����
        cboPriceMethod.Enabled = False
        cmdPriceMethod.Enabled = False
        optʱ��(0).Enabled = False
        optʱ��(1).Enabled = False
        dtpRunDate.Enabled = False
        Chk����.Enabled = False
        chkCostBatch.Enabled = False
        chkAutoPay.Enabled = False
        txtSummary.Enabled = False
        cmdClear.Visible = False
        cmdItem.Visible = False
        cmdOk.Visible = False
        vsfPrice.Cell(flexcpBackColor, 1, 0, vsfPrice.Rows - 1, vsfPrice.Cols - 1) = mconlngColor
        If vsfStore.Rows > 1 Then
            vsfStore.Cell(flexcpBackColor, 1, 0, vsfStore.Rows - 1, vsfStore.Cols - 1) = mconlngColor
        End If
        If vsfPay.Rows > 1 Then
            vsfPay.Cell(flexcpBackColor, 0, 0, vsfPay.Rows - 1, vsfPay.Cols - 1) = mconlngColor
        End If
    End If
    mblnLoad = True
    If mint���� = 1 Then
        Call CatalogModifyPrice
    End If
End Sub

Private Sub initComboBox()
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
        .InsertItem 0, "���䶯��", vsfStore.hwnd, 0
        .InsertItem 1, "Ӧ����䶯��", vsfPay.hwnd, 0
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
    txtNO.Left = Me.ScaleWidth - txtNO.Width
    LblNo.Left = txtNO.Left - LblNo.Width - 200
    lblDrugName.Left = Me.ScaleWidth / 2 - lblDrugName.Width / 2
    vsfPrice.Move 20, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, 3000
    picSplit.Left = 50
    picSplit.Top = vsfPrice.Top + vsfPrice.Height + 5
    picSplit.Width = Me.ScaleWidth
    txtSummary.Width = Me.ScaleWidth - lblSummary.Left - lblSummary.Width - 300
    TabCtlDetails.Move 20, picSplit.Height + picSplit.Top, Me.ScaleWidth, Me.ScaleHeight - picSplit.Top - picSplit.Height - picInfo.Height - cmdClear.Height - 300 - stbThis.Height
    picInfo.Move 0, TabCtlDetails.Top + TabCtlDetails.Height, Me.ScaleWidth
    lblFind.Top = picInfo.Top + picInfo.Height + 180
    lblFind.Left = 380
    txtFind.Top = lblFind.Top - 50
    txtFind.Left = lblFind.Left + lblFind.Width + 95
    cmdClear.Top = txtFind.Top
    cmdItem.Top = txtFind.Top
    cmdPrint.Top = txtFind.Top
    cmdOk.Top = txtFind.Top
    cmdCanc.Top = txtFind.Top
    cmdCanc.Left = Me.ScaleWidth - cmdCanc.Width - 300
    cmdOk.Left = cmdCanc.Left - cmdOk.Width - 200
    cmdPrint.Left = cmdOk.Left - cmdPrint.Width - 500
    cmdItem.Left = cmdPrint.Left - cmdPrint.Width - 20
    cmdClear.Left = cmdItem.Left - cmdItem.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mstrCaption)
    mdbl���� = 0
    mstr������ = ""
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
        .Rows = 2
        .RowHeight(1) = mconlngRowHeight
        .ColWidth(0) = 200
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mconlngRowHeight
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

        .TextMatrix(0, menuPriceCol.����ID) = "����id"
        .TextMatrix(0, menuPriceCol.ԭ��id) = "ԭ��id"
        .TextMatrix(0, menuPriceCol.Ʒ��) = "Ʒ��"
        .TextMatrix(0, menuPriceCol.���) = "���"
        .TextMatrix(0, menuPriceCol.�Ƿ���) = "�Ƿ���"
        .TextMatrix(0, menuPriceCol.����) = "����"
        .TextMatrix(0, menuPriceCol.��λ) = "��λ"
        .TextMatrix(0, menuPriceCol.��װϵ��) = "��װϵ��"
        .TextMatrix(0, menuPriceCol.�Ƿ��������) = "�Ƿ��������"
        .TextMatrix(0, menuPriceCol.�ӳ���) = "�ӳ���"
        .TextMatrix(0, menuPriceCol.���������) = "���������"
        .TextMatrix(0, menuPriceCol.�Ƿ��п��) = "�Ƿ��п��"
        .TextMatrix(0, menuPriceCol.������Ŀid) = "������Ŀid"
        .TextMatrix(0, menuPriceCol.ԭ�ɱ���) = "ԭ�ɱ���"
        .TextMatrix(0, menuPriceCol.�ֳɱ���) = "�ֳɱ���"
        .TextMatrix(0, menuPriceCol.ԭ���ۼ�) = "ԭ���ۼ�"
        .TextMatrix(0, menuPriceCol.�����ۼ�) = "�����ۼ�"
        .TextMatrix(0, menuPriceCol.ԭ�ɹ��޼�) = "ԭ�ɹ��޼�"
        .TextMatrix(0, menuPriceCol.�ֲɹ��޼�) = "�ֲɹ��޼�"
        .TextMatrix(0, menuPriceCol.ԭָ���ۼ�) = "ԭָ���ۼ�"
        .TextMatrix(0, menuPriceCol.��ָ���ۼ�) = "��ָ���ۼ�"

        '�����п�
        .ColWidth(menuPriceCol.����ID) = 0
        .ColWidth(menuPriceCol.ԭ��id) = 0
        .ColWidth(menuPriceCol.Ʒ��) = 3000
        .ColWidth(menuPriceCol.���) = 1500
        .ColWidth(menuPriceCol.�Ƿ���) = 0
        .ColWidth(menuPriceCol.����) = 2000
        .ColWidth(menuPriceCol.��λ) = 800
        .ColWidth(menuPriceCol.��װϵ��) = 0
        .ColWidth(menuPriceCol.�ӳ���) = 0
        .ColWidth(menuPriceCol.�Ƿ��������) = 0
        .ColWidth(menuPriceCol.���������) = 0
        .ColWidth(menuPriceCol.�Ƿ��п��) = 0
        .ColWidth(menuPriceCol.������Ŀid) = 0
        .ColWidth(menuPriceCol.ԭ�ɱ���) = 1000
        .ColWidth(menuPriceCol.�ֳɱ���) = 1000
        .ColWidth(menuPriceCol.ԭ���ۼ�) = 1000
        .ColWidth(menuPriceCol.�����ۼ�) = 1000
        .ColWidth(menuPriceCol.ԭ�ɹ��޼�) = 0
        .ColWidth(menuPriceCol.�ֲɹ��޼�) = 0
        .ColWidth(menuPriceCol.ԭָ���ۼ�) = 0
        .ColWidth(menuPriceCol.��ָ���ۼ�) = 0
        '���ö��뷽ʽ
        .ColAlignment(menuPriceCol.Ʒ��) = flexAlignLeftCenter
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
        .ColComboList(menuPriceCol.Ʒ��) = "|..."
    End With

    With vsfStore
        .Editable = flexEDNone
        .Cols = menuStoreCol.������
        .Rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mconlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mconlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
'        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMoveRows '�϶�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        '��������
        .TextMatrix(0, menuStoreCol.����ID) = "����id"
        .TextMatrix(0, menuStoreCol.�ⷿ) = "�ⷿ"
        .TextMatrix(0, menuStoreCol.�ⷿID) = "�ⷿid"
        .TextMatrix(0, menuStoreCol.��Ӧ��) = "��Ӧ��"
        .TextMatrix(0, menuStoreCol.��Ӧ��ID) = "��Ӧ��id"
        .TextMatrix(0, menuStoreCol.ҩƷ) = "����"
        .TextMatrix(0, menuStoreCol.���) = "���"
        .TextMatrix(0, menuStoreCol.��λ) = "��λ"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.Ч��) = "Ч��"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.��װϵ��) = "��װϵ��"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.���) = "���"
        .TextMatrix(0, menuStoreCol.ԭ���ۼ�) = "ԭ���ۼ�"
        .TextMatrix(0, menuStoreCol.�����ۼ�) = "�����ۼ�"
        .TextMatrix(0, menuStoreCol.�������) = "�������"
        .TextMatrix(0, menuStoreCol.�ӳ���) = "�ӳ���"
        .TextMatrix(0, menuStoreCol.ԭ�ɹ���) = "ԭ�ɹ���"
        .TextMatrix(0, menuStoreCol.�ֲɹ���) = "�ֲɹ���"
        .TextMatrix(0, menuStoreCol.��۲�) = "��۲�"
        '�����п�
        .ColWidth(0) = 0
        .ColWidth(menuStoreCol.�ⷿ) = 1500
        .ColWidth(menuStoreCol.�ⷿID) = 0
        .ColWidth(menuStoreCol.��Ӧ��) = 2000
        .ColWidth(menuStoreCol.��Ӧ��ID) = 0
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
        .ColWidth(menuStoreCol.�������) = 1000
        .ColWidth(menuStoreCol.�ӳ���) = 1000
        .ColWidth(menuStoreCol.ԭ�ɹ���) = 1000
        .ColWidth(menuStoreCol.�ֲɹ���) = 1000
        .ColWidth(menuStoreCol.��۲�) = 1000
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
        .ColAlignment(menuStoreCol.�������) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�ӳ���) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.ԭ�ɹ���) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�ֲɹ���) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.��۲�) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '��ͷ���ж���
    End With

    With vsfPay
        .Editable = flexEDNone
        .Cols = menuPayCol.������
        .Rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mconlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mconlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
'        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMoveRows '�϶�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        .TextMatrix(0, menuPayCol.����ID) = "����id"
        .TextMatrix(0, menuPayCol.��Ӧ��) = "��Ӧ��"
        .TextMatrix(0, menuPayCol.��Ӧ��ID) = "��Ӧ��id"
        .TextMatrix(0, menuPayCol.Ʒ��) = "Ʒ��"
        .TextMatrix(0, menuPayCol.��Ʊ��) = "��Ʊ��"
        .TextMatrix(0, menuPayCol.��Ʊ����) = "��Ʊ����"
        .TextMatrix(0, menuPayCol.��Ʊ���) = "��Ʊ���"
        .TextMatrix(0, menuPayCol.���) = "���"
        .TextMatrix(0, menuPayCol.����) = "����"
        '�����п�
        .ColWidth(menuPayCol.����ID) = 0
        .ColWidth(menuPayCol.��Ӧ��) = 1500
        .ColWidth(menuPayCol.Ʒ��) = 2000
        .ColWidth(menuPayCol.��Ʊ��) = 1500
        .ColWidth(menuPayCol.��Ʊ����) = 2000
        .ColWidth(menuPayCol.��Ʊ���) = 1500
        .ColHidden(menuPayCol.��Ӧ��ID) = True
        .ColHidden(menuPayCol.���) = True
        .ColHidden(menuPayCol.����) = True
        '���뷽ʽ
        .ColAlignment(menuPayCol.Ʒ��) = flexAlignLeftCenter
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

    On Error GoTo ErrHandle
    '���۷�ʽ 0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
    If mintMethod = 0 Then
        gstrSQL = "Select Distinct p.ԭ��id, i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "                 nvl(s.�ӳ���,0) / 100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, i.���� As ����, i.���㵥λ As ��λ," & vbNewLine & _
            "                s.��װ��λ,s.����ϵ��, s.�ɱ��� As ԭ�ɱ���, s.�ɱ��� As �³ɱ���, p.ԭ��, p.�ּ�," & vbNewLine & _
            "                p.������Ŀid, p.������, p.����˵��, s.���������, To_Char(a.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ����id," & vbNewLine & _
            "                Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select ҩƷid From ҩƷ��� where ����=1) K, ���ۻ��ܼ�¼ A, �շ���Ŀ���� B, �������� S, �շ���ĿĿ¼ I, �շѼ�Ŀ P" & vbNewLine & _
            "Where a.���ۺ� = p.���ۻ��ܺ� And b.�շ�ϸĿid(+) = s.����id And s.����id = i.Id And i.Id = k.ҩƷid(+) And i.Id = p.�շ�ϸĿid And" & vbNewLine & _
            "      p.���ۻ��ܺ� = [1] And a.���� = 1 And b.����(+) = 3 And a.���ۺ� = [1] " & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            GetPriceClassString("P") & " Order By ����id"
    ElseIf mintMethod = 1 Then
        gstrSQL = "Select Distinct i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "                nvl(s.�ӳ���,0) / 100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, m.���� As ����, i.���㵥λ As ��λ," & vbNewLine & _
            "                s.��װ��λ,s.����ϵ��, m.ԭ�ɱ���, m.�³ɱ���, p.�ּ� as ԭ��, p.�ּ�, p.������Ŀid," & vbNewLine & _
            "                a.������ As ������, a.˵�� As ����˵��, s.���������, To_Char(m.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ����id," & vbNewLine & _
            "                Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select Min(ԭ�ɱ���) As ԭ�ɱ���, Min(�³ɱ���) As �³ɱ���, min(����) as ����,���ۻ��ܺ�,ҩƷid,min(ִ������) as ִ������ From �ɱ��۵�����Ϣ Where ���ۻ��ܺ� = [1] Group By ���ۻ��ܺ�,ҩƷid) M, (Select ҩƷid From ҩƷ��� where ����=1) K, ���ۻ��ܼ�¼ A, �շ���Ŀ���� B, �������� S, �շ���ĿĿ¼ I, �շѼ�Ŀ P" & vbNewLine & _
            "Where m.���ۻ��ܺ�(+) = a.���ۺ� And b.�շ�ϸĿid(+) = s.����id And s.����id = i.Id And i.Id = k.ҩƷid(+) And m.ҩƷid = i.Id And" & vbNewLine & _
            "      i.Id = p.�շ�ϸĿid And Sysdate Between p.ִ������ And p.��ֹ���� And m.���ۻ��ܺ� = [1] And a.���� = 1 And b.����(+) = 3 And" & vbNewLine & _
            "      a.���ۺ� = [1] " & IIf(mintModal = 2, "", " And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            GetPriceClassString("P") & " Order By ����id"
    ElseIf mintMethod = 2 Then
        gstrSQL = "Select distinct p.ԭ��id, i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "       nvl(s.�ӳ���,0) / 100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, decode(m.����,null,i.����,m.����) As ����, i.���㵥λ As ��λ," & vbNewLine & _
            "       s.��װ��λ,s.����ϵ��, m.ԭ�ɱ���, m.�³ɱ���, p.ԭ��, p.�ּ�, p.������Ŀid, p.������, p.����˵��, s.���������," & vbNewLine & _
            "       To_Char(p.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ����id, Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select ҩƷid,Min(ԭ�ɱ���) As ԭ�ɱ���, Min(�³ɱ���) As �³ɱ���, min(����) as ����,���ۻ��ܺ� From �ɱ��۵�����Ϣ Where ���ۻ��ܺ� = [1] Group By ҩƷid,���ۻ��ܺ�) M, �շѼ�Ŀ P, ���ۻ��ܼ�¼ A, (Select ҩƷid From ҩƷ��� where ����=1) K, �շ���Ŀ���� B, �������� S, �շ���ĿĿ¼ I" & vbNewLine & _
            "Where m.���ۻ��ܺ� = a.���ۺ� and m.ҩƷid=i.id And p.���ۻ��ܺ� = a.���ۺ� And p.�շ�ϸĿid = k.ҩƷid(+) And p.�շ�ϸĿid = b.�շ�ϸĿid(+) And p.�շ�ϸĿid = s.����id And" & vbNewLine & _
            "      s.����id = i.Id And a.���ۺ� =[1] And b.����(+) = 3 And a.���� = 1 " & vbNewLine & _
            GetPriceClassString("P") & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & " order by ����id"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ۻ��ܺ�)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "�õ��ۼ�¼�Ѿ���ɾ���ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If

    With vsfPrice
        .Rows = 2
        rsTemp.MoveFirst
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp!����ID <> lngDrugID Then
                Select Case mintUnit
                    Case 0
                        db��װϵ�� = 1
                        strUnit = rsTemp!��λ
                    Case 1
                        db��װϵ�� = rsTemp!����ϵ��
                        strUnit = rsTemp!��װ��λ
                End Select

                lngDrugID = rsTemp!����ID

                If mintMethod = 0 Or mintMethod = 2 Then
                    .TextMatrix(.Rows - 1, menuPriceCol.ԭ��id) = IIf(IsNull(rsTemp!ԭ��id), "", rsTemp!ԭ��id)
                End If
                .TextMatrix(.Rows - 1, menuPriceCol.����ID) = lngDrugID

                .TextMatrix(.Rows - 1, menuPriceCol.Ʒ��) = "[" & rsTemp!���� & "]" & IIf(IsNull(rsTemp!��Ʒ��), rsTemp!ͨ����, rsTemp!��Ʒ��)
                .TextMatrix(.Rows - 1, menuPriceCol.���) = rsTemp!���
                .TextMatrix(.Rows - 1, menuPriceCol.�Ƿ���) = rsTemp!�Ƿ���
                
                If mintMethod = 1 Or mintMethod = 2 Then
                    gstrSQL = "select min(����) as ���� from �ɱ��۵�����Ϣ where ���ۻ��ܺ�=[1] and ҩƷid=[2]"
                    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "���ز�ѯ", mstr���ۻ��ܺ�, lngDrugID)
                    If rs����.RecordCount > 0 Then
                        .TextMatrix(.Rows - 1, menuPriceCol.����) = IIf(IsNull(rs����!����), "", rs����!����)
                    End If
                Else
                    .TextMatrix(.Rows - 1, menuPriceCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                End If
                
                .TextMatrix(.Rows - 1, menuPriceCol.��λ) = strUnit
                .TextMatrix(.Rows - 1, menuPriceCol.��װϵ��) = db��װϵ��

                .TextMatrix(.Rows - 1, menuPriceCol.�ӳ���) = rsTemp!�ӳ���
                .TextMatrix(.Rows - 1, menuPriceCol.���������) = rsTemp!���������
                .TextMatrix(.Rows - 1, menuPriceCol.�Ƿ��п��) = rsTemp!�Ƿ��п��
                .TextMatrix(.Rows - 1, menuPriceCol.������Ŀid) = IIf(IsNull(rsTemp!������Ŀid), "", rsTemp!������Ŀid)
                .TextMatrix(.Rows - 1, menuPriceCol.ԭ�ɱ���) = Format(rsTemp!ԭ�ɱ��� * db��װϵ��, mFMT.FM_�ɱ���)
                .TextMatrix(.Rows - 1, menuPriceCol.�ֳɱ���) = Format(rsTemp!�³ɱ��� * db��װϵ��, mFMT.FM_�ɱ���)
                .TextMatrix(.Rows - 1, menuPriceCol.ԭ���ۼ�) = Format(IIf(IsNull(rsTemp!ԭ��), rsTemp!�ּ�, rsTemp!ԭ��) * db��װϵ��, mFMT.FM_���ۼ�)
                .TextMatrix(.Rows - 1, menuPriceCol.�����ۼ�) = Format(rsTemp!�ּ� * db��װϵ��, mFMT.FM_���ۼ�)
                .TextMatrix(.Rows - 1, menuPriceCol.ԭ�ɹ��޼�) = Format(rsTemp!ָ������ * db��װϵ��, mFMT.FM_�ɱ���)
                .TextMatrix(.Rows - 1, menuPriceCol.�ֲɹ��޼�) = Format(rsTemp!ָ������ * db��װϵ��, mFMT.FM_�ɱ���)
                .TextMatrix(.Rows - 1, menuPriceCol.ԭָ���ۼ�) = Format(rsTemp!ָ���ۼ� * db��װϵ��, mFMT.FM_���ۼ�)
                .TextMatrix(.Rows - 1, menuPriceCol.��ָ���ۼ�) = Format(rsTemp!ָ���ۼ� * db��װϵ��, mFMT.FM_���ۼ�)

                txtValuer.Text = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                txtSummary.Text = IIf(IsNull(rsTemp!����˵��), "", rsTemp!����˵��)
                If mintModal = 1 Then
                    Me.dtpRunDate.MinDate = CDate(rsTemp!ִ������)
                End If
                If IsNull(rsTemp!ִ������) Then
                    StrToday = Format(sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
                Else
                    StrToday = Format(rsTemp!ִ������, "yyyy-MM-dd hh:mm:ss")
                End If
                Me.dtpRunDate.Value = CDate(StrToday)

                .Rows = .Rows + 1
                Call setColEdit
                .RowHeight(.Rows - 1) = mconlngRowHeight
            End If
            rsTemp.MoveNext
        Next
        Call GetDrugStore(Val(.TextMatrix(1, menuPriceCol.����ID)), 1)
    End With

    Exit Sub
ErrHandle:
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

    '��������
    On Error GoTo ErrHandle
    If strInput <> txtFind.Tag Then
        '��ʾ�µĲ���
        txtFind.Tag = strInput

        gstrSQL = "Select Distinct A.Id,'[' || A.���� || ']' As ҩƷ����, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                  "Where (A.վ�� = [3] Or A.վ�� is Null) And A.Id =B.�շ�ϸĿid And A.���='4' " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] ) " & _
                  "Order By ҩƷ���� "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "ȡƥ��Ĳ���id", strInput & "%", "%" & strInput & "%", gstrNodeNo)

        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If

    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    For n = 1 To mrsFindName.RecordCount
        '��������ˣ��򷵻ص�1����¼
        If mrsFindName.EOF Then mrsFindName.MoveFirst

        strҩ�� = mrsFindName!ҩƷ���� & IIf(IsNull(mrsFindName!��Ʒ��), mrsFindName!ͨ����, mrsFindName!��Ʒ��)

        For lngRow = 1 To vsfPrice.Rows - 1
            lngFindRow = vsfPrice.FindRow(strҩ��, lngRow, CLng(menuPriceCol.Ʒ��), True, True)
            If lngFindRow > 0 Then
                vsfPrice.Select lngFindRow, 1, lngFindRow, vsfPrice.Cols - 1
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If vsfPrice.Height + Y <= 800 Then Exit Sub
    If TabCtlDetails.Height - Y <= 1000 Then Exit Sub
    picSplit.Move 0, picSplit.Top + Y
    vsfPrice.Move 0, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, vsfPrice.Height + Y

    With TabCtlDetails
        .Top = picSplit.Top + picSplit.Height + 5
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = TabCtlDetails.Height - Y
    End With
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
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

    On Error GoTo ErrHandle
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
            " And ĩ��=1 And substr(����,5,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strTmp & "%", IIf(gstrMatchMethod = "0", "%", "") & strTmp & "%")

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
ErrHandle:
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
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPay_EnterCell()
    With vsfPrice
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
            If .Col = menuPayCol.Ʒ�� Then
                .Col = menuPayCol.��Ӧ��
            ElseIf .Col = menuPayCol.��Ӧ�� Then
                .Col = menuPayCol.��Ʊ��
            ElseIf .Col = menuPayCol.��Ʊ�� Then
                .Col = menuPayCol.��Ʊ����
            ElseIf .Col = menuPayCol.��Ʊ���� Then
                .Col = menuPayCol.��Ʊ���
            ElseIf .Col = menuPayCol.��Ʊ��� And .Row <> .Rows - 1 Then
                .Col = menuPayCol.Ʒ��
                .Row = .Row + 1
            End If
        End If
    End With
End Sub

Private Sub vsfPay_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPay
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If

        End With
    End If
End Sub

Private Sub vsfPay_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer

    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfPay
            If Col = menuPayCol.��Ʊ��� Then
                strKey = .EditText
                intDigit = Len(Mid(mFMT.FM_���, InStr(1, mFMT.FM_���, ".") + 1))
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strKey) Then Exit Sub
                    If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
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
    Dim strKey As String

    With vsfPay
        If Col = menuPayCol.��Ʊ���� Then
            strKey = .EditText
            If strKey <> "" Then
                If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                    strKey = TranNumToDate(strKey)
                    If strKey = "" Then
                        MsgBox "�Բ��𣬷�Ʊ���ڱ���Ϊ������,��ʽ(20000101����2000-01-01)��", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strKey
                    .TextMatrix(Row, menuPayCol.��Ʊ����) = .EditText
                End If

                If Not IsDate(strKey) Then
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

        If chkAppAllColumn.Value = 1 Then
            Call AutoCalc���п��۸�
        End If
    End With
End Sub

Private Sub AutoCalc���п��۸�()
    '-----------------------------------------------------------------------------------------------------------
    '����:�Զ��������п��ļ۸�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl�ֳɱ��� As Double, dbl�ּ� As Double, dbl�ӳ��� As Double, dbl�ɱ���� As Double, dbl��۵����� As Double, dbl������ As Double
    Dim lng����ID As Long, bln�ⷿ���� As Boolean, lng��Ӧ��ID As Long, lngTemp As Long, i As Long
    Dim blnHaveData As Boolean, lngStep As Long, lngSteps As Long
    Dim intCol As Integer
    Dim cllData As New Collection

    err = 0: On Error GoTo ErrHand:

    '��Ϊ���ڰ�װ�������⣬��ˣ�Ŀǰ����С��λ�������õ���
    dbl�ֳɱ��� = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.�ֳɱ���))
    dbl�ּ� = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.�����ۼ�))

    With vsfStore
        For lngRow = 1 To .Rows - 1
            If vsfPrice.Col = menuPriceCol.�ֳɱ��� Then
                .TextMatrix(lngRow, menuStoreCol.�ֲɹ���) = dbl�ֳɱ���
                '�ӳ���=�����ۼ�/�ֳɱ���-1
                If dbl�ֳɱ��� <> 0 Then
                    dbl�ӳ��� = Round(Val(.TextMatrix(lngRow, menuStoreCol.�����ۼ�)) / dbl�ֳɱ��� - 1, 7)
                Else
                    dbl�ӳ��� = 0

                End If
                '��۵�����=(�ֳɱ���-ԭ�ɱ���)
                dbl�ɱ���� = Round((Val(.TextMatrix(lngRow, menuStoreCol.ԭ�ɹ���)) - dbl�ֳɱ���), 7)
            ElseIf vsfPrice.Col = menuPriceCol.�����ۼ� Then
                .TextMatrix(lngRow, menuStoreCol.�����ۼ�) = dbl�ּ�
                '�ּ۷����ı�ʱ,��Ҫ���¸��ݼӳ��ʼ�����ص��ֳɱ���
                dbl�ӳ��� = Round(Val(.TextMatrix(lngRow, menuStoreCol.�ӳ���)) / 100, 7)
                If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
                '�ֳɱ���=�����ۼ�/(1+�ӳ���)
                dbl�ֳɱ��� = Round(dbl�ּ� / (1 + dbl�ӳ���), 7)
                '��۵�����=(�ֳɱ���-ԭ�ɱ���)
                dbl�ɱ���� = (dbl�ֳɱ��� - Val(.TextMatrix(lngRow, menuStoreCol.ԭ�ɹ���)))

                '������=����*(�ּ�-ԭ��)
                dbl������ = (dbl�ּ� - Val(.TextMatrix(lngRow, menuStoreCol.ԭ���ۼ�))) * Val(.TextMatrix(lngRow, menuStoreCol.����))
                .TextMatrix(lngRow, menuStoreCol.�������) = Format(dbl������, mFMT.FM_���)
            End If

            lng����ID = Val(.TextMatrix(lngRow, menuStoreCol.����ID))
            lng��Ӧ��ID = Val(.TextMatrix(lngRow, menuStoreCol.��Ӧ��ID))

            If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
            .TextMatrix(lngRow, menuStoreCol.�ӳ���) = Format(dbl�ӳ��� * 100, GFM_VBJCL)
            dbl�ɱ���� = (Val(.TextMatrix(lngRow, menuStoreCol.ԭ�ɹ���)) - dbl�ֳɱ���)
             '��۵�����=(�ֳɱ���-ԭ�ɱ���)*����
             dbl��۵����� = Round(dbl�ɱ���� * Val(.TextMatrix(lngRow, menuStoreCol.����)), 7)
            .TextMatrix(lngRow, menuStoreCol.��۲�) = Format(dbl��۵�����, mFMT.FM_���)
            lngTemp = Val(.TextMatrix(lngRow, menuStoreCol.����ID))
            lng��Ӧ��ID = Val(.TextMatrix(lngRow, menuStoreCol.��Ӧ��ID))

            If lng��Ӧ��ID <> 0 Then
                err = 0: On Error Resume Next
                cllData.Add Array(lngTemp, lng��Ӧ��ID, dbl��۵�����, .TextMatrix(lngRow, menuStoreCol.��Ӧ��ID), .TextMatrix(lngRow, menuStoreCol.����ID), .TextMatrix(lngRow, menuStoreCol.���), .TextMatrix(lngRow, menuStoreCol.����)), "K" & lng��Ӧ��ID & "_" & lngTemp
                If err <> 0 Then
                    '�ۼƲ�۵�����
                    dbl��۵����� = Val(cllData("K" & lng��Ӧ��ID & "_" & lngTemp)(2)) + dbl��۵�����
                    cllData.Remove "K" & lng��Ӧ��ID & "_" & lngTemp
                     err = 0: On Error GoTo ErrHand:
                    cllData.Add Array(lngTemp, lng��Ӧ��ID, dbl��۵�����, .TextMatrix(lngRow, menuStoreCol.��Ӧ��ID), .TextMatrix(lngRow, menuStoreCol.����ID), .TextMatrix(lngRow, menuStoreCol.���), .TextMatrix(lngRow, menuStoreCol.����)), "K" & lng��Ӧ��ID & "_" & lngTemp

                End If
                On Error GoTo ErrHand:
            End If
        Next

        If chkAutoPay.Value = 1 Then
            '��Ҫ�Զ�������ص�Ӧ���䶯��¼
            For i = 1 To cllData.Count
                With vsfPay
                    blnHaveData = False
                    For lngRow = 1 To .Rows - 1
                        lngTemp = Val(.TextMatrix(lngRow, menuPayCol.����ID))
                        lng��Ӧ��ID = Val(.TextMatrix(lngRow, menuPayCol.��Ӧ��ID))
                        If lngTemp = Val(cllData(i)(0)) _
                            And lng��Ӧ��ID = Val(cllData(i)(1)) Then
                            '���ļ���Ӧ����ͬ,�����ص�ֵ
                            .TextMatrix(lngRow, menuPayCol.��Ʊ���) = Format(Val(cllData(i)(2)), mFMT.FM_���)
                             blnHaveData = True
                        End If
                    Next
                    If blnHaveData = False Then
                        '��Ҫ���Ӹ��Ӧ�̵�����
                        If Val(.TextMatrix(.Rows - 1, menuStoreCol.����ID)) <> 0 Then
                            .Rows = .Rows + 1
                        End If
                        lngRow = .Rows - 1
                        .TextMatrix(lngRow, menuPayCol.��Ӧ��ID) = cllData(i)(3)
                        .TextMatrix(lngRow, menuPayCol.����ID) = cllData(i)(0)
                        .TextMatrix(lngRow, menuPayCol.���) = cllData(i)(5)
                        .TextMatrix(lngRow, menuPayCol.����) = cllData(i)(6)
                        .TextMatrix(lngRow, menuPayCol.��Ʊ���) = Format(Val(cllData(i)(2)), mFMT.FM_���)
                    End If
                End With
            Next
        End If
    End With

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
'    Call SetRowHidden(Val(vsfPrice.TextMatrix(NewRow, menuPriceCol.����id)))
End Sub

Private Sub SetRowHidden(ByVal lngDrugID As Long)
    '���ܣ��е���ʾ������
    '����������id
    Dim intRow As Integer

    If lngDrugID = 0 Then Exit Sub
    With vsfStore
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, menuStoreCol.����ID)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With

    With vsfPay
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, menuPayCol.����ID)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With
End Sub

'Private Sub vsfPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    With vsfPrice
'        mlngPrice = Val(.TextMatrix(Row, Col))
'        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
'            Cancel = True
'        End If
'    End With
'End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim mrsReturn As Recordset

    mblnClick = True
    mblnUpdateAdd = True

    On Error GoTo ErrHandle
    Set mrsReturn = SelectStuff("")
    If mrsReturn Is Nothing Then Exit Sub
    If mrsReturn.RecordCount = 0 Then Exit Sub

    Call GetDrugPirce(mrsReturn, Row)
    mblnUpdateAdd = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SelectStuff(ByVal strKey As String) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ������������
    '����:strKey-��ѡ�������
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/17
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim rsDrugInfo As ADODB.Recordset
    Dim blnCancel As Boolean, i As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim intϵ�� As Integer
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    err = 0: On Error GoTo ErrHand:
    Call CalcPosition(sngX, sngY, vsfPrice)

    Set rsDrugInfo = New ADODB.Recordset
    With rsDrugInfo
        If .State = 1 Then .Close
        .Fields.Append "id", adDouble, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "ͨ����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "���㵥λ", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����ϵ��", adDouble, 40, adFldIsNullable
        .Fields.Append "��װ��λ", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ʱ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "�ɱ���", adDouble, 40, adFldIsNullable
        .Fields.Append "ָ��������", adDouble, 40, adFldIsNullable
        .Fields.Append "ָ�����ۼ�", adDouble, 40, adFldIsNullable
        .Fields.Append "��������", adDouble, 1, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
        gstrSQL = "" & _
            "   Select distinct I.ID,I.����,b.���� As ��Ʒ��, i.���� As ͨ����,I.���,I.����,I.���㵥λ,P.����ϵ��,P.��װ��λ," & _
            "         decode(I.�Ƿ���,1,'ʱ��','����') ����,Decode(i.�Ƿ���, 0, '����', 1, 'ʱ��') As ʱ��," & _
            "         to_char(p.�ɱ���,'9999999999990.9999999') as �ɱ���," & _
            "         to_char(p.ָ��������,'9999999999990.9999999') ָ��������," & _
            "         to_char(p.ָ�����ۼ�,'9999999999990.9999999') ָ�����ۼ�," & _
            "          P.��������" & _
            "   From �շ���ĿĿ¼ I,�շ���Ŀ���� N,�������� P,�շ���Ŀ���� B" & _
            "   Where I.ID=N.�շ�ϸĿID And I.ID=P.����ID  and i.Id = b.�շ�ϸĿid(+) and b.����(+) = 3 And i.��� = '4' " & _
            "       and (I.���� like [1] or N.���� Like [1] or N.���� Like [1])" & _
            "       and (I.����ʱ�� Is Null Or I.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
     Else
        gstrSQL = "" & _
            "   Select distinct  I.ID,I.����,b.���� As ��Ʒ��, i.���� As ͨ����,I.���,I.����,I.���㵥λ,P.����ϵ��,P.��װ��λ, " & _
            "           decode(I.�Ƿ���,1,'ʱ��','����') ����,Decode(i.�Ƿ���, 0, '����', 1, 'ʱ��') As ʱ��," & _
            "           to_char(p.�ɱ���,'9999999999990.9999999') as �ɱ���," & _
            "           to_char(p.ָ��������,'9999999999990.9999999') ָ��������," & _
            "           to_char(p.ָ�����ۼ�,'9999999999990.9999999') ָ�����ۼ�," & _
            "           P.��������" & _
            "   From �շ���ĿĿ¼ I,�������� P,�շ���Ŀ���� B" & _
            "   Where I.ID=P.����ID and i.Id = b.�շ�ϸĿid(+) And" & _
            "   b.����(+) = 3 And i.��� = '4'" & _
            "           and (I.����ʱ�� Is Null Or I.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"

    End If

    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "��������ѡ��", False, "", "", False, False, True, sngX, sngY - vsfPrice.CellHeight, vsfPrice.CellHeight, blnCancel, False, False, strKey)
    If blnCancel = True Then Exit Function

    If rsTemp Is Nothing Then
        ShowMsgBox "������ָ������������,����!"
        Exit Function
    End If

    With rsDrugInfo
        .AddNew
        !Id = rsTemp!Id
        !���� = rsTemp!����
        !��Ʒ�� = rsTemp!��Ʒ��
        !ͨ���� = rsTemp!ͨ����
        !��� = rsTemp!���
        !���� = rsTemp!����
        !���㵥λ = rsTemp!���㵥λ
        !����ϵ�� = rsTemp!����ϵ��
        !��װ��λ = rsTemp!��װ��λ
        !ʱ�� = rsTemp!ʱ��
        !�ɱ��� = rsTemp!�ɱ���
        !ָ�������� = rsTemp!ָ��������
        !ָ�����ۼ� = rsTemp!ָ�����ۼ�
        !�������� = rsTemp!��������

        .Update
    End With

    Set SelectStuff = rsDrugInfo

    Exit Function
ErrHand:
    vsfPrice.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub GetDrugPirce(ByVal rsReturn As ADODB.Recordset, ByVal Row As Integer)
    '������ȡҩƷ��Ϣ
    Dim rsTemp As Recordset
    Dim lngDrugID As Long
    Dim lngRow As Long
    Dim i As Long
    Dim intCurrentPrice As Integer '�Ƿ���ʱ��
    Dim strUnit As String
    Dim db��װϵ�� As Double
    Dim dbl���� As Double

    On Error GoTo ErrHandle

    mlngOldStuffID = Val(vsfPrice.TextMatrix(Row, menuPriceCol.����ID))
    Set rsReturn = CheckDoubleDrug(rsReturn)
    If rsReturn.RecordCount = 0 Then Exit Sub

    rsReturn.MoveFirst
    For i = 0 To rsReturn.RecordCount - 1
        With vsfPrice
            lngDrugID = rsReturn!Id

            '����Ƿ����δִ�еļ۸�
            If checkNotExecutePrice(lngDrugID) = True Then Exit Sub

            Select Case mintUnit
                Case 0  'ɢװ��λ
                    db��װϵ�� = 1
                    strUnit = rsReturn!���㵥λ
                Case 1  '��װ��λ
                    db��װϵ�� = rsReturn!����ϵ��
                    strUnit = rsReturn!��װ��λ
            End Select

            .TextMatrix(Row, menuPriceCol.����ID) = lngDrugID

            .EditText = "[" & rsReturn!���� & "]" & IIf(IsNull(rsReturn!��Ʒ��) Or rsReturn!��Ʒ�� = "", rsReturn!ͨ����, rsReturn!��Ʒ��)
            .TextMatrix(Row, menuPriceCol.Ʒ��) = IIf(.EditText = "", "[" & rsReturn!���� & "]" & IIf(IsNull(rsReturn!��Ʒ��) Or rsReturn!��Ʒ�� = "", rsReturn!ͨ����, rsReturn!��Ʒ��), .EditText)

            .TextMatrix(Row, menuPriceCol.���) = IIf(IsNull(rsReturn!���), "", rsReturn!���)
            .TextMatrix(Row, menuPriceCol.�Ƿ���) = IIf(rsReturn!ʱ�� = "ʱ��", 1, 0)
            intCurrentPrice = IIf(rsReturn!ʱ�� = "ʱ��", 1, 0)
            .TextMatrix(Row, menuPriceCol.����) = IIf(IsNull(rsReturn!����), "", rsReturn!����)
            .TextMatrix(Row, menuPriceCol.��λ) = strUnit
            .TextMatrix(Row, menuPriceCol.��װϵ��) = db��װϵ��
            .TextMatrix(Row, menuPriceCol.�Ƿ��������) = zlStr.nvl(rsReturn!��������)
            .TextMatrix(Row, menuPriceCol.�ֳɱ���) = Format(Val(zlStr.nvl(rsReturn!�ɱ���)) * db��װϵ��, mFMT.FM_�ɱ���)
            .TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�) = Format(Val(zlStr.nvl(rsReturn!ָ��������)) * db��װϵ��, mFMT.FM_�ɱ���)
            .TextMatrix(Row, menuPriceCol.�ֲɹ��޼�) = .TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�)
            .TextMatrix(Row, menuPriceCol.ԭָ���ۼ�) = Format(Val(zlStr.nvl(rsReturn!ָ�����ۼ�)) * db��װϵ��, mFMT.FM_���ۼ�)
            .TextMatrix(Row, menuPriceCol.��ָ���ۼ�) = .TextMatrix(Row, menuPriceCol.ԭָ���ۼ�)

            gstrSQL = "select ҩƷid from ҩƷ��� where ҩƷid=[1] and ����=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����", lngDrugID)
            If rsTemp.RecordCount = 0 Then
                .TextMatrix(Row, menuPriceCol.�Ƿ��п��) = 0
            Else
                .TextMatrix(Row, menuPriceCol.�Ƿ��п��) = 1
            End If

            If intCurrentPrice = 0 Then '��������
                '��ʾ����ҩƷ���ۣ��ɱ���ȡƽ���۸��ۼ�ȡ�շѼ�Ŀ�ּ�
                gstrSQL = "Select b.Id, Decode(Nvl(k.�������, 0), 0, a.�ɱ���, (k.����� - k.�����) / k.�������) As �ɱ���, a.ָ��������, a.ָ�����ۼ�, b.�ּ�, a.���������," & vbNewLine & _
                            "       nvl(a.�ӳ���,0) / 100 As �ӳ���, b.������Ŀid" & vbNewLine & _
                            "From �������� A, �շѼ�Ŀ B," & vbNewLine & _
                            "     (Select Sum(ʵ�ʽ��) �����, Sum(ʵ�ʲ��) As �����, Sum(ʵ������) �������" & vbNewLine & _
                            "       From ҩƷ���" & vbNewLine & _
                            "       Where ���� = 1 And ҩƷid = [1]) K" & vbNewLine & _
                            "Where a.����id = b.�շ�ϸĿid And a.����id = [1] And (b.��ֹ���� Is Null Or b.��ֹ���� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                            GetPriceClassString("B")
            Else 'ʱ������
                '��ʾʱ�����ĵ��ۣ�ȡ�����/���������Ϊ��۸�
                gstrSQL = "" & _
                        "   Select  P.id,Decode(Nvl(K.�������,0),0,P.�ּ�,K.�����/Nvl(K.�������,1)) �ּ�,nvl(m.�ӳ���,0) / 100 as �ӳ���," & _
                        "           P.ִ������,P.������Ŀid,I.���� as ��������, " & IIf(mintUnit = 0, "1", " Nvl(M.����ϵ��,1)") & " as  ϵ��,decode(nvl(k.�������,0),0,m.�ɱ���,(k.�����-k.�����)/k.�������) as �ɱ���,m.��������,m.ָ��������,m.ָ�����ۼ�,m.���������" & _
                        "   From �շѼ�Ŀ P,������Ŀ I,�������� M," & _
                        "       (   Select Sum(ʵ�ʽ��) �����,Sum(ʵ������) �������,Sum(ʵ�ʲ��) As �����" & _
                        "           From ҩƷ��� " & _
                        "           Where  ����=1 and ҩƷID=[1] " & _
                        "        ) K" & _
                        " where p.�շ�ϸĿid=M.����id and P.������Ŀid=I.id and P.�շ�ϸĿid=[1] " & _
                        "       and (P.��ֹ���� is null or P.��ֹ����=to_date('3000-01-01','YYYY-MM-DD'))" & _
                        GetPriceClassString("P")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯҩƷ", lngDrugID)
            If rsTemp.RecordCount > 0 Then

                .TextMatrix(Row, menuPriceCol.ԭ��id) = rsTemp!Id
                .TextMatrix(Row, menuPriceCol.������Ŀid) = IIf(IsNull(rsTemp!������Ŀid), 0, rsTemp!������Ŀid)
                .TextMatrix(Row, menuPriceCol.�ӳ���) = GetFormat(IIf(IsNull(rsTemp!�ӳ���), 0, rsTemp!�ӳ���), 2)
                .TextMatrix(Row, menuPriceCol.���������) = IIf(IsNull(rsTemp!���������), 0, rsTemp!���������)
                .TextMatrix(Row, menuPriceCol.ԭ�ɱ���) = Format(IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���) * db��װϵ��, mFMT.FM_�ɱ���)
                .TextMatrix(Row, menuPriceCol.�ֳɱ���) = Format(IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���) * db��װϵ��, mFMT.FM_�ɱ���)
                .TextMatrix(Row, menuPriceCol.ԭ���ۼ�) = Format(IIf(IsNull(rsTemp!�ּ�), 0, rsTemp!�ּ�) * db��װϵ��, mFMT.FM_���ۼ�)
                If mstr������ = "" Or mint���� = 1 Then
                    .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format(IIf(IsNull(rsTemp!�ּ�), 0, rsTemp!�ּ�) * db��װϵ��, mFMT.FM_���ۼ�)
                Else
                    Select Case mintType
                        Case 1      '���ݳɱ��ۼӳ�
                            dbl���� = 1 + Val(mdbl����) / 100
                            .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format(Val(zlStr.nvl(rsTemp!�ɱ���)) * dbl���� * db��װϵ��, mFMT.FM_���ۼ�)
                        Case 2      '�������ۼ۰�����
                            dbl���� = 1 + Val(mdbl����) / 100
                            .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format(Val(zlStr.nvl(rsTemp!�ּ�)) * dbl���� * db��װϵ��, mFMT.FM_���ۼ�)
                        Case 3      '�������ۼ۰��̶����Ӽ�
                            dbl���� = Val(mdbl����)
                            .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format((Val(zlStr.nvl(rsTemp!�ּ�)) * db��װϵ��) + dbl����, mFMT.FM_���ۼ�)
                    End Select
                End If

                If Val(.TextMatrix(Row, menuPriceCol.�����ۼ�)) > Val(.TextMatrix(Row, menuPriceCol.��ָ���ۼ�)) And Val(.TextMatrix(Row, menuPriceCol.��ָ���ۼ�)) <> 0 Then
                    .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format(Val(.TextMatrix(Row, menuPriceCol.��ָ���ۼ�)), mFMT.FM_���ۼ�)
                End If
            Else
                .TextMatrix(Row, menuPriceCol.ԭ��id) = 0
                If Row > 1 Then
                    .TextMatrix(Row, menuPriceCol.������Ŀid) = .TextMatrix(Row - 1, menuPriceCol.������Ŀid)
                End If
                .TextMatrix(Row, menuPriceCol.ԭ���ۼ�) = Format(0, mFMT.FM_���ۼ�)
                .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format(0, mFMT.FM_���ۼ�)
                .TextMatrix(Row, menuPriceCol.ԭ�ɱ���) = Format(0, mFMT.FM_�ɱ���)
                .TextMatrix(Row, menuPriceCol.�ֳɱ���) = Format(0, mFMT.FM_�ɱ���)

                If mstr������ = "" Or mint���� = 1 Then
                    .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format(0, mFMT.FM_���ۼ�)
                Else
                    Select Case mintType
                        Case 1      '���ݳɱ��ۼӳ�
                            dbl���� = 1 + Val(mdbl����) / 100
                            .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format(0 * dbl���� * db��װϵ��, mFMT.FM_���ۼ�)
                        Case 2      '�������ۼ۰�����
                            dbl���� = 1 + Val(mdbl����) / 100
                            .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format(0 * dbl���� * db��װϵ��, mFMT.FM_���ۼ�)
                        Case 3      '�������ۼ۰��̶����Ӽ�
                            dbl���� = Val(mdbl����)
                            .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format(0 + dbl���� * db��װϵ��, mFMT.FM_���ۼ�)
                    End Select
                End If

                If Val(.TextMatrix(Row, menuPriceCol.�����ۼ�)) > Val(.TextMatrix(Row, menuPriceCol.��ָ���ۼ�)) And Val(.TextMatrix(Row, menuPriceCol.��ָ���ۼ�)) <> 0 Then
                    .TextMatrix(Row, menuPriceCol.�����ۼ�) = Format(Val(.TextMatrix(Row, menuPriceCol.��ָ���ۼ�)), mFMT.FM_���ۼ�)
                End If
            End If

            Call GetDrugStore(lngDrugID, Row)
            If Row = .Rows - 1 Then '���һ�в�������
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = mconlngRowHeight
                Row = Row + 1
            End If
        End With

        rsReturn.MoveNext
    Next
    Call setColEdit
    mstr������ = ""
    mdbl���� = 0
    Exit Sub
ErrHandle:
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
    Dim i As Long, n As Long
    Dim dbl��Ʊ��� As Double
    Dim strҩƷ���� As String
    Dim str��Ʊ As String
    Dim str��Ʊ���� As String
    Dim rsPirce As ADODB.Recordset
    Dim rsCost As ADODB.Recordset
    Dim dbl��װ���� As Double
    Dim bln��ͬҩƷ As Boolean
    Dim lng����ID As Long
    Dim str��λ As String
    Dim dbl���� As Double

    '���ܣ�Ϊ����б��������
    '����������id

    On Error GoTo ErrHandle

    '�ȼ���Ƿ����ظ������ݣ�����о���������ظ�������
    With vsfStore
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuStoreCol.����ID)) = mlngOldStuffID And mlngOldStuffID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    With vsfPay
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuPayCol.����ID)) = mlngOldStuffID And mlngOldStuffID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
        gstrSQL = "Select s.�ⷿid,s.ҩƷid as ����id, d.���� As �ⷿ, '[' || m.���� || ']' || m.���� As ҩƷ, m.���, m.����, m.���㵥λ �ۼ۵�λ, p.��װ��λ, s.�ϴ����� As ����, s.ʵ������ As ����," & vbNewLine & _
                    "       s.����, Nvl(m.�Ƿ���, 0) ���, m.Id," & vbNewLine & _
                    "       Decode(Nvl(m.�Ƿ���, 0), 0, e.�ּ�, Decode(s.���ۼ�,null,Decode(Nvl(s.ʵ������, 0), 0, e.�ּ�, s.ʵ�ʽ�� / s.ʵ������),s.���ۼ�)) As ʱ���ۼ�, p.ָ������� As �����,nvl(p.�ӳ���,0) as �ӳ���," & vbNewLine & _
                    "       Decode(s.ƽ���ɱ���, null, p.�ɱ���, s.ƽ���ɱ���) As �ɱ���, s.�ϴι�Ӧ��id, n.���� As ��Ӧ��, s.Ч��, s.�ϴβ��� As ����" & vbNewLine & _
                    "From ҩƷ��� S, ���ű� D, �շ���ĿĿ¼ M, �������� P, ��Ӧ�� N, �շѼ�Ŀ E" & vbNewLine & _
                    "Where d.Id = s.�ⷿid And s.ҩƷid = m.Id And m.Id = p.����id And Nvl(s.�ϴι�Ӧ��id, 0) = n.Id(+) And m.Id = e.�շ�ϸĿid And" & vbNewLine & _
                    "      s.���� = 1 And s.ҩƷid = [1] And Sysdate Between e.ִ������ And e.��ֹ����  " & vbNewLine & _
                    GetPriceClassString("E") & "Order By �ⷿ, s.�ϴ�����"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lngDrugID)

        If mlng��Ӧ��ID > 0 Then
            rsTemp.Filter = "�ϴι�Ӧ��ID=" & mlng��Ӧ��ID
        End If
    Else '�޸ģ�����
        If mintModal = 2 Then '����
            If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                gstrSQL = "select (sysdate-ִ������ ) as �Ƿ�ִ�� from ���ۻ��ܼ�¼ where ���ۺ�=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�ִ��", txtNO.Text)
                If rsTemp!�Ƿ�ִ�� > 0 Then
                    gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.ҩƷid as ����id, b.��ҩ��λid As �ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & vbNewLine & _
                        "                b.�³ɱ���, b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���, b.����, b.����, b.����, e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.��װ��λ As ҩ�ⵥλ," & vbNewLine & _
                        "                a.��д���� As ����, f.ָ������� As �����,nvl(f.�ӳ���,0) as �ӳ���, b.Ч��" & vbNewLine & _
                        "From ҩƷ�շ���¼ A,�ɱ��۵�����Ϣ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, �������� F" & vbNewLine & _
                        "Where a.id=b.�շ�id And a.�ⷿid = c.Id And b.��ҩ��λid = d.Id(+) And" & vbNewLine & _
                        "      a.ҩƷid = e.Id And e.Id = f.����id And b.���ۻ��ܺ� = [1] And a.���� = 18 order by �ⷿ,����"
                Else
                    gstrSQL = "Select Distinct a.�ⷿid,c.���� as �ⷿ, b.ҩƷid as ����id,a.�ϴι�Ӧ��id, '[' || e.���� || ']' ||e.���� as ҩƷ,e.���,d.���� as ��Ӧ��, b.�³ɱ���, b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���" & _
                            " ,a.�ϴβ��� as ����,a.����,a.�ϴ����� as ����,e.�Ƿ��� as ���,e.���㵥λ as �ۼ۵�λ,f.��װ��λ as ҩ�ⵥλ,a.ʵ������ as ����,f.ָ������� as �����,nvl(f.�ӳ���,0) as �ӳ���,a.Ч��" & _
                            " From ҩƷ��� A,���ű� C,��Ӧ�� D,�շ���ĿĿ¼ E,�������� F," & _
                                 " (Select Distinct ҩƷid, �ⷿid, ����, ����, Ч��, ����, ԭ�ɱ���, �³ɱ���, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ӧ����䶯, ִ������" & _
                                   " From �ɱ��۵�����Ϣ" & _
                                   " Where ���ۻ��ܺ� = [1]) B" & _
                            " Where a.ҩƷid = b.ҩƷid And a.�ⷿid = b.�ⷿid and nvl(a.����,0)=nvl(b.����,0) and a.�ⷿid=c.id and a.�ϴι�Ӧ��id=d.id(+) and a.ҩƷid=e.id and e.id=f.����id and a.����=1 order by �ⷿ,����"
                End If

            ElseIf cboPriceMethod.Text = "�����ۼ�" Then
                gstrSQL = "select (sysdate-ִ������ ) as �Ƿ�ִ�� from ���ۻ��ܼ�¼ where ���ۺ�=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�ִ��", txtNO.Text)
                If rsTemp!�Ƿ�ִ�� > 0 Then
                    gstrSQL = "Select Distinct a.����, a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ����id, a.��ҩ��λid As �ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���," & vbNewLine & _
                        "                d.���� As ��Ӧ��, f.�ɱ��� As �³ɱ���, f.�ɱ��� As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.����, a.����, a.����, e.�Ƿ��� As ���," & vbNewLine & _
                        "                e.���㵥λ As �ۼ۵�λ, f.��װ��λ As ҩ�ⵥλ, a.��д���� As ����, f.ָ������� As �����, nvl(f.�ӳ���,0) as �ӳ���,a.Ч��" & vbNewLine & _
                        "From ҩƷ�շ���¼ A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, �������� F" & vbNewLine & _
                        "Where a.�۸�id = b.Id And a.�ⷿid = c.Id And a.��ҩ��λid = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.����id And" & vbNewLine & _
                        "      b.���ۻ��ܺ� = [1] And ���� = 13 " & GetPriceClassString("B") & "order by �ⷿ,����"
                Else
                    gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ����id, a.�ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & _
                                            " a.ƽ���ɱ��� As �³ɱ���, a.ƽ���ɱ��� As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.�ϴβ��� As ����, a.����, a.�ϴ����� As ����," & _
                                            " e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.��װ��λ as ҩ�ⵥλ, a.ʵ������ As ����, f.ָ������� As �����, nvl(f.�ӳ���,0) as �ӳ���,a.Ч��" & _
                            " From ҩƷ��� A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, �������� F" & _
                            " Where a.ҩƷid = b.�շ�ϸĿid And a.�ⷿid = c.Id And a.�ϴι�Ӧ��id = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.����id And a.���� = 1 And" & _
                                  " b.���ۻ��ܺ� = [1]" & GetPriceClassString("B") & " order by �ⷿ,����"
                End If
            End If
        Else '�޸�
            If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                gstrSQL = "Select Distinct a.�ⷿid,c.���� as �ⷿ, b.ҩƷid as ����id,a.�ϴι�Ӧ��id, '[' || e.���� || ']' ||e.���� as ҩƷ,e.���,d.���� as ��Ӧ��, b.�³ɱ���, b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���" & _
                            " ,a.�ϴβ��� as ����,a.����,a.�ϴ����� as ����,e.�Ƿ��� as ���,e.���㵥λ as �ۼ۵�λ,f.��װ��λ as ҩ�ⵥλ,a.ʵ������ as ����,f.ָ������� as �����,nvl(f.�ӳ���,0) as �ӳ���,a.Ч��" & _
                            " From ҩƷ��� A,���ű� C,��Ӧ�� D,�շ���ĿĿ¼ E,�������� F," & _
                                 " (Select Distinct ҩƷid, �ⷿid, ����, ����, Ч��, ����, ԭ�ɱ���, �³ɱ���, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ӧ����䶯, ִ������" & _
                                   " From �ɱ��۵�����Ϣ" & _
                                   " Where ���ۻ��ܺ� = [1]) B" & _
                            " Where a.ҩƷid = b.ҩƷid And a.�ⷿid = b.�ⷿid and nvl(a.����,0)=nvl(b.����,0) and a.�ⷿid=c.id and a.�ϴι�Ӧ��id=d.id(+) and a.ҩƷid=e.id and e.id=f.����id and a.����=1 order by �ⷿ,����"
            ElseIf cboPriceMethod.Text = "�����ۼ�" Then
                gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ����id, a.�ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & _
                                            " a.ƽ���ɱ��� As �³ɱ���, a.ƽ���ɱ��� As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.�ϴβ��� As ����, a.����, a.�ϴ����� As ����," & _
                                            " e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.��װ��λ as ҩ�ⵥλ, a.ʵ������ As ����, f.ָ������� As �����, nvl(f.�ӳ���,0) as �ӳ���,a.Ч��" & _
                            " From ҩƷ��� A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, �������� F" & _
                            " Where a.ҩƷid = b.�շ�ϸĿid And a.�ⷿid = c.Id And a.�ϴι�Ӧ��id = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.����id And a.���� = 1 And" & _
                                  " b.���ۻ��ܺ� = [1] " & GetPriceClassString("B") & "order by �ⷿ,����"
            End If
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, txtNO.Text)
    End If
    
    With vsfStore
        Do While Not rsTemp.EOF
            dbl��װ���� = 0
            dbl��Ʊ��� = 0
            dblOldPrice = 0
            dblNewPrice = 0
            For i = 0 To vsfPrice.Rows - 1
                If rsTemp!����ID = vsfPrice.TextMatrix(i, menuPriceCol.����ID) Then
                    dbl��װ���� = vsfPrice.TextMatrix(i, menuPriceCol.��װϵ��)
                    dblOldPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.ԭ���ۼ�))
                    dblNewPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.�����ۼ�))
                    str��λ = vsfPrice.TextMatrix(i, menuPriceCol.��λ)
                    Exit For
                End If
            Next
        
            .Rows = .Rows + 1
            Call setColEdit
            .RowHeight(.Rows - 1) = mconlngRowHeight

            '�ӿհ��п�ʼ��������
            .TextMatrix(.Rows - 1, menuStoreCol.����ID) = rsTemp!����ID
            .TextMatrix(.Rows - 1, menuStoreCol.�ⷿ) = rsTemp!�ⷿ
            .TextMatrix(.Rows - 1, menuStoreCol.�ⷿID) = rsTemp!�ⷿID
            .TextMatrix(.Rows - 1, menuStoreCol.��Ӧ��) = zlStr.nvl(rsTemp!��Ӧ��, "")
            .TextMatrix(.Rows - 1, menuStoreCol.��Ӧ��ID) = IIf(mlng��Ӧ��ID > 0, mlng��Ӧ��ID, zlStr.nvl(rsTemp!�ϴι�Ӧ��id))
            .TextMatrix(.Rows - 1, menuStoreCol.ҩƷ) = rsTemp!ҩƷ
            strҩƷ���� = rsTemp!ҩƷ

            .TextMatrix(.Rows - 1, menuStoreCol.���) = rsTemp!���
            .TextMatrix(.Rows - 1, menuStoreCol.��λ) = str��λ
            .TextMatrix(.Rows - 1, menuStoreCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.Rows - 1, menuStoreCol.Ч��) = Format(IIf(IsNull(rsTemp!Ч��), "", rsTemp!Ч��), "YYYY-MM-DD")
            .TextMatrix(.Rows - 1, menuStoreCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.Rows - 1, menuStoreCol.����) = Format(rsTemp!���� / dbl��װ����, mFMT.FM_����)
            .TextMatrix(.Rows - 1, menuStoreCol.��װϵ��) = dbl��װ����
            .TextMatrix(.Rows - 1, menuStoreCol.����) = zlStr.nvl(rsTemp!����, 0)
            .TextMatrix(.Rows - 1, menuStoreCol.���) = rsTemp!���


            If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
                dblOldCost = IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���) * dbl��װ����

                If mdbl�ӳ��� > 0 Then
                    dbl�ӳ��� = Round(mdbl�ӳ��� / 100, 7)
                ElseIf dblOldCost > 0 Then
                    dbl�ӳ��� = Round(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice) / dblOldCost - 1, 7)
                Else
                   dbl�ӳ��� = nvl(rsTemp!�ӳ���, 0) / 100
                End If

                If 1 + dbl�ӳ��� = 0 Then
                    dblNewCost = 0
                Else
                    dblNewCost = rsTemp!ʱ���ۼ� * dbl��װ���� / (1 + dbl�ӳ���)
                End If
                If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0

                 .TextMatrix(.Rows - 1, menuStoreCol.ԭ���ۼ�) = Format(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice), mFMT.FM_���ۼ�)

                n = n + 1
                If (mblnʱ�����İ����ε��� = False Or Val(.TextMatrix(.Rows - 1, menuStoreCol.���)) = 0) _
                                                                                    And n <> 1 And mstr������ <> "" And mint���� <> 1 Then
                    .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = .TextMatrix(.Rows - 2, menuStoreCol.�����ۼ�)
                Else
                    If mstr������ = "" Or mint���� = 1 Then
                        .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = Format(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice), mFMT.FM_���ۼ�)
                    Else
                        Select Case mintType
                            Case 1      '���ݳɱ��ۼӳ�
                                dbl���� = 1 + Val(mdbl����) / 100
                                .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = Format(dblOldCost * dbl����, mFMT.FM_���ۼ�)
                            Case 2      '�������ۼ۰�����
                                dbl���� = 1 + Val(mdbl����) / 100
                                .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = Format(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl���� * dbl��װ����, dblOldPrice * dbl����), mFMT.FM_���ۼ�)
                            Case 3      '�������ۼ۰��̶����Ӽ�
                                dbl���� = Val(mdbl����)
                                .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = Format(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ���� + dbl����, dblOldPrice + dbl����), mFMT.FM_���ۼ�)
                        End Select
                    End If
                End If
                 
                 .TextMatrix(.Rows - 1, menuStoreCol.�������) = Format(rsTemp!���� / dbl��װ���� * (Val(.TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.ԭ���ۼ�))), mFMT.FM_���)
                 .TextMatrix(.Rows - 1, menuStoreCol.�ӳ���) = GetFormat(dbl�ӳ��� * 100, 2)
                 .TextMatrix(.Rows - 1, menuStoreCol.ԭ�ɹ���) = Format(dblOldCost, mFMT.FM_�ɱ���)
                 .TextMatrix(.Rows - 1, menuStoreCol.�ֲɹ���) = Format(dblNewCost, mFMT.FM_�ɱ���)
                 .TextMatrix(.Rows - 1, menuStoreCol.��۲�) = Format((Val(.TextMatrix(.Rows - 1, menuStoreCol.�ֲɹ���)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.ԭ�ɹ���))) * Val(.TextMatrix(.Rows - 1, menuStoreCol.����)), mFMT.FM_���)
                 dbl��Ʊ��� = dbl��Ʊ��� + (dblNewCost - dblOldCost) * Val(.TextMatrix(.Rows - 1, menuStoreCol.����))
                 
                 Call RefreshPayData("", dbl��Ʊ���)
            Else
                If mintModal = 2 And (cboPriceMethod.Text = "�����ۼ�" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����") Then   '����
                    gstrSQL = "Select a.�ɱ��� As ԭ��, a.���ۼ� As �ּ�" & vbNewLine & _
                        "From ҩƷ�շ���¼ A, �շѼ�Ŀ B" & vbNewLine & _
                        "Where a.�۸�id = b.Id And b.���ۻ��ܺ� = [1] And a.�ⷿid = [2] And a.ҩƷid = [3] And Nvl(a.����, 0) = [4]" & _
                        GetPriceClassString("B")
                        
                    Set rsPirce = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ۼ�", txtNO.Text, rsTemp!�ⷿID, rsTemp!����ID, zlStr.nvl(rsTemp!����, 0))
                    
                    If Not rsPirce.EOF Then
                        .TextMatrix(.Rows - 1, menuStoreCol.ԭ���ۼ�) = Format(Val(rsPirce!ԭ��) * dbl��װ����, mFMT.FM_���ۼ�)
                        .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = Format(Val(rsPirce!�ּ�) * dbl��װ����, mFMT.FM_���ۼ�)
                        .TextMatrix(.Rows - 1, menuStoreCol.�������) = Format(rsTemp!���� / dbl��װ���� * (Val(.TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.ԭ���ۼ�))), mFMT.FM_���)
                    Else
                        .TextMatrix(.Rows - 1, menuStoreCol.ԭ���ۼ�) = Format(dblOldPrice, mFMT.FM_���ۼ�)
                        .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = Format(dblNewPrice, mFMT.FM_���ۼ�)
                        .TextMatrix(.Rows - 1, menuStoreCol.�������) = Format(rsTemp!���� / dbl��װ���� * (dblNewPrice - IIf(rsTemp!��� = 1, dblNewPrice * dbl��װ����, dblOldPrice)), mFMT.FM_���)
                    End If
                    If cboPriceMethod.Text = "�����ۼ�" Then
                        gstrSQL = "Select �ɱ���" & vbNewLine & _
                                    "      From (Select ƽ���ɱ��� As �ɱ���" & vbNewLine & _
                                    "             From ҩƷ���" & vbNewLine & _
                                    "             Where ����=1 And �ⷿid = [1] And ҩƷid = [2] And nvl(����,0) = [3]" & vbNewLine & _
                                    "             Union All" & vbNewLine & _
                                    "             Select �ɱ��� From �������� Where ����id = [2])" & vbNewLine & _
                                    "      Where Rownum <= 1"

                        Set rsCost = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ɱ���", rsTemp!�ⷿID, rsTemp!����ID, nvl(rsTemp!����, 0))
                        .TextMatrix(.Rows - 1, menuStoreCol.ԭ�ɹ���) = Format(rsCost!�ɱ��� * dbl��װ����, mFMT.FM_�ɱ���)
                        .TextMatrix(.Rows - 1, menuStoreCol.�ֲɹ���) = Format(rsCost!�ɱ��� * dbl��װ����, mFMT.FM_�ɱ���)
                        .TextMatrix(.Rows - 1, menuStoreCol.��۲�) = Format(0, mFMT.FM_���)
                    Else
                        .TextMatrix(.Rows - 1, menuStoreCol.ԭ�ɹ���) = Format(rsTemp!ԭ�ɱ��� * dbl��װ����, mFMT.FM_�ɱ���)
                        .TextMatrix(.Rows - 1, menuStoreCol.�ֲɹ���) = Format(rsTemp!�³ɱ��� * dbl��װ����, mFMT.FM_�ɱ���)
                        .TextMatrix(.Rows - 1, menuStoreCol.��۲�) = Format((rsTemp!�³ɱ��� * dbl��װ���� - rsTemp!ԭ�ɱ��� * dbl��װ����) * Val(.TextMatrix(.Rows - 1, menuStoreCol.����)), mFMT.FM_���)
                    End If
                Else '�޸Ļ��߳ɱ��۵���
                    '����ֱ�Ӵ��շѼ�Ŀȡ�ּۣ�ʱ�����ȴӿ��ȡ�����û������շѼ�Ŀȡ
                    If nvl(rsTemp!���, 0) = 1 Then
                        gstrSQL = "Select Nvl(s.���ۼ�, Decode(Nvl(s.ʵ������, 0), 0, 0, Nvl(s.ʵ�ʽ��, 0) / s.ʵ������)) ʱ���ۼ�" & vbNewLine & _
                        "From ҩƷ��� S" & vbNewLine & _
                        "Where s.����=1 And s.�ⷿid = [1] And s.ҩƷid = [2] And nvl(s.����,0) = [3]"
                        
                        Set rsPirce = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ۼ�", rsTemp!�ⷿID, rsTemp!����ID, nvl(rsTemp!����, 0))
                        If rsPirce.RecordCount > 0 Then
                            If rsPirce!ʱ���ۼ� > 0 Then
                                .TextMatrix(.Rows - 1, menuStoreCol.ԭ���ۼ�) = Format(rsPirce!ʱ���ۼ� * dbl��װ����, mFMT.FM_���ۼ�)
                                .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = Format(rsPirce!ʱ���ۼ� * dbl��װ����, mFMT.FM_���ۼ�)
                            Else
                                .TextMatrix(.Rows - 1, menuStoreCol.ԭ���ۼ�) = Format(dblOldPrice, mFMT.FM_���ۼ�)
                                .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = Format(dblNewPrice, mFMT.FM_���ۼ�)
                            End If
                        Else
                            .TextMatrix(.Rows - 1, menuStoreCol.ԭ���ۼ�) = Format(dblOldPrice, mFMT.FM_���ۼ�)
                            .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = Format(dblNewPrice, mFMT.FM_���ۼ�)
                        End If
                    Else
                        .TextMatrix(.Rows - 1, menuStoreCol.ԭ���ۼ�) = Format(dblOldPrice, mFMT.FM_���ۼ�)
                        .TextMatrix(.Rows - 1, menuStoreCol.�����ۼ�) = Format(dblNewPrice, mFMT.FM_���ۼ�)
                    End If
                    .TextMatrix(.Rows - 1, menuStoreCol.�������) = Format(rsTemp!���� / dbl��װ���� * (dblNewPrice - IIf(rsTemp!��� = 1, dblNewPrice * dbl��װ����, dblOldPrice)), mFMT.FM_���)
                    .TextMatrix(.Rows - 1, menuStoreCol.ԭ�ɹ���) = Format(rsTemp!ԭ�ɱ��� * dbl��װ����, mFMT.FM_�ɱ���)
                    .TextMatrix(.Rows - 1, menuStoreCol.�ֲɹ���) = Format(rsTemp!�³ɱ��� * dbl��װ����, mFMT.FM_�ɱ���)
                    .TextMatrix(.Rows - 1, menuStoreCol.��۲�) = Format((Val(.TextMatrix(.Rows - 1, menuStoreCol.�ֲɹ���)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.ԭ�ɹ���))) * Val(.TextMatrix(.Rows - 1, menuStoreCol.����)), mFMT.FM_���)
                End If
                 
                If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    If rsTemp!ԭ�ɱ��� = 0 Then
                        dbl�ӳ��� = 0
                    Else
                        dbl�ӳ��� = Round(dblNewPrice / (rsTemp!�³ɱ��� * dbl��װ����) - 1, 7)
                    End If
                    .TextMatrix(.Rows - 1, menuStoreCol.�ӳ���) = GetFormat(dbl�ӳ��� * 100, 2)
                    .TextMatrix(.Rows - 1, menuStoreCol.ԭ�ɹ���) = Format(rsTemp!ԭ�ɱ��� * dbl��װ����, mFMT.FM_�ɱ���)
                    .TextMatrix(.Rows - 1, menuStoreCol.�ֲɹ���) = Format(rsTemp!�³ɱ��� * dbl��װ����, mFMT.FM_�ɱ���)
                    .TextMatrix(.Rows - 1, menuStoreCol.��۲�) = Format((Val(.TextMatrix(.Rows - 1, menuStoreCol.�ֲɹ���)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.ԭ�ɹ���))) * Val(.TextMatrix(.Rows - 1, menuStoreCol.����)), mFMT.FM_���)
                    dbl��Ʊ��� = dbl��Ʊ��� + (Val(.TextMatrix(.Rows - 1, menuStoreCol.�ֲɹ���)) - Val(.TextMatrix(.Rows - 1, menuStoreCol.ԭ�ɹ���))) * Val(.TextMatrix(.Rows - 1, menuStoreCol.����))
                    str��Ʊ = IIf(IsNull(rsTemp!��Ʊ��), "", rsTemp!��Ʊ��)
                    str��Ʊ���� = IIf(IsNull(rsTemp!��Ʊ����), "", rsTemp!��Ʊ����)
                    
                    Call RefreshPayData(str��Ʊ, dbl��Ʊ���)
                End If
            End If

            rsTemp.MoveNext
        Loop
    End With
    
    '�޸ĺͲ���ʱ�������б�ƽ���ɱ��ۣ��ۼ�
    'mintModal 0-���� 1-�޸� 2-����
    If mintModal = 1 Or mintModal = 2 Then
        With vsfStore
            For i = 1 To .Rows - 1
                If lng����ID <> .TextMatrix(i, menuStoreCol.����ID) Then
                    Call CaluateAverCost(Val(.TextMatrix(i, menuStoreCol.����ID)))
                    Call CaluateAverOldCost(Val(.TextMatrix(i, menuStoreCol.����ID)))
                    Call CaculateAverPirce(Val(.TextMatrix(i, menuStoreCol.����ID)))
                    Call CaculateAverOldPirce(Val(.TextMatrix(i, menuStoreCol.����ID)))
                    lng����ID = Val(.TextMatrix(i, menuStoreCol.����ID))
                End If
            Next
        End With
    End If

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RefreshPayData(ByVal str��Ʊ�� As String, ByVal str��Ʊ���� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���»�ȡӦ������䶯����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, dbl��Ʊ��� As Double
    Dim lng��Ӧ��ID As Long, lng����ID As Long, blnData As Boolean

    err = 0: On Error GoTo ErrHand:
    If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
        TabCtlDetails.Item(1).Visible = True
        chkAutoPay.Visible = True
        chkAutoPay.Value = 1
    End If
    
    If chkAutoPay.Value <> 1 Then RefreshPayData = True: Exit Function

    With vsfPay
        .Rows = 2
        .RowHeight(.Rows - 1) = mconlngRowHeight
        .Clear 1
    End With

    With vsfStore
        For i = 1 To .Rows - 1
            lng��Ӧ��ID = Val(.TextMatrix(i, menuStoreCol.��Ӧ��ID))

            lng����ID = Val(.TextMatrix(i, menuStoreCol.����ID))

            If lng��Ӧ��ID <> 0 And lng����ID <> 0 Then
                dbl��Ʊ��� = Val(.TextMatrix(i, menuStoreCol.��۲�))
'                If dbl��Ʊ��� <> 0 Then
                    '������صĹ�Ӧ���Ƿ����
                    With vsfPay
                        blnData = False
                        For j = 1 To .Rows - 1
                            If lng����ID = Val(.TextMatrix(j, menuPayCol.����ID)) And _
                               lng��Ӧ��ID = Val(.TextMatrix(j, menuPayCol.��Ӧ��ID)) Then
                                .TextMatrix(j, menuPayCol.��Ʊ���) = Format(Val(.TextMatrix(j, menuPayCol.��Ʊ���)) + dbl��Ʊ���, mFMT.FM_���)
                               blnData = True
                               Exit For
                            End If
                        Next
                        If blnData = False Then
                            'û�д˹�Ӧ�̻����,�����Ҫ��������
                            If Val(.TextMatrix(.Rows - 1, menuPayCol.��Ӧ��ID)) <> 0 Then
                                .Rows = .Rows + 1
                                .RowHeight(.Rows - 1) = mconlngRowHeight
                                Call setColEdit
                            End If
                            .TextMatrix(.Rows - 1, menuPayCol.��Ӧ��) = vsfStore.TextMatrix(i, menuStoreCol.��Ӧ��)
                            .TextMatrix(.Rows - 1, menuPayCol.��Ӧ��ID) = vsfStore.TextMatrix(i, menuStoreCol.��Ӧ��ID)
                            .TextMatrix(.Rows - 1, menuPayCol.����ID) = vsfStore.TextMatrix(i, menuStoreCol.����ID)
                            .TextMatrix(.Rows - 1, menuPayCol.Ʒ��) = vsfStore.TextMatrix(i, menuStoreCol.ҩƷ)
                            .TextMatrix(.Rows - 1, menuPayCol.���) = vsfStore.TextMatrix(i, menuStoreCol.���)
                            .TextMatrix(.Rows - 1, menuPayCol.����) = vsfStore.TextMatrix(i, menuStoreCol.����)
                            .TextMatrix(.Rows - 1, menuPayCol.��Ʊ��) = str��Ʊ��
                            If str��Ʊ���� <> "" Then
                                .TextMatrix(.Rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "YYYY-MM-DD")
                            End If
                            .TextMatrix(.Rows - 1, menuPayCol.��Ʊ���) = Format(dbl��Ʊ���, mFMT.FM_���)
                        End If
                    End With
'                End If
            End If
        Next
    End With

    RefreshPayData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsfPrice_DblClick()
    With vsfPrice
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPrice_EnterCell()
    Dim i As Integer

    With vsfPrice
        .Editable = flexEDNone
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
        If .Col = menuPriceCol.�ֳɱ��� Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuPriceCol.�ֳɱ���))
        ElseIf .Col = menuPriceCol.�����ۼ� Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuPriceCol.�����ۼ�))
        End If
    End With

    With vsfStore
        If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.����ID)) = 0 Then Exit Sub

        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.����ID)) = Val(.TextMatrix(i, menuStoreCol.����ID)) Then
                    .Select i, 0, i, .Cols - 1
                    .TopRow = i
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim lngDrugID As Long
    Dim strRow As String

    With vsfPrice
        If KeyCode = vbKeyReturn Then
            If .Col <> menuPriceCol.�����ۼ� Then
                If .Col = menuPriceCol.Ʒ�� And cboPriceMethod.Text = "�����ɱ���" Then
                    .Col = menuPriceCol.�ֳɱ���
'                    .EditCell
                ElseIf .Col = menuPriceCol.Ʒ�� And cboPriceMethod.Text = "�����ۼ�" Then
                    .Col = menuPriceCol.�����ۼ�
'                    .EditCell
                ElseIf .Col = menuPriceCol.�ֳɱ��� And cboPriceMethod.Text = "�����ɱ���" Then
                    If .Row = .Rows - 1 And Val(.TextMatrix(.Row, menuPriceCol.����ID)) <> 0 Then
                        .Rows = .Rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.Ʒ��
                        .RowHeight(.Rows - 1) = mconlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.����ID)) <> 0 Then
                        .ColComboList(menuPriceCol.Ʒ��) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.Ʒ��
                    End If
                ElseIf .Col = menuPriceCol.Ʒ�� And cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    .Col = menuPriceCol.�ֳɱ���
'                    .EditCell
                ElseIf .Col = menuPriceCol.�ֳɱ��� And cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    .Col = menuPriceCol.�����ۼ�
'                    .EditCell
                ElseIf .Col = menuPriceCol.�����ۼ� And cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.Ʒ��
                        .RowHeight(.Rows - 1) = mconlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.����ID)) <> 0 Then
                        .ColComboList(menuPriceCol.Ʒ��) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.Ʒ��
'                        .EditCell
                    End If
                Else
                    .Col = .Col + 1
'                    .EditCell
                End If
            Else
                If Val(.TextMatrix(.Row, menuPriceCol.����ID)) <> 0 And .Row = .Rows - 1 Then
                    .ColComboList(menuPriceCol.Ʒ��) = ""
                    .Rows = .Rows + 1
                    .Row = .Row + 1
                    .Col = menuPriceCol.Ʒ��
                    .RowHeight(.Rows - 1) = mconlngRowHeight
'                    .EditCell
                    Call setColEdit
                ElseIf Val(.TextMatrix(.Row, menuPriceCol.����ID)) <> 0 Then
                    .ColComboList(menuPriceCol.Ʒ��) = ""
                    .Row = .Row + 1
                    .Col = menuPriceCol.Ʒ��
'                    .EditCell
                End If
            End If
        ElseIf KeyCode = vbKeyDelete Then
            lngDrugID = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.����ID))

            If .Rows > 2 Then
                mdbl���� = 0
                mstr������ = ""
                .RemoveItem .Row
            Else
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.Row, intCol) = ""
                Next
            End If

            With vsfStore
                If lngDrugID = 0 Then Exit Sub
                For intRow = .Rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuStoreCol.����ID)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With

            With vsfPay
                If lngDrugID = 0 Then Exit Sub
                For intRow = .Rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuPayCol.����ID)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsReturn As Recordset
    Dim strKey As String

    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub

    With vsfPrice
        strKey = .EditText
        Select Case Col
        Case menuPriceCol.Ʒ��
            mblnUpdateAdd = True
            Set rsReturn = SelectStuff(strKey)
            If rsReturn Is Nothing Then Exit Sub
            If rsReturn.RecordCount = 0 Then Exit Sub
            Call GetDrugPirce(rsReturn, Row)
            mblnUpdateAdd = False
        End Select
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckDoubleDrug(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '����Ƿ����ظ���ҩƷ
    'lngDrugId ����id
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
            For j = 1 To .Rows - 1
                If Val(.TextMatrix(j, menuPriceCol.����ID)) = rsTemp!Id Then
                    strTemp = strTemp & " id <> " & rsTemp!Id & " and "
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
        MsgBox strName & "��" & intCount & "���������б����Ѿ����ڣ��Ѵ������Ĳ�����ӣ�", vbInformation, gstrSysName
    End If

    Set CheckDoubleDrug = rsTemp
End Function

Private Sub vsfPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPrice
            If .Col = menuPriceCol.Ʒ�� Then
                .Editable = flexEDKbdMouse
                Exit Sub
            End If
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer

    With vsfPrice
        strKey = .EditText
        If .Col = menuPriceCol.�ֳɱ��� Then
            mdbl�ɱ��� = Val(.TextMatrix(Row, Col))
        End If
    End With

    If Col = menuPriceCol.�ֳɱ��� Or Col = menuPriceCol.�����ۼ� Then
        If KeyAscii = vbKeyReturn Then Exit Sub
        If KeyAscii <> vbKeyBack Then
            Select Case Col
                Case menuPriceCol.�ֳɱ���
                    intDigit = Len(Mid(mFMT.FM_�ɱ���, InStr(1, mFMT.FM_�ɱ���, ".") + 1))
                Case menuPriceCol.�����ۼ�
                    intDigit = Len(Mid(mFMT.FM_���ۼ�, InStr(1, mFMT.FM_���ۼ�, ".") + 1))
            End Select

            If KeyAscii = vbKeyDelete Then
                If InStr(1, strKey, ".") > 0 Then
                    KeyAscii = 0
                End If
            ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                If vsfPrice.EditSelLength = Len(strKey) Then Exit Sub
                If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                    KeyAscii = 0
                    Exit Sub
                End If
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf Col = menuPriceCol.Ʒ�� Then
        If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsfPrice_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = menuPriceCol.Ʒ�� Then
        vsfPrice.ColComboList(menuPriceCol.Ʒ��) = "|..."
    End If
End Sub

Private Sub setColEdit()
    '���ܣ��������Ƿ�����޸�
    '�����޸ĵ�����ɫΪ��ɫ�����޸ĵ�����ɫΪ��ɫ
    Dim intCol As Integer
    Dim intRow As Integer

    With vsfPrice
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "�����ۼ�" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.Ʒ��, .Rows - 1, menuPriceCol.Ʒ��) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�����ۼ�, .Rows - 1, menuPriceCol.�����ۼ�) = mconlngCanColor
        ElseIf cboPriceMethod.Text = "�����ɱ���" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.Ʒ��, .Rows - 1, menuPriceCol.Ʒ��) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�ֳɱ���, .Rows - 1, menuPriceCol.�ֳɱ���) = mconlngCanColor
        Else
            .Cell(flexcpBackColor, 1, menuPriceCol.Ʒ��, .Rows - 1, menuPriceCol.Ʒ��) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�ֳɱ���, .Rows - 1, menuPriceCol.�ֳɱ���) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�����ۼ�, .Rows - 1, menuPriceCol.�����ۼ�) = mconlngCanColor
        End If

    End With

    With vsfStore
        If .Rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "�����ۼ�" Then
            .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = mconlngColor
        ElseIf cboPriceMethod.Text = "�����ɱ���" Then
            .Cell(flexcpBackColor, 1, menuStoreCol.�ӳ���, .Rows - 1, menuStoreCol.�ӳ���) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuStoreCol.�ֲɹ���, .Rows - 1, menuStoreCol.�ֲɹ���) = mconlngCanColor
        Else
            .Cell(flexcpBackColor, 1, menuStoreCol.�ӳ���, .Rows - 1, menuStoreCol.�ӳ���) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuStoreCol.�ֲɹ���, .Rows - 1, menuStoreCol.�ֲɹ���) = mconlngCanColor
            .Cell(flexcpBackColor, 1, menuStoreCol.�����ۼ�, .Rows - 1, menuStoreCol.�����ۼ�) = mconlngCanColor
        End If
        If .Rows > 1 Then
            For intRow = 1 To .Rows - 1
                If Val(.TextMatrix(intRow, menuStoreCol.���)) = 1 And mblnʱ�����İ����ε��� = True And mint���� <> 1 Then
                    .Cell(flexcpBackColor, intRow, menuStoreCol.�����ۼ�, intRow, menuStoreCol.�����ۼ�) = mconlngCanColor
                Else
                    .Cell(flexcpBackColor, intRow, menuStoreCol.�����ۼ�, intRow, menuStoreCol.�����ۼ�) = mconlngColor
                End If
            Next
        End If
    End With

    With vsfPay
        If .Rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = mconlngColor
        .Cell(flexcpBackColor, 1, menuPayCol.��Ʊ��, .Rows - 1, menuPayCol.��Ʊ��) = mconlngCanColor
        .Cell(flexcpBackColor, 1, menuPayCol.��Ʊ����, .Rows - 1, menuPayCol.��Ʊ����) = mconlngCanColor
        .Cell(flexcpBackColor, 1, menuPayCol.��Ʊ���, .Rows - 1, menuPayCol.��Ʊ���) = mconlngCanColor
    End With
End Sub

Private Sub vsfPrice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        vsfPrice.Editable = flexEDNone
        If vsfPrice.Col = menuPriceCol.Ʒ�� And mintModal <> 2 Then
            vsfPrice.ColComboList(menuPriceCol.Ʒ��) = "|..."
            vsfPrice.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub vsfPrice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngDrugID As Long
    Dim intRow As Integer
    Dim strKey As String

    strKey = Trim(vsfPrice.EditText)
    strKey = Replace(strKey, Chr(vbKeyReturn), "")
    strKey = Replace(strKey, Chr(10), "")

    With vsfPrice
        If .EditText = "" Then Exit Sub
        lngDrugID = Val(.TextMatrix(Row, menuPriceCol.����ID))
        If lngDrugID = 0 Then Exit Sub

        Select Case Col
            Case menuPriceCol.�ֳɱ���
                If Not IsNumeric(strKey) Then
                    Cancel = True
                    Exit Sub
                End If
                If .EditText > 9999999 Then
                    MsgBox "�ɱ��۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If mdblOldPrice = .EditText Then Exit Sub 'û�����޸�ʱ��ֱ���˳�

                If strKey <> "" Then
                    If Val(strKey) > Val(.TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�)) And Val(.TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�)) <> 0 Then
                        If MsgBox("�ɱ��۲��ܴ���ָ�����ۼۣ���" & Format(Val(.TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�)), mFMT.FM_�ɱ���) & "����������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            vsfPrice.EditText = Format(Val(strKey), mFMT.FM_�ɱ���)
                            vsfPrice.TextMatrix(Row, menuPriceCol.�ֲɹ��޼�) = Format(Val(strKey), mFMT.FM_�ɱ���)
                        End If
                    Else
                        vsfPrice.EditText = Format(Val(strKey), mFMT.FM_�ɱ���)
                    End If
                    If chkAppAllColumn.Value = 1 And mlngPrice <> vsfPrice.EditText Then
                        For intRow = 1 To .Rows - 1
                            If .TextMatrix(intRow, menuPriceCol.����ID) <> "" Then
                                .TextMatrix(intRow, menuPriceCol.�ֳɱ���) = vsfPrice.EditText
                            End If
                        Next
                    End If
                End If
                If chkAppAllColumn.Value = 0 Then
                    Call FullStoce�ɱ���(Val(.TextMatrix(Row, menuPriceCol.����ID)), vsfPrice.EditText)
                End If
            Case menuPriceCol.�����ۼ�
                If Not IsNumeric(strKey) Then
                    Cancel = True
                    Exit Sub
                End If
                If .EditText > 9999999 Then
                    MsgBox "���ۼ۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If mdblOldPrice = .EditText Then Exit Sub

                If strKey <> "" Then
'                    If zlCommFun.DblIsValid(strkey, 12, , False, , "�ּ�") = False Then Cancel = True: Exit Sub
                    If Val(.TextMatrix(Row, menuPriceCol.����ID)) = 0 Then
                        vsfPrice.EditText = Format(Val(strKey), mFMT.FM_���ۼ�)
                        Exit Sub
                    End If
                    If Val(strKey) > Val(.TextMatrix(Row, menuPriceCol.ԭָ���ۼ�)) And Val(.TextMatrix(Row, menuPriceCol.ԭָ���ۼ�)) <> 0 Then
                        If MsgBox("�ּ۲��ܴ���ָ�����ۼۣ���" & Format(Val(.TextMatrix(Row, menuPriceCol.ԭָ���ۼ�)), mFMT.FM_���ۼ�) & "����������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            vsfPrice.EditText = Format(Val(strKey), mFMT.FM_���ۼ�)
                            vsfPrice.TextMatrix(Row, menuPriceCol.��ָ���ۼ�) = Format(Val(strKey), mFMT.FM_���ۼ�)
                        End If
                    Else
                        vsfPrice.EditText = Format(Val(strKey), mFMT.FM_���ۼ�)
                    End If

                End If
                If chkAppAllColumn.Value = 1 And mlngPrice <> vsfPrice.EditText Then
                    For intRow = 1 To .Rows - 1
                        If .TextMatrix(intRow, menuPriceCol.����ID) <> "" Then
                            .TextMatrix(intRow, menuPriceCol.�����ۼ�) = vsfPrice.EditText
                        End If
                    Next
                End If
                If chkAppAllColumn.Value = 0 Then
                    Call FullStoce�ּ�(Val(.TextMatrix(Row, menuPriceCol.����ID)), vsfPrice.EditText)
                End If
        End Select
    End With
End Sub

Private Sub FullStoce�ɱ���(ByVal lng����ID, ByVal dbl�ɱ��� As Double)
    '�ɱ���
    Dim lngRow As Long, dbl������ As Double
    With vsfStore
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, menuStoreCol.����ID)) = lng����ID Then
                .TextMatrix(lngRow, menuStoreCol.�ֲɹ���) = Format(dbl�ɱ���, mFMT.FM_�ɱ���)
                 Call AutoCalcStoce(lngRow, menuStoreCol.�ֲɹ���)
            End If
        Next
    End With
End Sub

Private Sub FullStoce�ּ�(ByVal lng����ID As Long, ByVal dbl�ּ� As Double)
    '-----------------------------------------------------------------------------------------------------------
    '����:�����ּ�,�����䶯���ּۼ�������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-11-07 10:32:13
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl������ As Double
    With vsfStore
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, menuStoreCol.����ID)) = lng����ID Then
                .TextMatrix(lngRow, menuStoreCol.�����ۼ�) = Format(dbl�ּ�, mFMT.FM_���ۼ�)
                '������=����*(�ּ�-ԭ��)
                dbl������ = (dbl�ּ� - Val(.TextMatrix(lngRow, menuStoreCol.ԭ���ۼ�))) * Val(.TextMatrix(lngRow, menuStoreCol.����))
                .TextMatrix(lngRow, menuStoreCol.�������) = Format(dbl������, mFMT.FM_���)
                '��Ҫ���ݼӳ������¼�������ĳɱ���
'                 Call AutoCalcStoce(lngRow, menuStoreCol.�����ۼ�)
                If Val(.TextMatrix(lngRow, menuStoreCol.�ֲɹ���)) <> 0 Then
                    .TextMatrix(lngRow, menuStoreCol.�ӳ���) = Format(Val((.TextMatrix(lngRow, menuStoreCol.�����ۼ�)) / Val(.TextMatrix(lngRow, menuStoreCol.�ֲɹ���)) - 1) * 100, "#0.000")
                Else
                    .TextMatrix(lngRow, menuStoreCol.�ӳ���) = 0
                End If
            End If
        Next
    End With
End Sub

Private Sub AutoCalcStoce(ByVal lngEditRow As Long, ByVal lngEditCol As Long)
    '-----------------------------------------------------------------------------------------------------------
    '����:�Զ����������Ϣ(���ݼӳ��ʼ����ֳɱ��ۼ����,�����ֳɱ��ۼ�����ӳ���)
    '���:lngEditRow-��ǰ�༭����
    '     lngEditCol-��ǰ�༭����
    '����:
    '����:
    '����:���˺�
    '����:2008-11-06 17:03:02
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl�ֳɱ��� As Double, dbl�ӳ��� As Double, dbl�ɱ���� As Double, dbl��۵����� As Double
    Dim lng����ID As Long, bln�ⷿ���� As Boolean, lng��Ӧ��ID As Long, lngTemp As Long, i As Long
    Dim blnHaveData As Boolean, lngStep As Long, lngSteps As Long

    err = 0: On Error GoTo ErrHand:
    With vsfStore
        bln�ⷿ���� = chkCostBatch.Value = 1
        lngStep = IIf(bln�ⷿ����, lngEditRow, 1)
        lngSteps = IIf(bln�ⷿ����, lngEditRow, .Rows - 1)
        Select Case lngEditCol
        Case menuStoreCol.�ӳ���
            dbl�ӳ��� = Val(.TextMatrix(lngEditRow, lngEditCol)) / 100
            If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
            '�ֳɱ���=�����ۼ�/(1+�ӳ���)
            dbl�ֳɱ��� = Format(Val(.TextMatrix(lngEditRow, menuStoreCol.�����ۼ�)) / (1 + dbl�ӳ���), mFMT.FM_�ɱ���)
            '��۵�����=(ԭ�ɱ���-�ֳɱ���)
            dbl�ɱ���� = dbl�ֳɱ��� - Val(.TextMatrix(lngEditRow, menuStoreCol.ԭ�ɹ���))
        Case menuStoreCol.�ֲɹ���
            '��Ϊ���ڰ�װ�������⣬��ˣ�Ŀǰ����С��λ�������õ���
            dbl�ֳɱ��� = Val(.TextMatrix(lngEditRow, lngEditCol))
            '�ӳ���=�����ۼ�/�ֳɱ���-1
            If dbl�ֳɱ��� <> 0 Then
                dbl�ӳ��� = Round(Val(.TextMatrix(lngEditRow, menuStoreCol.�����ۼ�)) / dbl�ֳɱ��� - 1, 7)
            Else
                dbl�ӳ��� = 0
            End If
            '��۵�����=(�ֳɱ���-ԭ�ɱ���)
            dbl�ɱ���� = Format((dbl�ֳɱ��� - Val(.TextMatrix(lngEditRow, menuStoreCol.ԭ�ɹ���))), mFMT.FM_�ɱ���)
        Case menuStoreCol.��۲�
            Exit Sub
        Case menuStoreCol.�����ۼ�
            '�ּ۷����ı�ʱ,��Ҫ���¸��ݼӳ��ʼ�����ص��ֳɱ���
'            dbl�ӳ��� = Round(Val(.TextMatrix(lngEditRow, menuStoreCol.�ӳ���)) / 100, 7)
'            If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
'            '�ֳɱ���=�����ۼ�/(1+�ӳ���)
'            dbl�ֳɱ��� = Format(Val(.TextMatrix(lngEditRow, menuStoreCol.�����ۼ�)) / (1 + dbl�ӳ���), mFMT.FM_�ɱ���)
'            '��۵�����=(�ֳɱ���-ԭ�ɱ���)
'            dbl�ɱ���� = (dbl�ֳɱ��� - Val(.TextMatrix(lngEditRow, menuStoreCol.ԭ�ɹ���)))


            '�ּ۸ı�ʱ����Ҫ���¸��ݳɱ��ۼ�����صļӳ���
            dbl�ֳɱ��� = Val(.TextMatrix(lngEditRow, menuStoreCol.�ֲɹ���))
            If dbl�ֳɱ��� = 0 Then
                dbl�ӳ��� = 0
            Else
                dbl�ӳ��� = Round(Val(.TextMatrix(lngEditRow, menuStoreCol.�����ۼ�)) / dbl�ֳɱ��� - 1, 7)
            End If
            lngStep = lngEditRow
            lngSteps = lngEditRow
        Case Else
            Exit Sub
        End Select

        lng����ID = Val(.TextMatrix(lngEditRow, menuStoreCol.����ID))
        lng��Ӧ��ID = Val(.TextMatrix(lngEditRow, menuStoreCol.��Ӧ��ID))
        Dim cllData As New Collection
        For lngRow = lngStep To lngSteps
            If lng����ID = Val(.TextMatrix(lngRow, menuStoreCol.����ID)) Then
                If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
                .TextMatrix(lngRow, menuStoreCol.�ӳ���) = Format(dbl�ӳ��� * 100, GFM_VBJCL)
                '�óɱ���������С��λΪ׼�ģ����Ҫ��С����ϵ��.
                .TextMatrix(lngRow, menuStoreCol.�ֲɹ���) = Format(dbl�ֳɱ���, mFMT.FM_�ɱ���)
                dbl�ɱ���� = dbl�ֳɱ��� - Val(.TextMatrix(lngRow, menuStoreCol.ԭ�ɹ���))
                 '��۵�����=(�ֳɱ���-ԭ�ɱ���)*����
                 dbl��۵����� = Round(dbl�ɱ���� * Val(.TextMatrix(lngRow, menuStoreCol.����)), 7)
                .TextMatrix(lngRow, menuStoreCol.��۲�) = Format(dbl��۵�����, mFMT.FM_���)
'                .TextMatrix(lngRow, menuStoreCol.��۲�) = dbl��۵�����
                lngTemp = Val(.TextMatrix(lngRow, menuStoreCol.����ID))
                lng��Ӧ��ID = Val(.TextMatrix(lngRow, menuStoreCol.��Ӧ��ID))

                If lng��Ӧ��ID <> 0 Then
                    err = 0: On Error Resume Next
                    cllData.Add Array(lngTemp, lng��Ӧ��ID, dbl��۵�����, .TextMatrix(lngRow, menuStoreCol.��Ӧ��), .TextMatrix(lngRow, menuStoreCol.ҩƷ), .TextMatrix(lngRow, menuStoreCol.���), .TextMatrix(lngRow, menuStoreCol.����)), "K" & lng��Ӧ��ID & "_" & lngTemp
                    If err <> 0 Then
                        '�ۼƲ�۵�����
                        dbl��۵����� = Val(cllData("K" & lng��Ӧ��ID & "_" & lngTemp)(2)) + dbl��۵�����
                        cllData.Remove "K" & lng��Ӧ��ID & "_" & lngTemp
                         err = 0: On Error GoTo ErrHand:
                        cllData.Add Array(lngTemp, lng��Ӧ��ID, dbl��۵�����, .TextMatrix(lngRow, menuStoreCol.��Ӧ��), .TextMatrix(lngRow, menuStoreCol.ҩƷ), .TextMatrix(lngRow, menuStoreCol.���), .TextMatrix(lngRow, menuStoreCol.����)), "K" & lng��Ӧ��ID & "_" & lngTemp

                    End If
                    On Error GoTo ErrHand:
                End If
            End If
        Next
        If chkAutoPay.Value = 1 Then
            '��Ҫ�Զ�������ص�Ӧ���䶯��¼
            For i = 1 To cllData.Count
                With vsfPay
                    blnHaveData = False
                    For lngRow = 1 To .Rows - 1
                        lngTemp = Val(.TextMatrix(lngRow, menuPayCol.����ID))
                        lng��Ӧ��ID = Val(.TextMatrix(lngRow, menuPayCol.��Ӧ��ID))
                        If lngTemp = Val(cllData(i)(0)) And lng��Ӧ��ID = Val(cllData(i)(1)) Then
                            '���ļ���Ӧ����ͬ,�����ص�ֵ
                            .TextMatrix(lngRow, menuPayCol.��Ʊ���) = Format(Val(cllData(i)(2)), mFMT.FM_���)
                             blnHaveData = True
                        End If
                    Next
                    If blnHaveData = False Then
                        '��Ҫ���Ӹ��Ӧ�̵�����
                        If Val(.TextMatrix(.Rows - 1, menuPayCol.����ID)) <> 0 Or .Rows = 1 Then
                            .Rows = .Rows + 1
                            .RowHeight(.Rows - 1) = mconlngRowHeight
                            Call setColEdit
                        End If
                        lngRow = .Rows - 1
                        .TextMatrix(lngRow, menuPayCol.��Ӧ��) = cllData(i)(3)
                        .TextMatrix(lngRow, menuPayCol.��Ӧ��ID) = cllData(i)(1)
                        .TextMatrix(lngRow, menuPayCol.����ID) = cllData(i)(0)
                        .TextMatrix(lngRow, menuPayCol.Ʒ��) = cllData(i)(4)
                        .TextMatrix(lngRow, menuPayCol.���) = cllData(i)(5)
                        .TextMatrix(lngRow, menuPayCol.����) = cllData(i)(6)
                        .TextMatrix(lngRow, menuPayCol.��Ʊ���) = Format(Val(cllData(i)(2)), mFMT.FM_���)
                    End If
                End With
            Next
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
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

    If intRow = 0 Or mint���� = 1 Then Exit Sub

'    dblOldPrice = Val(vsfPrice.TextMatrix(intRow, menuPriceCol.ԭ���ۼ�))
'    dbl��װ = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.��װϵ��))
'
'    With vsfStore
'        For n = 1 To .Rows - 1
'            If .TextMatrix(n, 0) <> "" Then
'                If Val(.TextMatrix(n, menuStoreCol.����id)) = lngDrugID Then
'                    dblNum = Val(.TextMatrix(n, menuStoreCol.����))
'
'                    .TextMatrix(n, menuStoreCol.�����ۼ�) = format(dblNewPrice, mFMT.FM_���ۼ�)
'                    .TextMatrix(n, menuStoreCol.�������) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (dblNewPrice - dblOldPrice), mFMT.fm_���)
'
'                    If mint���� = 2 And chkAotuCost.Value = 1 Then
'                        dblOldCost = .TextMatrix(n, menuStoreCol.ԭ�ɹ���)
'                        dblNewCost = dblNewPrice / (1 + Round(Val(.TextMatrix(n, menuStoreCol.�ӳ���)) / 100, 7))
'                        .TextMatrix(n, menuStoreCol.�ֲɹ���) = format(dblNewCost, mFMT.FM_�ɱ���)
'                        .TextMatrix(n, menuStoreCol.��۲�) = Format((dblNewCost - dblOldCost) * dblNum, mFMT.fm_���)
'                        dbl��Ʊ��� = dbl��Ʊ��� + (dblNewCost - dblOldCost) * dblNum
'                    End If
'                End If
'            End If
'        Next
'    End With
'
'    If chkAutoPay.Value = 1 Then
'        With vsfPay
'            For n = 1 To .Rows - 1
'                If .TextMatrix(1, 0) <> "" Then
'                    If Val(.TextMatrix(n, menuPayCol.����id)) = lngDrugID Then
'                        .TextMatrix(n, menuPayCol.��Ʊ���) = format(dbl��Ʊ���, 2)
'                    End If
'                End If
'            Next
'        End With
'    End If
'
'    CaluateAverCost lngDrugID
End Sub

Private Sub CaluateAverOldCost(ByVal lng����ID As Long)
    '����ԭʼƽ���ɱ���
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .Rows - 1
            If .TextMatrix(i, menuStoreCol.����ID) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.����ID)) = lng����ID Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, menuPriceCol.����ID) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.����ID)) = lng����ID Then
                        .TextMatrix(i, menuPriceCol.ԭ�ɱ���) = Format(dblSumCost / dblSumNumber, mFMT.FM_�ɱ���)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaluateAverCost(ByVal lng����ID As Long)
    '����ƽ���ɱ���
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .Rows - 1
            If .TextMatrix(i, menuStoreCol.����ID) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.����ID)) = lng����ID Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.�ֲɹ���)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, menuPriceCol.����ID) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.����ID)) = lng����ID Then
                        .TextMatrix(i, menuPriceCol.�ֳɱ���) = Format(dblSumCost / dblSumNumber, mFMT.FM_�ɱ���)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
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
            .ColHidden(menuStoreCol.����) = True
            .ColHidden(menuStoreCol.���) = True
            .ColHidden(menuStoreCol.�ӳ���) = True
            .ColHidden(menuStoreCol.ԭ�ɹ���) = True
            .ColHidden(menuStoreCol.�ֲɹ���) = True
            .ColHidden(menuStoreCol.��۲�) = True
            .ColHidden(menuStoreCol.ԭ���ۼ�) = False
            .ColHidden(menuStoreCol.�����ۼ�) = False
        ElseIf cboPriceMethod.Text = "�����ɱ���" Then
            .ColHidden(menuStoreCol.ԭ���ۼ�) = True
            .ColHidden(menuStoreCol.�����ۼ�) = True
            .ColHidden(menuStoreCol.�������) = True
            .ColHidden(menuStoreCol.�ӳ���) = False
            .ColHidden(menuStoreCol.ԭ�ɹ���) = False
            .ColHidden(menuStoreCol.�ֲɹ���) = False
            .ColHidden(menuStoreCol.��۲�) = False
        ElseIf cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
            .ColHidden(menuStoreCol.ԭ���ۼ�) = False
            .ColHidden(menuStoreCol.�����ۼ�) = False
            .ColHidden(menuStoreCol.�������) = False
            .ColHidden(menuStoreCol.�ӳ���) = False
            .ColHidden(menuStoreCol.ԭ�ɹ���) = False
            .ColHidden(menuStoreCol.�ֲɹ���) = False
            .ColHidden(menuStoreCol.��۲�) = False
        End If
    End With
End Sub

Private Sub vsfStore_Click()
    Dim i As Integer
    With vsfStore
        For i = 1 To vsfPrice.Rows - 1
            If Val(.TextMatrix(.Row, menuStoreCol.����ID)) = Val(vsfPrice.TextMatrix(i, menuPriceCol.����ID)) Then
                vsfPrice.Tag = i
            End If
        Next
    End With
End Sub

Private Sub vsfStore_DblClick()
    With vsfStore
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
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
        ElseIf .Col = menuStoreCol.�ֲɹ��� Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.�ֲɹ���))
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
                If .Row <> .Rows - 1 Then
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
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfStore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer

    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfStore
            If Col = menuStoreCol.�ֲɹ��� Or Col = menuStoreCol.�����ۼ� Or Col = menuStoreCol.�ӳ��� Then
                strKey = .EditText
                Select Case Col
                    Case menuStoreCol.�ֲɹ���
                        intDigit = Len(Mid(mFMT.FM_�ɱ���, InStr(1, mFMT.FM_�ɱ���, ".") + 1))
                    Case menuStoreCol.�����ۼ�
                        intDigit = Len(Mid(mFMT.FM_�ɱ���, InStr(1, mFMT.FM_���ۼ�, ".") + 1))
                    Case menuStoreCol.�ӳ���
                        intDigit = 5
                End Select
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strKey) Then Exit Sub
                    If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
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
    Dim dbl���� As Double
    Dim Dbl��� As Double
    Dim dbl�ֲɹ��� As Double

    With vsfStore
        If .EditText = "" Then Exit Sub
        intRow = .Row
        Select Case .Col
            Case menuStoreCol.�����ۼ�
                If Not IsNumeric(.EditText) Then
                    MsgBox "���������֣�", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                Else
                    .EditText = Format(.EditText, mFMT.FM_���ۼ�)
                End If

'                If mdblOldPrice = .EditText Then Exit Sub

                If .EditText > 9999999 Then
                    MsgBox "���ۼ۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                .TextMatrix(intRow, menuStoreCol.�������) = Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * (Val(.EditText) - Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mFMT.FM_���)
                .TextMatrix(intRow, menuStoreCol.�����ۼ�) = Format(Val(.EditText), mFMT.FM_���ۼ�)
                If Val(.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) <> 0 Then
                    .TextMatrix(intRow, menuStoreCol.�ӳ���) = GetFormat((Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) / Val(.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) - 1) * 100, 3)
                End If
                
                For n = 1 To .Rows - 1
                    If .TextMatrix(intRow, menuStoreCol.����ID) = .TextMatrix(n, menuStoreCol.����ID) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.����)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.����)) = Val(.TextMatrix(n, menuStoreCol.����)) Then
                            .TextMatrix(n, menuStoreCol.�����ۼ�) = .TextMatrix(intRow, menuStoreCol.�����ۼ�)
                            .TextMatrix(n, menuStoreCol.�������) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (Val(.EditText) - Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�))), mFMT.FM_���)
                            If Val(.TextMatrix(n, menuStoreCol.�ֲɹ���)) <> 0 Then
                                .TextMatrix(n, menuStoreCol.�ӳ���) = GetFormat((Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) / Val(.TextMatrix(n, menuStoreCol.�ֲɹ���)) - 1) * 100, 3)
                            End If
                        End If
                        dbl���� = dbl���� + .TextMatrix(n, menuStoreCol.����)
                        Dbl��� = Dbl��� + .TextMatrix(n, menuStoreCol.����) * Val(.TextMatrix(n, menuStoreCol.�����ۼ�))
                    End If
                Next
                For n = 1 To vsfPrice.Rows - 1
                    If .TextMatrix(intRow, menuStoreCol.����ID) = vsfPrice.TextMatrix(n, menuPriceCol.����ID) Then
                        If dbl���� <> 0 Then
                            vsfPrice.TextMatrix(n, menuPriceCol.�����ۼ�) = Format(Dbl��� / dbl����, mFMT.FM_���ۼ�)
                        Else
                            vsfPrice.TextMatrix(n, menuPriceCol.�����ۼ�) = .TextMatrix(intRow, menuStoreCol.�����ۼ�)
                        End If
                    End If
                Next

                If mint���� > 0 Then
                    For n = 1 To .Rows - 1
                        If .TextMatrix(n, menuStoreCol.����ID) <> "" Then
                            If Val(.TextMatrix(n, menuStoreCol.����ID)) = Val(.TextMatrix(intRow, menuStoreCol.����ID)) Then
                                dbl��Ʊ��� = dbl��Ʊ��� + (Val(.TextMatrix(n, menuStoreCol.�ֲɹ���)) - Val(.TextMatrix(n, menuStoreCol.ԭ�ɹ���))) * Val(.TextMatrix(n, menuStoreCol.����))
                            End If
                        End If
                    Next

                    If chkAutoPay.Value = 1 Then
                        For n = 1 To vsfPay.Rows - 1
                            If vsfPay.TextMatrix(1, 0) <> "" Then
                                If Val(vsfPay.TextMatrix(n, menuPayCol.����ID)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.����ID)) Then
                                    vsfPay.TextMatrix(n, menuPayCol.��Ʊ���) = Format(dbl��Ʊ���, 2)
                                End If
                            End If
                        Next
                    End If
                End If
            Case menuStoreCol.�ӳ���
                If Val(.EditText) < 0 Then Exit Sub

                If Not IsNumeric(.EditText) Then
                    MsgBox "���������֣�", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
'                If mdblOldPrice = .EditText Then Exit Sub

                .TextMatrix(intRow, menuStoreCol.�ӳ���) = Format(Val(.EditText), "#0.000")
                .TextMatrix(intRow, menuStoreCol.�����ۼ�) = Format(Val(.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) * (1 + (Format(Val(.EditText), "#0.000")) / 100), mFMT.FM_���ۼ�)
                .TextMatrix(intRow, menuStoreCol.�������) = Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * (Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mFMT.FM_���)

                For n = 1 To .Rows - 1
                    If vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.����ID) = .TextMatrix(n, menuStoreCol.����ID) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.���)) = 0 Or mblnʱ�����İ����ε��� = False Then
                            .TextMatrix(n, menuStoreCol.�ӳ���) = Format(.TextMatrix(intRow, menuStoreCol.�ӳ���), "#0.000")
                            .TextMatrix(n, menuStoreCol.�����ۼ�) = Format(Val(.TextMatrix(n, menuStoreCol.�ֲɹ���)) * (1 + (Format(Val(.EditText), "#0.000")) / 100), mFMT.FM_���ۼ�)
                            .TextMatrix(n, menuStoreCol.�������) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�))), mFMT.FM_���)
                        End If
                        dbl���� = dbl���� + .TextMatrix(n, menuStoreCol.����)
                        Dbl��� = Dbl��� + .TextMatrix(n, menuStoreCol.����) * Val(.TextMatrix(n, menuStoreCol.�����ۼ�))
                    End If
                Next
                If dbl���� <> 0 Then
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.�����ۼ�) = Format(Dbl��� / dbl����, mFMT.FM_���ۼ�)
                Else
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.�����ۼ�) = .TextMatrix(intRow, menuStoreCol.�����ۼ�)
                End If
            Case menuStoreCol.�ֲɹ���
                If Not IsNumeric(.EditText) Then
                    MsgBox "���������֣�", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If .EditText > 9999999 Then
                    MsgBox "�ɹ��۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If Val(.EditText) < 0 Then
                    MsgBox "�ɱ��۲���Ϊ������", vbExclamation, gstrSysName
                    Cancel = True
                End If

                .EditText = Format(Val(.EditText), mFMT.FM_�ɱ���)
                .TextMatrix(intRow, menuStoreCol.�ֲɹ���) = Format(Val(.EditText), mFMT.FM_�ɱ���)
                .TextMatrix(intRow, menuStoreCol.��۲�) = Format((Val(.EditText) - .TextMatrix(intRow, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(intRow, menuStoreCol.����)), mFMT.FM_���)
                If Val(.TextMatrix(intRow, menuStoreCol.���)) = 1 And mblnʱ�����İ����ε��� = True And mint���� <> 1 Then
                    .TextMatrix(intRow, menuStoreCol.�����ۼ�) = Format(Val(.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) * (1 + (Val(.TextMatrix(intRow, menuStoreCol.�ӳ���)) / 100)), mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, menuStoreCol.�������) = Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * (Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mFMT.FM_���)
                End If
                
                dbl��Ʊ��� = (Val(.EditText) - .TextMatrix(intRow, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(intRow, menuStoreCol.����))

                For n = 1 To .Rows - 1
                    If .TextMatrix(n, menuStoreCol.����ID) <> "" Then
                        If Val(.TextMatrix(n, menuStoreCol.����ID)) = Val(.TextMatrix(intRow, menuStoreCol.����ID)) And n <> intRow Then
                            If chkCostBatch.Value = 0 Or (Val(.TextMatrix(intRow, menuStoreCol.����)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.����)) = Val(.TextMatrix(n, menuStoreCol.����))) Then
                                dbl�ֲɹ��� = Format(Val(.EditText), mFMT.FM_�ɱ���)
                                .TextMatrix(n, menuStoreCol.�ֲɹ���) = Format(dbl�ֲɹ���, mFMT.FM_�ɱ���)
                                .TextMatrix(n, menuStoreCol.��۲�) = Format((dbl�ֲɹ��� - .TextMatrix(n, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(n, menuStoreCol.����)), mFMT.FM_���)
                                If Val(.TextMatrix(intRow, menuStoreCol.���)) = 1 And mblnʱ�����İ����ε��� = True And mint���� <> 1 Then
                                    .TextMatrix(n, menuStoreCol.�����ۼ�) = Format(dbl�ֲɹ��� * (1 + (Val(.TextMatrix(n, menuStoreCol.�ӳ���)) / 100)), mFMT.FM_���ۼ�)
                                    .TextMatrix(n, menuStoreCol.�������) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�))), mFMT.FM_���)
                                End If
                            Else
                                dbl�ֲɹ��� = Val(.TextMatrix(n, menuStoreCol.�ֲɹ���))
                            End If
                            dbl��Ʊ��� = dbl��Ʊ��� + (dbl�ֲɹ��� - .TextMatrix(n, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(n, menuStoreCol.����))
                        End If
                    End If
                Next

                If chkAutoPay.Value = 1 Then
                    For n = 1 To vsfPay.Rows - 1
                        If vsfPay.TextMatrix(1, 0) <> "" Then
                            If Val(vsfPay.TextMatrix(n, menuPayCol.����ID)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.����ID)) Then
                                vsfPay.TextMatrix(n, menuPayCol.��Ʊ���) = Format(dbl��Ʊ���, mFMT.FM_���)
                            End If
                        End If
                    Next
                End If

                If chkCostBatch.Value = 0 Then
                    For n = 1 To vsfPrice.Rows - 1
                        If Val(.TextMatrix(intRow, menuStoreCol.����ID)) = Val(vsfPrice.TextMatrix(n, menuPriceCol.����ID)) Then
                            vsfPrice.TextMatrix(n, menuPriceCol.�ֳɱ���) = Format(.TextMatrix(intRow, menuStoreCol.�ֲɹ���), mFMT.FM_�ɱ���)
                            Exit For
                        End If
                    Next
                Else
                    CaluateAverCost Val(.TextMatrix(intRow, menuStoreCol.����ID))
                End If
                Call CaculateAverPirce(Val(.TextMatrix(intRow, menuStoreCol.����ID)))   '�ۼ۱䶯������ƽ���ۼ�
        End Select
    End With
End Sub

Private Sub CaculateAverPirce(ByVal lng����ID As Long)
    '�Զ�����ƽ���ۼ�
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .Rows - 1
            If .TextMatrix(i, menuStoreCol.����ID) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.����ID)) = lng����ID Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.�����ۼ�)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, menuPriceCol.����ID) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.����ID)) = lng����ID Then
                        .TextMatrix(i, menuPriceCol.�����ۼ�) = Format(dblSumPrice / dblSumNumber, mFMT.FM_���ۼ�)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaculateAverOldPirce(ByVal lng����ID As Long)
    '�Զ�ԭʼ����ƽ���ۼ�
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .Rows - 1
            If .TextMatrix(i, menuStoreCol.����ID) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.����ID)) = lng����ID Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.ԭ���ۼ�)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, menuPriceCol.����ID) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.����ID)) = lng����ID Then
                        .TextMatrix(i, menuPriceCol.ԭ���ۼ�) = Format(dblSumPrice / dblSumNumber, mFMT.FM_���ۼ�)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CatalogModifyPrice()
'����Ŀ¼ֱ�ӽ������
    Dim rsprice As New ADODB.Recordset
    gstrSQL = "Select Distinct i.Id,i.����,b.���� As ��Ʒ��,i.���� As ͨ����,i.���,i.����,i.���㵥λ,p.����ϵ��,p.��װ��λ," & vbNewLine & _
        "                Decode(i.�Ƿ���, 0, '����', 1, 'ʱ��') As ʱ��," & vbNewLine & _
        "                To_Char(p.�ɱ���, '9999999999990.9999999') As �ɱ���," & vbNewLine & _
        "                To_Char(p.ָ��������, '9999999999990.9999999') ָ��������," & vbNewLine & _
        "                To_Char(p.ָ�����ۼ�, '9999999999990.9999999') ָ�����ۼ�," & vbNewLine & _
        "                p.��������" & vbNewLine & _
        "From �շ���ĿĿ¼ i, �������� p, �շ���Ŀ���� b" & vbNewLine & _
        "Where i.Id = p.����id And i.Id = b.�շ�ϸĿid(+) And b.����(+) = 3 And i.��� = '4' and i.id=[1] And" & vbNewLine & _
        "      (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd'))"
    
    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "����ֱ�ӽ������", mlng���ID)
    
    Call GetDrugPirce(rsprice, 1)
    chkAppAllColumn.Enabled = False
End Sub
