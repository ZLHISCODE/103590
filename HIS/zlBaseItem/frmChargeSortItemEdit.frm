VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChargeSortItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ѱ����շ�����"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "frmChargeSortItemEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDel 
      Caption         =   "���(&D)"
      Height          =   350
      Left            =   3600
      TabIndex        =   84
      Top             =   3600
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.Frame fraItem 
      Caption         =   "��Ŀѡ��"
      Height          =   3285
      Left            =   3600
      TabIndex        =   75
      Top             =   120
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdMoveAll 
         Caption         =   "�Ƴ�����(&C)"
         Height          =   350
         Left            =   4560
         TabIndex        =   83
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "�Ƴ�(&M)"
         Height          =   350
         Left            =   4560
         TabIndex        =   82
         Top             =   240
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfItemList 
         Height          =   1935
         Left            =   120
         TabIndex        =   81
         Top             =   1200
         Width           =   5655
         _cx             =   9975
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
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
      Begin VB.CommandButton cmdFilter 
         Caption         =   "��"
         Height          =   270
         Left            =   3000
         TabIndex        =   80
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtInput 
         Height          =   270
         Left            =   960
         TabIndex        =   79
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cbo��Ŀ��� 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ����"
         Height          =   180
         Left            =   120
         TabIndex        =   78
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ���"
         Height          =   180
         Left            =   120
         TabIndex        =   76
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame fraҩƷӦ�� 
      Caption         =   "Ӧ�÷�Χ"
      Height          =   3300
      Left            =   4560
      TabIndex        =   68
      Top             =   3960
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton optӦ���� 
         Caption         =   "��Ӧ���ڱ����ҩƷ(&0)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   74
         Top             =   480
         Value           =   -1  'True
         Width           =   2955
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ�������С�����ҩ��(&2)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   73
         Top             =   1392
         Width           =   3795
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ�������С�Ƭ������ҩƷ(&3)"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   72
         Top             =   1848
         Width           =   4275
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ���ڱ�Ʒ��������ҩƷ(&1)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   71
         Top             =   936
         Width           =   2955
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ����ͬ��������ҩƷ(&4)"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   70
         Top             =   2304
         Width           =   2955
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ���ڱ����������ҩƷ(&5)"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   69
         Top             =   2760
         Width           =   2955
      End
   End
   Begin VB.Frame fra��ĿӦ�� 
      Caption         =   "Ӧ�÷�Χ"
      Height          =   3300
      Left            =   840
      TabIndex        =   63
      Top             =   4200
      Visible         =   0   'False
      Width           =   3495
      Begin VB.OptionButton optApply 
         Caption         =   "Ӧ���ڸ÷�����������Ŀ(&2)"
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   67
         Top             =   1680
         Width           =   3075
      End
      Begin VB.OptionButton optApply 
         Caption         =   "Ӧ���ڸ������������Ŀ(&3)"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   66
         Top             =   2280
         Width           =   3075
      End
      Begin VB.OptionButton optApply 
         Caption         =   "Ӧ����ͬ����������Ŀ(&1)"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   65
         Top             =   1080
         Width           =   3075
      End
      Begin VB.OptionButton optApply 
         Caption         =   "���Ա���Ŀ������(&0)"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   64
         Top             =   480
         Value           =   -1  'True
         Width           =   3075
      End
   End
   Begin VB.Frame fra�ѱ� 
      Caption         =   "�ѱ���ϸ"
      Height          =   3300
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   735
         TabIndex        =   43
         Text            =   "0.00"
         Top             =   2385
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   0
         Left            =   1845
         TabIndex        =   42
         Text            =   "100.000"
         Top             =   2385
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   1
         Left            =   735
         TabIndex        =   41
         Text            =   "0.00"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   1
         Left            =   1845
         TabIndex        =   40
         Text            =   "100.000"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   2
         Left            =   735
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   2895
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   2
         Left            =   1845
         TabIndex        =   38
         Text            =   "100.000"
         Top             =   2895
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   3
         Left            =   735
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   3150
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   3
         Left            =   1845
         TabIndex        =   36
         Text            =   "100.000"
         Top             =   3150
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   4
         Left            =   1845
         TabIndex        =   35
         Text            =   "100.000"
         Top             =   3405
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   4
         Left            =   735
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   3405
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   5
         Left            =   1845
         TabIndex        =   33
         Text            =   "100.000"
         Top             =   3660
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   5
         Left            =   735
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   3660
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   6
         Left            =   1845
         TabIndex        =   31
         Text            =   "100.000"
         Top             =   3915
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   6
         Left            =   735
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   3915
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   7
         Left            =   1845
         TabIndex        =   29
         Text            =   "100.000"
         Top             =   4170
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   7
         Left            =   735
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   4170
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   8
         Left            =   1845
         TabIndex        =   27
         Text            =   "100.000"
         Top             =   4425
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   8
         Left            =   735
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   4425
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   9
         Left            =   1845
         TabIndex        =   25
         Text            =   "100.000"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   9
         Left            =   735
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   10
         Left            =   1845
         TabIndex        =   23
         Text            =   "100.000"
         Top             =   4935
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   10
         Left            =   735
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   4935
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   11
         Left            =   1845
         TabIndex        =   21
         Text            =   "100.000"
         Top             =   5190
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   11
         Left            =   735
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   5190
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   12
         Left            =   1845
         TabIndex        =   19
         Text            =   "100.000"
         Top             =   5445
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   12
         Left            =   735
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   5445
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   13
         Left            =   1845
         TabIndex        =   17
         Text            =   "100.000"
         Top             =   5700
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   13
         Left            =   735
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   5700
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   14
         Left            =   1845
         TabIndex        =   15
         Text            =   "100.000"
         Top             =   5955
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   14
         Left            =   735
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   5955
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   15
         Left            =   1845
         TabIndex        =   13
         Text            =   "100.000"
         Top             =   6210
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   15
         Left            =   735
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   6210
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtStage 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "1"
         Top             =   1800
         Width           =   300
      End
      Begin VB.ComboBox cbo���㷽�� 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin MSComCtl2.UpDown UdStage 
         Height          =   300
         Left            =   2880
         TabIndex        =   9
         Top             =   1800
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtStage"
         BuddyDispid     =   196626
         OrigLeft        =   2010
         OrigTop         =   1200
         OrigRight       =   2250
         OrigBottom      =   1500
         Max             =   16
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblMoney 
         Caption         =   "Ӧ�շֶ����"
         Height          =   180
         Left            =   750
         TabIndex        =   62
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label lblTax 
         Caption         =   "ʵ�ձ���(%)"
         Height          =   195
         Left            =   1965
         TabIndex        =   61
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Label lblStage 
         Caption         =   "�ֶκ�"
         Height          =   225
         Left            =   120
         TabIndex        =   60
         Top             =   2175
         Width           =   540
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   0
         Left            =   225
         TabIndex        =   59
         Top             =   2430
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   58
         Top             =   2685
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   2
         Left            =   225
         TabIndex        =   57
         Top             =   2940
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   3
         Left            =   225
         TabIndex        =   56
         Top             =   3195
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   55
         Top             =   3450
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   180
         Index           =   5
         Left            =   225
         TabIndex        =   54
         Top             =   3705
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "7"
         Height          =   180
         Index           =   6
         Left            =   225
         TabIndex        =   53
         Top             =   3960
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "8"
         Height          =   180
         Index           =   7
         Left            =   225
         TabIndex        =   52
         Top             =   4215
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "9"
         Height          =   180
         Index           =   8
         Left            =   225
         TabIndex        =   51
         Top             =   4470
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "10"
         Height          =   180
         Index           =   9
         Left            =   180
         TabIndex        =   50
         Top             =   4725
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "11"
         Height          =   180
         Index           =   10
         Left            =   180
         TabIndex        =   49
         Top             =   4980
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "12"
         Height          =   180
         Index           =   11
         Left            =   180
         TabIndex        =   48
         Top             =   5235
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "13"
         Height          =   180
         Index           =   12
         Left            =   180
         TabIndex        =   47
         Top             =   5490
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "14"
         Height          =   180
         Index           =   13
         Left            =   180
         TabIndex        =   46
         Top             =   5745
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "15"
         Height          =   180
         Index           =   14
         Left            =   180
         TabIndex        =   45
         Top             =   6000
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "16"
         Height          =   180
         Index           =   15
         Left            =   180
         TabIndex        =   44
         Top             =   6255
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������Ŀ�����ֶ�"
         Height          =   180
         Left            =   1080
         TabIndex        =   11
         Top             =   1860
         Width           =   1440
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmChargeSortItemEdit.frx":000C
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblNote 
         Caption         =   "    ÿһ������Ŀ�ɰ�Ӧ�ս���Ϊ���(���16��)�����ò�ͬ��ʵ�ձ�����"
         Height          =   690
         Left            =   720
         TabIndex        =   8
         Top             =   1125
         Width           =   2595
      End
      Begin VB.Label lblMeasure 
         AutoSize        =   -1  'True
         Caption         =   "���㷽��"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ѡ��ѱ�"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1320
      TabIndex        =   0
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2400
      TabIndex        =   1
      Top             =   3600
      Width           =   1100
   End
End
Attribute VB_Name = "frmChargeSortItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�ӿڴ������
Private mintType As Integer           '����(���ڸô�������ڶദ���ô����ֲ�ͬʹ�û���)��0-�ѱ�������(������Ŀ)��1-�ѱ�������(������Ŀ)��2-�շ���Ŀ�����У�3-ҩƷ������
Private mstrGrade As String           '�ѱ𣺷ѱ������Ϊ����ֵ����������Ϊ��
Private mlngItemId As Long            '��ĿID���ѱ������Ϊ0����������Ϊ����ֵ
Private mStrItem As String            '��Ŀ����

'��������
Private mintStage As Integer
Private mblnChange As Boolean         '�Ƿ�ı���
Private mblnOk As Boolean

Private Const mconstListHead = "��Ŀid,7,0|����,1,1000|����,1,1500|���,1,1500|��λ,1,800|�۸�,7,800"
Private Enum ��Ŀ�б�
    ��Ŀid = 0
    ���� = 1
    ���� = 2
    ��� = 3
    ��λ = 4
    �۸� = 5
    
    ���� = 6
End Enum

Private Const mcstFormHeight As Double = 4600
Private Const mcstFormWidth As Double = 3750
Private Const mcstFormChargeHeight As Double = 3300
Private mstrҩƷ�۸�ȼ�  As String, mstr���ļ۸�ȼ� As String, mstr��ͨ�۸�ȼ� As String

Private Sub GetDrugOtherInfo()
    '��Ҫ����ҩƷĿ¼�����еõ���ǰҩƷ�ļ��ͺͲ���
    Dim rsTemp As ADODB.Recordset
    Dim str���� As String
    
    If mintType <> 3 Then Exit Sub
    If mlngItemId = 0 Then Exit Sub
    
    On Error GoTo ErrHandle
    gstrSQL = "Select Decode(A.���, '5', '����ҩ', '6', '�г�ҩ', '�в�ҩ') As ���, B.ҩƷ���� " & _
        " From �շ���ĿĿ¼ A, ҩƷ���� B, ҩƷ��� C " & _
        " Where A.ID = C.ҩƷid And B.ҩ��id = C.ҩ��id And A.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ��Ϣ", mlngItemId)
    
    If Not rsTemp.EOF Then
        optӦ����(2).Caption = "Ӧ�������С�" & rsTemp!��� & "��(&2)"
        optӦ����(3).Caption = "Ӧ�������С�" & rsTemp!ҩƷ���� & "����ҩƷ(&3)"
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub IniItemList()
    Dim i As Integer
    Dim strArr As Variant
    Dim strTemp As Variant
    
    strTemp = Split(mconstListHead, "|")
    
    With vsfItemList
        .redraw = flexRDNone
        .Rows = 1
        .Cols = ��Ŀ�б�.����
        .SelectionMode = flexSelectionByRow
        .RowHeightMin = 300
        
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        
        .redraw = flexRDDirect
    End With
End Sub
Private Sub GetItemList(ByVal strInput As String, ByVal strItemType As String, Optional ByVal lngItemID As Long = 0)
    Dim rsTmp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    Dim strReturn As String 'ѡ���������ַ���
    Dim strHyID As Long
    Dim strSqlCondition As String
    
    On Error GoTo ErrHandle
    
    rsTmp.CursorLocation = adUseClient

    If InStr(strInput, "'") > 0 Then
        MsgBox "�����˷Ƿ��ַ���", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If lngItemID > 0 Then
        strSqlCondition = " And A.Id = [5] "
    Else
        If strInput <> "" Then
            strSqlCondition = " And (A.���� like [1] or A.���� like [1] or  ('['||A.����||']'||A.����  =[3])  or  B.���� like [2]) "
        End If
        If strItemType <> "0" Then
            strSqlCondition = strSqlCondition & " And A.��� = [4] "
        End If
    End If
    
    Dim strWherePriceGrade As String
    If mstr��ͨ�۸�ȼ� = "" And mstrҩƷ�۸�ȼ� = "" And mstr���ļ۸�ȼ� = "" Then
        strWherePriceGrade = " And d.�۸�ȼ� Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And d.�۸�ȼ� = [6])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || a.��� || ';') > 0 And d.�۸�ȼ� = [7])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And d.�۸�ȼ� = [8])" & vbNewLine & _
            "      Or (d.�۸�ȼ� Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From �շѼ�Ŀ" & vbNewLine & _
            "                          Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [6])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [7])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And �۸�ȼ� = [8])))))"
    End If
    gstrSQL = _
        "SELECT A.����,A.����," & _
        "A.���,A.���㵥λ,ltrim(rtrim(to_char(Sum(nvl(D.�ּ�,0)),'9999999990.00'))) �۸�,A.ID" & _
        " FROM" & _
        " (Select Distinct A.ID,A.����,A.����,A.���,A.���㵥λ,a.���" & _
        "   From �շ���ĿĿ¼ A,�շ���Ŀ���� B" & _
        "   WHERE A.ID = B.�շ�ϸĿID" & _
        "       And (A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or A.����ʱ�� is null)" & strSqlCondition & _
        "   ) A,�շѼ�Ŀ D" & vbNewLine & _
        " Where A.ID=D.�շ�ϸĿID(+)" & _
        "       And D.ִ������ <= SYSDATE AND (D.��ֹ���� > SYSDATE OR D.��ֹ���� IS NULL)" & _
                strWherePriceGrade & vbNewLine & _
        " Group By A.����,A.����,A.���,A.���㵥λ,A.ID"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strInput & "%", "%" & UCase(strInput) & "%", strInput, _
        strItemType, lngItemID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    
    If rsTmp.RecordCount < 1 Then Exit Sub
    If rsTmp.RecordCount > 1 Then
        strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "����,1200,0,2;����,1800,0,2;���,1200,0,2;���㵥λ,800,0,2;�۸�,1000,1,2;ID,0,1,2", "�շ���Ŀѡ����", True, , , 1000 + 1500 + 1500 + 800 + 800 + 2000)
        If Trim(strReturn) = "" Then
            Exit Sub
        End If
    Else
        strReturn = Nvl(rsTmp!����) & "," & Nvl(rsTmp!����) & "," & Nvl(rsTmp!���) & "," & Nvl(rsTmp!���㵥λ) & "," & Nvl(rsTmp!�۸�) & "," & Nvl(rsTmp!ID, 0)
    End If
    
    With vsfItemList
        '����Ƿ��ظ�
        For i = 0 To .Rows - 1
            If Val(.TextMatrix(i, ��Ŀ�б�.��Ŀid)) = CLng(Split(strReturn, ",")(UBound(Split(strReturn, ",")))) Then
                Exit Sub
            End If
        Next
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, ��Ŀ�б�.����) = Split(strReturn, ",")(0)
        .TextMatrix(.Rows - 1, ��Ŀ�б�.����) = Split(strReturn, ",")(1)
        .TextMatrix(.Rows - 1, ��Ŀ�б�.���) = Split(strReturn, ",")(2)
        .TextMatrix(.Rows - 1, ��Ŀ�б�.��λ) = Split(strReturn, ",")(3)
        .TextMatrix(.Rows - 1, ��Ŀ�б�.�۸�) = Format(Val(Split(strReturn, ",")(4)), "###0.000;-##0.000;0.000;0.000")
        .TextMatrix(.Rows - 1, ��Ŀ�б�.��Ŀid) = Split(strReturn, ",")(5)
        
        '�����ؼ���С
        If .Rows > 3 And .Rows < 11 And .Top + .RowHeightMin * .Rows + 50 > fraItem.Height And UdStage.value > 5 Then
            Me.Height = Me.Height + (.Rows - 3) * .RowHeightMin
            .Height = .Height + (.Rows - 3) * .RowHeightMin
            fraItem.Height = fraItem.Height + (.Rows - 3) * .RowHeightMin
            
            If fra�ѱ�.Height < .Height Then
                fra�ѱ�.Height = fraItem.Height
                cmdHelp.Top = Me.Height - cmdHelp.Height - 500
                cmdOK.Top = cmdHelp.Top
                cmdCancel.Top = cmdOK.Top
            End If
        End If
        
        If .Rows > 2 Then
            lblItem.Caption = "[" & .TextMatrix(1, ��Ŀ�б�.����) & "��]" & "�ֶ�����"
        ElseIf .Rows = 2 Then
            lblItem.Caption = "[" & .TextMatrix(1, ��Ŀ�б�.����) & "]" & "�ֶ�����"
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveCharge() As Boolean
    Dim str���� As String
    Dim curStart As Currency, curEnd As Currency, dblTax As Double
    Dim i As Long
    Dim blnTrans As Boolean
    Dim intӦ�� As Integer
    
    curStart = Val(Me.txtMoney(0).Text)
    dblTax = Val(Me.txtTax(0).Text)
    
    Err = 0
    On Error GoTo ErrHand
    
    For mintStage = 0 To Me.UdStage.value - 1
        curStart = Val(Me.txtMoney(mintStage).Text)
        If mintStage >= Me.UdStage.value - 1 Then
            curEnd = Val("10000000000.00")
        Else
            curEnd = Val(Me.txtMoney(mintStage + 1).Text) - 0.01
        End If
        dblTax = Val(Me.txtTax(mintStage).Text)
        str���� = str���� & mintStage + 1 & ":" & curStart & ":" & curEnd & ":" & dblTax & ";"
    Next
    
    gcnOracle.BeginTrans
    blnTrans = False
    
    If mintType = 0 Then
        gstrSQL = "zl_�ѱ���ϸ_update('" & mstrGrade & "'," & mlngItemId & ",'" & str���� & "'," & Val(cbo���㷽��.Text) & "," & mintType & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    ElseIf mintType = 1 Then
        '���������շ���Ŀ
        For i = 1 To vsfItemList.Rows - 1
            gstrSQL = "zl_�ѱ���ϸ_update('" & mstrGrade & "'," & Val(vsfItemList.TextMatrix(i, ��Ŀ�б�.��Ŀid)) & ",'" & str���� & "'," & Val(cbo���㷽��.Text) & "," & mintType & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
    ElseIf mintType = 2 Then
        '�շ���ĿĿ¼�����÷ѱ�
        If optApply(0).value = True Then
            intӦ�� = 0
        ElseIf optApply(1).value = True Then
            intӦ�� = 1
        ElseIf optApply(2).value = True Then
            intӦ�� = 2
        ElseIf optApply(3).value = True Then
            intӦ�� = 3
        End If
        
        gstrSQL = "zl_�ѱ���ϸ_update('" & mstrGrade & "'," & mlngItemId & ",'" & str���� & "'," & Val(cbo���㷽��.Text) & "," & mintType & "," & intӦ�� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    ElseIf mintType = 3 Then
        'ҩƷĿ¼�����÷ѱ�
        If optӦ����(0).value = True Then
            intӦ�� = 0
        ElseIf optӦ����(1).value = True Then
            intӦ�� = 1
        ElseIf optӦ����(2).value = True Then
            intӦ�� = 2
        ElseIf optӦ����(3).value = True Then
            intӦ�� = 3
        ElseIf optӦ����(4).value = True Then
            intӦ�� = 4
        ElseIf optӦ����(5).value = True Then
            intӦ�� = 5
        End If
        
        gstrSQL = "zl_�ѱ���ϸ_update('" & mstrGrade & "'," & mlngItemId & ",'" & str���� & "'," & Val(cbo���㷽��.Text) & "," & mintType & "," & intӦ�� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    gcnOracle.CommitTrans
    
    mblnChange = False
    mblnOk = True
    blnTrans = True
    
    SaveCharge = True
    Exit Function
ErrHand:
    If Not blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function

Public Function ShowMe(objfrm As Object, ByVal int���� As Integer, ByVal str�ѱ� As String, ByVal lng��Ŀid As Long, ByVal str��Ŀ���� As String) As Boolean
    mintType = int����
    mstrGrade = str�ѱ�
    mlngItemId = lng��Ŀid
    mStrItem = str��Ŀ����
    
    Me.Show vbModal, objfrm
    
    ShowMe = mblnOk
End Function

Private Sub LoadCharge()
    Dim rsTemp As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ���� From �ѱ� Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ѱ�")
    
    cbo�ѱ�.Clear
    
    With rsTemp
        Do While Not .EOF
            cbo�ѱ�.AddItem !����
            
            If !���� = mstrGrade Then
                intIndex = cbo�ѱ�.ListCount - 1
            End If
            
            .MoveNext
        Loop
    End With
    
    If cbo�ѱ�.ListCount > 0 Then
        If mintType = 0 Or mintType = 1 Then
            cbo�ѱ�.Enabled = False
        End If
        cbo�ѱ�.ListIndex = intIndex
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadItemType()
    Dim rsTemp As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ����||'-'||���� As ���� From �շ���Ŀ��� Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ŀ���")
    
    cbo��Ŀ���.Clear
    cbo��Ŀ���.AddItem "0-�������"
    With rsTemp
        Do While Not .EOF
            cbo��Ŀ���.AddItem !����
            
            .MoveNext
        Loop
    End With
    
    If cbo��Ŀ���.ListCount > 0 Then
       cbo��Ŀ���.ListIndex = 0
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo�ѱ�_Click()
    mstrGrade = cbo�ѱ�.List(cbo�ѱ�.ListIndex)
    Call LoadChargeList(mstrGrade, mlngItemId)
End Sub


Private Sub cbo���㷽��_Click()
    '1-�ɱ��ۼ��ձ�������,���ֶ�
    If cbo���㷽��.ListIndex = 1 Then
        txtStage.Text = 1
        UdStage.value = 1
        txtStage.Enabled = False
        UdStage.Enabled = False
        lblNote.Caption = "  ҩƷʵ�ս��=�ɱ���*(1+���ձ���)���������ҩƷ�����Դ����ã������ۡ�"
        lblMoney.Caption = "�ֶ����"
        lblTax.Caption = "���ձ���(%)"
    '0-�ֶα�������
    Else
       txtStage.Enabled = True
       UdStage.Enabled = True
       lblNote.Caption = "    ÿһ������Ŀ�ɰ�Ӧ�ս���Ϊ���(���16��)�����ò�ͬ��ʵ�ձ�����"
       lblMoney.Caption = "Ӧ�շֶ����"
       lblTax.Caption = "ʵ�ձ���(%)"
    End If
    
End Sub

Private Sub cmdDel_Click()
    Dim intӦ�� As Integer
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    If mintType = 1 Then
        With vsfItemList
            If .Rows = 1 Then
                If mStrItem <> "" Then
                    If MsgBox("�Ƿ����[" & mStrItem & "]��Ŀ�ķѱ����ã�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    
                    gstrSQL = "zl_�ѱ���ϸ_update('" & mstrGrade & "'," & mlngItemId & ",Null,0," & mintType & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
            ElseIf .Rows = 2 Then
                If MsgBox("�Ƿ����[" & .TextMatrix(1, ��Ŀ�б�.����) & "]��Ŀ�ķѱ����ã�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                
                gstrSQL = "zl_�ѱ���ϸ_update('" & mstrGrade & "'," & Val(.TextMatrix(1, ��Ŀ�б�.��Ŀid)) & ",Null,0," & mintType & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Else
                If MsgBox("�Ƿ����[" & .TextMatrix(1, ��Ŀ�б�.����) & "]����Ŀ�ķѱ����ã�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                For i = 1 To .Rows - 1
                    gstrSQL = "zl_�ѱ���ϸ_update('" & mstrGrade & "'," & Val(.TextMatrix(i, ��Ŀ�б�.��Ŀid)) & ",Null,0," & mintType & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                Next
            End If
        End With
    ElseIf mintType = 2 Or mintType = 3 Then
        If mintType = 2 Then
            If optApply(0).value = True Then
                intӦ�� = 0
            ElseIf optApply(1).value = True Then
                intӦ�� = 1
            ElseIf optApply(2).value = True Then
                intӦ�� = 2
            ElseIf optApply(3).value = True Then
                intӦ�� = 3
            End If
        Else
            If optӦ����(0).value = True Then
                intӦ�� = 0
            ElseIf optӦ����(1).value = True Then
                intӦ�� = 1
            ElseIf optӦ����(2).value = True Then
                intӦ�� = 2
            ElseIf optӦ����(3).value = True Then
                intӦ�� = 3
            ElseIf optӦ����(4).value = True Then
                intӦ�� = 4
            ElseIf optӦ����(5).value = True Then
                intӦ�� = 5
            End If
        End If
        
        If MsgBox("�Ƿ����[" & mStrItem & "]��Ӧ�÷�Χ��������Ŀ�ķѱ����ã�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        gstrSQL = "zl_�ѱ���ϸ_update('" & mstrGrade & "'," & mlngItemId & ",Null,0," & mintType & "," & intӦ�� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
    End If
    
    MsgBox "����ɹ���", vbExclamation, gstrSysName
    If mintType = 1 Then
        Call IniChargeList
        Call IniItemList
    Else
        Call IniChargeList
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdFilter_Click()
'    If Trim(txtInput.Text) = "" Then Exit Sub
    
    Call GetItemList(txtInput.Text, Mid(cbo��Ŀ���.List(cbo��Ŀ���.ListIndex), 1, 1))
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdMove_Click()
    With vsfItemList
        If .Row > 0 Then
            .RemoveItem .Row
        End If
        If .Rows > 2 Then
            lblItem.Caption = "[" & .TextMatrix(1, ��Ŀ�б�.����) & "��]" & "�ֶ�����"
        ElseIf .Rows = 2 Then
            lblItem.Caption = "[" & .TextMatrix(1, ��Ŀ�б�.����) & "]" & "�ֶ�����"
        End If
    End With
End Sub

Private Sub cmdMoveAll_Click()
    lblItem.Caption = "�ֶ�����"
    Call IniItemList
End Sub

Private Sub Form_Load()
    mblnOk = False
    
    Me.Height = mcstFormHeight
    Me.Width = mcstFormWidth
    
    fra�ѱ�.Height = mcstFormChargeHeight
    
    'ByZT20030722
    If glngSys Like "8??" Then
        Caption = "��Ա�ȼ������շ�����"
    End If
    
    '���㷽��
    cbo���㷽��.AddItem "0-�ֶα�������", 0
    cbo���㷽��.AddItem "1-�ɱ��ۼ��ձ�������", 1
    
    Call GetPriceGrade(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    'ȡ�ѱ�
    Call LoadCharge
    
    'ȡ�ѱ���ϸ
    Call LoadChargeList(mstrGrade, mlngItemId)
    
    fraItem.Visible = False
    fra��ĿӦ��.Visible = False
    fraҩƷӦ��.Visible = False
    cmdDel.value = False
    
    If mintType = 0 Then
    ElseIf mintType = 1 Then
        Me.Width = Me.Width + fraItem.Width + 100
        fraItem.Visible = True
        fraItem.Top = fra�ѱ�.Top
        fraItem.Left = fra�ѱ�.Left + fra�ѱ�.Width + 100
        fraItem.Height = fra�ѱ�.Height
        cmdDel.Visible = True
        
        'ȡ��Ŀ���
        Call LoadItemType
        
        '��ʼ��Ŀ�б�
        Call IniItemList
        
        '�����������ĿID������ȡ����Ŀ��Ϣ
        If mlngItemId > 0 Then
            Call GetItemList("", "", mlngItemId)
        End If
    ElseIf mintType = 2 Then
        Me.Width = Me.Width + fra��ĿӦ��.Width + 100
        fra��ĿӦ��.Visible = True
        fra��ĿӦ��.Top = fra�ѱ�.Top
        fra��ĿӦ��.Left = fra�ѱ�.Left + fra�ѱ�.Width + 100
        fra��ĿӦ��.Height = fra�ѱ�.Height
        cmdDel.Visible = True
    ElseIf mintType = 3 Then
        Me.Width = Me.Width + fraҩƷӦ��.Width + 100
        fraҩƷӦ��.Visible = True
        fraҩƷӦ��.Top = fra�ѱ�.Top
        fraҩƷӦ��.Left = fra�ѱ�.Left + fra�ѱ�.Width + 100
        fraҩƷӦ��.Height = fra�ѱ�.Height
        cmdDel.Visible = True
        
        'ȡҩƷ���ʣ�������Ϣ
        Call GetDrugOtherInfo
    End If
    cmdOK.Left = Me.Width - cmdCancel.Width - cmdOK.Width - 240
    cmdCancel.Left = cmdOK.Left + cmdOK.Width
    cmdDel.Left = cmdOK.Left - cmdDel.Width - 250
End Sub

Private Sub LoadChargeList(ByVal str�ѱ� As String, ByVal lng��Ŀid As Long)
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    If mintType = 0 Then
        strSQL = " And ������Ŀid=[2]"
    Else
        strSQL = " And �շ�ϸĿid=[2]"
    End If
    
    On Error GoTo ErrHandle
    gstrSQL = "Select �κ�, Ӧ�ն���ֵ, Ӧ�ն�βֵ, ʵ�ձ���, ���㷽�� " & _
        " From �ѱ���ϸ Where �ѱ� = [1] " & strSQL & " Order By �κ�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ѱ���ϸ", str�ѱ�, lng��Ŀid)

    If rsTemp.RecordCount = 0 Then
        Call IniChargeList
        Exit Sub
    End If
    
    cbo���㷽��.ListIndex = IIF(rsTemp!���㷽�� = 0, 0, 1)
    
    With rsTemp
        txtStage.Text = .RecordCount
        UdStage.value = .RecordCount
        cbo���㷽��.ListIndex = Val(.Fields("���㷽��").value)     '����Click�¼�������ؿؼ�
        lblItem.Caption = "[" & mStrItem & "]" & "�ֶ�����"
        
        For i = 1 To .RecordCount
            If i > 16 Then Exit For
            
            lblNo(.AbsolutePosition - 1).Visible = True
            lblNo(.AbsolutePosition - 1).Caption = .AbsolutePosition
            txtMoney(.AbsolutePosition - 1).Visible = True
            txtMoney(.AbsolutePosition - 1).Text = Format(.Fields("Ӧ�ն���ֵ").value, "###########0.00;-##########0.00;0.00;0.00")
            txtTax(.AbsolutePosition - 1).Visible = True
            txtTax(.AbsolutePosition - 1).Text = Format(.Fields("ʵ�ձ���").value, "###0.000;-##0.000;0.000;0.000")
            
            .MoveNext
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub IniChargeList()
    cbo���㷽��.ListIndex = 0
    UdStage.Enabled = True
    UdStage.value = 1
    
    lblNo(0).Visible = True
    txtMoney(0).Visible = True
    txtTax(0).Visible = True
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
'    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
'        Cancel = 1
'    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    'У��
    If IsValidate = False Then Exit Sub
    If mintType = 2 Then
        If optApply(0).value = False Then
            For i = 0 To optApply.UBound
                If optApply(i).value = True Then
                    If MsgBox("�ѱ�����Ӧ�÷�ΧΪ��" & optApply(i).Caption & "���Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    '����
    If SaveCharge = False Then Exit Sub
    
    If mintType = 0 Then
        MsgBox "���óɹ���", vbExclamation, gstrSysName
        Call IniChargeList
    ElseIf mintType = 1 Then
        MsgBox "���óɹ���", vbExclamation, gstrSysName
        Call IniChargeList
        Call IniItemList
    Else
        Unload Me
    End If
End Sub

Private Function IsValidate() As Boolean
    Dim intStage As Integer
    Dim str���� As String
    Dim curStart As Currency, dblTax As Double
    Dim curStartBefore As Currency
    Dim dblTaxBefore As Double
    Dim i As Long
    
    If mintType = 1 And vsfItemList.Rows = 1 Then Exit Function
        
    For intStage = 1 To Me.UdStage.value - 1
        curStart = Val(Me.txtMoney(intStage).Text)
        dblTax = Val(Me.txtTax(intStage).Text)

        If curStart <= Val(Me.txtMoney(intStage - 1).Text) Then
            MsgBox "��" & intStage + 1 & "�δ���Ӧ�ն�ֵ������С����", vbExclamation, gstrSysName
            txtMoney(intStage).SetFocus
            Exit Function
        End If
        If dblTax = Val(Me.txtTax(intStage - 1).Text) Then
            MsgBox "��" & intStage + 1 & "�δ������ڶ�ʵ�ձ�����ͬ�������塣", vbExclamation, gstrSysName
            txtTax(intStage).SetFocus
            Exit Function
        End If
    Next
    
    IsValidate = True
End Function

Private Sub optApply_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To optApply.UBound
        If i = Index Then
            optApply(i).FontBold = True
        Else
            optApply(i).FontBold = False
        End If
    Next
End Sub

Private Sub optӦ����_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To optӦ����.UBound
        If i = Index Then
            optӦ����(i).FontBold = True
        Else
            optӦ����(i).FontBold = False
        End If
    Next
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
'    If Trim(txtInput.Text) = "" Then Exit Sub
    
    Call GetItemList(txtInput.Text, Mid(cbo��Ŀ���.List(cbo��Ŀ���.ListIndex), 1, 1))
End Sub


Private Sub txtMoney_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtMoney_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtMoney(Index)
End Sub

Private Sub txtMoney_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii > vbKey9 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtMoney_Validate(Index As Integer, Cancel As Boolean)
    Me.txtMoney(Index).Text = Format(Val(Me.txtMoney(Index).Text), "###########0.00;-##########0.00;0.00;0.00")
    If Val(Me.txtMoney(Index).Text) >= Val("10000000000.00") Or Val(Me.txtMoney(Index).Text) < 0 Then
        MsgBox "Ӧ�ս�����ֻ���� 0��10000000000.00֮�䡣", vbExclamation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub txtTax_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtTax_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtTax(Index)
End Sub

Private Sub txtTax_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii > vbKey9 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtTax_Validate(Index As Integer, Cancel As Boolean)
    Me.txtTax(Index).Text = Format(Val(Me.txtTax(Index).Text), "###0.000;-##0.000;0.000;0.000")
    If Val(Me.txtTax(Index).Text) > 500 Or Val(Me.txtTax(Index).Text) < 0 Then
        MsgBox "ʵ�ձ���ֻ���� 0��500֮�䡣", vbExclamation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub UdStage_Change()
    Dim dblRowHeight As Double
    Dim intValue As Integer
    
    intValue = Me.UdStage.value
    dblRowHeight = txtMoney(0).Height
        
    For mintStage = 0 To 15
        Me.lblNo(mintStage).Visible = (Me.UdStage.value > mintStage)
        Me.txtMoney(mintStage).Visible = (Me.UdStage.value > mintStage)
        Me.txtTax(mintStage).Visible = (Me.UdStage.value > mintStage)
    Next
    
    mblnChange = True
     
    If intValue < 4 Then Exit Sub
        
    fra�ѱ�.Height = 2750 + (intValue - 1) * dblRowHeight
    Me.Height = 3905 + (intValue - 1) * dblRowHeight
    cmdHelp.Top = Me.Height - cmdHelp.Height - 500
    cmdOK.Top = cmdHelp.Top
    cmdCancel.Top = cmdHelp.Top
    cmdDel.Top = cmdHelp.Top
    
    If fraItem.Visible = True Then
        fraItem.Height = fra�ѱ�.Height
        vsfItemList.Height = fraItem.Height - vsfItemList.Top - 50
    End If
End Sub

Private Function GetPriceGrade(ByRef strҩƷ�۸�ȼ� As String, _
    ByRef str���ļ۸�ȼ� As String, ByRef str��ͨ�۸�ȼ� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰվ��۸�ȼ�
    '���:
    '����:�۸�ȼ���ȡ�ɹ�����True�����򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    strҩƷ�۸�ȼ� = "": str���ļ۸�ȼ� = "": str��ͨ�۸�ȼ� = ""
    strSQL = "" & _
        "Select Max(Decode(b.�Ƿ�����ҩƷ, 1, �۸�ȼ�, Null)) As ҩƷ�ȼ�," & vbNewLine & _
        "       Max(Decode(b.�Ƿ���������, 1, �۸�ȼ�, Null)) As ���ĵȼ�," & vbNewLine & _
        "       Max(Decode(b.�Ƿ�������ͨ��Ŀ, 1, �۸�ȼ�, Null)) As ��ͨ�ȼ�" & vbNewLine & _
        "From �շѼ۸�ȼ�Ӧ�� A, �շѼ۸�ȼ� B" & vbNewLine & _
        "Where a.�۸�ȼ� = b.���� And a.���� = 0 And a.վ�� = [1]" & vbNewLine & _
        "      And (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�۸�ȼ�", gstrNodeNo)
    If Not rsTemp.EOF Then
        strҩƷ�۸�ȼ� = Nvl(rsTemp!ҩƷ�ȼ�)
        str���ļ۸�ȼ� = Nvl(rsTemp!���ĵȼ�)
        str��ͨ�۸�ȼ� = Nvl(rsTemp!��ͨ�ȼ�)
    End If
    GetPriceGrade = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

