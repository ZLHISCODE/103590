VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceWardReport 
   Caption         =   "����ִ�е���ӡ"
   ClientHeight    =   10815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   Icon            =   "frmAdviceWardReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10815
   ScaleMode       =   0  'User
   ScaleWidth      =   14415
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picParaAdd 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1423
      Left            =   275
      ScaleHeight     =   1440
      ScaleMode       =   0  'User
      ScaleWidth      =   3789.634
      TabIndex        =   29
      Top             =   5880
      Width           =   3729
      Begin VB.PictureBox picPara 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   999
         Left            =   1125
         ScaleHeight     =   495
         ScaleWidth      =   2655
         TabIndex        =   48
         Top             =   -3000
         Width           =   2655
         Begin VB.OptionButton optPara 
            Caption         =   "��˹�ٷҷ�����������"
            Height          =   255
            Index           =   999
            Left            =   0
            TabIndex        =   49
            Top             =   0
            Value           =   -1  'True
            Width           =   2100
         End
      End
      Begin VB.CheckBox chk��Ч 
         Caption         =   "����(&L)"
         Height          =   195
         Index           =   0
         Left            =   1125
         TabIndex        =   36
         Top             =   870
         Value           =   1  'Checked
         Width           =   930
      End
      Begin VB.CheckBox chk��Ч 
         Caption         =   "��ʱ(&T)"
         Height          =   195
         Index           =   1
         Left            =   2250
         TabIndex        =   35
         Top             =   870
         Value           =   1  'Checked
         Width           =   930
      End
      Begin VB.CheckBox chk�ظ���ӡ 
         Caption         =   "�����Ѵ�ӡ����(&A)"
         Height          =   195
         Left            =   1125
         TabIndex        =   34
         Top             =   1170
         Width           =   2295
      End
      Begin VB.CheckBox chkPara 
         Height          =   195
         Index           =   999
         Left            =   1125
         TabIndex        =   33
         Top             =   1335
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtWB 
         Height          =   300
         Index           =   999
         Left            =   1125
         TabIndex        =   32
         Top             =   1530
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox cboPara 
         Height          =   300
         Index           =   999
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1590
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.TextBox txtSZ 
         Height          =   300
         Index           =   999
         Left            =   1125
         TabIndex        =   30
         Top             =   1710
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1125
         TabIndex        =   37
         Top             =   0
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   176685059
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1125
         TabIndex        =   38
         Top             =   360
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   176685059
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtkPara 
         Height          =   300
         Index           =   999
         Left            =   1125
         TabIndex        =   39
         Top             =   1590
         Visible         =   0   'False
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   176685059
         CurrentDate     =   37953
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯʱ��  ��"
         Height          =   180
         Left            =   0
         TabIndex        =   47
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ֹ"
         Height          =   180
         Left            =   900
         TabIndex        =   46
         Top             =   420
         Width           =   180
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Left            =   0
         TabIndex        =   45
         Top             =   870
         Width           =   720
      End
      Begin VB.Label lblSZ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Index           =   999
         Left            =   360
         TabIndex        =   44
         Top             =   1590
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblRQ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Index           =   999
         Left            =   360
         TabIndex        =   43
         Top             =   1830
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblXL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Index           =   999
         Left            =   360
         TabIndex        =   42
         Top             =   1230
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblWB 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Index           =   999
         Left            =   360
         TabIndex        =   41
         Top             =   1470
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblDX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Index           =   999
         Left            =   360
         TabIndex        =   40
         Top             =   1950
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.PictureBox picCMD 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      ScaleHeight     =   615
      ScaleWidth      =   3855
      TabIndex        =   19
      Top             =   9840
      Width           =   3855
      Begin VB.CommandButton cmdPriv 
         Caption         =   "Ԥ  ��(&V)"
         Height          =   495
         Left            =   1320
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdExec 
         Caption         =   "��  ѯ(&Q)"
         Height          =   495
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��  ӡ(&P)"
         Height          =   495
         Left            =   2640
         TabIndex        =   20
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4200
      Top             =   120
   End
   Begin VB.Frame frmModle 
      Caption         =   "��ӡģʽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   7560
      Width           =   3975
      Begin VB.TextBox txt���� 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   1275
         TabIndex        =   14
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txt��λ�� 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   1275
         TabIndex        =   12
         Top             =   1140
         Width           =   2055
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   1275
         TabIndex        =   10
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdPati 
         Caption         =   "ѡ����"
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optPati 
         Caption         =   "���Ҳ���"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optPati 
         Caption         =   "ȫ��������"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��"
         Height          =   180
         Left            =   630
         TabIndex        =   15
         Top             =   1635
         Width           =   540
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ��"
         Height          =   180
         Left            =   630
         TabIndex        =   13
         Top             =   1215
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   630
         TabIndex        =   11
         Top             =   795
         Width           =   540
      End
   End
   Begin VB.Frame frmAnd 
      Caption         =   "�����趨"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7095
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3975
      Begin VB.PictureBox picReport 
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4035
         ScaleMode       =   0  'User
         ScaleWidth      =   3735
         TabIndex        =   23
         Top             =   600
         Width           =   3735
         Begin VB.CommandButton cmdSet 
            Caption         =   "����ִ�е�"
            Height          =   375
            Left            =   0
            TabIndex        =   27
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdSetup 
            Caption         =   "��ӡ����"
            Height          =   330
            Left            =   0
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + S"
            Top             =   360
            Width           =   1095
         End
         Begin MSComctlLib.ListView lvwReport 
            Height          =   4080
            Left            =   1155
            TabIndex        =   25
            Top             =   0
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   7197
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "img16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   6615
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ϵͳ"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ImageList img16 
            Left            =   120
            Top             =   1200
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdviceWardReport.frx":6852
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ִ�е�����"
            Height          =   180
            Left            =   30
            TabIndex        =   26
            Top             =   30
            Width           =   900
         End
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����(&U)"
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   5280
      ScaleHeight     =   4935
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   960
      Width           =   8895
      Begin VB.Frame fraColSel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   195
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frmAdviceWardReport.frx":69AC
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   4035
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   8160
         _cx             =   14393
         _cy             =   7117
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   1
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
      Begin VSFlex8Ctl.VSFlexGrid vsColumn 
         Height          =   3270
         Left            =   840
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1470
         _cx             =   2593
         _cy             =   5768
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAdviceWardReport.frx":6EFA
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5265
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   9287
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   10455
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceWardReport.frx":6F48
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22781
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   2
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
   Begin MSComctlLib.ImageList imgAdvice 
      Left            =   12720
      Top             =   120
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
            Picture         =   "frmAdviceWardReport.frx":77DC
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceWardReport.frx":7D76
            Key             =   "AllCheck"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceWardReport.frx":8310
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblAutoSize 
      AutoSize        =   -1  'True
      Caption         =   "lblAutoSize"
      Height          =   180
      Left            =   2040
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "frmAdviceWardReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1

Private mMainPrivs As String 'IN:���������������е�Ȩ��,ע����ڲ�ģ��Ȩ��
Private mlng����ID As Long 'IN
Private mlng����ID As Long 'IN
Private mobjDatas As Object
Private mintAdvice As Integer
Private mintIDIndex As Integer
Private mint���IDIndex As Integer
Private mstrIDName As String
Private mblnZPriw As Boolean 'ֱ��Ԥ���������»�ȡ
Private mstrFilter As String '����ظ�Ԥ��
Private mstr����IDs As String
Private mstrPrintedID As String
Private mlngLastRow As Long
Private mlngPre����ID As Long
Private mvarPar As Variant
Private Const colѡ�� = 1
Private mobjReportPrivw As Object

Public Function ShowMe(frmParent As Object, ByVal MainPrivs As String, _
    ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
'������
    mMainPrivs = MainPrivs
    
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mstrPrintedID = ""
        
    Me.Show , frmParent
End Function

Private Function InitReports() As Boolean
'���ܣ���ȡ���ñ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objItem As ListItem
    Dim strReports As String
    
    On Error GoTo errH
    lvwReport.ListItems.Clear
    
    strReports = zldatabase.GetPara("ִ�е����ñ���", glngSys, pסԺҽ������, , , , , cboUnit.ItemData(cboUnit.ListIndex))
    strSQL = "Select A.ID,A.���,A.����,NVL(A.����,B.����) AS ����,a.ϵͳ" & vbNewLine & _
            "From zlReports A, zlRPTPuts B" & vbNewLine & _
            "Where a.Id = b.����id(+)  And" & vbNewLine & _
            "      (b.����id = 1254 Or" & vbNewLine & _
            "       a.ϵͳ = [1] And a.��� In ('ZL1_INSIDE_1254_4', 'ZL1_INSIDE_1254_5', 'ZL1_INSIDE_1254_6', 'ZL1_INSIDE_1254_7', 'ZL1_INSIDE_1254_8'," & vbNewLine & _
            "                'ZL1_INSIDE_1254_9', 'ZL1_INSIDE_1254_10', 'ZL1_INSIDE_1254_11', 'ZL1_INSIDE_1254_12'," & vbNewLine & _
            "                'ZL1_INSIDE_1254_13', 'ZL1_INSIDE_1254_14', 'ZL1_INSIDE_1254_15', 'ZL1_INSIDE_1254_16'))" & vbNewLine & _
            "Order By a.Id"

    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, glngSys)
    Do While Not rsTmp.EOF
        If InStr(GetInsidePrivs(pסԺҽ������), ";" & rsTmp!���� & ";") > 0 Then
            If strReports = "" Or InStr("|" & strReports & "|", "|" & rsTmp!��� & "|") > 0 Then
                Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!���, rsTmp!����, , 1)
                objItem.SubItems(1) = Val(rsTmp!ϵͳ & "")
                objItem.Tag = Val(rsTmp!ID)
            End If
        End If
        rsTmp.MoveNext
    Loop
    InitReports = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '��������۲���
    If InStr(mMainPrivs, "ȫԺ����") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSQL = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
    End If
    
    cboUnit.Clear
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng����ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetColSortIndex(arrSort As Variant, strName As String) As Long
'���ܣ���ȡ�ϴα���������к�
    Dim i As Long
    
    If UBound(arrSort) = -1 Or strName = "" Then GetColSortIndex = -1: Exit Function
    For i = 0 To UBound(arrSort)
        If arrSort(i) = strName Then Exit For
    Next
    GetColSortIndex = i
End Function

Private Sub cboUnit_Click()
    Call InitReports
    tmrLoad.Enabled = True
End Sub

Private Sub cmdExec_Click()
    Dim strReports As String
    Dim curDate As Date, datBegin As Date, datEnd As Date
    Dim i As Long, j As Long, X As Long, objControl As Object
    Dim str���� As String, str��Ч As String, str�ظ���ӡ As String
    Dim objDatas As Object
    Dim rsData As Recordset, strColHidden As String, strColSort As String, strColSortEnd As String
    Dim arrSort As Variant, lngCount As Long, lngҽ������ As Long, lng���ID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim arrpara As Variant  '�Զ������
    Dim strValues(0 To 10) As String, strSQL As String, rsTmp As Recordset, rsSQL As Recordset
    
    '��������
    curDate = zldatabase.Currentdate
    Call zldatabase.SetPara("���ñ�����Ч", chk��Ч(0).Value & chk��Ч(1).Value, glngSys, pסԺҽ������)
    Call zldatabase.SetPara("���ñ���ʼʱ��", Format(dtpBegin.Value, "HH:mm:ss"), glngSys, pסԺҽ������)
    Call zldatabase.SetPara("���ñ���ʼ���", Int(CDate(Format(dtpBegin.Value, "yyyy-MM-dd")) - CDate(Format(curDate, "yyyy-MM-dd"))), glngSys, pסԺҽ������)
    Call zldatabase.SetPara("���ñ������ʱ��", Format(dtpEnd.Value, "HH:mm:ss"), glngSys, pסԺҽ������)
    Call zldatabase.SetPara("���ñ���������", Int(CDate(Format(dtpEnd.Value, "yyyy-MM-dd")) - CDate(Format(curDate, "yyyy-MM-dd"))), glngSys, pסԺҽ������)
    
    datBegin = dtpBegin.Value: datEnd = dtpEnd.Value
    '��Ч����
    If chk��Ч(0).Value = 1 And chk��Ч(1).Value = 1 Then
        str��Ч = "0,1"
    ElseIf chk��Ч(0).Value = 1 Then
        str��Ч = "0"
    Else
        str��Ч = "1"
    End If
    
    str�ظ���ӡ = IIF(chk�ظ���ӡ.Visible, chk�ظ���ӡ.Value, 0)
    If mstr����IDs <> "" Then
        str���� = "Select /*+cardinality(b,10)*/" & vbNewLine & _
                " C1 As ����id, C2 As ��ҳid" & vbNewLine & _
                "From Table(Cast(f_Str2list2('" & mstr����IDs & "') As Zltools.t_Strlist2)) B"
    Else
        str���� = "Select ����ID,��ҳID from ��Ժ���� Where ����ID=" & cboUnit.ItemData(cboUnit.ListIndex)
    End If
    cmdExec.Tag = "1"
    
    arrpara = Split("�Զ������1=0,�Զ������2=0,�Զ������3=0,�Զ������4=0,�Զ������5=0,�Զ������6=0,�Զ������7=0,�Զ������8=0,�Զ������9=0,�Զ������10=0", ",")
    
    '��ѡ
    For i = 0 To chkPara.Count - 2
        If i < 10 Then
            arrpara(i) = chkPara(i).Caption & "=" & IIF(chkPara(i).Value = 1, Split(chkPara(i).Tag, "|")(0), Split(chkPara(i).Tag, "|")(1))
            j = i
        End If
    Next
    '��ѡ
    For i = 0 To picPara.Count - 2
        j = j + 1
        If j < 10 Then
            For Each objControl In Me.Controls
                If TypeName(objControl) = "OptionButton" Then
                    If objControl.Container = picPara(i) Then
                        If objControl.Value = True Then
                            Exit For
                        End If
                    End If
                End If
            Next
            arrpara(j) = lblDX(i).Caption & "=" & objControl.Tag
        End If
    Next
    '����
    For i = 0 To cboPara.Count - 2
        j = j + 1
        If j < 10 Then
            arrpara(j) = lblXL(i).Caption & "=" & cboPara(i).ItemData(cboPara(i).ListIndex)
        End If
    Next
    '�ı�
    For i = 0 To txtWB.Count - 2
        j = j + 1
        If j < 10 Then
            
            If Trim(txtWB(i).Text) = "" Then
                MsgBox "������""" & lblWB(i).Caption & """������ֵ��", vbInformation, App.Title
                If txtWB(i).Enabled Then txtWB(i).SetFocus
                Exit Sub
            End If
            If Len(txtWB(i).Text) > 4000 Then
                MsgBox """" & lblWB(i).Caption & """������ֵ���Ȳ��ܳ���4000���ַ���", vbInformation, App.Title
                If txtWB(i).Enabled Then txtWB(i).SetFocus
                Exit Sub
            End If
                    
            arrpara(j) = lblWB(i).Caption & "=" & txtWB(i).Text
        End If
    Next
    '����
    For i = 0 To txtSZ.Count - 2
        j = j + 1
        If j < 10 Then
        
            If Trim(txtSZ(i).Text) = "" Then
                MsgBox "������""" & lblSZ(i).Caption & """������ֵ��", vbInformation, App.Title
                If txtSZ(i).Enabled Then txtSZ(i).SetFocus
                Exit Sub
            End If
            If Len(txtSZ(i).Text) > 4000 Then
                MsgBox """" & lblSZ(i).Caption & """������ֵ���Ȳ��ܳ���4000���ַ���", vbInformation, App.Title
                If txtSZ(i).Enabled Then txtSZ(i).SetFocus
                Exit Sub
            End If
            If Not IsNumeric(txtSZ(i).Text) Then
                MsgBox """" & lblSZ(i).Caption & """������ֵ����Ӧ��Ϊ�����ͣ�", vbInformation, App.Title
                If txtSZ(i).Enabled Then txtSZ(i).SetFocus
                Exit Sub
            End If
            
            arrpara(j) = lblSZ(i).Caption & "=" & txtSZ(i).Text
        End If
    Next
    '����
    For i = 0 To dtkPara.Count - 2
        j = j + 1
        If j < 10 Then
            arrpara(j) = lblRQ(i).Caption & "=" & Format(dtkPara(i).Value, "yyyy-MM-dd HH:mm:ss")
        End If
    Next

    mvarPar = Array()
    ReDim Preserve mvarPar(UBound(mvarPar) + 1)
    mvarPar(UBound(mvarPar)) = "��ʼʱ��=" & Format(datBegin, "yyyy-MM-dd HH:mm:ss")
    ReDim Preserve mvarPar(UBound(mvarPar) + 1)
    mvarPar(UBound(mvarPar)) = "����ʱ��=" & Format(datEnd, "yyyy-MM-dd HH:mm:ss")
    ReDim Preserve mvarPar(UBound(mvarPar) + 1)
    mvarPar(UBound(mvarPar)) = "��Ч=" & str��Ч
    ReDim Preserve mvarPar(UBound(mvarPar) + 1)
    mvarPar(UBound(mvarPar)) = "����=" & str����
    ReDim Preserve mvarPar(UBound(mvarPar) + 1)
    mvarPar(UBound(mvarPar)) = "�ظ���ӡ=" & str�ظ���ӡ
    ReDim Preserve mvarPar(UBound(mvarPar) + 1)
    mvarPar(UBound(mvarPar)) = "����ID=" & lvwReport.SelectedItem.Tag
    ReDim Preserve mvarPar(UBound(mvarPar) + 1)
    mvarPar(UBound(mvarPar)) = "PressWorkFirst=0"
    ReDim Preserve mvarPar(UBound(mvarPar) + 1)
    mvarPar(UBound(mvarPar)) = "��ʼ�к�=1"
    ReDim Preserve mvarPar(UBound(mvarPar) + 1)
    mvarPar(UBound(mvarPar)) = "������ID=0"

    For i = 0 To 9
        ReDim Preserve mvarPar(UBound(mvarPar) + 1)
        mvarPar(UBound(mvarPar)) = arrpara(i)
    Next
    
    Call mobjReport.LoadReport(gcnOracle, Val(lvwReport.SelectedItem.SubItems(1)), Mid(lvwReport.SelectedItem.Key, 2), Me, mobjReportPrivw, objDatas, _
        CStr(mvarPar(0)), CStr(mvarPar(1)), CStr(mvarPar(2)), CStr(mvarPar(3)), CStr(mvarPar(4)), CStr(mvarPar(5)), CStr(mvarPar(6)), CStr(mvarPar(7)), CStr(mvarPar(8)), CStr(mvarPar(9)), _
        CStr(mvarPar(10)), CStr(mvarPar(11)), CStr(mvarPar(12)), CStr(mvarPar(13)), CStr(mvarPar(14)), CStr(mvarPar(15)), CStr(mvarPar(16)), CStr(mvarPar(17)), CStr(mvarPar(18)), 1)
        
    If mobjReportPrivw Is Nothing Then MsgBox "����δ��ȡ�ɹ�������ϵ����Ա��", vbInformation, Me.Caption: Exit Sub
    Me.tbcSub.RemoveItem 1
    Me.tbcSub.InsertItem(1, lvwReport.SelectedItem.Text & "-��ӡԤ��", mobjReportPrivw.hwnd, 0).Tag = "Ԥ��"
    
    Set mobjDatas = objDatas     
     
    For i = 1 To objDatas.Count
        If objDatas(i).DataName = "ҽ������" Or objDatas(i).DataName Like "*ҽ����¼*" Then
            Set rsData = objDatas(i).dataset
            mintAdvice = i
            Exit For
        End If
    Next
    With vsAdvice
        .Clear
        .Rows = 1: .Cols = 1: .ColWidth(0) = 230
        .ExplorerBar = flexExSortShow
        
        strColHidden = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & "����������" & "", "_" & lvwReport.SelectedItem.Tag, "")
        strColSort = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & "������˳��" & "", "_" & lvwReport.SelectedItem.Tag, "")
        mintIDIndex = 0: mint���IDIndex = 0: mstrIDName = ""
        
        If Not rsData Is Nothing Then
            If rsData.RecordCount > 0 Then
                rsData.MoveFirst
                
                '���������ˣ�ȥ�����ֶΣ���Ҫ���������ֶ�Ҳȥ��
                If strColSort <> "" Then
                    For i = 0 To rsData.Fields.Count - 1
                        strColSortEnd = strColSortEnd & "|" & rsData.Fields(i).Name
                    Next
                    strColSortEnd = Mid(strColSortEnd, 2)
                    arrSort = Split(strColSort, "|")
                    strColSort = ""
                    For i = 0 To UBound(arrSort)
                        If arrSort(i) <> "" Then
                            If InStr("|" & strColSortEnd & "|", "|" & arrSort(i) & "|") > 0 Then strColSort = strColSort & "|" & arrSort(i)
                        End If
                    Next
                    strColSort = Mid(strColSort, 2)
                End If
                arrSort = Split(strColSort, "|")
                
                .Cols = rsData.Fields.Count + 2
                .ColDataType(colѡ��) = flexDTBoolean
                lngCount = 1
                .Cell(flexcpPicture, 0, colѡ��) = imgAdvice.ListImages("AllCheck").Picture
                .Cell(flexcpPictureAlignment, 0, colѡ��) = flexPicAlignCenterCenter
                .ColData(colѡ��) = "Check"
                
                For j = 2 To rsData.Fields.Count + 1
                    X = GetColSortIndex(arrSort, rsData.Fields(j - 2).Name)
                    If X = -1 Then
                        X = UBound(arrSort) + 2 + lngCount: lngCount = lngCount + 1
                    Else
                        X = X + 2
                    End If
                    .TextMatrix(0, X) = rsData.Fields(j - 2).Name
                    .ColData(X) = j - 2
                    If rsData.Fields(j - 2).Name = "ҽ������" Then lngҽ������ = X
                    If rsData.Fields(j - 2).Name = "���ID" Then lng���ID = X: mint���IDIndex = lng���ID
                    If rsData.Fields(j - 2).Name = "ID" Or rsData.Fields(j - 2).Name = "ҽ��ID" Then
                        mintIDIndex = X
                        If rsData.Fields(j - 2).Name = "ҽ��ID" Then
                            mstrIDName = "ҽ��ID"
                        Else
                            mstrIDName = "ID"
                        End If
                    End If
                    If UCase(.TextMatrix(0, X)) Like "*ID*" Then .ColHidden(X) = True
                    '�����ϴ�����Ч��
                    If strColHidden <> "" Then
                        If InStr(strColHidden, "|" & .TextMatrix(0, X) & "|") > 0 Then .ColHidden(X) = True
                    End If
                Next
                For i = 1 To rsData.RecordCount
                    .Rows = .Rows + 1
                    .Cell(flexcpChecked, i, colѡ��) = 1
                    For j = 2 To rsData.Fields.Count + 1
                        If j = mintIDIndex And lngҽ������ > 0 Then SetAdviceID strValues, rsData.Fields(.ColData(j)).Value & ""
                        .TextMatrix(i, j) = rsData.Fields(.ColData(j)).Value & ""
                    Next
                    rsData.MoveNext
                Next        
                
                If .FixedRows >= 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '��ҪDraw֮�����Ч
                .Redraw = flexRDDirect
                If lngҽ������ > 0 Then
                    For i = 0 To UBound(strValues)
                        If strValues(i) <> "" Then
                            strSQL = "Select ҽ��ID,�ϴδ�ӡʱ��,����ID" & vbNewLine & _
                                    "From ҽ��ִ�д�ӡ A" & vbNewLine & _
                                    "Where a.ҽ��id In (Select /*+cardinality(b,10)*/" & vbNewLine & _
                                    "                  Column_Value" & vbNewLine & _
                                    "                 From Table(f_Str2list([3])) b) And a.�ϴδ�ӡʱ�� Between [1] And [2]"
                            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")), Mid(strValues(i), 2))
                            If i = 0 Then
                                Set rsSQL = zldatabase.CopyNewRec(rsTmp)
                            Else
                                If Not rsTmp.EOF Then
                                    Do While Not rsTmp.EOF
                                        rsSQL.AddNew
                                        rsSQL!ҽ��ID = rsTmp!ҽ��ID
                                        rsSQL!�ϴδ�ӡʱ�� = rsTmp!�ϴδ�ӡʱ��
                                        rsSQL!����ID = rsTmp!����ID
                                        rsTmp.MoveNext
                                    Loop
                                    rsSQL.Update
                                End If

                            End If
                        End If
                    Next
                    If Not rsSQL Is Nothing Then
                        For i = 1 To .Rows - 1
                            rsSQL.Filter = "ҽ��ID=" & Val(.TextMatrix(i, mintIDIndex))
                            If rsSQL.RecordCount > 0 Then
                                .Cell(flexcpPicture, i, lngҽ������) = imgAdvice.ListImages("Print").Picture
                            Else
                                If lng���ID > 0 Then
                                    rsSQL.Filter = "ҽ��ID=" & Val(.TextMatrix(i, mint���IDIndex))
                                    If rsSQL.RecordCount > 0 Then
                                        .Cell(flexcpPicture, i, lngҽ������) = imgAdvice.ListImages("Print").Picture
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If

    End With
End Sub

Private Sub SetAdviceID(strValues As Variant, ByVal lngID As Long)
    Dim i As Long
    
    For i = 0 To UBound(strValues)
        If Len(strValues(i) & "," & lngID) < 4000 Then
            strValues(i) = strValues(i) & "," & lngID
            Exit Sub
        End If
    Next
End Sub

Private Sub cmdPati_Click()
    Dim str����IDs As String
    Dim str���� As String, strסԺ�� As String, str���� As String
    
    optPati(1).Value = True
    If frmPatiSelect.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex), str����IDs, str����, strסԺ��, str����) Then
        mstr����IDs = str����IDs
        txt��λ��.Text = str����
        txt����.Text = str����
        txtסԺ��.Text = strסԺ��
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim i As Long
    Dim strFilter As String
    
    With vsAdvice
        If cmdExec.Tag = "" Then
            '�Ȳ�ѯ
            cmdExec_Click
        End If
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 1) = 1 Then
                strFilter = strFilter & " Or " & mstrIDName & "=" & .TextMatrix(i, mintIDIndex)
            End If
        Next
        If strFilter <> "" Then
            strFilter = Mid(strFilter, 5)
            If mintIDIndex <> 0 Then
                mobjDatas(mintAdvice).dataset.Filter = strFilter
            Else
                MsgBox "ҽ����������Դ��δ����""ID/ҽ��ID""�ֶΣ��޷�����ѡ���ӡ��", vbInformation, Me.Caption
            End If
        End If
        If mobjDatas Is Nothing Then Exit Sub
        If .Rows > 1 And mintAdvice > 0 Then If mobjDatas(mintAdvice).dataset.RecordCount > 0 Then mobjDatas(mintAdvice).dataset.MoveFirst
        Call mobjReport.ReportOpenForRec(gcnOracle, Val(lvwReport.SelectedItem.SubItems(1)), Mid(lvwReport.SelectedItem.Key, 2), Me, mobjDatas, _
        CStr(mvarPar(0)), CStr(mvarPar(1)), CStr(mvarPar(2)), CStr(mvarPar(3)), CStr(mvarPar(4)), CStr(mvarPar(5)), CStr(mvarPar(6)), CStr(mvarPar(7)), CStr(mvarPar(8)), CStr(mvarPar(9)), _
        CStr(mvarPar(10)), CStr(mvarPar(11)), CStr(mvarPar(12)), CStr(mvarPar(13)), CStr(mvarPar(14)), CStr(mvarPar(15)), CStr(mvarPar(16)), CStr(mvarPar(17)), CStr(mvarPar(18)), 2)
    End With
End Sub

Private Sub cmdPriv_Click()
    If tbcSub.ItemCount > 1 Then
        tbcSub.Item(1).Selected = True
    Else
        cmdExec_Click
        mblnZPriw = True
        If tbcSub.ItemCount > 1 Then
            tbcSub.Item(1).Selected = True
        End If
        mblnZPriw = False
    End If
End Sub

Private Sub cmdSet_Click()
    If frmReportPara.ShowMe(Me, cboUnit.ItemData(cboUnit.ListIndex)) Then
        cboUnit_Click
    End If
End Sub

Private Sub cmdSetup_Click()
    Call mobjReport.ReportPrintSet(gcnOracle, glngSys, Mid(lvwReport.SelectedItem.Key, 2), Me)
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim strTmp As String
    Dim lngTmp As Long
    
    Call InitUnits  '��ȡ����
    'Call InitReports '��ȡ����
    'Call InitPara
    mstr����IDs = ""
    If lvwReport.ListItems.Count = 0 Then
        MsgBox "��û��Ȩ�޴�ӡ�κ�һ�ű�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    
    'ȱʡ����ʱ��
    curDate = zldatabase.Currentdate
    
    'ȱʡҽ����Ч
    strTmp = zldatabase.GetPara("���ñ�����Ч", glngSys, pסԺҽ������, "11", Array(chk��Ч(0), chk��Ч(1)))
    chk��Ч(0).Value = Val(Left(strTmp, 1))
    chk��Ч(1).Value = Val(Right(strTmp, 1))

    
    strTmp = zldatabase.GetPara("���ñ���ʼʱ��", glngSys, pסԺҽ������, "00:00:00", Array(dtpBegin))
    lngTmp = Val(zldatabase.GetPara("���ñ���ʼ���", glngSys, pסԺҽ������, "0", Array(dtpBegin)))
    dtpBegin.Value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
    
    strTmp = zldatabase.GetPara("���ñ������ʱ��", glngSys, pסԺҽ������, "23:59:59", Array(dtpEnd))
    lngTmp = Val(zldatabase.GetPara("���ñ���������", glngSys, pסԺҽ������, "0", Array(dtpEnd)))
    dtpEnd.Value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
    
    picParaAdd.Height = chk�ظ���ӡ.Top + chk�ظ���ӡ.Height + 100

    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(0, "ҽ����Ϣ", picPrint.hwnd, 0).Tag = "ҽ��"
       
        .Item(0).Selected = True '�½�ʱ���Զ�ѡ�������,�����ټ����¼�
        'ֻ����ѡ����Ӵ���
        'Call tbcSub_SelectedChanged(.Selected)
    End With
    
    Call RestoreWinState(Me, App.ProductName)
    Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tbcSub.Move tbcSub.Left, tbcSub.Top, Me.ScaleWidth - tbcSub.Left - 50, Me.ScaleHeight - tbcSub.Top - 550
    vsAdvice.Move 0, 0, picPrint.ScaleWidth, picPrint.ScaleHeight
    stbThis.Height = 500
    picCMD.Move picCMD.Left, Me.ScaleHeight - picCMD.Height - stbThis.Height
    frmModle.Move frmModle.Left, picCMD.Top - frmModle.Height - 150
    frmAnd.Move frmAnd.Left, frmAnd.Top, frmAnd.Width, frmModle.Top - frmAnd.Top - 150
    picReport.Height = frmAnd.Height - picParaAdd.Height - picReport.Top - 150
    picParaAdd.Top = picReport.Top + picReport.Height + 450
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mMainPrivs = ""
    mlng����ID = 0
    mlng����ID = 0
    mlngPre����ID = 0
    If Not mobjReportPrivw Is Nothing Then
        Unload mobjReportPrivw
    End If
    Set mobjReportPrivw = Nothing
    Set mobjDatas = Nothing
    Set mobjReport = Nothing
    
    Call SaveWinState(Me, App.ProductName)
End Sub '''

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsColumn
            If .Visible Then
                .Visible = False
                If vsAdvice.Enabled Then vsAdvice.SetFocus
            Else
                .Rows = 1
                For i = vsAdvice.FixedCols + 1 To vsAdvice.Cols - 1
                    .Rows = .Rows + 1
                    If vsAdvice.ColHidden(i) Or vsAdvice.ColWidth(i) = 0 Then
                        .TextMatrix(.Rows - 1, 0) = 0
                    Else
                        .TextMatrix(.Rows - 1, 0) = 1
                    End If
                    .TextMatrix(.Rows - 1, 1) = vsAdvice.TextMatrix(0, i)
                    .RowData(.Rows - 1) = i
                Next
                
                vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 150
                If vsColumn.Top + vsColumn.Height > Me.ScaleHeight Then
                    vsColumn.Height = Me.ScaleHeight - vsColumn.Top
                    vsColumn.Width = 1750
                Else
                    vsColumn.Width = 1470
                End If
                
                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub lvwReport_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strSQL As String, rsTmp As Recordset
    Dim i As Long, lngIndex As Long, strTmp As String, lngPic As Long, lngWidth As Long, lngLeft As Long, lngTop As Long
    Dim arrTmp As Variant, lngParaTop As Long
    
    If Val(Item.Tag & "") > 0 Then
        If mlngPre����ID = Val(Item.Tag & "") Then Exit Sub
        mlngPre����ID = Val(Item.Tag & "")
        If tbcSub.ItemCount > 1 Then
            Call tbcSub.RemoveItem(1)
        End If
        With vsAdvice
            .Clear
            .Rows = 1
            .Cols = 1
            .ColWidth(0) = 230
        End With
        
        lngParaTop = chk�ظ���ӡ.Top + chk�ظ���ӡ.Height + 100
        strSQL = "Select Distinct ����, ����, ȱʡֵ, ��ʽ, ֵ�б�" & vbNewLine & _
                "From zlRPTPars" & vbNewLine & _
                "Where Դid In (Select ID From zlRPTDatas Where ����id = [1]) And ���� Not IN('��ʼ�к�','��ʼʱ��','����ʱ��','��Ч','�ظ���ӡ','������ID','����ID')"
        For i = chkPara.Count - 2 To 0 Step -1
            Unload chkPara(i)
        Next
        For i = txtWB.Count - 2 To 0 Step -1
            Unload txtWB(i)
        Next
        For i = txtSZ.Count - 2 To 0 Step -1
            Unload txtSZ(i)
        Next
        For i = cboPara.Count - 2 To 0 Step -1
            Unload cboPara(i)
        Next
        For i = dtkPara.Count - 2 To 0 Step -1
            Unload dtkPara(i)
        Next
        For i = lblDX.Count - 2 To 0 Step -1
            Unload lblDX(i)
        Next
         For i = lblXL.Count - 2 To 0 Step -1
            Unload lblXL(i)
        Next
         For i = lblRQ.Count - 2 To 0 Step -1
            Unload lblRQ(i)
        Next
         For i = lblWB.Count - 2 To 0 Step -1
            Unload lblWB(i)
        Next
         For i = lblSZ.Count - 2 To 0 Step -1
            Unload lblSZ(i)
        Next
        For i = optPara.Count - 2 To 0 Step -1
            Unload optPara(i)
        Next
        For i = picPara.Count - 2 To 0 Step -1
            Unload picPara(i)
        Next
        On Error GoTo errH
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Item.Tag & ""))
        Do While Not rsTmp.EOF
            '֧�ָ�ѡ
            If rsTmp!ȱʡֵ & "" = "�̶�ֵ�б�" And rsTmp!��ʽ & "" = "2" Then
                lngIndex = chkPara.Count - 1
                Load chkPara(lngIndex)
                Set chkPara(lngIndex).Container = chkPara(999).Container
                chkPara(lngIndex).Caption = rsTmp!���� & ""
                arrTmp = Split(rsTmp!ֵ�б� & "", "|")
                For i = 0 To UBound(arrTmp)
                    If Mid(arrTmp(i), 1, 1) = "��" Then
                        strTmp = Mid(arrTmp(i), 2)
                    Else
                        strTmp = arrTmp(i)
                    End If
                    If Mid(arrTmp(i), 1, 1) = "��" And i = 0 Then chkPara(lngIndex).Value = 1: chkPara(lngIndex).Tag = Split(strTmp, ",")(1)
                    If i = 1 Then chkPara(lngIndex).Tag = chkPara(lngIndex).Tag & "|" & Split(strTmp, ",")(1)
                Next
                
            
                chkPara(lngIndex).Top = lngParaTop
                lngParaTop = lngParaTop + chkPara(lngIndex).Height + 100
                chkPara(lngIndex).ZOrder
                chkPara(lngIndex).Visible = True
            ElseIf rsTmp!ȱʡֵ & "" = "�̶�ֵ�б�" And rsTmp!��ʽ & "" = "1" Then
                '֧�ֵ�ѡ
                lngIndex = lblDX.Count - 1
                Load lblDX(lngIndex)
                Set lblDX(lngIndex).Container = lblDX(999).Container
                lblDX(lngIndex).Caption = rsTmp!���� & ""
                lblDX(lngIndex).Top = lngParaTop + 60
                
                lblDX(lngIndex).ZOrder
                lblDX(lngIndex).Visible = True
                
                lngPic = lngIndex
                Load picPara(lngPic)
                Set picPara(lngPic).Container = picPara(999).Container
                picPara(lngPic).Top = lngParaTop + 30
                arrTmp = Split(rsTmp!ֵ�б� & "", "|")
                For i = 0 To UBound(arrTmp)
                    If Mid(arrTmp(i), 1, 1) = "��" Then
                        strTmp = Mid(arrTmp(i), 2)
                    Else
                        strTmp = arrTmp(i)
                    End If
                    lngIndex = optPara.Count - 1
                    Load optPara(lngIndex)
                    Set optPara(lngIndex).Container = picPara(lngPic)
                    optPara(lngIndex).Caption = Split(strTmp, ",")(0)
                    lblAutoSize.Caption = optPara(lngIndex).Caption
                    optPara(lngIndex).Tag = Split(strTmp, ",")(1)
                    
                    optPara(lngIndex).Width = lblAutoSize.Width + 310
                    If lngLeft + lngWidth + 150 + optPara(lngIndex).Width > picPara(lngPic).Width And lngWidth <> 0 Then
                        '����
                        lngLeft = 0: lngWidth = 0: lngTop = lngTop + optPara(lngIndex).Height + 60
                    End If
                    optPara(lngIndex).Left = lngLeft + lngWidth + IIF(lngWidth > 0, 150, 0)
                    optPara(lngIndex).Top = lngTop
                    
                    
                    If Mid(arrTmp(i), 1, 1) = "��" Then picPara(lngPic).Tag = lngIndex
                    optPara(lngIndex).ZOrder
                    optPara(lngIndex).Visible = True
                    lngLeft = optPara(lngIndex).Left
                    lngTop = optPara(lngIndex).Top
                    lngWidth = optPara(lngIndex).Width
                Next
                If picPara(lngPic).Tag & "" <> "" Then optPara(Val(picPara(lngPic).Tag)).Value = True
                
                picPara(lngPic).ZOrder
                picPara(lngPic).Height = lngTop + optPara(999).Height + 60
                lngParaTop = lngParaTop + picPara(lngPic).Height + 80
                picPara(lngPic).Visible = True
                
            ElseIf rsTmp!ȱʡֵ & "" = "�̶�ֵ�б�" And rsTmp!��ʽ & "" = "0" Then
                '֧������
                lngIndex = lblXL.Count - 1
                Load lblXL(lngIndex)
                Set lblXL(lngIndex).Container = lblXL(999).Container
                lblXL(lngIndex).Caption = rsTmp!���� & ""
                lblXL(lngIndex).Top = lngParaTop + 60
                
                lblXL(lngIndex).ZOrder
                lblXL(lngIndex).Visible = True
                
                Load cboPara(lngIndex)
                Set cboPara(lngIndex).Container = cboPara(999).Container
                arrTmp = Split(rsTmp!ֵ�б� & "", "|")
                For i = 0 To UBound(arrTmp)
                    If Mid(arrTmp(i), 1, 1) = "��" Then
                        strTmp = Mid(arrTmp(i), 2)
                    Else
                        strTmp = arrTmp(i)
                    End If
                    cboPara(lngIndex).AddItem Split(strTmp, ",")(0)
                    cboPara(lngIndex).ItemData(cboPara(lngIndex).NewIndex) = Split(strTmp, ",")(1)
                    If Mid(arrTmp(i), 1, 1) = "��" Then cboPara(lngIndex).ListIndex = cboPara(lngIndex).NewIndex
                Next
                cboPara(lngIndex).Top = lngParaTop
                lngParaTop = lngParaTop + cboPara(lngIndex).Height + 80
                cboPara(lngIndex).ZOrder
                cboPara(lngIndex).Visible = True
            ElseIf rsTmp!���� & "" = "2" Then
                '֧������
                lngIndex = lblRQ.Count - 1
                Load lblRQ(lngIndex)
                Set lblRQ(lngIndex).Container = lblRQ(999).Container
                lblRQ(lngIndex).Caption = rsTmp!���� & ""
                lblRQ(lngIndex).Top = lngParaTop + 60
                
                lblRQ(lngIndex).ZOrder
                lblRQ(lngIndex).Visible = True
                
                Load dtkPara(lngIndex)
                Set dtkPara(lngIndex).Container = dtkPara(999).Container
                If rsTmp!ȱʡֵ & "" <> "�̶�ֵ�б�" And rsTmp!ȱʡֵ & "" <> "ѡ�������塭" Then
                    If InStr(rsTmp!ȱʡֵ & "", ":") > 0 Or InStr(rsTmp!ȱʡֵ & "", "ʱ��") > 0 Then
                        dtkPara(lngIndex).CustomFormat = "yyyy��MM��dd�� HH:mm:ss"
                    Else
                        dtkPara(lngIndex).CustomFormat = "yyyy��MM��dd��"
                    End If
                    If rsTmp!ȱʡֵ & "" <> "" Then
                        If Left(rsTmp!ȱʡֵ & "", 1) = "&" Then
                            dtkPara(lngIndex).Value = GetParVBMacro(rsTmp!ȱʡֵ & "")
                        Else
                            dtkPara(lngIndex).Value = Format(rsTmp!ȱʡֵ & "", dtkPara(lngIndex).CustomFormat)
                        End If
                    Else
                        dtkPara(lngIndex).Value = zldatabase.Currentdate
                    End If
                End If
                dtkPara(lngIndex).Top = lngParaTop
                lngParaTop = lngParaTop + dtkPara(lngIndex).Height + 80
                dtkPara(lngIndex).ZOrder
                dtkPara(lngIndex).Visible = True
            ElseIf rsTmp!���� & "" = "1" Then
                '֧�������ı���
                lngIndex = lblSZ.Count - 1
                Load lblSZ(lngIndex)
                Set lblSZ(lngIndex).Container = lblSZ(999).Container
                lblSZ(lngIndex).Caption = rsTmp!���� & ""
                lblSZ(lngIndex).Top = lngParaTop + 60
                
                lblSZ(lngIndex).ZOrder
                lblSZ(lngIndex).Visible = True
                
                Load txtSZ(lngIndex)
                Set txtSZ(lngIndex).Container = txtSZ(999).Container
                If rsTmp!ȱʡֵ & "" <> "�̶�ֵ�б�" And rsTmp!ȱʡֵ & "" <> "ѡ�������塭" Then
                    txtSZ(lngIndex).Text = rsTmp!ȱʡֵ & ""
                End If
                txtSZ(lngIndex).Top = lngParaTop
                lngParaTop = lngParaTop + txtSZ(lngIndex).Height + 80
                txtSZ(lngIndex).ZOrder
                txtSZ(lngIndex).Visible = True
            ElseIf rsTmp!���� & "" = "0" Then
                '֧���ַ��ı���
                lngIndex = lblWB.Count - 1
                Load lblWB(lngIndex)
                Set lblWB(lngIndex).Container = lblWB(999).Container
                lblWB(lngIndex).Caption = rsTmp!���� & ""
                lblWB(lngIndex).Top = lngParaTop + 60
                
                lblWB(lngIndex).ZOrder
                lblWB(lngIndex).Visible = True
                
                lngIndex = txtWB.Count - 1
                Load txtWB(lngIndex)
                Set txtWB(lngIndex).Container = txtWB(999).Container
                If rsTmp!ȱʡֵ & "" <> "�̶�ֵ�б�" And rsTmp!ȱʡֵ & "" <> "ѡ�������塭" Then
                    txtWB(lngIndex).Text = rsTmp!ȱʡֵ & ""
                End If
                txtWB(lngIndex).Top = lngParaTop
                lngParaTop = lngParaTop + txtWB(lngIndex).Height + 80
                txtWB(lngIndex).ZOrder
                txtWB(lngIndex).Visible = True
            End If
                
            rsTmp.MoveNext
        Loop
        picParaAdd.Height = lngParaTop
        Call Form_Resize
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetParVBMacro(Str As String) As String
'����:�������������,������ת�����VB����ֵ
    Dim curDate As Date
    
    If InStr(Str, "&") = 0 Then GetParVBMacro = Str: Exit Function
    
    curDate = zldatabase.Currentdate
    Select Case Str
        Case "&��ǰ����"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd")
        Case "&��ǰ����ʱ��"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd HH:mm:ss")
        Case "&ǰһ������"
            GetParVBMacro = Format(curDate - 7, "yyyy-MM-dd")
        Case "&ǰһ������"
            GetParVBMacro = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd")
        Case "&ǰһ������"
            GetParVBMacro = Format(DateAdd("m", -3, curDate), "yyyy-MM-dd")
        Case "&ǰһ������"
            GetParVBMacro = Format(DateAdd("yyyy", -1, curDate), "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(curDate + 7, "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(DateAdd("m", 1, curDate), "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(DateAdd("m", 3, curDate), "yyyy-MM-dd")
        Case "&��һ������"
            GetParVBMacro = Format(DateAdd("yyyy", 1, curDate), "yyyy-MM-dd")
        Case "&���쿪ʼʱ��"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 00:00:00")
        Case "&�������ʱ��"
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&ǰһ�쿪ʼʱ��"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd 00:00:00")
        Case "&ǰһ�����ʱ��"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd 23:59:59")
        Case "&ǰһ��ͬʱ��"
            GetParVBMacro = Format(curDate - 1, "yyyy-MM-dd HH:mm:ss")
        Case "&��һ��ͬʱ��"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd HH:mm:ss")
        Case "&��һ�����ʱ��"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd 23:59:59")
        Case "&��һ������"
            GetParVBMacro = Format(curDate + 1, "yyyy-MM-dd")
        Case "&���³�ʱ��"
            GetParVBMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&���³�ʱ��"
            curDate = DateAdd("m", -1, curDate)
            GetParVBMacro = Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParVBMacro = Format(curDate, "yyyy-MM-dd 23:59:59")
        Case "&�����ʱ��"
            GetParVBMacro = Format(Year(curDate) & "-01-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            GetParVBMacro = Format(Year(curDate) & "-12-31", "yyyy-MM-dd 23:59:59")
        Case "&�����ʱ��"
            GetParVBMacro = Format(Year(curDate) - 1 & "-01-01", "yyyy-MM-dd 00:00:00")
        Case "&����ĩʱ��"
            GetParVBMacro = Format(Year(curDate) - 1 & "-12-31", "yyyy-MM-dd 23:59:59")
    End Select
End Function

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    '���ܣ���ӡ֮�󣬸���ҽ�����ϴδ�ӡʱ��
    Dim rsTmp As ADODB.Recordset
    Dim arrPati As Variant, arrSQL As Variant
    Dim strSQL As String, i As Long
    Dim strSQLPati As String, strPatis As String, strTemp As String
    Dim strThis As String, p As Long, n As Long, lngParStar As Long
    Dim varPar(0 To 10) As String, blnTrans As Boolean, lngReportID As Long
    
    On Error GoTo errH
    If mstrPrintedID <> "" Then
        mstrPrintedID = Mid(mstrPrintedID, 2)
        n = 0
        Do While True
            If Len(mstrPrintedID) < 4000 Then
                p = Len(mstrPrintedID) + 1
            Else
                p = InStrRev(Mid(mstrPrintedID, 1, 4000), ",")
            End If
            strThis = Mid(mstrPrintedID, 1, p - 1)
            
            If n > 10 Then
                '̫�����ٴ���ʹ֮����
                varPar(10) = varPar(10) & "," & strThis
            Else
                varPar(n) = strThis
            End If
            
            n = n + 1
            mstrPrintedID = Mid(mstrPrintedID, p + 1)
            If mstrPrintedID = "" Then Exit Do
        Loop
        For i = 1 To lvwReport.ListItems.Count
            If Mid(lvwReport.ListItems(i).Key, 2) = ReportNum Then lngReportID = Val(lvwReport.ListItems(i).Tag): Exit For
        Next
        arrSQL = Array()
        For i = 0 To UBound(varPar)
            If varPar(i) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_ҽ��ִ�е�_��ӡ('" & varPar(i) & "'," & lngReportID & "," & _
                    "To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "To_Date('" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & 0 & ")"
            End If
        Next
        'ִ���ύ����
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            zldatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    mlngLastRow = 0
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjReport_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long)
'���ܣ��������ݴ�ӡ�¼�����¼ҽ����ӡ������
'˵�����������������Ҫ��ӡʱ���ǲ��ἤ����¼���
    If ID <> 0 Then
        If InStr(mstrPrintedID & ",", "," & ID & ",") = 0 Then
            mstrPrintedID = mstrPrintedID & "," & ID
        End If
        mlngLastRow = Row
    End If
End Sub

Private Sub optPati_Click(Index As Integer)
    If Index = 1 Then
        txtסԺ��.Enabled = True
        txtסԺ��.BackColor = &H80000005
        txt��λ��.Enabled = True
        txt��λ��.BackColor = &H80000005
        txt����.Enabled = True
        txt����.BackColor = &H80000005
    Else
        txtסԺ��.Enabled = False
        txtסԺ��.BackColor = &H8000000F
        txtסԺ��.Text = ""
        txt��λ��.Enabled = False
        txt��λ��.BackColor = &H8000000F
        txt��λ��.Text = ""
        txt����.Enabled = False
        txt����.BackColor = &H8000000F
        txt����.Text = ""
        mstr����IDs = ""
    End If
End Sub

Private Sub picReport_Resize()
    On Error Resume Next
    lvwReport.Height = picReport.Height - 100
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Long
    Dim strFilter As String
    
    If Item.Tag = "Ԥ��" And mblnZPriw = False Then
        With vsAdvice
    
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, 1) = 1 Then
                    strFilter = strFilter & " Or " & mstrIDName & "=" & .TextMatrix(i, mintIDIndex)
                End If
            Next
            strFilter = Mid(strFilter, 5)
            If mintIDIndex <> 0 Then
                mobjDatas(mintAdvice).dataset.Filter = IIF(strFilter = "", 0, strFilter)
            Else
                If .Rows > 1 And .Cols > 1 Then
                    MsgBox "ҽ����������Դ��δ����""ID/ҽ��ID""�ֶΣ��޷�����ѡ���ӡ��", vbInformation, Me.Caption
                End If
            End If
        End With
        If mstrFilter = strFilter Then Exit Sub
        
        mstrFilter = strFilter
        Call zlControl.FormLock(Me.hwnd)
        
        Call mobjReport.LoadReport(gcnOracle, Val(lvwReport.SelectedItem.SubItems(1)), Mid(lvwReport.SelectedItem.Key, 2), Me, mobjReportPrivw, mobjDatas, _
          CStr(mvarPar(0)), CStr(mvarPar(1)), CStr(mvarPar(2)), CStr(mvarPar(3)), CStr(mvarPar(4)), CStr(mvarPar(5)), CStr(mvarPar(6)), CStr(mvarPar(7)), CStr(mvarPar(8)), CStr(mvarPar(9)), _
          CStr(mvarPar(10)), CStr(mvarPar(11)), CStr(mvarPar(12)), CStr(mvarPar(13)), CStr(mvarPar(14)), CStr(mvarPar(15)), CStr(mvarPar(16)), CStr(mvarPar(17)), CStr(mvarPar(18)), 1)
        
        Me.tbcSub.RemoveItem 1
        Me.tbcSub.InsertItem(1, lvwReport.SelectedItem.Text & "-��ӡԤ��", mobjReportPrivw.hwnd, 0).Tag = "Ԥ��"
        
        mblnZPriw = True
        tbcSub.Item(1).Selected = True
        mblnZPriw = False
        Call zlControl.FormLock(0)
    End If
End Sub

Private Sub tmrLoad_Timer()
    '��cbo����ж�ؿؼ��ᱨ��365
    If tmrLoad.Enabled Then
        If lvwReport.ListItems.Count > 0 Then
            lvwReport.ListItems(1).Selected = True
            Call lvwReport_ItemClick(lvwReport.ListItems(1))
        End If
    End If
    tmrLoad.Enabled = False
End Sub

Private Sub txtWB_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtWB(Index)
End Sub

Private Sub txtSZ_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtSZ(Index)
End Sub

Private Sub txt��λ��_GotFocus()
    zlControl.TxtSelAll txt��λ��
End Sub

Private Sub txt��λ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FindPati 1, txt��λ��.Text
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FindPati 2, txt����.Text
    End If
End Sub

Private Sub txtסԺ��_GotFocus()
    zlControl.TxtSelAll txtסԺ��
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FindPati 0, txtסԺ��.Text
    End If
    
End Sub

Private Sub FindPati(ByVal intType As Integer, ByVal strText As String)
'���ܣ����ҵ�������
'����:inttype=0-סԺ�ţ�1-���ţ�2-����
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    If Trim(strText) = "" Then Exit Sub
    If intType = 0 Then
        strSQL = "Select rownum as id ,����ID,��ҳID,����,�Ա�,����,��Ժ����,סԺ��,��Ժ���� From ������ҳ Where סԺ��=[1]"
        vRect = zlControl.GetControlRect(txtסԺ��.hwnd)
    ElseIf intType = 1 Then
        strSQL = "Select rownum as id ,a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.��Ժ����, a.סԺ��,A.��Ժ����" & vbNewLine & _
                "From ������ҳ A, ��Ժ���� B" & vbNewLine & _
                "Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.��Ժ���� Like [2] And b.����id = [3]"
        vRect = zlControl.GetControlRect(txt��λ��.hwnd)
    ElseIf intType = 2 Then
        strSQL = "Select rownum as id ,a.����id, a.��ҳid, a.����, a.�Ա�, a.����, a.��Ժ����, a.סԺ��,A.��Ժ����" & vbNewLine & _
                "From ������ҳ A, ��Ժ���� B" & vbNewLine & _
                "Where a.����id = b.����id And a.��ҳid = b.��ҳid And A.���� Like [4] And b.����id = [3]"
        vRect = zlControl.GetControlRect(txt����.hwnd)
    End If
    
    On Error GoTo errH
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, "ѡ����", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txtסԺ��.Height, blnCancel, False, True, strText, "%" & strText & "%", cboUnit.ItemData(cboUnit.ListIndex), "%" & strText & "%")
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û�п��õĸò��ˣ���ȷ�ϲ��������Ƿ���ȷ��", vbInformation, gstrSysName
        End If
    Else
        mstr����IDs = Val(rsTmp!����ID & "") & ":" & Val(rsTmp!��ҳID & "")
        txt��λ��.Text = rsTmp!��Ժ���� & ""
        txt����.Text = rsTmp!���� & ""
        txtסԺ��.Text = rsTmp!סԺ�� & ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long, intval As Integer
    
    If Col = 1 Then
        intval = vsAdvice.Cell(flexcpChecked, Row, Col)
        If RowInһ����ҩ(Row, lngBegin, lngEnd) Then
            For i = lngBegin To lngEnd
                vsAdvice.Cell(flexcpChecked, i, Col) = intval
            Next
        End If
    End If
End Sub

Private Sub vsAdvice_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Dim i As Long, strTmp As String
    
    For i = 2 To vsAdvice.Cols - 1
        strTmp = strTmp & "|" & vsAdvice.TextMatrix(0, i)
    Next
    If strTmp <> "" Then strTmp = strTmp & "|"
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & "������˳��" & "", "_" & lvwReport.SelectedItem.Tag, strTmp
End Sub

Private Sub vsAdvice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsAdvice_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If vsAdvice.TextMatrix(0, Col) = "" Then Position = -1: Col = -1: Exit Sub
End Sub

Private Sub vsAdvice_BeforeSort(ByVal Col As Long, Order As Integer)
    Dim i As Long
    
    Order = 0
    If Col = colѡ�� Then
        With vsAdvice
            If .ColData(colѡ��) = "Check" Then
                .Cell(flexcpPicture, 0, colѡ��) = imgAdvice.ListImages("UnCheck").Picture
                .Cell(flexcpPictureAlignment, 0, colѡ��) = flexPicAlignCenterCenter
                .ColData(colѡ��) = "UnCheck"
                For i = 1 To .Rows - 1
                    .Cell(flexcpChecked, i, colѡ��) = 0
                Next
            Else
                .Cell(flexcpPicture, 0, colѡ��) = imgAdvice.ListImages("AllCheck").Picture
                .Cell(flexcpPictureAlignment, 0, colѡ��) = flexPicAlignCenterCenter
                .ColData(colѡ��) = "Check"
                For i = 1 To .Rows - 1
                    .Cell(flexcpChecked, i, colѡ��) = 1
                Next
            End If
        End With
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < 2 Then Cancel = True
End Sub

Private Function CheckColMegge(strName As String) As Integer
'���أ�����1=���˺ϲ���2=ҽ��һ���ϲ�
    If InStr("|����|�Ա�|����|����|��λ|סԺ��|����|����|", "|" & strName & "|") > 0 Then CheckColMegge = 1
    If InStr("|��ʼʱ��|��Чʱ��|��ֹʱ��|ֹͣʱ��|ͣ��ʱ��|����ʱ��|Ƶ��|Ƶ��|У��ʱ��|У����|����ҽ��|����|�÷�|����|", "|" & strName & "|") > 0 Then CheckColMegge = 2
End Function

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵��:PASS �е� ��RowInһ����ҩ�� ��˷�����ͬ,�޸Ĵ˷���Ҳ��Ҫͬ���޸� PASSͬ������
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If mint���IDIndex = 0 Or mintIDIndex = 0 Then Exit Function

        If Val(.TextMatrix(lngRow - 1, mint���IDIndex)) = Val(.TextMatrix(lngRow, mint���IDIndex)) Then
            blnTmp = True
        ElseIf Val(.TextMatrix(lngRow - 1, mint���IDIndex)) = Val(.TextMatrix(lngRow, mintIDIndex)) And Val(.TextMatrix(lngRow - 1, mint���IDIndex)) <> 0 Then
            blnTmp = True
        ElseIf Val(.TextMatrix(lngRow - 1, mintIDIndex)) = Val(.TextMatrix(lngRow, mint���IDIndex)) And Val(.TextMatrix(lngRow, mint���IDIndex)) <> 0 Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, mint���IDIndex)) = Val(.TextMatrix(lngRow, mint���IDIndex)) Then
                blnTmp = True
            ElseIf Val(.TextMatrix(lngRow + 1, mint���IDIndex)) = Val(.TextMatrix(lngRow, mintIDIndex)) And Val(.TextMatrix(lngRow + 1, mint���IDIndex)) <> 0 Then
                blnTmp = True
            ElseIf Val(.TextMatrix(lngRow + 1, mintIDIndex)) = Val(.TextMatrix(lngRow, mint���IDIndex)) And Val(.TextMatrix(lngRow, mint���IDIndex)) <> 0 Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, mint���IDIndex)) = Val(.TextMatrix(lngRow, mint���IDIndex)) Then
                    lngBegin = i
                ElseIf Val(.TextMatrix(i, mintIDIndex)) = Val(.TextMatrix(lngRow, mint���IDIndex)) And Val(.TextMatrix(lngRow, mint���IDIndex)) <> 0 Then
                    lngBegin = i
                ElseIf Val(.TextMatrix(i, mint���IDIndex)) = Val(.TextMatrix(lngRow, mintIDIndex)) And Val(.TextMatrix(i, mint���IDIndex)) <> 0 Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, mint���IDIndex)) = Val(.TextMatrix(lngRow, mint���IDIndex)) Then
                    lngEnd = i
                ElseIf Val(.TextMatrix(i, mintIDIndex)) = Val(.TextMatrix(lngRow, mint���IDIndex)) And Val(.TextMatrix(lngRow, mint���IDIndex)) <> 0 Then
                    lngEnd = i
                ElseIf Val(.TextMatrix(i, mint���IDIndex)) = Val(.TextMatrix(lngRow, mintIDIndex)) And Val(.TextMatrix(i, mint���IDIndex)) <> 0 Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Function RowInͬһ����(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵��:PASS �е� ��RowInһ����ҩ�� ��˷�����ͬ,�޸Ĵ˷���Ҳ��Ҫͬ���޸� PASSͬ������
    Dim i As Long, blnTmp As Boolean
    Dim lng����ID As Long, lngID As Long
    With vsAdvice
        For i = 2 To .Cols - 1
            If .TextMatrix(0, i) = "����ID" Or .TextMatrix(0, i) = "����" Or .TextMatrix(0, i) = "��������" Then lng����ID = i
        Next
        If lng����ID = 0 Then Exit Function

        If .TextMatrix(lngRow - 1, lng����ID) = .TextMatrix(lngRow, lng����ID) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If .TextMatrix(lngRow + 1, lng����ID) = .TextMatrix(lngRow, lng����ID) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If .TextMatrix(i, lng����ID) = .TextMatrix(lngRow, lng����ID) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If .TextMatrix(i, lng����ID) = .TextMatrix(lngRow, lng����ID) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInͬһ���� = blnTmp
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT, vBrush As LOGBRUSH
    Dim lngPen As Long, lngPenSel As Long
    Dim lngBrush As Long, lngBrushSel As Long
    Dim intReturn As Integer

    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '�����̶����еı����
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)

            '����߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ϱ߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���±߱����
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ұ߱����
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            'Ҫ�ϲ�����
            intReturn = CheckColMegge(.TextMatrix(0, Col))
            If intReturn = 0 Or Row = 0 Then
                Exit Sub
            ElseIf intReturn = 1 Then
                If Not RowInͬһ����(Row, lngBegin, lngEnd) Then Exit Sub
            ElseIf intReturn = 2 Then
                If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
            End If
            
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
        End If
    End With
End Sub

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    Dim i As Long, strTmp As String
    
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            vsAdvice.ColHidden(lngCol) = False
        Else
            vsAdvice.ColHidden(lngCol) = True
        End If
        For i = 2 To vsAdvice.Cols - 1
            If vsAdvice.ColHidden(i) = True Or vsAdvice.ColWidth(i) = 0 Then
                strTmp = strTmp & "|" & vsAdvice.TextMatrix(0, i)
            End If
        Next
        If strTmp <> "" Then strTmp = strTmp & "|"
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & "����������" & "", "_" & lvwReport.SelectedItem.Tag, strTmp
    End If
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then vsColumn.Visible = False
End Sub

Private Sub vsColumn_LostFocus()
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub


